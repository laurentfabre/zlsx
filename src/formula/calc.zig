//! `CT_CellFormula`'s attribute inventory, the shared-formula topology,
//! the slave translation matrix, and the workbook's calc state
//! (`goal_formula.md` §5.7.2, §5.7.6, M4b2).
//!
//! M4b1 stopped at the decoded `<f>` body and kept the raw attribute
//! region so this row would have the bytes to inventory. This is that
//! row: everything an `<f>` can say about itself, and everything
//! `<calcPr>` can say about how the workbook is calculated.
//!
//! Why the attributes are a table and not a switch
//! -----------------------------------------------
//! The typed reader discards `<f>` attributes wholesale
//! (`sheet_xml.zig:489-535`). Discarding `t="shared"` turns a slave into
//! a blank; discarding `dt2D` turns a data table into an ordinary
//! formula that computes one cell of it; discarding `bx` silently drops
//! a name assignment. So every attribute the schema defines gets a row,
//! the parser *reads the table* rather than restating it, and an
//! attribute with no row refuses. A thirteenth attribute invented by
//! some future Office is then a refusal rather than a silent
//! reinterpretation — which is the only safe default for an engine that
//! writes values back.
//!
//! Classification is by ATTRIBUTE, and the corpus is why
//! -----------------------------------------------------
//! `src/xlsx.zig:2099` recognizes a shared slave only when its `<f>` is
//! *self-closing*, and drops every other one.
//! `tests/corpus/calamine_non_monotonic_si.xlsx` writes its slaves as
//! `<f si="1" t="shared"></f>` — an open/close pair with an empty body —
//! so a shape-based reader loses nine of that workbook's twelve shared
//! cells and recalculates them as blanks. The rule here is the schema's:
//! **`t="shared"` with `ref` is a master, `t="shared"` without `ref` is a
//! slave**, whatever shape the element happens to have and whatever
//! order its attributes are written in (that same workbook writes
//! `ref si t`, with `t` last).
//!
//! The same workbook decides two more things. Its `si` values are
//! **non-monotonic** — masters appear in the order 1, 0, 2 — so nothing
//! here may assume `si` counts up; and its master `ref="A3:A7"` covers a
//! row that has no cell at all, so **`ref` is not a topology gate**. What
//! *is* a gate is order: a master must precede its slaves sheet-wide,
//! because a slave is defined by translating a formula it has not been
//! given yet.
//!
//! Translation is over the AST, not over the bytes
//! -----------------------------------------------
//! `rewriter.zig` shifts references by rewriting text and defers full
//! rows and columns to M10+. A slave cannot defer: `A:A` inside a shared
//! master has to translate now, and so does `$A1`, `Sheet1:Sheet3!A1`,
//! and `A1#`. So translation copies the AST, shifts every *relative*
//! half of every reference by (Δrow, Δcol), and prints the result. A
//! reference that leaves the grid becomes `#REF!` as a whole operand —
//! never a half-valid range — because that is the only form the file
//! format itself can express (`PtgRefErr`/`PtgAreaErr`; MS-XLS §2.5.198).
//!
//! Provenance
//! ----------
//! Corpus-decided: slave shape, attribute order, non-monotonic `si`, and
//! `ref` reaching past the occupied cells
//! (`tests/corpus/calamine_non_monotonic_si.xlsx`); `ca` on ordinary
//! shared groups and a 2-D master `ref`
//! (`tests/corpus/openxlsx_loadExample.xlsx`); the `<calcPr>` shapes that
//! must round-trip, including the empty element, the one with a trailing
//! space before `/>`, and the `iterate`/`refMode` set POI writes (22
//! corpus workbooks carry one). Everything else is spec-pinned and cited
//! at its table: the attribute inventory to ECMA-376 `CT_CellFormula`,
//! the calc-state inventory to `CT_CalcPr` and `CT_SheetCalcPr`, the
//! boolean lexical space to `xsd:boolean`, and the `#REF!` collapse to
//! MS-XLS's error-reference tokens.

const std = @import("std");
const assert = std.debug.assert;
const Allocator = std.mem.Allocator;

const coords = @import("zlsx_refs");
const decode = @import("decode.zig");
const parser = @import("parser.zig");
const serial_date = @import("serial_date.zig");

/// §10's plane-2 taxonomy has exactly one home; this file pays the same
/// import `decode.zig` and `metadata.zig` do, for the same reason.
pub const PlaneTwo = decode.PlaneTwo;
pub const CellSite = decode.CellSite;

// ─── refusals (§10) ──────────────────────────────────────────────

pub const Refusal = struct {
    reason: Reason,
    /// Byte offset into the part, when the refusal was found by
    /// scanning one.
    offset: u32 = 0,
    /// Set when the refusal is about one cell.
    cell: ?CellSite = null,

    pub const Reason = enum {
        // ── `CT_CellFormula` attributes (§5.7.2) ──
        /// An attribute with no row in `formula_attrs`. The typed
        /// reader discards attributes; this refuses instead, because
        /// three of the twelve change what the cell's formula *is*.
        unknown_formula_attribute,
        /// A recognized attribute whose value is not the lexical form
        /// its type demands — an `si` that is not an unsigned int, a
        /// `ca` that is not an `xsd:boolean`, a `ref` that is not a
        /// range.
        bad_formula_attribute_value,
        /// `t="dataTable"`, or any of the six data-table attributes.
        /// The construct is a what-if grid whose recalculation
        /// semantics v1 does not reproduce.
        data_table_formula,
        /// `bx="true"`. The formula assigns a name as a side effect,
        /// which is not something an engine that evaluates cells can
        /// honour. `bx="0"`/`"false"` — the value Office writes when it
        /// writes the attribute at all — is accepted and preserved.
        box_formula,
        /// A `t="array"` with no `ref`. §5.6h places a CSE result *in
        /// declared-range terms*, and there is no declared range.
        array_missing_ref,

        // ── shared topology (§5.7.1) ──
        /// `t="shared"` with no `si`. Nothing names the group.
        shared_missing_si,
        /// A slave whose `si` names no master anywhere on the sheet.
        shared_unknown_si,
        /// A slave that appears *before* the master it belongs to. The
        /// reader currently drops such a slave (`xlsx.zig:2099` reaches
        /// only self-closing shapes, and the group is not open yet);
        /// dropping a cell is what an engine that writes values back
        /// must not do.
        shared_slave_before_master,
        /// Two masters claiming one `si`. Which formula the slaves
        /// translate is then a coin toss.
        shared_duplicate_si,
        /// A master with an empty body. There is nothing for its slaves
        /// to translate.
        shared_master_empty,
        /// A master whose body does not parse. The slaves' formulas are
        /// derived from it, so the whole group is unreadable.
        shared_master_unparsable,

        // ── calculation granularity (§5.7.2) ──
        /// A legacy CSE array inside a cycle with `aca` false —
        /// per-cell granularity inside a cycle is a semantics this
        /// engine does not reproduce. Raised by the cycle detector
        /// (M5a), classified here.
        cse_array_per_cell_in_cycle,

        // ── calc state (§5.7.6) ──
        /// `fullPrecision="0"` — precision as displayed. Refused
        /// through v1: every stored value would have to be rounded to
        /// its number format before use.
        precision_as_displayed,
        /// A `<calcPr>`/`<sheetCalcPr>`/`<workbookPr>` attribute whose
        /// value is not its type's lexical form. Unknown *attributes*
        /// are preserved rather than refused — the calc element is
        /// carried back out byte-exact — but a `calcMode` nothing has
        /// classified would change whether the workbook recalculates.
        bad_calc_attribute_value,
        /// The part is not well-formed enough to find its calc state.
        malformed_calc_part,
    };

    /// Exhaustive by construction — a new `Reason` fails to compile
    /// until it has a §10 plane.
    pub fn planeTwo(self: Refusal) PlaneTwo {
        return switch (self.reason) {
            .unknown_formula_attribute,
            .box_formula,
            .cse_array_per_cell_in_cycle,
            => .FormulaUnsupportedConstruct,

            .bad_formula_attribute_value,
            .array_missing_ref,
            .shared_missing_si,
            .shared_unknown_si,
            .shared_slave_before_master,
            .shared_duplicate_si,
            .shared_master_empty,
            .shared_master_unparsable,
            .bad_calc_attribute_value,
            .malformed_calc_part,
            => .FormulaMalformedInput,

            .data_table_formula => .FormulaDataTableUnsupported,
            .precision_as_displayed => .FormulaPrecisionAsDisplayed,
        };
    }
};

// ─── the `CT_CellFormula` attribute inventory (§5.7.2) ───────────

/// What this engine does with an attribute. Not documentation:
/// `classifyFormula` switches on it, so a row cannot be added without
/// deciding.
pub const Treatment = enum {
    /// Read into the typed model and acted on.
    modeled,
    /// Read, kept, carried back out — and read by nothing in v1.
    preserved,
    /// Its presence names a construct v1 does not implement.
    refuses,
    /// One value is accepted, the other refuses. The row's note says
    /// which.
    conditional,
};

pub const FormulaAttrRow = struct {
    /// The qualified name as it appears in the file. Exactly one row is
    /// prefixed (`xml:space`), and its prefix is part of the match.
    name: []const u8,
    treatment: Treatment,
    /// The §10 plane a `.refuses` row raises. Null for the others.
    refusal: ?Refusal.Reason = null,
    note: []const u8,
};

/// ECMA-376 `CT_CellFormula`, complete, plus the one attribute the XML
/// namespace contributes. `classifyFormula` iterates this array — an
/// attribute absent from it is `unknown_formula_attribute` by
/// construction, not by a forgotten `else` branch.
pub const formula_attrs = [_]FormulaAttrRow{
    .{
        .name = "t",
        .treatment = .modeled,
        .note = "ST_CellFormulaType: normal | array | dataTable | shared. " ++
            "Absent is `normal`.",
    },
    .{
        .name = "ref",
        .treatment = .modeled,
        .note = "The range a shared group covers or a CSE array is declared " ++
            "over. Informational for shared groups — corpus masters reach " ++
            "past their occupied cells — and normative for arrays (§5.6h).",
    },
    .{
        .name = "si",
        .treatment = .modeled,
        .note = "Shared group index. Non-monotonic in the corpus, so it is " ++
            "a key and never an ordinal.",
    },
    .{
        .name = "ca",
        .treatment = .preserved,
        .note = "\"Calculate Cell\" — an always-recalculate/dirty HINT. Not " ++
            "function volatility: it never reaches RNG scheduling (§5.5), " ++
            "and 24 corpus formulas carry it on ordinary arithmetic.",
    },
    .{
        .name = "aca",
        .treatment = .preserved,
        .note = "Always-calculate-array — calculation GRANULARITY, not " ++
            "volatility. Both values evaluate whole here, which is " ++
            "identical in an acyclic graph; the one place granularity is " ++
            "observable is a cycle, and that refuses (`cycleRefusal`).",
    },
    .{
        .name = "bx",
        .treatment = .conditional,
        .refusal = .box_formula,
        .note = "`true`/`1` refuses — the formula assigns a name as a side " ++
            "effect. `false`/`0`, which is what Office writes when it " ++
            "writes the attribute, is accepted and preserved.",
    },
    .{
        .name = "dt2D",
        .treatment = .refuses,
        .refusal = .data_table_formula,
        .note = "Data table: two-variable rather than one.",
    },
    .{
        .name = "dtr",
        .treatment = .refuses,
        .refusal = .data_table_formula,
        .note = "Data table: the one-variable input is a row, not a column.",
    },
    .{
        .name = "del1",
        .treatment = .refuses,
        .refusal = .data_table_formula,
        .note = "Data table: the first input cell was deleted.",
    },
    .{
        .name = "del2",
        .treatment = .refuses,
        .refusal = .data_table_formula,
        .note = "Data table: the second input cell was deleted.",
    },
    .{
        .name = "r1",
        .treatment = .refuses,
        .refusal = .data_table_formula,
        .note = "Data table: the first input cell.",
    },
    .{
        .name = "r2",
        .treatment = .refuses,
        .refusal = .data_table_formula,
        .note = "Data table: the second input cell.",
    },
    .{
        .name = "xml:space",
        .treatment = .preserved,
        .note = "`default` | `preserve`. Excel writes it on formulas whose " ++
            "text has significant leading or trailing whitespace; the " ++
            "FORMULA carrier already preserves the body byte-exact, so " ++
            "this is recorded rather than acted on.",
    },
};

comptime {
    // The inventory is complete only if it is also unambiguous: two rows
    // for one name would make the treatment depend on iteration order.
    for (formula_attrs, 0..) |a, i| {
        for (formula_attrs[i + 1 ..]) |b| {
            assert(!std.mem.eql(u8, a.name, b.name));
        }
        assert((a.refusal != null) == (a.treatment == .refuses or a.treatment == .conditional));
    }
}

/// `ST_CellFormulaType`, with the schema's default named.
pub const Kind = enum {
    normal,
    array,
    data_table,
    shared,

    pub fn fromAttr(raw: ?[]const u8) ?Kind {
        const s = raw orelse return .normal;
        if (std.mem.eql(u8, s, "normal")) return .normal;
        if (std.mem.eql(u8, s, "array")) return .array;
        if (std.mem.eql(u8, s, "dataTable")) return .data_table;
        if (std.mem.eql(u8, s, "shared")) return .shared;
        return null;
    }
};

/// `xml:space`. Kept as an enum rather than a bool so the absent case is
/// distinguishable from an explicit `default` — a writer that dropped an
/// explicit `xml:space="default"` would change the bytes.
pub const XmlSpace = enum { absent, default, preserve };

/// One `<f>`, with every attribute the schema defines accounted for.
pub const CellFormula = struct {
    kind: Kind,
    /// Decoded body (FORMULA carrier — entities only). Empty for a
    /// bodiless or empty-bodied slave.
    text: []const u8,
    /// `si`, when present. Present on every shared `<f>`; the schema
    /// permits it elsewhere, where it means nothing and is preserved.
    si: ?u32,
    /// `ref`, normalized. Null when absent.
    ref: ?coords.Range,
    /// `ref`'s original spelling, so a writer can hand back what it was
    /// given rather than a re-formatted equivalent.
    ref_raw: ?[]const u8,
    ca: bool,
    aca: bool,
    /// Always false in an accepted formula — `bx="true"` refuses — and
    /// carried because `bx="0"` is a byte a round-trip owes back.
    bx: bool,
    xml_space: XmlSpace,
    /// The whole raw attribute region, borrowed from the part.
    raw_attrs: []const u8,

    /// Whether this formula is a shared *master*: the schema says a
    /// shared `<f>` carrying `ref` defines the group, and one without it
    /// belongs to a group. Shape plays no part.
    pub fn isSharedMaster(self: CellFormula) bool {
        return self.kind == .shared and self.ref != null;
    }

    pub fn isSharedSlave(self: CellFormula) bool {
        return self.kind == .shared and self.ref == null;
    }

    /// The refusal a cycle detector must raise if this formula turns out
    /// to be inside one (§5.7.2's `aca` row). Classified here so M5a's
    /// detector asks rather than re-derives; `aca`'s schema default is
    /// `false`, so an array with no `aca` at all reads as per-cell and
    /// refuses — the conservative direction.
    pub fn cycleRefusal(self: CellFormula) ?Refusal.Reason {
        if (self.kind == .array and !self.aca) return .cse_array_per_cell_in_cycle;
        return null;
    }
};

pub const FormulaResult = union(enum) {
    ok: CellFormula,
    refused: Refusal,
};

/// Classify one `<f>` against the inventory.
///
/// Pure, allocation-free, and total: every byte string reaching it
/// either produces a `CellFormula` or a typed refusal, which is what
/// makes the fuzz target's "never classifies two ways" invariant
/// checkable.
pub fn classifyFormula(f: decode.Formula, site: ?CellSite) FormulaResult {
    var out: CellFormula = .{
        .kind = .normal,
        .text = f.text,
        .si = null,
        .ref = null,
        .ref_raw = null,
        .ca = false,
        .aca = false,
        .bx = false,
        .xml_space = .absent,
        .raw_attrs = f.raw_attrs,
    };

    var it: decode.AttrIterator = .{ .attrs = f.raw_attrs };
    while (it.next()) |a| {
        // Foreign-namespace attributes are exempt from the inventory —
        // M4b1 decision 8, and the reason `x14ac:dyDescent` does not
        // refuse every row Excel writes. `xml:space` is the one prefixed
        // attribute with a row, matched on its full qualified name.
        const row = rowFor(a.qname) orelse {
            if (a.prefix().len != 0 or std.mem.eql(u8, a.qname, "xmlns")) continue;
            return refuse(.unknown_formula_attribute, site);
        };
        switch (row.treatment) {
            .refuses => return refuse(row.refusal.?, site),
            .modeled, .preserved, .conditional => {},
        }
        if (std.mem.eql(u8, row.name, "t")) {
            out.kind = Kind.fromAttr(a.raw_value) orelse
                return refuse(.bad_formula_attribute_value, site);
        } else if (std.mem.eql(u8, row.name, "ref")) {
            const r = coords.parseRange(a.raw_value, .{
                .case = .upper_only,
                .dollar = .accept,
            }) catch return refuse(.bad_formula_attribute_value, site);
            out.ref = r.normalized();
            out.ref_raw = a.raw_value;
        } else if (std.mem.eql(u8, row.name, "si")) {
            out.si = std.fmt.parseInt(u32, a.raw_value, 10) catch
                return refuse(.bad_formula_attribute_value, site);
        } else if (std.mem.eql(u8, row.name, "ca")) {
            out.ca = xsdBool(a.raw_value) orelse
                return refuse(.bad_formula_attribute_value, site);
        } else if (std.mem.eql(u8, row.name, "aca")) {
            out.aca = xsdBool(a.raw_value) orelse
                return refuse(.bad_formula_attribute_value, site);
        } else if (std.mem.eql(u8, row.name, "bx")) {
            const b = xsdBool(a.raw_value) orelse
                return refuse(.bad_formula_attribute_value, site);
            if (b) return refuse(row.refusal.?, site);
            out.bx = false;
        } else if (std.mem.eql(u8, row.name, "xml:space")) {
            if (std.mem.eql(u8, a.raw_value, "preserve")) {
                out.xml_space = .preserve;
            } else if (std.mem.eql(u8, a.raw_value, "default")) {
                out.xml_space = .default;
            } else {
                return refuse(.bad_formula_attribute_value, site);
            }
        } else unreachable; // every row is handled, and the comptime
        // block above proves the names are distinct
    }

    // `t="dataTable"` reaches the same refusal the six attributes do:
    // the type and its machinery are one construct, and a producer may
    // write either half first.
    if (out.kind == .data_table) return refuse(.data_table_formula, site);
    if (out.kind == .shared and out.si == null) return refuse(.shared_missing_si, site);
    if (out.kind == .array and out.ref == null) return refuse(.array_missing_ref, site);
    return .{ .ok = out };
}

fn rowFor(qname: []const u8) ?FormulaAttrRow {
    for (formula_attrs) |row| {
        if (std.mem.eql(u8, row.name, qname)) return row;
    }
    return null;
}

fn refuse(reason: Refusal.Reason, site: ?CellSite) FormulaResult {
    return .{ .refused = .{ .reason = reason, .cell = site } };
}

/// `xsd:boolean`'s complete lexical space, and nothing else. `"TRUE"` is
/// not in it — that is Excel's spelling for a cell value, not XML's for
/// an attribute.
pub fn xsdBool(s: []const u8) ?bool {
    if (std.mem.eql(u8, s, "1") or std.mem.eql(u8, s, "true")) return true;
    if (std.mem.eql(u8, s, "0") or std.mem.eql(u8, s, "false")) return false;
    return null;
}

// ─── sheet-wide shared topology (§5.7.1) ─────────────────────────

/// One shared group, keyed by its `si`.
pub const Group = struct {
    si: u32,
    /// The master's own coordinate. Translation is relative to this and
    /// not to `ref.first`: the two coincide in every corpus workbook,
    /// but the schema anchors a shared formula at the cell that carries
    /// it, and a producer that wrote a `ref` starting elsewhere would
    /// otherwise shift the whole group.
    anchor: coords.Cell,
    /// The declared range. Informational — corpus masters reach past
    /// their occupied cells — and kept because a writer owes it back.
    ref: coords.Range,
    /// The master's decoded body, borrowed.
    text: []const u8,
    /// Index into the cell slice `classifySheet` was given.
    cell: u32,
    /// Where the master sits, for a refusal that has to name it. Kept
    /// here rather than re-derived from `cell`, because a caller that
    /// has already dropped the cell slice would have nothing to derive
    /// it from.
    site: CellSite,
};

pub const Delta = struct {
    rows: i32 = 0,
    cols: i32 = 0,

    pub fn isZero(self: Delta) bool {
        return self.rows == 0 and self.cols == 0;
    }
};

pub const Role = union(enum) {
    /// The cell has no `<f>`.
    none,
    normal,
    /// A legacy CSE array, over its declared range.
    array: coords.Range,
    /// Defines the group at this index in `SheetShared.groups`.
    master: u32,
    slave: struct { group: u32, delta: Delta },
};

pub const Entry = struct {
    /// Index into the cell slice.
    cell: u32,
    formula: CellFormula,
    role: Role,
};

pub const SheetShared = struct {
    allocator: Allocator,
    /// One per cell that has an `<f>`, in document order.
    entries: []Entry,
    groups: []Group,

    pub fn deinit(self: *SheetShared) void {
        self.allocator.free(self.entries);
        self.allocator.free(self.groups);
        self.* = undefined;
    }

    /// The group a slave translates from, or null for anything else.
    pub fn groupOf(self: SheetShared, e: Entry) ?Group {
        return switch (e.role) {
            .slave => |s| self.groups[s.group],
            else => null,
        };
    }
};

pub const SheetResult = union(enum) {
    ok: SheetShared,
    refused: Refusal,
};

/// Classify every `<f>` on one sheet and resolve the shared topology.
///
/// `cells` must be in document order, which is what `decode.scanSheet`
/// returns. Two passes, because the two failure modes a slave has are
/// different statements: a master that appears *later* is an ordering
/// fault, and no master at all is a dangling reference. One pass could
/// only report the second.
pub fn classifySheet(
    gpa: Allocator,
    cells: []const decode.SheetCell,
) error{OutOfMemory}!SheetResult {
    var entries: std.ArrayListUnmanaged(Entry) = .empty;
    errdefer entries.deinit(gpa);
    var groups: std.ArrayListUnmanaged(Group) = .empty;
    errdefer groups.deinit(gpa);

    // Pass one: classify every formula, and collect the masters.
    for (cells, 0..) |c, i| {
        const f = c.formula orelse continue;
        const site: CellSite = .{ .row = c.row.oneBased(), .col = c.col.zeroBased() };
        const cf = switch (classifyFormula(f, site)) {
            .ok => |v| v,
            .refused => |r| {
                entries.deinit(gpa);
                groups.deinit(gpa);
                return .{ .refused = r };
            },
        };
        try entries.append(gpa, .{ .cell = @intCast(i), .formula = cf, .role = .normal });

        if (cf.isSharedMaster()) {
            const si = cf.si.?;
            for (groups.items) |g| {
                if (g.si == si) {
                    entries.deinit(gpa);
                    groups.deinit(gpa);
                    return .{ .refused = .{ .reason = .shared_duplicate_si, .cell = site } };
                }
            }
            if (cf.text.len == 0) {
                entries.deinit(gpa);
                groups.deinit(gpa);
                return .{ .refused = .{ .reason = .shared_master_empty, .cell = site } };
            }
            try groups.append(gpa, .{
                .si = si,
                .anchor = .{ .col = c.col, .row = c.row },
                .ref = cf.ref.?,
                .text = cf.text,
                .cell = @intCast(i),
                .site = site,
            });
        }
    }

    // Pass two: give every entry its role. Masters are known now, so a
    // slave can be told which of its two failures it has.
    for (entries.items) |*e| {
        const cf = e.formula;
        if (cf.kind == .array) {
            e.role = .{ .array = cf.ref.? };
            continue;
        }
        if (cf.isSharedMaster()) {
            e.role = .{ .master = groupIndexOf(groups.items, cf.si.?).? };
            continue;
        }
        if (!cf.isSharedSlave()) continue;

        const site: CellSite = .{
            .row = cells[e.cell].row.oneBased(),
            .col = cells[e.cell].col.zeroBased(),
        };
        const gi = groupIndexOf(groups.items, cf.si.?) orelse {
            entries.deinit(gpa);
            groups.deinit(gpa);
            return .{ .refused = .{ .reason = .shared_unknown_si, .cell = site } };
        };
        const g = groups.items[gi];
        if (g.cell > e.cell) {
            entries.deinit(gpa);
            groups.deinit(gpa);
            return .{ .refused = .{ .reason = .shared_slave_before_master, .cell = site } };
        }
        e.role = .{ .slave = .{ .group = gi, .delta = deltaBetween(g.anchor, .{
            .col = cells[e.cell].col,
            .row = cells[e.cell].row,
        }) } };
    }

    // Ownership moves one list at a time, each with its own `errdefer`:
    // `toOwnedSlice` empties the list it drains, so a failure on the
    // second would leave the first owned by nothing at all. Found by
    // this file's allocation-failure sweep, which is the only way a
    // two-allocation hand-off in a struct literal ever gets caught.
    const owned_entries = try entries.toOwnedSlice(gpa);
    errdefer gpa.free(owned_entries);
    const owned_groups = try groups.toOwnedSlice(gpa);
    return .{ .ok = .{
        .allocator = gpa,
        .entries = owned_entries,
        .groups = owned_groups,
    } };
}

fn groupIndexOf(groups: []const Group, si: u32) ?u32 {
    for (groups, 0..) |g, i| {
        if (g.si == si) return @intCast(i);
    }
    return null;
}

/// The shift from `from` to `to`. Both coordinates are in-grid, so the
/// difference fits an `i32` with three orders of magnitude to spare.
pub fn deltaBetween(from: coords.Cell, to: coords.Cell) Delta {
    return .{
        .rows = @as(i32, @intCast(to.row.oneBased())) - @as(i32, @intCast(from.row.oneBased())),
        .cols = @as(i32, @intCast(to.col.zeroBased())) - @as(i32, @intCast(from.col.zeroBased())),
    };
}

// ─── the translation matrix (§5.7.1) ─────────────────────────────

/// A translated formula: the copied AST and its printed text, both
/// owned by the arena.
pub const Translated = struct {
    arena: std.heap.ArenaAllocator,
    ast: parser.Ast,
    text: []const u8,

    pub fn deinit(self: *Translated) void {
        self.arena.deinit();
        self.* = undefined;
    }
};

/// The spelling every off-grid reference collapses to. One constant,
/// because the printer, the collapse rule and the tests must agree on
/// it byte for byte.
pub const ref_error = "#REF!";

/// Copy `src`, shifting every relative half of every reference by `d`.
///
/// The AST is copied rather than mutated, so a master can be translated
/// once per slave without re-parsing, and `d == 0` is provably the
/// identity (`translate(ast, .{}).text == ast.print()`).
///
/// **Reference operands collapse whole.** If any endpoint of a range,
/// full column, full row or spill leaves the grid, the entire operand
/// becomes `#REF!` — never `#REF!:B2`, which is a spelling the file
/// format has no token for (MS-XLS `PtgAreaErr`). A `#REF!` that was
/// already in the source is left exactly where it was: collapse is a
/// consequence of *this* translation, so a formula nothing moved comes
/// back byte-identical.
pub fn translate(
    gpa: Allocator,
    src: parser.Ast,
    d: Delta,
) error{OutOfMemory}!Translated {
    var arena = std.heap.ArenaAllocator.init(gpa);
    errdefer arena.deinit();
    const a = arena.allocator();

    const nodes = try a.dupe(parser.Node, src.nodes);
    const extra = try a.dupe(parser.Index, src.extra);

    // Which nodes *this* translation turned into `#REF!`. A source
    // `#REF!` is not in the set, so `=SUM(A1:#REF!)` translated by zero
    // stays `=SUM(A1:#REF!)`.
    const collapsed = try a.alloc(bool, nodes.len);
    @memset(collapsed, false);

    // Children always have lower indices than their parents — this
    // parser appends a node only once its children are built — so one
    // ascending pass *is* a post-order walk, and the collapse rule sees
    // its operands already decided. Asserted rather than assumed.
    for (nodes, 0..) |*n, i| {
        assertChildrenBelow(nodes[i], extra, i);
        switch (n.*) {
            .ref_cell => |c| {
                const moved = shiftCell(c.cell, d) orelse {
                    n.* = refErrorNode(c.span);
                    collapsed[i] = true;
                    continue;
                };
                if (moved.eqlExact(c.cell)) continue;
                var buf: [coords.format_buf_len]u8 = undefined;
                const text = coords.formatCell(&buf, moved);
                n.* = .{ .ref_cell = .{
                    .span = c.span,
                    .cell = moved,
                    .text = try a.dupe(u8, text),
                } };
            },
            .ref_full_col => |c| {
                const first = shiftColBound(c.first, d) orelse {
                    n.* = refErrorNode(c.span);
                    collapsed[i] = true;
                    continue;
                };
                const last = shiftColBound(c.last, d) orelse {
                    n.* = refErrorNode(c.span);
                    collapsed[i] = true;
                    continue;
                };
                n.* = .{ .ref_full_col = .{ .span = c.span, .first = first, .last = last } };
            },
            .ref_full_row => |c| {
                const first = shiftRowBound(c.first, d) orelse {
                    n.* = refErrorNode(c.span);
                    collapsed[i] = true;
                    continue;
                };
                const last = shiftRowBound(c.last, d) orelse {
                    n.* = refErrorNode(c.span);
                    collapsed[i] = true;
                    continue;
                };
                n.* = .{ .ref_full_row = .{ .span = c.span, .first = first, .last = last } };
            },
            // `A1:B2` with either end off the grid is `#REF!` entire.
            // The other binary operators are ordinary arithmetic and
            // propagate nothing.
            .binary => |b| {
                if (b.op != .range) continue;
                if (collapsed[b.lhs] or collapsed[b.rhs]) {
                    n.* = refErrorNode(b.span);
                    collapsed[i] = true;
                }
            },
            // `A1#` over a dead reference has no spill to name.
            .postfix => |p| {
                if (p.op != .spill) continue;
                if (collapsed[p.child]) {
                    n.* = refErrorNode(p.span);
                    collapsed[i] = true;
                }
            },
            // `Sheet1!A1` keeps its qualifier: the sheet still exists,
            // and Excel writes `Sheet1!#REF!` for exactly this case.
            .qualified => {},
            // Names, structured references, literals and calls carry no
            // coordinate to shift. A defined name whose *body* is
            // relative is §5.9's typed refusal, not a translation.
            else => {},
        }
    }

    var out: parser.Ast = .{
        .source = src.source,
        .body = src.body,
        .nodes = nodes,
        .extra = extra,
        .root = src.root,
    };
    const text = try out.print(a);
    return .{ .arena = arena, .ast = out, .text = text };
}

fn refErrorNode(span: parser.Span) parser.Node {
    return .{ .error_lit = .{ .span = span, .text = ref_error, .known = true } };
}

/// Null when the shifted cell leaves the grid. Absolute halves do not
/// move, which is the entire difference between `$A$1` and `A1`.
fn shiftCell(c: coords.Cell, d: Delta) ?coords.Cell {
    const row = if (c.anchor.row)
        c.row
    else
        coords.Row.fromOneBased(addClamped(c.row.oneBased(), d.rows) orelse return null) catch
            return null;
    const col = if (c.anchor.col)
        c.col
    else
        coords.Col.fromZeroBased(addClamped(c.col.zeroBased(), d.cols) orelse return null) catch
            return null;
    return .{ .col = col, .row = row, .anchor = c.anchor };
}

fn shiftColBound(b: parser.ColBound, d: Delta) ?parser.ColBound {
    if (b.absolute) return b;
    const v = addClamped(b.col.zeroBased(), d.cols) orelse return null;
    return .{ .col = coords.Col.fromZeroBased(v) catch return null, .absolute = false };
}

fn shiftRowBound(b: parser.RowBound, d: Delta) ?parser.RowBound {
    if (b.absolute) return b;
    const v = addClamped(b.row.oneBased(), d.rows) orelse return null;
    return .{ .row = coords.Row.fromOneBased(v) catch return null, .absolute = false };
}

/// Add a signed delta to an unsigned coordinate without wrapping.
/// Returns null on underflow, which the grid check would otherwise see
/// as an enormous in-range column.
fn addClamped(v: u32, delta: i32) ?u32 {
    const sum = @as(i64, v) + @as(i64, delta);
    if (sum < 0 or sum > std.math.maxInt(u32)) return null;
    return @intCast(sum);
}

fn assertChildrenBelow(n: parser.Node, extra: []const parser.Index, i: usize) void {
    switch (n) {
        .array => |x| for (extra[x.elems.start..][0..x.elems.len]) |c| assert(c < i),
        .qualified => |x| assert(x.target < i),
        .call => |x| {
            assert(x.callee < i);
            for (extra[x.args.start..][0..x.args.len]) |c| assert(c < i);
        },
        .paren => |x| assert(x.child < i),
        .unary => |x| assert(x.child < i),
        .postfix => |x| assert(x.child < i),
        .binary => |x| {
            assert(x.lhs < i);
            assert(x.rhs < i);
        },
        else => {},
    }
}

/// Translate a formula given as text: parse, translate, print.
///
/// The convenience the adapter wants — a slave's body is derived from a
/// master's *text*, and the group is parsed once per slave only because
/// caching the master's AST is the model builder's business (M5a).
pub fn translateText(
    gpa: Allocator,
    text: []const u8,
    d: Delta,
) error{OutOfMemory}!TranslatedText {
    var parsed = try parser.parse(gpa, text, .{});
    defer parsed.deinit(gpa);
    const ast = switch (parsed) {
        .refused => |r| return .{ .refused = r },
        .ok => |a| a,
    };
    // The translated AST is a copy in its own arena, so freeing the
    // source parse here is safe; what it still borrows is `text`, which
    // is the caller's and is documented as having to outlive the result.
    return .{ .ok = try translate(gpa, ast, d) };
}

pub const TranslatedText = union(enum) {
    ok: Translated,
    refused: parser.Refusal,
};

/// One sheet's shared masters, parsed once each.
///
/// The reason this type exists rather than a loop over `translateText`:
/// a group with a thousand slaves would otherwise parse its master a
/// thousand times, and the whole point of a shared formula is that the
/// producer wrote it once.
pub const Masters = struct {
    allocator: Allocator,
    /// One per `SheetShared.groups` entry, in the same order.
    asts: []parser.Ast,

    pub fn deinit(self: *Masters) void {
        for (self.asts) |*ast| ast.deinit(self.allocator);
        self.allocator.free(self.asts);
        self.* = undefined;
    }

    pub fn translateFor(
        self: Masters,
        gpa: Allocator,
        group: u32,
        d: Delta,
    ) error{OutOfMemory}!Translated {
        return translate(gpa, self.asts[group], d);
    }
};

pub const MastersResult = union(enum) {
    ok: Masters,
    refused: Refusal,
};

/// Parse every master on a sheet. A master that will not parse refuses
/// the whole group: its slaves' formulas are derived from it, so there
/// is nothing left to read.
pub fn parseMasters(
    gpa: Allocator,
    groups: []const Group,
) error{OutOfMemory}!MastersResult {
    var asts: std.ArrayListUnmanaged(parser.Ast) = .empty;
    errdefer {
        for (asts.items) |*ast| ast.deinit(gpa);
        asts.deinit(gpa);
    }
    try asts.ensureTotalCapacityPrecise(gpa, groups.len);

    for (groups) |g| {
        var parsed = try parser.parse(gpa, g.text, .{});
        switch (parsed) {
            .ok => |ast| asts.appendAssumeCapacity(ast),
            .refused => {
                parsed.deinit(gpa);
                for (asts.items) |*ast| ast.deinit(gpa);
                asts.deinit(gpa);
                return .{ .refused = .{
                    .reason = .shared_master_unparsable,
                    .cell = g.site,
                } };
            },
        }
    }
    return .{ .ok = .{ .allocator = gpa, .asts = try asts.toOwnedSlice(gpa) } };
}

// ─── calc state (§5.7.6) ─────────────────────────────────────────

pub const CalcMode = enum {
    manual,
    auto,
    auto_no_table,

    pub fn fromAttr(s: []const u8) ?CalcMode {
        if (std.mem.eql(u8, s, "manual")) return .manual;
        if (std.mem.eql(u8, s, "auto")) return .auto;
        if (std.mem.eql(u8, s, "autoNoTable")) return .auto_no_table;
        return null;
    }
};

pub const RefMode = enum {
    a1,
    r1c1,

    pub fn fromAttr(s: []const u8) ?RefMode {
        if (std.mem.eql(u8, s, "A1")) return .a1;
        if (std.mem.eql(u8, s, "R1C1")) return .r1c1;
        return null;
    }
};

/// §5.4d's compatibility version. Absent CV metadata is CV1, which is
/// what every pre-2024 workbook is and what a fresh zlsx file emits.
pub const TextCompat = enum { v1, v2 };

pub const DateSystem = serial_date.DateSystem;
/// Re-exported so the package layer reaches the epoch through the same
/// module that decided which epoch it is — `date_system` and the
/// conversions that read it are one contract, not two.
pub const dates = serial_date;

/// The `<ext>` Excel writes its calculation feature list into.
pub const calc_features_ext_uri = "{B58B0392-4F1F-4190-BB64-5DF3571DCE5F}";

pub const CalcFeatureRow = struct {
    name: []const u8,
    /// The compatibility version this feature *implies*, if any.
    compat: ?TextCompat = null,
    note: []const u8,
};

/// The feature names Excel writes into `<xcalcf:calcFeatures>`.
///
/// Every one of them is inert here: they tell Excel which calculation
/// behaviours the producing build had, and zlsx implements its own
/// (§5.3). The table exists because §5.4d's compatibility version rides
/// in the same list, and a name with **no** row leaves `text_compat` at
/// CV1 — §5.4d's documented default for absent CV metadata — rather than
/// being guessed at. Which name carries CV2 is pinned the way §4 pins
/// every other Office-vs-schema difference: from a byte-diffed Excel
/// reference (M7b), not from the base schema. Until that row exists the
/// list is *preserved whole*, so nothing is lost and nothing is
/// invented.
pub const calc_feature_inventory = [_]CalcFeatureRow{
    .{ .name = "microsoft.com:RD", .note = "Rich data types." },
    .{ .name = "microsoft.com:Single", .note = "`_xlfn.SINGLE` / implicit intersection." },
    .{ .name = "microsoft.com:FV", .note = "Formula versioning." },
    .{ .name = "microsoft.com:CNMTM", .note = "Concurrent multi-threaded calculation." },
    .{ .name = "microsoft.com:LET_WF", .note = "LET in a workbook function." },
    .{ .name = "microsoft.com:LAMBDA_WF", .note = "LAMBDA in a workbook function." },
    .{ .name = "microsoft.com:ARRAYTEXT_WF", .note = "Array text functions." },
};

/// A preserved `<ext>` from the workbook-level `<extLst>`.
pub const Extension = struct {
    uri: []const u8,
    /// The whole element, `<ext …>` through `</ext>`, borrowed from the
    /// part. Byte-exact re-emission is the only preservation an
    /// extension this engine does not interpret can honestly offer.
    raw: []const u8,
};

/// Everything `xl/workbook.xml` says about how the workbook calculates.
///
/// Every raw slice borrows the part bytes, so a `CalcState` is valid
/// only while the part it came from is.
pub const CalcState = struct {
    /// False when the workbook has no `<calcPr>` at all. Every typed
    /// field then holds the schema default, and `writeCalcPr` writes
    /// nothing — "absent elements created only when needed" (§5.7.6).
    present: bool = false,
    /// The attribute region of `<calcPr>`, verbatim. Includes the
    /// whitespace: `wdi_excel.xlsx` writes `<calcPr calcId="40001" />`,
    /// and the space before `/>` is a byte a round-trip owes back.
    raw_attrs: []const u8 = "",
    self_closing: bool = true,

    calc_id: u32 = 0,
    calc_mode: CalcMode = .auto,
    full_calc_on_load: bool = false,
    ref_mode: RefMode = .a1,
    iterate: bool = false,
    iterate_count: u32 = 100,
    iterate_delta: f64 = 0.001,
    /// `fullPrecision="0"` refuses, so an accepted state always has it
    /// true. Carried anyway: the field is what a later relaxation would
    /// switch on, and its absence would make that a schema change.
    full_precision: bool = true,
    calc_completed: bool = true,
    calc_on_save: bool = true,
    concurrent_calc: bool = true,
    concurrent_manual_count: ?u32 = null,
    force_full_calc: bool = false,

    /// `<workbookPr date1904>`. Workbook-derived and never
    /// caller-writable (§5.5) — the same text is a different serial
    /// under each epoch.
    date_system: DateSystem = .d1900,
    /// `<workbookPr>`'s attribute region, for the same reason
    /// `raw_attrs` exists.
    workbook_pr_raw: ?[]const u8 = null,

    /// Workbook-level `<ext>` elements, in document order.
    extensions: []const Extension = &.{},
    /// `<xcalcf:feature name="…">` names, in document order, borrowed.
    calc_features: []const []const u8 = &.{},
    text_compat: TextCompat = .v1,

    /// Write the `<calcPr>` element back exactly as it was read.
    ///
    /// Reconstructed from the parsed pieces rather than echoed from a
    /// saved span, so a corpus round-trip proves the parse *kept* what
    /// it needs — an echo would prove only that `memcpy` works.
    pub fn writeCalcPr(self: CalcState, w: *std.Io.Writer) std.Io.Writer.Error!void {
        if (!self.present) return;
        try w.writeAll("<calcPr");
        try w.writeAll(self.raw_attrs);
        if (self.self_closing) {
            try w.writeAll("/>");
        } else {
            try w.writeAll("></calcPr>");
        }
    }

    pub fn deinit(self: *CalcState, gpa: Allocator) void {
        gpa.free(self.extensions);
        gpa.free(self.calc_features);
        self.* = undefined;
    }
};

pub const CalcStateResult = union(enum) {
    ok: CalcState,
    refused: Refusal,
};

/// Parse the calc state out of `xl/workbook.xml`.
///
/// Allocates only the two lists it cannot bound: the workbook's
/// extensions and its calc-feature names. Everything else borrows.
pub fn parseCalcState(gpa: Allocator, xml: []const u8) error{OutOfMemory}!CalcStateResult {
    var out: CalcState = .{};
    var exts: std.ArrayListUnmanaged(Extension) = .empty;
    defer exts.deinit(gpa);
    var features: std.ArrayListUnmanaged([]const u8) = .empty;
    defer features.deinit(gpa);

    var sc = decode.Scanner.init(xml);
    var depth: usize = 0;
    // The workbook-level `<extLst>` only. A `<definedName>` or a sheet
    // may carry its own, and those are not calc state.
    var extlst_depth: ?usize = null;
    var ext_open: ?struct { start: usize, depth: usize, uri: []const u8 } = null;
    var in_calc_features = false;

    while (sc.next() catch {
        return .{ .refused = .{ .reason = .malformed_calc_part } };
    }) |ev| {
        switch (ev) {
            .close => {
                if (depth == 0) return .{ .refused = .{ .reason = .malformed_calc_part } };
                if (ext_open) |e| {
                    if (depth == e.depth) {
                        try exts.append(gpa, .{ .uri = e.uri, .raw = xml[e.start..sc.i] });
                        ext_open = null;
                        in_calc_features = false;
                    }
                }
                if (extlst_depth) |d| {
                    if (depth == d) extlst_depth = null;
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
        // A self-closing element never became part of `depth`, so its
        // own depth is one past the current one.
        const el_depth = if (self_closing) depth + 1 else depth;

        if (el_depth == 2 and std.mem.eql(u8, local, "calcPr")) {
            switch (parseCalcPr(el, self_closing, &out)) {
                .ok => {},
                .refused => |r| return .{ .refused = r },
            }
        } else if (el_depth == 2 and std.mem.eql(u8, local, "workbookPr")) {
            switch (parseWorkbookPr(el, &out)) {
                .ok => {},
                .refused => |r| return .{ .refused = r },
            }
        } else if (el_depth == 2 and std.mem.eql(u8, local, "extLst")) {
            if (!self_closing) extlst_depth = depth;
        } else if (extlst_depth != null and el_depth == extlst_depth.? + 1 and
            std.mem.eql(u8, local, "ext"))
        {
            const uri = el.attr("uri") orelse "";
            if (self_closing) {
                try exts.append(gpa, .{ .uri = uri, .raw = xml[el.offset..sc.i] });
            } else {
                ext_open = .{ .start = el.offset, .depth = depth, .uri = uri };
                in_calc_features = std.mem.eql(u8, uri, calc_features_ext_uri);
            }
        } else if (in_calc_features and std.mem.eql(u8, local, "feature")) {
            if (el.attr("name")) |n| try features.append(gpa, n);
        }
    }

    out.extensions = try exts.toOwnedSlice(gpa);
    errdefer gpa.free(out.extensions);
    out.calc_features = try features.toOwnedSlice(gpa);
    out.text_compat = textCompatOf(out.calc_features, &calc_feature_inventory);
    return .{ .ok = out };
}

/// §5.4d's version, from the feature list. Takes the table as a
/// parameter so the mapping is testable without waiting for the row
/// that pins CV2's spelling.
pub fn textCompatOf(
    features: []const []const u8,
    table: []const CalcFeatureRow,
) TextCompat {
    for (features) |f| {
        for (table) |row| {
            if (!std.mem.eql(u8, row.name, f)) continue;
            if (row.compat) |c| return c;
        }
    }
    return .v1;
}

const CalcParse = union(enum) { ok, refused: Refusal };

fn parseCalcPr(el: decode.Element, self_closing: bool, out: *CalcState) CalcParse {
    out.present = true;
    out.raw_attrs = el.attrs;
    out.self_closing = self_closing;

    var it = el.attrIterator();
    while (it.next()) |a| {
        // Unknown attributes are PRESERVED, not refused — `raw_attrs`
        // carries them back out. `<calcPr>` says how a workbook is
        // calculated, not what its cells contain, so an attribute this
        // engine does not read cannot silently change a value the way an
        // unread `<f>` attribute can.
        if (a.prefix().len != 0) continue;
        const name = a.local();
        const v = a.raw_value;
        if (std.mem.eql(u8, name, "calcId")) {
            out.calc_id = std.fmt.parseInt(u32, v, 10) catch return badCalc();
        } else if (std.mem.eql(u8, name, "calcMode")) {
            out.calc_mode = CalcMode.fromAttr(v) orelse return badCalc();
        } else if (std.mem.eql(u8, name, "fullCalcOnLoad")) {
            out.full_calc_on_load = xsdBool(v) orelse return badCalc();
        } else if (std.mem.eql(u8, name, "refMode")) {
            out.ref_mode = RefMode.fromAttr(v) orelse return badCalc();
        } else if (std.mem.eql(u8, name, "iterate")) {
            out.iterate = xsdBool(v) orelse return badCalc();
        } else if (std.mem.eql(u8, name, "iterateCount")) {
            out.iterate_count = std.fmt.parseInt(u32, v, 10) catch return badCalc();
        } else if (std.mem.eql(u8, name, "iterateDelta")) {
            out.iterate_delta = std.fmt.parseFloat(f64, v) catch return badCalc();
            if (!std.math.isFinite(out.iterate_delta)) return badCalc();
        } else if (std.mem.eql(u8, name, "fullPrecision")) {
            out.full_precision = xsdBool(v) orelse return badCalc();
            // §5.7.6: refused through v1. Every stored value would have
            // to be rounded to its number format before it is read, and
            // the number-format engine is M8a.
            if (!out.full_precision) {
                return .{ .refused = .{ .reason = .precision_as_displayed, .offset = offsetOf(el.offset) } };
            }
        } else if (std.mem.eql(u8, name, "calcCompleted")) {
            out.calc_completed = xsdBool(v) orelse return badCalc();
        } else if (std.mem.eql(u8, name, "calcOnSave")) {
            out.calc_on_save = xsdBool(v) orelse return badCalc();
        } else if (std.mem.eql(u8, name, "concurrentCalc")) {
            out.concurrent_calc = xsdBool(v) orelse return badCalc();
        } else if (std.mem.eql(u8, name, "concurrentManualCount")) {
            out.concurrent_manual_count = std.fmt.parseInt(u32, v, 10) catch return badCalc();
        } else if (std.mem.eql(u8, name, "forceFullCalc")) {
            out.force_full_calc = xsdBool(v) orelse return badCalc();
        }
    }
    return .ok;
}

fn parseWorkbookPr(el: decode.Element, out: *CalcState) CalcParse {
    out.workbook_pr_raw = el.attrs;
    if (el.attr("date1904")) |v| {
        const b = xsdBool(v) orelse return badCalc();
        out.date_system = if (b) .d1904 else .d1900;
    }
    return .ok;
}

fn badCalc() CalcParse {
    return .{ .refused = .{ .reason = .bad_calc_attribute_value } };
}

fn offsetOf(i: usize) u32 {
    return std.math.cast(u32, i) orelse std.math.maxInt(u32);
}

// ─── the worksheet's half (§5.7.6) ───────────────────────────────

/// `CT_SheetCalcPr`. One attribute, and v1 preserves it rather than
/// acting on it — same provenance policy as `calcPr`'s.
pub const SheetCalcPr = struct {
    present: bool = false,
    full_calc_on_load: bool = false,
    raw_attrs: []const u8 = "",
    self_closing: bool = true,

    pub fn write(self: SheetCalcPr, w: *std.Io.Writer) std.Io.Writer.Error!void {
        if (!self.present) return;
        try w.writeAll("<sheetCalcPr");
        try w.writeAll(self.raw_attrs);
        try w.writeAll(if (self.self_closing) "/>" else "></sheetCalcPr>");
    }
};

pub const SheetCalcResult = union(enum) {
    ok: SheetCalcPr,
    refused: Refusal,
};

/// Find `<sheetCalcPr>` in a worksheet part.
///
/// It is a direct child of `<worksheet>`, which is the only place the
/// schema puts it — so a `<sheetCalcPr>` nested somewhere else is
/// foreign content and is not read.
pub fn parseSheetCalcPr(xml: []const u8) SheetCalcResult {
    var sc = decode.Scanner.init(xml);
    var depth: usize = 0;
    while (sc.next() catch {
        return .{ .refused = .{ .reason = .malformed_calc_part } };
    }) |ev| {
        switch (ev) {
            .close => {
                if (depth == 0) return .{ .refused = .{ .reason = .malformed_calc_part } };
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
        const el_depth = if (self_closing) depth + 1 else depth;
        if (el_depth != 2 or !std.mem.eql(u8, el.local(), "sheetCalcPr")) continue;

        var out: SheetCalcPr = .{
            .present = true,
            .raw_attrs = el.attrs,
            .self_closing = self_closing,
        };
        if (el.attr("fullCalcOnLoad")) |v| {
            out.full_calc_on_load = xsdBool(v) orelse return .{ .refused = .{
                .reason = .bad_calc_attribute_value,
                .offset = offsetOf(el.offset),
            } };
        }
        return .{ .ok = out };
    }
    return .{ .ok = .{} };
}

// ─── tests ───────────────────────────────────────────────────────

const testing = std.testing;

const ns_attr = " xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"";

fn sheetXml(comptime body: []const u8) []const u8 {
    return "<worksheet" ++ ns_attr ++ "><sheetData>" ++ body ++ "</sheetData></worksheet>";
}

fn scanCells(xml: []const u8) !decode.Sheet {
    return switch (try decode.scanSheet(testing.allocator, xml, &.{}, .{})) {
        .ok => |s| s,
        .refused => |r| {
            std.debug.print("decode refused: {any}\n", .{r});
            return error.TestUnexpectedRefusal;
        },
    };
}

/// One `<f>` in one cell, classified. The `<f>` reaches
/// `classifyFormula` through the real decode boundary rather than a
/// hand-built `decode.Formula`, so the attribute region under test is
/// the one a part actually produces — and the sheet travels with the
/// result, because every text slice in a `CellFormula` borrows it.
const One = struct {
    sheet: decode.Sheet,
    result: FormulaResult,

    fn deinit(self: *One) void {
        self.sheet.deinit();
        self.* = undefined;
    }
};

fn classifyOne(comptime f: []const u8) !One {
    var sheet = try scanCells(sheetXml("<row r=\"1\"><c r=\"A1\">" ++ f ++ "</c></row>"));
    errdefer sheet.deinit();
    try testing.expectEqual(@as(usize, 1), sheet.cells.len);
    return .{
        .sheet = sheet,
        .result = classifyFormula(sheet.cells[0].formula.?, .{ .row = 1, .col = 0 }),
    };
}

// ─── the `CT_CellFormula` attribute inventory ────────────────────

const Expect = union(enum) {
    ok,
    refused: Refusal.Reason,
};

const AttrFixture = struct {
    /// The inventory row this fixture exercises. `""` for the rows that
    /// are *not* in the inventory — the unknown-attribute cases.
    row: []const u8,
    f: []const u8,
    expect: Expect,
    note: []const u8 = "",
};

/// One fixture per inventory row, refusing rows included. The test
/// below also proves the coverage: a row nothing exercises fails.
const attr_fixtures = [_]AttrFixture{
    // `t` — the four values plus a fifth that is not one.
    .{ .row = "t", .f = "<f>A1+1</f>", .expect = .ok, .note = "absent is normal" },
    .{ .row = "t", .f = "<f t=\"normal\">A1+1</f>", .expect = .ok },
    .{ .row = "t", .f = "<f t=\"array\" ref=\"A1:B2\">SUM(C1:C2)</f>", .expect = .ok },
    .{ .row = "t", .f = "<f t=\"shared\" ref=\"A1:A5\" si=\"0\">A2+1</f>", .expect = .ok },
    .{
        .row = "t",
        .f = "<f t=\"dataTable\" ref=\"A1:B2\">1</f>",
        .expect = .{ .refused = .data_table_formula },
    },
    .{
        .row = "t",
        .f = "<f t=\"bogus\">1</f>",
        .expect = .{ .refused = .bad_formula_attribute_value },
    },

    // `ref`
    .{ .row = "ref", .f = "<f t=\"array\" ref=\"$A$1:$B$2\">1</f>", .expect = .ok },
    .{
        .row = "ref",
        .f = "<f t=\"array\" ref=\"not-a-range\">1</f>",
        .expect = .{ .refused = .bad_formula_attribute_value },
    },
    .{
        .row = "ref",
        .f = "<f t=\"array\">1</f>",
        .expect = .{ .refused = .array_missing_ref },
    },

    // `si`
    .{ .row = "si", .f = "<f t=\"shared\" si=\"7\"/>", .expect = .ok, .note = "bodiless slave" },
    .{
        .row = "si",
        .f = "<f t=\"shared\" si=\"-1\" ref=\"A1:A2\">1</f>",
        .expect = .{ .refused = .bad_formula_attribute_value },
    },
    .{
        .row = "si",
        .f = "<f t=\"shared\" ref=\"A1:A2\">1</f>",
        .expect = .{ .refused = .shared_missing_si },
    },

    // `ca` — the dirty hint, on the ordinary arithmetic the corpus
    // carries it on.
    .{ .row = "ca", .f = "<f ca=\"1\">A1+1</f>", .expect = .ok },
    .{ .row = "ca", .f = "<f ca=\"false\">A1+1</f>", .expect = .ok },
    .{
        .row = "ca",
        .f = "<f ca=\"yes\">A1+1</f>",
        .expect = .{ .refused = .bad_formula_attribute_value },
    },

    // `aca` — granularity, both values accepted.
    .{ .row = "aca", .f = "<f t=\"array\" ref=\"A1:A2\" aca=\"1\">1</f>", .expect = .ok },
    .{ .row = "aca", .f = "<f t=\"array\" ref=\"A1:A2\" aca=\"false\">1</f>", .expect = .ok },
    .{
        .row = "aca",
        .f = "<f aca=\"maybe\">1</f>",
        .expect = .{ .refused = .bad_formula_attribute_value },
    },

    // `bx` — the conditional row: one value accepted, one refused.
    .{ .row = "bx", .f = "<f bx=\"0\">A1</f>", .expect = .ok },
    .{ .row = "bx", .f = "<f bx=\"false\">A1</f>", .expect = .ok },
    .{ .row = "bx", .f = "<f bx=\"true\">A1</f>", .expect = .{ .refused = .box_formula } },
    .{ .row = "bx", .f = "<f bx=\"1\">A1</f>", .expect = .{ .refused = .box_formula } },
    .{
        .row = "bx",
        .f = "<f bx=\"x\">A1</f>",
        .expect = .{ .refused = .bad_formula_attribute_value },
    },

    // The six data-table attributes, each on its own.
    .{ .row = "dt2D", .f = "<f dt2D=\"1\">1</f>", .expect = .{ .refused = .data_table_formula } },
    .{ .row = "dtr", .f = "<f dtr=\"1\">1</f>", .expect = .{ .refused = .data_table_formula } },
    .{ .row = "del1", .f = "<f del1=\"1\">1</f>", .expect = .{ .refused = .data_table_formula } },
    .{ .row = "del2", .f = "<f del2=\"1\">1</f>", .expect = .{ .refused = .data_table_formula } },
    .{ .row = "r1", .f = "<f r1=\"A1\">1</f>", .expect = .{ .refused = .data_table_formula } },
    .{ .row = "r2", .f = "<f r2=\"B1\">1</f>", .expect = .{ .refused = .data_table_formula } },

    // `xml:space` — the one prefixed row.
    .{ .row = "xml:space", .f = "<f xml:space=\"preserve\">A1</f>", .expect = .ok },
    .{ .row = "xml:space", .f = "<f xml:space=\"default\">A1</f>", .expect = .ok },
    .{
        .row = "xml:space",
        .f = "<f xml:space=\"squash\">A1</f>",
        .expect = .{ .refused = .bad_formula_attribute_value },
    },

    // Not in the inventory at all.
    .{
        .row = "",
        .f = "<f frobnicate=\"1\">A1</f>",
        .expect = .{ .refused = .unknown_formula_attribute },
    },
    .{
        .row = "",
        .f = "<f space=\"preserve\">A1</f>",
        .expect = .{ .refused = .unknown_formula_attribute },
        .note = "`xml:space` unqualified is a different attribute",
    },
    .{
        .row = "",
        .f = "<f x14ac:dyDescent=\"0.2\">A1</f>",
        .expect = .ok,
        .note = "foreign-namespace attributes are exempt (M4b1 decision 8)",
    },
};

test "inventory: every `CT_CellFormula` attribute has a fixture, and it holds" {
    inline for (attr_fixtures) |fx| {
        var one = try classifyOne(fx.f);
        defer one.deinit();
        const got = one.result;
        switch (fx.expect) {
            .ok => switch (got) {
                .ok => {},
                .refused => |r| {
                    std.debug.print("{s}: unexpected refusal {t}\n", .{ fx.f, r.reason });
                    return error.TestUnexpectedRefusal;
                },
            },
            .refused => |want| switch (got) {
                .ok => {
                    std.debug.print("{s}: expected {t}, got acceptance\n", .{ fx.f, want });
                    return error.TestExpectedRefusal;
                },
                .refused => |r| {
                    try testing.expectEqual(want, r.reason);
                    // Every refusal names the cell it happened at.
                    try testing.expectEqual(@as(u32, 1), r.cell.?.row);
                    _ = r.planeTwo();
                },
            },
        }
    }

    // Coverage, machine-checked: the inventory is the authority, so a
    // row added without a fixture fails here rather than shipping
    // untested.
    for (formula_attrs) |row| {
        var covered = false;
        for (attr_fixtures) |fx| {
            if (std.mem.eql(u8, fx.row, row.name)) covered = true;
        }
        if (!covered) {
            std.debug.print("no fixture for `{s}`\n", .{row.name});
            return error.UncoveredInventoryRow;
        }
    }
}

test "inventory: the typed view keeps what a round-trip owes back" {
    var one = try classifyOne("<f t=\"shared\" ref=\"A3:A7\" si=\"1\" ca=\"1\" bx=\"0\" xml:space=\"preserve\">A2+1</f>");
    defer one.deinit();
    const f = one.result.ok;
    try testing.expectEqual(Kind.shared, f.kind);
    try testing.expectEqual(@as(?u32, 1), f.si);
    try testing.expectEqualStrings("A3:A7", f.ref_raw.?);
    try testing.expectEqual(@as(u32, 3), f.ref.?.first.row.oneBased());
    try testing.expectEqual(@as(u32, 7), f.ref.?.last.row.oneBased());
    try testing.expect(f.ca);
    try testing.expect(!f.aca);
    try testing.expect(!f.bx);
    try testing.expectEqual(XmlSpace.preserve, f.xml_space);
    try testing.expectEqualStrings("A2+1", f.text);
    // The whole raw region survives, in the order it was written.
    try testing.expect(std.mem.indexOf(u8, f.raw_attrs, "ca=\"1\"") != null);
    try testing.expect(f.isSharedMaster());
    try testing.expect(!f.isSharedSlave());
}

test "inventory: `ca` is a dirty hint and `aca` is granularity" {
    // Neither is volatility. `ca` says nothing about scheduling, so the
    // only thing that reads it is a round-trip; `aca` is observable in
    // exactly one place, and that place refuses.
    var dirty_one = try classifyOne("<f ca=\"1\">A1+1</f>");
    defer dirty_one.deinit();
    const dirty = dirty_one.result.ok;
    try testing.expect(dirty.ca);
    try testing.expectEqual(@as(?Refusal.Reason, null), dirty.cycleRefusal());

    var whole_one = try classifyOne("<f t=\"array\" ref=\"A1:A2\" aca=\"1\">A1+1</f>");
    defer whole_one.deinit();
    const whole = whole_one.result.ok;
    try testing.expect(whole.aca);
    try testing.expectEqual(@as(?Refusal.Reason, null), whole.cycleRefusal());

    // Per-cell granularity inside a cycle is what refuses — and `aca`'s
    // schema default is false, so an array with no `aca` reads that way.
    var per_cell_one = try classifyOne("<f t=\"array\" ref=\"A1:A2\" aca=\"0\">A1+1</f>");
    defer per_cell_one.deinit();
    try testing.expectEqual(
        @as(?Refusal.Reason, .cse_array_per_cell_in_cycle),
        per_cell_one.result.ok.cycleRefusal(),
    );
    var defaulted_one = try classifyOne("<f t=\"array\" ref=\"A1:A2\">A1+1</f>");
    defer defaulted_one.deinit();
    try testing.expectEqual(
        @as(?Refusal.Reason, .cse_array_per_cell_in_cycle),
        defaulted_one.result.ok.cycleRefusal(),
    );
    // A non-array is never subject to the rule, whatever `aca` says.
    var scalar_one = try classifyOne("<f aca=\"0\">A1+1</f>");
    defer scalar_one.deinit();
    try testing.expectEqual(@as(?Refusal.Reason, null), scalar_one.result.ok.cycleRefusal());
}

test "inventory: every reason has a plane, and the planes are the right ones" {
    inline for (@typeInfo(Refusal.Reason).@"enum".fields) |fld| {
        const r: Refusal = .{ .reason = @enumFromInt(fld.value) };
        _ = r.planeTwo();
    }
    try testing.expectEqual(
        PlaneTwo.FormulaDataTableUnsupported,
        (Refusal{ .reason = .data_table_formula }).planeTwo(),
    );
    try testing.expectEqual(
        PlaneTwo.FormulaPrecisionAsDisplayed,
        (Refusal{ .reason = .precision_as_displayed }).planeTwo(),
    );
    try testing.expectEqual(
        PlaneTwo.FormulaUnsupportedConstruct,
        (Refusal{ .reason = .box_formula }).planeTwo(),
    );
}

// ─── sheet-wide shared topology ──────────────────────────────────

fn classifySheetXml(comptime body: []const u8) !struct {
    sheet: decode.Sheet,
    result: SheetResult,

    const Self = @This();
    fn deinit(self: *Self) void {
        switch (self.result) {
            .ok => |*ok| ok.deinit(),
            .refused => {},
        }
        self.sheet.deinit();
        self.* = undefined;
    }
} {
    var sheet = try scanCells(sheetXml(body));
    errdefer sheet.deinit();
    return .{ .sheet = sheet, .result = try classifySheet(testing.allocator, sheet.cells) };
}

fn expectTopologyRefusal(comptime body: []const u8, want: Refusal.Reason) !void {
    var c = try classifySheetXml(body);
    defer c.deinit();
    switch (c.result) {
        .ok => return error.TestExpectedRefusal,
        .refused => |r| try testing.expectEqual(want, r.reason),
    }
}

test "topology: a slave is recognized by its ATTRIBUTES, not by its shape" {
    // Both shapes in one sheet, and both must land in the group. The
    // empty-body form is what `tests/corpus/calamine_non_monotonic_si.xlsx`
    // writes and what `xlsx.zig:2099` drops; the self-closing form is
    // what Excel writes.
    var c = try classifySheetXml(
        \\<row r="1"><c r="A1"><f t="shared" ref="A1:A4" si="0">B1+1</f><v>1</v></c></row>
        \\<row r="2"><c r="A2"><f t="shared" si="0"/><v>2</v></c></row>
        \\<row r="3"><c r="A3"><f t="shared" si="0"></f><v>3</v></c></row>
        \\<row r="4"><c r="A4"><f si="0" t="shared"></f><v>4</v></c></row>
    );
    defer c.deinit();
    const ok = c.result.ok;
    try testing.expectEqual(@as(usize, 4), ok.entries.len);
    try testing.expectEqual(@as(usize, 1), ok.groups.len);
    try testing.expectEqual(@as(u32, 0), ok.groups[0].si);
    try testing.expectEqualStrings("B1+1", ok.groups[0].text);

    switch (ok.entries[0].role) {
        .master => |g| try testing.expectEqual(@as(u32, 0), g),
        else => return error.WrongRole,
    }
    // Three slaves, one row apart each, in every shape the format has.
    for (ok.entries[1..], 1..) |e, i| {
        switch (e.role) {
            .slave => |s| {
                try testing.expectEqual(@as(u32, 0), s.group);
                try testing.expectEqual(@as(i32, @intCast(i)), s.delta.rows);
                try testing.expectEqual(@as(i32, 0), s.delta.cols);
            },
            else => return error.WrongRole,
        }
    }
}

test "topology: a slave with a non-empty body still translates from its master" {
    // ECMA-376 requires a slave's body, when written, to *be* the
    // translated formula — so the two agree by construction and the
    // master is the single source of truth. The body is preserved and
    // read by nothing, which is what keeps one parse per group rather
    // than one per cell.
    var c = try classifySheetXml(
        \\<row r="1"><c r="A1"><f t="shared" ref="A1:A2" si="0">B1+1</f></c></row>
        \\<row r="2"><c r="A2"><f t="shared" si="0">B2+1</f></c></row>
    );
    defer c.deinit();
    const ok = c.result.ok;
    try testing.expectEqualStrings("B2+1", ok.entries[1].formula.text);
    switch (ok.entries[1].role) {
        .slave => |s| {
            try testing.expectEqualStrings("B1+1", ok.groups[s.group].text);
            try testing.expectEqual(@as(i32, 1), s.delta.rows);
        },
        else => return error.WrongRole,
    }
}

test "topology: the corpus workbook's shape — non-monotonic si, ref past the cells" {
    // `tests/corpus/calamine_non_monotonic_si.xlsx`, reduced: masters
    // in the order 1, 0, 2, and `ref="A3:A7"` naming two rows that have
    // no cell. `si` is a key, `ref` is not a gate.
    var c = try classifySheetXml(
        \\<row r="3"><c r="A3"><f ref="A3:A7" si="1" t="shared">A2+1</f></c><c r="B3"><f ref="B3:B6" si="0" t="shared">B2+2</f></c><c r="C3"><f ref="C3:C6" si="2" t="shared">C2+3</f></c></row>
        \\<row r="4"><c r="A4"><f si="1" t="shared"></f></c><c r="B4"><f si="0" t="shared"></f></c><c r="C4"><f si="2" t="shared"></f></c></row>
    );
    defer c.deinit();
    const ok = c.result.ok;
    try testing.expectEqual(@as(usize, 3), ok.groups.len);
    try testing.expectEqual(@as(u32, 1), ok.groups[0].si);
    try testing.expectEqual(@as(u32, 0), ok.groups[1].si);
    try testing.expectEqual(@as(u32, 2), ok.groups[2].si);
    // The master's own coordinate anchors the group, and `ref` reaches
    // past the last occupied row without that mattering.
    try testing.expectEqual(@as(u32, 3), ok.groups[0].anchor.row.oneBased());
    try testing.expectEqual(@as(u32, 7), ok.groups[0].ref.last.row.oneBased());
    for (ok.entries[3..]) |e| {
        switch (e.role) {
            .slave => |s| try testing.expectEqual(@as(i32, 1), s.delta.rows),
            else => return error.WrongRole,
        }
    }
}

test "topology: a 2-D master, and the deltas across it" {
    // `tests/corpus/openxlsx_loadExample.xlsx` shares a rectangle,
    // `ref="I2:N5"` with `ca="1"` throughout, so a slave's delta has
    // both components.
    var c = try classifySheetXml(
        \\<row r="2"><c r="I2"><f t="shared" ref="I2:N5" ca="1" si="0">A2*2</f></c><c r="J2"><f t="shared" ca="1" si="0"/></c></row>
        \\<row r="3"><c r="K3"><f t="shared" ca="1" si="0"/></c></row>
    );
    defer c.deinit();
    const ok = c.result.ok;
    try testing.expect(ok.entries[0].formula.ca);
    switch (ok.entries[1].role) {
        .slave => |s| {
            try testing.expectEqual(@as(i32, 0), s.delta.rows);
            try testing.expectEqual(@as(i32, 1), s.delta.cols);
        },
        else => return error.WrongRole,
    }
    switch (ok.entries[2].role) {
        .slave => |s| {
            try testing.expectEqual(@as(i32, 1), s.delta.rows);
            try testing.expectEqual(@as(i32, 2), s.delta.cols);
        },
        else => return error.WrongRole,
    }
}

test "topology: the four ways a shared group can be malformed" {
    // A slave before its master. Legal-looking XML, and unreadable:
    // the formula it is defined by has not been given yet.
    try expectTopologyRefusal(
        \\<row r="1"><c r="A1"><f t="shared" si="0"/></c></row>
        \\<row r="2"><c r="A2"><f t="shared" ref="A2:A4" si="0">B2+1</f></c></row>
    , .shared_slave_before_master);

    // A slave whose `si` names no master at all.
    try expectTopologyRefusal(
        \\<row r="1"><c r="A1"><f t="shared" ref="A1:A2" si="0">B1+1</f></c></row>
        \\<row r="2"><c r="A2"><f t="shared" si="9"/></c></row>
    , .shared_unknown_si);

    // Two masters claiming one `si`.
    try expectTopologyRefusal(
        \\<row r="1"><c r="A1"><f t="shared" ref="A1:A2" si="0">B1+1</f></c></row>
        \\<row r="2"><c r="A2"><f t="shared" ref="A2:A3" si="0">B2+9</f></c></row>
    , .shared_duplicate_si);

    // A master with nothing to share.
    try expectTopologyRefusal(
        \\<row r="1"><c r="A1"><f t="shared" ref="A1:A2" si="0"></f></c></row>
    , .shared_master_empty);
}

test "topology: non-shared formulas keep their own roles" {
    var c = try classifySheetXml(
        \\<row r="1"><c r="A1"><f>B1+1</f></c><c r="B1"><f t="array" ref="B1:C2">TRANSPOSE(D1:E2)</f></c><c r="C1"><v>3</v></c></row>
    );
    defer c.deinit();
    const ok = c.result.ok;
    // The value-only cell contributes no entry at all.
    try testing.expectEqual(@as(usize, 2), ok.entries.len);
    try testing.expectEqual(@as(usize, 0), ok.groups.len);
    try testing.expectEqual(Role.normal, ok.entries[0].role);
    switch (ok.entries[1].role) {
        .array => |r| try testing.expectEqual(@as(u32, 2), r.last.col.oneBased() - r.first.col.oneBased() + 1),
        else => return error.WrongRole,
    }
}

test "topology: a refused `<f>` refuses the sheet, at the cell that carried it" {
    var c = try classifySheetXml(
        \\<row r="1"><c r="A1"><f>B1+1</f></c></row>
        \\<row r="7"><c r="D7"><f dt2D="1">1</f></c></row>
    );
    defer c.deinit();
    switch (c.result) {
        .ok => return error.TestExpectedRefusal,
        .refused => |r| {
            try testing.expectEqual(Refusal.Reason.data_table_formula, r.reason);
            try testing.expectEqual(@as(u32, 7), r.cell.?.row);
            try testing.expectEqual(@as(u32, 3), r.cell.?.col);
        },
    }
}

// ─── the translation matrix ──────────────────────────────────────

fn expectTranslation(src: []const u8, d: Delta, want: []const u8) !void {
    var t = switch (try translateText(testing.allocator, src, d)) {
        .ok => |v| v,
        .refused => |r| {
            std.debug.print("{s}: parse refused {t}\n", .{ src, r.reason });
            return error.TestUnexpectedRefusal;
        },
    };
    defer t.deinit();
    testing.expectEqualStrings(want, t.text) catch |e| {
        std.debug.print("  translating `{s}` by ({d},{d})\n", .{ src, d.rows, d.cols });
        return e;
    };
}

const TranslationFixture = struct {
    /// The reference shape this row covers.
    shape: []const u8,
    src: []const u8,
    d: Delta,
    want: []const u8,
};

/// One row per reference shape §5.7.1 names, plus the off-grid column.
const translation_fixtures = [_]TranslationFixture{
    .{ .shape = "relative cell", .src = "A1", .d = .{ .rows = 1, .cols = 1 }, .want = "B2" },
    .{ .shape = "absolute cell", .src = "$A$1", .d = .{ .rows = 1, .cols = 1 }, .want = "$A$1" },
    .{ .shape = "mixed: $col", .src = "$A1", .d = .{ .rows = 1, .cols = 1 }, .want = "$A2" },
    .{ .shape = "mixed: $row", .src = "A$1", .d = .{ .rows = 1, .cols = 1 }, .want = "B$1" },
    .{ .shape = "range", .src = "A1:B2", .d = .{ .rows = 2 }, .want = "A3:B4" },
    .{
        .shape = "range, mixed anchors per corner",
        .src = "$A1:B$2",
        .d = .{ .rows = 3, .cols = 4 },
        .want = "$A4:F$2",
    },
    .{ .shape = "full column", .src = "A:A", .d = .{ .cols = 1 }, .want = "B:B" },
    .{ .shape = "full column, absolute", .src = "$A:$B", .d = .{ .cols = 5 }, .want = "$A:$B" },
    .{ .shape = "full column, mixed", .src = "$A:B", .d = .{ .cols = 2 }, .want = "$A:D" },
    .{ .shape = "full row", .src = "1:5", .d = .{ .rows = 2 }, .want = "3:7" },
    .{ .shape = "full row, absolute", .src = "$1:$5", .d = .{ .rows = 2 }, .want = "$1:$5" },
    .{
        .shape = "3D span",
        .src = "Sheet1:Sheet3!A1",
        .d = .{ .rows = 1 },
        .want = "Sheet1:Sheet3!A2",
    },
    .{
        .shape = "3D span, quoted",
        .src = "'Q1:Q4'!A1:B2",
        .d = .{ .rows = 1, .cols = 1 },
        .want = "'Q1:Q4'!B2:C3",
    },
    .{ .shape = "qualified", .src = "Sheet1!A1", .d = .{ .rows = 1 }, .want = "Sheet1!A2" },
    .{ .shape = "spill ref", .src = "A1#", .d = .{ .rows = 1, .cols = 1 }, .want = "B2#" },
    .{
        .shape = "refs in lazy branches",
        .src = "IF(A1,B1,C1)",
        .d = .{ .cols = 1 },
        .want = "IF(B1,C1,D1)",
    },
    .{
        .shape = "refs under a short-circuit, nested",
        .src = "IF(A1,IF(B2,C3,D4),E5)",
        .d = .{ .rows = 10, .cols = 1 },
        .want = "IF(B11,IF(C12,D13,E14),F15)",
    },
    .{
        .shape = "structured reference — name-based, never shifts",
        .src = "Table1[@Col]",
        .d = .{ .rows = 4, .cols = 4 },
        .want = "Table1[@Col]",
    },
    .{
        .shape = "defined name — never shifts",
        .src = "MyName+1",
        .d = .{ .rows = 4, .cols = 4 },
        .want = "MyName+1",
    },
    .{
        .shape = "array constant — no coordinates in it",
        .src = "{1,2;3,4}",
        .d = .{ .rows = 4, .cols = 4 },
        .want = "{1,2;3,4}",
    },
    .{
        .shape = "string literal that looks like a reference",
        .src = "\"A1\"&A1",
        .d = .{ .rows = 1 },
        .want = "\"A1\"&A2",
    },
    .{ .shape = "postfix percent", .src = "A1%", .d = .{ .rows = 1 }, .want = "A2%" },
    .{ .shape = "unary minus", .src = "-A1", .d = .{ .cols = 1 }, .want = "-B1" },
    .{
        .shape = "intersection operator",
        .src = "A1:A5 B1:D1",
        .d = .{ .rows = 1 },
        .want = "A2:A6 B2:D2",
    },
    .{
        .shape = "union operator",
        .src = "SUM((A1,B2))",
        .d = .{ .rows = 1 },
        .want = "SUM((A2,B3))",
    },
    .{
        .shape = "implicit intersection",
        .src = "@A1:A5",
        .d = .{ .rows = 1 },
        .want = "@A2:A6",
    },

    // Off-grid. A reference operand collapses WHOLE — the format has no
    // token for a half-valid area.
    .{ .shape = "off-grid: cell", .src = "A1", .d = .{ .rows = -1 }, .want = "#REF!" },
    .{ .shape = "off-grid: past XFD", .src = "XFD1", .d = .{ .cols = 1 }, .want = "#REF!" },
    .{
        .shape = "off-grid: past the last row",
        .src = "A1048576",
        .d = .{ .rows = 1 },
        .want = "#REF!",
    },
    .{
        .shape = "off-grid: one endpoint takes the whole range",
        .src = "SUM(A1:B2)",
        .d = .{ .rows = -1, .cols = -1 },
        .want = "SUM(#REF!)",
    },
    .{
        .shape = "off-grid: the qualifier survives",
        .src = "Sheet1!A1",
        .d = .{ .rows = -1 },
        .want = "Sheet1!#REF!",
    },
    .{ .shape = "off-grid: spill", .src = "A1#", .d = .{ .rows = -1 }, .want = "#REF!" },
    .{ .shape = "off-grid: full column", .src = "A:B", .d = .{ .cols = -1 }, .want = "#REF!" },
    .{ .shape = "off-grid: full row", .src = "1:5", .d = .{ .rows = -1 }, .want = "#REF!" },
    .{
        .shape = "off-grid: an absolute half is immune",
        .src = "$A1",
        .d = .{ .cols = -1 },
        .want = "$A1",
    },
    .{
        .shape = "off-grid: only the dead operand collapses",
        .src = "A1+B2",
        .d = .{ .rows = -1 },
        .want = "#REF!+B1",
    },
};

test "translation: one fixture per reference shape" {
    for (translation_fixtures) |fx| {
        try expectTranslation(fx.src, fx.d, fx.want);
    }
}

test "translation: a zero delta is the identity, on every fixture" {
    // The property that makes the matrix trustworthy: nothing moves
    // when nothing moved, byte for byte, including the spellings the
    // printer would otherwise normalize.
    for (translation_fixtures) |fx| {
        try expectTranslation(fx.src, .{}, fx.src);
    }
}

test "translation: a source `#REF!` is left where it is" {
    // Collapse is a consequence of *this* translation. A `#REF!` the
    // file already carried is not one, so a range around it survives.
    try expectTranslation("SUM(A1:#REF!)", .{}, "SUM(A1:#REF!)");
    try expectTranslation("SUM(A1:#REF!)", .{ .rows = 1 }, "SUM(A2:#REF!)");
    try expectTranslation("#REF!+1", .{ .rows = 3 }, "#REF!+1");
}

test "translation: every shape in the matrix is distinct" {
    // A fixture table earns its keep only if its rows differ; two rows
    // with one shape name would hide a missing case.
    for (translation_fixtures, 0..) |a, i| {
        for (translation_fixtures[i + 1 ..]) |b| {
            if (std.mem.eql(u8, a.shape, b.shape)) {
                std.debug.print("duplicate shape `{s}`\n", .{a.shape});
                return error.DuplicateFixtureShape;
            }
        }
    }
}

// ─── the reference translator, and the differential test ─────────

/// An independent translator, written the other way round on purpose:
/// this one scans the formula **text** for A1 references and shifts each
/// one, where `translate` walks a parsed tree and prints it. The two
/// share nothing but `coords`, which is what makes agreement on
/// thousands of randomized formulas evidence rather than a tautology.
///
/// It is deliberately naive — it would mis-handle a quoted sheet name or
/// a string literal containing something ref-shaped — so the generator
/// below produces only formulas whose text it can read. Every construct
/// it cannot is covered by a fixture in `translation_fixtures` instead.
fn referenceTranslate(gpa: Allocator, text: []const u8, d: Delta) error{OutOfMemory}![]u8 {
    var out: std.ArrayListUnmanaged(u8) = .empty;
    errdefer out.deinit(gpa);

    var i: usize = 0;
    while (i < text.len) {
        const before_ok = i == 0 or !isRefBody(text[i - 1]);
        if (before_ok) {
            if (matchRef(text[i..])) |m| {
                const after = i + m.len;
                const after_ok = after >= text.len or !isRefTail(text[after]);
                if (after_ok) {
                    var buf: [coords.format_buf_len]u8 = undefined;
                    if (shiftCell(m.cell, d)) |moved| {
                        try out.appendSlice(gpa, coords.formatCell(&buf, moved));
                    } else {
                        try out.appendSlice(gpa, ref_error);
                    }
                    i = after;
                    continue;
                }
            }
        }
        try out.append(gpa, text[i]);
        i += 1;
    }
    return out.toOwnedSlice(gpa);
}

fn isRefBody(c: u8) bool {
    return (c >= 'A' and c <= 'Z') or (c >= 'a' and c <= 'z') or
        (c >= '0' and c <= '9') or c == '_' or c == '.' or c == '$' or c == '!';
}

fn isRefTail(c: u8) bool {
    return (c >= 'A' and c <= 'Z') or (c >= 'a' and c <= 'z') or
        (c >= '0' and c <= '9') or c == '_' or c == '.' or c == '(';
}

fn matchRef(s: []const u8) ?struct { cell: coords.Cell, len: usize } {
    var i: usize = 0;
    if (i < s.len and s[i] == '$') i += 1;
    const letters_start = i;
    while (i < s.len and s[i] >= 'A' and s[i] <= 'Z') : (i += 1) {}
    if (i == letters_start or i - letters_start > coords.max_col_letters) return null;
    if (i < s.len and s[i] == '$') i += 1;
    const digits_start = i;
    while (i < s.len and s[i] >= '0' and s[i] <= '9') : (i += 1) {}
    if (i == digits_start) return null;
    const cell = coords.parseCell(s[0..i], .{
        .case = .upper_only,
        .dollar = .accept,
    }) catch return null;
    return .{ .cell = cell, .len = i };
}

/// A formula generator restricted to what `referenceTranslate` can read:
/// upper-case A1 references, numbers, arithmetic, ranges, and two
/// functions whose names carry no digits.
const Generator = struct {
    rnd: std.Random,
    buf: *std.ArrayListUnmanaged(u8),
    gpa: Allocator,

    fn expr(self: *Generator, depth: u8) error{OutOfMemory}!void {
        if (depth == 0 or self.rnd.uintLessThan(u8, 100) < 35) return self.leaf();
        switch (self.rnd.uintLessThan(u8, 5)) {
            0 => {
                try self.expr(depth - 1);
                try self.buf.appendSlice(self.gpa, switch (self.rnd.uintLessThan(u8, 4)) {
                    0 => "+",
                    1 => "-",
                    2 => "*",
                    else => "/",
                });
                try self.expr(depth - 1);
            },
            1 => {
                try self.buf.append(self.gpa, '(');
                try self.expr(depth - 1);
                try self.buf.append(self.gpa, ')');
            },
            2 => {
                try self.buf.appendSlice(self.gpa, "SUM(");
                try self.expr(depth - 1);
                try self.buf.append(self.gpa, ',');
                try self.expr(depth - 1);
                try self.buf.append(self.gpa, ')');
            },
            3 => {
                try self.buf.appendSlice(self.gpa, "IF(");
                try self.expr(depth - 1);
                try self.buf.append(self.gpa, ',');
                try self.expr(depth - 1);
                try self.buf.append(self.gpa, ',');
                try self.expr(depth - 1);
                try self.buf.append(self.gpa, ')');
            },
            else => {
                try self.ref();
                try self.buf.append(self.gpa, ':');
                try self.ref();
            },
        }
    }

    fn leaf(self: *Generator) error{OutOfMemory}!void {
        if (self.rnd.uintLessThan(u8, 4) == 0) {
            try self.buf.print(self.gpa, "{d}", .{self.rnd.uintLessThan(u16, 1000)});
            return;
        }
        try self.ref();
    }

    fn ref(self: *Generator) error{OutOfMemory}!void {
        // A middle band, so a delta of ±100 can never leave the grid —
        // the off-grid collapse has its own fixtures, and the reference
        // translator does not model it operand-wise.
        const col = coords.Col.fromZeroBased(500 + self.rnd.uintLessThan(u32, 1000)) catch
            unreachable;
        const row = coords.Row.fromOneBased(500 + self.rnd.uintLessThan(u32, 1000)) catch
            unreachable;
        var buf: [coords.format_buf_len]u8 = undefined;
        const s = coords.formatCell(&buf, .{
            .col = col,
            .row = row,
            .anchor = .{
                .col = self.rnd.boolean(),
                .row = self.rnd.boolean(),
            },
        });
        try self.buf.appendSlice(self.gpa, s);
    }
};

test "translation: randomized differential against an independent translator" {
    const a = testing.allocator;
    var prng = std.Random.DefaultPrng.init(0x4d34_6232_0000_0001);
    const rnd = prng.random();

    var src: std.ArrayListUnmanaged(u8) = .empty;
    defer src.deinit(a);

    var round: usize = 0;
    while (round < 2000) : (round += 1) {
        src.clearRetainingCapacity();
        var gen: Generator = .{ .rnd = rnd, .buf = &src, .gpa = a };
        try gen.expr(4);

        const d: Delta = .{
            .rows = rnd.intRangeAtMost(i32, -100, 100),
            .cols = rnd.intRangeAtMost(i32, -100, 100),
        };

        var mine = switch (try translateText(a, src.items, d)) {
            .ok => |v| v,
            .refused => |r| {
                std.debug.print("generated formula refused ({t}): {s}\n", .{ r.reason, src.items });
                return error.GeneratedFormulaRefused;
            },
        };
        defer mine.deinit();

        const theirs = try referenceTranslate(a, src.items, d);
        defer a.free(theirs);

        testing.expectEqualStrings(theirs, mine.text) catch |e| {
            std.debug.print("  source `{s}` by ({d},{d})\n", .{ src.items, d.rows, d.cols });
            return e;
        };

        // The generator's own spelling is canonical, so the identity
        // case is also a print round-trip — which is what rules out
        // "both implementations normalize the same way by accident".
        var identity = switch (try translateText(a, src.items, .{})) {
            .ok => |v| v,
            .refused => return error.GeneratedFormulaRefused,
        };
        defer identity.deinit();
        try testing.expectEqualStrings(src.items, identity.text);
    }
}

// ─── calc state ──────────────────────────────────────────────────

fn workbookXml(comptime body: []const u8) []const u8 {
    return "<workbook" ++ ns_attr ++ ">" ++ body ++ "</workbook>";
}

fn parseCalcOk(xml: []const u8) !CalcState {
    return switch (try parseCalcState(testing.allocator, xml)) {
        .ok => |s| s,
        .refused => |r| {
            std.debug.print("calc state refused: {t}\n", .{r.reason});
            return error.TestUnexpectedRefusal;
        },
    };
}

/// Re-emit the parsed `<calcPr>` and compare it with what was read.
/// Reconstructed from the parsed pieces, never echoed from a saved span
/// — an echo would prove only that `memcpy` works.
fn expectCalcPrRoundTrip(comptime element: []const u8) !void {
    var state = try parseCalcOk(workbookXml("<sheets/>" ++ element));
    defer state.deinit(testing.allocator);
    try testing.expect(state.present);

    var out: std.Io.Writer.Allocating = .init(testing.allocator);
    defer out.deinit();
    try state.writeCalcPr(&out.writer);
    try testing.expectEqualStrings(element, out.written());
}

/// Every distinct `<calcPr>` spelling in `tests/corpus/` (22 workbooks
/// carry one, in these twelve shapes). The corpus is the authority on
/// what a round-trip has to survive, including the two shapes a
/// reconstructing writer is most likely to get wrong: the empty element,
/// and `wdi_excel.xlsx`'s space before `/>`.
const corpus_calc_pr = [_][]const u8{
    "<calcPr calcId=\"181029\"/>",
    "<calcPr calcId=\"124519\"/>",
    "<calcPr calcId=\"145621\"/>",
    "<calcPr calcId=\"0\"/>",
    "<calcPr/>",
    "<calcPr calcId=\"191028\"/>",
    "<calcPr calcId=\"140000\" concurrentCalc=\"0\"/>",
    "<calcPr calcId=\"152511\"/>",
    "<calcPr calcId=\"125725\"/>",
    "<calcPr iterateCount=\"100\" refMode=\"A1\" iterate=\"false\" iterateDelta=\"0.001\"/>",
    "<calcPr calcId=\"152511\" calcOnSave=\"0\"/>",
    "<calcPr calcId=\"40001\" />",
};

test "calc state: every corpus `<calcPr>` shape round-trips byte-identically" {
    inline for (corpus_calc_pr) |element| {
        try expectCalcPrRoundTrip(element);
    }
    // …and so does the non-self-closing spelling, which the schema
    // permits even though no corpus workbook uses it.
    try expectCalcPrRoundTrip("<calcPr calcId=\"1\"></calcPr>");
}

test "calc state: an unknown attribute is preserved, not refused" {
    // The opposite policy to `<f>`'s, and deliberately: `<calcPr>` says
    // how a workbook is calculated, not what a cell contains, so an
    // attribute this engine does not read cannot silently change a
    // value. Refusing it would refuse files that open everywhere.
    try expectCalcPrRoundTrip("<calcPr calcId=\"1\" somethingNew=\"7\"/>");
    var state = try parseCalcOk(workbookXml("<calcPr calcId=\"1\" somethingNew=\"7\"/>"));
    defer state.deinit(testing.allocator);
    try testing.expectEqual(@as(u32, 1), state.calc_id);
}

test "calc state: the complete `CT_CalcPr` inventory reaches the typed view" {
    var state = try parseCalcOk(workbookXml(
        \\<calcPr calcId="191029" calcMode="manual" fullCalcOnLoad="1" refMode="R1C1"
        \\ iterate="true" iterateCount="50" iterateDelta="0.0001" fullPrecision="1"
        \\ calcCompleted="0" calcOnSave="0" concurrentCalc="0" concurrentManualCount="4"
        \\ forceFullCalc="1"/>
    ));
    defer state.deinit(testing.allocator);
    try testing.expectEqual(@as(u32, 191029), state.calc_id);
    try testing.expectEqual(CalcMode.manual, state.calc_mode);
    try testing.expect(state.full_calc_on_load);
    try testing.expectEqual(RefMode.r1c1, state.ref_mode);
    try testing.expect(state.iterate);
    try testing.expectEqual(@as(u32, 50), state.iterate_count);
    try testing.expect(@abs(state.iterate_delta - 0.0001) < 1e-12);
    try testing.expect(state.full_precision);
    try testing.expect(!state.calc_completed);
    try testing.expect(!state.calc_on_save);
    try testing.expect(!state.concurrent_calc);
    try testing.expectEqual(@as(?u32, 4), state.concurrent_manual_count);
    try testing.expect(state.force_full_calc);
}

test "calc state: absent `<calcPr>` is the schema's defaults, and writes nothing" {
    var state = try parseCalcOk(workbookXml("<sheets/>"));
    defer state.deinit(testing.allocator);
    try testing.expect(!state.present);
    try testing.expectEqual(CalcMode.auto, state.calc_mode);
    try testing.expectEqual(RefMode.a1, state.ref_mode);
    try testing.expectEqual(@as(u32, 100), state.iterate_count);
    try testing.expect(!state.iterate);
    try testing.expect(state.full_precision);

    var out: std.Io.Writer.Allocating = .init(testing.allocator);
    defer out.deinit();
    try state.writeCalcPr(&out.writer);
    try testing.expectEqualStrings("", out.written());
}

test "calc state: `fullPrecision=\"0\"` refuses through v1" {
    for ([_][]const u8{
        workbookXml("<calcPr fullPrecision=\"0\"/>"),
        workbookXml("<calcPr calcId=\"1\" fullPrecision=\"false\"/>"),
    }) |xml| {
        switch (try parseCalcState(testing.allocator, xml)) {
            .ok => |s| {
                var st = s;
                st.deinit(testing.allocator);
                return error.TestExpectedRefusal;
            },
            .refused => |r| {
                try testing.expectEqual(Refusal.Reason.precision_as_displayed, r.reason);
                try testing.expectEqual(PlaneTwo.FormulaPrecisionAsDisplayed, r.planeTwo());
            },
        }
    }
    // `fullPrecision="1"` is the default and passes through.
    var ok = try parseCalcOk(workbookXml("<calcPr fullPrecision=\"1\"/>"));
    defer ok.deinit(testing.allocator);
}

test "calc state: a value outside its type's lexical space refuses" {
    for ([_][]const u8{
        workbookXml("<calcPr calcMode=\"whenever\"/>"),
        workbookXml("<calcPr refMode=\"A1B2\"/>"),
        workbookXml("<calcPr calcId=\"lots\"/>"),
        workbookXml("<calcPr iterate=\"yes\"/>"),
        workbookXml("<calcPr iterateCount=\"many\"/>"),
        workbookXml("<calcPr iterateDelta=\"nan\"/>"),
        workbookXml("<calcPr concurrentManualCount=\"-1\"/>"),
        workbookXml("<workbookPr date1904=\"perhaps\"/>"),
    }) |xml| {
        switch (try parseCalcState(testing.allocator, xml)) {
            .ok => |s| {
                var st = s;
                st.deinit(testing.allocator);
                std.debug.print("accepted: {s}\n", .{xml});
                return error.TestExpectedRefusal;
            },
            .refused => |r| try testing.expectEqual(
                Refusal.Reason.bad_calc_attribute_value,
                r.reason,
            ),
        }
    }
}

test "calc state: `date1904` decides the epoch, in every corpus spelling" {
    // Three corpus workbooks write `date1904="false"` explicitly.
    var explicit_false = try parseCalcOk(workbookXml(
        \\<workbookPr backupFile="false" showObjects="all" date1904="false"/>
    ));
    defer explicit_false.deinit(testing.allocator);
    try testing.expectEqual(DateSystem.d1900, explicit_false.date_system);
    try testing.expect(explicit_false.workbook_pr_raw != null);

    inline for ([_][]const u8{ "1", "true" }) |spelling| {
        var s = try parseCalcOk(workbookXml("<workbookPr date1904=\"" ++ spelling ++ "\"/>"));
        defer s.deinit(testing.allocator);
        try testing.expectEqual(DateSystem.d1904, s.date_system);
    }

    // No `<workbookPr>` at all is the 1900 system, which is the default
    // every workbook without one has.
    var absent = try parseCalcOk(workbookXml("<sheets/>"));
    defer absent.deinit(testing.allocator);
    try testing.expectEqual(DateSystem.d1900, absent.date_system);
}

test "calc state: workbook extensions survive byte-exact, calcFeatures included" {
    const feature_ext =
        "<ext uri=\"" ++ calc_features_ext_uri ++ "\" xmlns:xcalcf=\"x\">" ++
        "<xcalcf:calcFeatures><xcalcf:feature name=\"microsoft.com:RD\"/>" ++
        "<xcalcf:feature name=\"microsoft.com:LAMBDA_WF\"/></xcalcf:calcFeatures></ext>";
    const other_ext = "<ext uri=\"{OTHER}\"><thing a=\"1\"/></ext>";

    var state = try parseCalcOk(workbookXml(
        "<calcPr calcId=\"1\"/><extLst>" ++ feature_ext ++ other_ext ++ "</extLst>",
    ));
    defer state.deinit(testing.allocator);

    try testing.expectEqual(@as(usize, 2), state.extensions.len);
    try testing.expectEqualStrings(calc_features_ext_uri, state.extensions[0].uri);
    try testing.expectEqualStrings(feature_ext, state.extensions[0].raw);
    try testing.expectEqualStrings(other_ext, state.extensions[1].raw);

    try testing.expectEqual(@as(usize, 2), state.calc_features.len);
    try testing.expectEqualStrings("microsoft.com:RD", state.calc_features[0]);
    try testing.expectEqualStrings("microsoft.com:LAMBDA_WF", state.calc_features[1]);

    // §5.4d: absent CV metadata is CV1, and no feature name in the
    // shipped inventory carries a version.
    try testing.expectEqual(TextCompat.v1, state.text_compat);
}

test "calc state: the compatibility version comes from the feature table" {
    // The mapping is exercised with a synthetic table rather than
    // waiting for the row that pins CV2's byte-exact spelling — §4's
    // policy for every Office-vs-schema difference is a byte-diffed
    // reference (M7b), and a guessed name enshrined in a test would be
    // exactly the invented expectation this ladder refuses to make.
    const synthetic = [_]CalcFeatureRow{
        .{ .name = "vendor:INERT", .note = "no version" },
        .{ .name = "vendor:CV2", .compat = .v2, .note = "the version marker" },
    };
    try testing.expectEqual(TextCompat.v1, textCompatOf(&.{}, &synthetic));
    try testing.expectEqual(TextCompat.v1, textCompatOf(&.{"vendor:INERT"}, &synthetic));
    try testing.expectEqual(TextCompat.v2, textCompatOf(&.{"vendor:CV2"}, &synthetic));
    try testing.expectEqual(
        TextCompat.v2,
        textCompatOf(&.{ "vendor:INERT", "vendor:CV2" }, &synthetic),
    );
    // An unrecognized name leaves the version alone rather than
    // guessing — the whole reason the table is a parameter.
    try testing.expectEqual(TextCompat.v1, textCompatOf(&.{"vendor:UNKNOWN"}, &synthetic));

    // Every shipped row is inert, and the names are distinct.
    for (calc_feature_inventory, 0..) |a, i| {
        try testing.expectEqual(@as(?TextCompat, null), a.compat);
        for (calc_feature_inventory[i + 1 ..]) |b| {
            try testing.expect(!std.mem.eql(u8, a.name, b.name));
        }
    }
}

test "calc state: only the workbook's own `<extLst>` is calc state" {
    // A `<definedName>`'s or a sheet's extension list is not, and
    // reading one would put foreign content in the calc state.
    var state = try parseCalcOk(workbookXml(
        \\<definedNames><definedName name="x"><extLst><ext uri="{NESTED}"/></extLst></definedName></definedNames>
    ));
    defer state.deinit(testing.allocator);
    try testing.expectEqual(@as(usize, 0), state.extensions.len);
}

test "calc state: `<sheetCalcPr>` is preserved, and round-trips" {
    const xml = "<worksheet" ++ ns_attr ++
        "><sheetCalcPr fullCalcOnLoad=\"1\"/><sheetData/></worksheet>";
    const got = switch (parseSheetCalcPr(xml)) {
        .ok => |s| s,
        .refused => return error.TestUnexpectedRefusal,
    };
    try testing.expect(got.present);
    try testing.expect(got.full_calc_on_load);

    var out: std.Io.Writer.Allocating = .init(testing.allocator);
    defer out.deinit();
    try got.write(&out.writer);
    try testing.expectEqualStrings("<sheetCalcPr fullCalcOnLoad=\"1\"/>", out.written());

    // Absent is the common case: no element, and nothing written.
    const none = switch (parseSheetCalcPr("<worksheet" ++ ns_attr ++ "><sheetData/></worksheet>")) {
        .ok => |s| s,
        .refused => return error.TestUnexpectedRefusal,
    };
    try testing.expect(!none.present);
    try testing.expect(!none.full_calc_on_load);

    // A bad value refuses rather than defaulting to false — the flag
    // decides whether Excel recalculates on open.
    switch (parseSheetCalcPr("<worksheet" ++ ns_attr ++
        "><sheetCalcPr fullCalcOnLoad=\"maybe\"/></worksheet>")) {
        .ok => return error.TestExpectedRefusal,
        .refused => |r| try testing.expectEqual(
            Refusal.Reason.bad_calc_attribute_value,
            r.reason,
        ),
    }
}

// ─── allocation failure (§8) ─────────────────────────────────────

test "checkAllAllocationFailures: nothing in this file leaks under OOM" {
    const H = struct {
        fn run(allocator: Allocator) !void {
            // 1. The sheet topology: a shared group with two shapes of
            //    slave, an array, and a plain formula, so both owned
            //    lists have to grow.
            const sheet_xml = "<worksheet" ++ ns_attr ++ "><sheetData>" ++
                "<row r=\"1\"><c r=\"A1\"><f t=\"shared\" ref=\"A1:A3\" si=\"0\">B1+1</f></c>" ++
                "<c r=\"B1\"><f t=\"array\" ref=\"B1:C2\">TRANSPOSE(D1:E2)</f></c></row>" ++
                "<row r=\"2\"><c r=\"A2\"><f t=\"shared\" si=\"0\"/></c></row>" ++
                "<row r=\"3\"><c r=\"A3\"><f t=\"shared\" si=\"0\"></f></c>" ++
                "<c r=\"B3\"><f>A3*2</f></c></row>" ++
                "</sheetData></worksheet>";
            var sheet = switch (try decode.scanSheet(allocator, sheet_xml, &.{}, .{})) {
                .ok => |s| s,
                .refused => return error.UnexpectedRefusal,
            };
            defer sheet.deinit();

            var shared = switch (try classifySheet(allocator, sheet.cells)) {
                .ok => |s| s,
                .refused => return error.UnexpectedRefusal,
            };
            defer shared.deinit();
            if (shared.groups.len != 1) return error.WrongGroupCount;

            // 2. Translation, including the arena the copied AST lives
            //    in and the printed text at the end of it.
            for (shared.entries) |e| {
                const slave = switch (e.role) {
                    .slave => |s| s,
                    else => continue,
                };
                var t = switch (try translateText(
                    allocator,
                    shared.groups[slave.group].text,
                    slave.delta,
                )) {
                    .ok => |v| v,
                    .refused => return error.UnexpectedRefusal,
                };
                defer t.deinit();
                if (t.text.len == 0) return error.EmptyTranslation;
            }

            // 3. Calc state, whose two owned lists are the extensions
            //    and the feature names.
            const wb_xml = "<workbook" ++ ns_attr ++ "><workbookPr date1904=\"1\"/>" ++
                "<calcPr calcId=\"1\" iterate=\"true\"/><extLst>" ++
                "<ext uri=\"" ++ calc_features_ext_uri ++ "\"><calcFeatures>" ++
                "<feature name=\"microsoft.com:RD\"/><feature name=\"microsoft.com:FV\"/>" ++
                "</calcFeatures></ext><ext uri=\"{OTHER}\"/></extLst></workbook>";
            var state = switch (try parseCalcState(allocator, wb_xml)) {
                .ok => |s| s,
                .refused => return error.UnexpectedRefusal,
            };
            defer state.deinit(allocator);
            if (state.extensions.len != 2 or state.calc_features.len != 2) {
                return error.WrongExtensionCount;
            }

            // 4. A refused parse must free everything it allocated —
            //    the refusal is a normal return, so `errdefer` never
            //    fires and only an explicit release can be right.
            const bad_xml = "<worksheet" ++ ns_attr ++ "><sheetData>" ++
                "<row r=\"1\"><c r=\"A1\"><f t=\"shared\" si=\"9\"/></c></row>" ++
                "</sheetData></worksheet>";
            var bad_sheet = switch (try decode.scanSheet(allocator, bad_xml, &.{}, .{})) {
                .ok => |s| s,
                .refused => return error.UnexpectedRefusal,
            };
            defer bad_sheet.deinit();
            switch (try classifySheet(allocator, bad_sheet.cells)) {
                .ok => return error.ExpectedRefusal,
                .refused => |r| if (r.reason != .shared_unknown_si) return error.WrongRefusal,
            }
        }
    };
    try testing.checkAllAllocationFailures(testing.allocator, H.run, .{});
}

// ─── fuzz (§8.1: attributes and calc state) ──────────────────────

fn fuzzCalcTarget(_: void, smith: *std.testing.Smith) anyerror!void {
    @disableInstrumentation();
    var buf: [512]u8 = undefined;
    const input = buf[0..smith.slice(&buf)];
    const a = std.testing.allocator;

    // 1. No `<f>` attribute string may panic — and none may classify
    //    two ways. The attribute region is spliced in raw, so the fuzzer
    //    reaches the inventory with whatever bytes it likes.
    {
        var doc: std.ArrayListUnmanaged(u8) = .empty;
        defer doc.deinit(a);
        try doc.appendSlice(a, "<worksheet");
        try doc.appendSlice(a, ns_attr);
        try doc.appendSlice(a, "><sheetData><row r=\"1\"><c r=\"A1\"><f ");
        try doc.appendSlice(a, input);
        try doc.appendSlice(a, ">A1+1</f></c></row></sheetData></worksheet>");

        var scanned = try decode.scanSheet(a, doc.items, &.{}, .{});
        switch (scanned) {
            .refused => |r| _ = r.planeTwo(),
            .ok => |*ok| {
                defer ok.deinit();
                for (ok.cells) |c| {
                    const f = c.formula orelse continue;
                    const first = classifyFormula(f, null);
                    const second = classifyFormula(f, null);
                    switch (first) {
                        .ok => |x| switch (second) {
                            .ok => |y| {
                                // Determinism, field by field: the same
                                // bytes classify the same way, twice.
                                try std.testing.expectEqual(x.kind, y.kind);
                                try std.testing.expectEqual(x.si, y.si);
                                try std.testing.expectEqual(x.ca, y.ca);
                                try std.testing.expectEqual(x.aca, y.aca);
                                try std.testing.expectEqual(x.bx, y.bx);
                                try std.testing.expectEqual(x.xml_space, y.xml_space);
                                try std.testing.expectEqual(x.ref == null, y.ref == null);
                                // An accepted formula is never a data
                                // table and never a `bx` assignment.
                                try std.testing.expect(x.kind != .data_table);
                                try std.testing.expect(!x.bx);
                                // …and a shared one always names a group.
                                if (x.kind == .shared) try std.testing.expect(x.si != null);
                                if (x.kind == .array) try std.testing.expect(x.ref != null);
                            },
                            .refused => return error.ClassifiedTwoWays,
                        },
                        .refused => |x| switch (second) {
                            .ok => return error.ClassifiedTwoWays,
                            .refused => |y| {
                                try std.testing.expectEqual(x.reason, y.reason);
                                _ = x.planeTwo();
                            },
                        },
                    }
                }

                // The sheet-wide pass over the same cells must not leak
                // or panic either, whatever the topology turned out to be.
                var shared = try classifySheet(a, ok.cells);
                switch (shared) {
                    .ok => |*s| s.deinit(),
                    .refused => |r| _ = r.planeTwo(),
                }
                shared = undefined;
            },
        }
    }

    // 2. No `<calcPr>` — or anything else in a workbook part — may
    //    panic or leak the two lists the parse owns.
    {
        var doc: std.ArrayListUnmanaged(u8) = .empty;
        defer doc.deinit(a);
        try doc.appendSlice(a, "<workbook");
        try doc.appendSlice(a, ns_attr);
        try doc.append(a, '>');
        try doc.appendSlice(a, input);
        try doc.appendSlice(a, "</workbook>");

        var state = try parseCalcState(a, doc.items);
        switch (state) {
            .ok => |*s| {
                // Whatever was parsed must re-emit; a state that cannot
                // be written back is a state that cannot round-trip.
                var out: std.Io.Writer.Allocating = .init(a);
                defer out.deinit();
                try s.writeCalcPr(&out.writer);
                if (s.present) {
                    try std.testing.expect(std.mem.startsWith(u8, out.written(), "<calcPr"));
                } else {
                    try std.testing.expectEqual(@as(usize, 0), out.written().len);
                }
                // An accepted state never carries precision-as-displayed.
                try std.testing.expect(s.full_precision);
                s.deinit(a);
            },
            .refused => |r| _ = r.planeTwo(),
        }
        state = undefined;

        _ = parseSheetCalcPr(doc.items);
    }

    // 3. Translation over arbitrary text: no delta may panic, leak, or
    //    produce a translation that differs between two runs.
    if (input.len <= 128) {
        for ([_]Delta{ .{}, .{ .rows = 1, .cols = 1 }, .{ .rows = -1_000_000, .cols = -20_000 } }) |d| {
            var first = try translateText(a, input, d);
            switch (first) {
                .refused => {},
                .ok => |*t| {
                    var second = try translateText(a, input, d);
                    switch (second) {
                        .refused => return error.TranslatedTwoWays,
                        .ok => |*u| {
                            defer u.deinit();
                            try std.testing.expectEqualStrings(t.text, u.text);
                        },
                    }
                    t.deinit();
                },
            }
            first = undefined;
        }
    }
}

test "fuzz: no `<f>` attribute or calc state can panic, leak, or classify two ways" {
    try std.testing.fuzz({}, fuzzCalcTarget, .{
        .corpus = &[_][]const u8{
            "t=\"shared\" ref=\"A1:A7\" si=\"1\"",
            "si=\"1\" t=\"shared\"",
            "t=\"array\" ref=\"I2:N5\" ca=\"1\" si=\"0\"",
            "bx=\"true\"",
            "dt2D=\"1\" dtr=\"1\" r1=\"A1\" r2=\"B1\"",
            "xml:space=\"preserve\"",
            "<calcPr calcId=\"40001\" />",
            "<calcPr iterateCount=\"100\" refMode=\"A1\" iterate=\"false\" iterateDelta=\"0.001\"/>",
            "<workbookPr date1904=\"1\"/><calcPr fullPrecision=\"0\"/>",
            "SUM($A$1:B2,Sheet1!C3#)",
        },
    });
}

// ─── masters, parsed once per group ──────────────────────────────

test "masters: one parse per group, translated per slave" {
    var c = try classifySheetXml(
        \\<row r="1"><c r="A1"><f t="shared" ref="A1:A3" si="0">$B1+C$1</f></c></row>
        \\<row r="2"><c r="A2"><f t="shared" si="0"/></c></row>
        \\<row r="3"><c r="A3"><f t="shared" si="0"/></c></row>
    );
    defer c.deinit();
    const ok = c.result.ok;

    var masters = switch (try parseMasters(testing.allocator, ok.groups)) {
        .ok => |m| m,
        .refused => return error.TestUnexpectedRefusal,
    };
    defer masters.deinit();
    try testing.expectEqual(@as(usize, 1), masters.asts.len);

    const want = [_][]const u8{ "$B2+C$1", "$B3+C$1" };
    var seen: usize = 0;
    for (ok.entries) |e| {
        const s = switch (e.role) {
            .slave => |v| v,
            else => continue,
        };
        var t = try masters.translateFor(testing.allocator, s.group, s.delta);
        defer t.deinit();
        try testing.expectEqualStrings(want[seen], t.text);
        seen += 1;
    }
    try testing.expectEqual(@as(usize, 2), seen);
}

test "masters: a master that will not parse refuses its group, at its cell" {
    var c = try classifySheetXml(
        \\<row r="4"><c r="C4"><f t="shared" ref="C4:C5" si="0">SUM(</f></c></row>
        \\<row r="5"><c r="C5"><f t="shared" si="0"/></c></row>
    );
    defer c.deinit();
    switch (try parseMasters(testing.allocator, c.result.ok.groups)) {
        .ok => |m| {
            var mm = m;
            mm.deinit();
            return error.TestExpectedRefusal;
        },
        .refused => |r| {
            try testing.expectEqual(Refusal.Reason.shared_master_unparsable, r.reason);
            try testing.expectEqual(@as(u32, 4), r.cell.?.row);
            try testing.expectEqual(@as(u32, 2), r.cell.?.col);
        },
    }
}
