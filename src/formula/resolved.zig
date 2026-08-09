//! §5.7.3 step 3 — the projection a recalc stages into, and the patcher
//! that writes its cached values back into a worksheet part (M5b1).
//!
//! What a projection is for
//! ------------------------
//! Evaluation produces values; a file transaction needs *bytes*. The
//! step between them is a `ResolvedSheet`: every `<c>` the part already
//! has, paired with its byte ranges and its raw `<f>`, plus the results
//! this run staged onto it. Nothing here decides what a cell computes
//! and nothing here writes a file. It is the one place that knows both
//! where a cached value lives in the document and what the run wants it
//! to say.
//!
//! Byte confinement is the whole claim
//! -----------------------------------
//! A recalculated workbook is mostly the workbook it was. Styles,
//! shared strings, drawings, conditional formats, the formulas
//! themselves, and every byte of every cell this run did not touch are
//! supposed to come back **identical** — and "supposed to" is not a
//! testable statement. So the patcher does not serialize a model back to
//! XML. It produces a list of `Edit`s over the source, each naming the
//! range it replaces and the kind of replacement it is, and
//! `approvedRange` maps each kind to the *one* range it is allowed to
//! address. A test then has two independent handles on the same claim:
//! re-applying the edits must reproduce the output exactly (so the list
//! explains every changed byte), and every edit must equal its approved
//! range (so no edit reaches outside a `<v>` or a `t` attribute). Neither
//! is a comment.
//!
//! There is one parser, and it hands out its coordinates
//! -----------------------------------------------------
//! The spans come from `decode.scanSheet` — the same walk that
//! classifies the cells — rather than from a locator written here.
//! A second scanner that disagreed with the classifier by one byte would
//! put a value in the wrong element, and the disagreement would be
//! invisible until a workbook came back wrong.
//!
//! `Sheet.slots` rather than `Sheet.cells` decides where a `<c>` exists,
//! because the merged view drops an empty styled cell and the document
//! does not: `<c r="A1" s="3"/>` carries no value and still occupies the
//! bytes a new `<c r="A1">` would need.
//!
//! The approved mutation set (§5.8b, M7b1)
//! ---------------------------------------
//! M5b1's kinds — `<v>`+`t` — plus four §5.8b additions: the anchor's
//! `f@ref` value, owned tail `<c>` create and clear, and `<dimension
//! ref>` expansion when created tails extend the used range. The DA
//! path that would USE the additions is gated per-transition by
//! `spill_transitions`: a row passes only against a committed
//! byte-diffed Excel-authored reference, refuses
//! `FormulaSpillPersistUnsupported` otherwise, and every row refuses
//! today. A covered legacy CSE passes with `<v>`+`t` alone. A `.plain`
//! publication with no `<c>` to land on still refuses — the projection
//! *carries* it (`ResolvedSheet.appends`) because the serializer path
//! (M5c) will want it, and refusing beats dropping it quietly. Only an
//! OWNED tail (`Role.spill_tail`) may become a created `<c>`, and only
//! under its anchor's proven row.
//!
//! M7c adds §5.8c's one authoring mutation: a whole `<f>…</f>` inserted
//! into an existing `<c>` that has none (`f_insert`; the self-closing
//! shape rides `reopen_self_closing`'s replacement). `.scalar`
//! authoring is live end-to-end on it; `.dynamic_array` and `.cse`
//! authoring consult the same `spill_transitions` discipline and refuse
//! `transition_unproven` until their byte-diffed Excel references land.
//!
//! The transition table is normative and shared
//! --------------------------------------------
//! `transitionFor` is §5.7.3's table as one function, and the blank→0
//! conversion is not in it: `value.publish` already owns that conversion
//! for every layer, so a publication arrives here as a
//! `PublishedScalar`, which has no blank arm to get wrong.
//!
//! Provenance
//! ----------
//! Spec-pinned: `CT_Cell`'s child order (`f`, then `v`/`is`) to
//! ECMA-376; `ST_CellType`'s default `n` to the same, which is why a
//! number publication *removes* `t` rather than writing `t="n"`;
//! ST_Xstring's escape grammar to the codec M4b1 landed. Row-decided:
//! which of `<v>`'s two shapes a transition produces, and that an
//! already-correct cell is left alone.

const std = @import("std");
const assert = std.debug.assert;
const Allocator = std.mem.Allocator;

const coords = @import("zlsx_refs");
const calc = @import("calc.zig");
const decode = @import("decode.zig");
const parser = @import("parser.zig");
const spill = @import("spill.zig");
const value = @import("value.zig");

/// §10's plane-2 taxonomy has exactly one home; this file pays the same
/// import `decode.zig` and `calc.zig` do, for the same reason.
pub const PlaneTwo = decode.PlaneTwo;
pub const CellSite = decode.CellSite;
pub const Span = decode.Span;

// ─── refusals (§10) ──────────────────────────────────────────────

pub const Refusal = struct {
    reason: Reason,
    /// Set when the refusal is about one cell, which is all of them
    /// except the projection-wide ones.
    cell: ?CellSite = null,
    /// The §5.8b transition row the refusal is about, when it is about
    /// one — a diagnostic that names WHICH byte-diffed reference is
    /// missing is one someone can act on.
    transition: ?SpillTransition.Id = null,

    pub const Reason = enum {
        // ── the spill gate (§5.7.3 / §5.8b) ──
        /// The staged result is not 1×1 and the target is not an array
        /// formula: an ordinary cell publishing an array has no declared
        /// or placed region to persist into.
        non_scalar_result,
        /// The target's `<f t="array" ref=…>` is a dynamic-array anchor
        /// (M4a resolved its `cm`/`vm` to the DA dialect) staged WITHOUT
        /// the model's placement outcome. The patcher never reconstructs
        /// a placement the model already made (§5.8b), so a role-less DA
        /// anchor refuses generically.
        dynamic_array_anchor,
        /// A DA transition row exists for this shape and no committed
        /// byte-diffed Excel reference pins it (§5.8b). `transition`
        /// names the row; the missing-reference list is
        /// `missingReferences()`.
        transition_unproven,
        /// The target carries `c@vm`. Value metadata is something this
        /// engine has never been shown (M4a decision 6), and writing a
        /// cached value under it would publish beneath semantics we
        /// cannot read. Permanent, not unlockable by reference.
        value_metadata_write,
        /// The target's `<f t="array" ref=…>` is a legacy CSE anchor and
        /// the projection does not cover its declared range on existing
        /// `<c>`s — persisting the anchor alone would leave the slaves'
        /// cached values contradicting it. Also the anchor-not-at-TL
        /// shape, which is the same statement about a different corner.
        cse_range_mismatch,
        /// A CSE anchor's `ref` does not parse as a range. The model
        /// refuses this at build; a standalone projection meets it here.
        cse_ref_unparseable,
        /// The publication lands where the part has no `<c>` at all and
        /// is not an owned spill tail. Inserting one is outside the
        /// approved mutation set (§5.8b names only the OWNED tail).
        cell_insertion_unsupported,

        // ── authoring (§5.8c, M7c) ──
        /// A `FormulaWrite` targets a cell that already carries an
        /// `<f>`. Replacing a formula's body would address bytes inside
        /// `spans.f`, and `f@ref` is THE one exception to that
        /// invariant (M7b1 decision 5) — authoring writes formulas
        /// where none exist, and a rewrite waits for its own approved
        /// mutation.
        formula_overwrite_unsupported,
        /// A `FormulaWrite` targets a cell carrying `c@cm`. The record
        /// describes the formula the cell used to hold; a fresh `<f>`
        /// under it would leave metadata narrating a formula that is
        /// no longer there.
        authored_under_cell_metadata,
        /// The authored text cannot become a legal `<f>` body: empty,
        /// or carrying a character the FORMULA carrier cannot encode
        /// (`decode.encodeAuthoredFormula` — there is no second escape
        /// layer to put it in).
        authored_text_unencodable,
        /// The staged role and the authoring dialect contradict each
        /// other — a `.scalar`/`.cse` write carrying a placement role,
        /// or an authored publication staged as someone else's tail.
        /// Like `duplicate_publication`: two stories about one cell,
        /// neither chosen silently.
        authored_role_contradiction,

        // ── tail geometry the approved set cannot address ──
        /// An owned tail lands in a row the part has no `<row>` element
        /// for (or has two of). Creating a `<row>` is a mutation §5.8b
        /// does not name.
        tail_row_missing,
        /// The enclosing `<row>` is self-closing; giving it content is
        /// a reopen §5.8b does not name.
        tail_row_self_closing,
        /// The enclosing `<row>` declares `spans` that would not cover
        /// the created `<c>`, and spans maintenance is not in the
        /// approved set — a stale hint we will not write and will not
        /// leave behind.
        tail_row_spans_stale,
        /// A spill extends the used range and the `<dimension ref>` is
        /// not a canonical parseable range to expand — §5.8b: such a
        /// spill refuses until the maintenance can be performed.
        dimension_unparseable,
        /// The DA anchor has no `ref` attribute to rewrite, or one that
        /// does not parse — the file's own record of the prior extent
        /// is the one thing tail clears derive from.
        anchor_ref_unusable,
        /// A tail publication whose anchor is not in the projection as
        /// a DA anchor. Half a construct; the other half is what makes
        /// a tail writable at all.
        tail_without_anchor,
        /// A coordinate the anchor's stored `ref` claims as a tail
        /// holds a formula, or is simultaneously a staged target —
        /// clearing either would remove content the ownership record
        /// does not actually own.
        tail_clear_foreign,

        // ── shapes a patch cannot be confined over ──
        /// Two `<c>` claim one coordinate. `scanSheet` refuses that
        /// among *modeled* cells; a duplicate where one of them is an
        /// empty styled cell reaches here, and patching either one
        /// leaves the other shadowing it.
        duplicate_cell_slot,
        /// Two publications claim one coordinate. Last-wins and
        /// first-wins are both defensible, so neither is chosen
        /// silently.
        duplicate_publication,
        /// A `<c>` carrying both `<v>` and `<is>`. `CT_Cell` allows one
        /// or the other; which a reader believes is a guess, and a
        /// patcher that guessed would publish into the loser.
        ambiguous_cell_content,

        // ── values that cannot be written back ──
        /// A rich error spelling that would not read back as the error
        /// it came from. "Preserved, never produced" (§5.3a) is a
        /// promise about spellings that came out of a file.
        unwritable_error_spelling,
        /// A non-finite number reached publication. N4a converts those
        /// to `#NUM!` where they are produced, so one arriving here is a
        /// value no cached `<v>` can hold.
        non_finite_number,
    };

    /// Exhaustive by construction — a new `Reason` fails to compile
    /// until it has a §10 plane.
    pub fn planeTwo(self: Refusal) PlaneTwo {
        return switch (self.reason) {
            .non_scalar_result,
            .dynamic_array_anchor,
            .transition_unproven,
            .value_metadata_write,
            .cse_range_mismatch,
            .cell_insertion_unsupported,
            .formula_overwrite_unsupported,
            .authored_under_cell_metadata,
            .tail_row_missing,
            .tail_row_self_closing,
            .tail_row_spans_stale,
            .dimension_unparseable,
            .anchor_ref_unusable,
            .tail_without_anchor,
            .tail_clear_foreign,
            => .FormulaSpillPersistUnsupported,

            .cse_ref_unparseable,
            .duplicate_cell_slot,
            .duplicate_publication,
            .ambiguous_cell_content,
            .unwritable_error_spelling,
            .authored_text_unencodable,
            .authored_role_contradiction,
            => .FormulaMalformedInput,

            .non_finite_number => .FormulaResultNotRepresentable,
        };
    }
};

fn refuseAt(reason: Refusal.Reason, row: coords.Row, col: coords.Col) Refusal {
    return .{
        .reason = reason,
        .cell = .{ .row = row.oneBased(), .col = col.zeroBased() },
    };
}

// ─── what a run stages (§5.7.3 step 3) ───────────────────────────

/// Where a staged value came from. The patcher treats both the same —
/// a cached `<v>` is a cached `<v>` — and the diagnostic does not, which
/// is the only reason it is carried.
pub const Origin = enum {
    /// Evaluated by this run.
    computed,
    /// A `setCell` the caller staged before the run; §5.7.3's "setCell
    /// replacements".
    set_cell,
};

/// §5.8c's authoring dialect (M7c, **Zig-only** — the versioned C
/// export and Python land at M9a2).
///
/// `text` is the formula body exactly as the parser reads it: no
/// leading `=`, FORMULA carrier (XML escaping only, no ST_Xstring
/// stage — `decode.encodeAuthoredFormula` is the one encoder). The
/// caller evaluated it before staging — the publication it rides
/// carries the result — so a text the parser would refuse never gets
/// this far on the shipped paths; what THIS layer still verifies is
/// what only bytes can get wrong (encodability, the target's shape,
/// the dialect's transition row).
pub const FormulaWrite = struct {
    text: []const u8,
    dialect: Dialect = .scalar,

    pub const Dialect = union(enum) {
        /// An ordinary cell formula. Live end-to-end at M7c: the
        /// authored `<f>` plus the cached value ride the proven
        /// mutation set, and the round-trip fixture is the proof.
        scalar,
        /// A dynamic-array anchor. Refuses `transition_unproven` at
        /// M7c: the authored `cm` and its XLDAPR record are part-graph
        /// mutations whose spec ARRIVES with the byte-diffed Excel
        /// references (§5.8b), so the rows park in
        /// `missingReferences()` until then.
        dynamic_array,
        /// A legacy CSE anchor over this declared range (A1 notation).
        /// No metadata byte participates — the anchor's fresh
        /// `<f t="array" ref>` is the ONE unproven mutation, and its
        /// row refuses until an Excel-authored reference pins what
        /// Office writes around one.
        cse: []const u8,
    };
};

/// One staged result, at the coordinate it belongs to.
///
/// `result` is a `PublishedScalar` and not a `ScalarValue`, so the
/// blank→0 conversion has already happened in the one place that owns it
/// (`value.publish`). A blank cannot reach the patcher, and the patcher
/// therefore has no opinion about what one would mean.
pub const Publication = struct {
    row: coords.Row,
    col: coords.Col,
    result: value.PublishedScalar,
    origin: Origin = .computed,
    /// The shape the evaluator produced. Anything but 1×1 meets the
    /// pre-M7 gate.
    shape: value.Shape = .{ .rows = 1, .cols = 1 },
    /// The dialect the target's formula was evaluated in (M4a resolves
    /// it from `c@cm`/`c@vm`). It does not change what a scalar cached
    /// value looks like; it names *which* of the two array refusals a
    /// `t="array"` target takes.
    dialect: value.Dialect = .legacy,
    /// What the model decided about this publication's placement —
    /// carried, never re-decided here (§5.8b: `stage()` persists what
    /// the model placed). `.plain` is every scalar result and setCell;
    /// the staging layer fills the other arms from `spill.Registry`.
    role: Role = .plain,
    /// §5.8c (M7c): the formula this publication also WRITES. Null for
    /// every publication that only caches a value — which is every one
    /// M7b1 knew about, so the field's default is the old behavior.
    authored: ?FormulaWrite = null,
};

/// A publication's placement role (§5.8b).
pub const Role = union(enum) {
    plain,
    /// A dynamic-array anchor, with the outcome the model placed. A DA
    /// anchor staged WITHOUT its outcome (`.plain`) refuses generically
    /// — the patcher will not reconstruct a placement the model already
    /// made.
    da_anchor: spill.Outcome,
    /// A tail cell owned by the anchor at (`row`, `col`) — persisted
    /// only through the anchor's own transition.
    spill_tail: struct { row: coords.Row, col: coords.Col },
};

/// The staged set, foldable into a projection exactly once.
///
/// §5.7.3 says deltas are "consumed once", and a flag on the staged set
/// is the only place that can be enforced: a second recalc over the same
/// workbook that re-applied a delta the first one already wrote would
/// publish a value nothing computed this run. `consume` is called on the
/// success path only, so a refused projection leaves the set intact for
/// a caller that fixes the cause and retries.
pub const StagedDeltas = struct {
    publications: []const Publication,
    consumed: bool = false,

    pub fn consume(self: *StagedDeltas) error{DeltasAlreadyConsumed}![]const Publication {
        if (self.consumed) return error.DeltasAlreadyConsumed;
        self.consumed = true;
        return self.publications;
    }
};

// ─── the projection ──────────────────────────────────────────────

/// A staged publication joined to the `<c>` it lands on.
pub const Target = struct {
    publication: Publication,
    /// Where the `<c>` and its interpreted children sit in the part.
    spans: decode.CellSpans,
    /// The `<f>` as M4b1 decoded it, raw attribute region included, or
    /// null when the cell has none. Carried unchanged — and the patcher
    /// proves it by never addressing a byte inside `spans.f`.
    formula: ?decode.Formula,
    /// What the cell contributed to the merged view before the patch, so
    /// a round-trip has both of its ends in one place.
    was: decode.InputCell,
    /// `c@cm` / `c@vm` as scanned (0 = absent). Carried for the gate: a
    /// `vm` names value metadata this engine has never been shown (M4a
    /// decision 6), and writing under one refuses HERE too, because a
    /// setCell publication reaches the patcher without ever meeting the
    /// resolver.
    cm: u32 = 0,
    vm: u32 = 0,
};

/// **Lifetime**: a projection borrows both of its inputs. `source` is
/// the caller's bytes and `Target.formula` points into the scan's arena,
/// so the part and the `decode.Sheet` it was scanned into must both
/// outlive this. Copying either would mean holding a second opinion
/// about a document that has not changed.
pub const ResolvedSheet = struct {
    arena: std.heap.ArenaAllocator,
    /// The source part. Borrowed.
    source: []const u8,
    /// Publications joined to an existing `<c>`, ascending by
    /// coordinate.
    targets: []const Target,
    /// Publications with no `<c>` to land on, ascending by coordinate.
    /// A `.plain` append is carried, never patched — see the module
    /// header. A `.spill_tail` append is the one shape M7b1 may turn
    /// into a created `<c>`, and only under its anchor's proven
    /// transition.
    appends: []const Publication,
    /// The scan's slots, borrowed like `source` — an owned tail the
    /// file stored and this run's shrink clears is not a target, so its
    /// bytes are findable only here (M7b1).
    slots: []const decode.CellSlot,
    /// The scan's row geometry, borrowed like `source` — a created
    /// `<c>` is confined by it (M7b1).
    rows: []const decode.RowSlot,
    /// The scan's `<dimension>` coordinates, or null when the part has
    /// none.
    dimension: ?decode.DimensionSpans,

    pub fn deinit(self: *ResolvedSheet) void {
        self.arena.deinit();
        self.* = undefined;
    }

    /// The target at a coordinate, or null. Binary over the projection's
    /// coordinate order — the CSE coverage gate asks this once per
    /// declared cell, and a declared range is bounded only by
    /// `max_matrix_cells`.
    pub fn targetAt(self: *const ResolvedSheet, row: coords.Row, col: coords.Col) ?Target {
        var lo: usize = 0;
        var hi: usize = self.targets.len;
        while (lo < hi) {
            const mid = lo + (hi - lo) / 2;
            const t = self.targets[mid];
            if (t.publication.row.oneBased() < row.oneBased() or
                (t.publication.row == row and t.publication.col.zeroBased() < col.zeroBased()))
            {
                lo = mid + 1;
            } else if (t.publication.row == row and t.publication.col == col) {
                return t;
            } else {
                hi = mid;
            }
        }
        return null;
    }
};

pub const ProjectResult = union(enum) {
    ok: ResolvedSheet,
    refused: Refusal,
};

/// Fold a run's staged results onto the part they belong to.
///
/// `scan` must be the accepted scan of `source` — the same bytes, the
/// same options. Everything the input-cell contract refuses has already
/// refused there, which is what makes every refusal below about the
/// *write* rather than about the document.
pub fn project(
    gpa: Allocator,
    source: []const u8,
    scan: *const decode.Sheet,
    staged: *StagedDeltas,
) error{ OutOfMemory, DeltasAlreadyConsumed }!ProjectResult {
    // Read without consuming: a refusal below must leave the set intact.
    const pubs = staged.publications;

    var arena = std.heap.ArenaAllocator.init(gpa);
    var keep = false;
    defer if (!keep) arena.deinit();
    const a = arena.allocator();

    var i: usize = 1;
    while (i < scan.slots.len) : (i += 1) {
        const cur = scan.slots[i];
        const prev = scan.slots[i - 1];
        if (cur.row == prev.row and cur.col == prev.col) {
            return .{ .refused = refuseAt(.duplicate_cell_slot, cur.row, cur.col) };
        }
    }

    // Sort a permutation, not a copy: the publications are 120 bytes
    // each and an index is four, and the copy lived in the projection's
    // arena until the patch died (§9.1). The keys the walk below sees
    // are unique — the duplicate check right here refuses otherwise —
    // so the unstable sort yields the same order the copying sort did.
    const order = try gpa.alloc(u32, pubs.len);
    defer gpa.free(order);
    for (order, 0..) |*slot, idx| slot.* = @intCast(idx);
    std.mem.sortUnstable(u32, order, pubs, lessThanPublicationAt);
    i = 1;
    while (i < order.len) : (i += 1) {
        const cur = pubs[order[i]];
        const prev = pubs[order[i - 1]];
        if (cur.row == prev.row and cur.col == prev.col) {
            return .{ .refused = refuseAt(.duplicate_publication, cur.row, cur.col) };
        }
    }

    // Both splits are knowable before either list exists, so each
    // backing is one exact-size allocation: a list that grows inside an
    // arena strands every buffer it abandons until the arena dies.
    var targets_n: usize = 0;
    for (pubs) |p| {
        if (slotAt(scan.slots, p.row, p.col) != null) targets_n += 1;
    }
    var targets: std.ArrayListUnmanaged(Target) = .empty;
    try targets.ensureTotalCapacityPrecise(a, targets_n);
    var appends: std.ArrayListUnmanaged(Publication) = .empty;
    try appends.ensureTotalCapacityPrecise(a, pubs.len - targets_n);
    for (order) |idx| {
        const p = pubs[idx];
        const slot = slotAt(scan.slots, p.row, p.col) orelse {
            try appends.append(a, p);
            continue;
        };
        if (slot.spans.v != null and slot.spans.is != null) {
            return .{ .refused = refuseAt(.ambiguous_cell_content, p.row, p.col) };
        }
        const modeled = cellAt(scan.cells, p.row, p.col);
        try targets.append(a, .{
            .publication = p,
            .spans = slot.spans,
            .formula = if (modeled) |m| m.formula else null,
            // A slot with no modeled cell is the empty styled cell the
            // merged view drops, and blank is exactly what it
            // contributed.
            .was = if (modeled) |m| m.input else .blank,
            .cm = if (modeled) |m| m.cm else 0,
            .vm = if (modeled) |m| m.vm else 0,
        });
    }

    // Materialized BEFORE the result literal: `.arena = arena` copies
    // the arena's state first, so an allocation in a later field
    // initializer that opened a fresh chunk would be known only to the
    // local copy — and leak when the returned one deinits. With the
    // exact capacities above these are non-allocating today; the order
    // keeps that from being load-bearing.
    const targets_out = try targets.toOwnedSlice(a);
    const appends_out = try appends.toOwnedSlice(a);

    _ = try staged.consume();
    keep = true;
    return .{ .ok = .{
        .arena = arena,
        .source = source,
        .targets = targets_out,
        .appends = appends_out,
        .slots = scan.slots,
        .rows = scan.rows,
        .dimension = scan.dimension,
    } };
}

fn lessThanPublication(_: void, x: Publication, y: Publication) bool {
    if (x.row.oneBased() != y.row.oneBased()) return x.row.oneBased() < y.row.oneBased();
    return x.col.zeroBased() < y.col.zeroBased();
}

fn lessThanPublicationAt(pubs: []const Publication, x: u32, y: u32) bool {
    return lessThanPublication({}, pubs[x], pubs[y]);
}

fn slotAt(slots: []const decode.CellSlot, row: coords.Row, col: coords.Col) ?decode.CellSlot {
    var lo: usize = 0;
    var hi: usize = slots.len;
    while (lo < hi) {
        const mid = lo + (hi - lo) / 2;
        const s = slots[mid];
        if (s.row.oneBased() < row.oneBased() or
            (s.row == row and s.col.zeroBased() < col.zeroBased()))
        {
            lo = mid + 1;
        } else if (s.row == row and s.col == col) {
            return s;
        } else {
            hi = mid;
        }
    }
    return null;
}

fn cellAt(cells: []const decode.SheetCell, row: coords.Row, col: coords.Col) ?decode.SheetCell {
    var lo: usize = 0;
    var hi: usize = cells.len;
    while (lo < hi) {
        const mid = lo + (hi - lo) / 2;
        const c = cells[mid];
        if (c.row.oneBased() < row.oneBased() or
            (c.row == row and c.col.zeroBased() < col.zeroBased()))
        {
            lo = mid + 1;
        } else if (c.row == row and c.col == col) {
            return c;
        } else {
            hi = mid;
        }
    }
    return null;
}

// ─── §5.7.3's transition table ───────────────────────────────────

/// The `<c>` shape one publication requires.
pub const Transition = struct {
    /// The `t` attribute value the patched cell must carry, or null when
    /// the attribute is *removed*. Null is not "leave it alone": a
    /// number's type is `ST_CellType`'s default, and writing `t="n"`
    /// where Excel writes nothing would change bytes to say what their
    /// absence already said.
    type_attr: ?[]const u8,
    /// The bytes that go between `<v>` and `</v>`, already encoded for
    /// the carrier the cell will be read through.
    v: []const u8,
};

pub const TransitionError = error{
    OutOfMemory,
    UnwritableErrorSpelling,
    NonFiniteNumber,
};

/// §5.7.3's table, as one function.
///
/// `""` needs no row of its own: it encodes to nothing and the emitter
/// writes `<v></v>` around it, which is exactly what the table asks for.
/// A blank needs no row either — `value.publish` converted it to `0`
/// before it could get here.
pub fn transitionFor(a: Allocator, published: value.PublishedScalar) TransitionError!Transition {
    switch (published) {
        .number => |n| {
            if (!std.math.isFinite(n)) return error.NonFiniteNumber;
            var buf: [value.format_buf_len]u8 = undefined;
            const text = value.formatNumber(&buf, n);
            return .{ .type_attr = null, .v = try a.dupe(u8, text) };
        },
        // A STRING carrier: ST_Xstring, then XML escaping. The carrier
        // class is the *site's*, and this site is a `t="str"` `<v>` —
        // which is why the same text inside an `<f>` would take only the
        // second pass (M4b1's split, and the reason a literal
        // `_x0041_` means two different things in two elements).
        .text => |t| return .{ .type_attr = "str", .v = try decode.encodeAuthoredString(a, t) },
        .boolean => |b| return .{ .type_attr = "b", .v = if (b) "1" else "0" },
        .err => |e| {
            const spelling = e.spelling();
            // Rich spellings are byte-preserved — but "preserved" is a
            // promise about a spelling that came out of a file, and one
            // this engine could not read back is one it must not write.
            if (decode.errorFromSpelling(spelling) == null) {
                return error.UnwritableErrorSpelling;
            }
            return .{ .type_attr = "e", .v = spelling };
        },
    }
}

// ─── §5.8b's cm/vm transition table ──────────────────────────────

/// One enumerated persistence transition (§5.8b): the exact metadata
/// collection, record type, one-based index rule and missing-record
/// behavior, plus the committed byte-diffed Excel-authored reference
/// that pins it. A row whose `reference` is null REFUSES — no
/// transition ships on a guess — and a `permanent` row refuses by
/// contract (M4a), unlockable by nothing.
///
/// The table is data on purpose: the refusing-rows test enumerates it,
/// `missingReferences()` is the surfaced park list, and committing a
/// reference flips exactly one row without touching the gate.
pub const SpillTransition = struct {
    id: Id,
    collection: Collection,
    /// The record type name, exactly as Office spells it — the match is
    /// case-sensitive (M4a decision 15).
    record_type: []const u8,
    index: IndexRule,
    missing_record: MissingRecord,
    /// Repo-relative path of the committed byte-diffed reference pair
    /// that pins this row, or null while none is committed. Authoring
    /// one needs Excel — `scripts/oracle/regenerate.sh` is the unblock.
    reference: ?[]const u8,
    /// Refuses by contract, not by missing proof.
    permanent: bool = false,

    pub const Id = enum {
        /// Anchor was spilled in the part, spills again: `<v>`+`t`, an
        /// `f@ref` rewrite when the extent moved, owned tail
        /// create/clear, `<dimension ref>` expansion. `c@cm` stays
        /// byte-identical, still naming its XLDAPR record.
        da_spill_rewrite,
        /// Anchor was spilled, now blocked: the cached value becomes
        /// `#SPILL!` (`t="e"`), `f@ref` collapses to the anchor, every
        /// owned tail clears. The cached-value side of the split ONLY —
        /// rich error metadata (`vm`) is never invented.
        da_spill_to_blocked,
        /// Anchor was blocked (`#SPILL!` cached), now spills: the
        /// recovery — `f@ref` grows from the anchor to the extent,
        /// tails create, the dimension may expand.
        da_blocked_to_spill,
        /// Anchor was blocked, stays blocked: the `#SPILL!` cached
        /// value and anchor-only `ref` rewrite in place.
        da_blocked_rewrite,
        /// Any publication landing on a `vm`-carrying cell. Permanent
        /// (M4a decision 6): reached through `vm`, a record means
        /// something this reader has never been shown.
        value_metadata_present,

        // ── authoring (§5.8c, M7c) — appended through the table's own
        //    seam; the M7b1 rows above are untouched ──
        /// A fresh `.dynamic_array` anchor the model placed spilled.
        /// The `cm` attribute, the XLDAPR record and every part-graph
        /// byte around them (the metadata part, its content type, its
        /// rel) are exactly what the reference pins — the builder
        /// LANDS with the reference set, so this row refuses even on
        /// an injected table (M7c decisions).
        da_author_spill,
        /// A fresh `.dynamic_array` anchor the model placed blocked:
        /// the authored `#SPILL!` shape. Same part-graph gate, same
        /// arrival.
        da_author_blocked,
        /// A fresh `.cse(ref)` anchor. No metadata byte participates —
        /// the one unproven mutation is the anchor's fresh
        /// `<f t="array" ref>`, whose sheet-part bytes ECMA-376 names,
        /// so the builder is provable through an injected table today
        /// and the reference re-pins empirically what Office writes
        /// AROUND one (calcChain, `aca`) when it lands.
        cse_author,

        pub fn name(self: Id) []const u8 {
            return @tagName(self);
        }
    };

    pub const Collection = enum {
        cell_metadata,
        value_metadata,
        /// No metadata collection participates (M7c: `cse_author`) —
        /// the row gates a fresh `<f t="array">` element, not a
        /// metadata transition, and an arm that pretended otherwise
        /// would name bytes the mutation never touches.
        none,
    };
    pub const IndexRule = enum {
        /// The cell's existing `c@cm`, one-based, preserved
        /// byte-identically — the patcher never addresses it.
        existing_cm,
        /// The cell's existing `c@vm`, one-based.
        existing_vm,
        /// The run must mint or attach the record — which index, which
        /// record, and every byte it takes is exactly what the row's
        /// reference pins (M7c: the `da_author_*` rows).
        authored_cm,
        /// No index participates (M7c: `cse_author`).
        none,
    };
    pub const MissingRecord = enum {
        /// A dangling index refuses. Upstream the resolver already
        /// refuses it (M4b1 decision 16), so the patcher never meets
        /// one; the row records the behavior so the reference can
        /// re-pin it empirically.
        refuse,
    };
};

/// §5.8b's table. Every DA row awaits its Excel-authored reference;
/// the `value_metadata_present` row is permanent.
pub const spill_transitions = [_]SpillTransition{
    .{ .id = .da_spill_rewrite, .collection = .cell_metadata, .record_type = "XLDAPR", .index = .existing_cm, .missing_record = .refuse, .reference = null },
    .{ .id = .da_spill_to_blocked, .collection = .cell_metadata, .record_type = "XLDAPR", .index = .existing_cm, .missing_record = .refuse, .reference = null },
    .{ .id = .da_blocked_to_spill, .collection = .cell_metadata, .record_type = "XLDAPR", .index = .existing_cm, .missing_record = .refuse, .reference = null },
    .{ .id = .da_blocked_rewrite, .collection = .cell_metadata, .record_type = "XLDAPR", .index = .existing_cm, .missing_record = .refuse, .reference = null },
    .{ .id = .value_metadata_present, .collection = .value_metadata, .record_type = "*", .index = .existing_vm, .missing_record = .refuse, .reference = null, .permanent = true },
    // §5.8c authoring (M7c), appended — extension happens at the end so
    // the M7b1 rows keep their positions and their park-list order.
    .{ .id = .da_author_spill, .collection = .cell_metadata, .record_type = "XLDAPR", .index = .authored_cm, .missing_record = .refuse, .reference = null },
    .{ .id = .da_author_blocked, .collection = .cell_metadata, .record_type = "XLDAPR", .index = .authored_cm, .missing_record = .refuse, .reference = null },
    .{ .id = .cse_author, .collection = .none, .record_type = "", .index = .none, .missing_record = .refuse, .reference = null },
};

fn rowById(table: []const SpillTransition, id: SpillTransition.Id) ?*const SpillTransition {
    for (table) |*row| {
        if (row.id == id) return row;
    }
    return null;
}

/// The park list: every transition that would pass once a byte-diffed
/// Excel reference is committed, in table order. What the final
/// summary surfaces, and what `regenerate.sh` + the spill reference
/// set will empty.
pub fn missingReferences(table: []const SpillTransition, buf: []SpillTransition.Id) []SpillTransition.Id {
    var n: usize = 0;
    for (table) |row| {
        if (row.permanent or row.reference != null) continue;
        if (n < buf.len) {
            buf[n] = row.id;
            n += 1;
        }
    }
    return buf[0..n];
}

// ─── the patch ───────────────────────────────────────────────────

/// What one edit did. The kind is not a label: `approvedRange` maps it
/// to the single byte range an edit of that kind may address, so a kind
/// that did not match its range is a bug the fuzz target catches.
pub const EditKind = enum {
    /// An existing `t="…"` rewritten in place.
    type_attr_replace,
    /// A `t="…"` inserted into a start tag that had none.
    type_attr_insert,
    /// An existing `t="…"` removed, with the whitespace that separated
    /// it from its neighbour.
    type_attr_remove,
    /// The bytes between `<v>` and `</v>` rewritten.
    v_content_replace,
    /// A `<v/>` given a body, which means replacing the element rather
    /// than a content region it does not have.
    v_element_replace,
    /// A whole `<v>…</v>` inserted into a cell that had none.
    v_insert,
    /// An `<is>…</is>` replaced by the `<v>…</v>` that supersedes it.
    /// A cell publishing through `<v>` cannot keep an inline string that
    /// a reader would believe instead.
    is_to_v,
    /// `<c …/>` reopened as `<c …>…</c>` so content can follow it.
    reopen_self_closing,

    // ── §5.8b's approved additions (M7b1) ──
    /// The raw value of an anchor's `<f ref="…">` rewritten — the ONE
    /// byte range inside `spans.f` any edit may address.
    f_ref_replace,
    /// A whole owned tail `<c r="…">…</c>` inserted into an existing
    /// `<row>`, at the point the row's slots dictate.
    cell_insert,
    /// A whole owned tail `<c>…</c>` removed — the clear a shrink or a
    /// block performs. Whole-element removal is this row's pinned
    /// spelling of "clear"; a keep-the-styled-shell variant waits for
    /// the Excel reference that would show Office writing one.
    cell_remove,
    /// The `<dimension ref>` value widened when created tails extend
    /// the used range. Bottom-right only, monotonic, top-left kept.
    dimension_ref_replace,

    // ── §5.8c's authoring addition (M7c) ──
    /// A whole `<f>…</f>` inserted into an existing `<c>` that has
    /// none, at the first-child position `CT_Cell` dictates (`f`
    /// precedes `v`/`is`). The self-closing shape carries its `<f>`
    /// inside `reopen_self_closing`'s replacement instead — one
    /// reopened tag, not an insertion into bytes that do not exist
    /// yet.
    f_insert,
};

pub const Edit = struct {
    /// The range of the SOURCE this replaces. `start == end` is an
    /// insertion, which is why edits order by `(start, end)` and not by
    /// `start` alone.
    at: Span,
    replacement: []const u8,
    kind: EditKind,
    cell: CellSite,
};

pub const Patch = struct {
    arena: std.heap.ArenaAllocator,
    /// The patched part.
    bytes: []const u8,
    /// Every range of the source that changed, ascending and
    /// non-overlapping. Re-applying these to the source reproduces
    /// `bytes` — which is the statement that nothing else moved.
    edits: []const Edit,

    pub fn deinit(self: *Patch) void {
        self.arena.deinit();
        self.* = undefined;
    }
};

pub const PatchResult = union(enum) {
    ok: Patch,
    refused: Refusal,
};

/// Write the projection's cached values back into the part.
///
/// Every refusal happens before a byte of output exists: the gate, the
/// §5.8b transition table and the geometry planning all run over the
/// whole projection first, so a refused patch has produced nothing to
/// roll back.
pub fn patch(self: *const ResolvedSheet, gpa: Allocator) error{OutOfMemory}!PatchResult {
    return patchWithTable(self, gpa, &spill_transitions);
}

/// `patch` with the transition table injected — the seam that lets the
/// builders, the confinement and the transaction ride be PROVEN while
/// the production table still refuses every row. Production callers
/// take `patch`; only fixtures pass a table with a reference filled in,
/// and no reference ships until a byte-diffed Excel pair is committed.
pub fn patchWithTable(
    self: *const ResolvedSheet,
    gpa: Allocator,
    table: []const SpillTransition,
) error{OutOfMemory}!PatchResult {
    var arena = std.heap.ArenaAllocator.init(gpa);
    var keep = false;
    defer if (!keep) arena.deinit();
    const a = arena.allocator();

    // Pass one: the gate and the transition table, over everything,
    // before anything is written.
    const transitions = try a.alloc(Transition, self.targets.len);
    for (self.targets, transitions) |t, *tr| {
        if (gateOf(self, t, table)) |g| {
            return .{ .refused = refuseGate(g, t.publication.row, t.publication.col) };
        }
        tr.* = transitionFor(a, t.publication.result) catch |err| switch (err) {
            error.OutOfMemory => return error.OutOfMemory,
            error.UnwritableErrorSpelling => return .{ .refused = refuseAt(
                .unwritable_error_spelling,
                t.publication.row,
                t.publication.col,
            ) },
            error.NonFiniteNumber => return .{ .refused = refuseAt(
                .non_finite_number,
                t.publication.row,
                t.publication.col,
            ) },
        };
    }

    // Still pass one: §5.8b's planned extras — anchor `ref` rewrites,
    // owned tail create/clear, dimension expansion — built and checked
    // here so every refusal below still precedes any output byte. On
    // the production table nothing plans: every DA row already refused
    // above.
    var extra: std.ArrayListUnmanaged(Edit) = .empty;
    var created: ?CellSite = null;
    for (self.targets) |t| {
        if (try planAnchorExtras(a, self, t, table, &extra)) |g| {
            return .{ .refused = refuseGate(g, t.publication.row, t.publication.col) };
        }
    }
    for (self.appends) |p| {
        switch (p.role) {
            .spill_tail => |owner| {
                // §5.8c: a tail is its anchor's — an authored formula
                // riding one is two stories about whose bytes these
                // are.
                if (p.authored != null) {
                    return .{ .refused = refuseAt(.authored_role_contradiction, p.row, p.col) };
                }
                if (tailGateFor(self, owner, table)) |g| {
                    return .{ .refused = refuseGate(g, p.row, p.col) };
                }
                if (try planTailInsert(a, self, p, &extra, &created)) |g| {
                    return .{ .refused = refuseGate(g, p.row, p.col) };
                }
            },
            // A publication with nowhere to land is not something to
            // skip, and only an OWNED tail may become a created `<c>`.
            else => return .{ .refused = refuseAt(.cell_insertion_unsupported, p.row, p.col) },
        }
    }
    if (created) |max| {
        if (try planDimension(a, self, max, &extra)) |g| {
            return .{ .refused = .{ .reason = g.reason, .cell = max, .transition = g.transition } };
        }
    }

    // Pass two: the edits.
    var edits: std.ArrayListUnmanaged(Edit) = .empty;
    for (self.targets, transitions) |t, tr| {
        try appendEdits(a, self.source, t, tr, &edits);
    }
    try edits.appendSlice(a, extra.items);
    const items = try edits.toOwnedSlice(a);
    // Stable: two tail creations at one insertion point keep the
    // column order they were planned in — `(start, end)` cannot tell
    // them apart.
    std.mem.sort(Edit, items, {}, lessThanEdit);

    // Pass three: the splice.
    const bytes = try applyEdits(a, self.source, items);

    keep = true;
    return .{ .ok = .{ .arena = arena, .bytes = bytes, .edits = items } };
}

fn lessThanEdit(_: void, x: Edit, y: Edit) bool {
    if (x.at.start != y.at.start) return x.at.start < y.at.start;
    return x.at.end < y.at.end;
}

/// What the gate answered: a reason, and the §5.8b row it is about
/// when it is about one.
const GateRefusal = struct {
    reason: Refusal.Reason,
    transition: ?SpillTransition.Id = null,
};

fn refuseGate(g: GateRefusal, row: coords.Row, col: coords.Col) Refusal {
    return .{
        .reason = g.reason,
        .cell = .{ .row = row.oneBased(), .col = col.zeroBased() },
        .transition = g.transition,
    };
}

/// §5.7.3's spill gate, §5.8b-narrowed: a covered legacy CSE passes (it
/// writes `<v>`+`t` on cells that all exist — M5b1's proven set), a DA
/// anchor consults the transition table and refuses until its row
/// carries a committed reference, and every other non-scalar shape
/// keeps refusing exactly as before M7b1.
fn gateOf(self: *const ResolvedSheet, t: Target, table: []const SpillTransition) ?GateRefusal {
    // `vm` first — M4a decision 7's order, for the same reason: a cell
    // carrying both marks must refuse on the one nothing can unlock.
    if (t.vm != 0) {
        return .{ .reason = .value_metadata_write, .transition = .value_metadata_present };
    }

    // §5.8c (M7c): an authored publication answers to the authoring
    // gate and to nothing after it — every shape the arms below judge
    // is about a formula the part already HAS, and an authored target
    // was just proven not to have one.
    if (t.publication.authored != null) {
        return authorGate(self, t, table);
    }

    if (t.formula) |f| {
        if (calc.Kind.fromAttr(f.kind) == calc.Kind.array) {
            return switch (t.publication.dialect) {
                .legacy => cseGate(self, t, f),
                .dynamic_array => daGate(t, table),
            };
        }
    }
    if (t.publication.role == .spill_tail) {
        return tailGateFor(self, t.publication.role.spill_tail, table);
    }
    if (!t.publication.shape.isScalar()) return .{ .reason = .non_scalar_result };
    return null;
}

/// A legacy CSE anchor passes only covered: every declared cell staged
/// this run, on an existing `<c>` (a target) or at least carried (an
/// append — which the append gate then names precisely). Persisting the
/// anchor while a slave keeps its old cache would have the range
/// contradict itself.
fn cseGate(self: *const ResolvedSheet, t: Target, f: decode.Formula) ?GateRefusal {
    const raw = f.ref orelse return .{ .reason = .cse_ref_unparseable };
    const range = (coords.parseRange(raw, .{
        .dollar = .accept,
        .case = .insensitive,
    }) catch return .{ .reason = .cse_ref_unparseable }).normalized();

    // The anchor is the range's top-left, or the file is telling two
    // stories about where the array starts.
    if (range.first.row.oneBased() != t.publication.row.oneBased() or
        range.first.col.zeroBased() != t.publication.col.zeroBased())
    {
        return .{ .reason = .cse_range_mismatch };
    }

    if (!declaredRangeCovered(self, range, t.publication.row, t.publication.col)) {
        return .{ .reason = .cse_range_mismatch };
    }
    return null;
}

/// Every declared cell except the anchor staged this run — target or
/// append. One derivation for the M5b1-era CSE gate and §5.8c's CSE
/// authoring gate, so the two cannot disagree about what "covered"
/// means.
fn declaredRangeCovered(
    self: *const ResolvedSheet,
    range: coords.Range,
    anchor_row: coords.Row,
    anchor_col: coords.Col,
) bool {
    var r = range.first.row.oneBased();
    while (r <= range.last.row.oneBased()) : (r += 1) {
        var c = range.first.col.zeroBased();
        while (c <= range.last.col.zeroBased()) : (c += 1) {
            if (r == anchor_row.oneBased() and c == anchor_col.zeroBased()) continue;
            const row = coords.Row.fromOneBased(r) catch unreachable;
            const col = coords.Col.fromZeroBased(c) catch unreachable;
            if (self.targetAt(row, col) != null) continue;
            if (appendAt(self, row, col) != null) continue;
            return false;
        }
    }
    return true;
}

/// §5.8c's authoring gate (M7c). The order is the order of what can
/// never be unlocked: the shapes no approved mutation addresses first,
/// then the text's own encodability, then the dialect's transition row.
fn authorGate(self: *const ResolvedSheet, t: Target, table: []const SpillTransition) ?GateRefusal {
    const w = t.publication.authored.?;

    // The modeled formula and the raw spans agree by scan construction;
    // the spans are what the EDITS answer to, and a formula-shaped run
    // of bytes the model did not carry is still one no insertion may
    // sit beside.
    if (t.spans.f != null or t.formula != null) {
        return .{ .reason = .formula_overwrite_unsupported };
    }
    if (t.cm != 0) return .{ .reason = .authored_under_cell_metadata };
    if (authoredTextRefusal(w.text)) |g| return g;

    switch (w.dialect) {
        .scalar => {
            if (t.publication.role != .plain) {
                return .{ .reason = .authored_role_contradiction };
            }
            if (!t.publication.shape.isScalar()) {
                return .{ .reason = .non_scalar_result };
            }
            return null;
        },
        .dynamic_array => {
            const outcome = switch (t.publication.role) {
                .da_anchor => |o| o,
                // M7b1 decision 7's statement, at authoring: the
                // patcher will not reconstruct a placement the model
                // already made.
                else => return .{ .reason = .dynamic_array_anchor },
            };
            const id: SpillTransition.Id = switch (outcome) {
                .spilled => .da_author_spill,
                .blocked => .da_author_blocked,
            };
            // Unconditional — reference or not (M7c decisions). The
            // authored `cm`, its XLDAPR record and the metadata part
            // around them are part-graph mutations whose BUILDER lands
            // with the reference set, so a reference alone cannot flip
            // these rows the way it flips M7b1's; the rows still sit
            // in the table because `missingReferences()` is the one
            // park list and these are parked on exactly that unblock.
            return .{ .reason = .transition_unproven, .transition = id };
        },
        .cse => |raw| {
            if (t.publication.role != .plain) {
                return .{ .reason = .authored_role_contradiction };
            }
            const range = (coords.parseRange(raw, .{
                .dollar = .accept,
                .case = .insensitive,
            }) catch return .{ .reason = .cse_ref_unparseable }).normalized();
            // The anchor is the declared top-left, and every declared
            // cell is staged this run — M7b1 decision 3's coverage
            // gate, applied before the range exists rather than after.
            if (range.first.row.oneBased() != t.publication.row.oneBased() or
                range.first.col.zeroBased() != t.publication.col.zeroBased())
            {
                return .{ .reason = .cse_range_mismatch };
            }
            if (!declaredRangeCovered(self, range, t.publication.row, t.publication.col)) {
                return .{ .reason = .cse_range_mismatch };
            }
            const row = rowById(table, .cse_author) orelse
                return .{ .reason = .transition_unproven, .transition = .cse_author };
            if (row.reference == null) {
                return .{ .reason = .transition_unproven, .transition = .cse_author };
            }
            return null;
        },
    }
}

/// What only bytes can get wrong about authored text: the FORMULA
/// carrier's own predicate (`decode.encodeAuthoredFormula`'s), applied
/// in pass one so the refusal precedes any output byte, and
/// allocation-free so the gate stays pure.
fn authoredTextRefusal(text: []const u8) ?GateRefusal {
    if (text.len == 0) return .{ .reason = .authored_text_unencodable };
    for (text) |c| {
        if (c == '\t' or c == '\n' or c == '\r') continue;
        if (c < 0x20) return .{ .reason = .authored_text_unencodable };
    }
    return null;
}

/// Which §5.8b row a DA anchor takes: the file's stored state (a cached
/// `#SPILL!` is the byte record of "was blocked") crossed with the
/// outcome the model placed. The placement itself is never re-decided
/// here — an anchor staged without it refuses generically.
fn daGate(t: Target, table: []const SpillTransition) ?GateRefusal {
    const outcome = switch (t.publication.role) {
        .da_anchor => |o| o,
        else => return .{ .reason = .dynamic_array_anchor },
    };
    const id: SpillTransition.Id = switch (outcome) {
        .spilled => if (wasBlocked(t)) .da_blocked_to_spill else .da_spill_rewrite,
        .blocked => if (wasBlocked(t)) .da_blocked_rewrite else .da_spill_to_blocked,
    };
    const row = rowById(table, id) orelse
        return .{ .reason = .transition_unproven, .transition = id };
    if (row.reference == null) {
        return .{ .reason = .transition_unproven, .transition = id };
    }
    return null;
}

/// A tail passes exactly when its anchor does: same projection, role
/// `.da_anchor`, proven row. Everything else is half a construct.
fn tailGateFor(
    self: *const ResolvedSheet,
    owner: anytype,
    table: []const SpillTransition,
) ?GateRefusal {
    const anchor = self.targetAt(owner.row, owner.col) orelse
        return .{ .reason = .tail_without_anchor };
    if (anchor.publication.role != .da_anchor) {
        return .{ .reason = .tail_without_anchor };
    }
    return daGate(anchor, table);
}

fn wasBlocked(t: Target) bool {
    return t.was == .err and t.was.err == .known and t.was.err.known == .spill;
}

fn appendAt(self: *const ResolvedSheet, row: coords.Row, col: coords.Col) ?Publication {
    var lo: usize = 0;
    var hi: usize = self.appends.len;
    while (lo < hi) {
        const mid = lo + (hi - lo) / 2;
        const p = self.appends[mid];
        if (p.row.oneBased() < row.oneBased() or
            (p.row == row and p.col.zeroBased() < col.zeroBased()))
        {
            lo = mid + 1;
        } else if (p.row == row and p.col == col) {
            return p;
        } else {
            hi = mid;
        }
    }
    return null;
}

// ─── §5.8b's planned extras (M7b1) ───────────────────────────────

/// The `f@ref` rewrite and the tail clears a PROVEN DA anchor plans.
/// On the production table this never plans anything — the gate
/// refused every DA row before it — so a planned edit here is always
/// downstream of a committed reference.
fn planAnchorExtras(
    a: Allocator,
    self: *const ResolvedSheet,
    t: Target,
    table: []const SpillTransition,
    out: *std.ArrayListUnmanaged(Edit),
) error{OutOfMemory}!?GateRefusal {
    const f = t.formula orelse return null;
    if (calc.Kind.fromAttr(f.kind) != calc.Kind.array) return null;
    if (t.publication.dialect != .dynamic_array) return null;
    const outcome = switch (t.publication.role) {
        .da_anchor => |o| o,
        else => return null,
    };
    const id: SpillTransition.Id = switch (outcome) {
        .spilled => if (wasBlocked(t)) .da_blocked_to_spill else .da_spill_rewrite,
        .blocked => if (wasBlocked(t)) .da_blocked_rewrite else .da_spill_to_blocked,
    };
    const row = rowById(table, id) orelse return null;
    if (row.reference == null) return null;

    const anchor_row = t.publication.row;
    const anchor_col = t.publication.col;
    const raw = f.ref orelse return .{ .reason = .anchor_ref_unusable };
    const ref_span = t.spans.f_ref orelse return .{ .reason = .anchor_ref_unusable };
    const old_range = (coords.parseRange(raw, .{
        .dollar = .accept,
        .case = .insensitive,
    }) catch return .{ .reason = .anchor_ref_unusable }).normalized();

    // The new extent: what the model placed, in file spelling. A
    // blocked anchor's region is itself alone — Excel's own record of
    // "nothing landed".
    const new_range: coords.Range = switch (outcome) {
        .blocked => .{
            .first = .{ .row = anchor_row, .col = anchor_col },
            .last = .{ .row = anchor_row, .col = anchor_col },
        },
        .spilled => |s| blk: {
            const lr = coords.Row.fromOneBased(anchor_row.oneBased() + s.rows - 1) catch
                return .{ .reason = .anchor_ref_unusable };
            const lc = coords.Col.fromZeroBased(anchor_col.zeroBased() + s.cols - 1) catch
                return .{ .reason = .anchor_ref_unusable };
            break :blk .{
                .first = .{ .row = anchor_row, .col = anchor_col },
                .last = .{ .row = lr, .col = lc },
            };
        },
    };

    const spelled = try spellRange(a, new_range);
    if (!std.mem.eql(u8, raw, spelled)) {
        try out.append(a, .{
            .at = ref_span,
            .replacement = spelled,
            .kind = .f_ref_replace,
            .cell = .{ .row = anchor_row.oneBased(), .col = anchor_col.zeroBased() },
        });
    }

    // Tail clears: every coordinate the STORED ref claims that the new
    // extent no longer covers. The stored ref is the file's one record
    // of the prior extent — the byte side of §5.8a's "shrink clears
    // own tails".
    var r = old_range.first.row.oneBased();
    while (r <= old_range.last.row.oneBased()) : (r += 1) {
        var c = old_range.first.col.zeroBased();
        while (c <= old_range.last.col.zeroBased()) : (c += 1) {
            if (r == anchor_row.oneBased() and c == anchor_col.zeroBased()) continue;
            const row_ = coords.Row.fromOneBased(r) catch continue;
            const col_ = coords.Col.fromZeroBased(c) catch continue;
            if (new_range.contains(.{ .row = row_, .col = col_ })) continue;
            const slot = slotAt(self.slots, row_, col_) orelse continue;
            // The record said "tail"; the bytes say formula or staged
            // target. Clearing either would remove content the record
            // does not actually own.
            if (slot.spans.f != null) return .{ .reason = .tail_clear_foreign };
            if (self.targetAt(row_, col_) != null) return .{ .reason = .tail_clear_foreign };
            try out.append(a, .{
                .at = slot.spans.cell,
                .replacement = "",
                .kind = .cell_remove,
                .cell = .{ .row = r, .col = c },
            });
        }
    }
    return null;
}

/// One owned tail becoming a created `<c>` inside an existing `<row>`,
/// plus the running bottom-right the dimension check consumes.
fn planTailInsert(
    a: Allocator,
    self: *const ResolvedSheet,
    p: Publication,
    out: *std.ArrayListUnmanaged(Edit),
    created: *?CellSite,
) error{OutOfMemory}!?GateRefusal {
    const row_1 = p.row.oneBased();
    const col_0 = p.col.zeroBased();

    const rs = uniqueRowIn(self.rows, row_1) orelse return .{ .reason = .tail_row_missing };
    if (rs.selfClosing()) return .{ .reason = .tail_row_self_closing };
    if (rs.spans_attr) |sa| {
        if (!spansCover(sa.slice(self.source), col_0 + 1)) {
            return .{ .reason = .tail_row_spans_stale };
        }
    }
    const at = tailInsertPoint(self.slots, self.rows, row_1, col_0) orelse
        return .{ .reason = .tail_row_missing };

    const tr = transitionFor(a, p.result) catch |err| switch (err) {
        error.OutOfMemory => return error.OutOfMemory,
        error.UnwritableErrorSpelling => return .{ .reason = .unwritable_error_spelling },
        error.NonFiniteNumber => return .{ .reason = .non_finite_number },
    };
    var buf: [coords.format_buf_len]u8 = undefined;
    const ref = coords.formatCell(&buf, .{ .row = p.row, .col = p.col });
    const replacement = if (tr.type_attr) |ta|
        try std.fmt.allocPrint(a, "<c r=\"{s}\" t=\"{s}\"><v>{s}</v></c>", .{ ref, ta, tr.v })
    else
        try std.fmt.allocPrint(a, "<c r=\"{s}\"><v>{s}</v></c>", .{ ref, tr.v });
    try out.append(a, .{
        .at = .{ .start = at, .end = at },
        .replacement = replacement,
        .kind = .cell_insert,
        .cell = .{ .row = row_1, .col = col_0 },
    });

    created.* = if (created.*) |m|
        .{ .row = @max(m.row, row_1), .col = @max(m.col, col_0) }
    else
        .{ .row = row_1, .col = col_0 };
    return null;
}

/// The `<dimension ref>` expansion created tails require. Absent
/// dimension: nothing to maintain — the OOXML spec lets a consumer
/// rescan `<sheetData>` (`docs/plans/structural-edits.md:100`). Present
/// but unparseable WHEN expansion is needed: the spill refuses, §5.8b's
/// "refuse until this mutation is proven possible" arm.
fn planDimension(
    a: Allocator,
    self: *const ResolvedSheet,
    max: CellSite,
    out: *std.ArrayListUnmanaged(Edit),
) error{OutOfMemory}!?GateRefusal {
    const dim = self.dimension orelse return null;
    const ref_span = dim.ref orelse return .{ .reason = .dimension_unparseable };
    const raw = ref_span.slice(self.source);
    const range = (coords.parseRange(raw, .{}) catch
        return .{ .reason = .dimension_unparseable }).normalized();

    const new_last_row = @max(range.last.row.oneBased(), max.row);
    const new_last_col = @max(range.last.col.zeroBased(), max.col);
    if (new_last_row == range.last.row.oneBased() and
        new_last_col == range.last.col.zeroBased()) return null;

    const lr = coords.Row.fromOneBased(new_last_row) catch
        return .{ .reason = .dimension_unparseable };
    const lc = coords.Col.fromZeroBased(new_last_col) catch
        return .{ .reason = .dimension_unparseable };
    // Bottom-right only, monotonic; the top-left keeps its own bytes.
    const tl_raw = if (std.mem.indexOfScalar(u8, raw, ':')) |i| raw[0..i] else raw;
    var buf: [coords.format_buf_len]u8 = undefined;
    const br = coords.formatCell(&buf, .{ .row = lr, .col = lc });
    try out.append(a, .{
        .at = ref_span,
        .replacement = try std.fmt.allocPrint(a, "{s}:{s}", .{ tl_raw, br }),
        .kind = .dimension_ref_replace,
        .cell = max,
    });
    return null;
}

/// Where a created `<c>` at (`row_1`, `col_0`) goes: after the last
/// existing `<c>` in its row with a smaller column, else at the row's
/// open end. Null when the row is missing, duplicated or self-closing —
/// the shapes §5.8b's tail creation refuses. Pure over the scan's own
/// coordinates, so the planner and the confinement checker answer from
/// one derivation.
pub fn tailInsertPoint(
    slots: []const decode.CellSlot,
    rows: []const decode.RowSlot,
    row_1: u32,
    col_0: u32,
) ?u32 {
    const rs = uniqueRowIn(rows, row_1) orelse return null;
    if (rs.selfClosing()) return null;
    var at = rs.open_end;
    for (slots) |s| {
        if (s.row.oneBased() != row_1) continue;
        if (s.col.zeroBased() < col_0) at = s.spans.cell.end;
    }
    return at;
}

fn uniqueRowIn(rows: []const decode.RowSlot, row_1: u32) ?decode.RowSlot {
    var found: ?decode.RowSlot = null;
    for (rows) |r| {
        if (r.number != row_1) continue;
        if (found != null) return null;
        found = r;
    }
    return found;
}

/// `spans="lo:hi"`, 1-based inclusive. Unparseable declares nothing —
/// and proves nothing, so it refuses at the caller.
fn spansCover(raw: []const u8, col_1: u32) bool {
    const colon = std.mem.indexOfScalar(u8, raw, ':') orelse return false;
    const lo = std.fmt.parseInt(u32, raw[0..colon], 10) catch return false;
    const hi = std.fmt.parseInt(u32, raw[colon + 1 ..], 10) catch return false;
    return col_1 >= lo and col_1 <= hi;
}

fn spellRange(a: Allocator, range: coords.Range) error{OutOfMemory}![]const u8 {
    var buf1: [coords.format_buf_len]u8 = undefined;
    var buf2: [coords.format_buf_len]u8 = undefined;
    const first = coords.formatCell(&buf1, range.first);
    if (range.first.eql(range.last)) return a.dupe(u8, first);
    const last = coords.formatCell(&buf2, range.last);
    return std.fmt.allocPrint(a, "{s}:{s}", .{ first, last });
}

/// The authored `<f>` element (§5.8c, M7c). The text was proven
/// encodable in pass one (`authoredTextRefusal`, the same predicate the
/// encoder applies), and a `.cse` ref was proven parseable at the gate,
/// so neither refusal arm is reachable here. The declared range is
/// re-spelled canonically — normalized, upper-case, `$`-free — rather
/// than echoed in caller bytes, for the same reason `planAnchorExtras`
/// spells the extent itself: the file records what the engine decided,
/// not how the caller typed it.
fn authoredFElement(a: Allocator, w: FormulaWrite) error{OutOfMemory}![]const u8 {
    const body = decode.encodeAuthoredFormula(a, w.text) catch |err| switch (err) {
        error.OutOfMemory => return error.OutOfMemory,
        error.UnencodableChar => unreachable,
    };
    return switch (w.dialect) {
        .scalar => std.fmt.allocPrint(a, "<f>{s}</f>", .{body}),
        .cse => |raw| blk: {
            const range = (coords.parseRange(raw, .{
                .dollar = .accept,
                .case = .insensitive,
            }) catch unreachable).normalized();
            const spelled = try spellRange(a, range);
            break :blk std.fmt.allocPrint(a, "<f t=\"array\" ref=\"{s}\">{s}</f>", .{ spelled, body });
        },
        // The gate refuses every `.dynamic_array` authoring at M7c —
        // the builder lands with the reference set.
        .dynamic_array => unreachable,
    };
}

fn appendEdits(
    a: Allocator,
    source: []const u8,
    t: Target,
    tr: Transition,
    out: *std.ArrayListUnmanaged(Edit),
) error{OutOfMemory}!void {
    const spans = t.spans;
    const site: CellSite = .{
        .row = t.publication.row.oneBased(),
        .col = t.publication.col.zeroBased(),
    };

    // §5.8c (M7c): the authored `<f>`, emitted BEFORE the value edits.
    // The list order is load-bearing at exactly one point: a cell with
    // no `<v>`/`<is>` takes both `f_insert` and `v_insert` at
    // `open_end`, `(start, end)` cannot order the tie, and the stable
    // sort (M7b1 decision 14) keeps `<f>` first — `CT_Cell`'s order.
    // The self-closing shape instead carries its `<f>` inside the
    // reopen replacement below.
    const authored_f: ?[]const u8 = if (t.publication.authored) |w|
        try authoredFElement(a, w)
    else
        null;
    if (authored_f) |f_elem| {
        if (!spans.selfClosing()) {
            try out.append(a, .{
                .at = .{ .start = spans.open_end, .end = spans.open_end },
                .replacement = f_elem,
                .kind = .f_insert,
                .cell = site,
            });
        }
    }

    // ── the `t` attribute ──
    if (spans.type_attr) |ts| {
        if (tr.type_attr) |want| {
            // Compared by value, not by spelling: a `t='str'` cell
            // publishing text is already the type the table asks for,
            // and rewriting it to change a quote would be a byte change
            // that says nothing.
            if (!std.mem.eql(u8, attrValueOf(ts.slice(source)), want)) {
                try out.append(a, .{
                    .at = ts,
                    .replacement = try std.fmt.allocPrint(a, "t=\"{s}\"", .{want}),
                    .kind = .type_attr_replace,
                    .cell = site,
                });
            }
        } else {
            try out.append(a, .{
                .at = removalSpan(source, spans, ts),
                .replacement = "",
                .kind = .type_attr_remove,
                .cell = site,
            });
        }
    } else if (tr.type_attr) |want| {
        const at = attrInsertPoint(source, spans.attrs);
        try out.append(a, .{
            .at = .{ .start = at, .end = at },
            .replacement = try std.fmt.allocPrint(a, " t=\"{s}\"", .{want}),
            .kind = .type_attr_insert,
            .cell = site,
        });
    }

    // ── the cached value ──
    //
    // The order below is the order `CT_Cell` allows the shapes to occur
    // in, so exactly one arm can be taken.
    if (spans.v) |vs| {
        if (spans.v_content) |vc| {
            if (!std.mem.eql(u8, vc.slice(source), tr.v)) {
                try out.append(a, .{
                    .at = vc,
                    .replacement = tr.v,
                    .kind = .v_content_replace,
                    .cell = site,
                });
            }
            return;
        }
        // A `<v/>` is replaced rather than left alone even when it
        // already reads as the empty text being published: the table
        // names `<v></v>` as the shape a text transition produces, and a
        // patcher with two answers for one row would be the drift this
        // file exists to prevent.
        try out.append(a, .{
            .at = vs,
            .replacement = try vElement(a, tr.v),
            .kind = .v_element_replace,
            .cell = site,
        });
        return;
    }

    if (spans.is) |is| {
        try out.append(a, .{
            .at = is,
            .replacement = try vElement(a, tr.v),
            .kind = .is_to_v,
            .cell = site,
        });
        return;
    }

    if (spans.selfClosing()) {
        // The last two bytes of a self-closing start tag are `/` and
        // `>`, whatever whitespace precedes them. An authored formula
        // rides the reopen (§5.8c): the cell's whole content comes
        // into being in one edit, in `CT_Cell`'s order.
        assert(spans.cell.end >= spans.cell.start + 2);
        const replacement = if (authored_f) |f_elem|
            try std.fmt.allocPrint(a, ">{s}<v>{s}</v></c>", .{ f_elem, tr.v })
        else
            try std.fmt.allocPrint(a, "><v>{s}</v></c>", .{tr.v});
        try out.append(a, .{
            .at = .{ .start = spans.cell.end - 2, .end = spans.cell.end },
            .replacement = replacement,
            .kind = .reopen_self_closing,
            .cell = site,
        });
        return;
    }

    // `<v>` follows `<f>` when there is one, and opens the content
    // otherwise.
    const at = if (spans.f) |f| f.end else spans.open_end;
    try out.append(a, .{
        .at = .{ .start = at, .end = at },
        .replacement = try vElement(a, tr.v),
        .kind = .v_insert,
        .cell = site,
    });
}

fn vElement(a: Allocator, body: []const u8) error{OutOfMemory}![]const u8 {
    return std.fmt.allocPrint(a, "<v>{s}</v>", .{body});
}

/// The value inside a raw `name="value"` attribute span. Quote-agnostic
/// because XML is: `t='b'` and `t="b"` say the same thing.
fn attrValueOf(raw: []const u8) []const u8 {
    const open = std.mem.indexOfAny(u8, raw, "\"'") orelse return "";
    const quote = raw[open];
    const rest = raw[open + 1 ..];
    const close = std.mem.indexOfScalar(u8, rest, quote) orelse return rest;
    return rest[0..close];
}

/// The bytes a `t` removal takes with it.
///
/// The attribute region is a whitespace-separated token list that always
/// *begins* with the whitespace separating it from the element name, so
/// removing a token means removing one adjacent whitespace run and which
/// run it is depends on where the token sits:
///
/// * with a neighbour to its left, take the run on the left —
///   `<c r="A1" t="s">` → `<c r="A1">`;
/// * first of several, take the run on the right, because the run on the
///   left is the element name's separator —
///   `<c t="s" r="A1">` → `<c r="A1">`;
/// * the only attribute, take both and leave the region empty —
///   `<c t="s" />` → `<c/>`.
///
/// Taking the wrong run leaves `<c  r="A1">`; taking neither leaves
/// `<c r="A1" >`; taking both in the first case leaves `<cr="A1">`.
fn removalSpan(source: []const u8, spans: decode.CellSpans, attr: Span) Span {
    var lo = attr.start;
    while (lo > spans.attrs.start and isSpace(source[lo - 1])) lo -= 1;
    var hi = attr.end;
    while (hi < spans.attrs.end and isSpace(source[hi])) hi += 1;

    const has_left = lo > spans.attrs.start;
    const has_right = hi < spans.attrs.end;
    if (has_left) return .{ .start = lo, .end = attr.end };
    if (has_right) return .{ .start = attr.start, .end = hi };
    return spans.attrs;
}

/// Where a new attribute goes: after the last one, ahead of whatever
/// whitespace a producer left between the region and the `/` or `>`.
fn attrInsertPoint(source: []const u8, attrs: Span) u32 {
    var at = attrs.end;
    while (at > attrs.start and isSpace(source[at - 1])) at -= 1;
    return at;
}

fn isSpace(c: u8) bool {
    return c == ' ' or c == '\t' or c == '\n' or c == '\r';
}

// ─── confinement ─────────────────────────────────────────────────

/// The one range an edit of this kind may address in this cell.
///
/// An equality rather than a bound. A patcher that wrote a *subrange* of
/// an approved region would still be confined, but the kind would have
/// stopped describing the edit — and the kind is what the fuzz target
/// checks the edit against. Null means the kind is not reachable for a
/// cell of this shape at all, which is itself a failure when an edit
/// claims it.
pub fn approvedRange(source: []const u8, spans: decode.CellSpans, kind: EditKind) ?Span {
    return switch (kind) {
        .type_attr_replace => spans.type_attr,
        .type_attr_remove => if (spans.type_attr) |t| removalSpan(source, spans, t) else null,
        .type_attr_insert => blk: {
            if (spans.type_attr != null) break :blk null;
            const at = attrInsertPoint(source, spans.attrs);
            break :blk .{ .start = at, .end = at };
        },
        .v_content_replace => spans.v_content,
        .v_element_replace => if (spans.v != null and spans.v_content == null) spans.v else null,
        .is_to_v => if (spans.v == null) spans.is else null,
        .reopen_self_closing => blk: {
            if (spans.v != null or spans.is != null or !spans.selfClosing()) break :blk null;
            break :blk .{ .start = spans.cell.end - 2, .end = spans.cell.end };
        },
        .v_insert => blk: {
            if (spans.v != null or spans.is != null or spans.selfClosing()) break :blk null;
            const at = if (spans.f) |f| f.end else spans.open_end;
            break :blk .{ .start = at, .end = at };
        },

        // §5.8b (M7b1). The anchor-ref exception to "no edit addresses
        // a byte inside `spans.f`" is exactly this sub-span; a tail
        // clear owns exactly its element — and an owned tail never
        // carries a formula, so a `<c>` with an `<f>` has no removable
        // range at all.
        .f_ref_replace => spans.f_ref,
        .cell_remove => if (spans.f == null) spans.cell else null,
        // §5.8c (M7c): only a cell with no `<f>` can receive one, and
        // only at the point right after its start tag — the reopened
        // self-closing shape carries its formula inside the reopen
        // edit, so the kind has no range there.
        .f_insert => blk: {
            if (spans.f != null or spans.selfClosing()) break :blk null;
            break :blk .{ .start = spans.open_end, .end = spans.open_end };
        },
        // Not answerable from one cell's spans: the insert point is the
        // ROW's geometry and the dimension is the SHEET's —
        // `verifyConfinement` answers both from `Geometry`.
        .cell_insert, .dimension_ref_replace => null,
    };
}

pub const ConfinementError = error{
    /// Two edits address overlapping bytes, so which one the output
    /// shows is a function of the sort and not of the intent.
    OverlappingEdits,
    /// An edit's range is not the one its kind is allowed to address.
    UnapprovedRange,
    /// Re-applying the edits to the source does not reproduce the
    /// output, so the edit list does not explain every changed byte.
    EditsDoNotExplainOutput,
    OutOfMemory,
};

/// What the confinement checker answers from — the scan's own
/// coordinates, passed rather than looked up through a callback, so the
/// checker cannot hold a second opinion about where anything is.
pub const Geometry = struct {
    slots: []const decode.CellSlot,
    rows: []const decode.RowSlot = &.{},
    dimension: ?decode.DimensionSpans = null,
};

/// Prove a patch changed only what it says it changed.
///
/// Two independent statements, because either alone is satisfiable by a
/// patcher that is wrong: every edit sits exactly on the range its kind
/// approves, AND replaying the edit list over the source reproduces the
/// output byte for byte. The second is what rules out a stray write; the
/// first is what rules out an edit that was honestly reported and still
/// had no business being made.
pub fn verifyConfinement(
    gpa: Allocator,
    source: []const u8,
    out: []const u8,
    edits: []const Edit,
    geo: Geometry,
) ConfinementError!void {
    var prev_end: u32 = 0;
    for (edits) |e| {
        if (e.at.start < prev_end) return error.OverlappingEdits;
        if (e.at.end < e.at.start or e.at.end > source.len) return error.UnapprovedRange;
        prev_end = e.at.end;

        const approved: ?Span = switch (e.kind) {
            // A created `<c>` may claim only the point its row's
            // geometry dictates — and only where no `<c>` exists.
            .cell_insert => blk: {
                if (spansOfSite(geo.slots, e.cell) != null) break :blk null;
                const at = tailInsertPoint(geo.slots, geo.rows, e.cell.row, e.cell.col) orelse
                    break :blk null;
                break :blk .{ .start = at, .end = at };
            },
            .dimension_ref_replace => if (geo.dimension) |d| d.ref else null,
            else => blk: {
                const spans = spansOfSite(geo.slots, e.cell) orelse break :blk null;
                break :blk approvedRange(source, spans, e.kind);
            },
        };
        const ap = approved orelse return error.UnapprovedRange;
        if (ap.start != e.at.start or ap.end != e.at.end) {
            return error.UnapprovedRange;
        }
    }

    const replayed = try applyEdits(gpa, source, edits);
    defer gpa.free(replayed);
    if (!std.mem.eql(u8, replayed, out)) return error.EditsDoNotExplainOutput;
}

fn spansOfSite(slots: []const decode.CellSlot, site: CellSite) ?decode.CellSpans {
    for (slots) |slot| {
        if (slot.row.oneBased() == site.row and slot.col.zeroBased() == site.col) {
            return slot.spans;
        }
    }
    return null;
}

/// Splice a sorted, non-overlapping edit list into the source.
pub fn applyEdits(a: Allocator, source: []const u8, edits: []const Edit) error{OutOfMemory}![]u8 {
    // The output's length is knowable before a byte moves. Sizing to
    // `source.len` and growing past it — every replacement here is
    // wider than what it replaces — stranded the whole first buffer in
    // the patch's arena until it died (§9.1).
    var total: usize = source.len;
    for (edits) |e| {
        total -= e.at.end - e.at.start;
        total += e.replacement.len;
    }

    var out: std.ArrayListUnmanaged(u8) = .empty;
    errdefer out.deinit(a);
    try out.ensureTotalCapacityPrecise(a, total);

    var cursor: u32 = 0;
    for (edits) |e| {
        assert(e.at.start >= cursor);
        try out.appendSlice(a, source[cursor..e.at.start]);
        try out.appendSlice(a, e.replacement);
        cursor = e.at.end;
    }
    try out.appendSlice(a, source[cursor..]);
    assert(out.items.len == total);
    return out.toOwnedSlice(a);
}

/// The single window in which `a` and `b` differ, as offsets into `a`,
/// or null when they are identical.
///
/// Independent of the edit list on purpose: a fixture that names the
/// span it expected to change wants an answer derived from the two
/// documents, not from the patcher's own account of what it did.
pub fn changedWindow(a: []const u8, b: []const u8) ?Span {
    var pre: usize = 0;
    while (pre < a.len and pre < b.len and a[pre] == b[pre]) pre += 1;
    if (pre == a.len and pre == b.len) return null;

    var suf: usize = 0;
    while (suf < a.len - pre and suf < b.len - pre and
        a[a.len - 1 - suf] == b[b.len - 1 - suf]) suf += 1;

    return .{ .start = @intCast(pre), .end = @intCast(a.len - suf) };
}

// ─── tests ───────────────────────────────────────────────────────

const testing = std.testing;

const ns_attr = " xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"";

fn sheetXml(comptime body: []const u8) []const u8 {
    return "<worksheet" ++ ns_attr ++ "><sheetData>" ++ body ++ "</sheetData></worksheet>";
}

const Fixture = struct {
    xml: []const u8,
    scan: decode.Sheet,
    resolved: ResolvedSheet,
    staged: *StagedDeltas,
    gpa: Allocator,

    fn deinit(self: *Fixture) void {
        self.resolved.deinit();
        self.scan.deinit();
        self.gpa.destroy(self.staged);
        self.* = undefined;
    }
};

fn buildFixture(
    gpa: Allocator,
    xml: []const u8,
    strings: []const []const u8,
    pubs: []const Publication,
) !Fixture {
    var scan = switch (try decode.scanSheet(gpa, xml, strings, .{})) {
        .ok => |s| s,
        .refused => |r| {
            std.debug.print("unexpected scan refusal: {any}\n", .{r});
            return error.TestUnexpectedRefusal;
        },
    };
    errdefer scan.deinit();

    const staged = try gpa.create(StagedDeltas);
    errdefer gpa.destroy(staged);
    staged.* = .{ .publications = pubs };

    const resolved = switch (try project(gpa, xml, &scan, staged)) {
        .ok => |r| r,
        .refused => |r| {
            std.debug.print("unexpected projection refusal: {any}\n", .{r});
            return error.TestUnexpectedRefusal;
        },
    };
    return .{ .xml = xml, .scan = scan, .resolved = resolved, .staged = staged, .gpa = gpa };
}

fn cellRef(comptime ref: []const u8) struct { row: coords.Row, col: coords.Col } {
    const c = coords.parseCell(ref, .{ .case = .insensitive }) catch unreachable;
    return .{ .row = c.row, .col = c.col };
}

fn pubAt(comptime ref: []const u8, result: value.PublishedScalar) Publication {
    const c = cellRef(ref);
    return .{ .row = c.row, .col = c.col, .result = result };
}

/// Patch one sheet, prove the confinement invariants, and hand back the
/// output. Every transition fixture goes through here, so none of them
/// can assert a value without also asserting the confinement.
fn patchOk(gpa: Allocator, f: *Fixture) ![]u8 {
    var p = switch (try patch(&f.resolved, gpa)) {
        .ok => |ok| ok,
        .refused => |r| {
            std.debug.print("unexpected patch refusal: {any}\n", .{r});
            return error.TestUnexpectedRefusal;
        },
    };
    defer p.deinit();

    try verifyConfinement(gpa, f.xml, p.bytes, p.edits, .{
        .slots = f.scan.slots,
        .rows = f.scan.rows,
        .dimension = f.scan.dimension,
    });
    return gpa.dupe(u8, p.bytes);
}

fn expectRefusal(gpa: Allocator, f: *Fixture, reason: Refusal.Reason) !void {
    switch (try patch(&f.resolved, gpa)) {
        .ok => |ok| {
            var p = ok;
            p.deinit();
            return error.TestExpectedRefusal;
        },
        .refused => |r| try testing.expectEqual(reason, r.reason),
    }
}

// ─── the transition table, row by row ────────────────────────────

test "transition: a number removes `t` and writes the shortest round-trip" {
    const xml = sheetXml(
        \\<row r="1"><c r="A1" t="str"><f>1+1</f><v>old</v></c></row>
    );
    var f = try buildFixture(testing.allocator, xml, &.{}, &.{
        pubAt("A1", .{ .number = 2 }),
    });
    defer f.deinit();

    const out = try patchOk(testing.allocator, &f);
    defer testing.allocator.free(out);
    try testing.expectEqualStrings(sheetXml(
        \\<row r="1"><c r="A1"><f>1+1</f><v>2</v></c></row>
    ), out);
}

test "transition: a number's `<v>` is the shortest text that round-trips" {
    // N5 reaches the file here: the binary64 nearest `0.1+0.2` must not
    // be written as `0.3` — a different number — and must not be written
    // with digits it does not need. The literal is spelled out because
    // Zig folds `0.1 + 0.2` at comptime in f128 and lands on 0.3 exactly,
    // which is the wrong value to be testing with.
    const xml = sheetXml(
        \\<row r="1"><c r="A1"><f>0.1+0.2</f><v>0</v></c></row>
    );
    var f = try buildFixture(testing.allocator, xml, &.{}, &.{
        pubAt("A1", .{ .number = 0.30000000000000004 }),
    });
    defer f.deinit();

    const out = try patchOk(testing.allocator, &f);
    defer testing.allocator.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "<v>0.30000000000000004</v>") != null);
}

test "transition: text becomes t=\"str\" with an ST_Xstring-encoded `<v>`" {
    const xml = sheetXml(
        \\<row r="1"><c r="A1"><f>CHAR(1)</f><v>0</v></c></row>
    );
    var f = try buildFixture(testing.allocator, xml, &.{}, &.{
        pubAt("A1", .{ .text = "\x01" }),
    });
    defer f.deinit();

    const out = try patchOk(testing.allocator, &f);
    defer testing.allocator.free(out);
    try testing.expectEqualStrings(sheetXml(
        \\<row r="1"><c r="A1" t="str"><f>CHAR(1)</f><v>_x0001_</v></c></row>
    ), out);
}

test "transition: a boolean is t=\"b\" with 1 or 0, never TRUE" {
    inline for (.{
        .{ true, "1" },
        .{ false, "0" },
    }) |case| {
        const xml = sheetXml(
            \\<row r="1"><c r="A1"><f>1=1</f><v>0</v></c></row>
        );
        var f = try buildFixture(testing.allocator, xml, &.{}, &.{
            pubAt("A1", .{ .boolean = case[0] }),
        });
        defer f.deinit();

        const out = try patchOk(testing.allocator, &f);
        defer testing.allocator.free(out);
        try testing.expectEqualStrings(sheetXml(
            "<row r=\"1\"><c r=\"A1\" t=\"b\"><f>1=1</f><v>" ++ case[1] ++ "</v></c></row>",
        ), out);
    }
}

test "transition: an error is t=\"e\" with its literal spelling" {
    const xml = sheetXml(
        \\<row r="1"><c r="A1"><f>1/0</f><v>0</v></c></row>
    );
    var f = try buildFixture(testing.allocator, xml, &.{}, &.{
        pubAt("A1", .{ .err = .{ .known = .div0 } }),
    });
    defer f.deinit();

    const out = try patchOk(testing.allocator, &f);
    defer testing.allocator.free(out);
    try testing.expectEqualStrings(sheetXml(
        \\<row r="1"><c r="A1" t="e"><f>1/0</f><v>#DIV/0!</v></c></row>
    ), out);
}

test "transition: a rich error spelling is written back byte-exact" {
    // §5.3a's "preserved, never produced": the engine did not invent
    // `#POWER_QUERY!`, and a patch that normalized it to one of the ten
    // would be inventing the *other* direction.
    const xml = sheetXml(
        \\<row r="1"><c r="A1" t="e"><f>X()</f><v>#GETTING_DATA</v></c></row>
    );
    var f = try buildFixture(testing.allocator, xml, &.{}, &.{
        pubAt("A1", .{ .err = .{ .rich = "#POWER_QUERY!" } }),
    });
    defer f.deinit();

    const out = try patchOk(testing.allocator, &f);
    defer testing.allocator.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "<v>#POWER_QUERY!</v>") != null);
    // `t` was already `e`, so the attribute is not rewritten at all.
    var p = switch (try patch(&f.resolved, testing.allocator)) {
        .ok => |ok| ok,
        .refused => return error.TestUnexpectedRefusal,
    };
    defer p.deinit();
    try testing.expectEqual(@as(usize, 1), p.edits.len);
    try testing.expectEqual(EditKind.v_content_replace, p.edits[0].kind);
}

test "transition: a blank publication caches numeric 0" {
    // The conversion is `value.publish`'s, and this is the row that
    // proves the patcher never sees the blank it would have to guess
    // about: `=A1` over an empty A1 caches `<v>0</v>`.
    const xml = sheetXml(
        \\<row r="1"><c r="B1" t="str"><f>A1</f><v>x</v></c></row>
    );
    var f = try buildFixture(testing.allocator, xml, &.{}, &.{
        .{
            .row = cellRef("B1").row,
            .col = cellRef("B1").col,
            .result = value.publish(.blank, .excel),
        },
    });
    defer f.deinit();

    const out = try patchOk(testing.allocator, &f);
    defer testing.allocator.free(out);
    try testing.expectEqualStrings(sheetXml(
        \\<row r="1"><c r="B1"><f>A1</f><v>0</v></c></row>
    ), out);
}

test "transition: every source of a blank reaches the file as the same 0" {
    // §5.7.3 names five ways a run can arrive at a blank — a reference
    // to an empty cell, a gap in a spill, an omitted argument, ISBLANK's
    // subject, and `""`. Four of them are the *evaluator's* business and
    // produce one `ScalarValue.blank` between them; the fifth is a text
    // value that happens to be empty and is a different row of the table.
    //
    // The four are one `ScalarValue.blank` by the time they leave the
    // evaluator, so what this layer can assert — and what the file
    // actually depends on — is that the conversion has exactly one
    // spelling, and that `""` does not take it.
    try testing.expect(value.PublishedScalar.eql(
        .{ .number = 0 },
        value.publish(.blank, .excel),
    ));
    try testing.expect(!value.PublishedScalar.eql(
        value.publish(.blank, .excel),
        value.publish(.{ .text = "" }, .excel),
    ));

    // And the two reach the file as different cells, so `ISBLANK`'s
    // subject and `=""` stay distinguishable after a recalc.
    const xml = sheetXml(
        \\<row r="1"><c r="B1" t="str"><f>X()</f><v>x</v></c></row>
    );
    inline for (.{
        .{ value.ScalarValue.blank, "<c r=\"B1\"><f>X()</f><v>0</v></c>" },
        .{
            value.ScalarValue{ .text = "" },
            "<c r=\"B1\" t=\"str\"><f>X()</f><v></v></c>",
        },
    }) |case| {
        var f = try buildFixture(testing.allocator, xml, &.{}, &.{
            .{
                .row = cellRef("B1").row,
                .col = cellRef("B1").col,
                .result = value.publish(case[0], .excel),
            },
        });
        defer f.deinit();

        const out = try patchOk(testing.allocator, &f);
        defer testing.allocator.free(out);
        try testing.expectEqualStrings(
            sheetXml("<row r=\"1\">" ++ case[1] ++ "</row>"),
            out,
        );
    }
}

test "transition: \"\" is t=\"str\" with an empty `<v></v>`" {
    const xml = sheetXml(
        \\<row r="1"><c r="A1"><f>""</f><v>0</v></c></row>
    );
    var f = try buildFixture(testing.allocator, xml, &.{}, &.{
        pubAt("A1", .{ .text = "" }),
    });
    defer f.deinit();

    const out = try patchOk(testing.allocator, &f);
    defer testing.allocator.free(out);
    try testing.expectEqualStrings(sheetXml(
        \\<row r="1"><c r="A1" t="str"><f>""</f><v></v></c></row>
    ), out);
}

test "projection: the raw `<f>` and its attribute region are carried, not copied" {
    // §5.7.3 asks step 3 to carry the formula. It is carried by
    // *reference* into the scan — attribute region included — so the
    // patcher has the bytes to leave alone rather than a re-rendering of
    // them it would have to prove equivalent.
    const xml = sheetXml(
        \\<row r="1"><c r="A1" t="str"><f t="shared" si="0" ref="A1:A2" ca="1">_x0041_&amp;"x"</f><v>a</v></c></row>
    );
    var f = try buildFixture(testing.allocator, xml, &.{}, &.{
        pubAt("A1", .{ .number = 1 }),
    });
    defer f.deinit();

    const t = f.resolved.targetAt(cellRef("A1").row, cellRef("A1").col).?;
    try testing.expectEqualStrings(" t=\"shared\" si=\"0\" ref=\"A1:A2\" ca=\"1\"", t.formula.?.raw_attrs);
    // A FORMULA carrier: entities decoded, ST_Xstring not applied.
    try testing.expectEqualStrings("_x0041_&\"x\"", t.formula.?.text);
    try testing.expectEqualStrings("a", t.was.text);
    try testing.expectEqualStrings(
        "<f t=\"shared\" si=\"0\" ref=\"A1:A2\" ca=\"1\">_x0041_&amp;\"x\"</f>",
        t.spans.f.?.slice(xml),
    );

    // And the patch leaves every one of those bytes where it found them.
    const out = try patchOk(testing.allocator, &f);
    defer testing.allocator.free(out);
    try testing.expect(std.mem.indexOf(u8, out, t.spans.f.?.slice(xml)) != null);
}

test "transition: a slave keeps its `<f>` in any shape and gains a `<v>`" {
    // Both slave spellings the corpus writes: the self-closing one
    // `xlsx.zig` recognizes, and the open/close pair it drops
    // (`calamine_non_monotonic_si.xlsx`). Neither `<f>` may move.
    const xml = sheetXml(
        \\<row r="1"><c r="A1"><f t="shared" ref="A1:A2" si="0">1+1</f><v>2</v></c></row>
    ++
        \\<row r="2"><c r="A2"><f t="shared" si="0"/></c><c r="B2"><f si="0" t="shared"></f></c></row>
    );
    var f = try buildFixture(testing.allocator, xml, &.{}, &.{
        pubAt("A2", .{ .number = 2 }),
        pubAt("B2", .{ .number = 2 }),
    });
    defer f.deinit();

    const out = try patchOk(testing.allocator, &f);
    defer testing.allocator.free(out);
    try testing.expectEqualStrings(sheetXml(
        \\<row r="1"><c r="A1"><f t="shared" ref="A1:A2" si="0">1+1</f><v>2</v></c></row>
    ++
        \\<row r="2"><c r="A2"><f t="shared" si="0"/><v>2</v></c><c r="B2"><f si="0" t="shared"></f><v>2</v></c></row>
    ), out);
}

test "transition: an inline string is superseded by the `<v>` that replaces it" {
    const xml = sheetXml(
        \\<row r="1"><c r="A1" t="inlineStr"><f>UPPER("a")</f><is><t>a</t></is></c></row>
    );
    var f = try buildFixture(testing.allocator, xml, &.{}, &.{
        pubAt("A1", .{ .text = "A" }),
    });
    defer f.deinit();

    const out = try patchOk(testing.allocator, &f);
    defer testing.allocator.free(out);
    try testing.expectEqualStrings(sheetXml(
        \\<row r="1"><c r="A1" t="str"><f>UPPER("a")</f><v>A</v></c></row>
    ), out);
}

test "transition: a `<v/>` is given the shape the table names" {
    const xml = sheetXml(
        \\<row r="1"><c r="A1" t="str"><f>""</f><v/></c></row>
    );
    var f = try buildFixture(testing.allocator, xml, &.{}, &.{
        pubAt("A1", .{ .text = "" }),
    });
    defer f.deinit();

    const out = try patchOk(testing.allocator, &f);
    defer testing.allocator.free(out);
    try testing.expectEqualStrings(sheetXml(
        \\<row r="1"><c r="A1" t="str"><f>""</f><v></v></c></row>
    ), out);
}

test "transition: a self-closing `<c/>` is reopened around its new value" {
    // An empty styled cell the merged view drops, published into by a
    // setCell. The slot is why the patch does not mistake it for a
    // coordinate with no `<c>`.
    const xml = sheetXml(
        \\<row r="1"><c r="A1" s="3"/></row>
    );
    var f = try buildFixture(testing.allocator, xml, &.{}, &.{
        .{
            .row = cellRef("A1").row,
            .col = cellRef("A1").col,
            .result = .{ .number = 7 },
            .origin = .set_cell,
        },
    });
    defer f.deinit();

    const out = try patchOk(testing.allocator, &f);
    defer testing.allocator.free(out);
    try testing.expectEqualStrings(sheetXml(
        \\<row r="1"><c r="A1" s="3"><v>7</v></c></row>
    ), out);
}

// ─── byte confinement, per transition ────────────────────────────

test "confinement: every transition changes exactly the ranges it named" {
    // The row's headline, as a table. Each case names the source bytes
    // it expects each edit to sit on — spelled out, so the intended span
    // is legible and not a pair of offsets — and the test then diffs the
    // two documents to find the window they actually disagree over,
    // which must fall inside those spans.
    const Case = struct {
        xml: []const u8,
        result: value.PublishedScalar,
        /// `{ kind, the source bytes it replaces }`, in byte order.
        edits: []const struct { EditKind, []const u8 },
    };
    const cases = [_]Case{
        .{
            .xml = sheetXml("<row r=\"1\"><c r=\"A1\" t=\"str\"><f>X()</f><v>a</v></c></row>"),
            .result = .{ .number = 1 },
            .edits = &.{ .{ .type_attr_remove, " t=\"str\"" }, .{ .v_content_replace, "a" } },
        },
        .{
            .xml = sheetXml("<row r=\"1\"><c r=\"A1\"><f>X()</f><v>1</v></c></row>"),
            .result = .{ .text = "a" },
            .edits = &.{ .{ .type_attr_insert, "" }, .{ .v_content_replace, "1" } },
        },
        .{
            .xml = sheetXml("<row r=\"1\"><c r=\"A1\" t=\"e\"><f>X()</f><v>#REF!</v></c></row>"),
            .result = .{ .boolean = true },
            .edits = &.{ .{ .type_attr_replace, "t=\"e\"" }, .{ .v_content_replace, "#REF!" } },
        },
        .{
            .xml = sheetXml("<row r=\"1\"><c r=\"A1\"><f>X()</f></c></row>"),
            .result = .{ .err = .{ .known = .na } },
            .edits = &.{ .{ .type_attr_insert, "" }, .{ .v_insert, "" } },
        },
        .{
            .xml = sheetXml("<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><f>X()</f><is><t>a</t></is></c></row>"),
            .result = .{ .number = 1 },
            .edits = &.{
                .{ .type_attr_remove, " t=\"inlineStr\"" },
                .{ .is_to_v, "<is><t>a</t></is>" },
            },
        },
        .{
            .xml = sheetXml("<row r=\"1\"><c r=\"A1\" s=\"1\"/></row>"),
            .result = .{ .number = 1 },
            .edits = &.{.{ .reopen_self_closing, "/>" }},
        },
        .{
            .xml = sheetXml("<row r=\"1\"><c r=\"A1\" t=\"str\"><f>X()</f><v/></c></row>"),
            .result = .{ .text = "" },
            .edits = &.{.{ .v_element_replace, "<v/>" }},
        },
    };

    for (cases) |case| {
        var f = try buildFixture(testing.allocator, case.xml, &.{}, &.{
            pubAt("A1", case.result),
        });
        defer f.deinit();

        var p = switch (try patch(&f.resolved, testing.allocator)) {
            .ok => |ok| ok,
            .refused => |r| {
                std.debug.print("unexpected refusal: {any}\n", .{r});
                return error.TestUnexpectedRefusal;
            },
        };
        defer p.deinit();

        try verifyConfinement(testing.allocator, case.xml, p.bytes, p.edits, .{
            .slots = f.scan.slots,
            .rows = f.scan.rows,
            .dimension = f.scan.dimension,
        });

        try testing.expectEqual(case.edits.len, p.edits.len);
        for (case.edits, p.edits) |want, got| {
            try testing.expectEqual(want[0], got.kind);
            // The span, in the source's own bytes. An insertion is
            // zero-length and names the empty string.
            try testing.expectEqualStrings(want[1], got.at.slice(case.xml));
        }

        // The independent half: the window the two documents disagree
        // over, derived without consulting the edit list, sits inside
        // what the edits claimed.
        const window = changedWindow(case.xml, p.bytes).?;
        try testing.expect(window.start >= p.edits[0].at.start);
        try testing.expect(window.end <= p.edits[p.edits.len - 1].at.end);
    }
}

test "confinement: a sheet's untouched cells come back byte-identical" {
    const xml = sheetXml(
        \\<row r="1" spans="1:3"><c r="A1" s="4" t="s"><v>0</v></c><c r="B1"><f>A1</f><v>1</v></c><c r="C1" t="e"><v>#N/A</v></c></row>
    ++
        \\<row r="2"><c r="A2" t="str"><f>"keep"</f><v>keep</v></c></row>
    );
    var f = try buildFixture(testing.allocator, xml, &.{"zero"}, &.{
        pubAt("B1", .{ .number = 42 }),
    });
    defer f.deinit();

    const out = try patchOk(testing.allocator, &f);
    defer testing.allocator.free(out);

    // Everything except B1's `<v>`: the styled shared-string cell, the
    // error cell, the whole second row, and the `spans` attribute the
    // patch has no business maintaining.
    try testing.expect(std.mem.indexOf(u8, out, "<c r=\"A1\" s=\"4\" t=\"s\"><v>0</v></c>") != null);
    try testing.expect(std.mem.indexOf(u8, out, "<c r=\"C1\" t=\"e\"><v>#N/A</v></c>") != null);
    try testing.expect(std.mem.indexOf(u8, out, "<row r=\"2\"><c r=\"A2\" t=\"str\"><f>\"keep\"</f><v>keep</v></c></row>") != null);
    try testing.expect(std.mem.indexOf(u8, out, "spans=\"1:3\"") != null);
    try testing.expect(std.mem.indexOf(u8, out, "<f>A1</f><v>42</v>") != null);
}

test "confinement: `t` removal takes the separator on the side with a neighbour" {
    inline for (.{
        .{ "<c r=\"A1\" t=\"str\"><f>X()</f><v>a</v></c>", "<c r=\"A1\"><f>X()</f><v>1</v></c>" },
        .{ "<c t=\"str\" r=\"A1\"><f>X()</f><v>a</v></c>", "<c r=\"A1\"><f>X()</f><v>1</v></c>" },
    }) |case| {
        const xml = sheetXml("<row r=\"1\">" ++ case[0] ++ "</row>");
        var f = try buildFixture(testing.allocator, xml, &.{}, &.{
            pubAt("A1", .{ .number = 1 }),
        });
        defer f.deinit();

        const out = try patchOk(testing.allocator, &f);
        defer testing.allocator.free(out);
        try testing.expectEqualStrings(sheetXml("<row r=\"1\">" ++ case[1] ++ "</row>"), out);
    }
}

test "confinement: an already-correct cell is not rewritten" {
    // The patcher's quietest property: publishing what a cell already
    // caches produces no edits at all, so a no-op recalc is a byte-
    // identical part.
    const xml = sheetXml(
        \\<row r="1"><c r="A1" t="str"><f>X()</f><v>a</v></c></row>
    );
    var f = try buildFixture(testing.allocator, xml, &.{}, &.{
        pubAt("A1", .{ .text = "a" }),
    });
    defer f.deinit();

    var p = switch (try patch(&f.resolved, testing.allocator)) {
        .ok => |ok| ok,
        .refused => return error.TestUnexpectedRefusal,
    };
    defer p.deinit();
    try testing.expectEqual(@as(usize, 0), p.edits.len);
    try testing.expectEqualStrings(xml, p.bytes);
}

// ─── ST_Xstring output encoding (M4b1's codec, at its writing end) ──

test "encoding: controls, a literal _x0041_, and an already-escaped _x005F_" {
    const cases = [_]struct { text: []const u8, want: []const u8 }{
        // A C0 control XML 1.0 cannot carry.
        .{ .text = "\x01", .want = "_x0001_" },
        // CR, which the corpus decided (`wdi_excel.xlsx`) and which XML
        // would normalize to LF if it were written literally.
        .{ .text = "\r", .want = "_x000D_" },
        // Tab and LF stay literal: XML carries them and Excel writes
        // them as themselves.
        .{ .text = "a\tb\nc", .want = "a\tb\nc" },
        // The rule that makes the encoding invertible.
        .{ .text = "_x0041_", .want = "_x005F_x0041_" },
        // An already-escaped underscore escapes again, because it is
        // itself text now.
        .{ .text = "_x005F_", .want = "_x005F_x005F_" },
        // XML escaping is the second pass, not a substitute for it.
        .{ .text = "a<b&c", .want = "a&lt;b&amp;c" },
    };

    for (cases) |case| {
        const xml = sheetXml(
            \\<row r="1"><c r="A1"><f>X()</f><v>0</v></c></row>
        );
        var f = try buildFixture(testing.allocator, xml, &.{}, &.{
            pubAt("A1", .{ .text = case.text }),
        });
        defer f.deinit();

        const out = try patchOk(testing.allocator, &f);
        defer testing.allocator.free(out);

        const want = try std.fmt.allocPrint(testing.allocator, "<v>{s}</v>", .{case.want});
        defer testing.allocator.free(want);
        try testing.expect(std.mem.indexOf(u8, out, want) != null);
    }
}

test "encoding: a FORMULA carrier does not take the ST_Xstring stage" {
    // The companion the row asks for. The same seven bytes on both
    // sides of one cell: `_x0041_` inside the `<f>` is a formula that
    // means those seven characters and is written back as them, while
    // `_x0041_` as a published *string* is escaped to `_x005F_x0041_`
    // so it reads back as seven characters too. One text, two sites,
    // two encodings — and the file says both.
    const xml = sheetXml(
        \\<row r="1"><c r="A1" t="str"><f>_x0041_()</f><v>old</v></c></row>
    );
    var f = try buildFixture(testing.allocator, xml, &.{}, &.{
        pubAt("A1", .{ .text = "_x0041_" }),
    });
    defer f.deinit();

    const out = try patchOk(testing.allocator, &f);
    defer testing.allocator.free(out);
    try testing.expectEqualStrings(sheetXml(
        \\<row r="1"><c r="A1" t="str"><f>_x0041_()</f><v>_x005F_x0041_</v></c></row>
    ), out);

    // And the authored-formula encoder agrees, so the asymmetry is not
    // an accident of which element the patcher happened to touch.
    const as_formula = try decode.encodeAuthoredFormula(testing.allocator, "_x0041_");
    defer testing.allocator.free(as_formula);
    try testing.expectEqualStrings("_x0041_", as_formula);
}

// ─── round-trip: what was published is what reads back ───────────

/// Re-scan a patched part and read the target coordinate back.
fn reread(gpa: Allocator, xml: []const u8, ref: []const u8) !decode.InputCell {
    var scan = switch (try decode.scanSheet(gpa, xml, &.{}, .{})) {
        .ok => |s| s,
        .refused => |r| {
            std.debug.print("patched part does not re-scan: {any}\n", .{r});
            return error.TestUnexpectedRefusal;
        },
    };
    defer scan.deinit();

    const c = try coords.parseCell(ref, .{ .case = .insensitive });
    for (scan.cells) |cell| {
        if (cell.row == c.row and cell.col == c.col) {
            // The scan arena dies with `scan`; text has to survive it.
            return switch (cell.input) {
                .text => |t| .{ .text = try gpa.dupe(u8, t) },
                .err => |e| switch (e) {
                    .rich => |t| .{ .err = .{ .rich = try gpa.dupe(u8, t) } },
                    .known => cell.input,
                },
                else => cell.input,
            };
        }
    }
    return error.TestUnexpectedResult;
}

test "round-trip: every row of the table reads back as what was published" {
    const cases = [_]struct {
        name: []const u8,
        published: value.PublishedScalar,
    }{
        .{ .name = "integer", .published = .{ .number = 42 } },
        .{ .name = "negative fraction", .published = .{ .number = -0.5 } },
        .{ .name = "N5 witness", .published = .{ .number = 0.30000000000000004 } },
        .{ .name = "large", .published = .{ .number = 1e300 } },
        .{ .name = "subnormal", .published = .{ .number = 5e-324 } },
        .{ .name = "text", .published = .{ .text = "hello" } },
        .{ .name = "empty text", .published = .{ .text = "" } },
        .{ .name = "control text", .published = .{ .text = "a\x01b" } },
        .{ .name = "escape-shaped text", .published = .{ .text = "_x0041_" } },
        .{ .name = "already-escaped text", .published = .{ .text = "_x005F_" } },
        .{ .name = "markup text", .published = .{ .text = "<&>\"'" } },
        .{ .name = "true", .published = .{ .boolean = true } },
        .{ .name = "false", .published = .{ .boolean = false } },
        .{ .name = "known error", .published = .{ .err = .{ .known = .div0 } } },
        .{ .name = "N/A", .published = .{ .err = .{ .known = .na } } },
        .{ .name = "rich error", .published = .{ .err = .{ .rich = "#POWER_QUERY!" } } },
        .{ .name = "blank publication", .published = value.publish(.blank, .excel) },
    };

    // Every starting shape a cell can be in, so the round-trip is over
    // the transition and not over one convenient predecessor.
    const starts = [_][]const u8{
        "<c r=\"A1\"><f>X()</f><v>1</v></c>",
        "<c r=\"A1\" t=\"str\"><f>X()</f><v>a</v></c>",
        "<c r=\"A1\" t=\"b\"><f>X()</f><v>1</v></c>",
        "<c r=\"A1\" t=\"e\"><f>X()</f><v>#N/A</v></c>",
        "<c r=\"A1\" t=\"inlineStr\"><f>X()</f><is><t>a</t></is></c>",
        "<c r=\"A1\"><f>X()</f></c>",
        "<c r=\"A1\" s=\"2\"/>",
    };

    for (cases) |case| {
        for (starts) |start| {
            const xml = try std.fmt.allocPrint(
                testing.allocator,
                "<worksheet{s}><sheetData><row r=\"1\">{s}</row></sheetData></worksheet>",
                .{ ns_attr, start },
            );
            defer testing.allocator.free(xml);

            var f = try buildFixture(testing.allocator, xml, &.{}, &.{
                pubAt("A1", case.published),
            });
            defer f.deinit();

            const out = try patchOk(testing.allocator, &f);
            defer testing.allocator.free(out);

            const back = try reread(testing.allocator, out, "A1");
            defer switch (back) {
                .text => |t| testing.allocator.free(t),
                .err => |e| switch (e) {
                    .rich => |t| testing.allocator.free(t),
                    .known => {},
                },
                else => {},
            };

            const got = value.publish(back.scalar(), .excel);
            if (!value.PublishedScalar.eql(case.published, got)) {
                std.debug.print(
                    "round-trip lost {s} from {s}: {any} -> {any}\n",
                    .{ case.name, start, case.published, got },
                );
                return error.TestUnexpectedResult;
            }
        }
    }
}

// ─── the pre-M7 spill gate ───────────────────────────────────────

test "gate: a non-1x1 result refuses with zero mutation" {
    const xml = sheetXml(
        \\<row r="1"><c r="A1"><f>SEQUENCE(3)</f><v>1</v></c></row>
    );
    var f = try buildFixture(testing.allocator, xml, &.{}, &.{
        .{
            .row = cellRef("A1").row,
            .col = cellRef("A1").col,
            .result = .{ .number = 1 },
            .shape = .{ .rows = 3, .cols = 1 },
        },
    });
    defer f.deinit();
    try expectRefusal(testing.allocator, &f, .non_scalar_result);
}

test "gate: a dynamic-array anchor refuses with zero mutation" {
    const xml = sheetXml(
        \\<row r="1"><c r="A1" cm="1"><f t="array" ref="A1" aca="1">UNIQUE(B:B)</f><v>1</v></c></row>
    );
    var f = try buildFixture(testing.allocator, xml, &.{}, &.{
        .{
            .row = cellRef("A1").row,
            .col = cellRef("A1").col,
            .result = .{ .number = 1 },
            .dialect = .dynamic_array,
        },
    });
    defer f.deinit();
    try expectRefusal(testing.allocator, &f, .dynamic_array_anchor);
}

test "cse: a covered legacy CSE persists — the anchor and its slaves together (M7b1)" {
    // §5.8c: M7b1 persists EXISTING CSE. Every mutation below is
    // `<v>`+`t` on a `<c>` the part already has — M5b1's proven set —
    // and no cm/vm byte exists to transition. The slave keeps its
    // bodiless place and gains the cache §5.7.3 says slaves gain.
    const xml = sheetXml(
        \\<row r="1"><c r="A1"><f t="array" ref="A1:A2">SEQUENCE(2)</f><v>999</v></c></row>
        \\<row r="2"><c r="A2"><v>999</v></c></row>
    );
    var f = try buildFixture(testing.allocator, xml, &.{}, &.{
        .{
            .row = cellRef("A1").row,
            .col = cellRef("A1").col,
            .result = .{ .number = 1 },
            .dialect = .legacy,
            .shape = .{ .rows = 2, .cols = 1 },
        },
        pubAt("A2", .{ .number = 2 }),
    });
    defer f.deinit();

    const out = try patchOk(testing.allocator, &f);
    defer testing.allocator.free(out);
    try testing.expectEqualStrings(sheetXml(
        \\<row r="1"><c r="A1"><f t="array" ref="A1:A2">SEQUENCE(2)</f><v>1</v></c></row>
        \\<row r="2"><c r="A2"><v>2</v></c></row>
    ), out);
}

test "cse: an uncovered declared range refuses with zero mutation" {
    // The anchor alone is staged; A2's stale cache would contradict it.
    const xml = sheetXml(
        \\<row r="1"><c r="A1"><f t="array" ref="A1:A2">SUM(B1:B2)</f><v>1</v></c></row>
        \\<row r="2"><c r="A2"><v>1</v></c></row>
    );
    var f = try buildFixture(testing.allocator, xml, &.{}, &.{
        .{
            .row = cellRef("A1").row,
            .col = cellRef("A1").col,
            .result = .{ .number = 1 },
            .dialect = .legacy,
        },
    });
    defer f.deinit();
    try expectRefusal(testing.allocator, &f, .cse_range_mismatch);
}

test "cse: an anchor away from its declared top-left refuses" {
    // B1 claims A1:A2 — the file tells two stories about where the
    // array starts, and the patcher believes neither.
    const xml = sheetXml(
        \\<row r="1"><c r="A1"><v>1</v></c><c r="B1"><f t="array" ref="A1:A2">X()</f><v>1</v></c></row>
        \\<row r="2"><c r="A2"><v>1</v></c></row>
    );
    var f = try buildFixture(testing.allocator, xml, &.{}, &.{
        .{
            .row = cellRef("B1").row,
            .col = cellRef("B1").col,
            .result = .{ .number = 1 },
            .dialect = .legacy,
        },
        pubAt("A1", .{ .number = 1 }),
        pubAt("A2", .{ .number = 1 }),
    });
    defer f.deinit();
    try expectRefusal(testing.allocator, &f, .cse_range_mismatch);
}

test "gate: a refused patch produces no bytes, and the source is untouched" {
    // "Zero mutation" as a byte statement rather than a promise: the
    // refusal path returns no `Patch` at all, and a second patch of the
    // same projection refuses identically — the projection did not
    // half-consume anything on its way out.
    const xml = sheetXml(
        \\<row r="1"><c r="A1" cm="1"><f t="array" ref="A1:A2">SEQUENCE(2)</f><v>1</v></c></row>
    );
    const before = try testing.allocator.dupe(u8, xml);
    defer testing.allocator.free(before);

    var f = try buildFixture(testing.allocator, xml, &.{}, &.{
        .{
            .row = cellRef("A1").row,
            .col = cellRef("A1").col,
            .result = .{ .number = 9 },
            .dialect = .dynamic_array,
        },
    });
    defer f.deinit();

    try expectRefusal(testing.allocator, &f, .dynamic_array_anchor);
    try expectRefusal(testing.allocator, &f, .dynamic_array_anchor);
    try testing.expectEqualStrings(before, f.resolved.source);
}

test "gate: an ordinary formula in the DA dialect is not an anchor" {
    // The gate refuses `t="array"`, not the dialect: a DA-dialect cell
    // whose formula is an ordinary scalar has nothing to spill and
    // nothing to refuse.
    const xml = sheetXml(
        \\<row r="1"><c r="A1" cm="1"><f>1+1</f><v>0</v></c></row>
    );
    var f = try buildFixture(testing.allocator, xml, &.{}, &.{
        .{
            .row = cellRef("A1").row,
            .col = cellRef("A1").col,
            .result = .{ .number = 2 },
            .dialect = .dynamic_array,
        },
    });
    defer f.deinit();

    const out = try patchOk(testing.allocator, &f);
    defer testing.allocator.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "<v>2</v>") != null);
}

// ─── §5.8b: the transition table and the DA path (M7b1) ──────────

/// The production table with every non-permanent reference filled in —
/// the seam fixtures prove the builders through. The path is a
/// sentinel, not a file: nothing reads it, and the PRODUCTION table is
/// separately pinned to refuse every row until a real committed
/// reference replaces a null.
const proven_table = blk: {
    var t = spill_transitions;
    for (&t) |*row| {
        if (!row.permanent) row.reference = "INJECTED-BY-TEST-not-a-committed-reference";
    }
    break :blk t;
};

fn daAnchorPub(comptime ref: []const u8, result: value.PublishedScalar, outcome: spill.Outcome) Publication {
    const c = cellRef(ref);
    return .{
        .row = c.row,
        .col = c.col,
        .result = result,
        .dialect = .dynamic_array,
        .role = .{ .da_anchor = outcome },
        .shape = switch (outcome) {
            .spilled => |s| s,
            .blocked => .{ .rows = 1, .cols = 1 },
        },
    };
}

fn tailPub(comptime ref: []const u8, comptime anchor: []const u8, result: value.PublishedScalar) Publication {
    const c = cellRef(ref);
    const a = cellRef(anchor);
    return .{
        .row = c.row,
        .col = c.col,
        .result = result,
        .role = .{ .spill_tail = .{ .row = a.row, .col = a.col } },
    };
}

/// `patchOk` against an injected table — same confinement proof.
fn patchOkWith(gpa: Allocator, f: *Fixture, table: []const SpillTransition) ![]u8 {
    var p = switch (try patchWithTable(&f.resolved, gpa, table)) {
        .ok => |ok| ok,
        .refused => |r| {
            std.debug.print("unexpected patch refusal: {any}\n", .{r});
            return error.TestUnexpectedRefusal;
        },
    };
    defer p.deinit();
    try verifyConfinement(gpa, f.xml, p.bytes, p.edits, .{
        .slots = f.scan.slots,
        .rows = f.scan.rows,
        .dimension = f.scan.dimension,
    });
    return gpa.dupe(u8, p.bytes);
}

fn expectRefusalWith(
    gpa: Allocator,
    f: *Fixture,
    table: []const SpillTransition,
    reason: Refusal.Reason,
) !void {
    switch (try patchWithTable(&f.resolved, gpa, table)) {
        .ok => |ok| {
            var p = ok;
            p.deinit();
            return error.TestExpectedRefusal;
        },
        .refused => |r| try testing.expectEqual(reason, r.reason),
    }
}

test "§5.8b: the production table refuses every row, and each refusal names its row" {
    // The four (was × now) combinations, each against the SHIPPED
    // table: this is DONE-WHEN 4's "with none committed, the whole
    // path's refusal is the pinned fixture".
    const cases = [_]struct {
        was_v: []const u8, // the stored cache: a number, or #SPILL!
        outcome: spill.Outcome,
        want: SpillTransition.Id,
    }{
        .{ .was_v = "<v>9</v>", .outcome = .{ .spilled = .{ .rows = 2, .cols = 1 } }, .want = .da_spill_rewrite },
        .{ .was_v = "<v>9</v>", .outcome = .{ .blocked = .obstruction }, .want = .da_spill_to_blocked },
        .{ .was_v = "<v>#SPILL!</v>", .outcome = .{ .spilled = .{ .rows = 2, .cols = 1 } }, .want = .da_blocked_to_spill },
        .{ .was_v = "<v>#SPILL!</v>", .outcome = .{ .blocked = .obstruction }, .want = .da_blocked_rewrite },
    };
    inline for (cases) |case| {
        const t_attr = if (comptime std.mem.indexOf(u8, case.was_v, "#SPILL!") != null) " t=\"e\"" else "";
        const xml = sheetXml(
            "<row r=\"1\"><c r=\"A1\" cm=\"1\"" ++ t_attr ++
                "><f t=\"array\" ref=\"A1\">SEQUENCE(2)</f>" ++ case.was_v ++ "</c></row>" ++
                "<row r=\"2\"></row>",
        );
        var f = try buildFixture(testing.allocator, xml, &.{}, &.{
            daAnchorPub("A1", .{ .number = 1 }, case.outcome),
        });
        defer f.deinit();

        switch (try patch(&f.resolved, testing.allocator)) {
            .ok => |ok| {
                var p = ok;
                p.deinit();
                return error.TestExpectedRefusal;
            },
            .refused => |r| {
                try testing.expectEqual(Refusal.Reason.transition_unproven, r.reason);
                try testing.expectEqual(case.want, r.transition.?);
                try testing.expectEqual(PlaneTwo.FormulaSpillPersistUnsupported, r.planeTwo());
            },
        }
    }
}

test "§5.8b: the table enumerates its refusing rows, and the park list is exact" {
    // The park list is DERIVED here, not restated: the expectation is
    // the table's own non-permanent reference-null rows in table
    // order, so a row added through the seam joins the list by
    // construction and a test that hard-coded four ids would have
    // gone stale at M7c.
    var buf: [spill_transitions.len]SpillTransition.Id = undefined;
    const missing = missingReferences(&spill_transitions, &buf);
    var expected: usize = 0;
    for (spill_transitions) |row| {
        if (row.permanent or row.reference != null) continue;
        try testing.expect(expected < missing.len);
        try testing.expectEqual(row.id, missing[expected]);
        expected += 1;
    }
    try testing.expectEqual(expected, missing.len);
    // M7b1's four rows plus M7c's three authoring rows — the one place
    // the COUNT is pinned, so a silently vanished row cannot hide
    // behind the derivation.
    try testing.expectEqual(@as(usize, 7), missing.len);
    try testing.expectEqual(SpillTransition.Id.da_spill_rewrite, missing[0]);
    try testing.expectEqual(SpillTransition.Id.cse_author, missing[6]);

    // Every row is one §5.8b/§5.8c statement, and the classification
    // is per-id: a DA persistence row is cellMetadata/XLDAPR over the
    // cell's own `cm`; an authoring row minted at the reference set;
    // the CSE authoring row touches no collection at all; the vm row
    // is permanent and never in the park list.
    for (spill_transitions) |row| {
        switch (row.id) {
            .da_spill_rewrite, .da_spill_to_blocked, .da_blocked_to_spill, .da_blocked_rewrite => {
                try testing.expect(row.reference == null);
                try testing.expectEqual(SpillTransition.Collection.cell_metadata, row.collection);
                try testing.expectEqualStrings("XLDAPR", row.record_type);
                try testing.expectEqual(SpillTransition.IndexRule.existing_cm, row.index);
                try testing.expectEqual(SpillTransition.MissingRecord.refuse, row.missing_record);
            },
            .value_metadata_present => {
                try testing.expect(row.permanent);
                try testing.expectEqual(SpillTransition.Collection.value_metadata, row.collection);
            },
            .da_author_spill, .da_author_blocked => {
                try testing.expect(row.reference == null);
                try testing.expectEqual(SpillTransition.Collection.cell_metadata, row.collection);
                try testing.expectEqualStrings("XLDAPR", row.record_type);
                try testing.expectEqual(SpillTransition.IndexRule.authored_cm, row.index);
            },
            .cse_author => {
                try testing.expect(row.reference == null);
                try testing.expectEqual(SpillTransition.Collection.none, row.collection);
                try testing.expectEqual(SpillTransition.IndexRule.none, row.index);
            },
        }
    }
}

test "§5.8b: a publication under `vm` refuses whatever the table says" {
    // A setCell publication reaches the patcher without meeting the
    // resolver — the gate refuses it here, permanently, on BOTH
    // tables.
    const xml = sheetXml(
        \\<row r="1"><c r="A1" vm="1"><v>1</v></c></row>
    );
    var f = try buildFixture(testing.allocator, xml, &.{}, &.{
        pubAt("A1", .{ .number = 2 }),
    });
    defer f.deinit();
    try expectRefusalWith(testing.allocator, &f, &proven_table, .value_metadata_write);
    try expectRefusal(testing.allocator, &f, .value_metadata_write);
}

test "§5.8b fixture: a grow rewrites the ref, fills the tail row, and touches nothing else" {
    // A1:A2 grows to A1:A3 — anchor `<v>` and `f@ref` rewrite, A2's
    // tail rewrites in place, A3 is CREATED inside its existing row.
    // `patchOkWith` proves every edit equals its approved span and the
    // edit list explains every changed byte — prior bytes outside the
    // approved spans are intact by construction.
    const xml = sheetXml(
        \\<row r="1"><c r="A1" cm="1"><f t="array" ref="A1:A2">SEQUENCE(3)</f><v>999</v></c></row>
        \\<row r="2"><c r="A2"><v>999</v></c></row>
        \\<row r="3"></row>
    );
    var f = try buildFixture(testing.allocator, xml, &.{}, &.{
        daAnchorPub("A1", .{ .number = 1 }, .{ .spilled = .{ .rows = 3, .cols = 1 } }),
        tailPub("A2", "A1", .{ .number = 2 }),
        tailPub("A3", "A1", .{ .number = 3 }),
    });
    defer f.deinit();

    const out = try patchOkWith(testing.allocator, &f, &proven_table);
    defer testing.allocator.free(out);
    try testing.expectEqualStrings(sheetXml(
        \\<row r="1"><c r="A1" cm="1"><f t="array" ref="A1:A3">SEQUENCE(3)</f><v>1</v></c></row>
        \\<row r="2"><c r="A2"><v>2</v></c></row>
        \\<row r="3"><c r="A3"><v>3</v></c></row>
    ), out);
}

test "§5.8b fixture: a shrink clears the owned tail cell whole" {
    const xml = sheetXml(
        \\<row r="1"><c r="A1" cm="1"><f t="array" ref="A1:A3">SEQUENCE(2)</f><v>999</v></c></row>
        \\<row r="2"><c r="A2"><v>999</v></c></row>
        \\<row r="3"><c r="A3"><v>999</v></c></row>
    );
    var f = try buildFixture(testing.allocator, xml, &.{}, &.{
        daAnchorPub("A1", .{ .number = 1 }, .{ .spilled = .{ .rows = 2, .cols = 1 } }),
        tailPub("A2", "A1", .{ .number = 2 }),
    });
    defer f.deinit();

    const out = try patchOkWith(testing.allocator, &f, &proven_table);
    defer testing.allocator.free(out);
    try testing.expectEqualStrings(sheetXml(
        \\<row r="1"><c r="A1" cm="1"><f t="array" ref="A1:A2">SEQUENCE(2)</f><v>1</v></c></row>
        \\<row r="2"><c r="A2"><v>2</v></c></row>
        \\<row r="3"></row>
    ), out);
}

test "§5.8b fixture: spill-to-blocked caches #SPILL! and keeps the split" {
    // The cached value becomes the bare `#SPILL!` (`t="e"`), the ref
    // collapses to the anchor, the tail clears. `cm` stays
    // byte-identical — and no `vm`, no rich metadata, is invented:
    // that is the cached-value vs rich-metadata split held.
    const xml = sheetXml(
        \\<row r="1"><c r="A1" cm="1"><f t="array" ref="A1:A2">SEQUENCE(2)</f><v>999</v></c></row>
        \\<row r="2"><c r="A2"><v>999</v></c></row>
    );
    var f = try buildFixture(testing.allocator, xml, &.{}, &.{
        daAnchorPub("A1", .{ .err = .{ .known = .spill } }, .{ .blocked = .obstruction }),
    });
    defer f.deinit();

    const out = try patchOkWith(testing.allocator, &f, &proven_table);
    defer testing.allocator.free(out);
    try testing.expectEqualStrings(sheetXml(
        \\<row r="1"><c r="A1" cm="1" t="e"><f t="array" ref="A1">SEQUENCE(2)</f><v>#SPILL!</v></c></row>
        \\<row r="2"></row>
    ), out);
}

test "§5.8b fixture: blocked-stays-blocked is byte-identical" {
    const xml = sheetXml(
        \\<row r="1"><c r="A1" cm="1" t="e"><f t="array" ref="A1">SEQUENCE(2)</f><v>#SPILL!</v></c></row>
    );
    var f = try buildFixture(testing.allocator, xml, &.{}, &.{
        daAnchorPub("A1", .{ .err = .{ .known = .spill } }, .{ .blocked = .obstruction }),
    });
    defer f.deinit();

    const out = try patchOkWith(testing.allocator, &f, &proven_table);
    defer testing.allocator.free(out);
    try testing.expectEqualStrings(xml, out);
}

test "§5.8b fixture: blocked-to-spill recovers — ref grows, tails create, in column order" {
    // Two tails into ONE empty row share an insertion point; the
    // stable sort keeps their planned column order.
    const xml = sheetXml(
        \\<row r="1"><c r="A1" cm="1" t="e"><f t="array" ref="A1">X()</f><v>#SPILL!</v></c></row>
        \\<row r="2"></row>
    );
    var f = try buildFixture(testing.allocator, xml, &.{}, &.{
        daAnchorPub("A1", .{ .number = 1 }, .{ .spilled = .{ .rows = 2, .cols = 2 } }),
        tailPub("B1", "A1", .{ .number = 2 }),
        tailPub("A2", "A1", .{ .number = 3 }),
        tailPub("B2", "A1", .{ .text = "x" }),
    });
    defer f.deinit();

    const out = try patchOkWith(testing.allocator, &f, &proven_table);
    defer testing.allocator.free(out);
    try testing.expectEqualStrings(sheetXml(
        \\<row r="1"><c r="A1" cm="1"><f t="array" ref="A1:B2">X()</f><v>1</v></c><c r="B1"><v>2</v></c></row>
        \\<row r="2"><c r="A2"><v>3</v></c><c r="B2" t="str"><v>x</v></c></row>
    ), out);
}

test "§5.8b fixture: created tails expand the dimension, bottom-right only" {
    const xml =
        "<worksheet" ++ ns_attr ++ ">" ++
        "<dimension ref=\"A1:A2\"/>" ++
        "<sheetData>" ++
        "<row r=\"1\"><c r=\"A1\" cm=\"1\"><f t=\"array\" ref=\"A1:A2\">SEQUENCE(3)</f><v>999</v></c></row>" ++
        "<row r=\"2\"><c r=\"A2\"><v>999</v></c></row>" ++
        "<row r=\"3\"></row>" ++
        "</sheetData></worksheet>";
    var f = try buildFixture(testing.allocator, xml, &.{}, &.{
        daAnchorPub("A1", .{ .number = 1 }, .{ .spilled = .{ .rows = 3, .cols = 1 } }),
        tailPub("A2", "A1", .{ .number = 2 }),
        tailPub("A3", "A1", .{ .number = 3 }),
    });
    defer f.deinit();

    const out = try patchOkWith(testing.allocator, &f, &proven_table);
    defer testing.allocator.free(out);
    try testing.expect(std.mem.indexOf(u8, out, "<dimension ref=\"A1:A3\"/>") != null);
    try testing.expect(std.mem.indexOf(u8, out, "<c r=\"A3\"><v>3</v></c>") != null);

    // A tail INSIDE the dimension expands nothing: rerun with a
    // dimension that already covers A3.
    const covered =
        "<worksheet" ++ ns_attr ++ ">" ++
        "<dimension ref=\"A1:B9\"/>" ++
        "<sheetData>" ++
        "<row r=\"1\"><c r=\"A1\" cm=\"1\"><f t=\"array\" ref=\"A1:A2\">SEQUENCE(3)</f><v>999</v></c></row>" ++
        "<row r=\"2\"><c r=\"A2\"><v>999</v></c></row>" ++
        "<row r=\"3\"></row>" ++
        "</sheetData></worksheet>";
    var f2 = try buildFixture(testing.allocator, covered, &.{}, &.{
        daAnchorPub("A1", .{ .number = 1 }, .{ .spilled = .{ .rows = 3, .cols = 1 } }),
        tailPub("A2", "A1", .{ .number = 2 }),
        tailPub("A3", "A1", .{ .number = 3 }),
    });
    defer f2.deinit();
    const out2 = try patchOkWith(testing.allocator, &f2, &proven_table);
    defer testing.allocator.free(out2);
    try testing.expect(std.mem.indexOf(u8, out2, "<dimension ref=\"A1:B9\"/>") != null);
}

test "§5.8b: tail geometry the approved set cannot address refuses, zero mutation" {
    // Into a missing row.
    {
        const xml = sheetXml(
            \\<row r="1"><c r="A1" cm="1"><f t="array" ref="A1">X()</f><v>9</v></c></row>
        );
        var f = try buildFixture(testing.allocator, xml, &.{}, &.{
            daAnchorPub("A1", .{ .number = 1 }, .{ .spilled = .{ .rows = 2, .cols = 1 } }),
            tailPub("A2", "A1", .{ .number = 2 }),
        });
        defer f.deinit();
        try expectRefusalWith(testing.allocator, &f, &proven_table, .tail_row_missing);
    }
    // Into a self-closing row — reopening it is a mutation §5.8b does
    // not name.
    {
        const xml = sheetXml(
            \\<row r="1"><c r="A1" cm="1"><f t="array" ref="A1">X()</f><v>9</v></c></row>
            \\<row r="2"/>
        );
        var f = try buildFixture(testing.allocator, xml, &.{}, &.{
            daAnchorPub("A1", .{ .number = 1 }, .{ .spilled = .{ .rows = 2, .cols = 1 } }),
            tailPub("A2", "A1", .{ .number = 2 }),
        });
        defer f.deinit();
        try expectRefusalWith(testing.allocator, &f, &proven_table, .tail_row_self_closing);
    }
    // Into a row whose `spans` would go stale.
    {
        const xml = sheetXml(
            \\<row r="1"><c r="A1" cm="1"><f t="array" ref="A1">X()</f><v>9</v></c></row>
            \\<row r="2" spans="1:1"><c r="A2"><v>1</v></c></row>
        );
        var f = try buildFixture(testing.allocator, xml, &.{}, &.{
            daAnchorPub("A1", .{ .number = 1 }, .{ .spilled = .{ .rows = 2, .cols = 2 } }),
            tailPub("B1", "A1", .{ .number = 2 }),
            tailPub("A2", "A1", .{ .number = 3 }),
            tailPub("B2", "A1", .{ .number = 4 }),
        });
        defer f.deinit();
        try expectRefusalWith(testing.allocator, &f, &proven_table, .tail_row_spans_stale);
    }
    // A spill that must expand an unparseable dimension refuses —
    // §5.8b's "refuse until this mutation is proven possible".
    {
        const xml =
            "<worksheet" ++ ns_attr ++ ">" ++
            "<dimension ref=\"GARBAGE\"/>" ++
            "<sheetData>" ++
            "<row r=\"1\"><c r=\"A1\" cm=\"1\"><f t=\"array\" ref=\"A1\">X()</f><v>9</v></c></row>" ++
            "<row r=\"2\"></row>" ++
            "</sheetData></worksheet>";
        var f = try buildFixture(testing.allocator, xml, &.{}, &.{
            daAnchorPub("A1", .{ .number = 1 }, .{ .spilled = .{ .rows = 2, .cols = 1 } }),
            tailPub("A2", "A1", .{ .number = 2 }),
        });
        defer f.deinit();
        try expectRefusalWith(testing.allocator, &f, &proven_table, .dimension_unparseable);
    }
    // A tail without its anchor in the projection is half a construct.
    {
        const xml = sheetXml(
            \\<row r="1"><c r="A1"><v>1</v></c></row>
            \\<row r="2"></row>
        );
        var f = try buildFixture(testing.allocator, xml, &.{}, &.{
            tailPub("A2", "A1", .{ .number = 2 }),
        });
        defer f.deinit();
        try expectRefusalWith(testing.allocator, &f, &proven_table, .tail_without_anchor);
    }
}

test "§5.8b: a clear the ownership record does not own refuses" {
    // The stored ref claims A2 as a tail; A2 holds a formula. Clearing
    // it would delete content the record never owned.
    const xml = sheetXml(
        \\<row r="1"><c r="A1" cm="1"><f t="array" ref="A1:A2">X()</f><v>9</v></c></row>
        \\<row r="2"><c r="A2"><f>1+1</f><v>2</v></c></row>
    );
    var f = try buildFixture(testing.allocator, xml, &.{}, &.{
        daAnchorPub("A1", .{ .number = 1 }, .{ .spilled = .{ .rows = 1, .cols = 1 } }),
    });
    defer f.deinit();
    try expectRefusalWith(testing.allocator, &f, &proven_table, .tail_clear_foreign);
}

test "§5.8b: the proven path refuses cleanly under allocation failure" {
    const xml = sheetXml(
        \\<row r="1"><c r="A1" cm="1"><f t="array" ref="A1:A2">SEQUENCE(3)</f><v>999</v></c></row>
        \\<row r="2"><c r="A2"><v>999</v></c></row>
        \\<row r="3"></row>
    );
    var f = try buildFixture(testing.allocator, xml, &.{}, &.{
        daAnchorPub("A1", .{ .number = 1 }, .{ .spilled = .{ .rows = 3, .cols = 1 } }),
        tailPub("A2", "A1", .{ .number = 2 }),
        tailPub("A3", "A1", .{ .number = 3 }),
    });
    defer f.deinit();

    try testing.checkAllAllocationFailures(testing.allocator, struct {
        fn run(gpa: Allocator, fixture: *Fixture) !void {
            var result = patchWithTable(&fixture.resolved, gpa, &proven_table) catch |err| {
                try testing.expectEqual(error.OutOfMemory, err);
                return err;
            };
            switch (result) {
                .ok => |*p| p.deinit(),
                .refused => return error.TestUnexpectedRefusal,
            }
        }
    }.run, .{&f});
}

// ─── the projection's own refusals ───────────────────────────────

test "projection: a publication with no `<c>` is carried, and the patch refuses it" {
    const xml = sheetXml(
        \\<row r="1"><c r="A1"><f>X()</f><v>1</v></c></row>
    );
    var f = try buildFixture(testing.allocator, xml, &.{}, &.{
        pubAt("A1", .{ .number = 1 }),
        pubAt("C9", .{ .number = 2 }),
    });
    defer f.deinit();

    // Carried, not dropped: the serializer path (M5c) is what will want
    // it, and a projection that silently lost it would make that row's
    // job impossible to test.
    try testing.expectEqual(@as(usize, 1), f.resolved.appends.len);
    try testing.expectEqual(@as(u32, 9), f.resolved.appends[0].row.oneBased());
    try expectRefusal(testing.allocator, &f, .cell_insertion_unsupported);
}

test "projection: two publications at one coordinate refuse" {
    const xml = sheetXml(
        \\<row r="1"><c r="A1"><f>X()</f><v>1</v></c></row>
    );
    var scan = switch (try decode.scanSheet(testing.allocator, xml, &.{}, .{})) {
        .ok => |s| s,
        .refused => return error.TestUnexpectedRefusal,
    };
    defer scan.deinit();

    var staged: StagedDeltas = .{ .publications = &.{
        pubAt("A1", .{ .number = 1 }),
        pubAt("A1", .{ .number = 2 }),
    } };
    switch (try project(testing.allocator, xml, &scan, &staged)) {
        .ok => |ok| {
            var r = ok;
            r.deinit();
            return error.TestExpectedRefusal;
        },
        .refused => |r| try testing.expectEqual(Refusal.Reason.duplicate_publication, r.reason),
    }
    // The refusal left the deltas unconsumed, so a caller that fixes the
    // cause can retry.
    try testing.expect(!staged.consumed);
}

test "projection: two `<c>` at one coordinate refuse before anything is written" {
    // `scanSheet` refuses duplicates among *modeled* cells; this pair
    // gets past it because one of them is the empty styled cell the
    // merged view drops.
    const xml = sheetXml(
        \\<row r="1"><c r="A1" s="1"/><c r="A1"><f>X()</f><v>1</v></c></row>
    );
    var scan = switch (try decode.scanSheet(testing.allocator, xml, &.{}, .{})) {
        .ok => |s| s,
        .refused => return error.TestUnexpectedRefusal,
    };
    defer scan.deinit();

    var staged: StagedDeltas = .{ .publications = &.{pubAt("A1", .{ .number = 1 })} };
    switch (try project(testing.allocator, xml, &scan, &staged)) {
        .ok => |ok| {
            var r = ok;
            r.deinit();
            return error.TestExpectedRefusal;
        },
        .refused => |r| try testing.expectEqual(Refusal.Reason.duplicate_cell_slot, r.reason),
    }
}

test "projection: a `<c>` carrying both `<v>` and `<is>` refuses" {
    const xml = sheetXml(
        \\<row r="1"><c r="A1" t="inlineStr"><f>X()</f><v>1</v><is><t>a</t></is></c></row>
    );
    var scan = switch (try decode.scanSheet(testing.allocator, xml, &.{}, .{})) {
        .ok => |s| s,
        .refused => return error.TestUnexpectedRefusal,
    };
    defer scan.deinit();

    var staged: StagedDeltas = .{ .publications = &.{pubAt("A1", .{ .number = 1 })} };
    switch (try project(testing.allocator, xml, &scan, &staged)) {
        .ok => |ok| {
            var r = ok;
            r.deinit();
            return error.TestExpectedRefusal;
        },
        .refused => |r| try testing.expectEqual(Refusal.Reason.ambiguous_cell_content, r.reason),
    }
}

test "projection: staged deltas are consumed exactly once" {
    const xml = sheetXml(
        \\<row r="1"><c r="A1"><f>X()</f><v>1</v></c></row>
    );
    var scan = switch (try decode.scanSheet(testing.allocator, xml, &.{}, .{})) {
        .ok => |s| s,
        .refused => return error.TestUnexpectedRefusal,
    };
    defer scan.deinit();

    var staged: StagedDeltas = .{ .publications = &.{pubAt("A1", .{ .number = 1 })} };
    var first = switch (try project(testing.allocator, xml, &scan, &staged)) {
        .ok => |r| r,
        .refused => return error.TestUnexpectedRefusal,
    };
    defer first.deinit();

    // A second recalc that re-applied a delta the first one already
    // wrote would publish a value nothing computed this run.
    try testing.expectError(
        error.DeltasAlreadyConsumed,
        project(testing.allocator, xml, &scan, &staged),
    );
}

test "refusal: a value no `<v>` can hold refuses rather than being written" {
    const xml = sheetXml(
        \\<row r="1"><c r="A1"><f>X()</f><v>1</v></c></row>
    );

    {
        var f = try buildFixture(testing.allocator, xml, &.{}, &.{
            pubAt("A1", .{ .err = .{ .rich = "#not a spelling" } }),
        });
        defer f.deinit();
        try expectRefusal(testing.allocator, &f, .unwritable_error_spelling);
    }
    {
        var f = try buildFixture(testing.allocator, xml, &.{}, &.{
            pubAt("A1", .{ .number = std.math.inf(f64) }),
        });
        defer f.deinit();
        try expectRefusal(testing.allocator, &f, .non_finite_number);
    }
}

test "refusals: every reason has a §10 plane, and the whole gate family shares one" {
    inline for (@typeInfo(Refusal.Reason).@"enum".fields) |field| {
        const reason: Refusal.Reason = @enumFromInt(field.value);
        _ = (Refusal{ .reason = reason }).planeTwo();
    }
    inline for (.{
        Refusal.Reason.non_scalar_result,
        Refusal.Reason.dynamic_array_anchor,
        Refusal.Reason.transition_unproven,
        Refusal.Reason.value_metadata_write,
        Refusal.Reason.cse_range_mismatch,
        Refusal.Reason.cell_insertion_unsupported,
        Refusal.Reason.tail_row_missing,
        Refusal.Reason.tail_row_self_closing,
        Refusal.Reason.tail_row_spans_stale,
        Refusal.Reason.dimension_unparseable,
        Refusal.Reason.anchor_ref_unusable,
        Refusal.Reason.tail_without_anchor,
        Refusal.Reason.tail_clear_foreign,
    }) |reason| {
        try testing.expectEqual(
            PlaneTwo.FormulaSpillPersistUnsupported,
            (Refusal{ .reason = reason }).planeTwo(),
        );
    }
}

// ─── the spans the walk hands out ────────────────────────────────

test "spans: every `<c>` gets a slot, including the ones the model drops" {
    const xml = sheetXml(
        \\<row r="1"><c r="A1" s="1"/><c r="B1"><v>1</v></c><c r="C1"></c></row>
    );
    var scan = switch (try decode.scanSheet(testing.allocator, xml, &.{}, .{})) {
        .ok => |s| s,
        .refused => return error.TestUnexpectedRefusal,
    };
    defer scan.deinit();

    try testing.expectEqual(@as(usize, 3), scan.slots.len);
    try testing.expectEqual(@as(usize, 1), scan.cells.len);
    for (scan.slots) |slot| {
        // The span is the element, exactly: it starts at a `<` and ends
        // one past a `>`.
        try testing.expectEqual(@as(u8, '<'), xml[slot.spans.cell.start]);
        try testing.expectEqual(@as(u8, '>'), xml[slot.spans.cell.end - 1]);
        try testing.expect(std.mem.startsWith(u8, slot.spans.cell.slice(xml), "<c"));
    }
    try testing.expect(scan.slots[0].spans.selfClosing());
    try testing.expect(!scan.slots[1].spans.selfClosing());
    try testing.expectEqualStrings("<v>1</v>", scan.slots[1].spans.v.?.slice(xml));
    try testing.expectEqualStrings("1", scan.slots[1].spans.v_content.?.slice(xml));
}

test "spans: implicit coordinates get slots at the reconstructed positions" {
    // A `<c>` with no `r` takes the column after its predecessor
    // (MS-OE376 §2.1.624). The patcher writes through the slot, so a
    // span at the wrong coordinate is a value in the wrong cell.
    const xml = sheetXml(
        \\<row><c><v>1</v></c><c t="str"><f>X()</f><v>a</v></c></row>
    );
    var scan = switch (try decode.scanSheet(testing.allocator, xml, &.{}, .{})) {
        .ok => |s| s,
        .refused => return error.TestUnexpectedRefusal,
    };
    defer scan.deinit();

    try testing.expectEqual(@as(usize, 2), scan.slots.len);
    try testing.expectEqual(@as(u32, 1), scan.slots[0].row.oneBased());
    try testing.expectEqual(@as(u32, 0), scan.slots[0].col.zeroBased());
    try testing.expectEqual(@as(u32, 1), scan.slots[1].col.zeroBased());
    try testing.expectEqualStrings("t=\"str\"", scan.slots[1].spans.type_attr.?.slice(xml));

    var staged: StagedDeltas = .{ .publications = &.{pubAt("B1", .{ .number = 2 })} };
    var resolved = switch (try project(testing.allocator, xml, &scan, &staged)) {
        .ok => |r| r,
        .refused => return error.TestUnexpectedRefusal,
    };
    defer resolved.deinit();

    var p = switch (try patch(&resolved, testing.allocator)) {
        .ok => |ok| ok,
        .refused => return error.TestUnexpectedRefusal,
    };
    defer p.deinit();
    try testing.expectEqualStrings(sheetXml(
        \\<row><c><v>1</v></c><c><f>X()</f><v>2</v></c></row>
    ), p.bytes);
}

test "spans: a formula's bytes are outside every approved range, except exactly the ref value" {
    // The projection carries the `<f>` and the patcher proves it: for
    // every cell shape, no approved range of any kind intersects the
    // `<f>` element — except §5.8b's ONE carve-out, `f_ref_replace`,
    // which must sit exactly on the raw ref value and nowhere else.
    const xml = sheetXml(
        \\<row r="1"><c r="A1" t="str"><f t="shared" si="0" ref="A1:A2">_x0041_&amp;"x"</f><v>a</v></c></row>
    );
    var scan = switch (try decode.scanSheet(testing.allocator, xml, &.{}, .{})) {
        .ok => |s| s,
        .refused => return error.TestUnexpectedRefusal,
    };
    defer scan.deinit();

    const spans = scan.slots[0].spans;
    const f = spans.f.?;
    inline for (@typeInfo(EditKind).@"enum".fields) |field| {
        const kind: EditKind = @enumFromInt(field.value);
        if (approvedRange(xml, spans, kind)) |r| {
            if (kind == .f_ref_replace) {
                try testing.expectEqual(spans.f_ref.?.start, r.start);
                try testing.expectEqual(spans.f_ref.?.end, r.end);
                try testing.expectEqualStrings("A1:A2", r.slice(xml));
            } else {
                const disjoint = r.end <= f.start or r.start >= f.end;
                if (!disjoint) {
                    std.debug.print("{s} may address the formula\n", .{field.name});
                    return error.TestUnexpectedResult;
                }
            }
        }
    }
}

// ─── fuzz targets ────────────────────────────────────────────────

/// Build a sheet and a publication set out of fuzzer bytes.
///
/// The bytes drive a *generator* rather than being fed in as XML: a
/// random byte string is a malformed part in essentially every draw, and
/// what this target is trying to reach is the patcher, not the scanner's
/// refusal path. The scanner has its own fuzz target for that.
const Generated = struct {
    xml: []const u8,
    pubs: []const Publication,

    fn deinit(self: Generated, a: Allocator) void {
        a.free(self.xml);
        a.free(self.pubs);
    }
};

const cell_shapes = [_][]const u8{
    "<c r=\"{s}\"><f>X()</f><v>1</v></c>",
    "<c r=\"{s}\" t=\"str\"><f>X()</f><v>a</v></c>",
    "<c r=\"{s}\" t=\"b\"><f>X()</f><v>0</v></c>",
    "<c r=\"{s}\" t=\"e\"><f>X()</f><v>#N/A</v></c>",
    "<c r=\"{s}\" t=\"inlineStr\"><f>X()</f><is><t>a</t></is></c>",
    "<c r=\"{s}\"><f>X()</f></c>",
    "<c r=\"{s}\" s=\"7\"/>",
    "<c r=\"{s}\" t=\"str\"><f>X()</f><v/></c>",
    "<c t=\"str\" r=\"{s}\" s=\"2\" ><f>X()</f><v>a</v></c>",
    "<c r=\"{s}\"><v>1</v></c>",
    // Slots the merged view drops, which is where the reopen edit lives
    // and where a `t` removal has to decide about a whitespace run with
    // no attribute on the other side of it.
    "<c r=\"{s}\"/>",
    "<c r=\"{s}\" t=\"s\"/>",
    "<c t=\"s\" r=\"{s}\" />",
    // A slave in both spellings the corpus writes.
    "<c r=\"{s}\"><f t=\"shared\" si=\"0\"/></c>",
    "<c r=\"{s}\"><f si=\"0\" t=\"shared\"></f><v>1</v></c>",
};

const published_shapes = [_]value.PublishedScalar{
    .{ .number = 0 },
    .{ .number = -1.5 },
    .{ .number = 0.30000000000000004 },
    .{ .text = "" },
    .{ .text = "a" },
    .{ .text = "_x0041_" },
    .{ .text = "\x01\r<&>" },
    .{ .boolean = true },
    .{ .boolean = false },
    .{ .err = .{ .known = .div0 } },
    .{ .err = .{ .rich = "#POWER_QUERY!" } },
};

fn generate(a: Allocator, seed: []const u8) !Generated {
    var xml: std.ArrayListUnmanaged(u8) = .empty;
    errdefer xml.deinit(a);
    var pubs: std.ArrayListUnmanaged(Publication) = .empty;
    errdefer pubs.deinit(a);

    try xml.appendSlice(a, "<worksheet");
    try xml.appendSlice(a, ns_attr);
    try xml.appendSlice(a, "><sheetData><row r=\"1\" spans=\"1:24\">");

    // Two bytes per cell: one picks the shape, one picks what gets
    // published into it. A short seed makes a short sheet, which is
    // exactly the shrink behaviour a crash report wants.
    const cells = @min(seed.len / 2, 24);
    var i: usize = 0;
    while (i < cells) : (i += 1) {
        const col = coords.Col.fromZeroBased(@intCast(i)) catch unreachable;
        var ref_buf: [16]u8 = undefined;
        const ref = coords.formatCell(&ref_buf, .{
            .row = coords.Row.fromOneBased(1) catch unreachable,
            .col = col,
        });

        const shape = cell_shapes[seed[i * 2] % cell_shapes.len];
        // `bufPrint`'s format string has to be comptime; the shape table
        // is not, so the substitution is done by hand.
        const at = std.mem.indexOf(u8, shape, "{s}").?;
        try xml.appendSlice(a, shape[0..at]);
        try xml.appendSlice(a, ref);
        try xml.appendSlice(a, shape[at + 3 ..]);

        const which = seed[i * 2 + 1];
        // Roughly one cell in eight is left unpublished, so the patch
        // has untouched cells to preserve as well as targets to write.
        if (which % 8 != 0) {
            try pubs.append(a, .{
                .row = coords.Row.fromOneBased(1) catch unreachable,
                .col = col,
                .result = published_shapes[(which / 8) % published_shapes.len],
            });
        }
    }

    try xml.appendSlice(a, "</row></sheetData></worksheet>");
    return .{ .xml = try xml.toOwnedSlice(a), .pubs = try pubs.toOwnedSlice(a) };
}

/// One generated run, patched. Null when the projection or the patch
/// refused — a refusal is a legitimate outcome for a generated sheet,
/// and the two targets below have nothing to say about one.
const Run = struct {
    gen: Generated,
    scan: decode.Sheet,
    resolved: ResolvedSheet,
    patched: Patch,

    fn deinit(self: *Run, a: Allocator) void {
        self.patched.deinit();
        self.resolved.deinit();
        self.scan.deinit();
        self.gen.deinit(a);
        self.* = undefined;
    }
};

fn runPatch(a: Allocator, seed: []const u8) !?Run {
    const gen = try generate(a, seed);
    errdefer gen.deinit(a);

    var scan = switch (try decode.scanSheet(a, gen.xml, &.{}, .{})) {
        .ok => |s| s,
        // A generator that produced an unscannable part is a bug in the
        // generator, and saying so is more useful than skipping.
        .refused => |r| {
            std.debug.print("generated part does not scan: {any}\n{s}\n", .{ r, gen.xml });
            return error.GeneratedPartDoesNotScan;
        },
    };
    errdefer scan.deinit();

    var staged: StagedDeltas = .{ .publications = gen.pubs };
    var resolved = switch (try project(a, gen.xml, &scan, &staged)) {
        .ok => |r| r,
        .refused => {
            scan.deinit();
            gen.deinit(a);
            return null;
        },
    };
    errdefer resolved.deinit();

    const patched = switch (try patch(&resolved, a)) {
        .ok => |ok| ok,
        .refused => {
            resolved.deinit();
            scan.deinit();
            gen.deinit(a);
            return null;
        },
    };
    return .{ .gen = gen, .scan = scan, .resolved = resolved, .patched = patched };
}

/// Target one's claim: the patch stayed inside the ranges it declared.
fn assertConfinement(a: Allocator, run: *const Run) !void {
    verifyConfinement(
        a,
        run.gen.xml,
        run.patched.bytes,
        run.patched.edits,
        .{
            .slots = run.scan.slots,
            .rows = run.scan.rows,
            .dimension = run.scan.dimension,
        },
    ) catch |err| {
        std.debug.print(
            "confinement broken ({t})\nsource: {s}\npatched: {s}\n",
            .{ err, run.gen.xml, run.patched.bytes },
        );
        return err;
    };
}

/// Target two's claim: every value the patch wrote reads back as itself
/// through a fresh scan of the patched part.
fn assertRoundTrip(a: Allocator, run: *const Run) !void {
    var back = switch (try decode.scanSheet(a, run.patched.bytes, &.{}, .{})) {
        .ok => |s| s,
        .refused => |r| {
            std.debug.print(
                "patched part does not re-scan: {any}\n{s}\n",
                .{ r, run.patched.bytes },
            );
            return error.PatchedPartDoesNotScan;
        },
    };
    defer back.deinit();

    for (run.resolved.targets) |t| {
        const cell = cellAt(back.cells, t.publication.row, t.publication.col) orelse
            return error.PublishedCellVanished;
        const got = value.publish(cell.input.scalar(), .excel);
        if (!value.PublishedScalar.eql(t.publication.result, got)) {
            std.debug.print(
                "round-trip lost a value at r{d}c{d}: {any} -> {any}\nsource: {s}\npatched: {s}\n",
                .{
                    t.publication.row.oneBased(), t.publication.col.zeroBased(),
                    t.publication.result,         got,
                    run.gen.xml,                  run.patched.bytes,
                },
            );
            return error.RoundTripLostAValue;
        }
    }
}

fn fuzzConfinementTarget(_: void, smith: *std.testing.Smith) anyerror!void {
    @disableInstrumentation();
    // 64 bytes is two per cell for the 24-cell ceiling, with room for
    // the generator to ignore a tail.
    var smith_buf: [64]u8 = undefined;
    const seed = smith_buf[0..smith.slice(&smith_buf)];
    var run = (try runPatch(testing.allocator, seed)) orelse return;
    defer run.deinit(testing.allocator);
    try assertConfinement(testing.allocator, &run);
}

fn fuzzRoundTripTarget(_: void, smith: *std.testing.Smith) anyerror!void {
    @disableInstrumentation();
    var smith_buf: [64]u8 = undefined;
    const seed = smith_buf[0..smith.slice(&smith_buf)];
    var run = (try runPatch(testing.allocator, seed)) orelse return;
    defer run.deinit(testing.allocator);
    try assertRoundTrip(testing.allocator, &run);
}

/// A corpus that reaches every cell shape and every published shape at
/// least once, shared by both targets: they explore the same space and
/// differ only in what they assert about it.
const fuzz_corpus = [_][]const u8{
    "",
    "\x00\x00",
    "\x06\x08", // the self-closing `<c/>` shape
    "\x04\x18", // inline string superseded by a `<v>`
    "\x07\x18", // a `<v/>` given a body
    "\x08\x00", // attributes in an unusual order, with padding
    "\x0b\x08", // `t="s"` on a slot the merged view drops
    "\x0c\x10", // the same, with `t` first and a trailing space
    "\x0d\x20\x0e\x28", // both shared-slave spellings
    "\x00\x00\x01\x08\x02\x10\x03\x18\x04\x20\x05\x28",
    "\x09\x30\x06\x38\x05\x40\x08\x48",
    "\x00\x00\x00\x00\x00\x00\x00\x00", // every cell the same
};

test "fuzz: a patch never writes outside the ranges it declares" {
    try std.testing.fuzz({}, fuzzConfinementTarget, .{ .corpus = &fuzz_corpus });
}

test "fuzz: every value a patch writes reads back as itself" {
    try std.testing.fuzz({}, fuzzRoundTripTarget, .{ .corpus = &fuzz_corpus });
}

test "fuzz: both targets' claims hold over seeded schedules" {
    // The same two assertions driven by a PRNG rather than by coverage.
    // The targets above explore; this runs a fixed, reproducible set on
    // every `zig build test`, so both invariants are gated even where
    // coverage-guided fuzzing is not available — the M5b0 pattern of one
    // body with two drivers.
    var prng = std.Random.DefaultPrng.init(0x5b1);
    for (0..256) |_| {
        var seed: [48]u8 = undefined;
        prng.random().bytes(&seed);
        const n = 2 * (1 + prng.random().uintLessThan(usize, seed.len / 2));
        var run = (try runPatch(testing.allocator, seed[0..n])) orelse continue;
        defer run.deinit(testing.allocator);
        try assertConfinement(testing.allocator, &run);
        try assertRoundTrip(testing.allocator, &run);
    }
}

test "fuzz: the confinement checker rejects a patch that lied" {
    // A checker that passed everything would make the two targets above
    // decorative, so it is checked against the two failures it exists to
    // catch: an edit outside its approved range, and an edit list that
    // does not explain the output.
    const xml = sheetXml(
        \\<row r="1"><c r="A1" t="str"><f>X()</f><v>a</v></c></row>
    );
    var scan = switch (try decode.scanSheet(testing.allocator, xml, &.{}, .{})) {
        .ok => |s| s,
        .refused => return error.TestUnexpectedRefusal,
    };
    defer scan.deinit();

    const site: CellSite = .{ .row = 1, .col = 0 };
    const spans = scan.slots[0].spans;

    // One byte to the left of the content region — inside the `<v>`
    // element, and still not the range the kind approves.
    const strayed = [_]Edit{.{
        .at = .{ .start = spans.v_content.?.start - 1, .end = spans.v_content.?.end },
        .replacement = "x",
        .kind = .v_content_replace,
        .cell = site,
    }};
    const out = try applyEdits(testing.allocator, xml, &strayed);
    defer testing.allocator.free(out);
    try testing.expectError(
        error.UnapprovedRange,
        verifyConfinement(testing.allocator, xml, out, &strayed, .{ .slots = scan.slots }),
    );

    // An honest edit list against an output that has an extra change in
    // it the list never mentions.
    const honest = [_]Edit{.{
        .at = spans.v_content.?,
        .replacement = "b",
        .kind = .v_content_replace,
        .cell = site,
    }};
    const tampered = try std.mem.concat(testing.allocator, u8, &.{ xml, "<!--x-->" });
    defer testing.allocator.free(tampered);
    try testing.expectError(
        error.EditsDoNotExplainOutput,
        verifyConfinement(testing.allocator, xml, tampered, &honest, .{ .slots = scan.slots }),
    );
}

test "allocation failure: a patch that cannot allocate refuses cleanly" {
    const xml = sheetXml(
        \\<row r="1"><c r="A1" t="str"><f>X()</f><v>a</v></c><c r="B1"><f>Y()</f></c></row>
    );
    try testing.checkAllAllocationFailures(testing.allocator, struct {
        fn run(a: Allocator, part: []const u8) !void {
            var scan = switch (try decode.scanSheet(a, part, &.{}, .{})) {
                .ok => |s| s,
                .refused => return error.TestUnexpectedRefusal,
            };
            defer scan.deinit();

            var staged: StagedDeltas = .{ .publications = &.{
                pubAt("A1", .{ .number = 1 }),
                pubAt("B1", .{ .text = "x" }),
            } };
            var resolved = switch (try project(a, part, &scan, &staged)) {
                .ok => |r| r,
                .refused => return error.TestUnexpectedRefusal,
            };
            defer resolved.deinit();

            var p = switch (try patch(&resolved, a)) {
                .ok => |ok| ok,
                .refused => return error.TestUnexpectedRefusal,
            };
            defer p.deinit();
        }
    }.run, .{xml});
}

// ─── §5.8c: authoring (M7c) ──────────────────────────────────────

const eval = @import("eval.zig");
const env_mod = @import("env.zig");
const casefold = @import("zlsx_casefold");

fn testFold(allocator: std.mem.Allocator, s: []const u8) anyerror![]u8 {
    return casefold.foldString(allocator, s);
}

fn authoredPub(comptime ref: []const u8, result: value.PublishedScalar, w: FormulaWrite) Publication {
    const c = cellRef(ref);
    return .{ .row = c.row, .col = c.col, .result = result, .authored = w };
}

/// DONE-WHEN 2's loop, at the layer that owns each leg: the patched
/// part IS the bytes a save writes (the transaction above stages them
/// verbatim), the re-open is a fresh `scanSheet` over them, and the
/// evaluation is the real parser and evaluator over the re-opened
/// stored values — so "agrees" is a statement about the file, not
/// about the fixture's memory of what it staged.
fn evaluateReopened(
    gpa: Allocator,
    part: []const u8,
    comptime authored_at: []const u8,
    formula_text: []const u8,
) !value.ScalarValue {
    var rescan = switch (try decode.scanSheet(gpa, part, &.{}, .{})) {
        .ok => |s| s,
        .refused => return error.TestUnexpectedRefusal,
    };
    defer rescan.deinit();

    const at = cellRef(authored_at);
    var authored_formula: ?decode.Formula = null;
    var fake = env_mod.Fake.init(gpa);
    defer fake.deinit();
    const sheet = try fake.addSheet("Sheet1");
    for (rescan.cells) |c| {
        if (c.row == at.row and c.col == at.col) {
            authored_formula = c.formula;
            continue;
        }
        const sv = c.input.scalar();
        if (sv == .blank) continue;
        try fake.put(sheet, .stored, .{ .row = c.row, .col = c.col, .v = sv });
    }
    const f = authored_formula orelse return error.TestExpectedFormula;
    try testing.expectEqualStrings(formula_text, f.text);

    var parsed = try parser.parse(gpa, f.text, .{});
    if (parsed == .refused) {
        parsed.deinit(gpa);
        return error.TestParseRefused;
    }
    var ast = parsed.ok;
    defer ast.deinit(gpa);

    var arena = std.heap.ArenaAllocator.init(gpa);
    defer arena.deinit();
    var draw_slot: f64 = 0;
    var draws = eval.DrawSource.constant(&draw_slot);
    var ev = eval.Evaluator.init(arena.allocator(), fake.evalEnv(), .{
        .current_sheet = sheet,
        .collation = .{ .fold = &testFold },
        .draws = &draws,
    });
    defer ev.deinit();
    const v = try ev.evaluate(ast);
    try testing.expect(v == .scalar);
    return v.scalar;
}

test "§5.8c: scalar authoring end-to-end — write, save, re-open, evaluate agrees" {
    // The authored cell is self-closing, so the write is the reopen
    // shape: `<f>` and `<v>` come into being in one edit, in
    // `CT_Cell`'s order.
    const xml = sheetXml(
        \\<row r="1"><c r="A1"><v>1</v></c><c r="B1"/></row><row r="2"><c r="A2"><v>2</v></c></row>
    );
    var f = try buildFixture(testing.allocator, xml, &.{}, &.{
        authoredPub("B1", .{ .number = 3 }, .{ .text = "SUM(A1:A2)" }),
    });
    defer f.deinit();

    const out = try patchOk(testing.allocator, &f);
    defer testing.allocator.free(out);
    try testing.expectEqualStrings(sheetXml(
        \\<row r="1"><c r="A1"><v>1</v></c><c r="B1"><f>SUM(A1:A2)</f><v>3</v></c></row><row r="2"><c r="A2"><v>2</v></c></row>
    ), out);

    const evaluated = try evaluateReopened(testing.allocator, out, "B1", "SUM(A1:A2)");
    try testing.expectEqual(@as(f64, 3), evaluated.number);

    // And the re-opened cache says what the evaluation says — the two
    // ends of "agrees".
    var rescan = switch (try decode.scanSheet(testing.allocator, out, &.{}, .{})) {
        .ok => |s| s,
        .refused => return error.TestUnexpectedRefusal,
    };
    defer rescan.deinit();
    const b1 = cellRef("B1");
    const cached = cellAt(rescan.cells, b1.row, b1.col) orelse return error.TestExpectedCell;
    try testing.expectEqual(@as(f64, 3), cached.input.number);
}

test "§5.8c: authoring onto a cached cell rides the proven kinds, carriers split" {
    // A `t`-typed text result through an existing `<v>`: the `<f>`
    // body takes the FORMULA carrier (entities only — `<` survives as
    // `&lt;`, no ST_Xstring stage) while the cached `<v>` takes the
    // STRING carrier, and the re-open reads both back to the bytes the
    // caller authored.
    const xml = sheetXml(
        \\<row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>0</v></c></row>
    );
    const text = "IF(A1<3,\"y&z\",\"n\")";
    var f = try buildFixture(testing.allocator, xml, &.{}, &.{
        authoredPub("B1", .{ .text = "y&z" }, .{ .text = text }),
    });
    defer f.deinit();

    const out = try patchOk(testing.allocator, &f);
    defer testing.allocator.free(out);
    try testing.expectEqualStrings(sheetXml(
        \\<row r="1"><c r="A1"><v>1</v></c><c r="B1" t="str"><f>IF(A1&lt;3,&quot;y&amp;z&quot;,&quot;n&quot;)</f><v>y&amp;z</v></c></row>
    ), out);

    const evaluated = try evaluateReopened(testing.allocator, out, "B1", text);
    try testing.expectEqualStrings("y&z", evaluated.text);
}

test "§5.8c: authoring next to an empty <v/> and an <is> keeps CT_Cell's order" {
    // The two remaining cached-value shapes an authored `<f>` can meet:
    // a self-closing `<v/>` (element-replaced — legal input only under
    // `t="str"`, where it caches the empty string) and an inline
    // string (superseded by `<v>`). In both, the `f_insert` lands at
    // the first-child point and the value edit stays its own kind.
    const xml = sheetXml(
        \\<row r="1"><c r="A1" t="str"><v/></c><c r="B1"><is><t>old</t></is></c></row>
    );
    var f = try buildFixture(testing.allocator, xml, &.{}, &.{
        authoredPub("A1", .{ .number = 7 }, .{ .text = "3+4" }),
        authoredPub("B1", .{ .number = 9 }, .{ .text = "4+5" }),
    });
    defer f.deinit();

    const out = try patchOk(testing.allocator, &f);
    defer testing.allocator.free(out);
    try testing.expectEqualStrings(sheetXml(
        \\<row r="1"><c r="A1"><f>3+4</f><v>7</v></c><c r="B1"><f>4+5</f><v>9</v></c></row>
    ), out);
}

test "§5.8c: the shapes authoring refuses, each by name" {
    const Case = struct {
        xml: []const u8,
        pub_: Publication,
        want: Refusal.Reason,
    };
    const cases = [_]Case{
        // A cell that already carries an `<f>` — even a bodiless one —
        // is a rewrite, not an authoring.
        .{
            .xml = sheetXml(
                \\<row r="1"><c r="B1"><f>1+1</f><v>2</v></c></row>
            ),
            .pub_ = authoredPub("B1", .{ .number = 3 }, .{ .text = "1+2" }),
            .want = .formula_overwrite_unsupported,
        },
        // `cm` names a record about a formula that is no longer there.
        .{
            .xml = sheetXml(
                \\<row r="1"><c r="B1" cm="1"><v>0</v></c></row>
            ),
            .pub_ = authoredPub("B1", .{ .number = 3 }, .{ .text = "1+2" }),
            .want = .authored_under_cell_metadata,
        },
        // `vm` refuses first, exactly as it does for a plain write.
        .{
            .xml = sheetXml(
                \\<row r="1"><c r="B1" vm="1"><v>0</v></c></row>
            ),
            .pub_ = authoredPub("B1", .{ .number = 3 }, .{ .text = "1+2" }),
            .want = .value_metadata_write,
        },
        // Text that cannot become a legal `<f>` body: empty, then a
        // control character the FORMULA carrier has no escape for.
        .{
            .xml = sheetXml(
                \\<row r="1"><c r="B1"><v>0</v></c></row>
            ),
            .pub_ = authoredPub("B1", .{ .number = 3 }, .{ .text = "" }),
            .want = .authored_text_unencodable,
        },
        .{
            .xml = sheetXml(
                \\<row r="1"><c r="B1"><v>0</v></c></row>
            ),
            .pub_ = authoredPub("B1", .{ .number = 3 }, .{ .text = "SUM(\x01)" }),
            .want = .authored_text_unencodable,
        },
        // No `<c>` to land on: authoring does not widen §5.8b's "only
        // an OWNED tail becomes a created cell".
        .{
            .xml = sheetXml(
                \\<row r="1"><c r="A1"><v>1</v></c></row>
            ),
            .pub_ = authoredPub("B1", .{ .number = 3 }, .{ .text = "1+2" }),
            .want = .cell_insertion_unsupported,
        },
        // A `.scalar` write carrying a placement role: two stories
        // about one cell.
        .{
            .xml = sheetXml(
                \\<row r="1"><c r="B1"><v>0</v></c></row>
            ),
            .pub_ = .{
                .row = cellRef("B1").row,
                .col = cellRef("B1").col,
                .result = .{ .number = 3 },
                .role = .{ .da_anchor = .{ .blocked = .obstruction } },
                .authored = .{ .text = "1+2" },
            },
            .want = .authored_role_contradiction,
        },
        // A non-1×1 result under a `.scalar` write.
        .{
            .xml = sheetXml(
                \\<row r="1"><c r="B1"><v>0</v></c></row>
            ),
            .pub_ = .{
                .row = cellRef("B1").row,
                .col = cellRef("B1").col,
                .result = .{ .number = 3 },
                .shape = .{ .rows = 2, .cols = 1 },
                .authored = .{ .text = "SEQUENCE(2)" },
            },
            .want = .non_scalar_result,
        },
    };
    for (cases) |case| {
        var f = try buildFixture(testing.allocator, case.xml, &.{}, &.{case.pub_});
        defer f.deinit();
        try expectRefusal(testing.allocator, &f, case.want);
    }
}

test "§5.8c: an authored formula staged as someone's tail refuses as a contradiction" {
    // No anchor is staged on purpose: an anchor target's own refusal
    // would precede every append (M7b1 decision 12), and what this
    // fixture pins is that the contradiction outranks even
    // `tail_without_anchor` — whose bytes these are is settled before
    // whether the owner exists.
    const xml = sheetXml(
        \\<row r="1"><c r="A1" cm="1"><f t="array" ref="A1">SEQUENCE(2)</f><v>9</v></c></row><row r="2"></row>
    );
    var f = try buildFixture(testing.allocator, xml, &.{}, &.{
        .{
            .row = cellRef("A2").row,
            .col = cellRef("A2").col,
            .result = .{ .number = 2 },
            .role = .{ .spill_tail = .{ .row = cellRef("A1").row, .col = cellRef("A1").col } },
            .authored = .{ .text = "1+1" },
        },
    });
    defer f.deinit();

    try expectRefusal(testing.allocator, &f, .authored_role_contradiction);
}

test "§5.8c: DA authoring refuses its row — and a reference alone cannot flip it" {
    const cases = [_]struct {
        outcome: spill.Outcome,
        want: SpillTransition.Id,
    }{
        .{ .outcome = .{ .spilled = .{ .rows = 2, .cols = 1 } }, .want = .da_author_spill },
        .{ .outcome = .{ .blocked = .obstruction }, .want = .da_author_blocked },
    };
    inline for (cases) |case| {
        const xml = sheetXml(
            \\<row r="1"><c r="B1"/></row><row r="2"></row>
        );
        var f = try buildFixture(testing.allocator, xml, &.{}, &.{
            .{
                .row = cellRef("B1").row,
                .col = cellRef("B1").col,
                .result = .{ .number = 1 },
                .dialect = .dynamic_array,
                .role = .{ .da_anchor = case.outcome },
                .shape = switch (case.outcome) {
                    .spilled => |s| s,
                    .blocked => .{ .rows = 1, .cols = 1 },
                },
                .authored = .{ .text = "SEQUENCE(2)", .dialect = .dynamic_array },
            },
        });
        defer f.deinit();

        // The production table parks the row…
        switch (try patch(&f.resolved, testing.allocator)) {
            .ok => return error.TestExpectedRefusal,
            .refused => |r| {
                try testing.expectEqual(Refusal.Reason.transition_unproven, r.reason);
                try testing.expectEqual(case.want, r.transition.?);
                try testing.expectEqual(PlaneTwo.FormulaSpillPersistUnsupported, r.planeTwo());
            },
        }
        // …and an injected reference parks it too (M7c decisions): the
        // authored `cm` is a part-graph mutation whose builder LANDS
        // with the reference set, so these rows do not flip the way
        // M7b1's do.
        try expectRefusalWith(testing.allocator, &f, &proven_table, .transition_unproven);
    }
}

test "§5.8c: a DA write staged without its placement refuses generically" {
    const xml = sheetXml(
        \\<row r="1"><c r="B1"/></row>
    );
    var f = try buildFixture(testing.allocator, xml, &.{}, &.{
        .{
            .row = cellRef("B1").row,
            .col = cellRef("B1").col,
            .result = .{ .number = 1 },
            .dialect = .dynamic_array,
            .authored = .{ .text = "SEQUENCE(2)", .dialect = .dynamic_array },
        },
    });
    defer f.deinit();

    try expectRefusal(testing.allocator, &f, .dynamic_array_anchor);
}

test "§5.8c: CSE authoring refuses its row on the shipped table" {
    const xml = sheetXml(
        \\<row r="1"><c r="B1"/></row><row r="2"><c r="B2"><v>0</v></c></row>
    );
    var f = try buildFixture(testing.allocator, xml, &.{}, &.{
        authoredPub("B1", .{ .number = 1 }, .{ .text = "SEQUENCE(2)", .dialect = .{ .cse = "B1:B2" } }),
        pubAt("B2", .{ .number = 2 }),
    });
    defer f.deinit();

    switch (try patch(&f.resolved, testing.allocator)) {
        .ok => return error.TestExpectedRefusal,
        .refused => |r| {
            try testing.expectEqual(Refusal.Reason.transition_unproven, r.reason);
            try testing.expectEqual(SpillTransition.Id.cse_author, r.transition.?);
        },
    }
}

test "§5.8c: a proven CSE authoring writes the anchor's declaration and the covered range" {
    // Injected-table proof of the builder, M7b1's own seam: the anchor
    // takes `<f t="array" ref>` — spelled canonically, not in caller
    // bytes — plus its cache; the slave rides M5b1's proven kinds; the
    // confinement holds over all of it. The committed reference
    // re-pins what Office writes AROUND one when it lands.
    const xml = sheetXml(
        \\<row r="1"><c r="B1"/></row><row r="2"><c r="B2"><v>0</v></c></row>
    );
    var f = try buildFixture(testing.allocator, xml, &.{}, &.{
        authoredPub("B1", .{ .number = 1 }, .{ .text = "SEQUENCE(2)", .dialect = .{ .cse = "b1:b2" } }),
        pubAt("B2", .{ .number = 2 }),
    });
    defer f.deinit();

    const out = try patchOkWith(testing.allocator, &f, &proven_table);
    defer testing.allocator.free(out);
    try testing.expectEqualStrings(sheetXml(
        \\<row r="1"><c r="B1"><f t="array" ref="B1:B2">SEQUENCE(2)</f><v>1</v></c></row><row r="2"><c r="B2"><v>2</v></c></row>
    ), out);
}

test "§5.8c: CSE authoring geometry refuses before the row's proof state" {
    const Case = struct {
        anchor: []const u8,
        ref: []const u8,
        want: Refusal.Reason,
    };
    // Uncovered slave, anchor off the top-left, unparseable range —
    // all contradictions about geometry, so they name themselves on
    // BOTH tables rather than hiding behind `transition_unproven`.
    const cases = [_]Case{
        .{ .anchor = "B1", .ref = "B1:B3", .want = .cse_range_mismatch },
        .{ .anchor = "B2", .ref = "B1:B2", .want = .cse_range_mismatch },
        .{ .anchor = "B1", .ref = "not-a-range", .want = .cse_ref_unparseable },
    };
    inline for (cases) |case| {
        const xml = sheetXml(
            \\<row r="1"><c r="B1"/></row><row r="2"><c r="B2"><v>0</v></c></row>
        );
        var f = try buildFixture(testing.allocator, xml, &.{}, &.{
            authoredPub(case.anchor, .{ .number = 1 }, .{ .text = "SEQUENCE(2)", .dialect = .{ .cse = case.ref } }),
            pubAt(if (std.mem.eql(u8, case.anchor, "B1")) "B2" else "B1", .{ .number = 2 }),
        });
        defer f.deinit();

        try expectRefusal(testing.allocator, &f, case.want);
        try expectRefusalWith(testing.allocator, &f, &proven_table, case.want);
    }
}

/// One authoring attempt per table row, for the enumeration test: the
/// switch is exhaustive over `SpillTransition.Id`, so a row added to
/// the table without an attempt here fails to COMPILE — which is what
/// keeps the refusal enumeration derived from the table rather than
/// from prose.
fn expectAuthoringRowRefusal(comptime id: SpillTransition.Id) !void {
    const xml = sheetXml(
        \\<row r="1"><c r="B1"/></row><row r="2"><c r="B2"><v>0</v></c></row>
    );
    const pubs: []const Publication = switch (id) {
        .da_author_spill, .da_author_blocked => &.{.{
            .row = cellRef("B1").row,
            .col = cellRef("B1").col,
            .result = .{ .number = 1 },
            .dialect = .dynamic_array,
            .role = .{ .da_anchor = if (id == .da_author_spill)
                .{ .spilled = .{ .rows = 2, .cols = 1 } }
            else
                .{ .blocked = .obstruction } },
            .shape = if (id == .da_author_spill) .{ .rows = 2, .cols = 1 } else .{ .rows = 1, .cols = 1 },
            .authored = .{ .text = "SEQUENCE(2)", .dialect = .dynamic_array },
        }},
        .cse_author => &.{
            authoredPub("B1", .{ .number = 1 }, .{ .text = "SEQUENCE(2)", .dialect = .{ .cse = "B1:B2" } }),
            pubAt("B2", .{ .number = 2 }),
        },
        else => @compileError("not an authoring row: " ++ @tagName(id)),
    };
    var f = try buildFixture(testing.allocator, xml, &.{}, pubs);
    defer f.deinit();

    switch (try patch(&f.resolved, testing.allocator)) {
        .ok => return error.TestExpectedRefusal,
        .refused => |r| {
            try testing.expectEqual(Refusal.Reason.transition_unproven, r.reason);
            try testing.expectEqual(id, r.transition.?);
        },
    }
}

test "§5.8c: the unproven-authoring list derives from the table, and each refusal names its row" {
    var buf: [spill_transitions.len]SpillTransition.Id = undefined;
    const missing = missingReferences(&spill_transitions, &buf);
    var authoring_rows: usize = 0;
    inline for (spill_transitions) |row| {
        // Exhaustive over `Id` — see `expectAuthoringRowRefusal`.
        switch (row.id) {
            // M7b1's persistence rows: their per-row refusals are
            // pinned in "the production table refuses every row"
            // above; here they only assert their park-list membership.
            .da_spill_rewrite, .da_spill_to_blocked, .da_blocked_to_spill, .da_blocked_rewrite => {
                try testing.expect(std.mem.indexOfScalar(SpillTransition.Id, missing, row.id) != null);
            },
            .value_metadata_present => {
                try testing.expect(std.mem.indexOfScalar(SpillTransition.Id, missing, row.id) == null);
            },
            .da_author_spill, .da_author_blocked, .cse_author => {
                try testing.expect(std.mem.indexOfScalar(SpillTransition.Id, missing, row.id) != null);
                try expectAuthoringRowRefusal(row.id);
                authoring_rows += 1;
            },
        }
    }
    try testing.expectEqual(@as(usize, 3), authoring_rows);
}

test {
    // `parser` is imported for the §10 taxonomy `decode` re-exports; the
    // reference keeps a build that drops the import honest.
    _ = parser.PlaneTwo;
}
