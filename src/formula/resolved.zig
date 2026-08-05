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
//! What this row does NOT do
//! -------------------------
//! Inserting a `<c>` the part does not have is outside the approved
//! mutation set (§5.8b): it would leave `<row spans>` and `<dimension>`
//! describing a sheet that no longer exists, and those maintenances are
//! M7b1's with their own byte-diffed proofs. The projection still
//! *carries* such a publication — `ResolvedSheet.appends` — because the
//! serializer path (M5c) will want it; the patcher refuses rather than
//! dropping it, so a staged delta can never go missing quietly.
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

    pub const Reason = enum {
        // ── the pre-M7 spill gate (§5.7.3) ──
        /// The staged result is not 1×1. Persisting it means writing the
        /// tail cells it spills into, and the mutation set that covers
        /// those is M7b1's.
        non_scalar_result,
        /// The target's `<f t="array" ref=…>` is a dynamic-array anchor
        /// (M4a resolved its `cm`/`vm` to the DA dialect). Its cached
        /// value and its spill region are one construct; writing half of
        /// it is what §5.8b refuses until the other half is proven.
        dynamic_array_anchor,
        /// The target's `<f t="array" ref=…>` is a legacy CSE array. The
        /// same reason, one dialect over: the declared range is part of
        /// the result.
        cse_array,
        /// The publication lands where the part has no `<c>` at all.
        /// Inserting one is outside the approved mutation set until
        /// `<row spans>` and `<dimension>` maintenance is proven (M7b1).
        cell_insertion_unsupported,

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
            .cse_array,
            .cell_insertion_unsupported,
            => .FormulaSpillPersistUnsupported,

            .duplicate_cell_slot,
            .duplicate_publication,
            .ambiguous_cell_content,
            .unwritable_error_spelling,
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
    /// Carried, never patched — see the module header.
    appends: []const Publication,

    pub fn deinit(self: *ResolvedSheet) void {
        self.arena.deinit();
        self.* = undefined;
    }

    /// The target at a coordinate, or null. Linear because a projection
    /// stages the cells one run touched, not the sheet.
    pub fn targetAt(self: *const ResolvedSheet, row: coords.Row, col: coords.Col) ?Target {
        for (self.targets) |t| {
            if (t.publication.row == row and t.publication.col == col) return t;
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

    const ordered = try a.dupe(Publication, pubs);
    std.mem.sortUnstable(Publication, ordered, {}, lessThanPublication);
    i = 1;
    while (i < ordered.len) : (i += 1) {
        const cur = ordered[i];
        const prev = ordered[i - 1];
        if (cur.row == prev.row and cur.col == prev.col) {
            return .{ .refused = refuseAt(.duplicate_publication, cur.row, cur.col) };
        }
    }

    var targets: std.ArrayListUnmanaged(Target) = .empty;
    var appends: std.ArrayListUnmanaged(Publication) = .empty;
    for (ordered) |p| {
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
        });
    }

    _ = try staged.consume();
    keep = true;
    return .{ .ok = .{
        .arena = arena,
        .source = source,
        .targets = try targets.toOwnedSlice(a),
        .appends = try appends.toOwnedSlice(a),
    } };
}

fn lessThanPublication(_: void, x: Publication, y: Publication) bool {
    if (x.row.oneBased() != y.row.oneBased()) return x.row.oneBased() < y.row.oneBased();
    return x.col.zeroBased() < y.col.zeroBased();
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
/// Every refusal happens before a byte of output exists: the gate and
/// the transition table run over the whole projection first, so a
/// refused patch has produced nothing to roll back.
pub fn patch(self: *const ResolvedSheet, gpa: Allocator) error{OutOfMemory}!PatchResult {
    var arena = std.heap.ArenaAllocator.init(gpa);
    var keep = false;
    defer if (!keep) arena.deinit();
    const a = arena.allocator();

    // A publication with nowhere to land is not something to skip.
    if (self.appends.len > 0) {
        const p = self.appends[0];
        return .{ .refused = refuseAt(.cell_insertion_unsupported, p.row, p.col) };
    }

    // Pass one: the pre-M7 gate and the transition table, over
    // everything, before anything is written.
    const transitions = try a.alloc(Transition, self.targets.len);
    for (self.targets, transitions) |t, *tr| {
        if (gateOf(t)) |reason| {
            return .{ .refused = refuseAt(reason, t.publication.row, t.publication.col) };
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

    // Pass two: the edits.
    var edits: std.ArrayListUnmanaged(Edit) = .empty;
    for (self.targets, transitions) |t, tr| {
        try appendEdits(a, self.source, t, tr, &edits);
    }
    const items = try edits.toOwnedSlice(a);
    std.mem.sortUnstable(Edit, items, {}, lessThanEdit);

    // Pass three: the splice.
    const bytes = try applyEdits(a, self.source, items);

    keep = true;
    return .{ .ok = .{ .arena = arena, .bytes = bytes, .edits = items } };
}

fn lessThanEdit(_: void, x: Edit, y: Edit) bool {
    if (x.at.start != y.at.start) return x.at.start < y.at.start;
    return x.at.end < y.at.end;
}

/// The pre-M7 spill gate: the three shapes whose persistence needs a
/// mutation set this row does not have.
fn gateOf(t: Target) ?Refusal.Reason {
    if (!t.publication.shape.isScalar()) return .non_scalar_result;
    const f = t.formula orelse return null;
    if (calc.Kind.fromAttr(f.kind) != calc.Kind.array) return null;
    // Both arms are the same refusal in every way except which
    // construct it names, and naming it is what makes a diagnostic
    // worth reading.
    return switch (t.publication.dialect) {
        .dynamic_array => .dynamic_array_anchor,
        .legacy => .cse_array,
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
        // `>`, whatever whitespace precedes them.
        assert(spans.cell.end >= spans.cell.start + 2);
        try out.append(a, .{
            .at = .{ .start = spans.cell.end - 2, .end = spans.cell.end },
            .replacement = try std.fmt.allocPrint(a, "><v>{s}</v></c>", .{tr.v}),
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
    /// The scan's slots — the same ones the projection was built from.
    /// Passed rather than looked up through a callback, so the checker
    /// answers from the walk's own coordinates and not from a second
    /// opinion about where a `<c>` is.
    slots: []const decode.CellSlot,
) ConfinementError!void {
    var prev_end: u32 = 0;
    for (edits) |e| {
        if (e.at.start < prev_end) return error.OverlappingEdits;
        if (e.at.end < e.at.start or e.at.end > source.len) return error.UnapprovedRange;
        prev_end = e.at.end;

        const spans = spansOfSite(slots, e.cell) orelse return error.UnapprovedRange;
        const approved = approvedRange(source, spans, e.kind) orelse return error.UnapprovedRange;
        if (approved.start != e.at.start or approved.end != e.at.end) {
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
    var out: std.ArrayListUnmanaged(u8) = .empty;
    errdefer out.deinit(a);
    try out.ensureTotalCapacity(a, source.len);

    var cursor: u32 = 0;
    for (edits) |e| {
        assert(e.at.start >= cursor);
        try out.appendSlice(a, source[cursor..e.at.start]);
        try out.appendSlice(a, e.replacement);
        cursor = e.at.end;
    }
    try out.appendSlice(a, source[cursor..]);
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

    try verifyConfinement(gpa, f.xml, p.bytes, p.edits, f.scan.slots);
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

        try verifyConfinement(testing.allocator, case.xml, p.bytes, p.edits, f.scan.slots);

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

test "gate: a legacy CSE array refuses with zero mutation" {
    const xml = sheetXml(
        \\<row r="1"><c r="A1"><f t="array" ref="A1:A2">SUM(B1:B2)</f><v>1</v></c></row>
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
    try expectRefusal(testing.allocator, &f, .cse_array);
}

test "gate: a refused patch produces no bytes, and the source is untouched" {
    // "Zero mutation" as a byte statement rather than a promise: the
    // refusal path returns no `Patch` at all, and a second patch of the
    // same projection refuses identically — the projection did not
    // half-consume anything on its way out.
    const xml = sheetXml(
        \\<row r="1"><c r="A1"><f t="array" ref="A1:A2">SUM(B1:B2)</f><v>1</v></c></row>
    );
    const before = try testing.allocator.dupe(u8, xml);
    defer testing.allocator.free(before);

    var f = try buildFixture(testing.allocator, xml, &.{}, &.{
        pubAt("A1", .{ .number = 9 }),
    });
    defer f.deinit();

    try expectRefusal(testing.allocator, &f, .cse_array);
    try expectRefusal(testing.allocator, &f, .cse_array);
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

test "refusals: every reason has a §10 plane, and the gate's four share one" {
    inline for (@typeInfo(Refusal.Reason).@"enum".fields) |field| {
        const reason: Refusal.Reason = @enumFromInt(field.value);
        _ = (Refusal{ .reason = reason }).planeTwo();
    }
    inline for (.{
        Refusal.Reason.non_scalar_result,
        Refusal.Reason.dynamic_array_anchor,
        Refusal.Reason.cse_array,
        Refusal.Reason.cell_insertion_unsupported,
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

test "spans: a formula's bytes are never inside an approved range" {
    // The projection carries the `<f>` and the patcher proves it: for
    // every cell shape, no approved range of any kind intersects the
    // `<f>` element.
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
            const disjoint = r.end <= f.start or r.start >= f.end;
            if (!disjoint) {
                std.debug.print("{s} may address the formula\n", .{field.name});
                return error.TestUnexpectedResult;
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
        run.scan.slots,
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
        verifyConfinement(testing.allocator, xml, out, &strayed, scan.slots),
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
        verifyConfinement(testing.allocator, xml, tampered, &honest, scan.slots),
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

test {
    // `parser` is imported for the §10 taxonomy `decode` re-exports; the
    // reference keeps a build that drops the import honest.
    _ = parser.PlaneTwo;
}
