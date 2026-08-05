//! §5.7.6's calc-state writes, and §5.7.7's mark-only eligibility.
//!
//! M5b2 of the tier-D1 ladder (`goal_formula.md`). `calc.zig` reads what
//! `xl/workbook.xml` says about how a workbook calculates; this file
//! writes the two things a recalc is allowed to change about it, and
//! nothing else.
//!
//! Why a patcher and not a serializer
//! ----------------------------------
//! `CalcState.writeCalcPr` reconstructs `<calcPr>` from its parsed
//! pieces, which is the right shape for a producer emitting a fresh
//! part. It is the wrong shape here. A recalc rewrites two attributes of
//! one element in a part that may carry defined names, external
//! references, pivot caches and vendor extensions this engine does not
//! model — and "everything else survived" is not a testable sentence
//! about a serializer, only a hope. So this file emits an **edit list**,
//! exactly as M5b1's cached-value patcher does, and `approvedRange` maps
//! each edit's kind to the single byte range it may address.
//!
//! What a recalc may change (§5.7.6, truthful producer state)
//! ---------------------------------------------------------
//! `calcId="0"` and `fullCalcOnLoad="1"`, and that is the whole list.
//! `calcId` becomes zero — the "unknown or older producer" value —
//! because preserving Excel's id would claim Excel produced zlsx's
//! caches, and Excel trusts `calcId`/`calcFeatures` when deciding whether
//! to recalculate. `fullCalcOnLoad` is set so Excel re-verifies on open,
//! harmlessly; the non-Excel consumers this engine targets read the fresh
//! `<v>` either way. `calcMode`, the iteration triple, `calcOnSave`,
//! `forceFullCalc`, `refMode`, `concurrentCalc`, `calcFeatures`, the CV
//! extension and `sheetCalcPr@fullCalcOnLoad` are all preserved, which
//! here means *not addressed by any edit*: the bytes are never visited,
//! so preservation is structural rather than a re-emission that has to be
//! checked. `fullPrecision="0"` never reaches this file — `calc.zig`
//! refuses it at parse.
//!
//! Absent elements are created only when needed, and `CT_Workbook` is an
//! `xsd:sequence`, so a created `<calcPr>` goes at its schema position:
//! before the first successor element the part actually has, or before
//! `</workbook>`. Both coordinates come from `calc.CalcPrSpans`, which
//! the parser recorded on its way past — one parser, and it hands out the
//! ranges it used.

const std = @import("std");
const assert = std.debug.assert;
const Allocator = std.mem.Allocator;

const decode = @import("decode.zig");
const calc = @import("calc.zig");
const resolved = @import("resolved.zig");

pub const Span = decode.Span;
pub const PlaneTwo = decode.PlaneTwo;

/// The two documents' single differing window. Re-exported from M5b1
/// rather than re-derived: a fixture that names the range it expected to
/// change wants an answer computed from the two byte strings, and there
/// is no second way to compute it that would be worth having.
pub const changedWindow = resolved.changedWindow;

// ─── what a run wants the calc state to say ──────────────────────

/// The desired post-run calc state, as the two independent bits it
/// actually is.
///
/// A struct rather than an enum with three cases because the two bits are
/// set by two different §5.7 rules — step 6's truthful-producer state and
/// step 7's mark-only path — and an enum would make "which rule set this"
/// unanswerable at the point the edits are built.
pub const Desired = struct {
    /// Write `calcId="0"`.
    calc_id_zero: bool = false,
    /// Write `fullCalcOnLoad="1"`.
    full_calc_on_load: bool = false,

    /// §5.7.6 after a successful recalc: both.
    pub const after_recalc: Desired = .{ .calc_id_zero = true, .full_calc_on_load = true };

    /// §5.7.7's `.keep_stale_and_mark`, and `Workbook.markRecalcOnLoad()`.
    /// The mark alone: the caches are the ones the file already had, so
    /// claiming a new producer for them would be the lie `calcId="0"`
    /// exists to avoid telling.
    pub const mark_only: Desired = .{ .full_calc_on_load = true };

    pub fn isEmpty(self: Desired) bool {
        return !self.calc_id_zero and !self.full_calc_on_load;
    }
};

/// The value `calcId` takes after a zlsx recalc. Named because it appears
/// in the emitted bytes, in the approved-range table, and in the
/// byte-level assertion the row is gated on.
pub const calc_id_value = "0";
pub const full_calc_on_load_value = "1";

const calc_id_attr = " calcId=\"" ++ calc_id_value ++ "\"";
const full_calc_attr = " fullCalcOnLoad=\"" ++ full_calc_on_load_value ++ "\"";

// ─── the patch ───────────────────────────────────────────────────

/// What one edit did. As in M5b1, the kind is not a label: it selects the
/// one range `approvedRange` permits, so an edit whose range does not
/// equal its kind's range is a bug a fixture catches rather than a
/// difference a reader has to notice.
pub const EditKind = enum {
    /// An existing `calcId="…"` run replaced whole — name, `=`, quotes
    /// and all. Replacing only the value would leave the quote style to
    /// be guessed at; replacing the run states it.
    calc_id_replace,
    /// `calcId="0"` inserted into a `<calcPr>` that had no `calcId`.
    calc_id_insert,
    /// An existing `fullCalcOnLoad="…"` run replaced whole.
    full_calc_on_load_replace,
    /// `fullCalcOnLoad="1"` inserted into a `<calcPr>` that had none.
    full_calc_on_load_insert,
    /// A whole `<calcPr …/>` created at its `CT_Workbook` sequence
    /// position, for a workbook that had no calc state at all.
    calc_pr_create,
};

pub const Edit = struct {
    /// The source range this replaces. `start == end` is an insertion.
    at: Span,
    replacement: []const u8,
    kind: EditKind,
};

pub const Plan = struct {
    /// The patched part. Zero edits means this aliases the source and
    /// `owns_bytes` is false — a no-op plan allocates nothing, which is
    /// what makes "a recalc that changes no calc state changes no bytes"
    /// a property rather than a comment.
    bytes: []const u8,
    owns_bytes: bool,
    /// At most two edits (`calcId` and `fullCalcOnLoad`) or one (the
    /// created element), so the list is inline rather than allocated:
    /// a fixed-size list is one fewer thing prepare can fail on.
    storage: [2]Edit = undefined,
    count: usize = 0,

    /// A method rather than a stored slice on purpose: a `[]const Edit`
    /// field pointing into `storage` would be left addressing the
    /// *previous* copy every time the struct is returned or assigned by
    /// value, which is a dangling slice that happens to read correctly
    /// until it does not.
    pub fn edits(self: *const Plan) []const Edit {
        return self.storage[0..self.count];
    }

    pub fn deinit(self: *Plan, gpa: Allocator) void {
        if (self.owns_bytes) gpa.free(self.bytes);
        self.* = undefined;
    }
};

pub const Result = union(enum) {
    ok: Plan,
    refused: calc.Refusal,
};

/// Build the edits that take `xml`'s calc state to `want`.
///
/// `state` must be the parse of `xml` — the spans it carries are offsets
/// into those exact bytes, and a `CalcState` from a different part would
/// address ranges that mean nothing here. Asserted, not documented.
pub fn plan(
    gpa: Allocator,
    xml: []const u8,
    state: calc.CalcState,
    want: Desired,
) error{OutOfMemory}!Result {
    if (state.present) {
        assert(state.spans.element.end <= xml.len);
        assert(state.spans.attrs.end <= state.spans.element.end);
    }

    var out: Plan = .{ .bytes = xml, .owns_bytes = false };
    var n: usize = 0;

    if (want.isEmpty()) return .{ .ok = out };

    if (!state.present) {
        // Nothing to amend: the element does not exist. Create exactly
        // as much of it as `want` asks for — "absent elements created
        // only when needed" is about the element, and the same sentence
        // decides which attributes it is born with.
        const at = state.spans.insert_at orelse
            return .{ .refused = .{ .reason = .malformed_calc_part } };
        assert(at <= xml.len);
        out.storage[0] = .{
            .at = .{ .start = at, .end = at },
            .replacement = createdElement(want),
            .kind = .calc_pr_create,
        };
        n = 1;
    } else {
        // Attribute order: `calcId` then `fullCalcOnLoad`, matching the
        // order §5.7.6 names them and the order Excel writes them. Two
        // insertions land on the same zero-width point, so the order they
        // are emitted in is the order they appear in the output — which
        // is why it is fixed here rather than left to a sort.
        const insert_at = attrInsertPoint(xml, state.spans.attrs);

        // Absence is not the target state. `calcId`'s schema default IS
        // zero, so a `<calcPr/>` with no `calcId` already *means* what a
        // recalc wants it to mean — but §5.7.6's assertion is that the
        // file **states** the pair, and the whole point of writing
        // `calcId="0"` is to tell a consumer something about the
        // producer. A default nobody wrote says nothing about who did.
        // A file that already states it is left alone, which is why the
        // condition is over the span and the value together and not over
        // either one.
        if (want.calc_id_zero and (state.spans.calc_id == null or state.calc_id != 0)) {
            if (state.spans.calc_id) |s| {
                out.storage[n] = .{ .at = s, .replacement = trimLeadingSpace(calc_id_attr), .kind = .calc_id_replace };
            } else {
                out.storage[n] = .{
                    .at = .{ .start = insert_at, .end = insert_at },
                    .replacement = calc_id_attr,
                    .kind = .calc_id_insert,
                };
            }
            n += 1;
        }
        if (want.full_calc_on_load and
            (state.spans.full_calc_on_load == null or !state.full_calc_on_load))
        {
            if (state.spans.full_calc_on_load) |s| {
                out.storage[n] = .{
                    .at = s,
                    .replacement = trimLeadingSpace(full_calc_attr),
                    .kind = .full_calc_on_load_replace,
                };
            } else {
                out.storage[n] = .{
                    .at = .{ .start = insert_at, .end = insert_at },
                    .replacement = full_calc_attr,
                    .kind = .full_calc_on_load_insert,
                };
            }
            n += 1;
        }
    }

    if (n == 0) {
        // The part already says what the run wants it to say. Zero edits
        // and the source bytes back — a second recalc of an unchanged
        // workbook rewrites nothing.
        return .{ .ok = out };
    }

    sortEdits(out.storage[0..n]);
    out.bytes = try applyEdits(gpa, xml, out.storage[0..n]);
    out.owns_bytes = true;
    out.count = n;
    return .{ .ok = out };
}

/// The element a workbook with no `<calcPr>` gets. Three constants and an
/// exhaustive choice, so a fourth desired state cannot be added without
/// deciding what it writes.
fn createdElement(want: Desired) []const u8 {
    if (want.calc_id_zero and want.full_calc_on_load) {
        return "<calcPr" ++ calc_id_attr ++ full_calc_attr ++ "/>";
    }
    if (want.calc_id_zero) return "<calcPr" ++ calc_id_attr ++ "/>";
    assert(want.full_calc_on_load);
    return "<calcPr" ++ full_calc_attr ++ "/>";
}

/// A replacement writes the attribute where one already is, so it must
/// not carry the separating space the insertion form needs.
fn trimLeadingSpace(s: []const u8) []const u8 {
    assert(s[0] == ' ');
    return s[1..];
}

/// Where a new attribute goes: after the last one, ahead of whatever
/// whitespace the producer left before `/>`.
///
/// `wdi_excel.xlsx` writes `<calcPr calcId="40001" />`, and that space is
/// a byte the round-trip owes back. Inserting at `attrs.end` would push
/// the new attribute past it and produce a double space; inserting at the
/// last non-space byte keeps the producer's spacing exactly where it was.
fn attrInsertPoint(source: []const u8, attrs: Span) u32 {
    var at = attrs.end;
    while (at > attrs.start and isSpace(source[at - 1])) at -= 1;
    return at;
}

fn isSpace(c: u8) bool {
    return c == ' ' or c == '\t' or c == '\n' or c == '\r';
}

/// Ascending by `(start, end)`. At most three entries and at most one
/// pair sharing a start, so an insertion sort is the whole algorithm —
/// and, unlike `std.sort`, it is stable, which is what keeps two
/// zero-width inserts in the order `plan` emitted them.
fn sortEdits(edits: []Edit) void {
    var i: usize = 1;
    while (i < edits.len) : (i += 1) {
        const e = edits[i];
        var j = i;
        while (j > 0 and (edits[j - 1].at.start > e.at.start or
            (edits[j - 1].at.start == e.at.start and edits[j - 1].at.end > e.at.end))) : (j -= 1)
        {
            edits[j] = edits[j - 1];
        }
        edits[j] = e;
    }
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

// ─── confinement ─────────────────────────────────────────────────

/// The one range an edit of this kind may address in this part.
///
/// An equality, not a bound — same reasoning as M5b1's: a patcher that
/// wrote a subrange would still be confined, but the kind would have
/// stopped describing the edit, and the kind is what the checker checks.
/// Null means the kind is unreachable for a part in this state, which is
/// itself a failure when an edit claims it.
pub fn approvedRange(source: []const u8, state: calc.CalcState, kind: EditKind) ?Span {
    return switch (kind) {
        .calc_pr_create => blk: {
            if (state.present) break :blk null;
            const at = state.spans.insert_at orelse break :blk null;
            break :blk .{ .start = at, .end = at };
        },
        .calc_id_replace => if (state.present) state.spans.calc_id else null,
        .full_calc_on_load_replace => if (state.present) state.spans.full_calc_on_load else null,
        .calc_id_insert => blk: {
            if (!state.present or state.spans.calc_id != null) break :blk null;
            const at = attrInsertPoint(source, state.spans.attrs);
            break :blk .{ .start = at, .end = at };
        },
        .full_calc_on_load_insert => blk: {
            if (!state.present or state.spans.full_calc_on_load != null) break :blk null;
            const at = attrInsertPoint(source, state.spans.attrs);
            break :blk .{ .start = at, .end = at };
        },
    };
}

pub const ConfinementError = error{
    /// Two edits address overlapping bytes, so which one the output shows
    /// is a function of the sort and not of the intent.
    OverlappingEdits,
    /// An edit's range is not the one its kind is allowed to address.
    UnapprovedRange,
    /// Replaying the edits over the source does not reproduce the output,
    /// so the list does not explain every changed byte.
    EditsDoNotExplainOutput,
    /// Two edits of the same kind. The calc state has one `calcId` and
    /// one `fullCalcOnLoad`; a list with two of either is describing a
    /// part this file did not read.
    DuplicateKind,
    OutOfMemory,
};

/// Prove a calc-state patch changed only what it says it changed.
///
/// The same two independent statements M5b1 settled on, because either
/// alone is satisfiable by a patcher that is wrong: every edit sits
/// exactly on the range its kind approves, AND replaying the list over
/// the source reproduces the output byte for byte.
pub fn verifyConfinement(
    gpa: Allocator,
    source: []const u8,
    out: []const u8,
    state: calc.CalcState,
    edits: []const Edit,
) ConfinementError!void {
    var prev_end: u32 = 0;
    var seen = std.EnumSet(EditKind).initEmpty();
    for (edits) |e| {
        if (e.at.start < prev_end) return error.OverlappingEdits;
        if (e.at.end < e.at.start or e.at.end > source.len) return error.UnapprovedRange;
        prev_end = e.at.end;

        if (seen.contains(e.kind)) return error.DuplicateKind;
        seen.insert(e.kind);

        const approved = approvedRange(source, state, e.kind) orelse return error.UnapprovedRange;
        if (approved.start != e.at.start or approved.end != e.at.end) return error.UnapprovedRange;
    }

    const replayed = try applyEdits(gpa, source, edits);
    defer gpa.free(replayed);
    if (!std.mem.eql(u8, replayed, out)) return error.EditsDoNotExplainOutput;
}

// ─── §5.7.7 mark-only eligibility (normative) ────────────────────

/// May `.keep_stale_and_mark` suppress this refusal?
///
/// **Exactly two planes may be suppressed**, and the switch is exhaustive
/// so a new plane cannot be added without answering the question. The
/// asymmetry is the point: mark-only says "these caches are stale, ask
/// Excel to redo them", which is an honest thing to say about a workbook
/// whose formulas this engine does not implement. It is a *lie* about a
/// workbook whose input was malformed, whose signature the mutation would
/// invalidate, whose embeddings would go stale under the change, that
/// exceeded a limit, that was cancelled, or that failed to read — in
/// every one of those the run does not know that a full recalculation
/// would even succeed, and a file marked `fullCalcOnLoad` is a file whose
/// caller believes the problem was handled.
///
/// `FormulaPrecisionAsDisplayed` is called out separately in §10 for the
/// same reason it is listed here as ineligible: `fullPrecision="0"` means
/// every stored value must be rounded to its number format before it is
/// read, so the *existing* caches are not the ones Excel would produce
/// either. Marking would not make them right.
pub fn markEligible(plane: PlaneTwo) bool {
    return switch (plane) {
        .FormulaUnsupportedFunction,
        .FormulaUnsupportedConstruct,
        => true,

        .FormulaPrecisionAsDisplayed,
        .FormulaMalformedInput,
        .FormulaLocaleSensitiveInput,
        .FormulaDataTableUnsupported,
        .FormulaSignedWorkbook,
        .FormulaStaleEmbeddings,
        .FormulaAnchorRequired,
        .FormulaCycle,
        .FormulaDynamicRefUnstable,
        .FormulaSpillPersistUnsupported,
        .FormulaResultNotRepresentable,
        .FormulaLimitExceeded,
        => false,
    };
}

/// The planes `.keep_stale_and_mark` may suppress, as data.
///
/// Derived from `markEligible` at comptime rather than written twice: a
/// second list is a second thing to forget to update, and the test that
/// compares them would only ever fail because someone updated one of
/// them.
pub const mark_eligible_planes = blk: {
    const all = @typeInfo(PlaneTwo).@"enum".fields;
    var n: usize = 0;
    for (all) |f| {
        if (markEligible(@field(PlaneTwo, f.name))) n += 1;
    }
    var out: [n]PlaneTwo = undefined;
    var i: usize = 0;
    for (all) |f| {
        const p = @field(PlaneTwo, f.name);
        if (markEligible(p)) {
            out[i] = p;
            i += 1;
        }
    }
    const frozen = out;
    break :blk frozen;
};

// ─── tests ───────────────────────────────────────────────────────

const testing = std.testing;

const wb_ns = " xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"";

fn parse(gpa: Allocator, xml: []const u8) !calc.CalcState {
    return switch (try calc.parseCalcState(gpa, xml)) {
        .ok => |s| s,
        .refused => |r| {
            std.debug.print("unexpected refusal: {t}\n", .{r.reason});
            return error.UnexpectedRefusal;
        },
    };
}

/// Plan, verify confinement, and hand back the patched bytes. Every
/// fixture below goes through this rather than calling `plan` directly,
/// so no test can assert an output without also asserting that the edits
/// explain it.
fn planned(gpa: Allocator, xml: []const u8, want: Desired) ![]u8 {
    var state = try parse(gpa, xml);
    defer state.deinit(gpa);
    var p = switch (try plan(gpa, xml, state, want)) {
        .ok => |v| v,
        .refused => return error.UnexpectedRefusal,
    };
    defer p.deinit(gpa);
    try verifyConfinement(gpa, xml, p.bytes, state, p.edits());
    return gpa.dupe(u8, p.bytes);
}

test "calcPr: absent element is created at the sequence position" {
    const gpa = testing.allocator;
    const xml =
        "<workbook" ++ wb_ns ++ "><sheets><sheet name=\"S\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>";
    const out = try planned(gpa, xml, Desired.after_recalc);
    defer gpa.free(out);
    try testing.expectEqualStrings(
        "<workbook" ++ wb_ns ++ "><sheets><sheet name=\"S\" sheetId=\"1\" r:id=\"rId1\"/></sheets>" ++
            "<calcPr calcId=\"0\" fullCalcOnLoad=\"1\"/></workbook>",
        out,
    );
}

test "calcPr: creation goes BEFORE a successor element, not at the end" {
    const gpa = testing.allocator;
    // `<extLst>` follows `calcPr` in `CT_Workbook`'s sequence. Appending
    // before `</workbook>` would put `<calcPr>` after it, which is an
    // invalid document — and one that opens as a repair prompt rather
    // than as an error, so nothing but a fixture would catch it.
    const xml = "<workbook" ++ wb_ns ++ "><sheets/><extLst><ext uri=\"{X}\"/></extLst></workbook>";
    const out = try planned(gpa, xml, Desired.after_recalc);
    defer gpa.free(out);
    try testing.expectEqualStrings(
        "<workbook" ++ wb_ns ++ "><sheets/><calcPr calcId=\"0\" fullCalcOnLoad=\"1\"/>" ++
            "<extLst><ext uri=\"{X}\"/></extLst></workbook>",
        out,
    );
}

test "calcPr: every successor in the table moves the insertion point" {
    const gpa = testing.allocator;
    for (calc.calc_pr_successors) |succ| {
        const xml = try std.fmt.allocPrint(
            gpa,
            "<workbook" ++ wb_ns ++ "><sheets/><{s}/></workbook>",
            .{succ},
        );
        defer gpa.free(xml);
        const out = try planned(gpa, xml, Desired.after_recalc);
        defer gpa.free(out);
        const expect = try std.fmt.allocPrint(
            gpa,
            "<workbook" ++ wb_ns ++ "><sheets/><calcPr calcId=\"0\" fullCalcOnLoad=\"1\"/><{s}/></workbook>",
            .{succ},
        );
        defer gpa.free(expect);
        try testing.expectEqualStrings(expect, out);
    }
}

test "calcPr: the created element carries only what was asked for" {
    const gpa = testing.allocator;
    const xml = "<workbook" ++ wb_ns ++ "><sheets/></workbook>";
    const marked = try planned(gpa, xml, Desired.mark_only);
    defer gpa.free(marked);
    try testing.expectEqualStrings(
        "<workbook" ++ wb_ns ++ "><sheets/><calcPr fullCalcOnLoad=\"1\"/></workbook>",
        marked,
    );
}

test "calcPr: both attributes present are replaced in place" {
    const gpa = testing.allocator;
    const xml = "<workbook" ++ wb_ns ++ "><calcPr calcId=\"191029\" fullCalcOnLoad=\"0\" calcMode=\"manual\"/></workbook>";
    const out = try planned(gpa, xml, Desired.after_recalc);
    defer gpa.free(out);
    try testing.expectEqualStrings(
        "<workbook" ++ wb_ns ++ "><calcPr calcId=\"0\" fullCalcOnLoad=\"1\" calcMode=\"manual\"/></workbook>",
        out,
    );
}

test "calcPr: the trailing space Excel writes survives an insertion" {
    const gpa = testing.allocator;
    const xml = "<workbook" ++ wb_ns ++ "><calcPr calcId=\"40001\" /></workbook>";
    const out = try planned(gpa, xml, Desired.after_recalc);
    defer gpa.free(out);
    try testing.expectEqualStrings(
        "<workbook" ++ wb_ns ++ "><calcPr calcId=\"0\" fullCalcOnLoad=\"1\" /></workbook>",
        out,
    );
}

test "calcPr: an attribute-less element gets both, separated correctly" {
    const gpa = testing.allocator;
    const xml = "<workbook" ++ wb_ns ++ "><calcPr/></workbook>";
    const out = try planned(gpa, xml, Desired.after_recalc);
    defer gpa.free(out);
    try testing.expectEqualStrings(
        "<workbook" ++ wb_ns ++ "><calcPr calcId=\"0\" fullCalcOnLoad=\"1\"/></workbook>",
        out,
    );
}

test "calcPr: a non-self-closing element is amended, not reshaped" {
    const gpa = testing.allocator;
    const xml = "<workbook" ++ wb_ns ++ "><calcPr calcId=\"7\"></calcPr></workbook>";
    const out = try planned(gpa, xml, Desired.after_recalc);
    defer gpa.free(out);
    try testing.expectEqualStrings(
        "<workbook" ++ wb_ns ++ "><calcPr calcId=\"0\" fullCalcOnLoad=\"1\"></calcPr></workbook>",
        out,
    );
}

test "calcPr: a state already at the target produces zero edits" {
    const gpa = testing.allocator;
    const xml = "<workbook" ++ wb_ns ++ "><calcPr calcId=\"0\" fullCalcOnLoad=\"1\"/></workbook>";
    var state = try parse(gpa, xml);
    defer state.deinit(gpa);
    var p = switch (try plan(gpa, xml, state, Desired.after_recalc)) {
        .ok => |v| v,
        .refused => return error.UnexpectedRefusal,
    };
    defer p.deinit(gpa);
    try testing.expectEqual(@as(usize, 0), p.edits().len);
    try testing.expect(!p.owns_bytes);
    try testing.expectEqual(xml.ptr, p.bytes.ptr);
}

test "calcPr: mark-only differs from the source in exactly fullCalcOnLoad" {
    const gpa = testing.allocator;
    // The whole of §5.7.7's byte-identity claim, stated as a diff over
    // the two documents rather than as a count of edits: the changed
    // window is the inserted attribute and nothing else.
    const xml = "<workbook" ++ wb_ns ++ "><calcPr calcId=\"191029\" calcMode=\"manual\" iterate=\"1\" iterateCount=\"77\"/></workbook>";
    const out = try planned(gpa, xml, Desired.mark_only);
    defer gpa.free(out);

    const w = changedWindow(xml, out).?;
    // Nothing in the source changed: the window is zero-width in `xml`.
    try testing.expectEqual(w.start, w.end);
    try testing.expectEqualStrings(
        " fullCalcOnLoad=\"1\"",
        out[w.start .. w.start + " fullCalcOnLoad=\"1\"".len],
    );
    // calcId, calcMode and the iteration pair are all still there.
    try testing.expect(std.mem.indexOf(u8, out, "calcId=\"191029\"") != null);
    try testing.expect(std.mem.indexOf(u8, out, "calcMode=\"manual\"") != null);
    try testing.expect(std.mem.indexOf(u8, out, "iterate=\"1\"") != null);
    try testing.expect(std.mem.indexOf(u8, out, "iterateCount=\"77\"") != null);
}

test "calcPr: mark-only on an already-marked workbook changes nothing" {
    const gpa = testing.allocator;
    const xml = "<workbook" ++ wb_ns ++ "><calcPr calcId=\"191029\" fullCalcOnLoad=\"true\"/></workbook>";
    const out = try planned(gpa, xml, Desired.mark_only);
    defer gpa.free(out);
    try testing.expectEqualStrings(xml, out);
    try testing.expect(changedWindow(xml, out) == null);
}

test "calcPr: an empty desire never touches the part" {
    const gpa = testing.allocator;
    const xml = "<workbook" ++ wb_ns ++ "><calcPr calcId=\"5\"/></workbook>";
    const out = try planned(gpa, xml, .{});
    defer gpa.free(out);
    try testing.expectEqualStrings(xml, out);
}

test "calcPr: every preserved row survives a recalc write" {
    const gpa = testing.allocator;
    // §5.7.6's preservation list, one attribute per row, all on one
    // element. The assertion is that each appears in the output with the
    // value it had — which for every one of them means the patcher never
    // addressed its bytes.
    const preserved = [_][]const u8{
        "calcMode=\"autoNoTable\"",
        "refMode=\"R1C1\"",
        "iterate=\"1\"",
        "iterateCount=\"250\"",
        "iterateDelta=\"1E-4\"",
        "calcCompleted=\"0\"",
        "calcOnSave=\"0\"",
        "concurrentCalc=\"0\"",
        "concurrentManualCount=\"4\"",
        "forceFullCalc=\"1\"",
        "fullPrecision=\"1\"",
        "vendorThing=\"kept\"",
    };
    var attrs: std.ArrayListUnmanaged(u8) = .empty;
    defer attrs.deinit(gpa);
    for (preserved) |a| {
        try attrs.append(gpa, ' ');
        try attrs.appendSlice(gpa, a);
    }
    const xml = try std.fmt.allocPrint(
        gpa,
        "<workbook" ++ wb_ns ++ "><calcPr calcId=\"191029\"{s}/></workbook>",
        .{attrs.items},
    );
    defer gpa.free(xml);

    const out = try planned(gpa, xml, Desired.after_recalc);
    defer gpa.free(out);
    for (preserved) |a| {
        try testing.expect(std.mem.indexOf(u8, out, a) != null);
    }
    try testing.expect(std.mem.indexOf(u8, out, "calcId=\"0\"") != null);
    try testing.expect(std.mem.indexOf(u8, out, "calcId=\"191029\"") == null);
    try testing.expect(std.mem.indexOf(u8, out, "fullCalcOnLoad=\"1\"") != null);
}

test "calcPr: the calcFeatures extension and workbookPr are untouched" {
    const gpa = testing.allocator;
    const ext = "<extLst><ext uri=\"" ++ calc.calc_features_ext_uri ++
        "\" xmlns:xcalcf=\"x\"><xcalcf:calcFeatures><xcalcf:feature name=\"microsoft.com:RD\"/>" ++
        "</xcalcf:calcFeatures></ext></extLst>";
    const xml = "<workbook" ++ wb_ns ++ "><workbookPr date1904=\"1\"/><calcPr calcId=\"9\"/>" ++ ext ++ "</workbook>";
    const out = try planned(gpa, xml, Desired.after_recalc);
    defer gpa.free(out);
    try testing.expect(std.mem.indexOf(u8, out, ext) != null);
    try testing.expect(std.mem.indexOf(u8, out, "<workbookPr date1904=\"1\"/>") != null);

    // And the round-trip still reads the same everything-else.
    var after = try parse(gpa, out);
    defer after.deinit(gpa);
    try testing.expectEqual(calc.DateSystem.d1904, after.date_system);
    try testing.expectEqual(@as(usize, 1), after.calc_features.len);
    try testing.expectEqual(@as(u32, 0), after.calc_id);
    try testing.expect(after.full_calc_on_load);
}

test "calcPr: fullPrecision=\"0\" refuses before any write exists" {
    const gpa = testing.allocator;
    const xml = "<workbook" ++ wb_ns ++ "><calcPr calcId=\"1\" fullPrecision=\"0\"/></workbook>";
    switch (try calc.parseCalcState(gpa, xml)) {
        .ok => return error.ExpectedRefusal,
        .refused => |r| {
            try testing.expectEqual(calc.Refusal.Reason.precision_as_displayed, r.reason);
            try testing.expectEqual(PlaneTwo.FormulaPrecisionAsDisplayed, r.planeTwo());
        },
    }
}

test "calcPr: sheetCalcPr is preserved — this file never sees a worksheet" {
    const sheet = "<worksheet" ++ wb_ns ++ "><sheetCalcPr fullCalcOnLoad=\"1\"/><sheetData/></worksheet>";
    const r = calc.parseSheetCalcPr(sheet);
    try testing.expect(r.ok.present);
    try testing.expect(r.ok.full_calc_on_load);
}

test "calcPr: a part with no insertion point refuses rather than guess" {
    const gpa = testing.allocator;
    // No `</workbook>` close and no successor element: `insert_at` is
    // null, and writing at offset 0 would put `<calcPr>` ahead of the
    // root. The honest answer is a refusal.
    const xml = "<sheets/>";
    var state = try parse(gpa, xml);
    defer state.deinit(gpa);
    try testing.expect(state.spans.insert_at == null);
    switch (try plan(gpa, xml, state, Desired.after_recalc)) {
        .ok => return error.ExpectedRefusal,
        .refused => |r| try testing.expectEqual(calc.Refusal.Reason.malformed_calc_part, r.reason),
    }
}

test "confinement: the checker rejects an edit outside its approved range" {
    const gpa = testing.allocator;
    const xml = "<workbook" ++ wb_ns ++ "><calcPr calcId=\"191029\"/></workbook>";
    var state = try parse(gpa, xml);
    defer state.deinit(gpa);
    var p = switch (try plan(gpa, xml, state, Desired.after_recalc)) {
        .ok => |v| v,
        .refused => return error.UnexpectedRefusal,
    };
    defer p.deinit(gpa);
    try verifyConfinement(gpa, xml, p.bytes, state, p.edits());

    // Slide the replace one byte to the left. The output is unchanged,
    // so only the range check can catch it.
    var tampered: [2]Edit = undefined;
    @memcpy(tampered[0..p.edits().len], p.edits());
    tampered[0].at.start -= 1;
    try testing.expectError(
        error.UnapprovedRange,
        verifyConfinement(gpa, xml, p.bytes, state, tampered[0..p.edits().len]),
    );
}

test "confinement: the checker rejects a write nobody reported" {
    const gpa = testing.allocator;
    const xml = "<workbook" ++ wb_ns ++ "><calcPr calcId=\"191029\"/></workbook>";
    var state = try parse(gpa, xml);
    defer state.deinit(gpa);
    var p = switch (try plan(gpa, xml, state, Desired.after_recalc)) {
        .ok => |v| v,
        .refused => return error.UnexpectedRefusal,
    };
    defer p.deinit(gpa);

    // An output with one extra byte the edit list does not account for.
    const smuggled = try std.mem.concat(gpa, u8, &.{ p.bytes, "<!--x-->" });
    defer gpa.free(smuggled);
    try testing.expectError(
        error.EditsDoNotExplainOutput,
        verifyConfinement(gpa, xml, smuggled, state, p.edits()),
    );
}

test "confinement: the checker rejects two edits of one kind" {
    const gpa = testing.allocator;
    const xml = "<workbook" ++ wb_ns ++ "><calcPr calcId=\"191029\"/></workbook>";
    var state = try parse(gpa, xml);
    defer state.deinit(gpa);
    const s = state.spans.calc_id.?;
    const doubled = [_]Edit{
        .{ .at = s, .replacement = "calcId=\"0\"", .kind = .calc_id_replace },
        .{ .at = .{ .start = s.end, .end = s.end }, .replacement = "", .kind = .calc_id_replace },
    };
    try testing.expectError(
        error.DuplicateKind,
        verifyConfinement(gpa, xml, xml, state, &doubled),
    );
}

test "confinement: an insert kind is unreachable when the attribute exists" {
    const gpa = testing.allocator;
    const xml = "<workbook" ++ wb_ns ++ "><calcPr calcId=\"191029\"/></workbook>";
    var state = try parse(gpa, xml);
    defer state.deinit(gpa);
    try testing.expect(approvedRange(xml, state, .calc_id_insert) == null);
    try testing.expect(approvedRange(xml, state, .calc_id_replace) != null);
    try testing.expect(approvedRange(xml, state, .calc_pr_create) == null);
    try testing.expect(approvedRange(xml, state, .full_calc_on_load_insert) != null);
    try testing.expect(approvedRange(xml, state, .full_calc_on_load_replace) == null);
}

test "mark eligibility: exactly two planes, and the table agrees" {
    try testing.expect(markEligible(.FormulaUnsupportedFunction));
    try testing.expect(markEligible(.FormulaUnsupportedConstruct));
    try testing.expectEqual(@as(usize, 2), mark_eligible_planes.len);

    // Every plane the §10 taxonomy names that mark-only must NOT
    // suppress, spelled out rather than derived — this list is the one
    // §5.7.7 is normative about, and a derivation from `markEligible`
    // would agree with a wrong `markEligible`.
    const always_refuse = [_]PlaneTwo{
        .FormulaPrecisionAsDisplayed,
        .FormulaMalformedInput,
        .FormulaLocaleSensitiveInput,
        .FormulaDataTableUnsupported,
        .FormulaSignedWorkbook,
        .FormulaStaleEmbeddings,
        .FormulaAnchorRequired,
        .FormulaCycle,
        .FormulaDynamicRefUnstable,
        .FormulaSpillPersistUnsupported,
        .FormulaResultNotRepresentable,
        .FormulaLimitExceeded,
    };
    for (always_refuse) |p| try testing.expect(!markEligible(p));
    try testing.expectEqual(
        @typeInfo(PlaneTwo).@"enum".fields.len,
        always_refuse.len + mark_eligible_planes.len,
    );
}

test "spans: the parser's coordinates address the element it read" {
    const gpa = testing.allocator;
    const xml = "<workbook" ++ wb_ns ++ "><calcPr calcId=\"40001\" iterate=\"1\"/></workbook>";
    var state = try parse(gpa, xml);
    defer state.deinit(gpa);
    try testing.expectEqualStrings(
        "<calcPr calcId=\"40001\" iterate=\"1\"/>",
        state.spans.element.slice(xml),
    );
    try testing.expectEqualStrings(
        " calcId=\"40001\" iterate=\"1\"",
        state.spans.attrs.slice(xml),
    );
    try testing.expectEqualStrings("calcId=\"40001\"", state.spans.calc_id.?.slice(xml));
    try testing.expect(state.spans.full_calc_on_load == null);
}

test "spans: a non-self-closing element's span reaches its close tag" {
    const gpa = testing.allocator;
    const xml = "<workbook" ++ wb_ns ++ "><calcPr calcId=\"1\"></calcPr></workbook>";
    var state = try parse(gpa, xml);
    defer state.deinit(gpa);
    try testing.expectEqualStrings(
        "<calcPr calcId=\"1\"></calcPr>",
        state.spans.element.slice(xml),
    );
}

test "round trip: every calc-state row reads back as itself after a write" {
    const gpa = testing.allocator;
    const Row = struct { attrs: []const u8, check: *const fn (calc.CalcState) bool };
    const rows = [_]Row{
        .{ .attrs = "calcMode=\"manual\"", .check = struct {
            fn f(s: calc.CalcState) bool {
                return s.calc_mode == .manual;
            }
        }.f },
        .{ .attrs = "calcMode=\"autoNoTable\"", .check = struct {
            fn f(s: calc.CalcState) bool {
                return s.calc_mode == .auto_no_table;
            }
        }.f },
        .{ .attrs = "refMode=\"R1C1\"", .check = struct {
            fn f(s: calc.CalcState) bool {
                return s.ref_mode == .r1c1;
            }
        }.f },
        .{ .attrs = "iterate=\"1\" iterateCount=\"33\" iterateDelta=\"0.25\"", .check = struct {
            fn f(s: calc.CalcState) bool {
                return s.iterate and s.iterate_count == 33 and s.iterate_delta == 0.25;
            }
        }.f },
        .{ .attrs = "calcCompleted=\"0\"", .check = struct {
            fn f(s: calc.CalcState) bool {
                return !s.calc_completed;
            }
        }.f },
        .{ .attrs = "calcOnSave=\"0\"", .check = struct {
            fn f(s: calc.CalcState) bool {
                return !s.calc_on_save;
            }
        }.f },
        .{ .attrs = "forceFullCalc=\"1\"", .check = struct {
            fn f(s: calc.CalcState) bool {
                return s.force_full_calc;
            }
        }.f },
        .{ .attrs = "concurrentCalc=\"0\" concurrentManualCount=\"3\"", .check = struct {
            fn f(s: calc.CalcState) bool {
                return !s.concurrent_calc and s.concurrent_manual_count == 3;
            }
        }.f },
        .{ .attrs = "fullPrecision=\"1\"", .check = struct {
            fn f(s: calc.CalcState) bool {
                return s.full_precision;
            }
        }.f },
    };

    for (rows) |row| {
        const xml = try std.fmt.allocPrint(
            gpa,
            "<workbook" ++ wb_ns ++ "><calcPr calcId=\"191029\" {s}/></workbook>",
            .{row.attrs},
        );
        defer gpa.free(xml);
        const out = try planned(gpa, xml, Desired.after_recalc);
        defer gpa.free(out);

        var after = try parse(gpa, out);
        defer after.deinit(gpa);
        try testing.expect(row.check(after));
        // The pair, at the byte level and through the parse.
        try testing.expect(std.mem.indexOf(u8, out, "calcId=\"0\"") != null);
        try testing.expect(std.mem.indexOf(u8, out, "fullCalcOnLoad=\"1\"") != null);
        try testing.expectEqual(@as(u32, 0), after.calc_id);
        try testing.expect(after.full_calc_on_load);
    }
}

test "round trip: a second write over a written part is a no-op" {
    const gpa = testing.allocator;
    const xml = "<workbook" ++ wb_ns ++ "><calcPr calcId=\"191029\" calcMode=\"manual\"/></workbook>";
    const once = try planned(gpa, xml, Desired.after_recalc);
    defer gpa.free(once);
    const twice = try planned(gpa, once, Desired.after_recalc);
    defer gpa.free(twice);
    try testing.expectEqualStrings(once, twice);
}

test "plan: an out-of-memory patch leaves nothing half-written" {
    const gpa = testing.allocator;
    const xml = "<workbook" ++ wb_ns ++ "><calcPr calcId=\"191029\"/></workbook>";
    var state = try parse(gpa, xml);
    defer state.deinit(gpa);
    try testing.checkAllAllocationFailures(gpa, struct {
        fn run(a: Allocator, source: []const u8, s: calc.CalcState) !void {
            var p = switch (try plan(a, source, s, Desired.after_recalc)) {
                .ok => |v| v,
                .refused => return error.UnexpectedRefusal,
            };
            defer p.deinit(a);
            try verifyConfinement(a, source, p.bytes, s, p.edits());
        }
    }.run, .{ xml, state });
}

test "calcPr: an element the document never closes refuses" {
    const gpa = testing.allocator;
    // The scanner is content with a truncated document, so nothing below
    // it would notice; the span the patcher would then address has a
    // start and no end.
    const xml = "<workbook" ++ wb_ns ++ "><calcPr calcId=\"1\">";
    switch (try calc.parseCalcState(gpa, xml)) {
        .ok => return error.ExpectedRefusal,
        .refused => |r| try testing.expectEqual(calc.Refusal.Reason.malformed_calc_part, r.reason),
    }
}
