//! §5.7.3 step 4's prepare/swap transaction — the in-memory half.
//!
//! M5b2 of the tier-D1 ladder (`goal_formula.md`). Everything a recalc
//! changes about a workbook stages into **one candidate**, and the swap
//! that installs it is the last operation and cannot fail. `recalculate()`
//! and `saveWithRecalc` (M5d2) are the entry points that will call this;
//! this file is the transaction they call, not the API they expose.
//!
//! What "the swap cannot fail" costs
//! --------------------------------
//! Everything the post-swap workbook needs has to exist before the swap
//! begins — not just the new part bytes, but the new typed views over
//! them, the report, and the list slot the retained generation goes into.
//! `prepare` therefore parses `xl/workbook.xml` and every sheet view the
//! workbook currently has materialised, builds the report in full, and
//! reserves capacity in `Workbook.retained`. `swap` is then a sequence of
//! moves: no allocation, no parse, no `try`.
//!
//! The gate is a proof, not an assertion. `checkAllAllocationFailures`
//! over `prepare` shows every failure landing before the swap with the
//! workbook unchanged, and a separate fixture runs the swap itself under
//! an allocator that fails every request — which passes only because the
//! swap asks it for nothing.
//!
//! What the swap must NOT free (§5.7.4)
//! -----------------------------------
//! `Workbook.sheet()` returns interior pointers and `cellByRef()` promises
//! Workbook-lifetime validity for the strings it hands back. Those strings
//! borrow **part bytes**. So the superseded generation — the whole
//! `PartStore`, plus the typed views over it — is retained until
//! `Workbook.deinit`, and the worksheet slot array is never reallocated,
//! so `&wb.worksheets[i]` survives any number of nonstructural recalcs.
//! Retention is counted (`max_retained_generations`, default 4, plus byte
//! accounting) and the count refuses **before** the swap; there is no
//! eviction, because evicting would mean deciding that a borrow the
//! caller was promised is now stale.
//!
//! M5b0 is what makes that affordable: every generation shares one
//! ref-counted `SourceBacking`, so a four-deep retained set costs one file
//! descriptor, not four.
//!
//! Not here
//! --------
//! The file transaction — temp file, `File.sync`, rename, directory fsync
//! — is M5d1/M5d2's, and its ordering (§5.7.9) puts the swap *after* the
//! rename. That is why `Candidate` is a value the caller holds rather than
//! something `prepare` installs: `recalculate()` swaps immediately,
//! `saveWithRecalc` serialises the candidate first and swaps only once the
//! rename has committed. The report already carries the dormant durability
//! slot that path needs, preallocated, because a post-rename fsync failure
//! is discovered after all allocation must be done.

const std = @import("std");
const assert = std.debug.assert;
const Allocator = std.mem.Allocator;

const store_mod = @import("store.zig");
// M5d1: the commit outcome §5.7.9 hands to the dormant durability slot.
const atomic_file = @import("atomic_file.zig");
const workbook_mod = @import("workbook.zig");
const typed_parts = @import("typed_parts/root.zig");
const engine = @import("zlsx_formula");

const PartStore = store_mod.PartStore;
const Workbook = workbook_mod.Workbook;
const RetainedGeneration = workbook_mod.RetainedGeneration;
const workbook_xml_mod = typed_parts.workbook_xml;
const sheet_xml_mod = typed_parts.sheet_xml;

const calc = engine.calc;
const calc_patch = engine.calc_patch;
const run_inputs = engine.run_inputs;

pub const PlaneTwo = engine.decode.PlaneTwo;
pub const CancelToken = run_inputs.CancelToken;

/// The relationship type that names a calculation chain. Resolution of
/// its `Target` is owner-relative to `xl/workbook.xml` — the norm real
/// files follow (`Target="calcChain.xml"`), not the absolute form
/// `PartStore.removePart`'s string matcher looks for.
pub const calc_chain_rel_type =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain";

pub const workbook_part = "xl/workbook.xml";
pub const workbook_rels_part = "xl/_rels/workbook.xml.rels";
pub const content_types_part = "[Content_Types].xml";

// ─── options ─────────────────────────────────────────────────────

/// §5.7.7's caller policy for constructs this engine does not implement.
pub const OnUnsupported = enum {
    /// The default: the whole recalc refuses, and nothing is mutated.
    refuse,
    /// Keep the workbook's existing caches and set `fullCalcOnLoad="1"`
    /// so the next consumer that *can* calculate them does. Eligible only
    /// for the two planes `calc_patch.markEligible` names.
    keep_stale_and_mark,
};

/// §5.7.4's default retention depth. Four rather than one because a
/// caller that recalculates in a loop should not have its earlier borrows
/// invalidated by the second iteration, and not unbounded because the
/// bytes are never reclaimed short of `deinit`.
pub const default_max_retained_generations: usize = 4;

/// §9's shape applied to retention: a ceiling on what the retained set may
/// hold resident. The default is generous — retention is the price of a
/// promise, and a run that trips this has a workbook large enough that the
/// caller wants to know.
pub const default_max_retained_bytes: u64 = 2 << 30;

/// §5.7.8's bound on the refusal census. Beyond this the report says it
/// truncated rather than growing without limit.
pub const max_census_entries: usize = 256;

pub const Options = struct {
    max_retained_generations: usize = default_max_retained_generations,
    max_retained_bytes: u64 = default_max_retained_bytes,
    on_unsupported: OnUnsupported = .refuse,
    /// Polled at the start of prepare and again immediately before the
    /// candidate is handed back. A cancelled prepare leaves memory
    /// untouched — the candidate is torn down on the way out.
    cancel: ?CancelToken = null,
    /// §5.7.8's echo. Carried verbatim into the report so a reader can
    /// see what the run was actually given.
    resolved: ?run_inputs.EffectiveRunInputs = null,
};

/// One entry of §5.7.7's census: a construct the evaluator could not
/// implement, and where it was.
pub const Unsupported = struct {
    plane: PlaneTwo,
    sheet: u32 = 0,
    /// One-based, as OOXML writes it. Zero when the refusal is not about
    /// a cell.
    row: u32 = 0,
    /// Zero-based.
    col: u32 = 0,
};

/// A part the evaluation phase produced new bytes for — M5b1's patch
/// output, addressed by part name. The transaction copies them into the
/// candidate, so the caller's buffers may die at the call boundary.
pub const StagedPart = struct {
    name: []const u8,
    bytes: []const u8,
};

// ─── refusals ────────────────────────────────────────────────────

pub const Refusal = struct {
    reason: Reason,
    /// Set when the refusal is a census entry passed through — the plane
    /// the *evaluator* raised, not one this file invented.
    plane: ?PlaneTwo = null,
    /// §5.7.7's census at the moment of the refusal (M9a2): the entries
    /// whose constructs refused, bounded by `max_census_entries` and
    /// owned — `deinit` with the workbook's allocator. Empty for the
    /// retention and calc-state reasons, which are not about cells.
    /// This is what makes a refusal's diagnostics honest: before M9a2
    /// the pipeline collapsed the census into `plane` alone and the
    /// refusing cells never crossed any boundary (decision M9a1-4).
    census: []const Unsupported = &.{},
    census_truncated: bool = false,

    pub fn deinit(self: *Refusal, gpa: Allocator) void {
        gpa.free(self.census);
        self.* = undefined;
    }

    pub const Reason = enum {
        /// §5.7.4: the retained set is full. Refused before the swap,
        /// because the alternative is freeing a generation a caller was
        /// promised would outlive the run.
        retained_generations_exhausted,
        /// §5.7.4's byte accounting reached its ceiling.
        retained_bytes_exhausted,
        /// The census is non-empty and the caller asked to refuse, or
        /// asked to mark and at least one entry is not mark-eligible.
        unsupported_construct,
        /// `xl/workbook.xml`'s calc state did not parse, or said
        /// `fullPrecision="0"`.
        calc_state,
    };

    /// §10's plane-2 vocabulary. Exhaustive, so a new reason cannot be
    /// added without deciding which error a caller sees.
    pub fn planeTwo(self: Refusal) PlaneTwo {
        return switch (self.reason) {
            .retained_generations_exhausted,
            .retained_bytes_exhausted,
            => .FormulaLimitExceeded,
            .unsupported_construct, .calc_state => self.plane orelse .FormulaMalformedInput,
        };
    }

    /// The same refusal, as the error a `Workbook` method returns.
    ///
    /// One name per plane, exhaustively — the census a refusal carries
    /// comes from the *evaluator*, so any of §10's fourteen can arrive
    /// here, and a switch that collapsed the ones this file cannot itself
    /// raise would report a cycle as malformed input.
    pub fn toWorkbookError(self: Refusal) workbook_mod.Error {
        return switch (self.planeTwo()) {
            .FormulaUnsupportedFunction => error.FormulaUnsupportedFunction,
            .FormulaUnsupportedConstruct => error.FormulaUnsupportedConstruct,
            .FormulaPrecisionAsDisplayed => error.FormulaPrecisionAsDisplayed,
            .FormulaMalformedInput => error.FormulaMalformedInput,
            .FormulaLocaleSensitiveInput => error.FormulaLocaleSensitiveInput,
            .FormulaDataTableUnsupported => error.FormulaDataTableUnsupported,
            .FormulaSignedWorkbook => error.FormulaSignedWorkbook,
            .FormulaStaleEmbeddings => error.FormulaStaleEmbeddings,
            .FormulaAnchorRequired => error.FormulaAnchorRequired,
            .FormulaCycle => error.FormulaCycle,
            .FormulaDynamicRefUnstable => error.FormulaDynamicRefUnstable,
            .FormulaSpillPersistUnsupported => error.FormulaSpillPersistUnsupported,
            .FormulaResultNotRepresentable => error.FormulaResultNotRepresentable,
            .FormulaLimitExceeded => error.FormulaLimitExceeded,
        };
    }
};

/// The transaction speaks `Workbook`'s error set plus cancellation.
///
/// Widening rather than mapping: this file sits *above* `Workbook` and
/// calls into it, and a translation table would have to invent an answer
/// for every error it does not expect — which is how a
/// `SheetIndexOutOfRange` ends up reported as a missing relationship.
pub const Error = error{
    /// §5.5's cooperative cancellation. A cancelled prepare has torn the
    /// candidate down before returning; the workbook is untouched.
    Cancelled,
} || workbook_mod.Error;

pub const Result = union(enum) {
    ok: Candidate,
    refused: Refusal,
};

// ─── the report (§5.7.8) ─────────────────────────────────────────

/// The dormant durability slot §5.7.9 needs.
///
/// A fixed flag and a fixed errno, preallocated with the report and
/// flipped without allocation. The outcome it records — a directory fsync
/// that failed *after* the rename committed — is discovered at the one
/// point in the pipeline where allocation is no longer permitted, and a
/// slot that had to be created then would be a failure that could
/// contradict an already-successful commit.
pub const Durability = struct {
    warning: bool = false,
    err_code: i32 = 0,

    /// Infallible by construction: two scalar stores.
    pub fn warn(self: *Durability, code: i32) void {
        self.warning = true;
        self.err_code = code;
    }
};

pub const Report = struct {
    /// Sheet parts the candidate rewrote.
    sheets_patched: u32 = 0,
    /// Cells whose cached value the run wrote. Supplied by the caller —
    /// the transaction stages bytes, and only the projection that produced
    /// them knows how many publications they carry.
    cells_written: u32 = 0,
    /// §5.6c: iteration passes, summed across components.
    passes: u32 = 0,
    /// §5.6c: members that reached the semantic bound without settling.
    non_converged_cells: u32 = 0,
    /// §5.6e: how many times the outer loop rebuilt the graph.
    dynamic_passes: u32 = 1,
    /// §5.7.7's census, bounded by `max_census_entries`.
    census: []const Unsupported = &.{},
    census_truncated: bool = false,
    /// §5.5's fingerprintable projection of the run's inputs. Null when
    /// the caller did not supply one (the standalone `markRecalcOnLoad`
    /// path has no run).
    resolved: ?run_inputs.EffectiveRunInputs = null,
    /// True when the run took §5.7.7's mark-only path: the workbook's
    /// existing caches were kept and only `fullCalcOnLoad="1"` was
    /// written. A successful recalc sets `fullCalcOnLoad` too, so a flag
    /// named after the attribute would say nothing; this one says which
    /// path was taken.
    kept_stale: bool = false,
    /// True when `xl/calcChain.xml` and its rel and content type were
    /// dropped.
    calc_chain_removed: bool = false,
    /// §5.7.4's accounting, as of the swap.
    retained_generations: usize = 0,
    retained_bytes: u64 = 0,
    /// §5.7.9's dormant slot.
    durability: Durability = .{},

    pub fn deinit(self: *Report, gpa: Allocator) void {
        gpa.free(self.census);
        self.* = undefined;
    }
};

// ─── the candidate ───────────────────────────────────────────────

/// One prepared generation, plus everything the swap will install with it.
///
/// Owns its contents until either `swap` (which moves them into the
/// workbook) or `abandon` (which frees them). Exactly one of the two must
/// be called; a `Candidate` that is neither swapped nor abandoned leaks,
/// and the swapped flag makes a double-swap a safety-checked assertion
/// rather than a double-free.
pub const Candidate = struct {
    gpa: Allocator,
    /// The next generation, with every replacement already staged.
    next: PartStore,
    /// The re-parsed `xl/workbook.xml` view.
    workbook_view: workbook_xml_mod.WorkbookXml,
    /// New per-sheet views, parallel to `Workbook.worksheets`. A slot is
    /// non-null exactly for a sheet whose part changed *and* whose view
    /// was already materialised — a lazy slot stays lazy, and re-parsing
    /// one would be work the caller never asked for.
    sheet_views: []?sheet_xml_mod.SheetXml,
    /// Empty slots for the superseded views, allocated here so the swap's
    /// `RetainedGeneration` needs no allocator.
    retired_slots: []?sheet_xml_mod.SheetXml,
    /// Bytes the superseded generation will hold once retained.
    retired_bytes: u64,
    report: Report,
    swapped: bool = false,

    /// Give up on the candidate. The workbook is exactly as it was.
    pub fn abandon(self: *Candidate) void {
        assert(!self.swapped);
        self.workbook_view.deinit(self.gpa);
        for (self.sheet_views) |*slot| {
            if (slot.*) |*v| {
                var view = v.*;
                view.deinit(self.gpa);
            }
        }
        self.gpa.free(self.sheet_views);
        self.gpa.free(self.retired_slots);
        self.next.deinit();
        self.report.deinit(self.gpa);
        self.* = undefined;
    }

    /// Install the candidate. **Infallible**: every line below is a move
    /// or a scalar store, and the one list that grows had its capacity
    /// reserved in `prepare`.
    ///
    /// `wb.worksheets` is never reallocated, so every `*Worksheet` a
    /// caller holds stays valid and keeps pointing at the same sheet.
    pub fn swap(self: *Candidate, wb: *Workbook) void {
        assert(!self.swapped);
        assert(wb.worksheets.len == self.sheet_views.len);
        assert(wb.retained.capacity > wb.retained.items.len);

        var gen: RetainedGeneration = .{
            .store = wb.store,
            .workbook = wb.workbook,
            .sheets = self.retired_slots,
            // The workbook-scope lazy views go with the generation they
            // read, unconditionally. Keeping them would mean deciding
            // which part backs each and whether that part moved — two
            // resolutions to get right in exchange for saving a re-parse
            // nobody has asked for.
            .sst = wb.sst_view,
            .styles = wb.styles_view,
            .bytes = self.retired_bytes,
        };
        wb.sst_view = null;
        wb.styles_view = null;

        for (wb.worksheets, 0..) |*ws, i| {
            if (self.sheet_views[i]) |new_view| {
                gen.sheets[i] = ws.parsed;
                ws.parsed = new_view;
            } else {
                gen.sheets[i] = null;
            }
        }

        wb.store = self.next;
        wb.workbook = self.workbook_view;
        wb.retained.appendAssumeCapacity(gen);
        wb.retained_bytes += gen.bytes;

        self.gpa.free(self.sheet_views);
        self.swapped = true;
        self.report.retained_generations = wb.retained.items.len;
        self.report.retained_bytes = wb.retained_bytes;
    }

    /// Take the report out. Valid before or after the swap — the report is
    /// built in `prepare` and is the caller's from that moment.
    pub fn takeReport(self: *Candidate) Report {
        const r = self.report;
        self.report = .{};
        return r;
    }
};

// ─── prepare ─────────────────────────────────────────────────────

/// Stage every mutation a recalc makes into one candidate.
///
/// `staged` carries the patched sheet parts M5b1's projection produced.
/// `census` carries what the evaluator could not implement; an empty
/// census is a clean run. Nothing here touches `wb`, with one exception
/// spelled out where it happens: `Worksheet.resolvePartName` caches the
/// name it resolved, which is a memo over a fact the swap does not change.
///
/// **No allocator parameter.** Almost everything this allocates is
/// installed into `wb` by the swap and freed by `Workbook.deinit`, so it
/// has to be `wb.allocator` — the alternative is a store allocated by one
/// allocator and freed by another, which is not an ergonomic difference
/// but a heap corruption. The allocator-first convention applies to
/// functions that have a choice; this one does not.
pub fn prepare(
    wb: *Workbook,
    staged: []const StagedPart,
    census: []const Unsupported,
    opts: Options,
) Error!Result {
    const gpa = wb.allocator;
    if (cancelled(opts)) return Error.Cancelled;

    // §5.7.4's counted retention, both halves, refused BEFORE anything is
    // built. Neither number depends on the candidate — what a swap would
    // retire is the generation that already exists — so discovering the
    // ceiling afterwards would cost a whole generation's work to learn
    // something knowable now.
    if (wb.retained.items.len >= opts.max_retained_generations) {
        return .{ .refused = .{ .reason = .retained_generations_exhausted } };
    }
    const retired_bytes = generationBytes(&wb.store);
    if (wb.retained_bytes + retired_bytes > opts.max_retained_bytes) {
        return .{ .refused = .{ .reason = .retained_bytes_exhausted } };
    }

    // §5.7.7's eligibility, decided once and here. `refuse` is the
    // default; `keep_stale_and_mark` may suppress exactly the two planes
    // `calc_patch.markEligible` names, and one ineligible entry refuses
    // the whole run — a partially-suppressed census would leave the
    // caller believing a marked file was a handled one.
    var mark_only = false;
    if (census.len > 0) {
        switch (opts.on_unsupported) {
            .refuse => return .{ .refused = .{
                .reason = .unsupported_construct,
                .plane = census[0].plane,
                .census = try boundedCensus(gpa, census),
                .census_truncated = census.len > max_census_entries,
            } },
            .keep_stale_and_mark => {
                for (census) |c| {
                    if (!calc_patch.markEligible(c.plane)) return .{ .refused = .{
                        .reason = .unsupported_construct,
                        .plane = c.plane,
                        .census = try boundedCensus(gpa, census),
                        .census_truncated = census.len > max_census_entries,
                    } };
                }
                mark_only = true;
            },
        }
    }

    // `defer if (!keep)` rather than `errdefer`, everywhere below, and
    // deliberately: a refusal returns *successfully* with a `.refused`
    // payload, so an `errdefer` would not fire for it — and every refusal
    // reachable from here is reachable with a candidate already built.
    // That asymmetry is what leaked a whole generation per refused run
    // until the fixtures caught it.
    var keep = false;

    var next = try wb.store.nextGeneration();
    defer if (!keep) next.deinit();

    // Mark-only stages nothing but the mark: §5.7.7's byte-identity claim
    // is that the file differs from an un-recalculated save in exactly
    // `fullCalcOnLoad="1"`, and applying the caches would be exactly the
    // thing the caller declined.
    var sheets_patched: u32 = 0;
    var calc_chain_removed = false;
    if (!mark_only) {
        for (staged) |p| {
            _ = (try next.part(p.name)) orelse return Error.MissingSheetPart;
            try next.replacePart(p.name, p.bytes);
            sheets_patched += 1;
        }
        calc_chain_removed = try removeCalcChain(gpa, &next);
    }

    if (cancelled(opts)) return Error.Cancelled;

    // §5.7.6. Read the calc state from the candidate's own
    // `xl/workbook.xml` — a staged part may already have replaced it, and
    // planning against the workbook's bytes would address ranges the
    // candidate no longer has.
    const wb_bytes = blk: {
        const p = (try next.part(workbook_part)) orelse return Error.MissingWorkbookPart;
        break :blk p.bytes;
    };
    var state = switch (try calc.parseCalcState(gpa, wb_bytes)) {
        .ok => |s| s,
        .refused => |r| return .{
            .refused = .{ .reason = .calc_state, .plane = r.planeTwo() },
        },
    };
    defer state.deinit(gpa);

    const want = if (mark_only) calc_patch.Desired.mark_only else calc_patch.Desired.after_recalc;
    var plan = switch (try calc_patch.plan(gpa, wb_bytes, state, want)) {
        .ok => |p| p,
        .refused => |r| return .{
            .refused = .{ .reason = .calc_state, .plane = r.planeTwo() },
        },
    };
    defer plan.deinit(gpa);
    if (plan.edits().len > 0) try next.replacePart(workbook_part, plan.bytes);

    // Every view the post-swap workbook will read, parsed now.
    var built = try buildViews(gpa, wb, &next, staged, mark_only);
    defer if (!keep) built.deinit(gpa);

    var report: Report = .{
        .sheets_patched = sheets_patched,
        .census = try boundedCensus(gpa, census),
        .census_truncated = census.len > max_census_entries,
        .resolved = opts.resolved,
        .kept_stale = mark_only,
        .calc_chain_removed = calc_chain_removed,
    };
    defer if (!keep) report.deinit(gpa);

    // The last allocation the swap would otherwise have to make.
    try wb.retained.ensureUnusedCapacity(gpa, 1);

    if (cancelled(opts)) return Error.Cancelled;

    keep = true;
    return .{ .ok = .{
        .gpa = gpa,
        .next = next,
        .workbook_view = built.workbook_view,
        .sheet_views = built.sheet_views,
        .retired_slots = built.retired_slots,
        .retired_bytes = retired_bytes,
        .report = report,
    } };
}

/// §5.7.7's `markRecalcOnLoad()`, as the transaction it is.
///
/// Honestly named, and honestly implemented: the same prepare/swap with an
/// empty staged set and mark-only intent, so it gets the same no-fail swap
/// and the same retention accounting a recalc does. A `replacePart` on the
/// live store would have been three lines and would have mutated a
/// workbook whose views still described the bytes it replaced.
pub fn markRecalcOnLoad(wb: *Workbook, opts: Options) Error!Result {
    var forced = opts;
    forced.on_unsupported = .keep_stale_and_mark;
    // One synthetic census entry, of a mark-eligible plane, so the
    // mark-only path is reached through the same gate a real run takes
    // rather than through a second door into the same room.
    const census = [_]Unsupported{.{ .plane = .FormulaUnsupportedConstruct }};
    return prepare(wb, &.{}, &census, forced);
}

fn cancelled(opts: Options) bool {
    const t = opts.cancel orelse return false;
    return t.isTriggered();
}

fn boundedCensus(gpa: Allocator, census: []const Unsupported) Allocator.Error![]Unsupported {
    const n = @min(census.len, max_census_entries);
    return gpa.dupe(Unsupported, census[0..n]);
}

/// What a retained generation holds resident.
///
/// The part arena, which is where a generation's bytes actually are:
/// decompressed part payloads, staged override payloads, part names,
/// relationship attributes. The typed views are spines *over* that arena's
/// bytes — they allocate their own row and cell arrays, but those are a
/// fraction of the payloads they index, and counting them would mean
/// asking every view type for a size it does not track.
fn generationBytes(s: *const PartStore) u64 {
    return s.arena.queryCapacity();
}

// ─── the views the swap installs ─────────────────────────────────

const BuiltViews = struct {
    workbook_view: workbook_xml_mod.WorkbookXml,
    sheet_views: []?sheet_xml_mod.SheetXml,
    retired_slots: []?sheet_xml_mod.SheetXml,

    fn deinit(self: *BuiltViews, gpa: Allocator) void {
        self.workbook_view.deinit(gpa);
        for (self.sheet_views) |*slot| {
            if (slot.*) |*v| {
                var view = v.*;
                view.deinit(gpa);
            }
        }
        gpa.free(self.sheet_views);
        gpa.free(self.retired_slots);
        self.* = undefined;
    }
};

fn buildViews(
    gpa: Allocator,
    wb: *Workbook,
    next: *PartStore,
    staged: []const StagedPart,
    mark_only: bool,
) Error!BuiltViews {
    const wb_part = (try next.part(workbook_part)) orelse return Error.MissingWorkbookPart;
    var workbook_view = try workbook_xml_mod.parse(gpa, wb_part.bytes);
    errdefer workbook_view.deinit(gpa);

    const n = wb.worksheets.len;
    const sheet_views = try gpa.alloc(?sheet_xml_mod.SheetXml, n);
    errdefer gpa.free(sheet_views);
    for (sheet_views) |*slot| slot.* = null;
    errdefer for (sheet_views) |*slot| {
        if (slot.*) |*v| {
            var view = v.*;
            view.deinit(gpa);
        }
    };

    if (!mark_only) {
        for (wb.worksheets, 0..) |*ws, i| {
            // A sheet nobody has parsed stays unparsed. Its next reader
            // will build a view over the new bytes, which is the same
            // answer for less work.
            if (ws.parsed == null) continue;
            const name = try ws.resolvePartName();
            if (!isStaged(staged, name)) continue;
            const p = (try next.part(name)) orelse return Error.MissingSheetPart;
            sheet_views[i] = try sheet_xml_mod.parse(gpa, p.bytes);
        }
    }

    const retired_slots = try gpa.alloc(?sheet_xml_mod.SheetXml, n);
    for (retired_slots) |*slot| slot.* = null;

    return .{
        .workbook_view = workbook_view,
        .sheet_views = sheet_views,
        .retired_slots = retired_slots,
    };
}

fn isStaged(staged: []const StagedPart, name: []const u8) bool {
    for (staged) |p| {
        if (std.mem.eql(u8, p.name, name)) return true;
    }
    return false;
}

// ─── calcChain removal (§5.7.3 step 5) ───────────────────────────

/// Drop the calculation chain: the part, its relationship, and its
/// content-type override.
///
/// The chain records the order Excel last calculated cells in. After a
/// recalc it describes a workbook that no longer exists, and Excel treats
/// a stale chain as authoritative — so removing it is not tidying, it is
/// the difference between a file that recalculates correctly and one that
/// does not.
///
/// The rel target is resolved **relative to `xl/workbook.xml`**, which is
/// what real files write (`Target="calcChain.xml"`). `PartStore.removePart`
/// matches the bare and leading-slash forms as literal strings, so it
/// would leave `Target="./calcChain.xml"` — and a relationship pointing at
/// a part that no longer exists is a package error. Matching by resolved
/// identity is the whole reason this is not just a `removePart` call.
///
/// Returns whether anything was removed. A workbook with no calcChain
/// relationship has no calculation chain, and that is not a failure.
fn removeCalcChain(gpa: Allocator, next: *PartStore) Error!bool {
    var target: ?[]const u8 = null;
    var rel_id: ?[]const u8 = null;
    for (next.rels(workbook_part)) |rel| {
        if (!std.mem.eql(u8, rel.type, calc_chain_rel_type)) continue;
        if (rel.target_mode == .external) continue;
        target = (try next.resolve(workbook_part, rel.target)) orelse continue;
        rel_id = rel.id;
        break;
    }
    const part_name = target orelse return false;

    // Order matters: both edits address the ORIGINAL indices, and the
    // array compaction has to be last.
    if (try next.part(content_types_part)) |ct| {
        const stripped = try removeElementByAttr(gpa, ct.bytes, "Override", "PartName", part_name);
        defer gpa.free(stripped);
        if (!std.mem.eql(u8, stripped, ct.bytes)) {
            try next.replacePart(content_types_part, stripped);
        }
    }

    if (try next.part(workbook_rels_part)) |rels| {
        const stripped = try removeElementByAttr(gpa, rels.bytes, "Relationship", "Id", rel_id.?);
        defer gpa.free(stripped);
        if (!std.mem.eql(u8, stripped, rels.bytes)) {
            try next.replacePart(workbook_rels_part, stripped);
        }
    }

    try next.removePart(part_name);
    return true;
}

/// Remove every element named `local` whose unprefixed `attr` equals
/// `value`, treating a leading `/` on either side as optional.
///
/// Attribute-aware rather than a substring search. The quote style, the
/// attribute order and the leading slash are all producer choices, and the
/// two things being removed here — a content-type override for a part that
/// is gone, and a relationship pointing at it — are exactly the residues a
/// strict consumer rejects a package for.
fn removeElementByAttr(
    gpa: Allocator,
    xml: []const u8,
    local: []const u8,
    attr: []const u8,
    value: []const u8,
) Allocator.Error![]u8 {
    var out: std.ArrayListUnmanaged(u8) = .empty;
    errdefer out.deinit(gpa);
    try out.ensureTotalCapacity(gpa, xml.len);

    var sc = engine.decode.Scanner.init(xml);
    var cursor: usize = 0;
    var depth: usize = 0;
    var open: ?struct { start: usize, depth: usize } = null;
    var malformed = false;

    while (true) {
        const ev = (sc.next() catch {
            malformed = true;
            break;
        }) orelse break;
        switch (ev) {
            .close => {
                if (depth == 0) break;
                depth -= 1;
                if (open) |o| {
                    if (depth == o.depth) {
                        try out.appendSlice(gpa, xml[cursor..o.start]);
                        cursor = sc.i;
                        open = null;
                    }
                }
            },
            .open => |el| {
                depth += 1;
                if (open != null) continue;
                if (!matches(el, local, attr, value)) continue;
                open = .{ .start = el.offset, .depth = depth - 1 };
            },
            .self_closing => |el| {
                if (open != null) continue;
                if (!matches(el, local, attr, value)) continue;
                try out.appendSlice(gpa, xml[cursor..el.offset]);
                cursor = sc.i;
            },
            .text, .cdata, .doctype => {},
        }
    }

    if (malformed) {
        // Not well-formed enough to edit safely. Hand the bytes back
        // unchanged; the caller compares and leaves the part alone. A
        // partial splice of tag soup is the one outcome worse than
        // leaving a stale override in place.
        out.clearRetainingCapacity();
        try out.appendSlice(gpa, xml);
        return out.toOwnedSlice(gpa);
    }
    try out.appendSlice(gpa, xml[cursor..]);
    return out.toOwnedSlice(gpa);
}

fn matches(
    el: engine.decode.Element,
    local: []const u8,
    attr: []const u8,
    value: []const u8,
) bool {
    if (!std.mem.eql(u8, el.local(), local)) return false;
    const v = el.attr(attr) orelse return false;
    return std.mem.eql(u8, stripSlash(v), stripSlash(value));
}

fn stripSlash(s: []const u8) []const u8 {
    return if (s.len > 0 and s[0] == '/') s[1..] else s;
}

// ─── tests ───────────────────────────────────────────────────────

const testing = std.testing;

test {
    // Zig analyses only what something reaches, and this file's public
    // surface is reached from exactly one method body in `workbook.zig`.
    // Without this, a decl could be broken and still ship green — the
    // lesson `src/dbx.zig` cost between #147 and #153.
    testing.refAllDecls(@This());
}

const ns_main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const ns_r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const ct_sheet = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
const ct_workbook = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
const ct_calc_chain = "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml";
const ct_rels = "application/vnd.openxmlformats-package.relationships+xml";
const rel_worksheet = ns_r ++ "/worksheet";
const rel_officedoc = ns_r ++ "/officeDocument";

const sheet_part = "xl/worksheets/sheet1.xml";
const calc_chain_part = "xl/calcChain.xml";

const sheet_before =
    "<worksheet xmlns=\"" ++ ns_main ++ "\"><sheetData><row r=\"1\">" ++
    "<c r=\"A1\"><v>1</v></c><c r=\"B1\"><f>A1+1</f><v>2</v></c>" ++
    "</row></sheetData></worksheet>";

/// The same sheet with B1's cache recalculated — what M5b1's patcher
/// hands the transaction.
const sheet_after =
    "<worksheet xmlns=\"" ++ ns_main ++ "\"><sheetData><row r=\"1\">" ++
    "<c r=\"A1\"><v>1</v></c><c r=\"B1\"><f>A1+1</f><v>99</v></c>" ++
    "</row></sheetData></worksheet>";

const Fixture = struct {
    /// How `xl/_rels/workbook.xml.rels` spells the calc-chain target.
    /// The three forms §5.7.3 step 5 has to resolve are all producer
    /// choices, and only the first is what real files write.
    calc_chain_target: []const u8 = "calcChain.xml",
    with_calc_chain: bool = true,
    calc_pr: []const u8 = "<calcPr calcId=\"191029\"/>",
};

/// Write a minimal but real `.xlsx` and return its path (caller frees).
///
/// Built through `PartStore.fresh` + `addPart` + `save` rather than
/// checked in as a binary: every shape this row has to handle is a
/// producer choice in one string, and a corpus file cannot be varied.
fn writeFixture(
    gpa: Allocator,
    io: std.Io,
    dir: []const u8,
    name: []const u8,
    f: Fixture,
) ![]u8 {
    const path = try std.fs.path.join(gpa, &.{ dir, name });
    errdefer gpa.free(path);

    var store = try PartStore.fresh(gpa, io);
    defer store.deinit();

    try store.addPart("_rels/.rels", ct_rels, "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" ++
        "<Relationship Id=\"rId1\" Type=\"" ++ rel_officedoc ++ "\" Target=\"xl/workbook.xml\"/>" ++
        "</Relationships>");

    const wb_xml = try std.fmt.allocPrint(
        gpa,
        "<workbook xmlns=\"" ++ ns_main ++ "\" xmlns:r=\"" ++ ns_r ++ "\">" ++
            "<sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets>{s}</workbook>",
        .{f.calc_pr},
    );
    defer gpa.free(wb_xml);
    try store.addPart("xl/workbook.xml", ct_workbook, wb_xml);

    var rels: std.ArrayListUnmanaged(u8) = .empty;
    defer rels.deinit(gpa);
    try rels.appendSlice(gpa, "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" ++
        "<Relationship Id=\"rId1\" Type=\"" ++ rel_worksheet ++ "\" Target=\"worksheets/sheet1.xml\"/>");
    if (f.with_calc_chain) {
        try rels.print(
            gpa,
            "<Relationship Id=\"rId2\" Type=\"" ++ calc_chain_rel_type ++ "\" Target=\"{s}\"/>",
            .{f.calc_chain_target},
        );
    }
    try rels.appendSlice(gpa, "</Relationships>");
    try store.addPart("xl/_rels/workbook.xml.rels", ct_rels, rels.items);

    try store.addPart(sheet_part, ct_sheet, sheet_before);

    // Two parts a recalc never touches, one on each side of PartStore's
    // 1 KiB stored/deflate threshold. Raw-entry identity is a claim about
    // *stored* bytes, and a fixture whose untouched parts are all tiny
    // would only ever prove it for the STORED path.
    try store.addPart("docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml", "<cp:coreProperties xmlns:cp=\"x\"><dc:title xmlns:dc=\"y\">t</dc:title></cp:coreProperties>");
    {
        var sst: std.ArrayListUnmanaged(u8) = .empty;
        defer sst.deinit(gpa);
        try sst.appendSlice(gpa, "<sst xmlns=\"" ++ ns_main ++ "\" count=\"400\" uniqueCount=\"400\">");
        for (0..400) |i| try sst.print(gpa, "<si><t>string number {d}</t></si>", .{i});
        try sst.appendSlice(gpa, "</sst>");
        std.debug.assert(sst.items.len > 1024);
        try store.addPart(
            "xl/sharedStrings.xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
            sst.items,
        );
    }

    if (f.with_calc_chain) {
        try store.addPart(calc_chain_part, ct_calc_chain, "<calcChain xmlns=\"" ++ ns_main ++ "\"><c r=\"B1\" i=\"1\"/></calcChain>");
    }

    try store.save(io, path);
    return path;
}

const staged_one = [_]StagedPart{.{ .name = sheet_part, .bytes = sheet_after }};

fn cellText(ws: *workbook_mod.Worksheet, ref: []const u8) ![]const u8 {
    const view = try ws.ensureParsed();
    for (view.rows) |row| {
        for (row.cells) |c| {
            if (std.mem.eql(u8, c.ref, ref)) return c.raw_value orelse "";
        }
    }
    return error.CellNotFound;
}

/// Every part's (pointer, length) as of now.
///
/// Identity by pointer, not by content: after a refused prepare the
/// workbook must be holding the *same* bytes, not bytes that happen to
/// compare equal — the latter is also what a store that quietly rebuilt
/// itself would show.
fn partIdentity(gpa: Allocator, s: *PartStore) ![]const []const u8 {
    const names = try s.partNames();
    const out = try gpa.alloc([]const u8, names.len);
    errdefer gpa.free(out);
    for (names, 0..) |n, i| {
        const p = (try s.part(n)) orelse return error.MissingPart;
        out[i] = p.bytes;
    }
    return out;
}

fn expectPartsIdentical(before: []const []const u8, s: *PartStore) !void {
    const names = try s.partNames();
    try testing.expectEqual(before.len, names.len);
    for (names, 0..) |n, i| {
        const p = (try s.part(n)) orelse return error.MissingPart;
        try testing.expectEqual(before[i].ptr, p.bytes.ptr);
        try testing.expectEqual(before[i].len, p.bytes.len);
    }
}

const Harness = struct {
    threaded: *std.Io.Threaded,
    tmp: testing.TmpDir,
    dir: [:0]u8,
    path: []u8,
    wb: Workbook,
    /// The fd-budget fixture has to tear the workbook down *inside* the
    /// test — the close count it asserts is only interesting after the
    /// last generation is gone. So closing early is a supported move
    /// rather than something a test works around.
    wb_open: bool = true,

    fn init(gpa: Allocator, f: Fixture) !Harness {
        const threaded = try gpa.create(std.Io.Threaded);
        errdefer gpa.destroy(threaded);
        threaded.* = .init(gpa, .{});
        errdefer threaded.deinit();
        const the_io = threaded.io();

        var tmp = testing.tmpDir(.{});
        errdefer tmp.cleanup();
        const dir = try tmp.dir.realPathFileAlloc(the_io, ".", gpa);
        errdefer gpa.free(dir);
        const path = try writeFixture(gpa, the_io, dir, "book.xlsx", f);
        errdefer gpa.free(path);
        const wb = try Workbook.open(gpa, the_io, path);
        return .{ .threaded = threaded, .tmp = tmp, .dir = dir, .path = path, .wb = wb };
    }

    fn io(self: *Harness) std.Io {
        return self.threaded.io();
    }

    fn closeWorkbook(self: *Harness) void {
        std.debug.assert(self.wb_open);
        self.wb.deinit();
        self.wb_open = false;
    }

    fn deinit(self: *Harness, gpa: Allocator) void {
        if (self.wb_open) self.wb.deinit();
        gpa.free(self.path);
        gpa.free(self.dir);
        self.tmp.cleanup();
        self.threaded.deinit();
        gpa.destroy(self.threaded);
    }
};

test "prepare/swap: the candidate's bytes are what the workbook reads after" {
    const gpa = testing.allocator;
    var h = try Harness.init(gpa, .{});
    defer h.deinit(gpa);

    const ws = try h.wb.sheet(0);
    try testing.expectEqualStrings("2", try cellText(ws, "B1"));

    var r = try prepare(&h.wb, &staged_one, &.{}, .{});
    var c = switch (r) {
        .ok => |*v| v,
        .refused => return error.UnexpectedRefusal,
    };
    // Before the swap the workbook is still the workbook it was.
    try testing.expectEqualStrings("2", try cellText(ws, "B1"));

    c.swap(&h.wb);
    var report = c.takeReport();
    defer report.deinit(gpa);

    try testing.expectEqualStrings("99", try cellText(ws, "B1"));
    try testing.expectEqual(@as(u32, 1), report.sheets_patched);
    try testing.expect(report.calc_chain_removed);
    try testing.expect(!report.kept_stale);
    try testing.expectEqual(@as(usize, 1), report.retained_generations);
    try testing.expect(report.retained_bytes > 0);
}

test "swap: worksheet slot addresses survive a nonstructural recalc" {
    const gpa = testing.allocator;
    var h = try Harness.init(gpa, .{});
    defer h.deinit(gpa);

    const before = try h.wb.sheet(0);
    _ = try before.ensureParsed();

    var r = try prepare(&h.wb, &staged_one, &.{}, .{});
    switch (r) {
        .ok => |*c| c.swap(&h.wb),
        .refused => return error.UnexpectedRefusal,
    }
    var rep = r.ok.takeReport();
    rep.deinit(gpa);

    const after = try h.wb.sheet(0);
    try testing.expectEqual(before, after);
    // And the handle the caller was already holding reads the new value.
    try testing.expectEqualStrings("99", try cellText(before, "B1"));
}

test "swap: a borrow taken before the recalc still reads what it read" {
    const gpa = testing.allocator;
    var h = try Harness.init(gpa, .{});
    defer h.deinit(gpa);

    const ws = try h.wb.sheet(0);
    // A slice borrowed from generation 0's part bytes. §5.7.4 promises it
    // stays readable for the Workbook's lifetime — which is only true
    // because the swap retains the whole superseded generation.
    const borrowed = try cellText(ws, "B1");
    try testing.expectEqualStrings("2", borrowed);

    for (0..3) |_| {
        var r = try prepare(&h.wb, &staged_one, &.{}, .{});
        switch (r) {
            .ok => |*c| c.swap(&h.wb),
            .refused => return error.UnexpectedRefusal,
        }
        var rep = r.ok.takeReport();
        rep.deinit(gpa);
        // Still the bytes it was, three generations later.
        try testing.expectEqualStrings("2", borrowed);
    }
    try testing.expectEqual(@as(usize, 3), h.wb.retained.items.len);
}

test "retention: four generations share one file descriptor" {
    const gpa = testing.allocator;
    var h = try Harness.init(gpa, .{});
    defer h.deinit(gpa);

    // M5b0's ledger, attached post-open: it is only read when the last
    // reference drops, so when it is installed does not matter — only
    // that it is installed before any of them do.
    var ledger: store_mod.CloseLedger = .{};
    h.wb.store.backing.ledger = &ledger;

    for (0..4) |_| {
        var r = try prepare(&h.wb, &staged_one, &.{}, .{});
        switch (r) {
            .ok => |*c| c.swap(&h.wb),
            .refused => return error.UnexpectedRefusal,
        }
        var rep = r.ok.takeReport();
        rep.deinit(gpa);
    }

    try testing.expectEqual(@as(usize, 4), h.wb.retained.items.len);
    // Four retained plus the live one, over ONE backing.
    try testing.expectEqual(@as(usize, 5), h.wb.store.backing.refCount());
    try testing.expectEqual(@as(usize, 0), ledger.closes);

    // M5b0's ownership claim, re-run against the retained path: every
    // superseded generation is still readable, and each holds the bytes
    // it held — generation 0 the original cache, the rest the new one.
    for (h.wb.retained.items, 0..) |*g, i| {
        const p = (try g.store.part(sheet_part)) orelse return error.MissingPart;
        if (i == 0) {
            try testing.expectEqualStrings(sheet_before, p.bytes);
        } else {
            try testing.expectEqualStrings(sheet_after, p.bytes);
        }
    }

    h.closeWorkbook();
    // Exactly one close, whichever generation went last. That is the
    // whole fd budget for a five-deep set: §5.7.4 retains generations,
    // M5b0 makes them share the file one of them opened.
    try testing.expectEqual(@as(usize, 1), ledger.closes);
}

test "retention: the cap refuses before the swap, and frees nothing" {
    const gpa = testing.allocator;
    var h = try Harness.init(gpa, .{});
    defer h.deinit(gpa);

    const opts: Options = .{ .max_retained_generations = 2 };
    for (0..2) |_| {
        var r = try prepare(&h.wb, &staged_one, &.{}, opts);
        switch (r) {
            .ok => |*c| c.swap(&h.wb),
            .refused => return error.UnexpectedRefusal,
        }
        var rep = r.ok.takeReport();
        rep.deinit(gpa);
    }

    const identity = try partIdentity(gpa, &h.wb.store);
    defer gpa.free(identity);

    const third = try prepare(&h.wb, &staged_one, &.{}, opts);
    switch (third) {
        .ok => return error.ExpectedRefusal,
        .refused => |ref| {
            try testing.expectEqual(Refusal.Reason.retained_generations_exhausted, ref.reason);
            try testing.expectEqual(PlaneTwo.FormulaLimitExceeded, ref.planeTwo());
            try testing.expectEqual(
                @as(workbook_mod.Error, error.FormulaLimitExceeded),
                ref.toWorkbookError(),
            );
        },
    }
    // Nothing was reclaimed to make room, and nothing changed.
    try testing.expectEqual(@as(usize, 2), h.wb.retained.items.len);
    try expectPartsIdentical(identity, &h.wb.store);
}

test "retention: the byte ceiling refuses too" {
    const gpa = testing.allocator;
    var h = try Harness.init(gpa, .{});
    defer h.deinit(gpa);

    const r = try prepare(&h.wb, &staged_one, &.{}, .{ .max_retained_bytes = 1 });
    switch (r) {
        .ok => return error.ExpectedRefusal,
        .refused => |ref| try testing.expectEqual(
            Refusal.Reason.retained_bytes_exhausted,
            ref.reason,
        ),
    }
    try testing.expectEqual(@as(usize, 0), h.wb.retained.items.len);
}

test "post-failure reads: a refused prepare leaves every part byte-identical" {
    const gpa = testing.allocator;
    var h = try Harness.init(gpa, .{ .calc_pr = "<calcPr calcId=\"1\" fullPrecision=\"0\"/>" });
    defer h.deinit(gpa);

    const ws = try h.wb.sheet(0);
    try testing.expectEqualStrings("2", try cellText(ws, "B1"));
    const identity = try partIdentity(gpa, &h.wb.store);
    defer gpa.free(identity);

    const r = try prepare(&h.wb, &staged_one, &.{}, .{});
    switch (r) {
        .ok => return error.ExpectedRefusal,
        .refused => |ref| {
            try testing.expectEqual(Refusal.Reason.calc_state, ref.reason);
            try testing.expectEqual(PlaneTwo.FormulaPrecisionAsDisplayed, ref.planeTwo());
        },
    }

    try expectPartsIdentical(identity, &h.wb.store);
    try testing.expectEqualStrings("2", try cellText(ws, "B1"));
    try testing.expectEqual(@as(usize, 1), h.wb.sheetCount());
    try testing.expectEqual(@as(usize, 0), h.wb.retained.items.len);
}

test "post-failure reads: a cancelled prepare leaves memory untouched" {
    const gpa = testing.allocator;
    var h = try Harness.init(gpa, .{});
    defer h.deinit(gpa);

    const ws = try h.wb.sheet(0);
    _ = try cellText(ws, "B1");
    const identity = try partIdentity(gpa, &h.wb.store);
    defer gpa.free(identity);

    var flag: u8 = 1;
    const token: CancelToken = .{ .flag = &flag };
    try testing.expectError(
        Error.Cancelled,
        prepare(&h.wb, &staged_one, &.{}, .{ .cancel = token }),
    );

    try expectPartsIdentical(identity, &h.wb.store);
    try testing.expectEqualStrings("2", try cellText(ws, "B1"));
    try testing.expectEqual(@as(usize, 0), h.wb.retained.items.len);
}

test "abandon: a prepared candidate the caller drops changes nothing" {
    const gpa = testing.allocator;
    var h = try Harness.init(gpa, .{});
    defer h.deinit(gpa);

    const ws = try h.wb.sheet(0);
    _ = try cellText(ws, "B1");
    const identity = try partIdentity(gpa, &h.wb.store);
    defer gpa.free(identity);

    var r = try prepare(&h.wb, &staged_one, &.{}, .{});
    switch (r) {
        .ok => |*c| c.abandon(),
        .refused => return error.UnexpectedRefusal,
    }

    try expectPartsIdentical(identity, &h.wb.store);
    try testing.expectEqualStrings("2", try cellText(ws, "B1"));
    try testing.expectEqual(@as(usize, 0), h.wb.retained.items.len);
}

test "no-fail swap: an allocator that fails everything still swaps" {
    const gpa = testing.allocator;
    var h = try Harness.init(gpa, .{});
    defer h.deinit(gpa);

    const ws = try h.wb.sheet(0);
    _ = try ws.ensureParsed();

    var r = try prepare(&h.wb, &staged_one, &.{}, .{});
    var c = switch (r) {
        .ok => |*v| v,
        .refused => return error.UnexpectedRefusal,
    };

    // The proof, rather than the assertion: run the swap under an
    // allocator whose very first request fails. It passes because the
    // swap makes no request — every byte it needs was allocated in
    // prepare, and the one list it appends to had its capacity reserved
    // there. `free` still reaches the real allocator, which is what the
    // swap does do.
    var failing = testing.FailingAllocator.init(gpa, .{ .fail_index = 0 });
    c.gpa = failing.allocator();
    c.swap(&h.wb);
    try testing.expectEqual(@as(usize, 0), failing.allocations);

    var report = c.takeReport();
    defer report.deinit(gpa);
    try testing.expectEqualStrings("99", try cellText(ws, "B1"));
}

/// One prepare, from a freshly-opened workbook, with every allocation on
/// `a`. Returns the error `checkAllAllocationFailures` needs to see, but
/// not before proving the workbook it failed on is unchanged.
fn preparePassUnderFailure(a: Allocator, io: std.Io, path: []const u8) !void {
    var wb = try Workbook.open(a, io, path);
    defer wb.deinit();

    const ws = try wb.sheet(0);
    _ = try ws.ensureParsed();
    const identity = try partIdentity(a, &wb.store);
    defer a.free(identity);

    var r = prepare(&wb, &staged_one, &.{}, .{}) catch |e| {
        // The whole claim of "all allocation in prepare": whichever
        // request failed, it failed with the workbook still whole.
        try expectPartsIdentical(identity, &wb.store);
        try testing.expectEqual(@as(usize, 0), wb.retained.items.len);
        try testing.expectEqualStrings("2", try cellText(ws, "B1"));
        return e;
    };
    switch (r) {
        .ok => |*c| {
            c.swap(&wb);
            var report = c.takeReport();
            report.deinit(a);
            try testing.expectEqualStrings("99", try cellText(ws, "B1"));
        },
        .refused => return error.UnexpectedRefusal,
    }
}

test "no-fail swap: every allocation failure lands before the swap" {
    const gpa = testing.allocator;
    var h = try Harness.init(gpa, .{});
    defer h.deinit(gpa);
    try testing.checkAllAllocationFailures(
        gpa,
        preparePassUnderFailure,
        .{ h.io(), h.path },
    );
}

// ─── calcChain removal (§5.7.3 step 5) ───────────────────────────

fn expectCalcChainGone(wb: *Workbook) !void {
    const names = try wb.store.partNames();
    for (names) |n| {
        if (std.mem.eql(u8, n, calc_chain_part)) return error.CalcChainPartStillPresent;
    }
    for (wb.store.rels(workbook_part)) |rel| {
        if (std.mem.eql(u8, rel.type, calc_chain_rel_type)) return error.CalcChainRelStillPresent;
    }
    const ct = (try wb.store.part(content_types_part)) orelse return error.MissingPart;
    if (std.mem.indexOf(u8, ct.bytes, "calcChain") != null) return error.CalcChainContentTypeStillPresent;
    const rels = (try wb.store.part(workbook_rels_part)) orelse return error.MissingPart;
    if (std.mem.indexOf(u8, rels.bytes, "calcChain") != null) return error.CalcChainRelBytesStillPresent;
}

test "calcChain: every rel-target spelling resolves to the same part" {
    const gpa = testing.allocator;
    // Relative (what producers write), absolute, and two noncanonical
    // forms. §4 recorded the bug this covers: `removeRelationshipsTo`
    // matches the absolute spelling as a literal string, so the relative
    // one — the common case — would have been left behind.
    const targets = [_][]const u8{
        "calcChain.xml",
        "/xl/calcChain.xml",
        "./calcChain.xml",
        "worksheets/../calcChain.xml",
    };
    for (targets) |t| {
        var h = try Harness.init(gpa, .{ .calc_chain_target = t });
        defer h.deinit(gpa);

        // Present before.
        try testing.expect((try h.wb.store.part(calc_chain_part)) != null);

        var r = try prepare(&h.wb, &staged_one, &.{}, .{});
        switch (r) {
            .ok => |*c| c.swap(&h.wb),
            .refused => return error.UnexpectedRefusal,
        }
        var rep = r.ok.takeReport();
        defer rep.deinit(gpa);
        try testing.expect(rep.calc_chain_removed);
        try expectCalcChainGone(&h.wb);
    }
}

test "calcChain: a workbook without one is not a failure" {
    const gpa = testing.allocator;
    var h = try Harness.init(gpa, .{ .with_calc_chain = false });
    defer h.deinit(gpa);

    var r = try prepare(&h.wb, &staged_one, &.{}, .{});
    switch (r) {
        .ok => |*c| c.swap(&h.wb),
        .refused => return error.UnexpectedRefusal,
    }
    var rep = r.ok.takeReport();
    defer rep.deinit(gpa);
    try testing.expect(!rep.calc_chain_removed);
}

test "calcChain: an external target is not a package part" {
    const gpa = testing.allocator;
    var h = try Harness.init(gpa, .{});
    defer h.deinit(gpa);
    // Not reachable through the fixture (TargetMode is an attribute the
    // builder does not set), so the guard is asserted where it lives.
    const stripped = try removeElementByAttr(
        gpa,
        "<Relationships><Relationship Id=\"rId2\" Target=\"x\"/></Relationships>",
        "Relationship",
        "Id",
        "rId9",
    );
    defer gpa.free(stripped);
    try testing.expectEqualStrings(
        "<Relationships><Relationship Id=\"rId2\" Target=\"x\"/></Relationships>",
        stripped,
    );
}

test "calcChain: the remover is attribute-aware, not a substring search" {
    const gpa = testing.allocator;
    // Single quotes, attribute order reversed, and a decoy whose PartName
    // merely *contains* the target. A substring matcher removes the decoy
    // and misses the real one.
    const xml =
        "<Types><Override PartName='/xl/calcChain.xml.bak' ContentType='a'/>" ++
        "<Override ContentType='b' PartName='/xl/calcChain.xml'/></Types>";
    const stripped = try removeElementByAttr(gpa, xml, "Override", "PartName", "xl/calcChain.xml");
    defer gpa.free(stripped);
    try testing.expectEqualStrings(
        "<Types><Override PartName='/xl/calcChain.xml.bak' ContentType='a'/></Types>",
        stripped,
    );
}

test "calcChain: a non-self-closing element is removed whole" {
    const gpa = testing.allocator;
    const xml = "<Types><Override PartName=\"/a.xml\"><x/></Override><Override PartName=\"/b.xml\"/></Types>";
    const stripped = try removeElementByAttr(gpa, xml, "Override", "PartName", "a.xml");
    defer gpa.free(stripped);
    try testing.expectEqualStrings("<Types><Override PartName=\"/b.xml\"/></Types>", stripped);
}

test "calcChain: tag soup is handed back unchanged rather than half-spliced" {
    const gpa = testing.allocator;
    const xml = "<Types><Override PartName=\"/a.xml\"";
    const stripped = try removeElementByAttr(gpa, xml, "Override", "PartName", "a.xml");
    defer gpa.free(stripped);
    try testing.expectEqualStrings(xml, stripped);
}

// ─── §5.7.7 refusal purity ───────────────────────────────────────

fn saveAndRead(gpa: Allocator, wb: *Workbook, io: std.Io, dir: []const u8, name: []const u8) ![]u8 {
    const out = try std.fs.path.join(gpa, &.{ dir, name });
    defer gpa.free(out);
    try wb.store.save(io, out);
    return std.Io.Dir.cwd().readFileAlloc(io, out, gpa, .limited(1 << 22));
}

test "refusal purity: a refused recalc saves the bytes an unrecalculated one does" {
    const gpa = testing.allocator;
    var h = try Harness.init(gpa, .{ .calc_pr = "<calcPr calcId=\"1\" fullPrecision=\"0\"/>" });
    defer h.deinit(gpa);

    const plain = try saveAndRead(gpa, &h.wb, h.io(), h.dir, "plain.xlsx");
    defer gpa.free(plain);

    const r = try prepare(&h.wb, &staged_one, &.{}, .{});
    try testing.expect(r == .refused);

    const after_refusal = try saveAndRead(gpa, &h.wb, h.io(), h.dir, "refused.xlsx");
    defer gpa.free(after_refusal);
    try testing.expectEqualSlices(u8, plain, after_refusal);
}

test "refusal purity: mark-only differs in exactly fullCalcOnLoad" {
    const gpa = testing.allocator;
    var h = try Harness.init(gpa, .{});
    defer h.deinit(gpa);

    const before_names = try h.wb.store.partNames();
    const before = try gpa.alloc([]u8, before_names.len);
    defer {
        for (before) |b| gpa.free(b);
        gpa.free(before);
    }
    for (before_names, 0..) |n, i| {
        const p = (try h.wb.store.part(n)) orelse return error.MissingPart;
        before[i] = try gpa.dupe(u8, p.bytes);
    }

    const census = [_]Unsupported{.{ .plane = .FormulaUnsupportedFunction, .row = 1, .col = 1 }};
    var r = try prepare(&h.wb, &staged_one, &census, .{ .on_unsupported = .keep_stale_and_mark });
    switch (r) {
        .ok => |*c| c.swap(&h.wb),
        .refused => return error.UnexpectedRefusal,
    }
    var rep = r.ok.takeReport();
    defer rep.deinit(gpa);
    try testing.expect(rep.kept_stale);
    // The caches were NOT applied, and the chain that describes them
    // stayed: a mark-only run did not calculate anything.
    try testing.expectEqual(@as(u32, 0), rep.sheets_patched);
    try testing.expect(!rep.calc_chain_removed);
    try testing.expectEqual(@as(usize, 1), rep.census.len);
    try testing.expect(!rep.census_truncated);

    // Every part except `xl/workbook.xml` is byte-identical, and that one
    // differs in exactly the inserted attribute — named as a range over
    // the two documents, not counted as a number of edits.
    const after_names = try h.wb.store.partNames();
    try testing.expectEqual(before_names.len, after_names.len);
    for (after_names, 0..) |n, i| {
        const p = (try h.wb.store.part(n)) orelse return error.MissingPart;
        if (!std.mem.eql(u8, n, workbook_part)) {
            try testing.expectEqualSlices(u8, before[i], p.bytes);
            continue;
        }
        const w = engine.calc_patch.changedWindow(before[i], p.bytes).?;
        try testing.expectEqual(w.start, w.end);
        const inserted = " fullCalcOnLoad=\"1\"";
        try testing.expectEqualStrings(inserted, p.bytes[w.start..][0..inserted.len]);
        // And `calcId` is untouched: mark-only makes no producer claim.
        try testing.expect(std.mem.indexOf(u8, p.bytes, "calcId=\"191029\"") != null);
    }
}

test "refusal purity: an ineligible plane refuses even under keep-stale-and-mark" {
    const gpa = testing.allocator;
    var h = try Harness.init(gpa, .{});
    defer h.deinit(gpa);

    // One eligible entry and one that is not. §5.7.7 refuses the whole
    // run rather than marking the part it could: a partially-suppressed
    // census leaves the caller believing a marked file was handled.
    const census = [_]Unsupported{
        .{ .plane = .FormulaUnsupportedConstruct },
        .{ .plane = .FormulaCycle },
    };
    const r = try prepare(&h.wb, &staged_one, &census, .{ .on_unsupported = .keep_stale_and_mark });
    switch (r) {
        .ok => return error.ExpectedRefusal,
        .refused => |ref| {
            var refusal = ref;
            defer refusal.deinit(gpa);
            try testing.expectEqual(Refusal.Reason.unsupported_construct, refusal.reason);
            try testing.expectEqual(PlaneTwo.FormulaCycle, refusal.planeTwo());
            try testing.expectEqual(
                @as(workbook_mod.Error, error.FormulaCycle),
                refusal.toWorkbookError(),
            );
            // M9a2: the refusal carries the census it refused over —
            // BOTH entries, not just the plane the error names.
            try testing.expectEqual(@as(usize, 2), refusal.census.len);
            try testing.expectEqual(PlaneTwo.FormulaUnsupportedConstruct, refusal.census[0].plane);
            try testing.expectEqual(PlaneTwo.FormulaCycle, refusal.census[1].plane);
            try testing.expect(!refusal.census_truncated);
        },
    }
    try testing.expectEqual(@as(usize, 0), h.wb.retained.items.len);
}

test "refusal purity: the default refuses on the first census entry" {
    const gpa = testing.allocator;
    var h = try Harness.init(gpa, .{});
    defer h.deinit(gpa);
    const census = [_]Unsupported{.{ .plane = .FormulaUnsupportedFunction, .sheet = 0, .row = 3, .col = 2 }};
    const r = try prepare(&h.wb, &staged_one, &census, .{});
    switch (r) {
        .ok => return error.ExpectedRefusal,
        .refused => |ref| {
            var refusal = ref;
            defer refusal.deinit(gpa);
            try testing.expectEqual(
                PlaneTwo.FormulaUnsupportedFunction,
                refusal.planeTwo(),
            );
            // M9a2: the refusing cell survives the collapse.
            try testing.expectEqual(@as(usize, 1), refusal.census.len);
            try testing.expectEqual(@as(u32, 3), refusal.census[0].row);
            try testing.expectEqual(@as(u32, 2), refusal.census[0].col);
        },
    }
}

test "markRecalcOnLoad: the mark, and nothing else" {
    const gpa = testing.allocator;
    var h = try Harness.init(gpa, .{});
    defer h.deinit(gpa);

    const before = blk: {
        const p = (try h.wb.store.part(workbook_part)) orelse return error.MissingPart;
        break :blk try gpa.dupe(u8, p.bytes);
    };
    defer gpa.free(before);

    try h.wb.markRecalcOnLoad();

    const after = (try h.wb.store.part(workbook_part)) orelse return error.MissingPart;
    const w = engine.calc_patch.changedWindow(before, after.bytes).?;
    try testing.expectEqualStrings(
        " fullCalcOnLoad=\"1\"",
        after.bytes[w.start..][0.." fullCalcOnLoad=\"1\"".len],
    );
    try testing.expect(std.mem.indexOf(u8, after.bytes, "calcId=\"191029\"") != null);
    // The chain still describes the caches, because the caches are still
    // the ones it described.
    try testing.expect((try h.wb.store.part(calc_chain_part)) != null);
    try testing.expectEqual(@as(usize, 1), h.wb.retained.items.len);
}

test "markRecalcOnLoad: fullPrecision=\"0\" refuses through the Workbook error set" {
    const gpa = testing.allocator;
    var h = try Harness.init(gpa, .{ .calc_pr = "<calcPr calcId=\"1\" fullPrecision=\"0\"/>" });
    defer h.deinit(gpa);
    try testing.expectError(error.FormulaPrecisionAsDisplayed, h.wb.markRecalcOnLoad());
}

// ─── §5.7.6 through the transaction ──────────────────────────────

test "calc state: a successful recalc writes the pair at the byte level" {
    const gpa = testing.allocator;
    var h = try Harness.init(gpa, .{ .calc_pr = "<calcPr calcId=\"191029\" calcMode=\"manual\" iterate=\"1\" iterateCount=\"42\"/>" });
    defer h.deinit(gpa);

    var r = try prepare(&h.wb, &staged_one, &.{}, .{});
    switch (r) {
        .ok => |*c| c.swap(&h.wb),
        .refused => return error.UnexpectedRefusal,
    }
    var rep = r.ok.takeReport();
    defer rep.deinit(gpa);

    const p = (try h.wb.store.part(workbook_part)) orelse return error.MissingPart;
    try testing.expect(std.mem.indexOf(u8, p.bytes, "calcId=\"0\"") != null);
    try testing.expect(std.mem.indexOf(u8, p.bytes, "fullCalcOnLoad=\"1\"") != null);
    try testing.expect(std.mem.indexOf(u8, p.bytes, "calcId=\"191029\"") == null);
    // Preserved, untouched.
    try testing.expect(std.mem.indexOf(u8, p.bytes, "calcMode=\"manual\"") != null);
    try testing.expect(std.mem.indexOf(u8, p.bytes, "iterate=\"1\"") != null);
    try testing.expect(std.mem.indexOf(u8, p.bytes, "iterateCount=\"42\"") != null);
}

test "calc state: a workbook with no calcPr gains one at its schema position" {
    const gpa = testing.allocator;
    var h = try Harness.init(gpa, .{ .calc_pr = "" });
    defer h.deinit(gpa);

    var r = try prepare(&h.wb, &staged_one, &.{}, .{});
    switch (r) {
        .ok => |*c| c.swap(&h.wb),
        .refused => return error.UnexpectedRefusal,
    }
    var rep = r.ok.takeReport();
    defer rep.deinit(gpa);

    const p = (try h.wb.store.part(workbook_part)) orelse return error.MissingPart;
    try testing.expect(std.mem.indexOf(u8, p.bytes, "</sheets><calcPr calcId=\"0\" fullCalcOnLoad=\"1\"/></workbook>") != null);
}

// ─── raw-entry identity (§5.7.9's byte promise, in memory) ───────

/// The stored (still compressed) payload of `name` inside `archive`.
fn rawPayload(s: *PartStore, archive: []const u8, name: []const u8) ![]const u8 {
    const names = try s.partNames();
    for (names, 0..) |n, i| {
        if (!std.mem.eql(u8, n, name)) continue;
        const e = s.entries[i];
        return archive[e.payload_offset..][0..e.compressed_size];
    }
    return error.MissingPart;
}

test "raw-entry identity: an untouched part keeps its stored compressed bytes" {
    const gpa = testing.allocator;
    var h = try Harness.init(gpa, .{});
    defer h.deinit(gpa);

    const source = try std.Io.Dir.cwd().readFileAlloc(h.io(), h.path, gpa, .limited(1 << 22));
    defer gpa.free(source);

    var r = try prepare(&h.wb, &staged_one, &.{}, .{});
    switch (r) {
        .ok => |*c| c.swap(&h.wb),
        .refused => return error.UnexpectedRefusal,
    }
    var rep = r.ok.takeReport();
    defer rep.deinit(gpa);

    const saved = try saveAndRead(gpa, &h.wb, h.io(), h.dir, "recalced.xlsx");
    defer gpa.free(saved);

    var src_store = try PartStore.openBuffer(gpa, h.io(), source);
    defer src_store.deinit();
    var out_store = try PartStore.openBuffer(gpa, h.io(), saved);
    defer out_store.deinit();

    // Everything the recalc did not address: same decompressed bytes,
    // same compression method, same CRC, and the same *stored* payload —
    // the part was copied, not re-emitted. One STORED part and one
    // DEFLATEd one, so "including its stored compressed bytes" is proven
    // on both sides of the store/deflate threshold.
    const untouched = [_][]const u8{ "_rels/.rels", "docProps/core.xml", "xl/sharedStrings.xml" };
    var saw_deflate = false;
    for (untouched) |name| {
        const a = (try src_store.part(name)) orelse return error.MissingPart;
        if (a.compression_method == 8) saw_deflate = true;
        const b = (try out_store.part(name)) orelse return error.MissingPart;
        try testing.expectEqualSlices(u8, a.bytes, b.bytes);
        try testing.expectEqual(a.compression_method, b.compression_method);
        try testing.expectEqualSlices(
            u8,
            try rawPayload(&src_store, source, name),
            try rawPayload(&out_store, saved, name),
        );
    }
    try testing.expect(saw_deflate);
    // And the one it did address is gone from the output entirely.
    for (try out_store.partNames()) |n| {
        try testing.expect(!std.mem.eql(u8, n, calc_chain_part));
    }
}

// ─── the report (§5.7.8) ─────────────────────────────────────────

test "report: bounded census, truncation flag, and the dormant durability slot" {
    const gpa = testing.allocator;
    var h = try Harness.init(gpa, .{});
    defer h.deinit(gpa);

    const big = try gpa.alloc(Unsupported, max_census_entries + 7);
    defer gpa.free(big);
    for (big) |*e| e.* = .{ .plane = .FormulaUnsupportedConstruct };

    var r = try prepare(&h.wb, &.{}, big, .{ .on_unsupported = .keep_stale_and_mark });
    switch (r) {
        .ok => |*c| c.swap(&h.wb),
        .refused => return error.UnexpectedRefusal,
    }
    var rep = r.ok.takeReport();
    defer rep.deinit(gpa);

    try testing.expectEqual(max_census_entries, rep.census.len);
    try testing.expect(rep.census_truncated);

    // Preallocated and dormant: flipping it is two scalar stores, which
    // is what makes it usable at the one point in §5.7.9 where allocation
    // is no longer permitted.
    try testing.expect(!rep.durability.warning);
    rep.durability.warn(5);
    try testing.expect(rep.durability.warning);
    try testing.expectEqual(@as(i32, 5), rep.durability.err_code);
}

test "M5d1: a real AtomicFile.Commit flips the dormant slot without allocating" {
    const gpa = testing.allocator;
    var h = try Harness.init(gpa, .{});
    defer h.deinit(gpa);

    var r = try prepare(&h.wb, &.{}, &.{}, .{});
    switch (r) {
        .ok => |*c| c.swap(&h.wb),
        .refused => return error.UnexpectedRefusal,
    }
    var rep = r.ok.takeReport();
    defer rep.deinit(gpa);
    try testing.expect(!rep.durability.warning);

    // The value M5d1's commit region actually produces, carried across
    // the layer boundary as data rather than re-derived here.
    const commit: atomic_file.Commit = .{
        .durability_warning = true,
        .durability_errno = @intFromEnum(std.posix.E.IO),
    };

    // Under an allocator that fails every request: §5.7.9 discovers a
    // post-rename fsync failure at the one point where allocation is no
    // longer permitted, so the transfer has to be pure stores. A `warn`
    // that ever grew anything would fail here rather than in production
    // on the day a disk filled up.
    var failing = testing.FailingAllocator.init(gpa, .{ .fail_index = 0 });
    const never = failing.allocator();
    try testing.expectError(error.OutOfMemory, never.alloc(u8, 1));

    if (commit.durability_warning) rep.durability.warn(commit.durability_errno);
    try testing.expect(rep.durability.warning);
    try testing.expectEqual(@intFromEnum(std.posix.E.IO), rep.durability.err_code);
    try testing.expectEqual(@as(usize, 0), failing.allocations);
}

test "report: the resolved run inputs are echoed verbatim" {
    const gpa = testing.allocator;
    var h = try Harness.init(gpa, .{});
    defer h.deinit(gpa);

    const run: run_inputs.RunInputs = .{ .now_utc_ms = 1_700_000_000_000, .rng_seed = 7, .limits = .{} };
    const eff = run.effective(.recalc);
    var r = try prepare(&h.wb, &staged_one, &.{}, .{ .resolved = eff });
    switch (r) {
        .ok => |*c| c.swap(&h.wb),
        .refused => return error.UnexpectedRefusal,
    }
    var rep = r.ok.takeReport();
    defer rep.deinit(gpa);
    try testing.expect(rep.resolved.?.eql(eff));
    // A recalc derives dialect per stored cell, so the projection has
    // none to echo.
    try testing.expect(rep.resolved.?.dialect == null);
}

test "prepare: a staged part the package does not have is an error, not a new part" {
    const gpa = testing.allocator;
    var h = try Harness.init(gpa, .{});
    defer h.deinit(gpa);

    const identity = try partIdentity(gpa, &h.wb.store);
    defer gpa.free(identity);

    const bogus = [_]StagedPart{.{ .name = "xl/worksheets/sheet9.xml", .bytes = sheet_after }};
    try testing.expectError(error.MissingSheetPart, prepare(&h.wb, &bogus, &.{}, .{}));
    try expectPartsIdentical(identity, &h.wb.store);
    try testing.expectEqual(@as(usize, 0), h.wb.retained.items.len);
}

test "swap: workbook-scope lazy views go with the generation that backs them" {
    const gpa = testing.allocator;
    var h = try Harness.init(gpa, .{});
    defer h.deinit(gpa);

    // Materialise the SST so there is a view to retire, and borrow from
    // it: the string has to still read as itself afterwards, which is
    // only true because the view went into the retained generation
    // instead of being freed.
    const borrowed = (try h.wb.sstText(3)).?;
    try testing.expectEqualStrings("string number 3", borrowed);

    var r = try prepare(&h.wb, &staged_one, &.{}, .{});
    switch (r) {
        .ok => |*c| c.swap(&h.wb),
        .refused => return error.UnexpectedRefusal,
    }
    var rep = r.ok.takeReport();
    defer rep.deinit(gpa);

    try testing.expect(h.wb.sst_view == null);
    try testing.expect(h.wb.retained.items[0].sst != null);
    try testing.expectEqualStrings("string number 3", borrowed);
    // And re-reading it through the new generation gives the same answer.
    try testing.expectEqualStrings("string number 3", (try h.wb.sstText(3)).?);
}
