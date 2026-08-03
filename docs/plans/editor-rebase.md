# Editor rebase onto Workbook (B2 plan, v1 draft)

> Tier B2 in `post-0.2.9-roadmap.md`. B0 (PartStore) and B1 (Workbook
> typed overlay) are shipped — `Workbook.open` / `setCell` /
> `save` / `renameSheet` / `rewriteAllFormulas` are the substrate
> this plan rebases onto. The current `Editor` (`src/xlsx.zig:4360`)
> is a substring-and-substitute mutation layer with five shadow
> tables. B2 retires those tables in favour of routing every
> mutation through the typed Workbook, while preserving Editor's
> conservative refusal posture and its public API for one minor
> line.

## Status (closed 2026-05-09)

**B2 fully shipped.** All iters merged through PR #89.

- iter-er-0 ✅ (Editor relocated to `pkg/editor.zig`; nested-types hoist) — PRs #61, #65
- iter-er-1 ✅ (Read-side parity — `Editor.open` constructs an internal `Workbook` via `Workbook.fromBook`) — PR #66
- iter-er-2 ✅ (`setCell` / `setCells` rebase — existing-sheet routes through `Worksheet.setCell` + `emitWithDeltas`) — PRs #67, #68
- iter-er-3 ✅ (`appendRows` rebase — `Worksheet.appendRows` API + `Worksheet.emitWithAppends` substring-splice fast-path; bench gate green at 1.016×) — PRs #69, #70, #71, #73
- iter-er-4 ✅ (structural edits — addSheet / renameSheet / deleteSheet / row+col edits all flow through Workbook): sheet-level via PRs #74 + #77; row/col edits via PR #82 (Editor delegates `insertRow`/`deleteRow`/`insertColumn`/`deleteColumn` to typed-overlay Workbook; legacy `pending_row_*`/`pending_col_*` queues retired)
- iter-er-5 ✅ (refusal-list lifts) — 5 of 5 axes lifted: rename_sheet (#79); four row/col axes (formula / defined-names / hyperlink / DV+CF — all in #82); `delete_sheet` rewriter variant (#81) + final `SheetDeleteWithDefinedNamesNotSupported` lift (#89). Per-sheet drawing / pivot / pane / autoFilter / table refusals stay (no rewriter exists, lifting silently corrupts — see [refusal-audit.md](refusal-audit.md))
- iter-er-6 ✅ (`Editor.save` rebase to `Workbook.save` — single emit path): phases 1–3 dead-state cleanup (#83, -2062 LOC) + thin-shim proper (#84, -1238 LOC); `Editor.save` is now a 14-line shim (passthrough preserves SHA256 byte-identity for no-mutation saves; mutated path delegates to `Workbook.save` → `PartStore.save`)
- iter-er-7 ✅ (corpus parity sweep + perf bench): task A corpus parity (#86, +9 tests, 0 bug surfaces); task B bench refresh (#85, all gates green at 1.076× of 174 ms canonical baseline; ≤ 1.10× strict and ≤ 1.50× loose ceilings both PASS); task C Workbook.save edge-case tests + retire legacy `Worksheet.emitWithAppends` (#87). Bonus: Codex review found a HIGH bug — empty SST string mis-resolved to rich entry index (#88).

**Final state**: pkg/editor.zig 6021 → 3231 LOC (-46%). `pkg/workbook.zig` owns the SST extension plan + per-sheet emit + atomic ZIP rebuild via `PartStore.save`. The four cross-sheet rewriters (`rewriteAllFormulas`, `rewriteAllDefinedNames`, `rewriteAllHyperlinkLocations`, `rewriteAllValidationsAndConditionalFormats`) are wired into `Workbook.{insertRow, deleteRow, insertColumn, deleteColumn, renameSheet, deleteSheet}`. The next track is **B3** (Writer rebase onto Workbook) — see `post-0.2.9-roadmap.md` Tier B and the proposed iter-wr-0..7 in `~/.claude/projects/-Users-lf-Projects-Pro-zlsx/memory/project_iter_er_3_pre_work.md`.

## Problem

`Editor` (`src/xlsx.zig:4360`) is the load-modify-save (Phase 3c)
mutation surface. It owns five shadow tables that the save path
walks in a fixed order:

- `pending_appends` (`AppendBuffer` per sheet, `src/xlsx.zig:4382`)
- `pending_new_sheets` / `pending_renames` / `pending_deletes`
  (`src/xlsx.zig:4389-4402`)
- `pending_row_inserts` / `pending_row_deletes` /
  `pending_col_inserts` / `pending_col_deletes`
  (`src/xlsx.zig:4411-4423`)
- `pending_mutations` (`MutatedSheet` per sheet, holds decompressed
  XML + an in-sync `CellSpan` index, `src/xlsx.zig:4432`)

Each table has its own conservative-pending-op rules. Cross-table
interactions are guarded by hand: `pending_appends` and
`pending_mutations` cannot coexist on the same sheet
(`error.SheetHasUnsavedAppends` / `error.SheetHasUnsavedMutations`);
`deleteSheet` refuses if any other table has a pending entry; row
and column edits both call `anySheetCrossSheetCarrier`
(`src/xlsx.zig:6024`) to refuse sheets carrying formulas, hyperlinks,
data validations, conditional formats, defined names, drawings,
panes, tables, or autofilters.

Three structural problems fall out:

1. **Two save pipelines.** `Editor.save` (`src/xlsx.zig:4683`) walks
   the five tables, decompresses-mutates-recompresses the affected
   parts, and rewrites the central directory. `Workbook.save`
   (`pkg/workbook.zig:410`) emits the same shape via PartStore.
   Two byte-preserving emit paths through the same ZIP layer is the
   primary duplicate-source-of-truth risk flagged in the B1 plan
   ("`Workbook.save` divergence from `Editor.save`",
   `docs/plans/workbook-overlay.md:265`). B2 collapses them.
2. **Per-category guard sprawl.** Every new structural-edit
   capability (insertRow, deleteCol, renameSheet) bolts a new guard
   onto the existing surface. Workbook centralises mutation in
   `Worksheet.setCell` + `Workbook.rewriteAllFormulas`; B2 lets the
   refusal list shrink to a single audit point per category.
3. **Paves B3.** The writer rebase (B3) needs a single source of
   truth for SST + styles + dirty-sheet emission. Once Editor is
   gone as an independent emit path, B3 is "swap fluent-writer's
   output for `Workbook.save`" rather than "unify three emit paths".

The win is concrete: one mutable model, fewer guard categories,
unblocks B3.

## Constraints

- **Refusal posture preserved.** Editor MUST never silently corrupt
  a workbook. Every refusal that exists today (cross-sheet carrier
  refuses on row/col edits, last-sheet refuses on `deleteSheet`,
  conflicting pending-op categories) stays unless iter-er-5 makes a
  considered decision to lift it (e.g. C1 M2 m1.5
  `rewriteAllFormulas` now liberates the formula axis of the
  refusal — see `Workbook.rewriteAllFormulas` shipped in PR #34).
- **`appendRows` substring fast-path retained while it remains the
  production format.** Phase 3a (load-modify-save row append) is
  the hottest path zlsx ships; a regression there is unacceptable.
  iter-er-3 keeps the existing substring-and-substitute emit under
  a feature gate (`Editor.appendRows` continues to use the legacy
  path until parity tests + bench gates are green on the
  Workbook-routed alternative).
- **One-minor-line compat facade.** `Editor.open` /
  `setCell` / `setCells` / `appendRows` / `addSheet` /
  `renameSheet` / `deleteSheet` / `insertRow` / `deleteRow` /
  `insertColumn` / `deleteColumn` / `save` / `deinit` keep their
  current signatures. Public field access (e.g. `editor.sheet_paths`)
  is preserved where production callers depend on it; internal
  shadow tables become private once the rebase ships.
- **Module-graph constraint (Zig 0.15.2).** Editor lives in
  `src/xlsx.zig`; Workbook lives in `pkg/`. The rebase imports
  `pkg/workbook.zig` from `src/xlsx.zig` (already done at
  `src/xlsx.zig`'s formula-rewriter call sites). No new
  cross-package collisions are introduced.
- **Single-threaded contract per `Editor`.** Same shape as today
  and as Workbook.

## Public surface (sketch)

Editor stays binary-compatible. The new shape is an
adapter-internal-Workbook:

```zig
pub const Editor = struct {
    allocator: Allocator,
    /// Lives inside Editor; constructed in `open` via
    /// `Workbook.fromBook`. All mutations flow through it.
    workbook: pkg.Workbook,
    /// Retained for one minor line so existing callers reading
    /// `editor.sheet_paths` keep working. Sourced from
    /// `workbook.sheetCount` + per-sheet path lookup at construction.
    sheet_paths: []const []const u8,

    pub fn open(allocator, path) !Editor;          // = Workbook.open + retain path mapping
    pub fn deinit(self: *Editor) void;             // = workbook.deinit + free sheet_paths

    pub fn setCell(self, sheet_idx, ref, value) !void;  // → Worksheet.setCell
    pub fn setCells(self, sheet_idx, refs, values) !void;
    pub fn appendRows(self, sheet_idx, rows) !void;     // → Worksheet append helper
    pub fn addSheet(self, name) !u32;                   // → Workbook.addSheet (B1 follow-up)
    pub fn renameSheet(self, idx, new_name) !void;      // → Workbook.renameSheet (PR #35)
    pub fn deleteSheet(self, idx) !void;                // → Workbook.deleteSheet (B1 follow-up)
    pub fn insertRow(self, idx, before_row) !void;      // → Worksheet structural helper
    pub fn deleteRow(self, idx, row) !void;
    pub fn insertColumn(self, idx, before_col) !void;
    pub fn deleteColumn(self, idx, col) !void;

    pub fn save(self, out_path) !void;                  // = workbook.save
};
```

The five `pending_*` tables disappear from the struct; their state
moves into Workbook's per-Worksheet delta map + Workbook's typed
state views.

## Iter sequencing

### iter-er-1 — Read-side parity (1 week)

**Scope:** `Editor.open` continues to call `Book.open` for path
resolution + sheet rels (Editor's existing setup), then promotes
the result to a `Workbook` via `Workbook.fromBook(book, path)`
(shipped in PR #29) and stores it as `editor.workbook`.
`editor.sheet_paths` is populated from the Workbook side.

**Delta:** no mutation paths are rerouted yet; this iter only
stands the substrate up. Existing `Editor.setCell` /
`appendRows` / `save` paths still walk the shadow tables.

**Tests:** `tests/editor_corpus.zig` opens every fixture under
both APIs; assert identical sheet count, sheet names, cell
content via `Worksheet.cell` and the existing `scanWorksheet`
spans.

**Walk-away:** if `Workbook.fromBook` ownership tangles with
`Editor`'s `src_buf` / `entries` lifetime (two readers over the
same backing file), revert to `Workbook.open(allocator, path)`
and accept the second open + parse cost. Document the layered
ownership.

### iter-er-2 — `setCell` / `setCells` rebase (1–2 weeks)

**Scope:** route every cell mutation through
`Worksheet.setCell` (`pkg/workbook.zig:1846`). Retire the
`pending_mutations` shadow table + the `MutatedSheet` /
`CellSpan` index it owns.

**Conversion:** map Editor's existing `Cell` shape to the
Workbook `CellValue` union (blank / number / boolean /
inlineStr / shared_string / formula). Both surfaces already
exist; the rebase is mechanical.

**Tests:** every `tests/editor_*.zig` setCell test runs
unchanged; add a test that `pending_mutations.count() == 0`
across an end-to-end mutate-and-save cycle (proves the table
is genuinely retired).

**Walk-away:** if `Worksheet.setCell` does NOT accept a value
shape Editor produces today (e.g. a Date variant Editor
synthesises), extend `CellValue` in B1 follow-up FIRST, then
resume iter-er-2. Do not fork the union for Editor.

### iter-er-3 — `appendRows` rebase (2 weeks) — TRICKIEST ITER

> **Risk callout.** This is the highest-stakes iter in B2.
> `Editor.appendRows` (Phase 3a, `src/xlsx.zig:4421`) is the
> production hot path: read-modify-append-save on multi-sheet
> workbooks is the workload zlsx is most often deployed for.
> The current substring-and-substitute fast-path measurably beats
> a per-cell `Worksheet.setCell` loop on append-heavy workloads
> (the Workbook delta map re-emits `<sheetData>` from scratch;
> the substring path splices `</sheetData>` and inlines the new
> rows). A naive rebase regresses Phase 3a perf by an
> unacceptable margin.

**Scope:** introduce `Worksheet.appendRows(rows)` on the
Workbook side that internally takes the substring fast-path
when no other delta has been recorded for the sheet, falls
back to the per-cell delta map otherwise. Editor's
`appendRows` becomes a thin pass-through. The legacy
`pending_appends` shadow table is retired only after the
parity + bench gates below are green for two consecutive runs.

**Walk-away gate (tightened):**

- Append-heavy bench (the existing `tests/bench/append_rows.zig`
  if present, else add one) shows ≤ 1.10× regression vs the
  legacy substring path on a 100k-row × 5-sheet fixture. Above
  1.10×, keep the legacy path in production behind a build-time
  flag and ship the rebase as opt-in only.
- Corpus parity sweep: every fixture that exercises
  `Editor.appendRows` round-trips byte-equivalent to the legacy
  path output (modulo whitespace-in-XML differences that don't
  affect Excel's open).

**Walk-away:** if Workbook's `appendRows` cannot match the
substring path within 1.10× even after a fast-path inside
Worksheet, leave Editor's `pending_appends` in place
indefinitely and document the divergence. Editor stays a
two-path mutator for that one method until B3 unifies the
emit pipeline anyway.

### iter-er-4 — Structural edits (2 weeks)

**Scope:** route `addSheet` / `renameSheet` / `deleteSheet` /
`insertRow` / `deleteRow` / `insertColumn` / `deleteColumn`
through Workbook. `renameSheet` already lands cleanly via
`Workbook.renameSheet` (PR #35). `deleteSheet` and `addSheet`
need Workbook-side helpers that don't yet exist as of
2026-05-03; either ship them as a B1 follow-up before B2
starts iter-er-4, or include them as the first two PRs of
iter-er-4.

**Cross-sheet formula handling:** the existing `deleteSheet`
guard refuses if cross-sheet refs exist. Now that
`Workbook.rewriteAllFormulas` (PR #34) is shipped, iter-er-4
can either (a) keep the refusal and document
`workbook.rewriteAllFormulas(.delete_sheet { … })` as the
caller's escape hatch, or (b) lift the refusal and call
`rewriteAllFormulas` internally. iter-er-5 makes the call
after auditing every guard.

**Tests:** every `tests/editor_sheet_*.zig` and
`tests/editor_row_*.zig` / `tests/editor_col_*.zig` runs
unchanged. Plus a `pending_*` count assertion symmetric to
iter-er-2's.

### iter-er-5 — Refusal-list audit (1 week)

**Scope:** walk every guard the existing Editor carries and
decide per-category:

- **Lift now** (paved by C1 M2 m1.5): formula-axis refusals on
  row/col edits — `rewriteAllFormulas` handles the rewrite.
- **Lift after C1 M2 m2** (DV / CF rewriter): data-validation
  and conditional-format axes.
- **Lift after C1 M2 m3** (defined-names + hyperlink rewriter):
  defined-names + internal-hyperlink axes.
- **Stay refused indefinitely**: drawings, pivots, panes,
  tables, autofilters (no rewriter exists; lifting silently
  corrupts).
- **Stay refused, ergonomic only**: `deleteSheet` on a workbook
  with one sheet (Excel rejects this on open).

**Output:** a checklist in the plan doc; one PR per axis
flipped, each adding parity tests that prove the lift doesn't
regress refused-shape detection.

### iter-er-6 — `Editor.save` rebase to `Workbook.save` (1 week)

**Scope:** retire `Editor.save`'s body. The function becomes:

```zig
pub fn save(self: *Editor, out_path: []const u8) !void {
    return self.workbook.save(out_path);
}
```

The five shadow tables are gone by this iter; everything
they tracked lives inside `Workbook` already.

**Walk-away gate:** corpus parity sweep — every fixture that
exercised `Editor.save` round-trips byte-equivalent to the
pre-rebase output. Plus the perf gate from the roadmap:
≤ 1.5× current ~5 ms ZIP-substitution latency on small
workbooks (`tests/bench/editor_save.zig`).

### iter-er-7 — Corpus parity sweep + perf bench (1 week)

**Scope:** end-to-end corpus sweep: open every fixture, apply
a representative mutation set (setCell / appendRows /
addSheet / renameSheet / insertRow), save, re-open, assert
parity vs the pre-rebase outputs captured in iter-er-1.

**Bench:** publish a before/after table covering `setCell`,
`appendRows`, `addSheet`, `save` on small / medium / large
fixtures. The roadmap's gate (≤ 1.5× ZIP-substitution
latency on small workbooks) is the merge-blocker; iter-er-3's
1.10× appendRows gate is the secondary gate.

**Walk-away:** if any single mutation regresses > 1.5× even
after iter-er-3's fast-path, keep the corresponding legacy
path under a feature flag and ship the rebase incremental.

## Walk-away gates (summary)

| Iter | Gate | Failure mode |
|---|---|---|
| iter-er-1 | Workbook-vs-Editor read parity on every corpus fixture | revert to `Workbook.open` not `fromBook` |
| iter-er-2 | `pending_mutations.count() == 0` post-save | freeze and debug delta-map vs span-index divergence |
| iter-er-3 | append-rows ≤ 1.10× legacy on 100k×5 fixture | keep legacy path in production behind a build flag |
| iter-er-4 | structural-edit corpus parity | per-category PR-by-PR rebase, not big-bang |
| iter-er-5 | every refusal classified into lift-now / lift-later / stay | document each "stay" with the rewriter that would unblock it |
| iter-er-6 | byte-equivalent save vs pre-rebase outputs | revert to dual-save-path until B3 |
| iter-er-7 | ≤ 1.5× ZIP-substitution latency on small workbooks | feature-flag the rebase, ship incremental |

## Risks

- **Phase 3a (appendRows) regression** — see iter-er-3 callout.
  The substring fast-path is measurably faster than a per-cell
  delta-map round-trip; the rebase MUST keep an internal
  fast-path or accept a permanent legacy fork for this one
  method.
- **Phase 3c (load-modify-save) regression** — Editor's
  `setCell` is widely used in production. iter-er-2's parity
  sweep + iter-er-7's full bench are the safety net; any > 1.5×
  regression on `Editor.setCell` blocks merge.
- **Refusal-list drift** — silently lifting a refusal that the
  existing rewriter doesn't fully cover corrupts files. iter-er-5
  is the dedicated audit; per-axis PRs add tests proving the
  refusal still fires on shapes the rewriter doesn't cover.
- **Internal-Book / internal-Workbook ownership** — iter-er-1
  inherits B1 iter-wb-2's lifetime risk: `Worksheet.rows()`
  borrows from the internal Book; `Editor.deinit` ordering must
  keep Book alive until all `Rows` are consumed. Mitigate via
  explicit lifetime comments + `std.testing.allocator` leak
  detection on every iter-er-1/2 test.
- **One-minor-line compat drift** — public Editor fields
  (`editor.sheet_paths`) need a deprecation path. iter-er-1
  retains them; iter-er-7 documents the deprecation; the next
  minor line removes them.

## Out of scope (explicit)

- **B3 Writer rebase onto Workbook** — separate plan, after B2.
- **C1 M2 m2** (DV / CF rewriter) and **C1 M2 m3** (defined-names
  + hyperlinks rewriter) — those rewriters are independent of
  Editor; iter-er-5 simply documents which Editor refusals lift
  once they ship.
- **Tier-D items** (D1 evaluator, D2 chart emit) — never touched
  by B2. *(Historical: accurate for B2's scope. D1's project-level
  "deferred indefinitely" status was reversed 2026-08-02 — see
  `goal_formula.md` — which does not change what B2 shipped.)*
- **C2b `addImage`** — depends on B1 + C2a; orthogonal to B2.
- **Editor public-field cleanup** — `editor.sheet_paths` and
  similar deprecate in this minor line; removal is the next
  minor.

## Estimation

| Iter | Estimate | Depends on |
|---|---|---|
| iter-er-1 read-side parity | 1 w | B1 (shipped) |
| iter-er-2 `setCell` rebase | 1–2 w | iter-er-1 |
| iter-er-3 `appendRows` rebase | 2 w | iter-er-2 + perf-bench infra |
| iter-er-4 structural edits | 2 w | iter-er-2 + Workbook `addSheet` / `deleteSheet` (B1 follow-up) |
| iter-er-5 refusal-list audit | 1 w | iter-er-4 |
| iter-er-6 `save` rebase | 1 w | iter-er-2..5 |
| iter-er-7 corpus + perf sweep | 1 w | iter-er-6 |

End-to-end: **4–6 weeks** (matches the roadmap estimate).
