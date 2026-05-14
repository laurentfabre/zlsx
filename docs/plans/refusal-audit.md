# Editor refusal-list audit (B2 iter-er-5)

> Walk every guard the current `Editor` carries and decide per-category:
> lift now, lift after a specific rewriter ships, stay refused. The
> lift PRs are scheduled separately from this audit; this doc is the
> checklist + decision rationale.

## Status

iter-er-5 audit complete (2026-05-05); all 5 lifts shipped (final
`SheetDeleteWithDefinedNamesNotSupported` lift on 2026-05-08 — the
`delete_sheet` rewriter variant from `feat(formula+workbook):
delete_sheet rewriter variant — final iter-er-5 lift unlock`
landed first). 4 axes stay refused (no rewriter exists; lifting
silently corrupts).

## Methodology

Read `pkg/editor.zig` after iter-er-4 sheet-level migration shipped
(commits b2f7fda + aa74a90). Surfaced every `return error.*` whose
shape is "refuse a structurally-correct edit because we can't yet
rewrite a downstream reference." Classified each by:

- **Axis** — what kind of cross-reference triggers the refusal
  (formula / hyperlink / DV / CF / defined-name / drawing / pivot /
  pane / autoFilter / table / last-sheet / pending-state /
  ZIP-impossible).
- **Trigger function** — `recordRowEdit` / `recordColEdit` /
  `deleteSheet`.
- **Today's rewriter status** — does a Workbook-side rewriter
  exist that could safely update the references?
- **Decision** — lift now, lift after rewriter X, or stay refused.

## Axes scheduled for lift

Each row maps to one follow-up PR. The PR pattern: drop the
specific guard branch, route the affected mutator through the
existing Workbook rewriter, add parity tests proving the lifted
edge cases survive a round-trip without `#REF!`.

| Axis | Trigger functions | Error today | Rewriter | Notes |
|---|---|---|---|---|
| **Formula** | `recordRowEdit`, `recordColEdit` (via `anySheetCrossSheetCarrier` `<f>` scan) | `RowEditWithFormulasNotSupported` / `ColEditWithFormulasNotSupported` | `Workbook.rewriteAllFormulas(edit)` (PR #34) | Insert/delete row+col rewrites land on every `<f>` cell across every sheet. Absolute-marker preservation, `target_sheet`-scoped bare refs, range collapse to `#REF!`. Lift = drop the `<f>` branch in `anySheetCrossSheetCarrier` + call `rewriteAllFormulas` from save. |
| **Defined names** | `recordRowEdit`, `recordColEdit`, `deleteSheet` (workbook.xml `<definedName>` scan) | `RowEditWithDefinedNamesNotSupported` / `ColEditWithDefinedNamesNotSupported` / `SheetDeleteWithDefinedNamesNotSupported` | `Workbook.rewriteAllDefinedNames(edit, target_sheet)` (PR #59) | Walks every `<definedName>` body, runs the formula rewriter on the inner ref. Workbook + sheet-scoped (`localSheetId`) handled. Lift = drop the workbook.xml `<definedName>` scan in all three sites + call `rewriteAllDefinedNames` from save. |
| **Hyperlink** | `recordRowEdit`, `recordColEdit` (via `anySheetCrossSheetCarrier` `<hyperlinks>` scan) | `RowEditUnsafeForSheet` / `ColEditUnsafeForSheet` (subset) | `Workbook.rewriteAllHyperlinkLocations(edit, target_sheet)` (PR #59) | Internal hyperlinks (`location="Sheet2!A1"`); external URLs are unaffected by row/col edits and don't need rewriting. Lift = drop the `<hyperlinks>` branch in `anySheetCrossSheetCarrier` + call `rewriteAllHyperlinkLocations` from save. |
| **Data validation** | `recordRowEdit`, `recordColEdit` (via `anySheetCrossSheetCarrier` `<dataValidations>` scan) | `RowEditUnsafeForSheet` / `ColEditUnsafeForSheet` (subset) | `Workbook.rewriteAllValidationsAndConditionalFormats(edit, target_sheet)` (PR #48) | Byte-level rewrite of `<formula1>` / `<formula2>` inner-text spans; preserves every surrounding attribute (errorTitle, etc.). Lift = drop the `<dataValidations>` branch + call rewriter. |
| **Conditional format** | `recordRowEdit`, `recordColEdit` (via `anySheetCrossSheetCarrier` `<conditionalFormatting>` scan) | `RowEditUnsafeForSheet` / `ColEditUnsafeForSheet` (subset) | `Workbook.rewriteAllValidationsAndConditionalFormats(edit, target_sheet)` (PR #48) | Same rewriter as DV. `<cfRule><formula>` body rewritten; dxf_id / priority / surrounding attrs preserved. Lift = drop the `<conditionalFormatting>` branch + call rewriter (already covered if DV lift runs the same call). |

After all five lifts ship, `anySheetCrossSheetCarrier` collapses to
a no-op and can be deleted. The `RowEditUnsafeForSheet` /
`ColEditUnsafeForSheet` errors retire — the per-sheet drawing /
pivot / pane / table / autoFilter checks below keep their own
narrower error codes.

## Axes that stay refused (no rewriter, lifting corrupts)

| Axis | Trigger | Error today | Why stay |
|---|---|---|---|
| **Drawings / pictures** | `recordColEdit` per-sheet `<drawing` / `<picture` / `<legacyDrawing` scan | `ColEditUnsafeForSheet` | Drawing anchors carry `<from>` / `<to>` cell coordinates. No rewriter exists; lifting would silently misalign every embedded image after a row/col shift. C2b shipped image emit (anchored at zero offset, fixed extent) but no rewriter for existing anchors yet. |
| **Pivots / tableParts / autoFilter** | `recordColEdit` per-sheet scan | `ColEditUnsafeForSheet` | Pivot caches and table ranges are byte-encoded across multiple parts (`xl/pivotCache/*`, `xl/tables/*.xml`). No rewriter exists and the cross-part ref graph is non-trivial. Stay refused indefinitely. |
| ~~**Frozen panes**~~ | ~~`recordColEdit` per-sheet `<pane ` scan~~ | ~~`ColEditUnsafeForSheet`~~ | ✅ **Lifted 2026-05-11**. `pkg/sheet_edit.zig` now shifts `xSplit`/`ySplit` + `topLeftCell` for `state="frozen"` / `state="frozenSplit"` panes during the byte transform. `state="split"` (or absent state, OOXML default = split) carries pixel offsets and surfaces `error.SplitPaneNotSupported`. Editor's row+col refusal guards no longer scan for `<pane>`. |
| **`deleteSheet` on a single-sheet workbook** | `Workbook.deleteSheet` via `Editor.deleteSheet` | `LastSheetUndeletable` (in `pkg/workbook.Error`) | Excel rejects the file on open if the workbook has zero sheets. Ergonomic refusal — stays. |

## Refusals that aren't axes (model invariants)

These refusals are model-level invariants, not "we can't yet
rewrite this" — they stay forever:

- `RowEditRequiresCleanSheet` / `ColEditRequiresCleanSheet` —
  refuses when a sheet has staged appends, deltas, or other
  pending row/col edits. The save-time substitution and the
  staging buffers build XML differently; merging needs design
  work that isn't scoped. **Stays.**
- `SheetDeleteRequiresCleanState` — refuses when *any* table has
  pending entries because `deleteSheet` rebuilds `sheet_paths`,
  invalidating cached indices. **Stays** unless iter-er-6 wires
  index-shifting at save time.
- `Zip64NotSupported` / `ZipSplitNotSupported` /
  `ZipDataDescriptorNotSupported` / `ZipEncryptedNotSupported` —
  archive-format limits we don't implement. **Stay.**

## Sequencing for the lift PRs

The five axes split by what blocks them:

**Lifted in `Workbook.renameSheet` (axis = `rename_sheet`)** — shipped
post-audit. Adding the three remaining rewriter calls
(`rewriteAllDefinedNames`, `rewriteAllHyperlinkLocations`,
`rewriteAllValidationsAndConditionalFormats`) to
`Workbook.renameSheet` covers cross-sheet refs from defined names,
internal hyperlinks, and DV/CF formulas. `Editor.renameSheet`
already routes through `Workbook.renameSheet` post iter-er-4 (2/N).

**Blocked on iter-er-4 (3/N)** — the four row/col edit lifts
(formula / defined-names / hyperlink / DV+CF) require typed-overlay
row/col shifts before they can compose with the rewriters'
delta-writes. Editor's `applyRowEditToWorksheet` operates on the
sheet's source bytes; the rewriter writes new formulas to the
workbook's delta map. The two paths emit the same cells via
different routes — without iter-er-4 (3/N) routing row/col edits
through the typed-overlay, lifting the row/col axis refusals
would ship a workbook where row attrs shift but formula refs
don't.

**Blocked on `delete_sheet` rewriter variant** — the
`deleteSheet` axis can't lift via the existing rewriters because
the formula-rewriter's `RewriteEdit` union has no `delete_sheet`
arm. Adding it (cross-sheet refs to the deleted sheet → `#REF!`)
is a small extension to `src/formula/rewriter.zig` + wiring into
`Workbook.deleteSheet`.

After iter-er-4 (3/N) ships, the four row/col axis lifts become
tractable single-PR drops (call existing rewriter, drop the
guard).

After the `delete_sheet` rewriter variant ships, the `deleteSheet`
defined-names refusal lifts trivially.

After all five ship: delete `anySheetCrossSheetCarrier` +
`RowEditUnsafeForSheet` / `ColEditUnsafeForSheet` (the latter two
stay only for the per-sheet drawing/pivot/pane/autoFilter cases,
which keep narrower error names).

## What this audit explicitly does NOT do

- It does not lift any guard. Lifts are per-PR work scheduled
  after this audit lands.
- It does not extend the rewriter coverage. The rewriters exist
  and are tested in their own suites; this audit just identifies
  where Editor can call them.
- It does not touch the row/col edit pipeline's per-sheet drawing
  / pivot / pane scan — those stay refused per "no rewriter."
- It does not address C2b drawing-anchor rewriter (a future iter,
  out of scope for B2).

## How this fits into iter-er-6

iter-er-6 turns `Editor.save` into a thin shim over
`Workbook.save`. The five lifts above can ship before iter-er-6 (
they call into Workbook's existing rewriters at save time and the
results land via the Editor's existing save pipeline). After
iter-er-6, the lifts naturally thread through Workbook's emit
path; no duplicate logic.
