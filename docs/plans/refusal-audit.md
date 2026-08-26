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
| ~~**Drawings (xdr)**~~ | ~~`recordRowEdit` / `recordColEdit` per-sheet `<drawing` scan~~ | ~~`RowEditUnsafeForSheet` / `ColEditUnsafeForSheet`~~ | ✅ **Lifted 2026-05-14** (iter dr-1). New `pkg/drawing_edit.zig` walks `<xdr:twoCellAnchor>` and `<xdr:oneCellAnchor>` blocks inside the referenced `xl/drawings/drawingN.xml`, shifting `<xdr:col>` (0-based, ECMA-376 §20.5.2.13) on col edits and `<xdr:row>` on row edits. Full-collapse anchors drop entirely. `<xdr:absoluteAnchor>` (pixel coords) passes through. Wired through `Workbook.applyDrawingEditForSheet`. v1 limitation: hard-codes `xdr:` namespace prefix. New typed errors: `MalformedDrawingXml`, `DrawingCoordinateOverflow`. |
| ~~**Legacy drawings (VML)**~~ | ~~`recordRowEdit` / `recordColEdit` per-sheet `<legacyDrawing` scan~~ | ~~`RowEditUnsafeForSheet` / `ColEditUnsafeForSheet`~~ | ✅ **Lifted 2026-05-14** (iter dr-2). New `pkg/vml_edit.zig` walks `<v:shape>` blocks in `xl/drawings/vmlDrawingN.vml`. For shapes with `<x:ClientData>`, shifts `<x:Row>` / `<x:Column>` (0-based) AND the matching from/to pair within the 8-int `<x:Anchor>` payload (FC/TC for cols, FR/TR for rows). Drop on delete-match against the anchor cell. PAIRED comments part rewrite (REL-705): every VML note has a `<comment ref>` in `xl/commentsN.xml` that MUST stay synchronized — `applyEditToCommentsXml` shifts/drops in step. Wired through `Workbook.applyVmlDrawingEditForSheet` (resolves both VML and comments parts via the `/relationships/comments` Type suffix). v1 limitation: hard-codes `v:` and `x:` prefixes. New typed errors: `MalformedVmlDrawing`, `VmlCoordinateOverflow`, `MalformedCommentsXml`. |
| ~~**Background `<picture>`**~~ | ~~`recordRowEdit` / `recordColEdit` per-sheet `<picture` scan~~ | ~~`RowEditUnsafeForSheet` / `ColEditUnsafeForSheet`~~ | ✅ **Lifted 2026-05-14** (iter dr-0). CT_SheetBackgroundPicture (ECMA-376 §18.3.1.67) is a single coordinate-free `r:id` reference to a tiled background image — row/col edits cannot misalign it. Editor's row+col guards no longer scan for `<picture`; the element passes through the byte transform unchanged. |
| ~~**Structured tables (`<tableParts>`)**~~ | ~~`recordRowEdit` / `recordColEdit` per-sheet `<tableParts` scan~~ | ~~`RowEditUnsafeForSheet` / `ColEditUnsafeForSheet`~~ | ✅ **Lifted 2026-05-14** (iter tbl-1). New `pkg/table_edit.zig` walks each `xl/tables/tableN.xml` referenced from the sheet's `<tableParts>` block (resolved through sheet rels). Shifts `<table ref>`, inner `<autoFilter ref>` (delegated to `sheet_edit.processAutoFilterTagCol`/`Row`, including `<filterColumn colId>` rebase), and inner `<sortState ref>` — closing the prior autoFilter caveat. On col edits, drops the matching `<tableColumn>` on a delete and adds a synthetic `<tableColumn id="max+1" name="ColumnN"/>` on an insert; `<tableColumns count=>` updates accordingly. Wired through `Workbook.applyTableEditsForSheet`. Editor pre-flight via `Workbook.preflightTableEditsForSheet` dry-runs the transform per table part BEFORE any sheet bytes are mutated, so the all-or-nothing contract holds. v1 limitations (still refused — schema-invalid table states, not unrewritten ones): row-delete on the table's header row → `TableHeaderRowDeleteUnsafe` (always when `headerRowCount >= 1`, the default); delete that would collapse the range to zero columns or zero rows → `TableCollapseUnsafe`; both surface as `RowEditUnsafeForSheet` / `ColEditUnsafeForSheet` to Editor callers. `<extLst>` table extensions pass through verbatim; `totalsRowCount > 0` BR-row deletes are not refused (lossy — let Excel recompute). 12 pure-function tests + 5 Editor round-trip tests pin the lift. New typed errors: `MalformedTableXml`, `TableCoordinateOverflow`, `TableCollapseUnsafe`, `TableHeaderRowDeleteUnsafe`. |
| **Pivots** | `recordRowEdit` / `recordColEdit` per-sheet **rels scan** (`isPivotRelType`: trailing-URI-segment match on `pivotTable` / `pivotCacheDefinition` / `pivotCacheRecords`) | `RowEditUnsafeForSheet` / `ColEditUnsafeForSheet` | ✅ **Refusal shipped 2026-07-29 (#139).** This row used to read "_no scan; refused at consumer level_" — that consumer-level refusal never existed: neither `pkg/editor.zig` nor `pkg/sheet_edit.zig` contained the string "pivot", so a row insert shifted the grid and left every pivot coordinate describing the old layout, silently. Exactly the class the method note below warns about — no guard to enumerate, so the scan-derived list could not see it. Detection is by relationship type, not sheet-body scan: a pivot leaves no *required* marker in worksheet XML (`<pivotSelection>` is optional), while the r:id edge to the pivot part is precisely what makes the pivot depend on this sheet's coordinates. Over-matching costs a refusal; under-matching costs a corrupted workbook. Checked before the table pre-flight — a rels lookup is cheaper than a dry-run transform per table part. The cross-part rewriter (`<location ref>` + cache field ranges across `xl/pivotTables/*` / `xl/pivotCache/*`) stays a later lift; zlsx's writer never emits pivots. |
| ~~**`<extLst>` extensions (`x14:`/`x15:`)**~~ | ~~`preflightExtensionEditsForSheet` (`<xm:f>` presence scan)~~ | ~~`ExtensionEditUnsafe` → `RowEditUnsafeForSheet` / `ColEditUnsafeForSheet`~~ | ✅ **Lifted in two steps: `xm:sqref` 2026-07-29 (#140), `xm:f` 2026-08-26 (S2).** `xm:sqref` element text (the range an extension covers — same meaning in every extension that carries it, appears nowhere else) is shifted by leaf-element name via `shiftRefOrRange` in the byte transform, so `x14:conditionalFormatting` / `x14:dataValidation` / `x14:sparklineGroup` / `x14:ignoredErrors` stay aligned after row/col edits, and a future `x14:` element carrying `xm:sqref` is correct for free. Failure posture there: an unterminated element or unparseable ref emits the original bytes — a partially-shifted list would be worse than an unshifted one. `<xm:f>` — the extensions' *formula* leaf: a sparkline's data range and its group's date axis, `x14:cfRule` / `x14:cfvo` expressions, `x14:formula1` / `formula2` — is a formula, not a range, so it rides the formula rewriter instead: `Workbook.rewriteAllExtensionFormulas` walks every sheet's carriers (`sheet_edit.nextXmFormula`, matched by leaf name like `xm:sqref`) with `on_sheet` = the host sheet and `target_sheet` = the edited sheet, under every row / col / sheet-rename / sheet-delete / table-column-rename edit, splicing bodies in place through the DV/CF path (decode-in, re-escape-out, no host). A sparkline on `Report` reading `Data!A2:A5` therefore shifts when `Data` gains a row and holds when `Report` does — the sheet-name context `sheet_edit.zig` lacked. Failure posture is the opposite of `xm:sqref`'s and deliberately so: the contract is all-or-nothing, so `Workbook.preflightExtensionFormulas` scans every sheet before an edit's first mutation and refuses the whole edit (`MalformedExtensionXml` → `RowEditUnsafeForSheet` / `ColEditUnsafeForSheet` at the Editor) when a carrier has no `</xm:f>` or holds markup — never a stale carrier beside a shifted `xm:sqref`. A source range deleted outright collapses to `Data!#REF!`, the same convention as a cell formula. |
| ~~**autoFilter**~~ | ~~`recordRowEdit` / `recordColEdit` per-sheet `<autoFilter` scan~~ | ~~`RowEditUnsafeForSheet` / `ColEditUnsafeForSheet`~~ | ✅ **Lifted 2026-05-14**. `pkg/sheet_edit.zig` now shifts the row/col halves of `<autoFilter ref="…">` during the byte transform, drops the entire element on full-range collapse (delete that wipes the only row or only column the range covers), and walks `<filterColumn colId="N">` children for col edits — `colId` is a 0-based offset within the autoFilter range, so survivors rebase to `new_abs - new_tl_col` and the filterColumn at the deleted column is dropped entirely. Editor's row+col refusal guards no longer scan for `<autoFilter`. ~~Caveat: nested `<sortState ref="…">` carries its own range that isn't yet rewritten.~~ **Closed 2026-07-26 (iter-sv-1)** — `<sortState ref>` and `<sortCondition ref>` are rewritten in every context (sheet-bare, autoFilter-nested, table-nested), from one implementation in `sheet_edit.zig` that `table_edit.zig` delegates to. |
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

**Status (2026-05-14)**: `anySheetCrossSheetCarrier` is gone — the
five rewriters shipped through B2 iter-er-5 (formula, hyperlink, DV,
CF, defined-names). `RowEditUnsafeForSheet` / `ColEditUnsafeForSheet`
remain solely as the umbrella errors for the structured-table
pre-flight (`pkg/editor.zig` remaps `TableCollapseUnsafe` /
`TableHeaderRowDeleteUnsafe` / `MalformedTableXml` /
`TableCoordinateOverflow` into them). The drawing / pane /
autoFilter axes lift via dedicated typed errors and no longer
surface the umbrella names.

## The other failure mode: silently unhandled coordinates

Everything above is about **refusals** — cases where the Editor
correctly declines an edit it cannot perform safely. A refusal is
loud, and the user can act on it.

There is a second, worse class: a coordinate-bearing attribute that
is neither rewritten nor refused, so the edit "succeeds" and leaves
the attribute pointing at the wrong cells. Nothing surfaces. This is
the exact outcome the north star exists to prevent — *"every row/col
edit either rewrites all coordinate-bearing parts correctly or
refuses with a typed error."*

Four were found and closed in iter-sv-1 (2026-07-26):

| Element | What went stale | Reached via |
|---|---|---|
| `<sheetView topLeftCell>` | the view's scroll anchor | every scrolled Excel-authored sheet |
| `<selection activeCell/sqref>` | the saved selection, incl. multi-range `sqref` lists | essentially every Excel-authored sheet |
| `<sortState ref>` | the sorted range, sheet-bare and autoFilter-nested | third-party files |
| `<sortCondition ref>` | the sort key range | **also inside `<table>`** — already-shipped code |

The last one is the notable finding: `<sortCondition>` was unrewritten
even on the structured-tables path that shipped in #111, whose own
test asserted the parent `<sortState>` shifted but never checked the
child. A test that pins the parent and ignores the child is how this
class survives review.

None of these were refused, and zlsx's own writer emits none of them —
which is why they stayed invisible. They only bite on
**load-modify-save of third-party files**, which is the product's
primary use case.

**Method note for future audits.** The refusal list is derived from
what the Editor *scans for*. This class is invisible to that method by
construction: there is no guard to enumerate. Finding these needed the
inverse question — for each coordinate-bearing element in
CT_Worksheet, does a rewriter dispatch on it? The remaining unaudited
surface by that question was `<extLst>` blocks (`x14:`/`x15:`
extensions), which passed through verbatim everywhere — **closed
2026-07-29 (#140)**: `xm:sqref` shifts, `<xm:f>` refused until the
S2 lift (2026-08-26) routed it through the formula rewriter (see the
axes table above). The same sweep closed the other guard-less class:
pivots, which this file mislabeled "refused at consumer level" when
no refusal existed anywhere (**#139**).

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
