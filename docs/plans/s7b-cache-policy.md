# S7b — Source-row edits and the pivot cache policy

> **What this decides.** S7a moved the one coordinate a pivot's *host*
> sheet owns (`pivotTableDefinition/location@ref`). S7b is the sheet a
> pivot *reads from*: a row inserted or deleted there moves the range
> the cache names (`worksheetSource@ref`) — a coordinate rewrite — and
> changes the data the cache holds a snapshot of — a policy question.
> This document lists every source spelling the corpus and the
> fixtures carry, says what each needs rewritten under a row edit on
> the source sheet, sets out the two cache policies with what Excel
> shows for each, and recommends one. The owner's answer is the gate
> (`goal_sigmoid.md` §5, row S7b); nothing in the row is built before
> it.

_Written 2026-08-28 against `main` at `74641af` (S7a merged). Row S7c
(source **columns**) is out of scope: a column edit changes the cache's
field schema, not its range._

---

## Contents

1. [Why a source edit is not a host edit](#1-why-a-source-edit-is-not-a-host-edit)
2. [Every spelling, and what a row edit must rewrite](#2-every-spelling-and-what-a-row-edit-must-rewrite)
3. [What the cache holds](#3-what-the-cache-holds)
4. [The two policies](#4-the-two-policies)
5. [Recommendation](#5-recommendation)
6. [What S7b builds under the recommendation](#6-what-s7b-builds-under-the-recommendation)
7. [The gate](#7-the-gate)

---

## 1. Why a source edit is not a host edit

A pivot is three parts and one direction of edges:

```
xl/workbook.xml <pivotCaches>  ──r:id──▶  xl/pivotCache/pivotCacheDefinitionN.xml
                                              │  <cacheSource><worksheetSource sheet= ref= | name= | r:id=>
                                              │  <cacheFields> … <sharedItems>          ── the value inventory
                                              └──r:id──▶  pivotCacheRecordsN.xml         ── the row snapshot
xl/worksheets/sheetK.xml  ──rels──▶  xl/pivotTables/pivotTableM.xml  ──rels──▶  pivotCacheDefinitionN.xml
                                       <location ref=>  (S7a)
```

The host sheet's relationships name the pivot part, so an edit on the
host finds its pivot by walking *its own* rels — that is what `#139`'s
guard and S7a's lift do. The source sheet has no relationship to
anything: the edge runs the other way, from the cache definition's
`worksheetSource` to the sheet by *name*. Nothing on the edited sheet
says it is a source. The only way to know is to read the graph and
resolve every source (`Workbook.pivotTables`, S6), which is why the
S6 audit found the guard blind to it and why the `S6 audit` tests in
`pkg/editor.zig` pin that blindness today.

Two things then differ from S7a:

| | Host edit (S7a) | Source edit (S7b) |
|---|---|---|
| What names the sheet | the sheet's rels (`pivotTable` type) | `worksheetSource@sheet`, a table's host, a defined name's body — resolved through the engine's symbol table |
| What the coordinate describes | where the pivot is *drawn* | where the data is *read from* |
| What else describes the same thing | nothing — `location@ref` is the one absolute coordinate | the cache's records and item inventories, a **snapshot** of that range at `refreshedDate` |
| Excel on the same edit | moves the rectangle, refreshes nothing | moves the reference, refreshes nothing — the pivot keeps showing the snapshot until *Refresh* |

The first difference is the guard's reverse edge (§6). The second is
the policy (§4): after the reference moves, the snapshot describes a
range that no longer has the same rows.

---

## 2. Every spelling, and what a row edit must rewrite

`CT_WorksheetSource` (ECMA-376 §18.10.1.99) carries four optional
attributes — `ref`, `name`, `sheet`, `r:id` — and a consolidation
`<rangeSet>` carries the same four. The reader (`pkg/typed_parts/
pivot_xml.zig`) parses both into one `WorksheetSource` and exports
`ref_span`, the byte span of the `ref` value, for S7b to splice at;
the resolver (`pkg/pivots.zig::Resolver.resolve`) tries `r:id`, then
`sheet`, then `name` (defined names first, then tables — the two share
a namespace in Excel). The table below is every spelling seen in the
corpus, the fixture writer (`pivots.fixture.SourceKind`) and the
parser's unit tests, plus the spec-permitted shapes no producer has
shown us.

The corpus carries **one** pivot workbook, `tests/corpus/
openxlsx_loadExample.xlsx` — two caches, both table-named:

| Cache | Source | Read by | Host | `refreshOnLoad` | `recordCount` |
|---|---|---|---|---|---|
| `pivotCacheDefinition1.xml` | `<worksheetSource name="Table2"/>` — a table on `IrisSample` | `PivotTable1` | `IrisSample` (host **and** source) | absent | 50 |
| `pivotCacheDefinition2.xml` | `<worksheetSource name="Table3"/>` — a table on `mtcars` | `PivotTable3` (part `pivotTable2.xml`) | `mtCars Pivot` (source-only sheet: `mtcars`) | absent | 29 |

Neither carries `ref`, so **the corpus never exercises the rewrite S7b
owns**. The synthetic fixture does.

### 2.1 The inventory

"Resolves via" is `pivots.ResolvedVia`; "today" is what a row edit on
the source sheet leaves behind on `main` at `74641af`, as the named
test proves.

| # | Spelling | Seen in | Resolves via | A row insert / delete on the source sheet must rewrite | Today |
|---|---|---|---|---|---|
| 1 | `<worksheetSource sheet="Data" ref="A1:C4"/>` | fixture `.sheet_ref`; parser tests | `.sheet_attr` | **`ref`**, by range semantics (§2.2) — the one rewrite S7b owns, spliced at `WorksheetSource.ref_span`. The `sheet` attribute is a name, never a coordinate; it is decoded for resolution and left alone. | **Stale.** Admitted, `ref="A1:C4"` after the data moved to `A1:C5` — `S6 audit: a sheet+ref source sheet is admitted, and worksheetSource@ref goes stale`; the contract is stated by `S7b (failing-first): row edits on a sheet+ref source sheet move worksheetSource@ref like a range` (skipped until the row lands). |
| 2 | `<worksheetSource name="SalesTbl"/>` — a table | corpus (both caches); fixture `.table_name` | `.table` | **Nothing in the cache.** The range is spelled through the table part, and `table_edit.applyEditToTable` already moves `table@ref` + `autoFilter@ref` on every row edit, grows on an insert inside, shrinks on a delete inside, and refuses a header-row delete (`TableHeaderRowDeleteUnsafe`) and a collapse (`TableCollapseUnsafe`). | **Valid.** `S6 audit: a table-named source stays valid because the table rewriter moves the table` (`table1.xml` → `A1:C5`); corpus `mtcars` insert → `table2.xml` `A1:K31`. Only the snapshot is stale (§3). |
| 3 | `<worksheetSource name="PivotSrc"/>` — a defined name, body `Data!$A$1:$C$4` | fixture `.defined_name`; parser tests | `.defined_name` | **Nothing in the cache.** The body is a formula carrier; `Workbook.rewriteAllDefinedNames` is the second of the five sweeps every row / col edit runs (`applySheetEditTransform`), so the body moves with the grid. **But the rewriter's endpoint policy differs from Excel's:** deleting *either* endpoint row spells `#REF!` at that endpoint (`A4:A5` minus row 5 → `A4:#REF!`, pinned in `src/formula/rewriter.zig`; the convention S2 carried), after which the source is `.unresolved` and Excel's *Refresh* fails — where Excel itself would have shrunk the name to `$A$1:$C$3`. So for this carrier S7b's refusal is on the **rewrite**, not only the rectangle: dry-run the name sweep for every body a cache reads through, and refuse the edit when the new body spells `#REF!` (§2.2) — that covers both endpoint rows of a static body and the deleted anchor of a dynamic one (`OFFSET(Report!$D$1,…)` minus row 1); interior deletes and every insert are already right. | **Valid under inserts and interior deletes**, pinned by `S7b analysis: a defined-name source follows the defined-name rewriter` (added with this document); **wrong under an endpoint delete** (the source goes `.unresolved`) — today admitted, S7b refuses it. |
| 4 | `<worksheetSource r:id="rIdExt" sheet="Sheet1" ref="A1:C4"/>` — another workbook | fixture `.external`; parser tests | `.external` | **Nothing.** `r:id` wins in the resolver: the `sheet` and `ref` describe a sheet of *another* file, so a local sheet named `Sheet1` is not a source and its edits do not touch this cache. | Correct by construction — `readsFromSheet` is false for every local sheet. |
| 5 | `<worksheetSource sheet="Nope" ref="A1:C4"/>` — a sheet the workbook lacks | fixture `.dangling` | `.unresolved` | **Nothing** — there is no sheet to edit. `.unresolved` is one arm for several provenances: a name whose body the resolver cannot bound — dynamic (`OFFSET(Report!$D$1,0,0,4,3)`, the `mayReadFromSheet` test in `pkg/pivots.zig`), 3D, a union, a bare range, or one that reaches its sheet only *through another name* (`PivotSrc = OFFSET(Anchor,…)`, `Anchor = Report!$D$1`) — and an `r:id` the resolver could not place (missing, internal, mistyped). A name-spelled one *does* read a local sheet: its body moves with the grid through the name sweep, but there is no rectangle to detect a content change by, and a deleted anchor row turns a body into `#REF!` — the rule for all of them is §7 Q4. | Nothing to do for a dangling name. S7a's *host* guard refuses when any source is unresolved (`mayReadFromSheet`); S7b's *source* guard cannot mirror that without refusing every row edit on every sheet of the workbook — the split is §7 Q4. |
| 6 | `<cacheSource type="consolidation"><consolidation><rangeSets><rangeSet sheet="Q1" ref="A1:B9"/><rangeSet name="Q2Data"/>…` | parser tests only — no corpus, **no fixture kind** | one `SourceResolution` per set (`range_set_resolutions`) | **Per set, as rows 1–5** — the same resolver, the same `r:id` → `sheet` → `name` precedence: a set with a placeable `r:id` is external (no-op), an unplaceable one is Q4 (i); only a locally resolved `sheet`+`ref` set splices at its own `ref_span`; a named set follows its carrier. One definition may need several splices — apply them in descending span order so earlier spans stay valid. | Untested end-to-end: `fixture.SourceKind` has no consolidation kind. S7b adds one. |
| 7 | `<x:worksheetSource …/>` — Strict (`purl.oclc.org`) prefix | parser tests | as 1–5 | As 1–5: the splice is at `ref_span`, which the parser fills whatever the prefix. | As 1–5. |

Spec-permitted, unseen from any producer — S7b handles them but does
not need a fixture per shape:

| Shape | Treatment |
|---|---|
| `sheet` + `name`, no `ref` | The resolver takes `sheet`; no `ref` to splice; the name's own carrier (table part or defined-name body) moves, and its bounds come from that carrier. |
| `sheet` alone — no `ref`, no `name` | The parser accepts it and the resolver returns `.sheet` with nothing to bound. Not a shape Excel writes. Row edits of that sheet **refuse** (Q4 iv): the cache claims the sheet and gives no rectangle to move or to judge a shift by. |
| `ref` alone — no `sheet` | `.unresolved` (no sheet is claimed). Admitted untouched (Q4 iv); nothing local is proven. |
| whole columns | A *direct* `ref` is `ST_Ref` — one or two complete cell references (MS-OI29500's note on the type), so Excel spells a whole-column source as the full-height rectangle `A1:C1048576`: a finite rectangle, and an insert inside it overflows → refuse (S7a's rule). Whole columns *do* arrive through a defined name (`Data!$A:$C`), which is why the bounds model (§6) has a whole-columns case: **no coordinate rewrite, but not a no-op** — every row edit of that sheet changes the range's content (a mark under the §5 predicate — the body is byte-identical, the content is not), and row 1 is the header, so an insert at row 1 (blanks it) and a delete at row 1 (promotes data to field names) refuse. A literal `ref="A:C"`, if a producer ever writes one, takes the same path; `edit.parseRect` rejects it today. |
| bottom edge at row 1 048 576 | An insert inside cannot grow: `PivotCoordinateOverflow`, refused — S7a's rule for the same overflow. |
| `ref="&#65;1:C4"` — character references in `ref` | `workbook_xml.decodeScalarAttr` before parsing (Codex #200 r3 REL-042); the splice writes a plain spelling. |
| `type="external"` / `"scenario"` | No `worksheetSource`; nothing to rewrite. |
| an unknown `type` | The parser keeps a `worksheetSource` / `rangeSet` child under *any* type (`parseCacheSource`); `Pivots` drops it today only because its `.unknown` arm is empty. Whether such a locator is authoritative is §7 Q5. An unknown type refuses host edits today (S7a). |

### 2.2 The range semantics for row 1 (and each `rangeSet`)

Excel treats a pivot's source as a range reference: the same rules a
cell formula's `A1:C4` follows under a row edit, with a refusal where
the table rewriter has one. For a source `r1:r2` (rows) and an edit at
row `i`:

| Edit | Condition | New span | Note |
|---|---|---|---|
| insert | `i ≤ r1` | `r1+1 : r2+1` | pure shift — the snapshot still matches the range |
| insert | `r1 < i ≤ r2` | `r1 : r2+1` | **content changed**: one blank row inside the range |
| insert | `i > r2` | unchanged | part byte-identical |
| delete | `i < r1` | `r1−1 : r2−1` | pure shift |
| delete | `i == r1` | — | **refuse** (`PivotEditUnsafe`): the header row feeds the field names; Excel's *Refresh* on a header-less range fails with *"The PivotTable field name is not valid"* — the table rewriter refuses the same delete for a headered table. **Exception, table-named source with `headerRowCount="0"`:** the field names come from `<tableColumns>`, not the top row, and `table_edit.checkEditSafe` admits that delete — so it is a *content change* (`r1 : r2−1`, marked), not a refusal |
| delete | `r1 < i ≤ r2` | `r1 : r2−1` | **content changed**: one record fewer |
| delete | `r1 == r2 == i` | — | **refuse**: collapse |
| delete | `i > r2` | unchanged | byte-identical |
| any | **name-spelled (non-table) carrier**, dry-run body spells `#REF!` | — | **refuse**: the formula rewriter spells `#REF!` at a deleted endpoint (`i == r1` or `i == r2` of a static body) and at a deleted anchor (a dynamic body), which would leave the source unresolved where Excel shrinks it — the refusal is on the *rewritten body*, so it needs no rectangle |

This is `table_edit.shiftTableBounds` + `checkEditSafe` on the row
axis, applied to a rectangle the cache names instead of a table part.
It differs from S7a's `shiftRect` in one place: an edit *inside* the
rectangle is admitted (the range grows or shrinks) where the host
rectangle refuses it — a source range is data, a drawn pivot is not.

> Refusals are decided on the **resolved rectangle**, whatever spelled
> it — with the table carrier's own header knowledge (`headerRowCount`)
> deciding which row is the header; rewrites are done **per carrier**. A header-row delete on a
> defined-name source must refuse for the same reason a `ref` one does,
> even though the rewrite of the name's body is another sweep's.
>
> The resolver does not yet *keep* that rectangle: `SourceResolution.
> LocalSheet` is sheet index, name, part and `via` — the table part's
> `ref` and the name's body are read for the sheet and discarded. Typed
> bounds on the resolution (direct `ref`, `table@ref`, a static name's
> range; `null` for a dynamic body) are S7b's first piece (§6).

---

## 3. What the cache holds

Everything below is a snapshot of the source range at `refreshedDate`.
None of it is a coordinate.

| Part | Element | What it snapshots |
|---|---|---|
| `pivotCacheDefinitionN.xml` | `recordCount`, `refreshedDate`, `refreshedBy` | how many rows, when |
| | `refreshOnLoad` (default `0`) | Excel's *PivotTable Options → Data → "Refresh data when opening the file"* checkbox, one per cache |
| | `saveData` (default `1`) | whether the records are *saved with the file* — an independent axis from whether a records part is present (`r:id`, which is what the reader goes by; the fixture's orphan cache writes `saveData="0"` with a records part) and from any refresh policy |
| | `invalid` (default `0`) | "the cache needs to be refreshed" (ECMA-376 §18.10.1.67) — a *state* flag, distinct from the `refreshOnLoad` *option*; what Excel does with it on open is oracle-pending (§4 A2) |
| | `enableRefresh` (default `1`) | whether the user may refresh at all; when `0`, Excel ignores refresh-on-open (`PivotCache.EnableRefresh` documentation) |
| | `cacheFields/cacheField/sharedItems` | the value inventory per field: the distinct strings / numbers / dates, `count`, `minValue` / `maxValue`, `containsBlank`, `containsNumber`, … |
| `pivotCacheRecordsN.xml` | one `<r>` per data row, in source order | each cell as an index into `sharedItems` (`<x v=>`) or inline (`<n>`, `<s>`, `<d>`, `<b>`, `<e>`, `<m>` = blank) |
| `pivotTableM.xml` (every table on the cache) | `pivotFields/items`, `rowItems`, `colItems`, `pageFields`, filters, `pivotArea` | item lists and the rendered layout, each an index into the same inventory |
| the host sheet (`sheetK.xml`) | ordinary `<c>` cells inside `location@ref` | the rendered **numbers** — the corpus' `mtCars Pivot` holds cache 2's totals as plain values in `A2:D5`; the pivot part carries layout and indices only |

A row insert inside the range adds a row the snapshot lacks; a delete
removes one it has; a value that vanishes from the range may still be
in an inventory, and one that appears is not. Excel tolerates all of
that — it is the state every workbook is in after any cell edit, and
Excel never refreshes a pivot on save. zlsx is in the same posture
today: `setCell` on a source cell leaves the snapshot as saved, and the
S6 table-named test says as much ("only the cached records are stale,
as after any cell edit").

So the row edit is *not* a new class of staleness. It is the existing
class, plus a moved reference. What is new is only that zlsx now knows,
at the moment of the edit, that a source range changed — which is what
makes a refresh marker possible.

---

## 4. The two policies

### Option A — move the reference, leave the snapshot

Rewrite `worksheetSource@ref` (and each `rangeSet@ref`) by §2.2; touch
nothing else in the cache. Two sub-choices on the refresh marker:

- **A0 — never mark.** The definition is byte-identical except `ref`.
  This is what Excel itself writes after the same edit.
- **A1 — mark on content change.** Set `refreshOnLoad="1"` on a cache
  whose rectangle an edit changed (the two *content changed* rows of
  §2.2); a pure shift leaves it as written.
- **A2 — flag the state, not the option.** Set `invalid="1"` on the
  same condition and leave `refreshOnLoad` as the user set it. The
  spec's meaning is exactly "needs refreshing"; whether Excel *acts* on
  it at open (refreshes, then clears it) or merely shows the stale
  layout until the user refreshes is not known here — Excel writes the
  flag itself after a source change it could not refresh, which
  suggests the latter. **Oracle-gated**: if Excel refreshes on `invalid`
  and clears it, A2 dominates A1 (no persistent user option flipped);
  if not, A2 is A0 with a hint.

Excel-visible consequences:

| | A0 | A1 |
|---|---|---|
| Open, no repair prompt | yes — the same bytes Excel writes | yes — `refreshOnLoad` is Excel's own attribute |
| The pivot on open | the **snapshot**: old rows, old totals, the inserted row absent, the deleted row present — until the user clicks *Refresh* | **refreshed** on open from the moved range: the inserted blank row appears as a `(blank)` item (counted as 0 in sums), the deleted row is gone. No security prompt (that is for external connections) and no repair prompt — but the refresh itself can fail *visibly*: a pivot whose refreshed layout would overlap other content gets Excel's overlap error, and a host sheet protected without pivot use enabled cannot refresh |
| Best-effort cases | — | inert when the cache has `enableRefresh="0"` (Excel ignores refresh-on-open then; S7b leaves that attribute alone rather than override a user choice) and when Excel is opened programmatically (`Workbooks.Open` skips the automatic refresh) |
| A shared cache | — | one cache can feed several pivot tables (`PivotCache.consumer_count`); the marker refreshes **every** consumer, on every host sheet |
| *PivotTable Options → Data → "Refresh data when opening the file"* | unchanged | **checked**, and stays checked — Excel persists the option; the user can untick it |
| Save prompt on close | none | yes: the refresh dirties the workbook |
| Every later open of that file | as before | refreshes again (the option is persistent) |
| LibreOffice Calc | believed to rebuild the DataPilot from the range on import rather than from the records snapshot — under either policy it would show the current data; **oracle to confirm** | same |

Cost: the S7a-style splice (`ref_span`, reverse order over rangeSets),
the reverse-edge guard (§6), and for A1 one attribute write on the
root element. Nothing the snapshot invariants depend on is touched, so
nothing can be made inconsistent that Excel does not already tolerate.

### Option B — move the reference and rewrite the snapshot

Everything in A, plus edit the snapshot so that it *is* the new range:
insert an `<r>` of `<m/>` per inserted blank row at the right ordinal,
delete the `<r>` at a deleted ordinal, keep `recordCount` and
`<pivotCacheRecords count>` true, add `containsBlank="1"` and a blank
item to every field's `sharedItems`, and recompute `minValue` /
`maxValue`. Dropping an inventory item no surviving record references
is *optional* — Excel keeps such items up to `missingItemsLimit`
(`PivotCache.MissingItemsLimit`) — but if B drops one it renumbers a
**typed index graph**, not one index space: records `<x v>` → shared items; `pivotField/items/item@x` →
shared items; `rowItems` / `colItems` `<x v>` → *pivot-field item
positions* (not shared items); `pivotArea/reference` indices → whatever
the reference's `field` says (a data-field selector under field
`4294967294` indexes data fields). The corpus shows the layers apart:
cache 2's `cyl` inventory is `[6, 4, 8]` while its pivot field's items
read `x="1" x="0" x="2"` — the sorted view over the unsorted
inventory — and the row layout indexes *those* positions. A blanket
renumber of every `<x v>` corrupts layout and chart selections.

Three facts about B:

1. **It needs a precondition it cannot check.** Mapping the deleted
   *source row* to a *record ordinal* assumes the snapshot is current —
   that record `k` is row `r1+k`. Any workbook edited since its last
   refresh (by Excel without *Refresh*, or by zlsx `setCell`) breaks
   that, silently: the wrong record is removed. `refreshedDate` says
   when the snapshot was taken, not whether the cells changed since.
2. **Showing the change is the engine, not the cache.** The numbers a
   reader sees are ordinary cells on the host sheet (§3): a cache and
   pivot part rewritten to perfection still display the old totals
   until a refresh recomputes them. So B either rebuilds every
   consumer's output cells too — the cache builder plus the layout
   (`rowItems` / `colItems`) plus the aggregation, which is S8's engine
   — or it is A with a heavier, riskier snapshot that still needs the
   refresh A needs.
3. **It has no Excel-visible upside without the engine.** Cache-only
   B — the B of Q1 and the table below — displays the old totals until
   *Refresh* exactly as A does, because the totals are host cells (§3);
   what it adds is a rewritten snapshot that may be wrong (B.1). And
   if any inventory it wrote is inconsistent, Excel's *"We found a
   problem with some content"* repair drops the pivot — the corruption
   class the ladder exists to avoid.

Cost: an S8-sized cache builder with an oracle per shape, for a result
Excel itself does not produce on this edit.

### Side by side

| | A0 | A1 | B |
|---|---|---|---|
| Parts rewritten | cache definition (`ref`) | cache definition (`ref`, `refreshOnLoad`) | + records, + every table on the cache (+ every host sheet's output cells, if it is to show anything — the engine) |
| Bytes preserved | all but `ref` | all but `ref` and one root attribute | no |
| Assumes the snapshot is current | no | no | **yes, unverifiably** |
| Excel shows the change | after *Refresh* | on open | after *Refresh* — the host cells are untouched; on open only with the aggregation engine |
| New repair-prompt risk | none | none | real |
| Matches what Excel writes after the same edit | exactly | plus one option Excel exposes in its UI | no |
| Size | one PR | one PR | S8-sized, several PRs |

---

## 5. Recommendation

**Option A, sub-choice A1** — move the reference; mark
`refreshOnLoad="1"` on a cache whose rectangle the edit changed; leave
the snapshot as saved.

Why A over B: B's correctness rests on a fact no reader can verify
(§4 B.1), and its consistent form is the pivot engine, which belongs
to S8 where "the snapshot is the source" holds by construction (a
fresh workbook builds its own cache). Excel does not rewrite records on
this edit either; zlsx should not be the first.

Why A1 over A0: zlsx is headless. Nobody clicks *Refresh* after a
scripted edit, so under A0 the pivot the next reader opens shows the
snapshot — a weaker form of the silent staleness `#139` closed. The
marker is Excel's own attribute, costs one byte-level attribute write,
produces no prompt, and is visible where the user can turn it off
(§4). It is **best-effort**, not a guarantee: inert under
`enableRefresh="0"` (which S7b does not override) and under a
programmatic open, and it refreshes every consumer of a shared cache.
The marker has **one predicate** wherever it appears — *the edit may
have changed the source's content, and is not a proven pure shift*:

- a source with a **finite rectangle** (a direct `ref`, a table, a
  static name body) marks when the edited row is inside it (`r1 < i ≤
  r2`, either kind); an edit above or below it is a proven shift or a
  no-op and does not mark;
- a source with **no finite rectangle** — whole columns, or an
  unbounded name body (Q4) — marks on **any** row edit of a sheet its
  dependency closure references (the `sheet` qualifiers and table
  hosts in the bodies, transitively), because no shift can be proven
  for it. *Not* "a body was rewritten": `Data!$A:$C` is byte-identical
  under every row edit, and `OFFSET(Report!$A:$C,…)` or a table-backed
  body likewise — the content changed, the body did not.

The same predicate serves Q3 for cell writes: inside a finite
rectangle, or on a referenced sheet of an unbounded source. A pure
shift of a bounded source does not mark, which keeps it byte-faithful
and answers the S7a gate's parked second question the same way: a
moved host rectangle changes no data, so it does not mark.

Why A1 over A2 *today*: A2 is the cleaner signal — a state flag, no
persistent option flipped — but its effect at open is unknown here,
and a marker Excel ignores is A0. The oracle (§6) settles it before
the build PR: **if Excel refreshes on `invalid="1"` and clears it, ship
A2 instead of A1**; the code difference is which attribute the same
write sets.

One consequence to name: A1 makes a row edit mark a cache that a
`setCell` inside the same rectangle does not. That asymmetry is
today's, not new — but the clean rule is one rule. **Sub-question
(§7 Q3):** extend the marker to cell writes inside a resolved source
rectangle as a save-time pass (the editor knows the dirtied sheets;
one graph read at save, one attribute per affected cache). Recommended
yes, as S7b's second PR, so the row leaves one rule behind.

---

## 6. What S7b builds under the recommendation

For the estimate to be honest, the shape of the row:

| Piece | Where | What |
|---|---|---|
| The reverse edge | `Workbook.preflightPivotEditsForSheet` / `applyPivotEditsForSheet` | Today the graph is read only when the edited sheet's rels name a pivot part. S7b reads it whenever the workbook carries a cache at all (`<pivotCaches>` in `workbook.xml` or a `pivotCacheDefinition` relationship — a cheap string gate before the full walk) and selects every cache that **depends on** the edited sheet, `worksheetSource` and `rangeSet` alike: a `.sheet` resolution to it, *and* an `.unresolved` one whose retained provenance names it — the spelling's `sheet` attribute (an unplaceable `r:id`, a locator-only spelling) or a sheet in its name closure (Q4). `resolvesTo` accepts only `.sheet` today, so the selection is new, not `readsFromSheet`. |
| Typed bounds + provenance | `pkg/pivots.zig::SourceResolution` | On `.sheet`: `bounds: Rect \| whole_columns` — parsed from a direct `ref`, from `table@ref` for a table-named source, from a static name's body (`Sheet!$A$1:$C$4`, or `Sheet!$A:$C` → whole columns); absent for a locator-only spelling. On `.unresolved`: a payload instead of a bare tag — *why* (dangling sheet, dangling name, unbounded body, unplaceable `r:id`, sheetless `ref`) plus the sheets and tables the name closure references, which is what Q4 and the §5 predicate consume. Today the resolver reads all of this and keeps only the sheet, so every rule below stands on this piece. |
| Refusal by rectangle, and by rewrite | `pkg/pivots.zig::edit` + a name-sweep dry-run | Header-row delete, collapse, overflow on the *resolved* rectangle, whatever spelled it; whole-column sources refuse row-1 edits. For a name-spelled (non-table) carrier, dry-run `rewriteAllDefinedNames`' body computation for the names caches read through and refuse when a new body spells `#REF!` (§2.2). Folded to `Row`/`ColEditUnsafeForSheet` by `Editor` like S7a's. |
| Rewrite by carrier | `pkg/pivots.zig::edit.applyToCacheDefinition` (new) | `sheet`+`ref` → splice at `ref_span` (descending span order over rangeSets); named → no-op (the carrier's own sweep moves it); external / dangling → no-op; a locator under an unknown `type` → by its spelling (Q5). All-or-nothing across the sheet's caches via `PartStore.replaceParts`, as S7a. |
| The marker | same | `refreshOnLoad="1"` (A1) or `invalid="1"` (A2) on the root under the one predicate of §5 (edited row inside a finite rectangle; any row edit of a referenced sheet for a source with none). This is an **upsert**, not a replace: both corpus definitions omit the attribute, and the shared attribute writer (`writeWithReplacedAttrs`) substitutes values it encounters and has no insertion path — S7b adds a root-attribute upsert that replaces a present value or inserts `name="1"` before the root's `>` / `/>`, tested for both the present and the absent case. |
| Host **and** source | composition | `IrisSample` in the corpus: S7a moves `location@ref`, S7b moves the table (already) and marks the cache. The refusal narrows from "any row" to "inside the pivot's footprint" — the `S7a: a host that is also a source still refuses (S7b's case)` test flips. |
| Fixture | `pivots.fixture.SourceKind` | a `.consolidation` kind (two `rangeSet`s, one `sheet`+`ref`, one named). |
| Tests that flip | `pkg/editor.zig` | `S6 audit: a sheet+ref source sheet is admitted, and worksheetSource@ref goes stale` (assert the moved `ref`), `S7b (failing-first)` (drop its skip guard), the S7a host-and-source test, the corpus test (`IrisSample` rows outside the footprint admitted). |
| Oracle | `scripts/oracle/` | Excel: (1) a synthetic `sheet`+`ref` workbook after a zlsx insert inside the source opens without a repair prompt and *Refresh* shows the blank row; (2) with `refreshOnLoad`, the blank row shows on open and the Options checkbox is on; (3) with `invalid="1"` alone — does Excel refresh on open and clear the flag (A2 vs A1); (4) a cache shared by two tables refreshes both; (5) a refresh whose layout would overlap a cell below the pivot — what Excel shows; (6) header-row and name-endpoint deletes are refused by zlsx, so nothing to open. LibreOffice: opens (1)–(3), shows the current data. |
| Untouched | — | `location@ref` (S7a), the column axis (S7c: `cacheFields`, `fld=` ordinals, items), decompression limits, the parity rows. |

---

## 7. The gate

Asked of the owner at analyse-end, in order:

| # | Question | Recommended |
|---|---|---|
| Q1 | **Cache policy:** A (move the reference, leave the snapshot) or B (also rewrite the records, inventories and indices — which still shows the old totals until *Refresh*, since the numbers are host cells, unless it also rebuilds every consumer's output: S8's engine)? | **A** |
| Q2 | **Refresh marker:** A0 (never — byte-faithful to Excel's own save), A1 (`refreshOnLoad="1"` under the §5 predicate — a finite-rectangle source: the edited row is inside it; a source with no finite rectangle (whole columns, an unbounded name body): any row edit of a sheet its dependency closure references; a proven shift never marks; best-effort under `enableRefresh="0"` and programmatic opens; refreshes every consumer of a shared cache), or A2 (`invalid="1"` under the same predicate — oracle-gated)? | **A1, or A2 if oracle (3) shows Excel refreshes on `invalid` and clears it** |
| Q3 | **One rule:** extend the marker to `setCell` writes inside a resolved source rectangle, as a save-time pass in S7b's second PR? | **yes** |
| Q4 | **Unresolved sources** on a row edit of a sheet that hosts nothing. `.unresolved` has several provenances (§2.1 row 5): a dangling sheet or name; a name-spelled source whose body the resolver cannot bound (dynamic, 3D, union, bare range, or reaching its sheet through another name); an `r:id` it could not place, with a `sheet` + `ref` beside it. One rule, or refuse workbook-wide (S7a's "not proven local is not proven elsewhere", applied to every sheet — every row edit on every sheet refused for one `OFFSET` name)? | **one rule, by what the spelling proves**: (i) a spelling whose `sheet` attribute *is* the edited sheet but which the resolver could not place (an unplaceable `r:id`) → **refuse** — it may be this sheet, and the `ref` cannot be moved; (ii) a name-spelled source → **admit**; its **dependency closure** (the names its body references, transitively — the engine's symbol table has it) gives the sheets and tables it depends on: a row edit of any of those sheets *marks* under the §5 predicate (no finite rectangle → no provable shift), and the dry-run name sweep over the closure **refuses** when any body would spell `#REF!`; (iii) a dangling spelling → admit untouched; (iv) a locator-only spelling (§2.1, spec-permitted table) → `sheet` without `ref` or `name` refuses row edits of that sheet (it claims the sheet and gives no rectangle), `ref` without `sheet` admits untouched (no sheet is claimed). The resolver keeps the provenance so the rule can tell the cases apart. |
| Q5 | **Unknown cache source `type`** (not `worksheet` / `external` / `consolidation` / `scenario`). First: is a `worksheetSource` / `rangeSet` the definition carries under an unknown type *authoritative*? The parser keeps it; `Pivots` discards it (`.unknown => {}`). Then, for an unknown type with no usable local locator, on a row edit of a sheet that hosts nothing: admit, or refuse workbook-wide as S7a does for a *host* (`mayReadFromSheet` is true for every sheet when any cache's type is unknown)? | **authoritative** — resolve a carried locator and apply the normal carrier rules (S7b fills the `.unknown` arm); **admit** when there is none — S7a's refusal protects a drawn rectangle on the edited sheet, and a definition that names no sheet names not this one |

The answers are recorded in `goal_sigmoid.md` §5 (row S7b) and this
file is amended with them; the row's `What it does` cell then names
the chosen policy, and the refusal audit and surface matrix update when
the lift lands.
