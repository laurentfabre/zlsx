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
field schema, not its range. **Gate answered 2026-08-28 — see §8**: the
owner chose **full B with the engine**, not the recommended A._

---

## Contents

1. [Why a source edit is not a host edit](#1-why-a-source-edit-is-not-a-host-edit)
2. [Every spelling, and what a row edit must rewrite](#2-every-spelling-and-what-a-row-edit-must-rewrite)
3. [What the cache holds](#3-what-the-cache-holds)
4. [The two policies](#4-the-two-policies)
5. [Recommendation](#5-recommendation)
6. [What S7b builds under the recommendation](#6-what-s7b-builds-under-the-recommendation)
7. [The gate](#7-the-gate)
8. [The answers](#8-the-answers)

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
| 6 | `<cacheSource type="consolidation"><consolidation><rangeSets><rangeSet sheet="Q1" ref="A1:B9"/><rangeSet name="Q2Data"/>…` | parser tests; fixture `.consolidation` (since S7b-1) — no corpus | one `SourceResolution` per set (`range_set_resolutions`) | **Per set, as rows 1–5** — the same resolver, the same `r:id` → `sheet` → `name` precedence: a set with a placeable `r:id` is external (no-op), an unplaceable one is Q4 (i); only a locally resolved `sheet`+`ref` set splices at its own `ref_span`; a named set follows its carrier. One definition may need several splices — each at its own span, the part rebuilt from its raw bytes in span order (the spans are offsets into the unchanged source, so none moves under another). | Pinned end-to-end since S7b-2: `S7b: the consolidation fixture — the direct set is respelled, the named set follows its body, the host still moves` (`pkg/editor.zig`) and the `edit:` consolidation tests (`pkg/pivots.zig`). |
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
| Rewrite by carrier | `pkg/pivots.zig::edit.applyToCacheDefinition` (new) | `sheet`+`ref` → splice at `ref_span` (each rangeSet at its own span, the part rebuilt from its raw bytes in span order); named → no-op (the carrier's own sweep moves it); external / dangling → no-op; a locator under an unknown `type` → by its spelling (Q5). All-or-nothing across the sheet's caches via `PartStore.replaceParts`, as S7a. |
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

---

## 8. The answers

Asked and answered 2026-08-28 (PR #201), after five Codex framing
rounds (8 → 6 → 5 → 4 → 2 findings).

| # | Answer | What it means for the build |
|---|---|---|
| Q1 | **B — full B with the engine.** Not the recommended A, and not "B when provable": the owner chose the reading in which zlsx *performs the refresh* — rebuild the cache from the source cells (records, every `sharedItems` inventory, `recordCount`, `refreshedDate`), rebuild every consumer's items and layout (`pivotFields/items`, `rowItems` / `colItems`), and recompute every consumer's output cells on its host sheet (aggregation). | B.1's precondition disappears: nothing is mapped by ordinal, the cache is *rebuilt*, not patched. B.2 is satisfied by construction: the engine is the row. The estimate is no longer 2–3 weeks — S7b absorbs the cache builder and the aggregation engine S8 was to bring, and **S8 reuses S7b's engine** (the spine inverts: S7b → S8, not the other way). The refusal posture holds at the engine's edge: a pivot shape the engine does not evaluate (the S7c/S8 list — calculated fields, grouping, OLAP, consolidation, external sources, …) **refuses** the source edit rather than write a partial rebuild; that list is the row's oracle matrix, per the phase discipline. |
| Q2 | **A1, or A2 if the oracle proves it.** | With B the marker is a safety net, not the mechanism: Excel refreshes again at open and must land on what zlsx wrote. Set under the §5 predicate; the `invalid` oracle (§6, oracle 3) still runs before the build PR. |
| Q3 | **Yes — one rule, S7b's second PR.** | Cell writes inside a source rectangle (or on a referenced sheet of an unbounded source) mark at save time; under B they are also candidates for the same rebuild — the second PR decides whether a `setCell` triggers the engine or only the marker. |
| Q4 + Q5 | **The doc's one rule.** | Refuse narrowly (an unplaceable `r:id` or a sheet-only spelling naming the edited sheet; a dry-run body spelling `#REF!`); admit and mark otherwise; a locator carried under an unknown `type` is authoritative. Under B an unbounded source cannot be rebuilt (no rectangle to read) — it takes the marker path only. |

What the recommendation got wrong, for the record: it weighed B as
"patch the snapshot" and found the patch unsound; the owner's B is
"rebuild the snapshot", which is sound and simply expensive — the
engine the ladder had scheduled two rows later.

---

## 9. What has landed

The row ships in pieces, each behind the previous one's tests:

| Piece | PR | What |
|---|---|---|
| S7b-1 — typed bounds + provenance | #202 | `SourceResolution.LocalSheet.bounds`, `Unresolved{why, sheets}`, `Pivots.dependsOnSheet`, the consolidation fixture kind. |
| S7b-2 — the reference move | #203 | The reverse edge (§6 row 1): the graph is read whenever the workbook carries a cache — a `pivotCacheDefinition` relationship, or a main-namespace `<pivotCaches>` list (a listed entry whose relationship is absent or mistyped is a graph that refuses, from every sheet) — and every cache that depends on the edited sheet is selected. `pivots.edit.applyToCacheDefinition` (§6 row 4): each `sheet` + `ref` spelling — `worksheetSource@ref`, every consolidation `rangeSet@ref` — is respelled by the §2.2 semantics at its `ref_span`, the part rebuilt from its raw bytes in span order; a table-named source moves with its table, a defined-name source with its body. The refusals of §6 row 3 by rectangle (header-row delete, collapse, overflow, a column edit inside the range — S7c's schema) and by rewrite (`rewrittenDefinedNameBody`, the one computation behind `rewriteAllDefinedNames`, dry-run over every name a cache reads through — `SourceResolution.names` carries the closure — refusing a body that would newly spell `#REF!` — counted outside string and sheet-name quotes — and a body the sweep never rewrites (embedded markup) when its source depends on the edited sheet); the Q4 refusals (an unplaceable `r:id` or a `sheet`-only spelling naming the edited sheet); the Q5 arm (a locator under an unknown `type` is resolved). The S7a "host some cache *may read*" refusal is lifted — a host that is also a source moves both coordinates. A graph that cannot be read now refuses every sheet's structural edit, since the source edge cannot be read either. |
| S7b-3 — the marker | #204 | **A1**, `refreshOnLoad="1"`, upserted on the definition's root — the value replaced where the attribute is present (`0` / `false` → `1`; a root already spelling `1` / `true` is left alone), inserted before the root's `>` where absent, which is both corpus definitions (`pivots.edit.markerSplice`; the shared attribute writer has no insertion path) — in the same span-ordered, all-or-nothing rebuild as the `ref` move (`applyToCacheDefinition`), so the preflight dry-runs it and `replaceParts` installs it with the rest. The **one predicate** of §5, per source resolution (`edit.editChangesContent`): a finite rectangle — a direct `ref`, a table's `ref`, a static name body, whole rows — marks when the edited row is inside it: an insert strictly inside (`r1 < i ≤ r2`; at `r1` it is a shift), a delete anywhere in the span (`r1 ≤ i ≤ r2`; at `r1` it is the admitted headerless-table case — every other carrier refused it before asking); a proven shift above or a no-op below never marks, so the part stays byte-faithful to what Excel writes after the same edit, and a column edit never marks (inside refuses — S7c; outside shifts). Whole columns mark on every row edit (row 1 refuses) and no column edit. An unbounded name body (`Unresolved.unbounded_body`) marks on **any edit of a sheet its closure references, on either axis** — the doc's §5 names row edits, but a column edit on such a sheet is not a proven shift either, and admitting it unmarked would be the one edit of a source the marker does not see. Never marked: external, dangling, sheetless, an unplaceable `r:id` (its `sheet` refuses the edit instead, Q4 i), a host rectangle that moved (the S7a gate's parked second question, answered: no data changed). **Q3, one rule for cell writes:** `Workbook.applySavePlans` (reached by `Editor.save`, `Editor.saveToOwnedBuffer` and the direct `Workbook.save`) reads the graph once per save — only when a sheet has staged `setCell` deltas or appended rows and the workbook carries a cache, the sweep's string gate — and marks every cache a write lands inside (`edit.cellWriteChangesSource`: inside a rectangle, in a whole-column source's columns, in a whole-row source's rows, anywhere on a referenced sheet of an unbounded body, anywhere on a sheet a `sheet`-only spelling claims), a delta at its coordinates, an appended row where the emitter places it (after the sheet's highest row, one column per non-empty cell — a source Excel wrote taller than its data catches appends; one that ends at the data does not), through `edit.markForRefresh`. Writes outside every source leave every definition byte-identical; a definition already marked is byte-identical under any write; a graph that cannot be read marks nothing and the save proceeds — the marker is best-effort, the snapshot it would have flagged is the state the file was already in, and a save is not where an edit the editor admitted is refused (the row edit on that workbook refuses at edit time, S7b-2). **A2 stays oracle-pending:** `scripts/oracle/` is the formula-recalc harness (`regenerate.sh`, `record_excel_mac.sh`) and carries no pivot leg, so oracle 3 of §6 — does Excel refresh on `invalid="1"` and clear it? — needs the owner's Excel on a marked fixture (`pivots.fixture.write(.sheet_ref)` + a zlsx insert at row 2, once with the attribute swapped by hand); if it does, `edit.marker_attr` (and `edit.markerSet`, which reads the parsed flag) is the one place that changes. Tests: the `edit:` upsert test in `pkg/pivots.zig` (absent, present-and-off single-quoted with whitespace, present-and-on as `true`), the cell-write predicate test, every S7b editor test now asserting the marker's presence or absence (`expectMarked`), the four `S7b-3` editor tests (cell writes inside / outside / on the host, by source kind, appended rows, idempotence + the unreadable graph), and the pure-shift test pinning a part that differs from the original in `ref` alone. |
| S7b-4 — the engine, first slice: the cache | #205 | **The snapshot is rebuilt from the cells.** When the S7b-3 predicate says a row edit changed a source's content (`edit.Plan.changed`), `pkg/pivots.zig::engine` rebuilds the cache from the source rectangle — the records part whole, every field's `sharedItems` inventory (strings, numbers, blanks, the `contains*` flags, `minValue` / `maxValue`, `count`, `longText`), `recordCount`, `refreshedDate` (the edit's wall clock as an Excel serial under the workbook's date system; dropped when the workbook was built from a bare store and has no clock) — and `Workbook.applyPivotEditsForSheet` installs the rebuilt definition and records parts in the same `replaceParts` transaction as the `ref` move, the marker kept: until the consumers' slice lands, Excel's refresh at open lays the consumers out over a snapshot that is already the source. The rectangle is read as the sheet is *before* the edit (the sweep runs before the first mutation) with the edit applied to the rows (`engine.rowsAfterEdit`: a blank record at the inserted row, the deleted row dropped), through the typed sheet view, the shared strings and the styles; the header row must spell the cache's field names (a headerless table — `headerRowCount="0"`, `LocalSheet.header_rows` — has none to check). **Two invariants keep every consumer index-valid without touching the consumer parts:** an inventory keeps every item it had, in its order — an unreferenced item stays, as Excel keeps items up to `missingItemsLimit` — and a value it lacks is appended in first-appearance order (numbers matched by value, strings case-insensitively under the workbook's collation, spelled by their first occurrence); a field whose records were inline (`<n>`, a data-only numeric field) stays inline unless it now holds a string, an enumerated one stays enumerated. Supported: a `worksheet`-type source with a finite rectangle — a direct `sheet` + `ref`, a table (`headerRowCount` 0 or 1), a static name body — and plain database fields whose inventories hold `<s>` / `<n>` / `<m>` items. **Refused (`PivotShapeUnsupported` → `PivotEditUnsafe` → `Row`/`ColEditUnsafeForSheet`), nothing written:** a source with no finite rectangle — whole columns or rows, an unbounded name body, a consolidation `rangeSet` — or a locator under a source `type` the engine does not read, finite or not (it moves, §7 Q5, but is not rebuilt), which S7b-3 admitted and marked; a calculated or group field (`formula`, `databaseField="0"`, a `fieldGroup` child); an OLAP or calculated root element (`cacheHierarchies`, `kpis`, `tupleCache`, `calculatedItems`, `calculatedMembers`, `dimensions`, `measureGroups`, `maps`); a field without a `sharedItems`, an inventory that held dates, booleans or errors, or an item with children; a rectangle whose width is not the field count or whose header is not the field names (S7c's schema); a source with no data row; a records part carrying anything but records; and cells the oracle matrix has not covered — a `t="d"` date, a number under a date format (the numfmt grammar's `describesDate`; a format it refuses cannot be told from one), a boolean, an error, a formula without a cached value, an inline string. Round 1 of the Codex review of #205 added to the refusals: a cell whose `r` names another row than the one holding it (`MalformedSheetXml`), a cell or row without an `r` (the typed view cannot place it), a `t="str"` formula without a `<v>` (uncomputed, like its numeric twin — an empty `<v/>` is the empty string), and a foreign-namespace element where the schema names only the part's own — an inventory item, a field or root child, a record — which a part regenerated around it would drop; a namespace alias or rebinding of the main namespace below a part's root refuses the graph read itself (`MalformedPivotXml`), the one-prefix scanner's precondition, as does a root whose one main binding is not its own prefix. Round 2 added: a numeric cell under a style whose format the workbook cannot spell — a locale built-in the table does not list (27–36, 50–58, 81; dates among them), a custom id without its `<numFmt>`, an index past `cellXfs` — refuses like a format the grammar refuses, only a style with no `numFmtId` reading as General; and an explicitly closed empty item (`<s v="a"></s>`) is the self-closing one, kept as written. Round 3 added: two definitions naming one records part are a graph that refuses; a rectangle wider than the field schema, or past the slice's cell budget (`engine.max_rebuild_cells`, 16 Mi cells — a million rows of sixteen fields), refuses before any cell is read, and the blank rows of a sparse source share one row, so a read costs the sheet's cells, not the rectangle's area; an enumerated inventory's `contains*` flags, `containsInteger`, `longText` and extrema describe the items the element holds — retained and appended — while an inline field's describe its rows; a main-namespace binding first introduced below a part's root refuses the read; and a record holding anything but childless value elements under the part's prefix refuses. Round 4 added: a table's totals rows (`totalsRowCount`, carried through the resolution) are not records — a delete on one refuses, an insert at one appends a data row; a headerless table's `<tableColumn>` names must be the field names; character data or a comment between the children of a records part, a record or an inventory refuses (a regenerated body would lose it); a row or cell coordinate twice in the source refuses (`MalformedSheetXml`); and a name twice on one start tag refuses the part. Round 5 added: the totals rows are not read at all (a label or an uncached total there is nobody's source cell) and the table's counts are added with a check; an item carrying a prefixed attribute refuses (a declaration on the inventory element would not survive its rebuild); a headerless table's column names are read through the part scanner — one preflighted tree, direct children only — and only for a headerless table; a cell whose `s`, or a style whose `numFmtId`, is written but not a number is not General and refuses. Round 6 added: a cell without `s` wears style 0 — a date there refuses; only a workbook with no styles part has no style 0 to wear; a qualified attribute on `sharedItems`, on a record or on a value element refuses (a regenerated element would lose it); a cell reference is read by the strict grid parser (uppercase inside `XFD`, a row inside 1048576, no leading zero, no `$`); consolidation markup beside a worksheet locator refuses, sets or none; the Editor's prepared collection is checked at run time — owner, store mutation count, sheet, axis, index, kind — and refuses when it is not this edit's; a start tag with more than 256 attributes refuses the part. Round 7 added: a splice whose spans are reversed, overlapping or past the part is `MalformedPivotXml`, not an assertion; `refreshedDate` is dated under `workbookPr@date1904` read on its own — a calc policy the recalculator refuses (`fullPrecision="0"`) no longer drops it; only a workbook with no clock, or an epoch that does not read, goes undated. Round 8 added: a cell whose `t` names no type the view knows is not a number and refuses; a number, an SST index or a retained `<n>` spelt with character references (`1&#x2E;5`) reads by what it is; a table whose `headerRowCount` or `totalsRowCount` is written but not a number refuses the graph rather than take the default; and a `xmlnsfoo:` prefix is a prefix, not a declaration. Round 9 added: a number is read by the xsd:double grammar, not Zig's wider one (`0x1p0`, `1_0` refuse); `refreshedDateIso`, where a part spells it, is redated to the same instant or removed with `refreshedDate`; what follows the records root — a comment, a processing instruction — is kept as written; `date1904` is decoded before it is read, and two `workbookPr` leave the rebuild undated. Round 10 added: whether a field's records were inline or indexed is read from the records as written (an explicit `count="0"` alone says nothing; a field spelt both ways, or a record of the wrong arity, refuses); and the typed views decode a cell's `r`, `s` and `t`, a row's `r`, a style's `numFmtId`, a table's `ref` and counts before reading them, so a character reference there is the value it names. Round 11 added: records spelt inline beside an inventory that holds items refuse — the two are not one shape, and an inline rebuild would drop every item a consumer indexes. Round 12 added: a `SheetEditSpec` naming neither axis or both is `InvalidSheetEditSpec`, refused before anything is read. A padded close tag (`</sharedItems >`) bounds the inventory span through the scanner's own close position, not arithmetic on the name. The Editor's pre-flight collection (`Workbook.PreparedPivotEdits`, from `preflightPivotEditsForSheet`) is the one its sweep installs through `Workbook.applySheetEdit`: one rebuild per cache per edit, not one for the pre-flight and another for the sweep. A pure shift never reaches the engine: the part stays byte-faithful to what Excel writes after the same edit. Cell writes (`setCell`, appended rows) keep S7b-3's marker-only rule at save; the engine does not run for them yet. **Oracle, parked with the owner:** a rebuilt cache opens in Excel without a repair prompt and *Refresh* agrees with it — `scripts/oracle/` has no pivot leg, so the owner's Excel decides on the fixture (`pivots.fixture.write(.sheet_ref)` + a zlsx insert at row 2) and on the corpus workbook after the `IrisSample` / `mtcars` inserts the editor tests perform. Tests: the `engine:` tests in `pkg/pivots.zig` (the insert and delete rebuilds to the byte, the stale-snapshot cases, the all-blank field, no clock, the shape refusals, a Strict-prefixed part, the plan's rectangle per carrier, the Codex #205 round-1 pins — a padded close tag, namespace hygiene, the signed 32-bit `containsInteger` interval) and the `S7b-4` editor tests (the calculated-field refusal beside an admitted shift, the cell refusals before any mutation, computed formulas, the bare-store clock, the records `extLst`, a cache without a records part, one rebuild per Editor edit, the round-1 cell and namespace refusals), plus every S7b editor test that now asserts the rebuilt counts and inventories — the corpus' two caches included. |
| S7b-5 — the engine, second slice: the consumers | this PR | **S7b-5 (the engine's second slice — the consumers, 2026-08-30):** every consumer of a rebuilt cache is laid out again in the same edit — `pkg/pivots.zig::engine.layout`: the row field's `<items>` (the written order kept, an inventory item it lacked appended after it under `sortType="manual"`, the schema's default, or every item re-sorted under `ascending` — numbers by value, text under the workbook's collation, the blank last; an item no record references is marked `m="1"`), `<rowItems>` (one `<i>` per item with a record, then the grand total), `location@ref` (the rectangle the rows now fill; `<colItems>` is left as written — no row edit changes the values axis), and the cells of the host rectangle: the header row (the row-labels caption the host already spells, else `Row Labels`; each `dataField@name`), one row per shown item (its label — a number as inventoried, text, or the `(blank)` caption — and one aggregate per data field), the grand total (`grandTotalCaption`, else the host's caption, else `Grand Total`). Aggregates: sum, count, countNums, average, min, max, product, in record order; a group with nothing to aggregate is an empty cell; text under a numeric aggregate is skipped, as `SUM` skips it. The grand total is what Excel computes, which the corpus fixes to the bit: a sum is the fold of the subtotals in item order, an average one running sum over every record divided by their count — no fold of the subtotals gives it — count, min and max are order-blind, product follows the running pass. The cells are written last in the sweep, after the byte transform, in post-edit coordinates (`Workbook.applyPivotHostWrites`; strings through the shared-string table, extended as a save extends it): edit first, refresh second — Excel's order — so a pivot on the edited sheet moves with the edit and then grows into the row an insert opened below it. Styles carry by row kind (the header's, the last item row's, the grand total's, per column); a cell the old rectangle covered and the new one does not is cleared. The corpus is the oracle: both `openxlsx_loadExample.xlsx` pivots re-laid from their own records reproduce every cell Excel wrote — 45 cells, the numbers to the bit (`294.80000000000007` is the sum of three subtotals, not of fifty records; `20.210344827586205` is twenty-nine records summed once, not three subtotals added) — and their parts byte for byte. **Cell writes join the rebuild (§7 Q3):** a save with a staged `setCell` or appended row inside a finite-rectangle source rebuilds the cache from the sheet as it is about to be (the staged writes laid over the read — `Workbook.overlayStagedWrites`) and re-lays its consumers, the host cells written to the parts before the sheet phase re-emits the remaining deltas over them; a write inside a re-laid rectangle is the pivot's to overwrite, as Excel's refresh overwrites a typed-over cell. A save never refuses: a shape the slice does not lay out, a host with staged appends, a rectangle that would grow over a cell take the marker alone, as S7b-3 left them. The marker stays on every rebuilt cache — the safety net §8 Q2 describes: Excel's refresh at open lands on what zlsx wrote, and re-lays what the slice leaves as written (a slicer's or timeline's item list). Refused on the edit path, nothing written (`PivotShapeUnsupported` → `PivotEditUnsafe` → `Row`/`ColEditUnsafeForSheet`): a form the slice does not lay out — more than one row field, a real field on the columns axis, a page field, `dataOnRows`, tabular / outline form (`compact="0"`), `showAll="1"`, a hidden or otherwise attributed item, `descending` or an unknown `sortType`, a top-N filter, a custom subtotal set, a blank row after items, a dispersion (`stdDev`, `var`, …), a `showDataAs` other than `normal`, `<formats>`, `<conditionalFormats>`, `<filters>`, an OLAP hierarchy, a chart format on a field's items (one on the values axis alone is admitted), a row field the cache does not enumerate, a `<rowItems>` / `<colItems>` not of the form, a `location` whose counts are not the form's; a rectangle that would grow over a cell holding anything (Excel refuses the overlap too); a pivot part two sheets host. Taken on faith, oracle pending: an empty cell (not 0) for a group with no value; a manual-sort field appending its new item last rather than in sorted position; `m="1"` on a retained item. |
