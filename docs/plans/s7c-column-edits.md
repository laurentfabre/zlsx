# S7c — column edits inside a pivot source: the per-case lift/refuse matrix

_Written 2026-08-30 against `main` at `72b5307` (S7b done, S3a done). Row S7c
of `goal_sigmoid.md`: "Column edits touch cache-field schema, ordinals, items,
filters, calculated fields — lift only what the oracle proves; the remainder
stays a documented, tested refusal. A refusal outcome closes this row."_

## 0. Where the line sits today

A structural column edit (`insertColumn` / `deleteColumn`) against a sheet a
pivot cache reads from is decided by `pkg/pivots.zig::edit.shiftSourceRect`
(and `shiftSourceBounds` above it):

| Column edit, vs the source rectangle | Today | Why |
|---|---|---|
| Left of `tl_col` (insert at or left of it; delete strictly left) | **shift** — the `ref` / name body / table follows, byte-provable | S7b-2 |
| Strictly inside (insert `tl < i ≤ br`; delete `tl ≤ i ≤ br`) | **refuse** `PivotSourceEditUnsafe` → `ColEditUnsafeForSheet` | the rectangle's width *is* the cache's field schema |
| Right of `br_col` | **no-op** | outside the range |
| Whole-column source (`Data!$A:$C` name body) | same three arms on the column axis | S7b-1/2 |
| Whole-row source | **refuse** every column edit | every column is inside |
| Unbounded name body | **refuse** any content-changing edit | the engine cannot rebuild what it cannot bound (policy B, S7b-4) |

The host rectangle is S7a's and unchanged here: a column edit inside the
pivot's footprint refuses (`PivotLocationEditUnsafe`), outside it shifts.

S7c decides the **inside** arm per case. The decision inherits the S7b gate's
policy B verbatim: *zlsx performs the refresh — a shape the engine does not
evaluate refuses the edit.* Nothing here re-opens marking-instead-of-rebuild;
a marked-but-stale schema is not a state policy B ships.

## 1. What a column edit inside a source touches — the carrier inventory

A row edit changes *records*. A column edit changes the **field schema**, and
field identity is spelled twice: by **name** (the header row ↔
`cacheField@name`) and by **ordinal** (document position in `<cacheFields>`).
Every carrier of either, per part:

**The cache (`pivotCacheDefinition` + `pivotCacheRecords`):**

| # | Carrier | Keyed by | Engine status today |
|---|---|---|---|
| 1 | `<cacheFields count>` + the positional `<cacheField>` list | ordinal = position | parsed; count checked (`checkShape`); **no element spans exported** |
| 2 | `cacheField@name` | name = header-row text | checked against the header row on every rebuild |
| 3 | each record `<r>` — one value per field, positionally | ordinal | regenerated whole by `engine.rebuild` |
| 4 | `sharedItems` per field | item index, *per field* — removal of another field does not move it | rebuilt append-only |
| 5 | calculated field (`@formula`, `databaseField="0"`) — the formula names other fields **by name** | name | `checkShape` **refuses the shape** |
| 6 | group field (`fieldGroup@base`, `@par`) — **field ordinals** | ordinal | `checkShape` refuses (`has_other_children`) |
| 7 | OLAP roots (`cacheHierarchies`, `kpis`, …) | — | refused (`has_other_children`) |

**Each consumer (`pivotTableDefinition`):**

| # | Carrier | Keyed by | Slice status today |
|---|---|---|---|
| 8 | `<pivotFields count>` + positional `<pivotField>` list (must mirror `cacheFields` 1:1) | ordinal = position | parsed; **no element spans** |
| 9 | `rowFields` / `colFields` `<field x>` (signed; `-2` = the values axis) | ordinal | parsed; layout admits one row field + `x="-2"` columns |
| 10 | `pageFields/pageField@fld` | ordinal | shape-refused |
| 11 | `dataFields/dataField@fld`, `@baseField` (`@baseItem` is an item index) | ordinal | parsed; `fld` drives layout; `baseField` inert under `showDataAs="normal"` but written (the corpus writes `baseField="1"`) |
| 12 | `pivotArea@field` + `references/reference@field` — under `formats`, `conditionalFormats`, `filters` (`@fld`, `@iMeasureFld`), `autoSortScope`, `chartFormats` | ordinal | all shape-refused **except** a chartFormat proven to select the values axis alone (`field="4294967294"` = −2) |
| 13 | `rowItems` / `colItems` `<x v>` | item positions, not ordinals | re-laid by the engine |
| 14 | `extLst` (x14) — `x14:pivotField` rides *inside* a `pivotField`'s own `extLst` (positional with its parent); the corpus' root-level `x14:pivotTableDefinition hideValuesRow` carries no field key | ordinal via parent | rides with its element |

**The consumer's host sheet:**

| # | Carrier | Keyed by | Status |
|---|---|---|---|
| 15 | `<pivotSelection>` → optional `<pivotArea>` child whose `references` name ordinals | ordinal | `sheet_edit` rewrites only the four absolute coordinates; the corpus sheets carry **no** `pivotSelection` |

**The package:**

| # | Carrier | Keyed by | Status |
|---|---|---|---|
| 16 | slicer caches (`xl/slicerCaches/*`, x14) — `sourceName` is a **field name**, `<pivotTables><pivotTable name=…>` the attachment list; timeline caches (`xl/timelineCaches/*`, x15) likewise | name | **untyped beyond the attachment list** — S7b-5 leaves their item lists to Excel's refresh, which cannot repair a *deleted* source field (the slicer dangles, visibly); the corpus carries one, `Slicer_Species` on the iris pivot |
| 17 | pivot charts — the chart part names the pivot, not a field; per-series pivot links live in `chartFormats` (#12) | — | gated by #12 |
| 18 | `GETPIVOTDATA("field", …)` in formulas — a string literal | name | out of scope: C1 ruled Excel never rewrites string literals; the function degrades to `#REF!` at recalc exactly as it does after Excel's own refresh drops a field |

The inventory is the audit surface: a lift is complete when every row above is
either rewritten, proven unaffected, or part of the refusal predicate. Rows
5–7, 10, and most of 12 are already outside the engine's admitted shape, so
under the S7b-4/5 checks the *live* rewrite surface for an admitted workbook
is rows 1–3, 8–9, 11, 14 — plus presence gates for 15 and 16.

## 2. The case matrix

Cases split by edit kind and by whether the edited ordinal is **referenced**
— carried by any of rows 9–12 (equivalently: its `pivotField` spells `axis=`
or `dataField="1"`, cross-checked both directions, a disagreement refusing as
the slice already refuses count disagreements).

Excel's refresh behavior below is stated from training data except where the
corpus pins it; each claim is marked. The **Excel-open oracle stays parked**
(`scripts/oracle/` has no pivot leg — S7b-4's parked line covers S7c's
fixtures too when it runs).

| Case | Edit | Excel after the same edit + refresh | Claim status | Ruling options |
|---|---|---|---|---|
| **K1** | insert strictly inside; header from the sheet (direct `ref`, static name, table with `headerRowCount="1"`) | the range grows; the new header cell is **blank**; refresh **fails** — *"The PivotTable field name is not valid"* — and at open, a marked cache raises that dialog | training data; consistent with the header check the engine already enforces | **refuse** (align the reason: the new field has no name the engine can prove). No lift is coherent: grow-and-mark ships a workbook that errors at open |
| **K2** | insert strictly inside a **headerless table** source (`headerRowCount="0"`) — `table_edit` synthesizes `Column<id>` in `<tableColumns>`, so the new field *is* named | refresh succeeds; a new unused field appears at the inserted ordinal; later ordinals shift **up** | training data | mechanical (same carriers as K3, `+1` instead of `−1`, one new bare `pivotField` + one blank-inventory `cacheField`) — but the shape is rare and each lift needs its own overflow/namespace/provenance pass. **Recommend refuse** in this PR |
| **K3** | delete strictly inside; the deleted ordinal **unreferenced**; the source keeps ≥ 2 columns | the range shrinks; refresh drops the `cacheField`, each record loses its value, later ordinals shift **down**, the rendered cells **do not change** | ordinal shift is ECMA-376 positional by definition; unchanged cells are **corpus-provable** (§3) | **lift** (recommended) or refuse |
| **K4a** | delete of a **referenced** ordinal that is one of ≥ 2 data fields (the form stays inside the slice: one row field, ≥ 1 data field left) | refresh drops the `dataField`; the values column vanishes; `location` narrows | training data; the surviving columns are corpus-pinned, the narrowing is not | lift later or refuse. **Recommend refuse** in this PR: silently dropping a rendered values column as a side effect of a sheet edit is the surprising outcome a refusal exists for; a follow-up slice can lift it if wanted |
| **K4b** | delete of the row field, the only data field, or an ordinal named by `baseField` | refresh re-lays into a **form the slice does not lay out** (no row axis / no values), `baseField` has no defined successor | training data | **refuse** |
| **K5** | delete that collapses the source to zero width | no meaningful result | — | **refuse** (the row-collapse twin) |
| **K6** | any column edit inside a source with **no finite rectangle** — whole columns, whole rows, an unbounded name body, a consolidation set, an unknown-`type` locator | — | — | **stays refused** — decided by policy B at the S7b gate, not re-opened here |
| **K7** | column edits *outside* every source; external / dangling spellings | shift / no-op | shipped (S7b-2) | unchanged |
| **K8** | a `setCell` into the **header row** at save (a rename — the schema by name) | refresh treats the renamed column as a new field: the old field's settings are lost, the layout drops it | training data | **unchanged**: the save path never refuses; the header-vs-names check already fails the rebuild and the cache takes the marker alone (S7b-5's fallback). Excel's open-refresh then does exactly the drop described — the safety net working as designed. Rewriting `cacheField@name` in place instead would *diverge* from Excel (Excel matches by name, not position), so a "rename lift" is ruled out on fidelity, not on effort |

### The K3 predicate, in full

Lift a `deleteColumn` at 1-based sheet column `i`, `tl_col ≤ i ≤ br_col`, of
cache C with source rectangle `[tl_col … br_col]`, iff **all** of:

1. C resolves to a finite-rectangle worksheet source (direct `sheet`+`ref`,
   table, static name body) and passes every S7b-4 admission (`checkShape`,
   header row = field names, cell reads, budgets, namespace hygiene);
2. every consumer of C passes every S7b-5 admission (the one-row-field form);
3. `br_col − tl_col ≥ 1` (K5 otherwise);
4. the deleted ordinal `k = i − tl_col` is unreferenced: `pivotField k`
   carries no `axis` and no `dataField="1"` **and** no `<field x="k">` on any
   axis, no `dataField@fld == k`, no `@baseField == k`, no admitted
   `reference@field == k` — a disagreement between the two directions
   refuses (`MalformedPivotXml`, as the slice's other cross-checks do);
5. `pivotField k` is removable whole: its children (an off-axis `<items>`
   list, its own `extLst`) leave with the element; no *other* pivotField's
   content indexes it (none can — item indexes are per-field);
6. for a table-named source, `table_edit.checkEditSafe` admits the same
   delete (it does: the matching `<tableColumn>` is dropped whole, its own
   formulas with it; only collapse refuses). A *sibling* column's
   `calculatedColumnFormula` naming the deleted column follows today's
   shipped table semantics — the identical delete on a pivot-less table is
   already admitted, so K3 inherits, not widens, that behavior;
7. no slicer cache or timeline cache is **attached to a consumer of this
   cache** — only the attachment list is read
   (`pivots.attachedPivotNames`: `<pivotTables><pivotTable name=…>` under
   the part's own x14 / x15 namespace, through the shared scanner); a part
   that cannot be read refuses, and `sourceName`-level matching stays
   untyped. *(Amended from the package-level presence gate during
   implementation: the corpus itself carries a slicer — on the iris pivot —
   so presence alone would refuse the row's own corpus proof for a mtcars
   edit. Attachment is the narrowest read that keeps the conservative
   direction.)*;
8. no consumer's host sheet carries a `<pivotSelection>` at all (v1
   presence gate — the schema makes `<pivotArea>` a required child, so
   element presence is the check; `sheet_edit` keeps rewriting the four
   absolute coordinates for every other edit, and the corpus sheets carry
   none).

Everything else inside → the per-case refusal, each with its own test.

### What the K3 lift writes

One transaction, the S7b sweep's existing all-or-nothing `replaceParts`:

- **the source coordinate**: `worksheetSource@ref` shrinks by the shipped
  range semantics (the refusal arm in `shiftSourceRect` becomes a caller
  decision); a table source moves with its table (`table_edit` drops the
  `<tableColumn>` and shrinks `table@ref`), a static name body with its name;
- **the definition**: the k-th `<cacheField>` element removed whole,
  `cacheFields@count` decremented — two new parser-exported spans
  (`CacheField.span`, the count's value span), never a second scanner;
- **the records + inventories + `recordCount` + `refreshedDate`(+`Iso`) +
  marker**: `engine.rebuildWith` — `rebuild` handed the schema edit — fed
  the post-edit rows (width − 1) against the effective field list, the same
  read path as a row edit (`rowsAfterColEdit` drops or blanks one value per
  row);
- **each consumer**: the k-th `<pivotField>` element removed whole,
  `pivotFields@count` decremented, every ordinal carrier `> k` decremented
  in place (`field@x`, `dataField@fld`, `@baseField`) — parser-exported
  value spans; then `engine.layout` re-lays items / rowItems / location /
  host cells exactly as S7b-5 does (unchanged bytes expected, and asserted
  in the corpus test);
- the marker stays on the rebuilt cache — S7b-5's safety-net rule, Excel's
  open-refresh landing on what zlsx wrote.

## 3. The proof story — what "the oracle" is for S7c

The corpus (`tests/corpus/openxlsx_loadExample.xlsx`) carries the K3 and K4
fixtures natively:

- **`mtCars Pivot`** (`pivotTable2`, cache 2, source `Table3`, 11 fields)
  references only ordinals 0 (`mpg`), 1 (`cyl`, the row axis and every
  `baseField`) and 5 (`wt`, twice). **Eight ordinals are unreferenced** —
  `disp` 2, `hp` 3, `drat` 4, `qsec` 6, `vs` 7, `am` 8, `gear` 9 (the one
  carrying an off-axis `<items>` list — predicate row 5's fixture), `carb`
  10.
  - Deleting `disp`'s sheet column exercises the **decrement** arm: `wt`'s
    two `dataField@fld` move 5 → 4, ordinals 3–10 shift, `x="1"` and
    `baseField="1"` hold still.
  - Deleting `qsec`'s column exercises the **no-decrement** arm (everything
    referenced sits below 6).
  - In both: the re-laid host cells must equal what Excel wrote **to the
    bit as values** — the pivot never read the deleted field, so the 20
    corpus cells (`A1:D5`) are the oracle, exactly as S7b-5's 45-cell
    proof. One nuance the implementation surfaced: the rebuild reads the
    *sheet's* lexical spellings (S7b-4's convention), while Excel's cache
    spelled the same doubles its own way — the sheet's `wt` cells spell
    `1.513` / `5.424` where the cache (and so the host) spelled
    `1.5129999999999999` / `5.4240000000000004`. The min/max cells respell,
    equal as doubles; the proof is value-exact, and byte-exact wherever the
    two lexicals agree (the average columns, every label).
- **`iris Pivot`** (`pivotTable1`, 5 fields) references **all five** ordinals
  (four data fields + the row field): every inside column delete refuses —
  K4a on `fld` 0–3 (four data fields — deleting one leaves three), K4b on
  the row field's column (`Species`) — and `Slicer_Species` is attached to
  it, so the v1 slicer gate refuses first anyway.

The synthetic fixtures (`pivots.fixture.write`) grow an unreferenced-field
variant per source spelling (direct `ref`, static name; the corpus covers
tables), pinning the direct-`ref` shrink and the name-body rewrite.

What stays on the **parked Excel oracle** (one new line in S7b-4's parked
entry, nothing else): Excel opens the K3 result without repair and *Refresh*
is a no-op on it. Identical standard to S7b-4/5, which shipped on the same
terms. Everything the corpus cannot pin beyond that — K2's synthetic-name
insert, K4a's narrowing — is *not lifted*, which is what "lift only what the
oracle proves" buys.

## 4. Recommendation and the owner's questions

**Recommendation: lift K3, refuse the rest per-case.** K3 is the one case
where the shipped machinery already proves the outcome end-to-end (the
corpus pins the cells; the ordinal map is positional by the schema; the
rebuild and re-lay are S7b-4/5's, re-fed), and the one where refusing is
pure friction — the edit provably changes nothing a consumer renders.
Refusals K1/K4b/K5/K6 align zlsx with what Excel's own refresh refuses or
mangles; K2/K4a are deferred, not ruled out.

| # | Question | Options | Recommendation | **Owner ruling (2026-08-31)** |
|---|---|---|---|---|
| Q1 | K3 — lift, or does S7c close as an all-refusal row? | lift / refuse-all | **lift** | **lift** |
| Q2 | K4a (delete one of ≥ 2 data fields) | in this PR / follow-up slice / stays refused | **stays refused**, revisit only on demand | **follow-up slice — S7c-2**, after K3 merges |
| Q3 | K2 (headerless-table insert) | in this PR / stays refused | **stays refused** | **in this PR** |
| Q4 | The v1 presence gates (7: any slicer/timeline cache part; 8: any `pivotSelection` with a `pivotArea` child) refuse K3 conservatively | accept / demand the typed narrow now | **accept** | **accept both** — gate 7 was then amended to attachment level during implementation (predicate row 7): the corpus carries `Slicer_Species` on the iris pivot, so package-level presence would refuse the row's own corpus proof |

**Gate answered 2026-08-31.** S7c-1 (this PR) lifts K3 and K2 and ships the
per-case refusals for K1 / K4a / K4b / K5; S7c-2 lifts K4a. The Codex
pressure-test of this framing could not run before the ruling (usage limit
until 2026-09-05) — it folds into the PR's review loop instead.
