# Structural-edit plan (v1 draft)

> Phase 3e (post-cell-mutate). Append-only LMS shipped (Phase 3c,
> `load-modify-save.md`). Cell-mutate plan drafted (Phase 3d,
> `cell-mutate.md`). This plan covers the next tier:
> shape-changing edits where every `(row, col)` reference in the
> file may need to shift.

## Problem

Cell-mutate (`Editor.setCell`) replaces ONE cell's bytes in place.
Real workflows often need shape changes that touch every cell ref
in the workbook:

- "Insert a row at position 5" — every row below shifts +1; every
  cell ref `=A6` in any formula stays semantically valid (rebases
  to A7 implicitly via the moved cell), but explicit
  `<mergeCells ref="A5:B7">` ranges, `<dataValidations sqref>`,
  `<hyperlinks ref>`, conditional-formatting `sqref`, pivot
  source ranges, calcChain entries, defined-name ranges — all
  need range rewrites.
- "Delete column C" — every cell to the right shifts -1; every
  formula body that mentions D / E / F columns needs textual
  rewrite (`=SUM(D2:F2)` → `=SUM(C2:E2)`); same range-rewrite
  matrix as above; the `<col min="N" max="M" width="...">`
  blocks need renumbering.
- "Rename Sheet1 to 'Summary'" — every formula in the workbook
  that says `'Sheet1'!A1` becomes `Summary!A1`; the workbook
  rels, defined names, pivot source refs, external links all
  need rewrites.
- "Delete Sheet1" — workbook.xml + rels lose the sheet entry;
  every formula in sibling sheets that references Sheet1
  becomes `#REF!` (Excel's behaviour) or needs to be tracked.

Each operation has a different blast radius. The append-only +
cell-mutate path can sidestep all of this; structural edits cannot.

## Scope choice — three split plans

Three operations, each its own multi-iter scope:

1. **Row insert / delete** — fewest range-rewrite consumers,
   fewest formula-body rewrites (numeric row component shifts in
   `r="N"`). The smallest of the three.
2. **Column insert / delete** — adds the column-letter rewrite
   layer. Every formula body that mentions a shifted column
   needs textual transform. Bigger.
3. **Sheet add / rename / delete** — workbook-level. Touches
   workbook.xml + rels + every formula's sheet prefix. Biggest.

This plan splits each into its own ship, in that order. Don't try
to ship two together — the range-rewrite matrix is shared but
each operation has its own correctness tests.

## Shared architecture: range-rewriter

All three operations need to walk and rewrite ranges in:

- `xl/worksheets/sheetN.xml`
  - `<dimension ref>`
  - `<mergeCells ref>`
  - `<dataValidations sqref>`
  - `<conditionalFormatting sqref>`
  - `<hyperlinks ref>`
  - `<f ref>` for shared / array formulas
  - `<row r=N>` (numeric shift only)
  - `<c r="A1">` (column-letter and/or row-num shift)
  - Formula body text inside `<f>` (column insert/delete only)
- `xl/_rels/sheetN.xml.rels` — usually no range refs but check
  external-link targets that embed a range
- `xl/sharedFormulas` — calcChain.xml (refs by sheet-rid + cell)
- `xl/workbook.xml` — `<definedName>` (named ranges)
- `xl/pivotTables/pivotN.xml` — source ranges
- `xl/tables/tableN.xml` — `<table ref="A1:Z100">`
- `xl/externalLinks/...` — when present
- Drawing anchors (`<xdr:from>` / `<xdr:to>`) embed row/col indices

A common `RangeRewriter` interface that takes a `(sheet_idx,
shift_kind, axis, before_index, delta)` tuple and walks every
known consumer would centralise the work. Each operation's plan
spec maps to a sequence of `RangeRewriter` calls.

## Plan A — Row insert / delete (Phase 3e)

### Scope

- `Editor.insertRow(sheet_idx, before_row)` — open a gap at
  `before_row`, push every existing row +1.
- `Editor.deleteRow(sheet_idx, row)` — remove the row, pull
  every later row -1.
- Multi-row variants: `insertRows(sheet, before, count)` /
  `deleteRows(sheet, range)`.

### Rewrites

| Place                            | Action                                |
|----------------------------------|---------------------------------------|
| sheet `<row r="N">`              | renumber `r=N` for shifted rows       |
| sheet `<c r="A1">`               | renumber row component of every ref   |
| sheet `<dimension>`              | recompute (best-effort) or leave stale |
| `<mergeCells ref="A1:B5">`       | shift row bounds; drop empty merges   |
| `<dataValidations sqref>`        | shift row bounds                      |
| `<conditionalFormatting sqref>`  | shift row bounds                      |
| `<hyperlinks ref>`               | shift row bounds                      |
| `<f ref>` shared/array           | shift row bounds                      |
| `<definedName>`                  | shift row component in body if it    |
|                                  | contains a `'Sheet'!A:Z` range        |
| pivot source ranges              | shift row bounds                      |
| table refs                       | shift row bounds                      |
| drawing anchors                  | shift row index                       |
| calcChain                       | drop entries for deleted rows         |

### Formula body rewrite

For `insertRow`: formula bodies mostly stay valid because Excel
re-resolves cell refs against current cell positions. **But**
formulas with explicit row literals like `=$A$5` need the literal
shifted from 5→6 if the row at position 5 was pushed down. This
is a real textual transform — same shape as column rewrite below.

For `deleteRow`: any formula that explicitly references the
deleted row becomes `#REF!`. Excel handles this on next save;
v1 leaves the formula text intact and lets Excel recompute.
*(Historical: "lets Excel recompute" was the only option when this
was written. D1 was reversed 2026-08-02 — once recalculation lands
on the save path, `goal_formula.md` M5d2, deferring to Excel becomes
a choice rather than the sole outcome. The v1 behaviour described
here is unchanged.)*

### Phasing — 4 iters

1. **iter-row-1**: `RangeRewriter` framework + tests on
   synthesised inputs.
2. **iter-row-2**: `Editor.insertRow` (single-row form), driving
   RangeRewriter across worksheet ranges. Defer formula body
   rewrites; rely on Excel's re-resolution.
3. **iter-row-3**: `Editor.deleteRow` (single-row form). calcChain
   entries for deleted rows must be dropped.
4. **iter-row-4**: multi-row variants + C ABI + Python + CLI.

### Risks

- **Self-closing rows + sparse layouts**: a worksheet with rows
  1, 2, 5, 7 (no 3, 4, 6) is fine — `insertRow(0, 5)` shifts
  rows 5 and 7 to 6 and 8, leaves the gap. Sparse layouts must
  not be densified by accident.
- **Row-1 insert + frozen panes**: `<sheetView>` `<pane ySplit="N">`
  refers to a row index. If we insert above the freeze line, the
  freeze should follow.

## Plan B — Column insert / delete (Phase 3f)

Strictly harder than Plan A because of formula-body rewrites and
the column-letter (string) shift. Don't ship until Plan A's
RangeRewriter is settled.

### Scope

- `Editor.insertColumn(sheet, before_col)` / `deleteColumn`.
- Single-column form first; multi-column variants in iter-col-4.

### Rewrites (additive on top of Plan A's matrix)

| Place                            | Action                                |
|----------------------------------|---------------------------------------|
| sheet `<c r="A1">`               | column-letter shift A→B, B→C, ...     |
| sheet `<col min="N" max="M">`    | renumber min/max bounds; merge or     |
|                                  | split blocks if a single col range    |
|                                  | falls inside a multi-col `<col>` block |
| Formula body text inside `<f>`   | textual transform: `A1`→`B1`,         |
|                                  | `$A$1`→`$B$1`, but ONLY column refs   |
|                                  | not function names (`AVERAGE` ≠ col   |
|                                  | A); 3-letter cols (`ZZ`→`AAA`)        |
| Range refs everywhere            | column component shift in addition    |
|                                  | to (or instead of) row shift          |

### Formula textual rewrite — the hard part

A safe formula-rewriter must:

1. Tokenize formula text against Excel's grammar (function names,
   string literals, sheet-prefixed refs, structured refs).
2. Identify every cell-ref token and shift its column component
   if `>= before_col` (insert) or `> deleted_col` (delete).
3. Emit `#REF!` for tokens equal to `deleted_col` on delete.

This is a small parser, not a regex transform. A regex would
mis-shift column letters embedded inside string literals
(`"Column A"`) or function names (`COLUMN`).

### Phasing — 5 iters (one extra vs row plan)

1. **iter-col-1**: formula tokenizer + tests against the OOXML
   grammar reference.
2. **iter-col-2**: extend `RangeRewriter` for column-axis shifts.
3. **iter-col-3**: `Editor.insertColumn` (single-col).
4. **iter-col-4**: `Editor.deleteColumn` (single-col, with
   `#REF!` emission).
5. **iter-col-5**: multi-col + bindings + CLI.

### Risks

- **3-letter column overflow**: `ZZ` → `AAA` is a real wrap-
  around. The shift must use the existing `columnIndexFromRef`
  / `formatCellRef` round-trip, not naive ASCII increment.
- **Structured table refs**: `Table1[Column A]` references
  shouldn't shift — they're name-resolved, not position-resolved.
- **Sheet-prefixed refs**: `Sheet2!A1` — the Plan B rewriter
  must shift only the column part, not the sheet name.

## Plan C — Sheet add / rename / delete (Phase 3g)

Workbook-level. Smallest geometric blast radius (no range shifts)
but deepest semantic blast radius (every formula in every other
sheet that references this one needs a string rewrite).

### Scope

- `Editor.addSheet(name) -> sheet_idx` — append a new sheet at
  the end (or at a caller-specified position).
- `Editor.renameSheet(sheet_idx, new_name)` — change the sheet
  name, propagating to every formula that prefixes it.
- `Editor.deleteSheet(sheet_idx)` — remove the sheet entirely;
  every formula referring to it becomes `#REF!`.

### Rewrites

| Place                            | Action                                |
|----------------------------------|---------------------------------------|
| `xl/workbook.xml`                | add/rename/remove the `<sheet>` entry |
| `xl/_rels/workbook.xml.rels`     | rels target update for add/delete     |
| `[Content_Types].xml`            | new Override for added sheet          |
| `xl/worksheets/sheetN.xml`       | for add: write a fresh sheet body;    |
|                                  | for delete: drop the entry entirely   |
| Every formula `'Sheet1'!A1` ref  | rename or `#REF!` substitution        |
| `<definedName>` body             | same                                  |
| Pivot source-sheet refs          | same                                  |
| External-link sheet refs         | same                                  |
| Calc chain                       | drop entries for deleted sheet        |
| Drawings / charts / oleObjects   | preserved as-is for add; for delete   |
|                                  | drop the rels entries                 |

### Phasing — 5 iters

1. **iter-sheet-1**: `Editor.addSheet` (simplest — no rewrites,
   just emit a new entry; existing tests stay green).
2. **iter-sheet-2**: formula-text sheet-ref rewriter (reuses the
   tokenizer from Plan B if landed first).
3. **iter-sheet-3**: `Editor.renameSheet`.
4. **iter-sheet-4**: `Editor.deleteSheet` with `#REF!` cascade.
5. **iter-sheet-5**: multi-sheet variants + C ABI + Python + CLI.

### Risks

- **Sheet-name escaping**: names with special chars require
  apostrophe-quoting (`'Sheet 1'!A1` not `Sheet 1!A1`). Already
  handled by the writer; renamer must apply the same rules.
- **Charts that reference the renamed sheet**: chart XML embeds
  series ranges with sheet-prefix. Need to update those too.
- **Index stability**: caller's stored `sheet_idx` becomes invalid
  after `deleteSheet` shifts indices. v1 contract: delete
  invalidates all subsequent indices; rename is index-stable.

## Cross-cutting concerns

### Test strategy

- **Round-trip parity**: for every operation, take a corpus
  fixture, apply the operation, save, re-open with `Book.open`,
  walk every cell and verify positions/values match expectation.
- **Excel + LibreOffice compatibility**: open the round-tripped
  output. Excel must not flag the file as repaired.
- **Fidelity**: parts the writer doesn't model (charts, pivots,
  drawings, custom XML, VBA) round-trip byte-for-byte.
- **Adversarial fixtures**: every operation must safely refuse
  malformed inputs (`error.MalformedXml` etc.) rather than
  emitting corrupt output.
- **Fuzz**: random sequences of insert/delete operations against
  a fixture; assert the resulting workbook re-opens.

### Deferred / out of scope (even after C)

- **Conditional formatting rewrites with formulas**: same
  formula-tokenizer concerns as Plan B. Defer until the
  tokenizer lands.
- **Pivot table refresh**: pivots cache their data; structural
  edits don't refresh the cache. Excel handles refresh on
  open, but a `<pivotCacheRefresh>` flag could be set
  defensively. v1 does not.
- **Cross-workbook external links**: `[file.xlsx]Sheet1!A1` refs
  don't get rewritten — we don't open the external workbook.
  v1 leaves them intact.
- **Macros / VBA refs**: VBA code may hard-code sheet names.
  No VBA parser; `.xlsm` files preserve VBA byte-for-byte but
  the macros may break semantically. Document in the API.

### Implementation ordering

Don't ship Plan B before Plan A — they share the
`RangeRewriter` framework. Don't ship Plan C before Plan B
unless you're willing to implement formula-text rewriting
twice (the column shift and the sheet-rename rewriter share
the same formula tokenizer).

Realistic order: 3e → 3f → 3g, each shipping fully (4-5
iters each) before the next starts.
