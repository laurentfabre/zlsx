# Cell-mutate plan (v1 draft)

> Phase 3d. Append-only LMS shipped (Phase 3c, see
> `load-modify-save.md`). Today the only mutation `Editor` supports
> is appending whole rows to a sheet. Real workflows want
> single-cell changes (status updates, recalculated totals,
> "fix the typo in B7") without re-running the whole producer.

## Problem

`Editor.appendRows` is built on a literal `</sheetData>` substring
search — find the closing tag, splice new `<row>` blocks before it.
That model can't express:

- "Set A1 to 'Done'" — the existing `<c r="A1">` span must be
  replaced in place, not appended.
- "Delete row 5" — every `<c r="…">` and the `<row r="5">` wrapper
  has to come out, and references in `<f>` array ranges or shared
  formulas may need follow-up.
- "Insert column C" — every cell ref to the right shifts; the
  worksheet, every formula's `<f>` body, every `mergeCell range,
  every `<dataValidations>` `sqref`, every `<hyperlinks>` ref, and
  the `<dimension>` need to move.

Each of those steps is its own correctness contract. The append-
only path can defer all of them; cell-mutate cannot.

## Scope choice — single-cell value mutation v1

Three increasingly ambitious targets:

1. **Single-cell value mutation**: `Editor.setCell(sheet, row, col, cell)`
   replaces the value AND/OR style on one cell, leaving every other
   byte of the worksheet untouched. No structural changes (no row
   add/delete, no column shift, no formula range rewriting).
2. **Row-level mutation**: `insertRow`, `deleteRow`. Requires
   shifting every cell ref below the affected row + updating
   ranges (merges, validations, hyperlinks, conditional formats,
   pivots).
3. **Column-level + workbook-level**: `insertColumn`, `deleteColumn`,
   sheet rename, sheet add/delete. Largest blast radius — touches
   every formula in the workbook, sharedStrings (sheet names in
   structured refs), styles.xml, theme, etc.

This plan targets **(1) single-cell value mutation** as v1. (2)
and (3) are their own follow-up plans because both involve range-
rewriting that the v1 tokenizer doesn't need to model. v1 covers
the headline workflow ("status A1 went from 'pending' to 'done'",
"recalculated total in B12"); (2)-(3) cover schema changes.

## Constraints

- **Fidelity**: anything we don't touch round-trips byte-for-byte.
  In particular: formulas (`<f>`), inline strings (`<is>`),
  phonetic hints (`<rPh>`), unknown `<c>` attributes (`vm`,
  `cm`, etc.), unknown `<c>` children (`extLst` extensions).
  This is the principal risk.
- **Sheet preservation**: parts the writer doesn't model (charts,
  pivots, drawings, custom XML, VBA, theme overrides) MUST
  round-trip byte-for-byte. Cell-mutate touches only worksheet
  XML, never these.
- **Reader contract**: `Book.open` keeps working unchanged.
- **Editor contract**: `Editor.appendRows` keeps working
  unchanged. `setCell` is a new method.
- **Concurrency**: single-threaded per `Editor`.
- **`<dimension>` update**: best-effort only, same as the append
  path. Excel recomputes on next save anyway.
- Stdlib only.

## Approach: span-preserving worksheet tokenizer

The append-only path operates at the byte level (substring search
for `</sheetData>`). Cell-mutate needs structure: for every `<c>`
in the worksheet body, we need to know:

```zig
const CellSpan = struct {
    /// Byte offset of the `<` opening the `<c` tag (in the
    /// decompressed sheet XML).
    start: usize,
    /// Byte offset just past `</c>` (or just past `/>` for
    /// self-closing cells).
    end: usize,
    /// Byte offset just past the `>` ending the opening `<c …>`
    /// (== `end` for self-closing). Used to splice cell values
    /// without re-emitting the existing attribute set.
    body_start: usize,
    /// 1-based row, 0-based column resolved from `r="A1"` (with
    /// the implicit-column fallback from iter-impl-col).
    row: u32,
    col: u32,
};
```

**Tokenizer iter-cm-1**: Walk worksheet XML once at first
mutation, building a `[]CellSpan` per row plus an index
`HashMap(row, []CellSpan)`. Same span-recording loop the existing
`Rows` iterator runs, just exposed as a side-table the Editor
keeps. Idempotent: rebuild only when the worksheet body has been
mutated since last walk (dirty flag).

**setCell iter-cm-2**: Lookup `(row, col)` in the span index.
Three cases:

1. **Cell exists**: replace `[start, end)` in the worksheet XML
   with a new `<c r="…" …>…</c>` whose attributes preserve every
   non-targeted attribute from the source span (style index `s`
   only changes if the caller asked; `t`, formula `<f>`, inline
   string `<is>`, etc. are preserved iff the caller is changing
   only the value).
2. **Cell missing in row**: insert a new `<c>` at the right
   position inside the existing `<row>` — find the lexicographic
   neighbours by column ref.
3. **Row missing entirely**: insert a new `<row r="N">` at the
   right position inside `<sheetData>` — same row-position
   search the v1 plan describes for `<dimension>`.

In all three cases: shift every later `start`/`end` offset in the
span index by the delta. Cheap when the index is sorted by start.

## Why this can't reuse the substring-splice path

The append-only `</sheetData>` search works because:

- The mutation point is always the same literal string.
- The mutation always APPENDS, never modifies in place.

Cell-mutate is the opposite shape:

- The mutation point is data-dependent (depends on `(row, col)`
  and the source's existing structure).
- The mutation REPLACES bytes, which means every subsequent span
  in the index must be re-anchored.

A span-preserving tokenizer is the smallest abstraction that
expresses both shapes correctly.

## Public API delta

```zig
pub const Editor = struct {
    // existing: open, deinit, save, appendRows
    // new in iter-cm-2:
    pub fn setCell(
        self: *Editor,
        sheet_idx: u32,
        row: u32,
        col: u32,
        cell: xlsx.Cell,
    ) !void;

    // optional iter-cm-3: bulk variant (mutate many cells in one
    // call, amortising span-table rebuild). Same semantics as
    // calling setCell N times.
    pub fn setCells(
        self: *Editor,
        sheet_idx: u32,
        edits: []const struct { row: u32, col: u32, cell: xlsx.Cell },
    ) !void;
};
```

`Editor.appendRows` and `Editor.save` stay unchanged. `Book.open`
unchanged.

## Phasing — 4 iters

1. **iter-cm-1** (read-only worksheet tokenizer): introduce
   `WorksheetSpans` — a `[]CellSpan` + row-index hashmap built by
   walking the worksheet body once. Behind an opt-in
   `Editor.scanWorksheet(sheet_idx)` API; no public mutation
   surface yet. Tests: round-trip every corpus file's sheet[0]
   through the scanner, assert `(row, col)` resolution matches
   `Book.rows` for every cell.

2. **iter-cm-2** (in-place `setCell`): expose
   `Editor.setCell(sheet, row, col, cell)`. Implements all three
   cases (cell exists / cell missing / row missing). Tests:
   round-trip every corpus file with one cell mutated; assert
   the surrounding sheet XML matches the source byte-for-byte
   except for the targeted `<c>` span. Add a fixture with a
   formula cell and assert `setCell` on a *different* cell
   leaves the formula intact.

3. **iter-cm-3** (bulk + style-only mutation): expose
   `Editor.setCells` for amortised mutation; expose a
   `setStyle(sheet, row, col, style_idx)` shorthand that touches
   only the `s="…"` attribute. Tests: pathological case where
   every cell in a 100×100 sheet gets a different value; assert
   linear-time perf (no quadratic span-rewrite).

4. **iter-cm-4** (bindings + CLI + docs): C ABI
   (`zlsx_editor_set_cell`), Python binding
   (`Editor.set_cell(sheet_idx, row, col, value)`), CLI
   `zlsx set-cell <file> --sheet N --ref A1 --value '"hello"' --out <file>`.

## Testing strategy

- **Tokenizer round-trip** (iter-cm-1): every corpus file's
  sheet[0] gets scanned; for every `(row, col)` the iterator
  yields, the scanner must surface a `CellSpan` whose `r=`
  attribute resolves to the same `(row, col)`. No drift.
- **Single-cell mutation byte-for-byte** (iter-cm-2): take a
  worksheet, change ONE cell value, confirm:
  - The mutated cell's `<c>` span has the new value.
  - Every other byte of the worksheet XML is identical to
    source (modulo `<dimension>` recompute).
  - The non-worksheet parts of the archive (rels, styles, SST,
    charts, drawings, VBA) are byte-identical.
- **Formula preservation** (iter-cm-2): a sheet with a `<f>` in
  cell B2 — `Editor.setCell(0, 1, 1, .{ .integer = 99 })` (cell
  A1) must leave B2's formula untouched.
- **Inline-string preservation** (iter-cm-2): a sheet with
  `t="inlineStr"` in C3 — `setCell` on A1 must leave C3's
  `<is>...</is>` body intact.
- **Phonetic / unknown-child preservation** (iter-cm-2): a
  worksheet with `<rPh>` (phonetic hints, common in Japanese
  workbooks) and `<extLst>` (unknown extensions) — mutating an
  unrelated cell must leave both intact.
- **Excel + LibreOffice compatibility**: open the round-tripped
  output. Excel must not flag the file as repaired.
- **Fuzz**: random `setCell` operations against a fixture; assert
  `Book.open` re-reads every cell as expected and every
  unchanged cell is identical to source.

## Open risks

- **Span shift cost on bulk mutations**: every `setCell` shifts
  every later span. Quadratic in worst case (mutate 1000 cells
  → 1M offset updates). Mitigation: `setCells` builds the new
  worksheet XML in one pass.
- **Stable cell-ref invariants**: `<c r="A1">` is the canonical
  form, but the tokenizer must accept r-less cells (per
  iter-impl-col). When INSERTING a missing cell, we always emit
  `r="A1"` form — even if surrounding cells are r-less. Risk:
  some readers may produce a slightly different canonical form
  on next save; tolerable.
- **Style index validation**: `s="N"` references styles.xml.
  v1 does not let `setCell` register new styles — the caller
  must pass a style index that already exists. Adding new
  styles to an existing styles.xml is its own xform (defer to
  iter-cm-5+).
- **Shared formula refs**: a `<c>` with `<f t="shared" si="N">`
  is part of a shared-formula group. If we replace such a cell
  without preserving the `<f t="shared">` opening, downstream
  cells with `<f t="shared" si="N"/>` (the slaves) lose their
  formula. v1 contract: `setCell` rejects mutation of shared-
  formula base cells with `error.SharedFormulaBaseCellMutate`;
  callers must mutate slave cells only or use a future
  `clearFormula` API.
- **Array-formula refs**: similar shape (`<f t="array" ref="X:Y">`).
  Same v1 rejection contract.
- **`<dimension>` widening**: same rule as append — only the
  canonical `ref="A1:Z100"` form is updated; others left
  unchanged (Excel recomputes).
- **Calc chain**: `xl/calcChain.xml`. Mutating a cell that's part
  of the calc chain doesn't break the chain (the chain just
  becomes stale; Excel rebuilds on next save). v1 leaves it
  unchanged.

## What's out of scope (queued)

- **Adding new styles**: `Editor.addStyle(...)` would mutate
  `xl/styles.xml` and re-number every `<c s="…">` reference.
  Mechanically clean but mostly mechanical XML rewriting; defer
  to iter-cm-5+.
- **Row insert / delete**: every cell ref below the affected
  row shifts. Plus range-rewrites in mergeCells, dataValidations,
  hyperlinks, conditionalFormatting, pivot caches. This is the
  big-ticket "structural edit" follow-up plan.
- **Column insert / delete**: same as row but with cell-ref
  *letter* shifts (A→B is a string transform, not just a number
  bump) and the additional pain of formula-body rewriting
  (every `=A1+B1` becomes `=B1+C1`).
- **Sheet add / rename / delete**: workbook.xml mutations + rels
  + every formula's sheet-prefixed ref ('Sheet1'!A1).
- **Cell-style mutation cross-sheet**: changing every style with
  index N to a different definition in styles.xml.

Each is its own multi-iter plan.
