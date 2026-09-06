# py-zlsx

Python binding for [zlsx](https://github.com/laurentfabre/zlsx) — a fast `.xlsx` reader + writer library written in Zig.

This package is a thin `ctypes` layer over `libzlsx` (no Rust, no PyO3, no third-party runtime deps — ctypes is stdlib). Reader benchmark: a 1,008-row workbook parses in **10.7 ms / 4.2 MB RSS** — 4× faster than `python-calamine`, 24× faster than `openpyxl`, at a tenth of the memory. Writer (Phase 3b, v0.2.4) produces styled workbooks with fonts, fills, borders, number formats, column widths, freeze panes, and auto-filter — the pragmatic openpyxl-parity set.

## Install

The wheel bundles a per-platform `libzlsx.{dylib,so,dll}` so `pip install py-zlsx` is self-contained. Running from source? Point `ZLSX_LIBRARY` at the shared library on disk, or install the Homebrew bottle (`brew install laurentfabre/zlsx/zlsx`) and the package will find it at `/opt/homebrew/opt/zlsx/lib/libzlsx.dylib`.

```bash
# (Preferred) from PyPI
pip install py-zlsx

# From source (requires a libzlsx on disk)
pip install -e ./bindings/python
export ZLSX_LIBRARY=/path/to/libzlsx.dylib   # optional — only if auto-discovery fails
```

## Read

```python
import zlsx

with zlsx.open("workbook.xlsx") as book:
    print(book.sheets)                       # ['Summary', 'Details', ...]
    print(book.sheet_state("Summary"))      # 'visible' | 'hidden' | 'veryHidden'

    for row in book.sheet(0).rows():
        # row is a list; cell types map to Python:
        #   empty → None    string → str
        #   integer → int   number → float   boolean → bool
        print(row)

    summary = book.sheet("Summary")          # by name also works
    header = next(summary.rows())

    with book.sheet(0).rows() as rows:       # formula text + error tags (0.9.0+)
        for row in rows:                      # row: cached values / error literals, as ever
            formulas = rows.formula_strings() # ['A1*2', None, ...] — own <f> text
            bases = rows.formula_refs()       # [None, CellRef(0, 3), ...] — a slave's base
            errors = rows.error_strings()     # [None, '#DIV/0!', ...] — t="e" cells
```

### From bytes — no file, no temp path

For workbook bytes that never touch a filesystem (SQL UDFs over binary
columns, object-store reads, network payloads), `open_bytes` parses the
buffer directly. The bytes are borrowed **only for the duration of the
call** — discard them right after:

```python
with zlsx.open_bytes(content) as book:      # bytes, bytearray, or memoryview
    for row in book.sheet(0).rows():
        ...
```

Requires libzlsx 0.6.0+ (`zlsx_book_open_buffer` in the C ABI).

## Write

The writer produces fresh workbooks; editing an existing one is `zlsx.edit` / `Editor`, below. Cell styles registered via `Writer.add_style` get a 1-based index; pass those indices alongside values in `write_row(styles=[…])`.

```python
import zlsx

with zlsx.write("out.xlsx") as w:
    # Register a "header" style — bold white text on blue, centred,
    # thin black border.
    header = w.add_style(zlsx.Style(
        font_bold=True,
        font_color_argb=0xFFFFFFFF,
        fill_pattern="solid",
        fill_fg_argb=0xFF1E3A8A,
        alignment_horizontal="center",
        border_bottom=zlsx.BorderSide(style="thin", color_argb=0xFF000000),
    ))
    money = w.add_style(zlsx.Style(number_format="$#,##0.00"))
    pct   = w.add_style(zlsx.Style(number_format="0.00%"))

    sheet = w.add_sheet("Summary")
    sheet.set_column_width(0, 24)    # 0-based column index
    sheet.set_column_width(1, 14)
    sheet.set_row_height(0, 24)      # header row, in points (0, 409.5]
    sheet.freeze_panes(rows=1, cols=0)              # legacy — clamps silently
    # sheet.freeze_panes_checked(rows=1, cols=0)    # raises on out-of-range
    sheet.set_auto_filter("A1:C1")

    sheet.write_row(["Name", "Amount", "Share"], styles=[header, header, header])
    sheet.write_row(["Alice", 12345.67, 0.42], styles=[0, money, pct])
    sheet.write_row(["Bob",    9876.54, 0.33], styles=[0, money, pct])

    # Workbook-level + sheet-scoped defined names (named ranges,
    # print areas, validation sources, etc.). Excel name rules
    # enforced up-front — A1-shaped, R1C1-shaped, length>255, and
    # case-insensitive duplicates per scope all raise ZlsxError.
    w.add_defined_name("Totals", "Summary!$B$2:$B$3")
    w.add_defined_name(
        "_xlnm.Print_Area",
        "Summary!$A$1:$C$3",
        local_sheet_id=0,
        hidden=True,
    )
# save happens automatically on clean exit; exception → no save
```

For round-trip verification through the same library, every `Style` field
emitted via `Writer.add_style` reads back through `Book.cell_font`,
`Book.cell_fill`, `Book.cell_border`, and `Book.cell_alignment` (alignment +
wrap_text + diagonal direction flags + entity-encoded font names all preserved).

### Style cheat sheet

The `Style` dataclass covers every openpyxl-parity style field shipped in Phase 3b:

| Field | Type | Values |
|---|---|---|
| `font_bold` / `font_italic` | `bool` | default `False` |
| `font_size` | `Optional[float]` | `None` = default (11 pt) |
| `font_name` | `Optional[str]` | `None` = "Calibri" |
| `font_color_argb` | `Optional[int]` | ARGB packed `0xAARRGGBB`, `None` = theme auto |
| `alignment_horizontal` | `str` literal | `"general"` / `"left"` / `"center"` / `"right"` / `"fill"` / `"justify"` / `"centerContinuous"` / `"distributed"` |
| `wrap_text` | `bool` | default `False` |
| `fill_pattern` | `str` literal | 19 OOXML patternTypes (`"none"`, `"solid"`, `"gray125"`, …) |
| `fill_fg_argb` / `fill_bg_argb` | `Optional[int]` | ARGB packed |
| `border_{left,right,top,bottom,diagonal}` | `BorderSide` | `BorderSide(style="thin", color_argb=0xFF000000)` |
| `diagonal_up` / `diagonal_down` | `bool` | default `False` |
| `number_format` | `Optional[str]` | OOXML format code, e.g. `"0.00%"`, `"m/d/yyyy"` |

`BorderSide.style` accepts 14 OOXML border style names: `"none"`, `"thin"`, `"medium"`, `"dashed"`, `"dotted"`, `"thick"`, `"double"`, `"hair"`, `"mediumDashed"`, `"dashDot"`, `"mediumDashDot"`, `"dashDotDot"`, `"mediumDashDotDot"`, `"slantDashDot"`.

## Recalculate & evaluate

libzlsx 0.8.0+ ships the formula engine across the C ABI. The binding
feature-probes every symbol group, so py-zlsx keeps importing against an
older dylib — the methods below then raise `RuntimeError`.

```python
import zlsx

ed = zlsx.Editor("model.xlsx")

# In-memory transaction: every formula cell recomputed, then swapped in
# as the final operation. On ANY failure the workbook is exactly as it was.
report = ed.recalculate()                       # RecalcReport
print(report.cells_written, report.resolved.now, report.resolved.seed)

# Atomic file transaction (§5.7.9): recalc, write, rename, THEN swap.
# A pre-commit failure leaves the destination's prior bytes (or its
# absence) and this editor's memory untouched.
report = ed.save_with_recalc("model_out.xlsx", timeout=30.0)

# Standalone cache-based evaluation — never mutates anything.
r = ed.evaluate("=SUM(A1:B9)")                  # EvalResult
r.value                                          # float | str | bool | ExcelError | Matrix
r.resolved                                       # the exact resolved context — replay it
                                                 # to reproduce a defaulted evaluation

# Whole-editor state to memory, and back.
blob = ed.save_to_buffer()                       # bytes
ed2 = zlsx.Editor.from_bytes(blob)               # copies; the borrow ends at the call

# Keep every cache, set fullCalcOnLoad="1" for the next consumer.
ed.mark_recalc_on_load()
```

Writer-side, `save(recalculate=...)` routes the save through the recalc
orchestrator, so every cached formula value in the destination is one the
engine computed:

```python
w = zlsx.Writer("fresh.xlsx")
s = w.add_sheet("Calc")
s.write_row_with_formulas([1, 0], [None, "A1+2"])
report = w.save(recalculate=zlsx.RecalcOptions())   # B1's <v> is now 3
```

**Context.** `now` (epoch ms or `datetime`) and `seed` default to the
clock / OS entropy *in the binding* — the library itself never reads
either, which is what makes "equal inputs ⇒ equal output" true. The
resolved context is echoed on every report and `EvalResult`, so a
defaulted run is reproducible by replaying `.resolved`. `mode` is
`"excel"` (default) or `"ieee"`; `on_unsupported` is `"refuse"`
(default) or `"keep_stale_and_mark"`.

**Results.** Blank never escapes — it publishes as `0.0` (§12.2). An
Excel error value (`#DIV/0!`) is a *successful result*: it arrives as
`ExcelError` (a `str` subclass), never as an exception. Rectangular
results arrive as `Matrix(rows, cols, cells)`.

**Refusals.** A construct the engine does not implement raises
`ZlsxFormulaRefusal` (a `ZlsxError` subclass) with `.error_name` (e.g.
`"FormulaUnsupportedFunction"`), `.cells` — the refusing cells as
`(sheet, row, col)`, row 1-based / col 0-based — and `.census` (full
`CensusEntry` records). Nothing is mutated on a refusal.

**Cancellation.** Cancellable calls run the FFI on a worker thread, so
Ctrl-C reaches the engine as a token trigger instead of a blocked signal
handler. `timeout=` (seconds) raises `TimeoutError` **only when the
cancellation is observed before the commit point**; a cancellation that
lands after the rename returns normally with
`report.cancelled_late=True` — the transaction completed, and saying
otherwise would be a lie about the filesystem.

**CSE rectangles.** `write_row_with_formulas` accepts
`FormulaSpec.cse(text, ref)` on the rectangle's top-left cell only; the
range's other cells arrive as plain value slots in later rows (empty
members become bare `<c>` placeholders), and the save refuses while any
rectangle is missing members. `dialect="dynamic_array"` is reserved and
currently refused.

`zlsx.engine_fingerprint()` returns the engine identity string (semver,
rule versions, target triple, build hash). Two processes may share
recalc results only when their fingerprints match.

## Structural edits & pivots

libzlsx 0.9.0+ exports the editor's structural edits and the typed pivot
read. Rows are 1-based, columns 0-based (`A = 0`, as in `set_cell`), sheet
indices 0-based; every edit is staged and committed by `save` /
`save_to_buffer`, and untouched parts stay byte-identical.

```python
import zlsx

with zlsx.edit("report.xlsx") as ed:
    ed.insert_row(0, 2)                    # a blank row 2; everything below shifts
    ed.delete_column(0, 3)                 # column D goes; E.. become D..
    idx = ed.add_sheet("Archive")          # -> 2
    ed.rename_sheet(1, "Summary 2026")     # 'Old'!A1 references follow (not a pivot's worksheetSource@sheet — see below)
    ed.rename_table_column("Sales", "Qty", "Quantity")   # Sales[Qty] follows too
    ed.save("report.xlsx")

for p in zlsx.pivots("report.xlsx"):       # the `zlsx pivots` records, as dicts
    if p["kind"] != "pivot":               # "pivot_cache": a cache no table reads
        print("orphan cache", p["cache"]["id"])
        continue
    cache = p["cache"]                     # null when the part reaches no cache
    src = cache["source"] if cache else None
    if src is None:
        where = "no cache"
    elif src["type"] == "worksheet":       # sheet / table / defined-name spellings
        where = src["resolved"]            # {"sheet":…,"sheet_idx":…,"via":…,"bounds":…}, {"external":…} or None
    elif src["type"] == "consolidation":
        where = [rs["resolved"] for rs in src["range_sets"]]
    else:                                  # "external", "scenario", "unknown"
        where = src["type"]
    print(p["name"], p["location"]["ref"], where)
```

A structural edit carries the rewriters the CLI's `insert-row` family
carries — formulas in every dialect (A1, 3D, R1C1, structured
references), defined names, hyperlinks, DV / CF, merges, panes,
autoFilter, tables, drawings, comments, `<xm:f>` extensions, chart
`<c:f>` series formulas, a hosted pivot's rectangle and a cache's source
range (a source whose content
changes is rebuilt during the edit and committed by `save`). One hole
the row inherits from the Zig editor: `rename_sheet` / `delete_sheet`
do not rewrite a pivot cache's `worksheetSource@sheet`, so a source
spelled by sheet name goes stale and `pivots()` reports it as
`"resolved": null`. Where the workbook cannot be kept consistent
the edit **refuses** rather than corrupt it, as a `ZlsxRefusal` whose
`error_name` says why:

| `error_name` | Raised by |
|---|---|
| `RowEditUnsafeForSheet` / `ColEditUnsafeForSheet` | an edit inside a hosted pivot's footprint, on a host sheet a pivot also reads from, one that would collapse a table or delete its header row, or a carrier the scan cannot read |
| `DuplicateSheetName` | `add_sheet` / `rename_sheet` (ASCII case-insensitive) |
| `CannotDeleteLastSheet` | `delete_sheet` on the only sheet |
| `TableColumnNameInUse` | `rename_table_column` to a name another column holds |
| `MalformedPivotXml` | `pivots()` on a graph it cannot read whole — never a partial inventory |
| `MalformedWorkbookXml` | also `defined_names()` on an inventory it cannot serve faithfully — a carrier that does not decode, malformed UTF-8, a body with embedded markup — never a record that lies — and `conditional_formats()` / `anchors()` / `sheet_props()` / `calc_props()` on a sheet list the strict workbook read cannot prove; `calc_props()` also on a `<calcPr>` slot it cannot report faithfully — two at the slot, one an MCE branch could project there, a duplicate attribute, a carrier that does not decode |
| `MalformedSheetXml` | also `conditional_formats()` on a sheet part the strict walk cannot serve faithfully — mismatched nesting, a namespace shape that could ghost a rule, an unterminated or markup-carrying formula, a carrier that does not decode — never a partial inventory (a broken second sheet refuses the first sheet's servable records too); and `sheet_props()` on a sheet part the strict walk cannot prove a pane / extent for — a second `<dimension>` / `<sheetViews>` / first-view `<pane>`, a duplicate attribute on that machinery, an MCE construct at a recognized slot, a carrier that does not decode |
| `MalformedDrawingXml` | also `anchors()` on a drawing graph the strict walk cannot read whole — a dangling or mistyped edge, an anchor that does not parse, a part along the chain the store cannot materialise, a series ref that does not decode, a spreadsheetDrawing binding under a name the walk cannot spell (past its 100-byte prefix limit or eight-alternate replay cap; such a workbook used to list the anchors under its other names), a drawing or chart part carrying a `<!DOCTYPE` or a `<` inside an attribute value (not well-formed XML) — never a partial inventory (a broken second sheet refuses the first sheet's servable record too); on the edit side, every row / column edit of that sheet refuses before the first mutation for the same reasons the read refuses its inventory (the drawing reference it cannot follow, the binding it cannot spell, a DTD, an anchor with no close, an unreadable corner or two corner blocks that overlap — `Editor.insert_row`'s docstring) |
| `DrawingOnUnlistedSheet` | `anchors()` on an anchored object whose worksheet part `xl/workbook.xml` does not list — no record could carry a truthful `sheet`, and dropping it would leave a hole |
| `SqrefCollapseUnsafe` | `delete_row` / `delete_column` that would collapse EVERY area of a `<conditionalFormatting>` or `<dataValidation>` `sqref` — Excel deletes such a rule outright; zlsx refuses rather than silently retarget it to the cells that slide into its place |
| `RowEditExceedsMaxRow` / `ColEditExceedsMaxCol` / `SplitPaneNotSupported` / `MalformedPaneSplit`, the carrier verdicts `MalformedSheetXml` / `MalformedDrawingXml` / `MalformedVmlDrawing` / `MalformedCommentsXml` / `MalformedTableXml` / `*CoordinateOverflow`, the workbook's own `MalformedWorkbookXml` / `IdSpaceExhausted` / … | the worksheet transform's and the sweeps' own verdicts, with their precise names — a cell that would leave the grid, a split pane, a part the walkers cannot read or materialise. The list is §10 of `docs/plans/c-abi-status-v1.md`; a generic `MalformedXml` from a rewriter's consistency guard stays a plain `ZlsxError` |

`ZlsxRefusal` is a `ZlsxError`; `ZlsxFormulaRefusal` (the engine's
Plane-2 refusals) now derives from it. Statements about the *call* stay
plain `ZlsxError`s named after the cause: `SheetIndexOutOfRange`,
`RowIndexOutOfRange`, `ColumnIndexOutOfRange`, `InvalidSheetName`,
`InvalidTableColumnName`, `TableNotFound` / `TableColumnNotFound` (a
selector that names nothing, like a sheet index), and the sequencing errors
`RowEditRequiresCleanSheet` / `ColEditRequiresCleanSheet` /
`SheetDeleteRequiresCleanState` — a structural edit needs the sheet (the
workbook, for a sheet delete) free of unsaved `set_cell` / `append_rows`
writes; save first. Indices are integers (`operator.index`; a float, a
string or a bool is a `TypeError`) in `[0, 2**32)` (`ValueError`),
checked before the call — ctypes would otherwise truncate or wrap them.

`Editor.pivots()` returns the records `zlsx pivots` prints
([docs/cli.md](../../docs/cli.md), "pivots"), parsed from the same bytes:
`{"kind": "pivot", …}` per pivot table in host-sheet order, then
`{"kind": "pivot_cache", …}` per cache no table reads; `[]` for a
workbook without pivots. It reads the editor's current workbook state:
structural edits are visible immediately — rename the host sheet and
the record names it — while staged `set_cell` / `append_rows` writes
reach the pivot graph at `save`, where a cache whose source they change
is rebuilt or marked; save, then read, to see them.

`Editor.defined_names()` / `zlsx.defined_names(path)` are the same
pattern over the `zlsx defined-names` records
([docs/cli.md](../../docs/cli.md), "defined-names"): one
`{"kind": "defined_name", …}` dict per `<definedName>` of
`xl/workbook.xml` in document order — `name`, `scope`
(`"workbook"` / `"sheet"` with `sheet` / `sheet_idx`), `body` (the
formula text as authored — nothing resolved or rewritten), `hidden`
(hidden names are streamed, not suppressed). Defined names live in
`xl/workbook.xml` only, so the read never waits for `save`: structural
edits and the name sweeps they carry — a sheet rename rewriting the
bodies — are visible immediately.

`Editor.conditional_formats()` / `zlsx.conditional_formats(path)` are
the same pattern over the `zlsx conditional-formats` records
([docs/cli.md](../../docs/cli.md), "conditional-formats"): one
`{"kind": "conditional_format", …}` dict per `<cfRule>` — sheets in
workbook order, rules in sheet-document order — carrying the rule
envelope (`sheet`, `sheet_idx`, `sqref`, `rule_type`, `formulas`,
`dxf_id`, `priority`), not the visual payload: `<colorScale>` /
`<dataBar>` / `<iconSet>` bodies and the `<dxfs>` styles stay in their
parts. The read never waits for `save` either: structural edits and
the DV/CF sweeps they carry are visible immediately, and staged
`set_cell` / `append_rows` writes never touch the rule machinery.

`Editor.anchors()` / `zlsx.anchors(path)` are the same pattern over the
`zlsx anchors` records ([docs/cli.md](../../docs/cli.md), "anchors"):
one `{"kind": "image_anchor", …}` dict per anchored image and one
`{"kind": "chart_anchor", …}` dict per anchored chart — sheets in
workbook order, a sheet's images before its charts, each class in
drawing-document order — carrying the anchor geometry (`anchor` in
`two_cell` / `one_cell` / `absolute`, `from` / `to` 1-based with EMU
offsets, `absolute` `{x, y, cx, cy}` in EMUs) and where the payload
lives (`part`; an image's `bytes` count; a chart's `chart_type` and
entity-decoded `series_refs`), never the payload: image bytes and chart
XML stay in their parts. Structural edits and the drawing sweeps they
carry are visible immediately — a rename renames `sheet`, a row insert
moves the edited sheet's anchors with the grid — and staged cell writes
never touch a drawing. A chart's `series_refs` ride the formula
rewriter with every other carrier (the chart `<c:f>` sweep), so after a
rename or a row / column edit on the sheet they name the read reports
the respelled part.

`Editor.sheet_props()` / `zlsx.sheet_props(path)` and
`Editor.calc_props()` / `zlsx.calc_props(path)` are the same pattern
over the `zlsx sheet-props` and `zlsx calc-props` records
([docs/cli.md](../../docs/cli.md), "sheet-props" / "calc-props"):
one `{"kind": "sheet_props", …}` dict per workbook sheet, workbook
order — the sheet's `<dimension ref>` as authored (`None` when the
element or the attribute is absent) and the `<pane>` of its first
`<sheetView>` as authored (`None` when there is none; `x_split` /
`y_split` / `top_left_cell` / `active_pane` / `state`, each `None` when
the source omits it, split panes reported as written) — and ONE
`{"kind": "calc_props", …}` dict for the workbook's `<calcPr>`
(`calc_id` / `full_calc_on_load` / `iterate` / `iterate_count` /
`iterate_delta` as authored, every field `None` when absent — a
workbook without `<calcPr>` is a dict of `None`s, the `doc_props()`
convention). Structural edits and the sheet sweeps they carry are
visible immediately — a rename renames `sheet`, a row insert grows
`dimension` and moves a frozen pane's split and `top_left_cell` with
the grid (a split pane is the one such an edit refuses,
`SplitPaneNotSupported`) — and `mark_recalc_on_load()` lands
`full_calc_on_load` in place; staged cell writes never touch the
extent, the views or `<calcPr>`.

## Embeddings

libzlsx 0.9.0+ writes the embedding set the E5 read surface reports:
`Editor.set_embeddings` is `Workbook.setEmbeddings` on the editor handle —
one call writes the index, a vector / hash part per coverage, the
workbook→index relationship and the recovery record in its two invisible
carriers, and replaces any previous set. `Editor.embeddable_rows` is the
read that feeds it — `Workbook.embeddableRows`, the `embed_row` records
`zlsx embed --extract` prints: the rows of a column over a range that
carry embeddable content, each with the text a model should see and the
canonical xxh3-64 content hash the write stores beside the vector. The
shape is the read side's, so read → embed → write round-trips on the same
arrays and the hashes read fresh under `zlsx embed --prune`:

```python
import numpy as np
import zlsx

with zlsx.edit("report.xlsx") as ed:
    rows = ed.embeddable_rows(0, "B2:B101", "B")       # [{"kind": "embed_row", "row": 2, "text": "…", "hash": 683…}, …]
    vectors = np.asarray(embed([r["text"] for r in rows]), dtype=np.float32)   # (len(rows), dim)
    by_row = {r["row"]: (v, r["hash"]) for r, v in zip(rows, vectors)}
    dim = vectors.shape[1]
    ed.set_embeddings("text-embedding-3-small", dim, [
        {"id": "body", "sheet": 0, "range": "B2:B101", "column": "B",
         # one slot per covered row: a row the read omitted has no vector
         "vectors": [by_row[r][0] if r in by_row else np.zeros(dim, np.float32) for r in range(2, 102)],
         "hashes": [by_row[r][1] if r in by_row else None for r in range(2, 102)]},   # None = tombstone
    ], dtype="f32")                             # or "int8-sym-per-vec": quantized in the library
    ed.save("report.xlsx")

with zlsx.embeddings("report.xlsx") as emb:
    assert emb.present
    emb.vectors("body"), emb.hashes("body"), emb.valid_mask("body")
```

`embeddable_rows(sheet, range, column, *, include_formulas=False)` returns
the records in range order — `row` 1-based, `text` as a reader sees the
cell (a shared or inline string's runs joined, entities resolved; a number's
`<v>` as written; an error's literal, `#N/A`; a boolean as `"1"` / `"0"`),
`hash` an `int` in
`[0, 2**64)` — and omits rows with nothing embeddable (`[]` for a range
with none); `include_formulas` admits formula cells with a cached value, the
coverage flag's reading. A sheet the editor holds staged `set_cell` writes
(or the header cell `rename_table_column` stages) or `append_rows` for
refuses with a `ZlsxError` (`SheetHasUnsavedMutations`
/ `SheetHasUnsavedAppends` — the parsed view the read walks does not carry
them; save and re-open, or read first); `InvalidRange` (the range, or a
column outside it) and `SheetIndexOutOfRange` are the call's; a workbook
the read cannot serve faithfully refuses with a `ZlsxRefusal`
(`MissingRelationship` / `MissingSheetPart`; `MalformedSheetXml` — a sheet
part the view cannot parse, or a row or cell it cannot place: no `r`, or one
it cannot read — 0, non-numeric, past the limit — or a ref under another row;
`MalformedSharedStringsXml`; and a cell value the
read cannot carry —
`UnsupportedCellValue`: a boolean `<v>` that is not 0 / 1, a `<v>` the number
canonicalizer cannot read, a `t="d"` ISO-8601 date, a `t` this reader does
not know, a shared-string index that is not a number, an entity the decoder
does not know; `SstIndexOutOfRange`,
`InvalidUtf8`, `UnicodeNormalizationFailed`) rather than return a record
that lies.

Statements about the write raise a plain `ZlsxError` named after the cause
— `InvalidEmbeddingInput`, `InvalidDtype` / `UnsupportedDtype`,
`SheetIndexOutOfRange`, `InvalidCoverageId`, `InvalidRange` (the range, or
a column outside it), `DuplicateCoverageId`, `CoverageOverlap`,
`InvalidXmlByte` (a control byte in the model name) — each before the
first part is written; a workbook the set cannot land in refuses with a
`ZlsxRefusal` (`MissingWorkbookRels` / `MalformedWorkbookRels`,
`IdSpaceExhausted`, `MissingRelationship`, `EmbeddingExceedsArchiveLimit` — a part past the
512 MiB cap, or the recovery record past its 16 × 200-byte ceiling: roughly
eighty coverages, or a ~3 KB model name — and `MalformedWorkbookXml`, an
`xl/workbook.xml` the previous record's strip cannot walk; each judged before
the first write, the index part's own cap excepted — it cannot fire, the
record ceiling bounds the same fields). A NumPy array crosses as one
contiguous float32 / uint64 buffer; values narrow to float32 as they are
(a float64 past its range lands as `inf`), `2**64 - 1` is the tombstone,
and a masked array's masked slots are "no value". A refusal after the first
part is written (an allocation failure, an index past the cap, a package
part the carriers cannot patch) leaves the staged set partially replaced:
close the editor without saving. A save after the write re-emits the
workbook's `<definedNames>` block: every existing name keeps `name`,
`localSheetId` and `hidden` only — its other attributes (`comment`,
`description`, `function`, `vbProcedure`, …) are dropped, as after any staged
defined-name edit (pre-existing, recorded). The recalc transactions
(`mark_recalc_on_load` then `save`, `save_with_recalc`, `recalculate`)
rebuild from the archive as opened and do not carry a staged embedding
write — call them before it, or save and re-open. `recovery_in_cells` (the
Numbers-durable carrier) is Zig-only until the editor grows a path for its
hidden sheet (its strip side shipped with the sweeps below).

`prune_embeddings()` and `strip_embeddings()` are the two sweeps `zlsx embed
--prune` / `--strip` run, on the editor handle (0.9.0+). Prune tombstones
every slot whose row is no longer embeddable and zeroes its vector — a row
deleted in plain Excel leaves its vector on disk — and returns the counts as
a dict, the fields of the CLI's `{"kind": "prune", …}` record:

```python
with zlsx.edit("report.xlsx") as ed:
    ed.prune_embeddings()      # {"redacted": 1, "stale": 0, "fresh": 99, "valid_empty": 0}
    ed.save("report.xlsx")
```

Content that drifted but is still embeddable counts `stale` and is never
redacted (re-embed those rows); the hashes `embeddable_rows` hands over prune
as all `fresh` once written; a workbook with no set, or a stripped one, is
all zeros. A staged `set_cell` on a covered row is judged as staged — a
`None` redacts its slot, any other value is `stale` — and a covered sheet
with staged `append_rows` raises a `ZlsxError` (`SheetHasUnsavedAppends`;
save first). A set the index read cannot read refuses with a `ZlsxRefusal`
(`MalformedEmbeddingSet`, `MissingEmbeddingPart`), as does a covered cell
the read cannot carry (the `embeddable_rows` names), each before the first
part write. `strip_embeddings()` removes the parts and the recovery record
from every carrier — the pre-share operation; `zlsx.embeddings(path)` then
reports `absent`, not `stripped`. It is idempotent and a no-op on a workbook
without embeddings; a `recovery_in_cells` sheet goes through the editor's
own `delete_sheet` path (so sheet indices stay honest — its rules apply:
`SheetDeleteRequiresCleanState`, `CannotDeleteLastSheet`), an
`xl/workbook.xml` the strip cannot walk refuses (`MalformedWorkbookXml`)
before the first removal, and a cell that merely spells the record's magic
is user text and stays.

## Spark (PySpark Data Source)

Spark 4.0+ / DBR 15.4+ (including serverless). `pip install py-zlsx[spark]`
locally; on Databricks the plain wheel is enough — pyspark is already there.

```python
from zlsx.spark import ZlsxDataSource
spark.dataSource.register(ZlsxDataSource)

df = (spark.read.format("zlsx")
      .option("sheet", "Sales")            # name, 0-based index, "a,b", or "*"
      .option("rowsPerPartition", 100000)  # split big sheets across executors
      .load("/Volumes/cat/schema/vol/*.xlsx"))

(df.coalesce(1).write.format("zlsx")
   .mode("overwrite")
   .save("/Volumes/cat/schema/vol/report.xlsx"))
```

| Read option | Applies to | Default | Meaning |
|---|---|---|---|
| `sheet` | batch + streaming | `"0"` | Sheet name, index, comma list, or `*` for all sheets |
| `header` | batch + streaming | `true` | First row is column names; `false` → `_c0..` |
| `inferRows` | batch + streaming | `1000` | Rows sampled for schema inference (widens across the whole sample, not just the first row) |
| `rowsPerPartition` | batch + streaming | `0` (off) | Also split each sheet into row ranges |
| `mode` | batch + streaming | `permissive` | `permissive` nulls cells that don't fit the schema; `failfast` raises naming the exact file/sheet/row/column |
| `zlsx.recalc` | **batch only** | `false` | Recompute every formula with the zlsx engine and read THOSE values instead of the workbook's cached ones |
| `zlsx.recalcUtcOffsetMin` | **batch only** | `0` (UTC) | UTC offset for the recalc context — the default is UTC like every other layer; the driver's zone never applies implicitly |
| `zlsx.recalcCacheMaxBytes` | **batch only** | `536870912` (512 MiB) | Byte bound of the per-executor snapshot cache; `0` disables it |

Reads partition per (file × sheet); paths accept a file, directory, glob, or
comma list. Writes: a `.xlsx` target is single-file mode and needs a
single-partition DataFrame (`coalesce(1)`); any other path becomes a directory
of `part-*.xlsx`. `date`/`timestamp` columns round-trip as styled Excel
serials. Write options: `sheet` (default `"Sheet1"`), `header`.

### Batch recalc

`zlsx.recalc="true"` makes the read observe the values the formula engine
computes, not whatever stale caches the workbook carries. The contract
(§12.4 of the formula plan):

- **Source files are never mutated.** The driver reads each workbook's
  bytes once, recalculates in memory, and serializes a snapshot buffer;
  schema inference and partition planning run on that recalced snapshot.
- **Partitions carry identity, not data**: the source's SHA-256 digest,
  the fully resolved recalc context (`now`, offset, seed, mode, profile —
  resolved once per read, so one job observes one logical instant), and
  the engine fingerprint. Executors read the source bytes once, hash that
  buffer, and recalc **the same buffer** — a workbook that changed since
  planning refuses by digest, and a task retry re-derives identical
  results instead of drifting.
- **Mixed-version fleets refuse**: an executor whose
  `zlsx.engine_fingerprint()` differs from the driver's raises rather
  than mixing results from two engine builds.
- The per-executor snapshot cache (keyed by digest + full context +
  fingerprint) is an optimization, never a correctness dependency.
- A workbook using a construct the engine refuses fails the read with
  `ZlsxFormulaRefusal` naming the cells — nothing silently falls back to
  stale caches.

### Streaming — Auto Loader for Excel

```python
stream = (spark.readStream.format("zlsx")
          .schema("region string, units bigint, revenue double")
          .option("startingPosition", "earliest")   # or "latest"
          .load("/Volumes/cat/schema/vol/landing/"))
```

Each workbook is ingested **exactly once** as it lands in the zone. The
checkpoint offset is a fingerprint map (`path → (mtime_ns, size)`), so a
restarted stream resumes from the offset alone. Semantics to know:

- Land files **atomically** (write elsewhere, then move in) — a half-written
  workbook fails the batch by name rather than ingesting garbage.
- Files are treated as **immutable**: a changed fingerprint re-ingests the
  whole workbook (the Auto Loader convention). Deletions are ignored.
- The offset grows with the file count — sized for a landing zone of
  thousands of workbooks, not millions.
- The batch read options marked **batch + streaming** above apply
  (`sheet`, `header`, `mode`, `rowsPerPartition`), plus the
  streaming-only `startingPosition` (`latest` skips files already present
  when the stream first starts). The `zlsx.recalc*` options do **not**:
  streaming + recalc is refused at option validation — recalculate in a
  batch job, then stream the result.

## Migration from openpyxl

### Reads

```python
# Before
from openpyxl import load_workbook
wb = load_workbook("data.xlsx", read_only=True, data_only=True)
for row in wb["Summary"].iter_rows(values_only=True):
    ...

# After
import zlsx
with zlsx.open("data.xlsx") as book:
    for row in book.sheet("Summary").rows():
        ...
```

Row shape is identical to openpyxl's `values_only=True` — a sequence of `None | bool | int | float | str`. zlsx yields `list` (not `tuple`) but anything that does `len(row)` and `row[i]` works unchanged.

### Writes

```python
# Before — openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
wb = Workbook()
ws = wb.active
ws.title = "Summary"
header = ws.cell(row=1, column=1, value="Name")
header.font = Font(bold=True, color="FFFFFFFF")
header.fill = PatternFill(fgColor="FF1E3A8A", patternType="solid")
header.alignment = Alignment(horizontal="center")
ws.column_dimensions["A"].width = 24
ws.freeze_panes = "A2"
wb.save("out.xlsx")

# After — zlsx
import zlsx
with zlsx.write("out.xlsx") as w:
    header = w.add_style(zlsx.Style(
        font_bold=True, font_color_argb=0xFFFFFFFF,
        fill_pattern="solid", fill_fg_argb=0xFF1E3A8A,
        alignment_horizontal="center",
    ))
    sheet = w.add_sheet("Summary")
    sheet.set_column_width(0, 24)
    sheet.freeze_panes(rows=1, cols=0)
    sheet.write_row(["Name"], styles=[header])
```

`zlsx.Style` is registered once and reused by index — no `cell.style = …` assignment per cell. Colours are `0xAARRGGBB` integers (openpyxl uses `"RRGGBB"` strings).

## Scope

**In**

- Read rows from any `.xlsx` / `.xlsm` — shared strings, inline strings, XML entities, UTF-8, numeric / boolean / error cells
- Write fresh workbooks with multiple sheets, typed cells, SST dedup, XML escaping
- Cell styles: fonts (bold / italic / size / name / color), horizontal alignment, wrap text, fills (19 patternTypes, fg + bg colors), borders (5 sides × 14 styles + diagonal up/down), number formats
- Per-sheet layout: column widths, freeze panes, auto-filter
- Merged cells, external + internal hyperlinks, comments
- Rich-text runs on write (`write_rich_row`)
- Load-modify-save via `zlsx.edit(path)` — append rows to existing sheets
  (numeric / int / float / bool / str cells), `set_cell` / `set_cells`,
  `doc_props` read + `strip_doc_props` scrub, atomic save, `save_to_buffer` /
  `Editor.from_bytes` for filesystem-less callers
- Structural edits (0.9.0+): `insert_row` / `delete_row` / `insert_column` /
  `delete_column` / `add_sheet` / `rename_sheet` / `delete_sheet` /
  `rename_table_column`, with every cross-part rewriter (formulas in every
  dialect, defined names, hyperlinks, DV / CF, merges, panes, autoFilter,
  tables, drawings, comments, `<xm:f>` extensions, chart `<c:f>` series
  formulas, pivot locations and sources); what cannot be kept consistent
  refuses with a typed
  `ZlsxRefusal` — see *Structural edits & pivots*
- Pivot tables, typed read (0.9.0+): `Editor.pivots()` / `zlsx.pivots(path)`
  — the `zlsx pivots` records as dicts
- Defined names, typed read (0.9.0+): `Editor.defined_names()` /
  `zlsx.defined_names(path)` — the `zlsx defined-names` records as dicts
- Conditional formats, typed read (0.9.0+): `Editor.conditional_formats()` /
  `zlsx.conditional_formats(path)` — the `zlsx conditional-formats` records
  as dicts
- Image / chart anchors, typed read (0.9.0+): `Editor.anchors()` /
  `zlsx.anchors(path)` — the `zlsx anchors` records as dicts (geometry and
  part names; the image bytes and chart XML stay in the archive)
- Panes, `<dimension>` and calc properties, typed read (0.9.0+):
  `Editor.sheet_props()` / `zlsx.sheet_props(path)` — the `zlsx sheet-props`
  records as dicts (the extent and the first view's pane as authored, split
  panes included) — and `Editor.calc_props()` / `zlsx.calc_props(path)` — the
  one `zlsx calc-props` record as a dict
- Sheet visibility (0.9.0+): `Book.sheet_state(selector)` / `Sheet.state` —
  the `<sheet state>` attribute as `zlsx list-sheets` spells it (`visible` /
  `hidden` / `veryHidden`; a missing or unrecognised value reads `visible`);
  hidden sheets stay in `Book.sheets` and read like any other
- Formula text and error tags on read (0.9.0+): `Rows.formula_strings()` /
  `Rows.formula_refs()` / `Rows.error_strings()` — the `<f>` body
  (entity-decoded), a shared / array slave's base cell, and the `t="e"`
  literal, one list per accessor aligned to the row `next()` yielded (the
  `formula` / `formula_ref` / `v` of `zlsx cells`); the row itself keeps the
  cached value and the literal string, so nothing that read before reads
  differently. A formula whose cached value is an error is a formula
- Embeddings on write (0.9.0+): `Editor.set_embeddings(model, dim, coverages,
  dtype=...)` — the vector set `zlsx.embeddings(path)` reads, on the same
  `(rows, dim)` float32 / uint64 shape, replacing any previous set — and
  `Editor.embeddable_rows(sheet, range, column)`, the rows to embed with the
  canonical content hash the write stores (the `embed_row` records of
  `zlsx embed --extract`), and the two sweeps `Editor.prune_embeddings()` /
  `Editor.strip_embeddings()` (`zlsx embed --prune` / `--strip`); see
  *Embeddings*
- Formula cells on write (`write_row_with_formulas`) — emits `<f>` + cached `<v>`; pass `recalculate=RecalcOptions()` to `save()` and the cached values are computed by zlsx's own engine, or leave it off and Excel recalculates on open. `FormulaSpec.cse(text, ref)` authors legacy CSE rectangles
- Formula engine (0.8.0+): `Editor.recalculate` / `save_with_recalc` (atomic §5.7.9 transaction) / `evaluate` / `save_to_buffer` / `Editor.from_bytes` / `mark_recalc_on_load` — see *Recalculate & evaluate*
- Data validation (list / numeric / custom) and conditional formatting (cellIs / expression / colorScale / dataBar)
- Refcounted handles — close the book while rows are still being consumed, the C ABI keeps the state alive until the last reference drops
- PySpark Data Source (`zlsx.spark`) — batch read with per-(file×sheet) partitions and optional row-range splits, batch write to single-file or `part-*.xlsx` targets

**Out** (by design, or queued)

- `.xls` / `.xlsb` / `.ods` — never
- Formula evaluation on the *read* path — the reader still returns the cached `<v>` value byte-for-byte and never computes. Since 0.8.0 the engine lives behind the explicit `recalculate` / `evaluate` / `save_with_recalc` surface (see *Recalculate & evaluate*); a plain read remains exactly what Excel stored
- Pictures — image *payloads* and image authoring are Zig-only today (S5; the anchors themselves read through `Editor.anchors()`); charts follow structural edits (their series formulas ride the formula rewriter — the chart `<c:f>` sweep) and their anchors and series refs read through `Editor.anchors()`, with chart authoring deferred (D2 → S9); pivot *authoring* is S8. The per-surface truth is [`docs/plans/surface-matrix.md`](../../docs/plans/surface-matrix.md)

## Thread safety

Distinct `Book` and `Writer` handles are fully independent — call them freely from any threads. Operations on the same handle must be externally synchronized, same as sqlite3 or libcurl. The C ABI's refcount lets a row iterator outlive its Book handle safely; all other cross-thread sharing is the caller's responsibility.

Cancellable formula-engine calls (`recalculate`, `save_with_recalc`, `evaluate`, `Writer.save(recalculate=...)`) run their FFI call on a private worker thread while the calling thread waits interruptibly; the handle-synchronization rule above still applies — the worker is an implementation detail, not a license to share the handle.

## Lifetime gotchas

String slices returned by the reader (`row[i]` where `row[i]` is a `str`) point into buffers owned by the `Book`. The Python binding decodes to `str` on every access, so you don't see this directly — each iteration materialises a fresh list.

Writer-side styles allocate on first registration and stay pinned for the Writer's lifetime. Registering the same `Style` twice returns the same index (content-compared dedup, including `font_name` and `number_format` strings).

## License

Proprietary — see [LICENSE](../../LICENSE).

- A 60-day free evaluation applies to any person or organization — released wheels/binaries only, no source-code rights.
- All other use requires a commercial license (which includes source rights for this Python wrapper layer, not the Zig core).

For commercial licensing, email **laurent.fabre@gmail.com**. See the [parent repository LICENSE](../../LICENSE) and [CONTRIBUTING.md](../../CONTRIBUTING.md) for full terms.
