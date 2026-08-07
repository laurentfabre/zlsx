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

    for row in book.sheet(0).rows():
        # row is a list; cell types map to Python:
        #   empty → None    string → str
        #   integer → int   number → float   boolean → bool
        print(row)

    summary = book.sheet("Summary")          # by name also works
    header = next(summary.rows())
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

The writer produces fresh workbooks (load-modify-save round-trip lands in Phase 3c). Cell styles registered via `Writer.add_style` get a 1-based index; pass those indices alongside values in `write_row(styles=[…])`.

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

libzlsx 0.9.0+ ships the formula engine across the C ABI. The binding
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
- Append-only load-modify-save via `zlsx.edit(path)` — append rows
  to existing sheets (numeric / int / float / bool / str cells)
  with atomic save
- Formula cells on write (`write_row_with_formulas`) — emits `<f>` + cached `<v>`; pass `recalculate=RecalcOptions()` to `save()` and the cached values are computed by zlsx's own engine, or leave it off and Excel recalculates on open. `FormulaSpec.cse(text, ref)` authors legacy CSE rectangles
- Formula engine (0.9.0+): `Editor.recalculate` / `save_with_recalc` (atomic §5.7.9 transaction) / `evaluate` / `save_to_buffer` / `Editor.from_bytes` / `mark_recalc_on_load` — see *Recalculate & evaluate*
- Data validation (list / numeric / custom) and conditional formatting (cellIs / expression / colorScale / dataBar)
- Refcounted handles — close the book while rows are still being consumed, the C ABI keeps the state alive until the last reference drops
- PySpark Data Source (`zlsx.spark`) — batch read with per-(file×sheet) partitions and optional row-range splits, batch write to single-file or `part-*.xlsx` targets

**Out** (by design, or queued)

- `.xls` / `.xlsb` / `.ods` — never
- Formula evaluation on the *read* path — the reader still returns the cached `<v>` value byte-for-byte and never computes. Since 0.9.0 the engine lives behind the explicit `recalculate` / `evaluate` / `save_with_recalc` surface (see *Recalculate & evaluate*); a plain read remains exactly what Excel stored
- Cell mutate / structural edits (insert column, delete row, etc.) on existing workbooks — append-only is shipped via `zlsx.edit(path)`; full round-trip is its own follow-up plan
- Pictures / charts / pivots — out of scope

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
