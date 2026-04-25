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
    sheet.freeze_panes(rows=1, cols=0)
    sheet.set_auto_filter("A1:C1")

    sheet.write_row(["Name", "Amount", "Share"], styles=[header, header, header])
    sheet.write_row(["Alice", 12345.67, 0.42], styles=[0, money, pct])
    sheet.write_row(["Bob",    9876.54, 0.33], styles=[0, money, pct])
# save happens automatically on clean exit; exception → no save
```

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
- Merged cells, hyperlinks, comments
- Rich-text runs on write (`write_rich_row`)
- Data validation (list / numeric / custom) and conditional formatting (cellIs / expression / colorScale / dataBar)
- Refcounted handles — close the book while rows are still being consumed, the C ABI keeps the state alive until the last reference drops

**Out** (by design, or queued)

- `.xls` / `.xlsb` / `.ods` — never
- Formula evaluation — never (the reader returns the cached `<v>` value Excel stored; zlsx never runs a formula engine)
- Formula cells on write — Python binding doesn't expose `write_row_with_formulas` yet (the Zig writer ships it; binding is queued)
- Load → modify → save round-trip — Phase 3c queued
- Pictures / charts / pivots — out of scope

## Thread safety

Distinct `Book` and `Writer` handles are fully independent — call them freely from any threads. Operations on the same handle must be externally synchronized, same as sqlite3 or libcurl. The C ABI's refcount lets a row iterator outlive its Book handle safely; all other cross-thread sharing is the caller's responsibility.

## Lifetime gotchas

String slices returned by the reader (`row[i]` where `row[i]` is a `str`) point into buffers owned by the `Book`. The Python binding decodes to `str` on every access, so you don't see this directly — each iteration materialises a fresh list.

Writer-side styles allocate on first registration and stay pinned for the Writer's lifetime. Registering the same `Style` twice returns the same index (content-compared dedup, including `font_name` and `number_format` strings).

## License

MIT — see [LICENSE](../../LICENSE) in the parent repository.
