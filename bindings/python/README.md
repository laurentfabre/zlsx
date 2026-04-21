# py-zlsx

Python binding for [zlsx](https://github.com/laurentfabre/zlsx) — a fast, read-mostly `.xlsx` library written in Zig.

This package ships Python bindings as a thin `ctypes` layer over `libzlsx` (no Rust, no PyO3, no third-party runtime deps). Benchmark context: zlsx reads a 1,008-row workbook in **10.7 ms / 4.2 MB RSS** — 4× faster than `python-calamine`, 24× faster than `openpyxl`, at a tenth of the memory.

## Install

The wheel bundles a per-platform `libzlsx.{dylib,so,dll}` so `pip install py-zlsx` is self-contained. If you're running from source, point `ZLSX_LIBRARY` at the shared library you have on disk — or install the Homebrew bottle (`brew install laurentfabre/zlsx/zlsx`) and the package will find it automatically at `/opt/homebrew/opt/zlsx/lib/libzlsx.dylib`.

```bash
# (Preferred) from PyPI
pip install py-zlsx

# From source (requires a libzlsx on disk)
pip install -e ./bindings/python
export ZLSX_LIBRARY=/path/to/libzlsx.dylib   # optional — only if auto-discovery fails
```

## Quick start

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

## Migration from openpyxl

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

The returned row shape is the same as openpyxl's `values_only=True` — a tuple-like sequence of values. zlsx yields `list`, not `tuple`, but anything that does `len(row)` and `row[i]` works unchanged.

## Scope

- **In**: reading rows from any `.xlsx` / `.xlsm` (shared strings, inline strings, XML entities, UTF-8, numeric / boolean / error cells), sheet enumeration, refcounted lifetimes (close the book while rows are still being consumed — the C ABI handles it).
- **Out**: writing (the Zig library has a write API but the Python binding exposes reads only in this version), `.xls` / `.xlsb` / `.ods`, formulas, styles, merged-cell semantics.

## Thread safety

Distinct `Book` handles are fully independent — call them freely from any threads. Operations on the same handle must be externally synchronized, same as sqlite3 or libcurl.

## License

MIT — see [LICENSE](../../LICENSE) in the parent repository.
