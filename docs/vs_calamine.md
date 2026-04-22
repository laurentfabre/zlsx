# `zlsx` vs `calamine` — feature gap analysis

`calamine` (Rust, [docs.rs/calamine/0.26.1](https://docs.rs/calamine/0.26.1/)) is the reference pure-native xlsx reader in the ecosystem. This page inventories the gap so you can choose consciously.

> **TL;DR**: zlsx is a narrower tool — xlsx-only, no formula evaluation, "just rows and cells plus styled writes" — deliberately limited to keep the core compact and fast ([10.7 ms / 4.2 MB on a 1,008-row xlsx vs calamine's 15.3 ms / 4.9 MB](benchmarks.md)). Since v0.2.4, zlsx also ships a pragmatic openpyxl-parity **writer** — something calamine doesn't offer at all. If your use case needs `.xls`/`.xlsb`/`.ods`, native `DateTime`, formula evaluation, defined names, or serde deserialization, use calamine. For reading+writing xlsx in Zig (or via C/Python), zlsx is the complete option.

## Supported file formats

| | calamine | zlsx |
|---|---|---|
| `.xlsx` | ✓ | ✓ |
| `.xlsm` (macro-enabled) | ✓ | ✓ *(same container; zlsx ignores the macro blob)* |
| `.xlsb` (binary OOXML) | ✓ | ✗ |
| `.xls` (legacy OLE/BIFF) | ✓ | ✗ |
| `.ods` (OpenDocument) | ✓ | ✗ |
| Auto-detect by content | ✓ (`open_workbook_auto`) | ✗ *(not needed — xlsx-only)* |

## Cell value types

Calamine's `Data` enum vs zlsx's `Cell` union:

| Calamine `Data` variant | zlsx `Cell` variant | Notes |
|---|---|---|
| `Empty` | `.empty` | identical |
| `String(String)` | `.string: []const u8` | calamine clones; zlsx borrows from xml when safe |
| `Int(i64)` | `.integer: i64` | identical; zlsx also prefers integer over float when the value parses cleanly |
| `Float(f64)` | `.number: f64` | identical |
| `Bool(bool)` | `.boolean: bool` | identical |
| `DateTime(ExcelDateTime)` | — | zlsx returns the raw serial number via `.number`; no DateTime decoding |
| `DateTimeIso(String)` | `.string` | ISO-formatted datetimes come back as strings in zlsx |
| `DurationIso(String)` | `.string` | same |
| `Error(CellErrorType)` | `.string` | zlsx returns the error code text (`#N/A`, `#REF!`, …) as a string; no typed enum |

**Gap to close (easy)**: a `Cell.error: ErrorCode` variant. The XML already carries `t="e"` and a string in `<v>`; we'd just map it. Hasn't been worth the surface area yet.

**Gap to close (hard)**: native `DateTime`. Requires parsing Excel's serial-number format **and** applying the number-format string from `styles.xml` to know whether `45000` is a date or a count. Non-trivial.

## Core API surface

| Concept | calamine | zlsx |
|---|---|---|
| Open workbook | `open_workbook(path)` / `open_workbook_auto(path)` | `Book.open(alloc, path)` |
| List sheets | `reader.sheet_names()` | `book.sheets: []Sheet` |
| Read a sheet | `reader.worksheet_range(name) → Range<Data>` | `book.rows(sheet, alloc) → Rows` |
| Nth sheet | `reader.worksheet_range_at(i)` | `book.sheets[i]` |
| Iterate rows | `range.rows() → impl Iterator<&[Data]>` | `rows.next() !?[]const Cell` |
| Iterate cells | `range.cells() / used_cells()` | *(iterate `rows`, then nested `for`; no cell-level iterator)* |
| Absolute cell `get(A5)` | `range.get_value((row, col))` | *(caller resolves column letter → index, peeks `cells[col]`)* |
| Sub-range | `range.range(start, end)` | ✗ |
| Range deserialize (serde) | `range.deserialize::<T>()` | ✗ *(N/A in Zig)* |
| Row / column counts | `range.height()` / `range.width()` / `get_size()` | *(not exposed; caller counts)* |
| First-row headers | `range.headers() → Option<Vec<String>>` | ✗ *(caller reads first row manually)* |

**Design difference that matters**: calamine **materializes** the whole sheet into a dense `Range<Data>` up front, so `get_value((r, c))` is O(1). zlsx **streams** rows via a state machine — you pay memory only for the current row's cells, but can't random-access. For the typical "iterate and ingest" workload the streaming shape wins on RSS.

## Metadata & advanced xlsx features

| Feature | calamine | zlsx |
|---|---|---|
| Sheet metadata (name, type, visibility) | `reader.sheets_metadata() → Vec<SheetMetadata>` | ✗ *(only name + path exposed)* |
| Sheet visibility (hidden/visible) | ✓ (`SheetVisible` enum) | ✗ |
| Sheet type (worksheet/chartsheet/dialog) | ✓ (`SheetType` enum) | ✗ *(zlsx indexes all `xl/worksheets/*.xml` indiscriminately)* |
| Defined names (named ranges) | `reader.defined_names() → &[(String, String)]` | ✗ |
| Formulas (not cached values) | `reader.worksheet_formula(name) → Range<String>` | ✗ *(zlsx reads `<v>`; the `<f>` child is skipped)* |
| Cached formula result | ✓ via `Data` in the range | ✓ via `Cell` (same `<v>` we all read) |
| Merged cell regions | ✗ *(per the 0.26 Reader trait, merged regions are not exposed — they're in XML but calamine currently doesn't surface them in its stable API)* | ✗ |
| Pictures / embedded images | `reader.pictures() → Option<Vec<(String, Vec<u8>)>>` | ✗ |
| VBA project | `reader.vba_project() → Option<Cow<VbaProject>>` | ✗ |
| Workbook / sheet protection | ✗ *(file-level password required before read)* | ✗ |
| Styles / number formats | ✗ in public API *(xlsx reader skips them)* | **✓ on write** *(bold/italic, font name/size/color, alignment, wrap, fills, borders, number formats — Phase 3b)*; ✗ on read (reader ignores xl/styles.xml) |
| Rich-text formatting (bold runs etc.) | ✗ *(returns the concatenated text, no spans)* | ✗ on read *(concatenates `<t>` runs)*; ✗ on write *(plain runs only — Phase 3b covers styling, not per-run formatting inside a cell)* |

## Writing (Phase 3b, shipped in v0.2.4)

| | calamine | zlsx |
|---|---|---|
| Write xlsx | ✗ *(calamine is read-only; the ecosystem's writer is `rust_xlsxwriter`)* | **✓** |
| Strings / numbers / booleans / empties | — | ✓ |
| Shared-string table with dedup | — | ✓ |
| Multi-sheet workbooks | — | ✓ |
| Fonts (bold, italic, size, name, color) | — | ✓ |
| Alignment (8 horizontal values) + wrap text | — | ✓ |
| Fills (19 OOXML patternTypes, fg + bg ARGB) | — | ✓ |
| Borders (5 sides × 14 styles, per-side colour, diagonal up/down) | — | ✓ |
| Custom number formats (numFmtId ≥ 164) | — | ✓ |
| Column widths (per-column override) | — | ✓ |
| Freeze panes (top rows + left cols) | — | ✓ |
| Auto-filter (A1-style range) | — | ✓ |
| Load → modify → save round-trip | — | ✗ *(Phase 3c queued — preserves only what zlsx parses today)* |
| Formulas (`<f>` emission on write) | — | ✗ *(explicitly out of scope; writer emits cached values only)* |
| Pictures / charts / pivots | — | ✗ *(out of scope)* |

## Ownership model

| | calamine | zlsx |
|---|---|---|
| String ownership | every `Data::String` clones into an owned `String` | borrow from source xml when the run is single-`<t>` and entity-free; own only when rich-text concatenation or entity decoding forces it |
| Whole-sheet materialization | yes — `Range<Data>` holds every cell for the sheet | no — `Rows.next()` yields one row at a time; RSS is O(single-row width) + sharedStrings.xml |
| Allocator | Rust global (`alloc::alloc`) | caller-provided, tested with `std.testing.allocator` leak detector |

## Platform & ecosystem

| | calamine | zlsx |
|---|---|---|
| Language | Rust | Zig 0.15 |
| Dependencies | 20+ transitive (serde, encoding, quick-xml, zip, log, …) | 0 — stdlib only |
| Binary size contribution | ~620 KB to a release binary | ~120 KB |
| Error types | `calamine::Error` (per-format variant) | inferred error set (`!Book`, `!Rows`, etc.) |
| Serde / deserialization | first-class (`deserialize_as_*`, `RangeDeserializerBuilder`) | N/A |
| Streaming writer bindings | in Python via `python-calamine` | N/A |

## Fuzzing

| | calamine | zlsx |
|---|---|---|
| In-repo fuzz harness | none committed | 14 targets (`fuzz parse*`, `fuzz Rows.next`, `fuzz Book.open`), random-byte + mutation, PRNG-driven; tested to 14 M inputs |
| OSS-Fuzz integration | yes (project `calamine`) | not yet |

## When to pick which

**Pick zlsx if**:
- You're writing Zig and want zero third-party deps.
- Your input is xlsx-only and you only need: rows, typed cells, UTF-8.
- RSS matters — embedded system, Lambda, or many parallel parses.
- You want a single-file library you can vendor.

**Pick calamine (or `python-calamine` from Python) if**:
- You need `.xls` / `.xlsb` / `.ods`.
- You need native `DateTime` / `DurationIso` / typed error cells.
- You need formulas, defined names, pictures, or VBA.
- You want the serde pipeline (`range.deserialize::<MyStruct>()`).
- Random-access into the sheet (`get_value((r, c))`) is in your hot path.

## Roadmap — what zlsx might add without losing its shape

In rough priority order, driven by real Alfred-style use cases, not feature parity for parity's sake:

1. **Typed error cells** (`Cell.err: ErrorCode`) — trivial, already exposed in XML. Low risk, under 20 lines.
2. **Row/col dimension helpers** on `Sheet` (`max_col`, `row_count` read from the sheet's `<dimension ref="A1:AI1008"/>` attribute). Cheap, useful for preallocation.
3. **`sheetByIndex(i)`** convenience — one-liner.
4. **Named-range lookup** (`book.definedName("Foo")`) — read from `workbook.xml`'s `<definedNames>`, return a `(sheet, start, end)` tuple. Useful for templated ingest.
5. **ISO-8601 date awareness** — detect `ExcelDateTime` from number-format code in styles.xml; optionally return `.datetime: i64` epoch-seconds. Heavier work; consider a separate optional module.
6. **Merged-region metadata** — a `Sheet.mergedRegions() → []Region` that reads `<mergeCell ref="A1:B3"/>`. Read-only, no interpretation.

Out of scope for zlsx intentionally: `.xls`/`.xlsb`/`.ods`, writing, pictures, VBA, style application. Those would compromise the "single-file, stdlib-only, xlsx-only" positioning that makes zlsx worth existing next to calamine.
