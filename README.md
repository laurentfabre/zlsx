# zlsx

Tiny `.xlsx` reader **and** writer for Zig. Single-file library, no third-party deps — just `std.zip` + `std.compress.flate` (for reads) + an in-house LZ77 + dynamic-huffman deflate compressor with lazy matching (for writes, since Zig 0.15.2's `std.compress.flate.Compress` is still mid-refactor) + a hand-rolled XML walker scoped to what spreadsheets actually need. Ships with a CLI, a C ABI, and Python bindings.

**Reader**: 1.6-2.0 ms on small files, **3.2 ms / 2.27 MB** on a 67 KB workbook with 1,144 shared strings — **1.20× faster than calamine-rust**, 7.5× faster than python-calamine, 41× faster than openpyxl on that file, at the **smallest RSS of the four** (~7.5-19× lower than the Python stack, ~1.4× lower than calamine). Native-speed tier on every corpus file. [Full benchmark table →](docs/benchmarks.md)

**Writer** (Phase 3b, v0.2.4): pragmatic openpyxl-parity styles — bold/italic, font size/name/color, horizontal alignment, wrap text, 19 OOXML fill patterns, 14 border styles × 5 sides, custom number formats, column widths, row heights, freeze panes, auto-filter, merged cell ranges, external + internal hyperlinks, full data-validation family (list / whole / decimal / date / time / textLength / custom with 8 comparison operators), formulas with cached values. Survived 1M-iter deep fuzz on every surface.

```zig
const xlsx = @import("zlsx");

// Reading
var book = try xlsx.Book.open(allocator, "in.xlsx");
defer book.deinit();
const sheet = book.sheetByName("Summary") orelse return error.MissingSheet;
var rows = try book.rows(sheet, allocator);
defer rows.deinit();
while (try rows.next()) |cells| {
    for (cells) |c| switch (c) {
        .empty   => {},
        .string  => |s| std.debug.print("str  {s}\n", .{s}),
        .integer => |i| std.debug.print("int  {d}\n", .{i}),
        .number  => |f| std.debug.print("num  {d}\n", .{f}),
        .boolean => |b| std.debug.print("bool {any}\n", .{b}),
    };
}

// Lazy open (iter54) — don't pre-decompress every sheet's XML up
// front. Use this when streaming one sheet at a time out of a
// large workbook, e.g. the upcoming jq-for-excel CLI. The file
// handle stays open for the Book's lifetime; Book.open is the
// eager facade that releases it on return.
var lazy = try xlsx.Book.openLazy(allocator, "big.xlsx");
defer lazy.deinit();
// Workbook-wide state (sheets list, SST, styles, theme) is ready
// on return. Per-sheet metadata (merged ranges, hyperlinks,
// validations, comments) materializes on demand — call
// streamSheet/rows or preloadSheet explicitly.
var r = try lazy.streamSheet(0, allocator);   // by index — CLI-friendly
defer r.deinit();
while (try r.next()) |_| {}
try lazy.preloadSheet(lazy.sheets[1]);        // populate metadata only

// Writing — including styles, fills, borders, number formats,
// column widths, freeze panes, auto-filter.
var w = xlsx.Writer.init(allocator);
defer w.deinit();

const header = try w.addStyle(.{
    .font_bold = true,
    .font_color_argb = 0xFFFFFFFF,
    .fill_pattern = .solid,
    .fill_fg_argb = 0xFF1E3A8A,
    .alignment_horizontal = .center,
    .border_bottom = .{ .style = .thin, .color_argb = 0xFF000000 },
});
const money = try w.addStyle(.{ .number_format = "$#,##0.00" });

var sheet_w = try w.addSheet("Summary");
try sheet_w.setColumnWidth(0, 24);
sheet_w.freezePanes(1, 0);
try sheet_w.setAutoFilter("A1:C1");
// For merged header spans: `try sheet_w.addMergedCell("A3:C3");`
try sheet_w.writeRowStyled(
    &.{ .{ .string = "Name" }, .{ .string = "Amount" }, .{ .string = "Active" } },
    &.{ header, header, header },
);
try sheet_w.writeRowStyled(
    &.{ .{ .string = "Alice" }, .{ .number = 12345.67 }, .{ .boolean = true } },
    &.{ 0, money, 0 },
);
try w.save("out.xlsx");
```

## Why

Zig's ecosystem has no official xlsx library yet. The options are [calamine](https://github.com/tafia/calamine) via FFI (Rust, read-only), Apache POI via subprocess (JVM), or rolling your own on top of `std.zip`. zlsx is the third option, packaged up, fuzz-tested (14 M inputs per release), and extended with a pragmatic write path so you don't need a second library for output.

Designed for a real use case: Alfred's hotel-concierge pipeline reads a 1,008-row × 35-col openpyxl-generated workbook with inline strings, shared strings, and UTF-8 content across multiple languages. `zlsx` reads it in ~10 ms and writes styled output without shipping openpyxl.

## What's in, what's out

**In**
- **Read** workbooks — shared strings (with rich-text runs + XML entities), inline strings, numeric / boolean / error / formula-cached cells, UTF-8 throughout, merged-cell ranges via `Book.mergedRanges(sheet)`, external-URL hyperlinks via `Book.hyperlinks(sheet)` (resolved through sheet `_rels`), all data validations via `Book.dataValidations(sheet)` — dropdowns (values entity-decoded), plus `kind` / `op` / `formula1` / `formula2` on numeric, date, time, text-length, and custom variants, rich-text run formatting (bold / italic / color / size / font) via `Book.richRuns(sst_idx)` for entries that used `<r>` wrappers, per-cell style indices via `Rows.styleIndices()` resolving to number-format codes via `Book.numberFormat(style_idx)` (date detection via `Book.isDateFormat(style_idx)`), cell comments via `Book.comments(sheet)` (author + plain-text body, decoded)
- **Write** workbooks — strings (SST-deduped), integers, numbers, booleans, empties, multi-sheet; cell styles with fonts, fills, borders, alignment, wrap, number formats; per-sheet column widths, row heights, freeze panes, auto-filter, merged cell ranges, external-URL hyperlinks (per-sheet `_rels`), internal hyperlinks (`location="Sheet2!A1"`), list-type data validations (dropdowns), number / decimal / date / time / text-length / custom data validations, formulas with optional cached value, rich-text cells (per-run bold / italic / color / size / font) via `SheetWriter.writeRichRow`, cell comments (notes) via `SheetWriter.addComment` (emits the full `xl/comments{N}.xml` + `vmlDrawing{N}.vml` + rels stack so Excel renders the yellow indicator), conditional formatting via `SheetWriter.addConditionalFormatCellIs` / `…Expression` (two core cfRule types) referencing differential formats registered via `Writer.addDxf`
- XML entity decoding (`&amp;`, `&lt;`, `&gt;`, `&quot;`, `&apos;`, `&#N;`, `&#xN;`) on read and escaping on write
- CLI (`zlsx {rows|cells|meta|list-sheets|comments|validations|hyperlinks|styles|sst} <file>`) emitting uniform NDJSON envelopes, with row/range/header/skip-take/all-sheets/sheet-glob/include-blanks/with-styles/output-mode flags and SIGPIPE-clean pipelines — see "CLI" below and [`docs/jq-for-excel.md`](docs/jq-for-excel.md)); C ABI (`libzlsx.{dylib,so,dll}` + `include/zlsx.h`); Python bindings (`pip install py-zlsx`)

**Out (by design)**
- No formula evaluation — the reader returns the cached `<v>` value; the writer accepts formula text + an optional cached result via `writeRowWithFormulas` but never computes the formula itself
- No automatic date decoding — dates surface as their raw Excel serial number via `.number`; a convenience helper `xlsx.fromExcelSerial(cell.number) -> ?DateTime` handles the 1900-03-01 through 9999-12-31 range (serials ≤ 60 return `null` because the Excel 1900 leap-year bug makes them ambiguous)
- No load-modify-save round-trip yet — Phase 3c queued. For now the writer only produces fresh workbooks
- No chart / pivot / image extraction or emission

## Feature matrix

How zlsx's current surface compares against the popular xlsx libraries. `✓` = first-class API; `helper` = exposed but caller-driven; `~` = partial / limited; `—` = not implemented; `?` = I'm not confident enough to claim one way or the other — PRs / corrections welcome.

### Reader capability

| Capability | **zlsx** | calamine-rust 0.26 | openpyxl 3.1 | python-calamine 0.6 |
|---|---|---|---|---|
| Shared strings (SST) | ✓ | ✓ | ✓ | ✓ |
| Inline strings | ✓ | ✓ | ✓ | ✓ |
| Rich-text runs concatenated | ✓ | ✓ | ✓ | ✓ |
| Numeric / integer / float split | ✓ | ~¹ | ✓ | ~¹ |
| Boolean / error cells | ✓ | ✓ | ✓ | ✓ |
| Formula cached value | ✓ | ✓ | ✓ | ✓ |
| Date as `DateTime` | ✓⁴ | ✓ | ✓ | ✓ |
| XML entity decoding | ✓ | ✓ | ✓ | ✓ |
| Merged cell ranges | ✓ | ✓ | ✓ | ✓ |
| External-URL hyperlinks | ✓ | ? | ✓ | ? |
| Data validations (list / dropdown) | ✓ | — | ✓ | — |
| Data validations (number / date / custom) | ✓ | — | ✓ | — |
| Rich-text formatting preserved | ✓³ | ~ | ✓ | — |
| Cell styles on read (bold / colour / fill / border) | ✓⁵ | — | ✓ | — |
| Comments / notes | ✓⁶ | ? | ✓ | — |
| Chart / image / pivot access | — | — | ~ | — |
| Load-modify-save | — | — | ✓ | — |

¹ Returns a single `Float` type for any non-text number — callers cast to integer if needed.
² `xlsx.fromExcelSerial(cell.number) -> ?DateTime`; out-of-range serials (1900 leap-bug window) return `null`.
³ `Book.richRuns(sst_idx)` surfaces per-`<r>` bold / italic + ARGB color / size / font name. Theme colors (`<color theme="N"/>`) are resolved via the workbook's `xl/theme/theme1.xml` palette (iter52); `<color indexed="N"/>` and `tint` are still not resolved.
⁴ `Rows.parseDate(col_idx)` combines style-lookup + date-format detection + serial decoding into one call. Returns `?DateTime` (or `datetime.datetime | None` in Python). The low-level chain (`styleIndices()` + `isDateFormat()` + `fromExcelSerial()`) is still exposed for callers that need the individual pieces.
⁵ `Book.cellFont(style_idx)` surfaces bold / italic / ARGB color / size / font name; `Book.cellFill(style_idx)` surfaces `patternType` + fg / bg ARGB; `Book.cellBorder(style_idx)` surfaces `style` + color per side (left / right / top / bottom / diagonal). Theme colors (`theme="N"`) are resolved via the `xl/theme/theme1.xml` palette (iter52). `indexed="N"` (legacy palette) and `tint` modifiers are still not resolved.
⁶ `Book.comments(sheet)` returns `{top_left, author, text, runs}` for every `<comment>` under `<commentList>`. `text` is always the concatenated plain-text form; `runs` is populated (per-run bold/italic/color/size/font_name) when the source body used `<r><rPr>` formatting, null otherwise.

### Writer capability

| Capability | **zlsx** | xlsxwriter 3.2 | openpyxl 3.1 |
|---|---|---|---|
| Multi-sheet | ✓ | ✓ | ✓ |
| SST-deduped strings | ✓ | ✓ | ✓ |
| Integer / float / bool / empty | ✓ | ✓ | ✓ |
| Font (bold, italic, size, name, colour) | ✓ | ✓ | ✓ |
| Fills (19 OOXML patterns) | ✓ | ✓ | ✓ |
| Borders (14 styles × 5 sides) | ✓ | ✓ | ✓ |
| Horizontal alignment / wrap text | ✓ | ✓ | ✓ |
| Custom number formats | ✓ | ✓ | ✓ |
| Column widths | ✓ | ✓ | ✓ |
| Row heights | ✓ | ✓ | ✓ |
| Freeze panes | ✓ | ✓ | ✓ |
| Auto-filter | ✓ | ✓ | ✓ |
| Merged cell ranges | ✓ | ✓ | ✓ |
| External-URL hyperlinks | ✓ | ✓ | ✓ |
| Internal (`Sheet!A1`) hyperlinks | ✓ | ✓ | ✓ |
| Data validations (list) | ✓ | ✓ | ✓ |
| Data validations (number / date / custom) | ✓ | ✓ | ✓ |
| Conditional formatting | ✓⁸ | ✓ | ✓ |
| Cell comments / notes | ✓ | ✓ | ✓ |
| Formulas (with cached value) | ✓ | ✓ | ✓ |
| Rich-text runs per cell | ✓ | ✓ | ✓ |
| Images (PNG / JPEG embed) | — | ✓ | ~ |
| Charts | — | ✓ | ~ |
| Deflate compression | ✓ | ✓ | ✓ |
| Load-modify-save | — | — | ✓ |
| Sheet-name validation (length / reserved chars / duplicates) | ✓ | ~⁷ | ~ |

⁷ xlsxwriter validates length and some chars but does not reject case-insensitive duplicates up front.
⁸ `SheetWriter.addConditionalFormatCellIs` / `…Expression` / `…ColorScale` / `…DataBar` cover the four most-used rule types. Other variants (iconSet / top10 / aboveAverage / duplicateValues) can layer on without breaking the `CfRule` union. Differential formats (`addDxf`) support bold / italic / font size / font color / fill color / per-side borders.

### Language / packaging

| Axis | **zlsx** | calamine-rust | openpyxl | xlsxwriter | python-calamine |
|---|---|---|---|---|---|
| Native language | Zig | Rust | Python | Python | Rust (via PyO3) |
| First-class Zig API | ✓ | — | — | — | — |
| C ABI + header | ✓ | — | — | — | — |
| Python bindings | ✓ (`py-zlsx`, ctypes over C ABI) | — (use python-calamine) | ✓ (native) | ✓ (native) | ✓ (native) |
| CLI | ✓ (read-side) | — | — | — | — |
| Third-party runtime deps | 0 (stdlib only) | ~5 Rust crates | 0 Python deps | 0 Python deps | 0 Python deps |
| Static-link-friendly binary | ✓ | ✓ | — | — | — |
| License | MIT | MIT / Apache-2 | MIT | BSD-3 | MIT |

## Install

### CLI binary

Prebuilt binaries for macOS (ARM64, Intel), Linux (x86_64, ARM64 — static musl), and Windows (x86_64) ship with every tagged release:

```bash
# macOS (Apple Silicon)
curl -fsSL -o zlsx.tar.gz "https://github.com/laurentfabre/zlsx/releases/latest/download/zlsx-0.2.4-aarch64-apple-darwin.tar.gz"
tar -xzf zlsx.tar.gz && sudo mv zlsx-*/bin/zlsx /usr/local/bin/

# Via Homebrew (once the tap is published)
brew tap laurentfabre/zlsx
brew install zlsx
```

Each release tarball also bundles `lib/libzlsx.{dylib,so,a}` and `include/zlsx.h` for C consumers.

### As a Zig dependency

`zlsx` is a plain Zig module. Add it to your `build.zig.zon`:

```zig
.dependencies = .{
    .zlsx = .{ .url = "https://github.com/laurentfabre/zlsx/archive/refs/heads/main.tar.gz" },
},
```

Then in your `build.zig`:

```zig
const zlsx = b.dependency("zlsx", .{ .target = target, .optimize = optimize });
exe.root_module.addImport("zlsx", zlsx.module("zlsx"));
```

Or, for local development, clone the repo next to your project and use a path dependency:

```zig
.dependencies = .{
    .zlsx = .{ .path = "../zlsx" },
},
```

## CLI

A thin binary `zlsx` ships with this repo — shell-friendly, pipeable, openpyxl-replacement speed, no Python interpreter floor. "jq for Excel": every sub-command emits a uniform NDJSON envelope that composes cleanly with `jq`, `rg`, `awk`, `duckdb read_ndjson`, or an LLM ingest harness. Full design in [`docs/jq-for-excel.md`](docs/jq-for-excel.md).

```bash
zig build -Doptimize=ReleaseFast
./zig-out/bin/zlsx file.xlsx                          # default: rows sub-command
./zig-out/bin/zlsx rows file.xlsx                     # explicit alias
./zig-out/bin/zlsx cells file.xlsx                    # per-cell NDJSON
./zig-out/bin/zlsx meta file.xlsx                     # workbook + sheet metadata
./zig-out/bin/zlsx list-sheets file.xlsx              # NDJSON sheet list
./zig-out/bin/zlsx comments file.xlsx                 # cell comments
./zig-out/bin/zlsx validations file.xlsx              # data validations
./zig-out/bin/zlsx hyperlinks file.xlsx               # external + internal hyperlinks
./zig-out/bin/zlsx styles file.xlsx                   # cell-XF style table
./zig-out/bin/zlsx sst file.xlsx                      # shared-strings table
```

### Sub-commands

| Command | `kind` | Per-line fields |
|---|---|---|
| `zlsx rows <file>` | `"row"` | `sheet, sheet_idx, row, cells[]` (each `{ref, col, t, v}`) |
| `zlsx cells <file>` | `"cell"` | `sheet, sheet_idx, ref, row, col, t, v, style?` |
| `zlsx comments <file>` | `"comment"` | `sheet, sheet_idx, ref, row, col, author, text, runs?` |
| `zlsx validations <file>` | `"validation"` | `sheet, sheet_idx, range, rule_type, op?, formula1, formula2?, values?` |
| `zlsx hyperlinks <file>` | `"hyperlink"` | `sheet, sheet_idx, range, url?, location?` |
| `zlsx styles <file>` | `"style"` | `idx, font, fill, border, num_fmt` (workbook-wide) |
| `zlsx sst <file>` | `"sst"` | `idx, text, runs?` (workbook-wide) |
| `zlsx meta <file>` | `"workbook"` + `"sheet"` | workbook record first, then per-sheet records |
| `zlsx list-sheets <file>` | `"sheet"` | `sheet, sheet_idx` — lighter-weight than `meta` |

### Default NDJSON row envelope

```jsonl
{"kind":"row","sheet":"Data","sheet_idx":0,"row":1,"cells":[{"ref":"A1","col":1,"t":"str","v":"name"},{"ref":"B1","col":2,"t":"str","v":"qty"}]}
{"kind":"row","sheet":"Data","sheet_idx":0,"row":2,"cells":[{"ref":"A2","col":1,"t":"str","v":"apple"},{"ref":"B2","col":2,"t":"int","v":3}]}
```

`t` ∈ `"str"` | `"int"` | `"num"` | `"bool"`; empty cells skipped from the `cells` array.

### Flags

**Sheet selection** — mutually exclusive:

```bash
--sheet N                     # 0-indexed
--name "Summary"              # by name
--all-sheets                  # every sheet concatenated
--sheet-glob 'Data*'          # glob on sheet name (UTF-8 `?` matches one codepoint)
```

**Row / cell bounds** (rows / cells / comments):

```bash
--start-row 2 --end-row 100   # 1-based inclusive, per sheet
--range B2:Z100               # A1 bounding rectangle (rows + cells only)
--header                      # rows only: promote first row to keys, emit `fields` dict
```

**Stream pagination** — applies globally across the concatenated output:

```bash
--skip N --take M
```

**Cell metadata opt-ins** (cells / rows):

```bash
--include-blanks              # emit t:"blank" records for empty cells
--with-styles                 # attach terse style: {bold?, italic?, fg?, bg?, nf?, border?}
```

**Output modes**:

```bash
--output ndjson               # default: invariant-envelope stream
--output compact-ndjson       # sheet-prologue variant (drops sheet/sheet_idx on data records)
--output pretty-json          # valid only on `meta`: single pretty-printed JSON object
```

**Row format (rows sub-command only)** — legacy escape hatches:

```bash
--format jsonl                # default: envelope
--format legacy-jsonl         # pre-iter55a bare arrays
--format legacy-jsonl-dict    # pre-iter55a bare objects (old `--format jsonl-dict` is a deprecated alias)
--format tsv                  # tab-separated, \N for empty
--format csv                  # RFC 4180
```

### Example pipelines

```bash
# All string cells across every sheet.
zlsx cells data.xlsx --all-sheets | jq 'select(.t=="str") | {sheet, ref, v}'

# Sum a column from the CLI without loading everything.
zlsx cells data.xlsx --range B2:B1000 | jq -r 'select(.t=="int" or .t=="num") | .v' | awk '{s+=$1} END {print s}'

# Every comment across every sheet, as TSV.
zlsx comments data.xlsx --all-sheets | jq -r '[.sheet, .ref, .author, .text] | @tsv'

# Schema check: every list-type validation + its range.
zlsx validations data.xlsx | jq 'select(.rule_type=="list") | {sheet, range, values}'

# Grep SST for emails.
zlsx sst data.xlsx | jq -r '.text' | rg '@\S+\.\S+'
```

### Pipeline safety

`zlsx cells huge.xlsx | head -10` exits 0 cleanly (no broken-pipe stderr noise). `SIGINT` → exit 130, `SIGTERM` → exit 143, both flushing in-flight records. Non-fatal parse errors (corrupt sheet in an otherwise-valid workbook) surface as inline `{"kind":"error",…}` records instead of aborting the pipeline — filter with `jq 'select(.kind!="error")'` if you want the data-only stream.

### Exit codes

| Code | Meaning |
|---|---|
| 0 | Success (inline `error` records may still have been emitted) |
| 1 | Bad CLI arguments |
| 2 | Could not open file / not a valid xlsx archive |
| 3 | Sheet not found (by name / index / glob) |
| 4 | Decompression limit exceeded (`ZipBombSuspected`, reserved) |
| 5 | OS error (permission denied, disk full on stdout, etc.) |
| 130 | SIGINT |
| 143 | SIGTERM |

Emission overhead is within 3% of the tally-only benchmark — `zlsx big.xlsx | jq` beats any Python-based xlsx reader by 4×+.

## Performance

### Read — 30-run `hyperfine -N` median on `worldbank_catalog.xlsx` (67 KB, 161 × 26, 1,144 shared strings)

| Library | Time | Peak RSS | Speedup |
|---|---|---|---|
| **zlsx** | **3.3 ms** | **2.25 MB** | **1.00×** |
| calamine-rust 0.26 | 4.0 ms | 3.08 MB | 1.19× slower |
| python-calamine 0.6 | 23.6 ms | 16.92 MB | 7.10× slower |
| openpyxl 3.1 (read_only) | 129.5 ms | 42.39 MB | 38.9× slower |

zlsx now leads calamine-rust on **every corpus file** (1.06-1.19× faster), holds the **smallest RSS of the four** on every workload (half of calamine, 8-19× below the Python stack), and beats every Python option 7-40×. See [docs/benchmarks.md](docs/benchmarks.md) for the per-file breakdown.

### Write — 20-run `hyperfine -N` median writing 1,001 styled rows × 10 cols

| Library | Time | Peak RSS | Output | Speedup |
|---|---|---|---|---|
| **zlsx Writer** | **7.2 ms** | **4.40 MB** | 54.9 KB | **1.00×** |
| xlsxwriter 3.2 (`constant_memory`) | 70.3 ms | 25.41 MB | 55.2 KB | 9.82× slower |
| openpyxl 3.1 (`write_only`) | 155.7 ms | 42.05 MB | 52.8 KB | 21.74× slower |

zlsx Writer ships an in-house LZ77 + dynamic-huffman deflate compressor with lazy matching and a word-size SIMD match-length compare — 8 bytes per XOR-then-`@ctz` pass in the LZ77 inner loop, ~6× fewer iterations than byte-at-a-time on typical 3-30-byte XML matches. Zig 0.15.2's stdlib `std.compress.flate.Compress` is still mid-refactor and does not compile (we piggy-back on `std.compress.flate.HuffmanEncoder`, the one module in `std.compress.flate` that *is* usable). Sub-1 KB entries bypass compression so the dynamic-block header overhead doesn't inflate tiny XML. **~139,000 styled rows/sec** — 9.82× xlsxwriter, 21.74× openpyxl, at a third of xlsxwriter's RSS. See [`docs/benchmarks.md`](docs/benchmarks.md) for the full matrix.

## Zig version

Built against **Zig 0.15.2**. Uses `std.Io` / writer-gate APIs, `std.zip.Iterator`, `std.compress.flate.Decompress`. Older Zig versions need minor surgery on the Reader/Writer types.

## API

### `Book`

```zig
pub fn open(allocator: Allocator, path: []const u8) !Book
pub fn deinit(self: *Book) void
pub fn sheetByName(self: *const Book, name: []const u8) ?Sheet
pub fn rows(self: *const Book, sheet: Sheet, allocator: Allocator) !Rows
pub fn mergedRanges(self: *const Book, sheet: Sheet) []const MergeRange
pub fn hyperlinks(self: *const Book, sheet: Sheet) []const Hyperlink
pub fn dataValidations(self: *const Book, sheet: Sheet) []const DataValidation
```

`Book.sheets` is a `[]const Sheet` (name + path pairs) exposed for enumeration. Shared strings live in `Book.shared_strings: []const []const u8`. `mergedRanges(sheet)` returns a slice of `MergeRange { top_left: CellRef, bottom_right: CellRef }` where `CellRef = { col: u32, row: u32 }` — column is 0-based (A=0), row is 1-based (row1=1). `hyperlinks(sheet)` returns a slice of `Hyperlink { top_left, bottom_right, url, location }` with exactly one of `url` / `location` non-empty per entry. External entries set `url` to the raw `Target` attribute resolved through `_rels/sheet{N}.xml.rels`; internal entries set `location` to the raw `location=` attribute (e.g. `Sheet2!A1`). XML entities are preserved in both fields so the values round-trip byte-for-byte; decode at the caller if you need a display form. Everything is owned by the `Book` until `deinit`.

### `Rows`

```zig
pub fn next(self: *Rows) !?[]const Cell
pub fn deinit(self: *Rows) void
```

Each `next()` returns a dense `[]Cell` where `cells[i]` is the cell in column `i` (0-based). Missing cells are `.empty`. Any `.string` slice is borrowed from the `Book`'s internal buffers or from the row's own short-lived scratch — copy it out if you need it past the next `next()` call.

### `Cell`

```zig
pub const Cell = union(enum) {
    empty,
    string: []const u8,
    integer: i64,
    number: f64,
    boolean: bool,
};
```

Integers are parsed first; pure integer values become `.integer` to avoid float-round-trip loss.

Dates live in `.number` as Excel serial days; decode via:

```zig
if (xlsx.fromExcelSerial(cell.number)) |dt| {
    std.debug.print("{d}-{d:0>2}-{d:0>2} {d:0>2}:{d:0>2}:{d:0>2}\n",
        .{ dt.year, dt.month, dt.day, dt.hour, dt.minute, dt.second });
}
```

Returns `null` outside `1900-03-01 .. 9999-12-31` or on `NaN` / `±inf` (serials ≤ 60 hit the 1900 leap-year-bug window).

## Tests

```bash
# unit + fuzz-smoke tests (1k fuzz iters/target default, ~700 ms)
zig build test

# deep fuzz: override iteration count
XLSX_FUZZ_ITERS=1_000_000 zig build test

# integration tests against real public xlsx files (see tests/corpus/)
zig build test-corpus

# refresh the corpus if it goes stale
bash scripts/fetch_test_corpus.sh
```

14 fuzz targets — random-byte + mutation-driven — against every internal parser and the public `Book.open`. Constraint enforced: no crashes, no panics, no unreachable, no OOM. Tested to 14 M inputs per release.

See [`docs/xlsx_test_corpus.md`](docs/xlsx_test_corpus.md) for the public datasets exercised by the integration suite (Frictionless Data sample, openpyxl `guess_types`, Apache-POI-style 3-sheet test, World Bank Data Catalog with 1,144 shared strings).

## License

MIT — see [LICENSE](LICENSE).
