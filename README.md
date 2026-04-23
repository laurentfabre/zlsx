# zlsx

Tiny `.xlsx` reader **and** writer for Zig. Single-file library, no third-party deps — just `std.zip` + `std.compress.flate` (for reads) + an in-house LZ77 + dynamic-huffman deflate compressor with lazy matching (for writes, since Zig 0.15.2's `std.compress.flate.Compress` is still mid-refactor) + a hand-rolled XML walker scoped to what spreadsheets actually need. Ships with a CLI, a C ABI, and Python bindings.

**Reader**: 1.9 ms on small files (alongside calamine-rust, process-startup-bound), 13.3 ms / 2.95 MB on a 67 KB workbook with 1,144 shared strings — 1.5× faster than python-calamine, 5.9× faster than openpyxl on that file, at the **smallest RSS of the four** (~6× lower than the Python stack, slightly below calamine). calamine-rust still leads on SST-heavy workloads (3.4 ms vs zlsx's 13.3 ms post single-pass rewrite); the remaining gap is stdlib-flate decompression overhead. [Full benchmark table →](docs/benchmarks.md)

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
- **Read** workbooks — shared strings (with rich-text runs + XML entities), inline strings, numeric / boolean / error / formula-cached cells, UTF-8 throughout, merged-cell ranges via `Book.mergedRanges(sheet)`, external-URL hyperlinks via `Book.hyperlinks(sheet)` (resolved through sheet `_rels`), list-type data validations via `Book.dataValidations(sheet)` (values entity-decoded)
- **Write** workbooks — strings (SST-deduped), integers, numbers, booleans, empties, multi-sheet; cell styles with fonts, fills, borders, alignment, wrap, number formats; per-sheet column widths, row heights, freeze panes, auto-filter, merged cell ranges, external-URL hyperlinks (per-sheet `_rels`), internal hyperlinks (`location="Sheet2!A1"`), list-type data validations (dropdowns), number / decimal / date / time / text-length / custom data validations, formulas with optional cached value
- XML entity decoding (`&amp;`, `&lt;`, `&gt;`, `&quot;`, `&apos;`, `&#N;`, `&#xN;`) on read and escaping on write
- CLI (`zlsx file.xlsx --format {jsonl,jsonl-dict,tsv,csv}`), C ABI (`libzlsx.{dylib,so,dll}` + `include/zlsx.h`), Python bindings (`pip install py-zlsx`)

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
| Date as `DateTime` | helper² | ✓ | ✓ | ✓ |
| XML entity decoding | ✓ | ✓ | ✓ | ✓ |
| Merged cell ranges | ✓ | ✓ | ✓ | ✓ |
| External-URL hyperlinks | ✓ | ? | ✓ | ? |
| Data validations (list / dropdown) | ✓ | — | ✓ | — |
| Data validations (number / date / custom) | — | — | ✓ | — |
| Rich-text formatting preserved | — | ~ | ✓ | — |
| Cell styles on read (bold / colour / fill) | — | — | ✓ | — |
| Comments / notes | — | ? | ✓ | — |
| Chart / image / pivot access | — | — | ~ | — |
| Load-modify-save | — | — | ✓ | — |

¹ Returns a single `Float` type for any non-text number — callers cast to integer if needed.
² `xlsx.fromExcelSerial(cell.number) -> ?DateTime`; out-of-range serials (1900 leap-bug window) return `null`.

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
| Conditional formatting | — | ✓ | ✓ |
| Cell comments / notes | — | ✓ | ✓ |
| Formulas (with cached value) | ✓ | ✓ | ✓ |
| Rich-text runs per cell | — | ✓ | ✓ |
| Images (PNG / JPEG embed) | — | ✓ | ~ |
| Charts | — | ✓ | ~ |
| Deflate compression | ✓ | ✓ | ✓ |
| Load-modify-save | — | — | ✓ |
| Sheet-name validation (length / reserved chars / duplicates) | ✓ | ~³ | ~ |

³ xlsxwriter validates length and some chars but does not reject case-insensitive duplicates up front.

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

A thin binary `zlsx` ships with this repo. It streams rows of the selected sheet to stdout — useful as an openpyxl replacement for shell / pipeline use:

```bash
zig build -Doptimize=ReleaseFast
./zig-out/bin/zlsx file.xlsx                          # JSONL (default)
./zig-out/bin/zlsx file.xlsx --format tsv             # TSV with \N for empty
./zig-out/bin/zlsx file.xlsx --format csv             # RFC 4180 CSV
./zig-out/bin/zlsx file.xlsx --format jsonl-dict      # {"A": val, "B": val, …}
./zig-out/bin/zlsx file.xlsx --sheet 2                # 0-indexed
./zig-out/bin/zlsx file.xlsx --name "Summary"         # by name
./zig-out/bin/zlsx file.xlsx --list-sheets            # one name per line
```

Emission overhead is within 3% of the tally-only benchmark — the CLI is fast enough that `zlsx big.xlsx | jq` beats any Python-based xlsx reader by 4×+.

## Performance

### Read — 30-run hyperfine median on `worldbank_catalog.xlsx` (67 KB, 161 × 26, 1,144 shared strings)

| Library | Time | Peak RSS | Speedup |
|---|---|---|---|
| calamine-rust 0.26 | **3.4 ms** | 3.09 MB | **1.00×** |
| **zlsx** | 13.3 ms | **2.95 MB** | 3.91× slower |
| python-calamine 0.6 | 20.1 ms | 16.94 MB | 5.91× slower |
| openpyxl 3.1 (read_only) | 78.7 ms | 29.82 MB | 23.15× slower |

zlsx sits alongside calamine-rust at 1.2-2.1 ms on smaller files (≤30 KB, both startup-bound), holds the **smallest RSS of the four** on every workload, and beats every Python option 9-40× across the whole corpus; SST-heavy reads vs calamine remain an open perf TODO (see [docs/benchmarks.md](docs/benchmarks.md) for the breakdown).

### Write — 20-run median writing 1,001 styled rows × 10 cols (Phase 3b, v0.2.4)

| Library | Time | Peak RSS | Output | Speedup |
|---|---|---|---|---|
| **zlsx Writer** | **37.2 ms** | **6.20 MB** | 54 KB | **1.00×** |
| xlsxwriter 3.2 (`constant_memory`) | 70.4 ms | 25.4 MB | 54 KB | 1.89× slower |
| openpyxl 3.1 (`write_only`) | 107.6 ms | 29.0 MB | 52 KB | 2.89× slower |

zlsx Writer ships an in-house LZ77 + dynamic-huffman deflate compressor with lazy matching and a word-size SIMD match-length compare — 8 bytes per XOR-then-`@ctz` pass in the LZ77 inner loop, ~6× fewer iterations than byte-at-a-time on typical 3-30-byte XML matches. Zig 0.15.2's stdlib `std.compress.flate.Compress` is still mid-refactor and does not compile (we piggy-back on `std.compress.flate.HuffmanEncoder`, the one module in `std.compress.flate` that *is* usable). Sub-1 KB entries bypass compression so the dynamic-block header overhead doesn't inflate tiny XML. Archive size matches xlsxwriter to within 0.5 %; wall time is ~half of xlsxwriter's and ~a third of openpyxl's at 4× lower RSS. See [`docs/benchmarks.md`](docs/benchmarks.md) for the full matrix.

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
