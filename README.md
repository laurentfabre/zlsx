# zlsx

Tiny `.xlsx` reader **and** writer for Zig. Single-file library, no third-party deps — just `std.zip` + `std.compress.flate` (for reads) + an in-house LZ77 + dynamic-huffman deflate compressor with lazy matching (for writes, since Zig 0.15.2's `std.compress.flate.Compress` is still mid-refactor) + a hand-rolled XML walker scoped to what spreadsheets actually need. Ships with a CLI, a C ABI, and Python bindings.

**Reader**: 1.5-2.5 ms on small files (alongside calamine-rust, process-startup-bound), 13 ms / 3.4 MB on a 67 KB workbook with 1,144 shared strings — 1.4× faster than python-calamine, 5.6× faster than openpyxl on that file, at ~5× lower RSS than the Python stack. calamine-rust still leads on SST-heavy workloads (3.5 ms vs zlsx's 13 ms post-arena). [Full benchmark table →](docs/benchmarks.md)

**Writer** (Phase 3b, v0.2.4): pragmatic openpyxl-parity styles — bold/italic, font size/name/color, horizontal alignment, wrap text, 19 OOXML fill patterns, 14 border styles × 5 sides, custom number formats, column widths, freeze panes, auto-filter, merged cell ranges, external-URL hyperlinks. Survived 1M-iter deep fuzz on every surface.

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
- **Read** workbooks — shared strings (with rich-text runs + XML entities), inline strings, numeric / boolean / error / formula-cached cells, UTF-8 throughout, merged-cell ranges via `Book.mergedRanges(sheet)`
- **Write** workbooks — strings (SST-deduped), integers, numbers, booleans, empties, multi-sheet; cell styles with fonts, fills, borders, alignment, wrap, number formats; per-sheet column widths, freeze panes, auto-filter, merged cell ranges, external-URL hyperlinks (per-sheet `_rels`)
- XML entity decoding (`&amp;`, `&lt;`, `&gt;`, `&quot;`, `&apos;`, `&#N;`, `&#xN;`) on read and escaping on write
- CLI (`zlsx file.xlsx --format {jsonl,jsonl-dict,tsv,csv}`), C ABI (`libzlsx.{dylib,so,dll}` + `include/zlsx.h`), Python bindings (`pip install py-zlsx`)

**Out (by design)**
- No formula evaluation — the reader returns the cached `<v>` value, the writer never synthesises formulas
- No date decoding — dates stay as their raw Excel serial number unless the generator pre-serialised them
- No load-modify-save round-trip yet — Phase 3c queued. For now the writer only produces fresh workbooks
- No chart / pivot / image extraction or emission

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

### Read — 20-run hyperfine median on `worldbank_catalog.xlsx` (67 KB, 161 × 26, 1,144 shared strings)

| Library | Time | Peak RSS | Speedup |
|---|---|---|---|
| calamine-rust 0.26 | **3.5 ms** | 3.09 MB | **1.00×** |
| **zlsx** | 13.4 ms | 3.38 MB | 3.84× slower |
| python-calamine 0.6 | 19.2 ms | 17.0 MB | 5.49× slower |
| openpyxl 3.1 (read_only) | 75.1 ms | 29.9 MB | 21.46× slower |

zlsx sits alongside calamine-rust at 1.5-2.5 ms on smaller files (≤30 KB, both startup-bound) and beats every Python option 4-40× across the whole corpus; SST-heavy reads vs calamine remain an open perf TODO (see [docs/benchmarks.md](docs/benchmarks.md) for the breakdown).

### Write — 20-run median writing 1,001 styled rows × 10 cols (Phase 3b, v0.2.4)

| Library | Time | Peak RSS | Output | Speedup |
|---|---|---|---|---|
| **zlsx Writer** | **37.4 ms** | **6.19 MB** | 54 KB | **1.00×** |
| xlsxwriter 3.2 (`constant_memory`) | 76.5 ms | 25.3 MB | 54 KB | 2.05× slower |
| openpyxl 3.1 (`write_only`) | 121.6 ms | 28.8 MB | 52 KB | 3.26× slower |

zlsx Writer ships an in-house LZ77 + dynamic-huffman deflate compressor with lazy matching and a word-size SIMD match-length compare — 8 bytes per XOR-then-`@ctz` pass in the LZ77 inner loop, ~6× fewer iterations than byte-at-a-time on typical 3-30-byte XML matches. Zig 0.15.2's stdlib `std.compress.flate.Compress` is still mid-refactor and does not compile (we piggy-back on `std.compress.flate.HuffmanEncoder`, the one module in `std.compress.flate` that *is* usable). Sub-1 KB entries bypass compression so the dynamic-block header overhead doesn't inflate tiny XML. Archive size matches xlsxwriter to the byte; wall time is half of xlsxwriter's and a third of openpyxl's. See [`docs/benchmarks.md`](docs/benchmarks.md) for the full matrix.

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
```

`Book.sheets` is a `[]const Sheet` (name + path pairs) exposed for enumeration. Shared strings live in `Book.shared_strings: []const []const u8`. `mergedRanges(sheet)` returns a slice of `MergeRange { top_left: CellRef, bottom_right: CellRef }` where `CellRef = { col: u32, row: u32 }` — column is 0-based (A=0), row is 1-based (row1=1). Everything is owned by the `Book` until `deinit`.

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
