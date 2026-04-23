# zlsx

Tiny `.xlsx` reader **and** writer for Zig. Single-file library, no third-party deps — just `std.zip` + `std.compress.flate` (for reads) + an in-house LZ77 + dynamic-huffman deflate compressor with lazy matching (for writes, since Zig 0.15.2's `std.compress.flate.Compress` is still mid-refactor) + a hand-rolled XML walker scoped to what spreadsheets actually need. Ships with a CLI, a C ABI, and Python bindings.

**Reader**: 10.7 ms / 4.2 MB on a 261 KB / 1,008-row workbook — 1.4× faster than calamine-rust, 4× faster than python-calamine, 24× faster than openpyxl, at one tenth the memory of the Python stack. [Full benchmark table →](docs/benchmarks.md)

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
- **Read** workbooks — shared strings (with rich-text runs + XML entities), inline strings, numeric / boolean / error / formula-cached cells, UTF-8 throughout
- **Write** workbooks — strings (SST-deduped), integers, numbers, booleans, empties, multi-sheet; cell styles with fonts, fills, borders, alignment, wrap, number formats; per-sheet column widths, freeze panes, auto-filter, merged cell ranges, external-URL hyperlinks (per-sheet `_rels`)
- XML entity decoding (`&amp;`, `&lt;`, `&gt;`, `&quot;`, `&apos;`, `&#N;`, `&#xN;`) on read and escaping on write
- CLI (`zlsx file.xlsx --format {jsonl,jsonl-dict,tsv,csv}`), C ABI (`libzlsx.{dylib,so,dll}` + `include/zlsx.h`), Python bindings (`pip install py-zlsx`)

**Out (by design)**
- No formula evaluation — the reader returns the cached `<v>` value, the writer never synthesises formulas
- No date decoding — dates stay as their raw Excel serial number unless the generator pre-serialised them
- No merged-cell authoring on the **reader** side — the underlying cell anchors are preserved; merging is caller-interpreted. (The **writer** does emit `<mergeCells>` — call `SheetWriter.addMergedCell("A1:B2")`.)
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

### Read — 20-run hyperfine median on `alfred_bdr_prospect_list.xlsx` (261 KB, 1,008 × 35)

| Library | Time | Peak RSS | Speedup |
|---|---|---|---|
| **zlsx** | **10.3 ms** | **4.16 MB** | **1.00×** |
| calamine-rust 0.26 | 15.3 ms | 4.94 MB | 1.48× slower |
| python-calamine 0.6 | 44.9 ms | 23.69 MB | 4.36× slower |
| openpyxl 3.1 (read_only) | 254.6 ms | 44.17 MB | 24.72× slower |

### Write — 20-run median writing 1,001 styled rows × 10 cols (Phase 3b, v0.2.4)

| Library | Time | Peak RSS | Output | Speedup |
|---|---|---|---|---|
| **zlsx Writer** | **71.4 ms** | **6.19 MB** | 54 KB | **1.00×** |
| xlsxwriter 3.2 (`constant_memory`) | 72.4 ms | 25.3 MB | 54 KB | 1.01× slower |
| openpyxl 3.1 (`write_only`) | 114.1 ms | 28.8 MB | 52 KB | 1.60× slower |

zlsx Writer ships an in-house LZ77 + dynamic-huffman deflate compressor with lazy matching (Zig 0.15.2's stdlib `std.compress.flate.Compress` is still mid-refactor and does not compile — we piggy-back on `std.compress.flate.HuffmanEncoder`, the one module in `std.compress.flate` that *is* usable). Sub-1 KB entries bypass compression so the dynamic-block header overhead doesn't inflate tiny XML. Archive size matches xlsxwriter to the byte and is within 3 % of openpyxl. See [`docs/benchmarks.md`](docs/benchmarks.md) for the full matrix.

## Zig version

Built against **Zig 0.15.2**. Uses `std.Io` / writer-gate APIs, `std.zip.Iterator`, `std.compress.flate.Decompress`. Older Zig versions need minor surgery on the Reader/Writer types.

## API

### `Book`

```zig
pub fn open(allocator: Allocator, path: []const u8) !Book
pub fn deinit(self: *Book) void
pub fn sheetByName(self: *const Book, name: []const u8) ?Sheet
pub fn rows(self: *const Book, sheet: Sheet, allocator: Allocator) !Rows
```

`Book.sheets` is a `[]const Sheet` (name + path pairs) exposed for enumeration. Shared strings live in `Book.shared_strings: []const []const u8`. Everything is owned by the `Book` until `deinit`.

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
