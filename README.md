# zlsx

Tiny, read-only `.xlsx` parser for Zig. Single-file, no third-party deps, just `std.zip` + `std.compress.flate` + a hand-rolled XML walker scoped to what spreadsheets actually need.

On a 261 KB / 1,008-row workbook: **10.7 ms** wall time, **4.2 MB** peak RSS. That's 1.4× faster than calamine-rust, 4× faster than python-calamine, 24× faster than openpyxl — at one tenth the memory of the Python stack. [Full benchmark table →](docs/benchmarks.md)

```zig
const xlsx = @import("zlsx");

var book = try xlsx.Book.open(allocator, "workbook.xlsx");
defer book.deinit();

const sheet = book.sheetByName("Summary") orelse return error.MissingSheet;
var rows = try book.rows(sheet, allocator);
defer rows.deinit();

while (try rows.next()) |cells| {
    for (cells, 0..) |c, col| switch (c) {
        .empty   => {},
        .string  => |s| std.debug.print("{d:>3}: str  {s}\n", .{ col, s }),
        .integer => |i| std.debug.print("{d:>3}: int  {d}\n", .{ col, i }),
        .number  => |f| std.debug.print("{d:>3}: num  {d}\n", .{ col, f }),
        .boolean => |b| std.debug.print("{d:>3}: bool {any}\n", .{ col, b }),
    };
}
```

## Why

Zig's ecosystem has no official xlsx reader yet. The options are [calamine](https://github.com/tafia/calamine) via FFI (Rust), Apache POI via subprocess (JVM), or rolling your own on top of `std.zip`. This is the third option, packaged up and fuzz-tested.

Designed for a real use case: Alfred's hotel-concierge pipeline needs to read a 1,008-row × 35-col openpyxl-generated workbook with inline strings, shared strings, and UTF-8 content across multiple languages. `zlsx` reads it in ~30 ms.

## What's in, what's out

**In**
- Read a workbook, enumerate sheets by name/index
- Iterate rows as dense `[]Cell` with typed cells (`empty`, `string`, `integer`, `number`, `boolean`)
- Shared strings (`sharedStrings.xml` with rich-text runs and XML entities)
- Inline strings (`<c t="inlineStr">`)
- Numeric/boolean/error/formula-cached cells
- XML entity decoding (`&amp;`, `&lt;`, `&gt;`, `&quot;`, `&apos;`, `&#N;`, `&#xN;`)
- UTF-8 content throughout (non-Latin scripts preserved)

**Out (by design)**
- No writing — read-only
- No formula evaluation — we read the cached `<v>` value, never re-compute
- No style / number-format application — dates stay as their raw Excel serial unless the generator pre-serialized them
- No merged-cell semantics — the underlying cell anchors are returned; merging is caller-interpreted
- No chart / pivot / image extraction

## Install

### CLI binary

Prebuilt binaries for macOS (ARM64, Intel), Linux (x86_64, ARM64 — static musl), and Windows (x86_64) ship with every tagged release:

```bash
# macOS (Apple Silicon)
curl -fsSL -o zlsx.tar.gz "https://github.com/laurentfabre/zlsx/releases/latest/download/zlsx-0.2.0-aarch64-apple-darwin.tar.gz"
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

20-run hyperfine median, real workload (`alfred_bdr_prospect_list.xlsx`, 261 KB, 1,008 rows × 35 cols):

| Library | Time | Memory | Speedup |
|---|---|---|---|
| **zlsx** | **10.7 ms** | **4.16 MB** | **1.00×** |
| calamine-rust 0.26 | 15.3 ms | 4.94 MB | 1.44× slower |
| python-calamine 0.6 | 44.9 ms | 23.69 MB | 4.21× slower |
| openpyxl 3.1 (read_only) | 254.6 ms | 44.17 MB | 23.89× slower |

zlsx edges calamine-rs because (a) the Zig binary is ~120 KB vs calamine's ~620 KB static link so startup is shorter, and (b) zlsx borrows string slices into the source xml buffer whenever possible, only allocating per-row owned strings when rich-text concatenation or entity decoding forces it. See [`docs/benchmarks.md`](docs/benchmarks.md) for the full 4-library × 5-file matrix, reproducibility commands, and counter-difference analysis.

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
