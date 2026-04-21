# zlsx

Tiny, read-only `.xlsx` parser for Zig. Single-file, no third-party deps, just `std.zip` + `std.compress.flate` + a hand-rolled XML walker scoped to what spreadsheets actually need.

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
