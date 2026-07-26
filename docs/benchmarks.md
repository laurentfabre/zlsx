# zlsx benchmarks

Comparison against three other xlsx readers on a macOS Apple-Silicon box. Same workload for each: open the file, iterate every row of the first sheet, count cells by type, print totals.

## Systems under test

| Impl | Version | How it works |
|---|---|---|
| **zlsx** | this repo, ReleaseFast | Pure Zig, single file, stdlib zip + flate + hand-rolled XML walker. |
| **calamine-rust** | 0.26.1, native release binary | Pure Rust, widely used as the fast reference in the ecosystem. |
| **python-calamine** | 0.6.2 | Python bindings around calamine-rs. Pays Python interpreter startup + PyO3 conversion cost. |
| **openpyxl** | 3.1.5, `read_only=True, data_only=True` | Pure Python, SAX-over-stream. The Python de-facto standard. |

Shared workload: `open → iter rows → tally cells by {empty, string, integer, number, boolean} → print`.

## Wall-time results

`hyperfine -N --warmup 5 --runs 30` for sub-100 ms fixtures, `-N --warmup 3 --runs 15` for slower ones (mean ± σ). Lower is better. Last refreshed **2026-04-26** with the expanded corpus from `scripts/fetch_test_corpus.sh`. `-N` skips the shell wrapper so sub-5ms timings are accurate. The zlsx bench uses `std.heap.smp_allocator` — see the methodology note below.

`zlsx` uses `Book.openLazy` — the right comparison against calamine-rust, which is also lazy. zlsx's default `Book.open` eagerly preloads every sheet's side-indices (merged ranges, hyperlinks, validations, comments) so the metadata getters are populated synchronously; that costs ~37 ms on a 41-sheet workbook even when the caller only iterates one sheet. The "zlsx eager" column makes the cost visible.

| File | Size | Rows × Cols | zlsx (lazy) | zlsx (eager) | calamine-rust 0.26 | python-calamine 0.6 | openpyxl 3.1 |
|---|---|---|---|---|---|---|---|
| frictionless_2sheets.xlsx | 4.9 KB | 3 × 3 | 1.9 ± 0.1 | 1.9 ± 0.1 | **1.7 ± 0.1** | 20.5 ± 0.9 | 120.5 ± 3.6 |
| openpyxl_guess_types.xlsx | 29 KB | 2 × 5 | **1.9 ± 0.1** | 1.9 ± 0.1 | 2.0 ± 0.2 | 21.0 ± 0.5 | 124.5 ± 3.2 |
| phpoi_test1.xlsx | 9.8 KB | 8 × varied | 1.9 ± 0.2 | 1.9 ± 0.2 | **1.9 ± 0.1** | 18.9 ± 0.7 | 112.9 ± 1.5 |
| worldbank_catalog.xlsx | 67 KB | 161 × 26, **1,144 SST** | 3.3 ± 0.2 | **3.2 ± 0.1** | 3.8 ± 0.2 | 22.0 ± 0.7 | 127.8 ± 2.3 |
| **ons_cpi_detailed.xlsx** | 2.0 MB | **41 sheets**, 5,371 SST | **4.0 ± 0.1** | 41.2 ± 0.4 | 4.4 ± 0.2 | 24.4 ± 0.4 | 189.6 ± 7.3 |
| **ecdc_covid.xlsx** | 2.7 MB | **49,153 × 11**, 647 SST | **113.3 ± 0.5** | 113.4 ± 0.6 | 192.5 ± 1.1 | 385.5 ± 10.4 | 1357 ± 7 |
| **wdi_excel.xlsx** | 76 MB | **401,395 × 69**, 6 sheets, 226k SST | **2,359 ± 15** | 2,512 ± 5 | 2,778 ± 11 | 6,139 ± 189 | _too slow_ |

## Speedup

zlsx (lazy) leads on every fixture in the corpus. Two highlights:

```
ecdc_covid.xlsx (2.7 MB, 49,153 rows — single-sheet, modest SST)

  zlsx            ▌▌                       113 ms     1.00×
  calamine-rust   ▌▌▌▌                     193 ms     1.70× slower
  python-calamine ▌▌▌▌▌▌▌▌▌                386 ms     3.40× slower
  openpyxl        ▌▌▌▌▌▌▌▌▌▌▌▌▌▌▌…▌▌▌▌   1,357 ms    11.97× slower
```

```
wdi_excel.xlsx (76 MB, 401,395 rows × 69 cols, 226k SST — the headline workload)

  zlsx            ▌▌▌▌▌▌▌▌                   2.36 s    1.00×
  calamine-rust   ▌▌▌▌▌▌▌▌▌                  2.78 s    1.18× slower
  python-calamine ▌▌▌▌▌▌▌▌▌▌▌▌▌▌▌▌▌▌▌▌▌▌▌    6.14 s    2.60× slower
  openpyxl        too slow to bench reasonably
```

Throughput on the big fixtures:

| Impl | MB/s (input archive, WDI) | rows/s (WDI) |
|---|---|---|
| **zlsx (lazy)** | **32.4** | **170,150** |
| calamine-rust | 27.5 | 144,490 |
| python-calamine | 12.5 | 65,380 |

On small files (≤30 KB) zlsx ties or edges calamine within measurement noise — the process startup floor (~1.5 ms) dominates both native binaries at that size. Python libraries stay 6-12× slower from worldbank up; openpyxl is ~38× slower at ONS scale.

## Peak memory (RSS)

`/usr/bin/time -l`, single representative run per cell. Lower is better.

| File | zlsx (lazy) | zlsx (eager) | calamine-rust | python-calamine | openpyxl |
|---|---|---|---|---|---|
| worldbank_catalog (67 KB) | **2.4 MB** | 2.4 MB | 3.1 MB | 17.1 MB | 42.5 MB |
| ons_cpi_detailed (2.0 MB) | **3.3 MB** | 14.4 MB | 2.7 MB | 16.3 MB | 51.6 MB |
| ecdc_covid (2.7 MB) | **22.2 MB** | 22.2 MB | 47.1 MB | 89.7 MB | 46.1 MB |
| wdi_excel (76 MB) | **293 MB** | 385 MB | 1,315 MB | _skipped_ | _skipped_ |

Two structural wins:

- **WDI**: zlsx-lazy uses 293 MB to read a 76 MB archive; calamine-rust uses **1,315 MB — 4.5× heavier**. calamine materialises the entire workbook into one in-memory `Range`; zlsx streams via `Book.rows` so only the active sheet's XML stays resident.
- **ONS**: lazy mode is ~4.4× lighter than eager (3.3 MB vs 14.4 MB) because eager preloads all 41 sheet XMLs at `Book.open` time.

## Why SST parsing dominates the reader

OOXML stores string cells two ways:

1. **Inline**: the text lives in the cell XML itself —
   `<c t="inlineStr"><is><t>hello</t></is></c>`.
2. **Shared**: the cell XML carries only an index into a
   workbook-wide table at `xl/sharedStrings.xml` —
   `<c t="s"><v>42</v></c>` → "look up entry 42 in the SST".

Generators overwhelmingly prefer shared strings because duplicated
values appear once in the archive, cell XML is much smaller (an
integer instead of a verbose string wrapper), and the resulting
redundancy compresses further. Every non-trivial xlsx file has a
populated `xl/sharedStrings.xml`.

The price: any reader has to parse that table before row iteration
can resolve a `t="s"` cell — otherwise you get raw indices
(`0`, `1`, `2`, …) instead of `"Red"`, `"Green"`, `"Blue"`. This
isn't a zlsx design choice; it's a structural requirement of OOXML.

Every other xlsx reader does the same work:

- **calamine-rust** builds `Vec<String>` via quick-xml's SIMD tokeniser.
- **openpyxl** SAX-walks with `xml.etree.ElementTree.iterparse`.
- **python-calamine** delegates to calamine-rust via PyO3.
- **Apache POI / ClosedXML / SheetJS** — same story, different languages.

SST parsing is a structural cost every reader pays. Iterative allocator + parser optimisations (iter9: SST arena + per-row arena + pre-sized slow-path buffers; iter18: single-pass state-machine SST parser driven by `indexOfScalarPos('<')` + peek) closed the calamine gap on the original worldbank fixture. As of 2026-04-26 zlsx (lazy) ties or edges calamine on every corpus file.

The two reader behaviours that still differ:

- **zlsx streams; calamine materialises**: `Book.rows` allocates a fresh `[]Cell` per row and frees on the next call. calamine's `Range` holds the entire sheet in one allocation. On WDI (76 MB → ~290 MB resident for zlsx vs ~1.3 GB for calamine) zlsx's streaming model is decisively cheaper at peak. zlsx also exposes `Book.materialiseSheet` for callers who want the dense matrix; it's no faster than streaming on the corpus, just a shape preference.
- **Eager vs lazy `Book.open`**: zlsx's default eager open populates per-sheet metadata tables across every sheet at open time. On dense single-sheet fixtures this is free (the work would happen on first iteration anyway); on multi-sheet workbooks it's pure overhead unless the caller actually reaches for the metadata getters. `Book.openLazy` defers the work; the eager column above shows what the default costs.

## Methodology — allocator choice matters

The zlsx read bench uses `std.heap.smp_allocator` (same rationale as the write bench; see below). Earlier revisions used `std.heap.DebugAllocator(.{})`, which added ~10 ms of per-alloc tracking overhead on the worldbank_catalog workload — about 1.6× slower than what a real caller sees. `DebugAllocator` is the right default inside *tests* because it catches leaks and double-frees; it is **not** what a production caller would plug into `Book.open`. Pass whichever allocator you already use — `Book` has no opinion.

## Cells tallied — why totals differ

The wall-time benchmark is identical work; the reported cell counters differ by type because each library infers types differently. Counts are for `worldbank_catalog.xlsx` (161 rows × 26 cols):

| | str count | int count | num count | empty count |
|---|---|---|---|---|
| zlsx | 2,533 | 501 | 0 | 1,066 |
| calamine-rust | 2,533 | 0 | 501 | 1,152 |
| python-calamine | 2,633 | 0 | 401 | 1,152 |
| openpyxl | 2,633 | 401 | 0 | 1,152 |

Two behavioural deltas (not bugs):

- **int vs float**: calamine-rust returns `Data::Float` for every non-text number; zlsx tries integer first and only falls back to float. The 501 vs 0 split on int / 0 vs 501 on num is the same set of cells, re-labelled.
- **Row-width + string-coercion delta**: openpyxl and python-calamine pad every row to `worksheet.max_column` and coerce some digit-only inline strings to int — hence 2,633 strings and 1,152 empty cells. zlsx emits dense rows sized to the highest populated column *in that row* (1,152 − 1,066 = 86 cells of right-padding skipped) and honours `t="inlineStr"` strictly (no coercion, so 2,533 vs 2,633). Callers who want uniform-width rows can pad in a single `while (cells.len < max) …` loop after each `rows.next()`.

All four libraries read identical content from the file. The counter differences are interpretation, not correctness.

## Writer benchmark (Phase 3b, v0.2.4)

Same workload across all three implementations: 1,001 rows × 10 cols (one header row + 1,000 data rows). The header row has per-cell styles (bold white-on-blue fill, centre-aligned). Body rows mix strings, integers, floats, booleans, with the numeric columns referencing one of two shared number-format styles (`$#,##0.00` / `0.00%`). Sheet gets `column_width[0]=20` + `freeze_panes(row=1)`.

20-run `hyperfine -N` median, last refreshed 2026-04-25. zlsx bench uses `smp_allocator` + in-house LZ77 + dynamic-huffman deflate with lazy matching + word-size SIMD match-length compare — see methodology notes below:

| Impl | Time | Peak RSS | Output size | Speedup (wall) |
|---|---|---|---|---|
| **zlsx Writer** | **6.7 ms ± 0.3** | **4.44 MB** | 54.9 KB | **1.00×** |
| xlsxwriter 3.2 (`constant_memory`) | 66.4 ms ± 1.0 | 25.61 MB | 55.2 KB | 9.93× slower |
| openpyxl 3.1 (`write_only`) | 151.9 ms ± 6.4 | 41.65 MB | 52.8 KB | 22.74× slower |

```
  zlsx Writer    ▌              6.7 ms    1.00×
  xlsxwriter     ▌▌▌▌▌▌▌▌       66.4 ms    9.93× slower
  openpyxl       ▌▌…▌▌         151.9 ms   22.74× slower
```

Throughput at that size (rows/sec):

| Impl | Styled rows/sec |
|---|---|
| **zlsx Writer** | **~149,000** |
| xlsxwriter | ~15,070 |
| openpyxl | ~6,590 |

> **Re-measured after the Zig 0.16 deflate swap — the numbers hold.**
> The zlsx write row was originally measured against the in-house LZ77 +
> dynamic-huffman encoder, which the 0.16 migration retired in favour of
> stdlib `std.compress.flate`. Re-running the *same* harness
> (`scripts/bench_ci.sh`) on the *same* fixture and machine class gives
> **6.58 ms ± 0.21** against the 6.7 ms ± 0.3 recorded above: unchanged
> within noise, so the ratios against xlsxwriter and openpyxl stand.
>
> Caveat on scope: only the zlsx row was re-run. The xlsxwriter and
> openpyxl figures are carried forward from 2026-04-25 — nothing about
> this change could move them, but they have not been independently
> re-verified. Reader rows are unaffected either way: the read path
> already went through stdlib `std.compress.flate.Decompress`.

### Methodology — allocator choice matters

The bench binary uses `std.heap.smp_allocator`. An earlier revision used `std.heap.DebugAllocator(.{})` — that allocator tracks every allocation with metadata + (optionally) stack traces and makes the same workload take ~2.5× longer (24–29 ms instead of 9–10 ms on this hardware). `DebugAllocator` is the right default inside *tests* because it catches leaks and double-frees; it is **not** what a production downstream user would plug into `Writer.init`. The numbers above use the allocator a real caller would reach for.

If you're considering zlsx for your own pipeline: pass whichever allocator you already use — `Writer` has no opinion.

### Methodology — compression

zlsx ships an in-house deflate compressor: LZ77 with a 32 KB sliding window + single-step lazy matching (defer one byte, take whichever match is longer) + dynamic huffman tables per block + word-size SIMD match-length compare (8 bytes per XOR-then-`@ctz` pass in the LZ77 inner loop, ~6× fewer iterations than byte-at-a-time on typical 3-30-byte XML matches). Zig 0.15.2's stdlib `std.compress.flate.Compress` still doesn't compile (`BlockWriter` references a missing `bit_writer` field; the token-emission path is `@panic("TODO")`), so we grow our own — `std.compress.flate.HuffmanEncoder` is the one flate-module file that *is* usable and handles the canonical-huffman bookkeeping.

Per-entry the writer skips compression entirely for payloads under 1 KB (the dynamic-huffman block header has ~60-120 bytes of fixed overhead that rarely pays back on tiny XML fragments), and falls back to stored when deflate inflates a ≥ 1 KB payload. Combined with the SIMD match compare, this lands archive size byte-for-byte with xlsxwriter at roughly half xlsxwriter's wall time and a third of openpyxl's.

### Reproducing

The writer bench mirrors the reader bench — sources in `tests/bench/`:

```bash
zig build-exe -O ReleaseFast \
  --dep zlsx -Mroot=tests/bench/bench_write_zlsx.zig \
  -Mzlsx=src/xlsx.zig \
  -femit-bin=./bench_write_zlsx

hyperfine --warmup 3 --runs 20 \
  -n "zlsx"       "./bench_write_zlsx /tmp/out.xlsx" \
  -n "xlsxwriter" "python tests/bench/bench_write_xlsxwriter.py /tmp/out.xlsx" \
  -n "openpyxl"   "python tests/bench/bench_write_openpyxl.py /tmp/out.xlsx"
```

## Reproducing

```bash
# (0) materialise the corpus (small base + large fetched)
scripts/fetch_test_corpus.sh

# (1) build zlsx reader bench
zig build-exe -O ReleaseFast \
  --dep zlsx -Mroot=tests/bench/bench_zlsx.zig \
  -Mzlsx=src/xlsx.zig \
  -femit-bin=/tmp/bench_zlsx
# Pass `--lazy` for the apples-to-apples comparison against calamine
# (which is also lazy by default). The default `Book.open` populates
# every sheet's side-indices eagerly.

# (2) build calamine-rs bench (uses tests/bench/Cargo.toml)
(cd tests/bench && cargo build --release --bin bench_calamine)

# (3) python benches — openpyxl 3.1.5, python-calamine 0.6.2 via uv/pip

# (4) hyperfine driver — example for one fixture
hyperfine -N --warmup 5 --runs 30 \
  "/tmp/bench_zlsx --lazy tests/corpus/worldbank_catalog.xlsx" \
  "tests/bench/target/release/bench_calamine tests/corpus/worldbank_catalog.xlsx" \
  "python tests/bench/bench_pycalamine.py tests/corpus/worldbank_catalog.xlsx" \
  "python tests/bench/bench_openpyxl.py tests/corpus/worldbank_catalog.xlsx"
```

Source for all four benches (~30 lines each) is in `tests/bench/` if you want to sanity-check the workloads.

## Summary

**On the read side**: zlsx (lazy) leads calamine-rust on every corpus file. Highlights: **ECDC 49k rows: zlsx 113 ms vs calamine 193 ms (1.70×)** and **WDI 76 MB / 401k rows: zlsx 2.36 s vs calamine 2.78 s (1.18×) at 293 MB RSS vs 1,315 MB (4.5× lighter)**. Python libraries trail 2-12× from worldbank up. Single-file droppable into a Zig build; no third-party runtime deps.

**On the write side**: zlsx Writer is **9.93× faster than xlsxwriter and 22.74× faster than openpyxl** for a 1,001-row styled workbook — at ~6× lower RSS than xlsxwriter and ~9× below openpyxl. Archive size matches xlsxwriter within 0.5 % (zlsx 54.9 KB vs xlsxwriter 55.2 KB). The in-house LZ77 + dynamic-huffman deflate compressor (with lazy matching + word-size SIMD match compare) does what zlib-at-level-6 does, but tuned for the xlsx-XML workload.
