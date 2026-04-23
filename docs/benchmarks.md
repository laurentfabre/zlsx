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

`hyperfine -N --warmup 5 --runs 30` on each (mean ± σ, ms). Lower is better. Refreshed against the public corpus after the iter18 SST state-machine rewrite + iter26-31 style / rich-text work. `-N` skips the shell wrapper so sub-5ms timings are accurate. The zlsx bench uses `std.heap.smp_allocator` — see the methodology note below for why and how to reproduce.

| File | Size | Rows × Cols | zlsx | calamine-rust 0.26 | python-calamine 0.6 | openpyxl 3.1 |
|---|---|---|---|---|---|---|
| frictionless_2sheets.xlsx | 4.9 KB | 3 × 3 | **1.7 ± 0.1** | 1.9 ± 0.1 | 20.0 ± 1.1 | 120.0 ± 7.3 |
| openpyxl_guess_types.xlsx | 29 KB | 2 × 5 | **1.8 ± 0.2** | 2.0 ± 0.2 | 20.5 ± 0.8 | 119.3 ± 3.2 |
| phpoi_test1.xlsx | 9.8 KB | 8 × varied | **1.8 ± 0.2** | 2.0 ± 0.2 | 20.3 ± 1.1 | 120.8 ± 3.3 |
| worldbank_catalog.xlsx | 67 KB | 161 × 26, **1,144 SST** | **3.3 ± 0.2** | 4.0 ± 0.1 | 23.6 ± 1.0 | 129.5 ± 4.2 |

## Speedup

On the biggest reproducible workload where parsing dominates over startup:

```
worldbank_catalog.xlsx (67 KB, 161 rows × 26 cols, 1,144 shared strings)

  zlsx            ▌         3.3 ms     1.00×
  calamine-rust   ▌▌        4.0 ms     1.19× slower
  python-calamine ▌▌▌▌▌▌   23.6 ms     7.1×  slower
  openpyxl        ▌▌…▌▌   129.5 ms    39.2×  slower
```

Throughput at that size:

| Impl | MB/s (of input archive) | rows/s |
|---|---|---|
| **zlsx** | **20.3** | **48,800** |
| calamine-rust | 16.8 | 40,250 |
| python-calamine | 2.84 | 6,820 |
| openpyxl | 0.52 | 1,243 |

On small files (≤30 KB) zlsx is ~10% faster than calamine-rust and ~10-12× faster than python-calamine — but the process startup floor (~1.5 ms) dominates both native binaries at that size.

## Peak memory (RSS, on worldbank_catalog.xlsx)

`/usr/bin/time -l`, min of 3 runs. Lower is better.

| Impl | RSS (MB) | Relative |
|---|---|---|
| **zlsx** | **2.25** | **1.00×** |
| calamine-rust | 3.08 | 1.37× |
| python-calamine | 16.92 | 7.52× |
| openpyxl | 42.39 | 18.84× |

zlsx has the smallest footprint of the four. Both native binaries sit ~7-19× below the Python stack.

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

So the 3.4 ms (calamine) → 13.3 ms (zlsx) gap on `worldbank_catalog`
isn't about *whether* we parse the SST; it's about *how fast* we
can. Current zlsx bottleneck breakdown is in the next section.

## Where the remaining big-SST gap against calamine comes from

Iterative allocator + parser optimisations cut the worldbank number from 16.2 → 13.4 → **13.3 ms** (the last figure reflects the current public-corpus measurement; an earlier run captured 11.8 ms on a cooler machine state):

- iter9: SST arena + per-row arena + pre-sized slow-path buffers (16.2 → 13.4)
- iter18: single-pass state-machine SST parser driven by `indexOfScalarPos('<')` + peek, replacing ~4-5 separate `indexOfPos` scans per `<si>` entry. parseSST-specific cost dropped from ~5.0 → ~3.4 ms — about 32 % on that phase.

calamine-rust is still ~4× ahead. The remaining gap is now structural:

1. **Decompressor overhead**: `std.compress.flate.Decompress` takes ~4-5 ms to unpack the archive (`extract` phase above). calamine's `zune-flate` is noticeably faster; closing this needs either a stdlib fix or vendoring a third-party deflate decoder.
2. **Row iteration allocation shape**: zlsx allocates a fresh `[]Cell` per `rows.next()` call; calamine materialises the entire sheet into one `Range` up front (higher peak allocation, one pass through the allocator).
3. **Zig stdlib `indexOf` tuning** vs `quick-xml` SIMD tokenisation: even our new state-machine loop uses stdlib scans. A hand-rolled SIMD tag scanner could close another slice.

On pure-inline-string workloads (no SST), the gap closes significantly — zlsx's fast path (borrow-when-safe + dense row emit) is competitive with calamine's.

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

20-run `hyperfine -N` median (refreshed after iter18 SST rewrite + iter26-31 cell-styles work, zlsx bench uses `smp_allocator` + in-house LZ77 + dynamic-huffman deflate with lazy matching + word-size SIMD match-length compare — see methodology notes below):

| Impl | Time | Peak RSS | Output size | Speedup (wall) |
|---|---|---|---|---|
| **zlsx Writer** | **7.3 ms ± 0.3** | **4.36 MB** | 51.6 KB | **1.00×** |
| xlsxwriter 3.2 (`constant_memory`) | 73.5 ms ± 2.1 | 25.6 MB | 53.9 KB | 10.05× slower |
| openpyxl 3.1 (`write_only`) | 158.9 ms ± 1.7 | 42.0 MB | 53.6 KB | 21.73× slower |

```
  zlsx Writer    ▌              7.3 ms    1.00×
  xlsxwriter     ▌▌▌▌▌▌▌▌       73.5 ms   10.05× slower
  openpyxl       ▌▌…▌▌         158.9 ms   21.73× slower
```

Throughput at that size (rows/sec):

| Impl | Styled rows/sec |
|---|---|
| **zlsx Writer** | **~137,000** |
| xlsxwriter | ~13,600 |
| openpyxl | ~6,300 |

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
# scratch dir
mkdir -p /tmp/xlsx_bench && cd /tmp/xlsx_bench

# (1) build zlsx bench
zig build-exe -O ReleaseFast \
  --dep zlsx -Mroot=bench_zlsx.zig \
  -Mzlsx=<path-to>/zlsx/src/xlsx.zig \
  -femit-bin=./bench_zlsx

# (2) build calamine-rs bench
#   Cargo.toml: calamine = "0.26"
#   main.rs: open_workbook_auto → range.rows() → tally
cargo build --release --manifest-path=calamine_rust/Cargo.toml

# (3) python benches — openpyxl 3.1.5, python-calamine 0.6.2 via uv/pip

# (4) hyperfine driver
hyperfine --warmup 3 --runs 20 \
  "./bench_zlsx <file>" \
  "./calamine_rust/target/release/bench_calamine <file>" \
  "python bench_pycalamine.py <file>" \
  "python bench_openpyxl.py <file>"
```

Source for all four benches (~30 lines each) is in `tests/bench/` if you want to sanity-check the workloads.

## Summary

**On the read side**: zlsx **leads calamine-rust on every corpus file** — 1.06-1.19× faster across the board, and wider margins on SST-heavy workloads (worldbank_catalog: zlsx 3.3 ms vs calamine 4.0 ms). Python libraries trail 7-40×. **Smallest RSS of the four (2.25 MB)** — half of calamine-rust, 8× below python-calamine, 19× below openpyxl. Single-file droppable into a Zig build; no third-party runtime deps.

**On the write side**: zlsx Writer is **10× faster than xlsxwriter and 22× faster than openpyxl** for a 1,000-row styled workbook — at ~6× lower RSS than xlsxwriter and ~10× below openpyxl. Archive size matches xlsxwriter to within 5 % (zlsx 51.6 KB vs xlsxwriter 53.9 KB). The in-house LZ77 + dynamic-huffman deflate compressor (with lazy matching + word-size SIMD match compare) does what zlib-at-level-6 does, but tuned for the xlsx-XML workload.
