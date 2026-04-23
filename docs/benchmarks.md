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

`hyperfine --warmup 5 --runs 20` on each (mean ± σ, ms). Lower is better. Refreshed 2026-04-23 against the public corpus (the 261 KB alfred_bdr workload that earlier revisions cited is not checked in). The zlsx bench now uses `std.heap.smp_allocator` — see the methodology note below for why and how to reproduce.

| File | Size | Rows × Cols | zlsx | calamine-rust 0.26 | python-calamine 0.6 | openpyxl 3.1 |
|---|---|---|---|---|---|---|
| frictionless_2sheets.xlsx | 4.9 KB | 3 × 3 | **1.6 ± 0.2** | 1.8 ± 0.5 | 16.2 ± 0.4 | 64.5 ± 0.6 |
| openpyxl_guess_types.xlsx | 29 KB | 2 × 5 | **1.6 ± 0.1** | **1.5 ± 0.2** | 16.4 ± 0.3 | 65.5 ± 0.9 |
| phpoi_test1.xlsx | 9.8 KB | 8 × varied | 1.9 ± 0.1 | **1.5 ± 0.1** | 16.6 ± 0.4 | 65.5 ± 1.0 |
| worldbank_catalog.xlsx | 67 KB | 161 × 26, **1,144 SST** | 16.2 ± 0.2 | **3.5 ± 0.3** | 19.2 ± 0.4 | 75.1 ± 0.8 |

## Speedup

On the biggest reproducible workload where parsing dominates over startup:

```
worldbank_catalog.xlsx (67 KB, 161 rows × 26 cols, 1,144 shared strings)

  calamine-rust   ▌         3.5 ms     1.00×
  zlsx            ▌▌▌▌▌    16.2 ms     4.6× slower
  python-calamine ▌▌▌▌▌▌   19.2 ms     5.5× slower
  openpyxl        ▌▌…▌▌    75.1 ms    21.5× slower
```

Throughput at that size:

| Impl | MB/s (of input archive) | rows/s |
|---|---|---|
| calamine-rust | 19.1 | 46,000 |
| zlsx | 4.1 | 9,940 |
| python-calamine | 3.5 | 8,380 |
| openpyxl | 0.89 | 2,140 |

On small files (≤30 KB) zlsx and calamine-rust are effectively tied at 1.5-2 ms — process startup dominates, and both are ~10× faster than python-calamine's ~16 ms floor (Python interpreter + PyO3 bridge).

## Peak memory (RSS, on worldbank_catalog.xlsx)

`/usr/bin/time -l`, min of 3 runs. Lower is better.

| Impl | RSS (MB) | Relative |
|---|---|---|
| calamine-rust | 3.09 | 1.00× |
| zlsx | 3.38 | 1.09× |
| python-calamine | 17.0 | 5.5× |
| openpyxl | 29.9 | 9.7× |

Both native binaries sit ~5-10× below the Python stack.

## Where the big-SST gap against calamine comes from

On workloads with many shared strings (here 1,144 in a 161-row file), calamine-rust is ~4.6× faster than zlsx. Likely factors:

1. **SST decode path**: zlsx currently decodes each shared-string entry lazily on first access but allocates an owned copy per decode when XML entities are present — which for workbooks with many `&amp;` / `&#N;` runs adds up. calamine appears to pool string buffers more aggressively. This is the most likely single cause.
2. **Row iteration allocation shape**: zlsx allocates a fresh `[]Cell` per `rows.next()` call; calamine materialises the entire sheet into one `Range` up front (higher peak allocation, but only one pass through the allocator).
3. **XML parser work**: both are hand-rolled; calamine's is older and more tuned.

On pure-inline-string workloads (no SST), the gap closes significantly — zlsx's fast path (borrow-when-safe + dense row emit) is competitive with calamine's. That's the regime our Alfred use-case sits in, so the default zlsx workload in Pro/ hits the fast path.

Reader perf room-for-improvement is an open roadmap item; the SST decode path is the obvious first target.

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

20-run hyperfine median (refreshed 2026-04-23, zlsx bench uses `smp_allocator` + in-house LZ77 + dynamic-huffman deflate with lazy matching — see methodology notes below):

| Impl | Time | Peak RSS | Output size | Speedup (wall) |
|---|---|---|---|---|
| **zlsx Writer** | **71.4 ms ± 0.7** | **6.19 MB** | 54 KB | **1.00×** |
| xlsxwriter 3.2 (`constant_memory`) | 72.4 ms ± 0.9 | 25.3 MB | 54 KB | 1.01× slower |
| openpyxl 3.1 (`write_only`) | 114.1 ms ± 1.4 | 28.8 MB | 52 KB | 1.60× slower |

```
  zlsx Writer    ▌▌▌▌▌▌▌         71.4 ms    1.00×
  xlsxwriter     ▌▌▌▌▌▌▌         72.4 ms    1.01× slower
  openpyxl       ▌▌▌▌▌▌▌▌▌▌▌▌   114.1 ms    1.60× slower
```

Throughput at that size (rows/sec):

| Impl | Styled rows/sec |
|---|---|
| zlsx Writer | ~14,000 |
| xlsxwriter | ~13,800 |
| openpyxl | ~8,800 |

### Methodology — allocator choice matters

The bench binary uses `std.heap.smp_allocator`. An earlier revision used `std.heap.DebugAllocator(.{})` — that allocator tracks every allocation with metadata + (optionally) stack traces and makes the same workload take ~2.5× longer (24–29 ms instead of 9–10 ms on this hardware). `DebugAllocator` is the right default inside *tests* because it catches leaks and double-frees; it is **not** what a production downstream user would plug into `Writer.init`. The numbers above use the allocator a real caller would reach for.

If you're considering zlsx for your own pipeline: pass whichever allocator you already use — `Writer` has no opinion.

### Methodology — compression

zlsx ships an in-house deflate compressor: LZ77 with a 32 KB sliding window + single-step lazy matching (defer one byte, take whichever match is longer) + dynamic huffman tables per block. Zig 0.15.2's stdlib `std.compress.flate.Compress` still doesn't compile (`BlockWriter` references a missing `bit_writer` field; the token-emission path is `@panic("TODO")`), so we grow our own — `std.compress.flate.HuffmanEncoder` is the one flate-module file that *is* usable and handles the canonical-huffman bookkeeping.

Per-entry the writer skips compression entirely for payloads under 1 KB (the dynamic-huffman block header has ~60-120 bytes of fixed overhead that rarely pays back on tiny XML fragments), and falls back to stored when deflate inflates a ≥ 1 KB payload. This matches the archive size of xlsxwriter/openpyxl (both use zlib at default level 6, same algorithm class) to within 3 % with comparable wall time.

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

**On the read side**: native parity with calamine-rust on small files (≤30 KB, both ~1.5-2 ms — process startup dominates), smallest RSS alongside calamine (~3 MB), single-file droppable into a Zig build. Against Python libraries it's 4-40× faster; the big Python win is on small files where the interpreter floor of ~16 ms dominates. On SST-heavy workloads (like worldbank_catalog's 1,144 shared strings) calamine-rust is currently ~4.6× faster than zlsx — the SST decode path is the obvious perf TODO.

**On the write side** (Phase 3b, v0.2.4): zlsx Writer matches xlsxwriter on wall time and output size and is 1.6× faster than openpyxl for a 1,000-row styled workbook — at ~4× lower RSS than either Python library. The in-house LZ77 + dynamic-huffman deflate compressor (with lazy matching) closes the archive-size gap to within 3 %, which was the last thing the older fixed-huffman + greedy-match code gave up to keep its size small.
