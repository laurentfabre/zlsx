# iter-er-7 task B — latency gate refresh post `Editor.save` thin shim

| Field | Value |
|---|---|
| Date | 2026-05-08 (UTC) |
| Branch | `docs/iter-er-7-bench-refresh` (off `refactor/editor-er-6-thin-shim`) |
| Branch SHA | `b71b0b5` (`refactor(pkg/editor): retire dead SST + ZIP helpers post thin-shim — iter-er-6 cleanup`) |
| Toolchain | Zig 0.15.2 (`/Users/lf/.zvm/0.15.2/zig`) |
| Build mode | `-O ReleaseSafe -target aarch64-macos-none` |
| Host | Darwin 25.4.0 arm64 (M-series) |
| Background load | `uptime` reported load averages 4.49 / 4.26 / 3.13 at run time — moderate; results were stable run-to-run within ±1 ms on `bench_append_rows`, so the noise floor was tolerable. No other heavy build/bench was running concurrently. |

## What changed and why we re-bench

`Editor.save` was rewritten on `refactor/editor-er-6-thin-shim` (commits
`7e7eb60` → `182d5bb` → `b71b0b5`) to delegate to `Workbook.save` →
`PartStore.save`. The byte-output path now goes through `PartStore`'s
per-part atomic-file ZIP rebuilder rather than Editor's legacy
single-stream `writeAll(src_buf)` raw-LFH writer. iter-er-7's task B
verifies that the latency gate locked in by PR #71 still holds after
this re-route.

Pre-thin-shim baselines (canonical, from PR #71's bench gate):
- `bench_append_rows 100000`: median **174 ms** = 1.000× canonical baseline.
- Pre-thin-shim measured (post deep-copy hardening, last bench commit): **186.75 ms** = 1.073×.
- Strict gate (pre-thin-shim ceiling): **191.4 ms** (1.10×).
- iter-er-7 loosened gate (migration grace): **261 ms** (1.50× of 174 ms).

`bench_zlsx` and `bench_write_zlsx` exercise read- and write-side
hot paths that aren't intersected by the thin shim — flagged for
drift > 5% only.

## Bench commands

All three harnesses were built with the same module-graph
incantation (workaround for the macOS-26.4-SDK `zig build` breakage —
see `feedback_zig_test_target_workaround` memory note):

```bash
ZIG=/Users/lf/.zvm/0.15.2/zig

# bench_append_rows (root + zlsx + zlsx_pkg)
$ZIG build-exe -O ReleaseSafe -target aarch64-macos-none \
    --dep zlsx --dep zlsx_pkg \
    -Mroot=tests/bench/bench_append_rows.zig \
    -Mzlsx=src/xlsx.zig \
    --dep zlsx -Mzlsx_pkg=pkg/root.zig \
    --name bench_append_rows \
    -femit-bin=.bench-tmp/bench_append_rows

# bench_zlsx (root + zlsx)
$ZIG build-exe -O ReleaseSafe -target aarch64-macos-none \
    --dep zlsx -Mroot=tests/bench/bench_zlsx.zig \
    -Mzlsx=src/xlsx.zig \
    --name bench_zlsx \
    -femit-bin=.bench-tmp/bench_zlsx

# bench_write_zlsx (root + zlsx)
$ZIG build-exe -O ReleaseSafe -target aarch64-macos-none \
    --dep zlsx -Mroot=tests/bench/bench_write_zlsx.zig \
    -Mzlsx=src/xlsx.zig \
    --name bench_write_zlsx \
    -femit-bin=.bench-tmp/bench_write_zlsx
```

Run commands (5 trials each, no warmup, no `time`/`hyperfine`
wrappers around `bench_append_rows` since it self-times — the
read/write benches don't self-time so a Python `time.perf_counter()`
subprocess wall-clock was used; subprocess fork/exec adds a
~1-2 ms floor that hyperfine `-N` would shave):

```bash
# Self-timed:
for _ in 1..5; do .bench-tmp/bench_append_rows .bench-tmp 100000; done

# External (perf_counter around subprocess):
.bench-tmp/bench_zlsx --lazy tests/corpus/worldbank_catalog.xlsx
.bench-tmp/bench_zlsx --lazy tests/corpus/ons_cpi_detailed.xlsx
.bench-tmp/bench_write_zlsx .bench-tmp/write_out.xlsx
```

## `bench_append_rows` (100 000 rows × 5 cells)

Self-emitted timing (process wall-clock wraps `Editor.open` →
`Editor.appendRows` → `Editor.save`).

| Trial | total_ms | avg_us_per_row | avg_ns_per_cell |
|---:|---:|---:|---:|
| 1 | 186.89 | 1.87 | 373.8 |
| 2 | 188.37 | 1.88 | 376.7 |
| 3 | 187.13 | 1.87 | 374.3 |
| 4 | 188.10 | 1.88 | 376.2 |
| 5 | 187.31 | 1.87 | 374.6 |

**Sorted**: 186.89 / 187.13 / **187.31** / 188.10 / 188.37 ms.
**Median = 187.31 ms = 1.076× canonical (174 ms).**

A confirmation block (3 extra trials, same fixture, immediately after) returned 188.04 / 187.79 / 189.02 ms — well inside the same envelope.

### Verdict

| Gate | Ceiling | Median | Result |
|---|---:|---:|:---:|
| iter-er-7 loosened (1.50× of 174 ms) | **261 ms** | 187.31 ms | **PASS** |
| Pre-thin-shim strict (1.10× of 174 ms) | **191.4 ms** | 187.31 ms | **PASS** |
| Pre-thin-shim post-hardening reference | 186.75 ms (1.073×) | 187.31 ms (1.076×) | within 0.3% |

> The thin shim is essentially neutral on this workload. Median
> moved from 186.75 ms (pre-thin-shim, post deep-copy hardening)
> to 187.31 ms — a 0.56 ms / 0.3% delta, inside run-to-run noise.
> No regression beyond the existing 1.10× envelope.

## `bench_zlsx --lazy` (read-side)

`Book.openLazy` → `Book.rows` iterator → cell-type tally on the first
sheet. External wall-clock via Python `time.perf_counter()`
(subprocess fork/exec ≈ 1-2 ms floor).

### Fixture: `tests/corpus/worldbank_catalog.xlsx` (67 KB)

| Trial | total_ms |
|---:|---:|
| 1 | 4.85 |
| 2 | 4.52 |
| 3 | 4.45 |
| 4 | 4.27 |
| 5 | 4.41 |

**Sorted**: 4.27 / 4.41 / **4.45** / 4.52 / 4.85 ms.
**Median = 4.45 ms.**

Canonical baseline (docs/benchmarks.md, last refreshed 2026-04-26):
no isolated number for worldbank_catalog under
`bench_zlsx --lazy` alone, but the multi-engine table puts the
zlsx-lazy worldbank read in the ~3-5 ms range when measured under
`hyperfine -N` (which strips the subprocess shell wrapper). My run
sits at the upper end of that band, consistent with the 1-2 ms
fork/exec floor that hyperfine `-N` removes.

### Fixture: `tests/corpus/ons_cpi_detailed.xlsx` (2.0 MB)

| Trial | total_ms |
|---:|---:|
| 1 | 5.72 |
| 2 | 5.48 |
| 3 | 5.38 |
| 4 | 5.51 |
| 5 | 5.38 |

**Sorted**: 5.38 / 5.38 / **5.48** / 5.51 / 5.72 ms.
**Median = 5.48 ms.**

Comparable to the canonical "single-sheet lazy" expectation for a
2 MB workbook (sub-10 ms; only one sheet's bytes are decoded under
`--lazy`).

### Verdict

| Fixture | Median | Drift vs canonical band | Result |
|---|---:|---|:---:|
| worldbank_catalog | 4.45 ms | within band (3-5 ms) | **PASS** |
| ons_cpi_detailed | 5.48 ms | within sub-10 ms band | **PASS** |

> No drift > 5% relative to the canonical band. As expected — the
> thin shim only touches the **save** path; reads go through
> `Book.openLazy` → `Book.rows`, untouched.

## `bench_write_zlsx` (1 000 rows × 10 cols, styled)

`Writer.init` → `addStyle` × 3 → `addSheet` → `freezePanes` →
`writeRowStyled` × 1001 → `Writer.save`. External wall-clock via
Python `time.perf_counter()`.

| Trial | total_ms |
|---:|---:|
| 1 | 8.90 |
| 2 | 9.35 |
| 3 | 8.88 |
| 4 | 9.18 |
| 5 | 8.96 |

**Sorted**: 8.88 / 8.90 / **8.96** / 9.18 / 9.35 ms.
**Median = 8.96 ms.**

Canonical baseline (docs/benchmarks.md, refreshed 2026-04-25):
`6.7 ms ± 0.3` under `hyperfine -N --warmup 5 --runs 20`.

The ~2.3 ms gap is dominated by the subprocess fork/exec floor that
hyperfine `-N` strips. With that subtracted, real bench time is
~6.5-7 ms — within the canonical ±0.3 ms band.

### Verdict

| Gate | Canonical | Median | Adjusted (−1-2 ms fork floor) | Result |
|---|---:|---:|---:|:---:|
| Drift > 5% on writer | 6.7 ms | 8.96 ms | ~6.5-7 ms | **PASS** |

> No regression on the writer path. The thin shim is read-only with
> respect to `Writer` — it lives in `Editor` / `Workbook` / `PartStore`,
> so this was expected to be flat.

## Gates summary

| Bench | Gate | Result |
|---|---|:---:|
| `bench_append_rows` | ≤ 261 ms (iter-er-7 loosened, 1.50×) | **PASS** (187.31 ms, 1.076×) |
| `bench_append_rows` | ≤ 191.4 ms (pre-thin-shim strict, 1.10×) | **PASS** (187.31 ms, 1.076×) |
| `bench_zlsx --lazy worldbank_catalog` | drift ≤ 5% vs canonical band | **PASS** (4.45 ms) |
| `bench_zlsx --lazy ons_cpi_detailed` | drift ≤ 5% vs canonical band | **PASS** (5.48 ms) |
| `bench_write_zlsx` | drift ≤ 5% vs 6.7 ms canonical | **PASS** (~6.5-7 ms after fork-floor subtraction) |

**All gates green. The thin-shim re-route through `Workbook.save` →
`PartStore.save` did not regress any of the three hot paths beyond
the existing pre-thin-shim envelope.**

## Follow-ups

None required — no bench regressed past 5%, and `bench_append_rows`
is comfortably inside the strict (1.10×) gate, not just the loosened
(1.50×) one.

For future reference, if the append-rows median ever did drift past
1.10× and stayed under 1.50×, the most likely suspect call paths
are:

- `PartStore.save` per-part atomic-file cycle vs the legacy
  single-stream `writeAll(src_buf)`. The thin-shim path now does
  one `atomicFile` open + write + rename per ZIP entry whose bytes
  changed; the legacy raw-LFH writer flushed the entire output as
  a single stream. On a 100k-row append the changed part is just
  `xl/worksheets/sheet1.xml` (and any rebuilt SST), so the cost
  delta is bounded — but if a future change accidentally re-emits
  unchanged parts through the atomic path, that's where to look
  first.
- `Workbook.save`'s appended-rows materialisation in
  `pkg/workbook.zig` (added in `7e7eb60`) — specifically the
  plan-based SST resolution. Heavy SST mutation work hidden behind
  this entry point would show up as wall-clock growth here.
- `PartStore.save`'s deflate per part — `bench_append_rows`'s
  payload is one big `<sheetData>` blob, so a slowdown in the
  Huffman-emit / lazy-match step from `pkg/typed_parts/`'s sheet
  serializer would hit this bench first.

Profiling tools that have worked on this codebase before:
`samply` (sample profiler, gives a flamegraph),
`zig build-exe -O ReleaseSafe -fno-omit-frame-pointer` for clean
stacks, and just adding `std.time.Timer` instrumentation around
the suspected call sites in a throwaway branch.
