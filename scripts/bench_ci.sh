#!/usr/bin/env bash
# A3 (post-0.2.9 roadmap): bench-regression CI harness — report-only.
#
# Builds the zlsx read + write + recalc benchmark binaries in
# ReleaseFast, runs them under hyperfine on a small fixed fixture set,
# and emits a JSON file consumable by scripts/compare_bench.py.
#
# Usage: bench_ci.sh <out_dir>
#
# Expects:
#   - hyperfine on PATH
#   - zig 0.16.0 on PATH
#   - tests/corpus/{worldbank_catalog,ons_cpi_detailed}.xlsx already
#     fetched (cache layer in .github/workflows/bench.yml)
#
# Env:
#   ZLSX_BENCH_RECALC_SIZES   F1-mix geometries for the recalc lane.
#                             Default "tiny small" — the two CI can
#                             afford. Add "named" (§9's 100k-cell
#                             workload) when recording a baseline; see
#                             the recalc section below for why that is
#                             not the default.
#   ZLSX_BENCH_CRITERIA_SIZES M7b2's criteria-mix geometries (whole-
#                             column + multi-criteria scans). Default
#                             "tiny small" — `small` is that workload's
#                             identity size, so the default list is
#                             already the baseline-recording one.
#
# Output:
#   <out_dir>/bench.json    aggregated hyperfine run results
#   <out_dir>/bench.md      pretty Markdown table for human review

set -euo pipefail

OUT_DIR="${1:-bench-out/current}"
mkdir -p "$OUT_DIR"

ZIG_BIN="${ZIG_BIN:-zig}"
RECALC_SIZES="${ZLSX_BENCH_RECALC_SIZES:-tiny small}"
CRITERIA_SIZES="${ZLSX_BENCH_CRITERIA_SIZES:-tiny small}"

echo "[bench-ci] building benchmark binaries (ReleaseFast)..."
# Built through `zig build` rather than hand-rolled `build-exe` lines.
# The old form restated the entire module graph here, so every new named
# module had to be added in two places or this script broke on its own —
# and because no build step compiled these files, they silently kept a
# stale API through the whole Zig 0.16 migration. build.zig is now the
# single source of truth; `zig build test` compile-checks them too.
#
# `-Doptimize=ReleaseFast` only started taking effect at M5d3: the bench
# modules hardcoded `bench_optimize` (ReleaseSafe, correct for the RSS
# probes, wrong for this lane), so every wall-clock number recorded
# before that commit was measured in ReleaseSafe under a ReleaseFast
# label. Baselines from before it are not comparable with ones after.
"$ZIG_BIN" build bench-exes -Doptimize=ReleaseFast --prefix "$OUT_DIR/zig-out"
cp "$OUT_DIR/zig-out/bin/zlsx-bench-read"  "$OUT_DIR/bench_zlsx_read"
cp "$OUT_DIR/zig-out/bin/zlsx-bench-write" "$OUT_DIR/bench_zlsx_write"
# Fail-soft: a base tree that predates M5d3 has no recalc bench, and
# bench.yml deliberately runs this script from a HEAD snapshot against
# an older worktree. Losing the recalc lane there is better than losing
# the read/write numbers too.
if [[ -f "$OUT_DIR/zig-out/bin/zlsx-bench-recalc" ]]; then
    cp "$OUT_DIR/zig-out/bin/zlsx-bench-recalc" "$OUT_DIR/bench_zlsx_recalc"
else
    echo "[bench-ci] note: no zlsx-bench-recalc in this tree; skipping the recalc lane"
fi

# Fixture set is intentionally small + stable: worldbank_catalog
# (67 KB, 1144 SST entries — exercises the SST + row stream) and
# ons_cpi_detailed (2 MB — exercises larger sheet bodies). Each is
# committed via fetch_test_corpus.sh.
READ_F1="tests/corpus/worldbank_catalog.xlsx"
READ_F2="tests/corpus/ons_cpi_detailed.xlsx"

if [[ ! -f "$READ_F1" || ! -f "$READ_F2" ]]; then
    echo "[bench-ci] error: corpus fixtures missing; run scripts/fetch_test_corpus.sh first" >&2
    exit 2
fi

# hyperfine knobs:
#   -N           skip the intermediate shell, lower noise
#   --warmup 5   discard the first 5 runs (FS cache, JIT-style fluctuation)
#   --runs 20    20 measurements — enough to get reasonable stddev
#   --export-json -> JSON we can post-process
echo "[bench-ci] running hyperfine on read benchmarks..."
hyperfine -N --warmup 5 --runs 20 \
    --export-json "$OUT_DIR/hyperfine_read.json" \
    --command-name "read:worldbank_catalog" \
    "$OUT_DIR/bench_zlsx_read --lazy $READ_F1" \
    --command-name "read:ons_cpi_detailed" \
    "$OUT_DIR/bench_zlsx_read --lazy $READ_F2"

echo "[bench-ci] running hyperfine on write benchmark..."
# Writer bench takes `<out.xlsx>` as argv[1] and saves the workbook
# there. Use a path inside $OUT_DIR so each measurement overwrites
# the previous output (no disk-fill from --runs 20).
WRITE_OUT="$OUT_DIR/bench_write_output.xlsx"
hyperfine -N --warmup 5 --runs 20 \
    --export-json "$OUT_DIR/hyperfine_write.json" \
    --command-name "write:1000_rows_styled" \
    "$OUT_DIR/bench_zlsx_write $WRITE_OUT"
rm -f "$WRITE_OUT"

# ── recalc lane (M5d3, §9's named workload) ──────────────────────────
#
# Three modes per size, because hyperfine measures whole processes and
# §9 wants two numbers out of one workload: `open` is the fixed cost,
# `recalc` is open + evaluate, `save` is the end-to-end file
# transaction. §9's *evaluate* ceiling is therefore read as
# recalc − open, which is also why more than one size is run — a single
# point cannot separate a fixed cost from a per-cell one, and a third
# makes the slope between them visible rather than assumed.
#
# `named` (100 000 cells) is NOT in the default size list. At the M5d3
# baseline one run of it takes ~2 minutes, so N=20 is a ~1-hour job per
# mode — a release-cut / baseline-recording activity, not a per-PR one,
# and §9 already splits the lanes that way (CI report-only, release cuts
# gated). Record a baseline with:
#
#   ZLSX_BENCH_RECALC_SIZES="tiny small named" scripts/bench_ci.sh <dir>
RECALC_JSONS=()
if [[ -x "$OUT_DIR/bench_zlsx_recalc" ]]; then
    RECALC_OUT="$OUT_DIR/bench_recalc_output.xlsx"
    for size in $RECALC_SIZES; do
        fixture="$OUT_DIR/f1_mix_$size.xlsx"
        echo "[bench-ci] emitting the F1-mix fixture ($size)..."
        # `emit` verifies the named workload's digest and exits nonzero
        # on drift, which is the point: a baseline measured against
        # different bytes describes a different workbook.
        "$OUT_DIR/bench_zlsx_recalc" emit "$fixture" --size "$size"

        echo "[bench-ci] running hyperfine on recalc benchmarks ($size)..."
        hyperfine -N --warmup 5 --runs 20 \
            --export-json "$OUT_DIR/hyperfine_recalc_$size.json" \
            --command-name "recalc-open:f1_mix_$size" \
            "$OUT_DIR/bench_zlsx_recalc open $fixture" \
            --command-name "recalc-eval:f1_mix_$size" \
            "$OUT_DIR/bench_zlsx_recalc recalc $fixture" \
            --command-name "recalc-save:f1_mix_$size" \
            "$OUT_DIR/bench_zlsx_recalc save $fixture --out $RECALC_OUT"
        RECALC_JSONS+=("$OUT_DIR/hyperfine_recalc_$size.json")

        # §9's per-phase report — one instrumented run, not a
        # distribution, so it is an artefact rather than a measurement.
        "$OUT_DIR/bench_zlsx_recalc" phases "$fixture" --out "$RECALC_OUT" \
            > "$OUT_DIR/phases_$size.md"
        rm -f "$RECALC_OUT" "$fixture"
    done

    # ── criteria lane (M7b2, §9's whole-column + multi-criteria scans) ──
    #
    # Same binary, same three modes; the workload is what changes. The
    # criteria mix keeps its 512 report formulas FIXED across sizes and
    # varies only the stored rows they scan, so the two sizes separate
    # the per-stored-cell cost of the aligned cursor from everything
    # paid per formula. Guarded on the usage string because bench.yml
    # runs this script from HEAD against an older worktree, and an old
    # binary would silently emit the F1 mix under a criteria filename.
    # `|| true` because the usage path exits 2 and `pipefail` would
    # otherwise fail the probe on the very output it is reading.
    if ("$OUT_DIR/bench_zlsx_recalc" 2>&1 || true) | grep -q -- --workload; then
        for size in $CRITERIA_SIZES; do
            fixture="$OUT_DIR/criteria_mix_$size.xlsx"
            echo "[bench-ci] emitting the criteria-mix fixture ($size)..."
            # `emit` verifies the identity size's digest and exits
            # nonzero on drift, exactly like the F1 mix's named gate.
            "$OUT_DIR/bench_zlsx_recalc" emit "$fixture" --workload criteria --size "$size"

            echo "[bench-ci] running hyperfine on criteria benchmarks ($size)..."
            hyperfine -N --warmup 5 --runs 20 \
                --export-json "$OUT_DIR/hyperfine_criteria_$size.json" \
                --command-name "criteria-open:criteria_mix_$size" \
                "$OUT_DIR/bench_zlsx_recalc open $fixture" \
                --command-name "criteria-eval:criteria_mix_$size" \
                "$OUT_DIR/bench_zlsx_recalc recalc $fixture" \
                --command-name "criteria-save:criteria_mix_$size" \
                "$OUT_DIR/bench_zlsx_recalc save $fixture --out $RECALC_OUT"
            RECALC_JSONS+=("$OUT_DIR/hyperfine_criteria_$size.json")
            rm -f "$RECALC_OUT" "$fixture"
        done
    else
        echo "[bench-ci] note: this tree's recalc bench predates --workload; skipping the criteria lane"
    fi
fi

# Merge the JSON outputs into a single bench.json with one `results[]`
# array. Schema follows hyperfine's own format
# (https://github.com/sharkdp/hyperfine/blob/master/scripts/README.md).
# `${arr[@]+…}` rather than a bare `${arr[@]}`: under `set -u` an empty
# array is an unbound variable in bash 4.3 and older, which is what
# macOS still ships as /bin/bash.
python3 - <<'PY' "$OUT_DIR/hyperfine_read.json" "$OUT_DIR/hyperfine_write.json" ${RECALC_JSONS[@]+"${RECALC_JSONS[@]}"} "$OUT_DIR/bench.json"
import json, sys

merged = {"results": []}
for src in sys.argv[1:-1]:
    with open(src) as f:
        merged["results"].extend(json.load(f).get("results", []))

with open(sys.argv[-1], "w") as f:
    json.dump(merged, f, indent=2)
PY

# Pretty Markdown summary (human-readable; CI artefact).
#
# Median and p95 join the table at M5d3: §9 makes the median the gating
# statistic and asks for the distribution alongside it, and a summary
# that reported only the mean would leave a reader unable to check the
# number `compare_bench.py --gate` actually acted on.
python3 - <<'PY' "$OUT_DIR/bench.json" "$OUT_DIR/bench.md"
import json, math, sys

with open(sys.argv[1]) as f:
    data = json.load(f)


def pct(values, q):
    xs = sorted(values)
    if not xs:
        return None
    if len(xs) == 1:
        return xs[0]
    pos = q * (len(xs) - 1)
    lo, hi = math.floor(pos), math.ceil(pos)
    return xs[lo] if lo == hi else xs[lo] + (xs[hi] - xs[lo]) * (pos - lo)


def ms(v):
    return "—" if v is None else f"{v * 1000:.2f} ms"


lines = [
    "| Benchmark | Mean | Median | p95 | Stddev | Min | Max |",
    "|---|---:|---:|---:|---:|---:|---:|",
]
for r in data["results"]:
    name = r.get("command", r.get("command_name", "<unnamed>"))
    times = r.get("times") or []
    median = r.get("median", pct(times, 0.5))
    lines.append(
        f"| {name} | {ms(r['mean'])} | {ms(median)} | {ms(pct(times, 0.95))} | "
        f"{ms(r.get('stddev', 0))} | {ms(r['min'])} | {ms(r['max'])} |"
    )

with open(sys.argv[2], "w") as f:
    f.write("\n".join(lines) + "\n")
PY

echo "[bench-ci] wrote: $OUT_DIR/bench.json  $OUT_DIR/bench.md"
