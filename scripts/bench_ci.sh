#!/usr/bin/env bash
# A3 (post-0.2.9 roadmap): bench-regression CI harness — report-only.
#
# Builds the zlsx read + write benchmark binaries in ReleaseFast,
# runs them under hyperfine on a small fixed fixture set, and emits
# a JSON file consumable by scripts/compare_bench.py.
#
# Usage: bench_ci.sh <out_dir>
#
# Expects:
#   - hyperfine on PATH
#   - zig 0.15.2 on PATH
#   - tests/corpus/{worldbank_catalog,ons_cpi_detailed}.xlsx already
#     fetched (cache layer in .github/workflows/bench.yml)
#
# Output:
#   <out_dir>/bench.json    aggregated hyperfine run results
#   <out_dir>/bench.md      pretty Markdown table for human review

set -euo pipefail

OUT_DIR="${1:-bench-out/current}"
mkdir -p "$OUT_DIR"

ZIG_BIN="${ZIG_BIN:-zig}"

echo "[bench-ci] building benchmark binaries (ReleaseFast)..."
# Note: -Mroot specifies the entry module — no positional path argument
# ("zig build-exe FILE" + "-Mroot=FILE" together is rejected as duplicate).
"$ZIG_BIN" build-exe \
    --dep zlsx -Mroot=tests/bench/bench_zlsx.zig \
    -Mzlsx=src/xlsx.zig \
    -O ReleaseFast \
    -femit-bin="$OUT_DIR/bench_zlsx_read"
"$ZIG_BIN" build-exe \
    --dep zlsx -Mroot=tests/bench/bench_write_zlsx.zig \
    -Mzlsx=src/xlsx.zig \
    -O ReleaseFast \
    -femit-bin="$OUT_DIR/bench_zlsx_write"

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

# Merge the two JSON outputs into a single bench.json with one
# `results[]` array. Schema follows hyperfine's own format
# (https://github.com/sharkdp/hyperfine/blob/master/scripts/README.md).
python3 - <<'PY' "$OUT_DIR/hyperfine_read.json" "$OUT_DIR/hyperfine_write.json" "$OUT_DIR/bench.json"
import json, sys

merged = {"results": []}
for src in sys.argv[1:-1]:
    with open(src) as f:
        merged["results"].extend(json.load(f).get("results", []))

with open(sys.argv[-1], "w") as f:
    json.dump(merged, f, indent=2)
PY

# Pretty Markdown summary (human-readable; CI artefact).
python3 - <<'PY' "$OUT_DIR/bench.json" "$OUT_DIR/bench.md"
import json, sys

with open(sys.argv[1]) as f:
    data = json.load(f)

lines = ["| Benchmark | Mean | Stddev | Min | Max |", "|---|---:|---:|---:|---:|"]
for r in data["results"]:
    name = r.get("command", r.get("command_name", "<unnamed>"))
    mean_ms = r["mean"] * 1000
    sd_ms = r.get("stddev", 0) * 1000
    min_ms = r["min"] * 1000
    max_ms = r["max"] * 1000
    lines.append(f"| {name} | {mean_ms:.2f} ms | {sd_ms:.2f} ms | {min_ms:.2f} ms | {max_ms:.2f} ms |")

with open(sys.argv[2], "w") as f:
    f.write("\n".join(lines) + "\n")
PY

echo "[bench-ci] wrote: $OUT_DIR/bench.json  $OUT_DIR/bench.md"
