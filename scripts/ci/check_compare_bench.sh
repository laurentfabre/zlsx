#!/usr/bin/env bash
# M5d3 gate: `compare_bench.py`'s two exit contracts.
#
# §9 (goal_formula.md) splits the comparator in two: CI stays
# report-only and always exits 0, while release cuts run `--gate`, which
# compares medians and exits nonzero on a regression. Both halves of
# that are load bearing, and neither is visible from a passing bench job
# — a gate that silently never fires and a report-only lane that
# suddenly starts failing PRs look identical until the day they matter.
#
# So both are asserted here against synthetic hyperfine JSON: an
# injected regression, an injected improvement, and the report-only path
# over the same regression.
#
# No arguments; writes into a temp dir it removes on exit.

set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
CMP="$ROOT/scripts/compare_bench.py"
WORK="$(mktemp -d)"
trap 'rm -rf "$WORK"' EXIT

# A hyperfine-shaped result. `times` is the raw sample: the gate reads
# `median`, and p95 is computed off `times`, so both fields have to be
# present for this to exercise the real path rather than a fallback.
write_json() {
    local path="$1" median="$2" stddev="$3"
    python3 - "$path" "$median" "$stddev" <<'PY'
import json, sys
path, median, stddev = sys.argv[1], float(sys.argv[2]), float(sys.argv[3])
# Nine samples straddling the median by ±stddev, so `median` and the
# `times` array agree and p95 is a real order statistic.
times = [median + (i - 4) * (stddev / 2) for i in range(9)]
json.dump(
    {
        "results": [
            {
                "command": "recalc:f1_mix_small",
                "mean": median,
                "median": median,
                "stddev": stddev,
                "min": min(times),
                "max": max(times),
                "times": times,
            }
        ]
    },
    open(path, "w"),
    indent=2,
)
PY
}

fail() { echo "FAIL: $*" >&2; exit 1; }

# 100 ms ± 1 ms baseline. The rule needs ≥10% AND ≥3σ AND ≥0.5 ms, so:
#   regressed   150 ms  →  +50%, ~35σ, +50 ms   → flagged
#   improved     80 ms  →  −20%                 → not flagged
write_json "$WORK/base.json"      0.100 0.001
write_json "$WORK/regressed.json" 0.150 0.001
write_json "$WORK/improved.json"  0.080 0.001

# 1. --gate on a regression: nonzero.
set +e
python3 "$CMP" "$WORK/base.json" "$WORK/regressed.json" --gate > "$WORK/gate_regress.md"
gate_regress_rc=$?
set -e
[ "$gate_regress_rc" -ne 0 ] || fail "--gate exited 0 on an injected regression"
grep -q "regression" "$WORK/gate_regress.md" || fail "--gate report does not name the regression"
grep -q "Base median" "$WORK/gate_regress.md" || fail "--gate did not compare medians"

# 2. --gate on an improvement: zero.
set +e
python3 "$CMP" "$WORK/base.json" "$WORK/improved.json" --gate > "$WORK/gate_improve.md"
gate_improve_rc=$?
set -e
[ "$gate_improve_rc" -eq 0 ] || fail "--gate exited $gate_improve_rc on an injected improvement"

# 3. Report-only over the SAME regression: still zero, and still means.
#    Same input as case 1, so the only variable is the mode.
set +e
python3 "$CMP" "$WORK/base.json" "$WORK/regressed.json" > "$WORK/report.md"
report_rc=$?
set -e
[ "$report_rc" -eq 0 ] || fail "report-only exited $report_rc on a regression; it must always exit 0"
grep -q "regression" "$WORK/report.md" || fail "report-only did not flag the regression in its body"
grep -q "Base mean" "$WORK/report.md" || fail "report-only did not compare means"
grep -q "Report-only" "$WORK/report.md" || fail "report-only did not label itself"

# 4. A malformed input is still exit 2 in both modes — the pre-existing
#    contract the two new exit codes must not have displaced.
echo 'not json' > "$WORK/bad.json"
set +e
python3 "$CMP" "$WORK/bad.json" "$WORK/regressed.json" --gate > /dev/null 2>&1
bad_rc=$?
set -e
[ "$bad_rc" -eq 2 ] || fail "malformed JSON exited $bad_rc, expected 2"

echo "OK: --gate fails on regression (rc=$gate_regress_rc), passes on improvement,"
echo "    report-only always exits 0, malformed input still exits 2."
