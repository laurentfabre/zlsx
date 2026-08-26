#!/usr/bin/env python3
"""
Bench-regression comparator.

Reads two `bench.json` files produced by `scripts/bench_ci.sh`
(base SHA + current HEAD), pairs entries by `command` name, and
emits a Markdown report.

Two modes, and the difference between them is the whole point (§9 of
goal_formula.md, extended at M5d3):

  report-only (default)
      A3's original contract. Compares **means**, flags regressions in
      the Markdown body, and **always exits 0**. This is what CI runs:
      a shared GitHub runner's wall clock is not a gate, and a bench job
      that can block a PR on thermal noise gets disabled within a month.

  --gate
      Release-cut mode. Compares **medians**, and **exits 1** if any
      benchmark regressed. A median ignores the one run that hit a
      migration stall; a mean does not, which is exactly why the
      report-only lane can afford means and a blocking lane cannot.

Regression rule (identical in both modes; only the location statistic
and the exit code differ):

- delta_pct      = (current.LOC - base.LOC) / base.LOC * 100
- pooled_sigma   = sqrt(base.stddev^2 + current.stddev^2)
- z              = (current.LOC - base.LOC) / pooled_sigma  (0 if sigma=0)
- absolute       = current.LOC - base.LOC (seconds)

Flag when ALL hold:
  delta_pct >= 10
  z         >= 3
  absolute  >= 0.0005     (avoid noise on sub-ms changes)

Exit codes:
  0  — comparison ran; no regression, or report-only mode (which never
       fails whatever it found)
  1  — `--gate` only: at least one benchmark regressed
  2  — malformed JSON or no overlapping benchmark names

Usage:
  compare_bench.py <base.json> <current.json> [--markdown OUT.md]
                   [--gate] [--baseline-commit SHA]
"""
from __future__ import annotations

import argparse
import json
import math
import sys
from pathlib import Path


REGRESSION_DELTA_PCT = 10.0
REGRESSION_Z = 3.0
REGRESSION_ABS_SECONDS = 0.0005


def load_runs(path: Path) -> dict[str, dict]:
    try:
        data = json.loads(path.read_text())
    except (FileNotFoundError, json.JSONDecodeError) as e:
        print(f"compare_bench: cannot load {path}: {e}", file=sys.stderr)
        sys.exit(2)
    if not isinstance(data, dict) or "results" not in data:
        print(f"compare_bench: {path} is not a hyperfine results JSON", file=sys.stderr)
        sys.exit(2)
    out = {}
    for r in data["results"]:
        name = r.get("command", r.get("command_name", ""))
        if not name:
            continue
        out[name] = r
    return out


def percentile(values: list[float], q: float) -> float:
    """Linear-interpolated percentile. `q` in [0, 1]."""
    if not values:
        return 0.0
    xs = sorted(values)
    if len(xs) == 1:
        return xs[0]
    pos = q * (len(xs) - 1)
    lo = math.floor(pos)
    hi = math.ceil(pos)
    if lo == hi:
        return xs[int(pos)]
    return xs[lo] + (xs[hi] - xs[lo]) * (pos - lo)


def location(run: dict, use_median: bool) -> float:
    """The statistic the comparison is made on.

    hyperfine reports `median` directly; `times` is the raw sample and is
    the fallback when an older export omits it. Falling back to the mean
    would silently turn a gate run into a report-only comparison wearing
    a gate's exit code, so the fallback chain never reaches the mean
    while `use_median` is set.
    """
    if not use_median:
        return float(run["mean"])
    if "median" in run:
        return float(run["median"])
    times = run.get("times") or []
    if times:
        return percentile([float(t) for t in times], 0.5)
    print(
        "compare_bench: --gate needs a `median` or `times` field; "
        f"this JSON has neither for one of its results",
        file=sys.stderr,
    )
    sys.exit(2)


def p95(run: dict) -> float | None:
    """§9 asks for the distribution alongside the gating statistic."""
    times = run.get("times") or []
    if not times:
        return None
    return percentile([float(t) for t in times], 0.95)


def fmt_ms(seconds: float) -> str:
    return f"{seconds * 1000:.2f} ms"


def fmt_opt_ms(seconds: float | None) -> str:
    return "—" if seconds is None else fmt_ms(seconds)


def fmt_pct(pct: float) -> str:
    sign = "+" if pct >= 0 else ""
    return f"{sign}{pct:.1f}%"


def main() -> int:
    p = argparse.ArgumentParser()
    p.add_argument("base", type=Path)
    p.add_argument("current", type=Path)
    p.add_argument("--markdown", type=Path, default=None)
    p.add_argument(
        "--gate",
        action="store_true",
        help="Release-cut mode: compare medians and exit 1 on regression.",
    )
    p.add_argument(
        "--baseline-commit",
        default=None,
        help="Commit the base measurements came from; recorded in the report.",
    )
    args = p.parse_args()

    base = load_runs(args.base)
    current = load_runs(args.current)

    overlap = sorted(set(base) & set(current))
    if not overlap:
        print("compare_bench: no overlapping benchmark names between base + current", file=sys.stderr)
        sys.exit(2)

    rows = []
    regressions = 0

    for name in overlap:
        b = base[name]
        c = current[name]
        b_loc = location(b, args.gate)
        c_loc = location(c, args.gate)
        b_sd = float(b.get("stddev", 0))
        c_sd = float(c.get("stddev", 0))

        delta_abs = c_loc - b_loc
        delta_pct = (delta_abs / b_loc * 100) if b_loc > 0 else 0.0
        pooled_sigma = math.sqrt(b_sd ** 2 + c_sd ** 2)
        z = (delta_abs / pooled_sigma) if pooled_sigma > 0 else 0.0

        flag = (
            delta_pct >= REGRESSION_DELTA_PCT
            and z >= REGRESSION_Z
            and delta_abs >= REGRESSION_ABS_SECONDS
        )
        if flag:
            regressions += 1
        verdict = "⚠️ regression" if flag else "ok"
        rows.append((name, b_loc, c_loc, p95(c), delta_pct, z, verdict))

    stat = "median" if args.gate else "mean"
    md_lines = [
        "<!-- zlsx-bench-ci -->",
        "### Benchmark Regression Report",
        "",
        f"Comparing **{args.base.parent.name}** → **{args.current.parent.name}**.",
        "",
        f"Threshold: ≥ {REGRESSION_DELTA_PCT:.0f}% slowdown AND ≥ {REGRESSION_Z:.0f}σ AND ≥ {REGRESSION_ABS_SECONDS*1000:.1f} ms absolute.",
    ]
    if args.baseline_commit:
        md_lines.append(f"Baseline commit: `{args.baseline_commit}`.")
    if args.gate:
        md_lines.append(
            "**Gate mode** — comparison on medians; a flagged benchmark fails this run "
            "(release cuts, per §9)."
        )
    else:
        md_lines.append(
            "Report-only — comparison on means; warnings do not fail CI "
            "(see docs/plans/archive/post-0.2.9-roadmap.md, A3)."
        )
    md_lines += [
        "",
        f"| Benchmark | Base {stat} | Current {stat} | Current p95 | Δ | z | Result |",
        "|---|---:|---:|---:|---:|---:|---|",
    ]
    for name, b_loc, c_loc, c_p95, delta_pct, z, verdict in rows:
        md_lines.append(
            f"| {name} | {fmt_ms(b_loc)} | {fmt_ms(c_loc)} | {fmt_opt_ms(c_p95)} | "
            f"{fmt_pct(delta_pct)} | {z:.1f} | {verdict} |"
        )
    md_lines.append("")
    if regressions:
        tail = (
            "Release cut blocked."
            if args.gate
            else "Investigate but CI stays green per the report-only gate."
        )
        md_lines.append(f"⚠️ **{regressions} benchmark(s) flagged as potential regression.** {tail}")
    else:
        md_lines.append("All measured benchmarks within threshold.")
    md_lines.append("")

    md = "\n".join(md_lines)

    if args.markdown:
        args.markdown.parent.mkdir(parents=True, exist_ok=True)
        args.markdown.write_text(md)
        print(f"compare_bench: wrote {args.markdown}")
    else:
        print(md)

    # The one behavioural difference. Report-only returns 0 whatever it
    # found — that is its contract, not an oversight.
    return 1 if (args.gate and regressions) else 0


if __name__ == "__main__":
    sys.exit(main())
