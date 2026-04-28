#!/usr/bin/env python3
"""
A3 (post-0.2.9 roadmap): bench-regression comparator.

Reads two `bench.json` files produced by `scripts/bench_ci.sh`
(base SHA + current HEAD), pairs entries by `command` name, and
emits a Markdown report flagging regressions per the report-only
contract documented in docs/plans/post-0.2.9-roadmap.md (A3).

Regression flag (warning, not failure):
- delta_pct      = (current.mean - base.mean) / base.mean * 100
- pooled_sigma   = sqrt(base.stddev^2 + current.stddev^2)
- z              = (current.mean - base.mean) / pooled_sigma  (0 if sigma=0)
- absolute       = current.mean - base.mean (seconds)

Flag when ALL hold:
  delta_pct >= 10
  z         >= 3
  absolute  >= 0.0005     (avoid noise on sub-ms changes)

Exit codes:
  0  — comparison ran (regression flag is in the Markdown body, never
       fails CI per the report-only gate)
  2  — malformed JSON or no overlapping benchmark names

Usage:
  compare_bench.py <base.json> <current.json> [--markdown OUT.md]
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


def fmt_ms(seconds: float) -> str:
    return f"{seconds * 1000:.2f} ms"


def fmt_pct(pct: float) -> str:
    sign = "+" if pct >= 0 else ""
    return f"{sign}{pct:.1f}%"


def main() -> int:
    p = argparse.ArgumentParser()
    p.add_argument("base", type=Path)
    p.add_argument("current", type=Path)
    p.add_argument("--markdown", type=Path, default=None)
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
        b_mean = float(b["mean"])
        c_mean = float(c["mean"])
        b_sd = float(b.get("stddev", 0))
        c_sd = float(c.get("stddev", 0))

        delta_abs = c_mean - b_mean
        delta_pct = (delta_abs / b_mean * 100) if b_mean > 0 else 0.0
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
        rows.append((name, b_mean, c_mean, delta_pct, z, verdict))

    md_lines = [
        "<!-- zlsx-bench-ci -->",
        "### Benchmark Regression Report",
        "",
        f"Comparing **{args.base.parent.name}** → **{args.current.parent.name}**.",
        "",
        f"Threshold: ≥ {REGRESSION_DELTA_PCT:.0f}% slowdown AND ≥ {REGRESSION_Z:.0f}σ AND ≥ {REGRESSION_ABS_SECONDS*1000:.1f} ms absolute.",
        "Report-only — warnings do not fail CI (see docs/plans/post-0.2.9-roadmap.md, A3).",
        "",
        "| Benchmark | Base mean | Current mean | Δ | z | Result |",
        "|---|---:|---:|---:|---:|---|",
    ]
    for name, b_mean, c_mean, delta_pct, z, verdict in rows:
        md_lines.append(
            f"| {name} | {fmt_ms(b_mean)} | {fmt_ms(c_mean)} | {fmt_pct(delta_pct)} | {z:.1f} | {verdict} |"
        )
    md_lines.append("")
    if regressions:
        md_lines.append(
            f"⚠️ **{regressions} benchmark(s) flagged as potential regression.** "
            "Investigate but CI stays green per the report-only gate."
        )
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

    return 0


if __name__ == "__main__":
    sys.exit(main())
