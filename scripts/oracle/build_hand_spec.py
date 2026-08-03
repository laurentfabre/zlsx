#!/usr/bin/env python3
"""
Build the hand-derived spec manifests (M1b, goal_formula.md §8.2).

The hand-derived suite is the only oracle leg whose values come from
READING rather than RUNNING. That is what makes it the tie-breaker: when
Excel and LibreOffice disagree, an independently-derived expectation is
the thing that says which of them is answering the question we asked.

It is therefore also the leg where guessing is most damaging. Every case
below cites the rule it comes from, and cases whose answer cannot be
defended from a documented rule are deliberately ABSENT — a hand-spec
entry that is really a recollection of what Excel probably does is worse
than no entry, because precedence will let it outrank the corpus and
LibreOffice.

Two manifests, because §8.2 makes precedence fidelity-specific:

  hand_spec_excel.json (.excel)  documented Excel semantics: which error
                                 a domain failure produces, how text and
                                 booleans coerce, operator precedence.
  hand_spec_ieee.json  (.ieee)   IEEE-754 bit goldens. These LEAD in
                                 `.ieee` mode, with Excel retained only
                                 as a recorded divergence witness.

Usage:
    scripts/oracle/build_hand_spec.py --input tests/oracle/inputs/oracle_suite.xlsx \\
        --out tests/oracle/fixtures
"""
from __future__ import annotations

import argparse
import datetime
import hashlib
import json
import struct
from pathlib import Path

EXTRACTOR_VERSION = "oracle-extractor-1"
SCHEMA = "zlsx-oracle-manifest-1"
SHEET = "Spec"


def bits(x: float) -> str:
    return "0x%016X" % struct.unpack("<Q", struct.pack("<d", x))[0]


def number(ref: str, value: float, formula: str) -> dict:
    return {
        "sheet": SHEET,
        "ref": ref,
        "kind": "number",
        "bits": bits(value),
        "source": repr(value),
        "formula": formula,
    }


def error(ref: str, kind: str, spelling: str, formula: str) -> dict:
    return {
        "sheet": SHEET,
        "ref": ref,
        "kind": "error",
        "error_kind": kind,
        "error_spelling": spelling,
        "formula": formula,
    }


def text(ref: str, value: str, formula: str) -> dict:
    return {"sheet": SHEET, "ref": ref, "kind": "text", "text": value, "formula": formula}


def boolean(ref: str, value: bool, formula: str) -> dict:
    return {"sheet": SHEET, "ref": ref, "kind": "boolean", "boolean": value, "formula": formula}


# ─── .excel fidelity: documented Excel semantics ─────────────────
#
# Each entry names the rule. Where the rule lives in goal_formula.md
# (which froze it after its own review), that is the citation.

EXCEL_CASES = [
    # Arithmetic that no dialect disputes — the baseline that proves the
    # suite is aligned with the workbook at all.
    number("A1", 2.0, "1+1"),
    # goal_formula.md §5.2 precedence table: `^` is LEFT-associative, so
    # 2^3^2 is (2^3)^2 = 64. Most languages give 512.
    number("A4", 64.0, "2^3^2"),
    # §5.2: unary minus binds tighter than `^`, so -1^2 is (-1)^2 = 1.
    number("A5", 1.0, "-1^2"),
    # §10 error taxonomy: division by zero is a producible classic error.
    error("A6", "div0", "#DIV/0!", "1/0"),
    # SQRT of a negative is a DOMAIN failure, not a type failure: Excel
    # documents #NUM! for it. LibreOffice returns #VALUE!, which is why
    # this case is in the suite.
    error("A7", "num", "#NUM!", "SQRT(-1)"),
    # Text that cannot coerce to a number in arithmetic → #VALUE!.
    error("A8", "value", "#VALUE!", '"a"+1'),
    # TRUE coerces to 1 in arithmetic context.
    number("A9", 2.0, "TRUE()+1"),
    # `&` coerces its operands to text.
    text("A10", "1x", '1&"x"'),
    # Excel's text comparison is case-INSENSITIVE, so "a" < "B" compares
    # a<b and is TRUE. A case-sensitive engine gives FALSE ("B" < "a" in
    # code-point order).
    boolean("A12", True, '"a"<"B"'),
    # The percent operator divides by 100.
    number("A13", 0.1, "10%"),
    # Excel refuses to produce infinity: arithmetic overflow is #NUM!.
    error("A14", "num", "#NUM!", "1E+308*10"),
]

# ─── .ieee fidelity: bit goldens ─────────────────────────────────
#
# Computed here rather than typed, so the constants cannot be mistyped.
# These lead in `.ieee` mode; Excel is a witness only (§8.2).

IEEE_CASES = [
    # The canonical case. 0.1+0.2 is NOT 0.3 in binary64 — it is one ULP
    # above. An engine that returns exactly 0.3 has rounded somewhere.
    number("A2", 0.1 + 0.2, "0.1+0.2"),
    # 1/3 to the last bit.
    number("A3", 1 / 3, "1/3"),
    # …and the comparison that follows from it. Under exact IEEE
    # semantics this is FALSE. Excel and LibreOffice both return TRUE,
    # because both fuzz comparisons near zero — a divergence the `.ieee`
    # leg exists to state plainly rather than accommodate.
    boolean("A11", (0.1 + 0.2) == 0.3, "(0.1+0.2)=0.3"),
    # Subnormal territory: a result too small for a normal double.
    number("A15", 1e-308 / 1e10, "1E-308/10000000000"),
    # Signed zero is a distinct bit pattern and must survive as one.
    number("A16", -0.0, "-0"),
    # The smallest positive subnormal.
    number("A17", 2.0**-1074, "2^-1074"),
    # Below the epsilon threshold: adding it changes nothing.
    number("A18", 1 + 1e-16, "1+1E-016"),
]


def manifest(case: str, fidelity: str, cells: list[dict], digest: str, citation: str) -> dict:
    return {
        "schema": SCHEMA,
        "case": case,
        "fidelity": fidelity,
        "provenance": {
            "adapter": "hand_spec",
            # For a leg that runs nothing, the "build" is the authority
            # the values were derived from. Blank would be a lie of
            # omission and `Record.validate` refuses it outright.
            "app_build": citation,
            "os": "n/a (derived, not executed)",
            "locale": "en_US.UTF-8",
            "extractor_version": EXTRACTOR_VERSION,
            # The INPUT workbook: these values answer the formulas in
            # that file, so that file is what identifies them.
            "workbook_digest": digest,
            "recorded": datetime.date.today().isoformat(),
        },
        "calc": {
            "calc_mode": "auto",
            "full_calc_on_load": False,
            "full_precision": True,
            "date1904": False,
        },
        "cells": cells,
    }


def main() -> int:
    p = argparse.ArgumentParser()
    p.add_argument("--input", type=Path, required=True, help="the oracle input workbook")
    p.add_argument("--out", type=Path, default=Path("tests/oracle/fixtures"))
    args = p.parse_args()

    digest = hashlib.sha256(args.input.read_bytes()).hexdigest()
    args.out.mkdir(parents=True, exist_ok=True)

    written = []
    for name, fidelity, cases, citation in (
        (
            "hand_spec_excel.json",
            "excel",
            EXCEL_CASES,
            "goal_formula.md §5.2 precedence + §10 error taxonomy; MS-OE376 operator semantics",
        ),
        (
            "hand_spec_ieee.json",
            "ieee",
            IEEE_CASES,
            "IEEE 754-2019 binary64; goldens computed, not transcribed",
        ),
    ):
        path = args.out / name
        path.write_text(
            json.dumps(manifest("oracle_suite", fidelity, cases, digest, citation), indent=2) + "\n"
        )
        written.append((path, len(cases)))

    for path, n in written:
        print(f"{path}  {n} cases")
    print(f"input digest: {digest}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
