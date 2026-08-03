#!/usr/bin/env python3
"""
Build the oracle harness's INPUT workbooks (M1b, goal_formula.md §8.2).

These files are what Excel and LibreOffice are asked to recalculate.
Each one carries formulas plus deliberately wrong cached values, so a
run that never actually recalculated can be detected and rejected.

Why this is hand-authored XML rather than written with zlsx
-----------------------------------------------------------
Two reasons, and the first is the important one:

 1. The whole point of a sentinel is a cached `<v>` that CONTRADICTS its
    own `<f>`. No sane writer offers that; zlsx's writer will not emit
    a formula with a knowingly-wrong cached value, and it should not.
 2. The oracle exists to check zlsx. Building its inputs with zlsx would
    put the implementation under test on both sides of the experiment.

The same reasoning applies to `tests/oracle/zip_reader.zig` and
`xml_scan.zig` on the reading side. Ordinary xlsx work in this repo
should still go through zlsx; this is the documented exception.

Usage:
    scripts/oracle/build_inputs.py --out tests/oracle/inputs
"""
from __future__ import annotations

import argparse
import hashlib
import zipfile
from dataclasses import dataclass, field
from pathlib import Path


# ─── minimal OOXML scaffolding ───────────────────────────────────

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
{sheet_overrides}
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
{calcchain_override}
</Types>"""

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""

STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="1"><fill><patternFill patternType="none"/></fill></fills>
<borders count="1"><border/></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
</styleSheet>"""


def row_of(ref: str) -> int:
    return int("".join(ch for ch in ref if ch.isdigit()))


def col_of(ref: str) -> int:
    """Bijective base-26 column index from an A1 reference."""
    n = 0
    for ch in ref:
        if not ch.isalpha():
            break
        n = n * 26 + (ord(ch.upper()) - ord("A") + 1)
    return n


def xml_escape(s: str) -> str:
    return (
        s.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


@dataclass
class Cell:
    ref: str
    # Exactly one of these three shapes:
    #   literal number/string  -> value set, formula None
    #   formula + cached value -> both set
    #   formula, no cache      -> formula set, value None
    value: str | None = None
    formula: str | None = None
    # OOXML `t` attribute: None (number), "s"/"str"/"b"/"e"/"inlineStr".
    kind: str | None = None

    def to_xml(self) -> str:
        attrs = f' r="{self.ref}"'
        if self.kind:
            attrs += f' t="{self.kind}"'
        body = ""
        if self.formula is not None:
            body += f"<f>{xml_escape(self.formula)}</f>"
        if self.value is not None:
            body += f"<v>{xml_escape(self.value)}</v>"
        if not body:
            return f"<c{attrs}/>"
        return f"<c{attrs}>{body}</c>"


@dataclass
class Sheet:
    name: str
    cells: list[Cell] = field(default_factory=list)

    def to_xml(self) -> str:
        rows: dict[int, list[Cell]] = {}
        for c in self.cells:
            rows.setdefault(row_of(c.ref), []).append(c)
        # Rows ascending by index, cells ascending by column WITHIN each
        # row. Excel enforces both and refuses the file outright — with
        # no error an automated caller can see — if either is violated.
        # zlsx and LibreOffice both accept the unsorted form, so this is
        # a constraint only Excel reveals, and only by silence.
        body = "".join(
            f'<row r="{r}">'
            + "".join(c.to_xml() for c in sorted(rows[r], key=lambda c: col_of(c.ref)))
            + "</row>"
            for r in sorted(rows)
        )
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            f"<sheetData>{body}</sheetData></worksheet>"
        )


@dataclass
class Workbook:
    sheets: list[Sheet]
    # calcChain entries as (ref, sheet_index_1based). Order is the point:
    # a deliberately wrong order is what the stale-dependency sentinel
    # rides on, since `CalculateFull` walks the recorded chain while
    # `CalculateFullRebuild` discards it and rebuilds the edges.
    calc_chain: list[tuple[str, int]] = field(default_factory=list)
    # `calcMode="manual"` would let an app skip calculating entirely;
    # oracle inputs are always automatic so the only reason a sentinel
    # survives is that the driver failed.
    full_calc_on_load: bool = False

    def workbook_xml(self) -> str:
        sheets = "".join(
            f'<sheet name="{xml_escape(s.name)}" sheetId="{i + 1}" r:id="rId{i + 1}"/>'
            for i, s in enumerate(self.sheets)
        )
        calc_pr = '<calcPr calcId="191029"'
        if self.full_calc_on_load:
            calc_pr += ' fullCalcOnLoad="1"'
        calc_pr += "/>"
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            f"<sheets>{sheets}</sheets>{calc_pr}</workbook>"
        )

    def workbook_rels(self) -> str:
        rels = "".join(
            f'<Relationship Id="rId{i + 1}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            f'Target="worksheets/sheet{i + 1}.xml"/>'
            for i in range(len(self.sheets))
        )
        n = len(self.sheets)
        rels += (
            f'<Relationship Id="rId{n + 1}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
            'Target="styles.xml"/>'
        )
        if self.calc_chain:
            rels += (
                f'<Relationship Id="rId{n + 2}" '
                'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" '
                'Target="calcChain.xml"/>'
            )
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f"{rels}</Relationships>"
        )

    def calc_chain_xml(self) -> str:
        entries = "".join(f'<c r="{ref}" i="{idx}"/>' for ref, idx in self.calc_chain)
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            f"{entries}</calcChain>"
        )

    def write(self, path: Path) -> str:
        sheet_overrides = "\n".join(
            f'<Override PartName="/xl/worksheets/sheet{i + 1}.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            for i in range(len(self.sheets))
        )
        calcchain_override = (
            '<Override PartName="/xl/calcChain.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>'
            if self.calc_chain
            else ""
        )

        path.parent.mkdir(parents=True, exist_ok=True)
        # Deterministic archive: fixed timestamps and a fixed part order,
        # so rebuilding the inputs does not churn their digests and make
        # every provenance record look stale.
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:

            def add(name: str, data: str) -> None:
                info = zipfile.ZipInfo(name, date_time=(2026, 1, 1, 0, 0, 0))
                info.compress_type = zipfile.ZIP_DEFLATED
                info.external_attr = 0o600 << 16
                z.writestr(info, data)

            add(
                "[Content_Types].xml",
                CONTENT_TYPES.format(
                    sheet_overrides=sheet_overrides,
                    calcchain_override=calcchain_override,
                ),
            )
            add("_rels/.rels", ROOT_RELS)
            add("xl/workbook.xml", self.workbook_xml())
            add("xl/_rels/workbook.xml.rels", self.workbook_rels())
            for i, sheet in enumerate(self.sheets):
                add(f"xl/worksheets/sheet{i + 1}.xml", sheet.to_xml())
            add("xl/styles.xml", STYLES)
            if self.calc_chain:
                add("xl/calcChain.xml", self.calc_chain_xml())

        return hashlib.sha256(path.read_bytes()).hexdigest()


# ─── the sentinel workbook ───────────────────────────────────────
#
# Every oracle input carries these three cells. They are the run's
# receipt: if they come back as planted, the driver did not actually
# recalculate and nothing from the run may be recorded.

SENTINEL_SHEET = "Sentinels"

# Planted cached values, mirrored in tests/oracle/sentinel_set.zig.
PLANTED_STALE_VALUE = "999"
PLANTED_STALE_DEPENDENCY = "111"
PLANTED_VOLATILE = "0.123456789"


def sentinel_cells() -> list[Cell]:
    return [
        # Inputs the chain reads.
        Cell("A1", value="1"),
        Cell("A2", value="2"),
        # 1. Stale VALUE: `=1+1` cannot be 999. Any calculation fixes it.
        Cell("B1", value=PLANTED_STALE_VALUE, formula="1+1"),
        # 2. Stale DEPENDENCY: a three-deep chain whose calcChain order is
        #    deliberately inverted (see `stale_dependency_chain`). A
        #    calculation that trusts the recorded chain can evaluate the
        #    tail before its precedents; one that rebuilds the edges
        #    cannot. The planted tail value is impossible either way, so
        #    this is also a second, independent stale-value proof.
        Cell("C1", value="777", formula="A1*10"),
        Cell("C2", value="888", formula="C1+A2"),
        Cell("B2", value=PLANTED_STALE_DEPENDENCY, formula="C2*2"),
        # 3. Volatile draw: RAND() pinned to a fixed cache. LibreOffice
        #    will re-save a document it merely opened; a redrawn volatile
        #    is the only proof it calculated.
        Cell("B3", value=PLANTED_VOLATILE, formula="RAND()"),
    ]


def stale_dependency_chain(sheet_index: int) -> list[tuple[str, int]]:
    """calcChain in the WRONG order: tail first, roots last."""
    return [("B2", sheet_index), ("C2", sheet_index), ("C1", sheet_index), ("B1", sheet_index)]


# ─── the hand-derived spec suite ─────────────────────────────────
#
# Cases chosen because they are DIVERGENCE POINTS — places where an
# implementation can be self-consistent and still wrong. Each one has a
# documented expected value in tests/oracle/fixtures/hand_spec.json;
# these cells are what the applications are asked to confirm or refute.

SPEC_SHEET = "Spec"

SPEC_CASES: list[tuple[str, str, str]] = [
    # (ref, formula, why it is a divergence point)
    ("A1", "1+1", "baseline: proves the sheet calculated at all"),
    ("A2", "0.1+0.2", "binary64 representation; must not be 0.3"),
    ("A3", "1/3", "non-terminating binary expansion"),
    # goal_formula.md §5.2 pins `^` as LEFT-associative: 2^3^2 = 64.
    # Most languages and calculators give 512, so an implementation that
    # "obviously" got this right is the one to check.
    ("A4", "2^3^2", "left-associative exponentiation: 64, not 512"),
    ("A5", "-1^2", "unary minus binds tighter than ^ in Excel: 1, not -1"),
    ("A6", "1/0", "#DIV/0!"),
    ("A7", "SQRT(-1)", "#NUM!"),
    ("A8", '"a"+1', "#VALUE! from text coercion in arithmetic"),
    ("A9", "TRUE+1", "boolean coerces to 1 in arithmetic"),
    ("A10", '1&"x"', "concatenation coerces number to text"),
    # NOT `1=1.0`: LibreOffice rewrites that to `1=1` on save, which
    # deletes the question. This one survives normalisation and asks
    # something sharper — whether equality compares raw binary64 or the
    # rounded display value.
    ("A11", "(0.1+0.2)=0.3", "equality on a value that is not exactly 0.3"),
    ("A12", '"a"<"B"', "text comparison is case-insensitive in Excel"),
    ("A13", "10%", "percent operator"),
    ("A14", "1E308*10", "overflow to #NUM!, not infinity"),
    ("A15", "1E-308/1E10", "subnormal territory"),
    ("A16", "-0.0", "signed zero: does Excel preserve it?"),
    ("A17", "2^-1074", "smallest positive subnormal"),
    ("A18", "1+1E-16", "addition below the epsilon threshold"),
]


def spec_cells() -> list[Cell]:
    # No cached values at all: an app that fails to calculate produces a
    # file with no `<v>`, which the extractor reports as a blank rather
    # than as a wrong answer.
    return [Cell(ref, formula=formula) for ref, formula, _ in SPEC_CASES]


def main() -> int:
    p = argparse.ArgumentParser()
    p.add_argument("--out", type=Path, default=Path("tests/oracle/inputs"))
    args = p.parse_args()

    sentinels = Sheet(SENTINEL_SHEET, sentinel_cells())
    spec = Sheet(SPEC_SHEET, spec_cells())

    wb = Workbook(
        sheets=[sentinels, spec],
        calc_chain=stale_dependency_chain(1),
    )
    path = args.out / "oracle_suite.xlsx"
    digest = wb.write(path)
    print(f"{path}  sha256={digest}")

    # A second input with fullCalcOnLoad set, to record what changes when
    # the workbook itself asks to be recalculated. The screen treats that
    # flag as a disqualifier for CORPUS workbooks (the caches are stale
    # by the author's own admission); recording it here is how we know
    # what the flag actually does rather than assuming.
    wb_on_load = Workbook(
        sheets=[Sheet(SENTINEL_SHEET, sentinel_cells())],
        calc_chain=stale_dependency_chain(1),
        full_calc_on_load=True,
    )
    path2 = args.out / "oracle_full_calc_on_load.xlsx"
    digest2 = wb_on_load.write(path2)
    print(f"{path2}  sha256={digest2}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
