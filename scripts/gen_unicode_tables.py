#!/usr/bin/env python3
"""
A1 (post-0.2.9 roadmap): Unicode table generator.

Four operating modes:

  --mode casefold:
    Reads `CaseFolding.txt` and emits a vendored Zig data file
    powering `unicode/casefold.zig`. Policy: non-Turkic full
    case fold (statuses C + F).

  --mode casing:
    Reads `UnicodeData.txt` + `SpecialCasing.txt` +
    `DerivedCoreProperties.txt` and emits the `casing_v1` tables
    powering `unicode/casing.zig` (M4f, §5.4b). Folding cannot
    implement casing — `fold("ß")` is `"ss"` where `UPPER("ß")` is
    `"SS"` — so this is a separate table with a separate policy:

      - full mappings, not simple ones: `ß` → `SS` is
        length-changing and a one-to-one table cannot express it;
      - `SpecialCasing.txt`'s **unconditional** rows override
        `UnicodeData.txt`'s simple mappings;
      - the one locale-INdependent conditional row, `Final_Sigma`,
        ships as its own runtime rule (Σ lowercases to ς at the end
        of a word, σ elsewhere), with the `Cased` / `Case_Ignorable`
        intervals it needs to decide "end of a word";
      - every locale-conditional row (`tr`, `az`, `lt`) is
        REJECTED. Turkish dotless-ı casing is a divergence zlsx
        records rather than implements: there is no locale in
        `RunInputs` to select it with (§5.4b).

  --mode nfc:
    Reads `UnicodeData.txt` + `CompositionExclusions.txt` and emits
    a vendored Zig data file powering `unicode/nfc.zig`.
    Includes:
      - canonical decomposition mappings (UnicodeData col 5,
        excluding `<compat>` prefixed entries);
      - canonical combining class (UnicodeData col 3);
      - composition pairs (inverted decomposition map minus
        Composition_Exclusions).

    Hangul (U+AC00..U+D7A3) is handled algorithmically by the
    runtime — its 11172 codepoints stay out of the table.

  --mode xid:
    Reads `DerivedCoreProperties.txt` and emits a vendored Zig data
    file powering `unicode/xid.zig` — the identifier grammar the
    formula tokenizer needs (M1a of the tier-D1 ladder). Two sorted,
    merged, non-overlapping interval tables: `XID_Start` and
    `XID_Continue`. The Zig side binary-searches them, so lookup is
    allocation-free.

    Zig's stdlib ships no XID tables, which is why this mode exists.

Usage:
    scripts/gen_unicode_tables.py --mode casefold \\
        --input /path/to/CaseFolding.txt \\
        --output unicode/tables/casefold_data.zig

    scripts/gen_unicode_tables.py --mode casing \\
        --input /path/to/UnicodeData.txt \\
        --special /path/to/SpecialCasing.txt \\
        --props /path/to/DerivedCoreProperties.txt \\
        --output unicode/tables/casing_data.zig

    scripts/gen_unicode_tables.py --mode nfc \\
        --input /path/to/UnicodeData.txt \\
        --excl /path/to/CompositionExclusions.txt \\
        --output unicode/tables/nfc_data.zig

    scripts/gen_unicode_tables.py --mode xid \\
        --input /path/to/DerivedCoreProperties.txt \\
        --output unicode/tables/xid_data.zig

Each generated file pins the Unicode version + SHA-256 of every
input in its header so re-generation is reproducible.
`scripts/ci/check_unicode_tables.sh` is the regen gate: it fetches
the pinned inputs, verifies those digests, re-runs the generator and
fails on any diff.
"""
from __future__ import annotations

import argparse
import hashlib
import re
import sys
from pathlib import Path


ATTRIBUTION = [
    "//",
    "// This file contains data derived from the Unicode Character",
    "// Database, used under the Unicode License v3. See",
    "// THIRD_PARTY_NOTICES.md at the repository root for the full",
    "// license text and attribution.",
]

CASEFOLD_LINE = re.compile(
    r"^(?P<code>[0-9A-F]+);\s*"
    r"(?P<status>[CFST]);\s*"
    r"(?P<mapping>[0-9A-F ]+);\s*#"
)

VERSION_LINE = re.compile(r"^# CaseFolding-(?P<version>[0-9.]+)\.txt")


def parse_input(path: Path) -> tuple[str, list[tuple[int, list[int]]]]:
    """Parse CaseFolding.txt; return (unicode_version, [(from, [to])])
    keeping only statuses 'C' and 'F' (non-Turkic full case fold)."""
    version = ""
    entries: list[tuple[int, list[int]]] = []
    with path.open() as f:
        for line in f:
            if not version:
                m = VERSION_LINE.match(line)
                if m:
                    version = m.group("version")
            m = CASEFOLD_LINE.match(line)
            if not m:
                continue
            status = m.group("status")
            if status not in ("C", "F"):
                # Skip simple-only ('S') — full ('F') already covers
                # the same source codepoints with a richer mapping.
                # Skip Turkic ('T') per the non-Turkic policy choice.
                continue
            src = int(m.group("code"), 16)
            mapping = [int(c, 16) for c in m.group("mapping").split()]
            entries.append((src, mapping))
    if not version:
        raise SystemExit("gen_unicode_tables: input missing # CaseFolding-X.Y.Z.txt header")
    # Sort by source codepoint (already sorted in UCD but be defensive).
    entries.sort(key=lambda e: e[0])
    # Sanity: no duplicate sources.
    seen = set()
    for src, _ in entries:
        if src in seen:
            raise SystemExit(f"gen_unicode_tables: duplicate source codepoint U+{src:04X}")
        seen.add(src)
    return version, entries


def emit_zig(
    out_path: Path, version: str, entries: list[tuple[int, list[int]]], input_sha256: str
) -> None:
    """Render the generated Zig file."""
    # Build the scalar pool: concatenate all mappings, indexed by offset.
    pool: list[int] = []
    fold_records: list[tuple[int, int, int]] = []  # (from, len, offset)
    for src, mapping in entries:
        offset = len(pool)
        pool.extend(mapping)
        fold_records.append((src, len(mapping), offset))

    lines = [
        "// AUTO-GENERATED by scripts/gen_unicode_tables.py — DO NOT EDIT",
        "// Source: https://www.unicode.org/Public/UCD/latest/ucd/CaseFolding.txt",
        f"// SHA-256 of input: {input_sha256}",
        f"// Unicode version: {version}",
        "// Policy: non-Turkic full case fold (statuses C + F).",
        *ATTRIBUTION,
        "",
        f'pub const unicode_version: []const u8 = "{version}";',
        "",
        "pub const FoldEntry = struct {",
        "    from: u21,",
        "    len: u8,",
        "    offset: u32,",
        "};",
        "",
        "pub const fold_entries: []const FoldEntry = &.{",
    ]
    for src, length, offset in fold_records:
        lines.append(f"    .{{ .from = 0x{src:04X}, .len = {length}, .offset = {offset} }},")
    lines.append("};")
    lines.append("")
    lines.append("pub const fold_scalars: []const u21 = &.{")
    # Pack 8 codepoints per line for readability + smaller diff churn.
    for i in range(0, len(pool), 8):
        chunk = ", ".join(f"0x{c:04X}" for c in pool[i : i + 8])
        lines.append(f"    {chunk},")
    lines.append("};")
    lines.append("")

    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text("\n".join(lines))
    print(f"gen_unicode_tables: {len(entries)} entries, {len(pool)} pool scalars → {out_path}")


# ─── casing_v1 (M4f, §5.4b) ──────────────────────────────────────────

# The three casing columns of UnicodeData.txt, which the NFC regex above
# deliberately stops short of: fields 12, 13 and 14 (uppercase, lowercase
# and titlecase simple mappings) of a 15-field record.
UNICODE_DATA_FIELDS = 15
UD_FIELD_UPPER = 12
UD_FIELD_LOWER = 13
UD_FIELD_TITLE = 14

SPECIAL_CASING_VERSION_LINE = re.compile(r"^# SpecialCasing-(?P<version>[0-9.]+)\.txt")

# The one conditional row that carries no language tag. Every other
# condition in SpecialCasing.txt is `lt`, `tr` or `az`.
FINAL_SIGMA = "Final_Sigma"


def parse_simple_casing(path: Path) -> dict[str, dict[int, list[int]]]:
    """UnicodeData.txt's simple upper/lower/title mappings.

    Identity mappings are dropped: the runtime returns an unmapped
    codepoint unchanged, so a row saying so would only cost a binary
    search step. Range records (`<..., First>` / `<..., Last>`) carry no
    casing and need no expansion."""
    out: dict[str, dict[int, list[int]]] = {"upper": {}, "lower": {}, "title": {}}
    with path.open() as f:
        for line in f:
            fields = line.rstrip("\n").split(";")
            if len(fields) != UNICODE_DATA_FIELDS:
                continue
            cp = int(fields[0], 16)
            for key, idx in (
                ("upper", UD_FIELD_UPPER),
                ("lower", UD_FIELD_LOWER),
                ("title", UD_FIELD_TITLE),
            ):
                raw = fields[idx].strip()
                if not raw:
                    continue
                mapping = [int(c, 16) for c in raw.split()]
                if mapping == [cp]:
                    continue
                out[key][cp] = mapping
    return out


def parse_special_casing(
    path: Path,
) -> tuple[str, dict[str, dict[int, list[int]]], dict[str, list[int]], list[str]]:
    """SpecialCasing.txt, split by this generator's policy.

    Returns `(version, unconditional, final_sigma, rejected)` where
    `unconditional` has the same shape as `parse_simple_casing`'s result,
    `final_sigma` holds the one locale-independent conditional row, and
    `rejected` names every locale-conditional row that was dropped — the
    generator prints it, so "we do not do Turkish" stays a visible
    decision rather than a silent absence."""
    version = ""
    uncond: dict[str, dict[int, list[int]]] = {"upper": {}, "lower": {}, "title": {}}
    final_sigma: dict[str, list[int]] = {}
    rejected: list[str] = []

    with path.open() as f:
        for raw_line in f:
            if not version:
                m = SPECIAL_CASING_VERSION_LINE.match(raw_line)
                if m:
                    version = m.group("version")
            line = raw_line.split("#", 1)[0].strip()
            if not line:
                continue
            fields = [x.strip() for x in line.split(";")]
            # `<code>; <lower>; <title>; <upper>;` then an optional
            # condition list, then the trailing empty field the `;`
            # terminator leaves behind.
            if len(fields) < 5:
                continue
            cp = int(fields[0], 16)
            lower = [int(c, 16) for c in fields[1].split()]
            title = [int(c, 16) for c in fields[2].split()]
            upper = [int(c, 16) for c in fields[3].split()]
            condition = fields[4]

            if not condition:
                for key, mapping in (("upper", upper), ("lower", lower), ("title", title)):
                    if mapping == [cp]:
                        continue
                    uncond[key][cp] = mapping
                continue

            if condition == FINAL_SIGMA:
                if cp != 0x03A3:
                    raise SystemExit(
                        f"gen_unicode_tables: unexpected {FINAL_SIGMA} source U+{cp:04X}"
                    )
                final_sigma = {"source": [cp], "lower": lower}
                continue

            rejected.append(f"U+{cp:04X} ({condition})")

    if not version:
        raise SystemExit("gen_unicode_tables: input missing # SpecialCasing-X.Y.Z.txt header")
    if not final_sigma:
        raise SystemExit(f"gen_unicode_tables: {FINAL_SIGMA} row absent from SpecialCasing.txt")
    return version, uncond, final_sigma, rejected


def emit_casing_zig(
    out_path: Path,
    version: str,
    simple: dict[str, dict[int, list[int]]],
    special: dict[str, dict[int, list[int]]],
    final_sigma: dict[str, list[int]],
    props: dict[str, list[tuple[int, int]]],
    digests: dict[str, str],
    rejected: list[str],
) -> None:
    """Render `casing_data.zig`: three full-mapping tables, the
    Final_Sigma rule, and the two property tables that rule needs."""
    merged: dict[str, dict[int, list[int]]] = {}
    for key in ("upper", "lower", "title"):
        # SpecialCasing wins: its rows are the full mappings, and a full
        # mapping that agrees with the simple one is not in the file.
        combined = dict(simple[key])
        combined.update(special[key])
        merged[key] = combined

    lines = [
        "// AUTO-GENERATED by scripts/gen_unicode_tables.py — DO NOT EDIT",
        "// Source files (UCD):",
        "//   https://www.unicode.org/Public/UCD/latest/ucd/UnicodeData.txt",
        "//   https://www.unicode.org/Public/UCD/latest/ucd/SpecialCasing.txt",
        "//   https://www.unicode.org/Public/UCD/latest/ucd/DerivedCoreProperties.txt",
    ]
    for k, v in digests.items():
        lines.append(f"// SHA-256 of {k}: {v}")
    lines.extend([
        f"// Unicode version: {version}",
        "// Policy: `casing_v1` — locale-neutral FULL casing. UnicodeData",
        "// simple mappings, overridden by unconditional SpecialCasing rows,",
        "// plus the locale-independent Final_Sigma rule. Locale-conditional",
        f"// rows (tr/az/lt) are rejected: {len(rejected)} dropped.",
        *ATTRIBUTION,
        "",
        f'pub const unicode_version: []const u8 = "{version}";',
        "",
        "/// A full case mapping: `len` scalars at `offset` in the matching",
        "/// pool. Length-changing by construction — `ß` (one scalar) maps to",
        "/// `SS` (two), which is the whole reason this is not a `u21` pair.",
        "pub const CaseEntry = struct { from: u21, len: u8, offset: u32 };",
        "",
    ])

    for key in ("upper", "lower", "title"):
        pool: list[int] = []
        records: list[tuple[int, int, int]] = []
        for cp in sorted(merged[key]):
            mapping = merged[key][cp]
            records.append((cp, len(mapping), len(pool)))
            pool.extend(mapping)
        lines.append(f"pub const {key}_entries: []const CaseEntry = &.{{")
        for cp, length, offset in records:
            lines.append(f"    .{{ .from = 0x{cp:04X}, .len = {length}, .offset = {offset} }},")
        lines.append("};")
        lines.append("")
        lines.append(f"pub const {key}_scalars: []const u21 = &.{{")
        for i in range(0, len(pool), 8):
            chunk = ", ".join(f"0x{c:04X}" for c in pool[i : i + 8])
            lines.append(f"    {chunk},")
        lines.append("};")
        lines.append("")

    sigma_lower = final_sigma["lower"]
    if len(sigma_lower) != 1:
        raise SystemExit("gen_unicode_tables: Final_Sigma lower mapping is not a single scalar")
    lines.extend([
        "/// The one conditional SpecialCasing row that carries no language",
        "/// tag: GREEK CAPITAL LETTER SIGMA lowercases to FINAL SIGMA at the",
        "/// end of a word and to ordinary SMALL SIGMA everywhere else. The",
        "/// condition is decided at runtime over the two property tables",
        "/// below, because it is a property of the neighbours rather than of",
        "/// the codepoint.",
        f"pub const final_sigma_source: u21 = 0x{final_sigma['source'][0]:04X};",
        f"pub const final_sigma_lower: u21 = 0x{sigma_lower[0]:04X};",
        "",
        "/// `Cased` and `Case_Ignorable` (UAX #29 / DerivedCoreProperties),",
        "/// the two properties Final_Sigma is defined over. Same interval",
        "/// shape as `xid_data.zig`: sorted, coalesced, binary-searchable,",
        "/// allocation-free.",
        "pub const Range = struct { lo: u21, hi: u21 };",
        "",
    ])
    for prop, name in (("Cased", "cased"), ("Case_Ignorable", "case_ignorable")):
        lines.append(f"pub const {name}: []const Range = &.{{")
        for lo, hi in props[prop]:
            lines.append(f"    .{{ .lo = 0x{lo:04X}, .hi = 0x{hi:04X} }},")
        lines.append("};")
        lines.append("")

    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text("\n".join(lines))
    print(
        f"gen_unicode_tables(casing): "
        f"{len(merged['upper'])} upper, {len(merged['lower'])} lower, "
        f"{len(merged['title'])} title, "
        f"{len(props['Cased'])} cased + {len(props['Case_Ignorable'])} "
        f"case-ignorable ranges → {out_path}"
    )
    print(f"gen_unicode_tables(casing): rejected {len(rejected)} locale rows: {', '.join(rejected)}")


UNICODE_DATA_LINE = re.compile(
    r"^(?P<code>[0-9A-F]+);"
    r"(?P<name>[^;]*);"
    r"(?P<gc>[^;]*);"
    r"(?P<ccc>[0-9]+);"
    r"(?P<bidi>[^;]*);"
    r"(?P<decomp>[^;]*);"
)


def parse_unicode_data(path: Path) -> tuple[
    list[tuple[int, list[int]]],   # canonical decompositions
    dict[int, int],                # ccc map
]:
    """Extract canonical decomposition mappings + canonical combining
    class from `UnicodeData.txt`. Skips compatibility decompositions
    (those whose mapping starts with `<...>`)."""
    decomps: list[tuple[int, list[int]]] = []
    ccc: dict[int, int] = {}
    with path.open() as f:
        for line in f:
            m = UNICODE_DATA_LINE.match(line)
            if not m:
                continue
            cp = int(m.group("code"), 16)
            ccc_val = int(m.group("ccc"))
            if ccc_val != 0:
                ccc[cp] = ccc_val
            decomp_field = m.group("decomp").strip()
            if not decomp_field or decomp_field.startswith("<"):
                # Empty or compatibility — skip.
                continue
            mapping = [int(c, 16) for c in decomp_field.split()]
            decomps.append((cp, mapping))
    return decomps, ccc


def parse_composition_exclusions(path: Path) -> set[int]:
    """Read `CompositionExclusions.txt` — each line lists a codepoint
    whose canonical decomposition must NOT be re-composed by NFC."""
    out: set[int] = set()
    with path.open() as f:
        for line in f:
            line = line.split("#", 1)[0].strip()
            if not line:
                continue
            # Range or single?
            if ".." in line:
                lo, hi = line.split("..")
                for cp in range(int(lo, 16), int(hi, 16) + 1):
                    out.add(cp)
            else:
                out.add(int(line, 16))
    return out


def emit_nfc_zig(
    out_path: Path,
    version: str,
    decomps: list[tuple[int, list[int]]],
    ccc: dict[int, int],
    exclusions: set[int],
    digests: dict[str, str],
) -> None:
    """Render `nfc_data.zig` with three tables:
      - decomp_entries: canonical decomposition lookup (sorted by `from`)
      - ccc_entries: combining class lookup (sorted by `cp`)
      - compose_entries: pair → composed codepoint (sorted by `(starter, combining)`)
    """
    # Sort decompositions by source codepoint.
    decomps.sort(key=lambda e: e[0])
    decomp_pool: list[int] = []
    decomp_records: list[tuple[int, int, int]] = []  # (from, len, offset)
    for cp, mapping in decomps:
        offset = len(decomp_pool)
        decomp_pool.extend(mapping)
        decomp_records.append((cp, len(mapping), offset))

    # CCC entries sorted by codepoint.
    ccc_records = sorted(ccc.items())

    # Build composition pairs from canonical decompositions, except:
    #  - skip if codepoint is in CompositionExclusions;
    #  - skip non-pair decompositions (only `pair` decomps compose);
    #  - skip if first codepoint has non-zero CCC (must be a starter).
    compose_pairs: list[tuple[int, int, int]] = []  # (starter, combining, composed)
    for cp, mapping in decomps:
        if cp in exclusions:
            continue
        if len(mapping) != 2:
            continue
        starter, combining = mapping
        if ccc.get(starter, 0) != 0:
            continue
        compose_pairs.append((starter, combining, cp))
    compose_pairs.sort(key=lambda e: (e[0], e[1]))

    lines = [
        "// AUTO-GENERATED by scripts/gen_unicode_tables.py — DO NOT EDIT",
        "// Source files (UCD):",
        "//   https://www.unicode.org/Public/UCD/latest/ucd/UnicodeData.txt",
        "//   https://www.unicode.org/Public/UCD/latest/ucd/CompositionExclusions.txt",
    ]
    for k, v in digests.items():
        lines.append(f"// SHA-256 of {k}: {v}")
    lines.extend([
        f"// Unicode version: {version}",
        *ATTRIBUTION,
        "",
        f'pub const unicode_version: []const u8 = "{version}";',
        "",
        "/// Hangul algorithmic ranges — handled at runtime, not via tables.",
        "pub const hangul_syllable_base: u21 = 0xAC00;",
        "pub const hangul_syllable_count: u21 = 11172;",
        "pub const hangul_l_base: u21 = 0x1100;",
        "pub const hangul_v_base: u21 = 0x1161;",
        "pub const hangul_t_base: u21 = 0x11A7;",
        "pub const hangul_l_count: u21 = 19;",
        "pub const hangul_v_count: u21 = 21;",
        "pub const hangul_t_count: u21 = 28;",
        "",
        "pub const DecompEntry = struct { from: u21, len: u8, offset: u32 };",
        "pub const decomp_entries: []const DecompEntry = &.{",
    ])
    for cp, length, offset in decomp_records:
        lines.append(f"    .{{ .from = 0x{cp:04X}, .len = {length}, .offset = {offset} }},")
    lines.append("};")
    lines.append("")
    lines.append("pub const decomp_scalars: []const u21 = &.{")
    for i in range(0, len(decomp_pool), 8):
        chunk = ", ".join(f"0x{c:04X}" for c in decomp_pool[i : i + 8])
        lines.append(f"    {chunk},")
    lines.append("};")
    lines.append("")
    lines.append("pub const CccEntry = struct { cp: u21, ccc: u8 };")
    lines.append("pub const ccc_entries: []const CccEntry = &.{")
    for cp, c in ccc_records:
        lines.append(f"    .{{ .cp = 0x{cp:04X}, .ccc = {c} }},")
    lines.append("};")
    lines.append("")
    lines.append("pub const ComposeEntry = struct { starter: u21, combining: u21, composed: u21 };")
    lines.append("pub const compose_entries: []const ComposeEntry = &.{")
    for s, c, comp in compose_pairs:
        lines.append(f"    .{{ .starter = 0x{s:04X}, .combining = 0x{c:04X}, .composed = 0x{comp:04X} }},")
    lines.append("};")
    lines.append("")

    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text("\n".join(lines))
    print(
        f"gen_unicode_tables(nfc): {len(decomp_records)} decomp, "
        f"{len(ccc_records)} ccc, {len(compose_pairs)} compose pairs → {out_path}"
    )


DERIVED_PROP_LINE = re.compile(
    r"^(?P<lo>[0-9A-F]{4,6})(?:\.\.(?P<hi>[0-9A-F]{4,6}))?\s*;\s*(?P<prop>[A-Za-z_]+)"
)

DERIVED_VERSION_LINE = re.compile(r"^# DerivedCoreProperties-(?P<version>[0-9.]+)\.txt")


def parse_derived_core_properties(
    path: Path, wanted: tuple[str, ...]
) -> tuple[str, dict[str, list[tuple[int, int]]]]:
    """Extract the requested binary properties from
    `DerivedCoreProperties.txt` as sorted, merged, non-overlapping
    [lo, hi] intervals. Returns (unicode_version, {prop: intervals})."""
    version = ""
    raw: dict[str, list[tuple[int, int]]] = {p: [] for p in wanted}
    with path.open() as f:
        for line in f:
            if line.startswith("#"):
                if not version:
                    m = DERIVED_VERSION_LINE.match(line)
                    if m:
                        version = m.group("version")
                continue
            m = DERIVED_PROP_LINE.match(line)
            if not m:
                continue
            prop = m.group("prop")
            if prop not in raw:
                continue
            lo = int(m.group("lo"), 16)
            hi = int(m.group("hi"), 16) if m.group("hi") else lo
            raw[prop].append((lo, hi))
    if not version:
        raise SystemExit(
            "gen_unicode_tables: input missing # DerivedCoreProperties-X.Y.Z.txt header"
        )
    out: dict[str, list[tuple[int, int]]] = {}
    for prop, intervals in raw.items():
        if not intervals:
            raise SystemExit(f"gen_unicode_tables: property {prop} absent from input")
        out[prop] = merge_intervals(intervals)
    return version, out


def merge_intervals(intervals: list[tuple[int, int]]) -> list[tuple[int, int]]:
    """Sort and coalesce touching or overlapping ranges. Coalescing
    matters: the UCD lists XID_Start in general-category order, so
    adjacent categories produce ranges that abut (e.g. `0041..005A`
    and `005B` would stay two rows) and every merged pair is one
    fewer binary-search step at runtime."""
    merged: list[tuple[int, int]] = []
    for lo, hi in sorted(intervals):
        if merged and lo <= merged[-1][1] + 1:
            prev_lo, prev_hi = merged[-1]
            merged[-1] = (prev_lo, max(prev_hi, hi))
        else:
            merged.append((lo, hi))
    return merged


def emit_xid_zig(
    out_path: Path,
    version: str,
    props: dict[str, list[tuple[int, int]]],
    input_sha256: str,
) -> None:
    """Render `xid_data.zig` with two interval tables."""
    start = props["XID_Start"]
    cont = props["XID_Continue"]

    # XID_Start ⊆ XID_Continue is a UCD invariant. Assert it here so a
    # future UCD revision that broke it could not slip into the tree.
    for lo, hi in start:
        if not any(c_lo <= lo and hi <= c_hi for c_lo, c_hi in cont):
            raise SystemExit(
                f"gen_unicode_tables: XID_Start range U+{lo:04X}..U+{hi:04X} "
                "is not contained in XID_Continue"
            )

    lines = [
        "// AUTO-GENERATED by scripts/gen_unicode_tables.py — DO NOT EDIT",
        "// Source: https://www.unicode.org/Public/17.0.0/ucd/DerivedCoreProperties.txt",
        f"// SHA-256 of input: {input_sha256}",
        f"// Unicode version: {version}",
        "// Properties: XID_Start, XID_Continue (UAX #31 identifier syntax).",
        *ATTRIBUTION,
        "",
        f'pub const unicode_version: []const u8 = "{version}";',
        "",
        "/// An inclusive codepoint interval. Tables are sorted by `lo`,",
        "/// non-overlapping, and coalesced, so a binary search over `lo`",
        "/// answers membership in one probe chain with no allocation.",
        "pub const Range = struct { lo: u21, hi: u21 };",
        "",
        "pub const xid_start: []const Range = &.{",
    ]
    for lo, hi in start:
        lines.append(f"    .{{ .lo = 0x{lo:04X}, .hi = 0x{hi:04X} }},")
    lines.append("};")
    lines.append("")
    lines.append("pub const xid_continue: []const Range = &.{")
    for lo, hi in cont:
        lines.append(f"    .{{ .lo = 0x{lo:04X}, .hi = 0x{hi:04X} }},")
    lines.append("};")
    lines.append("")

    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text("\n".join(lines))
    print(
        f"gen_unicode_tables(xid): {len(start)} XID_Start ranges, "
        f"{len(cont)} XID_Continue ranges → {out_path}"
    )


def parse_unicode_version_from_data(path: Path) -> str:
    """UnicodeData.txt has no version header. Use the path's
    sibling Derived files or just hard-code current. Better: read
    UCD's ReadMe.txt, but for now derive from CaseFolding.txt's
    version pin since the generator is run alongside it."""
    return "17.0.0"


def main() -> int:
    p = argparse.ArgumentParser()
    p.add_argument("--mode", choices=["casefold", "casing", "nfc", "xid"], default="casefold")
    p.add_argument("--input", type=Path, required=True)
    p.add_argument("--excl", type=Path, default=None, help="CompositionExclusions.txt (NFC only)")
    p.add_argument("--special", type=Path, default=None, help="SpecialCasing.txt (casing only)")
    p.add_argument(
        "--props", type=Path, default=None, help="DerivedCoreProperties.txt (casing only)"
    )
    p.add_argument("--output", type=Path, required=True)
    args = p.parse_args()

    if args.mode == "casing":
        if args.special is None or args.props is None:
            print(
                "gen_unicode_tables: --mode casing requires --special and --props",
                file=sys.stderr,
            )
            return 2
        simple = parse_simple_casing(args.input)
        version, special, final_sigma, rejected = parse_special_casing(args.special)
        prop_version, props = parse_derived_core_properties(
            args.props, ("Cased", "Case_Ignorable")
        )
        # Three files, one Unicode revision. A casing table built from
        # 17.0.0 mappings and 16.0.0 properties would decide Final_Sigma
        # by a different alphabet than it cases.
        if prop_version != version:
            raise SystemExit(
                "gen_unicode_tables: SpecialCasing is "
                f"{version} but DerivedCoreProperties is {prop_version}"
            )
        digests = {
            args.input.name: hashlib.sha256(args.input.read_bytes()).hexdigest(),
            args.special.name: hashlib.sha256(args.special.read_bytes()).hexdigest(),
            args.props.name: hashlib.sha256(args.props.read_bytes()).hexdigest(),
        }
        emit_casing_zig(
            args.output, version, simple, special, final_sigma, props, digests, rejected
        )
        return 0

    if args.mode == "casefold":
        raw = args.input.read_bytes()
        digest = hashlib.sha256(raw).hexdigest()
        version, entries = parse_input(args.input)
        emit_zig(args.output, version, entries, digest)
        return 0

    if args.mode == "xid":
        digest = hashlib.sha256(args.input.read_bytes()).hexdigest()
        version, props = parse_derived_core_properties(
            args.input, ("XID_Start", "XID_Continue")
        )
        emit_xid_zig(args.output, version, props, digest)
        return 0

    # NFC mode.
    if args.excl is None:
        print("gen_unicode_tables: --mode nfc requires --excl", file=sys.stderr)
        return 2
    decomps, ccc = parse_unicode_data(args.input)
    exclusions = parse_composition_exclusions(args.excl)
    version = parse_unicode_version_from_data(args.input)
    digests = {
        args.input.name: hashlib.sha256(args.input.read_bytes()).hexdigest(),
        args.excl.name: hashlib.sha256(args.excl.read_bytes()).hexdigest(),
    }
    emit_nfc_zig(args.output, version, decomps, ccc, exclusions, digests)
    return 0


if __name__ == "__main__":
    sys.exit(main())
