#!/usr/bin/env python3
"""Lint docs/plans/surface-matrix.md against the code it describes.

Every back-ticked symbol inside a per-surface cell must exist on that
surface, *on the owner the cell names*, so a cell cannot claim an entry
point that was renamed, moved or never shipped:

  Zig  — `Type.member` resolves to a `pub fn`, `pub const` or field inside
         the top-level `pub const Type = struct/union/enum { … }` block of
         some file under src/, pkg/, recalc/. `zlsx.fn` resolves to a
         module-level `pub fn`/`pub const` in src/xlsx.zig; `zlsx_pkg.fn` to an
         export in pkg/root.zig; `zlsx_recalc.fn` to one in recalc/recalc.zig.
         A ⛔ cell names a typed error; it must be spelled `error.Name`
         somewhere in the Zig tree. `zlsx.fn` resolves against src/xlsx.zig
         only — src/writer.zig is a private module, so its members are named
         `Writer.x` / `SheetWriter.x` and the matrix says so in §2.
  C    — `zlsx_*` names must be declared as functions in include/zlsx.h.
  Py   — `Class.method` resolves to a def inside that class (or a
         `Class.method = …` assignment) in bindings/python/zlsx/__init__.py;
         `zlsx.name` to a module-level def, class, `self.`-attribute owner, sub-package
         or sibling module.
  CLI  — a sub-command must appear as a string literal in the non-test part
         of src/cli.zig or src/formula_cli.zig; a `--flag` likewise;
         `zlsx-<name>` must have a src/<name>_main.zig.

Structure: every ✓ / ~ / ⛔ cell names at least one symbol; every `—` names
`— Sx` or `— gate`; row ids exist in goal_sigmoid.md's ladder table or are
marked *(proposed)* / *(scope change)* in this file's "Gaps by sigmoid row";
`n/a` needs the Rulings table to be out of its PROPOSED state; capability
tables keep their column count.

Known residual: a sub-command check passes for any literal compared with
`std.mem.eql` in the dispatch files, so a `--format` value such as `jsonl`
would pass as a sub-command; only a grammar-level parse would close that.

Exit 0 when clean; exit 1 with one line per problem otherwise.
Run from the repo root: python3 scripts/check_surface_matrix.py
"""

from __future__ import annotations

import ast
import re
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
MATRIX = REPO / "docs" / "plans" / "surface-matrix.md"
LADDER = REPO / "goal_sigmoid.md"
HEADER = REPO / "include" / "zlsx.h"
PY_PKG = REPO / "bindings" / "python" / "zlsx"
PY_INIT = PY_PKG / "__init__.py"
CLI_FILES = [REPO / "src" / "cli.zig", REPO / "src" / "formula_cli.zig"]
ZIG_ROOTS = [REPO / "src", REPO / "pkg", REPO / "recalc"]
ZIG_MODULE_ROOTS = {
    "zlsx": [REPO / "src" / "xlsx.zig"],  # the `zlsx` addModule root; src/writer.zig is a private module
    "zlsx_pkg": [REPO / "pkg" / "root.zig"],
    "zlsx_recalc": [REPO / "recalc" / "recalc.zig"],
}

SURFACES = ("Zig", "C", "Py", "CLI")
TICK = re.compile(r"`([^`]+)`")
ROW_REF = re.compile(r"—\s*(S\d+[a-z]?|gate)\b")
ROW_ID = re.compile(r"^S\d+[a-z]?$")
TOP_TYPE = re.compile(r"^pub const ([A-Z][A-Za-z0-9_]*) = (?:extern |packed )?(?:struct|union|enum)\b[^{]*\{")
TOP_PUB = re.compile(r"^pub (?:inline )?(?:fn|const) ([A-Za-z_][A-Za-z0-9_]*)")
INNER_PUB = re.compile(r"^\s+pub (?:inline )?(?:fn|const|var) ([A-Za-z_][A-Za-z0-9_]*)")
SUPERSCRIPT = "⁰¹²³⁴⁵⁶⁷⁸⁹"
FIELD = re.compile(r"^    ([a-z_][A-Za-z0-9_]*)\s*:")
ERR_TAG = re.compile(r"error\.([A-Z][A-Za-z0-9_]*)")
CLI_SUBCOMMAND = "eql\\(u8,\\s*[^,\"]+,\\s*\"{}\"\\)"


# ── Zig ─────────────────────────────────────────────────────────────────

def _strip_zig_line(line: str) -> str:
    """Drop comments and string contents so braces inside them don't count."""
    line = re.sub(r'"(?:\\.|[^"\\])*"', '""', line)
    line = re.sub(r"'(?:\\.|[^'\\])'", "''", line)
    return line.split("//", 1)[0]


def zig_index() -> tuple[dict[str, set[str]], dict[str, set[str]], set[str]]:
    """(members by (file, top-level type), module-level pubs by file, error tags)."""
    members: dict[tuple[str, str], set[str]] = {}
    top_level: dict[str, set[str]] = {}
    errors: set[str] = set()
    for root in ZIG_ROOTS:
        for path in root.rglob("*.zig"):
            text = path.read_text(encoding="utf-8", errors="replace")
            errors.update(ERR_TAG.findall("\n".join(_strip_zig_line(l) for l in text.splitlines())))
            tops = top_level.setdefault(str(path), set())
            current: str | None = None
            depth = 0
            for raw in text.splitlines():
                if current is None:
                    m = TOP_TYPE.match(raw)
                    if m:
                        current = m.group(1)
                        members.setdefault((str(path), current), set())
                        tops.add(current)
                        stripped = _strip_zig_line(raw)
                        depth = stripped.count("{") - stripped.count("}")
                        if depth <= 0:
                            current = None
                        continue
                    m = TOP_PUB.match(raw)
                    if m:
                        tops.add(m.group(1))
                    continue
                # Inside a top-level type block: only depth-1 declarations belong to it.
                if depth == 1:
                    m = INNER_PUB.match(raw)
                    if m:
                        members[(str(path), current)].add(m.group(1))
                    else:
                        f = FIELD.match(raw)
                        if f:
                            members[(str(path), current)].add(f.group(1))
                stripped = _strip_zig_line(raw)
                depth += stripped.count("{") - stripped.count("}")
                if depth <= 0:
                    current = None
    return members, top_level, errors


def zig_module_pubs(top_level: dict[str, set[str]], module: str) -> set[str]:
    out: set[str] = set()
    for root in ZIG_MODULE_ROOTS[module]:
        if root.is_file():
            out |= top_level.get(str(root), set())
        else:
            for path, names in top_level.items():
                if Path(path).parent == root:  # module level = files directly under src/
                    out |= names
    return out


# ── C ───────────────────────────────────────────────────────────────────

def c_symbols() -> set[str]:
    return set(re.findall(r"\b(zlsx_[a-z0-9_]+)\s*\(", HEADER.read_text(encoding="utf-8")))


# ── Python ──────────────────────────────────────────────────────────────

def py_index() -> tuple[dict[str, set[str]], set[str]]:
    tree = ast.parse(PY_INIT.read_text(encoding="utf-8"))
    classes: dict[str, set[str]] = {}
    module: set[str] = set()
    for node in tree.body:
        if isinstance(node, ast.ClassDef):
            names = {n.name for n in node.body if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef))}
            # Instance attributes assigned as `self.x = …` anywhere in the class body.
            for sub in ast.walk(node):
                if isinstance(sub, (ast.Assign, ast.AnnAssign)):
                    targets = sub.targets if isinstance(sub, ast.Assign) else [sub.target]
                    for t in targets:
                        if isinstance(t, ast.Attribute) and isinstance(t.value, ast.Name) and t.value.id == "self":
                            names.add(t.attr)
            classes[node.name] = names
            module.add(node.name)
        elif isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
            module.add(node.name)
        elif isinstance(node, ast.Assign):
            for tgt in node.targets:
                if isinstance(tgt, ast.Attribute) and isinstance(tgt.value, ast.Name):
                    classes.setdefault(tgt.value.id, set()).add(tgt.attr)  # SheetWriter.x = _sheet_x
                elif isinstance(tgt, ast.Name):
                    module.add(tgt.id)
    for sub in PY_PKG.iterdir():
        if (sub.is_dir() and (sub / "__init__.py").exists()) or (sub.suffix == ".py" and sub.name != "__init__.py"):
            module.add(sub.stem)  # sub-packages and sibling modules reachable as zlsx.<name>
    return classes, module


# ── CLI ─────────────────────────────────────────────────────────────────

def cli_text() -> str:
    """Source before the first `test "` block, per file — literals in tests don't count."""
    parts = []
    for p in CLI_FILES:
        text = p.read_text(encoding="utf-8")
        parts.append(re.split(r'^test "', text, maxsplit=1, flags=re.M)[0])
    return "\n".join(parts)


# ── Ladder ──────────────────────────────────────────────────────────────

def ladder_rows() -> set[str]:
    return set(re.findall(r"^\| (S\d+[a-z]?) \|", LADDER.read_text(encoding="utf-8"), re.M))


def proposed_rows(matrix: str) -> set[str]:
    return set(re.findall(r"^\| (S\d+[a-z]?) \*\((?:proposed|scope change)\)\*", matrix, re.M))


# ── Cells ───────────────────────────────────────────────────────────────

def check_symbol(surface: str, sym: str, ctx: dict) -> str | None:
    if surface == "C":
        return None if sym in ctx["c"] else f"C: `{sym}` is not declared in include/zlsx.h"
    if surface == "Py":
        if "/" in sym:
            return None if (REPO / sym).exists() else f"Py: path `{sym}` does not exist"
        owner, _, leaf = sym.rpartition(".")
        if owner in ("", "zlsx"):
            return None if leaf in ctx["py_module"] else f"Py: `{sym}` is not module-level in zlsx/__init__.py"
        if leaf in ctx["py_classes"].get(owner, set()):
            return None
        return f"Py: `{sym}` — class `{owner}` has no `{leaf}` in zlsx/__init__.py"
    if surface == "CLI":
        if sym.startswith("--"):
            return None if f'"{sym}"' in ctx["cli"] else f"CLI: flag `{sym}` not parsed in src/cli.zig / formula_cli.zig"
        if sym.startswith("zlsx-"):
            main = REPO / "src" / (sym[len("zlsx-"):].replace("-", "_") + "_main.zig")
            return None if main.exists() else f"CLI: sibling binary `{sym}` has no {main.name}"
        if re.search(CLI_SUBCOMMAND.format(re.escape(sym)), ctx["cli"]):
            return None
        return f"CLI: sub-command `{sym}` is not compared with std.mem.eql in src/cli.zig / formula_cli.zig"
    # Zig
    owner, _, leaf = sym.rpartition(".")
    if owner in ZIG_MODULE_ROOTS:
        pubs = zig_module_pubs(ctx["zig_top"], owner)
        return None if leaf in pubs else f"Zig: `{sym}` — `{owner}` has no module-level pub `{leaf}`"
    if owner == "":
        return None if any(leaf in s for s in ctx["zig_top"].values()) else f"Zig: `{sym}` is not a module-level pub"
    blocks = [names for (_, typ), names in ctx["zig_members"].items() if typ == owner]
    if not blocks:
        return f"Zig: `{sym}` — no top-level type `{owner}` under src/, pkg/, recalc/"
    if any(leaf in names for names in blocks):
        return None
    return f"Zig: `{sym}` — no `{owner}` block declares pub member `{leaf}`"


def check_cell(surface: str, cell: str, ctx: dict) -> list[str]:
    syms = TICK.findall(cell)
    mark = cell[:1]
    if mark == "⛔":
        if not syms:
            return [f"{surface}: ⛔ cell names no error"]
        return [f"{surface}: refusal `{s}` is not spelled `error.{s}` in the Zig tree"
                for s in syms if s not in ctx["zig_errors"]]
    if mark in ("✓", "~"):
        if not syms:
            return [f"{surface}: `{mark}` cell names no symbol"]
        residue = TICK.sub("", cell[1:])
        residue = re.sub(r"\([^)]*\)", "", residue)  # one parenthetical qualifier allowed
        residue = residue.translate({ord(c): None for c in SUPERSCRIPT + ", "})
        if residue:
            return [f"{surface}: prose in cell — `{residue.strip()}` (symbols, commas, footnote marks and one (…) only)"]
        return [p for s in syms if (p := check_symbol(surface, s, ctx))]
    if mark == "—":
        m = ROW_REF.search(cell)
        if not m:
            return [f"{surface}: `—` names no row (`— Sx` or `— gate`)"]
        if re.sub(r"[⁰¹²³⁴⁵⁶⁷⁸⁹ ]", "", cell[m.end():]):
            return [f"{surface}: prose after `— {m.group(1)}`"]
        row = m.group(1)
        if row != "gate" and row not in ctx["rows"]:
            return [f"{surface}: `— {row}` names a row that is neither in goal_sigmoid.md nor marked *(proposed)*"]
        return []
    if cell.startswith("n/a"):
        return [] if ctx["rulings_final"] else [f"{surface}: `n/a` while Rulings are still PROPOSED"]
    return [f"{surface}: cell `{cell[:30]}` does not start with ✓ ~ ⛔ — or n/a"]


BUILD = REPO / "build.zig"


def fuzz_row_problems(matrix: str) -> list[str]:
    """Every `.fuzz = true` root in build.zig must appear in the
    coverage-guided-fuzz row of the cross-cutting table, so a fuzz
    binary cannot be added (or renamed) without the inventory noticing
    (MNT-2402)."""
    src = "\n".join(
        line.split("//", 1)[0]
        for line in BUILD.read_text(encoding="utf-8").splitlines()
    )
    lines = src.splitlines()
    roots: set[str] = set()
    for i, line in enumerate(lines):
        m = re.search(r'\.root_source_file = b\.path\("([^"]+)"\)', line)
        if not m:
            continue
        for look in lines[i + 1 : i + 8]:
            if "});" in look:
                break
            if ".fuzz = true" in look:
                roots.add(m.group(1))
                break
    # The walker loop's roots come from its array literal
    # (`b.path(w.path)` carries no string to match on).
    walker = re.search(r"walker_fuzz = \[_\][^;]*?\};", src, re.S)
    if walker:
        roots.update(re.findall(r'\.path = "([^"]+)"', walker.group(0)))
    row = next(
        (l for l in matrix.splitlines() if l.startswith("| Coverage-guided fuzz binaries")),
        None,
    )
    if row is None:
        return ["surface-matrix: the coverage-guided fuzz row is missing"]
    expanded = row
    for m in re.finditer(r"([^\s`|{]*)\{([^}]*)\}([^\s`|]*)", row):
        expanded += " " + " ".join(
            m.group(1) + part + m.group(3) for part in m.group(2).split(",")
        )
    return [
        f"surface-matrix: fuzz row omits `.fuzz = true` root {r} from build.zig"
        for r in sorted(roots)
        if r not in expanded
    ]


def main() -> int:
    matrix = MATRIX.read_text(encoding="utf-8")
    members, top_level, errors = zig_index()
    classes, module = py_index()
    ctx = {
        "zig_members": members, "zig_top": top_level, "zig_errors": errors,
        "c": c_symbols(), "py_classes": classes, "py_module": module,
        "cli": cli_text(),
        "rows": ladder_rows() | proposed_rows(matrix),
        "rulings_final": "PROPOSED at the S0 gate" not in matrix,
    }
    problems: list[str] = []
    section = ""
    for lineno, line in enumerate(matrix.splitlines(), 1):
        if line.startswith("## "):
            section = line[3:4]
            continue
        if not line.startswith("|") or not section or section not in "1234567":
            continue
        cells = [c.strip() for c in line.strip().strip("|").split("|")]
        if cells[0] in ("Capability", "Property") or set(cells[1:]) <= {"---"}:
            continue
        if section == "6":
            continue  # prose table by design
        want = 5 if section == "7" else 6
        if len(cells) != want:
            problems.append(f"{MATRIX.name}:{lineno}: expected {want} columns in §{section}, got {len(cells)}")
            continue
        for surface, cell in zip(SURFACES, cells[1:5]):
            problems += [f"{MATRIX.name}:{lineno}: {p}" for p in check_cell(surface, cell, ctx)]
        if want == 6 and cells[5]:
            for tok in cells[5].replace("·", " ").split():
                if not ROW_ID.match(tok) or tok not in ctx["rows"]:
                    problems.append(f"{MATRIX.name}:{lineno}: Row column names unknown row `{tok}`")

    problems += fuzz_row_problems(matrix)

    if problems:
        print("\n".join(problems))
        print(f"\n{len(problems)} problem(s)")
        return 1
    print("surface-matrix: every symbol resolves on its owner")
    return 0


if __name__ == "__main__":
    sys.exit(main())
