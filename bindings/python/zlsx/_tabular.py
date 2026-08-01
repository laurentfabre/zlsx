"""Pure tabular logic behind :mod:`zlsx.spark` — no pyspark imports.

Everything here is driver/executor-agnostic and unit-testable without a
Spark session: the type-widening lattice, schema inference over a row
sample, permissive/failfast cell coercion, input-path expansion, and
partition planning. :mod:`zlsx.spark` is the thin PySpark shell over it.

Kinds are Spark DDL type names so inference can be rendered straight
into a DDL schema string: ``boolean`` / ``bigint`` / ``double`` /
``string``, plus ``void`` for a column with no observed values and the
schema-only kinds ``date`` / ``timestamp`` (never inferred — xlsx stores
dates as styled numbers — but accepted from a user-supplied schema).
"""
from __future__ import annotations

import datetime as _dt
import glob as _glob
import os

_EXCEL_EPOCH = _dt.datetime(1899, 12, 30)

# Widening lattice, applied per column across the sampled rows:
#
#   void < boolean < string          void < bigint < double < string
#
# boolean never widens into a numeric — a column mixing bools and
# numbers means the numbers were data and the bools were noise (or vice
# versa), and either way only string represents both losslessly.
_NUMERIC = {"bigint", "double"}
KINDS = ("void", "boolean", "bigint", "double", "string", "date", "timestamp")


def kind_of(value) -> str:
    """The narrowest kind holding this Python cell value."""
    if value is None:
        return "void"
    if isinstance(value, bool):  # before int: bool is an int subclass
        return "boolean"
    if isinstance(value, int):
        return "bigint"
    if isinstance(value, float):
        return "double"
    return "string"


def merge_kinds(a: str, b: str) -> str:
    if a == b:
        return a
    if a == "void":
        return b
    if b == "void":
        return a
    if a in _NUMERIC and b in _NUMERIC:
        return "double"
    return "string"


def infer_fields(header, sample_rows):
    """Infer ``[(name, kind, nullable)]`` from a header and a row sample.

    Widens across EVERY sampled row, not just the first — an int column
    that turns float on row 40 infers ``double``; one that turns textual
    infers ``string``. A column that never held a value infers
    ``string`` (the CSV-source convention for "no evidence").
    Short rows mark the missing columns nullable.
    """
    ncols = len(header) if header is not None else max(
        (len(r) for r in sample_rows), default=0
    )
    if ncols == 0:
        raise ValueError("cannot infer a schema from an empty sheet")

    names = _column_names(header, ncols)
    kinds = ["void"] * ncols
    nullable = [False] * ncols
    for row in sample_rows:
        for c in range(ncols):
            v = row[c] if c < len(row) else None
            if v is None:
                nullable[c] = True
            else:
                kinds[c] = merge_kinds(kinds[c], kind_of(v))
    kinds = [k if k != "void" else "string" for k in kinds]
    return list(zip(names, kinds, nullable))


def _column_names(header, ncols: int) -> list[str]:
    """Spark-safe column names: header cells stringified, blanks become
    ``_cN``, duplicates deduped with a numeric suffix."""
    raw = []
    for c in range(ncols):
        h = header[c] if header is not None and c < len(header) else None
        raw.append(str(h) if h is not None and str(h) != "" else f"_c{c}")
    seen: dict[str, int] = {}
    names = []
    for name in raw:
        n = seen.get(name, 0)
        seen[name] = n + 1
        names.append(name if n == 0 else f"{name}_{n}")
    return names


def ddl(fields) -> str:
    """Render inferred fields as a Spark DDL schema string."""
    return ", ".join(
        "`{}` {}".format(name.replace("`", "``"), kind)
        for name, kind, _ in fields
    )


def serial_to_datetime(serial: float) -> _dt.datetime:
    """Excel serial-date number → naive datetime (1900 date system)."""
    return _EXCEL_EPOCH + _dt.timedelta(days=float(serial))


class CellCoercionError(ValueError):
    """A cell that cannot be represented in the schema, under failfast."""


def coerce_row(row, kinds, mode: str, ctx: str, row_idx: int):
    """Fit one sheet row to the schema's column kinds.

    Short rows pad with ``None`` (ragged trailing cells are normal in
    real xlsx). Long rows and unrepresentable cells follow ``mode``:
    ``permissive`` truncates / nulls (the Spark CSV convention),
    ``failfast`` raises :class:`CellCoercionError` naming the exact
    file, sheet, row and column.
    """
    ncols = len(kinds)
    if len(row) > ncols:
        if mode == "failfast":
            raise CellCoercionError(
                f"{ctx} row {row_idx}: {len(row)} cells but the schema has "
                f"{ncols} columns"
            )
        row = row[:ncols]
    out = [None] * ncols
    for c in range(ncols):
        v = row[c] if c < len(row) else None
        if v is None:
            continue
        coerced = _coerce_cell(v, kinds[c])
        if coerced is _MISMATCH:
            if mode == "failfast":
                raise CellCoercionError(
                    f"{ctx} row {row_idx} col {c}: {v!r} does not fit "
                    f"schema type {kinds[c]}"
                )
            coerced = None
        out[c] = coerced
    return out


_MISMATCH = object()


def _coerce_cell(v, kind: str):
    if kind == "string":
        if isinstance(v, str):
            return v
        if isinstance(v, bool):
            return "true" if v else "false"
        return str(v)  # int / float — inference widened a mixed column
    if kind == "boolean":
        return v if isinstance(v, bool) else _MISMATCH
    if kind == "bigint":
        return v if isinstance(v, int) and not isinstance(v, bool) else _MISMATCH
    if kind == "double":
        if isinstance(v, bool):
            return _MISMATCH
        if isinstance(v, (int, float)):
            return float(v)
        return _MISMATCH
    if kind == "timestamp":
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            return serial_to_datetime(v)
        return _MISMATCH
    if kind == "date":
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            return serial_to_datetime(v).date()
        return _MISMATCH
    return _MISMATCH


def expand_paths(path_option: str) -> list[str]:
    """Expand a path option into a sorted list of workbook files.

    Accepts a comma-separated list; each entry may be a file, a
    directory (all ``*.xlsx`` / ``*.xlsm`` directly inside it), or a
    glob pattern. Raises when an entry matches nothing — an empty read
    should be a loud mistake, not an empty DataFrame.
    """
    files: list[str] = []
    for entry in (e.strip() for e in path_option.split(",")):
        if not entry:
            continue
        if os.path.isdir(entry):
            hits = sorted(
                os.path.join(entry, f)
                for f in os.listdir(entry)
                if f.endswith((".xlsx", ".xlsm")) and not f.startswith("~$")
            )
            if not hits:
                raise FileNotFoundError(f"no .xlsx/.xlsm files in directory: {entry}")
        elif _glob.has_magic(entry):
            hits = sorted(_glob.glob(entry))
            if not hits:
                raise FileNotFoundError(f"glob matched no files: {entry}")
        else:
            if not os.path.isfile(entry):
                raise FileNotFoundError(f"no such workbook: {entry}")
            hits = [entry]
        files.extend(hits)
    if not files:
        raise ValueError("path option resolved to no input files")
    # Dedup, preserving first-seen order.
    return list(dict.fromkeys(files))


def resolve_sheets(sheet_option: str, sheet_names) -> list[int]:
    """Resolve a sheet option against one workbook's sheet names.

    ``"*"`` selects every sheet; otherwise a comma-separated list of
    names and/or 0-based indices. Unknown names raise with the
    available names in the message.
    """
    if sheet_option.strip() == "*":
        return list(range(len(sheet_names)))
    out = []
    for entry in (e.strip() for e in sheet_option.split(",")):
        if not entry:
            continue
        if entry.isdigit():
            idx = int(entry)
            if idx >= len(sheet_names):
                raise IndexError(
                    f"sheet index {idx} out of range: workbook has "
                    f"{len(sheet_names)} sheet(s)"
                )
        else:
            try:
                idx = list(sheet_names).index(entry)
            except ValueError:
                raise KeyError(
                    f"no sheet named {entry!r}; available: {list(sheet_names)}"
                ) from None
        if idx not in out:
            out.append(idx)
    if not out:
        raise ValueError("sheet option resolved to no sheets")
    return out


def snapshot_files(path_option: str) -> dict:
    """Fingerprint every workbook the path option currently matches:
    ``{path: [mtime_ns, size]}``.

    Streaming variant of :func:`expand_paths`: an empty directory or an
    unmatched glob is a NORMAL state for a landing zone (nothing has
    arrived yet), so both yield ``{}`` instead of raising. A literal
    file path that does not exist still raises — naming one concrete
    file that is absent is a configuration error, not an empty zone.
    """
    files: dict = {}
    for entry in (e.strip() for e in path_option.split(",")):
        if not entry:
            continue
        if os.path.isdir(entry):
            hits = [
                os.path.join(entry, f)
                for f in os.listdir(entry)
                if f.endswith((".xlsx", ".xlsm")) and not f.startswith("~$")
            ]
        elif _glob.has_magic(entry):
            hits = _glob.glob(entry)
        else:
            if not os.path.isfile(entry):
                raise FileNotFoundError(f"no such workbook: {entry}")
            hits = [entry]
        for path in hits:
            try:
                st = os.stat(path)
            except FileNotFoundError:
                continue  # raced with a delete between list and stat
            files[path] = [st.st_mtime_ns, st.st_size]
    return files


def diff_snapshots(old: dict, new: dict) -> list:
    """Paths that are new or changed in ``new`` relative to ``old``,
    sorted for deterministic partition planning. Deletions are ignored:
    a removed workbook has nothing left to ingest. A changed
    fingerprint re-ingests the WHOLE file — the immutable-files
    convention (same as Auto Loader); dedupe downstream if producers
    rewrite in place."""
    return sorted(p for p, fp in new.items() if old.get(p) != fp)


def plan_row_ranges(n_data_rows: int, rows_per_partition: int):
    """Split ``n_data_rows`` into ``[(start, end)]`` ranges in data-row
    space (header excluded). One ``(0, n)`` range when splitting is off
    or unnecessary."""
    if rows_per_partition <= 0 or n_data_rows <= rows_per_partition:
        return [(0, n_data_rows)]
    return [
        (start, min(start + rows_per_partition, n_data_rows))
        for start in range(0, n_data_rows, rows_per_partition)
    ]
