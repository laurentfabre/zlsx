"""Pure tabular logic behind :mod:`zlsx.spark` — no pyspark imports.

Everything here is driver/executor-agnostic and unit-testable without a
Spark session: the type-widening lattice, schema inference over a row
sample (single- and multi-source), permissive/failfast cell coercion,
input-path expansion, partition planning, and atomic part-file landing.
:mod:`zlsx.spark` is the thin PySpark shell over it.

The split is load-bearing, not cosmetic: no CI job installs pyspark, so
anything living in ``spark.py`` ships untested. Logic goes here.

Kinds are Spark DDL type names so inference can be rendered straight
into a DDL schema string: ``boolean`` / ``bigint`` / ``double`` /
``string``, plus ``void`` for a column with no observed values and the
schema-only kinds ``date`` / ``timestamp`` (never inferred — xlsx stores
dates as styled numbers — but accepted from a user-supplied schema).
"""
from __future__ import annotations

import datetime as _dt
import glob as _glob
import hashlib as _hashlib
import os
import time as _time
import uuid as _uuid
from itertools import islice as _islice

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


def infer_fields(header, sample_rows, *, collapse_void: bool = True):
    """Infer ``[(name, kind, nullable)]`` from a header and a row sample.

    Widens across EVERY sampled row, not just the first — an int column
    that turns float on row 40 infers ``double``; one that turns textual
    infers ``string``. A column that never held a value infers
    ``string`` (the CSV-source convention for "no evidence").
    Short rows mark the missing columns nullable.

    ``collapse_void=False`` leaves a no-evidence column as ``"void"``
    instead, so callers merging several sources can keep "no evidence
    here" distinct from "genuinely textual" — an all-blank column in the
    first workbook would otherwise pin the merged column to ``string``
    and mask real types found in the others. Collapse once, after
    merging, via :func:`collapse_void_fields`.
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
    if collapse_void:
        kinds = [k if k != "void" else "string" for k in kinds]
    return list(zip(names, kinds, nullable))


def collapse_void_fields(fields):
    """Resolve any surviving ``"void"`` kind to ``string`` — the final
    step for a field list built with ``infer_fields(collapse_void=False)``
    and merged across sources."""
    return [(n, "string" if k == "void" else k, nl) for n, k, nl in fields]


def infer_schema(
    paths,
    *,
    sheet_option: str = "0",
    header: bool = True,
    infer_rows: int = 1000,
    infer_files: int = 10,
    open_book=None,
) -> str:
    """Infer one DDL schema across several workbooks and sheets.

    Samples ``infer_rows`` data rows from every selected sheet of the
    first ``infer_files`` workbooks (``0`` = all of them) and widens the
    result across all of them, so a column that is integral in one file
    and fractional in another resolves to ``double`` here rather than
    nulling (permissive) or raising (failfast) at read time.

    Empty sheets contribute nothing instead of aborting — one blank
    workbook in a landing zone should not fail the stream. Only when NO
    sampled sheet yields a header does this raise.

    ``open_book`` (path → Book context manager) defaults to
    ``zlsx.open``; the §12.4 recalc path passes one that opens the
    file's recalced in-memory snapshot instead, so inference sees the
    values the read will produce, not the stale cached ones.
    """
    import zlsx

    if open_book is None:
        open_book = zlsx.open
    paths = list(paths)
    if infer_files > 0:
        paths = paths[:infer_files]

    fields: list = []
    sampled = 0
    empty: list[str] = []
    for path in paths:
        with open_book(path) as book:
            for sheet_idx in resolve_sheets(sheet_option, book.sheets):
                with book.sheet(sheet_idx).rows() as rows:
                    it = iter(rows)
                    head = next(it, None) if header else None
                    sample = list(_islice(it, infer_rows))
                if (header and head is None) or (not header and not sample):
                    empty.append(f"{path} sheet {sheet_idx}")
                    continue
                fields = merge_fields(
                    fields, infer_fields(head, sample, collapse_void=False)
                )
                sampled += 1

    if sampled == 0:
        shown = ", ".join(empty[:5]) + (" …" if len(empty) > 5 else "")
        raise ValueError(
            f"cannot infer a schema: no sampled sheet had data ({shown}); "
            "supply one with .schema(...)"
        )
    return ddl(collapse_void_fields(fields))


def land_atomically(path: str, payload: bytes) -> None:
    """Put `payload` at `path` so no reader ever sees a partial workbook.

    Writes a sibling temp file and renames it over the target, which is
    atomic on POSIX and on the UC Volumes FUSE mount. Two attempts at the
    same partition — speculative execution, or a retry racing a
    straggler — therefore resolve to one complete file instead of
    interleaving their bytes into the destination.

    Filesystems that reject the rename fall back to one direct write: a
    far narrower torn window than streaming row-by-row, though not
    atomic. That is the best available there, and it is why this is a
    fallback rather than the primary path.
    """
    tmp = f"{path}.zlsx-tmp-{_uuid.uuid4().hex[:12]}"
    try:
        with open(tmp, "wb") as fh:
            fh.write(payload)
        try:
            os.replace(tmp, path)
        except OSError:
            with open(path, "wb") as fh:
                fh.write(payload)
    finally:
        if os.path.exists(tmp):
            try:
                os.remove(tmp)
            except OSError:
                pass


def merge_fields(a, b):
    """Widen two inferred field lists into one that holds both.

    Positional, because the read path coerces positionally
    (:func:`coerce_row` zips row cells against kinds by index) — so
    column *identity* here is the ordinal, not the name. The first
    source names each column; a later source that disagrees on the name
    at that position does not rename it. Kinds widen through
    :func:`merge_kinds`; a column present in only one source, or
    nullable in either, comes out nullable.

    Used to infer one schema across several (file, sheet) sources
    instead of trusting the first one to represent the rest.
    """
    if not a:
        return list(b)
    if not b:
        return list(a)

    out = []
    for i in range(max(len(a), len(b))):
        fa = a[i] if i < len(a) else None
        fb = b[i] if i < len(b) else None
        if fa is None or fb is None:
            # Ragged across sources: the column is absent somewhere, so
            # every row that source contributes is a null in it.
            name, kind, _ = fa or fb
            out.append((name, kind, True))
            continue
        name, kind_a, null_a = fa
        _, kind_b, null_b = fb
        out.append((name, merge_kinds(kind_a, kind_b), null_a or null_b))
    return out


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


# ─── Batch recalc (§12.4) ─────────────────────────────────────────────
#
# `zlsx.recalc="true"` makes the read observe the values the formula
# engine computes, not whatever stale <v> caches the workbook carries.
# Source files are NEVER mutated: recalc happens in memory per open,
# and the result exists only as a buffer.
#
# The one-buffer rule, driver and executor alike: read the source bytes
# ONCE, SHA-256 THAT buffer, and hand THE SAME buffer to the engine.
# Verification and parsing share one immutable snapshot, so there is no
# verify-then-reopen window for a concurrent writer to slip through.

RECALC_DEFAULT_CACHE_BYTES = 512 * 1024 * 1024

_TRUE_WORDS = ("true", "1", "yes")
_FALSE_WORDS = ("false", "0", "no")


class SnapshotDriftError(RuntimeError):
    """Executor-side refusal: the source bytes no longer match the
    digest the driver planned against."""


class EngineFingerprintMismatch(RuntimeError):
    """Executor-side refusal: this process's engine identity differs
    from the one the driver planned with."""


def parse_recalc_options(options, *, streaming: bool = False):
    """Validate the ``zlsx.recalc*`` option namespace.

    Returns ``None`` when recalc is off, else a dict of the resolved
    knobs. Unlike the boolean options elsewhere in the source, a value
    outside true/false vocabulary raises instead of reading as false —
    a typo must not silently hand back stale caches. Streaming + recalc
    is refused HERE, at option validation, per §12.4.
    """
    raw = options.get("zlsx.recalc")
    if raw is None:
        active = False
    else:
        word = raw.strip().lower()
        if word in _TRUE_WORDS:
            active = True
        elif word in _FALSE_WORDS:
            active = False
        else:
            raise ValueError(f"zlsx.recalc must be true or false, got {raw!r}")
    if not active:
        return None
    if streaming:
        raise ValueError(
            "zlsx.recalc is batch-only; streaming + recalc is refused "
            "(recalculate in a batch job, then stream the result)"
        )
    try:
        utc_offset_min = int(options.get("zlsx.recalcUtcOffsetMin", "0"))
    except ValueError:
        raise ValueError(
            "zlsx.recalcUtcOffsetMin must be an integer number of minutes, "
            f"got {options['zlsx.recalcUtcOffsetMin']!r}"
        ) from None
    try:
        cache_max = int(
            options.get("zlsx.recalcCacheMaxBytes", str(RECALC_DEFAULT_CACHE_BYTES))
        )
    except ValueError:
        raise ValueError(
            "zlsx.recalcCacheMaxBytes must be an integer byte count, "
            f"got {options['zlsx.recalcCacheMaxBytes']!r}"
        ) from None
    if cache_max < 0:
        raise ValueError(
            f"zlsx.recalcCacheMaxBytes must be >= 0 (0 disables the "
            f"cache), got {cache_max}"
        )
    return {
        "utc_offset_min": utc_offset_min,
        "cache_max_bytes": cache_max,
        # Fixed at every layer's defaults; not exposed as Spark options.
        "mode": "excel",
        "profile": "windows_1252",
        "on_unsupported": "refuse",
    }


def resolve_recalc_context(recalc_opts, *, now=None, seed=None) -> dict:
    """Resolve the full run context ONCE, on the driver, before any
    recalc runs. Every recalc in the job — the driver's inference pass
    and every executor partition — replays exactly these values, so one
    read observes one logical instant and one RNG stream, and a task
    retry cannot drift. The clock/entropy reads live here in the caller,
    never in the library (§5.5)."""
    if now is None:
        now = _time.time_ns() // 1_000_000
    if seed is None:
        seed = int.from_bytes(os.urandom(8), "little")
    return {
        "now": int(now),
        "utc_offset_min": recalc_opts["utc_offset_min"],
        "seed": int(seed) & 0xFFFFFFFFFFFFFFFF,
        "mode": recalc_opts["mode"],
        "profile": recalc_opts["profile"],
        "on_unsupported": recalc_opts["on_unsupported"],
    }


def _read_file_bytes(path: str) -> bytes:
    # Module-level seam: the race fixture monkeypatches this to mutate
    # the file immediately after the read, proving digest verification
    # and parsing share the one buffer this returns.
    with open(path, "rb") as fh:
        return fh.read()


def _recalc_editor_snapshot(buf: bytes, ctx: dict) -> bytes:
    """One in-memory recalc over ``buf`` → the serialized snapshot.
    The context is replayed verbatim, so equal (bytes, context, engine)
    means byte-identical snapshots."""
    import zlsx

    with zlsx.Editor.from_bytes(buf) as ed:
        ed.recalculate(
            now=ctx["now"],
            utc_offset_min=ctx["utc_offset_min"],
            seed=ctx["seed"],
            mode=ctx["mode"],
            profile=ctx["profile"],
            on_unsupported=ctx["on_unsupported"],
        )
        return ed.save_to_buffer()


def driver_snapshot(path: str, ctx: dict):
    """Driver side of §12.4: read the source bytes once, digest that
    buffer, recalc from the same buffer. Returns
    ``(digest_hex, snapshot_bytes)`` — the digest is what partition
    inputs carry, the snapshot is what schema inference and planning
    read. The source file is never written."""
    buf = _read_file_bytes(path)
    return _hashlib.sha256(buf).hexdigest(), _recalc_editor_snapshot(buf, ctx)


class RecalcCache:
    """Byte-bounded LRU of recalced snapshots.

    Keyed by (workbook digest, every resolved run input, engine
    fingerprint, recalc options) — never by digest alone, so two reads
    with different contexts can never share a snapshot. An entry is
    charged its snapshot's byte length; one larger than the whole bound
    is not admitted. Documented optimization, not a correctness
    dependency: a miss just recalculates."""

    def __init__(self, max_bytes: int):
        self.max_bytes = max_bytes
        self._entries: dict = {}  # key -> snapshot bytes, LRU-ordered
        self._bytes = 0

    def resize(self, max_bytes: int) -> None:
        self.max_bytes = max_bytes
        self._evict()

    def get(self, key):
        snap = self._entries.get(key)
        if snap is not None:
            self._entries.pop(key)
            self._entries[key] = snap  # re-insert = most recently used
        return snap

    def put(self, key, snap: bytes) -> None:
        if len(snap) > self.max_bytes:
            return
        old = self._entries.pop(key, None)
        if old is not None:
            self._bytes -= len(old)
        self._entries[key] = snap
        self._bytes += len(snap)
        self._evict()

    def _evict(self) -> None:
        while self._bytes > self.max_bytes and self._entries:
            oldest = next(iter(self._entries))
            self._bytes -= len(self._entries.pop(oldest))


# One cache per executor process; resized to each read's bound on use.
_EXECUTOR_CACHE = RecalcCache(0)


def recalc_cache_key(digest: str, ctx: dict, fingerprint: str) -> tuple:
    return (
        digest,
        ctx["now"],
        ctx["utc_offset_min"],
        ctx["seed"],
        ctx["mode"],
        ctx["profile"],
        ctx["on_unsupported"],
        fingerprint,
    )


def executor_snapshot(
    path: str,
    digest: str,
    ctx: dict,
    fingerprint: str,
    cache_max_bytes: int,
    *,
    _cache: RecalcCache | None = None,
) -> bytes:
    """Executor side of §12.4 — the one-buffer rule.

    Refuses when this process's engine fingerprint differs from the
    driver's (:class:`EngineFingerprintMismatch`) or when the source
    bytes no longer hash to the planned digest
    (:class:`SnapshotDriftError`). Otherwise recalcs THE SAME buffer
    the digest was computed over. Everything here is a pure function of
    (partition inputs, file bytes, engine), so a task retry re-derives
    the identical snapshot or the identical refusal."""
    import zlsx

    local = zlsx.engine_fingerprint()
    if local != fingerprint:
        raise EngineFingerprintMismatch(
            f"engine fingerprint mismatch: the driver planned with "
            f"{fingerprint!r} but this executor runs {local!r}; align "
            "py-zlsx/libzlsx across the cluster before recalc reads"
        )
    buf = _read_file_bytes(path)
    actual = _hashlib.sha256(buf).hexdigest()
    if actual != digest:
        raise SnapshotDriftError(
            f"{path}: source bytes changed since planning (sha256 "
            f"{actual[:16]}… != planned {digest[:16]}…); workbooks must "
            "stay immutable while a recalc read runs — land new "
            "versions atomically under a new name"
        )
    if cache_max_bytes <= 0:
        return _recalc_editor_snapshot(buf, ctx)
    cache = _EXECUTOR_CACHE if _cache is None else _cache
    cache.resize(cache_max_bytes)
    key = recalc_cache_key(digest, ctx, fingerprint)
    snap = cache.get(key)
    if snap is None:
        snap = _recalc_editor_snapshot(buf, ctx)
        cache.put(key, snap)
    return snap
