"""Unit tests for zlsx._tabular — the pure logic behind zlsx.spark.

No pyspark needed: everything here runs in plain CPython. The Spark
integration itself is covered by test_spark_integration.py (skipped
when pyspark / a local JVM is unavailable).
"""
from __future__ import annotations

import datetime as dt
import os

import pytest

from zlsx import _tabular as tab


# ─── widening lattice ─────────────────────────────────────────────────


def test_kind_of_bool_before_int():
    assert tab.kind_of(True) == "boolean"
    assert tab.kind_of(1) == "bigint"
    assert tab.kind_of(1.5) == "double"
    assert tab.kind_of("x") == "string"
    assert tab.kind_of(None) == "void"


@pytest.mark.parametrize("a,b,want", [
    ("void", "bigint", "bigint"),
    ("bigint", "bigint", "bigint"),
    ("bigint", "double", "double"),
    ("double", "bigint", "double"),
    ("bigint", "string", "string"),
    ("boolean", "bigint", "string"),
    ("boolean", "double", "string"),
    ("boolean", "string", "string"),
])
def test_merge_kinds(a, b, want):
    assert tab.merge_kinds(a, b) == want
    assert tab.merge_kinds(b, a) == want  # symmetric


# ─── inference beyond the first data row ──────────────────────────────


def test_infer_widens_int_to_double_on_later_row():
    fields = tab.infer_fields(["n"], [[1], [2], [3.5]])
    assert fields == [("n", "double", False)]


def test_infer_widens_to_string_on_later_row():
    fields = tab.infer_fields(["n"], [[1], [2], ["oops"]])
    assert fields == [("n", "string", False)]


def test_infer_all_none_column_is_nullable_string():
    fields = tab.infer_fields(["a", "b"], [[1, None], [2, None]])
    assert fields == [("a", "bigint", False), ("b", "string", True)]


def test_infer_short_rows_mark_missing_columns_nullable():
    fields = tab.infer_fields(["a", "b"], [[1, "x"], [2]])
    assert fields[1] == ("b", "string", True)


def test_infer_positional_names_without_header():
    fields = tab.infer_fields(None, [[1, "x"], [2, "y"]])
    assert [f[0] for f in fields] == ["_c0", "_c1"]


def test_infer_dedupes_duplicate_header_names():
    fields = tab.infer_fields(["a", "a", "a"], [[1, 2, 3]])
    assert [f[0] for f in fields] == ["a", "a_1", "a_2"]


def test_infer_blank_header_cell_gets_positional_name():
    fields = tab.infer_fields(["a", None, ""], [[1, 2, 3]])
    assert [f[0] for f in fields] == ["a", "_c1", "_c2"]


def test_infer_empty_sheet_raises():
    with pytest.raises(ValueError, match="empty sheet"):
        tab.infer_fields(None, [])


def test_ddl_renders_and_escapes_backticks():
    fields = [("plain", "bigint", False), ("we`ird", "string", True)]
    assert tab.ddl(fields) == "`plain` bigint, `we``ird` string"


# ─── coercion ─────────────────────────────────────────────────────────

KINDS = ["bigint", "double", "string", "boolean"]


def test_coerce_pads_short_rows_with_none():
    out = tab.coerce_row([1], KINDS, "permissive", "f", 2)
    assert out == [1, None, None, None]


def test_coerce_upcasts_int_for_double_column():
    out = tab.coerce_row([1, 2, "x", True], KINDS, "permissive", "f", 2)
    assert out == [1, 2.0, "x", True]
    assert isinstance(out[1], float)


def test_coerce_stringifies_numerics_in_string_column():
    out = tab.coerce_row([1, 2.5, 3, False], KINDS, "permissive", "f", 2)
    assert out[2] == "3"
    assert tab.coerce_row([1, 2.5, True, False], KINDS, "permissive", "f", 2)[2] == "true"


def test_coerce_permissive_nulls_mismatches():
    out = tab.coerce_row(["x", "y", "z", "w"], KINDS, "permissive", "f", 2)
    assert out == [None, None, "z", None]


def test_coerce_failfast_raises_with_context():
    with pytest.raises(tab.CellCoercionError, match=r"book.xlsx sheet 'S' row 7 col 0"):
        tab.coerce_row(["x"], KINDS, "failfast", "book.xlsx sheet 'S'", 7)


def test_coerce_long_row_truncates_or_raises():
    row = [1, 2.0, "x", True, "extra"]
    assert len(tab.coerce_row(row, KINDS, "permissive", "f", 2)) == 4
    with pytest.raises(tab.CellCoercionError, match="5 cells"):
        tab.coerce_row(row, KINDS, "failfast", "f", 2)


def test_coerce_bool_never_masquerades_as_number():
    assert tab.coerce_row([True], ["bigint"], "permissive", "f", 2) == [None]
    assert tab.coerce_row([True], ["double"], "permissive", "f", 2) == [None]


def test_coerce_serial_to_timestamp_and_date():
    # 45292 is 2024-01-01 in the 1900 date system.
    assert tab.coerce_row([45292.0], ["timestamp"], "failfast", "f", 2) == [
        dt.datetime(2024, 1, 1)
    ]
    assert tab.coerce_row([45292], ["date"], "failfast", "f", 2) == [
        dt.date(2024, 1, 1)
    ]


def test_serial_to_datetime_epoch():
    assert tab.serial_to_datetime(1.0) == dt.datetime(1899, 12, 31)
    assert tab.serial_to_datetime(45292.5) == dt.datetime(2024, 1, 1, 12, 0)


# ─── path expansion ───────────────────────────────────────────────────


def _touch(p):
    p.write_bytes(b"")
    return str(p)


def test_expand_paths_file_dir_glob_and_commas(tmp_path):
    a = _touch(tmp_path / "a.xlsx")
    b = _touch(tmp_path / "b.xlsm")
    _touch(tmp_path / "~$b.xlsm")  # Excel lock file: ignored in dir listing
    _touch(tmp_path / "c.csv")     # ignored in dir listing
    sub = tmp_path / "sub"
    sub.mkdir()
    c = _touch(sub / "c.xlsx")

    assert tab.expand_paths(a) == [a]
    assert tab.expand_paths(str(tmp_path)) == [a, b]
    assert tab.expand_paths(str(tmp_path / "*.xlsx")) == [a]
    assert tab.expand_paths(f"{a}, {c}") == [a, c]
    assert tab.expand_paths(f"{a},{a}") == [a]  # deduped


def test_expand_paths_missing_raises(tmp_path):
    with pytest.raises(FileNotFoundError, match="no such workbook"):
        tab.expand_paths(str(tmp_path / "ghost.xlsx"))
    with pytest.raises(FileNotFoundError, match="glob matched no files"):
        tab.expand_paths(str(tmp_path / "*.xlsx"))
    (tmp_path / "empty").mkdir()
    with pytest.raises(FileNotFoundError, match="no .xlsx"):
        tab.expand_paths(str(tmp_path / "empty"))


# ─── sheet resolution and row-range planning ──────────────────────────


def test_resolve_sheets_star_names_indices():
    names = ["Sales", "Targets", "Notes"]
    assert tab.resolve_sheets("*", names) == [0, 1, 2]
    assert tab.resolve_sheets("Targets", names) == [1]
    assert tab.resolve_sheets("2,Sales", names) == [2, 0]
    assert tab.resolve_sheets("Sales,0", names) == [0]  # deduped


def test_resolve_sheets_errors():
    names = ["Sales"]
    with pytest.raises(KeyError, match="Ghost"):
        tab.resolve_sheets("Ghost", names)
    with pytest.raises(IndexError, match="out of range"):
        tab.resolve_sheets("3", names)


def test_snapshot_files_empty_zone_and_fingerprints(tmp_path):
    zone = tmp_path / "zone"
    zone.mkdir()
    assert tab.snapshot_files(str(zone)) == {}
    assert tab.snapshot_files(str(zone / "*.xlsx")) == {}

    a = tmp_path / "zone" / "a.xlsx"
    a.write_bytes(b"AAAA")
    _touch(zone / "~$a.xlsx")  # lock file ignored
    snap = tab.snapshot_files(str(zone))
    assert list(snap) == [str(a)]
    mtime_ns, size = snap[str(a)]
    assert size == 4 and mtime_ns > 0


def test_snapshot_files_missing_literal_still_raises(tmp_path):
    with pytest.raises(FileNotFoundError, match="no such workbook"):
        tab.snapshot_files(str(tmp_path / "ghost.xlsx"))


def test_diff_snapshots_new_changed_deleted():
    old = {"/z/a.xlsx": [1, 10], "/z/gone.xlsx": [1, 5]}
    new = {
        "/z/a.xlsx": [1, 10],   # unchanged
        "/z/b.xlsx": [2, 20],   # new
        "/z/c.xlsx": [3, 30],   # new
    }
    assert tab.diff_snapshots(old, new) == ["/z/b.xlsx", "/z/c.xlsx"]
    assert tab.diff_snapshots(old, old) == []
    # changed mtime OR size re-ingests
    assert tab.diff_snapshots({"/z/a.xlsx": [1, 10]}, {"/z/a.xlsx": [9, 10]}) == ["/z/a.xlsx"]
    assert tab.diff_snapshots({"/z/a.xlsx": [1, 10]}, {"/z/a.xlsx": [1, 99]}) == ["/z/a.xlsx"]
    # from-empty ingests everything, sorted
    assert tab.diff_snapshots({}, new) == ["/z/a.xlsx", "/z/b.xlsx", "/z/c.xlsx"]


def test_plan_row_ranges():
    assert tab.plan_row_ranges(10, 0) == [(0, 10)]
    assert tab.plan_row_ranges(10, 100) == [(0, 10)]
    assert tab.plan_row_ranges(10, 5) == [(0, 5), (5, 10)]
    assert tab.plan_row_ranges(11, 5) == [(0, 5), (5, 10), (10, 11)]
    assert tab.plan_row_ranges(0, 5) == [(0, 0)]


# ─── cross-source widening (multi-file / multi-sheet inference) ───────


def test_merge_fields_widens_kind_across_sources():
    a = tab.infer_fields(["n"], [[1], [2]], collapse_void=False)
    b = tab.infer_fields(["n"], [[3.5]], collapse_void=False)
    assert tab.merge_fields(a, b) == [("n", "double", False)]
    assert tab.merge_fields(b, a) == [("n", "double", False)]


def test_merge_fields_incompatible_kinds_fall_to_string():
    a = tab.infer_fields(["c"], [[1]], collapse_void=False)
    b = tab.infer_fields(["c"], [["text"]], collapse_void=False)
    assert tab.merge_fields(a, b) == [("c", "string", False)]


def test_merge_fields_all_null_column_does_not_pin_string():
    """The regression that motivated collapse_void=False: a blank column
    in the first workbook must not force the merged column to string
    when a later workbook shows it is numeric."""
    a = tab.infer_fields(["v"], [[None], [None]], collapse_void=False)
    b = tab.infer_fields(["v"], [[7]], collapse_void=False)
    merged = tab.merge_fields(a, b)
    assert merged == [("v", "bigint", True)]
    assert tab.collapse_void_fields(merged) == [("v", "bigint", True)]
    # …and with no evidence anywhere it still lands on string.
    only_void = tab.merge_fields(a, a)
    assert tab.collapse_void_fields(only_void) == [("v", "string", True)]


def test_merge_fields_ragged_width_marks_missing_nullable():
    a = tab.infer_fields(["x", "y"], [[1, 2]], collapse_void=False)
    b = tab.infer_fields(["x"], [[3]], collapse_void=False)
    merged = tab.merge_fields(a, b)
    assert merged == [("x", "bigint", False), ("y", "bigint", True)]
    # Symmetric: the wider source may arrive second.
    assert tab.merge_fields(b, a) == merged


def test_merge_fields_nullability_is_sticky():
    a = tab.infer_fields(["x"], [[1]], collapse_void=False)
    b = tab.infer_fields(["x"], [[None], [2]], collapse_void=False)
    assert tab.merge_fields(a, b) == [("x", "bigint", True)]


def test_merge_fields_first_source_names_the_columns():
    a = tab.infer_fields(["region"], [["N"]], collapse_void=False)
    b = tab.infer_fields(["zone"], [["S"]], collapse_void=False)
    # Positional merge — the read path coerces by index, so a renamed
    # column in a later file does not rename the schema.
    assert tab.merge_fields(a, b) == [("region", "string", False)]


def test_merge_fields_identity_with_empty():
    a = tab.infer_fields(["x"], [[1]], collapse_void=False)
    assert tab.merge_fields([], a) == a
    assert tab.merge_fields(a, []) == a
    assert tab.merge_fields([], []) == []


def test_infer_fields_collapse_void_default_unchanged():
    """Default behaviour is the pre-existing one — void collapses to
    string in a single-source inference."""
    assert tab.infer_fields(["v"], [[None]]) == [("v", "string", True)]


# ─── infer_schema over real workbooks ─────────────────────────────────


def _book(path, rows, sheet="Sheet1"):
    import zlsx

    with zlsx.write(path) as w:
        s = w.add_sheet(sheet)
        for r in rows:
            s.write_row(r)


def test_infer_schema_widens_across_files(tmp_path):
    _book(tmp_path / "a.xlsx", [["region", "units"], ["N", 1]])
    _book(tmp_path / "b.xlsx", [["region", "units"], ["S", 2.5]])
    # File a alone would say bigint; the pair must say double.
    assert tab.infer_schema([str(tmp_path / "a.xlsx")]) == "`region` string, `units` bigint"
    assert tab.infer_schema(
        [str(tmp_path / "a.xlsx"), str(tmp_path / "b.xlsx")]
    ) == "`region` string, `units` double"


def test_infer_schema_widens_across_sheets(tmp_path):
    import zlsx

    p = tmp_path / "multi.xlsx"
    with zlsx.write(p) as w:
        s1 = w.add_sheet("One")
        s1.write_row(["v"])
        s1.write_row([1])
        s2 = w.add_sheet("Two")
        s2.write_row(["v"])
        s2.write_row(["text"])

    assert tab.infer_schema([str(p)], sheet_option="0") == "`v` bigint"
    assert tab.infer_schema([str(p)], sheet_option="*") == "`v` string"


def test_infer_schema_respects_infer_files_bound(tmp_path):
    _book(tmp_path / "a.xlsx", [["v"], [1]])
    _book(tmp_path / "b.xlsx", [["v"], [2.5]])
    paths = [str(tmp_path / "a.xlsx"), str(tmp_path / "b.xlsx")]
    # Bounded to the first file: the second never opens, so no widening.
    assert tab.infer_schema(paths, infer_files=1) == "`v` bigint"
    assert tab.infer_schema(paths, infer_files=0) == "`v` double"


def test_infer_schema_skips_empty_sheets(tmp_path):
    import zlsx

    with zlsx.write(tmp_path / "empty.xlsx") as w:
        w.add_sheet("Blank")
    _book(tmp_path / "data.xlsx", [["v"], [7]])

    # A blank workbook in the zone contributes nothing rather than
    # aborting the whole inference.
    assert tab.infer_schema(
        [str(tmp_path / "empty.xlsx"), str(tmp_path / "data.xlsx")]
    ) == "`v` bigint"


def test_infer_schema_all_empty_raises(tmp_path):
    import zlsx

    with zlsx.write(tmp_path / "empty.xlsx") as w:
        w.add_sheet("Blank")
    with pytest.raises(ValueError, match="no sampled sheet had data"):
        tab.infer_schema([str(tmp_path / "empty.xlsx")])


# ─── atomic part-file landing ─────────────────────────────────────────


def test_land_atomically_writes_payload(tmp_path):
    target = tmp_path / "part.xlsx"
    tab.land_atomically(str(target), b"PK\x03\x04payload")
    assert target.read_bytes() == b"PK\x03\x04payload"


def test_land_atomically_overwrites_and_leaves_no_temp(tmp_path):
    target = tmp_path / "part.xlsx"
    target.write_bytes(b"stale")
    tab.land_atomically(str(target), b"fresh")
    assert target.read_bytes() == b"fresh"
    assert [p.name for p in tmp_path.iterdir()] == ["part.xlsx"]


def test_land_atomically_cleans_up_when_the_write_fails(tmp_path):
    target = tmp_path / "sub" / "part.xlsx"   # parent does not exist
    with pytest.raises(OSError):
        tab.land_atomically(str(target), b"x")
    assert not tmp_path.joinpath("sub").exists()


def test_land_atomically_never_exposes_a_partial_file(tmp_path, monkeypatch):
    """The point of the temp+rename: a reader polling the destination
    sees either nothing or the complete payload, never a prefix."""
    target = tmp_path / "part.xlsx"
    seen = []

    real_replace = os.replace

    def spy(src, dst):
        # At the moment of the rename the destination must not exist
        # yet, and the temp must already hold every byte.
        seen.append((os.path.exists(dst), os.path.getsize(src)))
        return real_replace(src, dst)

    monkeypatch.setattr(os, "replace", spy)
    tab.land_atomically(str(target), b"0123456789")
    assert seen == [(False, 10)]
    assert target.read_bytes() == b"0123456789"


# ─── batch recalc (§12.4): option validation ──────────────────────────


def test_parse_recalc_options_off_by_default():
    assert tab.parse_recalc_options({}) is None
    assert tab.parse_recalc_options({"zlsx.recalc": "false"}) is None
    assert tab.parse_recalc_options({"zlsx.recalc": " FALSE "}) is None
    # Off + streaming is fine — only the combination refuses.
    assert tab.parse_recalc_options({}, streaming=True) is None


def test_parse_recalc_options_active_defaults():
    got = tab.parse_recalc_options({"zlsx.recalc": "true"})
    assert got == {
        "utc_offset_min": 0,                       # UTC, like every layer
        "cache_max_bytes": 512 * 1024 * 1024,
        "mode": "excel",
        "profile": "windows_1252",
        "on_unsupported": "refuse",
    }


def test_parse_recalc_options_rejects_typos():
    """A typo must not silently read as false and hand back stale
    caches — unlike the tolerant boolean options elsewhere."""
    with pytest.raises(ValueError, match="zlsx.recalc"):
        tab.parse_recalc_options({"zlsx.recalc": "ture"})


def test_parse_recalc_options_streaming_refused():
    with pytest.raises(ValueError, match="batch-only"):
        tab.parse_recalc_options({"zlsx.recalc": "true"}, streaming=True)


def test_parse_recalc_options_knob_validation():
    base = {"zlsx.recalc": "true"}
    got = tab.parse_recalc_options(
        {**base, "zlsx.recalcUtcOffsetMin": "60",
         "zlsx.recalcCacheMaxBytes": "0"}
    )
    assert got["utc_offset_min"] == 60 and got["cache_max_bytes"] == 0

    with pytest.raises(ValueError, match="recalcUtcOffsetMin"):
        tab.parse_recalc_options({**base, "zlsx.recalcUtcOffsetMin": "PST"})
    with pytest.raises(ValueError, match="recalcCacheMaxBytes"):
        tab.parse_recalc_options({**base, "zlsx.recalcCacheMaxBytes": "big"})
    with pytest.raises(ValueError, match=">= 0"):
        tab.parse_recalc_options({**base, "zlsx.recalcCacheMaxBytes": "-1"})


def test_resolve_recalc_context_echoes_and_defaults():
    opts = tab.parse_recalc_options(
        {"zlsx.recalc": "true", "zlsx.recalcUtcOffsetMin": "120"}
    )
    ctx = tab.resolve_recalc_context(opts, now=1700000000000, seed=42)
    assert ctx == {
        "now": 1700000000000, "utc_offset_min": 120, "seed": 42,
        "mode": "excel", "profile": "windows_1252",
        "on_unsupported": "refuse",
    }
    # Defaulted clock/entropy resolve to concrete replayable ints, and
    # the seed always fits u64.
    d = tab.resolve_recalc_context(opts, seed=-1)
    assert isinstance(d["now"], int) and d["now"] > 0
    assert d["seed"] == 0xFFFFFFFFFFFFFFFF


# ─── batch recalc: the byte-bounded LRU cache ─────────────────────────


def test_recalc_cache_lru_eviction_order():
    c = tab.RecalcCache(10)
    c.put("a", b"aaaa")
    c.put("b", b"bbbb")
    assert c.get("a") == b"aaaa"        # touches a: b is now the LRU
    c.put("c", b"cccc")                 # 12 bytes > 10: evict b
    assert c.get("b") is None
    assert c.get("a") == b"aaaa" and c.get("c") == b"cccc"


def test_recalc_cache_oversized_entry_not_admitted():
    c = tab.RecalcCache(4)
    c.put("big", b"12345")
    assert c.get("big") is None
    c.put("fits", b"1234")
    assert c.get("fits") == b"1234"


def test_recalc_cache_replacement_and_resize():
    c = tab.RecalcCache(8)
    c.put("k", b"1234")
    c.put("k", b"12345678")             # replace: accounting must not leak
    c.put("j", b"x")                    # 9 bytes: evicts k, not j
    assert c.get("k") is None and c.get("j") == b"x"
    c.put("k", b"1234")
    assert c.get("j") == b"x"           # touch j: k is now the LRU
    c.resize(1)                         # shrink evicts k; j still fits
    assert c.get("k") is None and c.get("j") == b"x"


# ─── batch recalc: driver / executor snapshot flows ───────────────────


def _stale_book(path, cached=999.0, a1=1):
    """A1=`a1`, B1='A1+2' with a STALE cached value — the engine
    computes a1+2, the file claims `cached`."""
    import zlsx

    with zlsx.write(str(path)) as w:
        s = w.add_sheet("Calc")
        s.write_row_with_formulas([a1, cached], [None, "A1+2"])
    return str(path)


def _b1(book_bytes):
    import zlsx

    with zlsx.open_bytes(book_bytes) as book:
        _, rows = book.sheet(0).read_all(header=False)
    return rows[0][1]


def _ctx(**over):
    base = tab.resolve_recalc_context(
        tab.parse_recalc_options({"zlsx.recalc": "true"}),
        now=1700000000000, seed=7,
    )
    base.update(over)
    return base


def test_driver_snapshot_recalcs_and_never_mutates_source(tmp_path):
    import hashlib

    path = _stale_book(tmp_path / "m.xlsx")
    before = open(path, "rb").read()
    assert _b1(before) == 999.0          # the read path trusts the cache

    digest, snap = tab.driver_snapshot(path, _ctx())
    assert digest == hashlib.sha256(before).hexdigest()
    assert _b1(snap) == 3.0              # the snapshot carries engine truth
    assert open(path, "rb").read() == before  # source bytes untouched


def test_executor_snapshot_rederives_the_driver_snapshot(tmp_path):
    import zlsx

    path = _stale_book(tmp_path / "m.xlsx")
    ctx = _ctx()
    digest, driver_snap = tab.driver_snapshot(path, ctx)
    exec_snap = tab.executor_snapshot(
        path, digest, ctx, zlsx.engine_fingerprint(), 0
    )
    assert exec_snap == driver_snap      # byte-identical re-derivation


def test_executor_snapshot_refuses_digest_drift_identically_on_retry(tmp_path):
    import zlsx

    path = _stale_book(tmp_path / "m.xlsx")
    ctx = _ctx()
    digest, _ = tab.driver_snapshot(path, ctx)
    _stale_book(path, cached=111.0, a1=50)   # the file drifts after planning

    fp = zlsx.engine_fingerprint()
    with pytest.raises(tab.SnapshotDriftError, match="m.xlsx") as first:
        tab.executor_snapshot(path, digest, ctx, fp, 0)
    with pytest.raises(tab.SnapshotDriftError) as retry:
        tab.executor_snapshot(path, digest, ctx, fp, 0)
    assert str(retry.value) == str(first.value)  # retries re-derive, not heal
    assert "immutable" in str(first.value)


def test_executor_snapshot_refuses_engine_mismatch(tmp_path):
    path = _stale_book(tmp_path / "m.xlsx")
    ctx = _ctx()
    digest, _ = tab.driver_snapshot(path, ctx)
    with pytest.raises(tab.EngineFingerprintMismatch, match="0.0.0"):
        tab.executor_snapshot(
            path, digest, ctx, "zlsx 0.0.0; bogus_rules; nowhere", 0
        )


def test_executor_snapshot_mutate_between_verify_and_open_race(
    tmp_path, monkeypatch
):
    """The §12.4 race fixture. A writer replaces the workbook at the
    worst possible instant — right after the executor's single read.
    Because verification and parsing share that one buffer, the recalc
    must still be of the verified bytes (B1 = 3), not the mutant's
    (B1 = 52), and the executor must not re-open the path."""
    import zlsx

    path = _stale_book(tmp_path / "m.xlsx")
    ctx = _ctx()
    digest, _ = tab.driver_snapshot(path, ctx)

    real_read = tab._read_file_bytes
    reads = []

    def read_then_lose_the_race(p):
        buf = real_read(p)
        reads.append(p)
        _stale_book(p, cached=999.0, a1=50)   # mutant lands immediately after
        return buf

    monkeypatch.setattr(tab, "_read_file_bytes", read_then_lose_the_race)
    snap = tab.executor_snapshot(path, digest, ctx, zlsx.engine_fingerprint(), 0)
    assert _b1(snap) == 3.0
    assert reads == [path]               # the source was read exactly once


def test_executor_snapshot_cache_hit_and_context_isolation(tmp_path):
    import zlsx

    path = _stale_book(tmp_path / "m.xlsx")
    ctx = _ctx()
    digest, _ = tab.driver_snapshot(path, ctx)
    fp = zlsx.engine_fingerprint()

    calls = []
    real = tab._recalc_editor_snapshot

    def counting(buf, c):
        calls.append(c["seed"])
        return real(buf, c)

    bound = 64 * 1024 * 1024
    try:
        tab._recalc_editor_snapshot = counting
        cache = tab.RecalcCache(bound)
        a = tab.executor_snapshot(path, digest, ctx, fp, bound, _cache=cache)
        b = tab.executor_snapshot(path, digest, ctx, fp, bound, _cache=cache)
        assert a == b
        assert len(calls) == 1           # second call was a cache hit
        # A different resolved context must NEVER share the snapshot.
        tab.executor_snapshot(
            path, digest, _ctx(seed=8), fp, bound, _cache=cache
        )
        assert len(calls) == 2
        # cache_max_bytes=0 bypasses the cache entirely.
        tab.executor_snapshot(path, digest, ctx, fp, 0, _cache=cache)
        assert len(calls) == 3
    finally:
        tab._recalc_editor_snapshot = real
