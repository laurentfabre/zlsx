"""Unit tests for zlsx._tabular — the pure logic behind zlsx.spark.

No pyspark needed: everything here runs in plain CPython. The Spark
integration itself is covered by test_spark_integration.py (skipped
when pyspark / a local JVM is unavailable).
"""
from __future__ import annotations

import datetime as dt

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


def test_plan_row_ranges():
    assert tab.plan_row_ranges(10, 0) == [(0, 10)]
    assert tab.plan_row_ranges(10, 100) == [(0, 10)]
    assert tab.plan_row_ranges(10, 5) == [(0, 5), (5, 10)]
    assert tab.plan_row_ranges(11, 5) == [(0, 5), (5, 10), (10, 11)]
    assert tab.plan_row_ranges(0, 5) == [(0, 0)]
