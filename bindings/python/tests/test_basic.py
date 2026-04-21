"""Smoke tests for py-zlsx.

Runs against the corpus tarballs materialised by
``scripts/fetch_test_corpus.sh`` at ``<repo>/tests/corpus/``. Skips the
corpus-heavy tests if the files aren't present.
"""

from __future__ import annotations

from pathlib import Path

import pytest

import zlsx

REPO_ROOT = Path(__file__).resolve().parents[3]
CORPUS = REPO_ROOT / "tests" / "corpus"


def _skip_if_missing(name: str) -> Path:
    path = CORPUS / name
    if not path.exists():
        pytest.skip(
            f"corpus file {name!r} not present — run scripts/fetch_test_corpus.sh"
        )
    return path


def test_version_string_matches_package():
    import zlsx._ffi as ffi

    lib_version = ffi.lib.zlsx_version_string().decode("utf-8")
    # Package version tracks the library's major.minor; patch may drift.
    assert lib_version.startswith("0.2."), (
        f"unexpected library version: {lib_version}"
    )


def test_open_invalid_path_raises():
    with pytest.raises(zlsx.ZlsxError):
        zlsx.open("/nonexistent/path/does/not/exist.xlsx")


def test_frictionless_two_sheets():
    path = _skip_if_missing("frictionless_2sheets.xlsx")
    with zlsx.open(path) as book:
        assert book.sheets == ["Sheet1", "Sheet2"]

        s1 = book.sheet(0)
        assert s1.name == "Sheet1"
        rows = list(s1.rows())
        # Header + 2 data rows.
        assert len(rows) == 3
        assert rows[0] == ["header1", "header2", "header3"]
        assert rows[1] == ["a", "b", "c"]
        assert rows[2] == ["d", "e", "f"]


def test_sheet_selection_by_name():
    path = _skip_if_missing("frictionless_2sheets.xlsx")
    with zlsx.open(path) as book:
        # Select by name should find index 1.
        sheet = book.sheet("Sheet2")
        assert sheet.index == 1
        assert sheet.name == "Sheet2"


def test_sheet_missing_name_raises_keyerror():
    path = _skip_if_missing("frictionless_2sheets.xlsx")
    with zlsx.open(path) as book:
        with pytest.raises(KeyError):
            book.sheet("NoSuchSheet")


def test_sheet_out_of_range_raises_indexerror():
    path = _skip_if_missing("frictionless_2sheets.xlsx")
    with zlsx.open(path) as book:
        with pytest.raises(IndexError):
            book.sheet(99)


def test_worldbank_row_count_matches_bench():
    path = _skip_if_missing("worldbank_catalog.xlsx")
    with zlsx.open(path) as book:
        rows = list(book.sheet(0).rows())
        # Matches the benchmark table (161 rows).
        assert len(rows) == 161


def test_cell_type_mapping_guess_types():
    path = _skip_if_missing("openpyxl_guess_types.xlsx")
    with zlsx.open(path) as book:
        rows = list(book.sheet(0).rows())
        # Just assert the call returns something without crashing; the
        # content is interpretation-sensitive across readers.
        assert rows  # non-empty
        for row in rows:
            for cell in row:
                # Every cell is one of the documented Python types.
                assert cell is None or isinstance(cell, (str, int, float, bool))


def test_close_book_while_rows_live():
    """Refcount keeps the underlying state alive — we can drop the Book
    handle and keep iterating rows without crashing."""
    path = _skip_if_missing("frictionless_2sheets.xlsx")

    book = zlsx.open(path)
    rows = book.sheet(0).rows()
    book.close()  # drop the Book handle — rows holds its own reference

    collected = list(rows)
    assert len(collected) == 3
