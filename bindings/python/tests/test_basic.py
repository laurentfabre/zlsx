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


# ─── Writer ────────────────────────────────────────────────────────────


def test_writer_round_trip(tmp_path):
    out = tmp_path / "out.xlsx"
    with zlsx.write(out) as w:
        sheet = w.add_sheet("Summary")
        sheet.write_row(["Name", "Age", "Active", "Pi"])
        sheet.write_row(["Alice", 30, True, 3.14159])
        sheet.write_row(["Bob", 25, False, None])

    assert out.exists()

    with zlsx.open(out) as book:
        assert book.sheets == ["Summary"]
        rows = list(book.sheet(0).rows())
        assert rows[0] == ["Name", "Age", "Active", "Pi"]
        assert rows[1][0] == "Alice"
        assert rows[1][1] == 30
        assert rows[1][2] is True
        assert abs(rows[1][3] - 3.14159) < 1e-9
        assert rows[2][0] == "Bob"
        assert rows[2][1] == 25
        assert rows[2][2] is False


def test_writer_multi_sheet_sst_dedup(tmp_path):
    out = tmp_path / "multi.xlsx"
    with zlsx.write(out) as w:
        s1 = w.add_sheet("Alpha")
        s1.write_row(["hello"])
        s1.write_row(["world"])
        s2 = w.add_sheet("Beta")
        s2.write_row(["hello"])   # dedups against s1
        s2.write_row(["zig"])

    with zlsx.open(out) as book:
        assert book.sheets == ["Alpha", "Beta"]
        a_rows = list(book.sheet("Alpha").rows())
        b_rows = list(book.sheet("Beta").rows())
        assert a_rows == [["hello"], ["world"]]
        assert b_rows == [["hello"], ["zig"]]


def test_writer_rejects_oversized_integer(tmp_path):
    out = tmp_path / "overflow.xlsx"
    with pytest.raises(zlsx.ZlsxError, match="IntegerExceedsExcelPrecision"):
        with zlsx.write(out) as w:
            sheet = w.add_sheet("S")
            sheet.write_row([(1 << 53) + 1])   # not exactly representable


def test_writer_no_save_on_exception(tmp_path):
    out = tmp_path / "aborted.xlsx"
    with pytest.raises(RuntimeError):
        with zlsx.write(out) as w:
            w.add_sheet("S").write_row(["a"])
            raise RuntimeError("caller aborted")
    assert not out.exists(), "exception should skip save"


def test_writer_bool_and_int_distinct(tmp_path):
    """Python bools are ints; verify we emit them as boolean cells, not
    integer cells (openpyxl has historically done the wrong thing here).
    """
    out = tmp_path / "bools.xlsx"
    with zlsx.write(out) as w:
        sheet = w.add_sheet("S")
        sheet.write_row([True, 1, False, 0])

    with zlsx.open(out) as book:
        row = next(book.sheet(0).rows())
        assert row[0] is True
        assert row[1] == 1 and isinstance(row[1], int)
        assert row[2] is False
        assert row[3] == 0 and isinstance(row[3], int)


def test_writer_xml_special_chars_escape(tmp_path):
    out = tmp_path / "entities.xlsx"
    with zlsx.write(out) as w:
        w.add_sheet("R&D").write_row(["a<b & c>d \"e\" 'f'"])

    with zlsx.open(out) as book:
        assert book.sheets == ["R&D"]
        row = next(book.sheet(0).rows())
        assert row[0] == "a<b & c>d \"e\" 'f'"


# ─── Styles (Phase 3b) ────────────────────────────────────────────────


def test_writer_add_style_dedups():
    with zlsx.write() as w:
        bold = w.add_style(zlsx.Style(font_bold=True))
        bold_again = w.add_style(zlsx.Style(font_bold=True))
        italic = w.add_style(zlsx.Style(font_italic=True))
        assert bold == bold_again
        assert italic != bold
        assert bold >= 1   # 0 is reserved for default
        assert italic >= 1


def test_writer_styled_round_trip(tmp_path):
    out = tmp_path / "styled.xlsx"
    with zlsx.write(out) as w:
        bold = w.add_style(zlsx.Style(font_bold=True))
        italic = w.add_style(zlsx.Style(font_italic=True))
        both = w.add_style(zlsx.Style(font_bold=True, font_italic=True))
        sheet = w.add_sheet("Styled")
        sheet.write_row(
            ["bold", "italic", "both", "plain"],
            styles=[bold, italic, both, 0],
        )

    # Reader ignores styles, but must still parse the file cleanly and
    # preserve the cell values.
    with zlsx.open(out) as book:
        row = next(book.sheet(0).rows())
        assert row == ["bold", "italic", "both", "plain"]


def test_writer_styles_length_mismatch(tmp_path):
    with zlsx.write(tmp_path / "x.xlsx") as w:
        bold = w.add_style(zlsx.Style(font_bold=True))
        sheet = w.add_sheet("S")
        with pytest.raises(ValueError, match="styles length"):
            sheet.write_row(["a", "b"], styles=[bold])


def test_writer_no_styles_xml_when_unused(tmp_path):
    """A writer that never calls add_style must produce a byte-identical
    output to v0.2.3 — no styles.xml entry in the archive. This is
    important so upgrades don't perturb hashes of previously-saved files.
    """
    import zipfile

    out = tmp_path / "plain.xlsx"
    with zlsx.write(out) as w:
        w.add_sheet("S").write_row(["hello"])

    with zipfile.ZipFile(out) as z:
        names = set(z.namelist())
    assert "xl/styles.xml" not in names
    assert "xl/sharedStrings.xml" in names  # still present as before
