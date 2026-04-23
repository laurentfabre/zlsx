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


def test_writer_stage2_style_fields(tmp_path):
    """Stage-2 fields (size, name, color, alignment, wrap_text) land in
    the emitted styles.xml."""
    import zipfile

    out = tmp_path / "stage2.xlsx"
    with zlsx.write(out) as w:
        fancy = w.add_style(zlsx.Style(
            font_size=18,
            font_name="Arial",
            font_color_argb=0xFFFF0000,
            alignment_horizontal="center",
            wrap_text=True,
        ))
        # Dedup: same spec from a fresh Python Style object returns same id.
        again = w.add_style(zlsx.Style(
            font_size=18,
            font_name="Arial",
            font_color_argb=0xFFFF0000,
            alignment_horizontal="center",
            wrap_text=True,
        ))
        assert fancy == again

        sheet = w.add_sheet("S")
        sheet.write_row(["styled"], styles=[fancy])

    with zipfile.ZipFile(out) as z:
        styles = z.read("xl/styles.xml").decode("utf-8")

    assert '<sz val="18"' in styles
    assert '<name val="Arial"' in styles
    assert 'rgb="FFFF0000"' in styles
    assert 'horizontal="center"' in styles
    assert 'wrapText="1"' in styles
    assert 'applyAlignment="1"' in styles


def test_writer_stage2_invalid_inputs():
    # No path — no save on exit, so errors from add_style don't chain
    # into a NoSheets save failure.
    with zlsx.write() as w:
        with pytest.raises(zlsx.ZlsxError, match="InvalidFontSize"):
            w.add_style(zlsx.Style(font_size=0))
        with pytest.raises(zlsx.ZlsxError, match="InvalidFontName"):
            w.add_style(zlsx.Style(font_name=""))
        with pytest.raises(ValueError, match="alignment_horizontal"):
            w.add_style(zlsx.Style(alignment_horizontal="not-a-real-alignment"))


def test_writer_stage3_fills(tmp_path):
    """Stage-3 fill fields (pattern + fg/bg colors) land in styles.xml."""
    import zipfile

    out = tmp_path / "fills.xlsx"
    with zlsx.write(out) as w:
        yellow = w.add_style(zlsx.Style(
            fill_pattern="solid",
            fill_fg_argb=0xFFFFFF00,
        ))
        striped = w.add_style(zlsx.Style(
            fill_pattern="darkHorizontal",
            fill_fg_argb=0xFF0000FF,
            fill_bg_argb=0xFFFFFFFF,
        ))
        # Dedup same spec.
        again = w.add_style(zlsx.Style(
            fill_pattern="solid",
            fill_fg_argb=0xFFFFFF00,
        ))
        assert yellow == again
        assert striped != yellow

        sheet = w.add_sheet("S")
        sheet.write_row(["a", "b"], styles=[yellow, striped])

    with zipfile.ZipFile(out) as z:
        styles = z.read("xl/styles.xml").decode("utf-8")

    assert 'patternType="solid"' in styles
    assert '<fgColor rgb="FFFFFF00"/>' in styles
    assert 'patternType="darkHorizontal"' in styles
    assert '<fgColor rgb="FF0000FF"/>' in styles
    assert '<bgColor rgb="FFFFFFFF"/>' in styles
    assert 'applyFill="1"' in styles


def test_writer_stage3_unknown_pattern_raises():
    with zlsx.write() as w:
        with pytest.raises(ValueError, match="fill_pattern"):
            w.add_style(zlsx.Style(fill_pattern="not-a-pattern"))


def test_writer_stage4_borders(tmp_path):
    import zipfile

    out = tmp_path / "borders.xlsx"
    with zlsx.write(out) as w:
        box = w.add_style(zlsx.Style(
            border_left=zlsx.BorderSide(style="thin", color_argb=0xFF000000),
            border_right=zlsx.BorderSide(style="thin", color_argb=0xFF000000),
            border_top=zlsx.BorderSide(style="thin", color_argb=0xFF000000),
            border_bottom=zlsx.BorderSide(style="thin", color_argb=0xFF000000),
        ))
        fancy = w.add_style(zlsx.Style(
            border_bottom=zlsx.BorderSide(style="thick", color_argb=0xFFFF0000),
            border_diagonal=zlsx.BorderSide(style="dashed"),
            diagonal_up=True,
        ))
        # Dedup.
        box_again = w.add_style(zlsx.Style(
            border_left=zlsx.BorderSide(style="thin", color_argb=0xFF000000),
            border_right=zlsx.BorderSide(style="thin", color_argb=0xFF000000),
            border_top=zlsx.BorderSide(style="thin", color_argb=0xFF000000),
            border_bottom=zlsx.BorderSide(style="thin", color_argb=0xFF000000),
        ))
        assert box == box_again
        assert fancy != box

        sheet = w.add_sheet("S")
        sheet.write_row(["a", "b"], styles=[box, fancy])

    with zipfile.ZipFile(out) as z:
        styles = z.read("xl/styles.xml").decode("utf-8")

    assert '<borders count="3">' in styles
    assert '<left style="thin"' in styles
    assert '<bottom style="thick"' in styles
    assert '<color rgb="FFFF0000"/>' in styles
    assert 'diagonalUp="1"' in styles
    assert '<diagonal style="dashed"' in styles
    assert 'applyBorder="1"' in styles


def test_writer_stage4_unknown_border_style_raises():
    with zlsx.write() as w:
        with pytest.raises(ValueError, match="border style"):
            w.add_style(zlsx.Style(
                border_left=zlsx.BorderSide(style="not-a-style"),
            ))


def test_writer_stage5_number_formats(tmp_path):
    import zipfile

    out = tmp_path / "numfmt.xlsx"
    with zlsx.write(out) as w:
        money = w.add_style(zlsx.Style(number_format="$#,##0.00"))
        pct = w.add_style(zlsx.Style(number_format="0.00%"))
        money_again = w.add_style(zlsx.Style(number_format="$#,##0.00"))
        assert money == money_again
        assert pct != money

        sheet = w.add_sheet("S")
        sheet.write_row([123.45, 0.9], styles=[money, pct])

    with zipfile.ZipFile(out) as z:
        styles = z.read("xl/styles.xml").decode("utf-8")

    assert '<numFmts count="2">' in styles
    assert 'numFmtId="164"' in styles
    assert 'numFmtId="165"' in styles
    assert 'formatCode="$#,##0.00"' in styles
    assert 'formatCode="0.00%"' in styles
    assert 'applyNumberFormat="1"' in styles


def test_writer_stage5_sheet_features(tmp_path):
    import zipfile

    out = tmp_path / "sheetfeat.xlsx"
    with zlsx.write(out) as w:
        sheet = w.add_sheet("Sheet1")
        sheet.set_column_width(0, 20.5)
        sheet.set_column_width(3, 12)
        sheet.freeze_panes(rows=1, cols=2)
        sheet.set_auto_filter("A1:D1")
        sheet.write_row(["a", "b", "c", "d"])

    with zipfile.ZipFile(out) as z:
        sheet_xml = z.read("xl/worksheets/sheet1.xml").decode("utf-8")

    # Ordering: sheetViews → cols → sheetData → autoFilter
    sv = sheet_xml.index("<sheetViews>")
    cols = sheet_xml.index("<cols>")
    data = sheet_xml.index("<sheetData>")
    af = sheet_xml.index("<autoFilter")
    assert sv < cols < data < af

    assert 'xSplit="2"' in sheet_xml
    assert 'ySplit="1"' in sheet_xml
    assert 'state="frozen"' in sheet_xml
    assert 'width="20.5"' in sheet_xml
    assert 'customWidth="1"' in sheet_xml
    assert 'ref="A1:D1"' in sheet_xml


def test_writer_stage5_invalid_inputs(tmp_path):
    out = tmp_path / "bad.xlsx"
    with zlsx.write(out) as w:
        sheet = w.add_sheet("S")
        sheet.write_row(["a"])
        with pytest.raises(zlsx.ZlsxError, match="InvalidColumnWidth"):
            sheet.set_column_width(0, -5)
        with pytest.raises(zlsx.ZlsxError, match="InvalidAutoFilterRange"):
            sheet.set_auto_filter("")


def test_writer_rejects_unknown_style_id(tmp_path):
    """writeRowStyled must range-check style ids against the registered
    styles — referencing id 1 before any addStyle() call would otherwise
    produce a workbook with `s="1"` but no matching <xf> record."""
    with zlsx.write(tmp_path / "bad.xlsx") as w:
        sheet = w.add_sheet("S")
        # No styles registered yet — id 1 is out of range.
        with pytest.raises(zlsx.ZlsxError, match="UnknownStyleId"):
            sheet.write_row(["x"], styles=[1])
        # Register one style → id 1 is now valid.
        sid = w.add_style(zlsx.Style(font_bold=True))
        assert sid == 1
        sheet.write_row(["ok"], styles=[sid])
        # id 2 is still out of range.
        with pytest.raises(zlsx.ZlsxError, match="UnknownStyleId"):
            sheet.write_row(["x"], styles=[2])


def test_writer_sheet_features_reject_negative_ints(tmp_path):
    """Python's signed ints would silently wrap to UINT32_MAX inside
    ctypes and then overflow inside Zig. Validate upfront with a
    clear ValueError."""
    with zlsx.write(tmp_path / "bad.xlsx") as w:
        sheet = w.add_sheet("S")
        with pytest.raises(ValueError, match="col_idx"):
            sheet.set_column_width(-1, 10)
        with pytest.raises(ValueError, match="rows/cols"):
            sheet.freeze_panes(rows=-1)
        with pytest.raises(ValueError, match="rows/cols"):
            sheet.freeze_panes(cols=-1)


def test_sheetwriter_invalidated_after_writer_close(tmp_path):
    """After Writer.close() (automatic on exit from a `with` block),
    cached SheetWriter references must refuse to call into the C ABI
    — the underlying handle is NULL and would crash on field access.
    """
    w = zlsx.write(tmp_path / "x.xlsx").__enter__()
    sheet = w.add_sheet("S")
    sheet.write_row(["ok"])
    w.__exit__(None, None, None)  # closes + invalidates sheet

    # Every SheetWriter method must raise cleanly, not segfault.
    with pytest.raises(RuntimeError, match="parent Writer was closed"):
        sheet.write_row(["bad"])
    with pytest.raises(RuntimeError, match="parent Writer was closed"):
        sheet.set_column_width(0, 10)
    with pytest.raises(RuntimeError, match="parent Writer was closed"):
        sheet.freeze_panes(1, 0)
    with pytest.raises(RuntimeError, match="parent Writer was closed"):
        sheet.set_auto_filter("A1:B1")


def test_argb_overflow_rejects_with_named_field():
    """ctypes.c_uint32 would silently mask 0x1FFFFFFFF → 0xFFFFFFFF;
    a user typo ships the wrong colour with no warning. Range-check
    upfront and name the offending field."""
    with zlsx.write() as w:
        with pytest.raises(ValueError, match="font_color_argb"):
            w.add_style(zlsx.Style(font_color_argb=0x1FFFFFFFF))
        with pytest.raises(ValueError, match="fill_fg_argb"):
            w.add_style(zlsx.Style(fill_pattern="solid", fill_fg_argb=-1))
        with pytest.raises(ValueError, match="border_left.color_argb"):
            w.add_style(zlsx.Style(
                border_left=zlsx.BorderSide(style="thin", color_argb=0x1_0000_0000),
            ))


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


def test_data_validations_extended_fields_for_list_kind(tmp_path):
    """Reader must surface kind/op/formula1/formula2 on every
    validation — exercise the plumbing through the Python writer,
    which only emits list kinds today. Numeric / custom kinds are
    covered by the Zig round-trip test in src/xlsx.zig."""
    import zlsx._ffi as ffi

    if not ffi._HAS_READER_DV_EXT:
        pytest.skip("loaded libzlsx predates extended DV ABI (0.2.6+)")

    out = tmp_path / "dv_ext_list.xlsx"
    with zlsx.write(out) as w:
        sheet = w.add_sheet("Pick")
        sheet.add_data_validation_list("A1", ["Yes", "No"])
        sheet.write_row(["hdr"])

    with zlsx.open(out) as book:
        dvs = book.data_validations(0)

    assert len(dvs) == 1
    assert dvs[0].kind == "list"
    assert dvs[0].op is None
    assert dvs[0].values == ("Yes", "No")
    # formula1 for a literal list comes through in its CSV form (entity-
    # decoded by the reader, so the outer `&quot;` becomes `"`).
    assert dvs[0].formula1 == "\"Yes,No\""
    assert dvs[0].formula2 == ""
