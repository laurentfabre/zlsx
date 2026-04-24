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


def test_writer_add_data_validation_numeric_and_custom_round_trip(tmp_path):
    """Round-trip numeric / custom data validations through the
    Python writer and read every extended field back. Guards the
    writer-DV extended ABI (0.2.6+)."""
    import zlsx._ffi as ffi

    if not ffi._HAS_DATA_VALIDATION_EXT or not ffi._HAS_READER_DV_EXT:
        pytest.skip("loaded libzlsx predates extended DV ABI (0.2.6+)")

    out = tmp_path / "dv_ext_writer.xlsx"
    with zlsx.write(out) as w:
        sheet = w.add_sheet("Num")
        sheet.add_data_validation_numeric("B2:B10", "whole", "between", "1", "100")
        sheet.add_data_validation_numeric("C3", "decimal", "greater_than", "0")
        sheet.add_data_validation_numeric("D4", "date", "less_than", "45658")
        sheet.add_data_validation_numeric("E5", "text_length", "between", "3", "20")
        # Custom — XML-special char `<` must round-trip clean.
        sheet.add_data_validation_custom("F6", "AND(F6>0,F6<LEN(A1))")
        sheet.add_data_validation_list("G7", ["Yes", "No"])  # mixed with list
        sheet.write_row(["hdr"])

    with zlsx.open(out) as book:
        dvs = book.data_validations(0)

    assert len(dvs) == 6
    # List entry emits first per writer ordering.
    assert dvs[0].kind == "list"
    assert dvs[0].values == ("Yes", "No")

    assert dvs[1].kind == "whole"
    assert dvs[1].op == "between"
    assert dvs[1].formula1 == "1"
    assert dvs[1].formula2 == "100"

    assert dvs[2].kind == "decimal"
    assert dvs[2].op == "greater_than"
    assert dvs[2].formula1 == "0"
    assert dvs[2].formula2 == ""

    assert dvs[3].kind == "date"
    assert dvs[3].op == "less_than"
    assert dvs[3].formula1 == "45658"

    assert dvs[4].kind == "text_length"
    assert dvs[4].op == "between"
    assert dvs[4].formula1 == "3"
    assert dvs[4].formula2 == "20"

    assert dvs[5].kind == "custom"
    assert dvs[5].op is None
    assert dvs[5].formula1 == "AND(F6>0,F6<LEN(A1))"


def test_shared_strings_enumeration_and_rich_discovery(tmp_path):
    """The iter37 audit flagged `Book.rich_text(sst_idx)` as
    effectively undiscoverable — Python callers couldn't enumerate
    which SST entries carry rich-text runs. iter45 closed it with
    `Book.shared_strings_count()` + `shared_string_at(idx)` +
    `shared_strings()`. This test proves the round-trip: write a
    book with one plain + one rich entry, then enumerate and
    rediscover which is which."""
    import zlsx._ffi as ffi

    if not ffi._HAS_SST_ENUM:
        pytest.skip("loaded libzlsx predates SST enum ABI (0.2.6+)")

    out = tmp_path / "sst_enum.xlsx"
    with zlsx.write(out) as w:
        sheet = w.add_sheet("S")
        sheet.write_rich_row([
            "plain-label",
            [
                zlsx.RichRun("bold-part ", bold=True),
                zlsx.RichRun("italic-part", italic=True),
            ],
        ])

    with zlsx.open(out) as book:
        # Count matches writer's emission order (plain first, then rich).
        assert book.shared_strings_count() == 2
        assert book.shared_string_at(0) == "plain-label"
        assert book.shared_string_at(1) == "bold-part italic-part"

        # shared_strings() materialises everything.
        all_sst = book.shared_strings()
        assert all_sst == ["plain-label", "bold-part italic-part"]

        # Discoverability loop: for every entry, check whether it's rich.
        rich_indices = []
        for i in range(book.shared_strings_count()):
            if book.rich_text(i) is not None:
                rich_indices.append(i)
        assert rich_indices == [1]

        # Out-of-range raises IndexError per the documented contract.
        with pytest.raises(IndexError, match="sst_idx .* out of range"):
            book.shared_string_at(99)


def test_rich_text_runs_parse_bold_italic(tmp_path):
    """Build a minimal xlsx with rich-text SST entries via raw zipfile
    (the writer doesn't emit rich text today) and verify the reader
    surfaces `<b/>` / `<i/>` correctly via `Book.rich_text(sst_idx)`.
    Plain single-run SST entries must return None (zero-overhead path)."""
    import zipfile
    import zlsx._ffi as ffi

    if not ffi._HAS_RICH_RUNS:
        pytest.skip("loaded libzlsx predates rich-text ABI (0.2.6+)")

    xlsx_path = tmp_path / "rich.xlsx"
    sst_xml = (
        b"<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
        b"<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"3\" uniqueCount=\"3\">"
        b"<si><t>plain</t></si>"
        b"<si><r><rPr><b/></rPr><t>bold</t></r><r><rPr><i/></rPr><t> italic</t></r></si>"
        b"<si><r><rPr><b/><i/></rPr><t>R&amp;D</t></r></si>"
        b"</sst>"
    )
    workbook_xml = (
        b"<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
        b"<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
        b"xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
        b"<sheets><sheet name=\"S\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>"
    )
    workbook_rels = (
        b"<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
        b"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        b"<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>"
        b"<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>"
        b"</Relationships>"
    )
    root_rels = (
        b"<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
        b"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        b"<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>"
        b"</Relationships>"
    )
    content_types = (
        b"<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
        b"<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
        b"<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
        b"<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
        b"<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"
        b"<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>"
        b"<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"
        b"</Types>"
    )
    sheet_xml = (
        b"<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
        b"<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
        b"<sheetData><row r=\"1\"><c r=\"A1\" t=\"s\"><v>0</v></c></row></sheetData></worksheet>"
    )

    with zipfile.ZipFile(xlsx_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", root_rels)
        z.writestr("xl/workbook.xml", workbook_xml)
        z.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        z.writestr("xl/sharedStrings.xml", sst_xml)
        z.writestr("xl/worksheets/sheet1.xml", sheet_xml)

    with zlsx.open(xlsx_path) as book:
        # Plain SST entry → None (zero-overhead path).
        assert book.rich_text(0) is None
        # Multi-run bold + italic.
        runs = book.rich_text(1)
        assert runs is not None
        assert len(runs) == 2
        assert runs[0].text == "bold"
        assert runs[0].bold and not runs[0].italic
        assert runs[1].text == " italic"
        assert runs[1].italic and not runs[1].bold
        # Entity-decoded rich text.
        runs = book.rich_text(2)
        assert runs is not None
        assert len(runs) == 1
        assert runs[0].text == "R&D"
        assert runs[0].bold and runs[0].italic
        # Out-of-range SST index → None (count returns 0).
        assert book.rich_text(999) is None


def test_rich_text_runs_color_size_font(tmp_path):
    """Rich-text color / size / font_name round-trip through the
    reader. Theme colors stay None (we don't resolve theme.xml)."""
    import zipfile
    import zlsx._ffi as ffi

    if not ffi._HAS_RICH_RUNS_EXT:
        pytest.skip("loaded libzlsx predates rich-text ext ABI (0.2.6+)")

    xlsx_path = tmp_path / "rich_ext.xlsx"
    sst_xml = (
        b"<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
        b"<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"2\">"
        b"<si><r><rPr><b/><sz val=\"14\"/><color rgb=\"FFFF0000\"/><rFont val=\"Arial\"/></rPr><t>styled</t></r></si>"
        b"<si><r><rPr><color theme=\"1\"/><sz val=\"11.5\"/></rPr><t>themed</t></r></si>"
        b"</sst>"
    )
    workbook_xml = (
        b"<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
        b"<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
        b"xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
        b"<sheets><sheet name=\"S\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>"
    )
    workbook_rels = (
        b"<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
        b"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        b"<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>"
        b"<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>"
        b"</Relationships>"
    )
    root_rels = (
        b"<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
        b"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        b"<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>"
        b"</Relationships>"
    )
    content_types = (
        b"<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
        b"<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
        b"<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
        b"<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
        b"<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"
        b"<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>"
        b"<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"
        b"</Types>"
    )
    sheet_xml = (
        b"<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
        b"<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
        b"<sheetData><row r=\"1\"><c r=\"A1\" t=\"s\"><v>0</v></c></row></sheetData></worksheet>"
    )

    with zipfile.ZipFile(xlsx_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", root_rels)
        z.writestr("xl/workbook.xml", workbook_xml)
        z.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        z.writestr("xl/sharedStrings.xml", sst_xml)
        z.writestr("xl/worksheets/sheet1.xml", sheet_xml)

    with zlsx.open(xlsx_path) as book:
        runs = book.rich_text(0)
        assert runs is not None
        assert len(runs) == 1
        assert runs[0].text == "styled"
        assert runs[0].bold
        assert runs[0].color_argb == 0xFFFF0000
        assert runs[0].size == 14.0
        assert runs[0].font_name == "Arial"

        # Theme color stays None; size still parses.
        runs = book.rich_text(1)
        assert runs is not None
        assert runs[0].color_argb is None
        assert runs[0].size == 11.5
        assert runs[0].font_name == ""


def test_sheet_read_all_and_module_read_helper(tmp_path):
    """`Sheet.read_all()` + `zlsx.read()` materialise rows into
    list-of-lists suitable for pandas.DataFrame / polars.DataFrame.
    No optional dependencies on those libraries — plain Python."""
    out = tmp_path / "readhelper.xlsx"
    with zlsx.write(out) as w:
        s = w.add_sheet("Data")
        s.write_row(["name", "qty", "price"])
        s.write_row(["apple", 3, 1.5])
        s.write_row(["banana", 7, 0.3])
        w.add_sheet("Other").write_row(["x"])

    # Sheet.read_all(header=False): every row in one list.
    with zlsx.open(out) as book:
        header, rows = book.sheet(0).read_all()
        assert header is None
        assert rows == [
            ["name", "qty", "price"],
            ["apple", 3, 1.5],
            ["banana", 7, 0.3],
        ]

    # Sheet.read_all(header=True): first row split out.
    with zlsx.open(out) as book:
        header, rows = book.sheet(0).read_all(header=True)
        assert header == ["name", "qty", "price"]
        assert rows == [
            ["apple", 3, 1.5],
            ["banana", 7, 0.3],
        ]

    # Module-level zlsx.read — one-shot, closes book.
    header, rows = zlsx.read(out, header=True)
    assert header == ["name", "qty", "price"]
    assert len(rows) == 2

    # Sheet by name.
    header, rows = zlsx.read(out, sheet="Other")
    assert rows == [["x"]]

    # Out-of-range index → ZlsxError.
    with pytest.raises(zlsx.ZlsxError, match="out of range"):
        zlsx.read(out, sheet=99)

    # Unknown sheet name → ZlsxError.
    with pytest.raises(zlsx.ZlsxError, match="not found"):
        zlsx.read(out, sheet="Missing")


def test_to_excel_serial_round_trip_with_parse_date(tmp_path):
    """Full date round-trip: Python datetime → to_excel_serial →
    write as numeric cell with date style → read via parse_date →
    Python datetime. Matches the iter46/47 intent: one-call
    conversion at both ends."""
    import datetime as _dt
    import zlsx._ffi as ffi

    if not ffi._HAS_TO_EXCEL_SERIAL or not ffi._HAS_PARSE_DATE:
        pytest.skip("loaded libzlsx predates to_excel_serial ABI (0.2.6+)")

    # A datetime.date and a datetime.datetime both round-trip.
    d_plain = _dt.date(2023, 1, 1)
    d_stamped = _dt.datetime(2024, 6, 15, 12, 34, 56)

    out = tmp_path / "dates_rt.xlsx"
    with zlsx.write(out) as w:
        date_style = w.add_style(zlsx.Style(number_format="yyyy-mm-dd"))
        dt_style = w.add_style(zlsx.Style(number_format="yyyy-mm-dd h:mm:ss"))
        sheet = w.add_sheet("S")
        sheet.write_row(
            [zlsx.to_excel_serial(d_plain), zlsx.to_excel_serial(d_stamped)],
            styles=[date_style, dt_style],
        )

    with zlsx.open(out) as book:
        with book.sheet(0).rows() as rows:
            next(rows)
            assert rows.parse_date(0) == _dt.datetime(2023, 1, 1)
            assert rows.parse_date(1) == _dt.datetime(2024, 6, 15, 12, 34, 56)

    # Rejection paths.
    with pytest.raises(ValueError, match="round-trippable date range"):
        zlsx.to_excel_serial(_dt.date(1800, 1, 1))
    with pytest.raises(ValueError, match="round-trippable date range"):
        zlsx.to_excel_serial(_dt.date(1900, 2, 28))  # pre-leap-bug exclusion
    with pytest.raises(TypeError, match="datetime.date or datetime.datetime"):
        zlsx.to_excel_serial("not a date")


def test_rows_parse_date_auto_converts_date_styled_cells(tmp_path):
    """Python callers can parse date-styled numeric cells directly
    via `Rows.parse_date(col_idx)` without manually chaining
    style_indices + is_date_format + fromExcelSerial."""
    import datetime as _dt
    import zlsx._ffi as ffi

    if not ffi._HAS_PARSE_DATE or not ffi._HAS_NUM_FMT:
        pytest.skip("loaded libzlsx predates parse_date ABI (0.2.6+)")

    out = tmp_path / "parse_date.xlsx"
    with zlsx.write(out) as w:
        date_style = w.add_style(zlsx.Style(number_format="yyyy-mm-dd"))
        pct_style = w.add_style(zlsx.Style(number_format="0.00%"))
        sheet = w.add_sheet("S")
        sheet.write_row(["hdr"])
        sheet.write_row(
            [44927, 0.25, 42, "txt"],
            styles=[date_style, pct_style, 0, 0],
        )

    with zlsx.open(out) as book:
        with book.sheet(0).rows() as rows:
            next(rows)  # header
            next(rows)  # data

            # col 0: date-styled — decodes to 2023-01-01.
            d0 = rows.parse_date(0)
            assert d0 == _dt.datetime(2023, 1, 1)
            # col 1: percentage-styled — not a date.
            assert rows.parse_date(1) is None
            # col 2: plain integer, no style — not a date.
            assert rows.parse_date(2) is None
            # col 3: string cell — not a date.
            assert rows.parse_date(3) is None
            # col 99: out of range — None.
            assert rows.parse_date(99) is None


def test_rows_style_indices_and_book_number_format(tmp_path):
    """Round-trip: writer emits styled cells with custom number
    formats, reader gets back per-cell style indices via
    `Rows.style_indices()` + resolves them via `Book.number_format` /
    `is_date_format`. Covers the iter29 symmetry-closer."""
    import zlsx._ffi as ffi

    if not ffi._HAS_NUM_FMT:
        pytest.skip("loaded libzlsx predates numFmt ABI (0.2.6+)")

    out = tmp_path / "numfmt.xlsx"
    with zlsx.write(out) as w:
        date_style = w.add_style(zlsx.Style(number_format="yyyy-mm-dd"))
        pct_style = w.add_style(zlsx.Style(number_format="0.00%"))
        sheet = w.add_sheet("S")
        sheet.write_row(["hdr"])
        sheet.write_row(
            [44927, 0.25, 42],
            styles=[date_style, pct_style, 0],
        )

    with zlsx.open(out) as book:
        with book.sheet(0).rows() as rows:
            next(rows)  # header row
            cells = next(rows)
            assert cells == [44927, 0.25, 42]
            styles = rows.style_indices()
            assert len(styles) == 3
            s0, s1, s2 = styles
            # Date column resolves back to custom numFmt + isDateFormat.
            assert s0 is not None
            assert book.number_format(s0) == "yyyy-mm-dd"
            assert book.is_date_format(s0) is True
            # Percentage custom code, not a date.
            assert s1 is not None
            assert book.number_format(s1) == "0.00%"
            assert book.is_date_format(s1) is False
            # Plain integer column with style 0 (default General).
            if s2 is not None:
                assert book.is_date_format(s2) is False

        # Out-of-range style index → None.
        assert book.number_format(99999) is None
        assert book.is_date_format(99999) is False


def test_book_cell_font_round_trip(tmp_path):
    """Writer emits bold/colored/sized/named font styles; reader
    resolves them via `Book.cell_font(style_idx)`."""
    import zlsx._ffi as ffi

    if not ffi._HAS_CELL_FONT:
        pytest.skip("loaded libzlsx predates cell_font ABI (0.2.6+)")

    out = tmp_path / "font.xlsx"
    with zlsx.write(out) as w:
        bold_style = w.add_style(zlsx.Style(
            font_bold=True,
            font_color_argb=0xFFFF0000,
            font_size=14,
            font_name="Courier New",
        ))
        plain_style = w.add_style(zlsx.Style(font_italic=True))
        sheet = w.add_sheet("S")
        sheet.write_row(
            ["bold-red", "italic", "bare"],
            styles=[bold_style, plain_style, 0],
        )

    with zlsx.open(out) as book:
        with book.sheet(0).rows() as rows:
            next(rows)
            styles = rows.style_indices()
            assert len(styles) == 3
            s0, s1, s2 = styles

            f0 = book.cell_font(s0)
            assert f0 is not None
            assert f0.bold and not f0.italic
            assert f0.color_argb == 0xFFFF0000
            assert f0.size == 14.0
            assert f0.name == "Courier New"

            f1 = book.cell_font(s1)
            assert f1 is not None and f1.italic and not f1.bold

            # Default font (xfId 0 or whatever the writer left) still
            # resolves to a non-None Font, even if all optionals are null.
            if s2 is not None:
                assert book.cell_font(s2) is not None

        # Out-of-range style idx → None.
        assert book.cell_font(99999) is None


def test_book_cell_fill_round_trip(tmp_path):
    """Writer emits a red solid fill; reader resolves via
    `Book.cell_fill(style_idx)`. Style 0 resolves to the writer's
    default fill (patternType="none")."""
    import zlsx._ffi as ffi

    if not ffi._HAS_CELL_FILL:
        pytest.skip("loaded libzlsx predates cell_fill ABI (0.2.6+)")

    out = tmp_path / "fill.xlsx"
    with zlsx.write(out) as w:
        red = w.add_style(zlsx.Style(
            fill_pattern="solid",
            fill_fg_argb=0xFFFF0000,
        ))
        sheet = w.add_sheet("S")
        sheet.write_row(["red", "plain"], styles=[red, 0])

    with zlsx.open(out) as book:
        with book.sheet(0).rows() as rows:
            next(rows)
            styles = rows.style_indices()
            s0, s1 = styles

            f0 = book.cell_fill(s0)
            assert f0 is not None
            assert f0.pattern == "solid"
            assert f0.fg_color_argb == 0xFFFF0000

            # Default writer style resolves to patternType="none".
            if s1 is not None:
                f1 = book.cell_fill(s1)
                assert f1 is not None
                assert f1.pattern == "none"
                assert f1.fg_color_argb is None

        assert book.cell_fill(99999) is None


def test_book_cell_border_round_trip(tmp_path):
    """Writer emits a boxed cell; reader resolves via
    `Book.cell_border(style_idx)`. Sides without a border come back
    with `style=""`."""
    import zlsx._ffi as ffi

    if not ffi._HAS_CELL_BORDER:
        pytest.skip("loaded libzlsx predates cell_border ABI (0.2.6+)")

    out = tmp_path / "border.xlsx"
    with zlsx.write(out) as w:
        boxed = w.add_style(zlsx.Style(
            border_left=zlsx.BorderSide(style="thin", color_argb=0xFF000000),
            border_right=zlsx.BorderSide(style="thin", color_argb=0xFF000000),
            border_top=zlsx.BorderSide(style="medium", color_argb=0xFFFF0000),
            border_bottom=zlsx.BorderSide(style="medium", color_argb=0xFFFF0000),
        ))
        sheet = w.add_sheet("S")
        sheet.write_row(["boxed", "plain"], styles=[boxed, 0])

    with zlsx.open(out) as book:
        with book.sheet(0).rows() as rows:
            next(rows)
            styles = rows.style_indices()
            s0, s1 = styles

            b0 = book.cell_border(s0)
            assert b0 is not None
            assert b0.left.style == "thin"
            assert b0.left.color_argb == 0xFF000000
            assert b0.right.style == "thin"
            assert b0.top.style == "medium"
            assert b0.top.color_argb == 0xFFFF0000
            assert b0.bottom.style == "medium"
            assert b0.diagonal.style == ""

            if s1 is not None:
                b1 = book.cell_border(s1)
                assert b1 is not None
                assert b1.left.style == ""
                assert b1.top.style == ""

        assert book.cell_border(99999) is None


def test_sheet_writer_write_rich_row_round_trip(tmp_path):
    """Python writes a row mixing plain + rich-text cells via
    `write_rich_row`; reader round-trips the formatting through
    `Book.rich_text(sst_idx)`. Guards the iter36 C-ABI + Python
    binding — iter33 landed the Zig API but not the FFI surface."""
    import zlsx._ffi as ffi

    if not ffi._HAS_WRITE_RICH_ROW or not ffi._HAS_RICH_RUNS:
        pytest.skip("loaded libzlsx predates write_rich_row ABI (0.2.6+)")

    out = tmp_path / "rich_writer.xlsx"
    with zlsx.write(out) as w:
        sheet = w.add_sheet("S")
        sheet.write_rich_row([
            "plain",
            [
                zlsx.RichRun("hello ", bold=True),
                zlsx.RichRun("world", italic=True, color_argb=0xFFFF0000,
                             size=12.0, font_name="Arial"),
            ],
            42,
        ])

    with zlsx.open(out) as book:
        # SST order: "plain" at 0, rich at 1.
        assert book.rich_text(0) is None
        runs = book.rich_text(1)
        assert runs is not None
        assert len(runs) == 2
        assert runs[0].text == "hello "
        assert runs[0].bold and not runs[0].italic
        assert runs[1].text == "world"
        assert runs[1].italic and not runs[1].bold
        assert runs[1].color_argb == 0xFFFF0000
        assert runs[1].size == 12.0
        assert runs[1].font_name == "Arial"


def test_dxf_extended_fields_border_size(tmp_path):
    """iter49 extends `Dxf` with font_size + per-side borders.
    Register a fully-populated Dxf and verify styles.xml contains
    the expected `<sz>` + `<border>` fragments."""
    import zipfile
    import zlsx._ffi as ffi

    if not ffi._HAS_CONDITIONAL_FORMAT:
        pytest.skip("loaded libzlsx predates CF ABI (0.2.6+)")

    out = tmp_path / "dxf_ext.xlsx"
    with zlsx.write(out) as w:
        rich_dxf = w.add_dxf(zlsx.Dxf(
            font_bold=True,
            font_color_argb=0xFFFF0000,
            font_size=16.0,
            fill_fg_argb=0xFFFFFF00,
            border_left=zlsx.BorderSide(style="thin", color_argb=0xFF000000),
            border_right=zlsx.BorderSide(style="thin", color_argb=0xFF000000),
            border_top=zlsx.BorderSide(style="medium", color_argb=0xFFFF00FF),
            border_bottom=zlsx.BorderSide(style="medium", color_argb=0xFFFF00FF),
        ))
        sheet = w.add_sheet("S")
        sheet.add_conditional_format_cell_is(
            "A1:A10", "greater_than", "100", None, rich_dxf
        )
        sheet.write_row(["hdr"])

    with zipfile.ZipFile(out) as z:
        styles_xml = z.read("xl/styles.xml").decode("utf-8")

    # font_size renders as <sz val="16"/>.
    assert '<sz val="16"/>' in styles_xml
    # Border block present with all 4 sides.
    assert '<border>' in styles_xml
    assert '<left style="thin">' in styles_xml
    assert '<right style="thin">' in styles_xml
    assert '<top style="medium">' in styles_xml
    assert '<bottom style="medium">' in styles_xml
    # Border colors.
    assert '<color rgb="FF000000"/>' in styles_xml
    assert '<color rgb="FFFF00FF"/>' in styles_xml


def test_conditional_formatting_round_trip(tmp_path):
    """Write cellIs + expression CF rules via Python; extract the
    generated xlsx and verify the sheet XML + styles.xml contain
    the expected conditionalFormatting / dxfs blocks."""
    import zipfile
    import zlsx._ffi as ffi

    if not ffi._HAS_CONDITIONAL_FORMAT:
        pytest.skip("loaded libzlsx predates CF ABI (0.2.6+)")

    out = tmp_path / "cf.xlsx"
    with zlsx.write(out) as w:
        red = w.add_dxf(zlsx.Dxf(font_bold=True, font_color_argb=0xFFFF0000))
        green = w.add_dxf(zlsx.Dxf(fill_fg_argb=0xFF00FF00))
        # Dedup check.
        red2 = w.add_dxf(zlsx.Dxf(font_bold=True, font_color_argb=0xFFFF0000))
        assert red == red2

        sheet = w.add_sheet("S")
        sheet.add_conditional_format_cell_is("B2:B10", "greater_than", "100", None, red)
        sheet.add_conditional_format_cell_is("C2:C10", "between", "0", "50", red)
        sheet.add_conditional_format_expression("A1:Z100", "MOD(ROW(),2)=0", green)
        sheet.write_row(["hdr"])

        # Rejection paths.
        with pytest.raises(ValueError, match="conditional-format operator"):
            sheet.add_conditional_format_cell_is("A1", "bogus", "1", None, red)
        with pytest.raises(zlsx.ZlsxError, match="InvalidDataValidation"):
            sheet.add_conditional_format_cell_is("A1", "equal", "", None, red)
        with pytest.raises(zlsx.ZlsxError, match="UnknownDxfId"):
            sheet.add_conditional_format_expression("A1", "ROW()=1", 99)

    # Extract the xlsx and verify the CF + dxfs wire up.
    with zipfile.ZipFile(out) as z:
        sheet_xml = z.read("xl/worksheets/sheet1.xml").decode("utf-8")
        styles_xml = z.read("xl/styles.xml").decode("utf-8")

    assert '<conditionalFormatting sqref="B2:B10">' in sheet_xml
    assert 'operator="greaterThan"' in sheet_xml
    assert '<formula>100</formula>' in sheet_xml
    assert '<conditionalFormatting sqref="C2:C10">' in sheet_xml
    assert 'operator="between"' in sheet_xml
    assert '<cfRule type="expression"' in sheet_xml
    assert 'MOD(ROW(),2)=0' in sheet_xml

    assert '<dxfs count="2">' in styles_xml
    assert '<color rgb="FFFF0000"/>' in styles_xml
    assert '<fgColor rgb="FF00FF00"/>' in styles_xml


def test_sheet_writer_add_comment_round_trip(tmp_path):
    """Python writer emits cell comments; reader round-trips them via
    `Book.comments(sheet_idx)`. Matrix-flip gate for iter38."""
    import zlsx._ffi as ffi

    if not ffi._HAS_COMMENT_WRITER or not ffi._HAS_COMMENTS:
        pytest.skip("loaded libzlsx predates comment writer ABI (0.2.6+)")

    out = tmp_path / "comments_writer.xlsx"
    with zlsx.write(out) as w:
        sheet = w.add_sheet("S")
        sheet.add_comment("B2", "Alice", "review this")
        sheet.add_comment("C3", "Bob & Co", "R&D notes")
        sheet.add_comment("D4", "Alice", "follow-up")  # author dedup
        sheet.write_row(["hdr"])

        # Rejection paths.
        with pytest.raises(zlsx.ZlsxError, match="InvalidCommentRef"):
            sheet.add_comment("", "a", "b")
        with pytest.raises(zlsx.ZlsxError, match="InvalidCommentRef"):
            sheet.add_comment("A1:B2", "a", "b")

    with zlsx.open(out) as book:
        cs = book.comments(0)
        assert len(cs) == 3
        assert cs[0].top_left == zlsx.CellRef(col=1, row=2)
        assert cs[0].author == "Alice"
        assert cs[0].text == "review this"
        assert cs[1].author == "Bob & Co"  # entity-decoded
        assert cs[1].text == "R&D notes"
        assert cs[2].author == "Alice"  # same dedup'd author
        assert cs[2].text == "follow-up"


def test_book_comments_parses_authors_refs_text(tmp_path):
    """Build a minimal xlsx with a comments1.xml part and verify
    `Book.comments(sheet_idx)` returns the right refs, authors, and
    entity-decoded plain text. Rich-text bodies get flattened
    (concatenated <t> slices, decoded)."""
    import zipfile
    import zlsx._ffi as ffi

    if not ffi._HAS_COMMENTS:
        pytest.skip("loaded libzlsx predates comments ABI (0.2.6+)")

    xlsx_path = tmp_path / "comments.xlsx"
    content_types = (
        b"<?xml version=\"1.0\"?>"
        b"<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
        b"<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
        b"<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
        b"<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"
        b"<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"
        b"<Override PartName=\"/xl/comments1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml\"/>"
        b"</Types>"
    )
    root_rels = (
        b"<?xml version=\"1.0\"?>"
        b"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        b"<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>"
        b"</Relationships>"
    )
    workbook = (
        b"<?xml version=\"1.0\"?>"
        b"<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
        b"xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
        b"<sheets><sheet name=\"S\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>"
    )
    wb_rels = (
        b"<?xml version=\"1.0\"?>"
        b"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        b"<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>"
        b"</Relationships>"
    )
    sheet1 = (
        b"<?xml version=\"1.0\"?>"
        b"<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
        b"<sheetData><row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>hello</t></is></c></row></sheetData></worksheet>"
    )
    sheet1_rels = (
        b"<?xml version=\"1.0\"?>"
        b"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        b"<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments\" Target=\"../comments1.xml\"/>"
        b"</Relationships>"
    )
    comments1 = (
        b"<?xml version=\"1.0\"?>"
        b"<comments xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
        b"<authors><author>Alice</author><author>Bob &amp; Co</author></authors>"
        b"<commentList>"
        b"<comment ref=\"B2\" authorId=\"0\"><text><r><t>review this</t></r></text></comment>"
        b"<comment ref=\"C3\" authorId=\"1\"><text><r><rPr><b/></rPr><t xml:space=\"preserve\">R&amp;D </t></r><r><t>notes</t></r></text></comment>"
        b"</commentList></comments>"
    )

    with zipfile.ZipFile(xlsx_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", root_rels)
        z.writestr("xl/workbook.xml", workbook)
        z.writestr("xl/_rels/workbook.xml.rels", wb_rels)
        z.writestr("xl/worksheets/sheet1.xml", sheet1)
        z.writestr("xl/worksheets/_rels/sheet1.xml.rels", sheet1_rels)
        z.writestr("xl/comments1.xml", comments1)

    with zlsx.open(xlsx_path) as book:
        cs = book.comments(0)
        assert len(cs) == 2

        assert cs[0].top_left == zlsx.CellRef(col=1, row=2)
        assert cs[0].author == "Alice"
        assert cs[0].text == "review this"

        assert cs[1].top_left == zlsx.CellRef(col=2, row=3)
        assert cs[1].author == "Bob & Co"
        assert cs[1].text == "R&D notes"


def test_writer_add_data_validation_rejects_invalid_inputs(tmp_path):
    """Exercise every error path on the extended writer DV APIs so the
    rejection behaviour from the Zig writer surfaces cleanly."""
    import zlsx._ffi as ffi

    if not ffi._HAS_DATA_VALIDATION_EXT:
        pytest.skip("loaded libzlsx predates extended writer DV ABI (0.2.6+)")

    with zlsx.write(tmp_path / "ignored.xlsx") as w:
        sheet = w.add_sheet("S")
        # Unknown kind / op → ValueError (Python-side validation).
        with pytest.raises(ValueError, match="data validation kind"):
            sheet.add_data_validation_numeric("A1", "bogus", "equal", "1")
        with pytest.raises(ValueError, match="data validation operator"):
            sheet.add_data_validation_numeric("A1", "whole", "bogus", "1")
        # two-formula mismatch: equal + formula2 set → InvalidDataValidation.
        with pytest.raises(zlsx.ZlsxError, match="InvalidDataValidation"):
            sheet.add_data_validation_numeric("A1", "whole", "equal", "1", "2")
        # between + missing formula2 → InvalidDataValidation.
        with pytest.raises(zlsx.ZlsxError, match="InvalidDataValidation"):
            sheet.add_data_validation_numeric("A1", "whole", "between", "1")
        # empty formula → InvalidDataValidation.
        with pytest.raises(zlsx.ZlsxError, match="InvalidDataValidation"):
            sheet.add_data_validation_numeric("A1", "whole", "equal", "")
        # bad range → InvalidHyperlinkRange (shared A1 validator).
        with pytest.raises(zlsx.ZlsxError, match="InvalidHyperlinkRange"):
            sheet.add_data_validation_numeric("", "whole", "equal", "1")
        # Custom rejection paths.
        with pytest.raises(zlsx.ZlsxError, match="InvalidDataValidation"):
            sheet.add_data_validation_custom("A1", "")
        with pytest.raises(zlsx.ZlsxError, match="InvalidHyperlinkRange"):
            sheet.add_data_validation_custom("", "A1>0")
        # Save something so the writer closes cleanly.
        sheet.write_row(["x"])
