"""S3d slice 2 — the per-sheet registrations on an opened sheet, through
the editor handle.

``Editor.set_column_width`` / ``set_row_height`` / ``freeze_panes`` /
``set_auto_filter`` / ``add_merged_cell`` / ``add_hyperlink`` /
``add_internal_hyperlink`` / ``add_comment`` /
``add_data_validation_{list,numeric,custom}`` /
``add_conditional_format_{cell_is,expression,color_scale,data_bar}`` are
``Worksheet.set*`` / ``add*`` through ``zlsx_editor_*``: the fresh
writer's registrations on an OPENED workbook, landed in the existing
sheet part at save — the element extended in place where the sheet holds
it, created at its schema slot where not; the sheet's relationships, its
comments part and its VML drawing extended or created with it. Every
workbook here comes from ``zlsx.Writer``: sheet ``S`` holds a merge, a
comment, an external hyperlink and a column width; sheet ``T`` is bare.
"""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest

import zlsx
from zlsx import Dxf


def _needs_attachments():
    import zlsx._ffi as ffi

    if not ffi._HAS_EDITOR_SHEET_ATTACHMENTS:
        pytest.skip("loaded libzlsx predates zlsx_editor_add_merged_cell (0.9.0+)")
    return ffi


def _write_fixture(path: Path) -> None:
    with zlsx.Writer(path) as w:
        w.add_dxf(Dxf(font_bold=True))
        s = w.add_sheet("S")
        s.write_row(["h", 2])
        s.write_row([1, 2])
        s.add_merged_cell("D1:E1")
        s.add_comment("A1", "alice", "old")
        s.add_hyperlink("B1", "https://a.example/")
        s.set_column_width(0, 12)
        t = w.add_sheet("T")
        t.write_row([5])


def _part(path: Path, name: str) -> str:
    with zipfile.ZipFile(path) as z:
        return z.read(name).decode("utf-8")


def _names(path: Path) -> set[str]:
    with zipfile.ZipFile(path) as z:
        return set(z.namelist())


def test_every_registration_lands_in_the_saved_sheet(tmp_path: Path) -> None:
    _needs_attachments()
    src = tmp_path / "src.xlsx"
    out = tmp_path / "out.xlsx"
    _write_fixture(src)
    with zlsx.edit(src) as ed:
        ed.set_column_width(0, 2, 30)
        ed.set_row_height(0, 1, 33)
        ed.freeze_panes(0, 1, 1)
        ed.set_auto_filter(0, "A1:C1")
        ed.add_merged_cell(0, "A5:B5")
        ed.add_hyperlink(0, "A4", "https://b.example/?q=1&r=2")
        ed.add_internal_hyperlink(0, "B4", "T!A1")
        ed.add_comment(0, "B2", "bob", "new")
        ed.add_comment(0, "C2", "alice", "again")
        ed.add_data_validation_list(0, "C1", ["x", "y"])
        ed.add_data_validation_numeric(0, "D1", "whole", "between", "1", "9")
        ed.add_data_validation_custom(0, "D2", "D2>0")
        ed.add_conditional_format_cell_is(0, "A2:A3", "greater_than", "0", None, 0)
        ed.add_conditional_format_expression(0, "B2:B3", "B2>1", 0)
        ed.add_conditional_format_color_scale(0, "A1:A9", 0xFF0000FF, 0xFF00FF00, 0xFFFF0000)
        ed.add_conditional_format_data_bar(0, "B1:B9", 0xFF0000FF)
        ed.add_comment(1, "A1", "carol", "t")
        ed.add_merged_cell(1, "A2:B2")
        ed.save(out)

    sheet1 = _part(out, "xl/worksheets/sheet1.xml")
    for needle in (
        '<sheetViews><sheetView workbookViewId="0"><pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/></sheetView></sheetViews>',
        '<cols><col min="1" max="1" width="12" customWidth="1"/><col min="3" max="3" width="30" customWidth="1"/></cols>',
        '<row r="2" ht="33" customHeight="1">',
        '</sheetData><autoFilter ref="A1:C1"/><mergeCells count="2"><mergeCell ref="D1:E1"/><mergeCell ref="A5:B5"/></mergeCells>',
        '<cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan">',
        '<cfRule type="dataBar" priority="4">',
        '<dataValidations count="3">',
        '<formula1>&quot;x,y&quot;</formula1>',
        '<hyperlinks><hyperlink ref="B1" r:id="rId1"/><hyperlink ref="A4" r:id="rId4"/><hyperlink ref="B4" location="T!A1"/></hyperlinks><legacyDrawing r:id="rId3"/></worksheet>',
    ):
        assert needle in sheet1, needle
    rels1 = _part(out, "xl/worksheets/_rels/sheet1.xml.rels")
    assert 'Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://b.example/?q=1&amp;r=2" TargetMode="External"/>' in rels1
    comments1 = _part(out, "xl/comments1.xml")
    assert "<authors><author>alice</author><author>bob</author></authors>" in comments1
    assert '<comment ref="C2" authorId="0">' in comments1
    vml1 = _part(out, "xl/drawings/vmlDrawing1.vml")
    assert 'id="_x0000_s1027"' in vml1
    sheet2 = _part(out, "xl/worksheets/sheet2.xml")
    assert '</sheetData><mergeCells count="1"><mergeCell ref="A2:B2"/></mergeCells><legacyDrawing r:id="rId2"/></worksheet>' in sheet2
    assert {"xl/comments2.xml", "xl/drawings/vmlDrawing2.vml", "xl/worksheets/_rels/sheet2.xml.rels"} <= _names(out)

    with zlsx.open(out) as book:
        assert len(book.merged_ranges(0)) == 2
        links = book.hyperlinks(0)
        assert len(links) == 3
        assert links[1].url == "https://b.example/?q=1&r=2"
        assert links[2].location == "T!A1"
        comments = book.comments(0)
        assert [c.author for c in comments] == ["alice", "bob", "alice"]
        assert len(book.data_validations(0)) == 3
        assert len(book.comments(1)) == 1
    with zlsx.edit(out) as ed2:
        assert len(ed2.conditional_formats()) == 4
        props = [p for p in ed2.sheet_props() if p["sheet_idx"] == 0][0]
        assert props["pane"]["x_split"] == 1
        assert props["pane"]["y_split"] == 1


def test_statements_about_the_call_stage_nothing_and_the_save_is_the_passthrough(tmp_path: Path) -> None:
    _needs_attachments()
    src = tmp_path / "src.xlsx"
    out = tmp_path / "out.xlsx"
    _write_fixture(src)
    before = src.read_bytes()
    with zlsx.edit(src) as ed:
        with pytest.raises(zlsx.ZlsxError, match="SheetIndexOutOfRange"):
            ed.set_column_width(2, 0, 9)
        with pytest.raises(zlsx.ZlsxError, match="InvalidColumnWidth"):
            ed.set_column_width(0, 0, -1)
        with pytest.raises(zlsx.ZlsxError, match="InvalidRowHeight"):
            ed.set_row_height(0, 0, 500)
        with pytest.raises(zlsx.ZlsxError, match="RowOutOfRange"):
            ed.freeze_panes(0, 1048576, 0)
        with pytest.raises(zlsx.ZlsxError, match="InvalidAutoFilterRange"):
            ed.set_auto_filter(0, "B1:A1")
        with pytest.raises(zlsx.ZlsxError, match="InvalidMergeRange"):
            ed.add_merged_cell(0, "A1")
        with pytest.raises(zlsx.ZlsxError, match="MergeRangeOverlaps"):
            ed.add_merged_cell(0, "E1:F2")
        with pytest.raises(zlsx.ZlsxError, match="InvalidHyperlinkUrl"):
            ed.add_hyperlink(0, "A9", "")
        with pytest.raises(zlsx.ZlsxError, match="InvalidHyperlinkLocation"):
            ed.add_internal_hyperlink(0, "A9", "")
        with pytest.raises(zlsx.ZlsxError, match="CommentRefTaken"):
            ed.add_comment(0, "A1", "x", "y")
        with pytest.raises(zlsx.ZlsxError, match="InvalidCommentRef"):
            ed.add_comment(0, "A1:A2", "x", "y")
        with pytest.raises(zlsx.ZlsxError, match="InvalidDataValidation"):
            ed.add_data_validation_list(0, "A9", ["a,b"])
        with pytest.raises(ValueError, match="unknown data validation kind"):
            ed.add_data_validation_numeric(0, "A9", "money", "between", "1", "2")
        with pytest.raises(zlsx.ZlsxError, match="InvalidDataValidation"):
            ed.add_data_validation_numeric(0, "A9", "whole", "between", "1", None)
        with pytest.raises(ValueError, match="unknown conditional-format operator"):
            ed.add_conditional_format_cell_is(0, "A9", "near", "1", None, 0)
        with pytest.raises(zlsx.ZlsxError, match="UnknownDxfId"):
            ed.add_conditional_format_cell_is(0, "A9", "equal", "1", None, 7)
        with pytest.raises(zlsx.ZlsxError, match="InvalidHyperlinkRange"):
            ed.add_conditional_format_data_bar(0, "9A", 0)
        # A control byte is judged at the registration, never at the save.
        with pytest.raises(zlsx.ZlsxError, match="InvalidXmlByte"):
            ed.add_comment(0, "D9", "b\x01d", "t")
        with pytest.raises(zlsx.ZlsxError, match="InvalidXmlByte"):
            ed.add_hyperlink(0, "D9", "https://x/\x02")
        with pytest.raises(zlsx.ZlsxError, match="InvalidXmlByte"):
            ed.add_data_validation_list(0, "D9", ["a\x1f"])
        with pytest.raises(ValueError):
            ed.add_merged_cell(2**32, "A1:B1")
        with pytest.raises(TypeError):
            ed.add_merged_cell(0, 42)
        ed.save(out)
    assert out.read_bytes() == before


def test_a_staged_registration_excludes_the_structural_edits_and_rides_beside_appends(tmp_path: Path) -> None:
    _needs_attachments()
    src = tmp_path / "src.xlsx"
    out = tmp_path / "out.xlsx"
    _write_fixture(src)
    with zlsx.edit(src) as ed:
        ed.add_merged_cell(1, "B2:C2")
        with pytest.raises(zlsx.ZlsxError, match="RowEditRequiresCleanSheet"):
            ed.insert_row(1, 1)
        with pytest.raises(zlsx.ZlsxError, match="ColEditRequiresCleanSheet"):
            ed.delete_column(1, 1)
        with pytest.raises(zlsx.ZlsxError, match="SheetDeleteRequiresCleanState"):
            ed.delete_sheet(1)
        ed.insert_row(0, 9)
        ed.append_rows(1, [[6]])
        ed.set_row_height(1, 1, 21)
        ed.save(out)
    sheet2 = _part(out, "xl/worksheets/sheet2.xml")
    assert '<row r="2" ht="21" customHeight="1"><c r="A2"><v>6</v></c></row></sheetData><mergeCells count="1"><mergeCell ref="B2:C2"/></mergeCells>' in sheet2


def test_a_sheet_part_the_splice_cannot_extend_is_a_refusal_at_the_first_registration(tmp_path: Path) -> None:
    _needs_attachments()
    src = tmp_path / "src.xlsx"
    torn = tmp_path / "torn.xlsx"
    out = tmp_path / "out.xlsx"
    _write_fixture(src)
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(torn, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/worksheets/sheet2.xml":
                data = (
                    b'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                    b'<sheetData/><x:mergeCells xmlns:x="u"/></worksheet>'
                )
            zout.writestr(item, data)
    before = torn.read_bytes()
    with zlsx.edit(torn) as ed:
        for call in (
            lambda: ed.set_column_width(1, 0, 9),
            lambda: ed.add_merged_cell(1, "A1:A2"),
            lambda: ed.add_comment(1, "A1", "x", "y"),
            lambda: ed.add_hyperlink(1, "A1", "https://x/"),
        ):
            with pytest.raises(zlsx.ZlsxRefusal) as info:
                call()
            assert info.value.error_name == "MalformedSheetXml"
        # The other sheet is readable.
        ed.add_merged_cell(0, "A5:B5")
        ed.save(out)
    assert out.read_bytes() != before
    assert '<mergeCell ref="A5:B5"/>' in _part(out, "xl/worksheets/sheet1.xml")
    with zlsx.edit(torn) as ed:
        with pytest.raises(zlsx.ZlsxRefusal):
            ed.add_merged_cell(1, "A1:A2")
        ed.save(out)
    assert out.read_bytes() == before


def test_save_with_recalc_carries_the_registrations(tmp_path: Path) -> None:
    _needs_attachments()
    src = tmp_path / "src.xlsx"
    with zlsx.Writer(src) as w:
        s = w.add_sheet("S")
        s.write_row_with_formulas([1, None], [None, "A1+1"])
    folded = tmp_path / "folded.xlsx"
    ordered = tmp_path / "ordered.xlsx"
    with zlsx.edit(src) as ed:
        ed.add_merged_cell(0, "C1:D1")
        ed.add_comment(0, "A1", "me", "n")
        ed.set_cell(0, 1, 0, 5)
        ed.save_with_recalc(folded)
    with zlsx.edit(src) as ed:
        ed.add_merged_cell(0, "C1:D1")
        ed.add_comment(0, "A1", "me", "n")
        ed.set_cell(0, 1, 0, 5)
        ed.recalculate()
        ed.save(ordered)
    assert folded.read_bytes() == ordered.read_bytes()
    with zlsx.open(folded) as book:
        assert len(book.merged_ranges(0)) == 1
        assert len(book.comments(0)) == 1
        sheet = book.sheet(0)
        header, rows = sheet.read_all()
        assert rows[0][1] == 6
