"""S3d slice 1 — styles on an opened workbook, through the editor handle.

``Editor.add_style`` / ``add_dxf`` / ``intern_num_fmt`` are
``Workbook.addStyle`` / ``addDxf`` / ``internNumFmt`` through
``zlsx_editor_add_style`` / ``_add_dxf`` / ``_intern_num_fmt``, and
``Editor.set_cell_style`` is ``Worksheet.setCellStyle`` through
``zlsx_editor_set_cell_style``: the fresh writer's registrations on an
OPENED workbook, the part extended at save. Every workbook here comes
from ``zlsx.Writer`` — two styles, one dxf, one custom format — so the
part the editor extends holds three ``<xf>``, one ``<dxf>`` and
``numFmtId`` 164 before the first registration.
"""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest

import zlsx
from zlsx import BorderSide, Dxf, Style


def _needs_editor_styles():
    import zlsx._ffi as ffi

    if not ffi._HAS_EDITOR_STYLES:
        pytest.skip("loaded libzlsx predates zlsx_editor_add_style (0.9.0+)")
    return ffi


def _write_fixture(path: Path) -> None:
    """Sheet ``S``: a bold header, a plain number, an italic ``0.00``
    cell; sheet ``T``: one number. A dxf registered for the table."""
    with zlsx.Writer(path) as w:
        bold = w.add_style(Style(font_bold=True))
        italic = w.add_style(Style(font_italic=True, number_format="0.00"))
        w.add_dxf(Dxf(font_bold=True))
        s = w.add_sheet("S")
        s.write_row(["h"], styles=[bold])
        s.write_row([1], styles=[0])
        s.write_row(["x"], styles=[italic])
        t = w.add_sheet("T")
        t.write_row([5])


def _styles_xml(path: Path) -> str:
    with zipfile.ZipFile(path) as z:
        return z.read("xl/styles.xml").decode("utf-8")


def _sheet_xml(path: Path, n: int) -> str:
    with zipfile.ZipFile(path) as z:
        return z.read(f"xl/worksheets/sheet{n}.xml").decode("utf-8")


HIGHLIGHT = Style(
    font_bold=True,
    fill_pattern="solid",
    fill_fg_argb=0xFFFFFF00,
    border_top=BorderSide(style="thin"),
    number_format="yyyy-mm-dd",
)


def test_registrations_extend_the_part_and_the_cells_take_the_index(tmp_path):
    _needs_editor_styles()
    src = tmp_path / "src.xlsx"
    dst = tmp_path / "dst.xlsx"
    _write_fixture(src)
    before = _styles_xml(src)
    assert '<cellXfs count="3">' in before

    with zlsx.Editor(src) as ed:
        idx = ed.add_style(HIGHLIGHT)
        assert idx == 3
        # Dedup within the save, against this editor's registrations.
        assert ed.add_style(HIGHLIGHT) == 3
        assert ed.add_style(Style(font_italic=True)) == 4
        assert ed.add_dxf(Dxf(font_italic=True)) == 1
        assert ed.add_dxf(Dxf(font_italic=True)) == 1
        # The style's format took 165; "0.00" is in the part as 164 and
        # lands again — the pool dedups within the save only.
        assert ed.intern_num_fmt("yyyy-mm-dd") == 165
        assert ed.intern_num_fmt("0%") == 166
        ed.set_cell_style(0, 2, 0, idx)  # an unstyled cell of the part
        ed.set_cell_style(0, 1, 0, 4)  # a styled cell: replaced
        ed.set_cell(0, 1, 1, 2)
        ed.set_cell_style(0, 1, 1, idx)  # a staged value
        ed.set_cell_style(0, 7, 2, idx)  # no such cell: created empty
        ed.set_cell_style(0, 9, 0, 0)  # a slot the part holds
        ed.save(dst)
        # Drained: the next registration reads the extended part.
        assert ed.add_style(Style(wrap_text=True)) == 5
        assert ed.add_dxf(Dxf(font_size=9.0)) == 2
        assert ed.intern_num_fmt("0.0") == 167

    after = _styles_xml(dst)
    # Everything before the first table and after the last one is the
    # writer's; the tables carry the new totals.
    assert before[: before.index("<numFmts")] == after[: after.index("<numFmts")]
    assert before[before.index("</dxfs>") :] == after[after.index("</dxfs>") :]
    for needle in (
        '<numFmts count="3">',
        '<numFmt numFmtId="165" formatCode="yyyy-mm-dd"/><numFmt numFmtId="166" formatCode="0%"/>',
        '<fonts count="5">',
        '<fills count="3">',
        '<borders count="2">',
        '<cellXfs count="5">',
        '<xf numFmtId="165" fontId="3" fillId="2" borderId="1" xfId="0" applyFont="1" applyNumberFormat="1" applyFill="1" applyBorder="1"/>',
        '<dxfs count="2">',
        "<dxf><font><i/></font></dxf></dxfs>",
    ):
        assert needle in after, needle
    sheet = _sheet_xml(dst, 1)
    assert '<c r="A2" s="3">' in sheet
    assert '<c r="A1" s="4"' in sheet
    assert '<c r="B1" s="3"><v>2</v></c>' in sheet
    assert '<c r="C7" s="3"/>' in sheet
    assert '<c r="A9" s="0"/>' in sheet
    # A sheet with no work is untouched.
    assert _sheet_xml(src, 2) == _sheet_xml(dst, 2)
    # The reader resolves the records as any producer's.
    with zlsx.open(dst) as book:
        rows = book.sheet(0).rows()
        next(iter(rows))
        assert rows.style_indices()[:2] == [4, 3]
        assert book.cell_font(3).bold
        assert book.cell_font(4).italic
        assert book.number_format(3) == "yyyy-mm-dd"


def test_statements_about_the_call_stage_nothing(tmp_path):
    _needs_editor_styles()
    src = tmp_path / "src.xlsx"
    dst = tmp_path / "dst.xlsx"
    _write_fixture(src)
    with zlsx.Editor(src) as ed:
        with pytest.raises(zlsx.ZlsxError, match="UnknownStyleIndex"):
            ed.set_cell_style(0, 1, 0, 3)
        with pytest.raises(zlsx.ZlsxError, match="SheetIndexOutOfRange"):
            ed.set_cell_style(2, 1, 0, 0)
        with pytest.raises(zlsx.ZlsxError, match="RowIndexOutOfRange"):
            ed.set_cell_style(0, 0, 0, 0)
        with pytest.raises(zlsx.ZlsxError, match="ColumnIndexOutOfRange"):
            ed.set_cell_style(0, 1, 16384, 0)
        # Bounded before ctypes narrows them: never wrapped to another
        # sheet, row or style (r6 A-PY-601).
        with pytest.raises(ValueError):
            ed.set_cell_style(2**32, 1, 0, 0)
        with pytest.raises(ValueError):
            ed.set_cell_style(0, 1, 0, 2**32 + 1)
        with pytest.raises(TypeError):
            ed.set_cell_style(0, 1.9, 0, 0)
        with pytest.raises(TypeError):
            ed.set_cell_style(0, 1, True, 0)
        # The same guard on the two writers that shared the gap (r6/r7).
        with pytest.raises(ValueError):
            ed.append_rows(2**32, [[9]])
        # A dxf keeps add_style's font-size rule (r27 A-DXF-2701, r28
        # A-PIN-2807): non-finite or non-positive is InvalidStyle.
        for bad in (float("nan"), float("inf"), -3.0, 0.0):
            with pytest.raises(zlsx.ZlsxError, match="InvalidStyle"):
                ed.add_dxf(Dxf(font_size=bad))
        with pytest.raises(ValueError):
            ed.set_cell(2**32, 1, 0, 1)
        with pytest.raises(zlsx.ZlsxError, match="InvalidStyle"):
            ed.intern_num_fmt("")
        with pytest.raises(zlsx.ZlsxError, match="InvalidStyle"):
            ed.add_style(Style(font_size=0.0))
        with pytest.raises(zlsx.ZlsxError, match="InvalidFontName"):
            ed.add_style(Style(font_name=""))
        with pytest.raises(ValueError):
            ed.add_style(Style(fill_pattern="plaid"))
        ed.append_rows(1, [[9]])
        with pytest.raises(zlsx.ZlsxError, match="SheetHasUnsavedAppends"):
            ed.set_cell_style(1, 1, 0, 0)
    # Untouched otherwise: the passthrough.
    with zlsx.Editor(src) as ed:
        with pytest.raises(zlsx.ZlsxError, match="UnknownStyleIndex"):
            ed.set_cell_style(0, 1, 0, 3)
        ed.save(dst)
    assert dst.read_bytes() == src.read_bytes()
    # A staged style is a staged cell write to the structural edits.
    with zlsx.Editor(src) as ed:
        ed.set_cell_style(0, 1, 0, 1)
        with pytest.raises(zlsx.ZlsxError, match="RowEditRequiresCleanSheet|SheetHasUnsavedMutations"):
            ed.insert_row(0, 1)
        with pytest.raises(zlsx.ZlsxError, match="SheetHasUnsavedMutations"):
            ed.append_rows(0, [[9]])
    # …and the fresh writer keeps the same rule under its own name (r28
    # A-DOC-2804, pinned r29 B-PIN-2901).
    with zlsx.Writer(tmp_path / "w.xlsx") as w:
        for bad in (float("nan"), float("inf"), -3.0, 0.0):
            with pytest.raises(zlsx.ZlsxError, match="InvalidFontSize"):
                w.add_dxf(Dxf(font_size=bad))
        w.add_sheet("S").write_row([1])


def test_a_styles_part_the_extension_cannot_read_refuses_before_anything_is_staged(tmp_path):
    _needs_editor_styles()
    src = tmp_path / "src.xlsx"
    torn = tmp_path / "torn.xlsx"
    dst = tmp_path / "dst.xlsx"
    _write_fixture(src)
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(torn, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/styles.xml":
                data = (
                    b'<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                    b'<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
                    b'<fonts count="1"><font/></fonts></styleSheet>'
                )
            zout.writestr(item, data)
    with zlsx.Editor(torn) as ed:
        for op in (
            lambda: ed.add_style(Style(font_bold=True)),
            lambda: ed.add_dxf(Dxf(font_bold=True)),
            lambda: ed.intern_num_fmt("0"),
            lambda: ed.set_cell_style(0, 1, 0, 0),
        ):
            with pytest.raises(zlsx.ZlsxRefusal) as info:
                op()
            assert info.value.error_name == "MalformedStylesXml"
            assert not isinstance(info.value, zlsx.ZlsxFormulaRefusal)
        ed.save(dst)
    assert dst.read_bytes() == torn.read_bytes()


def test_the_buffer_and_recalc_saves_carry_the_registrations(tmp_path):
    _needs_editor_styles()
    src = tmp_path / "src.xlsx"
    dst = tmp_path / "dst.xlsx"
    _write_fixture(src)
    with zlsx.Editor(src) as ed:
        idx = ed.add_style(Style(font_bold=True, font_color_argb=0xFFFF0000))
        ed.set_cell_style(1, 1, 0, idx)
        blob = ed.save_to_buffer()
    with zlsx.Editor.from_bytes(blob) as ed2:
        # The buffer's part is the extended one: the next slot follows.
        assert ed2.add_style(Style(font_italic=True)) == 4
    with zlsx.Editor(src) as ed:
        idx = ed.add_style(Style(font_bold=True, font_color_argb=0xFFFF0000))
        ed.set_cell_style(1, 1, 0, idx)
        ed.save_with_recalc(dst)
    assert '<cellXfs count="4">' in _styles_xml(dst)
    assert '<c r="A1" s="3"><v>5</v></c>' in _sheet_xml(dst, 2)
