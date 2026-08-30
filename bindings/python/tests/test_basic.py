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
    import zlsx
    import zlsx._ffi as ffi

    lib_version = ffi.lib.zlsx_version_string().decode("utf-8")
    # Package version tracks the library's major.minor; patch may
    # drift. Derive the expected prefix from zlsx.__version__ so this
    # test stays in sync with version bumps automatically rather than
    # going stale every release.
    expected_prefix = ".".join(zlsx.__version__.split(".")[:2]) + "."
    assert lib_version.startswith(expected_prefix), (
        f"library version {lib_version!r} does not match package "
        f"major.minor prefix {expected_prefix!r}"
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


def test_book_methods_after_close_raise():
    """Calling Book methods after close() must raise ZlsxError, not
    segfault. Regression for the NULL-handle crash path."""
    path = _skip_if_missing("frictionless_2sheets.xlsx")

    book = zlsx.open(path)
    book.close()

    with pytest.raises(zlsx.ZlsxError):
        book.sheet(0)
    with pytest.raises(zlsx.ZlsxError):
        book.sheet("Sheet1")


def test_sheet_methods_after_book_close_raise():
    """Sheet handles outlive close() but their methods that re-enter
    the C ABI through book._handle must surface a clean error."""
    path = _skip_if_missing("frictionless_2sheets.xlsx")

    book = zlsx.open(path)
    sheet = book.sheet(0)
    book.close()

    with pytest.raises(zlsx.ZlsxError):
        sheet.rows()
    with pytest.raises(zlsx.ZlsxError):
        sheet.read_all()


def test_rows_methods_after_close_raise():
    """Rows.close() drops the C handle. Subsequent next/style/parse
    calls must raise ZlsxError rather than crash."""
    path = _skip_if_missing("frictionless_2sheets.xlsx")

    with zlsx.open(path) as book:
        rows = book.sheet(0).rows()
        next(rows)  # populate _current_len for style_indices
        rows.close()

        with pytest.raises(zlsx.ZlsxError):
            next(rows)


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


def test_open_bytes_round_trip(tmp_path):
    out = tmp_path / "mem.xlsx"
    with zlsx.write(out) as w:
        sheet = w.add_sheet("Data")
        sheet.write_row(["k", "v"])
        sheet.write_row(["alpha", 42])
        sheet.write_row(["beta", 2.5])
        second = w.add_sheet("Second")
        second.write_row([True])

    data = out.read_bytes()
    # Delete the file before opening from bytes: proves no path is
    # touched and (post-open) that the buffer was only borrowed for
    # the duration of the call.
    out.unlink()

    book = zlsx.open_bytes(data)
    del data  # buffer contract: not needed after open_bytes returns
    with book:
        assert book.sheets == ["Data", "Second"]
        rows = list(book.sheet("Data").rows())
        assert rows[0] == ["k", "v"]
        assert rows[1] == ["alpha", 42]
        assert rows[2][0] == "beta"
        assert abs(rows[2][1] - 2.5) < 1e-9
        assert list(book.sheet("Second").rows()) == [[True]]


def test_open_bytes_accepts_bytearray_and_memoryview(tmp_path):
    out = tmp_path / "mem2.xlsx"
    with zlsx.write(out) as w:
        w.add_sheet("S").write_row(["x"])
    raw = out.read_bytes()

    for form in (bytearray(raw), memoryview(raw)):
        with zlsx.open_bytes(form) as book:
            assert list(book.sheet(0).rows()) == [["x"]]


def test_open_bytes_garbage_raises():
    with pytest.raises(zlsx.ZlsxError, match="BadZip"):
        zlsx.open_bytes(b"definitely not a zip archive")
    with pytest.raises(zlsx.ZlsxError):
        zlsx.open_bytes(b"")


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


# ─── Rows.skip ────────────────────────────────────────────────────────


@pytest.fixture
def skip_book(tmp_path):
    out = tmp_path / "skip.xlsx"
    with zlsx.write(out) as w:
        s = w.add_sheet("S")
        s.write_row(["name", "n"])
        for i in range(20):
            s.write_row([f"r{i}", i])
    return out


def test_rows_skip_lands_where_next_would(skip_book):
    """The property every range-partitioned read rests on: skip(k) then
    next() is the (k+1)-th row of a plain walk."""
    with zlsx.open(skip_book) as book:
        with book.sheet(0).rows() as rows:
            walked = [list(r) for r in rows]

    for k in range(len(walked)):
        with zlsx.open(skip_book) as book:
            with book.sheet(0).rows() as rows:
                assert rows.skip(k) == k
                assert list(next(iter(rows))) == walked[k]


def test_rows_skip_then_drain_yields_the_remainder(skip_book):
    with zlsx.open(skip_book) as book:
        with book.sheet(0).rows() as rows:
            rows.skip(5)
            rest = [list(r) for r in rows]
    assert len(rest) == 16          # 21 rows total, 5 skipped
    assert rest[0][0] == "r4"


def test_rows_skip_past_end_reports_actual_count(skip_book):
    with zlsx.open(skip_book) as book:
        with book.sheet(0).rows() as rows:
            assert rows.skip(10_000) == 21
            assert list(rows) == []


def test_rows_skip_zero_and_negative(skip_book):
    with zlsx.open(skip_book) as book:
        with book.sheet(0).rows() as rows:
            assert rows.skip(0) == 0
            assert list(next(iter(rows))) == ["name", "n"]
            with pytest.raises(ValueError, match="non-negative"):
                rows.skip(-1)


def test_rows_skip_after_close_raises(skip_book):
    with zlsx.open(skip_book) as book:
        rows = book.sheet(0).rows()
        rows.close()
        with pytest.raises(zlsx.ZlsxError, match="closed"):
            rows.skip(1)


def test_rows_skip_interleaves_with_next(skip_book):
    with zlsx.open(skip_book) as book:
        with book.sheet(0).rows() as rows:
            it = iter(rows)
            next(it)                      # header
            rows.skip(3)                  # r0..r2
            assert list(next(it))[0] == "r3"
            rows.skip(2)                  # r4, r5
            assert list(next(it))[0] == "r6"


def test_rows_skip_over_formula_cells(tmp_path):
    """Skipping across rows that carry formulas still lands correctly.

    Note this exercises the fast scan path, not the fallback: the zlsx
    writer emits plain ``<f>`` and never ``t="shared"`` / ``t="array"``,
    so no workbook it produces has cross-row formula state. The
    fallback — a shared base in a skipped row still resolving for its
    slave — is covered in Zig against raw sheet XML
    ("skipRows falls back to decoding when the sheet has shared
    formulas" in src/xlsx.zig), which is the only place such a sheet
    can be constructed.
    """
    out = tmp_path / "formulas.xlsx"
    with zlsx.write(out) as w:
        s = w.add_sheet("S")
        s.write_row(["a"])
        s.write_row_with_formulas([None], ["B1*2"])
        s.write_row_with_formulas([None], ["B2*2"])
        s.write_row(["last"])

    with zlsx.open(out) as book:
        with book.sheet(0).rows() as rows:
            assert rows.skip(3) == 3
            assert list(next(iter(rows)))[0] == "last"


# ─── Writer.to_bytes (writer-side mirror of open_bytes) ───────────────


def _sample_workbook(w):
    bold = w.add_style(zlsx.Style(font_bold=True))
    s = w.add_sheet("Summary")
    s.write_row(["region", "units"], styles=[bold, bold])
    s.write_row(["North", 120])
    s.write_row(["North", 7.5])       # SST dedup hit
    w.add_sheet("Notes").write_row(["second sheet"])


def test_writer_to_bytes_matches_save(tmp_path):
    out = tmp_path / "parity.xlsx"
    with zlsx.write() as w:
        _sample_workbook(w)
        payload = w.to_bytes()
        w.save(out)

    assert payload == out.read_bytes(), "to_bytes must equal what save writes"


def test_writer_to_bytes_round_trips_through_open_bytes():
    with zlsx.write() as w:
        _sample_workbook(w)
        payload = w.to_bytes()

    assert isinstance(payload, bytes)
    assert payload[:2] == b"PK"
    with zlsx.open_bytes(payload) as book:
        assert book.sheets == ["Summary", "Notes"]
        header, rows = book.sheet(0).read_all(header=True)
        assert header == ["region", "units"]
        assert rows == [["North", 120], ["North", 7.5]]


def test_writer_to_bytes_is_repeatable_and_non_consuming():
    with zlsx.write() as w:
        _sample_workbook(w)
        first = w.to_bytes()
        second = w.to_bytes()
        assert first == second
        # Still usable afterwards: appending changes the next payload.
        w.add_sheet("Third").write_row(["x"])
        assert w.to_bytes() != first


def test_writer_to_bytes_empty_workbook_raises():
    with zlsx.write() as w:
        with pytest.raises(zlsx.ZlsxError, match="NoSheets"):
            w.to_bytes()


def test_writer_to_bytes_after_close_raises():
    w = zlsx.write()
    w.add_sheet("S").write_row(["a"])
    w.close()
    with pytest.raises(zlsx.ZlsxError, match="closed"):
        w.to_bytes()


def test_writer_to_bytes_survives_many_calls():
    """The buffer is freed through zlsx_buffer_free on every call; a leak
    here would be invisible in a single round-trip, so loop it."""
    with zlsx.write() as w:
        _sample_workbook(w)
        sizes = {len(w.to_bytes()) for _ in range(200)}
    assert len(sizes) == 1


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


def test_conditional_formatting_color_scale_data_bar(tmp_path):
    """iter51: colorScale (2-stop + 3-stop) and dataBar CF rules."""
    import zipfile
    import zlsx._ffi as ffi

    if not ffi._HAS_CF_GRADIENT:
        pytest.skip("loaded libzlsx predates colorScale/dataBar ABI (0.2.6+)")

    out = tmp_path / "cf_gradient.xlsx"
    with zlsx.write(out) as w:
        sheet = w.add_sheet("S")
        sheet.add_conditional_format_color_scale(
            "A2:A100", 0xFFFF0000, 0xFFFFFF00, 0xFF00FF00,
        )
        sheet.add_conditional_format_color_scale(
            "B2:B100", 0xFFFFFFFF, None, 0xFF0000FF,
        )
        sheet.add_conditional_format_data_bar("C2:C100", 0xFF638EC6)
        sheet.write_row(["hdr"])

        with pytest.raises(zlsx.ZlsxError, match="InvalidHyperlinkRange"):
            sheet.add_conditional_format_color_scale("", 0, None, 0)
        with pytest.raises(zlsx.ZlsxError, match="InvalidHyperlinkRange"):
            sheet.add_conditional_format_data_bar("", 0)

    with zipfile.ZipFile(out) as z:
        sheet_xml = z.read("xl/worksheets/sheet1.xml").decode("utf-8")

    assert '<cfRule type="colorScale" priority="1">' in sheet_xml
    assert '<cfvo type="percentile" val="50"/>' in sheet_xml  # 3-stop
    assert '<color rgb="FFFF0000"/>' in sheet_xml
    assert '<color rgb="FFFFFF00"/>' in sheet_xml
    assert '<color rgb="FF00FF00"/>' in sheet_xml

    assert '<cfRule type="colorScale" priority="2">' in sheet_xml
    # 2-stop skips percentile cfvo; total count = 1 across both rules.
    assert sheet_xml.count('<cfvo type="percentile"') == 1

    assert '<cfRule type="dataBar" priority="3">' in sheet_xml
    assert '<color rgb="FF638EC6"/>' in sheet_xml


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


def test_book_comments_rich_text_runs_surface(tmp_path):
    """iter53 — rich-text comment bodies populate `Comment.runs`.
    Plain-text comments keep `runs=None` (zero-overhead path)."""
    import zipfile
    import zlsx._ffi as ffi

    if not ffi._HAS_COMMENTS or not ffi._HAS_COMMENT_RUNS:
        pytest.skip("loaded libzlsx predates comment-runs ABI (0.2.6+)")

    xlsx_path = tmp_path / "comments_rich.xlsx"
    sst_xml = (
        b"<?xml version=\"1.0\"?>"
        b"<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"0\"/>"
    )
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
        b"<sheetData><row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>x</t></is></c></row></sheetData></worksheet>"
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
        b"<authors><author>Alice</author></authors>"
        b"<commentList>"
        b"<comment ref=\"A1\" authorId=\"0\"><text><t>plain body</t></text></comment>"
        b"<comment ref=\"B2\" authorId=\"0\"><text><r><rPr><b/></rPr><t>bold </t></r><r><rPr><i/></rPr><t>italic</t></r></text></comment>"
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

        # Plain comment: runs=None, text populated.
        assert cs[0].text == "plain body"
        assert cs[0].runs is None

        # Rich comment: text is concat, runs tuple is populated.
        assert cs[1].text == "bold italic"
        assert cs[1].runs is not None
        assert len(cs[1].runs) == 2
        assert cs[1].runs[0].text == "bold "
        assert cs[1].runs[0].bold and not cs[1].runs[0].italic
        assert cs[1].runs[1].text == "italic"
        assert cs[1].runs[1].italic and not cs[1].runs[1].bold


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


def test_writer_add_defined_name_after_close_raises(tmp_path):
    """Calling add_defined_name on a closed Writer must raise rather
    than segfault into the C ABI with a null handle."""
    import zlsx._ffi as ffi
    if not ffi._HAS_DEFINED_NAME:
        pytest.skip("loaded libzlsx predates add_defined_name ABI (post-0.3.0)")
    out = str(tmp_path / "closed.xlsx")
    w = zlsx.Writer(out)
    w.close()
    with pytest.raises(zlsx.ZlsxError, match="Writer is closed"):
        w.add_defined_name("Foo", "Sheet1!$A$1")


def test_writer_add_defined_name_rejects_huge_local_sheet_id(tmp_path):
    """local_sheet_id > INT32_MAX would silently wrap through ctypes
    to a negative value, which the C ABI treats as workbook scope.
    The Python wrapper rejects up-front."""
    import zlsx._ffi as ffi
    if not ffi._HAS_DEFINED_NAME:
        pytest.skip("loaded libzlsx predates add_defined_name ABI (post-0.3.0)")
    out = str(tmp_path / "huge_lsi.xlsx")
    with zlsx.Writer(out) as w:
        w.add_sheet("S")
        with pytest.raises(ValueError, match="INT32_MAX"):
            w.add_defined_name("Foo", "S!$A$1", local_sheet_id=2**31)
        with pytest.raises(ValueError, match=">= 0"):
            w.add_defined_name("Foo", "S!$A$1", local_sheet_id=-1)


def test_writer_add_defined_name_round_trip(tmp_path):
    """Workbook + sheet-scoped defined names ship through to xl/workbook.xml
    and round-trip through the same Writer's save path."""
    import zlsx._ffi as ffi
    if not ffi._HAS_DEFINED_NAME:
        pytest.skip("loaded libzlsx predates add_defined_name ABI (post-0.3.0)")
    out = str(tmp_path / "defined_names.xlsx")
    with zlsx.Writer(out) as w:
        sheet = w.add_sheet("Sheet1")
        sheet.write_row([1, 2, 3])
        # Workbook scope.
        w.add_defined_name("MyRange", "Sheet1!$A$1:$C$1")
        # Sheet-scope + hidden.
        w.add_defined_name(
            "_xlnm.Print_Area",
            "Sheet1!$A$1:$C$1",
            local_sheet_id=0,
            hidden=True,
        )

        # Validation surface.
        with pytest.raises(zlsx.ZlsxError, match="InvalidDefinedName"):
            w.add_defined_name("A1", "Sheet1!$A$1")
        with pytest.raises(zlsx.ZlsxError, match="InvalidDefinedNameRefersTo"):
            w.add_defined_name("Foo", "")
        # Case-insensitive duplicate within scope.
        with pytest.raises(zlsx.ZlsxError, match="DuplicateDefinedName"):
            w.add_defined_name("myrange", "Sheet1!$B$1")


def test_sheet_writer_set_row_height_validates(tmp_path):
    """set_row_height accepts (0, 409.5] and rejects everything else."""
    import zlsx._ffi as ffi
    if not ffi._HAS_SET_ROW_HEIGHT:
        pytest.skip("loaded libzlsx predates set_row_height ABI (post-0.3.0)")
    out = str(tmp_path / "row_height.xlsx")
    with zlsx.Writer(out) as w:
        sheet = w.add_sheet("Sheet1")
        sheet.write_row(["x"])
        # Valid.
        sheet.set_row_height(0, 24.0)
        sheet.set_row_height(1, 409.5)  # at the cap
        # Invalid: zero, negative, above cap.
        with pytest.raises(zlsx.ZlsxError, match="InvalidRowHeight"):
            sheet.set_row_height(2, 0.0)
        with pytest.raises(zlsx.ZlsxError, match="InvalidRowHeight"):
            sheet.set_row_height(2, -1.0)
        with pytest.raises(zlsx.ZlsxError, match="InvalidRowHeight"):
            sheet.set_row_height(2, 410.0)


def test_sheet_writer_set_row_height_rejects_huge_row_idx(tmp_path):
    """row_idx > UINT32_MAX would wrap through ctypes; reject before FFI."""
    import zlsx._ffi as ffi
    if not ffi._HAS_SET_ROW_HEIGHT:
        pytest.skip("loaded libzlsx predates set_row_height ABI (post-0.3.0)")
    out = str(tmp_path / "huge_rowidx.xlsx")
    with zlsx.Writer(out) as w:
        sheet = w.add_sheet("S")
        sheet.write_row(["x"])
        with pytest.raises(ValueError, match="UINT32_MAX"):
            sheet.set_row_height(2**32 + 1, 24.0)


def test_sheet_writer_freeze_panes_checked_rejects_huge_counts(tmp_path):
    """rows/cols > UINT32_MAX would wrap through ctypes; reject before FFI."""
    import zlsx._ffi as ffi
    if not ffi._HAS_FREEZE_PANES_CHECKED:
        pytest.skip("loaded libzlsx predates freeze_panes_checked ABI (post-0.3.0)")
    out = str(tmp_path / "huge_freeze.xlsx")
    with zlsx.Writer(out) as w:
        sheet = w.add_sheet("S")
        sheet.write_row(["x"])
        with pytest.raises(ValueError, match="UINT32_MAX"):
            sheet.freeze_panes_checked(rows=2**32, cols=0)
        with pytest.raises(ValueError, match="UINT32_MAX"):
            sheet.freeze_panes_checked(rows=0, cols=2**32)


def test_sheet_writer_freeze_panes_checked_propagates_errors(tmp_path):
    """The checked variant raises ZlsxError on out-of-range counts
    instead of clamping silently."""
    import zlsx._ffi as ffi
    if not ffi._HAS_FREEZE_PANES_CHECKED:
        pytest.skip("loaded libzlsx predates freeze_panes_checked ABI (post-0.3.0)")
    out = str(tmp_path / "freeze_checked.xlsx")
    with zlsx.Writer(out) as w:
        sheet = w.add_sheet("Sheet1")
        sheet.write_row(["x"])
        # Valid.
        sheet.freeze_panes_checked(rows=1, cols=1)
        # Out-of-range.
        with pytest.raises(zlsx.ZlsxError, match="RowOutOfRange"):
            sheet.freeze_panes_checked(rows=1_048_576, cols=0)
        with pytest.raises(zlsx.ZlsxError, match="ColumnOutOfRange"):
            sheet.freeze_panes_checked(rows=0, cols=16_384)


def test_book_cell_alignment_out_of_range_returns_none(tmp_path):
    """Out-of-range style_idx returns None, including values that
    would wrap through ctypes c_uint32."""
    import zlsx._ffi as ffi
    if not ffi._HAS_CELL_ALIGNMENT:
        pytest.skip("loaded libzlsx predates cell_alignment ABI (post-0.3.0)")
    out = str(tmp_path / "align_oor.xlsx")
    with zlsx.Writer(out) as w:
        w.add_sheet("S").write_row(["x"])

    with zlsx.open(out) as book:
        # Past the end of the cell-xfs table.
        assert book.cell_alignment(99999) is None
        # Negative.
        assert book.cell_alignment(-1) is None
        # Past UINT32_MAX (would wrap to 0 and return the default).
        assert book.cell_alignment(2**32) is None
        assert book.cell_alignment(2**40) is None


def test_book_cell_alignment_round_trip(tmp_path):
    """A writer-emitted style with horizontal=center + wrap_text=True
    reads back through Book.cell_alignment."""
    import zlsx._ffi as ffi
    if not ffi._HAS_CELL_ALIGNMENT:
        pytest.skip("loaded libzlsx predates cell_alignment ABI (post-0.3.0)")
    out = str(tmp_path / "alignment.xlsx")
    with zlsx.Writer(out) as w:
        sheet = w.add_sheet("Sheet1")
        sid = w.add_style(zlsx.Style(
            alignment_horizontal="center",
            wrap_text=True,
        ))
        sheet.write_row(["x"], styles=[sid])

    with zlsx.open(out) as book:
        align = book.cell_alignment(sid)
        assert align is not None
        assert align.horizontal == "center"
        assert align.wrap_text is True

        # Index 0 (default no-style) must surface as horizontal="",
        # wrap_text=False — the OOXML default.
        align0 = book.cell_alignment(0)
        assert align0 is not None
        assert align0.horizontal == ""
        assert align0.wrap_text is False


# ─── Editor.set_cell + save round-trip (iter-cm-2) ─────────────────────
#
# The C ABI exposes the Editor as `zlsx_editor_open` / `zlsx_editor_set_cell`
# / `zlsx_editor_save` (added in libzlsx 0.2.7 / 0.2.9). The user-facing
# request mentions a "Workbook overlay" with `Workbook.open / setCell /
# save` — that surface is **not** a separate ABI; it's the Editor under
# its production name. The Python wrapper exposes it as `zlsx.edit()`
# returning a `zlsx.Editor`, with `set_cell()` and `save()` methods that
# match the requested round-trip contract. These tests cover that
# Workbook-overlay-equivalent path.
#
# Note: there is no `formula` cell tag in the C ABI — formulas are a
# row-level construct via `zlsx_sheet_writer_write_row_with_formulas`,
# not a per-cell `set_cell` call. The "formula" item from the user's
# scope is therefore not addressable through `Editor.set_cell` and is
# omitted intentionally; if a formula-mutating editor surface ships in
# a later iter, a sibling test belongs here.


def test_editor_set_cell_round_trip_number_int_bool_blank_string(tmp_path):
    """Workbook.open(tmp_path) → setCell × {int, float, bool, blank,
    string} → save → re-open → assert all values round-trip through
    the C ABI."""
    import zlsx._ffi as ffi
    if not ffi._HAS_EDITOR or not ffi._HAS_EDITOR_SET_CELL:
        pytest.skip("workbook overlay (Editor.set_cell) not exposed in loaded libzlsx")

    src = tmp_path / "src.xlsx"
    out = tmp_path / "out.xlsx"

    # Seed a 2-row × 5-col grid of integers. Plain integer rows produce a
    # canonical body with no `s=` style attr, so the editor's
    # SetCellSourceCellHasMetadata guard is satisfied for every cell we
    # rewrite below.
    with zlsx.write(src) as w:
        s = w.add_sheet("S")
        s.write_row([1, 2, 3, 4, 5])
        s.write_row([6, 7, 8, 9, 10])

    with zlsx.edit(src) as ed:
        # Cover every cell tag the ABI accepts via Editor.set_cell.
        ed.set_cell(0, 1, 0, 42)              # CELL_INTEGER
        ed.set_cell(0, 1, 1, 3.14159)         # CELL_NUMBER
        ed.set_cell(0, 1, 2, True)            # CELL_BOOLEAN
        ed.set_cell(0, 1, 3, None)            # CELL_EMPTY (blank)
        ed.set_cell(0, 1, 4, "Done & dusted") # CELL_STRING (inline; XML-escaped)
        # Also exercise set_cells bulk variant on row 2.
        ed.set_cells(0, [
            (2, 0, -7),
            (2, 1, 2.5),
            (2, 2, False),
            (2, 3, None),
            (2, 4, " trim me "),
        ])
        ed.save(out)

    assert out.exists()

    with zlsx.open(out) as book:
        rows = list(book.sheet(0).rows())
        # set_cell on col 3 with None blanks the cell — the row width may
        # therefore shrink if the trailing cell ended up blank, but cols
        # 0..4 all have some non-trivial content here so the iterator
        # surfaces them.
        assert rows[0][0] == 42
        assert isinstance(rows[0][0], int)
        assert abs(rows[0][1] - 3.14159) < 1e-9
        assert isinstance(rows[0][1], float)
        assert rows[0][2] is True
        # Blank cell sits between two non-blank neighbours — readers
        # surface it as None (canonical blank) when it's interior.
        assert rows[0][3] is None
        assert rows[0][4] == "Done & dusted"

        assert rows[1][0] == -7
        assert isinstance(rows[1][0], int)
        assert abs(rows[1][1] - 2.5) < 1e-9
        assert isinstance(rows[1][1], float)
        assert rows[1][2] is False
        assert rows[1][3] is None
        assert rows[1][4] == " trim me "


def test_editor_close_releases_handle_and_methods_raise(tmp_path):
    """Editor.close releases the C handle; subsequent set_cell/save
    must raise ZlsxError, not segfault. Mirrors the Book/Rows
    after-close contract for lifetime safety."""
    import zlsx._ffi as ffi
    if not ffi._HAS_EDITOR or not ffi._HAS_EDITOR_SET_CELL:
        pytest.skip("workbook overlay (Editor.set_cell) not exposed in loaded libzlsx")

    src = tmp_path / "src.xlsx"
    out = tmp_path / "out.xlsx"

    with zlsx.write(src) as w:
        w.add_sheet("S").write_row([1, 2])

    ed = zlsx.edit(src)
    ed.set_cell(0, 1, 0, 99)
    ed.close()

    with pytest.raises(zlsx.ZlsxError):
        ed.set_cell(0, 1, 1, 100)
    with pytest.raises(zlsx.ZlsxError):
        ed.save(out)

    # Double-close is a no-op (idempotent) — must not segfault.
    ed.close()


def test_editor_context_manager_drops_handle_on_exception(tmp_path):
    """`with zlsx.edit(...)` must close the editor even when the body
    raises. Verifies set_cell on the post-exit handle raises
    ZlsxError instead of crashing."""
    import zlsx._ffi as ffi
    if not ffi._HAS_EDITOR or not ffi._HAS_EDITOR_SET_CELL:
        pytest.skip("workbook overlay (Editor.set_cell) not exposed in loaded libzlsx")

    src = tmp_path / "src.xlsx"
    with zlsx.write(src) as w:
        w.add_sheet("S").write_row([1])

    captured = {}
    with pytest.raises(RuntimeError, match="caller aborted"):
        with zlsx.edit(src) as ed:
            captured["ed"] = ed
            ed.set_cell(0, 1, 0, 7)
            raise RuntimeError("caller aborted")

    ed = captured["ed"]
    with pytest.raises(zlsx.ZlsxError):
        ed.set_cell(0, 1, 0, 8)


def test_editor_open_invalid_path_raises():
    """Editor.open on a non-existent path must surface ZlsxError, not
    return a NULL handle the caller would dereference."""
    import zlsx._ffi as ffi
    if not ffi._HAS_EDITOR:
        pytest.skip("workbook overlay (Editor) not exposed in loaded libzlsx")

    with pytest.raises(zlsx.ZlsxError):
        zlsx.edit("/nonexistent/path/does/not/exist.xlsx")


# ---- Document properties (Z3) ---------------------------------------


def _write_docprops_fixture(path: Path) -> None:
    """A workbook carrying known PII in docProps.

    The Writer emits no docProps parts, so this drives the CLI's
    scrub-metadata counterpart from the other side: build a plain
    workbook, then confirm the binding reports "no metadata" for it.
    Populated-metadata coverage lives in the Zig suite, which can
    inject parts via PartStore.addPart directly.
    """
    with zlsx.Writer() as w:
        sheet = w.add_sheet("Data")
        sheet.write_row(["name", 1])
        sheet.write_row(["keep-me", 2])
        w.save(path)


def test_editor_doc_props_reports_absent_metadata(tmp_path):
    """A Writer-produced workbook has no docProps at all — every field
    is None rather than an empty string, so callers can tell "absent"
    from "blank"."""
    from zlsx import _ffi as ffi

    if not ffi._HAS_DOCPROPS:
        pytest.skip("loaded libzlsx predates the docProps ABI (0.5.0+)")

    src = tmp_path / "nodocprops.xlsx"
    _write_docprops_fixture(src)

    ed = zlsx.Editor(src)
    props = ed.doc_props()

    assert props["creator"] is None
    assert props["last_modified_by"] is None
    assert props["company"] is None
    assert props["has_custom_properties"] is False
    # The full field set is always present, so consumers can index
    # without existence checks.
    for key in ("title", "subject", "description", "keywords", "category",
                "created", "modified", "revision", "manager",
                "application", "hyperlink_base"):
        assert key in props


def test_editor_strip_doc_props_is_a_noop_without_metadata(tmp_path):
    """Scrubbing a workbook that has no docProps must not error and
    must not invent the parts."""
    from zlsx import _ffi as ffi

    if not ffi._HAS_DOCPROPS:
        pytest.skip("loaded libzlsx predates the docProps ABI (0.5.0+)")

    src = tmp_path / "scrub_src.xlsx"
    dst = tmp_path / "scrub_dst.xlsx"
    _write_docprops_fixture(src)

    ed = zlsx.Editor(src)
    ed.strip_doc_props()
    ed.save(dst)

    assert dst.exists()
    # Cell data survives the metadata scrub untouched.
    _header, rows = zlsx.read(dst)
    assert rows[0][0] == "name"
    assert rows[1][0] == "keep-me"

    after = zlsx.Editor(dst).doc_props()
    assert after["creator"] is None
    assert after["has_custom_properties"] is False


# ─── Formula engine (M9a2) ────────────────────────────────────────────


def _write_formula_fixture(path, formula="A1+2", cached=0.0):
    """A1 = 1 (plain), B1 = formula with a stale cached value — one
    cell for the engine to rewrite."""
    with zlsx.write(path) as w:
        s = w.add_sheet("Sheet1")
        s.write_row_with_formulas([1, cached], [None, formula])


def _skip_unless_recalc():
    import zlsx._ffi as ffi

    if not (ffi._HAS_RECALC and ffi._HAS_EVAL and ffi._HAS_CANCEL):
        pytest.skip("loaded libzlsx predates the formula engine ABI (0.9.0+)")
    return ffi


def test_editor_recalculate_rewrites_stale_cache(tmp_path):
    _skip_unless_recalc()
    src = tmp_path / "recalc_src.xlsx"
    _write_formula_fixture(src)

    ed = zlsx.Editor(src)
    report = ed.recalculate(now=1_700_000_000_000, seed=42)
    assert report.cells_written >= 1
    assert report.kept_stale is False
    assert report.cancelled_late is False
    assert report.durability_warning is False
    assert report.resolved is not None
    assert report.resolved.now == 1_700_000_000_000
    assert report.resolved.seed == 42
    # A recalc derives dialect per stored cell; the echo says so.
    assert report.resolved.dialect is None

    got = ed.evaluate("=B1", now=1_700_000_000_000, seed=1)
    assert got.value == 3.0
    ed.close()


def test_editor_evaluate_value_shapes(tmp_path):
    _skip_unless_recalc()
    src = tmp_path / "eval_src.xlsx"
    _write_formula_fixture(src)

    ed = zlsx.Editor(src)
    ctx = dict(now=1_700_000_000_000, seed=7)

    # Scalar float.
    assert ed.evaluate("=A1+2", **ctx).value == 3.0
    # Text.
    assert ed.evaluate('="a"&"b"', **ctx).value == "ab"
    # Bool.
    assert ed.evaluate("=1<2", **ctx).value is True
    # An Excel error VALUE is a successful result, not an exception.
    err = ed.evaluate("=1/0", **ctx).value
    assert isinstance(err, zlsx.ExcelError)
    assert err == "#DIV/0!"
    # Matrix, row-major, blanks publish as 0.
    m = ed.evaluate("={1,2;3,4}", **ctx).value
    assert isinstance(m, zlsx.Matrix)
    assert (m.rows, m.cols) == (2, 2)
    assert m.cells == [[1.0, 2.0], [3.0, 4.0]]
    assert ed.evaluate("=D7", **ctx).value == 0.0

    # Typed refusal raises ZlsxFormulaRefusal (a ZlsxError subclass).
    with pytest.raises(zlsx.ZlsxFormulaRefusal) as exc_info:
        ed.evaluate("=1+", **ctx)
    assert exc_info.value.error_name == "FormulaMalformedInput"
    ed.close()


def test_editor_evaluate_resolved_echo_replays(tmp_path):
    _skip_unless_recalc()
    src = tmp_path / "echo_src.xlsx"
    _write_formula_fixture(src)

    import time

    ed = zlsx.Editor(src)
    # Two defaulted evaluations resolve different contexts (the binding
    # reads the clock; the library never does).
    first = ed.evaluate("=NOW()")
    time.sleep(0.002)
    second = ed.evaluate("=NOW()")
    assert first.resolved.now != second.resolved.now
    assert first.value != second.value

    # Replaying either resolved context reproduces its result exactly.
    r = first.resolved
    replay = ed.evaluate(
        "=NOW()",
        now=r.now,
        utc_offset_min=r.utc_offset_min,
        seed=r.seed,
        mode=r.mode,
        profile=r.profile,
    )
    assert replay.value == first.value
    ed.close()


def test_editor_save_with_recalc_atomic(tmp_path):
    ffi = _skip_unless_recalc()
    if not ffi._HAS_SAVE_WITH_RECALC:
        pytest.skip("loaded libzlsx predates save_with_recalc (0.9.0+)")
    src = tmp_path / "swr_src.xlsx"
    dst = tmp_path / "swr_dst.xlsx"
    _write_formula_fixture(src)

    ed = zlsx.Editor(src)
    report = ed.save_with_recalc(dst, now=1_700_000_000_000, seed=42)
    assert report.cells_written >= 1
    assert report.durability_warning is False
    assert report.cancelled_late is False
    ed.close()

    # The destination holds the recalced cache; the source is untouched.
    _header, rows = zlsx.read(dst)
    assert rows[0][1] == 3.0
    _header, rows = zlsx.read(src)
    assert rows[0][1] == 0.0


def test_editor_save_to_buffer_and_from_bytes(tmp_path):
    ffi = _skip_unless_recalc()
    if not ffi._HAS_SAVE_BUFFER:
        pytest.skip("loaded libzlsx predates the editor buffer ABI (0.9.0+)")
    src = tmp_path / "buf_src.xlsx"
    _write_formula_fixture(src)
    src_bytes = src.read_bytes()

    # from_bytes copies: mutate the source object after, nothing breaks.
    blob = bytearray(src_bytes)
    ed = zlsx.Editor.from_bytes(bytes(blob))
    for i in range(len(blob)):
        blob[i] = 0xAA

    # An untouched editor round-trips the source bytes verbatim.
    assert ed.save_to_buffer() == src_bytes

    # A mutated one carries the mutation, and the buffer reopens.
    ed.set_cell(0, 1, 0, 7)
    out = ed.save_to_buffer()
    assert out != src_bytes
    ed2 = zlsx.Editor.from_bytes(out)
    got = ed2.evaluate("=A1", now=1, seed=1)
    assert got.value == 7.0
    ed2.close()
    ed.close()


def test_editor_recalc_refusal_carries_census(tmp_path):
    _skip_unless_recalc()
    src = tmp_path / "refusal_src.xlsx"
    _write_formula_fixture(src, formula="NOSUCHFN(A1)", cached=999.0)

    ed = zlsx.Editor(src)
    with pytest.raises(zlsx.ZlsxFormulaRefusal) as exc_info:
        ed.recalculate(now=1_700_000_000_000, seed=1)
    refusal = exc_info.value
    assert refusal.error_name == "FormulaUnsupportedFunction"
    # The refusing cell survives to Python: sheet 0, row 1 (1-based),
    # col 1 (0-based) — the M9a2 seam through recalc_run.prepare.
    assert refusal.cells == [(0, 1, 1)]
    assert len(refusal.census) == 1
    assert refusal.census[0].plane == "FormulaUnsupportedFunction"

    # And the workbook is untouched: the stale cache still reads back.
    assert ed.evaluate("=B1", now=1, seed=1).value == 999.0
    ed.close()


def test_editor_recalc_keep_stale_and_mark(tmp_path):
    _skip_unless_recalc()
    src = tmp_path / "mark_src.xlsx"
    dst = tmp_path / "mark_dst.xlsx"
    _write_formula_fixture(src, formula="NOSUCHFN(A1)", cached=999.0)

    ed = zlsx.Editor(src)
    report = ed.recalculate(
        now=1_700_000_000_000, seed=1, on_unsupported="keep_stale_and_mark"
    )
    assert report.kept_stale is True
    assert len(report.census) == 1
    assert report.census[0].plane == "FormulaUnsupportedFunction"
    ed.save(dst)
    ed.close()

    import zipfile

    with zipfile.ZipFile(dst) as z:
        wb_xml = z.read("xl/workbook.xml").decode("utf-8")
    assert 'fullCalcOnLoad="1"' in wb_xml


def test_editor_mark_recalc_on_load(tmp_path):
    ffi = _skip_unless_recalc()
    if not ffi._HAS_MARK_RECALC:
        pytest.skip("loaded libzlsx predates mark_recalc_on_load (0.9.0+)")
    src = tmp_path / "mrol_src.xlsx"
    dst = tmp_path / "mrol_dst.xlsx"
    _write_formula_fixture(src)

    ed = zlsx.Editor(src)
    ed.mark_recalc_on_load()
    ed.save(dst)
    ed.close()

    import zipfile

    with zipfile.ZipFile(dst) as z:
        wb_xml = z.read("xl/workbook.xml").decode("utf-8")
    assert 'fullCalcOnLoad="1"' in wb_xml


def test_editor_recalc_timeout_pre_commit(tmp_path):
    ffi = _skip_unless_recalc()
    if not ffi._HAS_SAVE_WITH_RECALC:
        pytest.skip("loaded libzlsx predates save_with_recalc (0.9.0+)")
    src = tmp_path / "timeout_src.xlsx"
    dst = tmp_path / "timeout_dst.xlsx"
    # Enough formula cells that a 1 ms deadline reliably trips one of
    # the engine's §5.5 poll points before the commit.
    with zlsx.write(src) as w:
        s = w.add_sheet("Big")
        s.write_row([1, 1])
        for i in range(2, 20_002):
            s.write_row_with_formulas([None, 0], [None, f"B{i - 1}+A$1"])

    ed = zlsx.Editor(src)
    with pytest.raises(TimeoutError):
        ed.save_with_recalc(dst, now=1_700_000_000_000, seed=1, timeout=0.001)
    # Pre-commit by contract: the destination never appeared and the
    # memory is untouched (the stale cache still reads back).
    assert not dst.exists()
    assert ed.evaluate("=B3", now=1, seed=1).value == 0.0

    # And with room to finish, the same transaction commits.
    report = ed.save_with_recalc(dst, now=1_700_000_000_000, seed=1, timeout=120.0)
    assert report.cells_written >= 20_000
    assert dst.exists()
    ed.close()


def test_writer_save_recalculate_option(tmp_path):
    ffi = _skip_unless_recalc()
    if not ffi._HAS_WRITER_RECALC:
        pytest.skip("loaded libzlsx predates writer save-with-recalc (0.9.0+)")
    dst = tmp_path / "writer_recalc.xlsx"

    w = zlsx.Writer(dst)
    s = w.add_sheet("Sheet1")
    s.write_row_with_formulas([1, 0], [None, "A1+2"])
    report = w.save(recalculate=zlsx.RecalcOptions(now=1_700_000_000_000, seed=5))
    assert report is not None
    assert report.cells_written >= 1
    # The writer is not consumed: a plain save still works after.
    w.save(tmp_path / "writer_plain.xlsx")
    w.close()

    _header, rows = zlsx.read(dst)
    assert rows[0][1] == 3.0


def test_write_row_with_formulas_v2_cse_rectangle(tmp_path):
    ffi = _skip_unless_recalc()
    if not ffi._HAS_FORMULAS_V2:
        pytest.skip("loaded libzlsx predates the formulas_v2 ABI (0.9.0+)")
    dst = tmp_path / "cse.xlsx"

    w = zlsx.Writer(dst)
    s = w.add_sheet("Sheet1")
    # A1 anchors A1:B2; B1/A2/B2 are members (A2 stays empty — the
    # writer emits its placeholder).
    s.write_row_with_formulas(
        [1, 2], [zlsx.FormulaSpec.cse("TRANSPOSE(D1:E2)", "A1:B2"), None]
    )
    s.write_row_with_formulas([None, 4], [None, None])
    w.save()
    w.close()

    import zipfile

    with zipfile.ZipFile(dst) as z:
        sheet_xml = z.read("xl/worksheets/sheet1.xml").decode("utf-8")
    assert '<f t="array" ref="A1:B2">TRANSPOSE(D1:E2)</f>' in sheet_xml
    assert '<c r="A2"/>' in sheet_xml

    # The state machine refuses across the boundary: a mismatched
    # anchor is an error, and an incomplete rectangle refuses the save.
    w2 = zlsx.Writer(tmp_path / "cse_bad.xlsx")
    s2 = w2.add_sheet("Sheet1")
    with pytest.raises(zlsx.ZlsxError):
        s2.write_row_with_formulas(
            [1, 2], [zlsx.FormulaSpec.cse("1+1", "B1:B2"), None]
        )
    s2.write_row_with_formulas(
        [1, 2], [zlsx.FormulaSpec.cse("1+1", "A1:A3"), None]
    )
    with pytest.raises(zlsx.ZlsxError):
        w2.save()
    w2.close()

    # FormulaSpec validates locally: cse needs a ref, others refuse one.
    with pytest.raises(ValueError):
        zlsx.FormulaSpec("1+1", "cse")
    with pytest.raises(ValueError):
        zlsx.FormulaSpec("1+1", "scalar", ref="A1:B2")
    with pytest.raises(ValueError):
        zlsx.FormulaSpec("1+1", "nonsense")


def test_write_row_with_formulas_dialect_kwarg(tmp_path):
    ffi = _skip_unless_recalc()
    if not ffi._HAS_FORMULAS_V2:
        pytest.skip("loaded libzlsx predates the formulas_v2 ABI (0.9.0+)")
    dst = tmp_path / "dialect_kwarg.xlsx"

    w = zlsx.Writer(dst)
    s = w.add_sheet("Sheet1")
    # Row-wide scalar dialect is sugar over the same v2 path.
    s.write_row_with_formulas([1, 3], [None, "A1+2"], dialect="scalar")
    # Row-wide 'cse' is per-cell only.
    with pytest.raises(ValueError):
        s.write_row_with_formulas([1], ["A1"], dialect="cse")
    # dynamic_array is parked: the writer refuses it (§5.8b).
    with pytest.raises(zlsx.ZlsxError):
        s.write_row_with_formulas([1], ["A1#"], dialect="dynamic_array")
    w.save()
    w.close()

    b = zlsx.open(dst)
    assert list(b.sheet(0).rows())[0][1] == 3.0
    b.close()


def test_engine_fingerprint_names_identity():
    ffi = _skip_unless_recalc()
    if not ffi._HAS_FINGERPRINT:
        pytest.skip("loaded libzlsx predates the fingerprint ABI (0.9.0+)")
    fp = zlsx.engine_fingerprint()
    assert fp.startswith("zlsx ")
    for component in ("excel_fp_rules_v1", "rng_v1", "collation_v1"):
        assert component in fp


# ── S3a: structural edits + the pivots read through the C ABI ─────────
#
# The Zig `Editor` has carried these since Phase 3e; S3a is the parity
# row that lets them cross the boundary. Every refusal is a typed
# `ZlsxRefusal` (status -2, `error_name` from the diag); statements
# about the call stay plain `ZlsxError`s.


def _require_structural():
    import zlsx._ffi as ffi
    if not ffi._HAS_EDITOR or not ffi._HAS_STRUCTURAL_EDITS:
        pytest.skip("structural edits not exposed in loaded libzlsx (requires 0.9.0+)")


def _three_by_three(path):
    with zlsx.write(path) as w:
        s = w.add_sheet("Data")
        s.write_row([1, 2, 3])
        s.write_row([4, 5, 6])
        s.write_row([7, 8, 9])
        w.add_sheet("Second").write_row(["two"])


def test_editor_structural_edits_round_trip(tmp_path):
    """insert_row / delete_column / insert_column / delete_row /
    add_sheet / rename_sheet / delete_sheet, then save → the reader
    sees the shifted grid and the new sheet list."""
    _require_structural()
    src = tmp_path / "src.xlsx"
    out = tmp_path / "out.xlsx"
    _three_by_three(src)

    with zlsx.edit(src) as ed:
        ed.insert_row(0, 2)        # [1,2,3] / blank / [4,5,6] / [7,8,9]
        ed.delete_column(0, 0)     # [2,3] / blank / [5,6] / [8,9]
        ed.insert_column(0, 0)     # [_,2,3] / blank / [_,5,6] / [_,8,9]
        ed.delete_row(0, 4)        # drops [_,8,9]
        assert ed.add_sheet("Third") == 2
        ed.rename_sheet(1, "Renamed")
        ed.delete_sheet(2)
        ed.save(out)

    with zlsx.open(out) as book:
        assert [book.sheet(i).name for i in range(2)] == ["Data", "Renamed"]
        rows = list(book.sheet(0).rows())
        # The reader skips the blank row; the two surviving data rows
        # start with the inserted blank column.
        assert rows == [[None, 2, 3], [None, 5, 6]]
        assert list(book.sheet(1).rows()) == [["two"]]


def test_editor_structural_refusals_are_typed(tmp_path):
    _require_structural()
    src = tmp_path / "src.xlsx"
    _three_by_three(src)

    with zlsx.edit(src) as ed:
        with pytest.raises(zlsx.ZlsxRefusal) as info:
            ed.add_sheet("data")            # ASCII case-insensitive duplicate
        assert info.value.error_name == "DuplicateSheetName"
        assert isinstance(info.value, zlsx.ZlsxError)
        assert not isinstance(info.value, zlsx.ZlsxFormulaRefusal)
        with pytest.raises(zlsx.ZlsxRefusal) as info:
            ed.rename_sheet(0, "SECOND")
        assert info.value.error_name == "DuplicateSheetName"
        with pytest.raises(zlsx.ZlsxRefusal) as info:
            ed.rename_table_column("Nope", "A", "B")
        assert info.value.error_name == "TableNotFound"
        ed.delete_sheet(1)
        with pytest.raises(zlsx.ZlsxRefusal) as info:
            ed.delete_sheet(0)
        assert info.value.error_name == "CannotDeleteLastSheet"
        assert "CannotDeleteLastSheet" in str(info.value)


def test_editor_structural_call_errors_are_plain_zlsx_errors(tmp_path):
    _require_structural()
    src = tmp_path / "src.xlsx"
    _three_by_three(src)

    with zlsx.edit(src) as ed:
        for call, name in (
            (lambda: ed.insert_row(9, 1), "SheetIndexOutOfRange"),
            (lambda: ed.delete_row(0, 0), "RowIndexOutOfRange"),
            (lambda: ed.insert_column(0, 16384), "ColumnIndexOutOfRange"),
            (lambda: ed.add_sheet(""), "InvalidSheetName"),
            (lambda: ed.rename_sheet(0, "a:b"), "InvalidSheetName"),
        ):
            with pytest.raises(zlsx.ZlsxError) as info:
                call()
            assert not isinstance(info.value, zlsx.ZlsxRefusal)
            assert name in str(info.value)
        # A staged cell write makes the sheet unclean for a structural
        # edit — a sequencing error, not a refusal.
        ed.set_cell(0, 1, 0, 42)
        with pytest.raises(zlsx.ZlsxError, match="RowEditRequiresCleanSheet") as info:
            ed.insert_row(0, 1)
        assert not isinstance(info.value, zlsx.ZlsxRefusal)

    ed = zlsx.edit(src)
    ed.close()
    with pytest.raises(zlsx.ZlsxError, match="closed"):
        ed.insert_row(0, 1)
    with pytest.raises(zlsx.ZlsxError, match="closed"):
        ed.pivots()


def test_editor_structural_edits_carry_the_rewriters_pivot_footprint_refuses(tmp_path):
    """The corpus pivot workbook: a row edit inside a hosted pivot's
    rectangle refuses (`RowEditUnsafeForSheet`), one above it lifts
    the pivot in step and `pivots()` reports the moved location."""
    _require_structural()
    import zlsx._ffi as ffi
    if not ffi._HAS_PIVOTS:
        pytest.skip("pivots read not exposed in loaded libzlsx")
    src = _skip_if_missing("openxlsx_loadExample.xlsx")
    work = tmp_path / "pivot.xlsx"
    work.write_bytes(src.read_bytes())

    before = zlsx.pivots(work)
    hosts = [p for p in before if p["kind"] == "pivot"]
    assert hosts, "corpus workbook carries pivot tables"
    host = hosts[0]
    ref = host["location"]["ref"]
    top = int("".join(ch for ch in ref.split(":")[0] if ch.isdigit()))

    with zlsx.edit(work) as ed:
        with pytest.raises(zlsx.ZlsxRefusal) as info:
            ed.delete_row(host["sheet_idx"], top)
        assert info.value.error_name == "RowEditUnsafeForSheet"
        ed.insert_row(host["sheet_idx"], 1)
        after = ed.pivots()
    moved = [p for p in after if p["kind"] == "pivot" and p["name"] == host["name"]][0]
    moved_top = int("".join(ch for ch in moved["location"]["ref"].split(":")[0] if ch.isdigit()))
    assert moved_top == top + 1


def test_pivots_frozen_shape_on_corpus_and_empty_on_plain_workbook(tmp_path):
    import zlsx._ffi as ffi
    if not ffi._HAS_EDITOR or not ffi._HAS_PIVOTS:
        pytest.skip("pivots read not exposed in loaded libzlsx (requires 0.9.0+)")

    plain = tmp_path / "plain.xlsx"
    _three_by_three(plain)
    assert zlsx.pivots(plain) == []
    with zlsx.edit(plain) as ed:
        assert ed.pivots() == []

    src = _skip_if_missing("openxlsx_loadExample.xlsx")
    records = zlsx.pivots(src)
    pivots = [r for r in records if r["kind"] == "pivot"]
    assert len(pivots) == 2
    frozen_keys = [
        "kind", "sheet", "sheet_idx", "name", "part", "location", "rows", "cols",
        "pages", "values", "data_caption", "grand_totals", "style", "cache",
    ]
    for p in pivots:
        assert list(p.keys()) == frozen_keys
        assert set(p["location"]) == {"ref", "first_header_row", "first_data_row", "first_data_col"}
        for axis in p["rows"] + p["cols"] + p["pages"]:
            assert axis == {"values": True} or set(axis) == {"field", "idx"}
        for v in p["values"]:
            assert set(v) == {"name", "field", "idx", "subtotal", "show_data_as", "num_fmt_id"}
        assert set(p["grand_totals"]) == {"rows", "cols"}
        cache = p["cache"]
        assert set(cache) == {
            "id", "part", "records_part", "record_count", "refreshed_by", "refreshed_date",
            "refresh_on_load", "save_data", "source", "fields",
        }
        assert cache["source"]["type"] == "worksheet"
        # Both corpus caches are table-named and resolve to a local sheet.
        assert cache["source"]["resolved"]["via"] == "table"
        assert cache["source"]["unresolved"] is None
        for f in cache["fields"]:
            assert set(f) == {"name", "num_fmt_id", "formula", "items", "types", "min", "max"}
    # No orphan caches in the corpus workbook.
    assert all(r["kind"] == "pivot" for r in records)


def test_editor_structural_indices_are_bounded_before_ctypes_narrowing(tmp_path):
    """c_uint32 wraps modulo 2**32: without a guard, rename_sheet(2**32, …)
    would rename sheet 0. Every integer-bearing structural method rejects
    out-of-range values before the FFI call, and the workbook is untouched."""
    _require_structural()
    src = tmp_path / "src.xlsx"
    _three_by_three(src)
    with zlsx.edit(src) as ed:
        for bad in (2**32, 2**32 + 1, -1, -(2**32) + 1):
            for call in (
                lambda: ed.insert_row(bad, 1),
                lambda: ed.insert_row(0, bad),
                lambda: ed.delete_row(bad, 1),
                lambda: ed.delete_row(0, bad),
                lambda: ed.insert_column(bad, 0),
                lambda: ed.insert_column(0, bad),
                lambda: ed.delete_column(bad, 0),
                lambda: ed.delete_column(0, bad),
                lambda: ed.rename_sheet(bad, "X"),
                lambda: ed.delete_sheet(bad),
            ):
                with pytest.raises(ValueError, match="4294967295"):
                    call()
        assert ed.save_to_buffer() == src.read_bytes()
    with zlsx.open(src) as book:
        assert [book.sheet(i).name for i in range(2)] == ["Data", "Second"]
