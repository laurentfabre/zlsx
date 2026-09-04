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

    # Typed refusal raises ZlsxFormulaRefusal — a ZlsxRefusal (S3a's base
    # for every typed refusal), hence a ZlsxError.
    with pytest.raises(zlsx.ZlsxFormulaRefusal) as exc_info:
        ed.evaluate("=1+", **ctx)
    assert exc_info.value.error_name == "FormulaMalformedInput"
    assert isinstance(exc_info.value, zlsx.ZlsxRefusal)
    assert isinstance(exc_info.value, zlsx.ZlsxError)
    assert issubclass(zlsx.ZlsxFormulaRefusal, zlsx.ZlsxRefusal)
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
    assert isinstance(refusal, zlsx.ZlsxRefusal)
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
        with pytest.raises(zlsx.ZlsxError, match="TableNotFound") as info:
            ed.rename_table_column("Nope", "A", "B")   # a selector, like a sheet index
        assert not isinstance(info.value, zlsx.ZlsxRefusal)
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


# ── S3b slice 2: the defined-names read through the C ABI ─────────────


def _require_defined_names():
    import zlsx._ffi as ffi
    if not ffi._HAS_EDITOR or not ffi._HAS_DEFINED_NAMES:
        pytest.skip("defined-names read not exposed in loaded libzlsx (requires 0.9.0+)")


def _named_workbook(path):
    with zlsx.write(path) as w:
        w.add_sheet("Data").write_row([1, 2, 3])
        w.add_sheet("Second").write_row(["two"])
        w.add_defined_name("Prices", "Data!$A$1:$C$4")
        w.add_defined_name("_xlnm.Print_Area", "Second!$A$1:$B$9", local_sheet_id=1)
        w.add_defined_name("Secret", "Data!$Z$1", hidden=True)


def test_defined_names_frozen_shape_and_empty_on_plain_workbook(tmp_path):
    """`Editor.defined_names()` / `zlsx.defined_names(path)` return the
    `zlsx defined-names` records as dicts — document order, hidden
    names streamed, `body` as authored; `[]` without names."""
    _require_defined_names()

    plain = tmp_path / "plain.xlsx"
    _three_by_three(plain)
    assert zlsx.defined_names(plain) == []
    with zlsx.edit(plain) as ed:
        assert ed.defined_names() == []

    src = tmp_path / "named.xlsx"
    _named_workbook(src)
    assert zlsx.defined_names(src) == [
        {"kind": "defined_name", "name": "Prices", "scope": "workbook",
         "sheet": None, "sheet_idx": None, "body": "Data!$A$1:$C$4", "hidden": False},
        {"kind": "defined_name", "name": "_xlnm.Print_Area", "scope": "sheet",
         "sheet": "Second", "sheet_idx": 1, "body": "Second!$A$1:$B$9", "hidden": False},
        {"kind": "defined_name", "name": "Secret", "scope": "workbook",
         "sheet": None, "sheet_idx": None, "body": "Data!$Z$1", "hidden": True},
    ]


def test_defined_names_read_the_editors_current_state(tmp_path):
    """A sheet rename is visible immediately: the name sweep rewrites
    the bodies and the view refreshes, no save in between."""
    _require_defined_names()
    _require_structural()
    src = tmp_path / "named.xlsx"
    _named_workbook(src)

    with zlsx.edit(src) as ed:
        ed.rename_sheet(0, "Facts")
        records = ed.defined_names()
    bodies = [r["body"] for r in records]
    assert bodies == ["Facts!$A$1:$C$4", "Second!$A$1:$B$9", "Facts!$Z$1"]
    assert records[1]["sheet"] == "Second"


def test_defined_names_refusal_is_typed(tmp_path):
    """An inventory the read cannot serve faithfully raises
    `ZlsxRefusal(MalformedWorkbookXml)` — never a partial list."""
    _require_defined_names()
    import zipfile

    src = tmp_path / "named.xlsx"
    _named_workbook(src)

    # A bad entity in one body: the editor opens (the open parser keeps
    # raw spans); the decode at read time refuses the whole view.
    broken = tmp_path / "broken.xlsx"
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(broken, "w") as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/workbook.xml":
                data = data.replace(b"Data!$Z$1", b"Data!$Z$1&bogus;")
            zout.writestr(item, data)

    with zlsx.edit(broken) as ed:
        with pytest.raises(zlsx.ZlsxRefusal) as info:
            ed.defined_names()
        assert info.value.error_name == "MalformedWorkbookXml"
        assert not isinstance(info.value, zlsx.ZlsxFormulaRefusal)

    ed = zlsx.edit(src)
    ed.close()
    with pytest.raises(zlsx.ZlsxError, match="closed"):
        ed.defined_names()


# ── S3b slice 6: the conditional-formats read through the C ABI ───────


def _require_conditional_formats():
    import zlsx._ffi as ffi
    if not ffi._HAS_EDITOR or not ffi._HAS_CONDITIONAL_FORMATS:
        pytest.skip("conditional-formats read not exposed in loaded libzlsx (requires 0.9.0+)")


def _cf_workbook(path):
    with zlsx.write(path) as w:
        dxf = w.add_dxf(zlsx.Dxf(font_bold=True, fill_fg_argb=0xFFFFC7CE))
        data = w.add_sheet("Data")
        data.write_row([1, 5, 9, 3])
        data.write_row([2, 6, 10, 4])
        data.add_conditional_format_cell_is("A1:A4", "between", "2", "4", dxf)
        data.add_conditional_format_expression("B1:B4", "B1>3", dxf)
        data.add_conditional_format_color_scale("C1:C4", 0xFFF8696B, None, 0xFF63BE7B)
        data.add_conditional_format_data_bar("D1:D4", 0xFF638EC6)
        report = w.add_sheet("Report")
        report.write_row(["R&D"])
        report.add_conditional_format_expression("A1:A2", '$A1="R&D"', dxf)


def test_conditional_formats_frozen_shape_and_empty_on_plain_workbook(tmp_path):
    """`Editor.conditional_formats()` / `zlsx.conditional_formats(path)`
    return the `zlsx conditional-formats` records as dicts — sheets in
    workbook order, rules in sheet-document order, the rule envelope
    only; `[]` without rules."""
    _require_conditional_formats()

    plain = tmp_path / "plain.xlsx"
    _three_by_three(plain)
    assert zlsx.conditional_formats(plain) == []
    with zlsx.edit(plain) as ed:
        assert ed.conditional_formats() == []

    src = tmp_path / "cf.xlsx"
    _cf_workbook(src)
    assert zlsx.conditional_formats(src) == [
        {"kind": "conditional_format", "sheet": "Data", "sheet_idx": 0, "sqref": "A1:A4",
         "rule_type": "cellIs", "formulas": ["2", "4"], "dxf_id": 0, "priority": 1},
        {"kind": "conditional_format", "sheet": "Data", "sheet_idx": 0, "sqref": "B1:B4",
         "rule_type": "expression", "formulas": ["B1>3"], "dxf_id": 0, "priority": 2},
        {"kind": "conditional_format", "sheet": "Data", "sheet_idx": 0, "sqref": "C1:C4",
         "rule_type": "colorScale", "formulas": [], "dxf_id": None, "priority": 3},
        {"kind": "conditional_format", "sheet": "Data", "sheet_idx": 0, "sqref": "D1:D4",
         "rule_type": "dataBar", "formulas": [], "dxf_id": None, "priority": 4},
        {"kind": "conditional_format", "sheet": "Report", "sheet_idx": 1, "sqref": "A1:A2",
         "rule_type": "expression", "formulas": ['$A1="R&D"'], "dxf_id": 0, "priority": 1},
    ]


def test_conditional_formats_read_the_editors_current_state(tmp_path):
    """A sheet rename is visible immediately, and a row insert moves
    the ENVELOPE with the bodies — sqref and formula on one grid, no
    save in between (Codex #216 r1 S3B-REL-301)."""
    _require_conditional_formats()
    _require_structural()
    src = tmp_path / "cf.xlsx"
    _cf_workbook(src)

    with zlsx.edit(src) as ed:
        ed.rename_sheet(0, "Facts")
        records = ed.conditional_formats()
        sheets = [r["sheet"] for r in records]
        assert sheets == ["Facts", "Facts", "Facts", "Facts", "Report"]
        assert records[0]["sqref"] == "A1:A4"

        ed.insert_row(0, 1)
        moved = ed.conditional_formats()
    assert [(r["sqref"], r["formulas"]) for r in moved] == [
        ("A2:A5", ["2", "4"]),
        ("B2:B5", ["B2>3"]),
        ("C2:C5", []),
        ("D2:D5", []),
        ("A1:A2", ['$A1="R&D"']),  # the other sheet did not move
    ]


def test_conditional_formats_entity_rename_round_trips(tmp_path):
    """Renaming to an entity-bearing name works and the CF read
    reports the decoded meaning; a second sheet cannot take the same
    meaning however the first is spelled (Codex #216 r16
    S3B-REL-1507)."""
    _require_conditional_formats()
    _require_structural()
    src = tmp_path / "cf.xlsx"
    _cf_workbook(src)

    with zlsx.edit(src) as ed:
        ed.rename_sheet(0, "R&D")
        records = ed.conditional_formats()
        assert records[0]["sheet"] == "R&D"
        with pytest.raises(zlsx.ZlsxRefusal):
            ed.rename_sheet(1, "r&d")


def test_conditional_formats_between_moves_both_formulas(tmp_path):
    """A `cellIs between` carries two formulas — a sweep that moved
    only the first left `B2` beside a moved sibling (Codex #216 r3
    S3B-REL-801)."""
    _require_conditional_formats()
    _require_structural()
    src = tmp_path / "cf_between.xlsx"
    with zlsx.write(src) as w:
        dxf = w.add_dxf(zlsx.Dxf(font_bold=True))
        data = w.add_sheet("Data")
        data.write_row([1, 5])
        data.write_row([2, 6])
        data.add_conditional_format_cell_is("A1:A4", "between", "B1", "B2", dxf)

    with zlsx.edit(src) as ed:
        ed.insert_row(0, 1)
        records = ed.conditional_formats()
    assert [(r["sqref"], r["formulas"]) for r in records] == [
        ("A2:A5", ["B2", "B3"]),
    ]


def test_conditional_formats_collapse_delete_refuses_typed(tmp_path):
    """Deleting the only row a rule targets raises
    `ZlsxRefusal(SqrefCollapseUnsafe)` and mutates nothing — Excel
    deletes such a rule outright; zlsx refuses rather than silently
    retarget it (Codex #216 r4 S3B-REL-805)."""
    _require_conditional_formats()
    _require_structural()
    src = tmp_path / "cf_collapse.xlsx"
    with zlsx.write(src) as w:
        dxf = w.add_dxf(zlsx.Dxf(font_bold=True))
        data = w.add_sheet("Data")
        data.write_row([1, 2, 3, 4])
        data.add_conditional_format_expression("A1:D1", "A1>0", dxf)

    with zlsx.edit(src) as ed:
        with pytest.raises(zlsx.ZlsxRefusal) as info:
            ed.delete_row(0, 1)
        assert info.value.error_name == "SqrefCollapseUnsafe"
        assert [(r["sqref"], r["formulas"]) for r in ed.conditional_formats()] == [
            ("A1:D1", ["A1>0"]),
        ]


def test_conditional_formats_refusal_is_typed(tmp_path):
    """An inventory the read cannot serve faithfully raises
    `ZlsxRefusal(MalformedSheetXml)` — never a partial list."""
    _require_conditional_formats()
    import zipfile

    src = tmp_path / "cf.xlsx"
    _cf_workbook(src)

    # A bad entity in one formula body: the editor opens (the open
    # parser keeps raw spans); the decode at read time refuses the
    # whole view with the sheet part's verdict.
    broken = tmp_path / "broken.xlsx"
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(broken, "w") as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/worksheets/sheet1.xml":
                data = data.replace(b"<formula>2</formula>", b"<formula>2&bogus;</formula>")
            zout.writestr(item, data)

    with zlsx.edit(broken) as ed:
        with pytest.raises(zlsx.ZlsxRefusal) as info:
            ed.conditional_formats()
        assert info.value.error_name == "MalformedSheetXml"
        assert not isinstance(info.value, zlsx.ZlsxFormulaRefusal)

    # A broken SECOND sheet refuses whole — the first sheet's four
    # perfectly servable records are never handed out.
    broken2 = tmp_path / "broken2.xlsx"
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(broken2, "w") as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/worksheets/sheet2.xml":
                data = data.replace(b"$A1=", b"$A1&bogus;=")
            zout.writestr(item, data)

    with zlsx.edit(broken2) as ed:
        with pytest.raises(zlsx.ZlsxRefusal) as info:
            ed.conditional_formats()
        assert info.value.error_name == "MalformedSheetXml"

    ed = zlsx.edit(src)
    ed.close()
    with pytest.raises(zlsx.ZlsxError, match="closed"):
        ed.conditional_formats()


def _require_anchors():
    import zlsx._ffi as ffi
    if not ffi._HAS_EDITOR or not ffi._HAS_ANCHORS:
        pytest.skip("anchors read not exposed in loaded libzlsx (requires 0.9.0+)")


_ANCHORS_PNG = b"\x89PNG\r\n\x1a\n01234567"

_ANCHORS_NS = (
    'xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
)

_ANCHORS_PIC = (
    '<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="{id}" name="{name}"/><xdr:cNvPicPr/></xdr:nvPicPr>'
    '<xdr:blipFill><a:blip r:embed="rIdI1"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>'
    '<xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>'
)

_ANCHORS_CHART_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    '<c:chart><c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:ser><c:idx val="0"/><c:order val="0"/>'
    '<c:tx><c:strRef><c:f>Data!$B$1</c:f></c:strRef></c:tx>'
    '<c:cat><c:strRef><c:f>Data!$A$2:$A$4</c:f></c:strRef></c:cat>'
    '<c:val><c:numRef><c:f>Data!$B$2:$B$4</c:f></c:numRef></c:val>'
    '</c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>'
)

# Report's drawing: the chart FIRST in document order, then a two-cell
# image, then an absolute image — the view regroups images before
# charts, so the fixture must not mirror the stream (the pkg fixture's
# shape, `anchor_ndjson.zig::fixture.write(.with_absolute)`).
_ANCHORS_DRAWING1 = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    f'<xdr:wsDr {_ANCHORS_NS}>'
    '<xdr:oneCellAnchor><xdr:from><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>'
    '<xdr:ext cx="3048000" cy="2286000"/><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="3" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>'
    '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
    '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rIdC1"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:oneCellAnchor>'
    '<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>1</xdr:col><xdr:colOff>9525</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>'
    '<xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>7</xdr:row><xdr:rowOff>19050</xdr:rowOff></xdr:to>'
    + _ANCHORS_PIC.format(id=2, name="Picture 1") + '<xdr:clientData/></xdr:twoCellAnchor>'
    '<xdr:absoluteAnchor><xdr:pos x="1000" y="2000"/><xdr:ext cx="914400" cy="457200"/>'
    + _ANCHORS_PIC.format(id=4, name="Picture 2") + '<xdr:clientData/></xdr:absoluteAnchor>'
    '</xdr:wsDr>'
)

# Data's drawing: one two-cell image, so the stream crosses a sheet
# boundary.
_ANCHORS_DRAWING2 = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    f'<xdr:wsDr {_ANCHORS_NS}>'
    '<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>'
    '<xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>'
    + _ANCHORS_PIC.format(id=2, name="Logo") + '<xdr:clientData/></xdr:twoCellAnchor>'
    '</xdr:wsDr>'
)

_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _rels_xml(*rels):
    body = "".join(
        f'<Relationship Id="{rid}" Type="{_REL_NS}/{leaf}" Target="{target}"/>'
        for rid, leaf, target in rels
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + body + "</Relationships>"
    )


def _anchors_workbook(path):
    """A real two-sheet workbook with anchors on BOTH sheets, the pkg
    fixture rebuilt by hand: the writer has no drawing surface, so the
    parts ride in through the archive."""
    import os
    import zipfile

    with zlsx.write(path) as w:
        data = w.add_sheet("Data")
        data.write_row(["Region", "Qty"])
        data.write_row(["East", 3])
        data.write_row(["West", 4])
        data.write_row(["East", 5])
        w.add_sheet("Report").write_row(["drawing host"])

    tmp = str(path) + ".tmp"
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            blob = zin.read(item.filename)
            if item.filename == "[Content_Types].xml":
                blob = blob.replace(
                    b"</Types>",
                    b'<Default Extension="png" ContentType="image/png"/>'
                    b'<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>'
                    b'<Override PartName="/xl/drawings/drawing2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>'
                    b'<Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>'
                    b"</Types>",
                )
            elif item.filename == "xl/worksheets/sheet1.xml":
                blob = blob.replace(b"</worksheet>", b'<drawing r:id="rIdD2"/></worksheet>')
            elif item.filename == "xl/worksheets/sheet2.xml":
                blob = blob.replace(b"</worksheet>", b'<drawing r:id="rIdD1"/></worksheet>')
            zout.writestr(item, blob)
        zout.writestr("xl/media/image1.png", _ANCHORS_PNG)
        zout.writestr("xl/charts/chart1.xml", _ANCHORS_CHART_XML)
        zout.writestr("xl/drawings/drawing1.xml", _ANCHORS_DRAWING1)
        zout.writestr("xl/drawings/_rels/drawing1.xml.rels", _rels_xml(("rIdI1", "image", "../media/image1.png"), ("rIdC1", "chart", "../charts/chart1.xml")))
        zout.writestr("xl/drawings/drawing2.xml", _ANCHORS_DRAWING2)
        zout.writestr("xl/drawings/_rels/drawing2.xml.rels", _rels_xml(("rIdI1", "image", "../media/image1.png")))
        zout.writestr("xl/worksheets/_rels/sheet1.xml.rels", _rels_xml(("rIdD2", "drawing", "../drawings/drawing2.xml")))
        zout.writestr("xl/worksheets/_rels/sheet2.xml.rels", _rels_xml(("rIdD1", "drawing", "../drawings/drawing1.xml")))
    os.replace(tmp, path)


def _patched_copy(src, dst, part, old, new):
    """`src` with the first `old` in `part` replaced by `new`, at `dst`."""
    import zipfile

    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            blob = zin.read(item.filename)
            if item.filename == part:
                assert old in blob
                blob = blob.replace(old, new, 1)
            zout.writestr(item, blob)


_ANCHORS_EXPECTED = [
    {"kind": "image_anchor", "sheet": "Data", "sheet_idx": 0, "part": "xl/media/image1.png",
     "anchor": "two_cell", "from": {"row": 1, "col": 1, "row_off": 0, "col_off": 0},
     "to": {"row": 4, "col": 3, "row_off": 0, "col_off": 0}, "absolute": None, "bytes": len(_ANCHORS_PNG)},
    {"kind": "image_anchor", "sheet": "Report", "sheet_idx": 1, "part": "xl/media/image1.png",
     "anchor": "two_cell", "from": {"row": 3, "col": 2, "row_off": 0, "col_off": 9525},
     "to": {"row": 8, "col": 5, "row_off": 19050, "col_off": 0}, "absolute": None, "bytes": len(_ANCHORS_PNG)},
    {"kind": "image_anchor", "sheet": "Report", "sheet_idx": 1, "part": "xl/media/image1.png",
     "anchor": "absolute", "from": None, "to": None,
     "absolute": {"x": 1000, "y": 2000, "cx": 914400, "cy": 457200}, "bytes": len(_ANCHORS_PNG)},
    {"kind": "chart_anchor", "sheet": "Report", "sheet_idx": 1, "part": "xl/charts/chart1.xml",
     "anchor": "one_cell", "from": {"row": 2, "col": 6, "row_off": 0, "col_off": 0}, "to": None,
     "absolute": None, "chart_type": "bar",
     "series_refs": ["Data!$B$1", "Data!$A$2:$A$4", "Data!$B$2:$B$4"]},
]


def test_anchors_frozen_shape_and_empty_on_plain_workbook(tmp_path):
    """`Editor.anchors()` / `zlsx.anchors(path)` return the `zlsx anchors`
    records as dicts — sheets in workbook order, a sheet's images before
    its charts (Report's chart is FIRST in its drawing), all three anchor
    kinds; `[]` without drawings."""
    _require_anchors()

    plain = tmp_path / "plain.xlsx"
    _three_by_three(plain)
    assert zlsx.anchors(plain) == []
    with zlsx.edit(plain) as ed:
        assert ed.anchors() == []

    src = tmp_path / "anchors.xlsx"
    _anchors_workbook(src)
    assert zlsx.anchors(src) == _ANCHORS_EXPECTED
    with zlsx.edit(src) as ed:
        assert ed.anchors() == _ANCHORS_EXPECTED


def test_anchors_read_the_editors_current_state(tmp_path):
    """A sheet rename is visible immediately in `sheet`, and a row
    insert moves the edited sheet's anchor with the grid while the
    other sheet's stay, no save in between. The chart's series
    formulas name the edited sheet, so they follow both edits (the
    chart ``<c:f>`` sweep) while the chart's own anchor on Report
    stays."""
    _require_anchors()
    _require_structural()
    _require_chart_sweep()
    src = tmp_path / "anchors.xlsx"
    _anchors_workbook(src)

    with zlsx.edit(src) as ed:
        ed.rename_sheet(0, "Facts")
        records = ed.anchors()
        assert [r["sheet"] for r in records] == ["Facts", "Report", "Report", "Report"]
        assert records[3]["series_refs"] == ["Facts!$B$1", "Facts!$A$2:$A$4", "Facts!$B$2:$B$4"]

        ed.insert_row(0, 1)
        moved = ed.anchors()
    assert (moved[0]["from"], moved[0]["to"]) == (
        {"row": 2, "col": 1, "row_off": 0, "col_off": 0},
        {"row": 5, "col": 3, "row_off": 0, "col_off": 0},
    )
    assert moved[1]["from"] == {"row": 3, "col": 2, "row_off": 0, "col_off": 9525}  # Report did not move
    assert moved[3]["from"] == {"row": 2, "col": 6, "row_off": 0, "col_off": 0}
    assert moved[3]["series_refs"] == ["Facts!$B$2", "Facts!$A$3:$A$5", "Facts!$B$3:$B$5"]


def _require_chart_sweep():
    """The chart ``<c:f>`` sweep is a behaviour of the structural edits
    and the anchors read, not an export of its own: it landed inside
    the unreleased 0.9.0, so there is nothing to probe beyond the two
    surfaces it rides — against a dylib that predates it these tests
    FAIL rather than skip, as a stale pin should."""
    _require_structural()
    _require_anchors()


def test_chart_series_formulas_ride_every_structural_edit(tmp_path):
    """Chart series formulas (``<c:f>``: series name, categories,
    values) ride the formula rewriter under a sheet rename, a row /
    column edit on the sheet they name, and a sheet delete — visible
    through ``Editor.anchors()`` with no save, and in the saved part."""
    _require_chart_sweep()
    import zipfile

    src = tmp_path / "anchors.xlsx"
    _anchors_workbook(src)
    dst = tmp_path / "moved.xlsx"
    with zlsx.edit(src) as ed:
        # An edit on the HOST sheet (Report carries the chart) moves the
        # anchor, not the carriers.
        ed.insert_row(1, 1)
        assert ed.anchors()[3]["series_refs"] == ["Data!$B$1", "Data!$A$2:$A$4", "Data!$B$2:$B$4"]
        # An insert inside the ranges on Data grows them; the
        # series-name cell above stays.
        ed.insert_row(0, 3)
        assert ed.anchors()[3]["series_refs"] == ["Data!$B$1", "Data!$A$2:$A$5", "Data!$B$2:$B$5"]
        # A column delete (0-based at the boundary: column A): the
        # category column collapses to the rewriter's qualified #REF!,
        # the value column slides left.
        ed.delete_column(0, 0)
        assert ed.anchors()[3]["series_refs"] == ["Data!$A$1", "Data!#REF!", "Data!$A$2:$A$5"]
        # A name needing quotes is quoted.
        ed.rename_sheet(0, "Raw Data")
        assert ed.anchors()[3]["series_refs"] == ["'Raw Data'!$A$1", "Data!#REF!", "'Raw Data'!$A$2:$A$5"]
        ed.save(dst)
    with zipfile.ZipFile(dst) as z:
        chart = z.read("xl/charts/chart1.xml").decode()
    assert "<c:f>'Raw Data'!$A$1</c:f>" in chart
    assert "<c:f>'Raw Data'!$A$2:$A$5</c:f>" in chart
    # Deleting the named sheet collapses every carrier into it. (The
    # deleted sheet's own drawing stays in the archive unreferenced —
    # `delete_sheet`'s documented orphan — so the strict anchors read
    # refuses `DrawingOnUnlistedSheet` afterwards; the saved chart part
    # is the witness.)
    gone = tmp_path / "gone.xlsx"
    with zlsx.edit(dst) as ed:
        ed.delete_sheet(0)
        ed.save(gone)
    with zipfile.ZipFile(gone) as z:
        chart = z.read("xl/charts/chart1.xml").decode()
    assert "<c:tx><c:strRef><c:f>#REF!</c:f></c:strRef></c:tx>" in chart
    # The category carrier had already collapsed to `Data!#REF!` under
    # the column delete; the rewriter leaves an error token as it lies
    # (its rule on every carrier), so the deleted qualifier stays.
    assert "<c:cat><c:strRef><c:f>Data!#REF!</c:f></c:strRef></c:cat>" in chart
    assert "<c:val><c:numRef><c:f>#REF!</c:f></c:numRef></c:val>" in chart


def test_chart_refusal_is_typed(tmp_path):
    """A chart part whose series carrier the walk cannot read whole
    refuses every structural edit before its first mutation:
    ``MalformedChartXml`` on a rename / delete, folded into the
    sheet-level names on a row / column edit — the ``<xm:f>`` shape."""
    _require_chart_sweep()

    src = tmp_path / "anchors.xlsx"
    _anchors_workbook(src)
    broken = tmp_path / "broken_chart.xlsx"
    _patched_copy(src, broken, "xl/charts/chart1.xml", b"<c:f>Data!$B$1</c:f>", b"<c:f>Data!$B$1")
    with zlsx.edit(broken) as ed:
        with pytest.raises(zlsx.ZlsxRefusal) as info:
            ed.rename_sheet(0, "Facts")
        assert info.value.error_name == "MalformedChartXml"
        assert not isinstance(info.value, zlsx.ZlsxFormulaRefusal)
        with pytest.raises(zlsx.ZlsxRefusal) as info:
            ed.delete_sheet(0)
        assert info.value.error_name == "MalformedChartXml"
        with pytest.raises(zlsx.ZlsxRefusal) as info:
            ed.insert_row(0, 1)
        assert info.value.error_name == "RowEditUnsafeForSheet"
        with pytest.raises(zlsx.ZlsxRefusal) as info:
            ed.delete_column(1, 0)
        assert info.value.error_name == "ColEditUnsafeForSheet"
        # The anchors read's own verdict on the same carrier.
        with pytest.raises(zlsx.ZlsxRefusal) as info:
            ed.anchors()
        assert info.value.error_name == "MalformedDrawingXml"


def test_anchors_refusal_is_typed(tmp_path):
    """An inventory the read cannot serve faithfully raises a typed
    `ZlsxRefusal` — never a partial list."""
    _require_anchors()
    import zipfile

    src = tmp_path / "anchors.xlsx"
    _anchors_workbook(src)

    # Data's image blip names a relationship its drawing does not hold:
    # the drawing graph cannot be read whole.
    broken = tmp_path / "broken.xlsx"
    _patched_copy(src, broken, "xl/drawings/drawing2.xml", b'r:embed="rIdI1"', b'r:embed="rIdXX"')
    with zlsx.edit(broken) as ed:
        with pytest.raises(zlsx.ZlsxRefusal) as info:
            ed.anchors()
        assert info.value.error_name == "MalformedDrawingXml"
        assert not isinstance(info.value, zlsx.ZlsxFormulaRefusal)

    # A broken SECOND sheet refuses whole — Data's perfectly servable
    # record is never handed out.
    broken2 = tmp_path / "broken2.xlsx"
    _patched_copy(src, broken2, "xl/drawings/drawing1.xml", b"<xdr:to>", b"<xdr:zz>")
    with zlsx.edit(broken2) as ed:
        with pytest.raises(zlsx.ZlsxRefusal) as info:
            ed.anchors()
        assert info.value.error_name == "MalformedDrawingXml"

    # A bad entity in a sheet-name carrier: the workbook-level read
    # refuses before any drawing is walked.
    broken3 = tmp_path / "broken3.xlsx"
    _patched_copy(src, broken3, "xl/workbook.xml", b'name="Report"', b'name="Rep&bogus;ort"')
    with zlsx.edit(broken3) as ed:
        with pytest.raises(zlsx.ZlsxRefusal) as info:
            ed.anchors()
        assert info.value.error_name == "MalformedWorkbookXml"

    # A copy of Report's part under a name no <sheet> entry reaches,
    # its drawing reference and rels riding along: the anchors on it
    # cannot be attributed, and the read refuses rather than drop them.
    orphan = tmp_path / "orphan.xlsx"
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(orphan, "w", zipfile.ZIP_DEFLATED) as zout:
        # Read the copies BEFORE the rewrite loop: writestr mutates the
        # shared ZipInfo objects (header offsets), after which zin can
        # no longer read by name.
        sheet2 = zin.read("xl/worksheets/sheet2.xml")
        rels2 = zin.read("xl/worksheets/_rels/sheet2.xml.rels")
        for item in zin.infolist():
            blob = zin.read(item.filename)
            if item.filename == "[Content_Types].xml":
                blob = blob.replace(
                    b"</Types>",
                    b'<Override PartName="/xl/worksheets/sheet9.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
                )
            zout.writestr(item, blob)
        zout.writestr("xl/worksheets/sheet9.xml", sheet2)
        zout.writestr("xl/worksheets/_rels/sheet9.xml.rels", rels2)
    with zlsx.edit(orphan) as ed:
        with pytest.raises(zlsx.ZlsxRefusal) as info:
            ed.anchors()
        assert info.value.error_name == "DrawingOnUnlistedSheet"

    ed = zlsx.edit(src)
    ed.close()
    with pytest.raises(zlsx.ZlsxError, match="closed"):
        ed.anchors()


def test_anchors_list_openpyxls_default_namespace_drawing_and_shift_it(tmp_path):
    """openpyxl 3.1 binds the spreadsheetDrawing namespace as its
    DEFAULT namespace (``<wsDr xmlns="…/spreadsheetDrawing">
    <oneCellAnchor><from><row>``): the read used to list nothing for
    such a workbook, and a row insert moved the grid and the chart's
    series formulas but not the anchor. The namespace-aware drawing
    slice resolves the prefix once for the read and the sweep — the
    committed openpyxl corpus workbook through both."""
    _require_anchors()
    _require_structural()
    path = _skip_if_missing("openpyxl_chart.xlsx")
    assert zlsx.anchors(path) == [
        {
            "kind": "chart_anchor",
            "sheet": "Data",
            "sheet_idx": 0,
            "part": "xl/charts/chart1.xml",
            "anchor": "one_cell",
            "from": {"row": 2, "col": 4, "row_off": 0, "col_off": 0},
            "to": None,
            "absolute": None,
            "chart_type": "bar",
            "series_refs": ["'Data'!B1", "'Data'!$A$2:$A$4", "'Data'!$B$2:$B$4"],
        }
    ]
    out = tmp_path / "openpyxl_moved.xlsx"
    with zlsx.edit(path) as ed:
        ed.insert_row(0, 1)
        moved = ed.anchors()
        assert moved[0]["from"] == {"row": 3, "col": 4, "row_off": 0, "col_off": 0}
        assert moved[0]["series_refs"] == ["'Data'!B2", "'Data'!$A$3:$A$5", "'Data'!$B$3:$B$5"]
        ed.save(out)
    assert zlsx.anchors(out)[0]["from"]["row"] == 3
    import zipfile

    with zipfile.ZipFile(out) as z:
        drawing = z.read("xl/drawings/drawing1.xml").decode()
    # The row is the one splice; the part stays as openpyxl spelled it.
    assert "<row>2</row>" in drawing and "<row>1</row>" not in drawing
    assert "xdr:" not in drawing


def test_anchors_unfollowable_drawing_binding_refuses_read_and_edit(tmp_path):
    """A spreadsheetDrawing binding under a name the anchor walk cannot
    spell (past the resolver's 100-byte prefix limit) would leave an
    anchor neither listed nor moved: the read refuses ``MalformedDrawingXml``
    and so does every row / column edit, before anything changes."""
    _require_anchors()
    _require_structural()
    import shutil
    import zipfile

    src = _skip_if_missing("openpyxl_chart.xlsx")
    path = tmp_path / "openpyxl_nsbind.xlsx"
    long_prefix = "p" * 101
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/drawings/drawing1.xml":
                data = data.replace(
                    b"<wsDr ",
                    b'<wsDr xmlns:' + long_prefix.encode() + b'="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" ',
                    1,
                )
            zout.writestr(item, data)
    with pytest.raises(zlsx.ZlsxRefusal) as info:
        zlsx.anchors(path)
    assert info.value.error_name == "MalformedDrawingXml"
    with zlsx.edit(path) as ed:
        with pytest.raises(zlsx.ZlsxRefusal) as info:
            ed.insert_row(0, 1)
        assert info.value.error_name == "MalformedDrawingXml"
        with pytest.raises(zlsx.ZlsxRefusal) as info:
            ed.delete_column(0, 0)
        assert info.value.error_name == "MalformedDrawingXml"
        # Pre-mutation: a rename, which runs no drawing sweep, still lands.
        ed.rename_sheet(0, "Facts")
        ed.save(tmp_path / "renamed.xlsx")
    del shutil


def _require_sheet_props():
    import zlsx._ffi as ffi
    if not ffi._HAS_EDITOR or not ffi._HAS_SHEET_PROPS:
        pytest.skip("sheet-props / calc-props reads not exposed in loaded libzlsx (requires 0.9.0+)")


def _patch_parts(path, patches):
    """Byte-replace the first `old` in each `part` of the saved workbook
    at `path`, in place — the pkg fixture's `patchPart`."""
    import os
    import zipfile

    tmp = str(path) + ".tmp"
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            blob = zin.read(item.filename)
            for part, old, new in patches:
                if item.filename == part:
                    assert old in blob, (part, old)
                    blob = blob.replace(old, new, 1)
            zout.writestr(item, blob)
    os.replace(tmp, path)


_CALC_PR = b'<calcPr calcId="191029" fullCalcOnLoad="1" iterate="true" iterateCount="100" iterateDelta="0.001"/>'


def _sheetless_copy(src, dst):
    """`src` with its whole `<sheets>…</sheets>` block replaced by an
    empty `<sheets/>` at `dst`, every sheet part and relationship kept
    — the sheetless shape the strict inventory refuses."""
    import re
    import zipfile

    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            blob = zin.read(item.filename)
            if item.filename == "xl/workbook.xml":
                new = re.sub(rb"<sheets>.*?</sheets>", b"<sheets/>", blob, count=1, flags=re.S)
                assert new != blob
                blob = new
            zout.writestr(item, blob)


def _sheet_props_workbook(path):
    """The pkg fixture (`sheet_props_ndjson.zig::fixture.write`) rebuilt
    by hand: Data frozen, Report split with a fractional split, Bare
    with nothing, a full `<calcPr>`. The writer emits no `<dimension>`
    and no `<calcPr>` and has no split-pane surface, so those ride in
    through the archive at their schema slots."""
    with zlsx.write(path) as w:
        data = w.add_sheet("Data")
        data.write_row(["Region", "Qty"])
        data.write_row(["East", 3])
        data.write_row(["West", 4])
        data.freeze_panes(1, 2)
        report = w.add_sheet("Report")
        report.write_row(["R&D", 1, 2])
        report.write_row(["Ops", 3, 4])
        w.add_sheet("Bare").write_row([1])
    _patch_parts(path, [
        ("xl/worksheets/sheet1.xml", b"<sheetViews>", b'<dimension ref="A1:B3"/><sheetViews>'),
        ("xl/worksheets/sheet2.xml", b"<sheetData>",
         b'<dimension ref="A1:C2"/><sheetViews><sheetView workbookViewId="0">'
         b'<pane xSplit="2865" ySplit="1215.5" topLeftCell="C4" activePane="bottomRight" state="split"/>'
         b"</sheetView></sheetViews><sheetData>"),
        ("xl/workbook.xml", b"</workbook>", _CALC_PR + b"</workbook>"),
    ])


_SHEET_PROPS_EXPECTED = [
    {"kind": "sheet_props", "sheet": "Data", "sheet_idx": 0, "dimension": "A1:B3",
     "pane": {"x_split": 2, "y_split": 1, "top_left_cell": "C2", "active_pane": "bottomRight", "state": "frozen"}},
    {"kind": "sheet_props", "sheet": "Report", "sheet_idx": 1, "dimension": "A1:C2",
     "pane": {"x_split": 2865, "y_split": 1215.5, "top_left_cell": "C4", "active_pane": "bottomRight", "state": "split"}},
    {"kind": "sheet_props", "sheet": "Bare", "sheet_idx": 2, "dimension": None, "pane": None},
]

_CALC_PROPS_EXPECTED = {
    "kind": "calc_props", "calc_id": 191029, "full_calc_on_load": True,
    "iterate": True, "iterate_count": 100, "iterate_delta": 0.001,
}

_CALC_PROPS_ABSENT = {
    "kind": "calc_props", "calc_id": None, "full_calc_on_load": None,
    "iterate": None, "iterate_count": None, "iterate_delta": None,
}


def test_sheet_props_frozen_shape_and_nulls_on_plain_workbook(tmp_path):
    """`Editor.sheet_props()` / `zlsx.sheet_props(path)` return the
    `zlsx sheet-props` records as dicts — one per sheet, workbook order,
    the split pane reported as written; a fresh writer's sheets are
    records of `None`s (no `<dimension>`, no views without a freeze),
    never an empty list."""
    _require_sheet_props()

    plain = tmp_path / "plain.xlsx"
    _three_by_three(plain)
    nulls = [
        {"kind": "sheet_props", "sheet": "Data", "sheet_idx": 0, "dimension": None, "pane": None},
        {"kind": "sheet_props", "sheet": "Second", "sheet_idx": 1, "dimension": None, "pane": None},
    ]
    assert zlsx.sheet_props(plain) == nulls
    with zlsx.edit(plain) as ed:
        assert ed.sheet_props() == nulls

    src = tmp_path / "props.xlsx"
    _sheet_props_workbook(src)
    assert zlsx.sheet_props(src) == _SHEET_PROPS_EXPECTED
    with zlsx.edit(src) as ed:
        records = ed.sheet_props()
    assert records == _SHEET_PROPS_EXPECTED
    # The wire's number spellings survive the parse: an integral split
    # is an int, a fractional one a float.
    assert type(records[0]["pane"]["x_split"]) is int
    assert type(records[1]["pane"]["y_split"]) is float


def test_calc_props_frozen_shape_and_absent_on_plain_workbook(tmp_path):
    """`Editor.calc_props()` / `zlsx.calc_props(path)` return the one
    `zlsx calc-props` record as a dict; a workbook without `<calcPr>` is
    a dict of `None`s, the `doc_props()` convention."""
    _require_sheet_props()

    plain = tmp_path / "plain.xlsx"
    _three_by_three(plain)
    assert zlsx.calc_props(plain) == _CALC_PROPS_ABSENT
    with zlsx.edit(plain) as ed:
        assert ed.calc_props() == _CALC_PROPS_ABSENT

    src = tmp_path / "props.xlsx"
    _sheet_props_workbook(src)
    rec = zlsx.calc_props(src)
    assert rec == _CALC_PROPS_EXPECTED
    # `True == 1` in Python; pin the JSON type, not just the value.
    assert rec["full_calc_on_load"] is True
    assert rec["iterate"] is True
    with zlsx.edit(src) as ed:
        assert ed.calc_props() == _CALC_PROPS_EXPECTED


def test_sheet_props_read_the_editors_current_state(tmp_path):
    """A sheet rename is visible immediately in `sheet`; a row insert
    below the frozen row grows the extent and moves the pane's top-left
    cell with the grid while the split holds; a split-pane sheet is the
    one such an edit refuses, and the refusal leaves the records as they
    were. No save in between."""
    _require_sheet_props()
    _require_structural()
    src = tmp_path / "props.xlsx"
    _sheet_props_workbook(src)

    with zlsx.edit(src) as ed:
        ed.rename_sheet(0, "Facts")
        renamed = ed.sheet_props()
        assert [r["sheet"] for r in renamed] == ["Facts", "Report", "Bare"]
        assert renamed[0]["dimension"] == "A1:B3"

        ed.insert_row(0, 2)
        moved = ed.sheet_props()
        assert moved[0] == {
            "kind": "sheet_props", "sheet": "Facts", "sheet_idx": 0, "dimension": "A1:B4",
            "pane": {"x_split": 2, "y_split": 1, "top_left_cell": "C3", "active_pane": "bottomRight", "state": "frozen"},
        }
        assert moved[1:] == _SHEET_PROPS_EXPECTED[1:]  # the other sheets did not move

        # The split pane the record reports is the one the row edit
        # refuses: the read and the editor's own contract agree.
        with pytest.raises(zlsx.ZlsxRefusal) as info:
            ed.insert_row(1, 1)
        assert info.value.error_name == "SplitPaneNotSupported"
        assert ed.sheet_props() == moved


def test_calc_props_read_the_editors_current_state(tmp_path):
    """The `<calcPr>` slot rides a rename's workbook.xml rewrite
    untouched, and `mark_recalc_on_load()` lands `full_calc_on_load` in
    the live part — visible with no save in between."""
    _require_sheet_props()
    _require_structural()
    import zlsx._ffi as ffi

    src = tmp_path / "props.xlsx"
    _sheet_props_workbook(src)
    with zlsx.edit(src) as ed:
        ed.rename_sheet(0, "Facts")
        assert ed.calc_props() == _CALC_PROPS_EXPECTED

    if not ffi._HAS_MARK_RECALC:
        pytest.skip("loaded libzlsx predates mark_recalc_on_load (0.9.0+)")
    plain = tmp_path / "plain.xlsx"
    _three_by_three(plain)
    with zlsx.edit(plain) as ed:
        assert ed.calc_props() == _CALC_PROPS_ABSENT
        ed.mark_recalc_on_load()
        rec = ed.calc_props()
    assert rec == {
        "kind": "calc_props", "calc_id": None, "full_calc_on_load": True,
        "iterate": None, "iterate_count": None, "iterate_delta": None,
    }
    assert rec["full_calc_on_load"] is True


def test_sheet_props_refusal_is_typed(tmp_path):
    """An inventory the read cannot serve faithfully raises a typed
    `ZlsxRefusal` — never a partial list."""
    _require_sheet_props()

    src = tmp_path / "props.xlsx"
    _sheet_props_workbook(src)

    # Two extents on Data: maxOccurs=1 is a refusal, not a pick.
    broken = tmp_path / "broken.xlsx"
    _patched_copy(src, broken, "xl/worksheets/sheet1.xml", b'<dimension ref="A1:B3"/>', b'<dimension ref="A1:B3"/><dimension ref="A1"/>')
    with zlsx.edit(broken) as ed:
        with pytest.raises(zlsx.ZlsxRefusal) as info:
            ed.sheet_props()
        assert info.value.error_name == "MalformedSheetXml"
        assert not isinstance(info.value, zlsx.ZlsxFormulaRefusal)

    # A duplicate attribute on the SECOND sheet's pane refuses whole —
    # Data's perfectly servable record is never handed out.
    broken2 = tmp_path / "broken2.xlsx"
    _patched_copy(src, broken2, "xl/worksheets/sheet2.xml", b'state="split"', b'state="split" state="frozen"')
    with zlsx.edit(broken2) as ed:
        with pytest.raises(zlsx.ZlsxRefusal) as info:
            ed.sheet_props()
        assert info.value.error_name == "MalformedSheetXml"

    # A bad entity in a sheet-name carrier: the workbook-level read
    # refuses before any sheet part is walked. The calc read walks the
    # same workbook.xml but attributes nothing, so it never decodes the
    # names — it still serves.
    broken3 = tmp_path / "broken3.xlsx"
    _patched_copy(src, broken3, "xl/workbook.xml", b'name="Report"', b'name="Rep&bogus;ort"')
    with zlsx.edit(broken3) as ed:
        with pytest.raises(zlsx.ZlsxRefusal) as info:
            ed.sheet_props()
        assert info.value.error_name == "MalformedWorkbookXml"
        assert ed.calc_props() == _CALC_PROPS_EXPECTED

    # An empty `<sheets/>`: the lenient opener accepts the sheetless
    # workbook, the strict inventory refuses it (CT_Sheets minOccurs=1)
    # — never `[]` from a read whose contract is one record per sheet.
    # Both reads share the walk's verdict.
    sheetless = tmp_path / "sheetless.xlsx"
    _sheetless_copy(src, sheetless)
    with zlsx.edit(sheetless) as ed:
        with pytest.raises(zlsx.ZlsxRefusal) as info:
            ed.sheet_props()
        assert info.value.error_name == "MalformedWorkbookXml"
        with pytest.raises(zlsx.ZlsxRefusal) as info:
            ed.calc_props()
        assert info.value.error_name == "MalformedWorkbookXml"

    ed = zlsx.edit(src)
    ed.close()
    with pytest.raises(zlsx.ZlsxError, match="closed"):
        ed.sheet_props()
    with pytest.raises(zlsx.ZlsxError, match="closed"):
        ed.calc_props()


def test_calc_props_refusal_is_typed(tmp_path):
    """A `<calcPr>` slot the read cannot report faithfully raises a
    typed `ZlsxRefusal` — never a guess at which element Excel honours."""
    _require_sheet_props()

    src = tmp_path / "props.xlsx"
    _sheet_props_workbook(src)

    # (name, old, new, whether the SHEET read still serves — it never
    # looks at the slot, but shares the workbook walk's own verdicts)
    cases = [
        # Two at the slot.
        ("two.xlsx", b"</workbook>", b'<calcPr calcId="1"/></workbook>', True),
        # One an MCE branch could project into the slot.
        ("mce.xlsx", _CALC_PR,
         b'<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
         b'<mc:Choice Requires="x15"><calcPr calcId="1"/></mc:Choice></mc:AlternateContent>', True),
        # A duplicate attribute; a carrier that does not decode.
        ("dup.xlsx", _CALC_PR, b'<calcPr calcId="1" calcId="2"/>', True),
        ("carrier.xlsx", _CALC_PR, b'<calcPr iterate="&bogus;"/>', True),
        # A `<sheets>` list the strict workbook walk cannot prove: the
        # calc read runs the same walk, the same verdict — for both.
        ("sheets.xlsx", b"</sheets>", b"</sheets><sheets/>", False),
    ]
    for name, old, new, sheet_read_serves in cases:
        broken = tmp_path / name
        _patched_copy(src, broken, "xl/workbook.xml", old, new)
        with zlsx.edit(broken) as ed:
            with pytest.raises(zlsx.ZlsxRefusal) as info:
                ed.calc_props()
            assert info.value.error_name == "MalformedWorkbookXml", name
            assert not isinstance(info.value, zlsx.ZlsxFormulaRefusal)
            if sheet_read_serves:
                assert ed.sheet_props() == _SHEET_PROPS_EXPECTED, name
            else:
                with pytest.raises(zlsx.ZlsxRefusal) as info:
                    ed.sheet_props()
                assert info.value.error_name == "MalformedWorkbookXml", name


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


def test_editor_structural_indices_reject_lossy_coercion(tmp_path):
    """`int(0.9)` is 0: an index that is not an integer is a TypeError,
    never truncated into a different cell; bool is refused too."""
    _require_structural()
    src = tmp_path / "src.xlsx"
    _three_by_three(src)

    class OnlyInt:
        def __int__(self):
            return 1

    with zlsx.edit(src) as ed:
        for bad in (0.9, 1.0, "2", True, OnlyInt(), None):
            for call in (
                lambda: ed.insert_row(bad, 1),
                lambda: ed.insert_row(0, bad),
                lambda: ed.delete_column(0, bad),
                lambda: ed.rename_sheet(bad, "X"),
                lambda: ed.delete_sheet(bad),
            ):
                with pytest.raises(TypeError):
                    call()
        assert ed.save_to_buffer() == src.read_bytes()


def test_editor_rename_table_column_round_trip_on_corpus(tmp_path):
    """Table2 on `IrisSample` (the corpus pivot workbook): rename
    `Species` → `Kind`, save, reopen — the header cell carries the new
    name, the pivot whose cache reads Table2 still resolves it, and the
    old name is now a selector that names nothing."""
    _require_structural()
    import zlsx._ffi as ffi
    if not ffi._HAS_PIVOTS:
        pytest.skip("pivots read not exposed in loaded libzlsx")
    src = _skip_if_missing("openxlsx_loadExample.xlsx")
    out = tmp_path / "renamed.xlsx"

    with zlsx.edit(src) as ed:
        ed.rename_table_column("Table2", "Species", "Kind")
        with pytest.raises(zlsx.ZlsxError, match="TableColumnNotFound") as info:
            ed.rename_table_column("Table2", "Species", "Other")
        assert not isinstance(info.value, zlsx.ZlsxRefusal)
        with pytest.raises(zlsx.ZlsxRefusal) as info:
            ed.rename_table_column("Table2", "Kind", "Sepal Width")
        assert info.value.error_name == "TableColumnNameInUse"
        ed.save(out)

    with zlsx.open(out) as book:
        header = next(iter(book.sheet("IrisSample").rows()))
        assert header[:5] == ["Sepal Length", "Sepal Width", "Petal Length", "Petal Width", "Kind"]
    sources = [p["cache"]["source"] for p in zlsx.pivots(out) if p["kind"] == "pivot"]
    table2 = [s for s in sources if s["name"] == "Table2"]
    assert table2 and table2[0]["resolved"]["via"] == "table"


def test_s3a_capability_probes_require_their_release_symbols():
    """A capability whose wrappers call a release function must not be
    advertised without it — the probes fold the prerequisites in."""
    import zlsx._ffi as ffi
    if ffi._HAS_STRUCTURAL_EDITS:
        assert ffi._HAS_DIAG_RELEASE and hasattr(ffi.lib, "zlsx_diag_release")
    if ffi._HAS_PIVOTS:
        assert ffi._HAS_DIAG_RELEASE and ffi._HAS_BUFFER_RELEASE
        assert hasattr(ffi.lib, "zlsx_buffer_release")


# ─── S3b slice 10: sheet visibility on the reader handle ─────────────


def _sheet_state_workbook(path):
    """Four sheets through the writer, then `state` attributes spliced
    into `xl/workbook.xml` through the archive — the writer authors no
    sheet state: `Ledger` hidden, `Secret` veryHidden, `Odd` an
    unrecognised value the reader folds to visible, `Data` untouched
    (no attribute at all)."""
    with zlsx.write(path) as w:
        for name in ("Data", "Ledger", "Secret", "Odd"):
            w.add_sheet(name).write_row([name])
    _patch_parts(path, [
        ("xl/workbook.xml", b'name="Ledger"', b'name="Ledger" state="hidden"'),
        ("xl/workbook.xml", b'name="Secret"', b'name="Secret" state="veryHidden"'),
        ("xl/workbook.xml", b'name="Odd"', b'name="Odd" state="bogus"'),
    ])


_SHEET_STATE_EXPECTED = ["visible", "hidden", "veryHidden", "visible"]


def _require_sheet_state():
    import zlsx._ffi as ffi
    if not ffi._HAS_SHEET_STATE:
        pytest.skip("libzlsx lacks zlsx_sheet_state (0.9.0+)")


def test_book_sheet_state_reads_the_workbook_attribute(tmp_path):
    _require_sheet_state()
    path = tmp_path / "sheet_state.xlsx"
    _sheet_state_workbook(path)
    with zlsx.open(path) as book:
        # Hidden sheets stay in the inventory and read like any other.
        assert book.sheets == ["Data", "Ledger", "Secret", "Odd"]
        assert [book.sheet_state(i) for i in range(4)] == _SHEET_STATE_EXPECTED
        assert book.sheet_state("Secret") == "veryHidden"
        assert book.sheet("Ledger").state == "hidden"
        assert list(book.sheet("Secret").rows()) == [["Secret"]]
        # The selector rule is `sheet()`'s, error for error.
        with pytest.raises(IndexError):
            book.sheet_state(4)
        with pytest.raises(IndexError):
            book.sheet_state(-1)
        with pytest.raises(KeyError):
            book.sheet_state("Nope")
        with pytest.raises(TypeError):
            book.sheet_state(1.5)
        secret = book.sheet(2)
    with pytest.raises(zlsx.ZlsxError, match="closed"):
        book.sheet_state(0)
    with pytest.raises(zlsx.ZlsxError, match="closed"):
        secret.state
    # The buffer opener models the same field from the same bytes.
    with zlsx.open_bytes(path.read_bytes()) as book:
        assert [book.sheet(i).state for i in range(4)] == _SHEET_STATE_EXPECTED


def test_book_sheet_state_fresh_writer_is_visible(tmp_path):
    """The writer authors no `state` attribute; the schema default is
    what the reader reports for every sheet it wrote."""
    _require_sheet_state()
    path = tmp_path / "sheet_state_fresh.xlsx"
    with zlsx.write(path) as w:
        w.add_sheet("A").write_row([1])
        w.add_sheet("B").write_row([2])
    with zlsx.open(path) as book:
        assert [book.sheet_state(i) for i in range(2)] == ["visible", "visible"]
        assert [book.sheet(n).state for n in ("A", "B")] == ["visible", "visible"]


def test_book_sheet_state_matches_the_cli_spelling(tmp_path):
    """`zlsx list-sheets` prints `SheetState.toString()` of the field
    this getter reports — the same spelling, sheet for sheet. Runs only
    where a local CLI build sits beside the dylib."""
    import json
    import subprocess

    _require_sheet_state()
    # `zig build` installs the CLI beside the dylib; CI's pytest lane
    # (windows-runtime) builds it too, as `zlsx.exe`.
    candidates = [REPO_ROOT / "zig-out" / "bin" / name for name in ("zlsx", "zlsx.exe")]
    cli = next((c for c in candidates if c.exists()), None)
    if cli is None:
        pytest.skip("no local zlsx CLI build at zig-out/bin/zlsx")
    path = tmp_path / "sheet_state_cli.xlsx"
    _sheet_state_workbook(path)
    out = subprocess.run(
        [str(cli), "list-sheets", str(path)], check=True, capture_output=True, encoding="utf-8"
    ).stdout
    records = [json.loads(line) for line in out.splitlines() if line]
    with zlsx.open(path) as book:
        assert [(r["sheet_idx"], r["sheet"], r["state"]) for r in records] == [
            (i, name, book.sheet_state(i)) for i, name in enumerate(book.sheets)
        ]
    assert [r["state"] for r in records] == _SHEET_STATE_EXPECTED


def test_book_sheet_state_defensive_branches(tmp_path, monkeypatch):
    """The two branches a healthy 0.9.0 dylib never takes: a code the
    binding does not know (a newer library) is a `ZlsxError`, an older
    dylib without the export is a `RuntimeError` on both the method and
    the property — and the selector rule runs before either probe."""
    import zlsx._ffi as ffi

    _require_sheet_state()
    path = tmp_path / "sheet_state_defensive.xlsx"
    with zlsx.write(path) as w:
        w.add_sheet("A").write_row([1])
    with zlsx.open(path) as book:
        monkeypatch.setattr(ffi.lib, "zlsx_sheet_state", lambda handle, idx: 7)
        with pytest.raises(zlsx.ZlsxError, match="returned 7"):
            book.sheet_state(0)
        # A -1 past the binding's own bound keeps the selector contract.
        monkeypatch.setattr(ffi.lib, "zlsx_sheet_state", lambda handle, idx: -1)
        with pytest.raises(IndexError):
            book.sheet_state("A")
        monkeypatch.setattr(ffi, "_HAS_SHEET_STATE", False)
        with pytest.raises(RuntimeError, match="0.9.0"):
            book.sheet_state(0)
        with pytest.raises(RuntimeError, match="0.9.0"):
            book.sheet(0).state
        with pytest.raises(IndexError):
            book.sheet_state(3)
        with pytest.raises(KeyError):
            book.sheet_state("Nope")


def test_sheet_state_probe_agrees_with_the_library_version():
    """A dylib at or past 0.9.0 exports `zlsx_sheet_state`; a probe that
    says otherwise is a packaging error, not a reason to skip the block
    above (the S3a release-symbol probe test's precedent)."""
    import zlsx._ffi as ffi

    major, minor = (int(part) for part in ffi.lib.zlsx_version_string().decode("utf-8").split(".")[:2])
    if (major, minor) >= (0, 9):
        assert ffi._HAS_SHEET_STATE, "libzlsx >= 0.9.0 must export zlsx_sheet_state"


# ─── S3b slice 11: formula text and error tags on the row iterator ───

_FORMULA_ROWS = (
    b'<row r="2">'
    b'<c r="A2"><f>A1*2</f><v>2</v></c>'
    b'<c r="B2" t="str"><f>"x"&amp;"y"&lt;&gt;A1</f><v>xy</v></c>'
    b'<c r="C2"><f>A1/0</f></c>'
    b'</row>'
    b'<row r="3">'
    b'<c r="A3"><f t="shared" ref="A3:B3" si="0">A2+1</f><v>3</v></c>'
    b'<c r="B3"><f t="shared" si="0"/><v>4</v></c>'
    b'<c r="C3" t="e"><v>#DIV/0!</v></c>'
    b'<c r="D3" t="e"><f>1/0</f><v>#DIV/0!</v></c>'
    b'</row>'
    b'<row r="4">'
    b'<c r="A4"><f t="array" ref="A4:B4">A1*{1,2}</f><v>1</v></c>'
    b'<c r="B4"><v>2</v></c>'
    b'<c r="C4" t="e"><v>#N/A</v></c>'
    b'</row>'
    b'<row r="5">'
    b'<c r="B5" t="e"><v>#REF!</v></c>'
    b'</row>'
    b'<row r="6">'
    b'<c r="A6"><f></f><v>5</v></c>'
    b'<c r="B6"><f t="shared" si="7"/><v>8</v></c>'
    b'<c r="C6"><f t="dataTable" ref="C6:C7" r1="A1">x</f><v>7</v></c>'
    b'</row>'
)

# A row the reader cannot finish (`<c r="B7"` with no `>`): the fixture
# has formula spreads, so a skip decodes through `next()` and tears here.
_FORMULA_TORN_TAIL = b'<row r="7"><c r="A7"><f>Z9</f><v>9</v></c><c r="B7"'


def _formula_workbook(path, tail=b""):
    """One row through the writer (A1 = 1), then rows 2–6 spliced into
    the sheet part before `</sheetData>` — the writer authors neither
    shared formulas nor `t="e"` cells: a stand-alone formula, an
    entity-bearing one, a formula-only cell, a shared base + slave, an
    error cell, a formula whose cached value is an error, an array base
    + slave, a gap before an error cell, an empty `<f></f>`, a slave
    whose base was never seen, a `t="dataTable"` formula (the
    src/c_abi.zig fixture)."""
    with zlsx.write(path) as w:
        w.add_sheet("Data").write_row([1])
    _patch_parts(path, [
        ("xl/worksheets/sheet1.xml", b"</sheetData>", _FORMULA_ROWS + tail + b"</sheetData>"),
    ])


# (values, formula_strings, formula_refs, error_strings) per row — the
# lists the three accessors return beside the row `next()` yields.
_FORMULA_EXPECTED = [
    ([1], [None], [None], [None]),
    ([2, "xy", None], ["A1*2", '"x"&"y"<>A1', "A1/0"], [None] * 3, [None] * 3),
    (
        [3, 4, "#DIV/0!", "#DIV/0!"],
        ["A2+1", None, None, "1/0"],
        [None, zlsx.CellRef(0, 3), None, None],
        [None, None, "#DIV/0!", None],
    ),
    ([1, 2, "#N/A"], ["A1*{1,2}", None, None], [None, zlsx.CellRef(0, 4), None], [None, None, "#N/A"]),
    ([None, "#REF!"], [None, None], [None, None], [None, "#REF!"]),
    # An empty body is own text of length 0; a slave with no base is a
    # value cell; a dataTable body is own text like any other.
    ([5, 8, 7], ["", None, "x"], [None] * 3, [None] * 3),
]


def _require_rows_formulas():
    import zlsx._ffi as ffi
    if not ffi._HAS_ROWS_FORMULAS:
        pytest.skip("libzlsx lacks the row formula / error getters (0.9.0+)")


def _read_with_side_channels(rows):
    return [
        (row, rows.formula_strings(), rows.formula_refs(), rows.error_strings())
        for row in rows
    ]


def test_rows_formula_and_error_side_channels(tmp_path):
    _require_rows_formulas()
    path = tmp_path / "formulas.xlsx"
    _formula_workbook(path)
    with zlsx.open(path) as book:
        with book.sheet(0).rows() as rows:
            # No current row before the first `next()`.
            assert rows.formula_strings() == []
            assert rows.formula_refs() == []
            assert rows.error_strings() == []
            assert _read_with_side_channels(rows) == _FORMULA_EXPECTED
            # Past the end there is no current row.
            assert rows.formula_strings() == []
            assert rows.formula_refs() == []
            assert rows.error_strings() == []
            assert rows.style_indices() == []
            assert rows.parse_date(0) is None
        # `skip` clears the current row.
        with book.sheet(0).rows() as rows:
            next(rows)
            assert rows.formula_strings() == [None]
            assert rows.skip(1) == 1
            assert rows.formula_strings() == []
            assert rows.formula_refs() == []
            assert rows.error_strings() == []
            assert next(rows) == [3, 4, "#DIV/0!", "#DIV/0!"]
            assert rows.formula_refs() == [None, zlsx.CellRef(0, 3), None, None]
        # A skip that fails leaves no current row either: the library
        # empties its view before it reads, and the binding zeroes its
        # length before it raises — a torn row never shows through the
        # previous row's shape.
        torn = tmp_path / "formulas_torn.xlsx"
        _formula_workbook(torn, tail=_FORMULA_TORN_TAIL)
        with book.sheet(0).rows() as rows:
            assert next(rows) == [1]
        with zlsx.open(torn) as torn_book:
            with torn_book.sheet(0).rows() as rows:
                next(rows)
                assert next(rows) == [2, "xy", None]
                assert rows.formula_strings() == ["A1*2", '"x"&"y"<>A1', "A1/0"]
                with pytest.raises(zlsx.ZlsxError, match="MalformedXml"):
                    rows.skip(10)
                assert rows.formula_strings() == []
                assert rows.formula_refs() == []
                assert rows.error_strings() == []
                assert rows.style_indices() == []
                assert rows.parse_date(0) is None
            # `next()` itself on the torn row: the same error, the same
            # empty answers — the `-1` half of `__next__`'s zeroing,
            # driven directly (in-house r3 S3B-MNT-303).
            with torn_book.sheet(0).rows() as rows:
                for _ in range(6):
                    next(rows)
                assert rows.formula_strings() == ["", None, "x"]
                with pytest.raises(zlsx.ZlsxError, match="MalformedXml"):
                    next(rows)
                assert rows.formula_strings() == []
                assert rows.formula_refs() == []
                assert rows.error_strings() == []
                assert rows.style_indices() == []
                assert rows.parse_date(0) is None
        # The value list is what it always was: cached values and the
        # error literal as a plain str.
        _, data = book.sheet(0).read_all()
        assert data == [values for values, _, _, _ in _FORMULA_EXPECTED]
    # A closed iterator refuses before any probe.
    with zlsx.open(path) as book:
        rows = book.sheet(0).rows()
        rows.close()
        for accessor in (rows.formula_strings, rows.formula_refs, rows.error_strings):
            with pytest.raises(zlsx.ZlsxError):
                accessor()
    # The buffer opener reads the same fields from the same bytes.
    with zlsx.open_bytes(path.read_bytes()) as book:
        with book.sheet(0).rows() as rows:
            assert _read_with_side_channels(rows) == _FORMULA_EXPECTED


def test_rows_formula_fresh_writer_has_no_side_channels(tmp_path):
    _require_rows_formulas()
    path = tmp_path / "formulas_fresh.xlsx"
    with zlsx.write(path) as w:
        # The gap in the middle is a positional empty cell to the
        # reader; a trailing None is not written at all.
        w.add_sheet("Data").write_row([1, None, "two", 3.5, True])
    with zlsx.open(path) as book:
        with book.sheet(0).rows() as rows:
            row = next(rows)
            assert row == [1, None, "two", 3.5, True]
            assert rows.formula_strings() == [None] * 5
            assert rows.formula_refs() == [None] * 5
            assert rows.error_strings() == [None] * 5


def _a1_to_ref(a1):
    col = 0
    i = 0
    while i < len(a1) and a1[i].isalpha():
        col = col * 26 + (ord(a1[i].upper()) - ord("A") + 1)
        i += 1
    return zlsx.CellRef(col - 1, int(a1[i:]))


def test_rows_formula_matches_the_cli_records(tmp_path):
    """`zlsx cells` prints `t:"formula"` (`formula` / `formula_ref`) and
    `t:"error"` (`v`) from the fields these accessors report — the same
    text, the same base, the same literal, cell for cell, and a plain
    tag everywhere else. Runs only where a local CLI build sits beside
    the dylib."""
    import json
    import subprocess

    _require_rows_formulas()
    candidates = [REPO_ROOT / "zig-out" / "bin" / name for name in ("zlsx", "zlsx.exe")]
    cli = next((c for c in candidates if c.exists()), None)
    if cli is None:
        pytest.skip("no local zlsx CLI build at zig-out/bin/zlsx")
    path = tmp_path / "formulas_cli.xlsx"
    _formula_workbook(path)
    out = subprocess.run(
        [str(cli), "cells", str(path), "--include-blanks"],
        check=True, capture_output=True, encoding="utf-8",
    ).stdout
    records = [json.loads(line) for line in out.splitlines() if line]
    assert all(r["kind"] == "cell" for r in records)
    cli_view = {
        (r["row"], r["col"] - 1): (
            r["t"],
            r.get("formula"),
            _a1_to_ref(r["formula_ref"]) if "formula_ref" in r else None,
            r["v"] if r["t"] == "error" else None,
            # The value beside the tag: `cached` on a formula record
            # (absent for a formula-only cell), `v` elsewhere — the
            # row element the binding yields for that cell.
            r.get("cached") if r["t"] == "formula" else r["v"],
        )
        for r in records
    }
    py_view = {}
    with zlsx.open(path) as book:
        with book.sheet(0).rows() as rows:
            for row_idx, row in enumerate(rows, start=1):
                for col, (value, formula, ref, err) in enumerate(
                    zip(row, rows.formula_strings(), rows.formula_refs(), rows.error_strings())
                ):
                    tag = "formula" if (formula is not None or ref is not None) else (
                        "error" if err is not None else "plain"
                    )
                    py_view[(row_idx, col)] = (tag, formula, ref, err, value)
    assert set(cli_view) == set(py_view)
    for key, (tag, formula, ref, err, value) in py_view.items():
        cli_tag, cli_formula, cli_ref, cli_err, cli_value = cli_view[key]
        if tag == "plain":
            assert cli_tag not in ("formula", "error"), key
        else:
            assert cli_tag == tag, key
        assert (cli_formula, cli_ref, cli_err, cli_value) == (formula, ref, err, value), key
    assert sum(1 for t, *_ in py_view.values() if t == "formula") == 10
    assert sum(1 for t, *_ in py_view.values() if t == "error") == 3


def test_rows_formula_defensive_branches(tmp_path, monkeypatch):
    """An older dylib without the trio is a `RuntimeError` on each
    accessor — after the closed-iterator check, so a closed iterator is
    the same `ZlsxError` whatever dylib is loaded."""
    import zlsx._ffi as ffi

    _require_rows_formulas()
    path = tmp_path / "formulas_defensive.xlsx"
    _formula_workbook(path)
    with zlsx.open(path) as book:
        with book.sheet(0).rows() as rows:
            next(rows)
            monkeypatch.setattr(ffi, "_HAS_ROWS_FORMULAS", False)
            for accessor in (rows.formula_strings, rows.formula_refs, rows.error_strings):
                with pytest.raises(RuntimeError):
                    accessor()
            rows.close()
            for accessor in (rows.formula_strings, rows.formula_refs, rows.error_strings):
                with pytest.raises(zlsx.ZlsxError):
                    accessor()


def test_rows_formulas_probe_agrees_with_the_library_version():
    """A dylib at or past 0.9.0 exports the trio; a probe that says
    otherwise is a packaging error, not a reason to skip the block
    above (the sheet_state precedent)."""
    import zlsx._ffi as ffi

    major, minor = (int(part) for part in ffi.lib.zlsx_version_string().decode("utf-8").split(".")[:2])
    if (major, minor) >= (0, 9):
        assert ffi._HAS_ROWS_FORMULAS, "libzlsx >= 0.9.0 must export zlsx_rows_formula_at / _formula_ref_at / _error_at"


def test_rows_parse_date_answers_for_the_current_row_only(tmp_path):
    """`parse_date` is the fifth per-column getter on the handle and
    shares the trio's rule: nothing before the first `next()`, after
    `skip()` (the fast path, which leaves the library's own row buffer
    holding the date) and past the end (in-house r1 S3B-REL-103/104)."""
    import datetime as _dt
    import zlsx._ffi as ffi

    if not ffi._HAS_PARSE_DATE:
        pytest.skip("loaded libzlsx predates parse_date ABI (0.2.6+)")
    path = tmp_path / "parse_date_current_row.xlsx"
    with zlsx.write(path) as w:
        date_style = w.add_style(zlsx.Style(number_format="yyyy-mm-dd"))
        sheet = w.add_sheet("S")
        sheet.write_row([44927, "one"], styles=[date_style, 0])
        sheet.write_row(["two", "two", "two"])
        sheet.write_row(["three"])
    with zlsx.open(path) as book:
        with book.sheet(0).rows() as rows:
            assert rows.parse_date(0) is None
            next(rows)
            assert rows.parse_date(0) == _dt.datetime(2023, 1, 1)
            assert rows.parse_date(1) is None
            assert rows.parse_date(-1) is None
            # A zero-length skip is a no-op: the row stays current.
            assert rows.skip(0) == 0
            assert rows.parse_date(0) == _dt.datetime(2023, 1, 1)
            assert rows.skip(1) == 1
            assert rows.parse_date(0) is None
            assert next(rows) == ["three"]
            assert rows.parse_date(0) is None
            with pytest.raises(StopIteration):
                next(rows)
            assert rows.parse_date(0) is None


def test_rows_skip_drain_fallback_leaves_no_current_row(tmp_path, monkeypatch):
    """The pre-0.8.0 fallback drains rows through `next()`; the drained
    rows were never yielded, so afterwards there is no current row —
    the same answer the library path gives (in-house r2 S3B-REL-202)."""
    import datetime as _dt
    import zlsx._ffi as ffi

    if not ffi._HAS_PARSE_DATE:
        pytest.skip("loaded libzlsx predates parse_date ABI (0.2.6+)")
    path = tmp_path / "skip_fallback.xlsx"
    with zlsx.write(path) as w:
        date_style = w.add_style(zlsx.Style(number_format="yyyy-mm-dd"))
        sheet = w.add_sheet("S")
        sheet.write_row([44927, "one"], styles=[date_style, 0])
        sheet.write_row(["two", "two", "two"])
        sheet.write_row(["three"])
    monkeypatch.setattr(ffi, "_HAS_ROWS_SKIP", False)
    with zlsx.open(path) as book:
        with book.sheet(0).rows() as rows:
            next(rows)
            assert rows.parse_date(0) == _dt.datetime(2023, 1, 1)
            assert rows.skip(0) == 0
            assert rows.style_indices() == [date_style, None]
            assert rows.skip(1) == 1
            assert rows.style_indices() == []
            assert rows.parse_date(0) is None
            assert next(rows) == ["three"]
            assert rows.skip(5) == 0
            assert rows.style_indices() == []
