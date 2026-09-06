"""S3c slice 2 — the embeddable-rows read on the editor handle.

``Editor.embeddable_rows`` is ``Workbook.embeddableRows`` through
``zlsx_editor_embeddable_rows_ndjson``: the rows of a column over a
range that carry embeddable content, each with the text a model should
see and the canonical xxh3-64 content hash ``set_embeddings`` stores
beside the vector — so read → embed → write is one shape, and the
hashes a caller hands back read fresh. Every workbook here comes from
``zlsx.Writer``; the ``valid_mask`` read-backs need NumPy, as the read
side's do.
"""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest

import zlsx

# The canonical hash of "alpha" at row 2 on `worksheets/sheet1.xml` —
# pinned in src/c_abi.zig on the same fixture, so a divergence between
# the surfaces (or a drift of the canonical form) fails under
# `zig build test`, not only under this advisory lane.
ALPHA_ROW2_HASH = 6830279115424181645


def _needs_read():
    import zlsx._ffi as ffi

    if not ffi._HAS_EMBEDDABLE_ROWS:
        pytest.skip("loaded libzlsx predates zlsx_editor_embeddable_rows_ndjson (0.9.0+)")
    return ffi


def _write_fixture(path: Path) -> None:
    """Two sheets: three text rows under a header on Docs, one on Second
    — the S3c fixture, the C tests' too."""
    with zlsx.Writer(path) as w:
        docs = w.add_sheet("Docs")
        docs.write_row(["title", "body"])
        docs.write_row(["alpha", "first body"])
        docs.write_row(["beta", "second body"])
        docs.write_row(["gamma", "third body"])
        second = w.add_sheet("Second")
        second.write_row(["two"])


def _patched(src: Path, dst: Path, part: str, old: bytes, new: bytes) -> Path:
    zin = zipfile.ZipFile(src)
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zo:
        for item in zin.namelist():
            data = zin.read(item)
            if item == part:
                assert data.count(old) == 1, (item, old)
                data = data.replace(old, new)
            zo.writestr(item, data)
    return dst


# ── the round trip ───────────────────────────────────────────────────


def test_embeddable_rows_fed_to_set_embeddings_read_back_valid(tmp_path):
    """The goal line: the hashes the read hands over, written by
    ``set_embeddings``, read back with ``valid_mask`` all true and equal
    to what was written — the read's hash IS the write's."""
    _needs_read()
    np = pytest.importorskip("numpy")
    src = tmp_path / "src.xlsx"
    out = tmp_path / "out.xlsx"
    _write_fixture(src)
    with zlsx.Editor(src) as ed:
        rows = ed.embeddable_rows(0, "A2:A4", "A")
        assert [r["row"] for r in rows] == [2, 3, 4]
        assert [r["text"] for r in rows] == ["alpha", "beta", "gamma"]
        hashes = [r["hash"] for r in rows]
        assert all(isinstance(h, int) and 0 <= h < 2**64 for h in hashes)
        vectors = np.arange(9, dtype=np.float32).reshape(3, 3)
        ed.set_embeddings("m", 3, [
            {"id": "title", "sheet": 0, "range": "A2:A4", "column": "A",
             "vectors": vectors, "hashes": hashes},
        ])
        ed.save(out)
    with zlsx.embeddings(out) as emb:
        assert emb.state == "present"
        assert emb.valid_mask("title").all()
        assert emb.hashes("title").tolist() == hashes
        assert emb.vectors("title").tolist() == vectors.tolist()


def test_embeddable_rows_records_are_the_cli_shape(tmp_path):
    """One ``embed_row`` record per row — the CLI's keys, the row
    1-based, the hash the pinned canonical value."""
    _needs_read()
    src = tmp_path / "src.xlsx"
    _write_fixture(src)
    with zlsx.Editor(src) as ed:
        rows = ed.embeddable_rows(0, "A2:A4", "A")
        assert len(rows) == 3
        for r in rows:
            assert set(r) == {"kind", "row", "text", "hash"}
            assert r["kind"] == "embed_row"
        assert rows[0] == {"kind": "embed_row", "row": 2, "text": "alpha", "hash": ALPHA_ROW2_HASH}
        # The second sheet by index — the write's own resolution — and
        # a range with nothing embeddable.
        assert [r["text"] for r in ed.embeddable_rows(1, "A1:A1", "A")] == ["two"]
        assert ed.embeddable_rows(0, "C2:C4", "C") == []
        assert ed.embeddable_rows(0, "A2:A4", "A", include_formulas=True) == rows


def test_embeddable_rows_every_kind_as_a_reader_sees_it(tmp_path):
    """Entities resolved, a rich string's runs joined, a number's
    ``<v>`` as written, a boolean as ``1``, a blank omitted, a formula's
    cached string only on request."""
    _needs_read()
    src = tmp_path / "kinds.xlsx"
    with zlsx.Writer(src) as w:
        s = w.add_sheet("Kinds")
        s.write_row(["title"])
        s.write_row(["a & b <c>"])
        s.write_row([1.5])
        s.write_row([None])
        s.write_row([True])
        s.write_row_with_formulas(["cached"], ["A2"])
        # Written last: the writer numbers a rich cell as it is written
        # but emits rich entries after every plain one (pre-existing,
        # recorded outside this slice).
        s.write_rich_row([[zlsx.RichRun("Hello", bold=True), zlsx.RichRun(" world")]])
    with zlsx.Editor(src) as ed:
        plain = ed.embeddable_rows(0, "A2:A7", "A")
        assert [(r["row"], r["text"]) for r in plain] == [
            (2, "a & b <c>"), (3, "1.5"), (5, "1"), (7, "Hello world"),
        ]
        with_formulas = ed.embeddable_rows(0, "A2:A7", "A", include_formulas=True)
        assert [(r["row"], r["text"]) for r in with_formulas] == [
            (2, "a & b <c>"), (3, "1.5"), (5, "1"), (6, "cached"), (7, "Hello world"),
        ]


def test_embeddable_rows_omitted_rows_become_tombstones_on_the_write(tmp_path):
    """The docstring's mapping: a covered row the read omits is a
    ``None`` hash on the write, and reads back as a tombstone."""
    _needs_read()
    np = pytest.importorskip("numpy")
    src = tmp_path / "gap.xlsx"
    out = tmp_path / "gap_out.xlsx"
    with zlsx.Writer(src) as w:
        s = w.add_sheet("Docs")
        s.write_row(["title"])
        s.write_row(["alpha"])
        s.write_row([None])
        s.write_row(["gamma"])
    with zlsx.Editor(src) as ed:
        rows = ed.embeddable_rows(0, "A2:A4", "A")
        assert [r["row"] for r in rows] == [2, 4]
        by_row = {r["row"]: r["hash"] for r in rows}
        dim = 2
        ed.set_embeddings("m", dim, [{
            "id": "title", "sheet": 0, "range": "A2:A4", "column": "A",
            "vectors": [[1.0, 2.0] if r in by_row else [0.0] * dim for r in range(2, 5)],
            "hashes": [by_row.get(r) for r in range(2, 5)],
        }])
        ed.save(out)
    with zlsx.embeddings(out) as emb:
        assert emb.valid_mask("title").tolist() == [True, False, True]
        assert emb.hashes("title").tolist() == [by_row[2], np.iinfo(np.uint64).max, by_row[4]]


# ── statements about the call ────────────────────────────────────────


@pytest.mark.parametrize(
    "name, args",
    [
        ("InvalidRange", (0, "A0:A2", "A")),
        ("InvalidRange", (0, "", "A")),
        ("InvalidRange", (0, "A2:A4", "Z")),
        ("InvalidRange", (0, "A2:A4", "")),
        ("SheetIndexOutOfRange", (2, "A2:A4", "A")),
        ("SheetIndexOutOfRange", (2**32 - 1, "A2:A4", "A")),
    ],
)
def test_embeddable_rows_call_errors_are_named(tmp_path, name, args):
    _needs_read()
    src = tmp_path / "src.xlsx"
    _write_fixture(src)
    with zlsx.Editor(src) as ed:
        with pytest.raises(zlsx.ZlsxError) as info:
            ed.embeddable_rows(*args)
        assert not isinstance(info.value, zlsx.ZlsxRefusal)
        assert f": {name}" in str(info.value)


def test_embeddable_rows_refuses_after_a_staged_cell_write(tmp_path):
    """A staged ``set_cell`` — outside the range, even — makes the sheet
    unreadable for this read (the parsed view does not carry it); the
    other sheet still answers, and the saved file does."""
    _needs_read()
    src = tmp_path / "src.xlsx"
    out = tmp_path / "out.xlsx"
    _write_fixture(src)
    with zlsx.Editor(src) as ed:
        ed.set_cell(0, 9, 3, 42)
        with pytest.raises(zlsx.ZlsxError) as info:
            ed.embeddable_rows(0, "A2:A4", "A")
        assert not isinstance(info.value, zlsx.ZlsxRefusal)
        assert ": SheetHasUnsavedMutations" in str(info.value)
        assert [r["text"] for r in ed.embeddable_rows(1, "A1:A1", "A")] == ["two"]
        ed.save(out)
        # The save commits the staged write: the same editor answers again.
        assert [r["text"] for r in ed.embeddable_rows(0, "A2:A4", "A")] == ["alpha", "beta", "gamma"]
    with zlsx.Editor(out) as ed:
        assert [r["text"] for r in ed.embeddable_rows(0, "A2:A4", "A")] == ["alpha", "beta", "gamma"]
        ed.append_rows(0, [["delta"]])
        with pytest.raises(zlsx.ZlsxError) as info:
            ed.embeddable_rows(0, "A2:A4", "A")
        assert ": SheetHasUnsavedAppends" in str(info.value)


# ── statements about the workbook ────────────────────────────────────


@pytest.mark.parametrize(
    "part, old, new, name, other_sheet_served",
    [
        ("xl/worksheets/sheet1.xml", b"</sheetData>", b"", "MalformedSheetXml", True),
        ("xl/worksheets/sheet1.xml", b't="s"><v>2</v>', b't="s"><v>99</v>', "SstIndexOutOfRange", True),
        ("xl/worksheets/sheet1.xml", b't="s"><v>2</v>', b't="b"><v>TRUE</v>', "UnsupportedCellValue", True),
        ("xl/sharedStrings.xml", b">alpha</t>", b">&bogus;</t>", "UnsupportedCellValue", True),
        ("xl/sharedStrings.xml", b">alpha</t>", b">\xff</t>", "InvalidUtf8", True),
        # A <v> the number canonicalizer cannot read, a t="d" date, a t
        # the reader does not know — the cell rule's one name, never the
        # canonicalizer's.
        ("xl/worksheets/sheet1.xml", b't="s"><v>2</v>', b"><v>1,5</v>", "UnsupportedCellValue", True),
        ("xl/worksheets/sheet1.xml", b't="s"><v>2</v>', b't="d"><v>2024-01-01</v>', "UnsupportedCellValue", True),
        ("xl/worksheets/sheet1.xml", b't="s"><v>2</v>', b't="zz"><v>42</v>', "UnsupportedCellValue", True),
        # The UTF-8 rule on the kinds the hash does not validate: an
        # error literal, a number (judged before the canonicalizer).
        ("xl/worksheets/sheet1.xml", b't="s"><v>2</v>', b't="e"><v>#N/\xff</v>', "InvalidUtf8", True),
        ("xl/worksheets/sheet1.xml", b't="s"><v>2</v>', b"><v>1\xff</v>", "InvalidUtf8", True),
        # A shared-string table the parser cannot read (the LAST entry's
        # close — an inner one is tolerated): the table is the
        # workbook's, so the other sheet refuses too.
        ("xl/sharedStrings.xml", b">two</t></si>", b">two</t>", "MalformedSharedStringsXml", False),
    ],
)
def test_embeddable_rows_refuses_a_workbook_it_cannot_serve(tmp_path, part, old, new, name, other_sheet_served):
    """A part the view cannot parse, or a cell value the read cannot
    carry: a typed refusal rather than a record that lies. A sheet
    part's verdict leaves the other sheet served; the shared-string
    table's does not."""
    _needs_read()
    src = tmp_path / "src.xlsx"
    _write_fixture(src)
    bad = _patched(src, tmp_path / "bad.xlsx", part, old, new)
    with zlsx.Editor(bad) as ed:
        with pytest.raises(zlsx.ZlsxRefusal) as info:
            ed.embeddable_rows(0, "A2:A4", "A")
        assert info.value.error_name == name
        if other_sheet_served:
            assert [r["text"] for r in ed.embeddable_rows(1, "A1:A1", "A")] == ["two"]
        else:
            with pytest.raises(zlsx.ZlsxRefusal) as other:
                ed.embeddable_rows(1, "A1:A1", "A")
            assert other.value.error_name == name


def test_embeddable_rows_carries_an_error_cell_as_its_literal(tmp_path):
    """An error cell is embeddable — its literal, the hash's kind ``e``."""
    _needs_read()
    src = tmp_path / "src.xlsx"
    _write_fixture(src)
    err = _patched(src, tmp_path / "err.xlsx", "xl/worksheets/sheet1.xml", b't="s"><v>2</v>', b't="e"><v>#N/A</v>')
    with zlsx.Editor(err) as ed:
        rows = ed.embeddable_rows(0, "A2:A4", "A")
        assert [(r["row"], r["text"]) for r in rows] == [(2, "#N/A"), (3, "beta"), (4, "gamma")]
        assert rows[0]["hash"] != ALPHA_ROW2_HASH


# ── the Python surface ───────────────────────────────────────────────


@pytest.mark.parametrize(
    "exc, args, kwargs",
    [
        (TypeError, (0, b"A2:A4", "A"), {}),
        (TypeError, (0, "A2:A4", 1), {}),
        (TypeError, (0, "A2:A4", "A"), {"include_formulas": 1}),
        (TypeError, (True, "A2:A4", "A"), {}),
        (ValueError, (-1, "A2:A4", "A"), {}),
        (ValueError, (2**32, "A2:A4", "A"), {}),
    ],
)
def test_embeddable_rows_shapes_are_checked_in_python_first(tmp_path, exc, args, kwargs):
    _needs_read()
    src = tmp_path / "src.xlsx"
    _write_fixture(src)
    with zlsx.Editor(src) as ed:
        with pytest.raises(exc):
            ed.embeddable_rows(*args, **kwargs)


def test_embeddable_rows_closed_editor_and_older_dylib(tmp_path, monkeypatch):
    ffi = _needs_read()
    src = tmp_path / "src.xlsx"
    _write_fixture(src)
    ed = zlsx.Editor(src)
    ed.close()
    with pytest.raises(zlsx.ZlsxError, match="closed"):
        ed.embeddable_rows(0, "A2:A4", "A")
    with zlsx.Editor(src) as live:
        monkeypatch.setattr(ffi, "_HAS_EMBEDDABLE_ROWS", False)
        with pytest.raises(RuntimeError, match="zlsx_editor_embeddable_rows_ndjson"):
            live.embeddable_rows(0, "A2:A4", "A")


def test_embeddable_rows_probe_agrees_with_the_library_version():
    """A dylib at or past 0.9.0 exports the read; a probe that says
    otherwise is a packaging error, not a reason to skip (the S3b
    slice-10 precedent)."""
    import zlsx._ffi as ffi

    major, minor = (int(part) for part in ffi.lib.zlsx_version_string().decode("utf-8").split(".")[:2])
    if (major, minor) >= (0, 9):
        assert ffi._HAS_EMBEDDABLE_ROWS, "libzlsx >= 0.9.0 must export zlsx_editor_embeddable_rows_ndjson"
