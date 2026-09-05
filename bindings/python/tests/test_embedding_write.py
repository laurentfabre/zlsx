"""S3c slice 1 — the embedding write on the editor handle.

``Editor.set_embeddings`` is ``Workbook.setEmbeddings`` through
``zlsx_editor_set_embeddings``: one call writes the whole set, and a
later ``zlsx.embeddings(path)`` reports ``present`` with the
provenance, vectors and hashes the call handed over. No fixture
generator is involved — every workbook here comes from ``zlsx.Writer``
— so the file runs wherever the dylib does. The vector / hash
read-backs need NumPy, as the read side's ``vectors()`` / ``hashes()``
do; the state, provenance and error tests do not.
"""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest

import zlsx

VECS = [[0.5, -1.25, 2.0], [0.0, 0.0, 0.0], [3.5, 4.0, -0.75]]
HASHES = [0x1111, None, 0x3333]
TITLE = {"id": "title", "sheet": 0, "range": "A2:A4", "column": "A", "vectors": VECS, "hashes": HASHES}


def _needs_write():
    import zlsx._ffi as ffi

    if not ffi._HAS_EMBEDDING_WRITE:
        pytest.skip("loaded libzlsx predates zlsx_editor_set_embeddings (0.9.0+)")
    return ffi


def _write_fixture(path: Path) -> None:
    """Two sheets: three text rows under a header on Docs, one on Second."""
    with zlsx.Writer(path) as w:
        docs = w.add_sheet("Docs")
        docs.write_row(["title", "body"])
        docs.write_row(["alpha", "first body"])
        docs.write_row(["beta", "second body"])
        docs.write_row(["gamma", "third body"])
        second = w.add_sheet("Second")
        second.write_row(["two"])


def _cov(**over) -> dict:
    c = dict(TITLE)
    c.update(over)
    return c


# ── the round trip ───────────────────────────────────────────────────


def test_set_embeddings_round_trips_to_present(tmp_path):
    """The goal line: ``embeddings()`` is ``present`` after the write,
    with the model / dim / dtype / coverages the call named, on two
    sheets, and the cells untouched (no hidden recovery sheet)."""
    _needs_write()
    src = tmp_path / "src.xlsx"
    out = tmp_path / "out.xlsx"
    _write_fixture(src)
    with zlsx.embeddings(src) as before:
        assert before.absent

    with zlsx.Editor(src) as ed:
        ed.set_embeddings(
            "test-model/v1",
            3,
            [
                TITLE,
                {
                    "id": "second", "sheet": 1, "range": "A1:A1", "column": "A",
                    "vectors": [[9, 8, 7]], "hashes": [0xABCD], "include_formulas": True,
                },
            ],
        )
        ed.save(out)

    with zlsx.embeddings(out) as emb:
        assert emb.present
        assert emb.state == "present"
        assert not emb.stripped and not emb.absent
        assert emb.model == "test-model/v1"
        assert emb.dim == 3
        assert emb.dtype == "f32"
        assert [(c.id, c.sheet, c.range, c.rows) for c in emb.coverages] == [
            ("title", "worksheets/sheet1.xml", "A2:A4", 3),
            ("second", "worksheets/sheet2.xml", "A1:A1", 1),
        ]
        assert emb.digest is None and emb.carrier is None

    _header, rows = zlsx.read(out)
    assert rows[0] == ["title", "body"] and rows[1][0] == "alpha" and rows[3][1] == "third body"
    names = zipfile.ZipFile(out).namelist()
    assert "xl/worksheets/sheet3.xml" not in names
    assert "xl/zlsxEmbeddings/index.xml" in names
    assert "xl/zlsxEmbeddings/title/vec.bin" in names
    assert "xl/zlsxEmbeddings/second/hashes.bin" in names
    # Both invisible carriers of the recovery record.
    assert "docProps/custom.xml" in names
    assert b"_zlsxRecovery" in zipfile.ZipFile(out).read("xl/workbook.xml")


def test_set_embeddings_vectors_and_hashes_read_back_on_one_shape(tmp_path):
    """What ``vectors()`` / ``hashes()`` return is what the write took:
    a (rows, dim) float32 matrix and a uint64 per row, ``None`` being
    the tombstone ``valid_mask`` reads."""
    np = pytest.importorskip("numpy")
    _needs_write()
    src = tmp_path / "src.xlsx"
    out = tmp_path / "out.xlsx"
    _write_fixture(src)
    matrix = np.array(VECS, dtype=np.float32)
    with zlsx.Editor(src) as ed:
        ed.set_embeddings("m", 3, [_cov(vectors=matrix)])
        ed.save(out)
    with zlsx.embeddings(out) as emb:
        assert emb.present
        got = emb.vectors("title")
        assert got.shape == (3, 3) and got.dtype == np.float32
        assert np.array_equal(got, matrix)
        hashes = emb.hashes("title")
        tomb = np.uint64(zlsx._ffi.lib.zlsx_emb_tombstone())
        assert hashes.tolist() == [0x1111, int(tomb), 0x3333]
        assert emb.valid_mask("title").tolist() == [True, False, True]
        # Read → re-embed → write on the read side's own arrays.
        with zlsx.Editor(out) as ed:
            ed.set_embeddings("m2", 3, [_cov(vectors=got, hashes=hashes)])
            ed.save(src)
    with zlsx.embeddings(src) as again:
        assert again.present and again.model == "m2"
        assert np.array_equal(again.vectors("title"), matrix)
        assert again.valid_mask("title").tolist() == [True, False, True]


def test_set_embeddings_int8_sym_quantizes_in_the_library(tmp_path):
    np = pytest.importorskip("numpy")
    _needs_write()
    src = tmp_path / "src.xlsx"
    out = tmp_path / "out.xlsx"
    _write_fixture(src)
    vecs = np.array([[1.0, -2.5, 0.25, 4.0], [0, 0, 0, 0], [100, -100, 50, 0.5]], dtype=np.float32)
    with zlsx.Editor(src) as ed:
        ed.set_embeddings("q", 4, [_cov(vectors=vecs, hashes=[1, 2, 3])], dtype="int8-sym-per-vec")
        ed.save(out)
    with zlsx.embeddings(out) as emb:
        assert emb.present and emb.dtype == "int8-sym-per-vec" and emb.dim == 4
        got = emb.vectors("title")
        for want, have in zip(vecs, got):
            step = float(np.abs(want).max()) / 127.0
            assert np.abs(want - have).max() <= step + 1e-6
        assert got[1].tolist() == [0.0, 0.0, 0.0, 0.0]
    # The compact layout on disk: 24-byte header + 3 × (4-byte scale + 4 codes).
    assert len(zipfile.ZipFile(out).read("xl/zlsxEmbeddings/title/vec.bin")) == 24 + 3 * 8


def test_set_embeddings_second_write_replaces_the_set(tmp_path):
    _needs_write()
    src = tmp_path / "src.xlsx"
    mid = tmp_path / "mid.xlsx"
    out = tmp_path / "out.xlsx"
    _write_fixture(src)
    with zlsx.Editor(src) as ed:
        ed.set_embeddings("m1", 3, [TITLE, _cov(id="body", range="B2:B4", column="B")])
        ed.set_embeddings("m2", 3, [_cov(id="body", range="B2:B2", column="B", vectors=[[1, 2, 3]], hashes=[7])])
        ed.save(mid)
    with zlsx.embeddings(mid) as emb:
        assert emb.present and emb.model == "m2"
        assert [(c.id, c.rows) for c in emb.coverages] == [("body", 1)]
    with zlsx.Editor(mid) as ed:
        ed.set_embeddings("m3", 3, [_cov(id="title", sheet=1, range="A1:A1", vectors=[[1, 2, 3]], hashes=[7])])
        ed.save(out)
    with zlsx.embeddings(out) as emb:
        assert emb.present and emb.model == "m3"
        assert [(c.id, c.sheet) for c in emb.coverages] == [("title", "worksheets/sheet2.xml")]
    z = zipfile.ZipFile(out)
    assert z.read("xl/_rels/workbook.xml.rels").count(b"zlsxEmbeddings/index.xml") == 1
    assert z.read("docProps/custom.xml").count(b"<property ") == 1


# ── refusals and call errors ─────────────────────────────────────────


@pytest.mark.parametrize(
    "name, kwargs",
    [
        ("InvalidEmbeddingInput", dict(dim=0)),
        ("InvalidEmbeddingInput", dict(coverages=[])),
        ("InvalidEmbeddingInput", dict(coverages=[_cov(vectors=VECS[:2], hashes=HASHES[:2])])),
        ("InvalidEmbeddingInput", dict(coverages=[_cov(hashes=[1, 2])])),
        ("InvalidDtype", dict(dtype="float32")),
        ("UnsupportedDtype", dict(dtype="binary16")),
        ("UnsupportedDtype", dict(dtype="int8-asym-per-vec")),
        ("SheetIndexOutOfRange", dict(coverages=[_cov(sheet=5)])),
        ("InvalidRange", dict(coverages=[_cov(range="A0:A2")])),
        ("InvalidRange", dict(coverages=[_cov(column="Z")])),
        ("InvalidCoverageId", dict(coverages=[_cov(id="bad id")])),
        ("DuplicateCoverageId", dict(coverages=[TITLE, _cov(range="B2:B4", column="B")])),
        ("CoverageOverlap", dict(coverages=[TITLE, _cov(id="other", range="A3:A5")])),
        ("InvalidXmlByte", dict(model="m\x00")),
    ],
)
def test_set_embeddings_call_errors_are_named_and_write_nothing(tmp_path, name, kwargs):
    _needs_write()
    src = tmp_path / "src.xlsx"
    _write_fixture(src)
    call = dict(model="m", dim=3, coverages=[TITLE], dtype="f32")
    call.update(kwargs)
    with zlsx.Editor(src) as ed:
        with pytest.raises(zlsx.ZlsxError) as info:
            ed.set_embeddings(call["model"], call["dim"], call["coverages"], dtype=call["dtype"])
        assert not isinstance(info.value, zlsx.ZlsxRefusal)
        assert name in str(info.value)
        assert ed.save_to_buffer() == src.read_bytes()


def test_set_embeddings_refuses_a_workbook_the_relationship_cannot_land_in(tmp_path):
    """A rels file without ``</Relationships>`` is the workbook's
    fault: a typed refusal, and nothing written — the pre-flight fires
    before the first part."""
    _needs_write()
    src = tmp_path / "src.xlsx"
    bad = tmp_path / "bad.xlsx"
    _write_fixture(src)
    zin = zipfile.ZipFile(src)
    with zipfile.ZipFile(bad, "w", zipfile.ZIP_DEFLATED) as zo:
        for item in zin.namelist():
            data = zin.read(item)
            if item == "xl/_rels/workbook.xml.rels":
                data = data.replace(b"</Relationships>", b"")
            zo.writestr(item, data)
    with zlsx.Editor(bad) as ed:
        with pytest.raises(zlsx.ZlsxRefusal) as info:
            ed.set_embeddings("m", 3, [TITLE])
        assert info.value.error_name == "MalformedWorkbookRels"
        assert ed.save_to_buffer() == bad.read_bytes()


@pytest.mark.parametrize(
    "exc, kwargs",
    [
        (TypeError, dict(coverages=[("title",)])),
        (TypeError, dict(coverages=[_cov(id=7)])),
        (TypeError, dict(coverages=[_cov(sheet=True)])),
        (TypeError, dict(coverages=[_cov(vectors=[[True, 1, 2]] + VECS[1:])])),
        (TypeError, dict(coverages=[_cov(vectors="abc")])),
        (TypeError, dict(coverages=[_cov(hashes=[True, 1, 2])])),
        (TypeError, dict(coverages=[_cov(include_formulas=1)])),
        (TypeError, dict(coverages="title")),
        (TypeError, dict(model=b"m")),
        (TypeError, dict(dtype=3)),
        (ValueError, dict(coverages=[_cov(extra=1)])),
        (ValueError, dict(coverages=[{"id": "t"}])),
        (ValueError, dict(coverages=[_cov(vectors=[[1, 2], [3, 4, 5], [6, 7, 8]])])),
        (ValueError, dict(coverages=[_cov(hashes=[2**64, 1, 2])])),
        (ValueError, dict(coverages=[_cov(hashes=[-1, 1, 2])])),
        (ValueError, dict(coverages=[_cov(sheet=2**32)])),
        (ValueError, dict(dim=2**32)),
    ],
)
def test_set_embeddings_shapes_are_checked_in_python_first(tmp_path, exc, kwargs):
    _needs_write()
    src = tmp_path / "src.xlsx"
    _write_fixture(src)
    call = dict(model="m", dim=3, coverages=[TITLE], dtype="f32")
    call.update(kwargs)
    with zlsx.Editor(src) as ed:
        with pytest.raises(exc):
            ed.set_embeddings(call["model"], call["dim"], call["coverages"], dtype=call["dtype"])
        assert ed.save_to_buffer() == src.read_bytes()


def test_set_embeddings_numpy_width_is_checked_against_dim(tmp_path):
    np = pytest.importorskip("numpy")
    _needs_write()
    src = tmp_path / "src.xlsx"
    _write_fixture(src)
    with zlsx.Editor(src) as ed:
        with pytest.raises(ValueError):
            ed.set_embeddings("m", 3, [_cov(vectors=np.zeros((3, 4), dtype=np.float32))])
        with pytest.raises(ValueError):
            ed.set_embeddings("m", 3, [_cov(vectors=np.zeros((3, 3, 1), dtype=np.float32))])
        # A flat array is fine: the C side judges the count against the range.
        ed.set_embeddings("m", 3, [_cov(vectors=np.zeros(9, dtype=np.float32))])


def test_set_embeddings_closed_editor_and_older_dylib(tmp_path, monkeypatch):
    import zlsx._ffi as ffi

    _needs_write()
    src = tmp_path / "src.xlsx"
    _write_fixture(src)
    ed = zlsx.Editor(src)
    ed.close()
    with pytest.raises(zlsx.ZlsxError, match="closed"):
        ed.set_embeddings("m", 3, [TITLE])
    with zlsx.Editor(src) as live:
        monkeypatch.setattr(ffi, "_HAS_EMBEDDING_WRITE", False)
        with pytest.raises(RuntimeError, match="zlsx_editor_set_embeddings"):
            live.set_embeddings("m", 3, [TITLE])


def test_set_embeddings_probe_agrees_with_the_library_version():
    """A dylib at or past 0.9.0 exports `zlsx_editor_set_embeddings`; a
    probe that says otherwise is a packaging error, not a reason to skip
    the tests above (the S3b slice-10 precedent)."""
    import zlsx._ffi as ffi

    major, minor = (int(part) for part in ffi.lib.zlsx_version_string().decode("utf-8").split(".")[:2])
    if (major, minor) >= (0, 9):
        assert ffi._HAS_EMBEDDING_WRITE, "libzlsx >= 0.9.0 must export zlsx_editor_set_embeddings"
