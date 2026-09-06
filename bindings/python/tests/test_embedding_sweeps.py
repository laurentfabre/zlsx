"""S3c slice 3 — the embedding sweeps on the editor handle.

``Editor.prune_embeddings`` is ``Workbook.pruneEmbeddings`` through
``zlsx_editor_prune_embeddings`` (the redaction sweep ``zlsx embed
--prune`` runs, its counts as a dict) and ``Editor.strip_embeddings``
is ``Workbook.stripEmbeddings`` through ``zlsx_editor_strip_embeddings``
(the pre-share operation ``zlsx embed --strip`` runs). Every workbook
here comes from ``zlsx.Writer`` and the write is ``set_embeddings`` over
``embeddable_rows``' hashes; the vector / mask read-backs need NumPy, as
the read side's do.
"""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest

import zlsx

ALL_FRESH = {"redacted": 0, "stale": 0, "fresh": 3, "valid_empty": 0}
ALL_ZERO = {"redacted": 0, "stale": 0, "fresh": 0, "valid_empty": 0}
SHEET1 = "xl/worksheets/sheet1.xml"
BETA_CELL = b'<c r="A3" t="s"><v>4</v></c>'


def _needs_sweeps():
    import zlsx._ffi as ffi

    if not ffi._HAS_EMBEDDING_SWEEPS:
        pytest.skip("loaded libzlsx predates zlsx_editor_prune_embeddings / _strip_embeddings (0.9.0+)")
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


def _write_embedded(tmp_path: Path, name: str = "emb.xlsx") -> Path:
    """The fixture with the ``title`` coverage written over the read's
    own hashes — every slot fresh — saved to ``name``."""
    src = tmp_path / f"src_{name}"
    out = tmp_path / name
    _write_fixture(src)
    with zlsx.Editor(src) as ed:
        rows = ed.embeddable_rows(0, "A2:A4", "A")
        ed.set_embeddings("m", 3, [
            {"id": "title", "sheet": 0, "range": "A2:A4", "column": "A",
             "vectors": [[1.0, 2.0, 3.0], [4.0, 5.0, 6.0], [7.0, 8.0, 9.0]],
             "hashes": [r["hash"] for r in rows]},
        ])
        ed.save(out)
    return out


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


def _dropped(src: Path, dst: Path, prefix: str) -> Path:
    """The archive without every entry under ``prefix`` — a consumer
    that rebuilds the package and drops the parts it does not know."""
    zin = zipfile.ZipFile(src)
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zo:
        for item in zin.namelist():
            if item.startswith(prefix):
                continue
            zo.writestr(item, zin.read(item))
    return dst


# ── the goal line ────────────────────────────────────────────────────


def test_hashes_from_the_read_prune_all_fresh_and_a_strip_reads_absent(tmp_path):
    """The goal line: a set written with ``embeddable_rows``' hashes
    prunes as all fresh, and a stripped file reads back ``absent``."""
    _needs_sweeps()
    emb = _write_embedded(tmp_path)
    stripped = tmp_path / "stripped.xlsx"
    with zlsx.embeddings(emb) as before:
        assert before.state == "present"
    with zlsx.Editor(emb) as ed:
        assert ed.prune_embeddings() == ALL_FRESH
        ed.strip_embeddings()
        ed.save(stripped)
    with zlsx.embeddings(stripped) as after:
        assert after.state == "absent"
        assert after.absent and not after.stripped and not after.present


# ── prune ────────────────────────────────────────────────────────────


def test_prune_report_is_the_cli_record_and_a_workbook_without_a_set_is_all_zeros(tmp_path):
    _needs_sweeps()
    src = tmp_path / "plain.xlsx"
    _write_fixture(src)
    with zlsx.Editor(src) as ed:
        report = ed.prune_embeddings()
        assert report == ALL_ZERO
        assert list(report) == ["redacted", "stale", "fresh", "valid_empty"]
        assert all(type(v) is int for v in report.values())


def test_prune_a_row_blanked_on_disk_is_redacted_and_reads_back_as_a_tombstone(tmp_path):
    """Excel's shape: the cell is gone from the saved part. Its slot
    becomes a tombstone over a zeroed vector; the next sweep calls it
    ``valid_empty``."""
    _needs_sweeps()
    np = pytest.importorskip("numpy")
    emb = _write_embedded(tmp_path)
    blanked = _patched(emb, tmp_path / "blanked.xlsx", SHEET1, BETA_CELL, b'<c r="A3"/>')
    out = tmp_path / "blanked_out.xlsx"
    with zlsx.Editor(blanked) as ed:
        assert ed.prune_embeddings() == {"redacted": 1, "stale": 0, "fresh": 2, "valid_empty": 0}
        ed.save(out)
    with zlsx.embeddings(out) as e:
        assert e.state == "present"
        assert e.valid_mask("title").tolist() == [True, False, True]
        assert int(e.hashes("title")[1]) == np.iinfo(np.uint64).max
        assert e.vectors("title").tolist() == [[1.0, 2.0, 3.0], [0.0, 0.0, 0.0], [7.0, 8.0, 9.0]]
    with zlsx.Editor(out) as ed:
        assert ed.prune_embeddings() == {"redacted": 0, "stale": 0, "fresh": 2, "valid_empty": 1}


def test_prune_an_edited_row_is_stale_and_left_alone(tmp_path):
    """Content that drifted but is still embeddable: counted, never
    redacted, and nothing rewritten — the save is the untouched
    editor's passthrough."""
    _needs_sweeps()
    emb = _write_embedded(tmp_path)
    edited = _patched(emb, tmp_path / "edited.xlsx", SHEET1, BETA_CELL, b'<c r="A3" t="s"><v>5</v></c>')
    out = tmp_path / "edited_out.xlsx"
    with zlsx.Editor(edited) as ed:
        assert ed.prune_embeddings() == {"redacted": 0, "stale": 1, "fresh": 2, "valid_empty": 0}
        ed.save(out)
    assert out.read_bytes() == edited.read_bytes()


def test_prune_judges_a_staged_set_cell_as_staged(tmp_path):
    """A blank redacts its slot; any other value is stale, never fresh
    — and the same editor's later save carries the redaction."""
    _needs_sweeps()
    emb = _write_embedded(tmp_path)
    out = tmp_path / "staged_out.xlsx"
    with zlsx.Editor(emb) as ed:
        ed.set_cell(0, 3, 0, None)
        ed.set_cell(0, 4, 0, "gamma edited")
        assert ed.prune_embeddings() == {"redacted": 1, "stale": 1, "fresh": 1, "valid_empty": 0}
        ed.save(out)
    with zlsx.Editor(out) as ed:
        assert ed.prune_embeddings() == {"redacted": 0, "stale": 1, "fresh": 1, "valid_empty": 1}


def test_prune_refuses_staged_appends_on_the_covered_sheet_only(tmp_path):
    _needs_sweeps()
    emb = _write_embedded(tmp_path)
    with zlsx.Editor(emb) as ed:
        ed.append_rows(1, [["more"]])
        assert ed.prune_embeddings() == ALL_FRESH
        ed.append_rows(0, [["more"]])
        with pytest.raises(zlsx.ZlsxError) as info:
            ed.prune_embeddings()
        assert not isinstance(info.value, zlsx.ZlsxRefusal)
        assert ": SheetHasUnsavedAppends" in str(info.value)


def test_prune_a_stripped_set_is_all_zeros(tmp_path):
    """The recovery record alone (a consumer dropped the parts): nothing
    to sweep, and a strip then clears the record too."""
    _needs_sweeps()
    emb = _write_embedded(tmp_path)
    stripped_by_tool = _dropped(emb, tmp_path / "by_tool.xlsx", "xl/zlsxEmbeddings/")
    out = tmp_path / "by_tool_out.xlsx"
    with zlsx.embeddings(stripped_by_tool) as e:
        assert e.state == "stripped"
    with zlsx.Editor(stripped_by_tool) as ed:
        assert ed.prune_embeddings() == ALL_ZERO
        ed.strip_embeddings()
        ed.save(out)
    with zlsx.embeddings(out) as e:
        assert e.state == "absent"


@pytest.mark.parametrize(
    "part, old, new, name, read_name",
    [
        # The parser's own name is what the read reports; the sweep
        # folds it under the set's.
        ("xl/zlsxEmbeddings/index.xml", b'range="A2:A4"', b'range="A0:A4"', "MalformedEmbeddingSet", "InvalidRange"),
        ("xl/zlsxEmbeddings/index.xml", b'dtype="f32"', b'dtype="f99"', "MalformedEmbeddingSet", "InvalidDtype"),
        ("xl/zlsxEmbeddings/index.xml", b'count="3"', b'count="4"', "MalformedEmbeddingSet", "CoverageCountMismatch"),
        ("xl/zlsxEmbeddings/title/vec.bin", b"ZVEC", b"ZVEX", "MalformedEmbeddingSet", "BadMagic"),
        ("xl/zlsxEmbeddings/_rels/index.xml.rels", b'Target="title/vec.bin"', b'Target="title/none.bin"', "MissingEmbeddingPart", "MissingEmbeddingPart"),
        # The read's own verdicts on a covered cell stop the sweep whole.
        (SHEET1, b'<c r="A3"', b"<c", "MalformedSheetXml", None),
        (SHEET1, b'<c r="A3" t="s"><v>4</v>', b'<c r="A3" t="b"><v>TRUE</v>', "UnsupportedCellValue", None),
        ("xl/sharedStrings.xml", b">beta</t>", b">&bogus;</t>", "UnsupportedCellValue", None),
    ],
)
def test_prune_refuses_a_workbook_it_cannot_sweep(tmp_path, part, old, new, name, read_name):
    """A typed refusal before the first part write — the save after it
    is the passthrough — and the read side keeps the parser's names."""
    _needs_sweeps()
    emb = _write_embedded(tmp_path)
    bad = _patched(emb, tmp_path / "bad.xlsx", part, old, new)
    out = tmp_path / "bad_out.xlsx"
    with zlsx.Editor(bad) as ed:
        with pytest.raises(zlsx.ZlsxRefusal) as info:
            ed.prune_embeddings()
        assert info.value.error_name == name
        ed.save(out)
    assert out.read_bytes() == bad.read_bytes()
    if read_name is not None:
        with pytest.raises(zlsx.ZlsxError) as read_info:
            zlsx.embeddings(bad)
        assert read_name in str(read_info.value)


# ── strip ────────────────────────────────────────────────────────────


def test_strip_removes_every_carrier_and_is_idempotent(tmp_path):
    _needs_sweeps()
    emb = _write_embedded(tmp_path)
    out = tmp_path / "stripped.xlsx"
    again = tmp_path / "stripped_again.xlsx"
    with zlsx.Editor(emb) as ed:
        ed.strip_embeddings()
        ed.strip_embeddings()
        ed.save(out)
    names = zipfile.ZipFile(out).namelist()
    assert not any(n.startswith("xl/zlsxEmbeddings/") for n in names)
    assert "docProps/custom.xml" not in names
    z = zipfile.ZipFile(out)
    assert b"_zlsxRecovery" not in z.read("xl/workbook.xml")
    assert b"zlsxEmbeddings" not in z.read("xl/_rels/workbook.xml.rels")
    with zlsx.embeddings(out) as e:
        assert e.state == "absent"
    with zlsx.Editor(out) as ed:
        # The cells are what they were, and a strip on a workbook
        # without a set is a no-op: the save is the passthrough.
        assert [r["text"] for r in ed.embeddable_rows(0, "A2:A4", "A")] == ["alpha", "beta", "gamma"]
        ed.strip_embeddings()
        ed.save(again)
    assert again.read_bytes() == out.read_bytes()


def test_strip_on_a_workbook_that_never_had_embeddings_changes_nothing(tmp_path):
    _needs_sweeps()
    src = tmp_path / "plain.xlsx"
    out = tmp_path / "plain_out.xlsx"
    _write_fixture(src)
    with zlsx.Editor(src) as ed:
        ed.strip_embeddings()
        ed.save(out)
    assert out.read_bytes() == src.read_bytes()


def test_strip_refuses_a_workbook_xml_it_cannot_walk_before_the_first_removal(tmp_path):
    """The chunk-name strip's verdict, judged first: the vectors are
    still in the archive after the refusal (the save is the
    passthrough)."""
    _needs_sweeps()
    emb = _write_embedded(tmp_path)
    bad = _patched(emb, tmp_path / "bad_wb.xlsx", "xl/workbook.xml", b"</workbook>", b'<definedName name="x>oops</definedName></workbook>')
    out = tmp_path / "bad_wb_out.xlsx"
    with zlsx.Editor(bad) as ed:
        with pytest.raises(zlsx.ZlsxRefusal) as info:
            ed.strip_embeddings()
        assert info.value.error_name == "MalformedWorkbookXml"
        ed.save(out)
    assert out.read_bytes() == bad.read_bytes()
    assert any(n.startswith("xl/zlsxEmbeddings/") for n in zipfile.ZipFile(out).namelist())


# ── the Python surface ───────────────────────────────────────────────


def test_sweeps_closed_editor_and_older_dylib(tmp_path, monkeypatch):
    ffi = _needs_sweeps()
    src = tmp_path / "src.xlsx"
    _write_fixture(src)
    ed = zlsx.Editor(src)
    ed.close()
    with pytest.raises(zlsx.ZlsxError, match="closed"):
        ed.prune_embeddings()
    with pytest.raises(zlsx.ZlsxError, match="closed"):
        ed.strip_embeddings()
    with zlsx.Editor(src) as live:
        monkeypatch.setattr(ffi, "_HAS_EMBEDDING_SWEEPS", False)
        with pytest.raises(RuntimeError, match="zlsx_editor_prune_embeddings"):
            live.prune_embeddings()
        with pytest.raises(RuntimeError, match="zlsx_editor_strip_embeddings"):
            live.strip_embeddings()


def test_sweeps_probe_agrees_with_the_library_version():
    """A dylib at or past 0.9.0 exports both sweeps; a probe that says
    otherwise is a packaging error, not a reason to skip (the S3b
    slice-10 precedent)."""
    import zlsx._ffi as ffi

    major, minor = (int(part) for part in ffi.lib.zlsx_version_string().decode("utf-8").split(".")[:2])
    if (major, minor) >= (0, 9):
        assert ffi._HAS_EMBEDDING_SWEEPS, "libzlsx >= 0.9.0 must export zlsx_editor_prune_embeddings and zlsx_editor_strip_embeddings"
