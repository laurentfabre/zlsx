"""E5 — Python embedding surface.

The tests are organised around the three states rather than the
methods, because the states are the contract. Some spreadsheet
applications rebuild the .xlsx archive on save and delete the vector
parts outright; a small recovery record survives that, so a workbook
which lost its vectors can still report what it held. A binding that
could not distinguish "stripped" from "never had any" would hand the
caller silence, which is exactly what the design rejects.

Fixtures come from ``zig build emb4-tools``, which produces the same
generator the compat matrices use. Tests skip when it has not been
built rather than failing, so a pure-Python checkout stays green.
"""

from __future__ import annotations

import shutil
import subprocess
import zipfile
from pathlib import Path

import pytest

import zlsx

REPO_ROOT = Path(__file__).resolve().parents[3]
FIXTURE_GEN = REPO_ROOT / "zig-out" / "bin" / "zlsx-emb4b-fixture"
EMB4_GEN = REPO_ROOT / "zig-out" / "bin" / "zlsx-emb4-fixture"

pytestmark = pytest.mark.skipif(
    not EMB4_GEN.exists(),
    reason="run `zig build emb4-tools` to build the embedding fixture generator",
)


@pytest.fixture(scope="module")
def embedded(tmp_path_factory) -> Path:
    """A workbook with vectors: model emb-4-fixture-v1, dim 4, 2 coverages."""
    out = tmp_path_factory.mktemp("emb") / "present.xlsx"
    subprocess.run([str(EMB4_GEN), str(out)], check=True, capture_output=True)
    return out


@pytest.fixture(scope="module")
def stripped(embedded: Path, tmp_path_factory) -> Path:
    """The same workbook with every embedding part deleted.

    Rebuilt directly rather than by driving LibreOffice: the archive
    surgery reproduces exactly what its export filter does — parts gone,
    workbook rel gone, recovery record untouched — without needing the
    application installed on the machine running the tests.
    """
    out = tmp_path_factory.mktemp("emb") / "stripped.xlsx"
    zin = zipfile.ZipFile(embedded)
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zo:
        for name in zin.namelist():
            if name.startswith("xl/zlsxEmbeddings/"):
                continue
            data = zin.read(name)
            if name == "xl/_rels/workbook.xml.rels":
                text = data.decode()
                # Drop the workbook→index relationship the same way a
                # rebuilding consumer does.
                import re

                text = re.sub(r"<Relationship[^>]*zlsxEmbeddings[^>]*/>", "", text)
                data = text.encode()
            zo.writestr(name, data)
    return out


# ── present ──────────────────────────────────────────────────────────


def test_present_reports_provenance(embedded: Path):
    with zlsx.embeddings(embedded) as e:
        assert e.present and not e.stripped and not e.absent
        assert e.state == "present"
        assert e.model == "emb-4-fixture-v1"
        assert e.dim == 4
        assert e.dtype == "int8-sym-per-vec"
        assert [c.id for c in e.coverages] == ["title", "body"]
        assert e.coverages[0].range == "A2:A5"
        assert e.coverages[0].rows == 4


def test_vectors_shape_and_dtype(embedded: Path):
    np = pytest.importorskip("numpy")
    with zlsx.embeddings(embedded) as e:
        v = e.vectors("title")
        assert v.shape == (4, 4)
        assert v.dtype == np.float32


def test_vectors_dequantize_correctly(embedded: Path):
    """The fixture uses scale 0.5 with q=[10,-10,5,-5], so row 0 must be
    0.5*q/127. Pins that the Zig-side decode is actually applied rather
    than raw bytes being reinterpreted."""
    pytest.importorskip("numpy")
    with zlsx.embeddings(embedded) as e:
        row = e.vectors("title")[0]
        expected = [0.5 * q / 127.0 for q in (10, -10, 5, -5)]
        for got, want in zip(row, expected):
            assert abs(got - want) < 1e-6


def test_coverage_addressable_by_name_and_index(embedded: Path):
    pytest.importorskip("numpy")
    with zlsx.embeddings(embedded) as e:
        assert (e.vectors("title") == e.vectors(0)).all()
        assert (e.vectors("body") == e.vectors(1)).all()


def test_unknown_coverage_raises(embedded: Path):
    pytest.importorskip("numpy")
    with zlsx.embeddings(embedded) as e:
        with pytest.raises(KeyError):
            e.vectors("nope")
        with pytest.raises(IndexError):
            e.vectors(99)


def test_valid_mask_all_live_when_no_tombstones(embedded: Path):
    pytest.importorskip("numpy")
    with zlsx.embeddings(embedded) as e:
        mask = e.valid_mask("title")
        assert mask.shape == (4,)
        assert mask.all()


def test_hashes_length_matches_rows(embedded: Path):
    pytest.importorskip("numpy")
    with zlsx.embeddings(embedded) as e:
        assert e.hashes("title").shape == (4,)


def test_digest_and_carrier_are_none_when_present(embedded: Path):
    with zlsx.embeddings(embedded) as e:
        assert e.digest is None
        assert e.carrier is None


# ── stripped ─────────────────────────────────────────────────────────


def test_stripped_recovers_provenance(stripped: Path):
    """The load-bearing case: vectors destroyed, workbook still says
    what it held."""
    with zlsx.embeddings(stripped) as e:
        assert e.stripped and not e.present and not e.absent
        assert e.state == "stripped"
        assert e.model == "emb-4-fixture-v1"
        assert e.dim == 4
        assert e.dtype == "int8-sym-per-vec"
        assert [c.id for c in e.coverages] == ["title", "body"]
        assert e.coverages[0].range == "A2:A5"
        assert e.coverages[0].rows == 4


def test_stripped_exposes_digest_and_carrier(stripped: Path):
    with zlsx.embeddings(stripped) as e:
        assert isinstance(e.digest, int)
        assert e.digest != 0
        assert e.carrier in ("defined_name", "doc_props")


def test_stripped_vectors_raise_rather_than_return_empty(stripped: Path):
    """An empty array would be the silent-nothing this design exists to
    prevent; the exception must carry the provenance instead."""
    pytest.importorskip("numpy")
    with zlsx.embeddings(stripped) as e:
        with pytest.raises(zlsx.EmbeddingsStripped) as ei:
            e.vectors("title")
        msg = str(ei.value)
        assert "emb-4-fixture-v1" in msg
        assert "Re-embed" in msg


def test_stripped_hashes_raise_too(stripped: Path):
    pytest.importorskip("numpy")
    with zlsx.embeddings(stripped) as e:
        with pytest.raises(zlsx.EmbeddingsStripped):
            e.hashes("title")


def test_embeddings_stripped_is_a_zlsx_error(stripped: Path):
    """So `except zlsx.ZlsxError` keeps catching everything from this
    module, and callers can narrow when they want to."""
    assert issubclass(zlsx.EmbeddingsStripped, zlsx.ZlsxError)


# ── absent ───────────────────────────────────────────────────────────


def test_absent_on_a_plain_workbook(tmp_path: Path):
    out = tmp_path / "plain.xlsx"
    with zlsx.write(out) as w:
        s = w.add_sheet("S")
        s.write_row(["a", "b"])
    with zlsx.embeddings(out) as e:
        assert e.absent and not e.present and not e.stripped
        assert e.state == "absent"
        assert e.coverages == []
        assert e.model == ""


def test_absent_vectors_raise_plain_error(tmp_path: Path):
    pytest.importorskip("numpy")
    out = tmp_path / "plain2.xlsx"
    with zlsx.write(out) as w:
        s = w.add_sheet("S")
        s.write_row(["a"])
    with zlsx.embeddings(out) as e:
        # Not EmbeddingsStripped — nothing was stripped.
        with pytest.raises(zlsx.ZlsxError) as ei:
            e.vectors()
        assert not isinstance(ei.value, zlsx.EmbeddingsStripped)


# ── lifecycle ────────────────────────────────────────────────────────


def test_context_manager_closes(embedded: Path):
    e = zlsx.embeddings(embedded)
    with e:
        assert e.present
    with pytest.raises(zlsx.ZlsxError):
        _ = e.model


def test_double_close_is_safe(embedded: Path):
    e = zlsx.embeddings(embedded)
    e.close()
    e.close()


def test_open_missing_file_raises(tmp_path: Path):
    with pytest.raises(zlsx.ZlsxError):
        zlsx.embeddings(tmp_path / "nope.xlsx")


def test_repr_states(embedded: Path, stripped: Path):
    with zlsx.embeddings(embedded) as e:
        assert "present" in repr(e)
    with zlsx.embeddings(stripped) as e:
        assert "stripped" in repr(e)


@pytest.mark.skipif(
    not FIXTURE_GEN.exists(), reason="emb4b fixture generator not built"
)
def test_carrier_falls_back_to_doc_props(tmp_path: Path):
    """With the defined names removed, the record must still come back
    from docProps — the two carriers exist because their removal
    mechanisms do not overlap."""
    import re

    src = tmp_path / "src.xlsx"
    subprocess.run([str(EMB4_GEN), str(src)], check=True, capture_output=True)
    out = tmp_path / "noname.xlsx"
    zin = zipfile.ZipFile(src)
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zo:
        for name in zin.namelist():
            if name.startswith("xl/zlsxEmbeddings/"):
                continue
            data = zin.read(name)
            if name == "xl/workbook.xml":
                text = re.sub(
                    r"<definedName[^>]*_zlsxRecovery[^>]*>.*?</definedName>",
                    "",
                    data.decode(),
                    flags=re.S,
                )
                data = text.encode()
            zo.writestr(name, data)
    with zlsx.embeddings(out) as e:
        assert e.stripped
        assert e.carrier == "doc_props"
        assert e.model == "emb-4-fixture-v1"
