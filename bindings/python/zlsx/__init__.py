"""py-zlsx — Python binding over the zlsx xlsx reader/writer library.

Quick start::

    import zlsx

    with zlsx.open("workbook.xlsx") as book:
        print(book.sheets)            # list[str]
        for row in book.sheet(0).rows():
            ...

The binding talks to ``libzlsx`` via ctypes — no Python interpreter floor
beyond ctypes itself (stdlib). On Homebrew, the dylib ships alongside
``brew install laurentfabre/zlsx/zlsx``; other platforms can set
``ZLSX_LIBRARY=/path/to/libzlsx.{so,dylib,dll}``.

Cell type mapping (``zlsx_cell_tag_t`` → Python):
    empty   → None
    string  → str  (UTF-8)
    integer → int  (never rounded)
    number  → float
    boolean → bool
"""

from __future__ import annotations

import ctypes
import os
import threading
import time
from datetime import datetime as _datetime
from pathlib import Path
from typing import Iterator, List, Optional, Tuple, Union

from . import _ffi

__version__ = "0.8.0"
"""Python-package version. Tracks the Zig library's major+minor; the patch
level may drift when the binding ships a Python-only fix."""

__all__ = [
    "open",
    "open_bytes",
    "write",
    "Book",
    "Sheet",
    "Rows",
    "Writer",
    "SheetWriter",
    "Style",
    "BorderSide",
    "CellRef",
    "MergeRange",
    "Hyperlink",
    "DataValidation",
    "RichRun",
    "Font",
    "Fill",
    "Border",
    "Alignment",
    "Comment",
    "Dxf",
    "CF_OPERATORS",
    "to_excel_serial",
    "read",
    "ZlsxError",
    "Editor",
    "edit",
    "Embeddings",
    "EmbeddingsStripped",
    "Coverage",
    "embeddings",
    "ZlsxFormulaRefusal",
    "FormulaSpec",
    "RecalcOptions",
    "RecalcReport",
    "CensusEntry",
    "Resolved",
    "EvalResult",
    "Matrix",
    "ExcelError",
    "engine_fingerprint",
]


# ─── Styles (Phase 3b) ─────────────────────────────────────────────────
#
# `Style` is a dataclass that mirrors the Zig writer's Style struct. Keep
# fields additive — openpyxl-parity fields land in subsequent releases
# (alignment, fills, borders, number formats, etc).

from dataclasses import dataclass, field
from typing import Literal, Optional


HAlignLiteral = Literal[
    "general", "left", "center", "right", "fill",
    "justify", "centerContinuous", "distributed",
]
PatternTypeLiteral = Literal[
    "none", "solid", "gray125", "gray0625", "darkGray", "mediumGray",
    "lightGray", "darkHorizontal", "darkVertical", "darkDown", "darkUp",
    "darkGrid", "darkTrellis", "lightHorizontal", "lightVertical",
    "lightDown", "lightUp", "lightGrid", "lightTrellis",
]
BorderStyleLiteral = Literal[
    "none", "thin", "medium", "dashed", "dotted", "thick", "double",
    "hair", "mediumDashed", "dashDot", "mediumDashDot", "dashDotDot",
    "mediumDashDotDot", "slantDashDot",
]

_HALIGN_VALUES = {
    "general": 0, "left": 1, "center": 2, "right": 3,
    "fill": 4, "justify": 5, "centerContinuous": 6, "distributed": 7,
}

_PATTERN_VALUES = {
    "none": 0, "solid": 1, "gray125": 2, "gray0625": 3,
    "darkGray": 4, "mediumGray": 5, "lightGray": 6,
    "darkHorizontal": 7, "darkVertical": 8, "darkDown": 9, "darkUp": 10,
    "darkGrid": 11, "darkTrellis": 12,
    "lightHorizontal": 13, "lightVertical": 14, "lightDown": 15,
    "lightUp": 16, "lightGrid": 17, "lightTrellis": 18,
}

_BORDER_STYLE_VALUES = {
    "none": 0, "thin": 1, "medium": 2, "dashed": 3, "dotted": 4,
    "thick": 5, "double": 6, "hair": 7, "mediumDashed": 8, "dashDot": 9,
    "mediumDashDot": 10, "dashDotDot": 11, "mediumDashDotDot": 12,
    "slantDashDot": 13,
}


@dataclass(frozen=True)
class BorderSide:
    """A single edge of a cell border. ``style="none"`` means no line;
    ``color_argb`` is optional (None = OOXML default / auto)."""
    style: BorderStyleLiteral = "none"
    color_argb: Optional[int] = None


@dataclass(frozen=True)
class Style:
    """A cell-style specification. Pass an instance to
    :meth:`Writer.add_style` and use the returned index with
    :meth:`SheetWriter.write_row`'s ``styles`` argument.

    Fields mirror the Zig Style struct. ``None`` means "unset (OOXML
    default)"; concrete values emit the corresponding XML attributes.
    Colour fields (``font_color_argb``, ``fill_fg_argb``,
    ``fill_bg_argb``) are packed ARGB (0xAARRGGBB); for fully opaque
    red use ``0xFFFF0000``.

    For a solid yellow highlight: ``Style(fill_pattern="solid",
    fill_fg_argb=0xFFFFFF00)``.
    """

    font_bold: bool = False
    font_italic: bool = False
    font_size: Optional[float] = None
    font_name: Optional[str] = None
    font_color_argb: Optional[int] = None
    alignment_horizontal: HAlignLiteral = "general"
    wrap_text: bool = False
    fill_pattern: PatternTypeLiteral = "none"
    fill_fg_argb: Optional[int] = None
    fill_bg_argb: Optional[int] = None
    # Borders — each side is a BorderSide; defaults emit nothing.
    # ``border_diagonal`` plus ``diagonal_up`` / ``diagonal_down`` control
    # the diagonal line (style gates rendering, the flags choose direction).
    border_left: "BorderSide" = field(default_factory=BorderSide)
    border_right: "BorderSide" = field(default_factory=BorderSide)
    border_top: "BorderSide" = field(default_factory=BorderSide)
    border_bottom: "BorderSide" = field(default_factory=BorderSide)
    border_diagonal: "BorderSide" = field(default_factory=BorderSide)
    diagonal_up: bool = False
    diagonal_down: bool = False
    # OOXML number format string (e.g. "0.00", "m/d/yyyy", "$#,##0.00").
    # None = General. Custom formats register at numFmtId >= 164 and
    # dedup across styles.
    number_format: Optional[str] = None


@dataclass(frozen=True)
class Dxf:
    """A differential format — font / fill / border overrides applied
    when a conditional-format rule matches. Register on the workbook
    via :meth:`Writer.add_dxf` to receive a ``dxf_id`` that
    :meth:`SheetWriter.add_conditional_format_cell_is` /
    :meth:`…_expression` can reference.

    Supported fields (iter49):
      - bold / italic / font color / font size
      - solid fill color
      - per-side borders (left / right / top / bottom)
    """
    font_bold: bool = False
    font_italic: bool = False
    font_color_argb: Optional[int] = None
    font_size: Optional[float] = None
    fill_fg_argb: Optional[int] = None
    border_left: "BorderSide" = field(default_factory=lambda: BorderSide())
    border_right: "BorderSide" = field(default_factory=lambda: BorderSide())
    border_top: "BorderSide" = field(default_factory=lambda: BorderSide())
    border_bottom: "BorderSide" = field(default_factory=lambda: BorderSide())


class ZlsxError(RuntimeError):
    """Raised when the zlsx C ABI returns an error. ``args[0]`` is the
    null-terminated diagnostic written into the error buffer by the
    library."""


_ERR_BUF_LEN = 256


def _decode_err(buf: ctypes.Array) -> str:
    return bytes(buf.value).decode("utf-8", errors="replace")


# DataValidation kind / operator code tables mirror the C ABI constants
# (see zlsx.h ZLSX_DV_KIND_* / ZLSX_DV_OP_*). Kept as plain dicts rather
# than Enums so callers can compare against simple strings.
_DV_KIND_FROM_CODE = {
    0: "list",
    1: "whole",
    2: "decimal",
    3: "date",
    4: "time",
    5: "text_length",
    6: "custom",
    7: "unknown",
}
_DV_OP_FROM_CODE = {
    0: "between",
    1: "not_between",
    2: "equal",
    3: "not_equal",
    4: "less_than",
    5: "less_than_or_equal",
    6: "greater_than",
    7: "greater_than_or_equal",
    # 0xFFFFFFFF → None (handled by dict.get returning None)
}


def _read_dv_formula(fn, handle, sheet_idx: int, dv_idx: int) -> str:
    """Call a `zlsx_data_validation_formulaN` getter and return the
    decoded string ("" on -1 or zero length). Shared by formula1 /
    formula2 paths."""
    ptr = ctypes.POINTER(ctypes.c_ubyte)()
    length = ctypes.c_size_t(0)
    rc = fn(handle, sheet_idx, dv_idx, ctypes.byref(ptr), ctypes.byref(length))
    if rc != 0 or length.value == 0:
        return ""
    return ctypes.string_at(ptr, length.value).decode("utf-8", errors="replace")


def _check_argb(name: str, value) -> int:
    """Validate an ARGB colour is in the u32 range. ctypes.c_uint32
    would silently mask a value like 0x1FFFFFFFF into 0xFFFFFFFF — a
    user typo that'd ship the wrong colour without warning. Range-
    check upfront and raise a ValueError that names the field."""
    if value is None:
        return 0
    v = int(value)
    if v < 0 or v > 0xFFFFFFFF:
        raise ValueError(
            f"{name} must be a 32-bit ARGB integer in [0, 0xFFFFFFFF], got {value!r}"
        )
    return v


def _cell_to_py(cell: _ffi.Cell) -> Union[None, str, int, float, bool]:
    tag = cell.tag
    if tag == _ffi.CELL_EMPTY:
        return None
    if tag == _ffi.CELL_STRING:
        if cell.str_len == 0:
            return ""
        raw = ctypes.string_at(cell.str_ptr, cell.str_len)
        return raw.decode("utf-8", errors="replace")
    if tag == _ffi.CELL_INTEGER:
        return cell.i
    if tag == _ffi.CELL_NUMBER:
        return cell.f
    if tag == _ffi.CELL_BOOLEAN:
        return bool(cell.b)
    # Defensive fallback for an ABI bump that adds a tag we don't know.
    return None


# ─── Reader metadata dataclasses ──────────────────────────────────────
#
# Mirror the Zig public types: column is 0-based (A=0), row is 1-based
# (row1=1). Immutable because Book.merged_ranges / Book.hyperlinks
# returns views into the library's internal buffers — any mutation
# users do here shouldn't leak back into other callers' view.


@dataclass(frozen=True)
class CellRef:
    """A1-style cell reference as `(col, row)`. ``col`` is 0-based
    (A=0, B=1, …); ``row`` is 1-based (row 1 is the first row)."""
    col: int
    row: int


@dataclass(frozen=True)
class MergeRange:
    """A rectangular merged cell range. ``top_left`` is component-wise
    ≤ ``bottom_right``; both corners are inclusive."""
    top_left: CellRef
    bottom_right: CellRef


@dataclass(frozen=True)
class Hyperlink:
    """An external-URL or internal-target hyperlink attached to a
    cell or cell range. ``url`` is the resolved ``Target`` from the
    sheet's rels file with XML entities decoded (``&amp;`` → ``&``,
    ``&apos;`` → ``'``, etc.) — ready to use as-is. ``location``
    carries the internal target (e.g. ``"Sheet2!A1"``) for
    ``<hyperlink location="…"/>`` entries (also decoded); it is
    empty for external links. Exactly one of ``url`` / ``location``
    is non-empty for any well-formed source. Requires libzlsx
    0.2.7+ to populate ``location`` — older dylibs leave it
    empty; libzlsx 0.2.10+ decodes XML entities (older dylibs
    surfaced the raw escaped form)."""
    top_left: CellRef
    bottom_right: CellRef
    url: str
    location: str = ""


@dataclass(frozen=True)
class DataValidation:
    """A data validation (dropdown / numeric / date / time / text-length
    / custom) on a cell or range.

    ``kind`` is one of ``"list"``, ``"whole"``, ``"decimal"``,
    ``"date"``, ``"time"``, ``"text_length"``, ``"custom"``, or
    ``"unknown"`` (forward-compat with generators that introduce new
    types).

    ``op`` is one of ``"between"``, ``"not_between"``, ``"equal"``,
    ``"not_equal"``, ``"less_than"``, ``"less_than_or_equal"``,
    ``"greater_than"``, ``"greater_than_or_equal"``, or ``None`` when
    the source had no ``operator=`` attribute (list / custom
    validations, or numeric with an omitted operator).

    ``values`` is populated for list-kind validations only (parsed
    from the literal quoted CSV in ``formula1``). Range-reference
    lists (``$A$1:$A$10``) come through as an empty tuple — callers
    can still read ``formula1`` to resolve the range themselves.

    ``formula1`` / ``formula2`` hold the entity-decoded formula
    content for non-list validations. ``formula2`` is populated only
    for ``between`` / ``not_between`` operators. All strings are
    decoded (``R&D`` not ``R&amp;D``)."""
    top_left: CellRef
    bottom_right: CellRef
    values: tuple[str, ...]
    kind: str = "list"
    op: str | None = None
    formula1: str = ""
    formula2: str = ""


@dataclass(frozen=True)
class RichRun:
    """A single formatting run inside a shared-string entry. Excel
    emits rich-text via ``<si><r><rPr/>...<t/></r>...</si>`` where
    every ``<r>`` can carry its own font properties.

    ``color_argb`` is the ARGB color from ``<color rgb="AARRGGBB"/>``
    or ``None`` when the run uses a theme color (not resolved today)
    or no color at all. ``size`` is in points, ``None`` when absent.
    ``font_name`` is ``""`` when the run had no ``<rFont val="…"/>``.
    The color / size / font fields require libzlsx 0.2.6+ — on older
    libraries they stay at their defaults."""
    text: str
    bold: bool = False
    italic: bool = False
    color_argb: int | None = None
    size: float | None = None
    font_name: str = ""


@dataclass(frozen=True)
class Font:
    """Cell-level font properties resolved via ``Book.cell_font(style_idx)``.
    Shape mirrors :class:`RichRun` minus ``text``. Theme colors aren't
    resolved — only explicit ``<color rgb="AARRGGBB"/>`` populates
    ``color_argb``."""
    bold: bool = False
    italic: bool = False
    color_argb: int | None = None
    size: float | None = None
    name: str = ""


@dataclass(frozen=True)
class Fill:
    """Cell-level fill properties resolved via ``Book.cell_fill(style_idx)``.
    ``pattern`` is the OOXML patternType attribute (``"none"``,
    ``"solid"``, ``"darkDown"``, …). ``fg_color_argb`` /
    ``bg_color_argb`` are ``None`` when the source used a theme or
    indexed color (not resolved today)."""
    pattern: str = "none"
    fg_color_argb: int | None = None
    bg_color_argb: int | None = None


@dataclass(frozen=True)
class Comment:
    """A cell comment / note parsed from ``xl/comments*.xml``.
    ``top_left`` points at the commented cell. ``author`` resolves
    through the ``<authors>`` table; ``text`` is always the
    concatenated plain-text body. ``runs`` is populated when the
    source body used ``<r><rPr>`` formatting (iter53); ``None`` for
    plain-text comments — the zero-overhead common case. All
    strings are entity-decoded."""
    top_left: CellRef
    author: str
    text: str
    runs: "tuple[RichRun, ...] | None" = None


@dataclass(frozen=True)
class Border:
    """Cell border resolved via ``Book.cell_border(style_idx)``.
    Every side is always present; absent sides have ``style=""``.
    Reuses the writer-side :class:`BorderSide` so read and write
    round-trip through the same type."""
    left: "BorderSide" = field(default_factory=lambda: BorderSide())
    right: "BorderSide" = field(default_factory=lambda: BorderSide())
    top: "BorderSide" = field(default_factory=lambda: BorderSide())
    bottom: "BorderSide" = field(default_factory=lambda: BorderSide())
    diagonal: "BorderSide" = field(default_factory=lambda: BorderSide())


@dataclass(frozen=True)
class Alignment:
    """Cell alignment resolved via ``Book.cell_alignment(style_idx)``.

    ``horizontal`` mirrors the OOXML ``<alignment horizontal="…">``
    enum value (``"left"``, ``"center"``, ``"right"``, etc.). When
    the cell uses the default ``"general"`` alignment (which the
    OOXML emitter omits), ``horizontal`` is the empty string.
    """
    horizontal: str = ""
    wrap_text: bool = False


# ─── Book ─────────────────────────────────────────────────────────────


class Book:
    """A workbook handle. Use :func:`zlsx.open` to construct one.

    Also usable as a context manager; exit closes the handle::

        with zlsx.open("file.xlsx") as book:
            ...
    """

    def __init__(self, path: Union[str, Path]):
        self._handle = None
        self._err = ctypes.create_string_buffer(_ERR_BUF_LEN)
        path_bytes = str(path).encode("utf-8")
        handle = _ffi.lib.zlsx_book_open(path_bytes, self._err, _ERR_BUF_LEN)
        if not handle:
            raise ZlsxError(f"zlsx_book_open({path!r}): {_decode_err(self._err)}")
        self._attach(handle)

    def _attach(self, handle) -> None:
        """Adopt an already-open C handle and cache sheet names — most
        callers enumerate them, and the list is short (<10 in typical
        workbooks). Shared by the path and buffer constructors."""
        self._handle = handle
        count = _ffi.lib.zlsx_sheet_count(self._handle)
        self.sheets: list[str] = []
        name_buf = ctypes.create_string_buffer(256)
        for i in range(count):
            full = _ffi.lib.zlsx_sheet_name(self._handle, i, name_buf, len(name_buf))
            if full >= len(name_buf):
                # Library reported a longer name than fit — grow and retry.
                name_buf = ctypes.create_string_buffer(full + 1)
                _ffi.lib.zlsx_sheet_name(self._handle, i, name_buf, len(name_buf))
            self.sheets.append(name_buf.value.decode("utf-8", errors="replace"))

    def sheet(self, selector: Union[int, str]) -> "Sheet":
        """Select a sheet by 0-based index or by name."""
        if not self._handle:
            raise ZlsxError("Book is closed")
        if isinstance(selector, int):
            if selector < 0 or selector >= len(self.sheets):
                raise IndexError(
                    f"sheet index {selector} out of range (workbook has {len(self.sheets)})"
                )
            return Sheet(self, selector)
        if isinstance(selector, str):
            name_bytes = selector.encode("utf-8")
            idx = _ffi.lib.zlsx_sheet_index_by_name(
                self._handle, name_bytes, len(name_bytes)
            )
            if idx < 0:
                raise KeyError(f"no sheet named {selector!r}")
            return Sheet(self, idx)
        raise TypeError(
            f"sheet selector must be int or str, got {type(selector).__name__}"
        )

    def merged_ranges(self, sheet_idx: int) -> list[MergeRange]:
        """Merged cell ranges declared in sheet ``sheet_idx``'s
        ``<mergeCells>`` block. Returns an empty list for sheets
        without merges."""
        if not self._handle:
            raise ZlsxError("book is closed")
        if not _ffi._HAS_READER_META:
            raise RuntimeError(
                "loaded libzlsx does not expose merged_ranges (requires 0.2.5+); "
                "upgrade libzlsx"
            )
        count = _ffi.lib.zlsx_merged_range_count(self._handle, sheet_idx)
        out: list[MergeRange] = []
        mr = _ffi.CMergeRange()
        for i in range(count):
            rc = _ffi.lib.zlsx_merged_range_at(self._handle, sheet_idx, i, ctypes.byref(mr))
            if rc != 0:
                # Defensive: count/at-index race shouldn't happen, but skip
                # gracefully rather than surface an internal error.
                continue
            out.append(MergeRange(
                top_left=CellRef(col=mr.top_left_col, row=mr.top_left_row),
                bottom_right=CellRef(col=mr.bottom_right_col, row=mr.bottom_right_row),
            ))
        return out

    def hyperlinks(self, sheet_idx: int) -> list[Hyperlink]:
        """Hyperlinks declared on sheet ``sheet_idx``, resolved through
        the sheet's ``_rels/sheet{N}.xml.rels`` file. Both external
        (``url``) and internal (``location``) targets are returned;
        for any well-formed source exactly one is non-empty. Returns
        an empty list for sheets without a ``<hyperlinks>`` block.
        Requires libzlsx 0.2.7+ to populate ``location``; older dylibs
        return ``location=""`` for every entry."""
        if not self._handle:
            raise ZlsxError("book is closed")
        if not _ffi._HAS_READER_META:
            raise RuntimeError(
                "loaded libzlsx does not expose hyperlinks (requires 0.2.5+); "
                "upgrade libzlsx"
            )
        count = _ffi.lib.zlsx_hyperlink_count(self._handle, sheet_idx)
        out: list[Hyperlink] = []
        hl = _ffi.CHyperlink()
        loc_ptr = ctypes.POINTER(ctypes.c_ubyte)()
        loc_len = ctypes.c_size_t(0)
        for i in range(count):
            rc = _ffi.lib.zlsx_hyperlink_at(self._handle, sheet_idx, i, ctypes.byref(hl))
            if rc != 0:
                continue
            url = ctypes.string_at(hl.url_ptr, hl.url_len).decode("utf-8", errors="replace")
            location = ""
            if _ffi._HAS_HYPERLINK_LOCATION:
                rc_loc = _ffi.lib.zlsx_hyperlink_location_at(
                    self._handle, sheet_idx, i,
                    ctypes.byref(loc_ptr), ctypes.byref(loc_len),
                )
                if rc_loc == 0 and loc_len.value > 0:
                    location = ctypes.string_at(loc_ptr, loc_len.value).decode("utf-8", errors="replace")
            out.append(Hyperlink(
                top_left=CellRef(col=hl.top_left_col, row=hl.top_left_row),
                bottom_right=CellRef(col=hl.bottom_right_col, row=hl.bottom_right_row),
                url=url,
                location=location,
            ))
        return out

    def comments(self, sheet_idx: int) -> list[Comment]:
        """Cell comments declared on sheet ``sheet_idx`` (from
        ``xl/comments*.xml`` discovered via the sheet's rels).
        Returns an empty list for sheets without a comments part.
        Requires libzlsx 0.2.6+."""
        if not self._handle:
            raise ZlsxError("book is closed")
        if not _ffi._HAS_COMMENTS:
            raise RuntimeError(
                "loaded libzlsx does not expose comments "
                "(requires 0.2.6+); upgrade libzlsx"
            )
        count = _ffi.lib.zlsx_comment_count(self._handle, sheet_idx)
        out: list[Comment] = []
        cc = _ffi.CComment()
        for i in range(count):
            rc = _ffi.lib.zlsx_comment_at(self._handle, sheet_idx, i, ctypes.byref(cc))
            if rc != 0:
                continue
            author = ""
            if cc.author_len > 0:
                author = ctypes.string_at(cc.author_ptr, cc.author_len).decode(
                    "utf-8", errors="replace"
                )
            text = ""
            if cc.text_len > 0:
                text = ctypes.string_at(cc.text_ptr, cc.text_len).decode(
                    "utf-8", errors="replace"
                )
            # iter53: surface rich-text runs when the comment body
            # used <r><rPr> formatting. Plain-text bodies have 0 runs
            # → runs=None so callers don't see a misleading tuple.
            runs_tuple: "tuple[RichRun, ...] | None" = None
            if _ffi._HAS_COMMENT_RUNS:
                rcount = _ffi.lib.zlsx_comment_run_count(self._handle, sheet_idx, i)
                if rcount > 0:
                    rlist: list[RichRun] = []
                    text_ptr = ctypes.POINTER(ctypes.c_ubyte)()
                    text_len = ctypes.c_size_t(0)
                    bold = ctypes.c_uint8(0)
                    italic = ctypes.c_uint8(0)
                    for r in range(rcount):
                        rc2 = _ffi.lib.zlsx_comment_run_at(
                            self._handle, sheet_idx, i, r,
                            ctypes.byref(text_ptr),
                            ctypes.byref(text_len),
                            ctypes.byref(bold),
                            ctypes.byref(italic),
                        )
                        if rc2 != 0:
                            continue
                        rlist.append(RichRun(
                            text=ctypes.string_at(text_ptr, text_len.value).decode(
                                "utf-8", errors="replace"
                            ),
                            bold=bool(bold.value),
                            italic=bool(italic.value),
                        ))
                    runs_tuple = tuple(rlist)
            out.append(Comment(
                top_left=CellRef(col=cc.cell_col, row=cc.cell_row),
                author=author,
                text=text,
                runs=runs_tuple,
            ))
        return out

    def data_validations(self, sheet_idx: int) -> list[DataValidation]:
        """Data validations on ``sheet_idx`` (dropdowns + numeric / date
        / time / text-length / custom). Empty list for sheets without a
        ``<dataValidations>`` block. Extended fields (``kind``, ``op``,
        ``formula1``, ``formula2``) require libzlsx 0.2.6+; on older
        libraries they fall back to the list-only defaults."""
        if not self._handle:
            raise ZlsxError("book is closed")
        if not _ffi._HAS_READER_DV:
            raise RuntimeError(
                "loaded libzlsx does not expose data_validations "
                "(requires 0.2.5+); upgrade libzlsx"
            )
        count = _ffi.lib.zlsx_data_validation_count(self._handle, sheet_idx)
        out: list[DataValidation] = []
        dv = _ffi.CDataValidation()
        for i in range(count):
            rc = _ffi.lib.zlsx_data_validation_at(
                self._handle, sheet_idx, i, ctypes.byref(dv)
            )
            if rc != 0:
                continue
            vals: list[str] = []
            vptr = ctypes.POINTER(ctypes.c_ubyte)()
            vlen = ctypes.c_size_t(0)
            for vi in range(dv.values_count):
                vrc = _ffi.lib.zlsx_data_validation_value_at(
                    self._handle, sheet_idx, i, vi,
                    ctypes.byref(vptr), ctypes.byref(vlen),
                )
                if vrc != 0:
                    continue
                vals.append(
                    ctypes.string_at(vptr, vlen.value).decode("utf-8", errors="replace")
                )
            kind = "list"
            op: str | None = None
            f1 = ""
            f2 = ""
            if _ffi._HAS_READER_DV_EXT:
                kind_code = _ffi.lib.zlsx_data_validation_kind(
                    self._handle, sheet_idx, i
                )
                kind = _DV_KIND_FROM_CODE.get(kind_code, "unknown")
                op_code = _ffi.lib.zlsx_data_validation_operator(
                    self._handle, sheet_idx, i
                )
                op = _DV_OP_FROM_CODE.get(op_code)
                f1 = _read_dv_formula(
                    _ffi.lib.zlsx_data_validation_formula1,
                    self._handle, sheet_idx, i,
                )
                f2 = _read_dv_formula(
                    _ffi.lib.zlsx_data_validation_formula2,
                    self._handle, sheet_idx, i,
                )
            out.append(DataValidation(
                top_left=CellRef(col=dv.top_left_col, row=dv.top_left_row),
                bottom_right=CellRef(col=dv.bottom_right_col, row=dv.bottom_right_row),
                values=tuple(vals),
                kind=kind,
                op=op,
                formula1=f1,
                formula2=f2,
            ))
        return out

    def shared_strings_count(self) -> int:
        """Total number of shared-string entries in the workbook.
        Returns 0 when the workbook has no ``xl/sharedStrings.xml``
        part (small xlsx files with only inline strings).

        Pair with :meth:`shared_string_at` + :meth:`rich_text` to
        enumerate every entry and discover which indices carry
        rich-text runs. Requires libzlsx 0.2.6+."""
        if not self._handle:
            raise ZlsxError("book is closed")
        if not _ffi._HAS_SST_ENUM:
            raise RuntimeError(
                "loaded libzlsx does not expose shared_strings_count "
                "(requires 0.2.6+); upgrade libzlsx"
            )
        return _ffi.lib.zlsx_shared_string_count(self._handle)

    def shared_string_at(self, sst_idx: int) -> str:
        """Return shared-string entry ``sst_idx`` as a decoded UTF-8
        ``str``. Raises :class:`IndexError` on out-of-range.
        Requires libzlsx 0.2.6+."""
        if not self._handle:
            raise ZlsxError("book is closed")
        if not _ffi._HAS_SST_ENUM:
            raise RuntimeError(
                "loaded libzlsx does not expose shared_string_at "
                "(requires 0.2.6+); upgrade libzlsx"
            )
        out_ptr = ctypes.POINTER(ctypes.c_ubyte)()
        out_len = ctypes.c_size_t(0)
        rc = _ffi.lib.zlsx_shared_string_at(
            self._handle, sst_idx, ctypes.byref(out_ptr), ctypes.byref(out_len)
        )
        if rc != 0:
            raise IndexError(f"sst_idx {sst_idx} out of range")
        if out_len.value == 0:
            return ""
        return ctypes.string_at(out_ptr, out_len.value).decode("utf-8", errors="replace")

    def shared_strings(self) -> list[str]:
        """Materialise every shared-string entry into a Python list.
        Each element is the entry's plain-text form (rich-text runs
        are concatenated into the same string by the parser — pair
        with :meth:`rich_text` to get formatting back).

        Prefer :meth:`shared_string_at` + :meth:`shared_strings_count`
        when iterating a large SST to avoid materialising the full
        list. Requires libzlsx 0.2.6+."""
        count = self.shared_strings_count()
        return [self.shared_string_at(i) for i in range(count)]

    def rich_text(self, sst_idx: int) -> list[RichRun] | None:
        """Rich-text runs for shared-string entry ``sst_idx``. Returns
        ``None`` for plain single-run strings (no ``<r>`` wrappers in
        the source XML — the common case, zero overhead). Returns a
        list of :class:`RichRun` for multi-run entries.

        SST indices can be discovered via iteration over cells: when a
        ``Cell`` is a string and you want to know if it was formatted,
        look up the corresponding SST index. Today that mapping isn't
        exposed — use this against arbitrary SST indices during
        exploration or when you've tracked the index yourself. A
        future iter will attach runs directly to string cells.

        Requires libzlsx 0.2.6+."""
        if not self._handle:
            raise ZlsxError("book is closed")
        if not _ffi._HAS_RICH_RUNS:
            raise RuntimeError(
                "loaded libzlsx does not expose rich_text "
                "(requires 0.2.6+); upgrade libzlsx"
            )
        count = _ffi.lib.zlsx_rich_run_count(self._handle, sst_idx)
        if count == 0:
            return None
        out: list[RichRun] = []
        text_ptr = ctypes.POINTER(ctypes.c_ubyte)()
        text_len = ctypes.c_size_t(0)
        bold = ctypes.c_uint8(0)
        italic = ctypes.c_uint8(0)
        color = ctypes.c_uint32(0)
        size = ctypes.c_float(0.0)
        font_ptr = ctypes.POINTER(ctypes.c_ubyte)()
        font_len = ctypes.c_size_t(0)
        for i in range(count):
            rc = _ffi.lib.zlsx_rich_run_at(
                self._handle, sst_idx, i,
                ctypes.byref(text_ptr),
                ctypes.byref(text_len),
                ctypes.byref(bold),
                ctypes.byref(italic),
            )
            if rc != 0:
                continue
            color_val: int | None = None
            size_val: float | None = None
            font_val = ""
            if _ffi._HAS_RICH_RUNS_EXT:
                crc = _ffi.lib.zlsx_rich_run_color(
                    self._handle, sst_idx, i, ctypes.byref(color)
                )
                if crc == 0:
                    color_val = int(color.value)
                src = _ffi.lib.zlsx_rich_run_size(
                    self._handle, sst_idx, i, ctypes.byref(size)
                )
                if src == 0:
                    size_val = float(size.value)
                frc = _ffi.lib.zlsx_rich_run_font_name(
                    self._handle, sst_idx, i,
                    ctypes.byref(font_ptr), ctypes.byref(font_len),
                )
                if frc == 0 and font_len.value > 0:
                    font_val = ctypes.string_at(
                        font_ptr, font_len.value
                    ).decode("utf-8", errors="replace")
            out.append(RichRun(
                text=ctypes.string_at(text_ptr, text_len.value).decode("utf-8", errors="replace"),
                bold=bool(bold.value),
                italic=bool(italic.value),
                color_argb=color_val,
                size=size_val,
                font_name=font_val,
            ))
        return out

    def number_format(self, style_idx: int) -> str | None:
        """Resolve a cell's style index (from ``Rows.style_indices()``)
        to its number-format code. Returns ``None`` on out-of-range
        indices or when the workbook has no ``xl/styles.xml``. Custom
        codes are whatever the source file declared; built-in ids
        decode to their canonical patterns (e.g. ``14`` →
        ``"m/d/yyyy"``). Requires libzlsx 0.2.6+."""
        if not self._handle:
            raise ZlsxError("book is closed")
        if not _ffi._HAS_NUM_FMT:
            raise RuntimeError(
                "loaded libzlsx does not expose number_format "
                "(requires 0.2.6+); upgrade libzlsx"
            )
        out_ptr = ctypes.POINTER(ctypes.c_ubyte)()
        out_len = ctypes.c_size_t(0)
        rc = _ffi.lib.zlsx_number_format(
            self._handle, style_idx, ctypes.byref(out_ptr), ctypes.byref(out_len)
        )
        if rc != 0:
            return None
        return ctypes.string_at(out_ptr, out_len.value).decode("utf-8", errors="replace")

    def cell_font(self, style_idx: int) -> Font | None:
        """Resolve a cell's style index to its :class:`Font` properties
        (bold / italic / color / size / name). Returns ``None`` on
        out-of-range indices or workbooks without ``xl/styles.xml``.
        Requires libzlsx 0.2.6+."""
        if not self._handle:
            raise ZlsxError("book is closed")
        if not _ffi._HAS_CELL_FONT:
            raise RuntimeError(
                "loaded libzlsx does not expose cell_font "
                "(requires 0.2.6+); upgrade libzlsx"
            )
        cf = _ffi.CCellFont()
        rc = _ffi.lib.zlsx_cell_font(self._handle, style_idx, ctypes.byref(cf))
        if rc != 0:
            return None
        name = ""
        if cf.name_len > 0:
            name = ctypes.string_at(cf.name_ptr, cf.name_len).decode(
                "utf-8", errors="replace"
            )
        return Font(
            bold=bool(cf.bold),
            italic=bool(cf.italic),
            color_argb=int(cf.color_argb) if cf.has_color else None,
            size=float(cf.size) if cf.has_size else None,
            name=name,
        )

    def cell_fill(self, style_idx: int) -> Fill | None:
        """Resolve a cell's style index to its :class:`Fill`
        (pattern + fg/bg ARGB). Returns ``None`` on out-of-range
        indices or workbooks without ``xl/styles.xml``. An all-defaults
        fill (``pattern="none"``, no colors) is still a non-None
        return. Requires libzlsx 0.2.6+."""
        if not self._handle:
            raise ZlsxError("book is closed")
        if not _ffi._HAS_CELL_FILL:
            raise RuntimeError(
                "loaded libzlsx does not expose cell_fill "
                "(requires 0.2.6+); upgrade libzlsx"
            )
        cf = _ffi.CCellFill()
        rc = _ffi.lib.zlsx_cell_fill(self._handle, style_idx, ctypes.byref(cf))
        if rc != 0:
            return None
        pattern = "none"
        if cf.pattern_len > 0:
            pattern = ctypes.string_at(cf.pattern_ptr, cf.pattern_len).decode(
                "utf-8", errors="replace"
            )
        return Fill(
            pattern=pattern,
            fg_color_argb=int(cf.fg_color_argb) if cf.has_fg else None,
            bg_color_argb=int(cf.bg_color_argb) if cf.has_bg else None,
        )

    def cell_border(self, style_idx: int) -> Border | None:
        """Resolve a cell's style index to its :class:`Border`
        (left/right/top/bottom/diagonal sides). Returns ``None`` on
        out-of-range indices or workbooks without ``xl/styles.xml``.
        Requires libzlsx 0.2.6+."""
        if not self._handle:
            raise ZlsxError("book is closed")
        if not _ffi._HAS_CELL_BORDER:
            raise RuntimeError(
                "loaded libzlsx does not expose cell_border "
                "(requires 0.2.6+); upgrade libzlsx"
            )
        cb = _ffi.CCellBorder()
        rc = _ffi.lib.zlsx_cell_border(self._handle, style_idx, ctypes.byref(cb))
        if rc != 0:
            return None

        def _side(s: _ffi.CBorderSide) -> BorderSide:
            style = ""
            if s.style_len > 0:
                style = ctypes.string_at(s.style_ptr, s.style_len).decode(
                    "utf-8", errors="replace"
                )
            return BorderSide(
                style=style,
                color_argb=int(s.color_argb) if s.has_color else None,
            )

        return Border(
            left=_side(cb.left),
            right=_side(cb.right),
            top=_side(cb.top),
            bottom=_side(cb.bottom),
            diagonal=_side(cb.diagonal),
        )

    def cell_alignment(self, style_idx: int) -> Alignment | None:
        """Resolve a cell's style index to its :class:`Alignment`
        (horizontal + wrap_text). Returns ``None`` on out-of-range
        indices. Cells without a nested ``<alignment>`` child surface
        as ``Alignment(horizontal="", wrap_text=False)`` — the OOXML
        default that the writer omits."""
        if not self._handle:
            raise ZlsxError("book is closed")
        if not _ffi._HAS_CELL_ALIGNMENT:
            raise RuntimeError(
                "loaded libzlsx does not expose cell_alignment "
                "(requires post-0.3.0 libzlsx); upgrade libzlsx"
            )
        # Bound-check before ctypes narrows to c_uint32. A Python int
        # outside [0, UINT32_MAX] would silently wrap and return
        # metadata for the wrong style; mirror the C ABI's "out of
        # range → None" contract.
        if style_idx < 0 or style_idx > 0xFFFFFFFF:
            return None
        ca = _ffi.CellAlignment()
        rc = _ffi.lib.zlsx_cell_alignment(self._handle, style_idx, ctypes.byref(ca))
        if rc != 0:
            return None
        horizontal = ""
        if ca.horizontal_len > 0:
            horizontal = ctypes.string_at(ca.horizontal_ptr, ca.horizontal_len).decode(
                "utf-8", errors="replace"
            )
        return Alignment(horizontal=horizontal, wrap_text=bool(ca.wrap_text))

    def is_date_format(self, style_idx: int) -> bool:
        """True when the style index resolves to a date / time /
        datetime pattern. Combine with ``xlsx.fromExcelSerial`` (or
        the Python equivalent) to auto-convert numeric cells to
        datetimes. Requires libzlsx 0.2.6+."""
        if not self._handle:
            raise ZlsxError("book is closed")
        if not _ffi._HAS_NUM_FMT:
            raise RuntimeError(
                "loaded libzlsx does not expose is_date_format "
                "(requires 0.2.6+); upgrade libzlsx"
            )
        return bool(_ffi.lib.zlsx_is_date_format(self._handle, style_idx))

    def close(self) -> None:
        """Drop our reference to the book. Active row iterators hold their
        own references, so this is safe to call before iteration finishes —
        the C ABI's refcount keeps the state alive until the last handle
        closes."""
        if self._handle:
            _ffi.lib.zlsx_book_close(self._handle)
            self._handle = None

    def __enter__(self) -> "Book":
        return self

    def __exit__(self, *exc_info) -> None:
        self.close()

    def __del__(self) -> None:
        try:
            self.close()
        except Exception:
            # __del__ must not raise.
            pass


# ─── Sheet ────────────────────────────────────────────────────────────


class Sheet:
    """A single worksheet within a :class:`Book`. Construct via
    :meth:`Book.sheet`."""

    def __init__(self, book: Book, index: int):
        self._book = book
        self.index = index
        self.name = book.sheets[index]

    def rows(self) -> "Rows":
        """Return a row iterator. Each iteration yields a ``list`` whose
        elements are Python values (see module docstring for the type
        mapping)."""
        if not self._book._handle:
            raise ZlsxError("Book is closed")
        return Rows(self._book, self.index)

    def read_all(self, header: bool = False) -> "tuple[list | None, list[list]]":
        """Materialise every row in this sheet into a ``list[list]``.

        Returns ``(header_row, data_rows)``. When ``header=True`` the
        first row is split out as the header; otherwise ``header_row``
        is ``None`` and ``data_rows`` contains every row.

        Convenience wrapper for callers who want to feed the result
        into ``pandas.DataFrame`` or ``polars.DataFrame``:

        .. code-block:: python

            with zlsx.open("data.xlsx") as book:
                headers, rows = book.sheet(0).read_all(header=True)
            df = pandas.DataFrame(rows, columns=headers)

        No optional dependency on pandas/polars — the return shape is
        plain Python lists, so any tabular library can consume it.

        Uses the bulk-FFI ``zlsx_matrix_open`` path when libzlsx 0.2.8+
        is loaded — one FFI call drains every row into a packed buffer,
        avoiding per-row dispatch overhead. Falls back to the per-row
        iterator on older libraries; result is identical."""
        if not self._book._handle:
            raise ZlsxError("Book is closed")
        if not _ffi._HAS_MATRIX:
            with self.rows() as r:
                all_rows = list(r)
            if not header or len(all_rows) == 0:
                return (None, all_rows)
            return (all_rows[0], all_rows[1:])

        err = ctypes.create_string_buffer(_ERR_BUF_LEN)
        handle = _ffi.lib.zlsx_matrix_open(
            self._book._handle, self.index, err, _ERR_BUF_LEN
        )
        if not handle:
            raise ZlsxError(f"zlsx_matrix_open: {_decode_err(err)}")
        try:
            cells_ptr = _ffi.cell_ptr()
            offsets_ptr = ctypes.POINTER(ctypes.c_size_t)()
            n_rows = ctypes.c_size_t()
            _ffi.lib.zlsx_matrix_data(
                handle,
                ctypes.byref(cells_ptr),
                ctypes.byref(offsets_ptr),
                ctypes.byref(n_rows),
            )
            n = n_rows.value
            all_rows = [None] * n
            # Inline cell decoding into the hot loop — saves the
            # Python function-call overhead × cell-count. On ECDC
            # (~441k cells) that's measurable.
            CE = _ffi.CELL_EMPTY
            CS = _ffi.CELL_STRING
            CI = _ffi.CELL_INTEGER
            CN = _ffi.CELL_NUMBER
            CB = _ffi.CELL_BOOLEAN
            string_at = ctypes.string_at
            for r_idx in range(n):
                start = offsets_ptr[r_idx]
                end = offsets_ptr[r_idx + 1]
                row = [None] * (end - start)
                ridx = 0
                for c in range(start, end):
                    cell = cells_ptr[c]
                    tag = cell.tag
                    if tag == CE:
                        v = None
                    elif tag == CS:
                        slen = cell.str_len
                        v = "" if slen == 0 else string_at(cell.str_ptr, slen).decode("utf-8", "replace")
                    elif tag == CI:
                        v = cell.i
                    elif tag == CN:
                        v = cell.f
                    elif tag == CB:
                        v = bool(cell.b)
                    else:
                        v = None
                    row[ridx] = v
                    ridx += 1
                all_rows[r_idx] = row
        finally:
            _ffi.lib.zlsx_matrix_close(handle)

        if not header or n == 0:
            return (None, all_rows)
        return (all_rows[0], all_rows[1:])


# ─── Rows ─────────────────────────────────────────────────────────────


class Rows:
    """Iterator over a sheet's rows. Normally constructed via
    :meth:`Sheet.rows`.

    The returned row lists are built fresh on each iteration — the
    underlying string slices point into library-owned buffers that are
    only valid until the next call, and we decode them to Python ``str``
    immediately to avoid dangling references.
    """

    def __init__(self, book: Book, sheet_idx: int):
        self._err = ctypes.create_string_buffer(_ERR_BUF_LEN)
        # Hold a reference to the Python Book so callers using iter29
        # helpers (`style_indices`, `number_format`) don't have to
        # thread it themselves, and to extend GC lifetime of the
        # underlying Book handle until this iterator is closed.
        self._book = book
        self._current_len = 0
        self._handle = _ffi.lib.zlsx_rows_open(
            book._handle, sheet_idx, self._err, _ERR_BUF_LEN
        )
        if not self._handle:
            raise ZlsxError(f"zlsx_rows_open: {_decode_err(self._err)}")

    def __iter__(self) -> Iterator[list]:
        return self

    def __next__(self) -> list:
        if not self._handle:
            raise ZlsxError("Rows iterator is closed")
        cells_ptr = _ffi.cell_ptr()
        cells_len = ctypes.c_size_t()
        rc = _ffi.lib.zlsx_rows_next(
            self._handle,
            ctypes.byref(cells_ptr),
            ctypes.byref(cells_len),
            self._err,
            _ERR_BUF_LEN,
        )
        if rc == 0:
            raise StopIteration
        if rc < 0:
            raise ZlsxError(f"zlsx_rows_next: {_decode_err(self._err)}")

        self._current_len = cells_len.value
        row = [_cell_to_py(cells_ptr[i]) for i in range(cells_len.value)]
        return row

    def style_indices(self) -> list[int | None]:
        """Style index for each cell in the most recently yielded row.
        ``None`` when the source `<c>` had no ``s`` attribute (General
        format). Layout mirrors the last row returned by ``next()`` so
        positional indexing matches. Raises :class:`RuntimeError` if
        the loaded libzlsx predates the 0.2.6+ numFmt ABI."""
        if not self._handle:
            raise ZlsxError("Rows iterator is closed")
        if not _ffi._HAS_NUM_FMT:
            raise RuntimeError(
                "loaded libzlsx does not expose per-cell style indices "
                "(requires 0.2.6+); upgrade libzlsx"
            )
        out: list[int | None] = []
        sidx = ctypes.c_uint32(0)
        for col in range(self._current_len):
            rc = _ffi.lib.zlsx_rows_style_at(self._handle, col, ctypes.byref(sidx))
            if rc == 0:
                out.append(int(sidx.value))
            elif rc == 1:
                out.append(None)
            else:
                # Out of range shouldn't happen within _current_len.
                out.append(None)
        return out

    def skip(self, n: int) -> int:
        """Advance past ``n`` rows without decoding their cells; returns
        how many were actually skipped (fewer than ``n`` only at end of
        sheet).

        Equivalent to calling ``next()`` ``n`` times and throwing the
        results away — same landing row, same numbering — but it does
        not build the Python row lists for what it passes. That makes
        range-partitioned reads affordable: partition *i* of a sheet
        must first get past *i·size* rows it will never look at, so
        decoding them turns K partitions into O(K²) work.

        Falls back to draining ``next()`` against a pre-0.8.0 libzlsx,
        so callers need no version check of their own.
        """
        if n < 0:
            raise ValueError(f"skip count must be non-negative, got {n}")
        if not self._handle:
            raise ZlsxError("Rows iterator is closed")
        if n == 0:
            return 0

        if not _ffi._HAS_ROWS_SKIP:
            skipped = 0
            for _ in range(n):
                try:
                    next(self)
                except StopIteration:
                    break
                skipped += 1
            return skipped

        out = ctypes.c_size_t(0)
        rc = _ffi.lib.zlsx_rows_skip(
            self._handle, n, ctypes.byref(out), self._err, _ERR_BUF_LEN
        )
        if rc != 0:
            raise ZlsxError(f"zlsx_rows_skip: {_decode_err(self._err)}")
        # The previous row's cells are gone; keep the side-channel
        # accessors (style_indices / parse_date) from reading a stale
        # length into a cleared buffer.
        self._current_len = 0
        return int(out.value)

    def parse_date(self, col_idx: int) -> "datetime.datetime | None":
        """Decode the current-row cell at ``col_idx`` as a date-styled
        number. Returns a Python ``datetime.datetime`` when the cell
        is a number/integer AND its style resolves to a date format
        AND the serial is in the valid Excel range (>= 61). Returns
        ``None`` otherwise (including out-of-range col_idx, string
        cells, and plain numbers without a date style).

        Rows only surface the current row — call after ``next()``.
        Requires libzlsx 0.2.6+."""
        import datetime as _dt
        if not self._handle:
            raise ZlsxError("Rows iterator is closed")
        if not _ffi._HAS_PARSE_DATE:
            raise RuntimeError(
                "loaded libzlsx does not expose rows_parse_date "
                "(requires 0.2.6+); upgrade libzlsx"
            )
        dt = _ffi.CDateTime()
        rc = _ffi.lib.zlsx_rows_parse_date(
            self._handle, col_idx, ctypes.byref(dt)
        )
        if rc != 0:
            return None
        return _dt.datetime(
            year=int(dt.year),
            month=int(dt.month),
            day=int(dt.day),
            hour=int(dt.hour),
            minute=int(dt.minute),
            second=int(dt.second),
        )

    def close(self) -> None:
        if self._handle:
            _ffi.lib.zlsx_rows_close(self._handle)
            self._handle = None

    def __enter__(self) -> "Rows":
        return self

    def __exit__(self, *exc_info) -> None:
        self.close()

    def __del__(self) -> None:
        try:
            self.close()
        except Exception:
            pass


# ─── Public entry point ───────────────────────────────────────────────


def to_excel_serial(dt) -> float:
    """Convert a Python ``datetime.datetime`` / ``datetime.date`` to
    an Excel serial-date number suitable for passing as a numeric
    cell. Combine with ``Style(number_format="yyyy-mm-dd")`` to
    write a date cell round-trippable via ``Rows.parse_date``.

    Raises ``ValueError`` when the date is outside the
    round-trippable range (year < 1900 or > 9999, or ≤ 1900-02-29 —
    the 1900 leap-year bug exclusion). Requires libzlsx 0.2.6+.
    """
    import datetime as _dt
    if not _ffi._HAS_TO_EXCEL_SERIAL:
        raise RuntimeError(
            "loaded libzlsx does not expose datetime_to_serial "
            "(requires 0.2.6+); upgrade libzlsx"
        )
    # datetime.date (not datetime.datetime) has no hour/minute/second
    # — treat it as midnight.
    if isinstance(dt, _dt.datetime):
        h, m, s = dt.hour, dt.minute, dt.second
    elif isinstance(dt, _dt.date):
        h, m, s = 0, 0, 0
    else:
        raise TypeError(
            f"expected datetime.date or datetime.datetime, got {type(dt).__name__}"
        )
    cdt = _ffi.CDateTime(
        year=dt.year, month=dt.month, day=dt.day,
        hour=h, minute=m, second=s, _pad=0,
    )
    out = ctypes.c_double(0.0)
    rc = _ffi.lib.zlsx_datetime_to_serial(ctypes.byref(cdt), ctypes.byref(out))
    if rc != 0:
        raise ValueError(
            f"{dt!r} is outside Excel's round-trippable date range "
            "(year 1900..9999, > 1900-02-29)"
        )
    return float(out.value)


def open(path: Union[str, Path]) -> Book:  # noqa: A001  (shadows builtin by design)
    """Open an ``.xlsx`` file for reading.

    Returns a :class:`Book` handle. The file must exist and be a valid
    xlsx archive. Raises :class:`ZlsxError` on parse failure.
    """
    return Book(path)


def open_bytes(data: Union[bytes, bytearray, memoryview]) -> Book:
    """Open an ``.xlsx`` workbook from bytes already in memory.

    No filesystem access, no temp file: the buffer is parsed eagerly and
    borrowed only for the duration of this call — the caller may discard
    ``data`` immediately after ``open_bytes`` returns. This is the entry
    point for callers that receive workbook bytes without a path (SQL
    UDFs over binary columns, network payloads, object-store reads)::

        with zlsx.open_bytes(content) as book:
            for row in book.sheet(0).rows():
                ...

    Raises :class:`ZlsxError` on parse failure.
    """
    raw = bytes(data)
    book = Book.__new__(Book)
    book._handle = None
    book._err = ctypes.create_string_buffer(_ERR_BUF_LEN)
    handle = _ffi.lib.zlsx_book_open_buffer(raw, len(raw), book._err, _ERR_BUF_LEN)
    if not handle:
        raise ZlsxError(f"zlsx_book_open_buffer({len(raw)} bytes): {_decode_err(book._err)}")
    book._attach(handle)
    return book


def read(
    path: Union[str, Path],
    sheet: Union[int, str] = 0,
    header: bool = False,
) -> "tuple[list | None, list[list]]":
    """Open ``path`` and materialise one sheet's rows in a single
    call. Closes the book before returning.

    ``sheet`` can be a 0-based index or a sheet name. ``header=True``
    splits the first row as the header; otherwise the entire sheet
    lands in the second element of the tuple.

    Wraps :meth:`Book.sheet` + :meth:`Sheet.read_all` for the
    "just-give-me-the-rows" case — typical entry point for callers
    that feed into pandas / polars:

    .. code-block:: python

        headers, rows = zlsx.read("data.xlsx", header=True)
        df = pandas.DataFrame(rows, columns=headers)
    """
    with open(path) as book:
        if isinstance(sheet, str):
            for i, name in enumerate(book.sheets):
                if name == sheet:
                    idx = i
                    break
            else:
                raise ZlsxError(f"sheet {sheet!r} not found; have {book.sheets!r}")
        else:
            idx = int(sheet)
            if idx < 0 or idx >= len(book.sheets):
                raise ZlsxError(
                    f"sheet index {idx} out of range (book has {len(book.sheets)} sheets)"
                )
        return book.sheet(idx).read_all(header=header)


# ─── Writer ───────────────────────────────────────────────────────────


def _py_value_to_cell(value):
    """Convert a Python value to a (ctypes Cell, optional keep-alive)
    tuple. For string cells, the keep-alive is the ctypes buffer holding
    the UTF-8 bytes — caller must hold it until the write call returns,
    otherwise cell.str_ptr becomes a dangling pointer."""
    cell = _ffi.Cell()
    if value is None:
        cell.tag = _ffi.CELL_EMPTY
        return cell, None
    if isinstance(value, bool):
        # Check bool BEFORE int — `isinstance(True, int)` is True in
        # Python, but we want True/False to emit as booleans.
        cell.tag = _ffi.CELL_BOOLEAN
        cell.b = 1 if value else 0
        return cell, None
    if isinstance(value, int):
        cell.tag = _ffi.CELL_INTEGER
        cell.i = value
        return cell, None
    if isinstance(value, float):
        cell.tag = _ffi.CELL_NUMBER
        cell.f = value
        return cell, None
    if isinstance(value, str):
        raw = value.encode("utf-8")
        cell.tag = _ffi.CELL_STRING
        cell.str_len = len(raw)
        # Create a ctypes array from the bytes and point str_ptr at it.
        # The bytes object + buffer must outlive the write call — we
        # return both so the caller holds the reference.
        buf = (ctypes.c_ubyte * len(raw)).from_buffer_copy(raw)
        cell.str_ptr = ctypes.cast(buf, ctypes.POINTER(ctypes.c_ubyte))
        return cell, buf
    raise TypeError(
        f"unsupported cell type: {type(value).__name__} (expected None, bool, int, float, str)"
    )


class SheetWriter:
    """A handle for writing rows to one sheet of a :class:`Writer`.

    Obtained via :meth:`Writer.add_sheet`. The underlying C handle is
    borrowed from the parent Writer and becomes invalid when the Writer
    is closed — do not hold on to a SheetWriter after its parent exits.
    """

    def __init__(self, parent: "Writer", handle):
        self._parent = parent
        self._handle = handle
        self._err = ctypes.create_string_buffer(_ERR_BUF_LEN)

    def _require_handle(self) -> None:
        """Raise a clear error if this SheetWriter was invalidated by
        ``Writer.close()``. Called at the top of every method that
        would otherwise pass a NULL pointer to the C ABI (whose
        signature is non-optional ``*SheetWriter`` and would null-deref
        on field access)."""
        if self._handle is None:
            raise RuntimeError(
                "SheetWriter used after its parent Writer was closed"
            )

    def write_row(self, values, styles=None) -> None:
        """Append a row. ``values`` is any iterable of ``None | bool |
        int | float | str``. Integers outside ±2^53-significant-bits
        raise :class:`ZlsxError` (Excel stores numerics as IEEE-754
        doubles — oversized ints would silently round on open).

        ``styles``, if provided, must be an iterable of the same length
        as ``values`` where each element is a style index returned by
        :meth:`Writer.add_style` (or 0 for the default no-style). If
        ``styles`` is None, every cell inherits the default formatting.
        """
        self._require_handle()
        cells_list = list(values)
        n = len(cells_list)

        if styles is not None:
            styles_list = list(styles)
            if len(styles_list) != n:
                raise ValueError(
                    f"styles length {len(styles_list)} must match values length {n}"
                )
        else:
            styles_list = None

        if n == 0:
            # Emit an empty row via the ABI's explicit null/zero path.
            rc = _ffi.lib.zlsx_sheet_writer_write_row(
                self._handle, None, 0, self._err, _ERR_BUF_LEN
            )
            if rc != 0:
                raise ZlsxError(
                    f"zlsx_sheet_writer_write_row (empty): {_decode_err(self._err)}"
                )
            return

        cell_array = (_ffi.Cell * n)()
        # Hold str buffers alive for the duration of the C call — the
        # cell's str_ptr points into these buffers and ctypes won't
        # keep them alive on its own.
        keepers = []
        for i, v in enumerate(cells_list):
            cell, keeper = _py_value_to_cell(v)
            cell_array[i] = cell
            if keeper is not None:
                keepers.append(keeper)

        if styles_list is None:
            rc = _ffi.lib.zlsx_sheet_writer_write_row(
                self._handle,
                ctypes.cast(cell_array, _ffi.cell_ptr),
                n,
                self._err,
                _ERR_BUF_LEN,
            )
            if rc != 0:
                raise ZlsxError(
                    f"zlsx_sheet_writer_write_row: {_decode_err(self._err)}"
                )
        else:
            if not _ffi._HAS_STYLES:
                raise RuntimeError(
                    "loaded libzlsx does not expose zlsx_sheet_writer_write_row_styled "
                    "(requires 0.2.4+); upgrade libzlsx or unset the styles argument"
                )
            style_array = (ctypes.c_uint32 * n)(*[int(s) for s in styles_list])
            rc = _ffi.lib.zlsx_sheet_writer_write_row_styled(
                self._handle,
                ctypes.cast(cell_array, _ffi.cell_ptr),
                style_array,
                n,
                self._err,
                _ERR_BUF_LEN,
            )
            if rc != 0:
                raise ZlsxError(
                    f"zlsx_sheet_writer_write_row_styled: {_decode_err(self._err)}"
                )

        # Reference `keepers` past the call so ctypes doesn't free the
        # backing str buffers while the C side is still reading them.
        del keepers

    def write_row_with_formulas(self, values, formulas, dialect=None) -> None:
        """Append a row mixing plain value cells with formula cells.

        ``values`` is the same iterable accepted by :meth:`write_row` —
        for formula cells it carries the cached ``<v>`` value Excel
        shows until the sheet is recalculated (pass ``None`` for "no
        cached value").

        ``formulas`` must be an iterable of the same length: each
        element is ``None`` (regular value cell), a ``str`` with the
        formula text (e.g. ``"A1+B1"``, no leading ``=``), or a
        :class:`FormulaSpec` (M9a2) carrying a dialect —
        ``FormulaSpec.cse(text, ref)`` declares a legacy CSE rectangle
        anchored at that cell; its members (the range's other cells,
        written by later calls) carry cached values and no formula,
        and the save refuses while any rectangle is incomplete.

        ``dialect=`` is row-wide sugar applying one dialect to every
        plain-``str`` formula in the row (``"cse"`` is per-cell only —
        it needs a ref). Plain-``str`` rows keep the legacy ABI path,
        so they still work against an older dylib; FormulaSpec or
        ``dialect=`` requires libzlsx 0.8.0+.

        Requires libzlsx 0.2.7+. Raises :class:`RuntimeError` against
        an older dylib that doesn't ship the symbol.
        """
        self._require_handle()
        if not _ffi._HAS_WRITE_ROW_WITH_FORMULAS:
            raise RuntimeError(
                "loaded libzlsx does not expose write_row_with_formulas "
                "(requires 0.2.7+); upgrade libzlsx"
            )

        values_list = list(values)
        formulas_list = list(formulas)
        n = len(values_list)
        if len(formulas_list) != n:
            raise ValueError(
                f"formulas length {len(formulas_list)} must match values length {n}"
            )

        needs_v2 = dialect is not None or any(
            isinstance(f, FormulaSpec) for f in formulas_list
        )
        if needs_v2:
            return self._write_row_with_formulas_v2(values_list, formulas_list, dialect)

        if n == 0:
            rc = _ffi.lib.zlsx_sheet_writer_write_row_with_formulas(
                self._handle, None, None, None, 0, self._err, _ERR_BUF_LEN
            )
            if rc != 0:
                raise ZlsxError(
                    f"zlsx_sheet_writer_write_row_with_formulas (empty): {_decode_err(self._err)}"
                )
            return

        cell_array = (_ffi.Cell * n)()
        formula_ptr_array = (ctypes.POINTER(ctypes.c_ubyte) * n)()
        formula_len_array = (ctypes.c_size_t * n)()
        # Hold all str buffers (cells + formulas) alive across the C call.
        keepers = []
        for i, v in enumerate(values_list):
            cell, keeper = _py_value_to_cell(v)
            cell_array[i] = cell
            if keeper is not None:
                keepers.append(keeper)

            f = formulas_list[i]
            if f is None:
                formula_ptr_array[i] = ctypes.cast(None, ctypes.POINTER(ctypes.c_ubyte))
                formula_len_array[i] = 0
            else:
                if not isinstance(f, str):
                    raise TypeError(
                        f"formulas[{i}] must be None or str, got {type(f).__name__}"
                    )
                if f == "":
                    # Empty string would round-trip to formula_len == 0,
                    # which the C ABI reads as "no formula on this column"
                    # — silently dropping the caller's intent. Reject up
                    # front instead. Pass None for "regular value cell".
                    raise ValueError(
                        f"formulas[{i}] is the empty string; pass None for a "
                        "regular value cell or a non-empty formula like 'A1+B1'"
                    )
                fbytes = f.encode("utf-8")
                # ctypes.c_char_p keeps `fbytes` referenced; cast to ubyte*
                # for the pointer array's element type.
                buf = ctypes.create_string_buffer(fbytes, len(fbytes))
                keepers.append(buf)
                formula_ptr_array[i] = ctypes.cast(buf, ctypes.POINTER(ctypes.c_ubyte))
                formula_len_array[i] = len(fbytes)

        rc = _ffi.lib.zlsx_sheet_writer_write_row_with_formulas(
            self._handle,
            ctypes.cast(cell_array, _ffi.cell_ptr),
            formula_ptr_array,
            formula_len_array,
            n,
            self._err,
            _ERR_BUF_LEN,
        )
        if rc != 0:
            raise ZlsxError(
                f"zlsx_sheet_writer_write_row_with_formulas: {_decode_err(self._err)}"
            )

        # Keep the buffers alive until after the C call returns.
        del keepers

    def _write_row_with_formulas_v2(self, values_list, formulas_list, dialect) -> None:
        """The M9a2 descriptor path: per-cell zlsx_formula_cell_v1
        marshalling for FormulaSpec rows and the row-wide ``dialect=``
        sugar."""
        if not _ffi._HAS_FORMULAS_V2:
            raise RuntimeError(
                "loaded libzlsx does not expose write_row_with_formulas_v2 "
                "(requires 0.8.0+); upgrade libzlsx"
            )
        if dialect is not None:
            if dialect not in ("scalar", "dynamic_array"):
                raise ValueError(
                    "row-wide dialect must be 'scalar' or 'dynamic_array'; "
                    "'cse' needs a per-cell ref — use FormulaSpec.cse(text, ref)"
                )

        n = len(values_list)
        if n == 0:
            rc = _ffi.lib.zlsx_sheet_writer_write_row_with_formulas_v2(
                self._handle, None, None, 0, self._err, _ERR_BUF_LEN
            )
            if rc != _ffi.ZLSX_OK:
                raise ZlsxError(
                    f"zlsx_sheet_writer_write_row_with_formulas_v2 (empty): "
                    f"{_decode_err(self._err)}"
                )
            return

        cell_array = (_ffi.Cell * n)()
        desc_array = (_ffi.FormulaCellV1 * n)()
        keepers = []
        for i, v in enumerate(values_list):
            cell, keeper = _py_value_to_cell(v)
            cell_array[i] = cell
            if keeper is not None:
                keepers.append(keeper)

            f = formulas_list[i]
            if f is None:
                continue  # ctypes zero-init: text NULL = plain slot
            if isinstance(f, str):
                if f == "":
                    raise ValueError(
                        f"formulas[{i}] is the empty string; pass None for a "
                        "regular value cell or a non-empty formula like 'A1+B1'"
                    )
                f = FormulaSpec(f, dialect or "scalar")
            elif not isinstance(f, FormulaSpec):
                raise TypeError(
                    f"formulas[{i}] must be None, str or FormulaSpec, "
                    f"got {type(f).__name__}"
                )
            tbytes = f.text.encode("utf-8")
            tbuf = ctypes.create_string_buffer(tbytes, len(tbytes))
            keepers.append(tbuf)
            desc_array[i].text = ctypes.cast(tbuf, ctypes.POINTER(ctypes.c_ubyte))
            desc_array[i].text_len = len(tbytes)
            desc_array[i].dialect = _FORMULA_DIALECT_CODES[f.dialect]
            if f.ref is not None:
                rbytes = f.ref.encode("utf-8")
                rbuf = ctypes.create_string_buffer(rbytes, len(rbytes))
                keepers.append(rbuf)
                desc_array[i].ref = ctypes.cast(rbuf, ctypes.POINTER(ctypes.c_ubyte))
                desc_array[i].ref_len = len(rbytes)

        rc = _ffi.lib.zlsx_sheet_writer_write_row_with_formulas_v2(
            self._handle,
            ctypes.cast(cell_array, _ffi.cell_ptr),
            desc_array,
            n,
            self._err,
            _ERR_BUF_LEN,
        )
        if rc != _ffi.ZLSX_OK:
            raise ZlsxError(
                f"zlsx_sheet_writer_write_row_with_formulas_v2: "
                f"{_decode_err(self._err)}"
            )
        del keepers

    def write_rich_row(self, values) -> None:
        """Append a row mixing plain cells with rich-text cells.

        Each element of ``values`` is either a plain Python value
        (``None``, ``bool``, ``int``, ``float``, ``str``) or an
        iterable of :class:`RichRun` for a rich-text cell. Rich
        cells get emitted as a single ``<si>`` containing one
        ``<r><rPr/>…<t/></r>`` per run; plain cells follow the same
        semantics as :meth:`write_row`.

        Requires libzlsx 0.2.6+."""
        self._require_handle()
        if not _ffi._HAS_WRITE_RICH_ROW:
            raise RuntimeError(
                "loaded libzlsx does not expose write_rich_row "
                "(requires 0.2.6+); upgrade libzlsx"
            )
        cells_list = list(values)
        n = len(cells_list)
        if n == 0:
            rc = _ffi.lib.zlsx_sheet_writer_write_row(
                self._handle, None, 0, self._err, _ERR_BUF_LEN
            )
            if rc != 0:
                raise ZlsxError(
                    f"zlsx_sheet_writer_write_row (empty): {_decode_err(self._err)}"
                )
            return

        cell_array = (_ffi.Cell * n)()
        lens_array = (ctypes.c_size_t * n)()
        ptrs_array = (ctypes.POINTER(_ffi.CRichRun) * n)()
        keepers: list = []

        for i, v in enumerate(cells_list):
            # A rich cell is any iterable of RichRun that isn't a str.
            if isinstance(v, (list, tuple)) and all(isinstance(r, RichRun) for r in v):
                runs_list = list(v)
                m = len(runs_list)
                if m == 0:
                    raise ValueError(
                        f"rich cell at column {i} has zero runs — pass a non-empty "
                        "list[RichRun] or use a plain value"
                    )
                run_array = (_ffi.CRichRun * m)()
                for j, r in enumerate(runs_list):
                    text_bytes = r.text.encode("utf-8")
                    text_buf = (ctypes.c_ubyte * max(len(text_bytes), 1)).from_buffer_copy(
                        text_bytes or b"\x00"
                    )
                    font_bytes = r.font_name.encode("utf-8") if r.font_name else b""
                    font_buf = (ctypes.c_ubyte * max(len(font_bytes), 1)).from_buffer_copy(
                        font_bytes or b"\x00"
                    )
                    run_array[j] = _ffi.CRichRun(
                        text_ptr=ctypes.cast(text_buf, ctypes.POINTER(ctypes.c_ubyte)),
                        text_len=len(text_bytes),
                        bold=1 if r.bold else 0,
                        italic=1 if r.italic else 0,
                        has_color=1 if r.color_argb is not None else 0,
                        has_size=1 if r.size is not None else 0,
                        color_argb=r.color_argb or 0,
                        size=r.size if r.size is not None else 0.0,
                        font_name_ptr=ctypes.cast(font_buf, ctypes.POINTER(ctypes.c_ubyte)),
                        font_name_len=len(font_bytes),
                    )
                    keepers.extend([text_buf, font_buf])
                # Placeholder plain cell — the C side ignores it when
                # rich_runs_lens[i] > 0.
                cell, keeper = _py_value_to_cell(None)
                cell_array[i] = cell
                if keeper is not None:
                    keepers.append(keeper)
                lens_array[i] = m
                ptrs_array[i] = ctypes.cast(run_array, ctypes.POINTER(_ffi.CRichRun))
                keepers.append(run_array)
            else:
                cell, keeper = _py_value_to_cell(v)
                cell_array[i] = cell
                if keeper is not None:
                    keepers.append(keeper)
                lens_array[i] = 0
                ptrs_array[i] = ctypes.POINTER(_ffi.CRichRun)()

        rc = _ffi.lib.zlsx_sheet_writer_write_rich_row(
            self._handle,
            ctypes.cast(cell_array, _ffi.cell_ptr),
            ptrs_array,
            lens_array,
            n,
            self._err,
            _ERR_BUF_LEN,
        )
        if rc != 0:
            raise ZlsxError(
                f"zlsx_sheet_writer_write_rich_row: {_decode_err(self._err)}"
            )
        del keepers


# Attach the stage-5 per-sheet methods to SheetWriter.
# ------------------------------------------------------
# Implemented as module-level function attachments so the class body
# above stays focused on the stage 1-4 row-writing API.

def _sheet_set_column_width(self: "SheetWriter", col_idx: int, width: float) -> None:
    """Set the display width of column ``col_idx`` (0-based) in
    character units (Excel default 8.43). Validated upfront."""
    self._require_handle()
    if not _ffi._HAS_SHEET_FEATURES:
        raise RuntimeError(
            "loaded libzlsx does not expose sheet layout features "
            "(requires 0.2.4+); upgrade libzlsx"
        )
    # Bound-check signed Python ints before ctypes wraps them into
    # u32 — a bare `ctypes.c_uint32(-1)` becomes UINT32_MAX, which
    # then overflows `col_idx + 1` inside the Zig writer.
    if col_idx < 0:
        raise ValueError(f"col_idx must be >= 0, got {col_idx}")
    rc = _ffi.lib.zlsx_sheet_writer_set_column_width(
        self._handle, int(col_idx), float(width), self._err, _ERR_BUF_LEN
    )
    if rc != 0:
        raise ZlsxError(
            f"zlsx_sheet_writer_set_column_width: {_decode_err(self._err)}"
        )


def _sheet_freeze_panes(self: "SheetWriter", rows: int = 0, cols: int = 0) -> None:
    """Freeze the top ``rows`` rows and left ``cols`` columns. Pass 0
    on an axis to leave it unfrozen. Overrides any previous freeze."""
    self._require_handle()
    if not _ffi._HAS_SHEET_FEATURES:
        raise RuntimeError(
            "loaded libzlsx does not expose freeze_panes (requires 0.2.4+); "
            "upgrade libzlsx"
        )
    if rows < 0 or cols < 0:
        raise ValueError(
            f"freeze_panes rows/cols must be >= 0, got rows={rows} cols={cols}"
        )
    _ffi.lib.zlsx_sheet_writer_freeze_panes(
        self._handle, int(rows), int(cols)
    )


def _sheet_set_row_height(self: "SheetWriter", row_idx: int, height: float) -> None:
    """Set the display height of row ``row_idx`` (0-based) in points.
    Excel accepts heights in (0, 409.5]. Raises :class:`ZlsxError`
    on ``InvalidRowHeight`` or ``RowOutOfRange``."""
    self._require_handle()
    if not _ffi._HAS_SET_ROW_HEIGHT:
        raise RuntimeError(
            "loaded libzlsx does not expose set_row_height "
            "(requires post-0.3.0 libzlsx); upgrade libzlsx"
        )
    if row_idx < 0:
        raise ValueError(f"row_idx must be >= 0, got {row_idx}")
    if row_idx > 0xFFFFFFFF:
        # ctypes c_uint32 narrowing would wrap mod 2^32 before the
        # C ABI can reject; bound-check up-front.
        raise ValueError(f"row_idx must be <= UINT32_MAX, got {row_idx}")
    rc = _ffi.lib.zlsx_sheet_writer_set_row_height(
        self._handle, int(row_idx), float(height), self._err, _ERR_BUF_LEN
    )
    if rc != 0:
        raise ZlsxError(
            f"zlsx_sheet_writer_set_row_height: {_decode_err(self._err)}"
        )


def _sheet_freeze_panes_checked(self: "SheetWriter", rows: int = 0, cols: int = 0) -> None:
    """Like :py:meth:`freeze_panes` but raises :class:`ZlsxError`
    on ``RowOutOfRange`` / ``ColumnOutOfRange`` instead of clamping
    silently. Prefer this over the legacy clamping form when callers
    want to catch out-of-range inputs."""
    self._require_handle()
    if not _ffi._HAS_FREEZE_PANES_CHECKED:
        raise RuntimeError(
            "loaded libzlsx does not expose freeze_panes_checked "
            "(requires post-0.3.0 libzlsx); upgrade libzlsx"
        )
    if rows < 0 or cols < 0:
        raise ValueError(
            f"freeze_panes_checked rows/cols must be >= 0, got rows={rows} cols={cols}"
        )
    if rows > 0xFFFFFFFF or cols > 0xFFFFFFFF:
        # ctypes c_uint32 narrowing wraps mod 2^32; the checked
        # variant's typed-error promise relies on bound-checking
        # before the FFI sees the value.
        raise ValueError(
            f"freeze_panes_checked rows/cols must be <= UINT32_MAX, got rows={rows} cols={cols}"
        )
    rc = _ffi.lib.zlsx_sheet_writer_freeze_panes_checked(
        self._handle, int(rows), int(cols), self._err, _ERR_BUF_LEN
    )
    if rc != 0:
        raise ZlsxError(
            f"zlsx_sheet_writer_freeze_panes_checked: {_decode_err(self._err)}"
        )


def _sheet_set_auto_filter(self: "SheetWriter", range_str: str) -> None:
    """Apply an auto-filter over ``range_str`` (A1-style, e.g. 'A1:E1')."""
    self._require_handle()
    if not _ffi._HAS_SHEET_FEATURES:
        raise RuntimeError(
            "loaded libzlsx does not expose set_auto_filter (requires 0.2.4+); "
            "upgrade libzlsx"
        )
    raw = range_str.encode("utf-8")
    buf = (ctypes.c_ubyte * max(len(raw), 1)).from_buffer_copy(raw or b"\x00")
    rc = _ffi.lib.zlsx_sheet_writer_set_auto_filter(
        self._handle,
        ctypes.cast(buf, ctypes.POINTER(ctypes.c_ubyte)),
        len(raw),
        self._err,
        _ERR_BUF_LEN,
    )
    # Keep buf alive through the call.
    del buf
    if rc != 0:
        raise ZlsxError(
            f"zlsx_sheet_writer_set_auto_filter: {_decode_err(self._err)}"
        )


def _sheet_add_merged_cell(self: "SheetWriter", range_str: str) -> None:
    """Register a rectangular merged cell range (A1-style, e.g. 'A1:B2').

    Single-cell ranges, inverted corners, lowercase, and references
    past Excel's 16 384 × 1 048 576 cap are rejected with
    :class:`ZlsxError`. Multiple merges per sheet are allowed but
    must not overlap — Excel rejects overlapping pairs at file-open
    time."""
    self._require_handle()
    if not _ffi._HAS_MERGED_CELL:
        raise RuntimeError(
            "loaded libzlsx does not expose add_merged_cell (requires 0.2.5+); "
            "upgrade libzlsx"
        )
    raw = range_str.encode("utf-8")
    buf = (ctypes.c_ubyte * max(len(raw), 1)).from_buffer_copy(raw or b"\x00")
    rc = _ffi.lib.zlsx_sheet_writer_add_merged_cell(
        self._handle,
        ctypes.cast(buf, ctypes.POINTER(ctypes.c_ubyte)),
        len(raw),
        self._err,
        _ERR_BUF_LEN,
    )
    # Keep buf alive through the call.
    del buf
    if rc != 0:
        raise ZlsxError(
            f"zlsx_sheet_writer_add_merged_cell: {_decode_err(self._err)}"
        )


def _sheet_add_hyperlink(self: "SheetWriter", range_str: str, url: str) -> None:
    """Attach an external-URL hyperlink to a cell or rectangular range.

    ``range_str`` is A1-style (``"A1"`` or ``"B2:C3"``); ``url`` is the
    external target (``http``/``https``/``mailto``/``file``/…). Both
    args are duped immediately; the URL is xml-escaped on emit, so
    query strings with ``&`` are safe. Raises :class:`ZlsxError` on
    malformed ranges (``InvalidHyperlinkRange``) or empty URLs
    (``InvalidHyperlinkUrl``)."""
    self._require_handle()
    if not _ffi._HAS_HYPERLINK:
        raise RuntimeError(
            "loaded libzlsx does not expose add_hyperlink (requires 0.2.5+); "
            "upgrade libzlsx"
        )
    range_raw = range_str.encode("utf-8")
    url_raw = url.encode("utf-8")
    range_buf = (ctypes.c_ubyte * max(len(range_raw), 1)).from_buffer_copy(
        range_raw or b"\x00"
    )
    url_buf = (ctypes.c_ubyte * max(len(url_raw), 1)).from_buffer_copy(
        url_raw or b"\x00"
    )
    rc = _ffi.lib.zlsx_sheet_writer_add_hyperlink(
        self._handle,
        ctypes.cast(range_buf, ctypes.POINTER(ctypes.c_ubyte)),
        len(range_raw),
        ctypes.cast(url_buf, ctypes.POINTER(ctypes.c_ubyte)),
        len(url_raw),
        self._err,
        _ERR_BUF_LEN,
    )
    # Keep the ctypes arrays alive through the call.
    del range_buf
    del url_buf
    if rc != 0:
        raise ZlsxError(
            f"zlsx_sheet_writer_add_hyperlink: {_decode_err(self._err)}"
        )


def _sheet_add_internal_hyperlink(self: "SheetWriter", range_str: str, location: str) -> None:
    """Attach an internal (same-workbook) hyperlink.

    ``range_str`` is A1-style (``"A1"`` / ``"B2:C3"``). ``location``
    is the target reference Excel writes verbatim into
    ``<hyperlink location="…"/>``, e.g. ``"Sheet2!A1"`` or
    ``"'Sheet With Spaces'!B2"``. Raises :class:`ZlsxError` on
    malformed ranges (``InvalidHyperlinkRange``) or empty location
    (``InvalidHyperlinkLocation``). Requires libzlsx 0.2.7+."""
    self._require_handle()
    if not _ffi._HAS_INTERNAL_HYPERLINK:
        raise RuntimeError(
            "loaded libzlsx does not expose add_internal_hyperlink "
            "(requires 0.2.7+); upgrade libzlsx"
        )
    range_raw = range_str.encode("utf-8")
    loc_raw = location.encode("utf-8")
    range_buf = (ctypes.c_ubyte * max(len(range_raw), 1)).from_buffer_copy(
        range_raw or b"\x00"
    )
    loc_buf = (ctypes.c_ubyte * max(len(loc_raw), 1)).from_buffer_copy(
        loc_raw or b"\x00"
    )
    rc = _ffi.lib.zlsx_sheet_writer_add_internal_hyperlink(
        self._handle,
        ctypes.cast(range_buf, ctypes.POINTER(ctypes.c_ubyte)),
        len(range_raw),
        ctypes.cast(loc_buf, ctypes.POINTER(ctypes.c_ubyte)),
        len(loc_raw),
        self._err,
        _ERR_BUF_LEN,
    )
    del range_buf
    del loc_buf
    if rc != 0:
        raise ZlsxError(
            f"zlsx_sheet_writer_add_internal_hyperlink: {_decode_err(self._err)}"
        )


def _sheet_add_comment(
    self: "SheetWriter",
    ref: str,
    author: str,
    text: str,
) -> None:
    """Attach a cell comment (note) to ``ref``.

    ``ref`` is a single-cell A1 reference (``"B2"``); ranges raise
    :class:`ZlsxError` (``InvalidCommentRef``). ``author`` shows in
    Excel's comment thread header — pass empty for anonymous. ``text``
    is the plain-text body; XML-special chars are escaped on emit.

    Requires libzlsx 0.2.6+."""
    self._require_handle()
    if not _ffi._HAS_COMMENT_WRITER:
        raise RuntimeError(
            "loaded libzlsx does not expose add_comment "
            "(requires 0.2.6+); upgrade libzlsx"
        )
    ref_raw = ref.encode("utf-8")
    author_raw = author.encode("utf-8")
    text_raw = text.encode("utf-8")
    ref_buf = (ctypes.c_ubyte * max(len(ref_raw), 1)).from_buffer_copy(
        ref_raw or b"\x00"
    )
    author_buf = (ctypes.c_ubyte * max(len(author_raw), 1)).from_buffer_copy(
        author_raw or b"\x00"
    )
    text_buf = (ctypes.c_ubyte * max(len(text_raw), 1)).from_buffer_copy(
        text_raw or b"\x00"
    )
    ptr_t = ctypes.POINTER(ctypes.c_ubyte)
    rc = _ffi.lib.zlsx_sheet_writer_add_comment(
        self._handle,
        ctypes.cast(ref_buf, ptr_t),
        len(ref_raw),
        ctypes.cast(author_buf, ptr_t),
        len(author_raw),
        ctypes.cast(text_buf, ptr_t),
        len(text_raw),
        self._err,
        _ERR_BUF_LEN,
    )
    del ref_buf, author_buf, text_buf
    if rc != 0:
        raise ZlsxError(
            f"zlsx_sheet_writer_add_comment: {_decode_err(self._err)}"
        )


def _sheet_add_data_validation_list(
    self: "SheetWriter",
    range_str: str,
    values: list,
) -> None:
    """Attach a list-type data validation (dropdown) to a cell or
    range. ``range_str`` is A1-style (``"A1"`` or ``"B2:B10"``);
    ``values`` is a non-empty iterable of strings that become the
    dropdown options. Embedded commas and bare double-quotes in
    values are rejected (Excel's list format can't represent them);
    XML-special chars like ``&``, ``<``, ``>`` are escaped on emit.
    Raises :class:`ZlsxError` on ``InvalidHyperlinkRange`` or
    ``InvalidDataValidation``."""
    self._require_handle()
    if not _ffi._HAS_DATA_VALIDATION:
        raise RuntimeError(
            "loaded libzlsx does not expose add_data_validation_list "
            "(requires 0.2.5+); upgrade libzlsx"
        )
    # Materialise to list to allow iteration multiple times.
    vals = list(values)
    if len(vals) > 256:
        raise ValueError(f"data validation list supports up to 256 values, got {len(vals)}")

    range_raw = range_str.encode("utf-8")
    range_buf = (ctypes.c_ubyte * max(len(range_raw), 1)).from_buffer_copy(
        range_raw or b"\x00"
    )

    # Build parallel arrays: a `POINTER(c_ubyte)` per value + matching length.
    # Keep the underlying `c_ubyte` arrays alive via a Python-side list so
    # the C side sees valid memory for the whole call.
    raw_values: list[bytes] = [v.encode("utf-8") for v in vals]
    value_bufs = [
        (ctypes.c_ubyte * max(len(raw), 1)).from_buffer_copy(raw or b"\x00")
        for raw in raw_values
    ]
    ptr_t = ctypes.POINTER(ctypes.c_ubyte)
    ptr_array = (ptr_t * len(vals))()
    len_array = (ctypes.c_size_t * len(vals))()
    for i, (b, raw) in enumerate(zip(value_bufs, raw_values)):
        ptr_array[i] = ctypes.cast(b, ptr_t)
        len_array[i] = len(raw)

    rc = _ffi.lib.zlsx_sheet_writer_add_data_validation_list(
        self._handle,
        ctypes.cast(range_buf, ptr_t),
        len(range_raw),
        ptr_array,
        len_array,
        len(vals),
        self._err,
        _ERR_BUF_LEN,
    )
    # Keep buffers alive through the call.
    del range_buf, value_bufs, ptr_array, len_array
    if rc != 0:
        raise ZlsxError(
            f"zlsx_sheet_writer_add_data_validation_list: {_decode_err(self._err)}"
        )


# Writer-side kind / op code tables mirror `ZLSX_DV_KIND_*` / `OP_*`.
# Kept separate from the reader-side `_DV_KIND_FROM_CODE` dict to make
# intent explicit (writer rejects list / unknown / custom codes here —
# list has its own entry point, custom uses `_sheet_add_data_validation_custom`,
# unknown is a forward-compat marker not a user-writeable kind).
_DV_WRITER_KIND_CODES = {
    "whole": 1,
    "decimal": 2,
    "date": 3,
    "time": 4,
    "text_length": 5,
}
_DV_WRITER_OP_CODES = {
    "between": 0,
    "not_between": 1,
    "equal": 2,
    "not_equal": 3,
    "less_than": 4,
    "less_than_or_equal": 5,
    "greater_than": 6,
    "greater_than_or_equal": 7,
}

#: Valid operator strings accepted by ``SheetWriter.add_data_validation_numeric``
#: and ``SheetWriter.add_conditional_format_cell_is``. Exposed as a
#: public frozenset so callers can introspect without reading tests.
CF_OPERATORS = frozenset(_DV_WRITER_OP_CODES.keys())


def _sheet_add_data_validation_numeric(
    self: "SheetWriter",
    range_str: str,
    kind: str,
    op: str,
    formula1: str,
    formula2: str | None = None,
) -> None:
    """Attach a numeric / date / time / text-length data validation.

    ``kind`` is one of ``"whole"``, ``"decimal"``, ``"date"``,
    ``"time"``, ``"text_length"``. ``op`` is one of ``"between"``,
    ``"not_between"``, ``"equal"``, ``"not_equal"``, ``"less_than"``,
    ``"less_than_or_equal"``, ``"greater_than"``,
    ``"greater_than_or_equal"``. ``formula2`` is required for
    ``between`` / ``not_between`` and must be ``None`` for the others
    (the C side rejects mismatches with ``InvalidDataValidation``).

    Raises :class:`ZlsxError` on invalid range / formula /
    two-formula mismatch, :class:`ValueError` on unknown kind / op."""
    self._require_handle()
    if not _ffi._HAS_DATA_VALIDATION_EXT:
        raise RuntimeError(
            "loaded libzlsx does not expose add_data_validation_numeric "
            "(requires 0.2.6+); upgrade libzlsx"
        )
    kind_code = _DV_WRITER_KIND_CODES.get(kind)
    if kind_code is None:
        raise ValueError(
            f"unknown data validation kind {kind!r}; expected one of "
            f"{sorted(_DV_WRITER_KIND_CODES)}"
        )
    op_code = _DV_WRITER_OP_CODES.get(op)
    if op_code is None:
        raise ValueError(
            f"unknown data validation operator {op!r}; expected one of "
            f"{sorted(_DV_WRITER_OP_CODES)}"
        )

    range_raw = range_str.encode("utf-8")
    range_buf = (ctypes.c_ubyte * max(len(range_raw), 1)).from_buffer_copy(
        range_raw or b"\x00"
    )
    ptr_t = ctypes.POINTER(ctypes.c_ubyte)

    f1_raw = formula1.encode("utf-8")
    f1_buf = (ctypes.c_ubyte * max(len(f1_raw), 1)).from_buffer_copy(
        f1_raw or b"\x00"
    )

    if formula2 is None:
        f2_ptr = ptr_t()  # NULL
        f2_len = 0
        f2_buf = None
    else:
        f2_raw = formula2.encode("utf-8")
        f2_buf = (ctypes.c_ubyte * max(len(f2_raw), 1)).from_buffer_copy(
            f2_raw or b"\x00"
        )
        f2_ptr = ctypes.cast(f2_buf, ptr_t)
        f2_len = len(f2_raw)

    rc = _ffi.lib.zlsx_sheet_writer_add_data_validation_numeric(
        self._handle,
        ctypes.cast(range_buf, ptr_t),
        len(range_raw),
        kind_code,
        op_code,
        ctypes.cast(f1_buf, ptr_t),
        len(f1_raw),
        f2_ptr,
        f2_len,
        self._err,
        _ERR_BUF_LEN,
    )
    del range_buf, f1_buf, f2_buf
    if rc != 0:
        raise ZlsxError(
            f"zlsx_sheet_writer_add_data_validation_numeric: {_decode_err(self._err)}"
        )


def _sheet_add_data_validation_custom(
    self: "SheetWriter",
    range_str: str,
    formula: str,
) -> None:
    """Attach a custom-formula data validation. ``formula`` is any
    Excel-parseable boolean expression — referenced cells get xml-
    escaped on emit. Empty formula raises :class:`ZlsxError`
    (``InvalidDataValidation``)."""
    self._require_handle()
    if not _ffi._HAS_DATA_VALIDATION_EXT:
        raise RuntimeError(
            "loaded libzlsx does not expose add_data_validation_custom "
            "(requires 0.2.6+); upgrade libzlsx"
        )
    range_raw = range_str.encode("utf-8")
    range_buf = (ctypes.c_ubyte * max(len(range_raw), 1)).from_buffer_copy(
        range_raw or b"\x00"
    )
    formula_raw = formula.encode("utf-8")
    formula_buf = (ctypes.c_ubyte * max(len(formula_raw), 1)).from_buffer_copy(
        formula_raw or b"\x00"
    )
    ptr_t = ctypes.POINTER(ctypes.c_ubyte)
    rc = _ffi.lib.zlsx_sheet_writer_add_data_validation_custom(
        self._handle,
        ctypes.cast(range_buf, ptr_t),
        len(range_raw),
        ctypes.cast(formula_buf, ptr_t),
        len(formula_raw),
        self._err,
        _ERR_BUF_LEN,
    )
    del range_buf, formula_buf
    if rc != 0:
        raise ZlsxError(
            f"zlsx_sheet_writer_add_data_validation_custom: {_decode_err(self._err)}"
        )


def _sheet_add_conditional_format_cell_is(
    self: "SheetWriter",
    range_str: str,
    op: str,
    formula1: str,
    formula2: str | None,
    dxf_id: int,
) -> None:
    """Attach a cellIs-type conditional-format rule. ``op`` is a
    writer-DV-style operator string (``"between"`` / ``"equal"`` /
    ``"greater_than"`` etc.). ``formula2`` is required for
    ``"between"`` / ``"not_between"`` and must be None otherwise.
    ``dxf_id`` comes from :meth:`Writer.add_dxf`.

    Requires libzlsx 0.2.6+."""
    self._require_handle()
    if not _ffi._HAS_CONDITIONAL_FORMAT:
        raise RuntimeError(
            "loaded libzlsx does not expose add_conditional_format_cell_is "
            "(requires 0.2.6+); upgrade libzlsx"
        )
    op_code = _DV_WRITER_OP_CODES.get(op)
    if op_code is None:
        raise ValueError(
            f"unknown conditional-format operator {op!r}; expected one of "
            f"{sorted(_DV_WRITER_OP_CODES)}"
        )
    range_raw = range_str.encode("utf-8")
    range_buf = (ctypes.c_ubyte * max(len(range_raw), 1)).from_buffer_copy(
        range_raw or b"\x00"
    )
    f1_raw = formula1.encode("utf-8")
    f1_buf = (ctypes.c_ubyte * max(len(f1_raw), 1)).from_buffer_copy(
        f1_raw or b"\x00"
    )
    ptr_t = ctypes.POINTER(ctypes.c_ubyte)
    if formula2 is None:
        f2_ptr = ptr_t()
        f2_len = 0
        f2_buf = None
    else:
        f2_raw = formula2.encode("utf-8")
        f2_buf = (ctypes.c_ubyte * max(len(f2_raw), 1)).from_buffer_copy(
            f2_raw or b"\x00"
        )
        f2_ptr = ctypes.cast(f2_buf, ptr_t)
        f2_len = len(f2_raw)

    rc = _ffi.lib.zlsx_sheet_writer_add_conditional_format_cell_is(
        self._handle,
        ctypes.cast(range_buf, ptr_t),
        len(range_raw),
        op_code,
        ctypes.cast(f1_buf, ptr_t),
        len(f1_raw),
        f2_ptr,
        f2_len,
        dxf_id,
        self._err,
        _ERR_BUF_LEN,
    )
    del range_buf, f1_buf, f2_buf
    if rc != 0:
        raise ZlsxError(
            f"zlsx_sheet_writer_add_conditional_format_cell_is: {_decode_err(self._err)}"
        )


def _sheet_add_conditional_format_expression(
    self: "SheetWriter",
    range_str: str,
    formula: str,
    dxf_id: int,
) -> None:
    """Attach an expression-type conditional-format rule. Same error
    semantics as :meth:`add_conditional_format_cell_is` minus the
    operator + formula2."""
    self._require_handle()
    if not _ffi._HAS_CONDITIONAL_FORMAT:
        raise RuntimeError(
            "loaded libzlsx does not expose add_conditional_format_expression "
            "(requires 0.2.6+); upgrade libzlsx"
        )
    range_raw = range_str.encode("utf-8")
    range_buf = (ctypes.c_ubyte * max(len(range_raw), 1)).from_buffer_copy(
        range_raw or b"\x00"
    )
    formula_raw = formula.encode("utf-8")
    formula_buf = (ctypes.c_ubyte * max(len(formula_raw), 1)).from_buffer_copy(
        formula_raw or b"\x00"
    )
    ptr_t = ctypes.POINTER(ctypes.c_ubyte)
    rc = _ffi.lib.zlsx_sheet_writer_add_conditional_format_expression(
        self._handle,
        ctypes.cast(range_buf, ptr_t),
        len(range_raw),
        ctypes.cast(formula_buf, ptr_t),
        len(formula_raw),
        dxf_id,
        self._err,
        _ERR_BUF_LEN,
    )
    del range_buf, formula_buf
    if rc != 0:
        raise ZlsxError(
            f"zlsx_sheet_writer_add_conditional_format_expression: {_decode_err(self._err)}"
        )


SheetWriter.set_column_width = _sheet_set_column_width   # type: ignore[attr-defined]
SheetWriter.set_row_height = _sheet_set_row_height       # type: ignore[attr-defined]
SheetWriter.freeze_panes = _sheet_freeze_panes           # type: ignore[attr-defined]
SheetWriter.freeze_panes_checked = _sheet_freeze_panes_checked  # type: ignore[attr-defined]
SheetWriter.set_auto_filter = _sheet_set_auto_filter     # type: ignore[attr-defined]
SheetWriter.add_merged_cell = _sheet_add_merged_cell     # type: ignore[attr-defined]
SheetWriter.add_hyperlink = _sheet_add_hyperlink         # type: ignore[attr-defined]
SheetWriter.add_internal_hyperlink = _sheet_add_internal_hyperlink  # type: ignore[attr-defined]
SheetWriter.add_comment = _sheet_add_comment             # type: ignore[attr-defined]
SheetWriter.add_data_validation_list = _sheet_add_data_validation_list  # type: ignore[attr-defined]
SheetWriter.add_data_validation_numeric = _sheet_add_data_validation_numeric  # type: ignore[attr-defined]
SheetWriter.add_data_validation_custom = _sheet_add_data_validation_custom  # type: ignore[attr-defined]
SheetWriter.add_conditional_format_cell_is = _sheet_add_conditional_format_cell_is  # type: ignore[attr-defined]
SheetWriter.add_conditional_format_expression = _sheet_add_conditional_format_expression  # type: ignore[attr-defined]


def _sheet_add_conditional_format_color_scale(
    self: "SheetWriter",
    range_str: str,
    low_color_argb: int,
    mid_color_argb: int | None,
    high_color_argb: int,
) -> None:
    """Attach a color-scale conditional format. 3-stop gradient when
    ``mid_color_argb`` is non-None (min → mid @ 50th percentile → max);
    2-stop (min → max) when None. No dxf_id needed.

    Requires libzlsx 0.2.6+."""
    self._require_handle()
    if not _ffi._HAS_CF_GRADIENT:
        raise RuntimeError(
            "loaded libzlsx does not expose add_conditional_format_color_scale "
            "(requires 0.2.6+); upgrade libzlsx"
        )
    range_raw = range_str.encode("utf-8")
    range_buf = (ctypes.c_ubyte * max(len(range_raw), 1)).from_buffer_copy(
        range_raw or b"\x00"
    )
    rc = _ffi.lib.zlsx_sheet_writer_add_conditional_format_color_scale(
        self._handle,
        ctypes.cast(range_buf, ctypes.POINTER(ctypes.c_ubyte)),
        len(range_raw),
        int(low_color_argb),
        1 if mid_color_argb is not None else 0,
        int(mid_color_argb) if mid_color_argb is not None else 0,
        int(high_color_argb),
        self._err,
        _ERR_BUF_LEN,
    )
    del range_buf
    if rc != 0:
        raise ZlsxError(
            f"zlsx_sheet_writer_add_conditional_format_color_scale: {_decode_err(self._err)}"
        )


def _sheet_add_conditional_format_data_bar(
    self: "SheetWriter",
    range_str: str,
    color_argb: int,
) -> None:
    """Attach a data-bar conditional format. `color_argb` is the bar
    fill — Excel's default is ``0xFF638EC6``."""
    self._require_handle()
    if not _ffi._HAS_CF_GRADIENT:
        raise RuntimeError(
            "loaded libzlsx does not expose add_conditional_format_data_bar "
            "(requires 0.2.6+); upgrade libzlsx"
        )
    range_raw = range_str.encode("utf-8")
    range_buf = (ctypes.c_ubyte * max(len(range_raw), 1)).from_buffer_copy(
        range_raw or b"\x00"
    )
    rc = _ffi.lib.zlsx_sheet_writer_add_conditional_format_data_bar(
        self._handle,
        ctypes.cast(range_buf, ctypes.POINTER(ctypes.c_ubyte)),
        len(range_raw),
        int(color_argb),
        self._err,
        _ERR_BUF_LEN,
    )
    del range_buf
    if rc != 0:
        raise ZlsxError(
            f"zlsx_sheet_writer_add_conditional_format_data_bar: {_decode_err(self._err)}"
        )


SheetWriter.add_conditional_format_color_scale = _sheet_add_conditional_format_color_scale  # type: ignore[attr-defined]
SheetWriter.add_conditional_format_data_bar = _sheet_add_conditional_format_data_bar  # type: ignore[attr-defined]


class Writer:
    """A xlsx workbook under construction.

    Use :func:`zlsx.write` to construct one. Finalise by calling
    :meth:`save` with a target path, then :meth:`close` to release
    resources. The context-manager protocol wraps this: ``with
    zlsx.write("out.xlsx") as w:`` saves automatically on clean exit.

    Writes strings, integers, floats, booleans, and empties; styles
    via :meth:`add_style` (bold/italic, fonts, fills, borders,
    alignment, wrap, number formats) and conditional-format dxfs via
    :meth:`add_dxf`. Per-sheet attachments include
    ``set_column_width``, ``freeze_panes``, ``set_auto_filter``,
    ``add_merged_cell``, ``add_hyperlink`` (external URLs),
    ``add_internal_hyperlink`` (workbook-internal targets),
    ``add_comment``, ``add_data_validation_{list,numeric,custom}``,
    ``add_conditional_format_{cell_is,expression,color_scale,data_bar}``,
    ``write_rich_row`` for inline rich-text runs, and
    ``write_row_with_formulas`` for formula cells with cached values.
    Editing an existing workbook is :class:`Editor` / :func:`edit`.
    """

    def __init__(self, path: Union[str, Path, None] = None):
        self._path = Path(path) if path is not None else None
        self._err = ctypes.create_string_buffer(_ERR_BUF_LEN)
        self._handle = _ffi.lib.zlsx_writer_create(self._err, _ERR_BUF_LEN)
        if not self._handle:
            raise ZlsxError(f"zlsx_writer_create: {_decode_err(self._err)}")
        # Track sheets so we can surface their names in Python.
        self._sheets: list[SheetWriter] = []

    def add_sheet(self, name: str) -> SheetWriter:
        """Add a sheet. The returned :class:`SheetWriter` is owned by
        this Writer — it becomes invalid after :meth:`close` (or the
        end of a ``with`` block)."""
        name_bytes = name.encode("utf-8")
        sw_handle = _ffi.lib.zlsx_writer_add_sheet(
            self._handle, name_bytes, len(name_bytes), self._err, _ERR_BUF_LEN
        )
        if not sw_handle:
            raise ZlsxError(f"zlsx_writer_add_sheet({name!r}): {_decode_err(self._err)}")
        sw = SheetWriter(self, sw_handle)
        self._sheets.append(sw)
        return sw

    def add_defined_name(
        self,
        name: str,
        refers_to: str,
        local_sheet_id: int | None = None,
        hidden: bool = False,
    ) -> None:
        """Register a workbook-level (or sheet-scoped) defined name.

        ``local_sheet_id`` ``None`` means workbook-scope; an integer
        ≥ 0 means a 0-based sheet index (must resolve at :meth:`save`
        time). ``hidden=True`` emits ``hidden="1"`` (the convention
        for ``_xlnm.Print_Area`` and similar built-in names).

        Raises :class:`ZlsxError` on ``InvalidDefinedName``,
        ``InvalidDefinedNameRefersTo``, or ``DuplicateDefinedName``
        (case-insensitive duplicate within the same scope).
        """
        if self._handle is None:
            raise ZlsxError("Writer is closed")
        if not _ffi._HAS_DEFINED_NAME:
            raise RuntimeError(
                "loaded libzlsx does not expose add_defined_name "
                "(requires post-0.3.0 libzlsx); upgrade libzlsx"
            )
        name_raw = name.encode("utf-8")
        refers_raw = refers_to.encode("utf-8")
        name_buf = (ctypes.c_ubyte * max(len(name_raw), 1)).from_buffer_copy(
            name_raw or b"\x00"
        )
        refers_buf = (ctypes.c_ubyte * max(len(refers_raw), 1)).from_buffer_copy(
            refers_raw or b"\x00"
        )
        ptr_t = ctypes.POINTER(ctypes.c_ubyte)
        # Negative sheet id signals workbook scope to the C ABI.
        lsi: int = -1 if local_sheet_id is None else int(local_sheet_id)
        if local_sheet_id is not None:
            if local_sheet_id < 0:
                raise ValueError(
                    f"local_sheet_id must be >= 0 or None, got {local_sheet_id}"
                )
            # The C ABI uses a signed int32 with negative = workbook
            # scope. Values >= 2**31 wrap to negative through ctypes
            # and would silently turn a sheet-scoped name into a
            # workbook-scoped one. Reject above the int32 cap.
            if local_sheet_id > 0x7FFFFFFF:
                raise ValueError(
                    f"local_sheet_id must be <= INT32_MAX, got {local_sheet_id}"
                )
        rc = _ffi.lib.zlsx_writer_add_defined_name(
            self._handle,
            ctypes.cast(name_buf, ptr_t),
            len(name_raw),
            ctypes.cast(refers_buf, ptr_t),
            len(refers_raw),
            lsi,
            1 if hidden else 0,
            self._err,
            _ERR_BUF_LEN,
        )
        del name_buf, refers_buf
        if rc != 0:
            raise ZlsxError(
                f"zlsx_writer_add_defined_name({name!r}): {_decode_err(self._err)}"
            )

    def add_style(self, style: "Style") -> int:
        """Register a cell style and return its 1-based index. Pass the
        returned value via ``styles=[…]`` to :meth:`SheetWriter.write_row`.
        Duplicate registrations return the same index.

        If the Style only sets ``font_bold``/``font_italic`` we call the
        stage-1 ``zlsx_writer_add_style`` for backward compatibility with
        libzlsx 0.2.3. Any stage-2 field (size, name, color, alignment,
        wrap_text) promotes the call to ``zlsx_writer_add_style_ex``
        (libzlsx 0.2.4+)."""
        if not _ffi._HAS_STYLES:
            raise RuntimeError(
                "loaded libzlsx does not expose zlsx_writer_add_style "
                "(requires 0.2.4+); upgrade libzlsx"
            )

        has_border = (
            style.border_left.style != "none"
            or style.border_right.style != "none"
            or style.border_top.style != "none"
            or style.border_bottom.style != "none"
            or style.border_diagonal.style != "none"
            or style.diagonal_up
            or style.diagonal_down
        )

        needs_ex = (
            style.font_size is not None
            or style.font_name is not None
            or style.font_color_argb is not None
            or style.alignment_horizontal != "general"
            or style.wrap_text
            or style.fill_pattern != "none"
            or style.fill_fg_argb is not None
            or style.fill_bg_argb is not None
            or has_border
            or style.number_format is not None
        )

        out_idx = ctypes.c_uint32(0)

        if not needs_ex:
            rc = _ffi.lib.zlsx_writer_add_style(
                self._handle,
                1 if style.font_bold else 0,
                1 if style.font_italic else 0,
                ctypes.byref(out_idx),
                self._err,
                _ERR_BUF_LEN,
            )
            if rc != 0:
                raise ZlsxError(f"zlsx_writer_add_style: {_decode_err(self._err)}")
            return int(out_idx.value)

        if not _ffi._HAS_STYLES_EX:
            raise RuntimeError(
                "loaded libzlsx does not expose zlsx_writer_add_style_ex "
                "(requires 0.2.4+) — stage-2 style fields need the newer dylib"
            )

        flags = 0
        if style.font_size is not None:
            flags |= _ffi.FONT_SIZE_SET
        if style.font_color_argb is not None:
            flags |= _ffi.FONT_COLOR_SET
        if style.fill_fg_argb is not None:
            flags |= _ffi.FILL_FG_SET
        if style.fill_bg_argb is not None:
            flags |= _ffi.FILL_BG_SET

        # Distinguish "unset" (None) from "empty string" — the latter
        # is invalid and must reach the Zig side as font_name_len=0
        # with an explicit sentinel that triggers InvalidFontName.
        if style.font_name is None:
            name_bytes = b""
        elif style.font_name == "":
            raise ZlsxError("InvalidFontName")
        else:
            name_bytes = style.font_name.encode("utf-8")

        if style.number_format is None:
            num_fmt_bytes = b""
        elif style.number_format == "":
            raise ZlsxError("InvalidNumberFormat")
        else:
            num_fmt_bytes = style.number_format.encode("utf-8")
        # Keep the bytes buffer alive through the FFI call.
        name_buf = (ctypes.c_ubyte * max(len(name_bytes), 1)).from_buffer_copy(
            name_bytes or b"\x00"
        )
        num_fmt_buf = (ctypes.c_ubyte * max(len(num_fmt_bytes), 1)).from_buffer_copy(
            num_fmt_bytes or b"\x00"
        )

        if style.alignment_horizontal not in _HALIGN_VALUES:
            raise ValueError(
                f"unknown alignment_horizontal: {style.alignment_horizontal!r}"
            )
        if style.fill_pattern not in _PATTERN_VALUES:
            raise ValueError(
                f"unknown fill_pattern: {style.fill_pattern!r}"
            )

        def _bstyle(side: BorderSide) -> int:
            if side.style not in _BORDER_STYLE_VALUES:
                raise ValueError(f"unknown border style: {side.style!r}")
            return _BORDER_STYLE_VALUES[side.style]

        flags2 = 0
        if style.border_left.color_argb is not None:
            flags2 |= _ffi.BORDER_LEFT_COLOR_SET
        if style.border_right.color_argb is not None:
            flags2 |= _ffi.BORDER_RIGHT_COLOR_SET
        if style.border_top.color_argb is not None:
            flags2 |= _ffi.BORDER_TOP_COLOR_SET
        if style.border_bottom.color_argb is not None:
            flags2 |= _ffi.BORDER_BOTTOM_COLOR_SET
        if style.border_diagonal.color_argb is not None:
            flags2 |= _ffi.BORDER_DIAGONAL_COLOR_SET

        spec = _ffi.CStyle(
            font_bold=1 if style.font_bold else 0,
            font_italic=1 if style.font_italic else 0,
            alignment_horizontal=_HALIGN_VALUES[style.alignment_horizontal],
            wrap_text=1 if style.wrap_text else 0,
            flags=flags,
            fill_pattern=_PATTERN_VALUES[style.fill_pattern],
            flags2=flags2,
            font_size=float(style.font_size or 0.0),
            font_color_argb=_check_argb("font_color_argb", style.font_color_argb),
            fill_fg_argb=_check_argb("fill_fg_argb", style.fill_fg_argb),
            fill_bg_argb=_check_argb("fill_bg_argb", style.fill_bg_argb),
            border_left_style=_bstyle(style.border_left),
            border_right_style=_bstyle(style.border_right),
            border_top_style=_bstyle(style.border_top),
            border_bottom_style=_bstyle(style.border_bottom),
            border_diagonal_style=_bstyle(style.border_diagonal),
            diagonal_up=1 if style.diagonal_up else 0,
            diagonal_down=1 if style.diagonal_down else 0,
            border_left_color_argb=_check_argb("border_left.color_argb", style.border_left.color_argb),
            border_right_color_argb=_check_argb("border_right.color_argb", style.border_right.color_argb),
            border_top_color_argb=_check_argb("border_top.color_argb", style.border_top.color_argb),
            border_bottom_color_argb=_check_argb("border_bottom.color_argb", style.border_bottom.color_argb),
            border_diagonal_color_argb=_check_argb("border_diagonal.color_argb", style.border_diagonal.color_argb),
            font_name_ptr=ctypes.cast(name_buf, ctypes.POINTER(ctypes.c_ubyte)),
            font_name_len=len(name_bytes),
            num_fmt_ptr=ctypes.cast(num_fmt_buf, ctypes.POINTER(ctypes.c_ubyte)),
            num_fmt_len=len(num_fmt_bytes),
        )
        rc = _ffi.lib.zlsx_writer_add_style_ex(
            self._handle,
            ctypes.byref(spec),
            ctypes.byref(out_idx),
            self._err,
            _ERR_BUF_LEN,
        )
        # Keep name_buf alive until the call returns.
        del name_buf
        if rc != 0:
            raise ZlsxError(f"zlsx_writer_add_style_ex: {_decode_err(self._err)}")
        return int(out_idx.value)

    def save(
        self,
        path: Union[str, Path, None] = None,
        recalculate: "Optional[RecalcOptions]" = None,
    ) -> "Optional[RecalcReport]":
        """Write the workbook to disk. Uses the path passed to
        :func:`zlsx.write` if none is provided here.

        ``recalculate=RecalcOptions(...)`` routes the save through the
        recalc orchestrator (§5.7.9): the fresh archive is emitted to
        memory, opened, recalculated, and committed atomically — every
        cached formula value in the destination is one the engine
        computed. Returns the :class:`RecalcReport` then, ``None``
        otherwise. Requires libzlsx 0.8.0+."""
        target = Path(path) if path is not None else self._path
        if target is None:
            raise ValueError("no save path: pass one to zlsx.write() or Writer.save()")
        raw = str(target).encode("utf-8")
        if recalculate is None:
            rc = _ffi.lib.zlsx_writer_save(
                self._handle, raw, len(raw), self._err, _ERR_BUF_LEN
            )
            if rc != 0:
                raise ZlsxError(f"zlsx_writer_save({target!r}): {_decode_err(self._err)}")
            return None
        if not _ffi._HAS_WRITER_RECALC:
            raise RuntimeError(
                "loaded libzlsx does not expose zlsx_writer_save_with_recalc "
                "(requires 0.8.0+); upgrade libzlsx"
            )
        opts = recalculate
        token = _new_cancel_token(self._err)
        try:
            run = _build_run(opts.now, opts.utc_offset_min, opts.seed, opts.mode,
                             opts.profile, "da", opts.on_unsupported, opts.timeout,
                             token)
            report, diag = _fresh_report_and_diag()
            pptr, pkeep = _path_as_ubyte(raw)

            def invoke():
                return _ffi.lib.zlsx_writer_save_with_recalc(
                    self._handle, pptr, len(raw), ctypes.byref(run),
                    ctypes.byref(report), ctypes.byref(diag),
                    self._err, _ERR_BUF_LEN,
                )

            result = _drive_recalc("zlsx_writer_save_with_recalc", invoke,
                                   token, opts.timeout, report, diag, self._err)
            del pkeep
            return result
        finally:
            if token is not None:
                _ffi.lib.zlsx_cancel_token_free(token)

    def to_bytes(self) -> bytes:
        """Serialise the workbook and return it as ``bytes`` instead of
        writing a file — the writer-side mirror of :func:`open_bytes`.

        Byte-for-byte identical to what :meth:`save` would have written,
        so anything that consumes an xlsx file consumes this. Use it when
        there is no usable filesystem to save through — a Spark executor
        writing to object storage, an upload body, an in-process pipeline
        that never wants a temp file::

            with zlsx.write() as w:
                w.add_sheet("Report").write_row(["ok"])
                payload = w.to_bytes()

        The Writer stays usable: append more rows and call again, or
        :meth:`save` to disk as well. Requires libzlsx 0.8.0+.
        """
        if not _ffi._HAS_SAVE_TO_BUFFER:
            raise RuntimeError(
                "loaded libzlsx does not expose writer save-to-buffer "
                "(requires 0.8.0+); upgrade libzlsx"
            )
        if not self._handle:
            raise ZlsxError("writer is closed")

        out_ptr = ctypes.POINTER(ctypes.c_ubyte)()
        out_len = ctypes.c_size_t(0)
        rc = _ffi.lib.zlsx_writer_save_to_buffer(
            self._handle,
            ctypes.byref(out_ptr),
            ctypes.byref(out_len),
            self._err,
            _ERR_BUF_LEN,
        )
        if rc != 0:
            raise ZlsxError(f"zlsx_writer_save_to_buffer: {_decode_err(self._err)}")
        try:
            # One copy, at a length the callee reported. ctypes.string_at
            # is length-delimited, so embedded NULs (every deflate stream
            # has them) survive.
            return ctypes.string_at(out_ptr, out_len.value)
        finally:
            _ffi.lib.zlsx_buffer_free(out_ptr, out_len)

    def add_dxf(self, dxf: "Dxf") -> int:
        """Register a differential format for conditional-formatting
        rules and return its dxf id. Content-dedup'd — same
        :class:`Dxf` returns the same id. Requires libzlsx 0.2.6+."""
        if not _ffi._HAS_CONDITIONAL_FORMAT:
            raise RuntimeError(
                "loaded libzlsx does not expose add_dxf "
                "(requires 0.2.6+); upgrade libzlsx"
            )

        def _side(s: "BorderSide") -> "_ffi.CDxfBorderSide":
            style_code = _BORDER_STYLE_VALUES.get(s.style, 0)
            return _ffi.CDxfBorderSide(
                style=style_code,
                has_color=1 if s.color_argb is not None else 0,
                _pad=(ctypes.c_uint8 * 2)(0, 0),
                color_argb=s.color_argb or 0,
            )

        c = _ffi.CDxf(
            bold=1 if dxf.font_bold else 0,
            italic=1 if dxf.font_italic else 0,
            has_color=1 if dxf.font_color_argb is not None else 0,
            has_fill=1 if dxf.fill_fg_argb is not None else 0,
            color_argb=dxf.font_color_argb or 0,
            fill_fg_argb=dxf.fill_fg_argb or 0,
            has_size=1 if dxf.font_size is not None else 0,
            _pad=(ctypes.c_uint8 * 3)(0, 0, 0),
            size=dxf.font_size if dxf.font_size is not None else 0.0,
            border_left=_side(dxf.border_left),
            border_right=_side(dxf.border_right),
            border_top=_side(dxf.border_top),
            border_bottom=_side(dxf.border_bottom),
        )
        out_id = ctypes.c_uint32(0)
        rc = _ffi.lib.zlsx_writer_add_dxf(
            self._handle,
            ctypes.byref(c),
            ctypes.byref(out_id),
            self._err,
            _ERR_BUF_LEN,
        )
        if rc != 0:
            raise ZlsxError(f"zlsx_writer_add_dxf: {_decode_err(self._err)}")
        return int(out_id.value)

    def close(self) -> None:
        """Release all writer state. Any :class:`SheetWriter` obtained
        from this Writer becomes invalid after close()."""
        if self._handle:
            _ffi.lib.zlsx_writer_close(self._handle)
            self._handle = None
            for sw in self._sheets:
                sw._handle = None
            self._sheets.clear()

    def __enter__(self) -> "Writer":
        return self

    def __exit__(self, exc_type, *exc_info) -> None:
        # Save on clean exit; propagate any exception. Always close.
        try:
            if exc_type is None and self._path is not None:
                self.save()
        finally:
            self.close()

    def __del__(self) -> None:
        try:
            self.close()
        except Exception:
            pass


def write(path: Union[str, Path, None] = None) -> Writer:
    """Begin a new xlsx workbook.

    If ``path`` is provided and this Writer is used as a context
    manager, the workbook is saved automatically on clean exit::

        with zlsx.write("out.xlsx") as w:
            sheet = w.add_sheet("Summary")
            sheet.write_row(["Name", "Count"])
            sheet.write_row(["Alice", 42])
    """
    return Writer(path)


# ─── Formula engine (M9a2) ───────────────────────────────────────────
#
# The Python face of the zlsx_status_v1 exports: recalculate / evaluate
# / save_with_recalc / save_to_buffer / mark_recalc_on_load on Editor,
# Writer.save(recalculate=...), and FormulaSpec on
# SheetWriter.write_row_with_formulas. Every call that receives
# library-owned memory releases it in a ``finally``; cancellable calls
# run on a worker thread so Ctrl-C and timeouts reach the engine as a
# token trigger rather than a blocked signal handler.


class ExcelError(str):
    """An Excel error *value* (e.g. ``#DIV/0!``) — a successful result,
    never a Python exception (``ZlsxError`` stays the exception type).
    Behaves as its spelling; ``isinstance(x, ExcelError)`` is the test."""

    __slots__ = ()

    def __repr__(self) -> str:  # pragma: no cover - cosmetic
        return f"ExcelError({str.__repr__(self)})"


@dataclass(frozen=True)
class CensusEntry:
    """One §5.7.7 census entry: a construct the engine could not
    implement, and where it was. ``row`` is 1-based; 0 means the entry
    is not about a cell. ``col`` is 0-based."""

    plane: str
    sheet: int
    row: int
    col: int


class ZlsxFormulaRefusal(ZlsxError):
    """A typed Plane-2 refusal (status -2): the engine refused the run
    rather than guessing. ``error_name`` is the Zig error name
    (e.g. ``"FormulaUnsupportedFunction"``); ``cells`` lists the
    refusing cells as ``(sheet, row, col)`` (row 1-based, col 0-based);
    ``census`` carries the full entries."""

    def __init__(
        self,
        error_name: str,
        cells: List[Tuple[int, int, int]],
        census: List[CensusEntry],
    ):
        at = f" at {len(cells)} cell(s)" if cells else ""
        super().__init__(f"{error_name}{at}")
        self.error_name = error_name
        self.cells = cells
        self.census = census


@dataclass(frozen=True)
class Resolved:
    """The §5.5 echo: the exact context a run resolved to. Replaying
    these values reproduces the run exactly — including one whose
    ``now``/``seed`` were defaulted by the binding."""

    now: int
    utc_offset_min: int
    seed: int
    mode: str
    profile: str
    dialect: Optional[str]
    anchor: Optional[Tuple[int, int]] = None


@dataclass(frozen=True)
class Matrix:
    """A rectangular evaluation result. ``cells`` is a list of rows;
    each cell is ``float``/``str``/``bool``/:class:`ExcelError` (blanks
    publish as 0, §12.2)."""

    rows: int
    cols: int
    cells: list


@dataclass(frozen=True)
class EvalResult:
    """What :meth:`Editor.evaluate` returns: the value plus the
    resolved-context echo that makes it reproducible."""

    value: object
    resolved: Resolved


@dataclass(frozen=True)
class RecalcReport:
    """What a recalculation did (§5.7.8). ``cancelled_late`` is
    binding-level truth: the cancellation (Ctrl-C or timeout) arrived
    only after the commit point, so the transaction completed."""

    sheets_patched: int
    cells_written: int
    passes: int
    non_converged_cells: int
    dynamic_passes: int
    kept_stale: bool
    calc_chain_removed: bool
    census: List[CensusEntry]
    census_truncated: bool
    retained_generations: int
    retained_bytes: int
    durability_warning: bool
    durability_errno: int
    resolved: Optional[Resolved]
    cancelled_late: bool = False


@dataclass(frozen=True)
class RecalcOptions:
    """The recalc context :meth:`Writer.save` composes through the
    orchestrator. Field semantics match :meth:`Editor.recalculate`."""

    now: object = None
    utc_offset_min: int = 0
    seed: Optional[int] = None
    mode: str = "excel"
    profile: str = "windows_1252"
    on_unsupported: str = "refuse"
    timeout: Optional[float] = None


@dataclass(frozen=True)
class FormulaSpec:
    """One cell's formula for ``write_row_with_formulas``: the text
    (no leading ``=``), the dialect, and — for ``"cse"`` only — the
    declared rectangle whose top-left must be the carrying cell."""

    text: str
    dialect: str = "scalar"
    ref: Optional[str] = None

    def __post_init__(self):
        if self.dialect not in ("scalar", "dynamic_array", "cse"):
            raise ValueError(f"unknown formula dialect {self.dialect!r}")
        if (self.dialect == "cse") != (self.ref is not None):
            raise ValueError("ref is required for dialect='cse' and illegal otherwise")
        if not self.text:
            raise ValueError("formula text must be non-empty; pass None for a value cell")

    @classmethod
    def cse(cls, text: str, ref: str) -> "FormulaSpec":
        return cls(text, "cse", ref)


_MODE_CODES = {"excel": 0, "ieee": 1}
_PROFILE_CODES = {"windows_1252": 0}
_DIALECT_CODES = {"da": 0, "dynamic_array": 0, "legacy": 1}
_ON_UNSUPPORTED_CODES = {"refuse": 0, "keep_stale_and_mark": 1}
_FORMULA_DIALECT_CODES = {
    "scalar": _ffi.ZLSX_FORMULA_SCALAR,
    "dynamic_array": _ffi.ZLSX_FORMULA_DYNAMIC_ARRAY,
    "cse": _ffi.ZLSX_FORMULA_CSE,
}


def engine_fingerprint() -> str:
    """The engine identity string (§12.4): semver + rule versions +
    target triple + build hash. Two processes may share recalc results
    only when these match."""
    if not _ffi._HAS_FINGERPRINT:
        raise RuntimeError(
            "loaded libzlsx does not expose zlsx_engine_fingerprint "
            "(requires 0.8.0+); upgrade libzlsx"
        )
    return _ffi.lib.zlsx_engine_fingerprint().decode("utf-8")


def _code_for(table, value, what):
    try:
        return table[value]
    except KeyError:
        raise ValueError(f"unknown {what} {value!r}; expected one of {sorted(table)}") from None


def _resolve_now(now) -> int:
    """None = the binding reads the clock (caller-side acquisition; the
    library itself never does — §5.5). datetime = its epoch millis
    (naive datetimes use the local zone); int/float = epoch millis."""
    if now is None:
        return time.time_ns() // 1_000_000
    if isinstance(now, _datetime):
        return int(now.timestamp() * 1000)
    return int(now)


def _resolve_seed(seed) -> int:
    if seed is None:
        return int.from_bytes(os.urandom(8), "little")
    return int(seed) & 0xFFFFFFFFFFFFFFFF


def _build_run(now, utc_offset_min, seed, mode, profile, dialect, on_unsupported, timeout, token):
    run = _ffi.RunV1()
    run.struct_size = ctypes.sizeof(_ffi.RunV1)
    run.now_utc_ms = _resolve_now(now)
    run.rng_seed = _resolve_seed(seed)
    run.utc_offset_min = int(utc_offset_min)
    run.fidelity = _code_for(_MODE_CODES, mode, "mode")
    run.profile = _code_for(_PROFILE_CODES, profile, "profile")
    run.dialect = _code_for(_DIALECT_CODES, dialect, "dialect")
    run.on_unsupported = _code_for(_ON_UNSUPPORTED_CODES, on_unsupported, "on_unsupported")
    if timeout is not None:
        if timeout <= 0:
            raise ValueError("timeout must be positive seconds")
        run.timeout_ms = max(1, int(timeout * 1000))
    if token is not None:
        run.cancel = token
    return run


def _new_cancel_token(err):
    """A fresh cancel token, or None when the dylib predates the API
    (the call then simply runs to completion on the main thread)."""
    if not _ffi._HAS_CANCEL:
        return None
    tok = _ffi.cancel_token_handle()
    rc = _ffi.lib.zlsx_cancel_token_new(ctypes.byref(tok), err, _ERR_BUF_LEN)
    if rc != _ffi.ZLSX_OK:
        raise ZlsxError(f"zlsx_cancel_token_new: {_decode_err(err)}")
    return tok


def _invoke_cancellable(invoke, token):
    """Run one FFI call on a worker thread and wait interruptibly.

    A Python signal handler cannot run while the main thread sits in a
    synchronous ctypes call, so the call moves off-thread; Ctrl-C
    triggers the token and keeps waiting — the engine observes it at
    its next §5.5 poll point and unwinds. Returns (rc, interrupted).
    """
    if token is None:
        return invoke(), False
    box = []

    def work():
        box.append(invoke())

    t = threading.Thread(target=work, daemon=True)
    t.start()
    interrupted = False
    while True:
        try:
            t.join(0.05)
            if not t.is_alive():
                break
        except KeyboardInterrupt:
            interrupted = True
            _ffi.lib.zlsx_cancel_token_trigger(token)
    if not box:
        raise ZlsxError("worker thread died without a status")
    return box[0], interrupted


def _decode_census(ptr, n) -> List[CensusEntry]:
    out = []
    for i in range(int(n)):
        e = ptr[i]
        plane = (
            _ffi.PLANE_NAMES[e.plane]
            if e.plane < len(_ffi.PLANE_NAMES)
            else f"plane_{e.plane}"
        )
        out.append(CensusEntry(plane=plane, sheet=e.sheet, row=e.row, col=e.col))
    return out


def _refusal_from_diag(diag) -> ZlsxFormulaRefusal:
    # ctypes returns a c_char array field as bytes cut at the first NUL.
    error_name = diag.error_name.decode("utf-8", errors="replace")
    census = _decode_census(diag.census, diag.census_len)
    cells = [(e.sheet, e.row, e.col) for e in census if e.row > 0]
    return ZlsxFormulaRefusal(error_name, cells, census)


def _decode_resolved(cres, anchor=None) -> Resolved:
    dialect = None
    if cres.dialect != _ffi.ZLSX_DIALECT_NONE:
        dialect = "da" if cres.dialect == 0 else "legacy"
    return Resolved(
        now=cres.now_utc_ms,
        utc_offset_min=cres.utc_offset_min,
        seed=cres.rng_seed,
        mode="ieee" if cres.fidelity == 1 else "excel",
        profile="windows_1252",
        dialect=dialect,
        anchor=anchor,
    )


def _decode_report(rep, cancelled_late) -> RecalcReport:
    resolved = _decode_resolved(rep.resolved) if rep.resolved_present else None
    return RecalcReport(
        sheets_patched=rep.sheets_patched,
        cells_written=rep.cells_written,
        passes=rep.passes,
        non_converged_cells=rep.non_converged_cells,
        dynamic_passes=rep.dynamic_passes,
        kept_stale=bool(rep.kept_stale),
        calc_chain_removed=bool(rep.calc_chain_removed),
        census=_decode_census(rep.census, rep.census_len),
        census_truncated=bool(rep.census_truncated),
        retained_generations=rep.retained_generations,
        retained_bytes=rep.retained_bytes,
        durability_warning=bool(rep.durability_warning),
        durability_errno=rep.durability_errno,
        resolved=resolved,
        cancelled_late=cancelled_late,
    )


def _drive_recalc(symbol, invoke, token, timeout, report, diag, err):
    """The shared tail of every recalc-shaped call: worker thread,
    status mapping, and release-in-finally (the M9a2 contract)."""
    started = time.monotonic()
    rc, interrupted = _invoke_cancellable(invoke, token)
    elapsed = time.monotonic() - started
    try:
        if rc == _ffi.ZLSX_OK:
            late = interrupted or (timeout is not None and elapsed >= timeout)
            return _decode_report(report, late)
        if rc == _ffi.ZLSX_CANCELLED:
            # Observed before commit, by the engine's own contract.
            if interrupted:
                raise KeyboardInterrupt
            raise TimeoutError(
                f"{symbol}: cancellation observed before commit (timeout={timeout}s)"
            )
        if rc == _ffi.ZLSX_REFUSED:
            raise _refusal_from_diag(diag)
        raise ZlsxError(f"{symbol}: {_decode_err(err)}")
    finally:
        _ffi.lib.zlsx_recalc_report_release(ctypes.byref(report))
        _ffi.lib.zlsx_diag_release(ctypes.byref(diag))


def _fresh_report_and_diag():
    report = _ffi.RecalcReportV1()
    report.struct_size = ctypes.sizeof(_ffi.RecalcReportV1)
    diag = _ffi.DiagV1()
    diag.struct_size = ctypes.sizeof(_ffi.DiagV1)
    return report, diag


def _path_as_ubyte(pbytes):
    buf = ctypes.create_string_buffer(pbytes, len(pbytes))
    return ctypes.cast(buf, ctypes.POINTER(ctypes.c_ubyte)), buf


class Editor:
    """Open an existing xlsx, append rows, save.

    Append-only v1: cell types are ``None`` / ``bool`` / ``int`` /
    ``float`` / ``str``. Rows are buffered in memory and applied
    atomically on :meth:`save`. The source workbook must already
    carry an ``xl/sharedStrings.xml`` part for string appends —
    workbooks with only inline strings raise ``NoSstInSource``.

    Use as a context manager so the underlying handle is dropped
    deterministically::

        with zlsx.edit("report.xlsx") as ed:
            ed.append_rows(0, [["Carol", 9.5], ["Dave", 7.0]])
            ed.save("report.xlsx")          # overwrite in place

    Single-disk archives only; ZIP64 / multi-disk / encrypted /
    data-descriptor archives are refused at open. Requires
    libzlsx 0.2.7+."""

    def __init__(self, path: Union[str, Path]):
        if not _ffi._HAS_EDITOR:
            raise RuntimeError(
                "loaded libzlsx does not expose Editor (requires 0.2.7+); "
                "upgrade libzlsx"
            )
        self._err = ctypes.create_string_buffer(_ERR_BUF_LEN)
        encoded_path = str(path).encode("utf-8")
        self._handle = _ffi.lib.zlsx_editor_open(encoded_path, self._err, _ERR_BUF_LEN)
        if not self._handle:
            raise ZlsxError(f"zlsx_editor_open: {_decode_err(self._err)}")

    def append_rows(self, sheet_idx: int, rows) -> None:
        """Buffer ``rows`` for append on ``sheet_idx``. Each row is
        an iterable of ``None | bool | int | float | str``. Rows
        are applied at :meth:`save` time, not now; multiple
        ``append_rows`` calls accumulate in order.

        Rows are streamed through the underlying editor one at a
        time, so a generator / DB cursor / CSV iterator can be
        passed without first materialising into a list. Mid-batch
        failure (a Python type error on row N+1, or an FFI rejection
        like ``IntegerExceedsExcelPrecision``) leaves any already-
        buffered rows in the editor — there's no rollback. Save
        only after the whole batch succeeded if you need
        all-or-nothing semantics."""
        if not self._handle:
            raise ZlsxError("editor is closed")
        for row in rows:
            cells_list = list(row)
            n = len(cells_list)
            cells_arg = None
            keepers = []
            if n > 0:
                cell_array = (_ffi.Cell * n)()
                for i, v in enumerate(cells_list):
                    cell, keeper = _py_value_to_cell(v)
                    cell_array[i] = cell
                    if keeper is not None:
                        keepers.append(keeper)
                cells_arg = ctypes.cast(cell_array, _ffi.cell_ptr)
            # Empty rows (n == 0) are forwarded — the native editor
            # emits `<row r="N"></row>` for them, which callers use
            # as visual gaps between data blocks.
            rc = _ffi.lib.zlsx_editor_append_row(
                self._handle,
                int(sheet_idx),
                cells_arg,
                n,
                self._err,
                _ERR_BUF_LEN,
            )
            del keepers
            if rc != 0:
                raise ZlsxError(
                    f"zlsx_editor_append_row: {_decode_err(self._err)}"
                )

    def set_cell(self, sheet_idx: int, row: int, col: int, value) -> None:
        """Replace or insert a single cell on ``sheet_idx`` at
        (``row``, ``col``). ``row`` is 1-based; ``col`` is 0-based
        (A=0, B=1, …). ``value`` follows the same Python → cell
        type mapping as :meth:`append_rows`:

        * ``None`` → empty
        * ``True``/``False`` → boolean
        * ``int`` (within ±2⁵³) → integer
        * ``float`` → number
        * ``str`` → inline-string (no SST dedup; iter-cm-2b)

        Errors propagate as :class:`ZlsxError`. Notable typed errors
        bubble up by name:

        * ``SetCellSourceCellHasMetadata`` — the source cell carries
          ``s="N"`` styles or non-canonical body (formulas,
          phonetic hints, extension blocks). The replacement is
          canonical; preserve-and-merge is iter-cm-2e.
        * ``SheetHasUnsavedAppends`` — :meth:`append_rows` was
          called on the same sheet first; mixing append and
          mutate isn't supported.
        """
        if not self._handle:
            raise ZlsxError("editor is closed")
        if not _ffi._HAS_EDITOR_SET_CELL:
            raise RuntimeError(
                "loaded libzlsx does not expose zlsx_editor_set_cell "
                "(requires 0.2.9+); upgrade libzlsx"
            )
        cell, keeper = _py_value_to_cell(value)
        cell_ref = _ffi.Cell.from_buffer_copy(bytes(cell))
        rc = _ffi.lib.zlsx_editor_set_cell(
            self._handle,
            int(sheet_idx),
            int(row),
            int(col),
            ctypes.byref(cell_ref),
            self._err,
            _ERR_BUF_LEN,
        )
        del keeper  # keep alive until after the FFI call
        if rc != 0:
            raise ZlsxError(f"zlsx_editor_set_cell: {_decode_err(self._err)}")

    def set_cells(self, sheet_idx: int, edits) -> None:
        """Bulk variant of :meth:`set_cell`. ``edits`` is an
        iterable of ``(row, col, value)`` tuples. Applied in source
        order; later edits see the byte offsets produced by earlier
        ones, same as calling :meth:`set_cell` N times. Mid-batch
        failure leaves the editor with the successful prefix
        applied — there's no rollback."""
        for edit in edits:
            row, col, value = edit
            self.set_cell(sheet_idx, row, col, value)

    def doc_props(self) -> dict:
        """Read the workbook's ``docProps`` metadata.

        Returns a dict with one key per field (``creator``,
        ``last_modified_by``, ``company``, …) plus
        ``has_custom_properties``. Absent fields are ``None``.

        These parts round-trip untouched through every edit, so a
        pipeline that masks cell values still ships the original
        author's name unless it scrubs them — see
        :meth:`strip_doc_props`.

        Requires libzlsx 0.5.0+.
        """
        if not _ffi._HAS_DOCPROPS:
            raise ZlsxError("libzlsx too old: zlsx_editor_docprop_at unavailable")

        out: dict = {}
        ptr = ctypes.POINTER(ctypes.c_ubyte)()
        length = ctypes.c_size_t(0)
        for name, field_id in _ffi.DOCPROP_FIELDS.items():
            rc = _ffi.lib.zlsx_editor_docprop_at(
                self._handle,
                ctypes.c_uint32(field_id),
                ctypes.byref(ptr),
                ctypes.byref(length),
            )
            if rc == -2:
                raise ZlsxError("zlsx_editor_docprop_at: could not read docProps")
            if rc != 0:
                raise ZlsxError(f"zlsx_editor_docprop_at({name}): rc={rc}")
            if length.value == 0:
                out[name] = None
            else:
                out[name] = bytes(
                    ctypes.cast(
                        ptr, ctypes.POINTER(ctypes.c_ubyte * length.value)
                    ).contents
                ).decode("utf-8", "replace")

        has_custom = _ffi.lib.zlsx_editor_has_custom_properties(self._handle)
        if has_custom < 0:
            raise ZlsxError("zlsx_editor_has_custom_properties failed")
        out["has_custom_properties"] = bool(has_custom)
        return out

    def strip_doc_props(self, strip_timestamps: bool = False) -> None:
        """Remove identifying document metadata, staged for the next
        :meth:`save`.

        Drops ``dc:creator``, ``cp:lastModifiedBy``, title/subject/
        description/keywords/category, ``Company``, ``Manager``,
        ``HyperlinkBase`` and the whole ``docProps/custom.xml`` part.
        Cell data is untouched.

        ``strip_timestamps`` additionally removes created/modified/
        revision. Those are kept by default: rarely identifying on
        their own, and removing them visibly empties Excel's
        document-info pane.

        Requires libzlsx 0.5.0+.
        """
        if not _ffi._HAS_DOCPROPS:
            raise ZlsxError("libzlsx too old: zlsx_editor_strip_doc_props unavailable")
        rc = _ffi.lib.zlsx_editor_strip_doc_props(
            self._handle,
            ctypes.c_int32(1 if strip_timestamps else 0),
            self._err,
            _ERR_BUF_LEN,
        )
        if rc != 0:
            raise ZlsxError(f"zlsx_editor_strip_doc_props: {_decode_err(self._err)}")

    def save(self, out_path: Union[str, Path]) -> None:
        """Write the (mutated) workbook atomically to ``out_path``.
        Pass the same path as the source to overwrite in place."""
        if not self._handle:
            raise ZlsxError("editor is closed")
        encoded = str(out_path).encode("utf-8")
        rc = _ffi.lib.zlsx_editor_save(
            self._handle,
            encoded,
            len(encoded),
            self._err,
            _ERR_BUF_LEN,
        )
        if rc != 0:
            raise ZlsxError(f"zlsx_editor_save: {_decode_err(self._err)}")

    # ── Formula engine (M9a2) ──────────────────────────────────────

    @classmethod
    def from_bytes(cls, data: bytes) -> "Editor":
        """Open an editor over a workbook already in memory. ``data``
        is copied — the borrow ends when this returns."""
        if not _ffi._HAS_SAVE_BUFFER:
            raise RuntimeError(
                "loaded libzlsx does not expose zlsx_open_buffer "
                "(requires 0.8.0+); upgrade libzlsx"
            )
        obj = cls.__new__(cls)
        obj._err = ctypes.create_string_buffer(_ERR_BUF_LEN)
        buf = (ctypes.c_ubyte * len(data)).from_buffer_copy(data)
        handle = _ffi.editor_handle()
        rc = _ffi.lib.zlsx_open_buffer(
            buf, len(data), ctypes.byref(handle), obj._err, _ERR_BUF_LEN
        )
        if rc != _ffi.ZLSX_OK:
            raise ZlsxError(f"zlsx_open_buffer: {_decode_err(obj._err)}")
        obj._handle = handle.value
        return obj

    def save_to_buffer(self) -> bytes:
        """Serialize the editor's current state — staged mutations
        included — to memory. An untouched editor returns the source
        bytes verbatim."""
        if not self._handle:
            raise ZlsxError("editor is closed")
        if not _ffi._HAS_SAVE_BUFFER:
            raise RuntimeError(
                "loaded libzlsx does not expose zlsx_editor_save_to_buffer "
                "(requires 0.8.0+); upgrade libzlsx"
            )
        out_ptr = ctypes.POINTER(ctypes.c_ubyte)()
        out_len = ctypes.c_size_t(0)
        rc = _ffi.lib.zlsx_editor_save_to_buffer(
            self._handle, ctypes.byref(out_ptr), ctypes.byref(out_len),
            self._err, _ERR_BUF_LEN,
        )
        if rc != _ffi.ZLSX_OK:
            raise ZlsxError(f"zlsx_editor_save_to_buffer: {_decode_err(self._err)}")
        try:
            # One copy at the reported length; string_at is
            # length-delimited so deflate's embedded NULs survive.
            return ctypes.string_at(out_ptr, out_len.value)
        finally:
            _ffi.lib.zlsx_buffer_release(out_ptr, out_len)

    def mark_recalc_on_load(self) -> None:
        """§5.7.7's mark-only transaction: keep every cached value, set
        ``fullCalcOnLoad="1"``, change nothing else."""
        if not self._handle:
            raise ZlsxError("editor is closed")
        if not _ffi._HAS_MARK_RECALC:
            raise RuntimeError(
                "loaded libzlsx does not expose zlsx_editor_mark_recalc_on_load "
                "(requires 0.8.0+); upgrade libzlsx"
            )
        _report, diag = _fresh_report_and_diag()
        rc = _ffi.lib.zlsx_editor_mark_recalc_on_load(
            self._handle, ctypes.byref(diag), self._err, _ERR_BUF_LEN
        )
        try:
            if rc == _ffi.ZLSX_REFUSED:
                raise _refusal_from_diag(diag)
            if rc != _ffi.ZLSX_OK:
                raise ZlsxError(
                    f"zlsx_editor_mark_recalc_on_load: {_decode_err(self._err)}"
                )
        finally:
            _ffi.lib.zlsx_diag_release(ctypes.byref(diag))

    def recalculate(
        self,
        now=None,
        utc_offset_min: int = 0,
        seed: Optional[int] = None,
        mode: str = "excel",
        profile: str = "windows_1252",
        on_unsupported: str = "refuse",
        timeout: Optional[float] = None,
    ) -> RecalcReport:
        """§5.7's in-memory transaction: recalculate every formula cell
        and swap the result in as the final operation. On refusal
        (:class:`ZlsxFormulaRefusal`), timeout (:class:`TimeoutError`,
        observed pre-commit only) or Ctrl-C the workbook is exactly as
        it was."""
        if not self._handle:
            raise ZlsxError("editor is closed")
        if not _ffi._HAS_RECALC:
            raise RuntimeError(
                "loaded libzlsx does not expose zlsx_editor_recalculate "
                "(requires 0.8.0+); upgrade libzlsx"
            )
        token = _new_cancel_token(self._err)
        try:
            run = _build_run(now, utc_offset_min, seed, mode, profile, "da",
                             on_unsupported, timeout, token)
            report, diag = _fresh_report_and_diag()

            def invoke():
                return _ffi.lib.zlsx_editor_recalculate(
                    self._handle, ctypes.byref(run), ctypes.byref(report),
                    ctypes.byref(diag), self._err, _ERR_BUF_LEN,
                )

            return _drive_recalc("zlsx_editor_recalculate", invoke, token,
                                 timeout, report, diag, self._err)
        finally:
            if token is not None:
                _ffi.lib.zlsx_cancel_token_free(token)

    def save_with_recalc(
        self,
        path: Union[str, Path],
        now=None,
        utc_offset_min: int = 0,
        seed: Optional[int] = None,
        mode: str = "excel",
        profile: str = "windows_1252",
        on_unsupported: str = "refuse",
        timeout: Optional[float] = None,
    ) -> RecalcReport:
        """§5.7.9's atomic file transaction: recalculate, write, rename,
        then swap in memory. Any pre-commit failure — refusal, timeout,
        Ctrl-C, I/O error — leaves the destination's prior bytes (or its
        absence) AND this editor's memory untouched. A cancellation that
        lands post-commit returns normally with
        ``report.cancelled_late=True``. A directory fsync failing after
        the rename is ``report.durability_warning``, never an error."""
        if not self._handle:
            raise ZlsxError("editor is closed")
        if not _ffi._HAS_SAVE_WITH_RECALC:
            raise RuntimeError(
                "loaded libzlsx does not expose zlsx_editor_save_with_recalc "
                "(requires 0.8.0+); upgrade libzlsx"
            )
        pbytes = str(path).encode("utf-8")
        token = _new_cancel_token(self._err)
        try:
            run = _build_run(now, utc_offset_min, seed, mode, profile, "da",
                             on_unsupported, timeout, token)
            report, diag = _fresh_report_and_diag()
            pptr, pkeep = _path_as_ubyte(pbytes)

            def invoke():
                return _ffi.lib.zlsx_editor_save_with_recalc(
                    self._handle, pptr, len(pbytes), ctypes.byref(run),
                    ctypes.byref(report), ctypes.byref(diag),
                    self._err, _ERR_BUF_LEN,
                )

            result = _drive_recalc("zlsx_editor_save_with_recalc", invoke,
                                   token, timeout, report, diag, self._err)
            del pkeep
            return result
        finally:
            if token is not None:
                _ffi.lib.zlsx_cancel_token_free(token)

    def evaluate(
        self,
        formula: str,
        sheet: int = 0,
        anchor: Optional[Tuple[int, int]] = None,
        dialect: str = "da",
        now=None,
        utc_offset_min: int = 0,
        seed: Optional[int] = None,
        mode: str = "excel",
        profile: str = "windows_1252",
        timeout: Optional[float] = None,
    ) -> EvalResult:
        """Standalone cache-based evaluation (M6 semantics): the
        workbook is byte-identical before and after; eval never
        commits. ``anchor`` is ``(row, col)`` with row 1-based and col
        0-based; without one, site-dependent formulas refuse.
        ``.value`` is ``float``/``str``/``bool``, an
        :class:`ExcelError` (a successful *result*, plane 1), or a
        :class:`Matrix`; ``.resolved`` echoes the exact context, so a
        defaulted volatile evaluation is reproducible by replay."""
        if not self._handle:
            raise ZlsxError("editor is closed")
        if not _ffi._HAS_EVAL:
            raise RuntimeError(
                "loaded libzlsx does not expose zlsx_editor_evaluate "
                "(requires 0.8.0+); upgrade libzlsx"
            )
        fbytes = formula.encode("utf-8")
        fptr, fkeep = _path_as_ubyte(fbytes) if fbytes else (None, None)
        anchor_row = 0
        anchor_col = 0
        if anchor is not None:
            anchor_row, anchor_col = int(anchor[0]), int(anchor[1])
            if anchor_row < 1:
                raise ValueError("anchor row is 1-based and must be >= 1")
        token = _new_cancel_token(self._err)
        try:
            run = _build_run(now, utc_offset_min, seed, mode, profile, dialect,
                             "refuse", timeout, token)
            val = _ffi.ValueV1()
            val.struct_size = ctypes.sizeof(_ffi.ValueV1)
            res = _ffi.ResolvedV1()
            res.struct_size = ctypes.sizeof(_ffi.ResolvedV1)
            diag = _ffi.DiagV1()
            diag.struct_size = ctypes.sizeof(_ffi.DiagV1)

            def invoke():
                return _ffi.lib.zlsx_editor_evaluate(
                    self._handle, fptr, len(fbytes), int(sheet),
                    anchor_row, anchor_col, ctypes.byref(run),
                    ctypes.byref(val), ctypes.byref(res), ctypes.byref(diag),
                    self._err, _ERR_BUF_LEN,
                )

            rc, interrupted = _invoke_cancellable(invoke, token)
            try:
                if rc == _ffi.ZLSX_CANCELLED:
                    if interrupted:
                        raise KeyboardInterrupt
                    raise TimeoutError(
                        f"zlsx_editor_evaluate: cancelled (timeout={timeout}s)"
                    )
                if rc == _ffi.ZLSX_REFUSED:
                    raise _refusal_from_diag(diag)
                if rc != _ffi.ZLSX_OK:
                    raise ZlsxError(
                        f"zlsx_editor_evaluate: {_decode_err(self._err)}"
                    )
                payload = (
                    ctypes.string_at(val.payload, val.payload_len)
                    if val.payload_len
                    else b""
                )

                def elem_to_py(e):
                    if e.tag == 0:
                        return float(e.num)
                    if e.tag == 2:
                        return e.num != 0
                    text = payload[e.payload_off : e.payload_off + e.payload_len].decode("utf-8")
                    if e.tag == 3:
                        return ExcelError(text)
                    return text

                if val.is_matrix:
                    cells = [
                        [elem_to_py(val.elems[r * val.cols + c]) for c in range(val.cols)]
                        for r in range(val.rows)
                    ]
                    value = Matrix(rows=val.rows, cols=val.cols, cells=cells)
                else:
                    value = elem_to_py(val.elems[0])
                return EvalResult(value=value, resolved=_decode_resolved(res, anchor))
            finally:
                _ffi.lib.zlsx_value_release(ctypes.byref(val))
                _ffi.lib.zlsx_diag_release(ctypes.byref(diag))
                del fkeep
        finally:
            if token is not None:
                _ffi.lib.zlsx_cancel_token_free(token)

    def close(self) -> None:
        if self._handle:
            _ffi.lib.zlsx_editor_close(self._handle)
            self._handle = None

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()

    def __del__(self):
        # Match Book/Writer: the underlying allocation (full source
        # buffer + entry table + buffered appends) can be sizeable.
        # Best-effort finalize — the C ABI is reentrant-safe and
        # null-tolerant, so a partially-initialized handle (e.g.
        # init raised after _handle was set) still cleans up.
        try:
            self.close()
        except Exception:
            pass


def edit(path: Union[str, Path]) -> Editor:
    """Open an existing xlsx for append-only mutation. See
    :class:`Editor` for the full contract."""
    return Editor(path)


# ─── Embeddings (E5) ─────────────────────────────────────────────────
#
# Semantic vectors stored inside the .xlsx, read back with their
# provenance. The API's shape is dictated by a measured fact: some
# spreadsheet applications rebuild the archive on save and delete the
# vector parts outright. A ~200-byte recovery record survives that in
# most of them, so a workbook which lost its vectors can still say what
# it held.
#
# Hence three states, not two. `vectors()` raising on a stripped
# workbook is deliberate — returning an empty array would recreate
# exactly the silent-nothing this design exists to prevent.
#
# ── The Numbers exception (measured 2026-07-27, Numbers 15.3) ────────
#
# Apple Numbers erases the recovery record too. It strips 5 of the 6
# carriers tested; the only survivor is cell data, which is visible to
# the user and therefore not where the record lives.
#
# So a workbook exported from Numbers reports `absent`, NOT `stripped`.
# It is indistinguishable from one that never had embeddings, and no
# amount of reader-side effort recovers it. This is a property of how
# Numbers exports — it rebuilds the file from its own document model, so
# only what that model represents survives — not something the binding
# can detect or work around.
#
# Practical consequence for callers: `absent` means "no embeddings
# here", not "never had any". If a pipeline needs to distinguish those,
# it has to track expectations outside the workbook.
#
# See docs/plans/emb-4b-carrier-matrix.md for the measurement.


class EmbeddingsStripped(ZlsxError):
    """Vectors were deleted by some tool; provenance survived.

    Raised by :meth:`Embeddings.vectors` and :meth:`Embeddings.hashes`.
    The workbook still knows its model, dimension, dtype and covered
    ranges — read them off the :class:`Embeddings` object and re-embed
    from source.

    .. note::
       This is the *recoverable* loss. Apple Numbers erases the
       recovery record along with the vectors, so a Numbers export
       reports ``absent`` and never raises this. See the module notes.
    """


@dataclass(frozen=True)
class Coverage:
    """One embedded range within a workbook.

    ``rows`` is the vector count, available whether or not the vectors
    themselves survived.
    """

    id: str
    sheet: str
    range: str
    rows: int


class Embeddings:
    """Embedding set of a workbook, in one of three states.

    Use as a context manager::

        with zlsx.embeddings("report.xlsx") as emb:
            if emb.present:
                vecs = emb.vectors("title")          # (rows, dim) float32
                live = vecs[emb.valid_mask("title")]  # drop deleted rows
            elif emb.stripped:
                print("stripped by a tool; was", emb.model, emb.coverages)
            else:
                print("no embeddings")

    ``present`` / ``stripped`` / ``absent`` are mutually exclusive.
    ``model``, ``dim``, ``dtype`` and ``coverages`` are populated for
    both ``present`` and ``stripped`` — recovering them after a strip is
    the entire point of the recovery record.

    .. warning::
       ``absent`` does **not** prove the workbook never had embeddings.
       Apple Numbers strips the recovery record along with the vectors
       (measured on 15.3), so a Numbers export is indistinguishable from
       a workbook that never had any. Pipelines that must tell those
       apart have to track the expectation outside the file.
    """

    __slots__ = ("_h", "_state", "_closed")

    def __init__(self, path: str | os.PathLike[str]) -> None:
        if not _ffi._HAS_EMB:
            raise ZlsxError(
                "libzlsx is too old for the embeddings API "
                "(zlsx_emb_open missing); rebuild or upgrade the library"
            )
        err = ctypes.create_string_buffer(256)
        h = _ffi.lib.zlsx_emb_open(str(path).encode(), err, len(err))
        if not h:
            raise ZlsxError(err.value.decode() or "failed to open workbook")
        self._h = h
        self._closed = False
        self._state = _ffi.lib.zlsx_emb_state(h)

    # ── lifecycle ────────────────────────────────────────────────────

    def close(self) -> None:
        if not self._closed:
            _ffi.lib.zlsx_emb_close(self._h)
            self._closed = True

    def __enter__(self) -> Embeddings:
        return self

    def __exit__(self, *exc: object) -> None:
        self.close()

    def __del__(self) -> None:
        try:
            self.close()
        except Exception:
            pass

    def _check(self) -> None:
        if self._closed:
            raise ZlsxError("Embeddings handle is closed")

    # ── state ────────────────────────────────────────────────────────

    @property
    def present(self) -> bool:
        """Vectors are available."""
        return self._state == _ffi.ZLSX_EMB_PRESENT

    @property
    def stripped(self) -> bool:
        """Vectors were deleted by a tool; provenance recovered."""
        return self._state == _ffi.ZLSX_EMB_STRIPPED

    @property
    def absent(self) -> bool:
        """No embeddings and no recovery record.

        Usually means the workbook never had any — but also what a
        Numbers export looks like, because Numbers erases the record
        too. The two cases are not distinguishable from the file.
        """
        return self._state == _ffi.ZLSX_EMB_ABSENT

    @property
    def state(self) -> str:
        return {
            _ffi.ZLSX_EMB_PRESENT: "present",
            _ffi.ZLSX_EMB_STRIPPED: "stripped",
        }.get(self._state, "absent")

    # ── provenance (present or stripped) ─────────────────────────────

    def _str(self, fn, *args: object) -> str:
        self._check()
        n = fn(self._h, *args, None, 0)
        if n == 0:
            return ""
        buf = ctypes.create_string_buffer(n + 1)
        fn(self._h, *args, buf, len(buf))
        return buf.value.decode()

    @property
    def model(self) -> str:
        return self._str(_ffi.lib.zlsx_emb_model)

    @property
    def dim(self) -> int:
        self._check()
        return int(_ffi.lib.zlsx_emb_dim(self._h))

    @property
    def dtype(self) -> str:
        return self._str(_ffi.lib.zlsx_emb_dtype)

    @property
    def coverages(self) -> list[Coverage]:
        self._check()
        n = _ffi.lib.zlsx_emb_coverage_count(self._h)
        out: list[Coverage] = []
        for i in range(n):
            out.append(
                Coverage(
                    id=self._str(_ffi.lib.zlsx_emb_coverage_id, i),
                    sheet=self._str(_ffi.lib.zlsx_emb_coverage_sheet, i),
                    range=self._str(_ffi.lib.zlsx_emb_coverage_range, i),
                    rows=int(_ffi.lib.zlsx_emb_coverage_rows(self._h, i)),
                )
            )
        return out

    # ── stripped-only ────────────────────────────────────────────────

    @property
    def digest(self) -> int | None:
        """Content fingerprint at embed time, or ``None`` unless stripped.

        Recomputable from the current cells, so an equal digest means the
        covered content has not drifted and a re-embed reproduces the
        same vectors.
        """
        self._check()
        if not self.stripped:
            return None
        return int(_ffi.lib.zlsx_emb_digest(self._h))

    @property
    def carrier(self) -> str | None:
        """Which carrier the record survived in, or ``None`` unless stripped."""
        self._check()
        if not self.stripped:
            return None
        c = _ffi.lib.zlsx_emb_carrier(self._h)
        return {
            _ffi.ZLSX_EMB_CARRIER_DOC_PROPS: "doc_props",
            _ffi.ZLSX_EMB_CARRIER_CELL_DATA: "cell_data",
        }.get(c, "defined_name")

    # ── vectors (present only) ───────────────────────────────────────

    def _coverage_index(self, coverage: str | int) -> tuple[int, int]:
        covs = self.coverages
        if isinstance(coverage, int):
            if not 0 <= coverage < len(covs):
                raise IndexError(f"coverage index {coverage} out of range")
            return coverage, covs[coverage].rows
        for i, c in enumerate(covs):
            if c.id == coverage:
                return i, c.rows
        raise KeyError(f"no coverage named {coverage!r}")

    def _require_present(self) -> None:
        if self.stripped:
            raise EmbeddingsStripped(
                f"vectors were stripped by another tool; the workbook recorded "
                f"model={self.model!r} dim={self.dim} dtype={self.dtype!r} "
                f"over {len(self.coverages)} coverage(s). Re-embed from source."
            )
        if self.absent:
            raise ZlsxError("this workbook has no embeddings")

    def vectors(self, coverage: str | int = 0):
        """Vectors for ``coverage`` as a ``(rows, dim)`` float32 array.

        Requires NumPy. Raises :class:`EmbeddingsStripped` when the
        vectors were deleted — the provenance is still readable off this
        object, and an empty array would hide that.

        Decoding happens in Zig: one FFI call per coverage rather than
        per row, and each dtype's layout has exactly one implementation.
        """
        import numpy as np

        self._check()
        self._require_present()
        i, rows = self._coverage_index(coverage)
        dim = self.dim
        out = np.empty(rows * dim, dtype=np.float32)
        rc = _ffi.lib.zlsx_emb_vectors(
            self._h,
            i,
            out.ctypes.data_as(ctypes.POINTER(ctypes.c_float)),
            out.size,
        )
        if rc != 0:
            raise ZlsxError(f"zlsx_emb_vectors failed (rc={rc})")
        return out.reshape(rows, dim)

    def hashes(self, coverage: str | int = 0):
        """Per-row content hashes for ``coverage`` as a uint64 array."""
        import numpy as np

        self._check()
        self._require_present()
        i, rows = self._coverage_index(coverage)
        out = np.empty(rows, dtype=np.uint64)
        rc = _ffi.lib.zlsx_emb_hashes(
            self._h,
            i,
            out.ctypes.data_as(ctypes.POINTER(ctypes.c_uint64)),
            out.size,
        )
        if rc != 0:
            raise ZlsxError(f"zlsx_emb_hashes failed (rc={rc})")
        return out

    def valid_mask(self, coverage: str | int = 0):
        """Boolean mask, False where the row was deleted (tombstoned).

        A deleted row keeps its slot in the vector array so indices stay
        aligned with the covered range; its hash is the tombstone
        sentinel. Mask before using the vectors::

            v = emb.vectors("title")[emb.valid_mask("title")]
        """
        import numpy as np

        tomb = np.uint64(_ffi.lib.zlsx_emb_tombstone())
        return self.hashes(coverage) != tomb

    def __repr__(self) -> str:
        if self._closed:
            return "<Embeddings closed>"
        if self.absent:
            return "<Embeddings absent>"
        return (
            f"<Embeddings {self.state} model={self.model!r} dim={self.dim} "
            f"dtype={self.dtype!r} coverages={len(self.coverages)}>"
        )


def embeddings(path: str | os.PathLike[str]) -> Embeddings:
    """Open a workbook's embedding set. See :class:`Embeddings`."""
    return Embeddings(path)
