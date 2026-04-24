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
from pathlib import Path
from typing import Iterator, Union

from . import _ffi

__version__ = "0.2.4"
"""Python-package version. Tracks the Zig library's major+minor; the patch
level may drift when the binding ships a Python-only fix."""

__all__ = [
    "open",
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
    "Comment",
    "Dxf",
    "CF_OPERATORS",
    "to_excel_serial",
    "read",
    "ZlsxError",
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
    """An external-URL hyperlink attached to a cell or cell range.
    ``url`` is the raw ``Target`` attribute from the sheet's rels
    file — XML entities like ``&amp;`` are preserved, so the URL
    round-trips byte-for-byte through save/reopen. Decode at the
    caller if a display form is needed."""
    top_left: CellRef
    bottom_right: CellRef
    url: str


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
    through the ``<authors>`` table; ``text`` is the concatenated
    plain-text body (rich runs inside a comment are flattened — the
    formatted form lands in a follow-up without breaking this
    shape). All strings are entity-decoded."""
    top_left: CellRef
    author: str
    text: str


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


# ─── Book ─────────────────────────────────────────────────────────────


class Book:
    """A workbook handle. Use :func:`zlsx.open` to construct one.

    Also usable as a context manager; exit closes the handle::

        with zlsx.open("file.xlsx") as book:
            ...
    """

    def __init__(self, path: Union[str, Path]):
        self._err = ctypes.create_string_buffer(_ERR_BUF_LEN)
        path_bytes = str(path).encode("utf-8")
        self._handle = _ffi.lib.zlsx_book_open(path_bytes, self._err, _ERR_BUF_LEN)
        if not self._handle:
            raise ZlsxError(f"zlsx_book_open({path!r}): {_decode_err(self._err)}")

        # Cache sheet names at open time — most callers enumerate them,
        # and the list is short (<10 in typical workbooks).
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
        """External-URL hyperlinks declared on sheet ``sheet_idx``,
        resolved through the sheet's ``_rels/sheet{N}.xml.rels`` file.
        Returns an empty list for sheets without a ``<hyperlinks>``
        block."""
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
        for i in range(count):
            rc = _ffi.lib.zlsx_hyperlink_at(self._handle, sheet_idx, i, ctypes.byref(hl))
            if rc != 0:
                continue
            url = ctypes.string_at(hl.url_ptr, hl.url_len).decode("utf-8", errors="replace")
            out.append(Hyperlink(
                top_left=CellRef(col=hl.top_left_col, row=hl.top_left_row),
                bottom_right=CellRef(col=hl.bottom_right_col, row=hl.bottom_right_row),
                url=url,
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
            out.append(Comment(
                top_left=CellRef(col=cc.cell_col, row=cc.cell_row),
                author=author,
                text=text,
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
        plain Python lists, so any tabular library can consume it."""
        with self.rows() as r:
            all_rows = list(r)
        if not header or len(all_rows) == 0:
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
SheetWriter.freeze_panes = _sheet_freeze_panes           # type: ignore[attr-defined]
SheetWriter.set_auto_filter = _sheet_set_auto_filter     # type: ignore[attr-defined]
SheetWriter.add_merged_cell = _sheet_add_merged_cell     # type: ignore[attr-defined]
SheetWriter.add_hyperlink = _sheet_add_hyperlink         # type: ignore[attr-defined]
SheetWriter.add_comment = _sheet_add_comment             # type: ignore[attr-defined]
SheetWriter.add_data_validation_list = _sheet_add_data_validation_list  # type: ignore[attr-defined]
SheetWriter.add_data_validation_numeric = _sheet_add_data_validation_numeric  # type: ignore[attr-defined]
SheetWriter.add_data_validation_custom = _sheet_add_data_validation_custom  # type: ignore[attr-defined]
SheetWriter.add_conditional_format_cell_is = _sheet_add_conditional_format_cell_is  # type: ignore[attr-defined]
SheetWriter.add_conditional_format_expression = _sheet_add_conditional_format_expression  # type: ignore[attr-defined]


class Writer:
    """A xlsx workbook under construction.

    Use :func:`zlsx.write` to construct one. Finalise by calling
    :meth:`save` with a target path, then :meth:`close` to release
    resources. The context-manager protocol wraps this: ``with
    zlsx.write("out.xlsx") as w:`` saves automatically on clean exit.

    Writes strings, integers, floats, booleans, and empties; styles
    via :meth:`add_style` (bold/italic, fonts, fills, borders,
    alignment, wrap, number formats); per-sheet ``set_column_width``,
    ``freeze_panes``, ``set_auto_filter``, ``add_merged_cell``,
    ``add_hyperlink`` (external URLs). Formulas and load-modify-save
    round-trip remain out of scope until Phase 3c.
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

    def save(self, path: Union[str, Path, None] = None) -> None:
        """Write the workbook to disk. Uses the path passed to
        :func:`zlsx.write` if none is provided here."""
        target = Path(path) if path is not None else self._path
        if target is None:
            raise ValueError("no save path: pass one to zlsx.write() or Writer.save()")
        raw = str(target).encode("utf-8")
        rc = _ffi.lib.zlsx_writer_save(
            self._handle, raw, len(raw), self._err, _ERR_BUF_LEN
        )
        if rc != 0:
            raise ZlsxError(f"zlsx_writer_save({target!r}): {_decode_err(self._err)}")

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
