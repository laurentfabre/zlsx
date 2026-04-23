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


class ZlsxError(RuntimeError):
    """Raised when the zlsx C ABI returns an error. ``args[0]`` is the
    null-terminated diagnostic written into the error buffer by the
    library."""


_ERR_BUF_LEN = 256


def _decode_err(buf: ctypes.Array) -> str:
    return bytes(buf.value).decode("utf-8", errors="replace")


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

        row = [_cell_to_py(cells_ptr[i]) for i in range(cells_len.value)]
        return row

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


def open(path: Union[str, Path]) -> Book:  # noqa: A001  (shadows builtin by design)
    """Open an ``.xlsx`` file for reading.

    Returns a :class:`Book` handle. The file must exist and be a valid
    xlsx archive. Raises :class:`ZlsxError` on parse failure.
    """
    return Book(path)


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


SheetWriter.set_column_width = _sheet_set_column_width   # type: ignore[attr-defined]
SheetWriter.freeze_panes = _sheet_freeze_panes           # type: ignore[attr-defined]
SheetWriter.set_auto_filter = _sheet_set_auto_filter     # type: ignore[attr-defined]
SheetWriter.add_merged_cell = _sheet_add_merged_cell     # type: ignore[attr-defined]
SheetWriter.add_hyperlink = _sheet_add_hyperlink         # type: ignore[attr-defined]


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
