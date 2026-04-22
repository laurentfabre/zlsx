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

__version__ = "0.2.3"
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
    "ZlsxError",
]


# ─── Styles (Phase 3b) ─────────────────────────────────────────────────
#
# `Style` is a dataclass that mirrors the Zig writer's Style struct. Keep
# fields additive — openpyxl-parity fields land in subsequent releases
# (alignment, fills, borders, number formats, etc).

from dataclasses import dataclass
from typing import Literal, Optional


HAlignLiteral = Literal[
    "general", "left", "center", "right", "fill",
    "justify", "centerContinuous", "distributed",
]

_HALIGN_VALUES = {
    "general": 0, "left": 1, "center": 2, "right": 3,
    "fill": 4, "justify": 5, "centerContinuous": 6, "distributed": 7,
}


@dataclass(frozen=True)
class Style:
    """A cell-style specification. Pass an instance to
    :meth:`Writer.add_style` and use the returned index with
    :meth:`SheetWriter.write_row`'s ``styles`` argument.

    Fields mirror the Zig Style struct. ``None`` means "unset (OOXML
    default)"; concrete values emit the corresponding XML attributes.
    ``font_color_argb`` is packed ARGB (0xAARRGGBB); for fully opaque
    red use ``0xFFFF0000``.
    """

    font_bold: bool = False
    font_italic: bool = False
    font_size: Optional[float] = None
    font_name: Optional[str] = None
    font_color_argb: Optional[int] = None
    alignment_horizontal: HAlignLiteral = "general"
    wrap_text: bool = False


class ZlsxError(RuntimeError):
    """Raised when the zlsx C ABI returns an error. ``args[0]`` is the
    null-terminated diagnostic written into the error buffer by the
    library."""


_ERR_BUF_LEN = 256


def _decode_err(buf: ctypes.Array) -> str:
    return bytes(buf.value).decode("utf-8", errors="replace")


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


class Writer:
    """A xlsx workbook under construction.

    Use :func:`zlsx.write` to construct one. Finalise by calling
    :meth:`save` with a target path, then :meth:`close` to release
    resources. The context-manager protocol wraps this: ``with
    zlsx.write("out.xlsx") as w:`` saves automatically on clean exit.

    Scope reminder: this MVP writer handles strings, integers, floats,
    booleans, and empties. Styles, merged regions, formulas, and
    load-modify-save round-trip are not yet supported — use openpyxl
    for those until Phase 3b/3c ship.
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

        needs_ex = (
            style.font_size is not None
            or style.font_name is not None
            or style.font_color_argb is not None
            or style.alignment_horizontal != "general"
            or style.wrap_text
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

        # Distinguish "unset" (None) from "empty string" — the latter
        # is invalid and must reach the Zig side as font_name_len=0
        # with an explicit sentinel that triggers InvalidFontName.
        if style.font_name is None:
            name_bytes = b""
        elif style.font_name == "":
            # Force Zig to see font_name = "" (len=0 with a non-null ptr
            # via a different sentinel than the "unset" case).
            raise ZlsxError("InvalidFontName")
        else:
            name_bytes = style.font_name.encode("utf-8")
        # Keep the bytes buffer alive through the FFI call.
        name_buf = (ctypes.c_ubyte * max(len(name_bytes), 1)).from_buffer_copy(
            name_bytes or b"\x00"
        )

        if style.alignment_horizontal not in _HALIGN_VALUES:
            raise ValueError(
                f"unknown alignment_horizontal: {style.alignment_horizontal!r}"
            )

        spec = _ffi.CStyle(
            font_bold=1 if style.font_bold else 0,
            font_italic=1 if style.font_italic else 0,
            alignment_horizontal=_HALIGN_VALUES[style.alignment_horizontal],
            wrap_text=1 if style.wrap_text else 0,
            flags=flags,
            font_size=float(style.font_size or 0.0),
            font_color_argb=int(style.font_color_argb or 0),
            font_name_ptr=ctypes.cast(name_buf, ctypes.POINTER(ctypes.c_ubyte)),
            font_name_len=len(name_bytes),
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
