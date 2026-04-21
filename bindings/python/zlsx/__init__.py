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

__version__ = "0.2.1"
"""Python-package version. Tracks the Zig library's major+minor; the patch
level may drift when the binding ships a Python-only fix."""

__all__ = ["open", "Book", "Sheet", "Rows", "ZlsxError"]


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
