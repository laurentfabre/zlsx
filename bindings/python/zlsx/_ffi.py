"""Internal ctypes bindings over libzlsx. Not part of the public API."""

from __future__ import annotations

import ctypes
import ctypes.util
import os
import sys
from pathlib import Path

# ─── Locate libzlsx ───────────────────────────────────────────────────


def _candidates() -> list[Path]:
    out: list[Path] = []
    here = Path(__file__).parent

    # 1. Explicit override — ZLSX_LIBRARY=/path/to/libzlsx.dylib
    if env := os.environ.get("ZLSX_LIBRARY"):
        out.append(Path(env))

    # 2. Bundled inside the wheel (same directory as this file). Populated
    #    by cibuildwheel in CI; absent in source-install mode.
    for name in ("libzlsx.dylib", "libzlsx.so", "zlsx.dll"):
        out.append(here / name)

    # 3. Local dev build at <repo>/zig-out/lib/. Placed BEFORE Homebrew
    #    so when working on the Zig side with `pip install -e .`, the
    #    freshly-built dylib shadows whatever brew installed. This makes
    #    "edit, zig build, run pytest" loops work without needing
    #    ZLSX_LIBRARY to be set manually.
    for rel in ("../../zig-out/lib", "../../../zig-out/lib"):
        out.append(here / rel / "libzlsx.dylib")
        out.append(here / rel / "libzlsx.so")

    # 4. Homebrew install path (fallback for end users).
    if sys.platform == "darwin":
        for prefix in ("/opt/homebrew/opt/zlsx", "/usr/local/opt/zlsx"):
            out.append(Path(prefix) / "lib" / "libzlsx.dylib")
    elif sys.platform.startswith("linux"):
        out.append(Path("/home/linuxbrew/.linuxbrew/opt/zlsx/lib/libzlsx.so"))
        out.append(Path("/usr/local/lib/libzlsx.so"))

    return out


def _load_library() -> ctypes.CDLL:
    tried: list[str] = []
    for cand in _candidates():
        cand = cand.resolve(strict=False)
        if cand.is_file():
            return ctypes.CDLL(str(cand))
        tried.append(str(cand))

    # Last-chance: system resolver.
    found = ctypes.util.find_library("zlsx")
    if found:
        return ctypes.CDLL(found)

    raise ImportError(
        "libzlsx not found. Install it via `brew install laurentfabre/zlsx/zlsx` "
        "or download a release tarball from "
        "https://github.com/laurentfabre/zlsx/releases and point ZLSX_LIBRARY "
        "at the .dylib / .so. Tried:\n  " + "\n  ".join(tried)
    )


lib = _load_library()

# ─── Types ─────────────────────────────────────────────────────────────


class Cell(ctypes.Structure):
    """Mirrors zlsx_cell_t in include/zlsx.h — flat struct, interpret via tag."""

    _fields_ = [
        ("tag", ctypes.c_uint32),
        ("str_len", ctypes.c_uint32),
        ("str_ptr", ctypes.POINTER(ctypes.c_ubyte)),
        ("i", ctypes.c_int64),
        ("f", ctypes.c_double),
        ("b", ctypes.c_uint8),
        ("_pad", ctypes.c_ubyte * 7),
    ]


# Cell tag constants (matches the C enum in zlsx.h).
CELL_EMPTY = 0
CELL_STRING = 1
CELL_INTEGER = 2
CELL_NUMBER = 3
CELL_BOOLEAN = 4

cell_ptr = ctypes.POINTER(Cell)
book_handle = ctypes.c_void_p
rows_handle = ctypes.c_void_p
writer_handle = ctypes.c_void_p
sheet_writer_handle = ctypes.c_void_p

# ─── Function signatures ──────────────────────────────────────────────

lib.zlsx_abi_version.argtypes = []
lib.zlsx_abi_version.restype = ctypes.c_uint32

lib.zlsx_version_string.argtypes = []
lib.zlsx_version_string.restype = ctypes.c_char_p

lib.zlsx_book_open.argtypes = [
    ctypes.c_char_p,  # path (null-terminated)
    ctypes.c_char_p,  # err_buf
    ctypes.c_size_t,  # err_buf_len
]
lib.zlsx_book_open.restype = book_handle

lib.zlsx_book_close.argtypes = [book_handle]
lib.zlsx_book_close.restype = None

lib.zlsx_sheet_count.argtypes = [book_handle]
lib.zlsx_sheet_count.restype = ctypes.c_uint32

lib.zlsx_sheet_name.argtypes = [
    book_handle,
    ctypes.c_uint32,
    ctypes.c_char_p,
    ctypes.c_size_t,
]
lib.zlsx_sheet_name.restype = ctypes.c_size_t

lib.zlsx_sheet_index_by_name.argtypes = [
    book_handle,
    ctypes.c_char_p,
    ctypes.c_size_t,
]
lib.zlsx_sheet_index_by_name.restype = ctypes.c_int32

lib.zlsx_rows_open.argtypes = [
    book_handle,
    ctypes.c_uint32,
    ctypes.c_char_p,
    ctypes.c_size_t,
]
lib.zlsx_rows_open.restype = rows_handle

lib.zlsx_rows_close.argtypes = [rows_handle]
lib.zlsx_rows_close.restype = None

lib.zlsx_rows_next.argtypes = [
    rows_handle,
    ctypes.POINTER(cell_ptr),
    ctypes.POINTER(ctypes.c_size_t),
    ctypes.c_char_p,
    ctypes.c_size_t,
]
lib.zlsx_rows_next.restype = ctypes.c_int32

# ─── Writer exports (v0.2.2+) ─────────────────────────────────────────

lib.zlsx_writer_create.argtypes = [ctypes.c_char_p, ctypes.c_size_t]
lib.zlsx_writer_create.restype = writer_handle

lib.zlsx_writer_close.argtypes = [writer_handle]
lib.zlsx_writer_close.restype = None

lib.zlsx_writer_add_sheet.argtypes = [
    writer_handle,
    ctypes.c_char_p,
    ctypes.c_size_t,
    ctypes.c_char_p,
    ctypes.c_size_t,
]
lib.zlsx_writer_add_sheet.restype = sheet_writer_handle

lib.zlsx_sheet_writer_write_row.argtypes = [
    sheet_writer_handle,
    cell_ptr,
    ctypes.c_size_t,
    ctypes.c_char_p,
    ctypes.c_size_t,
]
lib.zlsx_sheet_writer_write_row.restype = ctypes.c_int32

lib.zlsx_writer_save.argtypes = [
    writer_handle,
    ctypes.c_char_p,
    ctypes.c_size_t,
    ctypes.c_char_p,
    ctypes.c_size_t,
]
lib.zlsx_writer_save.restype = ctypes.c_int32

# ─── Styles (Phase 3b, available in libzlsx 0.2.4+) ───────────────────
#
# The `_ex` convention documented in the header leaves us with a single
# addStyle signature per ABI revision — we consume it here with a
# hasattr() guard so py-zlsx keeps importing against older dylibs.
# Callers that try to use styles against an older library get a clear
# AttributeError via the public Writer.add_style() wrapper.

_HAS_STYLES = hasattr(lib, "zlsx_writer_add_style")

if _HAS_STYLES:
    lib.zlsx_writer_add_style.argtypes = [
        writer_handle,
        ctypes.c_uint8,
        ctypes.c_uint8,
        ctypes.POINTER(ctypes.c_uint32),
        ctypes.c_char_p,
        ctypes.c_size_t,
    ]
    lib.zlsx_writer_add_style.restype = ctypes.c_int32

    lib.zlsx_sheet_writer_write_row_styled.argtypes = [
        sheet_writer_handle,
        cell_ptr,
        ctypes.POINTER(ctypes.c_uint32),
        ctypes.c_size_t,
        ctypes.c_char_p,
        ctypes.c_size_t,
    ]
    lib.zlsx_sheet_writer_write_row_styled.restype = ctypes.c_int32


# ─── Stage-2 style extension (libzlsx 0.2.4+) ──────────────────────────


class CStyle(ctypes.Structure):
    """Mirrors zlsx_style_t in include/zlsx.h."""
    _fields_ = [
        ("font_bold", ctypes.c_uint8),
        ("font_italic", ctypes.c_uint8),
        ("alignment_horizontal", ctypes.c_uint8),
        ("wrap_text", ctypes.c_uint8),
        ("flags", ctypes.c_uint8),
        ("_pad0", ctypes.c_ubyte * 3),
        ("font_size", ctypes.c_float),
        ("font_color_argb", ctypes.c_uint32),
        ("font_name_ptr", ctypes.POINTER(ctypes.c_ubyte)),
        ("font_name_len", ctypes.c_size_t),
    ]


FONT_SIZE_SET = 0x01
FONT_COLOR_SET = 0x02

_HAS_STYLES_EX = hasattr(lib, "zlsx_writer_add_style_ex")
if _HAS_STYLES_EX:
    lib.zlsx_writer_add_style_ex.argtypes = [
        writer_handle,
        ctypes.POINTER(CStyle),
        ctypes.POINTER(ctypes.c_uint32),
        ctypes.c_char_p,
        ctypes.c_size_t,
    ]
    lib.zlsx_writer_add_style_ex.restype = ctypes.c_int32

# ─── ABI version check ────────────────────────────────────────────────

EXPECTED_ABI_VERSION = 1
_found_abi = lib.zlsx_abi_version()
if _found_abi != EXPECTED_ABI_VERSION:
    raise ImportError(
        f"libzlsx ABI mismatch: py-zlsx expects v{EXPECTED_ABI_VERSION}, "
        f"loaded library reports v{_found_abi}. Upgrade one of them."
    )
