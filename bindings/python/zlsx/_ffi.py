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
matrix_handle = ctypes.c_void_p
writer_handle = ctypes.c_void_p
sheet_writer_handle = ctypes.c_void_p
editor_handle = ctypes.c_void_p

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

lib.zlsx_book_open_buffer.argtypes = [
    ctypes.c_char_p,  # data (raw bytes; length passed separately, no NUL needed)
    ctypes.c_size_t,  # len
    ctypes.c_char_p,  # err_buf
    ctypes.c_size_t,  # err_buf_len
]
lib.zlsx_book_open_buffer.restype = book_handle

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

# S3b slice 10 (0.9.0+): sheet visibility on the reader handle — the
# `<sheet state="…">` code of one sheet (0 visible / 1 hidden /
# 2 veryHidden, -1 out of range). Probe + skip so older dylibs keep
# importing; `Book.sheet_state` raises RuntimeError without it.
_HAS_SHEET_STATE = hasattr(lib, "zlsx_sheet_state")
if _HAS_SHEET_STATE:
    lib.zlsx_sheet_state.argtypes = [book_handle, ctypes.c_uint32]
    lib.zlsx_sheet_state.restype = ctypes.c_int32

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

# ─── Row skip (available in libzlsx 0.8.0+) ──────────────────────────
#
# Guarded like the other post-0.2.4 additions so py-zlsx keeps importing
# against an older dylib; Rows.skip() falls back to draining next() when
# the symbol is absent, which is what it replaces.
_HAS_ROWS_SKIP = hasattr(lib, "zlsx_rows_skip")

if _HAS_ROWS_SKIP:
    lib.zlsx_rows_skip.argtypes = [
        rows_handle,
        ctypes.c_size_t,                     # n
        ctypes.POINTER(ctypes.c_size_t),     # out_skipped
        ctypes.c_char_p,                     # err_buf
        ctypes.c_size_t,                     # err_buf_len
    ]
    lib.zlsx_rows_skip.restype = ctypes.c_int32

# ─── Matrix exports (v0.2.8+, bulk-FFI for sheet-at-a-time reads) ─────

_HAS_MATRIX = hasattr(lib, "zlsx_matrix_open")
if _HAS_MATRIX:
    lib.zlsx_matrix_open.argtypes = [
        book_handle,
        ctypes.c_uint32,
        ctypes.c_char_p,
        ctypes.c_size_t,
    ]
    lib.zlsx_matrix_open.restype = matrix_handle

    lib.zlsx_matrix_close.argtypes = [matrix_handle]
    lib.zlsx_matrix_close.restype = None

    lib.zlsx_matrix_data.argtypes = [
        matrix_handle,
        ctypes.POINTER(cell_ptr),
        ctypes.POINTER(ctypes.POINTER(ctypes.c_size_t)),
        ctypes.POINTER(ctypes.c_size_t),
    ]
    lib.zlsx_matrix_data.restype = None

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


class CRichRun(ctypes.Structure):
    _fields_ = [
        ("text_ptr", ctypes.POINTER(ctypes.c_ubyte)),
        ("text_len", ctypes.c_size_t),
        ("bold", ctypes.c_uint8),
        ("italic", ctypes.c_uint8),
        ("has_color", ctypes.c_uint8),
        ("has_size", ctypes.c_uint8),
        ("color_argb", ctypes.c_uint32),
        ("size", ctypes.c_float),
        ("font_name_ptr", ctypes.POINTER(ctypes.c_ubyte)),
        ("font_name_len", ctypes.c_size_t),
    ]


_HAS_WRITE_RICH_ROW = hasattr(lib, "zlsx_sheet_writer_write_rich_row")
if _HAS_WRITE_RICH_ROW:
    lib.zlsx_sheet_writer_write_rich_row.argtypes = [
        sheet_writer_handle,
        cell_ptr,                                              # cells_ptr
        ctypes.POINTER(ctypes.POINTER(CRichRun)),              # rich_runs_ptrs
        ctypes.POINTER(ctypes.c_size_t),                       # rich_runs_lens
        ctypes.c_size_t,                                       # cells_len
        ctypes.c_char_p,
        ctypes.c_size_t,
    ]
    lib.zlsx_sheet_writer_write_rich_row.restype = ctypes.c_int32

# Formula authoring (libzlsx 0.2.7+). Probed independently — older
# dylibs that ship the rest of the writer surface but not this symbol
# fall back through `_HAS_WRITE_ROW_WITH_FORMULAS = False`.
_HAS_WRITE_ROW_WITH_FORMULAS = hasattr(lib, "zlsx_sheet_writer_write_row_with_formulas")
if _HAS_WRITE_ROW_WITH_FORMULAS:
    lib.zlsx_sheet_writer_write_row_with_formulas.argtypes = [
        sheet_writer_handle,
        cell_ptr,                                              # cells_ptr
        ctypes.POINTER(ctypes.POINTER(ctypes.c_ubyte)),        # formula_ptrs
        ctypes.POINTER(ctypes.c_size_t),                       # formula_lens
        ctypes.c_size_t,                                       # cells_len
        ctypes.c_char_p,
        ctypes.c_size_t,
    ]
    lib.zlsx_sheet_writer_write_row_with_formulas.restype = ctypes.c_int32

lib.zlsx_writer_save.argtypes = [
    writer_handle,
    ctypes.c_char_p,
    ctypes.c_size_t,
    ctypes.c_char_p,
    ctypes.c_size_t,
]
lib.zlsx_writer_save.restype = ctypes.c_int32

# ─── Writer to buffer (available in libzlsx 0.8.0+) ───────────────────
#
# The writer-side mirror of zlsx_book_open_buffer. Guarded like the
# styles block so py-zlsx keeps importing against an older dylib; the
# public Writer.to_bytes() raises a clear error when the symbol is
# absent. restype is POINTER(c_ubyte) rather than c_char_p so ctypes
# hands back the raw address instead of eagerly copying to bytes — the
# copy is ours to make, once, at a known length.
_HAS_SAVE_TO_BUFFER = hasattr(lib, "zlsx_writer_save_to_buffer")

if _HAS_SAVE_TO_BUFFER:
    lib.zlsx_writer_save_to_buffer.argtypes = [
        writer_handle,
        ctypes.POINTER(ctypes.POINTER(ctypes.c_ubyte)),  # out_ptr
        ctypes.POINTER(ctypes.c_size_t),                 # out_len
        ctypes.c_char_p,                                 # err_buf
        ctypes.c_size_t,                                 # err_buf_len
    ]
    lib.zlsx_writer_save_to_buffer.restype = ctypes.c_int32

    lib.zlsx_buffer_free.argtypes = [
        ctypes.POINTER(ctypes.c_ubyte),
        ctypes.c_size_t,
    ]
    lib.zlsx_buffer_free.restype = None

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
        ("fill_pattern", ctypes.c_uint8),
        ("flags2", ctypes.c_uint8),
        ("_pad0", ctypes.c_ubyte * 1),
        ("font_size", ctypes.c_float),
        ("font_color_argb", ctypes.c_uint32),
        ("fill_fg_argb", ctypes.c_uint32),
        ("fill_bg_argb", ctypes.c_uint32),
        ("border_left_style", ctypes.c_uint8),
        ("border_right_style", ctypes.c_uint8),
        ("border_top_style", ctypes.c_uint8),
        ("border_bottom_style", ctypes.c_uint8),
        ("border_diagonal_style", ctypes.c_uint8),
        ("diagonal_up", ctypes.c_uint8),
        ("diagonal_down", ctypes.c_uint8),
        ("_pad1", ctypes.c_ubyte * 1),
        ("border_left_color_argb", ctypes.c_uint32),
        ("border_right_color_argb", ctypes.c_uint32),
        ("border_top_color_argb", ctypes.c_uint32),
        ("border_bottom_color_argb", ctypes.c_uint32),
        ("border_diagonal_color_argb", ctypes.c_uint32),
        ("font_name_ptr", ctypes.POINTER(ctypes.c_ubyte)),
        ("font_name_len", ctypes.c_size_t),
        ("num_fmt_ptr", ctypes.POINTER(ctypes.c_ubyte)),
        ("num_fmt_len", ctypes.c_size_t),
    ]


FONT_SIZE_SET = 0x01
FONT_COLOR_SET = 0x02
FILL_FG_SET = 0x04
FILL_BG_SET = 0x08

# flags2 bits
BORDER_LEFT_COLOR_SET = 0x01
BORDER_RIGHT_COLOR_SET = 0x02
BORDER_TOP_COLOR_SET = 0x04
BORDER_BOTTOM_COLOR_SET = 0x08
BORDER_DIAGONAL_COLOR_SET = 0x10


# Stage-5 per-sheet functions (libzlsx 0.2.4+).
_HAS_SHEET_FEATURES = hasattr(lib, "zlsx_sheet_writer_set_column_width")

if _HAS_SHEET_FEATURES:
    lib.zlsx_sheet_writer_set_column_width.argtypes = [
        sheet_writer_handle,
        ctypes.c_uint32,
        ctypes.c_float,
        ctypes.c_char_p,
        ctypes.c_size_t,
    ]
    lib.zlsx_sheet_writer_set_column_width.restype = ctypes.c_int32

    lib.zlsx_sheet_writer_freeze_panes.argtypes = [
        sheet_writer_handle,
        ctypes.c_uint32,
        ctypes.c_uint32,
    ]
    lib.zlsx_sheet_writer_freeze_panes.restype = None

    lib.zlsx_sheet_writer_set_auto_filter.argtypes = [
        sheet_writer_handle,
        ctypes.POINTER(ctypes.c_ubyte),
        ctypes.c_size_t,
        ctypes.c_char_p,
        ctypes.c_size_t,
    ]
    lib.zlsx_sheet_writer_set_auto_filter.restype = ctypes.c_int32

# Merged-cell authoring (libzlsx 0.2.5+ — independent of _HAS_SHEET_FEATURES
# because we want py-zlsx to keep importing against a 0.2.4 dylib and only
# fail when the caller actually requests the feature).
_HAS_MERGED_CELL = hasattr(lib, "zlsx_sheet_writer_add_merged_cell")
if _HAS_MERGED_CELL:
    lib.zlsx_sheet_writer_add_merged_cell.argtypes = [
        sheet_writer_handle,
        ctypes.POINTER(ctypes.c_ubyte),
        ctypes.c_size_t,
        ctypes.c_char_p,
        ctypes.c_size_t,
    ]
    lib.zlsx_sheet_writer_add_merged_cell.restype = ctypes.c_int32

# Data-validation (list / dropdown) — same feature-probe pattern.
_HAS_DATA_VALIDATION = hasattr(lib, "zlsx_sheet_writer_add_data_validation_list")
if _HAS_DATA_VALIDATION:
    lib.zlsx_sheet_writer_add_data_validation_list.argtypes = [
        sheet_writer_handle,
        ctypes.POINTER(ctypes.c_ubyte),
        ctypes.c_size_t,
        ctypes.POINTER(ctypes.POINTER(ctypes.c_ubyte)),
        ctypes.POINTER(ctypes.c_size_t),
        ctypes.c_size_t,
        ctypes.c_char_p,
        ctypes.c_size_t,
    ]
    lib.zlsx_sheet_writer_add_data_validation_list.restype = ctypes.c_int32

# Extended writer DV (numeric / custom) added in 0.2.6+.
_HAS_DATA_VALIDATION_EXT = (
    hasattr(lib, "zlsx_sheet_writer_add_data_validation_numeric")
    and hasattr(lib, "zlsx_sheet_writer_add_data_validation_custom")
)
if _HAS_DATA_VALIDATION_EXT:
    lib.zlsx_sheet_writer_add_data_validation_numeric.argtypes = [
        sheet_writer_handle,
        ctypes.POINTER(ctypes.c_ubyte),
        ctypes.c_size_t,
        ctypes.c_uint32,
        ctypes.c_uint32,
        ctypes.POINTER(ctypes.c_ubyte),
        ctypes.c_size_t,
        ctypes.POINTER(ctypes.c_ubyte),
        ctypes.c_size_t,
        ctypes.c_char_p,
        ctypes.c_size_t,
    ]
    lib.zlsx_sheet_writer_add_data_validation_numeric.restype = ctypes.c_int32
    lib.zlsx_sheet_writer_add_data_validation_custom.argtypes = [
        sheet_writer_handle,
        ctypes.POINTER(ctypes.c_ubyte),
        ctypes.c_size_t,
        ctypes.POINTER(ctypes.c_ubyte),
        ctypes.c_size_t,
        ctypes.c_char_p,
        ctypes.c_size_t,
    ]
    lib.zlsx_sheet_writer_add_data_validation_custom.restype = ctypes.c_int32

# Hyperlink authoring — same import-time feature-probe pattern.
_HAS_HYPERLINK = hasattr(lib, "zlsx_sheet_writer_add_hyperlink")
if _HAS_HYPERLINK:
    lib.zlsx_sheet_writer_add_hyperlink.argtypes = [
        sheet_writer_handle,
        ctypes.POINTER(ctypes.c_ubyte),
        ctypes.c_size_t,
        ctypes.POINTER(ctypes.c_ubyte),
        ctypes.c_size_t,
        ctypes.c_char_p,
        ctypes.c_size_t,
    ]
    lib.zlsx_sheet_writer_add_hyperlink.restype = ctypes.c_int32

# Internal-hyperlink authoring (libzlsx 0.2.7+). Same shape as
# `add_hyperlink` but the second string is the workbook-internal
# `location` (e.g. "Sheet2!A1") instead of an external URL.
_HAS_INTERNAL_HYPERLINK = hasattr(lib, "zlsx_sheet_writer_add_internal_hyperlink")
if _HAS_INTERNAL_HYPERLINK:
    lib.zlsx_sheet_writer_add_internal_hyperlink.argtypes = [
        sheet_writer_handle,
        ctypes.POINTER(ctypes.c_ubyte),
        ctypes.c_size_t,
        ctypes.POINTER(ctypes.c_ubyte),
        ctypes.c_size_t,
        ctypes.c_char_p,
        ctypes.c_size_t,
    ]
    lib.zlsx_sheet_writer_add_internal_hyperlink.restype = ctypes.c_int32


class CDxfBorderSide(ctypes.Structure):
    _fields_ = [
        ("style", ctypes.c_uint8),
        ("has_color", ctypes.c_uint8),
        ("_pad", ctypes.c_uint8 * 2),
        ("color_argb", ctypes.c_uint32),
    ]


class CDxf(ctypes.Structure):
    _fields_ = [
        ("bold", ctypes.c_uint8),
        ("italic", ctypes.c_uint8),
        ("has_color", ctypes.c_uint8),
        ("has_fill", ctypes.c_uint8),
        ("color_argb", ctypes.c_uint32),
        ("fill_fg_argb", ctypes.c_uint32),
        ("has_size", ctypes.c_uint8),
        ("_pad", ctypes.c_uint8 * 3),
        ("size", ctypes.c_float),
        ("border_left", CDxfBorderSide),
        ("border_right", CDxfBorderSide),
        ("border_top", CDxfBorderSide),
        ("border_bottom", CDxfBorderSide),
    ]


_HAS_CONDITIONAL_FORMAT = (
    hasattr(lib, "zlsx_writer_add_dxf")
    and hasattr(lib, "zlsx_sheet_writer_add_conditional_format_cell_is")
    and hasattr(lib, "zlsx_sheet_writer_add_conditional_format_expression")
)
if _HAS_CONDITIONAL_FORMAT:
    lib.zlsx_writer_add_dxf.argtypes = [
        writer_handle,
        ctypes.POINTER(CDxf),
        ctypes.POINTER(ctypes.c_uint32),
        ctypes.c_char_p,
        ctypes.c_size_t,
    ]
    lib.zlsx_writer_add_dxf.restype = ctypes.c_int32

    lib.zlsx_sheet_writer_add_conditional_format_cell_is.argtypes = [
        sheet_writer_handle,
        ctypes.POINTER(ctypes.c_ubyte),
        ctypes.c_size_t,
        ctypes.c_uint32,
        ctypes.POINTER(ctypes.c_ubyte),
        ctypes.c_size_t,
        ctypes.POINTER(ctypes.c_ubyte),
        ctypes.c_size_t,
        ctypes.c_uint32,
        ctypes.c_char_p,
        ctypes.c_size_t,
    ]
    lib.zlsx_sheet_writer_add_conditional_format_cell_is.restype = ctypes.c_int32

    lib.zlsx_sheet_writer_add_conditional_format_expression.argtypes = [
        sheet_writer_handle,
        ctypes.POINTER(ctypes.c_ubyte),
        ctypes.c_size_t,
        ctypes.POINTER(ctypes.c_ubyte),
        ctypes.c_size_t,
        ctypes.c_uint32,
        ctypes.c_char_p,
        ctypes.c_size_t,
    ]
    lib.zlsx_sheet_writer_add_conditional_format_expression.restype = ctypes.c_int32


_HAS_CF_GRADIENT = (
    hasattr(lib, "zlsx_sheet_writer_add_conditional_format_color_scale")
    and hasattr(lib, "zlsx_sheet_writer_add_conditional_format_data_bar")
)
if _HAS_CF_GRADIENT:
    lib.zlsx_sheet_writer_add_conditional_format_color_scale.argtypes = [
        sheet_writer_handle,
        ctypes.POINTER(ctypes.c_ubyte),
        ctypes.c_size_t,
        ctypes.c_uint32,
        ctypes.c_uint8,
        ctypes.c_uint32,
        ctypes.c_uint32,
        ctypes.c_char_p,
        ctypes.c_size_t,
    ]
    lib.zlsx_sheet_writer_add_conditional_format_color_scale.restype = ctypes.c_int32
    lib.zlsx_sheet_writer_add_conditional_format_data_bar.argtypes = [
        sheet_writer_handle,
        ctypes.POINTER(ctypes.c_ubyte),
        ctypes.c_size_t,
        ctypes.c_uint32,
        ctypes.c_char_p,
        ctypes.c_size_t,
    ]
    lib.zlsx_sheet_writer_add_conditional_format_data_bar.restype = ctypes.c_int32


_HAS_COMMENT_WRITER = hasattr(lib, "zlsx_sheet_writer_add_comment")
if _HAS_COMMENT_WRITER:
    lib.zlsx_sheet_writer_add_comment.argtypes = [
        sheet_writer_handle,
        ctypes.POINTER(ctypes.c_ubyte),
        ctypes.c_size_t,
        ctypes.POINTER(ctypes.c_ubyte),
        ctypes.c_size_t,
        ctypes.POINTER(ctypes.c_ubyte),
        ctypes.c_size_t,
        ctypes.c_char_p,
        ctypes.c_size_t,
    ]
    lib.zlsx_sheet_writer_add_comment.restype = ctypes.c_int32


# Reader metadata (libzlsx 0.2.5+): merged ranges + hyperlinks. Feature-
# probed like the writer additions so py-zlsx keeps importing against
# older dylibs.
class CMergeRange(ctypes.Structure):
    _fields_ = [
        ("top_left_col", ctypes.c_uint32),
        ("top_left_row", ctypes.c_uint32),
        ("bottom_right_col", ctypes.c_uint32),
        ("bottom_right_row", ctypes.c_uint32),
    ]


class CHyperlink(ctypes.Structure):
    _fields_ = [
        ("top_left_col", ctypes.c_uint32),
        ("top_left_row", ctypes.c_uint32),
        ("bottom_right_col", ctypes.c_uint32),
        ("bottom_right_row", ctypes.c_uint32),
        ("url_ptr", ctypes.POINTER(ctypes.c_ubyte)),
        ("url_len", ctypes.c_size_t),
    ]


_HAS_READER_META = hasattr(lib, "zlsx_merged_range_count")
if _HAS_READER_META:
    lib.zlsx_merged_range_count.argtypes = [book_handle, ctypes.c_uint32]
    lib.zlsx_merged_range_count.restype = ctypes.c_size_t
    lib.zlsx_merged_range_at.argtypes = [
        book_handle,
        ctypes.c_uint32,
        ctypes.c_size_t,
        ctypes.POINTER(CMergeRange),
    ]
    lib.zlsx_merged_range_at.restype = ctypes.c_int32

    lib.zlsx_hyperlink_count.argtypes = [book_handle, ctypes.c_uint32]
    lib.zlsx_hyperlink_count.restype = ctypes.c_size_t
    lib.zlsx_hyperlink_at.argtypes = [
        book_handle,
        ctypes.c_uint32,
        ctypes.c_size_t,
        ctypes.POINTER(CHyperlink),
    ]
    lib.zlsx_hyperlink_at.restype = ctypes.c_int32

# Internal-hyperlink `location` getter (libzlsx 0.2.7+). Probed
# independently because callers loading an older dylib that exposes
# `zlsx_hyperlink_at` but not the location accessor still get the
# external-URL surface.
_HAS_HYPERLINK_LOCATION = hasattr(lib, "zlsx_hyperlink_location_at")
if _HAS_HYPERLINK_LOCATION:
    lib.zlsx_hyperlink_location_at.argtypes = [
        book_handle,
        ctypes.c_uint32,
        ctypes.c_size_t,
        ctypes.POINTER(ctypes.POINTER(ctypes.c_ubyte)),
        ctypes.POINTER(ctypes.c_size_t),
    ]
    lib.zlsx_hyperlink_location_at.restype = ctypes.c_int32


class CComment(ctypes.Structure):
    _fields_ = [
        ("cell_col", ctypes.c_uint32),
        ("cell_row", ctypes.c_uint32),
        ("author_len", ctypes.c_size_t),
        ("author_ptr", ctypes.POINTER(ctypes.c_ubyte)),
        ("text_len", ctypes.c_size_t),
        ("text_ptr", ctypes.POINTER(ctypes.c_ubyte)),
    ]


_HAS_COMMENTS = (
    hasattr(lib, "zlsx_comment_count")
    and hasattr(lib, "zlsx_comment_at")
)
if _HAS_COMMENTS:
    lib.zlsx_comment_count.argtypes = [book_handle, ctypes.c_uint32]
    lib.zlsx_comment_count.restype = ctypes.c_size_t
    lib.zlsx_comment_at.argtypes = [
        book_handle,
        ctypes.c_uint32,
        ctypes.c_size_t,
        ctypes.POINTER(CComment),
    ]
    lib.zlsx_comment_at.restype = ctypes.c_int32


_HAS_COMMENT_RUNS = (
    hasattr(lib, "zlsx_comment_run_count")
    and hasattr(lib, "zlsx_comment_run_at")
)
if _HAS_COMMENT_RUNS:
    lib.zlsx_comment_run_count.argtypes = [
        book_handle,
        ctypes.c_uint32,
        ctypes.c_size_t,
    ]
    lib.zlsx_comment_run_count.restype = ctypes.c_size_t
    lib.zlsx_comment_run_at.argtypes = [
        book_handle,
        ctypes.c_uint32,
        ctypes.c_size_t,
        ctypes.c_size_t,
        ctypes.POINTER(ctypes.POINTER(ctypes.c_ubyte)),
        ctypes.POINTER(ctypes.c_size_t),
        ctypes.POINTER(ctypes.c_uint8),
        ctypes.POINTER(ctypes.c_uint8),
    ]
    lib.zlsx_comment_run_at.restype = ctypes.c_int32


class CDataValidation(ctypes.Structure):
    _fields_ = [
        ("top_left_col", ctypes.c_uint32),
        ("top_left_row", ctypes.c_uint32),
        ("bottom_right_col", ctypes.c_uint32),
        ("bottom_right_row", ctypes.c_uint32),
        ("values_count", ctypes.c_size_t),
    ]


_HAS_READER_DV = hasattr(lib, "zlsx_data_validation_count")
if _HAS_READER_DV:
    lib.zlsx_data_validation_count.argtypes = [book_handle, ctypes.c_uint32]
    lib.zlsx_data_validation_count.restype = ctypes.c_size_t
    lib.zlsx_data_validation_at.argtypes = [
        book_handle,
        ctypes.c_uint32,
        ctypes.c_size_t,
        ctypes.POINTER(CDataValidation),
    ]
    lib.zlsx_data_validation_at.restype = ctypes.c_int32
    lib.zlsx_data_validation_value_at.argtypes = [
        book_handle,
        ctypes.c_uint32,
        ctypes.c_size_t,
        ctypes.c_size_t,
        ctypes.POINTER(ctypes.POINTER(ctypes.c_ubyte)),
        ctypes.POINTER(ctypes.c_size_t),
    ]
    lib.zlsx_data_validation_value_at.restype = ctypes.c_int32


# Extended DV metadata (kind / operator / formula1 / formula2) was
# added in 0.2.6; probe each getter independently so bindings work
# against older libzlsx builds too.
_HAS_READER_DV_EXT = (
    _HAS_READER_DV
    and hasattr(lib, "zlsx_data_validation_kind")
    and hasattr(lib, "zlsx_data_validation_operator")
    and hasattr(lib, "zlsx_data_validation_formula1")
    and hasattr(lib, "zlsx_data_validation_formula2")
)
if _HAS_READER_DV_EXT:
    lib.zlsx_data_validation_kind.argtypes = [
        book_handle,
        ctypes.c_uint32,
        ctypes.c_size_t,
    ]
    lib.zlsx_data_validation_kind.restype = ctypes.c_uint32
    lib.zlsx_data_validation_operator.argtypes = [
        book_handle,
        ctypes.c_uint32,
        ctypes.c_size_t,
    ]
    lib.zlsx_data_validation_operator.restype = ctypes.c_uint32
    lib.zlsx_data_validation_formula1.argtypes = [
        book_handle,
        ctypes.c_uint32,
        ctypes.c_size_t,
        ctypes.POINTER(ctypes.POINTER(ctypes.c_ubyte)),
        ctypes.POINTER(ctypes.c_size_t),
    ]
    lib.zlsx_data_validation_formula1.restype = ctypes.c_int32
    lib.zlsx_data_validation_formula2.argtypes = [
        book_handle,
        ctypes.c_uint32,
        ctypes.c_size_t,
        ctypes.POINTER(ctypes.POINTER(ctypes.c_ubyte)),
        ctypes.POINTER(ctypes.c_size_t),
    ]
    lib.zlsx_data_validation_formula2.restype = ctypes.c_int32


# Shared-string enumeration — added 0.2.6+. Pairs with rich_text to let
# Python callers discover which SST entries carry formatted runs
# without hand-tracking indices.
_HAS_SST_ENUM = (
    hasattr(lib, "zlsx_shared_string_count")
    and hasattr(lib, "zlsx_shared_string_at")
)
if _HAS_SST_ENUM:
    lib.zlsx_shared_string_count.argtypes = [book_handle]
    lib.zlsx_shared_string_count.restype = ctypes.c_size_t
    lib.zlsx_shared_string_at.argtypes = [
        book_handle,
        ctypes.c_size_t,
        ctypes.POINTER(ctypes.POINTER(ctypes.c_ubyte)),
        ctypes.POINTER(ctypes.c_size_t),
    ]
    lib.zlsx_shared_string_at.restype = ctypes.c_int32


# Rich-text run reading — added in 0.2.6+. Plain single-run SST entries
# return 0 from rich_run_count so callers can skip them zero-cost.
_HAS_RICH_RUNS = (
    hasattr(lib, "zlsx_rich_run_count")
    and hasattr(lib, "zlsx_rich_run_at")
)
if _HAS_RICH_RUNS:
    lib.zlsx_rich_run_count.argtypes = [book_handle, ctypes.c_size_t]
    lib.zlsx_rich_run_count.restype = ctypes.c_size_t
    lib.zlsx_rich_run_at.argtypes = [
        book_handle,
        ctypes.c_size_t,
        ctypes.c_size_t,
        ctypes.POINTER(ctypes.POINTER(ctypes.c_ubyte)),
        ctypes.POINTER(ctypes.c_size_t),
        ctypes.POINTER(ctypes.c_uint8),
        ctypes.POINTER(ctypes.c_uint8),
    ]
    lib.zlsx_rich_run_at.restype = ctypes.c_int32

# Rich-text extended props (color / size / font_name) — feature-probed
# independently so a partial libzlsx still loads.
_HAS_RICH_RUNS_EXT = (
    _HAS_RICH_RUNS
    and hasattr(lib, "zlsx_rich_run_color")
    and hasattr(lib, "zlsx_rich_run_size")
    and hasattr(lib, "zlsx_rich_run_font_name")
)
if _HAS_RICH_RUNS_EXT:
    lib.zlsx_rich_run_color.argtypes = [
        book_handle,
        ctypes.c_size_t,
        ctypes.c_size_t,
        ctypes.POINTER(ctypes.c_uint32),
    ]
    lib.zlsx_rich_run_color.restype = ctypes.c_int32
    lib.zlsx_rich_run_size.argtypes = [
        book_handle,
        ctypes.c_size_t,
        ctypes.c_size_t,
        ctypes.POINTER(ctypes.c_float),
    ]
    lib.zlsx_rich_run_size.restype = ctypes.c_int32
    lib.zlsx_rich_run_font_name.argtypes = [
        book_handle,
        ctypes.c_size_t,
        ctypes.c_size_t,
        ctypes.POINTER(ctypes.POINTER(ctypes.c_ubyte)),
        ctypes.POINTER(ctypes.c_size_t),
    ]
    lib.zlsx_rich_run_font_name.restype = ctypes.c_int32


# Number-format / per-cell style-index surface — added in 0.2.6+.
class CDateTime(ctypes.Structure):
    _fields_ = [
        ("year", ctypes.c_uint16),
        ("month", ctypes.c_uint8),
        ("day", ctypes.c_uint8),
        ("hour", ctypes.c_uint8),
        ("minute", ctypes.c_uint8),
        ("second", ctypes.c_uint8),
        ("_pad", ctypes.c_uint8),
    ]


_HAS_PARSE_DATE = hasattr(lib, "zlsx_rows_parse_date")
if _HAS_PARSE_DATE:
    lib.zlsx_rows_parse_date.argtypes = [
        rows_handle,
        ctypes.c_size_t,
        ctypes.POINTER(CDateTime),
    ]
    lib.zlsx_rows_parse_date.restype = ctypes.c_int32


_HAS_TO_EXCEL_SERIAL = hasattr(lib, "zlsx_datetime_to_serial")
if _HAS_TO_EXCEL_SERIAL:
    lib.zlsx_datetime_to_serial.argtypes = [
        ctypes.POINTER(CDateTime),
        ctypes.POINTER(ctypes.c_double),
    ]
    lib.zlsx_datetime_to_serial.restype = ctypes.c_int32


_HAS_NUM_FMT = (
    hasattr(lib, "zlsx_rows_style_at")
    and hasattr(lib, "zlsx_number_format")
    and hasattr(lib, "zlsx_is_date_format")
)
if _HAS_NUM_FMT:
    lib.zlsx_rows_style_at.argtypes = [
        rows_handle,
        ctypes.c_size_t,
        ctypes.POINTER(ctypes.c_uint32),
    ]
    lib.zlsx_rows_style_at.restype = ctypes.c_int32
    lib.zlsx_number_format.argtypes = [
        book_handle,
        ctypes.c_uint32,
        ctypes.POINTER(ctypes.POINTER(ctypes.c_ubyte)),
        ctypes.POINTER(ctypes.c_size_t),
    ]
    lib.zlsx_number_format.restype = ctypes.c_int32
    lib.zlsx_is_date_format.argtypes = [book_handle, ctypes.c_uint32]
    lib.zlsx_is_date_format.restype = ctypes.c_uint8


class CCellFont(ctypes.Structure):
    _fields_ = [
        ("bold", ctypes.c_uint8),
        ("italic", ctypes.c_uint8),
        ("has_color", ctypes.c_uint8),
        ("has_size", ctypes.c_uint8),
        ("color_argb", ctypes.c_uint32),
        ("size", ctypes.c_float),
        ("name_len", ctypes.c_size_t),
        ("name_ptr", ctypes.POINTER(ctypes.c_ubyte)),
    ]


_HAS_CELL_FONT = hasattr(lib, "zlsx_cell_font")
if _HAS_CELL_FONT:
    lib.zlsx_cell_font.argtypes = [
        book_handle,
        ctypes.c_uint32,
        ctypes.POINTER(CCellFont),
    ]
    lib.zlsx_cell_font.restype = ctypes.c_int32


class CCellFill(ctypes.Structure):
    _fields_ = [
        ("has_fg", ctypes.c_uint8),
        ("has_bg", ctypes.c_uint8),
        ("_pad", ctypes.c_uint8 * 2),
        ("fg_color_argb", ctypes.c_uint32),
        ("bg_color_argb", ctypes.c_uint32),
        ("pattern_len", ctypes.c_size_t),
        ("pattern_ptr", ctypes.POINTER(ctypes.c_ubyte)),
    ]


_HAS_CELL_FILL = hasattr(lib, "zlsx_cell_fill")
if _HAS_CELL_FILL:
    lib.zlsx_cell_fill.argtypes = [
        book_handle,
        ctypes.c_uint32,
        ctypes.POINTER(CCellFill),
    ]
    lib.zlsx_cell_fill.restype = ctypes.c_int32


class CBorderSide(ctypes.Structure):
    _fields_ = [
        ("has_color", ctypes.c_uint8),
        ("_pad", ctypes.c_uint8 * 3),
        ("color_argb", ctypes.c_uint32),
        ("style_len", ctypes.c_size_t),
        ("style_ptr", ctypes.POINTER(ctypes.c_ubyte)),
    ]


class CCellBorder(ctypes.Structure):
    _fields_ = [
        ("left", CBorderSide),
        ("right", CBorderSide),
        ("top", CBorderSide),
        ("bottom", CBorderSide),
        ("diagonal", CBorderSide),
    ]


_HAS_CELL_BORDER = hasattr(lib, "zlsx_cell_border")
if _HAS_CELL_BORDER:
    lib.zlsx_cell_border.argtypes = [
        book_handle,
        ctypes.c_uint32,
        ctypes.POINTER(CCellBorder),
    ]
    lib.zlsx_cell_border.restype = ctypes.c_int32


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


# ─── CStyle layout guard ──────────────────────────────────────────────
#
# Matches Zig's `comptime` assertion in src/c_abi.zig. If either side
# reorders a field or changes padding, one binding will silently corrupt
# the other. Catch that at import time with a clear error that points
# the reader at both sides.

_EXPECTED_CSTYLE_SIZE_64 = 88
_EXPECTED_CSTYLE_SIZE_32 = 68
_actual_cstyle_size = ctypes.sizeof(CStyle)
if _actual_cstyle_size not in (_EXPECTED_CSTYLE_SIZE_64, _EXPECTED_CSTYLE_SIZE_32):
    raise ImportError(
        f"CStyle layout drift: expected {_EXPECTED_CSTYLE_SIZE_64} (64-bit) or "
        f"{_EXPECTED_CSTYLE_SIZE_32} (32-bit), got {_actual_cstyle_size}. "
        "bindings/python/zlsx/_ffi.py's CStyle._fields_ must match "
        "src/c_abi.zig's `extern struct CStyle` exactly."
    )

# Load-bearing field offsets — anything else the Zig comptime assertion
# pins, we pin here too.
for _name, _expected in [
    ("font_size", 8),
    ("font_color_argb", 12),
    ("fill_fg_argb", 16),
    ("fill_bg_argb", 20),
    ("border_left_style", 24),
    ("diagonal_down", 30),
    ("border_left_color_argb", 32),
    ("border_diagonal_color_argb", 48),
]:
    _got = getattr(CStyle, _name).offset
    if _got != _expected:
        raise ImportError(
            f"CStyle.{_name} offset drift: expected {_expected}, got {_got}"
        )
del _name, _expected, _got, _actual_cstyle_size

# ─── ABI version check ────────────────────────────────────────────────

EXPECTED_ABI_VERSION = 1
_found_abi = lib.zlsx_abi_version()
if _found_abi != EXPECTED_ABI_VERSION:
    raise ImportError(
        f"libzlsx ABI mismatch: py-zlsx expects v{EXPECTED_ABI_VERSION}, "
        f"loaded library reports v{_found_abi}. Upgrade one of them."
    )

# Editor (libzlsx 0.2.7+): load-modify-save append path. Probed
# independently so older dylibs that ship the rest of the surface
# without `zlsx_editor_open` keep importing.
_HAS_EDITOR = hasattr(lib, "zlsx_editor_open")
if _HAS_EDITOR:
    lib.zlsx_editor_open.argtypes = [
        ctypes.c_char_p,
        ctypes.c_char_p,
        ctypes.c_size_t,
    ]
    lib.zlsx_editor_open.restype = editor_handle

    lib.zlsx_editor_close.argtypes = [editor_handle]
    lib.zlsx_editor_close.restype = None

    lib.zlsx_editor_append_row.argtypes = [
        editor_handle,
        ctypes.c_uint32,
        cell_ptr,
        ctypes.c_size_t,
        ctypes.c_char_p,
        ctypes.c_size_t,
    ]
    lib.zlsx_editor_append_row.restype = ctypes.c_int32

    lib.zlsx_editor_save.argtypes = [
        editor_handle,
        ctypes.c_char_p,
        ctypes.c_size_t,
        ctypes.c_char_p,
        ctypes.c_size_t,
    ]
    lib.zlsx_editor_save.restype = ctypes.c_int32


# ---- Document properties (Z3) ---------------------------------------
#
# Feature-probed: an older libzlsx on the system must not break the
# import, per the binding's compatibility contract.
DOCPROP_CREATOR = 0
DOCPROP_LAST_MODIFIED_BY = 1
DOCPROP_TITLE = 2
DOCPROP_SUBJECT = 3
DOCPROP_DESCRIPTION = 4
DOCPROP_KEYWORDS = 5
DOCPROP_CATEGORY = 6
DOCPROP_CREATED = 7
DOCPROP_MODIFIED = 8
DOCPROP_REVISION = 9
DOCPROP_COMPANY = 10
DOCPROP_MANAGER = 11
DOCPROP_APPLICATION = 12
DOCPROP_HYPERLINK_BASE = 13

DOCPROP_FIELDS = {
    "creator": DOCPROP_CREATOR,
    "last_modified_by": DOCPROP_LAST_MODIFIED_BY,
    "title": DOCPROP_TITLE,
    "subject": DOCPROP_SUBJECT,
    "description": DOCPROP_DESCRIPTION,
    "keywords": DOCPROP_KEYWORDS,
    "category": DOCPROP_CATEGORY,
    "created": DOCPROP_CREATED,
    "modified": DOCPROP_MODIFIED,
    "revision": DOCPROP_REVISION,
    "company": DOCPROP_COMPANY,
    "manager": DOCPROP_MANAGER,
    "application": DOCPROP_APPLICATION,
    "hyperlink_base": DOCPROP_HYPERLINK_BASE,
}

_HAS_DOCPROPS = hasattr(lib, "zlsx_editor_docprop_at")
if _HAS_DOCPROPS:
    lib.zlsx_editor_docprop_at.argtypes = [
        editor_handle,
        ctypes.c_uint32,                            # field selector
        ctypes.POINTER(ctypes.POINTER(ctypes.c_ubyte)),
        ctypes.POINTER(ctypes.c_size_t),
    ]
    lib.zlsx_editor_docprop_at.restype = ctypes.c_int32

    lib.zlsx_editor_has_custom_properties.argtypes = [editor_handle]
    lib.zlsx_editor_has_custom_properties.restype = ctypes.c_int32

    lib.zlsx_editor_strip_doc_props.argtypes = [
        editor_handle,
        ctypes.c_int32,                             # strip_timestamps
        ctypes.c_char_p,
        ctypes.c_size_t,
    ]
    lib.zlsx_editor_strip_doc_props.restype = ctypes.c_int32

_HAS_EDITOR_SET_CELL = hasattr(lib, "zlsx_editor_set_cell")
if _HAS_EDITOR_SET_CELL:
    lib.zlsx_editor_set_cell.argtypes = [
        editor_handle,
        ctypes.c_uint32,           # sheet_idx
        ctypes.c_uint32,           # row (1-based)
        ctypes.c_uint32,           # col (0-based)
        ctypes.POINTER(Cell),
        ctypes.c_char_p,
        ctypes.c_size_t,
    ]
    lib.zlsx_editor_set_cell.restype = ctypes.c_int32

# ─── New writer/reader exports added in this iteration ───────────────

_HAS_DEFINED_NAME = hasattr(lib, "zlsx_writer_add_defined_name")
if _HAS_DEFINED_NAME:
    lib.zlsx_writer_add_defined_name.argtypes = [
        writer_handle,
        ctypes.POINTER(ctypes.c_ubyte),    # name_ptr
        ctypes.c_size_t,                   # name_len
        ctypes.POINTER(ctypes.c_ubyte),    # refers_to_ptr
        ctypes.c_size_t,                   # refers_to_len
        ctypes.c_int32,                    # local_sheet_id_neg (negative=workbook scope)
        ctypes.c_uint8,                    # hidden_flag
        ctypes.c_char_p,
        ctypes.c_size_t,
    ]
    lib.zlsx_writer_add_defined_name.restype = ctypes.c_int32

_HAS_SET_ROW_HEIGHT = hasattr(lib, "zlsx_sheet_writer_set_row_height")
if _HAS_SET_ROW_HEIGHT:
    lib.zlsx_sheet_writer_set_row_height.argtypes = [
        sheet_writer_handle,
        ctypes.c_uint32,                   # row_idx (0-based)
        ctypes.c_float,                    # height
        ctypes.c_char_p,
        ctypes.c_size_t,
    ]
    lib.zlsx_sheet_writer_set_row_height.restype = ctypes.c_int32

_HAS_FREEZE_PANES_CHECKED = hasattr(lib, "zlsx_sheet_writer_freeze_panes_checked")
if _HAS_FREEZE_PANES_CHECKED:
    lib.zlsx_sheet_writer_freeze_panes_checked.argtypes = [
        sheet_writer_handle,
        ctypes.c_uint32,                   # rows
        ctypes.c_uint32,                   # cols
        ctypes.c_char_p,
        ctypes.c_size_t,
    ]
    lib.zlsx_sheet_writer_freeze_panes_checked.restype = ctypes.c_int32


class CellAlignment(ctypes.Structure):
    """Mirror of `zlsx_cell_alignment_t` from include/zlsx.h.

    `horizontal_len == 0` means the alignment is the OOXML default
    ("general", which the emitter omits). `wrap_text == 1` when the
    cell's `<alignment wrapText="1"/>` was set.
    """
    _fields_ = [
        ("horizontal_len", ctypes.c_size_t),
        ("horizontal_ptr", ctypes.POINTER(ctypes.c_ubyte)),
        ("wrap_text", ctypes.c_uint8),
        ("_pad", ctypes.c_uint8 * 7),
    ]

_HAS_CELL_ALIGNMENT = hasattr(lib, "zlsx_cell_alignment")
if _HAS_CELL_ALIGNMENT:
    lib.zlsx_cell_alignment.argtypes = [
        book_handle,
        ctypes.c_uint32,
        ctypes.POINTER(CellAlignment),
    ]
    lib.zlsx_cell_alignment.restype = ctypes.c_int32


# ─── Embeddings (E5) ─────────────────────────────────────────────────

emb_handle = ctypes.c_void_p

ZLSX_EMB_ABSENT = 0
ZLSX_EMB_PRESENT = 1
ZLSX_EMB_STRIPPED = 2

ZLSX_EMB_CARRIER_DEFINED_NAME = 0
ZLSX_EMB_CARRIER_DOC_PROPS = 1
ZLSX_EMB_CARRIER_CELL_DATA = 2

# Feature-gated like the editor surface: a wheel built against an older
# libzlsx must still import, with the embedding API simply absent.
_HAS_EMB = hasattr(lib, "zlsx_emb_open")
if _HAS_EMB:
    lib.zlsx_emb_open.argtypes = [
        ctypes.c_char_p,
        ctypes.c_char_p,
        ctypes.c_size_t,
    ]
    lib.zlsx_emb_open.restype = emb_handle

    lib.zlsx_emb_close.argtypes = [emb_handle]
    lib.zlsx_emb_close.restype = None

    lib.zlsx_emb_state.argtypes = [emb_handle]
    lib.zlsx_emb_state.restype = ctypes.c_uint32

    for _name in ("zlsx_emb_model", "zlsx_emb_dtype"):
        _fn = getattr(lib, _name)
        _fn.argtypes = [emb_handle, ctypes.c_char_p, ctypes.c_size_t]
        _fn.restype = ctypes.c_size_t

    lib.zlsx_emb_dim.argtypes = [emb_handle]
    lib.zlsx_emb_dim.restype = ctypes.c_uint32

    lib.zlsx_emb_coverage_count.argtypes = [emb_handle]
    lib.zlsx_emb_coverage_count.restype = ctypes.c_size_t

    for _name in (
        "zlsx_emb_coverage_id",
        "zlsx_emb_coverage_range",
        "zlsx_emb_coverage_sheet",
    ):
        _fn = getattr(lib, _name)
        _fn.argtypes = [
            emb_handle,
            ctypes.c_size_t,
            ctypes.c_char_p,
            ctypes.c_size_t,
        ]
        _fn.restype = ctypes.c_size_t

    lib.zlsx_emb_coverage_rows.argtypes = [emb_handle, ctypes.c_size_t]
    lib.zlsx_emb_coverage_rows.restype = ctypes.c_uint32

    lib.zlsx_emb_digest.argtypes = [emb_handle]
    lib.zlsx_emb_digest.restype = ctypes.c_uint64

    lib.zlsx_emb_carrier.argtypes = [emb_handle]
    lib.zlsx_emb_carrier.restype = ctypes.c_uint32

    lib.zlsx_emb_tombstone.argtypes = []
    lib.zlsx_emb_tombstone.restype = ctypes.c_uint64

    lib.zlsx_emb_vectors.argtypes = [
        emb_handle,
        ctypes.c_size_t,
        ctypes.POINTER(ctypes.c_float),
        ctypes.c_size_t,
    ]
    lib.zlsx_emb_vectors.restype = ctypes.c_int32

    lib.zlsx_emb_hashes.argtypes = [
        emb_handle,
        ctypes.c_size_t,
        ctypes.POINTER(ctypes.c_uint64),
        ctypes.c_size_t,
    ]
    lib.zlsx_emb_hashes.restype = ctypes.c_int32


# ─── Formula engine (M9a1) ────────────────────────────────────────────
#
# zlsx_status_v1 exports. Every struct mirrors include/zlsx.h and the
# committed layout note (docs/plans/c-abi-status-v1.md). Each symbol
# group is feature-probed independently — an older dylib simply lacks
# the export, and the corresponding _HAS_* flag gates whatever the
# high-level API builds on it (M9a2). Callers set struct_size before
# every call; the library never writes beyond the v1 size it knows.

ZLSX_OK = 0
ZLSX_ERROR = -1
ZLSX_REFUSED = -2
ZLSX_NOMEM = -3
ZLSX_CANCELLED = -5

ZLSX_PLANE_NONE = 0xFFFFFFFF
ZLSX_DIALECT_NONE = 0xFFFFFFFF

cancel_token_handle = ctypes.c_void_p


class CensusEntryV1(ctypes.Structure):
    """Mirrors zlsx_census_entry_v1."""

    _fields_ = [
        ("plane", ctypes.c_uint32),
        ("sheet", ctypes.c_uint32),
        ("row", ctypes.c_uint32),
        ("col", ctypes.c_uint32),
    ]


class DiagV1(ctypes.Structure):
    """Mirrors zlsx_diag_v1. census is library-owned: zlsx_diag_release."""

    _fields_ = [
        ("struct_size", ctypes.c_size_t),
        ("plane", ctypes.c_uint32),
        ("census_truncated", ctypes.c_uint32),
        ("error_name", ctypes.c_char * 64),
        ("census", ctypes.POINTER(CensusEntryV1)),
        ("census_len", ctypes.c_size_t),
    ]


class RunV1(ctypes.Structure):
    """Mirrors zlsx_run_v1. Limit fields: 0 = documented default."""

    _fields_ = [
        ("struct_size", ctypes.c_size_t),
        ("now_utc_ms", ctypes.c_int64),
        ("rng_seed", ctypes.c_uint64),
        ("utc_offset_min", ctypes.c_int32),
        ("fidelity", ctypes.c_uint32),
        ("profile", ctypes.c_uint32),
        ("dialect", ctypes.c_uint32),
        ("on_unsupported", ctypes.c_uint32),
        ("_reserved0", ctypes.c_uint32),
        ("max_run_arena_bytes", ctypes.c_uint64),
        ("max_matrix_cells", ctypes.c_uint64),
        ("max_string_payload_bytes", ctypes.c_uint64),
        ("max_retained_ast_bytes", ctypes.c_uint64),
        ("max_diagnostics_bytes", ctypes.c_uint64),
        ("timeout_ms", ctypes.c_uint64),
        ("cancel", cancel_token_handle),
    ]


class ResolvedV1(ctypes.Structure):
    """Mirrors zlsx_resolved_v1 — the §5.5 echo."""

    _fields_ = [
        ("struct_size", ctypes.c_size_t),
        ("now_utc_ms", ctypes.c_int64),
        ("rng_seed", ctypes.c_uint64),
        ("utc_offset_min", ctypes.c_int32),
        ("fidelity", ctypes.c_uint32),
        ("profile", ctypes.c_uint32),
        ("dialect", ctypes.c_uint32),
        ("max_run_arena_bytes", ctypes.c_uint64),
        ("max_matrix_cells", ctypes.c_uint64),
        ("max_string_payload_bytes", ctypes.c_uint64),
        ("max_retained_ast_bytes", ctypes.c_uint64),
        ("max_diagnostics_bytes", ctypes.c_uint64),
    ]


class RecalcReportV1(ctypes.Structure):
    """Mirrors zlsx_recalc_report_v1. census: zlsx_recalc_report_release."""

    _fields_ = [
        ("struct_size", ctypes.c_size_t),
        ("sheets_patched", ctypes.c_uint32),
        ("cells_written", ctypes.c_uint32),
        ("passes", ctypes.c_uint32),
        ("non_converged_cells", ctypes.c_uint32),
        ("dynamic_passes", ctypes.c_uint32),
        ("kept_stale", ctypes.c_uint32),
        ("calc_chain_removed", ctypes.c_uint32),
        ("census_truncated", ctypes.c_uint32),
        ("retained_generations", ctypes.c_uint64),
        ("retained_bytes", ctypes.c_uint64),
        ("durability_warning", ctypes.c_uint32),
        ("durability_errno", ctypes.c_int32),
        ("resolved", ResolvedV1),
        ("resolved_present", ctypes.c_uint32),
        ("_reserved0", ctypes.c_uint32),
        ("census", ctypes.POINTER(CensusEntryV1)),
        ("census_len", ctypes.c_size_t),
    ]


class ValueElemV1(ctypes.Structure):
    """Mirrors zlsx_value_elem_v1 — {tag; num; payload_off, payload_len}."""

    _fields_ = [
        ("tag", ctypes.c_uint8),
        ("_reserved", ctypes.c_uint8 * 7),
        ("num", ctypes.c_double),
        ("payload_off", ctypes.c_uint64),
        ("payload_len", ctypes.c_uint64),
    ]


class ValueV1(ctypes.Structure):
    """Mirrors zlsx_value_v1. elems/payload: zlsx_value_release."""

    _fields_ = [
        ("struct_size", ctypes.c_size_t),
        ("rows", ctypes.c_uint32),
        ("cols", ctypes.c_uint32),
        ("is_matrix", ctypes.c_uint32),
        ("_reserved0", ctypes.c_uint32),
        ("elems", ctypes.POINTER(ValueElemV1)),
        ("elems_len", ctypes.c_size_t),
        ("payload", ctypes.POINTER(ctypes.c_uint8)),
        ("payload_len", ctypes.c_size_t),
    ]


# ctypes has no size_t-championed static assert; pin the mirrored sizes
# so a drifted field order fails at import, not at call time.
assert ctypes.sizeof(CensusEntryV1) == 16
assert ctypes.sizeof(DiagV1) == 96
assert ctypes.sizeof(RunV1) == 104
assert ctypes.sizeof(ResolvedV1) == 80
assert ctypes.sizeof(RecalcReportV1) == 168
assert ctypes.sizeof(ValueElemV1) == 32
assert ctypes.sizeof(ValueV1) == 56

_HAS_FINGERPRINT = hasattr(lib, "zlsx_engine_fingerprint")
if _HAS_FINGERPRINT:
    lib.zlsx_engine_fingerprint.argtypes = []
    lib.zlsx_engine_fingerprint.restype = ctypes.c_char_p

_HAS_CANCEL = hasattr(lib, "zlsx_cancel_token_new")
if _HAS_CANCEL:
    lib.zlsx_cancel_token_new.argtypes = [
        ctypes.POINTER(cancel_token_handle),
        ctypes.c_char_p,
        ctypes.c_size_t,
    ]
    lib.zlsx_cancel_token_new.restype = ctypes.c_int32

    lib.zlsx_cancel_token_trigger.argtypes = [cancel_token_handle]
    lib.zlsx_cancel_token_trigger.restype = None

    lib.zlsx_cancel_token_free.argtypes = [cancel_token_handle]
    lib.zlsx_cancel_token_free.restype = None

_HAS_MARK_RECALC = hasattr(lib, "zlsx_editor_mark_recalc_on_load")
if _HAS_MARK_RECALC:
    lib.zlsx_editor_mark_recalc_on_load.argtypes = [
        editor_handle,
        ctypes.POINTER(DiagV1),
        ctypes.c_char_p,
        ctypes.c_size_t,
    ]
    lib.zlsx_editor_mark_recalc_on_load.restype = ctypes.c_int32

    lib.zlsx_diag_release.argtypes = [ctypes.POINTER(DiagV1)]
    lib.zlsx_diag_release.restype = None

_HAS_RECALC = hasattr(lib, "zlsx_editor_recalculate")
if _HAS_RECALC:
    lib.zlsx_editor_recalculate.argtypes = [
        editor_handle,
        ctypes.POINTER(RunV1),
        ctypes.POINTER(RecalcReportV1),
        ctypes.POINTER(DiagV1),
        ctypes.c_char_p,
        ctypes.c_size_t,
    ]
    lib.zlsx_editor_recalculate.restype = ctypes.c_int32

    lib.zlsx_recalc_report_release.argtypes = [ctypes.POINTER(RecalcReportV1)]
    lib.zlsx_recalc_report_release.restype = None

_HAS_EVAL = hasattr(lib, "zlsx_editor_evaluate")
if _HAS_EVAL:
    lib.zlsx_editor_evaluate.argtypes = [
        editor_handle,
        ctypes.POINTER(ctypes.c_uint8),  # formula_ptr
        ctypes.c_size_t,  # formula_len
        ctypes.c_uint32,  # sheet_idx
        ctypes.c_uint32,  # anchor_row (1-based; 0 = absent)
        ctypes.c_uint32,  # anchor_col (0-based)
        ctypes.POINTER(RunV1),
        ctypes.POINTER(ValueV1),
        ctypes.POINTER(ResolvedV1),
        ctypes.POINTER(DiagV1),
        ctypes.c_char_p,
        ctypes.c_size_t,
    ]
    lib.zlsx_editor_evaluate.restype = ctypes.c_int32

    lib.zlsx_value_release.argtypes = [ctypes.POINTER(ValueV1)]
    lib.zlsx_value_release.restype = None

# ─── Formula engine (M9a2): buffers, file transaction, writer ─────────
#
# Part 2 of the zlsx_status_v1 surface. Same probing discipline: each
# group gates exactly the high-level methods built on it, and an older
# dylib keeps importing with the flag off.

ZLSX_FORMULA_SCALAR = 0
ZLSX_FORMULA_DYNAMIC_ARRAY = 1
ZLSX_FORMULA_CSE = 2

# §10's fourteen-plane vocabulary in ZLSX_PLANE_* order (the enum's
# declaration order is ABI, pinned by the Zig-side test). Index with a
# diag/census plane value to get the error name-shaped label.
PLANE_NAMES = (
    "FormulaUnsupportedFunction",
    "FormulaUnsupportedConstruct",
    "FormulaPrecisionAsDisplayed",
    "FormulaMalformedInput",
    "FormulaLocaleSensitiveInput",
    "FormulaDataTableUnsupported",
    "FormulaSignedWorkbook",
    "FormulaStaleEmbeddings",
    "FormulaAnchorRequired",
    "FormulaCycle",
    "FormulaDynamicRefUnstable",
    "FormulaSpillPersistUnsupported",
    "FormulaResultNotRepresentable",
    "FormulaLimitExceeded",
)


class FormulaCellV1(ctypes.Structure):
    """Mirrors zlsx_formula_cell_v1 — §12.3's per-cell descriptor."""

    _fields_ = [
        ("text", ctypes.POINTER(ctypes.c_ubyte)),
        ("text_len", ctypes.c_size_t),
        ("dialect", ctypes.c_uint32),
        ("_reserved0", ctypes.c_uint32),
        ("ref", ctypes.POINTER(ctypes.c_ubyte)),
        ("ref_len", ctypes.c_size_t),
    ]


assert ctypes.sizeof(FormulaCellV1) == 40

# Composite: Editor.save_to_buffer / Editor.from_bytes need all three
# symbols, and a dylib carrying one without the others is not a shape
# any release ever shipped.
_HAS_SAVE_BUFFER = (
    hasattr(lib, "zlsx_editor_save_to_buffer")
    and hasattr(lib, "zlsx_open_buffer")
    and hasattr(lib, "zlsx_buffer_release")
)
if _HAS_SAVE_BUFFER:
    lib.zlsx_editor_save_to_buffer.argtypes = [
        editor_handle,
        ctypes.POINTER(ctypes.POINTER(ctypes.c_ubyte)),  # out_ptr
        ctypes.POINTER(ctypes.c_size_t),                 # out_len
        ctypes.c_char_p,
        ctypes.c_size_t,
    ]
    lib.zlsx_editor_save_to_buffer.restype = ctypes.c_int32

    lib.zlsx_open_buffer.argtypes = [
        ctypes.POINTER(ctypes.c_ubyte),  # data
        ctypes.c_size_t,                 # data_len
        ctypes.POINTER(editor_handle),   # out
        ctypes.c_char_p,
        ctypes.c_size_t,
    ]
    lib.zlsx_open_buffer.restype = ctypes.c_int32

    lib.zlsx_buffer_release.argtypes = [
        ctypes.POINTER(ctypes.c_ubyte),
        ctypes.c_size_t,
    ]
    lib.zlsx_buffer_release.restype = None

_HAS_SAVE_WITH_RECALC = hasattr(lib, "zlsx_editor_save_with_recalc")
if _HAS_SAVE_WITH_RECALC:
    lib.zlsx_editor_save_with_recalc.argtypes = [
        editor_handle,
        ctypes.POINTER(ctypes.c_ubyte),  # out_path_ptr
        ctypes.c_size_t,                 # out_path_len
        ctypes.POINTER(RunV1),
        ctypes.POINTER(RecalcReportV1),
        ctypes.POINTER(DiagV1),
        ctypes.c_char_p,
        ctypes.c_size_t,
    ]
    lib.zlsx_editor_save_with_recalc.restype = ctypes.c_int32

_HAS_WRITER_RECALC = hasattr(lib, "zlsx_writer_save_with_recalc")
if _HAS_WRITER_RECALC:
    lib.zlsx_writer_save_with_recalc.argtypes = [
        writer_handle,
        ctypes.POINTER(ctypes.c_ubyte),  # out_path_ptr
        ctypes.c_size_t,                 # out_path_len
        ctypes.POINTER(RunV1),
        ctypes.POINTER(RecalcReportV1),
        ctypes.POINTER(DiagV1),
        ctypes.c_char_p,
        ctypes.c_size_t,
    ]
    lib.zlsx_writer_save_with_recalc.restype = ctypes.c_int32

_HAS_FORMULAS_V2 = hasattr(lib, "zlsx_sheet_writer_write_row_with_formulas_v2")
if _HAS_FORMULAS_V2:
    lib.zlsx_sheet_writer_write_row_with_formulas_v2.argtypes = [
        sheet_writer_handle,
        ctypes.POINTER(Cell),           # cells
        ctypes.POINTER(FormulaCellV1),  # formulas (parallel)
        ctypes.c_size_t,                # cells_len
        ctypes.c_char_p,
        ctypes.c_size_t,
    ]
    lib.zlsx_sheet_writer_write_row_with_formulas_v2.restype = ctypes.c_int32

# ---- S3a: structural edits + the pivots read (libzlsx 0.9.0+) --------
#
# zlsx_status_v1 exports: every one takes a nullable zlsx_diag_v1 and
# returns ZLSX_OK / ZLSX_ERROR / ZLSX_REFUSED / ZLSX_NOMEM. Rows are
# 1-based, columns 0-based, names are (ptr, len) UTF-8. Two capability
# probes — structural edits, pivots — and each also requires the
# release symbols its Python wrappers call unconditionally
# (`zlsx_diag_release`; the pivots read adds `zlsx_buffer_release`), so
# a feature-stripped dylib never advertises a capability it cannot
# clean up after (Codex #207 r6 REL-601).

ZLSX_NO_SHEET_IDX = 0xFFFFFFFF

# Probed and configured on their own: the same symbols are configured
# under the M9a1/M9a2 probes when those features are present, but the
# S3a capabilities must not depend on unrelated features carrying the
# declarations for them.
_HAS_DIAG_RELEASE = hasattr(lib, "zlsx_diag_release")
if _HAS_DIAG_RELEASE:
    lib.zlsx_diag_release.argtypes = [ctypes.POINTER(DiagV1)]
    lib.zlsx_diag_release.restype = None
_HAS_BUFFER_RELEASE = hasattr(lib, "zlsx_buffer_release")
if _HAS_BUFFER_RELEASE:
    lib.zlsx_buffer_release.argtypes = [
        ctypes.POINTER(ctypes.c_ubyte),
        ctypes.c_size_t,
    ]
    lib.zlsx_buffer_release.restype = None

_STRUCTURAL_EDIT_SYMBOLS = (
    "zlsx_editor_insert_row",
    "zlsx_editor_delete_row",
    "zlsx_editor_insert_column",
    "zlsx_editor_delete_column",
    "zlsx_editor_add_sheet",
    "zlsx_editor_rename_sheet",
    "zlsx_editor_delete_sheet",
    "zlsx_editor_rename_table_column",
)
_HAS_STRUCTURAL_EDITS = _HAS_DIAG_RELEASE and all(hasattr(lib, s) for s in _STRUCTURAL_EDIT_SYMBOLS)
if _HAS_STRUCTURAL_EDITS:
    _diag_tail = [ctypes.POINTER(DiagV1), ctypes.c_char_p, ctypes.c_size_t]
    for _sym in ("zlsx_editor_insert_row", "zlsx_editor_delete_row",
                 "zlsx_editor_insert_column", "zlsx_editor_delete_column"):
        _fn = getattr(lib, _sym)
        _fn.argtypes = [editor_handle, ctypes.c_uint32, ctypes.c_uint32] + _diag_tail
        _fn.restype = ctypes.c_int32

    lib.zlsx_editor_add_sheet.argtypes = [
        editor_handle,
        ctypes.POINTER(ctypes.c_ubyte),  # name
        ctypes.c_size_t,                 # name_len
        ctypes.POINTER(ctypes.c_uint32), # out_sheet_idx (nullable)
    ] + _diag_tail
    lib.zlsx_editor_add_sheet.restype = ctypes.c_int32

    lib.zlsx_editor_rename_sheet.argtypes = [
        editor_handle,
        ctypes.c_uint32,                 # sheet_idx
        ctypes.POINTER(ctypes.c_ubyte),  # name
        ctypes.c_size_t,                 # name_len
    ] + _diag_tail
    lib.zlsx_editor_rename_sheet.restype = ctypes.c_int32

    lib.zlsx_editor_delete_sheet.argtypes = [editor_handle, ctypes.c_uint32] + _diag_tail
    lib.zlsx_editor_delete_sheet.restype = ctypes.c_int32

    lib.zlsx_editor_rename_table_column.argtypes = [
        editor_handle,
        ctypes.POINTER(ctypes.c_ubyte), ctypes.c_size_t,  # table_name
        ctypes.POINTER(ctypes.c_ubyte), ctypes.c_size_t,  # old_name
        ctypes.POINTER(ctypes.c_ubyte), ctypes.c_size_t,  # new_name
    ] + _diag_tail
    lib.zlsx_editor_rename_table_column.restype = ctypes.c_int32

_HAS_PIVOTS = hasattr(lib, "zlsx_editor_pivots_ndjson") and _HAS_BUFFER_RELEASE and _HAS_DIAG_RELEASE
if _HAS_PIVOTS:
    lib.zlsx_editor_pivots_ndjson.argtypes = [
        editor_handle,
        ctypes.POINTER(ctypes.POINTER(ctypes.c_ubyte)),  # out
        ctypes.POINTER(ctypes.c_size_t),                 # out_len
        ctypes.POINTER(DiagV1),
        ctypes.c_char_p,
        ctypes.c_size_t,
    ]
    lib.zlsx_editor_pivots_ndjson.restype = ctypes.c_int32

# S3b slice 2: the defined-names read, the pivots buffer contract.
# Plural, the read — `_HAS_DEFINED_NAME` (singular) probes the writer's
# add_defined_name and predates this capability.
_HAS_DEFINED_NAMES = hasattr(lib, "zlsx_editor_defined_names_ndjson") and _HAS_BUFFER_RELEASE and _HAS_DIAG_RELEASE
if _HAS_DEFINED_NAMES:
    lib.zlsx_editor_defined_names_ndjson.argtypes = [
        editor_handle,
        ctypes.POINTER(ctypes.POINTER(ctypes.c_ubyte)),  # out
        ctypes.POINTER(ctypes.c_size_t),                 # out_len
        ctypes.POINTER(DiagV1),
        ctypes.c_char_p,
        ctypes.c_size_t,
    ]
    lib.zlsx_editor_defined_names_ndjson.restype = ctypes.c_int32

# S3b slice 6: the conditional-formats read, the same buffer contract.
# Plural + `_FORMATS`, the read — the writer's add_conditional_format_*
# probes are per-method and predate this capability.
_HAS_CONDITIONAL_FORMATS = hasattr(lib, "zlsx_editor_conditional_formats_ndjson") and _HAS_BUFFER_RELEASE and _HAS_DIAG_RELEASE
if _HAS_CONDITIONAL_FORMATS:
    lib.zlsx_editor_conditional_formats_ndjson.argtypes = [
        editor_handle,
        ctypes.POINTER(ctypes.POINTER(ctypes.c_ubyte)),  # out
        ctypes.POINTER(ctypes.c_size_t),                 # out_len
        ctypes.POINTER(DiagV1),
        ctypes.c_char_p,
        ctypes.c_size_t,
    ]
    lib.zlsx_editor_conditional_formats_ndjson.restype = ctypes.c_int32

# S3b slice 7: the anchors read, the same buffer contract.
_HAS_ANCHORS = hasattr(lib, "zlsx_editor_anchors_ndjson") and _HAS_BUFFER_RELEASE and _HAS_DIAG_RELEASE
if _HAS_ANCHORS:
    lib.zlsx_editor_anchors_ndjson.argtypes = [
        editor_handle,
        ctypes.POINTER(ctypes.POINTER(ctypes.c_ubyte)),  # out
        ctypes.POINTER(ctypes.c_size_t),                 # out_len
        ctypes.POINTER(DiagV1),
        ctypes.c_char_p,
        ctypes.c_size_t,
    ]
    lib.zlsx_editor_anchors_ndjson.restype = ctypes.c_int32

# S3b slice 9: the sheet-props and calc-props reads, the same buffer
# contract. One probe for the pair — the header's ZLSX_HAS_SHEET_PROPS
# covers both exports, so a dylib has either both or neither.
_HAS_SHEET_PROPS = (
    hasattr(lib, "zlsx_editor_sheet_props_ndjson")
    and hasattr(lib, "zlsx_editor_calc_props_ndjson")
    and _HAS_BUFFER_RELEASE
    and _HAS_DIAG_RELEASE
)
if _HAS_SHEET_PROPS:
    for _fn in (lib.zlsx_editor_sheet_props_ndjson, lib.zlsx_editor_calc_props_ndjson):
        _fn.argtypes = [
            editor_handle,
            ctypes.POINTER(ctypes.POINTER(ctypes.c_ubyte)),  # out
            ctypes.POINTER(ctypes.c_size_t),                 # out_len
            ctypes.POINTER(DiagV1),
            ctypes.c_char_p,
            ctypes.c_size_t,
        ]
        _fn.restype = ctypes.c_int32
    del _fn
