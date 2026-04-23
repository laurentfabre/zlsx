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


class CDxf(ctypes.Structure):
    _fields_ = [
        ("bold", ctypes.c_uint8),
        ("italic", ctypes.c_uint8),
        ("has_color", ctypes.c_uint8),
        ("has_fill", ctypes.c_uint8),
        ("color_argb", ctypes.c_uint32),
        ("fill_fg_argb", ctypes.c_uint32),
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
