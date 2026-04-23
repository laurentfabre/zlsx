/*
 * zlsx.h — C header for the zlsx xlsx reader.
 *
 * ABI contract
 * ------------
 * Opaque handles (zlsx_book_t*, zlsx_rows_t*) are allocated and freed
 * by this library. Callers must close them to release resources.
 *
 * Distinct handles are independent; operations on the SAME handle must
 * be externally synchronized — in particular, do not call zlsx_book_close()
 * concurrently with any other call taking the same handle. This matches
 * the sqlite3 / libcurl convention.
 *
 * The internal refcount lets a zlsx_rows_t* returned by zlsx_rows_open()
 * safely outlive the caller's zlsx_book_t* handle; the last close on
 * either frees the underlying state.
 *
 * Strings inside zlsx_cell_t.str_ptr point into buffers owned by the
 * Book (for SST-backed strings) or a short-lived per-row scratch (for
 * inline strings with entity decoding). They are valid only until the
 * next zlsx_rows_next() call or until either handle is closed. Copy
 * them if you need to outlive that window.
 *
 * Stability: bumps to ZLSX_ABI_VERSION signal binary-incompatible
 * changes. Additive changes leave the version untouched.
 */

#ifndef ZLSX_H
#define ZLSX_H

#include <stddef.h>
#include <stdint.h>

#ifdef __cplusplus
extern "C" {
#endif

/* ABI version — bumps on any binary-incompatible change. */
#define ZLSX_ABI_VERSION 1u

/* Opaque handles. Never dereference the struct contents directly. */
typedef struct zlsx_book_t zlsx_book_t;
typedef struct zlsx_rows_t zlsx_rows_t;
typedef struct zlsx_writer_t zlsx_writer_t;
typedef struct zlsx_sheet_writer_t zlsx_sheet_writer_t;

/* Cell tag discriminator. */
typedef enum {
    ZLSX_CELL_EMPTY   = 0,
    ZLSX_CELL_STRING  = 1,
    ZLSX_CELL_INTEGER = 2,
    ZLSX_CELL_NUMBER  = 3,
    ZLSX_CELL_BOOLEAN = 4
} zlsx_cell_tag_t;

/*
 * Flat cell struct — all fields present regardless of tag; interpret
 * based on `tag`:
 *   ZLSX_CELL_EMPTY    → ignore every other field
 *   ZLSX_CELL_STRING   → str_ptr, str_len
 *   ZLSX_CELL_INTEGER  → i
 *   ZLSX_CELL_NUMBER   → f
 *   ZLSX_CELL_BOOLEAN  → b (0 or 1)
 */
typedef struct {
    uint32_t        tag;      /* zlsx_cell_tag_t */
    uint32_t        str_len;
    const uint8_t * str_ptr;
    int64_t         i;
    double          f;
    uint8_t         b;
    uint8_t         _pad[7];
} zlsx_cell_t;

/* ABI version + build-time version string. */
uint32_t     zlsx_abi_version(void);
const char * zlsx_version_string(void);

/*
 * Open an xlsx file. On failure returns NULL and, if err_buf is non-NULL
 * with err_buf_len > 0, writes a null-terminated diagnostic into err_buf
 * (truncated to err_buf_len - 1 bytes).
 */
zlsx_book_t * zlsx_book_open(const char * path,
                             uint8_t     * err_buf,
                             size_t        err_buf_len);

/* Drop the caller's reference to a Book. NULL-safe (no-op). Active
 * row iterators hold their own references, so calling this while rows
 * are live is safe — the state is freed on the last reference. */
void zlsx_book_close(zlsx_book_t * book);

/* Number of sheets in the workbook. */
uint32_t zlsx_sheet_count(zlsx_book_t * book);

/*
 * Copy sheet `idx`'s name into out_buf, null-terminated. Returns the
 * full name length (may exceed out_buf_len - 1 — re-query with a
 * larger buffer if the return value is >= out_buf_len). Returns 0 if
 * idx is out of range.
 */
size_t zlsx_sheet_name(zlsx_book_t * book,
                       uint32_t      idx,
                       uint8_t     * out_buf,
                       size_t        out_buf_len);

/*
 * Find a sheet by name. Returns the 0-based index, or -1 if not found.
 * `name_ptr` does not need to be null-terminated; `name_len` bytes are
 * compared byte-for-byte against each sheet's declared name.
 */
int32_t zlsx_sheet_index_by_name(zlsx_book_t * book,
                                 const uint8_t * name_ptr,
                                 size_t          name_len);

/*
 * Merged cell range for a sheet. Columns are 0-based (A=0),
 * rows are 1-based (row1=1) — matches the Zig/Python API.
 */
typedef struct {
    uint32_t top_left_col;
    uint32_t top_left_row;
    uint32_t bottom_right_col;
    uint32_t bottom_right_row;
} zlsx_merge_range_t;

/*
 * Number of merged cell ranges on sheet `sheet_idx`. Returns 0 if
 * the index is out of range or the sheet has no merges.
 */
size_t zlsx_merged_range_count(zlsx_book_t * book, uint32_t sheet_idx);

/*
 * Copy merged range `range_idx` on sheet `sheet_idx` into `out`.
 * Returns 0 on success, -1 if either index is out of range.
 */
int32_t zlsx_merged_range_at(zlsx_book_t *        book,
                             uint32_t             sheet_idx,
                             size_t               range_idx,
                             zlsx_merge_range_t * out);

/*
 * Hyperlink entry. `url_ptr` points into the Book's rels XML and is
 * valid until `zlsx_book_close`; XML entities like `&amp;` are
 * preserved (URL round-trips byte-for-byte through save/reopen).
 */
typedef struct {
    uint32_t        top_left_col;
    uint32_t        top_left_row;
    uint32_t        bottom_right_col;
    uint32_t        bottom_right_row;
    const uint8_t * url_ptr;
    size_t          url_len;
} zlsx_hyperlink_t;

/*
 * Number of hyperlinks on sheet `sheet_idx`. Returns 0 if the index
 * is out of range or the sheet has none.
 */
size_t zlsx_hyperlink_count(zlsx_book_t * book, uint32_t sheet_idx);

/*
 * Copy hyperlink `link_idx` on sheet `sheet_idx` into `out`. Returns
 * 0 on success, -1 if either index is out of range.
 */
int32_t zlsx_hyperlink_at(zlsx_book_t *      book,
                          uint32_t           sheet_idx,
                          size_t             link_idx,
                          zlsx_hyperlink_t * out);

/*
 * Open a row iterator for sheet `sheet_idx`. On failure returns NULL
 * and writes a diagnostic into err_buf as per zlsx_book_open().
 *
 * The returned iterator retains a reference on the book, so it is safe
 * to close `book` while `rows` is still live — the underlying state
 * stays alive until the last reference is dropped.
 */
zlsx_rows_t * zlsx_rows_open(zlsx_book_t * book,
                             uint32_t      sheet_idx,
                             uint8_t     * err_buf,
                             size_t        err_buf_len);

/* Close and free a Rows handle. NULL-safe. Drops the reference on the
 * underlying Book; if this was the last reference, the Book is freed. */
void zlsx_rows_close(zlsx_rows_t * rows);

/*
 * Advance to the next row. On return:
 *    1 → a row is available; *out_cells points to an array of *out_len
 *        cells, valid until the next call to zlsx_rows_next() or until
 *        zlsx_rows_close() / zlsx_book_close() frees the underlying
 *        state. String pointers inside the cells have the same
 *        lifetime — copy them if you need to outlive the row.
 *    0 → end of sheet.
 *   -1 → parse error; if err_buf is non-NULL, writes a diagnostic.
 */
int32_t zlsx_rows_next(zlsx_rows_t         * rows,
                       const zlsx_cell_t ** out_cells,
                       size_t             * out_len,
                       uint8_t            * err_buf,
                       size_t               err_buf_len);

/* ─── Writer (ABI v1, added in 0.2.2) ─────────────────────────────── */

/*
 * Create a new empty Writer. Returns NULL on allocation failure; err_buf
 * receives a null-terminated diagnostic.
 */
zlsx_writer_t * zlsx_writer_create(uint8_t * err_buf, size_t err_buf_len);

/*
 * Release all Writer state. Any zlsx_sheet_writer_t handles obtained
 * from this Writer become invalid immediately — do not use them after
 * closing the parent. NULL-safe (no-op).
 */
void zlsx_writer_close(zlsx_writer_t * writer);

/*
 * Add a sheet. The returned sheet-writer handle is BORROWED from the
 * parent Writer — do not close it explicitly; it becomes invalid when
 * the Writer is closed. `name_ptr` does not need to be null-terminated.
 * Returns NULL on error.
 */
zlsx_sheet_writer_t * zlsx_writer_add_sheet(
    zlsx_writer_t * writer,
    const uint8_t * name_ptr,
    size_t          name_len,
    uint8_t       * err_buf,
    size_t          err_buf_len);

/*
 * Append a row of cells. Each `zlsx_cell_t` is interpreted exactly the
 * same way as on the read side — fill `tag` plus the field(s) matching
 * that tag. Integers outside ±2^53-significant-bits round on open in
 * Excel; the writer rejects those up front with err="IntegerExceedsExcelPrecision".
 *
 * On success returns 0 and the row is appended. On failure returns -1;
 * the row buffer is unchanged (the validation pass runs before any
 * mutation), so callers may retry / skip and keep writing.
 *
 * `cells_ptr` may be NULL iff `cells_len == 0` (emit an empty row).
 */
int32_t zlsx_sheet_writer_write_row(
    zlsx_sheet_writer_t * sw,
    const zlsx_cell_t   * cells_ptr,
    size_t                cells_len,
    uint8_t             * err_buf,
    size_t                err_buf_len);

/*
 * Serialise the in-memory workbook and write it to `path` (the path
 * does not need to be null-terminated; `path_len` bytes are used).
 * Returns 0 on success, -1 on failure. The Writer remains usable —
 * further rows may be appended and save() called again.
 */
int32_t zlsx_writer_save(
    zlsx_writer_t * writer,
    const uint8_t * path_ptr,
    size_t          path_len,
    uint8_t       * err_buf,
    size_t          err_buf_len);

/* ─── Styles (Phase 3b, added in 0.2.4) ──────────────────────────── */

/*
 * Register a cell style. Style index (1-based) is written into
 * *out_index. Returns 0 on success, -1 on allocation failure.
 *
 * Dedup: registering the same { font_bold, font_italic } combination
 * twice returns the same index. Style 0 is always the default no-style
 * slot reserved by the library.
 *
 * Future Style fields will be exposed through `_ex` variants to keep
 * this ABI stable.
 */
int32_t zlsx_writer_add_style(
    zlsx_writer_t * writer,
    uint8_t         font_bold,     /* 0 or 1 */
    uint8_t         font_italic,   /* 0 or 1 */
    uint32_t      * out_index,
    uint8_t       * err_buf,
    size_t          err_buf_len);

/* Stage 2-4 style fields (added 0.2.4, ABI v1 additive).
 *
 * `flags` (stage 1-3) + `flags2` (stage 4) let callers distinguish
 * "unset (default)" from explicitly-0 values for fields where C has
 * no natural Option<> type:
 *
 *   flags  bit 0  — font_size
 *   flags  bit 1  — font_color
 *   flags  bit 2  — fill_fg_argb
 *   flags  bit 3  — fill_bg_argb
 *   flags2 bit 0  — border_left_color_argb
 *   flags2 bit 1  — border_right_color_argb
 *   flags2 bit 2  — border_top_color_argb
 *   flags2 bit 3  — border_bottom_color_argb
 *   flags2 bit 4  — border_diagonal_color_argb
 *
 * `alignment_horizontal` enum:
 *   0=general, 1=left, 2=center, 3=right, 4=fill, 5=justify,
 *   6=centerContinuous, 7=distributed.
 * `fill_pattern` enum: 0=none, 1=solid, 2=gray125, 3=gray0625,
 *   4=darkGray, 5=mediumGray, 6=lightGray, 7..=12 dark*, 13..=18 light*.
 * `border_*_style` enum: 0=none, 1=thin, 2=medium, 3=dashed, 4=dotted,
 *   5=thick, 6=double, 7=hair, 8=mediumDashed, 9=dashDot,
 *   10=mediumDashDot, 11=dashDotDot, 12=mediumDashDotDot,
 *   13=slantDashDot.
 * Unknown enum values return -1 with err="BadAlignmentValue",
 * "BadFillPattern", or "BadBorderStyle". */
typedef struct {
    uint8_t         font_bold;            /* 0 or 1 */
    uint8_t         font_italic;          /* 0 or 1 */
    uint8_t         alignment_horizontal; /* 0..7 */
    uint8_t         wrap_text;            /* 0 or 1 */
    uint8_t         flags;
    uint8_t         fill_pattern;         /* 0..=18 */
    uint8_t         flags2;               /* stage-4 flag bits */
    uint8_t         _pad0[1];
    float           font_size;            /* used iff flags & 0x01 */
    uint32_t        font_color_argb;      /* used iff flags & 0x02 */
    uint32_t        fill_fg_argb;         /* used iff flags & 0x04 */
    uint32_t        fill_bg_argb;         /* used iff flags & 0x08 */
    uint8_t         border_left_style;
    uint8_t         border_right_style;
    uint8_t         border_top_style;
    uint8_t         border_bottom_style;
    uint8_t         border_diagonal_style;
    uint8_t         diagonal_up;          /* 0 or 1 */
    uint8_t         diagonal_down;        /* 0 or 1 */
    uint8_t         _pad1[1];
    uint32_t        border_left_color_argb;
    uint32_t        border_right_color_argb;
    uint32_t        border_top_color_argb;
    uint32_t        border_bottom_color_argb;
    uint32_t        border_diagonal_color_argb;
    const uint8_t * font_name_ptr;        /* NULL or unused iff font_name_len == 0 */
    size_t          font_name_len;
    /* Stage 5: OOXML number-format string (e.g., "0.00", "m/d/yyyy"). */
    const uint8_t * num_fmt_ptr;          /* NULL or unused iff num_fmt_len == 0 */
    size_t          num_fmt_len;
} zlsx_style_t;

#define ZLSX_FONT_SIZE_SET              0x01u
#define ZLSX_FONT_COLOR_SET             0x02u
#define ZLSX_FILL_FG_SET                0x04u
#define ZLSX_FILL_BG_SET                0x08u
#define ZLSX_BORDER_LEFT_COLOR_SET      0x01u /* flags2 bit 0 */
#define ZLSX_BORDER_RIGHT_COLOR_SET     0x02u /* flags2 bit 1 */
#define ZLSX_BORDER_TOP_COLOR_SET       0x04u /* flags2 bit 2 */
#define ZLSX_BORDER_BOTTOM_COLOR_SET    0x08u /* flags2 bit 3 */
#define ZLSX_BORDER_DIAGONAL_COLOR_SET  0x10u /* flags2 bit 4 */

int32_t zlsx_writer_add_style_ex(
    zlsx_writer_t      * writer,
    const zlsx_style_t * spec,
    uint32_t           * out_index,
    uint8_t            * err_buf,
    size_t               err_buf_len);

/*
 * Write a row with per-cell style indices. `styles_ptr` must point at
 * an array of `cells_len` uint32_t values; use 0 for cells that should
 * inherit the default (no-style) formatting.
 *
 * Atomicity contract is identical to zlsx_sheet_writer_write_row:
 * integer-precision validation runs before any mutation, so a failed
 * write leaves the sheet buffer unchanged and the caller can skip /
 * retry the row.
 *
 * Returns 0 on success, -1 on failure.
 */
int32_t zlsx_sheet_writer_write_row_styled(
    zlsx_sheet_writer_t * sw,
    const zlsx_cell_t   * cells_ptr,
    const uint32_t      * styles_ptr,
    size_t                cells_len,
    uint8_t             * err_buf,
    size_t                err_buf_len);

/* Stage 5: per-sheet layout features (added 0.2.4). */

/* Set the display width of column `col_idx` (0-based, A=0) in the
 * spreadsheet "character unit" that Excel uses. Returns 0 on success
 * or -1 with err="InvalidColumnWidth" for non-finite or non-positive
 * values. */
int32_t zlsx_sheet_writer_set_column_width(
    zlsx_sheet_writer_t * sw,
    uint32_t              col_idx,
    float                 width,
    uint8_t             * err_buf,
    size_t                err_buf_len);

/* Freeze the top `rows` rows and left `cols` columns on the sheet.
 * Pass 0 on an axis to leave it unfrozen. Overrides any previous
 * freeze on this sheet. Never fails. */
void zlsx_sheet_writer_freeze_panes(
    zlsx_sheet_writer_t * sw,
    uint32_t              rows,
    uint32_t              cols);

/* Apply an auto-filter over an A1-style range (e.g. "A1:E1"). The
 * writer dupes the range string immediately. Returns 0 on success or
 * -1 with err="InvalidAutoFilterRange" on empty input. */
int32_t zlsx_sheet_writer_set_auto_filter(
    zlsx_sheet_writer_t * sw,
    const uint8_t       * range_ptr,
    size_t                range_len,
    uint8_t             * err_buf,
    size_t                err_buf_len);

/* Register a rectangular merged cell range (A1-style, e.g. "A1:B2").
 * Validated + duped by the writer on intake. Returns 0 on success or
 * -1 with err="InvalidMergeRange" on empty / single-cell / inverted /
 * out-of-Excel-range input. Multiple merges per sheet are allowed;
 * callers must avoid overlaps (Excel rejects overlapping pairs at
 * file-open time). */
int32_t zlsx_sheet_writer_add_merged_cell(
    zlsx_sheet_writer_t * sw,
    const uint8_t       * range_ptr,
    size_t                range_len,
    uint8_t             * err_buf,
    size_t                err_buf_len);

/* Attach an external-URL hyperlink to a cell or rectangular range.
 * `range` is A1-style (single cell "A1" or span "B2:C3"); `url` is
 * the external target (http/https/mailto/file/...). Returns 0 or
 * -1 with err="InvalidHyperlinkRange" / "InvalidHyperlinkUrl". URL
 * is xml-escaped on emit so query-string `&` is safe. */
int32_t zlsx_sheet_writer_add_hyperlink(
    zlsx_sheet_writer_t * sw,
    const uint8_t       * range_ptr,
    size_t                range_len,
    const uint8_t       * url_ptr,
    size_t                url_len,
    uint8_t             * err_buf,
    size_t                err_buf_len);

#ifdef __cplusplus
}
#endif

#endif /* ZLSX_H */
