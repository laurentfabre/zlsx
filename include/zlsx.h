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

/* Stage-2 style fields (added 0.2.4, ABI v1 compatible).
 *
 * `flags` uses bit 0 = font_size_set, bit 1 = font_color_set so callers
 * can distinguish "unset (default)" from explicitly-0 values. Alignment
 * uses a compact enum (0 = general / 1 = left / 2 = center / 3 = right /
 * 4 = fill / 5 = justify / 6 = centerContinuous / 7 = distributed).
 * Unknown alignment values return -1 with err="BadAlignmentValue". */
typedef struct {
    uint8_t         font_bold;            /* 0 or 1 */
    uint8_t         font_italic;          /* 0 or 1 */
    uint8_t         alignment_horizontal; /* 0..7 */
    uint8_t         wrap_text;            /* 0 or 1 */
    uint8_t         flags;
    uint8_t         _pad0[3];
    float           font_size;            /* used iff flags & 0x01 */
    uint32_t        font_color_argb;      /* used iff flags & 0x02 */
    const uint8_t * font_name_ptr;        /* NULL or unused iff font_name_len == 0 */
    size_t          font_name_len;
} zlsx_style_t;

#define ZLSX_FONT_SIZE_SET  0x01u
#define ZLSX_FONT_COLOR_SET 0x02u

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

#ifdef __cplusplus
}
#endif

#endif /* ZLSX_H */
