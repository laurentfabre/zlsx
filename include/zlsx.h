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

#ifdef __cplusplus
}
#endif

#endif /* ZLSX_H */
