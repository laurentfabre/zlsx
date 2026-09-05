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
typedef struct zlsx_matrix_t zlsx_matrix_t;
typedef struct zlsx_writer_t zlsx_writer_t;
typedef struct zlsx_sheet_writer_t zlsx_sheet_writer_t;
typedef struct zlsx_editor_t zlsx_editor_t;
typedef struct zlsx_emb_t zlsx_emb_t;

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

/*
 * Open an xlsx workbook from bytes already in memory. Same semantics as
 * zlsx_book_open, but no filesystem access: the buffer is parsed eagerly
 * and borrowed only for the duration of this call — the caller may free
 * `data` immediately after zlsx_book_open_buffer returns.
 *
 * For callers that receive workbook bytes without a path: SQL UDFs over
 * binary columns, network payloads, archives-within-archives.
 */
zlsx_book_t * zlsx_book_open_buffer(const uint8_t * data,
                                    size_t          len,
                                    uint8_t       * err_buf,
                                    size_t          err_buf_len);

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
 * Sheet visibility — the <sheet state="…"> attribute of sheet `idx`
 * as the reader modelled it: ZLSX_SHEET_STATE_VISIBLE, _HIDDEN
 * (Excel's Hide) or _VERY_HIDDEN (unreachable from Excel's UI — only
 * VBA / the object model reveals such a sheet, so a caller scanning a
 * workbook has no other way to learn it exists). A missing or
 * unrecognised `state` reads as visible, the schema default — the
 * same rule `zlsx list-sheets` prints, from the same field. Hidden
 * sheets stay in the inventory: zlsx_sheet_count / zlsx_sheet_name /
 * the row iterators enumerate them regardless. Returns -1 if idx is
 * out of range. 0.9.0+ (probe: ZLSX_HAS_SHEET_STATE).
 */
#define ZLSX_SHEET_STATE_VISIBLE      0
#define ZLSX_SHEET_STATE_HIDDEN       1
#define ZLSX_SHEET_STATE_VERY_HIDDEN  2
int32_t zlsx_sheet_state(zlsx_book_t * book, uint32_t idx);

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
 * Copy the internal `location` (e.g. "Sheet2!A1") of hyperlink
 * `link_idx` on sheet `sheet_idx` into `out_ptr` / `out_len`. Pointer
 * lifetime matches the Book. External hyperlinks return 0 with
 * `*out_len = 0`. Returns -1 on out-of-range indices. Surfaces the
 * destination that `zlsx_hyperlink_at` discards for internal links.
 */
int32_t zlsx_hyperlink_location_at(zlsx_book_t *           book,
                                   uint32_t                sheet_idx,
                                   size_t                  link_idx,
                                   const unsigned char * * out_ptr,
                                   size_t *                out_len);

/*
 * Cell comment parsed from xl/comments*.xml. Author / text slices
 * point into the Book's internal arena; valid until
 * zlsx_book_close(). Comment bodies that use rich-text runs come
 * back as the concatenated plain text (rich-comment surface can be
 * added in a follow-up without breaking this struct).
 */
typedef struct {
    uint32_t        cell_col;
    uint32_t        cell_row;
    size_t          author_len;
    const uint8_t * author_ptr;
    size_t          text_len;
    const uint8_t * text_ptr;
} zlsx_comment_t;

/* Number of comments on sheet `sheet_idx`. Returns 0 on out-of-range
 * or no-comments. */
size_t zlsx_comment_count(zlsx_book_t * book, uint32_t sheet_idx);

/* Copy comment `comment_idx` on sheet `sheet_idx`. Returns 0 on
 * success, -1 on out-of-range indices. */
int32_t zlsx_comment_at(zlsx_book_t *    book,
                        uint32_t         sheet_idx,
                        size_t           comment_idx,
                        zlsx_comment_t * out);

/*
 * Number of rich-text runs for a comment. Returns 0 for plain-text
 * comments (the common case — zero overhead for callers that only
 * want `text` from zlsx_comment_at). Pair with zlsx_comment_run_at
 * to enumerate formatted runs.
 */
size_t zlsx_comment_run_count(zlsx_book_t * book,
                              uint32_t      sheet_idx,
                              size_t        comment_idx);

/*
 * Copy run `run_idx` of comment `comment_idx`. Returns 0 on success
 * with text + bold/italic populated; -1 on any out-of-range index
 * (including comments with no runs). Text pointer lifetime matches
 * the Book.
 */
int32_t zlsx_comment_run_at(zlsx_book_t *     book,
                            uint32_t          sheet_idx,
                            size_t            comment_idx,
                            size_t            run_idx,
                            const uint8_t * * out_text_ptr,
                            size_t *          out_text_len,
                            uint8_t *         out_bold,
                            uint8_t *         out_italic);

/*
 * Data-validation entry. `values_count` is the number of dropdown
 * options for type="list" validations; other variants still surface
 * the range with values_count=0. Values themselves are pulled via
 * `zlsx_data_validation_value_at` since extern structs can't hold
 * slice-of-slices.
 */
typedef struct {
    uint32_t top_left_col;
    uint32_t top_left_row;
    uint32_t bottom_right_col;
    uint32_t bottom_right_row;
    size_t   values_count;
} zlsx_data_validation_t;

/*
 * Number of data validations on sheet `sheet_idx`. Returns 0 if the
 * index is out of range or the sheet has none.
 */
size_t zlsx_data_validation_count(zlsx_book_t * book, uint32_t sheet_idx);

/*
 * Copy data validation `dv_idx` into `out`. Returns 0 on success or
 * -1 if either index is out of range.
 */
int32_t zlsx_data_validation_at(zlsx_book_t *            book,
                                uint32_t                 sheet_idx,
                                size_t                   dv_idx,
                                zlsx_data_validation_t * out);

/*
 * Copy dropdown value `value_idx` of validation `dv_idx` on sheet
 * `sheet_idx` into `*out_ptr` / `*out_len`. The pointer is into the
 * Book's internal buffers and is valid until `zlsx_book_close`.
 * Returns 0 on success or -1 if any index is out of range.
 */
int32_t zlsx_data_validation_value_at(zlsx_book_t *     book,
                                      uint32_t          sheet_idx,
                                      size_t            dv_idx,
                                      size_t            value_idx,
                                      const uint8_t * * out_ptr,
                                      size_t *          out_len);

/*
 * Data validation kind codes returned by zlsx_data_validation_kind().
 * Stable numeric codes so callers can switch on them.
 */
#define ZLSX_DV_KIND_LIST         0u
#define ZLSX_DV_KIND_WHOLE        1u
#define ZLSX_DV_KIND_DECIMAL      2u
#define ZLSX_DV_KIND_DATE         3u
#define ZLSX_DV_KIND_TIME         4u
#define ZLSX_DV_KIND_TEXT_LENGTH  5u
#define ZLSX_DV_KIND_CUSTOM       6u
#define ZLSX_DV_KIND_UNKNOWN      7u

/*
 * Data validation operator codes returned by
 * zlsx_data_validation_operator(). ZLSX_DV_OP_NONE means the source
 * had no `operator=` attribute (list / custom validations, or numeric
 * with an omitted operator — Excel treats the latter as `between`
 * but we preserve the absence so round-trips are exact).
 */
#define ZLSX_DV_OP_BETWEEN                  0u
#define ZLSX_DV_OP_NOT_BETWEEN              1u
#define ZLSX_DV_OP_EQUAL                    2u
#define ZLSX_DV_OP_NOT_EQUAL                3u
#define ZLSX_DV_OP_LESS_THAN                4u
#define ZLSX_DV_OP_LESS_THAN_OR_EQUAL       5u
#define ZLSX_DV_OP_GREATER_THAN             6u
#define ZLSX_DV_OP_GREATER_THAN_OR_EQUAL    7u
#define ZLSX_DV_OP_NONE                     0xFFFFFFFFu

/*
 * Return the kind code (see ZLSX_DV_KIND_*) for data validation
 * `dv_idx` on sheet `sheet_idx`. Returns ZLSX_DV_KIND_UNKNOWN on
 * out-of-range indices (callers should bounds-check via
 * zlsx_data_validation_count() first).
 */
uint32_t zlsx_data_validation_kind(zlsx_book_t * book,
                                   uint32_t      sheet_idx,
                                   size_t        dv_idx);

/*
 * Return the operator code (see ZLSX_DV_OP_*) for data validation
 * `dv_idx` on sheet `sheet_idx`. Returns ZLSX_DV_OP_NONE when the
 * source had no `operator=` attribute.
 */
uint32_t zlsx_data_validation_operator(zlsx_book_t * book,
                                       uint32_t      sheet_idx,
                                       size_t        dv_idx);

/*
 * Copy formula1 of data validation `dv_idx` on sheet `sheet_idx` into
 * `*out_ptr` / `*out_len`. The pointer is into the Book's internal
 * buffers and is valid until `zlsx_book_close`. Returns 0 on success,
 * -1 on out-of-range indices. An empty formula still returns 0 with
 * `*out_len = 0`.
 */
int32_t zlsx_data_validation_formula1(zlsx_book_t *     book,
                                      uint32_t          sheet_idx,
                                      size_t            dv_idx,
                                      const uint8_t * * out_ptr,
                                      size_t *          out_len);

/*
 * Copy formula2 of data validation `dv_idx` on sheet `sheet_idx`.
 * Same contract as zlsx_data_validation_formula1(); empty for
 * operators other than `between` / `not_between`.
 */
int32_t zlsx_data_validation_formula2(zlsx_book_t *     book,
                                      uint32_t          sheet_idx,
                                      size_t            dv_idx,
                                      const uint8_t * * out_ptr,
                                      size_t *          out_len);

/*
 * Total number of shared-string entries (0 when the workbook has
 * no xl/sharedStrings.xml part). Enumerate via zlsx_shared_string_at
 * together with zlsx_rich_run_count to discover which indices carry
 * rich-text runs.
 */
size_t zlsx_shared_string_count(zlsx_book_t * book);

/*
 * Copy SST entry `sst_idx` into `*out_ptr` / `*out_len`. Slice into
 * Book-owned storage; do not free. Returns 0 on success, -1 on
 * out-of-range.
 */
int32_t zlsx_shared_string_at(zlsx_book_t *     book,
                              size_t            sst_idx,
                              const uint8_t * * out_ptr,
                              size_t *          out_len);

/*
 * Number of rich-text runs for shared-string entry `sst_idx`, or 0
 * when that entry is a plain single-run string. Use this as a
 * presence probe before zlsx_rich_run_at(); SST entries without any
 * `<r>` wrappers in the source XML return 0.
 */
size_t zlsx_rich_run_count(zlsx_book_t * book, size_t sst_idx);

/*
 * Copy rich-text run `run_idx` of SST entry `sst_idx` into the out
 * pointers. Text is valid until zlsx_book_close(); bold/italic are
 * 0 or 1. Returns 0 on success, -1 on out-of-range indices.
 */
int32_t zlsx_rich_run_at(zlsx_book_t *     book,
                         size_t            sst_idx,
                         size_t            run_idx,
                         const uint8_t * * out_text_ptr,
                         size_t *          out_text_len,
                         uint8_t *         out_bold,
                         uint8_t *         out_italic);

/*
 * ARGB color of run `run_idx` on SST entry `sst_idx`. Returns 0 and
 * writes `*out_color` when the run had an explicit `<color rgb=…/>`.
 * Returns 1 when absent (no color, or a theme color we don't resolve);
 * `*out_color` is untouched. Returns -1 on out-of-range indices.
 */
int32_t zlsx_rich_run_color(zlsx_book_t * book,
                            size_t        sst_idx,
                            size_t        run_idx,
                            uint32_t *    out_color);

/*
 * Font size (points) of run `run_idx` on SST entry `sst_idx`. Same
 * present/absent/error tri-state as zlsx_rich_run_color.
 */
int32_t zlsx_rich_run_size(zlsx_book_t * book,
                           size_t        sst_idx,
                           size_t        run_idx,
                           float *       out_size);

/*
 * Font name of run `run_idx` on SST entry `sst_idx`. Pointer lifetime
 * matches the Book; empty slice (`*out_len == 0`) when the run had no
 * `<rFont val=…/>`. Returns 0 on success, -1 on out-of-range indices.
 */
int32_t zlsx_rich_run_font_name(zlsx_book_t *     book,
                                size_t            sst_idx,
                                size_t            run_idx,
                                const uint8_t * * out_ptr,
                                size_t *          out_len);

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
 * The row yielded by the last 1 is "the current row" for every
 * per-column getter on the handle (zlsx_rows_style_at(),
 * zlsx_rows_parse_date() and the formula / error getters); after a 0
 * or a -1, and after a zlsx_rows_skip() of n >= 1 whether it returned
 * 0 or -1, there is none and every getter returns -1 (a zero-length
 * skip is a no-op and keeps the row).
 */
int32_t zlsx_rows_next(zlsx_rows_t         * rows,
                       const zlsx_cell_t ** out_cells,
                       size_t             * out_len,
                       uint8_t            * err_buf,
                       size_t               err_buf_len);

/*
 * Advance past `n` rows without decoding their cells; writes the number
 * actually skipped to *out_skipped (fewer than `n` only at end of
 * sheet). Returns 0 on success, -1 on failure with a diagnostic in
 * err_buf.
 *
 * Semantically identical to calling zlsx_rows_next() `n` times and
 * discarding the results — same landing row, same row numbering — but
 * without building the cell arrays for what it passes. Intended for
 * range-partitioned reads, where each partition must first get past the
 * rows belonging to earlier partitions.
 *
 * Invalidates the cells of the most recently yielded row, exactly as
 * zlsx_rows_next() does — on -1 as well as on 0: there is no current
 * row afterwards either way. A zero-length skip (n == 0) is a no-op:
 * it writes 0 to *out_skipped and leaves the current row current.
 */
int32_t zlsx_rows_skip(zlsx_rows_t * rows,
                       size_t        n,
                       size_t      * out_skipped,
                       uint8_t     * err_buf,
                       size_t        err_buf_len);

/*
 * Style index for column `col_idx` of the most recently yielded row.
 * Valid between zlsx_rows_next() calls. Returns 0 and writes
 * `*out_style_idx` when the cell had an `s="…"` attribute; returns
 * 1 when the cell had no `s` (General / implicit style); returns -1
 * when `col_idx` is out of range for the current row.
 */
int32_t zlsx_rows_style_at(zlsx_rows_t * rows,
                           size_t        col_idx,
                           uint32_t    * out_style_idx);

/*
 * Formula text and error tags (S3b slice 11, 0.9.0+; probe:
 * ZLSX_HAS_ROWS_FORMULAS). The reader keeps three per-row side
 * channels beside the cells, so zlsx_cell_t and its tags are
 * untouched: a formula cell's zlsx_cell_t is its cached <v> value
 * (ZLSX_CELL_EMPTY for a formula-only cell) and an error cell's is
 * an ordinary ZLSX_CELL_STRING holding the literal — exactly what
 * every caller saw before. The three getters share
 * zlsx_rows_style_at's contract: 0 and the out params written when
 * the cell is that kind of cell; 1 when it is not (the out params
 * untouched); -1 when `col_idx` is out of range for the current row
 * — before the first zlsx_rows_next(), after its 0 / -1, after a
 * zlsx_rows_skip() of n >= 1, or past the row's end. Pointers are
 * valid until the next zlsx_rows_next() / a zlsx_rows_skip() of
 * n >= 1 or a close, the cells' own lifetime. For a formula cell
 * exactly one of zlsx_rows_formula_at()
 * and zlsx_rows_formula_ref_at() returns 0 and zlsx_rows_error_at()
 * returns 1 — a formula whose cached value is an error literal is a
 * formula (the CLI's `t:"formula"` with `cached:"#DIV/0!"`); for an
 * error cell only zlsx_rows_error_at() returns 0; for a value cell
 * all three return 1. The same rule `zlsx cells` prints as
 * `t:"formula"` (`formula` / `formula_ref` / `cached`) and
 * `t:"error"` (`v`), from the same fields.
 */

/*
 * Own formula text of column `col_idx` of the most recently yielded
 * row: the <f> body with XML entities decoded — a stand-alone
 * formula, a shared-formula base or an array-formula base. 1 for a
 * value cell, an error cell, or a shared / array slave (see
 * zlsx_rows_formula_ref_at()). An empty <f></f> returns 0 with
 * `*out_len = 0`.
 */
int32_t zlsx_rows_formula_at(zlsx_rows_t   *   rows,
                             size_t            col_idx,
                             const uint8_t * * out_ptr,
                             size_t *          out_len);

/*
 * Base cell of the shared- or array-formula slave at column
 * `col_idx` of the most recently yielded row — a cell with no <f>
 * body of its own (`<f t="shared" si="N"/>`, or a cell inside an
 * earlier `<f t="array" ref="…">` rectangle) whose formula is the
 * base's text. `*out_col` is 0-based (A = 0), `*out_row` 1-based —
 * the zlsx_merge_range_t convention. 1 when the cell is not a
 * slave; a slave whose base the reader never saw (an `si` with no
 * base above it in the sheet) reads as a value cell.
 */
int32_t zlsx_rows_formula_ref_at(zlsx_rows_t * rows,
                                 size_t        col_idx,
                                 uint32_t    * out_col,
                                 uint32_t    * out_row);

/*
 * Error literal of column `col_idx` of the most recently yielded
 * row: the <v> body of a `t="e"` cell (`#DIV/0!`, `#N/A`, `#REF!`,
 * `#VALUE!`, `#NUM!`, `#NAME?`, `#NULL!`, `#GETTING_DATA`) — the
 * bytes the cell array hands over as a ZLSX_CELL_STRING. 1 when the
 * cell is not an error cell, a formula with a cached error included.
 */
int32_t zlsx_rows_error_at(zlsx_rows_t   *   rows,
                           size_t            col_idx,
                           const uint8_t * * out_ptr,
                           size_t *          out_len);

/*
 * Bulk-FFI matrix surface (v0.2.8+). One zlsx_matrix_open() drains the
 * entire sheet into a flat zlsx_cell_t buffer + row-offsets array,
 * letting FFI consumers iterate the buffer in their own language with
 * zero further calls back into the C library. Intended for Python /
 * Node / etc. callers that pay per-call dispatch overhead on the
 * per-row zlsx_rows_next() loop at MB scale.
 *
 * Lifetime: out_cells / out_offsets buffers from zlsx_matrix_data()
 * stay valid until zlsx_matrix_close(). String slices inside cells
 * are duped into matrix-owned storage and have the same lifetime.
 */
zlsx_matrix_t * zlsx_matrix_open(zlsx_book_t * book,
                                 uint32_t      sheet_idx,
                                 uint8_t     * err_buf,
                                 size_t        err_buf_len);

void zlsx_matrix_close(zlsx_matrix_t * matrix);

/*
 * Read the matrix's flattened layout. After this call:
 *   *out_cells   points to the packed zlsx_cell_t buffer
 *   *out_offsets points to row-start offsets (length *out_n_rows + 1):
 *                row r runs cells[offsets[r] .. offsets[r+1]]
 *   *out_n_rows  is the row count
 * All three buffers stay valid until zlsx_matrix_close().
 */
void zlsx_matrix_data(zlsx_matrix_t      * matrix,
                      const zlsx_cell_t ** out_cells,
                      const size_t      ** out_offsets,
                      size_t             * out_n_rows);

/*
 * Decoded calendar date/time from an Excel-serial cell. Fields:
 *   year   — 1900..=9999
 *   month  — 1..=12, day 1..=31
 *   hour / minute / second — 0..=59 (23 for hour)
 *   _pad   — keep struct size/alignment predictable
 */
typedef struct {
    uint16_t year;
    uint8_t  month;
    uint8_t  day;
    uint8_t  hour;
    uint8_t  minute;
    uint8_t  second;
    uint8_t  _pad;
} zlsx_datetime_t;

/*
 * Parse the current-row cell at `col_idx` as a date-styled number.
 * Tri-state:
 *    0 → `*out` populated with the decoded DateTime
 *    1 → not a date (wrong type / non-date numFmt / out-of-range serial)
 *   -1 → `col_idx` is past the row width, or there is no current row
 *        (before the first zlsx_rows_next(), after its 0 / -1, after
 *        a zlsx_rows_skip() of n >= 1) — `*out` untouched
 *
 * Combines the existing `zlsx_rows_style_at` + `zlsx_is_date_format`
 * + `xlsx.fromExcelSerial` chain into one call.
 */
int32_t zlsx_rows_parse_date(zlsx_rows_t *     rows,
                             size_t            col_idx,
                             zlsx_datetime_t * out);

/*
 * Inverse of zlsx_rows_parse_date: DateTime → Excel serial number.
 * Returns 0 with `*out_serial` set on success, -1 when the input
 * is outside the round-trippable range (year<1900, invalid
 * month/day/hour/etc., or date ≤ 1900-02-29).
 *
 * Pair with a style carrying `number_format="yyyy-mm-dd"` (or any
 * date pattern) to write a date cell that Excel displays correctly
 * and zlsx_rows_parse_date decodes cleanly.
 */
int32_t zlsx_datetime_to_serial(const zlsx_datetime_t * dt,
                                double *                out_serial);

/*
 * Resolve a style index to its number-format code. Returns 0 and
 * populates `*out_ptr` / `*out_len` on success; returns -1 on
 * out-of-range indices or when the workbook has no styles.xml.
 * Pointer lifetime matches the Book.
 */
int32_t zlsx_number_format(zlsx_book_t *     book,
                           uint32_t          style_idx,
                           const uint8_t * * out_ptr,
                           size_t *          out_len);

/* Returns 1 if `style_idx` resolves to a date/time pattern, 0
 * otherwise (including out-of-range indices). */
uint8_t zlsx_is_date_format(zlsx_book_t * book, uint32_t style_idx);

/*
 * Per-cell font properties surfaced from xl/styles.xml `<fonts>`
 * indirected through `<cellXfs>`. `has_color` and `has_size` are
 * 0/1 flags — when 0, the respective field is meaningless (absent
 * in the source file). `name_ptr` / `name_len` borrow from the
 * Book's styles.xml; valid until zlsx_book_close().
 */
typedef struct {
    uint8_t         bold;
    uint8_t         italic;
    uint8_t         has_color;
    uint8_t         has_size;
    uint32_t        color_argb;
    float           size;
    size_t          name_len;
    const uint8_t * name_ptr;
} zlsx_cell_font_t;

/* Resolve a style index to its font. Returns 0 on success, -1 on
 * out-of-range indices or workbooks without styles.xml. */
int32_t zlsx_cell_font(zlsx_book_t *       book,
                       uint32_t            style_idx,
                       zlsx_cell_font_t *  out);

/*
 * Per-cell fill. `pattern_ptr` / `pattern_len` hold the OOXML
 * patternType attribute ("none", "solid", "darkDown", …). The
 * `has_fg` / `has_bg` flags indicate whether the ARGB fields are
 * populated; theme / indexed colors leave them at 0. Pointer lifetime
 * matches the Book.
 */
typedef struct {
    uint8_t         has_fg;
    uint8_t         has_bg;
    uint8_t         _pad[2];
    uint32_t        fg_color_argb;
    uint32_t        bg_color_argb;
    size_t          pattern_len;
    const uint8_t * pattern_ptr;
} zlsx_cell_fill_t;

/* Resolve a style index to its fill. Returns 0 on success, -1 on
 * out-of-range indices or workbooks without styles.xml. */
int32_t zlsx_cell_fill(zlsx_book_t *       book,
                       uint32_t            style_idx,
                       zlsx_cell_fill_t *  out);

/*
 * One side of a cell border. `style_ptr` / `style_len` hold the OOXML
 * style attribute ("thin", "medium", "thick", "double", "dashed", …)
 * or an empty slice when the side has no border. `has_color` + `pad`
 * keep the struct 4-byte aligned ahead of the u32 color.
 */
typedef struct {
    uint8_t         has_color;
    uint8_t         _pad[3];
    uint32_t        color_argb;
    size_t          style_len;
    const uint8_t * style_ptr;
} zlsx_border_side_t;

/*
 * Full cell border — five sides. Pointer lifetimes match the Book.
 */
typedef struct {
    zlsx_border_side_t left;
    zlsx_border_side_t right;
    zlsx_border_side_t top;
    zlsx_border_side_t bottom;
    zlsx_border_side_t diagonal;
} zlsx_cell_border_t;

/* Resolve a style index to its border. Returns 0 on success, -1 on
 * out-of-range indices or workbooks without styles.xml. Sides without
 * borders surface with `style_len == 0`. */
int32_t zlsx_cell_border(zlsx_book_t *          book,
                         uint32_t               style_idx,
                         zlsx_cell_border_t *   out);

/*
 * Cell alignment record. `horizontal_len == 0` means the alignment
 * is the OOXML default ("general", which the emitter omits).
 * `wrap_text == 1` when `wrapText="1"` was set on the <alignment>
 * child of the cell's <xf>. Pointer borrows from styles.xml; valid
 * for the lifetime of the Book.
 */
typedef struct {
    size_t          horizontal_len;
    const uint8_t * horizontal_ptr;
    uint8_t         wrap_text;
    uint8_t         _pad[7];
} zlsx_cell_alignment_t;

/* Resolve a style index to its alignment + wrap_text record.
 * Returns 0 on success, -1 on out-of-range index. Cells without a
 * nested <alignment> child surface as horizontal_len = 0,
 * wrap_text = 0. */
int32_t zlsx_cell_alignment(zlsx_book_t *           book,
                            uint32_t                style_idx,
                            zlsx_cell_alignment_t * out);

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
 * Rich-text run — one formatted piece of a rich-text cell.
 * `has_color` / `has_size` are 0/1 flags; when 0 the paired
 * `color_argb` / `size` field is ignored. `font_name_len == 0`
 * means "no rFont override". Text lifetime is the caller's — the
 * writer copies during zlsx_sheet_writer_write_rich_row().
 */
typedef struct {
    const uint8_t * text_ptr;
    size_t          text_len;
    uint8_t         bold;
    uint8_t         italic;
    uint8_t         has_color;
    uint8_t         has_size;
    uint32_t        color_argb;
    float           size;
    const uint8_t * font_name_ptr;
    size_t          font_name_len;
} zlsx_rich_run_t;

/*
 * Append a row mixing plain cells with rich-text cells. For each
 * column i in [0, cells_len):
 *   if rich_runs_lens[i] > 0 → rich cell; rich_runs_ptrs[i] points
 *     at rich_runs_lens[i] runs. cells_ptr[i] is ignored for that
 *     column (pass any placeholder).
 *   else → plain cell; cells_ptr[i] is a regular zlsx_cell_t.
 *
 * Either rich_runs_ptrs or rich_runs_lens may be NULL iff no
 * column is rich — passing both NULL degenerates to
 * zlsx_sheet_writer_write_row. Returns 0 on success, -1 on failure
 * (err_buf populated). Atomic: on failure next_row is not
 * advanced.
 */
int32_t zlsx_sheet_writer_write_rich_row(
    zlsx_sheet_writer_t         * sw,
    const zlsx_cell_t           * cells_ptr,
    const zlsx_rich_run_t * const* rich_runs_ptrs,
    const size_t                * rich_runs_lens,
    size_t                        cells_len,
    uint8_t                     * err_buf,
    size_t                        err_buf_len);

/*
 * Append a row that mixes plain value cells with formula cells. For
 * each column i in [0, cells_len):
 *   if formula_lens[i] > 0 → formula cell; formula_ptrs[i] points at
 *     formula_lens[i] bytes of formula text (e.g. "A1+B1"); cells_ptr[i]
 *     supplies the cached <v> value Excel shows until recalc (use the
 *     `empty` tag for "no cached value").
 *   else → plain value cell; cells_ptr[i] is the regular zlsx_cell_t.
 *
 * Both formula_ptrs and formula_lens may be NULL iff no column is a
 * formula — passing both NULL degenerates to zlsx_sheet_writer_write_row.
 * Returns 0 on success, -1 on failure (err_buf populated; "InvalidInput"
 * for caller-bug shapes like a non-zero formula_lens entry with NULL
 * formula_ptrs).
 */
int32_t zlsx_sheet_writer_write_row_with_formulas(
    zlsx_sheet_writer_t  * sw,
    const zlsx_cell_t    * cells_ptr,
    const uint8_t * const* formula_ptrs,
    const size_t         * formula_lens,
    size_t                 cells_len,
    uint8_t              * err_buf,
    size_t                 err_buf_len);

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

/*
 * Serialise the in-memory workbook into a freshly allocated buffer
 * instead of a file — the writer-side mirror of zlsx_book_open_buffer.
 * On success writes the base pointer to *out_ptr, the byte count to
 * *out_len, and returns 0; the caller owns those bytes and MUST release
 * them with zlsx_buffer_free(*out_ptr, *out_len).
 *
 * Returns -1 on failure with a diagnostic in `err_buf`, leaving
 * *out_ptr / *out_len untouched. Byte-for-byte identical to what
 * zlsx_writer_save would have written to disk. The Writer remains
 * usable and unmodified.
 */
int32_t zlsx_writer_save_to_buffer(
    zlsx_writer_t * writer,
    uint8_t      ** out_ptr,
    size_t        * out_len,
    uint8_t       * err_buf,
    size_t          err_buf_len);

/*
 * Release a buffer handed out by zlsx_writer_save_to_buffer. `len` must
 * be the exact length that call reported — the underlying allocator
 * frees by slice, not by base pointer alone. NULL is a no-op.
 */
void zlsx_buffer_free(
    uint8_t * ptr,
    size_t    len);

/* Register a workbook-level (or sheet-scoped) defined name.
 * `local_sheet_id_neg` < 0 means workbook-scope; >= 0 means
 * 0-based sheet index (must resolve at save() time).
 * `hidden_flag != 0` emits hidden="1". Returns 0 on success, -1
 * with err in {InvalidDefinedName, InvalidDefinedNameRefersTo,
 * DuplicateDefinedName, InvalidDefinedNameLocalSheetId,
 * OutOfMemory}. */
int32_t zlsx_writer_add_defined_name(
    zlsx_writer_t * writer,
    const uint8_t * name_ptr,
    size_t          name_len,
    const uint8_t * refers_to_ptr,
    size_t          refers_to_len,
    int32_t         local_sheet_id_neg,
    uint8_t         hidden_flag,
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

/* Set the height of row `row_idx` (0-based) in points. Excel
 * accepts heights in (0, 409.5]. Returns 0 on success or -1 with
 * err="InvalidRowHeight" / "RowOutOfRange". */
int32_t zlsx_sheet_writer_set_row_height(
    zlsx_sheet_writer_t * sw,
    uint32_t              row_idx,
    float                 height,
    uint8_t             * err_buf,
    size_t                err_buf_len);

/* Freeze the top `rows` rows and left `cols` columns on the sheet.
 * Pass 0 on an axis to leave it unfrozen. Overrides any previous
 * freeze on this sheet. Never fails — out-of-range counts are
 * clamped silently. Legacy entry point; new callers should prefer
 * zlsx_sheet_writer_freeze_panes_checked for typed errors. */
void zlsx_sheet_writer_freeze_panes(
    zlsx_sheet_writer_t * sw,
    uint32_t              rows,
    uint32_t              cols);

/* Checked freeze-panes: returns -1 with err="RowOutOfRange" /
 * "ColumnOutOfRange" instead of clamping. Newer FFI consumers
 * should prefer this over the legacy clamping form. */
int32_t zlsx_sheet_writer_freeze_panes_checked(
    zlsx_sheet_writer_t * sw,
    uint32_t              rows,
    uint32_t              cols,
    uint8_t             * err_buf,
    size_t                err_buf_len);

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

/* Attach a list-type data validation (dropdown) to a cell or range.
 * `range` is A1-style. `values_ptr[i][0..lens_ptr[i]]` is value i of
 * `values_count` dropdown options. Excel joins options with commas
 * inside a quoted formula1 string — commas or bare `"` in values
 * are rejected. Returns 0 or -1 with err="InvalidHyperlinkRange" /
 * "InvalidDataValidation". Max 256 values per call. */
int32_t zlsx_sheet_writer_add_data_validation_list(
    zlsx_sheet_writer_t  * sw,
    const uint8_t        * range_ptr,
    size_t                 range_len,
    const uint8_t * const* values_ptr,
    const size_t         * lens_ptr,
    size_t                 values_count,
    uint8_t              * err_buf,
    size_t                 err_buf_len);

/* Attach a numeric / date / time / text-length data validation to a
 * cell or rectangular range. `range` is A1-style. `kind_code` must be
 * one of ZLSX_DV_KIND_WHOLE / DECIMAL / DATE / TIME / TEXT_LENGTH
 * (using LIST / CUSTOM / UNKNOWN returns InvalidDataValidation —
 * those have dedicated entry points / aren't user-facing). `op_code`
 * must be one of ZLSX_DV_OP_* (never ZLSX_DV_OP_NONE). Pass
 * `formula2_ptr = NULL` with `formula2_len = 0` for single-formula
 * operators; non-NULL is required for BETWEEN / NOT_BETWEEN. */
int32_t zlsx_sheet_writer_add_data_validation_numeric(
    zlsx_sheet_writer_t * sw,
    const uint8_t       * range_ptr,
    size_t                range_len,
    uint32_t              kind_code,
    uint32_t              op_code,
    const uint8_t       * formula1_ptr,
    size_t                formula1_len,
    const uint8_t       * formula2_ptr,
    size_t                formula2_len,
    uint8_t             * err_buf,
    size_t                err_buf_len);

/* Attach a custom-formula data validation. Same error semantics as
 * zlsx_sheet_writer_add_data_validation_numeric() minus the operator
 * / formula2 (custom has neither). */
int32_t zlsx_sheet_writer_add_data_validation_custom(
    zlsx_sheet_writer_t * sw,
    const uint8_t       * range_ptr,
    size_t                range_len,
    const uint8_t       * formula_ptr,
    size_t                formula_len,
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

/* Attach an internal (same-workbook) hyperlink to a cell or range.
 * `location` is the target ref Excel writes verbatim into
 * <hyperlink location="…"/>, e.g. "Sheet2!A1" or "'Sheet With
 * Spaces'!B2". Returns 0 or -1 with err="InvalidHyperlinkRange" /
 * "InvalidHyperlinkLocation". */
int32_t zlsx_sheet_writer_add_internal_hyperlink(
    zlsx_sheet_writer_t * sw,
    const uint8_t       * range_ptr,
    size_t                range_len,
    const uint8_t       * location_ptr,
    size_t                location_len,
    uint8_t             * err_buf,
    size_t                err_buf_len);

/* Attach a cell comment (note). `ref` is a single-cell A1 ref
 * ("B2"); ranges are rejected. `author` + `text` are plain text,
 * xml-escaped on emit. Returns 0 or -1 with err="InvalidCommentRef"
 * / "InvalidHyperlinkRange" on bad ref, "OutOfMemory" on alloc. */
int32_t zlsx_sheet_writer_add_comment(
    zlsx_sheet_writer_t * sw,
    const uint8_t       * ref_ptr,
    size_t                ref_len,
    const uint8_t       * author_ptr,
    size_t                author_len,
    const uint8_t       * text_ptr,
    size_t                text_len,
    uint8_t             * err_buf,
    size_t                err_buf_len);

/*
 * Per-border-side payload for a dxf. `style` is a BorderStyle enum
 * byte (0 = none, 1 = thin, 2 = medium, 3 = dashed, …, 13 =
 * slantDashDot — see writer/BorderStyle source for the full table).
 * `has_color` + 3-byte pad align the u32 color ahead.
 */
typedef struct {
    uint8_t  style;
    uint8_t  has_color;
    uint8_t  _pad[2];
    uint32_t color_argb;
} zlsx_dxf_border_side_t;

/*
 * Differential format for conditional formatting — font / fill /
 * border overrides applied when a cfRule matches. Has-flags gate
 * the paired optional fields (0 means "not set"). `size` is in
 * points; borders default to .none (inherit cell style).
 */
typedef struct {
    uint8_t                bold;
    uint8_t                italic;
    uint8_t                has_color;
    uint8_t                has_fill;
    uint32_t               color_argb;
    uint32_t               fill_fg_argb;
    uint8_t                has_size;
    uint8_t                _pad[3];
    float                  size;
    zlsx_dxf_border_side_t border_left;
    zlsx_dxf_border_side_t border_right;
    zlsx_dxf_border_side_t border_top;
    zlsx_dxf_border_side_t border_bottom;
} zlsx_dxf_t;

/* Register a dxf on the workbook-wide `<dxfs>` table. Returns 0 on
 * success with `*out_dxf_id` set; -1 on alloc. Content-dedup'd. */
int32_t zlsx_writer_add_dxf(zlsx_writer_t *   w,
                            const zlsx_dxf_t* dxf,
                            uint32_t *        out_dxf_id,
                            uint8_t *         err_buf,
                            size_t            err_buf_len);

/* Attach a cellIs-type conditional-format rule. `op_code` reuses
 * the ZLSX_DV_OP_* codes (same OOXML tokens). `formula2_ptr` may
 * be NULL with formula2_len=0 when the operator doesn't need a
 * second formula (required for BETWEEN / NOT_BETWEEN). Returns 0
 * or -1 with err="InvalidDataValidation" / "InvalidHyperlinkRange"
 * / "UnknownDxfId". */
int32_t zlsx_sheet_writer_add_conditional_format_cell_is(
    zlsx_sheet_writer_t * sw,
    const uint8_t       * range_ptr,
    size_t                range_len,
    uint32_t              op_code,
    const uint8_t       * formula1_ptr,
    size_t                formula1_len,
    const uint8_t       * formula2_ptr,
    size_t                formula2_len,
    uint32_t              dxf_id,
    uint8_t             * err_buf,
    size_t                err_buf_len);

/* Attach an expression-type conditional-format rule. Same error
 * semantics as cellIs minus the operator / formula2. */
int32_t zlsx_sheet_writer_add_conditional_format_expression(
    zlsx_sheet_writer_t * sw,
    const uint8_t       * range_ptr,
    size_t                range_len,
    const uint8_t       * formula_ptr,
    size_t                formula_len,
    uint32_t              dxf_id,
    uint8_t             * err_buf,
    size_t                err_buf_len);

/* Attach a color-scale conditional format. `has_mid!=0` → 3-stop
 * gradient (min→mid→max via 50th percentile); otherwise 2-stop
 * (min→max). ARGB values embedded per stop; no dxf_id needed. */
int32_t zlsx_sheet_writer_add_conditional_format_color_scale(
    zlsx_sheet_writer_t * sw,
    const uint8_t       * range_ptr,
    size_t                range_len,
    uint32_t              low_color_argb,
    uint8_t               has_mid,
    uint32_t              mid_color_argb,
    uint32_t              high_color_argb,
    uint8_t             * err_buf,
    size_t                err_buf_len);

/* Attach a data-bar conditional format. `color_argb` is the bar
 * fill (Excel's default is 0xFF638EC6). */
int32_t zlsx_sheet_writer_add_conditional_format_data_bar(
    zlsx_sheet_writer_t * sw,
    const uint8_t       * range_ptr,
    size_t                range_len,
    uint32_t              color_argb,
    uint8_t             * err_buf,
    size_t                err_buf_len);

/* ── Editor (load-modify-save) ──────────────────────────────────────
 * Open an existing xlsx; append rows, set cells, read / strip docProps,
 * recalculate / evaluate, save to a path or a buffer. Numeric / integer /
 * boolean / empty / string cells; a source with no `xl/sharedStrings.xml`
 * part gets one created on the first string append. The editor's
 * own scanner refuses ZIP64 / multi-disk / encrypted / data-descriptor
 * archives up front. Structural edits (rows, columns, sheets, table
 * columns) and the `pivots` read are the zlsx_status_v1 block at the
 * end of this header (S3a). History:
 * docs/plans/archive/load-modify-save.md.
 */

/* Open an xlsx for editing. Returns NULL on failure with `err_buf`
 * populated. Path is null-terminated UTF-8.
 */
zlsx_editor_t * zlsx_editor_open(
    const char * path,
    uint8_t    * err_buf,
    size_t       err_buf_len);

/* Drop the editor handle. Safe with NULL. */
void zlsx_editor_close(zlsx_editor_t * ed);

/* Append a single row to the sheet at `sheet_idx`. Cells are
 * borrowed for the duration of this call; the editor dupes string
 * contents internally so callers can reuse / free their buffers
 * after this returns. Returns 0 on success, -1 on failure with
 * `err_buf` populated (e.g. `IntegerExceedsExcelPrecision`,
 * `RowIndexOutOfRange`, `SheetIndexOutOfRange`, `NoSstInSource`).
 */
int32_t zlsx_editor_append_row(
    zlsx_editor_t     * ed,
    uint32_t            sheet_idx,
    const zlsx_cell_t * cells_ptr,
    size_t              cells_len,
    uint8_t           * err_buf,
    size_t              err_buf_len);

/* In-place cell mutation (Phase 3d, iter-cm-2). Replaces or inserts a
 * single cell on `sheet_idx` at (`row`, `col`). `row` is 1-based;
 * `col` is 0-based (A=0, B=1, …). The cell is borrowed for the
 * duration of this call; the editor dupes string contents internally
 * so callers can reuse / free their buffers after this returns.
 *
 * Returns 0 on success, -1 on failure with `err_buf` populated.
 * Notable typed errors include `SetCellSourceCellHasMetadata` (the
 * source cell carries `s="N"` styles or non-canonical body —
 * preserve-and-merge is iter-cm-2e, not yet shipped),
 * `SheetHasUnsavedAppends`, `SheetIndexOutOfRange`, and
 * `RowIndexOutOfRange`.
 */
int32_t zlsx_editor_set_cell(
    zlsx_editor_t     * ed,
    uint32_t            sheet_idx,
    uint32_t            row,
    uint32_t            col,
    const zlsx_cell_t * cell,
    uint8_t           * err_buf,
    size_t              err_buf_len);

/* Save the workbook (with any pending appends applied) atomically
 * to `out_path` (`out_path_len` bytes; not null-terminated). Returns
 * 0 on success, -1 on failure with `err_buf` populated.
 */
int32_t zlsx_editor_save(
    zlsx_editor_t * ed,
    const char    * out_path,
    size_t          out_path_len,
    uint8_t       * err_buf,
    size_t          err_buf_len);

/* ---- Document properties (Z3) --------------------------------------
 *
 * Field selectors for zlsx_editor_docprop_at(). Numeric values are part
 * of the ABI contract: appending is safe, renumbering is not.
 */
#define ZLSX_DOCPROP_CREATOR           0u
#define ZLSX_DOCPROP_LAST_MODIFIED_BY  1u
#define ZLSX_DOCPROP_TITLE             2u
#define ZLSX_DOCPROP_SUBJECT           3u
#define ZLSX_DOCPROP_DESCRIPTION       4u
#define ZLSX_DOCPROP_KEYWORDS          5u
#define ZLSX_DOCPROP_CATEGORY          6u
#define ZLSX_DOCPROP_CREATED           7u
#define ZLSX_DOCPROP_MODIFIED          8u
#define ZLSX_DOCPROP_REVISION          9u
#define ZLSX_DOCPROP_COMPANY          10u
#define ZLSX_DOCPROP_MANAGER          11u
#define ZLSX_DOCPROP_APPLICATION      12u
#define ZLSX_DOCPROP_HYPERLINK_BASE   13u

/* Read one docProps field. Pointer lifetime matches the Editor.
 * Returns 0 on success (absent field yields *out_len == 0),
 * -1 on unknown field id, -2 if the properties could not be read. */
int32_t zlsx_editor_docprop_at(
    zlsx_editor_t       * ed,
    uint32_t              field,
    const unsigned char** out_ptr,
    size_t              * out_len);

/* Non-zero when docProps/custom.xml is present; -1 on read failure. */
int32_t zlsx_editor_has_custom_properties(zlsx_editor_t * ed);

/* Strip identifying document metadata, staged for the next save.
 * strip_timestamps also drops created/modified/revision, which the
 * default mask keeps. Returns 0 on success, -1 on failure. */
int32_t zlsx_editor_strip_doc_props(
    zlsx_editor_t * ed,
    int32_t         strip_timestamps,
    uint8_t       * err_buf,
    size_t          err_buf_len);


/* ── Embeddings (E5) ───────────────────────────────────────────────
 *
 * A read-only handle over an .xlsx's embedding set. Separate from
 * zlsx_book_t because embeddings live in the OPC part model while
 * zlsx_book_t is the streaming cell reader.
 *
 * The three states are the point of this surface. Some spreadsheet
 * applications rebuild the archive on save and delete the vector
 * parts; a ~200-byte recovery record survives that, so a stripped
 * workbook can still report what it used to hold. A consumer must be
 * able to tell the three cases apart. */
#define ZLSX_EMB_ABSENT   0u  /* never had embeddings */
#define ZLSX_EMB_PRESENT  1u  /* vectors available */
#define ZLSX_EMB_STRIPPED 2u  /* vectors deleted; provenance recovered */

#define ZLSX_EMB_CARRIER_DEFINED_NAME 0u
#define ZLSX_EMB_CARRIER_DOC_PROPS    1u
/* Opt-in carrier; the only one Apple Numbers preserves. */
#define ZLSX_EMB_CARRIER_CELL_DATA    2u

/* Open an .xlsx and resolve its embedding state. NULL on I/O or parse
 * failure; an absent or stripped set is a successful open. */
zlsx_emb_t * zlsx_emb_open(
    const char * path,
    uint8_t    * err_buf,
    size_t       err_buf_len);

void zlsx_emb_close(zlsx_emb_t * emb);

/* One of the ZLSX_EMB_* state constants. */
uint32_t zlsx_emb_state(zlsx_emb_t * emb);

/* Provenance. Available for PRESENT *and* STRIPPED — recovering it
 * after a strip is the reason the recovery record exists. String
 * getters copy into out_buf null-terminated and return the full
 * length; pass out_buf_len 0 to query the length. */
size_t   zlsx_emb_model(zlsx_emb_t * emb, uint8_t * out_buf, size_t out_buf_len);
uint32_t zlsx_emb_dim(zlsx_emb_t * emb);
size_t   zlsx_emb_dtype(zlsx_emb_t * emb, uint8_t * out_buf, size_t out_buf_len);

size_t   zlsx_emb_coverage_count(zlsx_emb_t * emb);
size_t   zlsx_emb_coverage_id(zlsx_emb_t * emb, size_t i, uint8_t * out_buf, size_t out_buf_len);
size_t   zlsx_emb_coverage_range(zlsx_emb_t * emb, size_t i, uint8_t * out_buf, size_t out_buf_len);
size_t   zlsx_emb_coverage_sheet(zlsx_emb_t * emb, size_t i, uint8_t * out_buf, size_t out_buf_len);
uint32_t zlsx_emb_coverage_rows(zlsx_emb_t * emb, size_t i);

/* STRIPPED only; 0 otherwise. Content fingerprint at embed time —
 * recomputable from current cells, so equal means the covered content
 * has not drifted and a re-embed reproduces the same vectors. */
uint64_t zlsx_emb_digest(zlsx_emb_t * emb);

/* STRIPPED only. Which carrier the record was recovered from. */
uint32_t zlsx_emb_carrier(zlsx_emb_t * emb);

/* Hash value marking a deleted row. Exposed rather than hard-coded in
 * each binding so the tombstone contract has one definition. */
uint64_t zlsx_emb_tombstone(void);

/* Decode coverage i's vectors as f32, row-major [rows][dim].
 * out_len must be exactly rows * dim.
 * Returns 0 on success, -1 on bad index or size, -2 when the state is
 * not PRESENT (a stripped coverage has provenance but no vectors). */
int32_t zlsx_emb_vectors(zlsx_emb_t * emb, size_t i, float * out, size_t out_len);

/* Copy coverage i's per-row content hashes; out_len must equal rows.
 * Same return convention. A row whose hash equals zlsx_emb_tombstone()
 * was deleted. */
int32_t zlsx_emb_hashes(zlsx_emb_t * emb, size_t i, uint64_t * out, size_t out_len);

/* ── Formula engine (M9a1) ──────────────────────────────────────────
 *
 * zlsx_status_v1 — NEW exports below only; everything above keeps its
 * shipped 0/-1 convention:
 *    0  OK
 *   -1  generic error (Zig error name in errbuf)
 *   -2  typed Plane-2 refusal (zlsx_diag_v1 populated when supplied)
 *   -3  allocation failure
 *   -4  reserved (never returned by v1)
 *   -5  cancellation/deadline observed before commit
 *
 * struct_size discipline: every struct below starts with size_t
 * struct_size, set by the caller to sizeof() at its compile time.
 * Readers use min(caller, known); writers touch only that prefix and
 * NEVER write beyond the known v1 size, even when the caller declares
 * more. Output structs are zero-initialised (excluding struct_size)
 * within the known prefix on entry. Sizes below the full v1 struct are
 * rejected with -1 "StructSizeTooSmall".
 *
 * Ownership: fields documented library-owned are released by the
 * per-type release fn; the structs themselves are the caller's.
 * Release fns are NULL-safe and no-ops on zeroed/already-released
 * structs.
 *
 * Layout contract: docs/plans/c-abi-status-v1.md; every offset is
 * pinned by comptime tests in src/c_abi.zig. */

#define ZLSX_OK             0
#define ZLSX_ERROR        (-1)
#define ZLSX_REFUSED      (-2)
#define ZLSX_NOMEM        (-3)
#define ZLSX_CANCELLED    (-5)

/* §10's fourteen-plane refusal vocabulary. Values are pinned by a Zig
 * test against the engine enum — they are ABI, not implementation. */
#define ZLSX_PLANE_UNSUPPORTED_FUNCTION        0u
#define ZLSX_PLANE_UNSUPPORTED_CONSTRUCT       1u
#define ZLSX_PLANE_PRECISION_AS_DISPLAYED      2u
#define ZLSX_PLANE_MALFORMED_INPUT             3u
#define ZLSX_PLANE_LOCALE_SENSITIVE_INPUT      4u
#define ZLSX_PLANE_DATA_TABLE_UNSUPPORTED      5u
#define ZLSX_PLANE_SIGNED_WORKBOOK             6u
#define ZLSX_PLANE_STALE_EMBEDDINGS            7u
#define ZLSX_PLANE_ANCHOR_REQUIRED             8u
#define ZLSX_PLANE_CYCLE                       9u
#define ZLSX_PLANE_DYNAMIC_REF_UNSTABLE       10u
#define ZLSX_PLANE_SPILL_PERSIST_UNSUPPORTED  11u
#define ZLSX_PLANE_RESULT_NOT_REPRESENTABLE   12u
#define ZLSX_PLANE_LIMIT_EXCEEDED             13u
#define ZLSX_PLANE_NONE               0xFFFFFFFFu

#define ZLSX_FIDELITY_EXCEL        0u
#define ZLSX_FIDELITY_IEEE         1u
#define ZLSX_PROFILE_WINDOWS_1252  0u
#define ZLSX_DIALECT_DYNAMIC_ARRAY 0u
#define ZLSX_DIALECT_LEGACY        1u
#define ZLSX_DIALECT_NONE          0xFFFFFFFFu  /* resolved echo of a recalc */
#define ZLSX_ON_UNSUPPORTED_REFUSE              0u
#define ZLSX_ON_UNSUPPORTED_KEEP_STALE_AND_MARK 1u

#define ZLSX_VALUE_NUMBER 0u
#define ZLSX_VALUE_TEXT   1u
#define ZLSX_VALUE_BOOL   2u
#define ZLSX_VALUE_ERROR  3u   /* an Excel error VALUE — a successful result */

/* Cross-thread cancellation token. Caller-owned; must outlive every
 * call it is passed to. trigger() is callable from any thread. */
typedef struct zlsx_cancel_token_t zlsx_cancel_token_t;

int32_t zlsx_cancel_token_new(zlsx_cancel_token_t ** out, char * errbuf, size_t errbuf_len);
void    zlsx_cancel_token_trigger(zlsx_cancel_token_t * tok);
void    zlsx_cancel_token_free(zlsx_cancel_token_t * tok);

/* One census entry (§5.7.7): a construct the evaluator could not
 * implement, and where it was. */
typedef struct zlsx_census_entry_v1 {
    uint32_t plane;   /* ZLSX_PLANE_* */
    uint32_t sheet;   /* 0-based */
    uint32_t row;     /* 1-based; 0 = not about a cell */
    uint32_t col;     /* 0-based */
} zlsx_census_entry_v1;

typedef struct zlsx_diag_v1 {
    size_t   struct_size;            /* caller sets to sizeof(zlsx_diag_v1) */
    uint32_t plane;                  /* ZLSX_PLANE_NONE when not a refusal */
    uint32_t census_truncated;       /* 0/1 */
    char     error_name[64];         /* NUL-terminated, e.g. "FormulaCycle" */
    const zlsx_census_entry_v1 * census;   /* library-owned; NULL iff census_len == 0 */
    size_t   census_len;
} zlsx_diag_v1;

void zlsx_diag_release(zlsx_diag_v1 * d);

/* §5.5 run inputs. now_utc_ms and rng_seed have no defaults on
 * purpose: a library that defaulted them would read a clock or an
 * entropy source and "equal inputs => equal output" would be false.
 * Limit fields: 0 = the documented §9 default (echoed numerically in
 * zlsx_resolved_v1, never as 0). */
typedef struct zlsx_run_v1 {
    size_t   struct_size;
    int64_t  now_utc_ms;
    uint64_t rng_seed;
    int32_t  utc_offset_min;         /* [-1440, 1440] */
    uint32_t fidelity;               /* ZLSX_FIDELITY_* */
    uint32_t profile;                /* ZLSX_PROFILE_* */
    uint32_t dialect;                /* ZLSX_DIALECT_*; standalone eval only */
    uint32_t on_unsupported;         /* ZLSX_ON_UNSUPPORTED_*; recalc only */
    uint32_t _reserved0;
    uint64_t max_run_arena_bytes;
    uint64_t max_matrix_cells;
    uint64_t max_string_payload_bytes;
    uint64_t max_retained_ast_bytes;
    uint64_t max_diagnostics_bytes;
    uint64_t timeout_ms;             /* 0 = none; absolute deadline derived at entry */
    zlsx_cancel_token_t * cancel;    /* NULL = non-cancellable */
} zlsx_run_v1;

/* The §5.5 echo: exactly what the run resolved to. */
typedef struct zlsx_resolved_v1 {
    size_t   struct_size;
    int64_t  now_utc_ms;
    uint64_t rng_seed;
    int32_t  utc_offset_min;
    uint32_t fidelity;
    uint32_t profile;
    uint32_t dialect;                /* ZLSX_DIALECT_NONE for a recalc */
    uint64_t max_run_arena_bytes;
    uint64_t max_matrix_cells;
    uint64_t max_string_payload_bytes;
    uint64_t max_retained_ast_bytes;
    uint64_t max_diagnostics_bytes;
} zlsx_resolved_v1;

typedef struct zlsx_recalc_report_v1 {
    size_t   struct_size;
    uint32_t sheets_patched;
    uint32_t cells_written;
    uint32_t passes;
    uint32_t non_converged_cells;
    uint32_t dynamic_passes;
    uint32_t kept_stale;             /* 0/1: §5.7.7's mark-only path taken */
    uint32_t calc_chain_removed;     /* 0/1 */
    uint32_t census_truncated;       /* 0/1 */
    uint64_t retained_generations;
    uint64_t retained_bytes;
    uint32_t durability_warning;     /* dormant §5.7.9 slot; 0 for in-memory recalc */
    int32_t  durability_errno;
    zlsx_resolved_v1 resolved;
    uint32_t resolved_present;       /* 0/1 */
    uint32_t _reserved0;
    const zlsx_census_entry_v1 * census;   /* library-owned */
    size_t   census_len;
} zlsx_recalc_report_v1;

void zlsx_recalc_report_release(zlsx_recalc_report_v1 * r);

/* One evaluated element (§12.3's descriptor). Blank never crosses the
 * boundary — it publishes as number 0 (§5.3a). */
typedef struct zlsx_value_elem_v1 {
    uint8_t  tag;                    /* ZLSX_VALUE_* */
    uint8_t  _reserved[7];
    double   num;                    /* number, or bool as 0/1 */
    uint64_t payload_off;            /* into the value's payload arena */
    uint64_t payload_len;            /* text / error spelling; 0 otherwise */
} zlsx_value_elem_v1;

typedef struct zlsx_value_v1 {
    size_t   struct_size;
    uint32_t rows;
    uint32_t cols;
    uint32_t is_matrix;              /* 0 = scalar (rows == cols == 1) */
    uint32_t _reserved0;
    const zlsx_value_elem_v1 * elems;      /* library-owned, row-major */
    size_t   elems_len;
    const uint8_t * payload;               /* library-owned, one arena */
    size_t   payload_len;
} zlsx_value_v1;

void zlsx_value_release(zlsx_value_v1 * v);

/* Engine identity (§12.4): "zlsx <semver>; excel_fp_rules_v1; rng_v1;
 * collation_v1; <arch>-<os>-<abi>; <build-hash>". Static storage,
 * never NULL. Feature-probe this symbol (absent export = feature off);
 * mixed-fingerprint fleets must refuse to share recalc results. */
const char * zlsx_engine_fingerprint(void);

/* §5.7.7's mark-only transaction: keep every cached value, set
 * fullCalcOnLoad="1", remove nothing else. Typed refusals (-2) include
 * FormulaPrecisionAsDisplayed. diag is optional (NULL ok). */
int32_t zlsx_editor_mark_recalc_on_load(zlsx_editor_t * ed,
        zlsx_diag_v1 * diag, char * errbuf, size_t errbuf_len);

/* §5.7's in-memory transaction (M5d2 pipeline): recalculate every
 * formula cell and swap the result in as the final operation. On any
 * non-zero status the workbook is exactly as it was. No file I/O. */
int32_t zlsx_editor_recalculate(zlsx_editor_t * ed, const zlsx_run_v1 * run,
        zlsx_recalc_report_v1 * report, zlsx_diag_v1 * diag,
        char * errbuf, size_t errbuf_len);

/* Standalone cache-based evaluation (M6 semantics). Scratch-only: the
 * workbook is byte-identical before and after. anchor_row = 0 means no
 * anchor; site-dependent formulas (`@`) then refuse rather than guess.
 * out_resolved and diag are optional (NULL ok). */
int32_t zlsx_editor_evaluate(zlsx_editor_t * ed,
        const uint8_t * formula_ptr, size_t formula_len,
        uint32_t sheet_idx,
        uint32_t anchor_row,         /* 1-based; 0 = absent */
        uint32_t anchor_col,         /* 0-based; read only when anchor_row != 0 */
        const zlsx_run_v1 * run,
        zlsx_value_v1 * out_value, zlsx_resolved_v1 * out_resolved,
        zlsx_diag_v1 * diag, char * errbuf, size_t errbuf_len);

/* Feature macros — compile-time counterpart of the dlsym probe. */
#define ZLSX_HAS_FINGERPRINT 1
#define ZLSX_HAS_MARK_RECALC 1
#define ZLSX_HAS_RECALC      1
#define ZLSX_HAS_EVAL        1
#define ZLSX_HAS_CANCEL      1

/* ── Formula engine (M9a2): buffers, the file transaction, writer ──
 *
 * Part 2 of the C ABI (§12.3). Same zlsx_status_v1 contract, same
 * struct_size discipline as the M9a1 block above; every M9a1 layout
 * stays frozen. Layout note: docs/plans/c-abi-status-v1.md. */

/* Release a buffer an M9a2 export allocated. NULL-safe. (The legacy
 * zlsx_buffer_free keeps its shipped contract; this is the status-era
 * name §12.3 pins.) */
void zlsx_buffer_release(uint8_t * ptr, size_t len);

/* Serialize the editor's current state — staged mutations included —
 * into a library-allocated buffer (§5.10). An untouched editor hands
 * back the source bytes verbatim. On non-zero status *out_ptr is NULL
 * and *out_len is 0. Release with zlsx_buffer_release. */
int32_t zlsx_editor_save_to_buffer(zlsx_editor_t * ed,
        uint8_t ** out_ptr, size_t * out_len,
        char * errbuf, size_t errbuf_len);

/* Open an editor over a workbook already in memory (§5.10). The
 * borrow ends when this returns: `data` is copied, so the caller may
 * free or reuse it immediately — the zlsx_book_open_buffer contract.
 * On ZLSX_OK *out holds the handle (close with zlsx_editor_close);
 * on any other status *out is NULL. */
int32_t zlsx_open_buffer(const uint8_t * data, size_t data_len,
        zlsx_editor_t ** out, char * errbuf, size_t errbuf_len);

/* §5.7.9's file transaction: recalculate, serialize from the prepared
 * candidate, rename, swap in memory between the rename and the
 * directory fsync. Any failure before the rename leaves BOTH the
 * destination's prior bytes (or its absence) and the editor's memory
 * untouched. A directory fsync failing after the rename is
 * report->durability_warning (+ durability_errno) on a ZLSX_OK return
 * — the §5.7.9 slot goes live here — never an error. A -2 refusal
 * carries the refusing cells in diag->census. */
int32_t zlsx_editor_save_with_recalc(zlsx_editor_t * ed,
        const uint8_t * out_path_ptr, size_t out_path_len,
        const zlsx_run_v1 * run,
        zlsx_recalc_report_v1 * report, zlsx_diag_v1 * diag,
        char * errbuf, size_t errbuf_len);

/* The producer-side file transaction (§12.3 Writer.save(recalculate=)):
 * emit the writer's archive to memory, open it as a workbook, run the
 * same §5.7.9 transaction zlsx_editor_save_with_recalc runs. The
 * writer handle is neither consumed nor mutated. */
int32_t zlsx_writer_save_with_recalc(zlsx_writer_t * w,
        const uint8_t * out_path_ptr, size_t out_path_len,
        const zlsx_run_v1 * run,
        zlsx_recalc_report_v1 * report, zlsx_diag_v1 * diag,
        char * errbuf, size_t errbuf_len);

/* Formula dialect tags for zlsx_formula_cell_v1. DYNAMIC_ARRAY is
 * reserved ABI: the writer currently refuses it (-1) — its authored
 * metadata's Excel reference set has not arrived (§5.8b). */
#define ZLSX_FORMULA_SCALAR        0u
#define ZLSX_FORMULA_DYNAMIC_ARRAY 1u
#define ZLSX_FORMULA_CSE           2u

/* §12.3's per-cell formula descriptor — the shape a CSE rectangle
 * needs and the parallel text arrays of
 * zlsx_sheet_writer_write_row_with_formulas cannot encode. An array
 * element (no struct_size); the v1 layout is frozen at 40 bytes. */
typedef struct zlsx_formula_cell_v1 {
    const uint8_t * text;     /* NULL = plain value slot */
    size_t   text_len;        /* > 0 when text != NULL */
    uint32_t dialect;         /* ZLSX_FORMULA_* */
    uint32_t _reserved0;      /* always 0 */
    const uint8_t * ref;      /* CSE only: declared range, uppercase A1
                               * ("A1" or "A1:B2"); NULL otherwise */
    size_t   ref_len;
} zlsx_formula_cell_v1;

/* The v2 formula row: `formulas` is parallel to `cells`. A CSE ref is
 * legal ONLY on the rectangle's top-left (the anchor writes
 * <f t="array" ref>); the range's other cells arrive as plain value
 * slots in later rows (empty ones become bare <c> placeholders), a
 * formula inside an open rectangle refuses, and save refuses while
 * any rectangle is missing members. Every refusal here is a statement
 * about the call: -1, never -2. */
int32_t zlsx_sheet_writer_write_row_with_formulas_v2(zlsx_sheet_writer_t * sw,
        const zlsx_cell_t * cells,
        const zlsx_formula_cell_v1 * formulas,
        size_t cells_len,
        char * errbuf, size_t errbuf_len);

/* Feature macros — compile-time counterpart of the dlsym probe. */
#define ZLSX_HAS_SAVE_BUFFER      1   /* editor save_to_buffer + open_buffer + buffer_release */
#define ZLSX_HAS_SAVE_WITH_RECALC 1
#define ZLSX_HAS_WRITER_RECALC    1
#define ZLSX_HAS_FORMULAS_V2      1

/* ── S3a: structural edits + the pivots read (zlsx_status_v1) ───────
 *
 * The Editor's structural edits and the S6 `pivots` NDJSON shape,
 * under the same status contract as the M9a1 / M9a2 exports above:
 * ZLSX_OK, ZLSX_ERROR with the error name in `errbuf`, ZLSX_REFUSED
 * with `diag->error_name` set and `diag->plane == ZLSX_PLANE_NONE`
 * (these refusals carry no formula plane), ZLSX_NOMEM. `diag` is
 * nullable on every call. Contract: docs/plans/c-abi-status-v1.md §10.
 *
 * What refuses (-2) is a statement about the workbook:
 *   RowEditUnsafeForSheet / ColEditUnsafeForSheet — the edit lands
 *     inside a hosted pivot's footprint, on a host sheet a pivot also
 *     reads from, would collapse a table or delete its header row, or
 *     a carrier the scan cannot read is in the way;
 *   CannotDeleteLastSheet, DuplicateSheetName (add / rename, compared
 *     ASCII case-insensitively), TableColumnNameInUse,
 *     MalformedPivotXml (the pivot graph cannot be read whole), the
 *     workbook's own structure found broken (InternalSheetNameTooLong,
 *     MalformedWorkbookXml, IdSpaceExhausted, MissingRelationship,
 *     SheetCountMismatch, …);
 *   with their precise names, the worksheet transform's own verdicts
 *     (RowEditExceedsMaxRow, ColEditExceedsMaxCol, SplitPaneNotSupported,
 *     MalformedPaneSplit, MalformedSheetXml) and a carrier a sweep cannot
 *     read, materialise or move (MalformedDrawingXml — a drawing part
 *     the archive cannot decompress, a sheet drawing reference the
 *     strict anchors read cannot follow (malformed, duplicate, dangling,
 *     absent), one binding the spreadsheetDrawing namespace under a name
 *     the anchor walk cannot spell, a DTD, a `<` inside an attribute
 *     value (not well-formed XML), or an anchor it cannot read whole (no
 *     close, a corner absent or unparseable, two corner blocks that
 *     overlap): the strict read's verdicts on the anchors both walk,
 *     refused before the first mutation —
 *     MalformedVmlDrawing,
 *     MalformedCommentsXml, MalformedTableXml, MalformedExtensionXml,
 *     MalformedChartXml — an <xm:f> extension or chart <c:f> series
 *     carrier the sweep cannot read whole, or a chart part carrying a
 *     DTD or a `<` inside an attribute value, refused before the first
 *     mutation — the *CoordinateOverflow
 *     trio, PivotEditUnsafe, SqrefCollapseUnsafe — a delete collapsing
 *     EVERY area of a DV/CF sqref, which Excel resolves by deleting the
 *     rule — MissingSheetPart, NoSheetData; a generic
 *     MalformedXml from a rewriter's consistency guard stays -1) — the
 *     full list is
 *     docs/plans/c-abi-status-v1.md §10; the typed worksheet parser's
 *     MalformedXml / UnexpectedEof cross as MalformedSheetXml, and a
 *     pivot part the archive cannot materialise as MalformedPivotXml.
 * What fails (-1) is a statement about the call: SheetIndexOutOfRange,
 * RowIndexOutOfRange, ColumnIndexOutOfRange, InvalidSheetName,
 * InvalidTableColumnName, TableNotFound / TableColumnNotFound (a
 * selector that names nothing, like a sheet index), InvalidInput (NULL
 * where bytes are required), and the sequencing errors RowEditRequiresCleanSheet /
 * ColEditRequiresCleanSheet / SheetDeleteRequiresCleanState — a
 * structural edit needs the sheet (the workbook, for a sheet delete)
 * free of staged cell writes and appended rows: save first.
 *
 * Every edit is staged in memory; zlsx_editor_save /
 * zlsx_editor_save_to_buffer commit it, with every cross-part
 * rewriter the Zig editor carries (formulas in every dialect, defined
 * names, hyperlinks, DV / CF, merges, panes, autoFilter, tables,
 * drawings, comments, `<xm:f>` extensions, chart `<c:f>` series
 * formulas, and — under a row / column edit — pivot locations and
 * sources). Rows are 1-based; columns
 * 0-based (A = 0), as zlsx_editor_set_cell spells them; sheet indices
 * 0-based. */

/* Insert a blank row before `before_row`; rows at or below it shift
 * down by one. */
int32_t zlsx_editor_insert_row(zlsx_editor_t * ed,
        uint32_t sheet_idx, uint32_t before_row,
        zlsx_diag_v1 * diag, char * errbuf, size_t errbuf_len);

/* Delete row `row`; rows below it shift up by one. */
int32_t zlsx_editor_delete_row(zlsx_editor_t * ed,
        uint32_t sheet_idx, uint32_t row,
        zlsx_diag_v1 * diag, char * errbuf, size_t errbuf_len);

/* Insert a blank column before `before_col` (0-based, A = 0). */
int32_t zlsx_editor_insert_column(zlsx_editor_t * ed,
        uint32_t sheet_idx, uint32_t before_col,
        zlsx_diag_v1 * diag, char * errbuf, size_t errbuf_len);

/* Delete column `col` (0-based, A = 0). */
int32_t zlsx_editor_delete_column(zlsx_editor_t * ed,
        uint32_t sheet_idx, uint32_t col,
        zlsx_diag_v1 * diag, char * errbuf, size_t errbuf_len);

/* Append an empty sheet named `name` (`name_len` UTF-8 bytes, not
 * NUL-terminated; judged by the fresh writer's rules). `out_sheet_idx`
 * (nullable) receives the new index on ZLSX_OK and UINT32_MAX
 * otherwise. */
int32_t zlsx_editor_add_sheet(zlsx_editor_t * ed,
        const uint8_t * name, size_t name_len,
        uint32_t * out_sheet_idx,
        zlsx_diag_v1 * diag, char * errbuf, size_t errbuf_len);

/* Rename sheet `sheet_idx`; cross-sheet references (formulas, defined
 * names, hyperlinks, DV / CF, <xm:f>, chart <c:f> series formulas)
 * follow. A pivot cache's
 * worksheetSource@sheet does NOT (a Zig-editor hole this row inherits):
 * the spelling goes stale and zlsx_editor_pivots_ndjson reports it as
 * "resolved":null. */
int32_t zlsx_editor_rename_sheet(zlsx_editor_t * ed,
        uint32_t sheet_idx, const uint8_t * name, size_t name_len,
        zlsx_diag_v1 * diag, char * errbuf, size_t errbuf_len);

/* Delete sheet `sheet_idx` (never the last one); references into it
 * become #REF!, indices above it shift down by one. A pivot cache that
 * read the sheet by name keeps the stale spelling (see rename). */
int32_t zlsx_editor_delete_sheet(zlsx_editor_t * ed,
        uint32_t sheet_idx,
        zlsx_diag_v1 * diag, char * errbuf, size_t errbuf_len);

/* Rename column `old_name` of table `table_name` to `new_name`: the
 * <tableColumn>, the table's formulas, every structured reference
 * workbook-wide, defined names, hyperlinks, DV / CF and the header
 * cell. Names are plain (decoded) UTF-8, not NUL-terminated. */
int32_t zlsx_editor_rename_table_column(zlsx_editor_t * ed,
        const uint8_t * table_name, size_t table_name_len,
        const uint8_t * old_name, size_t old_name_len,
        const uint8_t * new_name, size_t new_name_len,
        zlsx_diag_v1 * diag, char * errbuf, size_t errbuf_len);

/* The S6 `pivots` records — one {"kind":"pivot",…} line per pivot
 * table in host-sheet order, then one {"kind":"pivot_cache",…} line
 * per cache no table reads — as a library-allocated UTF-8 buffer,
 * byte-for-byte what `zlsx pivots <file>` prints (docs/cli.md,
 * "pivots"; the shape frozen at the S6 gate). Read over the editor's
 * current workbook state: structural edits (rows, columns, sheets,
 * table columns) are visible immediately; staged zlsx_editor_set_cell /
 * append_row writes reach the pivot graph at save, where a cache whose
 * source they change is rebuilt or marked — save, then read, to see
 * them. A workbook without pivots is
 * ZLSX_OK with (*out, *out_len) = (NULL, 0). Release with
 * zlsx_buffer_release. */
int32_t zlsx_editor_pivots_ndjson(zlsx_editor_t * ed,
        uint8_t ** out, size_t * out_len,
        zlsx_diag_v1 * diag, char * errbuf, size_t errbuf_len);

/* The S3b `defined-names` records — one {"kind":"defined_name",…} line
 * per <definedName> of xl/workbook.xml, in document order — as a
 * library-allocated UTF-8 buffer, byte-for-byte what
 * `zlsx defined-names <file>` prints with no selector (docs/cli.md,
 * "defined-names"). `body` is the formula text as authored — nothing
 * resolved or rewritten. Read over the editor's current workbook
 * state: structural edits and the name sweeps they carry (a sheet
 * rename rewriting the bodies) are visible immediately. A workbook
 * without defined names is ZLSX_OK with (*out, *out_len) = (NULL, 0).
 * An inventory that cannot be served faithfully — a carrier that does
 * not decode, malformed UTF-8, a body with embedded markup — refuses
 * whole (ZLSX_REFUSED, MalformedWorkbookXml) rather than hand over a
 * record that lies. Release with zlsx_buffer_release. */
int32_t zlsx_editor_defined_names_ndjson(zlsx_editor_t * ed,
        uint8_t ** out, size_t * out_len,
        zlsx_diag_v1 * diag, char * errbuf, size_t errbuf_len);

/* The S3b `conditional-formats` records — one
 * {"kind":"conditional_format",…} line per <cfRule>, sheets in
 * workbook order, rules in sheet-document order — as a
 * library-allocated UTF-8 buffer, byte-for-byte what
 * `zlsx conditional-formats <file>` prints with no selector
 * (docs/cli.md, "conditional-formats"). The record is the rule
 * envelope (sqref, rule_type, formulas, dxf_id, priority), not the
 * visual payload — <colorScale> / <dataBar> / <iconSet> bodies and
 * the <dxfs> styles stay in their parts. Read over the editor's
 * current parts: structural edits and the DV/CF sweeps they carry are
 * visible immediately; staged cell writes never touch the rule
 * machinery. A workbook without conditional formatting is ZLSX_OK
 * with (*out, *out_len) = (NULL, 0). An inventory that cannot be
 * served faithfully refuses whole — a sheet list the strict workbook
 * read cannot prove (MalformedWorkbookXml) or a sheet part the strict
 * walk cannot (MalformedSheetXml), both ZLSX_REFUSED — rather than
 * hand over a record that lies. (An archive past the decompression
 * caps fails at open — the caps are checked on the open-time
 * directory walk.) Release with zlsx_buffer_release. */
int32_t zlsx_editor_conditional_formats_ndjson(zlsx_editor_t * ed,
        uint8_t ** out, size_t * out_len,
        zlsx_diag_v1 * diag, char * errbuf, size_t errbuf_len);

/* The S3b `anchors` records — one {"kind":"image_anchor",…} line per
 * anchored image and one {"kind":"chart_anchor",…} line per anchored
 * chart, sheets in workbook order, a sheet's images before its
 * charts, each class in drawing-document order — as a
 * library-allocated UTF-8 buffer, byte-for-byte what
 * `zlsx anchors <file>` prints with no selector (docs/cli.md,
 * "anchors"). The record is the anchor geometry (anchor ∈ two_cell /
 * one_cell / absolute, from / to 1-based with EMU offsets, absolute
 * {x,y,cx,cy} in EMUs) and where the payload lives (part; an image's
 * byte count; a chart's chart_type + entity-decoded series_refs),
 * never the payload — image bytes and chart XML stay in their parts.
 * Read over the editor's current parts: structural edits and the
 * drawing sweeps they carry (a row insert moving an anchor, a rename
 * renaming `sheet`, a chart's series_refs respelled by the chart <c:f>
 * sweep — the new sheet name after a rename, shifted rows after an
 * insert on the sheet they name) are visible immediately; staged cell
 * writes never touch a drawing. Drawings are walked under every prefix
 * bound to the spreadsheetDrawing namespace — the canonical xdr:, any
 * other, and the DEFAULT namespace (openpyxl's `<wsDr xmlns="…">
 * <oneCellAnchor>` spelling) — the same resolution the row / column
 * sweep moves anchors under. A workbook without
 * anchored objects is ZLSX_OK with (*out, *out_len) = (NULL, 0). An
 * inventory that cannot be served faithfully refuses whole — a sheet
 * list the strict workbook read cannot prove (MalformedWorkbookXml),
 * a drawing graph the strict walk cannot read whole — a
 * spreadsheetDrawing binding under a name it cannot spell, a part
 * carrying a <!DOCTYPE or a `<` inside an attribute value included —
 * (MalformedDrawingXml), or an anchor
 * on a worksheet part the workbook
 * does not list (DrawingOnUnlistedSheet), all ZLSX_REFUSED — rather
 * than hand over a record that lies or a list with a hole. (An
 * archive past the decompression caps fails at open.) Release with
 * zlsx_buffer_release. */
int32_t zlsx_editor_anchors_ndjson(zlsx_editor_t * ed,
        uint8_t ** out, size_t * out_len,
        zlsx_diag_v1 * diag, char * errbuf, size_t errbuf_len);

/* The S3b `sheet-props` records — one {"kind":"sheet_props",…} line
 * per workbook sheet, workbook order — as a library-allocated UTF-8
 * buffer, byte-for-byte what `zlsx sheet-props <file>` prints with no
 * selector (docs/cli.md, "sheet-props"). Each record is the sheet's
 * <dimension ref> as authored (null when the element or the attribute
 * is absent) and the <pane> of its FIRST <sheetView> as authored (null
 * when there is none): x_split / y_split / top_left_cell / active_pane
 * / state, each null when the source omits it, no schema default
 * applied, split panes reported as written (the lenient
 * Worksheet.freezePane narrows to frozen panes; this read does not).
 * Later <sheetView> elements keep their own panes in the part. Read
 * over the editor's current parts: structural edits and the sheet
 * sweeps they carry (a rename renaming `sheet`, a row insert growing
 * `dimension` and moving a frozen pane's split and `top_left_cell`)
 * are visible immediately; staged cell writes never touch the extent
 * or the views. The buffer is never empty on ZLSX_OK: a sheetless
 * workbook (a missing or empty <sheets>) is refused, below. An
 * inventory that cannot be served faithfully refuses whole — a sheet
 * list the strict workbook read cannot prove (MalformedWorkbookXml —
 * the docs/cli.md "conditional-formats" contract: a carrier-less
 * entry, an unverifiable or duplicate relationship, two entries
 * reaching one part, two names decoding to one spelling, an empty
 * list; a ghost under <extLst> is simply not an entry), or a sheet
 * part the strict walk
 * cannot prove a pane / extent for (MalformedSheetXml: a second
 * <dimension> / <sheetViews> / first-view <pane>, a duplicate
 * attribute on that machinery, an MCE construct at a recognized slot,
 * a carrier that does not decode), both ZLSX_REFUSED — rather than
 * hand over a record that lies or a list with a hole. (An archive
 * past the decompression caps fails at open.) Release with
 * zlsx_buffer_release. */
int32_t zlsx_editor_sheet_props_ndjson(zlsx_editor_t * ed,
        uint8_t ** out, size_t * out_len,
        zlsx_diag_v1 * diag, char * errbuf, size_t errbuf_len);

/* The S3b `calc-props` record — the ONE {"kind":"calc_props",…} line
 * of xl/workbook.xml's <calcPr> — as a library-allocated UTF-8 buffer,
 * byte-for-byte what `zlsx calc-props <file>` prints (docs/cli.md,
 * "calc-props"): calc_id / full_calc_on_load / iterate / iterate_count
 * / iterate_delta as authored, every field null when the element or
 * the attribute is absent (a workbook without <calcPr> is a record of
 * nulls, never an empty buffer — the doc-props convention). Read over
 * the editor's current parts: zlsx_editor_mark_recalc_on_load and a
 * recalc that lands set fullCalcOnLoad="1" in place, visible
 * immediately; staged cell writes never touch the element. A slot the
 * read cannot report faithfully refuses (MalformedWorkbookXml,
 * ZLSX_REFUSED): two <calcPr> at the slot, one an MCE branch could
 * project there, a duplicate attribute, a carrier that does not
 * decode — and a <sheets> list the same strict walk cannot prove
 * (two wrappers, an empty one). Release with zlsx_buffer_release. */
int32_t zlsx_editor_calc_props_ndjson(zlsx_editor_t * ed,
        uint8_t ** out, size_t * out_len,
        zlsx_diag_v1 * diag, char * errbuf, size_t errbuf_len);

/* ── S3c slice 1: the embedding write (zlsx_status_v1) ─────────────
 *
 * Workbook.setEmbeddings on the editor handle. One call writes the
 * whole embedding set — xl/zlsxEmbeddings/index.xml, a vec.bin +
 * hashes.bin pair per coverage, the workbook→index relationship and
 * the recovery record in its two invisible carriers (a hidden defined
 * name, docProps/custom.xml) — and REPLACES any previous set (the
 * parts of a coverage id that disappears stay in the archive as
 * orphans, the Zig contract). Staged in memory; zlsx_editor_save
 * commits. Read it back with the zlsx_emb_* handle above: vectors
 * cross here as the f32 [rows][dim] matrix zlsx_emb_vectors hands
 * back and hashes as the uint64 per-row list zlsx_emb_hashes hands
 * back, so read → re-embed → write is one shape. */

/* One coverage. An array element (no struct_size, the
 * zlsx_formula_cell_v1 precedent); the v1 layout is frozen at 88
 * bytes. `rows` below is the range's row count (A2:A100 → 99). */
typedef struct zlsx_emb_coverage_v1 {
    const uint8_t  * id;         size_t id_len;       /* 1–63 of [A-Za-z0-9_-] */
    const uint8_t  * range;      size_t range_len;    /* A1 range, "A2:A100" */
    const uint8_t  * column;     size_t column_len;   /* the embedded column, inside the range */
    const float    * vectors;    size_t vectors_len;  /* == rows * dim, row-major, range order */
    const uint64_t * hashes;     size_t hashes_len;   /* == rows; zlsx_emb_tombstone() = no vector */
    uint32_t         sheet_idx;                       /* 0-based */
    uint32_t         include_formulas;                /* 0 or 1 */
} zlsx_emb_coverage_v1;

/* dtype is spelled as zlsx_emb_dtype spells it: "f32" or
 * "int8-sym-per-vec" (quantized here, one f32 scale per row); the
 * three other names the read knows have no writer (UnsupportedDtype).
 * flags is reserved and must be 0.
 *
 * -1 (a statement about the call, each raised BEFORE the first part
 * is written): InvalidInput — NULL handle, NULL bytes or arrays with
 * a non-zero length, a set flag, include_formulas past 1;
 * InvalidEmbeddingInput — no coverage, dim 0, a vectors_len or
 * hashes_len that disagrees with the range; InvalidDtype;
 * UnsupportedDtype; SheetIndexOutOfRange; the index's own rules
 * InvalidCoverageId, InvalidRange (the range, or a column outside
 * it), DuplicateCoverageId, CoverageOverlap; InvalidXmlByte (a C0
 * control byte in model — the rule every other text channel
 * enforces). -2 (ZLSX_REFUSED, a statement about the workbook, the
 * name in the diag with plane NONE): MissingWorkbookRels /
 * MalformedWorkbookRels — no xl/_rels/workbook.xml.rels, or one
 * without the </Relationships> the workbook→index relationship
 * lands before, or a _rels/.rels without it when docProps/custom.xml
 * has to be created for the recovery record (both checked before the
 * first write); IdSpaceExhausted — the rels file's rId space, or an
 * existing docProps/custom.xml's pid space, already at UINT32_MAX (a
 * hostile part; checked before the first write); MissingRelationship
 * — a sheet whose part the workbook's rels do not reach;
 * EmbeddingExceedsArchiveLimit — a
 * part past the 512 MiB read cap (sized here from the inputs, before
 * a vector byte is read), OR the recovery record past its ceiling of
 * 16 × 200 bytes — roughly eighty coverages at typical ids, or a
 * ~3 KB model name (encoded before the first write); and the
 * package's own MissingContentTypes / MalformedContentTypes /
 * MalformedWorkbookXml. A -2 or -3 that fires AFTER the first part
 * write — an allocation failure, an index past the cap, a content
 * types or workbook part the carriers cannot patch — leaves the
 * staged part set partially replaced: discard the editor without
 * saving. The recalc transactions — zlsx_editor_mark_recalc_on_load
 * then save, zlsx_editor_save_with_recalc, zlsx_editor_recalculate —
 * rebuild their candidate from the archive as opened and do NOT carry
 * this write: call them before it, or save and re-open (a recorded,
 * pre-existing rule of the transaction's generation model). The
 * record's hidden _zlsxRecoveryN defined names are staged with the
 * workbook plan and appear in zlsx_editor_defined_names_ndjson only
 * after a save. Inherited from
 * the Zig surface and unchanged here: the index read hands model /
 * id / target attributes back raw, so a model name carrying `&`, `<`
 * or `"` reads back entity-escaped, and a tab / LF / CR in it reads
 * back as written here while a conforming XML parser normalizes them
 * to spaces (attribute values) — use plain spaces. */
int32_t zlsx_editor_set_embeddings(zlsx_editor_t * ed,
        const uint8_t * model, size_t model_len,
        uint32_t dim,
        const uint8_t * dtype, size_t dtype_len,
        const zlsx_emb_coverage_v1 * coverages, size_t coverages_len,
        uint32_t flags,
        zlsx_diag_v1 * diag, char * errbuf, size_t errbuf_len);

/* Feature macros — compile-time counterpart of the dlsym probe. */
#define ZLSX_HAS_STRUCTURAL_EDITS 1   /* insert/delete row + column, add/rename/delete sheet, rename_table_column */
#define ZLSX_HAS_PIVOTS           1   /* editor pivots_ndjson */
#define ZLSX_HAS_DEFINED_NAMES    1   /* editor defined_names_ndjson (the read; the writer's add_defined_name predates the macros) */
#define ZLSX_HAS_CONDITIONAL_FORMATS 1   /* editor conditional_formats_ndjson (the read; the writer's add_conditional_format_* predate the macros) */
#define ZLSX_HAS_ANCHORS          1   /* editor anchors_ndjson */
#define ZLSX_HAS_SHEET_PROPS      1   /* editor sheet_props_ndjson + calc_props_ndjson */
#define ZLSX_HAS_SHEET_STATE      1   /* reader sheet_state (S3b slice 10) */
#define ZLSX_HAS_ROWS_FORMULAS    1   /* reader rows_formula_at + rows_formula_ref_at + rows_error_at (S3b slice 11) */
#define ZLSX_HAS_EMBEDDING_WRITE  1   /* editor set_embeddings + zlsx_emb_coverage_v1 (S3c slice 1) */


#ifdef __cplusplus
}
#endif

#endif /* ZLSX_H */
