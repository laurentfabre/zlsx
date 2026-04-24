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
 */
int32_t zlsx_rows_next(zlsx_rows_t         * rows,
                       const zlsx_cell_t ** out_cells,
                       size_t             * out_len,
                       uint8_t            * err_buf,
                       size_t               err_buf_len);

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
 *   -1 → `col_idx` is past the row width
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

#ifdef __cplusplus
}
#endif

#endif /* ZLSX_H */
