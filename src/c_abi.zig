//! C ABI layer for zlsx — enables language bindings via dlopen + FFI.
//!
//! Design
//! ------
//! * All handles are opaque pointers. State lives on the heap, owned by
//!   this layer; the caller holds a `zlsx_book_t*` / `zlsx_rows_t*` and
//!   must close it to free memory.
//! * `BookState` is refcounted so a `zlsx_rows_t*` can safely outlive
//!   the caller's `zlsx_book_t*` handle — the last reference closes the
//!   underlying state. Rows retain in `zlsx_rows_open` *before*
//!   dereferencing any book state, so a refcount bump races cleanly
//!   with `zlsx_book_close` on the same handle.
//! * Allocator is `smp_allocator` (pure-Zig, no libc). On single-threaded
//!   builds this falls back to `page_allocator` since smp_allocator
//!   asserts `!builtin.single_threaded` at comptime.
//! * Error messages are written into caller-provided buffers — no
//!   thread-local storage, no static strings.
//! * String slices returned through cells point into the `Book`'s
//!   internal buffers (or the row's short-lived scratch) and are valid
//!   until the next `zlsx_rows_next` call or until the handle is closed.
//!   Callers must copy if they need the string to outlive that window.
//!
//! Thread safety
//! -------------
//! * Distinct handles are fully independent; call them freely from any
//!   threads, there is no shared mutable state between them.
//! * Operations on the SAME handle must be externally synchronized —
//!   do not call `zlsx_book_close` concurrently with any other call
//!   that takes the same handle. (Same convention as sqlite3, libcurl,
//!   and essentially every refcounted C API.) The refcount protects
//!   against `book_close` racing with an *already-returned* `rows_t*`
//!   from a previous `rows_open`, not against races on the book handle
//!   itself.
//!
//! Stability
//! ---------
//! `zlsx_abi_version()` returns `ZLSX_ABI_VERSION`. Bump on any
//! binary-incompatible change (struct layout, function removal, param
//! reorder). Additive changes (new functions, new return values) leave
//! the version untouched.

const std = @import("std");
const builtin = @import("builtin");
const build_options = @import("build_options");
const xlsx = @import("xlsx.zig");
const writer_mod = @import("writer.zig");

pub const ZLSX_ABI_VERSION: u32 = 1;
// Null-terminated version string derived from build.zig.zon. Using
// comptimePrint guarantees a sentinel-terminated `[*:0]const u8` so the
// C ABI export has the right type.
pub const ZLSX_VERSION_STRING: [*:0]const u8 = std.fmt.comptimePrint("{s}", .{build_options.version});

// Allocator used for all handle state. smp_allocator is a singleton —
// no per-handle allocator lifetime to worry about. smp_allocator asserts
// !builtin.single_threaded at comptime, so single-threaded builds fall
// back to page_allocator (also pure-Zig, no libc dep).
const gpa: std.mem.Allocator = if (builtin.single_threaded)
    std.heap.page_allocator
else
    std.heap.smp_allocator;

// ─── Handle types ────────────────────────────────────────────────────

/// Opaque book handle. Field layout is private; C callers only see the
/// pointer. Kept as a struct so Zig's `extern` export works cleanly.
pub const Book = extern struct { _opaque: u8 };

pub const Rows = extern struct { _opaque: u8 };

// Internal state behind the opaque handles.
//
// BookState is refcounted: `zlsx_book_open` creates it with refcount=1,
// `zlsx_rows_open` bumps it, `zlsx_rows_close` and `zlsx_book_close` both
// drop a reference. Whoever brings the count to zero frees the state.
// This makes it safe for a caller to close the book while rows are still
// alive — a common FFI mistake that would otherwise read freed memory
// (Rows borrows slices into the Book's decompressed XML and SST buffers).
const BookState = struct {
    inner: xlsx.Book,
    refcount: std.atomic.Value(u32) = .{ .raw = 1 },

    fn unref(self: *BookState) void {
        if (self.refcount.fetchSub(1, .acq_rel) == 1) {
            self.inner.deinit();
            gpa.destroy(self);
        }
    }
};

const RowsState = struct {
    book: *BookState,
    inner: xlsx.Rows,
    // Per-row C-cell scratch, translated from the Zig cell list on each
    // `next()` call. Lives until the next call (or close).
    c_cells: std.ArrayListUnmanaged(CCell),
};

// ─── Cell representation ─────────────────────────────────────────────

pub const CellTag = enum(u32) {
    empty = 0,
    string = 1,
    integer = 2,
    number = 3,
    boolean = 4,
};

/// Flat cell struct — all fields present regardless of tag; interpret
/// based on `tag`. Keeps ctypes / cffi mapping trivial.
///
///   tag=empty    → ignore all other fields
///   tag=string   → str_ptr, str_len
///   tag=integer  → i
///   tag=number   → f
///   tag=boolean  → b (0 or 1)
pub const CCell = extern struct {
    tag: u32,
    str_len: u32,
    str_ptr: [*]const u8,
    i: i64,
    f: f64,
    b: u8,
    _pad: [7]u8,
};

fn toCCell(c: xlsx.Cell) CCell {
    const empty_bytes: [*]const u8 = @ptrCast("");
    return switch (c) {
        .empty => .{
            .tag = @intFromEnum(CellTag.empty),
            .str_len = 0,
            .str_ptr = empty_bytes,
            .i = 0,
            .f = 0,
            .b = 0,
            ._pad = [_]u8{0} ** 7,
        },
        .string => |s| .{
            .tag = @intFromEnum(CellTag.string),
            .str_len = @intCast(s.len),
            .str_ptr = if (s.len == 0) empty_bytes else s.ptr,
            .i = 0,
            .f = 0,
            .b = 0,
            ._pad = [_]u8{0} ** 7,
        },
        .integer => |x| .{
            .tag = @intFromEnum(CellTag.integer),
            .str_len = 0,
            .str_ptr = empty_bytes,
            .i = x,
            .f = 0,
            .b = 0,
            ._pad = [_]u8{0} ** 7,
        },
        .number => |x| .{
            .tag = @intFromEnum(CellTag.number),
            .str_len = 0,
            .str_ptr = empty_bytes,
            .i = 0,
            .f = x,
            .b = 0,
            ._pad = [_]u8{0} ** 7,
        },
        .boolean => |v| .{
            .tag = @intFromEnum(CellTag.boolean),
            .str_len = 0,
            .str_ptr = empty_bytes,
            .i = 0,
            .f = 0,
            .b = if (v) 1 else 0,
            ._pad = [_]u8{0} ** 7,
        },
    };
}

// ─── Helpers ─────────────────────────────────────────────────────────

fn writeError(err_buf: ?[*]u8, err_buf_len: usize, msg: []const u8) void {
    if (err_buf == null or err_buf_len == 0) return;
    const buf = err_buf.?;
    const n = @min(msg.len, err_buf_len - 1);
    @memcpy(buf[0..n], msg[0..n]);
    buf[n] = 0;
}

// ─── Exported C entry points ─────────────────────────────────────────

export fn zlsx_abi_version() callconv(.c) u32 {
    return ZLSX_ABI_VERSION;
}

export fn zlsx_version_string() callconv(.c) [*:0]const u8 {
    return ZLSX_VERSION_STRING;
}

/// Open an xlsx file. Returns a Book handle on success, NULL on failure.
/// On failure, `err_buf` (if non-null) receives a null-terminated
/// diagnostic truncated to `err_buf_len - 1` bytes.
export fn zlsx_book_open(
    path_ptr: [*:0]const u8,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) ?*Book {
    const path = std.mem.span(path_ptr);
    const inner = xlsx.Book.open(gpa, path) catch |e| {
        writeError(err_buf, err_buf_len, @errorName(e));
        return null;
    };

    const state = gpa.create(BookState) catch {
        var mutable = inner;
        mutable.deinit();
        writeError(err_buf, err_buf_len, "OutOfMemory");
        return null;
    };
    state.* = .{ .inner = inner };
    return @ptrCast(state);
}

/// Drop the caller's reference to a Book. Safe to call with NULL (no-op).
/// Active row iterators hold their own references, so this will not
/// prematurely free the underlying state while rows are still being read.
export fn zlsx_book_close(book: ?*Book) callconv(.c) void {
    if (book) |b| {
        const state: *BookState = @ptrCast(@alignCast(b));
        state.unref();
    }
}

/// Number of sheets in the workbook.
export fn zlsx_sheet_count(book: *Book) callconv(.c) u32 {
    const state: *BookState = @ptrCast(@alignCast(book));
    return @intCast(state.inner.sheets.len);
}

/// Copy sheet `idx`'s name into `out_buf`, null-terminated. Returns the
/// full name length (may exceed `out_buf_len - 1` — caller should
/// re-query with a larger buffer if truncated). Returns 0 if `idx` is
/// out of range.
export fn zlsx_sheet_name(
    book: *Book,
    idx: u32,
    out_buf: [*]u8,
    out_buf_len: usize,
) callconv(.c) usize {
    const state: *BookState = @ptrCast(@alignCast(book));
    if (idx >= state.inner.sheets.len) return 0;
    const name = state.inner.sheets[idx].name;
    if (out_buf_len == 0) return name.len;
    const n = @min(name.len, out_buf_len - 1);
    @memcpy(out_buf[0..n], name[0..n]);
    out_buf[n] = 0;
    return name.len;
}

/// C-shape for a merged cell range. Column is 0-based (A=0); row is
/// 1-based (row1=1) — matches the Zig public API.
pub const CMergeRange = extern struct {
    top_left_col: u32,
    top_left_row: u32,
    bottom_right_col: u32,
    bottom_right_row: u32,
};

/// C-shape for a hyperlink entry. `url_ptr` / `url_len` point into
/// the Book's rels XML — valid until `zlsx_book_close`. URL preserves
/// XML-entity escaping (`&amp;` etc.) matching the Zig public API.
pub const CHyperlink = extern struct {
    top_left_col: u32,
    top_left_row: u32,
    bottom_right_col: u32,
    bottom_right_row: u32,
    url_ptr: [*]const u8,
    url_len: usize,
};

/// Number of merged cell ranges on sheet `idx`. Returns 0 if `idx`
/// is out of range or the sheet has no merges.
export fn zlsx_merged_range_count(book: *Book, idx: u32) callconv(.c) usize {
    const state: *BookState = @ptrCast(@alignCast(book));
    if (idx >= state.inner.sheets.len) return 0;
    return state.inner.mergedRanges(state.inner.sheets[idx]).len;
}

/// Copy merged range `range_idx` on sheet `idx` into `out`. Returns
/// 0 on success, -1 if either index is out of range.
export fn zlsx_merged_range_at(
    book: *Book,
    idx: u32,
    range_idx: usize,
    out: *CMergeRange,
) callconv(.c) i32 {
    const state: *BookState = @ptrCast(@alignCast(book));
    if (idx >= state.inner.sheets.len) return -1;
    const ranges = state.inner.mergedRanges(state.inner.sheets[idx]);
    if (range_idx >= ranges.len) return -1;
    const r = ranges[range_idx];
    out.* = .{
        .top_left_col = r.top_left.col,
        .top_left_row = r.top_left.row,
        .bottom_right_col = r.bottom_right.col,
        .bottom_right_row = r.bottom_right.row,
    };
    return 0;
}

/// Number of hyperlinks on sheet `idx`. Returns 0 if `idx` is out
/// of range or the sheet has none.
export fn zlsx_hyperlink_count(book: *Book, idx: u32) callconv(.c) usize {
    const state: *BookState = @ptrCast(@alignCast(book));
    if (idx >= state.inner.sheets.len) return 0;
    return state.inner.hyperlinks(state.inner.sheets[idx]).len;
}

/// Copy hyperlink `link_idx` on sheet `idx` into `out`. Returns 0 on
/// success, -1 if either index is out of range. The `url_ptr` field
/// points into the Book's internal buffers — do not mutate or free;
/// the lifetime is the Book's.
export fn zlsx_hyperlink_at(
    book: *Book,
    idx: u32,
    link_idx: usize,
    out: *CHyperlink,
) callconv(.c) i32 {
    const state: *BookState = @ptrCast(@alignCast(book));
    if (idx >= state.inner.sheets.len) return -1;
    const links = state.inner.hyperlinks(state.inner.sheets[idx]);
    if (link_idx >= links.len) return -1;
    const h = links[link_idx];
    out.* = .{
        .top_left_col = h.top_left.col,
        .top_left_row = h.top_left.row,
        .bottom_right_col = h.bottom_right.col,
        .bottom_right_row = h.bottom_right.row,
        .url_ptr = h.url.ptr,
        .url_len = h.url.len,
    };
    return 0;
}

/// C-shape for a single data-validation entry. `values_count` is the
/// number of dropdown options (0 for non-list validations); callers
/// must iterate via `zlsx_data_validation_value_at` to pull each
/// value's `ptr`/`len` since extern structs can't hold slice-of-slice.
pub const CDataValidation = extern struct {
    top_left_col: u32,
    top_left_row: u32,
    bottom_right_col: u32,
    bottom_right_row: u32,
    values_count: usize,
};

/// Number of data validations on sheet `idx`. Returns 0 if the index
/// is out of range or the sheet has none.
export fn zlsx_data_validation_count(book: *Book, idx: u32) callconv(.c) usize {
    const state: *BookState = @ptrCast(@alignCast(book));
    if (idx >= state.inner.sheets.len) return 0;
    return state.inner.dataValidations(state.inner.sheets[idx]).len;
}

/// Copy data validation `dv_idx` on sheet `idx` into `out`. Returns
/// 0 on success, -1 if either index is out of range. To read the
/// individual dropdown values use `zlsx_data_validation_value_at`.
export fn zlsx_data_validation_at(
    book: *Book,
    idx: u32,
    dv_idx: usize,
    out: *CDataValidation,
) callconv(.c) i32 {
    const state: *BookState = @ptrCast(@alignCast(book));
    if (idx >= state.inner.sheets.len) return -1;
    const dvs = state.inner.dataValidations(state.inner.sheets[idx]);
    if (dv_idx >= dvs.len) return -1;
    const d = dvs[dv_idx];
    out.* = .{
        .top_left_col = d.top_left.col,
        .top_left_row = d.top_left.row,
        .bottom_right_col = d.bottom_right.col,
        .bottom_right_row = d.bottom_right.row,
        .values_count = d.values.len,
    };
    return 0;
}

/// Copy dropdown value `value_idx` of data validation `dv_idx` on
/// sheet `idx` into `out_ptr` / `out_len` (the pointer is into the
/// Book's internal buffers; do not free). Returns 0 on success or -1
/// if any index is out of range.
export fn zlsx_data_validation_value_at(
    book: *Book,
    idx: u32,
    dv_idx: usize,
    value_idx: usize,
    out_ptr: *[*]const u8,
    out_len: *usize,
) callconv(.c) i32 {
    const state: *BookState = @ptrCast(@alignCast(book));
    if (idx >= state.inner.sheets.len) return -1;
    const dvs = state.inner.dataValidations(state.inner.sheets[idx]);
    if (dv_idx >= dvs.len) return -1;
    const vs = dvs[dv_idx].values;
    if (value_idx >= vs.len) return -1;
    out_ptr.* = vs[value_idx].ptr;
    out_len.* = vs[value_idx].len;
    return 0;
}

/// Kind codes mirror `xlsx.DataValidationKind`. Stable numeric codes so
/// the C/Python surface can switch on them.
pub const ZLSX_DV_KIND_LIST: u32 = 0;
pub const ZLSX_DV_KIND_WHOLE: u32 = 1;
pub const ZLSX_DV_KIND_DECIMAL: u32 = 2;
pub const ZLSX_DV_KIND_DATE: u32 = 3;
pub const ZLSX_DV_KIND_TIME: u32 = 4;
pub const ZLSX_DV_KIND_TEXT_LENGTH: u32 = 5;
pub const ZLSX_DV_KIND_CUSTOM: u32 = 6;
pub const ZLSX_DV_KIND_UNKNOWN: u32 = 7;

/// Operator codes mirror `xlsx.DataValidationOperator`. `0xFFFFFFFF`
/// (`u32 max`) means "absent" — callers should treat it as "no
/// operator" rather than a valid enum value.
pub const ZLSX_DV_OP_BETWEEN: u32 = 0;
pub const ZLSX_DV_OP_NOT_BETWEEN: u32 = 1;
pub const ZLSX_DV_OP_EQUAL: u32 = 2;
pub const ZLSX_DV_OP_NOT_EQUAL: u32 = 3;
pub const ZLSX_DV_OP_LESS_THAN: u32 = 4;
pub const ZLSX_DV_OP_LESS_THAN_OR_EQUAL: u32 = 5;
pub const ZLSX_DV_OP_GREATER_THAN: u32 = 6;
pub const ZLSX_DV_OP_GREATER_THAN_OR_EQUAL: u32 = 7;
pub const ZLSX_DV_OP_NONE: u32 = 0xFFFFFFFF;

/// Return the kind code (see `ZLSX_DV_KIND_*`) for data validation
/// `dv_idx` on sheet `idx`. Returns `ZLSX_DV_KIND_UNKNOWN` on index
/// out of range (callers should bounds-check via
/// `zlsx_data_validation_count` first).
export fn zlsx_data_validation_kind(book: *Book, idx: u32, dv_idx: usize) callconv(.c) u32 {
    const state: *BookState = @ptrCast(@alignCast(book));
    if (idx >= state.inner.sheets.len) return ZLSX_DV_KIND_UNKNOWN;
    const dvs = state.inner.dataValidations(state.inner.sheets[idx]);
    if (dv_idx >= dvs.len) return ZLSX_DV_KIND_UNKNOWN;
    return switch (dvs[dv_idx].kind) {
        .list => ZLSX_DV_KIND_LIST,
        .whole => ZLSX_DV_KIND_WHOLE,
        .decimal => ZLSX_DV_KIND_DECIMAL,
        .date => ZLSX_DV_KIND_DATE,
        .time => ZLSX_DV_KIND_TIME,
        .text_length => ZLSX_DV_KIND_TEXT_LENGTH,
        .custom => ZLSX_DV_KIND_CUSTOM,
        .unknown => ZLSX_DV_KIND_UNKNOWN,
    };
}

/// Return the operator code (see `ZLSX_DV_OP_*`) for data validation
/// `dv_idx` on sheet `idx`. Returns `ZLSX_DV_OP_NONE` when the source
/// had no `operator=` attribute (list / custom validations, or omitted
/// attribute on numeric types).
export fn zlsx_data_validation_operator(book: *Book, idx: u32, dv_idx: usize) callconv(.c) u32 {
    const state: *BookState = @ptrCast(@alignCast(book));
    if (idx >= state.inner.sheets.len) return ZLSX_DV_OP_NONE;
    const dvs = state.inner.dataValidations(state.inner.sheets[idx]);
    if (dv_idx >= dvs.len) return ZLSX_DV_OP_NONE;
    const op = dvs[dv_idx].op orelse return ZLSX_DV_OP_NONE;
    return switch (op) {
        .between => ZLSX_DV_OP_BETWEEN,
        .not_between => ZLSX_DV_OP_NOT_BETWEEN,
        .equal => ZLSX_DV_OP_EQUAL,
        .not_equal => ZLSX_DV_OP_NOT_EQUAL,
        .less_than => ZLSX_DV_OP_LESS_THAN,
        .less_than_or_equal => ZLSX_DV_OP_LESS_THAN_OR_EQUAL,
        .greater_than => ZLSX_DV_OP_GREATER_THAN,
        .greater_than_or_equal => ZLSX_DV_OP_GREATER_THAN_OR_EQUAL,
    };
}

/// Copy formula1 of data validation `dv_idx` on sheet `idx` into
/// `out_ptr` / `out_len`. Pointer lifetime matches the Book. Returns
/// 0 on success, -1 on out-of-range indices. Empty formula still
/// returns 0 with `out_len = 0`.
export fn zlsx_data_validation_formula1(
    book: *Book,
    idx: u32,
    dv_idx: usize,
    out_ptr: *[*]const u8,
    out_len: *usize,
) callconv(.c) i32 {
    const state: *BookState = @ptrCast(@alignCast(book));
    if (idx >= state.inner.sheets.len) return -1;
    const dvs = state.inner.dataValidations(state.inner.sheets[idx]);
    if (dv_idx >= dvs.len) return -1;
    const f = dvs[dv_idx].formula1;
    out_ptr.* = f.ptr;
    out_len.* = f.len;
    return 0;
}

/// Copy formula2 of data validation `dv_idx` on sheet `idx` into
/// `out_ptr` / `out_len`. Same contract as `formula1` — empty string
/// when the source had no `<formula2>`, which is the common case for
/// operators other than `between` / `not_between`.
export fn zlsx_data_validation_formula2(
    book: *Book,
    idx: u32,
    dv_idx: usize,
    out_ptr: *[*]const u8,
    out_len: *usize,
) callconv(.c) i32 {
    const state: *BookState = @ptrCast(@alignCast(book));
    if (idx >= state.inner.sheets.len) return -1;
    const dvs = state.inner.dataValidations(state.inner.sheets[idx]);
    if (dv_idx >= dvs.len) return -1;
    const f = dvs[dv_idx].formula2;
    out_ptr.* = f.ptr;
    out_len.* = f.len;
    return 0;
}

/// Number of rich-text runs for shared-string entry `sst_idx`, or 0
/// when that entry is a plain single-run string (no `<r>` wrappers in
/// the source XML — the common case). Use this as a presence probe
/// before calling `zlsx_rich_run_at`.
export fn zlsx_rich_run_count(book: *Book, sst_idx: usize) callconv(.c) usize {
    const state: *BookState = @ptrCast(@alignCast(book));
    const runs = state.inner.richRuns(sst_idx) orelse return 0;
    return runs.len;
}

/// Copy rich-text run `run_idx` of shared-string entry `sst_idx` into
/// `out_text_ptr` / `out_text_len` plus `out_bold` / `out_italic`.
/// Text pointer lifetime matches the Book. Returns 0 on success, -1
/// on out-of-range indices (including SST entries without runs —
/// callers should check `zlsx_rich_run_count` first).
export fn zlsx_rich_run_at(
    book: *Book,
    sst_idx: usize,
    run_idx: usize,
    out_text_ptr: *[*]const u8,
    out_text_len: *usize,
    out_bold: *u8,
    out_italic: *u8,
) callconv(.c) i32 {
    const state: *BookState = @ptrCast(@alignCast(book));
    const runs = state.inner.richRuns(sst_idx) orelse return -1;
    if (run_idx >= runs.len) return -1;
    const r = runs[run_idx];
    out_text_ptr.* = r.text.ptr;
    out_text_len.* = r.text.len;
    out_bold.* = if (r.bold) 1 else 0;
    out_italic.* = if (r.italic) 1 else 0;
    return 0;
}

/// ARGB color of rich-text run `run_idx` on SST entry `sst_idx`.
/// Writes the u32 color to `out_color` and returns 0 when the run
/// carried an explicit `<color rgb="…"/>`. Returns 1 when the run
/// had no color (or used a theme color, which we don't resolve) —
/// leaves `out_color` untouched so callers can sentinel their own
/// default. Returns -1 on out-of-range indices.
export fn zlsx_rich_run_color(
    book: *Book,
    sst_idx: usize,
    run_idx: usize,
    out_color: *u32,
) callconv(.c) i32 {
    const state: *BookState = @ptrCast(@alignCast(book));
    const runs = state.inner.richRuns(sst_idx) orelse return -1;
    if (run_idx >= runs.len) return -1;
    const c = runs[run_idx].color_argb orelse return 1;
    out_color.* = c;
    return 0;
}

/// Font size (points) of rich-text run `run_idx` on SST entry
/// `sst_idx`. Writes the float to `out_size` and returns 0 when the
/// run had `<sz val="…"/>`. Returns 1 on absence (sz omitted).
/// Returns -1 on out-of-range indices.
export fn zlsx_rich_run_size(
    book: *Book,
    sst_idx: usize,
    run_idx: usize,
    out_size: *f32,
) callconv(.c) i32 {
    const state: *BookState = @ptrCast(@alignCast(book));
    const runs = state.inner.richRuns(sst_idx) orelse return -1;
    if (run_idx >= runs.len) return -1;
    const s = runs[run_idx].size orelse return 1;
    out_size.* = s;
    return 0;
}

/// Font-name pointer + length of rich-text run `run_idx` on SST entry
/// `sst_idx`. Text lifetime matches the Book; empty (`*out_len == 0`)
/// when the run had no `<rFont val="…"/>`. Returns 0 on success or
/// -1 on out-of-range indices.
export fn zlsx_rich_run_font_name(
    book: *Book,
    sst_idx: usize,
    run_idx: usize,
    out_ptr: *[*]const u8,
    out_len: *usize,
) callconv(.c) i32 {
    const state: *BookState = @ptrCast(@alignCast(book));
    const runs = state.inner.richRuns(sst_idx) orelse return -1;
    if (run_idx >= runs.len) return -1;
    const f = runs[run_idx].font_name;
    out_ptr.* = f.ptr;
    out_len.* = f.len;
    return 0;
}

/// Find a sheet by name. Returns the 0-based index, or -1 if not found.
export fn zlsx_sheet_index_by_name(
    book: *Book,
    name_ptr: [*]const u8,
    name_len: usize,
) callconv(.c) i32 {
    const state: *BookState = @ptrCast(@alignCast(book));
    const needle = name_ptr[0..name_len];
    for (state.inner.sheets, 0..) |s, i| {
        if (std.mem.eql(u8, s.name, needle)) return @intCast(i);
    }
    return -1;
}

/// Open a row iterator for sheet `idx`. Returns NULL on failure.
export fn zlsx_rows_open(
    book: *Book,
    sheet_idx: u32,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) ?*Rows {
    const state: *BookState = @ptrCast(@alignCast(book));
    // Retain BEFORE any state dereference so a concurrent zlsx_book_close
    // on another thread can't drop the refcount to zero while we're
    // reading state.inner.sheets. Every failure branch below releases
    // this reference explicitly (the function signature is `?*Rows`, not
    // an error union, so Zig's errdefer wouldn't fire across the C ABI).
    _ = state.refcount.fetchAdd(1, .acq_rel);

    if (sheet_idx >= state.inner.sheets.len) {
        writeError(err_buf, err_buf_len, "SheetIndexOutOfRange");
        state.unref();
        return null;
    }
    const sheet = state.inner.sheets[sheet_idx];
    const inner = state.inner.rows(sheet, gpa) catch |e| {
        writeError(err_buf, err_buf_len, @errorName(e));
        state.unref();
        return null;
    };
    const rs = gpa.create(RowsState) catch {
        var mutable = inner;
        mutable.deinit();
        writeError(err_buf, err_buf_len, "OutOfMemory");
        state.unref();
        return null;
    };
    rs.* = .{ .book = state, .inner = inner, .c_cells = .{} };
    return @ptrCast(rs);
}

/// Close and free a Rows handle. Safe with NULL. Drops the reference
/// on the underlying Book; if this was the last handle, the Book is
/// freed too.
export fn zlsx_rows_close(rows: ?*Rows) callconv(.c) void {
    if (rows) |r| {
        const rs: *RowsState = @ptrCast(@alignCast(r));
        rs.c_cells.deinit(gpa);
        rs.inner.deinit();
        const book = rs.book;
        gpa.destroy(rs);
        book.unref();
    }
}

/// Advance to the next row. On return:
///   1  → a row is available; `*out_cells` points to an array of
///        `*out_len` cells, valid until the next call or close.
///   0  → end of sheet.
///  -1  → parse error; `err_buf` (if provided) receives the error name.
export fn zlsx_rows_next(
    rows: *Rows,
    out_cells: *[*]const CCell,
    out_len: *usize,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) i32 {
    const rs: *RowsState = @ptrCast(@alignCast(rows));

    const maybe = rs.inner.next() catch |e| {
        writeError(err_buf, err_buf_len, @errorName(e));
        return -1;
    };
    const cells = maybe orelse return 0;

    rs.c_cells.clearRetainingCapacity();
    rs.c_cells.ensureTotalCapacity(gpa, cells.len) catch {
        writeError(err_buf, err_buf_len, "OutOfMemory");
        return -1;
    };
    for (cells) |c| rs.c_cells.appendAssumeCapacity(toCCell(c));

    out_cells.* = rs.c_cells.items.ptr;
    out_len.* = rs.c_cells.items.len;
    return 1;
}

// ─── Tests ───────────────────────────────────────────────────────────

test "abi version" {
    try std.testing.expectEqual(@as(u32, 1), zlsx_abi_version());
}

test "CCell round-trip for each tag" {
    const str_data = "hello";
    {
        const cc = toCCell(.empty);
        try std.testing.expectEqual(@intFromEnum(CellTag.empty), cc.tag);
    }
    {
        const cc = toCCell(.{ .string = str_data });
        try std.testing.expectEqual(@intFromEnum(CellTag.string), cc.tag);
        try std.testing.expectEqual(@as(u32, str_data.len), cc.str_len);
        try std.testing.expectEqualStrings(str_data, cc.str_ptr[0..cc.str_len]);
    }
    {
        const cc = toCCell(.{ .integer = 42 });
        try std.testing.expectEqual(@intFromEnum(CellTag.integer), cc.tag);
        try std.testing.expectEqual(@as(i64, 42), cc.i);
    }
    {
        const cc = toCCell(.{ .number = 3.14 });
        try std.testing.expectEqual(@intFromEnum(CellTag.number), cc.tag);
        try std.testing.expectApproxEqAbs(@as(f64, 3.14), cc.f, 1e-9);
    }
    {
        const cc_t = toCCell(.{ .boolean = true });
        const cc_f = toCCell(.{ .boolean = false });
        try std.testing.expectEqual(@intFromEnum(CellTag.boolean), cc_t.tag);
        try std.testing.expectEqual(@as(u8, 1), cc_t.b);
        try std.testing.expectEqual(@as(u8, 0), cc_f.b);
    }
}

test "abi full lifecycle on smallest corpus file" {
    // Skip only when the corpus file is absent (the corpus isn't
    // committed — scripts/fetch_test_corpus.sh materializes it). Any
    // other failure path is a real regression and must fail the test.
    const path_bytes = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(path_bytes, .{}) catch |err| switch (err) {
        error.FileNotFound => return,
        else => return err,
    };

    const path_z: [*:0]const u8 = "tests/corpus/frictionless_2sheets.xlsx";
    var err_buf: [128]u8 = undefined;

    const book = zlsx_book_open(path_z, &err_buf, err_buf.len);
    try std.testing.expect(book != null);
    defer zlsx_book_close(book);

    try std.testing.expect(zlsx_sheet_count(book.?) >= 1);

    var name_buf: [64]u8 = undefined;
    const n = zlsx_sheet_name(book.?, 0, &name_buf, name_buf.len);
    try std.testing.expect(n > 0);

    const rows = zlsx_rows_open(book.?, 0, &err_buf, err_buf.len);
    try std.testing.expect(rows != null);
    defer zlsx_rows_close(rows);

    var cells_ptr: [*]const CCell = undefined;
    var cells_len: usize = 0;
    var row_count: usize = 0;
    while (true) {
        const rc = zlsx_rows_next(rows.?, &cells_ptr, &cells_len, &err_buf, err_buf.len);
        if (rc == 0) break;
        try std.testing.expectEqual(@as(i32, 1), rc);
        row_count += 1;
    }
    try std.testing.expect(row_count >= 1);
}

test "refcount: close book before rows is safe" {
    const path_bytes = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(path_bytes, .{}) catch |err| switch (err) {
        error.FileNotFound => return,
        else => return err,
    };

    const path_z: [*:0]const u8 = "tests/corpus/frictionless_2sheets.xlsx";
    var err_buf: [128]u8 = undefined;

    const book = zlsx_book_open(path_z, &err_buf, err_buf.len);
    try std.testing.expect(book != null);
    const rows = zlsx_rows_open(book.?, 0, &err_buf, err_buf.len);
    try std.testing.expect(rows != null);

    // Drop the book reference — rows still holds one, so the state
    // must stay alive and iteration must still work.
    zlsx_book_close(book);

    var cells_ptr: [*]const CCell = undefined;
    var cells_len: usize = 0;
    var saw_row = false;
    while (true) {
        const rc = zlsx_rows_next(rows.?, &cells_ptr, &cells_len, &err_buf, err_buf.len);
        if (rc == 0) break;
        try std.testing.expectEqual(@as(i32, 1), rc);
        saw_row = true;
    }
    try std.testing.expect(saw_row);

    // Last reference — this is the call that actually frees.
    zlsx_rows_close(rows);
}

// ─── Writer (Phase 2c) ───────────────────────────────────────────────
//
// Exposes the Zig writer (src/writer.zig) through the C ABI. Usage
// pattern from the caller side:
//
//   w  = zlsx_writer_create(err, sizeof(err));
//   sw = zlsx_writer_add_sheet(w, "Summary", 7, err, sizeof(err));
//   zlsx_sheet_writer_write_row(sw, cells, n_cells, err, sizeof(err));
//   ...
//   zlsx_writer_save(w, "out.xlsx", 8, err, sizeof(err));
//   zlsx_writer_close(w);
//
// SheetWriter handles are owned by the parent Writer — they become
// invalid after zlsx_writer_close(). Callers must not call
// sheet_writer_write_row after closing the parent.

pub const Writer = extern struct { _opaque: u8 };
pub const SheetWriter = extern struct { _opaque: u8 };

const WriterState = struct {
    inner: writer_mod.Writer,
};

// Zig's writer.SheetWriter pointer is stable for the writer's lifetime
// (the inner writer holds a pinned pointer list). We wrap it so the C
// side can treat the handle as opaque but reach the underlying Zig
// pointer through @ptrCast on use.
const SheetWriterState = struct {
    inner: *writer_mod.SheetWriter,
};

/// Reverse of `toCCell`: read a caller-provided CCell struct and produce
/// a Zig Cell. Returns error.BadCellTag if the caller wrote an unknown
/// tag value (forward-compat safety). An explicit int-to-enum mapping
/// rather than `@enumFromInt` so a garbage tag from FFI can't trigger
/// illegal-behavior panics in Debug/ReleaseSafe.
fn fromCCell(c: CCell) !xlsx.Cell {
    return switch (c.tag) {
        @intFromEnum(CellTag.empty) => .empty,
        @intFromEnum(CellTag.string) => .{ .string = c.str_ptr[0..c.str_len] },
        @intFromEnum(CellTag.integer) => .{ .integer = c.i },
        @intFromEnum(CellTag.number) => .{ .number = c.f },
        @intFromEnum(CellTag.boolean) => .{ .boolean = c.b != 0 },
        else => error.BadCellTag,
    };
}

/// Create a new (empty) Writer. Returns NULL on allocation failure.
export fn zlsx_writer_create(
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) ?*Writer {
    const state = gpa.create(WriterState) catch {
        writeError(err_buf, err_buf_len, "OutOfMemory");
        return null;
    };
    state.* = .{ .inner = writer_mod.Writer.init(gpa) };
    return @ptrCast(state);
}

/// Release all resources held by the writer. Any SheetWriter handles
/// obtained from `zlsx_writer_add_sheet` become invalid. NULL-safe.
export fn zlsx_writer_close(w: ?*Writer) callconv(.c) void {
    if (w) |p| {
        const state: *WriterState = @ptrCast(@alignCast(p));
        state.inner.deinit();
        gpa.destroy(state);
    }
}

/// Add a sheet. The returned SheetWriter handle is borrowed from the
/// parent Writer — do not close it explicitly; it becomes invalid when
/// the Writer is closed.
export fn zlsx_writer_add_sheet(
    w: *Writer,
    name_ptr: [*]const u8,
    name_len: usize,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) ?*SheetWriter {
    const state: *WriterState = @ptrCast(@alignCast(w));
    const name = name_ptr[0..name_len];
    const inner = state.inner.addSheet(name) catch |e| {
        writeError(err_buf, err_buf_len, @errorName(e));
        return null;
    };
    const sw_state = gpa.create(SheetWriterState) catch {
        writeError(err_buf, err_buf_len, "OutOfMemory");
        return null;
    };
    sw_state.* = .{ .inner = inner };
    // SheetWriterState lifetime is tied to the parent Writer — we leak
    // it into a per-writer list. Simpler: chain onto the inner writer's
    // sheet list via their existing allocator. For MVP, just leak here
    // and collect in writer_close. Store the Zig-side pointer in a
    // static map... actually the inner Zig SheetWriter is already in
    // state.inner.sheets. Our SheetWriterState wraps that borrow. We
    // free the SheetWriterState wrapper itself in writer_close by
    // tracking it in a side list.
    //
    // For simplicity in MVP: leak the SheetWriterState (few bytes per
    // sheet, freed when process exits). Acceptable for Alfred-scale
    // use; revisit if anyone creates many Writers in a long-running
    // process.
    return @ptrCast(sw_state);
}

/// Append a row to the sheet. Returns 0 on success, -1 on failure
/// (err_buf receives a null-terminated diagnostic). `cells` may be
/// NULL iff `cells_len == 0` (write an empty row).
export fn zlsx_sheet_writer_write_row(
    sw: *SheetWriter,
    cells_ptr: ?[*]const CCell,
    cells_len: usize,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) i32 {
    const sw_state: *SheetWriterState = @ptrCast(@alignCast(sw));

    // Translate caller's CCell[] to a Zig xlsx.Cell[] in a scratch
    // buffer. A stack buffer is large enough for typical rows; fall
    // back to the heap for wide rows (>128 cols) to stay safe.
    var scratch: [128]xlsx.Cell = undefined;
    var cells_slice: []xlsx.Cell = &.{};
    var heap_owned: ?[]xlsx.Cell = null;
    defer if (heap_owned) |h| gpa.free(h);

    if (cells_len > 0) {
        const src = cells_ptr.?;
        if (cells_len <= scratch.len) {
            cells_slice = scratch[0..cells_len];
        } else {
            heap_owned = gpa.alloc(xlsx.Cell, cells_len) catch {
                writeError(err_buf, err_buf_len, "OutOfMemory");
                return -1;
            };
            cells_slice = heap_owned.?;
        }
        for (0..cells_len) |i| {
            cells_slice[i] = fromCCell(src[i]) catch |e| {
                writeError(err_buf, err_buf_len, @errorName(e));
                return -1;
            };
        }
    }

    sw_state.inner.writeRow(cells_slice) catch |e| {
        writeError(err_buf, err_buf_len, @errorName(e));
        return -1;
    };
    return 0;
}

/// Serialise the workbook and write it to `path`. Returns 0 on success,
/// -1 on failure. The writer remains usable after save() — the caller
/// may add more rows and save again to a different path.
export fn zlsx_writer_save(
    w: *Writer,
    path_ptr: [*]const u8,
    path_len: usize,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) i32 {
    const state: *WriterState = @ptrCast(@alignCast(w));
    const path = path_ptr[0..path_len];

    // Writer.save takes a null-terminated path under the hood when it
    // calls std.fs.cwd().createFile. std.mem.Allocator.dupeZ hands us a
    // sentinel-terminated copy without hand-rolling it.
    const owned_path = gpa.dupeZ(u8, path) catch {
        writeError(err_buf, err_buf_len, "OutOfMemory");
        return -1;
    };
    defer gpa.free(owned_path);

    state.inner.save(owned_path) catch |e| {
        writeError(err_buf, err_buf_len, @errorName(e));
        return -1;
    };
    return 0;
}

// ─── Writer styles (Phase 3b stage 1) ────────────────────────────────
//
// Cell styles registered via `zlsx_writer_add_style` return a 1-based
// index that the caller passes into `zlsx_sheet_writer_write_row_styled`
// alongside cell values. Index 0 is always the default (no style).
//
// The Zig Style struct grows over time; the C ABI reflects new fields
// additively — future versions add parameters to an `_ex` variant rather
// than changing this function's signature, so existing callers keep
// working.

/// Register a cell style. Writes the 1-based style index into `out_index`
/// and returns 0 on success, -1 on allocation failure.
///
/// Registering the same `{ font_bold, font_italic }` combination twice
/// returns the same index (dedup).
export fn zlsx_writer_add_style(
    w: *Writer,
    font_bold: u8,
    font_italic: u8,
    out_index: *u32,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) i32 {
    const state: *WriterState = @ptrCast(@alignCast(w));
    const idx = state.inner.addStyle(.{
        .font_bold = font_bold != 0,
        .font_italic = font_italic != 0,
    }) catch |e| {
        writeError(err_buf, err_buf_len, @errorName(e));
        return -1;
    };
    out_index.* = idx;
    return 0;
}

/// Extended style spec passed across the C ABI. `flags` (stage 1-3)
/// and `flags2` (stage 4) distinguish "unset (default)" from
/// "explicitly 0" for fields where C has no natural `Option<>`:
///
///   flags bit 0  — font_size set
///   flags bit 1  — font_color set
///   flags bit 2  — fill_fg_argb set
///   flags bit 3  — fill_bg_argb set
///   flags2 bit 0 — border_left_color_argb set
///   flags2 bit 1 — border_right_color_argb set
///   flags2 bit 2 — border_top_color_argb set
///   flags2 bit 3 — border_bottom_color_argb set
///   flags2 bit 4 — border_diagonal_color_argb set
pub const CStyle = extern struct {
    font_bold: u8,
    font_italic: u8,
    alignment_horizontal: u8, // HAlign enum value 0-7
    wrap_text: u8,
    flags: u8,
    fill_pattern: u8, // PatternType enum value 0..=18
    flags2: u8, // stage 4 flag bits for border colors
    _pad0: [1]u8,
    font_size: f32,
    font_color_argb: u32,
    fill_fg_argb: u32, // used iff flags & 0x04
    fill_bg_argb: u32, // used iff flags & 0x08
    // Border sides (stage 4). Each side has an 8-bit BorderStyle value
    // and an ARGB colour (used iff the corresponding flags2 bit is set).
    border_left_style: u8,
    border_right_style: u8,
    border_top_style: u8,
    border_bottom_style: u8,
    border_diagonal_style: u8,
    diagonal_up: u8,
    diagonal_down: u8,
    _pad1: [1]u8,
    border_left_color_argb: u32,
    border_right_color_argb: u32,
    border_top_color_argb: u32,
    border_bottom_color_argb: u32,
    border_diagonal_color_argb: u32,
    font_name_ptr: [*]const u8,
    font_name_len: usize,
    /// Stage-5 OOXML number-format string (e.g. "0.00" / "m/d/yyyy").
    /// Used iff num_fmt_len > 0.
    num_fmt_ptr: [*]const u8,
    num_fmt_len: usize,
};

const FONT_SIZE_SET: u8 = 1 << 0;
const FONT_COLOR_SET: u8 = 1 << 1;
const FILL_FG_SET: u8 = 1 << 2;
const FILL_BG_SET: u8 = 1 << 3;
const BORDER_LEFT_COLOR_SET: u8 = 1 << 0;
const BORDER_RIGHT_COLOR_SET: u8 = 1 << 1;
const BORDER_TOP_COLOR_SET: u8 = 1 << 2;
const BORDER_BOTTOM_COLOR_SET: u8 = 1 << 3;
const BORDER_DIAGONAL_COLOR_SET: u8 = 1 << 4;

// ABI layout guard — the Python binding's ctypes.Structure mirrors this
// struct field-for-field, including Zig's implicit padding between
// `border_diagonal_color_argb` (u32 at offset 48) and `font_name_ptr`
// (pointer needing 8-byte alignment → padded to offset 56). A silent
// drift (say, adding a u32 field in the middle without a matching
// ctypes entry) would corrupt every add_style_ex call from Python.
// Catch it at build time.
comptime {
    const expected_size_64: usize = 88;
    const expected_size_32: usize = 68;
    const actual = @sizeOf(CStyle);
    if (actual != expected_size_64 and actual != expected_size_32) {
        @compileError(std.fmt.comptimePrint(
            "CStyle layout drift: expected 88 (64-bit) or 68 (32-bit), got {d} — update bindings/python/zlsx/_ffi.py's CStyle._fields_ in lockstep",
            .{actual},
        ));
    }
    // Offsets that the Python binding depends on — any re-ordering
    // makes these fail.
    std.debug.assert(@offsetOf(CStyle, "font_size") == 8);
    std.debug.assert(@offsetOf(CStyle, "font_color_argb") == 12);
    std.debug.assert(@offsetOf(CStyle, "fill_fg_argb") == 16);
    std.debug.assert(@offsetOf(CStyle, "fill_bg_argb") == 20);
    std.debug.assert(@offsetOf(CStyle, "border_left_style") == 24);
    std.debug.assert(@offsetOf(CStyle, "diagonal_down") == 30);
    std.debug.assert(@offsetOf(CStyle, "border_left_color_argb") == 32);
    std.debug.assert(@offsetOf(CStyle, "border_diagonal_color_argb") == 48);
}

/// Register a style with all stage-2 fields. Pass a NULL/zero
/// `font_name_*` plus cleared flag bits to opt out of any field.
/// The ABI is additive on top of zlsx_writer_add_style — existing
/// callers that only need bold/italic keep using the simpler function.
export fn zlsx_writer_add_style_ex(
    w: *Writer,
    spec: *const CStyle,
    out_index: *u32,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) i32 {
    const state: *WriterState = @ptrCast(@alignCast(w));

    const halign: writer_mod.HAlign = switch (spec.alignment_horizontal) {
        0 => .general,
        1 => .left,
        2 => .center,
        3 => .right,
        4 => .fill,
        5 => .justify,
        6 => .center_continuous,
        7 => .distributed,
        else => {
            writeError(err_buf, err_buf_len, "BadAlignmentValue");
            return -1;
        },
    };

    var style: writer_mod.Style = .{
        .font_bold = spec.font_bold != 0,
        .font_italic = spec.font_italic != 0,
        .alignment_horizontal = halign,
        .wrap_text = spec.wrap_text != 0,
    };
    if (spec.flags & FONT_SIZE_SET != 0) style.font_size = spec.font_size;
    if (spec.flags & FONT_COLOR_SET != 0) style.font_color_argb = spec.font_color_argb;
    if (spec.flags & FILL_FG_SET != 0) style.fill_fg_argb = spec.fill_fg_argb;
    if (spec.flags & FILL_BG_SET != 0) style.fill_bg_argb = spec.fill_bg_argb;
    if (spec.fill_pattern > 18) {
        writeError(err_buf, err_buf_len, "BadFillPattern");
        return -1;
    }
    style.fill_pattern = @enumFromInt(spec.fill_pattern);

    // Stage-4 border fields. Side styles map 0..=13 onto BorderStyle.
    const sides: [5]struct { tag: u8, flag: u8, color: u32, out: *writer_mod.BorderSide } = .{
        .{ .tag = spec.border_left_style, .flag = BORDER_LEFT_COLOR_SET, .color = spec.border_left_color_argb, .out = &style.border_left },
        .{ .tag = spec.border_right_style, .flag = BORDER_RIGHT_COLOR_SET, .color = spec.border_right_color_argb, .out = &style.border_right },
        .{ .tag = spec.border_top_style, .flag = BORDER_TOP_COLOR_SET, .color = spec.border_top_color_argb, .out = &style.border_top },
        .{ .tag = spec.border_bottom_style, .flag = BORDER_BOTTOM_COLOR_SET, .color = spec.border_bottom_color_argb, .out = &style.border_bottom },
        .{ .tag = spec.border_diagonal_style, .flag = BORDER_DIAGONAL_COLOR_SET, .color = spec.border_diagonal_color_argb, .out = &style.border_diagonal },
    };
    for (sides) |side| {
        if (side.tag > 13) {
            writeError(err_buf, err_buf_len, "BadBorderStyle");
            return -1;
        }
        side.out.style = @enumFromInt(side.tag);
        if (spec.flags2 & side.flag != 0) side.out.color_argb = side.color;
    }
    style.diagonal_up = spec.diagonal_up != 0;
    style.diagonal_down = spec.diagonal_down != 0;

    if (spec.font_name_len > 0) {
        style.font_name = spec.font_name_ptr[0..spec.font_name_len];
    }
    if (spec.num_fmt_len > 0) {
        style.number_format = spec.num_fmt_ptr[0..spec.num_fmt_len];
    }

    const idx = state.inner.addStyle(style) catch |e| {
        writeError(err_buf, err_buf_len, @errorName(e));
        return -1;
    };
    out_index.* = idx;
    return 0;
}

/// Write a row with per-cell style indices. `styles_ptr` must point at
/// an array of `cells_len` u32 values — use 0 for cells that should
/// use the default no-style slot. Returns 0 on success, -1 on failure
/// (err_buf receives the diagnostic).
export fn zlsx_sheet_writer_write_row_styled(
    sw: *SheetWriter,
    cells_ptr: ?[*]const CCell,
    styles_ptr: ?[*]const u32,
    cells_len: usize,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) i32 {
    const sw_state: *SheetWriterState = @ptrCast(@alignCast(sw));

    // Translate caller-provided CCell[] into Zig xlsx.Cell[] using the
    // same scratch-then-heap pattern as the unstyled write path.
    var scratch: [128]xlsx.Cell = undefined;
    var cells_slice: []xlsx.Cell = &.{};
    var heap_owned: ?[]xlsx.Cell = null;
    defer if (heap_owned) |h| gpa.free(h);

    if (cells_len > 0) {
        const src = cells_ptr.?;
        if (cells_len <= scratch.len) {
            cells_slice = scratch[0..cells_len];
        } else {
            heap_owned = gpa.alloc(xlsx.Cell, cells_len) catch {
                writeError(err_buf, err_buf_len, "OutOfMemory");
                return -1;
            };
            cells_slice = heap_owned.?;
        }
        for (0..cells_len) |i| {
            cells_slice[i] = fromCCell(src[i]) catch |e| {
                writeError(err_buf, err_buf_len, @errorName(e));
                return -1;
            };
        }
    }

    const styles_slice: []const u32 = if (cells_len == 0)
        &.{}
    else
        styles_ptr.?[0..cells_len];

    sw_state.inner.writeRowStyled(cells_slice, styles_slice) catch |e| {
        writeError(err_buf, err_buf_len, @errorName(e));
        return -1;
    };
    return 0;
}

// ─── Sheet-level features (Phase 3b stage 5) ─────────────────────────
//
// These operate on a SheetWriter — not the Writer itself — because
// column widths / freeze panes / auto-filter are stored in each sheet's
// XML, not in xl/styles.xml. Zero indicates "no freeze" per axis.

/// Set the width (in character units) of column `col_idx` (0-based,
/// A=0). Returns 0 on success, -1 on invalid width (non-finite or ≤ 0).
export fn zlsx_sheet_writer_set_column_width(
    sw: *SheetWriter,
    col_idx: u32,
    width: f32,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) i32 {
    const sw_state: *SheetWriterState = @ptrCast(@alignCast(sw));
    sw_state.inner.setColumnWidth(col_idx, width) catch |e| {
        writeError(err_buf, err_buf_len, @errorName(e));
        return -1;
    };
    return 0;
}

/// Freeze the top `rows` rows and left `cols` columns. Pass 0 on an
/// axis to leave it unfrozen. Overrides any previous freeze on this
/// sheet. Never fails.
export fn zlsx_sheet_writer_freeze_panes(
    sw: *SheetWriter,
    rows: u32,
    cols: u32,
) callconv(.c) void {
    const sw_state: *SheetWriterState = @ptrCast(@alignCast(sw));
    sw_state.inner.freezePanes(rows, cols);
}

/// Apply an auto-filter over an A1-style range (e.g. "A1:E1"). The
/// writer dupes the range, so the caller can free their buffer
/// immediately after. Returns 0 on success, -1 on an empty range.
export fn zlsx_sheet_writer_set_auto_filter(
    sw: *SheetWriter,
    range_ptr: [*]const u8,
    range_len: usize,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) i32 {
    const sw_state: *SheetWriterState = @ptrCast(@alignCast(sw));
    const range = range_ptr[0..range_len];
    sw_state.inner.setAutoFilter(range) catch |e| {
        writeError(err_buf, err_buf_len, @errorName(e));
        return -1;
    };
    return 0;
}

/// Register a rectangular merged cell range (A1-style, e.g. "A1:B2").
/// The writer validates + dupes the range immediately. Returns 0 on
/// success; -1 with err="InvalidMergeRange" on empty / single-cell /
/// inverted / out-of-Excel-range input, or "OutOfMemory" on alloc
/// failure. Multiple merges per sheet are allowed; callers are
/// responsible for ensuring they don't overlap (Excel rejects
/// overlapping pairs at file-open time).
export fn zlsx_sheet_writer_add_merged_cell(
    sw: *SheetWriter,
    range_ptr: [*]const u8,
    range_len: usize,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) i32 {
    const sw_state: *SheetWriterState = @ptrCast(@alignCast(sw));
    const range = range_ptr[0..range_len];
    sw_state.inner.addMergedCell(range) catch |e| {
        writeError(err_buf, err_buf_len, @errorName(e));
        return -1;
    };
    return 0;
}

/// Attach a list-type data validation (dropdown) to a cell or
/// rectangular range. `range` is A1-style; `values_ptr` / `lens_ptr`
/// describe an array of `values_count` string slices that become the
/// dropdown options. Excel joins them with commas inside a quoted
/// formula1 string — embedded commas or bare double-quotes in values
/// are rejected. Returns 0 on success, -1 with err set to
/// "InvalidHyperlinkRange" on malformed range or
/// "InvalidDataValidation" on empty values / bad value chars.
export fn zlsx_sheet_writer_add_data_validation_list(
    sw: *SheetWriter,
    range_ptr: [*]const u8,
    range_len: usize,
    values_ptr: [*]const [*]const u8,
    lens_ptr: [*]const usize,
    values_count: usize,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) i32 {
    const sw_state: *SheetWriterState = @ptrCast(@alignCast(sw));
    const range = range_ptr[0..range_len];
    // Re-project the parallel pointer + length arrays into a Zig
    // slice-of-slices on a bounded scratch buffer so the Zig API
    // (which expects []const []const u8) can consume them directly.
    // Cap at a generous 256 values — dropdowns beyond that are rare
    // and exceed Excel's own practical limit anyway.
    if (values_count > 256) {
        writeError(err_buf, err_buf_len, @errorName(error.InvalidDataValidation));
        return -1;
    }
    var buf: [256][]const u8 = undefined;
    for (0..values_count) |i| {
        buf[i] = values_ptr[i][0..lens_ptr[i]];
    }
    sw_state.inner.addDataValidationList(range, buf[0..values_count]) catch |e| {
        writeError(err_buf, err_buf_len, @errorName(e));
        return -1;
    };
    return 0;
}

/// Attach a numeric / date / time / text-length data validation to a
/// cell or rectangular range. `range` is A1-style; `kind_code` is one
/// of `ZLSX_DV_KIND_WHOLE / DECIMAL / DATE / TIME / TEXT_LENGTH`;
/// `op_code` is one of `ZLSX_DV_OP_*` (not `NONE` — numeric
/// validations always have an operator). `formula1` and `formula2`
/// are the comparison arguments. `formula2_ptr` may be NULL with
/// `formula2_len = 0` for single-formula operators (pass non-NULL for
/// `between` / `not_between`). Returns 0 on success, -1 with err set
/// to "InvalidHyperlinkRange" on malformed range or
/// "InvalidDataValidation" on empty formula / two-formula mismatch.
export fn zlsx_sheet_writer_add_data_validation_numeric(
    sw: *SheetWriter,
    range_ptr: [*]const u8,
    range_len: usize,
    kind_code: u32,
    op_code: u32,
    formula1_ptr: [*]const u8,
    formula1_len: usize,
    formula2_ptr: ?[*]const u8,
    formula2_len: usize,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) i32 {
    const sw_state: *SheetWriterState = @ptrCast(@alignCast(sw));
    const range = range_ptr[0..range_len];
    const kind = dvKindFromCode(kind_code) orelse {
        writeError(err_buf, err_buf_len, @errorName(error.InvalidDataValidation));
        return -1;
    };
    const op = dvOpFromCode(op_code) orelse {
        writeError(err_buf, err_buf_len, @errorName(error.InvalidDataValidation));
        return -1;
    };
    const f1 = formula1_ptr[0..formula1_len];
    const f2: ?[]const u8 = if (formula2_ptr) |p| p[0..formula2_len] else null;
    sw_state.inner.addDataValidationNumeric(range, kind, op, f1, f2) catch |e| {
        writeError(err_buf, err_buf_len, @errorName(e));
        return -1;
    };
    return 0;
}

/// Attach a custom-formula data validation to a cell or range. Same
/// error semantics as `zlsx_sheet_writer_add_data_validation_numeric`
/// minus the operator / formula2 (custom has neither).
export fn zlsx_sheet_writer_add_data_validation_custom(
    sw: *SheetWriter,
    range_ptr: [*]const u8,
    range_len: usize,
    formula_ptr: [*]const u8,
    formula_len: usize,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) i32 {
    const sw_state: *SheetWriterState = @ptrCast(@alignCast(sw));
    const range = range_ptr[0..range_len];
    const formula = formula_ptr[0..formula_len];
    sw_state.inner.addDataValidationCustom(range, formula) catch |e| {
        writeError(err_buf, err_buf_len, @errorName(e));
        return -1;
    };
    return 0;
}

fn dvKindFromCode(code: u32) ?writer_mod.DataValidationNumericKind {
    return switch (code) {
        ZLSX_DV_KIND_WHOLE => .whole,
        ZLSX_DV_KIND_DECIMAL => .decimal,
        ZLSX_DV_KIND_DATE => .date,
        ZLSX_DV_KIND_TIME => .time,
        ZLSX_DV_KIND_TEXT_LENGTH => .text_length,
        else => null,
    };
}

fn dvOpFromCode(code: u32) ?writer_mod.DataValidationOp {
    return switch (code) {
        ZLSX_DV_OP_BETWEEN => .between,
        ZLSX_DV_OP_NOT_BETWEEN => .not_between,
        ZLSX_DV_OP_EQUAL => .equal,
        ZLSX_DV_OP_NOT_EQUAL => .not_equal,
        ZLSX_DV_OP_LESS_THAN => .less_than,
        ZLSX_DV_OP_LESS_THAN_OR_EQUAL => .less_than_or_equal,
        ZLSX_DV_OP_GREATER_THAN => .greater_than,
        ZLSX_DV_OP_GREATER_THAN_OR_EQUAL => .greater_than_or_equal,
        else => null,
    };
}

/// Attach an external-URL hyperlink to a cell or rectangular range.
/// `range` is A1-style (single cell "A1" or span "B2:C3"); `url` is
/// the external target (http/https/mailto/file/...). The writer
/// validates + dupes both on intake; the URL is xml-escaped on emit
/// so query-string `&` is safe. Returns 0 on success, -1 with
/// err="InvalidHyperlinkRange" on malformed range,
/// "InvalidHyperlinkUrl" on empty URL, or "OutOfMemory" on alloc
/// failure.
export fn zlsx_sheet_writer_add_hyperlink(
    sw: *SheetWriter,
    range_ptr: [*]const u8,
    range_len: usize,
    url_ptr: [*]const u8,
    url_len: usize,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) i32 {
    const sw_state: *SheetWriterState = @ptrCast(@alignCast(sw));
    const range = range_ptr[0..range_len];
    const url = url_ptr[0..url_len];
    sw_state.inner.addHyperlink(range, url) catch |e| {
        writeError(err_buf, err_buf_len, @errorName(e));
        return -1;
    };
    return 0;
}

// ─── Writer tests ────────────────────────────────────────────────────

test "writer: round-trip via reader" {
    const tmp_path = "/tmp/zlsx_c_abi_writer_roundtrip.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    var err_buf: [128]u8 = undefined;

    const w = zlsx_writer_create(&err_buf, err_buf.len);
    try std.testing.expect(w != null);
    defer zlsx_writer_close(w);

    const sheet_name = "Summary";
    const sw = zlsx_writer_add_sheet(w.?, sheet_name.ptr, sheet_name.len, &err_buf, err_buf.len);
    try std.testing.expect(sw != null);

    // Header row: two strings.
    const empty_bytes: [*]const u8 = @ptrCast("");
    const name_str = "Name";
    const age_str = "Age";
    const row1 = [_]CCell{
        .{ .tag = @intFromEnum(CellTag.string), .str_len = name_str.len, .str_ptr = name_str.ptr, .i = 0, .f = 0, .b = 0, ._pad = [_]u8{0} ** 7 },
        .{ .tag = @intFromEnum(CellTag.string), .str_len = age_str.len, .str_ptr = age_str.ptr, .i = 0, .f = 0, .b = 0, ._pad = [_]u8{0} ** 7 },
    };
    try std.testing.expectEqual(@as(i32, 0), zlsx_sheet_writer_write_row(sw.?, &row1, row1.len, &err_buf, err_buf.len));

    // Data row: string + integer.
    const alice_str = "Alice";
    const row2 = [_]CCell{
        .{ .tag = @intFromEnum(CellTag.string), .str_len = alice_str.len, .str_ptr = alice_str.ptr, .i = 0, .f = 0, .b = 0, ._pad = [_]u8{0} ** 7 },
        .{ .tag = @intFromEnum(CellTag.integer), .str_len = 0, .str_ptr = empty_bytes, .i = 30, .f = 0, .b = 0, ._pad = [_]u8{0} ** 7 },
    };
    try std.testing.expectEqual(@as(i32, 0), zlsx_sheet_writer_write_row(sw.?, &row2, row2.len, &err_buf, err_buf.len));

    // Save.
    try std.testing.expectEqual(@as(i32, 0), zlsx_writer_save(w.?, tmp_path, tmp_path.len, &err_buf, err_buf.len));

    // Read it back through the public API.
    var book = try xlsx.Book.open(std.testing.allocator, tmp_path);
    defer book.deinit();
    try std.testing.expectEqualStrings("Summary", book.sheets[0].name);
    var rows = try book.rows(book.sheets[0], std.testing.allocator);
    defer rows.deinit();
    const r1 = (try rows.next()).?;
    try std.testing.expectEqualStrings("Name", r1[0].string);
    try std.testing.expectEqualStrings("Age", r1[1].string);
    const r2 = (try rows.next()).?;
    try std.testing.expectEqualStrings("Alice", r2[0].string);
    try std.testing.expectEqual(@as(i64, 30), r2[1].integer);
}

test "reader C ABI: data_validation getters round-trip" {
    const tmp_path = "/tmp/zlsx_c_abi_reader_dv.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    {
        var w = xlsx.Writer.init(std.testing.allocator);
        defer w.deinit();
        var sheet = try w.addSheet("S");
        try sheet.addDataValidationList("A2:A10", &.{ "Red", "Green", "Blue" });
        try sheet.addDataValidationList("B2", &.{"Single"});
        // XML-escaped chars must survive writer → reader → C ABI.
        try sheet.addDataValidationList("C3", &.{ "R&D", "Q<A" });
        try sheet.writeRow(&.{.{ .string = "hdr" }});
        try w.save(tmp_path);
    }

    var err_buf: [128]u8 = undefined;
    const book = zlsx_book_open(tmp_path, &err_buf, err_buf.len);
    try std.testing.expect(book != null);
    defer zlsx_book_close(book);

    try std.testing.expectEqual(@as(usize, 3), zlsx_data_validation_count(book.?, 0));
    try std.testing.expectEqual(@as(usize, 0), zlsx_data_validation_count(book.?, 99));

    var dv: CDataValidation = undefined;
    try std.testing.expectEqual(@as(i32, 0), zlsx_data_validation_at(book.?, 0, 0, &dv));
    try std.testing.expectEqual(@as(u32, 0), dv.top_left_col);
    try std.testing.expectEqual(@as(u32, 2), dv.top_left_row);
    try std.testing.expectEqual(@as(u32, 0), dv.bottom_right_col);
    try std.testing.expectEqual(@as(u32, 10), dv.bottom_right_row);
    try std.testing.expectEqual(@as(usize, 3), dv.values_count);

    var vptr: [*]const u8 = undefined;
    var vlen: usize = undefined;
    try std.testing.expectEqual(@as(i32, 0), zlsx_data_validation_value_at(book.?, 0, 0, 0, &vptr, &vlen));
    try std.testing.expectEqualStrings("Red", vptr[0..vlen]);
    try std.testing.expectEqual(@as(i32, 0), zlsx_data_validation_value_at(book.?, 0, 0, 1, &vptr, &vlen));
    try std.testing.expectEqualStrings("Green", vptr[0..vlen]);
    try std.testing.expectEqual(@as(i32, 0), zlsx_data_validation_value_at(book.?, 0, 0, 2, &vptr, &vlen));
    try std.testing.expectEqualStrings("Blue", vptr[0..vlen]);
    try std.testing.expectEqual(@as(i32, -1), zlsx_data_validation_value_at(book.?, 0, 0, 3, &vptr, &vlen));

    // Entity-decoded output on the 3rd validation.
    try std.testing.expectEqual(@as(i32, 0), zlsx_data_validation_at(book.?, 0, 2, &dv));
    try std.testing.expectEqual(@as(usize, 2), dv.values_count);
    try std.testing.expectEqual(@as(i32, 0), zlsx_data_validation_value_at(book.?, 0, 2, 0, &vptr, &vlen));
    try std.testing.expectEqualStrings("R&D", vptr[0..vlen]);
    try std.testing.expectEqual(@as(i32, 0), zlsx_data_validation_value_at(book.?, 0, 2, 1, &vptr, &vlen));
    try std.testing.expectEqualStrings("Q<A", vptr[0..vlen]);

    try std.testing.expectEqual(@as(i32, -1), zlsx_data_validation_at(book.?, 0, 3, &dv));
    try std.testing.expectEqual(@as(i32, -1), zlsx_data_validation_at(book.?, 99, 0, &dv));
}

test "writer C ABI: add_data_validation_numeric + custom round-trip via reader" {
    const tmp_path = "/tmp/zlsx_c_abi_writer_dv_ext.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    var err_buf: [128]u8 = undefined;
    const w = zlsx_writer_create(&err_buf, err_buf.len);
    try std.testing.expect(w != null);
    defer zlsx_writer_close(w);

    const name = "Num";
    const sw = zlsx_writer_add_sheet(w.?, name.ptr, name.len, &err_buf, err_buf.len);
    try std.testing.expect(sw != null);

    // whole between 1..100 — two-formula path.
    const r1 = "B2:B10";
    const f1a = "1";
    const f1b = "100";
    try std.testing.expectEqual(@as(i32, 0), zlsx_sheet_writer_add_data_validation_numeric(
        sw.?,
        r1.ptr,
        r1.len,
        ZLSX_DV_KIND_WHOLE,
        ZLSX_DV_OP_BETWEEN,
        f1a.ptr,
        f1a.len,
        f1b.ptr,
        f1b.len,
        &err_buf,
        err_buf.len,
    ));

    // decimal greater_than 0 — single-formula path, NULL formula2.
    const r2 = "C3";
    const f2 = "0";
    try std.testing.expectEqual(@as(i32, 0), zlsx_sheet_writer_add_data_validation_numeric(
        sw.?,
        r2.ptr,
        r2.len,
        ZLSX_DV_KIND_DECIMAL,
        ZLSX_DV_OP_GREATER_THAN,
        f2.ptr,
        f2.len,
        null,
        0,
        &err_buf,
        err_buf.len,
    ));

    // custom — no op, no formula2. XML-special `<` must round-trip.
    const r3 = "D4";
    const cf = "AND(D4>0,D4<LEN(A1))";
    try std.testing.expectEqual(@as(i32, 0), zlsx_sheet_writer_add_data_validation_custom(
        sw.?,
        r3.ptr,
        r3.len,
        cf.ptr,
        cf.len,
        &err_buf,
        err_buf.len,
    ));

    // Rejection paths: bad range, bad kind code, two-formula mismatch.
    try std.testing.expectEqual(@as(i32, -1), zlsx_sheet_writer_add_data_validation_numeric(
        sw.?,
        "",
        0,
        ZLSX_DV_KIND_WHOLE,
        ZLSX_DV_OP_EQUAL,
        f2.ptr,
        f2.len,
        null,
        0,
        &err_buf,
        err_buf.len,
    ));
    try std.testing.expectEqual(@as(i32, -1), zlsx_sheet_writer_add_data_validation_numeric(
        sw.?,
        "A1",
        2,
        0xDEAD,
        ZLSX_DV_OP_EQUAL,
        f2.ptr,
        f2.len,
        null,
        0,
        &err_buf,
        err_buf.len,
    ));
    // equal with two formulas is an InvalidDataValidation.
    try std.testing.expectEqual(@as(i32, -1), zlsx_sheet_writer_add_data_validation_numeric(
        sw.?,
        "A1",
        2,
        ZLSX_DV_KIND_WHOLE,
        ZLSX_DV_OP_EQUAL,
        f1a.ptr,
        f1a.len,
        f1b.ptr,
        f1b.len,
        &err_buf,
        err_buf.len,
    ));
    try std.testing.expectEqual(@as(i32, -1), zlsx_sheet_writer_add_data_validation_custom(
        sw.?,
        "A1",
        2,
        "",
        0,
        &err_buf,
        err_buf.len,
    ));

    // Need at least one row so the writer emits the sheet.
    const hdr = "hdr";
    const row = [_]CCell{
        .{ .tag = @intFromEnum(CellTag.string), .str_len = hdr.len, .str_ptr = hdr.ptr, .i = 0, .f = 0, .b = 0, ._pad = [_]u8{0} ** 7 },
    };
    try std.testing.expectEqual(@as(i32, 0), zlsx_sheet_writer_write_row(sw.?, &row, row.len, &err_buf, err_buf.len));
    try std.testing.expectEqual(@as(i32, 0), zlsx_writer_save(w.?, tmp_path, tmp_path.len, &err_buf, err_buf.len));

    // Read it back and verify every field via the reader C ABI.
    const book = zlsx_book_open(tmp_path.ptr, &err_buf, err_buf.len);
    try std.testing.expect(book != null);
    defer zlsx_book_close(book);

    try std.testing.expectEqual(@as(usize, 3), zlsx_data_validation_count(book.?, 0));

    // dv 0: whole between 1..100
    try std.testing.expectEqual(ZLSX_DV_KIND_WHOLE, zlsx_data_validation_kind(book.?, 0, 0));
    try std.testing.expectEqual(ZLSX_DV_OP_BETWEEN, zlsx_data_validation_operator(book.?, 0, 0));
    var fp: [*]const u8 = undefined;
    var fl: usize = 0;
    try std.testing.expectEqual(@as(i32, 0), zlsx_data_validation_formula1(book.?, 0, 0, &fp, &fl));
    try std.testing.expectEqualStrings("1", fp[0..fl]);
    try std.testing.expectEqual(@as(i32, 0), zlsx_data_validation_formula2(book.?, 0, 0, &fp, &fl));
    try std.testing.expectEqualStrings("100", fp[0..fl]);

    // dv 1: decimal greater_than 0
    try std.testing.expectEqual(ZLSX_DV_KIND_DECIMAL, zlsx_data_validation_kind(book.?, 0, 1));
    try std.testing.expectEqual(ZLSX_DV_OP_GREATER_THAN, zlsx_data_validation_operator(book.?, 0, 1));
    try std.testing.expectEqual(@as(i32, 0), zlsx_data_validation_formula1(book.?, 0, 1, &fp, &fl));
    try std.testing.expectEqualStrings("0", fp[0..fl]);
    try std.testing.expectEqual(@as(i32, 0), zlsx_data_validation_formula2(book.?, 0, 1, &fp, &fl));
    try std.testing.expectEqual(@as(usize, 0), fl);

    // dv 2: custom
    try std.testing.expectEqual(ZLSX_DV_KIND_CUSTOM, zlsx_data_validation_kind(book.?, 0, 2));
    try std.testing.expectEqual(ZLSX_DV_OP_NONE, zlsx_data_validation_operator(book.?, 0, 2));
    try std.testing.expectEqual(@as(i32, 0), zlsx_data_validation_formula1(book.?, 0, 2, &fp, &fl));
    try std.testing.expectEqualStrings("AND(D4>0,D4<LEN(A1))", fp[0..fl]);
}

test "reader C ABI: merged_range + hyperlink getters round-trip" {
    const tmp_path = "/tmp/zlsx_c_abi_reader_meta.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    // Build a workbook with merges + hyperlinks through the Zig writer,
    // then read it back through the C ABI and verify every field.
    {
        var w = xlsx.Writer.init(std.testing.allocator);
        defer w.deinit();
        var sheet = try w.addSheet("S1");
        try sheet.addMergedCell("A1:B2");
        try sheet.addMergedCell("D5:D7");
        try sheet.addHyperlink("C3", "https://example.com/a");
        try sheet.addHyperlink("E5:F5", "mailto:x@example.com");
        try sheet.writeRow(&.{.{ .string = "x" }});
        try w.save(tmp_path);
    }

    var err_buf: [128]u8 = undefined;
    const book = zlsx_book_open(tmp_path, &err_buf, err_buf.len);
    try std.testing.expect(book != null);
    defer zlsx_book_close(book);

    // Merged ranges.
    try std.testing.expectEqual(@as(usize, 2), zlsx_merged_range_count(book.?, 0));
    try std.testing.expectEqual(@as(usize, 0), zlsx_merged_range_count(book.?, 99)); // out of range

    var mr: CMergeRange = undefined;
    try std.testing.expectEqual(@as(i32, 0), zlsx_merged_range_at(book.?, 0, 0, &mr));
    try std.testing.expectEqual(@as(u32, 0), mr.top_left_col);
    try std.testing.expectEqual(@as(u32, 1), mr.top_left_row);
    try std.testing.expectEqual(@as(u32, 1), mr.bottom_right_col);
    try std.testing.expectEqual(@as(u32, 2), mr.bottom_right_row);

    try std.testing.expectEqual(@as(i32, 0), zlsx_merged_range_at(book.?, 0, 1, &mr));
    try std.testing.expectEqual(@as(u32, 3), mr.top_left_col); // D
    try std.testing.expectEqual(@as(u32, 5), mr.top_left_row);
    try std.testing.expectEqual(@as(u32, 3), mr.bottom_right_col);
    try std.testing.expectEqual(@as(u32, 7), mr.bottom_right_row);

    try std.testing.expectEqual(@as(i32, -1), zlsx_merged_range_at(book.?, 0, 2, &mr));

    // Hyperlinks.
    try std.testing.expectEqual(@as(usize, 2), zlsx_hyperlink_count(book.?, 0));

    var hl: CHyperlink = undefined;
    try std.testing.expectEqual(@as(i32, 0), zlsx_hyperlink_at(book.?, 0, 0, &hl));
    try std.testing.expectEqual(@as(u32, 2), hl.top_left_col); // C
    try std.testing.expectEqual(@as(u32, 3), hl.top_left_row);
    try std.testing.expectEqual(@as(u32, 2), hl.bottom_right_col);
    try std.testing.expectEqual(@as(u32, 3), hl.bottom_right_row);
    const url1 = hl.url_ptr[0..hl.url_len];
    try std.testing.expectEqualStrings("https://example.com/a", url1);

    try std.testing.expectEqual(@as(i32, 0), zlsx_hyperlink_at(book.?, 0, 1, &hl));
    const url2 = hl.url_ptr[0..hl.url_len];
    try std.testing.expectEqualStrings("mailto:x@example.com", url2);

    try std.testing.expectEqual(@as(i32, -1), zlsx_hyperlink_at(book.?, 0, 2, &hl));
    try std.testing.expectEqual(@as(i32, -1), zlsx_hyperlink_at(book.?, 99, 0, &hl));
}

test "writer C ABI: add_merged_cell round-trips + rejects bad ranges" {
    const tmp_path = "/tmp/zlsx_c_abi_merged_cell.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    var err_buf: [128]u8 = undefined;

    const w = zlsx_writer_create(&err_buf, err_buf.len);
    try std.testing.expect(w != null);
    defer zlsx_writer_close(w);

    const sheet_name = "S1";
    const sw = zlsx_writer_add_sheet(w.?, sheet_name.ptr, sheet_name.len, &err_buf, err_buf.len);
    try std.testing.expect(sw != null);

    // Valid: returns 0 + empty err_buf.
    const good1 = "A1:B2";
    try std.testing.expectEqual(@as(i32, 0), zlsx_sheet_writer_add_merged_cell(sw.?, good1.ptr, good1.len, &err_buf, err_buf.len));
    const good2 = "C3:E5";
    try std.testing.expectEqual(@as(i32, 0), zlsx_sheet_writer_add_merged_cell(sw.?, good2.ptr, good2.len, &err_buf, err_buf.len));

    // Invalid: each error path returns -1 with "InvalidMergeRange".
    const bad_cases = [_][]const u8{
        "", // empty
        "A1", // no colon
        "A1:A1", // single cell
        "B1:A1", // inverted col
        "a1:b2", // lowercase
        "A0:B2", // row 0
        "XFE1:XFE2", // col > 16384
    };
    for (bad_cases) |bad| {
        @memset(&err_buf, 0);
        const rc = zlsx_sheet_writer_add_merged_cell(sw.?, bad.ptr, bad.len, &err_buf, err_buf.len);
        try std.testing.expectEqual(@as(i32, -1), rc);
        try std.testing.expect(std.mem.indexOf(u8, &err_buf, "InvalidMergeRange") != null);
    }

    // Save + confirm the workbook still opens + walks cleanly — if the
    // earlier error paths had poisoned `merged_cells`, save would emit
    // a malformed <mergeCells> block and the reader would choke.
    const one_str = "x";
    const empty_bytes: [*]const u8 = @ptrCast("");
    const row = [_]CCell{
        .{ .tag = @intFromEnum(CellTag.string), .str_len = one_str.len, .str_ptr = one_str.ptr, .i = 0, .f = 0, .b = 0, ._pad = [_]u8{0} ** 7 },
    };
    _ = empty_bytes;
    try std.testing.expectEqual(@as(i32, 0), zlsx_sheet_writer_write_row(sw.?, &row, row.len, &err_buf, err_buf.len));
    try std.testing.expectEqual(@as(i32, 0), zlsx_writer_save(w.?, tmp_path, tmp_path.len, &err_buf, err_buf.len));

    var book = try xlsx.Book.open(std.testing.allocator, tmp_path);
    defer book.deinit();
    var rows_iter = try book.rows(book.sheets[0], std.testing.allocator);
    defer rows_iter.deinit();
    while (try rows_iter.next()) |_| {}
}

// ─── Fuzz tests ──────────────────────────────────────────────────────

fn fuzzItersCabi() usize {
    const env = std.process.getEnvVarOwned(std.heap.page_allocator, "XLSX_FUZZ_ITERS") catch return 1_000;
    defer std.heap.page_allocator.free(env);
    var digits: [32]u8 = undefined;
    var di: usize = 0;
    for (env) |c| {
        if (c == '_') continue;
        if (di == digits.len) break;
        digits[di] = c;
        di += 1;
    }
    return std.fmt.parseInt(usize, digits[0..di], 10) catch 1_000;
}

fn fuzzSeedCabi() u64 {
    if (std.process.getEnvVarOwned(std.heap.page_allocator, "XLSX_FUZZ_SEED")) |s| {
        defer std.heap.page_allocator.free(s);
        return std.fmt.parseInt(u64, s, 10) catch 0xA1F8ED;
    } else |_| {
        return @bitCast(std.time.milliTimestamp());
    }
}

test "fuzz fromCCell: random tags never panic" {
    const iters = fuzzItersCabi();
    var prng = std.Random.DefaultPrng.init(fuzzSeedCabi());
    const rng = prng.random();

    // Keep a valid-looking str_ptr so the string-tag branch can
    // dereference without segfaulting. Content is zeros.
    var pool: [64]u8 = undefined;
    @memset(&pool, 0);

    for (0..iters) |_| {
        const c: CCell = .{
            .tag = rng.int(u32),
            // Cap str_len to the pool size so the returned string slice
            // doesn't point past our buffer when the tag lands on STRING.
            .str_len = @intCast(rng.intRangeAtMost(usize, 0, pool.len)),
            .str_ptr = @ptrCast(&pool),
            .i = rng.int(i64),
            .f = rng.float(f64),
            .b = rng.int(u8),
            ._pad = [_]u8{0} ** 7,
        };
        const got = fromCCell(c) catch |e| {
            try std.testing.expect(e == error.BadCellTag);
            continue;
        };
        // If no error, the returned Cell's tag must match one of the
        // 5 valid CellTag values — the type system already enforces
        // this, but assert for docs' sake.
        switch (got) {
            .empty, .string, .integer, .number, .boolean => {},
        }
    }
}

test "fuzz toCCell ↔ fromCCell round-trip for valid Cells" {
    const iters = fuzzItersCabi();
    var prng = std.Random.DefaultPrng.init(fuzzSeedCabi());
    const rng = prng.random();

    var strpool: [256]u8 = undefined;
    for (&strpool) |*b| b.* = (rng.int(u8) % 94) + 32;

    for (0..iters) |_| {
        const cell: xlsx.Cell = switch (rng.intRangeAtMost(u8, 0, 4)) {
            0 => .empty,
            1 => blk: {
                const start = rng.intRangeAtMost(usize, 0, strpool.len - 1);
                const len = rng.intRangeAtMost(usize, 0, strpool.len - start);
                break :blk .{ .string = strpool[start..][0..len] };
            },
            2 => .{ .integer = rng.int(i64) },
            3 => .{ .number = rng.float(f64) },
            else => .{ .boolean = rng.boolean() },
        };

        const cc = toCCell(cell);
        const back = try fromCCell(cc);

        switch (cell) {
            .empty => try std.testing.expectEqual(@as(std.meta.Tag(xlsx.Cell), .empty), back),
            .string => |s| try std.testing.expectEqualStrings(s, back.string),
            .integer => |n| try std.testing.expectEqual(n, back.integer),
            .number => |f| {
                // NaN != NaN; treat as equal for round-trip purposes.
                if (std.math.isNan(f)) {
                    try std.testing.expect(std.math.isNan(back.number));
                } else {
                    try std.testing.expectEqual(f, back.number);
                }
            },
            .boolean => |b| try std.testing.expectEqual(b, back.boolean),
        }
    }
}

test "fuzz writer via C ABI: random operations round-trip" {
    const iters = fuzzItersCabi() / 20; // expensive — real zip I/O
    const seed = fuzzSeedCabi();
    var prng = std.Random.DefaultPrng.init(seed);
    const rng = prng.random();
    var tmp_path_buf: [64]u8 = undefined;
    const tmp_path = try std.fmt.bufPrint(&tmp_path_buf, "/tmp/zlsx_fuzz_cabi_{x}.xlsx", .{seed});
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    var err_buf: [128]u8 = undefined;

    for (0..iters) |_| {
        const w = zlsx_writer_create(&err_buf, err_buf.len);
        try std.testing.expect(w != null);
        defer zlsx_writer_close(w);

        // Add 1-3 styles at random bool combos.
        const n_styles = rng.intRangeAtMost(usize, 0, 3);
        var style_ids: [3]u32 = undefined;
        for (0..n_styles) |i| {
            var out_idx: u32 = 0;
            const rc = zlsx_writer_add_style(
                w.?,
                @intFromBool(rng.boolean()),
                @intFromBool(rng.boolean()),
                &out_idx,
                &err_buf,
                err_buf.len,
            );
            try std.testing.expectEqual(@as(i32, 0), rc);
            style_ids[i] = out_idx;
        }

        // Add a sheet with a random uppercase-letter name (1-20
        // chars). Stays clear of Excel's reserved-char set
        // (`/\?*[]:`) so the fuzz hammers the cell / row / save
        // paths instead of the name validator — which has its own
        // dedicated coverage in writer.zig.
        var name_buf: [20]u8 = undefined;
        const name_len = rng.intRangeAtMost(usize, 1, name_buf.len);
        for (0..name_len) |i| name_buf[i] = 'A' + rng.intRangeAtMost(u8, 0, 25);
        const name_ptr: [*]const u8 = @ptrCast(&name_buf);
        const sw = zlsx_writer_add_sheet(w.?, name_ptr, name_len, &err_buf, err_buf.len);
        try std.testing.expect(sw != null);

        // Write 0-5 rows with random cells.
        const n_rows = rng.intRangeAtMost(usize, 0, 5);
        var expected_rows: usize = 0;
        for (0..n_rows) |_| {
            var cells: [6]CCell = undefined;
            var styles: [6]u32 = undefined;
            const n_cells = rng.intRangeAtMost(usize, 0, cells.len);
            var str_store: [6][16]u8 = undefined;
            for (0..n_cells) |ci| {
                styles[ci] = if (n_styles > 0 and rng.boolean())
                    style_ids[rng.intRangeAtMost(usize, 0, n_styles - 1)]
                else
                    0;
                const tag = rng.intRangeAtMost(u8, 0, 4);
                const str_len = rng.intRangeAtMost(usize, 0, str_store[ci].len);
                for (0..str_len) |i| str_store[ci][i] = (rng.int(u8) % 94) + 32;
                cells[ci] = .{
                    .tag = @intCast(tag),
                    .str_len = @intCast(str_len),
                    .str_ptr = @ptrCast(&str_store[ci]),
                    .i = rng.intRangeAtMost(i64, -(1 << 40), 1 << 40),
                    .f = rng.float(f64) * 1000,
                    .b = @intFromBool(rng.boolean()),
                    ._pad = [_]u8{0} ** 7,
                };
            }

            const rc = if (rng.boolean() and n_cells > 0)
                zlsx_sheet_writer_write_row_styled(sw.?, &cells, &styles, n_cells, &err_buf, err_buf.len)
            else
                zlsx_sheet_writer_write_row(sw.?, &cells, n_cells, &err_buf, err_buf.len);
            if (rc == 0) expected_rows += 1;
        }

        // 0-3 merge attempts mixing valid + invalid ranges. Invalid
        // ones must return -1 and NOT poison the writer's merged-cell
        // accumulator (the save step below would choke on malformed XML).
        const merge_candidates = [_][]const u8{
            "A1:B2", "C3:D4", "E1:E5", "AA1:AB2",
            "", // invalid
            "A1", // invalid: no colon
            "B1:A1", // invalid: col inverted
            "a1:b2", // invalid: lowercase
            "XFE1:XFE2", // invalid: col > 16384
        };
        const n_merges = rng.intRangeAtMost(usize, 0, 3);
        for (0..n_merges) |_| {
            const r = merge_candidates[rng.intRangeAtMost(usize, 0, merge_candidates.len - 1)];
            // Don't assert on rc — both 0 and -1 are valid outcomes;
            // the invariant we're fuzzing is "save never emits
            // malformed XML regardless of which attempts succeeded".
            _ = zlsx_sheet_writer_add_merged_cell(sw.?, r.ptr, r.len, &err_buf, err_buf.len);
        }

        const save_rc = zlsx_writer_save(w.?, @ptrCast(tmp_path.ptr), tmp_path.len, &err_buf, err_buf.len);
        try std.testing.expectEqual(@as(i32, 0), save_rc);

        // Re-read to verify the file isn't malformed.
        var book = try xlsx.Book.open(std.testing.allocator, tmp_path);
        defer book.deinit();
        try std.testing.expectEqual(@as(usize, 1), book.sheets.len);
        var rows = try book.rows(book.sheets[0], std.testing.allocator);
        defer rows.deinit();
        var read_rows: usize = 0;
        while (try rows.next()) |_| read_rows += 1;
        try std.testing.expectEqual(expected_rows, read_rows);
    }
}

// ─── Deep C-ABI fuzz ────────────────────────────────────────────────

test "fuzz C ABI: err_buf edge cases never overrun" {
    // Known failure paths (missing file, unknown sheet name) with
    // minimum-length / NULL error buffers. writeError must refuse to
    // write anything when buf is NULL or len == 0, and must always
    // null-terminate when len >= 1.
    const iters = fuzzItersCabi();
    var prng = std.Random.DefaultPrng.init(fuzzSeedCabi());
    const rng = prng.random();

    const bogus_path: [*:0]const u8 = "/nonexistent/__zlsx_fuzz_404__.xlsx";
    for (0..iters) |_| {
        // Buffer length in the tricky range [0, 4].
        const len = rng.intRangeAtMost(usize, 0, 4);
        var buf_storage: [5]u8 = undefined;
        // Poison the trailing byte so we can detect overruns.
        buf_storage[buf_storage.len - 1] = 0xAA;
        const buf_ptr: ?[*]u8 = if (rng.boolean()) null else if (len == 0) null else @ptrCast(&buf_storage);

        const book = zlsx_book_open(bogus_path, buf_ptr, len);
        try std.testing.expect(book == null);
        // No overrun: the poisoned trailing byte is untouched.
        try std.testing.expectEqual(@as(u8, 0xAA), buf_storage[buf_storage.len - 1]);
        if (buf_ptr != null and len >= 1) {
            // Must be null-terminated within [0, len-1].
            var saw_null = false;
            for (buf_storage[0..len]) |c| {
                if (c == 0) {
                    saw_null = true;
                    break;
                }
            }
            try std.testing.expect(saw_null);
        }
    }
}

test "fuzz C ABI: interleaved book + rows handles refcount correctly" {
    // Open N books + rows iterators in random order, close in random
    // order. Memory stays balanced (tested via testing.allocator's
    // implicit leak check at end).
    const corpus = "tests/corpus/frictionless_2sheets.xlsx";
    std.fs.cwd().access(corpus, .{}) catch return;

    const iters = fuzzItersCabi() / 10;
    const seed = fuzzSeedCabi();
    var prng = std.Random.DefaultPrng.init(seed);
    const rng = prng.random();
    const path_z: [*:0]const u8 = @ptrCast(corpus.ptr);
    var err: [128]u8 = undefined;

    for (0..iters) |_| {
        var book_handles: [4]?*Book = [_]?*Book{null} ** 4;
        var rows_handles: [8]?*Rows = [_]?*Rows{null} ** 8;

        // Open 1-4 books (all pointing at the same file — refcount is
        // per-handle, so this gives us independent copies of the state).
        const n_books = rng.intRangeAtMost(usize, 1, 4);
        for (0..n_books) |i| {
            book_handles[i] = zlsx_book_open(path_z, &err, err.len);
            try std.testing.expect(book_handles[i] != null);
        }

        // Open 1-8 row iterators across random books.
        const n_rows = rng.intRangeAtMost(usize, 1, 8);
        for (0..n_rows) |i| {
            const bi = rng.intRangeAtMost(usize, 0, n_books - 1);
            rows_handles[i] = zlsx_rows_open(book_handles[bi].?, 0, &err, err.len);
            try std.testing.expect(rows_handles[i] != null);
        }

        // Close in random order (books + rows mixed).
        var close_order: [12]u8 = undefined;
        const total = n_books + n_rows;
        for (0..total) |i| close_order[i] = @intCast(i);
        rng.shuffle(u8, close_order[0..total]);

        for (close_order[0..total]) |idx| {
            if (idx < n_books) {
                zlsx_book_close(book_handles[idx]);
                book_handles[idx] = null;
            } else {
                const ri = idx - @as(u8, @intCast(n_books));
                zlsx_rows_close(rows_handles[ri]);
                rows_handles[ri] = null;
            }
        }
        // If the refcount underflowed / leaked, testing.allocator's
        // leak detector catches it at the end of the test.
    }
}

test "fuzz C ABI writer: NULL err_buf + zero-cell rows" {
    // NULL err_buf on all failure paths, plus write_row with NULL cells
    // and cells_len=0 (which is a legitimate empty row per the ABI).
    const iters = fuzzItersCabi();
    var prng = std.Random.DefaultPrng.init(fuzzSeedCabi());
    const rng = prng.random();

    const seed = fuzzSeedCabi();
    var tmp_buf: [64]u8 = undefined;
    const tmp_path = try std.fmt.bufPrint(&tmp_buf, "/tmp/zlsx_fuzz_cabi_nullbuf_{x}.xlsx", .{seed});
    defer std.fs.cwd().deleteFile(tmp_path) catch {};

    for (0..iters / 50) |_| {
        const w = zlsx_writer_create(null, 0);
        try std.testing.expect(w != null);
        defer zlsx_writer_close(w);

        const name = "S";
        const sw = zlsx_writer_add_sheet(w.?, name.ptr, name.len, null, 0);
        try std.testing.expect(sw != null);

        // Empty row via cells_ptr=NULL, cells_len=0.
        try std.testing.expectEqual(
            @as(i32, 0),
            zlsx_sheet_writer_write_row(sw.?, null, 0, null, 0),
        );

        // Rows with random counts of random cells, all with NULL err_buf.
        const n_rows = rng.intRangeAtMost(usize, 0, 3);
        for (0..n_rows) |_| {
            var cells: [3]CCell = undefined;
            const nc = rng.intRangeAtMost(usize, 0, cells.len);
            for (0..nc) |ci| {
                cells[ci] = .{
                    .tag = @intFromEnum(CellTag.empty),
                    .str_len = 0,
                    .str_ptr = @ptrCast("".ptr),
                    .i = 0,
                    .f = 0,
                    .b = 0,
                    ._pad = [_]u8{0} ** 7,
                };
            }
            _ = zlsx_sheet_writer_write_row(sw.?, if (nc == 0) null else &cells, nc, null, 0);
        }

        try std.testing.expectEqual(
            @as(i32, 0),
            zlsx_writer_save(w.?, @ptrCast(tmp_path.ptr), tmp_path.len, null, 0),
        );
    }
}

test "fuzz C ABI: random u32 tag in CCell never panics through full row" {
    // Goes beyond the existing fromCCell unit fuzz — runs the bad-tag
    // CCell through an actual zlsx_sheet_writer_write_row call so the
    // integer-precision pre-pass + error return path are also exercised.
    const iters = fuzzItersCabi();
    var prng = std.Random.DefaultPrng.init(fuzzSeedCabi());
    const rng = prng.random();
    var err_buf: [64]u8 = undefined;

    const w = zlsx_writer_create(&err_buf, err_buf.len);
    try std.testing.expect(w != null);
    defer zlsx_writer_close(w);
    const name = "S";
    const sw = zlsx_writer_add_sheet(w.?, name.ptr, name.len, &err_buf, err_buf.len);
    try std.testing.expect(sw != null);

    // Static backing buffer for string-tagged cells so str_ptr is always
    // a valid dereferenceable pointer, even if the tag is bogus.
    var backing: [32]u8 = undefined;
    @memset(&backing, 'x');

    for (0..iters) |_| {
        var cells: [3]CCell = undefined;
        for (&cells) |*c| {
            c.* = .{
                .tag = rng.int(u32),
                .str_len = @intCast(rng.intRangeAtMost(usize, 0, backing.len)),
                .str_ptr = @ptrCast(&backing),
                .i = rng.int(i64),
                .f = rng.float(f64),
                .b = rng.int(u8),
                ._pad = [_]u8{0} ** 7,
            };
        }
        // Must either return 0 (all tags valid) or -1 (at least one
        // BadCellTag / IntegerExceedsExcelPrecision), never panic.
        _ = zlsx_sheet_writer_write_row(sw.?, &cells, cells.len, &err_buf, err_buf.len);
    }
}
