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
//! * Allocator is `smp_allocator` (pure-Zig, no libc). This module is
//!   always multi-threaded (R9-12): `zlsx_cancel_token_trigger` is
//!   documented callable from any thread, and -fsingle-threaded lowers
//!   atomics to plain ops, so build.zig never forwards the option here.
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
const fuzz_config = @import("fuzz_config");
const builtin = @import("builtin");
const build_options = @import("build_options");
const fingerprint_config = @import("fingerprint_config");
const xlsx = @import("zlsx");
const zlsx_pkg = @import("zlsx_pkg");
const zlsx_recalc = @import("zlsx_recalc");
const refs = @import("zlsx_refs");
const writer_mod = xlsx.writer_types;

pub const ZLSX_ABI_VERSION: u32 = 1;
// Null-terminated version string derived from build.zig.zon. Using
// comptimePrint guarantees a sentinel-terminated `[*:0]const u8` so the
// C ABI export has the right type.
pub const ZLSX_VERSION_STRING: [*:0]const u8 = std.fmt.comptimePrint("{s}", .{build_options.version});

// R9-12 (M9a1): refuse to compile single-threaded. build.zig hard-sets
// this module multi-threaded; the assertion catches any other build
// graph before it can ship a token whose trigger is a plain store.
comptime {
    if (builtin.single_threaded) {
        @compileError("the zlsx C ABI must be multi-threaded (R9-12): zlsx_cancel_token_t is a cross-thread atomic and -fsingle-threaded lowers atomics to plain ops");
    }
}

// Allocator used for all handle state. smp_allocator is a singleton —
// no per-handle allocator lifetime to worry about. Always available:
// R9-12 pins this module multi-threaded.
const gpa: std.mem.Allocator = std.heap.smp_allocator;

// ─── Handle types ────────────────────────────────────────────────────

/// Opaque book handle. Field layout is private; C callers only see the
/// pointer. Kept as a struct so Zig's `extern` export works cleanly.
pub const Book = extern struct { _opaque: u8 };

pub const Rows = extern struct { _opaque: u8 };

/// Opaque matrix handle — bulk-materialised view of one sheet, intended
/// for FFI consumers (Python, Node, etc.) that pay per-call dispatch
/// cost on the per-row `zlsx_rows_next` loop. One `zlsx_matrix_open`
/// drains the whole sheet into a packed `CCell[]` plus a row-offsets
/// array; the consumer iterates the buffer in-language with zero
/// further FFI calls until `zlsx_matrix_close`.
pub const Matrix = extern struct { _opaque: u8 };

/// Opaque editor handle. Created by `zlsx_editor_open`, freed by
/// `zlsx_editor_close`. Backed by an `zlsx_pkg.Editor` plus the heap
/// allocation that owns its source buffer + entry table.
pub const Editor = extern struct { _opaque: u8 };

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
    /// Owned Io backing every filesystem call made through this handle.
    /// Torn down with the handle; see zlsx_book_open for why it lives
    /// here rather than in a global.
    threaded: std.Io.Threaded,
    refcount: std.atomic.Value(u32) = .{ .raw = 1 },

    fn unref(self: *BookState) void {
        if (self.refcount.fetchSub(1, .acq_rel) == 1) {
            self.inner.deinit();
            self.threaded.deinit();
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

const MatrixState = struct {
    book: *BookState,
    // Owns string bytes duped from the per-row iterator (whose arena
    // resets on each `next()` call). Cell slices in `flat_cells`
    // pointing at SST / sheet-XML buffers stay live for the Book's
    // lifetime, so we only dupe row-arena slices.
    string_arena: std.heap.ArenaAllocator,
    // Flattened CCell buffer; row r runs cells[offsets[r]..offsets[r+1]].
    // Built once at open, alive until close.
    flat_cells: std.ArrayListUnmanaged(CCell),
    offsets: std.ArrayListUnmanaged(usize),
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
    // [*c] is the C-ABI pointer type that explicitly allows null —
    // matches the `const char *` shape on the C side and lets us
    // accept the common `{ str_ptr = NULL, str_len = 0 }` empty-
    // string pattern without invoking UB in `fromCCell`. ABI-
    // identical to `[*]const u8` in an extern struct (one machine
    // pointer either way).
    str_ptr: [*c]const u8,
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

    // 0.16 needs an `Io` for every filesystem call, and a C caller has
    // no way to hand us a Zig `std.Io`. The handle owns one: allocate
    // the state first so the `Threaded` sits at a stable address (it is
    // not safe to move after init), then open into it.
    const state = gpa.create(BookState) catch {
        writeError(err_buf, err_buf_len, "OutOfMemory");
        return null;
    };
    state.* = .{ .inner = undefined, .threaded = .init(gpa, .{}) };

    state.inner = xlsx.Book.open(gpa, state.threaded.io(), path) catch |e| {
        state.threaded.deinit();
        gpa.destroy(state);
        writeError(err_buf, err_buf_len, @errorName(e));
        return null;
    };
    return @ptrCast(state);
}

/// Open an xlsx workbook from bytes already in memory (B-cabi-1).
/// Same semantics as `zlsx_book_open`, but no filesystem access: the
/// buffer is parsed eagerly and **borrowed only for the duration of
/// this call** — the caller may free `data` immediately after return.
/// Returns NULL on failure with a diagnostic in `err_buf` (if non-null).
export fn zlsx_book_open_buffer(
    data: ?[*]const u8,
    len: usize,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) ?*Book {
    const bytes: []const u8 = if (data) |d| d[0..len] else {
        writeError(err_buf, err_buf_len, "NullBuffer");
        return null;
    };

    const state = gpa.create(BookState) catch {
        writeError(err_buf, err_buf_len, "OutOfMemory");
        return null;
    };
    state.* = .{ .inner = undefined, .threaded = .init(gpa, .{}) };

    state.inner = xlsx.Book.openBuffer(gpa, state.threaded.io(), bytes) catch |e| {
        state.threaded.deinit();
        gpa.destroy(state);
        writeError(err_buf, err_buf_len, @errorName(e));
        return null;
    };
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

/// Copy the `location` (internal target, e.g. `Sheet2!A1`) of
/// hyperlink `link_idx` on sheet `idx` into `out_ptr` / `out_len`.
/// Pointer lifetime matches the Book. Returns 0 on success, -1 on
/// out-of-range indices. External hyperlinks return 0 with
/// `out_len = 0`. Added to surface the internal-link destination
/// that `zlsx_hyperlink_at` discards.
export fn zlsx_hyperlink_location_at(
    book: *Book,
    idx: u32,
    link_idx: usize,
    out_ptr: *[*]const u8,
    out_len: *usize,
) callconv(.c) i32 {
    const state: *BookState = @ptrCast(@alignCast(book));
    if (idx >= state.inner.sheets.len) return -1;
    const links = state.inner.hyperlinks(state.inner.sheets[idx]);
    if (link_idx >= links.len) return -1;
    const loc = links[link_idx].location;
    out_ptr.* = loc.ptr;
    out_len.* = loc.len;
    return 0;
}

/// C-shape for a single cell comment. Author / text slices point
/// into the Book's internal arena; valid until `zlsx_book_close`.
pub const CComment = extern struct {
    cell_col: u32,
    cell_row: u32,
    author_len: usize,
    author_ptr: [*]const u8,
    text_len: usize,
    text_ptr: [*]const u8,
};

/// Number of cell comments on sheet `idx`. Returns 0 if `idx` is out
/// of range or the sheet has none.
export fn zlsx_comment_count(book: *Book, idx: u32) callconv(.c) usize {
    const state: *BookState = @ptrCast(@alignCast(book));
    if (idx >= state.inner.sheets.len) return 0;
    return state.inner.comments(state.inner.sheets[idx]).len;
}

/// Copy comment `comment_idx` on sheet `idx` into `out`. Returns 0
/// on success, -1 if either index is out of range.
export fn zlsx_comment_at(
    book: *Book,
    idx: u32,
    comment_idx: usize,
    out: *CComment,
) callconv(.c) i32 {
    const state: *BookState = @ptrCast(@alignCast(book));
    if (idx >= state.inner.sheets.len) return -1;
    const cs = state.inner.comments(state.inner.sheets[idx]);
    if (comment_idx >= cs.len) return -1;
    const c = cs[comment_idx];
    out.* = .{
        .cell_col = c.top_left.col,
        .cell_row = c.top_left.row,
        .author_len = c.author.len,
        .author_ptr = if (c.author.len == 0) @ptrCast("") else c.author.ptr,
        .text_len = c.text.len,
        .text_ptr = if (c.text.len == 0) @ptrCast("") else c.text.ptr,
    };
    return 0;
}

/// Number of rich-text runs for comment `comment_idx` on sheet `idx`.
/// Returns 0 when the comment is a plain single-run string (null
/// `runs`), so callers can probe with this before calling
/// `zlsx_comment_run_at`. -1 on out-of-range indices is not
/// distinguished from 0 — the caller should have bounds-checked via
/// `zlsx_comment_count` first.
export fn zlsx_comment_run_count(
    book: *Book,
    idx: u32,
    comment_idx: usize,
) callconv(.c) usize {
    const state: *BookState = @ptrCast(@alignCast(book));
    if (idx >= state.inner.sheets.len) return 0;
    const cs = state.inner.comments(state.inner.sheets[idx]);
    if (comment_idx >= cs.len) return 0;
    const runs = cs[comment_idx].runs orelse return 0;
    return runs.len;
}

/// Copy rich-text run `run_idx` of comment `comment_idx` on sheet
/// `idx` into the out pointers. Same tri-state return as
/// `zlsx_rich_run_at` from iter27: 0 → text populated, bold/italic
/// as 0/1; -1 → any index out of range (including comments that
/// have no runs). Mirrors the SST rich-run surface so callers can
/// reuse the same iteration idiom.
export fn zlsx_comment_run_at(
    book: *Book,
    idx: u32,
    comment_idx: usize,
    run_idx: usize,
    out_text_ptr: *[*]const u8,
    out_text_len: *usize,
    out_bold: *u8,
    out_italic: *u8,
) callconv(.c) i32 {
    const state: *BookState = @ptrCast(@alignCast(book));
    if (idx >= state.inner.sheets.len) return -1;
    const cs = state.inner.comments(state.inner.sheets[idx]);
    if (comment_idx >= cs.len) return -1;
    const runs = cs[comment_idx].runs orelse return -1;
    if (run_idx >= runs.len) return -1;
    const r = runs[run_idx];
    out_text_ptr.* = r.text.ptr;
    out_text_len.* = r.text.len;
    out_bold.* = if (r.bold) 1 else 0;
    out_italic.* = if (r.italic) 1 else 0;
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

/// Number of shared-string entries in the workbook. Returns 0 when
/// the workbook has no `xl/sharedStrings.xml` part (small xlsx files
/// with only inline strings). Use with `zlsx_shared_string_at` to
/// enumerate every entry — the pairing lets callers discover which
/// SST indices carry rich-text runs via `zlsx_rich_run_count`
/// without having to track the index themselves.
export fn zlsx_shared_string_count(book: *Book) callconv(.c) usize {
    const state: *BookState = @ptrCast(@alignCast(book));
    return state.inner.sharedStringsCount();
}

/// Copy shared-string entry `sst_idx` into `out_ptr` / `out_len`
/// (slice into the Book's internal buffers; do not free). Returns
/// 0 on success, -1 if `sst_idx` is out of range.
export fn zlsx_shared_string_at(
    book: *Book,
    sst_idx: usize,
    out_ptr: *[*]const u8,
    out_len: *usize,
) callconv(.c) i32 {
    const state: *BookState = @ptrCast(@alignCast(book));
    const s = state.inner.sharedStringAt(sst_idx) catch return -1;
    out_ptr.* = if (s.len == 0) @ptrCast("") else s.ptr;
    out_len.* = s.len;
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

/// `<sheet state="…">` codes for `zlsx_sheet_state`, mirroring
/// `xlsx.SheetState` in declaration order. Stable numeric codes so the
/// C / Python surface can switch on them; the header spells the same
/// three and the smoke test pins them.
pub const ZLSX_SHEET_STATE_VISIBLE: i32 = 0;
pub const ZLSX_SHEET_STATE_HIDDEN: i32 = 1;
pub const ZLSX_SHEET_STATE_VERY_HIDDEN: i32 = 2;

/// Sheet visibility — the `<sheet state="…">` attribute of sheet `idx`
/// as the reader modelled it (`Sheet.state`), or -1 when `idx` is out
/// of range. A missing or unrecognised `state` reads as visible, the
/// schema default (`SheetState.parse`'s rule: visibility never fails
/// an open), so this getter and `zlsx list-sheets` agree by
/// construction — the CLI prints `SheetState.toString()` of the same
/// field. Hidden sheets stay in the inventory (`zlsx_sheet_count`,
/// `zlsx_sheet_name`, the row iterators): a `veryHidden` sheet is
/// unreachable from Excel's UI, and this is how a caller learns it is
/// there.
export fn zlsx_sheet_state(book: *Book, idx: u32) callconv(.c) i32 {
    const state: *BookState = @ptrCast(@alignCast(book));
    if (idx >= state.inner.sheets.len) return -1;
    return switch (state.inner.sheets[idx].state) {
        .visible => ZLSX_SHEET_STATE_VISIBLE,
        .hidden => ZLSX_SHEET_STATE_HIDDEN,
        .very_hidden => ZLSX_SHEET_STATE_VERY_HIDDEN,
    };
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
    rs.* = .{ .book = state, .inner = inner, .c_cells = .empty };
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

    // The C-side view is "the current row" for every per-column
    // getter (`zlsx_rows_style_at`, `zlsx_rows_parse_date` and the
    // S3b slice 11 trio); it is rebuilt on 1 and emptied on 0 / -1,
    // so after the end of the sheet or a parse error no getter serves
    // a row the caller no longer has — the reader clears its own
    // per-row lists at the top of every `next()` and, on a mid-row
    // failure, leaves them partially refilled from the torn row, which
    // is why the view and not those lists is the bound.
    rs.c_cells.clearRetainingCapacity();
    const maybe = rs.inner.next() catch |e| {
        writeError(err_buf, err_buf_len, @errorName(e));
        return -1;
    };
    const cells = maybe orelse return 0;

    rs.c_cells.ensureTotalCapacity(gpa, cells.len) catch {
        writeError(err_buf, err_buf_len, "OutOfMemory");
        return -1;
    };
    for (cells) |c| rs.c_cells.appendAssumeCapacity(toCCell(c));

    out_cells.* = rs.c_cells.items.ptr;
    out_len.* = rs.c_cells.items.len;
    return 1;
}

/// Advance past `n` rows without decoding their cells, and write the
/// number actually skipped into `out_skipped` (fewer than `n` only at
/// end of sheet). Returns 0 on success, -1 on failure with a
/// diagnostic in `err_buf`.
///
/// Semantically identical to calling `zlsx_rows_next` `n` times and
/// discarding the results — same landing row, same row numbering — but
/// it does not build the cell arrays for what it passes. Intended for
/// range-partitioned reads, where every partition must first get past
/// the rows belonging to earlier partitions.
///
/// The cells of the most recently yielded row are invalidated, exactly
/// as `zlsx_rows_next` invalidates them — on -1 as well as on 0; a
/// zero-length skip is the no-op the contract reads as, and leaves the
/// current row current.
export fn zlsx_rows_skip(
    rows: *Rows,
    n: usize,
    out_skipped: ?*usize,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) i32 {
    const rs: *RowsState = @ptrCast(@alignCast(rows));

    // Zero rows: no `zlsx_rows_next` is stood in for, so the current
    // row stays current (py-zlsx returns before the call for the same
    // reason — in-house r2 quick win).
    if (n == 0) {
        if (out_skipped) |o| o.* = 0;
        return 0;
    }
    // The C-side cell view belongs to a row that is now behind us —
    // whether or not the skip lands: on the formula-spreads path the
    // reader decodes through `next()`, which resets its per-row lists
    // before a malformed row fails it, so a view kept across a -1
    // would bound the getters on the old row over the torn one
    // (in-house r1 S3B-REL-101/102).
    rs.c_cells.clearRetainingCapacity();
    const skipped = rs.inner.skipRows(n) catch |e| {
        writeError(err_buf, err_buf_len, @errorName(e));
        return -1;
    };
    if (out_skipped) |o| o.* = skipped;
    return 0;
}

/// Style index for column `col_idx` of the most recently yielded row.
/// Valid between `zlsx_rows_next` calls (the same lifetime contract
/// as the cells). Returns:
///    0 → `*out_style_idx` is set to the cell's `s="…"` attribute
///    1 → the cell had no `s` attribute (General / implicit style)
///   -1 → `col_idx` is out of range for the current row
export fn zlsx_rows_style_at(
    rows: *Rows,
    col_idx: usize,
    out_style_idx: *u32,
) callconv(.c) i32 {
    const rs: *RowsState = @ptrCast(@alignCast(rows));
    // Bound on the C-side view, not the reader's parallel list: the
    // view is what the last `zlsx_rows_next` handed the caller —
    // emptied at the end of the sheet, on a parse error and by a
    // non-empty `zlsx_rows_skip` — whereas the reader's per-row lists keep the
    // last decoded row through a fast-path skip: a stale style for a
    // row the caller never saw (S3b slice 11; the three side-channel
    // getters below share the rule).
    if (col_idx >= rs.c_cells.items.len) return -1;
    const styles = rs.inner.styleIndices();
    const s = (if (col_idx < styles.len) styles[col_idx] else null) orelse return 1;
    out_style_idx.* = s;
    return 0;
}

// ─── S3b slice 11: formula text and error tags on the row iterator ───
//
// The reader keeps three per-row side channels beside the cells
// (`Rows.formulaStrings` / `formulaRefs` / `errorStrings`) so the
// `Cell` union never grew a formula or an error arm: a formula cell's
// slot holds its cached `<v>` value and an error cell's slot holds the
// literal as a plain string. The C ABI mirrors that — `zlsx_cell_t`
// and its tags are untouched (an added tag would have turned every
// existing caller's error literal into "unknown tag") — and hands the
// channels over as per-column getters with `zlsx_rows_style_at`'s
// contract: 0 + out params, 1 "not that kind of cell", -1 out of range
// for the current row. The three are mutually exclusive by the
// reader's construction (`consumeCell`: exactly one formula slot for
// a formula cell, and the error slot is cleared when either is set).

/// Own formula text for column `col_idx` of the most recently yielded
/// row: the `<f>` body, entity-decoded — a stand-alone formula, a
/// shared-formula base or an array-formula base (the CLI's
/// `t:"formula"` + `formula`). Returns 0 and writes `*out_ptr` /
/// `*out_len` (the cells' lifetime: until the next `zlsx_rows_next` /
/// a `zlsx_rows_skip` of n >= 1 or a close); 1 when the cell carries no `<f>` body
/// of its own — a value cell, an error cell, or a shared / array slave
/// (see `zlsx_rows_formula_ref_at`); -1 when `col_idx` is out of range
/// for the current row. The out params are written only on 0.
export fn zlsx_rows_formula_at(
    rows: *Rows,
    col_idx: usize,
    out_ptr: *[*]const u8,
    out_len: *usize,
) callconv(.c) i32 {
    const rs: *RowsState = @ptrCast(@alignCast(rows));
    if (col_idx >= rs.c_cells.items.len) return -1;
    const strings = rs.inner.formulaStrings();
    const text = (if (col_idx < strings.len) strings[col_idx] else null) orelse return 1;
    const empty_bytes: [*]const u8 = @ptrCast("");
    out_ptr.* = if (text.len == 0) empty_bytes else text.ptr;
    out_len.* = text.len;
    return 0;
}

/// Base cell of the shared- or array-formula slave at column `col_idx`
/// of the most recently yielded row: the cell carries no `<f>` body of
/// its own — `<f t="shared" si="N"/>`, or a cell inside an earlier
/// `<f t="array" ref="…">` rectangle — and its formula is the base's
/// text (the CLI's `t:"formula"` + `formula_ref`). Returns 0 and writes
/// `*out_col` (0-based, A = 0) / `*out_row` (1-based); 1 when the cell
/// is not a slave; -1 when `col_idx` is out of range for the current
/// row. A slave whose base the reader never saw (a `si` with no base
/// above it) reads as a value cell — the reader's rule. The out params
/// are written only on 0.
export fn zlsx_rows_formula_ref_at(
    rows: *Rows,
    col_idx: usize,
    out_col: *u32,
    out_row: *u32,
) callconv(.c) i32 {
    const rs: *RowsState = @ptrCast(@alignCast(rows));
    if (col_idx >= rs.c_cells.items.len) return -1;
    const bases = rs.inner.formulaRefs();
    const base = (if (col_idx < bases.len) bases[col_idx] else null) orelse return 1;
    out_col.* = base.col;
    out_row.* = base.row;
    return 0;
}

/// Error literal at column `col_idx` of the most recently yielded row:
/// the `<v>` body of a `t="e"` cell (`#DIV/0!`, `#N/A`, `#REF!`,
/// `#VALUE!`, `#NUM!`, `#NAME?`, `#NULL!`, `#GETTING_DATA`), which the
/// cell array hands over as an ordinary ZLSX_CELL_STRING of the same
/// bytes (the CLI's `t:"error"` + `v`). Returns 0 and writes `*out_ptr`
/// / `*out_len` (the cells' lifetime); 1 when the cell is not an error
/// cell — including a formula cell whose cached value is an error
/// literal, where the formula wins (the CLI's `t:"formula"` with
/// `cached:"#DIV/0!"`); -1 when `col_idx` is out of range for the
/// current row. The out params are written only on 0.
export fn zlsx_rows_error_at(
    rows: *Rows,
    col_idx: usize,
    out_ptr: *[*]const u8,
    out_len: *usize,
) callconv(.c) i32 {
    const rs: *RowsState = @ptrCast(@alignCast(rows));
    if (col_idx >= rs.c_cells.items.len) return -1;
    const errors = rs.inner.errorStrings();
    const literal = (if (col_idx < errors.len) errors[col_idx] else null) orelse return 1;
    const empty_bytes: [*]const u8 = @ptrCast("");
    out_ptr.* = if (literal.len == 0) empty_bytes else literal.ptr;
    out_len.* = literal.len;
    return 0;
}

/// C-shape for `xlsx.DateTime`. All fields fit in their native
/// widths (year 1900..=9999 fits u16; month/day/hour/minute/second
/// fit u8). Fixed layout — no padding adjustments needed.
pub const CDateTime = extern struct {
    year: u16,
    month: u8,
    day: u8,
    hour: u8,
    minute: u8,
    second: u8,
    _pad: u8,
};

/// Inverse of `zlsx_rows_parse_date`: convert a `CDateTime` into
/// the Excel serial number that callers pass as `CellTag.number`
/// when writing a date cell. Returns 0 with `*out_serial` set on
/// success, -1 when the DateTime is outside the round-trippable
/// range (year < 1900, malformed fields, or date ≤ 1900-02-29 —
/// the 1900 leap-bug exclusion).
///
/// Pair with a style registered via `zlsx_writer_add_style_ex`
/// with `number_format = "yyyy-mm-dd"` (or any date pattern) to
/// emit a date cell round-trippable via `zlsx_rows_parse_date`.
export fn zlsx_datetime_to_serial(
    dt: *const CDateTime,
    out_serial: *f64,
) callconv(.c) i32 {
    const z_dt: xlsx.DateTime = .{
        .year = dt.year,
        .month = dt.month,
        .day = dt.day,
        .hour = dt.hour,
        .minute = dt.minute,
        .second = dt.second,
    };
    const serial = xlsx.toExcelSerial(z_dt) orelse return -1;
    out_serial.* = serial;
    return 0;
}

/// Convenience: parse the current-row cell at `col_idx` as a
/// date-styled number, writing the decoded `CDateTime` into `out`.
/// Returns 0 on success (non-null DateTime), 1 when the cell isn't
/// a date (wrong type / non-date numFmt / out-of-range serial), -1
/// when there is no current row or `col_idx` is past it (`out`
/// untouched).
///
/// Callers that want both the raw number AND the DateTime should
/// still use the cells array from `zlsx_rows_next` + this
/// function side-by-side.
export fn zlsx_rows_parse_date(
    rows: *Rows,
    col_idx: usize,
    out: *CDateTime,
) callconv(.c) i32 {
    const rs: *RowsState = @ptrCast(@alignCast(rows));
    // The C-side view is "the current row" here as for every other
    // per-column getter (see `zlsx_rows_style_at`): the reader's
    // `row_cells` keep the last decoded row through a fast-path skip,
    // so bounding on them alone served a skipped-past row's date
    // (in-house r1 S3B-REL-103/104).
    if (col_idx >= rs.c_cells.items.len) return -1;
    // Within the view every column exists in the reader's row too (the
    // two are built in lockstep on a yielded row), so a null here is
    // "not a date", never "out of range".
    const dt = rs.inner.parseDate(col_idx) orelse return 1;
    out.* = .{
        .year = dt.year,
        .month = dt.month,
        .day = dt.day,
        .hour = dt.hour,
        .minute = dt.minute,
        .second = dt.second,
        ._pad = 0,
    };
    return 0;
}

/// Resolve a style index to its number-format code. Returns:
///    0 → `*out_ptr` / `*out_len` point at the format string (lifetime
///        matches the Book; borrows from styles.xml for custom
///        codes, or a constant string for built-ins)
///   -1 → `style_idx` is out of range or the workbook has no
///        resolvable format for that index (malformed / missing
///        styles.xml)
export fn zlsx_number_format(
    book: *Book,
    style_idx: u32,
    out_ptr: *[*]const u8,
    out_len: *usize,
) callconv(.c) i32 {
    const state: *BookState = @ptrCast(@alignCast(book));
    const code = state.inner.numberFormat(style_idx) orelse return -1;
    out_ptr.* = code.ptr;
    out_len.* = code.len;
    return 0;
}

/// Returns 1 if `style_idx` resolves to a date/time pattern, 0
/// otherwise (including out-of-range indices and workbooks without
/// styles.xml).
export fn zlsx_is_date_format(book: *Book, style_idx: u32) callconv(.c) u8 {
    const state: *BookState = @ptrCast(@alignCast(book));
    return if (state.inner.isDateFormat(style_idx)) 1 else 0;
}

/// Per-cell font properties. Layout is fixed; `has_color` and
/// `has_size` disambiguate the optional fields. `name_ptr` / `name_len`
/// borrow from the Book's styles.xml — lifetime matches the Book.
pub const CFont = extern struct {
    bold: u8,
    italic: u8,
    has_color: u8,
    has_size: u8,
    color_argb: u32,
    size: f32,
    name_len: usize,
    name_ptr: [*]const u8,
};

/// Resolve a style index to its font properties. Returns 0 on success,
/// -1 on out-of-range style idx or missing styles.xml. The `has_color`
/// / `has_size` fields signal whether the optionals are populated;
/// `name_len == 0` means the font had no explicit `<name val="…"/>`.
export fn zlsx_cell_font(book: *Book, style_idx: u32, out: *CFont) callconv(.c) i32 {
    const state: *BookState = @ptrCast(@alignCast(book));
    const f = state.inner.cellFont(style_idx) orelse return -1;
    out.* = .{
        .bold = if (f.bold) 1 else 0,
        .italic = if (f.italic) 1 else 0,
        .has_color = if (f.color_argb != null) 1 else 0,
        .has_size = if (f.size != null) 1 else 0,
        .color_argb = f.color_argb orelse 0,
        .size = f.size orelse 0,
        .name_len = f.name.len,
        .name_ptr = if (f.name.len == 0) @ptrCast("") else f.name.ptr,
    };
    return 0;
}

/// Per-cell fill. `pattern_ptr` / `pattern_len` hold the OOXML
/// `patternType` attribute (e.g. "none", "solid", "darkDown"). The
/// `has_fg` / `has_bg` flags signal whether the ARGB fields are
/// populated; theme / indexed colors leave them at 0 and
/// `has_*` = 0.
pub const CFill = extern struct {
    has_fg: u8,
    has_bg: u8,
    _pad: [2]u8,
    fg_color_argb: u32,
    bg_color_argb: u32,
    pattern_len: usize,
    pattern_ptr: [*]const u8,
};

/// Resolve a style index to its fill. Returns 0 on success, -1 on
/// out-of-range indices or missing styles.xml. An all-defaults fill
/// (pattern="none", no colors) is a valid success return — absence
/// of a `<patternFill>` child means "no fill", not "no data".
export fn zlsx_cell_fill(book: *Book, style_idx: u32, out: *CFill) callconv(.c) i32 {
    const state: *BookState = @ptrCast(@alignCast(book));
    const f = state.inner.cellFill(style_idx) orelse return -1;
    out.* = .{
        .has_fg = if (f.fg_color_argb != null) 1 else 0,
        .has_bg = if (f.bg_color_argb != null) 1 else 0,
        ._pad = .{ 0, 0 },
        .fg_color_argb = f.fg_color_argb orelse 0,
        .bg_color_argb = f.bg_color_argb orelse 0,
        .pattern_len = f.pattern.len,
        .pattern_ptr = if (f.pattern.len == 0) @ptrCast("") else f.pattern.ptr,
    };
    return 0;
}

/// One side of a cell border. `has_color` signals whether
/// `color_argb` is populated (theme / indexed colors leave it 0
/// with `has_color = 0`). `style_len == 0` means "no border on this
/// side" (e.g. a `<bottom/>` self-closing element).
pub const CBorderSide = extern struct {
    has_color: u8,
    _pad: [3]u8,
    color_argb: u32,
    style_len: usize,
    style_ptr: [*]const u8,
};

/// Full cell border — five sides. `_pad` on each side keeps the
/// struct 4-byte aligned so the embedded-struct layout matches
/// across C compilers.
pub const CCellBorder = extern struct {
    left: CBorderSide,
    right: CBorderSide,
    top: CBorderSide,
    bottom: CBorderSide,
    diagonal: CBorderSide,
};

fn toCBorderSide(s: xlsx.BorderSide) CBorderSide {
    return .{
        .has_color = if (s.color_argb != null) 1 else 0,
        ._pad = .{ 0, 0, 0 },
        .color_argb = s.color_argb orelse 0,
        .style_len = s.style.len,
        .style_ptr = if (s.style.len == 0) @ptrCast("") else s.style.ptr,
    };
}

/// Resolve a style index to its border. Returns 0 on success, -1 on
/// out-of-range indices or missing styles.xml. Absent sides surface
/// with `style_len = 0` — this is the common case since most cells
/// only border 1-2 sides.
export fn zlsx_cell_border(book: *Book, style_idx: u32, out: *CCellBorder) callconv(.c) i32 {
    const state: *BookState = @ptrCast(@alignCast(book));
    const b = state.inner.cellBorder(style_idx) orelse return -1;
    out.* = .{
        .left = toCBorderSide(b.left),
        .right = toCBorderSide(b.right),
        .top = toCBorderSide(b.top),
        .bottom = toCBorderSide(b.bottom),
        .diagonal = toCBorderSide(b.diagonal),
    };
    return 0;
}

/// Cell alignment record for the FFI surface. `horizontal_len = 0`
/// means the alignment is the OOXML default ("general", which the
/// emitter omits). `wrap_text` is `1` when `wrapText="1"` was set.
pub const CCellAlignment = extern struct {
    horizontal_len: usize,
    horizontal_ptr: [*]const u8,
    wrap_text: u8,
    _pad: [7]u8,
};

/// Resolve a style index to its alignment + wrap_text record.
/// Returns 0 on success, -1 on out-of-range index. Cells without a
/// nested `<alignment>` child surface as `horizontal_len = 0,
/// wrap_text = 0`.
export fn zlsx_cell_alignment(book: *Book, style_idx: u32, out: *CCellAlignment) callconv(.c) i32 {
    const state: *BookState = @ptrCast(@alignCast(book));
    const a = state.inner.cellAlignment(style_idx) orelse return -1;
    out.* = .{
        .horizontal_len = if (a.horizontal) |h| h.len else 0,
        .horizontal_ptr = if (a.horizontal) |h|
            (if (h.len == 0) @ptrCast("") else h.ptr)
        else
            @ptrCast(""),
        .wrap_text = if (a.wrap_text) 1 else 0,
        ._pad = .{ 0, 0, 0, 0, 0, 0, 0 },
    };
    return 0;
}

// ─── Bulk row materialisation (Phase 3d, FFI-friendly) ───────────────
//
// Per-row `zlsx_rows_next` pays one FFI call per row. At MB scale
// (e.g. ECDC: 49k rows) the per-call dispatch overhead dominates
// total Python wall time. `zlsx_matrix_open` drains the whole sheet
// into one packed buffer in a single FFI call — the consumer then
// iterates in its own language with zero further FFI roundtrips.
//
// Lifetime: `out_cells` and `out_offsets` from `zlsx_matrix_data`
// stay valid until `zlsx_matrix_close`. Cell string slices point
// into the matrix arena (Book-lifetime SST slices stay valid for
// the matrix's life either way).

/// Materialise an entire sheet into a heap-resident matrix.
/// Walks the per-row iterator once, building a flat `CCell[]` buffer
/// + row-offsets. String slices that point at row-arena memory
/// (the iterator resets it per `next()` call) are duped into the
/// matrix's own arena; SST + sheet-XML borrows stay live for the
/// Book's lifetime so we don't over-copy them.
///
/// NULL on error with `err_buf` populated.
export fn zlsx_matrix_open(
    book: *Book,
    sheet_idx: u32,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) ?*Matrix {
    const state: *BookState = @ptrCast(@alignCast(book));
    if (sheet_idx >= state.inner.sheets.len) {
        writeError(err_buf, err_buf_len, "SheetIndexOutOfRange");
        return null;
    }
    // matrixOpenInner returns ![*]u8 instead of `?*Matrix` so its
    // own errdefers actually fire. The C ABI wrapper translates the
    // error into a NULL return + writeError. (errdefer never fires
    // on a `?T null` return path, only on `!T` error paths — Codex
    // round-2 correctly flagged that the previous shape leaked the
    // BookState refcount and per-iter resources on partial failure.)
    const ms = matrixOpenInner(state, sheet_idx) catch |e| {
        writeError(err_buf, err_buf_len, @errorName(e));
        return null;
    };
    return @ptrCast(ms);
}

fn matrixOpenInner(state: *BookState, sheet_idx: u32) !*MatrixState {
    _ = state.refcount.fetchAdd(1, .acq_rel);
    errdefer state.unref();

    const sheet = state.inner.sheets[sheet_idx];
    var iter = try state.inner.rows(sheet, gpa);
    errdefer iter.deinit();

    const ms = try gpa.create(MatrixState);
    errdefer gpa.destroy(ms);
    ms.* = .{
        .book = state,
        .string_arena = std.heap.ArenaAllocator.init(gpa),
        .flat_cells = .empty,
        .offsets = .empty,
    };
    errdefer ms.string_arena.deinit();
    errdefer ms.flat_cells.deinit(gpa);
    errdefer ms.offsets.deinit(gpa);

    const string_alloc = ms.string_arena.allocator();
    try ms.offsets.append(gpa, 0);

    while (try iter.next()) |row| {
        try ms.flat_cells.ensureUnusedCapacity(gpa, row.len);
        for (row) |c| {
            // Strings may borrow row-arena memory that resets on the
            // next iter.next(); dupe defensively into matrix-owned
            // storage. Over-copies SST/sheet-XML borrows but the cost
            // is bounded and avoids leaking ownership concerns
            // through the FFI layer.
            const out_c = switch (c) {
                .string => |s| toCCell(.{ .string = try string_alloc.dupe(u8, s) }),
                else => toCCell(c),
            };
            ms.flat_cells.appendAssumeCapacity(out_c);
        }
        try ms.offsets.append(gpa, ms.flat_cells.items.len);
    }

    iter.deinit();
    return ms;
}

/// Close and free a Matrix handle. Drops the reference on the
/// underlying Book; if this was the last handle, the Book is freed
/// too. NULL-safe.
export fn zlsx_matrix_close(matrix: ?*Matrix) callconv(.c) void {
    if (matrix) |m| {
        const ms: *MatrixState = @ptrCast(@alignCast(m));
        ms.flat_cells.deinit(gpa);
        ms.offsets.deinit(gpa);
        ms.string_arena.deinit();
        const book = ms.book;
        gpa.destroy(ms);
        book.unref();
    }
}

/// Read the matrix's flattened layout. After this call:
///   `*out_cells` points to the packed `CCell` buffer
///   `*out_offsets` points to the row-start offsets (length n_rows + 1)
///   `*out_n_rows` is the row count
/// All three buffers stay valid until `zlsx_matrix_close`.
export fn zlsx_matrix_data(
    matrix: *Matrix,
    out_cells: *[*]const CCell,
    out_offsets: *[*]const usize,
    out_n_rows: *usize,
) callconv(.c) void {
    const ms: *MatrixState = @ptrCast(@alignCast(matrix));
    out_cells.* = ms.flat_cells.items.ptr;
    out_offsets.* = ms.offsets.items.ptr;
    // offsets layout is (n_rows + 1) entries — first 0, last total.
    out_n_rows.* = ms.offsets.items.len - 1;
}

// ─── Tests ───────────────────────────────────────────────────────────

/// Per-test temporary file helper. Same shape as the helpers in
/// src/writer.zig / xlsx.zig / cli.zig — replaces hard-coded /tmp
/// paths so the suite is portable to Windows. Caller frees the
/// returned slice; `defer tt.deinit()` cleans up the directory.
const TestTmp = struct {
    dir: std.testing.TmpDir,
    pub fn init() TestTmp {
        return .{ .dir = std.testing.tmpDir(.{}) };
    }
    pub fn deinit(self: *TestTmp) void {
        self.dir.cleanup();
    }
    pub fn path(self: *TestTmp, alloc: std.mem.Allocator, io: std.Io, name: []const u8) ![:0]u8 {
        const d = try self.dir.dir.realPathFileAlloc(io, ".", alloc);
        defer alloc.free(d);
        return std.fs.path.joinZ(alloc, &.{ d, name });
    }
};

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
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    // Skip only when the corpus file is absent (the corpus isn't
    // committed — scripts/fetch_test_corpus.sh materializes it). Any
    // other failure path is a real regression and must fail the test.
    const path_bytes = "tests/corpus/frictionless_2sheets.xlsx";
    std.Io.Dir.cwd().access(io, path_bytes, .{}) catch |err| switch (err) {
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

test "zlsx_book_open_buffer: lifecycle on corpus bytes + buffer freed before use" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    const path = "tests/corpus/frictionless_2sheets.xlsx";
    std.Io.Dir.cwd().access(io, path, .{}) catch |err| switch (err) {
        error.FileNotFound => return, // corpus not fetched — same skip rule as the path test
        else => return err,
    };

    var err_buf: [128]u8 = undefined;

    // Free the source bytes before touching the Book: the ABI contract
    // says the buffer is borrowed only for the duration of the call.
    var book: ?*Book = null;
    {
        const bytes = try std.Io.Dir.cwd().readFileAlloc(io, path, std.testing.allocator, .limited(1 << 24));
        book = zlsx_book_open_buffer(bytes.ptr, bytes.len, &err_buf, err_buf.len);
        std.testing.allocator.free(bytes);
    }
    try std.testing.expect(book != null);
    defer zlsx_book_close(book);

    try std.testing.expect(zlsx_sheet_count(book.?) >= 1);

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

test "zlsx_book_open_buffer: garbage and NULL report errors, never crash" {
    var err_buf: [128]u8 = undefined;

    const garbage = "definitely not a zip";
    try std.testing.expectEqual(@as(?*Book, null), zlsx_book_open_buffer(garbage.ptr, garbage.len, &err_buf, err_buf.len));
    try std.testing.expectEqualStrings("BadZip", std.mem.sliceTo(&err_buf, 0));

    try std.testing.expectEqual(@as(?*Book, null), zlsx_book_open_buffer(null, 0, &err_buf, err_buf.len));
    try std.testing.expectEqualStrings("NullBuffer", std.mem.sliceTo(&err_buf, 0));

    // err_buf is optional — passing NULL must be safe on every path.
    try std.testing.expectEqual(@as(?*Book, null), zlsx_book_open_buffer(garbage.ptr, garbage.len, null, 0));
}

test "refcount: close book before rows is safe" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    const path_bytes = "tests/corpus/frictionless_2sheets.xlsx";
    std.Io.Dir.cwd().access(io, path_bytes, .{}) catch |err| switch (err) {
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
    /// See BookState.threaded — the writer handle owns the Io its
    /// save path writes through.
    threaded: std.Io.Threaded,
    /// SheetWriterStates we hand out via `zlsx_writer_add_sheet`.
    /// They're per-writer wrappers around inner-writer pointers, so
    /// their lifetime ends with the parent writer. Track them here
    /// so `zlsx_writer_close` can free each one — the previous
    /// "leak each, freed at process exit" MVP shape ballooned RSS
    /// for long-lived hosts (Python servers etc.).
    sheet_wrappers: std.ArrayListUnmanaged(*SheetWriterState) = .empty,
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
        @intFromEnum(CellTag.string) => blk: {
            // FFI callers commonly model an empty string as
            // `{ str_ptr = NULL, str_len = 0 }`. Slicing a null
            // many-pointer is UB, so normalise that shape to the
            // empty-string sentinel before slicing. str_ptr=NULL
            // with non-zero len is genuinely malformed — reject.
            const ptr_addr = @intFromPtr(c.str_ptr);
            if (c.str_len == 0 or ptr_addr == 0) {
                if (ptr_addr == 0 and c.str_len != 0) break :blk error.BadCellTag;
                break :blk xlsx.Cell{ .string = "" };
            }
            break :blk xlsx.Cell{ .string = c.str_ptr[0..c.str_len] };
        },
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
    state.* = .{
        .inner = writer_mod.Writer.init(gpa),
        .threaded = .init(gpa, .{}),
        .sheet_wrappers = .empty,
    };
    return @ptrCast(state);
}

/// Release all resources held by the writer. Any SheetWriter handles
/// obtained from `zlsx_writer_add_sheet` become invalid. NULL-safe.
export fn zlsx_writer_close(w: ?*Writer) callconv(.c) void {
    if (w) |p| {
        const state: *WriterState = @ptrCast(@alignCast(p));
        for (state.sheet_wrappers.items) |sw| gpa.destroy(sw);
        state.sheet_wrappers.deinit(gpa);
        state.inner.deinit();
        state.threaded.deinit();
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

    // Reserve wrapper-list capacity AND allocate the wrapper struct
    // BEFORE calling inner.addSheet. The previous order let
    // gpa.create or sheet_wrappers.append fail AFTER the inner
    // writer had already gained a new sheet — leaving the writer
    // with an extra orphan sheet a recovering caller would still
    // see in saved output. Doing both allocations first means the
    // only remaining failure point after addSheet is appendAssumeCapacity,
    // which is infallible by construction.
    state.sheet_wrappers.ensureUnusedCapacity(gpa, 1) catch {
        writeError(err_buf, err_buf_len, "OutOfMemory");
        return null;
    };
    const sw_state = gpa.create(SheetWriterState) catch {
        writeError(err_buf, err_buf_len, "OutOfMemory");
        return null;
    };
    const inner = state.inner.addSheet(name) catch |e| {
        gpa.destroy(sw_state);
        writeError(err_buf, err_buf_len, @errorName(e));
        return null;
    };
    sw_state.* = .{ .inner = inner };
    // Infallible: capacity reserved above.
    state.sheet_wrappers.appendAssumeCapacity(sw_state);
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

/// C-shape for a single rich-text run. `has_color` / `has_size` are
/// 0/1 flags; `font_name_len == 0` means "no rFont". Text slice
/// lifetime is the caller's — the writer copies the formatted form
/// into its own SST arena before this function returns.
pub const CRichRun = extern struct {
    text_ptr: [*]const u8,
    text_len: usize,
    bold: u8,
    italic: u8,
    has_color: u8,
    has_size: u8,
    color_argb: u32,
    size: f32,
    font_name_ptr: [*]const u8,
    font_name_len: usize,
};

/// Append a row that mixes plain cells with rich-text cells. For
/// each column: if `rich_runs_lens[col] > 0`, that column is a
/// rich-text cell with `rich_runs_lens[col]` runs pointed at by
/// `rich_runs_ptrs[col]`; otherwise `cells[col]` is a plain
/// value. Returns 0 on success, -1 on failure.
///
/// This is the C-ABI surface for `SheetWriter.writeRichRow`; the
/// Python binding layers on top. Plain-only rows should stay on
/// the existing `zlsx_sheet_writer_write_row` to avoid the extra
/// parallel-array plumbing.
export fn zlsx_sheet_writer_write_rich_row(
    sw: *SheetWriter,
    cells_ptr: ?[*]const CCell,
    rich_runs_ptrs: ?[*]const [*]const CRichRun,
    rich_runs_lens: ?[*]const usize,
    cells_len: usize,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) i32 {
    const sw_state: *SheetWriterState = @ptrCast(@alignCast(sw));

    // Translate each column into a RichRowCell. Plain cells reuse
    // the existing fromCCell conversion; rich cells build a
    // temporary `[]RichTextRun` slice that stays alive for the
    // duration of this call.
    var scratch_cells: [128]writer_mod.RichRowCell = undefined;
    var heap_cells: ?[]writer_mod.RichRowCell = null;
    defer if (heap_cells) |h| gpa.free(h);

    var cells_slice: []writer_mod.RichRowCell = &.{};
    if (cells_len > 0) {
        if (cells_len <= scratch_cells.len) {
            cells_slice = scratch_cells[0..cells_len];
        } else {
            heap_cells = gpa.alloc(writer_mod.RichRowCell, cells_len) catch {
                writeError(err_buf, err_buf_len, "OutOfMemory");
                return -1;
            };
            cells_slice = heap_cells.?;
        }
    }

    // Per-column runs scratch. Total runs across a row rarely exceeds
    // a handful so a single flat arena holds everything; fall back
    // to a heap alloc for very wide rich rows.
    var runs_scratch: [256]writer_mod.RichTextRun = undefined;
    var heap_runs: ?[]writer_mod.RichTextRun = null;
    defer if (heap_runs) |h| gpa.free(h);

    // First pass: count total runs so we can size the runs buffer.
    var total_runs: usize = 0;
    if (rich_runs_lens) |lens| {
        for (0..cells_len) |i| total_runs += lens[i];
    }
    // The extern signature treats both rich_runs_ptrs and _lens as
    // optional so callers with zero rich cells can pass NULL for
    // both. Guard against the invalid-but-legal-ABI case where the
    // caller supplied non-zero counts but a null pointer table —
    // otherwise the `.?` force-unwrap below would panic the process
    // instead of honouring the -1/err_buf return contract.
    if (total_runs > 0 and rich_runs_ptrs == null) {
        writeError(err_buf, err_buf_len, "InvalidInput");
        return -1;
    }
    var runs_all: []writer_mod.RichTextRun = &.{};
    if (total_runs > 0) {
        if (total_runs <= runs_scratch.len) {
            runs_all = runs_scratch[0..total_runs];
        } else {
            heap_runs = gpa.alloc(writer_mod.RichTextRun, total_runs) catch {
                writeError(err_buf, err_buf_len, "OutOfMemory");
                return -1;
            };
            runs_all = heap_runs.?;
        }
    }

    var runs_cursor: usize = 0;
    for (0..cells_len) |i| {
        const runs_len: usize = if (rich_runs_lens) |lens| lens[i] else 0;
        if (runs_len > 0) {
            const src_runs = rich_runs_ptrs.?[i];
            const dst = runs_all[runs_cursor .. runs_cursor + runs_len];
            for (0..runs_len) |r| {
                const s = src_runs[r];
                dst[r] = .{
                    .text = s.text_ptr[0..s.text_len],
                    .bold = s.bold != 0,
                    .italic = s.italic != 0,
                    .color_argb = if (s.has_color != 0) s.color_argb else null,
                    .size = if (s.has_size != 0) s.size else null,
                    .font_name = if (s.font_name_len > 0)
                        s.font_name_ptr[0..s.font_name_len]
                    else
                        null,
                };
            }
            cells_slice[i] = .{ .rich = dst };
            runs_cursor += runs_len;
        } else if (cells_ptr) |cp| {
            // Plain cell — route through fromCCell so the same
            // null-pointer + empty-string contract that
            // zlsx_sheet_writer_write_row honours applies here too.
            // Without this, a caller using the documented
            // `{ str_ptr = NULL, str_len = 0 }` empty-string shape
            // crashed with a null-pointer slice on rich rows.
            const cell = fromCCell(cp[i]) catch |e| {
                writeError(err_buf, err_buf_len, @errorName(e));
                return -1;
            };
            cells_slice[i] = switch (cell) {
                .empty => .empty,
                .string => |s| .{ .string = s },
                .integer => |n| .{ .integer = n },
                .number => |x| .{ .number = x },
                .boolean => |b| .{ .boolean = b },
            };
        } else {
            cells_slice[i] = .empty;
        }
    }

    sw_state.inner.writeRichRow(cells_slice) catch |e| {
        writeError(err_buf, err_buf_len, @errorName(e));
        return -1;
    };
    return 0;
}

/// Append a row that mixes plain value cells with formula cells.
/// For each column: if `formula_lens[i] > 0`, that column is a
/// formula cell — `formula_ptrs[i][0..formula_lens[i]]` is the
/// formula text and `cells[i]` is the cached `<v>` value Excel
/// shows until recalc. `formula_lens[i] == 0` means a regular
/// value cell (formula_ptrs[i] is ignored). Pass NULL for both
/// formula arrays if no column carries a formula — that's
/// equivalent to `zlsx_sheet_writer_write_row`.
///
/// C-ABI surface for `SheetWriter.writeRowWithFormulas`. Returns
/// 0 on success, -1 on failure (writes the error name into
/// `err_buf`). `FormulaCountMismatch` is the typical error from
/// the underlying writer when array lengths disagree.
export fn zlsx_sheet_writer_write_row_with_formulas(
    sw: *SheetWriter,
    cells_ptr: ?[*]const CCell,
    // Inner element is optional so a per-element NULL from the C side
    // surfaces as InvalidInput rather than slicing from a null pointer.
    formula_ptrs: ?[*]const ?[*]const u8,
    formula_lens: ?[*]const usize,
    cells_len: usize,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) i32 {
    const sw_state: *SheetWriterState = @ptrCast(@alignCast(sw));

    // Cells: same scratch / heap pattern as `zlsx_sheet_writer_write_row`.
    var scratch_cells: [128]xlsx.Cell = undefined;
    var heap_cells: ?[]xlsx.Cell = null;
    defer if (heap_cells) |h| gpa.free(h);
    var cells_slice: []xlsx.Cell = &.{};

    // Formulas: parallel `?[]const u8` slice. We size it to cells_len
    // even when no column carries a formula so writeRowWithFormulas
    // sees a length-matched array (its first check is len equality).
    var scratch_formulas: [128]?[]const u8 = undefined;
    var heap_formulas: ?[]?[]const u8 = null;
    defer if (heap_formulas) |h| gpa.free(h);
    var formulas_slice: []?[]const u8 = &.{};

    if (cells_len > 0) {
        if (cells_ptr == null) {
            writeError(err_buf, err_buf_len, "InvalidInput");
            return -1;
        }
        if (cells_len <= scratch_cells.len) {
            cells_slice = scratch_cells[0..cells_len];
            formulas_slice = scratch_formulas[0..cells_len];
        } else {
            heap_cells = gpa.alloc(xlsx.Cell, cells_len) catch {
                writeError(err_buf, err_buf_len, "OutOfMemory");
                return -1;
            };
            cells_slice = heap_cells.?;
            heap_formulas = gpa.alloc(?[]const u8, cells_len) catch {
                writeError(err_buf, err_buf_len, "OutOfMemory");
                return -1;
            };
            formulas_slice = heap_formulas.?;
        }
        const cp = cells_ptr.?;
        for (0..cells_len) |i| {
            cells_slice[i] = fromCCell(cp[i]) catch |e| {
                writeError(err_buf, err_buf_len, @errorName(e));
                return -1;
            };
            const flen: usize = if (formula_lens) |lens| lens[i] else 0;
            if (flen > 0) {
                if (formula_ptrs == null) {
                    writeError(err_buf, err_buf_len, "InvalidInput");
                    return -1;
                }
                const fp = formula_ptrs.?[i] orelse {
                    writeError(err_buf, err_buf_len, "InvalidInput");
                    return -1;
                };
                formulas_slice[i] = fp[0..flen];
            } else {
                formulas_slice[i] = null;
            }
        }
    }

    sw_state.inner.writeRowWithFormulas(cells_slice, formulas_slice) catch |e| {
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
    // calls std.Io.Dir.cwd().createFile. std.mem.Allocator.dupeZ hands us a
    // sentinel-terminated copy without hand-rolling it.
    const owned_path = gpa.dupeZ(u8, path) catch {
        writeError(err_buf, err_buf_len, "OutOfMemory");
        return -1;
    };
    defer gpa.free(owned_path);

    state.inner.save(state.threaded.io(), owned_path) catch |e| {
        writeError(err_buf, err_buf_len, @errorName(e));
        return -1;
    };
    return 0;
}

/// Serialise the in-memory workbook into a freshly allocated buffer
/// instead of a file — the writer-side mirror of `zlsx_book_open_buffer`.
/// On success writes the base pointer into `out_ptr`, the length into
/// `out_len`, and returns 0; the caller then owns those bytes and MUST
/// release them with `zlsx_buffer_free(ptr, len)`.
///
/// Returns -1 on failure with a diagnostic in `err_buf`, leaving
/// `out_ptr` / `out_len` untouched. The Writer stays usable and
/// unmodified — call it twice and you get two equal buffers.
export fn zlsx_writer_save_to_buffer(
    w: *Writer,
    out_ptr: ?*[*]u8,
    out_len: ?*usize,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) i32 {
    const op = out_ptr orelse {
        writeError(err_buf, err_buf_len, "NullOutPointer");
        return -1;
    };
    const ol = out_len orelse {
        writeError(err_buf, err_buf_len, "NullOutPointer");
        return -1;
    };

    const state: *WriterState = @ptrCast(@alignCast(w));
    const bytes = state.inner.saveToOwnedBuffer(gpa, state.threaded.io()) catch |e| {
        writeError(err_buf, err_buf_len, @errorName(e));
        return -1;
    };

    // A zero-length archive is not reachable (NoSheets fires first, and
    // any real archive carries an end-of-central-directory record), so
    // `bytes.ptr` is always a valid single-item-many pointer here.
    op.* = bytes.ptr;
    ol.* = bytes.len;
    return 0;
}

/// Release a buffer handed out by `zlsx_writer_save_to_buffer`. `len`
/// must be the exact length that call reported — Zig's allocator
/// interface frees by slice, not by base pointer alone. NULL is a no-op.
export fn zlsx_buffer_free(ptr: ?[*]u8, len: usize) callconv(.c) void {
    if (ptr) |p| gpa.free(p[0..len]);
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

/// Set the row height (in points) for `row_idx` (0-based). Excel
/// hard-rejects heights outside (0, 409.5]. Returns 0 on success,
/// -1 on `InvalidRowHeight` / `RowOutOfRange` (error name written
/// into `err_buf`).
export fn zlsx_sheet_writer_set_row_height(
    sw: *SheetWriter,
    row_idx: u32,
    height: f32,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) i32 {
    const sw_state: *SheetWriterState = @ptrCast(@alignCast(sw));
    sw_state.inner.setRowHeight(row_idx, height) catch |e| {
        writeError(err_buf, err_buf_len, @errorName(e));
        return -1;
    };
    return 0;
}

/// Freeze the top `rows` rows and left `cols` columns. Pass 0 on an
/// axis to leave it unfrozen. Overrides any previous freeze on this
/// sheet. Never fails — out-of-range counts are clamped to one less
/// than Excel's hard limits (1_048_575 rows, 16_383 cols) so a
/// visible pane always remains below / right of the freeze, which
/// is what freezePanes() requires. The void signature is preserved
/// for ABI back-compat.
export fn zlsx_sheet_writer_freeze_panes(
    sw: *SheetWriter,
    rows: u32,
    cols: u32,
) callconv(.c) void {
    const sw_state: *SheetWriterState = @ptrCast(@alignCast(sw));
    const clamped_rows = @min(rows, 1_048_575);
    const clamped_cols = @min(cols, 16_383);
    // Pair-assert the clamp invariant. freezePanes only fails when
    // rows >= EXCEL_MAX_ROW (1_048_576) or cols >= EXCEL_MAX_COL
    // (16_384), both ruled out by the @min above. If anyone ever
    // tightens those limits without updating this clamp, this assert
    // turns the resulting wrong-precondition into a Debug-build
    // tripwire instead of an opaque release-build host abort.
    std.debug.assert(clamped_rows < 1_048_576);
    std.debug.assert(clamped_cols < 16_384);
    sw_state.inner.freezePanes(clamped_rows, clamped_cols) catch unreachable;
}

/// Checked variant of freeze_panes: returns -1 with the typed error
/// name on out-of-range inputs (RowOutOfRange / ColumnOutOfRange)
/// instead of clamping silently. Newer FFI consumers should prefer
/// this; the legacy `zlsx_sheet_writer_freeze_panes` stays in place
/// for ABI back-compat.
export fn zlsx_sheet_writer_freeze_panes_checked(
    sw: *SheetWriter,
    rows: u32,
    cols: u32,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) i32 {
    const sw_state: *SheetWriterState = @ptrCast(@alignCast(sw));
    sw_state.inner.freezePanes(rows, cols) catch |e| {
        writeError(err_buf, err_buf_len, @errorName(e));
        return -1;
    };
    return 0;
}

/// Register a workbook-level (or sheet-scoped) defined name.
/// `local_sheet_id_neg` < 0 → workbook scope; ≥ 0 → 0-based sheet
/// index (must resolve at save() time). `hidden_flag != 0` →
/// `hidden="1"` attribute. Returns 0 on success, -1 on
/// `InvalidDefinedName` / `InvalidDefinedNameRefersTo` /
/// `DuplicateDefinedName` (error name written into `err_buf`).
export fn zlsx_writer_add_defined_name(
    w: *Writer,
    name_ptr: [*]const u8,
    name_len: usize,
    refers_to_ptr: [*]const u8,
    refers_to_len: usize,
    local_sheet_id_neg: i32,
    hidden_flag: u8,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) i32 {
    const w_state: *WriterState = @ptrCast(@alignCast(w));
    const name = name_ptr[0..name_len];
    const refers_to = refers_to_ptr[0..refers_to_len];
    const lsi: ?u32 = if (local_sheet_id_neg < 0) null else @intCast(local_sheet_id_neg);
    w_state.inner.addDefinedName(name, refers_to, .{
        .local_sheet_id = lsi,
        .hidden = hidden_flag != 0,
    }) catch |e| {
        writeError(err_buf, err_buf_len, @errorName(e));
        return -1;
    };
    return 0;
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
        const len = lens_ptr[i];
        const ptr = values_ptr[i];
        // Inner NULL guard: a caller can pass a values_ptr array
        // where one entry's pointer is NULL. With non-zero len that
        // would UB-slice; with zero len, slicing the null many-pointer
        // is also UB. Normalise NULL+0 to "" and reject NULL with
        // non-zero len as malformed.
        const ptr_addr = @intFromPtr(ptr);
        if (ptr_addr == 0) {
            if (len != 0) {
                writeError(err_buf, err_buf_len, "NullStringInDataValidation");
                return -1;
            }
            buf[i] = "";
        } else {
            buf[i] = ptr[0..len];
        }
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

/// Attach an internal (same-workbook) hyperlink to a cell or range.
/// `location` is the target reference Excel writes verbatim into
/// `<hyperlink location="…"/>` — typically `Sheet2!A1` or
/// `'Sheet With Spaces'!B2`. Equivalent of the Zig writer's
/// `addInternalHyperlink`; mirrors `add_hyperlink` for external URLs.
/// Returns 0 on success, -1 with `err="InvalidHyperlinkRange"` on
/// malformed range, `"InvalidHyperlinkLocation"` on empty location,
/// or `"OutOfMemory"` on alloc failure.
export fn zlsx_sheet_writer_add_internal_hyperlink(
    sw: *SheetWriter,
    range_ptr: [*]const u8,
    range_len: usize,
    location_ptr: [*]const u8,
    location_len: usize,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) i32 {
    const sw_state: *SheetWriterState = @ptrCast(@alignCast(sw));
    const range = range_ptr[0..range_len];
    const location = location_ptr[0..location_len];
    sw_state.inner.addInternalHyperlink(range, location) catch |e| {
        writeError(err_buf, err_buf_len, @errorName(e));
        return -1;
    };
    return 0;
}

/// Differential format for conditional-formatting rules. Plain
/// extern struct so callers can construct it inline. `has_color` /
/// `has_fill` gate the paired ARGB fields — zero means "not set",
/// matching the `?u32` semantics on the Zig-side Dxf.
/// Writer-side per-border-side payload for a Dxf. Distinct from the
/// reader's `CBorderSide` (which carries `style` as a string slice
/// from the parsed OOXML); here `style` is a byte from the
/// `BorderStyle` enum so callers don't have to string-match.
pub const CDxfBorderSide = extern struct {
    /// BorderStyle enum value. 0 = none (no border), 1 = thin,
    /// 2 = medium, 3 = dashed, … 13 = slant_dash_dot. See
    /// `writer.BorderStyle` for the full set.
    style: u8,
    has_color: u8,
    _pad: [2]u8,
    color_argb: u32,
};

pub const CDxf = extern struct {
    bold: u8,
    italic: u8,
    has_color: u8,
    has_fill: u8,
    color_argb: u32,
    fill_fg_argb: u32,
    has_size: u8,
    _pad: [3]u8,
    size: f32,
    border_left: CDxfBorderSide,
    border_right: CDxfBorderSide,
    border_top: CDxfBorderSide,
    border_bottom: CDxfBorderSide,
};

/// Register a differential format on the workbook-wide `<dxfs>`
/// table. Returns 0 on success with `*out_dxf_id` set; -1 on
/// alloc failure. Content-dedup'd: repeat registrations with the
/// same CDxf return the same id.
export fn zlsx_writer_add_dxf(
    w: *Writer,
    dxf: *const CDxf,
    out_dxf_id: *u32,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) i32 {
    const state: *WriterState = @ptrCast(@alignCast(w));
    const z_dxf: writer_mod.Dxf = .{
        .font_bold = dxf.bold != 0,
        .font_italic = dxf.italic != 0,
        .font_color_argb = if (dxf.has_color != 0) dxf.color_argb else null,
        .font_size = if (dxf.has_size != 0) dxf.size else null,
        .fill_fg_argb = if (dxf.has_fill != 0) dxf.fill_fg_argb else null,
        .border_left = cDxfBorderToZig(dxf.border_left),
        .border_right = cDxfBorderToZig(dxf.border_right),
        .border_top = cDxfBorderToZig(dxf.border_top),
        .border_bottom = cDxfBorderToZig(dxf.border_bottom),
    };
    const id = state.inner.addDxf(z_dxf) catch |e| {
        writeError(err_buf, err_buf_len, @errorName(e));
        return -1;
    };
    out_dxf_id.* = id;
    return 0;
}

fn cDxfBorderToZig(s: CDxfBorderSide) writer_mod.BorderSide {
    // Byte → BorderStyle enum. Out-of-range codes fall back to
    // `.none` (safe default — a misconfigured side renders as
    // inherit-from-cell instead of a random border shape).
    // 0.16: `std.meta.intToEnum` (error union) became `std.enums.fromInt`
    // (optional), so the out-of-range fallback is `orelse`, not `catch`.
    const style: writer_mod.BorderStyle = std.enums.fromInt(writer_mod.BorderStyle, s.style) orelse .none;
    return .{
        .style = style,
        .color_argb = if (s.has_color != 0) s.color_argb else null,
    };
}

fn cfOperatorFromCode(code: u32) ?writer_mod.CfOperator {
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

/// Attach a cellIs-type conditional-format rule. `op_code` reuses
/// the `ZLSX_DV_OP_*` table (shared OOXML tokens). `formula2_ptr`
/// may be NULL with `formula2_len = 0` when the operator doesn't
/// need a second formula. Returns 0 on success, -1 with
/// err="InvalidDataValidation" / "InvalidHyperlinkRange" /
/// "UnknownDxfId" on the respective validation failures.
export fn zlsx_sheet_writer_add_conditional_format_cell_is(
    sw: *SheetWriter,
    range_ptr: [*]const u8,
    range_len: usize,
    op_code: u32,
    formula1_ptr: [*]const u8,
    formula1_len: usize,
    formula2_ptr: ?[*]const u8,
    formula2_len: usize,
    dxf_id: u32,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) i32 {
    const sw_state: *SheetWriterState = @ptrCast(@alignCast(sw));
    const range = range_ptr[0..range_len];
    const op = cfOperatorFromCode(op_code) orelse {
        writeError(err_buf, err_buf_len, @errorName(error.InvalidDataValidation));
        return -1;
    };
    const f1 = formula1_ptr[0..formula1_len];
    const f2: ?[]const u8 = if (formula2_ptr) |p| p[0..formula2_len] else null;
    sw_state.inner.addConditionalFormatCellIs(range, op, f1, f2, dxf_id) catch |e| {
        writeError(err_buf, err_buf_len, @errorName(e));
        return -1;
    };
    return 0;
}

/// Attach an expression-type conditional-format rule. Same error
/// semantics as the cellIs export minus the operator / formula2.
export fn zlsx_sheet_writer_add_conditional_format_expression(
    sw: *SheetWriter,
    range_ptr: [*]const u8,
    range_len: usize,
    formula_ptr: [*]const u8,
    formula_len: usize,
    dxf_id: u32,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) i32 {
    const sw_state: *SheetWriterState = @ptrCast(@alignCast(sw));
    const range = range_ptr[0..range_len];
    const formula = formula_ptr[0..formula_len];
    sw_state.inner.addConditionalFormatExpression(range, formula, dxf_id) catch |e| {
        writeError(err_buf, err_buf_len, @errorName(e));
        return -1;
    };
    return 0;
}

/// Attach a color-scale conditional format. 3-stop when `has_mid != 0`
/// (min → mid → max gradient via percentile 50); 2-stop otherwise
/// (min → max). `low_color_argb` / `mid_color_argb` / `high_color_argb`
/// are ARGB values. No dxf_id needed — colors are embedded per-stop.
/// Returns 0 on success, -1 on bad range.
export fn zlsx_sheet_writer_add_conditional_format_color_scale(
    sw: *SheetWriter,
    range_ptr: [*]const u8,
    range_len: usize,
    low_color_argb: u32,
    has_mid: u8,
    mid_color_argb: u32,
    high_color_argb: u32,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) i32 {
    const sw_state: *SheetWriterState = @ptrCast(@alignCast(sw));
    const range = range_ptr[0..range_len];
    const mid: ?u32 = if (has_mid != 0) mid_color_argb else null;
    sw_state.inner.addConditionalFormatColorScale(range, low_color_argb, mid, high_color_argb) catch |e| {
        writeError(err_buf, err_buf_len, @errorName(e));
        return -1;
    };
    return 0;
}

/// Attach a data-bar conditional format. `color_argb` is the bar fill.
/// Returns 0 on success, -1 on bad range.
export fn zlsx_sheet_writer_add_conditional_format_data_bar(
    sw: *SheetWriter,
    range_ptr: [*]const u8,
    range_len: usize,
    color_argb: u32,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) i32 {
    const sw_state: *SheetWriterState = @ptrCast(@alignCast(sw));
    const range = range_ptr[0..range_len];
    sw_state.inner.addConditionalFormatDataBar(range, color_argb) catch |e| {
        writeError(err_buf, err_buf_len, @errorName(e));
        return -1;
    };
    return 0;
}

/// Attach a cell comment (note). `ref` is a single-cell A1 ref
/// ("B2"); ranges are rejected. `author` and `text` are both
/// plain-text; XML-special chars get escaped on emit. Returns 0
/// on success, -1 with err="InvalidCommentRef" /
/// "InvalidHyperlinkRange" on bad ref, "OutOfMemory" on alloc.
export fn zlsx_sheet_writer_add_comment(
    sw: *SheetWriter,
    ref_ptr: [*]const u8,
    ref_len: usize,
    author_ptr: [*]const u8,
    author_len: usize,
    text_ptr: [*]const u8,
    text_len: usize,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) i32 {
    const sw_state: *SheetWriterState = @ptrCast(@alignCast(sw));
    const ref = ref_ptr[0..ref_len];
    const author = author_ptr[0..author_len];
    const text = text_ptr[0..text_len];
    sw_state.inner.addComment(ref, author, text) catch |e| {
        writeError(err_buf, err_buf_len, @errorName(e));
        return -1;
    };
    return 0;
}

// ─── Writer tests ────────────────────────────────────────────────────

test "writer: round-trip via reader" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, io, "c_abi_writer_roundtrip.xlsx");
    defer std.testing.allocator.free(tmp_path);

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
    try std.testing.expectEqual(@as(i32, 0), zlsx_writer_save(w.?, tmp_path.ptr, tmp_path.len, &err_buf, err_buf.len));

    // Read it back through the public API.
    var book = try xlsx.Book.open(std.testing.allocator, io, tmp_path);
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
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, io, "c_abi_reader_dv.xlsx");
    defer std.testing.allocator.free(tmp_path);

    {
        var w = xlsx.Writer.init(std.testing.allocator);
        defer w.deinit();
        var sheet = try w.addSheet("S");
        try sheet.addDataValidationList("A2:A10", &.{ "Red", "Green", "Blue" });
        try sheet.addDataValidationList("B2", &.{"Single"});
        // XML-escaped chars must survive writer → reader → C ABI.
        try sheet.addDataValidationList("C3", &.{ "R&D", "Q<A" });
        try sheet.writeRow(&.{.{ .string = "hdr" }});
        try w.save(io, tmp_path);
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
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, io, "c_abi_writer_dv_ext.xlsx");
    defer std.testing.allocator.free(tmp_path);

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
    try std.testing.expectEqual(@as(i32, 0), zlsx_writer_save(w.?, tmp_path.ptr, tmp_path.len, &err_buf, err_buf.len));

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

test "reader C ABI: cell_font + cell_fill + cell_border + styleIndices + numFmt getters round-trip" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    // iter28-32 added styles-surface exports — this test hits the C
    // layer directly (separate from Python coverage in test_basic.py).
    const tmp_path = try tt.path(std.testing.allocator, io, "c_abi_cell_styles.xlsx");
    defer std.testing.allocator.free(tmp_path);

    {
        var w = xlsx.Writer.init(std.testing.allocator);
        defer w.deinit();
        const date_style = try w.addStyle(.{ .number_format = "yyyy-mm-dd" });
        const bold_red = try w.addStyle(.{
            .font_bold = true,
            .font_color_argb = 0xFFFF0000,
            .font_size = 14,
            .font_name = "Courier New",
            .fill_pattern = .solid,
            .fill_fg_argb = 0xFF00FF00,
            .border_left = .{ .style = .thin, .color_argb = 0xFF000000 },
            .border_top = .{ .style = .medium, .color_argb = 0xFFFF0000 },
        });
        var sheet = try w.addSheet("S");
        try sheet.writeRowStyled(
            &.{ .{ .number = 44927 }, .{ .string = "bold" }, .{ .integer = 42 } },
            &.{ date_style, bold_red, 0 },
        );
        try w.save(io, tmp_path);
    }

    var err_buf: [128]u8 = undefined;
    const book = zlsx_book_open(tmp_path, &err_buf, err_buf.len);
    try std.testing.expect(book != null);
    defer zlsx_book_close(book);

    const rows = zlsx_rows_open(book.?, 0, &err_buf, err_buf.len);
    try std.testing.expect(rows != null);
    defer zlsx_rows_close(rows);

    var cells_ptr: [*]const CCell = undefined;
    var cells_len: usize = 0;
    try std.testing.expectEqual(@as(i32, 1), zlsx_rows_next(
        rows.?,
        &cells_ptr,
        &cells_len,
        &err_buf,
        err_buf.len,
    ));
    try std.testing.expectEqual(@as(usize, 3), cells_len);

    // styleIndices — each cell returns 0 (present) or 1 (no style).
    var s_date: u32 = undefined;
    try std.testing.expectEqual(@as(i32, 0), zlsx_rows_style_at(rows.?, 0, &s_date));
    var s_bold: u32 = undefined;
    try std.testing.expectEqual(@as(i32, 0), zlsx_rows_style_at(rows.?, 1, &s_bold));
    var s_plain: u32 = undefined;
    const plain_rc = zlsx_rows_style_at(rows.?, 2, &s_plain);
    // Plain cell may legitimately return 1 (no `s` attr); accept both.
    try std.testing.expect(plain_rc == 0 or plain_rc == 1);

    // numberFormat on the date cell resolves to "yyyy-mm-dd".
    var nf_ptr: [*]const u8 = undefined;
    var nf_len: usize = 0;
    try std.testing.expectEqual(@as(i32, 0), zlsx_number_format(book.?, s_date, &nf_ptr, &nf_len));
    try std.testing.expectEqualStrings("yyyy-mm-dd", nf_ptr[0..nf_len]);
    try std.testing.expectEqual(@as(u8, 1), zlsx_is_date_format(book.?, s_date));
    try std.testing.expectEqual(@as(u8, 0), zlsx_is_date_format(book.?, s_bold));

    // cellFont on the bold/red/named cell.
    var font: CFont = undefined;
    try std.testing.expectEqual(@as(i32, 0), zlsx_cell_font(book.?, s_bold, &font));
    try std.testing.expectEqual(@as(u8, 1), font.bold);
    try std.testing.expectEqual(@as(u8, 0), font.italic);
    try std.testing.expectEqual(@as(u8, 1), font.has_color);
    try std.testing.expectEqual(@as(u32, 0xFFFF0000), font.color_argb);
    try std.testing.expectEqual(@as(u8, 1), font.has_size);
    try std.testing.expectEqual(@as(f32, 14.0), font.size);
    try std.testing.expectEqualStrings("Courier New", font.name_ptr[0..font.name_len]);

    // cellFill — solid green.
    var fill: CFill = undefined;
    try std.testing.expectEqual(@as(i32, 0), zlsx_cell_fill(book.?, s_bold, &fill));
    try std.testing.expectEqualStrings("solid", fill.pattern_ptr[0..fill.pattern_len]);
    try std.testing.expectEqual(@as(u8, 1), fill.has_fg);
    try std.testing.expectEqual(@as(u32, 0xFF00FF00), fill.fg_color_argb);

    // cellBorder — left thin black, top medium red, rest empty.
    var border: CCellBorder = undefined;
    try std.testing.expectEqual(@as(i32, 0), zlsx_cell_border(book.?, s_bold, &border));
    try std.testing.expectEqualStrings("thin", border.left.style_ptr[0..border.left.style_len]);
    try std.testing.expectEqual(@as(u8, 1), border.left.has_color);
    try std.testing.expectEqual(@as(u32, 0xFF000000), border.left.color_argb);
    try std.testing.expectEqualStrings("medium", border.top.style_ptr[0..border.top.style_len]);
    try std.testing.expectEqual(@as(u32, 0xFFFF0000), border.top.color_argb);
    try std.testing.expectEqual(@as(usize, 0), border.right.style_len);
    try std.testing.expectEqual(@as(usize, 0), border.bottom.style_len);
    try std.testing.expectEqual(@as(usize, 0), border.diagonal.style_len);

    // Out-of-range style idx → -1 on the pointer-out getters, 0 on
    // the predicate getters.
    try std.testing.expectEqual(@as(i32, -1), zlsx_cell_font(book.?, 99999, &font));
    try std.testing.expectEqual(@as(i32, -1), zlsx_cell_fill(book.?, 99999, &fill));
    try std.testing.expectEqual(@as(i32, -1), zlsx_cell_border(book.?, 99999, &border));
    try std.testing.expectEqual(@as(u8, 0), zlsx_is_date_format(book.?, 99999));
}

test "reader C ABI: merged_range + hyperlink getters round-trip" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, io, "c_abi_reader_meta.xlsx");
    defer std.testing.allocator.free(tmp_path);

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
        try sheet.addInternalHyperlink("G7", "S1!A1");
        try sheet.writeRow(&.{.{ .string = "x" }});
        try w.save(io, tmp_path);
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
    try std.testing.expectEqual(@as(usize, 3), zlsx_hyperlink_count(book.?, 0));

    var hl: CHyperlink = undefined;
    var loc_ptr: [*]const u8 = undefined;
    var loc_len: usize = 0;

    try std.testing.expectEqual(@as(i32, 0), zlsx_hyperlink_at(book.?, 0, 0, &hl));
    try std.testing.expectEqual(@as(u32, 2), hl.top_left_col); // C
    try std.testing.expectEqual(@as(u32, 3), hl.top_left_row);
    try std.testing.expectEqual(@as(u32, 2), hl.bottom_right_col);
    try std.testing.expectEqual(@as(u32, 3), hl.bottom_right_row);
    const url1 = hl.url_ptr[0..hl.url_len];
    try std.testing.expectEqualStrings("https://example.com/a", url1);
    // External hyperlinks have an empty location.
    try std.testing.expectEqual(@as(i32, 0), zlsx_hyperlink_location_at(book.?, 0, 0, &loc_ptr, &loc_len));
    try std.testing.expectEqual(@as(usize, 0), loc_len);

    try std.testing.expectEqual(@as(i32, 0), zlsx_hyperlink_at(book.?, 0, 1, &hl));
    const url2 = hl.url_ptr[0..hl.url_len];
    try std.testing.expectEqualStrings("mailto:x@example.com", url2);

    // Internal hyperlink: empty url, location populated.
    try std.testing.expectEqual(@as(i32, 0), zlsx_hyperlink_at(book.?, 0, 2, &hl));
    try std.testing.expectEqual(@as(usize, 0), hl.url_len);
    try std.testing.expectEqual(@as(i32, 0), zlsx_hyperlink_location_at(book.?, 0, 2, &loc_ptr, &loc_len));
    try std.testing.expectEqualStrings("S1!A1", loc_ptr[0..loc_len]);

    try std.testing.expectEqual(@as(i32, -1), zlsx_hyperlink_at(book.?, 0, 3, &hl));
    try std.testing.expectEqual(@as(i32, -1), zlsx_hyperlink_at(book.?, 99, 0, &hl));
    try std.testing.expectEqual(@as(i32, -1), zlsx_hyperlink_location_at(book.?, 0, 3, &loc_ptr, &loc_len));
    try std.testing.expectEqual(@as(i32, -1), zlsx_hyperlink_location_at(book.?, 99, 0, &loc_ptr, &loc_len));
}

test "writer C ABI: add_merged_cell round-trips + rejects bad ranges" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, io, "c_abi_merged_cell.xlsx");
    defer std.testing.allocator.free(tmp_path);

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
    try std.testing.expectEqual(@as(i32, 0), zlsx_writer_save(w.?, tmp_path.ptr, tmp_path.len, &err_buf, err_buf.len));

    var book = try xlsx.Book.open(std.testing.allocator, io, tmp_path);
    defer book.deinit();
    var rows_iter = try book.rows(book.sheets[0], std.testing.allocator);
    defer rows_iter.deinit();
    while (try rows_iter.next()) |_| {}
}

// ─── Fuzz tests ──────────────────────────────────────────────────────

fn fuzzItersCabi() usize {
    // Override comes from build.zig via -Dfuzz-iters or the
    // XLSX_FUZZ_ITERS environment variable; 0.16 test binaries
    // cannot read the environment themselves.
    return fuzz_config.iters_override orelse 1_000;
}

// ─── Editor (load-modify-save) ───────────────────────────────────────

const EditorState = struct {
    inner: zlsx_pkg.Editor,
    /// See BookState.threaded.
    threaded: std.Io.Threaded,
};

/// Open an existing xlsx for append-only mutation. Returns an Editor
/// handle on success, NULL on failure (`err_buf` populated).
export fn zlsx_editor_open(
    path_ptr: [*:0]const u8,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) ?*Editor {
    const path = std.mem.span(path_ptr);
    // Same ownership shape as zlsx_book_open: the handle owns its Io,
    // allocated first so the Threaded never moves after init.
    const state = gpa.create(EditorState) catch {
        writeError(err_buf, err_buf_len, "OutOfMemory");
        return null;
    };
    state.threaded = .init(gpa, .{});
    const inner = zlsx_pkg.Editor.open(gpa, state.threaded.io(), path) catch |e| {
        state.threaded.deinit();
        gpa.destroy(state);
        writeError(err_buf, err_buf_len, @errorName(e));
        return null;
    };
    state.inner = inner;
    return @ptrCast(state);
}

/// Drop the editor handle. Safe with NULL (no-op).
export fn zlsx_editor_close(ed: ?*Editor) callconv(.c) void {
    if (ed) |e| {
        const state: *EditorState = @ptrCast(@alignCast(e));
        state.inner.deinit();
        // The handle owns its Io runtime (see `zlsx_editor_open`):
        // workers joined and signal handlers restored here, as
        // `BookState.unref` does — an open/close per `zlsx.pivots(path)`
        // would otherwise leak one (Codex #207 r4 REL-403).
        state.threaded.deinit();
        gpa.destroy(state);
    }
}

/// Append a single row to the sheet at `sheet_idx`. Cell types
/// (numeric / integer / boolean / empty / string) follow the Zig
/// `appendRows` contract: lossy integers + out-of-bounds rows
/// surface as `IntegerExceedsExcelPrecision` /
/// `RowIndexOutOfRange` etc. on this call. Returns 0 on success,
/// -1 on failure with `err_buf` populated.
export fn zlsx_editor_append_row(
    ed: *Editor,
    sheet_idx: u32,
    cells_ptr: ?[*]const CCell,
    cells_len: usize,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) i32 {
    const state: *EditorState = @ptrCast(@alignCast(ed));

    var scratch: [128]xlsx.Cell = undefined;
    var heap_owned: ?[]xlsx.Cell = null;
    defer if (heap_owned) |h| gpa.free(h);
    var cells_slice: []xlsx.Cell = &.{};

    if (cells_len > 0) {
        if (cells_ptr == null) {
            writeError(err_buf, err_buf_len, "InvalidInput");
            return -1;
        }
        if (cells_len <= scratch.len) {
            cells_slice = scratch[0..cells_len];
        } else {
            heap_owned = gpa.alloc(xlsx.Cell, cells_len) catch {
                writeError(err_buf, err_buf_len, "OutOfMemory");
                return -1;
            };
            cells_slice = heap_owned.?;
        }
        const src = cells_ptr.?;
        for (0..cells_len) |i| {
            cells_slice[i] = fromCCell(src[i]) catch |e| {
                writeError(err_buf, err_buf_len, @errorName(e));
                return -1;
            };
        }
    }

    // The Editor.appendRows API takes `[]const []const Cell` — wrap
    // the single row in a length-1 outer slice. The editor dupes
    // string contents so the borrowed CCell strings can be freed
    // after this call returns.
    const single_row: [1][]const xlsx.Cell = .{cells_slice};
    state.inner.appendRows(sheet_idx, &single_row) catch |e| {
        writeError(err_buf, err_buf_len, @errorName(e));
        return -1;
    };
    return 0;
}

/// In-place cell mutation (Phase 3d, iter-cm-2). Replaces or
/// inserts a single cell on `sheet_idx`. `row` is 1-based;
/// `col` is 0-based. Cell types follow the same `CCell.tag`
/// encoding as the row API. Returns 0 on success, -1 on failure
/// with `err_buf` populated. Documented errors include
/// SetCellSourceCellHasMetadata (source carries `s=` styles or
/// non-canonical body — preserve-and-merge isn't shipped yet),
/// SheetHasUnsavedAppends, and SheetIndexOutOfRange.
export fn zlsx_editor_set_cell(
    ed: *Editor,
    sheet_idx: u32,
    row: u32,
    col: u32,
    cell_ptr: ?*const CCell,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) i32 {
    if (cell_ptr == null) {
        writeError(err_buf, err_buf_len, "InvalidInput");
        return -1;
    }
    const state: *EditorState = @ptrCast(@alignCast(ed));
    const cell = fromCCell(cell_ptr.?.*) catch |e| {
        writeError(err_buf, err_buf_len, @errorName(e));
        return -1;
    };
    state.inner.setCell(sheet_idx, row, col, cell) catch |e| {
        writeError(err_buf, err_buf_len, @errorName(e));
        return -1;
    };
    return 0;
}

/// Save the workbook (with any pending appends applied) atomically
/// to `out_path` (`out_path_len` bytes; not null-terminated).
/// Returns 0 on success, -1 on failure with `err_buf` populated.
export fn zlsx_editor_save(
    ed: *Editor,
    out_path_ptr: [*]const u8,
    out_path_len: usize,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) i32 {
    const state: *EditorState = @ptrCast(@alignCast(ed));
    const out_path = out_path_ptr[0..out_path_len];
    state.inner.save(state.threaded.io(), out_path) catch |e| {
        writeError(err_buf, err_buf_len, @errorName(e));
        return -1;
    };
    return 0;
}

fn fuzzSeedCabi(io: std.Io) u64 {
    if (fuzz_config.seed_override) |s| return s;
    // std.time lost every function in 0.16; a varying default
    // seed now comes from the monotonic clock via Io.
    const ts = std.Io.Clock.now(.awake, io);
    return @bitCast(@as(i64, @truncate(ts.nanoseconds)));
}

test "writer C ABI: write_row_with_formulas round-trips through reader" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, io, "c_abi_write_formulas.xlsx");
    defer std.testing.allocator.free(tmp_path);

    var err_buf: [128]u8 = undefined;
    const w = zlsx_writer_create(&err_buf, err_buf.len);
    try std.testing.expect(w != null);
    defer zlsx_writer_close(w);

    const sheet_name = "S1";
    const sw = zlsx_writer_add_sheet(w.?, sheet_name.ptr, sheet_name.len, &err_buf, err_buf.len);
    try std.testing.expect(sw != null);

    // Row 1: A1=2, B1=3, C1=A1+B1 (cached as 5).
    const row = [_]CCell{
        .{ .tag = @intFromEnum(CellTag.integer), .str_len = 0, .str_ptr = @ptrCast(""), .i = 2, .f = 0, .b = 0, ._pad = [_]u8{0} ** 7 },
        .{ .tag = @intFromEnum(CellTag.integer), .str_len = 0, .str_ptr = @ptrCast(""), .i = 3, .f = 0, .b = 0, ._pad = [_]u8{0} ** 7 },
        .{ .tag = @intFromEnum(CellTag.integer), .str_len = 0, .str_ptr = @ptrCast(""), .i = 5, .f = 0, .b = 0, ._pad = [_]u8{0} ** 7 },
    };
    const formula_c1 = "A1+B1";
    const formula_ptrs = [_][*]const u8{ @ptrCast(""), @ptrCast(""), formula_c1.ptr };
    const formula_lens = [_]usize{ 0, 0, formula_c1.len };

    try std.testing.expectEqual(
        @as(i32, 0),
        zlsx_sheet_writer_write_row_with_formulas(
            sw.?,
            &row,
            &formula_ptrs,
            &formula_lens,
            row.len,
            &err_buf,
            err_buf.len,
        ),
    );
    try std.testing.expectEqual(@as(i32, 0), zlsx_writer_save(w.?, tmp_path.ptr, tmp_path.len, &err_buf, err_buf.len));

    // Read back through the Zig reader, confirm formula text + cached value.
    var book = try xlsx.Book.open(std.testing.allocator, io, tmp_path);
    defer book.deinit();
    var rows = try book.rows(book.sheets[0], std.testing.allocator);
    defer rows.deinit();
    const cells = (try rows.next()) orelse return error.UnexpectedEndOfRows;
    try std.testing.expectEqual(@as(usize, 3), cells.len);
    try std.testing.expectEqual(@as(i64, 2), cells[0].integer);
    try std.testing.expectEqual(@as(i64, 3), cells[1].integer);
    try std.testing.expectEqual(@as(i64, 5), cells[2].integer); // cached value
    const fstrings = rows.formulaStrings();
    try std.testing.expectEqual(@as(usize, 3), fstrings.len);
    try std.testing.expect(fstrings[0] == null);
    try std.testing.expect(fstrings[1] == null);
    try std.testing.expectEqualStrings("A1+B1", fstrings[2].?);

    // Non-zero formula_lens entry with NULL formula_ptrs (whole table)
    // is caller bug — surface as InvalidInput rather than dereferencing
    // null.
    @memset(&err_buf, 0);
    const lens_with_one = [_]usize{ 0, 0, 5 };
    const rc = zlsx_sheet_writer_write_row_with_formulas(
        sw.?,
        &row,
        null, // formula_ptrs intentionally NULL
        &lens_with_one,
        row.len,
        &err_buf,
        err_buf.len,
    );
    try std.testing.expectEqual(@as(i32, -1), rc);
    try std.testing.expect(std.mem.indexOf(u8, &err_buf, "InvalidInput") != null);

    // Per-element NULL formula_ptrs[i] with formula_lens[i] > 0 is the
    // narrower caller bug — same InvalidInput contract, no slice from
    // a null pointer.
    @memset(&err_buf, 0);
    const elem_null_ptrs = [_]?[*]const u8{ null, null, null };
    const rc2 = zlsx_sheet_writer_write_row_with_formulas(
        sw.?,
        &row,
        &elem_null_ptrs,
        &lens_with_one,
        row.len,
        &err_buf,
        err_buf.len,
    );
    try std.testing.expectEqual(@as(i32, -1), rc2);
    try std.testing.expect(std.mem.indexOf(u8, &err_buf, "InvalidInput") != null);
}

test "writer C ABI: add_hyperlink + add_internal_hyperlink round-trip" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, io, "c_abi_hyperlink_writer.xlsx");
    defer std.testing.allocator.free(tmp_path);

    var err_buf: [128]u8 = undefined;
    const w = zlsx_writer_create(&err_buf, err_buf.len);
    try std.testing.expect(w != null);
    defer zlsx_writer_close(w);

    const sheet_name = "Main";
    const sw = zlsx_writer_add_sheet(w.?, sheet_name.ptr, sheet_name.len, &err_buf, err_buf.len);
    try std.testing.expect(sw != null);

    // External + internal hyperlinks via the C ABI.
    const r1 = "A1";
    const url1 = "https://example.com/a";
    try std.testing.expectEqual(@as(i32, 0), zlsx_sheet_writer_add_hyperlink(
        sw.?,
        r1.ptr,
        r1.len,
        url1.ptr,
        url1.len,
        &err_buf,
        err_buf.len,
    ));
    const r2 = "B2";
    const loc2 = "Main!A1";
    try std.testing.expectEqual(@as(i32, 0), zlsx_sheet_writer_add_internal_hyperlink(
        sw.?,
        r2.ptr,
        r2.len,
        loc2.ptr,
        loc2.len,
        &err_buf,
        err_buf.len,
    ));
    const x = "x";
    const row = [_]CCell{
        .{ .tag = @intFromEnum(CellTag.string), .str_len = x.len, .str_ptr = x.ptr, .i = 0, .f = 0, .b = 0, ._pad = [_]u8{0} ** 7 },
    };
    try std.testing.expectEqual(@as(i32, 0), zlsx_sheet_writer_write_row(sw.?, &row, row.len, &err_buf, err_buf.len));
    try std.testing.expectEqual(@as(i32, 0), zlsx_writer_save(w.?, tmp_path.ptr, tmp_path.len, &err_buf, err_buf.len));

    // Read back through the Zig reader, confirm both hyperlinks survived.
    var book = try xlsx.Book.open(std.testing.allocator, io, tmp_path);
    defer book.deinit();
    const links = book.hyperlinks(book.sheets[0]);
    try std.testing.expectEqual(@as(usize, 2), links.len);

    // External link: url populated, location empty.
    try std.testing.expectEqualStrings("https://example.com/a", links[0].url);
    try std.testing.expectEqual(@as(usize, 0), links[0].location.len);
    // Internal link: location populated, url empty.
    try std.testing.expectEqual(@as(usize, 0), links[1].url.len);
    try std.testing.expectEqualStrings("Main!A1", links[1].location);

    // Empty location is rejected as InvalidHyperlinkLocation.
    @memset(&err_buf, 0);
    const empty_loc: []const u8 = "";
    const rc = zlsx_sheet_writer_add_internal_hyperlink(
        sw.?,
        r2.ptr,
        r2.len,
        empty_loc.ptr,
        empty_loc.len,
        &err_buf,
        err_buf.len,
    );
    try std.testing.expectEqual(@as(i32, -1), rc);
    try std.testing.expect(std.mem.indexOf(u8, &err_buf, "InvalidHyperlinkLocation") != null);
}

test "editor C ABI: open + append_row + save round-trip" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src_path = try tt.path(std.testing.allocator, io, "c_abi_editor_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "c_abi_editor_dst.xlsx");
    defer std.testing.allocator.free(dst_path);

    // Build a source workbook through the writer.
    {
        var w = xlsx.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("D");
        try s.writeRow(&.{ .{ .string = "alpha" }, .{ .integer = 1 } });
        try w.save(io, src_path);
    }

    var err_buf: [128]u8 = undefined;
    const path_z = try std.testing.allocator.dupeZ(u8, src_path);
    defer std.testing.allocator.free(path_z);
    const ed = zlsx_editor_open(path_z.ptr, &err_buf, err_buf.len);
    try std.testing.expect(ed != null);
    defer zlsx_editor_close(ed);

    // Append a row via the C ABI: ["beta", 42, true].
    const beta_str = "beta";
    const empty_bytes: [*]const u8 = @ptrCast("");
    const row = [_]CCell{
        .{ .tag = @intFromEnum(CellTag.string), .str_len = beta_str.len, .str_ptr = beta_str.ptr, .i = 0, .f = 0, .b = 0, ._pad = [_]u8{0} ** 7 },
        .{ .tag = @intFromEnum(CellTag.integer), .str_len = 0, .str_ptr = empty_bytes, .i = 42, .f = 0, .b = 0, ._pad = [_]u8{0} ** 7 },
        .{ .tag = @intFromEnum(CellTag.boolean), .str_len = 0, .str_ptr = empty_bytes, .i = 0, .f = 0, .b = 1, ._pad = [_]u8{0} ** 7 },
    };
    const rc = zlsx_editor_append_row(ed.?, 0, &row, row.len, &err_buf, err_buf.len);
    try std.testing.expectEqual(@as(i32, 0), rc);

    const rc_save = zlsx_editor_save(ed.?, dst_path.ptr, dst_path.len, &err_buf, err_buf.len);
    try std.testing.expectEqual(@as(i32, 0), rc_save);

    // Verify via the reader.
    var book = try xlsx.Book.open(std.testing.allocator, io, dst_path);
    defer book.deinit();
    var rows = try book.rows(book.sheets[0], std.testing.allocator);
    defer rows.deinit();
    const r1 = (try rows.next()) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqualStrings("alpha", r1[0].string);
    try std.testing.expectEqual(@as(i64, 1), r1[1].integer);
    const r2 = (try rows.next()) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqualStrings("beta", r2[0].string);
    try std.testing.expectEqual(@as(i64, 42), r2[1].integer);
    try std.testing.expectEqual(true, r2[2].boolean);
    try std.testing.expectEqual(@as(?[]const xlsx.Cell, null), try rows.next());
}

test "fromCCell: null str_ptr with str_len=0 normalises to empty string" {
    // Common FFI shape for an empty string. Slicing a null many-pointer
    // would be UB; we should return the empty-string sentinel cleanly.
    const c: CCell = .{
        .tag = @intFromEnum(CellTag.string),
        .str_len = 0,
        .str_ptr = @ptrFromInt(0),
        .i = 0,
        .f = 0,
        .b = 0,
        ._pad = [_]u8{0} ** 7,
    };
    const got = try fromCCell(c);
    try std.testing.expectEqualStrings("", got.string);
}

test "fromCCell: null str_ptr with str_len>0 is rejected" {
    const c: CCell = .{
        .tag = @intFromEnum(CellTag.string),
        .str_len = 4,
        .str_ptr = @ptrFromInt(0),
        .i = 0,
        .f = 0,
        .b = 0,
        ._pad = [_]u8{0} ** 7,
    };
    try std.testing.expectError(error.BadCellTag, fromCCell(c));
}

test "fuzz fromCCell: random tags never panic" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    const iters = fuzzItersCabi();
    var prng = std.Random.DefaultPrng.init(fuzzSeedCabi(io));
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
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    const iters = fuzzItersCabi();
    var prng = std.Random.DefaultPrng.init(fuzzSeedCabi(io));
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
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    const iters = fuzzItersCabi() / 20; // expensive — real zip I/O
    const seed = fuzzSeedCabi(io);
    var prng = std.Random.DefaultPrng.init(seed);
    const rng = prng.random();
    var tmp_path_buf: [64]u8 = undefined;
    var tt = TestTmp.init();

    defer tt.deinit();

    const _fuzz_name = std.fmt.bufPrint(&tmp_path_buf, "fuzz_cabi_{x}.xlsx", .{seed}) catch unreachable;

    const tmp_path = try tt.path(std.testing.allocator, io, _fuzz_name);

    defer std.testing.allocator.free(tmp_path);
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
        var book = try xlsx.Book.open(std.testing.allocator, io, tmp_path);
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
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    // Known failure paths (missing file, unknown sheet name) with
    // minimum-length / NULL error buffers. writeError must refuse to
    // write anything when buf is NULL or len == 0, and must always
    // null-terminate when len >= 1.
    const iters = fuzzItersCabi();
    var prng = std.Random.DefaultPrng.init(fuzzSeedCabi(io));
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
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    // Open N books + rows iterators in random order, close in random
    // order. Memory stays balanced (tested via testing.allocator's
    // implicit leak check at end).
    const corpus = "tests/corpus/frictionless_2sheets.xlsx";
    std.Io.Dir.cwd().access(io, corpus, .{}) catch return;

    const iters = fuzzItersCabi() / 10;
    const seed = fuzzSeedCabi(io);
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
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    // NULL err_buf on all failure paths, plus write_row with NULL cells
    // and cells_len=0 (which is a legitimate empty row per the ABI).
    const iters = fuzzItersCabi();
    var prng = std.Random.DefaultPrng.init(fuzzSeedCabi(io));
    const rng = prng.random();

    const seed = fuzzSeedCabi(io);
    var tmp_buf: [64]u8 = undefined;
    var tt = TestTmp.init();

    defer tt.deinit();

    const _fuzz_name = std.fmt.bufPrint(&tmp_buf, "fuzz_cabi_nullbuf_{x}.xlsx", .{seed}) catch unreachable;

    const tmp_path = try tt.path(std.testing.allocator, io, _fuzz_name);

    defer std.testing.allocator.free(tmp_path);
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
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    // Goes beyond the existing fromCCell unit fuzz — runs the bad-tag
    // CCell through an actual zlsx_sheet_writer_write_row call so the
    // integer-precision pre-pass + error return path are also exercised.
    const iters = fuzzItersCabi();
    var prng = std.Random.DefaultPrng.init(fuzzSeedCabi(io));
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

test "zlsx_writer_add_defined_name: surfaces typed errors over FFI" {
    var open_err: [128]u8 = undefined;
    @memset(&open_err, 0);
    const w = zlsx_writer_create(&open_err, open_err.len);
    defer zlsx_writer_close(w);
    try std.testing.expect(w != null);
    var err_buf: [128]u8 = undefined;
    @memset(&err_buf, 0);

    // Need at least one sheet so save() is plausible (and so a
    // sheet-scoped name with localSheetId=0 resolves later).
    var sw_err: [128]u8 = undefined;
    @memset(&sw_err, 0);
    const sw = zlsx_writer_add_sheet(w.?, "Sheet1".ptr, "Sheet1".len, &sw_err, sw_err.len);
    try std.testing.expect(sw != null);
    // Sheet writers are owned by the parent Writer; freed on
    // zlsx_writer_close. No explicit close needed.

    // Valid: workbook-scope.
    @memset(&err_buf, 0);
    try std.testing.expectEqual(@as(i32, 0), zlsx_writer_add_defined_name(
        w.?,
        "MyRange".ptr,
        "MyRange".len,
        "Sheet1!$A$1:$B$1".ptr,
        "Sheet1!$A$1:$B$1".len,
        -1, // workbook scope
        0,
        &err_buf,
        err_buf.len,
    ));

    // Valid: sheet-scoped + hidden.
    @memset(&err_buf, 0);
    try std.testing.expectEqual(@as(i32, 0), zlsx_writer_add_defined_name(
        w.?,
        "_xlnm.Print_Area".ptr,
        "_xlnm.Print_Area".len,
        "Sheet1!$A$1:$B$1".ptr,
        "Sheet1!$A$1:$B$1".len,
        0, // local_sheet_id=0
        1, // hidden
        &err_buf,
        err_buf.len,
    ));

    // Invalid: A1-shaped name.
    @memset(&err_buf, 0);
    try std.testing.expectEqual(@as(i32, -1), zlsx_writer_add_defined_name(
        w.?,
        "A1".ptr,
        "A1".len,
        "Sheet1!$A$1".ptr,
        "Sheet1!$A$1".len,
        -1,
        0,
        &err_buf,
        err_buf.len,
    ));
    try std.testing.expect(std.mem.indexOf(u8, &err_buf, "InvalidDefinedName") != null);

    // Invalid: empty refers_to.
    @memset(&err_buf, 0);
    try std.testing.expectEqual(@as(i32, -1), zlsx_writer_add_defined_name(
        w.?,
        "Foo".ptr,
        "Foo".len,
        "".ptr,
        0,
        -1,
        0,
        &err_buf,
        err_buf.len,
    ));
    try std.testing.expect(std.mem.indexOf(u8, &err_buf, "InvalidDefinedNameRefersTo") != null);

    // Duplicate (case-insensitive, same scope).
    @memset(&err_buf, 0);
    try std.testing.expectEqual(@as(i32, -1), zlsx_writer_add_defined_name(
        w.?,
        "myrange".ptr,
        "myrange".len,
        "Sheet1!$A$2".ptr,
        "Sheet1!$A$2".len,
        -1,
        0,
        &err_buf,
        err_buf.len,
    ));
    try std.testing.expect(std.mem.indexOf(u8, &err_buf, "DuplicateDefinedName") != null);
}

test "zlsx_sheet_writer_set_row_height + freeze_panes_checked propagate errors" {
    var open_err: [128]u8 = undefined;
    @memset(&open_err, 0);
    const w = zlsx_writer_create(&open_err, open_err.len);
    defer zlsx_writer_close(w);
    try std.testing.expect(w != null);
    var sw_err: [128]u8 = undefined;
    @memset(&sw_err, 0);
    const sw = zlsx_writer_add_sheet(w.?, "S".ptr, "S".len, &sw_err, sw_err.len) orelse return error.AddSheetFailed;

    var err_buf: [128]u8 = undefined;

    // setRowHeight: valid + invalid.
    @memset(&err_buf, 0);
    try std.testing.expectEqual(@as(i32, 0), zlsx_sheet_writer_set_row_height(sw, 0, 24.0, &err_buf, err_buf.len));
    @memset(&err_buf, 0);
    try std.testing.expectEqual(@as(i32, -1), zlsx_sheet_writer_set_row_height(sw, 0, 0.0, &err_buf, err_buf.len));
    try std.testing.expect(std.mem.indexOf(u8, &err_buf, "InvalidRowHeight") != null);
    // Above Excel's 409.5 cap.
    @memset(&err_buf, 0);
    try std.testing.expectEqual(@as(i32, -1), zlsx_sheet_writer_set_row_height(sw, 0, 410.0, &err_buf, err_buf.len));
    try std.testing.expect(std.mem.indexOf(u8, &err_buf, "InvalidRowHeight") != null);
    // At the cap (boundary).
    @memset(&err_buf, 0);
    try std.testing.expectEqual(@as(i32, 0), zlsx_sheet_writer_set_row_height(sw, 1, 409.5, &err_buf, err_buf.len));

    // freeze_panes_checked: valid.
    @memset(&err_buf, 0);
    try std.testing.expectEqual(@as(i32, 0), zlsx_sheet_writer_freeze_panes_checked(sw, 1, 1, &err_buf, err_buf.len));
    // freeze_panes_checked: out-of-range surfaces typed error.
    @memset(&err_buf, 0);
    try std.testing.expectEqual(@as(i32, -1), zlsx_sheet_writer_freeze_panes_checked(sw, 1_048_576, 0, &err_buf, err_buf.len));
    try std.testing.expect(std.mem.indexOf(u8, &err_buf, "RowOutOfRange") != null);
    @memset(&err_buf, 0);
    try std.testing.expectEqual(@as(i32, -1), zlsx_sheet_writer_freeze_panes_checked(sw, 0, 16_384, &err_buf, err_buf.len));
    try std.testing.expect(std.mem.indexOf(u8, &err_buf, "ColumnOutOfRange") != null);
}

// ─── Document properties (Z3) ────────────────────────────────────────

/// Which `docProps` field `zlsx_editor_docprop_at` returns.
///
/// A field selector rather than 14 separate exports, and an explicit
/// `u32` rather than a Zig enum so the numeric values are part of the
/// ABI contract: appending is safe, renumbering is not.
pub const ZLSX_DOCPROP_CREATOR: u32 = 0;
pub const ZLSX_DOCPROP_LAST_MODIFIED_BY: u32 = 1;
pub const ZLSX_DOCPROP_TITLE: u32 = 2;
pub const ZLSX_DOCPROP_SUBJECT: u32 = 3;
pub const ZLSX_DOCPROP_DESCRIPTION: u32 = 4;
pub const ZLSX_DOCPROP_KEYWORDS: u32 = 5;
pub const ZLSX_DOCPROP_CATEGORY: u32 = 6;
pub const ZLSX_DOCPROP_CREATED: u32 = 7;
pub const ZLSX_DOCPROP_MODIFIED: u32 = 8;
pub const ZLSX_DOCPROP_REVISION: u32 = 9;
pub const ZLSX_DOCPROP_COMPANY: u32 = 10;
pub const ZLSX_DOCPROP_MANAGER: u32 = 11;
pub const ZLSX_DOCPROP_APPLICATION: u32 = 12;
pub const ZLSX_DOCPROP_HYPERLINK_BASE: u32 = 13;

/// Read one document-properties field from an Editor handle into
/// `out_ptr` / `out_len`. Pointer lifetime matches the Editor.
///
/// Returns 0 on success (including "field absent", which yields
/// `out_len = 0`), -1 on an unknown field id, -2 when the properties
/// could not be read at all.
///
/// Absent and empty are both reported as `out_len = 0`. The
/// distinction exists in the Zig API but is not worth an extra ABI
/// slot: no caller has needed to act on it, and a scrub treats them
/// identically.
export fn zlsx_editor_docprop_at(
    ed: *Editor,
    field: u32,
    out_ptr: *[*]const u8,
    out_len: *usize,
) callconv(.c) i32 {
    const state: *EditorState = @ptrCast(@alignCast(ed));
    const props = state.inner.docProps() catch return -2;

    const v: ?[]const u8 = switch (field) {
        ZLSX_DOCPROP_CREATOR => props.creator,
        ZLSX_DOCPROP_LAST_MODIFIED_BY => props.last_modified_by,
        ZLSX_DOCPROP_TITLE => props.title,
        ZLSX_DOCPROP_SUBJECT => props.subject,
        ZLSX_DOCPROP_DESCRIPTION => props.description,
        ZLSX_DOCPROP_KEYWORDS => props.keywords,
        ZLSX_DOCPROP_CATEGORY => props.category,
        ZLSX_DOCPROP_CREATED => props.created,
        ZLSX_DOCPROP_MODIFIED => props.modified,
        ZLSX_DOCPROP_REVISION => props.revision,
        ZLSX_DOCPROP_COMPANY => props.company,
        ZLSX_DOCPROP_MANAGER => props.manager,
        ZLSX_DOCPROP_APPLICATION => props.application,
        ZLSX_DOCPROP_HYPERLINK_BASE => props.hyperlink_base,
        else => return -1,
    };

    if (v) |slice| {
        out_ptr.* = slice.ptr;
        out_len.* = slice.len;
    } else {
        out_len.* = 0;
    }
    return 0;
}

/// Non-zero when the archive carries a `docProps/custom.xml` part.
/// Returns -1 if the properties could not be read.
export fn zlsx_editor_has_custom_properties(ed: *Editor) callconv(.c) i32 {
    const state: *EditorState = @ptrCast(@alignCast(ed));
    const props = state.inner.docProps() catch return -1;
    return if (props.has_custom_properties) 1 else 0;
}

/// Strip identifying document metadata, staged for the next save.
///
/// `strip_timestamps` also removes `dcterms:created` / `dcterms:modified`
/// / `cp:revision`, which the default mask keeps — they are rarely
/// identifying alone and removing them visibly empties Excel's
/// document-info pane.
///
/// Returns 0 on success, -1 on failure (`err_buf` populated).
export fn zlsx_editor_strip_doc_props(
    ed: *Editor,
    strip_timestamps: i32,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) i32 {
    const state: *EditorState = @ptrCast(@alignCast(ed));
    const mask: zlsx_pkg.DocPropsMask = if (strip_timestamps != 0)
        .{ .application = true, .created = true, .modified = true, .revision = true }
    else
        .{};
    state.inner.stripDocProps(mask) catch |e| {
        writeError(err_buf, err_buf_len, @errorName(e));
        return -1;
    };
    return 0;
}

// ─── E5: embeddings ──────────────────────────────────────────────────
//
// A dedicated read-only handle rather than accessors hung off `Book`:
// embeddings live on `zlsx_pkg.Workbook` (the OPC part model), while
// `Book` is the streaming cell reader. Bolting one onto the other would
// mean opening and holding both models for every caller who wants
// either.
//
// The state split is the whole point of the surface. A consumer must be
// able to tell "vectors here" from "vectors were stripped by some tool,
// and here is what they were" from "never had any" — see
// docs/plans/embeddings-in-xlsx.md §Durability contract.

/// Opaque embeddings handle. Created by `zlsx_emb_open`, freed by
/// `zlsx_emb_close`.
pub const Emb = extern struct { _opaque: u8 };

pub const ZLSX_EMB_ABSENT: u32 = 0;
pub const ZLSX_EMB_PRESENT: u32 = 1;
pub const ZLSX_EMB_STRIPPED: u32 = 2;

pub const ZLSX_EMB_CARRIER_DEFINED_NAME: u32 = 0;
pub const ZLSX_EMB_CARRIER_DOC_PROPS: u32 = 1;
/// Opt-in carrier: the only one Apple Numbers preserves.
pub const ZLSX_EMB_CARRIER_CELL_DATA: u32 = 2;

const EmbState = struct {
    inner: zlsx_pkg.Workbook,
    threaded: std.Io.Threaded,
    /// Snapshot taken at open. Held rather than re-queried per call so
    /// every accessor sees one consistent answer, and so a caller
    /// cannot observe the state change under it.
    state: zlsx_pkg.EmbeddingState,
};

/// Open an .xlsx and resolve its embedding state. Returns NULL on an
/// I/O or parse failure; an absent or stripped embedding set is a
/// successful open, reported through `zlsx_emb_state`.
export fn zlsx_emb_open(
    path_ptr: [*:0]const u8,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) ?*Emb {
    const path = std.mem.span(path_ptr);
    const st = gpa.create(EmbState) catch {
        writeError(err_buf, err_buf_len, "OutOfMemory");
        return null;
    };
    st.threaded = .init(gpa, .{});
    st.inner = zlsx_pkg.Workbook.open(gpa, st.threaded.io(), path) catch |e| {
        st.threaded.deinit();
        gpa.destroy(st);
        writeError(err_buf, err_buf_len, @errorName(e));
        return null;
    };
    st.state = st.inner.embeddings() catch |e| {
        st.inner.deinit();
        st.threaded.deinit();
        gpa.destroy(st);
        writeError(err_buf, err_buf_len, @errorName(e));
        return null;
    };
    return @ptrCast(st);
}

export fn zlsx_emb_close(h: ?*Emb) callconv(.c) void {
    if (h) |p| {
        const st: *EmbState = @ptrCast(@alignCast(p));
        st.inner.deinit();
        st.threaded.deinit();
        gpa.destroy(st);
    }
}

export fn zlsx_emb_state(h: *Emb) callconv(.c) u32 {
    const st: *EmbState = @ptrCast(@alignCast(h));
    return switch (st.state) {
        .absent => ZLSX_EMB_ABSENT,
        .present => ZLSX_EMB_PRESENT,
        .stripped => ZLSX_EMB_STRIPPED,
    };
}

fn copyOut(s: []const u8, out_buf: [*]u8, out_buf_len: usize) usize {
    if (out_buf_len == 0) return s.len;
    const n = @min(s.len, out_buf_len - 1);
    @memcpy(out_buf[0..n], s[0..n]);
    out_buf[n] = 0;
    return s.len;
}

/// Model provenance. Available for PRESENT *and* STRIPPED — recovering
/// it after a strip is the entire reason the ER record exists.
export fn zlsx_emb_model(h: *Emb, out_buf: [*]u8, out_buf_len: usize) callconv(.c) usize {
    const st: *EmbState = @ptrCast(@alignCast(h));
    return switch (st.state) {
        .present => |v| copyOut(v.index.model, out_buf, out_buf_len),
        .stripped => |r| copyOut(r.model, out_buf, out_buf_len),
        .absent => 0,
    };
}

export fn zlsx_emb_dim(h: *Emb) callconv(.c) u32 {
    const st: *EmbState = @ptrCast(@alignCast(h));
    return switch (st.state) {
        .present => |v| v.index.dim,
        .stripped => |r| r.dim,
        .absent => 0,
    };
}

export fn zlsx_emb_dtype(h: *Emb, out_buf: [*]u8, out_buf_len: usize) callconv(.c) usize {
    const st: *EmbState = @ptrCast(@alignCast(h));
    return switch (st.state) {
        .present => |v| copyOut(v.index.dtype.string(), out_buf, out_buf_len),
        .stripped => |r| copyOut(r.dtype, out_buf, out_buf_len),
        .absent => 0,
    };
}

export fn zlsx_emb_coverage_count(h: *Emb) callconv(.c) usize {
    const st: *EmbState = @ptrCast(@alignCast(h));
    return switch (st.state) {
        .present => |v| v.coverages.len,
        .stripped => |r| r.coverages.len,
        .absent => 0,
    };
}

export fn zlsx_emb_coverage_id(h: *Emb, i: usize, out_buf: [*]u8, out_buf_len: usize) callconv(.c) usize {
    const st: *EmbState = @ptrCast(@alignCast(h));
    return switch (st.state) {
        .present => |v| if (i < v.coverages.len)
            copyOut(v.coverages[i].coverage.id, out_buf, out_buf_len)
        else
            0,
        .stripped => |r| if (i < r.coverages.len)
            copyOut(r.coverages[i].id, out_buf, out_buf_len)
        else
            0,
        .absent => 0,
    };
}

export fn zlsx_emb_coverage_range(h: *Emb, i: usize, out_buf: [*]u8, out_buf_len: usize) callconv(.c) usize {
    const st: *EmbState = @ptrCast(@alignCast(h));
    return switch (st.state) {
        .present => |v| if (i < v.coverages.len)
            copyOut(v.coverages[i].coverage.range, out_buf, out_buf_len)
        else
            0,
        .stripped => |r| if (i < r.coverages.len)
            copyOut(r.coverages[i].range, out_buf, out_buf_len)
        else
            0,
        .absent => 0,
    };
}

export fn zlsx_emb_coverage_sheet(h: *Emb, i: usize, out_buf: [*]u8, out_buf_len: usize) callconv(.c) usize {
    const st: *EmbState = @ptrCast(@alignCast(h));
    return switch (st.state) {
        .present => |v| if (i < v.coverages.len)
            copyOut(v.coverages[i].coverage.worksheet_target, out_buf, out_buf_len)
        else
            0,
        .stripped => |r| if (i < r.coverages.len)
            copyOut(r.coverages[i].worksheet_target, out_buf, out_buf_len)
        else
            0,
        .absent => 0,
    };
}

/// Row count for coverage `i` — the number of vectors, present or
/// stripped. Sizes the caller's output buffers.
export fn zlsx_emb_coverage_rows(h: *Emb, i: usize) callconv(.c) u32 {
    const st: *EmbState = @ptrCast(@alignCast(h));
    return switch (st.state) {
        .present => |v| if (i < v.coverages.len) v.coverages[i].vec.header.count else 0,
        .stripped => |r| if (i < r.coverages.len) r.coverages[i].count else 0,
        .absent => 0,
    };
}

/// Content fingerprint recorded at embed time. STRIPPED only; 0
/// otherwise. Recomputable from the current cells, so equal means the
/// covered content has not drifted and a re-embed reproduces the same
/// vectors.
export fn zlsx_emb_digest(h: *Emb) callconv(.c) u64 {
    const st: *EmbState = @ptrCast(@alignCast(h));
    return switch (st.state) {
        .stripped => |r| r.digest,
        else => 0,
    };
}

/// Which carrier the recovery record came from. STRIPPED only.
export fn zlsx_emb_carrier(h: *Emb) callconv(.c) u32 {
    const st: *EmbState = @ptrCast(@alignCast(h));
    return switch (st.state) {
        .stripped => |r| switch (r.carrier) {
            .defined_name => ZLSX_EMB_CARRIER_DEFINED_NAME,
            .doc_props => ZLSX_EMB_CARRIER_DOC_PROPS,
            .cell_data => ZLSX_EMB_CARRIER_CELL_DATA,
        },
        else => 0,
    };
}

/// The hash value that marks a deleted row. Exposed rather than left
/// for the binding to hard-code, so the tombstone contract has one
/// definition.
export fn zlsx_emb_tombstone() callconv(.c) u64 {
    return zlsx_pkg.embedding_part.TOMBSTONE_HASH;
}

/// Decode coverage `i`'s vectors into `out` as f32, row-major
/// `[rows][dim]`. `out_len` must be exactly `rows * dim`.
///
/// One call per coverage: a 500-row × 1536-dim coverage would otherwise
/// be 500 FFI crossings, and the dtype layout stays on this side rather
/// than being reimplemented in every binding.
///
/// Returns 0 on success, -1 on a bad index or size, -2 when the state
/// is not PRESENT (a stripped coverage has provenance but no vectors —
/// the distinction the caller must not paper over).
export fn zlsx_emb_vectors(h: *Emb, i: usize, out: [*]f32, out_len: usize) callconv(.c) i32 {
    const st: *EmbState = @ptrCast(@alignCast(h));
    const v = switch (st.state) {
        .present => |x| x,
        else => return -2,
    };
    if (i >= v.coverages.len) return -1;
    zlsx_pkg.embedding_part.decodeAllF32(v.coverages[i].vec, out[0..out_len]) catch return -1;
    return 0;
}

/// Copy coverage `i`'s per-row content hashes into `out`. `out_len`
/// must be exactly `rows`. Same return convention as
/// `zlsx_emb_vectors`.
///
/// A row whose hash equals `zlsx_emb_tombstone()` was deleted; that is
/// what a binding turns into a validity mask.
export fn zlsx_emb_hashes(h: *Emb, i: usize, out: [*]u64, out_len: usize) callconv(.c) i32 {
    const st: *EmbState = @ptrCast(@alignCast(h));
    const v = switch (st.state) {
        .present => |x| x,
        else => return -2,
    };
    if (i >= v.coverages.len) return -1;
    const hp = v.coverages[i].hashes;
    if (out_len != hp.header.count) return -1;
    var k: usize = 0;
    while (k < out_len) : (k += 1) {
        out[k] = hp.value(@intCast(k)) catch return -1;
    }
    return 0;
}

// ─── Formula engine (M9a1) ───────────────────────────────────────────
//
// `zlsx_status_v1` + descriptor types + editor recalc/evaluate over the
// M5d2 pipeline. The committed layout contract is
// docs/plans/c-abi-status-v1.md; every offset it pins is asserted at
// comptime below. Legacy exports above keep their shipped 0/-1
// convention untouched — the status contract applies to NEW exports
// only.

const recalc_run = zlsx_pkg.recalc_run;
const RunInputs = zlsx_pkg.RunInputs;
const ResourceLimits = @FieldType(RunInputs, "limits");
const Fidelity = @FieldType(RunInputs, "fidelity");
const PlatformProfile = @FieldType(RunInputs, "platform_profile");
const FormulaDialect = @FieldType(RunInputs, "dialect");
const PlaneTwo = zlsx_pkg.recalc_txn.PlaneTwo;
const recalc_txn = zlsx_pkg.recalc_txn;
const EvaluateOptions = zlsx_pkg.EvaluateOptions;
const EvalValue = @FieldType(zlsx_pkg.Evaluation, "value");
const ScalarV = @FieldType(EvalValue, "scalar");
const parse_limits_default: @FieldType(EvaluateOptions, "parse_limits") = .{};

/// `zlsx_status_v1` (§12.3). New exports only.
pub const ZLSX_OK: i32 = 0;
pub const ZLSX_ERROR: i32 = -1;
pub const ZLSX_REFUSED: i32 = -2;
pub const ZLSX_NOMEM: i32 = -3;
// -4 reserved, never returned by v1.
pub const ZLSX_CANCELLED: i32 = -5;

/// `zlsx_diag_v1.plane` when the diag carries no Plane-2 refusal, and
/// `zlsx_resolved_v1.dialect` for a recalc (which derives dialect per
/// stored cell and normalizes it out of the fingerprint, §5.3b).
const plane_none: u32 = 0xFFFF_FFFF;
const dialect_none: u32 = 0xFFFF_FFFF;

const CCensusEntry = extern struct {
    plane: u32,
    sheet: u32,
    /// One-based, as OOXML writes it; 0 = not about a cell.
    row: u32,
    /// Zero-based.
    col: u32,
};

const CDiag = extern struct {
    struct_size: usize,
    plane: u32,
    census_truncated: u32,
    error_name: [64]u8,
    census: ?[*]const CCensusEntry,
    census_len: usize,
};

const CResolved = extern struct {
    struct_size: usize,
    now_utc_ms: i64,
    rng_seed: u64,
    utc_offset_min: i32,
    fidelity: u32,
    profile: u32,
    dialect: u32,
    max_run_arena_bytes: u64,
    max_matrix_cells: u64,
    max_string_payload_bytes: u64,
    max_retained_ast_bytes: u64,
    max_diagnostics_bytes: u64,
};

const CRun = extern struct {
    struct_size: usize,
    now_utc_ms: i64,
    rng_seed: u64,
    utc_offset_min: i32,
    fidelity: u32,
    profile: u32,
    dialect: u32,
    on_unsupported: u32,
    _reserved0: u32,
    max_run_arena_bytes: u64,
    max_matrix_cells: u64,
    max_string_payload_bytes: u64,
    max_retained_ast_bytes: u64,
    max_diagnostics_bytes: u64,
    timeout_ms: u64,
    cancel: ?*CancelTok,
};

const CRecalcReport = extern struct {
    struct_size: usize,
    sheets_patched: u32,
    cells_written: u32,
    passes: u32,
    non_converged_cells: u32,
    dynamic_passes: u32,
    kept_stale: u32,
    calc_chain_removed: u32,
    census_truncated: u32,
    retained_generations: u64,
    retained_bytes: u64,
    /// §5.7.9's dormant durability slot — 0 for the in-memory
    /// transaction; M9a2's save export fills it.
    durability_warning: u32,
    durability_errno: i32,
    resolved: CResolved,
    resolved_present: u32,
    _reserved0: u32,
    census: ?[*]const CCensusEntry,
    census_len: usize,
};

const CValueElem = extern struct {
    tag: u8,
    _reserved: [7]u8,
    num: f64,
    payload_off: u64,
    payload_len: u64,
};

const CValue = extern struct {
    struct_size: usize,
    rows: u32,
    cols: u32,
    is_matrix: u32,
    _reserved0: u32,
    elems: ?[*]const CValueElem,
    elems_len: usize,
    payload: ?[*]const u8,
    payload_len: usize,
};

const value_tag_number: u8 = 0;
const value_tag_text: u8 = 1;
const value_tag_bool: u8 = 2;
const value_tag_error: u8 = 3;

// Every offset the design note pins, enforced where the layout lives.
// A drifted field is a compile error, not a corrupted caller.
comptime {
    const assert = std.debug.assert;
    assert(@sizeOf(CCensusEntry) == 16);
    assert(@offsetOf(CDiag, "plane") == 8);
    assert(@offsetOf(CDiag, "error_name") == 16);
    assert(@offsetOf(CDiag, "census") == 80);
    assert(@offsetOf(CDiag, "census_len") == 88);
    assert(@sizeOf(CDiag) == 96);
    assert(@offsetOf(CResolved, "utc_offset_min") == 24);
    assert(@offsetOf(CResolved, "max_run_arena_bytes") == 40);
    assert(@sizeOf(CResolved) == 80);
    assert(@offsetOf(CRun, "utc_offset_min") == 24);
    assert(@offsetOf(CRun, "on_unsupported") == 40);
    assert(@offsetOf(CRun, "max_run_arena_bytes") == 48);
    assert(@offsetOf(CRun, "timeout_ms") == 88);
    assert(@offsetOf(CRun, "cancel") == 96);
    assert(@sizeOf(CRun) == 104);
    assert(@offsetOf(CRecalcReport, "retained_generations") == 40);
    assert(@offsetOf(CRecalcReport, "durability_warning") == 56);
    assert(@offsetOf(CRecalcReport, "resolved") == 64);
    assert(@offsetOf(CRecalcReport, "resolved_present") == 144);
    assert(@offsetOf(CRecalcReport, "census") == 152);
    assert(@offsetOf(CRecalcReport, "census_len") == 160);
    assert(@sizeOf(CRecalcReport) == 168);
    assert(@offsetOf(CValueElem, "num") == 8);
    assert(@offsetOf(CValueElem, "payload_off") == 16);
    assert(@offsetOf(CValueElem, "payload_len") == 24);
    assert(@sizeOf(CValueElem) == 32);
    assert(@offsetOf(CValue, "rows") == 8);
    assert(@offsetOf(CValue, "elems") == 24);
    assert(@offsetOf(CValue, "payload") == 40);
    assert(@sizeOf(CValue) == 56);
}

/// Opaque cancel token. Heap-allocated atomic; `trigger` is a release
/// store callable from any thread — the reason R9-12 pins this module
/// multi-threaded.
pub const CancelTok = extern struct { _opaque: u8 };

const CancelTokenState = struct { flag: std.atomic.Value(bool) };

export fn zlsx_cancel_token_new(
    out: ?*?*CancelTok,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) i32 {
    const slot = out orelse {
        writeError(err_buf, err_buf_len, "InvalidInput");
        return ZLSX_ERROR;
    };
    slot.* = null;
    const state = gpa.create(CancelTokenState) catch {
        writeError(err_buf, err_buf_len, "OutOfMemory");
        return ZLSX_NOMEM;
    };
    state.* = .{ .flag = .{ .raw = false } };
    slot.* = @ptrCast(state);
    return ZLSX_OK;
}

/// Thread-safe; any thread. Triggering an already-triggered token is a
/// no-op. NULL-safe.
export fn zlsx_cancel_token_trigger(tok: ?*CancelTok) callconv(.c) void {
    const state: *CancelTokenState = @ptrCast(@alignCast(tok orelse return));
    state.flag.store(true, .release);
}

/// Caller-owned lifetime: the token must outlive every call it was
/// passed to. NULL-safe.
export fn zlsx_cancel_token_free(tok: ?*CancelTok) callconv(.c) void {
    const state: *CancelTokenState = @ptrCast(@alignCast(tok orelse return));
    gpa.destroy(state);
}

/// §12.4's engine identity: semantic version + the three rule versions
/// + target triple + build hash, one static string. M9b keys Spark's
/// mixed-fleet refusal on it; the bare version string above is a
/// component, not the identity.
const engine_fingerprint_str = std.fmt.comptimePrint(
    "zlsx {s}; {s}; {s}; {s}; {s}-{s}-{s}; {s}",
    .{
        build_options.version,
        recalc_run.rule_versions.excel_fp,
        recalc_run.rule_versions.rng,
        recalc_run.rule_versions.collation,
        @tagName(builtin.target.cpu.arch),
        @tagName(builtin.target.os.tag),
        @tagName(builtin.target.abi),
        fingerprint_config.build_hash,
    },
);

export fn zlsx_engine_fingerprint() callconv(.c) [*:0]const u8 {
    return engine_fingerprint_str;
}

/// §12.3's struct_size discipline for output structs: reject below the
/// v1 minimum (before any byte is written), then zero exactly
/// [after struct_size, known) — never a byte beyond the known prefix,
/// even when the caller declared more.
fn prepOut(comptime T: type, ptr: *T, err_buf: ?[*]u8, err_buf_len: usize) bool {
    if (ptr.struct_size < @sizeOf(T)) {
        writeError(err_buf, err_buf_len, "StructSizeTooSmall");
        return false;
    }
    const bytes: [*]u8 = @ptrCast(ptr);
    @memset(bytes[@sizeOf(usize)..@sizeOf(T)], 0);
    return true;
}

fn checkIn(comptime T: type, ptr: *const T, err_buf: ?[*]u8, err_buf_len: usize) bool {
    if (ptr.struct_size < @sizeOf(T)) {
        writeError(err_buf, err_buf_len, "StructSizeTooSmall");
        return false;
    }
    return true;
}

fn prepDiag(diag: ?*CDiag, err_buf: ?[*]u8, err_buf_len: usize) bool {
    const d = diag orelse return true;
    if (!prepOut(CDiag, d, err_buf, err_buf_len)) return false;
    d.plane = plane_none;
    return true;
}

/// One error→status mapping (design note §2): OOM, cancellation, then
/// the fourteen-plane §10 vocabulary and the S3a structural vocabulary
/// (`structural_refusals`) as typed refusals; everything else is a
/// generic -1 with the error name in errbuf.
fn statusOf(e: anyerror) i32 {
    switch (e) {
        error.OutOfMemory => return ZLSX_NOMEM,
        error.Cancelled => return ZLSX_CANCELLED,
        else => {},
    }
    inline for (@typeInfo(PlaneTwo).@"enum".fields) |f| {
        if (std.mem.eql(u8, f.name, @errorName(e))) return ZLSX_REFUSED;
    }
    // S3a: the structural edits' and the pivots read's own vocabulary —
    // no plane, but a statement about the workbook all the same.
    if (isStructuralRefusal(e)) return ZLSX_REFUSED;
    return ZLSX_ERROR;
}

fn diagSetError(diag: ?*CDiag, name: []const u8) void {
    const d = diag orelse return;
    const n = @min(name.len, d.error_name.len - 1);
    @memcpy(d.error_name[0..n], name[0..n]);
    d.error_name[n] = 0;
    inline for (@typeInfo(PlaneTwo).@"enum".fields) |f| {
        if (std.mem.eql(u8, f.name, name)) d.plane = f.value;
    }
}

fn failMapped(e: anyerror, diag: ?*CDiag, err_buf: ?[*]u8, err_buf_len: usize) i32 {
    const status = statusOf(e);
    writeError(err_buf, err_buf_len, @errorName(e));
    if (status == ZLSX_REFUSED) diagSetError(diag, @errorName(e));
    return status;
}

fn refuseNamed(name: []const u8, diag: ?*CDiag, err_buf: ?[*]u8, err_buf_len: usize) i32 {
    writeError(err_buf, err_buf_len, name);
    diagSetError(diag, name);
    return ZLSX_REFUSED;
}

/// Narrow a `zlsx_run_v1` into the engine's `RunInputs`. Unknown enum
/// values and out-of-range fields are ABI-contract violations (-1),
/// not refusals — they are statements about the call.
fn runFromC(crun: *const CRun, io: std.Io, err_buf: ?[*]u8, err_buf_len: usize) ?RunInputs {
    const fidelity: Fidelity = switch (crun.fidelity) {
        0 => .excel,
        1 => .ieee,
        else => {
            writeError(err_buf, err_buf_len, "InvalidInput");
            return null;
        },
    };
    const profile: PlatformProfile = switch (crun.profile) {
        0 => .windows_1252,
        else => {
            writeError(err_buf, err_buf_len, "InvalidInput");
            return null;
        },
    };
    const dialect: FormulaDialect = switch (crun.dialect) {
        0 => .dynamic_array,
        1 => .legacy,
        else => {
            writeError(err_buf, err_buf_len, "InvalidInput");
            return null;
        },
    };
    var limits: ResourceLimits = .{};
    if (crun.max_run_arena_bytes != 0) limits.max_run_arena_bytes = crun.max_run_arena_bytes;
    if (crun.max_matrix_cells != 0) limits.max_matrix_cells = crun.max_matrix_cells;
    if (crun.max_string_payload_bytes != 0) limits.max_string_payload_bytes = crun.max_string_payload_bytes;
    if (crun.max_retained_ast_bytes != 0) limits.max_retained_ast_bytes = crun.max_retained_ast_bytes;
    if (crun.max_diagnostics_bytes != 0) limits.max_diagnostics_bytes = crun.max_diagnostics_bytes;
    var ri: RunInputs = .{
        .now_utc_ms = crun.now_utc_ms,
        .rng_seed = crun.rng_seed,
        .limits = limits,
        .utc_offset_min = crun.utc_offset_min,
        .fidelity = fidelity,
        .platform_profile = profile,
        .dialect = dialect,
    };
    if (crun.cancel) |tok| {
        const cstate: *CancelTokenState = @ptrCast(@alignCast(tok));
        ri.cancel = .{ .atomic = &cstate.flag };
    }
    if (crun.timeout_ms != 0) {
        const now = std.Io.Timestamp.now(io, .awake);
        const Ns = @TypeOf(now.nanoseconds);
        ri.deadline = .{ .nanoseconds = now.nanoseconds + @as(Ns, @intCast(crun.timeout_ms)) * 1_000_000 };
    }
    ri.validate() catch |e| {
        writeError(err_buf, err_buf_len, @errorName(e));
        return null;
    };
    return ri;
}

fn fillResolved(dst: *CResolved, eff: anytype) void {
    dst.struct_size = @sizeOf(CResolved);
    dst.now_utc_ms = eff.now_utc_ms;
    dst.rng_seed = eff.rng_seed;
    dst.utc_offset_min = eff.utc_offset_min;
    dst.fidelity = switch (eff.fidelity) {
        .excel => 0,
        .ieee => 1,
    };
    dst.profile = switch (eff.platform_profile) {
        .windows_1252 => 0,
    };
    dst.dialect = if (eff.dialect) |d| switch (d) {
        .dynamic_array => @as(u32, 0),
        .legacy => 1,
    } else dialect_none;
    dst.max_run_arena_bytes = eff.limits.max_run_arena_bytes;
    dst.max_matrix_cells = eff.limits.max_matrix_cells;
    dst.max_string_payload_bytes = eff.limits.max_string_payload_bytes;
    dst.max_retained_ast_bytes = eff.limits.max_retained_ast_bytes;
    dst.max_diagnostics_bytes = eff.limits.max_diagnostics_bytes;
}

fn censusToC(census: anytype) error{OutOfMemory}!?[]CCensusEntry {
    if (census.len == 0) return null;
    const arr = try gpa.alloc(CCensusEntry, census.len);
    for (census, arr) |u, *dst| dst.* = .{
        .plane = @intFromEnum(u.plane),
        .sheet = u.sheet,
        .row = u.row,
        .col = u.col,
    };
    return arr;
}

fn payloadLenOf(p: recalc_run.PublishedScalar) usize {
    return switch (p) {
        .text => |t| t.len,
        .err => |e| e.spelling().len,
        else => 0,
    };
}

fn pubElem(p: recalc_run.PublishedScalar, payload: []u8, cursor: *usize) CValueElem {
    var elem: CValueElem = .{
        .tag = value_tag_number,
        ._reserved = @splat(0),
        .num = 0,
        .payload_off = 0,
        .payload_len = 0,
    };
    switch (p) {
        .number => |v| elem.num = v,
        .boolean => |b| {
            elem.tag = value_tag_bool;
            elem.num = if (b) 1 else 0;
        },
        .text => |t| {
            elem.tag = value_tag_text;
            @memcpy(payload[cursor.* .. cursor.* + t.len], t);
            elem.payload_off = cursor.*;
            elem.payload_len = t.len;
            cursor.* += t.len;
        },
        .err => |e| {
            const s = e.spelling();
            elem.tag = value_tag_error;
            @memcpy(payload[cursor.* .. cursor.* + s.len], s);
            elem.payload_off = cursor.*;
            elem.payload_len = s.len;
            cursor.* += s.len;
        },
    }
    return elem;
}

/// Copy an evaluated value into C-owned descriptors. Blank never
/// crosses: §5.3a's publish seam converts as the elements are copied.
fn buildValue(out: *CValue, value: EvalValue, fidelity: Fidelity) error{OutOfMemory}!void {
    switch (value) {
        .array => |m| {
            const n = m.cells.len;
            var payload_bytes: usize = 0;
            for (m.cells) |s| payload_bytes += payloadLenOf(recalc_run.publish(s, fidelity));
            const elems = try gpa.alloc(CValueElem, n);
            errdefer gpa.free(elems);
            const payload: []u8 = if (payload_bytes > 0) try gpa.alloc(u8, payload_bytes) else &.{};
            var cursor: usize = 0;
            for (m.cells, elems) |s, *dst| {
                dst.* = pubElem(recalc_run.publish(s, fidelity), payload, &cursor);
            }
            out.rows = @intCast(m.rows);
            out.cols = @intCast(m.cols);
            out.is_matrix = 1;
            out.elems = elems.ptr;
            out.elems_len = n;
            if (payload_bytes > 0) {
                out.payload = payload.ptr;
                out.payload_len = payload_bytes;
            }
        },
        .reference => unreachable, // dereferenced before return (§5.3b)
        else => {
            const scalar: ScalarV = switch (value) {
                .scalar => |s| s,
                .missing_arg => .blank,
                else => unreachable,
            };
            const p = recalc_run.publish(scalar, fidelity);
            const payload_bytes = payloadLenOf(p);
            const elems = try gpa.alloc(CValueElem, 1);
            errdefer gpa.free(elems);
            const payload: []u8 = if (payload_bytes > 0) try gpa.alloc(u8, payload_bytes) else &.{};
            var cursor: usize = 0;
            elems[0] = pubElem(p, payload, &cursor);
            out.rows = 1;
            out.cols = 1;
            out.is_matrix = 0;
            out.elems = elems.ptr;
            out.elems_len = 1;
            if (payload_bytes > 0) {
                out.payload = payload.ptr;
                out.payload_len = payload_bytes;
            }
        },
    }
}

/// Release the library-owned interior of a `zlsx_value_v1`. The struct
/// itself is the caller's. NULL-safe; safe on a zeroed struct; resets
/// the released fields so a second call is a no-op.
export fn zlsx_value_release(v: ?*CValue) callconv(.c) void {
    const val = v orelse return;
    if (val.elems) |ptr| gpa.free(ptr[0..val.elems_len]);
    val.elems = null;
    val.elems_len = 0;
    if (val.payload) |ptr| {
        if (val.payload_len > 0) gpa.free(ptr[0..val.payload_len]);
    }
    val.payload = null;
    val.payload_len = 0;
}

export fn zlsx_recalc_report_release(r: ?*CRecalcReport) callconv(.c) void {
    const rep = r orelse return;
    if (rep.census) |ptr| gpa.free(ptr[0..rep.census_len]);
    rep.census = null;
    rep.census_len = 0;
}

export fn zlsx_diag_release(d: ?*CDiag) callconv(.c) void {
    const diag = d orelse return;
    if (diag.census) |ptr| gpa.free(ptr[0..diag.census_len]);
    diag.census = null;
    diag.census_len = 0;
}

/// §5.7.7's mark-only transaction: keep every cache, set
/// `fullCalcOnLoad="1"`, remove nothing else. Refusals (e.g.
/// `FormulaPrecisionAsDisplayed`) are typed -2 with the diag populated.
export fn zlsx_editor_mark_recalc_on_load(
    ed: ?*Editor,
    diag: ?*CDiag,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) i32 {
    const state: *EditorState = @ptrCast(@alignCast(ed orelse {
        writeError(err_buf, err_buf_len, "InvalidInput");
        return ZLSX_ERROR;
    }));
    if (!prepDiag(diag, err_buf, err_buf_len)) return ZLSX_ERROR;
    state.inner.workbook.markRecalcOnLoad() catch |e| {
        return failMapped(e, diag, err_buf, err_buf_len);
    };
    return ZLSX_OK;
}

/// §5.7's in-memory transaction over the M5d2 pipeline: recalculate
/// every formula cell and swap the result in as the final operation.
/// On refusal, cancellation or allocation failure the workbook is
/// exactly as it was. No file is opened or written.
export fn zlsx_editor_recalculate(
    ed: ?*Editor,
    run: ?*const CRun,
    report: ?*CRecalcReport,
    diag: ?*CDiag,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) i32 {
    const state: *EditorState = @ptrCast(@alignCast(ed orelse {
        writeError(err_buf, err_buf_len, "InvalidInput");
        return ZLSX_ERROR;
    }));
    const crun = run orelse {
        writeError(err_buf, err_buf_len, "InvalidInput");
        return ZLSX_ERROR;
    };
    const out = report orelse {
        writeError(err_buf, err_buf_len, "InvalidInput");
        return ZLSX_ERROR;
    };
    // Prep every output first — each gated on its own struct_size and
    // zeroed independently, so a failure anywhere (including a sibling
    // struct's size) leaves every ACCEPTED output zeroed and therefore
    // releasable. A rejected struct is left byte-for-byte untouched.
    var outputs_ok = prepOut(CRecalcReport, out, err_buf, err_buf_len);
    if (outputs_ok) out.resolved.struct_size = @sizeOf(CResolved);
    if (!prepDiag(diag, err_buf, err_buf_len)) outputs_ok = false;
    if (!outputs_ok) return ZLSX_ERROR;
    if (!checkIn(CRun, crun, err_buf, err_buf_len)) return ZLSX_ERROR;

    const io = state.threaded.io();
    const ri = runFromC(crun, io, err_buf, err_buf_len) orelse return ZLSX_ERROR;
    const opts: recalc_run.Options = .{
        .on_unsupported = switch (crun.on_unsupported) {
            0 => .refuse,
            1 => .keep_stale_and_mark,
            else => {
                writeError(err_buf, err_buf_len, "InvalidInput");
                return ZLSX_ERROR;
            },
        },
    };

    var refusal: recalc_txn.Refusal = .{ .reason = .unsupported_construct };
    var seamed = opts;
    seamed.refusal_out = &refusal;
    var rep = state.inner.workbook.recalculate(gpa, io, ri, seamed) catch |e| {
        return failMappedRefusal(e, &refusal, diag, err_buf, err_buf_len);
    };
    defer rep.deinit(gpa);

    reportToC(out, &rep) catch {
        writeError(err_buf, err_buf_len, "OutOfMemory");
        return ZLSX_NOMEM;
    };
    return ZLSX_OK;
}

/// The report copy-out shared by `zlsx_editor_recalculate`,
/// `zlsx_editor_save_with_recalc` and `zlsx_writer_save_with_recalc`.
/// `out` was prepped (zeroed) by the caller; census allocation is the
/// only fallible step and lands last, so an OOM here leaves a report
/// that is complete except for the census and still releasable.
fn reportToC(out: *CRecalcReport, rep: *const recalc_run.Report) error{OutOfMemory}!void {
    out.sheets_patched = rep.sheets_patched;
    out.cells_written = rep.cells_written;
    out.passes = rep.passes;
    out.non_converged_cells = rep.non_converged_cells;
    out.dynamic_passes = rep.dynamic_passes;
    out.kept_stale = @intFromBool(rep.kept_stale);
    out.calc_chain_removed = @intFromBool(rep.calc_chain_removed);
    out.census_truncated = @intFromBool(rep.census_truncated);
    out.retained_generations = @intCast(rep.retained_generations);
    out.retained_bytes = rep.retained_bytes;
    out.durability_warning = @intFromBool(rep.durability.warning);
    out.durability_errno = rep.durability.err_code;
    if (rep.resolved) |eff| {
        fillResolved(&out.resolved, eff);
        out.resolved_present = 1;
    }
    if (try censusToC(rep.census)) |arr| {
        out.census = arr.ptr;
        out.census_len = arr.len;
    }
}

/// `failMapped` + M9a2's refusal seam: when the pipeline moved a
/// refusal into `refusal` (`Options.refusal_out`), its cells become
/// the -2 diag's census. Consumes the refusal on every path. A refusal
/// that never reached the seam (e.g. `run.validate`'s
/// `FormulaLimitExceeded`) carries the default empty census, so the
/// diag stays truthful: error_name + plane, no cells.
fn failMappedRefusal(
    e: anyerror,
    refusal: *recalc_txn.Refusal,
    diag: ?*CDiag,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) i32 {
    const status = failMapped(e, diag, err_buf, err_buf_len);
    var r = refusal.*;
    defer r.deinit(gpa);
    if (status != ZLSX_REFUSED) return status;
    const d = diag orelse return status;
    // Best-effort: the refusal already told the truth in error_name;
    // an OOM copying the census leaves it empty rather than turning a
    // typed refusal into an allocation failure.
    const census = censusToC(r.census) catch return status;
    if (census) |arr| {
        d.census = arr.ptr;
        d.census_len = arr.len;
        d.census_truncated = @intFromBool(r.census_truncated);
    }
    return status;
}

/// M6's `zlsx eval` semantics over `Workbook.evaluate`: a cache-based
/// standalone read. Scratch-only — the workbook's logical state and
/// serialized bytes are byte-identical before and after (§5.6f purity);
/// eval never commits, so an observed trigger is always pre-commit.
export fn zlsx_editor_evaluate(
    ed: ?*Editor,
    formula_ptr: ?[*]const u8,
    formula_len: usize,
    sheet_idx: u32,
    anchor_row: u32,
    anchor_col: u32,
    run: ?*const CRun,
    out_value: ?*CValue,
    out_resolved: ?*CResolved,
    diag: ?*CDiag,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) i32 {
    const state: *EditorState = @ptrCast(@alignCast(ed orelse {
        writeError(err_buf, err_buf_len, "InvalidInput");
        return ZLSX_ERROR;
    }));
    const crun = run orelse {
        writeError(err_buf, err_buf_len, "InvalidInput");
        return ZLSX_ERROR;
    };
    const out = out_value orelse {
        writeError(err_buf, err_buf_len, "InvalidInput");
        return ZLSX_ERROR;
    };
    // Same output-first discipline as recalculate: every accepted
    // output is zeroed before any input check can fail the call.
    var outputs_ok = prepOut(CValue, out, err_buf, err_buf_len);
    if (out_resolved) |res| {
        if (!prepOut(CResolved, res, err_buf, err_buf_len)) outputs_ok = false;
    }
    if (!prepDiag(diag, err_buf, err_buf_len)) outputs_ok = false;
    if (!outputs_ok) return ZLSX_ERROR;
    if (!checkIn(CRun, crun, err_buf, err_buf_len)) return ZLSX_ERROR;

    // Bound before the slice exists: the boundary never constructs a
    // slice it hasn't checked. One-past-limit is the fixture.
    if (formula_len > parse_limits_default.max_formula_utf8_bytes) {
        return refuseNamed("FormulaLimitExceeded", diag, err_buf, err_buf_len);
    }
    if (formula_len > 0 and formula_ptr == null) {
        writeError(err_buf, err_buf_len, "InvalidInput");
        return ZLSX_ERROR;
    }
    const formula: []const u8 = if (formula_len == 0) "" else formula_ptr.?[0..formula_len];

    const io = state.threaded.io();
    const ri = runFromC(crun, io, err_buf, err_buf_len) orelse return ZLSX_ERROR;

    var site: ?@typeInfo(@FieldType(EvaluateOptions, "site")).optional.child = null;
    if (anchor_row != 0) {
        const row = refs.Row.fromOneBased(anchor_row) catch |e| {
            writeError(err_buf, err_buf_len, @errorName(e));
            return ZLSX_ERROR;
        };
        const col = refs.Col.fromZeroBased(anchor_col) catch |e| {
            writeError(err_buf, err_buf_len, @errorName(e));
            return ZLSX_ERROR;
        };
        site = .{ .row = row, .col = col };
    }

    // Observed-before-commit, checked the way the M6 CLI checks: at
    // entry and again before results cross the boundary.
    const ctl: zlsx_pkg.Control = .{ .cancel = ri.cancel, .deadline = ri.deadline };
    ctl.check(io) catch {
        writeError(err_buf, err_buf_len, "Cancelled");
        return ZLSX_CANCELLED;
    };

    var result = state.inner.workbook.evaluate(gpa, sheet_idx, formula, .{
        .collation = recalc_run.collation_v1,
        .fidelity = ri.fidelity,
        .dialect = ri.dialect,
        .site = site,
        .now_utc_ms = ri.now_utc_ms,
        .utc_offset_min = ri.utc_offset_min,
        .platform_profile = ri.platform_profile,
    }) catch |e| {
        return failMapped(e, diag, err_buf, err_buf_len);
    };
    defer result.deinit();

    switch (result) {
        .ok => |*evaluation| {
            ctl.check(io) catch {
                writeError(err_buf, err_buf_len, "Cancelled");
                return ZLSX_CANCELLED;
            };
            buildValue(out, evaluation.value, ri.fidelity) catch {
                writeError(err_buf, err_buf_len, "OutOfMemory");
                return ZLSX_NOMEM;
            };
            if (out_resolved) |res| fillResolved(res, ri.effective(.standalone_eval));
            return ZLSX_OK;
        },
        .refused => |r| return refuseNamed(@tagName(r.planeTwo()), diag, err_buf, err_buf_len),
        .parse_refused => |r| return refuseNamed(@tagName(r.planeTwo()), diag, err_buf, err_buf_len),
        .graph_refused => |r| return refuseNamed(@tagName(r.planeTwo()), diag, err_buf, err_buf_len),
        .iteration_refused => |r| return refuseNamed(@tagName(r.planeTwo()), diag, err_buf, err_buf_len),
        .eval_refused => |r| return refuseNamed(@tagName(r.plane), diag, err_buf, err_buf_len),
    }
}

// ─── Formula engine (M9a2): buffers, the file transaction, writer ────
//
// Part 2 of the C ABI (§12.3). Same `zlsx_status_v1` contract, same
// output-prep-before-input-validation discipline as the M9a1 block
// above. Layouts documented in `docs/plans/c-abi-status-v1.md`; the
// M9a2 additions keep every M9a1 offset frozen.

/// Release a buffer an M9a2 export allocated
/// (`zlsx_editor_save_to_buffer`). NULL-safe; zero-length-safe. The
/// legacy `zlsx_buffer_free` keeps its shipped contract untouched —
/// this is the same operation under the name §12.3 pins for the
/// status-era exports.
export fn zlsx_buffer_release(ptr: ?[*]u8, len: usize) callconv(.c) void {
    const p = ptr orelse return;
    if (len == 0) return;
    gpa.free(p[0..len]);
}

/// Serialize the editor's current state — staged mutations included —
/// into a library-allocated buffer (§5.10). An untouched editor hands
/// back the source bytes verbatim (the passthrough arm `zlsx_editor_save`
/// has). No file is opened; no commit point exists. Release the buffer
/// with `zlsx_buffer_release`.
export fn zlsx_editor_save_to_buffer(
    ed: ?*Editor,
    out_ptr: ?*?[*]u8,
    out_len: ?*usize,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) i32 {
    const state: *EditorState = @ptrCast(@alignCast(ed orelse {
        writeError(err_buf, err_buf_len, "InvalidInput");
        return ZLSX_ERROR;
    }));
    const op = out_ptr orelse {
        writeError(err_buf, err_buf_len, "NullOutPointer");
        return ZLSX_ERROR;
    };
    const ol = out_len orelse {
        writeError(err_buf, err_buf_len, "NullOutPointer");
        return ZLSX_ERROR;
    };
    // Outputs prepped before anything can fail (the M9a1 fuzz-caught
    // discipline): a failed save leaves a releasable (NULL, 0) pair.
    op.* = null;
    ol.* = 0;
    const bytes = state.inner.saveToOwnedBuffer(gpa) catch |e| {
        return failMapped(e, null, err_buf, err_buf_len);
    };
    op.* = bytes.ptr;
    ol.* = bytes.len;
    return ZLSX_OK;
}

/// Open an editor over a workbook already in memory (§5.10, B4's
/// `openBuffer` with the `Book.openBuffer` borrow contract: the borrow
/// ends when this returns — `data` is duped and may be freed, reused
/// or poisoned immediately). Status-style like every M9a2 export;
/// `*out` receives the handle on ZLSX_OK and NULL otherwise. Close
/// with `zlsx_editor_close`.
export fn zlsx_open_buffer(
    data: ?[*]const u8,
    data_len: usize,
    out: ?*?*Editor,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) i32 {
    const slot = out orelse {
        writeError(err_buf, err_buf_len, "NullOutPointer");
        return ZLSX_ERROR;
    };
    slot.* = null;
    const ptr = data orelse {
        writeError(err_buf, err_buf_len, "NullBuffer");
        return ZLSX_ERROR;
    };
    // Same ownership shape as zlsx_editor_open: the handle owns its Io,
    // allocated first so the Threaded never moves after init.
    const state = gpa.create(EditorState) catch {
        writeError(err_buf, err_buf_len, "OutOfMemory");
        return ZLSX_NOMEM;
    };
    state.threaded = .init(gpa, .{});
    state.inner = zlsx_pkg.Editor.openBuffer(gpa, state.threaded.io(), ptr[0..data_len]) catch |e| {
        state.threaded.deinit();
        gpa.destroy(state);
        return failMapped(e, null, err_buf, err_buf_len);
    };
    slot.* = @ptrCast(state);
    return ZLSX_OK;
}

/// §5.7.9's file transaction over the M5d2 pipeline: serialize from
/// the prepared candidate, rename, swap in memory between the rename
/// and the directory fsync. Any failure before the rename leaves BOTH
/// the destination's prior bytes (or its absence) and the editor's
/// memory untouched; a directory fsync that fails afterwards is the
/// report's durability warning — the §5.7.9 slot goes live here —
/// never an error. A -2 refusal carries the refusing cells in the
/// diag's census (M9a2's seam through `recalc_run.prepare`).
export fn zlsx_editor_save_with_recalc(
    ed: ?*Editor,
    out_path_ptr: ?[*]const u8,
    out_path_len: usize,
    run: ?*const CRun,
    report: ?*CRecalcReport,
    diag: ?*CDiag,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) i32 {
    const state: *EditorState = @ptrCast(@alignCast(ed orelse {
        writeError(err_buf, err_buf_len, "InvalidInput");
        return ZLSX_ERROR;
    }));
    const path_ptr = out_path_ptr orelse {
        writeError(err_buf, err_buf_len, "InvalidInput");
        return ZLSX_ERROR;
    };
    const crun = run orelse {
        writeError(err_buf, err_buf_len, "InvalidInput");
        return ZLSX_ERROR;
    };
    const out = report orelse {
        writeError(err_buf, err_buf_len, "InvalidInput");
        return ZLSX_ERROR;
    };
    // Output prep before input validation, each struct on its own
    // struct_size — the M9a1 discipline verbatim.
    var outputs_ok = prepOut(CRecalcReport, out, err_buf, err_buf_len);
    if (outputs_ok) out.resolved.struct_size = @sizeOf(CResolved);
    if (!prepDiag(diag, err_buf, err_buf_len)) outputs_ok = false;
    if (!outputs_ok) return ZLSX_ERROR;
    if (out_path_len == 0) {
        writeError(err_buf, err_buf_len, "InvalidInput");
        return ZLSX_ERROR;
    }
    if (!checkIn(CRun, crun, err_buf, err_buf_len)) return ZLSX_ERROR;

    const io = state.threaded.io();
    const ri = runFromC(crun, io, err_buf, err_buf_len) orelse return ZLSX_ERROR;
    var opts: recalc_run.Options = .{
        .on_unsupported = switch (crun.on_unsupported) {
            0 => .refuse,
            1 => .keep_stale_and_mark,
            else => {
                writeError(err_buf, err_buf_len, "InvalidInput");
                return ZLSX_ERROR;
            },
        },
    };
    var refusal: recalc_txn.Refusal = .{ .reason = .unsupported_construct };
    opts.refusal_out = &refusal;

    const path = path_ptr[0..out_path_len];
    var rep = state.inner.workbook.saveWithRecalc(gpa, io, path, ri, opts) catch |e| {
        return failMappedRefusal(e, &refusal, diag, err_buf, err_buf_len);
    };
    defer rep.deinit(gpa);

    reportToC(out, &rep) catch {
        writeError(err_buf, err_buf_len, "OutOfMemory");
        return ZLSX_NOMEM;
    };
    return ZLSX_OK;
}

/// The producer-side file transaction (§12.3's `Writer.save(recalculate=)`
/// leg): emit the writer's archive to memory, open it as a workbook,
/// then run the same §5.7.9 transaction `zlsx_editor_save_with_recalc`
/// runs. One composition, shipped in `zlsx_recalc.writerSaveWithRecalc`
/// — this export is that function across the boundary. The writer
/// handle itself is not consumed and not mutated.
export fn zlsx_writer_save_with_recalc(
    w: ?*Writer,
    out_path_ptr: ?[*]const u8,
    out_path_len: usize,
    run: ?*const CRun,
    report: ?*CRecalcReport,
    diag: ?*CDiag,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) i32 {
    const state: *WriterState = @ptrCast(@alignCast(w orelse {
        writeError(err_buf, err_buf_len, "InvalidInput");
        return ZLSX_ERROR;
    }));
    const path_ptr = out_path_ptr orelse {
        writeError(err_buf, err_buf_len, "InvalidInput");
        return ZLSX_ERROR;
    };
    const crun = run orelse {
        writeError(err_buf, err_buf_len, "InvalidInput");
        return ZLSX_ERROR;
    };
    const out = report orelse {
        writeError(err_buf, err_buf_len, "InvalidInput");
        return ZLSX_ERROR;
    };
    var outputs_ok = prepOut(CRecalcReport, out, err_buf, err_buf_len);
    if (outputs_ok) out.resolved.struct_size = @sizeOf(CResolved);
    if (!prepDiag(diag, err_buf, err_buf_len)) outputs_ok = false;
    if (!outputs_ok) return ZLSX_ERROR;
    if (out_path_len == 0) {
        writeError(err_buf, err_buf_len, "InvalidInput");
        return ZLSX_ERROR;
    }
    if (!checkIn(CRun, crun, err_buf, err_buf_len)) return ZLSX_ERROR;

    const io = state.threaded.io();
    const ri = runFromC(crun, io, err_buf, err_buf_len) orelse return ZLSX_ERROR;
    var opts: zlsx_recalc.Options = .{
        .on_unsupported = switch (crun.on_unsupported) {
            0 => .refuse,
            1 => .keep_stale_and_mark,
            else => {
                writeError(err_buf, err_buf_len, "InvalidInput");
                return ZLSX_ERROR;
            },
        },
    };
    var refusal: recalc_txn.Refusal = .{ .reason = .unsupported_construct };
    opts.refusal_out = &refusal;

    const path = path_ptr[0..out_path_len];
    var rep = zlsx_recalc.writerSaveWithRecalc(gpa, io, &state.inner, path, ri, opts) catch |e| {
        return failMappedRefusal(e, &refusal, diag, err_buf, err_buf_len);
    };
    defer rep.deinit(gpa);

    reportToC(out, &rep) catch {
        writeError(err_buf, err_buf_len, "OutOfMemory");
        return ZLSX_NOMEM;
    };
    return ZLSX_OK;
}

/// §12.3's per-cell formula descriptor as it crosses the boundary —
/// 40 bytes, offsets pinned below, in `tests/c_abi_smoke.c` and in
/// `_ffi.py`. An array element like `zlsx_census_entry_v1`, so no
/// `struct_size` field; the v1 layout is frozen.
const CFormulaCell = extern struct {
    /// NULL = plain value cell (the slot's `zlsx_cell_t` stands alone).
    text: ?[*]const u8,
    text_len: usize,
    /// ZLSX_FORMULA_* tag.
    dialect: u32,
    _reserved0: u32,
    /// CSE only: the declared range, uppercase A1 (`"A1"` / `"A1:B2"`).
    /// Must be NULL for every other dialect.
    ref: ?[*]const u8,
    ref_len: usize,
};

const formula_dialect_scalar: u32 = 0;
const formula_dialect_dynamic_array: u32 = 1;
const formula_dialect_cse: u32 = 2;

comptime {
    const assert = std.debug.assert;
    assert(@offsetOf(CFormulaCell, "text") == 0);
    assert(@offsetOf(CFormulaCell, "text_len") == 8);
    assert(@offsetOf(CFormulaCell, "dialect") == 16);
    assert(@offsetOf(CFormulaCell, "_reserved0") == 20);
    assert(@offsetOf(CFormulaCell, "ref") == 24);
    assert(@offsetOf(CFormulaCell, "ref_len") == 32);
    assert(@sizeOf(CFormulaCell) == 40);
}

/// The v2 formula row (§12.3): per-cell descriptors instead of the
/// parallel text arrays `zlsx_sheet_writer_write_row_with_formulas`
/// ships — the shape a CSE rectangle needs and the text-array shape
/// cannot encode. `formulas` is parallel to `cells`; `text == NULL`
/// marks a plain value slot. Follows `zlsx_status_v1`; every refusal
/// here — malformed ref, mismatched anchor, overlap, member formula,
/// unknown dialect — is a statement about the call, so -1, never -2.
export fn zlsx_sheet_writer_write_row_with_formulas_v2(
    sw: ?*SheetWriter,
    cells_ptr: ?[*]const CCell,
    formulas_ptr: ?[*]const CFormulaCell,
    cells_len: usize,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) i32 {
    const sw_state: *SheetWriterState = @ptrCast(@alignCast(sw orelse {
        writeError(err_buf, err_buf_len, "InvalidInput");
        return ZLSX_ERROR;
    }));

    // Same scratch / heap pattern as the legacy formulas shim.
    var scratch_cells: [128]xlsx.Cell = undefined;
    var heap_cells: ?[]xlsx.Cell = null;
    defer if (heap_cells) |h| gpa.free(h);
    var cells_slice: []xlsx.Cell = &.{};

    var scratch_formulas: [128]?writer_mod.FormulaCell = undefined;
    var heap_formulas: ?[]?writer_mod.FormulaCell = null;
    defer if (heap_formulas) |h| gpa.free(h);
    var formulas_slice: []?writer_mod.FormulaCell = &.{};

    if (cells_len > 0) {
        const cp = cells_ptr orelse {
            writeError(err_buf, err_buf_len, "InvalidInput");
            return ZLSX_ERROR;
        };
        const fp = formulas_ptr orelse {
            writeError(err_buf, err_buf_len, "InvalidInput");
            return ZLSX_ERROR;
        };
        if (cells_len <= scratch_cells.len) {
            cells_slice = scratch_cells[0..cells_len];
            formulas_slice = scratch_formulas[0..cells_len];
        } else {
            heap_cells = gpa.alloc(xlsx.Cell, cells_len) catch {
                writeError(err_buf, err_buf_len, "OutOfMemory");
                return ZLSX_NOMEM;
            };
            cells_slice = heap_cells.?;
            heap_formulas = gpa.alloc(?writer_mod.FormulaCell, cells_len) catch {
                writeError(err_buf, err_buf_len, "OutOfMemory");
                return ZLSX_NOMEM;
            };
            formulas_slice = heap_formulas.?;
        }
        for (0..cells_len) |i| {
            cells_slice[i] = fromCCell(cp[i]) catch |e| {
                writeError(err_buf, err_buf_len, @errorName(e));
                return ZLSX_ERROR;
            };
            const d = fp[i];
            const text_ptr = d.text orelse {
                // Plain value slot: dialect and ref must be silent too
                // — a descriptor that says "no formula, but CSE" is a
                // contract violation, not a no-op.
                if (d.dialect != formula_dialect_scalar or d.ref != null or d.text_len != 0) {
                    writeError(err_buf, err_buf_len, "InvalidInput");
                    return ZLSX_ERROR;
                }
                formulas_slice[i] = null;
                continue;
            };
            if (d.text_len == 0) {
                // The empty formula the legacy shim's length-encoding
                // could silently drop; here it is an explicit reject
                // (matches the Zig layer's FormulaTextEmpty).
                writeError(err_buf, err_buf_len, "InvalidInput");
                return ZLSX_ERROR;
            }
            const dialect: writer_mod.FormulaCell.Dialect = switch (d.dialect) {
                formula_dialect_scalar, formula_dialect_dynamic_array => blk: {
                    if (d.ref != null or d.ref_len != 0) {
                        writeError(err_buf, err_buf_len, "InvalidInput");
                        return ZLSX_ERROR;
                    }
                    break :blk if (d.dialect == formula_dialect_scalar) .scalar else .dynamic_array;
                },
                formula_dialect_cse => blk: {
                    const rp = d.ref orelse {
                        writeError(err_buf, err_buf_len, "InvalidInput");
                        return ZLSX_ERROR;
                    };
                    if (d.ref_len == 0) {
                        writeError(err_buf, err_buf_len, "InvalidInput");
                        return ZLSX_ERROR;
                    }
                    break :blk .{ .cse = rp[0..d.ref_len] };
                },
                else => {
                    writeError(err_buf, err_buf_len, "InvalidInput");
                    return ZLSX_ERROR;
                },
            };
            formulas_slice[i] = .{ .text = text_ptr[0..d.text_len], .dialect = dialect };
        }
    }

    sw_state.inner.writeRowWithFormulaCells(cells_slice, formulas_slice) catch |e| {
        return failMapped(e, null, err_buf, err_buf_len);
    };
    return ZLSX_OK;
}

// ─── S3a: structural edits + the pivots read (zlsx_status_v1) ────────
//
// The `Editor` structural edits — row / column insert and delete, sheet
// add / rename / delete, table-column rename — and the S6 `pivots`
// NDJSON shape, under the same `zlsx_status_v1` contract as the M9a1 /
// M9a2 blocks above. A refusal is a statement about the workbook: a
// construct the rewriter will not shift (`RowEditUnsafeForSheet`), a
// name the workbook already has (`DuplicateSheetName`,
// `TableColumnNameInUse`), the last sheet, a part the archive or the
// parsers cannot read, a pivot graph that cannot be read whole. Those
// are -2 with `zlsx_diag_v1.error_name` and
// `plane = ZLSX_PLANE_NONE`. A statement about the call — an index off
// the grid, a sheet, table or column the workbook does not have (a
// selector), a name Excel would not take, an edit on a sheet whose
// staged writes have not been saved — is -1 with the error name in
// `errbuf`. Contract: docs/plans/c-abi-status-v1.md §10.

/// The refusal vocabulary of the structural edits and the pivots read
/// (§10): every error an edit raises that is a statement about the
/// WORKBOOK — a construct the rewriter will not shift, a part it cannot
/// read, a grid it would push past its edge, a name the workbook holds.
/// `statusOf` checks it after the fourteen planes; a name here crosses
/// as -2, every other error the editor raises (an index, a name, a
/// selector that names nothing — `TableNotFound` is the table-shaped
/// `SheetIndexOutOfRange` (Codex #207 r3 REL-304) — a sequencing rule:
/// statements about the call) as -1. The editor folds
/// most of its pre-flights into the two `*UnsafeForSheet` names; the
/// rest reach the boundary as the transform or a later sweep spells
/// them (Codex #207 r1 REL-102), and a caller sees the precise cause.
const structural_refusals = [_]anyerror{
    // The editor's own verdicts.
    error.RowEditUnsafeForSheet,
    error.ColEditUnsafeForSheet,
    error.CannotDeleteLastSheet,
    error.DuplicateSheetName,
    error.TableColumnNameInUse,
    error.MalformedPivotXml,
    // The worksheet transform's, raised by its pre-mutation probe.
    error.RowEditExceedsMaxRow,
    error.ColEditExceedsMaxCol,
    error.SplitPaneNotSupported,
    error.MalformedPaneSplit,
    error.MalformedSheetXml,
    // The sweeps' — a carrier the walkers cannot read or move.
    error.MalformedDrawingXml,
    error.DrawingCoordinateOverflow,
    error.MalformedVmlDrawing,
    error.VmlCoordinateOverflow,
    error.MalformedCommentsXml,
    error.MalformedTableXml,
    error.TableCoordinateOverflow,
    error.TableCollapseUnsafe,
    error.TableHeaderRowDeleteUnsafe,
    // A delete that collapses EVERY area of a DV/CF `sqref` — Excel
    // deletes the rule outright; zlsx cannot excise it mid-walk and
    // refuses rather than silently retarget it (Codex #216 r4
    // S3B-REL-805, the TableCollapseUnsafe shape).
    error.SqrefCollapseUnsafe,
    error.PivotEditUnsafe,
    error.MalformedExtensionXml,
    // The chart `<c:f>` sweep's own verdict: a chart part the walk
    // cannot read whole, refused before the first mutation like
    // `MalformedExtensionXml` (the S2 shape) and folded into the two
    // `*UnsafeForSheet` names by the editor's row / column pre-flights.
    error.MalformedChartXml,
    error.MalformedSheetRels,
    error.MalformedWorkbookRels,
    error.MalformedDrawingRels,
    error.MissingSheetPart,
    error.NoSheetData,
    // The anchors read's own verdict (S3b slice 7): an anchored
    // object on a worksheet part `xl/workbook.xml` does not list — no
    // record could carry a truthful `sheet` / `sheet_idx`, and
    // dropping it would be a partial inventory.
    error.DrawingOnUnlistedSheet,
    // The workbook's own structure, found broken on the way — a name
    // the workbook holds that no rename argument can fix, a part or a
    // relationship the archive lost (Codex #207 r3 REL-303).
    error.InternalSheetNameTooLong,
    error.MissingRelationship,
    error.SheetElementNotFound,
    error.RelationshipElementNotFound,
    error.SheetCountMismatch,
    error.MalformedWorkbookXml,
    error.IdSpaceExhausted,
    error.MissingWorkbookPart,
    error.MissingWorkbookRels,
    error.MissingContentTypes,
    error.MalformedContentTypes,
    error.ContentTypesOverrideNotFound,
    // The workbook layer's spellings of two editor verdicts, should a
    // path ever surface them unfolded.
    error.LastSheetUndeletable,
    error.SheetNameInUse,
    // S3c slice 1: the embedding write's one limit — a part past the
    // 512 MiB read cap, refused rather than writing a workbook zlsx
    // could not reopen (§2's rule: limits are Plane-2 refusals at
    // every layer, the `FormulaLimitExceeded` precedent).
    error.EmbeddingExceedsArchiveLimit,
    // S3c slice 2: the embeddable-rows read's verdicts on a cell value
    // it cannot carry — a boolean `<v>` that is not 0 / 1, a
    // shared-string index that is not a number or past the table, an
    // entity the decoder does not know, text that is not UTF-8 or
    // that NFC cannot normalize. Statements about the workbook's
    // content (the `MalformedSheetXml` shape), refused whole rather
    // than a record that lies. Raised by no other status_v1 export.
    error.UnsupportedCellValue,
    error.SstIndexOutOfRange,
    error.InvalidUtf8,
    error.UnicodeNormalizationFailed,
    // Round 1: a shared-string table the parser cannot read — the
    // `MalformedSheetXml` shape for the other part every string cell
    // depends on (S3C2-REL-102).
    error.MalformedSharedStringsXml,
    // S3c slice 3: the redaction sweep's verdicts on the set it walks
    // — the index read's refusal folded under one name (a coverage
    // range the parser cannot read, a binary part with the wrong
    // magic, a count that disagrees with its header), and a part the
    // index names that the archive lacks. Statements about the
    // workbook; the parser's own names (`InvalidRange`, `InvalidDtype`,
    // `CountMismatch`, …) stay -1, the write's statement about its
    // inputs. Raised by no other status_v1 export.
    error.MalformedEmbeddingSet,
    error.MissingEmbeddingPart,
    // NOT here, deliberately: `error.ZipBombSuspected`. The S1 caps
    // admit every entry on the open-time directory walk, so the
    // verdict fires where the ABI has no diag to carry it — the
    // pointer-returning path open, and `zlsx_open_buffer`, whose
    // shipped contract is -1 (Codex #216 r2 S3B-ERR-702 ruled r1's
    // -2 remap an ABI break; a typed open refusal needs a
    // status-bearing open ABI, deferred). It stays a generic -1 with
    // its name in errbuf on every path, the way the CLI keeps its
    // exit-4 mapping as a defensive posture for a future path that
    // decompresses without the walk.
};

fn isStructuralRefusal(e: anyerror) bool {
    for (structural_refusals) |r| {
        if (e == r) return true;
    }
    return false;
}

/// The handle check every S3a export shares. NULL is a statement about
/// the call (-1 `InvalidInput`), never a crash.
fn editorStateOrNull(ed: ?*Editor, err_buf: ?[*]u8, err_buf_len: usize) ?*EditorState {
    const e = ed orelse {
        writeError(err_buf, err_buf_len, "InvalidInput");
        return null;
    };
    return @ptrCast(@alignCast(e));
}

/// A `(ptr, len)` byte string at the boundary: NULL with a non-zero
/// length is `InvalidInput`; NULL with length 0 is the empty string
/// (which the editor then judges on its own terms — an empty sheet
/// name is `InvalidSheetName`, not a boundary violation).
fn bytesArg(ptr: ?[*]const u8, len: usize, err_buf: ?[*]u8, err_buf_len: usize) ?[]const u8 {
    if (ptr) |p| return p[0..len];
    if (len == 0) return "";
    writeError(err_buf, err_buf_len, "InvalidInput");
    return null;
}

/// Columns cross the boundary 0-based, as `zlsx_editor_set_cell` and
/// the census entries spell them; the editor's structural API is
/// 1-based (A = 1). `UINT32_MAX` has no 1-based spelling and is refused
/// here rather than wrapped to 0.
fn colOneBased(col: u32, err_buf: ?[*]u8, err_buf_len: usize) ?u32 {
    if (col == std.math.maxInt(u32)) {
        writeError(err_buf, err_buf_len, "ColumnIndexOutOfRange");
        return null;
    }
    return col + 1;
}

/// Insert a blank row before `before_row` (1-based) on `sheet_idx`;
/// every row at or below it shifts down by one, and every carrier the
/// rewriters know — formulas, defined names, hyperlinks, DV / CF,
/// merges, panes, autoFilter, tables, drawings, comments, `<xm:f>`,
/// chart `<c:f>` series formulas, a hosted pivot's `location@ref`, a
/// pivot cache's source range — moves in step. Staged in memory;
/// `zlsx_editor_save` commits.
export fn zlsx_editor_insert_row(
    ed: ?*Editor,
    sheet_idx: u32,
    before_row: u32,
    diag: ?*CDiag,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) i32 {
    if (!prepDiag(diag, err_buf, err_buf_len)) return ZLSX_ERROR;
    const state = editorStateOrNull(ed, err_buf, err_buf_len) orelse return ZLSX_ERROR;
    state.inner.insertRow(sheet_idx, before_row) catch |e| {
        return failMapped(e, diag, err_buf, err_buf_len);
    };
    return ZLSX_OK;
}

/// Delete row `row` (1-based) on `sheet_idx`; every row below shifts up
/// by one. Same rewrite coverage and refusal contract as the insert.
export fn zlsx_editor_delete_row(
    ed: ?*Editor,
    sheet_idx: u32,
    row: u32,
    diag: ?*CDiag,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) i32 {
    if (!prepDiag(diag, err_buf, err_buf_len)) return ZLSX_ERROR;
    const state = editorStateOrNull(ed, err_buf, err_buf_len) orelse return ZLSX_ERROR;
    state.inner.deleteRow(sheet_idx, row) catch |e| {
        return failMapped(e, diag, err_buf, err_buf_len);
    };
    return ZLSX_OK;
}

/// Insert a blank column before `before_col` (0-based, A = 0) on
/// `sheet_idx`. Same rewrite coverage and refusal contract as the row
/// edits.
export fn zlsx_editor_insert_column(
    ed: ?*Editor,
    sheet_idx: u32,
    before_col: u32,
    diag: ?*CDiag,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) i32 {
    if (!prepDiag(diag, err_buf, err_buf_len)) return ZLSX_ERROR;
    const state = editorStateOrNull(ed, err_buf, err_buf_len) orelse return ZLSX_ERROR;
    const col_1 = colOneBased(before_col, err_buf, err_buf_len) orelse return ZLSX_ERROR;
    state.inner.insertColumn(sheet_idx, col_1) catch |e| {
        return failMapped(e, diag, err_buf, err_buf_len);
    };
    return ZLSX_OK;
}

/// Delete column `col` (0-based, A = 0) on `sheet_idx`.
export fn zlsx_editor_delete_column(
    ed: ?*Editor,
    sheet_idx: u32,
    col: u32,
    diag: ?*CDiag,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) i32 {
    if (!prepDiag(diag, err_buf, err_buf_len)) return ZLSX_ERROR;
    const state = editorStateOrNull(ed, err_buf, err_buf_len) orelse return ZLSX_ERROR;
    const col_1 = colOneBased(col, err_buf, err_buf_len) orelse return ZLSX_ERROR;
    state.inner.deleteColumn(sheet_idx, col_1) catch |e| {
        return failMapped(e, diag, err_buf, err_buf_len);
    };
    return ZLSX_OK;
}

/// `zlsx_editor_add_sheet`'s `*out_sheet_idx` before a sheet exists —
/// no index, never mistaken for sheet 0.
const no_sheet_idx: u32 = 0xFFFF_FFFF;

/// Append an empty sheet named `name` (`name_len` UTF-8 bytes, not
/// NUL-terminated). `*out_sheet_idx` (nullable) receives the new
/// sheet's index on ZLSX_OK and `UINT32_MAX` otherwise. The name is
/// judged by the fresh writer's rules (31 scalars, no `:\/?*[]`, not
/// `History`) and against every existing name ASCII case-insensitively.
export fn zlsx_editor_add_sheet(
    ed: ?*Editor,
    name_ptr: ?[*]const u8,
    name_len: usize,
    out_sheet_idx: ?*u32,
    diag: ?*CDiag,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) i32 {
    if (out_sheet_idx) |o| o.* = no_sheet_idx;
    if (!prepDiag(diag, err_buf, err_buf_len)) return ZLSX_ERROR;
    const state = editorStateOrNull(ed, err_buf, err_buf_len) orelse return ZLSX_ERROR;
    const name = bytesArg(name_ptr, name_len, err_buf, err_buf_len) orelse return ZLSX_ERROR;
    const idx = state.inner.addSheet(name) catch |e| {
        return failMapped(e, diag, err_buf, err_buf_len);
    };
    if (out_sheet_idx) |o| o.* = idx;
    return ZLSX_OK;
}

/// Rename sheet `sheet_idx` to `name`. Cross-sheet references
/// (`'Old'!A1`, defined names, hyperlink locations, DV / CF, `<xm:f>`,
/// chart `<c:f>` series formulas) follow the rename; a pivot cache's
/// `worksheetSource@sheet` does NOT
/// (the Zig editor's hole, stated in the header) and reads back as
/// `"resolved":null`.
export fn zlsx_editor_rename_sheet(
    ed: ?*Editor,
    sheet_idx: u32,
    name_ptr: ?[*]const u8,
    name_len: usize,
    diag: ?*CDiag,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) i32 {
    if (!prepDiag(diag, err_buf, err_buf_len)) return ZLSX_ERROR;
    const state = editorStateOrNull(ed, err_buf, err_buf_len) orelse return ZLSX_ERROR;
    const name = bytesArg(name_ptr, name_len, err_buf, err_buf_len) orelse return ZLSX_ERROR;
    state.inner.renameSheet(sheet_idx, name) catch |e| {
        return failMapped(e, diag, err_buf, err_buf_len);
    };
    return ZLSX_OK;
}

/// Delete sheet `sheet_idx`. Refuses the last sheet
/// (`CannotDeleteLastSheet`, -2); references into the deleted sheet
/// collapse to `#REF!`. Every index above `sheet_idx` shifts down by
/// one. Requires a clean editor — no staged cell writes or appended
/// rows on any sheet (`SheetDeleteRequiresCleanState`, -1: save first).
export fn zlsx_editor_delete_sheet(
    ed: ?*Editor,
    sheet_idx: u32,
    diag: ?*CDiag,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) i32 {
    if (!prepDiag(diag, err_buf, err_buf_len)) return ZLSX_ERROR;
    const state = editorStateOrNull(ed, err_buf, err_buf_len) orelse return ZLSX_ERROR;
    state.inner.deleteSheet(sheet_idx) catch |e| {
        return failMapped(e, diag, err_buf, err_buf_len);
    };
    return ZLSX_OK;
}

/// Rename column `old_name` of table `table_name` to `new_name`: the
/// `<tableColumn>`, the table's own formulas, every structured
/// reference workbook-wide, defined names, hyperlink locations, DV /
/// CF, and the header cell's text. Names are plain (decoded) text. A
/// table or column the workbook does not have is a selector, like a
/// sheet index — -1 `TableNotFound` / `TableColumnNotFound`; a name
/// another column holds is the workbook's — -2 `TableColumnNameInUse`.
export fn zlsx_editor_rename_table_column(
    ed: ?*Editor,
    table_ptr: ?[*]const u8,
    table_len: usize,
    old_ptr: ?[*]const u8,
    old_len: usize,
    new_ptr: ?[*]const u8,
    new_len: usize,
    diag: ?*CDiag,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) i32 {
    if (!prepDiag(diag, err_buf, err_buf_len)) return ZLSX_ERROR;
    const state = editorStateOrNull(ed, err_buf, err_buf_len) orelse return ZLSX_ERROR;
    const table = bytesArg(table_ptr, table_len, err_buf, err_buf_len) orelse return ZLSX_ERROR;
    const old = bytesArg(old_ptr, old_len, err_buf, err_buf_len) orelse return ZLSX_ERROR;
    const new = bytesArg(new_ptr, new_len, err_buf, err_buf_len) orelse return ZLSX_ERROR;
    state.inner.renameTableColumn(table, old, new) catch |e| {
        return failMapped(e, diag, err_buf, err_buf_len);
    };
    return ZLSX_OK;
}

/// The S6 `pivots` records — one `{"kind":"pivot",…}` line per pivot
/// table in host-sheet order, then one `{"kind":"pivot_cache",…}` line
/// per cache no table reads — as a library-allocated UTF-8 buffer,
/// byte-for-byte what `zlsx pivots <file>` prints (`docs/cli.md`,
/// "pivots"). Read over the editor's current workbook state:
/// structural edits are visible immediately; staged `set_cell` /
/// `append_row` writes reach the pivot graph at save (a cache whose
/// source they change is rebuilt or marked then, S7b) — save, then
/// read, to see them (Codex #207 r7 REL-701). A workbook without
/// pivots is ZLSX_OK with `(NULL, 0)`.
/// A graph that cannot be read whole refuses (`MalformedPivotXml`,
/// -2). Release with `zlsx_buffer_release`.
export fn zlsx_editor_pivots_ndjson(
    ed: ?*Editor,
    out_ptr: ?*?[*]u8,
    out_len: ?*usize,
    diag: ?*CDiag,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) i32 {
    // Every output prepped before anything can fail: a rejected sibling
    // leaves the accepted ones releasable.
    if (out_ptr) |op| op.* = null;
    if (out_len) |ol| ol.* = 0;
    if (!prepDiag(diag, err_buf, err_buf_len)) return ZLSX_ERROR;
    const op = out_ptr orelse {
        writeError(err_buf, err_buf_len, "NullOutPointer");
        return ZLSX_ERROR;
    };
    const ol = out_len orelse {
        writeError(err_buf, err_buf_len, "NullOutPointer");
        return ZLSX_ERROR;
    };
    const state = editorStateOrNull(ed, err_buf, err_buf_len) orelse return ZLSX_ERROR;

    const bytes = pivotsNdjsonOwned(gpa, &state.inner.workbook) catch |e| {
        return failMapped(e, diag, err_buf, err_buf_len);
    };
    if (bytes.len == 0) {
        gpa.free(bytes);
        return ZLSX_OK;
    }
    op.* = bytes.ptr;
    ol.* = bytes.len;
    return ZLSX_OK;
}

/// The pivots records as one owned buffer in `alloc`. The allocating
/// writer reports a failed growth as `WriteFailed`; at this boundary
/// that is an allocation failure and crosses as `-3`, not as a generic
/// error (Codex #207 r1 REL-103). Every other error is the graph
/// read's.
fn pivotsNdjsonOwned(alloc: std.mem.Allocator, wb: *zlsx_pkg.Workbook) ![]u8 {
    // A part the store cannot materialise (a bad CRC, a broken deflate
    // stream, a name the archive lost) is the graph that cannot be read
    // whole — `MalformedPivotXml`, like every other reason the read
    // refuses (Codex #207 r2 REL-203). Memory and the archive-wide
    // decompression budget keep their own statuses.
    var pv = wb.pivotTables() catch |e| switch (e) {
        error.OutOfMemory, error.ZipBombSuspected => return e,
        else => return error.MalformedPivotXml,
    };
    defer pv.deinit();
    var out: std.Io.Writer.Allocating = .init(alloc);
    defer out.deinit();
    zlsx_pkg.pivots.ndjson.writeAll(&out.writer, &pv) catch |e| switch (e) {
        error.WriteFailed => return error.OutOfMemory,
    };
    return out.toOwnedSlice();
}

/// The S3b `defined-names` records — one `{"kind":"defined_name",…}`
/// line per `<definedName>` of `xl/workbook.xml`, in document order —
/// as a library-allocated UTF-8 buffer, byte-for-byte what
/// `zlsx defined-names <file>` prints with no selector (`docs/cli.md`,
/// "defined-names"). `body` is the formula text as authored — nothing
/// resolved or rewritten. Read over the editor's current workbook
/// state: structural edits and the name sweeps they carry (a sheet
/// rename rewriting the bodies) are visible immediately; nothing about
/// a defined name waits for save. A workbook without defined names is
/// ZLSX_OK with `(NULL, 0)`. An inventory that cannot be served
/// faithfully — a carrier that does not decode, malformed UTF-8, a
/// body with embedded markup — refuses whole (`MalformedWorkbookXml`,
/// -2) rather than hand over a record that lies. Release with
/// `zlsx_buffer_release`.
export fn zlsx_editor_defined_names_ndjson(
    ed: ?*Editor,
    out_ptr: ?*?[*]u8,
    out_len: ?*usize,
    diag: ?*CDiag,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) i32 {
    // Every output prepped before anything can fail: a rejected sibling
    // leaves the accepted ones releasable.
    if (out_ptr) |op| op.* = null;
    if (out_len) |ol| ol.* = 0;
    if (!prepDiag(diag, err_buf, err_buf_len)) return ZLSX_ERROR;
    const op = out_ptr orelse {
        writeError(err_buf, err_buf_len, "NullOutPointer");
        return ZLSX_ERROR;
    };
    const ol = out_len orelse {
        writeError(err_buf, err_buf_len, "NullOutPointer");
        return ZLSX_ERROR;
    };
    const state = editorStateOrNull(ed, err_buf, err_buf_len) orelse return ZLSX_ERROR;

    const bytes = definedNamesNdjsonOwned(gpa, &state.inner.workbook) catch |e| {
        return failMapped(e, diag, err_buf, err_buf_len);
    };
    if (bytes.len == 0) {
        gpa.free(bytes);
        return ZLSX_OK;
    }
    op.* = bytes.ptr;
    ol.* = bytes.len;
    return ZLSX_OK;
}

/// The defined-name records as one owned buffer in `alloc` — the
/// shared writer (`pkg/defined_name_ndjson.zig`) over the workbook's
/// parsed `xl/workbook.xml` view, so the bytes are the CLI's. The
/// allocating writer reports a failed growth as `WriteFailed`; at this
/// boundary that is an allocation failure and crosses as `-3`, the
/// pivots builder's rule.
fn definedNamesNdjsonOwned(alloc: std.mem.Allocator, wb: *zlsx_pkg.Workbook) ![]u8 {
    var view = try zlsx_pkg.defined_names_ndjson.collect(alloc, &wb.workbook);
    defer view.deinit();
    var out: std.Io.Writer.Allocating = .init(alloc);
    defer out.deinit();
    zlsx_pkg.defined_names_ndjson.writeAll(&out.writer, &view) catch |e| switch (e) {
        error.WriteFailed => return error.OutOfMemory,
    };
    return out.toOwnedSlice();
}

/// The S3b `conditional-formats` records — one
/// `{"kind":"conditional_format",…}` line per `<cfRule>`, sheets in
/// workbook order, rules in sheet-document order — as a
/// library-allocated UTF-8 buffer, byte-for-byte what
/// `zlsx conditional-formats <file>` prints with no selector
/// (docs/cli.md, "conditional-formats"). The record is the rule
/// envelope (`sqref`, `rule_type`, `formulas`, `dxf_id`, `priority`),
/// not the visual payload — `<colorScale>` / `<dataBar>` / `<iconSet>`
/// bodies and the `<dxfs>` styles stay in their parts. Read over the
/// editor's current parts: structural edits and the DV/CF sweeps they
/// carry are visible immediately; staged cell writes never touch the
/// rule machinery. A workbook without conditional formatting is
/// ZLSX_OK with `(NULL, 0)`. An inventory that cannot be served
/// faithfully refuses whole — a sheet list the strict workbook read
/// cannot prove (`MalformedWorkbookXml`) or a sheet part the strict
/// walk cannot (`MalformedSheetXml`), both -2 — rather than hand over
/// a record that lies. Release with `zlsx_buffer_release`.
export fn zlsx_editor_conditional_formats_ndjson(
    ed: ?*Editor,
    out_ptr: ?*?[*]u8,
    out_len: ?*usize,
    diag: ?*CDiag,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) i32 {
    // Every output prepped before anything can fail: a rejected sibling
    // leaves the accepted ones releasable.
    if (out_ptr) |op| op.* = null;
    if (out_len) |ol| ol.* = 0;
    if (!prepDiag(diag, err_buf, err_buf_len)) return ZLSX_ERROR;
    const op = out_ptr orelse {
        writeError(err_buf, err_buf_len, "NullOutPointer");
        return ZLSX_ERROR;
    };
    const ol = out_len orelse {
        writeError(err_buf, err_buf_len, "NullOutPointer");
        return ZLSX_ERROR;
    };
    const state = editorStateOrNull(ed, err_buf, err_buf_len) orelse return ZLSX_ERROR;

    // Exhaustive over the closed `CollectError` — no `else`, so a
    // future member breaks this compile and forces a status decision
    // instead of silently crossing as a generic -1 (Codex #216 r8
    // S3B-MNT-911).
    const bytes = conditionalFormatsNdjsonOwned(gpa, &state.inner.workbook) catch |e| switch (e) {
        error.MalformedWorkbookXml,
        error.MalformedSheetXml,
        error.MissingSheetPart,
        error.ZipBombSuspected,
        error.OutOfMemory,
        => return failMapped(e, diag, err_buf, err_buf_len),
    };
    if (bytes.len == 0) {
        gpa.free(bytes);
        return ZLSX_OK;
    }
    op.* = bytes.ptr;
    ol.* = bytes.len;
    return ZLSX_OK;
}

/// The conditional-format records as one owned buffer in `alloc` — the
/// shared writer (`pkg/conditional_format_ndjson.zig`) over the
/// editor's current parts, so the bytes are the CLI's. The allocating
/// writer reports a failed growth as `WriteFailed`; at this boundary
/// that is an allocation failure and crosses as `-3`, the pivots
/// builder's rule.
fn conditionalFormatsNdjsonOwned(alloc: std.mem.Allocator, wb: *zlsx_pkg.Workbook) zlsx_pkg.conditional_formats_ndjson.CollectError![]u8 {
    var view = try zlsx_pkg.conditional_formats_ndjson.collect(alloc, wb);
    defer view.deinit();
    var out: std.Io.Writer.Allocating = .init(alloc);
    defer out.deinit();
    zlsx_pkg.conditional_formats_ndjson.writeAll(&out.writer, &view) catch |e| switch (e) {
        error.WriteFailed => return error.OutOfMemory,
    };
    return out.toOwnedSlice();
}

/// The S3b `anchors` records — one `{"kind":"image_anchor",…}` line
/// per anchored image and one `{"kind":"chart_anchor",…}` line per
/// anchored chart, sheets in workbook order, a sheet's images before
/// its charts, each class in drawing-document order — as a
/// library-allocated UTF-8 buffer, byte-for-byte what
/// `zlsx anchors <file>` prints with no selector (docs/cli.md,
/// "anchors"). The record is the anchor geometry and where the
/// payload lives (`part`, an image's `bytes` count, a chart's
/// `chart_type` + entity-decoded `series_refs`), never the payload:
/// image bytes and chart XML stay in their parts. Read over the
/// editor's current parts: structural edits and the drawing sweeps
/// they carry (a row insert moving an anchor's `from` / `to`, a
/// rename renaming `sheet`) are visible immediately; staged cell
/// writes never touch a drawing. A workbook without anchored objects
/// is ZLSX_OK with `(NULL, 0)`. An inventory that cannot be served
/// faithfully refuses whole — a sheet list the strict workbook read
/// cannot prove (`MalformedWorkbookXml`), a drawing graph the strict
/// walk cannot read whole (`MalformedDrawingXml`), or an anchor on a
/// worksheet part the workbook does not list
/// (`DrawingOnUnlistedSheet`), all -2 — rather than hand over a
/// record that lies or a list with a hole. Release with
/// `zlsx_buffer_release`.
export fn zlsx_editor_anchors_ndjson(
    ed: ?*Editor,
    out_ptr: ?*?[*]u8,
    out_len: ?*usize,
    diag: ?*CDiag,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) i32 {
    // Every output prepped before anything can fail: a rejected sibling
    // leaves the accepted ones releasable.
    if (out_ptr) |op| op.* = null;
    if (out_len) |ol| ol.* = 0;
    if (!prepDiag(diag, err_buf, err_buf_len)) return ZLSX_ERROR;
    const op = out_ptr orelse {
        writeError(err_buf, err_buf_len, "NullOutPointer");
        return ZLSX_ERROR;
    };
    const ol = out_len orelse {
        writeError(err_buf, err_buf_len, "NullOutPointer");
        return ZLSX_ERROR;
    };
    const state = editorStateOrNull(ed, err_buf, err_buf_len) orelse return ZLSX_ERROR;

    // Exhaustive over the closed `CollectError` — no `else`, so a
    // future member breaks this compile and forces a status decision
    // (the conditional-formats read's rule, Codex #216 r8
    // S3B-MNT-911).
    const bytes = anchorsNdjsonOwned(gpa, &state.inner.workbook) catch |e| switch (e) {
        error.MalformedWorkbookXml,
        error.DrawingOnUnlistedSheet,
        error.MalformedDrawingXml,
        error.ZipBombSuspected,
        error.OutOfMemory,
        => return failMapped(e, diag, err_buf, err_buf_len),
    };
    if (bytes.len == 0) {
        gpa.free(bytes);
        return ZLSX_OK;
    }
    op.* = bytes.ptr;
    ol.* = bytes.len;
    return ZLSX_OK;
}

/// The anchor records as one owned buffer in `alloc` — the shared
/// writer (`pkg/anchor_ndjson.zig`) over the editor's current parts,
/// so the bytes are the CLI's. The allocating writer reports a failed
/// growth as `WriteFailed`; at this boundary that is an allocation
/// failure and crosses as `-3`, the pivots builder's rule.
fn anchorsNdjsonOwned(alloc: std.mem.Allocator, wb: *zlsx_pkg.Workbook) zlsx_pkg.anchors_ndjson.CollectError![]u8 {
    var view = try zlsx_pkg.anchors_ndjson.collect(alloc, wb);
    defer view.deinit();
    var out: std.Io.Writer.Allocating = .init(alloc);
    defer out.deinit();
    zlsx_pkg.anchors_ndjson.writeAll(&out.writer, &view) catch |e| switch (e) {
        error.WriteFailed => return error.OutOfMemory,
    };
    return out.toOwnedSlice();
}

/// The S3b `sheet-props` records — one `{"kind":"sheet_props",…}`
/// line per workbook sheet, workbook order — as a library-allocated
/// UTF-8 buffer, byte-for-byte what `zlsx sheet-props <file>` prints
/// with no selector (docs/cli.md, "sheet-props"). Each record is the
/// sheet's `<dimension ref>` as authored (null when the element or
/// the attribute is absent) and the `<pane>` of its FIRST
/// `<sheetView>` as authored (null when there is none): `x_split` /
/// `y_split` / `top_left_cell` / `active_pane` / `state`, each null
/// when the source omits it, no schema default applied, split panes
/// reported as written (the lenient `Worksheet.freezePane` narrows to
/// frozen panes; this read does not). Read over the editor's current
/// parts: structural edits and the sheet sweeps they carry (a rename
/// renaming `sheet`, a row insert growing `dimension` and moving a
/// frozen pane's split and `top_left_cell`) are visible immediately;
/// staged cell writes never touch the extent or the views. An
/// inventory that cannot be served faithfully refuses whole — a sheet
/// list the strict workbook read cannot prove (`MalformedWorkbookXml`)
/// or a sheet part the strict walk cannot prove a pane / extent for
/// (`MalformedSheetXml`: a second `<dimension>` / `<sheetViews>` /
/// first-view `<pane>`, a duplicate attribute on that machinery, an
/// MCE construct at a recognized slot, a carrier that does not
/// decode), both -2 — rather than hand over a record that lies or a
/// list with a hole. Release with `zlsx_buffer_release`.
export fn zlsx_editor_sheet_props_ndjson(
    ed: ?*Editor,
    out_ptr: ?*?[*]u8,
    out_len: ?*usize,
    diag: ?*CDiag,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) i32 {
    // Every output prepped before anything can fail: a rejected sibling
    // leaves the accepted ones releasable.
    if (out_ptr) |op| op.* = null;
    if (out_len) |ol| ol.* = 0;
    if (!prepDiag(diag, err_buf, err_buf_len)) return ZLSX_ERROR;
    const op = out_ptr orelse {
        writeError(err_buf, err_buf_len, "NullOutPointer");
        return ZLSX_ERROR;
    };
    const ol = out_len orelse {
        writeError(err_buf, err_buf_len, "NullOutPointer");
        return ZLSX_ERROR;
    };
    const state = editorStateOrNull(ed, err_buf, err_buf_len) orelse return ZLSX_ERROR;

    // Exhaustive over the closed `CollectError` — no `else`, so a
    // future member breaks this compile and forces a status decision
    // (the conditional-formats read's rule, Codex #216 r8
    // S3B-MNT-911).
    const bytes = sheetPropsNdjsonOwned(gpa, &state.inner.workbook) catch |e| switch (e) {
        error.MalformedWorkbookXml,
        error.MalformedSheetXml,
        error.MissingSheetPart,
        error.ZipBombSuspected,
        error.OutOfMemory,
        => return failMapped(e, diag, err_buf, err_buf_len),
    };
    // The strict inventory refuses a sheetless workbook (a missing or
    // empty `<sheets>`, CT_Sheets minOccurs=1), so the stream is never
    // empty on success; the `(NULL, 0)` arm keeps the buffer contract
    // uniform with its siblings rather than assert that property of
    // another module here.
    if (bytes.len == 0) {
        gpa.free(bytes);
        return ZLSX_OK;
    }
    op.* = bytes.ptr;
    ol.* = bytes.len;
    return ZLSX_OK;
}

/// The sheet-props records as one owned buffer in `alloc` — the
/// shared writer (`pkg/sheet_props_ndjson.zig`) over the editor's
/// current parts, so the bytes are the CLI's. The allocating writer
/// reports a failed growth as `WriteFailed`; at this boundary that is
/// an allocation failure and crosses as `-3`, the pivots builder's
/// rule.
fn sheetPropsNdjsonOwned(alloc: std.mem.Allocator, wb: *zlsx_pkg.Workbook) zlsx_pkg.sheet_props_ndjson.CollectError![]u8 {
    var view = try zlsx_pkg.sheet_props_ndjson.collect(alloc, wb);
    defer view.deinit();
    var out: std.Io.Writer.Allocating = .init(alloc);
    defer out.deinit();
    zlsx_pkg.sheet_props_ndjson.writeAll(&out.writer, &view) catch |e| switch (e) {
        error.WriteFailed => return error.OutOfMemory,
    };
    return out.toOwnedSlice();
}

/// The S3b `calc-props` record — the ONE `{"kind":"calc_props",…}`
/// line of `xl/workbook.xml`'s `<calcPr>` — as a library-allocated
/// UTF-8 buffer, byte-for-byte what `zlsx calc-props <file>` prints
/// (docs/cli.md, "calc-props"): `calc_id` / `full_calc_on_load` /
/// `iterate` / `iterate_count` / `iterate_delta` as authored, every
/// field null when the element or the attribute is absent (a workbook
/// without `<calcPr>` is a record of nulls, never an empty buffer —
/// the doc-props convention). Read over the editor's current parts:
/// `zlsx_editor_mark_recalc_on_load` and a recalc that lands set
/// `fullCalcOnLoad="1"` in place, visible immediately; staged cell
/// writes never touch the element. A slot the read cannot report
/// faithfully refuses (`MalformedWorkbookXml`, -2): two `<calcPr>` at
/// the slot, one an MCE branch could project there, a duplicate
/// attribute, a carrier that does not decode. Release with
/// `zlsx_buffer_release`.
export fn zlsx_editor_calc_props_ndjson(
    ed: ?*Editor,
    out_ptr: ?*?[*]u8,
    out_len: ?*usize,
    diag: ?*CDiag,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) i32 {
    if (out_ptr) |op| op.* = null;
    if (out_len) |ol| ol.* = 0;
    if (!prepDiag(diag, err_buf, err_buf_len)) return ZLSX_ERROR;
    const op = out_ptr orelse {
        writeError(err_buf, err_buf_len, "NullOutPointer");
        return ZLSX_ERROR;
    };
    const ol = out_len orelse {
        writeError(err_buf, err_buf_len, "NullOutPointer");
        return ZLSX_ERROR;
    };
    const state = editorStateOrNull(ed, err_buf, err_buf_len) orelse return ZLSX_ERROR;

    // Exhaustive over the closed `CalcError` — no `else`.
    const bytes = calcPropsNdjsonOwned(gpa, &state.inner.workbook) catch |e| switch (e) {
        error.MalformedWorkbookXml,
        error.ZipBombSuspected,
        error.OutOfMemory,
        => return failMapped(e, diag, err_buf, err_buf_len),
    };
    // The record writer always emits its one line: an absent
    // `<calcPr>` is a record of nulls, so a caller never sees
    // `(NULL, 0)` on success.
    std.debug.assert(bytes.len != 0);
    op.* = bytes.ptr;
    ol.* = bytes.len;
    return ZLSX_OK;
}

/// The calc-props record as one owned buffer in `alloc` — the shared
/// writer over the editor's current `xl/workbook.xml`, so the bytes
/// are the CLI's. `WriteFailed` crosses as `-3`, as above.
fn calcPropsNdjsonOwned(alloc: std.mem.Allocator, wb: *zlsx_pkg.Workbook) zlsx_pkg.sheet_props_ndjson.CalcError![]u8 {
    const rec = try zlsx_pkg.sheet_props_ndjson.collectCalc(alloc, wb);
    var out: std.Io.Writer.Allocating = .init(alloc);
    defer out.deinit();
    zlsx_pkg.sheet_props_ndjson.writeCalcRecord(&out.writer, rec) catch |e| switch (e) {
        error.WriteFailed => return error.OutOfMemory,
    };
    return out.toOwnedSlice();
}

// ─── M9a1 tests ──────────────────────────────────────────────────────

/// A1 = 1 (plain), B1 = formula "A1+2" with cached <v>0</v> — one
/// stale cell for recalculate to rewrite, written through the C
/// surface so the test crosses the same boundary a binding does.
fn writeM9a1Fixture(io: std.Io, tt: *TestTmp) ![:0]u8 {
    const alloc = std.testing.allocator;
    const path = try tt.path(alloc, io, "m9a1.xlsx");
    errdefer alloc.free(path);
    var err_buf: [128]u8 = undefined;
    const w = zlsx_writer_create(&err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
    defer zlsx_writer_close(w);
    const name = "Sheet1";
    const sw = zlsx_writer_add_sheet(w, name.ptr, name.len, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
    const cells = [_]CCell{ toCCell(.{ .number = 1 }), toCCell(.{ .number = 0 }) };
    const formula = "A1+2";
    const fptrs = [_]?[*]const u8{ null, formula.ptr };
    const flens = [_]usize{ 0, formula.len };
    if (zlsx_sheet_writer_write_row_with_formulas(sw, &cells, &fptrs, &flens, cells.len, &err_buf, err_buf.len) != 0)
        return error.TestUnexpectedResult;
    if (zlsx_writer_save(w, path.ptr, path.len, &err_buf, err_buf.len) != 0)
        return error.TestUnexpectedResult;
    return path;
}

fn zeroRun() CRun {
    var crun = std.mem.zeroes(CRun);
    crun.struct_size = @sizeOf(CRun);
    return crun;
}

test "zlsx_status_v1: the fourteen plane values are pinned ABI" {
    // The C header hard-codes these numbers; the enum's declaration
    // order is what they pin. A reorder is an ABI break and fails here.
    const names = [_][]const u8{
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
    };
    const fields = @typeInfo(PlaneTwo).@"enum".fields;
    try std.testing.expectEqual(names.len, fields.len);
    inline for (fields, 0..) |f, i| {
        try std.testing.expectEqualStrings(names[i], f.name);
        try std.testing.expectEqual(i, f.value);
    }
}

test "zlsx_engine_fingerprint: names every identity component" {
    const fp = std.mem.span(zlsx_engine_fingerprint());
    try std.testing.expect(std.mem.indexOf(u8, fp, build_options.version) != null);
    try std.testing.expect(std.mem.indexOf(u8, fp, "excel_fp_rules_v1") != null);
    try std.testing.expect(std.mem.indexOf(u8, fp, "rng_v1") != null);
    try std.testing.expect(std.mem.indexOf(u8, fp, "collation_v1") != null);
    try std.testing.expect(std.mem.indexOf(u8, fp, @tagName(builtin.target.cpu.arch)) != null);
    // Two calls, one static string — no allocation to leak.
    try std.testing.expectEqual(zlsx_engine_fingerprint(), zlsx_engine_fingerprint());
}

test "cancel token: lifecycle, and a pre-triggered token cancels both entry points" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try writeM9a1Fixture(io, &tt);
    defer std.testing.allocator.free(path);

    var err_buf: [128]u8 = undefined;
    var tok: ?*CancelTok = null;
    try std.testing.expectEqual(ZLSX_OK, zlsx_cancel_token_new(&tok, &err_buf, err_buf.len));
    try std.testing.expect(tok != null);
    defer zlsx_cancel_token_free(tok);
    zlsx_cancel_token_trigger(tok);
    zlsx_cancel_token_trigger(tok); // re-trigger is a no-op

    const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
    defer zlsx_editor_close(ed);

    var crun = zeroRun();
    crun.cancel = tok;

    var report = std.mem.zeroes(CRecalcReport);
    report.struct_size = @sizeOf(CRecalcReport);
    try std.testing.expectEqual(ZLSX_CANCELLED, zlsx_editor_recalculate(ed, &crun, &report, null, &err_buf, err_buf.len));

    var val = std.mem.zeroes(CValue);
    val.struct_size = @sizeOf(CValue);
    const f = "=1+2";
    try std.testing.expectEqual(ZLSX_CANCELLED, zlsx_editor_evaluate(ed, f.ptr, f.len, 0, 0, 0, &crun, &val, null, null, &err_buf, err_buf.len));

    // NULL-safety of the token exports.
    zlsx_cancel_token_trigger(null);
    zlsx_cancel_token_free(null);
}

test "M9a1 end-to-end: recalculate rewrites the stale cell, evaluate reads, mark marks" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try writeM9a1Fixture(io, &tt);
    defer std.testing.allocator.free(path);

    var err_buf: [128]u8 = undefined;
    const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
    defer zlsx_editor_close(ed);

    var crun = zeroRun();
    crun.now_utc_ms = 1_700_000_000_000;
    crun.rng_seed = 42;

    var report = std.mem.zeroes(CRecalcReport);
    report.struct_size = @sizeOf(CRecalcReport);
    var diag = std.mem.zeroes(CDiag);
    diag.struct_size = @sizeOf(CDiag);
    try std.testing.expectEqual(ZLSX_OK, zlsx_editor_recalculate(ed, &crun, &report, &diag, &err_buf, err_buf.len));
    try std.testing.expect(report.cells_written >= 1);
    try std.testing.expectEqual(@as(u32, 1), report.resolved_present);
    try std.testing.expectEqual(@as(i64, 1_700_000_000_000), report.resolved.now_utc_ms);
    try std.testing.expectEqual(@as(u64, 42), report.resolved.rng_seed);
    // A recalc derives dialect per stored cell — the echo says so.
    try std.testing.expectEqual(dialect_none, report.resolved.dialect);
    // Defaults echo as numbers, never as 0.
    try std.testing.expect(report.resolved.max_run_arena_bytes > 0);
    try std.testing.expectEqual(@as(u32, 0), report.kept_stale);
    try std.testing.expectEqual(@as(u32, 0), report.durability_warning);
    try std.testing.expectEqual(plane_none, diag.plane);
    zlsx_recalc_report_release(&report);
    zlsx_diag_release(&diag);

    // Scalar number.
    var val = std.mem.zeroes(CValue);
    val.struct_size = @sizeOf(CValue);
    var resolved = std.mem.zeroes(CResolved);
    resolved.struct_size = @sizeOf(CResolved);
    const f_num = "=A1+2";
    try std.testing.expectEqual(ZLSX_OK, zlsx_editor_evaluate(ed, f_num.ptr, f_num.len, 0, 0, 0, &crun, &val, &resolved, null, &err_buf, err_buf.len));
    try std.testing.expectEqual(@as(u32, 0), val.is_matrix);
    try std.testing.expectEqual(@as(usize, 1), val.elems_len);
    try std.testing.expectEqual(value_tag_number, val.elems.?[0].tag);
    try std.testing.expectEqual(@as(f64, 3), val.elems.?[0].num);
    // Standalone eval states its dialect; the echo carries it.
    try std.testing.expectEqual(@as(u32, 0), resolved.dialect);
    zlsx_value_release(&val);

    // Text, through the payload arena.
    val = std.mem.zeroes(CValue);
    val.struct_size = @sizeOf(CValue);
    const f_text = "=\"a\"&\"b\"";
    try std.testing.expectEqual(ZLSX_OK, zlsx_editor_evaluate(ed, f_text.ptr, f_text.len, 0, 0, 0, &crun, &val, null, null, &err_buf, err_buf.len));
    try std.testing.expectEqual(value_tag_text, val.elems.?[0].tag);
    const t = val.elems.?[0];
    try std.testing.expectEqualStrings("ab", val.payload.?[t.payload_off .. t.payload_off + t.payload_len]);
    zlsx_value_release(&val);

    // An Excel error VALUE is a successful result (plane 1), not a status.
    val = std.mem.zeroes(CValue);
    val.struct_size = @sizeOf(CValue);
    const f_err = "=1/0";
    try std.testing.expectEqual(ZLSX_OK, zlsx_editor_evaluate(ed, f_err.ptr, f_err.len, 0, 0, 0, &crun, &val, null, null, &err_buf, err_buf.len));
    try std.testing.expectEqual(value_tag_error, val.elems.?[0].tag);
    const e = val.elems.?[0];
    try std.testing.expectEqualStrings("#DIV/0!", val.payload.?[e.payload_off .. e.payload_off + e.payload_len]);
    zlsx_value_release(&val);

    // Matrix, row-major.
    val = std.mem.zeroes(CValue);
    val.struct_size = @sizeOf(CValue);
    const f_mat = "={1,2;3,4}";
    try std.testing.expectEqual(ZLSX_OK, zlsx_editor_evaluate(ed, f_mat.ptr, f_mat.len, 0, 0, 0, &crun, &val, null, null, &err_buf, err_buf.len));
    try std.testing.expectEqual(@as(u32, 1), val.is_matrix);
    try std.testing.expectEqual(@as(u32, 2), val.rows);
    try std.testing.expectEqual(@as(u32, 2), val.cols);
    try std.testing.expectEqual(@as(usize, 4), val.elems_len);
    try std.testing.expectEqual(@as(f64, 3), val.elems.?[2].num);
    zlsx_value_release(&val);

    // Blank publishes as 0 — the word "blank" never crosses.
    val = std.mem.zeroes(CValue);
    val.struct_size = @sizeOf(CValue);
    const f_blank = "=D7";
    try std.testing.expectEqual(ZLSX_OK, zlsx_editor_evaluate(ed, f_blank.ptr, f_blank.len, 0, 0, 0, &crun, &val, null, null, &err_buf, err_buf.len));
    try std.testing.expectEqual(value_tag_number, val.elems.?[0].tag);
    try std.testing.expectEqual(@as(f64, 0), val.elems.?[0].num);
    zlsx_value_release(&val);

    // Typed refusals populate the diag.
    diag = std.mem.zeroes(CDiag);
    diag.struct_size = @sizeOf(CDiag);
    const f_bad = "=1+";
    try std.testing.expectEqual(ZLSX_REFUSED, zlsx_editor_evaluate(ed, f_bad.ptr, f_bad.len, 0, 0, 0, &crun, &val, null, &diag, &err_buf, err_buf.len));
    try std.testing.expectEqualStrings("FormulaMalformedInput", std.mem.sliceTo(&diag.error_name, 0));
    try std.testing.expectEqual(@intFromEnum(PlaneTwo.FormulaMalformedInput), diag.plane);
    zlsx_diag_release(&diag);

    // §5.7.7's mark-only transaction over the same handle.
    diag = std.mem.zeroes(CDiag);
    diag.struct_size = @sizeOf(CDiag);
    try std.testing.expectEqual(ZLSX_OK, zlsx_editor_mark_recalc_on_load(ed, &diag, &err_buf, err_buf.len));
}

test "M9a1 narrowing: every boundary rejects what the engine never sees" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try writeM9a1Fixture(io, &tt);
    defer std.testing.allocator.free(path);

    var err_buf: [128]u8 = undefined;
    const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
    defer zlsx_editor_close(ed);

    var report = std.mem.zeroes(CRecalcReport);
    var val = std.mem.zeroes(CValue);
    const f = "=1+2";

    // Unknown enum values are contract violations, not refusals.
    inline for (.{ "fidelity", "profile", "dialect", "on_unsupported" }) |field| {
        var crun = zeroRun();
        @field(crun, field) = 99;
        report = std.mem.zeroes(CRecalcReport);
        report.struct_size = @sizeOf(CRecalcReport);
        try std.testing.expectEqual(ZLSX_ERROR, zlsx_editor_recalculate(ed, &crun, &report, null, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings("InvalidInput", std.mem.sliceTo(&err_buf, 0));
    }

    // utc_offset_min outside [-1440, 1440], validated pre-narrowing.
    inline for (.{ @as(i32, 1441), @as(i32, -1441) }) |off| {
        var crun = zeroRun();
        crun.utc_offset_min = off;
        report = std.mem.zeroes(CRecalcReport);
        report.struct_size = @sizeOf(CRecalcReport);
        try std.testing.expectEqual(ZLSX_ERROR, zlsx_editor_recalculate(ed, &crun, &report, null, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings("UtcOffsetOutOfRange", std.mem.sliceTo(&err_buf, 0));
    }

    // A limit above the §9 hard ceiling (4× default) is rejected by the
    // same validator every layer uses.
    {
        var crun = zeroRun();
        crun.max_matrix_cells = std.math.maxInt(u64);
        report = std.mem.zeroes(CRecalcReport);
        report.struct_size = @sizeOf(CRecalcReport);
        try std.testing.expectEqual(ZLSX_ERROR, zlsx_editor_recalculate(ed, &crun, &report, null, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings("LimitOutOfRange", std.mem.sliceTo(&err_buf, 0));
    }

    // Field-width extremes on the evaluate boundary.
    {
        var crun = zeroRun();
        val = std.mem.zeroes(CValue);
        val.struct_size = @sizeOf(CValue);
        try std.testing.expectEqual(ZLSX_ERROR, zlsx_editor_evaluate(ed, f.ptr, f.len, std.math.maxInt(u32), 0, 0, &crun, &val, null, null, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings("SheetNotFound", std.mem.sliceTo(&err_buf, 0));

        try std.testing.expectEqual(ZLSX_ERROR, zlsx_editor_evaluate(ed, f.ptr, f.len, 0, std.math.maxInt(u32), 0, &crun, &val, null, null, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings("InvalidRef", std.mem.sliceTo(&err_buf, 0));

        try std.testing.expectEqual(ZLSX_ERROR, zlsx_editor_evaluate(ed, f.ptr, f.len, 0, 1, std.math.maxInt(u32), &crun, &val, null, null, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings("InvalidRef", std.mem.sliceTo(&err_buf, 0));
    }

    // One past the parser's byte limit refuses as the Plane-2 limit —
    // before the slice is ever formed.
    {
        var crun = zeroRun();
        var diag = std.mem.zeroes(CDiag);
        diag.struct_size = @sizeOf(CDiag);
        val = std.mem.zeroes(CValue);
        val.struct_size = @sizeOf(CValue);
        const big = try std.testing.allocator.alloc(u8, parse_limits_default.max_formula_utf8_bytes + 1);
        defer std.testing.allocator.free(big);
        @memset(big, 'A');
        try std.testing.expectEqual(ZLSX_REFUSED, zlsx_editor_evaluate(ed, big.ptr, big.len, 0, 0, 0, &crun, &val, null, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings("FormulaLimitExceeded", std.mem.sliceTo(&diag.error_name, 0));
        try std.testing.expectEqual(@intFromEnum(PlaneTwo.FormulaLimitExceeded), diag.plane);
    }

    // NULL where required.
    {
        var crun = zeroRun();
        try std.testing.expectEqual(ZLSX_ERROR, zlsx_editor_recalculate(ed, &crun, null, null, &err_buf, err_buf.len));
        try std.testing.expectEqual(ZLSX_ERROR, zlsx_editor_recalculate(null, &crun, &report, null, &err_buf, err_buf.len));
        val = std.mem.zeroes(CValue);
        val.struct_size = @sizeOf(CValue);
        try std.testing.expectEqual(ZLSX_ERROR, zlsx_editor_evaluate(ed, null, 5, 0, 0, 0, &crun, &val, null, null, &err_buf, err_buf.len));
    }
}

test "M9a1 canary-tail: bytes beyond the known prefix are never written" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try writeM9a1Fixture(io, &tt);
    defer std.testing.allocator.free(path);

    var err_buf: [128]u8 = undefined;
    const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
    defer zlsx_editor_close(ed);
    var crun = zeroRun();

    // A "newer caller" declares more than this library knows. Success
    // path: the tail stays untouched.
    {
        var buf: [@sizeOf(CRecalcReport) + 64]u8 align(@alignOf(CRecalcReport)) = undefined;
        @memset(&buf, 0xAA);
        const rp: *CRecalcReport = @ptrCast(&buf);
        rp.struct_size = buf.len;
        try std.testing.expectEqual(ZLSX_OK, zlsx_editor_recalculate(ed, &crun, rp, null, &err_buf, err_buf.len));
        for (buf[@sizeOf(CRecalcReport)..]) |b| try std.testing.expectEqual(@as(u8, 0xAA), b);
        zlsx_recalc_report_release(rp);
    }
    // Failure path (unknown fidelity): tail still untouched.
    {
        var bad = zeroRun();
        bad.fidelity = 99;
        var buf: [@sizeOf(CRecalcReport) + 64]u8 align(@alignOf(CRecalcReport)) = undefined;
        @memset(&buf, 0xAA);
        const rp: *CRecalcReport = @ptrCast(&buf);
        rp.struct_size = buf.len;
        try std.testing.expectEqual(ZLSX_ERROR, zlsx_editor_recalculate(ed, &bad, rp, null, &err_buf, err_buf.len));
        for (buf[@sizeOf(CRecalcReport)..]) |b| try std.testing.expectEqual(@as(u8, 0xAA), b);
    }
    // Below the v1 minimum: rejected before ANY byte is written — the
    // zero-init itself must not have happened.
    {
        var buf: [@sizeOf(CRecalcReport)]u8 align(@alignOf(CRecalcReport)) = undefined;
        @memset(&buf, 0xAA);
        const rp: *CRecalcReport = @ptrCast(&buf);
        rp.struct_size = @sizeOf(CRecalcReport) - 1;
        try std.testing.expectEqual(ZLSX_ERROR, zlsx_editor_recalculate(ed, &crun, rp, null, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings("StructSizeTooSmall", std.mem.sliceTo(&err_buf, 0));
        for (buf[@sizeOf(usize)..]) |b| try std.testing.expectEqual(@as(u8, 0xAA), b);
    }
    // Same discipline on the value and diag structs.
    {
        var buf: [@sizeOf(CValue) + 32]u8 align(@alignOf(CValue)) = undefined;
        @memset(&buf, 0xAA);
        const vp: *CValue = @ptrCast(&buf);
        vp.struct_size = buf.len;
        const f = "=1+2";
        try std.testing.expectEqual(ZLSX_OK, zlsx_editor_evaluate(ed, f.ptr, f.len, 0, 0, 0, &crun, vp, null, null, &err_buf, err_buf.len));
        for (buf[@sizeOf(CValue)..]) |b| try std.testing.expectEqual(@as(u8, 0xAA), b);
        zlsx_value_release(vp);
    }
    {
        var buf: [@sizeOf(CDiag) + 32]u8 align(@alignOf(CDiag)) = undefined;
        @memset(&buf, 0xAA);
        const dp: *CDiag = @ptrCast(&buf);
        dp.struct_size = buf.len;
        var val = std.mem.zeroes(CValue);
        val.struct_size = @sizeOf(CValue);
        const f_bad = "=1+";
        try std.testing.expectEqual(ZLSX_REFUSED, zlsx_editor_evaluate(ed, f_bad.ptr, f_bad.len, 0, 0, 0, &crun, &val, null, dp, &err_buf, err_buf.len));
        for (buf[@sizeOf(CDiag)..]) |b| try std.testing.expectEqual(@as(u8, 0xAA), b);
        zlsx_diag_release(dp);
    }
}

test "M9a1 release fns: NULL-safe, no-ops on zeroed structs, idempotent" {
    zlsx_value_release(null);
    zlsx_recalc_report_release(null);
    zlsx_diag_release(null);

    var val = std.mem.zeroes(CValue);
    zlsx_value_release(&val);
    zlsx_value_release(&val);
    var report = std.mem.zeroes(CRecalcReport);
    zlsx_recalc_report_release(&report);
    zlsx_recalc_report_release(&report);
    var diag = std.mem.zeroes(CDiag);
    zlsx_diag_release(&diag);
    zlsx_diag_release(&diag);
}

test "fuzz C ABI M9a1: random runs, formulas and anchors never panic, tails stay intact" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try writeM9a1Fixture(io, &tt);
    defer std.testing.allocator.free(path);

    var err_buf: [128]u8 = undefined;
    const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
    defer zlsx_editor_close(ed);

    var tok: ?*CancelTok = null;
    try std.testing.expectEqual(ZLSX_OK, zlsx_cancel_token_new(&tok, &err_buf, err_buf.len));
    defer zlsx_cancel_token_free(tok);

    var prng = std.Random.DefaultPrng.init(fuzz_config.seed_override orelse 0x9a1_c_ab1);
    const random = prng.random();

    var iter: usize = 0;
    while (iter < fuzzItersCabi()) : (iter += 1) {
        var crun: CRun = undefined;
        random.bytes(std.mem.asBytes(&crun));
        // The only pointer field must be NULL or a real token — a wild
        // pointer would be the caller's UB, not the boundary's.
        crun.cancel = if (random.boolean()) null else tok;
        // Keep the deadline either off or short-lived.
        if (random.boolean()) crun.timeout_ms = 0;
        // Exercise both sides of the struct_size gate.
        if (random.boolean()) crun.struct_size = @sizeOf(CRun);

        var fbuf: [24]u8 = undefined;
        random.bytes(&fbuf);
        const flen = random.uintLessThan(usize, fbuf.len + 1);

        var vbuf: [@sizeOf(CValue) + 16]u8 align(@alignOf(CValue)) = undefined;
        @memset(&vbuf, 0xAA);
        const vp: *CValue = @ptrCast(&vbuf);
        vp.struct_size = if (random.boolean()) vbuf.len else random.uintLessThan(usize, vbuf.len);

        var dbuf: [@sizeOf(CDiag) + 16]u8 align(@alignOf(CDiag)) = undefined;
        @memset(&dbuf, 0xAA);
        const dp: *CDiag = @ptrCast(&dbuf);
        dp.struct_size = if (random.boolean()) dbuf.len else random.uintLessThan(usize, dbuf.len);
        const use_diag = random.boolean();

        const st_eval = zlsx_editor_evaluate(
            ed,
            &fbuf,
            flen,
            random.uintLessThan(u32, 3),
            if (random.boolean()) 0 else random.int(u32),
            random.int(u32),
            &crun,
            vp,
            null,
            if (use_diag) dp else null,
            &err_buf,
            err_buf.len,
        );
        try std.testing.expect(st_eval == ZLSX_OK or st_eval == ZLSX_ERROR or
            st_eval == ZLSX_REFUSED or st_eval == ZLSX_NOMEM or st_eval == ZLSX_CANCELLED);
        for (vbuf[@sizeOf(CValue)..]) |b| try std.testing.expectEqual(@as(u8, 0xAA), b);
        for (dbuf[@sizeOf(CDiag)..]) |b| try std.testing.expectEqual(@as(u8, 0xAA), b);
        if (st_eval == ZLSX_OK) zlsx_value_release(vp);
        // Release only what the library accepted: a rejected
        // struct_size means the interior is still the caller's canary
        // garbage, and releasing THAT is the caller's UB, not ours.
        if (use_diag and dp.struct_size >= @sizeOf(CDiag)) zlsx_diag_release(dp);

        var rbuf: [@sizeOf(CRecalcReport) + 16]u8 align(@alignOf(CRecalcReport)) = undefined;
        @memset(&rbuf, 0xAA);
        const rp: *CRecalcReport = @ptrCast(&rbuf);
        rp.struct_size = if (random.boolean()) rbuf.len else random.uintLessThan(usize, rbuf.len);
        const st_recalc = zlsx_editor_recalculate(ed, &crun, rp, null, &err_buf, err_buf.len);
        try std.testing.expect(st_recalc == ZLSX_OK or st_recalc == ZLSX_ERROR or
            st_recalc == ZLSX_REFUSED or st_recalc == ZLSX_NOMEM or st_recalc == ZLSX_CANCELLED);
        for (rbuf[@sizeOf(CRecalcReport)..]) |b| try std.testing.expectEqual(@as(u8, 0xAA), b);
        if (st_recalc == ZLSX_OK) zlsx_recalc_report_release(rp);
    }
}

// ─── M9a2 tests ──────────────────────────────────────────────────────

/// The M9a1 fixture with the formula swapped for one the engine does
/// not implement — the refusal-census fixture.
fn writeM9a2RefusalFixture(io: std.Io, tt: *TestTmp) ![:0]u8 {
    const alloc = std.testing.allocator;
    const path = try tt.path(alloc, io, "m9a2_refusal.xlsx");
    errdefer alloc.free(path);
    var err_buf: [128]u8 = undefined;
    const w = zlsx_writer_create(&err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
    defer zlsx_writer_close(w);
    const name = "Sheet1";
    const sw = zlsx_writer_add_sheet(w, name.ptr, name.len, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
    const cells = [_]CCell{ toCCell(.{ .number = 1 }), toCCell(.{ .number = 999 }) };
    const formula = "NOSUCHFN(A1)";
    const fptrs = [_]?[*]const u8{ null, formula.ptr };
    const flens = [_]usize{ 0, formula.len };
    if (zlsx_sheet_writer_write_row_with_formulas(sw, &cells, &fptrs, &flens, cells.len, &err_buf, err_buf.len) != 0)
        return error.TestUnexpectedResult;
    if (zlsx_writer_save(w, path.ptr, path.len, &err_buf, err_buf.len) != 0)
        return error.TestUnexpectedResult;
    return path;
}

fn readFileBytes(io: std.Io, path: []const u8) ![]u8 {
    return std.Io.Dir.cwd().readFileAlloc(io, path, std.testing.allocator, .limited(1 << 24));
}

test "M9a2 end-to-end: open_buffer, save_to_buffer, save_with_recalc" {
    const alloc = std.testing.allocator;
    var threaded: std.Io.Threaded = .init(alloc, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try writeM9a1Fixture(io, &tt);
    defer alloc.free(path);
    const src_bytes = try readFileBytes(io, path);
    defer alloc.free(src_bytes);

    var err_buf: [128]u8 = undefined;

    // open_buffer: the borrow ends at the call — poison a copy after.
    var ed_slot: ?*Editor = null;
    {
        const borrowed = try alloc.dupe(u8, src_bytes);
        defer alloc.free(borrowed);
        try std.testing.expectEqual(ZLSX_OK, zlsx_open_buffer(borrowed.ptr, borrowed.len, &ed_slot, &err_buf, err_buf.len));
        @memset(borrowed, 0xAA);
    }
    const ed = ed_slot orelse return error.TestUnexpectedResult;
    defer zlsx_editor_close(ed);

    // save_to_buffer of an untouched editor: the source bytes verbatim.
    var out_ptr: ?[*]u8 = null;
    var out_len: usize = 0;
    try std.testing.expectEqual(ZLSX_OK, zlsx_editor_save_to_buffer(ed, &out_ptr, &out_len, &err_buf, err_buf.len));
    try std.testing.expectEqualSlices(u8, src_bytes, out_ptr.?[0..out_len]);
    zlsx_buffer_release(out_ptr, out_len);

    // save_with_recalc: the destination holds the recalced cache and
    // the durability slot reports a clean commit.
    const out_path = try tt.path(alloc, io, "m9a2_out.xlsx");
    defer alloc.free(out_path);
    var crun = zeroRun();
    crun.now_utc_ms = 1_700_000_000_000;
    crun.rng_seed = 42;
    var report = std.mem.zeroes(CRecalcReport);
    report.struct_size = @sizeOf(CRecalcReport);
    var diag = std.mem.zeroes(CDiag);
    diag.struct_size = @sizeOf(CDiag);
    try std.testing.expectEqual(ZLSX_OK, zlsx_editor_save_with_recalc(ed, out_path.ptr, out_path.len, &crun, &report, &diag, &err_buf, err_buf.len));
    try std.testing.expect(report.cells_written >= 1);
    try std.testing.expectEqual(@as(u32, 1), report.resolved_present);
    try std.testing.expectEqual(@as(u32, 0), report.durability_warning);
    try std.testing.expectEqual(@as(i32, 0), report.durability_errno);
    try std.testing.expectEqual(plane_none, diag.plane);
    zlsx_recalc_report_release(&report);
    zlsx_diag_release(&diag);

    // And the memory swapped with the file: evaluating on the SAME
    // handle reads the recalced cache.
    var val = std.mem.zeroes(CValue);
    val.struct_size = @sizeOf(CValue);
    const f_read = "=B1";
    try std.testing.expectEqual(ZLSX_OK, zlsx_editor_evaluate(ed, f_read.ptr, f_read.len, 0, 0, 0, &crun, &val, null, null, &err_buf, err_buf.len));
    try std.testing.expectEqual(@as(f64, 3), val.elems.?[0].num);
    zlsx_value_release(&val);

    // The destination reopens through open_buffer and reads the same.
    const out_bytes = try readFileBytes(io, out_path);
    defer alloc.free(out_bytes);
    var ed2_slot: ?*Editor = null;
    try std.testing.expectEqual(ZLSX_OK, zlsx_open_buffer(out_bytes.ptr, out_bytes.len, &ed2_slot, &err_buf, err_buf.len));
    const ed2 = ed2_slot orelse return error.TestUnexpectedResult;
    defer zlsx_editor_close(ed2);
    val = std.mem.zeroes(CValue);
    val.struct_size = @sizeOf(CValue);
    try std.testing.expectEqual(ZLSX_OK, zlsx_editor_evaluate(ed2, f_read.ptr, f_read.len, 0, 0, 0, &crun, &val, null, null, &err_buf, err_buf.len));
    try std.testing.expectEqual(@as(f64, 3), val.elems.?[0].num);
    zlsx_value_release(&val);
}

test "M9a2 refusal: the diag census names the refusing cell across the boundary" {
    const alloc = std.testing.allocator;
    var threaded: std.Io.Threaded = .init(alloc, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try writeM9a2RefusalFixture(io, &tt);
    defer alloc.free(path);

    var err_buf: [128]u8 = undefined;
    const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
    defer zlsx_editor_close(ed);

    var crun = zeroRun();
    crun.now_utc_ms = 1_700_000_000_000;
    crun.rng_seed = 7;
    var report = std.mem.zeroes(CRecalcReport);
    report.struct_size = @sizeOf(CRecalcReport);
    var diag = std.mem.zeroes(CDiag);
    diag.struct_size = @sizeOf(CDiag);
    try std.testing.expectEqual(ZLSX_REFUSED, zlsx_editor_recalculate(ed, &crun, &report, &diag, &err_buf, err_buf.len));
    try std.testing.expectEqualStrings("FormulaUnsupportedFunction", std.mem.sliceTo(&diag.error_name, 0));
    try std.testing.expectEqual(@intFromEnum(PlaneTwo.FormulaUnsupportedFunction), diag.plane);
    // M9a1 shipped this census empty (decision 4); M9a2 makes it name
    // the refusing cell: B1 = row 1 (1-based), col 1 (0-based).
    try std.testing.expectEqual(@as(usize, 1), diag.census_len);
    const entry = diag.census.?[0];
    try std.testing.expectEqual(@intFromEnum(PlaneTwo.FormulaUnsupportedFunction), entry.plane);
    try std.testing.expectEqual(@as(u32, 0), entry.sheet);
    try std.testing.expectEqual(@as(u32, 1), entry.row);
    try std.testing.expectEqual(@as(u32, 1), entry.col);
    try std.testing.expectEqual(@as(u32, 0), diag.census_truncated);
    zlsx_diag_release(&diag);
    // The same refusal through the file transaction: same census, and
    // the destination is never created.
    const out_path = try tt.path(alloc, io, "m9a2_refused_out.xlsx");
    defer alloc.free(out_path);
    diag = std.mem.zeroes(CDiag);
    diag.struct_size = @sizeOf(CDiag);
    try std.testing.expectEqual(ZLSX_REFUSED, zlsx_editor_save_with_recalc(ed, out_path.ptr, out_path.len, &crun, &report, &diag, &err_buf, err_buf.len));
    try std.testing.expectEqual(@as(usize, 1), diag.census_len);
    try std.testing.expectEqual(@as(u32, 1), diag.census.?[0].row);
    zlsx_diag_release(&diag);
    try std.testing.expectError(error.FileNotFound, std.Io.Dir.cwd().openFile(io, out_path, .{}));
}

test "M9a2 writer: save_with_recalc computes the fresh workbook's caches" {
    const alloc = std.testing.allocator;
    var threaded: std.Io.Threaded = .init(alloc, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    var err_buf: [128]u8 = undefined;

    const w = zlsx_writer_create(&err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
    defer zlsx_writer_close(w);
    const name = "Sheet1";
    const sw = zlsx_writer_add_sheet(w, name.ptr, name.len, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
    const cells = [_]CCell{ toCCell(.{ .number = 1 }), toCCell(.{ .number = 0 }) };
    const formula = "A1+2";
    const fptrs = [_]?[*]const u8{ null, formula.ptr };
    const flens = [_]usize{ 0, formula.len };
    try std.testing.expectEqual(@as(i32, 0), zlsx_sheet_writer_write_row_with_formulas(sw, &cells, &fptrs, &flens, cells.len, &err_buf, err_buf.len));

    const out_path = try tt.path(alloc, io, "m9a2_writer_out.xlsx");
    defer alloc.free(out_path);
    var crun = zeroRun();
    crun.now_utc_ms = 1_700_000_000_000;
    crun.rng_seed = 9;
    var report = std.mem.zeroes(CRecalcReport);
    report.struct_size = @sizeOf(CRecalcReport);
    var diag = std.mem.zeroes(CDiag);
    diag.struct_size = @sizeOf(CDiag);
    try std.testing.expectEqual(ZLSX_OK, zlsx_writer_save_with_recalc(w, out_path.ptr, out_path.len, &crun, &report, &diag, &err_buf, err_buf.len));
    try std.testing.expect(report.cells_written >= 1);
    try std.testing.expectEqual(@as(u32, 0), report.durability_warning);
    zlsx_recalc_report_release(&report);
    zlsx_diag_release(&diag);

    // The writer handle survives (not consumed): a plain save still works.
    const plain_path = try tt.path(alloc, io, "m9a2_writer_plain.xlsx");
    defer alloc.free(plain_path);
    try std.testing.expectEqual(@as(i32, 0), zlsx_writer_save(w, plain_path.ptr, plain_path.len, &err_buf, err_buf.len));

    // The recalced destination carries B1 = 3.
    const ed = zlsx_editor_open(out_path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
    defer zlsx_editor_close(ed);
    var val = std.mem.zeroes(CValue);
    val.struct_size = @sizeOf(CValue);
    const f_read = "=B1";
    try std.testing.expectEqual(ZLSX_OK, zlsx_editor_evaluate(ed, f_read.ptr, f_read.len, 0, 0, 0, &crun, &val, null, null, &err_buf, err_buf.len));
    try std.testing.expectEqual(@as(f64, 3), val.elems.?[0].num);
    zlsx_value_release(&val);
}

test "M9a2 v2 rows: the descriptor writes a CSE rectangle and refuses bad shapes" {
    const alloc = std.testing.allocator;
    var threaded: std.Io.Threaded = .init(alloc, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    var err_buf: [128]u8 = undefined;

    const w = zlsx_writer_create(&err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
    defer zlsx_writer_close(w);
    const name = "Sheet1";
    const sw = zlsx_writer_add_sheet(w, name.ptr, name.len, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;

    const plain: CFormulaCell = .{ .text = null, .text_len = 0, .dialect = formula_dialect_scalar, ._reserved0 = 0, .ref = null, .ref_len = 0 };
    const cse_text = "TRANSPOSE(A3:B3)";
    const cse_ref = "A1:B2";

    // Bad shapes first (nothing may land): unknown dialect / CSE
    // without ref / dynamic-array / anchor mismatch / empty text.
    const t = "1+2";
    var bad: [2]CFormulaCell = .{ .{ .text = t.ptr, .text_len = t.len, .dialect = 7, ._reserved0 = 0, .ref = null, .ref_len = 0 }, plain };
    const two_cells = [_]CCell{ toCCell(.{ .number = 3 }), toCCell(.{ .number = 4 }) };
    try std.testing.expectEqual(ZLSX_ERROR, zlsx_sheet_writer_write_row_with_formulas_v2(sw, &two_cells, &bad, two_cells.len, &err_buf, err_buf.len));
    bad[0] = .{ .text = t.ptr, .text_len = t.len, .dialect = formula_dialect_cse, ._reserved0 = 0, .ref = null, .ref_len = 0 };
    try std.testing.expectEqual(ZLSX_ERROR, zlsx_sheet_writer_write_row_with_formulas_v2(sw, &two_cells, &bad, two_cells.len, &err_buf, err_buf.len));
    bad[0] = .{ .text = t.ptr, .text_len = t.len, .dialect = formula_dialect_dynamic_array, ._reserved0 = 0, .ref = null, .ref_len = 0 };
    try std.testing.expectEqual(ZLSX_ERROR, zlsx_sheet_writer_write_row_with_formulas_v2(sw, &two_cells, &bad, two_cells.len, &err_buf, err_buf.len));
    try std.testing.expectEqualStrings("FormulaDynamicArrayUnsupported", std.mem.sliceTo(&err_buf, 0));
    const wrong_ref = "B1:B2";
    bad[0] = .{ .text = cse_text.ptr, .text_len = cse_text.len, .dialect = formula_dialect_cse, ._reserved0 = 0, .ref = wrong_ref.ptr, .ref_len = wrong_ref.len };
    try std.testing.expectEqual(ZLSX_ERROR, zlsx_sheet_writer_write_row_with_formulas_v2(sw, &two_cells, &bad, two_cells.len, &err_buf, err_buf.len));
    try std.testing.expectEqualStrings("FormulaCseAnchorMismatch", std.mem.sliceTo(&err_buf, 0));
    bad[0] = .{ .text = t.ptr, .text_len = 0, .dialect = formula_dialect_scalar, ._reserved0 = 0, .ref = null, .ref_len = 0 };
    try std.testing.expectEqual(ZLSX_ERROR, zlsx_sheet_writer_write_row_with_formulas_v2(sw, &two_cells, &bad, two_cells.len, &err_buf, err_buf.len));

    // Nothing landed: the sheet still starts at row 1. Now the real
    // anchor row: A1 anchors A1:B2, B1 is a member with a cached value.
    const anchor: CFormulaCell = .{ .text = cse_text.ptr, .text_len = cse_text.len, .dialect = formula_dialect_cse, ._reserved0 = 0, .ref = cse_ref.ptr, .ref_len = cse_ref.len };
    const row1 = [_]CFormulaCell{ anchor, plain };
    try std.testing.expectEqual(ZLSX_OK, zlsx_sheet_writer_write_row_with_formulas_v2(sw, &two_cells, &row1, two_cells.len, &err_buf, err_buf.len));

    // A formula inside the open rectangle refuses.
    var member_bad = [_]CFormulaCell{ .{ .text = t.ptr, .text_len = t.len, .dialect = formula_dialect_scalar, ._reserved0 = 0, .ref = null, .ref_len = 0 }, plain };
    try std.testing.expectEqual(ZLSX_ERROR, zlsx_sheet_writer_write_row_with_formulas_v2(sw, &two_cells, &member_bad, two_cells.len, &err_buf, err_buf.len));
    try std.testing.expectEqualStrings("FormulaCseMemberFormula", std.mem.sliceTo(&err_buf, 0));

    // An incomplete rectangle refuses the save.
    const early_path = try tt.path(alloc, io, "m9a2_v2_early.xlsx");
    defer alloc.free(early_path);
    try std.testing.expectEqual(@as(i32, -1), zlsx_writer_save(w, early_path.ptr, early_path.len, &err_buf, err_buf.len));
    try std.testing.expectEqualStrings("FormulaCseMemberMissing", std.mem.sliceTo(&err_buf, 0));

    // Row 2 completes it: A2 empty (becomes a placeholder), B2 valued.
    const row2_cells = [_]CCell{ toCCell(.empty), toCCell(.{ .number = 6 }) };
    const row2 = [_]CFormulaCell{ plain, plain };
    try std.testing.expectEqual(ZLSX_OK, zlsx_sheet_writer_write_row_with_formulas_v2(sw, &row2_cells, &row2, row2_cells.len, &err_buf, err_buf.len));

    const out_path = try tt.path(alloc, io, "m9a2_v2.xlsx");
    defer alloc.free(out_path);
    try std.testing.expectEqual(@as(i32, 0), zlsx_writer_save(w, out_path.ptr, out_path.len, &err_buf, err_buf.len));

    // The archive opens and the anchor's rectangle survives the reader.
    const book = zlsx_book_open(out_path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
    defer zlsx_book_close(book);
}

test "M9a2 boundary: the new exports reject what the library never sees" {
    const alloc = std.testing.allocator;
    var threaded: std.Io.Threaded = .init(alloc, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try writeM9a1Fixture(io, &tt);
    defer alloc.free(path);
    const src_bytes = try readFileBytes(io, path);
    defer alloc.free(src_bytes);
    var err_buf: [128]u8 = undefined;

    // zlsx_buffer_release: NULL-safe, zero-length safe.
    zlsx_buffer_release(null, 0);
    zlsx_buffer_release(null, 17);

    // open_buffer contract violations.
    try std.testing.expectEqual(ZLSX_ERROR, zlsx_open_buffer(src_bytes.ptr, src_bytes.len, null, &err_buf, err_buf.len));
    var slot: ?*Editor = @ptrFromInt(@alignOf(EditorState)); // poisoned: must be nulled
    try std.testing.expectEqual(ZLSX_ERROR, zlsx_open_buffer(null, 0, &slot, &err_buf, err_buf.len));
    try std.testing.expectEqual(@as(?*Editor, null), slot);
    const garbage = "not a zip archive at all";
    slot = @ptrFromInt(@alignOf(EditorState));
    try std.testing.expectEqual(ZLSX_ERROR, zlsx_open_buffer(garbage.ptr, garbage.len, &slot, &err_buf, err_buf.len));
    try std.testing.expectEqual(@as(?*Editor, null), slot);

    // save_to_buffer contract violations, and the prepped (NULL, 0) pair.
    var out_ptr: ?[*]u8 = @ptrFromInt(8);
    var out_len: usize = 99;
    try std.testing.expectEqual(ZLSX_ERROR, zlsx_editor_save_to_buffer(null, &out_ptr, &out_len, &err_buf, err_buf.len));
    var ed_slot: ?*Editor = null;
    try std.testing.expectEqual(ZLSX_OK, zlsx_open_buffer(src_bytes.ptr, src_bytes.len, &ed_slot, &err_buf, err_buf.len));
    const ed = ed_slot orelse return error.TestUnexpectedResult;
    defer zlsx_editor_close(ed);
    try std.testing.expectEqual(ZLSX_ERROR, zlsx_editor_save_to_buffer(ed, null, &out_len, &err_buf, err_buf.len));
    try std.testing.expectEqual(ZLSX_ERROR, zlsx_editor_save_to_buffer(ed, &out_ptr, null, &err_buf, err_buf.len));

    // save_with_recalc: NULL path / empty path / bad report size — and
    // a canary tail that survives every failure class.
    var crun = zeroRun();
    crun.now_utc_ms = 1;
    crun.rng_seed = 1;
    var rbuf: [@sizeOf(CRecalcReport) + 64]u8 align(@alignOf(CRecalcReport)) = undefined;
    @memset(&rbuf, 0xAA);
    const rp: *CRecalcReport = @ptrCast(&rbuf);
    rp.struct_size = rbuf.len;
    const out_path = try tt.path(alloc, io, "m9a2_boundary_out.xlsx");
    defer alloc.free(out_path);
    try std.testing.expectEqual(ZLSX_ERROR, zlsx_editor_save_with_recalc(ed, null, 5, &crun, rp, null, &err_buf, err_buf.len));
    try std.testing.expectEqual(ZLSX_ERROR, zlsx_editor_save_with_recalc(ed, out_path.ptr, 0, &crun, rp, null, &err_buf, err_buf.len));
    var bad_run = zeroRun();
    bad_run.now_utc_ms = 1;
    bad_run.rng_seed = 1;
    bad_run.fidelity = 99;
    try std.testing.expectEqual(ZLSX_ERROR, zlsx_editor_save_with_recalc(ed, out_path.ptr, out_path.len, &bad_run, rp, null, &err_buf, err_buf.len));
    for (rbuf[@sizeOf(CRecalcReport)..]) |b| try std.testing.expectEqual(@as(u8, 0xAA), b);
    // Below-minimum report: byte-for-byte untouched beyond struct_size.
    @memset(&rbuf, 0xAA);
    rp.struct_size = @sizeOf(CRecalcReport) - 1;
    try std.testing.expectEqual(ZLSX_ERROR, zlsx_editor_save_with_recalc(ed, out_path.ptr, out_path.len, &crun, rp, null, &err_buf, err_buf.len));
    try std.testing.expectEqualStrings("StructSizeTooSmall", std.mem.sliceTo(&err_buf, 0));
    for (rbuf[@sizeOf(usize)..]) |b| try std.testing.expectEqual(@as(u8, 0xAA), b);
    // And no destination appeared through any of it.
    try std.testing.expectError(error.FileNotFound, std.Io.Dir.cwd().openFile(io, out_path, .{}));

    // v2 writer export: NULL handle refuses.
    const dcell = [_]CCell{toCCell(.{ .number = 1 })};
    const dform = [_]CFormulaCell{.{ .text = null, .text_len = 0, .dialect = 0, ._reserved0 = 0, .ref = null, .ref_len = 0 }};
    try std.testing.expectEqual(ZLSX_ERROR, zlsx_sheet_writer_write_row_with_formulas_v2(null, &dcell, &dform, 1, &err_buf, err_buf.len));
}

test "fuzz C ABI M9a2: random formula descriptors never panic" {
    const alloc = std.testing.allocator;
    var threaded: std.Io.Threaded = .init(alloc, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var err_buf: [128]u8 = undefined;

    const w = zlsx_writer_create(&err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
    defer zlsx_writer_close(w);
    const name = "Fuzz";
    const sw = zlsx_writer_add_sheet(w, name.ptr, name.len, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;

    var prng = std.Random.DefaultPrng.init(fuzzSeedCabi(io));
    const random = prng.random();
    const text_pool = "SUM(A1:B2)";
    const ref_pool = "A1:B2XYZ:9";

    var iter: usize = 0;
    while (iter < fuzzItersCabi()) : (iter += 1) {
        var cells: [4]CCell = undefined;
        var descs: [4]CFormulaCell = undefined;
        const n = 1 + random.uintLessThan(usize, cells.len);
        for (0..n) |i| {
            cells[i] = toCCell(.{ .number = @floatFromInt(random.uintLessThan(u32, 1000)) });
            random.bytes(std.mem.asBytes(&descs[i]));
            // Re-pin every pointer to NULL or a valid pool slice —
            // the fuzz target is the descriptor grammar, not wild
            // memory.
            const text_len = random.uintLessThan(usize, text_pool.len + 1);
            descs[i].text = if (random.boolean()) text_pool.ptr else null;
            descs[i].text_len = if (descs[i].text != null) text_len else if (random.boolean()) 0 else text_len;
            const ref_len = random.uintLessThan(usize, ref_pool.len + 1);
            descs[i].ref = if (random.boolean()) ref_pool.ptr else null;
            descs[i].ref_len = if (descs[i].ref != null) ref_len else if (random.boolean()) 0 else ref_len;
            descs[i].dialect = random.uintLessThan(u32, 5);
        }
        const st = zlsx_sheet_writer_write_row_with_formulas_v2(sw, &cells, &descs, n, &err_buf, err_buf.len);
        try std.testing.expect(st == ZLSX_OK or st == ZLSX_ERROR or st == ZLSX_NOMEM);
    }
}

// ─── S3a tests ───────────────────────────────────────────────────────

/// Two sheets: `Data` = a 3×3 integer grid, `Second` = one string cell.
/// Written through the Zig writer — the fixture is not what these tests
/// prove; the boundary is.
fn writeS3aFixture(io: std.Io, tt: *TestTmp, name: []const u8) ![:0]u8 {
    const alloc = std.testing.allocator;
    const path = try tt.path(alloc, io, name);
    errdefer alloc.free(path);
    var w = xlsx.Writer.init(alloc);
    defer w.deinit();
    var data = try w.addSheet("Data");
    try data.writeRow(&.{ .{ .integer = 1 }, .{ .integer = 2 }, .{ .integer = 3 } });
    try data.writeRow(&.{ .{ .integer = 4 }, .{ .integer = 5 }, .{ .integer = 6 } });
    try data.writeRow(&.{ .{ .integer = 7 }, .{ .integer = 8 }, .{ .integer = 9 } });
    var second = try w.addSheet("Second");
    try second.writeRow(&.{.{ .string = "two" }});
    try w.save(io, path);
    return path;
}

fn freshDiag() CDiag {
    var diag = std.mem.zeroes(CDiag);
    diag.struct_size = @sizeOf(CDiag);
    return diag;
}

fn diagName(diag: *const CDiag) []const u8 {
    return std.mem.sliceTo(&diag.error_name, 0);
}

/// The bytes of one part of a saved workbook, through the package
/// store — the same reader every edit surface reopens with.
fn savedPartBytes(alloc: std.mem.Allocator, io: std.Io, path: []const u8, part_name: []const u8) ![]u8 {
    var wb = try zlsx_pkg.Workbook.open(alloc, io, path);
    defer wb.deinit();
    const part = (try wb.store.part(part_name)) orelse return error.TestUnexpectedResult;
    return alloc.dupe(u8, part.bytes);
}

test "S3a end-to-end: row, column and sheet edits cross the boundary and land in the saved file" {
    const alloc = std.testing.allocator;
    var threaded: std.Io.Threaded = .init(alloc, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try writeS3aFixture(io, &tt, "s3a_src.xlsx");
    defer alloc.free(path);
    const out_path = try tt.path(alloc, io, "s3a_out.xlsx");
    defer alloc.free(out_path);

    var err_buf: [128]u8 = undefined;
    const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
    defer zlsx_editor_close(ed);

    var diag = freshDiag();
    // Row 2 becomes blank; the old row 2 is row 3.
    try std.testing.expectEqual(ZLSX_OK, zlsx_editor_insert_row(ed, 0, 2, &diag, &err_buf, err_buf.len));
    // Column A (0-based 0) goes; B..C become A..B.
    try std.testing.expectEqual(ZLSX_OK, zlsx_editor_delete_column(ed, 0, 0, &diag, &err_buf, err_buf.len));
    // A blank column before the new A: the grid is now [_, 2, 3].
    try std.testing.expectEqual(ZLSX_OK, zlsx_editor_insert_column(ed, 0, 0, &diag, &err_buf, err_buf.len));
    // And the last row goes.
    try std.testing.expectEqual(ZLSX_OK, zlsx_editor_delete_row(ed, 0, 4, &diag, &err_buf, err_buf.len));

    var new_idx: u32 = 5;
    const third = "Third";
    try std.testing.expectEqual(ZLSX_OK, zlsx_editor_add_sheet(ed, third.ptr, third.len, &new_idx, &diag, &err_buf, err_buf.len));
    try std.testing.expectEqual(@as(u32, 2), new_idx);
    const renamed = "Renamed";
    try std.testing.expectEqual(ZLSX_OK, zlsx_editor_rename_sheet(ed, 1, renamed.ptr, renamed.len, &diag, &err_buf, err_buf.len));
    // A NULL diag is a legal way to not ask.
    try std.testing.expectEqual(ZLSX_OK, zlsx_editor_delete_sheet(ed, 2, null, &err_buf, err_buf.len));
    try std.testing.expectEqual(plane_none, diag.plane);
    zlsx_diag_release(&diag);

    try std.testing.expectEqual(@as(i32, 0), zlsx_editor_save(ed, out_path.ptr, out_path.len, &err_buf, err_buf.len));

    var book = try xlsx.Book.open(alloc, io, out_path);
    defer book.deinit();
    try std.testing.expectEqual(@as(usize, 2), book.sheets.len);
    try std.testing.expectEqualStrings("Data", book.sheets[0].name);
    try std.testing.expectEqualStrings("Renamed", book.sheets[1].name);

    // The grid the edits leave: row 1 = [_,2,3], row 2 blank, row 3 = [_,5,6].
    const sheet_xml = try savedPartBytes(alloc, io, out_path, "xl/worksheets/sheet1.xml");
    defer alloc.free(sheet_xml);
    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "<c r=\"B1\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "<c r=\"C3\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "<c r=\"A1\"") == null);
    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "<row r=\"4\"") == null);
    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "<v>9</v>") == null);
    try std.testing.expect(std.mem.indexOf(u8, sheet_xml, "<v>6</v>") != null);
}

test "S3a rename_table_column: the table part and the header cell follow the new name" {
    const alloc = std.testing.allocator;
    var threaded: std.Io.Threaded = .init(alloc, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(alloc, io, "s3a_table.xlsx");
    defer alloc.free(path);
    try zlsx_pkg.pivots.fixture.write(alloc, io, path, .table_name);
    const out_path = try tt.path(alloc, io, "s3a_table_out.xlsx");
    defer alloc.free(out_path);

    var err_buf: [128]u8 = undefined;
    const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
    defer zlsx_editor_close(ed);
    var diag = freshDiag();
    const tbl = "SalesTbl";
    const old = "Qty";
    const new = "Quantity";
    try std.testing.expectEqual(ZLSX_OK, zlsx_editor_rename_table_column(ed, tbl.ptr, tbl.len, old.ptr, old.len, new.ptr, new.len, &diag, &err_buf, err_buf.len));
    // The old name is gone: a selector that names nothing is about the
    // call, -1 (a fresh target, so the name-in-use check that precedes
    // the lookup stays out of the way). An empty new name likewise.
    const other = "Count";
    try std.testing.expectEqual(ZLSX_ERROR, zlsx_editor_rename_table_column(ed, tbl.ptr, tbl.len, old.ptr, old.len, other.ptr, other.len, &diag, &err_buf, err_buf.len));
    try std.testing.expectEqualStrings("TableColumnNotFound", std.mem.sliceTo(&err_buf, 0));
    try std.testing.expectEqual(@as(u8, 0), diag.error_name[0]);
    const price = "Price";
    try std.testing.expectEqual(ZLSX_ERROR, zlsx_editor_rename_table_column(ed, tbl.ptr, tbl.len, price.ptr, price.len, null, 0, &diag, &err_buf, err_buf.len));
    try std.testing.expectEqualStrings("InvalidTableColumnName", std.mem.sliceTo(&err_buf, 0));
    try std.testing.expectEqual(plane_none, diag.plane);
    zlsx_diag_release(&diag);
    try std.testing.expectEqual(@as(i32, 0), zlsx_editor_save(ed, out_path.ptr, out_path.len, &err_buf, err_buf.len));

    const table_xml = try savedPartBytes(alloc, io, out_path, "xl/tables/table1.xml");
    defer alloc.free(table_xml);
    try std.testing.expect(std.mem.indexOf(u8, table_xml, "name=\"Quantity\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, table_xml, "name=\"Qty\"") == null);
}

test "S3a refusals: -2, the error name in the diag, no plane, errbuf agrees" {
    const alloc = std.testing.allocator;
    var threaded: std.Io.Threaded = .init(alloc, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(alloc, io, "s3a_pivot.xlsx");
    defer alloc.free(path);
    try zlsx_pkg.pivots.fixture.write(alloc, io, path, .table_name);

    var err_buf: [128]u8 = undefined;
    const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
    defer zlsx_editor_close(ed);

    const cases = [_]struct { name: []const u8, status: i32 }{
        .{ .name = "RowEditUnsafeForSheet", .status = 0 },
        .{ .name = "ColEditUnsafeForSheet", .status = 0 },
        .{ .name = "DuplicateSheetName", .status = 0 },
        .{ .name = "TableColumnNameInUse", .status = 0 },
    };
    for (cases) |c| {
        var diag = std.mem.zeroes(CDiag);
        diag.struct_size = @sizeOf(CDiag);
        const status: i32 = if (std.mem.eql(u8, c.name, "RowEditUnsafeForSheet"))
            // Row 3 of `Report` is inside the pivot's A3:B6 footprint.
            zlsx_editor_delete_row(ed, 1, 3, &diag, &err_buf, err_buf.len)
        else if (std.mem.eql(u8, c.name, "ColEditUnsafeForSheet"))
            // Column A of `Report` is inside the footprint too (an insert
            // *before* it would be S7a's lift, not a refusal).
            zlsx_editor_delete_column(ed, 1, 0, &diag, &err_buf, err_buf.len)
        else if (std.mem.eql(u8, c.name, "DuplicateSheetName")) blk: {
            const dup = "report";
            break :blk zlsx_editor_add_sheet(ed, dup.ptr, dup.len, null, &diag, &err_buf, err_buf.len);
        } else blk: {
            const t = "SalesTbl";
            const o = "Qty";
            const n = "Price";
            break :blk zlsx_editor_rename_table_column(ed, t.ptr, t.len, o.ptr, o.len, n.ptr, n.len, &diag, &err_buf, err_buf.len);
        };
        try std.testing.expectEqual(ZLSX_REFUSED, status);
        try std.testing.expectEqualStrings(c.name, diagName(&diag));
        try std.testing.expectEqualStrings(c.name, std.mem.sliceTo(&err_buf, 0));
        try std.testing.expectEqual(plane_none, diag.plane);
        try std.testing.expectEqual(@as(usize, 0), diag.census_len);
        zlsx_diag_release(&diag);
    }

    // Rename to a name another sheet holds, case-folded.
    {
        var diag = freshDiag();
        const dup = "DATA";
        try std.testing.expectEqual(ZLSX_REFUSED, zlsx_editor_rename_sheet(ed, 1, dup.ptr, dup.len, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings("DuplicateSheetName", diagName(&diag));
    }
    // The last sheet cannot go: delete the two-sheet fixture down to one first.
    {
        var diag = freshDiag();
        try std.testing.expectEqual(ZLSX_OK, zlsx_editor_delete_sheet(ed, 1, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqual(ZLSX_REFUSED, zlsx_editor_delete_sheet(ed, 0, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings("CannotDeleteLastSheet", diagName(&diag));
        try std.testing.expectEqual(plane_none, diag.plane);
    }
}

test "S3a contract violations: -1 with the name in errbuf, the diag untouched or zeroed, canary tails intact" {
    const alloc = std.testing.allocator;
    var threaded: std.Io.Threaded = .init(alloc, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try writeS3aFixture(io, &tt, "s3a_contract.xlsx");
    defer alloc.free(path);

    var err_buf: [128]u8 = undefined;
    const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
    defer zlsx_editor_close(ed);

    var diag = freshDiag();
    try std.testing.expectEqual(ZLSX_ERROR, zlsx_editor_insert_row(ed, 9, 1, &diag, &err_buf, err_buf.len));
    try std.testing.expectEqualStrings("SheetIndexOutOfRange", std.mem.sliceTo(&err_buf, 0));
    try std.testing.expectEqual(ZLSX_ERROR, zlsx_editor_delete_row(ed, 0, 0, &diag, &err_buf, err_buf.len));
    try std.testing.expectEqualStrings("RowIndexOutOfRange", std.mem.sliceTo(&err_buf, 0));
    try std.testing.expectEqual(ZLSX_ERROR, zlsx_editor_insert_column(ed, 0, std.math.maxInt(u32), &diag, &err_buf, err_buf.len));
    try std.testing.expectEqualStrings("ColumnIndexOutOfRange", std.mem.sliceTo(&err_buf, 0));
    try std.testing.expectEqual(ZLSX_ERROR, zlsx_editor_delete_column(ed, 0, 16384, &diag, &err_buf, err_buf.len));
    try std.testing.expectEqualStrings("ColumnIndexOutOfRange", std.mem.sliceTo(&err_buf, 0));
    try std.testing.expectEqual(ZLSX_ERROR, zlsx_editor_insert_row(null, 0, 1, &diag, &err_buf, err_buf.len));
    try std.testing.expectEqualStrings("InvalidInput", std.mem.sliceTo(&err_buf, 0));
    // A refusal-free failure leaves the diag as prep left it: no name.
    try std.testing.expectEqual(@as(u8, 0), diag.error_name[0]);
    try std.testing.expectEqual(plane_none, diag.plane);

    // Names: NULL with a length is about the call; the empty name is
    // the editor's verdict (Excel would not take it).
    var idx: u32 = 3;
    try std.testing.expectEqual(ZLSX_ERROR, zlsx_editor_add_sheet(ed, null, 3, &idx, &diag, &err_buf, err_buf.len));
    try std.testing.expectEqualStrings("InvalidInput", std.mem.sliceTo(&err_buf, 0));
    try std.testing.expectEqual(no_sheet_idx, idx);
    try std.testing.expectEqual(ZLSX_ERROR, zlsx_editor_add_sheet(ed, null, 0, &idx, &diag, &err_buf, err_buf.len));
    try std.testing.expectEqualStrings("InvalidSheetName", std.mem.sliceTo(&err_buf, 0));
    const bad = "a:b";
    try std.testing.expectEqual(ZLSX_ERROR, zlsx_editor_rename_sheet(ed, 0, bad.ptr, bad.len, &diag, &err_buf, err_buf.len));
    try std.testing.expectEqualStrings("InvalidSheetName", std.mem.sliceTo(&err_buf, 0));
    const t = "T";
    try std.testing.expectEqual(ZLSX_ERROR, zlsx_editor_rename_table_column(ed, t.ptr, t.len, null, 1, t.ptr, t.len, &diag, &err_buf, err_buf.len));
    try std.testing.expectEqualStrings("InvalidInput", std.mem.sliceTo(&err_buf, 0));
    // A table the workbook does not have is a selector — the
    // table-shaped SheetIndexOutOfRange.
    try std.testing.expectEqual(ZLSX_ERROR, zlsx_editor_rename_table_column(ed, t.ptr, t.len, t.ptr, t.len, bad.ptr, bad.len, &diag, &err_buf, err_buf.len));
    try std.testing.expectEqualStrings("TableNotFound", std.mem.sliceTo(&err_buf, 0));
    try std.testing.expectEqual(@as(u8, 0), diag.error_name[0]);

    // A staged cell write makes the sheet unclean for a structural
    // edit: a sequencing statement, -1, never a refusal.
    const cell = toCCell(.{ .integer = 42 });
    try std.testing.expectEqual(@as(i32, 0), zlsx_editor_set_cell(ed, 0, 1, 0, &cell, &err_buf, err_buf.len));
    try std.testing.expectEqual(ZLSX_ERROR, zlsx_editor_insert_row(ed, 0, 1, &diag, &err_buf, err_buf.len));
    try std.testing.expectEqualStrings("RowEditRequiresCleanSheet", std.mem.sliceTo(&err_buf, 0));
    try std.testing.expectEqual(ZLSX_ERROR, zlsx_editor_insert_column(ed, 0, 0, &diag, &err_buf, err_buf.len));
    try std.testing.expectEqualStrings("ColEditRequiresCleanSheet", std.mem.sliceTo(&err_buf, 0));
    try std.testing.expectEqual(ZLSX_ERROR, zlsx_editor_delete_sheet(ed, 1, &diag, &err_buf, err_buf.len));
    try std.testing.expectEqualStrings("SheetDeleteRequiresCleanState", std.mem.sliceTo(&err_buf, 0));

    // struct_size below v1: rejected before a byte is written.
    var small = std.mem.zeroes(CDiag);
    small.struct_size = @sizeOf(CDiag) - 1;
    small.plane = 7;
    small.error_name[0] = 'x';
    try std.testing.expectEqual(ZLSX_ERROR, zlsx_editor_delete_sheet(ed, 1, &small, &err_buf, err_buf.len));
    try std.testing.expectEqualStrings("StructSizeTooSmall", std.mem.sliceTo(&err_buf, 0));
    try std.testing.expectEqual(@as(u32, 7), small.plane);
    try std.testing.expectEqual(@as(u8, 'x'), small.error_name[0]);

    // Canary tail: a caller compiled against a larger v-next struct
    // keeps its tail across a refusal and across a generic failure.
    var big: [@sizeOf(CDiag) + 64]u8 align(@alignOf(CDiag)) = undefined;
    @memset(&big, 0xAA);
    const bd: *CDiag = @ptrCast(&big);
    bd.struct_size = big.len;
    try std.testing.expectEqual(ZLSX_ERROR, zlsx_editor_insert_row(ed, 9, 1, bd, &err_buf, err_buf.len));
    for (big[@sizeOf(CDiag)..]) |b| try std.testing.expectEqual(@as(u8, 0xAA), b);
    try std.testing.expectEqual(plane_none, bd.plane);
    zlsx_diag_release(bd);
    for (big[@sizeOf(CDiag)..]) |b| try std.testing.expectEqual(@as(u8, 0xAA), b);
    const dup = "second";
    try std.testing.expectEqual(ZLSX_REFUSED, zlsx_editor_add_sheet(ed, dup.ptr, dup.len, null, bd, &err_buf, err_buf.len));
    try std.testing.expectEqualStrings("DuplicateSheetName", diagName(bd));
    try std.testing.expectEqual(plane_none, bd.plane);
    for (big[@sizeOf(CDiag)..]) |b| try std.testing.expectEqual(@as(u8, 0xAA), b);
    zlsx_diag_release(bd);
    for (big[@sizeOf(CDiag)..]) |b| try std.testing.expectEqual(@as(u8, 0xAA), b);
}

test "S3a pivots_ndjson: the package writer's bytes, empty on a plain workbook, refused on a broken graph" {
    const alloc = std.testing.allocator;
    var threaded: std.Io.Threaded = .init(alloc, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    var err_buf: [128]u8 = undefined;

    // The frozen record — the literal the CLI test and the package
    // test both pin.
    {
        const path = try tt.path(alloc, io, "s3a_pivots.xlsx");
        defer alloc.free(path);
        try zlsx_pkg.pivots.fixture.writeWithOrphanCache(alloc, io, path, .sheet_ref);
        const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
        defer zlsx_editor_close(ed);
        var diag = freshDiag();
        var out_ptr: ?[*]u8 = null;
        var out_len: usize = 0;
        try std.testing.expectEqual(ZLSX_OK, zlsx_editor_pivots_ndjson(ed, &out_ptr, &out_len, &diag, &err_buf, err_buf.len));
        const got = out_ptr.?[0..out_len];
        try std.testing.expect(std.mem.startsWith(u8, got, zlsx_pkg.pivots.ndjson.fixture_sheet_ref_record));
        try std.testing.expect(std.mem.indexOf(u8, got, "{\"kind\":\"pivot_cache\",\"cache\":{\"id\":") != null);
        try std.testing.expectEqual(@as(usize, 2), std.mem.count(u8, got, "\n"));
        zlsx_buffer_release(out_ptr, out_len);

        // A staged edit is what the read sees: rename the host sheet
        // and the record names it.
        const nm = "Host";
        try std.testing.expectEqual(ZLSX_OK, zlsx_editor_rename_sheet(ed, 1, nm.ptr, nm.len, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqual(ZLSX_OK, zlsx_editor_pivots_ndjson(ed, &out_ptr, &out_len, &diag, &err_buf, err_buf.len));
        try std.testing.expect(std.mem.startsWith(u8, out_ptr.?[0..out_len], "{\"kind\":\"pivot\",\"sheet\":\"Host\",\"sheet_idx\":1,"));
        zlsx_buffer_release(out_ptr, out_len);

        // NULL out pointers are about the call.
        try std.testing.expectEqual(ZLSX_ERROR, zlsx_editor_pivots_ndjson(ed, null, &out_len, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings("NullOutPointer", std.mem.sliceTo(&err_buf, 0));
        try std.testing.expectEqual(@as(usize, 0), out_len);
        try std.testing.expectEqual(ZLSX_ERROR, zlsx_editor_pivots_ndjson(null, &out_ptr, &out_len, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings("InvalidInput", std.mem.sliceTo(&err_buf, 0));
        try std.testing.expect(out_ptr == null);
    }
    // The timing the contract states (r7 REL-701): a staged cell write
    // inside the source is invisible to the read until save; the saved
    // workbook reopens with the cache rebuilt and marked.
    {
        const path = try tt.path(alloc, io, "s3a_pivots_staged.xlsx");
        defer alloc.free(path);
        try zlsx_pkg.pivots.fixture.write(alloc, io, path, .sheet_ref);
        const out_path = try tt.path(alloc, io, "s3a_pivots_staged_out.xlsx");
        defer alloc.free(out_path);
        const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
        defer zlsx_editor_close(ed);
        // B2 (Qty of the first record) 3 → 9: inside the source rect.
        const cell = toCCell(.{ .integer = 9 });
        try std.testing.expectEqual(@as(i32, 0), zlsx_editor_set_cell(ed, 0, 2, 1, &cell, &err_buf, err_buf.len));
        var diag = freshDiag();
        var out_ptr: ?[*]u8 = null;
        var out_len: usize = 0;
        try std.testing.expectEqual(ZLSX_OK, zlsx_editor_pivots_ndjson(ed, &out_ptr, &out_len, &diag, &err_buf, err_buf.len));
        // Byte-for-byte the pre-write record: the staged delta has not
        // reached the graph.
        try std.testing.expectEqualStrings(zlsx_pkg.pivots.ndjson.fixture_sheet_ref_record, out_ptr.?[0..out_len]);
        zlsx_buffer_release(out_ptr, out_len);
        try std.testing.expectEqual(@as(i32, 0), zlsx_editor_save(ed, out_path.ptr, out_path.len, &err_buf, err_buf.len));
        const ed2 = zlsx_editor_open(out_path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
        defer zlsx_editor_close(ed2);
        out_ptr = null;
        out_len = 0;
        try std.testing.expectEqual(ZLSX_OK, zlsx_editor_pivots_ndjson(ed2, &out_ptr, &out_len, &diag, &err_buf, err_buf.len));
        defer zlsx_buffer_release(out_ptr, out_len);
        const got = out_ptr.?[0..out_len];
        // The save refreshed the cache: marked, and the write visible
        // in the rebuilt inventory (Qty max 5 → 9).
        try std.testing.expect(std.mem.indexOf(u8, got, "\"refresh_on_load\":true") != null);
        try std.testing.expect(std.mem.indexOf(u8, got, "\"max\":\"9\"") != null);
    }
    // No pivots: success, nothing to release.
    {
        const path = try writeS3aFixture(io, &tt, "s3a_no_pivots.xlsx");
        defer alloc.free(path);
        const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
        defer zlsx_editor_close(ed);
        var out_ptr: ?[*]u8 = @ptrFromInt(0x1000);
        var out_len: usize = 99;
        try std.testing.expectEqual(ZLSX_OK, zlsx_editor_pivots_ndjson(ed, &out_ptr, &out_len, null, &err_buf, err_buf.len));
        try std.testing.expect(out_ptr == null);
        try std.testing.expectEqual(@as(usize, 0), out_len);
        zlsx_buffer_release(out_ptr, out_len);
    }
    // A graph that cannot be read whole: -2, named, nothing handed out.
    {
        const path = try tt.path(alloc, io, "s3a_pivots_broken.xlsx");
        defer alloc.free(path);
        try zlsx_pkg.pivots.fixture.write(alloc, io, path, .sheet_ref);
        try zlsx_pkg.pivots.fixture.patchPart(alloc, io, path, "xl/pivotTables/pivotTable1.xml", "<location ref=\"A3:B6\" firstHeaderRow=\"1\" firstDataRow=\"1\" firstDataCol=\"1\"/>", "");
        const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
        defer zlsx_editor_close(ed);
        var diag = freshDiag();
        var out_ptr: ?[*]u8 = null;
        var out_len: usize = 0;
        try std.testing.expectEqual(ZLSX_REFUSED, zlsx_editor_pivots_ndjson(ed, &out_ptr, &out_len, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings("MalformedPivotXml", diagName(&diag));
        try std.testing.expectEqual(plane_none, diag.plane);
        try std.testing.expect(out_ptr == null);
        try std.testing.expectEqual(@as(usize, 0), out_len);
    }
}

test "S3a refusals from the transform and the sweeps: the precise name crosses as -2" {
    const alloc = std.testing.allocator;
    var threaded: std.Io.Threaded = .init(alloc, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    var err_buf: [128]u8 = undefined;
    const sheet_part = "xl/worksheets/sheet1.xml";

    // A cell past the grid: the transform's pre-mutation probe refuses
    // the insert with its own name. (A pivot-free workbook: on a pivot
    // SOURCE sheet the S7b source read meets the strict parser first and
    // the verdict is `MalformedSheetXml` — also -2, a different name.)
    {
        const path = try writeS3aFixture(io, &tt, "s3a_maxrow.xlsx");
        defer alloc.free(path);
        try zlsx_pkg.pivots.fixture.patchPart(alloc, io, path, sheet_part, "<c r=\"B2\"><v>5</v></c>", "<c r=\"B1048577\"><v>5</v></c>");
        const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
        defer zlsx_editor_close(ed);
        var diag = freshDiag();
        try std.testing.expectEqual(ZLSX_REFUSED, zlsx_editor_insert_row(ed, 0, 2, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings("RowEditExceedsMaxRow", diagName(&diag));
        try std.testing.expectEqualStrings("RowEditExceedsMaxRow", std.mem.sliceTo(&err_buf, 0));
        try std.testing.expectEqual(plane_none, diag.plane);
    }
    // A split pane (pixel offsets) the rewriter does not shift.
    {
        const path = try writeS3aFixture(io, &tt, "s3a_splitpane.xlsx");
        defer alloc.free(path);
        try zlsx_pkg.pivots.fixture.patchPart(alloc, io, path, sheet_part, "<sheetData>", "<sheetViews><sheetView workbookViewId=\"0\"><pane ySplit=\"2400\" topLeftCell=\"A5\" activePane=\"bottomLeft\" state=\"split\"/></sheetView></sheetViews><sheetData>");
        const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
        defer zlsx_editor_close(ed);
        var diag = freshDiag();
        try std.testing.expectEqual(ZLSX_REFUSED, zlsx_editor_delete_column(ed, 0, 0, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings("SplitPaneNotSupported", diagName(&diag));
        try std.testing.expectEqual(plane_none, diag.plane);
        // Refused before any mutation: the sheet part is untouched.
        var out_ptr: ?[*]u8 = null;
        var out_len: usize = 0;
        try std.testing.expectEqual(ZLSX_OK, zlsx_editor_save_to_buffer(ed, &out_ptr, &out_len, &err_buf, err_buf.len));
        defer zlsx_buffer_release(out_ptr, out_len);
        const src_bytes = try readFileBytes(io, path);
        defer alloc.free(src_bytes);
        try std.testing.expectEqualSlices(u8, src_bytes, out_ptr.?[0..out_len]);
    }
}

fn pivotsNdjsonForFailures(alloc: std.mem.Allocator, wb: *zlsx_pkg.Workbook) !void {
    const bytes = try pivotsNdjsonOwned(alloc, wb);
    alloc.free(bytes);
}

test "S3a pivots_ndjson: every allocation failure while writing is OutOfMemory, never WriteFailed" {
    const alloc = std.testing.allocator;
    var threaded: std.Io.Threaded = .init(alloc, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(alloc, io, "s3a_pivots_oom.xlsx");
    defer alloc.free(path);
    try zlsx_pkg.pivots.fixture.writeWithOrphanCache(alloc, io, path, .sheet_ref);
    var wb = try zlsx_pkg.Workbook.open(alloc, io, path);
    defer wb.deinit();
    try std.testing.checkAllAllocationFailures(alloc, pivotsNdjsonForFailures, .{&wb});
    try std.testing.expectEqual(ZLSX_NOMEM, statusOf(error.OutOfMemory));
}

/// Two sheets and three defined names — workbook scope, sheet scope,
/// hidden — written through the C surface so the read-back crosses the
/// same boundary a binding does.
fn writeS3bDefinedNamesFixture(io: std.Io, tt: *TestTmp, name: []const u8) ![:0]u8 {
    const alloc = std.testing.allocator;
    const path = try tt.path(alloc, io, name);
    errdefer alloc.free(path);
    var err_buf: [128]u8 = undefined;
    const w = zlsx_writer_create(&err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
    defer zlsx_writer_close(w);
    const s1 = "Data";
    const sw1 = zlsx_writer_add_sheet(w, s1.ptr, s1.len, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
    const cells1 = [_]CCell{ toCCell(.{ .integer = 1 }), toCCell(.{ .integer = 2 }) };
    if (zlsx_sheet_writer_write_row(sw1, &cells1, cells1.len, &err_buf, err_buf.len) != 0)
        return error.TestUnexpectedResult;
    const s2 = "Second";
    const sw2 = zlsx_writer_add_sheet(w, s2.ptr, s2.len, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
    const cells2 = [_]CCell{toCCell(.{ .integer = 3 })};
    if (zlsx_sheet_writer_write_row(sw2, &cells2, cells2.len, &err_buf, err_buf.len) != 0)
        return error.TestUnexpectedResult;
    const n1 = "Prices";
    const b1 = "Data!$A$1:$C$4";
    if (zlsx_writer_add_defined_name(w, n1.ptr, n1.len, b1.ptr, b1.len, -1, 0, &err_buf, err_buf.len) != 0)
        return error.TestUnexpectedResult;
    const n2 = "_xlnm.Print_Area";
    const b2 = "Second!$A$1:$B$9";
    if (zlsx_writer_add_defined_name(w, n2.ptr, n2.len, b2.ptr, b2.len, 1, 0, &err_buf, err_buf.len) != 0)
        return error.TestUnexpectedResult;
    const n3 = "Secret";
    const b3 = "Data!$Z$1";
    if (zlsx_writer_add_defined_name(w, n3.ptr, n3.len, b3.ptr, b3.len, -1, 1, &err_buf, err_buf.len) != 0)
        return error.TestUnexpectedResult;
    if (zlsx_writer_save(w, path.ptr, path.len, &err_buf, err_buf.len) != 0)
        return error.TestUnexpectedResult;
    return path;
}

test "S3b defined_names_ndjson: the shared writer's bytes, current after a rename, empty without names" {
    const alloc = std.testing.allocator;
    var threaded: std.Io.Threaded = .init(alloc, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    var err_buf: [128]u8 = undefined;

    {
        const path = try writeS3bDefinedNamesFixture(io, &tt, "s3b_defined_names.xlsx");
        defer alloc.free(path);
        const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
        defer zlsx_editor_close(ed);
        var diag = freshDiag();
        var out_ptr: ?[*]u8 = null;
        var out_len: usize = 0;
        try std.testing.expectEqual(ZLSX_OK, zlsx_editor_defined_names_ndjson(ed, &out_ptr, &out_len, &diag, &err_buf, err_buf.len));
        // The frozen record — the literal the CLI test and the package
        // test both pin (docs/cli.md, "defined-names").
        const expected =
            "{\"kind\":\"defined_name\",\"name\":\"Prices\",\"scope\":\"workbook\",\"sheet\":null,\"sheet_idx\":null,\"body\":\"Data!$A$1:$C$4\",\"hidden\":false}\n" ++
            "{\"kind\":\"defined_name\",\"name\":\"_xlnm.Print_Area\",\"scope\":\"sheet\",\"sheet\":\"Second\",\"sheet_idx\":1,\"body\":\"Second!$A$1:$B$9\",\"hidden\":false}\n" ++
            "{\"kind\":\"defined_name\",\"name\":\"Secret\",\"scope\":\"workbook\",\"sheet\":null,\"sheet_idx\":null,\"body\":\"Data!$Z$1\",\"hidden\":true}\n";
        try std.testing.expectEqualStrings(expected, out_ptr.?[0..out_len]);
        zlsx_buffer_release(out_ptr, out_len);

        // A structural edit is what the read sees: renaming the sheet
        // the bodies reference shows the name sweep's rewrite and the
        // refreshed view, no save in between.
        const nm = "Facts";
        try std.testing.expectEqual(ZLSX_OK, zlsx_editor_rename_sheet(ed, 0, nm.ptr, nm.len, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqual(ZLSX_OK, zlsx_editor_defined_names_ndjson(ed, &out_ptr, &out_len, &diag, &err_buf, err_buf.len));
        const got = out_ptr.?[0..out_len];
        try std.testing.expect(std.mem.indexOf(u8, got, "\"body\":\"Facts!$A$1:$C$4\"") != null);
        try std.testing.expect(std.mem.indexOf(u8, got, "\"body\":\"Facts!$Z$1\"") != null);
        try std.testing.expect(std.mem.indexOf(u8, got, "\"sheet\":\"Second\",\"sheet_idx\":1,\"body\":\"Second!$A$1:$B$9\"") != null);
        zlsx_buffer_release(out_ptr, out_len);

        // NULL out pointers are about the call — either one, with the
        // present one reset from its poison.
        try std.testing.expectEqual(ZLSX_ERROR, zlsx_editor_defined_names_ndjson(ed, null, &out_len, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings("NullOutPointer", std.mem.sliceTo(&err_buf, 0));
        try std.testing.expectEqual(@as(usize, 0), out_len);
        out_ptr = @ptrFromInt(0x1000);
        try std.testing.expectEqual(ZLSX_ERROR, zlsx_editor_defined_names_ndjson(ed, &out_ptr, null, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings("NullOutPointer", std.mem.sliceTo(&err_buf, 0));
        try std.testing.expect(out_ptr == null);
        // Re-poisoned: the NULL-editor path's own reset is what this pins.
        out_ptr = @ptrFromInt(0x1000);
        out_len = 99;
        try std.testing.expectEqual(ZLSX_ERROR, zlsx_editor_defined_names_ndjson(null, &out_ptr, &out_len, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings("InvalidInput", std.mem.sliceTo(&err_buf, 0));
        try std.testing.expect(out_ptr == null);
        try std.testing.expectEqual(@as(usize, 0), out_len);

        // struct_size below v1: the outputs reset first, then the diag
        // is rejected before a byte of it is written — the whole struct
        // compared, not a field or two.
        var small = std.mem.zeroes(CDiag);
        small.struct_size = @sizeOf(CDiag) - 1;
        small.plane = 7;
        small.error_name[0] = 'x';
        const small_before = std.mem.toBytes(small);
        out_ptr = @ptrFromInt(0x1000);
        out_len = 99;
        try std.testing.expectEqual(ZLSX_ERROR, zlsx_editor_defined_names_ndjson(ed, &out_ptr, &out_len, &small, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings("StructSizeTooSmall", std.mem.sliceTo(&err_buf, 0));
        try std.testing.expect(out_ptr == null);
        try std.testing.expectEqual(@as(usize, 0), out_len);
        const small_after = std.mem.toBytes(small);
        try std.testing.expectEqualSlices(u8, &small_before, &small_after);
    }
    // No defined names: success, nothing to release, the poison reset.
    {
        const path = try writeS3aFixture(io, &tt, "s3b_no_names.xlsx");
        defer alloc.free(path);
        const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
        defer zlsx_editor_close(ed);
        var out_ptr: ?[*]u8 = @ptrFromInt(0x1000);
        var out_len: usize = 99;
        try std.testing.expectEqual(ZLSX_OK, zlsx_editor_defined_names_ndjson(ed, &out_ptr, &out_len, null, &err_buf, err_buf.len));
        try std.testing.expect(out_ptr == null);
        try std.testing.expectEqual(@as(usize, 0), out_len);
        zlsx_buffer_release(out_ptr, out_len);
    }
}

test "S3b defined_names_ndjson: an inventory the read cannot serve refuses whole, nothing handed out" {
    const alloc = std.testing.allocator;
    var threaded: std.Io.Threaded = .init(alloc, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    var err_buf: [128]u8 = undefined;
    const path = try writeS3bDefinedNamesFixture(io, &tt, "s3b_names_broken.xlsx");
    defer alloc.free(path);
    // A bad entity in one body: the open parser keeps raw spans, so the
    // editor opens; the decode at read time refuses the whole view.
    try zlsx_pkg.pivots.fixture.patchPart(alloc, io, path, "xl/workbook.xml", "Data!$Z$1", "Data!$Z$1&bogus;");
    const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
    defer zlsx_editor_close(ed);
    var diag = freshDiag();
    // Poisoned on entry: the refusal path itself must reset the pair.
    var out_ptr: ?[*]u8 = @ptrFromInt(0x1000);
    var out_len: usize = 99;
    try std.testing.expectEqual(ZLSX_REFUSED, zlsx_editor_defined_names_ndjson(ed, &out_ptr, &out_len, &diag, &err_buf, err_buf.len));
    try std.testing.expectEqualStrings("MalformedWorkbookXml", diagName(&diag));
    try std.testing.expectEqualStrings("MalformedWorkbookXml", std.mem.sliceTo(&err_buf, 0));
    try std.testing.expectEqual(plane_none, diag.plane);
    try std.testing.expect(out_ptr == null);
    try std.testing.expectEqual(@as(usize, 0), out_len);
}

fn definedNamesNdjsonForFailures(alloc: std.mem.Allocator, wb: *zlsx_pkg.Workbook) !void {
    const bytes = try definedNamesNdjsonOwned(alloc, wb);
    alloc.free(bytes);
}

test "S3b defined_names_ndjson: every allocation failure is OutOfMemory, never WriteFailed" {
    const alloc = std.testing.allocator;
    var threaded: std.Io.Threaded = .init(alloc, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try writeS3bDefinedNamesFixture(io, &tt, "s3b_names_oom.xlsx");
    defer alloc.free(path);
    var wb = try zlsx_pkg.Workbook.open(alloc, io, path);
    defer wb.deinit();
    try std.testing.checkAllAllocationFailures(alloc, definedNamesNdjsonForFailures, .{&wb});
}

/// The pkg fixture (`conditional_format_ndjson.zig::fixture.write`)
/// rebuilt through the C writer surface, so the frozen stream below is
/// pinned across the boundary a binding crosses: one dxf, four rule
/// kinds on `Data`, an escaping-heavy expression on `Report`.
fn writeS3bConditionalFormatsFixture(io: std.Io, tt: *TestTmp, name: []const u8) ![:0]u8 {
    const alloc = std.testing.allocator;
    const path = try tt.path(alloc, io, name);
    errdefer alloc.free(path);
    var err_buf: [128]u8 = undefined;
    const w = zlsx_writer_create(&err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
    defer zlsx_writer_close(w);
    var dxf = std.mem.zeroes(CDxf);
    dxf.bold = 1;
    dxf.has_fill = 1;
    dxf.fill_fg_argb = 0xFFFFC7CE;
    var dxf_id: u32 = 99;
    if (zlsx_writer_add_dxf(w, &dxf, &dxf_id, &err_buf, err_buf.len) != 0)
        return error.TestUnexpectedResult;
    const s1 = "Data";
    const sw1 = zlsx_writer_add_sheet(w, s1.ptr, s1.len, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
    const row1 = [_]CCell{ toCCell(.{ .integer = 1 }), toCCell(.{ .integer = 5 }), toCCell(.{ .integer = 9 }), toCCell(.{ .integer = 3 }) };
    if (zlsx_sheet_writer_write_row(sw1, &row1, row1.len, &err_buf, err_buf.len) != 0)
        return error.TestUnexpectedResult;
    const row2 = [_]CCell{ toCCell(.{ .integer = 2 }), toCCell(.{ .integer = 6 }), toCCell(.{ .integer = 10 }), toCCell(.{ .integer = 4 }) };
    if (zlsx_sheet_writer_write_row(sw1, &row2, row2.len, &err_buf, err_buf.len) != 0)
        return error.TestUnexpectedResult;
    const r1 = "A1:A4";
    const f1 = "2";
    const f2 = "4";
    if (zlsx_sheet_writer_add_conditional_format_cell_is(sw1, r1.ptr, r1.len, ZLSX_DV_OP_BETWEEN, f1.ptr, f1.len, f2.ptr, f2.len, dxf_id, &err_buf, err_buf.len) != 0)
        return error.TestUnexpectedResult;
    const r2 = "B1:B4";
    const fe = "B1>3";
    if (zlsx_sheet_writer_add_conditional_format_expression(sw1, r2.ptr, r2.len, fe.ptr, fe.len, dxf_id, &err_buf, err_buf.len) != 0)
        return error.TestUnexpectedResult;
    const r3 = "C1:C4";
    if (zlsx_sheet_writer_add_conditional_format_color_scale(sw1, r3.ptr, r3.len, 0xFFF8696B, 0, 0, 0xFF63BE7B, &err_buf, err_buf.len) != 0)
        return error.TestUnexpectedResult;
    const r4 = "D1:D4";
    if (zlsx_sheet_writer_add_conditional_format_data_bar(sw1, r4.ptr, r4.len, 0xFF638EC6, &err_buf, err_buf.len) != 0)
        return error.TestUnexpectedResult;
    const s2 = "Report";
    const sw2 = zlsx_writer_add_sheet(w, s2.ptr, s2.len, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
    const cells2 = [_]CCell{toCCell(.{ .string = "R&D" })};
    if (zlsx_sheet_writer_write_row(sw2, &cells2, cells2.len, &err_buf, err_buf.len) != 0)
        return error.TestUnexpectedResult;
    const r5 = "A1:A2";
    const fr = "$A1=\"R&D\"";
    if (zlsx_sheet_writer_add_conditional_format_expression(sw2, r5.ptr, r5.len, fr.ptr, fr.len, dxf_id, &err_buf, err_buf.len) != 0)
        return error.TestUnexpectedResult;
    if (zlsx_writer_save(w, path.ptr, path.len, &err_buf, err_buf.len) != 0)
        return error.TestUnexpectedResult;
    return path;
}

test "S3b conditional_formats_ndjson: the shared writer's bytes, current after a rename, empty without rules" {
    const alloc = std.testing.allocator;
    var threaded: std.Io.Threaded = .init(alloc, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    var err_buf: [128]u8 = undefined;

    {
        const path = try writeS3bConditionalFormatsFixture(io, &tt, "s3b_cf.xlsx");
        defer alloc.free(path);
        const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
        defer zlsx_editor_close(ed);
        var diag = freshDiag();
        var out_ptr: ?[*]u8 = null;
        var out_len: usize = 0;
        try std.testing.expectEqual(ZLSX_OK, zlsx_editor_conditional_formats_ndjson(ed, &out_ptr, &out_len, &diag, &err_buf, err_buf.len));
        // The frozen stream — the literal the shared writer's own test
        // pins over the same fixture (MNT-2302; docs/cli.md,
        // "conditional-formats").
        const expected =
            "{\"kind\":\"conditional_format\",\"sheet\":\"Data\",\"sheet_idx\":0,\"sqref\":\"A1:A4\"," ++
            "\"rule_type\":\"cellIs\",\"formulas\":[\"2\",\"4\"],\"dxf_id\":0,\"priority\":1}\n" ++
            "{\"kind\":\"conditional_format\",\"sheet\":\"Data\",\"sheet_idx\":0,\"sqref\":\"B1:B4\"," ++
            "\"rule_type\":\"expression\",\"formulas\":[\"B1>3\"],\"dxf_id\":0,\"priority\":2}\n" ++
            "{\"kind\":\"conditional_format\",\"sheet\":\"Data\",\"sheet_idx\":0,\"sqref\":\"C1:C4\"," ++
            "\"rule_type\":\"colorScale\",\"formulas\":[],\"dxf_id\":null,\"priority\":3}\n" ++
            "{\"kind\":\"conditional_format\",\"sheet\":\"Data\",\"sheet_idx\":0,\"sqref\":\"D1:D4\"," ++
            "\"rule_type\":\"dataBar\",\"formulas\":[],\"dxf_id\":null,\"priority\":4}\n" ++
            "{\"kind\":\"conditional_format\",\"sheet\":\"Report\",\"sheet_idx\":1,\"sqref\":\"A1:A2\"," ++
            "\"rule_type\":\"expression\",\"formulas\":[\"$A1=\\\"R&D\\\"\"],\"dxf_id\":0,\"priority\":1}\n";
        try std.testing.expectEqualStrings(expected, out_ptr.?[0..out_len]);
        zlsx_buffer_release(out_ptr, out_len);

        // A structural edit is what the read sees: the sheet inventory
        // is re-read from the current workbook.xml, no save in between.
        const nm = "Facts";
        try std.testing.expectEqual(ZLSX_OK, zlsx_editor_rename_sheet(ed, 0, nm.ptr, nm.len, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqual(ZLSX_OK, zlsx_editor_conditional_formats_ndjson(ed, &out_ptr, &out_len, &diag, &err_buf, err_buf.len));
        const got = out_ptr.?[0..out_len];
        try std.testing.expect(std.mem.indexOf(u8, got, "\"sheet\":\"Facts\",\"sheet_idx\":0,\"sqref\":\"A1:A4\"") != null);
        try std.testing.expect(std.mem.indexOf(u8, got, "\"sheet\":\"Data\"") == null);
        try std.testing.expect(std.mem.indexOf(u8, got, "\"sheet\":\"Report\",\"sheet_idx\":1,\"sqref\":\"A1:A2\"") != null);
        zlsx_buffer_release(out_ptr, out_len);

        // A row insert moves the ENVELOPE with the bodies — sqref and
        // formula on one grid, no save in between (Codex #216 r1
        // S3B-REL-301: before the sqref shift, `B1>3` moved to `B2>3`
        // while `sqref="B1:B4"` stayed).
        try std.testing.expectEqual(ZLSX_OK, zlsx_editor_insert_row(ed, 0, 1, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqual(ZLSX_OK, zlsx_editor_conditional_formats_ndjson(ed, &out_ptr, &out_len, &diag, &err_buf, err_buf.len));
        const moved = out_ptr.?[0..out_len];
        try std.testing.expect(std.mem.indexOf(u8, moved, "\"sqref\":\"A2:A5\",\"rule_type\":\"cellIs\",\"formulas\":[\"2\",\"4\"]") != null);
        try std.testing.expect(std.mem.indexOf(u8, moved, "\"sqref\":\"B2:B5\",\"rule_type\":\"expression\",\"formulas\":[\"B2>3\"]") != null);
        try std.testing.expect(std.mem.indexOf(u8, moved, "\"sqref\":\"C2:C5\",\"rule_type\":\"colorScale\"") != null);
        try std.testing.expect(std.mem.indexOf(u8, moved, "\"sqref\":\"D2:D5\",\"rule_type\":\"dataBar\"") != null);
        // The other sheet's rules did not move.
        try std.testing.expect(std.mem.indexOf(u8, moved, "\"sheet\":\"Report\",\"sheet_idx\":1,\"sqref\":\"A1:A2\"") != null);
        zlsx_buffer_release(out_ptr, out_len);

        // NULL out pointers are about the call — either one, with the
        // present one reset from its poison.
        try std.testing.expectEqual(ZLSX_ERROR, zlsx_editor_conditional_formats_ndjson(ed, null, &out_len, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings("NullOutPointer", std.mem.sliceTo(&err_buf, 0));
        try std.testing.expectEqual(@as(usize, 0), out_len);
        out_ptr = @ptrFromInt(0x1000);
        try std.testing.expectEqual(ZLSX_ERROR, zlsx_editor_conditional_formats_ndjson(ed, &out_ptr, null, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings("NullOutPointer", std.mem.sliceTo(&err_buf, 0));
        try std.testing.expect(out_ptr == null);
        // Re-poisoned: the NULL-editor path's own reset is what this pins.
        out_ptr = @ptrFromInt(0x1000);
        out_len = 99;
        try std.testing.expectEqual(ZLSX_ERROR, zlsx_editor_conditional_formats_ndjson(null, &out_ptr, &out_len, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings("InvalidInput", std.mem.sliceTo(&err_buf, 0));
        try std.testing.expect(out_ptr == null);
        try std.testing.expectEqual(@as(usize, 0), out_len);

        // struct_size below v1: the outputs reset first, then the diag
        // is rejected before a byte of it is written — the whole struct
        // compared, not a field or two.
        var small = std.mem.zeroes(CDiag);
        small.struct_size = @sizeOf(CDiag) - 1;
        small.plane = 7;
        small.error_name[0] = 'x';
        const small_before = std.mem.toBytes(small);
        out_ptr = @ptrFromInt(0x1000);
        out_len = 99;
        try std.testing.expectEqual(ZLSX_ERROR, zlsx_editor_conditional_formats_ndjson(ed, &out_ptr, &out_len, &small, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings("StructSizeTooSmall", std.mem.sliceTo(&err_buf, 0));
        try std.testing.expect(out_ptr == null);
        try std.testing.expectEqual(@as(usize, 0), out_len);
        const small_after = std.mem.toBytes(small);
        try std.testing.expectEqualSlices(u8, &small_before, &small_after);
    }
    // No conditional formatting: success, nothing to release, the
    // poison reset.
    {
        const path = try writeS3aFixture(io, &tt, "s3b_no_cf.xlsx");
        defer alloc.free(path);
        const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
        defer zlsx_editor_close(ed);
        var out_ptr: ?[*]u8 = @ptrFromInt(0x1000);
        var out_len: usize = 99;
        try std.testing.expectEqual(ZLSX_OK, zlsx_editor_conditional_formats_ndjson(ed, &out_ptr, &out_len, null, &err_buf, err_buf.len));
        try std.testing.expect(out_ptr == null);
        try std.testing.expectEqual(@as(usize, 0), out_len);
        zlsx_buffer_release(out_ptr, out_len);
    }
}

test "S3b conditional_formats_ndjson: an inventory the read cannot serve refuses whole, nothing handed out" {
    const alloc = std.testing.allocator;
    var threaded: std.Io.Threaded = .init(alloc, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    var err_buf: [128]u8 = undefined;
    const path = try writeS3bConditionalFormatsFixture(io, &tt, "s3b_cf_broken.xlsx");
    defer alloc.free(path);
    // A bad entity in one formula body: the open parser keeps raw
    // spans, so the editor opens; the decode at read time refuses the
    // whole view — the sheet part's verdict, not the workbook's.
    try zlsx_pkg.pivots.fixture.patchPart(alloc, io, path, "xl/worksheets/sheet1.xml", "<formula>2</formula>", "<formula>2&bogus;</formula>");
    const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
    defer zlsx_editor_close(ed);
    var diag = freshDiag();
    // Poisoned on entry: the refusal path itself must reset the pair.
    var out_ptr: ?[*]u8 = @ptrFromInt(0x1000);
    var out_len: usize = 99;
    try std.testing.expectEqual(ZLSX_REFUSED, zlsx_editor_conditional_formats_ndjson(ed, &out_ptr, &out_len, &diag, &err_buf, err_buf.len));
    try std.testing.expectEqualStrings("MalformedSheetXml", diagName(&diag));
    try std.testing.expectEqualStrings("MalformedSheetXml", std.mem.sliceTo(&err_buf, 0));
    try std.testing.expectEqual(plane_none, diag.plane);
    try std.testing.expect(out_ptr == null);
    try std.testing.expectEqual(@as(usize, 0), out_len);
}

test "S3b conditional_formats_ndjson: a broken SECOND sheet refuses whole — no partial stream" {
    const alloc = std.testing.allocator;
    var threaded: std.Io.Threaded = .init(alloc, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    var err_buf: [128]u8 = undefined;
    const path = try writeS3bConditionalFormatsFixture(io, &tt, "s3b_cf_broken2.xlsx");
    defer alloc.free(path);
    // The FIRST sheet's four records are perfectly servable; the bad
    // entity sits in the second sheet's formula. An implementation
    // that streamed per sheet would hand out the first four before
    // failing — the whole-inventory rule forbids exactly that.
    try zlsx_pkg.pivots.fixture.patchPart(alloc, io, path, "xl/worksheets/sheet2.xml", "$A1=", "$A1&bogus;=");
    const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
    defer zlsx_editor_close(ed);
    var diag = freshDiag();
    var out_ptr: ?[*]u8 = @ptrFromInt(0x1000);
    var out_len: usize = 99;
    try std.testing.expectEqual(ZLSX_REFUSED, zlsx_editor_conditional_formats_ndjson(ed, &out_ptr, &out_len, &diag, &err_buf, err_buf.len));
    try std.testing.expectEqualStrings("MalformedSheetXml", diagName(&diag));
    try std.testing.expect(out_ptr == null);
    try std.testing.expectEqual(@as(usize, 0), out_len);
}

test "S3b conditional_formats_ndjson: a sheet list the strict read cannot prove is MalformedWorkbookXml" {
    const alloc = std.testing.allocator;
    var threaded: std.Io.Threaded = .init(alloc, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    var err_buf: [128]u8 = undefined;
    const path = try writeS3bConditionalFormatsFixture(io, &tt, "s3b_cf_wb_broken.xlsx");
    defer alloc.free(path);
    // A bad entity in a sheet-name carrier: the workbook-level decode
    // refuses before any sheet part is walked.
    try zlsx_pkg.pivots.fixture.patchPart(alloc, io, path, "xl/workbook.xml", "name=\"Report\"", "name=\"Rep&bogus;ort\"");
    const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
    defer zlsx_editor_close(ed);
    var diag = freshDiag();
    var out_ptr: ?[*]u8 = @ptrFromInt(0x1000);
    var out_len: usize = 99;
    try std.testing.expectEqual(ZLSX_REFUSED, zlsx_editor_conditional_formats_ndjson(ed, &out_ptr, &out_len, &diag, &err_buf, err_buf.len));
    try std.testing.expectEqualStrings("MalformedWorkbookXml", diagName(&diag));
    try std.testing.expectEqualStrings("MalformedWorkbookXml", std.mem.sliceTo(&err_buf, 0));
    try std.testing.expectEqual(plane_none, diag.plane);
    try std.testing.expect(out_ptr == null);
    try std.testing.expectEqual(@as(usize, 0), out_len);
}

test "S3b conditional_formats_ndjson: a part the archive cannot materialise folds at the graph probe" {
    const alloc = std.testing.allocator;
    var threaded: std.Io.Threaded = .init(alloc, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    var err_buf: [128]u8 = undefined;
    const path = try writeS3bConditionalFormatsFixture(io, &tt, "s3b_cf_crc.xlsx");
    defer alloc.free(path);
    // A flipped byte deep in the second sheet's stored payload: the
    // graph probe materialises it first, so the zip layer's own error
    // folds to the graph's verdict there rather than escaping as a
    // generic -1 (Codex #216 r1 S3B-ERR-602).
    try corruptPartPayload(alloc, io, path, "xl/worksheets/sheet2.xml");
    const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
    defer zlsx_editor_close(ed);
    var diag = freshDiag();
    var out_ptr: ?[*]u8 = @ptrFromInt(0x1000);
    var out_len: usize = 99;
    try std.testing.expectEqual(ZLSX_REFUSED, zlsx_editor_conditional_formats_ndjson(ed, &out_ptr, &out_len, &diag, &err_buf, err_buf.len));
    try std.testing.expectEqualStrings("MalformedWorkbookXml", diagName(&diag));
    try std.testing.expectEqual(plane_none, diag.plane);
    try std.testing.expect(out_ptr == null);
    try std.testing.expectEqual(@as(usize, 0), out_len);
}

test "S3b conditional_formats_ndjson: a delete collapsing a rule's whole target refuses typed, nothing mutates" {
    const alloc = std.testing.allocator;
    var threaded: std.Io.Threaded = .init(alloc, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    var err_buf: [128]u8 = undefined;
    const path = try tt.path(alloc, io, "s3b_cf_collapse.xlsx");
    defer alloc.free(path);
    {
        const w = zlsx_writer_create(&err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
        defer zlsx_writer_close(w);
        var dxf = std.mem.zeroes(CDxf);
        dxf.bold = 1;
        var dxf_id: u32 = 0;
        if (zlsx_writer_add_dxf(w, &dxf, &dxf_id, &err_buf, err_buf.len) != 0)
            return error.TestUnexpectedResult;
        const s1 = "Data";
        const sw1 = zlsx_writer_add_sheet(w, s1.ptr, s1.len, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
        const row1 = [_]CCell{ toCCell(.{ .integer = 1 }), toCCell(.{ .integer = 2 }) };
        if (zlsx_sheet_writer_write_row(sw1, &row1, row1.len, &err_buf, err_buf.len) != 0)
            return error.TestUnexpectedResult;
        const r1 = "A1:D1";
        const fe = "A1>0";
        if (zlsx_sheet_writer_add_conditional_format_expression(sw1, r1.ptr, r1.len, fe.ptr, fe.len, dxf_id, &err_buf, err_buf.len) != 0)
            return error.TestUnexpectedResult;
        if (zlsx_writer_save(w, path.ptr, path.len, &err_buf, err_buf.len) != 0)
            return error.TestUnexpectedResult;
    }
    const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
    defer zlsx_editor_close(ed);
    var diag = freshDiag();
    // The rule's whole target is row 1: the delete refuses typed at
    // the pre-mutation probe (Excel deletes such a rule; zlsx cannot
    // excise it mid-walk and refuses rather than retarget it).
    try std.testing.expectEqual(ZLSX_REFUSED, zlsx_editor_delete_row(ed, 0, 1, &diag, &err_buf, err_buf.len));
    try std.testing.expectEqualStrings("SqrefCollapseUnsafe", diagName(&diag));
    try std.testing.expectEqualStrings("SqrefCollapseUnsafe", std.mem.sliceTo(&err_buf, 0));
    try std.testing.expectEqual(plane_none, diag.plane);
    // Nothing mutated: the read still serves the original envelope.
    var out_ptr: ?[*]u8 = null;
    var out_len: usize = 0;
    try std.testing.expectEqual(ZLSX_OK, zlsx_editor_conditional_formats_ndjson(ed, &out_ptr, &out_len, &diag, &err_buf, err_buf.len));
    const got = out_ptr.?[0..out_len];
    try std.testing.expect(std.mem.indexOf(u8, got, "\"sqref\":\"A1:D1\",\"rule_type\":\"expression\",\"formulas\":[\"A1>0\"]") != null);
    zlsx_buffer_release(out_ptr, out_len);
}

fn conditionalFormatsNdjsonForFailures(alloc: std.mem.Allocator, wb: *zlsx_pkg.Workbook) !void {
    const bytes = try conditionalFormatsNdjsonOwned(alloc, wb);
    alloc.free(bytes);
}

test "S3b conditional_formats_ndjson: every allocation failure is OutOfMemory, never WriteFailed" {
    const alloc = std.testing.allocator;
    var threaded: std.Io.Threaded = .init(alloc, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try writeS3bConditionalFormatsFixture(io, &tt, "s3b_cf_oom.xlsx");
    defer alloc.free(path);
    var wb = try zlsx_pkg.Workbook.open(alloc, io, path);
    defer wb.deinit();
    // Prime the workbook-owned caches (part names, part bytes) so the
    // sweep exercises the builder's own allocations — the shared
    // writer's own OOM test does the same.
    {
        const primed = try conditionalFormatsNdjsonOwned(alloc, &wb);
        alloc.free(primed);
    }
    try std.testing.checkAllAllocationFailures(alloc, conditionalFormatsNdjsonForFailures, .{&wb});
}

/// The pkg anchors fixture (`anchor_ndjson.zig::fixture.write`,
/// `.with_absolute`): `Data` carries one two-cell image; `Report`'s
/// drawing is chart-first in document order — a one-cell bar chart
/// with three series refs, a two-cell image, an absolute image — so
/// the stream crosses a sheet boundary, regroups images before
/// charts, and carries all three anchor kinds across the boundary a
/// binding crosses. No C authoring surface writes drawings, so the
/// fixture stays the package's.
fn writeS3bAnchorsFixture(io: std.Io, tt: *TestTmp, name: []const u8) ![:0]u8 {
    const alloc = std.testing.allocator;
    const path = try tt.path(alloc, io, name);
    errdefer alloc.free(path);
    try zlsx_pkg.anchors_ndjson.fixture.write(alloc, io, path, .with_absolute);
    return path;
}

/// The frozen stream over that fixture — the literal the CLI's own
/// `runAnchorsCommand` test pins (docs/cli.md, "anchors").
const s3b_anchors_frozen = std.fmt.comptimePrint(
    "{{\"kind\":\"image_anchor\",\"sheet\":\"Data\",\"sheet_idx\":0,\"part\":\"xl/media/image1.png\"," ++
        "\"anchor\":\"two_cell\",\"from\":{{\"row\":1,\"col\":1,\"row_off\":0,\"col_off\":0}}," ++
        "\"to\":{{\"row\":4,\"col\":3,\"row_off\":0,\"col_off\":0}},\"absolute\":null,\"bytes\":{d}}}\n" ++
        "{{\"kind\":\"image_anchor\",\"sheet\":\"Report\",\"sheet_idx\":1,\"part\":\"xl/media/image1.png\"," ++
        "\"anchor\":\"two_cell\",\"from\":{{\"row\":3,\"col\":2,\"row_off\":0,\"col_off\":9525}}," ++
        "\"to\":{{\"row\":8,\"col\":5,\"row_off\":19050,\"col_off\":0}},\"absolute\":null,\"bytes\":{d}}}\n" ++
        "{{\"kind\":\"image_anchor\",\"sheet\":\"Report\",\"sheet_idx\":1,\"part\":\"xl/media/image1.png\"," ++
        "\"anchor\":\"absolute\",\"from\":null,\"to\":null," ++
        "\"absolute\":{{\"x\":1000,\"y\":2000,\"cx\":914400,\"cy\":457200}},\"bytes\":{d}}}\n" ++
        "{{\"kind\":\"chart_anchor\",\"sheet\":\"Report\",\"sheet_idx\":1,\"part\":\"xl/charts/chart1.xml\"," ++
        "\"anchor\":\"one_cell\",\"from\":{{\"row\":2,\"col\":6,\"row_off\":0,\"col_off\":0}},\"to\":null," ++
        "\"absolute\":null,\"chart_type\":\"bar\"," ++
        "\"series_refs\":[\"Data!$B$1\",\"Data!$A$2:$A$4\",\"Data!$B$2:$B$4\"]}}\n",
    .{ zlsx_pkg.anchors_ndjson.fixture.png_bytes.len, zlsx_pkg.anchors_ndjson.fixture.png_bytes.len, zlsx_pkg.anchors_ndjson.fixture.png_bytes.len },
);

test "S3b anchors_ndjson: the shared writer's bytes, current after a rename and a row insert, empty without drawings" {
    const alloc = std.testing.allocator;
    var threaded: std.Io.Threaded = .init(alloc, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    var err_buf: [128]u8 = undefined;

    {
        const path = try writeS3bAnchorsFixture(io, &tt, "s3b_anchors.xlsx");
        defer alloc.free(path);
        const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
        defer zlsx_editor_close(ed);
        var diag = freshDiag();
        var out_ptr: ?[*]u8 = null;
        var out_len: usize = 0;
        try std.testing.expectEqual(ZLSX_OK, zlsx_editor_anchors_ndjson(ed, &out_ptr, &out_len, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings(s3b_anchors_frozen, out_ptr.?[0..out_len]);
        zlsx_buffer_release(out_ptr, out_len);

        // A structural edit is what the read sees: the sheet inventory
        // is re-read from the current workbook.xml, no save in between.
        const nm = "Facts";
        try std.testing.expectEqual(ZLSX_OK, zlsx_editor_rename_sheet(ed, 0, nm.ptr, nm.len, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqual(ZLSX_OK, zlsx_editor_anchors_ndjson(ed, &out_ptr, &out_len, &diag, &err_buf, err_buf.len));
        const got = out_ptr.?[0..out_len];
        try std.testing.expect(std.mem.startsWith(u8, got, "{\"kind\":\"image_anchor\",\"sheet\":\"Facts\",\"sheet_idx\":0,"));
        try std.testing.expect(std.mem.indexOf(u8, got, "\"sheet\":\"Data\"") == null);
        // The chart's series formulas rode the rename with the other
        // carriers (the chart `<c:f>` sweep, `Workbook.rewriteAllChartFormulas`)
        // and the read reports the respelled part.
        try std.testing.expect(std.mem.indexOf(u8, got, "\"series_refs\":[\"Facts!$B$1\",\"Facts!$A$2:$A$4\",\"Facts!$B$2:$B$4\"]") != null);
        try std.testing.expect(std.mem.indexOf(u8, got, "Data!") == null);
        zlsx_buffer_release(out_ptr, out_len);

        // A row insert on the image's sheet moves its anchor with the
        // grid — the drawing sweep's work, visible with no save.
        try std.testing.expectEqual(ZLSX_OK, zlsx_editor_insert_row(ed, 0, 1, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqual(ZLSX_OK, zlsx_editor_anchors_ndjson(ed, &out_ptr, &out_len, &diag, &err_buf, err_buf.len));
        const moved = out_ptr.?[0..out_len];
        try std.testing.expect(std.mem.indexOf(u8, moved, "\"sheet\":\"Facts\",\"sheet_idx\":0,\"part\":\"xl/media/image1.png\",\"anchor\":\"two_cell\",\"from\":{\"row\":2,\"col\":1,\"row_off\":0,\"col_off\":0},\"to\":{\"row\":5,\"col\":3,\"row_off\":0,\"col_off\":0}") != null);
        // The other sheet's anchors did not move; the chart on Report
        // names the edited sheet, so its series formulas shifted with
        // the grid while the anchor stayed.
        try std.testing.expect(std.mem.indexOf(u8, moved, "\"sheet\":\"Report\",\"sheet_idx\":1,\"part\":\"xl/media/image1.png\",\"anchor\":\"two_cell\",\"from\":{\"row\":3,\"col\":2,\"row_off\":0,\"col_off\":9525}") != null);
        try std.testing.expect(std.mem.indexOf(u8, moved, "\"anchor\":\"one_cell\",\"from\":{\"row\":2,\"col\":6,\"row_off\":0,\"col_off\":0}") != null);
        try std.testing.expect(std.mem.indexOf(u8, moved, "\"series_refs\":[\"Facts!$B$2\",\"Facts!$A$3:$A$5\",\"Facts!$B$3:$B$5\"]") != null);
        try std.testing.expect(std.mem.indexOf(u8, moved, "Data!") == null);
        zlsx_buffer_release(out_ptr, out_len);

        // NULL out pointers are about the call — either one, with the
        // present one reset from its poison.
        try std.testing.expectEqual(ZLSX_ERROR, zlsx_editor_anchors_ndjson(ed, null, &out_len, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings("NullOutPointer", std.mem.sliceTo(&err_buf, 0));
        try std.testing.expectEqual(@as(usize, 0), out_len);
        out_ptr = @ptrFromInt(0x1000);
        try std.testing.expectEqual(ZLSX_ERROR, zlsx_editor_anchors_ndjson(ed, &out_ptr, null, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings("NullOutPointer", std.mem.sliceTo(&err_buf, 0));
        try std.testing.expect(out_ptr == null);
        // Re-poisoned: the NULL-editor path's own reset is what this pins.
        out_ptr = @ptrFromInt(0x1000);
        out_len = 99;
        try std.testing.expectEqual(ZLSX_ERROR, zlsx_editor_anchors_ndjson(null, &out_ptr, &out_len, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings("InvalidInput", std.mem.sliceTo(&err_buf, 0));
        try std.testing.expect(out_ptr == null);
        try std.testing.expectEqual(@as(usize, 0), out_len);

        // struct_size below v1: the outputs reset first, then the diag
        // is rejected before a byte of it is written — the whole struct
        // compared, not a field or two.
        var small = std.mem.zeroes(CDiag);
        small.struct_size = @sizeOf(CDiag) - 1;
        small.plane = 7;
        small.error_name[0] = 'x';
        const small_before = std.mem.toBytes(small);
        out_ptr = @ptrFromInt(0x1000);
        out_len = 99;
        try std.testing.expectEqual(ZLSX_ERROR, zlsx_editor_anchors_ndjson(ed, &out_ptr, &out_len, &small, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings("StructSizeTooSmall", std.mem.sliceTo(&err_buf, 0));
        try std.testing.expect(out_ptr == null);
        try std.testing.expectEqual(@as(usize, 0), out_len);
        const small_after = std.mem.toBytes(small);
        try std.testing.expectEqualSlices(u8, &small_before, &small_after);
    }
    // No drawings: success, nothing to release, the poison reset.
    {
        const path = try writeS3aFixture(io, &tt, "s3b_no_anchors.xlsx");
        defer alloc.free(path);
        const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
        defer zlsx_editor_close(ed);
        var out_ptr: ?[*]u8 = @ptrFromInt(0x1000);
        var out_len: usize = 99;
        try std.testing.expectEqual(ZLSX_OK, zlsx_editor_anchors_ndjson(ed, &out_ptr, &out_len, null, &err_buf, err_buf.len));
        try std.testing.expect(out_ptr == null);
        try std.testing.expectEqual(@as(usize, 0), out_len);
        zlsx_buffer_release(out_ptr, out_len);
    }
}

test "S3b anchors_ndjson: an inventory the read cannot serve refuses whole, typed, nothing handed out" {
    const alloc = std.testing.allocator;
    var threaded: std.Io.Threaded = .init(alloc, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    var err_buf: [128]u8 = undefined;
    const patch = zlsx_pkg.anchors_ndjson.fixture.patchPart;

    const Case = struct { name: []const u8, part: []const u8, old: []const u8, new: []const u8, verdict: []const u8 };
    const cases = [_]Case{
        // Data's image blip names a relationship the drawing's rels
        // do not hold: the drawing graph cannot be read whole.
        .{ .name = "s3b_anchors_blip.xlsx", .part = "xl/drawings/drawing2.xml", .old = "r:embed=\"rIdI1\"", .new = "r:embed=\"rIdXX\"", .verdict = "MalformedDrawingXml" },
        // The SECOND sheet's drawing is the broken one while Data's
        // record is perfectly servable — the whole-inventory rule
        // hands out nothing.
        .{ .name = "s3b_anchors_second.xlsx", .part = "xl/drawings/drawing1.xml", .old = "<xdr:to>", .new = "<xdr:zz>", .verdict = "MalformedDrawingXml" },
        // A bad entity in a sheet-name carrier: the workbook-level
        // decode refuses before any drawing is walked.
        .{ .name = "s3b_anchors_wb.xlsx", .part = "xl/workbook.xml", .old = "name=\"Report\"", .new = "name=\"Rep&bogus;ort\"", .verdict = "MalformedWorkbookXml" },
        // A series ref whose carrier does not decode: the record
        // would lie, so the walk refuses it.
        .{ .name = "s3b_anchors_ref.xlsx", .part = "xl/charts/chart1.xml", .old = "<c:f>Data!$B$1</c:f>", .new = "<c:f>Data!$B&bogus;$1</c:f>", .verdict = "MalformedDrawingXml" },
        // A spreadsheetDrawing binding under a name the anchor walk
        // cannot spell (one byte past the resolver's 100-byte limit):
        // an anchor under it would be neither listed nor moved, so the
        // strict read refuses the drawing (the namespace-aware drawing
        // slice).
        .{ .name = "s3b_anchors_nsbind.xlsx", .part = "xl/drawings/drawing1.xml", .old = "<xdr:wsDr ", .new = "<xdr:wsDr xmlns:" ++ ("p" ** 101) ++ "=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" ", .verdict = "MalformedDrawingXml" },
    };
    for (cases) |case| {
        const path = try writeS3bAnchorsFixture(io, &tt, case.name);
        defer alloc.free(path);
        try patch(alloc, io, path, case.part, case.old, case.new);
        const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
        defer zlsx_editor_close(ed);
        var diag = freshDiag();
        // Poisoned on entry: the refusal path itself must reset the pair.
        var out_ptr: ?[*]u8 = @ptrFromInt(0x1000);
        var out_len: usize = 99;
        try std.testing.expectEqual(ZLSX_REFUSED, zlsx_editor_anchors_ndjson(ed, &out_ptr, &out_len, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings(case.verdict, diagName(&diag));
        try std.testing.expectEqualStrings(case.verdict, std.mem.sliceTo(&err_buf, 0));
        try std.testing.expectEqual(plane_none, diag.plane);
        try std.testing.expect(out_ptr == null);
        try std.testing.expectEqual(@as(usize, 0), out_len);
    }
}

test "S3b anchors_ndjson: an anchor on a worksheet part the workbook does not list is DrawingOnUnlistedSheet" {
    const alloc = std.testing.allocator;
    var threaded: std.Io.Threaded = .init(alloc, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    var err_buf: [128]u8 = undefined;
    const path = try writeS3bAnchorsFixture(io, &tt, "s3b_anchors_orphan.xlsx");
    defer alloc.free(path);
    // A copy of Report's part under a name no <sheet> entry reaches,
    // its drawing reference and rels riding along: the walkers still
    // key the anchors by the part, the inventory cannot attribute
    // them, and the read refuses rather than drop them.
    {
        var store = try zlsx_pkg.PartStore.open(alloc, io, path);
        defer store.deinit();
        const sheet2 = (try store.part("xl/worksheets/sheet2.xml")) orelse return error.TestUnexpectedResult;
        try store.addPart("xl/worksheets/sheet9.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml", sheet2.bytes);
        const rels = (try store.part("xl/worksheets/_rels/sheet2.xml.rels")) orelse return error.TestUnexpectedResult;
        try store.addPart("xl/worksheets/_rels/sheet9.xml.rels", "application/vnd.openxmlformats-package.relationships+xml", rels.bytes);
        try store.save(io, path);
    }
    const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
    defer zlsx_editor_close(ed);
    var diag = freshDiag();
    var out_ptr: ?[*]u8 = @ptrFromInt(0x1000);
    var out_len: usize = 99;
    try std.testing.expectEqual(ZLSX_REFUSED, zlsx_editor_anchors_ndjson(ed, &out_ptr, &out_len, &diag, &err_buf, err_buf.len));
    try std.testing.expectEqualStrings("DrawingOnUnlistedSheet", diagName(&diag));
    try std.testing.expectEqualStrings("DrawingOnUnlistedSheet", std.mem.sliceTo(&err_buf, 0));
    try std.testing.expectEqual(plane_none, diag.plane);
    try std.testing.expect(out_ptr == null);
    try std.testing.expectEqual(@as(usize, 0), out_len);
}

test "S3b anchors_ndjson: a part the archive cannot materialise folds to the graph it breaks" {
    const alloc = std.testing.allocator;
    var threaded: std.Io.Threaded = .init(alloc, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    var err_buf: [128]u8 = undefined;
    // A listed sheet part: the inventory's probe materialises it
    // first, so the zip layer's own error folds to the workbook's
    // verdict there (the conditional-formats read's rule, Codex #216
    // r1 S3B-ERR-602).
    {
        const path = try writeS3bAnchorsFixture(io, &tt, "s3b_anchors_crc_sheet.xlsx");
        defer alloc.free(path);
        try corruptPartPayload(alloc, io, path, "xl/worksheets/sheet2.xml");
        const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
        defer zlsx_editor_close(ed);
        var diag = freshDiag();
        var out_ptr: ?[*]u8 = @ptrFromInt(0x1000);
        var out_len: usize = 99;
        try std.testing.expectEqual(ZLSX_REFUSED, zlsx_editor_anchors_ndjson(ed, &out_ptr, &out_len, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings("MalformedWorkbookXml", diagName(&diag));
        try std.testing.expectEqual(plane_none, diag.plane);
        try std.testing.expect(out_ptr == null);
        try std.testing.expectEqual(@as(usize, 0), out_len);
    }
    // A drawing part: only the walk reaches it, and a chain it cannot
    // follow is a drawing graph that cannot be read whole — never the
    // zip layer's own name across the boundary.
    {
        const path = try writeS3bAnchorsFixture(io, &tt, "s3b_anchors_crc_drawing.xlsx");
        defer alloc.free(path);
        try corruptPartPayload(alloc, io, path, "xl/drawings/drawing1.xml");
        const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
        defer zlsx_editor_close(ed);
        var diag = freshDiag();
        var out_ptr: ?[*]u8 = @ptrFromInt(0x1000);
        var out_len: usize = 99;
        try std.testing.expectEqual(ZLSX_REFUSED, zlsx_editor_anchors_ndjson(ed, &out_ptr, &out_len, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings("MalformedDrawingXml", diagName(&diag));
        try std.testing.expectEqual(plane_none, diag.plane);
        try std.testing.expect(out_ptr == null);
        try std.testing.expectEqual(@as(usize, 0), out_len);
    }
}

fn anchorsNdjsonForFailures(alloc: std.mem.Allocator, wb: *zlsx_pkg.Workbook) !void {
    const bytes = try anchorsNdjsonOwned(alloc, wb);
    alloc.free(bytes);
}

test "S3b anchors_ndjson: every allocation failure is OutOfMemory, never WriteFailed" {
    const alloc = std.testing.allocator;
    var threaded: std.Io.Threaded = .init(alloc, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try writeS3bAnchorsFixture(io, &tt, "s3b_anchors_oom.xlsx");
    defer alloc.free(path);
    var wb = try zlsx_pkg.Workbook.open(alloc, io, path);
    defer wb.deinit();
    // Prime the workbook-owned caches (part names, part bytes) so the
    // sweep exercises the builder's own allocations — the shared
    // writer's own OOM test does the same.
    {
        const primed = try anchorsNdjsonOwned(alloc, &wb);
        alloc.free(primed);
    }
    try std.testing.checkAllAllocationFailures(alloc, anchorsNdjsonForFailures, .{&wb});
}

/// The pkg sheet-props fixture (`sheet_props_ndjson.zig::fixture.write`):
/// `Data` carries a frozen pane and a spliced extent, `Report` a
/// split pane with a fractional split and an extent, `Bare` neither,
/// and the workbook a full `<calcPr>` — so the stream carries a
/// frozen record, a split record (which the lenient view would
/// narrow away) and a record of nulls, and the calc record has every
/// field set. The writer emits no `<dimension>`, no `<calcPr>` and
/// has no split-pane surface, so the fixture stays the package's.
fn writeS3bSheetPropsFixture(io: std.Io, tt: *TestTmp, name: []const u8) ![:0]u8 {
    const alloc = std.testing.allocator;
    const path = try tt.path(alloc, io, name);
    errdefer alloc.free(path);
    try zlsx_pkg.sheet_props_ndjson.fixture.write(alloc, io, path);
    return path;
}

/// The frozen streams over that fixture — the literals the CLI's own
/// `runSheetPropsCommand` / `runCalcPropsCommand` tests pin
/// (docs/cli.md, "sheet-props" / "calc-props").
const s3b_sheet_props_frozen =
    "{\"kind\":\"sheet_props\",\"sheet\":\"Data\",\"sheet_idx\":0,\"dimension\":\"A1:B3\"," ++
    "\"pane\":{\"x_split\":2,\"y_split\":1,\"top_left_cell\":\"C2\",\"active_pane\":\"bottomRight\",\"state\":\"frozen\"}}\n" ++
    "{\"kind\":\"sheet_props\",\"sheet\":\"Report\",\"sheet_idx\":1,\"dimension\":\"A1:C2\"," ++
    "\"pane\":{\"x_split\":2865,\"y_split\":1215.5,\"top_left_cell\":\"C4\",\"active_pane\":\"bottomRight\",\"state\":\"split\"}}\n" ++
    "{\"kind\":\"sheet_props\",\"sheet\":\"Bare\",\"sheet_idx\":2,\"dimension\":null,\"pane\":null}\n";
const s3b_calc_props_frozen =
    "{\"kind\":\"calc_props\",\"calc_id\":191029,\"full_calc_on_load\":true,\"iterate\":true,\"iterate_count\":100,\"iterate_delta\":0.001}\n";
const s3b_calc_props_absent =
    "{\"kind\":\"calc_props\",\"calc_id\":null,\"full_calc_on_load\":null,\"iterate\":null,\"iterate_count\":null,\"iterate_delta\":null}\n";

test "S3b sheet_props_ndjson: the shared writer's bytes, current after a rename and a row insert, the call errors" {
    const alloc = std.testing.allocator;
    var threaded: std.Io.Threaded = .init(alloc, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    var err_buf: [128]u8 = undefined;

    const path = try writeS3bSheetPropsFixture(io, &tt, "s3b_sheet_props.xlsx");
    defer alloc.free(path);
    const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
    defer zlsx_editor_close(ed);
    var diag = freshDiag();
    var out_ptr: ?[*]u8 = null;
    var out_len: usize = 0;
    try std.testing.expectEqual(ZLSX_OK, zlsx_editor_sheet_props_ndjson(ed, &out_ptr, &out_len, &diag, &err_buf, err_buf.len));
    try std.testing.expectEqualStrings(s3b_sheet_props_frozen, out_ptr.?[0..out_len]);
    zlsx_buffer_release(out_ptr, out_len);

    // A structural edit is what the read sees: the sheet inventory is
    // re-read from the current workbook.xml, no save in between.
    const nm = "Facts";
    try std.testing.expectEqual(ZLSX_OK, zlsx_editor_rename_sheet(ed, 0, nm.ptr, nm.len, &diag, &err_buf, err_buf.len));
    try std.testing.expectEqual(ZLSX_OK, zlsx_editor_sheet_props_ndjson(ed, &out_ptr, &out_len, &diag, &err_buf, err_buf.len));
    const renamed = out_ptr.?[0..out_len];
    try std.testing.expect(std.mem.startsWith(u8, renamed, "{\"kind\":\"sheet_props\",\"sheet\":\"Facts\",\"sheet_idx\":0,\"dimension\":\"A1:B3\","));
    try std.testing.expect(std.mem.indexOf(u8, renamed, "\"sheet\":\"Data\"") == null);
    zlsx_buffer_release(out_ptr, out_len);

    // A row insert below the frozen row on that sheet: the sheet
    // sweep grows the extent and moves the pane's top-left cell with
    // the grid while the split itself (one frozen row above the
    // insertion) holds — visible with no save. The other sheets'
    // records are untouched.
    try std.testing.expectEqual(ZLSX_OK, zlsx_editor_insert_row(ed, 0, 2, &diag, &err_buf, err_buf.len));
    try std.testing.expectEqual(ZLSX_OK, zlsx_editor_sheet_props_ndjson(ed, &out_ptr, &out_len, &diag, &err_buf, err_buf.len));
    try std.testing.expectEqualStrings(
        "{\"kind\":\"sheet_props\",\"sheet\":\"Facts\",\"sheet_idx\":0,\"dimension\":\"A1:B4\"," ++
            "\"pane\":{\"x_split\":2,\"y_split\":1,\"top_left_cell\":\"C3\",\"active_pane\":\"bottomRight\",\"state\":\"frozen\"}}\n" ++
            "{\"kind\":\"sheet_props\",\"sheet\":\"Report\",\"sheet_idx\":1,\"dimension\":\"A1:C2\"," ++
            "\"pane\":{\"x_split\":2865,\"y_split\":1215.5,\"top_left_cell\":\"C4\",\"active_pane\":\"bottomRight\",\"state\":\"split\"}}\n" ++
            "{\"kind\":\"sheet_props\",\"sheet\":\"Bare\",\"sheet_idx\":2,\"dimension\":null,\"pane\":null}\n",
        out_ptr.?[0..out_len],
    );
    zlsx_buffer_release(out_ptr, out_len);

    // The split pane the record reports is the one the row edit
    // refuses: the read and the editor's own contract agree, and the
    // refused edit leaves the stream exactly as it was.
    try std.testing.expectEqual(ZLSX_REFUSED, zlsx_editor_insert_row(ed, 1, 1, &diag, &err_buf, err_buf.len));
    try std.testing.expectEqualStrings("SplitPaneNotSupported", diagName(&diag));
    try std.testing.expectEqual(ZLSX_OK, zlsx_editor_sheet_props_ndjson(ed, &out_ptr, &out_len, &diag, &err_buf, err_buf.len));
    try std.testing.expect(std.mem.indexOf(u8, out_ptr.?[0..out_len], "\"sheet\":\"Report\",\"sheet_idx\":1,\"dimension\":\"A1:C2\",\"pane\":{\"x_split\":2865,\"y_split\":1215.5,\"top_left_cell\":\"C4\"") != null);
    zlsx_buffer_release(out_ptr, out_len);

    // NULL out pointers are about the call — either one, with the
    // present one reset from its poison.
    try std.testing.expectEqual(ZLSX_ERROR, zlsx_editor_sheet_props_ndjson(ed, null, &out_len, &diag, &err_buf, err_buf.len));
    try std.testing.expectEqualStrings("NullOutPointer", std.mem.sliceTo(&err_buf, 0));
    try std.testing.expectEqual(@as(usize, 0), out_len);
    out_ptr = @ptrFromInt(0x1000);
    try std.testing.expectEqual(ZLSX_ERROR, zlsx_editor_sheet_props_ndjson(ed, &out_ptr, null, &diag, &err_buf, err_buf.len));
    try std.testing.expectEqualStrings("NullOutPointer", std.mem.sliceTo(&err_buf, 0));
    try std.testing.expect(out_ptr == null);
    // Re-poisoned: the NULL-editor path's own reset is what this pins.
    out_ptr = @ptrFromInt(0x1000);
    out_len = 99;
    try std.testing.expectEqual(ZLSX_ERROR, zlsx_editor_sheet_props_ndjson(null, &out_ptr, &out_len, &diag, &err_buf, err_buf.len));
    try std.testing.expectEqualStrings("InvalidInput", std.mem.sliceTo(&err_buf, 0));
    try std.testing.expect(out_ptr == null);
    try std.testing.expectEqual(@as(usize, 0), out_len);

    // struct_size below v1: the outputs reset first, then the diag
    // is rejected before a byte of it is written — the whole struct
    // compared, not a field or two.
    var small = std.mem.zeroes(CDiag);
    small.struct_size = @sizeOf(CDiag) - 1;
    small.plane = 7;
    small.error_name[0] = 'x';
    const small_before = std.mem.toBytes(small);
    out_ptr = @ptrFromInt(0x1000);
    out_len = 99;
    try std.testing.expectEqual(ZLSX_ERROR, zlsx_editor_sheet_props_ndjson(ed, &out_ptr, &out_len, &small, &err_buf, err_buf.len));
    try std.testing.expectEqualStrings("StructSizeTooSmall", std.mem.sliceTo(&err_buf, 0));
    try std.testing.expect(out_ptr == null);
    try std.testing.expectEqual(@as(usize, 0), out_len);
    const small_after = std.mem.toBytes(small);
    try std.testing.expectEqualSlices(u8, &small_before, &small_after);

    // A NULL diag is allowed on the success path, as on every sibling.
    try std.testing.expectEqual(ZLSX_OK, zlsx_editor_sheet_props_ndjson(ed, &out_ptr, &out_len, null, &err_buf, err_buf.len));
    try std.testing.expect(out_len != 0);
    zlsx_buffer_release(out_ptr, out_len);
}

test "S3b sheet_props_ndjson: a fresh writer's sheets are records of nulls, never an empty stream" {
    const alloc = std.testing.allocator;
    var threaded: std.Io.Threaded = .init(alloc, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    var err_buf: [128]u8 = undefined;
    const path = try writeS3aFixture(io, &tt, "s3b_sheet_props_plain.xlsx");
    defer alloc.free(path);
    const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
    defer zlsx_editor_close(ed);
    var out_ptr: ?[*]u8 = @ptrFromInt(0x1000);
    var out_len: usize = 99;
    try std.testing.expectEqual(ZLSX_OK, zlsx_editor_sheet_props_ndjson(ed, &out_ptr, &out_len, null, &err_buf, err_buf.len));
    const got = out_ptr.?[0..out_len];
    // One record per sheet, each with neither extent nor pane: the
    // fresh writer emits no `<dimension>` and no views without a
    // freeze.
    try std.testing.expectEqual(@as(usize, std.mem.count(u8, got, "\n")), std.mem.count(u8, got, "\"dimension\":null,\"pane\":null}\n"));
    try std.testing.expect(std.mem.startsWith(u8, got, "{\"kind\":\"sheet_props\",\"sheet\":"));
    zlsx_buffer_release(out_ptr, out_len);
}

test "S3b sheet_props_ndjson: an inventory the read cannot serve refuses whole, typed, nothing handed out" {
    const alloc = std.testing.allocator;
    var threaded: std.Io.Threaded = .init(alloc, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    var err_buf: [128]u8 = undefined;
    const patch = zlsx_pkg.sheet_props_ndjson.fixture.patchPart;

    const Case = struct { name: []const u8, part: []const u8, old: []const u8, new: []const u8, verdict: []const u8 };
    const cases = [_]Case{
        // Two extents on Data: maxOccurs=1 is a refusal, not a pick.
        .{ .name = "s3b_sp_dim2.xlsx", .part = "xl/worksheets/sheet1.xml", .old = "<dimension ref=\"A1:B3\"/>", .new = "<dimension ref=\"A1:B3\"/><dimension ref=\"A1\"/>", .verdict = "MalformedSheetXml" },
        // The SECOND sheet's pane carries a duplicate attribute while
        // Data's record is perfectly servable — the whole-inventory
        // rule hands out nothing.
        .{ .name = "s3b_sp_second.xlsx", .part = "xl/worksheets/sheet2.xml", .old = "state=\"split\"", .new = "state=\"split\" state=\"frozen\"", .verdict = "MalformedSheetXml" },
        // A pane carrier that does not decode: the record would lie.
        .{ .name = "s3b_sp_carrier.xlsx", .part = "xl/worksheets/sheet2.xml", .old = "topLeftCell=\"C4\"", .new = "topLeftCell=\"C&bogus;\"", .verdict = "MalformedSheetXml" },
        // An MCE branch at the views slot: the walk cannot rule a
        // projected pane in or out.
        .{ .name = "s3b_sp_mce.xlsx", .part = "xl/worksheets/sheet2.xml", .old = "<sheetViews>", .new = "<sheetViews><mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"><mc:Choice Requires=\"x14\"/></mc:AlternateContent>", .verdict = "MalformedSheetXml" },
        // A bad entity in a sheet-name carrier: the workbook-level
        // read refuses before any sheet part is walked.
        .{ .name = "s3b_sp_wb.xlsx", .part = "xl/workbook.xml", .old = "name=\"Report\"", .new = "name=\"Rep&bogus;ort\"", .verdict = "MalformedWorkbookXml" },
    };
    for (cases) |case| {
        const path = try writeS3bSheetPropsFixture(io, &tt, case.name);
        defer alloc.free(path);
        try patch(alloc, io, path, case.part, case.old, case.new);
        const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
        defer zlsx_editor_close(ed);
        var diag = freshDiag();
        // Poisoned on entry: the refusal path itself must reset the pair.
        var out_ptr: ?[*]u8 = @ptrFromInt(0x1000);
        var out_len: usize = 99;
        try std.testing.expectEqual(ZLSX_REFUSED, zlsx_editor_sheet_props_ndjson(ed, &out_ptr, &out_len, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings(case.verdict, diagName(&diag));
        try std.testing.expectEqualStrings(case.verdict, std.mem.sliceTo(&err_buf, 0));
        try std.testing.expectEqual(plane_none, diag.plane);
        try std.testing.expect(out_ptr == null);
        try std.testing.expectEqual(@as(usize, 0), out_len);
    }
    // A listed sheet part the archive cannot materialise: the
    // inventory's probe reaches it first, so the zip layer's own
    // error folds to the workbook's verdict there (the
    // conditional-formats read's rule, Codex #216 r1 S3B-ERR-602).
    {
        const path = try writeS3bSheetPropsFixture(io, &tt, "s3b_sp_crc.xlsx");
        defer alloc.free(path);
        try corruptPartPayload(alloc, io, path, "xl/worksheets/sheet3.xml");
        const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
        defer zlsx_editor_close(ed);
        var diag = freshDiag();
        var out_ptr: ?[*]u8 = @ptrFromInt(0x1000);
        var out_len: usize = 99;
        try std.testing.expectEqual(ZLSX_REFUSED, zlsx_editor_sheet_props_ndjson(ed, &out_ptr, &out_len, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings("MalformedWorkbookXml", diagName(&diag));
        try std.testing.expectEqual(plane_none, diag.plane);
        try std.testing.expect(out_ptr == null);
        try std.testing.expectEqual(@as(usize, 0), out_len);
    }
    // An empty `<sheets/>` — the wrapper present, its entries gone,
    // every sheet part still in the archive. The lenient opener
    // accepts it; the strict inventory refuses the sheetless workbook
    // (CT_Sheets minOccurs=1, the REL-602 rule closed for the
    // wrapper-present spelling by Codex #219 r1 S3B-REL-101), so the
    // sheet-props export can no longer hand back the empty success
    // its contract rules out — and the calc read, which used to serve
    // its record over the sheetless workbook, shares the walk's
    // verdict.
    {
        const path = try writeS3bSheetPropsFixture(io, &tt, "s3b_sp_sheetless.xlsx");
        defer alloc.free(path);
        try zlsx_pkg.conditional_formats_ndjson.fixture.emptySheets(alloc, io, path);
        const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
        defer zlsx_editor_close(ed);
        var diag = freshDiag();
        var out_ptr: ?[*]u8 = @ptrFromInt(0x1000);
        var out_len: usize = 99;
        try std.testing.expectEqual(ZLSX_REFUSED, zlsx_editor_sheet_props_ndjson(ed, &out_ptr, &out_len, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings("MalformedWorkbookXml", diagName(&diag));
        try std.testing.expectEqualStrings("MalformedWorkbookXml", std.mem.sliceTo(&err_buf, 0));
        try std.testing.expectEqual(plane_none, diag.plane);
        try std.testing.expect(out_ptr == null);
        try std.testing.expectEqual(@as(usize, 0), out_len);
        out_ptr = @ptrFromInt(0x1000);
        out_len = 99;
        try std.testing.expectEqual(ZLSX_REFUSED, zlsx_editor_calc_props_ndjson(ed, &out_ptr, &out_len, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings("MalformedWorkbookXml", diagName(&diag));
        try std.testing.expectEqualStrings("MalformedWorkbookXml", std.mem.sliceTo(&err_buf, 0));
        try std.testing.expectEqual(plane_none, diag.plane);
        try std.testing.expect(out_ptr == null);
        try std.testing.expectEqual(@as(usize, 0), out_len);
    }
}

test "S3b calc_props_ndjson: the shared writer's bytes, the absent record, mark-recalc visible with no save, the call errors" {
    const alloc = std.testing.allocator;
    var threaded: std.Io.Threaded = .init(alloc, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    var err_buf: [128]u8 = undefined;

    {
        const path = try writeS3bSheetPropsFixture(io, &tt, "s3b_calc_props.xlsx");
        defer alloc.free(path);
        const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
        defer zlsx_editor_close(ed);
        var diag = freshDiag();
        var out_ptr: ?[*]u8 = null;
        var out_len: usize = 0;
        try std.testing.expectEqual(ZLSX_OK, zlsx_editor_calc_props_ndjson(ed, &out_ptr, &out_len, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings(s3b_calc_props_frozen, out_ptr.?[0..out_len]);
        zlsx_buffer_release(out_ptr, out_len);

        // A rename is a workbook.xml rewrite the read re-walks; the
        // `<calcPr>` slot rides through it untouched.
        const nm = "Facts";
        try std.testing.expectEqual(ZLSX_OK, zlsx_editor_rename_sheet(ed, 0, nm.ptr, nm.len, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqual(ZLSX_OK, zlsx_editor_calc_props_ndjson(ed, &out_ptr, &out_len, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings(s3b_calc_props_frozen, out_ptr.?[0..out_len]);
        zlsx_buffer_release(out_ptr, out_len);

        // NULL out pointers are about the call — either one, with the
        // present one reset from its poison.
        try std.testing.expectEqual(ZLSX_ERROR, zlsx_editor_calc_props_ndjson(ed, null, &out_len, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings("NullOutPointer", std.mem.sliceTo(&err_buf, 0));
        try std.testing.expectEqual(@as(usize, 0), out_len);
        out_ptr = @ptrFromInt(0x1000);
        try std.testing.expectEqual(ZLSX_ERROR, zlsx_editor_calc_props_ndjson(ed, &out_ptr, null, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings("NullOutPointer", std.mem.sliceTo(&err_buf, 0));
        try std.testing.expect(out_ptr == null);
        out_ptr = @ptrFromInt(0x1000);
        out_len = 99;
        try std.testing.expectEqual(ZLSX_ERROR, zlsx_editor_calc_props_ndjson(null, &out_ptr, &out_len, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings("InvalidInput", std.mem.sliceTo(&err_buf, 0));
        try std.testing.expect(out_ptr == null);
        try std.testing.expectEqual(@as(usize, 0), out_len);

        var small = std.mem.zeroes(CDiag);
        small.struct_size = @sizeOf(CDiag) - 1;
        small.plane = 7;
        small.error_name[0] = 'x';
        const small_before = std.mem.toBytes(small);
        out_ptr = @ptrFromInt(0x1000);
        out_len = 99;
        try std.testing.expectEqual(ZLSX_ERROR, zlsx_editor_calc_props_ndjson(ed, &out_ptr, &out_len, &small, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings("StructSizeTooSmall", std.mem.sliceTo(&err_buf, 0));
        try std.testing.expect(out_ptr == null);
        try std.testing.expectEqual(@as(usize, 0), out_len);
        const small_after = std.mem.toBytes(small);
        try std.testing.expectEqualSlices(u8, &small_before, &small_after);
    }
    // No `<calcPr>`: a record of nulls, never `(NULL, 0)` — and the
    // mark-only transaction's `fullCalcOnLoad="1"` lands in the live
    // part, so the next read reports it with no save in between.
    {
        const path = try writeS3aFixture(io, &tt, "s3b_calc_props_plain.xlsx");
        defer alloc.free(path);
        const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
        defer zlsx_editor_close(ed);
        var diag = freshDiag();
        var out_ptr: ?[*]u8 = @ptrFromInt(0x1000);
        var out_len: usize = 99;
        try std.testing.expectEqual(ZLSX_OK, zlsx_editor_calc_props_ndjson(ed, &out_ptr, &out_len, null, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings(s3b_calc_props_absent, out_ptr.?[0..out_len]);
        zlsx_buffer_release(out_ptr, out_len);

        try std.testing.expectEqual(ZLSX_OK, zlsx_editor_mark_recalc_on_load(ed, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqual(ZLSX_OK, zlsx_editor_calc_props_ndjson(ed, &out_ptr, &out_len, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings(
            "{\"kind\":\"calc_props\",\"calc_id\":null,\"full_calc_on_load\":true,\"iterate\":null,\"iterate_count\":null,\"iterate_delta\":null}\n",
            out_ptr.?[0..out_len],
        );
        zlsx_buffer_release(out_ptr, out_len);
    }
}

test "S3b calc_props_ndjson: a slot the read cannot report faithfully refuses whole, typed, nothing handed out" {
    const alloc = std.testing.allocator;
    var threaded: std.Io.Threaded = .init(alloc, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    var err_buf: [128]u8 = undefined;
    const patch = zlsx_pkg.sheet_props_ndjson.fixture.patchPart;
    const full = "<calcPr calcId=\"191029\" fullCalcOnLoad=\"1\" iterate=\"true\" iterateCount=\"100\" iterateDelta=\"0.001\"/>";

    const Case = struct { name: []const u8, old: []const u8, new: []const u8 };
    const cases = [_]Case{
        // Two at the slot: which one Excel honours is not the reader's
        // to guess.
        .{ .name = "s3b_cp_two.xlsx", .old = "</workbook>", .new = "<calcPr calcId=\"1\"/></workbook>" },
        // A branch an MCE processor could project into the slot.
        .{ .name = "s3b_cp_mce.xlsx", .old = full, .new = "<mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"><mc:Choice Requires=\"x15\"><calcPr calcId=\"1\"/></mc:Choice></mc:AlternateContent>" },
        // A duplicate attribute; a carrier that does not decode.
        .{ .name = "s3b_cp_dup.xlsx", .old = full, .new = "<calcPr calcId=\"1\" calcId=\"2\"/>" },
        .{ .name = "s3b_cp_carrier.xlsx", .old = full, .new = "<calcPr iterate=\"&bogus;\"/>" },
        // A `<sheets>` list the strict workbook walk cannot prove: the
        // calc read runs the same walk, the same verdict.
        .{ .name = "s3b_cp_sheets.xlsx", .old = "</sheets>", .new = "</sheets><sheets/>" },
    };
    for (cases) |case| {
        const path = try writeS3bSheetPropsFixture(io, &tt, case.name);
        defer alloc.free(path);
        try patch(alloc, io, path, "xl/workbook.xml", case.old, case.new);
        const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
        defer zlsx_editor_close(ed);
        var diag = freshDiag();
        var out_ptr: ?[*]u8 = @ptrFromInt(0x1000);
        var out_len: usize = 99;
        try std.testing.expectEqual(ZLSX_REFUSED, zlsx_editor_calc_props_ndjson(ed, &out_ptr, &out_len, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings("MalformedWorkbookXml", diagName(&diag));
        try std.testing.expectEqualStrings("MalformedWorkbookXml", std.mem.sliceTo(&err_buf, 0));
        try std.testing.expectEqual(plane_none, diag.plane);
        try std.testing.expect(out_ptr == null);
        try std.testing.expectEqual(@as(usize, 0), out_len);
    }
}

fn sheetPropsNdjsonForFailures(alloc: std.mem.Allocator, wb: *zlsx_pkg.Workbook) !void {
    const bytes = try sheetPropsNdjsonOwned(alloc, wb);
    alloc.free(bytes);
}

fn calcPropsNdjsonForFailures(alloc: std.mem.Allocator, wb: *zlsx_pkg.Workbook) !void {
    const bytes = try calcPropsNdjsonOwned(alloc, wb);
    alloc.free(bytes);
}

test "S3b sheet_props_ndjson / calc_props_ndjson: every allocation failure is OutOfMemory, never WriteFailed" {
    const alloc = std.testing.allocator;
    var threaded: std.Io.Threaded = .init(alloc, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try writeS3bSheetPropsFixture(io, &tt, "s3b_sheet_props_oom.xlsx");
    defer alloc.free(path);
    var wb = try zlsx_pkg.Workbook.open(alloc, io, path);
    defer wb.deinit();
    // Prime the workbook-owned caches (part names, part bytes) so the
    // sweeps exercise the builders' own allocations — the shared
    // writer's own OOM test does the same.
    {
        const primed = try sheetPropsNdjsonOwned(alloc, &wb);
        alloc.free(primed);
        const primed_calc = try calcPropsNdjsonOwned(alloc, &wb);
        alloc.free(primed_calc);
    }
    try std.testing.checkAllAllocationFailures(alloc, sheetPropsNdjsonForFailures, .{&wb});
    try std.testing.checkAllAllocationFailures(alloc, calcPropsNdjsonForFailures, .{&wb});
}

/// Flip one byte deep inside the stored payload of `part`, found by
/// walking the local file headers by name (the name also appears in
/// `[Content_Types].xml` and the rels, so a plain search would land in
/// another part's payload). The central directory stays valid, the
/// editor opens, and the part fails when the store materialises it.
fn corruptPartPayload(alloc: std.mem.Allocator, io: std.Io, path: []const u8, part: []const u8) !void {
    const bytes = try std.Io.Dir.cwd().readFileAlloc(io, path, alloc, .limited(1 << 24));
    defer alloc.free(bytes);
    var pos: usize = 0;
    while (std.mem.indexOfPos(u8, bytes, pos, "PK\x03\x04")) |hdr| : (pos = hdr + 4) {
        if (hdr + 30 > bytes.len) break;
        const csize = std.mem.readInt(u32, bytes[hdr + 18 ..][0..4], .little);
        const nlen = std.mem.readInt(u16, bytes[hdr + 26 ..][0..2], .little);
        const elen = std.mem.readInt(u16, bytes[hdr + 28 ..][0..2], .little);
        const name_at = hdr + 30;
        if (name_at + nlen > bytes.len) break;
        if (!std.mem.eql(u8, bytes[name_at..][0..nlen], part)) continue;
        const payload = bytes[name_at + nlen + elen ..][0..csize];
        payload[payload.len / 2] ^= 0xFF;
        try std.Io.Dir.cwd().writeFile(io, .{ .sub_path = path, .data = bytes });
        return;
    }
    return error.TestUnexpectedResult;
}

test "S3a: what the lazy reads raise crosses as the workbook's verdict, not the call's" {
    const alloc = std.testing.allocator;
    var threaded: std.Io.Threaded = .init(alloc, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    var err_buf: [128]u8 = undefined;

    // A worksheet the typed parser cannot read, met when a sheet delete
    // parses the sheet it drops: `MalformedSheetXml`, -2.
    {
        const path = try writeS3aFixture(io, &tt, "s3a_lazy_sheet.xlsx");
        defer alloc.free(path);
        // The parser tolerates a lost close tag; it refuses a part
        // without its root element.
        try zlsx_pkg.pivots.fixture.patchPart(alloc, io, path, "xl/worksheets/sheet2.xml", "<worksheet ", "<wsheet ");
        const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
        defer zlsx_editor_close(ed);
        var diag = freshDiag();
        try std.testing.expectEqual(ZLSX_REFUSED, zlsx_editor_delete_sheet(ed, 1, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings("MalformedSheetXml", diagName(&diag));
        try std.testing.expectEqualStrings("MalformedSheetXml", std.mem.sliceTo(&err_buf, 0));
        try std.testing.expectEqual(plane_none, diag.plane);
    }
    // A stored name past the OLD internal 128-byte bound: that bound
    // fell in #216 r17 — a valid escape-heavy name legitimately
    // exceeds it — so the rename now ADMITS the oversized stored
    // carrier and repairs it; validation applies to the NEW name.
    {
        const path = try writeS3aFixture(io, &tt, "s3a_longname.xlsx");
        defer alloc.free(path);
        try zlsx_pkg.pivots.fixture.patchPart(alloc, io, path, "xl/workbook.xml", "name=\"Second\"", "name=\"" ++ ("S" ** 130) ++ "\"");
        const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
        defer zlsx_editor_close(ed);
        var diag = freshDiag();
        const nm = "Fine";
        try std.testing.expectEqual(ZLSX_OK, zlsx_editor_rename_sheet(ed, 1, nm.ptr, nm.len, &diag, &err_buf, err_buf.len));
    }
    // A worksheet part whose payload the store cannot materialise
    // (its CRC no longer matches): the archive opens, the lazy read
    // meets it, and it is the sheet's own verdict, not a generic
    // `BadZip` (Codex #207 r4 REL-401).
    {
        const path = try writeS3aFixture(io, &tt, "s3a_lazy_crc.xlsx");
        defer alloc.free(path);
        try corruptPartPayload(alloc, io, path, "xl/worksheets/sheet2.xml");
        const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
        defer zlsx_editor_close(ed);
        var diag = freshDiag();
        try std.testing.expectEqual(ZLSX_REFUSED, zlsx_editor_delete_sheet(ed, 1, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings("MalformedSheetXml", diagName(&diag));
        try std.testing.expectEqual(plane_none, diag.plane);
        // The same verdict from a row edit on the OTHER sheet: the
        // `<xm:f>` pre-flight reads every sheet lazily (r6 REL-603).
        diag = freshDiag();
        try std.testing.expectEqual(ZLSX_REFUSED, zlsx_editor_insert_row(ed, 0, 1, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings("MalformedSheetXml", diagName(&diag));
    }
    // A `<tablePart>` whose relationship is gone, and a table part
    // without its display name: workbook damage, not a wrong selector
    // (r6 REL-604).
    {
        const path = try tt.path(alloc, io, "s3a_table_norel.xlsx");
        defer alloc.free(path);
        try zlsx_pkg.pivots.fixture.write(alloc, io, path, .table_name);
        try zlsx_pkg.pivots.fixture.patchPart(alloc, io, path, "xl/worksheets/sheet1.xml", "<tablePart r:id=\"rIdT1\"/>", "<tablePart r:id=\"rIdZZ\"/>");
        const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
        defer zlsx_editor_close(ed);
        var diag = freshDiag();
        const tbl = "SalesTbl";
        const old = "Qty";
        const new = "Quantity";
        try std.testing.expectEqual(ZLSX_REFUSED, zlsx_editor_rename_table_column(ed, tbl.ptr, tbl.len, old.ptr, old.len, new.ptr, new.len, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings("MissingRelationship", diagName(&diag));
        try std.testing.expectEqual(plane_none, diag.plane);
    }
    {
        const path = try tt.path(alloc, io, "s3a_table_noname.xlsx");
        defer alloc.free(path);
        try zlsx_pkg.pivots.fixture.write(alloc, io, path, .table_name);
        // Both spellings go — the lookup falls back from displayName
        // to name, so a part with either still resolves.
        try zlsx_pkg.pivots.fixture.patchPart(alloc, io, path, "xl/tables/table1.xml", " name=\"SalesTbl\" displayName=\"SalesTbl\"", "");
        const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
        defer zlsx_editor_close(ed);
        var diag = freshDiag();
        const tbl = "SalesTbl";
        const old = "Qty";
        const new = "Quantity";
        try std.testing.expectEqual(ZLSX_REFUSED, zlsx_editor_rename_table_column(ed, tbl.ptr, tbl.len, old.ptr, old.len, new.ptr, new.len, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings("MalformedTableXml", diagName(&diag));
        try std.testing.expectEqual(plane_none, diag.plane);
    }
    // A workbook.xml the splice cannot patch (no `</sheets>` — the open
    // parser needs only the `<sheet>` entries): the workbook's own
    // verdict, the out index untouched.
    {
        const path = try writeS3aFixture(io, &tt, "s3a_nosheets_close.xlsx");
        defer alloc.free(path);
        try zlsx_pkg.pivots.fixture.patchPart(alloc, io, path, "xl/workbook.xml", "</sheets>", "");
        const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
        defer zlsx_editor_close(ed);
        var diag = freshDiag();
        var idx: u32 = 3;
        const nm = "Fresh";
        try std.testing.expectEqual(ZLSX_REFUSED, zlsx_editor_add_sheet(ed, nm.ptr, nm.len, &idx, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings("MalformedWorkbookXml", diagName(&diag));
        try std.testing.expectEqual(no_sheet_idx, idx);
        try std.testing.expectEqual(plane_none, diag.plane);
    }
    // A sheetId the workbook has already pushed to the top of u32: the
    // next identifier does not exist, and that is the workbook's verdict
    // (Codex #207 r5 REL-501), not a trap.
    {
        const path = try writeS3aFixture(io, &tt, "s3a_sheetid_exhausted.xlsx");
        defer alloc.free(path);
        try zlsx_pkg.pivots.fixture.patchPart(alloc, io, path, "xl/workbook.xml", "sheetId=\"2\"", "sheetId=\"4294967295\"");
        const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
        defer zlsx_editor_close(ed);
        var diag = freshDiag();
        var idx: u32 = 3;
        const nm = "Fresh";
        try std.testing.expectEqual(ZLSX_REFUSED, zlsx_editor_add_sheet(ed, nm.ptr, nm.len, &idx, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings("IdSpaceExhausted", diagName(&diag));
        try std.testing.expectEqual(no_sheet_idx, idx);
        try std.testing.expectEqual(plane_none, diag.plane);
    }
    // A table part whose payload the store cannot materialise: the
    // carrier's own verdict (Codex #207 r5 REL-502) — unfolded from the
    // rename, folded into the sheet-level name by a row edit's
    // pre-flight, a refusal either way.
    {
        const path = try tt.path(alloc, io, "s3a_lazy_table.xlsx");
        defer alloc.free(path);
        try zlsx_pkg.pivots.fixture.write(alloc, io, path, .table_name);
        try corruptPartPayload(alloc, io, path, "xl/tables/table1.xml");
        const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
        defer zlsx_editor_close(ed);
        var diag = freshDiag();
        const tbl = "SalesTbl";
        const old = "Qty";
        const new = "Quantity";
        try std.testing.expectEqual(ZLSX_REFUSED, zlsx_editor_rename_table_column(ed, tbl.ptr, tbl.len, old.ptr, old.len, new.ptr, new.len, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings("MalformedTableXml", diagName(&diag));
        try std.testing.expectEqual(plane_none, diag.plane);
        diag = freshDiag();
        try std.testing.expectEqual(ZLSX_REFUSED, zlsx_editor_insert_row(ed, 0, 6, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings("RowEditUnsafeForSheet", diagName(&diag));
    }
    // A pivot part whose payload the store cannot materialise: the
    // archive opens (the central directory is intact), the read refuses
    // as `MalformedPivotXml` with nothing handed out.
    {
        const path = try tt.path(alloc, io, "s3a_lazy_pivot.xlsx");
        defer alloc.free(path);
        try zlsx_pkg.pivots.fixture.write(alloc, io, path, .sheet_ref);
        try corruptPartPayload(alloc, io, path, "xl/pivotTables/pivotTable1.xml");
        const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
        defer zlsx_editor_close(ed);
        var diag = freshDiag();
        var out_ptr: ?[*]u8 = null;
        var out_len: usize = 0;
        try std.testing.expectEqual(ZLSX_REFUSED, zlsx_editor_pivots_ndjson(ed, &out_ptr, &out_len, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings("MalformedPivotXml", diagName(&diag));
        try std.testing.expect(out_ptr == null);
        try std.testing.expectEqual(@as(usize, 0), out_len);
    }
}

test "S3a: every member of the structural vocabulary crosses as -2 with its name in the diag and errbuf" {
    var err_buf: [128]u8 = undefined;
    for (structural_refusals) |r| {
        var diag = freshDiag();
        try std.testing.expect(prepDiag(&diag, &err_buf, err_buf.len));
        try std.testing.expectEqual(ZLSX_REFUSED, failMapped(r, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings(@errorName(r), diagName(&diag));
        try std.testing.expectEqualStrings(@errorName(r), std.mem.sliceTo(&err_buf, 0));
        try std.testing.expectEqual(plane_none, diag.plane);
        try std.testing.expectEqual(@as(usize, 0), diag.census_len);
    }
}

test "S3a: the structural vocabulary maps to -2 and nothing else does" {
    try std.testing.expectEqual(ZLSX_REFUSED, statusOf(error.RowEditUnsafeForSheet));
    try std.testing.expectEqual(ZLSX_REFUSED, statusOf(error.CannotDeleteLastSheet));
    try std.testing.expectEqual(ZLSX_REFUSED, statusOf(error.MalformedPivotXml));
    try std.testing.expectEqual(ZLSX_REFUSED, statusOf(error.RowEditExceedsMaxRow));
    try std.testing.expectEqual(ZLSX_REFUSED, statusOf(error.ColEditExceedsMaxCol));
    try std.testing.expectEqual(ZLSX_REFUSED, statusOf(error.SplitPaneNotSupported));
    try std.testing.expectEqual(ZLSX_REFUSED, statusOf(error.MalformedDrawingXml));
    try std.testing.expectEqual(ZLSX_REFUSED, statusOf(error.MalformedExtensionXml));
    try std.testing.expectEqual(ZLSX_REFUSED, statusOf(error.MalformedChartXml));
    try std.testing.expectEqual(ZLSX_REFUSED, statusOf(error.DrawingOnUnlistedSheet));
    try std.testing.expectEqual(ZLSX_REFUSED, statusOf(error.MalformedCommentsXml));
    try std.testing.expectEqual(ZLSX_REFUSED, statusOf(error.MissingSheetPart));
    try std.testing.expectEqual(ZLSX_REFUSED, statusOf(error.SqrefCollapseUnsafe));
    try std.testing.expectEqual(ZLSX_REFUSED, statusOf(error.InternalSheetNameTooLong));
    try std.testing.expectEqual(ZLSX_REFUSED, statusOf(error.MalformedWorkbookXml));
    try std.testing.expectEqual(ZLSX_REFUSED, statusOf(error.IdSpaceExhausted));
    // The decompression-caps verdict is a DELIBERATE -1: it fires at
    // open, where the ABI has no diag — `zlsx_open_buffer`'s shipped
    // contract — so remapping it to -2 was an ABI break (Codex #216
    // r2 S3B-ERR-702, overruling r1 S3B-ERR-601's remap). The name
    // still crosses in errbuf.
    try std.testing.expectEqual(ZLSX_ERROR, statusOf(error.ZipBombSuspected));
    try std.testing.expectEqual(ZLSX_ERROR, statusOf(error.TableNotFound));
    try std.testing.expectEqual(ZLSX_ERROR, statusOf(error.TableColumnNotFound));
    try std.testing.expectEqual(ZLSX_ERROR, statusOf(error.InvalidTableColumnName));
    try std.testing.expectEqual(ZLSX_ERROR, statusOf(error.MalformedXml));
    try std.testing.expectEqual(ZLSX_ERROR, statusOf(error.WriteFailed));
    try std.testing.expectEqual(ZLSX_ERROR, statusOf(error.SheetHasUnsavedMutations));
    try std.testing.expectEqual(ZLSX_ERROR, statusOf(error.RowEditRequiresCleanSheet));
    try std.testing.expectEqual(ZLSX_ERROR, statusOf(error.SheetIndexOutOfRange));
    try std.testing.expectEqual(ZLSX_ERROR, statusOf(error.InvalidSheetName));
    try std.testing.expectEqual(ZLSX_NOMEM, statusOf(error.OutOfMemory));
}

// ─── S3b slice 10: sheet visibility on the reader handle ─────────────

/// Four sheets through the writer, then `state` attributes spliced into
/// `xl/workbook.xml` through the archive — the writer authors no sheet
/// state (`<sheet name=… sheetId=… r:id=…/>`), so the fixture is the
/// test's: `Ledger` hidden, `Secret` veryHidden, `Odd` an unrecognised
/// value the reader folds to visible, `Data` untouched (no attribute).
fn writeS3bSheetStateFixture(io: std.Io, tt: *TestTmp, name: []const u8) ![:0]u8 {
    const alloc = std.testing.allocator;
    const path = try tt.path(alloc, io, name);
    errdefer alloc.free(path);
    {
        var w = xlsx.Writer.init(alloc);
        defer w.deinit();
        for ([_][]const u8{ "Data", "Ledger", "Secret", "Odd" }) |sheet_name| {
            var sheet = try w.addSheet(sheet_name);
            try sheet.writeRow(&.{.{ .string = sheet_name }});
        }
        try w.save(io, path);
    }
    const patch = zlsx_pkg.pivots.fixture.patchPart;
    try patch(alloc, io, path, "xl/workbook.xml", "name=\"Ledger\"", "name=\"Ledger\" state=\"hidden\"");
    try patch(alloc, io, path, "xl/workbook.xml", "name=\"Secret\"", "name=\"Secret\" state=\"veryHidden\"");
    try patch(alloc, io, path, "xl/workbook.xml", "name=\"Odd\"", "name=\"Odd\" state=\"bogus\"");
    return path;
}

// The three codes are the header's literals (`include/zlsx.h`
// ZLSX_SHEET_STATE_*), pinned here and static-asserted in
// tests/c_abi_smoke.c so the two hand-maintained spellings cannot
// drift apart.
test "S3b sheet_state: the codes are the header's, in the reader enum's order, and the spellings are the CLI's" {
    try std.testing.expectEqual(@as(i32, 0), ZLSX_SHEET_STATE_VISIBLE);
    try std.testing.expectEqual(@as(i32, 1), ZLSX_SHEET_STATE_HIDDEN);
    try std.testing.expectEqual(@as(i32, 2), ZLSX_SHEET_STATE_VERY_HIDDEN);
    // The codes follow `xlsx.SheetState`'s declaration order, and the
    // strings py-zlsx maps them back to (`_SHEET_STATE_NAMES`) are the
    // reader's own `toString` — what `zlsx list-sheets` prints. Pinned
    // here because `zig build test` is the hard CI gate; the Python
    // parity test runs only on the best-effort Windows lane.
    try std.testing.expectEqual(@as(u32, 0), @intFromEnum(xlsx.SheetState.visible));
    try std.testing.expectEqual(@as(u32, 1), @intFromEnum(xlsx.SheetState.hidden));
    try std.testing.expectEqual(@as(u32, 2), @intFromEnum(xlsx.SheetState.very_hidden));
    try std.testing.expectEqualStrings("visible", xlsx.SheetState.visible.toString());
    try std.testing.expectEqualStrings("hidden", xlsx.SheetState.hidden.toString());
    try std.testing.expectEqualStrings("veryHidden", xlsx.SheetState.very_hidden.toString());
}

test "S3b sheet_state: the reader's <sheet state> through the C ABI — the schema default, out of range, the buffer opener" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try writeS3bSheetStateFixture(io, &tt, "s3b_sheet_state.xlsx");
    defer std.testing.allocator.free(path);

    var err_buf: [128]u8 = undefined;
    const book = zlsx_book_open(path, &err_buf, err_buf.len);
    try std.testing.expect(book != null);
    defer zlsx_book_close(book);

    // Hidden sheets stay in the inventory — the count and the names
    // are what they were before the patch.
    try std.testing.expectEqual(@as(u32, 4), zlsx_sheet_count(book.?));
    try std.testing.expectEqual(ZLSX_SHEET_STATE_VISIBLE, zlsx_sheet_state(book.?, 0)); // no attribute
    try std.testing.expectEqual(ZLSX_SHEET_STATE_HIDDEN, zlsx_sheet_state(book.?, 1));
    try std.testing.expectEqual(ZLSX_SHEET_STATE_VERY_HIDDEN, zlsx_sheet_state(book.?, 2));
    try std.testing.expectEqual(ZLSX_SHEET_STATE_VISIBLE, zlsx_sheet_state(book.?, 3)); // `bogus` folds to the default
    // Out of range is -1, never a code — the first index past the end
    // and the far end of the u32 range alike.
    try std.testing.expectEqual(@as(i32, -1), zlsx_sheet_state(book.?, 4));
    try std.testing.expectEqual(@as(i32, -1), zlsx_sheet_state(book.?, std.math.maxInt(u32)));
    // Composes with the name lookup: the veryHidden sheet is found by
    // name, its name is intact, and its one row reads through the row
    // iterator — nothing about visibility gates the data.
    const secret = zlsx_sheet_index_by_name(book.?, "Secret", 6);
    try std.testing.expectEqual(@as(i32, 2), secret);
    try std.testing.expectEqual(ZLSX_SHEET_STATE_VERY_HIDDEN, zlsx_sheet_state(book.?, @intCast(secret)));
    var name_buf: [16]u8 = undefined;
    try std.testing.expectEqual(@as(usize, 6), zlsx_sheet_name(book.?, 2, &name_buf, name_buf.len));
    try std.testing.expectEqualStrings("Secret", std.mem.sliceTo(&name_buf, 0));
    {
        const rows = zlsx_rows_open(book.?, @intCast(secret), &err_buf, err_buf.len);
        try std.testing.expect(rows != null);
        defer zlsx_rows_close(rows);
        var cells_ptr: [*]const CCell = undefined;
        var cells_len: usize = 0;
        try std.testing.expectEqual(@as(i32, 1), zlsx_rows_next(rows.?, &cells_ptr, &cells_len, &err_buf, err_buf.len));
        try std.testing.expectEqual(@as(usize, 1), cells_len);
        try std.testing.expectEqual(@intFromEnum(CellTag.string), cells_ptr[0].tag);
        try std.testing.expectEqualStrings("Secret", cells_ptr[0].str_ptr[0..cells_ptr[0].str_len]);
        try std.testing.expectEqual(@as(i32, 0), zlsx_rows_next(rows.?, &cells_ptr, &cells_len, &err_buf, err_buf.len));
    }

    // The buffer opener models the same field from the same bytes.
    var from_bytes: ?*Book = null;
    {
        const bytes = try std.Io.Dir.cwd().readFileAlloc(io, path, std.testing.allocator, .limited(1 << 24));
        from_bytes = zlsx_book_open_buffer(bytes.ptr, bytes.len, &err_buf, err_buf.len);
        std.testing.allocator.free(bytes);
    }
    try std.testing.expect(from_bytes != null);
    defer zlsx_book_close(from_bytes);
    for ([_]i32{ ZLSX_SHEET_STATE_VISIBLE, ZLSX_SHEET_STATE_HIDDEN, ZLSX_SHEET_STATE_VERY_HIDDEN, ZLSX_SHEET_STATE_VISIBLE }, 0..) |want, i| {
        try std.testing.expectEqual(want, zlsx_sheet_state(from_bytes.?, @intCast(i)));
    }
    try std.testing.expectEqual(@as(i32, -1), zlsx_sheet_state(from_bytes.?, 4));
}

test "S3b sheet_state: a fresh writer's sheets carry no attribute and read visible" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try writeS3aFixture(io, &tt, "s3b_sheet_state_fresh.xlsx");
    defer std.testing.allocator.free(path);

    var err_buf: [128]u8 = undefined;
    const book = zlsx_book_open(path, &err_buf, err_buf.len);
    try std.testing.expect(book != null);
    defer zlsx_book_close(book);
    try std.testing.expectEqual(@as(u32, 2), zlsx_sheet_count(book.?));
    try std.testing.expectEqual(ZLSX_SHEET_STATE_VISIBLE, zlsx_sheet_state(book.?, 0));
    try std.testing.expectEqual(ZLSX_SHEET_STATE_VISIBLE, zlsx_sheet_state(book.?, 1));
    try std.testing.expectEqual(@as(i32, -1), zlsx_sheet_state(book.?, 2));
}

// ─── S3b slice 11: formula text and error tags on the row iterator ───

/// Rows 2–6 spliced into the writer's sheet part before `</sheetData>`
/// — the writer authors neither shared formulas nor `t="e"` cells, so
/// the fixture is the test's. Row 1 is the writer's (`A1` = 1).
///   row 2: A2 stand-alone `A1*2` cached 2 · B2 an entity-bearing body
///          (`"x"&amp;"y"&lt;&gt;A1`, decoded `"x"&"y"<>A1`) cached `xy`
///          · C2 formula-only (`A1/0`, no `<v>`)
///   row 3: A3 shared base `A2+1` (ref A3:B3, si 0) cached 3 · B3 its
///          slave (`<f t="shared" si="0"/>`) cached 4 · C3 `t="e"`
///          `#DIV/0!` · D3 a formula (`1/0`) whose cached value is the
///          same literal — the formula wins
///   row 4: A4 array base `A1*{1,2}` (ref A4:B4) cached 1 · B4 the
///          array slave (no `<f>` at all) cached 2 · C4 `t="e"` `#N/A`
///   row 5: A5 a gap · B5 `t="e"` `#REF!`
///   row 6: A6 an empty body (`<f></f>`, own text of length 0) cached 5
///          · B6 a slave whose base was never seen (`si="7"`) cached 8
///          — a value cell, the reader's rule · C6 a `t="dataTable"`
///          formula (`x`, own text like any other body) cached 7
const s3b11_rows =
    "<row r=\"2\">" ++
    "<c r=\"A2\"><f>A1*2</f><v>2</v></c>" ++
    "<c r=\"B2\" t=\"str\"><f>\"x\"&amp;\"y\"&lt;&gt;A1</f><v>xy</v></c>" ++
    "<c r=\"C2\"><f>A1/0</f></c>" ++
    "</row>" ++
    "<row r=\"3\">" ++
    "<c r=\"A3\"><f t=\"shared\" ref=\"A3:B3\" si=\"0\">A2+1</f><v>3</v></c>" ++
    "<c r=\"B3\"><f t=\"shared\" si=\"0\"/><v>4</v></c>" ++
    "<c r=\"C3\" t=\"e\"><v>#DIV/0!</v></c>" ++
    "<c r=\"D3\" t=\"e\"><f>1/0</f><v>#DIV/0!</v></c>" ++
    "</row>" ++
    "<row r=\"4\">" ++
    "<c r=\"A4\"><f t=\"array\" ref=\"A4:B4\">A1*{1,2}</f><v>1</v></c>" ++
    "<c r=\"B4\"><v>2</v></c>" ++
    "<c r=\"C4\" t=\"e\"><v>#N/A</v></c>" ++
    "</row>" ++
    "<row r=\"5\">" ++
    "<c r=\"B5\" t=\"e\"><v>#REF!</v></c>" ++
    "</row>" ++
    "<row r=\"6\">" ++
    "<c r=\"A6\"><f></f><v>5</v></c>" ++
    "<c r=\"B6\"><f t=\"shared\" si=\"7\"/><v>8</v></c>" ++
    "<c r=\"C6\"><f t=\"dataTable\" ref=\"C6:C7\" r1=\"A1\">x</f><v>7</v></c>" ++
    "</row>";

/// A row the reader cannot finish — `<c r="B7"` with no `>` — for the
/// failed-skip test: the fixture's spreads make `skipRows` decode
/// through `next()`, which tears on it.
const s3b11_torn_tail = "<row r=\"7\"><c r=\"A7\"><f>Z9</f><v>9</v></c><c r=\"B7\"";

fn writeS3bFormulaFixture(io: std.Io, tt: *TestTmp, name: []const u8) ![:0]u8 {
    return writeS3bFormulaFixtureWithTail(io, tt, name, "");
}

fn writeS3bFormulaFixtureWithTail(io: std.Io, tt: *TestTmp, name: []const u8, tail: []const u8) ![:0]u8 {
    const alloc = std.testing.allocator;
    const path = try tt.path(alloc, io, name);
    errdefer alloc.free(path);
    {
        var w = xlsx.Writer.init(alloc);
        defer w.deinit();
        var sheet = try w.addSheet("Data");
        try sheet.writeRow(&.{.{ .integer = 1 }});
        try w.save(io, path);
    }
    // `</sheetData>` occurs once in a one-sheet writer part; the rows
    // land after the writer's row 1, in order.
    const rows_xml = try std.mem.concat(alloc, u8, &.{ s3b11_rows, tail, "</sheetData>" });
    defer alloc.free(rows_xml);
    try zlsx_pkg.pivots.fixture.patchPart(alloc, io, path, "xl/worksheets/sheet1.xml", "</sheetData>", rows_xml);
    return path;
}

/// The three getters on one column, with sentinels in every out param
/// so a 1 / -1 that wrote anything is caught — the pointers are kept
/// raw and compared by identity, so a getter that wrote a different
/// pointer to the same bytes is caught too (in-house r1 S3B-MNT-105 +
/// a quick win).
const S3b11Probe = struct {
    const sentinel: [*]const u8 = @ptrCast("sentinel");

    formula: i32,
    formula_ptr: [*]const u8,
    formula_len: usize,
    ref: i32,
    ref_col: u32,
    ref_row: u32,
    err: i32,
    err_ptr: [*]const u8,
    err_len: usize,

    fn at(rows: *Rows, col: usize) S3b11Probe {
        var f_ptr: [*]const u8 = sentinel;
        var f_len: usize = 8;
        var r_col: u32 = 0xDEAD;
        var r_row: u32 = 0xBEEF;
        var e_ptr: [*]const u8 = sentinel;
        var e_len: usize = 8;
        const f = zlsx_rows_formula_at(rows, col, &f_ptr, &f_len);
        const r = zlsx_rows_formula_ref_at(rows, col, &r_col, &r_row);
        const e = zlsx_rows_error_at(rows, col, &e_ptr, &e_len);
        return .{
            .formula = f,
            .formula_ptr = f_ptr,
            .formula_len = f_len,
            .ref = r,
            .ref_col = r_col,
            .ref_row = r_row,
            .err = e,
            .err_ptr = e_ptr,
            .err_len = e_len,
        };
    }

    fn expectFormulaUntouched(self: S3b11Probe) !void {
        try std.testing.expectEqual(sentinel, self.formula_ptr);
        try std.testing.expectEqual(@as(usize, 8), self.formula_len);
    }

    fn expectRefUntouched(self: S3b11Probe) !void {
        try std.testing.expectEqual(@as(u32, 0xDEAD), self.ref_col);
        try std.testing.expectEqual(@as(u32, 0xBEEF), self.ref_row);
    }

    fn expectErrorUntouched(self: S3b11Probe) !void {
        try std.testing.expectEqual(sentinel, self.err_ptr);
        try std.testing.expectEqual(@as(usize, 8), self.err_len);
    }

    /// Every getter said "not that kind of cell" (or out of range) and
    /// left its out params alone.
    fn expectNone(self: S3b11Probe, code: i32) !void {
        try std.testing.expectEqual(code, self.formula);
        try std.testing.expectEqual(code, self.ref);
        try std.testing.expectEqual(code, self.err);
        try self.expectFormulaUntouched();
        try self.expectRefUntouched();
        try self.expectErrorUntouched();
    }

    fn expectFormula(self: S3b11Probe, text: []const u8) !void {
        try std.testing.expectEqual(@as(i32, 0), self.formula);
        try std.testing.expectEqualStrings(text, self.formula_ptr[0..self.formula_len]);
        // Written, not the sentinel — pinned by identity so an empty
        // body (len 0) still proves the pointer was set.
        try std.testing.expect(self.formula_ptr != sentinel);
        try std.testing.expectEqual(@as(i32, 1), self.ref);
        try std.testing.expectEqual(@as(i32, 1), self.err);
        try self.expectRefUntouched();
        try self.expectErrorUntouched();
    }

    fn expectSlave(self: S3b11Probe, col: u32, row: u32) !void {
        try std.testing.expectEqual(@as(i32, 1), self.formula);
        try std.testing.expectEqual(@as(i32, 0), self.ref);
        try std.testing.expectEqual(col, self.ref_col);
        try std.testing.expectEqual(row, self.ref_row);
        try std.testing.expectEqual(@as(i32, 1), self.err);
        try self.expectFormulaUntouched();
        try self.expectErrorUntouched();
    }

    fn expectError(self: S3b11Probe, literal: []const u8) !void {
        try std.testing.expectEqual(@as(i32, 1), self.formula);
        try std.testing.expectEqual(@as(i32, 1), self.ref);
        try std.testing.expectEqual(@as(i32, 0), self.err);
        try std.testing.expectEqualStrings(literal, self.err_ptr[0..self.err_len]);
        try std.testing.expect(self.err_ptr != sentinel);
        try self.expectFormulaUntouched();
        try self.expectRefUntouched();
    }
};

/// `zlsx_rows_parse_date` with a sentinel-filled out struct; `-1`
/// must leave it alone.
fn expectNoDateRow(rows: *Rows, col: usize) !void {
    var dt: CDateTime = .{ .year = 0xDEAD, .month = 1, .day = 2, .hour = 3, .minute = 4, .second = 5, ._pad = 6 };
    try std.testing.expectEqual(@as(i32, -1), zlsx_rows_parse_date(rows, col, &dt));
    try std.testing.expectEqual(@as(u16, 0xDEAD), dt.year);
    try std.testing.expectEqual(@as(u8, 5), dt.second);
}

/// All five per-column getters agree that `col` is not on the current
/// row — the trio, the style getter and the date getter answer -1 and
/// write nothing (in-house r2 S3B-DOC-206: one call per site, so the
/// claim "agreeing at each" is the test).
fn expectNoRowAt(rows: *Rows, col: usize) !void {
    try S3b11Probe.at(rows, col).expectNone(-1);
    var style: u32 = 0xABCD;
    try std.testing.expectEqual(@as(i32, -1), zlsx_rows_style_at(rows, col, &style));
    try std.testing.expectEqual(@as(u32, 0xABCD), style);
    try expectNoDateRow(rows, col);
}

fn expectCellString(cell: CCell, want: []const u8) !void {
    try std.testing.expectEqual(@intFromEnum(CellTag.string), cell.tag);
    try std.testing.expectEqualStrings(want, cell.str_ptr[0..cell.str_len]);
}

fn expectCellInt(cell: CCell, want: i64) !void {
    try std.testing.expectEqual(@intFromEnum(CellTag.integer), cell.tag);
    try std.testing.expectEqual(want, cell.i);
}

test "S3b rows formulas: the reader's three side channels through the C ABI — text, base ref, error literal, the precedence rule, the cached value" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try writeS3bFormulaFixture(io, &tt, "s3b_formulas.xlsx");
    defer std.testing.allocator.free(path);

    var err_buf: [128]u8 = undefined;
    const book = zlsx_book_open(path, &err_buf, err_buf.len);
    try std.testing.expect(book != null);
    defer zlsx_book_close(book);
    const rows = zlsx_rows_open(book.?, 0, &err_buf, err_buf.len);
    try std.testing.expect(rows != null);
    defer zlsx_rows_close(rows);
    var cells_ptr: [*]const CCell = undefined;
    var cells_len: usize = 0;

    // Before the first row there is no current row: -1, nothing written
    // — every getter on the handle agrees.
    try expectNoRowAt(rows.?, 0);

    // Row 1 — the writer's value cell: the trio says 1; past the end
    // of the row every getter on the handle says -1.
    try std.testing.expectEqual(@as(i32, 1), zlsx_rows_next(rows.?, &cells_ptr, &cells_len, &err_buf, err_buf.len));
    try std.testing.expectEqual(@as(usize, 1), cells_len);
    try expectCellInt(cells_ptr[0], 1);
    try S3b11Probe.at(rows.?, 0).expectNone(1);
    try expectNoRowAt(rows.?, 1);
    try expectNoRowAt(rows.?, std.math.maxInt(usize));

    // Row 2 — stand-alone formulas: the text is the `<f>` body,
    // entity-decoded; the cell is the cached `<v>`, or empty for a
    // formula-only cell.
    try std.testing.expectEqual(@as(i32, 1), zlsx_rows_next(rows.?, &cells_ptr, &cells_len, &err_buf, err_buf.len));
    try std.testing.expectEqual(@as(usize, 3), cells_len);
    try S3b11Probe.at(rows.?, 0).expectFormula("A1*2");
    try expectCellInt(cells_ptr[0], 2);
    try S3b11Probe.at(rows.?, 1).expectFormula("\"x\"&\"y\"<>A1");
    try expectCellString(cells_ptr[1], "xy");
    try S3b11Probe.at(rows.?, 2).expectFormula("A1/0");
    try std.testing.expectEqual(@intFromEnum(CellTag.empty), cells_ptr[2].tag);
    try expectNoRowAt(rows.?, 3);

    // Row 3 — a shared base and its slave, an error cell, and a formula
    // whose cached value is an error literal (the formula wins; the
    // cell array carries the literal as a string either way).
    try std.testing.expectEqual(@as(i32, 1), zlsx_rows_next(rows.?, &cells_ptr, &cells_len, &err_buf, err_buf.len));
    try std.testing.expectEqual(@as(usize, 4), cells_len);
    try S3b11Probe.at(rows.?, 0).expectFormula("A2+1");
    try expectCellInt(cells_ptr[0], 3);
    try S3b11Probe.at(rows.?, 1).expectSlave(0, 3);
    try expectCellInt(cells_ptr[1], 4);
    try S3b11Probe.at(rows.?, 2).expectError("#DIV/0!");
    try expectCellString(cells_ptr[2], "#DIV/0!");
    try S3b11Probe.at(rows.?, 3).expectFormula("1/0");
    try expectCellString(cells_ptr[3], "#DIV/0!");
    try expectNoRowAt(rows.?, 4);

    // Row 4 — an array base and the slave inside its rectangle, which
    // has no `<f>` of its own.
    try std.testing.expectEqual(@as(i32, 1), zlsx_rows_next(rows.?, &cells_ptr, &cells_len, &err_buf, err_buf.len));
    try std.testing.expectEqual(@as(usize, 3), cells_len);
    try S3b11Probe.at(rows.?, 0).expectFormula("A1*{1,2}");
    try expectCellInt(cells_ptr[0], 1);
    try S3b11Probe.at(rows.?, 1).expectSlave(0, 4);
    try expectCellInt(cells_ptr[1], 2);
    try S3b11Probe.at(rows.?, 2).expectError("#N/A");
    try expectNoRowAt(rows.?, 3);

    // Row 5 — a gap cell is a value cell to the trio; the error
    // literal keeps its bytes.
    try std.testing.expectEqual(@as(i32, 1), zlsx_rows_next(rows.?, &cells_ptr, &cells_len, &err_buf, err_buf.len));
    try std.testing.expectEqual(@as(usize, 2), cells_len);
    try std.testing.expectEqual(@intFromEnum(CellTag.empty), cells_ptr[0].tag);
    try S3b11Probe.at(rows.?, 0).expectNone(1);
    try S3b11Probe.at(rows.?, 1).expectError("#REF!");
    try expectCellString(cells_ptr[1], "#REF!");
    try expectNoRowAt(rows.?, 2);

    // Row 6 — the header's edge promises: an empty `<f></f>` is own
    // text of length 0 behind a written (non-sentinel) pointer; a slave
    // whose base was never seen is a value cell; `t="dataTable"` is
    // own text like any other body.
    try std.testing.expectEqual(@as(i32, 1), zlsx_rows_next(rows.?, &cells_ptr, &cells_len, &err_buf, err_buf.len));
    try std.testing.expectEqual(@as(usize, 3), cells_len);
    try S3b11Probe.at(rows.?, 0).expectFormula("");
    try expectCellInt(cells_ptr[0], 5);
    try S3b11Probe.at(rows.?, 1).expectNone(1);
    try expectCellInt(cells_ptr[1], 8);
    try S3b11Probe.at(rows.?, 2).expectFormula("x");
    try expectCellInt(cells_ptr[2], 7);
    try expectNoRowAt(rows.?, 3);

    // Past the end there is no current row again — for every getter
    // on the handle alike.
    try std.testing.expectEqual(@as(i32, 0), zlsx_rows_next(rows.?, &cells_ptr, &cells_len, &err_buf, err_buf.len));
    try expectNoRowAt(rows.?, 0);
}

test "S3b rows formulas: a skip clears the current row for every side-channel getter, and the buffer opener agrees" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try writeS3bFormulaFixture(io, &tt, "s3b_formulas_skip.xlsx");
    defer std.testing.allocator.free(path);

    var err_buf: [128]u8 = undefined;
    var book: ?*Book = null;
    {
        const bytes = try std.Io.Dir.cwd().readFileAlloc(io, path, std.testing.allocator, .limited(1 << 24));
        book = zlsx_book_open_buffer(bytes.ptr, bytes.len, &err_buf, err_buf.len);
        std.testing.allocator.free(bytes);
    }
    try std.testing.expect(book != null);
    defer zlsx_book_close(book);
    const rows = zlsx_rows_open(book.?, 0, &err_buf, err_buf.len);
    try std.testing.expect(rows != null);
    defer zlsx_rows_close(rows);
    var cells_ptr: [*]const CCell = undefined;
    var cells_len: usize = 0;

    // Row 2 has a formula in every column; after skipping row 3 nothing
    // is current — the reader's per-row lists still hold the last
    // decoded row (the fixture has formula spreads, so the skip decoded
    // it: row 3, four slots wide), but no getter serves it.
    try std.testing.expectEqual(@as(i32, 1), zlsx_rows_next(rows.?, &cells_ptr, &cells_len, &err_buf, err_buf.len));
    try std.testing.expectEqual(@as(i32, 1), zlsx_rows_next(rows.?, &cells_ptr, &cells_len, &err_buf, err_buf.len));
    try S3b11Probe.at(rows.?, 0).expectFormula("A1*2");
    var skipped: usize = 0;
    try std.testing.expectEqual(@as(i32, 0), zlsx_rows_skip(rows.?, 1, &skipped, &err_buf, err_buf.len));
    try std.testing.expectEqual(@as(usize, 1), skipped);
    // The premise that makes the -1 discriminate against the old
    // reader-list bound (in-house r3 S3B-MNT-304): row 3's four slots
    // (row 2 had three) sit stale behind the empty view.
    try std.testing.expectEqual(@as(usize, 4), rows_inner(rows.?).row_cells.items.len);
    try expectNoRowAt(rows.?, 0);

    // Row 4 lands next — the array slave still resolves to its base
    // (the spread state survives the skip) and the error literal reads.
    try std.testing.expectEqual(@as(i32, 1), zlsx_rows_next(rows.?, &cells_ptr, &cells_len, &err_buf, err_buf.len));
    try std.testing.expectEqual(@as(usize, 3), cells_len);
    try S3b11Probe.at(rows.?, 0).expectFormula("A1*{1,2}");
    try S3b11Probe.at(rows.?, 1).expectSlave(0, 4);
    try S3b11Probe.at(rows.?, 2).expectError("#N/A");
    try expectCellString(cells_ptr[2], "#N/A");
}

test "S3b rows formulas: a fresh writer's value cells answer 1 on every getter" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try writeS3aFixture(io, &tt, "s3b_formulas_fresh.xlsx");
    defer std.testing.allocator.free(path);

    var err_buf: [128]u8 = undefined;
    const book = zlsx_book_open(path, &err_buf, err_buf.len);
    try std.testing.expect(book != null);
    defer zlsx_book_close(book);
    const rows = zlsx_rows_open(book.?, 0, &err_buf, err_buf.len);
    try std.testing.expect(rows != null);
    defer zlsx_rows_close(rows);
    var cells_ptr: [*]const CCell = undefined;
    var cells_len: usize = 0;
    var seen: usize = 0;
    while (zlsx_rows_next(rows.?, &cells_ptr, &cells_len, &err_buf, err_buf.len) == 1) {
        seen += 1;
        for (0..cells_len) |col| try S3b11Probe.at(rows.?, col).expectNone(1);
        try S3b11Probe.at(rows.?, cells_len).expectNone(-1);
    }
    try std.testing.expect(seen > 0);
}

test "S3b rows formulas: a skip that fails leaves no current row either (in-house r1 S3B-REL-101/102)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try writeS3bFormulaFixtureWithTail(io, &tt, "s3b_formulas_torn.xlsx", s3b11_torn_tail);
    defer std.testing.allocator.free(path);

    var err_buf: [128]u8 = undefined;
    const book = zlsx_book_open(path, &err_buf, err_buf.len);
    try std.testing.expect(book != null);
    defer zlsx_book_close(book);
    const rows = zlsx_rows_open(book.?, 0, &err_buf, err_buf.len);
    try std.testing.expect(rows != null);
    defer zlsx_rows_close(rows);
    var cells_ptr: [*]const CCell = undefined;
    var cells_len: usize = 0;

    // Row 2 is current (three formula cells); a skip long enough to
    // reach the torn row 7 decodes through `next()` (the fixture has
    // spreads), which resets the reader's lists and tears — the view
    // must not bound the getters on row 2 over row 7's remains.
    try std.testing.expectEqual(@as(i32, 1), zlsx_rows_next(rows.?, &cells_ptr, &cells_len, &err_buf, err_buf.len));
    try std.testing.expectEqual(@as(i32, 1), zlsx_rows_next(rows.?, &cells_ptr, &cells_len, &err_buf, err_buf.len));
    try std.testing.expectEqual(@as(usize, 3), cells_len);
    try S3b11Probe.at(rows.?, 0).expectFormula("A1*2");
    var skipped: usize = 0xAAAA;
    @memset(&err_buf, 0);
    try std.testing.expectEqual(@as(i32, -1), zlsx_rows_skip(rows.?, 10, &skipped, &err_buf, err_buf.len));
    try std.testing.expectEqualStrings("MalformedXml", std.mem.sliceTo(&err_buf, 0));
    try std.testing.expectEqual(@as(usize, 0xAAAA), skipped);
    try expectNoRowAt(rows.?, 0);
    // The reader's lists do hold the torn row's remains — the view is
    // what keeps them off the ABI.
    try std.testing.expect(rows_inner(rows.?).formulaStrings().len > 0);
    // And no later call resurrects a row: the tear reports again or
    // the sheet ends, never a stale 1.
    try std.testing.expect(zlsx_rows_next(rows.?, &cells_ptr, &cells_len, &err_buf, err_buf.len) != 1);
    try expectNoRowAt(rows.?, 0);

    // `zlsx_rows_next` itself on the torn row: a fresh iterator reads
    // rows 1–6, then answers -1 with the same diagnostic and no current
    // row — the view is emptied before the reader is asked, not after
    // (in-house r3 S3B-MNT-303: the -1 half of the rule, driven
    // directly).
    {
        const fresh = zlsx_rows_open(book.?, 0, &err_buf, err_buf.len);
        try std.testing.expect(fresh != null);
        defer zlsx_rows_close(fresh);
        for (0..6) |_| try std.testing.expectEqual(@as(i32, 1), zlsx_rows_next(fresh.?, &cells_ptr, &cells_len, &err_buf, err_buf.len));
        try S3b11Probe.at(fresh.?, 2).expectFormula("x");
        @memset(&err_buf, 0);
        try std.testing.expectEqual(@as(i32, -1), zlsx_rows_next(fresh.?, &cells_ptr, &cells_len, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings("MalformedXml", std.mem.sliceTo(&err_buf, 0));
        try expectNoRowAt(fresh.?, 0);
    }
}

fn rows_inner(rows: *Rows) *const xlsx.Rows {
    const rs: *RowsState = @ptrCast(@alignCast(rows));
    return &rs.inner;
}

test "S3b rows formulas: the date getter answers for the current row only, through a fast-path skip (in-house r1 S3B-REL-103/104)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(std.testing.allocator, io, "s3b_formulas_date_skip.xlsx");
    defer std.testing.allocator.free(path);
    {
        // No formula spreads anywhere, so `zlsx_rows_skip` takes the
        // fast path that leaves the reader's per-row lists untouched.
        var w = xlsx.Writer.init(std.testing.allocator);
        defer w.deinit();
        const date_style = try w.addStyle(.{ .number_format = "yyyy-mm-dd" });
        var sheet = try w.addSheet("S");
        try sheet.writeRowStyled(&.{ .{ .number = 44927 }, .{ .string = "one" } }, &.{ date_style, 0 });
        try sheet.writeRow(&.{ .{ .string = "two" }, .{ .string = "two" }, .{ .string = "two" } });
        try sheet.writeRow(&.{.{ .string = "three" }});
        try w.save(io, path);
    }

    var err_buf: [128]u8 = undefined;
    const book = zlsx_book_open(path, &err_buf, err_buf.len);
    try std.testing.expect(book != null);
    defer zlsx_book_close(book);
    const rows = zlsx_rows_open(book.?, 0, &err_buf, err_buf.len);
    try std.testing.expect(rows != null);
    defer zlsx_rows_close(rows);
    var cells_ptr: [*]const CCell = undefined;
    var cells_len: usize = 0;

    try expectNoRowAt(rows.?, 0);
    try std.testing.expectEqual(@as(i32, 1), zlsx_rows_next(rows.?, &cells_ptr, &cells_len, &err_buf, err_buf.len));
    var dt: CDateTime = undefined;
    try std.testing.expectEqual(@as(i32, 0), zlsx_rows_parse_date(rows.?, 0, &dt));
    try std.testing.expectEqual(@as(u16, 2023), dt.year);
    try std.testing.expectEqual(@as(i32, 1), zlsx_rows_parse_date(rows.?, 1, &dt));
    try expectNoRowAt(rows.?, 2);

    // A zero-length skip is a no-op: the row stays current.
    var zero_skipped: usize = 7;
    try std.testing.expectEqual(@as(i32, 0), zlsx_rows_skip(rows.?, 0, &zero_skipped, &err_buf, err_buf.len));
    try std.testing.expectEqual(@as(usize, 0), zero_skipped);
    try std.testing.expectEqual(@as(i32, 0), zlsx_rows_parse_date(rows.?, 0, &dt));
    try std.testing.expectEqual(@as(u16, 2023), dt.year);

    // Skip row 2 on the fast path: the reader's `row_cells` still hold
    // row 1 (the date — two cells, not row 2's three), but nothing is
    // current to any getter.
    var skipped: usize = 0;
    try std.testing.expectEqual(@as(i32, 0), zlsx_rows_skip(rows.?, 1, &skipped, &err_buf, err_buf.len));
    try std.testing.expectEqual(@as(usize, 1), skipped);
    try std.testing.expectEqual(@as(usize, 2), rows_inner(rows.?).row_cells.items.len);
    try expectNoRowAt(rows.?, 0);

    // Row 3 lands: one string cell, no date; past its end -1.
    try std.testing.expectEqual(@as(i32, 1), zlsx_rows_next(rows.?, &cells_ptr, &cells_len, &err_buf, err_buf.len));
    try std.testing.expectEqual(@as(usize, 1), cells_len);
    try std.testing.expectEqual(@as(i32, 1), zlsx_rows_parse_date(rows.?, 0, &dt));
    try expectNoRowAt(rows.?, 1);
    try std.testing.expectEqual(@as(i32, 0), zlsx_rows_next(rows.?, &cells_ptr, &cells_len, &err_buf, err_buf.len));
    try expectNoRowAt(rows.?, 0);
}

test "chart sweep: a chart part the walk cannot read refuses every structural edit — MalformedChartXml, folded on row / column edits" {
    const alloc = std.testing.allocator;
    var threaded: std.Io.Threaded = .init(alloc, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    var err_buf: [128]u8 = undefined;

    const path = try writeS3bAnchorsFixture(io, &tt, "chart_sweep_bad.xlsx");
    defer alloc.free(path);
    try zlsx_pkg.anchors_ndjson.fixture.patchPart(alloc, io, path, "xl/charts/chart1.xml", "<c:f>Data!$B$1</c:f>", "<c:f>Data!<!-- x -->$B$1</c:f>");
    const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
    defer zlsx_editor_close(ed);

    const nm = "Facts";
    {
        var diag = freshDiag();
        try std.testing.expectEqual(ZLSX_REFUSED, zlsx_editor_rename_sheet(ed, 0, nm.ptr, nm.len, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings("MalformedChartXml", diagName(&diag));
        try std.testing.expectEqualStrings("MalformedChartXml", std.mem.sliceTo(&err_buf, 0));
        try std.testing.expectEqual(plane_none, diag.plane);
    }
    {
        var diag = freshDiag();
        try std.testing.expectEqual(ZLSX_REFUSED, zlsx_editor_delete_sheet(ed, 0, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings("MalformedChartXml", diagName(&diag));
    }
    {
        var diag = freshDiag();
        try std.testing.expectEqual(ZLSX_REFUSED, zlsx_editor_insert_row(ed, 0, 1, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings("RowEditUnsafeForSheet", diagName(&diag));
    }
    {
        var diag = freshDiag();
        try std.testing.expectEqual(ZLSX_REFUSED, zlsx_editor_insert_column(ed, 1, 1, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings("ColEditUnsafeForSheet", diagName(&diag));
    }
    // The read still serves the inventory it can: the same carrier is
    // the anchors read's own refusal (`MalformedDrawingXml`), so the
    // two surfaces agree on what a readable chart formula is.
    {
        var diag = freshDiag();
        var out_ptr: ?[*]u8 = null;
        var out_len: usize = 0;
        try std.testing.expectEqual(ZLSX_REFUSED, zlsx_editor_anchors_ndjson(ed, &out_ptr, &out_len, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings("MalformedDrawingXml", diagName(&diag));
    }
}

test "namespace-aware drawings: openpyxl's default-namespace drawing is listed and its anchor moves with a row insert; an unfollowable binding refuses the read and the edit" {
    const alloc = std.testing.allocator;
    var threaded: std.Io.Threaded = .init(alloc, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    var err_buf: [128]u8 = undefined;
    // Committed (`scripts/gen_openpyxl_chart_fixture.py`): `<wsDr
    // xmlns="…/spreadsheetDrawing"><oneCellAnchor><from>` — the drawing
    // the read used to list nothing for and the `xdr:`-literal sweep
    // left in place. `from` is 1-based on the wire (D2).
    const listed = "{\"kind\":\"chart_anchor\",\"sheet\":\"Data\",\"sheet_idx\":0,\"part\":\"xl/charts/chart1.xml\"," ++
        "\"anchor\":\"one_cell\",\"from\":{\"row\":2,\"col\":4,\"row_off\":0,\"col_off\":0},\"to\":null," ++
        "\"absolute\":null,\"chart_type\":\"bar\",\"series_refs\":[\"'Data'!B1\",\"'Data'!$A$2:$A$4\",\"'Data'!$B$2:$B$4\"]}\n";
    const moved_path = try tt.path(alloc, io, "openpyxl_moved.xlsx");
    defer alloc.free(moved_path);
    {
        const src: [*:0]const u8 = "tests/corpus/openpyxl_chart.xlsx";
        const ed = zlsx_editor_open(src, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
        defer zlsx_editor_close(ed);
        var diag = freshDiag();
        var out_ptr: ?[*]u8 = null;
        var out_len: usize = 0;
        try std.testing.expectEqual(ZLSX_OK, zlsx_editor_anchors_ndjson(ed, &out_ptr, &out_len, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings(listed, out_ptr.?[0..out_len]);
        zlsx_buffer_release(out_ptr, out_len);
        // A row insert above the chart: the anchor AND the series
        // formulas move (the drawing sweep and the chart sweep), in
        // the same read, no save between.
        try std.testing.expectEqual(ZLSX_OK, zlsx_editor_insert_row(ed, 0, 1, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqual(ZLSX_OK, zlsx_editor_anchors_ndjson(ed, &out_ptr, &out_len, &diag, &err_buf, err_buf.len));
        const moved = out_ptr.?[0..out_len];
        try std.testing.expect(std.mem.indexOf(u8, moved, "\"from\":{\"row\":3,\"col\":4,\"row_off\":0,\"col_off\":0}") != null);
        try std.testing.expect(std.mem.indexOf(u8, moved, "\"series_refs\":[\"'Data'!B2\",\"'Data'!$A$3:$A$5\",\"'Data'!$B$3:$B$5\"]") != null);
        zlsx_buffer_release(out_ptr, out_len);
        try std.testing.expectEqual(@as(i32, 0), zlsx_editor_save(ed, moved_path.ptr, moved_path.len, &err_buf, err_buf.len));
    }
    // The saved drawing carries the moved row, still unprefixed.
    {
        var store = try zlsx_pkg.PartStore.open(alloc, io, moved_path);
        defer store.deinit();
        const drawing = (try store.part("xl/drawings/drawing1.xml")).?;
        try std.testing.expect(std.mem.indexOf(u8, drawing.bytes, "<row>2</row>") != null);
        try std.testing.expect(std.mem.indexOf(u8, drawing.bytes, "xdr:") == null);
        const ed = zlsx_editor_open(moved_path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
        defer zlsx_editor_close(ed);
        var diag = freshDiag();
        var out_ptr: ?[*]u8 = null;
        var out_len: usize = 0;
        try std.testing.expectEqual(ZLSX_OK, zlsx_editor_anchors_ndjson(ed, &out_ptr, &out_len, &diag, &err_buf, err_buf.len));
        try std.testing.expect(std.mem.indexOf(u8, out_ptr.?[0..out_len], "\"from\":{\"row\":3,\"col\":4,") != null);
        zlsx_buffer_release(out_ptr, out_len);
    }
    // A second spreadsheetDrawing binding past the resolver's limit:
    // the read refuses under strict and every row / column edit refuses
    // whole, typed MalformedDrawingXml (the dr-1 name, not folded) —
    // pre-mutation, so a rename is still admitted after it.
    {
        const rejected_path = try tt.path(alloc, io, "openpyxl_nsbind.xlsx");
        defer alloc.free(rejected_path);
        {
            var store = try zlsx_pkg.PartStore.open(alloc, io, "tests/corpus/openpyxl_chart.xlsx");
            defer store.deinit();
            try store.save(io, rejected_path);
        }
        try zlsx_pkg.anchors_ndjson.fixture.patchPart(alloc, io, rejected_path, "xl/drawings/drawing1.xml", "<wsDr ", "<wsDr xmlns:" ++ ("p" ** 101) ++ "=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" ");
        const ed = zlsx_editor_open(rejected_path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
        defer zlsx_editor_close(ed);
        var diag = freshDiag();
        var out_ptr: ?[*]u8 = null;
        var out_len: usize = 0;
        try std.testing.expectEqual(ZLSX_REFUSED, zlsx_editor_anchors_ndjson(ed, &out_ptr, &out_len, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings("MalformedDrawingXml", diagName(&diag));
        try std.testing.expectEqual(ZLSX_REFUSED, zlsx_editor_insert_row(ed, 0, 1, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings("MalformedDrawingXml", diagName(&diag));
        try std.testing.expectEqualStrings("MalformedDrawingXml", std.mem.sliceTo(&err_buf, 0));
        try std.testing.expectEqual(ZLSX_REFUSED, zlsx_editor_insert_column(ed, 0, 0, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings("MalformedDrawingXml", diagName(&diag));
        const nm = "Facts";
        try std.testing.expectEqual(ZLSX_OK, zlsx_editor_rename_sheet(ed, 0, nm.ptr, nm.len, &diag, &err_buf, err_buf.len));
    }
}

test "namespace-aware drawings: the default-namespace fixture hands over the prefixed fixture's bytes — every anchor kind, through the C ABI" {
    const alloc = std.testing.allocator;
    var threaded: std.Io.Threaded = .init(alloc, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    var err_buf: [128]u8 = undefined;
    // Report's drawing in openpyxl's spelling (a two-cell image, an
    // absolute image, a one-cell chart, all unprefixed) beside Data's
    // `xdr:` one: the frozen `.with_absolute` bytes, byte for byte.
    const path = try tt.path(alloc, io, "s3b_anchors_default_ns.xlsx");
    defer alloc.free(path);
    try zlsx_pkg.anchors_ndjson.fixture.write(alloc, io, path, .default_namespace);
    const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
    defer zlsx_editor_close(ed);
    var diag = freshDiag();
    var out_ptr: ?[*]u8 = null;
    var out_len: usize = 0;
    try std.testing.expectEqual(ZLSX_OK, zlsx_editor_anchors_ndjson(ed, &out_ptr, &out_len, &diag, &err_buf, err_buf.len));
    try std.testing.expectEqualStrings(s3b_anchors_frozen, out_ptr.?[0..out_len]);
    zlsx_buffer_release(out_ptr, out_len);
    // A row insert on Report moves its cell-anchored image and chart
    // and leaves the absolute image where it is.
    try std.testing.expectEqual(ZLSX_OK, zlsx_editor_insert_row(ed, 1, 1, &diag, &err_buf, err_buf.len));
    try std.testing.expectEqual(ZLSX_OK, zlsx_editor_anchors_ndjson(ed, &out_ptr, &out_len, &diag, &err_buf, err_buf.len));
    const moved = out_ptr.?[0..out_len];
    try std.testing.expect(std.mem.indexOf(u8, moved, "\"from\":{\"row\":4,\"col\":2,\"row_off\":0,\"col_off\":9525},\"to\":{\"row\":9,\"col\":5,\"row_off\":19050,\"col_off\":0}") != null);
    try std.testing.expect(std.mem.indexOf(u8, moved, "\"absolute\":{\"x\":1000,\"y\":2000,\"cx\":914400,\"cy\":457200}") != null);
    try std.testing.expect(std.mem.indexOf(u8, moved, "\"anchor\":\"one_cell\",\"from\":{\"row\":3,\"col\":6,") != null);
    zlsx_buffer_release(out_ptr, out_len);
}

// ─── S3c slice 1: the embedding write (zlsx_status_v1) ────────────────
//
// `Workbook.setEmbeddings` crosses the boundary on the editor handle:
// one call writes the whole coverage set — the index, a vec / hashes
// part per coverage, the workbook→index relationship, the recovery
// record in its two invisible carriers — and replaces any previous
// set, the Zig surface's contract. Vectors cross as the f32
// `[rows][dim]` matrix `zlsx_emb_vectors` hands back and hashes as the
// u64 per-row list `zlsx_emb_hashes` hands back, so read → re-embed →
// write is symmetric on one shape; the on-disk encoding (int8-sym's
// per-row scale + codes) is `embedding_part.encodeVectorRecord`, the
// one encoder `zlsx embed --vectors` writes through.

/// `zlsx_emb_coverage_v1` (§18). An array element — no `struct_size`
/// (the `zlsx_formula_cell_v1` precedent); the v1 layout is frozen at
/// 88 bytes.
const CEmbCoverage = extern struct {
    id: ?[*]const u8,
    id_len: usize,
    range: ?[*]const u8,
    range_len: usize,
    column: ?[*]const u8,
    column_len: usize,
    /// `rows * dim` f32, row-major, range order.
    vectors: ?[*]const f32,
    vectors_len: usize,
    /// `rows` u64; `zlsx_emb_tombstone()` marks a row with no vector.
    hashes: ?[*]const u64,
    hashes_len: usize,
    /// 0-based, the `zlsx_editor_*` convention.
    sheet_idx: u32,
    /// 0 or 1.
    include_formulas: u32,
};

comptime {
    const assert = std.debug.assert;
    assert(@offsetOf(CEmbCoverage, "id") == 0);
    assert(@offsetOf(CEmbCoverage, "id_len") == 8);
    assert(@offsetOf(CEmbCoverage, "range") == 16);
    assert(@offsetOf(CEmbCoverage, "range_len") == 24);
    assert(@offsetOf(CEmbCoverage, "column") == 32);
    assert(@offsetOf(CEmbCoverage, "column_len") == 40);
    assert(@offsetOf(CEmbCoverage, "vectors") == 48);
    assert(@offsetOf(CEmbCoverage, "vectors_len") == 56);
    assert(@offsetOf(CEmbCoverage, "hashes") == 64);
    assert(@offsetOf(CEmbCoverage, "hashes_len") == 72);
    assert(@offsetOf(CEmbCoverage, "sheet_idx") == 80);
    assert(@offsetOf(CEmbCoverage, "include_formulas") == 84);
    assert(@sizeOf(CEmbCoverage) == 88);
}

/// `zlsx_prune_report_v1` — what `zlsx_editor_prune_embeddings` did
/// (`Workbook.PruneReport`), as `uint64_t` counts so a caller never
/// narrows a slot count; `struct_size` first, the §3 rule. The four
/// fields of the `{"kind":"prune",…}` record `zlsx embed --prune`
/// prints, in its order.
const CPruneReport = extern struct {
    struct_size: usize,
    redacted: u64,
    stale: u64,
    fresh: u64,
    valid_empty: u64,
};

comptime {
    const assert = std.debug.assert;
    assert(@offsetOf(CPruneReport, "redacted") == 8);
    assert(@offsetOf(CPruneReport, "stale") == 16);
    assert(@offsetOf(CPruneReport, "fresh") == 24);
    assert(@offsetOf(CPruneReport, "valid_empty") == 32);
    assert(@sizeOf(CPruneReport) == 40);
}

/// `zlsx_editor_set_embeddings`'s flags word. v1 defines no bit, so 0
/// is the only accepted value (a set bit is `InvalidInput`); reserved
/// so `recovery_in_cells` can cross later without a second export.
const emb_write_flags_known: u32 = 0;

/// A `(ptr, len)` array at the boundary, `bytesArg`'s rule on a typed
/// element: NULL with a non-zero length is `InvalidInput`; NULL with
/// length 0 is the empty array.
fn arrayArg(comptime T: type, ptr: ?[*]const T, len: usize, err_buf: ?[*]u8, err_buf_len: usize) ?[]const T {
    if (ptr) |p| return p[0..len];
    if (len == 0) return &[_]T{};
    writeError(err_buf, err_buf_len, "InvalidInput");
    return null;
}

/// Write the workbook's embedding set — `Workbook.setEmbeddings` on
/// the editor handle: `model` / `dim` / `dtype` and `coverages_len`
/// coverages, each a sheet index, an A1 `range`, the `column` inside
/// it, `rows * dim` f32 vectors and `rows` u64 hashes (rows being the
/// range's row count; a row with no vector carries
/// `zlsx_emb_tombstone()`). Replaces any previous set (a vanished
/// coverage id's parts stay in the archive as orphans, the Zig
/// contract). Staged in memory; `zlsx_editor_save` commits. `dtype`
/// is spelled as `zlsx_emb_dtype` spells it — `"f32"` or
/// `"int8-sym-per-vec"`; the three other names the read knows have
/// no writer (`UnsupportedDtype`). `flags` is reserved and must be 0.
///
/// -1, a statement about the call, each before the first part write:
/// `InvalidInput` (NULL handle, NULL bytes or arrays with a non-zero
/// length, a set flag, `include_formulas` past 1),
/// `InvalidEmbeddingInput` (no coverage, `dim == 0`, a `vectors_len`
/// or `hashes_len` that disagrees with the range), `InvalidDtype`,
/// `UnsupportedDtype`, `SheetIndexOutOfRange`, and the index's own
/// rules — `InvalidCoverageId`, `InvalidRange` (the range, or a column
/// outside it), `DuplicateCoverageId`, `CoverageOverlap`,
/// `InvalidXmlByte` (a C0 control byte in the model). -2, a statement
/// about the workbook: `MissingWorkbookRels` / `MalformedWorkbookRels`
/// (the workbook→index relationship, or the docProps carrier's package
/// relationship, has nowhere to land — checked before the first
/// write), `IdSpaceExhausted` (the rels file's `rId` space, or an
/// existing `docProps/custom.xml`'s `pid` space, already at
/// `UINT32_MAX` — checked before the first write),
/// `MissingRelationship` (the sheet's part is unreachable),
/// `EmbeddingExceedsArchiveLimit` (a part past the 512 MiB read cap,
/// sized here from the inputs before anything is read, OR a recovery
/// record past its 16 × 200-byte ceiling — roughly eighty coverages,
/// or a ~3 KB model — encoded before the first write), and the
/// package's own `MissingContentTypes` / `MalformedContentTypes`, and
/// `MalformedWorkbookXml` (an `xl/workbook.xml` the open admits but the
/// strip of the previous record's chunk names cannot walk — a
/// `<definedName` outside `<definedNames>` the scanner refuses; judged
/// before the first write). A -2 or -3 that fires AFTER the first part
/// write (an allocation failure, an index past the cap, a content
/// types or docProps part the carriers cannot patch) leaves the
/// staged part set partially replaced: discard the editor without
/// saving. A save after this write re-emits the workbook's
/// `<definedNames>` block: every existing name keeps `name`,
/// `localSheetId` and `hidden` only — its other attributes (`comment`,
/// `description`, `function`, `vbProcedure`, …) are dropped, as after
/// any staged defined-name edit (pre-existing, recorded).
/// The recalc transactions (`zlsx_editor_mark_recalc_on_load`
/// + save, `zlsx_editor_save_with_recalc`, `zlsx_editor_recalculate`)
/// rebuild their candidate from the archive as opened and do NOT
/// carry this write — call them before it, or save and re-open (a
/// recorded, pre-existing rule of the transaction's generation model).
export fn zlsx_editor_set_embeddings(
    ed: ?*Editor,
    model: ?[*]const u8,
    model_len: usize,
    dim: u32,
    dtype: ?[*]const u8,
    dtype_len: usize,
    coverages: ?[*]const CEmbCoverage,
    coverages_len: usize,
    flags: u32,
    diag: ?*CDiag,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) i32 {
    if (!prepDiag(diag, err_buf, err_buf_len)) return ZLSX_ERROR;
    const state = editorStateOrNull(ed, err_buf, err_buf_len) orelse return ZLSX_ERROR;
    if (flags & ~emb_write_flags_known != 0) {
        writeError(err_buf, err_buf_len, "InvalidInput");
        return ZLSX_ERROR;
    }
    const model_s = bytesArg(model, model_len, err_buf, err_buf_len) orelse return ZLSX_ERROR;
    const dtype_s = bytesArg(dtype, dtype_len, err_buf, err_buf_len) orelse return ZLSX_ERROR;
    const covs = arrayArg(CEmbCoverage, coverages, coverages_len, err_buf, err_buf_len) orelse return ZLSX_ERROR;
    const dt = zlsx_pkg.embedding_part.Dtype.fromString(dtype_s) catch |e| {
        return failMapped(e, diag, err_buf, err_buf_len);
    };
    // The Zig write's own first checks, ahead of encoding a body for
    // a call that cannot land (the Zig layer re-checks: one rule, two
    // callers).
    if (covs.len == 0 or dim == 0) {
        writeError(err_buf, err_buf_len, "InvalidEmbeddingInput");
        return ZLSX_ERROR;
    }
    const wb = &state.inner.workbook;
    const inputs = gpa.alloc(zlsx_pkg.EmbeddingCoverageInput, covs.len) catch {
        writeError(err_buf, err_buf_len, "OutOfMemory");
        return ZLSX_NOMEM;
    };
    defer gpa.free(inputs);
    var encoded: usize = 0;
    defer {
        for (inputs[0..encoded]) |in| gpa.free(in.vec_body);
    }
    for (covs, 0..) |c, i| {
        if (c.include_formulas > 1) {
            writeError(err_buf, err_buf_len, "InvalidInput");
            return ZLSX_ERROR;
        }
        const id = bytesArg(c.id, c.id_len, err_buf, err_buf_len) orelse return ZLSX_ERROR;
        const range = bytesArg(c.range, c.range_len, err_buf, err_buf_len) orelse return ZLSX_ERROR;
        const column = bytesArg(c.column, c.column_len, err_buf, err_buf_len) orelse return ZLSX_ERROR;
        const vectors = arrayArg(f32, c.vectors, c.vectors_len, err_buf, err_buf_len) orelse return ZLSX_ERROR;
        const hashes = arrayArg(u64, c.hashes, c.hashes_len, err_buf, err_buf_len) orelse return ZLSX_ERROR;
        const ws = wb.sheet(c.sheet_idx) catch |e| return failMapped(e, diag, err_buf, err_buf_len);
        const target = ws.embeddingTarget() catch |e| return failMapped(e, diag, err_buf, err_buf_len);
        const parsed = zlsx_pkg.embedding_part.parseA1Range(range) catch |e| {
            return failMapped(e, diag, err_buf, err_buf_len);
        };
        const rows: usize = parsed.rowCount();
        // The call's own lengths first (-1), then the workbook's cap
        // (-2): both O(1), nothing read (in-house r2 S3C-REL-202).
        if (vectors.len != rows * @as(usize, dim) or hashes.len != rows) {
            writeError(err_buf, err_buf_len, "InvalidEmbeddingInput");
            return ZLSX_ERROR;
        }
        // The Zig write's own pass-1 part sizing, from the inputs alone
        // and before anything is read or encoded — a coverage past the
        // cap used to be encoded in full (a second 512 MiB+ allocation)
        // to be refused (in-house r1 S3C-PERF-106). Checked arithmetic:
        // an unrepresentable size is past the cap, never a trap
        // (`recordBytes` itself multiplies in usize — 64-bit on every
        // shipped target).
        const cap = zlsx_pkg.embedding_part.PART_MAX_BYTES;
        const vec_bytes = std.math.mul(usize, rows, dt.recordBytes(dim)) catch std.math.maxInt(usize);
        if (vec_bytes > cap - zlsx_pkg.embedding_part.VEC_HEADER_BYTES or
            rows * @sizeOf(u64) > cap - zlsx_pkg.embedding_part.HASH_HEADER_BYTES)
        {
            return failMapped(error.EmbeddingExceedsArchiveLimit, diag, err_buf, err_buf_len);
        }
        const body = zlsx_pkg.embedding_part.encodeVectorBody(gpa, dt, dim, vectors) catch |e| {
            return failMapped(e, diag, err_buf, err_buf_len);
        };
        inputs[i] = .{
            .id = id,
            .worksheet_target = target,
            .range = range,
            .column = column,
            .include_formulas = c.include_formulas == 1,
            .vec_body = body,
            .hashes = hashes,
        };
        encoded = i + 1;
    }
    wb.setEmbeddings(model_s, dim, dt, inputs) catch |e| {
        return failMapped(e, diag, err_buf, err_buf_len);
    };
    return ZLSX_OK;
}

/// The S3c `embed --extract` records — one `{"kind":"embed_row",…}`
/// line per row of `column` over `range` on `sheet_idx` that carries
/// embeddable content, range order — as a library-allocated UTF-8
/// buffer, byte-for-byte what `zlsx embed --extract` prints
/// (docs/cli.md, "embed --extract"): the 1-based `row`, `text` as a
/// reader sees the cell (a shared or inline string's runs joined,
/// entities resolved; a number's `<v>` as written; an error's
/// literal, `#N/A`; a boolean as `1` / `0`) and `hash`, the canonical
/// xxh3-64 content hash
/// `zlsx_editor_set_embeddings` stores beside the vector, as an
/// unsigned 64-bit decimal. `include_formulas` (0 / 1) admits formula
/// cells with a cached value — the coverage flag's reading, so the
/// read matches the coverage it feeds. Rows with nothing embeddable
/// are omitted (a covered row missing here is a `zlsx_emb_tombstone()`
/// slot on the write); a range with none is ZLSX_OK with `(NULL, 0)`.
/// Read over the editor's current parts. A sheet the editor holds
/// staged cell writes (`set_cell`, or the header cell
/// `rename_table_column` stages on the host sheet) or appended rows
/// for refuses — the parsed view
/// this read walks does not carry them, so a row would answer with
/// its saved content and a hash the staged value turns stale the
/// moment it lands: -1 `SheetHasUnsavedMutations` /
/// `SheetHasUnsavedAppends`; save and re-open, or read before the
/// writes. -1 otherwise: `InvalidInput` (NULL handle, NULL bytes with
/// a non-zero length, `include_formulas` past 1), `NullOutPointer`,
/// `InvalidRange` (the range, or a column outside it),
/// `SheetIndexOutOfRange`, `StructSizeTooSmall`. -2, the name in the
/// diag with plane NONE: `MissingRelationship` / `MissingSheetPart`
/// (the sheet's part is unreachable), `MalformedSheetXml` (a sheet
/// part the view cannot parse, or a row or cell it cannot place — no
/// `r`, or one it cannot read: 0, non-numeric, past the limit — or a
/// ref under another row), `MalformedSharedStringsXml` (a
/// table it cannot parse), and a
/// cell value the read cannot carry — `UnsupportedCellValue` (a
/// boolean `<v>` that is not `0` / `1`, a `<v>` the number
/// canonicalizer cannot read — a comma decimal, `NaN` — a `t="d"`
/// ISO-8601 date, a `t` this reader does not know, a shared-string
/// index that is not a number, an entity the decoder does not know),
/// `SstIndexOutOfRange`, `InvalidUtf8`, `UnicodeNormalizationFailed` —
/// refused whole rather than hand over a record that lies. Release
/// with `zlsx_buffer_release`.
export fn zlsx_editor_embeddable_rows_ndjson(
    ed: ?*Editor,
    sheet_idx: u32,
    range: ?[*]const u8,
    range_len: usize,
    column: ?[*]const u8,
    column_len: usize,
    include_formulas: u32,
    out_ptr: ?*?[*]u8,
    out_len: ?*usize,
    diag: ?*CDiag,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) i32 {
    // Every output prepped before anything can fail: a rejected sibling
    // leaves the accepted ones releasable.
    if (out_ptr) |op| op.* = null;
    if (out_len) |ol| ol.* = 0;
    if (!prepDiag(diag, err_buf, err_buf_len)) return ZLSX_ERROR;
    const op = out_ptr orelse {
        writeError(err_buf, err_buf_len, "NullOutPointer");
        return ZLSX_ERROR;
    };
    const ol = out_len orelse {
        writeError(err_buf, err_buf_len, "NullOutPointer");
        return ZLSX_ERROR;
    };
    const state = editorStateOrNull(ed, err_buf, err_buf_len) orelse return ZLSX_ERROR;
    if (include_formulas > 1) {
        writeError(err_buf, err_buf_len, "InvalidInput");
        return ZLSX_ERROR;
    }
    const range_s = bytesArg(range, range_len, err_buf, err_buf_len) orelse return ZLSX_ERROR;
    const column_s = bytesArg(column, column_len, err_buf, err_buf_len) orelse return ZLSX_ERROR;
    const wb = &state.inner.workbook;
    // The write's own resolution of a sheet index to the coverage's
    // `worksheet_target` (`zlsx_editor_set_embeddings`): one spelling.
    const ws = wb.sheet(sheet_idx) catch |e| return failMapped(e, diag, err_buf, err_buf_len);
    const target = ws.embeddingTarget() catch |e| return failMapped(e, diag, err_buf, err_buf_len);
    const bytes = embeddableRowsNdjsonOwned(gpa, wb, target, range_s, column_s, include_formulas == 1) catch |e| {
        return failMapped(e, diag, err_buf, err_buf_len);
    };
    if (bytes.len == 0) {
        gpa.free(bytes);
        return ZLSX_OK;
    }
    op.* = bytes.ptr;
    ol.* = bytes.len;
    return ZLSX_OK;
}

/// The embeddable-row records as one owned buffer in `alloc` — the
/// shared writer (`pkg/embeddable_row_ndjson.zig`) over
/// `Workbook.embeddableRows`, so the bytes are the CLI's. The
/// allocating writer reports a failed growth as `WriteFailed`; at this
/// boundary that is an allocation failure and crosses as `-3`, the
/// pivots builder's rule.
fn embeddableRowsNdjsonOwned(
    alloc: std.mem.Allocator,
    wb: *zlsx_pkg.Workbook,
    target: []const u8,
    range: []const u8,
    column: []const u8,
    include_formulas: bool,
) ![]u8 {
    var rows = try wb.embeddableRows(alloc, target, range, column, include_formulas);
    defer rows.deinit();
    var out: std.Io.Writer.Allocating = .init(alloc);
    defer out.deinit();
    zlsx_pkg.embeddable_rows_ndjson.writeAll(&out.writer, rows.rows) catch |e| switch (e) {
        error.WriteFailed => return error.OutOfMemory,
    };
    return out.toOwnedSlice();
}

/// The sweeps' allocating writers and `allocPrint` spell an allocation
/// failure `WriteFailed`; at this boundary that is `-3`, the
/// embeddable-rows rule (`embeddableRowsNdjsonOwned`). Nothing else on
/// either sweep's path raises the name.
fn failSweep(e: anyerror, diag: ?*CDiag, err_buf: ?[*]u8, err_buf_len: usize) i32 {
    return failMapped(if (e == error.WriteFailed) error.OutOfMemory else e, diag, err_buf, err_buf_len);
}

/// The redaction sweep — `Workbook.pruneEmbeddings` on the editor
/// handle, what `zlsx embed --prune` runs: every slot whose row is no
/// longer embeddable becomes a tombstone (`zlsx_emb_tombstone()`) and
/// its vector is zeroed; the coverage's count and range never change
/// (the format is dense — slot `i` is row `first_row + i`, always).
/// Content that drifted but is still embeddable counts `stale` and
/// is never redacted — deleting a vector because its text was EDITED
/// would lose data the caller never asked to lose; re-embed those
/// rows. `report` receives the counts — `redacted`, `stale`, `fresh`,
/// `valid_empty` (a tombstone whose row is still empty), the fields
/// of the `{"kind":"prune",…}` record the CLI prints; a workbook with
/// no embedding set, or a stripped one (the recovery record alone),
/// is ZLSX_OK with every count 0. Each row is judged through the
/// embeddable-rows read's reading (`zlsx_editor_embeddable_rows_ndjson`),
/// so a row that read fresh there reads fresh here. Staged in memory;
/// `zlsx_editor_save` commits; a set with nothing to redact rewrites
/// nothing.
///
/// A staged `zlsx_editor_set_cell` on a covered row is judged as
/// staged, never fresh (its save-time encoding is not re-derived): a
/// blank or deleted cell redacts its slot, any other value counts
/// `stale`. A covered sheet with staged `zlsx_editor_append_row` rows
/// refuses instead — -1 `SheetHasUnsavedAppends`, the read's rule: the
/// parsed view does not carry them, and a slot for a row they will
/// become would be redacted as gone — save first. -1 otherwise:
/// `InvalidInput` (NULL handle), `NullOutPointer` (NULL report),
/// `StructSizeTooSmall`. -2, the name in the diag with plane NONE,
/// each before the first part write: `MalformedEmbeddingSet` (a set
/// the index read refuses — a coverage range it cannot parse, a
/// binary part with the wrong magic, a count that disagrees with its
/// header), `MissingEmbeddingPart` (the index's rels, or a vec / hash
/// part it names, gone from the archive), and the read's own verdicts
/// on a covered cell — `MissingRelationship` / `MissingSheetPart`,
/// `MalformedSheetXml`, `MalformedSharedStringsXml`,
/// `UnsupportedCellValue`, `SstIndexOutOfRange`, `InvalidUtf8`,
/// `UnicodeNormalizationFailed` — a cell the read cannot carry stops
/// the sweep whole rather than redact a row that has text. A -3 after
/// the first part write (an allocation failure between two coverages'
/// parts) leaves the staged set partially redacted: discard the
/// editor without saving, or call again — the sweep is re-runnable.
/// The recalc transactions (`zlsx_editor_mark_recalc_on_load` + save,
/// `zlsx_editor_save_with_recalc`, `zlsx_editor_recalculate`) rebuild
/// their candidate from the archive as opened and do NOT carry this
/// sweep — call them before it, or save and re-open (the rule
/// `zlsx_editor_set_embeddings` documents).
export fn zlsx_editor_prune_embeddings(
    ed: ?*Editor,
    report: ?*CPruneReport,
    diag: ?*CDiag,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) i32 {
    // Every output prepped before anything can fail (§3): a rejected
    // sibling leaves the accepted one zeroed and releasable.
    var outputs_ok = prepDiag(diag, err_buf, err_buf_len);
    const out = report orelse {
        writeError(err_buf, err_buf_len, "NullOutPointer");
        return ZLSX_ERROR;
    };
    if (!prepOut(CPruneReport, out, err_buf, err_buf_len)) outputs_ok = false;
    if (!outputs_ok) return ZLSX_ERROR;
    const state = editorStateOrNull(ed, err_buf, err_buf_len) orelse return ZLSX_ERROR;
    const r = state.inner.workbook.pruneEmbeddings() catch |e| {
        return failSweep(e, diag, err_buf, err_buf_len);
    };
    out.redacted = r.redacted;
    out.stale = r.stale;
    out.fresh = r.fresh;
    out.valid_empty = r.valid_empty;
    return ZLSX_OK;
}

/// Remove every embedding artefact — `Workbook.stripEmbeddings` on
/// the editor handle, what `zlsx embed --strip` runs: the
/// `xl/zlsxEmbeddings/` directory whole (whatever the archive holds
/// there, not only what the index names), the workbook→index
/// relationship, and the recovery record from all three carriers (the
/// hidden `_zlsxRecovery` defined names, the `docProps/custom.xml`
/// property — the part itself only when nothing else is left in it —
/// and the `recovery_in_cells` sheet). The pre-share operation: the
/// result reads `ZLSX_EMB_ABSENT`, not `ZLSX_EMB_STRIPPED` — the
/// caller asked for the nothing. Idempotent; ZLSX_OK on a workbook
/// that never had embeddings, and on a partially stripped one. Staged
/// in memory; `zlsx_editor_save` commits.
///
/// The `recovery_in_cells` sheet, when present, goes through the
/// editor's own `zlsx_editor_delete_sheet` path so sheet indices stay
/// honest, and only then do its rules apply — judged before the first
/// part is removed: -1 `SheetDeleteRequiresCleanState` (staged cell
/// writes or appended rows on any sheet — save first), -2
/// `CannotDeleteLastSheet`, and every index above the deleted sheet's
/// shifts down by one. -1 otherwise: `InvalidInput` (NULL handle),
/// `StructSizeTooSmall`. -2, the name in the diag with plane NONE:
/// `MalformedWorkbookXml` (an `xl/workbook.xml` the strip of the
/// record's chunk names cannot walk — judged before the first
/// removal), the package's own `MissingContentTypes` /
/// `MalformedContentTypes`, and the delete's verdicts should the
/// cells sheet be present (`MissingRelationship`, a carrier the
/// cross-sheet sweeps cannot read — `MalformedChartXml`,
/// `MalformedExtensionXml`, …). A -2 or -3 after the first removal (an
/// allocation failure, a content-types part a removal cannot patch)
/// leaves the staged set partially stripped: discard the editor
/// without saving, or call again — the strip is re-runnable. The
/// recalc transactions rebuild their candidate from the archive as
/// opened and do NOT carry this strip — call them before it, or save
/// and re-open (the rule `zlsx_editor_set_embeddings` documents).
export fn zlsx_editor_strip_embeddings(
    ed: ?*Editor,
    diag: ?*CDiag,
    err_buf: ?[*]u8,
    err_buf_len: usize,
) callconv(.c) i32 {
    if (!prepDiag(diag, err_buf, err_buf_len)) return ZLSX_ERROR;
    const state = editorStateOrNull(ed, err_buf, err_buf_len) orelse return ZLSX_ERROR;
    state.inner.stripEmbeddings() catch |e| {
        return failSweep(e, diag, err_buf, err_buf_len);
    };
    return ZLSX_OK;
}

// ── S3c slice 1 tests ─────────────────────────────────────────────────

/// Two sheets — three text rows under a header on `Docs`, one row on
/// `Second` — the workbook every S3c test embeds into.
fn writeS3cFixture(io: std.Io, tt: *TestTmp, name: []const u8) ![:0]u8 {
    const alloc = std.testing.allocator;
    const path = try tt.path(alloc, io, name);
    errdefer alloc.free(path);
    var w = xlsx.Writer.init(alloc);
    defer w.deinit();
    var docs = try w.addSheet("Docs");
    try docs.writeRow(&.{ .{ .string = "title" }, .{ .string = "body" } });
    try docs.writeRow(&.{ .{ .string = "alpha" }, .{ .string = "first body" } });
    try docs.writeRow(&.{ .{ .string = "beta" }, .{ .string = "second body" } });
    try docs.writeRow(&.{ .{ .string = "gamma" }, .{ .string = "third body" } });
    var second = try w.addSheet("Second");
    try second.writeRow(&.{.{ .string = "two" }});
    try w.save(io, path);
    return path;
}

fn s3cCoverage(id: []const u8, sheet_idx: u32, range: []const u8, column: []const u8, vectors: []const f32, hashes: []const u64) CEmbCoverage {
    return .{
        .id = id.ptr,
        .id_len = id.len,
        .range = range.ptr,
        .range_len = range.len,
        .column = column.ptr,
        .column_len = column.len,
        .vectors = vectors.ptr,
        .vectors_len = vectors.len,
        .hashes = hashes.ptr,
        .hashes_len = hashes.len,
        .sheet_idx = sheet_idx,
        .include_formulas = 0,
    };
}

fn s3cSet(ed: ?*Editor, model: []const u8, dim: u32, dtype: []const u8, covs: []const CEmbCoverage, flags: u32, diag: ?*CDiag, err_buf: []u8) i32 {
    return zlsx_editor_set_embeddings(ed, model.ptr, model.len, dim, dtype.ptr, dtype.len, covs.ptr, covs.len, flags, diag, err_buf.ptr, err_buf.len);
}

fn s3cEmbString(h: *Emb, comptime getter: anytype, args: anytype, buf: []u8) []const u8 {
    const n = @call(.auto, getter, .{h} ++ args ++ .{ buf.ptr, buf.len });
    return buf[0..@min(n, buf.len - 1)];
}

test "S3c set_embeddings: f32 coverages on two sheets cross the boundary, land in the file and read back through zlsx_emb_*" {
    const alloc = std.testing.allocator;
    var threaded: std.Io.Threaded = .init(alloc, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try writeS3cFixture(io, &tt, "s3c_src.xlsx");
    defer alloc.free(path);
    const out_path = try tt.path(alloc, io, "s3c_out.xlsx");
    defer alloc.free(out_path);

    var err_buf: [128]u8 = undefined;
    const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
    defer zlsx_editor_close(ed);

    // Docs!A2:A4 (dim 3) with the middle row tombstoned, Second!A1:A1.
    const tomb = zlsx_emb_tombstone();
    const title_vecs = [_]f32{ 0.5, -1.25, 2.0, 0, 0, 0, 3.5, 4.0, -0.75 };
    const title_hashes = [_]u64{ 0x1111, tomb, 0x3333 };
    const two_vecs = [_]f32{ 9, 8, 7 };
    const two_hashes = [_]u64{0xABCD};
    var covs = [_]CEmbCoverage{
        s3cCoverage("title", 0, "A2:A4", "A", &title_vecs, &title_hashes),
        s3cCoverage("second", 1, "A1:A1", "A", &two_vecs, &two_hashes),
    };
    covs[1].include_formulas = 1;
    var diag = freshDiag();
    try std.testing.expectEqual(ZLSX_OK, s3cSet(ed, "test-model/v1", 3, "f32", &covs, 0, &diag, &err_buf));
    try std.testing.expectEqual(plane_none, diag.plane);
    try std.testing.expectEqual(@as(usize, 0), diagName(&diag).len);
    zlsx_diag_release(&diag);
    try std.testing.expectEqual(@as(i32, 0), zlsx_editor_save(ed, out_path.ptr, out_path.len, &err_buf, err_buf.len));

    // The read side, through the same boundary a binding uses.
    const emb = zlsx_emb_open(out_path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
    defer zlsx_emb_close(emb);
    try std.testing.expectEqual(ZLSX_EMB_PRESENT, zlsx_emb_state(emb));
    var sbuf: [128]u8 = undefined;
    try std.testing.expectEqualStrings("test-model/v1", s3cEmbString(emb, zlsx_emb_model, .{}, &sbuf));
    try std.testing.expectEqual(@as(u32, 3), zlsx_emb_dim(emb));
    try std.testing.expectEqualStrings("f32", s3cEmbString(emb, zlsx_emb_dtype, .{}, &sbuf));
    try std.testing.expectEqual(@as(usize, 2), zlsx_emb_coverage_count(emb));
    try std.testing.expectEqualStrings("title", s3cEmbString(emb, zlsx_emb_coverage_id, .{@as(usize, 0)}, &sbuf));
    try std.testing.expectEqualStrings("worksheets/sheet1.xml", s3cEmbString(emb, zlsx_emb_coverage_sheet, .{@as(usize, 0)}, &sbuf));
    try std.testing.expectEqualStrings("A2:A4", s3cEmbString(emb, zlsx_emb_coverage_range, .{@as(usize, 0)}, &sbuf));
    try std.testing.expectEqual(@as(u32, 3), zlsx_emb_coverage_rows(emb, 0));
    try std.testing.expectEqualStrings("second", s3cEmbString(emb, zlsx_emb_coverage_id, .{@as(usize, 1)}, &sbuf));
    try std.testing.expectEqualStrings("worksheets/sheet2.xml", s3cEmbString(emb, zlsx_emb_coverage_sheet, .{@as(usize, 1)}, &sbuf));
    try std.testing.expectEqualStrings("A1:A1", s3cEmbString(emb, zlsx_emb_coverage_range, .{@as(usize, 1)}, &sbuf));
    try std.testing.expectEqual(@as(u32, 1), zlsx_emb_coverage_rows(emb, 1));
    var got_vecs: [9]f32 = undefined;
    try std.testing.expectEqual(@as(i32, 0), zlsx_emb_vectors(emb, 0, &got_vecs, got_vecs.len));
    try std.testing.expectEqualSlices(f32, &title_vecs, &got_vecs);
    var got_hashes: [3]u64 = undefined;
    try std.testing.expectEqual(@as(i32, 0), zlsx_emb_hashes(emb, 0, &got_hashes, got_hashes.len));
    try std.testing.expectEqualSlices(u64, &title_hashes, &got_hashes);
    try std.testing.expectEqual(@as(i32, 0), zlsx_emb_vectors(emb, 1, got_vecs[0..3], 3));
    try std.testing.expectEqualSlices(f32, &two_vecs, got_vecs[0..3]);
    try std.testing.expectEqual(@as(i32, 0), zlsx_emb_hashes(emb, 1, got_hashes[0..1], 1));
    try std.testing.expectEqual(@as(u64, 0xABCD), got_hashes[0]);
    try std.testing.expectEqual(@as(u64, 0), zlsx_emb_digest(emb));

    // The Zig surface agrees on the fields the C read does not spell:
    // the column and include_formulas per coverage, both carriers of
    // the recovery record, ONE workbook→index relationship.
    var wb = try zlsx_pkg.Workbook.open(alloc, io, out_path);
    defer wb.deinit();
    const view = (try wb.embeddings()).present;
    try std.testing.expectEqualStrings("A", view.coverages[0].coverage.column);
    try std.testing.expect(!view.coverages[0].coverage.include_formulas);
    try std.testing.expect(view.coverages[1].coverage.include_formulas);
    const rels = (try wb.store.part("xl/_rels/workbook.xml.rels")) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqual(@as(usize, 1), std.mem.count(u8, rels.bytes, zlsx_pkg.embedding_part.REL_TYPE_EMBEDDINGS));
    try std.testing.expect((try wb.store.part("docProps/custom.xml")) != null);
    const wbxml = (try wb.store.part("xl/workbook.xml")) orelse return error.TestUnexpectedResult;
    try std.testing.expect(std.mem.indexOf(u8, wbxml.bytes, zlsx_pkg.recovery_record.NAME_PREFIX) != null);
    // The cells are untouched: no hidden recovery sheet (flags = 0).
    try std.testing.expectEqual(@as(u32, 2), wb.sheetCount());
}

test "S3c set_embeddings: int8-sym quantizes in the library and reads back within one step of the per-row scale" {
    const alloc = std.testing.allocator;
    var threaded: std.Io.Threaded = .init(alloc, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try writeS3cFixture(io, &tt, "s3c_i8_src.xlsx");
    defer alloc.free(path);
    const out_path = try tt.path(alloc, io, "s3c_i8_out.xlsx");
    defer alloc.free(out_path);

    var err_buf: [128]u8 = undefined;
    const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
    defer zlsx_editor_close(ed);
    const vecs = [_]f32{ 1.0, -2.5, 0.25, 4.0, 0, 0, 0, 0, 100, -100, 50, 0.5 };
    const hashes = [_]u64{ 1, 2, 3 };
    const covs = [_]CEmbCoverage{s3cCoverage("q", 0, "A2:A4", "A", &vecs, &hashes)};
    var diag = freshDiag();
    try std.testing.expectEqual(ZLSX_OK, s3cSet(ed, "m", 4, "int8-sym-per-vec", &covs, 0, &diag, &err_buf));
    try std.testing.expectEqual(@as(i32, 0), zlsx_editor_save(ed, out_path.ptr, out_path.len, &err_buf, err_buf.len));

    const emb = zlsx_emb_open(out_path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
    defer zlsx_emb_close(emb);
    try std.testing.expectEqual(ZLSX_EMB_PRESENT, zlsx_emb_state(emb));
    var sbuf: [64]u8 = undefined;
    try std.testing.expectEqualStrings("int8-sym-per-vec", s3cEmbString(emb, zlsx_emb_dtype, .{}, &sbuf));
    var got: [12]f32 = undefined;
    try std.testing.expectEqual(@as(i32, 0), zlsx_emb_vectors(emb, 0, &got, got.len));
    for (vecs, got, 0..) |want, have, i| {
        const row_max: f32 = switch (i / 4) {
            0 => 4.0,
            1 => 0.0,
            else => 100.0,
        };
        try std.testing.expect(@abs(want - have) <= row_max / 127.0 + 1e-6);
    }
    // The part is the compact layout, not f32: 24-byte header + 3 × (4 + 4).
    const vec_part = try savedPartBytes(alloc, io, out_path, "xl/zlsxEmbeddings/q/vec.bin");
    defer alloc.free(vec_part);
    try std.testing.expectEqual(@as(usize, 24 + 3 * 8), vec_part.len);
}

test "S3c set_embeddings: a second write replaces the set — in one editor and across a save" {
    const alloc = std.testing.allocator;
    var threaded: std.Io.Threaded = .init(alloc, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try writeS3cFixture(io, &tt, "s3c_re_src.xlsx");
    defer alloc.free(path);
    const mid_path = try tt.path(alloc, io, "s3c_re_mid.xlsx");
    defer alloc.free(mid_path);
    const out_path = try tt.path(alloc, io, "s3c_re_out.xlsx");
    defer alloc.free(out_path);

    var err_buf: [128]u8 = undefined;
    const v3 = [_]f32{ 1, 2, 3, 4, 5, 6, 7, 8, 9 };
    const h3 = [_]u64{ 1, 2, 3 };
    const v1 = [_]f32{ 1, 2, 3 };
    const h1 = [_]u64{7};
    {
        const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
        defer zlsx_editor_close(ed);
        const first = [_]CEmbCoverage{
            s3cCoverage("title", 0, "A2:A4", "A", &v3, &h3),
            s3cCoverage("body", 0, "B2:B4", "B", &v3, &h3),
        };
        try std.testing.expectEqual(ZLSX_OK, s3cSet(ed, "m1", 3, "f32", &first, 0, null, &err_buf));
        // Replaced before the save: the set is what the LAST call said.
        const second = [_]CEmbCoverage{s3cCoverage("body", 0, "B2:B2", "B", &v1, &h1)};
        try std.testing.expectEqual(ZLSX_OK, s3cSet(ed, "m2", 3, "f32", &second, 0, null, &err_buf));
        try std.testing.expectEqual(@as(i32, 0), zlsx_editor_save(ed, mid_path.ptr, mid_path.len, &err_buf, err_buf.len));
    }
    {
        const emb = zlsx_emb_open(mid_path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
        defer zlsx_emb_close(emb);
        try std.testing.expectEqual(ZLSX_EMB_PRESENT, zlsx_emb_state(emb));
        var sbuf: [64]u8 = undefined;
        try std.testing.expectEqualStrings("m2", s3cEmbString(emb, zlsx_emb_model, .{}, &sbuf));
        try std.testing.expectEqual(@as(usize, 1), zlsx_emb_coverage_count(emb));
        try std.testing.expectEqualStrings("body", s3cEmbString(emb, zlsx_emb_coverage_id, .{@as(usize, 0)}, &sbuf));
        try std.testing.expectEqual(@as(u32, 1), zlsx_emb_coverage_rows(emb, 0));
    }
    // A fresh editor on the saved file re-embeds: the relationship and
    // the recovery record are replaced, never duplicated.
    {
        const ed = zlsx_editor_open(mid_path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
        defer zlsx_editor_close(ed);
        const third = [_]CEmbCoverage{s3cCoverage("title", 1, "A1:A1", "A", &v1, &h1)};
        try std.testing.expectEqual(ZLSX_OK, s3cSet(ed, "m3", 3, "f32", &third, 0, null, &err_buf));
        try std.testing.expectEqual(@as(i32, 0), zlsx_editor_save(ed, out_path.ptr, out_path.len, &err_buf, err_buf.len));
    }
    const emb = zlsx_emb_open(out_path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
    defer zlsx_emb_close(emb);
    try std.testing.expectEqual(ZLSX_EMB_PRESENT, zlsx_emb_state(emb));
    var sbuf: [64]u8 = undefined;
    try std.testing.expectEqualStrings("m3", s3cEmbString(emb, zlsx_emb_model, .{}, &sbuf));
    try std.testing.expectEqual(@as(usize, 1), zlsx_emb_coverage_count(emb));
    try std.testing.expectEqualStrings("worksheets/sheet2.xml", s3cEmbString(emb, zlsx_emb_coverage_sheet, .{@as(usize, 0)}, &sbuf));
    var wb = try zlsx_pkg.Workbook.open(alloc, io, out_path);
    defer wb.deinit();
    const rels = (try wb.store.part("xl/_rels/workbook.xml.rels")) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqual(@as(usize, 1), std.mem.count(u8, rels.bytes, zlsx_pkg.embedding_part.REL_TYPE_EMBEDDINGS));
    const custom = (try wb.store.part("docProps/custom.xml")) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqual(@as(usize, 1), std.mem.count(u8, custom.bytes, zlsx_pkg.recovery_record.DOC_PROP_NAME));
    // The primary carrier too: ONE `_zlsxRecovery0`, the last
    // generation's (a re-embed across a save used to splice the saved
    // chunk back beside the new one — in-house r1 S3C-REL-101 /
    // S3C-TEST-107).
    const wbxml = (try wb.store.part("xl/workbook.xml")) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqual(@as(usize, 1), std.mem.count(u8, wbxml.bytes, "name=\"_zlsxRecovery0\""));
    try std.testing.expectEqual(@as(usize, 1), std.mem.count(u8, wbxml.bytes, zlsx_pkg.recovery_record.NAME_PREFIX));
    try std.testing.expect(std.mem.indexOf(u8, wbxml.bytes, "|m3|") != null);
    try std.testing.expect(std.mem.indexOf(u8, wbxml.bytes, "|m2|") == null);
    const view = (try wb.embeddings()).present;
    try std.testing.expectEqual(@as(usize, 1), view.coverages.len);
}

test "S3c set_embeddings refusals that used to fire after the parts: the package rels file, the record ceiling and the part cap, each before the first write" {
    const alloc = std.testing.allocator;
    var threaded: std.Io.Threaded = .init(alloc, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    var err_buf: [128]u8 = undefined;
    const v3 = [_]f32{ 1, 2, 3, 4, 5, 6, 7, 8, 9 };
    const h3 = [_]u64{ 1, 2, 3 };
    const covs = [_]CEmbCoverage{s3cCoverage("title", 0, "A2:A4", "A", &v3, &h3)};

    // `_rels/.rels` without its close tag: the docProps carrier's
    // package relationship has nowhere to land.
    {
        const path = try writeS3cFixture(io, &tt, "s3c_pkgrels.xlsx");
        defer alloc.free(path);
        try zlsx_pkg.pivots.fixture.patchPart(alloc, io, path, "_rels/.rels", "</Relationships>", "");
        const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
        defer zlsx_editor_close(ed);
        var diag = freshDiag();
        try std.testing.expectEqual(ZLSX_REFUSED, s3cSet(ed, "m", 3, "f32", &covs, 0, &diag, &err_buf));
        try std.testing.expectEqualStrings("MalformedWorkbookRels", diagName(&diag));
        try std.testing.expectEqual(plane_none, diag.plane);
        var out_ptr: ?[*]u8 = null;
        var out_len: usize = 0;
        try std.testing.expectEqual(ZLSX_OK, zlsx_editor_save_to_buffer(ed, &out_ptr, &out_len, &err_buf, err_buf.len));
        defer zlsx_buffer_release(out_ptr, out_len);
        const src_bytes = try readFileBytes(io, path);
        defer alloc.free(src_bytes);
        try std.testing.expectEqualSlices(u8, src_bytes, out_ptr.?[0..out_len]);
    }
    // Beside a saved record, a `<definedName` with an unterminated quote
    // outside the `<definedNames>` block — a part the open admits, since
    // the view bounds the element inside the block only — is the strip's
    // to refuse. Its verdict used to cross from the install as the
    // scanner's own `MalformedXml`, a -1 AFTER every part, the index and
    // the relationship (in-house r5 S3C-REL-501); now the workbook's
    // refusal, before the first write.
    {
        const src_path = try writeS3cFixture(io, &tt, "s3c_strip_src.xlsx");
        defer alloc.free(src_path);
        const path = try tt.path(alloc, io, "s3c_strip.xlsx");
        defer alloc.free(path);
        {
            const ed = zlsx_editor_open(src_path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
            defer zlsx_editor_close(ed);
            try std.testing.expectEqual(ZLSX_OK, s3cSet(ed, "m1", 3, "f32", &covs, 0, null, &err_buf));
            try std.testing.expectEqual(@as(i32, 0), zlsx_editor_save(ed, path.ptr, path.len, &err_buf, err_buf.len));
        }
        try zlsx_pkg.pivots.fixture.patchPart(alloc, io, path, "xl/workbook.xml", "</workbook>", "<definedName name=\"x>oops</definedName></workbook>");
        const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
        defer zlsx_editor_close(ed);
        var diag = freshDiag();
        try std.testing.expectEqual(ZLSX_REFUSED, s3cSet(ed, "m9", 3, "f32", &covs, 0, &diag, &err_buf));
        try std.testing.expectEqualStrings("MalformedWorkbookXml", diagName(&diag));
        try std.testing.expectEqual(plane_none, diag.plane);
        var out_ptr: ?[*]u8 = null;
        var out_len: usize = 0;
        try std.testing.expectEqual(ZLSX_OK, zlsx_editor_save_to_buffer(ed, &out_ptr, &out_len, &err_buf, err_buf.len));
        defer zlsx_buffer_release(out_ptr, out_len);
        const src_bytes = try readFileBytes(io, path);
        defer alloc.free(src_bytes);
        try std.testing.expectEqualSlices(u8, src_bytes, out_ptr.?[0..out_len]);
    }
    // A rels file already at the last rId: the relationship has no id
    // to take (used to trap or wrap in pass 5, after every part).
    {
        const path = try writeS3cFixture(io, &tt, "s3c_rid.xlsx");
        defer alloc.free(path);
        try zlsx_pkg.pivots.fixture.patchPart(alloc, io, path, "xl/_rels/workbook.xml.rels", "</Relationships>", "<Relationship Id=\"rId4294967295\" Type=\"http://example.com/x\" Target=\"x.bin\"/></Relationships>");
        const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
        defer zlsx_editor_close(ed);
        var diag = freshDiag();
        try std.testing.expectEqual(ZLSX_REFUSED, s3cSet(ed, "m", 3, "f32", &covs, 0, &diag, &err_buf));
        try std.testing.expectEqualStrings("IdSpaceExhausted", diagName(&diag));
        try std.testing.expectEqual(plane_none, diag.plane);
        var out_ptr: ?[*]u8 = null;
        var out_len: usize = 0;
        try std.testing.expectEqual(ZLSX_OK, zlsx_editor_save_to_buffer(ed, &out_ptr, &out_len, &err_buf, err_buf.len));
        defer zlsx_buffer_release(out_ptr, out_len);
        const src_bytes = try readFileBytes(io, path);
        defer alloc.free(src_bytes);
        try std.testing.expectEqualSlices(u8, src_bytes, out_ptr.?[0..out_len]);
    }
    // A ~3 KB model: the recovery record past its 16-chunk ceiling.
    {
        const path = try writeS3cFixture(io, &tt, "s3c_record.xlsx");
        defer alloc.free(path);
        const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
        defer zlsx_editor_close(ed);
        const long_model = [_]u8{'m'} ** 3300;
        var diag = freshDiag();
        try std.testing.expectEqual(ZLSX_REFUSED, s3cSet(ed, &long_model, 3, "f32", &covs, 0, &diag, &err_buf));
        try std.testing.expectEqualStrings("EmbeddingExceedsArchiveLimit", diagName(&diag));
        try std.testing.expectEqualStrings("EmbeddingExceedsArchiveLimit", std.mem.sliceTo(&err_buf, 0));
        try std.testing.expectEqual(plane_none, diag.plane);
        // A coverage past the part cap, refused from the inputs before
        // any vector byte is read: a one-row, 2^27-wide f32 record is
        // 512 MiB + the header. The buffer behind the claimed length is
        // never dereferenced.
        var wide = s3cCoverage("wide", 0, "A2:A2", "A", &v3, h3[0..1]);
        wide.vectors_len = 1 << 27;
        try std.testing.expectEqual(ZLSX_REFUSED, s3cSet(ed, "m", 1 << 27, "f32", &.{wide}, 0, &diag, &err_buf));
        try std.testing.expectEqualStrings("EmbeddingExceedsArchiveLimit", diagName(&diag));
        // Every row of the grid times a 2^20-wide record: checked
        // arithmetic, refused, never a trap.
        var huge = s3cCoverage("huge", 0, "A1:A1048576", "A", &v3, &h3);
        huge.vectors_len = 1 << 40;
        huge.hashes_len = 1 << 20;
        try std.testing.expectEqual(ZLSX_REFUSED, s3cSet(ed, "m", 1 << 20, "f32", &.{huge}, 0, &diag, &err_buf));
        try std.testing.expectEqualStrings("EmbeddingExceedsArchiveLimit", diagName(&diag));
        var out_ptr: ?[*]u8 = null;
        var out_len: usize = 0;
        try std.testing.expectEqual(ZLSX_OK, zlsx_editor_save_to_buffer(ed, &out_ptr, &out_len, &err_buf, err_buf.len));
        defer zlsx_buffer_release(out_ptr, out_len);
        const src_bytes = try readFileBytes(io, path);
        defer alloc.free(src_bytes);
        try std.testing.expectEqualSlices(u8, src_bytes, out_ptr.?[0..out_len]);
    }
}

test "S3c set_embeddings contract violations: -1 with the name in errbuf, the diag left as prep left it, nothing written" {
    const alloc = std.testing.allocator;
    var threaded: std.Io.Threaded = .init(alloc, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try writeS3cFixture(io, &tt, "s3c_bad.xlsx");
    defer alloc.free(path);
    var err_buf: [128]u8 = undefined;
    const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
    defer zlsx_editor_close(ed);

    const v3 = [_]f32{ 1, 2, 3, 4, 5, 6, 7, 8, 9 };
    const h3 = [_]u64{ 1, 2, 3 };
    const ok_cov = s3cCoverage("title", 0, "A2:A4", "A", &v3, &h3);
    const Case = struct { name: []const u8, model: []const u8 = "m", dim: u32 = 3, dtype: []const u8 = "f32", covs: []const CEmbCoverage, flags: u32 = 0, ed: bool = true };
    var bad_len = ok_cov;
    bad_len.vectors_len = 8;
    var bad_hlen = ok_cov;
    bad_hlen.hashes_len = 2;
    var null_vec = ok_cov;
    null_vec.vectors = null;
    var null_hash = ok_cov;
    null_hash.hashes = null;
    var null_id = ok_cov;
    null_id.id = null;
    var null_range = ok_cov;
    null_range.range = null;
    var null_col = ok_cov;
    null_col.column = null;
    var formulas2 = ok_cov;
    formulas2.include_formulas = 2;
    const cases = [_]Case{
        .{ .name = "InvalidInput", .covs = &.{ok_cov}, .ed = false },
        .{ .name = "InvalidInput", .covs = &.{ok_cov}, .flags = 1 },
        .{ .name = "InvalidInput", .covs = &.{ok_cov}, .flags = 0x8000_0000 },
        .{ .name = "InvalidInput", .covs = &.{formulas2} },
        .{ .name = "InvalidInput", .covs = &.{null_vec} },
        .{ .name = "InvalidInput", .covs = &.{null_hash} },
        .{ .name = "InvalidInput", .covs = &.{null_id} },
        .{ .name = "InvalidInput", .covs = &.{null_range} },
        .{ .name = "InvalidInput", .covs = &.{null_col} },
        .{ .name = "InvalidEmbeddingInput", .covs = &.{} },
        .{ .name = "InvalidEmbeddingInput", .covs = &.{ok_cov}, .dim = 0 },
        .{ .name = "InvalidEmbeddingInput", .covs = &.{ok_cov}, .dim = 4 },
        .{ .name = "InvalidEmbeddingInput", .covs = &.{bad_len} },
        .{ .name = "InvalidEmbeddingInput", .covs = &.{bad_hlen} },
        .{ .name = "InvalidDtype", .covs = &.{ok_cov}, .dtype = "float32" },
        .{ .name = "InvalidDtype", .covs = &.{ok_cov}, .dtype = "" },
        .{ .name = "UnsupportedDtype", .covs = &.{ok_cov}, .dtype = "binary16" },
        .{ .name = "UnsupportedDtype", .covs = &.{ok_cov}, .dtype = "bfloat16" },
        .{ .name = "UnsupportedDtype", .covs = &.{ok_cov}, .dtype = "int8-asym-per-vec" },
        .{ .name = "SheetIndexOutOfRange", .covs = &.{s3cCoverage("title", 2, "A2:A4", "A", &v3, &h3)} },
        .{ .name = "SheetIndexOutOfRange", .covs = &.{s3cCoverage("title", std.math.maxInt(u32), "A2:A4", "A", &v3, &h3)} },
        .{ .name = "InvalidRange", .covs = &.{s3cCoverage("title", 0, "A0:A2", "A", &v3, &h3)} },
        .{ .name = "InvalidRange", .covs = &.{s3cCoverage("title", 0, "A2:A4", "Z", &v3, &h3)} },
        .{ .name = "InvalidRange", .covs = &.{s3cCoverage("title", 0, "A2:A4", "", &v3, &h3)} },
        .{ .name = "InvalidCoverageId", .covs = &.{s3cCoverage("bad id", 0, "A2:A4", "A", &v3, &h3)} },
        .{ .name = "InvalidCoverageId", .covs = &.{s3cCoverage("", 0, "A2:A4", "A", &v3, &h3)} },
        .{ .name = "DuplicateCoverageId", .covs = &.{ ok_cov, s3cCoverage("title", 0, "B2:B4", "B", &v3, &h3) } },
        .{ .name = "CoverageOverlap", .covs = &.{ ok_cov, s3cCoverage("other", 0, "A3:A5", "A", &v3, &h3) } },
        .{ .name = "InvalidXmlByte", .covs = &.{ok_cov}, .model = "m\x00" },
        .{ .name = "InvalidXmlByte", .covs = &.{ok_cov}, .model = "\x1fm" },
    };
    for (cases) |case| {
        var diag = freshDiag();
        @memset(&err_buf, 0xAA);
        const rc = s3cSet(if (case.ed) ed else null, case.model, case.dim, case.dtype, case.covs, case.flags, &diag, &err_buf);
        try std.testing.expectEqual(ZLSX_ERROR, rc);
        try std.testing.expectEqualStrings(case.name, std.mem.sliceTo(&err_buf, 0));
        try std.testing.expectEqual(@as(usize, 0), diagName(&diag).len);
        try std.testing.expectEqual(plane_none, diag.plane);
    }
    // NULL bytes with a non-zero length, on the two scalar strings.
    {
        var diag = freshDiag();
        try std.testing.expectEqual(ZLSX_ERROR, zlsx_editor_set_embeddings(ed, null, 1, 3, "f32", 3, &.{ok_cov}, 1, 0, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings("InvalidInput", std.mem.sliceTo(&err_buf, 0));
        try std.testing.expectEqual(ZLSX_ERROR, zlsx_editor_set_embeddings(ed, "m", 1, 3, null, 3, &.{ok_cov}, 1, 0, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings("InvalidInput", std.mem.sliceTo(&err_buf, 0));
        try std.testing.expectEqual(ZLSX_ERROR, zlsx_editor_set_embeddings(ed, "m", 1, 3, "f32", 3, null, 1, 0, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings("InvalidInput", std.mem.sliceTo(&err_buf, 0));
        // NULL with length 0 is the empty string / array: judged by
        // the write, never a boundary violation.
        try std.testing.expectEqual(ZLSX_ERROR, zlsx_editor_set_embeddings(ed, "m", 1, 3, "f32", 3, null, 0, 0, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings("InvalidEmbeddingInput", std.mem.sliceTo(&err_buf, 0));
        try std.testing.expectEqual(ZLSX_ERROR, zlsx_editor_set_embeddings(ed, "m", 1, 3, null, 0, &.{ok_cov}, 1, 0, &diag, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings("InvalidDtype", std.mem.sliceTo(&err_buf, 0));
    }
    // A diag below the v1 size: -1 StructSizeTooSmall, the diag byte-for-byte.
    {
        var small = freshDiag();
        small.struct_size = @sizeOf(CDiag) - 1;
        const before = std.mem.toBytes(small);
        try std.testing.expectEqual(ZLSX_ERROR, s3cSet(ed, "m", 3, "f32", &.{ok_cov}, 0, &small, &err_buf));
        try std.testing.expectEqualStrings("StructSizeTooSmall", std.mem.sliceTo(&err_buf, 0));
        try std.testing.expectEqualSlices(u8, &before, &std.mem.toBytes(small));
    }
    // Nothing was written: the editor still saves the source bytes.
    var out_ptr: ?[*]u8 = null;
    var out_len: usize = 0;
    try std.testing.expectEqual(ZLSX_OK, zlsx_editor_save_to_buffer(ed, &out_ptr, &out_len, &err_buf, err_buf.len));
    defer zlsx_buffer_release(out_ptr, out_len);
    const src_bytes = try readFileBytes(io, path);
    defer alloc.free(src_bytes);
    try std.testing.expectEqualSlices(u8, src_bytes, out_ptr.?[0..out_len]);
    // And the empty model is the write's to judge: it lands.
    try std.testing.expectEqual(ZLSX_OK, zlsx_editor_set_embeddings(ed, null, 0, 3, "f32", 3, &.{ok_cov}, 1, 0, null, &err_buf, err_buf.len));
}

test "S3c set_embeddings refusals: a rels file the relationship cannot land in is -2 with the name in the diag, nothing written" {
    const alloc = std.testing.allocator;
    var threaded: std.Io.Threaded = .init(alloc, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try writeS3cFixture(io, &tt, "s3c_rels.xlsx");
    defer alloc.free(path);
    try zlsx_pkg.pivots.fixture.patchPart(alloc, io, path, "xl/_rels/workbook.xml.rels", "</Relationships>", "");
    var err_buf: [128]u8 = undefined;
    const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
    defer zlsx_editor_close(ed);
    const v3 = [_]f32{ 1, 2, 3, 4, 5, 6, 7, 8, 9 };
    const h3 = [_]u64{ 1, 2, 3 };
    const covs = [_]CEmbCoverage{s3cCoverage("title", 0, "A2:A4", "A", &v3, &h3)};
    // Poisoned on entry: the refusal path itself must reset the pair.
    var diag = freshDiag();
    diag.plane = 3;
    @memcpy(diag.error_name[0..5], "stale");
    try std.testing.expectEqual(ZLSX_REFUSED, s3cSet(ed, "m", 3, "f32", &covs, 0, &diag, &err_buf));
    try std.testing.expectEqualStrings("MalformedWorkbookRels", diagName(&diag));
    try std.testing.expectEqualStrings("MalformedWorkbookRels", std.mem.sliceTo(&err_buf, 0));
    try std.testing.expectEqual(plane_none, diag.plane);
    zlsx_diag_release(&diag);
    var out_ptr: ?[*]u8 = null;
    var out_len: usize = 0;
    try std.testing.expectEqual(ZLSX_OK, zlsx_editor_save_to_buffer(ed, &out_ptr, &out_len, &err_buf, err_buf.len));
    defer zlsx_buffer_release(out_ptr, out_len);
    const src_bytes = try readFileBytes(io, path);
    defer alloc.free(src_bytes);
    try std.testing.expectEqualSlices(u8, src_bytes, out_ptr.?[0..out_len]);
    // The vocabulary is pinned on the mapping itself, refusals and
    // call errors alike.
    try std.testing.expectEqual(ZLSX_REFUSED, statusOf(error.MissingWorkbookRels));
    try std.testing.expectEqual(ZLSX_REFUSED, statusOf(error.EmbeddingExceedsArchiveLimit));
    try std.testing.expectEqual(ZLSX_REFUSED, statusOf(error.MissingRelationship));
    try std.testing.expectEqual(ZLSX_ERROR, statusOf(error.InvalidEmbeddingInput));
    try std.testing.expectEqual(ZLSX_ERROR, statusOf(error.InvalidDtype));
    try std.testing.expectEqual(ZLSX_ERROR, statusOf(error.UnsupportedDtype));
    try std.testing.expectEqual(ZLSX_ERROR, statusOf(error.InvalidRange));
    try std.testing.expectEqual(ZLSX_ERROR, statusOf(error.InvalidCoverageId));
    try std.testing.expectEqual(ZLSX_ERROR, statusOf(error.DuplicateCoverageId));
    try std.testing.expectEqual(ZLSX_ERROR, statusOf(error.CoverageOverlap));
    try std.testing.expectEqual(ZLSX_ERROR, statusOf(error.InvalidXmlByte));
    try std.testing.expectEqual(ZLSX_ERROR, statusOf(error.SheetIndexOutOfRange));
}

fn encodeVectorBodyForFailures(alloc: std.mem.Allocator) !void {
    const vectors = [_]f32{ 1, 2, 3, 4, 5, 6 };
    const body = try zlsx_pkg.embedding_part.encodeVectorBody(alloc, .int8_sym_per_vec, 3, &vectors);
    alloc.free(body);
}

test "S3c set_embeddings: the body encoder's allocation failure is OutOfMemory, the boundary's -3" {
    try std.testing.checkAllAllocationFailures(std.testing.allocator, encodeVectorBodyForFailures, .{});
    try std.testing.expectEqual(ZLSX_NOMEM, statusOf(error.OutOfMemory));
}

// ── S3c slice 2 tests ─────────────────────────────────────────────────

fn s3cRows(ed: ?*Editor, sheet_idx: u32, range: []const u8, column: []const u8, include_formulas: u32, out_ptr: ?*?[*]u8, out_len: ?*usize, diag: ?*CDiag, err_buf: []u8) i32 {
    return zlsx_editor_embeddable_rows_ndjson(ed, sheet_idx, range.ptr, range.len, column.ptr, column.len, include_formulas, out_ptr, out_len, diag, err_buf.ptr, err_buf.len);
}

/// The canonical hash of `cell` at `row` on the first sheet — what the
/// read must hand over, spelled by the library's own hasher.
fn s3cHash(row: u32, cell: zlsx_pkg.embedding_part.CanonicalCell) !u64 {
    var scratch: std.ArrayListUnmanaged(u8) = .empty;
    defer scratch.deinit(std.testing.allocator);
    return zlsx_pkg.embedding_part.xxh3Canonical(std.testing.allocator, "worksheets/sheet1.xml", row, cell, &scratch);
}

/// One `embed_row` line as the shared writer spells it.
fn s3cRecord(alloc: std.mem.Allocator, row: u32, text: []const u8, hash: u64) ![]u8 {
    return std.fmt.allocPrint(alloc, "{{\"kind\":\"embed_row\",\"row\":{d},\"text\":\"{s}\",\"hash\":{d}}}\n", .{ row, text, hash });
}

test "S3c embeddable_rows: the records cross the boundary byte-for-byte, and fed back to set_embeddings the hashes read fresh" {
    const alloc = std.testing.allocator;
    var threaded: std.Io.Threaded = .init(alloc, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try writeS3cFixture(io, &tt, "s3c2_src.xlsx");
    defer alloc.free(path);
    const out_path = try tt.path(alloc, io, "s3c2_out.xlsx");
    defer alloc.free(out_path);

    var err_buf: [128]u8 = undefined;
    const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
    defer zlsx_editor_close(ed);

    var out_ptr: ?[*]u8 = null;
    var out_len: usize = 0;
    var diag = freshDiag();
    try std.testing.expectEqual(ZLSX_OK, s3cRows(ed, 0, "A2:A4", "A", 0, &out_ptr, &out_len, &diag, &err_buf));
    defer zlsx_buffer_release(out_ptr, out_len);
    try std.testing.expectEqual(plane_none, diag.plane);
    try std.testing.expectEqual(@as(usize, 0), diagName(&diag).len);
    zlsx_diag_release(&diag);

    // The bytes are the shared writer's, over the canonical hashes.
    const h2 = try s3cHash(2, .{ .string = "alpha" });
    const h3 = try s3cHash(3, .{ .string = "beta" });
    const h4 = try s3cHash(4, .{ .string = "gamma" });
    var want: std.ArrayListUnmanaged(u8) = .empty;
    defer want.deinit(alloc);
    for ([_]struct { r: u32, t: []const u8, h: u64 }{ .{ .r = 2, .t = "alpha", .h = h2 }, .{ .r = 3, .t = "beta", .h = h3 }, .{ .r = 4, .t = "gamma", .h = h4 } }) |e| {
        const line = try s3cRecord(alloc, e.r, e.t, e.h);
        defer alloc.free(line);
        try want.appendSlice(alloc, line);
    }
    try std.testing.expectEqualStrings(want.items, out_ptr.?[0..out_len]);
    // Pinned across surfaces: the Python leg asserts the same literal
    // on the same fixture (`test_embeddable_rows.py`), under this
    // suite because CI's pytest lane is advisory.
    try std.testing.expectEqual(@as(u64, 6830279115424181645), h2);

    // Fed straight back to the write, every slot reads fresh under the
    // sweep: the read's hash IS the write's.
    const vecs = [_]f32{ 1, 2, 3, 4, 5, 6, 7, 8, 9 };
    const hashes = [_]u64{ h2, h3, h4 };
    const covs = [_]CEmbCoverage{s3cCoverage("title", 0, "A2:A4", "A", &vecs, &hashes)};
    try std.testing.expectEqual(ZLSX_OK, s3cSet(ed, "m", 3, "f32", &covs, 0, null, &err_buf));
    try std.testing.expectEqual(@as(i32, 0), zlsx_editor_save(ed, out_path.ptr, out_path.len, &err_buf, err_buf.len));
    var wb = try zlsx_pkg.Workbook.open(alloc, io, out_path);
    defer wb.deinit();
    const report = try wb.pruneEmbeddings();
    try std.testing.expectEqual(@as(usize, 3), report.fresh);
    try std.testing.expectEqual(@as(usize, 0), report.stale);
    try std.testing.expectEqual(@as(usize, 0), report.redacted);

    // The second sheet by index, the write's own resolution.
    var second_ptr: ?[*]u8 = null;
    var second_len: usize = 0;
    try std.testing.expectEqual(ZLSX_OK, s3cRows(ed, 1, "A1:A1", "A", 0, &second_ptr, &second_len, null, &err_buf));
    defer zlsx_buffer_release(second_ptr, second_len);
    var scratch: std.ArrayListUnmanaged(u8) = .empty;
    defer scratch.deinit(alloc);
    const two = try zlsx_pkg.embedding_part.xxh3Canonical(alloc, "worksheets/sheet2.xml", 1, .{ .string = "two" }, &scratch);
    const two_line = try s3cRecord(alloc, 1, "two", two);
    defer alloc.free(two_line);
    try std.testing.expectEqualStrings(two_line, second_ptr.?[0..second_len]);

    // A range with nothing embeddable: OK, the outputs reset to (NULL, 0).
    var empty_ptr: ?[*]u8 = out_ptr;
    var empty_len: usize = 99;
    try std.testing.expectEqual(ZLSX_OK, s3cRows(ed, 0, "C2:C4", "C", 0, &empty_ptr, &empty_len, null, &err_buf));
    try std.testing.expectEqual(@as(?[*]u8, null), empty_ptr);
    try std.testing.expectEqual(@as(usize, 0), empty_len);
}

/// Every cell kind under one column: an entity-bearing shared string,
/// a number, a blank, a boolean, a formula with a cached string, a
/// rich shared string (written last — the writer numbers a rich cell
/// as it is written but emits rich entries after every plain one, a
/// pre-existing defect recorded outside this slice).
fn writeS3c2KindsFixture(io: std.Io, tt: *TestTmp, name: []const u8) ![:0]u8 {
    const alloc = std.testing.allocator;
    const path = try tt.path(alloc, io, name);
    errdefer alloc.free(path);
    var w = xlsx.Writer.init(alloc);
    defer w.deinit();
    var kinds = try w.addSheet("Kinds");
    try kinds.writeRow(&.{.{ .string = "title" }});
    try kinds.writeRow(&.{.{ .string = "a & b <c>" }});
    try kinds.writeRow(&.{.{ .number = 1.5 }});
    try kinds.writeRow(&.{.empty});
    try kinds.writeRow(&.{.{ .boolean = true }});
    try kinds.writeRowWithFormulas(&.{.{ .string = "cached" }}, &.{"A2"});
    try kinds.writeRichRow(&.{.{ .rich = &.{ .{ .text = "Hello", .bold = true }, .{ .text = " world" } } }});
    try w.save(io, path);
    return path;
}

test "S3c embeddable_rows: every kind as a reader sees it — entities resolved, runs joined, a blank omitted, formulas only on request" {
    const alloc = std.testing.allocator;
    var threaded: std.Io.Threaded = .init(alloc, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try writeS3c2KindsFixture(io, &tt, "s3c2_kinds.xlsx");
    defer alloc.free(path);
    var err_buf: [128]u8 = undefined;
    const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
    defer zlsx_editor_close(ed);

    const Line = struct { r: u32, t: []const u8, h: u64 };
    const l2: Line = .{ .r = 2, .t = "a & b <c>", .h = try s3cHash(2, .{ .string = "a & b <c>" }) };
    const l3: Line = .{ .r = 3, .t = "1.5", .h = try s3cHash(3, .{ .number = "1.5" }) };
    const l5: Line = .{ .r = 5, .t = "1", .h = try s3cHash(5, .{ .boolean = true }) };
    const l6: Line = .{ .r = 6, .t = "cached", .h = try s3cHash(6, .{ .string = "cached" }) };
    const l7: Line = .{ .r = 7, .t = "Hello world", .h = try s3cHash(7, .{ .string = "Hello world" }) };
    const cases = [_]struct { inc: u32, lines: []const Line }{
        .{ .inc = 0, .lines = &.{ l2, l3, l5, l7 } },
        .{ .inc = 1, .lines = &.{ l2, l3, l5, l6, l7 } },
    };
    for (cases) |case| {
        var want: std.ArrayListUnmanaged(u8) = .empty;
        defer want.deinit(alloc);
        for (case.lines) |e| {
            const line = try s3cRecord(alloc, e.r, e.t, e.h);
            defer alloc.free(line);
            try want.appendSlice(alloc, line);
        }
        var out_ptr: ?[*]u8 = null;
        var out_len: usize = 0;
        try std.testing.expectEqual(ZLSX_OK, s3cRows(ed, 0, "A2:A7", "A", case.inc, &out_ptr, &out_len, null, &err_buf));
        defer zlsx_buffer_release(out_ptr, out_len);
        try std.testing.expectEqualStrings(want.items, out_ptr.?[0..out_len]);
    }
}

test "S3c embeddable_rows contract violations: -1 with the name in errbuf, the diag as prep left it, the outputs reset" {
    const alloc = std.testing.allocator;
    var threaded: std.Io.Threaded = .init(alloc, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try writeS3cFixture(io, &tt, "s3c2_bad.xlsx");
    defer alloc.free(path);
    var err_buf: [128]u8 = undefined;
    const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
    defer zlsx_editor_close(ed);

    const Case = struct { name: []const u8, ed: bool = true, sheet: u32 = 0, range: []const u8 = "A2:A4", column: []const u8 = "A", inc: u32 = 0 };
    const cases = [_]Case{
        .{ .name = "InvalidInput", .ed = false },
        .{ .name = "InvalidInput", .inc = 2 },
        .{ .name = "InvalidInput", .inc = std.math.maxInt(u32) },
        .{ .name = "InvalidRange", .range = "A0:A2" },
        .{ .name = "InvalidRange", .range = "" },
        .{ .name = "InvalidRange", .range = "A2:A4", .column = "Z" },
        .{ .name = "InvalidRange", .column = "" },
        .{ .name = "SheetIndexOutOfRange", .sheet = 2 },
        .{ .name = "SheetIndexOutOfRange", .sheet = std.math.maxInt(u32) },
    };
    // A poisoned output pair: every -1 resets it before judging.
    var poison: [1]u8 = .{0};
    for (cases) |case| {
        var diag = freshDiag();
        @memset(&err_buf, 0xAA);
        var out_ptr: ?[*]u8 = &poison;
        var out_len: usize = 7;
        const rc = s3cRows(if (case.ed) ed else null, case.sheet, case.range, case.column, case.inc, &out_ptr, &out_len, &diag, &err_buf);
        try std.testing.expectEqual(ZLSX_ERROR, rc);
        try std.testing.expectEqualStrings(case.name, std.mem.sliceTo(&err_buf, 0));
        try std.testing.expectEqual(@as(usize, 0), diagName(&diag).len);
        try std.testing.expectEqual(plane_none, diag.plane);
        try std.testing.expectEqual(@as(?[*]u8, null), out_ptr);
        try std.testing.expectEqual(@as(usize, 0), out_len);
    }
    // NULL bytes with a non-zero length; NULL with length 0 is the
    // empty string, the read's to judge (InvalidRange).
    {
        var out_ptr: ?[*]u8 = null;
        var out_len: usize = 0;
        try std.testing.expectEqual(ZLSX_ERROR, zlsx_editor_embeddable_rows_ndjson(ed, 0, null, 5, "A", 1, 0, &out_ptr, &out_len, null, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings("InvalidInput", std.mem.sliceTo(&err_buf, 0));
        try std.testing.expectEqual(ZLSX_ERROR, zlsx_editor_embeddable_rows_ndjson(ed, 0, "A2:A4", 5, null, 1, 0, &out_ptr, &out_len, null, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings("InvalidInput", std.mem.sliceTo(&err_buf, 0));
        try std.testing.expectEqual(ZLSX_ERROR, zlsx_editor_embeddable_rows_ndjson(ed, 0, null, 0, "A", 1, 0, &out_ptr, &out_len, null, &err_buf, err_buf.len));
        try std.testing.expectEqualStrings("InvalidRange", std.mem.sliceTo(&err_buf, 0));
        // NULL out pointers.
        try std.testing.expectEqual(ZLSX_ERROR, s3cRows(ed, 0, "A2:A4", "A", 0, null, &out_len, null, &err_buf));
        try std.testing.expectEqualStrings("NullOutPointer", std.mem.sliceTo(&err_buf, 0));
        try std.testing.expectEqual(ZLSX_ERROR, s3cRows(ed, 0, "A2:A4", "A", 0, &out_ptr, null, null, &err_buf));
        try std.testing.expectEqualStrings("NullOutPointer", std.mem.sliceTo(&err_buf, 0));
        try std.testing.expectEqual(@as(?[*]u8, null), out_ptr);
    }
    // A diag below the v1 size: -1 StructSizeTooSmall, the diag
    // byte-for-byte — and the outputs reset BEFORE the size is judged
    // (poisoned here; round 1 TEST-104).
    {
        var small = freshDiag();
        small.struct_size = @sizeOf(CDiag) - 1;
        const before = std.mem.toBytes(small);
        var out_ptr: ?[*]u8 = &poison;
        var out_len: usize = 7;
        try std.testing.expectEqual(ZLSX_ERROR, s3cRows(ed, 0, "A2:A4", "A", 0, &out_ptr, &out_len, &small, &err_buf));
        try std.testing.expectEqualStrings("StructSizeTooSmall", std.mem.sliceTo(&err_buf, 0));
        try std.testing.expectEqualSlices(u8, &before, &std.mem.toBytes(small));
        try std.testing.expectEqual(@as(?[*]u8, null), out_ptr);
        try std.testing.expectEqual(@as(usize, 0), out_len);
    }
    // A staged cell write — outside the range, even — makes the sheet
    // unreadable for this read: a sequencing statement, -1, never a
    // refusal. The other sheet still answers.
    {
        const cell = toCCell(.{ .integer = 42 });
        try std.testing.expectEqual(@as(i32, 0), zlsx_editor_set_cell(ed, 0, 9, 3, &cell, &err_buf, err_buf.len));
        var out_ptr: ?[*]u8 = null;
        var out_len: usize = 0;
        var diag = freshDiag();
        try std.testing.expectEqual(ZLSX_ERROR, s3cRows(ed, 0, "A2:A4", "A", 0, &out_ptr, &out_len, &diag, &err_buf));
        try std.testing.expectEqualStrings("SheetHasUnsavedMutations", std.mem.sliceTo(&err_buf, 0));
        try std.testing.expectEqual(@as(usize, 0), diagName(&diag).len);
        try std.testing.expectEqual(@as(?[*]u8, null), out_ptr);
        try std.testing.expectEqual(ZLSX_OK, s3cRows(ed, 1, "A1:A1", "A", 0, &out_ptr, &out_len, &diag, &err_buf));
        zlsx_buffer_release(out_ptr, out_len);
    }
}

test "S3c embeddable_rows refusals: -2 with the name in the diag — a sheet part the view cannot parse, a value the read cannot carry" {
    const alloc = std.testing.allocator;
    var threaded: std.Io.Threaded = .init(alloc, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    var err_buf: [128]u8 = undefined;

    const Case = struct { file: []const u8, part: []const u8, old: []const u8, new: []const u8, name: []const u8, other_served: bool = true };
    const cases = [_]Case{
        .{ .file = "s3c2_sheet.xlsx", .part = "xl/worksheets/sheet1.xml", .old = "</sheetData>", .new = "", .name = "MalformedSheetXml" },
        .{ .file = "s3c2_sst.xlsx", .part = "xl/worksheets/sheet1.xml", .old = "t=\"s\"><v>2</v>", .new = "t=\"s\"><v>99</v>", .name = "SstIndexOutOfRange" },
        .{ .file = "s3c2_bool.xlsx", .part = "xl/worksheets/sheet1.xml", .old = "t=\"s\"><v>2</v>", .new = "t=\"b\"><v>TRUE</v>", .name = "UnsupportedCellValue" },
        .{ .file = "s3c2_entity.xlsx", .part = "xl/sharedStrings.xml", .old = ">alpha</t>", .new = ">&bogus;</t>", .name = "UnsupportedCellValue" },
        .{ .file = "s3c2_utf8.xlsx", .part = "xl/sharedStrings.xml", .old = ">alpha</t>", .new = ">\xff</t>", .name = "InvalidUtf8" },
        // Round 1: a `<v>` the number canonicalizer cannot read and a
        // `t="d"` date (REL-101), a `t` the reader does not know
        // (REL-103), the UTF-8 check on the kinds the hash does not
        // validate (TEST-104), a table the parser cannot read (REL-102).
        .{ .file = "s3c2_comma.xlsx", .part = "xl/worksheets/sheet1.xml", .old = "t=\"s\"><v>2</v>", .new = "><v>1,5</v>", .name = "UnsupportedCellValue" },
        .{ .file = "s3c2_date.xlsx", .part = "xl/worksheets/sheet1.xml", .old = "t=\"s\"><v>2</v>", .new = "t=\"d\"><v>2024</v>", .name = "UnsupportedCellValue" },
        .{ .file = "s3c2_unknown_t.xlsx", .part = "xl/worksheets/sheet1.xml", .old = "t=\"s\"><v>2</v>", .new = "t=\"zz\"><v>42</v>", .name = "UnsupportedCellValue" },
        .{ .file = "s3c2_err_utf8.xlsx", .part = "xl/worksheets/sheet1.xml", .old = "t=\"s\"><v>2</v>", .new = "t=\"e\"><v>#N/\xff</v>", .name = "InvalidUtf8" },
        .{ .file = "s3c2_num_utf8.xlsx", .part = "xl/worksheets/sheet1.xml", .old = "t=\"s\"><v>2</v>", .new = "><v>1\xff</v>", .name = "InvalidUtf8" },
        // The LAST entry's close (the parser tolerates an inner one); the
        // table is the workbook's, so the other sheet refuses too.
        .{ .file = "s3c2_sst_part.xlsx", .part = "xl/sharedStrings.xml", .old = ">two</t></si>", .new = ">two</t>", .name = "MalformedSharedStringsXml", .other_served = false },
        // Round 2: a row or cell without `r` (positional OOXML the typed
        // view cannot place) refuses rather than read as blank (REL-201 B).
        .{ .file = "s3c2_row_no_r.xlsx", .part = "xl/worksheets/sheet1.xml", .old = "<row r=\"3\"", .new = "<row", .name = "MalformedSheetXml" },
        .{ .file = "s3c2_cell_no_r.xlsx", .part = "xl/worksheets/sheet1.xml", .old = "<c r=\"A3\"", .new = "<c", .name = "MalformedSheetXml" },
    };
    for (cases) |case| {
        const path = try writeS3cFixture(io, &tt, case.file);
        defer alloc.free(path);
        try zlsx_pkg.pivots.fixture.patchPart(alloc, io, path, case.part, case.old, case.new);
        const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
        defer zlsx_editor_close(ed);
        var poison: [1]u8 = .{0};
        var out_ptr: ?[*]u8 = &poison;
        var out_len: usize = 7;
        var diag = freshDiag();
        diag.plane = 3;
        @memset(&err_buf, 0xAA);
        try std.testing.expectEqual(ZLSX_REFUSED, s3cRows(ed, 0, "A2:A4", "A", 0, &out_ptr, &out_len, &diag, &err_buf));
        try std.testing.expectEqualStrings(case.name, diagName(&diag));
        try std.testing.expectEqualStrings(case.name, std.mem.sliceTo(&err_buf, 0));
        try std.testing.expectEqual(plane_none, diag.plane);
        try std.testing.expectEqual(@as(?[*]u8, null), out_ptr);
        try std.testing.expectEqual(@as(usize, 0), out_len);
        zlsx_diag_release(&diag);
        // The verdict is the part's: a sheet part's leaves the other
        // sheet served, the shared-string table's refuses it too.
        const other = s3cRows(ed, 1, "A1:A1", "A", 0, &out_ptr, &out_len, null, &err_buf);
        try std.testing.expectEqual(if (case.other_served) ZLSX_OK else ZLSX_REFUSED, other);
        zlsx_buffer_release(out_ptr, out_len);
    }

    // r2 REL-201: a cell placed under another row OUTSIDE the range
    // refuses too — the rule is the sheet's, not the range's.
    {
        const path = try writeS3cFixture(io, &tt, "s3c2_misplaced.xlsx");
        defer alloc.free(path);
        try zlsx_pkg.pivots.fixture.patchPart(alloc, io, path, "xl/worksheets/sheet1.xml", "<c r=\"A4\"", "<c r=\"A2\"");
        const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
        defer zlsx_editor_close(ed);
        var out_ptr: ?[*]u8 = null;
        var out_len: usize = 0;
        var diag = freshDiag();
        try std.testing.expectEqual(ZLSX_REFUSED, s3cRows(ed, 0, "A2:A3", "A", 0, &out_ptr, &out_len, &diag, &err_buf));
        try std.testing.expectEqualStrings("MalformedSheetXml", diagName(&diag));
        zlsx_diag_release(&diag);
    }
}

test "S3c embeddable_rows: the plane of every verdict the read can raise" {
    for ([_]anyerror{ error.UnsupportedCellValue, error.SstIndexOutOfRange, error.InvalidUtf8, error.UnicodeNormalizationFailed, error.MalformedSharedStringsXml, error.MalformedSheetXml, error.MissingSheetPart, error.MissingRelationship }) |e| {
        try std.testing.expectEqual(ZLSX_REFUSED, statusOf(e));
    }
    // The canonicalizer's own name never reaches the boundary from this
    // read (folded under UnsupportedCellValue); should it, it is a -1.
    try std.testing.expectEqual(ZLSX_ERROR, statusOf(error.MalformedNumber));
    for ([_]anyerror{ error.SheetHasUnsavedMutations, error.SheetHasUnsavedAppends, error.InvalidRange, error.SheetIndexOutOfRange }) |e| {
        try std.testing.expectEqual(ZLSX_ERROR, statusOf(e));
    }
}

// ── S3c slice 3 tests ─────────────────────────────────────────────────

fn s3cPrune(ed: ?*Editor, report: ?*CPruneReport, diag: ?*CDiag, err_buf: []u8) i32 {
    return zlsx_editor_prune_embeddings(ed, report, diag, err_buf.ptr, err_buf.len);
}

fn s3cStrip(ed: ?*Editor, diag: ?*CDiag, err_buf: []u8) i32 {
    return zlsx_editor_strip_embeddings(ed, diag, err_buf.ptr, err_buf.len);
}

/// A report poisoned in every count, so a zero after the call is the
/// library's, never the caller's.
fn poisonedPruneReport() CPruneReport {
    return .{ .struct_size = @sizeOf(CPruneReport), .redacted = 0xAA, .stale = 0xAA, .fresh = 0xAA, .valid_empty = 0xAA };
}

fn expectPruneCounts(report: CPruneReport, redacted: u64, stale: u64, fresh: u64, valid_empty: u64) !void {
    try std.testing.expectEqual(redacted, report.redacted);
    try std.testing.expectEqual(stale, report.stale);
    try std.testing.expectEqual(fresh, report.fresh);
    try std.testing.expectEqual(valid_empty, report.valid_empty);
}

const s3c3_blank_cell = CCell{ .tag = @intFromEnum(CellTag.empty), .str_len = 0, .str_ptr = null, .i = 0, .f = 0, .b = 0, ._pad = [_]u8{0} ** 7 };

fn s3cStringCell(s: []const u8) CCell {
    return .{ .tag = @intFromEnum(CellTag.string), .str_len = @intCast(s.len), .str_ptr = s.ptr, .i = 0, .f = 0, .b = 0, ._pad = [_]u8{0} ** 7 };
}

/// The S3c fixture with the `title` coverage written over the read's
/// own hashes (every slot fresh), saved to `name` — the set every
/// sweep test starts from.
fn writeS3c3Embedded(io: std.Io, tt: *TestTmp, src_name: []const u8, name: []const u8) ![:0]u8 {
    const alloc = std.testing.allocator;
    const src = try writeS3cFixture(io, tt, src_name);
    defer alloc.free(src);
    const path = try tt.path(alloc, io, name);
    errdefer alloc.free(path);
    var err_buf: [128]u8 = undefined;
    const ed = zlsx_editor_open(src.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
    defer zlsx_editor_close(ed);
    const hashes = [_]u64{ try s3cHash(2, .{ .string = "alpha" }), try s3cHash(3, .{ .string = "beta" }), try s3cHash(4, .{ .string = "gamma" }) };
    const vecs = [_]f32{ 1, 2, 3, 4, 5, 6, 7, 8, 9 };
    const covs = [_]CEmbCoverage{s3cCoverage("title", 0, "A2:A4", "A", &vecs, &hashes)};
    if (s3cSet(ed, "m", 3, "f32", &covs, 0, null, &err_buf) != ZLSX_OK) return error.TestUnexpectedResult;
    if (zlsx_editor_save(ed, path.ptr, path.len, &err_buf, err_buf.len) != 0) return error.TestUnexpectedResult;
    return path;
}

/// The same set with the recovery record ALSO in its hidden cells
/// sheet (`recovery_in_cells`, the Zig-only opt-in) — three sheets,
/// the recovery sheet last.
fn writeS3c3EmbeddedInCells(io: std.Io, tt: *TestTmp, src_name: []const u8, name: []const u8) ![:0]u8 {
    const alloc = std.testing.allocator;
    const src = try writeS3cFixture(io, tt, src_name);
    defer alloc.free(src);
    const path = try tt.path(alloc, io, name);
    errdefer alloc.free(path);
    var wb = try zlsx_pkg.Workbook.open(alloc, io, src);
    defer wb.deinit();
    const hashes = [_]u64{ try s3cHash(2, .{ .string = "alpha" }), try s3cHash(3, .{ .string = "beta" }), try s3cHash(4, .{ .string = "gamma" }) };
    const vecs = [_]f32{ 1, 2, 3, 4, 5, 6, 7, 8, 9 };
    const body = try zlsx_pkg.embedding_part.encodeVectorBody(alloc, .f32, 3, &vecs);
    defer alloc.free(body);
    try wb.setEmbeddingsOpts("m", 3, .f32, &[_]zlsx_pkg.EmbeddingCoverageInput{.{
        .id = "title",
        .worksheet_target = "worksheets/sheet1.xml",
        .range = "A2:A4",
        .column = "A",
        .include_formulas = false,
        .vec_body = body,
        .hashes = &hashes,
    }}, .{ .recovery_in_cells = true });
    try wb.save(io, path);
    return path;
}

fn expectSameBytes(io: std.Io, a: []const u8, b: []const u8) !void {
    const alloc = std.testing.allocator;
    const ab = try readFileBytes(io, a);
    defer alloc.free(ab);
    const bb = try readFileBytes(io, b);
    defer alloc.free(bb);
    try std.testing.expectEqualSlices(u8, ab, bb);
}

test "S3c prune_embeddings: the read's hashes prune all fresh and rewrite nothing; a staged blank redacts its slot; the saved tombstone reads valid_empty" {
    const alloc = std.testing.allocator;
    var threaded: std.Io.Threaded = .init(alloc, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try writeS3c3Embedded(io, &tt, "s3c3_src.xlsx", "s3c3_emb.xlsx");
    defer alloc.free(path);
    const same_path = try tt.path(alloc, io, "s3c3_same.xlsx");
    defer alloc.free(same_path);
    const pruned_path = try tt.path(alloc, io, "s3c3_pruned.xlsx");
    defer alloc.free(pruned_path);
    var err_buf: [128]u8 = undefined;

    {
        const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
        defer zlsx_editor_close(ed);
        var report = poisonedPruneReport();
        var diag = freshDiag();
        try std.testing.expectEqual(ZLSX_OK, s3cPrune(ed, &report, &diag, &err_buf));
        try std.testing.expectEqual(plane_none, diag.plane);
        try std.testing.expectEqual(@as(usize, 0), diagName(&diag).len);
        zlsx_diag_release(&diag);
        try expectPruneCounts(report, 0, 0, 3, 0);
        // Nothing to redact, nothing rewritten: the save is the
        // untouched editor's passthrough, byte for byte.
        try std.testing.expectEqual(@as(i32, 0), zlsx_editor_save(ed, same_path.ptr, same_path.len, &err_buf, err_buf.len));
        try expectSameBytes(io, path, same_path);

        // A staged blank on a covered row is judged as staged: its
        // slot is redacted, the other two stay fresh; a staged string
        // is never fresh.
        try std.testing.expectEqual(@as(i32, 0), zlsx_editor_set_cell(ed, 0, 3, 0, &s3c3_blank_cell, &err_buf, err_buf.len));
        const edited = s3cStringCell("gamma edited");
        try std.testing.expectEqual(@as(i32, 0), zlsx_editor_set_cell(ed, 0, 4, 0, &edited, &err_buf, err_buf.len));
        report = poisonedPruneReport();
        try std.testing.expectEqual(ZLSX_OK, s3cPrune(ed, &report, null, &err_buf));
        try expectPruneCounts(report, 1, 1, 1, 0);
        try std.testing.expectEqual(@as(i32, 0), zlsx_editor_save(ed, pruned_path.ptr, pruned_path.len, &err_buf, err_buf.len));
    }
    // The saved file: slot 1 is a tombstone over a zeroed vector, the
    // others untouched.
    {
        const emb = zlsx_emb_open(pruned_path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
        defer zlsx_emb_close(emb);
        try std.testing.expectEqual(ZLSX_EMB_PRESENT, zlsx_emb_state(emb));
        var hashes: [3]u64 = undefined;
        try std.testing.expectEqual(@as(i32, 0), zlsx_emb_hashes(emb, 0, &hashes, hashes.len));
        try std.testing.expectEqual(try s3cHash(2, .{ .string = "alpha" }), hashes[0]);
        try std.testing.expectEqual(zlsx_emb_tombstone(), hashes[1]);
        try std.testing.expectEqual(try s3cHash(4, .{ .string = "gamma" }), hashes[2]);
        var vecs: [9]f32 = undefined;
        try std.testing.expectEqual(@as(i32, 0), zlsx_emb_vectors(emb, 0, &vecs, vecs.len));
        try std.testing.expectEqualSlices(f32, &.{ 1, 2, 3, 0, 0, 0, 7, 8, 9 }, &vecs);
    }
    // Re-opened: the tombstone over the blank is valid_empty, the
    // edited row stale (its hash is the old text's), alpha fresh.
    {
        const ed = zlsx_editor_open(pruned_path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
        defer zlsx_editor_close(ed);
        var report = poisonedPruneReport();
        try std.testing.expectEqual(ZLSX_OK, s3cPrune(ed, &report, null, &err_buf));
        try expectPruneCounts(report, 0, 1, 1, 1);
    }
}

test "S3c prune_embeddings: a row blanked on disk redacts, an edited row is stale and left alone, a workbook without a set is all zeros" {
    const alloc = std.testing.allocator;
    var threaded: std.Io.Threaded = .init(alloc, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    var err_buf: [128]u8 = undefined;
    const sheet_part = "xl/worksheets/sheet1.xml";

    // Excel's shape: the cell is gone from the saved part.
    {
        const path = try writeS3c3Embedded(io, &tt, "s3c3_b_src.xlsx", "s3c3_blanked.xlsx");
        defer alloc.free(path);
        try zlsx_pkg.pivots.fixture.patchPart(alloc, io, path, sheet_part, "<c r=\"A3\" t=\"s\"><v>4</v></c>", "<c r=\"A3\"/>");
        const out = try tt.path(alloc, io, "s3c3_blanked_out.xlsx");
        defer alloc.free(out);
        {
            const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
            defer zlsx_editor_close(ed);
            var report = poisonedPruneReport();
            try std.testing.expectEqual(ZLSX_OK, s3cPrune(ed, &report, null, &err_buf));
            try expectPruneCounts(report, 1, 0, 2, 0);
            try std.testing.expectEqual(@as(i32, 0), zlsx_editor_save(ed, out.ptr, out.len, &err_buf, err_buf.len));
        }
        const emb = zlsx_emb_open(out.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
        defer zlsx_emb_close(emb);
        var hashes: [3]u64 = undefined;
        try std.testing.expectEqual(@as(i32, 0), zlsx_emb_hashes(emb, 0, &hashes, hashes.len));
        try std.testing.expectEqual(zlsx_emb_tombstone(), hashes[1]);
    }
    // An edit: stale, not redacted, and nothing rewritten.
    {
        const path = try writeS3c3Embedded(io, &tt, "s3c3_e_src.xlsx", "s3c3_edited.xlsx");
        defer alloc.free(path);
        try zlsx_pkg.pivots.fixture.patchPart(alloc, io, path, sheet_part, "<c r=\"A3\" t=\"s\"><v>4</v></c>", "<c r=\"A3\" t=\"s\"><v>5</v></c>");
        const out = try tt.path(alloc, io, "s3c3_edited_out.xlsx");
        defer alloc.free(out);
        const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
        defer zlsx_editor_close(ed);
        var report = poisonedPruneReport();
        try std.testing.expectEqual(ZLSX_OK, s3cPrune(ed, &report, null, &err_buf));
        try expectPruneCounts(report, 0, 1, 2, 0);
        try std.testing.expectEqual(@as(i32, 0), zlsx_editor_save(ed, out.ptr, out.len, &err_buf, err_buf.len));
        try expectSameBytes(io, path, out);
    }
    // No set at all: every count 0, the diag as prep left it.
    {
        const plain = try writeS3cFixture(io, &tt, "s3c3_plain.xlsx");
        defer alloc.free(plain);
        const ed = zlsx_editor_open(plain.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
        defer zlsx_editor_close(ed);
        var report = poisonedPruneReport();
        var diag = freshDiag();
        try std.testing.expectEqual(ZLSX_OK, s3cPrune(ed, &report, &diag, &err_buf));
        try std.testing.expectEqual(plane_none, diag.plane);
        zlsx_diag_release(&diag);
        try expectPruneCounts(report, 0, 0, 0, 0);
    }
}

test "S3c prune_embeddings: statements about the call — NULL handle, NULL report, a rejected struct_size on either output, staged appends on the covered sheet" {
    const alloc = std.testing.allocator;
    var threaded: std.Io.Threaded = .init(alloc, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try writeS3c3Embedded(io, &tt, "s3c3_c_src.xlsx", "s3c3_call.xlsx");
    defer alloc.free(path);
    var err_buf: [128]u8 = undefined;
    const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
    defer zlsx_editor_close(ed);

    var report = poisonedPruneReport();
    try std.testing.expectEqual(ZLSX_ERROR, s3cPrune(null, &report, null, &err_buf));
    try std.testing.expectEqualStrings("InvalidInput", std.mem.sliceTo(&err_buf, 0));
    // Prepped (zeroed) before the handle was judged: the caller can
    // read it either way.
    try expectPruneCounts(report, 0, 0, 0, 0);

    try std.testing.expectEqual(ZLSX_ERROR, s3cPrune(ed, null, null, &err_buf));
    try std.testing.expectEqualStrings("NullOutPointer", std.mem.sliceTo(&err_buf, 0));

    // A report too small: untouched byte for byte; the diag beside it
    // still prepped.
    {
        var small = poisonedPruneReport();
        small.struct_size = @sizeOf(CPruneReport) - 1;
        const before = std.mem.toBytes(small);
        var diag = freshDiag();
        try std.testing.expectEqual(ZLSX_ERROR, s3cPrune(ed, &small, &diag, &err_buf));
        try std.testing.expectEqualStrings("StructSizeTooSmall", std.mem.sliceTo(&err_buf, 0));
        try std.testing.expectEqualSlices(u8, &before, &std.mem.toBytes(small));
        try std.testing.expectEqual(plane_none, diag.plane);
        zlsx_diag_release(&diag);
    }
    // A diag too small: the report, accepted, is zeroed.
    {
        var diag = freshDiag();
        diag.struct_size = @sizeOf(CDiag) - 1;
        report = poisonedPruneReport();
        try std.testing.expectEqual(ZLSX_ERROR, s3cPrune(ed, &report, &diag, &err_buf));
        try std.testing.expectEqualStrings("StructSizeTooSmall", std.mem.sliceTo(&err_buf, 0));
        try expectPruneCounts(report, 0, 0, 0, 0);
    }
    // Appended rows on the OTHER sheet do not touch the sweep; on the
    // covered sheet they refuse it, -1, before anything is judged.
    {
        const cells = [_]CCell{s3cStringCell("more")};
        try std.testing.expectEqual(@as(i32, 0), zlsx_editor_append_row(ed, 1, &cells, cells.len, &err_buf, err_buf.len));
        report = poisonedPruneReport();
        try std.testing.expectEqual(ZLSX_OK, s3cPrune(ed, &report, null, &err_buf));
        try expectPruneCounts(report, 0, 0, 3, 0);
        try std.testing.expectEqual(@as(i32, 0), zlsx_editor_append_row(ed, 0, &cells, cells.len, &err_buf, err_buf.len));
        var diag = freshDiag();
        report = poisonedPruneReport();
        try std.testing.expectEqual(ZLSX_ERROR, s3cPrune(ed, &report, &diag, &err_buf));
        try std.testing.expectEqualStrings("SheetHasUnsavedAppends", std.mem.sliceTo(&err_buf, 0));
        try std.testing.expectEqual(plane_none, diag.plane);
        try expectPruneCounts(report, 0, 0, 0, 0);
        zlsx_diag_release(&diag);
    }
}

test "S3c prune_embeddings: statements about the workbook — a set the index read refuses, a part it names gone, a covered sheet the read cannot serve; nothing staged after a refusal" {
    const alloc = std.testing.allocator;
    var threaded: std.Io.Threaded = .init(alloc, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    var err_buf: [128]u8 = undefined;
    const Case = struct { name: []const u8, part: []const u8, old: []const u8, new: []const u8, want: []const u8, read_name: ?[]const u8 };
    const cases = [_]Case{
        // The parser's own name (`InvalidRange`) is what the read
        // reports; the sweep folds it under the set's.
        .{ .name = "s3c3_index_range.xlsx", .part = "xl/zlsxEmbeddings/index.xml", .old = "range=\"A2:A4\"", .new = "range=\"A0:A4\"", .want = "MalformedEmbeddingSet", .read_name = "InvalidRange" },
        .{ .name = "s3c3_index_dtype.xlsx", .part = "xl/zlsxEmbeddings/index.xml", .old = "dtype=\"f32\"", .new = "dtype=\"f99\"", .want = "MalformedEmbeddingSet", .read_name = "InvalidDtype" },
        .{ .name = "s3c3_vec_magic.xlsx", .part = "xl/zlsxEmbeddings/title/vec.bin", .old = "ZVEC", .new = "ZVEX", .want = "MalformedEmbeddingSet", .read_name = "BadMagic" },
        .{ .name = "s3c3_rels_target.xlsx", .part = "xl/zlsxEmbeddings/_rels/index.xml.rels", .old = "Target=\"title/vec.bin\"", .new = "Target=\"title/none.bin\"", .want = "MissingEmbeddingPart", .read_name = "MissingEmbeddingPart" },
        .{ .name = "s3c3_sheet_no_r.xlsx", .part = "xl/worksheets/sheet1.xml", .old = "<c r=\"A3\"", .new = "<c", .want = "MalformedSheetXml", .read_name = null },
        .{ .name = "s3c3_sheet_bool.xlsx", .part = "xl/worksheets/sheet1.xml", .old = "<c r=\"A3\" t=\"s\"><v>4</v>", .new = "<c r=\"A3\" t=\"b\"><v>TRUE</v>", .want = "UnsupportedCellValue", .read_name = null },
    };
    for (cases, 0..) |case, i| {
        const src_name = try std.fmt.allocPrint(alloc, "s3c3_w_src{d}.xlsx", .{i});
        defer alloc.free(src_name);
        const path = try writeS3c3Embedded(io, &tt, src_name, case.name);
        defer alloc.free(path);
        try zlsx_pkg.pivots.fixture.patchPart(alloc, io, path, case.part, case.old, case.new);
        const out_name = try std.fmt.allocPrint(alloc, "s3c3_w_out{d}.xlsx", .{i});
        defer alloc.free(out_name);
        const out = try tt.path(alloc, io, out_name);
        defer alloc.free(out);
        const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
        defer zlsx_editor_close(ed);
        var report = poisonedPruneReport();
        var diag = freshDiag();
        try std.testing.expectEqual(ZLSX_REFUSED, s3cPrune(ed, &report, &diag, &err_buf));
        try std.testing.expectEqualStrings(case.want, diagName(&diag));
        try std.testing.expectEqualStrings(case.want, std.mem.sliceTo(&err_buf, 0));
        try std.testing.expectEqual(plane_none, diag.plane);
        zlsx_diag_release(&diag);
        try expectPruneCounts(report, 0, 0, 0, 0);
        // Refused before the first part write: the save is the passthrough.
        try std.testing.expectEqual(@as(i32, 0), zlsx_editor_save(ed, out.ptr, out.len, &err_buf, err_buf.len));
        try expectSameBytes(io, path, out);
        if (case.read_name) |read_name| {
            try std.testing.expect(zlsx_emb_open(path.ptr, &err_buf, err_buf.len) == null);
            try std.testing.expectEqualStrings(read_name, std.mem.sliceTo(&err_buf, 0));
        }
    }
}

test "S3c strip_embeddings: a set strips to ABSENT with every carrier gone; idempotent; a workbook without a set is untouched" {
    const alloc = std.testing.allocator;
    var threaded: std.Io.Threaded = .init(alloc, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    var err_buf: [128]u8 = undefined;
    const path = try writeS3c3Embedded(io, &tt, "s3c3_s_src.xlsx", "s3c3_strip.xlsx");
    defer alloc.free(path);
    const out = try tt.path(alloc, io, "s3c3_stripped.xlsx");
    defer alloc.free(out);
    {
        const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
        defer zlsx_editor_close(ed);
        var diag = freshDiag();
        try std.testing.expectEqual(ZLSX_OK, s3cStrip(ed, &diag, &err_buf));
        try std.testing.expectEqual(plane_none, diag.plane);
        try std.testing.expectEqual(@as(usize, 0), diagName(&diag).len);
        zlsx_diag_release(&diag);
        // Twice: the second finds nothing and says so with ZLSX_OK.
        try std.testing.expectEqual(ZLSX_OK, s3cStrip(ed, null, &err_buf));
        try std.testing.expectEqual(@as(i32, 0), zlsx_editor_save(ed, out.ptr, out.len, &err_buf, err_buf.len));
    }
    {
        const emb = zlsx_emb_open(out.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
        defer zlsx_emb_close(emb);
        try std.testing.expectEqual(ZLSX_EMB_ABSENT, zlsx_emb_state(emb));
    }
    // Every carrier: no part under the directory, no relationship, no
    // hidden name, no docProps part (nothing else lived in it).
    {
        var wb = try zlsx_pkg.Workbook.open(alloc, io, out);
        defer wb.deinit();
        for (wb.store.parts) |p| {
            try std.testing.expect(!std.mem.startsWith(u8, p.name, "xl/zlsxEmbeddings"));
        }
        const rels = (try wb.store.part("xl/_rels/workbook.xml.rels")) orelse return error.TestUnexpectedResult;
        try std.testing.expect(std.mem.indexOf(u8, rels.bytes, "zlsxEmbeddings") == null);
        const wbxml = (try wb.store.part("xl/workbook.xml")) orelse return error.TestUnexpectedResult;
        try std.testing.expect(std.mem.indexOf(u8, wbxml.bytes, "_zlsxRecovery") == null);
        try std.testing.expect((try wb.store.part("docProps/custom.xml")) == null);
        try std.testing.expectEqual(@as(u32, 2), wb.sheetCount());
    }
    // The cells are what they were: the read still serves them.
    {
        const ed = zlsx_editor_open(out.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
        defer zlsx_editor_close(ed);
        var out_ptr: ?[*]u8 = null;
        var out_len: usize = 0;
        try std.testing.expectEqual(ZLSX_OK, s3cRows(ed, 0, "A2:A4", "A", 0, &out_ptr, &out_len, null, &err_buf));
        defer zlsx_buffer_release(out_ptr, out_len);
        try std.testing.expect(std.mem.indexOf(u8, out_ptr.?[0..out_len], "\"text\":\"beta\"") != null);
        // A strip on a workbook without a set: OK and the passthrough.
        const same = try tt.path(alloc, io, "s3c3_stripped_again.xlsx");
        defer alloc.free(same);
        try std.testing.expectEqual(ZLSX_OK, s3cStrip(ed, null, &err_buf));
        try std.testing.expectEqual(@as(i32, 0), zlsx_editor_save(ed, same.ptr, same.len, &err_buf, err_buf.len));
        try expectSameBytes(io, out, same);
    }
}

test "S3c strip_embeddings: the recovery_in_cells sheet goes through the editor's delete — indices stay honest after the strip, and a dirty editor refuses before the first removal" {
    const alloc = std.testing.allocator;
    var threaded: std.Io.Threaded = .init(alloc, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    var err_buf: [128]u8 = undefined;
    const path = try writeS3c3EmbeddedInCells(io, &tt, "s3c3_cells_src.xlsx", "s3c3_cells.xlsx");
    defer alloc.free(path);
    {
        var wb = try zlsx_pkg.Workbook.open(alloc, io, path);
        defer wb.deinit();
        try std.testing.expectEqual(@as(u32, 3), wb.sheetCount());
        try std.testing.expectEqual(@as(?u32, 2), wb.recoveryCellSheetIndex());
    }
    const out = try tt.path(alloc, io, "s3c3_cells_out.xlsx");
    defer alloc.free(out);
    {
        const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
        defer zlsx_editor_close(ed);
        var diag = freshDiag();
        try std.testing.expectEqual(ZLSX_OK, s3cStrip(ed, &diag, &err_buf));
        zlsx_diag_release(&diag);
        // The mirror agrees with the workbook: the next sheet is 2,
        // and a write to it lands on it.
        var idx: u32 = no_sheet_idx;
        try std.testing.expectEqual(ZLSX_OK, zlsx_editor_add_sheet(ed, "New", 3, &idx, null, &err_buf, err_buf.len));
        try std.testing.expectEqual(@as(u32, 2), idx);
        const x = s3cStringCell("x");
        try std.testing.expectEqual(@as(i32, 0), zlsx_editor_set_cell(ed, 2, 1, 0, &x, &err_buf, err_buf.len));
        try std.testing.expectEqual(@as(i32, 0), zlsx_editor_save(ed, out.ptr, out.len, &err_buf, err_buf.len));
    }
    {
        const emb = zlsx_emb_open(out.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
        defer zlsx_emb_close(emb);
        try std.testing.expectEqual(ZLSX_EMB_ABSENT, zlsx_emb_state(emb));
        var wb = try zlsx_pkg.Workbook.open(alloc, io, out);
        defer wb.deinit();
        try std.testing.expectEqual(@as(u32, 3), wb.sheetCount());
        try std.testing.expectEqual(@as(?u32, null), wb.recoveryCellSheetIndex());
        try std.testing.expect((try wb.sheetByName("New")) != null);
        const ed = zlsx_editor_open(out.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
        defer zlsx_editor_close(ed);
        var out_ptr: ?[*]u8 = null;
        var out_len: usize = 0;
        try std.testing.expectEqual(ZLSX_OK, s3cRows(ed, 2, "A1:A1", "A", 0, &out_ptr, &out_len, null, &err_buf));
        defer zlsx_buffer_release(out_ptr, out_len);
        try std.testing.expect(std.mem.indexOf(u8, out_ptr.?[0..out_len], "\"row\":1,\"text\":\"x\"") != null);
    }
    // Dirty: the delete's pre-flight, -1, and nothing removed — the set
    // and the sheet are still there after the save.
    {
        const dirty_out = try tt.path(alloc, io, "s3c3_cells_dirty.xlsx");
        defer alloc.free(dirty_out);
        const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
        defer zlsx_editor_close(ed);
        const z = s3cStringCell("z");
        try std.testing.expectEqual(@as(i32, 0), zlsx_editor_set_cell(ed, 1, 5, 0, &z, &err_buf, err_buf.len));
        var diag = freshDiag();
        try std.testing.expectEqual(ZLSX_ERROR, s3cStrip(ed, &diag, &err_buf));
        try std.testing.expectEqualStrings("SheetDeleteRequiresCleanState", std.mem.sliceTo(&err_buf, 0));
        try std.testing.expectEqual(plane_none, diag.plane);
        zlsx_diag_release(&diag);
        try std.testing.expectEqual(@as(i32, 0), zlsx_editor_save(ed, dirty_out.ptr, dirty_out.len, &err_buf, err_buf.len));
        const emb = zlsx_emb_open(dirty_out.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
        defer zlsx_emb_close(emb);
        try std.testing.expectEqual(ZLSX_EMB_PRESENT, zlsx_emb_state(emb));
        var wb = try zlsx_pkg.Workbook.open(alloc, io, dirty_out);
        defer wb.deinit();
        try std.testing.expectEqual(@as(?u32, 2), wb.recoveryCellSheetIndex());
    }
}

test "S3c strip_embeddings: an xl/workbook.xml the chunk-name strip cannot walk refuses before the first removal — with and without the cells sheet" {
    const alloc = std.testing.allocator;
    var threaded: std.Io.Threaded = .init(alloc, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    var err_buf: [128]u8 = undefined;
    const junk_old = "</workbook>";
    const junk_new = "<definedName name=\"x>oops</definedName></workbook>";
    {
        const path = try writeS3c3Embedded(io, &tt, "s3c3_wx_src.xlsx", "s3c3_wx.xlsx");
        defer alloc.free(path);
        try zlsx_pkg.pivots.fixture.patchPart(alloc, io, path, "xl/workbook.xml", junk_old, junk_new);
        const out = try tt.path(alloc, io, "s3c3_wx_out.xlsx");
        defer alloc.free(out);
        const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
        defer zlsx_editor_close(ed);
        var diag = freshDiag();
        try std.testing.expectEqual(ZLSX_REFUSED, s3cStrip(ed, &diag, &err_buf));
        try std.testing.expectEqualStrings("MalformedWorkbookXml", diagName(&diag));
        try std.testing.expectEqual(plane_none, diag.plane);
        zlsx_diag_release(&diag);
        // Nothing removed: the passthrough save.
        try std.testing.expectEqual(@as(i32, 0), zlsx_editor_save(ed, out.ptr, out.len, &err_buf, err_buf.len));
        try expectSameBytes(io, path, out);
    }
    {
        const path = try writeS3c3EmbeddedInCells(io, &tt, "s3c3_wxc_src.xlsx", "s3c3_wxc.xlsx");
        defer alloc.free(path);
        try zlsx_pkg.pivots.fixture.patchPart(alloc, io, path, "xl/workbook.xml", junk_old, junk_new);
        const out = try tt.path(alloc, io, "s3c3_wxc_out.xlsx");
        defer alloc.free(out);
        const ed = zlsx_editor_open(path.ptr, &err_buf, err_buf.len) orelse return error.TestUnexpectedResult;
        defer zlsx_editor_close(ed);
        var diag = freshDiag();
        try std.testing.expectEqual(ZLSX_REFUSED, s3cStrip(ed, &diag, &err_buf));
        try std.testing.expectEqualStrings("MalformedWorkbookXml", diagName(&diag));
        zlsx_diag_release(&diag);
        // The cells sheet was not deleted ahead of the verdict.
        try std.testing.expectEqual(@as(i32, 0), zlsx_editor_save(ed, out.ptr, out.len, &err_buf, err_buf.len));
        try expectSameBytes(io, path, out);
    }
}

test "S3c sweeps: the plane of every verdict; the parser's own names stay -1; WriteFailed crosses as -3" {
    for ([_]anyerror{ error.MalformedEmbeddingSet, error.MissingEmbeddingPart, error.MalformedWorkbookXml, error.CannotDeleteLastSheet, error.MalformedSheetXml, error.UnsupportedCellValue, error.MissingContentTypes }) |e| {
        try std.testing.expectEqual(ZLSX_REFUSED, statusOf(e));
    }
    for ([_]anyerror{ error.CountMismatch, error.InvalidRange, error.InvalidDtype, error.BadMagic, error.InvalidEmbeddingInput, error.SheetHasUnsavedAppends, error.SheetDeleteRequiresCleanState, error.WriteFailed }) |e| {
        try std.testing.expectEqual(ZLSX_ERROR, statusOf(e));
    }
    var err_buf: [64]u8 = undefined;
    try std.testing.expectEqual(ZLSX_NOMEM, failSweep(error.WriteFailed, null, &err_buf, err_buf.len));
    try std.testing.expectEqualStrings("OutOfMemory", std.mem.sliceTo(&err_buf, 0));
    var diag = freshDiag();
    try std.testing.expectEqual(ZLSX_REFUSED, failSweep(error.MalformedEmbeddingSet, &diag, &err_buf, err_buf.len));
    try std.testing.expectEqualStrings("MalformedEmbeddingSet", diagName(&diag));
    zlsx_diag_release(&diag);
}
