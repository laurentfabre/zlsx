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
const xlsx = @import("xlsx.zig");

pub const ZLSX_ABI_VERSION: u32 = 1;
pub const ZLSX_VERSION_STRING: [*:0]const u8 = "0.2.0-dev";

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
    // reading state.inner.sheets. errdefer unrefs on any failure path.
    _ = state.refcount.fetchAdd(1, .acq_rel);
    errdefer state.unref();

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
