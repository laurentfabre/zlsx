//! B0 phase 1 (post-0.2.9 roadmap): PartStore — read-only OOXML
//! package layer.
//!
//! Opens a .xlsx ZIP archive, walks the central directory,
//! decompresses every part eagerly, parses `[Content_Types].xml` and
//! every `_rels/*.rels`, and exposes a typed read-only API:
//!
//!   - `partNames()` — central-directory order, source-preserving.
//!   - `part(name)` — `Part { name, content_type, bytes, compression_method }`.
//!   - `rels(owner)` — `[]Relationship`.
//!   - `resolve(owner, target)` — collapse `..`/`.` segments per OPC.
//!
//! Out of scope for this milestone: `save`, `replacePart`, `addPart`,
//! dirty flags, typed overlays for known parts. Those land in B0 M2/M3
//! per docs/plans/post-0.2.9-roadmap.md.
//!
//! Relationship to `Editor` (src/xlsx.zig): the Editor has a battle-
//! tested ZIP scanner that supports the byte-preserving save path; we
//! deliberately keep PartStore independent of Editor's internals here
//! to ship M1 without disturbing the Editor surface. M2 will extract
//! a shared scanner (planned src/package/zip.zig) so both modules
//! consume the same code.

const std = @import("std");
const atomic_file_mod = @import("atomic_file.zig");
const AtomicFile = atomic_file_mod.AtomicFile;
pub const Commit = atomic_file_mod.Commit;

// M5d1: §5.5's cancellation / deadline / 64 KiB-chunk seam. Named import
// so `pkg/control.zig` belongs to exactly one module tree.
const control = @import("zlsx_control");
pub const Poller = control.Poller;

// §9.1 M10q's fill accounting. Zero cost with no tally installed: the
// arena's own allocator is returned untouched.
const fill = @import("fill.zig");

/// §5.7.9's one seam *inside* the commit region (M5d2).
///
/// The ordering the spec makes normative is rename → swap → directory
/// fsync, and the swap is the caller's: `saveWithRecalc` installs a
/// prepared generation the instant the rename has published the bytes it
/// was serialized from. That instant is inside `saveControlled`, between
/// two statements, so the only way to reach it is a callback.
///
/// **Infallible by signature, and that is the whole design.** §5.7.9 says
/// nothing after the commit point may report failure as an error;
/// `AtomicFile.syncDir` obeys that by returning a `Commit` instead of an
/// error union, and this obeys it by returning `void`. A hook that could
/// fail would put a second failure mode after the point where there is no
/// longer a failure to report — and `recalc_txn.Candidate.swap` is
/// already no-fail, so nothing is being given up.
pub const CommitHook = struct {
    ctx: ?*anyopaque = null,
    call: ?*const fn (?*anyopaque) void = null,

    /// What every plain save passes: the rename is the last thing that
    /// happens, and nothing needs to run between it and the fsync.
    pub const none: CommitHook = .{};

    pub inline fn fire(self: CommitHook) void {
        const f = self.call orelse return;
        f(self.ctx);
    }
};

/// Read exactly `dest.len` bytes at `offset`.
///
/// Zig 0.16 removed `File.seekTo` / `File.readAll`; positional reads go
/// through a `File.Reader`. An empty reader buffer is deliberate — it
/// makes `readSliceAll` read straight into `dest` with no intermediate
/// copy, which is what the old `seekTo` + `readAll` pair did.
///
/// A short read was `n != expected` -> BadZip before; it is now
/// `error.EndOfStream`, mapped to the same BadZip so the archive-level
/// contract is unchanged.
fn readAtExact(file: std.Io.File, io: std.Io, offset: u64, dest: []u8) Error!void {
    var fr = file.reader(io, &.{});
    fr.seekTo(offset) catch return Error.BadZip;
    fr.interface.readSliceAll(dest) catch return Error.BadZip;
}

pub const Part = struct {
    name: []const u8,
    /// Resolved via `[Content_Types].xml` (Override > Default by
    /// extension). `null` when neither matches — callers can treat
    /// that as "unknown content type".
    content_type: ?[]const u8,
    /// Decompressed part bytes. Owned by the PartStore arena;
    /// borrowed by the caller for the PartStore's lifetime.
    ///
    /// **Lazy materialization (since iter-wb-6 ratio gate fix):**
    /// `bytes.len` is 0 for parts with `uncompressed_size > 0` that
    /// haven't been accessed yet. `PartStore.part(name)` decompresses
    /// + caches on first call. Inherently-empty parts (size == 0)
    /// stay empty across the boundary.
    /// Always-materialized parts: `[Content_Types].xml` and every
    /// `_rels/*.rels` (needed at open() for content-type / relationship
    /// resolution).
    bytes: []const u8,
    /// ZIP compression method as recorded in the central directory
    /// (0 = stored, 8 = deflate). Useful for callers that want to
    /// re-emit byte-for-byte later.
    ///
    /// **Stale while a replacement is pending.** M10b defers the
    /// deflate to the save that needs it, so between `replacePart` and
    /// that save this field still names the method of the bytes the
    /// replacement *displaced* — the method for the new bytes does not
    /// exist yet, because no compressor has chosen it. Re-emission is a
    /// post-save concern and reads a coherent pair there; a caller that
    /// wants the pair coherent earlier has to save first.
    compression_method: u16,

    // ─── Lazy-materialization metadata ────────────────────────────────
    // These fields let `PartStore.part()` decompress + verify on first
    // access. Internal-only for callers (read for resilience only).
    payload_offset: u32 = 0,
    compressed_size: u32 = 0,
    uncompressed_size: u32 = 0,
    crc32: u32 = 0,
};

pub const TargetMode = enum { internal, external };

pub const Relationship = struct {
    id: []const u8,
    type: []const u8,
    target: []const u8,
    target_mode: TargetMode,
};

pub const Error = error{
    NotPkzip,
    BadZip,
    UnsupportedCompression,
    Zip64NotSupported,
    EncryptedNotSupported,
    SplitArchiveNotSupported,
    DataDescriptorNotSupported,
    DuplicateContentType,
    PartNotFound,
    PartAlreadyExists,
    MissingContentTypes,
    MalformedContentTypes,
    /// save() rejects archives whose total saved size would exceed
    /// 4 GiB (ZIP32 offset limit) or whose entry count would exceed
    /// 65 535 (EOCD u16 record-count field). Surfaced before any
    /// bytes are written so the atomic-file output is never left in
    /// a partial state.
    ZipArchiveTooLarge,
} || std.Io.File.OpenError || std.mem.Allocator.Error || std.Io.File.Reader.Error ||
    std.Io.File.SeekError || std.Io.File.StatError || std.Io.File.SyncError ||
    // §5.5's cooperative cancellation (M5d1). Note the spelling: this is
    // NOT `std.Io.Cancelable`'s `Canceled`, which the sets above already
    // contribute and which means "the Io runtime cancelled the syscall".
    // This one means the caller's token fired.
    control.Error;

/// Observability seam for the fd-budget invariant. §5.7.4 gates
/// repeated recalc on "RSS, allocation accounting, borrow validity, and
/// fd count", and the fd count is the one of those four that cannot be
/// read off the backing itself — by the time the answer is interesting
/// the backing has been destroyed. So the counter lives with the
/// caller: a `SourceBacking` bumps `closes` exactly once, when its last
/// reference drops. Production constructors pass `null`.
pub const CloseLedger = struct {
    closes: usize = 0,
};

/// Ref-counted source of a package's compressed bytes, shared by every
/// `PartStore` generation opened over it.
///
/// Before M5b0 a store held `?std.Io.File` and closed it in `deinit`,
/// which made the store the handle's *exclusive* owner: a shallow copy
/// closed it twice, and moving it out to keep the bytes reachable left
/// the original store unable to materialize anything. §5.7.4's
/// prepare/swap transaction retains the entire superseded generation
/// until `Workbook.deinit`, so several generations are alive at once by
/// construction — and neither of those outcomes is survivable.
///
/// Both variants answer `readAt` and `size`, so nothing above this type
/// branches on where the bytes live: `open`, `openBuffer` and
/// `nextGeneration` all funnel into `openOver`, and `materializeAt` and
/// `save` read through the same call.
///
/// Deliberately not atomic. A backing and every generation over it
/// belong to one thread — the same contract `PartStore` itself has —
/// and an atomic count would advertise a sharing discipline nothing
/// here tests.
pub const SourceBacking = struct {
    pub const Source = union(enum) {
        /// The source archive, kept open for the backing's lifetime.
        file: std.Io.File,
        /// A backing-owned copy of the source archive. `openBuffer`
        /// dupes the caller's slice, so the caller's borrow ends when
        /// the call returns.
        buffer: []const u8,
    };

    allocator: std.mem.Allocator,
    /// The `Io` the backing was created with. Closing a file needs one
    /// long after `open()` returned, and the generation that goes last
    /// is not necessarily the one that opened it — so the backing
    /// carries its own rather than borrowing a generation's.
    io: std.Io,
    source: Source,
    /// Live generations. Starts at 1, for the store the constructor
    /// hands the backing to.
    refs: usize,
    ledger: ?*CloseLedger,

    fn create(
        allocator: std.mem.Allocator,
        io: std.Io,
        source: Source,
        ledger: ?*CloseLedger,
    ) std.mem.Allocator.Error!*SourceBacking {
        const self = try allocator.create(SourceBacking);
        self.* = .{
            .allocator = allocator,
            .io = io,
            .source = source,
            .refs = 1,
            .ledger = ledger,
        };
        return self;
    }

    /// Adopt an already-open handle. The backing closes it on last
    /// release; the caller must not close it itself, and must not keep
    /// an `errdefer file.close` armed past this call — that pairing is
    /// the double-close this type exists to remove.
    fn createFile(
        allocator: std.mem.Allocator,
        io: std.Io,
        file: std.Io.File,
        ledger: ?*CloseLedger,
    ) std.mem.Allocator.Error!*SourceBacking {
        return create(allocator, io, .{ .file = file }, ledger);
    }

    /// Copy `bytes` into backing-owned memory. Duping rather than
    /// borrowing is what lets the caller's slice die at the call
    /// boundary, which is the borrow rule `Book` already follows for
    /// buffer-sourced workbooks.
    fn createBuffer(
        allocator: std.mem.Allocator,
        io: std.Io,
        bytes: []const u8,
        ledger: ?*CloseLedger,
    ) std.mem.Allocator.Error!*SourceBacking {
        const owned = try allocator.dupe(u8, bytes);
        errdefer allocator.free(owned);
        return create(allocator, io, .{ .buffer = owned }, ledger);
    }

    /// Take a reference for a new generation.
    pub fn retain(self: *SourceBacking) *SourceBacking {
        self.refs += 1;
        return self;
    }

    /// Drop one reference. The last one closes the file (or frees the
    /// buffer) and destroys the backing — exactly once, whichever
    /// generation happens to go last, in whatever order they go.
    pub fn release(self: *SourceBacking) void {
        std.debug.assert(self.refs > 0);
        self.refs -= 1;
        if (self.refs > 0) return;

        switch (self.source) {
            .file => |f| f.close(self.io),
            .buffer => |b| self.allocator.free(b),
        }
        if (self.ledger) |l| l.closes += 1;
        const allocator = self.allocator;
        allocator.destroy(self);
    }

    pub fn refCount(self: *const SourceBacking) usize {
        return self.refs;
    }

    /// Read exactly `dest.len` bytes at `offset`.
    ///
    /// Out of range is `BadZip` in both variants: for a file that is
    /// the short read `seekTo` + `readAll` already produced, and for a
    /// buffer it is the answer the old `self.file orelse return
    /// Error.BadZip` gave — a `fresh()` store's backing is the empty
    /// buffer, so asking it for source bytes it never had still
    /// refuses, without a null check anywhere above.
    pub fn readAt(self: *const SourceBacking, offset: u64, dest: []u8) Error!void {
        switch (self.source) {
            .file => |f| try readAtExact(f, self.io, offset, dest),
            .buffer => |b| {
                if (offset > b.len) return Error.BadZip;
                const start: usize = @intCast(offset);
                if (dest.len > b.len - start) return Error.BadZip;
                @memcpy(dest, b[start..][0..dest.len]);
            },
        }
    }

    /// Total source bytes. `openOver` gates on this for both variants,
    /// so the ZIP32 size refusal does not depend on where the archive
    /// came from. (`openBuffer` checks the same bound once more, up
    /// front, so an oversize slice is refused before it is copied.)
    fn size(self: *const SourceBacking) std.Io.File.StatError!u64 {
        return switch (self.source) {
            .file => |f| (try f.stat(self.io)).size,
            .buffer => |b| b.len,
        };
    }
};

pub const PartStore = struct {
    allocator: std.mem.Allocator,
    /// The `Io` this store was opened with. `save()` needs one for the
    /// atomic-file output long after `open()` returned, so the store
    /// carries its own. Source *reads* go through `backing.io`, which
    /// is the same value — the backing keeps a copy because it outlives
    /// any one generation.
    io: std.Io,
    /// Arena-owned storage for every borrowed slice exposed via the
    /// public API (part names, content types, rel attrs, decompressed
    /// part bytes). Caller-visible slices stay valid until `deinit`.
    arena: fill.Arena,
    /// Replaced payloads too large to take through `arena` (§9.1 M10m).
    ///
    /// A 0.16 arena sizes a new chunk at 1.5 × (what it already holds +
    /// what was asked for), so the recalc's one 8.37 MB sheet body
    /// bought a 20.27 MB chunk — 11.9 MB of headroom nothing else in
    /// the store would ever ask for, alive from the swap to `deinit`.
    /// Above the threshold the payload gets an exact block instead and
    /// is freed here. Same lifetime the arena gave it: valid until
    /// `deinit`, and a re-replace of the same part strands its
    /// predecessor exactly as an arena reset-free would.
    big_parts: std.ArrayListUnmanaged([]u8) = .empty,
    /// One reference to the source bytes, shared with every other
    /// generation opened over the same archive. The previous design
    /// slurped the entire file into `src_buf` at open(); now we read
    /// CD + EOCD + structural parts at open() then
    /// close-the-buffer-but-keep-the-backing so RSS doesn't retain
    /// compressed bytes for parts callers never touch. `materializeAt`
    /// and the byte-preserving `save()` path both re-read through it.
    ///
    /// Never null: a `PartStore.fresh()` store gets an empty-buffer
    /// backing rather than an absent one. Every part of a fresh store
    /// lives in `overrides[i]`, so no source read is ever reached; if
    /// one were, the backing refuses it the same way the old `?File`
    /// null check did.
    backing: *SourceBacking,
    /// Per-part raw ZIP entries (LFH/CDFH/payload offsets). Same
    /// length + ordering as `parts`.
    entries: []ZipEntry,
    parts: []Part,
    /// Per-part override (when caller has called `replacePart`).
    /// Same length as `parts`. `null` = use original src_buf bytes.
    overrides: []?Override,
    /// Map from part name → relationships parsed from
    /// `_rels/<owner>.rels`. Empty list when the part has no rels.
    rels_by_owner: std.StringHashMapUnmanaged([]Relationship),
    /// Trailing ZIP comment after the EOCD (typically empty).
    /// Preserved across save() so byte-identical comments survive.
    eocd_comment: []const u8,

    pub fn open(allocator: std.mem.Allocator, io: std.Io, path: []const u8) Error!PartStore {
        const file = try std.Io.Dir.cwd().openFile(io, path, .{});
        // Hand the handle to the backing immediately, and never arm an
        // `errdefer file.close` past this point: from here the backing
        // is the sole closer, and a second one is precisely the bug
        // M5b0 removes. The one hand-close below covers the only window
        // where the backing does not exist yet.
        const backing = SourceBacking.createFile(allocator, io, file, null) catch |err| {
            file.close(io);
            return err;
        };
        errdefer backing.release();
        return try openOver(allocator, backing, .none);
    }

    /// Open a package that is already in memory. The bytes are copied
    /// into the backing, so the caller's slice may die at the call
    /// boundary.
    ///
    /// This is the substrate for `Workbook.openBuffer` (M5b2), not that
    /// entry point itself: what lands here is the *backing* half — a
    /// buffer-sourced store that reaches every read, materialize and
    /// save path a file-sourced one does.
    pub fn openBuffer(allocator: std.mem.Allocator, io: std.Io, bytes: []const u8) Error!PartStore {
        return openBufferControlled(allocator, io, bytes, .none);
    }

    /// `openBuffer` with §5.5's poll seam armed (M5d1). Substrate for
    /// `Workbook.openBufferControlled` (§5.10).
    ///
    /// Opening a buffer is not free: the central directory is scanned,
    /// every structural part is decompressed eagerly, and the whole
    /// archive is copied into the backing first. On a multi-hundred-MiB
    /// package that is seconds of work happening *before* a recalc
    /// starts, which is exactly the window §5.10 says an orchestrator's
    /// deadline has to reach.
    ///
    /// Nothing is mutated on the way out: a cancelled open releases the
    /// backing and tears down the arena, so the caller's `bytes` and
    /// every other generation over the same archive are untouched.
    pub fn openBufferControlled(
        allocator: std.mem.Allocator,
        io: std.Io,
        bytes: []const u8,
        poller: Poller,
    ) Error!PartStore {
        // Before the dupe: cancelling should not first copy the archive.
        try poller.check();
        // Same refusal `openOver` makes for both variants, taken before
        // the copy rather than after it: there is no reason to dupe
        // 4 GiB in order to reject it.
        if (bytes.len > std.math.maxInt(u32)) return Error.Zip64NotSupported;
        const backing = try SourceBacking.createBuffer(allocator, io, bytes, null);
        errdefer backing.release();
        return try openOver(allocator, backing, poller);
    }

    /// Open another `PartStore` over the same source bytes.
    ///
    /// The new generation gets its own arena, parts, overrides and
    /// rels; it shares only the backing, whose reference count this
    /// call bumps. Both generations stay independently readable and may
    /// be `deinit`ed in any order — the fd (or buffer) is released
    /// once, by whichever goes last. That is what lets §5.7.4 retain a
    /// superseded generation until `Workbook.deinit` on an fd budget of
    /// one.
    ///
    /// Overrides are NOT inherited: a generation is the source archive
    /// as it was opened, not the previous generation's staged state.
    /// The transaction that stages new bytes on top (M5b2) is the
    /// caller's, not this primitive's.
    ///
    /// A `fresh()` store has no source archive to re-scan, so this
    /// returns `error.BadZip` for one — its empty backing has no EOCD.
    pub fn nextGeneration(self: *const PartStore) Error!PartStore {
        const backing = self.backing.retain();
        errdefer backing.release();
        return try openOver(self.allocator, backing, .none);
    }

    /// Shared tail of `open` / `openBuffer` / `nextGeneration`: scan and
    /// parse an archive out of a backing whose reference this call
    /// consumes on success. On failure the caller's `errdefer` releases.
    ///
    /// The only place a variant is named: a file has to be copied into
    /// a contiguous scratch buffer to be scanned, a buffer already is
    /// one. Everything downstream of `buildFromArchive` is identical.
    fn openOver(allocator: std.mem.Allocator, backing: *SourceBacking, poller: Poller) Error!PartStore {
        const size_u64 = try backing.size();
        if (size_u64 > std.math.maxInt(u32)) return Error.Zip64NotSupported;

        switch (backing.source) {
            // No scratch copy: the backing's own bytes ARE the
            // contiguous view, and they outlive this call.
            .buffer => |b| return try buildFromArchive(allocator, backing, b, poller),
            .file => {
                // Read the whole file into a SCRATCH buffer
                // (page-allocator, not arena). scanCentralDirectory
                // needs random access to every CDFH + each LFH header
                // to compute payload offsets; doing that without a
                // contiguous buffer would require hundreds of small
                // disk reads. We free the scratch buffer before
                // returning so RSS doesn't retain it. Pages are mmap'd
                // by `page_allocator`, so `free` returns them via
                // munmap rather than holding them inside the process
                // arena.
                const scratch = try std.heap.page_allocator.alloc(u8, @intCast(size_u64));
                defer std.heap.page_allocator.free(scratch);
                try readChunked(backing, 0, scratch, poller);
                return try buildFromArchive(allocator, backing, scratch, poller);
            },
        }
    }

    /// Parse `archive` (a contiguous view of the whole package) into a
    /// store over `backing`.
    fn buildFromArchive(
        allocator: std.mem.Allocator,
        backing: *SourceBacking,
        archive: []const u8,
        poller: Poller,
    ) Error!PartStore {
        var arena: fill.Arena = .init(allocator, .parts);
        errdefer arena.deinit();
        const ar_alloc = arena.allocator();
        const scratch = archive;

        const entries = try scanCentralDirectory(ar_alloc, scratch);

        // Lazy decompression (iter-wb-6 ratio gate fix): only the
        // structural parts ([Content_Types].xml + every _rels/*.rels)
        // are decompressed at open() — they're needed inline for
        // content-type / relationship resolution. Everything else
        // stays compressed-on-disk until `PartStore.part()` first
        // surfaces them via `materializeAt` (seek + readAll).
        const parts = try ar_alloc.alloc(Part, entries.len);
        for (entries, 0..) |e, i| {
            // Per entry, ahead of the decompress: a package with tens of
            // thousands of parts is a long operation even when each part
            // is small, and the eager pass touches every one of them.
            try poller.check();
            const eager = isStructuralPart(e.name) or e.uncompressed_size == 0;
            var bytes: []const u8 = &.{};
            if (eager) {
                const compressed = scratch[e.payload_offset .. e.payload_offset + e.compressed_size];
                bytes = try decompressPayload(
                    ar_alloc,
                    compressed,
                    e.compression_method,
                    e.uncompressed_size,
                    poller,
                );
                if (std.hash.Crc32.hash(bytes) != e.crc32) return Error.BadZip;
            }
            parts[i] = .{
                .name = e.name,
                .content_type = null, // resolved next pass
                .bytes = bytes,
                .compression_method = e.compression_method,
                .payload_offset = @intCast(e.payload_offset),
                .compressed_size = e.compressed_size,
                .uncompressed_size = e.uncompressed_size,
                .crc32 = e.crc32,
            };
        }

        try resolveContentTypes(ar_alloc, parts);

        var rels_by_owner: std.StringHashMapUnmanaged([]Relationship) = .empty;
        for (parts) |p| {
            const owner = (try relsOwner(ar_alloc, p.name)) orelse continue;
            const relationships = try parseRelationships(ar_alloc, p.bytes);
            try rels_by_owner.put(ar_alloc, owner, relationships);
        }

        // Recover the EOCD comment for byte-preserving save(). DUPE
        // into arena because the source bytes (scratch) are about to
        // be freed.
        const eocd_off = try findEocd(scratch);
        const comment_len = std.mem.readInt(u16, scratch[eocd_off + 20 ..][0..2], .little);
        const comment_start = eocd_off + eocd_min_size;
        const eocd_comment: []const u8 = if (comment_len > 0)
            try ar_alloc.dupe(u8, scratch[comment_start .. comment_start + comment_len])
        else
            &[_]u8{};

        const overrides = try ar_alloc.alloc(?Override, parts.len);
        for (overrides) |*o| o.* = null;

        return .{
            .allocator = allocator,
            .io = backing.io,
            .arena = arena,
            .backing = backing,
            .entries = entries,
            .parts = parts,
            .overrides = overrides,
            .rels_by_owner = rels_by_owner,
            .eocd_comment = eocd_comment,
        };
    }

    /// Construct a from-scratch PartStore with no source archive.
    /// Every part the caller wants in the output is added via
    /// `addPart`; nothing else is on disk yet. Save emits a fresh
    /// LFH+CDFH+payload for every entry — never reads from a source
    /// file (there is none).
    ///
    /// Seeds an empty `[Content_Types].xml` so subsequent
    /// `addPart(name, ct, bytes)` calls have a valid CT document to
    /// stage `<Override>` entries into. Without the seed, the very
    /// first `addPart` would fail with `error.MissingContentTypes`.
    ///
    /// Calling `save(path)` immediately after `fresh()` (no addPart)
    /// emits a 1-entry ZIP containing only the empty
    /// `[Content_Types].xml` — spec-legal but useless to OOXML
    /// readers (no workbook, no rels). This is intentional: the
    /// store doesn't enforce OOXML well-formedness; that's the
    /// caller's job (Workbook.create or similar).
    pub fn fresh(allocator: std.mem.Allocator, io: std.Io) Error!PartStore {
        // A fresh store still gets a backing — the empty-buffer one.
        // Keeping the field non-optional is what removes the "is there
        // a source?" branch from `materializeAt` and `save`: those
        // paths are unreachable for a fresh store anyway, and if the
        // invariant ever breaks, an out-of-range read on an empty
        // buffer refuses with the same `BadZip` the null check gave.
        const backing = try SourceBacking.createBuffer(allocator, io, &.{}, null);
        errdefer backing.release();

        var arena: fill.Arena = .init(allocator, .parts);
        errdefer arena.deinit();
        const ar_alloc = arena.allocator();

        // Minimum spec-legal [Content_Types].xml. No <Default> or
        // <Override> entries — addPart will append them lazily.
        const seed_ct_xml =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" ++
            "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" ++
            "</Types>";
        const owned_ct_name = try ar_alloc.dupe(u8, "[Content_Types].xml");
        const owned_ct_bytes = try ar_alloc.dupe(u8, seed_ct_xml);

        // STORED (method 0): the seed XML is well under the 1 KiB
        // deflate threshold used elsewhere in the file. Mirrors the
        // `bytes.len < 1024` branch in `addPart` / `replacePart`.
        const owned_payload = try ar_alloc.dupe(u8, seed_ct_xml);

        const entries = try ar_alloc.alloc(ZipEntry, 1);
        entries[0] = .{
            .name = owned_ct_name,
            .lfh_offset = 0,
            .lfh_total_len = 0,
            .cdfh_offset = 0,
            .cdfh_total_len = 0,
            .payload_offset = 0,
            .compressed_size = @intCast(owned_payload.len),
            .uncompressed_size = @intCast(seed_ct_xml.len),
            .compression_method = 0,
            .crc32 = std.hash.Crc32.hash(seed_ct_xml),
            .data_descriptor_len = 0,
            .has_data_descriptor = false,
        };
        const parts = try ar_alloc.alloc(Part, 1);
        parts[0] = .{
            .name = owned_ct_name,
            // [Content_Types].xml has no content_type of its own —
            // it IS the content-type registry. Matches the result of
            // `resolveContentTypes` for the same input.
            .content_type = null,
            .bytes = owned_ct_bytes,
            .compression_method = 0,
            .payload_offset = 0,
            .compressed_size = @intCast(owned_payload.len),
            .uncompressed_size = @intCast(seed_ct_xml.len),
            .crc32 = std.hash.Crc32.hash(seed_ct_xml),
        };
        const overrides = try ar_alloc.alloc(?Override, 1);
        overrides[0] = .{ .compressed = .{
            .payload = owned_payload,
            .compression_method = 0,
            .crc32 = std.hash.Crc32.hash(seed_ct_xml),
            .uncompressed_size = @intCast(seed_ct_xml.len),
        } };

        return .{
            .allocator = allocator,
            .io = io,
            .arena = arena,
            .backing = backing,
            .entries = entries,
            .parts = parts,
            .overrides = overrides,
            .rels_by_owner = .empty,
            .eocd_comment = &.{},
        };
    }

    /// Drop this generation. Releases one reference to the shared
    /// backing — the source handle is closed only if this was the last
    /// generation holding it, so tearing down generation N never
    /// invalidates generation N+1.
    pub fn deinit(self: *PartStore) void {
        self.backing.release();
        for (self.big_parts.items) |b| self.allocator.free(b);
        self.big_parts.deinit(self.allocator);
        self.arena.deinit();
    }

    /// Add a new part to the package. The part name must NOT already
    /// exist (use `replacePart` for that case). Updates
    /// `[Content_Types].xml` to declare the new part's content type
    /// via an `<Override>` entry, since most consumer tooling won't
    /// open a workbook whose parts aren't declared.
    ///
    /// The new part is held entirely in arena-owned memory; nothing
    /// is written to disk until the next `save()`.
    pub fn addPart(
        self: *PartStore,
        name: []const u8,
        content_type: []const u8,
        bytes: []const u8,
    ) !void {
        return self.addPartControlled(name, content_type, bytes, .none);
    }

    /// `addPart` with §5.5's poll seam (M5d1). Same atomicity: every
    /// fallible step runs before the commit block, so a cancelled add
    /// leaves the store exactly as it was.
    pub fn addPartControlled(
        self: *PartStore,
        name: []const u8,
        content_type: []const u8,
        bytes: []const u8,
        poller: Poller,
    ) !void {
        if (self.findIndex(name) != null) return error.PartAlreadyExists;
        // Strict `>=`: 0xFFFFFFFF is the Zip64 sentinel — emitting that
        // in compressed_size or uncompressed_size produces an archive
        // the reader treats as Zip64 and rejects.
        if (bytes.len >= std.math.maxInt(u32)) return Error.Zip64NotSupported;
        if (name.len > std.math.maxInt(u16)) return Error.ZipArchiveTooLarge;

        const ar_alloc = self.arena.allocator();

        // Atomicity: do every fallible allocation up front BEFORE
        // committing any field mutation. If any allocation fails,
        // the store stays unchanged (arena cleans the leftovers
        // up later). Only the final assignment block is infallible.
        const owned_name = try ar_alloc.dupe(u8, name);
        const owned_ct = try ar_alloc.dupe(u8, content_type);
        const owned_user_bytes = try ar_alloc.dupe(u8, bytes);

        // Compress user payload via the same policy as replacePart.
        var compressed: std.ArrayListUnmanaged(u8) = .empty;
        defer compressed.deinit(ar_alloc);
        const method = try compressInto(ar_alloc, bytes, &compressed, poller);
        if (compressed.items.len >= std.math.maxInt(u32)) return Error.Zip64NotSupported;
        const owned_payload = try compressed.toOwnedSlice(ar_alloc);

        // Stage the [Content_Types].xml update WITHOUT calling
        // replacePart — we don't want to commit that mutation until
        // we know the array reallocs below also succeed.
        const ct_staging = try self.stageContentTypeOverride(name, content_type, poller);
        const ct_idx = ct_staging.idx;
        const ct_new_part_bytes = ct_staging.new_part_bytes;
        const ct_new_override = ct_staging.new_override;

        // Synthetic ZipEntry for the new part — save() rebuilds
        // LFH/CDFH from the override slot, so source-offset fields
        // are unused.
        const new_entry: ZipEntry = .{
            .name = owned_name,
            .lfh_offset = 0,
            .lfh_total_len = 0,
            .cdfh_offset = 0,
            .cdfh_total_len = 0,
            .payload_offset = 0,
            .compressed_size = @intCast(owned_payload.len),
            .uncompressed_size = @intCast(bytes.len),
            .compression_method = method,
            .crc32 = std.hash.Crc32.hash(bytes),
            .data_descriptor_len = 0,
            .has_data_descriptor = false,
        };
        const new_part: Part = .{
            .name = owned_name,
            .content_type = owned_ct,
            .bytes = owned_user_bytes,
            .compression_method = method,
        };
        const new_override: Override = .{ .compressed = .{
            .payload = owned_payload,
            .compression_method = method,
            .crc32 = std.hash.Crc32.hash(bytes),
            .uncompressed_size = @intCast(bytes.len),
        } };

        // Grow the parallel arrays via alloc+copy rather than
        // realloc. realloc invalidates the source slice on
        // relocation, so a successful realloc on entries followed
        // by a failed realloc on parts would leave self.entries
        // pointing into freed memory. With alloc+copy the old
        // slices stay valid (the arena holds their backing storage
        // until deinit), so a later allocation failure leaves the
        // store fully unchanged.
        const grown_entries = try ar_alloc.alloc(ZipEntry, self.entries.len + 1);
        @memcpy(grown_entries[0..self.entries.len], self.entries);
        const grown_parts = try ar_alloc.alloc(Part, self.parts.len + 1);
        @memcpy(grown_parts[0..self.parts.len], self.parts);
        const grown_overrides = try ar_alloc.alloc(?Override, self.overrides.len + 1);
        @memcpy(grown_overrides[0..self.overrides.len], self.overrides);

        // ─── Commit point: every step from here on is infallible ──
        grown_entries[grown_entries.len - 1] = new_entry;
        grown_parts[grown_parts.len - 1] = new_part;
        grown_overrides[grown_overrides.len - 1] = new_override;
        self.entries = grown_entries;
        self.parts = grown_parts;
        self.overrides = grown_overrides;
        // Apply the staged content-types update.
        self.overrides[ct_idx] = ct_new_override;
        self.parts[ct_idx].bytes = ct_new_part_bytes;
        self.parts[ct_idx].compression_method = ct_new_override.compressed.compression_method;
    }

    const StagedContentTypeUpdate = struct {
        idx: usize,
        new_part_bytes: []u8,
        new_override: Override,
    };

    /// Build the new [Content_Types].xml with the requested
    /// `<Override>` element appended. Does not mutate the store —
    /// caller commits via `self.overrides[idx] = new_override; ...`
    /// only after every other addPart allocation has succeeded.
    fn stageContentTypeOverride(
        self: *PartStore,
        part_name: []const u8,
        content_type: []const u8,
        poller: Poller,
    ) !StagedContentTypeUpdate {
        const ct_idx = self.findIndex("[Content_Types].xml") orelse
            return error.MissingContentTypes;
        const ct_part = self.parts[ct_idx];
        const old_xml = ct_part.bytes;
        const close_tag = "</Types>";
        const close_pos = std.mem.lastIndexOf(u8, old_xml, close_tag) orelse
            return error.MalformedContentTypes;

        const ar_alloc = self.arena.allocator();
        var buf: std.ArrayListUnmanaged(u8) = .empty;
        // No defer-deinit: on success the bytes are kept (referenced
        // from the staged result); on error the arena cleans up.
        try buf.appendSlice(ar_alloc, old_xml[0..close_pos]);
        try buf.appendSlice(ar_alloc, "<Override ContentType=\"");
        try appendXmlEscaped(ar_alloc, &buf, content_type);
        try buf.appendSlice(ar_alloc, "\" PartName=\"/");
        try appendXmlEscaped(ar_alloc, &buf, part_name);
        try buf.appendSlice(ar_alloc, "\"/>");
        try buf.appendSlice(ar_alloc, old_xml[close_pos..]);
        const new_xml = try buf.toOwnedSlice(ar_alloc);

        // Bound the uncompressed-size field too. The compressed
        // check below would catch anything that didn't shrink, but
        // a payload that compresses smaller while its uncompressed
        // size lands at exactly 0xFFFFFFFF would still write the
        // Zip64 sentinel into the LFH/CDFH uncompressed_size field.
        if (new_xml.len >= std.math.maxInt(u32)) return Error.Zip64NotSupported;

        // Compress the new CT XML — same policy as replacePart.
        var ct_compressed: std.ArrayListUnmanaged(u8) = .empty;
        defer ct_compressed.deinit(ar_alloc);
        const ct_method = try compressInto(ar_alloc, new_xml, &ct_compressed, poller);
        if (ct_compressed.items.len >= std.math.maxInt(u32)) return Error.Zip64NotSupported;
        const ct_payload = try ct_compressed.toOwnedSlice(ar_alloc);

        return .{
            .idx = ct_idx,
            .new_part_bytes = new_xml,
            .new_override = .{ .compressed = .{
                .payload = ct_payload,
                .compression_method = ct_method,
                .crc32 = std.hash.Crc32.hash(new_xml),
                .uncompressed_size = @intCast(new_xml.len),
            } },
        };
    }

    /// Append `s` to `buf` with XML attribute-value escaping. Covers
    /// the five XML-significant characters; sufficient for emitting
    /// caller-provided part names + content types into the
    /// `[Content_Types].xml` override entries.
    ///
    /// B3 iter-wr-6 NOTE: kept local rather than forwarded to
    /// `pkg/sheet_plan.zig::appendXmlEscaped`. The plan-side variant
    /// returns the broader `sheet_plan.Error` set (includes
    /// `RowOutOfRange`, `InvalidMergeRange`, …) which `PartStore`'s
    /// `Error` set deliberately omits. Forwarding would widen
    /// `addPart`'s declared error set beyond what its callers
    /// (Workbook.* paths) accept. The 5-entity escape stays inline;
    /// the canonical home for the rejecting variant remains
    /// `pkg/sheet_plan.zig`.
    fn appendXmlEscaped(
        alloc: std.mem.Allocator,
        buf: *std.ArrayListUnmanaged(u8),
        s: []const u8,
    ) !void {
        for (s) |c| switch (c) {
            '&' => try buf.appendSlice(alloc, "&amp;"),
            '<' => try buf.appendSlice(alloc, "&lt;"),
            '>' => try buf.appendSlice(alloc, "&gt;"),
            '"' => try buf.appendSlice(alloc, "&quot;"),
            '\'' => try buf.appendSlice(alloc, "&apos;"),
            else => try buf.append(alloc, c),
        };
    }

    /// Drop a part from the archive, along with its
    /// `[Content_Types].xml` override and any `<Relationship>` in any
    /// `.rels` part that targets it.
    ///
    /// Idempotent: removing an absent part is a no-op, not an error.
    /// That is the friendlier contract for scrub-style callers which
    /// ask for removal without first checking presence.
    ///
    /// Ordering matters. The content-type and relationship edits run
    /// FIRST, while the parallel arrays still hold their original
    /// indices, and the compaction happens last — doing it the other
    /// way round would have those edits addressing shifted slots.
    ///
    /// Not for structural parts. Removing `xl/workbook.xml` or a live
    /// worksheet would produce an archive that no reader accepts; this
    /// exists for genuinely optional parts such as `docProps/custom.xml`.
    pub fn removePart(self: *PartStore, name: []const u8) !void {
        const idx = self.findIndex(name) orelse return;
        const ar_alloc = self.arena.allocator();

        // 1. Content-type override, if the part had one.
        if (self.findIndex("[Content_Types].xml")) |ct_idx| {
            const ct = try self.part("[Content_Types].xml") orelse return Error.MissingContentTypes;
            const stripped = try removeContentTypeOverride(ar_alloc, ct.bytes, name);
            defer ar_alloc.free(stripped);
            if (!std.mem.eql(u8, stripped, ct.bytes)) {
                // Guard against removing the CT part itself, which
                // would make the archive unreadable.
                std.debug.assert(ct_idx != idx);
                try self.replacePart("[Content_Types].xml", stripped);
            }
        }

        // 2. Any relationship pointing at it. Targets are written
        //    relative to the rels file's owner, so both the bare name
        //    and a leading-slash absolute form are matched.
        for (self.parts, 0..) |p, i| {
            if (i == idx) continue;
            if (!std.mem.endsWith(u8, p.name, ".rels")) continue;
            const rels_part = try self.part(p.name) orelse continue;
            const stripped = try removeRelationshipsTo(ar_alloc, rels_part.bytes, name);
            defer ar_alloc.free(stripped);
            if (!std.mem.eql(u8, stripped, rels_part.bytes)) {
                try self.replacePart(p.name, stripped);
            }
        }

        // 3. Compact the three parallel arrays. alloc+copy rather than
        //    in-place shifting for the same reason addPart grows that
        //    way: a failed allocation must leave the store untouched.
        const n = self.parts.len;
        std.debug.assert(n > 0);
        const new_entries = try ar_alloc.alloc(ZipEntry, n - 1);
        const new_parts = try ar_alloc.alloc(Part, n - 1);
        const new_overrides = try ar_alloc.alloc(?Override, n - 1);

        var w: usize = 0;
        for (0..n) |r| {
            if (r == idx) continue;
            new_entries[w] = self.entries[r];
            new_parts[w] = self.parts[r];
            new_overrides[w] = self.overrides[r];
            w += 1;
        }
        std.debug.assert(w == n - 1);

        self.entries = new_entries;
        self.parts = new_parts;
        self.overrides = new_overrides;
    }

    /// Replace the bytes of an existing part. The compressed payload is
    /// staged as an override and written at `save()`; `parts[idx].bytes`
    /// is updated in the same call, so **`part(name)` returns the new
    /// bytes immediately**. `error.PartNotFound` if `name` is absent.
    ///
    /// (This doc block spent several iterations attached to `addPart`
    /// saying the opposite — that overrides were write-only until save.
    /// The mirror below has been here since iter-er-4, and M5b2's
    /// transaction depends on it: the candidate parses its new typed
    /// views out of the store it just staged into.)
    ///
    /// `bytes` is duped into store-owned storage; the caller may free
    /// its own buffer as soon as the call returns.
    pub fn replacePart(self: *PartStore, name: []const u8, bytes: []const u8) !void {
        return self.replacePartControlled(name, bytes, .none);
    }

    /// Payloads at or above this get an exact block instead of the
    /// arena. The threshold is about the arena's chunk arithmetic, not
    /// about the part: below it a payload generally fits the chunk the
    /// store already has and costs nothing beyond its own bytes, above
    /// it the request is itself what sizes the next chunk.
    pub const big_payload_bytes = 1 << 20;

    /// Bytes this store keeps resident: the arena's capacity **plus**
    /// every out-of-arena block `dupePayload` carved for a large
    /// payload.
    ///
    /// Retention accounting has to see both. Before M10m every replaced
    /// payload was duped into the arena, so `arena.queryCapacity()` was
    /// the whole figure; M10m moved payloads ≥ `big_payload_bytes` out
    /// of it to stop the chunk ladder doubling around them, and a
    /// generation holding hundreds of MiB of replaced parts would
    /// otherwise report only the few KiB its arena still spans —
    /// letting `max_retained_bytes` be overshot by however much the
    /// blocks weigh. Saturating, because this bounds a ceiling check
    /// and a wrap would read as "plenty of room".
    pub fn residentBytes(self: *const PartStore) u64 {
        var total: u64 = self.arena.queryCapacity();
        for (self.big_parts.items) |b| total +|= b.len;
        return total;
    }

    fn dupePayload(
        self: *PartStore,
        ar_alloc: std.mem.Allocator,
        bytes: []const u8,
    ) ![]u8 {
        if (bytes.len < big_payload_bytes) return ar_alloc.dupe(u8, bytes);
        // Capacity first, so the block's existence is the last fallible
        // step: an append that failed after the dupe would have to undo
        // it, and the callers state their atomicity as "every
        // allocation succeeded before anything is installed".
        try self.big_parts.ensureUnusedCapacity(self.allocator, 1);
        const block = try self.allocator.dupe(u8, bytes);
        self.big_parts.appendAssumeCapacity(block);
        return block;
    }

    /// `replacePart` with §5.5's poll seam (M5d1).
    ///
    /// Compression is DEFERRED (M10b): this call stages the raw bytes
    /// and a `.pending` override, and the deflate runs when a save
    /// path first needs the compressed form — under that save's own
    /// poller, which keeps M5d1's seam where the work actually is. A
    /// recalc that is never saved never compresses at all, and every
    /// in-memory reader was already served by the `parts[idx].bytes`
    /// mirror rather than the payload. `error.Cancelled` leaves the
    /// store byte-identical — the poll precedes any field write.
    pub fn replacePartControlled(
        self: *PartStore,
        name: []const u8,
        bytes: []const u8,
        poller: Poller,
    ) !void {
        const idx = self.findIndex(name) orelse return error.PartNotFound;
        const ar_alloc = self.arena.allocator();

        // Strict `>=`: 0xFFFFFFFF is the Zip64 sentinel — emitting that
        // in compressed_size or uncompressed_size produces an archive
        // the reader treats as Zip64 and rejects.
        if (bytes.len >= std.math.maxInt(u32)) return Error.Zip64NotSupported;

        // The seam's poll survives the deferral: a replace on an
        // already-cancelled run still refuses before mutating.
        try poller.check();

        // Build the new arena-owned value BEFORE installing anything
        // so a mid-allocation OOM leaves the store unchanged (no
        // partial-mutation observable to a caller that recovers from
        // the error).
        const dupe_bytes = try self.dupePayload(ar_alloc, bytes);
        self.overrides[idx] = .pending;
        // Mirror into parts[idx].bytes so subsequent part() lookups see
        // the updated content. NOTE: derived content_type entries
        // inferred from a replaced [Content_Types].xml are NOT
        // refreshed until the next open() — same v1 contract as before.
        // `compression_method` is the materialization's to fill: the
        // method does not exist until the compressor picks it, which
        // M10b defers to the save that needs it. It therefore still
        // names the DISPLACED bytes' method until then — documented on
        // the field, because a caller reading the public `Part` can see
        // that pair before any save makes it coherent.
        self.parts[idx].bytes = dupe_bytes;

        // Rels-cache refresh: when the replaced part is itself a
        // `_rels/<base>.rels` file, re-parse its relationships so
        // `store.rels(owner)` returns the post-replace view.
        // Without this, `Workbook.addSheet` and similar callers
        // could patch a rels file but downstream
        // `Worksheet.resolvePartName` (which queries `store.rels`)
        // wouldn't see the new entry until the workbook was
        // re-opened. iter-er-4.
        if (try relsOwner(ar_alloc, name)) |owner| {
            const refreshed = try parseRelationships(ar_alloc, dupe_bytes);
            // `put` overwrites any existing entry for this owner.
            try self.rels_by_owner.put(ar_alloc, owner, refreshed);
        }
    }

    /// Atomic write of the package to `path`. Untouched parts are
    /// emitted verbatim (LFH + payload bytes copied from the backing,
    /// CDFH copied with patched lfh_offset). Overridden parts get
    /// fresh LFH + payload but reuse the source CDFH (with patched
    /// fields). EOCD comment is preserved.
    pub fn save(self: *PartStore, io: std.Io, path: []const u8) !void {
        // The post-commit durability warning is dropped here and only
        // here: a bare `save` has no report to carry it in, and §5.7.9 is
        // explicit that it is not an error. `saveControlled` is what
        // `saveWithRecalc` (M5d2) calls to get it.
        _ = try self.saveControlled(io, path, .none);
    }

    /// `save` with §5.5's poll seam, returning §5.7.9's commit outcome
    /// (M5d1).
    ///
    /// Polls per entry while the archive streams out and once more
    /// immediately before `AtomicFile.finish` — that last one is the
    /// final poll §5.7.9 places before the commit point, so a cancelled
    /// save can never have renamed anything. Everything from `finish`
    /// onward is the non-cancellable commit region.
    pub fn saveControlled(self: *PartStore, io: std.Io, path: []const u8, poller: Poller) !Commit {
        return self.saveCommitted(io, path, poller, .none);
    }

    /// `saveControlled` with §5.7.9's swap point exposed (M5d2).
    ///
    /// `hook` runs after the rename has committed and before the
    /// directory fsync — the position §5.7.9 makes normative for
    /// `saveWithRecalc`'s in-memory swap. It cannot fail, cannot be
    /// cancelled, and cannot be skipped: once `finish` returns, the
    /// destination already names the new bytes, so a memory state that
    /// still describes the old ones is the inconsistency the ordering
    /// exists to prevent.
    pub fn saveCommitted(
        self: *PartStore,
        io: std.Io,
        path: []const u8,
        poller: Poller,
        hook: CommitHook,
    ) !Commit {
        try self.materializeOverrides(poller);
        try self.checkArchiveBounds();

        // Ahead of creating the temp file: a save that is already
        // cancelled should not leave the destination's directory holding
        // even a zero-length `.ztmp-N` for the instant before `deinit`
        // removes it.
        try poller.check();

        var write_buf: [4096]u8 = undefined;
        var atomic_file = try AtomicFile.init(io, path, &write_buf);
        defer atomic_file.deinit();

        try self.emitArchive(&atomic_file.file_writer.interface, poller);

        // §5.7.9's final poll. Everything after it — the `File.sync`, the
        // rename, the directory fsync — is the non-cancellable commit
        // region, so this is the last instant at which a cancelled save
        // is still a save that changed nothing.
        try poller.check();
        try atomic_file.finish();
        // The commit point has passed. §5.7.9's swap goes here — between
        // the rename and the directory fsync — so that no observer can
        // find the file published and the memory stale.
        hook.fire();
        return atomic_file.syncDir();
    }

    /// `saveCommitted`'s archive, emitted into caller-owned memory
    /// instead of a temp file (M9a2, §5.10's producer for a
    /// store-backed workbook). No file is opened and no commit point
    /// exists, so §5.7.9's vocabulary does not apply: a cancelled emit
    /// frees the partial buffer and leaves both the store and the
    /// caller's view of it untouched. The returned bytes are the
    /// caller's, freed with `allocator`.
    pub fn saveToOwnedBuffer(self: *PartStore, allocator: std.mem.Allocator, poller: Poller) ![]u8 {
        try self.materializeOverrides(poller);
        try self.checkArchiveBounds();
        try poller.check();
        var sink: std.Io.Writer.Allocating = .init(allocator);
        defer sink.deinit();
        try self.emitArchive(&sink.writer, poller);
        return sink.toOwnedSlice();
    }

    /// The deferred half of `replacePartControlled` (M10b): compress
    /// every `.pending` override, under the save's poller — deflate on
    /// a large sheet is exactly the stretch M5d1's seam exists for.
    /// Runs before `checkArchiveBounds` on every save path, so the
    /// bounds are always checked against real compressed sizes and
    /// the emit loops below never meet a pending slot. The compressor,
    /// its stored-vs-deflate policy and its input are the ones the
    /// eager path used, which is why the emitted bytes are identical.
    fn materializeOverrides(self: *PartStore, poller: Poller) !void {
        const ar_alloc = self.arena.allocator();
        for (self.overrides, 0..) |*slot, i| {
            const ov = slot.* orelse continue;
            if (ov != .pending) continue;
            const bytes = self.parts[i].bytes;

            var compressed: std.ArrayListUnmanaged(u8) = .empty;
            defer compressed.deinit(ar_alloc);
            const method = try compressInto(ar_alloc, bytes, &compressed, poller);
            if (compressed.items.len >= std.math.maxInt(u32)) return Error.Zip64NotSupported;

            slot.* = .{ .compressed = .{
                .payload = try compressed.toOwnedSlice(ar_alloc),
                .compression_method = method,
                .crc32 = std.hash.Crc32.hash(bytes),
                .uncompressed_size = @intCast(bytes.len),
            } };
            self.parts[i].compression_method = method;
        }
    }

    /// Preflight ZIP32 limits BEFORE any byte is produced. Every
    /// offset / size field on the wire is u32 (offsets, CD size,
    /// payload sizes) or u16 (name length, comment length, entry
    /// count). Compute the projected total saved size and reject
    /// upfront so the file emitter never leaves a partial archive in
    /// the atomic-file's tmp slot and the buffer emitter never grows
    /// an allocation it would have to throw away.
    fn checkArchiveBounds(self: *const PartStore) !void {
        if (self.entries.len > std.math.maxInt(u16)) return Error.ZipArchiveTooLarge;
        if (self.eocd_comment.len > std.math.maxInt(u16)) return Error.ZipArchiveTooLarge;
        // Compute the serialized ZIP32 fields up front (cd_offset and
        // cd_size) and reject if either would write the Zip64
        // sentinel. cd_offset = total LFH phase bytes; cd_size =
        // total CDFH phase bytes. Total file size is NOT a
        // serialized field, so we don't reject on it — only on the
        // values that actually go on wire.
        var lfh_phase_total: u64 = 0;
        var cdfh_phase_total: u64 = 0;
        for (self.entries, 0..) |e, i| {
            if (e.name.len > std.math.maxInt(u16)) return Error.ZipArchiveTooLarge;
            const lfh_total: u64 = if (self.overrides[i]) |ov| blk: {
                // Materialization precedes bounds on every save path;
                // a pending slot here is a caller that skipped it.
                std.debug.assert(ov != .pending);
                const c = ov.compressed;
                if (c.payload.len >= std.math.maxInt(u32)) return Error.ZipArchiveTooLarge;
                break :blk 30 + @as(u64, e.name.len) + @as(u64, c.payload.len);
            } else blk: {
                break :blk @as(u64, e.lfh_total_len) +
                    @as(u64, e.compressed_size) +
                    @as(u64, e.data_descriptor_len);
            };
            const cdfh_total: u64 = if (self.overrides[i] != null)
                46 + @as(u64, e.name.len)
            else
                @as(u64, e.cdfh_total_len);
            lfh_phase_total += lfh_total;
            cdfh_phase_total += cdfh_total;
        }
        if (lfh_phase_total >= std.math.maxInt(u32)) return Error.ZipArchiveTooLarge;
        if (cdfh_phase_total >= std.math.maxInt(u32)) return Error.ZipArchiveTooLarge;
        // Round-trip guard: zlsx's own reader rejects any file
        // whose stat.size exceeds u32 (we don't support Zip64).
        // Producing an archive larger than that would be unreadable
        // by ourselves, so cap total length too.
        const total_projected = lfh_phase_total + cdfh_phase_total +
            @as(u64, eocd_min_size) + @as(u64, self.eocd_comment.len);
        if (total_projected > std.math.maxInt(u32)) return Error.ZipArchiveTooLarge;
    }

    /// The archive byte stream itself, shared by `saveCommitted` (into
    /// the atomic temp file) and `saveToOwnedBuffer` (into memory).
    /// Callers run `checkArchiveBounds` first; this ends with a flush so
    /// both sinks hold the complete archive when it returns.
    fn emitArchive(self: *PartStore, w: *std.Io.Writer, poller: Poller) !void {
        var written: u64 = 0;
        const new_lfh_offsets = try self.allocator.alloc(u32, self.entries.len);
        defer self.allocator.free(new_lfh_offsets);

        // One reusable 64 KiB window for the byte-preserving copies
        // below. It replaces a per-entry `page_allocator.alloc(total)`,
        // which meant a 200 MiB untouched part was read fully into RSS
        // before a single byte reached the temp file — and was, being one
        // read and one write, the longest unpollable stretch in a save.
        const copy_buf = try self.allocator.alloc(u8, control.chunk_bytes);
        defer self.allocator.free(copy_buf);

        for (self.entries, 0..) |e, i| {
            try poller.check();
            new_lfh_offsets[i] = @intCast(written);
            if (self.overrides[i]) |ov| {
                // Build a fresh LFH with the override's compression
                // method + size. Reuses the original LFH name + extra
                // bytes verbatim so we keep any extension fields the
                // source carried (timestamps etc.).
                std.debug.assert(ov != .pending);
                const c = ov.compressed;
                var lfh_bytes: [30]u8 = undefined;
                std.mem.writeInt(u32, lfh_bytes[0..4], lfh_signature, .little);
                std.mem.writeInt(u16, lfh_bytes[4..6], 20, .little); // version
                std.mem.writeInt(u16, lfh_bytes[6..8], 0, .little); // flags
                std.mem.writeInt(u16, lfh_bytes[8..10], c.compression_method, .little);
                std.mem.writeInt(u16, lfh_bytes[10..12], 0, .little); // mod time
                std.mem.writeInt(u16, lfh_bytes[12..14], 0x21, .little); // mod date (1980-01-01)
                std.mem.writeInt(u32, lfh_bytes[14..18], c.crc32, .little);
                std.mem.writeInt(u32, lfh_bytes[18..22], @intCast(c.payload.len), .little);
                std.mem.writeInt(u32, lfh_bytes[22..26], c.uncompressed_size, .little);
                std.mem.writeInt(u16, lfh_bytes[26..28], @intCast(e.name.len), .little);
                std.mem.writeInt(u16, lfh_bytes[28..30], 0, .little); // no extra
                try w.writeAll(&lfh_bytes);
                try w.writeAll(e.name);
                var it = poller.chunks(c.payload);
                while (try it.next()) |chunk| try w.writeAll(chunk);
                written += @as(u64, lfh_bytes.len) + @as(u64, e.name.len) + @as(u64, c.payload.len);
            } else {
                // Untouched: stream LFH + payload from the source
                // file byte-for-byte. For entries with a data
                // descriptor (flag 0x0008), ALSO copy the trailing
                // 12/16-byte descriptor — the CDFH still advertises
                // that flag, so a reader will expect those bytes
                // after the payload.
                const total = e.lfh_total_len + e.compressed_size + e.data_descriptor_len;
                // Source-byte branch: only reachable when `overrides[i]
                // == null`, which implies the store came from `open()`
                // or `openBuffer()`. Fresh stores override every entry,
                // so this branch never fires for them — and if it did,
                // their empty backing would refuse the read.
                var copied: usize = 0;
                while (copied < total) {
                    try poller.check();
                    const n = @min(total - copied, copy_buf.len);
                    try self.backing.readAt(e.lfh_offset + copied, copy_buf[0..n]);
                    try w.writeAll(copy_buf[0..n]);
                    copied += n;
                }
                written += @as(u64, total);
            }
        }

        // Sentinel-safety: 0xFFFFFFFF in the EOCD's cd_offset field
        // means "look for Zip64 extras", which we don't emit.
        if (written >= std.math.maxInt(u32)) return Error.ZipArchiveTooLarge;
        const new_cd_offset: u32 = @intCast(written);
        for (self.entries, 0..) |e, i| {
            try poller.check();
            if (self.overrides[i]) |ov| {
                // Fresh CDFH for the override. 46-byte header + name.
                std.debug.assert(ov != .pending);
                const c = ov.compressed;
                var cdfh_bytes: [46]u8 = undefined;
                std.mem.writeInt(u32, cdfh_bytes[0..4], cdfh_signature, .little);
                std.mem.writeInt(u16, cdfh_bytes[4..6], 20, .little); // version made by
                std.mem.writeInt(u16, cdfh_bytes[6..8], 20, .little); // version needed
                std.mem.writeInt(u16, cdfh_bytes[8..10], 0, .little); // flags
                std.mem.writeInt(u16, cdfh_bytes[10..12], c.compression_method, .little);
                std.mem.writeInt(u16, cdfh_bytes[12..14], 0, .little);
                std.mem.writeInt(u16, cdfh_bytes[14..16], 0x21, .little);
                std.mem.writeInt(u32, cdfh_bytes[16..20], c.crc32, .little);
                std.mem.writeInt(u32, cdfh_bytes[20..24], @intCast(c.payload.len), .little);
                std.mem.writeInt(u32, cdfh_bytes[24..28], c.uncompressed_size, .little);
                std.mem.writeInt(u16, cdfh_bytes[28..30], @intCast(e.name.len), .little);
                std.mem.writeInt(u16, cdfh_bytes[30..32], 0, .little); // no extra
                std.mem.writeInt(u16, cdfh_bytes[32..34], 0, .little); // no comment
                std.mem.writeInt(u16, cdfh_bytes[34..36], 0, .little); // disk number
                std.mem.writeInt(u16, cdfh_bytes[36..38], 0, .little); // internal attrs
                std.mem.writeInt(u32, cdfh_bytes[38..42], 0, .little); // external attrs
                std.mem.writeInt(u32, cdfh_bytes[42..46], new_lfh_offsets[i], .little);
                try w.writeAll(&cdfh_bytes);
                try w.writeAll(e.name);
                written += @as(u64, cdfh_bytes.len) + @as(u64, e.name.len);
            } else {
                // Untouched: read CDFH bytes from the source file,
                // patch the lfh_offset field at byte 42-46.
                const cdfh = try self.allocator.alloc(u8, e.cdfh_total_len);
                defer self.allocator.free(cdfh);
                // Same invariant as the LFH branch above: untouched
                // entries imply we came from `open()` / `openBuffer()`
                // with real source bytes.
                try self.backing.readAt(e.cdfh_offset, cdfh);
                std.mem.writeInt(u32, cdfh[42..46], new_lfh_offsets[i], .little);
                try w.writeAll(cdfh);
                written += @as(u64, cdfh.len);
            }
        }
        // Sentinel-safety: same Zip64 reservation applies to cd_size.
        const cd_size_u64 = written - new_cd_offset;
        if (cd_size_u64 >= std.math.maxInt(u32)) return Error.ZipArchiveTooLarge;
        const new_cd_size: u32 = @intCast(cd_size_u64);

        // EOCD.
        var eocd_bytes: [eocd_min_size]u8 = undefined;
        std.mem.writeInt(u32, eocd_bytes[0..4], eocd_signature, .little);
        std.mem.writeInt(u16, eocd_bytes[4..6], 0, .little);
        std.mem.writeInt(u16, eocd_bytes[6..8], 0, .little);
        std.mem.writeInt(u16, eocd_bytes[8..10], @intCast(self.entries.len), .little);
        std.mem.writeInt(u16, eocd_bytes[10..12], @intCast(self.entries.len), .little);
        std.mem.writeInt(u32, eocd_bytes[12..16], new_cd_size, .little);
        std.mem.writeInt(u32, eocd_bytes[16..20], new_cd_offset, .little);
        std.mem.writeInt(u16, eocd_bytes[20..22], @intCast(self.eocd_comment.len), .little);
        try w.writeAll(&eocd_bytes);
        try w.writeAll(self.eocd_comment);

        try w.flush();
    }

    fn findIndex(self: *const PartStore, name: []const u8) ?usize {
        for (self.parts, 0..) |p, i| {
            if (std.mem.eql(u8, p.name, name)) return i;
        }
        return null;
    }

    pub fn partNames(self: *const PartStore) ![]const []const u8 {
        // Build a names-only view on demand. Cheap because parts.len
        // is small; alternative is to cache a separate slice up-front.
        // For M1 keep the data structures minimal. Returns an error
        // on arena alloc failure so callers can surface OOM rather
        // than silently see an empty list.
        const ar_alloc = @constCast(&self.arena).allocator();
        const out = try ar_alloc.alloc([]const u8, self.parts.len);
        for (self.parts, 0..) |p, i| out[i] = p.name;
        return out;
    }

    pub fn part(self: *const PartStore, name: []const u8) Error!?Part {
        return self.partControlled(name, .none);
    }

    /// `part` with §5.5's poll seam (M5d1). First access to a part is
    /// what decompresses it, so "model materialization" is exactly this
    /// call — and on a workbook whose SST is 120 MiB it is a long
    /// operation with no other seam in it.
    pub fn partControlled(self: *const PartStore, name: []const u8, poller: Poller) Error!?Part {
        const idx = self.findIndex(name) orelse return null;
        try materializeAt(self, idx, poller);
        return self.parts[idx];
    }

    /// Materialize the part at `idx` if it hasn't been already. The
    /// `@constCast` here is safe because the cache fill is morally
    /// mutable — every reader that asks for the same bytes gets the
    /// same answer; we just compute it once. Inherently-empty parts
    /// (uncompressed_size == 0) skip the decompress entirely.
    ///
    /// Reads the compressed payload through `self.backing` into a
    /// scratch buffer (page-allocator), decompresses into the arena,
    /// frees the scratch. Decompressed bytes are cached on
    /// `Part.bytes` for the rest of THIS generation's lifetime — the
    /// cache is per-generation, the bytes it is filled from are not.
    fn materializeAt(self: *const PartStore, idx: usize, poller: Poller) Error!void {
        const p = &@constCast(self).parts[idx];
        // An override IS this part's content, the empty one included.
        // `bytes.len == 0` alone cannot mean "not materialized yet":
        // a caller who replaced a part with `""` leaves exactly that
        // shape behind, and reloading the source over it would hand
        // back the original. Since M10b defers the deflate to save,
        // that reloaded original is also what would be WRITTEN — a
        // blanked part silently reverting. The override slot is the
        // authoritative signal; the byte length is not.
        if (self.overrides[idx] != null) return;
        if (p.bytes.len > 0 or p.uncompressed_size == 0) return;
        const ar_alloc = @constCast(&self.arena).allocator();

        const compressed = try std.heap.page_allocator.alloc(u8, p.compressed_size);
        defer std.heap.page_allocator.free(compressed);
        // Fresh stores never have non-materialized parts (every part
        // is either inherently empty or carried in an `overrides[i]`
        // slot with bytes already in the arena). If we land here on
        // one, its empty backing refuses the read — the store
        // invariant is violated and `BadZip` says so.
        try readChunked(self.backing, p.payload_offset, compressed, poller);

        // §9.1c M10t, candidate 1 — the READ path gets M10m's exact
        // block, which until now only the *replace* path had. The arena
        // sizes a chunk at 1.5 × (what it holds + what was asked for),
        // so a 5 288 922 B sheet body bought a 7 951 114 B chunk: `id7`,
        // the only site live in all 24 eras. `decompressPayload`
        // allocates `declared_uncompressed` exactly, so handing it the
        // raw allocator gives the payload its own block and no ladder.
        //
        // Same lifetime the arena gave it — `deinit` frees `big_parts` —
        // so `Part.bytes` still means "valid until this generation
        // dies", which is the contract goal_codex.md §4 was retracted
        // for breaking. Nothing is released early here; the block is
        // only shaped differently.
        const exact = p.uncompressed_size >= big_payload_bytes;
        const target = if (exact) self.allocator else ar_alloc;
        if (exact) {
            // Capacity first, so the block's existence is the last
            // fallible step — `dupePayload`'s discipline, for the same
            // reason: an append that failed after the decompress would
            // have to undo it.
            try @constCast(self).big_parts.ensureUnusedCapacity(self.allocator, 1);
        }
        const bytes = try decompressPayload(
            target,
            compressed,
            p.compression_method,
            p.uncompressed_size,
            poller,
        );
        errdefer if (exact) self.allocator.free(bytes);
        if (std.hash.Crc32.hash(bytes) != p.crc32) return Error.BadZip;
        if (exact) @constCast(self).big_parts.appendAssumeCapacity(bytes);
        p.bytes = bytes;
    }

    /// Structural parts are decompressed eagerly at `open()` because
    /// content-type / relationship resolution needs their bytes
    /// inline. Everything else defers to first-access via `part()`.
    fn isStructuralPart(name: []const u8) bool {
        if (std.mem.eql(u8, name, "[Content_Types].xml")) return true;
        if (std.mem.endsWith(u8, name, ".rels")) return true;
        return false;
    }

    /// Filtered view of parts whose content type starts with
    /// `image/` (PNG, JPEG, GIF, etc.). Caller-friendly C2a
    /// MVP — anchor / per-sheet attribution lives in the future
    /// `drawings()` parser.
    ///
    /// Allocated inside the store's arena; valid until `deinit`.
    /// Returned as `[]const Part` because the slice is a read-only
    /// filtered view; callers should treat the contents as immutable
    /// borrows from the store.
    pub fn imageParts(self: *const PartStore) ![]const Part {
        const ar_alloc = @constCast(&self.arena).allocator();
        var out: std.ArrayListUnmanaged(Part) = .empty;
        for (self.parts, 0..) |p, idx| {
            const ct = p.content_type orelse continue;
            if (std.mem.startsWith(u8, ct, "image/")) {
                // Materialize bytes before exposing — image-walking
                // callers (drawing parser, zlsx-extract-images) need
                // the decompressed payload to write the image to disk
                // or compute its size. Lazy mode means this is a
                // no-op the second time around.
                try materializeAt(self, idx, .none);
                try out.append(ar_alloc, self.parts[idx]);
            }
        }
        return try out.toOwnedSlice(ar_alloc);
    }

    pub fn rels(self: *const PartStore, owner_part_name: []const u8) []const Relationship {
        return self.rels_by_owner.get(owner_part_name) orelse &.{};
    }

    /// Read-only predicate: does the store have at least one part
    /// override (from `replacePart` or `addPart`) that hasn't been
    /// flushed to disk yet?
    ///
    /// Note: `PartStore.save` does NOT clear overrides post-save —
    /// they persist across save calls. So this predicate reflects
    /// "diff vs the original on-disk archive opened by `open()`",
    /// not "uncommitted-since-last-save". Most callers want the
    /// former (e.g. for "do I need to save before exit?" — the
    /// answer should remain true even after a previous save).
    pub fn hasUnsavedChanges(self: *const PartStore) bool {
        for (self.overrides) |o| if (o != null) return true;
        return false;
    }

    /// Whether the part at `name` has been mutated since `open()`.
    /// Returns `false` for parts that don't exist (use `part()` to
    /// disambiguate). Used by `Editor.save` to know which source-
    /// covered parts need fresh substitution from PartStore bytes.
    pub fn isOverridden(self: *const PartStore, name: []const u8) bool {
        const idx = self.findIndex(name) orelse return false;
        return self.overrides[idx] != null;
    }

    /// Resolve a relationship `target` (which is interpreted relative
    /// to `owner_part_name`'s parent directory) into a normalised
    /// absolute part name. Returns `null` for external targets and
    /// for paths that escape the package root.
    ///
    /// External detection is heuristic on the target string itself
    /// because this method takes a raw target rather than a full
    /// Relationship. It catches URL schemes (`https://`, `mailto:`,
    /// `file://`), UNC paths (`\\server\share`), and Windows drive
    /// letters (`C:\foo`) — the shapes that external relationships
    /// actually take in real workbooks. Callers that already have a
    /// Relationship and want to be exact should branch on
    /// `Relationship.target_mode == .external` upstream.
    pub fn resolve(
        self: *const PartStore,
        owner_part_name: []const u8,
        target: []const u8,
    ) !?[]const u8 {
        if (target.len == 0) return null;
        if (looksExternal(target)) return null;
        // Absolute target (rare): "/xl/foo.xml" → "xl/foo.xml".
        if (target[0] == '/') {
            return try self.dupeArena(target[1..]);
        }
        // Relative target: collapse against owner's parent dir.
        const owner_dir = parentDir(owner_part_name);
        var stack: std.ArrayListUnmanaged([]const u8) = .empty;
        defer stack.deinit(self.allocator);
        // Seed stack with owner_dir segments.
        if (owner_dir.len > 0) {
            var it = std.mem.splitScalar(u8, owner_dir, '/');
            while (it.next()) |seg| {
                if (seg.len == 0) continue;
                try stack.append(self.allocator, seg);
            }
        }
        // Walk target.
        var it = std.mem.splitScalar(u8, target, '/');
        while (it.next()) |seg| {
            if (seg.len == 0 or std.mem.eql(u8, seg, ".")) continue;
            if (std.mem.eql(u8, seg, "..")) {
                if (stack.items.len == 0) return null; // escaped package
                _ = stack.pop();
                continue;
            }
            try stack.append(self.allocator, seg);
        }
        // Join with `/`. Total length = sum + (n-1) separators.
        var total: usize = 0;
        for (stack.items, 0..) |s, i| {
            total += s.len;
            if (i + 1 < stack.items.len) total += 1;
        }
        const ar_alloc = @constCast(&self.arena).allocator();
        const out = try ar_alloc.alloc(u8, total);
        var w: usize = 0;
        for (stack.items, 0..) |s, i| {
            @memcpy(out[w .. w + s.len], s);
            w += s.len;
            if (i + 1 < stack.items.len) {
                out[w] = '/';
                w += 1;
            }
        }
        return out;
    }

    fn dupeArena(self: *const PartStore, s: []const u8) ![]u8 {
        const ar_alloc = @constCast(&self.arena).allocator();
        return ar_alloc.dupe(u8, s);
    }
};

/// Detect targets that are external to the package and must NOT be
/// joined with a part-name parent directory. The shapes covered are:
///   - URL schemes:    `https://...`, `mailto:...`, `file:///...`
///   - UNC paths:      `\\server\share\...`
///   - Drive letters:  `C:\foo` or `C:/foo`
/// Anything else is treated as a relative or absolute package path
/// (the `/`-rooted branch in `resolve`).
fn looksExternal(target: []const u8) bool {
    if (target.len < 2) return false;
    if (target[0] == '\\' and target[1] == '\\') return true;
    if (target.len >= 3 and isAsciiAlpha(target[0]) and target[1] == ':' and
        (target[2] == '\\' or target[2] == '/'))
    {
        return true;
    }
    // URL scheme per RFC 3986 §3.1:
    //   scheme = ALPHA *( ALPHA / DIGIT / "+" / "-" / "." )
    // Walk the target until we hit a colon (URL scheme found) or any
    // non-scheme char (definitely not a URL — it's a path segment).
    // No fixed length cap — schemes are unbounded by spec, and the
    // first non-scheme char terminates the walk regardless of length.
    var i: usize = 0;
    while (i < target.len) : (i += 1) {
        const c = target[i];
        if (c == ':') {
            // First char must be a letter (RFC 3986 §3.1).
            return i > 0 and isAsciiAlpha(target[0]);
        }
        if (!(isAsciiAlpha(c) or std.ascii.isDigit(c) or c == '+' or c == '-' or c == '.')) {
            return false;
        }
    }
    return false;
}

fn isAsciiAlpha(c: u8) bool {
    return (c >= 'a' and c <= 'z') or (c >= 'A' and c <= 'Z');
}

// ─── ZIP scanner ──────────────────────────────────────────────────────

const ZipEntry = struct {
    name: []const u8,
    /// Offset (into src_buf) of the local file header (signature 0x04034b50).
    lfh_offset: usize,
    /// Total length of the LFH including the variable-length name +
    /// extra fields. = 30 + lfh_name_len + lfh_extra_len. Payload
    /// begins at `lfh_offset + lfh_total_len`.
    lfh_total_len: usize,
    /// Offset of the central-directory file header (signature 0x02014b50).
    cdfh_offset: usize,
    /// Total length of the CDFH including the variable-length tail
    /// (filename + extra + comment) — useful for byte-preserving
    /// copies that keep extra fields the source carried.
    cdfh_total_len: usize,
    /// Compressed payload size and start offset (= lfh_offset +
    /// lfh_total_len). Recorded in the CDFH; LFH copies have the
    /// same value EXCEPT when a data-descriptor (flag 0x0008) sat
    /// after the payload — CDFH is canonical either way.
    payload_offset: usize,
    compressed_size: u32,
    uncompressed_size: u32,
    compression_method: u16,
    crc32: u32,
    /// Length of the data descriptor that follows the payload, in
    /// bytes. 0 when the data-descriptor flag (0x0008) is clear; 12
    /// or 16 when set (16 if the optional 0x08074b50 signature is
    /// present, 12 otherwise). save() copies these bytes verbatim
    /// for untouched entries so the archive stays valid. Disambiguation
    /// uses the gap to the next entry's LFH, not a signature peek
    /// (which would mis-classify 1-in-2^32 inputs where CRC ==
    /// 0x08074b50).
    data_descriptor_len: usize,
    /// True iff the source CDFH had the 0x0008 general-purpose flag.
    /// Used during the second-pass `data_descriptor_len` computation.
    has_data_descriptor: bool,
};

/// Override = caller-supplied replacement bytes for an existing part,
/// with a fresh LFH + CDFH built at save() time.
const Override = union(enum) {
    /// Compressed and ready to emit.
    compressed: Compressed,
    /// Staged with compression deferred (M10b): the raw bytes are the
    /// mirror in `parts[i].bytes`, and every save path materializes
    /// through `materializeOverrides` before any bound is checked or
    /// byte emitted. A recalc that is never saved never pays the
    /// deflate — §9.1's profile put that at ~7 % of the whole
    /// evaluate lane — and the bytes a save does emit are identical,
    /// because the compressor, its policy and its input are.
    pending,

    const Compressed = struct {
        /// Compressed payload bytes (owned by arena).
        payload: []const u8,
        compression_method: u16,
        crc32: u32,
        uncompressed_size: u32,
    };
};

const eocd_signature: u32 = 0x06054b50;
const cdfh_signature: u32 = 0x02014b50;
const lfh_signature: u32 = 0x04034b50;
const eocd_min_size: usize = 22;
const eocd_scan_window: usize = 65535 + eocd_min_size;

fn scanCentralDirectory(arena: std.mem.Allocator, buf: []const u8) ![]ZipEntry {
    if (buf.len < eocd_min_size) return Error.NotPkzip;
    const eocd_off = try findEocd(buf);

    // Reject split / multi-disk archives. For single-disk ZIPs all
    // four disk-related fields are zero and the per-disk record
    // count equals the total. Otherwise CDFH/LFH offsets are
    // disk-relative and treating them as buffer offsets would
    // either error inscrutably or read unrelated bytes.
    const this_disk = std.mem.readInt(u16, buf[eocd_off + 4 ..][0..2], .little);
    const cd_disk = std.mem.readInt(u16, buf[eocd_off + 6 ..][0..2], .little);
    const records_on_disk = std.mem.readInt(u16, buf[eocd_off + 8 ..][0..2], .little);
    const total_records = std.mem.readInt(u16, buf[eocd_off + 10 ..][0..2], .little);
    if (this_disk != 0 or cd_disk != 0 or records_on_disk != total_records) {
        return Error.SplitArchiveNotSupported;
    }

    const cd_size = std.mem.readInt(u32, buf[eocd_off + 12 ..][0..4], .little);
    const cd_offset = std.mem.readInt(u32, buf[eocd_off + 16 ..][0..4], .little);

    if (cd_size == 0xFFFFFFFF or cd_offset == 0xFFFFFFFF) return Error.Zip64NotSupported;
    // Widen to usize BEFORE the saturating add. Both fields are u32,
    // so `cd_offset +| cd_size` would otherwise saturate at u32::max
    // (0xFFFFFFFF), and a genuine 4 GiB+1 sum on a 64-bit build with
    // a >4 GiB buf would compare equal to buf.len = 4 GiB and slip
    // through. Casting first lets the saturation operate at usize
    // width — correct on both 32-bit and 64-bit targets.
    const cd_off_us: usize = cd_offset;
    const cd_size_us: usize = cd_size;
    if (cd_off_us +| cd_size_us > buf.len) return Error.BadZip;

    var out: std.ArrayListUnmanaged(ZipEntry) = .empty;
    try out.ensureTotalCapacity(arena, total_records);

    var cur: usize = cd_off_us;
    // Safe: `cd_off_us +| cd_size_us > buf.len` was just rejected,
    // so the non-saturating sum fits usize.
    const cd_end = cd_off_us + cd_size_us;
    var idx: usize = 0;
    while (cur + 46 <= cd_end and idx < total_records) : (idx += 1) {
        const sig = std.mem.readInt(u32, buf[cur..][0..4], .little);
        if (sig != cdfh_signature) return Error.BadZip;

        const flags = std.mem.readInt(u16, buf[cur + 8 ..][0..2], .little);
        if (flags & 0x0001 != 0) return Error.EncryptedNotSupported;
        // Data-descriptor flag (0x0008) means the LFH's CRC + sizes
        // are zero and the real values trail the payload. The CDFH
        // copies we read have the real sizes either way, so this is
        // OK as long as we use CDFH-recorded sizes (we do).

        const compression_method = std.mem.readInt(u16, buf[cur + 10 ..][0..2], .little);
        const crc32 = std.mem.readInt(u32, buf[cur + 16 ..][0..4], .little);
        const compressed_size = std.mem.readInt(u32, buf[cur + 20 ..][0..4], .little);
        const uncompressed_size = std.mem.readInt(u32, buf[cur + 24 ..][0..4], .little);
        const filename_len = std.mem.readInt(u16, buf[cur + 28 ..][0..2], .little);
        const extra_len = std.mem.readInt(u16, buf[cur + 30 ..][0..2], .little);
        const comment_len = std.mem.readInt(u16, buf[cur + 32 ..][0..2], .little);
        const lfh_offset = std.mem.readInt(u32, buf[cur + 42 ..][0..4], .little);

        if (compressed_size == 0xFFFFFFFF or uncompressed_size == 0xFFFFFFFF or lfh_offset == 0xFFFFFFFF) {
            return Error.Zip64NotSupported;
        }
        if (compression_method != 0 and compression_method != 8) {
            return Error.UnsupportedCompression;
        }

        const name_start = cur + 46;
        if (name_start +| filename_len > cd_end) return Error.BadZip;
        const name = buf[name_start .. name_start + filename_len];

        // Compute payload offset by reading the LFH (its filename +
        // extra fields can differ from the CDFH copy).
        const lfh_off_us: usize = lfh_offset; // widen for the bounds math
        if (lfh_off_us +| 30 > buf.len) return Error.BadZip;
        const lfh_sig = std.mem.readInt(u32, buf[lfh_offset..][0..4], .little);
        if (lfh_sig != lfh_signature) return Error.BadZip;
        const lfh_name_len = std.mem.readInt(u16, buf[lfh_offset + 26 ..][0..2], .little);
        const lfh_extra_len = std.mem.readInt(u16, buf[lfh_offset + 28 ..][0..2], .little);
        // 30 + 2×u16 ≤ 30 + 131070 — fits both u32 and usize without
        // overflow concerns. lfh_total_len is bounded by definition.
        const lfh_total_len = 30 + @as(usize, lfh_name_len) + @as(usize, lfh_extra_len);
        const payload_offset = lfh_off_us +| lfh_total_len;
        if (payload_offset > buf.len) return Error.BadZip;
        const compressed_us: usize = compressed_size;
        if (payload_offset +| compressed_us > buf.len) return Error.BadZip;

        const cdfh_total = 46 + @as(usize, filename_len) + @as(usize, extra_len) + @as(usize, comment_len);
        // Reject when the CDFH's full tail (filename + extra +
        // comment) would overrun cd_end. Without this, a malformed
        // ZIP whose extra/comment lengths point past cd_end is
        // accepted, and save() later copies src_buf[cdfh_offset ..
        // cdfh_offset + cdfh_total_len] verbatim — pulling in EOCD
        // / comment bytes as if they were CDFH data.
        if (cur +| cdfh_total > cd_end) return Error.BadZip;

        try out.append(arena, .{
            .name = try arena.dupe(u8, name),
            .lfh_offset = lfh_offset,
            .lfh_total_len = lfh_total_len,
            .cdfh_offset = cur,
            .cdfh_total_len = cdfh_total,
            .payload_offset = payload_offset,
            .compressed_size = compressed_size,
            .uncompressed_size = uncompressed_size,
            .compression_method = compression_method,
            .crc32 = crc32,
            .data_descriptor_len = 0, // patched below
            .has_data_descriptor = (flags & 0x0008) != 0,
        });

        // Advance past this CDFH. The components are u16/usize sums
        // bounded by the per-entry limits already checked above; use
        // saturating so a pathological tail length doesn't wrap on
        // 32-bit. The next-iteration loop guard re-validates against
        // cd_end.
        cur = name_start +| filename_len +| extra_len +| comment_len;
    }

    if (idx != total_records) return Error.BadZip;

    // Second pass: compute data_descriptor_len for every DDF entry.
    // The descriptor sits immediately after the payload and is
    // either 12 or 16 bytes; disambiguate by measuring the gap to
    // the next entry's LFH (or to the central-directory start for
    // the last entry). This is exact — peeking at the first 4 bytes
    // would mis-classify the rare case where the CRC happens to
    // equal the descriptor signature 0x08074b50.
    const slice = out.items;
    // Sort entries by lfh_offset to compute gaps. Sort indices, not
    // the items themselves, because central-directory order is
    // semantic and must be preserved.
    const order = try arena.alloc(usize, slice.len);
    defer arena.free(order);
    for (order, 0..) |*o, i| o.* = i;
    const Ctx = struct {
        items: []ZipEntry,
        pub fn lessThan(self: *@This(), a: usize, b: usize) bool {
            return self.items[a].lfh_offset < self.items[b].lfh_offset;
        }
    };
    var ctx = Ctx{ .items = slice };
    std.sort.pdq(usize, order, &ctx, Ctx.lessThan);
    for (order, 0..) |idx2, k| {
        const e = &slice[idx2];
        if (!e.has_data_descriptor) continue;
        const payload_end = e.payload_offset + e.compressed_size;
        const next_off: usize = if (k + 1 < order.len) slice[order[k + 1]].lfh_offset else cd_offset;
        if (next_off < payload_end) return Error.BadZip;
        const gap = next_off - payload_end;
        if (gap != 12 and gap != 16) return Error.BadZip;
        e.data_descriptor_len = gap;
    }

    return out.toOwnedSlice(arena);
}

fn findEocd(buf: []const u8) !usize {
    // Scan back from the end of buf looking for the EOCD signature.
    // ZIP allows up to 65535 bytes of comment after EOCD, so the
    // scan window is bounded; longer would mean a malformed archive.
    // Local invariant: the signature is 4 bytes, so we need at least
    // 4 bytes to read it. Callers in this file enforce ≥ eocd_min_size
    // (22) transitively, but `findEocd` is also exported via the
    // save() round-trip path, so guard locally.
    if (buf.len < 4) return Error.NotPkzip;
    const max_back = @min(eocd_scan_window, buf.len);
    var i: usize = buf.len - 4;
    const lo = if (buf.len >= max_back) buf.len - max_back else 0;
    while (true) {
        const sig = std.mem.readInt(u32, buf[i..][0..4], .little);
        if (sig == eocd_signature) {
            const eocd_end = i + eocd_min_size;
            if (eocd_end > buf.len) return Error.BadZip;
            const comment_len = std.mem.readInt(u16, buf[i + 20 ..][0..2], .little);
            if (eocd_end + comment_len != buf.len) {
                // Comment must consume the file tail exactly. If not,
                // this is a stray signature inside another structure.
            } else {
                return i;
            }
        }
        if (i == lo) break;
        i -= 1;
    }
    return Error.NotPkzip;
}

/// The one compression policy every staging path shares, with §5.5's
/// poll seam through it (M5d1). Returns the ZIP compression method.
///
///   - Sub-1 KiB (and empty) inputs: STORED. Deflate's dynamic-block
///     header overhead dominates the gain on tiny XML.
///   - Larger inputs: deflate, falling back to STORED when compression
///     did not actually shrink the payload.
///
/// It was three verbatim copies of that policy before this row — in
/// `addPart`, `replacePart` and `stageContentTypeOverride` — and the
/// chunked seam had to reach all three, so they became one.
fn compressInto(
    alloc: std.mem.Allocator,
    bytes: []const u8,
    out: *std.ArrayListUnmanaged(u8),
    poller: Poller,
) !u16 {
    if (bytes.len < 1024 or bytes.len == 0) {
        try storeChunked(alloc, bytes, out, poller);
        return 0;
    }
    const zlsx = @import("zlsx");
    try zlsx.deflateCompressControlled(alloc, bytes, out, poller);
    if (out.items.len >= bytes.len) {
        out.clearRetainingCapacity();
        try storeChunked(alloc, bytes, out, poller);
        return 0;
    }
    return 8;
}

/// The STORED arm of `compressInto`. Chunked for the same reason the
/// deflate arm is: a part that fails the shrink test is exactly the part
/// that is large and incompressible (an embedded PNG, a signature blob),
/// so the fallback copy is not the cheap case.
fn storeChunked(
    alloc: std.mem.Allocator,
    bytes: []const u8,
    out: *std.ArrayListUnmanaged(u8),
    poller: Poller,
) !void {
    try out.ensureTotalCapacity(alloc, bytes.len);
    var it = poller.chunks(bytes);
    while (try it.next()) |chunk| out.appendSliceAssumeCapacity(chunk);
}

/// `SourceBacking.readAt` in `chunk_bytes` pieces. The backing read
/// itself is one `preadv`-shaped call per chunk, so this is a poll seam
/// rather than a buffering change.
fn readChunked(backing: *SourceBacking, offset: u64, dest: []u8, poller: Poller) Error!void {
    var done: usize = 0;
    while (done < dest.len) {
        try poller.check();
        const n = @min(dest.len - done, control.chunk_bytes);
        try backing.readAt(offset + done, dest[done .. done + n]);
        done += n;
    }
}

// ZIP-bomb defenses for `decompressPayload`. Both checks fire BEFORE
// the upfront `arena.alloc(declared_uncompressed)` so a crafted CDFH
// declaring multi-GB uncompressed size can't OOM the process.
//
//  - max_part_size: per-part hard cap. 512 MiB is ~4× larger than the
//    biggest legitimate part observed in our corpus (~120 MiB SST in
//    a 76 MiB workbook). xlsx files exceeding this almost certainly
//    aren't legitimate single-file workbooks.
//
//  - max_deflate_ratio: sanity cap on declared_uncompressed /
//    compressed. Real-world deflate hits ~1000:1 only on highly
//    redundant payloads (long runs of zeros). 4096:1 is a generous
//    margin that still blocks classic ZIP bombs (10MB compressed
//    declaring 4GB uncompressed = 400:1 — far below the cap, so
//    intentionally narrower bombs that fit within 4096:1 just have
//    to actually contain ~250KB of plausible compressed data per
//    1GB declared, at which point they're already being rate-limited
//    by the LFH/CDFH signature and decompression cost themselves).
const max_part_size: usize = 512 * 1024 * 1024;
const max_deflate_ratio: usize = 4096;

fn decompressPayload(
    arena: std.mem.Allocator,
    payload: []const u8,
    method: u16,
    declared_uncompressed: u32,
    poller: Poller,
) ![]u8 {
    if (declared_uncompressed > max_part_size) return Error.BadZip;
    // Saturating multiply guards against `payload.len * ratio`
    // overflow on pathological inputs. usize @ 64-bit can't reach
    // saturation in practice, but this stays correct on 32-bit too.
    const ratio_cap = std.math.mul(usize, payload.len, max_deflate_ratio) catch std.math.maxInt(usize);
    if (declared_uncompressed > ratio_cap) return Error.BadZip;

    if (method == 0) {
        if (payload.len != declared_uncompressed) return Error.BadZip;
        const out = try arena.alloc(u8, declared_uncompressed);
        var written: usize = 0;
        var it = poller.chunks(payload);
        while (try it.next()) |chunk| {
            @memcpy(out[written .. written + chunk.len], chunk);
            written += chunk.len;
        }
        return out;
    } else if (method == 8) {
        var src_reader = std.Io.Reader.fixed(payload);
        var flate_buffer: [std.compress.flate.max_window_len]u8 = undefined;
        var dec = std.compress.flate.Decompress.init(&src_reader, .raw, &flate_buffer);
        const out = try arena.alloc(u8, declared_uncompressed);
        var out_writer = std.Io.Writer.fixed(out);
        // §5.5: pull the inflated stream out in `chunk_bytes` pieces
        // rather than one `streamExact64` of the whole part. The
        // decompressor is a stream over one window, so the boundaries
        // change nothing about the bytes produced — `streamExact64`
        // resumes exactly where the previous call stopped — but they give
        // a 500 MiB sharedStrings.xml the same poll density a 4 KiB rels
        // file gets.
        var remaining: u64 = declared_uncompressed;
        while (remaining > 0) {
            try poller.check();
            const n = @min(remaining, control.chunk_bytes);
            dec.reader.streamExact64(&out_writer, n) catch return Error.BadZip;
            remaining -= n;
        }
        return out;
    }
    return Error.UnsupportedCompression;
}

// ─── Content types ─ [Content_Types].xml ─────────────────────────────

fn resolveContentTypes(arena: std.mem.Allocator, parts: []Part) !void {
    const ct_part = blk: {
        for (parts) |p| {
            if (std.mem.eql(u8, p.name, "[Content_Types].xml")) break :blk p;
        }
        // No content-types part — leave every Part.content_type as null.
        return;
    };
    const xml = ct_part.bytes;

    // Phase 1: collect Default extension → content type.
    // Phase 2: collect Override partName → content type.
    // Phase 3: assign each part's content_type (Override wins).

    var i: usize = 0;
    while (liveIndexOfPos(xml, i, "<Default")) |pos| {
        const end = xmlStartTagEnd(xml, pos) orelse break;
        const attrs = xml[pos..end];
        const ext_raw = attrAtSlice(attrs, "Extension") orelse {
            i = end + 1;
            continue;
        };
        const ct_raw = attrAtSlice(attrs, "ContentType") orelse {
            i = end + 1;
            continue;
        };
        const ext = try decodeXmlEntities(arena, ext_raw);
        const ct = try decodeXmlEntities(arena, ct_raw);
        // Apply default to every part with that extension.
        for (parts) |*p| {
            if (p.content_type != null) continue; // Override may set later
            if (extensionEql(p.name, ext)) p.content_type = ct;
        }
        i = end + 1;
    }
    i = 0;
    while (liveIndexOfPos(xml, i, "<Override")) |pos| {
        const end = xmlStartTagEnd(xml, pos) orelse break;
        const attrs = xml[pos..end];
        const part_name_raw = attrAtSlice(attrs, "PartName") orelse {
            i = end + 1;
            continue;
        };
        const ct_raw = attrAtSlice(attrs, "ContentType") orelse {
            i = end + 1;
            continue;
        };
        // PartName / ContentType are XML attribute values: a `/`
        // in a name appears literal but `&`, `<`, `>` etc must be
        // escaped on wire. Decode before comparing against ZIP
        // entry names (which carry the literal bytes), otherwise a
        // part named `xl/a&b.xml` (escaped as `xl/a&amp;b.xml`)
        // never matches and gets a null content_type.
        const part_name = try decodeXmlEntities(arena, part_name_raw);
        const ct = try decodeXmlEntities(arena, ct_raw);
        // PartName starts with `/`; strip to match part.name.
        const stripped = if (part_name.len > 0 and part_name[0] == '/') part_name[1..] else part_name;
        for (parts) |*p| {
            if (std.mem.eql(u8, p.name, stripped)) {
                p.content_type = ct;
                break;
            }
        }
        i = end + 1;
    }
}

fn attrAtSlice(attrs: []const u8, key: []const u8) ?[]const u8 {
    return xmlAttrValue(attrs, key);
}

/// Find the `>` that closes the start tag opening at `open_pos`,
/// skipping `>` characters inside quoted attribute values — `>` is
/// legal there per XML 1.0 §2.4 (only `<` and `&` must be escaped),
/// and shape names like `name="a>b"` occur in real parts. Returns
/// null when the tag never closes.
pub fn xmlStartTagEnd(xml: []const u8, open_pos: usize) ?usize {
    std.debug.assert(open_pos < xml.len);
    std.debug.assert(xml[open_pos] == '<');
    var i = open_pos + 1;
    while (i < xml.len) {
        const c = xml[i];
        if (c == '"' or c == '\'') {
            const close = std.mem.indexOfScalarPos(u8, xml, i + 1, c) orelse return null;
            i = close + 1;
            continue;
        }
        if (c == '>') return i;
        i += 1;
    }
    return null;
}

/// Extract one attribute's value from a start tag by lexing the
/// attribute list left to right — never by substring search, so a
/// value containing `=`, `>` or a lookalike key can neither
/// terminate nor satisfy the scan. Accepts both quote styles and
/// XML 1.0 §3.1 `Eq` whitespace on either side of `=`; both are
/// legal spellings some non-Microsoft producers (libreoffice,
/// hand-edited .rels files) emit, and missing them would silently
/// leave content types unresolved, relationships unparsed, and — on
/// the drawing-append path — ids double-allocated.
///
/// `tag` is the start tag WITHOUT its closing `>` (as cut by
/// `xmlStartTagEnd`), either from its `<` (the element name is
/// skipped) or from anywhere inside the attribute list. The value is
/// returned raw — entity decoding is the caller's choice. Malformed
/// attribute syntax stops the lex and returns null.
pub fn xmlAttrValue(tag: []const u8, name: []const u8) ?[]const u8 {
    std.debug.assert(name.len > 0);
    var i: usize = 0;
    // Skip `<` + element name when handed a full start tag.
    if (i < tag.len and tag[i] == '<') {
        i += 1;
        while (i < tag.len and
            !std.ascii.isWhitespace(tag[i]) and tag[i] != '/' and tag[i] != '>')
        {
            i += 1;
        }
    }
    while (i < tag.len) {
        if (std.ascii.isWhitespace(tag[i]) or tag[i] == '/') {
            i += 1;
            continue;
        }
        const name_start = i;
        while (i < tag.len and !std.ascii.isWhitespace(tag[i]) and tag[i] != '=') i += 1;
        const attr_name = tag[name_start..i];
        while (i < tag.len and std.ascii.isWhitespace(tag[i])) i += 1;
        if (i >= tag.len or tag[i] != '=') return null;
        i += 1;
        while (i < tag.len and std.ascii.isWhitespace(tag[i])) i += 1;
        if (i >= tag.len) return null;
        const quote = tag[i];
        if (quote != '"' and quote != '\'') return null;
        const val_start = i + 1;
        const val_end = std.mem.indexOfScalarPos(u8, tag, val_start, quote) orelse return null;
        if (std.mem.eql(u8, attr_name, name)) return tag[val_start..val_end];
        i = val_end + 1;
    }
    return null;
}

/// Index of the first occurrence of `needle` at or after `start` that
/// lies in LIVE markup — outside comments, CDATA sections, and
/// processing instructions (the XML declaration included). Element
/// scans that key on raw bytes use this so commented-out lookalikes,
/// PI payloads (which may legally contain `<!--`), and CDATA text are
/// never mistaken for markup. An unterminated special region hides
/// everything after its opener. Doctypes are not handled — they do
/// not occur in OOXML parts.
pub fn liveIndexOfPos(xml: []const u8, start: usize, needle: []const u8) ?usize {
    std.debug.assert(needle.len > 0);
    const Special = struct {
        open: []const u8,
        close: []const u8,
    };
    const specials = [_]Special{
        .{ .open = "<!--", .close = "-->" },
        .{ .open = "<![CDATA[", .close = "]]>" },
        .{ .open = "<?", .close = "?>" },
    };
    var i = start;
    while (i < xml.len) {
        const cand = std.mem.indexOfPos(u8, xml, i, needle) orelse return null;
        // The earliest special-region opener in [i, cand) decides:
        // none → cand is live; one → skip its whole region (which may
        // swallow cand) and rescan.
        var nearest: ?usize = null;
        var nearest_close: []const u8 = undefined;
        var nearest_open_len: usize = 0;
        for (specials) |sp| {
            if (std.mem.indexOfPos(u8, xml[0..cand], i, sp.open)) |p| {
                if (nearest == null or p < nearest.?) {
                    nearest = p;
                    nearest_close = sp.close;
                    nearest_open_len = sp.open.len;
                }
            }
        }
        const open = nearest orelse return cand;
        const close = std.mem.indexOfPos(u8, xml, open + nearest_open_len, nearest_close) orelse
            return null;
        i = close + nearest_close.len;
    }
    return null;
}

/// Last live occurrence of `needle`, or null. The splice helpers use
/// this so a close tag quoted inside an epilog comment cannot steal
/// the splice point from the real one.
pub fn liveLastIndexOf(xml: []const u8, needle: []const u8) ?usize {
    var last: ?usize = null;
    var i: usize = 0;
    while (liveIndexOfPos(xml, i, needle)) |p| {
        last = p;
        i = p + 1;
    }
    return last;
}

fn extensionEql(name: []const u8, ext: []const u8) bool {
    if (std.mem.lastIndexOfScalar(u8, name, '.')) |dot| {
        return std.ascii.eqlIgnoreCase(name[dot + 1 ..], ext);
    }
    return false;
}

// ─── Relationships ─ _rels/X.rels ────────────────────────────────────

/// Return the part name a `_rels/Y.rels` file describes, allocating
/// the result inside `arena`. Examples:
///   "_rels/.rels"                    → ""             (package level)
///   "xl/_rels/workbook.xml.rels"     → "xl/workbook.xml"
///   "xl/worksheets/_rels/sheet1.xml.rels" → "xl/worksheets/sheet1.xml"
/// Returns null if `name` isn't a `_rels/<base>.rels` shape.
/// Remove `<Override PartName="/<name>" …/>` from a `[Content_Types].xml`
/// body. Returns allocator-owned bytes; unchanged input yields a copy.
fn removeContentTypeOverride(
    allocator: std.mem.Allocator,
    ct_xml: []const u8,
    part_name: []const u8,
) ![]u8 {
    // OOXML writes override targets absolute-from-package-root.
    const needle = try std.fmt.allocPrint(allocator, "PartName=\"/{s}\"", .{part_name});
    defer allocator.free(needle);
    return removeSelfClosingElementContaining(allocator, ct_xml, "<Override", needle);
}

/// Remove every `<Relationship … Target="…"/>` whose target resolves to
/// `part_name`. Both the bare relative form and the leading-slash
/// absolute form appear in the wild, so both are matched.
fn removeRelationshipsTo(
    allocator: std.mem.Allocator,
    rels_xml: []const u8,
    part_name: []const u8,
) ![]u8 {
    const rel = try std.fmt.allocPrint(allocator, "Target=\"{s}\"", .{part_name});
    defer allocator.free(rel);
    const abs = try std.fmt.allocPrint(allocator, "Target=\"/{s}\"", .{part_name});
    defer allocator.free(abs);

    const once = try removeSelfClosingElementContaining(allocator, rels_xml, "<Relationship", rel);
    defer allocator.free(once);
    return removeSelfClosingElementContaining(allocator, once, "<Relationship", abs);
}

/// Copy `xml` minus every `<open … needle … >` element. Used for the
/// flat, self-closing elements that Content_Types and .rels are built
/// from, so no nesting handling is required.
fn removeSelfClosingElementContaining(
    allocator: std.mem.Allocator,
    xml: []const u8,
    open_tag: []const u8,
    needle: []const u8,
) ![]u8 {
    var out: std.ArrayListUnmanaged(u8) = .empty;
    errdefer out.deinit(allocator);

    var i: usize = 0;
    while (i < xml.len) {
        const start = std.mem.indexOfPos(u8, xml, i, open_tag) orelse break;
        const gt = std.mem.indexOfScalarPos(u8, xml, start, '>') orelse break;
        const elem = xml[start .. gt + 1];
        if (std.mem.indexOf(u8, elem, needle) != null) {
            // Drop the element: copy up to it, resume after it.
            try out.appendSlice(allocator, xml[i..start]);
        } else {
            try out.appendSlice(allocator, xml[i .. gt + 1]);
        }
        i = gt + 1;
    }
    try out.appendSlice(allocator, xml[i..]);
    return out.toOwnedSlice(allocator);
}

fn relsOwner(arena: std.mem.Allocator, name: []const u8) !?[]const u8 {
    if (!std.mem.endsWith(u8, name, ".rels")) return null;
    // Find the `_rels` directory segment. It's either prefix-less
    // ("_rels/...") or trailing-prefix ("X/_rels/...").
    const marker = "_rels/";
    const marker_pos = std.mem.lastIndexOf(u8, name, marker) orelse return null;
    // Marker must end with the last `/` before the basename.
    const last_slash = std.mem.lastIndexOfScalar(u8, name, '/') orelse return null;
    if (marker_pos + marker.len - 1 != last_slash) return null;

    const prefix = if (marker_pos == 0) "" else name[0..marker_pos];
    const filename_start = last_slash + 1;
    const filename_end = name.len - ".rels".len;
    if (filename_end < filename_start) return null;
    const filename = name[filename_start..filename_end];

    if (filename.len == 0) {
        // Package-level rels: "_rels/.rels" → owner = "".
        return try arena.dupe(u8, "");
    }
    if (prefix.len == 0) {
        return try arena.dupe(u8, filename);
    }
    // OPC part names always use '/' regardless of host OS. Concatenate
    // explicitly rather than via std.fs.path.join, which switches to
    // '\' on Windows and would silently break relationship lookup
    // (every callsite — store.rels, drawing walkers — keys on '/').
    // `prefix` is `name[0..marker_pos]` where marker is "_rels/", so
    // it already ends with the trailing '/' — no extra separator.
    std.debug.assert(prefix[prefix.len - 1] == '/');
    const out = try arena.alloc(u8, prefix.len + filename.len);
    @memcpy(out[0..prefix.len], prefix);
    @memcpy(out[prefix.len..], filename);
    return out;
}

/// Parse a rels part's bytes into Relationship entries. Public so
/// callers holding bytes the `rels_by_owner` cache doesn't cover (a
/// rels part created by `addPart` in the same session — only
/// `replacePart` refreshes the cache) can chase them with the SAME
/// quote-tolerant, entity-decoding parser that fills the cache,
/// rather than a divergent scanner. `arena` owns every returned
/// slice; callers typically pass a temporary ArenaAllocator.
pub fn parseRelationships(arena: std.mem.Allocator, xml: []const u8) ![]Relationship {
    var out: std.ArrayListUnmanaged(Relationship) = .empty;
    var i: usize = 0;
    while (liveIndexOfPos(xml, i, "<Relationship")) |pos| {
        // Skip `<Relationships ...>` (the wrapper) — only match
        // `<Relationship ` or `<Relationship>` followed by
        // attributes; the root tag is `<Relationships>` (with `s`).
        const after = pos + "<Relationship".len;
        if (after >= xml.len) break;
        const c = xml[after];
        if (c != ' ' and c != '\t' and c != '\n' and c != '\r' and c != '/' and c != '>') {
            i = after;
            continue;
        }
        const end = xmlStartTagEnd(xml, pos) orelse break;
        const attrs = xml[pos..end];
        const id = attrAtSlice(attrs, "Id") orelse {
            i = end + 1;
            continue;
        };
        const rtype = attrAtSlice(attrs, "Type") orelse {
            i = end + 1;
            continue;
        };
        const target = attrAtSlice(attrs, "Target") orelse {
            i = end + 1;
            continue;
        };
        const mode_str = attrAtSlice(attrs, "TargetMode");
        const target_mode: TargetMode = if (mode_str) |m|
            (if (std.mem.eql(u8, m, "External")) .external else .internal)
        else
            .internal;

        try out.append(arena, .{
            // Id and Type stored decoded — relTargetForId compares
            // decoded forms on both sides so lookups stay consistent
            // even when the referring `r:id="…"` attribute contains
            // entities. Target gets the same treatment because its
            // value flows directly into ZIP part-name resolution.
            .id = try decodeXmlEntities(arena, id),
            .type = try decodeXmlEntities(arena, rtype),
            .target = try decodeXmlEntities(arena, target),
            .target_mode = target_mode,
        });
        i = end + 1;
    }
    return out.toOwnedSlice(arena);
}

/// Decode the five canonical XML named entities (`&amp; &lt; &gt;
/// &quot; &apos;`) plus numeric character references (`&#N;` and
/// `&#xN;`) into their literal forms, returning a fresh arena-owned
/// slice. Code points are UTF-8 encoded. Unknown named entities pass
/// through verbatim. Malformed numeric refs (no closing `;`, empty
/// digit run, out-of-range code point) also pass through verbatim,
/// matching the lenient behaviour of OOXML readers.
/// Public alongside `parseRelationships`: id-allocation scans decode
/// attribute values through the same routine the parsers use.
pub fn decodeXmlEntities(arena: std.mem.Allocator, s: []const u8) ![]u8 {
    if (std.mem.indexOfScalar(u8, s, '&') == null) return arena.dupe(u8, s);
    var out: std.ArrayListUnmanaged(u8) = .empty;
    errdefer out.deinit(arena);
    try out.ensureTotalCapacity(arena, s.len);
    var i: usize = 0;
    while (i < s.len) {
        if (s[i] == '&') {
            const remain = s[i..];
            if (std.mem.startsWith(u8, remain, "&amp;")) {
                try out.append(arena, '&');
                i += 5;
                continue;
            }
            if (std.mem.startsWith(u8, remain, "&lt;")) {
                try out.append(arena, '<');
                i += 4;
                continue;
            }
            if (std.mem.startsWith(u8, remain, "&gt;")) {
                try out.append(arena, '>');
                i += 4;
                continue;
            }
            if (std.mem.startsWith(u8, remain, "&quot;")) {
                try out.append(arena, '"');
                i += 6;
                continue;
            }
            if (std.mem.startsWith(u8, remain, "&apos;")) {
                try out.append(arena, '\'');
                i += 6;
                continue;
            }
            // Numeric character reference: `&#N;` (decimal) or
            // `&#xN;` (hex). Decode the code point and append its
            // UTF-8 representation. Any malformation falls through
            // to the literal-`&` append below.
            if (std.mem.startsWith(u8, remain, "&#")) {
                if (decodeNumericRef(remain)) |info| {
                    try out.appendSlice(arena, info.utf8[0..info.utf8_len]);
                    i += info.consumed;
                    continue;
                }
            }
        }
        try out.append(arena, s[i]);
        i += 1;
    }
    return out.toOwnedSlice(arena);
}

pub const NumericRef = struct {
    utf8: [4]u8,
    utf8_len: u3,
    consumed: usize,
};

pub fn decodeNumericRef(s: []const u8) ?NumericRef {
    // s starts with "&#". Find the closing ';' and the digit run.
    if (s.len < 4) return null; // need at least "&#0;"
    var digit_start: usize = 2;
    var base: u8 = 10;
    if (s[2] == 'x' or s[2] == 'X') {
        digit_start = 3;
        base = 16;
    }
    const semi = std.mem.indexOfScalarPos(u8, s, digit_start, ';') orelse return null;
    if (semi == digit_start) return null; // empty digit run
    // XML 1.0 §4.1 restricts numeric refs to digit chars only —
    // no `+`, `-`, `_`, or whitespace. parseInt accepts `+` and
    // `_` which would smuggle malformed refs into the decoder, so
    // validate the digit run first.
    const digits = s[digit_start..semi];
    for (digits) |c| {
        const ok = if (base == 10)
            (c >= '0' and c <= '9')
        else
            ((c >= '0' and c <= '9') or (c >= 'a' and c <= 'f') or (c >= 'A' and c <= 'F'));
        if (!ok) return null;
    }
    const code = std.fmt.parseInt(u32, digits, base) catch return null;
    if (code > 0x10FFFF) return null;
    // XML 1.0 forbids most C0 controls; skip them rather than
    // emitting invalid UTF-8 / non-XML content. Allow tab / LF / CR.
    if (code < 0x20 and code != 0x09 and code != 0x0A and code != 0x0D) return null;
    var ref: NumericRef = .{ .utf8 = undefined, .utf8_len = 0, .consumed = semi + 1 };
    const len = std.unicode.utf8Encode(@intCast(code), &ref.utf8) catch return null;
    ref.utf8_len = @intCast(len);
    return ref;
}

fn parentDir(name: []const u8) []const u8 {
    if (std.mem.lastIndexOfScalar(u8, name, '/')) |slash| return name[0..slash];
    return "";
}

// ─── Tests ────────────────────────────────────────────────────────────

test "PartStore.open: enumerates parts of frictionless_2sheets.xlsx" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    const fixture = "tests/corpus/frictionless_2sheets.xlsx";
    std.Io.Dir.cwd().access(io, fixture, .{}) catch return error.SkipZigTest;

    var store = try PartStore.open(std.testing.allocator, io, fixture);
    defer store.deinit();

    const names = try store.partNames();
    try std.testing.expect(names.len > 0);

    // Required OOXML parts must always be present.
    var seen_ct = false;
    var seen_wb = false;
    for (names) |n| {
        if (std.mem.eql(u8, n, "[Content_Types].xml")) seen_ct = true;
        if (std.mem.eql(u8, n, "xl/workbook.xml")) seen_wb = true;
    }
    try std.testing.expect(seen_ct);
    try std.testing.expect(seen_wb);
}

test "PartStore.part: workbook.xml has a workbook content type" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    const fixture = "tests/corpus/frictionless_2sheets.xlsx";
    std.Io.Dir.cwd().access(io, fixture, .{}) catch return error.SkipZigTest;

    var store = try PartStore.open(std.testing.allocator, io, fixture);
    defer store.deinit();

    const wb = try store.part("xl/workbook.xml") orelse return error.TestUnexpectedResult;
    const ct = wb.content_type orelse return error.TestUnexpectedResult;
    // OOXML ContentType for the workbook part. The canonical
    // workbook content-type is `…spreadsheetml.sheet.main+xml`
    // (NOT "…workbook+xml" — common confusion).
    try std.testing.expect(std.mem.indexOf(u8, ct, "spreadsheetml") != null);
    try std.testing.expect(std.mem.indexOf(u8, ct, "sheet.main") != null);
    // Bytes must be non-empty.
    try std.testing.expect(wb.bytes.len > 0);
    // workbook.xml is XML; should start with '<'.
    try std.testing.expectEqual(@as(u8, '<'), wb.bytes[0]);
}

test "PartStore.rels: package-root + workbook rels parse" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    const fixture = "tests/corpus/frictionless_2sheets.xlsx";
    std.Io.Dir.cwd().access(io, fixture, .{}) catch return error.SkipZigTest;

    var store = try PartStore.open(std.testing.allocator, io, fixture);
    defer store.deinit();

    // Package-level rels (`_rels/.rels`, owner = "") must point at
    // `xl/workbook.xml`.
    const root_rels = store.rels("");
    try std.testing.expect(root_rels.len > 0);
    var found_wb = false;
    for (root_rels) |r| {
        if (std.mem.endsWith(u8, r.target, "workbook.xml")) found_wb = true;
    }
    try std.testing.expect(found_wb);

    // Workbook rels (`xl/_rels/workbook.xml.rels`, owner =
    // `xl/workbook.xml`) must list at least one worksheet.
    const wb_rels = store.rels("xl/workbook.xml");
    try std.testing.expect(wb_rels.len > 0);
    var found_ws = false;
    for (wb_rels) |r| {
        if (std.mem.indexOf(u8, r.target, "worksheets/sheet") != null) found_ws = true;
    }
    try std.testing.expect(found_ws);
}

test "relsOwner: shape decoder" {
    var arena = std.heap.ArenaAllocator.init(std.testing.allocator);
    defer arena.deinit();
    const a = arena.allocator();

    try std.testing.expectEqualStrings("", (try relsOwner(a, "_rels/.rels")).?);
    try std.testing.expectEqualStrings(
        "xl/workbook.xml",
        (try relsOwner(a, "xl/_rels/workbook.xml.rels")).?,
    );
    try std.testing.expectEqualStrings(
        "xl/worksheets/sheet1.xml",
        (try relsOwner(a, "xl/worksheets/_rels/sheet1.xml.rels")).?,
    );
    // Negative cases.
    try std.testing.expectEqual(@as(?[]const u8, null), try relsOwner(a, "xl/workbook.xml"));
    try std.testing.expectEqual(@as(?[]const u8, null), try relsOwner(a, "xl/foo.rels"));
}

test "PartStore.resolve: relative + absolute targets" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    const fixture = "tests/corpus/frictionless_2sheets.xlsx";
    std.Io.Dir.cwd().access(io, fixture, .{}) catch return error.SkipZigTest;

    var store = try PartStore.open(std.testing.allocator, io, fixture);
    defer store.deinit();

    // Relative: workbook.xml's rels target "worksheets/sheet1.xml" →
    // "xl/worksheets/sheet1.xml".
    const r1 = (try store.resolve("xl/workbook.xml", "worksheets/sheet1.xml")).?;
    try std.testing.expectEqualStrings("xl/worksheets/sheet1.xml", r1);

    // Relative with `..`: "../sharedStrings.xml" from a worksheet
    // resolves to "xl/sharedStrings.xml".
    const r2 = (try store.resolve("xl/worksheets/sheet1.xml", "../sharedStrings.xml")).?;
    try std.testing.expectEqualStrings("xl/sharedStrings.xml", r2);

    // Absolute: "/xl/workbook.xml" → "xl/workbook.xml".
    const r3 = (try store.resolve("anywhere", "/xl/workbook.xml")).?;
    try std.testing.expectEqualStrings("xl/workbook.xml", r3);
}

test "PartStore.imageParts: extract embedded images (C2a MVP)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    const fixture = "tests/corpus/poi_58325_db.xlsx";
    std.Io.Dir.cwd().access(io, fixture, .{}) catch return error.SkipZigTest;

    var store = try PartStore.open(std.testing.allocator, io, fixture);
    defer store.deinit();

    const images = try store.imageParts();
    // poi_58325_db.xlsx ships 4 image parts under xl/media/. Confirm
    // imageParts surfaces them as bytes-bearing Parts (no XML parse,
    // no per-sheet anchor attribution — that's the future drawings()
    // parser's job).
    try std.testing.expect(images.len > 0);
    for (images) |p| {
        try std.testing.expect(std.mem.startsWith(u8, p.name, "xl/media/"));
        try std.testing.expect(p.bytes.len > 0);
        const ct = p.content_type orelse return error.TestUnexpectedResult;
        try std.testing.expect(std.mem.startsWith(u8, ct, "image/"));
    }
}

test "PartStore: data descriptors detected in fixtures with flag 0x0008" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    const fixture = "tests/corpus/frictionless_2sheets.xlsx";
    std.Io.Dir.cwd().access(io, fixture, .{}) catch return error.SkipZigTest;

    var store = try PartStore.open(std.testing.allocator, io, fixture);
    defer store.deinit();

    // frictionless_2sheets.xlsx has data-descriptor flag set on every
    // entry (verified via Python's zipfile module). Confirm the
    // scanner detects + records the descriptor length so save()
    // copies those bytes.
    var with_dd: usize = 0;
    for (store.entries) |e| {
        if (e.data_descriptor_len > 0) with_dd += 1;
    }
    try std.testing.expect(with_dd > 0);
}

test "PartStore.save: byte-preserving round-trip with no mutations" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    const fixture = "tests/corpus/frictionless_2sheets.xlsx";
    std.Io.Dir.cwd().access(io, fixture, .{}) catch return error.SkipZigTest;

    var tmp = std.testing.tmpDir(.{});
    defer tmp.cleanup();
    const dir = try tmp.dir.realPathFileAlloc(io, ".", std.testing.allocator);
    defer std.testing.allocator.free(dir);
    const out_path = try std.fs.path.join(std.testing.allocator, &.{ dir, "round_trip.xlsx" });
    defer std.testing.allocator.free(out_path);

    {
        var store = try PartStore.open(std.testing.allocator, io, fixture);
        defer store.deinit();
        try store.save(io, out_path);
    }

    // Re-open the saved file. Every part must decompress to the
    // same bytes as the source.
    var src = try PartStore.open(std.testing.allocator, io, fixture);
    defer src.deinit();
    var dst = try PartStore.open(std.testing.allocator, io, out_path);
    defer dst.deinit();

    try std.testing.expectEqual(src.parts.len, dst.parts.len);
    for (src.parts, dst.parts) |s, d| {
        try std.testing.expectEqualStrings(s.name, d.name);
        try std.testing.expectEqualSlices(u8, s.bytes, d.bytes);
    }
}

test "PartStore.replacePart + save: replaced part has new bytes; others untouched" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    const fixture = "tests/corpus/frictionless_2sheets.xlsx";
    std.Io.Dir.cwd().access(io, fixture, .{}) catch return error.SkipZigTest;

    var tmp = std.testing.tmpDir(.{});
    defer tmp.cleanup();
    const dir = try tmp.dir.realPathFileAlloc(io, ".", std.testing.allocator);
    defer std.testing.allocator.free(dir);
    const out_path = try std.fs.path.join(std.testing.allocator, &.{ dir, "modified.xlsx" });
    defer std.testing.allocator.free(out_path);

    const replacement: []const u8 =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" ++
        "<test xmlns=\"http://example.com/zlsx\">replaced</test>";

    var src_workbook_bytes: []const u8 = undefined;
    {
        var store = try PartStore.open(std.testing.allocator, io, fixture);
        defer store.deinit();
        const wb_part = try store.part("xl/workbook.xml") orelse return error.TestUnexpectedResult;
        src_workbook_bytes = wb_part.bytes;
        // Pick a small XML part that's safe to overwrite for the
        // round-trip test. workbook.xml ensures we exercise the
        // override path on a part that exists in every fixture.
        try store.replacePart("xl/workbook.xml", replacement);
        try store.save(io, out_path);
    }

    var dst = try PartStore.open(std.testing.allocator, io, out_path);
    defer dst.deinit();

    // Replaced part has the new bytes.
    const wb = try dst.part("xl/workbook.xml") orelse return error.TestUnexpectedResult;
    try std.testing.expectEqualSlices(u8, replacement, wb.bytes);

    // Sanity: at least one OTHER part still matches the source's
    // decompressed bytes byte-for-byte. sharedStrings.xml in
    // frictionless_2sheets is a stable Override-typed part.
    var src = try PartStore.open(std.testing.allocator, io, fixture);
    defer src.deinit();
    if (try src.part("xl/sharedStrings.xml")) |s| {
        const d = try dst.part("xl/sharedStrings.xml") orelse return error.TestUnexpectedResult;
        try std.testing.expectEqualSlices(u8, s.bytes, d.bytes);
    }
}

test "PartStore.replacePart: large input round-trips through deflate" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    const fixture = "tests/corpus/frictionless_2sheets.xlsx";
    std.Io.Dir.cwd().access(io, fixture, .{}) catch return error.SkipZigTest;

    var tmp = std.testing.tmpDir(.{});
    defer tmp.cleanup();
    const dir = try tmp.dir.realPathFileAlloc(io, ".", std.testing.allocator);
    defer std.testing.allocator.free(dir);
    const out_path = try std.fs.path.join(std.testing.allocator, &.{ dir, "deflate_round_trip.xlsx" });
    defer std.testing.allocator.free(out_path);

    // Build a large, highly-compressible payload (10 KiB of repeated
    // ASCII). M2.5 routes inputs >= 1 KiB through deflate and falls
    // back to STORED only if deflate doesn't shrink — this should
    // hit the deflate path.
    var buf: [10 * 1024]u8 = undefined;
    @memset(&buf, 'A');
    const replacement: []const u8 = &buf;

    {
        var store = try PartStore.open(std.testing.allocator, io, fixture);
        defer store.deinit();
        try store.replacePart("xl/workbook.xml", replacement);
        // Deferred (M10b): the replace stages raw bytes only; the save
        // below is what compresses.
        const idx = store.findIndex("xl/workbook.xml").?;
        try std.testing.expect(store.overrides[idx].? == .pending);
        try store.save(io, out_path);
        // Compression must shrink the payload — 10 KiB of one byte
        // is the trivial deflate case.
        const ov = store.overrides[idx].?.compressed;
        try std.testing.expectEqual(@as(u16, 8), ov.compression_method);
        try std.testing.expect(ov.payload.len < replacement.len);
    }

    var dst = try PartStore.open(std.testing.allocator, io, out_path);
    defer dst.deinit();
    const wb = try dst.part("xl/workbook.xml") orelse return error.TestUnexpectedResult;
    try std.testing.expectEqualSlices(u8, replacement, wb.bytes);
}

test "PartStore.replacePart: a payload past the exact-block threshold round-trips (§9.1 M10m)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    const fixture = "tests/corpus/frictionless_2sheets.xlsx";
    std.Io.Dir.cwd().access(io, fixture, .{}) catch return error.SkipZigTest;

    var tmp = std.testing.tmpDir(.{});
    defer tmp.cleanup();
    const dir = try tmp.dir.realPathFileAlloc(io, ".", std.testing.allocator);
    defer std.testing.allocator.free(dir);
    const out_path = try std.fs.path.join(std.testing.allocator, &.{ dir, "big_block.xlsx" });
    defer std.testing.allocator.free(out_path);

    // Straddle the threshold in one test: the small replace rides the
    // arena, the large one gets its own exact block, and both must be
    // readable through `part()` afterwards and land in the saved
    // archive. The allocator's leak check is the other half of the
    // assertion — the block is freed by `deinit`, not by the arena.
    const big = try std.testing.allocator.alloc(u8, PartStore.big_payload_bytes + 7);
    defer std.testing.allocator.free(big);
    for (big, 0..) |*b, i| b.* = @intCast('a' + (i % 26));
    const small: []const u8 = "<workbook/>";

    {
        var store = try PartStore.open(std.testing.allocator, io, fixture);
        defer store.deinit();
        try store.replacePart("xl/workbook.xml", small);
        try std.testing.expectEqual(@as(usize, 0), store.big_parts.items.len);
        try store.replacePart("xl/worksheets/sheet1.xml", big);
        try std.testing.expectEqual(@as(usize, 1), store.big_parts.items.len);

        const staged = try store.part("xl/worksheets/sheet1.xml") orelse
            return error.TestUnexpectedResult;
        try std.testing.expectEqualSlices(u8, big, staged.bytes);
        try store.save(io, out_path);
    }

    var dst = try PartStore.open(std.testing.allocator, io, out_path);
    defer dst.deinit();
    const sheet = try dst.part("xl/worksheets/sheet1.xml") orelse
        return error.TestUnexpectedResult;
    try std.testing.expectEqualSlices(u8, big, sheet.bytes);
    const wb = try dst.part("xl/workbook.xml") orelse return error.TestUnexpectedResult;
    try std.testing.expectEqualSlices(u8, small, wb.bytes);
}

test "materializeAt: a large part is decompressed into an exact block, not the arena (§9.1c M10t)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    const fixture = "tests/corpus/frictionless_2sheets.xlsx";
    std.Io.Dir.cwd().access(io, fixture, .{}) catch return error.SkipZigTest;

    var tmp = std.testing.tmpDir(.{});
    defer tmp.cleanup();
    const dir = try tmp.dir.realPathFileAlloc(io, ".", std.testing.allocator);
    defer std.testing.allocator.free(dir);
    const out_path = try std.fs.path.join(std.testing.allocator, &.{ dir, "big_read.xlsx" });
    defer std.testing.allocator.free(out_path);

    const big = try std.testing.allocator.alloc(u8, PartStore.big_payload_bytes + 7);
    defer std.testing.allocator.free(big);
    for (big, 0..) |*b, i| b.* = @intCast('a' + (i % 26));

    // Build an archive that HAS a large part, then read it back. M10m
    // gave the replace path its exact block; the read path kept the
    // arena, and the arena sizes a chunk at 1.5 × (held + asked) — so
    // materializing this part bought half a megabyte of chunk nothing
    // else would ever ask for. `id7`, live in all 24 eras.
    {
        var src = try PartStore.open(std.testing.allocator, io, fixture);
        defer src.deinit();
        try src.replacePart("xl/worksheets/sheet1.xml", big);
        try src.save(io, out_path);
    }

    var dst = try PartStore.open(std.testing.allocator, io, out_path);
    defer dst.deinit();
    // Nothing is materialized yet: `open` decompresses only the
    // structural parts, and this one is neither.
    try std.testing.expectEqual(@as(usize, 0), dst.big_parts.items.len);
    const before = dst.arena.queryCapacity();

    const sheet = try dst.part("xl/worksheets/sheet1.xml") orelse
        return error.TestUnexpectedResult;
    try std.testing.expectEqualSlices(u8, big, sheet.bytes);

    // The block exists and the arena did not grow to hold it. Both
    // halves matter: the first says the payload took the exact path,
    // the second says the ladder never stepped. The testing
    // allocator's leak check is the third — `deinit` frees
    // `big_parts`, and a payload that took this path without being
    // recorded there would leak.
    try std.testing.expectEqual(@as(usize, 1), dst.big_parts.items.len);
    try std.testing.expectEqual(before, dst.arena.queryCapacity());
    try std.testing.expect(dst.residentBytes() >= big.len);
}

test "PartStore.residentBytes counts the out-of-arena blocks too" {
    // The retention ceiling reads this figure. M10m moved payloads
    // ≥ `big_payload_bytes` out of the arena, so a store that reported
    // only `arena.queryCapacity()` would say a 1 MiB replaced part
    // weighs nothing and let `max_retained_bytes` be overshot by the
    // whole block. Guard it where the block is created.
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    const fixture = "tests/corpus/frictionless_2sheets.xlsx";
    std.Io.Dir.cwd().access(io, fixture, .{}) catch return error.SkipZigTest;

    var store = try PartStore.open(std.testing.allocator, io, fixture);
    defer store.deinit();

    const before = store.residentBytes();
    try std.testing.expectEqual(@as(usize, 0), store.big_parts.items.len);

    const big = try std.testing.allocator.alloc(u8, PartStore.big_payload_bytes + 7);
    defer std.testing.allocator.free(big);
    @memset(big, 'z');
    try store.replacePart("xl/worksheets/sheet1.xml", big);

    try std.testing.expectEqual(@as(usize, 1), store.big_parts.items.len);
    // The block is out of the arena, so the growth has to come from the
    // `big_parts` term — arena capacity alone could not account for it.
    const after = store.residentBytes();
    try std.testing.expect(after - before >= big.len);
}

test "PartStore.replacePart: an EMPTY replacement survives a read before save" {
    // `materializeAt` treated `bytes.len == 0` as "not materialized yet"
    // and reloaded the source part over it. Before M10b that was only a
    // wrong `part()` read, because save compressed the override eagerly
    // at replace time; once M10b deferred the deflate to save, the
    // reloaded ORIGINAL became what got written — a caller that blanked
    // a part silently got the original back. The override slot, not the
    // byte length, says whether a part has content of its own.
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    const fixture = "tests/corpus/frictionless_2sheets.xlsx";
    std.Io.Dir.cwd().access(io, fixture, .{}) catch return error.SkipZigTest;

    var tmp = std.testing.tmpDir(.{});
    defer tmp.cleanup();
    const dir = try tmp.dir.realPathFileAlloc(io, ".", std.testing.allocator);
    defer std.testing.allocator.free(dir);
    const out_path = try std.fs.path.join(std.testing.allocator, &.{ dir, "blanked.xlsx" });
    defer std.testing.allocator.free(out_path);

    const target = "xl/worksheets/sheet1.xml";
    {
        var store = try PartStore.open(std.testing.allocator, io, fixture);
        defer store.deinit();

        const original = try store.part(target) orelse return error.TestUnexpectedResult;
        try std.testing.expect(original.bytes.len > 0);

        try store.replacePart(target, "");
        // The read is the trigger: it is what used to reload the source.
        const staged = try store.part(target) orelse return error.TestUnexpectedResult;
        try std.testing.expectEqual(@as(usize, 0), staged.bytes.len);

        try store.save(io, out_path);
    }

    var dst = try PartStore.open(std.testing.allocator, io, out_path);
    defer dst.deinit();
    const round = try dst.part(target) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqual(@as(usize, 0), round.bytes.len);
}

// ─── M5d1: cancellation at the archive seams ─────────────────────────

/// A `Watch` armed with both a flag and a far-future deadline, plus the
/// injected clock that arms the flag on its Nth read.
///
/// Both halves are needed to make "cancel arrives *during* the operation"
/// deterministic. The deadline forces a clock read at every poll; the
/// injected clock counts those reads and sets the flag at a chosen one;
/// the next poll sees the flag. Without the deadline the flag would have
/// to be set before the call, which only proves the entry-point check.
fn tripAfter(base: std.Io, polls: u64, flag: *volatile u8) control.Watch {
    const io = control.inject.wrap(base, .{ .trip_at = polls, .trip_flag = flag });
    return .init(io, .{
        .cancel = .{ .flag = flag },
        .deadline = .{ .nanoseconds = std.math.maxInt(i64) },
    });
}

/// 512 KiB of high-entropy bytes: eight `chunk_bytes` pieces that deflate
/// cannot shrink, so the STORED fallback keeps the staged payload the
/// same size as the input and the chunk arithmetic in these tests is the
/// arithmetic the code actually performs.
fn incompressible(alloc: std.mem.Allocator) ![]u8 {
    const buf = try alloc.alloc(u8, control.chunk_bytes * 8);
    var rng: std.Random.DefaultPrng = .init(0x5d1_5d1);
    rng.random().bytes(buf);
    return buf;
}

test "M5d1: cancel inside replacePart leaves the store byte-identical" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    const fixture = "tests/corpus/frictionless_2sheets.xlsx";
    std.Io.Dir.cwd().access(io, fixture, .{}) catch return error.SkipZigTest;

    var store = try PartStore.open(std.testing.allocator, io, fixture);
    defer store.deinit();

    const idx = store.findIndex("xl/workbook.xml").?;
    const before = (try store.part("xl/workbook.xml")).?.bytes;
    try std.testing.expect(store.overrides[idx] == null);

    // Large enough that the compression spans several chunks, so the
    // save-side trip below lands inside the deflate rather than at an
    // entry poll.
    const replacement = try std.testing.allocator.alloc(u8, control.chunk_bytes * 4);
    defer std.testing.allocator.free(replacement);
    @memset(replacement, 'A');

    // The replace's own poll (M10b: the deflate moved to the save, so
    // this is the one check the call still owes): an already-cancelled
    // run refuses before any field is written.
    var cancelled: u8 = 1;
    const w: control.Watch = .init(io, .{ .cancel = .{ .flag = &cancelled } });
    try std.testing.expectError(
        control.Error.Cancelled,
        store.replacePartControlled("xl/workbook.xml", replacement, w.poller()),
    );

    // Nothing staged, nothing mirrored: `part()` still answers with the
    // source bytes and the override slot is still empty.
    try std.testing.expect(store.overrides[idx] == null);
    try std.testing.expectEqualSlices(u8, before, (try store.part("xl/workbook.xml")).?.bytes);
    try std.testing.expect(!store.hasUnsavedChanges());

    // The deflate itself now runs under the SAVE's poller
    // (`materializeOverrides`), and a trip inside it must leave the
    // destination untouched and the staged replacement intact.
    try store.replacePart("xl/workbook.xml", replacement);
    var tmp = std.testing.tmpDir(.{});
    defer tmp.cleanup();
    const dir = try tmp.dir.realPathFileAlloc(io, ".", std.testing.allocator);
    defer std.testing.allocator.free(dir);
    const out_path = try std.fs.path.join(std.testing.allocator, &.{ dir, "book.xlsx" });
    defer std.testing.allocator.free(out_path);

    var flag: u8 = 0;
    var w2 = tripAfter(io, 3, &flag);
    try std.testing.expectError(
        control.Error.Cancelled,
        store.saveControlled(io, out_path, w2.poller()),
    );
    // The trip precedes `AtomicFile.init`, so not even a temp file
    // exists, and the mirror still answers with the replacement — a
    // cancelled save half-materialized at most, it un-staged nothing.
    try std.testing.expectError(
        error.FileNotFound,
        std.Io.Dir.cwd().access(io, out_path, .{}),
    );
    try std.testing.expectEqualSlices(u8, replacement, (try store.part("xl/workbook.xml")).?.bytes);

    // A retry without the trip completes and round-trips the bytes the
    // cancelled attempt was carrying.
    _ = try store.saveControlled(io, out_path, .none);
    var dst = try PartStore.open(std.testing.allocator, io, out_path);
    defer dst.deinit();
    try std.testing.expectEqualSlices(u8, replacement, (try dst.part("xl/workbook.xml")).?.bytes);
}

test "M5d1: cancel inside materialization mutates nothing" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    const fixture = "tests/corpus/frictionless_2sheets.xlsx";
    std.Io.Dir.cwd().access(io, fixture, .{}) catch return error.SkipZigTest;

    var store = try PartStore.open(std.testing.allocator, io, fixture);
    defer store.deinit();

    const name = "xl/worksheets/sheet1.xml";
    const idx = store.findIndex(name) orelse return error.SkipZigTest;
    // Must not already be materialized, or there is nothing to cancel.
    try std.testing.expectEqual(@as(usize, 0), store.parts[idx].bytes.len);

    var flag: u8 = 1; // already up: the first poll refuses
    var w = tripAfter(io, 1, &flag);
    try std.testing.expectError(
        control.Error.Cancelled,
        store.partControlled(name, w.poller()),
    );

    // The cache stayed empty, so a later uncancelled read still does the
    // work — a cancelled materialization must not poison the slot.
    try std.testing.expectEqual(@as(usize, 0), store.parts[idx].bytes.len);
    const p = (try store.part(name)).?;
    try std.testing.expect(p.bytes.len > 0);
    try std.testing.expectEqual(p.crc32, std.hash.Crc32.hash(p.bytes));
}

test "M5d1: cancel part-way THROUGH a materialization mutates nothing" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    const fixture = "tests/corpus/frictionless_2sheets.xlsx";
    std.Io.Dir.cwd().access(io, fixture, .{}) catch return error.SkipZigTest;

    var tmp = std.testing.tmpDir(.{});
    defer tmp.cleanup();
    const dir = try tmp.dir.realPathFileAlloc(io, ".", std.testing.allocator);
    defer std.testing.allocator.free(dir);
    const out_path = try std.fs.path.join(std.testing.allocator, &.{ dir, "big_part.xlsx" });
    defer std.testing.allocator.free(out_path);

    const big = try incompressible(std.testing.allocator);
    defer std.testing.allocator.free(big);

    // A package whose `xl/workbook.xml` is eight chunks long, written to
    // disk so the re-open below has something genuinely lazy to inflate.
    {
        var src = try PartStore.open(std.testing.allocator, io, fixture);
        defer src.deinit();
        try src.replacePart("xl/workbook.xml", big);
        try src.save(io, out_path);
    }

    var store = try PartStore.open(std.testing.allocator, io, out_path);
    defer store.deinit();
    const idx = store.findIndex("xl/workbook.xml").?;
    try std.testing.expectEqual(@as(usize, 0), store.parts[idx].bytes.len);

    var flag: u8 = 0;
    var w = tripAfter(io, 3, &flag);
    try std.testing.expectError(
        control.Error.Cancelled,
        store.partControlled("xl/workbook.xml", w.poller()),
    );
    // Cancelled three chunks in, with nothing published: the cache slot
    // is only written after the whole part has inflated and its CRC has
    // matched, so a partially-read part can never become visible.
    try std.testing.expectEqual(@as(usize, 0), store.parts[idx].bytes.len);

    const p = (try store.part("xl/workbook.xml")).?;
    try std.testing.expectEqualSlices(u8, big, p.bytes);
}

test "M5d1: cancel during save leaves no temp file and no destination" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    const fixture = "tests/corpus/frictionless_2sheets.xlsx";
    std.Io.Dir.cwd().access(io, fixture, .{}) catch return error.SkipZigTest;

    var tmp = std.testing.tmpDir(.{ .iterate = true });
    defer tmp.cleanup();
    const dir = try tmp.dir.realPathFileAlloc(io, ".", std.testing.allocator);
    defer std.testing.allocator.free(dir);
    const out_path = try std.fs.path.join(std.testing.allocator, &.{ dir, "cancelled.xlsx" });
    defer std.testing.allocator.free(out_path);

    var store = try PartStore.open(std.testing.allocator, io, fixture);
    defer store.deinit();

    // Trip a few polls in, i.e. part-way through streaming entries out.
    var flag: u8 = 0;
    var w = tripAfter(io, 3, &flag);
    try std.testing.expectError(
        control.Error.Cancelled,
        store.saveControlled(io, out_path, w.poller()),
    );

    // The destination never existed, so §5.7.9's "no output file"
    // promise applies in its literal form…
    try std.testing.expectError(
        error.FileNotFound,
        std.Io.Dir.cwd().access(io, out_path, .{}),
    );
    // …and the half-written temp file is gone with it.
    var it = tmp.dir.iterate();
    try std.testing.expectEqual(@as(?std.Io.Dir.Entry, null), try it.next(io));
}

test "M5d1: save polls at least once per 64 KiB of archive" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    const fixture = "tests/corpus/frictionless_2sheets.xlsx";
    std.Io.Dir.cwd().access(io, fixture, .{}) catch return error.SkipZigTest;

    var tmp = std.testing.tmpDir(.{});
    defer tmp.cleanup();
    const dir = try tmp.dir.realPathFileAlloc(io, ".", std.testing.allocator);
    defer std.testing.allocator.free(dir);
    const out_path = try std.fs.path.join(std.testing.allocator, &.{ dir, "measured.xlsx" });
    defer std.testing.allocator.free(out_path);

    var store = try PartStore.open(std.testing.allocator, io, fixture);
    defer store.deinit();

    // One deliberately large part, so the entry loop is not the only
    // thing generating polls.
    const big = try incompressible(std.testing.allocator);
    defer std.testing.allocator.free(big);
    try store.replacePart("xl/workbook.xml", big);

    // A deadline that cannot fire (the injected clock never advances)
    // turns every poll into exactly one counted clock read.
    const counting = control.inject.wrap(io, .{});
    const w: control.Watch = .init(counting, .{
        .deadline = .{ .nanoseconds = std.math.maxInt(i64) },
    });
    _ = try store.saveControlled(counting, out_path, w.poller());

    // What §5.5 owes for this save, computed from the code's own shape
    // rather than eyeballed: one poll before the temp file is created,
    // one per entry in each of the two passes, one immediately before
    // the commit region, and one per 64 KiB of the staged payload.
    const staged = store.overrides[store.findIndex("xl/workbook.xml").?].?.compressed;
    const payload_polls = control.chunkCount(staged.payload.len);
    const owed = 2 + 2 * store.entries.len + payload_polls;
    try std.testing.expect(control.inject.state.now_calls >= owed);
    // Not a vacuous bound: the large part alone accounts for eight of
    // them, so an unchunked seam cannot reach the total on entry polls.
    try std.testing.expectEqual(@as(usize, 8), payload_polls);
}

test "M5d1: an injected rename failure leaves memory AND the destination untouched" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    const fixture = "tests/corpus/frictionless_2sheets.xlsx";
    std.Io.Dir.cwd().access(io, fixture, .{}) catch return error.SkipZigTest;

    var tmp = std.testing.tmpDir(.{ .iterate = true });
    defer tmp.cleanup();
    const dir = try tmp.dir.realPathFileAlloc(io, ".", std.testing.allocator);
    defer std.testing.allocator.free(dir);
    const out_path = try std.fs.path.join(std.testing.allocator, &.{ dir, "book.xlsx" });
    defer std.testing.allocator.free(out_path);
    try tmp.dir.writeFile(io, .{ .sub_path = "book.xlsx", .data = "PRIOR BYTES" });

    var store = try PartStore.open(std.testing.allocator, io, fixture);
    defer store.deinit();
    try store.replacePart("xl/workbook.xml", "<workbook/>");
    const idx = store.findIndex("xl/workbook.xml").?;
    // Deferred (M10b): the replace stages raw bytes only.
    try std.testing.expect(store.overrides[idx].? == .pending);

    const faulty = control.inject.wrap(io, .{ .fail_rename = true });
    try std.testing.expectError(error.AccessDenied, store.saveControlled(faulty, out_path, .none));

    // Destination: the prior bytes, not a truncated or half-written
    // archive. §5.7.9's promise is precisely "unchanged until the commit
    // point", and the rename IS the commit point.
    var buf: [64]u8 = undefined;
    try std.testing.expectEqualStrings(
        "PRIOR BYTES",
        try std.Io.Dir.cwd().readFile(io, out_path, &buf),
    );
    // Memory: the staged override is still staged — materialized by
    // the attempt (compression precedes the failing rename) and still
    // describing the same replacement bytes. A failed save is not a
    // save that half-consumed the candidate.
    const after = store.overrides[idx].?.compressed;
    try std.testing.expectEqual(std.hash.Crc32.hash("<workbook/>"), after.crc32);
    try std.testing.expectEqual(@as(u32, "<workbook/>".len), after.uncompressed_size);
    try std.testing.expectEqualStrings("<workbook/>", (try store.part("xl/workbook.xml")).?.bytes);
    // …and no `.ztmp-N` beside the user's workbook.
    var it = tmp.dir.iterate();
    while (try it.next(io)) |entry| {
        try std.testing.expect(!std.mem.startsWith(u8, entry.name, ".ztmp-"));
    }
}

test "M5d1: an injected post-commit dir fsync failure is a warning on a committed save" {
    if (@import("builtin").os.tag == .windows) return error.SkipZigTest;

    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    const fixture = "tests/corpus/frictionless_2sheets.xlsx";
    std.Io.Dir.cwd().access(io, fixture, .{}) catch return error.SkipZigTest;

    var tmp = std.testing.tmpDir(.{});
    defer tmp.cleanup();
    const dir = try tmp.dir.realPathFileAlloc(io, ".", std.testing.allocator);
    defer std.testing.allocator.free(dir);
    const out_path = try std.fs.path.join(std.testing.allocator, &.{ dir, "warned.xlsx" });
    defer std.testing.allocator.free(out_path);

    var store = try PartStore.open(std.testing.allocator, io, fixture);
    defer store.deinit();

    // Sync #1 is the temp file (must succeed, or the rename never runs);
    // #2 is the directory, reached only after the commit.
    const faulty = control.inject.wrap(io, .{ .fail_file_sync_at = 2 });
    const commit = try store.saveControlled(faulty, out_path, .none);

    try std.testing.expect(commit.durability_warning);
    try std.testing.expectEqual(@intFromEnum(std.posix.E.IO), commit.durability_errno);
    // Success, not an error — and the file it committed opens.
    var reopened = try PartStore.open(std.testing.allocator, io, out_path);
    defer reopened.deinit();
    try std.testing.expect((try reopened.part("xl/workbook.xml")) != null);
}

test "PartStore.replacePart: unknown part name returns PartNotFound" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    const fixture = "tests/corpus/frictionless_2sheets.xlsx";
    std.Io.Dir.cwd().access(io, fixture, .{}) catch return error.SkipZigTest;

    var store = try PartStore.open(std.testing.allocator, io, fixture);
    defer store.deinit();

    try std.testing.expectError(error.PartNotFound, store.replacePart("xl/does_not_exist.xml", "x"));
}

test "PartStore.open: rejects non-PK file" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tmp = std.testing.tmpDir(.{});
    defer tmp.cleanup();
    try tmp.dir.writeFile(io, .{ .sub_path = "garbage.xlsx", .data = "not a zip" });
    const dir = try tmp.dir.realPathFileAlloc(io, ".", std.testing.allocator);
    defer std.testing.allocator.free(dir);
    const path = try std.fs.path.join(std.testing.allocator, &.{ dir, "garbage.xlsx" });
    defer std.testing.allocator.free(path);

    try std.testing.expectError(Error.NotPkzip, PartStore.open(std.testing.allocator, io, path));
}

test "PartStore.open: rejects split-disk EOCD" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    // Build a minimal 22-byte EOCD with this_disk=1 (split). No CD
    // entries needed — the disk check fires before the CD walk.
    var eocd: [22]u8 = undefined;
    @memset(&eocd, 0);
    std.mem.writeInt(u32, eocd[0..4], 0x06054b50, .little); // EOCD signature
    std.mem.writeInt(u16, eocd[4..6], 1, .little); // this_disk = 1 (split)
    // cd_disk, records_on_disk, total_records, cd_size, cd_offset, comment_len
    // all zero — only the non-zero this_disk should trip the check.

    var tmp = std.testing.tmpDir(.{});
    defer tmp.cleanup();
    try tmp.dir.writeFile(io, .{ .sub_path = "split.xlsx", .data = &eocd });
    const dir = try tmp.dir.realPathFileAlloc(io, ".", std.testing.allocator);
    defer std.testing.allocator.free(dir);
    const path = try std.fs.path.join(std.testing.allocator, &.{ dir, "split.xlsx" });
    defer std.testing.allocator.free(path);

    try std.testing.expectError(
        Error.SplitArchiveNotSupported,
        PartStore.open(std.testing.allocator, io, path),
    );
}

test "decodeXmlEntities decodes the five canonical entities" {
    var arena = std.heap.ArenaAllocator.init(std.testing.allocator);
    defer arena.deinit();
    const a = arena.allocator();

    // Fast path: no entity → verbatim copy.
    try std.testing.expectEqualStrings("hello", try decodeXmlEntities(a, "hello"));

    // The five canonical entities.
    try std.testing.expectEqualStrings("a&b", try decodeXmlEntities(a, "a&amp;b"));
    try std.testing.expectEqualStrings("<", try decodeXmlEntities(a, "&lt;"));
    try std.testing.expectEqualStrings(">", try decodeXmlEntities(a, "&gt;"));
    try std.testing.expectEqualStrings("\"", try decodeXmlEntities(a, "&quot;"));
    try std.testing.expectEqualStrings("'", try decodeXmlEntities(a, "&apos;"));

    // Real-world OOXML rels target with embedded `&`.
    try std.testing.expectEqualStrings(
        "../media/logo&1.png",
        try decodeXmlEntities(a, "../media/logo&amp;1.png"),
    );

    // Unknown entity passes through verbatim.
    try std.testing.expectEqualStrings("&unknown;", try decodeXmlEntities(a, "&unknown;"));

    // Numeric character references (decimal + hex).
    try std.testing.expectEqualStrings("&", try decodeXmlEntities(a, "&#38;"));
    try std.testing.expectEqualStrings("&", try decodeXmlEntities(a, "&#x26;"));
    try std.testing.expectEqualStrings("*", try decodeXmlEntities(a, "&#x2A;"));
    // Multi-byte UTF-8 (€ = U+20AC).
    try std.testing.expectEqualStrings("\u{20AC}", try decodeXmlEntities(a, "&#x20AC;"));
    try std.testing.expectEqualStrings("\u{20AC}", try decodeXmlEntities(a, "&#8364;"));
    // Real-world rels target with numeric ref for `&`.
    try std.testing.expectEqualStrings(
        "../media/logo&1.png",
        try decodeXmlEntities(a, "../media/logo&#38;1.png"),
    );

    // Malformed numeric refs pass through verbatim.
    try std.testing.expectEqualStrings("&#;", try decodeXmlEntities(a, "&#;"));
    try std.testing.expectEqualStrings("&#xZ;", try decodeXmlEntities(a, "&#xZ;"));
    // Out-of-range code point passes through.
    try std.testing.expectEqualStrings("&#1114112;", try decodeXmlEntities(a, "&#1114112;"));

    // XML 1.0 §4.1 forbids non-digit chars in numeric refs.
    // parseInt would otherwise accept `+` / `_`, which is not what
    // the spec allows.
    try std.testing.expectEqualStrings("&#+38;", try decodeXmlEntities(a, "&#+38;"));
    try std.testing.expectEqualStrings("&#3_8;", try decodeXmlEntities(a, "&#3_8;"));
    try std.testing.expectEqualStrings("&#x2_6;", try decodeXmlEntities(a, "&#x2_6;"));
}

test "looksExternal classifies URL / UNC / drive-letter targets" {
    // Real OOXML external relationship targets.
    try std.testing.expect(looksExternal("https://example.com/a.png"));
    try std.testing.expect(looksExternal("http://example.com"));
    try std.testing.expect(looksExternal("mailto:foo@bar.com"));
    try std.testing.expect(looksExternal("file:///etc/hosts"));
    try std.testing.expect(looksExternal("\\\\server\\share\\foo.png"));
    try std.testing.expect(looksExternal("C:\\foo\\bar.png"));
    try std.testing.expect(looksExternal("C:/foo/bar.png"));
    try std.testing.expect(looksExternal("z:/foo")); // lowercase drive

    // Long custom scheme — RFC 3986 doesn't bound scheme length;
    // the heuristic must not either.
    try std.testing.expect(looksExternal("verylongcustomscheme:https://host/file.png"));
    try std.testing.expect(looksExternal("ms-officeapp.test+v1:https://example.com"));

    // In-package targets — must NOT trip the heuristic.
    try std.testing.expect(!looksExternal("worksheets/sheet1.xml"));
    try std.testing.expect(!looksExternal("../sharedStrings.xml"));
    try std.testing.expect(!looksExternal("/xl/workbook.xml"));
    try std.testing.expect(!looksExternal("xl/media/image1.png"));
    try std.testing.expect(!looksExternal("foo.xml"));
    try std.testing.expect(!looksExternal(""));
    try std.testing.expect(!looksExternal("a"));
}

test "PartStore.resolve: external targets return null" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    const fixture = "tests/corpus/frictionless_2sheets.xlsx";
    std.Io.Dir.cwd().access(io, fixture, .{}) catch return error.SkipZigTest;
    var store = try PartStore.open(std.testing.allocator, io, fixture);
    defer store.deinit();

    try std.testing.expectEqual(@as(?[]const u8, null), try store.resolve("xl/worksheets/sheet1.xml", "https://example.com/a.png"));
    try std.testing.expectEqual(@as(?[]const u8, null), try store.resolve("xl/worksheets/sheet1.xml", "mailto:foo@bar.com"));
    try std.testing.expectEqual(@as(?[]const u8, null), try store.resolve("xl/worksheets/sheet1.xml", "C:\\foo.xlsx"));
}

test "attrAtSlice tolerates single-quoted XML attributes" {
    // Content_Types.xml and .rels from non-Microsoft producers
    // (libreoffice, pandoc, hand-edits) sometimes single-quote
    // attributes. Missing them silently dropped content-type
    // resolution and relationship parsing.
    try std.testing.expectEqualStrings(
        "image/png",
        attrAtSlice("Extension=\"png\" ContentType=\"image/png\"", "ContentType").?,
    );
    try std.testing.expectEqualStrings(
        "image/png",
        attrAtSlice("Extension='png' ContentType='image/png'", "ContentType").?,
    );
    try std.testing.expectEqualStrings(
        "rId1",
        attrAtSlice("Id='rId1' Type='foo' Target='bar'", "Id").?,
    );
}

test "decompressPayload: rejects ZIP-bomb declared sizes" {
    var arena = std.heap.ArenaAllocator.init(std.testing.allocator);
    defer arena.deinit();
    const a = arena.allocator();

    // Tiny payload, declared > 512 MiB hard cap → BadZip.
    const tiny = "x";
    try std.testing.expectError(
        Error.BadZip,
        decompressPayload(a, tiny, 8, max_part_size + 1, .none),
    );

    // Tiny payload, declared within hard cap but ratio > 4096:1 → BadZip.
    // tiny is 1 byte so the ratio cap is 4096; declared = 8192 trips it.
    try std.testing.expectError(
        Error.BadZip,
        decompressPayload(a, tiny, 8, 8192, .none),
    );

    // Stored (method 0) entries are still validated against the cap so
    // a CDFH claiming 4 GiB stored can't allocate it.
    try std.testing.expectError(
        Error.BadZip,
        decompressPayload(a, tiny, 0, max_part_size + 1, .none),
    );
}

test "PartStore.addPart + save: new part survives round-trip with content type" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    const fixture = "tests/corpus/frictionless_2sheets.xlsx";
    std.Io.Dir.cwd().access(io, fixture, .{}) catch return error.SkipZigTest;

    var tmp = std.testing.tmpDir(.{});
    defer tmp.cleanup();
    const dir = try tmp.dir.realPathFileAlloc(io, ".", std.testing.allocator);
    defer std.testing.allocator.free(dir);
    const out_path = try std.fs.path.join(std.testing.allocator, &.{ dir, "added.xlsx" });
    defer std.testing.allocator.free(out_path);

    const new_part_name = "xl/customData.xml";
    const new_part_ct = "application/xml";
    const new_part_bytes = "<?xml version=\"1.0\"?><custom>added</custom>";

    {
        var store = try PartStore.open(std.testing.allocator, io, fixture);
        defer store.deinit();
        try store.addPart(new_part_name, new_part_ct, new_part_bytes);
        try store.save(io, out_path);
    }

    var dst = try PartStore.open(std.testing.allocator, io, out_path);
    defer dst.deinit();

    const part_in_dst = try dst.part(new_part_name) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqualSlices(u8, new_part_bytes, part_in_dst.bytes);
    try std.testing.expectEqualStrings(new_part_ct, part_in_dst.content_type.?);
}

test "PartStore.addPart: multiple parts in one session all register content types" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    const fixture = "tests/corpus/frictionless_2sheets.xlsx";
    std.Io.Dir.cwd().access(io, fixture, .{}) catch return error.SkipZigTest;

    var tmp = std.testing.tmpDir(.{});
    defer tmp.cleanup();
    const dir = try tmp.dir.realPathFileAlloc(io, ".", std.testing.allocator);
    defer std.testing.allocator.free(dir);
    const out_path = try std.fs.path.join(std.testing.allocator, &.{ dir, "added_two.xlsx" });
    defer std.testing.allocator.free(out_path);

    {
        var store = try PartStore.open(std.testing.allocator, io, fixture);
        defer store.deinit();
        try store.addPart("xl/customA.xml", "application/xml", "<a/>");
        try store.addPart("xl/customB.xml", "application/xml", "<b/>");
        try store.save(io, out_path);
    }

    var dst = try PartStore.open(std.testing.allocator, io, out_path);
    defer dst.deinit();

    // Both parts present.
    const a = try dst.part("xl/customA.xml") orelse return error.TestUnexpectedResult;
    const b = try dst.part("xl/customB.xml") orelse return error.TestUnexpectedResult;
    try std.testing.expectEqualSlices(u8, "<a/>", a.bytes);
    try std.testing.expectEqualSlices(u8, "<b/>", b.bytes);
    // Both content types registered (the bug Codex caught was that
    // the second addPart's content-type rebuild lost the first
    // addPart's <Override>, leaving customA.xml undeclared).
    try std.testing.expectEqualStrings("application/xml", a.content_type.?);
    try std.testing.expectEqualStrings("application/xml", b.content_type.?);
}

test "PartStore.addPart: XML-escapes part name + content type into Content_Types" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    const fixture = "tests/corpus/frictionless_2sheets.xlsx";
    std.Io.Dir.cwd().access(io, fixture, .{}) catch return error.SkipZigTest;

    var tmp = std.testing.tmpDir(.{});
    defer tmp.cleanup();
    const dir = try tmp.dir.realPathFileAlloc(io, ".", std.testing.allocator);
    defer std.testing.allocator.free(dir);
    const out_path = try std.fs.path.join(std.testing.allocator, &.{ dir, "added_meta.xlsx" });
    defer std.testing.allocator.free(out_path);

    // ZIP names CAN contain `&` (rare but legal). The on-disk part
    // name is the raw `&`, but the [Content_Types].xml entry must
    // serialise it as `&amp;` to keep the XML well-formed.
    const tricky_name = "xl/a&b.xml";
    {
        var store = try PartStore.open(std.testing.allocator, io, fixture);
        defer store.deinit();
        try store.addPart(tricky_name, "application/xml", "<x/>");
        try store.save(io, out_path);
    }

    var dst = try PartStore.open(std.testing.allocator, io, out_path);
    defer dst.deinit();
    // Reopened part is found under its raw name (the .rels parser
    // decodes entities to recover the literal).
    const got = try dst.part(tricky_name) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqualSlices(u8, "<x/>", got.bytes);
    // Content_Types.xml must contain the escaped form, not the raw `&`.
    const ct = try dst.part("[Content_Types].xml") orelse return error.TestUnexpectedResult;
    try std.testing.expect(std.mem.indexOf(u8, ct.bytes, "PartName=\"/xl/a&amp;b.xml\"") != null);
    // ...and must NOT contain the raw `&` in an attribute (otherwise
    // the XML is malformed).
    try std.testing.expect(std.mem.indexOf(u8, ct.bytes, "PartName=\"/xl/a&b.xml\"") == null);
    // Round-trip: the part's content_type field must resolve too —
    // resolveContentTypes decodes the escaped PartName back to the
    // raw form so the lookup against the ZIP entry name succeeds.
    try std.testing.expectEqualStrings("application/xml", got.content_type.?);
}

test "PartStore.addPart: rejects duplicate part name" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    const fixture = "tests/corpus/frictionless_2sheets.xlsx";
    std.Io.Dir.cwd().access(io, fixture, .{}) catch return error.SkipZigTest;

    var store = try PartStore.open(std.testing.allocator, io, fixture);
    defer store.deinit();

    try std.testing.expectError(
        error.PartAlreadyExists,
        store.addPart("xl/workbook.xml", "application/xml", "irrelevant"),
    );
}

// Fuzz targets for attacker-controlled parsers. These run as plain
// smoke tests on `zig build test` (a few iterations on the seed-only
// corpus) and become coverage-guided under `zig build fuzz` once
// the package-layer test target is wired into the fuzz module.
// Contract: the parser must not panic / deadlock / OOB-read on
// any input — typed errors are fine.

fn fuzzDecodeXmlEntitiesTarget(_: void, smith: *std.testing.Smith) anyerror!void {
    @disableInstrumentation();
    // 4 KiB matches the PRNG harness's scratch bound, so a crash
    // found here reproduces against the same input shape.
    var smith_buf: [4096]u8 = undefined;
    const input = smith_buf[0..smith.slice(&smith_buf)];
    var arena = std.heap.ArenaAllocator.init(std.testing.allocator);
    defer arena.deinit();
    const decoded = decodeXmlEntities(arena.allocator(), input) catch return;
    // Output must be valid UTF-8 — numeric refs go through
    // utf8Encode which rejects invalid code points.
    _ = std.unicode.utf8ValidateSlice(decoded);
}

test "fuzz: decodeXmlEntities never crashes on adversarial input" {
    try std.testing.fuzz({}, fuzzDecodeXmlEntitiesTarget, .{
        .corpus = &[_][]const u8{
            "&amp;",                     "&lt;",       "&#38;",      "&#x26;",
            "&unknown;",                 "&#",         "&#;",        "&#xZ;",
            "&#9999999999999;",          "&#x10FFFF;", "&#x110000;", "&#0;",
            "&amp;&lt;&gt;&quot;&apos;", "a&b",        "&&&&",
            "\xC0\x80", "\xFF\xFE", // invalid UTF-8 inputs — must not crash
        },
    });
}

fn fuzzLooksExternalTarget(_: void, smith: *std.testing.Smith) anyerror!void {
    @disableInstrumentation();
    // 4 KiB matches the PRNG harness's scratch bound, so a crash
    // found here reproduces against the same input shape.
    var smith_buf: [4096]u8 = undefined;
    const input = smith_buf[0..smith.slice(&smith_buf)];
    _ = looksExternal(input);
}

fn fuzzParseRelationshipsTarget(_: void, smith: *std.testing.Smith) anyerror!void {
    @disableInstrumentation();
    // 4 KiB matches the PRNG harness's scratch bound, so a crash
    // found here reproduces against the same input shape.
    var smith_buf: [4096]u8 = undefined;
    const input = smith_buf[0..smith.slice(&smith_buf)];
    var arena = std.heap.ArenaAllocator.init(std.testing.allocator);
    defer arena.deinit();
    _ = parseRelationships(arena.allocator(), input) catch {};
}

test "fuzz: parseRelationships never crashes on adversarial XML" {
    try std.testing.fuzz({}, fuzzParseRelationshipsTarget, .{
        .corpus = &[_][]const u8{
            "",
            "<?xml version=\"1.0\"?><Relationships/>",
            "<Relationships><Relationship Id='rId1' Type='img' Target='x.png'/></Relationships>",
            "<Relationship Id=\"a\" Target=\"&amp;\"/>",
            "<Relationship Id=\"a\" Target=\"&\"/>", // unclosed entity
            "<Relationship Id=\"a\" Target=\"&#1114112;\"/>", // out-of-range numeric
            "<Relationship Id=\"a\" Target=\"\"/>",
            "<Relationship Id=\"a\" TargetMode=\"External\" Target=\"https://example.com\"/>",
            "<Relationship", // truncated
            "<Relationship>", // no attrs
            "<<<<<<<<<<<<<<<<<<<<<<<<", // pathological
            "<Relationship Id='\xC0\x80'/>", // overlong UTF-8
        },
    });
}

test "fuzz: looksExternal never crashes on adversarial input" {
    try std.testing.fuzz({}, fuzzLooksExternalTarget, .{
        .corpus = &[_][]const u8{
            "",                                      "a",                      "/",
            "https://example.com",                   "C:\\foo",                "\\\\unc",
            ":colon-only",                           "scheme:",                "xmlns:foo",
            "12scheme:",                             "+:",                     "scheme++:",
            "verylongschemenamethatexceedssixteen:", "../../../../etc/passwd", "a:/b",
        },
    });
}

test "PartStore.addPart: large input round-trips through deflate" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    const fixture = "tests/corpus/frictionless_2sheets.xlsx";
    std.Io.Dir.cwd().access(io, fixture, .{}) catch return error.SkipZigTest;

    var tmp = std.testing.tmpDir(.{});
    defer tmp.cleanup();
    const dir = try tmp.dir.realPathFileAlloc(io, ".", std.testing.allocator);
    defer std.testing.allocator.free(dir);
    const out_path = try std.fs.path.join(std.testing.allocator, &.{ dir, "added_large.xlsx" });
    defer std.testing.allocator.free(out_path);

    // 16 KiB highly-redundant payload — well above the 1 KiB
    // STORED-vs-DEFLATE threshold, so the deflate path is exercised.
    const big_bytes = try std.testing.allocator.alloc(u8, 16 * 1024);
    defer std.testing.allocator.free(big_bytes);
    @memset(big_bytes, 'A');

    {
        var store = try PartStore.open(std.testing.allocator, io, fixture);
        defer store.deinit();
        try store.addPart("xl/extra.bin", "application/octet-stream", big_bytes);
        try store.save(io, out_path);
    }

    var dst = try PartStore.open(std.testing.allocator, io, out_path);
    defer dst.deinit();
    const got = try dst.part("xl/extra.bin") orelse return error.TestUnexpectedResult;
    try std.testing.expectEqualSlices(u8, big_bytes, got.bytes);
}

test "PartStore.addPart: atomic on every allocation-failure step" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    const fixture = "tests/corpus/frictionless_2sheets.xlsx";
    std.Io.Dir.cwd().access(io, fixture, .{}) catch return error.SkipZigTest;

    // Drive addPart through std.testing.checkAllAllocationFailures
    // so every single fallible allocation along the way takes a
    // turn returning OOM. The contract: on any error, the store's
    // observable state (parts.len, the [Content_Types].xml bytes
    // a fresh part() call returns) is unchanged from before the
    // call. The atomicity rebuild — alloc-then-commit, parallel
    // arrays grown by alloc+copy not realloc, staged CT update —
    // is exactly what this verifies.
    const Closure = struct {
        fn run(alloc: std.mem.Allocator, run_io: std.Io, src_fixture: []const u8) !void {
            // The store itself is opened under the failing
            // allocator too — checkAllAllocationFailures has its
            // own contract: every OOM either propagates as-is or
            // is converted to a different error. open() failing
            // is fine; we just need to propagate it.
            var store = try PartStore.open(alloc, run_io, src_fixture);
            defer store.deinit();

            const before_count = store.parts.len;
            const ct_before = try store.part("[Content_Types].xml") orelse
                return error.MissingContentTypes;
            const ct_before_bytes = ct_before.bytes;

            // The actual call we're stressing.
            store.addPart("xl/extra.xml", "application/xml", "<x/>") catch |e| {
                // On failure, store state must be unchanged.
                try std.testing.expectEqual(before_count, store.parts.len);
                const ct_after = try store.part("[Content_Types].xml") orelse return e;
                try std.testing.expect(ct_after.bytes.ptr == ct_before_bytes.ptr);
                return e;
            };
            // On success the store grew by exactly one part.
            try std.testing.expectEqual(before_count + 1, store.parts.len);
        }
    };

    try std.testing.checkAllAllocationFailures(
        std.testing.allocator,
        Closure.run,
        // io travels through the extra-args tuple: an inner struct
        // fn cannot capture it from the enclosing test scope.
        .{ io, fixture },
    );
}

test "PartStore.fresh: returns empty store" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var store = try PartStore.fresh(std.testing.allocator, io);
    defer store.deinit();

    // A fresh store's backing is the empty buffer, not an absent one:
    // the field is non-optional so no read path needs a null check.
    try std.testing.expect(store.backing.source == .buffer);
    try std.testing.expectEqual(@as(usize, 0), store.backing.source.buffer.len);
    try std.testing.expectEqual(@as(usize, 1), store.backing.refCount());
    // No EOCD comment.
    try std.testing.expectEqual(@as(usize, 0), store.eocd_comment.len);
    // No relationships seeded.
    try std.testing.expectEqual(@as(u32, 0), store.rels_by_owner.count());

    // The store is seeded with `[Content_Types].xml` so subsequent
    // `addPart` calls have a CT document to stage Override entries
    // into. The seed is held as an override too, which is what makes
    // the all-overrides save path fire on save().
    const names = try store.partNames();
    try std.testing.expectEqual(@as(usize, 1), names.len);
    try std.testing.expectEqualStrings("[Content_Types].xml", names[0]);
    try std.testing.expect(store.overrides.len == 1);
    try std.testing.expect(store.overrides[0] != null);

    // Allocator is wired through (arena child of testing.allocator).
    // No way to assert that directly; testing.allocator's leak check
    // at end-of-test catches any orphan allocation.
}

test "PartStore.fresh: addPart populates the store" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var store = try PartStore.fresh(std.testing.allocator, io);
    defer store.deinit();

    try store.addPart("hello.txt", "text/plain", "hello world");

    const p = (try store.part("hello.txt")) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqualSlices(u8, "hello world", p.bytes);
    try std.testing.expectEqualStrings("text/plain", p.content_type.?);

    // [Content_Types].xml must now declare the new part via an
    // <Override> — same contract as addPart on a non-fresh store.
    const ct = (try store.part("[Content_Types].xml")) orelse
        return error.TestUnexpectedResult;
    try std.testing.expect(std.mem.indexOf(u8, ct.bytes, "/hello.txt") != null);
    try std.testing.expect(std.mem.indexOf(u8, ct.bytes, "text/plain") != null);
}

test "PartStore.fresh: save round-trips through PartStore.open" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tmp = std.testing.tmpDir(.{});
    defer tmp.cleanup();
    const dir = try tmp.dir.realPathFileAlloc(io, ".", std.testing.allocator);
    defer std.testing.allocator.free(dir);
    const out_path = try std.fs.path.join(std.testing.allocator, &.{ dir, "fresh_round.xlsx" });
    defer std.testing.allocator.free(out_path);

    // Two parts: one tiny (forced STORED by the <1 KiB heuristic),
    // one >= 1 KiB (deflated when it shrinks).
    const small_name = "tiny.xml";
    const small_bytes = "<x/>";
    const big_name = "xl/big.xml";
    var big_buf: [2048]u8 = undefined;
    // Highly compressible payload so deflate definitely wins.
    @memset(&big_buf, 'A');
    const big_bytes: []const u8 = &big_buf;

    {
        var store = try PartStore.fresh(std.testing.allocator, io);
        defer store.deinit();
        try store.addPart(small_name, "application/xml", small_bytes);
        try store.addPart(big_name, "application/xml", big_bytes);
        try store.save(io, out_path);
    }

    var dst = try PartStore.open(std.testing.allocator, io, out_path);
    defer dst.deinit();

    const small = (try dst.part(small_name)) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqualSlices(u8, small_bytes, small.bytes);
    try std.testing.expectEqualStrings("application/xml", small.content_type.?);
    // Tiny input: STORED.
    try std.testing.expectEqual(@as(u16, 0), small.compression_method);

    const big = (try dst.part(big_name)) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqualSlices(u8, big_bytes, big.bytes);
    try std.testing.expectEqualStrings("application/xml", big.content_type.?);
    // 2 KiB of 'A' compresses to a few bytes — deflate should win.
    try std.testing.expectEqual(@as(u16, 8), big.compression_method);

    // Three entries on disk: seeded CT.xml + the two added parts.
    const dst_names = try dst.partNames();
    try std.testing.expectEqual(@as(usize, 3), dst_names.len);
}

test "PartStore.fresh: save with zero addParts produces a valid 1-entry ZIP" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    // Strict "0-entry ZIP" isn't reachable through `fresh()` because
    // the constructor seeds `[Content_Types].xml` (precondition for
    // addPart). The closest contract is: `fresh()` immediately
    // followed by `save()` produces a 1-entry ZIP holding only the
    // seeded CT.xml. That archive must round-trip cleanly through
    // `open()`.
    var tmp = std.testing.tmpDir(.{});
    defer tmp.cleanup();
    const dir = try tmp.dir.realPathFileAlloc(io, ".", std.testing.allocator);
    defer std.testing.allocator.free(dir);
    const out_path = try std.fs.path.join(std.testing.allocator, &.{ dir, "fresh_empty.xlsx" });
    defer std.testing.allocator.free(out_path);

    {
        var store = try PartStore.fresh(std.testing.allocator, io);
        defer store.deinit();
        try store.save(io, out_path);
    }

    var dst = try PartStore.open(std.testing.allocator, io, out_path);
    defer dst.deinit();

    const dst_names = try dst.partNames();
    try std.testing.expectEqual(@as(usize, 1), dst_names.len);
    try std.testing.expectEqualStrings("[Content_Types].xml", dst_names[0]);

    const ct = (try dst.part("[Content_Types].xml")) orelse
        return error.TestUnexpectedResult;
    try std.testing.expect(std.mem.indexOf(u8, ct.bytes, "<Types") != null);
    try std.testing.expect(std.mem.indexOf(u8, ct.bytes, "</Types>") != null);
}

// ─── M5b0: SourceBacking ownership ───────────────────────────────────
//
// One sentence, tested from four directions: N generations may share
// one backing, each stays independently readable for as long as it
// lives, they may be retired in any order, and the source is closed
// exactly once.

/// Open a backing whose close is reported through `ledger`. Mirrors the
/// adopt-then-never-arm-a-second-closer shape of `PartStore.open`; the
/// ledger is the only difference.
fn openLedgeredBacking(
    allocator: std.mem.Allocator,
    io: std.Io,
    path: []const u8,
    ledger: *CloseLedger,
) Error!*SourceBacking {
    const file = try std.Io.Dir.cwd().openFile(io, path, .{});
    return SourceBacking.createFile(allocator, io, file, ledger) catch |err| {
        file.close(io);
        return err;
    };
}

fn openLedgered(
    allocator: std.mem.Allocator,
    io: std.Io,
    path: []const u8,
    ledger: *CloseLedger,
) Error!PartStore {
    const backing = try openLedgeredBacking(allocator, io, path, ledger);
    errdefer backing.release();
    return try PartStore.openOver(allocator, backing, .none);
}

/// Local-file-header signature. Any live backing over a real archive
/// must still be able to hand these four bytes back.
const pk_lfh_magic = "PK\x03\x04";

/// Close-order permutations over four generations — §5.7.4's
/// `max_retained_generations` default, so this is the shape the
/// transaction will actually produce.
const close_orders = [_][4]usize{
    .{ 0, 1, 2, 3 }, // forward: the generation that opened the file goes first
    .{ 3, 2, 1, 0 }, // reverse
    .{ 1, 3, 0, 2 }, // interleaved
};

fn exerciseCloseOrder(io: std.Io, fixture: []const u8, order: [4]usize) !void {
    const allocator = std.testing.allocator;
    var ledger: CloseLedger = .{};

    var gens: [order.len]PartStore = undefined;
    gens[0] = try openLedgered(allocator, io, fixture, &ledger);
    for (gens[1..], 1..) |*g, i| g.* = try gens[i - 1].nextGeneration();

    // One backing, four references: the fd budget is one for the whole
    // retained set, not one per generation.
    try std.testing.expectEqual(gens.len, gens[0].backing.refCount());
    for (gens[1..]) |g| try std.testing.expect(g.backing == gens[0].backing);

    // Each generation caches into its own arena, so equal bytes here
    // are four independent reads of one source, not a shared slice.
    var first_bytes: []const u8 = &.{};
    for (&gens, 0..) |*g, i| {
        const p = (try g.part("xl/workbook.xml")) orelse
            return error.TestUnexpectedResult;
        try std.testing.expect(p.bytes.len > 0);
        if (i == 0) {
            first_bytes = p.bytes;
        } else {
            try std.testing.expectEqualSlices(u8, first_bytes, p.bytes);
            try std.testing.expect(first_bytes.ptr != p.bytes.ptr);
        }
    }
    try std.testing.expectEqual(@as(usize, 0), ledger.closes);

    for (order, 0..) |idx, round| {
        gens[idx].deinit();
        const survivors = order.len - round - 1;
        const last = survivors == 0;
        try std.testing.expectEqual(@as(usize, if (last) 1 else 0), ledger.closes);
        if (!last) {
            try std.testing.expectEqual(survivors, gens[order[round + 1]].backing.refCount());
        }

        // Every survivor can still reach the source — both raw, and
        // through the lazy-materialization path with a part name no
        // generation has touched yet this round.
        for (order[round + 1 ..]) |alive| {
            var probe: [4]u8 = undefined;
            try gens[alive].backing.readAt(0, &probe);
            try std.testing.expectEqualSlices(u8, pk_lfh_magic, &probe);

            const names = try gens[alive].partNames();
            const want = names[round % names.len];
            const p = (try gens[alive].part(want)) orelse
                return error.TestUnexpectedResult;
            try std.testing.expectEqualStrings(want, p.name);
        }
    }
    try std.testing.expectEqual(@as(usize, 1), ledger.closes);
}

test "SourceBacking: four generations, closed in every order, close once" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    const fixture = "tests/corpus/frictionless_2sheets.xlsx";
    std.Io.Dir.cwd().access(io, fixture, .{}) catch return error.SkipZigTest;

    for (close_orders) |order| try exerciseCloseOrder(io, fixture, order);
}

/// Drive a retain/release schedule over one backing and prove the close
/// lands exactly once, at the end. Bit 0 of each op byte picks retain
/// vs release; a release that would drop the last reference ends the
/// schedule instead, because reading a destroyed backing is the
/// caller's bug, not the backing's.
///
/// Shared by the fuzz target (buffer backing) and the seeded test (file
/// backing) so the two variants are held to one refcount discipline.
fn exerciseRefcountSchedule(
    backing: *SourceBacking,
    ledger: *CloseLedger,
    ops: []const u8,
    expect_readable: []const u8,
) !void {
    var live: usize = 1; // the reference `create` handed us
    for (ops) |op| {
        if (op & 1 == 0) {
            _ = backing.retain();
            live += 1;
        } else {
            if (live == 1) break;
            backing.release();
            live -= 1;
        }
        try std.testing.expectEqual(live, backing.refCount());
        try std.testing.expectEqual(@as(usize, 0), ledger.closes);

        // Whatever sequence of retains and releases got us here, a
        // backing with references outstanding is still readable.
        var probe: [4]u8 = undefined;
        try backing.readAt(0, &probe);
        try std.testing.expectEqualSlices(u8, expect_readable, &probe);
    }
    while (live > 0) : (live -= 1) backing.release();
    try std.testing.expectEqual(@as(usize, 1), ledger.closes);
}

fn fuzzBackingRefcountTarget(io: std.Io, smith: *std.testing.Smith) anyerror!void {
    @disableInstrumentation();
    // 256 ops is well past the depth any real transaction reaches; the
    // schedule, not the archive, is what varies here.
    var smith_buf: [256]u8 = undefined;
    const ops = smith_buf[0..smith.slice(&smith_buf)];

    var ledger: CloseLedger = .{};
    const backing = try SourceBacking.createBuffer(
        std.testing.allocator,
        io,
        pk_lfh_magic,
        &ledger,
    );
    try exerciseRefcountSchedule(backing, &ledger, ops, pk_lfh_magic);
}

test "fuzz: SourceBacking never double-closes under randomized clone/drop" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    try std.testing.fuzz(threaded.io(), fuzzBackingRefcountTarget, .{
        .corpus = &[_][]const u8{
            "", // no ops: create then release
            "\x01", // immediate last-release
            "\x00\x01", // retain, release
            "\x00\x00", // two retains, both dropped by the tail loop
            "\x01\x01\x01", // releases that must not underflow
            "\x00\x01\x00\x01\x00\x01", // alternating
            "\x00\x00\x00\x00\x01\x01\x01\x01", // fan out, fan in
            "\x00\x00\x01\x00\x01\x01\x00\x01", // interleaved
        },
    });
}

test "SourceBacking: seeded clone/drop schedules over a file backing" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    const fixture = "tests/corpus/frictionless_2sheets.xlsx";
    std.Io.Dir.cwd().access(io, fixture, .{}) catch return error.SkipZigTest;

    // The fuzz target above runs the buffer variant, where "close" is a
    // free. This runs the same schedule where "close" is an fd close —
    // the case where getting it wrong is not merely a leak.
    var prng = std.Random.DefaultPrng.init(0x5b0);
    for (0..64) |_| {
        var ops: [32]u8 = undefined;
        prng.random().bytes(&ops);
        var ledger: CloseLedger = .{};
        const backing = try openLedgeredBacking(
            std.testing.allocator,
            io,
            fixture,
            &ledger,
        );
        try exerciseRefcountSchedule(backing, &ledger, &ops, pk_lfh_magic);
    }
}

test "PartStore: generation N+1 reads bytes generation N still holds" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    const fixture = "tests/corpus/frictionless_2sheets.xlsx";
    std.Io.Dir.cwd().access(io, fixture, .{}) catch return error.SkipZigTest;

    const allocator = std.testing.allocator;
    var ledger: CloseLedger = .{};
    var gen0 = try openLedgered(allocator, io, fixture, &ledger);

    const wb0 = (try gen0.part("xl/workbook.xml")) orelse
        return error.TestUnexpectedResult;
    const gen0_bytes = wb0.bytes;
    try std.testing.expect(gen0_bytes.len > 0);

    var gen1 = try gen0.nextGeneration();
    const wb1 = (try gen1.part("xl/workbook.xml")) orelse
        return error.TestUnexpectedResult;
    try std.testing.expectEqualSlices(u8, gen0_bytes, wb1.bytes);
    try std.testing.expect(gen0_bytes.ptr != wb1.bytes.ptr);

    // Retire the generation that opened the file, first. Under the old
    // exclusive ownership this closed the fd out from under gen1 and
    // every later materialization failed.
    gen0.deinit();
    try std.testing.expectEqual(@as(usize, 0), ledger.closes);
    try std.testing.expectEqual(@as(usize, 1), gen1.backing.refCount());

    // gen1 can still lazily materialize a part nobody has touched...
    const styles = (try gen1.part("xl/styles.xml")) orelse
        return error.TestUnexpectedResult;
    try std.testing.expect(styles.bytes.len > 0);

    // ...and still spawn generation 2 from the same source, which is
    // what makes recalc repeatable rather than once-per-open.
    var gen2 = try gen1.nextGeneration();
    const wb2 = (try gen2.part("xl/workbook.xml")) orelse
        return error.TestUnexpectedResult;
    try std.testing.expectEqualSlices(u8, wb1.bytes, wb2.bytes);

    gen1.deinit();
    try std.testing.expectEqual(@as(usize, 0), ledger.closes);
    const styles2 = (try gen2.part("xl/styles.xml")) orelse
        return error.TestUnexpectedResult;
    try std.testing.expect(styles2.bytes.len > 0);

    gen2.deinit();
    try std.testing.expectEqual(@as(usize, 1), ledger.closes);
}

test "PartStore.nextGeneration: a fresh store has no source to re-scan" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();

    var store = try PartStore.fresh(std.testing.allocator, io);
    defer store.deinit();

    // Documented, not accidental: a generation is the source archive
    // re-scanned, and a fresh store's parts live in overrides. The
    // empty backing has no EOCD, so the scan refuses.
    try std.testing.expectError(error.NotPkzip, store.nextGeneration());
    // The failed attempt released its own reference and nothing else.
    try std.testing.expectEqual(@as(usize, 1), store.backing.refCount());
}

test "PartStore: buffer-backed and file-backed stores take the same path" {
    comptime {
        // Two variants, both exercised below. A third would make this
        // test's claim to cover "both paths" false, so break the build
        // rather than quietly under-test.
        std.debug.assert(@typeInfo(SourceBacking.Source).@"union".fields.len == 2);
    }

    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    const fixture = "tests/corpus/frictionless_2sheets.xlsx";
    std.Io.Dir.cwd().access(io, fixture, .{}) catch return error.SkipZigTest;
    const allocator = std.testing.allocator;

    var from_file = try PartStore.open(allocator, io, fixture);
    defer from_file.deinit();

    var from_buf = blk: {
        const raw = try std.Io.Dir.cwd().readFileAlloc(io, fixture, allocator, .limited(1 << 24));
        defer allocator.free(raw);
        const s = try PartStore.openBuffer(allocator, io, raw);
        // Poison the caller's slice before a single part is read: the
        // backing dupes, so nothing below may depend on these bytes.
        @memset(raw, 0xAA);
        break :blk s;
    };
    defer from_buf.deinit();

    // The two stores differ in exactly one observable.
    try std.testing.expect(from_file.backing.source == .file);
    try std.testing.expect(from_buf.backing.source == .buffer);

    // Identical scan.
    try std.testing.expectEqual(from_file.entries.len, from_buf.entries.len);
    for (from_file.entries, from_buf.entries) |a, b| {
        try std.testing.expectEqualStrings(a.name, b.name);
        try std.testing.expectEqual(a.lfh_offset, b.lfh_offset);
        try std.testing.expectEqual(a.lfh_total_len, b.lfh_total_len);
        try std.testing.expectEqual(a.cdfh_offset, b.cdfh_offset);
        try std.testing.expectEqual(a.cdfh_total_len, b.cdfh_total_len);
        try std.testing.expectEqual(a.payload_offset, b.payload_offset);
        try std.testing.expectEqual(a.compressed_size, b.compressed_size);
        try std.testing.expectEqual(a.uncompressed_size, b.uncompressed_size);
        try std.testing.expectEqual(a.compression_method, b.compression_method);
        try std.testing.expectEqual(a.crc32, b.crc32);
        try std.testing.expectEqual(a.data_descriptor_len, b.data_descriptor_len);
        try std.testing.expectEqual(a.has_data_descriptor, b.has_data_descriptor);
    }

    // Identical lazy materialization, part by part — this is the
    // `backing.readAt` path in `materializeAt`, once per variant.
    const names_file = try from_file.partNames();
    const names_buf = try from_buf.partNames();
    try std.testing.expectEqual(names_file.len, names_buf.len);
    for (names_file, names_buf) |nf, nb| {
        try std.testing.expectEqualStrings(nf, nb);
        const pf = (try from_file.part(nf)) orelse return error.TestUnexpectedResult;
        const pb = (try from_buf.part(nb)) orelse return error.TestUnexpectedResult;
        try std.testing.expectEqualSlices(u8, pf.bytes, pb.bytes);
        try std.testing.expectEqual(pf.compression_method, pb.compression_method);
        if (pf.content_type) |ctf| {
            try std.testing.expectEqualStrings(ctf, pb.content_type.?);
        } else {
            try std.testing.expect(pb.content_type == null);
        }
        // Identical relationships for every owner that has any.
        const rf = from_file.rels(nf);
        const rb = from_buf.rels(nb);
        try std.testing.expectEqual(rf.len, rb.len);
        for (rf, rb) |x, y| {
            try std.testing.expectEqualStrings(x.id, y.id);
            try std.testing.expectEqualStrings(x.type, y.type);
            try std.testing.expectEqualStrings(x.target, y.target);
            try std.testing.expectEqual(x.target_mode, y.target_mode);
        }
    }
    try std.testing.expectEqualSlices(u8, from_file.eocd_comment, from_buf.eocd_comment);

    // Identical output. `save` copies untouched LFH and CDFH bytes
    // straight out of the backing, so a byte-equal archive from both
    // is the strongest statement available that the two variants are
    // one path: the ONLY difference between these stores is which arm
    // of `Source` answered every one of those reads.
    var tmp = std.testing.tmpDir(.{});
    defer tmp.cleanup();
    const dir = try tmp.dir.realPathFileAlloc(io, ".", allocator);
    defer allocator.free(dir);
    const out_file = try std.fs.path.join(allocator, &.{ dir, "from_file.xlsx" });
    defer allocator.free(out_file);
    const out_buf = try std.fs.path.join(allocator, &.{ dir, "from_buf.xlsx" });
    defer allocator.free(out_buf);

    try from_file.save(io, out_file);
    try from_buf.save(io, out_buf);

    const saved_file = try std.Io.Dir.cwd().readFileAlloc(io, out_file, allocator, .limited(1 << 24));
    defer allocator.free(saved_file);
    const saved_buf = try std.Io.Dir.cwd().readFileAlloc(io, out_buf, allocator, .limited(1 << 24));
    defer allocator.free(saved_buf);
    try std.testing.expectEqualSlices(u8, saved_file, saved_buf);
}

test "PartStore.openBuffer: generations share one buffer backing" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    const fixture = "tests/corpus/frictionless_2sheets.xlsx";
    std.Io.Dir.cwd().access(io, fixture, .{}) catch return error.SkipZigTest;
    const allocator = std.testing.allocator;

    const raw = try std.Io.Dir.cwd().readFileAlloc(io, fixture, allocator, .limited(1 << 24));
    defer allocator.free(raw);

    var gen0 = try PartStore.openBuffer(allocator, io, raw);
    var gen1 = try gen0.nextGeneration();
    try std.testing.expect(gen0.backing == gen1.backing);
    try std.testing.expectEqual(@as(usize, 2), gen0.backing.refCount());

    // A generation over a buffer retains bytes the same way one over a
    // file retains an fd: gen0 goes first, gen1 still materializes.
    gen0.deinit();
    const styles = (try gen1.part("xl/styles.xml")) orelse
        return error.TestUnexpectedResult;
    try std.testing.expect(styles.bytes.len > 0);
    gen1.deinit();
}

test "xmlAttrValue: lexes quotes, Eq whitespace, and hostile values" {
    // Both quote styles and XML 1.0 §3.1 Eq whitespace.
    try std.testing.expectEqualStrings("rId1", xmlAttrValue("<Relationship Id=\"rId1\"", "Id").?);
    try std.testing.expectEqualStrings("rId1", xmlAttrValue("<Relationship Id='rId1'", "Id").?);
    try std.testing.expectEqualStrings("rId1", xmlAttrValue("<Relationship Id = 'rId1'", "Id").?);
    try std.testing.expectEqualStrings("rId1", xmlAttrValue("<Relationship\n  Id\t= \"rId1\"", "Id").?);

    // A value containing `=` or a lookalike key never satisfies nor
    // derails the lex — attributes are walked, not substring-matched.
    try std.testing.expectEqualStrings("y", xmlAttrValue("<R Target=\"a=b\" Id=\"y\"", "Id").?);
    try std.testing.expectEqualStrings("y", xmlAttrValue("<R MyId=\"x\" Id=\"y\"", "Id").?);
    try std.testing.expect(xmlAttrValue("<R MyId=\"x\"", "Id") == null);
    // Attribute order is free.
    try std.testing.expectEqualStrings("t", xmlAttrValue("<R Target='t' Id='i'", "Target").?);
    // Missing attribute, malformed Eq.
    try std.testing.expect(xmlAttrValue("<R Id=\"x\"", "Type") == null);
    try std.testing.expect(xmlAttrValue("<R Id=rId1", "Id") == null);
}

test "xmlStartTagEnd: skips '>' inside quoted values" {
    const xml = "<xdr:cNvPr name=\"a>b\" id=\"1\"/><next/>";
    const end = xmlStartTagEnd(xml, 0).?;
    // The real close is the one after id="1"/, not the > inside name.
    try std.testing.expectEqual(@as(u8, '>'), xml[end]);
    try std.testing.expect(std.mem.startsWith(u8, xml[end + 1 ..], "<next/>"));
    // An unterminated tag is null, not a scan past the buffer.
    try std.testing.expect(xmlStartTagEnd("<a href=\"x", 0) == null);
}

test "parseRelationships: accepts Eq whitespace around = (Codex r2)" {
    var arena = std.heap.ArenaAllocator.init(std.testing.allocator);
    defer arena.deinit();
    const rels = try parseRelationships(
        arena.allocator(),
        "<Relationships><Relationship Id = 'rId1' Type = 'x' Target = '../drawings/drawing1.xml'/></Relationships>",
    );
    try std.testing.expectEqual(@as(usize, 1), rels.len);
    try std.testing.expectEqualStrings("rId1", rels[0].id);
    try std.testing.expectEqualStrings("../drawings/drawing1.xml", rels[0].target);
}

test "xmlAttrValue: representative legal attributes across repo parsers (Codex r3)" {
    const rel = "<Relationship xml:space=\"preserve\" Id=\"rId1\" Target=\"O'Brien=a>b\" TargetMode=\"External\"";
    try std.testing.expectEqualStrings("rId1", xmlAttrValue(rel, "Id").?);
    try std.testing.expectEqualStrings("O'Brien=a>b", xmlAttrValue(rel, "Target").?);
    try std.testing.expectEqualStrings("preserve", xmlAttrValue(rel, "xml:space").?);
    const override = "<Override ext:flag='true' PartName='/xl/a.xml' ContentType='application/xml'";
    try std.testing.expectEqualStrings("/xl/a.xml", xmlAttrValue(override, "PartName").?);
    const default = "<Default ns:flag=\"1\" Extension = 'xml' ContentType = \"application/xml\"";
    try std.testing.expectEqualStrings("xml", xmlAttrValue(default, "Extension").?);
}

test "xmlStartTagEnd: a quote in following text cannot extend a valid start tag" {
    const xml = "<Relationship Id=\"rId1\">text with a stray \" quote > after</Relationship>";
    const end = xmlStartTagEnd(xml, 0).?;
    try std.testing.expectEqualStrings("<Relationship Id=\"rId1\"", xml[0..end]);
}

test "parseRelationships: commented relationships are not live (Codex r4)" {
    var arena = std.heap.ArenaAllocator.init(std.testing.allocator);
    defer arena.deinit();
    // A stale commented rId1 pointing elsewhere precedes the live
    // rId1: parsing the comment would silently redirect an append
    // into another sheet's drawing part.
    const rels = try parseRelationships(
        arena.allocator(),
        "<Relationships>" ++
            "<!-- <Relationship Id=\"rId1\" Type=\"x\" Target=\"../drawings/drawing2.xml\"/> -->" ++
            "<Relationship Id=\"rId1\" Type=\"x\" Target=\"../drawings/drawing1.xml\"/>" ++
            "</Relationships>",
    );
    try std.testing.expectEqual(@as(usize, 1), rels.len);
    try std.testing.expectEqualStrings("../drawings/drawing1.xml", rels[0].target);
}
