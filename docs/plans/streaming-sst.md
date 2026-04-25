# Streaming SST plan (v6)

> Mitigation for design-doc question Q5: SST in memory for very large
> workbooks (e.g., 500 MB archives with millions of unique strings)
> blows RAM on `Book.open`. v0.2.x ships eager-only SST resolution.

## Problem

`parseSharedStrings(book, sst_xml)` walks every `<si>` element on
open, builds a `[][]const u8` shared-strings table in `sst_arena`,
and (for entries that need entity decoding) allocates a fresh
decoded buffer per entry. For a 50 MB SST with 5 M unique entries
(~30 bytes average) the eager pass produces ~150 MB of decoded
copies + ~80 MB of pointer-len pairs on top of the resident
`shared_strings_xml`.

Most callers never touch every entry — typical pipelines resolve
only the strings that actually appear in cells.

## Constraints

- **Lifetime contract** (src/xlsx.zig:14): strings returned from
  `Rows.next()` are borrowed from the `Book` and remain valid for
  the `Book`'s lifetime. The lazy backend MUST preserve this.
- **Perf floor**: 3.3 ms median on `worldbank_catalog.xlsx`
  (1,144 SST entries) is the no-regress line for the eager
  backend. Hot-path SST lookup must stay array-index-fast.
- **Threading** (CLAUDE.md): zlsx is single-threaded per `Book`.
  Plan does NOT add internal locking; documents the constraint.
- **Existing `Book.openLazy`** (sheet-lazy, eager SST,
  src/xlsx.zig:639+) keeps its current shape.
- Stdlib only. No new dependencies.

## Two backends, one accessor

`Book` carries a tagged union for SST storage:

```zig
const SstBackend = union(enum) {
    /// Eager — every entry pre-decoded into Book-owned storage.
    /// Today's representation, unchanged shape and perf.
    eager: [][]const u8,

    /// Lazy — pay decode cost only for accessed entries. The
    /// offsets/lengths index is always resident (~8 bytes/entry,
    /// 80 MB for 10 M entries vs 230+ MB for eager-decoded). The
    /// `resolved` map is sparse — RAM scales with accessed
    /// entries, not SST size.
    lazy: struct {
        offsets: []u32,
        lengths: []u32,
        resolved: std.AutoHashMapUnmanaged(u32, []const u8),
    },
};
sst: SstBackend,
```

Public access through one accessor:

```zig
/// Resolved shared string at idx. O(1) on eager (array index);
/// O(1) avg on lazy (hashmap lookup, decode-on-miss into
/// `sst_arena`). Errors only on out-of-range idx.
pub fn sharedStringAt(self: *Book, idx: usize) ![]const u8

/// Total entries in the SST. Both backends know this at open()
/// without materialising any entry.
pub fn sharedStringsCount(self: *const Book) usize
```

The hashmap path is slower per access than array indexing
(measured ~30 ns vs ~5 ns) but only fires on lazy mode. Lazy is
opt-in; callers asking for lazy accept that trade.

## Why hashmap, not optional-array

Codex review v4/v5 surfaced the right concern: a `[]?[]const u8`
of `uniqueCount` length still pays one slot per entry —
~120 MB on 10 M entries. A sparse hashmap holds only resolved
entries: at 1% access on 10 M entries, the map holds 100K entries
(~3 MB). That's the actual win for sparse-access workloads.

The eager backend keeps `[][]const u8` unchanged — the hashmap
overhead doesn't exist on the eager hot path.

## Public surface migration

The current public field `Book.shared_strings: [][]const u8` is
read directly by:

- `Rows.resolveSharedString` (src/xlsx.zig:1843).
- `src/cli.zig:2385,2426,3110` — the `meta` and `sst` sub-commands.
- `src/c_abi.zig:642,655` — `zlsx_sst_count` / `zlsx_sst_at`.
- ~20 internal test/fuzz sites in `src/xlsx.zig`.

iter-sst-1 changes only `Rows.resolveSharedString` to call
`book.sharedStringAt(idx)`; `Rows.book` stays `?*const Book`
because `sharedStringAt` on the eager backend doesn't mutate. The
new accessor takes `*Book` (mutable) but the eager path never
touches the union's discriminant — the type signature is wider
than the runtime mutation requires.

iter-sst-2 migrates the CLI + C ABI direct readers to the
accessor. iter-sst-3 introduces the lazy backend, by which point
every reader is already routed correctly and the tagged-union
flip is behavioural-only.

The public field name changes (`shared_strings` →
`sst_storage_internal` or removed) on the same release the lazy
backend ships, so external callers see one breaking change at a
known boundary, not a churn series.

## Rich runs — eagerly built on both backends

`rich_runs_by_sst_idx` and its `*const Book` accessor stay
unchanged. Rich entries are typically <1% of an SST; eagerly
walking and populating the map on `openSstLazy` adds ~1% of the
eager cost and preserves the const accessor + Book-lifetime slice
contract.

If a workload appears with high rich-text density, revisit in v7.

## Ownership model on `deinit`

Today: `a.free(self.shared_strings)` then `self.sst_arena.deinit()`
(src/xlsx.zig:1028). Slice headers freed first; arena destroys
the bodies.

After lazy ships, `Book.deinit` must:

1. Free the discriminant-specific containers:
   - **Eager**: free the `[][]const u8` slice header (today).
   - **Lazy**: free `offsets`, `lengths`, `resolved.deinit(allocator)`.
2. Destroy `sst_arena` (which owns every materialised entry's
   bytes for both backends).
3. Free `shared_strings_xml` (today, unchanged).

The hashmap `.resolved` only stores `[]const u8` slice headers —
the bodies live in `sst_arena`. The hashmap's own backing is
freed by `.deinit(allocator)` before the arena is destroyed.

## Phasing — 4 iters

1. **iter-sst-1** (minimal route through accessor): add
   `Book.sharedStringAt(*Book, idx)` and
   `Book.sharedStringsCount(*const Book)`. Migrate
   `Rows.resolveSharedString` to call the accessor. Keep
   `Book.shared_strings` field, keep `Rows.book: ?*const Book`.
   Bench gate: no regression on `worldbank_catalog.xlsx`.

2. **iter-sst-2** (CLI + C ABI migration): migrate the direct
   `book.shared_strings[i]` readers in `src/cli.zig` and
   `src/c_abi.zig` to the accessor. Tests stay on the field
   directly until iter-sst-3 (they're internal). Still zero
   behavioural change.

3. **iter-sst-3** (the lazy entrypoint, the actual feature):
   - Replace `Book.shared_strings: [][]const u8` with
     `sst: SstBackend` tagged union. Eager constructor is
     `.eager = [...]`; lazy constructor is the offset table +
     empty hashmap.
   - `Book.openSstLazy(allocator, path)`: skip eager decode;
     build the offset table during the SST walk; eagerly
     populate `rich_runs_by_sst_idx`.
   - `sharedStringAt` branches on the union tag; lazy hits do
     hashmap lookup + decode-on-miss into `sst_arena`.
   - Migrate the remaining test/fuzz sites to the accessor.
   - Update `Book.deinit` per the ownership model above.
   - Synthetic 1 M-entry SST test: lazy with 1% access stays
     under ~120 MB resident (~80 MB index + ~3 MB resolved
     map + small `sst_arena`); eager refuses to alloc under
     `std.testing.allocator`'s OOM trap.

4. **iter-sst-4** (CLI + docs): wire `Book.openSstLazy` into
   opt-in CLI invocations (`--sst-lazy` on `cells` / `rows`).
   `zlsx sst` rejects lazy or transparently swaps to eager.
   Document the new entrypoint in the library README + the
   design doc; update Q5 status from "queued" to "shipped".

## Threading contract

> One `Book` per thread (CLAUDE.md). `sharedStringAt` is not
> thread-safe on the lazy backend — first-touch mutates the
> resolved hashmap. Multi-threaded SST access on a lazy Book
> requires the caller to serialise externally or pre-materialise
> the SST (via a public `Book.materialiseAllSst()` helper, added
> in iter-sst-4 if asked).

Internal locking is rejected: it adds runtime cost on the eager
backend and contradicts the existing per-Book single-thread
convention.

## Testing strategy

- **Equivalence**: every corpus file gives byte-identical
  resolution across both backends, for every cell.
- **Perf gate**: 4-file read benchmark must not regress past
  1.5% wall-time on the eager backend after any iter lands.
  Lazy on `worldbank_catalog.xlsx` should be within 5% of
  eager.
- **RSS ceiling**: synthetic 1 M-entry SST (~50 MB xml). Eager
  refuses to alloc under the testing allocator's OOM trap;
  lazy with 1% access stays under ~120 MB resident.
- **Fuzz**: extend `fuzz parseSharedStrings` to seed both
  backends; assert resolution agrees byte-for-byte across 1M
  iters. Add an "access-pattern" axis: random sparse vs full
  sweep.
- **Concurrency negative test**: a debug-mode assertion in the
  lazy materialise path detects re-entrance (thread-local
  "currently materialising" flag in safe builds).

## Open risks

- **iter-sst-3 is the heaviest slice**: tagged-union flip + 4
  ownership transitions + hashmap plumbing + test migration.
  Probably 2-3 days of work. iters 1, 2, 4 are each <½ day.
- **Hashmap perf on hot lazy path**: 30 ns per cell vs 5 ns on
  eager. For a 1M-cell workload that's 25 ms extra. Worth
  measuring with a synthetic before iter-sst-3 lands; can swap
  to a paged array (`[]?[]const u8` chunked into 64 K-entry
  pages, freed individually) if hashmap proves too slow.
- **`shared_strings_xml` retention**: required for lazy decode.
  Halving the SST XML's resident footprint requires
  mmap-the-zip — separate plan.
- **Bindings**: `bindings/python` doesn't reach into
  `shared_strings` directly (goes through C ABI). C ABI
  accessors migrate in iter-sst-2. No Python-side change
  needed.
- **Status as "final" plan**: this is v6 after four Codex
  review rounds. Further architectural refactors (mmap'd zip,
  sparse-rich-runs, multi-threaded resolution) are deferred to
  separate plans.
