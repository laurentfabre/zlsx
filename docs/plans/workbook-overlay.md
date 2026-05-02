# Workbook typed-overlay plan (v1 draft)

> Tier B1 in `post-0.2.9-roadmap.md`. PartStore (B0) is shipped —
> read, byte-preserving save, `replacePart`, `addPart`, hardening.
> The next layer is a typed `Workbook` overlay sitting on top of
> the part store, surfacing `Worksheet` / cell / style / SST / DV /
> CF / hyperlinks / comments / merges as a single mutable model.
> B2 (Editor rebase) and B3 (Writer rebase) follow once this layer
> is bedded down.

## Status (as of 2026-05-02 — tier complete)

- iter-wb-1 (typed-overlay handles on PartStore) — ✅ shipped (PR #21, 5 typed parsers + 44 inline tests + corpus sweep, ~4181 LOC)
- iter-wb-2 (Workbook + Worksheet root types, read-only) — ✅ shipped (PR #22)
- iter-wb-3 (`Workbook.fromBook(book, path)` adapter — Book-side `toWorkbook` shim skipped to avoid a circular module dep) — ✅ shipped (PR #29)
- iter-wb-4 (Worksheet mutation surface + delta-emit save) — ✅ shipped across m1+m2+m3+m4 (PRs #24-#27): `setCell` for blank / number / boolean / string (inlineStr) / formula / shared_string + `Workbook.save` byte-preserving outside `<sheetData>` + SST extension (creates `xl/sharedStrings.xml` when absent)
- iter-wb-5 (workbook-level state views — defined-names scope, sstText, cellByRef) — ✅ shipped (PR #23)
- iter-wb-6 (RSS gate + corpus parity sweep) — ✅ shipped (PR #28 measurement infra; gate closed at 0.78× via PRs #30 deferred-decompress + #31 file-streaming PartStore)

## Problem

`Book` (`src/xlsx.zig:547`) is a 14 706-line all-in-one reader
state, with every part's parsed view (SST, styles, themes, sheets,
defined names, merges, hyperlinks, validations, conditional
formats, comments, drawings) resident on the same struct.
`Editor` (`src/xlsx.zig:4360`) layers a substring-and-substitute
mutation model on top, with `pending_mutations` /
`pending_new_sheets` shadow tables (`src/xlsx.zig:4389,4432`)
tracked as auxiliary state.

Three structural problems fall out:

1. **No write-side type for known parts.** PartStore's typed
   overlay milestone (B0 M3) was deliberately left for B1 because
   the typed Workbook view is what makes typed overlays useful.
2. **Reader / editor coupling.** Editor mutations touch Book
   internals directly; structural edits require careful guard
   negotiation per category (`anySheetCrossSheetCarrier` at
   `src/xlsx.zig:5865`). A typed Workbook lets B2 rebase Editor on
   a stable model without churning the reader.
3. **Writer / reader divergence.** `src/writer.zig` is a fresh-
   file emitter (6 521 lines, separate code path). B3 rebases it
   onto Workbook so SST and styles flow through one model.

B1 lands the type. B2 + B3 rebase consumers. This plan is B1
only.

## Constraints

- **Stdlib-only contract preserved** (no third-party Zig deps).
- **Module-graph constraint (Zig 0.15.2).** Today writer / cli /
  pkg can't all coexist in one compilation because `src/package/`
  used to live under writer's package dir. The package layer
  already moved to `pkg/`; B1 places `Workbook` in `pkg/` so the
  same workaround applies — `cli`, `writer`, and `workbook` stay
  in separate package dirs.
- **Reader stability.** Adapter-only path first per the roadmap
  critique ("Mitigate by adapter-only path first; rebase reader
  internals last"). `Book` keeps its current shape for one minor
  line; iter-wb-3 supplies the `Book.toWorkbook(options)` bridge.
- **PartStore byte-preservation.** Untouched parts' compressed
  payload, name, and order stay intact across `Workbook.save`.
  Identical contract to B0's `replacePart`.
- **RSS floor.** A 100k-row × 10-sheet workbook MUST need ≤ 2×
  current `Book.openLazy` RSS before any sheet is touched. This
  is the roadmap's walk-away gate; iter-wb-6 measures.
- **Single-threaded contract per `Workbook`.** Same shape as
  `Book` today. No internal locking.

## Public surface (sketch)

```zig
// Module: zlsx_pkg (pkg/root.zig) — re-exports.
// Implementation: pkg/workbook.zig + pkg/typed_parts/*.zig.

pub const Workbook = struct {
    pub fn open(allocator, path) !Workbook;        // PartStore.open + lazy overlay
    pub fn openLazy(allocator, path) !Workbook;    // sheet-lazy variant
    pub fn create(allocator) Workbook;             // greenfield, no PartStore yet
    pub fn deinit(self: *Workbook) void;
    pub fn save(self: *Workbook, path) !void;      // byte-preserve untouched + emit dirty

    pub fn sheet(self: *Workbook, idx: u32) !*Worksheet;
    pub fn sheetByName(self: *Workbook, name) ?*Worksheet;
    pub fn sheetCount(self: *const Workbook) u32;

    pub fn definedNames(self: *Workbook) DefinedNamesView;
    pub fn styles(self: *Workbook) StylesView;
    pub fn sst(self: *Workbook) SstView;          // read-side wraps existing lazy backend
};

pub const Worksheet = struct {
    pub fn rows(self: *Worksheet) !Rows;
    pub fn cell(self: *Worksheet, ref: A1Ref) ?CellValue;
    pub fn setCell(self: *Worksheet, ref: A1Ref, value: CellValue) !void;  // iter-wb-4

    pub fn merges(self: *const Worksheet) []const MergeRange;
    pub fn hyperlinks(self: *const Worksheet) []const Hyperlink;
    pub fn validations(self: *const Worksheet) []const DataValidation;
    pub fn conditionalFormats(self: *const Worksheet) []const ConditionalFormat;
    pub fn comments(self: *const Worksheet) []const Comment;
};

// Compat facade (one minor line):
pub fn (self: *Book) toWorkbook(allocator, opts: ToWorkbookOptions) !Workbook;
```

`Workbook.open` is `PartStore.open(allocator, path)` plus a lazy
typed-overlay attach. It does NOT eagerly parse every sheet — the
`Worksheet` is a lightweight handle until `rows()` / `cell()` /
`setCell()` materialises the parse via the existing
`ensureSheetLoaded` (`src/xlsx.zig:926`) plumbing.

## Why typed-overlay handles before Workbook

iter-wb-1 alone gives PartStore typed views of well-known parts
(workbook.xml, sheet, sst, styles, theme) without any new
top-level type. This is B0 M3 ("Typed overlays for known parts")
which the roadmap marks "likely consolidated into B1." Shipping
it first gives Workbook a clean substrate to compose from
(`pkg/typed_parts/workbook_xml.zig`, `sheet_xml.zig`, etc.) and
keeps the iter-wb-1 PR diffable: no API surface change to the
existing `zlsx`, `zlsx_pkg`, or `Editor`.

## Iter sequencing

### iter-wb-1 — Typed-overlay handles on PartStore (1–2 weeks)

**Scope:** `pkg/typed_parts/{workbook_xml,sheet_xml,sst_xml,styles_xml,theme_xml}.zig`.
Each module exposes a `parse(allocator, bytes) !T` returning a
borrow-from-bytes typed view (no copies of leaf strings).

**Extends:** `PartStore` gains `partTyped(name) !TypedPart`
returning a `union(enum) { workbook: WorkbookXml, sheet:
SheetXml, sst: SstXml, styles: StylesXml, theme: ThemeXml,
opaque: Part }`. Resolves the part by name, looks up content type,
dispatches to the parser. Falls back to `.opaque` for unknown
content types (preserves byte-preserving contract).

**Tests:** typed-overlay round-trip on every fixture in
`tests/corpus/` (decode then re-emit; bytes equivalent for an
opaque part, semantically equivalent for a typed one). Add a
`pkg/typed_parts/test_basic.zig` covering happy-path each parser.

**Walk-away:** if a typed parser balloons past ~400 lines or
needs a third-party dep, demote that part to `.opaque` and defer
the typed view to a follow-up iter.

### iter-wb-2 — Workbook + Worksheet root types, read-only (2 weeks)

**Scope:** `pkg/workbook.zig` introduces `Workbook` and
`Worksheet`. Read-only surface: `cell`, `rows`, `merges`,
`hyperlinks`, `validations`, `conditionalFormats`, `comments`,
`sheet`, `sheetByName`, `sheetCount`, `styles`, `sst`,
`definedNames`.

**Implementation:** `Workbook` holds a `PartStore` and the typed
overlays from iter-wb-1. For sheet-data access it composes an
internal `xlsx.Book` (the existing reader) — does NOT duplicate
parsing. The internal Book is held by-value, opened against the
same backing file, and exposes its `Rows` / `cell` paths through
Worksheet methods.

**Why composition over reimplementation:** preserves reader
stability (the roadmap's primary risk for B1) and avoids
duplicating ~14 700 lines of parser. The trade-off is that
`Workbook` holds a `Book` until B2/B3 rebase the consumers; iter-
wb-4's mutation path tracks deltas separately rather than
mutating Book's state.

**Tests:** the existing `tests/xlsx_corpus.zig` is mirrored as
`tests/workbook_corpus.zig`, opening every fixture through both
APIs and asserting parity (cell counts, row totals, sheet names,
styles, merges, hyperlinks, validations, conditional formats,
comments).

**Walk-away:** if internal-Book lifetime gets tangled (e.g.
`Worksheet.rows()` returns slices owned by Book that outlive
their owner), revert to a thinner Workbook that delegates open()
calls and document the layered ownership.

### iter-wb-3 — `Book.toWorkbook` adapter (1 week)

**Scope:** `pub fn (self: *Book) toWorkbook(allocator,
ToWorkbookOptions) !Workbook` on the existing `Book` type, plus
`Workbook.fromBook(*Book)` constructor for the same effect from
the Workbook side.

**Use case:** existing callers that already opened a `Book` via
`Book.open` / `Book.openLazy` / `Book.openSstLazy` can promote
their handle to a `Workbook` without re-reading the file.
`ToWorkbookOptions` covers eager-vs-lazy SST + sheet preferences;
defaults match the source `Book`.

**Tests:** every `Book.open*` corpus test gets a
`.toWorkbook()` mirror in `tests/workbook_corpus.zig`; assert
identical cell content, sheet count, SST entries.

### iter-wb-4 — Worksheet mutation surface + delta-emit save (2–3 weeks)

**Scope:** `Worksheet.setCell(ref, value)`,
`Worksheet.setCells(refs, values)`. Mutations accumulate in a
per-Worksheet delta map (`std.AutoHashMapUnmanaged(CellRef,
CellValue)`) — does NOT touch the internal Book's parsed state.

**`Workbook.save(path)`:**
1. Iterate dirty worksheets (delta map non-empty).
2. For each, materialise the source sheet XML via the typed
   `SheetXml` overlay (iter-wb-1), apply deltas (replace existing
   `<c>` spans, append new ones, widen sheet `<dimension>`),
   re-emit through PartStore's `replacePart`.
3. SST extension: new strings flow through the existing append
   path used by `Editor` today (`src/xlsx.zig:6993` SstAppender).
4. Untouched parts: PartStore byte-preserves them.

**Out of scope for iter-wb-4:** structural shape changes (insert
row, delete column, rename sheet), formula rewriting, defined-
names mutation. Those land in C1 / B2 / B3.

**Tests:** corpus round-trip with single-cell mutations;
SST-extension on string-typed setCell; SheetMatrix-driven
parity vs `Editor.setCell` on the same fixture.

**Walk-away:** if delta-emit reuses ≥ 70% of `Editor.setCell`'s
existing logic, the value-add is small — pause iter-wb-4 and let
B2 (Editor rebase) drive the unification instead.

### iter-wb-5 — Workbook-level state views (1–2 weeks)

**Scope:** `SstView` (read-side wrapper over existing
`Book.sharedStringAt`), `StylesView` (over the existing styles
parse), `DefinedNamesView` (workbook + sheet scope). All read-
only in this iter; mutation lands when consumers need it
(B2/B3).

**Why split from iter-wb-2:** iter-wb-2 covers per-Worksheet
surface; this iter covers workbook-scope views that don't fit on
Worksheet. Splitting keeps the iter-wb-2 PR scoped.

### iter-wb-6 — RSS gate + corpus parity sweep (1 week)

**Walk-away gate from the roadmap:** "100k-row × 10-sheet
workbook needs ≤ 2× current `Book.openLazy` RSS before any sheet
is touched."

**Scope:** synthesise the 100k×10 fixture (or use an existing
large-corpus entry), measure RSS via `getrusage` /
`mach_task_info` / Windows equivalent on the bench job, fail the
gate if `Workbook.openLazy` exceeds 2× the `Book.openLazy`
baseline. Plus a sweep of `tests/corpus/` confirming
Workbook-vs-Book parity on every fixture.

**Tests:** new `tests/bench/workbook_rss.zig` (off the default
`zig build test` path; runs in the bench CI job alongside
`hyperfine`).

## Risks

- **Internal-Book ownership tangling.** The biggest risk:
  `Worksheet.rows()` returns iterator state that borrows from the
  internal Book; `Workbook.deinit` ordering must keep Book alive
  until all `Rows` are consumed. Mitigate via explicit lifetime
  comments + `std.testing.allocator` leak detection on every
  iter-wb-2/3 test.
- **`Workbook.save` divergence from `Editor.save`.** Two save
  paths through PartStore is a duplicate-source-of-truth risk
  until B2/B3 unify them. Mitigate by extracting the shared
  delta-emit core into `pkg/save_core.zig` early and having both
  call sites depend on it.
- **Module-graph collision (Zig 0.15.2).** Workbook in
  `pkg/workbook.zig` plus `cli` plus `writer` may surface the
  same "file exists in modules X and Y" error documented in the
  roadmap. Mitigate by keeping `pkg/` as the single home for
  Workbook and not adding direct `@import` paths from `src/cli.zig`
  or `src/writer.zig` into `pkg/`.
- **Compat-facade drift.** `Book.toWorkbook` must stay accurate
  across one minor line. Test it against every option combination
  on every corpus fixture.

## Out of scope (explicit)

- B2 Editor rebase (separate plan, after this).
- B3 Writer rebase (separate plan, after B2).
- C1 formula rewriting on Workbook (separate plan).
- C2b `addImage` on Workbook (separate plan, depends on B1
  iter-wb-2 + C2a).
- D1 evaluator, D2 chart emit — deferred indefinitely per
  roadmap.

## Estimation

| Iter | Estimate | Depends on |
|---|---|---|
| iter-wb-1 typed-overlay handles | 1–2 w | B0 (shipped) |
| iter-wb-2 Workbook + Worksheet (read-only) | 2 w | iter-wb-1 |
| iter-wb-3 `Book.toWorkbook` adapter | 1 w | iter-wb-2 |
| iter-wb-4 mutation surface + delta-save | 2–3 w | iter-wb-2 |
| iter-wb-5 workbook-level views | 1–2 w | iter-wb-2 |
| iter-wb-6 RSS gate + parity sweep | 1 w | iter-wb-2..5 |

End-to-end: **6–10 weeks** (matches the roadmap estimate).
