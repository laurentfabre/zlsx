# Load-modify-save plan (v3 — append-only path shipped)

> Phase 3c per the design-doc roadmap. Today's writer is fresh-file
> only — `Writer.init → addSheet → writeRow* → save`. No path to
> read an existing xlsx, mutate it, and save the result.

## Status

- iter-lms-1 (Editor scaffold + byte-identical passthrough) — ✅ shipped
- iter-lms-1b (raw-ZIP scanner + entry table) — ✅ shipped
- iter-lms-2 (numeric/bool/blank append + sheet substitute) — ✅ shipped
- iter-lms-3 (string append + SST extension) — ✅ shipped
- iter-lms-4 (C ABI + Python binding + docs) — ✅ shipped
- CLI sub-command (`zlsx append-rows … --out … < ndjson`) — ✅ shipped
- SST-less workbook string append — ✅ shipped (creates
  `xl/sharedStrings.xml` on demand; patches rels + Content_Types)
- Dimension column-bound widening on append — ✅ shipped

`Editor.open(allocator, path) / appendRows / save` is callable from
Zig today against any single-disk, non-ZIP64, non-encrypted xlsx
that already has an `xl/sharedStrings.xml` part.

## Problem

Real workflows want minimal mutations on an existing file —
typically appending rows (ETL output, audit logs) or updating
specific cells (status fields, recalculated metrics). Today the
only way is to re-create the entire workbook with `Writer`, which
is expensive, lossy on parts the writer doesn't model (charts,
pivots, drawings, custom XML, VBA), and bypasses Excel's careful
preservation of unrelated state.

## Scope choice — append-only v1

Three increasingly ambitious targets:

1. **Append-only**: read, append rows, save.
2. **Cell mutate**: read, change individual cell values, save.
3. **Full round-trip**: any cell / style / merge / sheet / etc.

This plan targets **(1) append-only** as v1. (2) is its own
follow-up plan because replacing a `<c>` span correctly requires
preserving formulas, inline strings, phonetics, and unknown
attributes/children — that's a real worksheet-tokenizer
commitment. (3) is multi-quarter scope.

Append covers ETL / audit / logging workflows ("write today's
rows to the bottom of the report"). It does NOT cover status-cell
updates ("change A1 to 'Done'"). Two plans, two ships.

## Constraints

- **Fidelity**: parts we don't model (charts, pivots, drawings,
  custom XML, VBA in `.xlsm`, theme overrides) round-trip
  byte-for-byte. The v1 design must not even decompress them.
- **Reader contract**: `Book.open` keeps working unchanged.
- **Writer contract**: `Writer.init` keeps working unchanged.
  `Editor` is a new entrypoint.
- **Concurrency**: single-threaded per `Editor`.
- Stdlib only.

## Approach: raw-ZIP scanner + worksheet append injection

```zig
pub const Editor = struct {
    pub fn open(allocator: Allocator, path: []const u8) !Editor;
    pub fn deinit(self: *Editor) void;

    /// Append rows to an existing sheet. v1 accepts only
    /// numeric / boolean / blank cells (no string cells until
    /// iter-lms-3 ships the SST-append path). Style indices
    /// refer to entries in the source's existing styles.xml —
    /// adding new styles is out of scope for v1.
    pub fn appendRows(self: *Editor, sheet_idx: u32,
                      rows: []const []const xlsx.Cell) !void;

    /// Write the modified workbook to `out_path`. Source archive
    /// stays untouched; pass the same path to overwrite (atomic
    /// rename via `out_path + ".tmp"`).
    pub fn save(self: *Editor, out_path: []const u8) !void;
};
```

### Raw-ZIP scanner (the iter-lms-1 deliverable Codex v2 surfaced)

Codex v2 was right that `Book.openLazy` doesn't preserve enough
ZIP state for byte-identical passthrough. The Editor builds its
own raw scanner pass at open:

- Read the source archive into an mmap'd / `[]u8` buffer.
- Parse the End-Of-Central-Directory record (capture its
  comment field verbatim — for non-zlsx-produced archives this
  may be non-empty).
- Walk the central directory: for each entry, record
  `(name, lfh_offset, lfh_byte_span, payload_byte_span,
  data_descriptor_span_or_null, cdfh_byte_span, gp_flags,
  compression_method)`.
- For each entry, validate the LFH signature + filename match,
  capture the LFH span (including filename + extra_field +
  extras after `<filename_len + extra_len>` — it varies per
  entry), payload span (compressed bytes), and post-data
  descriptor if GP flag bit 3 is set.

This is ~150 lines of careful ZIP code, not a one-line wrapper
around the existing reader. Worth its own iter.

### Save algorithm

`Editor.save(out_path)`:

1. Open `out_path + ".tmp"` for write.
2. Walk source ZIP entries in original order:
   - **Unmodified entry**: write LFH span + payload span + data
     descriptor (if any) verbatim. Track new local-header
     offset.
   - **Modified sheet's worksheet XML** (path resolved via
     `Book.sheets[sheet_idx].path` — NOT a synthesized
     `xl/worksheets/sheetN.xml`): rewrite to inject appended
     rows (see below). Compress with the existing in-house
     deflate. Emit a fresh LFH with new sizes / CRC. v1 does
     not preserve the original LFH's extra fields for the
     rewritten entry — it emits a minimal LFH (matching what
     `Writer.save` produces today).
   - **Modified `xl/sharedStrings.xml`**: only when iter-lms-3
     ships. v1 (numeric/bool/blank only) never touches the SST.
3. Re-emit the central directory using the captured CDFH spans
   for unmodified entries (with the new offsets) and freshly-
   built CDFH entries for modified ones.
4. Re-emit the EOCD record with the captured comment field
   preserved.
5. `fsync` + `rename` to `out_path`.

### Worksheet append-only rewrite

- Source sheet XML retained as a slice into the source archive
  buffer (decompressed once, kept resident).
- Find `</sheetData>` (literal string match — not within an XML
  comment because OOXML sheets don't put `</sheetData>` inside
  comments; defensive guard nonetheless rejects sheets where
  the literal appears inside a `<!--...-->` span).
- Everything before goes through verbatim; everything after
  goes through verbatim.
- Just before `</sheetData>`, inject the appended `<row r="N">
  <c r="...">…</c>… </row>` blocks with row indices computed
  from the source's highest used row (read once via the lazy
  `Rows` iterator, cached on first `appendRows` per sheet).

### `<dimension>` update — canonical form only

Codex v2: I was hand-waving variant handling. Concrete rule:

- If the source has a self-closing `<dimension ref="A1:Z100"/>`
  with a `top:bottom` rectangle ref, update it: replace the
  ref string in place to extend the row bound. Column bound
  unchanged (we only append).
- Any other shape (no `<dimension>`, opened-tag form, column-
  only ref, namespaced attribute, etc.): leave unchanged.
  Excel recomputes `<dimension>` on its next save anyway, so
  staleness is tolerable.

Document this in the API: appended rows expand the used range,
but `<dimension>` updates are best-effort.

## Why no SST reuse on string append

Codex v2 caught a real semantic risk: zlsx's `parseSharedStrings`
stores rich-text metadata in a parallel `rich_runs_by_sst_idx`
map (src/xlsx.zig:3383+). String equality match against the
existing SST could alias a plain-text append onto a rich-text
entry, making the newly written cell inherit the rich
formatting on next open.

iter-lms-3 (when it ships) follows the conservative rule:
**every new appended string gets a new SST entry**. No reuse.
This costs a small amount of bloat (duplicate plain-text
entries) and trades it for content-fidelity correctness.
Compaction is a future-iter feature behind an explicit
`Editor.compactSst()` call (out of scope; would require
re-numbering every `t="s"` reference).

A safer reuse is possible by checking `book.richRuns(idx) == null`
before alias-by-equality. Defer that optimisation past v1.

## Public API delta

```zig
// New top-level type. No changes to Book or Writer.
pub const Editor = struct { ... };
```

`Book` and `Writer` stay unchanged. `Editor` opens its own
internal `Book` (for reads) and an internal serialiser separate
from the public `Writer` — the writer's fresh-file model
doesn't compose with passthrough-existing-zip.

## Phasing — 4 iters

1. **iter-lms-1** (raw-ZIP scanner + byte-identical passthrough):
   build the raw-ZIP scanner described above. Introduce
   `Editor.open` / `Editor.deinit` / `Editor.save`. Save walks
   the source ZIP and writes every entry verbatim. Round-trip
   every corpus file: byte-identical SHA256 vs source. Add a
   corpus file produced by Excel (not zlsx) to test
   non-canonical extras / EOCD-comment / data-descriptor
   handling — the existing corpus is curated for the reader
   and may not cover ZIP edge cases.

2. **iter-lms-2** (numeric/bool/blank append): add
   `Editor.appendRows` accepting numeric / boolean / blank
   cells only. Implement the `</sheetData>` injection and
   canonical-form `<dimension>` update. Cache the source's
   highest used row per sheet on first call. Test: append rows
   to every corpus file, re-open with `Book.open`, confirm
   new rows readable AND original cells unchanged.

3. **iter-lms-3** (string-cell append with no-reuse SST): extend
   `appendRows` to accept string cells. New shared strings
   ALWAYS append a new SST entry; no reuse. Tests: appended
   rows with new strings, with strings carrying XML entities,
   with strings whose plain-text equals an existing rich-text
   entry (must NOT alias).

4. **iter-lms-4** (CLI + bindings + docs): expose `Editor` via
   the C ABI (`zlsx_editor_open` / `zlsx_editor_append_rows` /
   `zlsx_editor_save`) and Python binding (`zlsx.edit(path)
   → Editor` context manager). CLI: `zlsx append-rows <file>
   --sheet N` for shell pipelines. Update README + design
   doc; mark Phase 3c "append landed; cell-mutate deferred to
   a follow-up plan".

## Testing strategy

- **Byte-identical passthrough** (iter-lms-1): every corpus
  file `Editor.open → Editor.save` produces a file with
  identical SHA256 to the source. Includes a non-zlsx-
  produced corpus file (to catch ZIP shape variations).
- **Append correctness** (iter-lms-2/3): append known rows;
  re-open with `Book.open`; confirm new rows present at the
  expected indices, original cells preserved verbatim, no
  dupes/missing rows.
- **Excel + LibreOffice compatibility**: open the
  round-tripped output. Excel must not flag the file as
  repaired.
- **Fidelity**: corpus file with charts / drawings / VBA
  added before iter-lms-2. Editor must `open → appendRows →
  save` and the chart/drawing/VBA part must extract
  byte-identically (verifiable via unzipping both archives
  and comparing).
- **Fuzz**: random row appends; assert `Book.open` re-reads
  the new rows back as written, existing rows unchanged.

## Open risks

- **ZIP edge cases**:
  - **Data descriptors** (GP flag bit 3): captured in the
    raw scanner. Test corpus must include an Excel-written
    file (Excel sometimes uses these).
  - **ZIP64**: archives > 4 GB. Out of scope for v1; document
    the limit and refuse with a clear error.
  - **Encrypted entries**: refuse to open.
  - **Non-canonical EOCD**: archive-level comment is preserved
    verbatim. Spans validated with the central-directory
    record count.
- **Worksheet XML edge cases**:
  - **Trailing whitespace / CRLF after `</sheetData>`**: the
    injection point is the literal `</sheetData>`, so trailing
    bytes flow through as the post-suffix slice.
  - **Sheet without `<sheetData>`** (rare empty workbooks):
    refuse to append; emit an error.
  - **Multiple `</sheetData>` instances** (impossible in valid
    OOXML but defensive): use the first match; the post-suffix
    is from there to end.
- **`<dimension>` non-canonical forms**: skip update silently.
  Excel recomputes on next save.
- **Calc chain** (`xl/calcChain.xml`): unchanged in v1.
  Append doesn't touch formula references; cache stays valid.
- **Sheet path resolution**: via `Book.sheets[sheet_idx].path`,
  not synthesized.
- **Append vs read interaction**: appending after walking the
  full source via `Rows.next()` is fine; appending mid-walk is
  undefined. Document the constraint.
- **Most likely first failure**: a non-zlsx-produced corpus
  file with an EOCD comment, or with ZIP entries that have
  extras the existing reader silently skips. iter-lms-1's
  test corpus must include at least one Excel-written file
  to catch these.
