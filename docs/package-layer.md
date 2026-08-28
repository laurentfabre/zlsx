# `zlsx_pkg` — OOXML package layer

`zlsx_pkg` is a Zig module sitting alongside the main `zlsx`
reader/writer. It exposes the OOXML package layer (ZIP entries,
content types, relationships, drawings) as typed objects without
pulling the full reader into the build. Consumers who want to
extract images / charts / opaque parts from a workbook can use it
directly; the reader's overhead (sheet streaming, SST decode,
style lookup) doesn't apply.

> The title used to say **read-only**. It was true through Tier B0 and
> stopped being true in stages: the byte-preserving write API below, then
> the `Workbook` overlay's cell mutation, and at M5d the formula engine's
> recalculation entry points. See [Recalculation](#recalculation-m5d).

> Lives at `pkg/root.zig` in the repo. Wired into `build.zig` as
> the `zlsx_pkg` module name. Add it to your own `build.zig.zon`
> via a path or git dep on this repo.

## What you get

```zig
const pkg = @import("zlsx_pkg");

var store = try pkg.PartStore.open(allocator, "in.xlsx");
defer store.deinit();
```

### Read API

| Symbol | Notes |
|---|---|
| `PartStore.open(alloc, path)` | Slurps the file, walks the ZIP central directory, parses `[Content_Types].xml` and every `_rels/*.rels` part. |
| `partNames()` | Central-directory order; never sorted. |
| `part(name)` | Returns `?Part { name, content_type, bytes, compression_method }`. Bytes are decompressed (eager); arena-borrowed. |
| `rels(owner_part_name)` | Returns the relationships described by `_rels/<owner>.rels`. Owner `""` is the package root. |
| `resolve(owner, target)` | Joins a relationship `target` against `owner`'s parent dir, collapsing `..`/`.`. Absolute (`/xl/...`) targets bypass the join. |
| `imageParts()` | Filtered view: every part whose ContentType starts with `image/`. |

### Write API (byte-preserving)

| Symbol | Notes |
|---|---|
| `replacePart(name, bytes)` | Queues an override. Sub-1 KiB inputs ship STORED; larger inputs go through deflate (with a fall-back to STORED if compression doesn't shrink). |
| `addPart(name, content_type, bytes)` | Append a new part. Updates `[Content_Types].xml` to declare the part via an `<Override>` so consumers (Excel, openpyxl) accept the saved package. Same compression policy as `replacePart`. Returns `error.PartAlreadyExists` if the name is already in use. |
| `save(path)` | Atomic write. Untouched parts copy LFH + payload bytes byte-for-byte from the source. Replaced / added parts get fresh LFH + CDFH. EOCD comment preserved. Data-descriptor bytes (flag 0x0008) preserved verbatim. |

### Drawing helpers (C2a)

| Symbol | Notes |
|---|---|
| `imageAnchors(store, alloc)` | Walks every sheet's `<drawing r:id=...>` chain and returns `[]ImageAnchor` with `image_part_name` + `sheet_part_name` + cell-grid `from`/`to` anchors + image bytes. Surfaces `<xdr:absoluteAnchor>` images via the optional `absolute: ?AbsoluteAnchor` field (pixel-coordinate `x` / `y` / `cx` / `cy` in EMUs). |
| `chartAnchors(store, alloc)` | Same shape for `<xdr:graphicFrame>` containing `<c:chart>`. Each anchor exposes `chart_type` (`bar` / `line` / `pie` / `scatter` / `area` / `bubble` / `radar` / `other`) + `series_refs` (every `<c:f>` formula ref flattened in document order, across both Transitional and Strict chart-prefix bindings) + `raw_xml`. Also carries `absolute: ?AbsoluteAnchor` for pixel-anchored charts. |

OOXML namespace handling:
- Both Transitional URIs (`http://schemas.openxmlformats.org/...`) and Strict URIs (`http://purl.oclc.org/ooxml/...`) are accepted.
- Non-canonical prefixes (`dr:` / `dml:` / `chrt:` etc.) are resolved from the document's xmlns declarations, with whitespace tolerance around `=` and prefix lengths up to 100 chars.
- Multiple prefixes bound to the same xdr URI are all tracked (capped at 8); the scanner replays per prefix so anchors using ANY bound prefix are surfaced. Same-URI alts win over unused other-conformance declarations.
- xdr alt-prefix lookup walks the whole drawing (not just the first 4 KiB), so descendant-element xmlns declarations are reachable. The narrower 4 KiB scope is preserved for `a` / `c` lookups so a stray mid-document `xmlns:dml=...` can't shadow the canonical fallback.
- Chart elements are located by tag (`<*:chart`) rather than by pre-formatted needle, with the prefix verified per match against an in-scope binding lookup. An in-scope local binding is authoritative — a `<c:chart xmlns:c="non-chart"/>` is rejected even when `c` matches the drawing-root primary.
- XML namespace scoping is approximated via element-extent tracking: a binding on `<foo xmlns:p=".../>` ends at the matching `/>`, and a binding on `<foo xmlns:p="...">…</foo>` ends at the matching `</foo>` (with a depth counter for same-name nesting). Closed siblings don't leak their bindings into adjacent state.
- Comments (`<!-- ... -->`), CDATA (`<![CDATA[ ... ]]>`), and processing instructions (`<?...?>`) are skipped throughout — `xmlns:` text inside them isn't treated as a real binding, fake `<*:chart>` markup inside them isn't picked up as a candidate, fake `</foo>` text inside a PI doesn't unbalance the extent depth counter, and the forward state machine handles delimiter-shaped content nested inside another section type (`<![CDATA[<!--]]>`) without false transitions.
- Tag-end and xmlns scans are quote-aware: a literal `>` inside `descr="a > b"` doesn't end the tag prematurely, and an `xmlns:`-shaped substring inside a quoted attribute value is rejected as a real binding.
- The chart-element scan is O(n) flat: skip regions are eaten inline so adversarial inputs with many fake `:chart` substrings inside many comment/CDATA/PI sections don't trigger quadratic re-scans.

Caller-side lifetime contract for `chartAnchors`:
- Outer slice + each anchor's `series_refs` slice are **caller-allocated** — free both.
- Strings inside `series_refs` borrow from `raw_xml`; do not free.

## Example: extract every embedded image

The `zlsx-extract-images` binary in the release tarball does
exactly this:

```zig
const std = @import("std");
const pkg = @import("zlsx_pkg");

pub fn main() !u8 {
    var gpa: std.heap.GeneralPurposeAllocator(.{}) = .{};
    defer _ = gpa.deinit();
    const alloc = gpa.allocator();

    var store = try pkg.PartStore.open(alloc, "in.xlsx");
    defer store.deinit();

    var dir = try std.fs.cwd().makeOpenPath("out", .{});
    defer dir.close();

    for (try store.imageParts()) |p| {
        const basename = std.fs.path.basename(p.name);
        try dir.writeFile(.{ .sub_path = basename, .data = p.bytes });
    }
    return 0;
}
```

## Example: list charts and their data ranges

```zig
const charts = try pkg.chartAnchors(&store, alloc);
defer {
    for (charts) |c| alloc.free(c.series_refs);
    alloc.free(charts);
}
for (charts) |c| {
    std.debug.print("{s} ({s}): on {s}\n", .{
        c.chart_part_name,
        @tagName(c.chart_type),
        c.sheet_part_name,
    });
    for (c.series_refs) |r| std.debug.print("    {s}\n", .{r});
}
```

## Example: replace a part and save

```zig
var store = try pkg.PartStore.open(alloc, "in.xlsx");
defer store.deinit();

const new_workbook = "<?xml version=\"1.0\" ...";
try store.replacePart("xl/workbook.xml", new_workbook);
try store.save("out.xlsx");
```

Untouched parts in `out.xlsx` share their compressed payload bytes
with `in.xlsx` byte-for-byte; only `xl/workbook.xml` re-deflates.

## Recalculation (M5d)

The layer is no longer read-mostly: `zlsx_pkg.Workbook` carries the
formula engine's two entry points, and a third public module composes
them with the writer.

| Symbol | Notes |
|---|---|
| `Workbook.recalculate(alloc, io, run, opts)` | §5.7's in-memory transaction. Recalculates every formula cell and swaps the result in as the final pipeline operation. No file is opened; on any refusal or cancellation the workbook is exactly as it was. |
| `Workbook.saveWithRecalc(alloc, io, path, run, opts)` | The same pipeline plus §5.7.9's file transaction: serialize from the *unswapped* candidate → temp write → `File.sync` → final cancellation poll → rename → swap → directory fsync. Any failure before the rename leaves both the destination's prior bytes and the workbook's memory untouched. |
| `Workbook.openBuffer(alloc, io, bytes)` | Open a package already in memory. The borrow ends when it returns — the store copies. |
| `Workbook.openBufferControlled(alloc, io, bytes, ctl)` | Same, with §5.5's cancel/deadline reaching the archive scan and the eager part decompression. |
| `Workbook.markRecalcOnLoad()` | Set `<calcPr fullCalcOnLoad="1">` and change nothing else. Honestly named: it does not calculate. |
| `zlsx_recalc.writerSaveWithRecalc(alloc, io, writer, path, run, opts)` | The composition, in a **third module** (`zlsx_recalc`) that imports `zlsx` and `zlsx_pkg`. `Writer` bytes → `Workbook.openBuffer` → `saveWithRecalc`, with the `Control` threaded into all three stages. |

`RunInputs` (clock, UTC offset, seed, fidelity, platform profile, plus
the cancel token and deadline) makes a run reproducible from its inputs;
`RecalcOptions` carries the policy around it. The returned `RecalcReport`
is the caller's — `deinit` it with the allocator the workbook was opened
with.

`zlsx_recalc` is a separate module rather than a method on either of the
other two because the composition needs both halves at once, and putting
it in either would close a loop that today runs one way (`zlsx_pkg →
zlsx`). That direction is load bearing: `pkg/zip.zig` and
`pkg/fresh_emit.zig` are deliberately stdlib-only and take deflate as a
function pointer to keep the graph a DAG.

## Out of scope

- **Linked images** (`r:link` instead of `r:embed`) — these point at
  external files and aren't carried in the package; the drawing
  walker silently skips them.
- **External-workbook chart series refs** (`[Book.xlsx]Sheet1!A1`
  patterns inside `<c:f>`) — surfaced verbatim in `series_refs`;
  no path-resolution or external-part fetching.
- **Pivot tables** — read as a typed graph (`Workbook.pivotTables`;
  `zlsx_pkg.pivots.collect` over a `PartStore` plus its parsed
  `typed_parts.workbook_xml` view): tables with host
  sheet and output rectangle, caches with their source resolved to the
  sheet it reads from, field schema, records part. Never emitted or
  rewritten; the parts stay byte-preserved through every edit.
- **Per-part inferred metadata refresh** — `replacePart` /
  `addPart` install fresh bytes for the touched part(s), but
  metadata derived from `[Content_Types].xml` or `_rels/*.rels`
  on OTHER parts is NOT re-inferred until the next `open()`.
  Saved archives carry the new state correctly; in-process reads
  of derived state may be stale.

## Stability

`zlsx_pkg` is shipped as part of the post-0.2.9 roadmap's Tier-B0
foundation. The Zig API is expected to stay stable through Tier-B1
(`Workbook` overlay) — the future `Workbook` type will sit on top
of `PartStore` rather than replacing it.

If you need an editable workbook with cell-level access (read +
mutate cells, formulas, styles), use the main `zlsx` reader /
`Editor` instead. `zlsx_pkg` is for callers who want raw OOXML
package access without the cell-level overhead.
