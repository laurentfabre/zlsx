# `zlsx_pkg` — read-only OOXML package layer

`zlsx_pkg` is a Zig module sitting alongside the main `zlsx`
reader/writer. It exposes the OOXML package layer (ZIP entries,
content types, relationships, drawings) as typed objects without
pulling the full reader into the build. Consumers who want to
extract images / charts / opaque parts from a workbook can use it
directly; the reader's overhead (sheet streaming, SST decode,
style lookup) doesn't apply.

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
| `save(path)` | Atomic write. Untouched parts copy LFH + payload bytes byte-for-byte from the source. Replaced parts get fresh LFH + CDFH. EOCD comment preserved. Data-descriptor bytes (flag 0x0008) preserved verbatim. |

### Drawing helpers (C2a)

| Symbol | Notes |
|---|---|
| `imageAnchors(store, alloc)` | Walks every sheet's `<drawing r:id=...>` chain and returns `[]ImageAnchor` with `image_part_name` + `sheet_part_name` + cell-grid `from`/`to` anchors + image bytes. |
| `chartAnchors(store, alloc)` | Same shape for `<xdr:graphicFrame>` containing `<c:chart>`. Each anchor exposes `chart_type` (`bar` / `line` / `pie` / `scatter` / `area` / `bubble` / `radar` / `other`) + `series_refs` (every `<c:f>` formula ref flattened in document order) + `raw_xml`. |

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

## Out of scope

- **Linked images** (`r:link` instead of `r:embed`) — these point at
  external files and aren't carried in the package; the drawing
  walker silently skips them.
- **Absolute-pixel anchors** (`<xdr:absoluteAnchor>`) — uncommon
  enough that no corpus fixture exercises them today; tracker note
  in the roadmap.
- **Namespace-prefix tolerance** — the parser hard-codes `xdr:` and
  `a:`. Every Microsoft Excel + LibreOffice + xlsxwriter +
  openpyxl + python-calamine fixture in the repo's corpus uses
  these prefixes, but OOXML producers can technically pick any
  prefix. Workbooks with non-standard prefixes will silently
  surface zero anchors.
- **Pivot tables** — detected, opaque-byte preserved, never
  materialised as a typed object.
- **Adding a brand-new part** — `addPart` is the next public
  surface to grow; today only `replacePart` is shipped.

## Stability

`zlsx_pkg` is shipped as part of the post-0.2.9 roadmap's Tier-B0
foundation. The Zig API is expected to stay stable through Tier-B1
(`Workbook` overlay) — the future `Workbook` type will sit on top
of `PartStore` rather than replacing it.

If you need an editable workbook with cell-level access (read +
mutate cells, formulas, styles), use the main `zlsx` reader /
`Editor` instead. `zlsx_pkg` is for callers who want raw OOXML
package access without the cell-level overhead.
