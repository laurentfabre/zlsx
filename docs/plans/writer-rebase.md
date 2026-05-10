# Writer rebase onto Workbook (B3 plan, wr-0 inventory)

> Tier B3 in `post-0.2.9-roadmap.md:401`. B0 (PartStore), B1 (Workbook
> typed overlay) and B2 (Editor rebase) are shipped. `Workbook.save`
> (`pkg/workbook.zig:608`) → `PartStore.save` (`pkg/store.zig:502`) is
> the substrate. `Editor.save` is now a 14-line shim. **B3 makes
> `xlsx.Writer.save` (`src/writer.zig:665`) a thin shim over
> `Workbook.save` too**, retiring the 6579 LOC fluent-builder fork.
>
> This document is the **wr-0 inventory** — read-only enumeration of
> every distinct emit path, every byte-stable invariant the rebase
> must preserve, and the proposed iter-wr-1..7 sequencing. Companion
> to `editor-rebase.md`. No code is touched in wr-0.

## Status (as of 2026-05-10)

- B0 PartStore — ✅ shipped
- B1 Workbook typed overlay — ✅ shipped
- B2 Editor rebase — ✅ closed; `Editor.save` is a 14-line shim, `pkg/editor.zig` shrank 6021 → 3231 LOC (-46%)
- **B3 Writer rebase — ✅ closed (7/7 iters shipped)**
  - wr-0 inventory ✅ #91 (this doc)
  - wr-1 SST ✅ #94 — `pkg/sst_plan.zig` (309 LOC, std-only)
  - wr-2 Styles ✅ #95 — `pkg/styles_plan.zig` (842 LOC; most byte-fragile axis — pinned by a byte-equivalence parity test that locks every §1.10 invariant)
  - wr-3 workbook.xml ✅ #97 — `pkg/workbook_xml_plan.zig` (472 LOC) + `Workbook.addDefinedName` with full Excel name-rule validation
  - wr-5 ZIP ✅ #96 — `pkg/zip.zig` (354 LOC, std-only; takes a `DeflateFn` callback to avoid the `writer → pkg → workbook → writer` module-graph cycle)
  - wr-4 sheet emit ✅ #98 — `pkg/sheet_plan.zig`; 7 byte-equivalence parity tests + 24 standalone tests; bench 6.5 ms = 1.00× of 6.7 ms baseline.
  - wr-6 helper dedup + `validateSheetName` reconciliation + `SheetState` extraction ✅ #100 — `pkg/sheet_plan.zig` (~1565 LOC) gains the `SheetState` registry struct; `src/writer.zig` 5947 → 5606 LOC (-341, -5.7%); `validateSheetName` lifted to Writer's Unicode-scalar-aware version (Workbook delegates). The 13 SheetWriter `add*` / `set*` methods are now thin forwarders over `SheetState`.
  - **wr-7 Workbook fresh-emit + Writer.save shim + corpus parity sweep ✅ closed** — new `pkg/fresh_emit.zig` (527 LOC, std-only) hosts the entire archive orchestration ([Content_Types].xml + rels + workbook.xml + per-sheet sheet/rels/comments/vml + sst + styles + ZIP CD/EOCD). `Writer.save` collapsed to a **17-line thin shim** that projects Writer state onto `fresh_emit.ArchiveInputs` and delegates. `Workbook.saveFreshEmit(path)` ships through the same substrate against per-Worksheet `body` + `sheet_state` + workbook-level `fresh_sst_plan` / `fresh_sst_count` / `styles_plan` / `workbook_xml_plan`. 13 `add*`/`set*` forwarder methods added to `pkg.Worksheet` (mirroring `xlsx.Writer.SheetWriter`). 5 new parity tests pin Workbook fresh-emit byte-equivalence to Writer.save. `src/writer.zig` 5606 → 5256 LOC (-350, -6.2% this iter; cumulative 6579 → 5256 = -1323 LOC = -20.1%).

3 B3 prep PRs landed alongside the iters: #92 `PartStore.fresh()`, #93 `Workbook.empty()` + `SstExtensionPlan` rich-string axis, #90 docs flip B2 to ✅.

**Architectural pattern locked in**: each Writer subsystem extracts into a std-only `pkg/<subsystem>_plan.zig` module that Writer + Workbook both consume. The wr-7 closer adds `pkg/fresh_emit.zig` as the shared archive-build orchestrator. `src/writer.zig`: 6579 → 5256 LOC (-1323, -20.1%) cumulative.

## Problem

`xlsx.Writer` (`src/writer.zig:393`) is the **fresh-file / Phase 3b**
producer surface. The intended call shape is:

```
Writer.init → addStyle? → addDxf? → addDefinedName? → addSheet
  → SheetWriter.{writeRow, writeRowStyled, writeRowWithFormulas,
    writeRichRow, setColumnWidth, setRowHeight, freezePanes,
    setAutoFilter, addMergedCell, addHyperlink,
    addInternalHyperlink, addComment, addDataValidation*,
    addConditionalFormat*}
  → Writer.save
```

`Writer.save` walks every registered sheet plus the writer's own
shared-string table, styles list, dxf list, numFmt pool, and defined
names; serialises the workbook from scratch into a fresh
`[Content_Types].xml` + `_rels/.rels` + `xl/workbook.xml` +
`xl/_rels/workbook.xml.rels` + `xl/sharedStrings.xml` + (optional)
`xl/styles.xml` + per-sheet `xl/worksheets/sheetN.xml` (+ optional
sheet-rels, `xl/commentsN.xml`, `xl/drawings/vmlDrawingN.vml`); then
hands the buffer to a private `ZipWriter` that emits LFH + payload +
CDFH + EOCD into a `std.ArrayListUnmanaged(u8)`.

Workbook's `save` model is fundamentally different — it loads an
existing file via PartStore, applies per-Worksheet `setCell` /
`appendRows` deltas to the parsed-and-edited part bytes, and writes
back. Three structural problems fall out:

1. **Two emit pipelines, zero shared code (other than
   `deflateCompress`).** Writer owns `sst_strings` / `sst_index` /
   `sst_is_rich` / `styles` / `dxfs` / `num_fmts` / `num_fmt_index`
   / `defined_names`; Workbook owns the equivalents typed at
   `pkg/typed_parts/{sst_xml,styles_xml,workbook_xml,sheet_xml}.zig`.
   Two SST builders, two style serialisers, two `<workbook>` emitters,
   two `<sheetData>` row emitters, two ZIP archivers — each one a
   duplicate-source-of-truth risk for byte-stability invariants
   (`<cellStyles>` order, Defaults-before-Overrides in
   `[Content_Types].xml`, `<definedNames>` between `</sheets>` and
   `</workbook>`, etc.).
2. **Workbook can't yet emit fresh-file shapes.** Workbook today
   handles `setCell` deltas + `appendRows` substring-splice; it has
   no path for fresh `<conditionalFormatting>` / `<dataValidations>`
   / `<hyperlinks>` / `<autoFilter>` / `<cols>` / `<sheetView>`
   panes / `<mergeCells>` / `<commentList>` / `vmlDrawing` from
   scratch (it can rewrite existing instances via
   `rewriteAllValidationsAndConditionalFormats`,
   `rewriteAllHyperlinkLocations`, etc., but that's an edit
   surface, not a producer). B3 needs Workbook to gain a
   "fresh-emit" mode for these constructs OR for Writer to keep
   building them locally and hand them to Workbook through a typed
   handoff.
3. **Three duplicated XML helpers.** `appendXmlEscaped` exists at
   `src/writer.zig:2706`, `pkg/store.zig:416`, and
   `pkg/workbook.zig:2173`/`4391` (the latter is
   `appendXmlEscapedText`, attribute-vs-text variant). `validateSheetName`
   is in `src/writer.zig:2210` and `pkg/workbook.zig:2009`. These are
   free wins for wr-6 cleanup once one is the canonical home.

The win is concrete: one SST, one styles registry, one ZIP emit,
one byte-stability surface to defend, ~5500 LOC of duplicate Writer
code retired.

## Constraints

- **Byte-stability preserved.** Existing Excel / openpyxl /
  LibreOffice / calamine round-trips MUST continue to produce
  byte-identical (or at minimum byte-equivalent — whitespace inside
  XML is the only acceptable drift) Writer output post-rebase.
  See **Section 2 invariants** for the catalogue.
- **Phase 3b perf gate.** `tests/bench/bench_write_zlsx.zig`
  measures 1 k rows × 10 cols at **6.7 ms ± 0.3** (4.44 MB peak —
  see `docs/benchmarks.md:145`). The rebase MUST not regress past
  **1.10× → ~7.4 ms ceiling**; > 1.10× walks back to a Writer-local
  fork for the regressed axis.
- **One-minor-line API compat.** `xlsx.Writer.{init, deinit, save,
  addSheet, addStyle, addDxf, addDefinedName}` and
  `SheetWriter.{writeRow, writeRowStyled, writeRowWithFormulas,
  writeRichRow, setColumnWidth, setRowHeight, freezePanes,
  setAutoFilter, addMergedCell, addHyperlink, addInternalHyperlink,
  addComment, addDataValidation*, addConditionalFormat*}` keep
  current signatures. The c_abi surface (`src/c_abi.zig:1403`+) is
  also frozen — `zlsx_writer_*` exports stay byte-compatible.
- **Module-graph.** `src/writer.zig` lives in `src/`; `pkg/workbook.zig`
  lives in `pkg/`. The rebase imports `pkg.Workbook` from
  `src/writer.zig` (already shipped at `src/xlsx.zig` for
  `Workbook.fromBook`). No new cross-package collisions.
- **Single-threaded contract** preserved.

## 1. Surface inventory — every distinct OOXML emit path in `src/writer.zig`

The following catalogues all 13 distinct emit phases inside
`Writer.save` (`src/writer.zig:665-1233`) plus the row emitter and
styles emitter. Each entry: function (line range) → what it emits →
state it depends on → byte-stability invariants that MUST be
preserved.

### 1.1 `[Content_Types].xml` — `Writer.save:694-739`

- **Emits.** `<?xml…?><Types …><Default …/>… <Override …/>…</Types>`.
- **State.** `Writer.sheets` (count + per-sheet comments presence),
  `have_styles` derived from `Writer.styles` + `Writer.dxfs`.
- **Invariants.**
  - **Every `<Default>` precedes every `<Override>`** (OPC schema requirement; `src/writer.zig:683`).
  - When any sheet has comments, the `Default Extension="vml"` declaration is emitted **before** the fixed `<Override>` block (`50ed225` = "writer: VML Default precedes ALL Overrides").
  - Override `PartName` is the leading-slash-prefixed absolute path (`/xl/worksheets/sheetN.xml`).
  - The fixed `xml` + `rels` Defaults come from `CONTENT_TYPES_DEFAULTS` (`src/writer.zig:65`); the workbook + sharedStrings Overrides come from `CONTENT_TYPES_FIXED_OVERRIDES` (`:71`); `</Types>` close from `CONTENT_TYPES_TAIL` (`:75`).

### 1.2 `_rels/.rels` — `Writer.save:742`

- **Emits.** `ROOT_RELS` static blob (`src/writer.zig:77`).
- **State.** Pure constant. No inputs.
- **Invariant.** Single relationship pointing to `xl/workbook.xml` with `Type=…/officeDocument`, `Id="rId1"`.

### 1.3 `xl/workbook.xml` — `Writer.save:744-782`

- **Emits.** `<workbook …><sheets>…</sheets>[<definedNames>…]</workbook>`.
- **State.** `Writer.sheets` (name + ordinal index), `Writer.defined_names`.
- **Invariants.**
  - `<sheets>` populated in registration order; per-sheet `name`, `sheetId="N"` (1-based), `r:id="rIdN"`.
  - `name` attribute is **XML-attribute-escaped** (sheet names like `R&D` / `x<y` are valid Excel — `src/writer.zig:2706` `appendXmlEscaped` covers `<>&"'`).
  - **`<definedNames>` block sits between `</sheets>` and `</workbook>`** (OOXML schema; `src/writer.zig:758`).
  - `<definedName>` attributes ordering: `name`, optional `localSheetId`, optional `hidden="1"` (writer position-fixed; readers tolerate any order, but byte-stability tests will catch drift).
  - `localSheetId` ≥ `Writer.sheets.len` is rejected at save (`error.InvalidDefinedNameLocalSheetId`).

### 1.4 `xl/_rels/workbook.xml.rels` — `Writer.save:784-811`

- **Emits.** `<Relationships …>(N × worksheet) + sharedStrings + (optional styles)</Relationships>`.
- **State.** `Writer.sheets.len`, `have_styles`.
- **Invariants.**
  - rId numbering: `rId1..rIdN` for sheets in declaration order, `rId(N+1)` for `sharedStrings`, `rId(N+2)` for `styles` (when present).
  - `Type` URLs are the OOXML 2006 strings hard-coded in the format string at `:792`.

### 1.5 Per-sheet `xl/worksheets/sheetN.xml` — `Writer.save:813-1023`

This is the most complex section — the OOXML CT_Worksheet schema fixes the child-element order, and Excel's "repaired" prompt fires on any drift.

- **Emits.** `<?xml…?><worksheet …>` + (`<sheetViews>`?) + (`<cols>`?) + `<sheetData>` + (`<autoFilter>`?) + (`<mergeCells>`?) + (`<conditionalFormatting>`*) + (`<dataValidations>`?) + (`<hyperlinks>`?) + (`<legacyDrawing>`?) + `</worksheet>`.
- **State.** `SheetWriter.{freeze_rows, freeze_cols, column_widths, body, auto_filter_range, merged_cells, conditional_formats, data_validations, data_validation_ranges, hyperlinks, internal_hyperlinks, comments}`.
- **Invariants.**
  - Prolog from `WORKSHEET_PROLOG` (`src/writer.zig:97`) — fixed XML decl + worksheet root with default + r namespace.
  - **CT_Worksheet child order (ECMA-376):** `sheetViews → cols → sheetData → autoFilter → mergeCells → conditionalFormatting → dataValidations → hyperlinks → legacyDrawing`. Reorder = Excel "repaired" prompt.
  - `<sheetView workbookViewId="0">` fixed; pane attribute order: `xSplit, ySplit, topLeftCell, activePane, state="frozen"`. `topLeftCell` computed via `formatCellRef(freeze_rows + 1, freeze_cols)`.
  - `activePane` selection: `"bottomRight"` if both rows + cols frozen, else `"bottomLeft"` for rows-only, `"topRight"` for cols-only.
  - `<col>` attributes: `min, max, width, customWidth="1"` — fixed order.
  - `<sheetData>` body comes from `SheetWriter.body.items` — pre-built by `writeRowImpl` (`src/writer.zig:2007`); SST indices are baked in by then.
  - `<mergeCells count="N">` — count attribute preserved even if N == 1.
  - `<conditionalFormatting>` — one block per rule, `priority` increments per rule (`src/writer.zig:878`); `dxfId` references `Writer.dxfs` ids.
  - `<dataValidations count="N">` — count = list entries + range entries combined; **list entries emitted FIRST**, then numeric/custom range entries (preserves the iter13 ordering — see `:954`).
  - List validations: `formula1` is `&quot;v1,v2,…&quot;` (commas join values; XML-entity-escaped quotes wrap them); embedded commas / bare `"` in values are rejected at intake (`addDataValidationList:1593-1594`).
  - `<dataValidation>` attribute order: `type, [operator], allowBlank="1", showInputMessage="1", showErrorMessage="1", sqref` (fixed).
  - `<hyperlinks>` — external entries (rIds) FIRST, then internal entries (`location="…"`); `r:id` numbering matches the per-sheet rels file (1..N for hyperlinks).
  - `<legacyDrawing r:id="rId{N+2}"/>` only when `comments.items.len > 0`. rId scheme: `1..N` external hyperlinks, `N+1` = comments part, `N+2` = vmlDrawing part.

### 1.6 Per-sheet `xl/worksheets/_rels/sheetN.xml.rels` — `Writer.save:1031-1061`

- **Emits.** Fresh `<Relationships>` doc per sheet that has either `hyperlinks` or `comments`.
- **State.** `SheetWriter.{hyperlinks, comments}`, sheet ordinal.
- **Invariants.**
  - Skipped entirely (no part written) when both lists empty.
  - rId scheme matches the in-sheet `<hyperlinks>` + `<legacyDrawing>` rIds.
  - External hyperlink Target is **XML-attribute-escaped** so `?q=1&x=2` survives (`src/writer.zig:1039`); `TargetMode="External"` always.
  - Comments rel target is `../comments{i+1}.xml`; vmlDrawing rel target is `../drawings/vmlDrawing{i+1}.vml`.

### 1.7 `xl/commentsN.xml` — `Writer.save:1070-1120`

- **Emits.** `<comments xmlns=…><authors>…</authors><commentList>…</commentList></comments>`.
- **State.** `SheetWriter.comments[*].author` (deduped on emit), `SheetWriter.comments`.
- **Invariants.**
  - Authors deduped O(N²) on emit; first occurrence wins on `authorId` numbering.
  - Plain-text comments emit `<text><t xml:space="preserve">…</t></text>` (NO synthetic `<r>` wrapper — see `:1106-1112`: a `<r>` would make the reader treat every Writer-produced comment as rich, breaking the plain/rich contract). When rich-comment support lands later this becomes a per-comment branch.
  - `xml:space="preserve"` on `<t>` is mandatory (preserves leading / trailing / runs of whitespace).

### 1.8 `xl/drawings/vmlDrawingN.vml` — `Writer.save:1127-1193`

- **Emits.** `<xml xmlns:v=… xmlns:o=… xmlns:x=…><o:shapelayout>…</o:shapelayout><v:shapetype id="_x0000_t202" …/>(<v:shape …>×N)</xml>`.
- **State.** `SheetWriter.comments` (one shape per comment).
- **Invariants.**
  - `<o:idmap data="1,2,3,…">` — chunks of 1024 shape IDs each; `num_idmaps = comments.len / 1024 + 1` (over-provision by one is harmless; under-provision = unrendered notes).
  - Shape IDs start at **1025** and increment per comment (`src/writer.zig:1187`); `_x0000_t202` is the canonical text-box shape type.
  - `<x:Anchor>` is 8-tuple `fromCol, 15, fromRow, 2, toCol, 31, toRow, 3`. **Both** `from_col` and `to_col` are clamped to `EXCEL_MAX_COL - 1` (and same for rows) — un-clamped emit produces inverted anchors when the comment sits on column XFD or row 1048576 (`a966e29` callout).
  - VML namespace declarations are URN-form, not http; LibreOffice + Excel both parse only the URN form for the legacy notes layer.

### 1.9 `xl/sharedStrings.xml` — `Writer.save:1196-1215`

- **Emits.** `<sst xmlns=… count="X" uniqueCount="Y"><si>…</si>×Y</sst>`.
- **State.** `Writer.sst_strings` (the dedup pool — owned slices), `Writer.sst_is_rich` (parallel bool array), `Writer.sst_count` (running count of string-typed cells), `Writer.sst_index` (text → idx hashmap).
- **Invariants.**
  - `count` = total string-cell occurrences; `uniqueCount` = `sst_strings.len`.
  - Plain entries: `<si><t xml:space="preserve">…</t></si>` (xml:space mandatory).
  - Rich entries: `<si>{pre-serialised <r>…</r> body}</si>` — the body was built at intern time by `Writer.sstInternRich` (`src/writer.zig:631`).
  - Plain entries are deduped via `StringHashMap`; rich entries are NEVER deduped (rich bodies are rare; hashing the full formatted form costs more than it saves — see `src/writer.zig:626`).
  - **String-content invariant.** Plain entries are escaped through `appendXmlEscaped`; XML 1.0 forbidden control bytes are pre-validated by `assertNoForbiddenXmlBytes` (`:2724`) at `writeRow*` intake — DEL (0x7F) is permitted (`944a2e6`).

### 1.10 `xl/styles.xml` — `emitStylesXml` at `src/writer.zig:2386-2598`

- **Emits.** `<styleSheet …><numFmts…>?<fonts…><fills…><borders…><cellStyleXfs…><cellXfs…><cellStyles…><dxfs…>?</styleSheet>`.
- **State.** `Writer.styles` (per-style font/fill/border/numFmt/alignment), `Writer.num_fmts` + `Writer.num_fmt_index`, `Writer.dxfs`.
- **Invariants (the most byte-fragile axis).**
  - **CT_Stylesheet element order (OOXML schema):** `numFmts → fonts → fills → borders → cellStyleXfs → cellXfs → cellStyles → dxfs`. **`<cellStyles>` MUST sit between `<cellXfs>` and `<dxfs>`** (`7f8cbe3` = "drop duplicate <cellStyles> emission"; the prior bug emitted it twice when dxfs were absent).
  - `<numFmts>` is OMITTED entirely when `num_fmts.len == 0` (built-ins 0..=49 don't appear here); user numFmts start at id **164** (`NUM_FMT_BASE`, `:137`).
  - `<fonts count="N+1">` — default font at index 0 (from `STYLES_FONTS_DEFAULT`), then one per registered style (1:1; deduping fonts independently is a future iter — `:2381`).
  - `<fills count="2 + user_fills">` — slots 0=`patternType="none"`, 1=`patternType="gray125"` are reserved (conventional OOXML defaults), then one per style with non-default fill.
  - `<borders count="1 + user_borders">` — slot 0 is the empty default `<border><left/><right/><top/><bottom/><diagonal/></border>`, then one per style with any border field set.
  - `<cellStyleXfs>` is the static `STYLES_CELL_STYLE_XFS` blob (`:124`).
  - `<cellXfs count="N+1">` — slot 0 = default no-style (`STYLES_DEFAULT_CELL_XF`), then one per style. **`addStyle` returns 1-based index** so callers feed it directly to `s="N"`.
  - `<xf>` attribute order (fixed): `numFmtId, fontId, fillId, borderId, xfId="0", applyFont="1", [applyNumberFormat], [applyFill], [applyBorder], [applyAlignment]` + optional `<alignment horizontal=… wrapText=…/>`.
  - `<cellStyles count="1">` — single `<cellStyle name="Normal" xfId="0" builtinId="0"/>` (`STYLES_CELL_STYLES`, `:130`); strict-mode validators reject styles.xml without it.
  - `<dxfs>` block (when present) — emitted LAST per the schema; `<font>`, `<fill>`, `<border>` children only when at least one field differs from default.
  - **Color encoding.** Always 8-hex ARGB (`{X:0>8}` format) — drop the leading hash; theme indices are not used.

### 1.11 ZIP layout — `ZipWriter` at `src/writer.zig:3265-3429`

- **Emits.** Sequential LFH + payload phase, then CDFH phase, then EOCD record. ZIP32 only (no Zip64).
- **State.** `ZipWriter.entries` (per-entry name, sizes, offsets, CRC32, compression method).
- **Invariants.**
  - **Per-entry compression policy.** Payloads under 1 KB go uncompressed (method 0); ≥ 1 KB go through `deflateCompress` with method 8 (the writer's hand-rolled DEFLATE; see `:2848-3263`). Stored fallback if the compressed output is bigger than the input.
  - Method 8 DEFLATE: 32 KB sliding window, single-step lazy matching, hash-3 anchor (DEFLATE_HASH_BITS=15), Huffman frequency cap **u13** (raised from u15 → u13 in `410b4bc` to unblock 100 k-row deflates that overflowed the stdlib HuffmanEncoder's u16 internal accumulator).
  - **CRC32** computed over the uncompressed bytes (per ZIP32 spec).
  - **No data descriptors** — sizes / CRC live in the LFH.
  - LFH version-needed: 20 (DEFLATE) or 10 (stored). General-purpose flag: 0 (no encryption, no data descriptors, no UTF-8 name flag).
  - **EOCD comment.** Empty (`comment_length=0`); `disk_number=0`; `central_directory_disk_number=0`.
  - **ZIP32 sentinel guards.** Every serialized u32 size / offset field rejected at `0xFFFFFFFF` (Zip64 sentinel — `66e4ccd`); CD-size + CD-offset re-checked AFTER writing the CD (`962fc86`); the writer guarantees its own output is readable by zlsx's reader (`d6235f3` total-size guard).
  - **Entry-name policy.** ASCII / UTF-8 raw bytes (no general-purpose-bit-11 UTF-8 flag); slashes are forward; no leading slash.

### 1.12 Row emit — `writeRowImpl` at `src/writer.zig:2007-2119`

- **Emits.** `<row r="N"[ ht="…" customHeight="1"]>(<c …>…</c>×M)</row>` into `SheetWriter.body`.
- **State.** `SheetWriter.{next_row, row_heights, body}`, `Writer.{sst_strings, sst_index, sst_is_rich, sst_count}`.
- **Invariants.**
  - **Row atomicity.** Pre-validates the entire row (Excel hard limits `EXCEL_MAX_ROW=1_048_576`, `EXCEL_MAX_COL=16_384`; integers via `fitsExactlyInF64`; finite f64; no XML-1.0 forbidden bytes via `assertNoForbiddenXmlBytes`) **BEFORE** any byte is appended to `body`. A failed row leaves `body` untouched (`350a50d`, `7f8cbe3`-era robustness work).
  - **Empty-cell elision.** `cell == .empty && style_id == 0 && formula == null` → cell omitted entirely (OOXML treats missing cells as empty); empty + styled OR empty + formula → emitted with appropriate attributes.
  - **Self-closing `<c r=… s=…/>`** for empty + styled (preserves byte-for-byte output across the formula API addition — `:2061`).
  - **Type attributes.** `t="s"` for SST entries; `t="str"` for formula-cached strings (NOT `t="s"` — Excel marks formula string caches inline, not via SST; `:2073`); `t="b"` for booleans; numeric / blank cells emit no `t` attribute.
  - Cell attribute order: `r, [s], [t]` (fixed).
  - Formula emission: `<f>{xml-escaped}</f><v>{value}</v>` — value is the cached result Excel displays until recalc; for `.empty` formula cells the `<v>` is omitted.
  - Number-cell emit uses `{d}` formatting on f64 — Zig's default (round-trip-stable but not necessarily byte-identical across Zig versions; this is a known fragility — see Section 6 risks).
  - Boolean emit: `<v>0</v>` or `<v>1</v>`.

### 1.13 Rich-row emit — `writeRichRow` at `src/writer.zig:1942-2005`

- **Emits.** Same `<row>…</row>` shape as `writeRowImpl`, but each `.rich` cell pre-builds a `<si>{<r><rPr/>...<t>...</t></r>×K}</si>` body via `Writer.sstInternRich` (`:631`).
- **State.** Same as `writeRowImpl`; rich entries always append to SST (no dedup, see 1.9).
- **Invariants.**
  - `<rPr>` block emitted only when at least one prop is set (bold / italic / size / color / font_name).
  - Run-property order inside `<rPr>`: `<b/> <i/> <sz val=…/> <color rgb=…/> <rFont val=…/>` (fixed).
  - `<t xml:space="preserve">` always.
  - Rich-run text + font_name validated for forbidden XML bytes BEFORE any append (`fbf7f60`, `20fe2b8`).

### 1.14 Helpers used during emit (free wins for wr-6)

| Helper | Writer location | Workbook equivalent | Note |
|---|---|---|---|
| `appendXmlEscaped` | `src/writer.zig:2706` | `pkg/workbook.zig:2173`; `pkg/store.zig:416` | **3× duplicated**. wr-6 unify into `pkg/typed_parts/` or `ziglib`. |
| `appendXmlEscapedText` (text variant — no apostrophe escape) | (inlined in writer's `appendXmlEscaped`) | `pkg/workbook.zig:4391` | Workbook splits attr vs text; Writer single-mode. wr-6 align. |
| `formatCellRef` | `src/writer.zig:2124` | (re-exported `src/xlsx.zig:6835`; consumed at `pkg/editor.zig:600`) | Already re-exported — single source of truth. wr-2/4 keep using it directly. |
| `validateSheetName` | `src/writer.zig:2210` | `pkg/workbook.zig:2009` (and `src/xlsx.zig:6833` re-exports writer's) | **2× duplicated**. wr-6 collapse to writer's (the canonical surface). |
| `validateMergeRange` | `src/writer.zig:2172` | (none — Workbook has no merge fresh-emit) | wr-4 lift into Workbook when merge-cells gain typed-fresh-emit. |
| `validateAutoFilterRange` | `src/writer.zig:2320` | (none) | Same — wr-4. |
| `validateHyperlinkRange` | `src/writer.zig:2331` | (none) | Same — wr-4. |
| `validateDefinedName` | `src/writer.zig:2242` | (none — Workbook reads names but doesn't fresh-emit) | wr-3 lift. |
| `looksLikeR1C1Ref` / `looksLikeCellRef` | `src/writer.zig:2265-2319` | (none) | Defined-name validators; wr-3. |
| `asciiEqlFold` | `src/writer.zig:2189` | (none — Workbook uses `sheetNameMatchesDecoded` w/ entity-decode) | wr-3 reconcile. Note: writer uses Unicode `casefold.excelSheetNameEql` for full duplicate detection. |
| `assertNoForbiddenXmlBytes` | `src/writer.zig:2724` | (none — readers don't validate; Workbook trusts source bytes + `setCell` deltas don't pass through here) | wr-1/4 will need this for the SST + sheetData emit paths. |
| `deflateCompress` | `src/writer.zig:3244` (file-scope `pub fn`) | already consumed at `pkg/store.zig:262, 390, 453` and `pkg/editor.zig:79` | **Already shared.** wr-5 will route the new ZIP-path through this same function. |
| `fitsExactlyInF64` | `src/writer.zig:52` | (Workbook integer-cell delta path uses raw f64 cast) | wr-4 share when integer cells reach Workbook fresh-emit. |
| `ZipWriter` (struct) | `src/writer.zig:3265` (private) | `pkg/store.zig:502` (`PartStore.save`) emits LFH+CDFH+EOCD inline | **Architectural fork.** wr-5 retires `ZipWriter` in favour of a `PartStore.fromScratch()` mode. |

## 2. Pin golden byte outputs

The following tests / fixtures encode byte-stability invariants the
rebase MUST not break. They are the wr-7 parity gate.

### 2.1 In-tree writer tests — `src/writer.zig` (all run via `zig build test`'s `writer_tests` step at `build.zig:112`)

Total: **66 tests** in writer.zig. The byte-asserting subset (those
that go beyond round-trip-via-reader and check raw XML / ZIP bytes)
is the parity surface. Selected examples:

| Test | Line | What's pinned |
|---|---|---|
| `formatCellRef A1, B2, Z1, AA1, AAA1` | 3470 | A1 column-letter math |
| `appendXmlEscaped covers all 5 entities` | 3479 | `< > & " '` → entity form |
| `empty workbook fails with NoSheets` | 3486 | early-exit branch |
| `single-sheet round-trip via zlsx reader` | 3496 | end-to-end shape |
| `multi-sheet round-trip + SST dedup` | 3562 | SST dedup invariant |
| `xml entities in strings are escaped` | 3597 | per-cell escape |
| `writeRowStyled rejects out-of-range style id` | 3619 | refusal contract |
| `stage-5 number format registers + emits numFmts` | 3643 | numFmt slot 164+ |
| `writeRowWithFormulas emits <f> + cached <v>` | 3702 | formula cell shape |
| `setRowHeight emits ht + customHeight, only on marked rows` | 3784 | row-height delta emit |
| `stage-5 sheet-level features (cols, freeze, autoFilter)` | 3853 | CT_Worksheet child order |
| `addMergedCell validates + emits <mergeCells> block` | 3921 | mergeCells |
| `addDataValidationNumeric + Custom emit correct XML` | 4001 | DV ranges |
| `VML idmap expands for >1023 comments per sheet` | 4090 | idmap chunking |
| `comment on XFD column emits non-inverted VML anchor` | 4175 | anchor clamp (`a966e29`) |
| `conditional formatting — colorScale (2+3 stop) + dataBar` | 4225 | CF stops |
| `conditional formatting — cellIs + expression rules + dxfs table` | 4310 | dxfs ordering |
| `addDataValidationList validates + emits <dataValidations> block` | 4427 | DV list |
| `addDataValidationList — no block when none registered` | 4499 | optional-block elision |
| `addHyperlink validates + emits <hyperlinks> + per-sheet _rels` | 4535 | rId numbering |
| `no <hyperlinks> block or _rels entry when none registered` | 4635 | rels skip |
| `no <mergeCells> block when none registered` | 4668 | optional-block elision |
| `stage-4 border sides emit into styles.xml` | 4705 | border layout |
| `stage-3 fill fields emit into styles.xml` | 4778 | fill layout |
| `stage-2 style fields emit into styles.xml` | 4856 | style layout |
| `addSheet validates sheet names …` | 5039–5179 | sheet-name rules |
| `reject only integers that round on IEEE-754 conversion` | 5205 | `fitsExactlyInF64` |
| `writeRow is atomic on IntegerExceedsExcelPrecision` | 5233 | row atomicity |
| `random stage 2-5 style combos survive round-trip` | 5589 | style fuzz |
| `random stage-5 per-sheet feature combos` | 5669 | per-sheet feature fuzz |
| `random op ordering with invariants` | 5971 | op-order fuzz |
| `multi-save preserves all prior rows` | 6083 | save idempotence |
| `boundary numeric values survive round-trip` | 6131 | f64 boundary |
| `addDefinedName accepts/rejects valid names` | 6303–6398 | name-rule validators |
| `forbidden-XML-byte string leaves no half-written row` | 6412 | atomicity |
| `font_name with XML specials, alignment, wrap, diagonals` | 6475 | font_name escape |
| `large repetitive payload doesn't trip Huffman assert` | 6528 | u13 freq scale (`410b4bc`) |
| `fuzz Writer end-to-end round-trip via reader` | 5515 | end-to-end fuzz |
| `fuzz ZipWriter produces archives our reader can walk` | 5802 | ZIP round-trip |

The `expectEqualStrings` / `expectEqualSlices` assertions across
these tests pin specific raw XML byte fragments — these are the
golden byte outputs the rebase MUST replicate.

### 2.2 Bench harnesses (perf gates)

- `tests/bench/bench_write_zlsx.zig` — 1 k rows × 10 cols, target **6.7 ms ± 0.3** (`docs/benchmarks.md:145`). wr-7 gate: ≤ 1.10× = 7.4 ms.
- `tests/bench/synth_100k_x_10.zig` — 100 k-row stress. wr-7 secondary gate.
- `tests/bench/bench_write_openpyxl.py` + `bench_write_xlsxwriter.py` — cross-tool reference (informational, not gating).
- `tests/bench/workbook_rss.zig` — RSS profile (would catch streaming-write regressions if wr-7 turns up memory pressure).

### 2.3 Editor consumer tests (B2 closed those — Writer doesn't touch them, but they assert end-to-end shape)

`pkg/editor.zig:1160`+ has ~30 tests that build a Writer, save, and
re-open. These remain valuable as a producer round-trip canary; the
rebase MUST keep them green (they continue to call `Writer.init`
through wr-6's frozen public API).

## 3. Workbook gaps to close

For `Writer.save` to become a thin shim over `Workbook.save`,
Workbook needs new emit capabilities. Today Workbook handles
`setCell` / `appendRows` / `addSheet` / `deleteSheet` / `renameSheet`
/ `addImage` / `rewriteAll*` (formulas, hyperlinks, defined names,
DV, CF) — but its mutation model is **delta-on-existing-bytes**, not
fresh-file producer.

| Gap | Workbook today | Smallest API addition needed |
|---|---|---|
| Fresh `<conditionalFormatting>` emit | `rewriteAllValidationsAndConditionalFormats` mutates existing | `Worksheet.addConditionalFormat*` accepting the same struct shapes Writer takes; emit slots into `<sheetData>`-following CT_Worksheet position via the delta engine or a new "fresh sheet body" mode |
| Fresh `<dataValidations>` emit | same — rewriter only | `Worksheet.addDataValidation{List,Numeric,Custom}` |
| Fresh `<hyperlinks>` + per-sheet rels | `rewriteAllHyperlinkLocations` mutates existing | `Worksheet.addHyperlink` / `addInternalHyperlink` + a way to extend the sheet's `.rels` part on save |
| Fresh `<autoFilter>` | none | `Worksheet.setAutoFilter` |
| Fresh `<cols>` | none | `Worksheet.setColumnWidth` (multi-call appends, like Writer) |
| Fresh `<sheetView>` panes | none (frozen-pane refusal still in place per `editor-rebase.md` iter-er-5) | `Worksheet.freezePanes` — and the per-sheet-pane refusal lifts naturally because Workbook now owns the emit |
| Fresh `<mergeCells>` | none | `Worksheet.addMergedCell` |
| Comments + VML drawing | none | `Worksheet.addComment` + a Workbook-side VML emitter (port `Writer.save:1127-1193` into `pkg/`) |
| Rich strings (`RichTextRun`) in fresh emit | `CellValue.shared_string` is plain-only; rich entries pre-rendered into SST by Writer | Extend `SstExtensionPlan` with a `rich` axis OR add `CellValue.rich: []const RichTextRun` |
| Workbook-level `<definedNames>` fresh emit | reads via `definedNames()` but doesn't accept fresh adds | `Workbook.addDefinedName(name, refers_to, opts)` + name-rule validation lift from `src/writer.zig:2242` |
| Per-cell number-format / style indexing on fresh emit | `Worksheet.setCell` doesn't take a style id | `Worksheet.setCellStyled(ref, value, style_id)` OR per-row `setRow` helper accepting `[]const u32 styles` (mirrors `writeRowStyled`) |
| Style registry (`addStyle` / `addDxf` / numFmt pool) | `pkg/typed_parts/styles_xml.zig` is parse-only today | New `Workbook.{addStyle, addDxf, internNumFmt}` writing into `styles_xml_mod` typed state |
| Auto-author dedup for comments | none | Folds into the comments emit gap — Workbook-side comments emitter dedups authors O(N²) like Writer does |
| Atomic row writes (atomicity contract from `writeRowImpl`) | `Worksheet.setCell` per-cell, no batch | `Worksheet.appendRowsAtomic(rows)` — already in flight via `Worksheet.appendRows` (`pkg/workbook.zig:3900`); reuses |

The largest single gap is **styles unification**: Writer's `Style`
/`Dxf`/`Border`/`Fill`/`Font`/`HAlign`/`PatternType`/`BorderStyle`
need to either move into `pkg/typed_parts/styles_xml.zig` or have a
typed bridge with bidirectional mapping. Section 5's wr-2 owns this.

## 4. Proposed iter-wr-1..7 plan

Refines the wr-0..7 sketch (memory file
`project_iter_er_3_pre_work.md:59-67`) with concrete file/function
targets per iter. **wr-2 (styles) and wr-4 (sheet emit) are the two
big iters** — wr-2 because styles is the most byte-fragile axis;
wr-4 because `<sheetData>` emit is the perf-critical path.

### iter-wr-1 — SST unification (1 week)

**Scope.** Retire `Writer.{sst_strings, sst_index, sst_is_rich, sst_count, sstIntern, sstInternRich}` in favour of `Workbook.SstExtensionPlan` (`pkg/workbook.zig:buildSstExtensionPlan`). Writer's row emitter routes string interns through Workbook's plan; rich-text becomes a new `SstExtensionPlan.rich` axis.

**Files.** `src/writer.zig` (gut SST state); `pkg/workbook.zig` (extend `SstExtensionPlan`); `pkg/typed_parts/sst_xml.zig` (rich-emit helper).

**Walk-away.** SST round-trip + rich-string round-trip parity tests green; bench within 1.10×.

### iter-wr-2 — Styles unification (2 weeks) — BIG ITER

**Scope.** Move `Style`, `Dxf`, `BorderSide`, `BorderStyle`, `PatternType`, `HAlign` types from `src/writer.zig` to a new home in `pkg/typed_parts/styles_xml.zig`. Workbook gains `Workbook.{addStyle, addDxf, internNumFmt}`. `emitStylesXml` (`src/writer.zig:2386`) ports to `styles_xml_mod.emit`. Writer's `addStyle` / `addDxf` become pass-throughs.

**Files.** `src/writer.zig` (gut styles state, port `emitStylesXml` to typed-parts); `pkg/typed_parts/styles_xml.zig` (extend with fresh-emit); `pkg/workbook.zig` (new pub fns).

**Walk-away.** All 66 in-tree writer tests green; styles-byte-equivalence test added — diff compares Writer-saved vs Workbook-saved styles.xml byte-for-byte.

**Risk callout.** Styles is the most byte-fragile surface (CT_Stylesheet child order; numFmt id 164 base; fill slots 0+1 reserved). One missed attribute order = "repaired" prompt across the corpus.

### iter-wr-3 — workbook.xml unification (1 week)

**Scope.** Retire `Writer.defined_names` + the workbook.xml emit branch in `Writer.save` (`:744-782`). Workbook gains `Workbook.addDefinedName(name, refers_to, opts)` lifting `validateDefinedName` from `src/writer.zig:2242`. workbook.xml fresh-emit moves into `pkg/typed_parts/workbook_xml.zig`.

**Files.** `src/writer.zig`; `pkg/workbook.zig`; `pkg/typed_parts/workbook_xml.zig`.

**Walk-away.** Defined-name corpus tests green; workbook.xml byte-equivalence test added.

### iter-wr-4 — Sheet emit unification (2-3 weeks) — BIG ITER (PERF-CRITICAL)

**Scope.** Move per-sheet emit (`Writer.save:813-1023` + `writeRowImpl`/`writeRichRow` body builders) into Workbook. Workbook gains the gap-list from Section 3: `Worksheet.{addConditionalFormat*, addDataValidation*, addHyperlink, addInternalHyperlink, setAutoFilter, setColumnWidth, setRowHeight, freezePanes, addMergedCell, addComment, setCellStyled}`. Writer's `SheetWriter` becomes a thin facade.

**Files.** `src/writer.zig` (huge cut — most of `SheetWriter`'s body emit moves out); `pkg/workbook.zig`; `pkg/typed_parts/sheet_xml.zig` (extend with fresh-emit).

**Walk-away.** **Bench gate strict** (≤ 1.10× of 6.7 ms = 7.4 ms ceiling) — `<sheetData>` emit is the hot loop; any regression here breaks Phase 3b. CT_Worksheet child-order parity test added.

### iter-wr-5 — ZIP emit unification (1-2 weeks)

**Scope.** Retire `Writer`'s private `ZipWriter` (`src/writer.zig:3265`) in favour of a new `PartStore.fromScratch()` mode that lets Writer hand a list of `(name, content_type, bytes)` tuples and gets back a saveable PartStore. `deflateCompress` is already shared (`pkg/store.zig:262`).

**Files.** `src/writer.zig` (delete `ZipWriter` ~165 LOC); `pkg/store.zig` (extend with `fromScratch`).

**Walk-away.** ZIP round-trip fuzz test (`tests/.../fuzz`) green; existing 100 k-row deflate stress test (`410b4bc` regression) green.

### iter-wr-6 — `Writer.save` thin shim + helper dedup (1 week)

**Scope.** `Writer.save` becomes:

```zig
pub fn save(self: *Writer, path: []const u8) !void {
    var wb = try self.toWorkbook();  // populates a Workbook from the staged state
    defer wb.deinit();
    return wb.save(path);
}
```

Retire ~5500 LOC of dead Writer helpers (the per-emit functions
moved out in wr-1..5). Collapse the duplicated helpers from Section
1.14: one `appendXmlEscaped`, one `validateSheetName`, one
`asciiEqlFold` strategy.

**Files.** `src/writer.zig` (final cleanup pass — should drop from 6579 LOC to ~700-1000 LOC of facade + tests); `pkg/workbook.zig` (canonical helpers). `pkg/store.zig` (canonical `appendXmlEscaped`).

**Walk-away.** All 66 in-tree writer tests green; line-count target ≤ 1100 LOC.

### iter-wr-7 — Corpus parity sweep + perf bench (1 week)

**Scope.** End-to-end corpus sweep — for every writer test, capture
pre-rebase output bytes (against tagged commit `8XXXX` = pre-wr-1)
and assert post-rebase output is byte-equivalent (modulo XML
whitespace inside elements). Bench: publish before/after table for
1 k × 10 / 10 k × 10 / 100 k × 10. Fix any regression > 1.10× under
a feature flag or revert the offending iter.

**Walk-away.** ≤ 1.10× perf gate hold; corpus parity at 100%.

## 5. Walk-away gates (summary)

| Iter | Gate | Failure mode |
|---|---|---|
| wr-1 | SST round-trip parity (plain + rich); bench ≤ 1.10× | revert to Writer-local SST; document divergence |
| wr-2 | styles.xml byte-equivalence; all 66 tests green | freeze and debug schema-order regression |
| wr-3 | defined-name corpus parity | retain Writer-local `defined_names` until name-rule validator parity proven |
| wr-4 | bench ≤ 1.10× = 7.4 ms ceiling; CT_Worksheet child order parity | retain Writer-local sheet emit; ship rebase as opt-in |
| wr-5 | ZIP fuzz green; 100 k-row deflate stress green | retain `ZipWriter` until `PartStore.fromScratch` proven |
| wr-6 | line-count ≤ 1100 LOC for `src/writer.zig`; helper dedup complete | revert per-helper if dedup breaks any test |
| wr-7 | end-to-end corpus parity 100%; perf ≤ 1.10× across 1 k / 10 k / 100 k | feature-flag offending axis; ship incremental |

## 6. Risks

- **Phase 3b perf regression.** Writer is the fastest production path
  — `tests/bench/bench_write_zlsx.zig` measures **6.7 ms** for 1 k ×
  10 (`docs/benchmarks.md:145`); any rebase that goes past 1.10× =
  7.4 ms breaks the production gate. wr-4 is where this happens (or
  doesn't); the fix is a fast-path inside the new
  `Worksheet.appendRowsAtomic` — same shape as iter-er-3's substring
  fast-path, but for fresh emit.

- **Byte-stability across the 0.2.9 invariant set.** The following
  invariants are pinned by tests AND by reader-side parity (Excel,
  openpyxl, calamine all assume them). One missed = "repaired"
  prompt or silent corruption:
  - **`<cellStyles>` between `<cellXfs>` and `<dxfs>`** — `7f8cbe3` (drop duplicate cellStyles emission).
  - **VML `Default Extension="vml"` precedes ALL `<Override>`** in `[Content_Types].xml` — `50ed225`.
  - **`<definedNames>` between `</sheets>` and `</workbook>`** — `src/writer.zig:758` (per OOXML schema).
  - **Defined names reject R1C1 shape, accept A1-shape only** — `350a50d`.
  - **Defined names reject case-insensitive duplicates per scope** — `490f0b8`.
  - **CT_Worksheet child order** — `src/writer.zig:861-1018` ordering preserved.
  - **VML anchor clamp on XFD column / row 1048576** — `a966e29`, `2c4a3d4`.
  - **DEL (0x7F) is XML 1.0 legal** — `944a2e6`.
  - **Huffman freq cap u13** — `410b4bc` (100 k-row deflate fix).
  - **ZIP32 sentinel guards on serialized fields** — `a0a84a2`, `66e4ccd`, `962fc86`, `d6235f3`.
  - **freezePanes split must leave a visible pane** — `2c4a3d4`.

- **Streaming write.** Writer + Workbook are **both fully buffered**
  today — every part is built in `std.ArrayListUnmanaged(u8)` and
  the whole archive is held in memory before flush. If wr-7's RSS
  bench surfaces memory pressure on 100 k-row workbooks (or the
  10 k-sheet stress test), a streaming variant becomes part of the
  rebase scope. For now, treat as wr-7 risk only.

- **f64 formatting fragility.** Number cells emit via Zig's `{d}`
  formatter which is round-trip-stable but not byte-identical
  across Zig stdlib versions. The toolchain pin (Zig 0.15.2, see
  `~/Projects/Pro/CLAUDE.md`) holds for now, but a future Zig bump
  could shift number serialisation. Note for wr-7 — the parity
  diff must tolerate per-byte numeric reformatting OR the gate
  becomes "open in Excel + read back the f64s + assert ≈".

- **c_abi byte-compat.** `src/c_abi.zig:1403`+ exposes
  `zlsx_writer_*` exports consumed by `py-zlsx`. Any error-code
  drift breaks downstream callers. wr-6 must not change error
  enum ordinal values.

- **Rich-text dedup policy divergence.** Writer never dedups rich
  entries (`src/writer.zig:626`); Workbook's SST plan would
  naturally dedup. The dedup rule must be retained verbatim or
  the SST output bytes drift.

## 7. Open questions

1. **Where do Workbook-side fresh-emit constructs live — `pkg/workbook.zig` or new `pkg/typed_parts/sheet_emit.zig`?** The typed_parts subdirectory is parse-focused today (`sheet_xml.zig` is reader-only). wr-4 needs a "fresh-emit" home; a new module avoids ballooning workbook.zig past its 330 KB.

2. **Should `CellValue` gain a `.rich: []const RichTextRun` variant or should `SstExtensionPlan` own the rich axis end-to-end?** First option is symmetric with `.shared_string` but bloats the union; second keeps the union flat but forces rich entries through a separate Worksheet method.

3. **Style id stability across rebase.** Writer's `addStyle` returns 1-based ids; the new Workbook surface must preserve that contract for c_abi compat. Does `Workbook.addStyle` also return 1-based? (Recommended: yes.)

4. **NumFmt id 164 base.** `NUM_FMT_BASE = 164` is hard-coded in writer; Workbook-side typed-parts has no such constant today. Is this the canonical home (`pkg/typed_parts/styles_xml.zig`) or should it live in a new shared `pkg/typed_parts/num_fmt.zig`?

5. **Per-sheet rels file emit on fresh-write workbooks.** Workbook today touches existing `xl/worksheets/_rels/sheetN.xml.rels` via `replacePart`; on a fresh-write the file doesn't exist. wr-4 needs `PartStore.addPart` to be the path — does it support the `xl/worksheets/_rels/` directory layout? (Spot-check in wr-0: yes, `PartStore.addPart` takes any name; the rels-cache refresh in PR #74 covered the lookup-after-add case.)

6. **Comment + VML emit ownership.** wr-4 ports the VML drawing emitter (`src/writer.zig:1127-1193`) into Workbook. Should it live next to drawings (`pkg/drawing_emit.zig`) or in a new `pkg/vml_emit.zig`? VML legacy notes are distinct from `xl/drawings/drawingN.xml` (the modern picture/chart drawings); `drawing_emit.zig` is currently picture-focused.

7. **wr-1 vs wr-3 ordering.** SST first vs workbook.xml first — the sketch lists SST first, but workbook.xml is simpler (~40 LOC of emit) and could ship as a confidence-builder before wr-2's styles risk. Re-confirm before starting.

8. **Bench-baseline tagging.** wr-7 needs a pre-rebase output snapshot to diff against. Tag commit at start-of-wr-1 (`v0.X.Y-pre-writer-rebase`)? Capture bytes for every test fixture into `tests/bench/golden/` as a one-shot at wr-0 close? (Recommended: tag + capture at wr-0 commit.)

9. **Streaming write scope creep.** If wr-7 surfaces memory pressure, do we (a) ship streaming as wr-8, or (b) revert wr-1..7 and rebuild on a streaming base? The cheap answer is (a); the architecturally clean answer is (b). User's "always pick the long-term option" preference (memory file `feedback_long_term_decisions.md`) suggests (b) — but only if wr-7 actually finds the regression.

10. **Editor + Writer co-existence during rebase.** B2 closed Editor as a thin shim over Workbook. During wr-1..7, Writer is mid-rebase. Are Editor's c_abi tests at risk if `Workbook.save` evolves under both? Mitigation: wr-1..6 each preserve `Workbook.save` semantics for the existing `setCell` / `appendRows` paths; only the *fresh-emit* surface is new.

## Out of scope (explicit)

- **C2b drawing-anchor rewriter** — independent of B3; lifts the per-sheet drawing/picture refusal (Editor side, not Writer).
- **C1 M2 m4 + later formula evaluator** — never touched by B3.
- **Streaming write surface** — risk only; full streaming is potential wr-8 / B-stream.
- **Coverage-guided fuzzing of write path** — B-fuzz is parallel.
- **Public field cleanup on Writer / SheetWriter** — current minor line keeps everything; deprecation lives in the next minor.
