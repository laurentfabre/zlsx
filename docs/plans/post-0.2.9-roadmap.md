# zlsx post-0.2.9 roadmap

Synthesises three rounds of Codex deep-research and two rounds of
Codex critique on the seven weaknesses surfaced in the project
assessment. Material changes through revisions: A1 reclassified as
medium-effort (Unicode subsystem), B-tier split into four sub-tiers
(B0–B3), C2 split into C2a (extract/replace) + C2b (addImage) and
pulled forward of formula work, chart emit demoted to Tier D,
walk-away gates tightened (grammar-class instead of corpus-percentage,
report-only bench CI).

## Status (as of 2026-05-03)

Tier-A, the Tier-B keystone (B0), and **B1 Workbook overlay are
fully shipped** — every iter (wb-1 typed parsers, wb-2 read-only
roots, wb-3 `Workbook.fromBook`, wb-4 setCell over 5 CellValue
variants + Workbook.save, wb-5 view ergonomics, wb-6 RSS gate)
landed across PRs #21-#31. iter-wb-6's RSS ratio went from 44× →
0.78× (Workbook now uses LESS resident memory than `Book.openLazy`)
via two perf passes: PR #30 deferred decompression, PR #31 dropped
the file slurp in favour of seek+readAll on demand.

C2a is fully shipped (image + chart anchors + namespace-aware
drawing parser — multi-prefix bindings, full XML scope tracking
for closed siblings, comment / CDATA / PI awareness, quote-aware
attribute scanning, O(n) flat under adversarial input). addPart
shipped on PartStore. C1 M1 (formula tokenizer + loss-preserving
printer) shipped. **C1 M2 m1** (pure `rewriteFormula` — PR #32),
**m1.5 Workbook wiring** (`rewriteAllFormulas` — PR #34),
**m1.6 `Workbook.renameSheet`** (PR #35), and **m2 DV / CF
rewriter with full splice persistence** (`rewriteAllValidationsAnd
ConditionalFormats` with `target_sheet` parameter — PR #48) all
landed. **m3 defined-names + hyperlinks rewriter is still pending**
— a fresh-author attempt is recommended; the prior subagent
attempts on 2026-05-02/03 were closed for cascading rebase
conflicts on `pkg/workbook.zig`, but the agent's logic (target_sheet
API + collect-then-splice fix for the sheet-scope panic) is
referenced in the closed PRs' comments.

Convenience surface added in this batch:
- `Worksheet.cellStyle(ref)` composite accessor (font / fill /
  border / alignment / number-format-code) — PR #39
- `Worksheet.deleteCell(ref)` via new `CellValue.deleted` arm —
  PR #49
- `Workbook.numberFormatFor(style_idx)` with the well-known
  built-in numFmt table (0-4, 9-22, 37-40, 45-49) — PR #51
- `Workbook.hasUnsavedChanges()` + `PartStore.hasUnsavedChanges()`
  predicates — PR #56

Test surface extended:
- Python binding round-trip + lifecycle tests (`Editor` C ABI) —
  PR #42

Plans authored:
- `docs/plans/editor-rebase.md` — B2 plan (PR #41)

Re-authored on stable main and in flight as draft PRs (2026-05-04),
all locally verified via `zig test -target aarch64-macos-none …`
(workaround for the macOS 26.4 SDK arm64e-only libSystem.B.tbd
incompatibility with zig 0.15.2 host-SDK detection; CI runners
unaffected):
- `Workbook.deleteSheet` — draft PR #58 (127/127 tests pass)
- C1 M2 m3 redo — draft PR #59 (128/128 tests pass; collect-then-
  splice pattern fixes the PR #37 sheet-scope panic)
- C2b minimal `addImage` — draft PR #60 (137/137 tests pass; new
  `pkg/drawing_emit.zig` carries the OOXML emit helpers)

The `resolved_part_name` leak in `Worksheet.ensureParsed` (every
`parsed = null → ensureParsed` cycle leaked the prior dupe) was
fixed in PR #38 alongside the prior doc refresh.

Greenfield writer surface is feature-complete for non-image
workbooks: cells, formulas, styles, validations, conditional
formats, comments, hyperlinks, defined names (workbook + sheet-
scoped via `Writer.addDefinedName` with full Excel name-rule
validation incl. R1C1 / A1-shape / case-insensitive duplicate
rejection), freeze panes (legacy clamp + checked typed-error
variant), merged cells, auto-filter (A1-validated), column
widths, row heights (validated against the (0, 409.5] cap).
OOXML schema-correct on emit (Defaults-before-Overrides in
[Content_Types].xml, `<cellStyles>` between `<cellXfs>` and
`<dxfs>`, string-cached formulas use `t="str"`), atomic on error
(no half-written `<row>`), XML-1.0-forbidden bytes rejected at
every user-text channel (cells, formulas, comments, URLs, rich
runs, font names). Reader round-trips every Style field zlsx's
own writer emits — `Book.cellAlignment`, diagonal direction
flags, entity-decoded font names, and `xsd:boolean "true"` form
all land. C ABI + Python bindings expose the full surface with
typed errors and FFI-narrowing bounds checks (no silent uint32
wraparound).

| Item | Status |
|---|---|
| **A1** Unicode sheet-name dedup (case-fold + NFC + char-length) | ✅ shipped |
| **A2** Windows runtime tests | ✅ shipped (continue-on-error pending green-streak gate) |
| **A3** Benchmark-regression CI (report-only) | ✅ shipped |
| **B0 M1** PartStore read-side | ✅ shipped |
| **B0 M2** PartStore byte-preserving save + replacePart | ✅ shipped (incl. data-descriptor preservation) |
| **B0 M2.5** replacePart deflate-encoded overrides | ✅ shipped |
| **B0 addPart** Append a new part with content-type registration | ✅ shipped (atomic on allocation failure, XML-escapes attribute values, stays sentinel-safe) |
| **B0 M3** Typed overlays for known parts | ✅ shipped — consolidated into B1 iter-wb-1 (PR #21) |
| **B0 hardening** ZIP32 sentinel safety, eager CRC32, CDFH bounds, ZIP-bomb caps, split-archive rejection, XML-entity round-trip in [Content_Types] + .rels (named + numeric), external-rel filter | ✅ shipped |
| **B0 perf** Lazy decompress + file-streaming | ✅ shipped (PRs #30 + #31 — `PartStore.open` no longer slurps the file or eagerly decompresses; payloads stream from disk via seek + readAll on first `part(name)` access) |
| **B1** Workbook typed overlay | ✅ shipped — [plan](workbook-overlay.md), PRs #21–#28, #30, #31. RSS gate green at 0.78× (ceiling 1.5×). |
| **B2** Editor rebase onto Workbook | ✅ shipped 2026-05-09 — [plan](editor-rebase.md). All iters landed: iter-er-0 (#61, #65); iter-er-1 (#66); iter-er-2 (#67, #68); iter-er-3 1/2 (#69, #70, #71) + 2/2 (#73); iter-er-4 sheet-level (#74, #77) + row/col (#82); iter-er-5 axis lifts (#79 rename_sheet; #82 four row/col axes; #81 `delete_sheet` rewriter variant; #89 `SheetDeleteWithDefinedNamesNotSupported` final lift); iter-er-6 phases 1–3 (#83, -2062 LOC) + thin-shim proper (#84, -1238 LOC); iter-er-7 task A corpus parity (#86, +9 tests, 0 bug surfaces), task B bench refresh (#85, all gates green at 1.076× of 174 ms canonical), task C Workbook.save tests + retire legacy `emitWithAppends` (#87); Codex-found bug fix on empty-string vs rich SST entry (#88). **Final state**: pkg/editor.zig 6021 → 3231 LOC (-46%); `Editor.save` is a 14-line shim over `Workbook.save` (passthrough preserves SHA256 byte-identity for no-mutation saves); all four cross-sheet rewriters wired into `Workbook.{insertRow, deleteRow, insertColumn, deleteColumn, renameSheet, deleteSheet}`; only model-invariant refusals remain (drawings/pivots/panes/autoFilter/table/last-sheet — no rewriter exists for those, staying refused). |
| **B3** Writer rebase onto Workbook | ✅ shipped 2026-05-10 — [plan](writer-rebase.md). All 7 iters landed: iter-wr-0 inventory ✅ #91 (568-LOC doc cataloguing 14 emit paths + 39 golden tests + 13-row gap table); iter-wr-1 SST ✅ #94 (`pkg/sst_plan.zig` 309 LOC, std-only); iter-wr-2 Styles ✅ #95 (`pkg/styles_plan.zig` 842 LOC, byte-fragile axis with parity test); iter-wr-3 workbook.xml ✅ #97 (`pkg/workbook_xml_plan.zig` 472 LOC + Workbook.addDefinedName); iter-wr-5 ZIP ✅ #96 (`pkg/zip.zig` 354 LOC, std-only takes `DeflateFn` callback); iter-wr-4 sheet emit ✅ #98 (`pkg/sheet_plan.zig` 1565 LOC; 7 byte-equivalence parity tests); iter-wr-6 helper dedup + `validateSheetName` reconciliation + `SheetState` extraction ✅ #100 (13 SheetWriter methods become thin forwarders); **iter-wr-7 Workbook fresh-emit + Writer.save thin shim + corpus parity ✅** — new `pkg/fresh_emit.zig` (527 LOC, std-only) hosts the entire archive orchestration. `Writer.save` collapsed to a **17-line shim** (target ≤ 80). `Workbook.saveFreshEmit(path)` ships through the same substrate. 13 `add*`/`set*` forwarder methods added to `pkg.Worksheet` (mirroring `xlsx.Writer.SheetWriter`). 5 new parity tests pin Workbook fresh-emit byte-equivalence to Writer.save. Plus 3 prep PRs ✅: #92 `PartStore.fresh()`, #93 `Workbook.empty()` + SST rich-axis. **Architectural pattern**: each Writer subsystem extracts to a std-only `pkg/<sub>_plan.zig` consumed by Writer + Workbook; the wr-7 closer adds `pkg/fresh_emit.zig` as the shared archive-build orchestrator. src/writer.zig: 6579 → 5256 LOC (-1323, -20.1% cumulative). |
| **B-fuzz** Coverage-guided fuzz nightly | ✅ shipped (reader + package layer fuzz binaries on ubuntu-22.04 nightly) |
| **C2a** Object extraction (images / charts / opaque) | ✅ shipped: image + chart anchors + series refs + Strict OOXML content-type detection + `<xdr:absoluteAnchor>` pixel-coordinate parsing (`absolute: ?AbsoluteAnchor`). Namespace handling is comprehensive: multi-prefix tracking (xdr_alts list, same-URI preference, late-declared bindings via full-document scan), proper XML scope (closed self-closing AND container siblings don't leak bindings, depth counter for nesting), in-scope local-binding authority (root fallback only when no local), per-tag chart-element verification (no false matches from unused declarations), comment/CDATA/PI awareness everywhere (forward state machine, fake-markup filtering, quote-aware tag-end + attribute-value scanning, candidate filtering at source), O(n) flat under adversarial input. |
| **C2b** addImage | ✅ shipped (PR #60 — oneCellAnchor at zero offset, fixed 1\"×1\" extent, drawingless-sheets-only; `pkg/drawing_emit.zig` carries the OOXML helpers; 6 typed errors covering anchor / mime / drawing-existing / rels-malformed cases) |
| **C1 M1** Formula tokenizer + loss-preserving printer | ✅ shipped (`src/formula/tokenizer.zig` — A1 refs incl. case-insensitive, sheet qualifiers, ranges, names, function-call disambiguation, number/string/bool/error literals, every operator, array constants, whitespace preserved, external-wb refs as `.unknown`) |
| **C1 M2 m1** A1 cell-formula rewriter (`rewriteFormula`) | ✅ shipped (PR #32 — `src/formula/rewriter.zig`; insert/delete rows/cols + `rename_sheet`; absolute-marker preservation; `target_sheet`-scoped bare refs; range collapse to `#REF!`; 20 inline tests + `checkAllAllocationFailures`) |
| **C1 M2 m1.5** Workbook wiring (`rewriteAllFormulas`) | ✅ shipped (PR #34 — walks every sheet, every formula cell; stages rewritten text via `setCell`; counts rewrites; byte-identical no-ops skipped) |
| **C1 M2 m1.6** `Workbook.renameSheet` convenience | ✅ shipped (PR #35 — validates new name (length, forbidden chars, "history" reserved, case-insensitive duplicate detection); rewrites cross-sheet qualifiers via `rewriteAllFormulas`; patches `xl/workbook.xml`'s `<sheet name=>` attr; refreshes the typed view via re-parse so `wb.sheet(idx).name()` reflects the new value immediately) |
| **C1 M2 m2** DV / CF formula rewriter | ✅ shipped (PR #48 — `rewriteAllValidationsAndConditionalFormats(edit, target_sheet)` with full splice persistence; byte-level rewrite of `<formula1>` / `<formula2>` / `<cfRule><formula>` inner-text spans, preserving every surrounding attribute (errorTitle, dxf_id, priority, etc.); 3 inline tests with synth helpers) |
| **C1 M2 m3** Defined-names + internal hyperlink rewriter | ✅ shipped (PR #59 — `rewriteAllDefinedNames(edit, target_sheet)` + `rewriteAllHyperlinkLocations(edit, target_sheet)`; collect-then-splice pattern fixes the PR #37 sheet-scope panic; 4 inline tests with synthetic-block injection helpers) |
| **Convenience: cellStyle accessor** | ✅ shipped (PR #39 — `Worksheet.cellStyle(ref)` returns composite `ResolvedStyle { font, fill, border, alignment, number_format_code }`; v1 simplifications: `apply_X` false ⇒ null sub-style; out-of-range sub-ids ⇒ null) |
| **Convenience: Worksheet.deleteCell** | ✅ shipped (PR #49 — new `CellValue.deleted` arm; emits the cell as fully removed from `<sheetData>`, distinct from `.blank`) |
| **Convenience: numberFormatFor** | ✅ shipped (PR #51 — `Workbook.numberFormatFor(style_idx)` with built-in numFmt table covering 0-4, 9-22, 37-40, 45-49 + custom-table fallback) |
| **Convenience: hasUnsavedChanges** | ✅ shipped (PR #56 — `Workbook.hasUnsavedChanges()` + `PartStore.hasUnsavedChanges()`; documented quirk: `PartStore.save` does NOT clear overrides post-save, so the predicate is "diff vs original on-disk archive") |
| **Convenience: deleteSheet** | ✅ shipped (PR #58 — `Workbook.deleteSheet(idx)` patches workbook.xml + rels + Content_Types Override, sheet-count-aware re-parse, shrinks worksheets[] with renumbered survivors; 3 inline tests; orphan-part trade-off documented since PartStore lacks `removePart`) |
| **Convenience: Python binding setCell+lifecycle tests** | ✅ shipped (PR #42 — Editor C ABI round-trip; documented gap: Workbook overlay surface NOT yet exposed in C ABI, so binding still uses `Editor`) |
| **Bench: appendRows wall-clock harness** | ✅ shipped (PR #69 — `tests/bench/bench_append_rows.zig` + `zig build bench-append-rows`; baseline 174 ms median on 100k×5; gate ceiling for iter-er-3 at 191.4 ms) |
| **Fix: writer.zig Huffman u15 → u13 freq scale** | ✅ shipped (PR #70 — `lit_enc.generate(&lit_freq, 15)` crashed inside `std.compress.flate.HuffmanEncoder.bitCounts` on `<sheetData>` payloads ≥ ~10k repetitive numeric rows; tightening `scaleFreqs` cap to u13 gives the level walker enough numerical headroom to converge; output is still valid DEFLATE; 50_000×5 round-trip regression test) |
| **D1** Formula evaluator | deferred indefinitely |
| **D2** Typed chart emit | deferred |
| **Refusal lift: frozen panes** | ✅ shipped 2026-05-11 — `pkg/sheet_edit.zig` shifts `xSplit`/`ySplit` + `topLeftCell` for `state="frozen"` / `state="frozenSplit"` panes during the row/col byte transform; editor refusal guards dropped the `<pane>` scan; `state="split"` (pixel offsets) surfaces `error.SplitPaneNotSupported`. 11 pure-function tests on the byte transform + 4 Editor round-trip tests pin the lift. |
| **Refusal lift: autoFilter** | ✅ shipped 2026-05-14 — `pkg/sheet_edit.zig::processAutoFilterTagRow`/`processAutoFilterTagCol` shifts the row/col halves of `<autoFilter ref>` during the byte transform; full-range collapse drops the element; `<filterColumn colId="N">` children rebase to `new_abs - new_tl_col` on col edits, with the filterColumn at the deleted column dropped entirely. Editor row+col guards no longer scan for `<autoFilter`. 11 pure-function tests + 4 Editor round-trip tests pin the lift. Caveat: nested `<sortState ref>` inside open-form autoFilter isn't yet rewritten — pre-existing third-party files only, since zlsx never emits open-form autoFilter. Remaining refused axes (per `docs/plans/refusal-audit.md`): drawings/pictures (anchor rewriter not yet wired), pivots/tableParts (cross-part ref graph, stays). |

Public surface added by this batch:
- New module `zlsx_pkg` (root: `src/package/root.zig`) re-exports
  `PartStore`, `Part`, `Relationship`, `imageAnchors`,
  `ChartAnchor`, `chartAnchors`. Consumable independently of the
  reader/writer surface.
- `src/unicode/casefold.zig` + `src/unicode/nfc.zig` for the
  sheet-name dedup pipeline.

### Known module-graph constraint (Zig 0.15.2)

A `zlsx extract-images` CLI subcommand was attempted as part of
this batch but reverted. The Zig 0.15 module-graph computation
treats every file under a module's package directory as part of
that module's tree, even files the root file doesn't transitively
import. With:

- `cli_mod` rooted at `src/cli.zig` (package dir = `src/`).
- `package_mod` rooted at `src/package/root.zig` (package dir =
  `src/package/`).
- `writer_mod` rooted at `src/writer.zig` (package dir = `src/`).

When `cli_mod` adds both `zlsx_pkg` and `writer` as imports, Zig
sees `src/package/store.zig` reachable from both `zlsx_pkg`'s tree
(via `root.zig`) AND `writer`'s tree (because `src/package/` lives
under `src/`, writer's package dir). It rejects with:

    error: file exists in modules 'root' and 'writer'

Workarounds for the next attempt:
1. Move `src/package/` to a sibling directory (`pkg/` or
   `subpkg/package/`) so it isn't under writer's package dir.
2. Ship `extract-images` as a separate executable that links only
   the package layer — no Editor / Writer.
3. Use direct path imports in `store.zig` (`@import("../writer.zig")`)
   so no named-module mechanism is needed; trade-off is a tighter
   coupling that pulls writer into every consumer.

Until one of those lands, callers wanting to extract images
programmatically can use `zlsx_pkg` directly from a Zig program.
The CLI surface stays as it was before this attempt.

## TL;DR

Three quick wins (A1 Unicode + char-length, A2 Windows tests, A3 bench
CI report-only) ship inside a quarter. Foundational work is a 4–6+
month effort organised as a Package store → Workbook overlay → Editor
rebase → Writer rebase chain. Formula core (tokenizer + rewriter, no
evaluator) and image extraction ride parallel after the part store
exists. Pivot creation, chart creation, and a formula evaluator are
all explicitly out of scope until production proves the prior tier.

## Sequencing graph

```
                 Tier A (parallel quick wins)
                 ├── A1 Unicode dedup + char length     [2–4 w]
                 ├── A2 Windows runtime tests           [1 w]
                 └── A3 Bench-regression CI (report-only) [1 w MVP]

                 Tier B (foundation, sequential after A)
                 ├── B0 PartStore + rel resolver + content-types   [4–6 w]   ← unblocks B1, C2-extract
                 ├── B1 Workbook typed overlay (cells/sheets)      [6–10 w]  ← depends on B0
                 ├── B2 Editor rebase onto Workbook                [4–6 w]   ← depends on B1
                 └── B3 Writer rebase onto Workbook                [4–6 w]   ← depends on B2

                 Tier B-side (parallel to B)
                 └── B-fuzz Coverage-guided fuzzing (Linux nightly) [3–4 w]

                 Tier C (foundation+, depends on B0)
                 ├── C2a Object extract / list / replace opaque bytes [3–5 w]
                 │   (depends on B0 only; does NOT need formula core)
                 ├── C2b addImage (image creation)                   [2–3 w]
                 │   (depends on C2a + B1)
                 └── C1 Formula Core: tokenizer + rewriter           [10–14 w]
                     (depends on B1)

                 C2a ships before C1 because it only needs B0; C1
                 requires the typed Worksheet model from B1.

                 Tier D (long-tail, optional)
                 ├── D1 Formula evaluator (minimal)
                 ├── D2 Typed chart emit (line/bar/scatter)
                 └── (pivot creation explicitly out)
```

Total foundational path B0 → B1 → B2 → B3: **4–6+ months**, dominated
by single-file complexity of `xlsx.zig` and `writer.zig`.

## Tier A — Quick wins

### A1. Unicode sheet-name dedup + char-length validation (2–4 weeks)

**Problem (expanded):**
- `asciiEqlFold` in `src/writer.zig` only handles 7-bit ASCII; `café` /
  `CAFÉ` and `ß` / `SS` come out as distinct.
- The 31-char limit at `src/writer.zig:2006` measures **bytes** not
  characters, so legal multi-byte Excel names get rejected.

**Approach:** non-Turkic full Unicode case fold + NFC canonicalisation;
fix char-length to count Unicode scalar values; ASCII fast path stays.

**API delta:**
```zig
pub fn excelSheetNameKey(allocator, name) ![]u8     // canonical NFC(casefold) bytes
pub fn excelSheetNameEql(a, b) !bool                 // semantic equal
```
Replace `asciiEqlFold` callers and the byte-length check.

**Implementation:**
- Vendor generated tables from `CaseFolding.txt` + minimal NFC tables
  (zero third-party-runtime contract preserved; no `utf8proc` linkage).
- Cap generated table size; one-page review policy on the generation
  script before each Unicode version bump.
- Empirical Excel matrix as the source of truth, encoded as test
  fixtures: `café/CAFÉ`, composed/decomposed `é`, `ß/SS`, Greek sigma,
  Turkish dotted I, Kelvin sign, fullwidth ASCII.

**Walk-away gate (tightened):** an empirical Excel duplicate matrix is
recorded; tests encode expected behaviour; equivalences NOT supported by
the v1 fold (e.g. Turkish locale variants, NFKC compatibility) are
**explicitly rejected** with a documented error rather than silently
passing through. No "ship NFC and document" hand-wave.

**Risks:** generated tables can bloat compile time; Excel's actual
semantics may not equal Unicode default casefold + NFC; char length
fix affects existing test fixtures.

### A2. Windows runtime tests (1 week)

**Problem:** release pipeline cross-compiles a Windows binary but never
runs the test suite, C ABI tests, or Python wheel smoke on Windows.

**Approach:** add a `windows-latest` job to `.github/workflows/ci.yml`
matrix running `zig fmt --check`, `zig build test`,
`scripts/fetch_test_corpus.sh` via Git Bash, `zig build test-corpus`,
2–3 CLI fixture smoke tests. Plus a `windows-smoke` job in
`release.yml` that exercises the released zip.

**Tightened gate (per critique):**
- Windows job MUST run the C ABI test target.
- Windows job MUST exercise at least one Python wheel
  import/load-fixture smoke test.
- Both pass on `windows-latest` for two consecutive PRs and add
  ≤ 10 min runtime.

**Likely friction (concrete):**
- Hard-coded `/tmp` paths in writer tests around `src/writer.zig:3288`
  and elsewhere — Windows lacks `/tmp`.
- Git Bash may hide problems native PowerShell users hit; add a
  PowerShell-only smoke step too.
- `src/cli.zig` already branches on Windows for signal handling;
  exercise that branch.

**Walk-away:** no-go only if the path/temp-dir refactor balloons past
two PRs of churn.

### A3. Benchmark-regression CI — report-only (1 week MVP)

**Problem:** no CI catches performance regressions across commits.

**Approach:** custom hyperfine-based pipeline. MVP is report-only —
emits warning + Markdown summary as a PR comment but does not block
merges. A blocking gate may be introduced only after the variance
and baseline gates below pass.

**Pipeline:**
1. `scripts/bench_ci.sh` builds ReleaseFast, runs `hyperfine -N --warmup 5 --runs 20 --export-json` on a small fixed fixture set (worldbank_catalog, ons_cpi_detailed). Cache corpus fetch to remove that variance.
2. PRs check out base SHA, build + run same fixtures into `bench-base.json`. Pin runner class to one image to control variance.
3. `scripts/compare_bench.py` emits a Markdown table + warning comment when current is slower by > 10% AND > 3σ on two consecutive runs. Never sets the job to red.
4. Upload artefacts; on `main`, retain for trend history.

**Walk-away gate (tightened):**
- Stay report-only until ≥ 20 main-branch samples show per-fixture
  variance < 7%.
- Blocking gate requires historical baseline, runner class pinning,
  AND a manual override knob (e.g. PR label or commit-trailer to
  acknowledge intentional regression).

**Risks (added):** GitHub Actions runner variance; PR base-checkout
doubles runtime; corpus fetch can dominate variance unless cached.

## Tier B — Foundation

### B0. Package store + rel resolver + content-types model (4–6 weeks)

**Why split:** the v1 plan folded this into B1 ("Workbook"), but the
PartStore is an independently shippable layer that:
- preserves macros / drawings / charts / images / custom parts byte-for-byte;
- gives C2 image extraction an immediate data structure to work on;
- makes B1 + B2 + B3 a typed overlay rather than parallel parsers.

**API delta:**
```zig
pub const PartStore = struct {
    pub fn open(allocator, path) !PartStore;            // mmap or file-backed
    pub fn save(self, path) !void;                       // byte-preserve untouched
    pub fn part(self, name) ?Part;                       // resolved relative path
    pub fn rel(self, owner: []const u8) ![]Relationship; // owner = part name
    pub fn replacePart(self, name, bytes) !void;         // typed overlay path
    pub fn addPart(self, name, content_type, bytes) !void;
};
```
Public so C2 image extraction can hang directly off it.

**Milestones:**
1. Read-side: enumerate parts, parse `[Content_Types].xml`, parse
   every `_rels/*.rels`, expose resolver.
2. Write-side: emit unchanged parts byte-for-byte (same compressed
   payload, name, length); emit dirty parts via fresh deflate.
3. Typed-overlay handle for known parts (workbook.xml, sheet, sst,
   styles, theme).

**Walk-away gate (revised per critique):**
- Untouched ZIP entries' compressed payload bytes AND names AND order
  preserved.
- Relationship graph equivalent to source.
- Excel opens the round-trip output without prompting "repair?".

(NOT byte-identical whole-file round-trip — central directory metadata
re-emission would otherwise eat weeks for no functional gain.)

### B1. Workbook typed overlay (6–10 weeks, depends on B0)

**API delta:**
```zig
pub const Workbook = struct {
    pub fn open(allocator, path) !Workbook;          // = PartStore.open + lazy overlay
    pub fn openLazy(allocator, path) !Workbook;
    pub fn create(allocator) Workbook;
    pub fn save(self, path) !void;
    pub fn sheet(self, idx) !*Worksheet;
    // …
};
pub const Worksheet = struct {
    pub fn rows(self) !Rows;
    pub fn cell(self, ref) ?CellValue;
    pub fn setCell(self, ref, CellValue) !void;
    // …
};

// Compat facades for one minor line:
Book.toWorkbook(options)        // existing reader → Workbook
Editor.openWorkbook(path)       // = Workbook.open
```

**Milestones:**
1. Lazy `WorksheetModel` — sparse cells, styles, merges, hyperlinks,
   validations, comments, conditional formats. Backed by existing
   `ensureSheetLoaded` + row parser.
2. Workbook-level state: SST view, styles view, defined names view.
3. Compat layer: `Book.toWorkbook` adapts existing reader output.

**Walk-away gate:** a 100k-row × 10-sheet workbook needs ≤ 2× current
`Book.openLazy` RSS before any sheet is touched.

**Risks (added per critique):** the existing `Book` is a large
all-in-one reader state with many maps (`src/xlsx.zig:529`). Mutating
it into a shared model risks destabilising the reader. Mitigate by
adapter-only path first; rebase reader internals last.

### B2. Editor rebase onto Workbook (4–6 weeks, depends on B1)

**Per critique:** **editor rebase precedes writer rebase**. The editor
already needs preservation + structural safety; the writer can stay a
fast fresh-file emitter longer.

The current editor lives at `src/xlsx.zig:4240` with many conservative
pending-operation interactions. Rebase onto Workbook mutations,
retaining ZIP substitution for untouched parts.

**Walk-away gate:** Editor round-trip parity (current corpus tests
green on the new path); ≤ 1.5× current ~5 ms ZIP-substitution
latency on small workbooks.

### B3. Writer rebase onto Workbook (4–6 weeks, depends on B2)

`Writer.save` emits from a `Workbook` populated by the fluent builder.
Unifies SST / styles between fresh-file and load-modify-save paths.
Last because it's the mature path; touching it without B0–B2
production bake-time risks regressions in the fastest production code.

### B-fuzz. Coverage-guided fuzzing — Linux nightly (3–4 weeks, parallel to B)

**Problem:** current fuzz harnesses are PRNG-driven; CI runs
fuzz-smoke only.

**Approach:** Zig 0.15.2 built-in fuzzing on Linux x64 first.

**Status of Zig fuzz:**
- `zig test --help` exposes `-ffuzz`; `std.testing.fuzz(context, testOne, .{ .corpus = ... })` exists.
- Zig source: not implemented on Windows (shared-memory + COFF/PE debug).
- macOS Mach-O `addEntryPoint` crash documented in this repo.
- Instrumentation = inline counters + trace-cmp; not AFL `trace_pc_guard`.

**Milestones:**
1. Convert 3–5 reader targets to `input → target` form:
   `parseSharedStrings`, `parseWorkbookSheets`, `Rows.next`, `Book.open`.
2. Seed corpora at `tests/fuzz/corpus/{shared_strings,workbook_sheets,rows,book_open}`.
3. Nightly `ubuntu-22.04` workflow: 15–30 min per target; uploads
   crash artefacts; deterministic replay first.
4. **macOS fuzz job allow-failure** (per critique): runs the same
   targets but expects to fail with a clear "blocked on Zig upstream
   issue/link" message. Surfaces upstream-status.

**Risks (added):** Zig fuzz API may shift under us; ZIP-parsing fuzz
hits archive code more than worksheet semantics — write targets that
bypass archive parsing where worksheet correctness is the goal.

**Walk-away (go/no-go):** go if `zig build test --fuzz` runs 30 min on
ubuntu-22.04 and finds + replays at least one entry. No-go for
macOS/Windows fuzz until Zig upstream verified.

## Tier C — Foundation+

### C1. Formula Core: tokenizer + rewriter (10–14 weeks, depends on B1)

**Problem:** `anySheetCrossSheetCarrier` at `src/xlsx.zig:5865` and
related guards refuse formulas, hyperlinks, validations, conditional
formats, defined names, drawings, panes, tables, autofilters. Per
critique: tokenising formulas alone does NOT liberate structural edits
— each guard category has its own ref-rewrite story.

**Approach (tightened):** tokenizer + rewriter handles formulas FIRST,
then in successive iters expands to data-validation formulas,
conditional-format formulas, defined names, internal hyperlink
locations. Each iter shrinks the refusal set. Evaluator stays deferred
to Tier D.

**API delta:**
```zig
Formula.parse(allocator, text) !FormulaAst
Formula.rewriteRefs(ast_or_text, RewriteContext) ![]u8
Workbook.rewriteReferences(edit: StructuralEdit) !void
Writer.setRecalcOnOpen(bool)
```

**Milestones:**
1. **Tokenizer + loss-preserving printer** — A1 refs, sheet-qualified
   refs, ranges, named ranges, string literals, functions, operators,
   array constants. Public parse/format utility.
2. **A1 cell-formula rewriter** — row/col/sheet rename/delete on
   `<f>` bodies. Liberates the formula-only branch of structural-edit
   refusal.
3. **DV / CF formula rewriter** — same axis, new owners. Liberates
   `<dataValidation>` and `<conditionalFormatting>` formulas.
4. **Defined-names rewriter** — workbook-scope refs.
5. **Internal hyperlink `location`** — same axis.
6. **Shared / array formula support** — base + dependent ranges.

**Walk-away gate (replaced "95% corpus coverage"):** grammar-class
gates. The rewriter must handle, with golden round-trip tests:
- A1 refs, absolute / mixed / relative
- Quoted-sheet refs (`'My Sheet'!A1`)
- Range refs
- Defined names (workbook + sheet scope)
- Shared formula bases + dependents
- Array formula bases + spread ranges
- DV / CF / hyperlink-location refs

If any class fails, that class's owners stay on the refusal list; do
NOT silently downgrade.

**Risks:** Excel grammar is large (structured refs, 3D refs, external
workbooks, dynamic arrays, locale separators). Rewriter bugs silently
corrupt files — golden round-trips catch this only for known classes.

### C2a. Object extract / list / replace opaque bytes (3–5 weeks, depends on B0 only)

**Per critique:** image extraction + opaque-replace lives off the
relationship graph (B0), NOT formulas. C2a ships ahead of C1 because
it has no formula-rewriter dependency.

**Scope:** v1 is **read-side + opaque-replace only**. Discover, list,
and surface drawings/charts/pivots/images as typed objects with raw
payloads attached; let callers replace those payloads byte-for-byte.
No image *creation*, no chart creation. (See C2b for image creation.)

**API delta:**
```zig
Workbook.drawings(sheet) []DrawingObject
DrawingObject = union(enum) {
    image: { path, content_type, anchor, bytes },
    chart: { chart_path, chart_type, series: []SeriesRef, anchor, raw_xml },
    pivot: { table_path, cache_id, source_ref, raw_xml },
}
Workbook.replaceOpaquePart(path, bytes) !void
```

**OOXML structures handled:**
- Sheet `<drawing r:id>`, `xl/worksheets/_rels/sheetN.xml.rels`
- `xl/drawings/drawingN.xml` + `_rels`
- `xl/charts/chartN.xml` (series via `c:barChart` / `c:lineChart` /
  `c:scatterChart` / `c:ser` / `c:f`)
- `xl/media/imageN.{png,jpeg,gif}`
- `xl/pivotTables/*` + `xl/pivotCache/*` (detect + preserve only)
- workbook `<pivotCaches>`, `[Content_Types].xml`

**Milestones:**
1. Drawing parser → anchors + opaque chart/pivot/image objects.
2. Lazy raw-bytes extraction API.
3. Opaque-replace API (`replaceOpaquePart`) — same compressed-payload
   contract as B0's untouched-part round-trip.

**Walk-away gate:** stop before image creation, chart creation, chart
**rendering**, chart-style editing, slicers, SmartArt, OLE, VBA,
pivot-cache recalc, or any "edit row/col on a drawing-bearing sheet"
work. Per critique: chart **emit** is D2; image **emit** is C2b.

**Risks:** anchors are non-trivial; row/col edits on drawing-bearing
sheets stay refused (the existing guard stands until C2b lifts it).

### C2b. Image creation — `addImage` (2–3 weeks, depends on C2a + B1)

**Scope:** add a brand-new image to an existing workbook with a typed
anchor. Depends on B1 because the anchor sits inside the typed
Worksheet model; depends on C2a because it reuses the drawing-part
emission shape.

**API delta:**
```zig
Workbook.addImage(sheet, anchor, bytes, mime) !void
```

**Milestones:**
1. Drawing/media/relationship/content-types stack emission for a
   single new image.
2. Anchor math: cell-anchor (one A1 ref) + range-anchor (A1:B5).
3. Refusal lifted for row/col edits on sheets where ALL drawings are
   `addImage`-emitted (we own the anchors).

**Walk-away gate:** if anchor math turns out to require pixel/EMU
coordinate computations beyond cell-anchor + range-anchor, defer to
D-tier and keep `addImage` cell-anchor-only.

## Tier D — Long-tail, optional

### D1. Formula evaluator — minimal (deferred indefinitely)

After C1 has at least one quarter of production bake. Covers literals,
arithmetic / comparison, cell+range refs, `SUM`, `MIN`, `MAX`,
`AVERAGE`, `IF`, boolean ops, errors. Writes updated cached `<v>` for
supported formulas; otherwise marks recalc-on-open.

**Per critique:** an evaluator will consume the project if allowed.
Stay optional indefinitely; only build if a concrete user demand
appears.

### D2. Typed chart emit (deferred)

After C2 ships and image creation is exercised in production. Typed
`addChart` for line/bar/scatter; styles minimal. Chart **style editing**,
slicers, pivot creation are explicitly out of scope forever.

## Cross-cutting principles (unchanged from v1)

1. **Zero third-party runtime deps stays inviolable** — Unicode tables
   vendored; no `utf8proc` linkage.
2. **Conservative refusal posture preserved** — Editor never silently
   corrupts a workbook.
3. **One minor line of compat facades** — Book / Writer / Editor stay
   functional after Workbook ships, with deprecation notes.
4. **Walk-away criteria are real** — each tier's gate is measurable;
   per critique gates are tighter than v1.
5. **Tests precede CI gates** — bench-regression CI stays report-only
   until variance proves controlled; fuzz CI stays nightly.

## Estimation summary (revised per critique)

| Tier | v1 estimate | v2 estimate (per Codex critique) |
|---|---|---|
| A1 Unicode dedup | 1–2 w | 2–4 w |
| A2 Windows tests | 1 w | 1 w |
| A3 Bench CI MVP | 2–3 w | 1 w (report-only) |
| B0 PartStore | (folded) | 4–6 w |
| B1 Workbook overlay | 2–3 mo (whole tier) | 6–10 w |
| B2 Editor rebase | (folded) | 4–6 w |
| B3 Writer rebase | (folded) | 4–6 w |
| B-fuzz nightly | 3–4 w | 3–4 w |
| C1 Formula rewriter | 2–3 mo | 10–14 w |
| C2a Object extract / opaque-replace | 1–2 mo | 3–5 w |
| C2b addImage | (folded) | 2–3 w |

Tier B end-to-end: **4–6+ months** of focused work.
