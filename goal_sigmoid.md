# 🎯 goal_sigmoid — close every missing feature of zlsx

> Living plan for analysing, implementing, reviewing, testing, and improving
> **each missing feature** of zlsx, under the standing constraints: stdlib-only,
> Zig 0.16.0, TigerStyle-defensive, byte-splice fidelity, refuse-rather-than-half-do,
> measured performance. Companion to `docs/ROADMAP.md` (which stays the plan of
> record — sigmoid rows graduate into its table as they start, the way the D1
> ladder did from `goal_formula.md`).
>
> **Owner-interaction contract (locked, 2026-08-26):** every row carries a named
> **OWNER GATE** at the end of its analyse phase. Work inside a row does not pass
> its gate without the owner answering the gate's questions; mechanical steps
> inside a row (review rounds, test gates, benches) never pause for approval.
> Questions are asked when they are *necessary* — at semantic decision points —
> exactly as directed.

_Created: 2026-08-26 · `main` at `d99a386` · v0.8.0, Zig 0.16.0 · Status: **S0 done (gate answered 2026-08-26) · S1 done (gate open: the three numbers) · S2 done (no gate fired) · S6 partial (Zig + CLI + audit merged 2026-08-28; gate answered — shape frozen, C + Py legs = S6's second PR, hole to S7b) · S7a done (oracle question parked, §5) · S3a or S7b next**_

---

## 1. Scope decisions (locked by owner, 2026-08-26)

| Decision | Owner's choice |
|---|---|
| What "missing feature" means | **Full competitor parity** — correctness gaps + public-surface parity + un-deferred authoring, until the README matrices show no in-scope `—` against openpyxl / xlsxwriter / calamine |
| Pivot depth | **Full programme including authoring** — detection audit → typed read → staged edit lifts → fresh-workbook authoring |
| Surfaces per row | **Strict 4-way parity** — a row is `done` only when the capability lands in Zig, the C ABI, Python, *and* the CLI |
| Deferred reversals | **Writer/Workbook image-authoring parity** and **D2 chart authoring (scoped)** enter the ladder. **Formula-literal masking stays deferred** (`docs/plans/formula-literal-masking.md` remains the record) |

Standing facts that shaped the frame (verified 2026-08-26):

- The old perf "gap" is closed and gets **no row**: named evaluate ≈ **347.9 ms**
  against the 500 ms ceiling; the blocking RSS gate is the ratcheted absolute
  budget **50 189 598 B** (M10x figure 49 692 672 B = 47.39 MiB), met. The
  15.15 MiB / 3× figure stays a non-blocking research hypothesis (§9.1b–g of
  `goal_formula.md`). Every sigmoid row keeps the ordinary regression gates.
- Refusal state on `main`: every row/col-edit axis with a rewriter is lifted;
  **two guards remained** at creation — pivots (#139) and `<xm:f>` sparkline
  formulas (#140); S2 lifted the second on 2026-08-26, pivots remain.
- `src/writer.zig` excludes drawings/charts/pivots wholesale; README writer-matrix
  footnote ⁵ described an emission surface the fresh Writer does not have (fixed
  in S0). Doc truth-reconciliation is row zero, not a side effect.
- S0's inventory (`docs/plans/surface-matrix.md`) found four more Zig-only
  surfaces no row covered, a fresh `Writer` that is a *private* build module,
  no aggregate ZIP budget on any path, and a pivot guard that sees hosting
  sheets only — hence rows S3d, S3e, S5b, S11 and the S1 / S3b / S5 text below.

## 2. Out of scope (unchanged)

- `.xls`, `.xlsb`, `.ods` — permanently out; no ladder rows, ever.
- **Formula-literal masking** — stays deferred by this plan's own owner gate.
- **E4W** (Excel-on-Windows survival) and **PyPI publishing** — parked owner-action
  items, not engineering rows. They keep their existing plans.
- Embedding v1.1 follow-ups (hash-to-slot remap, quantization recall) — parked.
- Reopening the RSS research hypothesis — only by explicit owner directive.

---

## 3. The ladder

Status vocabulary is ROADMAP's: `done` / `partial` / `planned` / `todo` /
`blocked` / `deferred`. Estimates are originals and stay uncorrected.

| ID | Piece | Status | Estimate | Needs | What it does |
|---|---|---|---|---|---|
| S0 | Doc truth + surface inventory | done | 3–5 d | — | Reconciled every stale claim (Writer image footnote ⁵, Python "structural edits are a follow-up" / "0.9.0+", `to_bytes()` queue note, the pivot-guard reach, the archive-defense wording, the fuzz platform and target list), and committed the four-surface capability matrix (`docs/plans/surface-matrix.md`) with `scripts/check_surface_matrix.py` in CI. Three Codex rounds; owner gate answered 2026-08-26 (§5). |
| S1 | Core reader decompression defenses | done | 1 wk | — | The three limits (per-part cap, ratio cap, **new** whole-archive aggregate) live once, in `pkg/control.zig::decompress_limits`, and every opener admits every central-directory entry against them before anything is inflated: the core reader's walk (`Book.openLazyWithSst`, with a per-part re-check in `extractEntryToBuffer`), the editor's own structural scan (so `Editor.open` refuses before `Book.open` runs — the S0 hole), and `PartStore.scanCentralDirectory` (which had per-part + ratio at decompress time only, no aggregate). One error, `ZipBombSuspected`, on every surface; CLI exit 4 on the read, edit and embed families (`openFailureExit`). Hostile-archive tests on all three openers via `zlsx.zip_probe`. Owner gate (§5): confirm or move the three numbers. |
| S2 | `<xm:f>` sparkline refusal lift | done | 1–2 wk | — | `Workbook.rewriteAllExtensionFormulas` routes every `<xm:f>` carrier (sparkline data ranges and date axes, `x14:` CF / DV formulas; walked by `sheet_edit.nextXmFormula`, matched by leaf name like `xm:sqref`) through the formula rewriter with `on_sheet` = host sheet and `target_sheet` = edited sheet — the context `sheet_edit.zig` lacked — under every row / col / sheet-rename / sheet-delete / table-column-rename edit. All-or-nothing: `preflightExtensionFormulas` scans every sheet before an edit's first mutation and refuses (`MalformedExtensionXml` → `Row`/`ColEditUnsafeForSheet`) on a carrier it cannot read. #140's `ExtensionEditUnsafe` is gone. No owner gate fired: the corpus carries no `<xm:f>`, and a deleted source range collapses to `#REF!` by the same convention as cell formulas. |
| S3a | Parity: structural edits → C + Python | planned | 2–3 wk | S0 | Row/col insert–delete and sheet-level edits reach the C ABI (three-file transaction) and py-zlsx `Editor`; refusals surface as typed statuses in both. |
| S3b | Parity: typed reads → all surfaces | planned | 1–2 wk | S0 | Widened at the S0 gate. CLI: merged ranges, defined names, conditional formats, drawing anchors, document properties as typed reads (NDJSON, `docs/cli.md` contract). C + Python: defined names, conditional formats, anchors, panes / dimension / calc properties, sheet visibility, formula text and error tags — the C reader has no error tag and no formula text today. |
| S3c | Parity: embedding mutation → C + Python | planned | 1 wk | S0 | The Zig/CLI embedding write/prune/strip surface reaches the C ABI and Python; read-only handles retire. Also: `recovery_in_cells` on C + Python + CLI, a CLI vector / state dump, and a control-byte check on embedding metadata (S0 found `setEmbeddingsOpts` escapes `& < > "` only). |
| S3d | Parity: existing-workbook authoring → C + Python + CLI | planned | 2–3 wk | S0 | Added at the S0 gate. `Workbook.addStyle` / `addDxf` / `internNumFmt`, the `Worksheet.set*` / `add*` layout, merge, hyperlink, comment, DV and CF methods, `addDefinedName`, `deleteCell`, and a standalone mark-recalc-on-load — Zig-only today — reach the C ABI, py-zlsx `Editor` and the CLI edit family. |
| S3e | Parity: opening strategies → C + Python | planned | 1 wk | S0 | Added at the S0 gate. Lazy per-sheet loading (`Book.openLazy` / `preloadSheet` / `streamSheet`) and the lazy SST backend reach the C ABI and Python (per-sheet also the CLI); `--sst-lazy` already exists. |
| S4 | Indexed palette + tint resolution | planned | 2–3 wk | S0 | Legacy `indexed="N"` (default table + workbook `<indexedColors>` overrides + reserved 64/65), `tint` math, every carrier (font, fill, border, rich-text runs); oracle-pinned; all four surfaces. |
| S5 | Image authoring parity (reversal) | planned | 2–3 wk | S0, S11 (CLI leg) | The fresh Writer routes through the **one canonical** Workbook/C2b drawing emitter — native extents, `twoCellAnchor`, dialect fidelity, typed refusals; then C ABI + Python + CLI image authoring. No second emitter, ever. Also exports `src/writer.zig` as a public module (`zlsx_writer`) — S0 found it is a private `createModule` no dependent project can name. |
| S5b | Typed object extract / replace / remove → all surfaces | planned | 2–3 wk | S0 | Added at the S0 gate. C2a's product promise as a typed operation on every surface: extract, replace and remove images and charts with the drawing, its relationships and content types repaired (today Zig has raw `PartStore` composition only; `ChartAnchor` is read-only); retires the `zlsx-extract-images` sibling into `zlsx`. |
| S6 | Pivot topology + typed read | partial | 3–4 wk | S1 | **Shipped (Zig + CLI):** `pkg/typed_parts/pivot_xml.zig` — cache-definition + table-definition parsers, Strict / Transitional by namespace prefix, the two `ref` attributes exported as byte spans for S7a / S7b; `pkg/pivots.zig` — the graph walked from both roots (`<pivotCaches>` → cacheDefinition → records; sheet rels → pivotTable), every `worksheetSource` resolved through the engine's symbol table (`sheet` attribute, a table's host, a defined name's body) or to another workbook; `Workbook.pivotTables`; `zlsx pivots` (contract: `docs/cli.md` "pivots"). Conservative and read-only — formats, hierarchies, OLAP stay raw; parts byte-preserved. **Audit answered:** #139's guard does *not* see source-only sheets — pinned by three `S6 audit` tests in `pkg/editor.zig` (synthetic `sheet`+`ref` source: admitted, `ref` left stale; synthetic table-named source: admitted, the table rewriter keeps it valid; corpus `openxlsx_loadExample.xlsx`: both hosts refuse, source-only `mtcars` admitted). The corpus carries exactly one pivot workbook, both caches table-named. Six Codex rounds (6 → 8 → 8 → 7 → 1 → 2 findings, plateau reached; one finding declined and documented: a prefix-bound `<tableParts>` is invisible to the table index exactly as it is to the table editor). **Remaining:** the C ABI + Python legs, after the gate freezes the shape. |
| S7a | Pivot edit lift: output-location | done | 1–2 wk | S6, S2 | Row/col edits on a sheet that only *hosts* a pivot move `pivotTableDefinition/location@ref` — `pkg/pivots.zig::edit.applyToTableDefinition` splices the new rectangle at the parser's `Location.ref_span` (no second scanner), `Workbook.preflightPivotEditsForSheet` dry-runs it before the first mutation, and the sweep applies it first, all-or-nothing across the sheet's pivots; an edit below / right of the rectangle leaves the part byte-identical. Everything else still refuses (`PivotEditUnsafe` → `Row`/`ColEditUnsafeForSheet`, CLI exit 3): an edit inside the pivot's footprint — the rectangle plus a conservative report-filter band above it (`rowPageCount + 1` rows, `3 · colPageCount` columns; Excel refuses that edit too) — a host some cache *may read* (`Pivots.mayReadFromSheet`: resolved to it — the corpus' `IrisSample`, S7b — or unresolved (a dynamic defined name, a dangling spelling) or of unknown type), a pivot part two sheets host, a shift past the grid, an unreadable graph. The host's `<pivotSelection>` coordinates move with the grid; the install is `PartStore.replaceParts`, transactional under allocation failure. The sheet's relationships gate the read, so the S6 source-only admission is unchanged. Pinned by the `edit:` tests in `pkg/pivots.zig` (moves, refusals, the filter band, decoy + entity splice, grid edges, allocation failure) and the `S7a` tests in `pkg/editor.zig` (host-only lift, footprint refusal, host-and-source refusal, no-graph-read on a pivot-less sheet, all-or-nothing on the direct `Workbook` path, unresolved-source and shared-part refusals, transactional install under allocation failure, corpus `mtCars Pivot` → `A2:D6`) and the `pivotSelection` tests in `pkg/sheet_edit.zig`. Codex round 1: 5 findings (1 HIGH — an `OFFSET(...)` named source reading the host was admitted), all fixed. Oracle question parked (§5). |
| S7b | Pivot edit lift: source rows | planned | 2–3 wk | S7a | Source-sheet row edits rewrite `worksheetSource@ref` under an explicit, owner-chosen cache policy; refuses without it. |
| S7c | Pivot edit lift: source columns | planned | 2–3 wk | S7b | Column edits touch cache-field schema, ordinals, items, filters, calculated fields — lift only what the oracle proves; the remainder stays a *documented, tested* refusal. A refusal outcome closes this row. |
| S8 | Pivot authoring (reversal) | planned | 6–10 wk | S6, S7a, S11 (CLI leg) | Fresh workbooks only: one pivot type, one contiguous source range, no external/consolidation/OLAP sources, explicit refresh policy, Excel + LibreOffice oracle. Editing existing pivots stays refused. |
| S9 | D2 chart authoring (reversal) | planned | 6–10 wk | S5, S11 (CLI leg) | Fresh workbooks only: one chart family first, explicit series/category formulas, unique axis IDs, cached-values policy, Office round-trip oracle. Never typed-reemit an existing unknown chart. |
| S11 | CLI fresh-workbook authoring (`zlsx write`) | planned | 2–3 wk | S0 | Added at the S0 gate instead of an `n/a` ruling. A JSON workbook spec → `zlsx write out.xlsx`, covering the §2 `Writer` family (sheets, typed rows, styles, rich text, formulas, layout, merges, hyperlinks, comments, DV / CF, defined names, date → serial, hidden sheets). The CLI legs of S5, S8 and S9 land on it. |
| S10 | Parity closure audit | planned | 1 wk | all above | Re-run the README matrices + surface matrix; every remaining `—` is either closed or carries a recorded owner ruling; sweep the parked two-liners (engine-fingerprint accessor on Zig / CLI, hidden sheets on write for Zig / C / Py). Final whole-ladder Codex review round. |

Dependency spine: S0 → {S3a,S3b,S3c,S3d,S3e,S4,S5,S5b,S11} · S1 → S6 → S7a → S7b → S7c · {S6,S7a} → S8 · S5 → S9 · S11 → the CLI legs of {S5,S8,S9} · everything → S10. S1/S2 are independent early wins.

---

## 4. Phase discipline (every row, in order)

1. **Analyse / decide** — read the real parts (corpus + spec), enumerate the
   graph, write the row's contract. Ends at the row's **OWNER GATE**.
2. **Failing test or oracle first** — the behavior is pinned by a test that
   fails before the implementation exists. For authoring rows, the oracle is
   an Office/LibreOffice round-trip, regenerated not hand-written.
3. **Implement** — refuse-over-corrupt; byte-splice fidelity; typed errors.
4. **Review** — Codex loop (`gpt-5.6-sol`, high), schema-strict, until
   ship-ready or the owner ends the loop. In-house agents for round-1 breadth.
5. **Gates** — full suite + fuzz; byte-diff on untouched parts; corpus
   round-trip; time gates in **ReleaseFast**, the RSS gate in **ReleaseSafe**;
   allocation-failure paths where the row touches allocation.
6. **Mutation proof** — *targeted inverse patches*: every critical new test is
   shown to fail when its owned behavior is reverted. No generic campaign.
7. **Improve** — only against a fresh measurement, and only if the row moved a
   measured lane. No speculative tuning.

Phases are **gates, not PRs**. A row ships as however many PRs its semantics
need (S7\* will be several; S0 is one). Analysis and review never get their own
roadmap rows.

## 5. Owner gates — the questions, named in advance

| Row | Gate question(s) the owner will be asked at analyse-end |
|---|---|
| S0 | **Answered 2026-08-26.** (1) Surface-matrix format + CI lint — approved as is. (2) Six `n/a` rulings — all accepted (each amends strict four-way parity for that cell only; recorded in `docs/plans/surface-matrix.md` §Rulings). (3) Rows S3d, S3e, S5b, S11 adopted; S3b widened to all four surfaces; S11 in the spine ahead of S10. (4) S1's corrected text confirmed. (5) The `zlsx_writer` module export rides S5. |
| S1 | Confirm the three limits (per-part cap, ratio, aggregate budget) and their values. **Shipped with defaults pending the answer:** 512 MiB per part and 4096:1 (the package path's numbers, unchanged) plus a new 2 GiB whole-archive aggregate — chosen as ~5× the corpus' largest legitimate total (379 MiB, `wdi_excel.xlsx`; largest part 274 MiB, highest real ratio 40:1). All three are one struct literal in `pkg/control.zig::decompress_limits`; moving any of them is a one-line change with no other code to touch. Two things only the owner can decide: (1) keep, raise or lower each number; (2) whether `eval` / `recalc` should report a breach as 3 (typed refusal — "limits are refusals at every layer") rather than the 2 (open/parse) they emit today. |
| S2 | **None fired (2026-08-26).** The corpus carries no `<xm:f>`, so the row shipped on the rewriter's existing conventions: a deleted source range collapses to `#REF!` (as for a cell formula), an unreadable carrier refuses the whole edit. Both are one-line changes if the owner wants a different posture. |
| S3a–e | Confirm C ABI naming/versioning for the new exports (v1 pattern vs v2 suffix, per `c-abi-status-v1.md`). S3d: which authoring methods reach the CLI edit family as sub-commands vs a spec file. |
| S4 | The color contract: raw provenance + effective ARGB (recommended) vs effective-only; tint rounding + alpha rules. |
| S5 | Confirm Writer API shape (mirror `addImage`/`addImageRange`/`addImageAnchored` or a reduced set); the public module name (`zlsx_writer`). |
| S5b | Approve the typed object API (extract / replace / remove; what "repaired" means for a removed image's drawing) and the `zlsx images` sub-command shape. |
| S11 | Approve the workbook-spec JSON schema before it freezes (it is a public contract like the NDJSON envelope). |
| S6 | **Answered 2026-08-28** (asked 2026-08-27, PR #199). (1) The `zlsx pivots` record shape (`docs/cli.md` "pivots") **approved as built** and frozen at merge: axes as `{"field":NAME,"idx":N}` objects with `{"values":true}` for the values axis; `cache` nested inside the pivot record; orphan caches as separate `pivot_cache` records; `source.resolved` as a sheet / external / `null` union; `types` as a list of present kinds; no `v` version key (the read family carries none). (2) The C ABI + Python legs ship as **S6's second PR**, mirroring the frozen shape; the row stays `partial` until they land. (3) The source-only-sheet hole **stays open until S7b**, which owns the refusal and the `worksheetSource@ref` rewrite together; the three `S6 audit` tests keep it pinned meanwhile. |
| S7a | **Parked 2026-08-28 (PR built without it).** The oracle: does Excel reopen a workbook whose pivot rectangle zlsx moved (`location@ref` shifted, every other pivot byte untouched, `pivotCacheRecords` as saved) and draw the pivot at the new place without a repair prompt — and does the report-filter band follow it? The lift assumes yes for a pure shift (the same bytes Excel writes after its own row insert above a pivot) and refuses inside the footprint. Two things only the owner's Excel can settle: (1) the band width for over-then-down filters (the guard refuses `3 · colPageCount` columns; Excel draws `3 · colPageCount − 1`) — narrow it if the oracle shows the blank column is not the pivot's; (2) whether a workbook saved by zlsx after a shift needs `refreshOnLoad` — today it does not set it. Regenerate through `scripts/oracle/` once Excel is free. |
| S7b | **The cache policy**: on source-row edits, adjust `worksheetSource@ref` and mark refresh-on-load, vs also rewriting cache records. This is the row's semantic core. |
| S7c | Per-case lift/refuse ruling once the oracle matrix is in hand. |
| S8 | Which pivot type ships first; the refresh policy; the oracle bar (Excel alone vs Excel+LibreOffice both green). |
| S9 | Which chart family ships first; cached-values policy; the oracle bar. |
| S10 | Final ruling on every surviving `—`. |

## 6. Where the truth will live

- This file: the ladder + gates, updated per row like `goal_formula.md` was.
- `docs/plans/surface-matrix.md` (born in S0): the 4-surface capability truth.
- `docs/ROADMAP.md`: rows graduate into the plan-of-record table as they start;
  refusal-state paragraph updates as #139/#140 fall.
- Per-row detail specs, where a row needs one: `docs/plans/sigmoid-<row>.md`.
