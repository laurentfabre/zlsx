# 🎯 Goals — zlsx

> Living goal file. The **north star** plus the **active track** and near-term objectives.
> Detail lives in `docs/plans/`; narrated context + collaboration rules live in the
> knowledge base at `docs/kb/` (open `docs/kb/kb.html`, gitignored).

_Last updated: 2026-08-01_

---

## North star

> **The fastest, safest `.xlsx` library — read *and* mutate-in-place — usable from Zig, C,
> and Python, stdlib-only.** Every row/col edit either rewrites all coordinate-bearing
> parts correctly or refuses with a typed error. Never silently corrupt a workbook.

Performance posture (the bar to hold): 1.7–12× over calamine, 38× over openpyxl,
4× less RAM at MB scale. Write-path strict gate ≤ 7.4 ms.

---

## ✅ Completed track — embeddings in xlsx (all on `main` as of #134)

Store semantic vector embeddings inside `.xlsx` via vendor-namespaced OPC parts under
`xl/zlsxEmbeddings/`, in the `fabre.me` namespace. E1/E2/E3/E0 shipped in **#123**;
the carrier matrix in **#124**; ER in **#127**; E5 in **#130**; the emb-6 CLI in
**#132/#133/#134**. (PR #115 is dead — it was cut from pre-migration `main` and was
re-applied by hand.)

```mermaid
%%{init: {'theme': 'base', 'themeVariables': {'primaryColor': '#1a1a2e', 'primaryTextColor': '#e0e0e0', 'primaryBorderColor': '#00d4ff', 'lineColor': '#00d4ff', 'secondaryColor': '#16213e', 'tertiaryColor': '#0f3460', 'fontFamily': 'monospace'}}}%%
graph LR
    D["design doc ✅"] --> E1["emb-1a/1b ✅<br/>wire-format + manifest"]
    E1 --> E2["emb-2/3a/3b ✅<br/>read + write + rel"]
    E2 --> E4["emb-4 ✅ 3-tool<br/>matrix done<br/>(Win pending host)"]
    E4 --> E4B["emb-4B ◐ carrier matrix<br/>3 carriers survive<br/>both rebuilders"]
    E4B --> DEC{{"durability contract<br/>DECIDED 2026-07-26"}}
    DEC --> ER["emb-R ✅ recovery record<br/>hidden definedName + docProps"]
    ER --> E5["emb-5 ✅ Python<br/>NumPy + valid_mask"]
    E5 --> E6["✅ emb-6 — CLI (#132/#133/#134)<br/>embed --strip / --prune / --extract / --vectors"]
    style E4 fill:#16213e
    style E4B fill:#16213e
    style DEC fill:#0f3460
    style ER fill:#16213e
    style E5 fill:#16213e
    style E6 fill:#0f3460
```

**emb-4 — DONE for the 3 reachable desktop tools (validated 2026-05-30/31).**

- ✅ Harness + runner (`tests/emb-4/run-matrix.sh`); baseline + zlsx control PASS.
- ✅ **Excel for Mac 16.109.2 → PASS** — preserves all parts + the workbook→index rel.
- ⚠️ **Numbers 14.5 → STRIPPED** · **LibreOffice 26.2.3.2 → STRIPPED** · openpyxl → STRIPPED.
- ⛔ **Excel for Windows → not run** — no Windows+Excel env reachable from this Mac
  (checked 2026-05-31: no local VM; SSH config only has Linux company infra). Run on a
  Windows host or a Windows CI runner, then `emb4-verify` the saved file. Staged copy:
  `/tmp/zlsx-emb4/excel-win.xlsx` (note: `/tmp` is volatile — regenerate via the runner).

> **Key finding (the emb-4 payoff):** only Excel preserves embeddings on save; Numbers +
> LibreOffice strip them (parts removed, unrecoverable). The v1 "PASS on all targets" bar
> is **NOT met**.

**emb-4B — carrier matrix built, automated legs measured (2026-07-26).**

Six carriers, one fixture, one round-trip per tool. Against the two
archive-rebuilding consumers reachable without a GUI:

- ✅ **Survive both** — `docProps/custom.xml`, cell data, `<definedName>`.
- ❌ **Stripped by both** — `xl/zlsxEmbeddings/*` (the emb-4 control) and `<extLst>`.
- ⚠️ **Split** — `customXml/` survives LibreOffice, not openpyxl.
- ✅ **Excel for Mac → 0/6 lost** (AppleScript leg, automated) — and it opens the
  six-carrier fixture with no dialog, which the parent matrix treats as a blocking bar.
- ❌ **Numbers 15.3 → 5/6 lost** (run 2026-07-27) — strips the OPC part, customXml,
  docProps, defined names and extLst. **Only cell data survives.** Drive it with
  `open -a Numbers` (the `open` Apple event hangs; LaunchServices does not) then
  `export document 1`.
- ⛔ Excel-Win leg still needs a licensed Windows host.

> **Key finding:** a **recovery record** — model id, dim, dtype, coverage ranges,
> content hash; ~100–200 bytes, *not* the vectors — can be carried durably through a
> tool that erases the vectors. `<definedName>` ranks first: survives both rebuilders
> *and* is enumerated by no Document Inspector module. It was never considered in the
> design doc's "Why NOT" section. Matrix: `docs/plans/emb-4b-carrier-matrix.md`.

**✅ DURABILITY CONTRACT — DECIDED 2026-07-26.** The reconciliation that gated
emb-5/emb-6 is done; §Goals.0 is amended.

> **Excel-durable vectors; evidence that survives every measured consumer except Apple
> Numbers.** zlsx guarantees that a workbook which loses its vectors **says so** — unless
> it went through Numbers, which erases the evidence too.
>
> ⚠️ Corrected 2026-07-27: this read "universally-durable evidence" while the Numbers leg
> was unrun. That claim was false and is withdrawn, not softened.

Rejected **silent best-effort** on the same standard as the row/col edit contract: either
the operation is correct or it refuses, never silently wrong. Rejected **putting vectors
somewhere universally durable** — the only carrier that survives everything *and* could
hold them is cell data, which is user-visible, pollutes the SST, and costs 4× in size to
serve two of four targets.

Ships instead: a ~200-byte **recovery record** (model id, dim, dtype, coverage ranges,
hash digest) in a **hidden `<definedName>`** (primary) + **`docProps/custom.xml`**
(secondary) — both, because their removal mechanisms are disjoint (Document Inspector
strips docProps, not defined names). `hidden="1"` keeps it out of Excel's Name Manager,
which is what narrows Goal 3 rather than breaking it.

Measured encoding requirements a reader MUST tolerate (found by round-tripping, not
assumed): LibreOffice rewrites `hidden="1"` → `hidden="true"`, XML-escapes the payload
(`"x"` → `&quot;x&quot;`), and adds `function="false"` / `vbProcedure="false"`. Match on
`name=`, never on the whole tag.

**⛔ The risk materialised.** Numbers strips both recovery carriers, so a Numbers export
reports `absent`, not `stripped`. It cannot be engineered around: every invisible carrier
dies there, and the only survivor — cell data — is visible by definition, because Numbers
rebuilds from its own document model. **Closed 2026-07-27:** both positions ship. Default is
"invisible, not universal" (Goal 3 intact); `recovery_in_cells = true` adds the cell
carrier and survives Numbers, at the cost of a sheet the user can unhide. Verified against
a real Numbers export in both configurations.

**✅ emb-R SHIPPED (2026-07-27).** `pkg/recovery_record.zig` + both carriers. `embeddings()`
now returns `EmbeddingState` — `present` / `stripped` / `absent` — so a stripped workbook
cannot be mistaken for one that never had vectors. Validated end-to-end: LibreOffice and
openpyxl both destroy the vectors and the provenance still comes back, via `defined_name`;
with `docProps` removed it still recovers, and with the names removed it falls back to
`doc_props`. Each carrier is independently sufficient.

**✅ emb-5 SHIPPED (2026-07-27).** `zlsx.embeddings(path)` returns present / stripped /
absent. Vectors come back as a NumPy `(rows, dim)` float32 array decoded in Zig — one FFI
call per coverage, one implementation of each dtype layout — with `valid_mask()` for
tombstoned rows. On a stripped workbook `vectors()` raises `EmbeddingsStripped` carrying
the recovered model, dim, dtype and ranges, rather than returning an empty array: an empty
array would be exactly the silent-nothing the contract rejects.

**✅ emb-6 (CLI) shipped 2026-07-28** — `--strip` (#132), `--prune` (#133),
`--extract` / `--vectors` (#134). **The embedding arc is complete; what remains
on it is IANA registration and the unrun Excel-Windows leg.**
Full reasoning: `docs/plans/embeddings-in-xlsx.md` §Durability contract.

Spec: `docs/plans/embeddings-in-xlsx.md` · matrix + Findings: `docs/plans/emb-4-compat-matrix.md`.

**Byte-ship decisions:** OPC relationship-URI namespace ✅ **decided 2026-07-28
— `schemas.fabre.me`** (MIME tree moved with it: `application/vnd.fabre.zlsx.*`).
Remaining: IANA MIME registration (deferred to v1.0).

---

## ✅ Done (load-bearing)

- **B2/B3 unification** — Writer + Editor collapsed onto `Workbook`; each subsystem a
  std-only `pkg/*_plan.zig`. `Editor.save` 14 LOC, `Writer.save` 17 LOC.
- **Every row/col-edit refusal axis with a rewriter is lifted** on `main` (panes,
  autoFilter, picture, xdr, VML+comments, structured tables, extLst `xm:sqref` #140).
  Two axes stay refused, now with *actual* guards: **pivots** (#139 — previously
  claimed "refused at consumer level" while no refusal existed anywhere; a row edit
  silently stranded every pivot coordinate) and **`<xm:f>`** sparkline formulas
  (#140 — needs the formula rewriter + sheet-name context).
- **v0.5.0 released 2026-07-30** — first release carrying the Zig 0.16 migration and
  the whole embedding arc (#123–#140). Exercises #138's release automation end-to-end.
- **v0.6.0 released 2026-08-01** — `open_bytes()` at all three layers (#144:
  `Book.openBuffer` / `zlsx_book_open_buffer` / `zlsx.open_bytes`; memory-backed
  `File.Reader`, the borrow ends when the call returns). 5 binaries + 5 official
  wheels + SHA256SUMS; Homebrew tap at 0.6.0.

## 🅿️ Parked / out-of-band

- ~~**Relicense PR #102** (MIT → commercial)~~ — **decided and merged
  2026-08-01**: zlsx is **proprietary** — repository public for reading,
  60-day artifact-only evaluation, commercial license (wrapper-layer
  source rights only; never the Zig core). The prior-versions MIT
  acknowledgment was removed 2026-08-01 at Laurent's direction.
- ~~**Zig 0.16 migration**~~ — shipped: the toolchain in #120, the embedding arc's
  forward-port in #123. `stash@{0}` is discharged and safe to drop.
- **zlsx-cloud SaaS** — separate design arc (`docs/plans/saas-*`, gitignored).

## 🧭 Pro track — Databricks interface (ACTIVE; proofs 2026-07-30 → 08-01)

First pro feature: interface zlsx with Databricks, whose native Excel support is
weak (JVM/POI-based `spark-excel` or driver-side pandas+openpyxl). Ideation lives
in the private KB (`docs/kb/`, gitignored) alongside the SaaS arc; the verified
experiments are in-tree at `integrations/databricks/` (#143).

Proven end-to-end on a live workspace, using released artifacts only:

- **Both directions** — xlsx → Volume → Delta → SQL aggregates, and a warehouse
  table → styled report xlsx (2026-07-30).
- **PySpark Data Source** — `spark.read.format("zlsx")` on serverless compute
  over a Volume workbook, no Delta copy (a ~60-line `ZlsxDataSource` + the
  released aarch64 wheel).
- **`read_xlsx()` UC Python UDF in pure DBSQL** — the wheel loads inside the UDF
  sandbox; the `wb_sales_live` view parses the workbook *file* at query time.
  Liveness proven: edit the file, the next query reflects it — zero
  re-ingestion. Runs shim-free on 0.6.0's `open_bytes()`, this track's first
  code ask (#144).
- **Genie on a live workbook** (2026-08-01) — the Genie space now carries
  `wb_sales_live` in its data sources; asked what the live workbook says right
  now, Genie generated SQL against the view and answered correctly. A
  natural-language room over an Excel file, no landing step.

Shipped since: the Data Source hardening (per-file×sheet partitions, row-range
splits, type widening, the writer half) in #146, the `zlsx dbx` CLI family
(push / pull / genie) over std.http in #147, and the streaming source — Auto
Loader for Excel — in #148. The relicense (#102) is load-bearing here: the
proprietary boundary is the pro tier.

Next, in value order:

- **`Writer.saveToOwnedBuffer` → `to_bytes()`** (Zig / C ABI / Python) so a
  workbook can be produced without a filesystem, and cross-file schema
  inference for the Data Source. Spark parts now serialise in memory and land
  by rename, so no reader sees a partial workbook.
- ~~**`zlsx dbx audit`**~~ — ✅ done: content-hash the zone, diff against a
  manifest and/or a table's ingestion record, report drift / orphan / missing
  as NDJSON with exit 3 on findings. Live-verified against
  `workspace.default.zlsx_sales`.
- ~~**Data Source, second pass**: row-range partitions re-parse~~ — ✅ done:
  `Rows.skipRows` (Zig) → `zlsx_rows_skip` → `Rows.skip` (Python), used by the
  Spark row-range read path. Skipping a row costs **175 ns** against **2470 ns**
  to decode one (ReleaseFast, 200k-row sheet) — 14× on the traversal, 2.4–2.6×
  end-to-end at 16–64 partitions.

  **The quadratic is not gone, it is cheaper.** Each partition still re-opens
  and re-inflates the whole sheet part, and that per-partition cost is now what
  dominates a partitioned read. Removing it needs a shared inflated-sheet cache
  across partitions on one executor — a different piece of work, and one that
  only pays off when partitions of the same file land together.

## 🔭 Candidate follow-ups (value/effort order)

1. ~~Latent silent-corruption fixes (`<sheetView topLeftCell>`, `<selection>`, bare-sheet
   `<sortState>`)~~ — ✅ **done (#125)**. Also caught a fourth, `<sortCondition ref>`,
   unrewritten even inside `<table>` on the path that shipped in #111.
2. ~~**Fuzz the byte-walkers** (`pkg/{table,drawing,vml,sheet}_edit.zig`, ~3000 LOC)~~
   — ✅ **done (#131)**. Found three crashes: `matchTagAt` read one past the end,
   `writeWithReplacedAttr` sliced backwards on an unterminated value, and
   `shiftSingleA1Col/Row` overran a fixed 16-byte buffer. #125 was the argument
   for it, and it was right: that code had four unhandled elements the refusal
   audit's method could not see.
3. ~~`<extLst>` coordinate fixups (`x14:`/`x15:` blocks)~~ — ✅ **done (#140)**:
   `xm:sqref` shifts by leaf-element name; `<xm:f>` refuses via `ExtensionEditUnsafe`.
   Same pass shipped the pivot refusal (#139) — the last guard-less silent-corruption
   class the refusal audit's method could not see.
4. **Route `<xm:f>` through the formula rewriter** — removes #140's
   `ExtensionEditUnsafe` guard, which today refuses row/col edits on any sheet
   carrying a sparkline formula. Needs a sheet-name context `sheet_edit.zig`
   does not have. The highest-value remaining lift. **No longer standalone**
   as of 2026-08-02: the D1 ladder (`goal_formula.md`) builds the parser and
   name-resolution layer this needs, so the route-through is pullable once
   M2 lands and is carried on that plan's M10+ backlog. Doing it before M2
   means building the sheet-name context twice.
5. **Cross-part pivot rewriter** — removes #139's refusal. Bigger lift:
   `<location ref>` + cache field ranges across `xl/pivotTables/*` and
   `xl/pivotCache/*`, a ref graph zlsx has never walked.
6. CDATA-aware shared tag scanner (candidate for `ziglib`).

---

## Where the truth lives

| For… | Look at |
|---|---|
| The work queue (authoritative) | `docs/plans/post-0.2.9-roadmap.md`, `docs/plans/refusal-audit.md` |
| Narrated architecture / conventions / gotchas | `docs/kb/` → `kb.html` |
| Rules for agents editing code | `AGENTS.md` |
| Active-track detail | `docs/plans/embeddings-in-xlsx.md` |
