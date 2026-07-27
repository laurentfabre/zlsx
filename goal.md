# 🎯 Goals — zlsx

> Living goal file. The **north star** plus the **active track** and near-term objectives.
> Detail lives in `docs/plans/`; narrated context + collaboration rules live in the
> knowledge base at `docs/kb/` (open `docs/kb/kb.html`, gitignored).

_Last updated: 2026-07-26_

---

## North star

> **The fastest, safest `.xlsx` library — read *and* mutate-in-place — usable from Zig, C,
> and Python, stdlib-only.** Every row/col edit either rewrites all coordinate-bearing
> parts correctly or refuses with a typed error. Never silently corrupt a workbook.

Performance posture (the bar to hold): 1.7–12× over calamine, 38× over openpyxl,
4× less RAM at MB scale. Write-path strict gate ≤ 7.4 ms.

---

## 🔥 Active track — embeddings in xlsx (on `main` as of #123 + #124)

Store semantic vector embeddings inside `.xlsx` via vendor-namespaced OPC parts under
`xl/zlsxEmbeddings/`. E1/E2/E3/E0 shipped in **#123**; the carrier matrix in **#124**.
(PR #115 is dead — it was cut from pre-migration `main` and was re-applied by hand.)

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
    E5 --> E6["⬅ NEXT: emb-6<br/>CLI: zlsx embed / --prune / --strip"]
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

**▶ NEXT — emb-6 (CLI).**
Full reasoning: `docs/plans/embeddings-in-xlsx.md` §Durability contract.

Spec: `docs/plans/embeddings-in-xlsx.md` · matrix + Findings: `docs/plans/emb-4-compat-matrix.md`.

**Blocking byte-ship decisions:** confirm the OPC relationship-URI namespace
(`laurentfabre.dev` or alternative); IANA MIME registration (deferred to v1.0).

---

## ✅ Done (load-bearing)

- **B2/B3 unification** — Writer + Editor collapsed onto `Workbook`; each subsystem a
  std-only `pkg/*_plan.zig`. `Editor.save` 14 LOC, `Writer.save` 17 LOC.
- **Every row/col-edit refusal axis with a rewriter is lifted** on `main` (panes,
  autoFilter, picture, xdr, VML+comments, structured tables). Only Pivots stay refused
  (no rewriter; the writer never emits them).

## 🅿️ Parked / out-of-band

- **Relicense PR #102** (MIT → PolyForm NC) — parked by decision.
- ~~**Zig 0.16 migration**~~ — shipped: the toolchain in #120, the embedding arc's
  forward-port in #123. `stash@{0}` is discharged and safe to drop.
- **zlsx-cloud SaaS** — separate design arc (`docs/plans/saas-*`, gitignored).

## 🔭 Candidate follow-ups (value/effort order)

1. ~~Latent silent-corruption fixes (`<sheetView topLeftCell>`, `<selection>`, bare-sheet
   `<sortState>`)~~ — ✅ **done (#125)**. Also caught a fourth, `<sortCondition ref>`,
   unrewritten even inside `<table>` on the path that shipped in #111.
2. **Fuzz the byte-walkers** (`pkg/{table,drawing,vml,sheet}_edit.zig`, ~3000 LOC, none
   yet) — now the top candidate, and #125 is the argument for it: that code had four
   unhandled elements and the gap was invisible to the refusal audit's method.
3. CDATA-aware shared tag scanner (candidate for `ziglib`).
4. `<extLst>` coordinate fixups (`x14:`/`x15:` blocks) — the one surface still passing
   through verbatim everywhere, per #125's audit note.

---

## Where the truth lives

| For… | Look at |
|---|---|
| The work queue (authoritative) | `docs/plans/post-0.2.9-roadmap.md`, `docs/plans/refusal-audit.md` |
| Narrated architecture / conventions / gotchas | `docs/kb/` → `kb.html` |
| Rules for agents editing code | `AGENTS.md` |
| Active-track detail | `docs/plans/embeddings-in-xlsx.md` |
