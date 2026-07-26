# 🎯 Goals — zlsx

> Living goal file. The **north star** plus the **active track** and near-term objectives.
> Detail lives in `docs/plans/`; narrated context + collaboration rules live in the
> knowledge base at `docs/kb/` (open `docs/kb/kb.html`, gitignored).

_Last updated: 2026-05-31_

---

## North star

> **The fastest, safest `.xlsx` library — read *and* mutate-in-place — usable from Zig, C,
> and Python, stdlib-only.** Every row/col edit either rewrites all coordinate-bearing
> parts correctly or refuses with a typed error. Never silently corrupt a workbook.

Performance posture (the bar to hold): 1.7–12× over calamine, 38× over openpyxl,
4× less RAM at MB scale. Write-path strict gate ≤ 7.4 ms.

---

## 🔥 Active track — embeddings in xlsx (PR #115)

Store semantic vector embeddings inside `.xlsx` via vendor-namespaced OPC parts under
`xl/zlsxEmbeddings/`. Branch `feat/emb-1a-embedding-part`.

```mermaid
%%{init: {'theme': 'base', 'themeVariables': {'primaryColor': '#1a1a2e', 'primaryTextColor': '#e0e0e0', 'primaryBorderColor': '#00d4ff', 'lineColor': '#00d4ff', 'secondaryColor': '#16213e', 'tertiaryColor': '#0f3460', 'fontFamily': 'monospace'}}}%%
graph LR
    D["design doc ✅"] --> E1["emb-1a/1b ✅<br/>wire-format + manifest"]
    E1 --> E2["emb-2/3a/3b ✅<br/>read + write + rel"]
    E2 --> E4["emb-4 ✅ 3-tool<br/>matrix done<br/>(Win pending host)"]
    E4 --> E4B["emb-4B ◐ carrier matrix<br/>3 carriers survive<br/>both rebuilders"]
    E4B --> RX["⬅ NEXT: reconcile<br/>compat finding in<br/>design doc"]
    RX --> E5["emb-5<br/>Python (NumPy + valid_mask)"]
    E5 --> E6["emb-6<br/>CLI: zlsx embed / --prune / --strip"]
    style E4 fill:#16213e
    style E4B fill:#16213e
    style RX fill:#0f3460
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
- ⛔ Numbers + Excel-Win legs still manual (AppleScript `export` against Numbers hangs
  `-1712` on three attempts; `osascript` lacks assistive access here). Numbers is the
  most aggressive rebuilder and could still change the ranking.

> **Key finding:** a **recovery record** — model id, dim, dtype, coverage ranges,
> content hash; ~100–200 bytes, *not* the vectors — can be carried durably through a
> tool that erases the vectors. `<definedName>` ranks first: survives both rebuilders
> *and* is enumerated by no Document Inspector module. It was never considered in the
> design doc's "Why NOT" section. Matrix: `docs/plans/emb-4b-carrier-matrix.md`.

**▶ NEXT OBJECTIVE — reconcile the compat finding in `embeddings-in-xlsx.md`** (§Goals.0)
before emb-5/emb-6. The decision is now better informed but still a product call: accept
**"Excel-durable, best-effort elsewhere"** as the v1 contract and document it loudly,
**or** carry a recovery record in a second carrier so a stripped workbook still says so
and emb-6 can **recompute from source** (re-embed when vec/hash parts are missing but
covered cells still match — the hash column exists to detect exactly this drift).
emb-4B moved the second option from "would need a hiding spot we don't have" to
"three measured candidates". Then: emb-5 (Python), emb-6 (CLI).

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
- **Zig 0.16 migration** — forward-port preserved in `stash@{0}`; own branch when it lands.
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
