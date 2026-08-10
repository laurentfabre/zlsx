# goal_codex.md — the architectural memory programme

_Created 2026-08-11. **Status: PROPOSED — nothing here is agreed.** The
§9.1 ceilings are an owner decision and this file does not change them._

_Provenance: a Codex advisory pass on the §9.1 ceiling → verification of
every claim against the code → **four** rounds of Codex adversarial
review of this plan. R1: 2 CRITICAL / 5 MAJOR, overturning three of my
four "verified" corrections. R2: 0 CRITICAL / 2 MAJOR, catching that I
had never applied this file's own budget rule to its own top-ranked row.
R3: 0 CRITICAL / 1 MAJOR, catching **the same error a third time, one era
further back**. R4: 0 CRITICAL / 2 MAJOR, catching it a fourth time in a
third currency, and deriving the programme-wide bound. Every ruling was
independently re-verified before
acceptance. §10 records all of it, because the pattern is the most useful
thing in this document._

The M10 ladder (M10a…M10p) took first-recalc RSS on the named workload
from 506.7 MiB to **57.77 MiB** — 3.81× the current 15.15 MiB ceiling —
by narrowing records. That seam is worked out. Everything below is a
representation or contract change.

---

## 0. The four rules that govern every row

**The budget rule** (M10n/M10o) — a cut pays only down to the first era
it does not reach. **Apply it against every era, not the adjacent one.**
That mistake was made three times in this file's own drafting.

Era trace at the M10p baseline (peak live 65 181 995):

| era | height | phase | what a cut must clear |
|---:|---:|---|---|
| 10 | 52 684 820 | decode | the true floor for model-held cuts |
| **15** | **60 441 827** | **graph build** | **pre-publication: no result-side cut can lower it** |
| **20** | **65 181 995** | **the drive's publish — the peak** | |
| 21 | 64 074 072 | staging's plateau | drive-exclusive cuts cap here (1 107 923 B) |
| 22 | 63 810 178 | the projection | |
| 23 | 61 127 094 | the plan | |

Three caps follow, and between them they bound everything proposed here:

- **A drive-exclusive cut is capped at 1 107 923 B** (era 21).
- **A publication-side cut is capped at 4 740 168 B** (era 15), because
  era 15 predates publication and no amount of result-side work lowers
  it.
- **§1, §2, §5, §6 and §7 together are capped at 12 497 175 B** (era 10,
  decode — `65 181 995 − 52 684 820`). The model is absent at era 10 (no
  `id21` in its trace; M10o recorded the model as live era 12–23), so no
  model-, result- or staging-held cut lowers it. After §2 spends
  4 740 168, the rest has **7 757 007 B** before decode wins.
- **§4 is the one exception, and it is why §4 is not merely "retracted"
  but unscheduled.** It targets the part bytes, and `id7:7 951 114` **is
  live at era 10**. Per-part eviction *during* decode could lower era 10;
  eviction after it could not. So scheduling §4 requires repricing every
  era, and until its ownership design and eviction point are specified it
  cannot be priced at all — `id7`'s backing request is not a saving
  estimate. The bound above holds for the programme **as scheduled**.

**The instrument rule** (M10p) — the profiler wraps the *backing*
allocator, so for an arena it records the **chunk request**, never the
fill. Precisely: *RSS prices the resident pages attributable to the
removed payload.* A fully-filled, touched, exclusive chunk **can** equal
its RSS contribution; the backing request alone does not prove it.

**The gross/net rule** — a block removed is not a block saved if its
replacement lives in the same era.

**The four-quantities rule** (forced by R3) — every proposal must state
these separately and never conflate them:

1. **gross fill demand removed** (bytes of records that stop existing)
2. **trace-side backing reduction** (which allocation the profiler
   actually loses)
3. **peak-live cap** (the budget rule against every era)
4. **RSS effect** — unknown for arena-held blocks until §3 exists

---

## 1. Component reports — summary mode

**The change.** Replace 89 999 `ComponentReport` records (5 759 936 B,
exactly 64 B each) with prior membership + compact contributions.

**Verified.** `reports`/`previous_reports` are gpa-backed and exact. The
recalc path folds the array to three scalars and frees it before staging
(`recalc_run.zig:470`). On this workload the records are output-only:
M10b's gate decides the no-graph-change case without building anything
(`iterate.zig:769`), so carry-forward is never reached. There is no third
use of `previousReport`; replacement-not-cumulative semantics are correct.

**The four quantities.** Gross 5 759 936 B · **trace-side
`5 759 936 − replacement backing`, UNKNOWN** until the summary
representation is specified — the replacement lives in the same drive
era, and literal 48-byte `graph.Key` membership would retain most of the
block · **peak-live cap 1 107 923 B — drive-exclusive** · RSS
predictable *once the representation is fixed*, being non-arena.

**Net is lower still.** `execute` creates component state before the gate
can decide (`iterate.zig:1097`); rebuild paths can hold current and
previous buffers (`:895`, `:920`, `:1110`, `:1119`); literal `graph.Key`
membership would preserve most of the 48-byte identity cost.

**API constraint.** `Report.components` is read by tests at
`iterate.zig:2102–3056`. Detailed mode stays the **default**; summary is
selected only by `recalc_run`.

---

## 2. In-place publication ⭐ the largest realizable cut

**The change.** For **formula records in the recalc candidate only**,
overwrite the topmost record's value instead of writing a second `Cell`.

**Verified — the transactional assumption holds.**
- `WorkbookEnv` is a per-run candidate (`recalc_run.zig:325`, `defer
  model.deinit()` `:333`); the live workbook is mutated only at
  `Candidate.swap`. The model, computed entries included, is live
  through era 21 — staging receives `&model` (`:497`).
- Candidate discard is sufficient across production refusal,
  `driver.failure` (`:459`), cancellation, staging and transaction
  paths; nothing reads the model after a refusal (`censusRefusal`
  `:462` uses `wb`).
- Convergence keeps its own priors (`Engine.held`, `iterate.zig:534`,
  read `:1250`, compared `:1255`), gated on `settings.iterate`.
- §5.6e's rebuild reads the model after publications deliberately, and
  `merged()` already ranks computed over stored.
- On a formula record `Cell` has no omitted payload field: preserving
  `layer`, `formula_text`, `cache`, `extra` and mutating only `v` is
  complete; `row`/`col` remain the identity.

**Three conditions.**
1. **Formula records only — never a spill tail.** `publishResult` is
   public and its computed-layer `errdefer` (`workbook.zig:9631`) is the
   cleanup `spill.place` requires. The test at `:18009-18017` publishes
   C1 **over B1's spill tail** and requires `retractResult(C1)` to reveal
   it again; overwriting that record in place while preserving `layer`
   and `extra` would corrupt tail ownership.
2. **Candidate-discard is a host policy, not a global replacement.** The
   generic `Host` contract requires infallible retraction
   (`iterate.zig:291`); `run` performs journal rollback on refusal
   (`:493-496`) and fixtures assert real retractions (`:2180-2182`).
   Journal rollback stays the default; recalc opts in.
3. **In-place and candidate-discard are one atomic change.** `journal`
   holds only `CellRef` (`iterate.zig:517`) and `retractResult` only
   removes the computed-layer record (`workbook.zig:9654`).

**The four quantities — and R3's correction.**

| quantity | value |
|---|---:|
| gross fill demand removed (1 563 × 5 128 + 80 000 × 12) | 8 975 064 B |
| **trace-side backing reduction** | **7 720 116 + 1 079 988** |
| **peak-live cap (era 15)** | **4 740 168 B** |
| RSS effect | unknown until §3 |

(The journal *reserves* 89 999 slots — 1 079 988 B of backing — but the
workload publishes 80 000, so its fill is 960 000 B. Capacity is the
trace-side figure; fill is the demand figure. They are different numbers
and this table needs both.)

The trace confirms the mechanism: the model arena is **12 330 726 at
era 15 and 12 338 604 at era 20** — only 7 878 more. The publish's
80 000 extra cells are not growth in the model arena; they are the
separate publication-triggered chunk **`id82` = 7 720 116 B**, which the
all-numeric named workload replaces with nothing
(`synth_f1_mix.zig:42`, `workbook.zig:9455`).

Provisional post-change trace: era 20 → 56 381 891, era 21 → 56 353 956,
**era 15 unchanged at 60 441 827 and therefore the new peak.** So the
realizable peak-live saving is **4 740 168 B**, not 9.1 MB. The
8 015 064 B of removed record demand is fill, and cannot be subtracted
from allocator-boundary era heights at all.

---

## 3. Instrumentation — the fill/capacity split

**The change.** Report per pipeline checkpoint: arena capacity, arena
**used**, exact-block live bytes, sampled RSS.

**Why it is not small.** The profiler records only backing-allocator
allocation lengths (`bench_recalc.zig:352`); `EraRecord` (`:280`) has no
fill or RSS state, and **the backing boundary cannot reconstruct
allocations made inside an arena**. It needs used-byte counters in the
selected arena *owners*, explicit checkpoints, and platform RSS sampling.

**Why it is first.** §2's entire value is fill and arena backing. Without
this, §2 ships blind.

---

## 4. Releasing the decompressed parts — ⚠️ mostly retracted

The original premise is wrong. **Formula text is already owned**:
`sharedFormulaText` does `arena.dupe(u8, …)` (`workbook.zig:10190`+) and
`dupeValue` duplicates text and rich errors (`:9952`). The 7 951 114 B is
an **arena backing request** around a ≈5 294 703 B payload
(`store.zig:1411`), and re-inflating at staging dirties the same pages.
`Part.bytes` is a public lifetime contract (`store.zig:79`, `:320`) with
no legal per-part eviction; staging's `part()` is unpolled
(`recalc_run.zig:1112`, `store.zig:1395`); `openBuffer` retains a
backing-owned copy of the archive (`store.zig:238`), so any benefit is
path-dependent.

**What survives:** an API/lifetime design, using `partControlled` with
the run poller. **Confidence LOW. Unscheduled.**

---

## 5. `input.cells` as a view — parked

A genuine projection of the model (`graph.zig:387`), with M10m's roots
retirement as precedent — but **drive-exclusive**: freed after the drive,
before staging (`recalc_run.zig:478`). Capped at 1 107 923 B, and ≈0 once
§1 has taken that budget. **Parked.**

---

## 6. Columnar model records — the row that reaches era 15

`Cell` 80 B → 32–40 B generically, or 16–24 B as structure-of-arrays here.

**Why it is now the relevant successor to §2.** It is a *whole-model*
representation cut, so it reaches the pre-drive cells already live at
era 15 — the era that caps §2. Nothing else **scheduled** does; §4 would
also reach era 15 and earlier, but it is unscheduled and unpriceable.

**But state it in the right currency.** It reduces model **used bytes**
at era 15 by roughly 4.0–6.4 MB (the 100 010 pre-drive cells are
allocated through `WorkbookEnv.arena`). Its **trace-side** reduction is
unknown: `id21:12 330 726` is arena *backing*, and a narrower record need
not step the chunk ladder at all — which is precisely what M10p
demonstrated. **So the peak after §2 + §6 cannot be named yet.** It may
remain era 15, or move to an earlier era once the ladder steps. Era 10
bounds it either way.

**Risk.** `Sheet` is a directory of 64-entry chunks whose **order is the
contract**, with many readers — `merged`, `advance`, `lowerBound`, `at`,
`next`, the range iterators, the spill host. Largest and least de-risked;
arena-backed, so unpriceable until §3.

---

## 7. Staging into a compressed override or rope

Bears on staging's plateau. **The publication chunk (7 720 116) does NOT
belong here** — it is model-arena publication capacity and is §2's, since
a staging rope cannot remove it.

**Why it is not urgent.** Reducing era 21 alone leaves era 20 untouched,
and after §2 the peak is era 15 rather than staging. Revisit only if a
post-§2 trace makes staging competitive again.

**Prerequisites:** `replacePart` promises `part(name)` immediately
returns contiguous raw bytes (`store.zig:960`); transaction prepare
installs those into a new generation and may parse a replacement view
(`recalc_txn.zig:505`, `:659`); `residentBytes` assumes arena/raw-block
ownership.

---

## 8. Sequencing

R3's arithmetic makes this short: **§3 → §2 → re-profile.**

| # | row | reaches | peak-live cap | predictable? |
|---|---|---|---:|---|
| 1 | **§3** instrumentation | — | — | makes the rest priceable |
| 2 | **§2** in-place publication | eras 20 + 21 | **4 740 168 B** (era 15) | ❌ arena-backed |
| — | **re-profile** | | | mandatory |
| 3 | **§6** columnar model | eras 15 + 20 + 21 | reaches era 15 | ❌ |
| — | **§1** summary reports | era 20 only | 1 107 923 B | ✅ but temporary |
| — | **§7** staging rope | era 21 only | — | parked |
| — | **§5** `input.cells` | era 20 only | ≈ 0 | parked |
| — | **§4** part lifetime | ? | ? | unscheduled |

**Why §1 is no longer first.** It buys at most 1 107 923 B and **does not
change the post-§2 endpoint** — era 15 caps the result either way. It is
worth doing for its own sake (exact, predictable, low risk), not as a
step toward a target.

**Why §7 is not second.** Reducing era 21 alone leaves era 20 untouched;
and after §2 the peak is era 15, not staging.

**Why §6 follows §2.** It is the only proposal that reaches era 15.

## Verification matrix every row must satisfy

File-backed **and** `openBuffer` workbooks; an already-parsed
`Worksheet`; dynamic rebuilds; iterative cycles; spill arrays;
mid-re-inflate cancellation; retention limits; saved-byte identity across
all four workloads × tiny/small/named.

---

## 9. What this reaches

**No total is stated, deliberately.** The two bounded figures are §1's
1 107 923 B and §2's 4 740 168 B, both peak-live caps rather than RSS
predictions; §2's RSS effect is unknowable until §3 measures used
capacity and checkpoint RSS. An earlier draft claimed "≈48.7 MiB" and a
later one "≈9.1 MB realizable" — both withdrawn.

**The programme bound is the number that matters to the ceiling
decision.** Everything proposed here is model- or result-held, and the
model is absent at era 10. So the peak-live floor for this whole
programme is **52 684 820 B** — a maximum reduction of 12 497 175 B from
the M10p baseline. Whether that corresponds to an RSS figure is exactly
what §3 exists to answer (M10p proved the two decouple), but the
qualitative conclusion is safe: **this programme cannot approach 32 MiB,
and reaching further means attacking the decode era, which nothing here
does.**

The advisory floor for a general non-specializing engine was
**10–12 MiB**, which makes 15.15 MiB a *stretch goal* rather than an
impossibility — but reaching it requires a decode-era change on top of
§6, not the result-side work.

**Recommendation to the owner (§9.1 unchanged by this file):** a
**64 MiB hard ceiling**, and **no numeric target yet**. Land §3, then §2,
then re-profile. A target set before that trace exists would repeat the
error that produced 15.15.

---

## 10. What the reviews overturned, and the pattern

**Round 1** — three of four "verified" corrections wrong, all failing the
same way: *I verified that a mechanism existed and inferred its value.*

| claim | ruling | what I missed |
|---|---|---|
| §4 saves ≈3.8 MB | **WRONG** | formula text is already owned; the block is arena backing |
| Summary reports impossible | **WRONG** | carry-forward never runs here; they are output-only |
| In-place can keep rollback | **WRONG** | the journal holds only `CellRef`; the change is atomic |

**Round 2**

| claim | ruling | what I missed |
|---|---|---|
| §1 first, saving 5 759 936 B | **WRONG** | never applied §0's budget rule to it — drive-exclusive, capped at 1 107 923 B |
| In-place preserves `layer`/`extra` | **INCOMPLETE** | corrupts spill-tail ownership; formula-records-only |

**Round 3 — the same error a third time, one era further back**

| claim | ruling | what I missed |
|---|---|---|
| §2 realizes ≈9.1 MB → peak ≈56.09 MB | **WRONG** | **era 15 predates publication**, so no result-side cut lowers it; the cap is 4 740 168 B |
| 8 015 064 B of record demand is the trace-side reduction | **WRONG** | that is *fill*; the trace loses `id82` (7 720 116), the publication chunk |

**Round 4 — the same family again, in a third currency**

| claim | ruling | what I missed |
|---|---|---|
| §6 reaches era 15 | **WRONG CURRENCY** | it reaches era 15 in *used bytes*; its trace-side effect is unknown, because `id21` is arena backing and a narrower record need not step the ladder |
| §1's trace-side reduction is "the same" as its gross | **WRONG** | the replacement lives in the same era; net is `gross − replacement backing` |
| journal contributes 1 079 988 B of demand | **WRONG** | that is reserved *capacity*; fill is 80 000 × 12 = 960 000 |
| "the two caps bound everything proposed" | **INCOMPLETE** | era 10 bounds the programme at 12 497 175 B |

The transferable rules:

> **An arena-backed block's size is not a saving.** The profiler prices
> the chunk request; RSS prices the resident pages attributable to the
> removed payload.

> **Apply the budget rule against every era, not the adjacent one.**
> Three drafts of this file capped a cut at the era below it and missed
> an earlier era that survives the cut untouched. Era 21 caught the first
> two; era 15 caught the third.

> **Name the four quantities separately every time** — gross demand,
> trace-side backing, peak-live cap, RSS. Every error above is one of
> them wearing another's number.
