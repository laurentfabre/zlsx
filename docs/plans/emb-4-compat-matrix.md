# emb-4 — Compatibility matrix for `xl/zlsxEmbeddings/*`

> **v1 blocker.** No Python binding (emb-5) or CLI surface (emb-6)
> ships before this matrix runs green on Excel mac, Excel Win,
> LibreOffice Calc, and Apple Numbers. Per the design doc, any
> "open with warning" / "we removed features" / "we recovered the
> file" dialog from any v1 target is a blocking failure that needs
> the design to be amended before code lands.

---

## What this matrix tests

The `setEmbeddings` writer produces five new OPC parts under
`xl/zlsxEmbeddings/` plus mutations to `[Content_Types].xml` and
`xl/_rels/workbook.xml.rels`. ECMA-376 Part 2 §9.1.7 permits but
does not mandate that consumers preserve unknown parts on save;
the matrix is the empirical confirmation that the v1 target tools
behave as the design depends on. Three things are checked per tool:

1. **Open without warning** — the cell grid renders normally, no
   "recovered file" dialog, no "removed features" notice.
2. **Workbook→index rel preserved** — the
   `<Relationship Type=".../embeddings"
   Target="zlsxEmbeddings/index.xml"/>` in
   `xl/_rels/workbook.xml.rels` survives a passive save (open →
   save → close, no edits).
3. **Embedding parts preserved byte-structurally** — the index,
   per-coverage vec.bin / hashes.bin, and content-types overrides
   all still exist; `Workbook.embeddings()` returns a non-null
   view whose `model` / `dim` / `dtype` / coverage ids match the
   fixture.

---

## Procedure

### 1. Generate the fixture

```bash
zig build emb4-fixture -- /tmp/zlsx-emb4.xlsx
```

This writes a 12-part .xlsx (~5 KB) with one sheet ("Items", 5 rows)
and two coverages ("title" on A2:A5, "body" on B2:B5) under
model `emb-4-fixture-v1`, dim 4, dtype `int8-sym-per-vec`.

The fixture is byte-stable across runs — re-running emits the same
bytes, so diffs after the matrix runs come from the consumer tool,
not the writer.

### 2. Per-tool round-trip

For each of the four targets:

1. **Copy** the fixture to a tool-specific path so the runs don't
   stomp each other:
   ```bash
   cp /tmp/zlsx-emb4.xlsx /tmp/zlsx-emb4-<tool>.xlsx
   ```
2. **Open** the copy in the tool. Note the dialog state
   (no-dialog / recovered / warning / refused).
3. **Save** the file in-place via the tool's File→Save menu (NOT
   "Save As" — that resets enough metadata that the test stops
   being about preservation behaviour). For tools that refuse
   to overwrite without "Save As" semantics, save to a sibling
   path and update the verify command.
4. **Close** the workbook in the tool before running step 5; some
   tools hold an exclusive lock that breaks the zlsx reopen.
5. **Verify**:
   ```bash
   zig build emb4-verify -- /tmp/zlsx-emb4-<tool>.xlsx
   ```
   Exit codes:
   - `0` PASS — full preservation
   - `2` PARTIAL — parts survived, fields drifted (model / dim /
         dtype mismatch — flag as a tool-specific bug)
   - `3` STRIPPED — embeddings view missing, no rel — tool removed
         the parts entirely
   - `4` PARTS-ONLY — parts survived but workbook rel stripped
         (degraded; readers must full-scan)
   - `5` ORPHANED REL — rel survived but parts gone (worst case,
         workbook is inconsistent)

### 3. Log the result

Update the matrix table below with the verdict per tool. PASS is
the only acceptable outcome for the v1 critical path; anything
else is a design-doc amendment + design-review cycle before any
follow-up emb slice ships.

---

## Automated runner

`tests/emb-4/run-matrix.sh [workdir]` builds the helpers, generates the
fixture, and drives every leg that needs no GUI (zlsx control, openpyxl,
LibreOffice headless if installed), then stages a per-tool copy for the
GUI-only tools and prints their open → save → verify steps. Re-run it any
time the writer changes.

> **Build path (since the 0.16 migration).** The runner drives the canonical
> `zig build emb4-fixture` / `emb4-verify` / `emb4-passive-save` steps on every
> platform, and prefers the pinned `~/.zvm/0.16.0/zig` over PATH. Under 0.15.2
> the helpers do not compile at all — they use 0.16's `std.process.Init` entry
> point. The old macOS-only standalone `zig build-exe -target aarch64-macos-none`
> fallback (0.15.2's bundled `libSystem` had no `arm64-macos` slice, so the build
> runner would not link) has been removed.

> **LibreOffice headless on macOS 26.4 needs one GUI launch first (2026-05-30).**
> A cold `soffice --convert-to` hangs (empty log) until LibreOffice has completed
> its first-run setup once via the GUI (`open -a LibreOffice`, let it settle, quit).
> After that, headless conversion works. On Linux CI this isn't an issue.

## Result matrix

v1 critical path — **PASS is the only acceptable outcome**:

| Tool                              | Version  | Open  | Save → Reopen | `emb4-verify` exit | Notes |
|-----------------------------------|----------|-------|---------------|--------------------|-------|
| Microsoft Excel for Mac           | 16.109.2 | clean | **preserved** | **0 PASS**         | Verified 2026-05-30. Saved from a trusted location (`~/Documents`); **all 6 zlsxEmbeddings parts + the workbook→index rel survive a passive save.** From `/tmp` Excel opens in Protected View, which blocks save until "Enable Editing". |
| Microsoft Excel for Windows (365) |          | —     | —             | _pending (no host)_| Checked 2026-05-31: no Windows+Excel env reachable from the dev Mac (no local VM; SSH config has only Linux company infra). Run on a Windows host or Windows CI runner, then `emb4-verify` the saved file. |
| LibreOffice Calc                  | 26.2.3.2 | clean | **stripped**  | **3 STRIPPED**     | Verified 2026-05-30. Calc's OOXML export filter rebuilds the archive and drops all zlsxEmbeddings parts + the rel. (Headless `--convert-to` only works after one GUI first-run launch on macOS 26.4 — see runner note.) |
| Apple Numbers                     | 14.5     | clean | **stripped**  | **3 STRIPPED**     | Verified 2026-05-30. File▸Export To▸Excel rebuilds the archive and drops all parts + the rel. (Export must target a non-TCC folder, e.g. `~/`, not `~/Documents`.) |

Control (zlsx itself — confirms the part format round-trips through our own
delta-on-bytes writer, so any strip elsewhere is a property of that tool):

| Tool                  | Version | Open | Save → Reopen | `emb4-verify` exit | Notes |
|-----------------------|---------|------|---------------|--------------------|-------|
| zlsx (`emb4-passive-save`) | 0.15.2 | clean | preserved | **0 PASS** | Verified 2026-05-30 on macOS 26.4 (standalone build). |

Informational (expected to strip — out of v1 scope per the design):

| Tool          | Version | Open | Save → Reopen | `emb4-verify` exit | Notes |
|---------------|---------|------|---------------|--------------------|-------|
| Google Sheets |         |      |               | _pending_          | Manual; documented stripper. |
| openpyxl      | 3.1.5   | n/a  | **stripped**  | **3 STRIPPED**     | Verified 2026-05-30 — rebuilds the archive, drops all 6 `xl/zlsxEmbeddings/*` parts + the workbook→index rel. Matches the design expectation. |

> **Static pre-check (2026-05-30).** Pretty-printing the writer output confirms
> the OOXML-sensitive parts are spec-correct: per-part content-type `<Override>`s
> for the vendor MIME types, a relative `Target="zlsxEmbeddings/index.xml"`
> workbook relationship under the `schemas.laurentfabre.dev/zlsx/2026` namespace,
> and well-formed `index.xml` + `_rels/index.xml.rels`. So the file *should* open
> without warning; whether each consumer *preserves on save* is the empirical
> question the GUI rows close.

## Findings (2026-05-30)

Three of the four v1 targets were validated on macOS 26.4 (Excel-for-Windows still
needs a Windows host). The headline result:

> **Only Microsoft Excel preserves the embeddings on a passive save. Apple Numbers
> and LibreOffice Calc both strip them.** Excel round-trips unknown OPC parts by
> design; Numbers (Export▸Excel) and LibreOffice (Calc OOXML filter) rebuild the
> archive from their own model and drop every `xl/zlsxEmbeddings/*` part plus the
> workbook→index relationship.

| Consumer | Verdict | Mechanism |
|---|---|---|
| Excel for Mac | **PASS** | preserves unknown parts on save |
| Numbers | STRIPPED | rebuilds archive on Export▸Excel |
| LibreOffice Calc | STRIPPED | rebuilds archive via the Calc OOXML filter |
| openpyxl (informational) | STRIPPED | rebuilds archive on save |
| zlsx (control) | PASS | delta-on-bytes writer preserves |

**Design implication.** The matrix's stated v1 bar was *PASS on all four critical
tools*; that bar is **not met** — Numbers and LibreOffice strip. This is the empirical
input emb-4 exists to surface, and per this doc's own rule it warrants a design note
before emb-5/emb-6:

- Embeddings are **durable through Excel** but **best-effort elsewhere** — a Numbers or
  LibreOffice save silently discards them (no warning dialog; the cell grid is fine).
- Because the parts are *removed* (not just unreferenced), no reader-side full-scan
  recovers them after such a save. The data is gone, not hidden.
- Options for the design to weigh (out of scope for emb-4 itself): accept
  "Excel-durable, best-effort elsewhere" as the v1 contract and document it loudly;
  or add a recompute-from-source path (emb-6 can re-embed via `xxh3Canonical` when the
  hashes/vecs are missing but the covered cells still match). The hash column exists
  precisely to detect this drift.

This does not block the *harness* or the writer — both are correct. It blocks declaring
the compat bar "green," and should be reconciled in `embeddings-in-xlsx.md` (§Goals.0
compat set) before the binding/CLI slices.

---

## What to do on failure

- **Open with warning dialog.** Capture the exact text. The
  candidate root causes are content-type override syntax,
  relationship `Target` URI form, or `[Content_Types].xml`
  ordering. Inspect the round-tripped file with `unzip -l` and
  `unzip -p ... [Content_Types].xml` first.
- **Parts stripped (exit 3).** The consumer prunes unreferenced
  parts; either our workbook→index rel didn't survive its rels
  rewrite, or the consumer's rels parser rejected our
  vendor-namespaced Type URI. Compare
  `xl/_rels/workbook.xml.rels` before and after.
- **Rel stripped but parts preserved (exit 4).** Less catastrophic:
  the reader can still find the index via direct part lookup
  (which is what `Workbook.embeddings()` does today). Mark it
  graceful-degraded and move on, but record the tool.
- **Orphaned rel (exit 5).** The consumer kept our rel but
  deleted the parts it points at. Shouldn't be possible under
  any well-behaved OPC consumer; if observed, file a bug
  upstream and amend the matrix's expectation column.

---

## Cross-references

- `docs/plans/embeddings-in-xlsx.md` — full spec (compat target set
  is in §Goals.0).
- `tests/emb-4/fixture_gen.zig` — fixture producer source.
- `tests/emb-4/verify.zig` — verifier source.
- `pkg/workbook.zig::setEmbeddings` — emb-3a writer; emb-3b
  workbook-rel registration + replacePart path.
