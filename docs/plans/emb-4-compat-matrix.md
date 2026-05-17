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

## Result matrix

| Tool                              | Version | Open      | Save → Reopen | `emb4-verify` exit | Notes |
|-----------------------------------|---------|-----------|---------------|--------------------|-------|
| Microsoft Excel for Mac           |         |           |               |                    |       |
| Microsoft Excel for Windows (365) |         |           |               |                    |       |
| LibreOffice Calc                  |         |           |               |                    |       |
| Apple Numbers                     |         |           |               |                    |       |

Informational (expected to strip — out of v1 scope per the design):

| Tool          | Version | Open | Save → Reopen | `emb4-verify` exit | Notes |
|---------------|---------|------|---------------|--------------------|-------|
| Google Sheets |         |      |               |                    |       |
| openpyxl      |         |      |               |                    |       |

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
