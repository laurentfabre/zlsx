# emb-4B — Carrier survival matrix

> **What this closes.** emb-4 measured one carrier and got a bad
> answer: `xl/zlsxEmbeddings/*` survives Excel but is erased by Apple
> Numbers and LibreOffice Calc, which rebuild the archive from their
> own model. emb-4B measures whether *anywhere else* in the package
> survives those rebuilds, so a small recovery record can ride in a
> second carrier instead of betting everything on one hiding spot.

---

## Why a second carrier at all

The two failure modes deliberately do not overlap:

| | Erased by archive rebuild | Enumerated by Document Inspector |
|---|---|---|
| `xl/zlsxEmbeddings/*` custom part | **yes** | no |
| Cell data | no | **yes** |

A carrier that is invisible to the Inspector dies in Numbers. A carrier
that survives Numbers is visible to the Inspector. Neither is durable
alone, which is the argument for carrying provenance in two places
rather than choosing one.

**A recovery record is not the vectors.** It is model id, dim, dtype,
coverage ranges and a content hash — roughly 100–200 bytes. That is
what makes the capacity-limited carriers viable here even though the
design doc rejected them as *primary* carriers: `docProps/custom.xml`
is scalar-only and Excel caps strings near 255 chars, and a
`<definedName>` formula is similarly bounded. Neither could ever hold
the vectors. Both can hold enough to tell a reader *that the vectors
were stripped and by roughly what*, so it can re-embed from source
rather than silently returning nothing.

---

## Method

One fixture carries the same marker in six carriers at once, so a
single open → save round-trip per tool measures all six. That matters
most for the GUI-only legs, where the alternative is six manual saves
per tool.

```bash
tests/emb-4b/run-carriers.sh [workdir]      # default /tmp/zlsx-emb4b
```

Markers are `ZLSX-E4B-<CARRIER>-8f3a2c1d`, distinct per carrier so a
tool that relocates or merges content cannot make one carrier's
survival look like another's. The nonce is fixed, not random: the
fixture is byte-stable across runs, so any diff after a leg is
attributable to the consumer and never to the generator.

Detection is byte-level rather than structural — a consumer that
preserves the payload but reindents or re-encodes the containing XML
still counts as preserving it. The one exception is `opc_part`, read
back through `Workbook.embeddings()` so it stays a like-for-like
control against emb-4's verdict.

`zlsx-emb4b-verify` exits with **the number of carriers lost** (0–6);
64 is reserved for usage/IO failure so it can never be mistaken for a
count. The helpers are invoked as installed binaries rather than through
`zig build emb4b-verify`, which would flatten every non-zero count to
build-failure 1.

---

## Result matrix

Verified 2026-07-26 on macOS 26.4, Zig 0.16.0, fixture nonce `8f3a2c1d`.

| Carrier | Location | zlsx (control) | Excel mac 16.109.2 | openpyxl 3.1.5 | LibreOffice Calc 26.2.3.2 | Numbers 14.5 | Excel Win |
|---|---|---|---|---|---|---|---|
| `opc_part` | `xl/zlsxEmbeddings/index.xml` | ✅ | ✅ | ❌ | ❌ | _pending_ | _pending_ |
| `custom_xml` | `customXml/item1.xml` | ✅ | ✅ | ❌ | ✅ | _pending_ | _pending_ |
| `doc_props` | `docProps/custom.xml` | ✅ | ✅ | **✅** | **✅** | _pending_ | _pending_ |
| `cell_data` | hidden sheet `zlsxE4B` A1 | ✅ | ✅ | **✅** | **✅** | _pending_ | _pending_ |
| `defined_name` | `<definedName ZlsxE4BRecovery>` | ✅ | ✅ | **✅** | **✅** | _pending_ | _pending_ |
| `ext_lst` | `xl/workbook.xml` `<extLst><ext>` | ✅ | ✅ | ❌ | ❌ | _pending_ | _pending_ |
| | **carriers lost** | 0/6 | **0/6** | 3/6 | 2/6 | — | — |

Informational: `state="hidden"` on the marker sheet survived every leg
run so far, so the `cell_data` carrier keeps its concealment through a
rebuild — it just never had concealment from the Inspector in the first
place.

**Excel opens the six-carrier fixture without a dialog.** That is not a
throwaway detail: this doc's parent sets "any warning / recovered-file /
removed-features dialog is a blocking failure", and the fixture adds
five carriers of surface beyond what emb-4 tested. The AppleScript leg
is itself the evidence — a modal would have blocked the Apple event,
which is exactly how the Numbers leg fails.

### Legs not yet run

- **Apple Numbers** — the other archive-rebuilding tool, and the leg
  that matters most, since it could change the ranking below. Not
  automatable from here: AppleScript `export … as Microsoft Excel`
  returns `-1712 AppleEvent timed out` on three attempts, including
  inside an explicit 400 s timeout block and with the app activated.
  Diagnosing it needs assistive access for `osascript` that this host
  does not grant. Numbers quits cleanly afterwards, so nothing is
  wedged — the same script shape drives Excel to completion, so this is
  a Numbers-specific scripting fault rather than a harness bug. Run
  manually: open the staged copy, File ▸ Export To ▸ Excel into a
  non-TCC folder (`~/`, not `~/Documents`), then verify.
- **Excel for Windows** — still blocked on a Windows host with Excel,
  exactly as `E4W` is. A GitHub Actions `windows-latest` runner does
  **not** close this: the CI job proves the binary runs on Windows, not
  that Excel preserves anything, and Excel is not installed on hosted
  runners.

---

## Findings

**Three carriers survive both archive-rebuilding consumers measured so
far:** `docProps/custom.xml`, cell data, and `<definedName>`. Excel for
Mac preserves all six, reproducing emb-4's verdict on `opc_part` and
confirming that nothing in the wider fixture upsets it.

That is the result emb-4B existed to get. A recovery record *can* be
carried durably through a tool that erases the vectors, which turns
"the data is gone and nothing recovers it" into "the data is gone and
the workbook still says so" — a materially different product promise.

Two results are worth calling out because they contradict the prior
reasoning:

- **`<extLst>` is stripped by both openpyxl and LibreOffice.** This is
  the extension point ECMA-376 actually sanctions for vendor data, and
  the intuitive first choice for a second carrier. It is not durable
  here. Anyone revisiting this design should not re-derive `extLst` as
  the obvious answer — it was measured and it loses.
- **`customXml/` survives LibreOffice but not openpyxl.** It is
  therefore strictly worse than `doc_props` / `defined_name` on
  durability *and* worse on exposure (Document Inspector ▸ Custom XML
  Data ▸ Remove All targets it by name). The design doc's rejection of
  `customXml/` stands, now on measured rather than predicted grounds.

**Ranking for a recovery-record carrier**, on the evidence so far:

1. **`defined_name`** — survives both rebuilders, and is not
   enumerated by any Document Inspector module. Capacity-bounded, which
   a recovery record can live within.
2. **`doc_props`** — survives both rebuilders, but Document Inspector ▸
   Document Properties and Personal Information removes it, and that is
   a commonly-run corporate compliance flow.
3. **`cell_data`** — survives everything, but is the most visible
   option by a wide margin: Sheet ▸ Unhide reveals it to any curious
   user, and it pollutes the cell grid and the SST.

`defined_name` leading is a genuinely new input to the durability
decision: it was never considered in `embeddings-in-xlsx.md`, whose
"Why NOT" section covers `customXml/`, hidden worksheets,
`docProps/custom.xml` and sidecar files — but not defined names.

### What this does *not* settle

emb-4B measures **survival**, not the product decision. It does not by
itself answer what zlsx promises about workbooks that pass through
Numbers or LibreOffice; it removes the excuse that the answer had to be
a guess. Two legs are still open (Numbers, Excel Win) and Numbers in
particular could change the ranking — it is the tool whose rebuild is
most aggressive.

---

## Cross-references

- `docs/plans/emb-4-compat-matrix.md` — the single-carrier matrix this
  follows up, and its Findings section.
- `docs/plans/embeddings-in-xlsx.md` §Goals.0 — the compat target set,
  and the "Why NOT" alternatives that emb-4B has now measured.
- `tests/emb-4b/carriers.zig` — carrier catalogue and marker scheme.
- `tests/emb-4b/carrier_gen.zig` — fixture producer (two-pass: typed
  writer, then raw `PartStore` injection).
- `tests/emb-4b/carrier_verify.zig` — per-carrier survival report.
