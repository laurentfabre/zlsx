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

| Carrier | Location | zlsx (control) | Excel mac 16.109.2 | openpyxl 3.1.5 | LibreOffice Calc 26.2.3.2 | **Numbers 15.3** | Excel Win |
|---|---|---|---|---|---|---|---|
| `opc_part` | `xl/zlsxEmbeddings/index.xml` | ✅ | ✅ | ❌ | ❌ | ❌ | _pending_ |
| `custom_xml` | `customXml/item1.xml` | ✅ | ✅ | ❌ | ✅ | ❌ | _pending_ |
| `doc_props` | `docProps/custom.xml` | ✅ | ✅ | **✅** | **✅** | ❌ | _pending_ |
| `cell_data` | hidden sheet `zlsxE4B` A1 | ✅ | ✅ | **✅** | **✅** | **✅** | _pending_ |
| `defined_name` | `<definedName ZlsxE4BRecovery>` | ✅ | ✅ | **✅** | **✅** | ❌ | _pending_ |
| `ext_lst` | `xl/workbook.xml` `<extLst><ext>` | ✅ | ✅ | ❌ | ❌ | ❌ | _pending_ |
| | **carriers lost** | 0/6 | **0/6** | 3/6 | 2/6 | **5/6** | — |

Informational: `state="hidden"` on the marker sheet survived every leg,
Numbers included — so the `cell_data` carrier keeps its concealment
through a rebuild. It just never had concealment from the Document
Inspector in the first place.

> **Numbers erases everything except cell data (measured 2026-07-27,
> Numbers 15.3).** This is the result the matrix's "live risk" note
> warned about, and it landed on the bad side. Numbers strips 5 of 6
> carriers — including **both** recovery-record carriers. Only content
> that is part of the workbook model itself survives its export.
>
> Note the version: the emb-4 matrix records Numbers 14.5; the machine
> now runs 15.3, so this is a *newer* build, not an older one, and the
> behaviour is not a regression that a future update will fix.

### How the Numbers leg is driven (it is scriptable after all)

Earlier attempts concluded Numbers could not be automated, on three
`-1712 AppleEvent timed out` failures. That was the wrong conclusion
from a correct observation: the **`open` Apple event** hangs
indefinitely, but the application is otherwise fully responsive —
`get version` and `count of documents` answer instantly.

Opening through LaunchServices instead of Apple events sidesteps it
entirely, and once a document is open, `export` works first time:

```bash
open -a Numbers fixture.xlsx        # LaunchServices, not an Apple event
# poll: osascript -e 'tell application "Numbers" to count of documents'
osascript -e 'tell application "Numbers" to \
  export document 1 to POSIX file "'"$HOME"'/out.xlsx" as Microsoft Excel'
```

Export must target a non-TCC folder (`~/`, not `~/Documents`), and the
destination must not already exist.

**Excel opens the six-carrier fixture without a dialog.** That is not a
throwaway detail: this doc's parent sets "any warning / recovered-file /
removed-features dialog is a blocking failure", and the fixture adds
five carriers of surface beyond what emb-4 tested. The AppleScript leg
is itself the evidence — a modal would have blocked the Apple event,
which is exactly how the Numbers leg fails.

### Legs not yet run

- ~~**Apple Numbers**~~ — **run 2026-07-27**, see the matrix above and
  the runner note. It did change the ranking, decisively.
- **Excel for Windows** — still blocked on a Windows host with Excel,
  exactly as `E4W` is. A GitHub Actions `windows-latest` runner does
  **not** close this: the CI job proves the binary runs on Windows, not
  that Excel preserves anything, and Excel is not installed on hosted
  runners.

---

## Findings

**Three carriers survive openpyxl and LibreOffice:**
`docProps/custom.xml`, cell data, and `<definedName>`. Excel for Mac
preserves all six, reproducing emb-4's verdict on `opc_part` and
confirming that nothing in the wider fixture upsets it.

That is the result emb-4B existed to get. A recovery record *can* be
carried through a tool that erases the vectors, which turns "the data
is gone and nothing recovers it" into "the data is gone and the
workbook still says so".

**But only one carrier survives Numbers, and it is the visible one.**
Numbers strips 5 of 6, including both recovery-record carriers. So the
promise is real for openpyxl and LibreOffice and *false for Numbers* —
a workbook that passes through Numbers is once again indistinguishable
from one that never had embeddings. See the trade-off section below;
this is the single most consequential row in the matrix.

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

**Ranking for a recovery-record carrier**, with every reachable leg run:

1. **`defined_name`** — survives openpyxl and LibreOffice, and is not
   enumerated by any Document Inspector module. Capacity-bounded, which
   a recovery record can live within. **Erased by Numbers.**
2. **`doc_props`** — same survival profile, but Document Inspector ▸
   Document Properties and Personal Information removes it, and that is
   a commonly-run corporate compliance flow. **Erased by Numbers.**
3. **`cell_data`** — **the only carrier that survives Numbers**, and it
   survives everything else too. Also the most visible option by a wide
   margin: Sheet ▸ Unhide reveals it to any curious user, and it
   pollutes the cell grid and the SST.

`defined_name` leading was a genuinely new input to the durability
decision — it was never considered in `embeddings-in-xlsx.md`'s "Why
NOT" section. It remains the right *primary* carrier. What the Numbers
leg changes is that no combination of invisible carriers is universal.

### The trade-off is now explicit, and it is real

Every carrier that is invisible to the user is erased by Numbers. The
only carrier Numbers preserves is the one that is part of the workbook
model — and therefore visible. That is not an implementation gap to
engineer around; it follows from *why* Numbers keeps things at all. It
rebuilds the file from its own document model, so exactly the content
that model represents survives, and nothing else.

So there are two honest positions, and they cannot both be held:

- **Invisible, not universal** (what ships): the record hides from the
  Document Inspector and the Name Manager, and a Numbers round-trip
  erases it. Goal 3 intact; the promise carries an exception.
- **Universal, not invisible**: add cell data as a third carrier, and a
  Numbers round-trip keeps the provenance — at the cost of a hidden
  sheet a user can reveal with Sheet ▸ Unhide. Goal 3 breached.

**Both are now implemented, and the choice is the caller's.**
`setEmbeddingsOpts(..., .{ .recovery_in_cells = true })` adds the cell
carrier; the default leaves it off, so Goal 3 holds unless a caller
deliberately trades it away.

Verified against a real Numbers 15.3 export, not simulated:

| fixture | after Numbers export |
|---|---|
| default carriers | `absent` — vectors *and* evidence gone |
| `--cells` opt-in | `stripped`, `carrier=cell_data`, full provenance |

So the trade-off is real in both directions and measured in both
directions. What the library does *not* do is pick for you: the default
is invisible because that is the stated product goal, and the escape
hatch is one flag away because "your vectors silently vanished" is a
worse outcome for some callers than "there is a sheet you can unhide".

### What this does *not* settle

emb-4B measures **survival**, not the product decision. Excel for
Windows remains unrun (`E4W`) and is the one gap left in the vector
half of the contract.

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
