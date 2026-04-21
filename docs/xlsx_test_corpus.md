# `ziglib.xlsx` — public test corpus

Curated set of publicly-downloadable xlsx files that exercise different parts of the parser. All are free, accessible without auth, and stable (government / open-data portals, long-standing library fixtures). Run `scripts/fetch_test_corpus.sh` to materialize them under `tests/corpus/`.

## Why this matters

The parser was initially validated only against Alfred's BDR xlsx (openpyxl-generated, inline strings only, no sharedStrings.xml). That left whole code paths untested — most notably the shared-string index. This corpus closes those gaps with real-world files.

## Verified corpus (as of 2026-04-21)

All four URLs below returned 200 OK and parsed through `ziglib.xlsx` without errors. Results captured in the "Live validation" section at the bottom.

| # | File | Direct URL | Size | Sheets | sharedStrings | What it tests |
|---|---|---|---|---|---|---|
| 1 | Frictionless `sample-2-sheets.xlsx` | [github.com/.../files/excel](https://raw.githubusercontent.com/frictionlessdata/datasets/main/files/excel/sample-2-sheets.xlsx) | 7 KB | 2 | 18 entries | Smallest multi-sheet with shared strings. Quick sanity. |
| 2 | openpyxl `guess_types.xlsx` | [github.com/fluidware/openpyxl](https://github.com/fluidware/openpyxl/raw/master/openpyxl/tests/test_data/genuine/guess_types.xlsx) | 31 KB | 1 | 3 entries | Mixed cell types — numbers (3.14), dates stored as strings (3/14/15), scientific notation (3E5), inline text. Has a thumbnail. |
| 3 | ph-poi `test1.xlsx` | [github.com/phax/ph-poi](https://github.com/phax/ph-poi/raw/master/src/test/resources/excel/test1.xlsx) | 9 KB | 3 | 3 entries | Sparse/diagonal layout (A1, B2, C3 on different rows), cells containing newlines, printer settings binary blob in the archive. |
| 4 | World Bank Data Catalog | [databankfiles.worldbank.org](https://databankfiles.worldbank.org/public/ddpext_download/world_bank_data_catalog.xlsx) | 60 KB | 2 | **1,144 entries, 143 KB sharedStrings.xml** | Heavy shared-strings path, 161 data rows × 26 cols. First file in this corpus that actually stresses the SST index. |

## Recommended additions (untested, documented for later)

| Category | Source | URL / note |
|---|---|---|
| **Very large** (1M+ rows) | SheetJS formula stress | `https://oss.sheetjs.com/test_files/formula_stress_test.xlsx` (~43 MB, A1:H1048576). Drops RAM/time floor. Note: the SheetJS `test_files` GitHub repo itself was disabled for TOS, but the hosted file at oss.sheetjs.com is still reachable at time of writing. |
| **Dates as serial numbers** | openpyxl `test_data` | Other files under [`fluidware/openpyxl`'s `/openpyxl/tests/test_data/genuine/`](https://github.com/fluidware/openpyxl/tree/master/openpyxl/tests/test_data/genuine) — includes `date_styles.xlsx`, `formats.xlsx`, etc. Requires git-cloning the fork to browse fully. |
| **Rich text runs** | openpyxl, caxlsx rich_text_example | Rich text (bold spans, multi-font) stores multiple `<r><rPr>…</rPr><t>…</t></r>` inside an `<si>`. Our parser concatenates the text runs correctly but drops formatting — acceptable for our use case (distillation reads text, not format). |
| **Formulas + cached values** | BLS Employment Situation | [bls.gov/web/empsit.supp.toc.htm](https://www.bls.gov/web/empsit.supp.toc.htm) publishes xlsx with formulas for monthly aggregates. Cached `<v>` values are what we want; we don't need to evaluate. |
| **Merged cells, hidden rows, styles** | ExcelBench fixtures | [github.com/SynthGL/ExcelBench/tree/main/fixtures](https://github.com/SynthGL/ExcelBench) — 17 xlsx features across 3 tiers (core, formatting, advanced). Clone `ExcelBench/fixtures/excel/` for the full set. |
| **CJK / Arabic / emoji** | ONS / INSEE / govtrack | Hand-pick from national statistics agencies — e.g. [ONS filter-outputs](https://www.ons.gov.uk/filter-outputs) publishes xlsx with non-Latin-1 characters. |

## Live validation (2026-04-21)

```bash
$ /tmp/xlsx_smoke data/worldbank.xlsx "World Bank Data Catalog"
sheets: 2
  - World Bank Data Catalog  →  xl/worksheets/sheet1.xml
  - About World Bank Data Catalog  →  xl/worksheets/sheet2.xml
shared strings: 1144

header (26 cols):
  [ 0] DataCatalog_id
  [ 1] Name
  [ 2] Acronym
  …
total rows: 161
```

All four files decoded without `error.MalformedXml`, `error.BadZip`, or `error.UnsupportedCompression`. UTF-8 preserved (Alfred's em-dash rendering carried through here too). Shared-string indices resolved correctly — no off-by-one seen.

## Known gaps this corpus does NOT cover

| Edge case | Why it matters for `ziglib.xlsx` |
|---|---|
| **`store` (uncompressed) entries** | Every file above uses deflate. Our code has a `.store` branch, but it's untested end-to-end. |
| **Zip64 archives** (>4 GB) | `std.zip.Iterator` handles these, but `extractEntryToBuffer` hasn't been exercised against a >4 GB archive. |
| **Cells with `<f>` (formula) children** | Our parser skips the formula and reads `<v>` (cached value). Tested in theory via the "ph-poi test1" file (which has some cached values) but not in a dedicated fixture. |
| **`t="e"` (error cells)** | Mapped to `.string` holding the error code (`#N/A`, `#REF!`, …). Not yet seen in the corpus. |
| **Namespace prefixes** (e.g. `<x:sst>`) | Some generators prefix the spreadsheetml namespace. Our substring search doesn't account for this. Potential silent-failure mode. |
| **Horizontal images, charts, pivot caches** | Ignored by design — we skip every archive entry except SST, workbook, rels, and worksheets. No regression risk from their presence, but untested. |

## Download script

A reproducible script would live at `scripts/fetch_test_corpus.sh`:

```bash
#!/usr/bin/env bash
set -euo pipefail
dir="${1:-tests/corpus}"
mkdir -p "$dir"
curl -sfL -o "$dir/frictionless_2sheets.xlsx" \
  "https://raw.githubusercontent.com/frictionlessdata/datasets/main/files/excel/sample-2-sheets.xlsx"
curl -sfL -o "$dir/openpyxl_guess_types.xlsx" \
  "https://github.com/fluidware/openpyxl/raw/master/openpyxl/tests/test_data/genuine/guess_types.xlsx"
curl -sfL -o "$dir/phpoi_test1.xlsx" \
  "https://github.com/phax/ph-poi/raw/master/src/test/resources/excel/test1.xlsx"
curl -sfL -o "$dir/worldbank_catalog.xlsx" \
  "https://databankfiles.worldbank.org/public/ddpext_download/world_bank_data_catalog.xlsx"
echo "corpus in $dir:"
ls -la "$dir"
```

Total ~110 KB — small enough to commit to the repo under `tests/corpus/` if we prefer offline tests.

## Next actions (in priority order)

1. **Write a corpus-backed integration test** in `ziglib` that opens each file and asserts expected row counts / header cells. Covers Frictionless + openpyxl + ph-poi + World Bank in one go.
2. **Find and verify a `store`-method xlsx** (no deflate). Construct one via `zip -0` if a real one is hard to find.
3. **Find a namespaced-SST xlsx** (`<x:sst>` prefix) — if it exists in the wild, our `<si` substring match silently fails and we'd report 0 shared strings while iterating cells that point into them. High-severity if real.
4. **Download SheetJS formula stress test** and run through the library — confirms wall-clock on 1M-row workloads and pins the memory ceiling.
