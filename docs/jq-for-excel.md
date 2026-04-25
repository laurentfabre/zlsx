# jq for Excel — streaming CLI design (v4.1)

> **Status**: design **shipped** — every rollout slice iter54 → iter60c plus iter61-a/b/c is landed on `main` and Codex-reviewed clean. The sub-command dispatch, envelope, pagination/range/header, output modes, signal handling, inline error-record mechanism, AND cell-type completeness (`t:"date"` / `"error"` / `"formula"` with `formula_ref` + `cached` per shared-formula semantics) are all live. 1900 and 1904 (Mac) date epochs both decode correctly. Four rounds of design-level Codex review (v1 → v4.1) preceded implementation; each code slice went through its own review rounds. The `t` tag now spans all 8 design-doc values: `str` / `int` / `num` / `bool` / `blank` / `date` / `error` / `formula`. See `src/cli.zig` for the implementation and the README's "CLI" section for the user-facing reference. Array formulas spread the `t:"formula"` + `formula_ref` shape across the full `<f t="array" ref="…">` rectangle, matching the shared-formula slave shape.

## Goal

Turn `zlsx` into the thing you reach for when piping an xlsx into a Unix pipeline or an LLM harness. Same ergonomics as `jq`:

- Every sub-command emits **uniform-envelope NDJSON** — one record per line, every record carries `kind`; sheet-scoped records additionally carry `sheet` + `sheet_idx`, so multi-source streams compose cleanly.
- No built-in query DSL. Compose with `jq`, `rg`, `awk`, `duckdb` — the existing CLI toolkit is already excellent.
- Streaming as the **default**, not the fast path. Memory should be a function of SST size, not file size.
- LLM-optimized metadata: every cell carries explicit type + stable numeric row/col, so downstream never re-parses A1.
- Explicit error records for corrupt parts / unsupported features — long-running pipelines should never fail opaquely.

## Non-goals

- **No embedded query DSL.** Users pipe to `jq`. Owning a mini-language is a tar-pit.
- **No async / concurrent output.** One subcommand, one process, one stdout stream. Shell composition handles parallelism.
- **No round-trip / mutate-in-place** from the CLI. Read-only by design; use the library API for writes.
- **No theme / layout preservation** beyond what the reader already surfaces.

## The common envelope

The invariant is precise:

1. **Every record has a `kind` field** that picks the schema (`"cell"`, `"row"`, `"comment"`, `"validation"`, `"hyperlink"`, `"style"`, `"sst"`, `"sheet"`, `"workbook"`, `"error"`).
2. **Sheet-scoped records additionally have `sheet` and `sheet_idx`** (all except `workbook`, `style`, `sst` — which are workbook-wide and have no sheet).
3. **The invariant never varies by flag.** No flag ever drops `kind`, `sheet`, or `sheet_idx` from records that carry them.

Shape:

```json
{"kind": "<record-type>", "sheet": "<name>", "sheet_idx": <int>, ...}   // sheet-scoped
{"kind": "<record-type>", ...}                                           // workbook-wide
```

### Compact mode (opt-in, schema stays stable)

For very large cell streams where `"sheet":"VeryLongSheetName","sheet_idx":0` per record is a measurable token tax, callers can opt into:

```
zlsx cells data.xlsx --output compact-ndjson
```

Which emits a `"sheet"` prologue record before each sheet's cells, then sheet-local records WITHOUT `sheet`/`sheet_idx`:

```jsonl
{"kind":"sheet","sheet":"Data","sheet_idx":0}
{"kind":"cell","ref":"A1","row":1,"col":1,"t":"str","v":"name"}
{"kind":"cell","ref":"B1","row":1,"col":2,"t":"str","v":"qty"}
{"kind":"sheet","sheet":"Other","sheet_idx":1}
{"kind":"cell","ref":"A1","row":1,"col":1,"t":"str","v":"x"}
```

This is a **different schema** that consumers opt into explicitly — not a silent field-drop. Callers that don't pass `--output compact-ndjson` always see the full envelope. The default schema (`--output ndjson`, implicit) is invariant.

## Sub-command surface

| Command | Record `kind` | Per-line fields |
|---|---|---|
| `zlsx cells <file>` | `"cell"` | `sheet, sheet_idx, ref, row, col, t, v, serial?, formula?, formula_ref?, cached?, style?, runs?` |
| `zlsx rows <file>` | `"row"` | `sheet, sheet_idx, row, cells[]` (each cell: `{ref, col, t, v}`) |
| `zlsx comments <file>` | `"comment"` | `sheet, sheet_idx, ref, row, col, author, text, runs?` |
| `zlsx validations <file>` | `"validation"` | `sheet, sheet_idx, range, rule_type, op?, formula1, formula2?, values?` |
| `zlsx hyperlinks <file>` | `"hyperlink"` | `sheet, sheet_idx, range, url?, location?` |
| `zlsx styles <file>` | `"style"` | `idx, font, fill, border, num_fmt` *(no sheet — workbook-wide)* |
| `zlsx sst <file>` | `"sst"` | `idx, text, runs?` *(no sheet — workbook-wide)* |
| `zlsx meta <file>` | `"workbook"` + `"sheet"` | workbook record first, then one per sheet |
| `zlsx list-sheets <file>` | `"sheet"` | `name, sheet_idx, rows` — lighter-weight than `meta` |

Short `zlsx <file>` stays as an alias for `zlsx cells <file>` — which, per the v3-revised default, streams **the first sheet only**. Users who want every sheet pass `--all-sheets` explicitly (on either form). The existing `--format {jsonl,jsonl-dict,tsv,csv}` on `rows` stays for backward compat; the default output on new commands is pure NDJSON with no format selector.

## Record shapes

### `cells` — the LLM-optimized stream

```jsonl
{"kind":"cell","sheet":"Data","sheet_idx":0,"ref":"A1","row":1,"col":1,"t":"str","v":"name"}
{"kind":"cell","sheet":"Data","sheet_idx":0,"ref":"B1","row":1,"col":2,"t":"str","v":"qty"}
{"kind":"cell","sheet":"Data","sheet_idx":0,"ref":"A2","row":2,"col":1,"t":"str","v":"apple"}
{"kind":"cell","sheet":"Data","sheet_idx":0,"ref":"B2","row":2,"col":2,"t":"int","v":3}
{"kind":"cell","sheet":"Data","sheet_idx":0,"ref":"B3","row":3,"col":2,"t":"date","v":"2024-06-15T00:00:00","serial":45458}
{"kind":"cell","sheet":"Data","sheet_idx":0,"ref":"C2","row":2,"col":3,"t":"formula","formula":"A2+B2","cached":30}
{"kind":"cell","sheet":"Data","sheet_idx":0,"ref":"D2","row":2,"col":4,"t":"error","v":"#DIV/0!"}
```

`t` values (always present, human-readable):
- `"str"` — string
- `"int"` — integer
- `"num"` — non-integer number
- `"bool"` — boolean
- `"date"` — numeric serial with a date-styled format (`v` = ISO-8601, `serial` = raw Excel serial for callers that want it)
- `"formula"` — cell carries a formula (`formula` always present; `cached` present when the file stored a cached value, absent if unavailable)
- `"error"` — Excel error cell (`v` = literal error string like `"#DIV/0!"`, `"#N/A"`, `"#REF!"`)
- `"blank"` — empty cell that has a style or was referenced. By default blank cells are not emitted at all; `--include-blanks` flips this.

Opt-in with `--with-styles`:
```jsonl
{"kind":"cell","sheet":"Data","sheet_idx":0,"ref":"A1","row":1,"col":1,"t":"str","v":"name","style":{"bold":true,"fg":"FFFFFFFF","bg":"FF1F4E79"}}
```

Style shorthand (terse to save tokens):
- `bold`, `italic` — booleans, emitted only when `true`
- `fg` / `bg` — ARGB hex strings, emitted only when set (theme-resolved via iter52)
- `nf` — number format code string, emitted only when not the built-in "General"
- `border` — `{l, r, t, b}` with `{s, c}` (style, color ARGB), emitted only when any side is set

### `rows` — flat tabular extract

```jsonl
{"kind":"row","sheet":"Data","sheet_idx":0,"row":1,"cells":[{"ref":"A1","col":1,"t":"str","v":"name"},{"ref":"B1","col":2,"t":"str","v":"qty"}]}
{"kind":"row","sheet":"Data","sheet_idx":0,"row":2,"cells":[{"ref":"A2","col":1,"t":"str","v":"apple"},{"ref":"B2","col":2,"t":"int","v":3}]}
```

`--header` flag promotes the first row to keys per sheet and flattens subsequent rows into dicts:
```jsonl
{"kind":"row","sheet":"Data","sheet_idx":0,"row":2,"fields":{"name":"apple","qty":3}}
```

When `--header` is set, the header row itself is NOT emitted — consumers want just the data dicts.

### `meta` — workbook summary (still NDJSON)

```jsonl
{"kind":"workbook","path":"data.xlsx","sheets":2,"sst":{"count":1144,"rich":12},"has_styles":true,"has_theme":true,"has_comments":true,"format_version":"0.2.6"}
{"kind":"sheet","sheet":"Data","sheet_idx":0,"rows":161,"cols":26,"first_cell":"A1","last_cell":"Z161","has_comments":true}
{"kind":"sheet","sheet":"Other","sheet_idx":1,"rows":3,"cols":5,"first_cell":"A1","last_cell":"E3","has_comments":false}
```

A workbook record followed by one sheet record per sheet. All records share the common envelope (`kind`), so `zlsx meta ... | jq 'select(.kind=="sheet")'` is a trivial filter.

### `error` — failure events in the stream

```jsonl
{"kind":"error","sheet":"Data","sheet_idx":0,"scope":"sheet","code":"MalformedXml","message":"unterminated <c> at byte 12345"}
```

Emitted inline when a non-fatal parse error hits. Fatal errors exit non-zero with a final `error` record on stderr. Pipelines can `jq 'select(.kind!="error")'` to strip.

## CLI flag conventions

- `--sheet NAME` — select a sheet by name (string). Excel-native semantic.
- `--sheet-index N` — select by 0-based index. Escape hatch for scripting where sheet names aren't known.
- `--all-sheets` — stream every sheet concatenated.
- `--sheet-glob 'Data*'` — match sheet names against a simple glob.
- **Default (no `--sheet*` flag)**: first sheet only. Users who want all sheets pass `--all-sheets` explicitly. This mirrors Excel's "open the first sheet" default and avoids surprising large outputs.
- `--range A1:Z100` — bounding rectangle for `cells` / `rows`. A1-style, scoped to the current sheet.
- `--header` — on `rows`, treat the first row (1-based `row:1`) as keys and emit `fields` dict per data row. Consistent with `row:1` / `--start-row` elsewhere; no 0-based row addressing exists in the surface.
- `--with-styles` — opt-in metadata on `cells` / `rows`. Off by default.
- `--include-blanks` — emit `t:"blank"` records for empty-but-addressed cells. Off by default.
- `--skip N` / `--take N` — stream-native pagination (symmetric, `jq`-native phrasing; replaces the DB-flavored `--offset` / `--limit`). Scoping: `--skip` / `--take` apply **globally over the concatenated output stream**, AFTER sheet selection (`--sheet`, `--sheet-index`, `--sheet-glob`, `--all-sheets`) and any `--range` / `--start-row` / `--end-row` filters. Under `--all-sheets`, `--skip 1000 --take 500` takes records 1001–1500 across the full cross-sheet stream, not per sheet. This matches `head`/`tail`/`jq` stream semantics; callers wanting per-sheet windows run one process per sheet.
- `--start-row R` / `--end-row R` — alternative row-bounded pagination on `rows` / `cells` when callers think in 1-based sheet rows. Applied **per sheet** (each sheet's own rows), unlike `--skip` / `--take` which are global.
- `--output {ndjson|compact-ndjson|pretty-json}` — output mode.
  - `ndjson` (default) — invariant-envelope NDJSON stream.
  - `compact-ndjson` — sheet-prologue variant described above.
  - `pretty-json` — valid only on `meta`. When set, `meta` collapses its workbook + sheets NDJSON into **one** pretty-printed JSON object of the shape `{"kind":"workbook", …, "sheets":[ {"kind":"sheet", …}, … ]}`. This is an explicit different schema (single document, not a stream), flagged by the user. Streaming commands reject `pretty-json` with exit 1.

## Operational guarantees

Explicit behavior for the rough edges of real workbooks and real pipelines:

### Encoding & input validation

- **Charset**: zlsx produces UTF-8 output only. OOXML is spec-UTF-8, but some generators emit UTF-16 or Windows-1252 despite the header claiming UTF-8. On decode failure, the offending bytes are replaced with U+FFFD and an inline `kind:"error"` record is emitted naming the offending part + byte offset; processing continues.
- **Invalid XML control characters** (C0 controls outside the OOXML allowlist, U+FFFE, U+FFFF): JSON serialization always escapes these — U+0000..U+001F and U+007F are emitted as `\u00XX` escapes, and `U+FFFE`/`U+FFFF` are replaced with U+FFFD. No raw control byte ever reaches stdout; every line is valid JSON. There is no "pass-through" mode.
- **Max cell text length**: capped at `2^28` bytes (256 MB per cell). Exceeds-cap cells emit a `kind:"error"` record and are skipped.
- **Max run count per rich-text cell**: capped at `2^16` runs per SST entry. Same error-record behavior on overflow.

### ZIP decompression limits

All three thresholds below are **fatal and consistent**: on breach, zlsx emits one `kind:"error"` record on stderr, flushes in-flight stdout, closes the archive, and exits with code 4 (`ZipBombSuspected`). No partial output is "salvaged" from a suspected bomb — that would defeat the guarantee.

- **Hard cap on any decompressed part**: `2^32` bytes (4 GB).
- **Per-entry compression ratio**: capped at 10,000:1 (matches zlib's recommendation for bomb detection).
- **Total decompressed size across the archive**: `2^34` bytes (16 GB).

### Backpressure & signal handling

- **SIGPIPE**: `zlsx cells huge.xlsx | head -10` exits cleanly. The CLI installs a SIGPIPE handler that sets an internal "stop streaming" flag; the current sheet's remaining rows are abandoned, partial output is flushed, and `exit 0` follows. No broken-pipe traceback noise.
- **SIGINT / SIGTERM**: in-flight records are flushed, then exit 130 / 143 respectively. If the signal arrives mid-record, the partial record is discarded (not emitted) so the stream stays valid NDJSON.
- **Flush policy**: every record is written with an explicit newline + flush on `stdout`. Pipelines always see records in emission order, no coalescing.
- **stderr format**: plain text lines (not JSONL) for the human reader. `stdout` inline `kind:"error"` records are for pipelines; `stderr` is the same content as free-form English for a human tail-f.

### Exit codes

| Code | Meaning |
|---|---|
| 0 | Success. `kind:"error"` records may still have been emitted inline — non-fatal parse issues. |
| 1 | Bad CLI arguments. |
| 2 | Could not open file / not a valid xlsx archive. |
| 3 | Sheet not found (by name / index / glob). |
| 4 | Decompression limit exceeded (`ZipBombSuspected`). |
| 5 | OS error (permission denied, disk full on stdout, etc.). |
| 130 | SIGINT. |
| 143 | SIGTERM. |

### Formula and external-reference handling

- **Shared formulas** (`<f t="shared" ref="…" si="…">`): the base cell carries the expanded formula text in its `formula` field. Dependent cells do NOT repeat the formula; instead they emit an explicit `"formula_ref": "A1"` field (string, A1-style, the base cell's ref) alongside any `cached` value. Consumers can look the base formula up in the earlier record or recompute from the pattern. Example:
  ```jsonl
  {"kind":"cell","sheet":"Data","sheet_idx":0,"ref":"C2","row":2,"col":3,"t":"formula","formula":"A2+B2","cached":30}
  {"kind":"cell","sheet":"Data","sheet_idx":0,"ref":"C3","row":3,"col":3,"t":"formula","formula_ref":"C2","cached":47}
  ```
- **External references** (`[1]Sheet1!A1`): emitted as literal text in `formula`. `zlsx` does not resolve across workbooks.
- **Malformed formulas**: invalid XML inside `<f>` emits `kind:"error"` inline and the cell falls back to its cached value if one exists.

### Determinism

- **Record order within a sheet**: top-to-bottom, left-to-right. Matches the OOXML document order of the source `<row>` / `<c>` elements.
- **Record order across sheets** (`--all-sheets`): workbook's declared sheet order.
- **No implicit sorting**: `zlsx` never reorders cells or rows. Pipelines that need sort use `jq -s 'sort_by(…)'`.

## Architecture

### Current reality

`Book.open` eagerly decompresses every part. For the 67 KB test fixture that's fine; for 100 MB+ SEC filings or finance exports it's memory-unfriendly.

### Target: streaming first

```
zlsx.cli.main()
  → xlsx.Book.openLazy(path)   ← central dir + SST + styles.xml + theme.xml only
  → per subcommand:
     → stream the relevant part (one sheet at a time for cells/rows)
     → emit NDJSON record-by-record, each flushed before the next
     → release per-sheet buffers between sheets (when --all-sheets)
```

`openLazy`:
- Central dir: already streaming via `std.zip.Iterator` — O(1) per entry.
- Eagerly load: `sharedStrings.xml` (cells reference by index), `styles.xml` + `theme1.xml` (style lookup is random-access), per-sheet rels (hyperlinks / comments need them).
- Lazily load: sheet XML parts, comments parts, vmlDrawings.
- Scope: ~500 LOC refactor. Current `Book` stays as a facade over `BookLazy` for library callers who want eager behavior.

### Sub-command dispatch

```
zlsx <subcommand> <file> [flags]
```

Parser is two-phase: first token picks the sub-command (or defaults to `cells` / `rows` for backward compat), then delegates.

## Rollout plan (iter54+) — **SHIPPED**

Revised after Codex review — ship the streaming primitives and envelope BEFORE multiplying sub-commands, so nothing cements inconsistent schemas. Every slice landed on `main`; Codex signed off on each.

1. **iter54 — `openLazy` foundation** ✅ (A scaffolding + B lazy extraction + C `streamSheet(idx)`): `Book.openLazy` / `preloadSheet` / `streamSheet` all public on `Book`; `Book.open` preserved as the eager facade that closes the file handle on return.
2. **iter55 — common envelope in the CLI** ✅ (55a envelope + 55b README/deprecation): `--format jsonl` is the uniform `{kind,sheet,sheet_idx,row,cells:[…]}` stream; `legacy-jsonl` / `legacy-jsonl-dict` carry the pre-iter55a shapes; `jsonl-dict` alias emits a stderr deprecation warning.
3. **iter56 — `cells` sub-command** ✅: `zlsx cells <file>` emits per-cell NDJSON with `{kind,sheet,sheet_idx,ref,row,col,t,v}` and sparse empties. Date / formula / error cells are still basic types (`str`/`int`/`num`/`bool`) pending reader surface extensions.
4. **iter57 — `meta` + `list-sheets` as NDJSON** ✅: workbook + per-sheet records; legacy `--list-sheets` flag preserved for plain-text callers.
5. **iter58 — `comments` / `validations` / `hyperlinks` / `sst` / `styles`** ✅ (5 wrappers in one slice).
6. **iter59 — pagination + filtering flags** ✅ (shipped as 59a / 59b-1..4 / 59c): `--skip`, `--take`, `--start-row`, `--end-row`, `--range`, `--header`, `--all-sheets`, `--sheet-glob` (UTF-8 `?`), `--include-blanks`, `--with-styles`. `--sheet` / `--sheet-index` covered by iter55+.
7. **iter60 — operational guarantees + error records + compact schema** ✅ (60a signals/exit codes + 60b `--output` modes + 60c inline errors).

Every iter shipped independently, each user-observable. See `git log --oneline -- src/cli.zig src/xlsx.zig` for the commit chain.

## Example pipelines

```bash
# All cells across all sheets, only strings, piped to an LLM-friendly ingester.
zlsx cells data.xlsx | jq 'select(.t=="str") | {sheet, ref, v}'

# Date columns only, reformatted as YYYY-MM-DD.
zlsx cells data.xlsx | jq 'select(.t=="date") | .v | .[:10]'

# Sum a column from the CLI without loading everything.
zlsx cells data.xlsx --range B2:B1000 | jq -r 'select(.t=="int" or .t=="num") | .v' | awk '{s+=$1} END {print s}'

# Every comment across every sheet, as TSV.
zlsx comments data.xlsx | jq -r '[.sheet, .ref, .author, .text] | @tsv'

# Grep SST for emails.
zlsx sst data.xlsx | jq -r '.text' | rg '@\S+\.\S+'

# Schema check: every list-type validation + its cell range.
zlsx validations data.xlsx | jq 'select(.rule_type=="list") | {sheet, range, values}'

# Cross-join: "which cells reference a missing sheet in their formula?"
zlsx cells data.xlsx | jq 'select(.t=="formula" and (.formula | test("[A-Z][a-z]+!")))'

# Pipeline with error handling.
zlsx cells huge.xlsx 2>errors.log | jq 'select(.kind!="error")' | ./llm-ingest

# Sheet glob — everything matching "2024-*".
zlsx cells financials.xlsx --sheet-glob '2024-*' | ./my-model
```

## Open design questions

1. **Formula cached values**: `t:"formula"` always has `formula`; `cached` field is present only when Excel stored a cached result. Should we auto-recalculate? **Proposal: no** — zlsx is a reader, not a spreadsheet engine. Callers that need the computed value can shell out to libreoffice / excel.
2. ~~**`--all-sheets` as default?**~~ **Resolved in v3**: default is first sheet only; `--all-sheets` is explicit opt-in. The `jq`-style "operate on all input" argument was outweighed by Excel-user surprise on multi-sheet workbooks where the first sheet is typically the "main" sheet and the rest is support/scratch. (The short alias `zlsx <file>` was re-pointed in v4 to match — see the CLI section.)
3. **Error record placement**: inline in stdout or only stderr? **Proposal: both** — emit to stdout (callers can filter via `jq`) AND to stderr (scripts that care about failure can grep). The stderr copy drops sheet/sheet_idx provenance since stderr is unordered.
4. **Styles identity**: do cells carry `style:{bold:…}` (inlined) or `style_idx:42` with a separate `zlsx styles` stream for the lookup table? **Proposal: inline** — keeps each cell record self-contained, avoids pipeline composition order. Callers who care about style dedup can do it in jq.
5. **Large SST in memory**: for 500 MB workbooks with 10M SST entries, the SST pre-load blows RAM. **Proposal: acceptable for iter54-60**; mitigate in a later iter by streaming SST + building an on-disk mmap index.
6. **Rich-run storage**: `runs?` on cell records or separate `zlsx sst`? **Proposal: both** — `runs` inline opt-in via `--with-styles`, always queryable via `zlsx sst`.

## Why not embed a query DSL?

Same answer as v1, still correct:

1. `jq` already works. Every user has it. Zero lift.
2. Re-implementing `select`, `map`, `@tsv`, etc. is a multi-year project for no gain.
3. The shape of our output — uniform-envelope NDJSON with `kind` discriminators — is the leverage. `jq` operates on this shape natively.
4. If we ever want a query DSL, it goes in a separate binary (`zlsxq`?) that sits between `zlsx` and `jq`.

## What changed across review rounds

### Round 1: v1 → v2

| v1 claim | v2 resolution |
|---|---|
| `rows` emits bare arrays — breaks the self-describing promise | Every record carries `{kind, sheet, sheet_idx, row}`; `rows` emits `{kind:"row", cells:[…]}` |
| `cells` missing `row` / `col` numerics | Added — consumers never re-parse A1 |
| `t:"e"` overloaded for blank + error | Split: `"blank"` vs `"error"` |
| `t:"f"` ambiguous `v`/`formula` | Renamed `t:"formula"` with explicit `formula` + `cached?` fields |
| Optional `t` "when unambiguous" | `t` always present |
| `--pretty` on NDJSON | Removed from streaming cmds; only on `meta` |
| `--seek` / `--take` | Renamed `--offset` / `--limit` |
| `--sheet N` + `--name X` | Unified: `--sheet` accepts either |
| 0-indexed sheets surprising for Excel users | `--sheet` name works; index is the escape hatch |
| `openLazy` scheduled late | Moved to iter54 (foundation) |
| Missing multi-sheet composition | `--all-sheets`, `--sheet-glob`, `sheet_idx` in every record |
| `meta` as single-object JSON | Now NDJSON: workbook + sheet-per-record |
| No error events in stream | New `kind:"error"` inline records with scope + code + message |

### Round 2: v2 → v3

| v2 problem | v3 resolution |
|---|---|
| `--no-provenance` makes the envelope schema conditional | Removed. Replaced with `--output compact-ndjson` — a different explicit schema with a sheet-prologue record. Default schema stays invariant. |
| Invariant "every record has kind+sheet+sheet_idx" is false for workbook-wide records | Tightened wording: "every record has `kind`; sheet-scoped records additionally have `sheet` + `sheet_idx`". |
| `validations.kind` collides with the envelope `kind` discriminator | Renamed the validation subtype field to `rule_type`. |
| Production-gap silence (UTF-8, ZIP bombs, signals, SIGPIPE, exit codes, shared formulas) | New "Operational guarantees" section. |
| `--offset` / `--limit` is DB language in a stream context | Renamed `--skip` / `--max-records` *(v3 intermediate — renamed again to `--take` in v4 for symmetry, see round-3 table)*; added `--start-row` / `--end-row`. |
| `--sheet` accepting both name + index is ambiguous | Split: `--sheet NAME` + `--sheet-index N`. |
| Defaulting to `--all-sheets` surprises Excel users with huge output | Default is first sheet; `--all-sheets` is explicit opt-in. |

### Round 3: v3 → v4

| v3 problem | v4 resolution |
|---|---|
| "Open design question #2" still proposed `--all-sheets` as default — directly contradicted the v3 CLI section | Resolved in-place (first sheet only); question now records the decision + why. Short `zlsx <file>` alias re-pointed to `cells <file>` (was `cells --all-sheets`). |
| Rollout iter59 still listed `--offset` / `--limit` | Replaced with `--skip` / `--take` / `--start-row` / `--end-row` / `--sheet` / `--sheet-index`. |
| Rollout iter60 still mentioned `--no-provenance` | Replaced with `--output compact-ndjson` (sheet-prologue) and `--output pretty-json` (meta single-object). |
| Op-guarantees "invalid XML chars passed through as UTF-8" was unsafe — could emit raw control bytes and break downstream JSON parsers | Tightened: U+0000..U+001F and U+007F are always JSON-escaped (`\u00XX`); U+FFFE/U+FFFF replaced with U+FFFD. No pass-through mode. |
| ZIP-bomb thresholds inconsistent: oversized part was fatal but per-entry ratio was "skip" — weakens the guarantee | All three thresholds (part-size, per-entry ratio, total decompressed size) are now uniformly fatal with exit 4. |
| `--header` docced as "row 0" but everything else in the doc is 1-based | Changed to "first row (1-based `row:1`)". |
| `pretty-json` output mode conflicted with `meta` being NDJSON | Redefined: `pretty-json` is valid only on `meta`, and collapses the NDJSON stream into a single pretty-printed `{workbook, sheets:[…]}` JSON object — an explicit alternative schema, not a mode toggle on the NDJSON stream. |
| `--skip` / `--max-records` pagination pair was asymmetric | Renamed to `--skip` / `--take` (jq-native, symmetric). |
| `formula_ref:<base-ref>` written in prose looked like notation, not a field | Made explicit: `"formula_ref": "A1"` field on dependent cells, with example. |

## Summary for the impatient

> Ship `openLazy` first (iter54). Then wrap every CLI output in a uniform `{kind, sheet, sheet_idx, …}` envelope (iter55). Then `cells` (iter56) with always-present `t`, `row`, `col`, and an explicit `error` type. Then the thin wrapper commands together (iter58). Don't build a query DSL — pipe to `jq`.
