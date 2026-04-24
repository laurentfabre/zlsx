# jq for Excel — streaming CLI design (v2)

> **Status**: design, not yet implemented. See `src/cli.zig` for the current CLI surface (single-command, row-value jsonl/tsv/csv). v2 incorporates Codex's review of v1 — see the "what changed" note at the bottom.

## Goal

Turn `zlsx` into the thing you reach for when piping an xlsx into a Unix pipeline or an LLM harness. Same ergonomics as `jq`:

- Every sub-command emits **uniform-envelope NDJSON** — one record per line, every record carries `kind` + `sheet` + `sheet_idx` so multi-source streams compose cleanly.
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

Every NDJSON record from every sub-command is a JSON object with this minimum shape:

```json
{"kind": "<record-type>", "sheet": "<name>", "sheet_idx": <int>, ...}
```

Rules:
- `kind` is always present and picks the record schema (`"cell"`, `"row"`, `"comment"`, `"validation"`, `"hyperlink"`, `"style"`, `"sst"`, `"sheet"`, `"workbook"`, `"error"`).
- `sheet` + `sheet_idx` are always present for sheet-scoped records — so `jq 'select(.sheet=="Data")'` works uniformly.
- Additional fields are schema-specific (see tables below).
- Downstream consumers can filter / group / merge multi-sheet streams without sniffing positional arrays.

Benefit: a single pipeline can consume `zlsx cells ... & zlsx comments ... & wait` and interleave; `kind` keeps records addressable.

## Sub-command surface

| Command | Record `kind` | Per-line fields |
|---|---|---|
| `zlsx cells <file>` | `"cell"` | `sheet, sheet_idx, ref, row, col, t, v, style?, formula?, runs?` |
| `zlsx rows <file>` | `"row"` | `sheet, sheet_idx, row, cells[]` (each cell: `{ref, col, t, v}`) |
| `zlsx comments <file>` | `"comment"` | `sheet, sheet_idx, ref, row, col, author, text, runs?` |
| `zlsx validations <file>` | `"validation"` | `sheet, sheet_idx, range, kind, op?, formula1, formula2?, values?` |
| `zlsx hyperlinks <file>` | `"hyperlink"` | `sheet, sheet_idx, range, url?, location?` |
| `zlsx styles <file>` | `"style"` | `idx, font, fill, border, num_fmt` *(no sheet — workbook-wide)* |
| `zlsx sst <file>` | `"sst"` | `idx, text, runs?` *(no sheet — workbook-wide)* |
| `zlsx meta <file>` | `"workbook"` + `"sheet"` | workbook record first, then one per sheet |
| `zlsx list-sheets <file>` | `"sheet"` | `name, sheet_idx, rows` — lighter-weight than `meta` |

Short `zlsx <file>` stays as an alias for `zlsx cells <file> --all-sheets`. The existing `--format {jsonl,jsonl-dict,tsv,csv}` on `rows` stays for backward compat; the default output on new commands is pure NDJSON with no format selector.

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

- `--sheet <N|NAME>` — selector accepts either 0-based index or sheet name. No `--name`.
- `--all-sheets` — stream every sheet concatenated (default behavior when `--sheet` is absent).
- `--sheet-glob 'Data*'` — match sheet names against a simple glob.
- `--range A1:Z100` — bounding rectangle for `cells` / `rows`. A1-style, scoped to the current sheet.
- `--header` — on `rows`, treat row 0 as keys and emit `fields` dict per data row.
- `--with-styles` — opt-in metadata on `cells` / `rows`. Off by default.
- `--include-blanks` — emit `t:"blank"` records for empty-but-addressed cells. Off by default.
- `--offset N` / `--limit N` — pagination (renamed from iter54 draft's `--seek/--take`).
- `--pretty` — **only on `meta`** (single-object commands). Not valid on NDJSON streams; emits an error on misuse.
- `--no-provenance` — drop `sheet` / `sheet_idx` fields for tiny single-sheet streams where the consumer doesn't need them. Default is to include.

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

## Rollout plan (iter54+)

Revised after Codex review — ship the streaming primitives and envelope BEFORE multiplying sub-commands, so nothing cements inconsistent schemas.

1. **iter54 — `openLazy` foundation**: refactor `Book.open` into a lazy-core + eager-facade. Keep `Book` signature for existing library callers. Introduce `Book.streamSheet(idx)` that returns a `Rows` iterator without the full pre-load. ~500 LOC.
2. **iter55 — common envelope in the CLI**: wrap every existing output (rows) with `{kind, sheet, sheet_idx, …}`. Bump the existing `--format jsonl` to the new shape; keep `--format legacy-jsonl` as an escape hatch for one release cycle.
3. **iter56 — `cells` sub-command**: per-cell NDJSON with full envelope, `row` / `col` numerics, `t` always present. Date detection via iter46's `Rows.parseDate`. Formula support via the iter22 writer round-trip shape.
4. **iter57 — `meta` + `list-sheets` as NDJSON**: workbook + sheet records. Trivial over existing reader APIs.
5. **iter58 — `comments` / `validations` / `hyperlinks` / `sst` / `styles`**: thin CLI wrappers. ~50 LOC each. Ship together — one iter.
6. **iter59 — pagination + filtering flags**: `--offset`, `--limit`, `--range`, `--header`, `--all-sheets`, `--sheet-glob`, `--include-blanks`, `--with-styles`.
7. **iter60 — error records + `--no-provenance`**: inline `kind:"error"` events, opt-out envelope fields for lean streams.

Every iter ships independently, each under ~500 LOC, each user-observable.

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
zlsx validations data.xlsx | jq 'select(.kind=="list") | {sheet, range, values}'

# Cross-join: "which cells reference a missing sheet in their formula?"
zlsx cells data.xlsx | jq 'select(.t=="formula" and (.formula | test("[A-Z][a-z]+!")))'

# Pipeline with error handling.
zlsx cells huge.xlsx 2>errors.log | jq 'select(.kind!="error")' | ./llm-ingest

# Sheet glob — everything matching "2024-*".
zlsx cells financials.xlsx --sheet-glob '2024-*' | ./my-model
```

## Open design questions

1. **Formula cached values**: `t:"formula"` always has `formula`; `cached` field is present only when Excel stored a cached result. Should we auto-recalculate? **Proposal: no** — zlsx is a reader, not a spreadsheet engine. Callers that need the computed value can shell out to libreoffice / excel.
2. **`--all-sheets` as default?**: if no `--sheet` given, should we stream everything or require explicit opt-in? **Proposal: stream all as default** — matches `jq`'s "operate on all input" spirit. Users who want one sheet pass `--sheet`.
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

## What changed in v2 (Codex review findings)

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

## Summary for the impatient

> Ship `openLazy` first (iter54). Then wrap every CLI output in a uniform `{kind, sheet, sheet_idx, …}` envelope (iter55). Then `cells` (iter56) with always-present `t`, `row`, `col`, and an explicit `error` type. Then the thin wrapper commands together (iter58). Don't build a query DSL — pipe to `jq`.
