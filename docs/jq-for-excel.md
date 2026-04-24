# jq for Excel — streaming CLI design (draft)

> **Status**: design draft, not implemented. See `src/cli.zig` for the current CLI surface (single-command, row-value jsonl/tsv/csv).

## Goal

Turn `zlsx` into the thing you reach for when piping an xlsx into a Unix pipeline or an LLM harness. Same ergonomics as `jq`:

- Sub-commands emit **self-describing NDJSON** — one record per line.
- Each sub-command addresses a single concern (rows, cells, comments, validations, styles, SST).
- No built-in query DSL. Compose with `jq`, `rg`, `awk`, `duckdb` — the existing CLI toolkit is already excellent.
- Streaming where possible. Fall back to eager load when OOXML's archive layout forces it.
- LLM-optimized metadata: every cell carries type + optional style ref, so downstream can reason without re-guessing.

## Non-goals

- **No embedded query DSL.** Users can pipe to `jq`. Owning a mini-language is a tar-pit.
- **No async / concurrent output.** One subcommand, one process, one stdout stream. Shell composition handles parallelism.
- **No round-trip / mutate-in-place** from the CLI. Read-only by design; use the library API for writes.
- **No theme / layout preservation** beyond what the reader already surfaces.

## Sub-command surface

| Command | Per-line record | When |
|---|---|---|
| `zlsx rows <file> [--sheet N\|--name X]` | `[v, v, v, …]` or `{"A": v, …}` if `--header` | Flat tabular extract for CSV-style consumers |
| `zlsx cells <file> [--sheet] [--range A1:Z100]` | `{"sheet", "ref", "t", "v", "style_ref"?}` | LLM-optimized per-cell stream |
| `zlsx sst <file>` | `{"idx", "text", "runs"?}` | Dedup'd string table — cheap to stream |
| `zlsx comments <file> [--sheet]` | `{"sheet", "ref", "author", "text", "runs"?}` | Iter34 + iter53 surface |
| `zlsx validations <file> [--sheet]` | `{"sheet", "range", "kind", "op"?, "formula1", "formula2"?, "values"?}` | Iter24 surface |
| `zlsx hyperlinks <file> [--sheet]` | `{"sheet", "range", "url"?, "location"?}` | External + internal |
| `zlsx styles <file>` | `{"idx", "font", "fill", "border", "num_fmt"}` | Iter28-32 + iter52 surface |
| `zlsx meta <file>` | Single JSON object: sheets, counts, parts | Workbook-level summary |
| `zlsx list-sheets <file>` | Line per sheet name | Already exists — keep |

Existing `zlsx <file> --format {jsonl,jsonl-dict,tsv,csv}` stays as an alias for `zlsx rows`.

## Record shapes

### `cells` — the LLM-optimized stream

```jsonl
{"sheet":"Data","ref":"A1","t":"s","v":"name"}
{"sheet":"Data","ref":"B1","t":"s","v":"qty"}
{"sheet":"Data","ref":"A2","t":"s","v":"apple"}
{"sheet":"Data","ref":"B2","t":"n","v":3}
{"sheet":"Data","ref":"B3","t":"d","v":"2024-06-15T00:00:00"}
{"sheet":"Data","ref":"C2","t":"f","v":30,"formula":"A2+B2"}
```

Type tag `t`:
- `"s"` string, `"n"` number, `"i"` integer, `"b"` boolean, `"d"` date (auto via `Rows.parseDate`), `"f"` formula, `"e"` empty

Opt-in with `--with-styles`:
```jsonl
{"sheet":"Data","ref":"A1","t":"s","v":"name","style":{"bold":true,"fg":"FFFFFFFF","bg":"FF1F4E79"}}
```

Style fields collapsed to shorthand for token economy:
- `bold`, `italic` — booleans, emitted only when true
- `fg` / `bg` — ARGB hex strings, emitted only when set
- `nf` — number format code
- `border` — `{l, r, t, b}` each ARGB+style, emitted only when any side is non-none

### `meta` — workbook summary

```json
{
  "path": "data.xlsx",
  "sheets": [{"name": "Data", "rows": 161, "first_cell": "A1"}, {"name": "Other", "rows": 3}],
  "sst": {"count": 1144, "rich_count": 12},
  "has_styles": true,
  "has_theme": true,
  "has_comments": true,
  "format_version": "0.2.6"
}
```

Emits a single JSON object (not NDJSON) because there's exactly one record.

## CLI flag conventions

- `--sheet N` / `--name X` — mutually exclusive; 0-indexed int or name string.
- `--range A1:Z100` — bounding rectangle for `cells` / `rows`. A1-style.
- `--header` — on `rows`, treat row 0 as keys and emit `{…}` per data row.
- `--with-styles` — opt-in metadata; off by default to keep streams cheap.
- `--types` — on `cells`, include the `t` field (default: omit when value type is unambiguous).
- `--pretty` — human-readable indent. Off by default (NDJSON purity).
- `--seek N` — skip N records before emitting. For pagination.
- `--take N` — emit at most N records then exit.
- `--null-mode {omit|emit|string}` — how to represent empty cells. Default `omit` (no record for empty cells in `cells`; `null` in array for `rows`).

## Architecture

### Current reality

```
zlsx.cli.main()
  → xlsx.Book.open(path)  ← reads all parts into memory
  → Rows iterator            ← streaming within a sheet
  → writeRow per format
```

`Book.open` eagerly decompresses every sheet XML part. Fine for files up to ~10 MB. For 100 MB+ workbooks it's memory-unfriendly.

### Target: streaming book

```
zlsx.cli.main()
  → xlsx.Book.openLazy(path)   ← reads central dir only; parts loaded on demand
  → for each subcommand:
     → stream a single part at a time
     → emit NDJSON record-by-record
     → release buffers between records
```

Scope of the lazy-open refactor:
- Central directory scan: O(1) per entry, already lazy via `std.zip.Iterator`
- Per-sheet XML: load + decompress only when that sheet is requested
- SST: needs pre-load since cells reference by index — but SST size is bounded by the sst count, not sheet size
- Rows: already streaming; no change

Estimated LOC: ~300-500 for `openLazy` + sub-command plumbing. Big but not architectural rewrite.

### Sub-command dispatch

```
zlsx <subcommand> <file> [flags]
```

Parser becomes two-phase: first token picks the sub-command, then delegates. Keep the existing single-command path as `zlsx <file> --format …` for backward compat (treat as `zlsx rows <file> --format …`).

## Rollout plan (iter54+)

1. **iter54** — `cells` sub-command (most impactful for LLM streaming). NDJSON per-cell stream with opt-in `--with-styles`. Covers the 80 % use case.
2. **iter55** — `meta` + `list-sheets` as sub-commands. Trivial wrappers.
3. **iter56** — `comments` / `validations` / `hyperlinks` / `styles` / `sst` — each ~50 LOC of CLI glue over existing reader APIs.
4. **iter57** — `--range`, `--header`, `--seek`, `--take` flags on `rows` / `cells`.
5. **iter58** — `Book.openLazy` — per-sheet on-demand load. Enables larger workbooks.
6. **iter59+** — `--pretty`, `--null-mode`, `--format`.

Each iter stays under ~200 LOC, ships independently, user-observable per commit.

## Example pipelines

```bash
# All cells in one sheet with styles, as JSON for LLM.
zlsx cells data.xlsx --sheet 0 --with-styles | jq 'select(.t=="s")'

# Date columns only.
zlsx cells data.xlsx | jq 'select(.t=="d")'

# Sum a column from the CLI.
zlsx cells data.xlsx --range B2:B1000 | jq -r '.v' | awk '{s+=$1} END {print s}'

# Extract every comment across every sheet to feed an LLM.
zlsx comments data.xlsx | jq -r '[.ref, .author, .text] | @tsv'

# Grep for SST entries that look like email addresses.
zlsx sst data.xlsx | jq -r '.text' | rg '@\S+\.\S+'

# List every data validation with the range they cover, for a schema check.
zlsx validations data.xlsx | jq 'select(.kind=="whole")'
```

## Open design questions

1. **Dates**: do we emit ISO-8601 strings (`"2024-06-15T00:00:00"`) or Excel serials (`44927`)? **Proposal: ISO when the cell is date-styled, serial otherwise**. Caller can re-parse ISO cheaply; not every numeric is a date.
2. **Errors**: Excel cells can carry `#DIV/0!`-style error strings. **Proposal: emit as `{"t":"e","v":"#DIV/0!"}`**.
3. **Large SST**: 100k+ entries is realistic. Streaming `zlsx sst` is fine, but cell streams that reference SST need the whole SST in memory during emit. Acceptable — SSTs rarely exceed 10 MB.
4. **Rich text in cells**: opt-in via `--with-styles` (include the `runs` array when present). Off by default to keep the stream terse.
5. **Invariant**: every subcommand's output is a valid JSON-lines stream. `--pretty` breaks NDJSON — make it explicit that piped consumers should NOT use `--pretty`.

## Why not embed a query DSL?

1. `jq` already works. Every user has it. Zero lift.
2. We'd be re-implementing `select`, `map`, `@tsv`, etc. — a multi-year project for no gain.
3. The shape of our output is the leverage: structured NDJSON that `jq` can touch.
4. If we ever want a query DSL, it goes in a separate binary (`zlsxq`?) that sits between `zlsx` and `jq`.

---

## Summary for the impatient

> `zlsx rows` already exists and streams row arrays. Add `cells` next (per-cell NDJSON with type tags), then `meta` / `comments` / `validations` as thin wrappers over existing library APIs. Don't build a query DSL — tell users to pipe to `jq`.
