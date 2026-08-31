# `zlsx` CLI reference

The `zlsx` binary is "jq for Excel": the read/query sub-commands emit a
uniform NDJSON envelope by default, composing cleanly with `jq`, `rg`,
`awk`, `duckdb read_ndjson`, or an LLM ingest harness. Three deliberate
exceptions: `rows` offers CSV/TSV/legacy escape hatches via `--format`,
`meta` / `list-sheets` offer a single collapsed JSON object via
`--output pretty-json`, and the formula commands `eval` / `recalc` speak
their own versioned NDJSON grammars (below), not the row envelope.
Mutation sub-commands (the edit family, `scrub-metadata`, `embed --strip` /
`--vectors`) write the output workbook and stay silent on stdout on success —
except `embed --prune`, which emits one NDJSON summary. Design rationale
lives in [`jq-for-excel.md`](jq-for-excel.md) (historical; this page is the
current contract).

```bash
zlsx file.xlsx                          # default: rows sub-command
zlsx rows file.xlsx                     # explicit alias
zlsx cells file.xlsx --range B2:Z100    # per-cell NDJSON, bounded
zlsx meta file.xlsx --output pretty-json
```

**Contents**: [Sub-commands](#sub-commands) ·
[Row envelope](#the-ndjson-row-envelope) · [Flags](#flags) ·
[Example pipelines](#example-pipelines) ·
[Pipeline safety](#pipeline-safety) · [Exit codes](#exit-codes) ·
[License](#license)

---

## Sub-commands

### Read (NDJSON out)

| Command | `kind` | Per-line fields |
|---|---|---|
| `zlsx rows <file>` | `"row"` | `sheet, sheet_idx, row, cells[]` (each `{ref, col, t, v}`) |
| `zlsx cells <file>` | `"cell"` | `sheet, sheet_idx, ref, row, col, t, v, style?` |
| `zlsx comments <file>` | `"comment"` | `sheet, sheet_idx, ref, row, col, author, text, runs?` |
| `zlsx validations <file>` | `"validation"` | `sheet, sheet_idx, range, rule_type, op?, formula1, formula2?, values?` |
| `zlsx hyperlinks <file>` | `"hyperlink"` | `sheet, sheet_idx, range, url?, location?` |
| `zlsx pivots <file>` | `"pivot"`, then `"pivot_cache"` | `sheet, sheet_idx, name, part, location{}, rows[], cols[], pages[], values[], data_caption, grand_totals{}, style, cache{}` — [contract](#pivots--the-typed-pivot-read) |
| `zlsx merges <file>` | `"merge"` | `sheet, sheet_idx, range, start_row, start_col, end_row, end_col` — 1-based, corners inclusive |
| `zlsx defined-names <file>` | `"defined_name"` | `name, scope, sheet, sheet_idx, body, hidden` — [contract](#defined-names--the-workbook-name-inventory) |
| `zlsx styles <file>` | `"style"` | `idx, font, fill, border, num_fmt` (workbook-wide) |
| `zlsx sst <file>` | `"sst"` | `idx, text, runs?` (workbook-wide) |
| `zlsx meta <file>` | `"workbook"` + `"sheet"` | workbook record first, then per-sheet records |
| `zlsx list-sheets <file>` | `"sheet"` | `sheet, sheet_idx, state` — lighter-weight than `meta` |

#### `pivots` — the typed pivot read

`zlsx pivots <file>` walks the pivot graph (`xl/workbook.xml` `<pivotCaches>` →
cache definitions → records parts; every sheet's relationships → pivot-table
definitions) and emits one record per **pivot table**, in host-sheet order,
followed by one `pivot_cache` record per cache **no pivot table reads**. Orphan
caches belong to no sheet, so they ride along by default and under
`--all-sheets`, and a concrete selector — `--sheet`, `--name`, `--sheet-glob` —
suppresses them. Nothing is modified; this is the read half of
`goal_sigmoid.md` row S6, and the shape below is the S6 owner-gate contract.

```jsonl
{"kind":"pivot","sheet":"Report","sheet_idx":1,"name":"PivotTable1","part":"xl/pivotTables/pivotTable1.xml",
 "location":{"ref":"A3:B6","first_header_row":1,"first_data_row":1,"first_data_col":1},
 "rows":[{"field":"Region","idx":0}],"cols":[],"pages":[],
 "values":[{"name":"Sum of Qty","field":"Qty","idx":1,"subtotal":"sum","show_data_as":null,"num_fmt_id":null}],
 "data_caption":"Values","grand_totals":{"rows":true,"cols":true},"style":"PivotStyleLight16",
 "cache":{"id":7,"part":"xl/pivotCache/pivotCacheDefinition1.xml","records_part":"xl/pivotCache/pivotCacheRecords1.xml",
  "record_count":3,"refreshed_by":"zlsx","refreshed_date":"45000.5","refresh_on_load":false,"save_data":true,
  "source":{"type":"worksheet","sheet":"Data","ref":"A1:C4","name":null,
            "resolved":{"sheet":"Data","sheet_idx":0,"via":"sheet_attr","bounds":"A1:C4"},"unresolved":null},
  "fields":[{"name":"Region","num_fmt_id":0,"formula":null,"items":2,"types":["string"],"min":null,"max":null},
            {"name":"Qty","num_fmt_id":0,"formula":null,"items":null,"types":["number","integer"],"min":"3","max":"5"}]}}
{"kind":"pivot_cache","cache":{…}}
```

(One line per record on the wire; wrapped here for reading.)

| Field | Meaning |
|---|---|
| `sheet`, `sheet_idx` | The **host** sheet — where the pivot renders; `location.ref` is in its grid. Dropped under `--output compact-ndjson` in favour of the sheet prologue. |
| `name`, `part` | `pivotTableDefinition@name` (decoded) and the part's archive name. |
| `location` | `ref` (entity-decoded) plus `first_header_row` / `first_data_row` / `first_data_col`; an absent attribute is `null`. |
| `rows`, `cols`, `pages` | The axes, in order. Each entry is `{"field":NAME,"idx":N}` — `NAME` is the cache field at pivot-field ordinal `N` (`null` when the cache has no such field) — or `{"values":true}` for the values axis (`x="-2"`). |
| `values` | The data fields: caption (`name`), source field (`field` / `idx`), `subtotal` as the ST_DataConsolidateFunction token (`sum`, `count`, `average`, `max`, `min`, `product`, `countNums`, `stdDev`, `stdDevp`, `var`, `varp`; an unknown token passes through verbatim), `show_data_as` (raw; `null` means normal), `num_fmt_id`. |
| `data_caption`, `grand_totals`, `style` | The values-axis caption, `rowGrandTotals` / `colGrandTotals` (schema default `true`), `pivotTableStyleInfo@name`. |
| `cache` | The cache the pivot reads, or `null` when neither the part's relationship nor its `cacheId` reaches one. |
| `cache.id` | `<pivotCache cacheId>` from `xl/workbook.xml`; `null` for a cache only a pivot's own relationship reaches. |
| `cache.records_part`, `cache.record_count` | The `pivotCacheRecords` part and its declared `recordCount`. Records are never parsed. |
| `cache.refreshed_by`, `refreshed_date`, `refresh_on_load`, `save_data` | Refresh facts as written. `refreshed_date` is the `refreshedDate` serial, or the `refreshedDateIso` string when a producer wrote only that; `null` with neither. |
| `cache.source` | `{"type":"worksheet",…}`, `{"type":"consolidation","range_sets":[…]}`, `{"type":"external","connection_id":N}`, `{"type":"scenario"}` or `{"type":"unknown","raw":"…"}`. |
| `cache.source.sheet`, `.ref`, `.name` | The `worksheetSource` spellings, decoded (`sheet` and `name` by their string carrier, `ref` entity-decoded). A worksheet source is spelled either as `sheet` + `ref`, or as `name` — a table or a defined name. |
| `cache.source.resolved` | **What the spelling led to** — the surface-matrix footnote-¹⁰ answer: `{"sheet":…,"sheet_idx":N,"via":"sheet_attr"|"table"|"defined_name","bounds":AREA}` for a sheet of this workbook; `{"external":TARGET}` when the source lives in another workbook (the `r:id` relationship target); `null` when the spelling names nothing the workbook has (a dangling sheet name, a defined name with a dynamic or 3D body). Each consolidation `range_sets[]` entry carries the same five keys. |
| `cache.source.resolved.bounds` | The area the spelling **proves** it reads, as A1 text (S7b-1): a rectangle (`"A1:C4"` — a direct `ref`, the table's `ref`, a name body such as `Data!$A$1:$C$4`), whole columns (`"A:C"`, from a name body `Data!$A:$C`) or whole rows (`"1:4"`); `null` when the spelling names the sheet and nothing on it (a `sheet` attribute alone, or a `ref` the bounds parser rejects — anything but plain uppercase `A1`, `A1:C4`, `A:C`, `1:4`). Rectangles are corner-normalised. |
| `cache.source.unresolved` | Why a `"resolved":null` spelling could not be placed, and what it still proves (S7b-1): `{"why":WHY,"sheets":[{"sheet":NAME,"sheet_idx":N},…]}`, `null` when the spelling resolved. Present wherever `resolved` is — a worksheet source and each consolidation `range_sets[]` entry; the `external`, `scenario` and `unknown` source objects carry neither key. `why` is one of `dangling_sheet`, `dangling_name`, `unbounded_body` (a defined name whose body is not one static sheet-qualified area — dynamic, 3D, a union, a bare range, or reaching its sheet through another name), `unplaceable_rid` (an `r:id` that names no External-mode external-link relationship), `sheetless_ref` (`ref` with neither `sheet` nor `name`), `no_locator`. `sheets` lists, ascending, every sheet the spelling still names — the local `sheet` beside an unplaceable `r:id`; every sheet qualifier (each sheet of a 3D span), table host and referenced name's body reachable from a name body, recursively — or is empty when nothing local is proven. |
| `cache.fields` | The cache-field schema in ordinal order: decoded `name`, `num_fmt_id`, `formula` (calculated fields; `null` otherwise), `items` (`sharedItems@count`; `null` when absent), `types` (the `contains*` inventory as a list drawn from `string`, `number`, `integer`, `blank`, `date`, `mixed`; `null` without `<sharedItems>`), `min` / `max` (`minValue` / `maxValue`, else `minDate` / `maxDate`, raw). |

Sheet selection follows the read family — `--sheet` / `--name` narrow to one
host sheet, `--all-sheets` / `--sheet-glob` widen, the default streams every
sheet — and `--skip` / `--take` page the record stream. A workbook without
pivots is an empty, successful stream. The graph is read whole or not at all
(`MalformedPivotXml`, exit 2): a part a pivot, cache or sheet relationship
names but that is missing, mistyped or unreadable, a `<pivotCache>` entry
without its `cacheId` or `r:id`, a cache whose identity disagrees between
`xl/workbook.xml` and the pivot that reads it, two caches under one id, a pivot
with two cache edges, a records part that is named but absent, a defined-name
inventory the engine refuses or a `<tablePart>` that attaches nothing (checked,
in that order, when a source is spelled by name and no defined name has it) all fail the command rather than listing the pivots
that did parse — a partial inventory is the shape of a guard hole. Formats, conditional formats,
chart formats, hierarchies and every OLAP-only element are not exposed; the
parts stay byte-preserved for callers that need them raw
(`zlsx_pkg.PartStore`).

#### `defined-names` — the workbook name inventory

`zlsx defined-names <file>` (S3b) emits one record per `<definedName>` of
`xl/workbook.xml`, in document order — the order Excel's name manager
shows. Nothing is resolved or rewritten: `body` is the formula text as
authored (entity-decoded), whether it is a static area, a 3D span, a
dynamic `OFFSET`, or a constant.

```jsonl
{"kind":"defined_name","name":"Prices","scope":"workbook","sheet":null,"sheet_idx":null,"body":"Data!$A$1:$C$4","hidden":false}
{"kind":"defined_name","name":"_xlnm.Print_Area","scope":"sheet","sheet":"Report","sheet_idx":1,"body":"Report!$A$1:$B$9","hidden":false}
```

| Field | Meaning |
|---|---|
| `name` | The `name` attribute, decoded by its string carrier (entities + ST_Xstring). Excel's built-ins keep their `_xlnm.` prefix. |
| `scope`, `sheet`, `sheet_idx` | `"workbook"` for a name with no `localSheetId` (`sheet` / `sheet_idx` are `null`); `"sheet"` for a sheet-scoped name, with the scope sheet's decoded name and its zero-based index (`localSheetId` indexes the `<sheets>` order). A `localSheetId` past the sheet list keeps `scope:"sheet"` and the index as written, with `sheet:null` — the attribute is what the producer wrote. |
| `body` | The element's inner text — the formula the name refers to — entity-decoded, verbatim otherwise. |
| `hidden` | The `hidden` attribute (schema default `false`). Hidden names (filter databases, slicer caches) are streamed, not suppressed — filter with `jq 'select(.hidden|not)'`. |

Selection mirrors the `pivots` orphan-cache rule: a concrete selector —
`--sheet`, `--name`, `--sheet-glob` — narrows to the names **scoped** to
matching sheets and suppresses workbook-scope names; the default and
`--all-sheets` stream every name, workbook scope included. `--skip` /
`--take` page the stream; `--output compact-ndjson` is accepted as a
no-op (a defined name's sheet fields are its scope, not a host
location, so there is no prologue to hoist them into). The command
routes through the package layer like `pivots`, so an archive the
lenient reader tolerates but the package layer refuses (ZIP data
descriptors, for one) is exit 2; a workbook without defined names is an
empty, successful stream. A name or body the read cannot serve
faithfully refuses the whole command (exit 2) rather than emit a
record that lies: a carrier that does not decode (a bad entity, an
ill-formed `_xHHHH_` escape), a decoded name, body or sheet name that
is not UTF-8 (NDJSON must stay parseable), or a body carrying embedded
markup (a CDATA section, a comment — element text no consumer of the
name reads through; the rewriters refuse such a body too) — a partial
or wrong name inventory is the shape of a guard hole. Two shapes are
not part of the workbook.xml view at all and emit no record, on any
surface that reads it: a bodiless self-closing `<definedName/>` and an
entry with no `name` attribute — spellings Excel never writes, and
names the engine's symbol table (which `pivots` resolution and the
edit-time name sweep read through) equally does not hold, so the CLI
cannot show a name the rest of the toolchain would not honour.

#### `merges` — merged ranges

`merges` (S3b) is the same read family as `validations` / `hyperlinks`:
one record per `<mergeCell>`, every sheet by default, `--sheet` /
`--name` / `--all-sheets` / `--sheet-glob` narrowing and widening,
`--skip` / `--take` paging, a sheet prologue under
`--output compact-ndjson`. `range` is the A1 rectangle as one merged
range declares it; `start_row` / `start_col` / `end_row` / `end_col`
are its inclusive corners, 1-based on both axes (column A = 1, like the
`cells` envelope), so `jq` consumers can intersect without re-parsing
A1 text.

### Edit (load-modify-save)

Every edit sub-command reads `<file>`, applies the change, and atomic-renames
the result to `--out PATH` — the input is never modified in place. A refused
edit (a construct the rewriter cannot shift safely yet — see
[`plans/refusal-audit.md`](plans/refusal-audit.md)) exits with a diagnostic
instead of writing a corrupt workbook; a successful save is byte-safe.

Pivot tables on the edited sheet (S7a, `goal_sigmoid.md`): a row/col edit on
a sheet that only *hosts* a pivot moves the pivot's output rectangle
(`pivotTableDefinition/location@ref`) in step — an insert at or above the
rectangle, a delete above it, an edit below or to its right (which leaves the
part byte-identical); the host's own `<pivotSelection>` coordinates move with
the grid. The edit exits 3 instead when it lands *inside* the pivot's
footprint (the rectangle, or the report-filter band Excel draws above it —
Excel refuses that edit too), when some cache *may read* the host sheet — a
source resolved to it (row S7b), or one `zlsx pivots` reports as
`"resolved":null` (a dynamic defined name, a dangling spelling) or of a type
it does not know — when one pivot part is hosted by two sheets, or when the
pivot graph cannot be read whole. One
known hole in the promise, pinned by the S6 audit: a sheet a pivot only *reads
from* is admitted today, and for a source spelled as `sheet` + `ref` the
cache's range goes stale (a table-named source follows the table rewriter and
stays valid). `zlsx pivots` names the sheet every pivot reads from
(`cache.source.resolved`); row S7b closes the hole.

| Command | Required flags | What it does |
|---|---|---|
| `append-rows` | `--sheet N --out P` | Append rows read from stdin — one JSON array per line (`null`→empty, `true/false`→bool, integer→int, number→float, string→str) |
| `set-cell` | `--sheet N --ref A1 --value <JSON> --out P` | Rewrite one cell in place; `--value` is a single JSON token with the same type mapping |
| `insert-row` / `delete-row` | `--sheet N --row N --out P` | Structural row edit (1-based row) |
| `insert-column` / `delete-column` | `--sheet N --col LETTER --out P` | Structural column edit (`A`..`XFD`) |
| `add-sheet` | `--new-name NAME --out P` | Append a new empty sheet |
| `rename-sheet` | `--sheet N --new-name NAME --out P` | Rename a sheet (Unicode-aware duplicate check) |
| `delete-sheet` | `--sheet N --out P` | Drop a sheet (cannot drop the last one) |
| `rename-table-column` | `--table NAME --old-name OLD --new-name NEW --out P` | Rename a column of a named table; the `<tableColumn>`, the table's formulas, every structured reference workbook-wide (`Table1[Old]`, bare `[Old]` / `[@Old]` inside the table), defined names, hyperlinks, DV / CF and the header cell follow. Exit 3 on `TableNotFound` / `TableColumnNotFound` / `TableColumnNameInUse` / `InvalidTableColumnName` |
| `scrub-metadata` | `--out P` | Strip `docProps` identity metadata (author, last-modified-by, company, …) — elements are removed outright, not blanked; cell data is byte-preserved |

```bash
cat rows.ndjson | zlsx append-rows in.xlsx --sheet 0 --out out.xlsx
zlsx set-cell in.xlsx --sheet 0 --ref C5 --value '"hello"' --out out.xlsx
zlsx scrub-metadata in.xlsx --out clean.xlsx
```

### Embeddings (`embed`)

Manage embedding vectors stored *inside* the workbook (zlsx's embedding parts —
`vnd.fabre.zlsx.*` OPC parts, invisible to Excel). The model is invoked out of
band — the embed pipeline never makes a network call — so extract and write
compose through a pipe:

```bash
zlsx embed book.xlsx --extract --column A --coverage A2:A100 > rows.ndjson
my-embedder < rows.ndjson > vecs.ndjson
zlsx embed book.xlsx --vectors vecs.ndjson --model M \
    --column A --coverage A2:A100 --out out.xlsx
```

| Mode | Required flags | What it does |
|---|---|---|
| `embed --extract` | `--column A --coverage A2:A100` | Read-only: one `{"kind":"embed_row","row":N,"text":"…"}` record per embeddable row |
| `embed --vectors P` | `--column A --coverage A2:A100 --model NAME --out P` | Write the embedding parts from `{"row":N,"vector":[…]}` NDJSON; covered rows with no vector become tombstones. Optional: `--id NAME`, `--dtype f32\|int8-sym`, `--sheet N` |
| `embed --prune` | `--out P` | The redaction sweep: tombstone every slot whose row is no longer embeddable and zero its vector (a row deleted in plain Excel leaves its vector on disk — this removes it). Changed-but-present content is reported stale, not redacted |
| `embed --strip` | `--out P` | Remove the embedding parts *and* the recovery record — the pre-share operation; the result reports `absent`, not `stripped` |

### Databricks (`dbx`)

The one network-touching family. Auth from `DATABRICKS_HOST` /
`DATABRICKS_TOKEN`; Genie space from `GENIE_SPACE_ID` (or `--space ID`);
SQL warehouse from `DATABRICKS_WAREHOUSE_ID` (or `--warehouse ID`).

| Command | What it does |
|---|---|
| `zlsx dbx push local.xlsx /Volumes/cat/schema/vol/f.xlsx [--overwrite]` | Upload a workbook to a Unity Catalog Volume — the file is parsed *before* upload, so garbage is refused client-side (exit 3) |
| `zlsx dbx pull /Volumes/... local.xlsx` | Download a workbook — parsed before the atomic rename, so a truncated or non-workbook body never lands |
| `zlsx dbx genie "question" [--space ID] [--timeout-secs N]` | Ask a Genie space; streams status / generated SQL / columns / rows / text as NDJSON |
| `zlsx dbx audit /Volumes/cat/schema/vol/zone/` | Hash every workbook in a landing zone and report drift / orphans / disappearances as NDJSON (exit 3 when there are findings) |

Workbooks and response bodies are capped at **64 MiB** in both directions:
`push` refuses a larger local input before parsing or upload, and `pull` /
`genie` reject a larger response after it is buffered (a correctness cap,
not a hard memory bound). `genie` polls up to `--timeout-secs`
(default **120**).

#### `dbx audit` — does the zone match what was ingested?

Walks the Volume directory (recursively, to 8 levels), downloads each
`.xlsx` / `.xlsm`, and records its **SHA-256**. Identity is the content
hash, not the `(mtime, size)` fingerprint the streaming source keys its
offsets on — that is the only way to catch a rewrite that preserved size
and moved mtime backwards.

| Flag | Effect |
|---|---|
| `--write-manifest <file>` | Record this run's hashes |
| `--manifest <file>` | Compare against a prior run → `drift` / `new` / `missing` |
| `--table <cat.sch.tbl>` | Compare against a table's ingestion record → `orphan` / `missing` |
| `--source-column <name>` | Provenance column (default `_source_file`) |
| `--warehouse <id>`, `--timeout-secs <n>` | SQL warehouse and poll budget (default **120**) |

| Status | Meaning |
|---|---|
| `ok` | Hash matches the manifest; the ingestion record has seen it |
| `new` | Present and readable, but the manifest has no entry — not a finding |
| `drift` | Content hash changed — the immutable-files convention is broken |
| `orphan` | In the zone, absent from the ingestion record: data the table never reflected |
| `missing` | In the manifest or the record, gone from the zone |
| `unreadable` | Bytes are there but do not parse as a workbook |

```bash
zlsx dbx audit /Volumes/main/default/landing/ --write-manifest audit.json
# …later, in CI:
zlsx dbx audit /Volumes/main/default/landing/ \
    --manifest audit.json --table main.default.sales   # exit 3 → fail the job
```

`--table` splices identifiers into a `SELECT DISTINCT`, so both the table
and the column must be plain `[A-Za-z0-9_]` identifiers (dots separating
the parts); anything else is refused before a request is made. If the
warehouse reports the result set as truncated, orphan findings are
flagged as untrustworthy on stderr rather than silently under-reported.

### Formula (`eval` / `recalc`)

The formula engine on the command line (M6, §12.2 of the design doc).
Both commands have their own argument grammar, their own NDJSON stream
contract (below — **not** the row envelope), and their own exit table.

```
zlsx eval   <file.xlsx> --formula "<formula>" (--sheet N | --name NAME) [--anchor A1]
            [--dialect da|legacy] [--now ISO8601] [--utc-offset MIN] [--seed N]
            [--mode excel|ieee] [--profile windows_1252] [--deadline SECONDS]
zlsx recalc <file.xlsx> --out <out.xlsx> [--now ISO8601] [--utc-offset MIN]
            [--seed N] [--mode excel|ieee] [--profile windows_1252]
            [--on-unsupported refuse|keep-stale-and-mark] [--report]
            [--deadline SECONDS]
```

- `--formula` carries the text; the next token is taken verbatim, so
  `-A1`, `--A1`, and other flag-shaped text need no escaping. A leading
  `=` is accepted and ignored. `--sheet N | --name NAME` is mandatory
  and exclusive for `eval`.
- `--anchor A1` supplies the evaluation site for site-dependent
  constructs (`ROW()`, `COLUMN()`, `@`); absent when one is needed →
  exit 3 (`FormulaAnchorRequired`).
- `--now` / `--seed` default to the system wall clock / secure random
  source; a source that cannot be read is **exit 6** (never conflated
  with OOM). `--seed` is a decimal u64; in every record it serializes
  as a **decimal string** (u64 exceeds the JSON/JS safe-integer range).
- `recalc --out` must not be the input — refused (exit 1) before
  anything is opened. The write is one atomic transaction (§5.7.9): on
  every refusal or cancellation path the destination's prior bytes are
  untouched ("no file" only when it was absent).

#### The stream state machine

Versioned NDJSON: every record carries `"kind"` and `"v":1`. The
grammars are normative — refusal and cancellation can occur **before**
any header exists:

```
eval stream   := refusal | cancelled
               | eval-header ( eval-cell* ) diagnostic* ( eval-complete | refusal | cancelled )
recalc stream := diagnostic* ( recalc-report | refusal | cancelled )    (stdout only with --report)
```

```jsonl
{"kind":"eval-header","v":1,"type":"number","value":3,"resolved":{"now":"2026-08-05T12:00:00.000Z","utcOffsetMin":0,"seed":"7","mode":"excel","profile":"windows_1252","dialect":"da","anchor":"B7"}}
{"kind":"eval-header","v":1,"type":"matrix","rows":2,"cols":2,"resolved":{…}}
{"kind":"eval-cell","v":1,"r":1,"c":1,"type":"number","value":1}
{"kind":"diagnostic","v":1,"severity":"warning","message":"…"}
{"kind":"eval-complete","v":1,"cells":4}
{"kind":"refusal","v":1,"error":"FormulaCycle","cells":["Sheet1!A1"],"truncated":false}
{"kind":"cancelled","v":1,"after":"eval-cell"}
{"kind":"recalc-report","v":1,"sheets":1,"cells":1,"passes":1,"nonConverged":0,"dynamicPasses":1,"keptStale":false,"calcChainRemoved":false,"census":[{"error":"FormulaUnsupportedFunction","cell":"Sheet1!B1"}],"censusTruncated":false,"resolved":{…}}
```

- Scalar results ride in the header (`type` ∈ `number` | `text` | `bool`
  | `error`) and emit no `eval-cell` records — `eval-complete` then
  counts 0. Matrix results put the shape in the header and stream cells
  row-major. **`blank` never appears**: an empty cell publishes as `0`
  (§5.3a's one mandatory conversion), in both commands.
- Terminal states are explicit: success = `eval-complete` /
  `recalc-report`; typed refusal = `refusal` (the `error` field is §10's
  plane-2 name, verbatim); cancellation = `cancelled`, whose `after`
  names the last record emitted (`none` when the stream was still
  empty). Completed records stand — the stream is prefix-valid.
- **SIGPIPE exception**: a closed pipe cannot receive a terminal record.
  SIGPIPE ends the stream prefix-valid with **no terminal and exit 0**
  (`zlsx eval … | head -1` is not an error). Any other EOF without a
  terminal record exits nonzero — the exit code is what distinguishes
  the two.
- `resolved` echoes every input the run actually used (§5.7.8), so a
  stream is reproducible from its own header. `dialect` / `anchor`
  appear on `eval` only; a recalc derives dialect per stored cell.

#### Exit codes (`eval` / `recalc`)

| Code | `eval` | `recalc` |
|---|---|---|
| 0 | evaluated (Excel error values like `#DIV/0!` are results) | recalced + written |
| 1 | usage (incl. missing `--sheet`/`--name`) | usage (incl. `--out` = input) |
| 2 | open/parse (incl. sheet not found) | open/parse |
| 3 | typed refusal — **incl. `FormulaLimitExceeded`** (limits are refusals at every layer) — or `--deadline` | refusal / cancellation — **destination untouched** |
| 4 | allocation failure (genuine OOM only) | allocation failure |
| 5 | stdout write failure | output write/rename failure |
| 6 | default-context acquisition failure: the wall clock or random source could not be read while resolving an omitted `--now`/`--seed` (never reported as OOM) | same |
| 130/143 | SIGINT / SIGTERM | same — **before the rename only** |

**Exit mapping is commit-aware** (`recalc`): the rename is the commit,
and a signal that arrives after it reports **0**, never 130/143 — the
destination already holds the recalced bytes, and telling a scripted
caller to retry finished work would be a lie. The cancellation
guarantee covers workbook-file mutation only; the final cancellation
poll sits immediately before the rename, and the rename-through-swap
span is non-cancellable.

---

## The NDJSON row envelope

```jsonl
{"kind":"row","sheet":"Data","sheet_idx":0,"row":1,"cells":[{"ref":"A1","col":1,"t":"str","v":"name"},{"ref":"B1","col":2,"t":"str","v":"qty"}]}
{"kind":"row","sheet":"Data","sheet_idx":0,"row":2,"cells":[{"ref":"A2","col":1,"t":"str","v":"apple"},{"ref":"B2","col":2,"t":"int","v":3}]}
```

`t` ∈ `"str"` | `"int"` | `"num"` | `"bool"` | `"blank"` | `"date"` | `"error"`
| `"formula"`; empty cells are skipped from `cells[]` unless `--include-blanks`
is set.

- **Date cells** carry an extra `serial` field with the raw Excel serial
  alongside the ISO `v`. Both the 1900 and 1904 (`<workbookPr date1904="1"/>`)
  epochs decode correctly — `serial` always carries the source value unshifted.
- **Formula cells** carry `formula` (own text) or `formula_ref`
  (shared-formula base ref) plus an optional `cached` value instead of `v`.
- **Error cells** carry the literal Excel error string in `v` (`"#DIV/0!"`,
  `"#N/A"`, …).

---

## Flags

**Default sheet scope** differs by family: `rows` / `cells` read sheet 0
unless told otherwise; `comments` / `validations` / `hyperlinks` / `pivots` /
`merges` stream every sheet; `styles` / `sst` / `meta` / `list-sheets` are
workbook-wide, and so is `defined-names` (a sheet selector there narrows by
a name's *scope* — see its contract above).

**Sheet selection** — mutually exclusive:

```bash
--sheet N                     # 0-indexed
--name "Summary"              # by name
--all-sheets                  # every sheet concatenated
--sheet-glob 'Data*'          # glob on sheet name (UTF-8 `?` matches one codepoint)
```

**Row / cell bounds** (rows / cells / comments):

```bash
--start-row 2 --end-row 100   # 1-based inclusive, per sheet
--range B2:Z100               # A1 bounding rectangle (rows + cells only)
--header                      # rows only: promote first row to keys, emit `fields` dict
```

**Stream pagination** — applies globally across the concatenated output:

```bash
--skip N --take M
```

**Cell metadata opt-ins** (cells / rows):

```bash
--include-blanks              # emit t:"blank" records for empty cells
--with-styles                 # attach terse style: {bold?, italic?, fg?, bg?, nf?, border?}
--sst-lazy                    # open with Book.openSstLazy — defer SST decode until first
                              # cell access (huge-workbook RAM mitigation; sparse access
                              # wins, full sweeps cost a bit more). Trade-off: malformed
                              # <t> in the SST surfaces on first access, not at open.
```

**Output modes**:

```bash
--output ndjson               # default: invariant-envelope stream
--output compact-ndjson       # sheet-prologue variant (drops sheet/sheet_idx on data
                              # records); applies to rows / cells / comments /
                              # validations / hyperlinks / pivots / merges — a no-op
                              # on workbook-scoped sub-commands (defined-names included)
--output pretty-json          # meta + list-sheets only: single collapsed JSON object
                              # (rejected on the streaming sub-commands)
```

**Row format (rows sub-command only)** — legacy escape hatches:

```bash
--format jsonl                # default: envelope
--format legacy-jsonl         # pre-iter55a bare arrays
--format legacy-jsonl-dict    # pre-iter55a bare objects (old `--format jsonl-dict` is a
                              # deprecated alias)
--format tsv                  # tab-separated, \N for empty
--format csv                  # RFC 4180-style quoting (LF record separators)
```

---

## Example pipelines

```bash
# All string cells across every sheet.
zlsx cells data.xlsx --all-sheets | jq 'select(.t=="str") | {sheet, ref, v}'

# Sum a column from the CLI without loading everything.
zlsx cells data.xlsx --range B2:B1000 | jq -r 'select(.t=="int" or .t=="num") | .v' | awk '{s+=$1} END {print s}'

# Every comment across every sheet, as TSV.
zlsx comments data.xlsx --all-sheets | jq -r '[.sheet, .ref, .author, .text] | @tsv'

# Schema check: every list-type validation + its range.
zlsx validations data.xlsx | jq 'select(.rule_type=="list") | {sheet, range, values}'

# Grep SST for emails.
zlsx sst data.xlsx | jq -r '.text' | rg '@\S+\.\S+'

# Merged header bands: every merge that touches row 1.
zlsx merges data.xlsx --all-sheets | jq 'select(.start_row==1)'

# Visible workbook-scope names with their bodies, as TSV.
zlsx defined-names data.xlsx | jq -r 'select(.scope=="workbook" and (.hidden|not)) | [.name, .body] | @tsv'

# Push a report to a UC Volume, then ask Genie about it.
zlsx dbx push report.xlsx /Volumes/main/default/landing/report.xlsx
zlsx dbx genie "what were total units last month?"
```

---

## Pipeline safety

On POSIX platforms, `zlsx cells huge.xlsx | head -10` exits 0 cleanly (no
broken-pipe stderr noise), `SIGINT` → exit 130, and `SIGTERM` → exit 143 —
both flushing in-flight records. The Windows build does not install these
signal handlers.

Malformed input handling is part-dependent — some parsers propagate
`error.MalformedXml` and abort `Book.open` with exit 2, others are lenient and
degrade to empty / partial metadata. In practice:

- If `Book.open` fails, the CLI exits 2 with a stderr diagnostic and no stdout
  stream.
- If `Book.open` succeeds, the CLI emits as much as it can. Malformation
  surfacing from `Rows.next()` during per-cell iteration is caught per-sheet
  and emitted as an inline
  `{"kind":"error","scope":"sheet","code":"MalformedXml",…}` record —
  processing continues with neighbour sheets. Filter with
  `jq 'select(.kind!="error")'` for the data-only stream.

Scripts that need a precise part-by-part failure map should read the reader
source (`src/xlsx.zig`) — the design-doc "Operational guarantees" section
tracks the intent, but current behaviour is what `parseSharedStrings` /
`parseStyles` / `parseTheme` / `parseMergedRangesForSheet` /
`parseHyperlinksForSheet` / `parseDataValidationsForSheet` /
`parseCommentsForSheet` actually do today.

---

## Exit codes

The table below is the contract for the **read family**. The formula
family (`eval` / `recalc`) has its own, richer table — see
[Formula](#formula-eval--recalc) above. The edit, embed,
and dbx families reuse the same codes with per-command meanings — e.g.
`dbx push` returns 3 when the upload preflight refuses a non-workbook, edit
sub-commands return 3 on a refused structural edit, and the embed family
returns 4 on a vector-buffer allocation failure. Exit 4 for a breached
decompression limit is shared by every family that opens through the core
reader or the editor (read, edit, embed); `eval` / `recalc` report it as 2.
Scripts branching on exact codes for non-read sub-commands should test the
specific command they use.

| Code | Meaning (read family) |
|---|---|
| 0 | Success (inline `error` records may still have been emitted for recoverable sheet-level MalformedXml) |
| 1 | Bad CLI arguments |
| 2 | Could not open the input: missing file, permission denied, not a valid xlsx archive, malformed parts at open time — or, on `pivots`, a pivot graph that cannot be read whole (a named part missing or unreadable, a cache identity that disagrees) |
| 3 | Sheet not found (by name / index). A `--sheet-glob` matching zero sheets is an empty *successful* stream (exit 0), not an error |
| 4 | A decompression limit was breached (`ZipBombSuspected`): a part declared past the per-part cap, past the ratio cap, or a whole archive declared past the aggregate budget — checked on the central directory before anything is inflated, so no partial output precedes it. Numbers in [Pipeline safety](#pipeline-safety). The embed family also returns 4 on a vector-buffer allocation failure |
| 5 | OS error writing output (stdout write failure, disk full, mutation-save I/O) |
| 130 | SIGINT (POSIX) |
| 143 | SIGTERM (POSIX) |

---

## License

zlsx is proprietary — see the repository [LICENSE](../LICENSE). The 60-day
free evaluation covers the released binaries and wheels only; continued or
non-evaluation use requires a commercial license
(**laurent.fabre@gmail.com**).
