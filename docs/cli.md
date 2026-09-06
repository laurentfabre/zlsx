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
| `zlsx doc-props <file>` | `"doc_props"` | one record: `creator, last_modified_by, title, subject, description, keywords, category, created, modified, revision, company, manager, application, hyperlink_base, has_custom_properties` — [contract](#doc-props--the-document-properties-report) |
| `zlsx anchors <file>` | `"image_anchor"`, `"chart_anchor"` | `sheet, sheet_idx, part, anchor, from{}, to{}, absolute{}`, then `bytes` (images) or `chart_type, series_refs[]` (charts) — [contract](#anchors--anchored-images-and-charts) |
| `zlsx conditional-formats <file>` | `"conditional_format"` | `sheet, sheet_idx, sqref, rule_type, formulas[], dxf_id, priority` — one record per `<cfRule>` — [contract](#conditional-formats--the-rule-inventory) |
| `zlsx sheet-props <file>` | `"sheet_props"` | `sheet, sheet_idx, dimension, pane{x_split, y_split, top_left_cell, active_pane, state}` — one record per sheet — [contract](#sheet-props--panes-and-the-extent) |
| `zlsx calc-props <file>` | `"calc_props"` | one record: `calc_id, full_calc_on_load, iterate, iterate_count, iterate_delta` — [contract](#calc-props--the-calculation-properties) |
| `zlsx styles <file>` | `"style"` | `idx, font, fill, border, num_fmt` (workbook-wide) |
| `zlsx sst <file>` | `"sst"` | `idx, text, runs?` (workbook-wide) |
| `zlsx meta <file>` | `"workbook"` + `"sheet"` | workbook record first, then per-sheet records |
| `zlsx list-sheets <file>` | `"sheet"` | `sheet, sheet_idx, state` — lighter-weight than `meta` — [contract](#list-sheets--names-and-visibility) |

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
A1 text. Sheet names — in the `sheet` field and against `--name` /
`--sheet-glob` — carry the reader's spelling, exactly as on every
Book-route command (`rows`, `cells`, `comments`, `validations`,
`hyperlinks`): XML entities decoded, ST_Xstring `_xHHHH_` escapes as
written. The package-route commands (`pivots`, `defined-names`,
`anchors`, `conditional-formats`, `sheet-props`) decode ST_Xstring too, so the rare
sheet name spelled with such an escape reads differently across the
two families — a reader-level lift on a later S3b slice, not a
per-command patch (one family, one spelling).
One validity floor is this command's own: the sheet name is a merge
record's only user-text channel, so a selected sheet that has merges
under a non-UTF-8 name refuses the whole command (exit 2) before any
record — the stream is valid NDJSON or it is nothing. A bad-named
sheet with no merges, or one the selection excludes, emits nothing
and does not refuse.

#### `doc-props` — the document-properties report

`zlsx doc-props <file>` (S3b) emits the document-properties field set as
**one** NDJSON record: every `docProps/core.xml` / `docProps/app.xml`
field the typed view models, plus a presence flag for
`docProps/custom.xml` (whose arbitrary key/value pairs zlsx deliberately
does not model). This is the read `meta` embeds as its `doc_props`
object and `Editor.doc_props` returns over the C ABI, promoted to a
typed report of its own — same field set, same order, same raw values,
with two boundary-convention divergences: a present-but-empty element
(`<dc:creator></dc:creator>`) is `""` here and in `meta`, while the C
ABI's length-zero convention makes Python's `Editor.doc_props` report
it as `None`; and a malformed-UTF-8 value refuses here (below) where
Python decodes it with replacement characters.

```jsonl
{"kind":"doc_props","creator":"Alex","last_modified_by":"Alex","title":null,"subject":null,"description":null,"keywords":null,"category":null,"created":"2014-03-07T16:08:25Z","modified":"2017-08-27T03:03:25Z","revision":null,"company":null,"manager":null,"application":"Microsoft Excel","hyperlink_base":null,"has_custom_properties":false}
```

Absent fields are `null` — never omitted, so consumers can index without
existence checks — and a workbook with no docProps parts at all is a
record of nulls under exit 0: the absence itself is the one fact the
line reports. Values are the element text **as stored** (XML entities as
written, timestamps verbatim W3CDTF), not a re-decoding; a field value
that is not UTF-8 refuses the whole command (exit 2) with nothing
written, the family's validity floor — every read and validation
failure lands before the first byte (only a stdout I/O failure mid-line,
exit 5, can truncate the record, as on any streaming command). The
command routes through the package layer like `pivots`, so an archive
the lenient reader tolerates but the package layer refuses is exit 2
(where `meta`, a Book-route command, degrades to `"doc_props": null`
instead — its own contract). There is no sheet dimension: sheet
selectors and `--format` are tolerated and ignored (the `meta` family's
wrapper-friendliness), and `--skip` / `--take` do not apply to a
single-record report. The identifying subset of this report is what
`scrub-metadata` removes by default — `application`, `created`,
`modified` and `revision` are deliberately preserved by its default
mask (its contract, not this one).

#### `anchors` — anchored images and charts

`zlsx anchors <file>` (S3b) emits one record per anchored image
(`"image_anchor"`) and per anchored chart (`"chart_anchor"`): sheets in
workbook order, a sheet's images before its charts, each class in
drawing-document order. The view is `zlsx_pkg.imageAnchors` +
`zlsx_pkg.chartAnchors` attributed to workbook sheets — the walkers key
anchors by worksheet *part*, the command resolves each part back to the
`<sheets>` entry that owns it (through the same strict sheet inventory
`conditional-formats` reads, below — not the lenient projection) so
every record carries the read family's `sheet` / `sheet_idx` envelope.
The payloads stay in the archive: an image record reports the part's
decompressed size (`bytes`), a chart record its detected type and
formula refs — `zlsx-extract-images` or `zlsx_pkg.PartStore` fetch the
bytes by `part`. The C ABI (`zlsx_editor_anchors_ndjson`) and py-zlsx
(`Editor.anchors()` / `zlsx.anchors(path)`) hand over these exact
bytes with no selector.

```jsonl
{"kind":"image_anchor","sheet":"Report","sheet_idx":1,"part":"xl/media/image1.png","anchor":"two_cell","from":{"row":3,"col":2,"row_off":0,"col_off":9525},"to":{"row":8,"col":5,"row_off":19050,"col_off":0},"absolute":null,"bytes":4096}
{"kind":"chart_anchor","sheet":"Report","sheet_idx":1,"part":"xl/charts/chart1.xml","anchor":"one_cell","from":{"row":2,"col":6,"row_off":0,"col_off":0},"to":null,"absolute":null,"chart_type":"bar","series_refs":["Data!$B$1","Data!$A$2:$A$4","Data!$B$2:$B$4"]}
```

| Field | Meaning |
|---|---|
| `sheet`, `sheet_idx` | The sheet whose drawing carries the anchor. Dropped under `--output compact-ndjson` in favour of the sheet prologue. |
| `part` | The image or chart part's archive name — the key for fetching the payload. |
| `anchor` | `"two_cell"` (`<xdr:twoCellAnchor>`), `"one_cell"` (`<xdr:oneCellAnchor>` — sized by `<xdr:ext>`, not modelled here) or `"absolute"` (`<xdr:absoluteAnchor>` — pixel-pinned, no cell anchor). `xdr:` is the canonical spelling, not a requirement: the walk follows every prefix bound to the spreadsheetDrawing namespace, the default namespace included (openpyxl's `<wsDr xmlns="…"><oneCellAnchor>`); a binding under a name it cannot spell (past its 100-byte prefix limit or its eight-alternate replay cap) refuses the inventory (`MalformedDrawingXml`, exit 2), as does a drawing or chart part carrying a `<!DOCTYPE` (OPC forbids a DTD; an entity could rewrite the markup) or a `<` inside an attribute value (not well-formed XML; a close or a corner spelled there would be taken for the real one) — both new with the namespace-aware drawing slice: such a workbook used to list the anchors under its other names, and a DTD's entity values could surface as anchors. |
| `from`, `to` | Cell-grid corners: `row` / `col` 1-based on the wire (like the `cells` / `merges` envelopes; the drawing XML itself is 0-based), `row_off` / `col_off` the EMU offsets within that cell, verbatim (1 EMU = 1/914400 inch). `to` is `null` for `one_cell`; both are `null` for `absolute`. |
| `absolute` | `{"x","y","cx","cy"}` in EMUs, verbatim, for `"absolute"` anchors; `null` otherwise. |
| `bytes` | Image records only: the image part's decompressed size in bytes. The payload itself is never on the wire. |
| `chart_type` | Chart records only: the detected chart-XML element — `bar`, `line`, `pie`, `scatter`, `area`, `bubble`, `radar`, or `other` for unrecognised / compound charts. |
| `series_refs` | Chart records only: every `<c:f>` formula ref of the chart part, entity-decoded, flattened in document order (series names, categories, values); an empty carrier (`<c:f></c:f>`, `<c:f/>`) is an empty string; the walk is element-agnostic, so a linked chart title's `<c:f>` rides in the same list. Empty for a chart of inline literal data. The refs follow structural edits (the chart `<c:f>` sweep), so after a rename or a row / column edit on the sheet they name the record spells the respelled part. |

Sheet selection follows the read family — `--sheet` / `--name` narrow to
one sheet, `--all-sheets` / `--sheet-glob` widen, the default streams
every sheet — and `--skip` / `--take` page the merged record stream. A
workbook without drawings is an empty, successful stream, and so is a
drawing of shapes only (shapes are not modelled). The command routes
through the package layer like `pivots`, so an archive the lenient
reader tolerates but the package layer refuses is exit 2. A read the
command cannot serve faithfully refuses whole (exit 2) rather than emit
a record that lies or thin the inventory: an anchor on a worksheet part
`xl/workbook.xml` does not list (the record could not carry a truthful
`sheet`); a sheet list the strict workbook read cannot prove — the
`conditional-formats` contract below: a listed sheet the read cannot
place (its relationship dangling, mistyped or external, its part
absent from the archive or unreadable), two `<sheet>` entries reaching
one part (the same anchor would ride out twice under two identities),
two names decoding to one spelling, an empty `<sheets>` list (a
sheetless workbook is refused, never served as an empty inventory);
a drawing structure the walkers
recognise but cannot read whole — a sheet's drawing relationship, a
pic's image relationship or a frame's chart relationship that dangles,
is of the wrong relationship type, or names an absent part, a
`<drawing>` element whose reference cannot be read, a second live
`<drawing>` element on one sheet (the schema allows one), an anchor whose
`from` / `to` / `pos` / `ext` does not parse (a `two_cell` anchor with
an unreadable `<to>` must not ride out looking like `one_cell`, and a
`one_cell` anchor's schema-required extent is validated even though it
stays off the wire), an anchor block or a `<c:f>` carrier that never
closes; a sheet name or series ref whose carrier does not decode or is
not UTF-8, a series ref with embedded markup, a part name that is not
UTF-8 — a partial anchor inventory is the shape of a guard hole. (The
Zig walkers `zlsx_pkg.imageAnchors` / `zlsx_pkg.chartAnchors` keep
their historical lenient contract — skip what does not parse, no
crash; the CLI reads their strict variant.) Scope notes
from the underlying walkers: namespace prefixes are resolved from the
drawing document itself (alternate bindings to the drawing URIs are
followed — `pkg/drawings.zig` documents the exact scope); legacy VML
drawings (`<legacyDrawing>`) are not walked; and the walkers cover
**worksheet** drawings only, so a chartsheet's full-sheet chart emits
no record (chartsheets stay in the `sheet_idx` numbering — they are
`<sheets>` entries — they just host no walkable anchors).

#### `conditional-formats` — the rule inventory

`zlsx conditional-formats <file>` (S3b) emits one record per `<cfRule>`:
sheets in workbook order, rules in sheet-document order, each carrying
its parent `<conditionalFormatting>` block's `sqref`. The record
reports the rule envelope (where it applies, what kind it is, its
formula bodies, its differential style id and priority), not the
visual payload: `<colorScale>` / `<dataBar>` / `<iconSet>` children
and the `<dxfs>` styles they point into stay in their parts,
byte-preserved, for callers that need them raw (`zlsx_pkg.PartStore`;
`styles` does not list dxfs). The records come from a strict,
namespace- and depth-aware walk of each sheet part
(`pkg/conditional_format_ndjson.zig`) — the same inventory the Zig
view `Worksheet.conditionalFormats` models, read so the wire can
prove it whole; the Zig view itself keeps its historical lenient
lexical contract, the split the `anchors` walkers established.

```jsonl
{"kind":"conditional_format","sheet":"Data","sheet_idx":0,"sqref":"A1:A4","rule_type":"cellIs","formulas":["2","4"],"dxf_id":0,"priority":1}
{"kind":"conditional_format","sheet":"Data","sheet_idx":0,"sqref":"C1:C4","rule_type":"colorScale","formulas":[],"dxf_id":null,"priority":3}
```

| Field | Meaning |
|---|---|
| `sheet`, `sheet_idx` | The sheet whose part carries the rule. Dropped under `--output compact-ndjson` in favour of the sheet prologue. |
| `sqref` | The parent block's target list as authored — one or more space-separated A1 areas, entity-decoded, not normalised or split. |
| `rule_type` | The `type` attribute as written (`cellIs`, `expression`, `colorScale`, `dataBar`, `iconSet`, `containsBlanks`, …) — a token inventory this read does not police. |
| `formulas` | The rule's `<formula>` bodies, entity-decoded, in document order — up to the schema's three (a `cellIs` `between` carries two, most rules one, the visual-payload rules none). Each body is text-only and a direct child of its rule; a fourth formula, one that never closes, or one carrying markup (an element, a comment, a CDATA section — markup is not the formula, the `defined-names` ruling) refuses the read. |
| `dxf_id` | The differential-style id (0-based into the styles part's `<dxfs>`), or `null`. |
| `priority` | The rule's cascade priority (lower wins), or `null`. |

Sheet selection follows the read family — `--sheet` / `--name` narrow
to one sheet, `--all-sheets` / `--sheet-glob` widen, the default
streams every sheet — and `--skip` / `--take` page the record stream.
A workbook without conditional formatting is an empty, successful
stream. The command routes through the package layer like `pivots`, so
an archive the lenient reader tolerates but the package layer refuses
is exit 2, and so is any workbook sheet the package layer cannot read
whole — sheet identities come from a strict namespace- and depth-aware
read of the workbook's `<sheets>` list (a ghost `<sheet>` under
`<extLst>` is not an identity, an entry missing a required carrier
refuses instead of vanishing, an empty list refuses — a sheetless
workbook is never served as an empty inventory — and any root-declared
relationships-namespace prefix spells the id reference — `q:id` under
`xmlns:q` reads like `r:id`), the sheet graph resolves strictly (the
relationship under the entry's id — matched by its DECODED spelling —
must exist, be a sheet-family type, be internal, and reach a part the
archive holds that no other sheet reaches — the `anchors` rule; the
rels part itself verifies strictly too, so a nested `Relationship`
decoy or a duplicate id refuses), and a worksheet-rooted part
additionally shares the
typed view's parse verdict (the same open-time verdicts every
`Worksheet` reader shares; a macrosheet part has no typed view, so
the strict walk's verdict stands alone there) — even for a sheet the
selection would exclude: the inventory is proven whole before
selection and pagination apply.

The strict walk is what lets the stream claim "one record per rule,
no more, no fewer". A `conditionalFormatting` element is a rule block
only as a main-namespace direct child of the sheet root (Transitional
or Strict namespace, bound as the document's default — a root that
binds neither, or a part that aliases the main namespace to a prefix,
refuses, the host-predicate rule), a `cfRule` only as a direct child
of a block, a `formula` only as a text-only direct child of a rule —
so an extension tree spelling the same names (an `x14` subtree
rebinding the default namespace, a decoy under `<extLst>`) can never
ghost a record. Namespace bindings compare by their decoded value —
an entity-spelled alias (`…spreadsheetml&#47;2006/main`) is an alias.
The grammar is XML's, not a byte pattern: whitespace
around `=` and either quote read as authored, a closing tag may space
before its `>` (`</formula >`), attribute names follow XML's Name
grammar (`<cfRule/x="y"` is not rule machinery), and a raw `<` in any
attribute value refuses. Refused whole (exit 2), nothing
written: mismatched or unterminated nesting anywhere in the part —
a truncated part, a second root, and nesting past the walker's
1024-level ceiling included; a `<!DOCTYPE>` or any other markup
declaration (an internal DTD could define entities whose expansion a
byte walk cannot see — Excel writes none); a `<formula>` that never
closes (a refusal, not an absence) or carries markup; a fourth
formula (the schema's `maxOccurs` is 3); a duplicate attribute on a
block or rule tag; a `sqref`, rule type or formula whose carrier does not decode (a
bad entity — the character-reference grammar is exact: `&#+65;` and
`&#1_0;` are not references) or decodes to non-UTF-8 (NDJSON must
stay parseable); a sheet-name carrier that does not decode; a numeric
carrier (`dxfId`, `priority`) whose entities do not decode — decoding
runs first, and only a carrier that decodes falls through to the
written-but-invalid `null` convention below; and — the
walk runs no markup-compatibility (MCE) processor — the MCE constructs
whose projection could reach a rule-carrying slot: an MCE-bound
element as a direct child of the sheet root, a block or a rule
(`mc:AlternateContent` substitutes a branch in place), the
`mc:ProcessContent` attribute anywhere (it lifts a wrapper's
children), a default binding of the MCE namespace, and an MCE element
directly inside the workbook's `<sheets>` list. The `mc` declaration
and `mc:Ignorable` on a root are inert — every modern Excel part
carries them — and a deeper `mc:AlternateContent` (the real
`<oleObject>` alternate shape) sits inside an already-opaque subtree
and stays walkable. A partial or wrong rule inventory is the shape of a
guard hole. Boundary conventions, pinned as the contract: an absent
`sqref` or `type` and an empty one are one shape (the empty string);
a `dxfId` or `priority` that decodes but is not a plain string of
decimal digits in `u32` range reads as `null` (the written-but-invalid
convention `tableHeaderRowCount` set — `+1` and `1_0` are not
numbers; entities resolve first, so `&#49;` is `1` and an undecodable
carrier is exit 2, not `null`); a `<conditionalFormatting>` block with no `<cfRule>`
children — the self-closing spelling included — emits nothing; a
`chartsheet` or `dialogsheet` part holds no rules by schema and
contributes an empty inventory (it is still walked whole — its
nesting must prove like any other part's), while a `macrosheet`
(which can carry conditional formatting) is walked like a worksheet. Extension-list
conditional formatting (`x14` data bars, icon sets carried under
`<extLst>`) is a different tree and not part of this read — the edit
paths move its `<xm:sqref>` / `<xm:f>` carriers with the grid (S2),
but no read models it yet (chart `<c:f>` series formulas move the same
way — the chart sweep — and `anchors` reports the respelled refs). The main-tree rule machinery moves whole
under row / column edits too (S3b slice 6):
`conditionalFormatting@sqref` and `dataValidation@sqref` shift with
the grid — merge-rect interval semantics: shift, grow on an insert
inside, shrink on a delete inside, a collapsed area dropped from the
list, whole-column (`A:A`) and whole-row (`1:1`) areas absorbing the
cross-axis edit and shifting as intervals along their own, `$`
anchors preserved, the value entity-decoded before parsing — in the
same edit that the formula sweep moves the rule's `<formula>` bodies
(all three schema slots), so this read reports envelope and bodies
on one grid after an edit. An sqref area the strict grammar cannot
read refuses the whole edit before anything mutates, and so does a
delete that collapses EVERY area of a rule (`SqrefCollapseUnsafe`):
Excel deletes such a rule outright, and zlsx refuses rather than
silently retarget it to the cells that slide into its place.

#### `sheet-props` — panes and the extent

`zlsx sheet-props <file>` (S3b) emits one record per workbook sheet,
workbook order: the sheet's `<dimension ref>` — the used-range extent as
authored — and the `<pane>` of its first `<sheetView>` as authored:
`xSplit` / `ySplit` / `topLeftCell` / `activePane` / `state`. This is
the read the Zig views `Worksheet.dimension` and `Worksheet.freezePane`
model, promoted to a typed report with one deliberate divergence: the
lenient `freezePane` narrows to frozen panes (`state="frozen"` /
`"frozenSplit"`, integer splits, the first such pane anywhere in the
part) and keeps that contract; this read reports split panes too, as
written, from the schema slot only. The records come from a strict,
namespace- and depth-aware walk of each sheet part
(`pkg/sheet_props_ndjson.zig`, on the `conditional-formats` scanner's
lexical layer). The C ABI (`zlsx_editor_sheet_props_ndjson`) and
py-zlsx (`Editor.sheet_props()` / `zlsx.sheet_props(path)`) hand over
these exact bytes with no selector.

```jsonl
{"kind":"sheet_props","sheet":"Data","sheet_idx":0,"dimension":"A1:B3","pane":{"x_split":2,"y_split":1,"top_left_cell":"C2","active_pane":"bottomRight","state":"frozen"}}
{"kind":"sheet_props","sheet":"Report","sheet_idx":1,"dimension":"A1:C2","pane":{"x_split":2865,"y_split":1215.5,"top_left_cell":"C4","active_pane":"bottomRight","state":"split"}}
{"kind":"sheet_props","sheet":"Bare","sheet_idx":2,"dimension":null,"pane":null}
```

| Field | Meaning |
|---|---|
| `sheet`, `sheet_idx` | The sheet. Dropped under `--output compact-ndjson` in favour of the sheet prologue (one per record here — every sheet contributes exactly one). |
| `dimension` | The `<dimension ref>` as authored (entity-decoded, not normalised, not recomputed from the cells), or `null` when the sheet spells neither the element nor its attribute — a fresh zlsx `Writer` workbook, for one. |
| `pane` | The first `<sheetView>`'s `<pane>`, or `null` when the sheet or that view has none. Present, each field is the attribute as authored or `null` when absent — no schema default is applied (Excel reads an absent `activePane` as `bottomRight` and an absent `state` as `split`; the wire says what the file says). |
| `pane.x_split`, `pane.y_split` | `xSplit` / `ySplit` as JSON numbers: a frozen pane's column / row counts, a split pane's position in twips (1/20 pt), a fractional value kept. |
| `pane.top_left_cell` | The `topLeftCell` A1 reference — the first visible cell of the bottom-right pane. |
| `pane.active_pane` | The `activePane` token (`bottomRight`, `topRight`, `bottomLeft`, `topLeft`) — a token inventory this read does not police. |
| `pane.state` | The `state` token (`frozen`, `frozenSplit`, `split`) — likewise. |

Sheet selection follows the read family — `--sheet` / `--name` narrow
to one sheet, `--all-sheets` / `--sheet-glob` widen, the default
streams every sheet — and `--skip` / `--take` page the record stream.
Every sheet emits exactly one record, so the unselected stream's length
is the workbook's sheet count. The command routes through the package
layer like `pivots`, under the `conditional-formats` inventory and
refusal rules: sheet identities come from the strict `<sheets>` read,
the sheet graph resolves strictly, a worksheet-rooted part shares the
typed view's parse verdict, and any sheet part the walk cannot prove —
even one the selection would exclude — refuses the whole command (exit
2), nothing written.

The strict walk is what lets each record claim "this is the slot". A
`dimension` is the extent only as a main-namespace direct child of a
`worksheet` or `macrosheet` root (the roots the schema gives one); a
`sheetViews` only as a direct child of a root whose views carry panes
(`worksheet`, `macrosheet`, `dialogsheet` — a `chartsheet`'s
`<sheetView>` cannot hold a pane by schema, so a chartsheet's record is
two nulls); a `sheetView` only as its direct child; a `pane` only as a
direct child of the FIRST `sheetView`. Everything else is opaque: a
`<pane>` under `<customSheetViews>` (a real corpus shape), a
`<dimension>` under `<extLst>` or inside a subtree that rebinds the
default namespace, a commented-out slot — none of these is reported,
none refuses. Later `<sheetView>` elements (Excel writes one per extra
window open on the sheet) keep their own panes in the part; the record
reports the primary view's. Refused whole (exit 2), nothing written: a
second `<dimension>`, `<sheetViews>` or first-view `<pane>` (each
maxOccurs=1 — which one Excel honours is not the reader's to guess), a
duplicate attribute on that machinery, the main namespace bound to a
prefix, mismatched or unterminated nesting, a DOCTYPE, an MCE
construct at a recognized slot (the walk has no MCE processor: an
`mc:AlternateContent` under the root, the views wrapper or the first
view; `mc:ProcessContent` anywhere), and a carrier the stream cannot
carry faithfully — a `ref` / `topLeftCell` / `activePane` / `state`
that does not decode, is not UTF-8, or carries markup. Boundary
conventions, pinned as the contract: `x_split` / `y_split` decode
their entities first (`&#50;` is `2`; an undecodable carrier is exit
2, not `null`), collapse boundary XML whitespace (XSD's
`whiteSpace="collapse"` for the atomic types — `" 2 "` is `2`; an
interior run is not a number), then read as an `xsd:double` lexical
or `null` — the written-but-invalid convention the family's numeric
attributes set (`1_0`, `0x10`, `abc` are not doubles; `INF` / `NaN`
are, but have no JSON spelling and read `null` too). A `<dimension>` without `ref` and
a `<pane/>` without attributes are present-and-empty: `dimension` is
`null` either way, `pane` is an object of nulls.

#### `calc-props` — the calculation properties

`zlsx calc-props <file>` (S3b) emits the workbook's `<calcPr>` as
**one** NDJSON record: `calcId`, `fullCalcOnLoad`, `iterate`,
`iterateCount`, `iterateDelta` — the field set the Zig view
`Workbook.calcProperties` models, read strictly from `xl/workbook.xml`
and reported as authored. The lenient view applies its own defaults
(`iterate` / `fullCalcOnLoad` false when absent, anything but `true` /
`1` false) and keeps them; this read reports `null` for what the file
does not say. The C ABI (`zlsx_editor_calc_props_ndjson`) and py-zlsx
(`Editor.calc_props()` / `zlsx.calc_props(path)`) hand over this exact
line.

```jsonl
{"kind":"calc_props","calc_id":191029,"full_calc_on_load":true,"iterate":false,"iterate_count":100,"iterate_delta":0.001}
```

Absent attributes are `null` — never omitted — and a workbook with no
`<calcPr>` at all is a record of nulls under exit 0 (the `doc-props`
convention; an empty `<calcPr/>` is the same record — absent and empty
are one shape on the wire). `calc_id` / `iterate_count` are
`xsd:unsignedInt` (a digit-only `u32` or `null`), `full_calc_on_load`
/ `iterate` are `xsd:boolean` (`true` / `false` / `1` / `0`, else
`null` — the lenient view reads `yes` as false, this read reports that
nothing lexical was said), `iterate_delta` an `xsd:double` (or `null`
— a non-finite value has no JSON spelling); every typed attribute
decodes its entities first (an undecodable carrier refuses, exit 2),
then collapses boundary XML whitespace (XSD's `whiteSpace="collapse"`
— `" 100 "` is `100`), then types lexically. An entity-spelled or non-numeric `calcId` / `iterateCount` /
`iterateDelta` never reaches this read through any surface — the CLI,
the C ABI and py-zlsx all open through `Workbook.open`, whose lenient
parse refuses it first, at open (its own contract — the strict decode
above is the family's, kept as the record's contract and tested below
the opener).
The element is the slot only as a main-namespace direct child
of the `<workbook>` root: a `<calcPr>` under `<extLst>` or inside a
rebound subtree is not reported; two at the slot refuse — which one
Excel honours is not the reader's to guess — and so does a `<calcPr>`
reached from the root through MCE elements alone (an
`mc:AlternateContent` / `mc:Choice` / `mc:Fallback` chain — the one
path an MCE processor, which the walk lacks, could project it into
the slot by; the real root-level `mc:AlternateContent` Excel writes
for `x15ac:absPath` stays inert, and a `<calcPr>` below any ordinary
wrapper — `<extLst>` included, even inside such a branch — is neither
reported nor a refusal). The record is read and validated
whole before the first byte is written. The read runs the same strict
workbook walk the sheet inventory comes from, so a `<sheets>` list it
cannot prove refuses here too (exit 2), though no sheet part is read.
The command routes through the package layer like `pivots`, so an
archive the lenient reader tolerates but the package layer refuses is
exit 2. There is no sheet dimension: sheet selectors, `--format` and
`--output compact-ndjson` are tolerated and ignored (the `doc-props` /
`meta` wrapper-friendliness), and `--skip` / `--take` do not apply to
a single-record report.

#### `list-sheets` — names and visibility

`zlsx list-sheets <file>` emits one `{"kind":"sheet",…}` record per
workbook sheet, in workbook order:

| Field | Meaning |
|---|---|
| `sheet` | The sheet's name, entity-decoded. |
| `sheet_idx` | Its 0-based position in the reader's sheet inventory — `<sheets>` order, minus any `<sheet>` the reader cannot resolve (no `name`, no `r:id`, or a dangling relationship), which consumes no index. The same index `--sheet`, `rows`, `cells` and `zlsx_sheet_state` use. |
| `state` | The `<sheet state="…">` attribute as the reader modelled it: `visible`, `hidden` (Excel's *Hide*) or `veryHidden` (unreachable from Excel's UI — only VBA / the object model reveals such a sheet, so this stream is how a caller scanning a workbook learns it exists). The value is compared as written — an entity-spelled value is unrecognised — and a missing or unrecognised attribute reads as `visible`, the schema default: visibility never fails an open. |

Hidden sheets stay in the inventory: `sheet_idx` counts them, `--sheet` /
`--name` select them, `rows` reads them. The same field reaches the C ABI
as `zlsx_sheet_state` (a code — 0 `visible` / 1 `hidden` / 2
`veryHidden`, `-1` out of range) and py-zlsx as `Book.sheet_state(selector)`
/ `Sheet.state` (this exact spelling), so the three surfaces agree sheet
for sheet. `--output pretty-json` collapses the stream into one
`{"sheets": […]}` object; `meta` carries the same per-sheet records (plus
`has_comments`) after its workbook record, which tallies them as
`hidden_sheet_count` / `very_hidden_sheet_count`.

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
| `embed --extract` | `--column A --coverage A2:A100` | Read-only: one `{"kind":"embed_row","row":N,"text":"…","hash":H}` record per embeddable row — `text` as a reader sees the cell (a shared or inline string's runs joined, entities resolved; a number's `<v>` as written; an error's literal, `#N/A`; a boolean as `1` / `0`), `hash` the canonical xxh3-64 content hash `--vectors` stores beside the vector, as an unsigned 64-bit decimal (a JSON reader that narrows integers to doubles — `jq` before 1.7 — rounds it; Python's `json` does not). Rows with nothing embeddable are omitted. Refuses (exit 3) rather than hand over a record that lies: `MalformedSheetXml` (a sheet part the view cannot parse, or a row or cell it cannot place — no `r`, or a ref under another row), `MalformedSharedStringsXml`, `UnsupportedCellValue` (a boolean `<v>` that is not 0 / 1, a `<v>` the number canonicalizer cannot read, a `t="d"` ISO-8601 date, a `t` this reader does not know, a bad shared-string index or entity), `SstIndexOutOfRange`, `InvalidUtf8`, `UnicodeNormalizationFailed`. A sheet the editor holds staged cell writes for is not this read's concern (the CLI opens fresh) |
| `embed --vectors P` | `--column A --coverage A2:A100 --model NAME --out P` | Write the embedding parts from `{"row":N,"vector":[…]}` NDJSON; covered rows with no vector become tombstones. Optional: `--id NAME`, `--dtype f32\|int8-sym`, `--sheet N` |
| `embed --prune` | `--out P` | The redaction sweep: tombstone every slot whose row is no longer embeddable and zero its vector (a row deleted in plain Excel leaves its vector on disk — this removes it). Changed-but-present content is reported stale, not redacted. Judges through `--extract`'s reading and refuses (exit 3) the same way it does — a covered cell the read cannot carry, or a row or cell the view cannot place, stops the sweep rather than redact a row that has text |
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

The C ABI and py-zlsx report the same three fields beside the row
iterator's cells (0.9.0+): `zlsx_rows_formula_at` / `zlsx_rows_formula_ref_at`
/ `zlsx_rows_error_at` with `zlsx_rows_style_at`'s per-column contract, and
`Rows.formula_strings()` / `formula_refs()` / `error_strings()` — one list per
accessor aligned to the row `next()` yielded. The cell itself stays what it
always was there: the cached value for a formula cell (`cached` here,
omitted for a formula-only cell, whose cell is empty) and the literal string
for an error cell. The
precedence is the same on every surface — a formula whose cached value is an
error literal is `t:"formula"`, never `t:"error"`. Contract:
`docs/plans/c-abi-status-v1.md` §17.

---

## Flags

**Default sheet scope** differs by family: `rows` / `cells` read sheet 0
unless told otherwise; `comments` / `validations` / `hyperlinks` / `pivots` /
`merges` / `anchors` / `conditional-formats` / `sheet-props` stream every
sheet; `styles` / `sst` / `meta` / `list-sheets` are workbook-wide, and so are
`defined-names` (a sheet selector there narrows by a name's *scope* — see its
contract above), `doc-props` and `calc-props` (no sheet dimension at all —
selectors are tolerated and ignored).

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
                              # validations / hyperlinks / pivots / merges / anchors /
                              # conditional-formats / sheet-props — a no-op on
                              # workbook-scoped sub-commands (defined-names,
                              # doc-props and calc-props included)
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

# Every cellIs rule with exactly two formula bodies (between / notBetween — the
# operator itself is not on the wire), as TSV.
zlsx conditional-formats data.xlsx | jq -r 'select(.rule_type=="cellIs" and (.formulas|length)==2) | [.sheet, .sqref, .formulas[0], .formulas[1]] | @tsv'

# Every frozen pane: sheet, frozen column / row counts, first scrolling cell.
zlsx sheet-props data.xlsx | jq -r 'select(.pane.state=="frozen" or .pane.state=="frozenSplit") | [.sheet, .pane.x_split, .pane.y_split, .pane.top_left_cell] | @tsv'

# Does the workbook ask Excel to recalculate everything on open?
zlsx calc-props data.xlsx | jq '.full_calc_on_load == true'

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
| 2 | Could not open the input: missing file, permission denied, not a valid xlsx archive, malformed parts at open time — or, on `pivots`, a pivot graph that cannot be read whole (a named part missing or unreadable, a cache identity that disagrees) — or, on `defined-names`, a name inventory the read cannot serve faithfully (a carrier that does not decode, malformed UTF-8, a body with embedded markup — its contract above) — or, on `merges`, a selected sheet with merges under a non-UTF-8 name — or, on `doc-props`, a docProps part the store cannot read or a field value that is not UTF-8 — or, on `anchors`, an anchor inventory the read cannot serve faithfully (a sheet the read cannot place, a drawing / image / chart relationship that dangles, an anchor that does not parse, a carrier that does not decode, malformed UTF-8, a series ref with embedded markup — its contract above) — or, on `conditional-formats`, a rule inventory the read cannot serve faithfully (a sheet part the strict walk cannot prove whole — mismatched nesting, a namespace shape that could ghost a rule, an unterminated or markup-carrying formula, a duplicate attribute on the rule machinery — or a sqref / rule type / formula / sheet-name carrier that does not decode or is not UTF-8 — its contract above) — or, on `sheet-props`, a sheet part the strict walk cannot prove a pane / extent for (a second `<dimension>` / `<sheetViews>` / first-view `<pane>`, a duplicate attribute on that machinery, an MCE construct at a recognized slot, or a `ref` / pane carrier that does not decode or is not UTF-8 — its contract above) — or, on `calc-props`, a `<calcPr>` slot the read cannot report faithfully (two at the slot, one an MCE branch could project there, a duplicate attribute, a carrier that does not decode — its contract above) |
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
