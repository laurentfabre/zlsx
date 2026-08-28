# Surface matrix — the four-surface capability truth

> Born in `goal_sigmoid.md` row S0 (2026-08-26). **A PR that adds, lifts or
> refuses a capability on any surface updates this file in the same PR.** A
> sigmoid row is `done` only when its line here reads `✓ ✓ ✓ ✓` or carries an
> owner ruling. Cells name the entry point rather than a bare tick, and
> `scripts/check_surface_matrix.py` (CI: `unicode-and-notices`) fails when a
> named symbol does not exist *on its owner* on that surface — run it before
> committing an edit here.

_Last audited: 2026-08-26 against `main` at `d99a386` (v0.8.0, Zig 0.16.0).
The audit line moves only when an audit row (S0, S10) re-verifies every cell;
capability PRs edit cells, not this line._

## Legend and cell grammar

| Cell | Meaning |
|---|---|
| ✓ `sym` | Shipped. `sym` is the entry point — function, export, method, or sub-command / flag. |
| ~ `sym` ⁿ | Partial. `sym` is where the partial behaviour lives; the footnote says exactly what is missing. |
| ⛔ `Error` | Refused by design with that typed error. Refusing is the contract, not a gap. |
| — Sx | Absent; `Sx` is the sigmoid row that closes it. |
| n/a | Permanently absent on that surface by owner ruling — one line per cell in [Rulings](#rulings). |

Every symbol is back-ticked and resolvable on its own surface, on its named
owner: C symbols are the full export names in `include/zlsx.h`; Python names
are `Class.method` (a `def` inside that class) or `zlsx.function` (module
level) in `bindings/python/zlsx/__init__.py`; CLI names are sub-commands or
flags parsed in `src/cli.zig` / `src/formula_cli.zig`; Zig names are
`Type.member` (a `pub fn`, `pub const` or field inside that type's block),
`zlsx.fn` (module level in `src/xlsx.zig`), `zlsx_pkg.fn` (exported by
`pkg/root.zig`) or `zlsx_recalc.fn` (`recalc/recalc.zig`). No prose in
cells — a capability that needs a sentence gets a footnote.

The four surfaces:

| Surface | What it means here |
|---|---|
| **Zig** | The three public modules `build.zig` exports with `addModule`: `zlsx` (`src/xlsx.zig`), `zlsx_pkg` (`pkg/`, export list in `pkg/root.zig`), `zlsx_recalc` (`recalc/recalc.zig`) — plus `src/writer.zig`, which today is a *private* `createModule` wired into the CLI, the C ABI and `zlsx_recalc` only (§2, first row). |
| **C** | `include/zlsx.h`, implemented in `src/c_abi.zig`. |
| **Py** | `bindings/python/zlsx` (ctypes over the C ABI — it can never lead C). |
| **CLI** | The `zlsx` binary (`src/cli.zig`, `src/formula_cli.zig`). The `zlsx-extract-images` sibling binary is called out where it matters. |

---

## 1 · Read

| Capability | Zig | C | Py | CLI | Row |
|---|---|---|---|---|---|
| Open a workbook file | ✓ `Book.open`, `Workbook.open` | ✓ `zlsx_book_open` | ✓ `zlsx.open` | ✓ `rows` ⁰ | |
| Open from in-memory bytes | ✓ `Book.openBuffer` | ✓ `zlsx_book_open_buffer` | ✓ `zlsx.open_bytes` | n/a | |
| Lazy per-sheet loading | ✓ `Book.openLazy`, `Book.preloadSheet`, `Book.streamSheet` | — S3e | — S3e | — S3e | S3e |
| Lazy SST backend (millions of unique strings) | ✓ `Book.openSstLazy` | — S3e | — S3e | ✓ `--sst-lazy` | S3e |
| Row iteration, typed cells (str / int / num / bool / empty) | ✓ `Book.rows`, `Rows.next` | ✓ `zlsx_rows_open`, `zlsx_rows_next`, `zlsx_rows_skip` | ✓ `Sheet.rows` | ✓ `rows`, `cells` | |
| Whole-sheet bulk read | ✓ `Book.materialiseSheet` | ✓ `zlsx_matrix_open`, `zlsx_matrix_data` | ✓ `Sheet.read_all`, `zlsx.read` | n/a | |
| Sheet names, lookup by name | ✓ `Book.sheets`, `Book.sheetByName` | ✓ `zlsx_sheet_count`, `zlsx_sheet_name`, `zlsx_sheet_index_by_name` | ✓ `Book.sheets`, `Book.sheet` | ✓ `list-sheets`, `--name` | |
| Sheet visibility (`hidden` / `veryHidden`) | ✓ `Sheet.state` | — S3b | — S3b | ✓ `list-sheets` | S3b |
| Shared strings + rich-text runs | ✓ `Book.sharedStringAt`, `Book.richRuns` | ✓ `zlsx_shared_string_at`, `zlsx_rich_run_at` | ✓ `Book.shared_string_at`, `Book.rich_text` | ✓ `sst` | |
| Dates, both epochs → `DateTime` | ✓ `Rows.parseDate`, `zlsx.fromExcelSerial`, `zlsx.fromExcelSerial1904` | ✓ `zlsx_rows_parse_date`, `zlsx_is_date_format` | ✓ `Rows.parse_date`, `Book.is_date_format` | ✓ `rows`, `cells` ¹ | |
| Date → Excel serial | ✓ `zlsx.toExcelSerial` | ✓ `zlsx_datetime_to_serial` | ✓ `zlsx.to_excel_serial` | — S11 | S11 |
| Formula text on read (own + shared-formula base) | ✓ `Rows.formulaStrings`, `Rows.formulaRefs` | — S3b | — S3b | ✓ `cells` ² | S3b |
| Error cells distinguished from text | ✓ `Rows.errorStrings` | ~ `zlsx_rows_next` ³ | ~ `Rows.__next__` ³ | ✓ `cells` ² | S3b |
| Cell style index + resolved font / fill / border / alignment / number format | ✓ `Rows.styleIndices`, `Book.cellFont`, `Book.cellFill`, `Book.cellBorder`, `Book.cellAlignment`, `Book.numberFormat` | ✓ `zlsx_rows_style_at`, `zlsx_cell_font`, `zlsx_cell_fill`, `zlsx_cell_border`, `zlsx_cell_alignment`, `zlsx_number_format` | ✓ `Rows.style_indices`, `Book.cell_font`, `Book.cell_fill`, `Book.cell_border`, `Book.cell_alignment`, `Book.number_format` | ✓ `cells`, `styles` | |
| Indexed palette (`indexed="N"`) + `tint` resolution | — S4 ⁴ | — S4 | — S4 | — S4 | S4 |
| Merged ranges | ✓ `Book.mergedRanges` | ✓ `zlsx_merged_range_at`, `zlsx_merged_range_count` | ✓ `Book.merged_ranges` | — S3b | S3b |
| Hyperlinks (external + internal) | ✓ `Book.hyperlinks` | ✓ `zlsx_hyperlink_at`, `zlsx_hyperlink_count`, `zlsx_hyperlink_location_at` | ✓ `Book.hyperlinks` | ✓ `hyperlinks` | |
| Data validations | ✓ `Book.dataValidations` | ✓ `zlsx_data_validation_at`, `zlsx_data_validation_count` | ✓ `Book.data_validations` | ✓ `validations` | |
| Comments (+ rich runs) | ✓ `Book.comments` | ✓ `zlsx_comment_at`, `zlsx_comment_count`, `zlsx_comment_run_at` | ✓ `Book.comments` | ✓ `comments` | |
| Defined names | ✓ `Workbook.definedNames`, `Workbook.definedNamesForSheet` | — S3b | — S3b | — S3b | S3b |
| Conditional formats | ✓ `Worksheet.conditionalFormats` | — S3b | — S3b | — S3b | S3b |
| Freeze panes, `<dimension>`, calc properties | ✓ `Worksheet.freezePane`, `Worksheet.dimension`, `Workbook.calcProperties` | — S3b | — S3b | — S3b | S3b |
| Image / chart anchors | ✓ `zlsx_pkg.imageAnchors`, `zlsx_pkg.chartAnchors` | — S3b | — S3b | — S3b | S3b |
| Document properties | ✓ `Workbook.docProps` | ✓ `zlsx_editor_docprop_at`, `zlsx_editor_has_custom_properties` | ✓ `Editor.doc_props` | — S3b | S3b |
| Embedding vectors (state / model / dim / dtype / coverage / vectors / hashes / digest / carrier / tombstones) | ✓ `Workbook.embeddings` | ✓ `zlsx_emb_open`, `zlsx_emb_state`, `zlsx_emb_model`, `zlsx_emb_dim`, `zlsx_emb_dtype`, `zlsx_emb_coverage_count`, `zlsx_emb_coverage_id`, `zlsx_emb_coverage_sheet`, `zlsx_emb_coverage_range`, `zlsx_emb_coverage_rows`, `zlsx_emb_vectors`, `zlsx_emb_hashes`, `zlsx_emb_digest`, `zlsx_emb_carrier`, `zlsx_emb_tombstone` | ✓ `zlsx.embeddings` | — S3c ⁵ | S3c |
| Pivot tables, typed — tables (host sheet, `location`, axes, data fields), caches (source resolved to its sheet, field schema, records part) | ✓ `Workbook.pivotTables` ⁶ | — S6 | — S6 | ✓ `pivots` ⁶ | S6 |

⁰ Every read sub-command opens the file; `rows` is the default and the lint anchor. `dbx genie` is the one sub-command that takes no workbook.
¹ `rows` / `cells` tag date cells `t:"date"` unconditionally, with the ISO value and the raw `serial` — `docs/cli.md`, "The NDJSON row envelope".
² `cells` / `rows` emit `t:"formula"` (`formula` / `formula_ref` / `cached`) and `t:"error"` (`v` is the literal).
³ The C cell tags are `ZLSX_CELL_{STRING,INTEGER,NUMBER,BOOLEAN,EMPTY}` (`include/zlsx.h`); an error literal (`#DIV/0!`, `#N/A`, …) arrives as an ordinary string cell and Python converts it as one. Zig readers get the tag from `Rows.errorStrings` alongside `Rows.next` (`src/xlsx.zig`: "the Cell union slot stays `.string`").
⁴ Theme colours resolve through the workbook palette on every surface; the legacy indexed table and `tint` math do not, anywhere.
⁵ `embed --extract` lists the rows that need embedding; no mode dumps the stored vectors or the recovery state.
⁶ Read-only typed graph (`pkg/pivots.zig`; parsers in `pkg/typed_parts/pivot_xml.zig`, Strict + Transitional by namespace prefix): every pivot table with its host sheet and `location`, every cache with its `worksheetSource` resolved to a sheet of the workbook (the `sheet` attribute, a table's host, or a defined name's body — through the engine's symbol table), to another workbook, or to nothing. Formats, hierarchies and OLAP elements stay raw; the parts are byte-preserved through every edit on every surface. The NDJSON contract is `docs/cli.md` "pivots"; the C and Python legs follow the S6 gate. See §3 footnote ¹⁰ for how far the refusal guard reaches.

## 2 · Write — fresh workbooks (`Writer`)

| Capability | Zig | C | Py | CLI | Row |
|---|---|---|---|---|---|
| The fresh `Writer` reachable from outside the repo | — S5 ⁷ | ✓ `zlsx_writer_create` | ✓ `zlsx.Writer` | — S11 | S5 · S11 |
| Create; add sheets | ✓ `Writer.init`, `Writer.addSheet` | ✓ `zlsx_writer_create`, `zlsx_writer_add_sheet` | ✓ `zlsx.Writer`, `Writer.add_sheet` | — S11 | S11 |
| Hidden sheets on write | — S10 | — S10 | — S10 | — S11 | S10 · S11 |
| Typed rows (str / int / num / bool / empty) | ✓ `SheetWriter.writeRow` | ✓ `zlsx_sheet_writer_write_row` | ✓ `SheetWriter.write_row` | — S11 | S11 |
| Styles — fonts, 19 fills, 14 × 5 borders, alignment, number formats | ✓ `Writer.addStyle`, `SheetWriter.writeRowStyled` | ✓ `zlsx_writer_add_style`, `zlsx_writer_add_style_ex`, `zlsx_sheet_writer_write_row_styled` | ✓ `Writer.add_style`, `SheetWriter.write_row` | — S11 | S11 |
| Rich-text runs per cell | ✓ `SheetWriter.writeRichRow` | ✓ `zlsx_sheet_writer_write_rich_row` | ✓ `SheetWriter.write_rich_row` | — S11 | S11 |
| Formulas + caller-supplied cached value; CSE rectangles | ✓ `SheetWriter.writeRowWithFormulas`, `SheetWriter.writeRowWithFormulaCells` | ✓ `zlsx_sheet_writer_write_row_with_formulas`, `zlsx_sheet_writer_write_row_with_formulas_v2` | ✓ `SheetWriter.write_row_with_formulas`, `FormulaSpec.cse` | — S11 | S11 |
| Column widths, row heights, freeze panes, auto-filter | ✓ `SheetWriter.setColumnWidth`, `SheetWriter.setRowHeight`, `SheetWriter.freezePanes`, `SheetWriter.setAutoFilter` | ✓ `zlsx_sheet_writer_set_column_width`, `zlsx_sheet_writer_set_row_height`, `zlsx_sheet_writer_freeze_panes`, `zlsx_sheet_writer_freeze_panes_checked`, `zlsx_sheet_writer_set_auto_filter` | ✓ `SheetWriter.set_column_width`, `SheetWriter.set_row_height`, `SheetWriter.freeze_panes`, `SheetWriter.set_auto_filter` | — S11 | S11 |
| Merged cells | ✓ `SheetWriter.addMergedCell` | ✓ `zlsx_sheet_writer_add_merged_cell` | ✓ `SheetWriter.add_merged_cell` | — S11 | S11 |
| Hyperlinks, external + internal | ✓ `SheetWriter.addHyperlink`, `SheetWriter.addInternalHyperlink` | ✓ `zlsx_sheet_writer_add_hyperlink`, `zlsx_sheet_writer_add_internal_hyperlink` | ✓ `SheetWriter.add_hyperlink`, `SheetWriter.add_internal_hyperlink` | — S11 | S11 |
| Comments | ✓ `SheetWriter.addComment` | ✓ `zlsx_sheet_writer_add_comment` | ✓ `SheetWriter.add_comment` | — S11 | S11 |
| Data validations — list / numeric / custom | ✓ `SheetWriter.addDataValidationList`, `SheetWriter.addDataValidationNumeric`, `SheetWriter.addDataValidationCustom` | ✓ `zlsx_sheet_writer_add_data_validation_list`, `zlsx_sheet_writer_add_data_validation_numeric`, `zlsx_sheet_writer_add_data_validation_custom` | ✓ `SheetWriter.add_data_validation_list`, `SheetWriter.add_data_validation_numeric`, `SheetWriter.add_data_validation_custom` | — S11 | S11 |
| Conditional formats — cellIs / expression / colorScale / dataBar, with `dxf` | ✓ `SheetWriter.addConditionalFormatCellIs`, `SheetWriter.addConditionalFormatExpression`, `SheetWriter.addConditionalFormatColorScale`, `SheetWriter.addConditionalFormatDataBar`, `Writer.addDxf` | ✓ `zlsx_sheet_writer_add_conditional_format_cell_is`, `zlsx_sheet_writer_add_conditional_format_expression`, `zlsx_sheet_writer_add_conditional_format_color_scale`, `zlsx_sheet_writer_add_conditional_format_data_bar`, `zlsx_writer_add_dxf` | ✓ `SheetWriter.add_conditional_format_cell_is`, `SheetWriter.add_conditional_format_expression`, `SheetWriter.add_conditional_format_color_scale`, `SheetWriter.add_conditional_format_data_bar`, `Writer.add_dxf` | — S11 | S11 |
| Defined names (incl. hidden) | ✓ `Writer.addDefinedName` | ✓ `zlsx_writer_add_defined_name` | ✓ `Writer.add_defined_name` | — S11 | S11 |
| Save to file | ✓ `Writer.save` | ✓ `zlsx_writer_save` | ✓ `Writer.save` | — S11 | S11 |
| Save to bytes | ✓ `Writer.saveToOwnedBuffer` | ✓ `zlsx_writer_save_to_buffer` | ✓ `Writer.to_bytes` | — S11 | S11 |
| Save with recalc (§5.7.9 transaction) | ✓ `zlsx_recalc.writerSaveWithRecalc` | ✓ `zlsx_writer_save_with_recalc` | ✓ `Writer.save` | — S11 | S11 |
| Images | — S5 ⁸ | — S5 | — S5 | — S11 | S5 · S11 |
| Charts | — S9 | — S9 | — S9 | — S11 | S9 · S11 |
| Pivot tables (one type, one contiguous source) | — S8 | — S8 | — S8 | — S11 | S8 · S11 |
| Sheet-name validation on the fresh writer — 31-scalar cap, reserved names, NFC + full-casefold duplicate check | ✓ `Writer.addSheet` | ✓ `zlsx_writer_add_sheet` | ✓ `Writer.add_sheet` | — S11 | S11 |

⁷ `src/writer.zig` is built with `b.createModule` and imported as `writer` by the CLI, `src/c_abi.zig` and `recalc/`; it is not one of the three `addModule` exports, so a Zig project that depends on zlsx (e.g. `nemonym`, which imports `zlsx` and `zlsx_pkg`) cannot name `Writer` at all — its fresh-workbook path is `Workbook.empty` → `Workbook.saveFreshEmit` or the C ABI. Every ✓ below in this column is true *inside* the repo. One `addModule("zlsx_writer", …)` closes it; proposed for S5, the row that reshapes the Writer API.
⁸ The fresh `Writer` has no image API at all. Image authoring exists today on the *editing* layer only — §3, `Workbook.addImage`. S5 routes the Writer through that one emitter; no second emitter, ever.

## 3 · Edit — load-modify-save (`Editor` / `Workbook`)

| Capability | Zig | C | Py | CLI | Row |
|---|---|---|---|---|---|
| Open for editing | ✓ `Editor.open`, `Workbook.open` | ✓ `zlsx_editor_open` | ✓ `zlsx.edit`, `zlsx.Editor` | ✓ `append-rows` ⁸ | |
| Open from bytes | ✓ `Editor.openBuffer`, `Workbook.openBuffer` | ✓ `zlsx_open_buffer` | ✓ `Editor.from_bytes` | n/a | |
| Append rows | ✓ `Editor.appendRows` | ✓ `zlsx_editor_append_row` | ✓ `Editor.append_rows` | ✓ `append-rows` | |
| Set cells | ✓ `Editor.setCell`, `Editor.setCells` | ✓ `zlsx_editor_set_cell` | ✓ `Editor.set_cell`, `Editor.set_cells` | ✓ `set-cell` | |
| Delete a cell outright (not "write a blank") | ✓ `Worksheet.deleteCell` | — S3d | — S3d | — S3d | S3d |
| Save — atomic rename, untouched parts byte-preserved | ✓ `Editor.save` | ✓ `zlsx_editor_save` | ✓ `Editor.save` | ✓ `--out` | |
| Save to bytes | ✓ `Editor.saveToOwnedBuffer` | ✓ `zlsx_editor_save_to_buffer` | ✓ `Editor.save_to_buffer` | n/a | |
| Document-properties scrub | ✓ `Workbook.stripDocProps` | ✓ `zlsx_editor_strip_doc_props` | ✓ `Editor.strip_doc_props` | ✓ `scrub-metadata` | |
| Add / rename / delete sheet | ✓ `Editor.addSheet`, `Editor.renameSheet`, `Editor.deleteSheet` | — S3a | — S3a | ✓ `add-sheet`, `rename-sheet`, `delete-sheet` | S3a |
| Insert / delete row | ✓ `Editor.insertRow`, `Editor.deleteRow` | — S3a | — S3a | ✓ `insert-row`, `delete-row` | S3a |
| Insert / delete column | ✓ `Editor.insertColumn`, `Editor.deleteColumn` | — S3a | — S3a | ✓ `insert-column`, `delete-column` | S3a |
| Rename a table column (structured-ref rewrite) | ✓ `Editor.renameTableColumn` | — S3a | — S3a | — S3a | S3a |
| Sheet-name duplicate check on the edit path (add / rename) | ~ `Workbook.addSheet`, `Workbook.renameSheet` ¹⁴ | — S3a | — S3a | ~ `add-sheet`, `rename-sheet` ¹⁴ | S3a |
| Cross-part rewriters the structural edits carry — formulas (A1, 3D, R1C1, structured refs), defined names, hyperlinks, DV / CF, merges, panes, autoFilter, tables, drawings, comments | ✓ `Workbook.rewriteAllFormulas`, `Workbook.rewriteAllDefinedNames`, `Workbook.rewriteAllHyperlinkLocations`, `Workbook.rewriteAllValidationsAndConditionalFormats` ⁹ | — S3a | — S3a | ✓ `insert-row` ⁹ | S3a |
| Row / col edit on a sheet that only **hosts** a pivot — `location@ref` moves in step (S7a) | ✓ `Workbook.preflightPivotEditsForSheet` ¹⁰ | — S3a | — S3a | ✓ `insert-row` ⁹ | S3a |
| Refusal: row / col edit inside a hosted pivot's footprint, on a host sheet a pivot **may read from**, on a pivot part two sheets host, or on a pivot graph that cannot be read | ⛔ `PivotEditUnsafe` ¹⁰ | — S3a | — S3a | ⛔ `RowEditUnsafeForSheet`, `ColEditUnsafeForSheet` ¹⁰ | S7b · S7c |
| `<extLst>` `<xm:f>` formulas (sparkline ranges and date axes, x14 CF / DV) rewritten under every structural edit, with the host / target sheet context; a carrier the scan cannot read refuses the whole edit before its first mutation | ✓ `Workbook.rewriteAllExtensionFormulas`, `Workbook.preflightExtensionFormulas` ¹⁵ | — S3a | — S3a | ✓ `insert-row` ⁹ | S3a |
| Existing-workbook authoring — styles / dxf / number formats; column widths, row heights, panes, auto-filter, merges, hyperlinks, comments, DV, CF on an existing sheet | ✓ `Workbook.addStyle`, `Workbook.addDxf`, `Workbook.internNumFmt`, `Worksheet.setColumnWidth`, `Worksheet.setRowHeight`, `Worksheet.freezePanes`, `Worksheet.setAutoFilter`, `Worksheet.addMergedCell`, `Worksheet.addHyperlink`, `Worksheet.addInternalHyperlink`, `Worksheet.addComment`, `Worksheet.addDataValidationList`, `Worksheet.addDataValidationRange`, `Worksheet.addDataValidationCustom`, `Worksheet.addConditionalFormatCellIs`, `Worksheet.addConditionalFormatExpression`, `Worksheet.addConditionalFormatColorScale`, `Worksheet.addConditionalFormatDataBar` | — S3d | — S3d | — S3d | S3d |
| Add a defined name to an existing workbook | ✓ `Workbook.addDefinedName` | — S3d | — S3d | — S3d | S3d |
| Add images — native size, cell range (`twoCellAnchor`), explicit extent, append into an existing drawing | ✓ `Workbook.addImage`, `Workbook.addImageRange`, `Workbook.addImageAnchored` | — S5 | — S5 | — S5 | S5 |
| Extract / replace / remove embedded objects (images, charts), typed | ~ `PartStore.imageParts`, `PartStore.replacePart`, `PartStore.removePart` ¹¹ | — S5b | — S5b | ~ `zlsx-extract-images` ¹¹ | S5b |
| Mark recalc-on-load (`fullCalcOnLoad`) | ✓ `Workbook.markRecalcOnLoad` | ✓ `zlsx_editor_mark_recalc_on_load` | ✓ `Editor.mark_recalc_on_load` | ~ `--on-unsupported` ¹² | S3d |
| Recalculate — in memory, and as the atomic file transaction | ✓ `Workbook.recalculate`, `Workbook.saveWithRecalc` | ✓ `zlsx_editor_recalculate`, `zlsx_editor_save_with_recalc` | ✓ `Editor.recalculate`, `Editor.save_with_recalc` | ✓ `recalc` | |
| Evaluate one formula against the workbook | ✓ `Workbook.evaluate` | ✓ `zlsx_editor_evaluate` | ✓ `Editor.evaluate` | ✓ `eval` | |
| Cancellation / deadline on engine calls | ✓ `zlsx_pkg.Control`, `zlsx_pkg.CancelToken` | ✓ `zlsx_cancel_token_new`, `zlsx_cancel_token_trigger`, `zlsx_cancel_token_free` | ✓ `Editor.recalculate` (timeout kwarg) | ✓ `--deadline` | |
| Engine fingerprint | — S10 ¹³ | ✓ `zlsx_engine_fingerprint` | ✓ `zlsx.engine_fingerprint` | — S10 | S10 |

⁸ Every edit sub-command opens the file for editing; `append-rows` is the lint anchor.
⁹ The rewriters run under every structural edit on every surface that has the edit — the CLI cell names one sub-command as the lint anchor. `pkg/sheet_edit.zig`, `pkg/table_edit.zig`, `pkg/drawing_edit.zig` are the per-part transforms behind them.
¹⁰ The pivot guard starts from the relationships **of the edited sheet**: a sheet whose relationships name no pivot part is admitted without the graph being read, so a sheet a pivot only *reads from* (named by `worksheetSource` inside the cache definition) is **not detected** — S6's audit (2026-08-27) pinned it: the `S6 audit` tests in `pkg/editor.zig` show the source-only sheet admitted, with `worksheetSource@ref` left stale for a `sheet` + `ref` source (a table-named source follows the table rewriter and stays valid). `Workbook.pivotTables` resolves every source sheet (`Pivots.readsFromSheet`), which is what S7b guards on. For a sheet that *hosts* a pivot, S7a (2026-08-28) reads the graph and moves `pivotTableDefinition/location@ref` — the definition's one absolute coordinate, spliced at the parser's `Location.ref_span` by `pkg/pivots.zig::edit.applyToTableDefinition` — for an edit at or above / left of the rectangle; an edit below or right of it leaves the part byte-identical. `Workbook.preflightPivotEditsForSheet` dry-runs the move before the first mutation; the sweep applies it first and whole. The host's `<pivotSelection>` coordinates move with the grid like `<selection>`. `PivotEditUnsafe` remains for an edit inside the pivot's footprint (the rectangle plus a conservative report-filter band above it: `rowPageCount + 1` rows, `3 · colPageCount` columns — Excel refuses that edit too), a host some cache *may read* (`Pivots.mayReadFromSheet`: resolved to it, unresolved, or of unknown type), a pivot part two sheets host, a shift past the grid, and a sheet whose relationships name a cache part directly; a graph that cannot be read whole is `MalformedPivotXml`. `Editor` remaps both (and the S2 residual `MalformedExtensionXml`, footnote ¹⁵) to `RowEditUnsafeForSheet` (row edits) / `ColEditUnsafeForSheet` (column edits); the CLI reports those with exit 3. The S7a oracle question — how Excel re-lays a pivot whose rectangle moved — is parked with the owner (`goal_sigmoid.md` §5).
¹¹ Zig composes over *raw* parts: `PartStore.imageParts` lists image parts, `replacePart` / `removePart` swap or drop bytes by part name. There is no typed removal that also repairs the drawing, its relationships and content types, and no typed chart replacement (`ChartAnchor` is read-only). The sibling binary `zlsx-extract-images` extracts images by part name — no replace / remove, no chart parts, and not reachable as `zlsx <sub-command>`. S5b therefore includes Zig.
¹² `recalc --on-unsupported keep-stale-and-mark` sets the flag only as the fallback arm of a recalc; there is no standalone mark sub-command.
¹³ The fingerprint string is assembled privately in `src/c_abi.zig` from the `fingerprint_config` build option; no `zlsx` / `zlsx_pkg` / `zlsx_recalc` accessor exists.
¹⁴ The edit path compares sheet names **ASCII case-insensitively** (`pkg/workbook.zig`, `addSheet` / `renameSheet` / `sheetNameMatchesDecoded`); the fresh writer's NFC + full-casefold comparison (`casefold.excelSheetNameEql`, so `café` / `CAFÉ` and `ß` / `SS` collapse) is not used here. Length and reserved-name validation are shared.
¹⁵ S2 (2026-08-26) replaced #140's `ExtensionEditUnsafe` refusal with a rewrite: `<xm:f>` is the formula leaf of the `x14:` extensions, so it rides the formula rewriter (`on_sheet` = host sheet, `target_sheet` = edited sheet) rather than the byte transform that shifts its sibling `<xm:sqref>`. The one refusal left on the axis is a carrier `sheet_edit.nextXmFormula` cannot read (no `</xm:f>`, markup in the body): `Workbook.preflightExtensionFormulas` scans every sheet before an edit's first mutation, so the edit is all-or-nothing. C and Python observe the lift only once S3a gives them structural edits.

## 4 · Embeddings

| Capability | Zig | C | Py | CLI | Row |
|---|---|---|---|---|---|
| Read (state / model / dim / dtype / coverage / vectors / hashes / digest / carrier / tombstones) | ✓ `Workbook.embeddings` | ✓ `zlsx_emb_open`, `zlsx_emb_state`, `zlsx_emb_model`, `zlsx_emb_dim`, `zlsx_emb_dtype`, `zlsx_emb_coverage_count`, `zlsx_emb_coverage_id`, `zlsx_emb_coverage_sheet`, `zlsx_emb_coverage_range`, `zlsx_emb_coverage_rows`, `zlsx_emb_vectors`, `zlsx_emb_hashes`, `zlsx_emb_digest`, `zlsx_emb_carrier`, `zlsx_emb_tombstone` | ✓ `zlsx.embeddings` | — S3c ⁵ | S3c |
| List embeddable rows | ✓ `Workbook.embeddableRows` | — S3c | — S3c | ✓ `--extract` | S3c |
| Write vectors (+ tombstones for uncovered rows) | ✓ `Workbook.setEmbeddings`, `Workbook.setEmbeddingsOpts` | — S3c | — S3c | ✓ `--vectors` | S3c |
| Recovery record in cells (`recovery_in_cells`, Numbers-durable) | ✓ `Workbook.setEmbeddingsOpts` | — S3c | — S3c | — S3c | S3c |
| Prune (tombstone stale slots, zero vectors) | ✓ `Workbook.pruneEmbeddings` | — S3c | — S3c | ✓ `--prune` | S3c |
| Strip (parts + recovery record) | ✓ `Workbook.stripEmbeddings` | — S3c | — S3c | ✓ `--strip` | S3c |

## 5 · Safety and limits

| Capability | Zig | C | Py | CLI | Row |
|---|---|---|---|---|---|
| Archive defenses on the package path — per-part cap, ratio cap, Zip64 / split / encrypted refused | ✓ `PartStore.open`, `PartStore.openBuffer` ¹⁵ | ✓ `zlsx_editor_open`, `zlsx_open_buffer` ¹⁶ | ✓ `zlsx.edit`, `Editor.from_bytes` ¹⁶ | ✓ `append-rows` ¹⁶ | |
| Decompression caps on the core reader (`Book.open`) — per-part cap, ratio cap | ✓ `Book.open`, `Book.openBuffer`, `zlsx.extractEntryToBuffer` | ✓ `zlsx_book_open`, `zlsx_book_open_buffer` | ✓ `zlsx.open`, `zlsx.open_bytes` | ✓ `rows` ¹⁷ | |
| Aggregate decompression budget (whole archive) | ✓ `zlsx.decompress_limits`, `DecompressBudget.admit` ¹⁸ | ✓ `zlsx_book_open`, `zlsx_editor_open` | ✓ `zlsx.open`, `zlsx.edit` | ✓ `rows`, `append-rows` ¹⁷ | |

¹⁵ `PartStore` admits every central-directory entry against `zlsx_control.decompress_limits` on the scan (`scanCentralDirectory`) — the per-part cap, the ratio cap and the whole-archive aggregate, before any part is inflated — and re-checks the per-part half in `decompressPayload` before allocating that part's output; the file-backed open reads the whole archive into a scratch buffer first, so the limits bound decompression, not the archive read. Data-descriptor entries are accepted (sizes come from the central directory). `Workbook.open` goes through `PartStore` and inherits this.
¹⁶ `Editor.open` — the path under `zlsx_editor_open`, `zlsx.edit` and every CLI edit sub-command — admits every entry on its own structural scan (Zip64 / split / encrypted / data-descriptor refused on the same pass), *then* opens the core reader (`Book.open` / `Book.openBuffer`), which admits every entry again on its own walk; the archive is refused before any part is inflated whichever layer sees it first, and the editor's later `readEntry` decompressions re-check the per-part half. On C the refusal is `ZLSX_ERROR` with `ZipBombSuspected` in the error buffer; on Python it is `ZlsxError("ZipBombSuspected")`; every CLI edit sub-command exits 4.
¹⁷ Exit code 4 (`ZipBombSuspected`) on the read family and the edit family — `src/cli.zig::openFailureExit`. `eval` / `recalc` keep their own table (`docs/cli.md`) and report a breach at open as 2.
¹⁸ The three numbers live in `pkg/control.zig` (`decompress_limits`: 512 MiB per part, 4096:1, 2 GiB aggregate — the S1 owner gate's to confirm or move) and are re-exported as `zlsx.decompress_limits` and `zlsx_pkg.decompress_limits`. `DecompressBudget.admit` is the one implementation of the arithmetic all three openers run.

## 6 · Cross-cutting properties

Not per-surface capabilities — they hold for the code every surface shares, and are listed so the README's claims about them have a checkable home.

| Property | Where it holds | Where it does not |
|---|---|---|
| Control-byte rejection on user-text channels | Sheet names, cell text, comments, defined names, hyperlinks — the writer and editor input paths. | Embedding metadata: `Workbook.setEmbeddingsOpts` writes the `model` / coverage id text through `appendXmlEscaped` (`pkg/embedding_part.zig`), which escapes `& < > "` only — a control byte in a model id reaches the part. Closed with S3c. |
| Coverage-guided fuzz binaries (`zig build fuzz --fuzz`; each a `-ffuzz` `addTest` in `build.zig`) | `src/xlsx.zig` (`unit_fuzz_tests`), `pkg/store.zig` (`package_fuzz_tests`), the byte walkers `pkg/{sheet_edit,table_edit,drawing_edit,vml_edit}.zig` (`walker_fuzz`), and the formula engine `src/formula/{tokenizer,numfmt,parser,value,spill,resolved,criteria,metadata,calc,names,decode,eval}.zig`. Nightly on `macos-14` under 0.16.0 (the Linux build runner panics collecting coverage). | `pkg/workbook.zig` has `std.testing.fuzz` blocks but is wired only as an ordinary test module (smoke, not coverage-guided); `pkg/drawings.zig` (anchor parsers) and `pkg/typed_parts/*` have no fuzz target at all — ordinary `Workbook` and corpus coverage only. |
| Byte-preservation of untouched parts on save | Every save path (`PartStore.save`, `Editor.save` passthrough). | — |

## 7 · Surfaces that belong to one platform

| Capability | Zig | C | Py | CLI |
|---|---|---|---|---|
| PySpark Data Source — batch read/write, streaming source | n/a | n/a | ✓ `zlsx.spark` | n/a |
| Unity Catalog Python UDF pattern | n/a | n/a | ✓ `integrations/databricks/` | n/a |
| Databricks transfer + governance | n/a | n/a | n/a | ✓ `dbx` |
| Raw OPC substrate — `PartStore` part access, `Workbook.empty` → `Workbook.saveFreshEmit` | ✓ `zlsx_pkg.PartStore`, `Workbook.empty`, `Workbook.saveFreshEmit` | n/a | n/a | n/a |

---

## Gaps by sigmoid row

Derived from the tables above; the ladder's **Row** column is the join key.
S3d, S3e, S5b and S11 were added to the ladder, and S3b widened, at the S0
gate (2026-08-26).

| Row | Cells it closes |
|---|---|
| S2 | **Done 2026-08-26** — §3 `<xm:f>` row: the refusal became `Workbook.rewriteAllExtensionFormulas` + `Workbook.preflightExtensionFormulas` in the shared code (footnote ¹⁵); C and Py *observe* the lift only once S3a gives them structural edits — the row's four-way `done` is transitive through S3a. |
| S3a | §3 sheet-level edits, row / col edits, table-column rename and the rewriters → C + Py; table-column rename → CLI. |
| S3b | The ladder text says merged ranges, defined names, conditional formats, anchors → **CLI**; §1 adds document properties to that CLI list. §1 also shows defined names, conditional formats, anchors, panes / dimension / calc properties, sheet visibility, formula text and error tags missing on **C and Py** (merged ranges and document properties are already there). Widened at the gate: one typed-read parity row covering all four surfaces. |
| S3c | §4 embeddable rows, write, prune, strip → C + Py; `recovery_in_cells` → C + Py + CLI; the vector / state dump → CLI; §6 control-byte check on embedding metadata (all surfaces, shared code). |
| S3d | §3 existing-workbook authoring (styles, layout, merges, hyperlinks, comments, DV / CF, defined names, `deleteCell`) and a standalone mark-recalc → C + Py + CLI. Zig-only today; no row names it. |
| S3e | §1 opening strategies — lazy per-sheet loading and the lazy SST backend → C + Py (+ CLI for per-sheet). |
| S4 | §1 indexed palette + tint, all four. |
| S5 | §2 Writer images; §3 `Workbook.addImage*` → C + Py + CLI; §2 first row — export `src/writer.zig` as a public module (footnote ⁷), since S5 is the row that reshapes the Writer API anyway. The CLI leg presupposes a fresh-workbook path — see S11. |
| S5b | §3 typed object extract / replace / remove (images, charts) on **all four** surfaces — Zig has raw-part composition only (footnote ¹¹) — retiring the `zlsx-extract-images` sibling into `zlsx`. C2a's product promise, distinct from the raw OPC substrate. |
| S6 | §1 typed pivot read — Zig + CLI shipped; the C and Python legs follow the S6 gate. Footnote ¹⁰'s audit answered (2026-08-27). |
| S7a | **Done 2026-08-28** — §3 the host-only pivot row: `location@ref` moves under every row / col edit (footnote ¹⁰); C and Py observe it through S3a. |
| S7b · S7c | §3 the remaining `PivotEditUnsafe` lifts — source sheets, then cache fields — staged. |
| S8 | §2 pivot authoring, fresh workbooks. CLI leg presupposes S11. |
| S9 | §2 charts, fresh workbooks. CLI leg presupposes S11. |
| S11 | §2 the CLI column — fresh-workbook authoring from a workbook spec (`zlsx write`). In the spine as a dependency of the CLI legs of S5, S8 and S9 and ahead of S10 (the id is a label, not a position). |
| S10 | Every `—` left is closed or carries a ruling below; plus the closing-sweep items parked here — engine-fingerprint accessor (Zig, CLI) and hidden sheets on write (Zig, C, Py). |

## Rulings

**Owner rulings recorded at the S0 gate (2026-08-26)** — all six proposals
accepted as recommended; the CLI fresh-authoring question became row S11
instead of a ruling. Accepting a ruling amends the
owner-locked strict four-way parity rule **for that cell only** — a row whose
remaining `—` cells are all `n/a` counts as `done`; nothing else about the
rule changes.

| Surface | Permanently absent | Why | Ruling |
|---|---|---|---|
| CLI | Fresh-workbook authoring — the whole §2 `Writer` family. | The CLI's authoring grammar is NDJSON scalars over an *existing* workbook; styling and rule authoring need a workbook-spec schema. Ruling it `n/a` contradicts the owner-locked strict four-way parity, and S5 / S8 / S9 plan CLI image, pivot and chart authoring — which have nowhere to land without a fresh-workbook path. | Not `n/a` — **row S11** (`zlsx write` from a JSON workbook spec). |
| CLI | In-memory bytes I/O (`open_bytes` / `to_bytes` / `save_to_buffer` / `from_bytes`). | Files are the CLI's memory; buffers exist to serve an in-process boundary. Streaming a workbook through stdin / stdout would be a feature, not parity. | Accepted — `n/a` |
| CLI | Bulk-FFI matrix handle. | `rows` already streams a sheet in one pass; the handle amortises per-call FFI dispatch, which a process boundary has none of. | Accepted — `n/a` |
| CLI | Date → Excel serial encoding; hidden sheets on write. | Library-side helpers of the fresh writer; the CLI cells follow S11. Hidden sheets on write are missing on Zig, C and Py as well — two-line additions with no design. | Not `n/a` — CLI cells ride S11; the Zig / C / Py hidden-sheet cells ride S10's closing sweep. |
| Zig · C · Py | The `dbx` family. | Network transfer and governance live in the static binary by design; Python reaches Databricks through the Spark Data Source and the UDF pattern. | Accepted — `n/a` |
| Zig · C · CLI | PySpark Data Source, UC UDF pattern. | Spark is a Python runtime. | Accepted — `n/a` |
| C · Py · CLI | Raw OPC substrate — `PartStore` part access and the package-layer fresh emit. | Byte-level part access is what the product surfaces are built on, not a product surface. The *typed* object operations C2a promised are **not** covered by this ruling — they are S5b. | Accepted — `n/a` |
| Zig · CLI | Engine fingerprint accessor. | Identity metadata for the engine; today only the C ABI and Python expose it. Two-line additions, no design. | Not `n/a` — rides S10's closing sweep. |

## Updating this file

1. Land the capability. In the same PR, change the cell from `— Sx` to
   `✓ \`sym\`` (or `~ \`sym\`` + a footnote naming the remaining gap), and run
   `python3 scripts/check_surface_matrix.py` — it fails on any symbol that
   does not exist on its owner, any cell without a symbol, any row id that is
   not in the ladder (a row proposed here before it exists in
   `goal_sigmoid.md` is marked *(proposed)* until the owner gate adopts it).
2. Remove the row's line from *Gaps by sigmoid row* when it empties.
3. Never write `n/a` without a line in *Rulings*; never add a ruling without
   the owner's answer recorded in `goal_sigmoid.md` §5.
4. Leave the *Last audited* line alone unless the PR re-verifies every cell
   (S0 and S10 do; capability PRs do not).
