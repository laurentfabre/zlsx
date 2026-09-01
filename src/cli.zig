//! `zlsx` — command-line interface over the zlsx library.
//!
//! Streams rows of the selected sheet to stdout in one of four formats;
//! the edit/embed families mutate through the package layer; `eval` /
//! `recalc` (M6, `formula_cli.zig`) put the formula engine on the
//! command line. Designed as a drop-in openpyxl replacement: shell-
//! friendly, pipeable into jq / awk, no Python interpreter floor.

const std = @import("std");
const fuzz_config = @import("fuzz_config");
const builtin = @import("builtin");
const xlsx = @import("zlsx");
const zlsx_pkg = @import("zlsx_pkg");
const dbx = @import("dbx.zig");
const formula_cli = @import("formula_cli.zig");
const coords = @import("zlsx_refs");

const Format = enum {
    /// NEW default: row envelope `{kind,sheet,sheet_idx,row,cells:[…]}`.
    jsonl,
    /// Bare `[…]` arrays — what iter54's `jsonl` emitted.
    legacy_jsonl,
    /// Bare `{col:val,…}` objects — what iter54's `jsonl-dict` emitted.
    legacy_jsonl_dict,
    tsv,
    csv,
};

/// iter56/57/58: first positional decides sub-command. `rows` is the
/// legacy envelope-row emitter; `cells` is the per-cell NDJSON stream;
/// `meta` emits a workbook record followed by per-sheet records;
/// `list_sheets` is the lighter NDJSON variant. iter58 adds the
/// five-way reader-surface exposure: `comments` / `validations` /
/// `hyperlinks` iterate every sheet (sheet-scoped records); `styles`
/// / `sst` are workbook-wide. Bare `zlsx file.xlsx` (no sub-command
/// token) still means `rows` so existing scripts keep working — the
/// short-alias re-point to `cells` is an iter60+ breaking change with
/// its own rollout.
const Subcommand = enum {
    rows,
    cells,
    meta,
    list_sheets,
    comments,
    validations,
    hyperlinks,
    /// S6: typed pivot read — one record per pivot table (host-sheet
    /// order), then one per cache no table reads. Package-layer
    /// route (`Workbook.pivotTables`), dispatched before the Book open.
    pivots,
    styles,
    sst,
    /// S3b: merged ranges — one record per `<mergeCell>` across the
    /// selected sheets. Book route; the validations / hyperlinks
    /// family (range-keyed, every sheet by default).
    merges,
    /// S3b: defined names — one record per `<definedName>` of
    /// `xl/workbook.xml`, document order. Package-layer route like
    /// `pivots` (the reader-only Book has no workbook.xml view);
    /// a concrete sheet selector narrows to the names SCOPED to that
    /// sheet and suppresses workbook-scope names.
    defined_names,
    /// S3b: document properties — one `{"kind":"doc_props",…}` record
    /// carrying the `docProps/core.xml` + `app.xml` field set and the
    /// `custom.xml` presence flag (the `Editor.doc_props` field set).
    /// Package-layer route like `pivots`; no sheet dimension, so it
    /// joins `meta`'s workbook-scoped flag tolerance.
    doc_props,
    /// S3b slice 4: drawing anchors — one record per anchored image or
    /// chart across the selected sheets (`pkg/anchor_ndjson.zig` over
    /// `zlsx_pkg.imageAnchors` + `zlsx_pkg.chartAnchors`).
    /// Package-layer route like `pivots`: the walkers read drawing
    /// parts the reader-only Book has no view of.
    anchors,
    /// iter-lms-4 follow-up: append rows from stdin (NDJSON, one
    /// JSON array per row) to a sheet of an existing xlsx and save
    /// to `--out`. Requires `--sheet N` and `--out PATH`.
    append_rows,
    /// iter-cm-4: rewrite a single cell in place via `Editor.setCell`
    /// and save to `--out`. Requires `--sheet N`, `--ref A1`, and
    /// `--value <JSON>`. Value is a single JSON token (string, int,
    /// float, bool, null) — same shape as one slot of an
    /// `append-rows` row.
    set_cell,
    /// iter-row-4 / iter-col-5 / iter-sheet-5 — structural-edit
    /// CLI sub-commands. Each opens via Editor, applies one
    /// structural mutation, and saves to `--out`.
    insert_row,
    delete_row,
    insert_column,
    delete_column,
    add_sheet,
    rename_sheet,
    delete_sheet,
    /// S3a: rename a column of a named table — `<tableColumn name>`
    /// and every structured reference workbook-wide follow
    /// (`Editor.renameTableColumn`). Requires `--table NAME`,
    /// `--old-name OLD`, `--new-name NEW` and `--out PATH`.
    rename_table_column,
    /// Z3: strip identifying document metadata (docProps/core.xml,
    /// docProps/app.xml, docProps/custom.xml) and save to `--out`.
    /// Cell data is untouched — this is the metadata counterpart to
    /// masking cell values.
    scrub_metadata,
    /// emb-6a: embedding maintenance. `--strip` removes the embedding
    /// parts *and* the recovery record, so the result reports `absent`
    /// rather than `stripped` — the pre-share operation from the
    /// design's Caveats. Requires `--out PATH`.
    embed,
};

/// iter60b: `--output` wire-shape switch.
/// - `ndjson` (default) — the invariant envelope NDJSON stream iter55+
///   already emits. Every record carries its own `sheet`/`sheet_idx`.
/// - `compact_ndjson` — per-sheet prologue record `{"kind":"sheet",…}`
///   emitted once per sheet; subsequent per-record lines OMIT
///   `sheet`/`sheet_idx`. Sheet-scoped sub-commands only (cells / rows /
///   comments / validations / hyperlinks). On workbook-scoped commands
///   it's implemented for consistency but is effectively a no-op:
///   `meta` just drops the `sheet`/`sheet_idx` fields from its sheet
///   records; `list-sheets` is literally identical to ndjson (it
///   already emits sheet prologues only); `styles` / `sst` have no
///   `sheet` field at all.
/// - `pretty_json` — meta-only. Collapses the workbook + sheets
///   records into one 2-space-indented JSON object. Every other
///   sub-command rejects this mode at parse time.
const OutputMode = enum { ndjson, compact_ndjson, pretty_json };

const Args = struct {
    subcommand: Subcommand = .rows,
    file: []const u8,
    sheet_index: ?usize = null,
    sheet_name: ?[]const u8 = null,
    format: Format = .jsonl,
    list_sheets: bool = false,
    /// Set when the user passed the deprecated `--format jsonl-dict`
    /// spelling. `main` emits a one-line stderr deprecation warning
    /// so existing scripts keep working while their authors learn
    /// about the rename.
    deprecated_jsonl_dict: bool = false,
    /// iter59a: stream-native pagination over the emitted-record
    /// stream (rows / cells / comments / validations / hyperlinks /
    /// styles / sst). Both are applied GLOBALLY after sheet selection.
    skip: ?usize = null,
    take: ?usize = null,
    /// iter59b-1: per-sheet row-bounded filtering on the three
    /// sub-commands that emit row-keyed records (rows / cells /
    /// comments). Both endpoints are 1-based OOXML row numbers and
    /// inclusive: `start_row=3, end_row=5` emits rows 3, 4, 5.
    /// Applied BEFORE --skip/--take, so --skip counts post-row-filter
    /// records per the jq-for-excel design doc.
    start_row: ?u32 = null,
    end_row: ?u32 = null,
    /// iter59b-2: A1-style bounding-rectangle filter (`--range A1:Z100`).
    /// Populated only on `rows` and `cells`; rejected elsewhere.
    /// Stored with `top_left ≤ bottom_right` on both axes; the CLI
    /// parser rejects inverted corners rather than silently swapping
    /// (differs from `xlsx.parseA1Range` which normalises silently).
    /// When paired with --start-row / --end-row, the row bounds are
    /// intersected (most restrictive wins) at the filter site.
    range: ?xlsx.MergeRange = null,
    /// iter59b-3: promote the first emitted row to header keys on the
    /// `rows --format jsonl` path. Header row is consumed silently;
    /// subsequent rows emit `{…,"fields":{key:val,…}}` instead of
    /// `{…,"cells":[…]}`. Rejected for every other sub-command and
    /// every non-default format — see parseArgs for scoping rules.
    header: bool = false,
    /// iter59b-4: on `cells` / `rows --format jsonl`, emit records for
    /// empty cells using the `t:"blank","v":null` shape instead of
    /// skipping them. On the `rows --header` dict path the flag is
    /// a no-op (the dict already emits `key:null` for missing cells)
    /// but accepted silently so scripts can set it unconditionally.
    /// On legacy flat formats (csv / tsv / legacy-jsonl / legacy-jsonl-
    /// dict) the flag is accepted but shape-neutral — those formats
    /// already serialise empties per their own convention. Rejected on
    /// every other sub-command; see parseArgs for the scoping matrix.
    include_blanks: bool = false,
    /// iter59b-4: on `cells` / `rows --format jsonl` (envelope only),
    /// attach a terse per-cell `style:{…}` object when the cell's
    /// style index resolves to an effective format (any of bold /
    /// italic / fg / bg / non-General num_fmt / any border side).
    /// Cells with no effective style OMIT the field entirely.
    /// Rejected on `rows --header` (the fields dict has no place for
    /// per-cell metadata) and on non-jsonl formats (csv/tsv/legacy
    /// shapes don't accommodate nested records). Rejected on every
    /// other sub-command — those have their own style exposure via
    /// the `styles` sub-command.
    with_styles: bool = false,
    /// iter59c: expand sheet selection to every sheet in the workbook.
    /// Mutually exclusive with `--sheet` / `--name` / `--sheet-glob`.
    /// Ignored on workbook-scoped sub-commands (same tolerance group
    /// as the other sheet-selector flags).
    all_sheets: bool = false,
    /// iter59c: simple-glob pattern (`*` any run, `?` single char,
    /// case-sensitive). Selects every sheet whose name matches. Mutually
    /// exclusive with `--sheet` / `--name` / `--all-sheets`. Ignored on
    /// workbook-scoped sub-commands.
    sheet_glob: ?[]const u8 = null,
    /// iter60b: wire-shape switch. See OutputMode doc for per-mode
    /// semantics and the sub-command scoping matrix.
    output: OutputMode = .ndjson,
    /// iter-lms-4 follow-up: target file for the `append-rows`
    /// sub-command. Required for `append-rows`, ignored / rejected
    /// elsewhere. Set via `--out PATH`.
    out_path: ?[]const u8 = null,
    /// emb-6a: `embed --strip` — remove the embedding parts and the
    /// recovery record. Rejected on every other sub-command.
    strip: bool = false,
    /// emb-6b: `embed --prune` — tombstone slots whose row is no longer
    /// embeddable. Rejected on every other sub-command.
    prune: bool = false,
    /// emb-6c: `embed --extract` — emit the rows that need embedding as
    /// NDJSON on stdout. Phase one of the out-of-band write path.
    extract: bool = false,
    /// emb-6c: `--vectors PATH` — NDJSON vectors to write back. Phase
    /// two.
    vectors_path: ?[]const u8 = null,
    /// emb-6c: `--model NAME` — provenance recorded in the index.
    model_name: ?[]const u8 = null,
    /// emb-6c: `--column A` — the column letters being embedded.
    /// Distinct from `--col`, which is a numeric index for column edits.
    column_name: ?[]const u8 = null,
    /// emb-6c: `--coverage A2:A100` — the covered A1 range. A dedicated
    /// flag rather than `--range`, which parses to a rectangle for the
    /// reader sub-commands and is rejected on `embed`.
    coverage_range: ?[]const u8 = null,
    /// emb-6c: `--id NAME` — `<coverage id>`; defaults to "default".
    coverage_id: ?[]const u8 = null,
    /// emb-6c: `--dtype f32|int8-sym` — on-disk vector encoding.
    dtype_name: ?[]const u8 = null,
    /// iter-cm-4: A1-style cell ref for the `set-cell` sub-command.
    /// Set via `--ref A1`.
    cell_ref: ?[]const u8 = null,
    /// iter-cm-4: JSON-encoded scalar value for `set-cell`. Same
    /// shape as one cell in an `append-rows` row. Set via
    /// `--value <JSON>`.
    cell_value_json: ?[]const u8 = null,
    /// iter-row-4 / iter-col-5: 1-based row index for `insert-row` /
    /// `delete-row`. Set via `--row N`.
    row_1based: ?u32 = null,
    /// iter-col-5: column letter (A, B, …, XFD) for `insert-column` /
    /// `delete-column`. Set via `--col LETTER`. Decoded into a 1-based
    /// column index at runtime.
    col_letter: ?[]const u8 = null,
    /// iter-sheet-4 / iter-sheet-5: target sheet name for
    /// `add-sheet` / `rename-sheet`. Set via `--name NAME`.
    /// Distinct from `sheet_name` (the sheet selector) — name conflict
    /// is resolved at runtime per-subcommand.
    new_sheet_name: ?[]const u8 = null,
    /// S3a: the table and the column `rename-table-column` renames.
    /// Set via `--table NAME` / `--old-name OLD`; the new name rides
    /// `--new-name`.
    table_name: ?[]const u8 = null,
    old_column_name: ?[]const u8 = null,
    /// iter-sst-4: opt into the lazy SST backend (`Book.openSstLazy`).
    /// Workbooks with millions of unique strings skip the eager
    /// decode arena; resolution happens on first cell access. See
    /// docs/plans/archive/streaming-sst.md for the trade-off (sparse access
    /// wins; full sweeps cost slightly more than eager). Accepted on
    /// every sub-command for wrapper-friendly setting; ignored by
    /// `meta` / `list-sheets` / `styles` / `sst` (those are
    /// workbook-scoped and don't benefit from per-cell laziness).
    sst_lazy: bool = false,
};

const ArgError = error{
    NoFile,
    HelpRequested,
    UnknownFlag,
    MissingValue,
    BadFormat,
    BadSheetIndex,
    BadArgValue,
    SheetArgConflict,
    TooManyArgs,
};

/// First-pass scan: identify the sub-command without validating
/// flag values. Lets the main pass relax --sheet / --name / --format
/// validation for workbook-scoped sub-commands that wrappers may
/// append those flags to universally. Skips `--sheet` / `--name` /
/// `--format` pairs so their values aren't mistaken for positionals.
fn detectSubcommand(argv: []const []const u8) Subcommand {
    var i: usize = 0;
    while (i < argv.len) : (i += 1) {
        const a = argv[i];
        if (std.mem.eql(u8, a, "--sheet") or
            std.mem.eql(u8, a, "--name") or
            std.mem.eql(u8, a, "--format") or
            std.mem.eql(u8, a, "--skip") or
            std.mem.eql(u8, a, "--take") or
            std.mem.eql(u8, a, "--start-row") or
            std.mem.eql(u8, a, "--end-row") or
            std.mem.eql(u8, a, "--range") or
            std.mem.eql(u8, a, "--sheet-glob") or
            std.mem.eql(u8, a, "--output") or
            std.mem.eql(u8, a, "--out") or
            std.mem.eql(u8, a, "--ref") or
            std.mem.eql(u8, a, "--value") or
            std.mem.eql(u8, a, "--row") or
            std.mem.eql(u8, a, "--col") or
            std.mem.eql(u8, a, "--new-name") or
            std.mem.eql(u8, a, "--table") or
            std.mem.eql(u8, a, "--old-name") or
            std.mem.eql(u8, a, "--vectors") or
            std.mem.eql(u8, a, "--model") or
            std.mem.eql(u8, a, "--column") or
            std.mem.eql(u8, a, "--coverage") or
            std.mem.eql(u8, a, "--id") or
            std.mem.eql(u8, a, "--dtype"))
        {
            i += 1; // skip paired value (bounds-checked by caller)
            continue;
        }
        if (a.len > 0 and a[0] == '-') continue; // flag with no value
        if (std.mem.eql(u8, a, "cells")) return .cells;
        if (std.mem.eql(u8, a, "rows")) return .rows;
        if (std.mem.eql(u8, a, "append-rows")) return .append_rows;
        if (std.mem.eql(u8, a, "set-cell")) return .set_cell;
        if (std.mem.eql(u8, a, "insert-row")) return .insert_row;
        if (std.mem.eql(u8, a, "delete-row")) return .delete_row;
        if (std.mem.eql(u8, a, "insert-column")) return .insert_column;
        if (std.mem.eql(u8, a, "delete-column")) return .delete_column;
        if (std.mem.eql(u8, a, "add-sheet")) return .add_sheet;
        if (std.mem.eql(u8, a, "rename-sheet")) return .rename_sheet;
        if (std.mem.eql(u8, a, "delete-sheet")) return .delete_sheet;
        if (std.mem.eql(u8, a, "rename-table-column")) return .rename_table_column;
        if (std.mem.eql(u8, a, "meta")) return .meta;
        if (std.mem.eql(u8, a, "list-sheets")) return .list_sheets;
        if (std.mem.eql(u8, a, "scrub-metadata")) return .scrub_metadata;
        if (std.mem.eql(u8, a, "embed")) return .embed;
        if (std.mem.eql(u8, a, "comments")) return .comments;
        if (std.mem.eql(u8, a, "validations")) return .validations;
        if (std.mem.eql(u8, a, "hyperlinks")) return .hyperlinks;
        if (std.mem.eql(u8, a, "pivots")) return .pivots;
        if (std.mem.eql(u8, a, "styles")) return .styles;
        if (std.mem.eql(u8, a, "sst")) return .sst;
        if (std.mem.eql(u8, a, "merges")) return .merges;
        if (std.mem.eql(u8, a, "defined-names")) return .defined_names;
        if (std.mem.eql(u8, a, "doc-props")) return .doc_props;
        if (std.mem.eql(u8, a, "anchors")) return .anchors;
        return .rows; // first positional is the file path
    }
    return .rows;
}

fn parseArgs(raw_argv: []const []const u8) ArgError!Args {
    // Pre-normalize `--key=value` to `[--key, value]` token pairs so
    // the rest of the parser can use its existing two-token form
    // uniformly. The buffer is stack-allocated with a generous
    // bound (256 tokens — real invocations are well under 50);
    // overflow returns TooManyArgs so the caller sees a typed
    // error instead of an OOB write.
    var split_buf: [256][]const u8 = undefined;
    var split_count: usize = 0;
    // Boolean flags that don't consume a value. `--bool=anything`
    // is invalid syntax and must be rejected — silently splitting
    // it would let the value leak into the positional slot and
    // pick the wrong file path.
    const boolean_flags = [_][]const u8{
        "--list-sheets", "--header",     "--include-blanks", "--with-styles",
        "--sst-lazy",    "--all-sheets", "--help",           "--strip",
        "--prune",       "--extract",
    };
    // Value-bearing flags. `--key=value` is split into [--key, value]
    // ONLY when key is one of these — otherwise the token is left
    // verbatim, so an arbitrary `--Q=1` value passed via the
    // two-token form (`--name --Q=1`) is consumed as a literal value
    // by the preceding flag rather than misparsed as a new flag.
    const value_flags = [_][]const u8{
        "--sheet",      "--name",      "--format",  "--skip",
        "--take",       "--start-row", "--end-row", "--range",
        "--sheet-glob", "--output",    "--out",     "--ref",
        "--value",      "--row",       "--col",     "--new-name",
        "--vectors",    "--model",     "--column",  "--coverage",
        "--id",         "--dtype",     "--table",   "--old-name",
    };
    // Context-aware splitter: the token IMMEDIATELY following a
    // value-bearing flag is its literal value and must pass through
    // verbatim, even if it textually matches `--known-flag=value`.
    // Without this, `zlsx rows f.xlsx --name --format=csv` would
    // split the literal sheet name into a separate flag-token pair.
    var prev_was_value_flag: bool = false;
    for (raw_argv) |raw| {
        if (prev_was_value_flag) {
            // This token is a value for the previous flag — pass
            // verbatim regardless of shape.
            if (split_count >= split_buf.len) return ArgError.TooManyArgs;
            split_buf[split_count] = raw;
            split_count += 1;
            prev_was_value_flag = false;
            continue;
        }
        if (raw.len >= 2 and raw[0] == '-' and raw[1] == '-') {
            if (std.mem.indexOfScalar(u8, raw, '=')) |eq| {
                const key = raw[0..eq];
                for (boolean_flags) |bf| {
                    if (std.mem.eql(u8, key, bf)) return ArgError.BadArgValue;
                }
                var is_value_flag = false;
                for (value_flags) |vf| {
                    if (std.mem.eql(u8, key, vf)) {
                        is_value_flag = true;
                        break;
                    }
                }
                if (is_value_flag) {
                    if (split_count + 2 > split_buf.len) return ArgError.TooManyArgs;
                    split_buf[split_count] = key;
                    split_buf[split_count + 1] = raw[eq + 1 ..];
                    split_count += 2;
                    // The split form already supplied the value for
                    // this flag; the next token is unrelated.
                    prev_was_value_flag = false;
                    continue;
                }
                // Unknown `--key=…` token — fall through to verbatim.
            }
            // Bare `--flag` form: if it's value-bearing, the next
            // token is its literal value.
            for (value_flags) |vf| {
                if (std.mem.eql(u8, raw, vf)) {
                    prev_was_value_flag = true;
                    break;
                }
            }
        }
        if (split_count >= split_buf.len) return ArgError.TooManyArgs;
        split_buf[split_count] = raw;
        split_count += 1;
    }
    const argv = split_buf[0..split_count];

    const detected_sub = detectSubcommand(argv);
    // Workbook-scoped commands don't consume --sheet / --name /
    // --format, so wrappers that always append those flags should
    // not hit a hard error. Parse them tolerantly: missing-value is
    // still an error (user typo), but a malformed value is silently
    // dropped. Non-workbook commands keep strict validation.
    //
    // iter58: the three sheet-scoped newcomers (`comments` /
    // `validations` / `hyperlinks`) iterate every sheet by default,
    // so they join this group for flag tolerance even though their
    // records do carry `sheet` / `sheet_idx`. Narrowing via `--sheet`
    // is deferred to iter58-follow-up.
    const workbook_scoped = switch (detected_sub) {
        .meta,
        .list_sheets,
        .styles,
        .sst,
        // S3b: doc-props has no sheet dimension at all — sheet and
        // format flags are tolerated and dropped, the meta family's
        // wrapper-friendliness.
        .doc_props,
        => true,
        .rows, .cells, .comments, .validations, .hyperlinks, .pivots, .merges, .defined_names, .anchors, .append_rows, .set_cell, .insert_row, .delete_row, .insert_column, .delete_column, .add_sheet, .rename_sheet, .delete_sheet, .rename_table_column, .scrub_metadata, .embed => false,
    };

    var out: Args = .{ .file = "", .subcommand = detected_sub };
    var first_positional_seen = false;
    var i: usize = 0;
    while (i < argv.len) : (i += 1) {
        const a = argv[i];
        if (std.mem.eql(u8, a, "-h") or std.mem.eql(u8, a, "--help")) {
            return ArgError.HelpRequested;
        } else if (std.mem.eql(u8, a, "--list-sheets")) {
            out.list_sheets = true;
        } else if (std.mem.eql(u8, a, "--header")) {
            // Boolean flag — no value consumed. Scoping checked below.
            out.header = true;
        } else if (std.mem.eql(u8, a, "--include-blanks")) {
            out.include_blanks = true;
        } else if (std.mem.eql(u8, a, "--with-styles")) {
            out.with_styles = true;
        } else if (std.mem.eql(u8, a, "--sst-lazy")) {
            out.sst_lazy = true;
        } else if (std.mem.eql(u8, a, "--out")) {
            i += 1;
            if (i >= argv.len) return ArgError.MissingValue;
            out.out_path = argv[i];
        } else if (std.mem.eql(u8, a, "--ref")) {
            i += 1;
            if (i >= argv.len) return ArgError.MissingValue;
            out.cell_ref = argv[i];
        } else if (std.mem.eql(u8, a, "--value")) {
            i += 1;
            if (i >= argv.len) return ArgError.MissingValue;
            out.cell_value_json = argv[i];
        } else if (std.mem.eql(u8, a, "--row")) {
            i += 1;
            if (i >= argv.len) return ArgError.MissingValue;
            const v = std.fmt.parseInt(u32, argv[i], 10) catch return ArgError.BadArgValue;
            // 1-based per the structural-edit contract; mirror
            // --start-row / --end-row's parse-time zero rejection.
            if (v == 0) return ArgError.BadArgValue;
            out.row_1based = v;
        } else if (std.mem.eql(u8, a, "--col")) {
            i += 1;
            if (i >= argv.len) return ArgError.MissingValue;
            out.col_letter = argv[i];
        } else if (std.mem.eql(u8, a, "--new-name")) {
            i += 1;
            if (i >= argv.len) return ArgError.MissingValue;
            out.new_sheet_name = argv[i];
        } else if (std.mem.eql(u8, a, "--table")) {
            i += 1;
            if (i >= argv.len) return ArgError.MissingValue;
            out.table_name = argv[i];
        } else if (std.mem.eql(u8, a, "--old-name")) {
            i += 1;
            if (i >= argv.len) return ArgError.MissingValue;
            out.old_column_name = argv[i];
        } else if (std.mem.eql(u8, a, "--sheet")) {
            if (!workbook_scoped and (out.sheet_name != null or out.all_sheets or out.sheet_glob != null))
                return ArgError.SheetArgConflict;
            i += 1;
            if (i >= argv.len) return ArgError.MissingValue;
            const parsed = std.fmt.parseInt(usize, argv[i], 10) catch {
                if (workbook_scoped) continue; // ignore bad value for meta/list-sheets
                return ArgError.BadSheetIndex;
            };
            out.sheet_index = parsed;
        } else if (std.mem.eql(u8, a, "--name")) {
            if (!workbook_scoped and (out.sheet_index != null or out.all_sheets or out.sheet_glob != null))
                return ArgError.SheetArgConflict;
            i += 1;
            if (i >= argv.len) return ArgError.MissingValue;
            out.sheet_name = argv[i];
        } else if (std.mem.eql(u8, a, "--strip")) {
            // emb-6a: no value. Only `embed` acts on it; anywhere else
            // it is a typo for something, and silently ignoring it on a
            // destructive-sounding flag would be the wrong tolerance.
            if (out.subcommand != .embed) return ArgError.UnknownFlag;
            out.strip = true;
        } else if (std.mem.eql(u8, a, "--prune")) {
            if (out.subcommand != .embed) return ArgError.UnknownFlag;
            out.prune = true;
        } else if (std.mem.eql(u8, a, "--extract")) {
            if (out.subcommand != .embed) return ArgError.UnknownFlag;
            out.extract = true;
        } else if (std.mem.eql(u8, a, "--vectors")) {
            if (out.subcommand != .embed) return ArgError.UnknownFlag;
            i += 1;
            if (i >= argv.len) return ArgError.MissingValue;
            out.vectors_path = argv[i];
        } else if (std.mem.eql(u8, a, "--model")) {
            if (out.subcommand != .embed) return ArgError.UnknownFlag;
            i += 1;
            if (i >= argv.len) return ArgError.MissingValue;
            out.model_name = argv[i];
        } else if (std.mem.eql(u8, a, "--column")) {
            if (out.subcommand != .embed) return ArgError.UnknownFlag;
            i += 1;
            if (i >= argv.len) return ArgError.MissingValue;
            out.column_name = argv[i];
        } else if (std.mem.eql(u8, a, "--coverage")) {
            if (out.subcommand != .embed) return ArgError.UnknownFlag;
            i += 1;
            if (i >= argv.len) return ArgError.MissingValue;
            out.coverage_range = argv[i];
        } else if (std.mem.eql(u8, a, "--id")) {
            if (out.subcommand != .embed) return ArgError.UnknownFlag;
            i += 1;
            if (i >= argv.len) return ArgError.MissingValue;
            out.coverage_id = argv[i];
        } else if (std.mem.eql(u8, a, "--dtype")) {
            if (out.subcommand != .embed) return ArgError.UnknownFlag;
            i += 1;
            if (i >= argv.len) return ArgError.MissingValue;
            out.dtype_name = argv[i];
        } else if (std.mem.eql(u8, a, "--all-sheets")) {
            // iter59c: no value; expands selection to every sheet.
            // On workbook-scoped sub-commands silently accept (same
            // tolerance group as --sheet/--name) so wrappers can set
            // it universally without an exit-1.
            if (!workbook_scoped and (out.sheet_index != null or out.sheet_name != null or out.sheet_glob != null))
                return ArgError.SheetArgConflict;
            out.all_sheets = true;
        } else if (std.mem.eql(u8, a, "--sheet-glob")) {
            i += 1;
            if (i >= argv.len) return ArgError.MissingValue;
            if (!workbook_scoped and (out.sheet_index != null or out.sheet_name != null or out.all_sheets))
                return ArgError.SheetArgConflict;
            out.sheet_glob = argv[i];
        } else if (std.mem.eql(u8, a, "--format")) {
            i += 1;
            if (i >= argv.len) return ArgError.MissingValue;
            const v = argv[i];
            if (std.mem.eql(u8, v, "jsonl")) {
                out.format = .jsonl;
            } else if (std.mem.eql(u8, v, "legacy-jsonl")) {
                out.format = .legacy_jsonl;
            } else if (std.mem.eql(u8, v, "legacy-jsonl-dict")) {
                out.format = .legacy_jsonl_dict;
            } else if (std.mem.eql(u8, v, "jsonl-dict")) {
                // Deprecated alias for `legacy-jsonl-dict` — routed
                // through the deprecation flag so `main` emits one
                // stderr warning. Pre-iter55a the only dict shape we
                // shipped was the bare object, so the intent is clear.
                out.format = .legacy_jsonl_dict;
                out.deprecated_jsonl_dict = true;
            } else if (std.mem.eql(u8, v, "tsv")) {
                out.format = .tsv;
            } else if (std.mem.eql(u8, v, "csv")) {
                out.format = .csv;
            } else {
                if (workbook_scoped) continue; // ignore unknown format for meta/list-sheets
                return ArgError.BadFormat;
            }
        } else if (std.mem.eql(u8, a, "--skip")) {
            i += 1;
            if (i >= argv.len) return ArgError.MissingValue;
            // --skip / --take are strict for EVERY sub-command
            // (unlike --sheet / --format whose tolerance depends on
            // workbook_scoped). Pagination is too useful on styles /
            // sst — those commands dump huge streams and a typoed
            // --take that silently returned everything would be a
            // very expensive surprise. For meta / list-sheets which
            // don't paginate, rejecting a --skip typo is also the
            // clearer user-signal: the flag is not effective there.
            out.skip = std.fmt.parseInt(usize, argv[i], 10) catch return ArgError.BadArgValue;
        } else if (std.mem.eql(u8, a, "--take")) {
            i += 1;
            if (i >= argv.len) return ArgError.MissingValue;
            out.take = std.fmt.parseInt(usize, argv[i], 10) catch return ArgError.BadArgValue;
        } else if (std.mem.eql(u8, a, "--start-row")) {
            i += 1;
            if (i >= argv.len) return ArgError.MissingValue;
            // Strict on every sub-command (same rationale as --skip/--take):
            // silently dropping a typoed row bound is an expensive surprise.
            // OOXML rows are 1-based; 0 is a user error and we reject it.
            const v = std.fmt.parseInt(u32, argv[i], 10) catch return ArgError.BadArgValue;
            if (v == 0) return ArgError.BadArgValue;
            out.start_row = v;
        } else if (std.mem.eql(u8, a, "--end-row")) {
            i += 1;
            if (i >= argv.len) return ArgError.MissingValue;
            const v = std.fmt.parseInt(u32, argv[i], 10) catch return ArgError.BadArgValue;
            if (v == 0) return ArgError.BadArgValue;
            out.end_row = v;
        } else if (std.mem.eql(u8, a, "--range")) {
            i += 1;
            if (i >= argv.len) return ArgError.MissingValue;
            // `<topLeft>:<bottomRight>`, both A1-style. Single-cell
            // input (no colon) is rejected — the flag's contract is a
            // rectangle per docs/jq-for-excel.md v4.1. Inverted corners
            // (e.g. `Z1:A1`) are also rejected rather than silently
            // normalised: the user wrote them in that order on purpose
            // or by mistake, and a typo-tolerant swap hides the mistake.
            const raw = argv[i];
            const colon = std.mem.indexOfScalar(u8, raw, ':') orelse return ArgError.BadArgValue;
            const tl = xlsx.parseA1Ref(raw[0..colon]) catch return ArgError.BadArgValue;
            const br = xlsx.parseA1Ref(raw[colon + 1 ..]) catch return ArgError.BadArgValue;
            if (tl.col > br.col or tl.row > br.row) return ArgError.BadArgValue;
            out.range = .{ .top_left = tl, .bottom_right = br };
        } else if (std.mem.eql(u8, a, "--output")) {
            i += 1;
            if (i >= argv.len) return ArgError.MissingValue;
            // Strict for every sub-command (not in the workbook_scoped
            // tolerance group): silently dropping a typoed wire-shape
            // value would hide a bug in consumer scripts that depend on
            // the alternate envelope. Fail loud instead.
            const v = argv[i];
            if (std.mem.eql(u8, v, "ndjson")) {
                out.output = .ndjson;
            } else if (std.mem.eql(u8, v, "compact-ndjson")) {
                out.output = .compact_ndjson;
            } else if (std.mem.eql(u8, v, "pretty-json")) {
                out.output = .pretty_json;
            } else {
                return ArgError.BadArgValue;
            }
        } else if (a.len > 0 and a[0] == '-') {
            return ArgError.UnknownFlag;
        } else {
            if (!first_positional_seen) {
                first_positional_seen = true;
                // Sub-command token already handled by detectSubcommand.
                // Skip it here so it isn't mistaken for the file path.
                if (std.mem.eql(u8, a, "cells") or
                    std.mem.eql(u8, a, "rows") or
                    std.mem.eql(u8, a, "meta") or
                    std.mem.eql(u8, a, "list-sheets") or
                    std.mem.eql(u8, a, "scrub-metadata") or
                    std.mem.eql(u8, a, "embed") or
                    std.mem.eql(u8, a, "comments") or
                    std.mem.eql(u8, a, "validations") or
                    std.mem.eql(u8, a, "hyperlinks") or
                    std.mem.eql(u8, a, "pivots") or
                    std.mem.eql(u8, a, "styles") or
                    std.mem.eql(u8, a, "sst") or
                    std.mem.eql(u8, a, "merges") or
                    std.mem.eql(u8, a, "defined-names") or
                    std.mem.eql(u8, a, "doc-props") or
                    std.mem.eql(u8, a, "anchors") or
                    std.mem.eql(u8, a, "append-rows") or
                    std.mem.eql(u8, a, "set-cell") or
                    std.mem.eql(u8, a, "insert-row") or
                    std.mem.eql(u8, a, "delete-row") or
                    std.mem.eql(u8, a, "insert-column") or
                    std.mem.eql(u8, a, "delete-column") or
                    std.mem.eql(u8, a, "add-sheet") or
                    std.mem.eql(u8, a, "rename-sheet") or
                    std.mem.eql(u8, a, "delete-sheet") or
                    std.mem.eql(u8, a, "rename-table-column"))
                {
                    continue;
                }
            }
            if (out.file.len == 0) out.file = a else return ArgError.UnknownFlag;
        }
    }
    if (out.file.len == 0) return ArgError.NoFile;

    // iter59b-1: --start-row / --end-row only map to sub-commands
    // that emit row-keyed records (rows / cells / comments). The
    // range-keyed commands (validations / hyperlinks) and the
    // workbook-scoped commands (meta / list-sheets / styles / sst)
    // have no per-record row number — reject the flag rather than
    // silently ignoring it.
    if (out.start_row != null or out.end_row != null) {
        switch (detected_sub) {
            .rows, .cells, .comments => {},
            .validations, .hyperlinks, .pivots, .merges, .defined_names, .doc_props, .anchors, .meta, .list_sheets, .styles, .sst, .append_rows, .set_cell, .insert_row, .delete_row, .insert_column, .delete_column, .add_sheet, .rename_sheet, .delete_sheet, .rename_table_column, .scrub_metadata, .embed => {
                return ArgError.BadArgValue;
            },
        }
    }
    // iter59b-2: --range is tighter than --start-row / --end-row — it
    // filters by BOTH row AND column, so `comments` (which emits per
    // cell-ref but has no col-keyed wire contract yet) is deliberately
    // NOT included here. Only `rows` and `cells` accept --range.
    if (out.range != null) {
        switch (detected_sub) {
            .rows, .cells => {},
            .comments, .validations, .hyperlinks, .pivots, .merges, .defined_names, .doc_props, .anchors, .meta, .list_sheets, .styles, .sst, .append_rows, .set_cell, .insert_row, .delete_row, .insert_column, .delete_column, .add_sheet, .rename_sheet, .delete_sheet, .rename_table_column, .scrub_metadata, .embed => {
                return ArgError.BadArgValue;
            },
        }
    }
    // Empty emission ranges are caught at parse time — `start > end`
    // can never produce a record, which is almost certainly a typo.
    if (out.start_row) |s| if (out.end_row) |e| {
        if (s > e) return ArgError.BadArgValue;
    };
    // The legacy --list-sheets flag takes an early return in main
    // and emits plain sheet names — no row concept. Row bounds
    // passed alongside it would silently no-op, hiding typos.
    if (out.list_sheets and (out.start_row != null or out.end_row != null or out.range != null)) {
        return ArgError.BadArgValue;
    }
    // iter59b-3: --header promotes the first row to keys. It only
    // composes with the `rows` sub-command on the NDJSON envelope
    // (the flat formats each have their own well-defined row shape).
    // Silently no-op'ing on mismatch would hide user typos, so reject.
    if (out.header) {
        if (detected_sub != .rows) return ArgError.BadArgValue;
        if (out.format != .jsonl) return ArgError.BadArgValue;
        if (out.list_sheets) return ArgError.BadArgValue;
    }
    // iter59b-4: --include-blanks and --with-styles both scope to the
    // two cell-shape emitters (`cells` / `rows`) and neither makes sense
    // on the legacy --list-sheets flag. Reject rather than silently
    // no-op to surface user typos (same rationale as --header).
    if (out.include_blanks) {
        switch (detected_sub) {
            .rows, .cells => {},
            .comments, .validations, .hyperlinks, .pivots, .merges, .defined_names, .doc_props, .anchors, .meta, .list_sheets, .styles, .sst, .append_rows, .set_cell, .insert_row, .delete_row, .insert_column, .delete_column, .add_sheet, .rename_sheet, .delete_sheet, .rename_table_column, .scrub_metadata, .embed => {
                return ArgError.BadArgValue;
            },
        }
        if (out.list_sheets) return ArgError.BadArgValue;
    }
    if (out.with_styles) {
        switch (detected_sub) {
            .rows, .cells => {},
            .comments, .validations, .hyperlinks, .pivots, .merges, .defined_names, .doc_props, .anchors, .meta, .list_sheets, .styles, .sst, .append_rows, .set_cell, .insert_row, .delete_row, .insert_column, .delete_column, .add_sheet, .rename_sheet, .delete_sheet, .rename_table_column, .scrub_metadata, .embed => {
                return ArgError.BadArgValue;
            },
        }
        if (out.list_sheets) return ArgError.BadArgValue;
        // `cells` shape is fixed (ignores --format) so --with-styles is
        // always welcome there. `rows` only has a style-shaped place on
        // the envelope; csv/tsv/legacy-jsonl/legacy-jsonl-dict shapes
        // don't nest, and the --header dict has no per-cell slot.
        if (detected_sub == .rows) {
            if (out.format != .jsonl) return ArgError.BadArgValue;
            if (out.header) return ArgError.BadArgValue;
        }
    }
    // iter60b: `pretty-json` is the collapsed-object variant
    // (docs/jq-for-excel.md v4.1). Every streaming sub-command emits a
    // record-per-line shape that a collapsed single-object rewrite
    // doesn't compose with, so those are still rejected at parse time
    // rather than silently falling back.
    //
    // `list-sheets` joined `meta` here: it is workbook-scoped and
    // bounded (one record per sheet), so a single object is a coherent
    // shape — and callers checking sheet visibility want to `jq` the
    // whole answer at once rather than stream it.
    if (out.output == .pretty_json and
        detected_sub != .meta and detected_sub != .list_sheets)
    {
        return ArgError.BadArgValue;
    }
    // `--format` shapes only `rows` output. On sheet-scoped sub-commands
    // (cells / comments / validations / hyperlinks) the emitter ignores
    // it, so reject anything but the implicit `jsonl` default rather than
    // silently producing the wrong shape. Workbook-scoped sub-commands
    // (meta / list-sheets / styles / sst) keep their existing tolerance —
    // wrappers append `--format` universally and we don't want to break
    // them.
    if (!workbook_scoped and detected_sub != .rows and out.format != .jsonl) {
        return ArgError.BadFormat;
    }

    // Structural-edit-only flags are rejected on every other sub-
    // command. Mirrors the strictness of --header / --include-blanks
    // / --start-row so user typos surface as exit-1 BadArgValue
    // rather than silently no-op'ing the requested edit.
    if (out.row_1based != null and !(detected_sub == .insert_row or detected_sub == .delete_row)) {
        return ArgError.BadArgValue;
    }
    if (out.col_letter != null and !(detected_sub == .insert_column or detected_sub == .delete_column)) {
        return ArgError.BadArgValue;
    }
    if (out.new_sheet_name != null and !(detected_sub == .add_sheet or detected_sub == .rename_sheet or detected_sub == .rename_table_column)) {
        return ArgError.BadArgValue;
    }
    if ((out.table_name != null or out.old_column_name != null) and detected_sub != .rename_table_column) {
        return ArgError.BadArgValue;
    }
    if (out.cell_ref != null and detected_sub != .set_cell) {
        return ArgError.BadArgValue;
    }
    if (out.cell_value_json != null and detected_sub != .set_cell) {
        return ArgError.BadArgValue;
    }
    // --out is required by every edit command and silently ignored
    // on read-only paths if we don't reject. Without this guard,
    // `zlsx rows in.xlsx --out out.jsonl` exits 0 but never writes
    // the file — surprising and quietly destructive when the user
    // expects out.jsonl to land.
    if (out.out_path != null) switch (detected_sub) {
        .append_rows,
        .set_cell,
        .insert_row,
        .delete_row,
        .insert_column,
        .delete_column,
        .add_sheet,
        .rename_sheet,
        .delete_sheet,
        .rename_table_column,
        .scrub_metadata,
        .embed,
        => {},
        else => return ArgError.BadArgValue,
    };
    return out;
}

fn writeUsage(w: *std.Io.Writer) !void {
    try w.writeAll(
        \\usage: zlsx [<subcommand>] <file.xlsx> [options]
        \\
        \\formula sub-commands (M6, own grammars — see each --help):
        \\  eval              evaluate one formula against a workbook (NDJSON)
        \\  recalc            recalculate every formula cell into --out
        \\
        \\  --sheet N         0-indexed sheet to read (default: 0; on
        \\                    pivots every host sheet, on merges and
        \\                    anchors every sheet, on defined-names
        \\                    every name)
        \\  --name NAME       select sheet by name (conflicts with --sheet)
        \\  --all-sheets      (iter59c) iterate every sheet. Mutually
        \\                    exclusive with --sheet / --name / --sheet-glob.
        \\                    On `cells` / `rows`, --skip / --take apply
        \\                    GLOBALLY across the concatenated cross-sheet
        \\                    stream (not per sheet). --header / --start-row
        \\                    / --end-row / --range are PER-SHEET — each
        \\                    sheet independently resolves its first-row
        \\                    keys and row bounds. Silently accepted (and
        \\                    ignored) on workbook-scoped sub-commands.
        \\  --sheet-glob PAT  (iter59c) simple-glob over sheet names:
        \\                    `*` matches any run, `?` one char; case-
        \\                    sensitive. Mutually exclusive with --sheet /
        \\                    --name / --all-sheets. Scope and per-sheet
        \\                    vs global interactions match --all-sheets.
        \\  --format FMT      jsonl | legacy-jsonl | legacy-jsonl-dict | jsonl-dict | tsv | csv
        \\                    (default: jsonl — NDJSON row envelope; iter55a.
        \\                    Applies to the `rows` sub-command only; ignored
        \\                    by `cells`, which always emits per-cell NDJSON.
        \\                    `jsonl-dict` is a deprecated alias for
        \\                    `legacy-jsonl-dict` — accepted this release.)
        \\  --list-sheets     print sheet names, one per line, and exit
        \\                    (legacy plain-text flag — still works.
        \\                    The `list-sheets` sub-command emits NDJSON.)
        \\  --skip N          drop the first N emitted records (iter59a).
        \\                    Applies globally to the record stream of
        \\                    rows / cells / comments / validations /
        \\                    hyperlinks / pivots / merges / defined-names /
        \\                    anchors / styles / sst. Ignored by meta,
        \\                    list-sheets and doc-props (a single-record
        \\                    report).
        \\  --take N          stop after N emitted records. Same scope
        \\                    as --skip; combine for middle-slice paging.
        \\  --start-row R     (iter59b) 1-based OOXML row; drop records
        \\                    whose row < R. Per-sheet scope (each
        \\                    sheet's own rows, unlike --skip which is
        \\                    global). Valid for rows / cells / comments
        \\                    only; rejected on validations / hyperlinks
        \\                    / pivots / merges / defined-names / doc-props
        \\                    / anchors / meta / list-sheets / styles / sst.
        \\  --end-row R       (iter59b) 1-based OOXML row; stop emitting
        \\                    after row R (inclusive). Same scope and
        \\                    sub-command constraints as --start-row.
        \\                    Applied BEFORE --skip / --take.
        \\  --header          (iter59b-3) promote the first emitted row
        \\                    to field keys on `rows --format jsonl`. The
        \\                    header row itself is NOT emitted; subsequent
        \\                    rows become {"kind":"row",…,"fields":{k:v,…}}.
        \\                    Numeric/boolean header cells stringify; empty
        \\                    header cells fall back to "col_A"/"col_B"/…
        \\                    Duplicate keys are emitted verbatim (JSON
        \\                    accepts them; consumer handles de-dup).
        \\                    Rejected on every other sub-command and on
        \\                    --format other than jsonl — those shapes
        \\                    don't compose with the field-dict contract.
        \\  --range A1:Z100   (iter59b-2) A1-style bounding rectangle;
        \\                    valid for `rows` and `cells` only. Inverted
        \\                    corners (e.g. `Z1:A1`) are rejected. When
        \\                    combined with --start-row / --end-row the
        \\                    row bounds are intersected (most-restrictive
        \\                    wins). For `rows` a row is emitted iff it
        \\                    has at least one in-range cell; out-of-range
        \\                    columns are masked to empty so the cells[]
        \\                    array stays column-indexed.
        \\  --include-blanks  (iter59b-4) emit empty cells as
        \\                    {"t":"blank","v":null} records instead of
        \\                    skipping them. Applies to `cells` and the
        \\                    `rows --format jsonl` envelope. No-op on
        \\                    `rows --header` (the fields dict already
        \\                    emits `key:null` for missing cells) and on
        \\                    legacy flat formats. Rejected on every other
        \\                    sub-command.
        \\  --with-styles     (iter59b-4) attach a terse `style:{…}` field
        \\                    to per-cell records, populated only when the
        \\                    cell has effective formatting (bold / italic
        \\                    / fg / bg / non-General nf / any border).
        \\                    Unstyled cells OMIT the field entirely.
        \\                    Terse shape: {"bold":true,"italic":true,
        \\                      "fg":"FF…","bg":"FF…","nf":"0.00",
        \\                      "border":{"l":{"s":"thin","c":"FF…"},…}}.
        \\                    Valid on `cells` and `rows --format jsonl`
        \\                    only; rejected on `rows --header` (no slot
        \\                    in the fields dict) and flat formats.
        \\  --output MODE     (iter60b) wire-shape switch:
        \\                    ndjson           (default) invariant-envelope
        \\                                     NDJSON — every record carries
        \\                                     `sheet`/`sheet_idx`.
        \\                    compact-ndjson   per-sheet prologue record
        \\                                     {"kind":"sheet","sheet":"S",
        \\                                     "sheet_idx":N} emitted once per
        \\                                     sheet; subsequent records OMIT
        \\                                     `sheet`/`sheet_idx`. Applies to
        \\                                     cells / rows / comments /
        \\                                     validations / hyperlinks / pivots
        \\                                     / merges / anchors.
        \\                                     `--skip`/`--take` still slice
        \\                                     data records (prologues aren't
        \\                                     counted). On `meta` the
        \\                                     per-sheet records drop
        \\                                     `sheet`/`sheet_idx`; on
        \\                                     `list-sheets` / `styles` / `sst`
        \\                                     / `defined-names` / `doc-props`
        \\                                     it's effectively a no-op (a
        \\                                     defined name's sheet fields are
        \\                                     its SCOPE, not a host location;
        \\                                     doc-props has no sheet fields).
        \\                    pretty-json      meta-only. Collapses workbook
        \\                                     + sheets into one 2-space-
        \\                                     indented JSON object. The
        \\                                     scalar sheet count is
        \\                                     `sheets_count` in this mode
        \\                                     (to avoid the `sheets: [...]`
        \\                                     array-field collision); NDJSON
        \\                                     keeps the original `sheets:N`.
        \\                                     Rejected on every other
        \\                                     sub-command.
        \\  -h, --help        show this help
        \\
        \\Sub-commands
        \\  rows               (default) one NDJSON envelope per row — see Formats.
        \\                     Bare `zlsx file.xlsx` is an alias for `zlsx rows file.xlsx`.
        \\  cells              one NDJSON record per non-empty cell (iter56):
        \\                     {"kind":"cell","sheet":"S","sheet_idx":0,"ref":"A1",
        \\                      "row":1,"col":1,"t":"str","v":"x"}
        \\                     t ∈ {"str","int","num","bool"}. Empty cells skipped.
        \\                     --format is ignored; output shape is fixed.
        \\  meta               workbook summary as NDJSON (iter57). One
        \\                     workbook record first, then one sheet record per sheet:
        \\                     {"kind":"workbook","path":"f.xlsx","sheets":N,
        \\                      "sst":{"count":C,"rich":R},
        \\                      "has_styles":bool,"has_theme":bool,"has_comments":bool}
        \\                     {"kind":"sheet","sheet":"S","sheet_idx":0,
        \\                      "has_comments":bool}
        \\                     --format / --sheet / --name are ignored.
        \\  list-sheets        lighter NDJSON variant of `meta` (iter57):
        \\                     one {"kind":"sheet","sheet":…,"sheet_idx":…}
        \\                     record per sheet. For the plain-text one-name-
        \\                     per-line shape, use the legacy `--list-sheets` flag.
        \\  comments           one NDJSON record per cell comment across every
        \\                     sheet (iter58):
        \\                     {"kind":"comment","sheet":"S","sheet_idx":0,
        \\                      "ref":"A1","row":1,"col":1,"author":"Alice",
        \\                      "text":"…","runs":null}
        \\  validations        one NDJSON record per data-validation range
        \\                     across every sheet (iter58):
        \\                     {"kind":"validation","sheet":"S","sheet_idx":0,
        \\                      "range":"B2:B100","rule_type":"list","op":null,
        \\                      "formula1":"a,b","formula2":null,
        \\                      "values":["a","b"]}
        \\  hyperlinks         one NDJSON record per hyperlink across every
        \\                     sheet (iter58):
        \\                     {"kind":"hyperlink","sheet":"S","sheet_idx":0,
        \\                      "range":"A1","url":"https://…","location":null}
        \\  pivots             one NDJSON record per pivot table across every
        \\                     sheet (S6), then one per pivot cache no table
        \\                     reads. Host sheet in `sheet`; the cache's
        \\                     `source.resolved` names the sheet the pivot
        \\                     READS from (or another workbook, or null):
        \\                     {"kind":"pivot","sheet":"S","sheet_idx":0,
        \\                      "name":"PivotTable1","part":"xl/pivotTables/…",
        \\                      "location":{"ref":"A3:B6",…},"rows":[{"field":"Region","idx":0}],
        \\                      "cols":[],"pages":[],"values":[{"name":"Sum of Qty",
        \\                      "field":"Qty","idx":1,"subtotal":"sum",…}],
        \\                      "data_caption":"Values","grand_totals":{…},"style":"…",
        \\                      "cache":{"id":7,"part":"…","records_part":"…",
        \\                      "record_count":3,…,"source":{"type":"worksheet",
        \\                      "sheet":"Data","ref":"A1:C4","name":null,
        \\                      "resolved":{"sheet":"Data","sheet_idx":0,"via":"sheet_attr",
        \\                      "bounds":"A1:C4"},"unresolved":null},
        \\                      "fields":[{"name":"Region","types":["string"],…}]}}
        \\  styles             one NDJSON record per cell-XF style entry
        \\                     (workbook-wide, iter58):
        \\                     {"kind":"style","idx":0,"font":{…}|null,
        \\                      "fill":{…}|null,"border":{…}|null,
        \\                      "num_fmt":"General"|null}
        \\  sst                one NDJSON record per shared-string entry
        \\                     (workbook-wide, iter58):
        \\                     {"kind":"sst","idx":0,"text":"…","runs":null}
        \\  merges             one NDJSON record per merged range across every
        \\                     sheet (S3b):
        \\                     {"kind":"merge","sheet":"S","sheet_idx":0,
        \\                      "range":"A1:B3","start_row":1,"start_col":1,
        \\                      "end_row":3,"end_col":2}
        \\                     Rows and cols are 1-based, corners inclusive.
        \\  defined-names      one NDJSON record per <definedName> of
        \\                     xl/workbook.xml, document order (S3b):
        \\                     {"kind":"defined_name","name":"Prices",
        \\                      "scope":"workbook","sheet":null,"sheet_idx":null,
        \\                      "body":"Data!$A$1:$C$4","hidden":false}
        \\                     scope ∈ {"workbook","sheet"}; a sheet-scoped
        \\                     name carries its sheet + 0-based sheet_idx
        \\                     (localSheetId). --sheet / --name / --sheet-glob
        \\                     narrow to names SCOPED to matching sheets and
        \\                     suppress workbook-scope names; the default and
        \\                     --all-sheets stream every name.
        \\  doc-props          the document-properties field set as ONE
        \\                     NDJSON record (S3b):
        \\                     {"kind":"doc_props","creator":…,
        \\                      "last_modified_by":…,"title":…,"subject":…,
        \\                      "description":…,"keywords":…,"category":…,
        \\                      "created":…,"modified":…,"revision":…,
        \\                      "company":…,"manager":…,"application":…,
        \\                      "hyperlink_base":…,
        \\                      "has_custom_properties":bool}
        \\                     Absent fields are null; text is as stored
        \\                     (meta's values; Python's Editor.doc_props
        \\                     maps a present-but-empty element to None).
        \\                     A workbook with no docProps parts is a
        \\                     record of nulls. Workbook-scoped: sheet
        \\                     selectors are tolerated and ignored.
        \\  anchors            one NDJSON record per anchored image or
        \\                     chart across every sheet (S3b) — images
        \\                     before charts within a sheet:
        \\                     {"kind":"image_anchor","sheet":"S","sheet_idx":0,
        \\                      "part":"xl/media/image1.png","anchor":"two_cell",
        \\                      "from":{"row":3,"col":2,"row_off":0,"col_off":9525},
        \\                      "to":{…},"absolute":null,"bytes":4096}
        \\                     {"kind":"chart_anchor",…,"part":"xl/charts/chart1.xml",
        \\                      "anchor":"one_cell","from":{…},"to":null,
        \\                      "absolute":null,"chart_type":"bar",
        \\                      "series_refs":["Data!$B$2:$B$4"]}
        \\                     anchor ∈ {"two_cell","one_cell","absolute"};
        \\                     rows/cols 1-based, offsets in EMUs verbatim;
        \\                     absolute anchors carry {"x","y","cx","cy"}
        \\                     and null from/to. The payload stays in the
        \\                     archive — "bytes" is the image part's size.
        \\  append-rows        load-modify-save: append rows to an existing
        \\                     sheet, atomic-rename to --out. Reads NDJSON
        \\                     row arrays from stdin (one JSON array per
        \\                     line). Cell types: null→empty, true/false→
        \\                     bool, integer→int, number→float, string→str.
        \\                     Required: --out PATH and --sheet N.
        \\                     `cat rows.ndjson | zlsx append-rows in.xlsx
        \\                       --sheet 0 --out out.xlsx`
        \\  set-cell           load-modify-save: rewrite a single cell in
        \\                     place. Required: --out PATH, --sheet N,
        \\                     --ref A1, --value <JSON>. Value is one
        \\                     JSON token (string, integer, float, true,
        \\                     false, null) — same type-mapping as
        \\                     append-rows. Insert / empty-row paths handled.
        \\                     `zlsx set-cell in.xlsx --sheet 0 --ref C5
        \\                       --value '"hello"' --out out.xlsx`
        \\  insert-row /       structural row edit. Required: --out PATH,
        \\  delete-row         --sheet N, --row N (1-based). Refuses on
        \\                     sheets carrying formulas, hyperlinks, data
        \\                     validations, conditional formatting, frozen
        \\                     panes, or other constructs the rewriter
        \\                     doesn't yet shift.
        \\  insert-column /    structural column edit. Required: --out PATH,
        \\  delete-column      --sheet N, --col LETTER (A..XFD). Same
        \\                     refusal contract as the row edits.
        \\  add-sheet          load-modify-save: append a new (empty)
        \\                     sheet. Required: --out PATH, --new-name NAME.
        \\  rename-sheet       load-modify-save: rename a sheet. Required:
        \\                     --out PATH, --sheet N, --new-name NAME.
        \\  rename-table-column
        \\                     load-modify-save: rename a column of a named
        \\                     table; every structured reference follows.
        \\                     Required: --out PATH, --table NAME,
        \\                     --old-name OLD, --new-name NEW.
        \\  delete-sheet       load-modify-save: drop a sheet (cannot drop
        \\                     the last remaining sheet). Required:
        \\                     --out PATH, --sheet N.
        \\  embed --strip      load-modify-save: remove the embedding parts
        \\                     AND the recovery record, so the result reports
        \\                     `absent` rather than `stripped`. The pre-share
        \\                     operation — use it before handing a workbook to
        \\                     someone who should not have the vectors or the
        \\                     provenance. Required: --strip, --out PATH.
        \\                     e.g. `zlsx embed book.xlsx --strip --out clean.xlsx`
        \\  embed --prune      load-modify-save: tombstone every slot whose
        \\                     row is no longer embeddable and zero its
        \\                     vector — the redaction sweep. A row deleted
        \\                     in plain Excel leaves its vector on disk;
        \\                     this removes it. Content that merely changed
        \\                     is reported stale, NOT redacted. Emits one
        \\                     NDJSON summary on stdout. `count` and the
        \\                     coverage range never move. Required:
        \\                     --prune, --out PATH. Mutually exclusive
        \\                     with --strip.
        \\  embed --extract    read-only: emit one NDJSON record per row
        \\                     that has something worth embedding —
        \\                     {"kind":"embed_row","row":N,"text":"…"}.
        \\                     Rows with nothing embeddable are omitted,
        \\                     not emitted empty. Required: --column A,
        \\                     --coverage A2:A100. Takes no --out.
        \\  embed --vectors P  write the embedding parts from NDJSON
        \\                     vectors — {"row":N,"vector":[…]}. Covered
        \\                     rows with no vector become tombstones, so a
        \\                     partial embedding is representable. All
        \\                     vectors must share one dimension. Required:
        \\                     --column A, --coverage A2:A100, --model
        \\                     NAME, --out PATH. Optional: --id NAME
        \\                     (default "default"), --dtype f32|int8-sym
        \\                     (default f32), --sheet N (default 0).
        \\
        \\                     The model is invoked out of band — the
        \\                     embed pipeline never makes a network call
        \\                     — so the two phases compose through a pipe:
        \\                       zlsx embed b.xlsx --extract --column A \
        \\                         --coverage A2:A100 > rows.ndjson
        \\                       my-embedder < rows.ndjson > vecs.ndjson
        \\                       zlsx embed b.xlsx --vectors vecs.ndjson \
        \\                         --model M --column A --coverage A2:A100 \
        \\                         --out out.xlsx
        \\  dbx push|pull|genie  Databricks over REST (the one network-
        \\                     touching family). Auth from DATABRICKS_HOST
        \\                     / DATABRICKS_TOKEN; genie space from
        \\                     GENIE_SPACE_ID. push/pull transfer workbooks
        \\                     to/from a UC Volume and REFUSE bytes that do
        \\                     not parse as a workbook (upload preflight,
        \\                     verify-before-rename download); genie asks a
        \\                     Genie space and streams status + SQL + rows
        \\                     as NDJSON. `zlsx dbx --help` for details.
        \\
        \\Formats (rows only)
        \\  jsonl              NDJSON row envelope (default, iter55a):
        \\                     {"kind":"row","sheet":"S","sheet_idx":0,"row":1,
        \\                      "cells":[{"ref":"A1","col":1,"t":"str","v":"x"},…]}
        \\                     t ∈ {"str","int","num","bool"}; empty cells skipped.
        \\  legacy-jsonl       pre-iter55a bare arrays:  [1, "foo", null, true]
        \\  legacy-jsonl-dict  pre-iter55a bare objects: {"A": 1, "B": "foo"}
        \\                     (alias `jsonl-dict` accepted this release for back-
        \\                     compat; will warn in a future release)
        \\  tsv                tab-separated, \N for empty cells, control chars escaped
        \\  csv                RFC 4180, empty string for empty cells
        \\
        \\Error records (iter60c)
        \\  Non-fatal parse errors (e.g. a single corrupt sheet inside an
        \\  otherwise-valid workbook) are emitted inline as
        \\  {"kind":"error","sheet":"…","sheet_idx":N,"scope":"sheet",
        \\   "code":"MalformedXml","message":"…"}
        \\  records and the run still exits 0. Pipelines may strip them
        \\  via `jq 'select(.kind!="error")'`. Workbook-level open
        \\  failures stay fatal (exit 2 / 5 — see Exit codes).
        \\
        \\Exit codes
        \\  0  success
        \\  1  bad arguments
        \\  2  could not open or parse workbook
        \\  3  sheet not found
        \\
    );
}

fn colLetter(buf: *[8]u8, idx: usize) []const u8 {
    // Unchecked writer: `idx` is a 0-based position in an emitted row,
    // never validated against the grid on this path. An 8-byte buffer
    // covers any `u32`, so the failure branch is unreachable.
    const n = coords.writeColNumberLetters(buf, @intCast(idx + 1)) catch unreachable;
    return buf[0..n];
}

/// The one JSON escaper every surface shares (`pkg/json_text.zig`):
/// the C ABI's NDJSON buffers and these records spell a string the
/// same way by construction.
fn writeJsonString(w: *std.Io.Writer, s: []const u8) !void {
    return zlsx_pkg.json_text.writeString(w, s);
}

fn writeJsonCell(w: *std.Io.Writer, cell: xlsx.Cell) !void {
    switch (cell) {
        .empty => try w.writeAll("null"),
        .string => |s| try writeJsonString(w, s),
        .integer => |x| try w.print("{d}", .{x}),
        .number => |f| {
            if (std.math.isFinite(f)) {
                try w.print("{d}", .{f});
            } else {
                // JSON has no NaN/Inf — emit null so parsers don't choke.
                try w.writeAll("null");
            }
        },
        .boolean => |b| try w.writeAll(if (b) "true" else "false"),
    }
}

fn writeTsvField(w: *std.Io.Writer, s: []const u8) !void {
    for (s) |c| switch (c) {
        '\t' => try w.writeAll("\\t"),
        '\n' => try w.writeAll("\\n"),
        '\r' => try w.writeAll("\\r"),
        '\\' => try w.writeAll("\\\\"),
        else => try w.writeByte(c),
    };
}

fn writeCsvField(w: *std.Io.Writer, s: []const u8) !void {
    var needs_quote = false;
    for (s) |c| {
        if (c == ',' or c == '"' or c == '\n' or c == '\r') {
            needs_quote = true;
            break;
        }
    }
    if (!needs_quote) {
        try w.writeAll(s);
        return;
    }
    try w.writeByte('"');
    for (s) |c| {
        if (c == '"') try w.writeAll("\"\"") else try w.writeByte(c);
    }
    try w.writeByte('"');
}

/// Per-cell `t` type tag for the envelope schema. Mirrors the
/// design-doc "cells" record but limited to the four primitive
/// types this slice emits — formula is future work. The
/// `t:"date"` variant lives on a parallel boolean channel (see
/// `Rows.dateTypes()`), not in the `Cell` union itself; same shape
/// for `t:"error"` via `Rows.errorStrings()`.
fn envelopeTypeTag(cell: xlsx.Cell) []const u8 {
    return switch (cell) {
        .empty => unreachable, // caller skips empties
        .string => "str",
        .integer => "int",
        .number => "num",
        .boolean => "bool",
    };
}

/// Emit an ISO-8601 date-time as `YYYY-MM-DDTHH:MM:SS` (no trailing
/// `Z`, no fractional seconds). Matches the jq-for-excel design-doc
/// shape for `t:"date"` cells.
fn writeIsoDateTime(w: *std.Io.Writer, dt: xlsx.DateTime) !void {
    try w.print(
        "{d:0>4}-{d:0>2}-{d:0>2}T{d:0>2}:{d:0>2}:{d:0>2}",
        .{ dt.year, dt.month, dt.day, dt.hour, dt.minute, dt.second },
    );
}

/// Emit a date cell's `v` + `serial` pair: `"<ISO>",serial:<N>` (no
/// leading comma, the caller writes `"v":` first). Used by both
/// `writeCell` and `writeEnvelopeCells` so the on-wire shape stays
/// in lockstep. Caller guarantees the serial is inside
/// `fromExcelSerial`'s accepted range (the `.date` side channel on
/// `Rows` already filtered it).
fn writeDateValueAndSerial(w: *std.Io.Writer, serial: f64, uses_1904: bool) !void {
    // Pick the epoch's decoder: 1904 books shift every serial by
    // 1462 days. The Rows date side-channel set its flag using the
    // same decoder, so the caller's invariant still holds either way.
    const dt = (if (uses_1904) xlsx.fromExcelSerial1904(serial) else xlsx.fromExcelSerial(serial)) orelse unreachable;
    try w.writeByte('"');
    try writeIsoDateTime(w, dt);
    try w.writeByte('"');
    // Serial prints as an integer when it has no fractional part (the
    // date-only case — majority). {d} already does this for f64. The
    // raw value is what's stored in the source — emit it unchanged
    // so callers know which epoch the serial is on (consumers can
    // cross-check by re-reading the source if epoch matters).
    try w.print(",\"serial\":{d}", .{serial});
}

/// Coerce a numeric `Cell` to its raw f64 serial. Caller has already
/// gated on `Rows.dateTypes()[i] == true`, so the cell is guaranteed
/// to be `.integer` or `.number`.
fn cellToSerial(cell: xlsx.Cell) f64 {
    return switch (cell) {
        .number => |n| n,
        .integer => |n| @floatFromInt(n),
        else => unreachable, // callers gate on dateTypes() == true
    };
}

/// iter61-b: emit a formula record's cached-value tail. The wire
/// shape is `,"cached":<JSON>` for `.integer` / `.number` / `.string`
/// / `.boolean`, or NOTHING for `.empty` (the `cached` field is
/// omitted entirely, per the design doc — a formula with no cached
/// value carries no `cached` key).
///
/// Note: this writes the leading comma; callers must NOT have written
/// one. Mirrors the convention `writeJsonCell` uses on its own.
fn writeFormulaCached(w: *std.Io.Writer, cell: xlsx.Cell) !void {
    switch (cell) {
        .empty => {}, // omit the cached field entirely
        .integer => |x| try w.print(",\"cached\":{d}", .{x}),
        .number => |f| {
            if (std.math.isFinite(f)) {
                try w.print(",\"cached\":{d}", .{f});
            } else {
                // JSON has no NaN/Inf — match writeJsonCell and emit null.
                try w.writeAll(",\"cached\":null");
            }
        },
        .boolean => |b| try w.writeAll(if (b) ",\"cached\":true" else ",\"cached\":false"),
        .string => |s| {
            try w.writeAll(",\"cached\":");
            try writeJsonString(w, s);
        },
    }
}

/// iter61-b: emit an A1-style ref from a `CellRef`. Uses `colLetter`
/// to render the column. The cell ref is the slave's base cell (per
/// `Rows.formulaRefs()`), so by construction `row >= 1` and
/// `col < 16384`.
fn writeFormulaRef(w: *std.Io.Writer, ref: xlsx.CellRef) !void {
    var col_buf: [8]u8 = undefined;
    const letters = colLetter(&col_buf, @intCast(ref.col));
    try w.writeAll("\"formula_ref\":\"");
    try w.print("{s}{d}", .{ letters, ref.row });
    try w.writeByte('"');
}

/// Emit just the `[{ref,col,t,v},…]` array. By default sparse —
/// `.empty` slots are skipped. `row_number` is the 1-based OOXML row
/// used to build each cell's `ref`.
///
/// iter59b-4: when `include_blanks` is set, every `.empty` cell is
/// materialised as `{"ref":…,"col":…,"t":"blank","v":null}`. When
/// `style_ctx` is non-null AND the cell's style index resolves to
/// effective formatting, a terse `style:{…}` field is appended per
/// the design doc.
fn writeEnvelopeCells(
    w: *std.Io.Writer,
    cells: []const xlsx.Cell,
    row_number: u32,
    include_blanks: bool,
    style_ctx: ?EnvelopeStyleCtx,
    col_offset: u32,
    date_types: []const bool,
    error_strings: []const ?[]const u8,
    formula_strings: []const ?[]const u8,
    formula_refs: []const ?xlsx.CellRef,
    uses_1904: bool,
) !void {
    try w.writeByte('[');
    var first = true;
    for (cells, 0..) |c, i| {
        // iter61-b P2 follow-up: a formula cell whose source XML has
        // no cached <v> element comes through as Cell.empty with the
        // formula text in row_formula_strings (or a ref in
        // row_formula_refs). Pre-checking here so the .empty skip
        // doesn't drop those records — runCellsOnSheetCore already
        // does the same for the cells sub-command.
        const has_formula_here = (i < formula_strings.len and formula_strings[i] != null) or
            (i < formula_refs.len and formula_refs[i] != null);
        if (c == .empty and !include_blanks and !has_formula_here) continue;
        if (!first) try w.writeByte(',');
        first = false;

        const absolute_col: u32 = col_offset + @as(u32, @intCast(i));
        var col_buf: [8]u8 = undefined;
        const letters = colLetter(&col_buf, absolute_col);
        var ref_buf: [16]u8 = undefined;
        const ref = std.fmt.bufPrint(&ref_buf, "{s}{d}", .{ letters, row_number }) catch unreachable;

        // date_types / error_strings / formula_strings / formula_refs
        // may be shorter than `cells` when the caller passes an empty
        // slice (e.g. default-opt-out paths); treat the tail as false /
        // null. Tag precedence (most-specific wins): formula > error >
        // date > primitive. The reader's consumeCell already enforces
        // mutual exclusion (a formula cell never has the error side
        // channel set; a numeric formula never has the date flag set),
        // so this ordering is purely defensive — but keeps the rule
        // local and readable.
        const fmla_str: ?[]const u8 = if (i < formula_strings.len) formula_strings[i] else null;
        const fmla_ref: ?xlsx.CellRef = if (i < formula_refs.len) formula_refs[i] else null;
        const err_str: ?[]const u8 =
            if (fmla_str == null and fmla_ref == null and i < error_strings.len) error_strings[i] else null;
        const is_date: bool =
            fmla_str == null and fmla_ref == null and err_str == null and
            i < date_types.len and date_types[i];

        try w.writeAll("{\"ref\":");
        try writeJsonString(w, ref);
        switch (c) {
            .empty => {
                // iter61-b: a formula cell with NO cached `<v>` is
                // legal — `<c r="C2"><f>A2+B2</f></c>` (formula-only).
                // We still surface it as `t:"formula"` rather than
                // collapsing to `t:"blank"`, matching the design doc:
                // the wire record carries `formula:<text>` (or
                // `formula_ref:<A1>`) and omits `cached` entirely.
                if (fmla_str) |text| {
                    try w.print(",\"col\":{d},\"t\":\"formula\",\"formula\":", .{absolute_col + 1});
                    try writeJsonString(w, text);
                } else if (fmla_ref) |base| {
                    try w.print(",\"col\":{d},\"t\":\"formula\",", .{absolute_col + 1});
                    try writeFormulaRef(w, base);
                } else {
                    try w.print(",\"col\":{d},\"t\":\"blank\",\"v\":null", .{absolute_col + 1});
                }
            },
            else => {
                if (fmla_str) |text| {
                    try w.print(",\"col\":{d},\"t\":\"formula\",\"formula\":", .{absolute_col + 1});
                    try writeJsonString(w, text);
                    try writeFormulaCached(w, c);
                } else if (fmla_ref) |base| {
                    try w.print(",\"col\":{d},\"t\":\"formula\",", .{absolute_col + 1});
                    try writeFormulaRef(w, base);
                    try writeFormulaCached(w, c);
                } else if (err_str) |literal| {
                    try w.print(",\"col\":{d},\"t\":\"error\",\"v\":", .{absolute_col + 1});
                    try writeJsonString(w, literal);
                } else if (is_date) {
                    try w.print(",\"col\":{d},\"t\":\"date\",\"v\":", .{absolute_col + 1});
                    try writeDateValueAndSerial(w, cellToSerial(c), uses_1904);
                } else {
                    try w.print(",\"col\":{d},\"t\":\"{s}\",\"v\":", .{ absolute_col + 1, envelopeTypeTag(c) });
                    try writeJsonCell(w, c);
                }
            },
        }
        if (style_ctx) |ctx| {
            // The envelope slice may be wider (padded .empty) or narrower
            // (range-sliced) than the row's actual styleIndices — guard.
            const sidx_opt: ?u32 = if (i < ctx.style_indices.len)
                ctx.style_indices[i]
            else
                null;
            if (sidx_opt) |sidx| if (styleBlockEffective(ctx.book, sidx)) {
                try w.writeAll(",\"style\":");
                _ = try writeTerseStyleBlock(w, ctx.book, sidx);
            };
        }
        try w.writeByte('}');
    }
    try w.writeByte(']');
}

/// iter59b-4: style context for `writeEnvelopeCells`. `style_indices`
/// is the row's per-column style id slice (from `Rows.styleIndices`)
/// aligned by the same position as `cells`. When `--range` masks the
/// cells array, the caller passes an indices slice prepared the same
/// way (masked / sliced identically) so `cells[i]` and `style_indices[i]`
/// agree.
const EnvelopeStyleCtx = struct {
    book: *const xlsx.Book,
    style_indices: []const ?u32,
};

/// Emit one NDJSON envelope line:
/// `{"kind":"row","sheet":…,"sheet_idx":…,"row":…,"cells":[…]}\n`.
/// All-empty rows still emit the envelope with `"cells":[]` so
/// consumers can count rows without a second pass.
fn writeRowEnvelope(
    w: *std.Io.Writer,
    sheet_name: []const u8,
    sheet_idx: usize,
    row_number: u32,
    cells: []const xlsx.Cell,
    include_blanks: bool,
    style_ctx: ?EnvelopeStyleCtx,
    col_offset: u32,
    compact: bool,
    date_types: []const bool,
    error_strings: []const ?[]const u8,
    formula_strings: []const ?[]const u8,
    formula_refs: []const ?xlsx.CellRef,
    uses_1904: bool,
) !void {
    if (compact) {
        try w.print("{{\"kind\":\"row\",\"row\":{d},\"cells\":", .{row_number});
    } else {
        try w.writeAll("{\"kind\":\"row\",\"sheet\":");
        try writeJsonString(w, sheet_name);
        try w.print(",\"sheet_idx\":{d},\"row\":{d},\"cells\":", .{ sheet_idx, row_number });
    }
    try writeEnvelopeCells(
        w,
        cells,
        row_number,
        include_blanks,
        style_ctx,
        col_offset,
        date_types,
        error_strings,
        formula_strings,
        formula_refs,
        uses_1904,
    );
    try w.writeAll("}\n");
}

/// iter59b-3: emit one dict-shape envelope line:
/// `{"kind":"row","sheet":…,"sheet_idx":…,"row":…,"fields":{k:v,…}}\n`.
/// `keys` is one string per header column; `data_cells` is the row's
/// materialised cells, positionally aligned to `keys`. Missing cells
/// (row shorter than `keys`) or empty cells emit `"key": null`; extra
/// cells past `keys.len` are dropped (no key for them). Duplicate keys
/// are emitted as-is — JSON accepts them, the consumer deduplicates.
fn writeRowEnvelopeDict(
    w: *std.Io.Writer,
    sheet_name: []const u8,
    sheet_idx: usize,
    row_number: u32,
    keys: []const []const u8,
    data_cells: []const xlsx.Cell,
    compact: bool,
) !void {
    if (compact) {
        try w.print("{{\"kind\":\"row\",\"row\":{d},\"fields\":{{", .{row_number});
    } else {
        try w.writeAll("{\"kind\":\"row\",\"sheet\":");
        try writeJsonString(w, sheet_name);
        try w.print(",\"sheet_idx\":{d},\"row\":{d},\"fields\":{{", .{ sheet_idx, row_number });
    }
    for (keys, 0..) |k, i| {
        if (i > 0) try w.writeByte(',');
        try writeJsonString(w, k);
        try w.writeByte(':');
        if (i < data_cells.len and data_cells[i] != .empty) {
            try writeJsonCell(w, data_cells[i]);
        } else {
            try w.writeAll("null");
        }
    }
    try w.writeAll("}}\n");
}

/// iter56: emit one NDJSON record for a single cell, matching the
/// `cells` sub-command wire format:
/// `{"kind":"cell","sheet":…,"sheet_idx":…,"ref":…,"row":…,"col":…,"t":…,"v":…}\n`.
///
/// iter59b-4: `.empty` cells are permitted ONLY when the caller opted
/// into `--include-blanks` and materialises them as `t:"blank","v":null`.
/// Without the flag, the caller is still required to skip empties
/// (sparse-by-default cell stream). `style_block` optionally appends
/// `,"style":{…}` when the cell has effective formatting — callers
/// pass null to omit the key entirely; a non-null book + style_idx
/// pair triggers the lookup.
fn writeCell(
    w: *std.Io.Writer,
    sheet_name: []const u8,
    sheet_idx: usize,
    ref: []const u8,
    row: u32,
    col: u32,
    cell: xlsx.Cell,
    style_ctx: ?CellStyleCtx,
    compact: bool,
    date_type: bool,
    error_string: ?[]const u8,
    formula_string: ?[]const u8,
    formula_ref: ?xlsx.CellRef,
    uses_1904: bool,
) !void {
    if (compact) {
        try w.writeAll("{\"kind\":\"cell\",\"ref\":");
        try writeJsonString(w, ref);
        try w.print(",\"row\":{d},\"col\":{d},\"t\":", .{ row, col });
    } else {
        try w.writeAll("{\"kind\":\"cell\",\"sheet\":");
        try writeJsonString(w, sheet_name);
        try w.print(",\"sheet_idx\":{d},\"ref\":", .{sheet_idx});
        try writeJsonString(w, ref);
        try w.print(",\"row\":{d},\"col\":{d},\"t\":", .{ row, col });
    }
    // iter61-b: tag precedence — formula > error > date > primitive.
    // The reader already enforces mutual exclusion (consumeCell
    // clears the error slot when a formula is present, and the date
    // gate skips formula cells), so this ordering is defensive but
    // keeps the rule local. Formula cells emit `formula` /
    // `formula_ref` + optional `cached`; non-formula cells fall
    // through to the existing error / date / primitive paths.
    if (formula_string) |text| {
        try w.writeAll("\"formula\",\"formula\":");
        try writeJsonString(w, text);
        try writeFormulaCached(w, cell);
    } else if (formula_ref) |base| {
        try w.writeAll("\"formula\",");
        try writeFormulaRef(w, base);
        try writeFormulaCached(w, cell);
    } else switch (cell) {
        .empty => try w.writeAll("\"blank\",\"v\":null"),
        else => {
            if (error_string) |literal| {
                try w.writeAll("\"error\",\"v\":");
                try writeJsonString(w, literal);
            } else if (date_type) {
                try w.writeAll("\"date\",\"v\":");
                try writeDateValueAndSerial(w, cellToSerial(cell), uses_1904);
            } else {
                try w.print("\"{s}\",\"v\":", .{envelopeTypeTag(cell)});
                try writeJsonCell(w, cell);
            }
        },
    }
    if (style_ctx) |ctx| if (ctx.style_idx) |sidx| {
        // `writeTerseStyleBlock` returns false when the resolved style
        // has no effective formatting; in that case we omit the field
        // rather than emitting `"style":{}` so consumers can test
        // presence instead of comparing against an empty object.
        // We speculatively write the prefix and patch by writing into
        // a fixed staging buffer would require a second pass. Simpler:
        // the callee emits nothing iff ineffective, so we must decide
        // before writing the prefix. Do the effectiveness check twice
        // (once here, once inside the callee) — cheap lookups.
        if (styleBlockEffective(ctx.book, sidx)) {
            try w.writeAll(",\"style\":");
            _ = try writeTerseStyleBlock(w, ctx.book, sidx);
        }
    };
    try w.writeAll("}\n");
}

/// iter59b-4: context required to emit a per-cell style block. Passed
/// by value; `book` is a const pointer so callers can't accidentally
/// mutate workbook state mid-iteration.
const CellStyleCtx = struct {
    book: *const xlsx.Book,
    /// Null when the source row had no `s=` attribute for this column.
    /// The callee still short-circuits on an ineffective style, so
    /// passing a non-null id for a default-styled cell is fine.
    style_idx: ?u32,
};

/// iter59b-4: fast pre-check mirroring writeTerseStyleBlock's own
/// effectiveness test. Lets the caller decide whether to emit the
/// `,"style":` prefix BEFORE committing to the block — avoids writing
/// a dangling key with nothing after it.
fn styleBlockEffective(book: *const xlsx.Book, style_idx: u32) bool {
    const font = book.cellFont(style_idx);
    const fill = book.cellFill(style_idx);
    const border = book.cellBorder(style_idx);
    const nf = book.numberFormat(style_idx);

    const fill_effective: bool = if (fill) |fl|
        !((std.mem.eql(u8, fl.pattern, "none") or fl.pattern.len == 0) and
            fl.fg_color_argb == null and fl.bg_color_argb == null)
    else
        false;
    const side_empty = struct {
        fn f(s: xlsx.BorderSide) bool {
            return s.style.len == 0 and s.color_argb == null;
        }
    }.f;
    const border_effective: bool = if (border) |b|
        !(side_empty(b.left) and side_empty(b.right) and
            side_empty(b.top) and side_empty(b.bottom) and side_empty(b.diagonal))
    else
        false;
    const font_effective: bool = if (font) |f|
        (f.bold or f.italic or f.color_argb != null)
    else
        false;
    const nf_effective: bool = if (nf) |s|
        !(std.mem.eql(u8, s, "General"))
    else
        false;
    return font_effective or fill_effective or border_effective or nf_effective;
}

/// Legacy emitter — covers the four bare/flat formats. The new
/// envelope format (`.jsonl`) goes through `writeRowEnvelope`, not
/// this function. Calling this with `.jsonl` is a programmer error.
///
/// `col_offset` is the 0-based absolute column of `cells[0]`. It
/// matters only for `legacy_jsonl_dict`, whose keys are column
/// letters and must reflect the TRUE column (not restart from A)
/// when the caller passes a sliced row. Callers outside --range
/// paths pass 0.
fn writeRow(w: *std.Io.Writer, cells: []const xlsx.Cell, fmt: Format, col_offset: u32) !void {
    switch (fmt) {
        .jsonl => unreachable, // envelope path; use writeRowEnvelope
        .legacy_jsonl => {
            try w.writeByte('[');
            for (cells, 0..) |c, i| {
                if (i > 0) try w.writeAll(", ");
                try writeJsonCell(w, c);
            }
            try w.writeAll("]\n");
        },
        .legacy_jsonl_dict => {
            try w.writeByte('{');
            var first = true;
            for (cells, 0..) |c, i| {
                if (c == .empty) continue;
                if (!first) try w.writeAll(", ");
                first = false;
                var col_buf: [8]u8 = undefined;
                const col = colLetter(&col_buf, col_offset + @as(u32, @intCast(i)));
                try w.writeByte('"');
                try w.writeAll(col);
                try w.writeAll("\": ");
                try writeJsonCell(w, c);
            }
            try w.writeAll("}\n");
        },
        .tsv => {
            for (cells, 0..) |c, i| {
                if (i > 0) try w.writeByte('\t');
                switch (c) {
                    .empty => try w.writeAll("\\N"),
                    .string => |s| try writeTsvField(w, s),
                    .integer => |x| try w.print("{d}", .{x}),
                    .number => |f| {
                        if (std.math.isFinite(f)) try w.print("{d}", .{f}) else try w.writeAll("\\N");
                    },
                    .boolean => |b| try w.writeAll(if (b) "true" else "false"),
                }
            }
            try w.writeByte('\n');
        },
        .csv => {
            for (cells, 0..) |c, i| {
                if (i > 0) try w.writeByte(',');
                switch (c) {
                    .empty => {},
                    .string => |s| try writeCsvField(w, s),
                    .integer => |x| try w.print("{d}", .{x}),
                    .number => |f| {
                        if (std.math.isFinite(f)) try w.print("{d}", .{f});
                    },
                    .boolean => |b| try w.writeAll(if (b) "true" else "false"),
                }
            }
            try w.writeByte('\n');
        },
    }
}

// ─── iter60a: process hygiene (signals + exit codes) ────────────────
//
// Three goals, per docs/jq-for-excel.md v4.1:
//  1. `zlsx cells huge.xlsx | head -10` exits cleanly on SIGPIPE — no
//     broken-pipe traceback, no stderr noise, exit 0.
//  2. SIGINT / SIGTERM flush in-flight records, exit 130 / 143. A
//     partial record mid-emission is dropped (not written) so the
//     stream stays valid NDJSON.
//  3. Exit-code table is wired: 0/1/2/3/4/5/130/143 per the doc.
//
// Signal handling is async-signal-safe by construction: the handlers
// only touch three atomic flags and nothing else. No allocation, no
// writers, no stdlib calls beyond atomic stores. Exit codes are routed
// from `main()` at the normal return path.
//
// Mid-record discard contract: every emission loop polls
// `signals.shouldStop()` BEFORE starting a new record. Once the flag
// is set, we return early without writing the opening brace — so
// every line on stdout is a complete, valid NDJSON record.
const signals = struct {
    var stop_streaming = std.atomic.Value(bool).init(false);
    var sigint_fired = std.atomic.Value(bool).init(false);
    var sigterm_fired = std.atomic.Value(bool).init(false);
    var sigpipe_fired = std.atomic.Value(bool).init(false);
    /// M6 (§12.2): `eval` / `recalc` compute their exit codes
    /// commit-aware — a signal that lands after the rename is a
    /// success, not a 130/143 — so once they have mapped an exit code
    /// it is final and `exitCode` below must not re-map it. The
    /// streaming sub-commands never set this and keep the shipped
    /// override.
    var exit_is_final = std.atomic.Value(bool).init(false);

    /// 0.16 types the POSIX signal-handler parameter per platform —
    /// Linux uses a generated `SIG` enum, other targets an int — so
    /// derive it from `Sigaction.handler_fn` instead of hard-coding
    /// `i32`, which only compiled on some targets.
    const SigArg = @typeInfo(@typeInfo(std.posix.Sigaction.handler_fn).pointer.child).@"fn".params[0].type.?;

    fn onSigpipe(_: SigArg) callconv(.c) void {
        sigpipe_fired.store(true, .release);
        stop_streaming.store(true, .release);
    }
    fn onSigint(_: SigArg) callconv(.c) void {
        sigint_fired.store(true, .release);
        stop_streaming.store(true, .release);
    }
    fn onSigterm(_: SigArg) callconv(.c) void {
        sigterm_fired.store(true, .release);
        stop_streaming.store(true, .release);
    }

    fn install() void {
        if (builtin.os.tag == .windows) return;
        const pipe_act: std.posix.Sigaction = .{
            .handler = .{ .handler = onSigpipe },
            .mask = std.posix.sigemptyset(),
            .flags = 0,
        };
        const int_act: std.posix.Sigaction = .{
            .handler = .{ .handler = onSigint },
            .mask = std.posix.sigemptyset(),
            .flags = 0,
        };
        const term_act: std.posix.Sigaction = .{
            .handler = .{ .handler = onSigterm },
            .mask = std.posix.sigemptyset(),
            .flags = 0,
        };
        std.posix.sigaction(std.posix.SIG.PIPE, &pipe_act, null);
        std.posix.sigaction(std.posix.SIG.INT, &int_act, null);
        std.posix.sigaction(std.posix.SIG.TERM, &term_act, null);
    }

    inline fn shouldStop() bool {
        return stop_streaming.load(.acquire);
    }

    /// Map the current signal state to the right exit code per the
    /// design-doc table. Caller has already finished emission — we
    /// only choose the shell status byte. SIGINT beats SIGTERM beats
    /// SIGPIPE in the rare case two fire in the same process lifetime;
    /// the ordering matches the severity the user experiences (Ctrl-C
    /// is user intent, SIGTERM is orderly shutdown, SIGPIPE is normal
    /// pipeline teardown).
    fn exitCode(existing: u8) u8 {
        if (exit_is_final.load(.acquire)) return existing;
        if (sigint_fired.load(.acquire)) return 130;
        if (sigterm_fired.load(.acquire)) return 143;
        if (sigpipe_fired.load(.acquire)) return 0;
        return existing;
    }
};

/// iter60a: classify a top-level CLI error against the exit-code
/// table. Stdout write failures (the pipe peer closed, disk full,
/// etc.) become exit 5 — but only if SIGPIPE didn't fire first, in
/// which case the signal state takes precedence via `signals.exitCode`.
/// Map a top-level runtime error to the design-doc exit-code table.
/// Returns the exit code only; the caller is responsible for printing
/// a human-readable diagnostic when the classification is "unknown"
/// (exit 1) so field failures aren't swallowed silently.
fn classifyTopLevelError(e: anyerror) u8 {
    return switch (e) {
        error.WriteFailed, error.Unexpected, error.BrokenPipe => 5,
        // S1: unreachable in practice — every entry is admitted on the
        // open-time directory walk, so a lazy sheet extraction cannot
        // trip a limit mid-stream — but the mapping exists so the
        // contract holds if a future path decompresses without the walk.
        error.ZipBombSuspected => 4,
        else => 1,
    };
}

/// S1: the one open-time failure that is not exit 2. A decompression
/// limit (`ZipBombSuspected`) means the archive is hostile, not
/// unreadable, and `docs/cli.md` reserved exit 4 for it on every family
/// that opens through the core reader or the editor. Every "cannot
/// open" site routes through here so the read and edit families cannot
/// drift apart.
fn openFailureExit(e: anyerror) u8 {
    return if (e == error.ZipBombSuspected) 4 else 2;
}

pub fn main(init: std.process.Init) u8 {
    // iter60a: the process-hygiene slice. Install SIGPIPE / SIGINT /
    // SIGTERM handlers before any emission can begin so a fast pipe
    // teardown (e.g. `| head -0`) never races the first write.
    signals.install();

    const code = runMain(init) catch |e| blk: {
        const classified = classifyTopLevelError(e);
        // Preserve the diagnostic for unclassified errors so field
        // users still get a handle to file a bug — exit 1 without
        // any stderr is the runtime equivalent of "something broke,
        // good luck figuring out what."
        if (classified == 1) {
            var stderr_buf: [128]u8 = undefined;
            var stderr_file = std.Io.File.stderr().writer(init.io, &stderr_buf);
            const err = &stderr_file.interface;
            err.print("zlsx: {s}\n", .{@errorName(e)}) catch {};
            err.flush() catch {};
        }
        break :blk classified;
    };
    return signals.exitCode(code);
}

/// Process-wide Io for the CLI.
///
/// Zig 0.16 requires an `Io` for every filesystem and stdio call. It is
/// handed to `main` via `std.process.Init` and stashed here so the
/// dispatch switch can reach it without threading a parameter through
/// every command signature. The command functions that actually touch
/// the filesystem still take an `io` parameter — that keeps them
/// testable, since tests never run `runMain` and so never set this.
///
/// Named `proc_io` rather than `io` so it cannot collide with the
/// locally bound `io` in test blocks.
var proc_io: std.Io = undefined;

fn runMain(init: std.process.Init) !u8 {
    // Debug builds use the leak-detecting allocator; release builds use
    // smp_allocator — fast, pure-Zig (no libc dep). smp_allocator asserts
    // !builtin.single_threaded, so single-threaded builds fall back to
    // page_allocator (also pure-Zig, slightly higher per-alloc cost but
    // fine for short-lived CLIs).
    var gpa: std.heap.DebugAllocator(.{}) = .init;
    defer if (builtin.mode == .Debug) {
        _ = gpa.deinit();
    };
    const release_alloc: std.mem.Allocator = if (builtin.single_threaded)
        std.heap.page_allocator
    else
        std.heap.smp_allocator;
    const alloc = if (builtin.mode == .Debug) gpa.allocator() else release_alloc;

    // 0.16 hands main its Io and argv through `std.process.Init`
    // rather than exposing them ambiently — `std.process.argsAlloc` is
    // gone. Take both from there instead of standing up our own
    // Threaded instance.
    proc_io = init.io;

    // toSlice wants an arena — process.Init carries one whose lifetime
    // is the whole process, which is exactly argv's lifetime. Passing
    // the general-purpose allocator instead leaks, since 0.16 has no
    // argsFree counterpart.
    const raw_args = try init.minimal.args.toSlice(init.arena.allocator());

    var stdout_buf: [16 * 1024]u8 = undefined;
    var stdout_file = std.Io.File.stdout().writer(proc_io, &stdout_buf);
    const out = &stdout_file.interface;
    // iter60a: flush on normal exit AND on a signal-triggered stop.
    // Per-record flushes below surface SIGPIPE promptly; this trailing
    // flush is the belt-and-braces for the success path and for the
    // partial last-record case after SIGINT/SIGTERM. Errors are
    // silenced — by the time we're in a defer there's nothing sensible
    // to do with a write failure.
    defer out.flush() catch {};

    var stderr_buf: [4 * 1024]u8 = undefined;
    var stderr_file = std.Io.File.stderr().writer(proc_io, &stderr_buf);
    const err = &stderr_file.interface;
    defer err.flush() catch {};

    // dbx-1: the Databricks family has its own argument grammar (no
    // local workbook positional, env-based auth), so it delegates the
    // whole tail before parseArgs rather than growing the Subcommand
    // scoping matrix.
    if (raw_args.len >= 2 and std.mem.eql(u8, raw_args[1], "dbx")) {
        return try dbx.run(alloc, proc_io, init.minimal.environ, raw_args[2..], out, err);
    }

    // M6 (§12.2): `eval` / `recalc` delegate the whole tail the same
    // way — their argument grammar, stream state machine and exit table
    // live in `formula_cli.zig`. They own their exit mapping completely
    // (commit-aware; never propagate errors), so latch it as final
    // before returning it through `signals.exitCode`.
    if (raw_args.len >= 2 and
        (std.mem.eql(u8, raw_args[1], "eval") or std.mem.eql(u8, raw_args[1], "recalc")))
    {
        const code = formula_cli.run(alloc, proc_io, raw_args[1..], out, err, .{ .sig = .{
            .stop = &signals.stop_streaming,
            .sigint = &signals.sigint_fired,
            .sigterm = &signals.sigterm_fired,
            .sigpipe = &signals.sigpipe_fired,
        } });
        signals.exit_is_final.store(true, .release);
        return code;
    }

    const args = parseArgs(raw_args[1..]) catch |e| switch (e) {
        ArgError.HelpRequested => {
            try writeUsage(out);
            try out.flush();
            return 0;
        },
        ArgError.NoFile => {
            try err.writeAll("zlsx: no input file\n\n");
            try writeUsage(err);
            return 1;
        },
        ArgError.UnknownFlag,
        ArgError.MissingValue,
        ArgError.BadFormat,
        ArgError.BadSheetIndex,
        ArgError.BadArgValue,
        ArgError.SheetArgConflict,
        ArgError.TooManyArgs,
        => {
            try err.print("zlsx: bad arguments ({s})\n\n", .{@errorName(e)});
            try writeUsage(err);
            return 1;
        },
    };

    if (args.deprecated_jsonl_dict) {
        try err.writeAll(
            "zlsx: --format jsonl-dict is deprecated, use --format legacy-jsonl-dict (this alias will be removed in a future release)\n",
        );
        try err.flush();
    }

    // iter-lms-4 follow-up: append-rows uses Editor (mutates the
    // archive) instead of Book (read-only), so dispatch BEFORE the
    // Book open. Returns its own exit code.
    if (args.subcommand == .append_rows) {
        return try runAppendRowsCommand(alloc, proc_io, args, err);
    }
    // iter-cm-4: same dispatch pattern for set-cell — Editor route.
    if (args.subcommand == .set_cell) {
        return try runSetCellCommand(alloc, proc_io, args, err);
    }
    // iter-row-4 / iter-col-5 / iter-sheet-5: structural-edit
    // sub-commands also route through Editor.
    switch (args.subcommand) {
        .insert_row, .delete_row => return try runRowEditCommand(alloc, proc_io, args, err),
        .insert_column, .delete_column => return try runColEditCommand(alloc, proc_io, args, err),
        .add_sheet => return try runAddSheetCommand(alloc, proc_io, args, err),
        .rename_sheet => return try runRenameSheetCommand(alloc, proc_io, args, err),
        .delete_sheet => return try runDeleteSheetCommand(alloc, proc_io, args, err),
        .rename_table_column => return try runRenameTableColumnCommand(alloc, proc_io, args, err),
        .scrub_metadata => return try runScrubMetadataCommand(alloc, proc_io, args, err),
        .embed => return try runEmbedCommand(alloc, proc_io, args, out, err),
        // The legacy --list-sheets flag overrides every read
        // sub-command, this one included: it takes the Book path below.
        .pivots => if (!args.list_sheets) return try runPivotsCommand(alloc, proc_io, args, out, err),
        // S3b: defined names live in xl/workbook.xml, which only the
        // package layer parses — same pre-Book dispatch as `pivots`.
        .defined_names => if (!args.list_sheets) return try runDefinedNamesCommand(alloc, proc_io, args, out, err),
        // S3b slice 3: document properties live in docProps/*.xml,
        // which only the package layer holds — same pre-Book dispatch.
        .doc_props => if (!args.list_sheets) return try runDocPropsCommand(alloc, proc_io, args, out, err),
        // S3b slice 4: drawing anchors — the walkers read drawing parts
        // the reader-only Book has no view of; same pre-Book dispatch.
        .anchors => if (!args.list_sheets) return try runAnchorsCommand(alloc, proc_io, args, out, err),
        else => {},
    }

    // iter-sst-4: dispatch on --sst-lazy.
    var book = if (args.sst_lazy)
        xlsx.Book.openSstLazy(alloc, proc_io, args.file) catch |e| {
            try err.print("zlsx: cannot open '{s}': {s}\n", .{ args.file, @errorName(e) });
            return openFailureExit(e);
        }
    else
        xlsx.Book.open(alloc, proc_io, args.file) catch |e| {
            try err.print("zlsx: cannot open '{s}': {s}\n", .{ args.file, @errorName(e) });
            return openFailureExit(e);
        };
    defer book.deinit();

    if (args.list_sheets) {
        for (book.sheets) |s| {
            if (signals.shouldStop()) return 0;
            try out.writeAll(s.name);
            try out.writeByte('\n');
        }
        try out.flush();
        return 0;
    }

    // iter57/58 sub-commands — no per-sheet selection. meta /
    // list-sheets / styles / sst are workbook-wide; comments /
    // validations / hyperlinks iterate every sheet internally.
    switch (args.subcommand) {
        .meta => {
            // Unix argv is raw bytes; only emit `path` as JSON when
            // valid UTF-8 so the NDJSON line stays parseable. Invalid
            // bytes → JSON null + stderr warning.
            const path_opt: ?[]const u8 = if (std.unicode.utf8ValidateSlice(args.file))
                args.file
            else blk: {
                try err.print(
                    "zlsx: workbook path contains non-UTF-8 bytes; emitting \"path\":null in meta output\n",
                    .{},
                );
                try err.flush();
                break :blk null;
            };
            // Document properties come from the package layer, which
            // the reader-only Book has no view of. The reader accepts
            // some archives the package layer refuses (ZIP data
            // descriptors, for one), so a failure here degrades to
            // `"doc_props": null` rather than failing `meta` outright.
            var dp_wb: ?zlsx_pkg.Workbook = zlsx_pkg.Workbook.open(alloc, proc_io, args.file) catch null;
            defer if (dp_wb) |*w| w.deinit();
            const dp: ?zlsx_pkg.DocProps = if (dp_wb) |*w| (w.docProps() catch null) else null;

            try runMetaCommand(out, &book, path_opt, args.output, dp);
            return 0;
        },
        .list_sheets => {
            try runListSheetsCommand(out, &book, args.output);
            return 0;
        },
        .comments => {
            const filter = resolveSheetFilter(&book, args) catch {
                try err.writeAll("zlsx: sheet not found\n");
                return 3;
            };
            try runCommentsCommand(out, &book, filter, args, args.skip, args.take, args.start_row, args.end_row);
            return 0;
        },
        .validations => {
            const filter = resolveSheetFilter(&book, args) catch {
                try err.writeAll("zlsx: sheet not found\n");
                return 3;
            };
            try runValidationsCommand(out, &book, filter, args, args.skip, args.take);
            return 0;
        },
        .hyperlinks => {
            const filter = resolveSheetFilter(&book, args) catch {
                try err.writeAll("zlsx: sheet not found\n");
                return 3;
            };
            try runHyperlinksCommand(out, &book, filter, args, args.skip, args.take);
            return 0;
        },
        .merges => {
            const filter = resolveSheetFilter(&book, args) catch {
                try err.writeAll("zlsx: sheet not found\n");
                return 3;
            };
            return try runMergesCommand(out, err, &book, filter, args, args.skip, args.take);
        },
        .styles => {
            try runStylesCommand(out, &book, args.skip, args.take);
            return 0;
        },
        .sst => {
            try runSstCommand(out, &book, args.skip, args.take);
            return 0;
        },
        // Editor-route subcommands all dispatched before the Book open.
        .append_rows,
        .set_cell,
        .insert_row,
        .delete_row,
        .insert_column,
        .delete_column,
        .add_sheet,
        .rename_sheet,
        .delete_sheet,
        .rename_table_column,
        .scrub_metadata,
        .embed,
        .pivots,
        .defined_names,
        .doc_props,
        .anchors,
        => unreachable,
        .rows, .cells => {},
    }

    // iter59c: resolve the sheet selection up-front. --all-sheets /
    // --sheet-glob expand to every matching sheet; --sheet / --name
    // narrow to one; default is still sheet 0. Errors stay on the same
    // exit paths as before.
    if (book.sheets.len == 0) {
        try err.writeAll("zlsx: workbook has no sheets\n");
        return 3;
    }
    if (args.sheet_name) |n| {
        var found: bool = false;
        for (book.sheets) |s| {
            if (std.mem.eql(u8, s.name, n)) {
                found = true;
                break;
            }
        }
        if (!found) {
            try err.print("zlsx: no sheet named '{s}'\n", .{n});
            return 3;
        }
    }
    if (args.sheet_index) |idx| {
        if (idx >= book.sheets.len) {
            try err.print("zlsx: sheet index {d} out of range (workbook has {d})\n", .{ idx, book.sheets.len });
            return 3;
        }
    }

    switch (args.subcommand) {
        .rows => try runRowsAcrossSheets(out, &book, args, alloc),
        .cells => try runCellsAcrossSheets(out, &book, args, alloc),
        // Handled by the workbook-scoped early return above.
        .meta,
        .list_sheets,
        .comments,
        .validations,
        .hyperlinks,
        .styles,
        .sst,
        .append_rows,
        .set_cell,
        .insert_row,
        .delete_row,
        .insert_column,
        .delete_column,
        .add_sheet,
        .rename_sheet,
        .delete_sheet,
        .rename_table_column,
        .scrub_metadata,
        .embed,
        .pivots,
        .merges,
        .defined_names,
        .doc_props,
        .anchors,
        => unreachable,
    }
    return 0;
}

/// iter59a: stream-native pagination. `consume()` returns one of
/// three verdicts per candidate record. The counters apply GLOBALLY
/// over the emitted-record stream of a single sub-command run, per
/// the jq-for-excel CLI conventions in docs/jq-for-excel.md.
const Pagination = struct {
    skip: ?usize,
    take: ?usize,
    skipped: usize = 0,
    taken: usize = 0,

    const Verdict = enum { drop, emit, stop };

    fn init(skip: ?usize, take: ?usize) Pagination {
        return .{ .skip = skip, .take = take };
    }

    /// Call once per candidate record before emitting. `.drop` means
    /// advance past this record; `.emit` means emit then mark taken;
    /// `.stop` means --take already satisfied — return early without
    /// emitting anything further.
    fn consume(self: *Pagination) Verdict {
        if (self.take) |t| if (self.taken >= t) return .stop;
        if (self.skip) |s| if (self.skipped < s) {
            self.skipped += 1;
            return .drop;
        };
        self.taken += 1;
        return .emit;
    }
};

/// iter60b: emit the compact-ndjson sheet prologue. Called once per
/// sheet, just before the sheet's first emitted data record in
/// `compact-ndjson` mode. Callers own the "did I emit this sheet's
/// prologue yet?" bookkeeping — typically via a `last_sheet_idx:
/// ?usize` threaded across a cross-sheet run, so pagination can skip a
/// sheet entirely (no prologue leaks for empty/skipped sheets).
fn writeCompactSheetPrologue(w: *std.Io.Writer, sheet_name: []const u8, sheet_idx: usize) !void {
    try w.writeAll("{\"kind\":\"sheet\",\"sheet\":");
    try writeJsonString(w, sheet_name);
    try w.print(",\"sheet_idx\":{d}}}\n", .{sheet_idx});
}

/// iter60c: emit one inline `{"kind":"error",…}` NDJSON record for a
/// non-fatal parse error. `sheet`/`sheet_idx` are optional so workbook-
/// scoped errors can omit them; under compact-ndjson they're dropped
/// from sheet-scoped records too — matching the per-record envelope
/// rules of the surrounding stream. Callers own the decision of which
/// errors to surface vs propagate; this helper is shape-only.
///
/// Per docs/jq-for-excel.md v4.1: pipelines may strip these via
/// `jq 'select(.kind!="error")'` and a run that emits them still exits 0.
fn writeErrorRecord(
    w: *std.Io.Writer,
    sheet_name: ?[]const u8,
    sheet_idx: ?usize,
    scope: []const u8,
    code: []const u8,
    message: []const u8,
) !void {
    // Error records ALWAYS carry sheet/sheet_idx when they're sheet-
    // scoped, even under --output compact-ndjson. Rationale (iter60c
    // P1 follow-up): a malformed sheet can fail BEFORE any per-sheet
    // prologue is written — dropping identity would leave downstream
    // consumers unable to tell which sheet failed. Error records are
    // one per bad sheet, so the identity overhead is negligible.
    try w.writeAll("{\"kind\":\"error\"");
    if (sheet_name) |n| {
        try w.writeAll(",\"sheet\":");
        try writeJsonString(w, n);
    }
    if (sheet_idx) |i| try w.print(",\"sheet_idx\":{d}", .{i});
    try w.writeAll(",\"scope\":");
    try writeJsonString(w, scope);
    try w.writeAll(",\"code\":");
    try writeJsonString(w, code);
    try w.writeAll(",\"message\":");
    try writeJsonString(w, message);
    try w.writeAll("}\n");
}

/// iter60c: per-error human-readable description. Kept tight by design
/// — byte-offset plumbing lives behind a future reader API change. The
/// returned slice is a static literal; safe to embed verbatim in JSON.
fn nonFatalErrorMessage(err: anyerror) []const u8 {
    return switch (err) {
        error.MalformedXml => "malformed sheet XML",
        else => @errorName(err),
    };
}

/// iter59c: single-sheet row driver. Kept for call-site compat with
/// the existing test suite — constructs a fresh Pagination internally.
/// Multi-sheet callers go through `runRowsAcrossSheets` so pagination
/// persists across sheets per the design-doc global-stream semantics.
fn runRowsCommand(
    out: *std.Io.Writer,
    book: *xlsx.Book,
    sheet: xlsx.Sheet,
    sheet_idx: usize,
    format: Format,
    alloc: std.mem.Allocator,
    skip: ?usize,
    take: ?usize,
    start_row: ?u32,
    end_row: ?u32,
    range: ?xlsx.MergeRange,
    header: bool,
    include_blanks: bool,
    with_styles: bool,
) !void {
    var pg = Pagination.init(skip, take);
    try runRowsOnSheet(out, book, sheet, sheet_idx, format, alloc, &pg, start_row, end_row, range, header, include_blanks, with_styles, false);
}

/// iter59c: per-sheet body of `rows`. Takes the Pagination by pointer
/// so cross-sheet drivers can thread one counter through every sheet —
/// `--skip N --take M` slices the concatenated stream, not per sheet.
/// Header state is local to this call (per-sheet reset by design).
///
/// iter60c: wraps `runRowsOnSheetCore`. Same non-fatal-error contract
/// as `runCellsOnSheet` — see there.
fn runRowsOnSheet(
    out: *std.Io.Writer,
    book: *xlsx.Book,
    sheet: xlsx.Sheet,
    sheet_idx: usize,
    format: Format,
    alloc: std.mem.Allocator,
    pg: *Pagination,
    start_row: ?u32,
    end_row: ?u32,
    range: ?xlsx.MergeRange,
    header: bool,
    include_blanks: bool,
    with_styles: bool,
    compact: bool,
) !void {
    // Track whether the compact-mode sheet prologue has been written
    // by the Core. On an error catch below we need to know this so
    // we can emit the prologue FIRST (if not yet written), then the
    // error record without identity — preserves the
    // prologue-carries-identity envelope contract even when a sheet
    // fails before any data record.
    var prologue_emitted: bool = false;
    runRowsOnSheetCore(out, book, sheet, sheet_idx, format, alloc, pg, start_row, end_row, range, header, include_blanks, with_styles, compact, &prologue_emitted) catch |err| switch (err) {
        error.MalformedXml => {
            // iter60c-P2: error records participate in --skip / --take
            // pagination so a later sheet failure doesn't violate the
            // "global over concatenated stream" contract.
            switch (pg.consume()) {
                .drop => return,
                .stop => return,
                .emit => {
                    if (compact and !prologue_emitted) {
                        try writeCompactSheetPrologue(out, sheet.name, sheet_idx);
                        prologue_emitted = true;
                    }
                    // In compact mode the prologue carries identity —
                    // drop sheet/sheet_idx from the error record to
                    // match the data-record envelope contract. In
                    // ndjson mode keep them (no prologue mechanism).
                    const err_sheet: ?[]const u8 = if (compact) null else sheet.name;
                    const err_idx: ?usize = if (compact) null else sheet_idx;
                    try writeErrorRecord(out, err_sheet, err_idx, "sheet", "MalformedXml", nonFatalErrorMessage(err));
                },
            }
            return;
        },
        else => return err,
    };
}

fn runRowsOnSheetCore(
    out: *std.Io.Writer,
    book: *xlsx.Book,
    sheet: xlsx.Sheet,
    sheet_idx: usize,
    format: Format,
    alloc: std.mem.Allocator,
    pg: *Pagination,
    start_row: ?u32,
    end_row: ?u32,
    range: ?xlsx.MergeRange,
    header: bool,
    include_blanks: bool,
    with_styles: bool,
    compact: bool,
    prologue_emitted_out: *bool,
) !void {
    // iter60b: compact-ndjson wire shape emits a `{"kind":"sheet",…}`
    // prologue before the first data record of this sheet. Deferred
    // to just-before-emit so paginated-away or all-blank sheets
    // don't leak a stray prologue with no records behind it. The
    // flag lives in the WRAPPER's stack (iter60c-P2 follow-up) so
    // the catch block can detect an early failure and emit the
    // prologue before the error record.
    // Scoping is enforced at parse time: parseArgs rejects --header on
    // any format other than .jsonl. Reassert here so any accidental
    // future caller that bypasses parseArgs fails loudly in Debug.
    std.debug.assert(!header or format == .jsonl);
    // iter59b-4: --with-styles is envelope-only and header-incompatible.
    // parseArgs enforces both; reassert for offensive-programming parity.
    std.debug.assert(!with_styles or (format == .jsonl and !header));

    var rows = try book.rows(sheet, alloc);
    defer rows.deinit();

    // iter59b-3: owned key strings derived from the header row.
    // Lifetime is this function's scope; row iteration yields fresh
    // cell buffers per row, so we must copy header cell contents out
    // before the next `rows.next()` call reuses the buffer.
    var header_keys: std.ArrayListUnmanaged([]u8) = .empty;
    defer {
        for (header_keys.items) |k| alloc.free(k);
        header_keys.deinit(alloc);
    }
    var header_consumed: bool = !header; // if --header off, skip the dance

    // iter59b-2: --range + --start-row / --end-row take the INTERSECTION
    // on the row axis. The user said "both bounds apply"; the only
    // self-consistent reading is most-restrictive-wins.
    const row_lo: ?u32 = blk: {
        const a = start_row;
        const b = if (range) |r| r.top_left.row else null;
        if (a == null) break :blk b;
        if (b == null) break :blk a;
        break :blk @max(a.?, b.?);
    };
    const row_hi: ?u32 = blk: {
        const a = end_row;
        const b = if (range) |r| r.bottom_right.row else null;
        if (a == null) break :blk b;
        if (b == null) break :blk a;
        break :blk @min(a.?, b.?);
    };

    // Masked buffer for --range on the envelope path: positional
    // contract (cells[i] lives in column i) requires we write `.empty`
    // into out-of-range columns rather than compacting the slice.
    // Only allocated when --range is actually present.
    var masked: std.ArrayListUnmanaged(xlsx.Cell) = .empty;
    defer masked.deinit(alloc);
    // iter59b-4: parallel masked style indices — lives next to `masked`
    // so `masked.items[i]` and `masked_styles.items[i]` stay paired
    // through every view transformation below. Only populated when
    // --range is present AND we're on the envelope-positional path.
    var masked_styles: std.ArrayListUnmanaged(?u32) = .empty;
    defer masked_styles.deinit(alloc);
    // Parallel masked date-type flags — same lockstep invariant as
    // `masked_styles`, wired through `writeRowEnvelope` so the sliced
    // envelope still surfaces `t:"date"` for date-styled numerics.
    var masked_dates: std.ArrayListUnmanaged(bool) = .empty;
    defer masked_dates.deinit(alloc);
    // iter61-c: parallel masked error-string slice. Same lockstep
    // invariant as `masked_styles` / `masked_dates`, wired through
    // `writeRowEnvelope` so the sliced envelope still surfaces
    // `t:"error"` for cells whose source `<c>` was `t="e"`.
    var masked_errors: std.ArrayListUnmanaged(?[]const u8) = .empty;
    defer masked_errors.deinit(alloc);
    // iter61-b: parallel masked formula-string + formula-ref slices.
    // Same lockstep invariant — wired through `writeRowEnvelope` so
    // the sliced envelope still surfaces `t:"formula"` records for
    // cells whose source `<c>` carried `<f>`. The two are mutually
    // exclusive per cell (a cell either has its own formula text or
    // is a shared-formula slave referencing one).
    var masked_formulas: std.ArrayListUnmanaged(?[]const u8) = .empty;
    defer masked_formulas.deinit(alloc);
    var masked_formula_refs: std.ArrayListUnmanaged(?xlsx.CellRef) = .empty;
    defer masked_formula_refs.deinit(alloc);

    while (try rows.next()) |cells| {
        // iter60a: outer-loop stop-poll — mid-record discard contract.
        // Once flagged, we bail before touching the row's cells.
        if (signals.shouldStop()) return;
        const row_number = rows.currentRowNumber();
        // Row bounds run BEFORE pagination (design doc v4.1).
        if (row_lo) |s| if (row_number < s) continue;
        // OOXML rows are monotonic — once past the upper bound, no
        // more records in this sheet's stream can satisfy it.
        if (row_hi) |e| if (row_number > e) break;

        // Per design doc: a row is emitted iff at least one cell is
        // inside the rectangle. The envelope path masks out-of-col
        // cells to .empty so the positional cells[i] == col-i contract
        // holds. Flat formats (csv/tsv/legacy-jsonl/legacy-jsonl-dict)
        // instead SLICE to the in-range column span and pass the
        // absolute col offset so legacy-jsonl-dict keys stay truthful
        // (`{"XFD": …}`, not `{"A": …}` for a ranged XFD1:XFD10).
        const EmitView = struct {
            cells: []const xlsx.Cell,
            /// iter59b-4: parallel style-index slice — may be shorter
            /// than `cells` when --with-styles is off (empty slice) or
            /// when the row's raw styleIndices is shorter than the
            /// masked cells width (callee guards positional reads).
            style_indices: []const ?u32,
            /// iter61-a: parallel date-type flag slice. Aligned to
            /// `cells` the same way `style_indices` is — a `true` slot
            /// means the matching cell is a date-styled numeric serial
            /// and should be emitted as `t:"date"` with a `serial`
            /// sidecar. Empty when the row has no date cells.
            date_types: []const bool,
            /// iter61-c: parallel error-string slice. Aligned to `cells`;
            /// a non-null slot means the matching cell was `t="e"` in
            /// OOXML and should be emitted as `t:"error"` with the
            /// literal string as `v`. Empty when the row has no error
            /// cells.
            error_strings: []const ?[]const u8,
            /// iter61-b: parallel formula-text slice. Aligned to
            /// `cells`; a non-null slot means the cell carried its
            /// own `<f>` body (stand-alone formula, shared-formula
            /// base, or array-formula base) and should be emitted
            /// as `t:"formula"` with `formula:<text>` + optional
            /// `cached:<v>`. Empty when the row has no formula
            /// bases.
            formula_strings: []const ?[]const u8,
            /// iter61-b: parallel shared-formula base-ref slice.
            /// Aligned to `cells`; a non-null slot means the cell
            /// is a shared-formula slave (`<f t="shared" si="N"/>`)
            /// and should be emitted as `t:"formula"` with
            /// `formula_ref:"<A1>"` + optional `cached:<v>`. Empty
            /// when the row has no slave cells.
            formula_refs: []const ?xlsx.CellRef,
            col_offset: u32,
            any_non_empty: bool,
        };
        const raw_styles = rows.styleIndices();
        const raw_dates = rows.dateTypes();
        const raw_errors = rows.errorStrings();
        const raw_formulas = rows.formulaStrings();
        const raw_formula_refs = rows.formulaRefs();
        // Unified slice view for --range: span exactly [tl.col..br.col]
        // with padded empties on sparse rows. writeRowEnvelope now
        // takes col_offset and each cell record carries absolute col
        // explicitly, so the envelope doesn't need positional
        // cells[i]==col-i alignment anymore. style_indices slices in
        // parallel so --with-styles still reaches the right per-cell
        // metadata.
        const view: EmitView = if (range) |r| blk: {
            const range_width: usize = @as(usize, r.bottom_right.col) - r.top_left.col + 1;
            masked.clearRetainingCapacity();
            masked_styles.clearRetainingCapacity();
            masked_dates.clearRetainingCapacity();
            masked_errors.clearRetainingCapacity();
            masked_formulas.clearRetainingCapacity();
            masked_formula_refs.clearRetainingCapacity();
            try masked.ensureTotalCapacity(alloc, range_width);
            try masked_styles.ensureTotalCapacity(alloc, range_width);
            try masked_dates.ensureTotalCapacity(alloc, range_width);
            try masked_errors.ensureTotalCapacity(alloc, range_width);
            try masked_formulas.ensureTotalCapacity(alloc, range_width);
            try masked_formula_refs.ensureTotalCapacity(alloc, range_width);
            var any = false;
            var col: u32 = r.top_left.col;
            while (col <= r.bottom_right.col) : (col += 1) {
                const src_idx: usize = col;
                if (src_idx < cells.len) {
                    masked.appendAssumeCapacity(cells[src_idx]);
                    masked_styles.appendAssumeCapacity(
                        if (src_idx < raw_styles.len) raw_styles[src_idx] else null,
                    );
                    masked_dates.appendAssumeCapacity(
                        src_idx < raw_dates.len and raw_dates[src_idx],
                    );
                    masked_errors.appendAssumeCapacity(
                        if (src_idx < raw_errors.len) raw_errors[src_idx] else null,
                    );
                    masked_formulas.appendAssumeCapacity(
                        if (src_idx < raw_formulas.len) raw_formulas[src_idx] else null,
                    );
                    masked_formula_refs.appendAssumeCapacity(
                        if (src_idx < raw_formula_refs.len) raw_formula_refs[src_idx] else null,
                    );
                    // iter61-b: a formula cell with cached `.empty`
                    // (formula-only, no <v>) still warrants emission
                    // — it's not a structural blank. Keep it counted
                    // as non-empty so the sliced row isn't dropped.
                    const has_formula = (src_idx < raw_formulas.len and raw_formulas[src_idx] != null) or
                        (src_idx < raw_formula_refs.len and raw_formula_refs[src_idx] != null);
                    if (cells[src_idx] != .empty or has_formula) any = true;
                } else {
                    masked.appendAssumeCapacity(.empty);
                    masked_styles.appendAssumeCapacity(null);
                    masked_dates.appendAssumeCapacity(false);
                    masked_errors.appendAssumeCapacity(null);
                    masked_formulas.appendAssumeCapacity(null);
                    masked_formula_refs.appendAssumeCapacity(null);
                }
            }
            break :blk .{
                .cells = masked.items,
                .style_indices = masked_styles.items,
                .date_types = masked_dates.items,
                .error_strings = masked_errors.items,
                .formula_strings = masked_formulas.items,
                .formula_refs = masked_formula_refs.items,
                .col_offset = r.top_left.col,
                .any_non_empty = any,
            };
        } else .{
            .cells = cells,
            .style_indices = raw_styles,
            .date_types = raw_dates,
            .error_strings = raw_errors,
            .formula_strings = raw_formulas,
            .formula_refs = raw_formula_refs,
            .col_offset = 0,
            .any_non_empty = true,
        };

        // Skip all-blank rows by default. --include-blanks preserves
        // them, but ONLY on the envelope path — on --header the blank
        // row would poison the key set with `col_*` placeholders, and
        // on flat formats --include-blanks is a documented no-op so a
        // preserved blank row would leak extra empty lines. The
        // envelope path is the only shape where blank rows carry
        // useful `t:"blank"` cell records.
        const preserve_blank = include_blanks and format == .jsonl and !header;
        if (!view.any_non_empty and !preserve_blank) continue;

        // iter59b-3: the header row lives BEFORE pagination so --skip N
        // counts N *data* records. The header cells are captured here
        // and the row itself is swallowed (no envelope emitted).
        if (!header_consumed) {
            try captureHeaderKeys(&header_keys, alloc, view.cells, view.col_offset);
            header_consumed = true;
            continue;
        }

        switch (pg.consume()) {
            .drop => continue,
            .stop => return,
            .emit => {},
        }
        if (compact and !prologue_emitted_out.*) {
            try writeCompactSheetPrologue(out, sheet.name, sheet_idx);
            prologue_emitted_out.* = true;
        }
        if (header) {
            try writeRowEnvelopeDict(out, sheet.name, sheet_idx, row_number, header_keys.items, view.cells, compact);
        } else switch (format) {
            .jsonl => {
                const style_ctx: ?EnvelopeStyleCtx = if (with_styles)
                    .{ .book = book, .style_indices = view.style_indices }
                else
                    null;
                try writeRowEnvelope(
                    out,
                    sheet.name,
                    sheet_idx,
                    row_number,
                    view.cells,
                    include_blanks,
                    style_ctx,
                    view.col_offset,
                    compact,
                    view.date_types,
                    view.error_strings,
                    view.formula_strings,
                    view.formula_refs,
                    book.uses_1904_epoch,
                );
            },
            // iter59b-4: flat formats are shape-neutral w.r.t. both
            // --include-blanks (they serialise empties per their own
            // convention) and --with-styles (no place to put metadata).
            // parseArgs rejects --with-styles on flat formats; allow
            // --include-blanks through as a documented no-op so scripts
            // can set it unconditionally.
            else => try writeRow(out, view.cells, format, view.col_offset),
        }
    }
}

/// iter59b-3: derive one owned key string per header cell. String
/// headers pass through verbatim; numeric/boolean headers are
/// stringified via bufPrint; empty headers become `"col_<letter>"`
/// so consumers can still reference the column. `col_offset` is the
/// absolute 0-based column of cells[0] — matters when --range
/// produced a sliced view so fallback labels reflect the true column.
fn captureHeaderKeys(
    keys: *std.ArrayListUnmanaged([]u8),
    alloc: std.mem.Allocator,
    cells: []const xlsx.Cell,
    col_offset: u32,
) !void {
    // Caller owns the list; clear so re-capture in a future multi-sheet
    // mode stays correct even though today we only ever fill once.
    for (keys.items) |k| alloc.free(k);
    keys.clearRetainingCapacity();
    try keys.ensureTotalCapacity(alloc, cells.len);

    var scratch: [64]u8 = undefined;
    for (cells, 0..) |c, i| {
        const absolute_col: u32 = col_offset + @as(u32, @intCast(i));
        const key: []u8 = switch (c) {
            .empty => blk: {
                var letter_buf: [8]u8 = undefined;
                const letters = colLetter(&letter_buf, absolute_col);
                break :blk try std.fmt.allocPrint(alloc, "col_{s}", .{letters});
            },
            .string => |s| try alloc.dupe(u8, s),
            .integer => |x| blk: {
                const s = std.fmt.bufPrint(&scratch, "{d}", .{x}) catch unreachable;
                break :blk try alloc.dupe(u8, s);
            },
            .number => |f| blk: {
                const s = if (std.math.isFinite(f))
                    std.fmt.bufPrint(&scratch, "{d}", .{f}) catch unreachable
                else
                    // Non-finite headers are a pathological input; fall
                    // back to the column-letter placeholder rather than
                    // emitting "nan" which collides across columns.
                    std.fmt.bufPrint(&scratch, "col_{d}", .{absolute_col + 1}) catch unreachable;
                break :blk try alloc.dupe(u8, s);
            },
            .boolean => |b| try alloc.dupe(u8, if (b) "true" else "false"),
        };
        keys.appendAssumeCapacity(key);
    }
}

/// iter56: stream one NDJSON record per non-empty cell of the selected
/// sheet. Empty cells are suppressed (matches envelope semantics on
/// the rows path). `--format` is intentionally ignored here — the
/// `cells` sub-command has a single fixed wire shape.
///
/// iter59c: single-sheet entry kept for test-call compat. Multi-sheet
/// drivers use `runCellsOnSheet` directly so Pagination persists across
/// sheets (cross-sheet --skip / --take slice the concatenated stream).
fn runCellsCommand(
    out: *std.Io.Writer,
    book: *xlsx.Book,
    sheet: xlsx.Sheet,
    sheet_idx: usize,
    alloc: std.mem.Allocator,
    skip: ?usize,
    take: ?usize,
    start_row: ?u32,
    end_row: ?u32,
    range: ?xlsx.MergeRange,
    include_blanks: bool,
    with_styles: bool,
) !void {
    var pg = Pagination.init(skip, take);
    try runCellsOnSheet(out, book, sheet, sheet_idx, alloc, &pg, start_row, end_row, range, include_blanks, with_styles, false);
}

/// iter59c: per-sheet cell emitter — takes Pagination by pointer so a
/// cross-sheet driver can thread the same counter across every sheet.
///
/// iter60c: wraps `runCellsOnSheetCore`. Non-fatal parse errors (today:
/// `error.MalformedXml`) are caught at sheet boundary and surfaced as
/// an inline `{"kind":"error",…}` record per docs/jq-for-excel.md v4.1.
/// Other errors propagate unchanged so resource / I/O failures stay
/// fatal.
fn runCellsOnSheet(
    out: *std.Io.Writer,
    book: *xlsx.Book,
    sheet: xlsx.Sheet,
    sheet_idx: usize,
    alloc: std.mem.Allocator,
    pg: *Pagination,
    start_row: ?u32,
    end_row: ?u32,
    range: ?xlsx.MergeRange,
    include_blanks: bool,
    with_styles: bool,
    compact: bool,
) !void {
    var prologue_emitted: bool = false;
    runCellsOnSheetCore(out, book, sheet, sheet_idx, alloc, pg, start_row, end_row, range, include_blanks, with_styles, compact, &prologue_emitted) catch |err| switch (err) {
        error.MalformedXml => {
            // iter60c-P2 follow-up: error records participate in
            // --skip / --take AND preserve compact-ndjson's
            // prologue-carries-identity envelope. See the matching
            // note in runRowsOnSheet.
            switch (pg.consume()) {
                .drop => return,
                .stop => return,
                .emit => {
                    if (compact and !prologue_emitted) {
                        try writeCompactSheetPrologue(out, sheet.name, sheet_idx);
                        prologue_emitted = true;
                    }
                    const err_sheet: ?[]const u8 = if (compact) null else sheet.name;
                    const err_idx: ?usize = if (compact) null else sheet_idx;
                    try writeErrorRecord(out, err_sheet, err_idx, "sheet", "MalformedXml", nonFatalErrorMessage(err));
                },
            }
            return;
        },
        else => return err,
    };
}

fn runCellsOnSheetCore(
    out: *std.Io.Writer,
    book: *xlsx.Book,
    sheet: xlsx.Sheet,
    sheet_idx: usize,
    alloc: std.mem.Allocator,
    pg: *Pagination,
    start_row: ?u32,
    end_row: ?u32,
    range: ?xlsx.MergeRange,
    include_blanks: bool,
    with_styles: bool,
    compact: bool,
    prologue_emitted_out: *bool,
) !void {
    // See runCellsOnSheet for why the flag lives in the wrapper's
    // stack — the catch block needs to detect an early failure and
    // emit the prologue itself (iter60c-P2 follow-up).
    var rows = try book.rows(sheet, alloc);
    defer rows.deinit();

    // iter59b-2: intersect --range row bounds with --start-row / --end-row.
    const row_lo: ?u32 = blk: {
        const a = start_row;
        const b = if (range) |r| r.top_left.row else null;
        if (a == null) break :blk b;
        if (b == null) break :blk a;
        break :blk @max(a.?, b.?);
    };
    const row_hi: ?u32 = blk: {
        const a = end_row;
        const b = if (range) |r| r.bottom_right.row else null;
        if (a == null) break :blk b;
        if (b == null) break :blk a;
        break :blk @min(a.?, b.?);
    };

    while (try rows.next()) |cells| {
        if (signals.shouldStop()) return;
        const row_number = rows.currentRowNumber();
        if (row_lo) |s| if (row_number < s) continue;
        if (row_hi) |e| if (row_number > e) break;
        const raw_styles = rows.styleIndices();
        const raw_dates = rows.dateTypes();
        const raw_errors = rows.errorStrings();
        const raw_formulas = rows.formulaStrings();
        const raw_formula_refs = rows.formulaRefs();
        for (cells, 0..) |c, i| {
            // iter61-b: a formula cell with cached `.empty` is still
            // a formula record and must not be skipped — `<c><f>…</f></c>`
            // (no <v>) lives between empty and primitive. Bypass the
            // skip-empty rule when either formula side channel is set.
            const has_formula = (i < raw_formulas.len and raw_formulas[i] != null) or
                (i < raw_formula_refs.len and raw_formula_refs[i] != null);
            // iter59b-4: --include-blanks flips the empty-skip into
            // emit-as-blank. Without the flag, the old sparse-by-default
            // behaviour holds.
            if (c == .empty and !include_blanks and !has_formula) continue;
            if (range) |r| {
                const col: u32 = @intCast(i);
                if (col < r.top_left.col or col > r.bottom_right.col) continue;
            }
            // iter60a: poll BEFORE pg.consume()+write so a signal
            // during a hot inner loop doesn't leave a half-written
            // record on stdout. Atomic acquire-load per candidate is
            // ~free relative to the xlsx parse cost.
            if (signals.shouldStop()) return;

            switch (pg.consume()) {
                .drop => continue,
                .stop => return,
                .emit => {},
            }

            var col_buf: [8]u8 = undefined;
            const letters = colLetter(&col_buf, i);
            var ref_buf: [16]u8 = undefined;
            const ref = std.fmt.bufPrint(&ref_buf, "{s}{d}", .{ letters, row_number }) catch unreachable;

            const style_ctx: ?CellStyleCtx = if (with_styles) blk: {
                const sidx: ?u32 = if (i < raw_styles.len) raw_styles[i] else null;
                break :blk .{ .book = book, .style_idx = sidx };
            } else null;

            if (compact and !prologue_emitted_out.*) {
                try writeCompactSheetPrologue(out, sheet.name, sheet_idx);
                prologue_emitted_out.* = true;
            }
            // iter61-b: same precedence ordering as writeCell —
            // formula > error > date. The raw side channels enforce
            // mutual exclusion at the reader, so this is defensive
            // bookkeeping for callers that pass arbitrary slices.
            const fmla_str: ?[]const u8 = if (i < raw_formulas.len) raw_formulas[i] else null;
            const fmla_ref: ?xlsx.CellRef = if (i < raw_formula_refs.len) raw_formula_refs[i] else null;
            const err_str: ?[]const u8 =
                if (fmla_str == null and fmla_ref == null and i < raw_errors.len) raw_errors[i] else null;
            const is_date: bool =
                fmla_str == null and fmla_ref == null and err_str == null and
                i < raw_dates.len and raw_dates[i];
            try writeCell(
                out,
                sheet.name,
                sheet_idx,
                ref,
                row_number,
                @intCast(i + 1),
                c,
                style_ctx,
                compact,
                is_date,
                err_str,
                fmla_str,
                fmla_ref,
                book.uses_1904_epoch,
            );
            // Per docs/jq-for-excel.md v4.1: "every record is written
            // with an explicit newline + flush on stdout." The flush
        }
    }
}

/// iter59c: cross-sheet predicate for `cells` / `rows`. Centralises
/// the 4-way selection matrix (--sheet / --name / --all-sheets /
/// --sheet-glob / default=first) so both drivers stay in lockstep.
/// Returns true iff the sheet at (name, idx) is in the selection.
/// Assumes parseArgs already rejected mutually-exclusive combinations
/// and main() already bounds-checked --sheet / --name against the book.
fn sheetSelectedForCellsRows(args: Args, sheet_name: []const u8, sheet_idx: usize) bool {
    if (args.sheet_index) |idx| return sheet_idx == idx;
    if (args.sheet_name) |n| return std.mem.eql(u8, sheet_name, n);
    if (args.sheet_glob) |pat| return globMatch(pat, sheet_name);
    if (args.all_sheets) return true;
    return sheet_idx == 0;
}

/// iter59c: multi-sheet driver for the `cells` sub-command. Walks the
/// workbook once, emitting through `runCellsOnSheet` for every sheet
/// that matches the selector. Pagination lives HERE (not per sheet) so
/// `--skip N --take M` slices the concatenated cross-sheet stream per
/// docs/jq-for-excel.md v4.1: "--skip 1000 --take 500 takes records
/// 1001-1500 across the full cross-sheet stream, not per sheet."
fn runCellsAcrossSheets(
    out: *std.Io.Writer,
    book: *xlsx.Book,
    args: Args,
    alloc: std.mem.Allocator,
) !void {
    var pg = Pagination.init(args.skip, args.take);
    for (book.sheets, 0..) |s, sheet_idx| {
        if (signals.shouldStop()) return;
        if (!sheetSelectedForCellsRows(args, s.name, sheet_idx)) continue;
        // Short-circuit once --take is satisfied — checked BEFORE
        // opening the next sheet's row stream to avoid useless I/O.
        if (args.take) |t| if (pg.taken >= t) return;
        try runCellsOnSheet(
            out,
            book,
            s,
            sheet_idx,
            alloc,
            &pg,
            args.start_row,
            args.end_row,
            args.range,
            args.include_blanks,
            args.with_styles,
            args.output == .compact_ndjson,
        );
    }
    // iter60a-P1 follow-up: final explicit flush so stdout write
    // failures (ENOSPC, closed pipe on completion, etc.) surface as
    // error.WriteFailed → classifyTopLevelError → exit 5 instead of
    // being silently swallowed by the defer-catch-swallow in runMain.
    try out.flush();
}

/// iter59c: multi-sheet driver for `rows`. Same cross-sheet pagination
/// contract as `runCellsAcrossSheets`. `--header` is per-sheet by
/// design (each sheet's first in-bounds row becomes that sheet's keys)
/// — the header state lives inside `runRowsOnSheet`, so calling it
/// once per sheet naturally resets keys between sheets.
fn runRowsAcrossSheets(
    out: *std.Io.Writer,
    book: *xlsx.Book,
    args: Args,
    alloc: std.mem.Allocator,
) !void {
    var pg = Pagination.init(args.skip, args.take);
    for (book.sheets, 0..) |s, sheet_idx| {
        if (signals.shouldStop()) return;
        if (!sheetSelectedForCellsRows(args, s.name, sheet_idx)) continue;
        if (args.take) |t| if (pg.taken >= t) return;
        try runRowsOnSheet(
            out,
            book,
            s,
            sheet_idx,
            args.format,
            alloc,
            &pg,
            args.start_row,
            args.end_row,
            args.range,
            args.header,
            args.include_blanks,
            args.with_styles,
            args.output == .compact_ndjson,
        );
    }
    try out.flush();
}

/// iter57: emit the workbook record followed by one sheet record per
/// sheet. Fields deliberately limited to ones that are O(1) over the
/// reader APIs Book already exposes — `rows` / `cols` / `first_cell` /
/// `last_cell` / `format_version` are follow-up work (they need
/// sheet-iteration or version plumbing) per the iter57 scope note.
fn runMetaCommand(
    out: *std.Io.Writer,
    book: *const xlsx.Book,
    path: ?[]const u8,
    output: OutputMode,
    /// Document properties, or null when the package layer could not
    /// open the archive. The reader accepts some archives the Editor
    /// refuses (ZIP data descriptors, for one), so `meta` must still
    /// work without them rather than fail the whole command.
    doc_props: ?zlsx_pkg.DocProps,
) !void {
    // Hidden-sheet tally. Exposed as scalars so a caller can gate on it
    // with jq alone rather than reducing over the sheets array.
    var hidden_count: usize = 0;
    var very_hidden_count: usize = 0;
    for (book.sheets) |s| switch (s.state) {
        .hidden => hidden_count += 1,
        .very_hidden => very_hidden_count += 1,
        .visible => {},
    };
    // Workbook-level `has_comments` is the OR across every sheet —
    // saves callers a reduce step when they only want "does this file
    // have any comments at all?".
    var any_comments = false;
    for (book.sheets) |s| {
        if (book.comments(s).len != 0) {
            any_comments = true;
            break;
        }
    }

    if (output == .pretty_json) {
        // iter60b: collapse workbook + per-sheet records into one
        // 2-space-indented JSON object. The scalar sheet count is
        // renamed from `sheets` to `sheets_count` in this mode ONLY
        // so it doesn't collide with the `sheets: [...]` array that
        // carries the per-sheet records. NDJSON output keeps the
        // original `sheets:N` scalar for back-compat (see below).
        try out.writeAll("{\n");
        try out.writeAll("  \"kind\": \"workbook\",\n");
        try out.writeAll("  \"path\": ");
        if (path) |p| try writeJsonString(out, p) else try out.writeAll("null");
        try out.print(
            ",\n  \"sheets_count\": {d},\n",
            .{book.sheets.len},
        );
        try out.print(
            "  \"sst\": {{\"count\": {d}, \"rich\": {d}}},\n",
            .{ book.sharedStringsCount(), book.rich_runs_by_sst_idx.count() },
        );
        try out.print(
            "  \"has_styles\": {s},\n  \"has_theme\": {s},\n  \"has_comments\": {s},\n",
            .{
                if (book.styles_xml != null) "true" else "false",
                if (book.theme_xml != null) "true" else "false",
                if (any_comments) "true" else "false",
            },
        );
        try out.print(
            "  \"hidden_sheet_count\": {d},\n  \"very_hidden_sheet_count\": {d},\n",
            .{ hidden_count, very_hidden_count },
        );
        try writeDocPropsPretty(out, doc_props);
        try out.writeAll("  \"sheets\": [");
        if (book.sheets.len == 0) {
            try out.writeAll("]\n}\n");
            try out.flush();
            return;
        }
        try out.writeByte('\n');
        for (book.sheets, 0..) |s, i| {
            if (signals.shouldStop()) return;
            const sheet_has_comments = book.comments(s).len != 0;
            try out.writeAll("    {\"kind\": \"sheet\", \"sheet\": ");
            try writeJsonString(out, s.name);
            try out.print(
                ", \"sheet_idx\": {d}, \"state\": \"{s}\", \"has_comments\": {s}}}",
                .{ i, s.state.toString(), if (sheet_has_comments) "true" else "false" },
            );
            if (i + 1 < book.sheets.len) try out.writeByte(',');
            try out.writeByte('\n');
        }
        try out.writeAll("  ]\n}\n");
        try out.flush();
        return;
    }

    // `path` is null when the caller detected non-UTF-8 bytes in the
    // original argv — emit JSON `null` so the NDJSON line stays
    // parseable. main() has already logged the reason to stderr.
    try out.writeAll("{\"kind\":\"workbook\",\"path\":");
    if (path) |p| try writeJsonString(out, p) else try out.writeAll("null");
    try out.print(
        ",\"sheets\":{d},\"sst\":{{\"count\":{d},\"rich\":{d}}}",
        .{ book.sheets.len, book.sharedStringsCount(), book.rich_runs_by_sst_idx.count() },
    );
    try out.print(
        ",\"has_styles\":{s},\"has_theme\":{s},\"has_comments\":{s}",
        .{
            if (book.styles_xml != null) "true" else "false",
            if (book.theme_xml != null) "true" else "false",
            if (any_comments) "true" else "false",
        },
    );
    try out.print(
        ",\"hidden_sheet_count\":{d},\"very_hidden_sheet_count\":{d}",
        .{ hidden_count, very_hidden_count },
    );
    try writeDocPropsCompact(out, doc_props);
    try out.writeAll("}\n");

    // iter60b-P1: compact mode is a documented no-op for meta's
    // per-sheet records. Stripping sheet/sheet_idx here would orphan
    // the records — meta has no sheet-prologue mechanism to share the
    // identifier across multiple records (unlike the sheet-scoped
    // sub-commands whose prologue carries the context). Each per-
    // sheet record IS the identifier.
    for (book.sheets, 0..) |s, i| {
        if (signals.shouldStop()) return;
        const sheet_has_comments = book.comments(s).len != 0;
        try out.writeAll("{\"kind\":\"sheet\",\"sheet\":");
        try writeJsonString(out, s.name);
        try out.print(
            ",\"sheet_idx\":{d},\"state\":\"{s}\",\"has_comments\":{s}}}\n",
            .{ i, s.state.toString(), if (sheet_has_comments) "true" else "false" },
        );
    }
    try out.flush();
}

/// The doc-props wire fields, in the order `Editor.doc_props` and
/// `meta`'s `doc_props` object list them. Every JSON key equals the
/// `zlsx_pkg.DocProps` field it reads, so `@field` walks this one
/// table for the `meta` object, the `doc-props` record and its UTF-8
/// floor — three emissions that cannot disagree on order or spelling.
const doc_prop_fields = [_][]const u8{
    "creator",
    "last_modified_by",
    "title",
    "subject",
    "description",
    "keywords",
    "category",
    "created",
    "modified",
    "revision",
    "company",
    "manager",
    "application",
    "hyperlink_base",
};

/// Emit the `doc_props` object for `meta --output pretty-json`.
///
/// Absent parts render as `null` rather than being omitted, so a
/// consumer can distinguish "zlsx could not read them" from "the
/// workbook genuinely has none" without a schema lookup.
fn writeDocPropsPretty(out: *std.Io.Writer, props: ?zlsx_pkg.DocProps) !void {
    const dp = props orelse {
        try out.writeAll("  \"doc_props\": null,\n");
        return;
    };
    try out.writeAll("  \"doc_props\": {\n");
    inline for (doc_prop_fields) |name| {
        try writeDocPropField(out, name, @field(dp, name), "    ", true);
    }
    try out.print(
        "    \"has_custom_properties\": {s}\n  }},\n",
        .{if (dp.has_custom_properties) "true" else "false"},
    );
}

/// Same object, single-line, for the NDJSON envelope.
fn writeDocPropsCompact(out: *std.Io.Writer, props: ?zlsx_pkg.DocProps) !void {
    const dp = props orelse {
        try out.writeAll(",\"doc_props\":null");
        return;
    };
    try out.writeAll(",\"doc_props\":{");
    inline for (doc_prop_fields) |name| {
        try writeDocPropField(out, name, @field(dp, name), "", false);
    }
    try out.print(
        "\"has_custom_properties\":{s}}}",
        .{if (dp.has_custom_properties) "true" else "false"},
    );
}

/// S3b: `zlsx doc-props` — the document-properties field set as one
/// `{"kind":"doc_props",…}` record: every `docProps/core.xml` /
/// `app.xml` field the typed view models (`null` when absent, text as
/// stored — `meta`'s values; Python's `Editor.doc_props` diverges only
/// by the C boundary's conventions, mapping a present-but-empty
/// element to None and replacing malformed UTF-8) plus
/// `has_custom_properties`. Routes through the package layer like
/// `pivots`; a workbook with no docProps parts is a record of nulls —
/// the absence itself is the one fact the line reports. A field value
/// that is not UTF-8 refuses the whole command (exit 2) with nothing
/// written: every read and validation failure lands before the first
/// byte (only a stdout I/O failure mid-line can truncate the record,
/// as on any streaming command). Contract in docs/cli.md, "doc-props".
fn runDocPropsCommand(
    alloc: std.mem.Allocator,
    io: std.Io,
    args: Args,
    out: *std.Io.Writer,
    err: *std.Io.Writer,
) !u8 {
    var wb = zlsx_pkg.Workbook.open(alloc, io, args.file) catch |e| {
        try err.print("zlsx: cannot open '{s}': {s}\n", .{ args.file, @errorName(e) });
        try err.flush();
        return openFailureExit(e);
    };
    defer wb.deinit();
    const dp = wb.docProps() catch |e| {
        try err.print("zlsx: cannot read document properties in '{s}': {s}\n", .{ args.file, @errorName(e) });
        try err.flush();
        return 2;
    };
    // Every present value is validated before the first byte of the
    // record, so no read or validation failure can half-write it (a
    // stdout I/O failure mid-line can, as on any streaming command).
    // The values pass through writeJsonString verbatim, so malformed
    // UTF-8 here would be invalid NDJSON under exit 0.
    inline for (doc_prop_fields) |name| {
        if (@field(dp, name)) |v| {
            if (!std.unicode.utf8ValidateSlice(v)) {
                try err.print("zlsx: cannot read document properties in '{s}': {s} is not UTF-8\n", .{ args.file, name });
                try err.flush();
                return 2;
            }
        }
    }
    try out.writeAll("{\"kind\":\"doc_props\"");
    inline for (doc_prop_fields) |name| {
        try out.writeAll(",\"" ++ name ++ "\":");
        if (@field(dp, name)) |v| try writeJsonString(out, v) else try out.writeAll("null");
    }
    try out.print(
        ",\"has_custom_properties\":{s}}}\n",
        .{if (dp.has_custom_properties) "true" else "false"},
    );
    try out.flush();
    return 0;
}

/// One `"key": value` pair. Null fields emit JSON `null` so the object
/// shape is stable across workbooks — consumers can index without
/// existence checks.
fn writeDocPropField(
    out: *std.Io.Writer,
    key: []const u8,
    value: ?[]const u8,
    indent: []const u8,
    pretty: bool,
) !void {
    try out.writeAll(indent);
    try out.writeByte('"');
    try out.writeAll(key);
    try out.writeAll(if (pretty) "\": " else "\":");
    if (value) |v| try writeJsonString(out, v) else try out.writeAll("null");
    try out.writeAll(",");
    if (pretty) try out.writeByte('\n');
}

/// `zlsx scrub-metadata <in.xlsx> --out <clean.xlsx>`
///
/// Strips authorship metadata and saves. Everything else — every cell,
/// every untouched part — flows through the Editor's byte-preserving
/// save path unchanged.
fn runScrubMetadataCommand(
    alloc: std.mem.Allocator,
    io: std.Io,
    args: Args,
    err: *std.Io.Writer,
) !u8 {
    const out_path = args.out_path orelse {
        try err.writeAll("zlsx: scrub-metadata requires --out PATH\n");
        try err.flush();
        return 2;
    };

    var ed = zlsx_pkg.Editor.open(alloc, io, args.file) catch |e| {
        try err.print("zlsx: cannot open '{s}': {s}\n", .{ args.file, @errorName(e) });
        try err.flush();
        return openFailureExit(e);
    };
    defer ed.deinit();

    ed.stripDocProps(.{}) catch |e| {
        try err.print("zlsx: scrub-metadata: {s}\n", .{@errorName(e)});
        try err.flush();
        return 3;
    };

    ed.save(io, out_path) catch |e| {
        try err.print("zlsx: save '{s}': {s}\n", .{ out_path, @errorName(e) });
        try err.flush();
        return 5;
    };
    return 0;
}

/// The coverage's worksheet target, relative to `xl/` — the form
/// `<coverage worksheet_target>` stores.
fn embedWorksheetTarget(ws: *zlsx_pkg.Worksheet) ![]const u8 {
    const part = try ws.resolvePartName();
    return if (std.mem.startsWith(u8, part, "xl/")) part["xl/".len..] else part;
}

/// emb-6c phase one: `zlsx embed <file> --extract --column A --coverage A2:A100`.
///
/// Emits one NDJSON record per row that has something worth embedding.
/// Rows with nothing embeddable are omitted rather than emitted empty —
/// the consumer pays per row, and a vector for "" means nothing.
fn runEmbedExtract(
    alloc: std.mem.Allocator,
    io: std.Io,
    args: Args,
    out: *std.Io.Writer,
    err: *std.Io.Writer,
) !u8 {
    const column = args.column_name orelse {
        try err.writeAll("zlsx: embed --extract requires --column A\n");
        try err.flush();
        return 2;
    };
    const range = args.coverage_range orelse {
        try err.writeAll("zlsx: embed --extract requires --coverage A2:A100\n");
        try err.flush();
        return 2;
    };

    var wb = zlsx_pkg.Workbook.open(alloc, io, args.file) catch |e| {
        try err.print("zlsx: cannot open '{s}': {s}\n", .{ args.file, @errorName(e) });
        try err.flush();
        return openFailureExit(e);
    };
    defer wb.deinit();

    const ws = wb.sheet(@intCast(args.sheet_index orelse 0)) catch |e| {
        try err.print("zlsx: sheet: {s}\n", .{@errorName(e)});
        try err.flush();
        return 3;
    };
    const target = embedWorksheetTarget(ws) catch |e| {
        try err.print("zlsx: sheet: {s}\n", .{@errorName(e)});
        try err.flush();
        return 3;
    };

    const rows = wb.embeddableRows(alloc, target, range, column, false) catch |e| {
        try err.print("zlsx: embed --extract: {s}\n", .{@errorName(e)});
        try err.flush();
        return 3;
    };
    defer alloc.free(rows);

    for (rows) |r| {
        try out.print("{{\"kind\":\"embed_row\",\"row\":{d},\"text\":", .{r.row});
        try writeJsonString(out, r.text);
        try out.writeAll("}\n");
    }
    try out.flush();
    return 0;
}

/// emb-6c phase two: `zlsx embed <file> --vectors vecs.ndjson --model M
/// --column A --coverage A2:A100 --out out.xlsx`.
///
/// Reads `{"row":N,"vector":[…]}` records, writes the embedding parts.
/// Covered rows with no vector become tombstones — the same state
/// `--prune` leaves them in, so a partial embedding is representable
/// rather than an error.
fn runEmbedApply(
    alloc: std.mem.Allocator,
    io: std.Io,
    args: Args,
    vectors_path: []const u8,
    err: *std.Io.Writer,
) !u8 {
    const column = args.column_name orelse {
        try err.writeAll("zlsx: embed --vectors requires --column A\n");
        try err.flush();
        return 2;
    };
    const range = args.coverage_range orelse {
        try err.writeAll("zlsx: embed --vectors requires --coverage A2:A100\n");
        try err.flush();
        return 2;
    };
    const model = args.model_name orelse {
        try err.writeAll("zlsx: embed --vectors requires --model NAME\n");
        try err.flush();
        return 2;
    };
    const out_path = args.out_path orelse {
        try err.writeAll("zlsx: embed --vectors requires --out PATH\n");
        try err.flush();
        return 2;
    };
    const dtype: zlsx_pkg.embedding_part.Dtype = blk: {
        const name = args.dtype_name orelse break :blk .f32;
        if (std.mem.eql(u8, name, "f32")) break :blk .f32;
        if (std.mem.eql(u8, name, "int8-sym")) break :blk .int8_sym_per_vec;
        try err.print("zlsx: unknown --dtype '{s}' (want f32 | int8-sym)\n", .{name});
        try err.flush();
        return 2;
    };

    var ed = zlsx_pkg.Editor.open(alloc, io, args.file) catch |e| {
        try err.print("zlsx: cannot open '{s}': {s}\n", .{ args.file, @errorName(e) });
        try err.flush();
        return openFailureExit(e);
    };
    defer ed.deinit();

    const ws = ed.workbook.sheet(@intCast(args.sheet_index orelse 0)) catch |e| {
        try err.print("zlsx: sheet: {s}\n", .{@errorName(e)});
        try err.flush();
        return 3;
    };
    const target = embedWorksheetTarget(ws) catch |e| {
        try err.print("zlsx: sheet: {s}\n", .{@errorName(e)});
        try err.flush();
        return 3;
    };

    const rows = ed.workbook.embeddableRows(alloc, target, range, column, false) catch |e| {
        try err.print("zlsx: embed --vectors: {s}\n", .{@errorName(e)});
        try err.flush();
        return 3;
    };
    defer alloc.free(rows);

    var vecs = readVectorFile(alloc, io, vectors_path) catch |e| {
        try err.print("zlsx: --vectors '{s}': {s}\n", .{ vectors_path, @errorName(e) });
        try err.flush();
        return 3;
    };
    defer {
        var it = vecs.valueIterator();
        while (it.next()) |v| alloc.free(v.*);
        vecs.deinit(alloc);
    }
    if (vecs.count() == 0) {
        try err.writeAll("zlsx: --vectors carried no records\n");
        try err.flush();
        return 3;
    }

    const parsed_range = zlsx_pkg.embedding_part.parseA1Range(range) catch |e| {
        try err.print("zlsx: --coverage: {s}\n", .{@errorName(e)});
        try err.flush();
        return 2;
    };
    const first_row = parsed_range.first.row;
    const count: u32 = parsed_range.last.row - first_row + 1;

    // Dimension comes from the data, and every vector must agree — a
    // ragged set would produce a part whose header lies about its rows.
    var dim: u32 = 0;
    {
        var it = vecs.valueIterator();
        while (it.next()) |v| {
            const n: u32 = @intCast(v.len);
            if (dim == 0) dim = n;
            if (n != dim) {
                try err.print("zlsx: --vectors: inconsistent dimensions ({d} vs {d})\n", .{ dim, n });
                try err.flush();
                return 3;
            }
        }
    }
    if (dim == 0) {
        try err.writeAll("zlsx: --vectors: zero-length vector\n");
        try err.flush();
        return 3;
    }

    const rec_len = dtype.recordBytes(dim);
    const vec_body = alloc.alloc(u8, @as(usize, count) * rec_len) catch return 4;
    defer alloc.free(vec_body);
    @memset(vec_body, 0);
    const hashes = alloc.alloc(u64, count) catch return 4;
    defer alloc.free(hashes);
    // Every slot starts tombstoned; a row only becomes live if it is
    // both embeddable and carries a vector.
    @memset(hashes, zlsx_pkg.embedding_part.TOMBSTONE_HASH);

    var written: usize = 0;
    for (rows) |r| {
        const vec = vecs.get(r.row) orelse continue;
        const slot = r.row - first_row;
        const dst = vec_body[@as(usize, slot) * rec_len ..][0..rec_len];
        switch (dtype) {
            .f32 => {
                for (vec, 0..) |f, i| {
                    std.mem.writeInt(u32, dst[i * 4 ..][0..4], @bitCast(f), .little);
                }
            },
            .int8_sym_per_vec => {
                const codes = alloc.alloc(i8, dim) catch return 4;
                defer alloc.free(codes);
                const res = zlsx_pkg.embedding_part.quantizeF32ToI8Sym(vec, codes);
                std.mem.writeInt(u32, dst[0..4], @bitCast(res.scale), .little);
                for (codes, 0..) |c, i| dst[4 + i] = @bitCast(c);
            },
            else => unreachable,
        }
        hashes[slot] = r.hash;
        written += 1;
    }

    ed.workbook.setEmbeddings(model, dim, dtype, &[_]zlsx_pkg.EmbeddingCoverageInput{.{
        .id = args.coverage_id orelse "default",
        .worksheet_target = target,
        .range = range,
        .column = column,
        .include_formulas = false,
        .vec_body = vec_body,
        .hashes = hashes,
    }}) catch |e| {
        try err.print("zlsx: setEmbeddings: {s}\n", .{@errorName(e)});
        try err.flush();
        return 3;
    };

    ed.save(io, out_path) catch |e| {
        try err.print("zlsx: save '{s}': {s}\n", .{ out_path, @errorName(e) });
        try err.flush();
        return 5;
    };
    return 0;
}

/// Parse `{"row":N,"vector":[…]}` NDJSON into row → owned f32 slice.
fn readVectorFile(
    alloc: std.mem.Allocator,
    io: std.Io,
    path: []const u8,
) !std.AutoHashMapUnmanaged(u32, []f32) {
    const bytes = try std.Io.Dir.cwd().readFileAlloc(io, path, alloc, .limited(512 * 1024 * 1024));
    defer alloc.free(bytes);

    var map: std.AutoHashMapUnmanaged(u32, []f32) = .empty;
    errdefer {
        var it = map.valueIterator();
        while (it.next()) |v| alloc.free(v.*);
        map.deinit(alloc);
    }

    var lines = std.mem.splitScalar(u8, bytes, '\n');
    while (lines.next()) |raw| {
        const line = std.mem.trim(u8, raw, " \t\r");
        if (line.len == 0) continue;

        var parsed = try std.json.parseFromSlice(std.json.Value, alloc, line, .{});
        defer parsed.deinit();
        const obj = switch (parsed.value) {
            .object => |o| o,
            else => return error.InvalidVectorRecord,
        };
        const row_v = obj.get("row") orelse return error.InvalidVectorRecord;
        const row: u32 = switch (row_v) {
            .integer => |n| if (n > 0 and n <= std.math.maxInt(u32))
                @intCast(n)
            else
                return error.InvalidVectorRecord,
            else => return error.InvalidVectorRecord,
        };
        const vec_v = obj.get("vector") orelse return error.InvalidVectorRecord;
        const arr = switch (vec_v) {
            .array => |a| a,
            else => return error.InvalidVectorRecord,
        };
        const vals = try alloc.alloc(f32, arr.items.len);
        errdefer alloc.free(vals);
        for (arr.items, 0..) |item, i| {
            vals[i] = switch (item) {
                .float => |f| @floatCast(f),
                .integer => |n| @floatFromInt(n),
                else => return error.InvalidVectorRecord,
            };
        }
        // Last record for a row wins, so a regenerated tail can simply
        // be appended rather than requiring the file be rewritten.
        if (map.fetchPut(alloc, row, vals) catch return error.OutOfMemory) |old| {
            alloc.free(old.value);
        }
    }
    return map;
}

/// emb-6a: `zlsx embed <file> --strip --out PATH`.
///
/// Removes the embedding parts and the recovery record together, so
/// the saved workbook reports `absent` rather than `stripped`. That is
/// the point of the operation: this is the pre-share path, and leaving
/// recoverable provenance behind would defeat it.
fn runEmbedCommand(
    alloc: std.mem.Allocator,
    io: std.Io,
    args: Args,
    out: *std.Io.Writer,
    err: *std.Io.Writer,
) !u8 {
    // Exactly one mode. These do contradictory things to the same
    // parts, so combining them is a confusion to surface rather than an
    // order to resolve.
    var mode_count: u8 = 0;
    if (args.strip) mode_count += 1;
    if (args.prune) mode_count += 1;
    if (args.extract) mode_count += 1;
    if (args.vectors_path != null) mode_count += 1;
    if (mode_count == 0) {
        try err.writeAll("zlsx: embed requires one of --extract, --vectors PATH, --prune, --strip\n");
        try err.flush();
        return 2;
    }
    if (mode_count > 1) {
        try err.writeAll("zlsx: embed modes are mutually exclusive (--extract / --vectors / --prune / --strip)\n");
        try err.flush();
        return 2;
    }

    // Phase one of the write path: read-only, emits on stdout, never
    // touches the workbook — so it deliberately does NOT want --out.
    if (args.extract) return try runEmbedExtract(alloc, io, args, out, err);
    if (args.vectors_path) |vp| return try runEmbedApply(alloc, io, args, vp, err);

    const mode: []const u8 = if (args.strip) "--strip" else "--prune";
    const out_path = args.out_path orelse {
        try err.print("zlsx: embed {s} requires --out PATH\n", .{mode});
        try err.flush();
        return 2;
    };

    var ed = zlsx_pkg.Editor.open(alloc, io, args.file) catch |e| {
        try err.print("zlsx: cannot open '{s}': {s}\n", .{ args.file, @errorName(e) });
        try err.flush();
        return openFailureExit(e);
    };
    defer ed.deinit();

    if (args.strip) {
        ed.workbook.stripEmbeddings() catch |e| {
            try err.print("zlsx: embed --strip: {s}\n", .{@errorName(e)});
            try err.flush();
            return 3;
        };
    } else {
        const report = ed.workbook.pruneEmbeddings() catch |e| {
            try err.print("zlsx: embed --prune: {s}\n", .{@errorName(e)});
            try err.flush();
            return 3;
        };
        // Prune's result is the point of running it — how many vectors
        // outlived their text. Reported on stdout in the same NDJSON
        // envelope the reader sub-commands use, so it composes with jq
        // rather than needing to be scraped from a log line.
        try out.print(
            "{{\"kind\":\"prune\",\"redacted\":{d},\"stale\":{d},\"fresh\":{d},\"valid_empty\":{d}}}\n",
            .{ report.redacted, report.stale, report.fresh, report.valid_empty },
        );
        try out.flush();
    }

    ed.save(io, out_path) catch |e| {
        try err.print("zlsx: save '{s}': {s}\n", .{ out_path, @errorName(e) });
        try err.flush();
        return 5;
    };
    return 0;
}

/// iter57: lighter NDJSON variant of `meta` — one record per sheet,
/// name + index only. Same envelope shape as `meta`'s sheet record
/// minus the workbook-scoped `has_comments` field, so consumers can
/// trivially swap between the two commands.
fn runListSheetsCommand(out: *std.Io.Writer, book: *const xlsx.Book, output: OutputMode) !void {
    if (output == .pretty_json) {
        try out.writeAll("{\n  \"sheets\": [");
        if (book.sheets.len == 0) {
            try out.writeAll("]\n}\n");
            try out.flush();
            return;
        }
        try out.writeByte('\n');
        for (book.sheets, 0..) |s, i| {
            if (signals.shouldStop()) return;
            try out.writeAll("    {\"kind\": \"sheet\", \"sheet\": ");
            try writeJsonString(out, s.name);
            try out.print(
                ", \"sheet_idx\": {d}, \"state\": \"{s}\"}}",
                .{ i, s.state.toString() },
            );
            if (i + 1 < book.sheets.len) try out.writeByte(',');
            try out.writeByte('\n');
        }
        try out.writeAll("  ]\n}\n");
        try out.flush();
        return;
    }

    for (book.sheets, 0..) |s, i| {
        if (signals.shouldStop()) return;
        try out.writeAll("{\"kind\":\"sheet\",\"sheet\":");
        try writeJsonString(out, s.name);
        // `state` surfaces `<sheet state="…">`. veryHidden sheets are
        // unreachable from Excel's UI, so a caller scanning a workbook
        // has no other way to learn they exist.
        try out.print(",\"sheet_idx\":{d},\"state\":\"{s}\"}}\n", .{ i, s.state.toString() });
    }
    try out.flush();
}

// ─── iter58: reader-surface sub-commands ─────────────────────────────

/// Resolve the sheet-selector flags (--sheet index / --name) to an
/// optional sheet filter. Null means "iterate every sheet" (the
/// default for sheet-scoped-but-multi-sheet commands like comments /
/// validations / hyperlinks). Returns error.SheetNotFound when a
/// concrete selector was given but doesn't match the workbook.
///
/// iter59c: --all-sheets / --sheet-glob also collapse to null here so
/// the existing comments/validations/hyperlinks loop iterates every
/// sheet; per-sheet inclusion is then decided by `isSheetIncluded`.
fn resolveSheetFilter(book: *const xlsx.Book, args: Args) !?usize {
    // iter59c: --all-sheets / --sheet-glob both mean "visit every
    // sheet and let isSheetIncluded decide". Return null so the caller
    // takes the multi-sheet branch.
    if (args.all_sheets or args.sheet_glob != null) return null;
    if (args.sheet_index) |idx| {
        if (idx >= book.sheets.len) return error.SheetNotFound;
        return idx;
    }
    if (args.sheet_name) |name| {
        for (book.sheets, 0..) |s, i| {
            if (std.mem.eql(u8, s.name, name)) return i;
        }
        return error.SheetNotFound;
    }
    return null;
}

/// iter59c: simple-glob matcher. `*` matches any run (including empty),
/// `?` matches exactly one char. Case-sensitive, no escapes. Recursive
/// — pattern depth is bounded by `*` count, and user-supplied patterns
/// are tiny in practice (sheet-name glob, not path glob).
fn globMatch(pattern: []const u8, text: []const u8) bool {
    if (pattern.len == 0) return text.len == 0;
    if (pattern[0] == '*') {
        // Skip consecutive stars to keep the recursion shallow.
        var p_rest = pattern[1..];
        while (p_rest.len > 0 and p_rest[0] == '*') p_rest = p_rest[1..];
        if (p_rest.len == 0) return true; // trailing star matches the rest
        var i: usize = 0;
        while (i <= text.len) : (i += 1) {
            if (globMatch(p_rest, text[i..])) return true;
        }
        return false;
    }
    if (text.len == 0) return false;
    if (pattern[0] == '?') {
        // `?` matches exactly one UTF-8 codepoint, not one byte. For
        // non-ASCII sheet names (e.g. `表1`, `Résumé`) a byte-stride
        // advance would land inside a multi-byte sequence and poison
        // every subsequent literal compare.
        const n = std.unicode.utf8ByteSequenceLength(text[0]) catch 1;
        if (text.len < n) return false;
        return globMatch(pattern[1..], text[n..]);
    }
    if (pattern[0] == text[0]) {
        return globMatch(pattern[1..], text[1..]);
    }
    return false;
}

/// iter59c: per-sheet inclusion test used inside the multi-sheet loops
/// of comments / validations / hyperlinks / rows / cells. When the
/// caller already narrowed to a concrete index via `resolveSheetFilter`
/// → Some(idx), it only defers to that. Otherwise:
/// - --sheet-glob matches sheet name against the pattern;
/// - --all-sheets accepts every sheet;
/// - default (none set) accepts only sheet 0.
fn isSheetIncluded(args: Args, sheet_name: []const u8, sheet_idx: usize) bool {
    if (args.sheet_glob) |pat| return globMatch(pat, sheet_name);
    if (args.all_sheets) return true;
    // iter55a default: first sheet only.
    return sheet_idx == 0;
}

/// Emit an `"A1"`-style ref into `buf` from a reader-shape CellRef
/// (`col` is 0-based — A→0, B→1 — and `row` is 1-based, matching
/// `xlsx.parseA1Ref`). Panics if the generated ref exceeds 16 bytes —
/// OOXML's max column XFD (=16 383) plus max row 1 048 576 fits in
/// 10 bytes, so the budget has a lot of slack. Callers must not hold
/// the returned slice past the buffer's lifetime.
fn refFromCellRef(buf: *[16]u8, ref: xlsx.CellRef) []const u8 {
    std.debug.assert(ref.row >= 1);
    var letters_buf: [8]u8 = undefined;
    const letters = colLetter(&letters_buf, ref.col);
    return std.fmt.bufPrint(buf, "{s}{d}", .{ letters, ref.row }) catch unreachable;
}

/// Emit `{"text":…,"bold":…,…}` fields for a single RichRun. Caller
/// wraps in the surrounding `[` / `]`. `bold`, `italic`, `color`,
/// `size`, `font_name` are each emitted only when set (matches the
/// design-doc "emitted only when true/non-null" shorthand).
fn writeRichRun(w: *std.Io.Writer, run: xlsx.RichRun) !void {
    try w.writeAll("{\"text\":");
    try writeJsonString(w, run.text);
    if (run.bold) try w.writeAll(",\"bold\":true");
    if (run.italic) try w.writeAll(",\"italic\":true");
    if (run.color_argb) |c| try w.print(",\"color\":\"{X:0>8}\"", .{c});
    if (run.size) |s| {
        if (std.math.isFinite(s)) try w.print(",\"size\":{d}", .{s});
    }
    if (run.font_name.len != 0) {
        try w.writeAll(",\"font_name\":");
        try writeJsonString(w, run.font_name);
    }
    try w.writeByte('}');
}

/// Emit `null` for plain strings, otherwise `[<run>,…]`. Shared by
/// `comments` and `sst` which use the same runs wire-shape.
fn writeRichRunsOrNull(w: *std.Io.Writer, runs: ?[]const xlsx.RichRun) !void {
    const rs = runs orelse {
        try w.writeAll("null");
        return;
    };
    try w.writeByte('[');
    for (rs, 0..) |r, i| {
        if (i != 0) try w.writeByte(',');
        try writeRichRun(w, r);
    }
    try w.writeByte(']');
}

/// Map the reader's DataValidationKind to the OOXML wire-form string
/// the design doc pins in the `rule_type` field. `.unknown` surfaces
/// as the literal `"unknown"` so consumers can still filter.
fn validationKindName(kind: xlsx.DataValidationKind) []const u8 {
    return switch (kind) {
        .list => "list",
        .whole => "whole",
        .decimal => "decimal",
        .date => "date",
        .time => "time",
        .text_length => "textLength",
        .custom => "custom",
        .unknown => "unknown",
    };
}

/// Map DataValidationOperator to its OOXML camelCase token.
fn validationOpName(op: xlsx.DataValidationOperator) []const u8 {
    return switch (op) {
        .between => "between",
        .not_between => "notBetween",
        .equal => "equal",
        .not_equal => "notEqual",
        .less_than => "lessThan",
        .less_than_or_equal => "lessThanOrEqual",
        .greater_than => "greaterThan",
        .greater_than_or_equal => "greaterThanOrEqual",
    };
}

/// Emit `"A1"` for a single-cell range or `"A1:B2"` for a rectangle
/// into the caller-provided 32-byte buffer. Uses `refFromColRow`
/// under the hood so both endpoints get identical formatting.
fn rangeFromBounds(buf: *[32]u8, top_left: xlsx.CellRef, bottom_right: xlsx.CellRef) []const u8 {
    var tl_buf: [16]u8 = undefined;
    const tl = refFromCellRef(&tl_buf, top_left);
    if (top_left.col == bottom_right.col and top_left.row == bottom_right.row) {
        return std.fmt.bufPrint(buf, "{s}", .{tl}) catch unreachable;
    }
    var br_buf: [16]u8 = undefined;
    const br = refFromCellRef(&br_buf, bottom_right);
    return std.fmt.bufPrint(buf, "{s}:{s}", .{ tl, br }) catch unreachable;
}

/// Emit one NDJSON record per comment. Sheet selection follows the
/// unified iter59c rules:
///  - `filter = Some(idx)` → only that sheet (--sheet / --name);
///  - `filter = null` → fall back to `isSheetIncluded(args, …)` —
///    `--all-sheets` / `--sheet-glob` gate, else every sheet (legacy
///    default for this sub-command, preserved for back-compat).
///
/// Pagination persists across sheets so `--skip` / `--take` slice the
/// concatenated cross-sheet stream (per docs/jq-for-excel.md v4.1).
fn runCommentsCommand(
    out: *std.Io.Writer,
    book: *const xlsx.Book,
    filter: ?usize,
    args: Args,
    skip: ?usize,
    take: ?usize,
    start_row: ?u32,
    end_row: ?u32,
) !void {
    var pg = Pagination.init(skip, take);
    const compact = args.output == .compact_ndjson;
    // iter60b: in compact mode, prologue is emitted on first
    // to-be-emitted record of each sheet. Tracking last emitted
    // sheet_idx rather than a per-sheet local bool lets pagination
    // that straddles sheet boundaries still get prologues interleaved
    // correctly (a --take that stops mid-sheet 1 emits sheet 0's
    // prologue but never sheet 1's).
    var last_prologue: ?usize = null;
    for (book.sheets, 0..) |s, sheet_idx| {
        if (signals.shouldStop()) return;
        if (filter) |f| {
            if (sheet_idx != f) continue;
        } else if (args.sheet_glob != null or args.all_sheets) {
            // iter59c: honour the glob/--all-sheets pair. When neither
            // is set, the legacy "iterate every sheet" default is kept
            // (this sub-command has no natural `sheet 0 only` anchor).
            if (!isSheetIncluded(args, s.name, sheet_idx)) continue;
        }
        for (book.comments(s)) |c| {
            // Comments are not guaranteed monotonic by row across a
            // sheet's comment list (OOXML preserves author/insertion
            // order). `continue` on both bounds — don't `break`.
            if (start_row) |sr| if (c.top_left.row < sr) continue;
            if (end_row) |er| if (c.top_left.row > er) continue;
            if (signals.shouldStop()) return;
            switch (pg.consume()) {
                .drop => continue,
                .stop => return,
                .emit => {},
            }
            if (compact and (last_prologue == null or last_prologue.? != sheet_idx)) {
                try writeCompactSheetPrologue(out, s.name, sheet_idx);
                last_prologue = sheet_idx;
            }
            var ref_buf: [16]u8 = undefined;
            const ref = refFromCellRef(&ref_buf, c.top_left);

            if (compact) {
                try out.writeAll("{\"kind\":\"comment\",\"ref\":");
                try writeJsonString(out, ref);
            } else {
                try out.writeAll("{\"kind\":\"comment\",\"sheet\":");
                try writeJsonString(out, s.name);
                try out.print(",\"sheet_idx\":{d},\"ref\":", .{sheet_idx});
                try writeJsonString(out, ref);
            }
            // Reader `col` is 0-based (A=0); wire format is 1-based
            // (A=1) for consistency with `cells` / `rows` envelopes.
            try out.print(
                ",\"row\":{d},\"col\":{d},\"author\":",
                .{ c.top_left.row, c.top_left.col + 1 },
            );
            try writeJsonString(out, c.author);
            try out.writeAll(",\"text\":");
            try writeJsonString(out, c.text);
            try out.writeAll(",\"runs\":");
            try writeRichRunsOrNull(out, c.runs);
            try out.writeAll("}\n");
        }
    }
    try out.flush();
}

/// Emit one NDJSON record per data-validation range. Sheet selection
/// follows the same iter59c rules as runCommentsCommand — see there.
fn runValidationsCommand(
    out: *std.Io.Writer,
    book: *const xlsx.Book,
    filter: ?usize,
    args: Args,
    skip: ?usize,
    take: ?usize,
) !void {
    var pg = Pagination.init(skip, take);
    const compact = args.output == .compact_ndjson;
    var last_prologue: ?usize = null;
    for (book.sheets, 0..) |s, sheet_idx| {
        if (signals.shouldStop()) return;
        if (filter) |f| {
            if (sheet_idx != f) continue;
        } else if (args.sheet_glob != null or args.all_sheets) {
            if (!isSheetIncluded(args, s.name, sheet_idx)) continue;
        }
        for (book.dataValidations(s)) |dv| {
            if (signals.shouldStop()) return;
            switch (pg.consume()) {
                .drop => continue,
                .stop => return,
                .emit => {},
            }
            if (compact and (last_prologue == null or last_prologue.? != sheet_idx)) {
                try writeCompactSheetPrologue(out, s.name, sheet_idx);
                last_prologue = sheet_idx;
            }
            var range_buf: [32]u8 = undefined;
            const range = rangeFromBounds(&range_buf, dv.top_left, dv.bottom_right);

            if (compact) {
                try out.writeAll("{\"kind\":\"validation\",\"range\":");
                try writeJsonString(out, range);
            } else {
                try out.writeAll("{\"kind\":\"validation\",\"sheet\":");
                try writeJsonString(out, s.name);
                try out.print(",\"sheet_idx\":{d},\"range\":", .{sheet_idx});
                try writeJsonString(out, range);
            }
            try out.print(",\"rule_type\":\"{s}\",\"op\":", .{validationKindName(dv.kind)});
            if (dv.op) |op| try out.print("\"{s}\"", .{validationOpName(op)}) else try out.writeAll("null");

            try out.writeAll(",\"formula1\":");
            try writeJsonString(out, dv.formula1);
            try out.writeAll(",\"formula2\":");
            if (dv.formula2.len != 0) try writeJsonString(out, dv.formula2) else try out.writeAll("null");

            try out.writeAll(",\"values\":");
            if (dv.kind == .list and dv.values.len != 0) {
                try out.writeByte('[');
                for (dv.values, 0..) |v, i| {
                    if (i != 0) try out.writeByte(',');
                    try writeJsonString(out, v);
                }
                try out.writeByte(']');
            } else {
                try out.writeAll("null");
            }
            try out.writeAll("}\n");
        }
    }
    try out.flush();
}

/// Emit one NDJSON record per hyperlink. Sheet selection follows the
/// same iter59c rules as runCommentsCommand — see there.
fn runHyperlinksCommand(
    out: *std.Io.Writer,
    book: *const xlsx.Book,
    filter: ?usize,
    args: Args,
    skip: ?usize,
    take: ?usize,
) !void {
    var pg = Pagination.init(skip, take);
    const compact = args.output == .compact_ndjson;
    var last_prologue: ?usize = null;
    for (book.sheets, 0..) |s, sheet_idx| {
        if (signals.shouldStop()) return;
        if (filter) |f| {
            if (sheet_idx != f) continue;
        } else if (args.sheet_glob != null or args.all_sheets) {
            if (!isSheetIncluded(args, s.name, sheet_idx)) continue;
        }
        for (book.hyperlinks(s)) |h| {
            if (signals.shouldStop()) return;
            switch (pg.consume()) {
                .drop => continue,
                .stop => return,
                .emit => {},
            }
            if (compact and (last_prologue == null or last_prologue.? != sheet_idx)) {
                try writeCompactSheetPrologue(out, s.name, sheet_idx);
                last_prologue = sheet_idx;
            }
            var range_buf: [32]u8 = undefined;
            const range = rangeFromBounds(&range_buf, h.top_left, h.bottom_right);

            if (compact) {
                try out.writeAll("{\"kind\":\"hyperlink\",\"range\":");
                try writeJsonString(out, range);
            } else {
                try out.writeAll("{\"kind\":\"hyperlink\",\"sheet\":");
                try writeJsonString(out, s.name);
                try out.print(",\"sheet_idx\":{d},\"range\":", .{sheet_idx});
                try writeJsonString(out, range);
            }
            try out.writeAll(",\"url\":");
            if (h.url.len != 0) try writeJsonString(out, h.url) else try out.writeAll("null");
            try out.writeAll(",\"location\":");
            if (h.location.len != 0) try writeJsonString(out, h.location) else try out.writeAll("null");
            try out.writeAll("}\n");
        }
    }
    try out.flush();
}

/// S6: `zlsx pivots` — one `{"kind":"pivot",…}` record per pivot table in
/// host-sheet order, then one `{"kind":"pivot_cache",…}` record per cache
/// no pivot table reads. Routes through the package layer: `Workbook` is
/// the surface that walks relationships, and a pivot is nothing but
/// relationships (`pkg/pivots.zig`) — the reader-only `Book` has no view
/// of the parts. Sheet selection follows the read family (`--sheet` /
/// `--name` narrow to one host sheet, `--all-sheets` / `--sheet-glob`
/// widen, the default visits every sheet); orphan caches are workbook-
/// scoped and only ride along when no sheet was selected. Contract in
/// docs/cli.md, "pivots".
fn runPivotsCommand(
    alloc: std.mem.Allocator,
    io: std.Io,
    args: Args,
    out: *std.Io.Writer,
    err: *std.Io.Writer,
) !u8 {
    var wb = zlsx_pkg.Workbook.open(alloc, io, args.file) catch |e| {
        try err.print("zlsx: cannot open '{s}': {s}\n", .{ args.file, @errorName(e) });
        try err.flush();
        return openFailureExit(e);
    };
    defer wb.deinit();
    var pivots = wb.pivotTables() catch |e| {
        try err.print("zlsx: cannot read pivots in '{s}': {s}\n", .{ args.file, @errorName(e) });
        try err.flush();
        return 2;
    };
    defer pivots.deinit();

    const filter: ?usize = blk: {
        if (args.all_sheets or args.sheet_glob != null) break :blk null;
        if (args.sheet_index) |idx| {
            if (idx >= pivots.sheet_names.len) {
                try err.writeAll("zlsx: sheet not found\n");
                try err.flush();
                return 3;
            }
            break :blk idx;
        }
        if (args.sheet_name) |name| {
            for (pivots.sheet_names, 0..) |s, i| {
                if (std.mem.eql(u8, s, name)) break :blk i;
            }
            try err.writeAll("zlsx: sheet not found\n");
            try err.flush();
            return 3;
        }
        break :blk null;
    };

    var pg = Pagination.init(args.skip, args.take);
    const compact = args.output == .compact_ndjson;
    var last_prologue: ?usize = null;
    for (pivots.tables) |pt| {
        if (signals.shouldStop()) return 0;
        if (filter) |f| {
            if (pt.sheet_idx != f) continue;
        } else if (args.sheet_glob != null or args.all_sheets) {
            if (!isSheetIncluded(args, pt.sheet_name, pt.sheet_idx)) continue;
        }
        switch (pg.consume()) {
            .drop => continue,
            .stop => {
                try out.flush();
                return 0;
            },
            .emit => {},
        }
        if (compact and (last_prologue == null or last_prologue.? != pt.sheet_idx)) {
            try writeCompactSheetPrologue(out, pt.sheet_name, pt.sheet_idx);
            last_prologue = pt.sheet_idx;
        }
        try zlsx_pkg.pivots.ndjson.writeTable(out, &pivots, pt, if (compact) .compact else .full);
    }

    if (filter == null and args.sheet_glob == null) {
        for (pivots.caches) |c| {
            if (c.consumer_count != 0) continue;
            if (signals.shouldStop()) return 0;
            switch (pg.consume()) {
                .drop => continue,
                .stop => {
                    try out.flush();
                    return 0;
                },
                .emit => {},
            }
            try zlsx_pkg.pivots.ndjson.writeCacheRecord(out, &pivots, c);
        }
    }
    try out.flush();
    return 0;
}

/// S3b: emit one NDJSON record per merged range. Sheet selection
/// follows the same iter59c rules as runCommentsCommand — see there.
fn runMergesCommand(
    out: *std.Io.Writer,
    err: *std.Io.Writer,
    book: *const xlsx.Book,
    filter: ?usize,
    args: Args,
    skip: ?usize,
    take: ?usize,
) !u8 {
    // A merge record's only user-text channel is the sheet name (the
    // range and corners are ASCII by construction), so one that is
    // not UTF-8 would make the whole stream unparseable NDJSON under
    // exit 0 (Codex #211 r4). Refuse up front — before any record —
    // when a sheet the selection includes would emit records under a
    // name the stream cannot carry; a bad-named sheet with no merges
    // emits nothing and does not lie.
    for (book.sheets, 0..) |s, sheet_idx| {
        if (filter) |f| {
            if (sheet_idx != f) continue;
        } else if (args.sheet_glob != null or args.all_sheets) {
            if (!isSheetIncluded(args, s.name, sheet_idx)) continue;
        }
        if (book.mergedRanges(s).len == 0) continue;
        if (!std.unicode.utf8ValidateSlice(s.name)) {
            try err.print(
                "zlsx: cannot read merges in '{s}': sheet {d} has a non-UTF-8 name\n",
                .{ args.file, sheet_idx },
            );
            try err.flush();
            return 2;
        }
    }

    var pg = Pagination.init(skip, take);
    const compact = args.output == .compact_ndjson;
    var last_prologue: ?usize = null;
    for (book.sheets, 0..) |s, sheet_idx| {
        if (signals.shouldStop()) return 0;
        if (filter) |f| {
            if (sheet_idx != f) continue;
        } else if (args.sheet_glob != null or args.all_sheets) {
            if (!isSheetIncluded(args, s.name, sheet_idx)) continue;
        }
        for (book.mergedRanges(s)) |m| {
            if (signals.shouldStop()) return 0;
            switch (pg.consume()) {
                .drop => continue,
                // Flush before the early return: runMain's deferred
                // flush swallows errors, so a --take'd stream that
                // relied on it would report exit 0 on a failed final
                // write instead of the documented exit 5 (Codex #211
                // r1).
                .stop => {
                    try out.flush();
                    return 0;
                },
                .emit => {},
            }
            if (compact and (last_prologue == null or last_prologue.? != sheet_idx)) {
                try writeCompactSheetPrologue(out, s.name, sheet_idx);
                last_prologue = sheet_idx;
            }
            var range_buf: [32]u8 = undefined;
            const range = rangeFromBounds(&range_buf, m.top_left, m.bottom_right);

            if (compact) {
                try out.writeAll("{\"kind\":\"merge\",\"range\":");
                try writeJsonString(out, range);
            } else {
                try out.writeAll("{\"kind\":\"merge\",\"sheet\":");
                try writeJsonString(out, s.name);
                try out.print(",\"sheet_idx\":{d},\"range\":", .{sheet_idx});
                try writeJsonString(out, range);
            }
            // Reader cols are 0-based (A=0); the wire is 1-based like
            // the `cells` / `comments` envelopes.
            try out.print(
                ",\"start_row\":{d},\"start_col\":{d},\"end_row\":{d},\"end_col\":{d}}}\n",
                .{ m.top_left.row, m.top_left.col + 1, m.bottom_right.row, m.bottom_right.col + 1 },
            );
        }
    }
    try out.flush();
    return 0;
}

/// S3b: `zlsx defined-names` — one `{"kind":"defined_name",…}` record
/// per `<definedName>` of `xl/workbook.xml`, in document order. Routes
/// through the package layer: the reader-only Book has no workbook.xml
/// view. A concrete selector (`--sheet` / `--name`) narrows to the
/// names SCOPED to that sheet (`localSheetId`) and suppresses
/// workbook-scope names, the way a selector suppresses orphan caches
/// on `pivots`; `--sheet-glob` matches the scope sheet's name; the
/// default and `--all-sheets` stream every name. Contract in
/// docs/cli.md, "defined-names".
fn runDefinedNamesCommand(
    alloc: std.mem.Allocator,
    io: std.Io,
    args: Args,
    out: *std.Io.Writer,
    err: *std.Io.Writer,
) !u8 {
    var wb = zlsx_pkg.Workbook.open(alloc, io, args.file) catch |e| {
        try err.print("zlsx: cannot open '{s}': {s}\n", .{ args.file, @errorName(e) });
        try err.flush();
        return openFailureExit(e);
    };
    defer wb.deinit();
    var view = zlsx_pkg.defined_names_ndjson.collect(alloc, &wb.workbook) catch |e| {
        try err.print("zlsx: cannot read defined names in '{s}': {s}\n", .{ args.file, @errorName(e) });
        try err.flush();
        return 2;
    };
    defer view.deinit();

    const filter: ?usize = blk: {
        if (args.all_sheets or args.sheet_glob != null) break :blk null;
        if (args.sheet_index) |idx| {
            if (idx >= view.sheet_names.len) {
                try err.writeAll("zlsx: sheet not found\n");
                try err.flush();
                return 3;
            }
            break :blk idx;
        }
        if (args.sheet_name) |name| {
            for (view.sheet_names, 0..) |s, i| {
                if (std.mem.eql(u8, s, name)) break :blk i;
            }
            try err.writeAll("zlsx: sheet not found\n");
            try err.flush();
            return 3;
        }
        break :blk null;
    };

    var pg = Pagination.init(args.skip, args.take);
    for (view.names) |d| {
        if (signals.shouldStop()) return 0;
        if (filter) |f| {
            const sid = d.scope_sheet_idx orelse continue;
            if (sid != f) continue;
        } else if (args.sheet_glob) |pat| {
            // A name whose localSheetId is past the sheet list has no
            // scope sheet to match — skipped, like every non-matching
            // scope. Workbook-scope names are suppressed here too.
            const scope = d.scope_sheet orelse continue;
            if (!globMatch(pat, scope)) continue;
        }
        switch (pg.consume()) {
            .drop => continue,
            .stop => {
                try out.flush();
                return 0;
            },
            .emit => {},
        }
        try zlsx_pkg.defined_names_ndjson.writeName(out, d);
    }
    try out.flush();
    return 0;
}

/// S3b slice 4: `zlsx anchors` — one record per anchored image or
/// chart, images before charts within a sheet, sheets in workbook
/// order. Routes through the package layer like `pivots`: the drawing
/// walkers read parts the reader-only Book has no view of. Sheet
/// selection follows the read family — `--sheet` / `--name` narrow to
/// one host sheet, `--all-sheets` / `--sheet-glob` widen, the default
/// streams every sheet. Contract in docs/cli.md, "anchors".
fn runAnchorsCommand(
    alloc: std.mem.Allocator,
    io: std.Io,
    args: Args,
    out: *std.Io.Writer,
    err: *std.Io.Writer,
) !u8 {
    var wb = zlsx_pkg.Workbook.open(alloc, io, args.file) catch |e| {
        try err.print("zlsx: cannot open '{s}': {s}\n", .{ args.file, @errorName(e) });
        try err.flush();
        return openFailureExit(e);
    };
    defer wb.deinit();
    var view = zlsx_pkg.anchors_ndjson.collect(alloc, &wb.store, &wb.workbook) catch |e| {
        try err.print("zlsx: cannot read anchors in '{s}': {s}\n", .{ args.file, @errorName(e) });
        try err.flush();
        return 2;
    };
    defer view.deinit();

    const filter: ?usize = blk: {
        if (args.all_sheets or args.sheet_glob != null) break :blk null;
        if (args.sheet_index) |idx| {
            if (idx >= view.sheet_names.len) {
                try err.writeAll("zlsx: sheet not found\n");
                try err.flush();
                return 3;
            }
            break :blk idx;
        }
        if (args.sheet_name) |name| {
            for (view.sheet_names, 0..) |s, i| {
                if (std.mem.eql(u8, s, name)) break :blk i;
            }
            try err.writeAll("zlsx: sheet not found\n");
            try err.flush();
            return 3;
        }
        break :blk null;
    };

    var pg = Pagination.init(args.skip, args.take);
    const compact = args.output == .compact_ndjson;
    var last_prologue: ?usize = null;
    for (view.records) |r| {
        if (signals.shouldStop()) return 0;
        const sheet_idx: usize = r.sheetIdx();
        if (filter) |f| {
            if (sheet_idx != f) continue;
        } else if (args.sheet_glob != null or args.all_sheets) {
            if (!isSheetIncluded(args, r.sheetName(), sheet_idx)) continue;
        }
        switch (pg.consume()) {
            .drop => continue,
            // Flush before the early return: runMain's deferred flush
            // swallows errors, so a --take'd stream that relied on it
            // would report exit 0 on a failed final write instead of
            // the documented exit 5 (the Codex #211 r1 lesson).
            .stop => {
                try out.flush();
                return 0;
            },
            .emit => {},
        }
        if (compact and (last_prologue == null or last_prologue.? != sheet_idx)) {
            try writeCompactSheetPrologue(out, r.sheetName(), sheet_idx);
            last_prologue = sheet_idx;
        }
        try zlsx_pkg.anchors_ndjson.writeRecord(out, r, if (compact) .compact else .full);
    }
    try out.flush();
    return 0;
}

/// Emit `{…}` for a BorderSide or `null` when the side has no style.
fn writeBorderSideOrNull(w: *std.Io.Writer, side: xlsx.BorderSide) !void {
    if (side.style.len == 0) {
        try w.writeAll("null");
        return;
    }
    try w.writeAll("{\"style\":");
    try writeJsonString(w, side.style);
    try w.writeAll(",\"color\":");
    if (side.color_argb) |c| try w.print("\"{X:0>8}\"", .{c}) else try w.writeAll("null");
    try w.writeByte('}');
}

/// iter59b-4: terse-shape border side for the per-cell `style.border`
/// block: `{"s":"<style>","c":"<argb>"}` with the color field elided
/// when absent. Returns true iff the side contributed bytes.
fn writeTerseBorderSide(w: *std.Io.Writer, side: xlsx.BorderSide) !bool {
    if (side.style.len == 0) return false;
    try w.writeAll("{\"s\":");
    try writeJsonString(w, side.style);
    if (side.color_argb) |c| {
        try w.print(",\"c\":\"{X:0>8}\"", .{c});
    }
    try w.writeByte('}');
    return true;
}

/// iter59b-4: emit the terse `style:{…}` block for a cell when its
/// resolved style has any effective formatting. Returns true when a
/// block was written (so the caller's leading `,\"style\":` prefix was
/// needed), false when the style is structurally empty and the caller
/// must omit the field entirely.
///
/// Terse shape per docs/jq-for-excel.md v4.1:
///   `{"bold":true,"italic":true,"fg":"FF…","bg":"FF…","nf":"0.00",
///     "border":{"l":{"s":"thin","c":"FF000000"},…}}`
/// Each key is omitted when the underlying value is the default /
/// unset — so an unstyled cell with only a non-General numFmt emits
/// just `{"nf":"m/d/yyyy"}`.
fn writeTerseStyleBlock(
    w: *std.Io.Writer,
    book: *const xlsx.Book,
    style_idx: u32,
) !bool {
    // Resolve once; each getter is a direct random-access lookup.
    const font = book.cellFont(style_idx);
    const fill = book.cellFill(style_idx);
    const border = book.cellBorder(style_idx);
    const nf = book.numberFormat(style_idx);

    // Replicate the styles sub-command's null-detection so an all-
    // default fill/border doesn't register as "has style" just because
    // cellFill returned a zero-valued struct rather than null.
    const fill_effective: bool = if (fill) |fl|
        !((std.mem.eql(u8, fl.pattern, "none") or fl.pattern.len == 0) and
            fl.fg_color_argb == null and fl.bg_color_argb == null)
    else
        false;
    const side_empty = struct {
        fn f(s: xlsx.BorderSide) bool {
            return s.style.len == 0 and s.color_argb == null;
        }
    }.f;
    // Terse block emits only l/r/t/b — diagonal is intentionally
    // omitted per the design doc. Exclude diagonal from the
    // effectiveness check so a diagonal-only cell doesn't produce
    // an empty `"border":{}` block.
    const border_effective: bool = if (border) |b|
        !(side_empty(b.left) and side_empty(b.right) and
            side_empty(b.top) and side_empty(b.bottom))
    else
        false;

    const font_effective: bool = if (font) |f|
        (f.bold or f.italic or f.color_argb != null)
    else
        false;

    // "General" is numFmtId 0 — zlsx.numberFormat resolves built-ins so
    // we compare the string. A null return means no styles.xml at all
    // or no format attached; either way, nothing to emit.
    const nf_effective: bool = if (nf) |s|
        !(std.mem.eql(u8, s, "General"))
    else
        false;

    if (!(font_effective or fill_effective or border_effective or nf_effective)) {
        return false;
    }

    try w.writeByte('{');
    var first = true;

    if (font) |f| {
        if (f.bold) {
            if (!first) try w.writeByte(',');
            first = false;
            try w.writeAll("\"bold\":true");
        }
        if (f.italic) {
            if (!first) try w.writeByte(',');
            first = false;
            try w.writeAll("\"italic\":true");
        }
    }

    // Terse-shape colour contract (matches the design-doc example
    // `{"bold":true,"fg":"FFFFFFFF","bg":"FF1F4E79"}`):
    //   fg → font (text) colour when set
    //   bg → cell background from `<fill>` (prefer `fgColor` on a
    //         solid pattern, else `bgColor`). Fill's fgColor is a
    //         misnomer inherited from OOXML — it's the pattern's
    //         foreground, which for `solid` IS the visible background.
    if (font) |f| if (f.color_argb) |c| {
        if (!first) try w.writeByte(',');
        first = false;
        try w.print("\"fg\":\"{X:0>8}\"", .{c});
    };
    if (fill_effective) {
        // Prefer fgColor (solid fills stash the visible colour there);
        // fall back to bgColor for other pattern types.
        const bg_argb: ?u32 = fill.?.fg_color_argb orelse fill.?.bg_color_argb;
        if (bg_argb) |c| {
            if (!first) try w.writeByte(',');
            first = false;
            try w.print("\"bg\":\"{X:0>8}\"", .{c});
        }
    }

    if (nf_effective) {
        if (!first) try w.writeByte(',');
        first = false;
        try w.writeAll("\"nf\":");
        try writeJsonString(w, nf.?);
    }

    if (border_effective) {
        if (!first) try w.writeByte(',');
        first = false;
        try w.writeAll("\"border\":{");
        const b = border.?;
        var border_first = true;
        // Emit only set sides (l/r/t/b). Diagonal is intentionally
        // excluded from the terse shape — the design doc lists l/r/t/b.
        const sides = [_]struct { key: []const u8, side: xlsx.BorderSide }{
            .{ .key = "l", .side = b.left },
            .{ .key = "r", .side = b.right },
            .{ .key = "t", .side = b.top },
            .{ .key = "b", .side = b.bottom },
        };
        for (sides) |sd| {
            if (sd.side.style.len == 0) continue;
            if (!border_first) try w.writeByte(',');
            border_first = false;
            try w.writeByte('"');
            try w.writeAll(sd.key);
            try w.writeAll("\":");
            _ = try writeTerseBorderSide(w, sd.side);
        }
        try w.writeByte('}');
    }

    try w.writeByte('}');
    return true;
}

/// Emit one NDJSON record per cell-XF style entry. Workbook-scoped.
/// Every nested block (`font` / `fill` / `border`) is either the
/// resolved struct or JSON `null` when the getter returns null.
fn runStylesCommand(
    out: *std.Io.Writer,
    book: *const xlsx.Book,
    skip: ?usize,
    take: ?usize,
) !void {
    var pg = Pagination.init(skip, take);
    for (book.cell_xf_numfmt_ids, 0..) |_, i| {
        if (signals.shouldStop()) return;
        switch (pg.consume()) {
            .drop => continue,
            .stop => return,
            .emit => {},
        }
        const idx: u32 = @intCast(i);

        try out.print("{{\"kind\":\"style\",\"idx\":{d},\"font\":", .{idx});
        if (book.cellFont(idx)) |f| {
            try out.writeAll("{\"bold\":");
            try out.writeAll(if (f.bold) "true" else "false");
            try out.writeAll(",\"italic\":");
            try out.writeAll(if (f.italic) "true" else "false");
            try out.writeAll(",\"color\":");
            if (f.color_argb) |c| try out.print("\"{X:0>8}\"", .{c}) else try out.writeAll("null");
            try out.writeAll(",\"size\":");
            if (f.size) |s| {
                if (std.math.isFinite(s)) try out.print("{d}", .{s}) else try out.writeAll("null");
            } else try out.writeAll("null");
            try out.writeAll(",\"name\":");
            if (f.name.len != 0) try writeJsonString(out, f.name) else try out.writeAll("null");
            try out.writeByte('}');
        } else try out.writeAll("null");

        try out.writeAll(",\"fill\":");
        if (book.cellFill(idx)) |fl| {
            // Treat the default zlsx Fill (pattern="none", both
            // colors null) as "no fill" on the wire, same as when
            // cellFill returned null. Consumers can then trust
            // `fill != null` to mean "the style actually defines
            // a fill."
            const no_fill = (std.mem.eql(u8, fl.pattern, "none") or fl.pattern.len == 0) and
                fl.fg_color_argb == null and fl.bg_color_argb == null;
            if (no_fill) {
                try out.writeAll("null");
            } else {
                try out.writeAll("{\"pattern\":");
                try writeJsonString(out, fl.pattern);
                try out.writeAll(",\"fg\":");
                if (fl.fg_color_argb) |c| try out.print("\"{X:0>8}\"", .{c}) else try out.writeAll("null");
                try out.writeAll(",\"bg\":");
                if (fl.bg_color_argb) |c| try out.print("\"{X:0>8}\"", .{c}) else try out.writeAll("null");
                try out.writeByte('}');
            }
        } else try out.writeAll("null");

        try out.writeAll(",\"border\":");
        if (book.cellBorder(idx)) |b| {
            // Same contract: if every side is the zero BorderSide
            // (empty style + null color), emit `null` so unstyled
            // XFs read as "no border" on the wire.
            const side_empty = struct {
                fn f(s: xlsx.BorderSide) bool {
                    return s.style.len == 0 and s.color_argb == null;
                }
            }.f;
            const no_border = side_empty(b.left) and side_empty(b.right) and
                side_empty(b.top) and side_empty(b.bottom) and side_empty(b.diagonal);
            if (no_border) {
                try out.writeAll("null");
            } else {
                try out.writeAll("{\"left\":");
                try writeBorderSideOrNull(out, b.left);
                try out.writeAll(",\"right\":");
                try writeBorderSideOrNull(out, b.right);
                try out.writeAll(",\"top\":");
                try writeBorderSideOrNull(out, b.top);
                try out.writeAll(",\"bottom\":");
                try writeBorderSideOrNull(out, b.bottom);
                try out.writeAll(",\"diagonal\":");
                try writeBorderSideOrNull(out, b.diagonal);
                try out.writeByte('}');
            }
        } else try out.writeAll("null");

        try out.writeAll(",\"num_fmt\":");
        if (book.numberFormat(idx)) |nf| try writeJsonString(out, nf) else try out.writeAll("null");
        try out.writeAll("}\n");
    }
    try out.flush();
}

/// Emit one NDJSON record per shared-string entry.
fn runSstCommand(
    out: *std.Io.Writer,
    book: *xlsx.Book,
    skip: ?usize,
    take: ?usize,
) !void {
    var pg = Pagination.init(skip, take);
    const sst_count = book.sharedStringsCount();
    var i: usize = 0;
    while (i < sst_count) : (i += 1) {
        if (signals.shouldStop()) return;
        switch (pg.consume()) {
            .drop => continue,
            .stop => return,
            .emit => {},
        }
        const s = try book.sharedStringAt(i);
        try out.print("{{\"kind\":\"sst\",\"idx\":{d},\"text\":", .{i});
        try writeJsonString(out, s);
        try out.writeAll(",\"runs\":");
        try writeRichRunsOrNull(out, book.richRuns(i));
        try out.writeAll("}\n");
    }
    try out.flush();
}

// ─── append-rows (load-modify-save CLI surface, iter-lms-4 follow-up) ──

/// Read NDJSON rows from stdin, append to a sheet of `args.file`,
/// save to `args.out_path`. Each line is a JSON array; each cell is
/// `null` (empty), `true`/`false` (bool), a number (int / float), or
/// a string. Empty lines are skipped. Returns 0 on success, 1 on
/// argument error, 2 on Editor open failure, 3 on per-line parse
/// errors, 5 on save failure.
fn runAppendRowsCommand(
    alloc: std.mem.Allocator,
    io: std.Io,
    args: Args,
    err: *std.Io.Writer,
) !u8 {
    const out_path = args.out_path orelse {
        try err.writeAll("zlsx: append-rows requires --out PATH\n");
        try err.flush();
        return 1;
    };
    const sheet_idx_opt = args.sheet_index;
    if (sheet_idx_opt == null) {
        try err.writeAll("zlsx: append-rows requires --sheet N (0-based)\n");
        try err.flush();
        return 1;
    }
    // The Editor accepts u32; refuse here so a usize that overflows
    // the cast (--sheet 5000000000 on 64-bit) returns the documented
    // exit-1 path instead of trapping mid-cast.
    if (sheet_idx_opt.? > std.math.maxInt(u32)) {
        try err.writeAll("zlsx: --sheet value too large (must fit in u32)\n");
        try err.flush();
        return 1;
    }
    const sheet_idx: u32 = @intCast(sheet_idx_opt.?);

    var ed = zlsx_pkg.Editor.open(alloc, io, args.file) catch |e| {
        try err.print("zlsx: cannot open '{s}': {s}\n", .{ args.file, @errorName(e) });
        try err.flush();
        return openFailureExit(e);
    };
    defer ed.deinit();

    // Slurp stdin once. NDJSON volumes for append are typically
    // bounded (audit logs, ETL deltas); refuse > 256 MiB to keep
    // pathological inputs from OOM-ing the process.
    const stdin = std.Io.File.stdin();
    var stdin_buf: [8192]u8 = undefined;
    var stdin_reader = stdin.reader(proc_io, &stdin_buf);
    const all_input = stdin_reader.interface.allocRemaining(alloc, .limited(256 * 1024 * 1024)) catch |e| {
        try err.print("zlsx: failed to read stdin: {s}\n", .{@errorName(e)});
        try err.flush();
        return 1;
    };
    defer alloc.free(all_input);

    var line_no: usize = 0;
    var i: usize = 0;
    while (i < all_input.len) {
        line_no += 1;
        const nl = std.mem.indexOfScalarPos(u8, all_input, i, '\n') orelse all_input.len;
        var line = all_input[i..nl];
        // Trim trailing CR (Windows line endings).
        if (line.len > 0 and line[line.len - 1] == '\r') line = line[0 .. line.len - 1];
        i = nl + 1;
        if (line.len == 0) continue;

        var parsed = std.json.parseFromSlice(std.json.Value, alloc, line, .{}) catch |e| {
            try err.print("zlsx: line {d}: invalid JSON: {s}\n", .{ line_no, @errorName(e) });
            try err.flush();
            return 3;
        };
        defer parsed.deinit();
        const root = parsed.value;
        if (root != .array) {
            try err.print("zlsx: line {d}: expected JSON array\n", .{line_no});
            try err.flush();
            return 3;
        }

        // Cap per-row cell count at Excel's hard column limit
        // (16384 = XFD). Without this, a malicious / buggy caller
        // submitting `[null, null, ...]` arrays of millions of
        // slots per line can balloon the per-row alloc; the row
        // would later fail at the writer anyway, but rejecting up
        // front avoids the wasted allocation and gives a clearer
        // error.
        if (root.array.items.len > 16_384) {
            try err.print(
                "zlsx: line {d}: row has {d} cells; Excel maximum is 16384\n",
                .{ line_no, root.array.items.len },
            );
            try err.flush();
            return 3;
        }
        const cells = try alloc.alloc(xlsx.Cell, root.array.items.len);
        defer alloc.free(cells);
        for (root.array.items, 0..) |v, ci| {
            cells[ci] = jsonValueToCell(v) catch |e| {
                try err.print(
                    "zlsx: line {d}, cell {d}: {s}\n",
                    .{ line_no, ci + 1, @errorName(e) },
                );
                try err.flush();
                return 3;
            };
        }
        const single_row: [1][]const xlsx.Cell = .{cells};
        ed.appendRows(sheet_idx, &single_row) catch |e| {
            try err.print("zlsx: line {d}: appendRows: {s}\n", .{ line_no, @errorName(e) });
            try err.flush();
            return 3;
        };
    }

    ed.save(io, out_path) catch |e| {
        try err.print("zlsx: save '{s}': {s}\n", .{ out_path, @errorName(e) });
        try err.flush();
        return 5;
    };
    return 0;
}

/// iter-cm-4: rewrite a single cell of an existing workbook in place
/// and save to `--out`. Required flags: `--sheet N`, `--ref A1`,
/// `--value <JSON>`. Value parsing reuses jsonValueToCell so the
/// accepted types match `append-rows` exactly.
fn runSetCellCommand(
    alloc: std.mem.Allocator,
    io: std.Io,
    args: Args,
    err: *std.Io.Writer,
) !u8 {
    const out_path = args.out_path orelse {
        try err.writeAll("zlsx: set-cell requires --out PATH\n");
        try err.flush();
        return 1;
    };
    if (args.sheet_index == null) {
        try err.writeAll("zlsx: set-cell requires --sheet N (0-based)\n");
        try err.flush();
        return 1;
    }
    if (args.sheet_index.? > std.math.maxInt(u32)) {
        try err.writeAll("zlsx: --sheet value too large (must fit in u32)\n");
        try err.flush();
        return 1;
    }
    const sheet_idx: u32 = @intCast(args.sheet_index.?);

    const ref = args.cell_ref orelse {
        try err.writeAll("zlsx: set-cell requires --ref A1\n");
        try err.flush();
        return 1;
    };
    const value_json = args.cell_value_json orelse {
        try err.writeAll("zlsx: set-cell requires --value <JSON>\n");
        try err.flush();
        return 1;
    };

    // Parse "A1" → (row, col0). Reuse parseA1Range over a single ref;
    // a bare "A1" parses as a 1-cell rectangle, which is exactly what
    // we want.
    const range = xlsx.parseA1Range(ref) catch {
        try err.print("zlsx: invalid --ref '{s}' (expected A1-style)\n", .{ref});
        try err.flush();
        return 1;
    };
    if (range.top_left.row != range.bottom_right.row or range.top_left.col != range.bottom_right.col) {
        try err.print("zlsx: --ref must be a single cell (got '{s}')\n", .{ref});
        try err.flush();
        return 1;
    }
    // parseA1Ref returns row 1-based, col 0-based — match setCell's
    // signature directly (sheet, row_1based, col_0based, cell).
    const row_1based: u32 = @intCast(range.top_left.row);
    const col_0based: u32 = @intCast(range.top_left.col);

    var parsed = std.json.parseFromSlice(std.json.Value, alloc, value_json, .{}) catch |e| {
        try err.print("zlsx: --value invalid JSON: {s}\n", .{@errorName(e)});
        try err.flush();
        return 3;
    };
    defer parsed.deinit();
    const cell = jsonValueToCell(parsed.value) catch |e| {
        try err.print("zlsx: --value: {s}\n", .{@errorName(e)});
        try err.flush();
        return 3;
    };

    var ed = zlsx_pkg.Editor.open(alloc, io, args.file) catch |e| {
        try err.print("zlsx: cannot open '{s}': {s}\n", .{ args.file, @errorName(e) });
        try err.flush();
        return openFailureExit(e);
    };
    defer ed.deinit();

    ed.setCell(sheet_idx, row_1based, col_0based, cell) catch |e| {
        try err.print("zlsx: setCell {s}: {s}\n", .{ ref, @errorName(e) });
        try err.flush();
        return 3;
    };
    ed.save(io, out_path) catch |e| {
        try err.print("zlsx: save '{s}': {s}\n", .{ out_path, @errorName(e) });
        try err.flush();
        return 5;
    };
    return 0;
}

/// Decode an A1 column letter run ("A"…"XFD") into the 1-based
/// column index Editor.{insert,delete}Column expect. Mirrors the
/// reader's `parseA1Ref` letter loop. Returns null on malformed
/// input.
fn parseColLettersToOneBased(s: []const u8) ?u32 {
    return coords.parseColNumber(s, .{
        .case = .upper_only,
        .max_letters = 3,
    }) catch null;
}

fn requireSheetIdxU32(args: Args, err: *std.Io.Writer, who: []const u8) !u32 {
    if (args.sheet_index == null) {
        try err.print("zlsx: {s} requires --sheet N (0-based)\n", .{who});
        try err.flush();
        return error.SheetIndexMissing;
    }
    if (args.sheet_index.? > std.math.maxInt(u32)) {
        try err.writeAll("zlsx: --sheet value too large (must fit in u32)\n");
        try err.flush();
        return error.SheetIndexTooLarge;
    }
    return @intCast(args.sheet_index.?);
}

fn runRowEditCommand(
    alloc: std.mem.Allocator,
    io: std.Io,
    args: Args,
    err: *std.Io.Writer,
) !u8 {
    const out_path = args.out_path orelse {
        try err.writeAll("zlsx: row-edit requires --out PATH\n");
        try err.flush();
        return 1;
    };
    const sheet_idx = requireSheetIdxU32(args, err, switch (args.subcommand) {
        .insert_row => "insert-row",
        .delete_row => "delete-row",
        else => unreachable,
    }) catch return 1;
    const row = args.row_1based orelse {
        try err.writeAll("zlsx: row-edit requires --row N (1-based)\n");
        try err.flush();
        return 1;
    };

    var ed = zlsx_pkg.Editor.open(alloc, io, args.file) catch |e| {
        try err.print("zlsx: cannot open '{s}': {s}\n", .{ args.file, @errorName(e) });
        try err.flush();
        return openFailureExit(e);
    };
    defer ed.deinit();

    switch (args.subcommand) {
        .insert_row => ed.insertRow(sheet_idx, row) catch |e| {
            try err.print("zlsx: insertRow {d}: {s}\n", .{ row, @errorName(e) });
            try err.flush();
            return 3;
        },
        .delete_row => ed.deleteRow(sheet_idx, row) catch |e| {
            try err.print("zlsx: deleteRow {d}: {s}\n", .{ row, @errorName(e) });
            try err.flush();
            return 3;
        },
        else => unreachable,
    }
    ed.save(io, out_path) catch |e| {
        try err.print("zlsx: save '{s}': {s}\n", .{ out_path, @errorName(e) });
        try err.flush();
        return 5;
    };
    return 0;
}

fn runColEditCommand(
    alloc: std.mem.Allocator,
    io: std.Io,
    args: Args,
    err: *std.Io.Writer,
) !u8 {
    const out_path = args.out_path orelse {
        try err.writeAll("zlsx: col-edit requires --out PATH\n");
        try err.flush();
        return 1;
    };
    const sheet_idx = requireSheetIdxU32(args, err, switch (args.subcommand) {
        .insert_column => "insert-column",
        .delete_column => "delete-column",
        else => unreachable,
    }) catch return 1;
    const col_letter = args.col_letter orelse {
        try err.writeAll("zlsx: col-edit requires --col LETTER (A..XFD)\n");
        try err.flush();
        return 1;
    };
    const col_1based = parseColLettersToOneBased(col_letter) orelse {
        try err.print("zlsx: invalid --col '{s}' (expected A..XFD)\n", .{col_letter});
        try err.flush();
        return 1;
    };

    var ed = zlsx_pkg.Editor.open(alloc, io, args.file) catch |e| {
        try err.print("zlsx: cannot open '{s}': {s}\n", .{ args.file, @errorName(e) });
        try err.flush();
        return openFailureExit(e);
    };
    defer ed.deinit();

    switch (args.subcommand) {
        .insert_column => ed.insertColumn(sheet_idx, col_1based) catch |e| {
            try err.print("zlsx: insertColumn {s}: {s}\n", .{ col_letter, @errorName(e) });
            try err.flush();
            return 3;
        },
        .delete_column => ed.deleteColumn(sheet_idx, col_1based) catch |e| {
            try err.print("zlsx: deleteColumn {s}: {s}\n", .{ col_letter, @errorName(e) });
            try err.flush();
            return 3;
        },
        else => unreachable,
    }
    ed.save(io, out_path) catch |e| {
        try err.print("zlsx: save '{s}': {s}\n", .{ out_path, @errorName(e) });
        try err.flush();
        return 5;
    };
    return 0;
}

fn runAddSheetCommand(
    alloc: std.mem.Allocator,
    io: std.Io,
    args: Args,
    err: *std.Io.Writer,
) !u8 {
    const out_path = args.out_path orelse {
        try err.writeAll("zlsx: add-sheet requires --out PATH\n");
        try err.flush();
        return 1;
    };
    const name = args.new_sheet_name orelse {
        try err.writeAll("zlsx: add-sheet requires --new-name NAME\n");
        try err.flush();
        return 1;
    };

    var ed = zlsx_pkg.Editor.open(alloc, io, args.file) catch |e| {
        try err.print("zlsx: cannot open '{s}': {s}\n", .{ args.file, @errorName(e) });
        try err.flush();
        return openFailureExit(e);
    };
    defer ed.deinit();

    _ = ed.addSheet(name) catch |e| {
        try err.print("zlsx: addSheet '{s}': {s}\n", .{ name, @errorName(e) });
        try err.flush();
        return 3;
    };
    ed.save(io, out_path) catch |e| {
        try err.print("zlsx: save '{s}': {s}\n", .{ out_path, @errorName(e) });
        try err.flush();
        return 5;
    };
    return 0;
}

fn runRenameSheetCommand(
    alloc: std.mem.Allocator,
    io: std.Io,
    args: Args,
    err: *std.Io.Writer,
) !u8 {
    const out_path = args.out_path orelse {
        try err.writeAll("zlsx: rename-sheet requires --out PATH\n");
        try err.flush();
        return 1;
    };
    const sheet_idx = requireSheetIdxU32(args, err, "rename-sheet") catch return 1;
    const new_name = args.new_sheet_name orelse {
        try err.writeAll("zlsx: rename-sheet requires --new-name NAME\n");
        try err.flush();
        return 1;
    };

    var ed = zlsx_pkg.Editor.open(alloc, io, args.file) catch |e| {
        try err.print("zlsx: cannot open '{s}': {s}\n", .{ args.file, @errorName(e) });
        try err.flush();
        return openFailureExit(e);
    };
    defer ed.deinit();

    ed.renameSheet(sheet_idx, new_name) catch |e| {
        try err.print("zlsx: renameSheet {d} -> '{s}': {s}\n", .{ sheet_idx, new_name, @errorName(e) });
        try err.flush();
        return 3;
    };
    ed.save(io, out_path) catch |e| {
        try err.print("zlsx: save '{s}': {s}\n", .{ out_path, @errorName(e) });
        try err.flush();
        return 5;
    };
    return 0;
}

fn runDeleteSheetCommand(
    alloc: std.mem.Allocator,
    io: std.Io,
    args: Args,
    err: *std.Io.Writer,
) !u8 {
    const out_path = args.out_path orelse {
        try err.writeAll("zlsx: delete-sheet requires --out PATH\n");
        try err.flush();
        return 1;
    };
    const sheet_idx = requireSheetIdxU32(args, err, "delete-sheet") catch return 1;

    var ed = zlsx_pkg.Editor.open(alloc, io, args.file) catch |e| {
        try err.print("zlsx: cannot open '{s}': {s}\n", .{ args.file, @errorName(e) });
        try err.flush();
        return openFailureExit(e);
    };
    defer ed.deinit();

    ed.deleteSheet(sheet_idx) catch |e| {
        try err.print("zlsx: deleteSheet {d}: {s}\n", .{ sheet_idx, @errorName(e) });
        try err.flush();
        return 3;
    };
    ed.save(io, out_path) catch |e| {
        try err.print("zlsx: save '{s}': {s}\n", .{ out_path, @errorName(e) });
        try err.flush();
        return 5;
    };
    return 0;
}

/// S3a: `zlsx rename-table-column <file> --table T --old-name A
/// --new-name B --out P`. Routes through `Editor.renameTableColumn`;
/// every error it raises — a selector that names nothing
/// (`TableNotFound`, `TableColumnNotFound`), a name in use
/// (`TableColumnNameInUse`), a name Excel would not take
/// (`InvalidTableColumnName`) — exits 3 like every other failed
/// structural edit.
fn runRenameTableColumnCommand(
    alloc: std.mem.Allocator,
    io: std.Io,
    args: Args,
    err: *std.Io.Writer,
) !u8 {
    const out_path = args.out_path orelse {
        try err.writeAll("zlsx: rename-table-column requires --out PATH\n");
        try err.flush();
        return 1;
    };
    const table = args.table_name orelse {
        try err.writeAll("zlsx: rename-table-column requires --table NAME\n");
        try err.flush();
        return 1;
    };
    const old_name = args.old_column_name orelse {
        try err.writeAll("zlsx: rename-table-column requires --old-name OLD\n");
        try err.flush();
        return 1;
    };
    const new_name = args.new_sheet_name orelse {
        try err.writeAll("zlsx: rename-table-column requires --new-name NEW\n");
        try err.flush();
        return 1;
    };

    var ed = zlsx_pkg.Editor.open(alloc, io, args.file) catch |e| {
        try err.print("zlsx: cannot open '{s}': {s}\n", .{ args.file, @errorName(e) });
        try err.flush();
        return openFailureExit(e);
    };
    defer ed.deinit();

    ed.renameTableColumn(table, old_name, new_name) catch |e| {
        try err.print("zlsx: renameTableColumn {s}[{s}] -> '{s}': {s}\n", .{ table, old_name, new_name, @errorName(e) });
        try err.flush();
        return 3;
    };
    ed.save(io, out_path) catch |e| {
        try err.print("zlsx: save '{s}': {s}\n", .{ out_path, @errorName(e) });
        try err.flush();
        return 5;
    };
    return 0;
}

fn jsonValueToCell(v: std.json.Value) !xlsx.Cell {
    return switch (v) {
        .null => .{ .empty = {} },
        .bool => |b| .{ .boolean = b },
        .integer => |n| .{ .integer = n },
        .float => |f| .{ .number = f },
        .number_string => |s| blk: {
            // std.json hands integers > i64 as `number_string`. Try
            // parseInt first (so "9007199254740993" still routes to
            // the precision check the Editor enforces). If parseInt
            // overflows AND the JSON token is a pure integer literal,
            // refuse rather than silently routing through f64 and
            // rounding — Excel can't represent it exactly either.
            // Pure floats / scientific notation route to .number as
            // expected.
            if (std.fmt.parseInt(i64, s, 10)) |n| {
                break :blk .{ .integer = n };
            } else |_| {
                const looks_like_integer =
                    std.mem.indexOfScalar(u8, s, '.') == null and
                    std.mem.indexOfScalar(u8, s, 'e') == null and
                    std.mem.indexOfScalar(u8, s, 'E') == null;
                if (looks_like_integer) return error.IntegerExceedsI64;
                const f = std.fmt.parseFloat(f64, s) catch return error.UnsupportedJsonNumber;
                break :blk .{ .number = f };
            }
        },
        .string => |s| .{ .string = s },
        .array, .object => return error.UnsupportedJsonType,
    };
}

// ─── Tests ───────────────────────────────────────────────────────────

/// Per-test temporary file helper. Same shape as the helpers in
/// src/writer.zig and src/xlsx.zig — replaces hard-coded /tmp paths
/// so the suite is portable to Windows. Caller frees the returned
/// slice; `defer tt.deinit()` cleans up the directory.
const TestTmp = struct {
    dir: std.testing.TmpDir,
    pub fn init() TestTmp {
        return .{ .dir = std.testing.tmpDir(.{}) };
    }
    pub fn deinit(self: *TestTmp) void {
        self.dir.cleanup();
    }
    pub fn path(self: *TestTmp, alloc: std.mem.Allocator, io: std.Io, name: []const u8) ![:0]u8 {
        const d = try self.dir.dir.realPathFileAlloc(io, ".", alloc);
        defer alloc.free(d);
        return std.fs.path.joinZ(alloc, &.{ d, name });
    }
};

test {
    // Pull `dbx.zig` into the test build's analysis set. Zig analyses
    // lazily: `dbx` is referenced only from `run()`, which no test calls,
    // so without this the compiler never looks at the file and its tests
    // are never collected — they silently did not run from #147 until
    // this reference was added. Any future src/*.zig that hangs off a
    // command dispatch needs the same line.
    _ = dbx;
    _ = formula_cli;
}

test "colLetter A,B,Z,AA,AZ,BA,ZZ,AAA" {
    var buf: [8]u8 = undefined;
    try std.testing.expectEqualStrings("A", colLetter(&buf, 0));
    try std.testing.expectEqualStrings("B", colLetter(&buf, 1));
    try std.testing.expectEqualStrings("Z", colLetter(&buf, 25));
    try std.testing.expectEqualStrings("AA", colLetter(&buf, 26));
    try std.testing.expectEqualStrings("AZ", colLetter(&buf, 51));
    try std.testing.expectEqualStrings("BA", colLetter(&buf, 52));
    try std.testing.expectEqualStrings("ZZ", colLetter(&buf, 701));
    try std.testing.expectEqualStrings("AAA", colLetter(&buf, 702));
}

test "writeJsonString escapes" {
    var scratch: [256]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    try writeJsonString(&w, "hi\n\"\\\t\x01");
    try std.testing.expectEqualStrings("\"hi\\n\\\"\\\\\\t\\u0001\"", w.buffered());
}

test "writeCsvField quoting" {
    var scratch: [256]u8 = undefined;
    {
        var w = std.Io.Writer.fixed(&scratch);
        try writeCsvField(&w, "plain");
        try std.testing.expectEqualStrings("plain", w.buffered());
    }
    {
        var w = std.Io.Writer.fixed(&scratch);
        try writeCsvField(&w, "has,comma");
        try std.testing.expectEqualStrings("\"has,comma\"", w.buffered());
    }
    {
        var w = std.Io.Writer.fixed(&scratch);
        try writeCsvField(&w, "has\"quote");
        try std.testing.expectEqualStrings("\"has\"\"quote\"", w.buffered());
    }
}

test "parseArgs: set-cell subcommand token is skipped, --ref / --value parse" {
    const argv = [_][]const u8{
        "set-cell", "in.xlsx", "--sheet", "0", "--ref", "B1", "--value", "\"hello\"", "--out", "out.xlsx",
    };
    const a = try parseArgs(&argv);
    try std.testing.expectEqual(Subcommand.set_cell, a.subcommand);
    try std.testing.expectEqualStrings("in.xlsx", a.file);
    try std.testing.expectEqual(@as(?usize, 0), a.sheet_index);
    try std.testing.expectEqualStrings("B1", a.cell_ref.?);
    try std.testing.expectEqualStrings("\"hello\"", a.cell_value_json.?);
    try std.testing.expectEqualStrings("out.xlsx", a.out_path.?);
}

test "runSetCellCommand rewrites a single cell and saves to --out" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx.writer_types;
    const src_path = try tt.path(std.testing.allocator, io, "cli_set_cell_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "cli_set_cell_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRow(&.{ .{ .integer = 1 }, .{ .integer = 2 } });
        try s.writeRow(&.{ .{ .integer = 3 }, .{ .integer = 4 } });
        try w.save(io, src_path);
    }
    var err_buf: [1024]u8 = undefined;
    var err_w = std.Io.Writer.fixed(&err_buf);
    const args: Args = .{
        .file = src_path,
        .subcommand = .set_cell,
        .sheet_index = 0,
        .out_path = dst_path,
        .cell_ref = "B1",
        .cell_value_json = "\"hello\"",
    };
    const rc = try runSetCellCommand(std.testing.allocator, io, args, &err_w);
    try std.testing.expectEqual(@as(u8, 0), rc);

    // Verify by re-opening the saved file: B1 must be "hello", A1
    // unchanged, A2 unchanged.
    var book = try xlsx.Book.open(std.testing.allocator, io, dst_path);
    defer book.deinit();
    var rows = try book.rows(book.sheets[0], std.testing.allocator);
    defer rows.deinit();
    const r1 = (try rows.next()) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqual(@as(i64, 1), r1[0].integer);
    try std.testing.expectEqualStrings("hello", r1[1].string);
    const r2 = (try rows.next()) orelse return error.TestUnexpectedResult;
    try std.testing.expectEqual(@as(i64, 3), r2[0].integer);
    try std.testing.expectEqual(@as(i64, 4), r2[1].integer);
}

test "runSetCellCommand rejects missing --ref / --value / --out" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx.writer_types;
    const src_path = try tt.path(std.testing.allocator, io, "cli_set_cell_missing.xlsx");
    defer std.testing.allocator.free(src_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRow(&.{.{ .integer = 1 }});
        try w.save(io, src_path);
    }
    var err_buf: [1024]u8 = undefined;
    {
        var err_w = std.Io.Writer.fixed(&err_buf);
        const a: Args = .{ .file = src_path, .subcommand = .set_cell, .sheet_index = 0, .cell_ref = "A1", .cell_value_json = "1" };
        try std.testing.expectEqual(@as(u8, 1), try runSetCellCommand(std.testing.allocator, io, a, &err_w));
    }
    {
        var err_w = std.Io.Writer.fixed(&err_buf);
        // Dummy out path — set-cell rejects on missing --ref before
        // ever attempting to write, so this never hits the filesystem.
        const a: Args = .{ .file = src_path, .subcommand = .set_cell, .sheet_index = 0, .out_path = "unused.xlsx", .cell_value_json = "1" };
        try std.testing.expectEqual(@as(u8, 1), try runSetCellCommand(std.testing.allocator, io, a, &err_w));
    }
    {
        var err_w = std.Io.Writer.fixed(&err_buf);
        const a: Args = .{ .file = src_path, .subcommand = .set_cell, .sheet_index = 0, .out_path = "unused.xlsx", .cell_ref = "A1" };
        try std.testing.expectEqual(@as(u8, 1), try runSetCellCommand(std.testing.allocator, io, a, &err_w));
    }
}

test "parseArgs rejects --row 0 (1-based contract)" {
    const argv = [_][]const u8{ "insert-row", "in.xlsx", "--sheet", "0", "--row", "0", "--out", "x.xlsx" };
    try std.testing.expectError(ArgError.BadArgValue, parseArgs(&argv));
}

test "parseArgs rejects structural flags on read-only commands" {
    {
        const argv = [_][]const u8{ "rows", "in.xlsx", "--row", "2" };
        try std.testing.expectError(ArgError.BadArgValue, parseArgs(&argv));
    }
    {
        const argv = [_][]const u8{ "cells", "in.xlsx", "--col", "B" };
        try std.testing.expectError(ArgError.BadArgValue, parseArgs(&argv));
    }
    {
        const argv = [_][]const u8{ "meta", "in.xlsx", "--new-name", "Foo" };
        try std.testing.expectError(ArgError.BadArgValue, parseArgs(&argv));
    }
    {
        const argv = [_][]const u8{ "rows", "in.xlsx", "--ref", "A1" };
        try std.testing.expectError(ArgError.BadArgValue, parseArgs(&argv));
    }
    {
        const argv = [_][]const u8{ "rows", "in.xlsx", "--value", "1" };
        try std.testing.expectError(ArgError.BadArgValue, parseArgs(&argv));
    }
}

test "parseColLettersToOneBased: A=1, Z=26, AA=27, XFD=16384, XFE rejected" {
    try std.testing.expectEqual(@as(?u32, 1), parseColLettersToOneBased("A"));
    try std.testing.expectEqual(@as(?u32, 26), parseColLettersToOneBased("Z"));
    try std.testing.expectEqual(@as(?u32, 27), parseColLettersToOneBased("AA"));
    try std.testing.expectEqual(@as(?u32, 16384), parseColLettersToOneBased("XFD"));
    try std.testing.expectEqual(@as(?u32, null), parseColLettersToOneBased("XFE"));
    try std.testing.expectEqual(@as(?u32, null), parseColLettersToOneBased(""));
    try std.testing.expectEqual(@as(?u32, null), parseColLettersToOneBased("a"));
    try std.testing.expectEqual(@as(?u32, null), parseColLettersToOneBased("ABCD"));
}

test "runRowEditCommand insert-row + delete-row round-trip" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx.writer_types;
    const src_path = try tt.path(std.testing.allocator, io, "cli_row_edit_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const dst_path = try tt.path(std.testing.allocator, io, "cli_row_edit_dst.xlsx");
    defer std.testing.allocator.free(dst_path);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("S");
        try s.writeRow(&.{.{ .integer = 1 }});
        try s.writeRow(&.{.{ .integer = 2 }});
        try s.writeRow(&.{.{ .integer = 3 }});
        try w.save(io, src_path);
    }
    var err_buf: [1024]u8 = undefined;

    // insert-row 2 → expect [1, empty, 2, 3]
    {
        var err_w = std.Io.Writer.fixed(&err_buf);
        const a: Args = .{
            .file = src_path,
            .subcommand = .insert_row,
            .sheet_index = 0,
            .out_path = dst_path,
            .row_1based = 2,
        };
        try std.testing.expectEqual(@as(u8, 0), try runRowEditCommand(std.testing.allocator, io, a, &err_w));
    }
    {
        var book = try xlsx.Book.open(std.testing.allocator, io, dst_path);
        defer book.deinit();
        var rows = try book.rows(book.sheets[0], std.testing.allocator);
        defer rows.deinit();
        const r1 = (try rows.next()).?;
        try std.testing.expectEqual(@as(i64, 1), r1[0].integer);
        // The inserted row may surface as a blank row entirely — the
        // reader's iteration contract skips empty rows, so the next
        // emitted row should be the original "2".
        const r2 = (try rows.next()).?;
        try std.testing.expectEqual(@as(i64, 2), r2[0].integer);
    }
}

test "runAddSheetCommand + runRenameSheetCommand + runDeleteSheetCommand round-trip" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const writer_mod = xlsx.writer_types;
    const src_path = try tt.path(std.testing.allocator, io, "cli_sheet_ops_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const after_add = try tt.path(std.testing.allocator, io, "cli_sheet_ops_add.xlsx");
    defer std.testing.allocator.free(after_add);
    const after_rename = try tt.path(std.testing.allocator, io, "cli_sheet_ops_rename.xlsx");
    defer std.testing.allocator.free(after_rename);
    const after_delete = try tt.path(std.testing.allocator, io, "cli_sheet_ops_delete.xlsx");
    defer std.testing.allocator.free(after_delete);
    {
        var w = writer_mod.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("First");
        try s.writeRow(&.{.{ .integer = 1 }});
        try w.save(io, src_path);
    }
    var err_buf: [1024]u8 = undefined;

    // add-sheet "Second"
    {
        var err_w = std.Io.Writer.fixed(&err_buf);
        const a: Args = .{
            .file = src_path,
            .subcommand = .add_sheet,
            .out_path = after_add,
            .new_sheet_name = "Second",
        };
        try std.testing.expectEqual(@as(u8, 0), try runAddSheetCommand(std.testing.allocator, io, a, &err_w));
    }
    {
        var book = try xlsx.Book.open(std.testing.allocator, io, after_add);
        defer book.deinit();
        try std.testing.expectEqual(@as(usize, 2), book.sheets.len);
        try std.testing.expectEqualStrings("Second", book.sheets[1].name);
    }

    // rename-sheet 0 -> "Renamed"
    {
        var err_w = std.Io.Writer.fixed(&err_buf);
        const a: Args = .{
            .file = after_add,
            .subcommand = .rename_sheet,
            .sheet_index = 0,
            .out_path = after_rename,
            .new_sheet_name = "Renamed",
        };
        try std.testing.expectEqual(@as(u8, 0), try runRenameSheetCommand(std.testing.allocator, io, a, &err_w));
    }
    {
        var book = try xlsx.Book.open(std.testing.allocator, io, after_rename);
        defer book.deinit();
        try std.testing.expectEqualStrings("Renamed", book.sheets[0].name);
    }

    // delete-sheet 1 (drops "Second")
    {
        var err_w = std.Io.Writer.fixed(&err_buf);
        const a: Args = .{
            .file = after_rename,
            .subcommand = .delete_sheet,
            .sheet_index = 1,
            .out_path = after_delete,
        };
        try std.testing.expectEqual(@as(u8, 0), try runDeleteSheetCommand(std.testing.allocator, io, a, &err_w));
    }
    {
        var book = try xlsx.Book.open(std.testing.allocator, io, after_delete);
        defer book.deinit();
        try std.testing.expectEqual(@as(usize, 1), book.sheets.len);
        try std.testing.expectEqualStrings("Renamed", book.sheets[0].name);
    }
}

test "runRenameTableColumnCommand: the table part follows the new name; a missing table exits 3; flags are parsed and fenced" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const src_path = try tt.path(std.testing.allocator, io, "cli_rtc_src.xlsx");
    defer std.testing.allocator.free(src_path);
    const out_path = try tt.path(std.testing.allocator, io, "cli_rtc_out.xlsx");
    defer std.testing.allocator.free(out_path);
    try zlsx_pkg.pivots.fixture.write(std.testing.allocator, io, src_path, .table_name);

    var err_buf: [512]u8 = undefined;
    {
        var err_w = std.Io.Writer.fixed(&err_buf);
        const argv = [_][]const u8{ "rename-table-column", src_path, "--table", "SalesTbl", "--old-name", "Qty", "--new-name", "Quantity", "--out", out_path };
        const a = try parseArgs(&argv);
        try std.testing.expectEqual(Subcommand.rename_table_column, a.subcommand);
        try std.testing.expectEqualStrings("SalesTbl", a.table_name.?);
        try std.testing.expectEqualStrings("Qty", a.old_column_name.?);
        try std.testing.expectEqualStrings("Quantity", a.new_sheet_name.?);
        try std.testing.expectEqual(@as(u8, 0), try runRenameTableColumnCommand(std.testing.allocator, io, a, &err_w));
    }
    {
        var wb = try zlsx_pkg.Workbook.open(std.testing.allocator, io, out_path);
        defer wb.deinit();
        const part = (try wb.store.part("xl/tables/table1.xml")) orelse return error.TestUnexpectedResult;
        try std.testing.expect(std.mem.indexOf(u8, part.bytes, "name=\"Quantity\"") != null);
        try std.testing.expect(std.mem.indexOf(u8, part.bytes, "name=\"Qty\"") == null);
    }
    // A table the workbook does not have: exit 3, the name in the diagnostic.
    {
        var err_w = std.Io.Writer.fixed(&err_buf);
        const a: Args = .{ .file = src_path, .subcommand = .rename_table_column, .out_path = out_path, .table_name = "Nope", .old_column_name = "Qty", .new_sheet_name = "Q" };
        try std.testing.expectEqual(@as(u8, 3), try runRenameTableColumnCommand(std.testing.allocator, io, a, &err_w));
        try std.testing.expect(std.mem.indexOf(u8, err_w.buffered(), "TableNotFound") != null);
    }
    // Each required flag missing is exit 1.
    {
        var err_w = std.Io.Writer.fixed(&err_buf);
        const a: Args = .{ .file = src_path, .subcommand = .rename_table_column, .out_path = out_path, .table_name = "SalesTbl", .new_sheet_name = "Q" };
        try std.testing.expectEqual(@as(u8, 1), try runRenameTableColumnCommand(std.testing.allocator, io, a, &err_w));
    }
    // The table flags belong to this sub-command alone.
    {
        const argv = [_][]const u8{ "rows", "f.xlsx", "--table", "T" };
        try std.testing.expectError(ArgError.BadArgValue, parseArgs(&argv));
    }
    {
        const argv = [_][]const u8{ "rename-sheet", "f.xlsx", "--sheet", "0", "--old-name", "T", "--new-name", "U", "--out", "o.xlsx" };
        try std.testing.expectError(ArgError.BadArgValue, parseArgs(&argv));
    }
}

test "parseArgs basics" {
    const argv = [_][]const u8{ "file.xlsx", "--sheet", "2", "--format", "csv" };
    const a = try parseArgs(&argv);
    try std.testing.expectEqualStrings("file.xlsx", a.file);
    try std.testing.expectEqual(@as(?usize, 2), a.sheet_index);
    try std.testing.expectEqual(Format.csv, a.format);
}

test "parseArgs accepts --key=value syntax" {
    // GNU-style --key=value should work uniformly with --key value.
    const argv = [_][]const u8{ "file.xlsx", "--sheet=2", "--format=csv", "--take=10" };
    const a = try parseArgs(&argv);
    try std.testing.expectEqualStrings("file.xlsx", a.file);
    try std.testing.expectEqual(@as(?usize, 2), a.sheet_index);
    try std.testing.expectEqual(Format.csv, a.format);
    try std.testing.expectEqual(@as(?usize, 10), a.take);
}

test "parseArgs --key= with empty value is accepted as empty string" {
    // `--out=` produces an empty string value — let the downstream
    // path-handling reject empties, not the parser. Use a sub-
    // command that actually accepts --out (post-iter65 the parser
    // rejects --out on read-only commands).
    const argv = [_][]const u8{ "set-cell", "in.xlsx", "--sheet", "0", "--ref", "B1", "--value", "1", "--out=" };
    const a = try parseArgs(&argv);
    try std.testing.expectEqualStrings("", a.out_path.?);
}

test "parseArgs preserves unknown --foo=bar as literal value" {
    // `--name --Q=1` should set the sheet name to literally "--Q=1",
    // NOT split "--Q" as a flag (which would steal the literal from
    // --name and leave 1 dangling as a positional). Only known
    // value-flags get the `=` split.
    const argv = [_][]const u8{ "rows", "f.xlsx", "--name", "--Q=1" };
    const a = try parseArgs(&argv);
    try std.testing.expectEqualStrings("--Q=1", a.sheet_name.?);
}

test "parseArgs preserves --known-flag=value as literal value when used after value-flag" {
    // `--name --format=csv` should set sheet name to literally
    // "--format=csv" (rare but valid xlsx sheet name), NOT split
    // --format=csv into [--format, csv]. The context-aware splitter
    // must see "the previous token is a value-bearing flag" and
    // leave THIS token verbatim regardless of shape.
    const argv = [_][]const u8{ "rows", "f.xlsx", "--name", "--format=csv" };
    const a = try parseArgs(&argv);
    try std.testing.expectEqualStrings("--format=csv", a.sheet_name.?);
}

test "parseArgs preserves --bool=value as literal when used after value-flag" {
    // Same context rule: `--name --header=1` is a literal sheet name
    // even though `--header=1` would otherwise trigger BadArgValue
    // (since --header is a boolean flag).
    const argv = [_][]const u8{ "rows", "f.xlsx", "--name", "--header=1" };
    const a = try parseArgs(&argv);
    try std.testing.expectEqualStrings("--header=1", a.sheet_name.?);
}

test "parseArgs rejects --out on read-only commands" {
    // `zlsx rows in.xlsx --out out.jsonl` would silently swallow
    // --out and exit 0 without writing the file. Reject as
    // BadArgValue so the user sees the typo / mis-use.
    {
        const argv = [_][]const u8{ "rows", "f.xlsx", "--out", "out.jsonl" };
        try std.testing.expectError(ArgError.BadArgValue, parseArgs(&argv));
    }
    {
        const argv = [_][]const u8{ "meta", "f.xlsx", "--out", "out.jsonl" };
        try std.testing.expectError(ArgError.BadArgValue, parseArgs(&argv));
    }
    // Edit commands continue to accept --out.
    {
        const argv = [_][]const u8{ "set-cell", "in.xlsx", "--sheet", "0", "--ref", "B1", "--value", "1", "--out", "out.xlsx" };
        const a = try parseArgs(&argv);
        try std.testing.expectEqualStrings("out.xlsx", a.out_path.?);
    }
}

test "parseArgs rejects =value on boolean flags" {
    // `--list-sheets=false` is invalid syntax — the value would
    // otherwise leak into the positional file slot and pick the
    // wrong workbook.
    {
        const argv = [_][]const u8{ "file.xlsx", "--list-sheets=false" };
        try std.testing.expectError(ArgError.BadArgValue, parseArgs(&argv));
    }
    {
        const argv = [_][]const u8{ "file.xlsx", "--header=1" };
        try std.testing.expectError(ArgError.BadArgValue, parseArgs(&argv));
    }
    // Bare `--list-sheets` without `=` still works (boolean flag).
    {
        const argv = [_][]const u8{ "file.xlsx", "--list-sheets" };
        const a = try parseArgs(&argv);
        try std.testing.expect(a.list_sheets);
    }
}

test "parseArgs rejects both --sheet and --name" {
    const argv = [_][]const u8{ "f.xlsx", "--sheet", "0", "--name", "Sheet1" };
    try std.testing.expectError(ArgError.SheetArgConflict, parseArgs(&argv));
}

test "parseArgs help" {
    const argv = [_][]const u8{"-h"};
    try std.testing.expectError(ArgError.HelpRequested, parseArgs(&argv));
}

test "parseArgs maps jsonl to envelope and legacy-jsonl to bare array" {
    {
        const argv = [_][]const u8{ "f.xlsx", "--format", "jsonl" };
        const a = try parseArgs(&argv);
        try std.testing.expectEqual(Format.jsonl, a.format);
    }
    {
        const argv = [_][]const u8{ "f.xlsx", "--format", "legacy-jsonl" };
        const a = try parseArgs(&argv);
        try std.testing.expectEqual(Format.legacy_jsonl, a.format);
    }
    {
        const argv = [_][]const u8{ "f.xlsx", "--format", "legacy-jsonl-dict" };
        const a = try parseArgs(&argv);
        try std.testing.expectEqual(Format.legacy_jsonl_dict, a.format);
    }
    {
        // Deprecated alias still lands on the bare-dict path AND
        // flips the deprecation flag for main's stderr warning.
        const argv = [_][]const u8{ "f.xlsx", "--format", "jsonl-dict" };
        const a = try parseArgs(&argv);
        try std.testing.expectEqual(Format.legacy_jsonl_dict, a.format);
        try std.testing.expect(a.deprecated_jsonl_dict);
    }
    {
        // Canonical `legacy-jsonl-dict` must NOT trip the warning.
        const argv = [_][]const u8{ "f.xlsx", "--format", "legacy-jsonl-dict" };
        const a = try parseArgs(&argv);
        try std.testing.expect(!a.deprecated_jsonl_dict);
    }
}

test "writeRowEnvelope emits kind + sheet + sheet_idx + row + sparse cells" {
    var scratch: [512]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    const cells = [_]xlsx.Cell{
        .{ .string = "name" },
        .{ .integer = 42 },
        .empty, // sparse — must be skipped in the cells array
        .{ .number = 3.5 },
        .{ .boolean = true },
    };
    try writeRowEnvelope(&w, "Data", 0, 1, &cells, false, null, 0, false, &.{}, &.{}, &.{}, &.{}, false);
    const expected =
        "{\"kind\":\"row\",\"sheet\":\"Data\",\"sheet_idx\":0,\"row\":1,\"cells\":[" ++
        "{\"ref\":\"A1\",\"col\":1,\"t\":\"str\",\"v\":\"name\"}," ++
        "{\"ref\":\"B1\",\"col\":2,\"t\":\"int\",\"v\":42}," ++
        "{\"ref\":\"D1\",\"col\":4,\"t\":\"num\",\"v\":3.5}," ++
        "{\"ref\":\"E1\",\"col\":5,\"t\":\"bool\",\"v\":true}" ++
        "]}\n";
    try std.testing.expectEqualStrings(expected, w.buffered());
}

test "writeRowEnvelope all-empty row emits envelope with empty cells array" {
    var scratch: [256]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    const cells = [_]xlsx.Cell{ .empty, .empty, .empty };
    try writeRowEnvelope(&w, "S", 2, 7, &cells, false, null, 0, false, &.{}, &.{}, &.{}, &.{}, false);
    try std.testing.expectEqualStrings(
        "{\"kind\":\"row\",\"sheet\":\"S\",\"sheet_idx\":2,\"row\":7,\"cells\":[]}\n",
        w.buffered(),
    );
}

test "writeRowEnvelope escapes sheet name" {
    var scratch: [256]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    const cells = [_]xlsx.Cell{.{ .integer = 1 }};
    try writeRowEnvelope(&w, "She\"et\n", 0, 1, &cells, false, null, 0, false, &.{}, &.{}, &.{}, &.{}, false);
    try std.testing.expectEqualStrings(
        "{\"kind\":\"row\",\"sheet\":\"She\\\"et\\n\",\"sheet_idx\":0,\"row\":1,\"cells\":[" ++
            "{\"ref\":\"A1\",\"col\":1,\"t\":\"int\",\"v\":1}" ++
            "]}\n",
        w.buffered(),
    );
}

test "writeRowEnvelope non-finite number becomes null v" {
    var scratch: [256]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    const cells = [_]xlsx.Cell{.{ .number = std.math.nan(f64) }};
    try writeRowEnvelope(&w, "S", 0, 1, &cells, false, null, 0, false, &.{}, &.{}, &.{}, &.{}, false);
    // `t` stays `"num"` for the non-finite case — the type of the
    // cell didn't change, only its JSON-serializable value did.
    // This matches the pre-iter55a behaviour of writeJsonCell.
    try std.testing.expectEqualStrings(
        "{\"kind\":\"row\",\"sheet\":\"S\",\"sheet_idx\":0,\"row\":1,\"cells\":[" ++
            "{\"ref\":\"A1\",\"col\":1,\"t\":\"num\",\"v\":null}" ++
            "]}\n",
        w.buffered(),
    );
}

test "parseArgs routes 'cells' as the cells sub-command" {
    // Bare file-path defaults to rows (back-compat).
    {
        const argv = [_][]const u8{"file.xlsx"};
        const a = try parseArgs(&argv);
        try std.testing.expectEqual(Subcommand.rows, a.subcommand);
        try std.testing.expectEqualStrings("file.xlsx", a.file);
    }
    // Explicit `rows` is parsed as rows, file-path is the next positional.
    {
        const argv = [_][]const u8{ "rows", "file.xlsx" };
        const a = try parseArgs(&argv);
        try std.testing.expectEqual(Subcommand.rows, a.subcommand);
        try std.testing.expectEqualStrings("file.xlsx", a.file);
    }
    // `cells` flips the sub-command.
    {
        const argv = [_][]const u8{ "cells", "file.xlsx" };
        const a = try parseArgs(&argv);
        try std.testing.expectEqual(Subcommand.cells, a.subcommand);
        try std.testing.expectEqualStrings("file.xlsx", a.file);
    }
    // `cells` with flags behind it.
    {
        const argv = [_][]const u8{ "cells", "file.xlsx", "--sheet", "2" };
        const a = try parseArgs(&argv);
        try std.testing.expectEqual(Subcommand.cells, a.subcommand);
        try std.testing.expectEqualStrings("file.xlsx", a.file);
        try std.testing.expectEqual(@as(?usize, 2), a.sheet_index);
    }
    // Flags before the sub-command still work — first POSITIONAL is
    // what decides, not the first argv slot.
    {
        const argv = [_][]const u8{ "--sheet", "1", "cells", "file.xlsx" };
        const a = try parseArgs(&argv);
        try std.testing.expectEqual(Subcommand.cells, a.subcommand);
        try std.testing.expectEqualStrings("file.xlsx", a.file);
    }
}

test "writeCell emits kind + sheet + sheet_idx + ref + row + col + t + v" {
    var scratch: [512]u8 = undefined;

    // string
    {
        var w = std.Io.Writer.fixed(&scratch);
        try writeCell(&w, "Data", 0, "A1", 1, 1, .{ .string = "name" }, null, false, false, null, null, null, false);
        try std.testing.expectEqualStrings(
            "{\"kind\":\"cell\",\"sheet\":\"Data\",\"sheet_idx\":0,\"ref\":\"A1\",\"row\":1,\"col\":1,\"t\":\"str\",\"v\":\"name\"}\n",
            w.buffered(),
        );
    }
    // integer
    {
        var w = std.Io.Writer.fixed(&scratch);
        try writeCell(&w, "Data", 0, "B2", 2, 2, .{ .integer = 3 }, null, false, false, null, null, null, false);
        try std.testing.expectEqualStrings(
            "{\"kind\":\"cell\",\"sheet\":\"Data\",\"sheet_idx\":0,\"ref\":\"B2\",\"row\":2,\"col\":2,\"t\":\"int\",\"v\":3}\n",
            w.buffered(),
        );
    }
    // number
    {
        var w = std.Io.Writer.fixed(&scratch);
        try writeCell(&w, "Data", 0, "C3", 3, 3, .{ .number = 3.5 }, null, false, false, null, null, null, false);
        try std.testing.expectEqualStrings(
            "{\"kind\":\"cell\",\"sheet\":\"Data\",\"sheet_idx\":0,\"ref\":\"C3\",\"row\":3,\"col\":3,\"t\":\"num\",\"v\":3.5}\n",
            w.buffered(),
        );
    }
    // boolean
    {
        var w = std.Io.Writer.fixed(&scratch);
        try writeCell(&w, "Data", 0, "D4", 4, 4, .{ .boolean = true }, null, false, false, null, null, null, false);
        try std.testing.expectEqualStrings(
            "{\"kind\":\"cell\",\"sheet\":\"Data\",\"sheet_idx\":0,\"ref\":\"D4\",\"row\":4,\"col\":4,\"t\":\"bool\",\"v\":true}\n",
            w.buffered(),
        );
    }
}

test "writeCell escapes sheet name" {
    var scratch: [256]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    try writeCell(&w, "She\"et\n", 2, "A1", 1, 1, .{ .integer = 7 }, null, false, false, null, null, null, false);
    try std.testing.expectEqualStrings(
        "{\"kind\":\"cell\",\"sheet\":\"She\\\"et\\n\",\"sheet_idx\":2,\"ref\":\"A1\",\"row\":1,\"col\":1,\"t\":\"int\",\"v\":7}\n",
        w.buffered(),
    );
}

test "cells loop skips empty cells from the stream" {
    // Mirrors runCellsCommand's inner loop: feed a mixed row, confirm
    // only non-empty cells surface, refs are built from (col,row_number).
    var scratch: [512]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);

    const cells = [_]xlsx.Cell{
        .{ .string = "name" }, // A1
        .empty, // B1 — must produce no output
        .{ .integer = 42 }, // C1
        .empty, // D1 — must produce no output
        .{ .boolean = false }, // E1
    };
    const row_number: u32 = 1;
    for (cells, 0..) |c, i| {
        if (c == .empty) continue;
        var col_buf: [8]u8 = undefined;
        const letters = colLetter(&col_buf, i);
        var ref_buf: [16]u8 = undefined;
        const ref = std.fmt.bufPrint(&ref_buf, "{s}{d}", .{ letters, row_number }) catch unreachable;
        try writeCell(&w, "S", 0, ref, row_number, @intCast(i + 1), c, null, false, false, null, null, null, false);
    }

    try std.testing.expectEqualStrings(
        "{\"kind\":\"cell\",\"sheet\":\"S\",\"sheet_idx\":0,\"ref\":\"A1\",\"row\":1,\"col\":1,\"t\":\"str\",\"v\":\"name\"}\n" ++
            "{\"kind\":\"cell\",\"sheet\":\"S\",\"sheet_idx\":0,\"ref\":\"C1\",\"row\":1,\"col\":3,\"t\":\"int\",\"v\":42}\n" ++
            "{\"kind\":\"cell\",\"sheet\":\"S\",\"sheet_idx\":0,\"ref\":\"E1\",\"row\":1,\"col\":5,\"t\":\"bool\",\"v\":false}\n",
        w.buffered(),
    );
}

test "writeRow legacy-jsonl produces bare arrays (regression guard)" {
    var scratch: [256]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    const cells = [_]xlsx.Cell{
        .{ .string = "x" },
        .empty,
        .{ .integer = 9 },
    };
    try writeRow(&w, &cells, .legacy_jsonl, 0);
    try std.testing.expectEqualStrings("[\"x\", null, 9]\n", w.buffered());
}

test "writeRow legacy-jsonl-dict produces bare objects (regression guard)" {
    var scratch: [256]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    const cells = [_]xlsx.Cell{
        .{ .string = "x" },
        .empty,
        .{ .integer = 9 },
    };
    try writeRow(&w, &cells, .legacy_jsonl_dict, 0);
    try std.testing.expectEqualStrings("{\"A\": \"x\", \"C\": 9}\n", w.buffered());
}

test "parseArgs routes 'meta' and 'list-sheets' correctly" {
    // `meta` as first positional.
    {
        const argv = [_][]const u8{ "meta", "file.xlsx" };
        const a = try parseArgs(&argv);
        try std.testing.expectEqual(Subcommand.meta, a.subcommand);
        try std.testing.expectEqualStrings("file.xlsx", a.file);
    }
    // `list-sheets` as first positional flips the sub-command.
    {
        const argv = [_][]const u8{ "list-sheets", "file.xlsx" };
        const a = try parseArgs(&argv);
        try std.testing.expectEqual(Subcommand.list_sheets, a.subcommand);
        try std.testing.expectEqualStrings("file.xlsx", a.file);
    }
    // Sub-command token AFTER flags still works (positional decides).
    {
        const argv = [_][]const u8{ "--sheet", "1", "meta", "file.xlsx" };
        const a = try parseArgs(&argv);
        try std.testing.expectEqual(Subcommand.meta, a.subcommand);
    }
    // Legacy `--list-sheets` flag is NOT the `list-sheets` sub-command.
    // The flag flips `list_sheets` (legacy plain text), not `subcommand`.
    {
        const argv = [_][]const u8{ "--list-sheets", "file.xlsx" };
        const a = try parseArgs(&argv);
        try std.testing.expect(a.list_sheets);
        try std.testing.expectEqual(Subcommand.rows, a.subcommand);
    }
}

test "parseArgs tolerates bogus --sheet / --format values on workbook-scoped sub-commands" {
    // Wrappers that append --sheet/--format universally must still
    // reach `meta` / `list-sheets` without an exit-1. Values are
    // silently dropped on those sub-commands, not validated.
    {
        const argv = [_][]const u8{ "meta", "f.xlsx", "--sheet", "nope" };
        const a = try parseArgs(&argv);
        try std.testing.expectEqual(Subcommand.meta, a.subcommand);
        try std.testing.expect(a.sheet_index == null);
    }
    {
        const argv = [_][]const u8{ "list-sheets", "f.xlsx", "--format", "bogus" };
        const a = try parseArgs(&argv);
        try std.testing.expectEqual(Subcommand.list_sheets, a.subcommand);
    }
    // Non-workbook-scoped commands stay strict — bogus --sheet still
    // errors on `rows` / `cells`.
    {
        const argv = [_][]const u8{ "cells", "f.xlsx", "--sheet", "nope" };
        try std.testing.expectError(ArgError.BadSheetIndex, parseArgs(&argv));
    }
    {
        const argv = [_][]const u8{ "rows", "f.xlsx", "--format", "bogus" };
        try std.testing.expectError(ArgError.BadFormat, parseArgs(&argv));
    }
}

test "parseArgs --skip / --take round-trip and tolerance" {
    // Both flags parse as usize and live on Args.
    {
        const argv = [_][]const u8{ "rows", "f.xlsx", "--skip", "5", "--take", "10" };
        const a = try parseArgs(&argv);
        try std.testing.expectEqual(@as(?usize, 5), a.skip);
        try std.testing.expectEqual(@as(?usize, 10), a.take);
    }
    // Bogus --skip / --take are hard errors on record-scoped commands.
    {
        const argv = [_][]const u8{ "rows", "f.xlsx", "--skip", "bogus" };
        try std.testing.expectError(ArgError.BadArgValue, parseArgs(&argv));
    }
    {
        const argv = [_][]const u8{ "cells", "f.xlsx", "--take", "nope" };
        try std.testing.expectError(ArgError.BadArgValue, parseArgs(&argv));
    }
    // --skip / --take are strict on every sub-command (unlike
    // --sheet / --format whose tolerance follows the workbook_scoped
    // group). Pagination is too useful on styles / sst — a typoed
    // --take that silently returned the full stream would be an
    // expensive surprise. On meta / list-sheets which don't paginate,
    // the error is also the clearer signal than silent no-op.
    inline for (.{ "meta", "list-sheets", "styles", "sst" }) |cmd| {
        {
            const argv = [_][]const u8{ cmd, "f.xlsx", "--skip", "bogus" };
            try std.testing.expectError(ArgError.BadArgValue, parseArgs(&argv));
        }
        {
            const argv = [_][]const u8{ cmd, "f.xlsx", "--take", "nope" };
            try std.testing.expectError(ArgError.BadArgValue, parseArgs(&argv));
        }
    }
    // --skip and --take default to null when absent — legacy callers
    // must see identical behavior to pre-iter59a.
    {
        const argv = [_][]const u8{ "cells", "f.xlsx" };
        const a = try parseArgs(&argv);
        try std.testing.expect(a.skip == null);
        try std.testing.expect(a.take == null);
    }
}

test "parseArgs --start-row / --end-row round-trip and rejections" {
    // Happy path: both parse as u32 and live on Args.
    {
        const argv = [_][]const u8{ "rows", "f.xlsx", "--start-row", "5", "--end-row", "10" };
        const a = try parseArgs(&argv);
        try std.testing.expectEqual(@as(?u32, 5), a.start_row);
        try std.testing.expectEqual(@as(?u32, 10), a.end_row);
    }
    // Bogus values error (strict on every sub-command).
    {
        const argv = [_][]const u8{ "rows", "f.xlsx", "--start-row", "bogus" };
        try std.testing.expectError(ArgError.BadArgValue, parseArgs(&argv));
    }
    {
        const argv = [_][]const u8{ "cells", "f.xlsx", "--end-row", "nope" };
        try std.testing.expectError(ArgError.BadArgValue, parseArgs(&argv));
    }
    // 0 is a user error: OOXML rows are 1-based.
    {
        const argv = [_][]const u8{ "rows", "f.xlsx", "--start-row", "0" };
        try std.testing.expectError(ArgError.BadArgValue, parseArgs(&argv));
    }
    // start_row > end_row is an empty emission range — caught at parse.
    {
        const argv = [_][]const u8{ "cells", "f.xlsx", "--start-row", "10", "--end-row", "5" };
        try std.testing.expectError(ArgError.BadArgValue, parseArgs(&argv));
    }
    // start_row == end_row is a valid single-row slice.
    {
        const argv = [_][]const u8{ "cells", "f.xlsx", "--start-row", "7", "--end-row", "7" };
        const a = try parseArgs(&argv);
        try std.testing.expectEqual(@as(?u32, 7), a.start_row);
        try std.testing.expectEqual(@as(?u32, 7), a.end_row);
    }
    // Sub-commands without a row key reject --start-row / --end-row.
    inline for (.{ "validations", "hyperlinks", "pivots", "merges", "defined-names", "doc-props", "meta", "list-sheets", "styles", "sst" }) |cmd| {
        {
            const argv = [_][]const u8{ cmd, "f.xlsx", "--start-row", "2" };
            try std.testing.expectError(ArgError.BadArgValue, parseArgs(&argv));
        }
        {
            const argv = [_][]const u8{ cmd, "f.xlsx", "--end-row", "5" };
            try std.testing.expectError(ArgError.BadArgValue, parseArgs(&argv));
        }
    }
    // Explicitly allowed on the three row-keyed sub-commands.
    inline for (.{ "rows", "cells", "comments" }) |cmd| {
        const argv = [_][]const u8{ cmd, "f.xlsx", "--start-row", "2", "--end-row", "4" };
        const a = try parseArgs(&argv);
        try std.testing.expectEqual(@as(?u32, 2), a.start_row);
        try std.testing.expectEqual(@as(?u32, 4), a.end_row);
    }
    // Defaults to null when absent.
    {
        const argv = [_][]const u8{ "cells", "f.xlsx" };
        const a = try parseArgs(&argv);
        try std.testing.expect(a.start_row == null);
        try std.testing.expect(a.end_row == null);
    }
    // Legacy --list-sheets flag takes the early-return path in main
    // and emits plain sheet names; row bounds passed alongside it
    // would silently no-op. parseArgs must reject.
    {
        const argv = [_][]const u8{ "f.xlsx", "--list-sheets", "--start-row", "2" };
        try std.testing.expectError(ArgError.BadArgValue, parseArgs(&argv));
    }
    {
        const argv = [_][]const u8{ "f.xlsx", "--list-sheets", "--end-row", "10" };
        try std.testing.expectError(ArgError.BadArgValue, parseArgs(&argv));
    }
}

test "runCellsCommand --start-row / --end-row bound the emitted cell stream" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, io, "cli_rowbounds_iter59b.xlsx");
    defer std.testing.allocator.free(tmp_path);
    {
        const writer = xlsx.writer_types;
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s0 = try w.addSheet("Data");
        // 5 rows × 1 cell each → rows 1..5 in the OOXML sense.
        try s0.writeRow(&.{.{ .string = "c1" }});
        try s0.writeRow(&.{.{ .string = "c2" }});
        try s0.writeRow(&.{.{ .string = "c3" }});
        try s0.writeRow(&.{.{ .string = "c4" }});
        try s0.writeRow(&.{.{ .string = "c5" }});
        try w.save(io, tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, io, tmp_path);
    defer book.deinit();

    const countLines = struct {
        fn f(s: []const u8) usize {
            var n: usize = 0;
            for (s) |c| if (c == '\n') {
                n += 1;
            };
            return n;
        }
    }.f;

    // --start-row 2 --end-row 4 → rows 2, 3, 4.
    {
        var scratch: [4096]u8 = undefined;
        var w = std.Io.Writer.fixed(&scratch);
        try runCellsCommand(&w, &book, book.sheets[0], 0, std.testing.allocator, null, null, 2, 4, null, false, false);
        const out = w.buffered();
        try std.testing.expectEqual(@as(usize, 3), countLines(out));
        try std.testing.expect(std.mem.indexOf(u8, out, "\"v\":\"c1\"") == null);
        try std.testing.expect(std.mem.indexOf(u8, out, "\"v\":\"c2\"") != null);
        try std.testing.expect(std.mem.indexOf(u8, out, "\"v\":\"c3\"") != null);
        try std.testing.expect(std.mem.indexOf(u8, out, "\"v\":\"c4\"") != null);
        try std.testing.expect(std.mem.indexOf(u8, out, "\"v\":\"c5\"") == null);
    }
    // Row bounds run BEFORE --skip/--take. Of rows 2/3/4, --skip 1
    // drops c2 and --take 1 keeps exactly c3.
    {
        var scratch: [4096]u8 = undefined;
        var w = std.Io.Writer.fixed(&scratch);
        try runCellsCommand(&w, &book, book.sheets[0], 0, std.testing.allocator, 1, 1, 2, 4, null, false, false);
        const out = w.buffered();
        try std.testing.expectEqual(@as(usize, 1), countLines(out));
        try std.testing.expect(std.mem.indexOf(u8, out, "\"v\":\"c3\"") != null);
    }
}

test "parseArgs --range round-trip and rejections" {
    // Happy path: `A1:C10` parses on rows / cells.
    inline for (.{ "rows", "cells" }) |cmd| {
        const argv = [_][]const u8{ cmd, "f.xlsx", "--range", "A1:C10" };
        const a = try parseArgs(&argv);
        try std.testing.expect(a.range != null);
        try std.testing.expectEqual(@as(u32, 0), a.range.?.top_left.col);
        try std.testing.expectEqual(@as(u32, 1), a.range.?.top_left.row);
        try std.testing.expectEqual(@as(u32, 2), a.range.?.bottom_right.col);
        try std.testing.expectEqual(@as(u32, 10), a.range.?.bottom_right.row);
    }
    // Malformed input.
    inline for (.{ "bogus", "A1-C10", "", ":", "A1:", ":B2", "a1:b2", "A0:B2" }) |bad| {
        const argv = [_][]const u8{ "cells", "f.xlsx", "--range", bad };
        try std.testing.expectError(ArgError.BadArgValue, parseArgs(&argv));
    }
    // Inverted corners are rejected (no silent normalisation).
    {
        const argv = [_][]const u8{ "cells", "f.xlsx", "--range", "Z1:A1" };
        try std.testing.expectError(ArgError.BadArgValue, parseArgs(&argv));
    }
    {
        const argv = [_][]const u8{ "cells", "f.xlsx", "--range", "A10:A1" };
        try std.testing.expectError(ArgError.BadArgValue, parseArgs(&argv));
    }
    // Missing value.
    {
        const argv = [_][]const u8{ "cells", "f.xlsx", "--range" };
        try std.testing.expectError(ArgError.MissingValue, parseArgs(&argv));
    }
    // Sub-commands without row+col keys reject --range.
    inline for (.{ "comments", "validations", "hyperlinks", "merges", "defined-names", "doc-props", "meta", "list-sheets", "styles", "sst" }) |cmd| {
        const argv = [_][]const u8{ cmd, "f.xlsx", "--range", "A1:B2" };
        try std.testing.expectError(ArgError.BadArgValue, parseArgs(&argv));
    }
    // Legacy --list-sheets flag also rejects.
    {
        const argv = [_][]const u8{ "f.xlsx", "--list-sheets", "--range", "A1:B2" };
        try std.testing.expectError(ArgError.BadArgValue, parseArgs(&argv));
    }
    // Single-cell "A1" (no colon) is rejected — contract is a rectangle.
    {
        const argv = [_][]const u8{ "cells", "f.xlsx", "--range", "A1" };
        try std.testing.expectError(ArgError.BadArgValue, parseArgs(&argv));
    }
    // Defaults to null when absent.
    {
        const argv = [_][]const u8{ "cells", "f.xlsx" };
        const a = try parseArgs(&argv);
        try std.testing.expect(a.range == null);
    }
    // detectSubcommand skips the paired value: an A1 ref that looks
    // like a positional (e.g. `A1:B2` begins with a letter) must not
    // be mistaken for the file path.
    {
        const argv = [_][]const u8{ "rows", "--range", "A1:B2", "f.xlsx" };
        const a = try parseArgs(&argv);
        try std.testing.expectEqualStrings("f.xlsx", a.file);
        try std.testing.expect(a.range != null);
    }
}

test "runCellsCommand --range filters by bounding rectangle" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, io, "cli_range_iter59b2.xlsx");
    defer std.testing.allocator.free(tmp_path);
    {
        const writer = xlsx.writer_types;
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s0 = try w.addSheet("Data");
        // 5×5 grid; ref-style values so we can assert exact inclusion.
        // Row 1: A1..E1, row 2: A2..E2, …
        try s0.writeRow(&.{ .{ .string = "A1" }, .{ .string = "B1" }, .{ .string = "C1" }, .{ .string = "D1" }, .{ .string = "E1" } });
        try s0.writeRow(&.{ .{ .string = "A2" }, .{ .string = "B2" }, .{ .string = "C2" }, .{ .string = "D2" }, .{ .string = "E2" } });
        try s0.writeRow(&.{ .{ .string = "A3" }, .{ .string = "B3" }, .{ .string = "C3" }, .{ .string = "D3" }, .{ .string = "E3" } });
        try s0.writeRow(&.{ .{ .string = "A4" }, .{ .string = "B4" }, .{ .string = "C4" }, .{ .string = "D4" }, .{ .string = "E4" } });
        try s0.writeRow(&.{ .{ .string = "A5" }, .{ .string = "B5" }, .{ .string = "C5" }, .{ .string = "D5" }, .{ .string = "E5" } });
        try w.save(io, tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, io, tmp_path);
    defer book.deinit();

    const countLines = struct {
        fn f(s: []const u8) usize {
            var n: usize = 0;
            for (s) |c| if (c == '\n') {
                n += 1;
            };
            return n;
        }
    }.f;

    // --range B2:C3 → exactly 4 cells: B2, C2, B3, C3.
    {
        var scratch: [8192]u8 = undefined;
        var w = std.Io.Writer.fixed(&scratch);
        const range: xlsx.MergeRange = .{
            .top_left = .{ .col = 1, .row = 2 },
            .bottom_right = .{ .col = 2, .row = 3 },
        };
        try runCellsCommand(&w, &book, book.sheets[0], 0, std.testing.allocator, null, null, null, null, range, false, false);
        const out = w.buffered();
        try std.testing.expectEqual(@as(usize, 4), countLines(out));
        try std.testing.expect(std.mem.indexOf(u8, out, "\"v\":\"B2\"") != null);
        try std.testing.expect(std.mem.indexOf(u8, out, "\"v\":\"C2\"") != null);
        try std.testing.expect(std.mem.indexOf(u8, out, "\"v\":\"B3\"") != null);
        try std.testing.expect(std.mem.indexOf(u8, out, "\"v\":\"C3\"") != null);
        // Corner spot-checks outside the rect.
        try std.testing.expect(std.mem.indexOf(u8, out, "\"v\":\"A1\"") == null);
        try std.testing.expect(std.mem.indexOf(u8, out, "\"v\":\"A2\"") == null);
        try std.testing.expect(std.mem.indexOf(u8, out, "\"v\":\"D2\"") == null);
        try std.testing.expect(std.mem.indexOf(u8, out, "\"v\":\"B4\"") == null);
        try std.testing.expect(std.mem.indexOf(u8, out, "\"v\":\"E5\"") == null);
    }

    // Intersection with --start-row / --end-row: --range B2:C4 ∩
    // [start=3, end=5] → rows {3, 4}, cols {1, 2} → 4 cells.
    {
        var scratch: [8192]u8 = undefined;
        var w = std.Io.Writer.fixed(&scratch);
        const range: xlsx.MergeRange = .{
            .top_left = .{ .col = 1, .row = 2 },
            .bottom_right = .{ .col = 2, .row = 4 },
        };
        try runCellsCommand(&w, &book, book.sheets[0], 0, std.testing.allocator, null, null, 3, 5, range, false, false);
        const out = w.buffered();
        try std.testing.expectEqual(@as(usize, 4), countLines(out));
        try std.testing.expect(std.mem.indexOf(u8, out, "\"v\":\"B3\"") != null);
        try std.testing.expect(std.mem.indexOf(u8, out, "\"v\":\"C3\"") != null);
        try std.testing.expect(std.mem.indexOf(u8, out, "\"v\":\"B4\"") != null);
        try std.testing.expect(std.mem.indexOf(u8, out, "\"v\":\"C4\"") != null);
        try std.testing.expect(std.mem.indexOf(u8, out, "\"v\":\"B2\"") == null);
        try std.testing.expect(std.mem.indexOf(u8, out, "\"v\":\"B5\"") == null);
    }
}

test "runRowsCommand --range filters rows + masks out-of-range cells" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, io, "cli_range_rows_iter59b2.xlsx");
    defer std.testing.allocator.free(tmp_path);
    {
        const writer = xlsx.writer_types;
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s0 = try w.addSheet("Data");
        try s0.writeRow(&.{ .{ .string = "A1" }, .{ .string = "B1" }, .{ .string = "C1" }, .{ .string = "D1" }, .{ .string = "E1" } });
        try s0.writeRow(&.{ .{ .string = "A2" }, .{ .string = "B2" }, .{ .string = "C2" }, .{ .string = "D2" }, .{ .string = "E2" } });
        try s0.writeRow(&.{ .{ .string = "A3" }, .{ .string = "B3" }, .{ .string = "C3" }, .{ .string = "D3" }, .{ .string = "E3" } });
        try s0.writeRow(&.{ .{ .string = "A4" }, .{ .string = "B4" }, .{ .string = "C4" }, .{ .string = "D4" }, .{ .string = "E4" } });
        try s0.writeRow(&.{ .{ .string = "A5" }, .{ .string = "B5" }, .{ .string = "C5" }, .{ .string = "D5" }, .{ .string = "E5" } });
        try w.save(io, tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, io, tmp_path);
    defer book.deinit();

    const countLines = struct {
        fn f(s: []const u8) usize {
            var n: usize = 0;
            for (s) |c| if (c == '\n') {
                n += 1;
            };
            return n;
        }
    }.f;

    // --range B2:C3 on rows → 2 envelope lines (rows 2 and 3).
    // Out-of-range columns are masked to empty, so only B2/C2 and B3/C3
    // appear as quoted values.
    var scratch: [8192]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    const range: xlsx.MergeRange = .{
        .top_left = .{ .col = 1, .row = 2 },
        .bottom_right = .{ .col = 2, .row = 3 },
    };
    try runRowsCommand(&w, &book, book.sheets[0], 0, .jsonl, std.testing.allocator, null, null, null, null, range, false, false, false);
    const out = w.buffered();
    try std.testing.expectEqual(@as(usize, 2), countLines(out));
    try std.testing.expect(std.mem.indexOf(u8, out, "\"B2\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"C2\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"B3\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"C3\"") != null);
    // Row 1, 4, 5 entirely absent.
    try std.testing.expect(std.mem.indexOf(u8, out, "\"A1\"") == null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"B4\"") == null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"E5\"") == null);
    // Out-of-col cells in kept rows are masked, so A2/D2/E2/A3/D3/E3
    // must NOT appear as quoted values in the envelope.
    try std.testing.expect(std.mem.indexOf(u8, out, "\"A2\"") == null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"D2\"") == null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"E2\"") == null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"A3\"") == null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"D3\"") == null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"E3\"") == null);
}

test "parseArgs --header scoping" {
    // Happy: `rows` + default format.
    {
        const argv = [_][]const u8{ "rows", "f.xlsx", "--header" };
        const a = try parseArgs(&argv);
        try std.testing.expect(a.header);
        try std.testing.expectEqual(Subcommand.rows, a.subcommand);
        try std.testing.expectEqual(Format.jsonl, a.format);
    }
    // Happy: `rows` + explicit `--format jsonl`.
    {
        const argv = [_][]const u8{ "rows", "f.xlsx", "--header", "--format", "jsonl" };
        const a = try parseArgs(&argv);
        try std.testing.expect(a.header);
    }
    // Reject: --header on any non-jsonl format (tsv/csv/legacy variants
    // have their own row shapes and --header would silently no-op).
    inline for (.{ "tsv", "csv", "legacy-jsonl", "legacy-jsonl-dict" }) |fmt| {
        const argv = [_][]const u8{ "rows", "f.xlsx", "--header", "--format", fmt };
        try std.testing.expectError(ArgError.BadArgValue, parseArgs(&argv));
    }
    // Reject: --header on any other sub-command.
    inline for (.{ "cells", "comments", "validations", "hyperlinks", "meta", "list-sheets", "styles", "sst" }) |cmd| {
        const argv = [_][]const u8{ cmd, "f.xlsx", "--header" };
        try std.testing.expectError(ArgError.BadArgValue, parseArgs(&argv));
    }
    // Reject: --header with the legacy plain-text --list-sheets flag.
    {
        const argv = [_][]const u8{ "f.xlsx", "--list-sheets", "--header" };
        try std.testing.expectError(ArgError.BadArgValue, parseArgs(&argv));
    }
    // Default: --header off when absent.
    {
        const argv = [_][]const u8{ "rows", "f.xlsx" };
        const a = try parseArgs(&argv);
        try std.testing.expect(!a.header);
    }
}

test "runRowsCommand --header promotes first row and emits fields dict" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, io, "cli_header_iter59b3.xlsx");
    defer std.testing.allocator.free(tmp_path);
    {
        const writer = xlsx.writer_types;
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s0 = try w.addSheet("Data");
        try s0.writeRow(&.{ .{ .string = "name" }, .{ .string = "qty" } });
        try s0.writeRow(&.{ .{ .string = "apple" }, .{ .integer = 3 } });
        try s0.writeRow(&.{ .{ .string = "pear" }, .{ .integer = 7 } });
        try w.save(io, tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, io, tmp_path);
    defer book.deinit();

    const countLines = struct {
        fn f(s: []const u8) usize {
            var n: usize = 0;
            for (s) |c| if (c == '\n') {
                n += 1;
            };
            return n;
        }
    }.f;

    var scratch: [8192]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    try runRowsCommand(&w, &book, book.sheets[0], 0, .jsonl, std.testing.allocator, null, null, null, null, null, true, false, false);
    const out = w.buffered();

    // 3 rows in, header consumed → exactly 2 records out.
    try std.testing.expectEqual(@as(usize, 2), countLines(out));
    // Header row must NOT appear as a record.
    try std.testing.expect(std.mem.indexOf(u8, out, "\"v\":\"name\"") == null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"cells\":") == null);
    // Data rows emit as fields dicts keyed by header cell values.
    try std.testing.expect(std.mem.indexOf(u8, out, "\"fields\":{\"name\":\"apple\",\"qty\":3}") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"fields\":{\"name\":\"pear\",\"qty\":7}") != null);
    // Envelope scaffolding still present.
    try std.testing.expect(std.mem.indexOf(u8, out, "\"kind\":\"row\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"sheet\":\"Data\"") != null);
}

test "runRowsCommand --header duplicate header keys emitted verbatim" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, io, "cli_header_dup_iter59b3.xlsx");
    defer std.testing.allocator.free(tmp_path);
    {
        const writer = xlsx.writer_types;
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s0 = try w.addSheet("Data");
        try s0.writeRow(&.{ .{ .string = "x" }, .{ .string = "x" } });
        try s0.writeRow(&.{ .{ .integer = 1 }, .{ .integer = 2 } });
        try w.save(io, tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, io, tmp_path);
    defer book.deinit();

    var scratch: [1024]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    try runRowsCommand(&w, &book, book.sheets[0], 0, .jsonl, std.testing.allocator, null, null, null, null, null, true, false, false);
    const out = w.buffered();

    // Both duplicate "x" keys appear in the dict as-is.
    try std.testing.expect(std.mem.indexOf(u8, out, "\"fields\":{\"x\":1,\"x\":2}") != null);
}

test "runRowsCommand --header empty header cells fall back to col_<letter>" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, io, "cli_header_empty_iter59b3.xlsx");
    defer std.testing.allocator.free(tmp_path);
    {
        const writer = xlsx.writer_types;
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s0 = try w.addSheet("Data");
        // Header: A="name", B=empty, C="qty" → keys "name","col_B","qty".
        try s0.writeRow(&.{ .{ .string = "name" }, .empty, .{ .string = "qty" } });
        try s0.writeRow(&.{ .{ .string = "apple" }, .{ .integer = 42 }, .{ .integer = 3 } });
        try w.save(io, tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, io, tmp_path);
    defer book.deinit();

    var scratch: [1024]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    try runRowsCommand(&w, &book, book.sheets[0], 0, .jsonl, std.testing.allocator, null, null, null, null, null, true, false, false);
    const out = w.buffered();

    try std.testing.expect(std.mem.indexOf(u8, out, "\"col_B\":42") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"name\":\"apple\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"qty\":3") != null);
}

test "runRowsCommand --header + --range derives keys only from in-range cols" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    // Header row has 4 cells across A..D ("w","x","y","z"); a --range
    // B:C must consume only the B/C header cells and emit data dicts
    // keyed exactly {"x","y"} — no `col_A` / `col_D` leak from the
    // masked full-width view.
    const tmp_path = try tt.path(std.testing.allocator, io, "cli_header_range_iter59b3.xlsx");
    defer std.testing.allocator.free(tmp_path);
    {
        const writer = xlsx.writer_types;
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s0 = try w.addSheet("Data");
        try s0.writeRow(&.{ .{ .string = "w" }, .{ .string = "x" }, .{ .string = "y" }, .{ .string = "z" } });
        try s0.writeRow(&.{ .{ .integer = 1 }, .{ .integer = 2 }, .{ .integer = 3 }, .{ .integer = 4 } });
        try w.save(io, tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, io, tmp_path);
    defer book.deinit();

    var scratch: [1024]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);

    // B:C → cols 1..2 (0-based).
    const range: xlsx.MergeRange = .{
        .top_left = .{ .row = 1, .col = 1 },
        .bottom_right = .{ .row = 2, .col = 2 },
    };
    try runRowsCommand(&w, &book, book.sheets[0], 0, .jsonl, std.testing.allocator, null, null, null, null, range, true, false, false);
    const out = w.buffered();

    // Only the in-range header keys should appear.
    try std.testing.expect(std.mem.indexOf(u8, out, "\"x\":2") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"y\":3") != null);
    // Out-of-range headers and their `col_<letter>` fallbacks must NOT.
    try std.testing.expect(std.mem.indexOf(u8, out, "\"w\"") == null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"z\"") == null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"col_A\"") == null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"col_D\"") == null);
}

test "runRowsCommand --include-blanks on csv/header is a no-op for blank rows" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    // Tight contract per iter59b-4 P2 follow-up: --include-blanks
    // preserves all-blank rows ONLY on the envelope (.jsonl) path.
    // On csv / tsv / legacy-jsonl / legacy-jsonl-dict the flag is a
    // documented no-op and must NOT inject extra blank output lines.
    // On --header the flag is also a no-op — a blank row must not
    // promote to a `col_*`-keyed header.
    const tmp_path = try tt.path(std.testing.allocator, io, "cli_blanks_flat_iter59b4.xlsx");
    defer std.testing.allocator.free(tmp_path);
    {
        const writer = xlsx.writer_types;
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s0 = try w.addSheet("Data");
        // Row 1: A only; B/C blank. Row 2: C only; A/B blank.
        try s0.writeRow(&.{ .{ .string = "x" }, .empty, .empty });
        try s0.writeRow(&.{ .empty, .empty, .{ .string = "y" } });
        try w.save(io, tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, io, tmp_path);
    defer book.deinit();

    // csv + --range B:B + --include-blanks — range is all-empty
    // for both rows. Must emit nothing.
    var scratch: [1024]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    const range_b_only: xlsx.MergeRange = .{
        .top_left = .{ .row = 1, .col = 1 },
        .bottom_right = .{ .row = 2, .col = 1 },
    };
    try runRowsCommand(&w, &book, book.sheets[0], 0, .csv, std.testing.allocator, null, null, null, null, range_b_only, false, true, false);
    try std.testing.expectEqual(@as(usize, 0), w.buffered().len);
}

test "runRowsCommand --range + --include-blanks keeps blank-only ranged rows" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    // A row with data only in A/D (both outside the B:C range) and
    // --include-blanks must still emit with two t:"blank" cells —
    // the whole point of --include-blanks is to surface empties.
    const tmp_path = try tt.path(std.testing.allocator, io, "cli_range_blank_iter59b4.xlsx");
    defer std.testing.allocator.free(tmp_path);
    {
        const writer = xlsx.writer_types;
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s0 = try w.addSheet("Data");
        // Row 1: "x" in A, "y" in D. Nothing in B/C.
        try s0.writeRow(&.{ .{ .string = "x" }, .empty, .empty, .{ .string = "y" } });
        try w.save(io, tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, io, tmp_path);
    defer book.deinit();

    var scratch: [1024]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);

    // B:C → cols 1..2 (0-based).
    const range: xlsx.MergeRange = .{
        .top_left = .{ .row = 1, .col = 1 },
        .bottom_right = .{ .row = 1, .col = 2 },
    };
    try runRowsCommand(&w, &book, book.sheets[0], 0, .jsonl, std.testing.allocator, null, null, null, null, range, false, true, false);
    const out = w.buffered();

    // The row must appear with two t:"blank" cells.
    try std.testing.expect(std.mem.indexOf(u8, out, "\"kind\":\"row\"") != null);
    // Count t:"blank" occurrences — should be 2 (B and C).
    var count: usize = 0;
    var i: usize = 0;
    while (std.mem.indexOfPos(u8, out, i, "\"t\":\"blank\"")) |pos| : (i = pos + 1) count += 1;
    try std.testing.expectEqual(@as(usize, 2), count);
}

test "writeTerseStyleBlock doesn't leak empty border for diagonal-only sides" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    // Codex P2: a cell whose border has ONLY the diagonal side set
    // must not serialize `"border":{}` — the terse emitter omits
    // diagonal entirely, so emitting an empty border object would
    // be a shape leak. A style block may still appear because the
    // Zig writer attaches a default font to every styled cell
    // (which may have a color), but the "border" key must never
    // appear with an empty object.
    const tmp_path = try tt.path(std.testing.allocator, io, "cli_diag_only_iter59b4.xlsx");
    defer std.testing.allocator.free(tmp_path);
    {
        const writer = xlsx.writer_types;
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        const diag_style = try w.addStyle(.{
            .border_diagonal = .{ .style = .thin, .color_argb = 0xFF000000 },
        });
        var s0 = try w.addSheet("Data");
        try s0.writeRowStyled(&.{.{ .string = "x" }}, &.{diag_style});
        try w.save(io, tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, io, tmp_path);
    defer book.deinit();

    var scratch: [1024]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    try runCellsCommand(&w, &book, book.sheets[0], 0, std.testing.allocator, null, null, null, null, null, false, true);
    const out = w.buffered();

    // Cell must appear.
    try std.testing.expect(std.mem.indexOf(u8, out, "\"v\":\"x\"") != null);
    // `"border"` must not appear AT ALL — no `"border":{…}` for the
    // diagonal-only case because the terse block excludes diagonal.
    try std.testing.expect(std.mem.indexOf(u8, out, "\"border\"") == null);
}

test "parseArgs --include-blanks scoping" {
    // Happy: cells + default.
    {
        const argv = [_][]const u8{ "cells", "f.xlsx", "--include-blanks" };
        const a = try parseArgs(&argv);
        try std.testing.expect(a.include_blanks);
        try std.testing.expectEqual(Subcommand.cells, a.subcommand);
    }
    // Happy: rows + default envelope.
    {
        const argv = [_][]const u8{ "rows", "f.xlsx", "--include-blanks" };
        const a = try parseArgs(&argv);
        try std.testing.expect(a.include_blanks);
    }
    // Happy (no-op but accepted): rows + --header.
    {
        const argv = [_][]const u8{ "rows", "f.xlsx", "--include-blanks", "--header" };
        const a = try parseArgs(&argv);
        try std.testing.expect(a.include_blanks);
        try std.testing.expect(a.header);
    }
    // Happy (no-op but accepted): rows + flat formats.
    inline for (.{ "tsv", "csv", "legacy-jsonl", "legacy-jsonl-dict" }) |fmt| {
        const argv = [_][]const u8{ "rows", "f.xlsx", "--include-blanks", "--format", fmt };
        const a = try parseArgs(&argv);
        try std.testing.expect(a.include_blanks);
    }
    // Reject: every non-cells/rows sub-command.
    inline for (.{ "comments", "validations", "hyperlinks", "meta", "list-sheets", "styles", "sst" }) |cmd| {
        const argv = [_][]const u8{ cmd, "f.xlsx", "--include-blanks" };
        try std.testing.expectError(ArgError.BadArgValue, parseArgs(&argv));
    }
    // Reject: legacy --list-sheets flag.
    {
        const argv = [_][]const u8{ "f.xlsx", "--list-sheets", "--include-blanks" };
        try std.testing.expectError(ArgError.BadArgValue, parseArgs(&argv));
    }
    // Default off.
    {
        const argv = [_][]const u8{ "cells", "f.xlsx" };
        const a = try parseArgs(&argv);
        try std.testing.expect(!a.include_blanks);
    }
}

test "parseArgs --with-styles scoping" {
    // Happy: cells + default.
    {
        const argv = [_][]const u8{ "cells", "f.xlsx", "--with-styles" };
        const a = try parseArgs(&argv);
        try std.testing.expect(a.with_styles);
    }
    // Happy: rows + jsonl (no --header).
    {
        const argv = [_][]const u8{ "rows", "f.xlsx", "--with-styles" };
        const a = try parseArgs(&argv);
        try std.testing.expect(a.with_styles);
    }
    // Reject: rows + flat formats (no place for nested metadata).
    inline for (.{ "tsv", "csv", "legacy-jsonl", "legacy-jsonl-dict" }) |fmt| {
        const argv = [_][]const u8{ "rows", "f.xlsx", "--with-styles", "--format", fmt };
        try std.testing.expectError(ArgError.BadArgValue, parseArgs(&argv));
    }
    // Reject: rows + --header (fields dict has no per-cell slot).
    {
        const argv = [_][]const u8{ "rows", "f.xlsx", "--with-styles", "--header" };
        try std.testing.expectError(ArgError.BadArgValue, parseArgs(&argv));
    }
    // Reject: every non-cells/rows sub-command.
    inline for (.{ "comments", "validations", "hyperlinks", "meta", "list-sheets", "styles", "sst" }) |cmd| {
        const argv = [_][]const u8{ cmd, "f.xlsx", "--with-styles" };
        try std.testing.expectError(ArgError.BadArgValue, parseArgs(&argv));
    }
    // Reject: legacy --list-sheets flag.
    {
        const argv = [_][]const u8{ "f.xlsx", "--list-sheets", "--with-styles" };
        try std.testing.expectError(ArgError.BadArgValue, parseArgs(&argv));
    }
    // Default off.
    {
        const argv = [_][]const u8{ "cells", "f.xlsx" };
        const a = try parseArgs(&argv);
        try std.testing.expect(!a.with_styles);
    }
}

test "runCellsCommand --include-blanks emits t:\"blank\" for empty cells" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, io, "cli_blanks_iter59b4.xlsx");
    defer std.testing.allocator.free(tmp_path);
    {
        const writer = xlsx.writer_types;
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s0 = try w.addSheet("Data");
        // Row 1: A="x", B=empty, C=7 → a single sparse row with a gap.
        try s0.writeRow(&.{ .{ .string = "x" }, .empty, .{ .integer = 7 } });
        try w.save(io, tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, io, tmp_path);
    defer book.deinit();

    var scratch: [2048]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    try runCellsCommand(&w, &book, book.sheets[0], 0, std.testing.allocator, null, null, null, null, null, true, false);
    const out = w.buffered();

    // Non-empty cells still emit with their proper types.
    try std.testing.expect(std.mem.indexOf(u8, out, "\"ref\":\"A1\",\"row\":1,\"col\":1,\"t\":\"str\",\"v\":\"x\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"ref\":\"C1\",\"row\":1,\"col\":3,\"t\":\"int\",\"v\":7") != null);
    // The gap at B1 must surface as a blank record.
    try std.testing.expect(std.mem.indexOf(u8, out, "\"ref\":\"B1\",\"row\":1,\"col\":2,\"t\":\"blank\",\"v\":null") != null);
}

test "runCellsCommand without --include-blanks still skips empties" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    // Regression guard: default behaviour preserved when the flag is off.
    const tmp_path = try tt.path(std.testing.allocator, io, "cli_blanks_off_iter59b4.xlsx");
    defer std.testing.allocator.free(tmp_path);
    {
        const writer = xlsx.writer_types;
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s0 = try w.addSheet("Data");
        try s0.writeRow(&.{ .{ .string = "x" }, .empty, .{ .integer = 7 } });
        try w.save(io, tmp_path);
    }
    var book = try xlsx.Book.open(std.testing.allocator, io, tmp_path);
    defer book.deinit();

    var scratch: [2048]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    try runCellsCommand(&w, &book, book.sheets[0], 0, std.testing.allocator, null, null, null, null, null, false, false);
    const out = w.buffered();

    try std.testing.expect(std.mem.indexOf(u8, out, "\"blank\"") == null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"B1\"") == null);
}

test "runCellsCommand --with-styles emits terse style block for styled cells" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, io, "cli_with_styles_iter59b4.xlsx");
    defer std.testing.allocator.free(tmp_path);
    {
        const writer = xlsx.writer_types;
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        // Bold + white font on dark-blue solid fill — the canonical
        // header-row look from the design-doc example.
        const styled = try w.addStyle(.{
            .font_bold = true,
            .font_color_argb = 0xFFFFFFFF,
            .fill_pattern = .solid,
            .fill_fg_argb = 0xFF1F4E79,
        });
        var s0 = try w.addSheet("Data");
        try s0.writeRowStyled(
            &.{ .{ .string = "name" }, .{ .string = "qty" } },
            &.{ styled, styled },
        );
        try s0.writeRow(&.{ .{ .string = "apple" }, .{ .integer = 3 } });
        try w.save(io, tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, io, tmp_path);
    defer book.deinit();

    var scratch: [4096]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    try runCellsCommand(&w, &book, book.sheets[0], 0, std.testing.allocator, null, null, null, null, null, false, true);
    const out = w.buffered();

    // Styled header cells surface the terse block.
    try std.testing.expect(std.mem.indexOf(u8, out, "\"style\":{") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"bold\":true") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"fg\":\"FFFFFFFF\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"bg\":\"FF1F4E79\"") != null);
    // Unstyled data cells ("apple", 3) must NOT carry a style field.
    const apple_pos = std.mem.indexOf(u8, out, "\"v\":\"apple\"").?;
    const apple_line_end = std.mem.indexOfScalarPos(u8, out, apple_pos, '\n').?;
    const apple_line = out[apple_pos..apple_line_end];
    try std.testing.expect(std.mem.indexOf(u8, apple_line, "\"style\"") == null);
}

test "runRowsCommand --with-styles on envelope attaches style to per-cell records" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, io, "cli_rows_styles_iter59b4.xlsx");
    defer std.testing.allocator.free(tmp_path);
    {
        const writer = xlsx.writer_types;
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        const italic = try w.addStyle(.{ .font_italic = true });
        var s0 = try w.addSheet("Data");
        try s0.writeRowStyled(
            &.{ .{ .string = "a" }, .{ .string = "b" } },
            &.{ italic, 0 },
        );
        try w.save(io, tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, io, tmp_path);
    defer book.deinit();

    var scratch: [2048]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    try runRowsCommand(&w, &book, book.sheets[0], 0, .jsonl, std.testing.allocator, null, null, null, null, null, false, false, true);
    const out = w.buffered();

    // Styled A1 gets the terse block with just italic:true.
    try std.testing.expect(std.mem.indexOf(u8, out, "\"ref\":\"A1\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"italic\":true") != null);
    // Default-styled B1 has NO style field.
    const b1_pos = std.mem.indexOf(u8, out, "\"ref\":\"B1\"").?;
    // Scan from B1's start to end-of-object (next `}`). Cheap since
    // each cell record is < 200 bytes in this tiny fixture.
    const rel_close = std.mem.indexOfScalarPos(u8, out, b1_pos, '}').?;
    const b1_record = out[b1_pos..rel_close];
    try std.testing.expect(std.mem.indexOf(u8, b1_record, "\"style\"") == null);
}

test "runCellsCommand --skip / --take slice the emitted cell stream" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, io, "cli_pagination_iter59a.xlsx");
    defer std.testing.allocator.free(tmp_path);
    {
        const writer = xlsx.writer_types;
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s0 = try w.addSheet("Data");
        // 5 rows × 1 cell each → 5 candidate cells in emit order.
        try s0.writeRow(&.{.{ .string = "c1" }});
        try s0.writeRow(&.{.{ .string = "c2" }});
        try s0.writeRow(&.{.{ .string = "c3" }});
        try s0.writeRow(&.{.{ .string = "c4" }});
        try s0.writeRow(&.{.{ .string = "c5" }});
        try w.save(io, tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, io, tmp_path);
    defer book.deinit();

    const countLines = struct {
        fn f(s: []const u8) usize {
            var n: usize = 0;
            for (s) |c| if (c == '\n') {
                n += 1;
            };
            return n;
        }
    }.f;

    // Baseline — no pagination.
    {
        var scratch: [4096]u8 = undefined;
        var w = std.Io.Writer.fixed(&scratch);
        try runCellsCommand(&w, &book, book.sheets[0], 0, std.testing.allocator, null, null, null, null, null, false, false);
        try std.testing.expectEqual(@as(usize, 5), countLines(w.buffered()));
    }
    // --skip 2 drops the first two cells (c1, c2).
    {
        var scratch: [4096]u8 = undefined;
        var w = std.Io.Writer.fixed(&scratch);
        try runCellsCommand(&w, &book, book.sheets[0], 0, std.testing.allocator, 2, null, null, null, null, false, false);
        const out = w.buffered();
        try std.testing.expectEqual(@as(usize, 3), countLines(out));
        try std.testing.expect(std.mem.indexOf(u8, out, "\"v\":\"c1\"") == null);
        try std.testing.expect(std.mem.indexOf(u8, out, "\"v\":\"c2\"") == null);
        try std.testing.expect(std.mem.indexOf(u8, out, "\"v\":\"c3\"") != null);
    }
    // --take 3 keeps exactly the first three.
    {
        var scratch: [4096]u8 = undefined;
        var w = std.Io.Writer.fixed(&scratch);
        try runCellsCommand(&w, &book, book.sheets[0], 0, std.testing.allocator, null, 3, null, null, null, false, false);
        const out = w.buffered();
        try std.testing.expectEqual(@as(usize, 3), countLines(out));
        try std.testing.expect(std.mem.indexOf(u8, out, "\"v\":\"c3\"") != null);
        try std.testing.expect(std.mem.indexOf(u8, out, "\"v\":\"c4\"") == null);
    }
    // --skip 1 --take 2 yields the exact middle slice: c2, c3.
    {
        var scratch: [4096]u8 = undefined;
        var w = std.Io.Writer.fixed(&scratch);
        try runCellsCommand(&w, &book, book.sheets[0], 0, std.testing.allocator, 1, 2, null, null, null, false, false);
        const out = w.buffered();
        try std.testing.expectEqual(@as(usize, 2), countLines(out));
        try std.testing.expect(std.mem.indexOf(u8, out, "\"v\":\"c1\"") == null);
        try std.testing.expect(std.mem.indexOf(u8, out, "\"v\":\"c2\"") != null);
        try std.testing.expect(std.mem.indexOf(u8, out, "\"v\":\"c3\"") != null);
        try std.testing.expect(std.mem.indexOf(u8, out, "\"v\":\"c4\"") == null);
    }
}

test "runMetaCommand emits path:null on non-UTF-8 workbook path" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var scratch: [512]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);

    // Build a minimal Book-shaped view without actually opening a
    // file — runMetaCommand only dereferences book.sheets / sst /
    // styles_xml / theme_xml / rich_runs_by_sst_idx / comments.
    var empty_book: xlsx.Book = .{
        .io = io,
        .allocator = std.testing.allocator,
        .sst_arena = std.heap.ArenaAllocator.init(std.testing.allocator),
    };
    defer empty_book.deinit();

    try runMetaCommand(&w, &empty_book, null, .ndjson, null);

    const out = scratch[0..w.end];
    // The path field must serialize as literal `null`, not a string.
    try std.testing.expect(std.mem.indexOf(u8, out, "\"path\":null") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"kind\":\"workbook\"") != null);
}

test "runListSheetsCommand emits one sheet record per sheet" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, io, "cli_list_sheets_iter57.xlsx");
    defer std.testing.allocator.free(tmp_path);
    {
        const writer = xlsx.writer_types;
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s0 = try w.addSheet("Data");
        try s0.writeRow(&.{.{ .string = "hdr" }});
        var s1 = try w.addSheet("Other");
        try s1.writeRow(&.{.{ .integer = 1 }});
        var s2 = try w.addSheet("She\"et"); // name with a quote — must be JSON-escaped
        try s2.writeRow(&.{.{ .boolean = true }});
        try w.save(io, tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, io, tmp_path);
    defer book.deinit();

    var scratch: [1024]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    try runListSheetsCommand(&w, &book, .ndjson);
    // Z4: every record now carries `state`. Writer-produced sheets are
    // all visible; the veryHidden path is covered by the sheet-state
    // test below.
    try std.testing.expectEqualStrings(
        "{\"kind\":\"sheet\",\"sheet\":\"Data\",\"sheet_idx\":0,\"state\":\"visible\"}\n" ++
            "{\"kind\":\"sheet\",\"sheet\":\"Other\",\"sheet_idx\":1,\"state\":\"visible\"}\n" ++
            "{\"kind\":\"sheet\",\"sheet\":\"She\\\"et\",\"sheet_idx\":2,\"state\":\"visible\"}\n",
        w.buffered(),
    );
}

test "runMetaCommand emits workbook record with sst/has_* fields then sheet records" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, io, "cli_meta_iter57.xlsx");
    defer std.testing.allocator.free(tmp_path);
    {
        const writer = xlsx.writer_types;
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s0 = try w.addSheet("Data");
        // Two distinct strings + one repeat → SST count of 2,
        // exercises the deduped path.
        try s0.writeRow(&.{ .{ .string = "alpha" }, .{ .string = "beta" } });
        try s0.writeRow(&.{.{ .string = "alpha" }});
        try s0.addComment("A1", "me", "hi there"); // forces has_comments=true for this sheet
        var s1 = try w.addSheet("NoComments");
        try s1.writeRow(&.{.{ .integer = 42 }});
        try w.save(io, tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, io, tmp_path);
    defer book.deinit();

    var scratch: [4096]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    try runMetaCommand(&w, &book, tmp_path, .ndjson, null);

    const out = w.buffered();
    // Parse NDJSON line by line and assert field presence + values.
    var line_it = std.mem.splitScalar(u8, out, '\n');

    const wb_line = line_it.next() orelse return error.MissingWorkbookLine;
    // Structural probes — avoid order-sensitive equality because the
    // exact field ordering is an implementation detail the wire format
    // only loosely pins down. We pin the presence + values.
    try std.testing.expect(std.mem.indexOf(u8, wb_line, "\"kind\":\"workbook\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, wb_line, "\"sheets\":2") != null);
    try std.testing.expect(std.mem.indexOf(u8, wb_line, "\"sst\":{\"count\":2,\"rich\":0}") != null);
    // has_styles / has_theme reflect whether the writer chose to emit
    // those parts — we only pin field *presence* here, not the writer's
    // part-emission policy. The workbook-scoped `has_comments` is
    // deterministic given the addComment call above.
    try std.testing.expect(
        std.mem.indexOf(u8, wb_line, "\"has_styles\":true") != null or
            std.mem.indexOf(u8, wb_line, "\"has_styles\":false") != null,
    );
    try std.testing.expect(
        std.mem.indexOf(u8, wb_line, "\"has_theme\":true") != null or
            std.mem.indexOf(u8, wb_line, "\"has_theme\":false") != null,
    );
    try std.testing.expect(std.mem.indexOf(u8, wb_line, "\"has_comments\":true") != null);
    try std.testing.expect(std.mem.indexOf(u8, wb_line, "\"path\":") != null);

    const sheet0 = line_it.next() orelse return error.MissingSheet0;
    try std.testing.expect(std.mem.indexOf(u8, sheet0, "\"kind\":\"sheet\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, sheet0, "\"sheet\":\"Data\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, sheet0, "\"sheet_idx\":0") != null);
    try std.testing.expect(std.mem.indexOf(u8, sheet0, "\"has_comments\":true") != null);

    const sheet1 = line_it.next() orelse return error.MissingSheet1;
    try std.testing.expect(std.mem.indexOf(u8, sheet1, "\"sheet\":\"NoComments\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, sheet1, "\"sheet_idx\":1") != null);
    try std.testing.expect(std.mem.indexOf(u8, sheet1, "\"has_comments\":false") != null);

    // Trailing empty token after the final '\n' — but no more records.
    const trailing = line_it.next();
    if (trailing) |t| try std.testing.expectEqualStrings("", t);
    try std.testing.expectEqual(@as(?[]const u8, null), line_it.next());
}

test "legacy --list-sheets flag still emits plain text (regression guard)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    // Regression guard: the legacy plain-text shape is exactly
    // `<name>\n` per sheet, no JSON, no sub-command routing. This
    // mirrors the code path in main() line-for-line so the flag
    // keeps working across iter57.
    const tmp_path = try tt.path(std.testing.allocator, io, "cli_legacy_list_sheets.xlsx");
    defer std.testing.allocator.free(tmp_path);
    {
        const writer = xlsx.writer_types;
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s0 = try w.addSheet("Data");
        try s0.writeRow(&.{.{ .string = "x" }});
        var s1 = try w.addSheet("More");
        try s1.writeRow(&.{.{ .string = "y" }});
        try w.save(io, tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, io, tmp_path);
    defer book.deinit();

    var scratch: [256]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    for (book.sheets) |s| {
        try w.writeAll(s.name);
        try w.writeByte('\n');
    }
    try std.testing.expectEqualStrings("Data\nMore\n", w.buffered());
}

// ─── iter58 tests ────────────────────────────────────────────────────

test "parseArgs routes iter58 sub-commands correctly" {
    const names = [_][]const u8{ "comments", "validations", "hyperlinks", "styles", "sst" };
    const expected = [_]Subcommand{ .comments, .validations, .hyperlinks, .styles, .sst };
    for (names, expected) |n, want| {
        const argv = [_][]const u8{ n, "file.xlsx" };
        const a = try parseArgs(&argv);
        try std.testing.expectEqual(want, a.subcommand);
        try std.testing.expectEqualStrings("file.xlsx", a.file);
    }
    // Styles / sst are workbook-scoped — bogus --sheet / --format
    // must be tolerated (per iter57's P2 fix).
    {
        const argv = [_][]const u8{ "styles", "f.xlsx", "--format", "bogus" };
        const a = try parseArgs(&argv);
        try std.testing.expectEqual(Subcommand.styles, a.subcommand);
    }
    {
        const argv = [_][]const u8{ "sst", "f.xlsx", "--sheet", "bogus" };
        const a = try parseArgs(&argv);
        try std.testing.expectEqual(Subcommand.sst, a.subcommand);
    }
    // Comments / validations / hyperlinks ARE sheet-scoped as of
    // iter58-P2 follow-up — bogus --sheet / --format must error so
    // callers don't get silently-misrouted output.
    {
        const argv = [_][]const u8{ "comments", "f.xlsx", "--sheet", "bogus" };
        try std.testing.expectError(ArgError.BadSheetIndex, parseArgs(&argv));
    }
    {
        const argv = [_][]const u8{ "hyperlinks", "f.xlsx", "--format", "bogus" };
        try std.testing.expectError(ArgError.BadFormat, parseArgs(&argv));
    }
    // Valid --sheet narrows the filter on sheet-scoped sub-commands.
    {
        const argv = [_][]const u8{ "comments", "f.xlsx", "--sheet", "1" };
        const a = try parseArgs(&argv);
        try std.testing.expectEqual(@as(?usize, 1), a.sheet_index);
    }
}

test "runCommentsCommand emits one record per comment across every sheet" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, io, "cli_comments_iter58.xlsx");
    defer std.testing.allocator.free(tmp_path);
    {
        const writer = xlsx.writer_types;
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s0 = try w.addSheet("Data");
        try s0.writeRow(&.{.{ .string = "hdr" }});
        try s0.addComment("A1", "Alice", "needs review");
        var s1 = try w.addSheet("Other");
        try s1.writeRow(&.{.{ .integer = 1 }});
        try s1.addComment("B2", "Bob", "hi");
        try w.save(io, tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, io, tmp_path);
    defer book.deinit();

    var scratch: [2048]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    // Args carries default flags; filter=null + no --all-sheets/--glob
    // preserves the legacy "iterate every sheet" default for this sub-cmd.
    const default_args: Args = .{ .file = "", .subcommand = .comments };
    try runCommentsCommand(&w, &book, null, default_args, null, null, null, null);

    const out = w.buffered();
    try std.testing.expect(std.mem.startsWith(u8, out, "{\"kind\":\"comment\""));
    try std.testing.expect(std.mem.indexOf(u8, out, "\"sheet\":\"Data\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"ref\":\"A1\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"row\":1,\"col\":1") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"author\":\"Alice\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"text\":\"needs review\"") != null);
    // Plain comments emit a bare `<text><t>...</t></text>` body
    // (no synthetic `<r>` run wrapper), so the reader surfaces
    // `runs:null` per the plain/rich distinction.
    try std.testing.expect(std.mem.indexOf(u8, out, "\"runs\":null") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"sheet\":\"Other\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"ref\":\"B2\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"row\":2,\"col\":2") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"author\":\"Bob\"") != null);
}

test "runValidationsCommand emits list validation with values array" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, io, "cli_validations_iter58.xlsx");
    defer std.testing.allocator.free(tmp_path);
    {
        const writer = xlsx.writer_types;
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s0 = try w.addSheet("Data");
        try s0.writeRow(&.{.{ .string = "fruit" }});
        try s0.addDataValidationList("B2:B100", &.{ "apple", "banana" });
        try w.save(io, tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, io, tmp_path);
    defer book.deinit();

    var scratch: [2048]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    const default_args: Args = .{ .file = "", .subcommand = .validations };
    try runValidationsCommand(&w, &book, null, default_args, null, null);

    const out = w.buffered();
    try std.testing.expect(std.mem.startsWith(u8, out, "{\"kind\":\"validation\""));
    try std.testing.expect(std.mem.indexOf(u8, out, "\"sheet\":\"Data\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"range\":\"B2:B100\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"rule_type\":\"list\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"op\":null") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"formula2\":null") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"values\":[\"apple\",\"banana\"]") != null);
}

test "runHyperlinksCommand emits url set + location null for external links" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, io, "cli_hyperlinks_iter58.xlsx");
    defer std.testing.allocator.free(tmp_path);
    {
        const writer = xlsx.writer_types;
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s0 = try w.addSheet("Data");
        try s0.writeRow(&.{.{ .string = "site" }});
        try s0.addHyperlink("A2", "https://example.com/");
        try w.save(io, tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, io, tmp_path);
    defer book.deinit();

    var scratch: [2048]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    const default_args: Args = .{ .file = "", .subcommand = .hyperlinks };
    try runHyperlinksCommand(&w, &book, null, default_args, null, null);

    const out = w.buffered();
    try std.testing.expect(std.mem.startsWith(u8, out, "{\"kind\":\"hyperlink\""));
    try std.testing.expect(std.mem.indexOf(u8, out, "\"sheet\":\"Data\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"range\":\"A2\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"url\":\"https://example.com/\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"location\":null") != null);
}

test "runStylesCommand emits one record per cell-XF entry" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, io, "cli_styles_iter58.xlsx");
    defer std.testing.allocator.free(tmp_path);
    {
        const writer = xlsx.writer_types;
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        _ = try w.addStyle(.{ .font_bold = true });
        var s0 = try w.addSheet("Data");
        try s0.writeRow(&.{.{ .string = "hdr" }});
        try w.save(io, tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, io, tmp_path);
    defer book.deinit();

    var scratch: [4096]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    try runStylesCommand(&w, &book, null, null);

    const out = w.buffered();
    try std.testing.expect(std.mem.startsWith(u8, out, "{\"kind\":\"style\""));
    try std.testing.expect(std.mem.indexOf(u8, out, "\"idx\":0") != null);
    // The bold style registered at addStyle idx=1 (idx 0 is the default
    // no-style xf slot); the record MUST surface with bold:true.
    try std.testing.expect(std.mem.indexOf(u8, out, "\"idx\":1") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"bold\":true") != null);
    // Each record also pins font / fill / border / num_fmt fields.
    try std.testing.expect(std.mem.indexOf(u8, out, "\"font\":") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"fill\":") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"border\":") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"num_fmt\":") != null);
}

test "runSstCommand emits one record per shared-string entry" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, io, "cli_sst_iter58.xlsx");
    defer std.testing.allocator.free(tmp_path);
    {
        const writer = xlsx.writer_types;
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s0 = try w.addSheet("Data");
        try s0.writeRow(&.{ .{ .string = "header" }, .{ .string = "qty" } });
        try s0.writeRow(&.{ .{ .string = "apple" }, .{ .integer = 3 } });
        try w.save(io, tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, io, tmp_path);
    defer book.deinit();

    var scratch: [4096]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    try runSstCommand(&w, &book, null, null);

    const out = w.buffered();
    try std.testing.expect(std.mem.startsWith(u8, out, "{\"kind\":\"sst\""));
    try std.testing.expect(std.mem.indexOf(u8, out, "\"idx\":0") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"text\":\"header\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"text\":\"qty\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"text\":\"apple\"") != null);
    // Plain strings — runs must be null on every record.
    try std.testing.expect(std.mem.indexOf(u8, out, "\"runs\":null") != null);
}

// ─── Fuzz tests ──────────────────────────────────────────────────────

fn fuzzItersCli() usize {
    // Override comes from build.zig via -Dfuzz-iters or the
    // XLSX_FUZZ_ITERS environment variable; 0.16 test binaries
    // cannot read the environment themselves.
    return fuzz_config.iters_override orelse 1_000;
}

fn fuzzSeedCli(io: std.Io) u64 {
    if (fuzz_config.seed_override) |s| return s;
    // std.time lost every function in 0.16; a varying default
    // seed now comes from the monotonic clock via Io.
    const ts = std.Io.Clock.now(.awake, io);
    return @bitCast(@as(i64, @truncate(ts.nanoseconds)));
}

test "fuzz colLetter: output is uppercase A-Z" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    const iters = fuzzItersCli();
    var prng = std.Random.DefaultPrng.init(fuzzSeedCli(io));
    const rng = prng.random();
    var buf: [8]u8 = undefined;
    for (0..iters) |_| {
        // xlsx max is column 16383 (XFD); cap at 2^20 — beyond that the
        // 8-byte buffer can't fit all letters and the function would
        // wrap around via pos underflow. This is documented: caller is
        // expected to stay within OOXML's column range.
        const idx = rng.intRangeAtMost(usize, 0, 1_048_575);
        const letters = colLetter(&buf, idx);
        try std.testing.expect(letters.len >= 1);
        for (letters) |c| {
            try std.testing.expect(c >= 'A' and c <= 'Z');
        }
    }
}

test "fuzz parseArgs: arbitrary tokens never panic" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    const iters = fuzzItersCli();
    var prng = std.Random.DefaultPrng.init(fuzzSeedCli(io));
    const rng = prng.random();

    var token_pool: [8][32]u8 = undefined;
    for (0..token_pool.len) |i| rng.bytes(&token_pool[i]);

    for (0..iters) |_| {
        const n_tokens = rng.intRangeAtMost(usize, 0, 12);
        var argv_buf: [12][]const u8 = undefined;
        for (0..n_tokens) |i| {
            const k = rng.intRangeAtMost(usize, 0, token_pool.len - 1);
            const l = rng.intRangeAtMost(usize, 0, token_pool[k].len);
            argv_buf[i] = token_pool[k][0..l];
        }
        // Mix in a few well-known tokens so we hit more branches.
        if (n_tokens >= 1 and rng.boolean()) argv_buf[0] = "--help";
        if (n_tokens >= 2 and rng.boolean()) argv_buf[1] = "--format";

        // Must never panic; errors are fine.
        _ = parseArgs(argv_buf[0..n_tokens]) catch {};
    }
}

test "fuzz writeJsonString: no raw control chars survive" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    const iters = fuzzItersCli();
    var prng = std.Random.DefaultPrng.init(fuzzSeedCli(io));
    const rng = prng.random();

    var input: [256]u8 = undefined;
    var scratch: [4096]u8 = undefined;

    for (0..iters) |_| {
        const l = rng.intRangeAtMost(usize, 0, input.len);
        rng.bytes(input[0..l]);
        var w = std.Io.Writer.fixed(&scratch);
        writeJsonString(&w, input[0..l]) catch continue;

        const out = w.buffered();
        try std.testing.expect(out.len >= 2); // at least "\"\""
        try std.testing.expect(out[0] == '"');
        try std.testing.expect(out[out.len - 1] == '"');

        // Walk the interior (between the outer quotes). No bare control
        // char (0..0x1f) except when preceded by a backslash. Quote and
        // backslash always escaped too.
        var i: usize = 1;
        while (i < out.len - 1) : (i += 1) {
            const c = out[i];
            if (c == '\\') {
                // Skip the escape sequence (\", \\, \n, \r, \t, \b, \f, \uXXXX).
                i += 1;
                if (i < out.len - 1 and out[i] == 'u') i += 4;
                continue;
            }
            try std.testing.expect(c >= 0x20);
            try std.testing.expect(c != '"');
        }
    }
}

test "fuzz writeCsvField: balanced quotes + no bare quote outside them" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    const iters = fuzzItersCli();
    var prng = std.Random.DefaultPrng.init(fuzzSeedCli(io));
    const rng = prng.random();

    var input: [256]u8 = undefined;
    var scratch: [4096]u8 = undefined;

    for (0..iters) |_| {
        const l = rng.intRangeAtMost(usize, 0, input.len);
        rng.bytes(input[0..l]);
        var w = std.Io.Writer.fixed(&scratch);
        writeCsvField(&w, input[0..l]) catch continue;

        const out = w.buffered();
        // If any RFC-4180 trigger byte was present, output must be
        // quoted. Otherwise unquoted is fine.
        var needs_quote = false;
        for (input[0..l]) |c| {
            if (c == ',' or c == '"' or c == '\n' or c == '\r') {
                needs_quote = true;
                break;
            }
        }
        if (needs_quote) {
            try std.testing.expect(out.len >= 2);
            try std.testing.expectEqual(@as(u8, '"'), out[0]);
            try std.testing.expectEqual(@as(u8, '"'), out[out.len - 1]);
            // Every `"` inside must be doubled.
            var i: usize = 1;
            while (i < out.len - 1) : (i += 1) {
                if (out[i] == '"') {
                    try std.testing.expect(i + 1 < out.len - 1 and out[i + 1] == '"');
                    i += 1;
                }
            }
        }
    }
}

// ─── iter59c: --all-sheets / --sheet-glob ──────────────────────────

test "parseArgs --all-sheets alone sets the flag" {
    const argv = [_][]const u8{ "cells", "f.xlsx", "--all-sheets" };
    const a = try parseArgs(&argv);
    try std.testing.expect(a.all_sheets);
    try std.testing.expect(a.sheet_glob == null);
    try std.testing.expect(a.sheet_index == null);
    try std.testing.expect(a.sheet_name == null);
}

test "parseArgs --sheet-glob alone stores the pattern" {
    const argv = [_][]const u8{ "cells", "f.xlsx", "--sheet-glob", "Data*" };
    const a = try parseArgs(&argv);
    try std.testing.expect(!a.all_sheets);
    try std.testing.expectEqualStrings("Data*", a.sheet_glob.?);
}

test "parseArgs rejects --all-sheets combined with --sheet" {
    {
        const argv = [_][]const u8{ "cells", "f.xlsx", "--all-sheets", "--sheet", "0" };
        try std.testing.expectError(ArgError.SheetArgConflict, parseArgs(&argv));
    }
    {
        const argv = [_][]const u8{ "cells", "f.xlsx", "--sheet", "0", "--all-sheets" };
        try std.testing.expectError(ArgError.SheetArgConflict, parseArgs(&argv));
    }
}

test "parseArgs rejects --sheet-glob combined with --name" {
    const argv = [_][]const u8{ "cells", "f.xlsx", "--name", "Sheet1", "--sheet-glob", "S*" };
    try std.testing.expectError(ArgError.SheetArgConflict, parseArgs(&argv));
}

test "parseArgs rejects --all-sheets combined with --sheet-glob" {
    const argv = [_][]const u8{ "cells", "f.xlsx", "--all-sheets", "--sheet-glob", "S*" };
    try std.testing.expectError(ArgError.SheetArgConflict, parseArgs(&argv));
}

test "parseArgs tolerates --all-sheets / --sheet-glob on workbook-scoped sub-commands" {
    // Wrappers that set these flags universally must still reach
    // meta / list-sheets / styles / sst without exit-1 (same tolerance
    // group as --sheet / --name per the iter58 design).
    inline for (.{ "meta", "list-sheets", "styles", "sst" }) |cmd| {
        {
            const argv = [_][]const u8{ cmd, "f.xlsx", "--all-sheets" };
            const a = try parseArgs(&argv);
            try std.testing.expect(a.all_sheets);
        }
        {
            const argv = [_][]const u8{ cmd, "f.xlsx", "--sheet-glob", "*" };
            const a = try parseArgs(&argv);
            try std.testing.expect(a.sheet_glob != null);
        }
    }
}

test "parseArgs --sheet-glob value isn't mistaken for a sub-command token" {
    // detectSubcommand must skip the value of --sheet-glob so a pattern
    // that happens to equal a sub-command name ("cells", "rows", …)
    // doesn't re-route the subcommand decision. Regression guard for
    // detectSubcommand's skip-pair list.
    const argv = [_][]const u8{ "rows", "--sheet-glob", "cells", "f.xlsx" };
    const a = try parseArgs(&argv);
    try std.testing.expectEqual(Subcommand.rows, a.subcommand);
    try std.testing.expectEqualStrings("f.xlsx", a.file);
    try std.testing.expectEqualStrings("cells", a.sheet_glob.?);
}

test "parseArgs --output parses the three valid modes" {
    {
        const argv = [_][]const u8{ "cells", "--output", "ndjson", "f.xlsx" };
        const a = try parseArgs(&argv);
        try std.testing.expectEqual(OutputMode.ndjson, a.output);
    }
    {
        const argv = [_][]const u8{ "cells", "--output", "compact-ndjson", "f.xlsx" };
        const a = try parseArgs(&argv);
        try std.testing.expectEqual(OutputMode.compact_ndjson, a.output);
    }
    {
        const argv = [_][]const u8{ "meta", "--output", "pretty-json", "f.xlsx" };
        const a = try parseArgs(&argv);
        try std.testing.expectEqual(OutputMode.pretty_json, a.output);
    }
}

test "parseArgs --output rejects unknown values" {
    const argv = [_][]const u8{ "cells", "--output", "yaml", "f.xlsx" };
    try std.testing.expectError(ArgError.BadArgValue, parseArgs(&argv));
}

test "parseArgs --output missing value is MissingValue" {
    const argv = [_][]const u8{ "cells", "--output" };
    try std.testing.expectError(ArgError.MissingValue, parseArgs(&argv));
}

test "parseArgs --output pretty-json is rejected on streaming sub-commands" {
    // Z4: `list-sheets` moved OUT of this list — it is workbook-scoped
    // and bounded, so a collapsed object is a coherent shape and
    // callers gating on sheet visibility want the whole answer at once.
    // Everything below still streams record-per-line.
    const subs = [_][]const u8{
        "cells",       "rows",       "comments",
        "validations", "hyperlinks", "styles",
        "sst",
    };
    for (subs) |sub| {
        const argv = [_][]const u8{ sub, "--output", "pretty-json", "f.xlsx" };
        try std.testing.expectError(ArgError.BadArgValue, parseArgs(&argv));
    }
    // list-sheets now accepts it.
    const argv_ls = [_][]const u8{ "list-sheets", "--output", "pretty-json", "f.xlsx" };
    const parsed = try parseArgs(&argv_ls);
    try std.testing.expectEqual(OutputMode.pretty_json, parsed.output);
    // Bare `zlsx file.xlsx` defaults to the `rows` sub-command, which
    // also cannot accept pretty-json.
    const argv_bare = [_][]const u8{ "--output", "pretty-json", "f.xlsx" };
    try std.testing.expectError(ArgError.BadArgValue, parseArgs(&argv_bare));
}

test "detectSubcommand skips --output value (value that collides with a sub-command token)" {
    // Regression guard: if detectSubcommand didn't consume --output's
    // paired value, `--output cells` would re-route the sub-command to
    // `.cells`. The paired-value skip list must include `--output`.
    const argv = [_][]const u8{ "rows", "--output", "ndjson", "f.xlsx" };
    const a = try parseArgs(&argv);
    try std.testing.expectEqual(Subcommand.rows, a.subcommand);
    try std.testing.expectEqualStrings("f.xlsx", a.file);
}

test "runCellsAcrossSheets compact-ndjson emits per-sheet prologue and omits sheet/sheet_idx on cell records" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, io, "cli_compact_cells_iter60b.xlsx");
    defer std.testing.allocator.free(tmp_path);
    {
        const writer = xlsx.writer_types;
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s0 = try w.addSheet("Data");
        try s0.writeRow(&.{ .{ .string = "name" }, .{ .string = "qty" } });
        try s0.writeRow(&.{ .{ .string = "a" }, .{ .integer = 1 } });
        var s1 = try w.addSheet("Other");
        try s1.writeRow(&.{.{ .string = "x" }});
        try w.save(io, tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, io, tmp_path);
    defer book.deinit();

    var buf: [4096]u8 = undefined;
    var w = std.Io.Writer.fixed(&buf);
    const args: Args = .{
        .file = tmp_path,
        .subcommand = .cells,
        .all_sheets = true,
        .output = .compact_ndjson,
    };
    try runCellsAcrossSheets(&w, &book, args, std.testing.allocator);

    const out = w.buffered();

    // Count prologue records — one per sheet.
    const prologue = "{\"kind\":\"sheet\",";
    var i: usize = 0;
    var prologues: usize = 0;
    while (std.mem.indexOfPos(u8, out, i, prologue)) |p| {
        prologues += 1;
        i = p + prologue.len;
    }
    try std.testing.expectEqual(@as(usize, 2), prologues);

    // Count cell records — 4 non-empty cells total (2 on Data row 1,
    // 2 on Data row 2, and 1 on Other; that's 5).
    const cell_tag = "{\"kind\":\"cell\",";
    i = 0;
    var cells_count: usize = 0;
    while (std.mem.indexOfPos(u8, out, i, cell_tag)) |p| {
        cells_count += 1;
        i = p + cell_tag.len;
    }
    try std.testing.expectEqual(@as(usize, 5), cells_count);

    // Every cell line must start with `{"kind":"cell","ref":` — the
    // `sheet`/`sheet_idx` fields must be absent on cell records.
    var line_it = std.mem.splitScalar(u8, out, '\n');
    while (line_it.next()) |line| {
        if (line.len == 0) continue;
        if (std.mem.startsWith(u8, line, "{\"kind\":\"cell\",")) {
            try std.testing.expect(std.mem.indexOf(u8, line, "\"sheet\":") == null);
            try std.testing.expect(std.mem.indexOf(u8, line, "\"sheet_idx\":") == null);
            try std.testing.expect(std.mem.startsWith(u8, line, "{\"kind\":\"cell\",\"ref\":"));
        }
    }

    // First line is the Data prologue, since Data is sheet 0.
    try std.testing.expect(std.mem.startsWith(
        u8,
        out,
        "{\"kind\":\"sheet\",\"sheet\":\"Data\",\"sheet_idx\":0}\n",
    ));
}

test "runMetaCommand pretty-json collapses workbook + sheets into one JSON object" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, io, "cli_pretty_meta_iter60b.xlsx");
    defer std.testing.allocator.free(tmp_path);
    {
        const writer = xlsx.writer_types;
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s0 = try w.addSheet("Data");
        try s0.writeRow(&.{.{ .string = "hdr" }});
        var s1 = try w.addSheet("Other");
        try s1.writeRow(&.{.{ .integer = 1 }});
        try w.save(io, tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, io, tmp_path);
    defer book.deinit();

    var buf: [4096]u8 = undefined;
    var w = std.Io.Writer.fixed(&buf);
    try runMetaCommand(&w, &book, tmp_path, .pretty_json, null);

    const out = w.buffered();

    // std.json must accept it — that alone is the strongest structural
    // assertion (balanced braces, valid strings/numbers, etc.).
    const parsed = try std.json.parseFromSlice(std.json.Value, std.testing.allocator, out, .{});
    defer parsed.deinit();

    const root = parsed.value.object;
    try std.testing.expectEqualStrings("workbook", root.get("kind").?.string);
    // sheets_count scalar — see OutputMode doc for why we renamed the
    // scalar in this mode only.
    try std.testing.expectEqual(@as(i64, 2), root.get("sheets_count").?.integer);
    const sheets_arr = root.get("sheets").?.array;
    try std.testing.expectEqual(@as(usize, 2), sheets_arr.items.len);
    try std.testing.expectEqualStrings(
        "sheet",
        sheets_arr.items[0].object.get("kind").?.string,
    );
    try std.testing.expectEqualStrings(
        "Data",
        sheets_arr.items[0].object.get("sheet").?.string,
    );
    try std.testing.expectEqual(
        @as(i64, 0),
        sheets_arr.items[0].object.get("sheet_idx").?.integer,
    );
    try std.testing.expectEqualStrings(
        "Other",
        sheets_arr.items[1].object.get("sheet").?.string,
    );

    // Indentation sanity: 2-space prefix on each inner key.
    try std.testing.expect(std.mem.indexOf(u8, out, "\n  \"kind\":") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\n  \"sheets\":") != null);
}

test "parseArgs default --output is ndjson" {
    // Regression guard: every sub-command without an explicit --output
    // resolves to OutputMode.ndjson so the wire-shape default stays
    // exactly what iter60a shipped.
    const argv = [_][]const u8{ "cells", "f.xlsx" };
    const a = try parseArgs(&argv);
    try std.testing.expectEqual(OutputMode.ndjson, a.output);
}

test "globMatch literal / wildcards / edge cases" {
    // Literal
    try std.testing.expect(globMatch("Sheet1", "Sheet1"));
    try std.testing.expect(!globMatch("Sheet1", "Sheet2"));
    // `*` runs
    try std.testing.expect(globMatch("*", ""));
    try std.testing.expect(globMatch("*", "anything"));
    try std.testing.expect(globMatch("Data*", "Data"));
    try std.testing.expect(globMatch("Data*", "Data123"));
    try std.testing.expect(!globMatch("Data*", "NoData"));
    try std.testing.expect(globMatch("*Data", "XData"));
    try std.testing.expect(globMatch("*Data*", "XDataY"));
    try std.testing.expect(globMatch("a*b*c", "abc"));
    try std.testing.expect(globMatch("a*b*c", "a123b456c"));
    try std.testing.expect(!globMatch("a*b*c", "a123b456"));
    // `?` exact single char
    try std.testing.expect(globMatch("S?2", "Sh2"));
    try std.testing.expect(!globMatch("S?2", "S2"));
    try std.testing.expect(!globMatch("S?2", "Sh22"));
    // Empty pattern vs empty text
    try std.testing.expect(globMatch("", ""));
    try std.testing.expect(!globMatch("", "x"));
    try std.testing.expect(!globMatch("x", ""));
    // Pattern longer than input
    try std.testing.expect(!globMatch("abcd", "abc"));
    // Case-sensitive
    try std.testing.expect(!globMatch("sheet", "Sheet"));
    // Consecutive `*` collapse
    try std.testing.expect(globMatch("**", "hello"));
    try std.testing.expect(globMatch("a***b", "axyzb"));
    // UTF-8: `?` matches one codepoint, not one byte.
    // "é" = 2 bytes (0xC3 0xA9); "表" = 3 bytes; "𝕊" = 4 bytes.
    try std.testing.expect(globMatch("R?sumé", "Résumé"));
    try std.testing.expect(globMatch("?1", "表1"));
    try std.testing.expect(globMatch("?", "é"));
    try std.testing.expect(globMatch("?", "表"));
    try std.testing.expect(globMatch("?", "𝕊"));
    // Multi-`?` + non-ASCII.
    try std.testing.expect(globMatch("??", "éé"));
    try std.testing.expect(!globMatch("??", "é")); // only one char
}

test "runCellsAcrossSheets --all-sheets emits every sheet with correct sheet_idx" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, io, "cli_all_sheets_iter59c.xlsx");
    defer std.testing.allocator.free(tmp_path);
    {
        const writer = xlsx.writer_types;
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s0 = try w.addSheet("Alpha");
        try s0.writeRow(&.{.{ .string = "A1_alpha" }});
        var s1 = try w.addSheet("Beta");
        try s1.writeRow(&.{.{ .string = "A1_beta" }});
        try w.save(io, tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, io, tmp_path);
    defer book.deinit();

    var scratch: [2048]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    const args: Args = .{ .file = "", .subcommand = .cells, .all_sheets = true };
    try runCellsAcrossSheets(&w, &book, args, std.testing.allocator);
    const out = w.buffered();

    // Alpha record first, Beta second — sheet_idx monotonic.
    try std.testing.expect(std.mem.indexOf(u8, out, "\"sheet\":\"Alpha\",\"sheet_idx\":0") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"sheet\":\"Beta\",\"sheet_idx\":1") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"v\":\"A1_alpha\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"v\":\"A1_beta\"") != null);
}

test "runCellsAcrossSheets --sheet-glob selects only matching sheets" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, io, "cli_glob_iter59c.xlsx");
    defer std.testing.allocator.free(tmp_path);
    {
        const writer = xlsx.writer_types;
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s0 = try w.addSheet("Sheet1");
        try s0.writeRow(&.{.{ .string = "v1" }});
        var s1 = try w.addSheet("Sheet2");
        try s1.writeRow(&.{.{ .string = "v2" }});
        var s2 = try w.addSheet("Data3");
        try s2.writeRow(&.{.{ .string = "v3" }});
        try w.save(io, tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, io, tmp_path);
    defer book.deinit();

    var scratch: [2048]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    // `S*2` matches "Sheet2" only.
    const args: Args = .{ .file = "", .subcommand = .cells, .sheet_glob = "S*2" };
    try runCellsAcrossSheets(&w, &book, args, std.testing.allocator);
    const out = w.buffered();

    try std.testing.expect(std.mem.indexOf(u8, out, "\"v\":\"v2\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"v\":\"v1\"") == null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"v\":\"v3\"") == null);
    // `sheet_idx` for Sheet2 is 1, not 0 — the emitter uses the real
    // book position, not a filtered-stream ordinal.
    try std.testing.expect(std.mem.indexOf(u8, out, "\"sheet_idx\":1") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"sheet_idx\":0") == null);
}

test "runCellsAcrossSheets --all-sheets --skip --take slices the cross-sheet stream" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, io, "cli_cross_pag_iter59c.xlsx");
    defer std.testing.allocator.free(tmp_path);
    {
        const writer = xlsx.writer_types;
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s0 = try w.addSheet("A");
        // 3 cells on sheet 0 → a, b, c.
        try s0.writeRow(&.{ .{ .string = "a" }, .{ .string = "b" }, .{ .string = "c" } });
        var s1 = try w.addSheet("B");
        // 3 cells on sheet 1 → d, e, f.
        try s1.writeRow(&.{ .{ .string = "d" }, .{ .string = "e" }, .{ .string = "f" } });
        try w.save(io, tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, io, tmp_path);
    defer book.deinit();

    var scratch: [4096]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    // Concatenated stream: a, b, c, d, e, f. --skip 2 --take 3 → c, d, e.
    const args: Args = .{
        .file = "",
        .subcommand = .cells,
        .all_sheets = true,
        .skip = 2,
        .take = 3,
    };
    try runCellsAcrossSheets(&w, &book, args, std.testing.allocator);
    const out = w.buffered();

    const countLines = struct {
        fn f(s: []const u8) usize {
            var n: usize = 0;
            for (s) |c| if (c == '\n') {
                n += 1;
            };
            return n;
        }
    }.f;

    try std.testing.expectEqual(@as(usize, 3), countLines(out));
    try std.testing.expect(std.mem.indexOf(u8, out, "\"v\":\"a\"") == null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"v\":\"b\"") == null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"v\":\"c\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"v\":\"d\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"v\":\"e\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"v\":\"f\"") == null);
    // Cross-sheet: both sheets must contribute at least one record.
    try std.testing.expect(std.mem.indexOf(u8, out, "\"sheet\":\"A\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"sheet\":\"B\"") != null);
}

test "writeErrorRecord sheet-scoped + workbook-scoped shapes (iter60c)" {
    // Sheet-scoped, ndjson — every identity field present.
    {
        var scratch: [256]u8 = undefined;
        var w = std.Io.Writer.fixed(&scratch);
        try writeErrorRecord(&w, "Data", 0, "sheet", "MalformedXml", "malformed sheet XML");
        try std.testing.expectEqualStrings(
            "{\"kind\":\"error\",\"sheet\":\"Data\",\"sheet_idx\":0,\"scope\":\"sheet\",\"code\":\"MalformedXml\",\"message\":\"malformed sheet XML\"}\n",
            w.buffered(),
        );
    }
    // Workbook-scoped — null sheet/sheet_idx, both omitted.
    {
        var scratch: [256]u8 = undefined;
        var w = std.Io.Writer.fixed(&scratch);
        try writeErrorRecord(&w, null, null, "workbook", "ArchiveClosed", "archive closed");
        try std.testing.expectEqualStrings(
            "{\"kind\":\"error\",\"scope\":\"workbook\",\"code\":\"ArchiveClosed\",\"message\":\"archive closed\"}\n",
            w.buffered(),
        );
    }
    // String fields containing JSON-special chars are escaped.
    {
        var scratch: [256]u8 = undefined;
        var w = std.Io.Writer.fixed(&scratch);
        try writeErrorRecord(&w, "Q\"uoted", 1, "sheet", "Code", "line\nbreak");
        const out = w.buffered();
        try std.testing.expect(std.mem.indexOf(u8, out, "\"sheet\":\"Q\\\"uoted\"") != null);
        try std.testing.expect(std.mem.indexOf(u8, out, "\"message\":\"line\\nbreak\"") != null);
    }
}

test "runCellsAcrossSheets emits inline kind:error for a malformed sheet and continues (iter60c)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    // Two-sheet workbook: sheet 0 valid, sheet 1's loaded XML is
    // hot-replaced after open with bytes that trip
    // `consumeRow → indexOfScalarPos('<') == null → error.MalformedXml`.
    // The CLI must emit one inline `kind:"error"` record at sheet
    // boundary, keep going, and exit cleanly (no propagation).
    const tmp_path = try tt.path(std.testing.allocator, io, "cli_iter60c_malformed.xlsx");
    defer std.testing.allocator.free(tmp_path);
    {
        const writer = xlsx.writer_types;
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s0 = try w.addSheet("Good");
        try s0.writeRow(&.{ .{ .string = "g1" }, .{ .string = "g2" } });
        var s1 = try w.addSheet("Bad");
        try s1.writeRow(&.{.{ .string = "b1" }});
        try w.save(io, tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, io, tmp_path);
    defer book.deinit();

    // Replace sheet 1's loaded XML with a payload that opens a `<row>`
    // but never closes it / emits no further `<` — `consumeRow` fails on
    // the missing close-tag scan. The replacement uses `book.allocator`
    // so `book.deinit` frees it through the same path as the original.
    const bad_path = book.sheets[1].path;
    const corrupt =
        "<?xml version=\"1.0\"?>" ++
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" ++
        "<sheetData><row r=\"1\">no-close-tag-here";
    const dup = try book.allocator.dupe(u8, corrupt);
    if (book.sheet_data.getEntry(bad_path)) |entry| {
        book.allocator.free(entry.value_ptr.*);
        entry.value_ptr.* = dup;
    } else {
        book.allocator.free(dup);
        return error.SheetDataMissing;
    }

    var scratch: [4096]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    const args: Args = .{ .file = "", .subcommand = .cells, .all_sheets = true };
    try runCellsAcrossSheets(&w, &book, args, std.testing.allocator);
    const out = w.buffered();

    // Good sheet's two cells survive.
    try std.testing.expect(std.mem.indexOf(u8, out, "\"v\":\"g1\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"v\":\"g2\"") != null);
    // Bad sheet emits exactly one error record carrying its identity.
    try std.testing.expect(std.mem.indexOf(u8, out, "\"kind\":\"error\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"sheet\":\"Bad\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"sheet_idx\":1") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"scope\":\"sheet\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"code\":\"MalformedXml\"") != null);
}

test "runCellsCommand single-sheet malformed sheet emits one error record without propagating (iter60c)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, io, "cli_iter60c_single.xlsx");
    defer std.testing.allocator.free(tmp_path);
    {
        const writer = xlsx.writer_types;
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s0 = try w.addSheet("Only");
        try s0.writeRow(&.{.{ .string = "x" }});
        try w.save(io, tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, io, tmp_path);
    defer book.deinit();

    const bad_path = book.sheets[0].path;
    const corrupt =
        "<?xml version=\"1.0\"?>" ++
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" ++
        "<sheetData><row r=\"1\">unterminated";
    const dup = try book.allocator.dupe(u8, corrupt);
    if (book.sheet_data.getEntry(bad_path)) |entry| {
        book.allocator.free(entry.value_ptr.*);
        entry.value_ptr.* = dup;
    } else {
        book.allocator.free(dup);
        return error.SheetDataMissing;
    }

    var scratch: [2048]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    // Must NOT propagate — this is the "non-fatal" contract the design
    // doc describes for a corrupt sheet inside an otherwise-open book.
    try runCellsCommand(&w, &book, book.sheets[0], 0, std.testing.allocator, null, null, null, null, null, false, false);

    const out = w.buffered();
    try std.testing.expect(std.mem.indexOf(u8, out, "\"kind\":\"error\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"sheet\":\"Only\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"code\":\"MalformedXml\"") != null);
}

test "writeCell emits t:date with ISO v and raw serial" {
    var scratch: [512]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    // Excel serial 45458 = 2024-06-15 (date-only, so the fractional
    // part is 0 and {d} prints as `45458`).
    try writeCell(&w, "Data", 0, "B3", 3, 2, .{ .integer = 45458 }, null, false, true, null, null, null, false);
    try std.testing.expectEqualStrings(
        "{\"kind\":\"cell\",\"sheet\":\"Data\",\"sheet_idx\":0,\"ref\":\"B3\",\"row\":3,\"col\":2,\"t\":\"date\",\"v\":\"2024-06-15T00:00:00\",\"serial\":45458}\n",
        w.buffered(),
    );
}

test "writeRowEnvelope emits t:date inside cells array" {
    var scratch: [512]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    const cells = [_]xlsx.Cell{
        .{ .string = "name" }, // A1, not a date
        .{ .integer = 45458 }, // B1, date-styled
    };
    const dates = [_]bool{ false, true };
    try writeRowEnvelope(&w, "Data", 0, 1, &cells, false, null, 0, false, &dates, &.{}, &.{}, &.{}, false);
    try std.testing.expectEqualStrings(
        "{\"kind\":\"row\",\"sheet\":\"Data\",\"sheet_idx\":0,\"row\":1,\"cells\":[" ++
            "{\"ref\":\"A1\",\"col\":1,\"t\":\"str\",\"v\":\"name\"}," ++
            "{\"ref\":\"B1\",\"col\":2,\"t\":\"date\",\"v\":\"2024-06-15T00:00:00\",\"serial\":45458}" ++
            "]}\n",
        w.buffered(),
    );
}

test "runCellsCommand emits t:date for a date-styled numeric cell (iter61-a)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, io, "cli_iter61a_cells.xlsx");
    defer std.testing.allocator.free(tmp_path);
    {
        const writer = xlsx.writer_types;
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        const date_style = try w.addStyle(.{ .number_format = "yyyy-mm-dd" });
        var sheet = try w.addSheet("Data");
        // Row 1: a plain int (A1) + a date-styled int (B1). We check
        // the plain int still emits `t:"int"` while the date one
        // emits `t:"date"` + `serial`.
        try sheet.writeRowStyled(
            &.{ .{ .integer = 7 }, .{ .integer = 45458 } },
            &.{ 0, date_style },
        );
        try w.save(io, tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, io, tmp_path);
    defer book.deinit();

    var scratch: [4096]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    try runCellsCommand(&w, &book, book.sheets[0], 0, std.testing.allocator, null, null, null, null, null, false, false);

    const out = w.buffered();
    // Plain int stays int.
    try std.testing.expect(std.mem.indexOf(u8, out, "\"ref\":\"A1\",\"row\":1,\"col\":1,\"t\":\"int\",\"v\":7") != null);
    // Date-styled int becomes t:"date" with ISO `v` and raw `serial`.
    try std.testing.expect(std.mem.indexOf(u8, out, "\"ref\":\"B1\",\"row\":1,\"col\":2,\"t\":\"date\",\"v\":\"2024-06-15T00:00:00\",\"serial\":45458") != null);
}

test "runRowsCommand envelope emits t:date inside cells array (iter61-a)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, io, "cli_iter61a_rows.xlsx");
    defer std.testing.allocator.free(tmp_path);
    {
        const writer = xlsx.writer_types;
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        const date_style = try w.addStyle(.{ .number_format = "yyyy-mm-dd" });
        var sheet = try w.addSheet("Data");
        try sheet.writeRowStyled(
            &.{ .{ .integer = 7 }, .{ .integer = 45458 } },
            &.{ 0, date_style },
        );
        try w.save(io, tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, io, tmp_path);
    defer book.deinit();

    var scratch: [4096]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    var pg = Pagination.init(null, null);
    var prologue: bool = false;
    try runRowsOnSheetCore(&w, &book, book.sheets[0], 0, .jsonl, std.testing.allocator, &pg, null, null, null, false, false, false, false, &prologue);

    const out = w.buffered();
    // Single envelope line, int + date side by side.
    try std.testing.expect(std.mem.indexOf(u8, out, "\"kind\":\"row\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "{\"ref\":\"A1\",\"col\":1,\"t\":\"int\",\"v\":7}") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "{\"ref\":\"B1\",\"col\":2,\"t\":\"date\",\"v\":\"2024-06-15T00:00:00\",\"serial\":45458}") != null);
}

test "runCellsCommand skips t:date auto-convert on 1904-epoch workbook" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    // iter61-a P1 follow-up: workbooks with <workbookPr date1904="1"/>
    // shift every serial by 1462 days. Until proper 1904 decoding ships,
    // the CLI must NOT auto-convert these cells to t:"date" — the
    // numeric value is the authoritative signal.
    const tmp_path = try tt.path(std.testing.allocator, io, "cli_iter61a_date1904.xlsx");
    defer std.testing.allocator.free(tmp_path);
    {
        const writer = xlsx.writer_types;
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        const date_style = try w.addStyle(.{ .number_format = "yyyy-mm-dd" });
        var s = try w.addSheet("Data");
        // Serial 45458 = 2024-06-15 under 1900 system. We write a 1900
        // workbook, then flip the flag post-open to simulate a 1904
        // file (the zlsx Writer doesn't emit 1904 workbooks today).
        try s.writeRowStyled(&.{.{ .integer = 45458 }}, &.{date_style});
        try w.save(io, tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, io, tmp_path);
    defer book.deinit();

    // Simulate a 1904-epoch workbook.
    book.uses_1904_epoch = true;

    var scratch: [1024]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    try runCellsCommand(&w, &book, book.sheets[0], 0, std.testing.allocator, null, null, null, null, null, false, false);
    const out = w.buffered();

    // 1904-epoch decoding now ships: serial 45458 in the 1904 system
    // shifts by +1462 days to a 1900-system serial of 46920, which
    // decodes to 2028-06-16T00:00:00. The raw serial is preserved
    // so consumers can re-decode under their own epoch if they want.
    try std.testing.expect(std.mem.indexOf(u8, out, "\"t\":\"date\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"v\":\"2028-06-16T00:00:00\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"serial\":45458") != null);
}

test "writeCell emits t:error with the literal v (iter61-c)" {
    var scratch: [512]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    // OOXML stores the error literal inside <v>; the CLI surfaces it
    // verbatim as the JSON string value. Cell slot stays .string (the
    // same literal) but the `t` comes from the parallel side channel.
    try writeCell(
        &w,
        "Data",
        0,
        "D2",
        2,
        4,
        .{ .string = "#DIV/0!" },
        null,
        false,
        false,
        "#DIV/0!",
        null,
        null,
        false,
    );
    try std.testing.expectEqualStrings(
        "{\"kind\":\"cell\",\"sheet\":\"Data\",\"sheet_idx\":0,\"ref\":\"D2\",\"row\":2,\"col\":4,\"t\":\"error\",\"v\":\"#DIV/0!\"}\n",
        w.buffered(),
    );
}

test "writeRowEnvelope emits t:error inside cells array (iter61-c)" {
    var scratch: [512]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    const cells = [_]xlsx.Cell{
        .{ .string = "name" }, // A1 — plain string, not an error
        .{ .string = "#N/A" }, // B1 — error literal
    };
    const dates = [_]bool{ false, false };
    const errors = [_]?[]const u8{ null, "#N/A" };
    try writeRowEnvelope(&w, "Data", 0, 1, &cells, false, null, 0, false, &dates, &errors, &.{}, &.{}, false);
    try std.testing.expectEqualStrings(
        "{\"kind\":\"row\",\"sheet\":\"Data\",\"sheet_idx\":0,\"row\":1,\"cells\":[" ++
            "{\"ref\":\"A1\",\"col\":1,\"t\":\"str\",\"v\":\"name\"}," ++
            "{\"ref\":\"B1\",\"col\":2,\"t\":\"error\",\"v\":\"#N/A\"}" ++
            "]}\n",
        w.buffered(),
    );
}

test "runCellsCommand emits t:error for a t=\"e\" cell (iter61-c)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    // The zlsx Writer can't emit OOXML t="e" cells directly, so we
    // mirror the iter60c pattern: write a valid workbook, open it,
    // then post-inject a sheet1.xml carrying `<c t="e"><v>#DIV/0!</v></c>`
    // through `book.sheet_data`. The same allocator handles free on
    // deinit.
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, io, "cli_iter61c_error.xlsx");
    defer std.testing.allocator.free(tmp_path);
    {
        const writer = xlsx.writer_types;
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("Data");
        try s.writeRow(&.{.{ .integer = 1 }});
        try w.save(io, tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, io, tmp_path);
    defer book.deinit();

    const sheet_path = book.sheets[0].path;
    // Row 1: A1 plain int, B1 error literal, C1 another error literal.
    // The plain int cell still surfaces as t:"int"; both error cells
    // surface as t:"error" with their literal v, matching docs/jq-for-excel.md.
    const injected =
        "<?xml version=\"1.0\"?>" ++
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" ++
        "<sheetData>" ++
        "<row r=\"1\">" ++
        "<c r=\"A1\"><v>1</v></c>" ++
        "<c r=\"B1\" t=\"e\"><v>#DIV/0!</v></c>" ++
        "<c r=\"C1\" t=\"e\"><v>#REF!</v></c>" ++
        "</row>" ++
        "</sheetData>" ++
        "</worksheet>";
    const dup = try book.allocator.dupe(u8, injected);
    if (book.sheet_data.getEntry(sheet_path)) |entry| {
        book.allocator.free(entry.value_ptr.*);
        entry.value_ptr.* = dup;
    } else {
        book.allocator.free(dup);
        return error.SheetDataMissing;
    }

    var scratch: [4096]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    try runCellsCommand(&w, &book, book.sheets[0], 0, std.testing.allocator, null, null, null, null, null, false, false);

    const out = w.buffered();
    try std.testing.expect(std.mem.indexOf(u8, out, "\"ref\":\"A1\",\"row\":1,\"col\":1,\"t\":\"int\",\"v\":1") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"ref\":\"B1\",\"row\":1,\"col\":2,\"t\":\"error\",\"v\":\"#DIV/0!\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"ref\":\"C1\",\"row\":1,\"col\":3,\"t\":\"error\",\"v\":\"#REF!\"") != null);
}

test "runRowsCommand envelope emits t:error inside cells array (iter61-c)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const tmp_path = try tt.path(std.testing.allocator, io, "cli_iter61c_error_rows.xlsx");
    defer std.testing.allocator.free(tmp_path);
    {
        const writer = xlsx.writer_types;
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("Data");
        try s.writeRow(&.{.{ .integer = 1 }});
        try w.save(io, tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, io, tmp_path);
    defer book.deinit();

    const sheet_path = book.sheets[0].path;
    const injected =
        "<?xml version=\"1.0\"?>" ++
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" ++
        "<sheetData>" ++
        "<row r=\"1\">" ++
        "<c r=\"A1\"><v>1</v></c>" ++
        "<c r=\"B1\" t=\"e\"><v>#DIV/0!</v></c>" ++
        "</row>" ++
        "</sheetData>" ++
        "</worksheet>";
    const dup = try book.allocator.dupe(u8, injected);
    if (book.sheet_data.getEntry(sheet_path)) |entry| {
        book.allocator.free(entry.value_ptr.*);
        entry.value_ptr.* = dup;
    } else {
        book.allocator.free(dup);
        return error.SheetDataMissing;
    }

    var scratch: [4096]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    var pg = Pagination.init(null, null);
    var prologue: bool = false;
    try runRowsOnSheetCore(&w, &book, book.sheets[0], 0, .jsonl, std.testing.allocator, &pg, null, null, null, false, false, false, false, &prologue);

    const out = w.buffered();
    try std.testing.expect(std.mem.indexOf(u8, out, "\"kind\":\"row\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "{\"ref\":\"A1\",\"col\":1,\"t\":\"int\",\"v\":1}") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "{\"ref\":\"B1\",\"col\":2,\"t\":\"error\",\"v\":\"#DIV/0!\"}") != null);
}

// ─── iter61-b: t:"formula" cells end-to-end ─────────────────────────

test "writeCell emits t:formula with formula text + cached value (iter61-b)" {
    var scratch: [512]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    // Stand-alone formula: <c r="C2"><f>A2+B2</f><v>30</v></c>. The
    // CLI emits the design-doc shape: t:"formula", formula:<text>,
    // cached:<value>. The Cell slot carries the cached numeric, the
    // formula side channel carries the expression.
    try writeCell(
        &w,
        "Data",
        0,
        "C2",
        2,
        3,
        .{ .integer = 30 },
        null,
        false,
        false,
        null,
        "A2+B2",
        null,
        false,
    );
    try std.testing.expectEqualStrings(
        "{\"kind\":\"cell\",\"sheet\":\"Data\",\"sheet_idx\":0,\"ref\":\"C2\",\"row\":2,\"col\":3,\"t\":\"formula\",\"formula\":\"A2+B2\",\"cached\":30}\n",
        w.buffered(),
    );
}

test "writeCell emits t:formula with formula_ref + cached value (iter61-b)" {
    var scratch: [512]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    // Shared-formula slave: <c r="C3"><f t="shared" si="0"/><v>47</v></c>.
    // The base cell at C2 already exposed `A2+B2`; this slave just
    // points back at it via formula_ref. The cached integer comes
    // from the slave's <v>.
    try writeCell(
        &w,
        "Data",
        0,
        "C3",
        3,
        3,
        .{ .integer = 47 },
        null,
        false,
        false,
        null,
        null,
        .{ .col = 2, .row = 2 }, // C2
        false,
    );
    try std.testing.expectEqualStrings(
        "{\"kind\":\"cell\",\"sheet\":\"Data\",\"sheet_idx\":0,\"ref\":\"C3\",\"row\":3,\"col\":3,\"t\":\"formula\",\"formula_ref\":\"C2\",\"cached\":47}\n",
        w.buffered(),
    );
}

test "writeRowEnvelope emits formula records inside cells array (iter61-b)" {
    var scratch: [1024]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    const cells = [_]xlsx.Cell{
        .{ .integer = 1 }, // A1 — plain int, not a formula
        .{ .integer = 30 }, // B1 — stand-alone formula, cached 30
        .{ .integer = 47 }, // C1 — shared-formula slave, cached 47
    };
    const dates = [_]bool{ false, false, false };
    const errors = [_]?[]const u8{ null, null, null };
    const formulas = [_]?[]const u8{ null, "A2+B2", null };
    const formula_refs = [_]?xlsx.CellRef{ null, null, .{ .col = 1, .row = 1 } }; // → B1
    try writeRowEnvelope(&w, "Data", 0, 1, &cells, false, null, 0, false, &dates, &errors, &formulas, &formula_refs, false);
    try std.testing.expectEqualStrings(
        "{\"kind\":\"row\",\"sheet\":\"Data\",\"sheet_idx\":0,\"row\":1,\"cells\":[" ++
            "{\"ref\":\"A1\",\"col\":1,\"t\":\"int\",\"v\":1}," ++
            "{\"ref\":\"B1\",\"col\":2,\"t\":\"formula\",\"formula\":\"A2+B2\",\"cached\":30}," ++
            "{\"ref\":\"C1\",\"col\":3,\"t\":\"formula\",\"formula_ref\":\"B1\",\"cached\":47}" ++
            "]}\n",
        w.buffered(),
    );
}

test "writeEnvelopeCells emits formula-only cells when Cell is .empty (iter61-b P2)" {
    // Codex P2 follow-up: a formula cell whose source had no cached
    // <v> element comes through as Cell.empty with formula text in
    // row_formula_strings. The row-envelope path must not skip those
    // records via the .empty short-circuit.
    var scratch: [512]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    const cells = [_]xlsx.Cell{
        .{ .integer = 1 },
        .empty, // formula cell with no <v> body
        .empty, // shared-formula slave with no <v>
    };
    const date_types = [_]bool{ false, false, false };
    const error_strings = [_]?[]const u8{ null, null, null };
    const formula_strings = [_]?[]const u8{ null, "A2+B2", null };
    const formula_refs = [_]?xlsx.CellRef{ null, null, .{ .row = 1, .col = 1 } };

    try writeRowEnvelope(
        &w,
        "Data",
        0,
        7,
        &cells,
        false, // include_blanks
        null, // style_ctx
        0, // col_offset
        false, // compact
        &date_types,
        &error_strings,
        &formula_strings,
        &formula_refs,
        false, // uses_1904
    );

    const out = w.buffered();
    // Cell A (.integer) emits as int.
    try std.testing.expect(std.mem.indexOf(u8, out, "\"t\":\"int\",\"v\":1") != null);
    // Cell B is .empty + has formula → must emit as t:"formula" with NO cached.
    try std.testing.expect(std.mem.indexOf(u8, out, "\"col\":2,\"t\":\"formula\",\"formula\":\"A2+B2\"}") != null);
    // Cell C is .empty + has formula_ref → must emit as t:"formula" with formula_ref.
    // CellRef{row=1, col=1} renders as "B1" (col 0-based → 'B', row 1-based stays 1).
    try std.testing.expect(std.mem.indexOf(u8, out, "\"col\":3,\"t\":\"formula\",\"formula_ref\":\"B1\"}") != null);
}

test "Rows.formulaStrings returns entity-decoded text (iter61-b P2)" {
    // Codex round-2 P2 #2 claim: formula text emitted XML-escaped.
    // Counter-claim verified here: internOrBorrow runs appendDecoded
    // when the body contains '&', so &lt;/&gt;/&amp;/&quot;/&apos;
    // are decoded to their literal characters before reaching the
    // CLI emitter or any Rows.formulaStrings() consumer.
    const xml =
        "<sheetData>" ++
        "<row r=\"1\">" ++
        "<c r=\"A1\"><f>IF(A1&lt;B1,&quot;x&quot;,A1&amp;B1)</f><v>0</v></c>" ++
        "</row>" ++
        "</sheetData>";

    var rows: xlsx.Rows = .{
        .xml = xml,
        .pos = 0,
        .shared_strings = &.{},
        .allocator = std.testing.allocator,
        .row_cells = .empty,
        .row_styles = .empty,
        .row_date_types = .empty,
        .row_error_strings = .empty,
        .row_formula_strings = .empty,
        .row_formula_refs = .empty,
        .shared_si_to_base_ref = .{},
        .array_ranges = .empty,
        .arena = std.heap.ArenaAllocator.init(std.testing.allocator),
    };
    defer rows.deinit();

    _ = (try rows.next()) orelse return error.UnexpectedEndOfRows;
    const formulas = rows.formulaStrings();
    try std.testing.expect(formulas.len >= 1);
    const f = formulas[0] orelse return error.MissingFormula;
    // Decoded: < " & all expanded.
    try std.testing.expectEqualStrings("IF(A1<B1,\"x\",A1&B1)", f);
}

test "runCellsCommand emits t:formula for stand-alone, shared-base, shared-slave (iter61-b)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    // Mirror the iter60c / iter61-c post-injection trick: write a valid
    // workbook, then replace sheet1.xml with a hand-crafted blob that
    // exercises all three formula shapes in one row.
    const tmp_path = try tt.path(std.testing.allocator, io, "cli_iter61b_formula.xlsx");
    defer std.testing.allocator.free(tmp_path);
    {
        const writer = xlsx.writer_types;
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("Data");
        try s.writeRow(&.{.{ .integer = 1 }});
        try w.save(io, tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, io, tmp_path);
    defer book.deinit();

    const sheet_path = book.sheets[0].path;
    // Row 1: A=1, B=2 (data inputs).
    // Row 2: C2 stand-alone formula (=A2+B2 → 3); D2 shared-base
    //        formula (=A2*B2 → 2) with si=0 ref="D2:D3"; E2 array
    //        formula (=A2-B2 → -1) ref="E2:E3" — surfaces top-left
    //        cell as a formula record.
    // Row 3: A=10, B=20.
    //        C3 stand-alone (=A3+B3 → 30); D3 shared-formula slave
    //        (si=0, no body) → resolves to D2; E3 has no <c> for it
    //        (we elide because the array-spread for E2 doesn't surface).
    const injected =
        "<?xml version=\"1.0\"?>" ++
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" ++
        "<sheetData>" ++
        "<row r=\"1\">" ++
        "<c r=\"A1\"><v>1</v></c>" ++
        "<c r=\"B1\"><v>2</v></c>" ++
        "</row>" ++
        "<row r=\"2\">" ++
        "<c r=\"A2\"><v>1</v></c>" ++
        "<c r=\"B2\"><v>2</v></c>" ++
        "<c r=\"C2\"><f>A2+B2</f><v>3</v></c>" ++
        "<c r=\"D2\"><f t=\"shared\" ref=\"D2:D3\" si=\"0\">A2*B2</f><v>2</v></c>" ++
        "<c r=\"E2\"><f t=\"array\" ref=\"E2:E3\">A2-B2</f><v>-1</v></c>" ++
        "</row>" ++
        "<row r=\"3\">" ++
        "<c r=\"A3\"><v>10</v></c>" ++
        "<c r=\"B3\"><v>20</v></c>" ++
        "<c r=\"C3\"><f>A3+B3</f><v>30</v></c>" ++
        "<c r=\"D3\"><f t=\"shared\" si=\"0\"/><v>200</v></c>" ++
        "</row>" ++
        "</sheetData>" ++
        "</worksheet>";
    const dup = try book.allocator.dupe(u8, injected);
    if (book.sheet_data.getEntry(sheet_path)) |entry| {
        book.allocator.free(entry.value_ptr.*);
        entry.value_ptr.* = dup;
    } else {
        book.allocator.free(dup);
        return error.SheetDataMissing;
    }

    var scratch: [8192]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    try runCellsCommand(&w, &book, book.sheets[0], 0, std.testing.allocator, null, null, null, null, null, false, false);

    const out = w.buffered();
    // Row 1 — plain integers, no formula records.
    try std.testing.expect(std.mem.indexOf(u8, out, "\"ref\":\"A1\",\"row\":1,\"col\":1,\"t\":\"int\",\"v\":1") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"ref\":\"B1\",\"row\":1,\"col\":2,\"t\":\"int\",\"v\":2") != null);
    // Row 2 — three formula records:
    //   C2: stand-alone — formula text + cached.
    try std.testing.expect(std.mem.indexOf(u8, out, "\"ref\":\"C2\",\"row\":2,\"col\":3,\"t\":\"formula\",\"formula\":\"A2+B2\",\"cached\":3") != null);
    //   D2: shared-formula base — formula text + cached.
    try std.testing.expect(std.mem.indexOf(u8, out, "\"ref\":\"D2\",\"row\":2,\"col\":4,\"t\":\"formula\",\"formula\":\"A2*B2\",\"cached\":2") != null);
    //   E2: array-formula base — formula text + cached (top-left only).
    try std.testing.expect(std.mem.indexOf(u8, out, "\"ref\":\"E2\",\"row\":2,\"col\":5,\"t\":\"formula\",\"formula\":\"A2-B2\",\"cached\":-1") != null);
    // Row 3 — C3 stand-alone, D3 shared-formula slave (formula_ref=D2):
    try std.testing.expect(std.mem.indexOf(u8, out, "\"ref\":\"C3\",\"row\":3,\"col\":3,\"t\":\"formula\",\"formula\":\"A3+B3\",\"cached\":30") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"ref\":\"D3\",\"row\":3,\"col\":4,\"t\":\"formula\",\"formula_ref\":\"D2\",\"cached\":200") != null);
}

test "runCellsCommand emits t:formula with formula_ref for array-formula slaves" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    // Same post-injection trick: write a workbook then replace the
    // sheet body. C2 is the array base (=A2:A4*B2:B4 with ref=C2:C4);
    // C3 + C4 carry only cached <v> bodies. Reader must spread the
    // array formula context to the slaves so all three rows surface
    // as t:"formula" records.
    const tmp_path = try tt.path(std.testing.allocator, io, "cli_array_spread.xlsx");
    defer std.testing.allocator.free(tmp_path);
    {
        const writer = xlsx.writer_types;
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s = try w.addSheet("Data");
        try s.writeRow(&.{.{ .integer = 1 }});
        try w.save(io, tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, io, tmp_path);
    defer book.deinit();

    const sheet_path = book.sheets[0].path;
    const injected =
        "<?xml version=\"1.0\"?>" ++
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" ++
        "<sheetData>" ++
        "<row r=\"2\">" ++
        "<c r=\"C2\"><f t=\"array\" ref=\"C2:C4\">A2:A4*B2:B4</f><v>10</v></c>" ++
        "</row>" ++
        "<row r=\"3\">" ++
        "<c r=\"C3\"><v>20</v></c>" ++
        "</row>" ++
        "<row r=\"4\">" ++
        "<c r=\"C4\"><v>30</v></c>" ++
        "</row>" ++
        "</sheetData>" ++
        "</worksheet>";
    const dup = try book.allocator.dupe(u8, injected);
    if (book.sheet_data.getEntry(sheet_path)) |entry| {
        book.allocator.free(entry.value_ptr.*);
        entry.value_ptr.* = dup;
    } else {
        book.allocator.free(dup);
        return error.SheetDataMissing;
    }

    var scratch: [8192]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    try runCellsCommand(&w, &book, book.sheets[0], 0, std.testing.allocator, null, null, null, null, null, false, false);

    const out = w.buffered();
    // Base C2: own formula text + cached.
    try std.testing.expect(std.mem.indexOf(u8, out, "\"ref\":\"C2\",\"row\":2,\"col\":3,\"t\":\"formula\",\"formula\":\"A2:A4*B2:B4\",\"cached\":10") != null);
    // Slave C3: formula_ref to C2 + cached.
    try std.testing.expect(std.mem.indexOf(u8, out, "\"ref\":\"C3\",\"row\":3,\"col\":3,\"t\":\"formula\",\"formula_ref\":\"C2\",\"cached\":20") != null);
    // Slave C4: formula_ref to C2 + cached.
    try std.testing.expect(std.mem.indexOf(u8, out, "\"ref\":\"C4\",\"row\":4,\"col\":3,\"t\":\"formula\",\"formula_ref\":\"C2\",\"cached\":30") != null);
}

// ─── S6: `pivots` ────────────────────────────────────────────────────

test "runPivotsCommand: one record per pivot in the frozen field order" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(std.testing.allocator, io, "cli_pivots.xlsx");
    defer std.testing.allocator.free(path);
    try zlsx_pkg.pivots.fixture.write(std.testing.allocator, io, path, .sheet_ref);

    var scratch: [8192]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    var err_buf: [256]u8 = undefined;
    var err_w = std.Io.Writer.fixed(&err_buf);
    const rc = try runPivotsCommand(std.testing.allocator, io, .{ .file = path, .subcommand = .pivots }, &w, &err_w);
    try std.testing.expectEqual(@as(u8, 0), rc);
    const expected =
        "{\"kind\":\"pivot\",\"sheet\":\"Report\",\"sheet_idx\":1,\"name\":\"PivotTable1\"," ++
        "\"part\":\"xl/pivotTables/pivotTable1.xml\"," ++
        "\"location\":{\"ref\":\"A3:B6\",\"first_header_row\":1,\"first_data_row\":1,\"first_data_col\":1}," ++
        "\"rows\":[{\"field\":\"Region\",\"idx\":0}],\"cols\":[],\"pages\":[]," ++
        "\"values\":[{\"name\":\"Sum of Qty\",\"field\":\"Qty\",\"idx\":1,\"subtotal\":\"sum\",\"show_data_as\":null,\"num_fmt_id\":null}]," ++
        "\"data_caption\":\"Values\",\"grand_totals\":{\"rows\":true,\"cols\":true},\"style\":\"PivotStyleLight16\"," ++
        "\"cache\":{\"id\":7,\"part\":\"xl/pivotCache/pivotCacheDefinition1.xml\"," ++
        "\"records_part\":\"xl/pivotCache/pivotCacheRecords1.xml\",\"record_count\":3," ++
        "\"refreshed_by\":\"zlsx\",\"refreshed_date\":\"45000.5\",\"refresh_on_load\":false,\"save_data\":true," ++
        "\"source\":{\"type\":\"worksheet\",\"sheet\":\"Data\",\"ref\":\"A1:C4\",\"name\":null," ++
        "\"resolved\":{\"sheet\":\"Data\",\"sheet_idx\":0,\"via\":\"sheet_attr\",\"bounds\":\"A1:C4\"},\"unresolved\":null}," ++
        "\"fields\":[" ++
        "{\"name\":\"Region\",\"num_fmt_id\":0,\"formula\":null,\"items\":2,\"types\":[\"string\"],\"min\":null,\"max\":null}," ++
        "{\"name\":\"Qty\",\"num_fmt_id\":0,\"formula\":null,\"items\":null,\"types\":[\"number\",\"integer\"],\"min\":\"3\",\"max\":\"5\"}," ++
        "{\"name\":\"Price\",\"num_fmt_id\":0,\"formula\":null,\"items\":null,\"types\":[\"number\"],\"min\":\"1.5\",\"max\":\"3.5\"}" ++
        "]}}\n";
    try std.testing.expectEqualStrings(expected, w.buffered());
}

test "runPivotsCommand: table-name, defined-name, external and dangling sources in the resolved slot" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();

    const cases = [_]struct { kind: zlsx_pkg.pivots.fixture.SourceKind, want: []const u8 }{
        .{ .kind = .table_name, .want = "\"source\":{\"type\":\"worksheet\",\"sheet\":null,\"ref\":null,\"name\":\"SalesTbl\",\"resolved\":{\"sheet\":\"Data\",\"sheet_idx\":0,\"via\":\"table\",\"bounds\":\"A1:C4\"},\"unresolved\":null}" },
        .{ .kind = .defined_name, .want = "\"source\":{\"type\":\"worksheet\",\"sheet\":null,\"ref\":null,\"name\":\"PivotSrc\",\"resolved\":{\"sheet\":\"Data\",\"sheet_idx\":0,\"via\":\"defined_name\",\"bounds\":\"A1:C4\"},\"unresolved\":null}" },
        .{ .kind = .external, .want = "\"source\":{\"type\":\"worksheet\",\"sheet\":\"Sheet1\",\"ref\":\"A1:C4\",\"name\":null,\"resolved\":{\"external\":\"file:///C:/data/other.xlsx\"},\"unresolved\":null}" },
        .{ .kind = .dangling, .want = "\"source\":{\"type\":\"worksheet\",\"sheet\":\"Nope\",\"ref\":\"A1:C4\",\"name\":null,\"resolved\":null,\"unresolved\":{\"why\":\"dangling_sheet\",\"sheets\":[]}}" },
        .{ .kind = .consolidation, .want = "\"source\":{\"type\":\"consolidation\",\"range_sets\":[{\"sheet\":\"Data\",\"ref\":\"A1:C4\",\"name\":null,\"resolved\":{\"sheet\":\"Data\",\"sheet_idx\":0,\"via\":\"sheet_attr\",\"bounds\":\"A1:C4\"},\"unresolved\":null},{\"sheet\":null,\"ref\":null,\"name\":\"PivotSrc\",\"resolved\":{\"sheet\":\"Report\",\"sheet_idx\":1,\"via\":\"defined_name\",\"bounds\":\"A1:B2\"},\"unresolved\":null}]}" },
    };
    for (cases, 0..) |case, i| {
        var name_buf: [32]u8 = undefined;
        const name = try std.fmt.bufPrint(&name_buf, "cli_pivots_{d}.xlsx", .{i});
        const path = try tt.path(std.testing.allocator, io, name);
        defer std.testing.allocator.free(path);
        try zlsx_pkg.pivots.fixture.write(std.testing.allocator, io, path, case.kind);

        var scratch: [8192]u8 = undefined;
        var w = std.Io.Writer.fixed(&scratch);
        var err_buf: [256]u8 = undefined;
        var err_w = std.Io.Writer.fixed(&err_buf);
        const rc = try runPivotsCommand(std.testing.allocator, io, .{ .file = path, .subcommand = .pivots }, &w, &err_w);
        try std.testing.expectEqual(@as(u8, 0), rc);
        const line = w.buffered();
        try std.testing.expect(std.mem.indexOf(u8, line, case.want) != null);
        try std.testing.expectEqual(@as(usize, 1), std.mem.count(u8, line, "\n"));
    }
}

test "runPivotsCommand: an unbounded name reports its provenance; whole columns report as bounds (S7b-1)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(std.testing.allocator, io, "cli_pivots_provenance.xlsx");
    defer std.testing.allocator.free(path);

    const cases = [_]struct { body: []const u8, want: []const u8 }{
        .{ .body = "OFFSET(Report!$D$1,0,0,4,3)", .want = "\"resolved\":null,\"unresolved\":{\"why\":\"unbounded_body\",\"sheets\":[{\"sheet\":\"Report\",\"sheet_idx\":1}]}}" },
        .{ .body = "Data!$A:$C", .want = "\"resolved\":{\"sheet\":\"Data\",\"sheet_idx\":0,\"via\":\"defined_name\",\"bounds\":\"A:C\"},\"unresolved\":null}" },
    };
    for (cases) |case| {
        try zlsx_pkg.pivots.fixture.write(std.testing.allocator, io, path, .defined_name);
        try zlsx_pkg.pivots.fixture.patchPart(std.testing.allocator, io, path, "xl/workbook.xml", "Data!$A$1:$C$4", case.body);
        var scratch: [8192]u8 = undefined;
        var w = std.Io.Writer.fixed(&scratch);
        var err_buf: [256]u8 = undefined;
        var err_w = std.Io.Writer.fixed(&err_buf);
        const rc = try runPivotsCommand(std.testing.allocator, io, .{ .file = path, .subcommand = .pivots }, &w, &err_w);
        try std.testing.expectEqual(@as(u8, 0), rc);
        try std.testing.expect(std.mem.indexOf(u8, w.buffered(), case.want) != null);
    }
}

test "runPivotsCommand: --sheet / --name select the host sheet; compact-ndjson drops the envelope" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(std.testing.allocator, io, "cli_pivots_filter.xlsx");
    defer std.testing.allocator.free(path);
    try zlsx_pkg.pivots.fixture.write(std.testing.allocator, io, path, .sheet_ref);

    // The source sheet hosts nothing: an empty, successful stream.
    {
        var scratch: [8192]u8 = undefined;
        var w = std.Io.Writer.fixed(&scratch);
        var err_buf: [256]u8 = undefined;
        var err_w = std.Io.Writer.fixed(&err_buf);
        const rc = try runPivotsCommand(std.testing.allocator, io, .{ .file = path, .subcommand = .pivots, .sheet_index = 0 }, &w, &err_w);
        try std.testing.expectEqual(@as(u8, 0), rc);
        try std.testing.expectEqualStrings("", w.buffered());
    }
    // By name, compact: prologue + envelope-less record.
    {
        var scratch: [8192]u8 = undefined;
        var w = std.Io.Writer.fixed(&scratch);
        var err_buf: [256]u8 = undefined;
        var err_w = std.Io.Writer.fixed(&err_buf);
        const rc = try runPivotsCommand(std.testing.allocator, io, .{ .file = path, .subcommand = .pivots, .sheet_name = "Report", .output = .compact_ndjson }, &w, &err_w);
        try std.testing.expectEqual(@as(u8, 0), rc);
        const got = w.buffered();
        try std.testing.expect(std.mem.startsWith(u8, got, "{\"kind\":\"sheet\",\"sheet\":\"Report\",\"sheet_idx\":1}\n{\"kind\":\"pivot\",\"name\":\"PivotTable1\","));
    }
    // Unknown sheet: exit 3, like every read sub-command.
    {
        var scratch: [8192]u8 = undefined;
        var w = std.Io.Writer.fixed(&scratch);
        var err_buf: [256]u8 = undefined;
        var err_w = std.Io.Writer.fixed(&err_buf);
        const rc = try runPivotsCommand(std.testing.allocator, io, .{ .file = path, .subcommand = .pivots, .sheet_name = "Nope" }, &w, &err_w);
        try std.testing.expectEqual(@as(u8, 3), rc);
    }
    // A workbook without pivots is an empty, successful stream.
    {
        const plain = try tt.path(std.testing.allocator, io, "cli_pivots_none.xlsx");
        defer std.testing.allocator.free(plain);
        {
            var wr = xlsx.writer_types.Writer.init(std.testing.allocator);
            defer wr.deinit();
            var s = try wr.addSheet("S");
            try s.writeRow(&.{.{ .integer = 1 }});
            try wr.save(io, plain);
        }
        var scratch: [1024]u8 = undefined;
        var w = std.Io.Writer.fixed(&scratch);
        var err_buf: [256]u8 = undefined;
        var err_w = std.Io.Writer.fixed(&err_buf);
        const rc = try runPivotsCommand(std.testing.allocator, io, .{ .file = plain, .subcommand = .pivots }, &w, &err_w);
        try std.testing.expectEqual(@as(u8, 0), rc);
        try std.testing.expectEqualStrings("", w.buffered());
    }
}

test "runPivotsCommand: orphan caches follow the pivots, only when no sheet is selected, and page through" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(std.testing.allocator, io, "cli_pivots_orphan.xlsx");
    defer std.testing.allocator.free(path);
    try zlsx_pkg.pivots.fixture.writeWithOrphanCache(std.testing.allocator, io, path, .sheet_ref);

    const orphan_line =
        "{\"kind\":\"pivot_cache\",\"cache\":{\"id\":8,\"part\":\"xl/pivotCache/pivotCacheDefinition2.xml\"," ++
        "\"records_part\":\"xl/pivotCache/pivotCacheRecords2.xml\",\"record_count\":0,\"refreshed_by\":null," ++
        "\"refreshed_date\":null,\"refresh_on_load\":false,\"save_data\":false," ++
        "\"source\":{\"type\":\"worksheet\",\"sheet\":\"Report\",\"ref\":\"A1:A1\",\"name\":null," ++
        "\"resolved\":{\"sheet\":\"Report\",\"sheet_idx\":1,\"via\":\"sheet_attr\",\"bounds\":\"A1:A1\"},\"unresolved\":null}," ++
        "\"fields\":[{\"name\":\"Note\",\"num_fmt_id\":0,\"formula\":null,\"items\":null,\"types\":[\"string\"],\"min\":null,\"max\":null}]}}\n";

    // No selection: the pivot, then the orphan.
    {
        var scratch: [8192]u8 = undefined;
        var w = std.Io.Writer.fixed(&scratch);
        var err_buf: [256]u8 = undefined;
        var err_w = std.Io.Writer.fixed(&err_buf);
        const rc = try runPivotsCommand(std.testing.allocator, io, .{ .file = path, .subcommand = .pivots }, &w, &err_w);
        try std.testing.expectEqual(@as(u8, 0), rc);
        const got = w.buffered();
        try std.testing.expectEqual(@as(usize, 2), std.mem.count(u8, got, "\n"));
        try std.testing.expect(std.mem.startsWith(u8, got, "{\"kind\":\"pivot\","));
        const second = std.mem.indexOf(u8, got, "\n").? + 1;
        try std.testing.expectEqualStrings(orphan_line, got[second..]);
    }
    // A sheet selected: the orphan belongs to no sheet and is not emitted.
    {
        var scratch: [8192]u8 = undefined;
        var w = std.Io.Writer.fixed(&scratch);
        var err_buf: [256]u8 = undefined;
        var err_w = std.Io.Writer.fixed(&err_buf);
        const rc = try runPivotsCommand(std.testing.allocator, io, .{ .file = path, .subcommand = .pivots, .sheet_index = 1 }, &w, &err_w);
        try std.testing.expectEqual(@as(u8, 0), rc);
        try std.testing.expectEqual(@as(usize, 1), std.mem.count(u8, w.buffered(), "\n"));
        try std.testing.expect(std.mem.indexOf(u8, w.buffered(), "pivot_cache") == null);
    }
    // --all-sheets is not a selection; --skip 1 pages past the pivot to the orphan.
    {
        var scratch: [8192]u8 = undefined;
        var w = std.Io.Writer.fixed(&scratch);
        var err_buf: [256]u8 = undefined;
        var err_w = std.Io.Writer.fixed(&err_buf);
        const rc = try runPivotsCommand(std.testing.allocator, io, .{ .file = path, .subcommand = .pivots, .all_sheets = true, .skip = 1 }, &w, &err_w);
        try std.testing.expectEqual(@as(u8, 0), rc);
        try std.testing.expectEqualStrings(orphan_line, w.buffered());
    }
    // --take 1 stops before the orphan.
    {
        var scratch: [8192]u8 = undefined;
        var w = std.Io.Writer.fixed(&scratch);
        var err_buf: [256]u8 = undefined;
        var err_w = std.Io.Writer.fixed(&err_buf);
        const rc = try runPivotsCommand(std.testing.allocator, io, .{ .file = path, .subcommand = .pivots, .take = 1 }, &w, &err_w);
        try std.testing.expectEqual(@as(u8, 0), rc);
        try std.testing.expectEqual(@as(usize, 1), std.mem.count(u8, w.buffered(), "\n"));
        try std.testing.expect(std.mem.indexOf(u8, w.buffered(), "pivot_cache") == null);
    }
}

test "runPivotsCommand: --sheet-glob suppresses orphans, an ISO-only refresh date shows, an unreadable part is exit 2 with no stdout" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(std.testing.allocator, io, "cli_pivots_glob.xlsx");
    defer std.testing.allocator.free(path);
    try zlsx_pkg.pivots.fixture.writeWithOrphanCache(std.testing.allocator, io, path, .sheet_ref);
    try zlsx_pkg.pivots.fixture.patchPart(std.testing.allocator, io, path, "xl/pivotCache/pivotCacheDefinition1.xml", "refreshedDate=\"45000.5\"", "refreshedDateIso=\"2023-03-15T12:00:00Z\"");

    // A glob is a selection: the host matches, the orphan is not emitted.
    {
        var scratch: [8192]u8 = undefined;
        var w = std.Io.Writer.fixed(&scratch);
        var err_buf: [256]u8 = undefined;
        var err_w = std.Io.Writer.fixed(&err_buf);
        const rc = try runPivotsCommand(std.testing.allocator, io, .{ .file = path, .subcommand = .pivots, .sheet_glob = "Rep*" }, &w, &err_w);
        try std.testing.expectEqual(@as(u8, 0), rc);
        const got = w.buffered();
        try std.testing.expectEqual(@as(usize, 1), std.mem.count(u8, got, "\n"));
        try std.testing.expect(std.mem.indexOf(u8, got, "pivot_cache") == null);
        try std.testing.expect(std.mem.indexOf(u8, got, "\"refreshed_date\":\"2023-03-15T12:00:00Z\"") != null);
    }
    // A glob matching no host: an empty, successful stream.
    {
        var scratch: [1024]u8 = undefined;
        var w = std.Io.Writer.fixed(&scratch);
        var err_buf: [256]u8 = undefined;
        var err_w = std.Io.Writer.fixed(&err_buf);
        const rc = try runPivotsCommand(std.testing.allocator, io, .{ .file = path, .subcommand = .pivots, .sheet_glob = "Nope*" }, &w, &err_w);
        try std.testing.expectEqual(@as(u8, 0), rc);
        try std.testing.expectEqualStrings("", w.buffered());
    }
    // A pivot part that cannot be read: exit 2, a diagnostic, nothing on stdout.
    try zlsx_pkg.pivots.fixture.patchPart(std.testing.allocator, io, path, "xl/pivotTables/pivotTable1.xml", "<location ref=\"A3:B6\" firstHeaderRow=\"1\" firstDataRow=\"1\" firstDataCol=\"1\"/>", "");
    {
        var scratch: [1024]u8 = undefined;
        var w = std.Io.Writer.fixed(&scratch);
        var err_buf: [256]u8 = undefined;
        var err_w = std.Io.Writer.fixed(&err_buf);
        const rc = try runPivotsCommand(std.testing.allocator, io, .{ .file = path, .subcommand = .pivots }, &w, &err_w);
        try std.testing.expectEqual(@as(u8, 2), rc);
        try std.testing.expectEqualStrings("", w.buffered());
        try std.testing.expect(std.mem.indexOf(u8, err_w.buffered(), "MalformedPivotXml") != null);
    }
}

test "parseArgs routes 'pivots' and rejects row-keyed flags on it" {
    {
        const argv = [_][]const u8{ "pivots", "f.xlsx" };
        const a = try parseArgs(&argv);
        try std.testing.expectEqual(Subcommand.pivots, a.subcommand);
        try std.testing.expectEqualStrings("f.xlsx", a.file);
    }
    {
        // The legacy override parses alongside the sub-command; the
        // dispatch honours it before the pivot path.
        const argv = [_][]const u8{ "pivots", "f.xlsx", "--list-sheets" };
        const a = try parseArgs(&argv);
        try std.testing.expectEqual(Subcommand.pivots, a.subcommand);
        try std.testing.expect(a.list_sheets);
    }
    {
        const argv = [_][]const u8{ "pivots", "f.xlsx", "--range", "A1:B2" };
        try std.testing.expectError(ArgError.BadArgValue, parseArgs(&argv));
    }
}

// ─── S3b: `merges` + `defined-names` ─────────────────────────────────

test "parseArgs: merges / defined-names tokens and flag rejections" {
    {
        const argv = [_][]const u8{ "merges", "f.xlsx" };
        const a = try parseArgs(&argv);
        try std.testing.expectEqual(Subcommand.merges, a.subcommand);
        try std.testing.expectEqualStrings("f.xlsx", a.file);
    }
    {
        const argv = [_][]const u8{ "defined-names", "f.xlsx" };
        const a = try parseArgs(&argv);
        try std.testing.expectEqual(Subcommand.defined_names, a.subcommand);
        try std.testing.expectEqualStrings("f.xlsx", a.file);
    }
    // Range-keyed / workbook-keyed records: no row or rectangle key.
    inline for (.{ "merges", "defined-names" }) |cmd| {
        {
            const argv = [_][]const u8{ cmd, "f.xlsx", "--range", "A1:B2" };
            try std.testing.expectError(ArgError.BadArgValue, parseArgs(&argv));
        }
        {
            const argv = [_][]const u8{ cmd, "f.xlsx", "--start-row", "2" };
            try std.testing.expectError(ArgError.BadArgValue, parseArgs(&argv));
        }
        {
            const argv = [_][]const u8{ cmd, "f.xlsx", "--end-row", "5" };
            try std.testing.expectError(ArgError.BadArgValue, parseArgs(&argv));
        }
        {
            const argv = [_][]const u8{ cmd, "f.xlsx", "--output", "pretty-json" };
            try std.testing.expectError(ArgError.BadArgValue, parseArgs(&argv));
        }
        {
            // `--format` shapes only `rows` output; non-workbook-scoped
            // sub-commands reject anything but the implicit default.
            const argv = [_][]const u8{ cmd, "f.xlsx", "--format", "csv" };
            try std.testing.expectError(ArgError.BadFormat, parseArgs(&argv));
        }
    }
}

test "runMergesCommand: every sheet by default, exact wire shape" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(std.testing.allocator, io, "cli_merges.xlsx");
    defer std.testing.allocator.free(path);
    {
        const writer = xlsx.writer_types;
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s0 = try w.addSheet("Data");
        try s0.writeRow(&.{.{ .string = "a" }});
        try s0.addMergedCell("A1:B3");
        try s0.addMergedCell("D5:D6");
        var s1 = try w.addSheet("Report");
        try s1.writeRow(&.{.{ .string = "b" }});
        try s1.addMergedCell("C2:E2");
        try w.save(io, path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, io, path);
    defer book.deinit();

    // Default: every sheet, full envelope, corners 1-based.
    {
        var scratch: [4096]u8 = undefined;
        var w = std.Io.Writer.fixed(&scratch);
        var err_buf: [256]u8 = undefined;
        var err_w = std.Io.Writer.fixed(&err_buf);
        try std.testing.expectEqual(@as(u8, 0), try runMergesCommand(&w, &err_w, &book, null, .{ .file = path, .subcommand = .merges }, null, null));
        try std.testing.expectEqualStrings(
            "{\"kind\":\"merge\",\"sheet\":\"Data\",\"sheet_idx\":0,\"range\":\"A1:B3\",\"start_row\":1,\"start_col\":1,\"end_row\":3,\"end_col\":2}\n" ++
                "{\"kind\":\"merge\",\"sheet\":\"Data\",\"sheet_idx\":0,\"range\":\"D5:D6\",\"start_row\":5,\"start_col\":4,\"end_row\":6,\"end_col\":4}\n" ++
                "{\"kind\":\"merge\",\"sheet\":\"Report\",\"sheet_idx\":1,\"range\":\"C2:E2\",\"start_row\":2,\"start_col\":3,\"end_row\":2,\"end_col\":5}\n",
            w.buffered(),
        );
    }
    // --sheet 1 narrows to the Report sheet.
    {
        var scratch: [1024]u8 = undefined;
        var w = std.Io.Writer.fixed(&scratch);
        var err_buf: [256]u8 = undefined;
        var err_w = std.Io.Writer.fixed(&err_buf);
        try std.testing.expectEqual(@as(u8, 0), try runMergesCommand(&w, &err_w, &book, 1, .{ .file = path, .subcommand = .merges }, null, null));
        try std.testing.expectEqualStrings(
            "{\"kind\":\"merge\",\"sheet\":\"Report\",\"sheet_idx\":1,\"range\":\"C2:E2\",\"start_row\":2,\"start_col\":3,\"end_row\":2,\"end_col\":5}\n",
            w.buffered(),
        );
    }
    // compact-ndjson: one prologue per sheet, records drop the envelope.
    {
        var scratch: [4096]u8 = undefined;
        var w = std.Io.Writer.fixed(&scratch);
        var err_buf: [256]u8 = undefined;
        var err_w = std.Io.Writer.fixed(&err_buf);
        try std.testing.expectEqual(@as(u8, 0), try runMergesCommand(&w, &err_w, &book, null, .{ .file = path, .subcommand = .merges, .output = .compact_ndjson }, null, null));
        try std.testing.expectEqualStrings(
            "{\"kind\":\"sheet\",\"sheet\":\"Data\",\"sheet_idx\":0}\n" ++
                "{\"kind\":\"merge\",\"range\":\"A1:B3\",\"start_row\":1,\"start_col\":1,\"end_row\":3,\"end_col\":2}\n" ++
                "{\"kind\":\"merge\",\"range\":\"D5:D6\",\"start_row\":5,\"start_col\":4,\"end_row\":6,\"end_col\":4}\n" ++
                "{\"kind\":\"sheet\",\"sheet\":\"Report\",\"sheet_idx\":1}\n" ++
                "{\"kind\":\"merge\",\"range\":\"C2:E2\",\"start_row\":2,\"start_col\":3,\"end_row\":2,\"end_col\":5}\n",
            w.buffered(),
        );
    }
    // --skip / --take page the concatenated cross-sheet stream.
    {
        var scratch: [1024]u8 = undefined;
        var w = std.Io.Writer.fixed(&scratch);
        var err_buf: [256]u8 = undefined;
        var err_w = std.Io.Writer.fixed(&err_buf);
        try std.testing.expectEqual(@as(u8, 0), try runMergesCommand(&w, &err_w, &book, null, .{ .file = path, .subcommand = .merges }, 1, 1));
        try std.testing.expectEqualStrings(
            "{\"kind\":\"merge\",\"sheet\":\"Data\",\"sheet_idx\":0,\"range\":\"D5:D6\",\"start_row\":5,\"start_col\":4,\"end_row\":6,\"end_col\":4}\n",
            w.buffered(),
        );
    }
    // --sheet-glob narrows by sheet name.
    {
        var scratch: [1024]u8 = undefined;
        var w = std.Io.Writer.fixed(&scratch);
        var err_buf: [256]u8 = undefined;
        var err_w = std.Io.Writer.fixed(&err_buf);
        try std.testing.expectEqual(@as(u8, 0), try runMergesCommand(&w, &err_w, &book, null, .{ .file = path, .subcommand = .merges, .sheet_glob = "Rep*" }, null, null));
        try std.testing.expectEqualStrings(
            "{\"kind\":\"merge\",\"sheet\":\"Report\",\"sheet_idx\":1,\"range\":\"C2:E2\",\"start_row\":2,\"start_col\":3,\"end_row\":2,\"end_col\":5}\n",
            w.buffered(),
        );
    }
}

test "runMergesCommand: a non-UTF-8 sheet name refuses before any record (Codex #211 r4)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(std.testing.allocator, io, "cli_merges_bad_name.xlsx");
    defer std.testing.allocator.free(path);
    {
        const writer = xlsx.writer_types;
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s0 = try w.addSheet("Bad");
        try s0.writeRow(&.{.{ .string = "a" }});
        try s0.addMergedCell("A1:B2");
        var s1 = try w.addSheet("Fine");
        try s1.writeRow(&.{.{ .string = "b" }});
        try s1.addMergedCell("C3:D4");
        try w.save(io, path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, io, path);
    defer book.deinit();
    // Post-injection: the writer refuses invalid names, so poison the
    // parsed inventory directly — the merge map is keyed by part path,
    // not by name, so the lookup still resolves.
    book.sheets[0].name = "B\xffd";

    // Both output modes refuse with exit 2 and an empty stream.
    inline for (.{ OutputMode.ndjson, OutputMode.compact_ndjson }) |mode| {
        var scratch: [1024]u8 = undefined;
        var w = std.Io.Writer.fixed(&scratch);
        var err_buf: [256]u8 = undefined;
        var err_w = std.Io.Writer.fixed(&err_buf);
        const rc = try runMergesCommand(&w, &err_w, &book, null, .{ .file = path, .subcommand = .merges, .output = mode }, null, null);
        try std.testing.expectEqual(@as(u8, 2), rc);
        try std.testing.expectEqualStrings("", w.buffered());
        try std.testing.expect(std.mem.indexOf(u8, err_w.buffered(), "non-UTF-8") != null);
    }
    // Narrowed away from the bad sheet, the stream is clean and the
    // bad name is never emitted — no reason to refuse.
    {
        var scratch: [1024]u8 = undefined;
        var w = std.Io.Writer.fixed(&scratch);
        var err_buf: [256]u8 = undefined;
        var err_w = std.Io.Writer.fixed(&err_buf);
        const rc = try runMergesCommand(&w, &err_w, &book, 1, .{ .file = path, .subcommand = .merges }, null, null);
        try std.testing.expectEqual(@as(u8, 0), rc);
        try std.testing.expectEqualStrings(
            "{\"kind\":\"merge\",\"sheet\":\"Fine\",\"sheet_idx\":1,\"range\":\"C3:D4\",\"start_row\":3,\"start_col\":3,\"end_row\":4,\"end_col\":4}\n",
            w.buffered(),
        );
    }
}

test "runDefinedNamesCommand: document order, scope narrowing, exit codes" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(std.testing.allocator, io, "cli_defined_names.xlsx");
    defer std.testing.allocator.free(path);
    {
        const writer = xlsx.writer_types;
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s0 = try w.addSheet("Data");
        try s0.writeRow(&.{.{ .integer = 1 }});
        var s1 = try w.addSheet("Report");
        try s1.writeRow(&.{.{ .integer = 2 }});
        try w.addDefinedName("Prices", "Data!$A$1:$C$4", .{});
        try w.addDefinedName("_xlnm.Print_Area", "Report!$A$1:$B$9", .{ .local_sheet_id = 1 });
        try w.addDefinedName("Secret", "Data!$Z$1", .{ .hidden = true });
        try w.save(io, path);
    }

    // Default: every name, document order, exact wire shape.
    {
        var scratch: [4096]u8 = undefined;
        var w = std.Io.Writer.fixed(&scratch);
        var err_buf: [256]u8 = undefined;
        var err_w = std.Io.Writer.fixed(&err_buf);
        const rc = try runDefinedNamesCommand(std.testing.allocator, io, .{ .file = path, .subcommand = .defined_names }, &w, &err_w);
        try std.testing.expectEqual(@as(u8, 0), rc);
        try std.testing.expectEqualStrings(
            "{\"kind\":\"defined_name\",\"name\":\"Prices\",\"scope\":\"workbook\",\"sheet\":null,\"sheet_idx\":null,\"body\":\"Data!$A$1:$C$4\",\"hidden\":false}\n" ++
                "{\"kind\":\"defined_name\",\"name\":\"_xlnm.Print_Area\",\"scope\":\"sheet\",\"sheet\":\"Report\",\"sheet_idx\":1,\"body\":\"Report!$A$1:$B$9\",\"hidden\":false}\n" ++
                "{\"kind\":\"defined_name\",\"name\":\"Secret\",\"scope\":\"workbook\",\"sheet\":null,\"sheet_idx\":null,\"body\":\"Data!$Z$1\",\"hidden\":true}\n",
            w.buffered(),
        );
    }
    // --sheet 1: only the names SCOPED to that sheet; workbook-scope
    // names are suppressed like orphan caches under a pivots selector.
    {
        var scratch: [1024]u8 = undefined;
        var w = std.Io.Writer.fixed(&scratch);
        var err_buf: [256]u8 = undefined;
        var err_w = std.Io.Writer.fixed(&err_buf);
        const rc = try runDefinedNamesCommand(std.testing.allocator, io, .{ .file = path, .subcommand = .defined_names, .sheet_index = 1 }, &w, &err_w);
        try std.testing.expectEqual(@as(u8, 0), rc);
        try std.testing.expectEqualStrings(
            "{\"kind\":\"defined_name\",\"name\":\"_xlnm.Print_Area\",\"scope\":\"sheet\",\"sheet\":\"Report\",\"sheet_idx\":1,\"body\":\"Report!$A$1:$B$9\",\"hidden\":false}\n",
            w.buffered(),
        );
    }
    // --sheet 0: a sheet with no scoped names is an empty, successful
    // stream.
    {
        var scratch: [256]u8 = undefined;
        var w = std.Io.Writer.fixed(&scratch);
        var err_buf: [256]u8 = undefined;
        var err_w = std.Io.Writer.fixed(&err_buf);
        const rc = try runDefinedNamesCommand(std.testing.allocator, io, .{ .file = path, .subcommand = .defined_names, .sheet_index = 0 }, &w, &err_w);
        try std.testing.expectEqual(@as(u8, 0), rc);
        try std.testing.expectEqualStrings("", w.buffered());
    }
    // --name resolves by decoded sheet name.
    {
        var scratch: [1024]u8 = undefined;
        var w = std.Io.Writer.fixed(&scratch);
        var err_buf: [256]u8 = undefined;
        var err_w = std.Io.Writer.fixed(&err_buf);
        const rc = try runDefinedNamesCommand(std.testing.allocator, io, .{ .file = path, .subcommand = .defined_names, .sheet_name = "Report" }, &w, &err_w);
        try std.testing.expectEqual(@as(u8, 0), rc);
        try std.testing.expect(std.mem.indexOf(u8, w.buffered(), "\"name\":\"_xlnm.Print_Area\"") != null);
        try std.testing.expect(std.mem.indexOf(u8, w.buffered(), "\"name\":\"Prices\"") == null);
    }
    // --sheet-glob matches scope sheets; workbook-scope suppressed.
    {
        var scratch: [1024]u8 = undefined;
        var w = std.Io.Writer.fixed(&scratch);
        var err_buf: [256]u8 = undefined;
        var err_w = std.Io.Writer.fixed(&err_buf);
        const rc = try runDefinedNamesCommand(std.testing.allocator, io, .{ .file = path, .subcommand = .defined_names, .sheet_glob = "Rep*" }, &w, &err_w);
        try std.testing.expectEqual(@as(u8, 0), rc);
        try std.testing.expect(std.mem.indexOf(u8, w.buffered(), "\"name\":\"_xlnm.Print_Area\"") != null);
        try std.testing.expect(std.mem.indexOf(u8, w.buffered(), "\"scope\":\"workbook\"") == null);
    }
    // --skip / --take page the stream.
    {
        var scratch: [1024]u8 = undefined;
        var w = std.Io.Writer.fixed(&scratch);
        var err_buf: [256]u8 = undefined;
        var err_w = std.Io.Writer.fixed(&err_buf);
        const rc = try runDefinedNamesCommand(std.testing.allocator, io, .{ .file = path, .subcommand = .defined_names, .skip = 1, .take = 1 }, &w, &err_w);
        try std.testing.expectEqual(@as(u8, 0), rc);
        try std.testing.expect(std.mem.startsWith(u8, w.buffered(), "{\"kind\":\"defined_name\",\"name\":\"_xlnm.Print_Area\""));
        try std.testing.expectEqual(@as(?usize, null), std.mem.indexOf(u8, w.buffered(), "\"name\":\"Secret\""));
    }
    // A sheet selector that names no sheet is exit 3.
    {
        var scratch: [256]u8 = undefined;
        var w = std.Io.Writer.fixed(&scratch);
        var err_buf: [256]u8 = undefined;
        var err_w = std.Io.Writer.fixed(&err_buf);
        const rc = try runDefinedNamesCommand(std.testing.allocator, io, .{ .file = path, .subcommand = .defined_names, .sheet_name = "Nope" }, &w, &err_w);
        try std.testing.expectEqual(@as(u8, 3), rc);
        const rc2 = try runDefinedNamesCommand(std.testing.allocator, io, .{ .file = path, .subcommand = .defined_names, .sheet_index = 9 }, &w, &err_w);
        try std.testing.expectEqual(@as(u8, 3), rc2);
    }
}

test "runDefinedNamesCommand: a workbook without names is an empty stream" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(std.testing.allocator, io, "cli_defined_names_none.xlsx");
    defer std.testing.allocator.free(path);
    {
        const writer = xlsx.writer_types;
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s0 = try w.addSheet("Data");
        try s0.writeRow(&.{.{ .integer = 1 }});
        try w.save(io, path);
    }
    var scratch: [256]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    var err_buf: [256]u8 = undefined;
    var err_w = std.Io.Writer.fixed(&err_buf);
    const rc = try runDefinedNamesCommand(std.testing.allocator, io, .{ .file = path, .subcommand = .defined_names }, &w, &err_w);
    try std.testing.expectEqual(@as(u8, 0), rc);
    try std.testing.expectEqualStrings("", w.buffered());
}

// ─── S3b slice 4: `anchors` ──────────────────────────────────────────

test "parseArgs: anchors token and flag rejections" {
    {
        const argv = [_][]const u8{ "anchors", "f.xlsx" };
        const a = try parseArgs(&argv);
        try std.testing.expectEqual(Subcommand.anchors, a.subcommand);
        try std.testing.expectEqualStrings("f.xlsx", a.file);
    }
    // Anchor records are range-keyed, not row-keyed, and stream.
    {
        const argv = [_][]const u8{ "anchors", "f.xlsx", "--range", "A1:B2" };
        try std.testing.expectError(ArgError.BadArgValue, parseArgs(&argv));
    }
    {
        const argv = [_][]const u8{ "anchors", "f.xlsx", "--start-row", "2" };
        try std.testing.expectError(ArgError.BadArgValue, parseArgs(&argv));
    }
    {
        const argv = [_][]const u8{ "anchors", "f.xlsx", "--output", "pretty-json" };
        try std.testing.expectError(ArgError.BadArgValue, parseArgs(&argv));
    }
    {
        const argv = [_][]const u8{ "anchors", "f.xlsx", "--format", "csv" };
        try std.testing.expectError(ArgError.BadFormat, parseArgs(&argv));
    }
    {
        const argv = [_][]const u8{ "anchors", "f.xlsx", "--with-styles" };
        try std.testing.expectError(ArgError.BadArgValue, parseArgs(&argv));
    }
}

test "runAnchorsCommand: every sheet by default, exact wire shape" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(std.testing.allocator, io, "cli_anchors.xlsx");
    defer std.testing.allocator.free(path);
    try zlsx_pkg.anchors_ndjson.fixture.write(std.testing.allocator, io, path, .with_absolute);

    var scratch: [8192]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    var err_buf: [256]u8 = undefined;
    var err_w = std.Io.Writer.fixed(&err_buf);
    const rc = try runAnchorsCommand(std.testing.allocator, io, .{ .file = path, .subcommand = .anchors }, &w, &err_w);
    try std.testing.expectEqual(@as(u8, 0), rc);
    const png_len = zlsx_pkg.anchors_ndjson.fixture.png_bytes.len;
    var expected_buf: [2048]u8 = undefined;
    const expected = try std.fmt.bufPrint(
        &expected_buf,
        "{{\"kind\":\"image_anchor\",\"sheet\":\"Report\",\"sheet_idx\":1,\"part\":\"xl/media/image1.png\"," ++
            "\"anchor\":\"two_cell\",\"from\":{{\"row\":3,\"col\":2,\"row_off\":0,\"col_off\":9525}}," ++
            "\"to\":{{\"row\":8,\"col\":5,\"row_off\":19050,\"col_off\":0}},\"absolute\":null,\"bytes\":{d}}}\n" ++
            "{{\"kind\":\"image_anchor\",\"sheet\":\"Report\",\"sheet_idx\":1,\"part\":\"xl/media/image1.png\"," ++
            "\"anchor\":\"absolute\",\"from\":null,\"to\":null," ++
            "\"absolute\":{{\"x\":1000,\"y\":2000,\"cx\":914400,\"cy\":457200}},\"bytes\":{d}}}\n" ++
            "{{\"kind\":\"chart_anchor\",\"sheet\":\"Report\",\"sheet_idx\":1,\"part\":\"xl/charts/chart1.xml\"," ++
            "\"anchor\":\"one_cell\",\"from\":{{\"row\":2,\"col\":6,\"row_off\":0,\"col_off\":0}},\"to\":null," ++
            "\"absolute\":null,\"chart_type\":\"bar\"," ++
            "\"series_refs\":[\"Data!$B$1\",\"Data!$A$2:$A$4\",\"Data!$B$2:$B$4\"]}}\n",
        .{ png_len, png_len },
    );
    try std.testing.expectEqualStrings(expected, w.buffered());
}

test "runAnchorsCommand: selection, pagination, compact prologue, exit codes" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(std.testing.allocator, io, "cli_anchors_sel.xlsx");
    defer std.testing.allocator.free(path);
    try zlsx_pkg.anchors_ndjson.fixture.write(std.testing.allocator, io, path, .image_and_chart);

    // The anchor-less sheet selected by index: an empty, successful stream.
    {
        var scratch: [4096]u8 = undefined;
        var w = std.Io.Writer.fixed(&scratch);
        var err_buf: [256]u8 = undefined;
        var err_w = std.Io.Writer.fixed(&err_buf);
        const rc = try runAnchorsCommand(std.testing.allocator, io, .{ .file = path, .subcommand = .anchors, .sheet_index = 0 }, &w, &err_w);
        try std.testing.expectEqual(@as(u8, 0), rc);
        try std.testing.expectEqualStrings("", w.buffered());
    }
    // By name, compact: one prologue, envelope-less records.
    {
        var scratch: [4096]u8 = undefined;
        var w = std.Io.Writer.fixed(&scratch);
        var err_buf: [256]u8 = undefined;
        var err_w = std.Io.Writer.fixed(&err_buf);
        const rc = try runAnchorsCommand(std.testing.allocator, io, .{ .file = path, .subcommand = .anchors, .sheet_name = "Report", .output = .compact_ndjson }, &w, &err_w);
        try std.testing.expectEqual(@as(u8, 0), rc);
        const got = w.buffered();
        try std.testing.expect(std.mem.startsWith(u8, got, "{\"kind\":\"sheet\",\"sheet\":\"Report\",\"sheet_idx\":1}\n{\"kind\":\"image_anchor\",\"part\":"));
        try std.testing.expectEqual(@as(usize, 1), std.mem.count(u8, got, "\"kind\":\"sheet\""));
        try std.testing.expect(std.mem.indexOf(u8, got, "\"sheet_idx\":1,\"part\"") == null);
    }
    // --sheet-glob widens by name; --skip / --take page the stream.
    {
        var scratch: [4096]u8 = undefined;
        var w = std.Io.Writer.fixed(&scratch);
        var err_buf: [256]u8 = undefined;
        var err_w = std.Io.Writer.fixed(&err_buf);
        const rc = try runAnchorsCommand(std.testing.allocator, io, .{ .file = path, .subcommand = .anchors, .sheet_glob = "Rep*", .skip = 1, .take = 1 }, &w, &err_w);
        try std.testing.expectEqual(@as(u8, 0), rc);
        const got = w.buffered();
        try std.testing.expect(std.mem.startsWith(u8, got, "{\"kind\":\"chart_anchor\","));
        try std.testing.expectEqual(@as(usize, 1), std.mem.count(u8, got, "\n"));
    }
    // Unknown sheet: exit 3, like every read sub-command.
    {
        var scratch: [256]u8 = undefined;
        var w = std.Io.Writer.fixed(&scratch);
        var err_buf: [256]u8 = undefined;
        var err_w = std.Io.Writer.fixed(&err_buf);
        const rc = try runAnchorsCommand(std.testing.allocator, io, .{ .file = path, .subcommand = .anchors, .sheet_name = "Nope" }, &w, &err_w);
        try std.testing.expectEqual(@as(u8, 3), rc);
        const rc2 = try runAnchorsCommand(std.testing.allocator, io, .{ .file = path, .subcommand = .anchors, .sheet_index = 9 }, &w, &err_w);
        try std.testing.expectEqual(@as(u8, 3), rc2);
    }
}

test "runAnchorsCommand: a workbook without drawings is an empty stream" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(std.testing.allocator, io, "cli_anchors_none.xlsx");
    defer std.testing.allocator.free(path);
    {
        const writer = xlsx.writer_types;
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s0 = try w.addSheet("Data");
        try s0.writeRow(&.{.{ .integer = 1 }});
        try w.save(io, path);
    }
    var scratch: [256]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    var err_buf: [256]u8 = undefined;
    var err_w = std.Io.Writer.fixed(&err_buf);
    const rc = try runAnchorsCommand(std.testing.allocator, io, .{ .file = path, .subcommand = .anchors }, &w, &err_w);
    try std.testing.expectEqual(@as(u8, 0), rc);
    try std.testing.expectEqualStrings("", w.buffered());
}

test "runAnchorsCommand: a series ref the stream cannot carry refuses whole (exit 2)" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try tt.path(std.testing.allocator, io, "cli_anchors_bad_ref.xlsx");
    defer std.testing.allocator.free(path);
    try zlsx_pkg.anchors_ndjson.fixture.write(std.testing.allocator, io, path, .image_and_chart);
    // A bad entity in one `<c:f>` body: the whole command refuses —
    // a partial anchor inventory is the shape of a guard hole.
    try zlsx_pkg.pivots.fixture.patchPart(std.testing.allocator, io, path, "xl/charts/chart1.xml", "Data!$B$1", "Data!$B$1&bogus;");
    var scratch: [4096]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    var err_buf: [512]u8 = undefined;
    var err_w = std.Io.Writer.fixed(&err_buf);
    const rc = try runAnchorsCommand(std.testing.allocator, io, .{ .file = path, .subcommand = .anchors }, &w, &err_w);
    try std.testing.expectEqual(@as(u8, 2), rc);
    try std.testing.expectEqualStrings("", w.buffered());
    try std.testing.expect(std.mem.indexOf(u8, err_w.buffered(), "MalformedDrawingXml") != null);
}

// ─── S3b slice 3: `doc-props` ────────────────────────────────────────

/// A saved fresh-writer workbook with docProps parts spliced in through
/// the package store — `core_xml` / `app_xml` / `custom` optional, so
/// one helper serves the populated, absent and refused cases.
fn writeDocPropsFixture(
    io: std.Io,
    tt: *TestTmp,
    name: []const u8,
    core_xml: ?[]const u8,
    app_xml: ?[]const u8,
    custom: bool,
) ![:0]u8 {
    const alloc = std.testing.allocator;
    const path = try tt.path(alloc, io, name);
    {
        // Block-scoped: popped once the fresh file exists, so the
        // splice half's `defer` below is the only owner of `path`.
        errdefer alloc.free(path);
        const writer = xlsx.writer_types;
        var w = writer.Writer.init(alloc);
        defer w.deinit();
        var s0 = try w.addSheet("Data");
        try s0.writeRow(&.{.{ .integer = 1 }});
        try w.save(io, path);
    }
    if (core_xml == null and app_xml == null and !custom) return path;
    // Splice the parts and save to a second file: the store's lazy
    // part reads may still be backed by the source archive.
    defer alloc.free(path);
    const out_path = try tt.path(alloc, io, "spliced.xlsx");
    errdefer alloc.free(out_path);
    var wb = try zlsx_pkg.Workbook.open(alloc, io, path);
    defer wb.deinit();
    if (core_xml) |x| try wb.store.addPart("docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml", x);
    if (app_xml) |x| try wb.store.addPart("docProps/app.xml", "application/vnd.openxmlformats-officedocument.extended-properties+xml", x);
    if (custom) try wb.store.addPart("docProps/custom.xml", "application/vnd.openxmlformats-officedocument.custom-properties+xml", "<Properties/>");
    try wb.store.save(io, out_path);
    return out_path;
}

test "runDocPropsCommand: the full field set, text as stored, one record" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    // `&amp;` stays as stored: the record carries the same bytes
    // `Editor.doc_props` and `meta` hand over, not a re-decoding.
    const core =
        "<cp:coreProperties><dc:creator>A &amp; B</dc:creator>" ++
        "<cp:lastModifiedBy>C</cp:lastModifiedBy><dc:title>T</dc:title>" ++
        "<dcterms:created xsi:type=\"dcterms:W3CDTF\">2026-01-02T03:04:05Z</dcterms:created>" ++
        "<cp:revision>7</cp:revision></cp:coreProperties>";
    const app =
        "<Properties><Company>Acme</Company><Manager>M</Manager>" ++
        "<Application>zlsx-test</Application><HyperlinkBase>https://x</HyperlinkBase></Properties>";
    const path = try writeDocPropsFixture(io, &tt, "cli_doc_props.xlsx", core, app, true);
    defer std.testing.allocator.free(path);

    var scratch: [2048]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    var err_buf: [256]u8 = undefined;
    var err_w = std.Io.Writer.fixed(&err_buf);
    const rc = try runDocPropsCommand(std.testing.allocator, io, .{ .file = path, .subcommand = .doc_props }, &w, &err_w);
    try std.testing.expectEqual(@as(u8, 0), rc);
    try std.testing.expectEqualStrings(
        "{\"kind\":\"doc_props\",\"creator\":\"A &amp; B\",\"last_modified_by\":\"C\"," ++
            "\"title\":\"T\",\"subject\":null,\"description\":null,\"keywords\":null," ++
            "\"category\":null,\"created\":\"2026-01-02T03:04:05Z\",\"modified\":null," ++
            "\"revision\":\"7\",\"company\":\"Acme\",\"manager\":\"M\"," ++
            "\"application\":\"zlsx-test\",\"hyperlink_base\":\"https://x\"," ++
            "\"has_custom_properties\":true}\n",
        w.buffered(),
    );
}

test "runDocPropsCommand: no docProps parts is a record of nulls" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try writeDocPropsFixture(io, &tt, "cli_doc_props_none.xlsx", null, null, false);
    defer std.testing.allocator.free(path);

    var scratch: [1024]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    var err_buf: [256]u8 = undefined;
    var err_w = std.Io.Writer.fixed(&err_buf);
    const rc = try runDocPropsCommand(std.testing.allocator, io, .{ .file = path, .subcommand = .doc_props }, &w, &err_w);
    try std.testing.expectEqual(@as(u8, 0), rc);
    try std.testing.expectEqualStrings(
        "{\"kind\":\"doc_props\",\"creator\":null,\"last_modified_by\":null," ++
            "\"title\":null,\"subject\":null,\"description\":null,\"keywords\":null," ++
            "\"category\":null,\"created\":null,\"modified\":null,\"revision\":null," ++
            "\"company\":null,\"manager\":null,\"application\":null," ++
            "\"hyperlink_base\":null,\"has_custom_properties\":false}\n",
        w.buffered(),
    );
}

test "runDocPropsCommand: a non-UTF-8 field refuses whole before any byte" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    // The last field the table walks carries the bad bytes, so a
    // half-writing implementation would have emitted 13 fields first —
    // the empty stdout is the test.
    const app = "<Properties><HyperlinkBase>bad\xffbytes</HyperlinkBase></Properties>";
    const path = try writeDocPropsFixture(io, &tt, "cli_doc_props_bad.xlsx", "<cp:coreProperties><dc:creator>Ok</dc:creator></cp:coreProperties>", app, false);
    defer std.testing.allocator.free(path);

    var scratch: [1024]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    var err_buf: [512]u8 = undefined;
    var err_w = std.Io.Writer.fixed(&err_buf);
    const rc = try runDocPropsCommand(std.testing.allocator, io, .{ .file = path, .subcommand = .doc_props }, &w, &err_w);
    try std.testing.expectEqual(@as(u8, 2), rc);
    try std.testing.expectEqualStrings("", w.buffered());
    try std.testing.expect(std.mem.indexOf(u8, err_w.buffered(), "hyperlink_base is not UTF-8") != null);
}

test "parseArgs: doc-props token, workbook-scoped flag tolerance" {
    {
        const argv = [_][]const u8{ "doc-props", "f.xlsx" };
        const a = try parseArgs(&argv);
        try std.testing.expectEqual(Subcommand.doc_props, a.subcommand);
        try std.testing.expectEqualStrings("f.xlsx", a.file);
    }
    // The meta family's wrapper-friendliness: sheet and format flags
    // parse without error and are ignored downstream.
    {
        const argv = [_][]const u8{ "doc-props", "f.xlsx", "--sheet", "2", "--format", "csv" };
        const a = try parseArgs(&argv);
        try std.testing.expectEqual(Subcommand.doc_props, a.subcommand);
    }
}

test "runDocPropsCommand: a present-but-empty element is \"\", not null (Codex #213 r1)" {
    // `<dc:creator></dc:creator>` is present-and-empty: the CLI and
    // `meta` report "", where Python's Editor.doc_props maps the C
    // boundary's length-zero convention to None — the documented
    // divergence, pinned from this side.
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const core = "<cp:coreProperties><dc:creator></dc:creator></cp:coreProperties>";
    const path = try writeDocPropsFixture(io, &tt, "cli_doc_props_empty.xlsx", core, null, false);
    defer std.testing.allocator.free(path);

    var scratch: [1024]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    var err_buf: [256]u8 = undefined;
    var err_w = std.Io.Writer.fixed(&err_buf);
    const rc = try runDocPropsCommand(std.testing.allocator, io, .{ .file = path, .subcommand = .doc_props }, &w, &err_w);
    try std.testing.expectEqual(@as(u8, 0), rc);
    try std.testing.expect(std.mem.startsWith(u8, w.buffered(), "{\"kind\":\"doc_props\",\"creator\":\"\","));
}

test "runDocPropsCommand: a stdout that cannot take the record surfaces WriteFailed" {
    // The validity floor's promise is scoped to read and validation
    // failures; an I/O failure mid-line is the stream's, reported, not
    // swallowed.
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    var tt = TestTmp.init();
    defer tt.deinit();
    const path = try writeDocPropsFixture(io, &tt, "cli_doc_props_tiny_out.xlsx", null, null, false);
    defer std.testing.allocator.free(path);

    var scratch: [16]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    var err_buf: [256]u8 = undefined;
    var err_w = std.Io.Writer.fixed(&err_buf);
    try std.testing.expectError(
        error.WriteFailed,
        runDocPropsCommand(std.testing.allocator, io, .{ .file = path, .subcommand = .doc_props }, &w, &err_w),
    );
}

test "writeDocPropsPretty / writeDocPropsCompact: populated objects, exact bytes (Codex #213 r1)" {
    // The table refactor's byte identity: both meta emitters walked a
    // hand-listed field sequence before doc_prop_fields; these literals
    // pin the loop to the same bytes over a fully-populated view.
    const dp = zlsx_pkg.DocProps{
        .creator = "C",
        .last_modified_by = "L",
        .title = "T",
        .subject = "S",
        .description = "D",
        .keywords = "K",
        .category = "G",
        .created = "2026-01-02T03:04:05Z",
        .modified = "2026-01-02T03:04:06Z",
        .revision = "9",
        .company = "Co",
        .manager = "M",
        .application = "App",
        .hyperlink_base = "H",
        .has_custom_properties = true,
    };
    {
        var scratch: [1024]u8 = undefined;
        var w = std.Io.Writer.fixed(&scratch);
        try writeDocPropsPretty(&w, dp);
        try std.testing.expectEqualStrings(
            "  \"doc_props\": {\n" ++
                "    \"creator\": \"C\",\n" ++
                "    \"last_modified_by\": \"L\",\n" ++
                "    \"title\": \"T\",\n" ++
                "    \"subject\": \"S\",\n" ++
                "    \"description\": \"D\",\n" ++
                "    \"keywords\": \"K\",\n" ++
                "    \"category\": \"G\",\n" ++
                "    \"created\": \"2026-01-02T03:04:05Z\",\n" ++
                "    \"modified\": \"2026-01-02T03:04:06Z\",\n" ++
                "    \"revision\": \"9\",\n" ++
                "    \"company\": \"Co\",\n" ++
                "    \"manager\": \"M\",\n" ++
                "    \"application\": \"App\",\n" ++
                "    \"hyperlink_base\": \"H\",\n" ++
                "    \"has_custom_properties\": true\n  },\n",
            w.buffered(),
        );
    }
    {
        var scratch: [1024]u8 = undefined;
        var w = std.Io.Writer.fixed(&scratch);
        try writeDocPropsCompact(&w, dp);
        try std.testing.expectEqualStrings(
            ",\"doc_props\":{\"creator\":\"C\",\"last_modified_by\":\"L\",\"title\":\"T\"," ++
                "\"subject\":\"S\",\"description\":\"D\",\"keywords\":\"K\",\"category\":\"G\"," ++
                "\"created\":\"2026-01-02T03:04:05Z\",\"modified\":\"2026-01-02T03:04:06Z\"," ++
                "\"revision\":\"9\",\"company\":\"Co\",\"manager\":\"M\",\"application\":\"App\"," ++
                "\"hyperlink_base\":\"H\",\"has_custom_properties\":true}",
            w.buffered(),
        );
    }
}
