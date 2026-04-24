//! `zlsx` — read-only command-line interface over the zlsx library.
//!
//! Streams rows of the selected sheet to stdout in one of four formats.
//! Designed as a drop-in openpyxl replacement for reads: shell-friendly,
//! pipeable into jq / awk, no Python interpreter floor.

const std = @import("std");
const builtin = @import("builtin");
const xlsx = @import("xlsx.zig");

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
    styles,
    sst,
};

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
            std.mem.eql(u8, a, "--sheet-glob"))
        {
            i += 1; // skip paired value (bounds-checked by caller)
            continue;
        }
        if (a.len > 0 and a[0] == '-') continue; // flag with no value
        if (std.mem.eql(u8, a, "cells")) return .cells;
        if (std.mem.eql(u8, a, "rows")) return .rows;
        if (std.mem.eql(u8, a, "meta")) return .meta;
        if (std.mem.eql(u8, a, "list-sheets")) return .list_sheets;
        if (std.mem.eql(u8, a, "comments")) return .comments;
        if (std.mem.eql(u8, a, "validations")) return .validations;
        if (std.mem.eql(u8, a, "hyperlinks")) return .hyperlinks;
        if (std.mem.eql(u8, a, "styles")) return .styles;
        if (std.mem.eql(u8, a, "sst")) return .sst;
        return .rows; // first positional is the file path
    }
    return .rows;
}

fn parseArgs(argv: []const []const u8) ArgError!Args {
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
        => true,
        .rows, .cells, .comments, .validations, .hyperlinks => false,
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
                    std.mem.eql(u8, a, "comments") or
                    std.mem.eql(u8, a, "validations") or
                    std.mem.eql(u8, a, "hyperlinks") or
                    std.mem.eql(u8, a, "styles") or
                    std.mem.eql(u8, a, "sst"))
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
            .validations, .hyperlinks, .meta, .list_sheets, .styles, .sst => {
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
            .comments, .validations, .hyperlinks, .meta, .list_sheets, .styles, .sst => {
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
            .comments, .validations, .hyperlinks, .meta, .list_sheets, .styles, .sst => {
                return ArgError.BadArgValue;
            },
        }
        if (out.list_sheets) return ArgError.BadArgValue;
    }
    if (out.with_styles) {
        switch (detected_sub) {
            .rows, .cells => {},
            .comments, .validations, .hyperlinks, .meta, .list_sheets, .styles, .sst => {
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
    return out;
}

fn writeUsage(w: *std.Io.Writer) !void {
    try w.writeAll(
        \\usage: zlsx [<subcommand>] <file.xlsx> [options]
        \\
        \\  --sheet N         0-indexed sheet to read (default: 0)
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
        \\                    hyperlinks / styles / sst. Ignored by meta
        \\                    and list-sheets.
        \\  --take N          stop after N emitted records. Same scope
        \\                    as --skip; combine for middle-slice paging.
        \\  --start-row R     (iter59b) 1-based OOXML row; drop records
        \\                    whose row < R. Per-sheet scope (each
        \\                    sheet's own rows, unlike --skip which is
        \\                    global). Valid for rows / cells / comments
        \\                    only; rejected on validations / hyperlinks
        \\                    / meta / list-sheets / styles / sst.
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
        \\  styles             one NDJSON record per cell-XF style entry
        \\                     (workbook-wide, iter58):
        \\                     {"kind":"style","idx":0,"font":{…}|null,
        \\                      "fill":{…}|null,"border":{…}|null,
        \\                      "num_fmt":"General"|null}
        \\  sst                one NDJSON record per shared-string entry
        \\                     (workbook-wide, iter58):
        \\                     {"kind":"sst","idx":0,"text":"…","runs":null}
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
        \\Exit codes
        \\  0  success
        \\  1  bad arguments
        \\  2  could not open or parse workbook
        \\  3  sheet not found
        \\
    );
}

fn colLetter(buf: *[8]u8, idx: usize) []const u8 {
    var i: usize = idx + 1; // xlsx columns are 1-based
    var pos: usize = buf.len;
    while (i > 0) {
        i -= 1;
        pos -= 1;
        buf[pos] = 'A' + @as(u8, @intCast(i % 26));
        i /= 26;
    }
    return buf[pos..];
}

fn writeJsonString(w: *std.Io.Writer, s: []const u8) !void {
    try w.writeByte('"');
    for (s) |c| switch (c) {
        '"' => try w.writeAll("\\\""),
        '\\' => try w.writeAll("\\\\"),
        '\n' => try w.writeAll("\\n"),
        '\r' => try w.writeAll("\\r"),
        '\t' => try w.writeAll("\\t"),
        0x08 => try w.writeAll("\\b"),
        0x0c => try w.writeAll("\\f"),
        0...0x07, 0x0b, 0x0e...0x1f => try w.print("\\u{x:0>4}", .{c}),
        else => try w.writeByte(c),
    };
    try w.writeByte('"');
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
/// types this slice emits — date/formula/error are future work.
fn envelopeTypeTag(cell: xlsx.Cell) []const u8 {
    return switch (cell) {
        .empty => unreachable, // caller skips empties
        .string => "str",
        .integer => "int",
        .number => "num",
        .boolean => "bool",
    };
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
) !void {
    try w.writeByte('[');
    var first = true;
    for (cells, 0..) |c, i| {
        if (c == .empty and !include_blanks) continue;
        if (!first) try w.writeByte(',');
        first = false;

        const absolute_col: u32 = col_offset + @as(u32, @intCast(i));
        var col_buf: [8]u8 = undefined;
        const letters = colLetter(&col_buf, absolute_col);
        var ref_buf: [16]u8 = undefined;
        const ref = std.fmt.bufPrint(&ref_buf, "{s}{d}", .{ letters, row_number }) catch unreachable;

        try w.writeAll("{\"ref\":");
        try writeJsonString(w, ref);
        switch (c) {
            .empty => try w.print(",\"col\":{d},\"t\":\"blank\",\"v\":null", .{absolute_col + 1}),
            else => {
                try w.print(",\"col\":{d},\"t\":\"{s}\",\"v\":", .{ absolute_col + 1, envelopeTypeTag(c) });
                try writeJsonCell(w, c);
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
) !void {
    try w.writeAll("{\"kind\":\"row\",\"sheet\":");
    try writeJsonString(w, sheet_name);
    try w.print(",\"sheet_idx\":{d},\"row\":{d},\"cells\":", .{ sheet_idx, row_number });
    try writeEnvelopeCells(w, cells, row_number, include_blanks, style_ctx, col_offset);
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
) !void {
    try w.writeAll("{\"kind\":\"row\",\"sheet\":");
    try writeJsonString(w, sheet_name);
    try w.print(",\"sheet_idx\":{d},\"row\":{d},\"fields\":{{", .{ sheet_idx, row_number });
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
) !void {
    try w.writeAll("{\"kind\":\"cell\",\"sheet\":");
    try writeJsonString(w, sheet_name);
    try w.print(",\"sheet_idx\":{d},\"ref\":", .{sheet_idx});
    try writeJsonString(w, ref);
    try w.print(",\"row\":{d},\"col\":{d},\"t\":", .{ row, col });
    switch (cell) {
        .empty => try w.writeAll("\"blank\",\"v\":null"),
        else => {
            try w.print("\"{s}\",\"v\":", .{envelopeTypeTag(cell)});
            try writeJsonCell(w, cell);
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

pub fn main() !u8 {
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

    const raw_args = try std.process.argsAlloc(alloc);
    defer std.process.argsFree(alloc, raw_args);

    var stdout_buf: [16 * 1024]u8 = undefined;
    var stdout_file = std.fs.File.stdout().writer(&stdout_buf);
    const out = &stdout_file.interface;
    defer out.flush() catch {};

    var stderr_buf: [4 * 1024]u8 = undefined;
    var stderr_file = std.fs.File.stderr().writer(&stderr_buf);
    const err = &stderr_file.interface;
    defer err.flush() catch {};

    const args = parseArgs(raw_args[1..]) catch |e| switch (e) {
        ArgError.HelpRequested => {
            try writeUsage(out);
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

    var book = xlsx.Book.open(alloc, args.file) catch |e| {
        try err.print("zlsx: cannot open '{s}': {s}\n", .{ args.file, @errorName(e) });
        return 2;
    };
    defer book.deinit();

    if (args.list_sheets) {
        for (book.sheets) |s| {
            try out.writeAll(s.name);
            try out.writeByte('\n');
        }
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
            try runMetaCommand(out, &book, path_opt);
            return 0;
        },
        .list_sheets => {
            try runListSheetsCommand(out, &book);
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
        .styles => {
            try runStylesCommand(out, &book, args.skip, args.take);
            return 0;
        },
        .sst => {
            try runSstCommand(out, &book, args.skip, args.take);
            return 0;
        },
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
    try runRowsOnSheet(out, book, sheet, sheet_idx, format, alloc, &pg, start_row, end_row, range, header, include_blanks, with_styles);
}

/// iter59c: per-sheet body of `rows`. Takes the Pagination by pointer
/// so cross-sheet drivers can thread one counter through every sheet —
/// `--skip N --take M` slices the concatenated stream, not per sheet.
/// Header state is local to this call (per-sheet reset by design).
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
) !void {
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
    var header_keys: std.ArrayListUnmanaged([]u8) = .{};
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
    var masked: std.ArrayListUnmanaged(xlsx.Cell) = .{};
    defer masked.deinit(alloc);
    // iter59b-4: parallel masked style indices — lives next to `masked`
    // so `masked.items[i]` and `masked_styles.items[i]` stay paired
    // through every view transformation below. Only populated when
    // --range is present AND we're on the envelope-positional path.
    var masked_styles: std.ArrayListUnmanaged(?u32) = .{};
    defer masked_styles.deinit(alloc);

    while (try rows.next()) |cells| {
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
            col_offset: u32,
            any_non_empty: bool,
        };
        const raw_styles = rows.styleIndices();
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
            try masked.ensureTotalCapacity(alloc, range_width);
            try masked_styles.ensureTotalCapacity(alloc, range_width);
            var any = false;
            var col: u32 = r.top_left.col;
            while (col <= r.bottom_right.col) : (col += 1) {
                const src_idx: usize = col;
                if (src_idx < cells.len) {
                    masked.appendAssumeCapacity(cells[src_idx]);
                    masked_styles.appendAssumeCapacity(
                        if (src_idx < raw_styles.len) raw_styles[src_idx] else null,
                    );
                    if (cells[src_idx] != .empty) any = true;
                } else {
                    masked.appendAssumeCapacity(.empty);
                    masked_styles.appendAssumeCapacity(null);
                }
            }
            break :blk .{
                .cells = masked.items,
                .style_indices = masked_styles.items,
                .col_offset = r.top_left.col,
                .any_non_empty = any,
            };
        } else .{
            .cells = cells,
            .style_indices = raw_styles,
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
        if (header) {
            try writeRowEnvelopeDict(out, sheet.name, sheet_idx, row_number, header_keys.items, view.cells);
        } else switch (format) {
            .jsonl => {
                const style_ctx: ?EnvelopeStyleCtx = if (with_styles)
                    .{ .book = book, .style_indices = view.style_indices }
                else
                    null;
                try writeRowEnvelope(out, sheet.name, sheet_idx, row_number, view.cells, include_blanks, style_ctx, view.col_offset);
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
    try runCellsOnSheet(out, book, sheet, sheet_idx, alloc, &pg, start_row, end_row, range, include_blanks, with_styles);
}

/// iter59c: per-sheet cell emitter — takes Pagination by pointer so a
/// cross-sheet driver can thread the same counter across every sheet.
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
) !void {
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
        const row_number = rows.currentRowNumber();
        if (row_lo) |s| if (row_number < s) continue;
        if (row_hi) |e| if (row_number > e) break;
        const raw_styles = rows.styleIndices();
        for (cells, 0..) |c, i| {
            // iter59b-4: --include-blanks flips the empty-skip into
            // emit-as-blank. Without the flag, the old sparse-by-default
            // behaviour holds.
            if (c == .empty and !include_blanks) continue;
            if (range) |r| {
                const col: u32 = @intCast(i);
                if (col < r.top_left.col or col > r.bottom_right.col) continue;
            }

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

            try writeCell(
                out,
                sheet.name,
                sheet_idx,
                ref,
                row_number,
                @intCast(i + 1),
                c,
                style_ctx,
            );
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
        );
    }
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
        );
    }
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
) !void {
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

    // `path` is null when the caller detected non-UTF-8 bytes in the
    // original argv — emit JSON `null` so the NDJSON line stays
    // parseable. main() has already logged the reason to stderr.
    try out.writeAll("{\"kind\":\"workbook\",\"path\":");
    if (path) |p| try writeJsonString(out, p) else try out.writeAll("null");
    try out.print(
        ",\"sheets\":{d},\"sst\":{{\"count\":{d},\"rich\":{d}}}",
        .{ book.sheets.len, book.shared_strings.len, book.rich_runs_by_sst_idx.count() },
    );
    try out.print(
        ",\"has_styles\":{s},\"has_theme\":{s},\"has_comments\":{s}}}\n",
        .{
            if (book.styles_xml != null) "true" else "false",
            if (book.theme_xml != null) "true" else "false",
            if (any_comments) "true" else "false",
        },
    );

    for (book.sheets, 0..) |s, i| {
        const sheet_has_comments = book.comments(s).len != 0;
        try out.writeAll("{\"kind\":\"sheet\",\"sheet\":");
        try writeJsonString(out, s.name);
        try out.print(
            ",\"sheet_idx\":{d},\"has_comments\":{s}}}\n",
            .{ i, if (sheet_has_comments) "true" else "false" },
        );
    }
}

/// iter57: lighter NDJSON variant of `meta` — one record per sheet,
/// name + index only. Same envelope shape as `meta`'s sheet record
/// minus the workbook-scoped `has_comments` field, so consumers can
/// trivially swap between the two commands.
fn runListSheetsCommand(out: *std.Io.Writer, book: *const xlsx.Book) !void {
    for (book.sheets, 0..) |s, i| {
        try out.writeAll("{\"kind\":\"sheet\",\"sheet\":");
        try writeJsonString(out, s.name);
        try out.print(",\"sheet_idx\":{d}}}\n", .{i});
    }
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
    for (book.sheets, 0..) |s, sheet_idx| {
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
            switch (pg.consume()) {
                .drop => continue,
                .stop => return,
                .emit => {},
            }
            var ref_buf: [16]u8 = undefined;
            const ref = refFromCellRef(&ref_buf, c.top_left);

            try out.writeAll("{\"kind\":\"comment\",\"sheet\":");
            try writeJsonString(out, s.name);
            try out.print(",\"sheet_idx\":{d},\"ref\":", .{sheet_idx});
            try writeJsonString(out, ref);
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
    for (book.sheets, 0..) |s, sheet_idx| {
        if (filter) |f| {
            if (sheet_idx != f) continue;
        } else if (args.sheet_glob != null or args.all_sheets) {
            if (!isSheetIncluded(args, s.name, sheet_idx)) continue;
        }
        for (book.dataValidations(s)) |dv| {
            switch (pg.consume()) {
                .drop => continue,
                .stop => return,
                .emit => {},
            }
            var range_buf: [32]u8 = undefined;
            const range = rangeFromBounds(&range_buf, dv.top_left, dv.bottom_right);

            try out.writeAll("{\"kind\":\"validation\",\"sheet\":");
            try writeJsonString(out, s.name);
            try out.print(",\"sheet_idx\":{d},\"range\":", .{sheet_idx});
            try writeJsonString(out, range);
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
    for (book.sheets, 0..) |s, sheet_idx| {
        if (filter) |f| {
            if (sheet_idx != f) continue;
        } else if (args.sheet_glob != null or args.all_sheets) {
            if (!isSheetIncluded(args, s.name, sheet_idx)) continue;
        }
        for (book.hyperlinks(s)) |h| {
            switch (pg.consume()) {
                .drop => continue,
                .stop => return,
                .emit => {},
            }
            var range_buf: [32]u8 = undefined;
            const range = rangeFromBounds(&range_buf, h.top_left, h.bottom_right);

            try out.writeAll("{\"kind\":\"hyperlink\",\"sheet\":");
            try writeJsonString(out, s.name);
            try out.print(",\"sheet_idx\":{d},\"range\":", .{sheet_idx});
            try writeJsonString(out, range);
            try out.writeAll(",\"url\":");
            if (h.url.len != 0) try writeJsonString(out, h.url) else try out.writeAll("null");
            try out.writeAll(",\"location\":");
            if (h.location.len != 0) try writeJsonString(out, h.location) else try out.writeAll("null");
            try out.writeAll("}\n");
        }
    }
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
}

/// Emit one NDJSON record per shared-string entry.
fn runSstCommand(
    out: *std.Io.Writer,
    book: *const xlsx.Book,
    skip: ?usize,
    take: ?usize,
) !void {
    var pg = Pagination.init(skip, take);
    for (book.shared_strings, 0..) |s, i| {
        switch (pg.consume()) {
            .drop => continue,
            .stop => return,
            .emit => {},
        }
        try out.print("{{\"kind\":\"sst\",\"idx\":{d},\"text\":", .{i});
        try writeJsonString(out, s);
        try out.writeAll(",\"runs\":");
        try writeRichRunsOrNull(out, book.richRuns(i));
        try out.writeAll("}\n");
    }
}

// ─── Tests ───────────────────────────────────────────────────────────

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

test "parseArgs basics" {
    const argv = [_][]const u8{ "file.xlsx", "--sheet", "2", "--format", "csv" };
    const a = try parseArgs(&argv);
    try std.testing.expectEqualStrings("file.xlsx", a.file);
    try std.testing.expectEqual(@as(?usize, 2), a.sheet_index);
    try std.testing.expectEqual(Format.csv, a.format);
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
    try writeRowEnvelope(&w, "Data", 0, 1, &cells, false, null, 0);
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
    try writeRowEnvelope(&w, "S", 2, 7, &cells, false, null, 0);
    try std.testing.expectEqualStrings(
        "{\"kind\":\"row\",\"sheet\":\"S\",\"sheet_idx\":2,\"row\":7,\"cells\":[]}\n",
        w.buffered(),
    );
}

test "writeRowEnvelope escapes sheet name" {
    var scratch: [256]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    const cells = [_]xlsx.Cell{.{ .integer = 1 }};
    try writeRowEnvelope(&w, "She\"et\n", 0, 1, &cells, false, null, 0);
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
    try writeRowEnvelope(&w, "S", 0, 1, &cells, false, null, 0);
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
        try writeCell(&w, "Data", 0, "A1", 1, 1, .{ .string = "name" }, null);
        try std.testing.expectEqualStrings(
            "{\"kind\":\"cell\",\"sheet\":\"Data\",\"sheet_idx\":0,\"ref\":\"A1\",\"row\":1,\"col\":1,\"t\":\"str\",\"v\":\"name\"}\n",
            w.buffered(),
        );
    }
    // integer
    {
        var w = std.Io.Writer.fixed(&scratch);
        try writeCell(&w, "Data", 0, "B2", 2, 2, .{ .integer = 3 }, null);
        try std.testing.expectEqualStrings(
            "{\"kind\":\"cell\",\"sheet\":\"Data\",\"sheet_idx\":0,\"ref\":\"B2\",\"row\":2,\"col\":2,\"t\":\"int\",\"v\":3}\n",
            w.buffered(),
        );
    }
    // number
    {
        var w = std.Io.Writer.fixed(&scratch);
        try writeCell(&w, "Data", 0, "C3", 3, 3, .{ .number = 3.5 }, null);
        try std.testing.expectEqualStrings(
            "{\"kind\":\"cell\",\"sheet\":\"Data\",\"sheet_idx\":0,\"ref\":\"C3\",\"row\":3,\"col\":3,\"t\":\"num\",\"v\":3.5}\n",
            w.buffered(),
        );
    }
    // boolean
    {
        var w = std.Io.Writer.fixed(&scratch);
        try writeCell(&w, "Data", 0, "D4", 4, 4, .{ .boolean = true }, null);
        try std.testing.expectEqualStrings(
            "{\"kind\":\"cell\",\"sheet\":\"Data\",\"sheet_idx\":0,\"ref\":\"D4\",\"row\":4,\"col\":4,\"t\":\"bool\",\"v\":true}\n",
            w.buffered(),
        );
    }
}

test "writeCell escapes sheet name" {
    var scratch: [256]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    try writeCell(&w, "She\"et\n", 2, "A1", 1, 1, .{ .integer = 7 }, null);
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
        try writeCell(&w, "S", 0, ref, row_number, @intCast(i + 1), c, null);
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
    inline for (.{ "validations", "hyperlinks", "meta", "list-sheets", "styles", "sst" }) |cmd| {
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
    const tmp_path = "/tmp/zlsx_cli_rowbounds_iter59b.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};
    {
        const writer = @import("writer.zig");
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s0 = try w.addSheet("Data");
        // 5 rows × 1 cell each → rows 1..5 in the OOXML sense.
        try s0.writeRow(&.{.{ .string = "c1" }});
        try s0.writeRow(&.{.{ .string = "c2" }});
        try s0.writeRow(&.{.{ .string = "c3" }});
        try s0.writeRow(&.{.{ .string = "c4" }});
        try s0.writeRow(&.{.{ .string = "c5" }});
        try w.save(tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, tmp_path);
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
    inline for (.{ "comments", "validations", "hyperlinks", "meta", "list-sheets", "styles", "sst" }) |cmd| {
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
    const tmp_path = "/tmp/zlsx_cli_range_iter59b2.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};
    {
        const writer = @import("writer.zig");
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
        try w.save(tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, tmp_path);
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
    const tmp_path = "/tmp/zlsx_cli_range_rows_iter59b2.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};
    {
        const writer = @import("writer.zig");
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s0 = try w.addSheet("Data");
        try s0.writeRow(&.{ .{ .string = "A1" }, .{ .string = "B1" }, .{ .string = "C1" }, .{ .string = "D1" }, .{ .string = "E1" } });
        try s0.writeRow(&.{ .{ .string = "A2" }, .{ .string = "B2" }, .{ .string = "C2" }, .{ .string = "D2" }, .{ .string = "E2" } });
        try s0.writeRow(&.{ .{ .string = "A3" }, .{ .string = "B3" }, .{ .string = "C3" }, .{ .string = "D3" }, .{ .string = "E3" } });
        try s0.writeRow(&.{ .{ .string = "A4" }, .{ .string = "B4" }, .{ .string = "C4" }, .{ .string = "D4" }, .{ .string = "E4" } });
        try s0.writeRow(&.{ .{ .string = "A5" }, .{ .string = "B5" }, .{ .string = "C5" }, .{ .string = "D5" }, .{ .string = "E5" } });
        try w.save(tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, tmp_path);
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
    const tmp_path = "/tmp/zlsx_cli_header_iter59b3.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};
    {
        const writer = @import("writer.zig");
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s0 = try w.addSheet("Data");
        try s0.writeRow(&.{ .{ .string = "name" }, .{ .string = "qty" } });
        try s0.writeRow(&.{ .{ .string = "apple" }, .{ .integer = 3 } });
        try s0.writeRow(&.{ .{ .string = "pear" }, .{ .integer = 7 } });
        try w.save(tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, tmp_path);
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
    const tmp_path = "/tmp/zlsx_cli_header_dup_iter59b3.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};
    {
        const writer = @import("writer.zig");
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s0 = try w.addSheet("Data");
        try s0.writeRow(&.{ .{ .string = "x" }, .{ .string = "x" } });
        try s0.writeRow(&.{ .{ .integer = 1 }, .{ .integer = 2 } });
        try w.save(tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, tmp_path);
    defer book.deinit();

    var scratch: [1024]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    try runRowsCommand(&w, &book, book.sheets[0], 0, .jsonl, std.testing.allocator, null, null, null, null, null, true, false, false);
    const out = w.buffered();

    // Both duplicate "x" keys appear in the dict as-is.
    try std.testing.expect(std.mem.indexOf(u8, out, "\"fields\":{\"x\":1,\"x\":2}") != null);
}

test "runRowsCommand --header empty header cells fall back to col_<letter>" {
    const tmp_path = "/tmp/zlsx_cli_header_empty_iter59b3.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};
    {
        const writer = @import("writer.zig");
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s0 = try w.addSheet("Data");
        // Header: A="name", B=empty, C="qty" → keys "name","col_B","qty".
        try s0.writeRow(&.{ .{ .string = "name" }, .empty, .{ .string = "qty" } });
        try s0.writeRow(&.{ .{ .string = "apple" }, .{ .integer = 42 }, .{ .integer = 3 } });
        try w.save(tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, tmp_path);
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
    // Header row has 4 cells across A..D ("w","x","y","z"); a --range
    // B:C must consume only the B/C header cells and emit data dicts
    // keyed exactly {"x","y"} — no `col_A` / `col_D` leak from the
    // masked full-width view.
    const tmp_path = "/tmp/zlsx_cli_header_range_iter59b3.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};
    {
        const writer = @import("writer.zig");
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s0 = try w.addSheet("Data");
        try s0.writeRow(&.{ .{ .string = "w" }, .{ .string = "x" }, .{ .string = "y" }, .{ .string = "z" } });
        try s0.writeRow(&.{ .{ .integer = 1 }, .{ .integer = 2 }, .{ .integer = 3 }, .{ .integer = 4 } });
        try w.save(tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, tmp_path);
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
    // Tight contract per iter59b-4 P2 follow-up: --include-blanks
    // preserves all-blank rows ONLY on the envelope (.jsonl) path.
    // On csv / tsv / legacy-jsonl / legacy-jsonl-dict the flag is a
    // documented no-op and must NOT inject extra blank output lines.
    // On --header the flag is also a no-op — a blank row must not
    // promote to a `col_*`-keyed header.
    const tmp_path = "/tmp/zlsx_cli_blanks_flat_iter59b4.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};
    {
        const writer = @import("writer.zig");
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s0 = try w.addSheet("Data");
        // Row 1: A only; B/C blank. Row 2: C only; A/B blank.
        try s0.writeRow(&.{ .{ .string = "x" }, .empty, .empty });
        try s0.writeRow(&.{ .empty, .empty, .{ .string = "y" } });
        try w.save(tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, tmp_path);
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
    // A row with data only in A/D (both outside the B:C range) and
    // --include-blanks must still emit with two t:"blank" cells —
    // the whole point of --include-blanks is to surface empties.
    const tmp_path = "/tmp/zlsx_cli_range_blank_iter59b4.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};
    {
        const writer = @import("writer.zig");
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s0 = try w.addSheet("Data");
        // Row 1: "x" in A, "y" in D. Nothing in B/C.
        try s0.writeRow(&.{ .{ .string = "x" }, .empty, .empty, .{ .string = "y" } });
        try w.save(tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, tmp_path);
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
    // Codex P2: a cell whose border has ONLY the diagonal side set
    // must not serialize `"border":{}` — the terse emitter omits
    // diagonal entirely, so emitting an empty border object would
    // be a shape leak. A style block may still appear because the
    // Zig writer attaches a default font to every styled cell
    // (which may have a color), but the "border" key must never
    // appear with an empty object.
    const tmp_path = "/tmp/zlsx_cli_diag_only_iter59b4.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};
    {
        const writer = @import("writer.zig");
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        const diag_style = try w.addStyle(.{
            .border_diagonal = .{ .style = .thin, .color_argb = 0xFF000000 },
        });
        var s0 = try w.addSheet("Data");
        try s0.writeRowStyled(&.{.{ .string = "x" }}, &.{diag_style});
        try w.save(tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, tmp_path);
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
    const tmp_path = "/tmp/zlsx_cli_blanks_iter59b4.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};
    {
        const writer = @import("writer.zig");
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s0 = try w.addSheet("Data");
        // Row 1: A="x", B=empty, C=7 → a single sparse row with a gap.
        try s0.writeRow(&.{ .{ .string = "x" }, .empty, .{ .integer = 7 } });
        try w.save(tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, tmp_path);
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
    // Regression guard: default behaviour preserved when the flag is off.
    const tmp_path = "/tmp/zlsx_cli_blanks_off_iter59b4.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};
    {
        const writer = @import("writer.zig");
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s0 = try w.addSheet("Data");
        try s0.writeRow(&.{ .{ .string = "x" }, .empty, .{ .integer = 7 } });
        try w.save(tmp_path);
    }
    var book = try xlsx.Book.open(std.testing.allocator, tmp_path);
    defer book.deinit();

    var scratch: [2048]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    try runCellsCommand(&w, &book, book.sheets[0], 0, std.testing.allocator, null, null, null, null, null, false, false);
    const out = w.buffered();

    try std.testing.expect(std.mem.indexOf(u8, out, "\"blank\"") == null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"B1\"") == null);
}

test "runCellsCommand --with-styles emits terse style block for styled cells" {
    const tmp_path = "/tmp/zlsx_cli_with_styles_iter59b4.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};
    {
        const writer = @import("writer.zig");
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
        try w.save(tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, tmp_path);
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
    const tmp_path = "/tmp/zlsx_cli_rows_styles_iter59b4.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};
    {
        const writer = @import("writer.zig");
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        const italic = try w.addStyle(.{ .font_italic = true });
        var s0 = try w.addSheet("Data");
        try s0.writeRowStyled(
            &.{ .{ .string = "a" }, .{ .string = "b" } },
            &.{ italic, 0 },
        );
        try w.save(tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, tmp_path);
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
    const tmp_path = "/tmp/zlsx_cli_pagination_iter59a.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};
    {
        const writer = @import("writer.zig");
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s0 = try w.addSheet("Data");
        // 5 rows × 1 cell each → 5 candidate cells in emit order.
        try s0.writeRow(&.{.{ .string = "c1" }});
        try s0.writeRow(&.{.{ .string = "c2" }});
        try s0.writeRow(&.{.{ .string = "c3" }});
        try s0.writeRow(&.{.{ .string = "c4" }});
        try s0.writeRow(&.{.{ .string = "c5" }});
        try w.save(tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, tmp_path);
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
    var scratch: [512]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);

    // Build a minimal Book-shaped view without actually opening a
    // file — runMetaCommand only dereferences book.sheets / sst /
    // styles_xml / theme_xml / rich_runs_by_sst_idx / comments.
    var empty_book: xlsx.Book = .{
        .allocator = std.testing.allocator,
        .sst_arena = std.heap.ArenaAllocator.init(std.testing.allocator),
    };
    defer empty_book.deinit();

    try runMetaCommand(&w, &empty_book, null);

    const out = scratch[0..w.end];
    // The path field must serialize as literal `null`, not a string.
    try std.testing.expect(std.mem.indexOf(u8, out, "\"path\":null") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"kind\":\"workbook\"") != null);
}

test "runListSheetsCommand emits one sheet record per sheet" {
    const tmp_path = "/tmp/zlsx_cli_list_sheets_iter57.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};
    {
        const writer = @import("writer.zig");
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s0 = try w.addSheet("Data");
        try s0.writeRow(&.{.{ .string = "hdr" }});
        var s1 = try w.addSheet("Other");
        try s1.writeRow(&.{.{ .integer = 1 }});
        var s2 = try w.addSheet("She\"et"); // name with a quote — must be JSON-escaped
        try s2.writeRow(&.{.{ .boolean = true }});
        try w.save(tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, tmp_path);
    defer book.deinit();

    var scratch: [1024]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    try runListSheetsCommand(&w, &book);
    try std.testing.expectEqualStrings(
        "{\"kind\":\"sheet\",\"sheet\":\"Data\",\"sheet_idx\":0}\n" ++
            "{\"kind\":\"sheet\",\"sheet\":\"Other\",\"sheet_idx\":1}\n" ++
            "{\"kind\":\"sheet\",\"sheet\":\"She\\\"et\",\"sheet_idx\":2}\n",
        w.buffered(),
    );
}

test "runMetaCommand emits workbook record with sst/has_* fields then sheet records" {
    const tmp_path = "/tmp/zlsx_cli_meta_iter57.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};
    {
        const writer = @import("writer.zig");
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
        try w.save(tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, tmp_path);
    defer book.deinit();

    var scratch: [4096]u8 = undefined;
    var w = std.Io.Writer.fixed(&scratch);
    try runMetaCommand(&w, &book, tmp_path);

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
    // Regression guard: the legacy plain-text shape is exactly
    // `<name>\n` per sheet, no JSON, no sub-command routing. This
    // mirrors the code path in main() line-for-line so the flag
    // keeps working across iter57.
    const tmp_path = "/tmp/zlsx_cli_legacy_list_sheets.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};
    {
        const writer = @import("writer.zig");
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s0 = try w.addSheet("Data");
        try s0.writeRow(&.{.{ .string = "x" }});
        var s1 = try w.addSheet("More");
        try s1.writeRow(&.{.{ .string = "y" }});
        try w.save(tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, tmp_path);
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
    const tmp_path = "/tmp/zlsx_cli_comments_iter58.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};
    {
        const writer = @import("writer.zig");
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s0 = try w.addSheet("Data");
        try s0.writeRow(&.{.{ .string = "hdr" }});
        try s0.addComment("A1", "Alice", "needs review");
        var s1 = try w.addSheet("Other");
        try s1.writeRow(&.{.{ .integer = 1 }});
        try s1.addComment("B2", "Bob", "hi");
        try w.save(tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, tmp_path);
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
    // The writer's comment body wraps text in `<r><t>` (even for
    // plain bodies), so the reader populates `runs` as a one-entry
    // array of `{text:"…"}`. `runs:null` would require an `<r>`-less
    // body, which the writer doesn't emit today — exercise the
    // populated-runs path instead.
    try std.testing.expect(std.mem.indexOf(u8, out, "\"runs\":[{\"text\":\"needs review\"}]") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"sheet\":\"Other\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"ref\":\"B2\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"row\":2,\"col\":2") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "\"author\":\"Bob\"") != null);
}

test "runValidationsCommand emits list validation with values array" {
    const tmp_path = "/tmp/zlsx_cli_validations_iter58.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};
    {
        const writer = @import("writer.zig");
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s0 = try w.addSheet("Data");
        try s0.writeRow(&.{.{ .string = "fruit" }});
        try s0.addDataValidationList("B2:B100", &.{ "apple", "banana" });
        try w.save(tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, tmp_path);
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
    const tmp_path = "/tmp/zlsx_cli_hyperlinks_iter58.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};
    {
        const writer = @import("writer.zig");
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s0 = try w.addSheet("Data");
        try s0.writeRow(&.{.{ .string = "site" }});
        try s0.addHyperlink("A2", "https://example.com/");
        try w.save(tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, tmp_path);
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
    const tmp_path = "/tmp/zlsx_cli_styles_iter58.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};
    {
        const writer = @import("writer.zig");
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        _ = try w.addStyle(.{ .font_bold = true });
        var s0 = try w.addSheet("Data");
        try s0.writeRow(&.{.{ .string = "hdr" }});
        try w.save(tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, tmp_path);
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
    const tmp_path = "/tmp/zlsx_cli_sst_iter58.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};
    {
        const writer = @import("writer.zig");
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s0 = try w.addSheet("Data");
        try s0.writeRow(&.{ .{ .string = "header" }, .{ .string = "qty" } });
        try s0.writeRow(&.{ .{ .string = "apple" }, .{ .integer = 3 } });
        try w.save(tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, tmp_path);
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
    const env = std.process.getEnvVarOwned(std.heap.page_allocator, "XLSX_FUZZ_ITERS") catch return 1_000;
    defer std.heap.page_allocator.free(env);
    var digits: [32]u8 = undefined;
    var di: usize = 0;
    for (env) |c| {
        if (c == '_') continue;
        if (di == digits.len) break;
        digits[di] = c;
        di += 1;
    }
    return std.fmt.parseInt(usize, digits[0..di], 10) catch 1_000;
}

fn fuzzSeedCli() u64 {
    if (std.process.getEnvVarOwned(std.heap.page_allocator, "XLSX_FUZZ_SEED")) |s| {
        defer std.heap.page_allocator.free(s);
        return std.fmt.parseInt(u64, s, 10) catch 0xA1F8ED;
    } else |_| {
        return @bitCast(std.time.milliTimestamp());
    }
}

test "fuzz colLetter: output is uppercase A-Z" {
    const iters = fuzzItersCli();
    var prng = std.Random.DefaultPrng.init(fuzzSeedCli());
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
    const iters = fuzzItersCli();
    var prng = std.Random.DefaultPrng.init(fuzzSeedCli());
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
    const iters = fuzzItersCli();
    var prng = std.Random.DefaultPrng.init(fuzzSeedCli());
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
    const iters = fuzzItersCli();
    var prng = std.Random.DefaultPrng.init(fuzzSeedCli());
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
    const tmp_path = "/tmp/zlsx_cli_all_sheets_iter59c.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};
    {
        const writer = @import("writer.zig");
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s0 = try w.addSheet("Alpha");
        try s0.writeRow(&.{.{ .string = "A1_alpha" }});
        var s1 = try w.addSheet("Beta");
        try s1.writeRow(&.{.{ .string = "A1_beta" }});
        try w.save(tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, tmp_path);
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
    const tmp_path = "/tmp/zlsx_cli_glob_iter59c.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};
    {
        const writer = @import("writer.zig");
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s0 = try w.addSheet("Sheet1");
        try s0.writeRow(&.{.{ .string = "v1" }});
        var s1 = try w.addSheet("Sheet2");
        try s1.writeRow(&.{.{ .string = "v2" }});
        var s2 = try w.addSheet("Data3");
        try s2.writeRow(&.{.{ .string = "v3" }});
        try w.save(tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, tmp_path);
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
    const tmp_path = "/tmp/zlsx_cli_cross_pag_iter59c.xlsx";
    defer std.fs.cwd().deleteFile(tmp_path) catch {};
    {
        const writer = @import("writer.zig");
        var w = writer.Writer.init(std.testing.allocator);
        defer w.deinit();
        var s0 = try w.addSheet("A");
        // 3 cells on sheet 0 → a, b, c.
        try s0.writeRow(&.{ .{ .string = "a" }, .{ .string = "b" }, .{ .string = "c" } });
        var s1 = try w.addSheet("B");
        // 3 cells on sheet 1 → d, e, f.
        try s1.writeRow(&.{ .{ .string = "d" }, .{ .string = "e" }, .{ .string = "f" } });
        try w.save(tmp_path);
    }

    var book = try xlsx.Book.open(std.testing.allocator, tmp_path);
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
