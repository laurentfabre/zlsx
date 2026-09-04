# C ABI v1, part 1 (M9a1) — `zlsx_status_v1`, descriptors, editor recalc/evaluate

The committed design note §12.3 requires: exact field layouts, the status
contract, ownership rules, and the staging decisions, reviewed before the
code that implements them. This document is normative for the M9a1 exports;
`include/zlsx.h` and `src/c_abi.zig` implement it and pin every offset
below with comptime tests.

## 1. Scope: what lands at M9a1, what is staged to M9a2

**M9a1 exports** (each lands in header + impl + `_ffi.py` together — the
3-file transaction):

| Export | Probe (`_ffi.py`) | Header macro |
|---|---|---|
| `zlsx_engine_fingerprint` | `_HAS_FINGERPRINT` | `ZLSX_HAS_FINGERPRINT` |
| `zlsx_editor_mark_recalc_on_load` | `_HAS_MARK_RECALC` | `ZLSX_HAS_MARK_RECALC` |
| `zlsx_editor_recalculate` + `zlsx_recalc_report_release` | `_HAS_RECALC` | `ZLSX_HAS_RECALC` |
| `zlsx_editor_evaluate` + `zlsx_value_release` | `_HAS_EVAL` | `ZLSX_HAS_EVAL` |
| `zlsx_cancel_token_new` / `_trigger` / `_free` | `_HAS_CANCEL` | `ZLSX_HAS_CANCEL` |
| `zlsx_diag_release` | (rides `_HAS_RECALC`/`_HAS_EVAL`/`_HAS_MARK_RECALC`) | — |

**Staged to M9a2**: `zlsx_editor_save_with_recalc`, `zlsx_editor_save_to_buffer`,
`zlsx_open_buffer`, the `…_with_formulas_v2` writer export, `zlsx_buffer_release`,
and every high-level Python method (`Editor.recalculate` etc. — the worker-thread
cancellation mechanics live there). M9a1's binding leg is the `_ffi.py` symbol
declarations, structures, and probes; `_HAS_SAVE_BUFFER`, `_HAS_SAVE_WITH_RECALC`,
`_HAS_WRITER_RECALC`, `_HAS_FORMULAS_V2` stage with their exports.

**Why the cancel token is part 1** (decision M9a1-1): R9-12 hard-sets this
module multi-threaded *because* `-fsingle-threaded` atomics cannot back a
cross-thread token — a decision that is only about something that exists.
The recalc/evaluate exports take the token parameter, and a parameter no
caller can construct is untestable; the three token exports are two
allocations and a store.

## 2. `zlsx_status_v1` — the numeric contract

New exports only. Legacy exports (everything shipped before M9a1,
e.g. the `0/-1` writer/editor surface around `c_abi.zig:1927-1953`) keep
their `0/-1` contract verbatim; an old binding against a new library sees
identical behaviour on every symbol it knows.

| Code | Name | Meaning |
|---|---|---|
| `0` | `ZLSX_OK` | success |
| `-1` | `ZLSX_ERROR` | generic error; message (Zig error name) in caller's `errbuf` |
| `-2` | `ZLSX_REFUSED` | typed Plane-2 refusal; `zlsx_diag_v1` populated when supplied. Includes `FormulaLimitExceeded` — limits are Plane-2 refusals at every layer |
| `-3` | `ZLSX_NOMEM` | genuine allocation failure |
| `-4` | — | reserved (never returned by v1) |
| `-5` | `ZLSX_CANCELLED` | cancellation/deadline observed before commit |

**Zig error → status mapping** (decision M9a1-2, one mapping, no per-layer
reclassification): `error.OutOfMemory → -3`; `error.Cancelled → -5`; every
`Formula*` member of `pkg/workbook.Error` (the fourteen-plane §10 vocabulary)
`→ -2` with diag populated; everything else `→ -1` with `@errorName` in
`errbuf`. Input-contract violations raised at the ABI boundary itself
(NULL where required, `struct_size` below minimum, unknown enum value,
`UtcOffsetOutOfRange`) are `-1`: they are statements about the call, not
about the workbook.

Every fallible export takes `(char *errbuf, size_t errbuf_len)` (existing
convention — `errbuf` nullable, truncating write, always NUL-terminated)
AND, where a refusal is possible, a nullable `zlsx_diag_v1 *`.

## 3. `struct_size` rules

Every caller-allocated struct that crosses the boundary starts with
`size_t struct_size`, set by the caller to the size it was compiled
against.

- Readers use `min(caller_size, known_size)`; fields beyond the caller's
  declared size take their documented defaults.
- Writers touch only `min(caller_size, known_size)` bytes — **bytes beyond
  the known prefix are never written**, even when `caller_size` is larger
  (canary-tail tests pin this).
- The `struct_size` field itself is caller-owned: the library never writes
  it, in particular zero-initialisation skips it.
- Output structs are zero-initialised within `[sizeof(size_t), known_size)`
  on entry, before any failure can return. Prep is per-struct and runs
  before input validation: a failure anywhere — including a *sibling*
  struct's rejected `struct_size` — leaves every accepted output zeroed
  and therefore releasable. A rejected struct is left byte-for-byte
  untouched, and releasing an untouched (never-zeroed) struct is the
  caller's UB.
- Minimum accepted size = offset+size of the last v1 field, i.e. the full
  v1 `sizeof`. Below that: `-1`, message `StructSizeTooSmall`, no bytes
  written beyond the zero-init that never happened (the check precedes it).
- v1 structs never shrink; future fields append. Nested structs
  (`zlsx_resolved_v1` inside the report) do not get independent prefix
  treatment — the outer struct's rule governs; their `struct_size` is
  written by the library as documentation of the layout it filled.

## 4. Field layouts (LP64 / LLP64; offsets identical on both)

All structs are natural-alignment C structs; every offset is pinned by a
comptime test in `src/c_abi.zig`. Pointers/`size_t` are 8 bytes on every
supported target.

### 4.1 `zlsx_census_entry_v1` — 16 bytes

| off | type | field | notes |
|---|---|---|---|
| 0 | `uint32_t` | `plane` | `ZLSX_PLANE_*` |
| 4 | `uint32_t` | `sheet` | 0-based |
| 8 | `uint32_t` | `row` | 1-based; 0 = not about a cell |
| 12 | `uint32_t` | `col` | 0-based |

### 4.2 `zlsx_diag_v1` — 96 bytes

| off | type | field | notes |
|---|---|---|---|
| 0 | `size_t` | `struct_size` | caller sets |
| 8 | `uint32_t` | `plane` | leading refusal's plane; `ZLSX_PLANE_NONE` (`0xFFFFFFFF`) when absent |
| 12 | `uint32_t` | `census_truncated` | 0/1 |
| 16 | `char[64]` | `error_name` | NUL-terminated Zig error name (`"FormulaCycle"`) |
| 80 | `const zlsx_census_entry_v1 *` | `census` | library-owned; NULL iff `census_len == 0` |
| 88 | `size_t` | `census_len` | |

Released with `zlsx_diag_release` (frees `census`, resets it to
NULL/0; the struct itself is the caller's). Safe on a diag that was never
populated (zero-init leaves `census == NULL`). Census entries cross only
on paths that produce a census today: a successful mark-eligible
recalc report. On a refusal the pkg pipeline collapses the census into
the error before the ABI sees it (`recalc_txn.Refusal` carries
reason+plane only), so a `-2` diag carries `error_name` + `plane` with
an empty census — surfacing the refusing cells through
`prepare` is flagged for M9a2's Python `ZlsxFormulaRefusal(cells, census)`
leg (decision M9a1-3).

### 4.3 `zlsx_run_v1` — 104 bytes (input)

| off | type | field | notes |
|---|---|---|---|
| 0 | `size_t` | `struct_size` | |
| 8 | `int64_t` | `now_utc_ms` | no default — §5.5, a library that defaulted it would read a clock |
| 16 | `uint64_t` | `rng_seed` | no default, same reason |
| 24 | `int32_t` | `utc_offset_min` | validated in [-1440, 1440] pre-narrowing |
| 28 | `uint32_t` | `fidelity` | `ZLSX_FIDELITY_*`; unknown → -1 |
| 32 | `uint32_t` | `profile` | `ZLSX_PROFILE_*` |
| 36 | `uint32_t` | `dialect` | `ZLSX_DIALECT_*`; standalone eval only, recalc reads stored cells |
| 40 | `uint32_t` | `on_unsupported` | `ZLSX_ON_UNSUPPORTED_*`; recalc only |
| 44 | `uint32_t` | `_reserved0` | ignored in v1 |
| 48 | `uint64_t` | `max_run_arena_bytes` | 0 = default (§9 numeric defaults) |
| 56 | `uint64_t` | `max_matrix_cells` | 0 = default |
| 64 | `uint64_t` | `max_string_payload_bytes` | 0 = default |
| 72 | `uint64_t` | `max_retained_ast_bytes` | 0 = default |
| 80 | `uint64_t` | `max_diagnostics_bytes` | 0 = default |
| 88 | `uint64_t` | `timeout_ms` | 0 = none; converted to an absolute `.awake` deadline at entry |
| 96 | `zlsx_cancel_token_t *` | `cancel` | nullable = non-cancellable |

`deadline`/`cancel` stay outside the fingerprintable projection by
construction (`EffectiveRunInputs` has no such fields).

### 4.4 `zlsx_resolved_v1` — 80 bytes (§5.5 echo)

| off | type | field |
|---|---|---|
| 0 | `size_t` | `struct_size` |
| 8 | `int64_t` | `now_utc_ms` |
| 16 | `uint64_t` | `rng_seed` |
| 24 | `int32_t` | `utc_offset_min` |
| 28 | `uint32_t` | `fidelity` |
| 32 | `uint32_t` | `profile` |
| 36 | `uint32_t` | `dialect` — `ZLSX_DIALECT_NONE` (`0xFFFFFFFF`) for a recalc, which derives dialect per stored cell |
| 40..79 | `uint64_t ×5` | the five §9 limits, resolved (defaults echoed as numbers, never as 0) |

### 4.5 `zlsx_recalc_report_v1` — 168 bytes

| off | type | field |
|---|---|---|
| 0 | `size_t` | `struct_size` |
| 8 | `uint32_t` | `sheets_patched` |
| 12 | `uint32_t` | `cells_written` |
| 16 | `uint32_t` | `passes` |
| 20 | `uint32_t` | `non_converged_cells` |
| 24 | `uint32_t` | `dynamic_passes` |
| 28 | `uint32_t` | `kept_stale` (0/1) |
| 32 | `uint32_t` | `calc_chain_removed` (0/1) |
| 36 | `uint32_t` | `census_truncated` (0/1) |
| 40 | `uint64_t` | `retained_generations` |
| 48 | `uint64_t` | `retained_bytes` |
| 56 | `uint32_t` | `durability_warning` (dormant §5.7.9 slot — 0 for the in-memory transaction; M9a2's save fills it) |
| 60 | `int32_t` | `durability_errno` |
| 64 | `zlsx_resolved_v1` | `resolved` (embedded by value) |
| 144 | `uint32_t` | `resolved_present` (0/1) |
| 148 | `uint32_t` | `_reserved0` |
| 152 | `const zlsx_census_entry_v1 *` | `census` (library-owned) |
| 160 | `size_t` | `census_len` |

Released with `zlsx_recalc_report_release` (frees `census`).

### 4.6 `zlsx_value_elem_v1` — 32 bytes (§12.3's descriptor, exactly)

| off | type | field | notes |
|---|---|---|---|
| 0 | `uint8_t` | `tag` | `ZLSX_VALUE_*` |
| 1 | `uint8_t[7]` | `_reserved` | explicit padding, always 0 |
| 8 | `double` | `num` | number value; bool as 0/1; 0 otherwise |
| 16 | `uint64_t` | `payload_off` | into the value's payload arena |
| 24 | `uint64_t` | `payload_len` | text bytes / error spelling; 0 for number+bool |

Tags: `ZLSX_VALUE_NUMBER 0`, `ZLSX_VALUE_TEXT 1`, `ZLSX_VALUE_BOOL 2`,
`ZLSX_VALUE_ERROR 3` (payload = Excel spelling, e.g. `#DIV/0!` — an error
*value* is a successful result, plane 1, never a status code). Blank never
crosses: §12.2's publish rule turns it into number 0 before the boundary.

### 4.7 `zlsx_value_v1` — 56 bytes

| off | type | field |
|---|---|---|
| 0 | `size_t` | `struct_size` |
| 8 | `uint32_t` | `rows` |
| 12 | `uint32_t` | `cols` |
| 16 | `uint32_t` | `is_matrix` (0 = scalar, rows=cols=1) |
| 20 | `uint32_t` | `_reserved0` |
| 24 | `const zlsx_value_elem_v1 *` | `elems` (library-owned, row-major, `rows*cols`) |
| 32 | `size_t` | `elems_len` |
| 40 | `const uint8_t *` | `payload` (library-owned, one arena) |
| 48 | `size_t` | `payload_len` |

References are dereferenced before crossing (the evaluator already
guarantees `.reference` cannot reach a result). Released with
`zlsx_value_release` (frees `elems` + `payload`).

## 5. Export signatures

```c
const char *zlsx_engine_fingerprint(void);            /* static storage, never NULL */

int32_t zlsx_cancel_token_new(zlsx_cancel_token_t **out, char *errbuf, size_t errbuf_len);
void    zlsx_cancel_token_trigger(zlsx_cancel_token_t *tok);   /* thread-safe, any thread */
void    zlsx_cancel_token_free(zlsx_cancel_token_t *tok);      /* NULL-safe */

int32_t zlsx_editor_mark_recalc_on_load(zlsx_editor_t *ed,
        zlsx_diag_v1 *diag, char *errbuf, size_t errbuf_len);

int32_t zlsx_editor_recalculate(zlsx_editor_t *ed, const zlsx_run_v1 *run,
        zlsx_recalc_report_v1 *report, zlsx_diag_v1 *diag,
        char *errbuf, size_t errbuf_len);

int32_t zlsx_editor_evaluate(zlsx_editor_t *ed,
        const uint8_t *formula_ptr, size_t formula_len,
        uint32_t sheet_idx,
        uint32_t anchor_row,   /* 1-based; 0 = no anchor (site-dependent formulas refuse) */
        uint32_t anchor_col,   /* 0-based; read only when anchor_row != 0 */
        const zlsx_run_v1 *run,
        zlsx_value_v1 *out_value, zlsx_resolved_v1 *out_resolved /* nullable */,
        zlsx_diag_v1 *diag, char *errbuf, size_t errbuf_len);

void zlsx_value_release(zlsx_value_v1 *v);
void zlsx_recalc_report_release(zlsx_recalc_report_v1 *r);
void zlsx_diag_release(zlsx_diag_v1 *d);
```

- `report`, `out_value` are required (`NULL` → -1): an evaluation whose
  result cannot land anywhere is a call that meant nothing.
- `formula_len > 32768` (the parser's `max_formula_utf8_bytes`) is refused
  `-2 FormulaLimitExceeded` **before** the slice is formed — the boundary
  never constructs a slice it hasn't bounded.
- `zlsx_editor_evaluate` is M6's `zlsx eval` semantics over the same
  `Workbook.evaluate`: a cache-based read; the standalone path draws
  nothing volatile (the engine's fixed-draw seam), `rng_seed` is echoed in
  `resolved` for reproducibility of the context, and recalc is where a
  seed actually draws (decision M9a1-4 — matches the shipped M6 CLI
  behaviour exactly).
- Cancellation is observed-before-commit. `evaluate` never commits;
  `recalculate`'s commit point is the in-memory swap, and the pipeline
  polls the same token/deadline the CLI path does. A trigger observed
  pre-commit returns `-5` with memory untouched.

## 6. `zlsx_engine_fingerprint()`

```
zlsx <semver>; excel_fp_rules_v1; rng_v1; collation_v1; <arch>-<os>-<abi>; <build-hash>
```

- semver from `build.zig.zon` (the bare `c_abi.zig:51-54` version string is
  a component, not the identity);
- the three rule versions are read from the engine at comptime
  (`pkg/recalc_run.zig` re-exports them as `rule_versions` — the string
  cannot drift from the code that implements the rule);
- target triple from `builtin.target`;
- build hash: `-Dbuild-hash=<str>` override, else `git rev-parse --short=12 HEAD`
  at build time, else `"unknown"` (source tarballs must still build). It
  lives in its own options module (`fingerprint_config`) imported only by
  the C ABI module, so a new commit does not invalidate the CLI's build
  cache.

M9b keys Spark's engine-identity refusal on this string; mismatch refuses.

## 7. R9-12: the C ABI module is multi-threaded, always

`build.zig` stops forwarding `-Dsingle-threaded` into `c_abi_mod`
(`.single_threaded = false`, stated); the option's description narrows to
the CLI. `src/c_abi.zig` carries a comptime assertion refusing to compile
single-threaded — under `-fsingle-threaded` Zig lowers atomics to plain
ops and the cross-thread cancel token would be silently broken in exactly
the supported configuration. The CLI keeps its signal-safe `flag` token
kind. Both configurations stay CI-compiled (`-Dsingle-threaded=true` lane
builds the CLI single-threaded and the ABI multi-threaded from the same
invocation). The `gpa` fallback to `page_allocator` for single-threaded
builds is dead by construction and removed.

## 8. Test matrix (narrowing / canary / boundary / fuzz)

| Gate | What it pins |
|---|---|
| comptime offset asserts | every offset in §4, `@sizeOf` of every struct |
| plane-value pinning test | `ZLSX_PLANE_*` ↔ `@intFromEnum(PlaneTwo.*)` for all fourteen + `NONE` |
| narrowing tests | unknown `fidelity`/`profile`/`dialect`/`on_unsupported` → -1; `utc_offset_min` ±1441 → -1 (`UtcOffsetOutOfRange`); `sheet_idx = UINT32_MAX` → -1; `formula_len` one-past-limit → -2 `FormulaLimitExceeded` |
| boundary tests | limit fields at 0 (default), `UINT64_MAX` (rejected by `ResourceLimits.validate`'s 4× hard ceiling → -1 `LimitOutOfRange`) |
| canary-tail tests | output structs with `struct_size = known + 64`, tail filled `0xAA` — tail intact after success AND after each failure class |
| `StructSizeTooSmall` | every struct-taking export, `struct_size = known - 1` → -1, output untouched |
| status tests | 0 / -1 / -2 / -3(via failing allocator N/A at ABI — covered by refusal+generic classes) / -5 (pre-triggered token) |
| release fns | idempotent-shape: release, then release the zeroed struct again (no-op); NULL-safe |
| old-binding compat | legacy exports' 0/-1 behaviour unchanged (existing suite already pins); probes absent on old dylibs = feature off (`hasattr` gate in `_ffi.py`) |
| ABI fuzz | random `zlsx_run_v1` bytes (enums, limits, struct_size), random formula bytes, random anchors against a live editor: status ∈ {0,-1,-2,-3,-5}, no panic, canary tails intact — `fuzzItersCabi()` iterations |
| header compile gate | `tests/c_abi_smoke.c` compiled as an object in `zig build test`: `#error` unless every `ZLSX_HAS_*` macro is defined; takes the address of every M9a1 export |

## 9. Decisions index (mirrored into `goal_formula.md`'s M9a1 block)

1. Cancel token API ships in part 1 (R9-12 is about it; the token
   parameter is otherwise untestable).
2. One error→status mapping; ABI-boundary contract violations are -1.
3. Refusal-path census stays empty in v1 diag; pkg-level surfacing flagged
   for M9a2's Python leg.
4. `zlsx_editor_evaluate` = M6 CLI semantics (fixed draw source, seed
   echoed); volatile-drawing evaluate is not a thing any shipped layer has.
5. `zlsx_buffer_release` stages with the buffer exports (M9a2).
6. Build hash in a dedicated `fingerprint_config` options module; CLI
   build cache unaffected by commits.

## 10. S3a — structural edits + the `pivots` read (2026-08-30)

The `Editor`'s structural edits and S6's `pivots` NDJSON shape cross the
boundary under the §2 contract — new exports, so `zlsx_status_v1`, never
the legacy `0/-1`. Nine exports, one header macro pair, one probe group:

| Export | Probe (`_ffi.py`) | Header macro |
|---|---|---|
| `zlsx_editor_insert_row` / `_delete_row` / `_insert_column` / `_delete_column` | `_HAS_STRUCTURAL_EDITS` | `ZLSX_HAS_STRUCTURAL_EDITS` |
| `zlsx_editor_add_sheet` / `_rename_sheet` / `_delete_sheet` | `_HAS_STRUCTURAL_EDITS` | `ZLSX_HAS_STRUCTURAL_EDITS` |
| `zlsx_editor_rename_table_column` | `_HAS_STRUCTURAL_EDITS` | `ZLSX_HAS_STRUCTURAL_EDITS` |
| `zlsx_editor_pivots_ndjson` (+ `zlsx_buffer_release`) | `_HAS_PIVOTS` | `ZLSX_HAS_PIVOTS` |

Each probe also requires the release symbols the Python wrappers call
unconditionally: `_HAS_STRUCTURAL_EDITS` needs `zlsx_diag_release`,
`_HAS_PIVOTS` needs it and `zlsx_buffer_release`, each probed and
configured on its own rather than borrowed from an unrelated feature's
block (r6 REL-601).

**Naming** (the S3a–e gate question): the M9a2 precedent holds — a new
export takes its bare name under the status contract; the `_v2` suffix
is reserved for a name a legacy export already owns
(`zlsx_sheet_writer_write_row_with_formulas_v2`). None of the S3a names
collide, so none carry a suffix. No new struct crosses: the exports take
scalars, `(ptr, len)` byte strings and the existing `zlsx_diag_v1`.

**The refusal vocabulary** (decision S3a-1). §2 maps `Formula*` planes to
`-2`; the structural edits have a vocabulary of their own, and the same
rule applies — a refusal is a statement about the *workbook*, a failure
a statement about the *call*. `statusOf` checks one more list after the
fourteen planes (`c_abi.zig::structural_refusals`); the diag carries the
name with `plane = ZLSX_PLANE_NONE` and an empty census:

| `-2` (`ZLSX_REFUSED`, `diag.error_name`) | `-1` (`ZLSX_ERROR`, `errbuf`) |
|---|---|
| `RowEditUnsafeForSheet`, `ColEditUnsafeForSheet` — the editor's fold of its pre-flights: inside a hosted pivot's footprint, a host a pivot also reads, a table collapse / header-row delete, an `<xm:f>` carrier the scan cannot read, a chart `<c:f>` carrier the walk cannot read whole | `SheetIndexOutOfRange`, `RowIndexOutOfRange`, `ColumnIndexOutOfRange` (a 0-based `UINT32_MAX` column has no 1-based spelling and is refused before the conversion) |
| The worksheet transform's own verdicts, from its pre-mutation probe: `RowEditExceedsMaxRow`, `ColEditExceedsMaxCol`, `SplitPaneNotSupported`, `MalformedPaneSplit`, `MalformedSheetXml` | |
| The sweeps' — a carrier the walkers cannot read or move: `MalformedDrawingXml`, `DrawingCoordinateOverflow`, `MalformedVmlDrawing`, `VmlCoordinateOverflow`, `MalformedCommentsXml`, `MalformedTableXml`, `TableCoordinateOverflow`, `TableCollapseUnsafe`, `TableHeaderRowDeleteUnsafe`, `SqrefCollapseUnsafe` (a delete collapsing EVERY area of a DV/CF `sqref` — S3b slice 6 r4), `PivotEditUnsafe`, `MalformedExtensionXml`, `MalformedChartXml` (a chart `<c:f>` series carrier the sweep cannot read whole — the chart sweep, the `MalformedExtensionXml` shape: refused before the first mutation, folded into the two `*UnsafeForSheet` names on row / column edits), `MalformedSheetRels`, `MalformedWorkbookRels`, `MalformedDrawingRels`, `MissingSheetPart`, `NoSheetData` (and the workbook layer's `LastSheetUndeletable` / `SheetNameInUse`, should a path surface them unfolded) | |
| `CannotDeleteLastSheet` | `InvalidSheetName`, `InvalidTableColumnName` — Excel would not take the name |
| `DuplicateSheetName` (add / rename; ASCII case-insensitive, footnote ¹⁴ of the surface matrix), `TableColumnNameInUse` — a name the workbook holds | `TableNotFound`, `TableColumnNotFound` — a selector that names nothing among the workbook's readable tables (decision S3a-9); a `<tablePart>` whose relationship, part or display name is broken refuses instead (`MissingRelationship` / `MalformedTableXml`, r6 REL-604) |
| The workbook's own structure, found broken on the way: `InternalSheetNameTooLong` (a stored sheet name that is EMPTY — an OOXML-invariant violation no argument fixes; the historical 128-byte carrier bound fell in #216 r17, a valid escape-heavy name legitimately exceeds it), `MalformedWorkbookXml` (`xl/workbook.xml` without the `</sheets>` a splice needs), `IdSpaceExhausted` (a `sheetId`, `rId` or worksheet part number already at `UINT32_MAX` — checked arithmetic, never a trap), `MissingRelationship`, `SheetElementNotFound`, `RelationshipElementNotFound`, `SheetCountMismatch`, `MissingWorkbookPart`, `MissingWorkbookRels`, `MissingContentTypes`, `MalformedContentTypes`, `ContentTypesOverrideNotFound` | `InvalidInput` — NULL where bytes are required (NULL with length 0 is the empty string, judged by the editor) |
| | `RowEditRequiresCleanSheet`, `ColEditRequiresCleanSheet`, `SheetDeleteRequiresCleanState` — sequencing: the sheet (the workbook, for a sheet delete) has staged cell writes or appended rows; save first |
| `MalformedPivotXml` — the pivot graph cannot be read whole | `NullOutPointer`, `StructSizeTooSmall` |

The list is `c_abi.zig::structural_refusals`, one place. The editor folds
its own pre-flights into the two `*UnsafeForSheet` names; what the
transform's probe and the later sweeps raise reaches the boundary
unfolded, so a caller sees the precise cause (`RowEditExceedsMaxRow`
rather than "unsafe") — Codex #207 r1 REL-102. Sequencing errors and
the transform's post-probe contract (a sweep failing after the probe
passed: discard the editor, reopen — `Workbook.applySheetEditTransform`)
are unchanged by this row.

The Python leg mirrors the split: `-2` raises `ZlsxRefusal(error_name)`
(the new base class; `ZlsxFormulaRefusal` now derives from it and is
chosen when the diag names a plane), `-1` raises a plain `ZlsxError`.

**Coordinates**: rows 1-based, columns 0-based (A = 0) — what
`zlsx_editor_set_cell` and `zlsx_census_entry_v1` already spell; the Zig
editor's 1-based column API is converted at the boundary. Sheet indices
0-based; `zlsx_editor_add_sheet` writes `UINT32_MAX` to `*out_sheet_idx`
on entry and the new index on `ZLSX_OK`.

**`zlsx_editor_pivots_ndjson`** (decision S3a-2): the S6 gate froze a
record *shape*, and the shape is text. The export hands over the NDJSON
bytes themselves — `pkg/pivot_ndjson.zig`, the one writer `zlsx pivots`
prints through too — in a library-allocated buffer released with
`zlsx_buffer_release`; a workbook without pivots is `ZLSX_OK` with
`(NULL, 0)`. A typed struct graph would have been a second spelling of a
frozen contract, drifting on its own schedule. Python parses one JSON
object per line (`Editor.pivots()` / `zlsx.pivots(path)`); a C caller
parses with whatever it has. The read runs over the editor's current
workbook state — structural edits visible immediately, staged
`set_cell` / `append_row` writes only after `save` refreshes or marks
the caches they touch (r7 REL-701 narrowed the claim; the alternative,
a non-mutating overlay read, is S8-adjacent machinery this row does
not need). The buffer is built by
`c_abi.zig::pivotsNdjsonOwned(alloc, wb)`: the allocating writer reports
a failed growth as `WriteFailed`, which the builder maps to
`OutOfMemory` so the boundary answers `-3`, never a generic `-1`
(Codex #207 r1 REL-103; pinned by an allocation-failure sweep).

**Two more mappings** (decisions S3a-6, S3a-7): a worksheet part the
typed parser cannot read — reached lazily by a sheet delete and by the
sweeps — is the sheet transform's own `MalformedSheetXml`, spelled so at
the one call site (`Worksheet.ensureParsed`) rather than renamed at the
boundary, where a generic `MalformedXml` from any other part would have
been mislabelled (r3 REL-302); a pivot part the store cannot materialise
(a bad CRC, a broken stream) makes `zlsx_editor_pivots_ndjson` refuse as
`MalformedPivotXml`, with `OutOfMemory` and `ZipBombSuspected` keeping
their own statuses.

**`addSheet` is all-or-nothing** (decision S3a-8, r3 REL-305):
`Workbook.addSheet` builds both patched parts, parses the fresh view over
its own copy (`workbook_xml.parseOwning`) and grows the slot table before
the first store write, then adds the part and replaces the two references
as one atomic pair; `Editor.addSheet` computes the part name ahead
(`Workbook.nextSheetPartName`) and allocates its mirror before the
mutation, so nothing after it can fail. An allocation failure anywhere
leaves the workbook and the view as they were (pinned by
`checkAllAllocationFailures` over `Workbook.addSheet` in
`pkg/workbook.zig`; the editor's mirror is allocation-free after the
mutation by construction — the sweep cannot run through `Editor.open`,
whose reader path aborts under an injected failure, a pre-existing
reader defect noted for a follow-up); the one residue is an unreferenced
empty part when `replaceParts` fails after `addPart` — `deleteSheet`'s
orphan trade-off.

**Known hole, inherited:** `Editor.renameSheet` / `deleteSheet` do not
rewrite a pivot cache's `worksheetSource@sheet` / `rangeSet@sheet`; a
source spelled by sheet name goes stale and the pivots read reports it
as `"resolved":null`. Pre-existing in the Zig editor and the CLI; the
exports carry it unchanged, the header and the Python docs say so, and
the lift belongs to the S7 family on every surface at once.

**Tests** (`src/c_abi.zig`, "S3a …"): a boundary round trip whose saved
grid is checked in the sheet part; the editor's verdicts, the
transform's (`RowEditExceedsMaxRow`, `SplitPaneNotSupported`), the lazy
reads' (`MalformedSheetXml` from a sheet delete, `MalformedPivotXml` from
a corrupted part) and the workbook's own (`InternalSheetNameTooLong`)
each driven through the exports with the name in both the diag and
`errbuf` and `plane == NONE`, plus every member of `structural_refusals`
through `failMapped`; the `-1` classes driven through the exports —
`SheetIndexOutOfRange`, `RowIndexOutOfRange`, `ColumnIndexOutOfRange`,
`InvalidInput`, `InvalidSheetName`, `InvalidTableColumnName`,
`TableNotFound`, `TableColumnNotFound`, `RowEditRequiresCleanSheet`,
`ColEditRequiresCleanSheet`, `SheetDeleteRequiresCleanState`,
`NullOutPointer`, `StructSizeTooSmall` — with the diag left as prep left
it; `StructSizeTooSmall` leaving the diag byte-for-byte; a `0xAA` canary
tail across a generic failure and a refusal, before and after release; the pivots buffer equal to the package writer's frozen
record, `(NULL, 0)` on a plain workbook, `-2 MalformedPivotXml` on a
broken graph; `statusOf` pinned on both vocabularies. The header compile
gate (`tests/c_abi_smoke.c`) takes every S3a address and `#error`s
without either macro.

## 11. Decisions index (S3a)

1. Structural refusals are a second `-2` vocabulary next to the planes,
   with `plane = ZLSX_PLANE_NONE`; sequencing and argument errors stay
   `-1`.
2. The pivots read crosses as the frozen NDJSON bytes, not a struct
   graph; one writer serves the CLI and the ABI.
3. Bare names, no `_v2` — nothing collides.
4. Columns 0-based at the boundary, converted to the editor's 1-based
   API; `UINT32_MAX` refused before the conversion.
5. The transform's and the sweeps' workbook-safety errors cross with
   their precise names (one list, `structural_refusals`), not folded
   into `*UnsafeForSheet`; a growth failure in the pivots buffer is
   `-3`. The Python leg bounds every index to `[0, 2³²)` before ctypes
   narrows it (`ValueError`).
6. A worksheet part the typed parser cannot read is `MalformedSheetXml`
   at `Worksheet.ensureParsed` itself — every surface, one name.
7. A pivot part the store cannot materialise refuses the pivots read as
   `MalformedPivotXml`; `OutOfMemory` / `ZipBombSuspected` keep theirs.
8. `addSheet` is all-or-nothing under allocation failure, on the
   workbook and on the editor.
9. `TableNotFound` / `TableColumnNotFound` are selectors — `-1`, like a
   sheet index — not refusals; `TableColumnNameInUse` stays `-2`.
10. The `worksheetSource@sheet` hole under sheet rename / delete is
    inherited and documented, not lifted here.
11. `Worksheet.ensureParsed` spells a part the store cannot materialise
    (a CRC mismatch, a broken stream) as `MalformedSheetXml` too;
    `xl/workbook.xml` without `</sheets>` is `MalformedWorkbookXml`. The
    rewriters' own consistency guards ("the view no longer describes
    these bytes", `MalformedXml`) keep the generic name and `-1`.
12. `deleteSheet` allocates its surviving slot table (and the editor
    its mirror) before the first mutation; the commit after the fresh
    view parses cannot fail, so a handle that met an allocation failure
    still agrees with itself and closes (pinned by an allocation-failure
    sweep in `pkg/workbook.zig`). The sweeps' partial work under a
    failure stays the documented discard-and-reopen contract.
13. `zlsx_editor_close` tears down the handle's `std.Io.Threaded`, as
    `zlsx_book_close` always did.
14. A carrier part a structural path reads lazily — the `<xm:f>`
    pre-flight's whole-workbook sheet scan included (r6 REL-603) — drawing, VML,
    comments, table — that the store cannot materialise is that
    carrier's own verdict (`Workbook.carrierPart`: `MalformedDrawingXml`
    / `MalformedVmlDrawing` / `MalformedCommentsXml` /
    `MalformedTableXml` / `MalformedChartXml` — the chart sweep's
    preflight walks every chart part the same way), as the worksheet and
    pivot parts already were;
    `OutOfMemory` / `ZipBombSuspected` keep theirs. The three
    identifier increments of `addSheet` are checked
    (`IdSpaceExhausted`).
15. A `Worksheet.setCell` overwrite installs the replacement before it
    frees the displaced owned value (`getOrPut`, the one fallible
    step first): an allocation failure leaves the previous delta live
    and the handle closable (r6 REL-602; allocation-failure sweep in
    `pkg/workbook.zig`).

## 12. S3b slice 2 — the `defined-names` read (2026-09-01)

One export, the S3a pivots pattern verbatim:

| Export | Probe (`_ffi.py`) | Header macro |
|---|---|---|
| `zlsx_editor_defined_names_ndjson` (+ `zlsx_buffer_release`) | `_HAS_DEFINED_NAMES` | `ZLSX_HAS_DEFINED_NAMES` |

(`_HAS_DEFINED_NAME`, singular, probes the writer's
`zlsx_writer_add_defined_name` and predates this capability.)

**The bytes are the CLI's** (decision S3a-2 carried over): the S3b gate
froze the `defined-names` record in `docs/cli.md`, and the record is
text. The export hands over the NDJSON bytes of
`pkg/defined_name_ndjson.zig` — the shared writer `zlsx defined-names`
prints through — built by `c_abi.zig::definedNamesNdjsonOwned(alloc, wb)`
over the workbook's parsed `xl/workbook.xml` view: every name, document
order, no selector (the CLI's default stream). A workbook without
defined names is `ZLSX_OK` with `(NULL, 0)`; release with
`zlsx_buffer_release`. The allocating writer's `WriteFailed` crosses as
`-3`, the pivots builder's rule. Python parses one JSON object per line
(`Editor.defined_names()` / `zlsx.defined_names(path)`).

**One refusal, already in the vocabulary**: an inventory the read cannot
serve faithfully — a carrier that does not decode, malformed UTF-8
after decode, a body carrying embedded markup (the `docs/cli.md`
contract) — is `MalformedWorkbookXml`, `-2`, a member of
`structural_refusals` since S3a. No new name crosses.

**Timing, simpler than pivots**: defined names live in `xl/workbook.xml`
only, and the editor re-parses that view after every structural edit
(`refreshWorkbookXmlView`), so edits *and* the name sweeps they carry (a
sheet rename rewriting the bodies) are visible immediately; staged cell
writes never touch the part, so nothing about this read waits for save.

**Tests** (`src/c_abi.zig`, "S3b …"): the buffer equal to the shared
writer's frozen record over a fixture built through the C writer surface
(workbook scope, sheet scope, hidden); the rename-sweep rewrite visible
with no save; `(NULL, 0)` on a workbook without names with the poison
reset; `NullOutPointer` / `InvalidInput` as call errors; poisoned
outputs reset on the refusal and undersized-diag paths, the rejected
diag byte-for-byte untouched; `-2` `MalformedWorkbookXml` on a bad
entity with nothing handed out and the name in diag + errbuf; an
allocation-failure sweep over `definedNamesNdjsonOwned`. The smoke gate
takes the address and `#error`s without the macro; the Python leg pins
the parsed records, the refusal type and the closed-editor error.

## 13. S3b slice 6 — the `conditional-formats` read (2026-09-02)

One export, the slice-2 pattern verbatim:

| Export | Probe (`_ffi.py`) | Header macro |
|---|---|---|
| `zlsx_editor_conditional_formats_ndjson` (+ `zlsx_buffer_release`) | `_HAS_CONDITIONAL_FORMATS` | `ZLSX_HAS_CONDITIONAL_FORMATS` |

(`_HAS_CONDITIONAL_FORMAT`, singular, probes the writer's
`add_conditional_format_*` family and predates this capability.)

**The bytes are the CLI's**: the S3b gate froze the
`conditional-formats` record in `docs/cli.md`, and the record is text.
The export hands over the NDJSON bytes of
`pkg/conditional_format_ndjson.zig` — the shared strict walk + record
writer `zlsx conditional-formats` prints through — built by
`c_abi.zig::conditionalFormatsNdjsonOwned(alloc, wb)` over the editor's
current parts: every rule, sheets in workbook order, rules in
sheet-document order, no selector (the CLI's default stream). A workbook
without conditional formatting is `ZLSX_OK` with `(NULL, 0)`; release
with `zlsx_buffer_release`. The allocating writer's `WriteFailed`
crosses as `-3`, the pivots builder's rule. Python parses one JSON
object per line (`Editor.conditional_formats()` /
`zlsx.conditional_formats(path)`).

**The refusals, all already in the vocabulary**: a sheet list the
strict workbook read cannot prove is `MalformedWorkbookXml` — a part
the archive cannot materialise at the graph probe folds there too
(round 1, S3B-ERR-602: a bad CRC no longer escapes with the zip
layer's own name) — and a sheet part the strict walk cannot serve
faithfully (the `docs/cli.md` contract: nesting, namespace shape,
formula arity and markup, carriers that do not decode) is
`MalformedSheetXml`; both `-2`, members of `structural_refusals` since
S3a. `collect`'s error surface is the closed, compiler-checked
`CollectError`: those two plus `OutOfMemory` (`-3`),
`MissingSheetPart` (`-2`, typed but unreachable once the graph probe
holds), and `ZipBombSuspected` — the S1 decompression-caps verdict,
typed in the set for honesty but effectively open-time: the caps
admit every entry on the open-time directory walk (`src/cli.zig`,
`classifyTopLevelError`), so it fires where the ABI has no diag to
carry it — the pointer-returning path open and `zlsx_open_buffer`,
whose shipped contract is `-1`. Round 1 (S3B-ERR-601) remapped it to
`-2`; round 2 (S3B-ERR-702) ruled that an ABI break on the diag-less
opener and reverted it — it stays a DELIBERATE generic `-1` with its
name in errbuf, pinned by a `statusOf` test, until a status-bearing
open ABI exists. No new name crosses.

**Timing**: conditional formats live in the sheet parts, which
structural edits and their DV/CF sweeps rewrite in place before the
call returns, and the sheet inventory is a fresh strict read of the
current `xl/workbook.xml` — so edits (a rename renaming the `sheet`
field, an insert moving `sqref` and the formula bodies) are visible
immediately; staged cell writes never touch the rule machinery, so
nothing about this read waits for save. The `sqref` half of that
claim is round 1's own repair (S3B-REL-301): the worksheet transform
historically moved only the formula bodies — an insert rewrote
`B1>3` to `B2>3` while `sqref="B1:B4"` stayed, a skew every consumer
(Excel included, after save) could observe — and
`pkg/sheet_edit.zig::processSqrefListTag` now shifts
`conditionalFormatting@sqref` and `dataValidation@sqref` with the
merge-rect interval semantics — shift / grow / shrink, a collapsed
area dropped from the list, whole-column (`A:A`) and whole-row
(`1:1`) spellings absorbing the cross-axis edit and shifting as
intervals along their own, `$` anchors parsed and preserved, the
value entity-DECODED before parsing, BOTH axes validated by the
strict ST_Ref grammar — not just the edited one (rounds 3–4,
S3B-REL-802/803). An area the grammar refuses fails the whole edit
pre-mutation, the `<xm:f>` all-or-nothing posture — a frozen
envelope beside swept bodies is exactly the skew this repair closes
— and a delete that collapses EVERY area refuses too
(`SqrefCollapseUnsafe`, -2, the `TableCollapseUnsafe` shape): Excel
deletes the rule outright, the walker cannot excise an element with
children mid-walk, and kept bytes would silently retarget the rule
to whatever slides into its old coordinates (r4 S3B-REL-805).
Comments, CDATA sections, PIs and DOCTYPE constructs pass through
both walkers verbatim — a tag spelled inside one is prose, neither
rewritten nor refused on (r4 S3B-REL-804). The bodies
themselves: round 3 (S3B-REL-801) extended the DV/CF sweep from CF
formula slot 1 to all three schema slots, so a `cellIs` `between`
moves both its formulas with the envelope. Repeated reads stay
bounded: sheet targets resolve into the view's own arena
(`PartStore.resolveOwned`, S3B-MEM-603), not the store's lifetime
arena.

**Tests** (`src/c_abi.zig`, "S3b conditional_formats_ndjson …"): the
buffer equal to the shared writer's frozen stream (MNT-2302's literal)
over the pkg fixture rebuilt through the C writer surface (one dxf,
four rule kinds, an escaping-heavy expression); the rename AND a row
insert visible with no save — the insert pinning `sqref` and formula
on one grid (`A2:A5` beside `B2>3`), the other sheet unmoved; `(NULL,
0)` on a workbook without rules with the poison reset;
`NullOutPointer` / `InvalidInput` as call errors; poisoned outputs
reset on the refusal and undersized-diag paths, the rejected diag
byte-for-byte untouched; `-2` `MalformedSheetXml` on a bad entity in a
formula body with nothing handed out and the name in diag + errbuf;
the whole-refusal pins round 1 added (S3B-TEST-605): a broken SECOND
sheet behind four servable records refuses whole, a bad entity in a
sheet-name carrier is `MalformedWorkbookXml`, a flipped payload byte
folds at the graph probe, and `statusOf(ZipBombSuspected)` is the
deliberate `-1` (r2);
an allocation-failure sweep over `conditionalFormatsNdjsonOwned`
(caches primed first, the shared writer's own OOM-test shape); the
sqref-shift unit tests in `pkg/sheet_edit.zig` (insert above / inside,
delete inside, collapse-drop, the all-collapse refusal for CF and DV,
the column axis, the `dataValidations` wrapper decoy, whole-column /
whole-row absorption + interval shifting + collapse, `$`-anchor
round-trips, both-axes grid validation, entity decode with a
byte-preserving no-op, garbage refusal, comment / CDATA decoy
pass-through). The smoke gate takes the address and `#error`s without
the macro; the Python leg pins the parsed records, the insert-row
move, the `between` both-slots move, both refusal shapes and the
closed-editor error.

## 14. S3b slice 7 — the `anchors` read (2026-09-02)

One export, the slice-2 pattern verbatim:

| Export | Probe (`_ffi.py`) | Header macro |
|---|---|---|
| `zlsx_editor_anchors_ndjson` (+ `zlsx_buffer_release`) | `_HAS_ANCHORS` | `ZLSX_HAS_ANCHORS` |

**The bytes are the CLI's**: the S3b gate froze the `anchors` record in
`docs/cli.md`, and the record is text. The export hands over the NDJSON
bytes of `pkg/anchor_ndjson.zig` — the shared collector + record writer
`zlsx anchors` prints through — built by
`c_abi.zig::anchorsNdjsonOwned(alloc, wb)` over the editor's current
parts: every anchored image (`image_anchor`) and chart
(`chart_anchor`), sheets in workbook order, a sheet's images before
its charts, each class in drawing-document order, no selector (the
CLI's default stream). The record is the anchor geometry and where the
payload lives — never the payload: image bytes and chart XML stay in
their parts. A workbook without anchored objects is `ZLSX_OK` with
`(NULL, 0)`; release with `zlsx_buffer_release`. The allocating
writer's `WriteFailed` crosses as `-3`, the pivots builder's rule.
Python parses one JSON object per line (`Editor.anchors()` /
`zlsx.anchors(path)`).

**The refusals — one new name**: `DrawingOnUnlistedSheet` joins
`structural_refusals` (§10's vocabulary, `-2`): an anchored object on
a worksheet part `xl/workbook.xml` does not list, which no record
could attribute truthfully and no faithful inventory could drop. The
other two were already there: a sheet list the strict workbook read
cannot prove is `MalformedWorkbookXml` — the same strict inventory the
conditional-formats read established (§13), now SHARED as
`conditional_format_ndjson.resolveSheets` and used by the anchors
collector in place of its lenient `WorkbookXml.sheets` projection, so
the ghost-entry, carrier-less-entry, alternate-prefix, decoded-rid,
verified-rels, duplicate-decoded-name and graph-probe rules of #215 /
#216 hold for anchors too (a part the archive cannot materialise
folds there, S3B-ERR-602's shape) — and a drawing graph the strict
walk cannot read whole is `MalformedDrawingXml` (the `docs/cli.md`
contract: a dangling, mistyped or external edge, an absent part, an
anchor whose `from` / `to` / `pos` / `ext` does not parse, an
unclosed anchor or `<c:f>`, a second live `<drawing>`, a carrier that
does not decode; and, folded at the one place the walk is entered, a
drawing-family part the store cannot materialise — the walkers' error
surface is inferred, so the collector folds it to the closed set
rather than let the zip layer's own name cross). `collect`'s error
surface is the closed, compiler-checked `CollectError`: those three
plus `OutOfMemory` (`-3`) and `ZipBombSuspected` — typed for honesty,
effectively open-time, the deliberate `-1` §13 records. The export's
switch is exhaustive with no `else` (S3B-MNT-911's rule).

**Timing**: anchors live in the drawing parts, which structural edits
and their drawing sweeps rewrite in place before the call returns, and
the sheet inventory is a fresh strict read of the current
`xl/workbook.xml` — so a rename renames `sheet`, a row insert on the
image's sheet moves its `from` / `to` with the grid while the other
sheet's anchors stay, all visible immediately; staged cell writes never
touch a drawing, so nothing about this read waits for save. **The
chart `<c:f>` sweep (landed after slice 11, the S3b follow-up this
read made observable)**: a chart's series formulas — `<c:tx>`,
`<c:cat>`, `<c:val>` / `<c:xVal>` / `<c:yVal>` / `<c:bubbleSize>`
carriers — ride the formula rewriter under every structural edit
(`Workbook.rewriteAllChartFormulas`: the `<xm:f>` precedent — carrier
walk, decode, rewrite, byte-preserving splice, all-or-nothing refusal
by `Workbook.preflightChartFormulas` BEFORE the first mutation), so
`series_refs` spell the new sheet name after a rename, shifted rows
after an insert on the sheet they name, `#REF!` after that sheet's
delete (a qualified reference into it; a carrier already at
`Sheet!#REF!` — an error token, not a reference — keeps its qualifier
under a rename and a delete, the rewriter's single-qualifier rule on
every carrier, cell formulas included), and the rest of the chart part
is byte-identical. The parts walked are every part of the chart
content type, every `xl/charts/chart<N>.xml`, and every internal
chart-typed relationship target of any part — the edge this read
follows, from a drawing it reaches by the sheet's relationship alone
and need not recognise by name or content type either — so the two
enumerations cannot disagree on a chart the read serves. The
carrier walk (`drawings.ChartFormulaWalk`) and the body acceptance
(`Workbook.decodeChartFormulaBody`) are shared with this read, so what
the read serves is exactly what the sweep moves; a carrier the walk
cannot read whole — no close (a start tag truncated at the end of the
part included), markup in a body, a carrier start tag inside an
unterminated comment / CDATA / PI, a chart namespace bound anywhere in
the part under a prefix the walk does not follow — longer than the
resolver's 100-byte limit, a second prefix on a chart URI, or declared
beyond the resolver's 4 KiB root window (an `xmlns:` inside a comment,
a PI or an attribute value is text; the part would otherwise be walked
under a prefix it does not use and move nothing) — is this read's
`MalformedDrawingXml` and
the edit's `MalformedChartXml` (§10). Pivot charts ride too — their
carriers name the pivot's host cells and shift with the hosted
rectangle (S7a); a re-layout that changes the rectangle's extent
(S7b-5, S7c-2) is not re-derived into the chart, Excel's refresh-on-load
rebuilds the series — while `<c:pivotSource><c:name>` (a
`[book]sheet!pivot` locator, not a formula) keeps its spelling, as the
cache's `worksheetSource@sheet` does under a rename. A part binding
the chart namespace as its DEFAULT namespace (`<chartSpace
xmlns="…/chart"><f>`) is openpyxl's spelling and is walked under `<f>`
(in-house CF-REL-401 — it had been documented as unproduced and left
unwalked, so every openpyxl chart went stale silently). **The
namespace-aware drawing slice (2026-09-05)** closed the sibling gap
that round recorded (CF-DOC-501): the anchors read and the drawing
sweep (`drawing_edit`, the dr-1 row / column rewrite) share one prefix
resolution — `drawings.resolveDrawingPrefixes`: the root element's
prefix, every alternate bound to a spreadsheetDrawing URI anywhere in
the part, and the DEFAULT namespace as the empty prefix (`<wsDr
xmlns="…/spreadsheetDrawing"><oneCellAnchor><from><row>`, openpyxl
3.1's drawings, which the read listed as nothing and the
`xdr:`-literal sweep left in place while the grid and the series
formulas moved) — so a row insert above an openpyxl chart moves its
`from` with the grid, and this export hands over the record `zlsx
anchors` prints. A spreadsheetDrawing binding under a name the walk
cannot spell — longer than the resolver's 100-byte limit, or past its
eight-alternate replay cap; an anchor under it would be neither listed
nor shifted — is this read's `MalformedDrawingXml` (strict) and the
row / column edit's `MalformedDrawingXml` (§10, the dr-1 name; not
folded by the Editor), dry-run by `Workbook.preflightDrawingEditForSheet`
before the edit's first mutation. Pinned on every surface by the
corpus test on `tests/corpus/openpyxl_chart.xlsx`. Not
walked, by design: chartex parts (`<cx:f>`) and the `c15:` data-label
range extension. Repeated
reads stay bounded: the drawing walkers used to resolve every part
name along the relationship chain into the store's lifetime arena —
the S3B-MEM-603 shape, unbounded growth for a long-lived editor
repeating the read per call — and now resolve into the walk's own
allocator as scratch freed before return (the resolved names were only
ever lookup keys; the anchors carry the store's part names).

**Tests** (`src/c_abi.zig`, "S3b anchors_ndjson …"): the buffer equal
to the shared writer's frozen stream (the CLI test's literal) over the
pkg fixture (all three anchor kinds, a sheet boundary, chart-first
document order regrouped); the rename AND a row insert visible with no
save — the rename pinning `sheet`, the insert pinning the edited
sheet's `from` / `to` moved beside the other sheet's unmoved anchors,
and the chart's refs pinned RESPELLED under both (the chart sweep:
`Facts!$B$1` after the rename, `Facts!$B$2` / `Facts!$A$3:$A$5` /
`Facts!$B$3:$B$5` after the insert on the sheet they name);
`(NULL, 0)` on a workbook without drawings with the poison reset;
`NullOutPointer` / `InvalidInput` as call errors; poisoned outputs
reset on the refusal and undersized-diag paths, the rejected diag
byte-for-byte untouched; `-2` with the name in diag + errbuf and
nothing handed out for a dangling blip edge, a broken SECOND sheet
behind a servable first record, a bad entity in a sheet-name carrier,
a series ref that does not decode, an orphan worksheet part
(`DrawingOnUnlistedSheet`), and a flipped payload byte in a listed
sheet part (`MalformedWorkbookXml`, the graph probe) versus a drawing
part (`MalformedDrawingXml`, the walk fold); `statusOf` covers the new
name; an allocation-failure sweep over `anchorsNdjsonOwned`. In
`pkg/anchor_ndjson.zig` the collector tests moved to the new
signature and the lenient-era `relById` pin became a strict-inventory
pin (a 200-byte rid, an entity-spelled rid, an unbound rid). The
smoke gate takes the address and `#error`s without the macro; the
Python leg builds the fixture through the archive (the writer has no
drawing surface), pins the parsed records, the rename + insert-row
moves, four refusal shapes and the closed-editor error.

## 15. S3b slice 9 — the `sheet-props` and `calc-props` reads (2026-09-02)

Two exports, the slice-2 pattern verbatim, under ONE macro and ONE
probe (a dylib has either both or neither):

| Export | Probe (`_ffi.py`) | Header macro |
|---|---|---|
| `zlsx_editor_sheet_props_ndjson` (+ `zlsx_buffer_release`) | `_HAS_SHEET_PROPS` | `ZLSX_HAS_SHEET_PROPS` |
| `zlsx_editor_calc_props_ndjson` (+ `zlsx_buffer_release`) | `_HAS_SHEET_PROPS` | `ZLSX_HAS_SHEET_PROPS` |

**The bytes are the CLI's**: the S3b gate froze both records in
`docs/cli.md`, and the records are text. The exports hand over the
NDJSON bytes of `pkg/sheet_props_ndjson.zig` — the shared collectors +
record writers `zlsx sheet-props` / `zlsx calc-props` print through —
built by `c_abi.zig::sheetPropsNdjsonOwned(alloc, wb)` (`collect` +
`writeAll`: one `{"kind":"sheet_props",…}` line per workbook sheet,
workbook order, no selector — the CLI's default stream) and
`c_abi.zig::calcPropsNdjsonOwned(alloc, wb)` (`collectCalc` +
`writeCalcRecord`: the ONE `{"kind":"calc_props",…}` line) over the
editor's current parts. A sheet record is the `<dimension ref>` as
authored (null when the element or the attribute is absent) and the
FIRST `<sheetView>`'s `<pane>` as authored (null when there is none;
`x_split` / `y_split` / `top_left_cell` / `active_pane` / `state`, each
null when the source omits it, no schema default applied, split panes
reported as written — the lenient `Worksheet.freezePane` narrows to
frozen panes, the read does not); later views' panes stay in the part
(slice 8's recorded owner-shape deferral). The calc record is
`calc_id` / `full_calc_on_load` / `iterate` / `iterate_count` /
`iterate_delta` as authored, every field null when the element or the
attribute is absent — a workbook without `<calcPr>` is a record of
nulls, never `(NULL, 0)` (the doc-props convention; the export asserts
a non-empty buffer). The sheet-props stream is never empty on
success either: the strict inventory refuses a sheetless workbook —
a missing `<sheets>` (REL-602) and, since this slice's review, an
empty `<sheets/>` (CT_Sheets minOccurs=1; Codex #219 r1 S3B-REL-101:
the lenient opener accepts the shape as a zero-length inventory, so
every read over the strict inventory — conditional formats, anchors,
sheet-props — served an empty success for it, contradicting this
contract, and the calc read served its record over a sheetless
workbook; the walk now counts its entries and all four refuse). The export's
`(NULL, 0)` arm is kept for the buffer contract's uniformity, not as a
reachable shape. Release with
`zlsx_buffer_release`. The allocating writer's `WriteFailed` crosses as
`-3`, the pivots builder's rule. Python parses one JSON object per line
(`Editor.sheet_props()` / `zlsx.sheet_props(path)` → `list[dict]`;
`Editor.calc_props()` / `zlsx.calc_props(path)` → the one `dict`, a
count other than one raising `ZlsxError`), through one shared
`_ndjson_read(symbol)` helper the pair's probe gates.

**The refusals — no new name**: `collect`'s error surface is the
closed `CollectError` = `{MalformedWorkbookXml, MalformedSheetXml,
MissingSheetPart, ZipBombSuspected, OutOfMemory}`, `collectCalc`'s the
closed `CalcError` = `{MalformedWorkbookXml, ZipBombSuspected,
OutOfMemory}`; both exports switch exhaustively with no `else`
(S3B-MNT-911's rule). Every member was already mapped: a sheet list the
strict workbook read cannot prove is `MalformedWorkbookXml` (§13's
shared inventory, `conditional_format_ndjson.resolveSheets`; a listed
part the archive cannot materialise folds there, S3B-ERR-602's shape);
a sheet part the strict walk cannot prove a pane / extent for is
`MalformedSheetXml` (the `docs/cli.md` contract: a second `<dimension>`
/ `<sheetViews>` / first-view `<pane>`, a duplicate attribute on that
machinery, an MCE construct at a recognized slot, a `ref` / pane
carrier that does not decode, plus the typed sheet view's own parse
verdict on worksheet roots, REL-404); a `<calcPr>` slot the read
cannot report faithfully is `MalformedWorkbookXml` (two at the slot,
one an MCE branch could project there, a duplicate attribute, a
carrier that does not decode — and the same `<sheets>` verdict, since
the calc read runs the same strict workbook walk without resolving
sheet parts); `MissingSheetPart` is §10's (slice 8's S3B-MNT-101
ruling kept it in the set for parity with the conditional-formats
read); `ZipBombSuspected` stays the deliberate open-time `-1` §13
records; `OutOfMemory` is `-3`. Both refuse whole — a broken SECOND
sheet hands out nothing for the servable first.

**Timing**: the sheet inventory is a fresh strict read of the current
`xl/workbook.xml`, and the extent and views live in the sheet parts
the structural edits rewrite in place before the call returns — so a
rename renames `sheet`, and a row insert on a sheet with a FROZEN pane
grows `dimension` and moves the pane's `top_left_cell` (and its split,
when the insertion is above it) with the grid, the sheet sweep's
existing `processDimensionTag` / `processPaneTagRow` work, all
visible immediately. A row / column edit on a sheet whose pane is
SPLIT refuses (`SplitPaneNotSupported`, the editor's pre-existing
contract) and leaves the stream exactly as it was: the split pane the
record reports is the one the edit refuses, so the read and the
editor agree. The `<calcPr>` slot rides a rename's workbook.xml
rewrite untouched; `zlsx_editor_mark_recalc_on_load` (and a recalc
that lands) swap `fullCalcOnLoad="1"` into the live part through the
§5.7.7 transaction, so the next `calc_props` read reports it with no
save in between. Staged cell writes never touch the extent, the views
or `<calcPr>`. Repeated reads stay bounded: `collect` resolves the
inventory into the view's own arena and walks each part on gpa scratch
reclaimed per sheet (`residentBytes()` pinned flat in the package
tests); `collectCalc` allocates only scratch.

**Tests** (`src/c_abi.zig`, "S3b sheet_props_ndjson …" / "S3b
calc_props_ndjson …"): the buffers equal to the shared writers' frozen
streams (the CLI tests' literals) over the pkg fixture (a frozen
record, a split record with a fractional split, a record of nulls; a
calc record with every field set); the rename AND a row insert below
the frozen row visible with no save (the exact post-edit stream:
`A1:B4`, `top_left_cell` `C3`, `y_split` held at 1, the other sheets
byte-identical), the split-pane refusal leaving the stream as it was,
the calc record riding the rename untouched; a fresh writer's sheets
as records of nulls (never an empty stream) and a fresh writer's
`<calcPr>` as the absent record flipping to `full_calc_on_load: true`
after `zlsx_editor_mark_recalc_on_load`; `NullOutPointer` /
`InvalidInput` as call errors on both, the NULL-diag success path;
poisoned outputs reset on the refusal and undersized-diag paths, the
rejected diag byte-for-byte untouched; `-2` with the name in diag +
errbuf and nothing handed out for two extents, a duplicate pane
attribute on a broken SECOND sheet behind a servable first, a pane
carrier that does not decode, an MCE branch at the views slot, a bad
entity in a sheet-name carrier (`MalformedWorkbookXml`, before any
walk), and a flipped payload byte in a listed sheet part
(`MalformedWorkbookXml`, the inventory probe); for the calc read, two
`<calcPr>`, an MCE-projected one, a duplicate attribute, an
undecodable carrier and a `<sheets>` list the walk cannot prove; an
allocation-failure sweep over both builders. The smoke gate takes both
addresses and `#error`s without the macro; the Python leg builds the
fixture through the archive (the writer emits no `<dimension>` / no
`<calcPr>` and has no split-pane surface), pins the parsed records
and the JSON types (`True`, not `1`), the rename + insert-row moves,
the split-pane refusal, the mark-recalc flip, the refusal shapes on
both reads and the closed-editor error.

## 16. S3b slice 10 — the sheet-visibility read (2026-09-02)

One export on the READER handle. The row's Zig surface is the reader's
`Sheet.state`, and the C reader family (`zlsx_sheet_count` /
`zlsx_sheet_name` / `zlsx_sheet_index_by_name`) is where per-sheet
scalars live; the editor's C surface has no sheet count / name
export (its per-sheet reads are typed NDJSON records), so an NDJSON
handover over `Workbook` (the slice-2/6/7/9 shape) would have invented
a second inventory for one enum per sheet. One macro,
one probe, no release function:

| Export | Probe (`_ffi.py`) | Header macro |
|---|---|---|
| `zlsx_sheet_state(book, idx) → int32_t` | `_HAS_SHEET_STATE` | `ZLSX_HAS_SHEET_STATE` |

**The value is the CLI's**: `zlsx list-sheets` prints
`SheetState.toString()` of `Book.sheets[idx].state`; the export returns
that field as a code — `ZLSX_SHEET_STATE_VISIBLE` 0 / `_HIDDEN` 1 /
`_VERY_HIDDEN` 2, `xlsx.SheetState`'s declaration order, spelled in
`include/zlsx.h` and `src/c_abi.zig` and pinned to the literals on
both sides (a Zig test, a smoke-test static assert) — and Python maps
the code back to the OOXML spelling the CLI prints (`"visible"` /
`"hidden"` / `"veryHidden"`): `Book.sheet_state(selector)` by index or
by name under `Book.sheet()`'s selector rule (the two now share
`Book._sheet_index`, so a bad selector is the same `IndexError` /
`KeyError` / `TypeError` on both), and `Sheet.state` for a selected
sheet. The reader's rule carries over unchanged: a missing or
unrecognised `state` attribute reads as visible, the schema default
(`SheetState.parse` — visibility never fails an open), and hidden
sheets stay in the inventory: count, names, lookup and the row
iterators enumerate them regardless, which is the point — a
`veryHidden` sheet is unreachable from Excel's UI, and this getter is
how a caller scanning a workbook learns it exists. Out of range is
`-1`, never a code (the `zlsx_sheet_index_by_name` convention); there
is no diag, no errbuf and no allocation — the book owns the field, and
both openers (`zlsx_book_open`, `zlsx_book_open_buffer`) model it from
the same parse.

**Not a `zlsx_status_v1` export, deliberately**: nothing here can
refuse. The lenient reader decided the field at open (an archive it
cannot open never yields a handle), and a scalar getter on a live
handle has no failure the caller did not cause. The §10 vocabulary is
untouched; no refusal name, no `-2`.

**Tests** (`src/c_abi.zig`, "S3b sheet_state …"): the three codes are
the header's literals; a writer fixture with `state` attributes
spliced through the archive (the writer authors none — `Ledger`
hidden, `Secret` veryHidden, `Odd` an unrecognised value, `Data` no
attribute) reads 0 / 1 / 2 / 0 through the path opener and the buffer
opener, `-1` at the first index past the end and at `UINT32_MAX`, the
veryHidden sheet found by name, its name intact, its one row read
through `zlsx_rows_open` / `zlsx_rows_next`, and the count unchanged;
a fresh writer's sheets read visible. Since the review (in-house r1,
S3B-MNT-101): the codes test also pins `@intFromEnum(xlsx.SheetState.*)`
= 0 / 1 / 2 and `SheetState.toString()` = the three spellings — the
strings py-zlsx maps the codes back to — under the hard `zig build test`
gate, a `src/cli.zig` test pins `list-sheets`' literal `"state"`
strings on the same hidden / veryHidden / bogus fixture, and
`src/xlsx.zig` pins `SheetState.parse`'s fold at its source; CI's only
pytest lane (windows-runtime) is best-effort. `tests/c_abi_smoke.c`
`#error`s without the macro, static-asserts the three codes and takes
the address. Python (`test_basic.py`, "sheet_state"): the spellings by
index and by name, `Sheet.state`, the selector errors, the closed-book
error on both, `open_bytes` parity, the fresh-writer default, the two
defensive branches (an unknown code → `ZlsxError`, `-1` after a resolved
index → `IndexError`, an older dylib → `RuntimeError` on the method and
the property, the selector errors still first), a probe that must agree
with the library version (a ≥ 0.9 dylib without the export fails, not
skips), and — where a CLI build sits beside the dylib (locally, and CI's
windows-runtime lane as `zlsx.exe`) — `zlsx list-sheets`'s `state` equal
to `Book.sheet_state` sheet for sheet.

## 17. S3b slice 11 — formula text and error tags on the row iterator (2026-09-03)

Three exports on the ROWS handle, beside `zlsx_rows_style_at`. The row's
Zig surface is the reader's three per-row side channels
(`Rows.formulaStrings` / `Rows.formulaRefs` / `Rows.errorStrings`,
`src/xlsx.zig`), which exist so the `Cell` union never grew a formula or
an error arm: a formula cell's slot is its cached `<v>` value and an
error cell's slot is the literal as a plain string. The C ABI mirrors
that exactly — `zlsx_cell_t` and `zlsx_cell_tag_t` are untouched (an
added `ZLSX_CELL_ERROR` tag would have turned every shipped caller's
error literal into "unknown tag": py-zlsx's `_cell_to_py` folds an
unknown tag to `None`) — and hands the channels over as per-column
getters with `zlsx_rows_style_at`'s contract. One macro, one probe, no
release function, no `zlsx_status_v1`:

| Export | Probe (`_ffi.py`) | Header macro |
|---|---|---|
| `zlsx_rows_formula_at(rows, col_idx, out_ptr, out_len) → int32_t` | `_HAS_ROWS_FORMULAS` | `ZLSX_HAS_ROWS_FORMULAS` |
| `zlsx_rows_formula_ref_at(rows, col_idx, out_col, out_row) → int32_t` | (same) | (same) |
| `zlsx_rows_error_at(rows, col_idx, out_ptr, out_len) → int32_t` | (same) | (same) |

**The contract is `zlsx_rows_style_at`'s**: `0` and the out params
written when column `col_idx` of the current row is that kind of cell;
`1` when it is not (the out params untouched); `-1` when `col_idx` is out
of range for the current row. "The current row" is the one the last
`zlsx_rows_next` returned `1` for — before the first call, after a `0`
(end of sheet) or a `-1` (parse error), and after a non-zero
`zlsx_rows_skip` whether it returned `0` or `-1`, there is none and
every per-column getter on the handle answers `-1` (`zlsx_rows_style_at`
and `zlsx_rows_parse_date` included); a zero-length skip is a no-op
and keeps the row. The pointers have the cells' lifetime (until the
next `zlsx_rows_next` / a `zlsx_rows_skip` of n >= 1 or a close);
`out_col` is 0-based (A = 0) and `out_row` 1-based, the
`zlsx_merge_range_t` convention.

**The values are the CLI's**: `zlsx cells` prints `t:"formula"` with
`formula` (the `<f>` body, entity-decoded — a stand-alone formula, a
shared-formula base or an array-formula base) or `formula_ref` (a shared
/ array slave's base cell) plus `cached` (the cell value), and
`t:"error"` with `v` (the `t="e"` literal), from the same three fields
the getters read. Exactly one of `zlsx_rows_formula_at` and
`zlsx_rows_formula_ref_at` returns `0` for a formula cell and
`zlsx_rows_error_at` returns `1` for it — a formula whose cached value is
an error literal is a formula (the reader clears the error slot when
either formula slot is set: `consumeCell`, the CLI's precedence rule
`formula > error > date > primitive`); for an error cell only
`zlsx_rows_error_at` returns `0`; for a value cell — a gap included —
all three return `1`. A slave whose base the reader never saw (an `si`
with no base above it) reads as a value cell, the reader's rule. Python
maps each getter to one list aligned to the row `next()` yielded:
`Rows.formula_strings() -> list[str | None]`, `Rows.formula_refs() ->
list[CellRef | None]`, `Rows.error_strings() -> list[str | None]`, each
checking the closed-iterator `ZlsxError` before the dylib probe's
`RuntimeError`; `Rows.__next__` now zeroes its current length on `0` /
`-1` so the accessors (`style_indices` included) return `[]` past the
end, as the C side answers `-1`.

**Two siblings fixed on the way**: `zlsx_rows_style_at` and
`zlsx_rows_parse_date` bounded `col_idx` on the reader's parallel lists,
which a fast-path `zlsx_rows_skip` never clears — after a skip they
served the last decoded row's style / date for a row the caller never
saw (py-zlsx shielded `style_indices` by zeroing its length on a
successful `skip`; `parse_date` passed the index straight through). All
five getters now bound on the C-side view, which `zlsx_rows_next`
empties on `0` / `-1` and `zlsx_rows_skip` empties BEFORE it reads (a
skip that fails on the spreads path has already reset the reader's
lists and left them partially refilled from the torn row — in-house r1
S3B-REL-101/102 — so a view kept across the `-1` bounded the getters on
the old row over the torn one); `Rows.skip` zeroes its length before it
raises — on its pre-0.8.0 drain fallback too (in-house r2 S3B-REL-202)
— and `Rows.parse_date` bounds on it before the index narrows. A
zero-length skip is the no-op the contract reads as on both surfaces:
no `zlsx_rows_next` is stood in for, so the current row stays current —
the four header sentences that state the rule (`zlsx_rows_next`,
`zlsx_rows_skip`, the getters' block, `zlsx_rows_parse_date`) and the
binding's docstrings each carry the `n >= 1` qualifier or name the
`n == 0` no-op (in-house r3 S3B-DOC-301).
The only observable change to the two siblings is `-1` / `None` where
they used to answer for a row that was not current.

**Not a `zlsx_status_v1` export, deliberately**: nothing here can refuse
— the reader decided the fields while yielding the row, and a getter on
a live iterator has no failure the caller did not cause. The §10
vocabulary is untouched; no refusal name, no `-2`. Not on the bulk
matrix path either: `zlsx_matrix_data` / `Sheet.read_all` stay
values-only, as `Book.materialiseSheet` is.

**Tests** (`src/c_abi.zig`, "S3b rows formulas …" ×5): a writer fixture
with rows 2–6 spliced into the sheet part through the archive (the
writer authors neither shared formulas nor `t="e"` cells) — a
stand-alone formula, an entity-bearing body (`"x"&amp;"y"&lt;&gt;A1` →
`"x"&"y"<>A1`), a formula-only cell (`ZLSX_CELL_EMPTY` + text), a shared
base and its slave (`(0, 3)`), a `t="e"` literal, a formula whose cached
value is `#DIV/0!` (formula wins, the cell a string), an array base and
the slave inside its rectangle (`(0, 4)`), a gap before an error cell,
an empty `<f></f>` (own text of length 0 behind a written pointer), a
slave whose base was never seen (a value cell), a `t="dataTable"` body
— probed with sentinels in every out param, the pointers compared by
identity, so a `1` / `-1` that wrote anything is caught on every cell
kind (in-house r1 S3B-MNT-105/106); no current row before the first
`next`, past each row's end, at `maxInt(usize)`, past the end of the
sheet, after a skip, after a skip that FAILS on a torn row and after a
`zlsx_rows_next` that fails on it directly — the `-1` half of the rule
on both surfaces (in-house r3 S3B-MNT-303) — (one `expectNoRowAt` per
site: the trio, the style getter and the date
getter agreeing, the reader's lists shown to hold the torn row's
remains); the date getter through a fast-path skip over a dated row
(the reader's `row_cells` still two wide — row 1 — while nothing is
current) and a zero-length skip keeping the row current; the buffer
opener reading the same fields; a fresh writer's cells answering `1`
everywhere. `tests/c_abi_smoke.c`
`#error`s without the macro and takes the three addresses. Python
(`test_basic.py`, "rows_formula" / "parse_date_answers"): the six rows'
four lists, the empty lists before the first row / past the end / after
`skip` / after a failed `skip` (`parse_date` `None` at each),
`read_all` unchanged, the closed-iterator error on each accessor,
`open_bytes` parity, a fresh writer's row, the older-dylib
`RuntimeError` after the closed check, a probe that must agree with the
library version, `parse_date` through a fast-path skip and a
zero-length skip, the pre-0.8.0 drain fallback (`_HAS_ROWS_SKIP` forced
off) leaving no current row, and — where a CLI build sits beside the
dylib — `zlsx cells --include-blanks`'s `t` /
`formula` / `formula_ref` / `v` / `cached` equal to the accessors and
the row values cell for cell (ten formula cells, three error cells).
