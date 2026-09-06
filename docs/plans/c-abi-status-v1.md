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

### 4.8 `zlsx_prune_report_v1` — 40 bytes (S3c slice 3, §20)

| off | type | field |
|---|---|---|
| 0 | `size_t` | `struct_size` |
| 8 | `uint64_t` | `redacted` (slots turned into tombstones this call) |
| 16 | `uint64_t` | `stale` (content drifted, still embeddable — left as-is) |
| 24 | `uint64_t` | `fresh` (hash matches the row's content) |
| 32 | `uint64_t` | `valid_empty` (a tombstone whose row is still empty) |

Caller-owned, nothing to release; `Workbook.PruneReport`'s four `usize`
counts widened to `uint64_t` so no LLP64 caller narrows a slot count. The
fields of the `{"kind":"prune",…}` record `zlsx embed --prune` prints, in
its order.

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
folded by the Editor), run by `Workbook.applySheetEditTransform`
before the edit's first mutation (the sweep's result is what it
installs after the sheet). The sweep walks the read's lexical layer,
follows the sheet's drawing edge as the read does (`drawings.findDrawingRef`
+ the typed relationship lookup — a reference the strict read cannot
follow refuses the edit as it refuses the inventory) and reads corners
through the read's own parser (`drawings.parseCornerIn`), so on the
anchors both walk the two judge a drawing the same way: a wrapper and
its children may mix two followed spellings (`<xdr:twoCellAnchor><from>`),
comment / CDATA / PI / DTD text is never an anchor under either
spelling (a live `<!DOCTYPE` refuses both — and the strict chart walk,
so the chart sweep's preflight refuses a chart part carrying one,
`MalformedChartXml`), a `<` inside an attribute value — not well-formed XML,
the one decoy surface no exact-QName rule covers — refuses both (and
the strict chart walk), XSD-collapsed whitespace around a scalar parses
for both (digits only otherwise — `1_0` is not 10), and what the sweep
cannot move — a wrapper with no close, a corner block absent or with a
scalar that does not parse, two blocks that overlap (a `<to>` nested in
`<from>`) — is the same `MalformedDrawingXml` the strict read raises. The walks differ only
where their jobs do: the sweep reads the corners of a shape the read
never lists, the read validates a one-cell anchor's `<ext>` the sweep
does not move, a reversed `<to>` / `<from>` pair is listed and moved
in document order, a self-closing wrapper is stepped over by both. Pinned on every surface
by the corpus test on `tests/corpus/openpyxl_chart.xlsx`. One
consequence the listing makes reachable: `deleteSheet` leaves the
doomed sheet's drawing and charts in the archive (the recorded
orphan-part follow-up), and this read then refuses that workbook
whole — `DrawingOnUnlistedSheet` — openpyxl workbooks included now
that their drawings are listed (round 6, ND-REL-602). Not
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

## 18. S3c slice 1 — the embedding write (2026-09-05)

`Workbook.setEmbeddings` crosses the boundary on the editor handle — the
first S3c export, under the §2 contract. One export, one header macro,
one probe, one new struct (an array element):

| Export | Probe (`_ffi.py`) | Header macro |
|---|---|---|
| `zlsx_editor_set_embeddings(ed, model, model_len, dim, dtype, dtype_len, coverages, coverages_len, flags, diag, errbuf, errbuf_len) → int32_t` | `_HAS_EMBEDDING_WRITE` (+ `zlsx_diag_release`) | `ZLSX_HAS_EMBEDDING_WRITE` |

**The shape is the read side's** (decision S3c-1). Vectors cross as the
f32 `[rows][dim]` matrix `zlsx_emb_vectors` hands back and hashes as the
u64 per-row list `zlsx_emb_hashes` hands back, `zlsx_emb_tombstone()`
marking a row with no vector (Python: `None`); `dtype` is the string
`zlsx_emb_dtype` returns (`"f32"`, `"int8-sym-per-vec"`); the sheet is
an index, `zlsx_editor_*`'s convention, resolved to the coverage's
`worksheet_target` by `Worksheet.embeddingTarget` — the CLI's private
helper, moved so `embed --vectors` and the export spell it once. So
read → re-embed → write is ONE shape on every surface, and the on-disk
encoding lives in the library: `embedding_part.encodeVectorRecord` (the
f32 wire bytes, or the symmetric int8 quantizer's per-row
`f32 scale; i8[dim]`) is the one encoder, and `zlsx embed --vectors`
now writes through it too (its inline loop is gone). The Zig surface's
raw `vec_body` bytes stay Zig-only: a C or Python caller never
assembles an on-disk record. The three dtypes the read knows but the
writer cannot encode (`binary16`, `bfloat16`, `int8-asym-per-vec`) are
`UnsupportedDtype`, a new `embedding_part.Error` member, refused before
any byte moves.

**`zlsx_emb_coverage_v1`** (decision S3c-2): an array element with a
frozen 88-byte layout and no `struct_size` — the `zlsx_formula_cell_v1`
precedent (§3's prefix rule governs a caller-allocated struct the
library reads whole; an element inside the caller's array has no
per-element size to honour). Offsets pinned by a comptime block in
`src/c_abi.zig`, `_Static_assert`s in `tests/c_abi_smoke.c` and asserts
in `_ffi.py`:

| off | type | field |
|---|---|---|
| 0 / 8 | `const uint8_t *` / `size_t` | `id`, `id_len` (1–63 of `[A-Za-z0-9_-]`) |
| 16 / 24 | `const uint8_t *` / `size_t` | `range`, `range_len` (A1) |
| 32 / 40 | `const uint8_t *` / `size_t` | `column`, `column_len` (inside the range) |
| 48 / 56 | `const float *` / `size_t` | `vectors`, `vectors_len` (`== rows * dim`, row-major) |
| 64 / 72 | `const uint64_t *` / `size_t` | `hashes`, `hashes_len` (`== rows`) |
| 80 | `uint32_t` | `sheet_idx` (0-based) |
| 84 | `uint32_t` | `include_formulas` (0 / 1) |

Both lengths are explicit and checked against the range's row count
(`InvalidEmbeddingInput`) — the read side's `out_len` rule mirrored, so
a mis-sized caller buffer is a named error, never an out-of-bounds
read. NULL with length 0 is the empty string / array, judged by the
write (an empty model lands; an empty coverage set is
`InvalidEmbeddingInput`; an empty dtype is `InvalidDtype`); NULL with a
non-zero length is `InvalidInput`, the `bytesArg` rule.

**`flags`** (decision S3c-3) is reserved: v1 defines no bit and refuses
a set one (`InvalidInput`), so `recovery_in_cells` can cross later
without a `_v2`. It does NOT cross in this slice:
`Workbook.setEmbeddingsOpts(.recovery_in_cells)` adds its hidden sheet
through `Workbook.addSheet`, beneath the `Editor`'s sheet mirror — the
mirror would not know the sheet (the S3a all-or-nothing
`Editor.addSheet` exists precisely because the two must agree), so the
flag waits for an Editor-level path.

**The refusal split** (decision S3c-4) is §10's rule. `-1`, statements
about the call, every one raised before the first part write:
`InvalidInput`, `InvalidEmbeddingInput`, `InvalidDtype`,
`UnsupportedDtype`, `SheetIndexOutOfRange`, `InvalidCoverageId`,
`InvalidRange` (the range, or a column outside it),
`DuplicateCoverageId`, `CoverageOverlap`, `InvalidXmlByte`,
`StructSizeTooSmall`. `-2`, statements about the workbook, the name in
the diag with `plane = NONE`: `MissingWorkbookRels` /
`MalformedWorkbookRels` (`xl/_rels/workbook.xml.rels` without the
`</Relationships>` the workbook→index relationship lands before, or
`_rels/.rels` without it when `docProps/custom.xml` has to be created
for the recovery record — both pre-flighted in pass 1) /
`MissingRelationship` / `IdSpaceExhausted` (the rels file's `rId`
space, or an existing `docProps/custom.xml`'s `pid` space, already at
`UINT32_MAX` — pre-flighted in pass 1, round 2) / `MissingContentTypes`
/ `MalformedContentTypes` / `MalformedWorkbookXml` (already in
`structural_refusals`; since round 5 also the strip's verdict on an
`xl/workbook.xml` the open admits but the scanner refuses — a
`<definedName` outside `<definedNames>` — judged in pass 2c) and
`EmbeddingExceedsArchiveLimit` — new in the
list — a part past the
512 MiB read cap (`embedding_part.PART_MAX_BYTES`, the S1 per-part
limit spelled once; sized from the inputs in the C layer before a
vector byte is read, and again in the Zig write's pass 1) OR the
recovery record past the defined-name carrier's ceiling of 16 × 200
bytes (roughly eighty coverages at typical ids, or a ~3 KB model name;
encoded in pass 2c, installed last): a limit either way, so a refusal
(§2, the `FormulaLimitExceeded` rule), the one name kept rather than a
second spelling of "too big for the carrier". The Python leg mirrors
S3a: `ZlsxError` / `ZlsxRefusal(error_name)`, with the shapes
(`TypeError` / `ValueError`) checked before ctypes could truncate,
through the same `_structural_call` helper; a NumPy matrix or hash
array crosses as one contiguous, aligned float32 / uint64 buffer
(`np.require(…, ("C", "A"))`, the array itself when it already is one),
never a Python float per value.

**Three Zig-layer fixes the export exposed** (the
`Workbook.setEmbeddings` contract, every surface at once):

- The write never validated `column`; the index read does
  (`parseCoverage`: a column name inside the range) — so `column = "Z"`
  on `A2:A4` saved cleanly and `Workbook.embeddings()` refused the
  index it had just written (`InvalidRange`). One rule now,
  `embedding_part.validateCoverageColumn`, called by both.
- The S0 audit's one unchecked text channel (surface matrix §6):
  `appendXmlEscaped` in `pkg/embedding_part.zig` escaped `& < > "` only,
  so a C0 control byte in `model` reached `index.xml`.
  `embedding_part.validateMetadataText` — the writer's own
  `sheet_plan.isForbiddenXmlByte`, one definition; `pkg/embedding_part.zig`
  imports `zlsx_sheet_plan` and its test root now declares it — refuses
  `model` and `worksheet_target` in pass 1, and the encoder refuses
  again (`InvalidXmlByte`, a new `embedding_part.Error` member). Tab /
  LF / CR / DEL are XML characters and pass, as on every other channel.
- Refusals fired AFTER the parts were installed —
  `MissingWorkbookRels` / `MalformedWorkbookRels` in pass 5 (the
  relationship's splice target) and from the docProps carrier's
  `_rels/.rels` (round 1), `EmbeddingExceedsArchiveLimit` per coverage
  in pass 3 (a second oversized coverage left the first one's parts)
  and from the recovery record's chunk ceiling in the install (round
  1: a set with a new index and no record at all) — leaving a torn
  staged set. `preflightEmbeddingsRels` (both rels files), exact part
  sizing (header + body, what the encoders emit) and the record's
  encoding (`encodeRecoveryRecord`, pass 2c; `installRecoveryRecord`
  last) now run before the first write. Residue, documented on every
  surface: an allocation failure, an index XML past the cap, or a
  `[Content_Types].xml` or `docProps/custom.xml` the carriers cannot
  patch can still leave the set partially replaced — discard the editor
  without saving (the `xl/workbook.xml` strip left the list in round 5:
  prepared in pass 2c, committed after the relationship). The three docProps `allocPrint` folds no
  longer turn an allocation failure into a `-1 WriteFailed`; the
  remaining `WriteFailed` folds are `bufPrint`s the id and rId bounds
  cannot overflow.
- **Round 1's HIGH (S3C-REL-101)**, pre-existing on the Zig and CLI
  surfaces and shipped here under "replaces any previous set": a
  re-embed across a save resurrected the previous recovery generation.
  `stripRecoveryNamesFromWorkbookXml` replaced `xl/workbook.xml` in the
  store but never refreshed the parsed view, and the save-time splice
  (`applyWorkbookXmlPlanDefinedNames`) rebuilds `<definedNames>` from
  the VIEW plus the plan — so the saved chunk came back beside the new
  one (two `_zlsxRecovery0`, Excel's repair prompt) and a stripped read
  reported the OLD model. The strip now parses the fresh view over its
  own copy before the store write and swaps it, the add-sheet pattern;
  both replacement tests count the name carrier (S3C-TEST-107).

**Round 2** (two agents, every round-1 fix verified, no HIGH): the one
register-time failure the pre-flight did not mirror — `nextMaxNumericAttr(…)
+ 1` unchecked in `registerWorkbookEmbeddingsRel` (`rId`) and
`upsertRecoveryDocProp` (`pid`), a hostile `rId4294967295` /
`pid="4294967295"` trapping in Debug or wrapping in ReleaseFast after
every part — is checked arithmetic now (`IdSpaceExhausted`, `addSheet`'s
spelling, already `-2`) and pre-flighted in pass 1 on both files. The
strip judged "ours" by a substring over the attribute region and
case-sensitively: a user's `_zlsxRecoveryMine` was deleted (final since
the round-1 view refresh) and a foreign tool's `_ZLSXRECOVERY0` survived
beside ours; `recovery_record.isChunkName` — the exact
`_zlsxRecovery<digits>` form, case-insensitive — now governs the plan
drop and the strip, pinned with three user names (one hidden, one
prefixed) surviving a re-embed across a save once, flags intact, and the
case variant stripped. `Workbook.setEmbeddings` invalidates the cached
`embeddings()` view (Zig-only: the same workbook read the previous set
back). The docProps upsert writes through an `ArrayListUnmanaged` (the
`Allocating` writer spelled an allocation failure `WriteFailed`, a `-1`
after the parts). The C layer judges the call's lengths (`-1`) before the
cap (`-2`). Python: `np.require(…, ("C", "A"))` — a misaligned
`frombuffer(offset=1)` view is copied, an aligned array still crosses as
itself; a masked hash is the tombstone and a masked vector value 0; the
float32 cast runs under `np.errstate(over="ignore")` (NumPy's overflow
`RuntimeWarning` was an exception under warnings-as-errors).
`Workbook.definedNames()`'s lifetime is stated (until the next view
swap). `Dtype.recordBytes` multiplies in `usize` unchecked — 64-bit on
every shipped target.

**Round 3** (two agents, every round-2 fix verified): the strip judged
`name` only in its `name="…"` spelling while the typed parser (and so
the fresh view the save-time splice reads) accepts either quote and XML
whitespace around `=` — a single-quoted or spaced chunk from a foreign
serializer survived the strip and the round-1 shape returned for that
spelling. The strip now reads the attribute through
`workbook_xml.getAttr`, the parser's one acceptance (the S3B-REL-1201
rule); a third generation over saved bytes pins the strip itself (the
user's three names kept once, a `name = '_zlsxRecovery0'` chunk
stripped, the new record alone). Python: `_hash_buffer`'s masked branch
skipped the plain path's shape and dtype rules (a masked `(n, 1)` array
landed as tombstones) — judged first now; the masked-vector pin masks a
non-zero row; the alignment pin asserts the mechanism (an aligned,
C-contiguous buffer; an aligned array crossing as itself), since the
shipped targets load a misaligned float without a fault; the stale
`np.ascontiguousarray` wording and the `embeddings()` getter's lifetime
are corrected, and the unreachable "staged name collides with a chunk"
residue clause is gone.

**Round 4** (two agents, every round-3 fix verified, both ship-ready;
the edges closed in-round): the strip's element, attribute-region and
close scans were still lexical (a bare `indexOf` on `<definedName`, the
first `>`, a bare `indexOf` on the close) while the fresh view uses the
quote-aware `findTagEnd` and the comment-aware `findClosingTag` — a chunk
with a quoted `>` in an attribute before `name`, or a decoy inside a
comment, was seen differently by the strip and the view. The strip now
walks with the parser's own `findTagOpen` / `findClosingTag` (a chunk
with no close is `MalformedWorkbookXml`) and
judges the DECODED name (`store_mod.decodeXmlEntities`, the splice's own
reading — a character reference such as `&#95;zlsxRecovery0` was the one
way the chunk form could hide; the early return considers `&#` too). The
recovery-record READER, lexical by design over bytes a foreign tool may
have rewritten, gained the same comment skip: a decoy chunk inside a
comment used to be read as the record and the stripped read answered
`absent`. Pinned in the third generation: the saved chunk re-spelled
`comment="a>b" name = '&#95;zlsxRecovery0'` plus a single-quoted decoy
inside a comment before the block — one live `_zlsxRecovery0` after, the
decoy comment intact, the user's three names once, the stripped read the
last model; `pkg/recovery_record.zig` pins the reader on the decoy, on a
decoy-only part and on an unterminated comment. Python: the float32 cast
runs under `np.errstate(all="ignore")` — round 3's "no cast in this
direction raises `invalid`" was false (a signalling NaN does, and a
caller's `np.seterr(under="raise")` fires on a subnormal); both pinned.
The residue sentence is now on every surface (the docstring names the
index past the cap, the README carries it), the PR body's counts agree
with the suite. Carried: the record reader accepts `name='…'` / `name="…"`
but not `name =` spacing and does not decode entities in the name — a
foreign strip that also re-spells attributes would read as `absent`
(pre-existing, the reader's own scope).

**Round 5** (two agents, every round-4 fix verified through the built
dylib; A ship-ready with four LOWs, B blocking on one MEDIUM — the same
shape): the round-4 strip walks the WHOLE part with the parser's scanner
while the view bounds `<definedName` inside `<definedNames>` only, so a
`<definedName` with an unterminated quote outside the block — a part the
open admits — tripped the scanner's own `MalformedXml` from the install,
a `-1` AFTER every part, the index and the relationship (a save then
shipped an index at the new model beside a record at the old one; the
`&#` early-return clause sent record-free workbooks down the same walk).
The strip is now two halves: `prepareRecoveryNameStrip` in pass 2c —
the stripped bytes and the view parsed over them, a pure function of a
part no later pass touches, every scanner or parser verdict mapped to
`MalformedWorkbookXml` — and `commitRecoveryNameStrip` from the install
(`replacePart` + the view swap). Pinned in Zig (the junk beside a saved
record and on a record-free part spelling `&#`: the refusal, no new part,
the index still the old model), C (`ZLSX_REFUSED` + `MalformedWorkbookXml`,
`save_to_buffer == source`) and Python (the same, through a
zipfile-patched second generation). The recovery-record READER's round-4
skip covered comments only; a decoy chunk inside a CDATA section or a PI
— which the parser and the strip rightly skip, so it survived every
re-embed — was read as the record and a stripped read answered the
PREVIOUS generation. The reader now skips with the parser's own
`skipNonElement` (the module imports nothing but std, so
`pkg/recovery_record.zig` stays a leaf; its root now runs the parser's
tests too) — comments, CDATA, PIs and DOCTYPE, one acceptance; pinned
on all four plus the unterminated forms. The `&#` clause of the early
return gained its own pin (a saved chunk spelled `&#95;zlsxRecovery0`
with the prefix nowhere else). **Recorded follow-up (round 5,
S3C-REL-504 B — pre-existing, outside this slice)**: a save after any
staged defined-name edit — `set_embeddings` stages the record's hidden
names, so every one of its saves — re-emits `<definedNames>` from the
view plus the plan (`applyWorkbookXmlPlanDefinedNames` /
`spliceDefinedNamesBlock`) keeping `name` / `localSheetId` / `hidden`
only: `comment`, `description`, `help`, `statusBar`, `customMenu`,
`shortcutKey`, `function`, `vbProcedure`, `workbookParameter`,
`publishToServer`, `xlm`, `functionGroupId` are dropped from every
existing name (an `insert_row` save keeps them — the byte-preserving
`rewriteAllDefinedNames` shape). Fix = re-emit unchanged entries
byte-for-byte and regenerate only the staged ones; until then the
sentence is on every surface.

**Round 6** (two agents, both ship-ready — converged): every round-5 fix
verified in the code, in a pin that fails with its fix reverted, and
through the built library (the c15 junk on every surface, every
non-element decoy before, after and inside a chunk body, the
character-reference-only chunk, three generations with the round-2–5
spellings, the CLI's `--strip` / `--vectors` on the junk). One LOW: the
Zig setter doc and the Python README said every refusal lands before the
first write while listing the index XML's cap (pass 4) as a residue — the
cap cannot fire (the record ceiling checked in pass 2c bounds the same
fields to a few KB); both now say so.

**Recorded follow-up (round 2, S3C-REL-201 B — pre-existing, outside
this slice, an owner decision)**: the recalc transactions rebuild their
candidate from the archive AS OPENED (`PartStore.nextGeneration`:
"overrides are NOT inherited") and `recalc_txn.prepare` re-applies only
the Editor's cell parts — so `mark_recalc_on_load` + `save` after
`set_embeddings` persists a STRIPPED file (the record spliced, the parts
gone), `save_with_recalc` writes the parts but drops the hidden-name
carrier, `rename_sheet` + `mark_recalc_on_load` silently reverts the
rename, and `insert_row` / `add_sheet` + `mark_recalc_on_load` + `save`
ABORTS the process on the `sheetCount` assert. The fix belongs to the
transactions (refuse while the live generation has installs — the
`requireCompleteStructuralState` precedent — or build the candidate
over the live generation). This slice states the ordering rule on the
three surfaces ("call the recalc transactions before the write, or
save and re-open") and pins both safe orders in Python.

**Recorded, not lifted**: the index read hands `model` / `id` /
`worksheet_target` attributes back raw (no entity decoding), so a model
name carrying `&`, `<` or `"` reads back as `&amp;` … on every surface —
the defined-names ST_Xstring shape (#190), a read-side slice of its own;
a tab / LF / CR in `model` passes the byte rule (an XML 1.0 character,
as on every other channel) into an ATTRIBUTE value, which a conforming
parser normalizes to a space while zlsx's lexical read returns it raw
— documented on every surface, plain spaces recommended; the hidden
`_zlsxRecoveryN` names are staged with the workbook plan and appear in
`defined_names()` only after a save (the S3b read walks the parsed
view); "the one encoder" is the one PRODUCTION encoder — the emb-4 /
emb-4b fixture generators keep an independent spelling of the int8-sym
record as the layout oracle; the `recovery_in_cells` flag (above);
embeddable rows, prune, strip → C + Py (the rest of S3c).

**Tests** (`src/c_abi.zig`, "S3c set_embeddings …" ×7): the f32 round
trip on two sheets with a tombstone and `include_formulas`, read back
through `zlsx_emb_*` and the Zig read (the column, the flag, ONE
relationship, both carriers, no hidden sheet); int8-sym within one step
of the per-row scale and the compact part size; a second write replacing
the set in one editor and across a save (one relationship, one property,
ONE `_zlsxRecovery0` carrying the last model); thirty `-1` cases with the
name in errbuf, the diag as prep left it, the NULL-with-length-0 rule,
`StructSizeTooSmall` byte-for-byte, and `save_to_buffer` equal to the
source (nothing written); the `-2` `MalformedWorkbookRels` with a
poisoned diag reset and nothing written, plus `statusOf` pinned on both
vocabularies; the refusals that used to fire after the parts —
`_rels/.rels` without its close tag, a rels file at `rId4294967295`,
a 3300-byte model, a one-row 2^27-wide f32 coverage and a full-grid
2^20-wide one refused from the inputs before a vector byte is read —
each with `save_to_buffer` equal to the source; the body encoder under
`checkAllAllocationFailures`.
`pkg/embedding_part.zig`: the encoders read back through
`decodeAllF32`, shape checks before any byte moves, the column rule, the
byte rule over every value 0x00–0x7F and through `encodeIndexXml`.
`pkg/workbook.zig`: the column and control-byte refusals leaving no
part, both rels pre-flights leaving no part, an exhausted `rId` and
`pid` space leaving no part, the record ceiling (120 one-row
coverages, a 3300-byte model) leaving no part while forty coverages
land, and the re-embed across a save over three generations (one name
in the view and the file, three user names kept once with their flags
through the plan drop AND the strip, a foreign case-variant and a
spaced single-quoted chunk stripped, the same workbook's `embeddings()`
answering the new model, a stripped read reporting the LAST
generation), the strip's scanner verdict beside a saved record and on a
record-free part spelling `&#` (`MalformedWorkbookXml`, no new part, the
index still the old model), and a saved chunk spelled only as a
character reference stripped on the re-embed; `pkg/recovery_record.zig`
pins `isChunkName` and the reader's non-element skip (a decoy in a
comment, a CDATA section, a PI and a DOCTYPE subset, plus the
unterminated forms). Each fix of rounds 1–5 was mutation-checked:
reverted, its Zig, C or Python tests fail.
`tests/c_abi_smoke.c` `#error`s without the macro, takes the address
and pins the struct. Python (`test_embedding_write.py`, 45): the goal
line — `embeddings()` `present` after the write with the provenance,
the cells untouched, both carriers; vectors / hashes / `valid_mask`
read back and re-embedded on the read side's own arrays; int8-sym;
replacement across a save with one hidden name and a stripped read
reporting the last model; fourteen named `ZlsxError`s writing nothing;
the `MalformedWorkbookRels` `ZlsxRefusal` on either rels file, the
`MalformedWorkbookXml` refusal on a second generation whose part the
strip cannot walk, and the record-ceiling refusal, nothing written;
NumPy crossing without
`tolist` (an ndarray subclass whose `tolist` raises), signed / float /
object / bool arrays judged, `2**64 - 1` as the tombstone; a
misaligned view (the mechanism: an aligned buffer, an aligned array
crossing as itself), masked hashes and a masked non-zero vector row, a
masked array's shape and dtype rules, a float64 overflow, a signalling
NaN under warnings-as-errors and a subnormal under a caller's
`under="raise"`; the two documented recalc orders landing; seventeen
Python-side shape errors; NumPy width against `dim`; the closed editor
and the older-dylib `RuntimeError`; the probe against the version.

## 19. S3c slice 2 — the embeddable-rows read (2026-09-06)

`Workbook.embeddableRows` crosses the boundary on the editor handle — the
second S3c export, under the §2 contract. One export, one header macro,
one probe, no new struct:

| Export | Probe (`_ffi.py`) | Header macro |
|---|---|---|
| `zlsx_editor_embeddable_rows_ndjson(ed, sheet_idx, range, range_len, column, column_len, include_formulas, out, out_len, diag, errbuf, errbuf_len) → int32_t` | `_HAS_EMBEDDABLE_ROWS` (+ `zlsx_buffer_release`, `zlsx_diag_release`) | `ZLSX_HAS_EMBEDDABLE_ROWS` |

**The shape is the CLI's** (decision S3c-5): the S3b buffer contract —
the `{"kind":"embed_row","row":N,"text":"…","hash":H}` records `zlsx
embed --extract` prints, one line per embeddable row in range order,
written once for every surface by `pkg/embeddable_row_ndjson.zig` (the
`defined_name_ndjson.zig` precedent; the CLI writes through it too, its
inline loop is gone). The record gains `hash` on every surface: the
CLI's used to carry `row` and `text` only — `embed --vectors`
recomputes the hash from the workbook itself — but a C or Python caller
feeding `zlsx_editor_set_embeddings` cannot, and the goal of the slice
is precisely the canonical hash. It is spelled as an unsigned 64-bit
decimal, the value `zlsx_emb_hashes` reads back (a JSON reader that
narrows integers to doubles — `jq` before 1.7 — rounds it; Python's
`json` does not; documented on the CLI and the header). The arguments
are the write's coverage fields — the sheet as an index (resolved by
`Worksheet.embeddingTarget`, the write's one spelling), the A1 range,
the column inside it, `include_formulas` as 0 / 1 — so the read matches
the coverage it feeds. Rows with nothing embeddable are omitted, the
Zig contract: a covered row missing from the read is a
`zlsx_emb_tombstone()` slot (Python `None`) on the write; the Python
docstring carries the mapping. A range with none is `ZLSX_OK` with
`(NULL, 0)`.

**The refusal split** (decision S3c-6) is §10's rule. `-1`, statements
about the call: `InvalidInput` (NULL handle, NULL bytes with a non-zero
length, `include_formulas` past 1), `NullOutPointer`, `InvalidRange`
(the range, or a column outside it), `SheetIndexOutOfRange`,
`StructSizeTooSmall`, and `SheetHasUnsavedMutations` /
`SheetHasUnsavedAppends` — a sequencing statement (the caller staged
writes on the sheet it is reading), the plane `applySheetEdit`'s own
refusal of the same state already has (`statusOf` pins it `-1`). `-2`,
statements about the workbook, the name in the diag with `plane =
NONE`: `MissingRelationship` / `MissingSheetPart` (the sheet's part is
unreachable) and `MalformedSheetXml` (a sheet part the view cannot
parse, or — since round 2 — a row or cell it cannot place: no `r`, or
one the view cannot read — 0, non-numeric, past the limit — or a ref
under another row), all already in `structural_refusals`, plus
five new entries —
`UnsupportedCellValue` (a boolean `<v>` that is not `0` / `1`, a `<v>`
the number canonicalizer cannot read — a comma decimal, `NaN` — a
`t="d"` ISO-8601 date, a `t` this reader does not know, a shared-string
index that is not a number, an entity the decoder does not know),
`SstIndexOutOfRange` (an index past the table), `InvalidUtf8`,
`UnicodeNormalizationFailed` — a cell value the read cannot carry,
refused whole rather than a record that lies (the `MalformedSheetXml`
shape) — and `MalformedSharedStringsXml`, a shared-string table the
view cannot parse (round 1). None of the five is raised by any other
status_v1 export (`InvalidUtf8` / `UnicodeNormalizationFailed` come
from the canonicalizer alone, the other three from this read's own
cell and part rules), so the global classifier changes no shipped
verdict. The Python
leg mirrors the S3b reads: `ZlsxError` named after the cause,
`ZlsxRefusal(error_name)`, the shapes (`TypeError` / `ValueError`)
checked before ctypes, the buffer released on every path.

**Three Zig-layer fixes the export exposed** (the
`Workbook.embeddableRows` contract, every surface at once — `zlsx embed
--extract` and `--vectors` included):

- The read handed over cell text RAW. `sheet_xml.Cell.raw_value` is
  undecoded by contract ("caller decodes if needed") and so is the SST
  view's `.plain`, so a cell spelling `a &amp; b &lt;c&gt;` went to the
  embedder as the entities and was hashed as such; a rich shared string
  refused the whole read (`SstEntryIsRich` — one bold word in the
  column and `--extract` failed); a rich inline string was its first
  run alone. `canonicalCellFor` (`canonicalCellOf` since round 1, the
  cell in hand) now reads a cell the way `readHostGrid`
  (the S7b pivot rebuild) reads one — a shared string plain or its runs
  joined, an inline string's runs joined through `inlineStringText`, a
  formula's cached `t="str"` value decoded — through
  `sst_xml.decodeText`, the ONE strict decoder every cell reader uses;
  an entity it does not know is `UnsupportedCellValue`. The hash is
  taken over that text — the design's "visible text" — and
  `pruneEmbeddings` derives its verdict through the same function, so
  a rich or entity-bearing row reads fresh after a re-embed rather
  than permanently stale. Hash compatibility: a stored hash computed
  over an entity-bearing raw `<t>` by the previous code reads `stale`
  once (the old hash was over bytes no reader sees); every other row's
  hash is unchanged.
- `embeddableRows` returned borrowed text and folded EVERY hasher
  error into `WriteFailed` — an allocation failure crossed as a `-1`
  and invalid UTF-8 as a nonsense name. It now returns an arena-owned
  `EmbeddableRows` (the `DefinedNames` shape; `deinit` frees the
  decoded texts) and propagates the hasher's own names; a text that is
  not UTF-8 is `InvalidUtf8` for every cell kind, judged before the
  hash, because the NDJSON writer passes bytes through.
- Staged cell writes were invisible to the read: `cellByRef` walks the
  parsed view, and `Worksheet.deltas` / `appended_rows` are not in it,
  so `setCell` then `embeddableRows` answered with the SAVED content and
  a hash the staged value turns stale the moment it lands.
  `classifySlot` already refused to call a staged value fresh
  ("re-deriving its canonical hash would mean reproducing the
  save-time encoding"); the read now refuses the sheet whole
  (`SheetHasUnsavedMutations` / `SheetHasUnsavedAppends`, the
  `applySheetEdit` rule) — save and re-open, or read before the
  writes. The CLI opens fresh and never meets it.

**Recorded, outside the slice**: the xlsx WRITER numbers a rich cell
as it is written (`sstInternRich`, "rich entries are indexed after
plain new") but emits every plain SST entry before the rich ones, so a
plain string written AFTER a rich row takes the rich cell's index — the
rich cell then shows the later plain string in Excel (found by this
slice's fixture: `writeRow("a & b")`, `writeRichRow(…)`,
`writeRow("placeholder")` → the rich cell read "placeholder"). A
writer fix of its own with a regression pin on all three surfaces
(`write_rich_row`); the fixtures here write the rich row last. Also
carried: prune / strip → C + Py, `recovery_in_cells`, the CLI vector
dump (the rest of S3c).

**Round 1** (two agents; A 3 MEDIUM + 2 LOW, B 1 MEDIUM + 6 LOW, both
not ship-ready on the same MEDIUM; every fix in the round-1 commit): the
`try` on the hasher surfaced the number canonicalizer's own
`MalformedNumber` — a comma decimal, `NaN`, or every `t="d"` ISO-8601
date cell (a Strict-conformance workbook's dates) refused the whole
read as an unclassified `-1` on no surface's list, and `--prune`
aborted the same way (REL-101) → `hashCanonical` folds it under
`UnsupportedCellValue` at both call sites, and `.date` refuses there
explicitly (the canonical form has no date kind). A shared-string
table the parser cannot read crossed as `-1 MalformedXml` while the
sheet part's failure is `-2 MalformedSheetXml` (REL-102) →
`sstDecodedText` maps the store's and the parser's names to a new
`MalformedSharedStringsXml` (`Workbook.Error`, `structural_refusals`,
every surface). One `cellByRef` scan per covered row made the read and
the sweep quadratic in the sheet — 34 s for a 40 000-row column through
the dylib (PERF-103) → `columnCellsOverRange` resolves the column's
cells from ONE pass over the view for both (`canonicalCellOf` takes
the cell; a cell in the wrong `<row>` or with a ref this reader cannot
parse is `MalformedSheetXml`, the host-grid rule). A cell whose `t`
names no known type was served as a number with a number-kind hash (B
REL-103) → `UnsupportedCellValue`. The pre-hash UTF-8 check was
load-bearing for number / error cells only and pinned on strings only
(TEST-104) → an error literal and a number with a `0xFF` byte on all
three surfaces, and the `StructSizeTooSmall` case now poisons the
outputs. Docs: `rename_table_column` stages the header cell as a delta
on the host sheet, so it is a trigger of `SheetHasUnsavedMutations`
too (B DOC-105 — named on every surface); the error-cell kind (`#N/A`,
hash kind `e`) was missing from every new surface's description of
`text`, and `EmbeddableRow.text` overstated the arena's ownership (a
number's or an error's `<v>` borrows the parsed view) (A DOC-105); the
Python count (DOC-106). Recorded: the read is stricter than
`readHostGrid` where a value's MEANING is unknown (a non-numeric
shared-string index refuses here, the grid carries "occupied, no
text"); the staged-write refusal is per sheet, by design; the same
editor answers again after its save (pinned in Python).

**Round 2** (two agents; A ship-ready with 3 LOW, B 1 MEDIUM + 3 LOW —
every round-1 fix verified in code and by a pin that fails on a
revert; fixes in the round-2 commit): a `<row>` or `<c>` written
without `r` — legal positional OOXML the typed parser drops and counts
(`unaddressed_rows` / `unaddressed_cells`) — was invisible to the
resolver, so `--extract` / the export omitted a row the lenient reader
shows and `--prune` ZEROED its vector (B REL-201; the pivot edit
already refuses the shape, Codex #205 r1 REL-103) → the resolver
refuses `MalformedSheetXml` on either count, read and sweep alike,
pinned on all three surfaces and through `pruneEmbeddings`. The
"wrong `<row>`" verdict was applied inside the range only, while the
comment and §19 claimed the whole-sheet host-grid rule (A REL-201) →
every cell of the sheet is judged before the range filter, pinned with
a misplaced cell outside the range on all three surfaces. The `.date`
pin (`2024-01-01`) exercised the round-1 fold, not the arm (A
TEST-201) → `<v>2024</v>`. `EmbeddableRow.text` for a number or an
error borrowed the parsed view, which the next save or structural edit
drops (B REL-202) → duped into the rows' arena; the doc is "owned by
the arena" again. docs/cli.md named none of the refusals the embed
family exits 3 on (B DOC-203) → listed on `--extract` and `--prune`.
The round-1 fold was pinned through the read only (B TEST-204) → the
sweep's own pins (a comma decimal, a date, an unknown `t`, a row and a
cell without `r`). §19's "four new entries" and the causes list (A
DOC-201) → five and seven. **Round 3** (two agents, convergence; A
ship-ready with 3 LOW): positional OOXML is in the project's own corpus
— `wdi_excel.xlsx` (the World Bank exporter) carries no `r` on any row
of its six sheets, and on most cells (the first sheet: 401 395 rows,
9 994 345 of 10 492 903 cells `r`-less — a mixed shape, the rows being
the trigger), so every embed surface now refuses it `MalformedSheetXml` where
`main` answered an EMPTY row list silently (nothing served for a sheet
full of text — the same silent omission the round-2 fix closed);
recorded here, pinned corpus-gated in `tests/workbook_corpus.zig`, and
the follow-up named: the typed parser inferring positions the way
`src/xlsx.zig` and the formula decoder already do would lift this
refusal and `PivotEditUnsafe` on the same shape — a slice of its own.
B's corpus sweep (31 workbooks, 98 sheet pairs, HEAD against a
`main` build): 85 pairs byte-identical modulo `hash`; besides WDI,
`poi_poc_shared_strings.xlsx` (one `<row r="0">`, which the view
counts as unplaced) now refuses where `main` served 49 rows — the
clause on every surface names the `r` shapes the view cannot read;
eight `poi_clusterfuzz_xssf` sheets `main` refused `SstEntryIsRich`
are served, and `calamine_encoded_entities` embeds the decoded text.
The round-2 dupe of a number's text into the arena had no pin →
pointer identity against the view's `<v>` (TEST-302); docs/cli.md
omitted `MissingRelationship` / `MissingSheetPart` and §19's decision
paragraph the round-2 clause (DOC-303). Recorded: `worksheetForTarget` resolves
every sibling sheet's part name in index order, so a sibling without
`r:id` refuses a healthy sheet's read (pre-existing, the sweep's
too); a `t="d"` cell in a covered column stops `--prune` whole (the
pre-round-1 code aborted the same way as `WriteFailed`); the lenient
reader admits `A03` where the typed layer refuses; `--extract`'s bytes
rest on the shared writer, unpinned in `cli.zig` (the S3b precedent).

**Tests** (`src/c_abi.zig`, "S3c embeddable_rows …" ×5): the records
on the S3c fixture byte-for-byte against the shared writer over the
library's own hasher, the row-2 literal pinned, the hashes fed back to
`set_embeddings` and `pruneEmbeddings` calling every slot fresh, the
second sheet by index, an empty range resetting the outputs; every cell
kind (an entity-bearing string, a number, a blank omitted, a boolean, a
formula's cached string on request, a rich string's runs joined); nine
`-1` cases with the name in errbuf, the diag as prep left it and the
outputs reset, the NULL-with-length rules, `NullOutPointer` on either
output, `StructSizeTooSmall` byte-for-byte with the outputs poisoned
and reset, the staged write on one sheet leaving the other readable;
thirteen `-2` patches (`</sheetData>` gone, an index past the table,
`<v>TRUE</v>`, `&bogus;`, a `0xFF` byte in a string, `<v>1,5</v>`, a
`t="d"` date whose `<v>` is a number, `t="zz"`, a `0xFF` byte in an
error literal and in a number, the table's last `</si>` gone, a row
and a cell without `r`) with the name in the diag, plane NONE, the
other sheet still served after a sheet part's verdict and refused
after the table's, plus a misplaced cell outside the range; the plane
of every verdict, the canonicalizer's own name pinned `-1` should it
ever escape.
`pkg/workbook.zig` ("embeddableRows …" ×7): the three existing pins on
the arena shape; the decoded text and its hash on an entity-bearing
shared string, a rich shared string and an entity-bearing inline
string, fed back and pruned fresh; numbers, booleans and a cached
formula string with and without `include_formulas`; the staged-write
and append refusals (a write outside the range refuses too); thirteen
cell / part verdicts on patched parts (the round-1 seven among them, a
row and a cell without `r`, a cell whose ref names another row —
inside the range, and outside it) and the inline run walker's, the
sweep refusing on the same five shapes (a comma decimal, a date, an
unknown `t`, a row and a cell without `r`), and an error cell carried
as its literal with the kind-`e` hash.
`pkg/embeddable_row_ndjson.zig` ×2 (the wire shape, no rows no bytes).
`tests/c_abi_smoke.c` `#error`s without the macro and takes the
address. Python (`test_embeddable_rows.py`, 34): the goal line — the
read's hashes written by `set_embeddings` read back with `valid_mask`
all true and equal; the CLI shape with the pinned literal, the second
sheet, the empty range; every kind; omitted rows mapped to tombstones;
six named `ZlsxError`s; the staged write and append with the other
sheet, the same editor after its save and the saved file answering;
thirteen `ZlsxRefusal`s through zipfile-patched copies (the table's
refusing the other sheet too) plus a misplaced cell outside the range;
an error cell carried as its literal; six shape errors; the closed
editor and the older-dylib `RuntimeError`; the probe against the
version.

## 20. S3c slice 3 — the embedding sweeps (2026-09-06)

`Workbook.pruneEmbeddings` and `Workbook.stripEmbeddings` cross the
boundary on the editor handle — the third and last S3c export pair on the
mutation side, under the §2 contract. Two exports, one header macro, one
probe, one new struct (§4.8):

| Export | Probe (`_ffi.py`) | Header macro |
|---|---|---|
| `zlsx_editor_prune_embeddings(ed, report, diag, errbuf, errbuf_len) → int32_t` | `_HAS_EMBEDDING_SWEEPS` (+ `zlsx_diag_release`) | `ZLSX_HAS_EMBEDDING_SWEEPS` |
| `zlsx_editor_strip_embeddings(ed, diag, errbuf, errbuf_len) → int32_t` | `_HAS_EMBEDDING_SWEEPS` (+ `zlsx_diag_release`) | `ZLSX_HAS_EMBEDDING_SWEEPS` |

**The shape is the CLI's** (decision S3c-7): what `zlsx embed --prune` and
`zlsx embed --strip` run, on the S3a status_v1 shape — no arguments beyond
the handle, staged in memory, `zlsx_editor_save` commits. Prune's result is
the point of running it, so it crosses as a caller-owned
`zlsx_prune_report_v1` (§4.8; a struct rather than four out-pointers, so a
fifth count can append under the §3 rule): the four counts of the CLI's
`{"kind":"prune",…}` record, in its order, as `uint64_t`; NULL is
`NullOutPointer` (the S3b reads' name for a missing output), and the report
is prepped beside the diag before anything is judged — a rejected
`struct_size` on either leaves the accepted one zeroed and the rejected one
untouched byte for byte. Python returns the report as a dict with the same
four keys (`Editor.prune_embeddings()`), and `Editor.strip_embeddings()`
returns nothing; both reach the Zig layer through the one `_structural_call`
helper the S3a exports use. A workbook with no set, or a stripped one (the
recovery record alone), prunes to every count 0 with ZLSX_OK; a strip on a
workbook without embeddings is ZLSX_OK and stages nothing (the save is the
untouched editor's passthrough, pinned byte for byte on every surface); a
sweep with nothing to redact rewrites nothing.

**The refusal split** (decision S3c-8) is §10's rule. `-1`, statements
about the call: `InvalidInput` (NULL handle), `NullOutPointer`,
`StructSizeTooSmall`, `SheetHasUnsavedAppends` (prune: a covered sheet
with staged `append_row` rows — the read's rule, see below) and
`SheetDeleteRequiresCleanState` (strip: staged cell writes or appended
rows on any sheet while a `recovery_in_cells` sheet has to go — the
delete's sequencing statement, `zlsx_editor_delete_sheet`'s own `-1`).
`-2`, statements about the workbook, the name in the diag with `plane =
NONE`, each before the first part write or removal: two new entries in
`structural_refusals` — `MalformedEmbeddingSet` and `MissingEmbeddingPart`
— plus, on prune, the embeddable-rows read's own verdicts on a covered cell
(`MissingRelationship` / `MissingSheetPart`, `MalformedSheetXml`,
`MalformedSharedStringsXml`, `UnsupportedCellValue`, `SstIndexOutOfRange`,
`InvalidUtf8`, `UnicodeNormalizationFailed`: a cell the read cannot carry
stops the sweep whole rather than redact a row that has text — §19's rule,
the same `columnCellsOverRange` / `classifySlot` reading), and, on strip,
`MalformedWorkbookXml` (an `xl/workbook.xml` the chunk-name strip cannot
walk — the slice-1 REL-501 shape), `CannotDeleteLastSheet` and the sheet
delete's verdicts should the cells sheet be present (`MissingRelationship`, a
carrier the cross-sheet sweeps cannot read — `MalformedChartXml`,
`MalformedExtensionXml`, …). `-3`: the sweeps' allocating writers and
`allocPrint` spell an allocation failure `WriteFailed`, folded to
`OutOfMemory` at the boundary (`failSweep`, the §19
`embeddableRowsNdjsonOwned` rule; nothing else on either path raises the
name). Neither new `-2` name is raised by any other status_v1 export
(`setEmbeddings` never reads the index; the E5 read handle `zlsx_emb_open`
is a pointer-returning open and keeps the parser's own names in errbuf), so
the global classifier changes no shipped verdict.

**`MalformedEmbeddingSet`, and why the sweep folds** (the §19
`UnsupportedCellValue` / `MalformedSharedStringsXml` precedent): the index
read (`Workbook.embeddings`) raises the embedding parser's own names — a
coverage `@range` it cannot parse is `InvalidRange`, a `dtype` it does not
know `InvalidDtype`, a `vec.bin` with the wrong magic `BadMagic`, a count
that disagrees with its header `CountMismatch` / `CoverageCountMismatch` —
and `zlsx_editor_set_embeddings` spells several of those same names for a
statement about ITS inputs, pinned `-1` since slice 1. From a sweep they
are statements about the workbook, so `pruneEmbeddings` folds every
`embedding_part.Error` name the index read raises under
`MalformedEmbeddingSet` (`embeddingSetVerdict`, `isEmbeddingPartError` over
the parser's error set — one rule, no list to drift) and keeps the read's
`MissingEmbeddingPart` (the index's rels, or a vec / hash part it names,
gone) and the archive's own verdicts; the two defensive bounds inside the
walk (`hashes.value`, `redactSlot`) spell the same name. The read side is
unchanged: `zlsx_emb_open` on the same file reports `InvalidRange`, pinned
beside the sweep's verdict on every surface.

**The mirror-safe strip** (decision S3c-9): the `recovery_in_cells`
carrier is a hidden sheet, and `Workbook.stripEmbeddings` deleted it
through `Workbook.deleteSheet` — beneath the Editor's `sheet_paths`
mirror, which then held one entry more than the workbook: an index the
mirror read as one sheet and the workbook as another, a `set_cell` after
the strip landing on the wrong sheet, and the next `add_sheet` compounding
it (the CLI's `--strip` opened fresh, saved at once and never met it).
`Editor.stripEmbeddings` (new) deletes that sheet through the Editor's own
`deleteSheet` first — its pre-flights (`SheetDeleteRequiresCleanState`,
`CannotDeleteLastSheet`, allocated-before-mutation, the S3a REL-402 rule)
ahead of the first removal, the doomed part removed the way the workbook's
own path removes it — and the workbook sweep then finds nothing left to
delete; the strip's own `xl/workbook.xml` verdict is judged ahead of the
delete as a pure check (`Workbook.preflightEmbeddingStrip`). The export and
the CLI both run through it: one path.

**Two Zig-layer holes the export exposed** (the `Workbook.stripEmbeddings`
contract, every surface at once — `zlsx embed --strip` included):

- A strip of a SAVED `recovery_in_cells` set read back `.stripped` with the
  whole record. The record reader scans every worksheet part and
  `xl/sharedStrings.xml` for the record text (a Numbers export lifts the
  cell's string into the table), and a save does the same to the recovery
  cell — so deleting the sheet left the string in the table and the
  reader found it. The in-memory pin ("strip: removes the recovery sheet
  the cell opt-in creates") never saved, so the cell was still staged and
  vanished with the sheet; the C test on a saved fixture found it.
  `scrubRecoveryCellText` now blanks the record in place wherever the
  reader would find it (the table and every worksheet part — an inline
  string, a renumbered sheet, an orphaned part), through the reader's own
  locator (`recovery_record.recordSpanInText`, factored out of
  `findRecordInText` so what one finds the other blanks), so no
  shared-string index moves; and `removeRecoveryCellSheet` removes the
  doomed sheet's part, which `deleteSheet` leaves as an orphan by design.
  Pinned on the Zig layer (the table's bytes before and after: one span
  gone, every other byte kept), the Editor, the C surface and Python
  (`absent` after a strip is the goal line).
- Every verdict the strip can give fired AFTER the first removal: the
  chunk-name strip walked `xl/workbook.xml` after the parts and the
  relationship were gone, and the sheet delete ran last — a refusal left
  the vectors gone and the record standing, or the parts gone and the
  sheet standing. The strip now prepares the name strip first (the
  slice-1 pass-1 shape: `prepareRecoveryNameStrip` before the first
  removal, `commitRecoveryNameStrip` in place of the old one-shot) and,
  when a cells sheet is present, judges the pure verdict, deletes the
  sheet (whose own verdicts then precede every removal) and prepares the
  name strip over the post-delete bytes — the order matters: a strip
  prepared before the delete would commit the deleted `<sheet>` back.
  Pinned with the store's mutation counter unchanged after each refusal,
  with and without the cells sheet, on every surface.

**One prune rule** the export needed: a covered sheet with staged appended
rows now refuses (`SheetHasUnsavedAppends`) before the first part write.
The parsed view the sweep walks does not carry `Worksheet.appended_rows`,
so a slot for a row they will become classified as gone and was redacted
— the §19 shape, on the sweep. Staged `setCell` deltas are judged as they
were (`classifySlot`: a blank redacts, any other value is `stale`, never
`fresh`), the documented `setCell(.blank)` → `prune` → `save` sequence.
Appended rows on an uncovered sheet do not touch the sweep.

**Round 1** (two agents; A ship-ready with 1 MEDIUM + 4 LOW, B not
ship-ready on 1 MEDIUM + 4 LOW; every fix in the round-1 commit, each
pinned on the surface that found it): `pruneEmbeddings` re-parsed the
index rels into a fixed 64-entry buffer and spelled every parse failure —
`BufferTooSmall` at the 65th relationship included — `MissingEmbeddingPart`,
so a set of 33+ coverages that `embeddings()` accepts through its grow loop
refused the sweep on every surface under a name the docs define as "a part
gone from the archive" (A REL-101, pre-existing on the Zig / CLI path) →
`EmbeddingView` carries `rels` (the retained grow-loop storage, borrowed
like the coverages) and the sweep resolves part names from the view;
pinned with 33 single-row coverages. The scrub replaced worksheet parts
and the table without dropping `Worksheet.parsed` / `sst_view`, and a save
with staged deltas regenerates `<sheetData>` from the view — prune (parses
the covered sheet) → strip → `setCell` → save wrote an inline record back
and the file read `.stripped` (B REL-101) → `dropViewOverPart`, the
every-other-replace-site shape; pinned in the table's shape and the
inline one. The scrub blanked every locator hit while the reader honours
only a span it can decode under its 3 200-byte cap — a cell that merely
spelled `zlsxER1` was cut by a strip advertised as a no-op (A + B REL-102)
→ `recordSpanDecodes`, the reader's own gates (`recovery_record.decode`
under the same grow loop) before a span is blanked; pinned on the Zig
layer and Python (`see zlsxER1 for the magic`, `xzlsxER1abc`, the
passthrough byte for byte). `PartStore.hasUnsavedChanges` saw overrides
only and `removePart` installs none, so a strip whose only store mutation
was removals — a foreign package with the parts under `<Default>` content
types, no relationship, no record — saved the source bytes with every
vector intact and reported OK (A REL-103 / B REL-105, pre-existing on the
CLI path; no zlsx-written file reaches it) → the store counts committed
removals (`removed`) and the predicate ORs it in; pinned store-level (a
removal with no override, no relationship) and end-to-end in Python (the
foreign shape strips to `absent`, the archive without the parts). The
scrub decompressed every worksheet part into the store arena on every
strip, the documented no-op included (B PERF-103) → `StripScope`: the
table is always scrubbed (where zlsx's writer and a Numbers export put the
string), the worksheet parts only for a workbook that carried the
`zlsxRecovery` sheet at entry (the Editor passing what it saw before its
own delete), `_rels` files skipped — recorded: a user-RENAMED recovery
sheet holding an INLINE record is not scanned (its table string still is).
B's notes taken: the sweep's final loop installed the redacted parts one
`replacePart` at a time, so an allocation failure landed N−1 of N — now
`PartStore.replaceParts`, the atomic install, and the "-3 after the first
part write" clause is gone from every surface; `embeddingSetVerdict`
names the error names the parser shares with the workbook vocabulary
(`MissingRelationship`, `InvalidUtf8`, `UnicodeNormalizationFailed`, `BufferTooSmall`, `InvalidXmlByte`) so a future workbook-level raise inside `embeddings()`
is not folded by accident; a NULL report beside a rejected diag is two
-1s and errbuf carries the later (`NullOutPointer`). Docs (A + B DOC-104,
A DOC-105): the Python README still called the sweeps "CLI / Zig-only
today" above the paragraph that ships them; "a blank or deleted cell" (no
C delete export) → "a blank (empty) cell"; "call again — the strip is
re-runnable" did not hold when the cells-sheet delete tore past its
pre-flights — `StructuralEditIncomplete` on the retry and on the save —
so the strip's -1 vocabulary names it on every surface and the clause
says discard; `MalformedContentTypes` dropped from the strip's -2 list
(unreachable: `removeContentTypeOverride` allocates only).

**Round 2** (two agents on the round-1 tree; both ship-ready, A 2 LOW,
B 4 LOW — every round-1 fix verified in code and by the test that fails
on its revert; fixes in the round-2 commit): the `StripScope` gate took
its evidence from the live sheet list alone, so a strip torn by a `-3`
between the cells-sheet delete and the scrub (or a hidden sheet a user
had already deleted) left an orphaned worksheet part the reader still
scans out of the retry's scope — "call again" under-delivered for an
inline-string orphan (A + B REL-201) → `hasOrphanWorksheetPart`: a
worksheet part no `<sheet>` resolves to is evidence too, OR-ed into the
scope on both paths; pinned with an orphan holding an inline record and
no `zlsxRecovery` sheet stripping to `.absent`. The Editor removed the
doomed part under the mirror's spelling (`Book.sheets[i].path`, no
dot-segment collapse) where the workbook's path uses `resolvePartName`
(B MAINT-201) → the workbook's spelling on both, and the Editor pin
asserts the part present before the strip so the removal pin is not
vacuous (A's note). Docs (A + B DOC-201/202): §20 still listed
`MissingContentTypes` / `MalformedContentTypes` as strip refusals —
neither can fire on the strip path (`removePart` resolves the part it
then `part()`s) — dropped from every surface; the Python count read 20
where the file holds 22; the lenient-reader claim named a pin that read
a stripped file whose table carried no record — the Zig typed-views pin
now reads the SCRUBBED table through the lenient reader. Recorded (B):
the reader honours only the first magic per part and stops at the first
part with one, so a user span sorting before a real record hides it from
the read while the scrub blanks every decodable span (the safe
direction); a foreign producer writing two `zlsxRecovery` sheets sees
the second deleted beneath the mirror (illegal OOXML); a re-embed across
a save leaves the previous generation's `<si>` unreferenced in the table
and a `.stripped` read reports the OLDEST record (slice-1 territory);
`recoverEmbeddingRecord`'s own `BufferTooSmall` (a record past 1 MiB of
scratch — unreachable under the 3 200-byte carrier bound) would fold
under `MalformedEmbeddingSet`.

**Round 3** (two agents on the round-2 tree, the convergence round; both
ship-ready, A 3 LOW, B 2 LOW — every earlier fix verified in code and by
its reverting test; fixes in the round-3 commit): the round-2 orphan
check resolved every sheet's part with `try`, so a `<sheet>` whose
relationship the workbook lacks — a shape `Workbook.open` admits and
`main` stripped clean — refused the strip `MissingRelationship`, a name
the four surfaces reserve for the cells-sheet delete (A + B REL-301) → an
unresolvable sheet is reason to scan wider, never a verdict (only an
allocation failure keeps its name); pinned with a dangling `r:id` on
the Zig surface (the Editor / C / Python / CLI refuse that workbook at
open, `SheetCountMismatch`). The round-2 Editor fix — the doomed part
removed under `resolvePartName`'s spelling — had no fixture on which
the mirror's and the workbook's spellings differ, so a revert to
`sheet_paths[idx]` passed the suite (A MAINT-302 / B TEST-301) → tried:
a relationship Target with a dot segment (`./worksheets/sheet2.xml`) is
refused at `Editor.open` (`MissingSheet` — no ZIP entry under the
verbatim spelling), so the shape on which the two spellings differ is
unreachable through the Editor; the one-spelling rule stays as
consistency with the workbook path, recorded on the code and here, and
the pin asserts the mirror's and the workbook's spellings agree on the
fixture. Docs
(A DOC-303): the c_abi comment still named "a content-types part a
removal cannot patch" as a post-removal -2 (dropped from the header in
round 2) — gone; the fold comment and this section listed the names the
parser shares with the workbook vocabulary from memory — now the exact
set (`MissingRelationship`, `InvalidUtf8`, `UnicodeNormalizationFailed`, `BufferTooSmall`, `InvalidXmlByte`).

**Round 4** (two agents on the round-3 tree, the confirmation round): A
converged — every vector "ANALYSED — no defect", every earlier fix
verified in code and by its reverting test; B ship-ready with one doc
LOW (DOC-401: this census listed eight of the ten `pkg/workbook.zig`
slice-3 tests — the round-2 orphan pin and the round-3 dangling-`r:id`
pin were missing; fixed here) and confirmed the round-3 `catch` swallows
only `MissingRelationship`, the sole non-OOM name `resolvePartName` can
raise, with nothing downstream re-resolving a sheet.

**Recorded, outside the slice**: the recalc transactions
(`zlsx_editor_mark_recalc_on_load` + save, `zlsx_editor_save_with_recalc`,
`zlsx_editor_recalculate`) rebuild their candidate from the archive as
opened and carry neither sweep, as they carry no `set_embeddings` (slice
1's recorded follow-up — the ordering rule is on every surface's
documentation of the write; the sweeps inherit it). `recovery_in_cells` on
the write side is still Zig-only (`Workbook.setEmbeddingsOpts` adds its
hidden sheet beneath the Editor's mirror; the strip side of that path is
what this slice built). The CLI vector / state dump.

**Tests** (`src/c_abi.zig`, "S3c prune_embeddings …" ×4, "S3c
strip_embeddings …" ×3, "S3c sweeps …" ×1; round-1 pins on the Zig layer
and Python — `pkg/workbook.zig` "S3c slice 3 r1 …" ×4: 33 coverages
through the view's relationships, the typed views dropped by the scrub in
both shapes with a staged edit and a save reading `.absent`, a cell that
merely spells the magic left with the store's mutation counter unchanged
under both scopes, a removal with no override and no relationship flipping
`hasUnsavedChanges`; `test_embedding_sweeps.py` +2, 22: the magic-only
cell's passthrough and the lenient read of it, the foreign package whose
only change is removals stripping to `absent` with the archive lacking
the parts; the lenient reader over a stripped file's cells on the strip
test — and, in round 2, over the SCRUBBED table in the Zig
typed-views pin): the read's hashes pruning all
fresh with the passthrough save byte for byte, a staged blank redacting and
a staged string stale, the saved tombstone over a zeroed vector through
`zlsx_emb_hashes` / `zlsx_emb_vectors` and `valid_empty` on the re-open; a
row blanked on disk, an edited row stale with nothing rewritten, a workbook
without a set all zeros; NULL handle, NULL report, a rejected `struct_size`
on either output (the report untouched byte for byte, the diag prepped;
the diag rejected, the report zeroed), appended rows on the uncovered sheet
ignored and on the covered sheet `-1`; six `-2` patches (the index's
`range` and `dtype`, the vec part's magic, the rels target, a cell without
`r`, a boolean `<v>`) with the name in the diag, plane NONE, the
passthrough save after each and the E5 read handle keeping the parser's
own name; the strip to ABSENT with every carrier checked in the archive,
twice, the cells still served and the no-set strip a passthrough; the
`recovery_in_cells` sheet through the mirror (`add_sheet` at 2, a write
landing there, ABSENT after the save, the dirty editor `-1` with the set
and the sheet still standing); the `xl/workbook.xml` verdict with and
without the cells sheet, the passthrough after each; the plane of every
verdict, the parser's names `-1`, `WriteFailed` `-3`.
`pkg/workbook.zig` ("S3c slice 3 …" ×4, "S3c slice 3 r1 …" ×4, "S3c slice 3
r2 …" ×1, "S3c slice 3 r3 …" ×1 — ten): the fold on five patches with the
read keeping the parser's name and the store's mutation counter unchanged;
appended rows covered and uncovered; the SAVED cells strip to `.absent`
with the table blanked in place (one span gone, every other byte kept),
the orphan gone, the cells served and the second strip changing nothing;
the workbook.xml verdict before the first removal with and without the
cells sheet; the round-1 four (33 coverages, the typed views dropped in
both shapes with the lenient read over the scrubbed table, the magic-only
cell left under both scopes, a removal flipping `hasUnsavedChanges`); the
round-2 orphan holding an inline record with no `zlsxRecovery` sheet
stripping to `.absent`; the round-3 dangling `r:id` widening the scan
and stripping clean. `pkg/editor.zig` ×1: the mirror in step after the strip (a
sheet added at 1, a write landing on it, read back from the saved file),
the orphan gone, the dirty editor refusing with nothing removed.
`tests/c_abi_smoke.c` `#error`s without the macro, pins the 40-byte
layout and takes both addresses. Python (`test_embedding_sweeps.py`, 22 — 14 tests, one an 8-case parametrize):
the goal line — the read's hashes prune all fresh and a stripped file
reads `absent`; the report's keys and a workbook without a set; a row
blanked on disk with the tombstone and the zeroed vector read back; an
edited row stale with the passthrough; the staged `set_cell` shapes; the
append refusal on the covered sheet only; a set stripped by a tool pruning
to zeros then stripping to `absent`; eight `ZlsxRefusal`s through
zipfile-patched copies with the passthrough after each and the read
keeping the parser's name; the strip's carriers in the archive, idempotent,
a no-op without a set; the `MalformedWorkbookXml` pre-flight; the closed
editor, the older-dylib `RuntimeError`, the probe against the version.
