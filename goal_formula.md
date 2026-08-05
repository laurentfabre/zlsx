# 🎯 Goal — the zlsx formula interpreter (tier D1)

> Living plan for building, testing, and improving a **formula evaluation engine**
> inside zlsx, under the same constraints as the rest of the project: stdlib-only,
> Zig 0.16.0, TigerStyle-defensive, byte-splice fidelity, refuse-rather-than-half-do,
> measured performance. Companion to `goal.md` / `goal_plan.md`; detail specs in
> `docs/plans/formula-*.md`. **Self-contained: every normative contract appears in
> this document in full — no reference to any prior revision is normative.**

_Created: 2026-08-02 · Revision: **v21** (post Codex rounds 1–20 — 412 findings
dispositioned; see §15) · Status: **PARKED (2026-08-02)** — the SHIP-READY review
loop was stopped by owner directive after round 20; this document is committed as
the design record, not as scheduled work. Round 21 was never run._

---

## Table of contents

1. [What / Why — the D1 reversal](#1-what--why)
2. [Decision log (locked)](#2-decision-log)
3. [Constraints inherited](#3-constraints)
4. [What exists today](#4-what-exists-today)
5. [Architecture](#5-architecture)
6. [Milestone ladder](#6-milestone-ladder)
7. [Function registry](#7-function-registry)
8. [Testing & oracles](#8-testing--oracles)
9. [Performance & limits](#9-performance--limits)
10. [Refusal & error taxonomy](#10-refusal--error-taxonomy)
11. [Interactions with existing arcs](#11-interactions)
12. [API surface per layer](#12-api-surface)
13. [Documentation flips](#13-documentation-flips)
14. [Risks](#14-risks)
15. [Review log](#15-review-log)

---

## 1. What / Why

**What**: (a) workbook recalculation filling every formula cell's cached `<v>`,
including as part of save on the Workbook/Editor path and — via the `zlsx_recalc`
orchestrator module and `Writer.saveToOwnedBuffer` (§5.10) — the fresh-Writer
path; (b) standalone formula evaluation against a workbook context. One engine;
customer layers CLI, C ABI, Python, Spark-batch; Zig `pkg` as the source-visible
internal foundation (no core-source rights in any tier, `LICENSE:5-30`).

**Why now**: (1) `emitCell` drops cached values on rewritten/new formulas
(`pkg/workbook.zig:366,5476` — "Tier D1" named); edits leave non-uniform
staleness. (2) Non-Excel consumers (Spark source, UDF, Genie) read `<v>` —
stale caches are silent data corruption; the commercial argument. (3) #140's
refusal wants a real parser (`goal.md`: "highest-value remaining lift").

**Reversal recorded** at §13's enumerated checklist; the planning flip is
**M-1** and precedes all code.

**Non-goals** (deferred, do not drift): chart/pivot authoring; `.xls`/`.xlsb`;
VBA/add-in/UDF execution; external links; RTD/CUBE; Python-in-Excel; what-if
data tables (refusal); LAMBDA/LET (parse-and-refuse); relative-ref defined
names (refusal); localized function names; locale-sensitive *input* outside
pinned grammars (typed refusal, never a fabricated error value — §5.4b);
streaming-source recalc; namespace-prefixed SpreadsheetML (preflight refusal;
scanners match literal names today — `sheet_xml.zig:421`,
`xlsx_test_corpus.md:58`); **`fullPrecision="0"`
(precision-as-displayed) refuses through all of v1** — support is an M10+
capability after numfmt, not an M8 side effect; **top-level multi-area
reference results** (typed refusal at every public layer in v1 — §5.4c);
OS-timezone resolution (stdlib-only Zig 0.16 has no portable resolver —
**all layers default to UTC** unless an explicit `--utc-offset` is given;
TZif integration is an M10+ milestone). Compatibility Version 2 is **in
scope** (§5.4d) — CV2 is the default for new Current-Channel workbooks since
April 2026, so refusing it would refuse normal files.

---

## 2. Decision log

Locked 2026-08-02 by Laurent (axioms; changing one reopens the plan):

| # | Decision | Consequence |
|---|---|---|
| D-1 | Both recalc-on-save and standalone eval | `recalculate()` (in-memory transaction) + `saveWithRecalc` (file transaction, §5.7.9) + Writer path via orchestrator; eval at every layer |
| D-2 | Registry phased to 300+, Core ~60 first | Core gate = M4e ✅ — **59** cumulative (M4c 20 + M4d 17 + M4e 22), counted from the frozen inventory committed at M3a rather than from this row. Past the gate: M4f ✅ adds 19 (F1c-text) → **78**, M4g ✅ adds 15 (F1c-date) → **93** |
| D-3 | Dual fidelity modes | §5.4; `.excel` carries `platform_profile` + `collation_v1` |
| D-4 | Dynamic arrays from the start | Array-first eval from M3a with the normative shape/coercion table (§5.3b); persistence staged behind proof gates; authoring via explicit `FormulaWrite.dialect` |
| D-5 | Iterative calc honoring `calcPr` | CalcState workbook-derived; multi-SCC schedule normative (§5.6c) |
| D-6 | Pinned clock/seed | RunInputs = instant + UTC offset + seed; volatile draw schedule keyed per call site (§5.6d) |
| D-7 | All customer layers in v1 (Spark batch-only) | v1 = **M9d** |
| D-8 | Four ground-truth sources | All adapters at M1b; volatiles excluded from **every external** value oracle (§8.2) |

---

## 3. Constraints

- Zero deps; Zig 0.16.0 — `std.Io` explicit in every file/buffer API
  (`AGENTS.md:250-260`; precedent `Book.openBuffer(allocator, io, bytes)`,
  `src/xlsx.zig:795-811`); TigerStyle; bounds-checked narrowing (`AGENTS.md:107`).
- **Module graph is law** (`build.zig:20-23` forbids
  `writer → zlsx_pkg → workbook → zlsx → writer`): Writer-path recalc composes
  in a **third public module `zlsx_recalc`** (§5.10) exported alongside the two
  existing public modules (`build.zig:143-157,408-438`), with a
  `tests/consumer` dependency test. `src/formula/` never imports `pkg/`
  (`EvalEnv` only; CI grep gate, M0). Fuzz wiring explicit (`build.zig:746`);
  wiring is part of each PR gate.
- Byte-splice fidelity with raw-entry identity assertions (§8.4;
  `store.zig:695`).
- Refuse rather than half-do.
- Performance under existing lanes; single-mode whole-graph bench builds
  (the `build.zig:1030` mixed-mode trap fixed in the bench PR) (§9).
- Five test layers; `checkAllAllocationFailures` incl. reports (prepared
  pre-swap, §5.7.4).
- C ABI: 3-file transaction; `zlsx_status_v1` numeric contract inline
  (§12.3); error channel on every fallible export; **write only
  `min(caller_size, known_size)` — bytes beyond the known prefix are never
  touched** (canary-tail tests); release fn per owned descriptor; field-width
  boundary tests.
- Licensing boundary; lazy-analysis gate.

---

## 4. What exists today

Verified 2026-08-02 on `main` @ `07c99f0`:

| Fact | Evidence | Consequence |
|---|---|---|
| Tokenizer round-trip lossless; `Table1[A1]` mis-tokenized; **identifier predicate ASCII-only; error whitelist closed** | `tokenizer.zig:77,266-274,309-339` | M1a: correction fixtures; **Unicode identifier grammar** (sheet/defined names are not ASCII-only) + **extensible error-literal rule** (unknown `#…!`/`#…?` lexemes tokenize as `.error_lit`, byte-preserved) with astral/combining/unknown-error round-trips |
| Many A1/ref helpers, mixed bases | rewriter; `xlsx.zig:241,2564`; writer; `workbook.zig:381,391,5396`; `sheet_plan.zig`; `sheet_edit.zig:117,1955`; `table_edit.zig:354`; `cli.zig:1065,4392` | M0: rg-derived consolidation + typed `SheetIndex`/`Row1`/`Col0` |
| Shared slaves recognized only self-closing; early slaves dropped | `xlsx.zig:2099` | **CLOSED (M4b2)**: classification is by attribute (`t="shared"` + `ref` = master, without = slave), sheet-wide topology validated, slaves translated. `calamine_non_monotonic_si.xlsx` writes `<f si="1" t="shared"></f>` — the shape the reader drops |
| Formula text + identifiers XML-escaped; **`<f>` attributes discarded by the typed parser** | `sheet_xml.zig:489-535`, `workbook_xml.zig:24,58` | M4b1 decode boundary + decoded symbol layer; **CLOSED (M4b2)**: the complete `CT_CellFormula` inventory is a table `classifyFormula` reads, one fixture per row, unknown attribute refuses (§5.7.2) |
| Scanners literal-match `<sheetData>/<row>/<c>` | `sheet_xml.zig:421` | Namespace preflight refusal (M4b1) |
| Delta model can't carry formula+cache; **`Writer.save` is path-only; its archive inputs are private** | `workbook.zig:353,5410,6206`; `writer.zig:484-512,940` | `ResolvedSheet` (§5.7.3); **`Writer.saveToOwnedBuffer(allocator, io)`** producer API (M5c; allocator-first per `AGENTS.md:194`) — `pkg/fresh_emit.zig:107` alone cannot serve an external orchestrator |
| `PartStore.open` path-only, retains `io`; calcChain rels owner-relative (`Target="calcChain.xml"`), `removeRelationshipsTo` matches absolute | `store.zig:105-144,1510`; corpus `phpoi_test1.xlsx` | `Workbook.openBuffer(allocator, io, bytes)` — borrow ends at return, store copies (Book precedent); rel-target resolution relative to `xl/workbook.xml` (M5b2) |
| `calcPr` partial; `date1904` absent from typed view; CV extension unparsed | `workbook_xml.zig:89,286-318` | **CLOSED (M4b2)**: complete `CT_CalcPr`, `sheetCalcPr`, `date1904` → `t="d"`'s epoch, extensions preserved byte-exact; every corpus `<calcPr>` round-trips (§5.7.6). CV2's feature name stays unpinned until M7b's byte-diff |
| `cm`/`vm` = one-based metadata indexes; no parser; **spec-vs-Office collection resolution differs** | MS-OE376 | M4a typed reader; **transition rows name the exact collection (cellMetadata vs valueMetadata), record type, indexing, and missing-record behavior — pinned empirically by byte-diffed Excel references at M7b**, not assumed from the base schema |
| Tables can carry `calculatedColumnFormula`/totals formulas | `table_edit.zig:39` | M4b3 producer inventory + refusal when member cells lack `<f>` |
| Cached text may need C0 controls; emitters reject forbidden XML bytes | `sheet_plan.zig:1153` | **CLOSED (M4f)**: the codec landed at M4b1 (`decode.zig:596-657`); M4f is the row that first PRODUCES such text — `CHAR(1)` — and round-trips it from a formula result through `encodeAuthoredString` → `decodeCarrier(.string)`, including the `_x005F_` escape of a literal `_x0041_`. M5b1 writes it back |
| ABI errors today: `0/-1` + error buffer; `write_row_with_formulas` has no dialect param | `c_abi.zig:1837-1920,1927` | New exports keep errbuf + add diag; **FormulaWrite stays Zig-only at M7c; the versioned C export lands at M9a2** |
| CLI reserves 130/143; flushes completed records on signal | `docs/cli.md:224-227,268`; `cli.zig:1718` | Exit tables include 130/143; **no-output guarantee applies to workbook-file mutation only; eval NDJSON streams may be prefix-valid on cancellation** (§12.2) |

Hazards fixed, not inherited: shared-attr strip; typed-overlay gaps;
`formula-literal-masking.md:42-48` misstatement (M-1 correction + cache policy
table).

---

## 5. Architecture

### 5.1 Module map

```mermaid
%%{init: {'theme': 'base', 'themeVariables': {'primaryColor': '#1a1a2e', 'primaryTextColor': '#e0e0e0', 'primaryBorderColor': '#00d4ff', 'lineColor': '#00d4ff', 'secondaryColor': '#16213e', 'tertiaryColor': '#0f3460', 'fontFamily': 'monospace'}}}%%
graph TD
    TOK["tokenizer.zig (M1a: Unicode idents,<br/>extensible error lits)"] --> PARSE["parser.zig — AST + printer"]
    REFS["refs/refs.zig — module zlsx_refs<br/>(M0, typed coords + import gate)"] --> PARSE
    PARSE --> EVAL["eval.zig — engine"]
    VAL["value.zig"] --> EVAL
    ENV["env.zig — EvalEnv (ordered sparse<br/>iteration + logicalBlankCount)<br/>+ NameResolver seam (M4b3)"] --> EVAL
    NAMES["names.zig (M4b3) — CT_DefinedName +<br/>CT_TableFormula inventories, §5.9 order<br/>drivers, 3D matrix"] --> EVAL
    NAMES --> SYM["symbols.zig — decoded symbol layer<br/>(§5.9 tiers, resolver seam)"]
    SYM --> WB
    CRIT["criteria.zig"] --> FN["registry.zig — comptime table<br/>+ frozen v1 inventory"]
    SDATE["serial_date.zig"] --> FN
    XSTR["xstring.zig"] --> FN
    NUMFMT["numfmt.zig (M8)"] --> FN
    FN --> EVAL
    EVAL --> GRAPH["graph.zig — nodes, multi-SCC schedule,<br/>callsite-keyed volatile draws"]
    GRAPH --> WB["pkg/workbook.zig — adapter, ResolvedSheet,<br/>prepare/swap, openBuffer"]
    WRT["src/writer.zig — saveToOwnedBuffer"] --> ORCH
    WB --> ORCH["zlsx_recalc — third public module<br/>(imports zlsx + zlsx_pkg; no cycle)"]
    ORCH --> CLI["src/cli.zig"]
    ORCH --> CABI["C ABI"]
    CABI --> PY["py-zlsx"] --> SPARK["zlsx.spark (batch, digest-verified)"]
```

**M0 placement (decided 2026-08-03).** The coordinate module roots at
top-level `refs/`, not `src/formula/`, for the same reason as `unicode/`:
a file belongs to exactly one module's package tree, and the consumers
span `zlsx` (`src/`), `zlsx_pkg` (`pkg/`), `zlsx_sheet_plan`, and
`zlsx_workbook_xml_plan`. Under `src/` or `pkg/` the compile fails with
"file exists in modules 'zlsx' and 'zlsx_refs'".

M0 also found more duplication than the ladder row anticipated: **six**
A1 parsers and **seven** column-letter formatters, disagreeing on column
base (0- vs 1-based), case handling, and whether `A01` is a valid row —
including two structs both named `CellRef` with different bases, and one
file (`pkg/sheet_plan.zig`) that disagreed with itself. All are now
adapters over `zlsx_refs`, each preserving its own policy via explicit
options rather than being silently unified; `refs/import_gate.zig` runs
in `zig build test` and fails on any new hand-rolled base-26 codec.

### 5.2 Grammar contract

**Tokens (M1a)**: `.op_spill`, `.op_at`, `.structured_ref` (with `'[ '] '# '@`
escapes), `.external_ref` (refused typed); 3D/full-col structural. **Unicode identifiers (normative grammar)**: start = `XID_Start ∪ {_, \}`,
continue = `XID_Continue ∪ {_, .}` (**backslash start-only**; `a\b`
refuses, fixtured); max 255 code points; invalid UTF-8 →
typed refusal; **classification precedence**: cell reference → R1C1-shape
rejection → TRUE/FALSE → function/name. Fixtures: combining marks, astral
starts, cell-looking names.
**Extensible error literals**: known ten match the whitelist; any other
`#…!`/`#…?` lexeme tokenizes as `.error_lit` with bytes preserved (rich/future
errors round-trip). Round-trip holds for all. Compat gate: previously-
recognized constructs rewrite byte-identically; `Table1[A1]` correction
fixtured.

**M1a decisions (shipped 2026-08-03).** Nine points the row left open or
got wrong, pinned here because M2 builds directly on them:

1. **Refusals are out-of-band, not errors.** The tokenizer never fails
   on malformed input (only OOM) — failing would break the round-trip
   every downstream byte-identity gate rests on. `scan()` returns
   `{tokens, refusals}` with a typed `Refusal.Reason` per construct
   (`invalid_utf8`, `identifier_too_long`, `backslash_after_start`,
   `r1c1_reference`, `external_reference`, `unterminated_*`), each
   annotated with the §10 plane-2 error M2 raises for it. `tokenize()`
   keeps the pre-M1a signature for the rewriter.
2. **`?` dropped from the identifier set** (it was in the pre-M1a
   predicate). Excel does not accept `?` in defined names, and the
   normative grammar is `XID_Continue ∪ {_, .}`. Correction, fixtured.
3. **Extensible-error rule is bounded**: `#` + 1..62 bytes of
   `[A-Za-z0-9_/.]` + `!`/`?`, cap 64 bytes total. Unbounded, a stray
   `#` swallows the rest of the expression looking for a terminator. A
   `#` matching neither the whitelist nor this rule is `.op_spill`.
4. **Bare `RC` is NOT rejected as R1C1** — `RC` is also column 471, so
   `RC:RC` is a live A1 full-column reference and the A1 reading wins
   under the §5.2 precedence. Rejection covers `R<n>C<n>` (≥1 digit)
   plus `R[`/`C[` by lookahead. `RC1` never reaches the rule: it is a
   cell reference.
5. **`.external_ref` also covers the `[1]` workbook-index form**, and
   the *whole chain it qualifies* is untouchable — `[1]Sheet1!A1`
   tokenizes as an ordinary `name bang cell_ref` triple that the
   rewriter would otherwise happily shift. `isOpaqueQualifier` +
   `endOfExternalReference` in `rewriter.zig` are the guard; the compat
   suite is what proves it.
6. **`$1` needs its own scan branch.** Digits cannot start an
   identifier, so the `$` of the absolute full-row form `$1:$1` would
   otherwise strand as a lone lexeme and split a construct the pre-M1a
   tokenizer kept whole.
7. **`a\b` yields two adjacent names, not name-unknown-name.** `\` is
   start-only, so it terminates the first identifier and starts a
   second. The refusal is the signal; the token shape is incidental.
8. **The regen gate normalises through `zig fmt`** before diffing: the
   committed tables are formatted and `zig fmt` column-aligns long
   scalar lists, so a raw diff reports whitespace churn on byte-correct
   data. Running it also proved the committed casefold/NFC tables
   reproduce exactly from their pinned inputs.
9. **`refs/` was missing from `build.zig.zon` `.paths`** (M0 gap — the
   packaged module could not have built); added with
   `THIRD_PARTY_NOTICES.md`.

**Still open (owner action):** `LICENSE` needs the third-party-data
carve-out saying the Unicode-derived tables are licensed separately.
`THIRD_PARTY_NOTICES.md` flags it; the carve-out itself is not a code
change and is not made here.

**Parser (M2)**: AST + canonical printer (parse→print→parse structural
equality). Precedence (oracle-pinned rows): `:` > ` ` > `,`(ref) >
unary`±`(right; `=-1^2=1`) > `%` (interaction fixtures) > `^` (left; `2^3^2=64`)
> `*` `/` > `+` `-` > `&` > comparisons.

**Full-row references**: `1:1` / `1:5` join full-column refs as
structurally-recognized, evaluable constructs (bounds `[1, 1 048 576]`,
sparse `RangeSet` membership, graph range-node edges; fixtures `SUM(1:1)`,
`COUNTBLANK(2:3)`); structural *rewriting* of full-row/col refs remains the
rewriter's documented out-of-scope (`rewriter.zig:22`) and an M10+ item —
evaluation and rewriting are independent capabilities.

**Array constants**: `{…}`, `,` columns / `;` rows (re-disambiguated inside
braces), elements = signed numeric literals, strings, TRUE/FALSE, error
literals; rectangularity required (ragged → typed parse refusal); no
refs/calls/nesting. **Leading `=`**: stored text is body-only; standalone APIs
strip exactly one optional `=` after whitespace; `==` → parse refusal.
**Structured refs**: full item-specifier grammar (`#All #Data #Headers
#Totals #This Row`, `@`, column ranges, combined forms, escapes). **Prefixes**:
layered `_xlfn.` / `_xlfn._xlws.` / `_xlfn.SINGLE`→`@`; `_xlpm.` refused;
original spelling preserved.

**M2 decisions (shipped 2026-08-03).** Fifteen points the row left open
or got wrong, pinned here because M3a builds directly on them:

1. **Number literals are not converted.** The AST keeps the source
   spelling and defers `parseDecimal` to M3a1. Converting here would
   force a rounding policy before §5.4 pins one, and the canonical
   printer would then have to *re-render* every literal — turning a
   round-trip property into a formatting bet.
2. **The canonical printer is canonical, not byte-preserving.** Byte
   fidelity stays the tokenizer's contract (`format(tokenize(x)) == x`);
   the parser drops insignificant whitespace and normalises booleans to
   `TRUE`/`FALSE`. Everything a canonical form cannot re-derive rides in
   the AST verbatim: number spellings, string bodies, error spellings,
   sheet-name quoting and its `''` escapes, structured-ref column names,
   and the `_xlfn.` layers. Parenthesised groups keep an explicit node —
   dropping redundant parens would change the structure the round-trip
   is asserting.
3. **A call requires `(` adjacent to the name.** M1a's classification
   precedence is normative, so `SUM (A1)` keeps the kinds M1a assigned
   it and parses as an *intersection*, not a call. Excel would read it
   as a call; the divergence is unreachable through the workbook path
   because stored `<f>` never spaces a call. Fixtured, and preferred
   over shipping two classifiers that disagree.
4. **`A1 -1` is subtraction, not an intersection.** `+` and `-` are
   excluded from the primary-starter set that turns whitespace into the
   intersection operator. Without the exclusion every `A1 -1` in the
   corpus silently becomes an intersection with a negative literal.
5. **`_xlfn.SINGLE(x)` and `@x` unify into one node.** Same operator,
   two spellings: the node carries a `form` discriminator and the
   original callee text, so downstream sees one shape and the printer
   still hands back what was written. Unification requires exactly one
   argument; anything else stays an ordinary call for the registry to
   reject on arity.
6. **`;` outside braces is `FormulaLocaleSensitiveInput`, not a syntax
   error.** Inside braces it is the array row separator. The tokenizer
   emits `.arg_sep` for both bytes, so the parser re-disambiguates on
   the byte — the one place a token kind alone is not enough.
7. **Full-span references are matched on the tokens, before any node
   exists.** `A:A` and `1:1` are recognised speculatively, which keeps
   the flat node array free of orphaned operands and makes
   `max_ast_nodes` count only live nodes. Out-of-grid spans (`XFE:XFE`,
   `1:1048577`) are deliberately **not** full spans: they stay generic
   ranges over names and reach §5.9, which is where `#NAME?` is
   provable.
8. **Structured-ref item specifiers are a set.** The canonical form
   orders them `#All #Data #Headers #Totals #This Row` regardless of how
   they were written, so two spellings of one selection compare equal.
   `@` sets `this_row` and records `at_shorthand` purely as a print
   form. A bare part starting with an unescaped `#` **must** be a known
   item specifier — `Table1[#Nope]` is malformed, not a column named
   `#Nope` (that is `Table1['#Nope]`).
9. **Column names are kept raw, with their `'` escapes intact**, so
   printing needs no allocation and no re-escaping; `decodeColumnName`
   resolves them for the M4b1 symbol layer. The bracketing rule is
   stricter than the grammar needs — only `,` and `:` are genuinely
   ambiguous, but the canonical form brackets anything outside
   letters/digits/`_`/`.`/non-ASCII, because Excel writes
   `Table1[@[Col A]]` and a canonical `Table1[@Col A]` would be a
   spelling no other reader emits.
10. **Refusals are a union, not an error set.** A refusal without a span
    is not actionable and a Zig error carries no payload, so `parse`
    returns `Parsed{ok|refused}` and reserves `error.OutOfMemory` for
    the only real error. Tokenizer refusals are checked **before**
    parsing and the first in detection order wins, so the diagnostic
    names the construct rather than whatever syntax error it went on to
    cause.
11. **Three of the seven §9 parse limits are dominated at defaults.** A
    code point is at most four UTF-8 bytes and
    `max_formula_utf8_bytes = 4 × max_formula_chars`, so the byte cap
    can be *reached* but never exceeded first; every token spans at
    least one code point and `max_tokens > max_formula_chars`; nodes are
    bounded by tokens. All three stay enforced and boundary-tested at
    lowered ceilings, because `Limits` is caller-adjustable.
12. **"No lost bytes" is an invariant, not a test.** `parse` accumulates
    every consumed token's length — speculative matches rewind it with
    the cursor — and asserts the total equals the input length on
    success; `trailing_input` is what makes reaching the end mandatory.
    Since M1a proved the tokens tile the input exactly, that total is a
    proof no byte was skipped. The fuzz target adds the structural half:
    spans nest inside their parent, and siblings are ordered and
    disjoint.
13. **Oracle pinning is by fidelity.** `ieee` manifests are compared
    bit-exactly; `excel` manifests to 15 significant digits, which is
    §5.4's own display rule — an excel-fidelity manifest records
    `0.1+0.2` as `0.3`, so demanding exact bits there would test zlsx
    against a rounding policy M3a1 has not landed. The tolerance cannot
    blur a precedence question: the discriminating cases differ by a
    factor of 8 (`2^3^2` = 64 vs 512) and by sign (`-1^2` = 1 vs −1).
14. **`&` and the reference-operator ranks are spec-pinned, not
    oracle-pinned.** No committed manifest mixes `&` with `+` or with a
    comparison, and none exercises `,` or the intersection space. They
    ship as fixtures **labelled as such** rather than claimed as
    oracle-derived; closing that gap needs new oracle cases, not a new
    parser.
15. **§5.9 ships as a spec artifact plus classification, not
    resolution.** There is no workbook at M2. The parser records
    call-vs-value use, the peeled prefix layers, the original spelling,
    and the `_xlnm.` flag; both resolution orders are exported as
    checked arrays that M4b3 consumes.

**Deferred to M10+:** call classification when whitespace separates the
name from `(` (decision 3) — it needs a lexer-level rule the rewriter
shares, not a parser-local override.

### 5.3 Value model & array semantics

**5.3a Types**: `ScalarValue` = number(f64, finite) | text | boolean |
err(known|rich) | blank — the only Matrix element, **internal to the evaluator**; the public boundary uses `PublishedScalar`/`PublishedMatrix` (no blank variant) with **one mandatory blank→0 conversion** shared by Zig, CLI, C, and Python.
`Value` adds missing_arg / array(*Matrix, non-recursive) / reference(RangeSet)
at evaluator layer. Blank ≠ missing_arg ≠ `""`. Rich errors preserved,
never produced. **Per-form evaluation contracts** (each fixture-pinned with
volatile-draw, error, and dependency-capture tests in every branch position):
`IF` and `CHOOSE` are **lazy for scalar selectors** (taken branch only;
zero runtime draws in dead branches); **array selectors switch to
per-element masking** (both branches evaluate broadcast to the mask shape;
per-element errors stay per-element). **Static vs runtime split
(normative)**: the dependency GRAPH always carries static syntactic edges
for every arm; laziness governs only runtime evaluation and volatile draws.
Mixed-mask IF/IFERROR/IFNA fixtures at M4c; **CHOOSE fixtures at M4e**; `IFERROR`/`IFNA` evaluate the value arg, then the fallback only on
error; **`IFS` and `SWITCH` are EAGER** — Excel evaluates all arms
(observable via volatiles and dependencies), so zlsx does too; `AND`/`OR`
eager.
**Empty matrices are unrepresentable results**: any zero-row/zero-column
array normalizes to `#CALC!` at the producing function's boundary (Excel's
own answer; both modes — never persisted, never streamed). Criteria module
(M3b). Per-run arena.

**5.3b Shape & coercion table (normative, lands with M3a; dialect-indexed)**:

| Context | DA dialect | Legacy dialects |
|---|---|---|
| Scalar where array expected | lift to 1×1 | same |
| Binary op, shapes (r₁×c₁)·(r₂×c₂) | broadcast dims of size 1; result (max r × max c); incompatible (both >1 and unequal) → elementwise `#N/A` fill per Excel (fixture-pinned) | same evaluation, but result consumed per intersection rules |
| **Reference** where scalar expected | no implicit intersection (the reference spills/iterates per function DA-awareness) | **implicit intersection**: same-row → row-projected element; same-column → column-projected; else `#VALUE!` (fixture-pinned) |
| **Array** (non-reference) where scalar expected | spills/iterates per DA-awareness | **top-left reduction** — arrays reduce to their top-left element, NOT row/col intersection (Excel distinguishes references from arrays; fixture-pinned) |
| Reference in value context | dereference (single rect → matrix/scalar) | same + implicit intersection rules |
| `@expr`, expr is a **single-cell reference** | **the reference unchanged** (`=@A1` yields A1's value regardless of the evaluation row/col — Excel's single-item exception precedes intersection) | same |
| `@expr`, expr is a **multi-cell reference** | row/column intersection with the evaluation site: 1-D vector → element on the shared axis; **2-D range → the cell at the intersection of the evaluation row AND column iff the range spans both**; no intersection → `#VALUE!` | identical, explicit spelling |
| `@expr`, expr is an **array** | **top-left element** (`@SEQUENCE(3)` → 1) | same |

Oracle tests per row: operators, scalar-arg functions, array-arg functions,
spill references, legacy formulas, `@SEQUENCE(…)`, `@{1,2;3,4}`,
off-axis `=@A1`, 2-D `@` ranges (spanning and non-spanning),
horizontal/vertical ranges, out-of-range intersections.

**Standalone-eval dialect (normative)**: dialect is a stored-cell property
(`EvalEnv.dialectOf`), so standalone evaluation — which has no stored cell —
takes an explicit `dialect` at **every layer that has standalone eval**
(Zig: `RunInputs.dialect`, default `.dynamic_array`; CLI `--dialect
da|legacy`; C run-inputs field; Python kwarg), echoed in every
resolved-context record and in every cache/fingerprint key. **Spark has no
standalone-eval operation in v1**, so it exposes no dialect option — its
recalc derives dialect from workbook provenance like any stored-cell recalc.
Identical text can legitimately evaluate differently under `legacy`, which
is why the choice is the caller's, never inferred.

**Scalar coercion matrix (normative, M3a)** — provenance × context ×
fidelity. Provenances: numeric/text/bool literal, blank cell, text cell,
bool cell, error. Contexts: arithmetic op, comparison, `&`, function arg by
coercion class. The matrix (every cell = a fixture):

| Provenance ↓ / Context → | arithmetic | comparison | `&` concat | direct fn arg (numeric class) | via range (aggregate class) | lookup key | criteria operand | SORT element |
|---|---|---|---|---|---|---|---|---|
| blank cell | 0 | cross-type rules (pinned) | `""` | 0 | **skipped** | matches blank | blank rules | sorts first (pinned) |
| `""` text | `#VALUE!` | text compare | `""` | `#VALUE!` | skipped | text match | `""` criterion | text order |
| numeric text (invariant grammar) | coerced | **cross-type: number < text < logical; never cross-equal** (pinned) | as text | coerced | **NOT coerced** (Excel ignores text in ranges) | exact-text unless numeric-key (pinned) | coerced | text order |
| locale-flavored text | refusal | text compare | as text | refusal | ignored | text | refusal | text order |
| non-numeric text | `#VALUE!` | text compare | as text | `#VALUE!` | ignored | text | text | text order |
| boolean | 1/0 | logical > text > number (pinned) | `"TRUE"`/`"FALSE"` | 1/0 | **ignored via ranges, counted as direct args** (Excel split, pinned per aggregate) | logical match | logical rules | logical order |
| error | propagates | propagates | propagates | propagates | propagates (unless skip-class) | propagates | propagates | pinned position |

Rules: text in arithmetic parses **through the pinned
invariant grammar only** (plain decimal/scientific forms; `"1"+1=2`); text
that parses only under some locale (e.g. `"1,5"`) →
`FormulaLocaleSensitiveInput` refusal, never a guessed number and never a
guessed `#VALUE!`; non-numeric text in arithmetic → `#VALUE!` (that IS
Excel's locale-independent answer for e.g. `"abc"+1`). Every cell of the
matrix is a fixture. **Ordinary comparisons (`=`, `<>`, `<`, `<=`, `>`,
`>=`) are case-INsensitive on text** (Excel semantics; `EXACT` is the
case-sensitive function): equality and ordering use case-folded comparison
via the full non-Turkic fold (`unicode/casefold.zig` — the single
normative algorithm of `collation_v1`, §5.4b) — `A/a` equal, `ß/ss/SS`
fold-equal (fixtured), no Unicode normalization applied (code points as
stored); divergences from Excel's locale collation recorded per `collation_v1`.

**5.3c Error propagation order (normative)**: operator operands evaluate
left-to-right; the first error encountered is the result. Eager function
arguments evaluate in declaration order; first error wins unless the
registry's propagation class says otherwise (**provenance-aware per
function, never family-wide**: `COUNT` counts numbers only — errors in
ranges neither counted nor propagated, direct scalar args coerce;
**`COUNTA` counts error-valued cells and `""`**; `COUNTBLANK` per its blank
class; `COUNTIF(S)` criteria can match errors — direct/reference/error/
bool/text/`""`/blank fixtures each;
`observe` — ISERROR/IFERROR/IFNA; `per_element` — elementwise array ops keep
errors per cell). Errors inside ranges surface in the deterministic iteration
order of §5.6a (first stored error in area/sheet/row-major order wins for
propagating aggregates). Mixed-error fixtures: IF, AND/OR, SUM over
error-bearing ranges, VLOOKUP, array ops.

**M3a1 decisions (shipped 2026-08-03).** Sixteen points the row left
open or got wrong, pinned here because M3a2 builds directly on them:

1. **Two number constructors, not one.** `ScalarValue.fromNumber`
   asserts finiteness — an evaluator handing over a number it computed
   knows it is finite, and a non-finite one is a bug in the caller.
   `fromArithmetic` converts instead, because N4a's whole point is that
   an overflow *is* a legitimate outcome (`#NUM!`). One constructor
   doing both would either assert on a real overflow or let a
   non-finite value into a `ScalarValue`, and both are worse than
   having two names.
2. **`provenanceOf` returns null for an actual number.** The §5.3b
   matrix has seven rows and none of them is "already a number".
   Returning a row anyway would invite a caller to look up a
   disposition the normative table never stated.
3. **The locale classifier's bias runs one way, deliberately.** A false
   positive turns a case Excel answers with `#VALUE!` — a *successful*
   plane-1 result — into a plane-2 refusal that stops a whole
   recalculation; a false negative only costs a less informative error.
   So `LocaleSensitive` requires a specific shape: a single decimal
   comma, digit groups of exactly three under one consistent separator,
   or currency/percent affixes around an otherwise-invariant number.
   `1.2.3`, `1,23,456`, and `1 23` are not numbers under any locale and
   get `#VALUE!`. Both sides are fixtured.
4. **N1a applies to every ingress except `.cache_import`.** §5.4 pins
   `literal` → N1a and `cache_import` → N1b but leaves
   `text_coercion` / `function_arg` / `criteria` unstated. They round
   like literals — Excel's `="1.2345678901234567"+0` yields the
   15-digit value — and the cached `<v>` is the one input that is
   *already* binary64, where re-rounding would corrupt a value nobody
   asked us to reinterpret. Fixtured across all five ingresses in both
   modes.
5. **Space trimming is a property of the ingress, not of the text.**
   `.literal` and `.cache_import` are stored forms and take their bytes
   exactly; the three coercion ingresses trim ASCII spaces because that
   is what Excel does with user text. `" 1 "` therefore parses on one
   path and not the other — from one parser, which is the point of
   there being only one.
6. **N2's threshold is spec-pinned and provisional.** No committed
   manifest contains a zero-snap case and the Excel oracle leg is
   parked (M1b), so nothing on disk decides it. `2^-48` relative
   (`zero_snap_relative_shift_v1`) snaps the textbook cases —
   `1.1-1.0-0.1` and `0.1+0.2-0.3` — and is a **named constant** so
   pinning it later is a one-line change with a fixture behind it.
   Additive scope is enforced by construction rather than by comment:
   `multiply` and `divide` never call the snap at all.
7. **N5 has to choose between two renderings, and the choice is not
   cosmetic.** Zig's `{d}` and `{e}` share the shortest-round-trip
   digit generation and differ only in layout, so the shortest
   round-tripping text is simply the shorter of the two. Positionally,
   `5e-324` is **326 bytes** — a caller sizing a buffer for "a number"
   would be wrong by an order of magnitude. `formatNumber` therefore
   generates the positional form into private scratch and copies it out
   only when it wins; `format_buf_len` is 32 and is a real bound.
8. **`-0` is where the committed excel-fidelity evidence and §5.4
   disagree — and it is the adapter, not the rule.** §5.4 says `.excel`
   normalizes `-0` → `+0` at publication "(fixture-backed)". The only
   committed excel-fidelity manifest is **LibreOffice**, which records
   `-0` with bits `0x8000000000000000`. LibreOffice is not Excel and
   the Excel leg is parked, so nothing on disk can settle it. §5.4's
   rule is implemented as written; the row is listed in
   `excel_adapter_divergences` and skipped from the excel-leg oracle
   tie **by name**, with the skip count asserted — not silently
   dropped.
9. **Excel-fidelity manifests pin to 15 significant digits, not to
   bits.** LibreOffice records `0.1+0.2` as `0.3` (bits
   `0x3FD3333333333333`): its `<v>` carries the display-rounded value.
   That is a serialization property of the adapter, **not** evidence
   that excel-fidelity arithmetic rounds — §5.4 says publication is
   post-N2/N3 binary64 and says nothing about rounding results. So the
   tie compares excel manifests at 15 significant digits and `ieee`
   manifests bit-exactly. M2 reached the same rule independently for
   its precedence pins.
10. **Subnormals are oracle-decided; most of §5.4 is not.** Both
    manifests record `2^-1074` as `0x1` and `1E-308/10000000000` as
    `0x316A2`, across fidelities. Exactly **three** of the nine
    divergence points are oracle-backed — subnormals, overflow, and
    division by zero — *(M3a2 makes it four: see M3a2 decision 1)* —
    and each point carries an `evidence` field
    saying `oracle` or `spec_pinned` rather than implying support it
    does not have. A test asserts the count.
11. **The Divergence ×2 gate asserts agreement as well as
    disagreement.** A gate that only looked for differences would pass
    if the modes diverged *everywhere*, which is exactly as wrong as
    not diverging at all. Every point carries `must_differ` or
    `must_agree`, both halves are asserted to have fired, and the probe
    array is length-checked against the point array so a rule cannot be
    added without a probe.
12. **`collation_v1` takes the fold as a parameter, and the concrete
    fold is imported only by the test section.**
    `src/unicode/casefold.zig` already belongs to the `zlsx` module's
    package tree (`src/xlsx.zig:25` imports it relatively), so a named
    module rooted on the same file collides the moment `zlsx` imports
    the formula engine — the failure M0 hit with `refs/` and M1a with
    `unicode/xid.zig`. (**M4f moved the file to top-level `unicode/`**,
    which dissolves the collision half of this argument; the parameter
    stays, because injecting the fold is the better design and not
    merely the compiling one.) **Verified on 0.16.0**: a file-scope
    `const` referenced only from a `test` block is not resolved in a non-test
    build, so M3a2+ consumers compile `value.zig` without declaring the
    import and the collision never arises. Injection keeps the
    semantics independent of the build graph regardless.
13. **Fold-equal implies equal for *ordering*, not just equality**, and
    comparing folded sequences gives that for free: `A`/`a` and
    `ß`/`ss`/`SS` come back `.eq` under `<` and `>`. A raw tie-break
    would break it, which is why SORT/SORTBY's tie-break lives outside
    the comparator under its own name (`sort_tiebreak_policy`).
14. **Byte order over folded UTF-8 *is* code-point order**, so the
    comparator needs no decoding pass. A property of the encoding,
    recorded so nobody later "fixes" it into a slower loop.
15. **`collation_v1` deliberately does not normalize, unlike sheet-name
    dedup.** `casefold.excelSheetNameEql` applies NFC because
    composed/decomposed sheet names should dedup; §5.3b says text
    comparison uses code points as stored. Both behaviours are fixtured
    side by side — `café` precomposed vs decomposed is *equal* as a
    sheet name and *unequal* as text — so the difference reads as
    intentional rather than as a missing call.
16. **The shape and coercion tables ship as executable lookups.** Eight
    shape rows × two dialects and the 7×8 = 56 coercion cells are
    `switch` tables, with a fixture asserting every cell and a coverage
    check proving that neither the table nor the fixture can lose a row
    without the other noticing.

**Still open (needs the parked Excel oracle leg):** N2's zero-snap
*threshold* (decision 6) and the `.excel` signed-zero policy
(decision 8). M3a2 turned the snap's **existence** into oracle-backed
evidence (M3a2 decision 1), but the threshold is now only bounded
below — `> ~1.85e-16` relative — and nothing on disk bounds it above;
the signed-zero policy is unchanged and still `spec_pinned`. Neither
can be confirmed until the M1b Excel adapter runs.

**Deferred to M4f — CLOSED (2026-08-04):** moving
`src/unicode/casefold.zig` to top-level `unicode/` (decision 12). The
file now sits at `unicode/casefold.zig` beside `nfc.zig`, `xid.zig` and
M4f's new `casing.zig`; `zlsx` imports it by name like every other
consumer, and `src/unicode/` no longer exists.

**M3a2 decisions (shipped 2026-08-03).** Seventeen points the row left
open, got wrong, or discovered — pinned here because M3b, M4b1, and
every F-batch build directly on them:

1. **The oracle decides a rule §5.4 never stated, and it upgrades N2
   from text to evidence.** `(0.1+0.2)=0.3` is recorded **TRUE** by the
   LibreOffice excel-fidelity manifest and **FALSE** by the hand-spec
   `ieee` one. A comparison is a subtraction against zero, which puts it
   inside N2's *additive* scope — and that reading is the only one under
   which both committed manifests are satisfied at once. Numeric
   comparison therefore routes through `applyZeroSnap`, the N2
   divergence point's `evidence` flips to `oracle` (three → four), and
   M3a1's "no committed manifest contains a zero-snap case" is
   superseded. This is the **one change M3a2 makes to `value.zig`**: one
   enum literal and the count assertion that guards it. Leaving the
   label at `spec_pinned` would have understated evidence that is now on
   disk, which is the failure mode the `evidence` field exists to
   prevent.
2. **`TRUE` and `FALSE` were missing from the frozen v1 list.** The
   ladder's M4c row enumerates 18 names; the committed manifests contain
   `TRUE()+1`, which M1a already tokenizes as a *call* (`.name`,
   `paren_open`, `paren_close`) and not as a boolean literal. A row the
   oracle can decide and the registry cannot answer is a gap, so both
   names join F1a-1: M4c goes 18 → 20 and the Core gate 57 → 59. The
   inventory file is the count source, so the ladder row was corrected
   to match it rather than the reverse.
3. **The frozen inventory is a committed TSV, not a Zig array.**
   `src/formula/function_inventory_v1.tsv` — 175 rows of
   `name<TAB>milestone<TAB>batch`, sorted, unique, tagged. Data because
   §7 makes it the authoritative count source for *every later PR*:
   regenerating a number from a file a shell can read is a different
   activity from re-reading a table someone has to compile. Three tests
   guard it — the total, per-milestone counts against the ladder, and
   strict ascending order (uniqueness and sort in one assertion).
4. **A formula's result is a value, so a top-level reference
   dereferences through the same §5.3b row as an operand.** `=A1:A3`
   spills under `dynamic_array` and intersects under `legacy` because
   `reference_in_value` says so, not because the top level is special.
   The one thing the top level *does* add is §10's multi-area refusal,
   checked before the dereference so `(A1:A2,C1:C2)` refuses while
   `SUM((A1:A2,C1:C2))` still works.
5. **Runtime dependency capture happens where a reference is
   CONSTRUCTED, not where it is read.** `IFS` is eager, so its untaken
   arms evaluate to references it never dereferences — and capturing at
   the point of read would report them as unread, which is exactly
   backwards. §5.3a's split is between *evaluated* and *not evaluated*;
   capture is placed to measure that.
6. **The volatile-draw counter lives in the seam.** `DrawSource.draw`
   increments before it calls out, so "zero draws in the dead branch" is
   a property of the evaluator rather than of how carefully a fixture
   was wired. M3b replaces the callback with `rng_v1`; the counter and
   its meaning do not move.
7. **Eagerness is asserted positively.** `IFS(FALSE,RAND(),TRUE,RAND())`
   draws **2** and `AND(FALSE,RAND()>2)` draws **1** — fixtures that
   fail if anyone "optimizes" a short-circuit in. A test that only
   checked results would pass either way.
8. **Coercion classes are behaviour, not documentation.** The dispatcher
   applies each slot's class before the implementation runs, so `SQRT`
   receives a number or an error and never re-derives a coercion. This
   is what keeps `SQRT(-1)` a statement about the radicand — `#NUM!` —
   with no possibility of it being a missing coercion in disguise.
9. **Whether an array lifts is the dialect's decision, not the
   function's.** `SQRT({4,9})` is `2` under `legacy` (the table's
   `top_left_reduction`, which is a reduction and explicitly *not* an
   intersection) and `{2,3}` under `dynamic_array`. A mixed signature —
   a range slot beside a scalar one — refuses as `NotYetImplemented`
   rather than guessing ahead of M7a's decision table.
10. **`SQRT(-1)` is a second excel-adapter divergence, and the two
    excel-fidelity manifests disagree with EACH OTHER about it.** The
    hand-spec suite records `#NUM!` (Excel's answer); LibreOffice
    records `#VALUE!`. That disagreement is itself the proof that the
    row is about the adapter and not about the rule, so it joins `-0` in
    a named `excel_adapter_divergences` list with the skip count
    asserted — never silently dropped.
11. **Wrong arity is `FormulaMalformedInput`, not `#VALUE!`.** Excel
    cannot store a formula whose argument count its own registry
    rejects, so such input did not come from a workbook. An unregistered
    *call* is `FormulaUnsupportedFunction` (§7) and an unresolvable
    *name* in value position is a plane-1 `#NAME?` (§5.9) — three
    different answers to three different questions.
12. **Not-yet-implemented is its own named refusal with an enumerated
    membership.** `NotYetImplemented` maps to
    `FormulaUnsupportedConstruct` but is listed separately in
    `not_yet_implemented`, currently one entry: 3D sheet spans (M4b3).
    A later row deletes a line and watches a test fail until it does —
    which conflating the two would have made impossible.
    **Settled at M4b3**: the row deleted that entry, the list is empty,
    and the watching test is the one that used to expect
    `NotYetImplemented` from a span. M4b3 also split off
    `UnsupportedConstruct` for what v1 refuses *on purpose* (§5.6g's
    ineligible 3D contexts, §5.9's disqualified names, a table
    reference), so the enumerated list keeps meaning "not at this row"
    rather than drifting into "never".
13. **The five required registry fields carry no defaults, and a test
    reads `@typeInfo` to prove it.** An omitted propagation class is how
    `COUNTA` quietly becomes `COUNT`. Three comptime checks back it up:
    the two parallel per-slot tables must agree on slot count, a
    `.plain` form must have an impl and no lazy slot, and a deferring
    form must have a lazy slot and no impl.
14. **The empty matrix normalizes at the call boundary and nowhere
    else**, and is tested by invoking a test-local `Function` whose impl
    returns `error.EmptyMatrix` — rather than by adding a fake entry to
    the shipped table. The first v1 function that can produce one is
    `FILTER` at M7a; the boundary is ready before it.
15. **The fake implements the merge, not a stub of it.** Entries sort by
    `(row, col, layer descending)` on insert, so ordered iteration is
    independent of insertion order *by construction* — there is no code
    path that could return backing order because there is no backing
    order — and a merged read is a lower bound rather than a scan. Both
    iterators carry their state inline behind a compile-time size bound,
    so iteration allocates nothing.
16. **Zig runs an imported file's tests only once something in it is
    referenced.** `eval.zig`'s test artifact reported *0 tests* while
    the file had no test of its own, silently taking `registry.zig`'s
    eleven with it. `test { _ = registry; _ = env; }` is not decoration.
17. **Operator chains are folded iteratively, and the evaluator's depth
    limit is its own field.** Every value operator is left-associative,
    so `1+1+1+…` parses as a tree as deep as it is long — and a
    5 000-term sum, comfortably inside Excel's 8 192-character limit,
    overflowed the stack when the walk recursed into it (caught by a
    fixture, not by a review). `Evaluator.binary` now descends the left
    spine on the heap and folds it back, which preserves §5.3c's
    left-to-right operand order exactly while leaving recursion bounded
    by *parenthesis* nesting — something §9 already bounds at 256. The
    residual `max_expr_depth` is a separate `Options` field. **Three
    depths exist and none substitutes for another**:
    `Limits.max_parse_depth` (256, recursing grammar productions),
    §9's `max_eval_depth` (512, dependency-closure recursion — M5a's
    graph, not an expression), and `max_expr_depth` (AST nodes on the
    stack). Reusing the parser's 256 would have refused a 300-term sum
    that parsed perfectly well; reusing §9's name would have collided
    two unrelated quantities. §9 gains a row for the new one.

**Spec-pinned at M3a2 (no committed manifest decides them):** `IF`'s
omitted third argument is `FALSE`; `CHOOSE` truncates its index toward
zero and is `#VALUE!` outside `1..n`; `IFS`/`SWITCH` with no match are
`#N/A`; `AND`/`OR` over no logical value at all is `#VALUE!`; a
disjoint `intersect` is `#NULL!`; a multi-area reference in a *value*
context is `#VALUE!`; `A1#` against a non-anchor is `#REF!`; a
reference to an unknown sheet is `#REF!`.

### 5.4 Numeric & text fidelity

**`excel_fp_rules_v1`** (fixture-pinned, oracle build recorded): N1a 15-digit
literal ingress (`.excel`) / N1b full-binary64 `<v>` import / N2 zero-snap (**additive scope only**: applies to `+`/`-` results near zero — never multiplication, division, or function results; counterexample fixtures for each) /
N3 subnormal & `-0` / N4 per-function quirks / **N4a finiteness** (overflow →
`#NUM!`, `0/0` → `#DIV/0!`; asserted at returns + pre-serialization) / N5
shortest-round-trip (+`fitsExactlyInF64`). Publication = post-N2/N3 binary64.
**One decimal-ingress primitive**: `parseDecimal(fidelity, ingress)` where
`ingress ∈ {literal (N1a), cache_import (N1b), text_coercion, function_arg,
criteria}` — the ONLY decimal-text→f64 path. **Caller→ingress table
(normative)**: formula literals → `literal`; **present numeric `<v>` values only** → `cache_import` (SST indices parse as bounded integers, and numeric-looking SST *content* stays text until §5.3b coercion — there is no SST-number ingress path); §5.3b arithmetic coercion → `text_coercion`;
`VALUE`/`NUMBERVALUE`/`DATEVALUE`/`TIMEVALUE` components → `function_arg`;
criteria operands → `criteria`. Paired fixtures push identical decimal text
through every path in both fidelity modes; a second parser diverging is
structurally impossible.

**`ieee_fp_rules_v1` (normative — the `.ieee` mode is a contract, not an
absence)**: literals convert binary64-nearest (full text, no truncation);
every operation is IEEE-754 binary64 round-to-nearest-even; **no zero-snap**;
**signed zero preserved bitwise**; **subnormals preserved**; N4 per-function
quirks still apply (they are correctness, e.g. `MOD` sign); N4a finiteness
still applies (overflow → `#NUM!` — Excel's value domain is shared by both
modes); N5 serialization identical. **Signed-zero comparison policy**:
`.excel` normalizes `-0` → `+0` at cell publication (fixture-backed);
`.ieee` preserves it; §8.3's bit-exact comparisons apply the mode's policy
consistently for scalars, arrays, cached values, and goldens. M3a's
"Divergence ×2" gate runs every divergence-point fixture under both rule
tables.

**5.4a Serial dates**: 0–60 domain (60 = fictitious 1900-02-29; 0 =
1900-01-00), 1904 clean; boundaries pinned; reader contract untouched.
Landed M3b as `src/formula/serial_date.zig` — **the 1900 epoch is two
epochs** (serials 1–59 count from 1899-12-31, 61+ from 1899-12-30) and
both invented days are representable rather than errors; see the M3b
decisions block in §5.5. **`WEEKDAY` counts serials, not calendar days
(M4g decision 1)**: the phantom day makes the two disagree by one
everywhere below serial 61, and Excel counts serials — `WEEKDAY(1)` is
Sunday though 1900-01-01 was a Monday. The 1904 system has no phantom
and no drift, only a different phase.

**5.4b Locale, collation, platform**: locale-sensitive **parses**
(VALUE/DATEVALUE/TIMEVALUE/TEXT/NUMBERVALUE) carry pinned invariant grammars;
outside-grammar input → `FormulaLocaleSensitiveInput` refusal (never a
fabricated `#VALUE!`). **The line the grammar draws (M4g decision 2): a
spelling refuses exactly when the locale would change what it MEANS, not
merely how it looks.** So `"1.5"` parses and `"1,5"` refuses; `DATEVALUE`
accepts ISO and every named-month form unconditionally, accepts `M/D/YYYY`
whenever one field settles the order (`"1/15/2020"` can only be January
15th anywhere), and refuses `"1/2/2020"`, which is two dates. `TIMEVALUE`
has no such refusal — a clock has one field order everywhere. **Locale-sensitive OUTPUT is pinned invariant in v1**
— there is deliberately no locale field in `RunInputs`: `TEXT`/`FIXED`/
`DOLLAR` render invariant en-US forms and `NUMBERVALUE`'s omitted separators
default to `.`/`,` — each a recorded, fixtured divergence from Excel's
locale behavior; a locale profile is an M10+ addition that would join every
fingerprint/cache key when it lands. **`collation_v1`** — ONE comparator, stated once: **lexicographic order of
full-non-Turkic-folded code-point sequences** (`unicode/casefold.zig`;
`ß` folds to `ss`) governs ordinary `=` `<>` `<` `<=` `>` `>=`, SEARCH,
wildcards, lookup equality AND ordering, criteria, and SORT/SORTBY.
**Positional matching over expanding folds**: the fold keeps a
folded-unit→original-unit map; `?` consumes **one code point — version-INdependent** (CV changes exactly
the index units of LEN/MID/FIND/SEARCH/REPLACE **and, corrected at M4f,
LEFT/RIGHT — seven, not five**: a count of characters cannot mean UTF-16
code units in MID and code points in LEFT inside one workbook (M4f
decision 1) — wildcards,
criteria, COUNTIF/MATCH/XMATCH are NOT CV-dependent; oracle-pinned);
`*`/`~` operate on original units; SEARCH returns positions in its
CV-dependent unit — fixtures: `ß` expansion, ligatures, combining marks,
astral. Comparator exceptions —
**MIN/MAX are NOT in this list** (they never compare text: direct text args
coerce-or-`#VALUE!`, text in ranges ignored, no numbers → 0 — §5.3b row,
fixtures); **one total preorder** governs ALL text ordering and caseless equality:
compare **full-folded code-point sequences, nothing else** — fold-equal
strings (`A`/`a`, `ß`/`ss`/`SS`) are EQUAL for `=`, `<>`, `<`, `>`, lookups,
and sorting semantics alike (a raw tie-break would make them unequal and is
therefore *not* part of the semantic order; the shipped fold primitive
defines equality this way, `unicode/casefold.zig:53-70`). SORT/SORTBY
use **stable source position** as a private, non-semantic tie-break among
fold-equal elements. Registry-level **match policies**: `.folded` (default — `=`, SEARCH, wildcard consumers), `.raw` (**FIND and SUBSTITUTE are case-sensitive**, like EXACT; CODE/UNICODE raw), `.arg_selected` (TEXTBEFORE/TEXTAFTER/TEXTSPLIT via `match_mode`). Every text function's policy is explicit registry data. Each divergence from Excel's
locale collation recorded; `ß`/`SS`/`ST` ordering fixtures included. Registry metadata flags every collation-touching
function. **`casing_v1` (module + SpecialCasing generator IN M4f with LOWER/UPPER; M8b adds only PROPER's word segmentation)**: versioned, locale-neutral **full Unicode casing** (**UnicodeData simple mappings + unconditional SpecialCasing + the locale-independent Final_Sigma rule; locale-conditional (tr/az/lt) rejected** — the generator stops before UnicodeData casing fields, `gen_unicode_tables.py:138-172`; SpecialCasing-backed — `ß`→`SS` is a length-changing full mapping that simple one-to-one mappings cannot express; a vendored SpecialCasing table joins the existing generator, `scripts/gen_unicode_tables.py`, which today emits only case-fold + NFC). Fixtures: every length-changing mapping, dotted/dotless I (invariant; Turkish divergence recorded), `ß`→`SS` (oracle-pinned), ligatures, combining marks, astral letters. `CHAR`/`CODE`: `platform_profile` (default `.windows_1252`);
Mac-oracle exclusion 128–255 + CP-1252 spec cases.

**5.4c Multi-area results**: a top-level result that is still a multi-area or
multi-sheet reference after evaluation → **typed refusal
`FormulaResultNotRepresentable` at every public layer in v1** (single rects
dereference; in-formula consumption of unions/intersections is unaffected;
Excel-match for top-level multi-area display is M10+ if demanded).

**5.4d Compatibility Version — BOTH versions in v1**: the 2024 workbook CV
extension changes **surrogate-pair (code-point) handling** — CV2 treats a
surrogate pair as one character in LEN/MID/FIND/SEARCH/REPLACE **plus
LEFT/RIGHT (M4f decision 1)** (not grapheme
clustering; variation selectors/modifiers stay separate). **A CV1 index
that falls between the halves of a surrogate pair is
`FormulaResultNotRepresentable`** — Excel returns a lone surrogate, which
UTF-8 cannot carry (M4f decision 2). **CV2 is the
default for newly created Current-Channel workbooks since April 2026**, so
refusing it would refuse normal files. Parsed + preserved (M4b2);
`CalcState.text_compat` v1|v2; **absent CV metadata = CV1** (matches pre-CV workbooks; fresh Writer emits no metadata part, `fresh_emit.zig:45` → CV1; CV2 authoring for fresh files = M10+; oracles: absent/CV1/CV2/fresh); **both semantics implemented in M4f** (the
shared text layer counts UTF-16 code units for CV1 and code points for CV2
in the five affected functions); oracle fixtures authored under each version
where the host Excel permits, hand-spec cases otherwise (astral-plane
fixtures for both).

### 5.5 Run inputs vs workbook calc state

```zig
pub const RunInputs = struct {
    now_utc_ms: i64,
    utc_offset_min: i32 = 0,         // ABI int32_t; validated [-1440,1440]
                                     // pre-narrowing (AGENTS.md:107).         // fixed civil offset for NOW()/TODAY().
                                     // DEFAULT IS UTC AT EVERY LAYER — Zig 0.16
                                     // stdlib has no portable local-tz resolver
                                     // and zlsx is stdlib-only; callers pass an
                                     // explicit offset (CLI --utc-offset) or get
                                     // UTC, documented. TZif is M10+. DST and
                                     // midnight boundary fixtures both epochs.
    rng_seed: u64,                   // rng_v1 (xoshiro256**). Seeded randomness
                                     // is an INTENTIONAL DIVERGENCE from Excel in
                                     // both modes (Excel's RNG is unspecified);
                                     // RAND-family Excel-fidelity tests check
                                     // shape/type/range/same-context repeatability
                                     // only — never sequences.
    fidelity: FidelityMode = .excel,
    platform_profile: PlatformProfile = .windows_1252,  // v1 enum CLOSED: {windows_1252}.
                                     // --now: RFC 3339, 'Z'/numeric offset, seconds
                                     // precision, years 1900-9999; timestamp-offset
                                     // vs --utc-offset conflict = error.
    dialect: EvalDialect = .dynamic_array,  // standalone eval only. Effective-input
                                     // PROJECTIONS differ by operation: standalone
                                     // eval keys/echoes dialect; recalc derives
                                     // dialect per stored cell (EvalEnv.dialectOf),
                                     // NORMALIZES this field out of its fingerprint
                                     // and omits it from echoes — no phantom key.
    limits: Limits,
    deadline: ?std.Io.Timestamp = null,  // .awake clock of the SAME std.Io
                                     // (Zig 0.16 has no 'Instant'; Io.zig:778,793);       // absolute monotonic (std.Io.Clock.now(.awake, io));
                                     // same safe points as cancel; no helper thread;
                                     // first to fire wins. EXCLUDED (with cancel) from
                                     // EffectiveRunInputs — never fingerprinted/keyed;
                                     // an optional requested-timeout duration is echoed
                                     // separately; byte-determinism is scoped to runs
                                     // that complete uncancelled.
    cancel: ?CancelToken = null,     // no-alloc union over the two storage
                                     // kinds: *const std.atomic.Value(bool)
                                     // (multi-threaded) | *const volatile
                                     // sig_atomic_t-style flag (signal-safe,
                                     // single-threaded CLI); one
                                     // isTriggered() seam. Zig surfaces
                                     // cancellation as error.Cancelled in
                                     // RecalcError/EvalError, mapping to C
                                     // -5 / CLI 130-143-3.
                                     // EXCLUDED from EffectiveRunInputs: never part of
                                     // fingerprints, cache keys, or echoes — cancellation
                                     // appears only in terminal/report status.
                                     // cooperative cancellation:
                                     // polled between cell evaluations and at
                                     // bounded work intervals inside long ops;
                                     // §5.7.9 defines the non-cancellable commit
                                     // region. CLI --deadline arms it; C ABI
                                     // exposes a token (§12.3); Python maps
                                     // timeout= / KeyboardInterrupt onto it.
};
pub const CalcState = struct {       // WORKBOOK-DERIVED, never caller-writable
    date_system, iteration, full_precision, text_compat,
};
pub const EvalSite = struct { row_1based: Row1, col_0based: Col0 };
```

Adapters resolve defaults once and echo resolved values (report, manifests,
Spark plan) — **per-layer default table (normative)**: CLI — `now` = one
`std.Io` wall-clock read at startup (ms), `seed` = 8 bytes from the
`std.Io` random source (**entropy or clock failure → exit 6**, its own status — exit 4 stays exclusively allocator failure, §12.2), offset 0; Python — same
sources once per call; Spark — driver values; **Zig and C have NO defaults
— `RunInputs.now_utc_ms` and `.rng_seed` are required caller-supplied fields
and the library never reads a clock or an entropy source on its own; only the
CLI, Python, and Spark adapters resolve omitted values**; every layer echoes its exact
resolved values: CLI `--now/--utc-offset/--seed/--mode/--profile/--dialect`;
Python kwargs; Spark `zlsx.recalc` (activation, `"true"`), `zlsx.recalcNow`,
`zlsx.recalcUtcOffsetMin` (**default 0 = UTC**, like every layer),
`zlsx.recalcSeed`, `zlsx.recalcMode`, `zlsx.recalcProfile` (defaults:
driver instant / **UTC** / random-once / excel / windows_1252; no dialect
option — Spark has no standalone eval, §5.3b). Typed coordinates throughout (`SheetIndex`, `Row1`, `Col0`;
conversions at named boundaries only). `current_sheet` required at every
layer (CLI: shipped `--sheet N | --name NAME` grammar, one mandatory for
`eval`; Python: required positional). A1 never shifts at eval; `EvalSite`
only for argless ROW/COLUMN, `@` intersection row/col, `#This Row`; missing →
`FormulaAnchorRequired`. Determinism: equal RunInputs + bytes + build +
target ⇒ byte-equal; cross-target CI goldens; wall-clock only as cancellation.

**M3b decisions (shipped 2026-08-03).** Eighteen points the row left
open, got wrong, or discovered. They span §5.4a (serial dates), §5.3b
(criteria), §5.5 (run inputs), and §9 (the byte budget), and are
collected here because that is the section whose subject — what a run is
*given* — covers all four.

1. **The library reads no clock and no entropy source, and the type
   system says so.** `now_utc_ms` and `rng_seed` are required fields
   with no defaults, and a test reads `@typeInfo` to prove every other
   field has one. A default would mean the library resolved an input on
   its own, and "equal inputs ⇒ equal output" would be false in a way no
   test could see.
2. **`ResourceLimits` is a second limits struct, not an extension of
   `parser.Limits`.** §9 holds three different kinds of limit: parse
   *shape* (M2, enforced while parsing), aggregate *bytes and counts*
   (this row, enforced by an allocator wrapper), and *work* counters —
   cell evaluations, SCC passes — which can burn CPU without allocating
   and therefore need explicit counters (M5a2). Three mechanisms, three
   structs; merging them would put an enforcement point in a struct that
   cannot reach it.
3. **An exhausted budget is a refusal, not an allocation failure.**
   `std.mem.Allocator` can only say `OutOfMemory`, so `Budget` records
   the **first** category to trip and the evaluator maps that onto
   `FormulaLimitExceeded`. Without the record, "your formula asked for
   too much" and "this machine is out of memory" are the same error, and
   only one of them is the caller's to fix.
4. **Nothing mutates on a refusal.** The charge is checked *before* the
   backing allocator is called, so a rejected allocation leaves the
   counter and the heap exactly as they were. Below/at/above boundary
   tests for all five categories.
5. **`matrix_cells` is charged as a count, through an explicit
   `charge`, not through the allocator.** §9 bounds *cells* so that the
   limit means the same thing whatever `@sizeOf(ScalarValue)` happens to
   be; routing it through byte accounting would silently re-scale it.
6. **Three of the five categories have a charge site at M3b, and the
   gap is named rather than hidden.** `run_arena` (the arena is built
   over the budget), `string_payload` (concatenation, number formatting,
   string-literal unescaping — borrowed text is *not* charged, because
   the run did not create it), and `matrix_cells`. `retained_asts` and
   `diagnostics` have no producer inside the evaluator: ASTs belong to
   the caller and the diagnostic sink is M6's. The **mapping** is proven
   for all five; the charge sites arrive with their producers.
7. **The 1900 epoch is two epochs, and is written as two constants.**
   Serials 1–59 count from 1899-12-31 and serials 61+ from 1899-12-30.
   Expressing that as one epoch plus a correction is the same arithmetic
   wearing a disguise; two named constants make the discontinuity
   impossible to lose, and a comptime assertion pins their difference at
   exactly one day.
8. **Both invented days are representable, not errors.** Serial 0 is
   `1900-01-00` and serial 60 is `1900-02-29`, a date that never
   happened. Both appear in real workbooks, so a converter that refused
   them would refuse files Excel opens. A `Fictitious` tag names which
   one, so a caller that ignores it still gets a sensible `year`/
   `month`/`day` rather than an out-of-band sentinel.
9. **The round trip is tested over the whole domain, not a sample.**
   ~3 M serials per system, walked in milliseconds. Every interesting
   failure in date arithmetic is within two days of a boundary, which is
   exactly what sampling misses. A separate monotonic-and-gapless test
   covers the failure a round trip alone permits: two serials mapping to
   one date.
10. **Excel's documented maxima are asserted against the calendar, not
    trusted.** `2958465` and `2957003` are checked at comptime against
    `daysFromCivil(9999,12,31)`, and their 1462-day difference — the
    constant every 1900↔1904 conversion turns on — is pinned too.
11. **A criterion is not a comparison: it is type-restricted.** `">5"`
    never matches a text cell, even though §5.3b's total preorder puts
    every text above every number. Applying the preorder here would make
    `COUNTIF(A:A,">5")` count words — wrong, and the kind of wrong that
    looks right in a small fixture. **Spec-pinned**: no committed
    manifest contains a criteria row.
12. **Wildcards need a folded-unit → original-code-point map, and
    per-code-point folding is what produces it.** §5.4b says `?`
    consumes one code point of the *original*, so `"?"` matches `ß`
    while `"?s"` does not, and `"ss"` matches it because a literal run
    may span an expansion — provided it lands on a code-point boundary,
    which is what stops `"s"` from matching half a folded `ß`. The
    non-Turkic full fold is defined per code point, so folding
    code-point-wise and folding the whole string agree byte for byte
    (fixtured over a 16-string corpus) and only the former yields the
    map. No change to `casefold.zig` was needed.
13. **"Has a wildcard" and "needs the pattern matcher" are different
    questions.** The first implementation conflated them and
    `COUNTIF(range,"~*")` silently failed to match a literal `*`: the
    escape still has to be stripped even when no wildcard survives it.
    The field is `is_pattern`, and `hasWildcards` is tested separately.
14. **Blank positions are scanned as runs, and whether a blank
    satisfies the criteria is computed once per scan.** It is the same
    answer at every position in a run; computing it per position is
    exactly what would make `COUNTIF(A:A,"")` walk 1 048 576
    coordinates. `ScanResult.visited` counts every position covered, so
    a run-based cursor can be audited for double-counting.
15. **`ScanResult` must not be blanket-reset.** `result.* = .{}`
    discarded the caller's `hits` storage — found by a fixture reading
    back `0xAAAAAAAA`, not by review. The reset preserves it.
16. **`rng_v1` is written out rather than taken from
    `std.Random.DefaultPrng`.** The plan names xoshiro256\*\*; the
    stdlib's default is a different member of the family, and a stdlib
    default is free to change between Zig releases — which is precisely
    what a *versioned* generator must not do. Twenty lines of explicit
    state is the cheaper guarantee.
17. **The KATs were cross-checked against an independent
    implementation, not recorded from this one.** A golden generated by
    the code it guards proves only that the code agrees with itself.
    `splitmix64(0)`'s published first output `0xE220A8397B1DCDAF`
    anchors the seeding half from outside the repository.
18. **`deadline` and `cancel` are absent from `EffectiveRunInputs` by
    construction, and the dialect projection differs per operation.** A
    field that is not in the struct cannot leak into a cache key, so the
    exclusion needs no filter anyone could forget. `effective` takes an
    `Operation` because standalone eval keys on the caller's dialect
    while a recalc derives it per stored cell and normalizes it out —
    keying on it there would split identical work across two entries.

**Spec-pinned at M3b (no committed manifest decides them):** the
type-restriction rule above; `""` matches true blanks and `""` cells
while `"<>"` matches everything else; an error *spelling* (`"#N/A"`) is
an error operand and matches error cells, which is why `COUNTIF` can
match what `COUNT` cannot count; a text criterion never matches a
logical cell and vice versa (§5.3b's cross-type rule); a trailing `~`
escapes nothing and is dropped; wildcards are inert under an ordering
operator.

**Also verified at M3b:** criteria scanning reads through `EvalEnv`
directly rather than through `Evaluator.readRange`, and the areas are
still captured as runtime dependencies — because M3a2 records a
reference where it is *constructed*, not where it is read. Fixture-pinned
rather than assumed.

### 5.6 Graph, order, cycles, dynamic refs

**5.6a `EvalEnv`** (M3a; typed coords; explicit error sets):
`cellValue(SheetIndex, Row1, Col0)`; `rangeIterator(range)` — a **sparse
LOGICAL iterator (normative)**: one ordered pass merging (a) stored cells,
(b) staged deltas/appends of the logical view, (c) values computed earlier in
this run, and (d) **virtual spill tails** — cells materialized by an
already-evaluated dynamic-array anchor that remain scratch-only until
transactional staging (§5.8). A stored-cells-only interface would make
`C1=SUM(A1:B2)` read stale or blank inputs for a 2×2 spill computed at A1 in
the same run. `cellValue` reads the same merged view. Layer precedence:
computed > staged > stored, with spill tails owned by their anchor
(§5.8a ownership). Iteration stays sparse (no 1M-coordinate walk) and is
ordered per area → sheet → row-major, independent of backing insertion order
and of which layer supplied a cell (randomized-insertion test);
**same-run spill-grow / spill-shrink range-dependency tests at M7a** — a
grow must be visible to dependents evaluated after the anchor, a shrink must
clear its tails from every subsequent read; **`logicalBlankCount(range, class)`** with **operation-specific blank
classification** (Excel's own split): `.isblank_class` = true blanks only;
`.countblank_class` = true blanks **plus cells whose (cached or computed)
value is the empty string** — COUNTBLANK counts `=""` results while
ISBLANK is false and COUNTA is true for them (fixture matrix: true blank /
`=""` / `0` / error, over sparse full columns); **both classes count over
the SAME merged logical view as `rangeIterator`** (a cell occupied only by a
same-run spill tail is not blank in either class — fixture-pinned); both
computed without walking 1M coordinates (A:A benches);
**criteria alignment (normative)** — `SUMIF`/`AVERAGEIF` **project** a
differently-sized sum/average range from its top-left cell using the
criteria range's dimensions (Excel's documented rule — the ranges need not
be same-shape); `*IFS` functions require equal-dimension ranges (else
`#VALUE!`, pinned) and use an **N-way sparse cursor** (not repeated pairwise
zips) over all criteria ranges + the aggregation range, with logical blanks
as runs. Fixtures: unequal SUMIF ranges, 3+-criteria SUMIFS, full-row and
full-column ranges; whole-column perf tests; lands before any criteria
function that takes a separate range; multi-area union overlap double-counting
fixture-pinned; `resolveSheet/Name/Table`, `spillShape`, `dialectOf`,
`calcState`. In-memory fake (M3a); pkg adapter (M4b1).

**5.6b Nodes** (M5a1, shipped): formula cells; **spill-tail nodes**
(each depending on its anchor, so a shape change is a node-set change);
range nodes, one per distinct area, which is what makes the edge count
readers **+** producers rather than readers × producers; **3D span
nodes**, one edge per member sheet (§5.6g); name nodes; table producers.
A reference to a coordinate holding no formula contributes **no edge** —
a constant cannot be recalculated, so it cannot constrain an order.
Areas resolve against a producer index sorted **twice** (row-major and
column-major), probed through whichever order makes the narrower band
the leading key — the corrected form of "interval buckets", and what
keeps `SUM(A:A)` and `SUM(1:1)` both proportional to what is stored.
Order: SCC condensation, then **Kahn with a min-heap** over the
condensation DAG — canonical, not merely reproducible — tie-break
(SheetIndex, Row1, Col0), generalized across kinds by a total order on
node keys whose every term is content rather than an insertion ordinal.

**5.6c Multi-SCC iteration schedule (normative)**: build the condensation DAG;
process components in topological order; an **SCC iterates to its own
convergence before any downstream node evaluates** (downstream sees final
values only); each SCC gets its own pass counter bounded by **both** the
semantic `iterateCount` (clamped to Excel's max 32 767) **and** the resource
ceiling `max_scc_iterations` — whichever binds first decides the outcome, and
the two outcomes differ (see the exhaustion rule below); defaults off/100/0.001;
**the transition table, pinned (M5a2)**: `iterateCount` 0 → 100 (Excel's own
minimum is 1, so a zero is an unset attribute and an unset attribute is the
schema default) · `iterateCount` > 32 767 → 32 767 · `iterateDelta` < 0 → 0.001 ·
**`iterateDelta` = 0 → 0.001, the SAME row as negative** — under a strict
`< iterateDelta` a zero tolerance is satisfied by nothing at all, not even by a
value that did not move, so it is an unsatisfiable bound rather than a request
for exact equality; reading it as exact equality would need an exception inside
the comparison, and a comparison with an exception in it is one two
implementations will eventually disagree about · non-finite `iterateDelta` never
reaches the table (`calc.parseCalcState` refuses the part). Gauss–Seidel
visibility inside a pass. **Order divergence, declared**: Excel iterates via a mutable
calculation chain whose order evolves during recalc; zlsx's fixed
coordinate-order Gauss–Seidel is an **intentional documented divergence**
(determinism requires it), gated by order-sensitive circular-workbook
fixtures that record where converged values differ from Excel's and assert
convergence itself agrees. Convergence per cell: numbers **`abs(new − previous) < iterateDelta`**
(magnitude — a raw signed difference would instantly 'converge' any
decreasing value; increasing/decreasing/sign-crossing + boundary + signed-
zero fixtures); text/bool/blank/errors two-pass equality; arrays: shape equality AND per-element convergence — same-type elements
use their type's rule; **any type transition or shape change = not
converged** (mixed/oscillating/error-transition fixtures; NaN impossible
per N4a). **SCC
seed table**: numeric cache → value; text/bool/error cache → 0; **absent `<v>` (uncached formula) → 0; malformed `<v>` → pre-mutation typed refusal, never a zero seed** (a malformed cache cannot be both 'unparseable input' and a number); array anchors **with a declared/metadata shape** → zero-filled
that shape; array anchors with **no recoverable shape** (e.g. newly authored
dynamic formulas) → a pre-iteration **shape pass** evaluates the anchor once
outside the cycle to fix its shape; if the shape then changes between
iterations → `#SPILL!` (indeterminate) — shape-mutating cycles never spin.
First-run vs resume pinned (`A1=(A1+1)/2` parity fixture). **Two distinct
exhaustion outcomes — semantic bound vs resource ceiling (normative)**: the
**semantic** bound is the workbook's own `calcPr@iterateCount` (clamped to
32 767); reaching it is Excel's documented behavior and returns **success +
`non_converged_cells`**. The **resource** ceiling `max_scc_iterations` (§9)
is caller-supplied; when it is **strictly lower** than the workbook's
`iterateCount` and is the bound actually hit, the run returns
**`FormulaLimitExceeded` with zero mutation** (equal bounds are the
workbook's answer — the caller permitted exactly what the file asked for,
so calling that a resource refusal would refuse a run nothing constrained) — a resource cap must never silently write caches computed
with fewer iterations than the workbook requested (that would contradict D-5,
and §9 limits are Plane-2 refusals at every layer). Which bound fired is
recorded per SCC. Fixtures: caller ceiling above / equal to / below
`iterateCount`; convergence before either bound (success); ceiling hit in one
SCC while another converges (the whole run refuses, zero mutation).
Iteration-off cycles → `FormulaCycle`. Idempotence scoped to
acyclic/converged. Fixtures with interacting cyclic + acyclic components.

**5.6d Volatile draw schedule (rng_v1)**: draws keyed by **(invocation path, stable AST callsite ordinal, SCC-pass, element ordinal)** — **exactly those four, and no component term (M5a2)**: §5.6e resets a changed SCC's pass counter and re-seeds it, so the same cell re-runs pass 1 while belonging to a *different* component, and a component term would make that a new key and a fresh draw — which is precisely the "a discovery pass cannot perturb a result" §5.6e requires. The path already names the cell and a cell is in one component at a time, so the term buys nothing and costs the property it sits beside. The invocation path is the CALLING owner plus the chain of name/table expansion **edges, each segment carrying its reference-occurrence ordinal** (so `A1=N+N` with `N=RAND()` draws twice — the two `N` references are distinct occurrences; nested repeated names and repeated table/RANDARRAY producers likewise), plus the materialized row for table producers; standalone roots use a constant path. **Volatile oracle policy (unified — supersedes any other statement)**: external oracles verify only enumerated observable properties via statistical/property protocol — repeated-reference inequality (`N+N`), per-reference re-execution, result type/range; **draw counts and sequencing are internal-KAT-only** — two `RAND()` in one cell are
distinct call sites; `RANDARRAY` elements draw by element ordinal; memoized
for the rest of the recalc; dynamic-edge rebuild and shape-stabilization
passes **reuse** memoized draws. KATs: PRNG seed/sequence (M3b);
multi-callsite `RAND()+RAND()` (M4d); draw order under lazy branches (M4d);
graph-order + rebuild-reuse (M5a); RANDARRAY element ordinals (M7a);
iterative-SCC pass keys (M5a).

**5.6e Dynamic refs — outer/inner loop integration (normative)**: the
**outer loop** is graph rebuild (runtime-edge capture → recondense → at most
`max_dynamic_passes` = 3); the **inner loop** is per-SCC iteration (§5.6c).
After a topology change: SCCs whose membership changed (merged, split, or
gained/lost edges) **reset their pass counters and re-seed from current cell
values**; unchanged SCCs keep their converged state; the re-evaluation set is
the transitive dependents of every changed node plus all changed SCCs.
Volatile draws stay memoized across rebuilds (§5.6d) — discovery passes
cannot perturb results. Exhaustion → `FormulaDynamicRefUnstable`.
Volatile-shaped spills must repeat their shape across passes else `#SPILL!`
(indeterminate). Fixtures: INDIRECT flipping a cycle open/closed, OFFSET
merging two SCCs, spill-shape change re-splitting a component.

**5.6f Standalone eval staging**: M4b3 `Workbook.evaluate` is cache-based
(referenced formula cells yield cached `<v>`; missing → blank), labeled
internal; M5a adds dependency-closure evaluation. Exposure: **CLI at M6,
C ABI + Python at M9a, Spark at M9b** — all after closure semantics exist.
Both behaviors tested; no silent switch. **Purity contract**: `evaluate`
is scratch-only — closure evaluation, spills, volatile draws, and graph
state live in the run arena and never mutate Workbook logical state or
serialized bytes; gated by logical-state + serialized-byte identity tests
across success, refusal, cancellation, and injected OOM.

**5.6g 3D references**: inclusive sheet-span expansion in workbook order;
**frozen v1 eligible list — exactly {SUM, COUNT, COUNTA, AVERAGE, MIN,
MAX}**, each oracle-tested; every other reference-consuming function
refuses 3D spans typed; **context legality**: 3D refs inside CSE/DA array
formulas or intersection (`@`/implicit) contexts refuse pre-eval and
pre-persist (oracle cases); edges per member sheet;
missing/reordered endpoints → pinned `#REF!` semantics.

**5.6h Legacy CSE placement (normative, in declared-range terms)**: let the
declared range be D (rows_D × cols_D) and the evaluated array R (rows_R ×
cols_R). **Only a 1×1 scalar R broadcasts** — it fills every cell of D.
**Every non-scalar R — including 1×N and N×1 vectors — is placed by
coordinate, never replicated**: cell (i,j) of D takes R(i,j) when i ≤ rows_R
**and** j ≤ cols_R; **`#N/A` wherever D extends beyond R in EITHER
dimension**; surplus of R beyond D is **truncated** (discarded). The round-7
per-dimension broadcast rule was wrong — Excel replicates scalars only and
pads a CSE range larger than its returned array with `#N/A`. Error R
propagates to every D cell. 2D fixtures for each combination (D>R in both
dimensions, D<R, mixed per-dimension, and **1×N into M×N / N×1 into N×M,
each asserting the `#N/A` fill rather than replication** — oracle-pinned,
because that is precisely where the superseded rule diverged).

### 5.7 Recalc pipeline

`Workbook.recalculate(alloc, io, run, opts, diag) RecalcError!RecalcReport` —
an **in-memory transaction** (§5.7.9 separates the file transaction).

1. **Model build** over the logical view (staged deltas + appends); namespace
   preflight; XML decode boundary + decoded symbol layer; dataTable refusal;
   attribute-based shared classification + topology validation; slave translation via an **AST copy-translation matrix** (independent of the byte rewriter's M10+ full-row/col deferral): (Δrow,Δcol) over relative refs incl. **full rows/columns, mixed absolutes, 3D spans, spill refs, refs in lazy branches**; off-grid → `#REF!`; fuzzed + oracled per row; CSE per
   §5.6h; dialect via M4a metadata (ambiguity → refusal).
3b. **Embedding stale-coverage policy** (numbered where it RUNS — after
   step 3's staging, before candidate serialization; ordering normative): coverage hashes bind to canonical cell payloads
   (`docs/plans/embeddings-in-xlsx.md:455-462,610-618`; drift detection
   `pkg/workbook.zig:775,1942-1960`). After evaluation and staging, compare
   the **canonical before/after payload of EVERY staged cell intersecting
   ANY coverage** — incl. created/changed/cleared spill tails (value cells;
   stale-able even under `include_formulas=false`); AST reachability is
   never the test; any hash-affecting overlap **refuses in v1** (`FormulaStaleEmbeddings`, coverages in diag) —
   `zlsxER1` has no status field and `hashes.bin` no stale sentinel
   (`recovery_record.zig:29-33,112-150`; `embeddings-in-xlsx.md:540-558`);
   `.mark_stale` requires a versioned record migration, **M10+**.
   Never a silent commit of changed `<v>` under old hashes. Tests: success,
   refusal, cancellation, injected OOM.
2. **`CT_CellFormula` attribute policy (complete inventory)**: `t`, `si`,
   `ref` — modeled (§5.6h/shared); `ca` ("Calculate Cell") — an always-recalculate/dirty **hint**, preserved;
   NOT function volatility, never touches RNG scheduling; **`aca` (always-calculate-array) — CALCULATION GRANULARITY, not
   volatility**: `true` = the legacy array calculates whole; `false` =
   Excel may calculate cells individually. Parsed + preserved; zlsx
   evaluates CSE arrays whole under both values (identical results in
   acyclic graphs — oracle for both states); a **circular CSE array with
   `aca="false"` refuses** (per-cell granularity inside a cycle is
   semantics we don't reproduce) (M4b2);
   `dt2D`/`dtr`/`del1`/`del2`/`r1`/`r2` — data table machinery →
   `FormulaDataTableUnsupported`; `bx` → **`bx="true"` refuses** (name-assignment semantics, unsupported); `bx="0"`/`false` — which Office requires when present — is accepted + preserved; **`xml:space` parsed + preserved** (Excel emits it on `<f>`; whitespace-sensitive fixtures); **unknown attributes → pre-mutation
   typed refusal** (the current parser discards attributes,
   `sheet_xml.zig:489-535` — silent normalization is exactly what we refuse
   to do). Raw spans always preserved.
   **Input cell-type contract (M4b1)** — the mapping the current parser
   lacks (unknown `t` silently becomes a number today,
   `sheet_xml.zig:540-549`): `t` absent/`n` **with a present `<v>`** → `parseDecimal(cache_import)`; absent `<v>` on a non-formula cell → `blank`; absent `<v>` on a formula cell → uncached (§5.6c seeds / §5.6f pre-closure reads);
   `s` → SST index → decoded string (ST_Xstring input decode);
   `inlineStr` → decoded string; `str` → text; **`t` never overrides the
   uncached rule (normative precedence)**: formula-cell-plus-absent-`<v>` is
   decided FIRST and is always *uncached*, whatever `t` says — a formula cell
   may legitimately carry `t="b"`/`t="e"` with no cached value, and the b/e
   lexical tables below apply **only when `<v>` is present**; `b` → boolean via a **normative lexical table** — exactly
   `0` or `1`, whitespace-free (the reader's lax rule, `xlsx.zig:1981-1985`,
   is a reader concern); other tokens, or an **empty** `<v>`, → pre-mutation
   refusal; **absent `<v>` on a NON-formula `t="b"` cell → pre-mutation
   refusal** (nothing to interpret); `e` → `ErrorValue` via its table — known-ten
   grammar or rich `#…!`/`#…?` pattern, byte-preserved; malformed or **empty**
   `<v>` → pre-mutation refusal; **absent `<v>` on a NON-formula `t="e"` cell
   → pre-mutation refusal**; oracle rows are authored separately for the
   formula and non-formula cases of both `t="b"` and `t="e"`; `d` → serial via a **normative lexical table (M4b2, with `date1904`)**: accepted forms exactly `YYYY-MM-DD` and `YYYY-MM-DDTHH:MM:SS[.fff]` (≤3 fractional digits); **timezone offsets refuse** (Office supports a limited ISO-8601 subset, not the full grammar); range-checked against the active epoch; invalid text → pre-mutation typed refusal; Excel/LO fixtures per row; **unknown `t` →
   pre-mutation typed refusal**, never a silent number.
3. **Evaluate** (§5.6). **Stage into `ResolvedSheet`** — raw spans + (raw
   `<f>` + attrs) + new cached results + setCell replacements + appends;
   deltas consumed once. Transitions: number → `t` removed, shortest-rt;
   text → `t="str"` + ST_Xstring-encoded `<v>`; boolean → `t="b"` 0/1;
   error → `t="e"` literal (rich preserved); **blank publication → numeric 0**
   (`=A1` on empty A1 caches `<v>0</v>`; spill gaps publish 0; fixtures:
   `=A1`, omitted arg, spill-with-gaps, ISBLANK, `""`; internal blank exists pre-publication only — standalone eval publishes through the SAME blank→0 conversion, §12.2); `""` →
   `t="str"` + empty `<v></v>`. Slaves keep `<f>` any-shape, gain `<v>`.
   Pre-M7 gate: non-1×1 / DA anchor / CSE → `FormulaSpillPersistUnsupported`,
   zero mutation.
4. **Prepare/swap — ONE candidate, swap LAST**: every mutation — formula
   caches, spill mutations, **calcChain removal and calc-state changes
   (steps 5–6 describe staging components, not post-swap actions)**,
   report, diagnostics — stages into a single candidate; `recalculate()`
   swaps as the final pipeline operation; `saveWithRecalc` orders
   serialize-candidate → rename → swap (§5.7.9). No workbook/package bytes or candidate allocations
   mutate after the swap (sole exemption: the caller-owned report's
   preallocated durability slot, §5.7.9). Over complete PartStore state
   (parts, overrides, rels,
   `rels_by_owner`, typed views) **plus fully-constructed report +
   diagnostics** — all allocation in prepare. **Borrow-lifetime preservation
   (normative)**: `Workbook.sheet()` returns stable pointers
   (`workbook.zig:1119-1127`) and `cellByRef()` promises Workbook-lifetime
   validity (`:6484-6492`), while view replacement currently frees old
   arenas (`:4757-4761`) — the swap keeps worksheet-slot addresses stable **for nonstructural recalc
   only** (structural add/delete keeps its documented invalidation,
   `workbook.zig:3574`) and **retains the ENTIRE superseded generation — old PartStore (source handle + part arena)
   AND typed views — until `Workbook.deinit`** (sheet leaf strings borrow
   part bytes, `sheet_xml.zig:3-9`, not view arenas — retaining views alone
   would dangle); retention is **counted, not documentary**: `max_retained_generations`
   (default 4) plus retained-byte accounting — with `SourceBacking` (M5b0) all generations SHARE one ref-counted source handle (fd budget = one total); reclaiming requires
   **`Workbook.deinit` and reopen** (borrows are Workbook-lifetime,
   `workbook.zig:6484-6492`); repeated-recalc tests gate RSS, allocation accounting,
   borrow validity, and fd count; hold-across-REPEATED-recalc borrow tests across success,
   refusal, cancellation, and injected OOM; swap and everything after
   no-fail (failing-allocator-proven; post-failure reads; bounded report
   storage + truncation flags; the report carries a **preallocated dormant
   durability slot** — a fixed flag + errno field flipped without
   allocation if post-rename dir-fsync fails, since that outcome is
   discovered only after all allocation must be done). Cancellation polled
   during eval and before swap; cancelled runs leave memory untouched.
5. **calcChain removal**: rel target resolved relative to `xl/workbook.xml`
   (owner-relative norm, corpus-verified); part + rel + content-type removed
   atomically; relative/absolute/noncanonical tests; Excel-opens-clean.
6. **Calc-state policy** (complete; unknown attrs byte-preserved;
   round-tripped): `calcId` — see the truthful-producer state below (byte-level
   assertion: every successful recalc emits `calcId="0"` +
   `fullCalcOnLoad="1"`) · `calcMode` preserved
   (`markRecalcOnLoad` per-mode pinned) · iteration fields honored+preserved
   (clamps §5.6c) · `fullPrecision="0"` → **refusal through v1** ·
   **truthful producer state (supersedes both prior policies)**: after a successful zlsx recalc, `calcId` is set to `"0"` (unknown/older producer — preserving Excel's ID would claim Excel produced zlsx's caches, and Excel trusts calcId/calcFeatures when deciding whether to recalc) and `fullCalcOnLoad="1"` is SET (Excel re-verifies on open, harmlessly; non-Excel consumers — the commercial target — read the fresh `<v>` either way); `calcMode` + iteration settings preserved; `calcCompleted` per pinned oracle state; relaxation of these conservative signals = M10+ after Excel-open oracles across calcId/calcFeatures combos ·
   `calcCompleted` transition pinned · `calcOnSave`/`forceFullCalc`/`refMode`
   parsed+preserved · `concurrentCalc(ManualCount)` preserved · worksheet
   `sheetCalcPr@fullCalcOnLoad` **preserved in v1** (same provenance policy); `calcFeatures` parsed + preserved (M4b2) ·
   CV extension parsed+preserved · absent elements created only when needed.
7. **Refusals**: whole-recalc default + census; `.keep_stale_and_mark`
   (byte-identical except `fullCalcOnLoad="1"`) — **eligibility
   (normative): mark-only may suppress ONLY `FormulaUnsupportedFunction`/
   `FormulaUnsupportedConstruct`; `FormulaMalformedInput`, signature,
   embedding, limit, cancellation, and I/O failures ALWAYS refuse**;
   `markRecalcOnLoad()` honestly named. Purity baseline: staged-state save without recalc ≡ save
   after refused recalc.
8. **Report**: prepared pre-swap; counts, passes, `non_converged_cells`,
   dynamic passes, census, resolved-RunInputs echo.
9. **`saveWithRecalc` — the file transaction (ordering normative)**: run
   prepare fully → serialize output bytes **from the prepared (unswapped)
   state** → write temp → **`File.sync(io)`** (file contents durable) →
   **final cancellation poll** → atomic rename (**the commit point**) →
   swap in-memory state (no-fail, immediately after rename) → directory
   fsync on POSIX **after the swap** — a dir-fsync failure at that point is
   a **committed-with-durability-warning diagnostic** (the rename succeeded;
   memory and file are consistent; status stays success with a
   `durability_warning` flag), never an error that could contradict the
   already-committed state. **The rename-through-swap span is
   non-cancellable**; the last poll sits immediately before rename, so
   cancellation can never commit a file; the CLI's exit mapping is
   commit-aware — a signal arriving after rename reports success, not
   130/143. **Existing/overwritten destinations (incl. Python's permitted
   save-over-source, `__init__.py:2994`)**: the temp file lives in the
   destination's directory and rename replaces atomically, so the promise
   is precisely "**destination bytes are unchanged until the commit
   point**" — pre-commit failure leaves the prior destination intact;
   "no output file exists" applies only to initially-absent destinations.
   `pkg/atomic_file.zig:119-127` currently flushes without syncing:
   **`AtomicFile.finish` gains `File.sync(io)` before rename** (M5d scope);
   injected-failure tests around sync, rename, post-rename dir-fsync, and
   the state commit. Any failure before rename ⇒ memory untouched AND
   destination bytes unchanged (or destination still absent if it never
   existed) — tested at every injected failure point for BOTH cases;
   rename success ⇒ memory and file consistent.
   `recalculate()` alone swaps at step 4 with no file involved. Both
   allocation-failure-swept.

### 5.8 Dynamic arrays & spill

**5.8a Decision table** (fixture-pinned): fits+owned/empty → spill; foreign
non-empty → `#SPILL!`(obstruction); sheet edge → (edge); table → (table);
merge → (merge); volatile-shape unstable → (indeterminate); >
`max_matrix_cells` → `FormulaLimitExceeded`. Ownership: spilled cells record
their anchor; shrink clears own tails; competing anchors resolve in calc
order; tail mutations ride the transaction. Tests: shrink/grow/anchor-error/
overwritten-tail/racing.

**5.8b Persistence — approved mutation set**: M7b1 may mutate `<v>`+`t`,
anchor `f@ref`, owned tail create/clear, **worksheet `<dimension ref>`
expansion when a spill extends the used range** (dimension maintenance is
already a recognized distinct mutation, `docs/plans/structural-edits.md:100`;
byte-diff gate covers it; spills that would extend the dimension refuse
until this mutation is proven), and enumerated `cm`/`vm`
transitions — **each row names the exact metadata collection
(cellMetadata vs valueMetadata), record type, one-based index, and
missing-record behavior, pinned empirically from byte-diffed Excel-authored
references** (spec and Office behavior differ; we pin what Office writes).
Every transition carries an Excel-opens-clean proof or refuses. Others →
`FormulaSpillPersistUnsupported` until M7c (authoring; part-graph spec from
byte-diffed references). `#SPILL!` cached value vs rich metadata split
maintained.

**5.8c Authoring dialect**: `FormulaWrite{ text, dialect: .scalar (default) |
.dynamic_array | .cse(ref) }` — **Zig-only at M7c**; the versioned C export
(`…_v2` with dialect; existing `zlsx_sheet_writer_write_row_with_formulas`
untouched, `c_abi.zig:1837-1920`) + Python land at M9a2 with probes. v1
authors `.scalar`; M7b1 recalculates/persists **existing** CSE only; **both `.cse` and `.dynamic_array` authoring land at M7c**.

**M4a decisions (shipped 2026-08-03).** Seventeen points the row left
open, got wrong, or discovered. They span §5.3b (dialect), §5.6a
(`dialectOf`), §5.8b (`cm`/`vm`) and §10 (refusals), and are collected
here because §5.8b is where the `cm`/`vm` contract lives.

1. **Everything in the part is spec-pinned, and the sources are named.**
   No committed manifest in `tests/oracle/fixtures/` contains a metadata
   row and no workbook in `tests/corpus/` carries an `xl/metadata.xml`
   part, so nothing here is oracle-decided. The element and attribute
   inventory comes from ECMA-376 `sml-sheetMetadata.xsd`, `c@cm`/`c@vm`
   from `CT_Cell` (`xsd:unsignedInt`, default `0`), one-based indexing
   from `:133` above, and the type names from what Office writes. M7b
   still re-pins the *transitions* from byte-diffed references; nothing
   at M4a writes, so nothing at M4a front-runs that.
2. **The classification table is the parser, not documentation.**
   `element_inventory` has one row per schema element — 20 — and
   `childOf(parent, name)` is a loop over it. An element with no row has
   no legal position anywhere, so the parse refuses instead of skipping.
   A table that only *described* the reader would drift from it in the
   first commit that added an element.
3. **Two tables, because the part has two kinds of "type".** Elements are
   schema-defined and closed; metadata *type names* are producer-defined
   strings and open. That is why `unknown` is a first-class row of
   `type_classification` with a treatment rather than an error case —
   there is no schema to consult about a name Excel invents next year.
4. **Names that the schema overloads by position get distinct members.**
   `bk` under `cellMetadata` and `bk` under `futureMetadata` are
   different records; so are `t` under `mdx` and the `t` attribute of
   `rc`. A reader that resolved element names globally would classify
   whichever it saw first.
5. **Only `XLDAPR` is interpreted, and the refusal is a property of the
   reference — not of the file.** A workbook full of rich values
   recalculates cleanly as long as the run neither reads nor writes a
   cell that points at one. That is what makes `CellRole` a parameter of
   resolution rather than a fact about the workbook: the input side and
   the result side ask the same question and get the same answer, and the
   role only decides which side the diagnostic names.
6. **All value metadata refuses, which is deliberately wider than the
   row's `XLRICHVALUE`/unknown.** Office declares `XLDAPR` with
   `cellMeta="1"` — it is a *cell* metadata type. Reached through `vm` it
   means something this reader has never been shown, and the safe reading
   of "never been shown" is a refusal (`dynamic_array_in_value_metadata`),
   not a dialect.
7. **Value metadata is checked before cell metadata.** A cell carrying
   both a dynamic-array mark and a rich value must refuse; checking `cm`
   first would hand back a dialect a caller could act on before noticing
   the part it must not touch.
8. **Pre-mutation is a property of the API, not of the caller's
   discipline.** `resolveRun` is two-phase: phase 1 classifies every cell
   and writes nothing, phase 2 fills the output only once no refusal is
   possible. A single-pass loop would leave every cell before the
   rich-value one already resolved. Fixtured with a poisoned output
   buffer, on the input side and the result side, and asserted by the
   fuzz target on every input it survives parsing.
9. **Parse-time and resolution-time refusals are different questions.**
   The part being unreadable — bad XML, an unclassifiable element, an
   attribute the schema does not define — refuses at parse. A rich value
   refuses at the reference. Collapsing the two would make an untouched
   rich value refuse a whole workbook, which is the over-refusal that
   makes people disable the check.
10. **One-based, and the `0` default is the argument.** `cm="0"` means
    "no metadata"; a zero-based index could not distinguish that from
    "the first block". The base schema's prose and Office's behaviour
    differ here (`:133`), and this reader follows Office.
11. **`count` attributes are recorded and not enforced.** Resolution
    indexes by position. `tests/corpus/poi_MalformedSSTCount.xlsx` is
    this repo's standing evidence that producers miscount, and refusing a
    file Excel opens over a hint is a worse failure than ignoring it —
    the hint is exposed for diagnostics instead.
12. **An unknown attribute on an interpreted element refuses**, so
    `CT_MetadataType`'s 28 attributes ship as data. The booleans are
    parsed and validated, then discarded except `cellMeta`: a
    "validated then dropped" attribute is classified, an unread one is
    not.
13. **`<!DOCTYPE>` refuses outright.** An entity-expanding metadata
    reader is a denial-of-service surface with no upside;
    `tests/corpus/poi_xxe_in_schema.xlsx` is why this repo takes that
    seriously. Foreign content under `<ext>` is skipped wholesale by
    depth, bounded by `max_skip_depth`.
14. **Every inert element's payload is reachable only through a type that
    refuses.** That is the invariant that makes "recognized and skipped"
    safe for `metadataStrings`, `mdxMetadata` and `futureMetadata`: no
    cell can consume them without first meeting `XLMDX`, `XLRICHVALUE`
    or an unknown name. The interpreted set is exactly seven elements and
    a test names them.
15. **Type names match exactly, case included.** A differently-cased or
    entity-escaped spelling classifies as `unknown` and therefore
    refuses. The alternative is inferring "this is probably `XLDAPR`"
    about a part we are about to overwrite values under.
16. **`env.zig` stays a leaf: the dependency points the other way.**
    `EvalEnv` gains `DialectResolver` — a context pointer and a function
    pointer — and `metadata.zig` binds itself to it
    (`CellDialectResolver`). The `Fake` gains `cm`/`vm` per cell and
    answers `dialectOf` through the resolver when one is attached, so
    every test written before the reader existed still means what it
    said.
17. **The env-level error collapses onto the class that always
    refuses.** `EvalEnv` can only say `error.MetadataRefused`; the
    precise plane — `FormulaUnsupportedConstruct` for a rich value,
    `FormulaMalformedInput` for a broken part — travels with the
    resolver's own `Refusal`, which the report carries. Where the detail
    is already gone (`eval.planeTwo`) the mapping is
    `FormulaMalformedInput`, because collapsing onto a mark-eligible
    class could let `.keep_stale_and_mark` suppress a refusal §5.7.7 says
    must stand. §10's taxonomy keeps one home: `metadata.zig` imports
    `parser.PlaneTwo` rather than restating it.

**M4b1 decisions (shipped 2026-08-03).** Nineteen points the row left
open, got wrong, or discovered. They span §5.6a (`EvalEnv`), §5.7.1–2
(decode boundary, input cell-type contract), §5.9 (symbols) and §9
(limits), and the corpus decided four of them against what the row had
written down.

1. **The corpus decided the row axis, and it decided it against the
   plan.** The row scoped implicit coordinates to `<c>` and said a
   `<row>` without `r` should refuse. `tests/corpus/wdi_excel.xlsx` —
   the World Bank's WDI export, 13.9M cells — omits `r` on **both**
   `<row>` and `<c>` throughout, and `src/xlsx.zig:1858-1864` already
   reconstructs it. Refusing would have refused a workbook this repo
   reads today. The rule is now the same on both axes: the row after the
   predecessor, first row 1.
2. **A row's number is provisional until its first cell.** A producer
   that omits `r` on the row may still put it on the cells, and then the
   cells are the authority — which is exactly what the reader's
   `recoverRowFromFirstCell` does. Deferring means both readings agree on
   the same file instead of silently placing values two rows apart.
3. **`<row r="0">` refuses where the reader skips it.**
   `tests/corpus/poi_poc_shared_strings.xlsx` carries one. A reader may
   drop a row it cannot place; an engine that writes values back cannot,
   because a dropped row recalculates as blank cells.
4. **Three carrier classes, not two.** FORMULA and STRING are the row's
   named pair, but numeric/boolean/error `<v>` bodies needed a third
   (`lexical`): they decode like formulas (entities only) and *author*
   like neither — this engine generates their spelling, so there is
   nothing to escape and nothing that can fail to encode.
5. **Authoring a formula can fail, and that is the asymmetry's price.**
   With no ST_Xstring stage, a formula containing a character XML cannot
   represent has nowhere to go: `unencodable_formula_char`, not an
   invented escape. Emitting `_x0001_` there would produce a formula that
   reads back as seven literal characters.
6. **Identifiers are STRING carriers; bodies are FORMULA carriers.**
   `CT_Sheet@name`, `CT_DefinedName@name`, `CT_Table@name` and
   `CT_TableColumn@name` are ST_Xstring-typed and decode both passes; the
   `<definedName>` body, `<f>`, `<calculatedColumnFormula>` and
   `<totalsRowFormula>` are the exception the row already named. One
   fixture puts the same seven bytes in a name and in its own body and
   pins the two answers side by side.
7. **The preflight is a sweep, not a check.** Every part the run will
   read — workbook, shared strings, every sheet, every table — is
   namespace-checked before *any* of them is decoded, so a bad
   vocabulary on sheet 9 cannot be discovered after sheet 1 has been
   modeled. ISO Strict is refused as **classified** rather than unknown:
   the vocabulary is nearly the same, and the parts it differs in are
   exactly the parts recalc reads.
8. **Foreign-namespace attributes are exempt from the inventory;
   foreign-namespace elements inside `<sheetData>` are not.**
   `x14ac:dyDescent` is on nearly every row Excel writes, so refusing a
   namespaced extension would refuse the ordinary case. An unrecognized
   element inside `<sheetData>` could be wrapping cells, so skipping it
   would drop them: that refuses. Everything outside `<sheetData>` is
   skipped wholesale without inspection.
9. **Two `<c>` at one coordinate refuse.** Last-wins and first-wins are
   both defensible, which is why neither is chosen silently.
10. **`max_modeled_cells` is §9's name and §9's default (64M) — and it
    does not bound memory.** Modeling `wdi_excel.xlsx` costs ~7 GB at
    13.9M cells, well inside the limit. The cell count is a shape bound;
    the **byte** budget (`max_run_arena_bytes`, §9 aggregates, counted
    allocator) is what actually protects the process, and wiring it
    through the model builder is M5's. Recorded here because the gap is
    real and the number in §9 does not close it. The adapter charges the
    limit workbook-wide (each sheet scans against what the sheets before
    it left), not per sheet.
11. **The engine is ONE module, and it had to become one.**
    `pkg/workbook.zig` reaching `env`, `value` and `decode` as separate
    modules would compile `env.zig` twice and build an `EvalEnv` the
    evaluator could not accept. Worse, `src/xlsx.zig` imported
    `formula/rewriter.zig` by relative path, which would have put
    `tokenizer.zig` in both `zlsx` and the engine module — "file exists
    in modules" is a build error, not a subtlety. `xlsx.zig` now reaches
    the rewriter through the named module, and the engine is named once.
12. **The pkg-facing module declares no `zlsx_casefold`.** That
    dependency exists only for the TEST sections of `value.zig` and
    `symbols.zig`; declaring it in a compilation that also contains
    `zlsx` claims `src/unicode/casefold.zig` for two modules. Two module
    objects, one root file, one of them test-only. (M4f moved the file
    out of `src/`, so the claim is no longer contested — but the
    pkg-facing module DOES now declare `zlsx_casing`, which cannot be
    test-confined because `UPPER` and `LOWER` compute with it.)
13. **A refusal is a normal return, so `errdefer` does not fire.** Every
    scan entry point releases its arena through `defer if (!keep)`, and
    the fuzz target is what proves a refused parse frees what it
    allocated.
14. **`Builder.deinit` is idempotent.** An ownership rule a caller has to
    remember is one a caller gets wrong under injected OOM — the
    `consumed` flag makes the obvious `defer b.deinit()` correct on every
    path, including the one where the allocator fails between the last
    `add` and `finish`.
15. **`Workbook.open` double-freed its `PartStore` on every failing
    open**, found by this row's allocation-failure sweep:
    `fromStore` takes ownership *including on failure*, and both it and
    `open` had an armed `errdefer`. A workbook with no
    `xl/workbook.xml` segfaulted instead of returning
    `MissingWorkbookPart`. Fixed in `open` and `empty` by disarming
    before the hand-off.
16. **`dialectOf` routes through M4a, and a dangling `cm` refuses.** A
    `cm="1"` on a workbook with no `xl/metadata.xml` is
    `index_out_of_range`, not a guessed dialect; the precise refusal is
    retrievable from the adapter (`lastDialectRefusal`) because the env
    interface can only say `error.MetadataRefused` (M4a decision 17).
17. **`spillShape` answers null through this row.** A spill's extent is
    not recoverable from cached `<v>` values, and inventing one would
    make `A1#` resolve to a guess. §5.8's stored spill state arrives at
    M7a.
18. **The adapter is proven by the fake's own suite.** Ordering,
    precedence, both blank classes, the N-way cursor's modes and its
    run/position accounting, unknown-sheet, sheet resolution — one test
    body, two harnesses, and the adapter's harness authors real sheet XML
    and reads it back through the decode boundary rather than injecting
    cells directly.
19. **`PartStore.addPart` does not refresh the rels cache;
    `replacePart` does** (`store.zig:688`). Only a test-construction
    consequence — real packages come from `open`, which parses every rels
    part — but it is why the in-memory workbook builder seeds an empty
    rels part and then replaces it, the same two-step `Workbook.empty` +
    `addSheet` takes.

**M4b2 decisions (shipped 2026-08-04).** Fourteen points, in
`src/formula/calc.zig` (the attribute inventory, the topology, the
translation matrix, the calc state), `serial_date.zig` (the `t="d"`
table), `decode.zig` (the epoch plumbing) and the adapter. The corpus
decided five of them.

1. **One corpus workbook decided the whole classification rule.**
   `tests/corpus/calamine_non_monotonic_si.xlsx` writes its shared
   slaves as `<f si="1" t="shared"></f>` — an open/close pair with an
   empty body, not the self-closing form `xlsx.zig:2099` recognizes — so
   a shape-based reader loses nine of that workbook's twelve shared
   cells and recalculates them as blanks. The rule is the schema's:
   **`t="shared"` with `ref` is a master, without `ref` is a slave**,
   whatever shape the element has. The same file writes attributes in
   the order `ref si t`, so order-independence is corpus-decided too.
2. **`si` is a key, not an ordinal, and `ref` is not a topology gate.**
   Same workbook: masters appear in the order 1, 0, 2, and
   `ref="A3:A7"` names two rows that have no cell at all. So nothing
   assumes `si` counts up, and a slave is never checked against its
   master's declared range. What *is* checked is order — a master must
   precede its slaves sheet-wide, because a slave is defined by
   translating a formula it has not been given yet.
3. **Two failure modes, two passes.** A slave whose master appears
   *later* and a slave whose `si` names no master are different
   statements about a file, and a single pass could only report the
   second. Masters are collected first; slaves are classified against
   the finished set.
4. **The master is the single source of truth for a slave's formula.**
   ECMA-376 requires a slave's body, when written, to *be* the
   translated formula, so the two agree by construction and either would
   be conformant. The master wins because that is one parse per group
   rather than one per cell — `Masters` exists for exactly that — and
   the slave's body is preserved and read by nothing.
5. **Translation is an AST copy, and a reference operand collapses
   whole.** `rewriter.zig` shifts text and defers full rows and columns
   to M10+; a slave cannot defer, so translation copies the tree, shifts
   every *relative* half by (Δrow, Δcol), and prints. An endpoint that
   leaves the grid takes its entire operand to `#REF!` — never
   `#REF!:B2`, which is a spelling the format has no token for
   (MS-XLS `PtgRefErr`/`PtgAreaErr`). A qualifier survives its target:
   `Sheet1!#REF!`, which is what Excel writes.
6. **Collapse is a consequence of *this* translation, tracked as one.**
   A `#REF!` the file already carried is left where it is, so
   `translate(ast, 0)` is provably the identity — asserted on every
   fixture in the matrix, byte for byte. Without the distinction, a
   formula nothing moved would come back rewritten.
7. **Children always have lower indices than their parents**, so one
   ascending pass over the node array *is* a post-order walk and the
   collapse rule sees its operands already decided. Asserted rather than
   assumed (`assertChildrenBelow`).
8. **The differential test needed a genuinely independent translator.**
   The reference implementation scans the formula *text* for A1
   references where `translate` walks a tree; they share nothing but
   `coords`. It is deliberately naive, so the generator emits only
   formulas it can read and only deltas that stay on the grid — every
   construct it cannot handle (quoted sheet names, spills, structured
   references, the off-grid collapse) has a fixture in the matrix
   instead. 2 000 randomized cases, and the identity case doubles as a
   print round-trip.
9. **`<calcPr>` preserves unknown attributes where `<f>` refuses
   them.** The opposite policy, deliberately: an unread `<f>` attribute
   can change what a cell's formula *is*, and an unread `<calcPr>`
   attribute cannot change what any cell contains. Refusing the latter
   would refuse files that open everywhere. A *known* attribute with an
   unknown value still refuses — an unclassified `calcMode` decides
   whether the workbook recalculates.
10. **The round-trip is reconstructed, never echoed.** `writeCalcPr`
    rebuilds the element from the parsed attribute region and the
    open/close form, so the corpus test proves the parse *kept* what it
    needs rather than proving `memcpy` works. The two shapes that catch
    a naive writer are both in the corpus: `<calcPr/>` with no
    attributes at all, and `wdi_excel.xlsx`'s `<calcPr calcId="40001" />`
    with a space before the `/>`.
11. **The `t="d"` table lives in `serial_date.zig`, not with the calc
    state.** It needs the epoch, which is calc state, but putting it in
    `calc.zig` would have made `decode.zig` import `calc.zig` while
    `calc.zig` imports `decode.zig`. `serial_date` already owns the
    calendar including both invented 1900 days, and its
    `DateOutOfRange` *is* the range check §5.7.2 asks for — a second one
    would have been a second answer. `decode.Options` gains
    `date_system`, and the adapter reads `<workbookPr date1904>` **before
    the first sheet is scanned**, which is the only ordering under which
    the same eight bytes cannot mean two serials.
12. **`t="d"` with no `<v>` joins `b` and `e`.** A `t` that promises a
    value the cell does not have is `missing_cached_value`, and the
    uncached rule still wins first: a formula cell with no `<v>` is
    uncached whatever `t` says. `date_cell_unsupported` is gone; its
    successor `bad_date_cache` is a malformed-input refusal, not an
    unsupported-construct one, because the construct is now supported.
13. **CV2's marker is left unpinned on purpose.** The calcFeatures
    extension is recognized by URI, its feature names are collected, and
    every `<ext>` is preserved byte-exact — but which feature name
    carries §5.4d's compatibility version is pinned the way §4 pins
    every Office-vs-schema difference: from a byte-diffed Excel
    reference (M7b). A name absent from the table leaves `text_compat`
    at CV1, §5.4d's documented default, so nothing is lost and nothing
    is invented; `textCompatOf` takes the table as a parameter so the
    mapping is tested today against a synthetic row.
14. **`classifySheet` leaked its first list when the second failed**,
    found by this row's allocation-failure sweep: two `toOwnedSlice`
    calls inside one struct literal have no `errdefer` between them, and
    the first list is already owned by nothing by the time the second
    fails. Ownership now moves one list at a time. The same shape is why
    every refusal path in this file releases explicitly — a refusal is a
    normal return, so `errdefer` never fires (M4b1 decision 13).

### 5.9 Name & identifier resolution

Call position → strip layered prefixes → registry; unregistered →
`FormulaUnsupportedFunction`. Value position → sheet-scoped name (shadowing)
→ workbook name → table → `_xlnm.` builtins → `#NAME?` (provable there). All
matching over the decoded symbol layer, case-folded. **`CT_DefinedName`
attribute inventory (M4b3, complete — 16 rows)**: refusal-when-referenced —
`function`, `vbProcedure`, `xlm`; inert-preserved — `functionGroupId`,
`publishToServer`, `workbookParameter`, `comment`, `customMenu`,
`description`, `help`, `shortcutKey`, `statusBar`, `xml:space`; **unknown
attributes → pre-mutation typed refusal** (typed view keeps
name/formula/scope/hidden, `workbook_xml.zig:58-68`); explicit M4b3
deliverable.
Name bodies = graph
nodes (M5a; depth-guard interim M4b3); opaque payloads inert unless
referenced; relative-ref names → typed refusal v1; `_xlpm.`/LAMBDA/LET refused.

**M4b3 decisions (shipped 2026-08-04).** Nineteen points, in
`src/formula/names.zig` (the two inventories, §5.9's drivers, the 3D
matrix), `symbols.zig` (the tiers and the resolver seam), `eval.zig`
(expansion and 3D evaluation), `env.zig` (the seam), `decode.zig` (table
geometry) and the adapter. The corpus decided three of them and
**refused to decide the rest** — which is itself the row's most
important finding.

1. **The order is data, and the driver reads it.** `resolveInOrder`
   takes §5.9's sequence as a *parameter* and iterates it;
   `symbols.SymbolTable.Tiers` supplies one tier per entry and states no
   sequence at all. The proof is a test that hands the driver a permuted
   array and watches the winner change — a chain of `if`s could not do
   that. M2 exported `value_resolution_order` and
   `call_resolution_order` for exactly this, and a comptime check now
   proves each is a permutation of its enum: a stage missing from the
   array is a stage that never runs, a stage listed twice runs twice.
2. **The `CT_DefinedName` inventory is sixteen rows, and the corpus
   knows only three of them.** Across `tests/corpus/`, defined names
   carry `name` (19), `localSheetId` (17) and `hidden` (10) and nothing
   else. Every other row — including all three macro flags — is
   **spec-pinned to ECMA-376 §18.2.5 and labelled as such** in the
   table itself, rather than claimed as corpus- or oracle-derived.
3. **The macro three refuse when *referenced*, not when read.** A
   workbook may carry a `vbProcedure` name no formula mentions, and
   refusing to open it would refuse files Excel opens. Same shape
   `_xlpm.` and unknown `_xlnm.` already had, which is why
   `macro_defined_name` joins them in `decode.Refusal.Reason` instead of
   starting a second vocabulary for one case.
4. **`function="0"` is an ordinary name.** The three flags are
   `xsd:boolean`, so their presence is not the fact — their *value* is.
   A producer that writes them explicitly false means nothing unusual,
   and refusing it would refuse a file that says so.
5. **The scan reads the part, not the typed view.** `workbook_xml.zig`
   keeps four fields and drops the attribute region, which is precisely
   the region the inventory is about. `names.scanDefinedNames` walks
   `xl/workbook.xml` itself — the same choice M4b2 made for `<calcPr>`,
   for the same reason.
6. **Depth is what distinguishes a defined name from an element that
   shares its tag.** A `<definedName>` inside somebody's `<extLst>` is
   not a defined name, and a scanner matching on the tag alone would
   invent names out of a foreign vocabulary.
7. **A relative body refuses when referenced, and the body is asked
   last.** Three things can disqualify a name that resolved — its
   spelling (M4b1), its attributes (this row), and its body — and they
   are checked in that order because only the third costs a parse.
8. **A body that will not parse is the expansion's refusal, not the
   name's.** `bodyRefusal` answers null for it; the parse that expands
   it names the construct, with the construct's own offset. Flattening
   both into "this name is bad" would lose which.
9. **LAMBDA and LET refuse through the registry, not a second rule.**
   They are unregistered calls, so §5.9's call-position order already
   answers `FormulaUnsupportedFunction`. A name-level rule would have
   been one refusal stated twice.
10. **The `_xlnm.` tier answers only for a spelling nothing declares.**
    A declared builtin is found at the name tier, before the builtin
    tier runs. A *recognized* builtin the workbook never defined names
    nothing, so the tier declines and `#NAME?` stands; an unrecognized
    spelling under the reserved prefix refuses, because treating it as
    a user name that happens to start with `_xlnm.` is a guess.
11. **Name expansion is inline and depth-guarded, and says so.**
    `max_name_expansion_depth = 8` is interim: M5a makes bodies graph
    nodes, where `A = B`, `B = A` is a cycle like any other. Until then
    inline expansion has to stop somewhere, and it stops with a §9 limit
    that has a name rather than a stack overflow that does not.
12. **A table producer's members must exist.** Excel materializes a
    calculated column into every data cell, so the table part and the
    sheet state it twice; when they disagree — a producer declared, a
    member with no `<f>` — an engine that writes values back would
    recalculate that member as a blank and delete a column. Neither side
    is authoritative enough to pick, so it refuses. `CT_TableFormula`
    itself is two attributes (`array`, plus `xml:space`), and
    `decode.Table` gained `headerRowCount`/`totalsRowCount` because the
    member set cannot be computed without them.
13. **The fixture that had to change was wrong about Excel.** M4b1's
    table test declared a calculated column over `B2:B4` on a sheet
    holding only `A1` — a table Excel never wrote. It now carries its
    members, which is what made the refusal visible as a fixture rather
    than as a surprise.
14. **§5.6g ships entirely spec-pinned, and the row says so rather than
    implying otherwise.** No corpus workbook contains a 3D span; no
    committed oracle manifest contains one either (§8.2). So the matrix
    is pinned to the spec and labelled at its table, the per-function
    oracle legs arrive with the functions (AVERAGE, MIN and MAX are
    M4e), and this row's fixtures are the three eligible functions the
    registry already holds plus the matrix itself.
15. **Eligibility is judged at the enclosing call argument, and the
    three reference operators are transparent.** `:`, ` ` and `,`
    *compose* references, so `SUM(S1:S3!A1:B1)` is still SUM's argument;
    every other operator consumes a value, so `SUM(S1:S3!A1*2)` hands
    `*` a multi-area reference and refuses. Conservative on purpose:
    refusing is the direction an engine that writes values back can
    afford to be wrong in.
16. **A span and a union are the same shape and not the same thing.**
    `Reference.three_d` records which. After a union, `A1:B1` takes the
    bounding box of two areas; after a span it takes one box per member
    sheet, because a span is one reference repeated across sheets.
    Nothing else in the evaluator branches on the flag — aggregating
    over N areas is aggregating over N areas, which is why the six
    eligible functions needed no 3D-specific code.
17. **Reordered endpoints are not normalized.** `Sheet2:Sheet1` is a
    span someone's edit broke; computing over the sheets between them
    would answer a question nobody asked. It joins the missing endpoint
    at `#REF!`, which is a *value* — the proof the span was evaluated
    rather than refused.
18. **`UnsupportedConstruct` and `NotYetImplemented` are two errors
    because they answer two questions.** "Never in v1" is not "not at
    this row", and only the second has an enumerated membership. M4b3
    deleted the list's one entry (3D sheet spans) and a test now fails
    if a span refuses that way again — the deletion being watched is
    what decision 12 of §5.3 was for.
19. **`evaluate` is pure by construction, and gated anyway.** It never
    calls `putComputed`; every byte it produces lives in the arena it
    hands back. The gate compares every part's bytes before and after on
    four paths — success, evaluator refusal, a run cut short by a §9
    limit, and the injected-OOM sweep — because "it does not mutate" is
    a claim about code that changes, not about code as written. The
    resolver seam it wires (`env.NameResolver`, `error.NameRefused`)
    mirrors M4a's `DialectResolver`/`MetadataRefused` exactly: the
    interface says the resolution failed, and the typed reason stays
    with the resolver that owns the diagnostic.

### 5.10 Cycle-free composition (`zlsx_recalc`)

Third public module (importing `zlsx` + `zlsx_pkg`; no cycle;
`tests/consumer` dependency test): `writerSaveWithRecalc` = Writer
**`saveToOwnedBuffer(allocator, io)`** (new, gated — allocator-first per `AGENTS.md:194`; ownership documented,
size-limited, allocation-failure-swept, byte-equivalent to path save) →
`Workbook.openBuffer(allocator, io, bytes)` (borrow ends at return; store
copies — Book precedent `xlsx.zig:795-811`) → recalculate → save. CLI, C ABI
(`zlsx_editor_open_buffer`, `zlsx_writer_save_with_recalc`), Python compose
through it.

**Cancellation reaches BOTH pre-recalc stages (normative)**: fresh
serialization and buffer open can each process gigabytes *before* recalc
begins, so an uncontrolled Writer path would violate the §5.5 polling bound
that M5d1 promises. Each therefore gains a **control-aware variant** —
`Writer.saveToOwnedBufferControlled(alloc, io, ctl)` and
`Workbook.openBufferControlled(alloc, io, bytes, ctl)`, where
`ctl: Control = { cancel: ?CancelToken, deadline: ?std.Io.Timestamp }`; the
plain signatures remain and forward a null control, so §12.1 stays stable.
`writerSaveWithRecalc` threads the orchestrator's own cancel token and
deadline through **both**. They poll on the same M5d1 chunked seams
(compression, decompression, raw-entry copies, XML scans), return
`error.Cancelled` with the partial buffer freed, and mutate nothing — both
stages sit entirely before the commit point (§5.7.9). Tests: cancel
mid-fresh-serialization, cancel mid-buffer-open, deadline expiry in each, and
the polling-latency bound measured on a deliberately large fresh workbook.

---

## 6. Milestone ladder

Tier **D1**. One PR per row — **41 rows, M-1 … M9d** (count = the table; every v1 function name frozen — no ellipses);
counts regenerate from the frozen registry inventory (M3a). v1 = M9d.
`<xm:f>` route-through pullable after M2.

| PR | Ships | Own gate |
|---|---|---|
| **M-1** | Planning flip + §13 checklist + literal-masking correction + cache policy table | Docs-only |
| **M0** | `refs.zig` + typed coords + adapters + import gate | Byte-identical call sites |
| **M1a** | Tokenizer: new kinds + **Unicode identifiers incl. the XID data work item** (generated `XID_Start`/`XID_Continue` interval tables from pinned Unicode 17 `DerivedCoreProperties.txt` via a new `gen_unicode_tables.py` mode — stdlib has no XID tables; generator does casefold/nfc only, `:302`; SHA/version/license headers; allocation-free lookup; regen gate) + **`THIRD_PARTY_NOTICES`** (Unicode License V3 + header attribution; **`LICENSE` third-party-data carve-out = M-1 OWNER action, flagged**; notices in EVERY artifact — zon paths `build.zig.zon:7-19`, release staging `release.yml:81-96`, Homebrew, wheels+sdists `pyproject.toml:10-11,41-45` — with artifact-content gates) + **extensible error literals** + compat suite + correction fixtures + fuzz | Untouched-construct identity; astral/unknown-error round-trips; XID boundary tests |
| **M1b** | Oracle harness: 4 adapters, frozen extractor, semantic manifests, provenance, precedence rules, forced-full-calc + stale sentinel, reviewed regeneration | CI replay; sentinel proof |
| **M2** | Parser + printer + precedence + structured refs + array constants + leading-`=` + prefixes + refusals + name-resolution spec + parse limits | parse→print→parse; fuzz; alloc-failure |
| **M3a1** | `value.zig` (types, ScalarValue/Matrix invariants) + **shape/coercion table** (§5.3b) + **error-order contract** (§5.3c) + **`collation_v1`** (before any collation-dependent operator/function) + `excel_fp_rules_v1`/`ieee_fp_rules_v1` + `parseDecimal` primitive | Divergence ×2; collation fixtures |
| **M3a2** | `eval.zig` core (broadcast, lazy forms) + `env.zig` EvalEnv (+`logicalBlankCount`, `alignedRangeIterator`) + in-memory fake + registry framework + **frozen inventory** + minimal fns | Non-finite escape fuzz; env-fake unit suite |
| **M3b** | serial_date + criteria + RNG (PRNG KATs) + RunInputs + counted-allocator limits | Boundary oracles; criteria fuzz |
| **M4a** | metadata.xml typed reader + cm/vm resolution + dialect primitives; **every record type classified — only `XLDAPR` interpreted; `XLRICHVALUE`/unknown value metadata on any input or result cell → pre-mutation typed refusal** | Fixtures incl. rich-value refusal; fuzz |
| **M4b1** | EvalEnv adapter + decode boundary (**decoding split by carrier class (corrected)**: FORMULA carriers — `<f>` bodies, defined-name bodies, table formulas — **XML-entity decoding ONLY** (ST_Xstring does NOT apply; literal `"_x0041_"` in a formula survives byte-exact, round-trip oracle); STRING carriers — SST, inline, `t="str"` `<v>` — XML-entity THEN ST_Xstring, RICH strings concatenating all visible `<r>/<t>` runs in document order, `<rPh>` excluded (inline reader returns first `<t>` only, `sheet_xml.zig:552-563`; SST rich `sst_xml.zig:56-66`); authored STRING output ST_Xstring-encodes; authored FORMULA output XML-escapes only (raw borrowed XML today: `sheet_xml.zig:522`, `workbook_xml.zig:61`); authored formula text **XML-escapes ONLY — no ST_Xstring stage** (string carriers alone use ST_Xstring); authored-formula tests prove literal `_x0041_` stays literal; fixtures with `_x005F_`, encoded controls, and entity escapes in formulas AND names — `_xHHHH_`, escaped `_x005F_`, C0 controls, across SST/inline/`t="str"`; the SST view decodes XML entities only today, `sst_xml.zig:174-193`) + **input cell-type contract** (§5.7.2) + **implicit-coordinate reconstruction** (a `<c>` without `r` takes the column after the preceding cell — Office semantics, MS-OE376 §2.1.624; the parser SKIPS such cells today, `sheet_xml.zig:507`; fixtures: first-cell-no-r, gaps, formulas, out-of-grid) + symbol layer + namespace preflight | Decode/`'R&D'`/prefixed fixtures |
| **M4b2** ✅ | Full calc-state parse (`CT_CalcPr` complete, `sheetCalcPr`, `date1904`, extensions byte-exact, `fullPrecision="0"` refuses) + **CT_CellFormula attribute inventory** (13 rows, one fixture each, unknown refuses) + **attribute-based** shared classification + sheet-wide topology validation + **AST copy-translation** (relative halves by (Δrow,Δcol); off-grid collapses the whole operand to `#REF!`) + `t="d"`'s normative lexical table | Slave-shape + topology matrix; attribute fixtures; 2 000-case randomized differential vs an independent translator; corpus `<calcPr>` byte round-trip; attribute/calc-state fuzz |
| **M4b3** ✅ | Name resolution + table producers + 3D matrix + cache-based `evaluate` | Site semantics; opaque names; 3D fixtures |
| **M4c** ✅ | F1a-1 (20: operators; IF, AND, OR, NOT, IFERROR, IFNA, IFS, SWITCH; ISBLANK, ISNUMBER, ISTEXT, ISERROR, ISERR, ISNA, ISLOGICAL, NA, N, T; **TRUE, FALSE** — added at M3a2, see the decisions block) | Oracle-first |
| **M4d** ✅ | F1a-2 (17: ABS, ROUND, ROUNDUP, ROUNDDOWN, INT, TRUNC, MOD, POWER, SQRT, EXP, LN, LOG, LOG10, SIGN, PI, RAND, RANDBETWEEN — **SQRT and RAND pinned here, registered at M3a2**) + multi-callsite/lazy-branch draw KATs | Oracle-first; KATs |
| **M4e** ✅ | F1b (22: SUM, COUNT, COUNTA, COUNTBLANK, AVERAGE, MIN, MAX, SUMIF, COUNTIF, AVERAGEIF, SUMPRODUCT; VLOOKUP, HLOOKUP, INDEX, MATCH, XLOOKUP, XMATCH, CHOOSE, ROW, ROWS, COLUMN, COLUMNS — **the seven M3a2 framework subjects pinned here, registered at M3a2**) — **Core gate 59** | Oracle-first |
| **M4f** ✅ | F1c-text (19: LEFT, RIGHT, MID, LEN, LOWER, UPPER, TRIM, CONCAT, CONCATENATE, TEXTJOIN, SUBSTITUTE, REPLACE, FIND, SEARCH, EXACT, VALUE, REPT, CHAR, CODE) + **CV1/CV2 shared text layer** (§5.4d; collation_v1 landed at M3a) + **`casing_v1`** + the `unicode/` move | Oracle-first; codec tests; per-CV fixtures |
| **M4g** ✅ | F1c-date (15: DATE, YEAR, MONTH, DAY, HOUR, MINUTE, SECOND, TODAY, NOW, EOMONTH, EDATE, WEEKDAY, DATEVALUE, TIMEVALUE, TIME) + the **invariant date grammar** (§5.4b) + `RunInputs`' clock reaching the evaluator | Oracle-first; per-epoch fixtures |
| **M5a1** ✅ | graph.zig: node model, SCC condensation, deterministic order, **seed table**, range-order contract; closure eval semantics | Scaling assertion; order fixtures; **randomized differential test vs a brute-force graph builder** (overlaps, full rows/cols, 3D spans, names, spill resize/invalidation — a missed edge passes perf tests but corrupts caches) |
| **M5a2** ✅ | Iteration engine (multi-SCC schedule, convergence, clamps) + callsite-keyed volatile schedule + rebuild-reuse KATs + dynamic-edge fixpoint + **complete oracle-gated INDIRECT + OFFSET contracts** (the fixpoint's test subjects; registered fully here so M6's public CLI never exposes a half-function) | Iteration oracles; stabilization fuzz; INDIRECT/OFFSET fixtures |
| **M5b0** | **`SourceBacking`** — ref-counted file/buffer backing shared across PartStore generations (each store exclusively owns + closes one `std.Io.File`, `store.zig:105-129,326-329`; shallow clone double-closes, moving breaks retention); backing unified; repeated-recalc + ownership tests. **Ladder-ordered FIRST of the M5b group — physically before M5b1/M5b2, because the transaction that requires it cannot land earlier than it** | Ownership tests; double-close fuzz |
| **M5b1** | `ResolvedSheet` projection + cached-value patcher + transitions (**incl. ST_Xstring output encoding**) + fuzz | Byte-confinement; round-trip |
| **M5b2** | Prepare/swap transaction (complete state, reports pre-swap) + calcChain rel-resolution + calc-state writes + `markRecalcOnLoad` + diagnostics. **Hard dependency on M5b0** — whole-generation retention (§5.7.4) is unsafe while `PartStore` exclusively owns and closes its own file, so M5b2's gate re-runs M5b0's ownership tests | No-fail-swap proof; post-failure reads; raw-entry identity; refusal purity; **M5b0 ownership tests green** |
| **M5c** | `Workbook.openBuffer(alloc, io, bytes)` + **`Writer.saveToOwnedBuffer`** + `zlsx_recalc` **importable shell** (module graph only — its public composition ops land in M5d, where the consumer test moves) | Module-graph gate; buffer≡path byte-equivalence |
| **M5d1** | Archive/durability substrate: **`AtomicFile.finish` fsync fix** + commit region + **cancellation-aware serialization seam** (the context-free `DeflateFn` at `pkg/zip.zig:64-68,140-144` AND PartStore's whole-input compression during preparation, `pkg/store.zig:370-385,498-512,638-656` — the shared compressor becomes context/callback-aware with 64 KiB chunks; the same chunking covers **decompression during model materialization** (`store.zig:881-916,1361-1384`), **raw-entry copies at save** (`:783-792`), XML scans, and temp-file writes, so the §5.5 polling bound holds across every long operation; cancel-inside-entry, cancel-inside-replacePart, and cancel-inside-materialization tests. **Control-aware buffer variants** — `saveToOwnedBufferControlled` / `openBufferControlled` (§5.10) wired to these seams, with cancel-mid-fresh-serialization and cancel-mid-buffer-open tests, so the Writer path meets the same bound. **Documented SLA exceptions: the blocking `File.sync` AND the post-commit POSIX directory fsync cannot be polled** — both uncancellable waits (no timeout; post-commit status is already success per §5.7.9), fault-injected tests for each, incl. Python worker-thread wait behavior) | Injected sync/rename failures; cancel-inside-* tests |
| **M5d2** | `recalculate()` + `saveWithRecalc` (ordering §5.7.9) + report + pre-M7 gate + logical-view gate + embedding-staleness preflight | Determinism; scoped idempotence; no-formula identity; confinement |
| **M5d3** | Writer compose + `zlsx_recalc` composition ops + **consumer dependency test** + committed bench workloads | Module-graph gate; bench baseline |
| **M6** | CLI: NDJSON schemas, exit tables (incl. 130/143 + prefix-valid stream rule), mandatory `--sheet/--name`, `--out` identity | Contract tests |
| **M7a** | DA evaluation + decision table + ownership + `A1#`/`@` + F2-DA natives (FILTER, SORT, SORTBY, UNIQUE, SEQUENCE, RANDARRAY, TRANSPOSE) + RANDARRAY KATs | Fixtures; obstruction fuzz |
| **M7b1** | DA + CSE **persistence** (approved set; cm/vm collection-pinned transitions; tail ownership; tolerated-state proofs) | Excel-opens-clean per transition |
| **M7b2** | F2 criteria batch (SUMIFS, COUNTIFS, AVERAGEIFS, MINIFS, MAXIFS; ADDRESS) — INDIRECT/OFFSET completed at M5a2 | Oracle-first; whole-column + multi-criteria benches |
| **M7b3** | F2 statistics batch (MEDIAN, MODE.SNGL, STDEV.P/S, VAR.P/S, PERCENTILE.INC, QUARTILE.INC, RANK.EQ, LARGE, SMALL) | Oracle-first |
| **M7c** | DA authoring (byte-diffed spec → impl) + `FormulaWrite` (**Zig-only**) | Byte-diff vs references |
| **M8a** | **`numfmt_v1` versioned grammar + support matrix FIRST** (sections ≤4, conditions, escapes/fills, fractions, scientific, elapsed `[h]`/`[mm]`, locale `[$-409]` tags — each row supported-with-exact-rendering or typed-refusal; the repo has only a date-detection heuristic today, `src/xlsx.zig:3909`) + numfmt + TEXT | Format fuzz; TEXT matrix; per-row grammar fixtures |
| **M8b** | PROPER (word segmentation over the M4f `casing_v1` module) | Segmentation fixtures |
| **M8c** | F3 batch (NUMBERVALUE, FIXED, DOLLAR, CLEAN, UNICHAR, UNICODE, TEXTBEFORE, TEXTAFTER, TEXTSPLIT; NETWORKDAYS(.INTL), WORKDAY(.INTL), DATEDIF, DAYS, DAYS360, YEARFRAC, ISOWEEKNUM, WEEKNUM) | Oracle-first |
| **M9a1** | C ABI part 1: `zlsx_status_v1` + descriptor types + editor recalc/evaluate + release fns + **`zlsx_engine_fingerprint()` export (header + `_HAS_FINGERPRINT` probe — M9b depends on it)** + **`zlsx_editor_mark_recalc_on_load` (header + `_HAS_MARK_RECALC` probe + old-dylib skip)** + narrowing tests + design note | 3-file txn; probes; ABI fuzz |
| **M9a2** | C ABI part 2: `save_to_buffer`, `open_buffer`, writer exports (incl. `…_with_formulas_v2` dialect) + Python Editor/Writer methods + `finally` cleanup | 3-file txn; probes; boundary tests |
| **M9b** | Spark batch recalc: `zlsx.recalc` activation, read-only guarantee, digest-verified partitions, driver-inference-on-recalced-snapshot, per-executor digest-keyed cache note, retry tests, streaming refusal | Integration; retries; serverless verification |
| **M9c1** | **Shared deterministic solver contract FIRST** (iterations ≤128 charged to a shared **`WorkBudget`** threaded through evaluator + solvers — units: node 1, solver iteration 4, nested callbacks re-charge; combined-exhaustion tests; poll points; Excel-compatible guesses — RATE/IRR 0.1 — and root selection; pinned tolerance; `#NUM!` on domain/convergence failure) + F4a-TVM (7, frozen: PMT, IPMT, PPMT, PV, FV, RATE, NPER) | Oracle-first; convergence/non-convergence/cancellation fixtures |
| **M9c2** | F4a-flows (8, frozen: NPV, IRR, XNPV, XIRR, SLN, SYD, DB, DDB) | Oracle-first; solver fixtures |
| **M9d** | F4b engineering (20, frozen: CONVERT, DELTA, GESTEP, BIN2DEC, DEC2BIN, HEX2DEC, DEC2HEX, OCT2DEC, DEC2OCT, BITAND, BITOR, BITXOR, BITLSHIFT, BITRSHIFT, COMPLEX, IMREAL, IMAGINARY, IMABS, IMSUM, IMPRODUCT); **v1 complete**; §13 release gate | Oracle-first; rg allowlist; **absolute + regression perf checks (§9)** |

**M10+ backlog**: F5 (census-ordered); `<xm:f>` route-through; namespace-aware
scanners; **future compatibility versions beyond CV2** (CV1+CV2 are v1 scope,
§5.4d); `fullPrecision` support; relative names; TZif local-timezone
resolution; streaming recalc; rewriter spill-shifting; LAMBDA/LET;
drop-cached-values; multi-area top-level Excel-match; rich-metadata authoring
beyond M7c.

---

## 7. Function registry

Framework (M3a): comptime table — name, arity + per-slot laziness, coercion
classes, volatility (`ca` excluded — cell scheduling only), propagation class
(`propagate|observe|per_element|per_function_provenance` — no generic
skip-errors class survives §5.3c), reference-producing,
DA-awareness, locale/collation/platform/CV flags, impl fn. **Frozen inventory
file (committed M3a) is the authoritative count source** —
`src/formula/function_inventory_v1.tsv`, **175 rows** of
`name<TAB>milestone<TAB>batch`, sorted and unique; every F-batch PR
regenerates its count from it. The five metadata fields carry no defaults and
a `@typeInfo` test proves it. Unregistered calls → typed refusal; `#NAME?`
only for value-position resolution failure. Batch contents are enumerated in
the ladder rows above (single source, no duplicate list to drift).

**M4c decisions (shipped 2026-08-04).** Twelve points, in
`src/formula/registry.zig` (six new rows and their implementations, the
batch gates) and `eval.zig` (fixtures only). The first F-batch, and the
row that discovered how little the committed oracle decides about
functions.

1. **The committed manifests decide exactly one cell of this batch, and
   the fixtures say so out loud.** `TRUE()+1` = 2, recorded by both the
   hand-spec excel suite and the LibreOffice one; nothing else in
   §8.2's evidence touches F1a-1. Every other fixture ships
   `spec_pinned` — and the label is **checked against the manifests**
   rather than asserted: a row claiming oracle evidence no manifest
   holds fails, and so does a row that ships spec-pinned while a
   manifest decides it. The oracle-row count is pinned at 1, so when
   the parked Excel leg runs and the suite grows F1a-1 rows, the row
   that moves the count is the row that re-labels. Same discipline
   §5.6g's 3D matrix shipped under at M4b3 (decision 14).
2. **`N` is not the `.number` coercion class, and that is this row's one
   real trap.** Excel's `N` carries its own conversion table — a number
   is itself, `TRUE`/`FALSE` are 1 and 0, an error is the error, and
   *everything else* is 0 — so `N("7")` is **0** where the numeric class
   answers 7. Its slot is therefore `.value_any` and the table lives in
   the implementation. `T` mirrors it: text is itself, an error is the
   error, everything else is `""`, and `T(1)` is `""` rather than `"1"`.
   Both are fixtured at the row a reasonable reading gets wrong.
3. **`N` and `T` propagate where the IS-family observes.** §5.3c's
   `observe` means "looks at an error without becoming one", and N and T
   *become* it. So they sit two rows from `ISERR` in the same table
   carrying the opposite class — which is precisely what §5.3c means by
   provenance-aware **per function, never family-wide**.
4. **`ISNA` asks by spelling, not by enum tag.** `ErrorValue` has two
   arms and only `known` carries a `KnownError`; matching the tag alone
   would answer FALSE for an `#N/A` that reached the evaluator through
   M1a's extensible-literal rule as a rich spelling. `ISNA` is a
   question about the error a user sees, and what they see is the
   spelling.
5. **The ladder's "lazy forms" is a grouping, not a contract.** The row
   names IF, IFS, SWITCH, IFERROR and IFNA together, but §5.3a defers
   three of them and declares the other two **EAGER**: Excel evaluates
   every `IFS` and `SWITCH` arm, observably through volatiles. So the
   proof for the eager pair is the *opposite* of the proof for the lazy
   trio — `IFS(FALSE(),RAND(),TRUE(),RAND())` must draw **twice**, and a
   fixture proving its untaken arm neither evaluated nor drew would have
   been a fixture proving zlsx wrong. Both directions are draw-counted,
   because a draw count is the one instrument a right answer arrived at
   wrongly cannot satisfy.
6. **`IFERROR` and `IFNA` invert §5.3c's first-error rule, and one pair
   of cells proves it.** `IFERROR(A5,A6)` is `#N/A` — the *fallback's*
   error, not the first argument's — while `IFNA(A5,A6)` is `#DIV/0!`,
   the first argument's after all, because `#DIV/0!` is not what IFNA
   catches. Same two cells, two functions, opposite answers; every
   multi-argument name in the batch is fixtured in both argument orders,
   and the coverage list is derived from the registry's own arity so a
   function that gains an argument later cannot slip past unordered.
7. **The IS-family reduces an array to its top-left, and the spill is
   deferred to M7a on purpose.** M3a2 shipped `ISERROR`/`ISNUMBER`/
   `ISBLANK` as `.value_any` with top-left reduction; the four new
   members follow that convention rather than starting a second one.
   Excel 365 spills `=ISNUMBER(A1:A3)`, and the registry already carries
   `da_aware` for exactly that decision — taking it here would have been
   M7a's call made early and in the wrong file.
8. **The batch's size is regenerated in both directions and written
   down in neither.** §7 makes the TSV the count source, so the gates
   derive the twenty from `milestone == "M4c"`: every tagged row must
   resolve through `registry.lookup` (no omissions), and every
   registered function whose inventory row says `M4c` must be listed
   under `F1a-1` (no substitutions). The ladder's own "20" is checked
   against the same file by M3a2's per-milestone test, so prose and code
   meet at the data rather than at each other.
9. **`TRUE` and `FALSE` are proved ordinary rather than assumed so.**
   They joined F1a-1 at M3a2 (decision 2) instead of being written for
   it, and a zero-argument signature is the shape most likely to have
   been waved through. Their five fields are asserted row by row —
   including that a no-slot signature is **not** liftable, which is true
   for a reason (there is nothing to lift) and not by accident.
10. **The fuzz target varies argument *shapes* and asserts determinism,
    not merely survival.** `fuzzEvalTarget` already fuzzes formula text;
    this one fixes the grammar and varies what reaches a registered slot
    — every §5.3b provenance, references, multi-area sets, both array
    orientations, an omitted argument, an intersection, a nested call
    and a locale-flavoured refusal. Each input is evaluated **twice**,
    by two live evaluators over one arena, and the two results must
    agree; a refusal must be the same refusal both times. "Evaluates two
    ways" is the failure a single run cannot see.
11. **The five conditional names needed no evaluator change.** M3a2
    built `if_form`, `iferror_form` and `ifna_form` and left
    `IFS`/`SWITCH` as eager `.plain` implementations, so F1a-1 is six
    registry rows plus their implementations and `eval.zig` gained
    fixtures only. The row's scope allowed touching the evaluator "where
    a lazy form needs it"; nothing did.
12. **What this row does NOT pin, said here so the table cannot be
    misread.** The registry holds 29 entries and F1a-1 is 20 of them.
    The other nine — `SUM`, `COUNT`, `COUNTA`, `COUNTBLANK`, `COUNTIF`,
    `SUMIF`, `SQRT`, `RAND`, `CHOOSE` — are M3a2's framework test
    subjects, registered and tested but **not oracle-pinned**; M4d and
    M4e own them. "The registry has 29 entries" is not "29 functions
    have shipped".

**M4d decisions (shipped 2026-08-04).** Thirteen points, in
`src/formula/registry.zig` (fifteen new rows and their implementations,
the batch gates), `eval.zig` (fixtures and KATs only), and `rng.zig`
(one seam). The batch where the two fidelity rule tables stop being a
property of the *parser* and become a property of a **function**.

1. **The oracle decides one cell of seventeen functions, and its
   evidence is a disagreement.** `SQRT(-1)` is the only F1a-2 formula
   any committed manifest contains — and both excel-fidelity manifests
   contain it with *different* answers: the hand-spec suite records
   `#NUM!` (Excel's), LibreOffice records `#VALUE!`. Those two facts are
   not in tension, they are the same fact: at most one of two files
   claiming `"fidelity": "excel"` is Excel, which is what makes the row
   about the adapter rather than about the rule. It stays named in
   `excel_adapter_divergences` beside `-0`, the tie's skip count stays
   pinned at 2, and a new M4d test asserts the **disagreement itself**
   rather than only the skip — averaging the two would produce an answer
   neither adapter gave. zlsx answers `#NUM!` in both modes, because a
   radicand's domain is not a floating-point question.
2. **A recorded manifest cell is not necessarily a deciding one, and
   this row is where that started to matter.** M4c asked a two-valued
   question — does any manifest contain this formula? F1a-1 held no
   volatile, so it never met the third answer. The LibreOffice suite
   records a `RAND()` cell marked `"excluded": "volatile_formula"`;
   under M4c's question that reads as evidence, and this batch registers
   `RAND`. So M4d's evidence check is three-valued — `silent`,
   `decided`, `excluded` — with `excluded` its own assertion and its own
   pinned count, and a fixture claiming evidence from an excluded cell
   fails by name. §8.2 excludes volatiles from every external value
   oracle; the checker now enforces that instead of assuming it.
3. **`SQRT` and `RAND` are pinned where they stand.** Both were
   registered at M3a2 as framework subjects. M4d holds them to the same
   five-field check, the same fixture-per-name coverage, and the same
   evidence labelling as the fifteen the row writes — but does not move
   their table rows, because relocating a row to demonstrate that it
   belongs to a batch would demonstrate something about the file rather
   than about the table. The tests read the inventory instead.
4. **The two rule tables reach this batch through exactly one function.**
   N2's zero-snap is additive-scope-only and never applies to a function
   result; N3's signed-zero policy applies at *publication*, not at
   production; N4a is shared by both modes. What is left is which
   decimal a rounding decision is taken on, and that is `decimalView` —
   `excel_fp_rules_v1` reads N1a's 15 significant digits, and
   `ieee_fp_rules_v1` says `literal_significant_digits = null` precisely
   so the same call is a no-op there. One place, both directions,
   fixtured in both modes.
5. **The textbook `ROUND(2.675, 2)` case does NOT diverge, and finding
   out why located the real divergence.** 2.675 is really
   2.67499999999999982 — but multiplying it by 100 lands within half a
   ULP of 267.5, so the *scaling* rounds up and both modes answer 2.68.
   The divergence lives where the scaled value stays strictly below the
   half-way point, which needs an argument carrying 16–17 significant
   digits. The fixture cells are therefore **stored values**, not
   literals: N1a would round a 17-digit literal to 15 at ingress under
   `.excel`, and the divergence under test would have been the parser's.
6. **`-0` is produced here and normalized later, so the collapse branch
   carries a sign.** `ROUNDDOWN(-0.4, 0)` is `-0` before publication —
   preserved bitwise by `.ieee`, normalized to `+0` by `.excel`. That
   holds for the branch where every significant digit is rounded away
   too, which is why it returns `copysign(0, n)` rather than a literal
   zero; a bare `0` would have silently collapsed a mode divergence into
   agreement. Comparisons in the fidelity fixtures are on **published**
   values, bit for bit, because a comparison anywhere else cannot see
   this at all.
7. **The scale factor need not be representable, and the fix is a
   bound rather than an invented digit limit.** `10^d` overflows above
   `d = 308` and underflows below `-324`, while `d` arrives as an
   arbitrary f64 a user typed. `roundAt` decides both extremes *before*
   scaling, by comparing the requested place with the value's own
   decimal exponent: `d ≥ 17 − e` is already exact (binary64 carries no
   more than 17 significant digits), and `d + e ≤ −2` removes all of
   them. What survives is bounded — `|n·10^d| < 10^17` by construction —
   and the scaling splits into ≤300-decade steps so no *intermediate*
   leaves range either. `ROUNDUP` at a collapsed place is the one case
   that legitimately overflows, and N4a answers `#NUM!`.
8. **`POWER` is the function spelling of `^` and shares its arithmetic
   exactly — including where that is arguably wrong.** `0^0` is 1 here,
   inherited from the operator; Excel answers `#NUM!` and LibreOffice
   answers 1, and **no committed manifest records the cell**. Changing
   it is an operator-level decision, and a functions row with no
   evidence is the wrong place to take one. So the row pins the
   *identity* instead — `POWER(a,b)` and `a^b` agree bit for bit in both
   modes across negative bases, fractional exponents, overflow and
   `0^0` — which is the property a workbook actually depends on, and
   which makes the operator the single place a future oracle leg has to
   change.
9. **`MOD`'s quirk is the sign, and only the sign.** §5.4's N4 names it
   specifically, so the implementation is floored modulus written out
   (`n − d·floor(n/d)`) and the quotient is deliberately **not** read
   through `decimalView`: widening MOD's quirk list would be inventing
   an Excel behaviour no manifest recorded. Both modes therefore agree
   on every MOD fixture, and the overflow of an extreme ratio reaches
   `#NUM!` through N4a rather than through a magnitude test.
10. **`RANDBETWEEN` draws exactly once, and that is a decision about
    the instrument.** `rng_v1` has an exactly-uniform rejection sampler
    (`nextIntInclusive`, M3b) and this function does not use it: a
    rejection sampler draws a data-dependent number of times, which
    would make the draw *counter* — the instrument every §5.6d KAT is
    built on — unable to state anything. Scaling one draw is off perfect
    uniformity by at most one part in 2^53. The `@min` clamp covers the
    one case scaling cannot, a draw just under 1 against a wide span,
    and a sweep over the ranges and draws that break it asserts every
    result is an integer inside `[bottom, top]`.
11. **Non-integer `RANDBETWEEN` bounds move INWARD.** `ceil` the bottom,
    `floor` the top, so `RANDBETWEEN(1.5, 3.5)` draws from `{2,3}`. The
    alternative reading — truncate both toward zero — lets it answer 1,
    a value outside the interval the caller wrote. Nothing on disk
    decides this, so the fixture ships `spec_pinned` and pins the
    invariant a reader can check instead: every result lies within
    `[bottom, top]`, and a range holding no integer is `#NUM!`.
12. **`LOG`'s base-1 failure is `#DIV/0!`, not `#NUM!`.** It divides by
    `LN(1)`, and that is the one place this family answers with
    something other than a domain error — so it is its own line, its own
    fixture, and its own error-order case proving an argument's error
    still beats it. `LOG(x)` and `LOG10(x)` are one operation under two
    spellings, as `TRUNC(x)` and `ROUNDDOWN(x,0)` are; both equivalences
    are asserted rather than left to be inferred from two
    implementations.
13. **Reproducibility is a property of the seam, and the fuzz property
    is proved by enumeration.** `Rng.fromRunInputs` is the only way a
    run's generator is built, so "the draw sequence comes from
    `RunInputs` alone" is a statement about an argument list — the KAT
    ties the *evaluator's* first two draws to `rng_v1`'s own stream for
    that seed, and proves a dead lazy arm consumes nothing by showing
    the live arm receives the stream's **first** value. Separately: the
    `fuzz` step is Linux-only (coverage-guided fuzzing is broken
    upstream on macOS and Windows), so the batch is also swept
    **exhaustively** — every name against every argument shape at one
    and two arguments, in both modes, each input evaluated twice. A
    sweep that always runs beats a search that runs on one platform.

**M4e decisions (shipped 2026-08-04).** Thirteen points, in
`src/formula/registry.zig` (fifteen new rows, their implementations,
the batch gates) and `eval.zig` (fixtures only). The row that closes
**Core gate 59** — and the row where the dispatcher's first-error rule
turned out to be unable to state §5.3c for a whole family.

1. **The committed manifests decide NOTHING about this batch, and the
   row states that as a number rather than as an absence.** Not one
   cell of the three manifests calls any of the twenty-two — §8.2's
   evidence is eighteen operator and literal cells plus `SQRT(-1)`. So
   every fixture ships `spec_pinned`, the oracle-row count is pinned at
   **zero**, and a second check reads the *manifests* and fails if any
   cell mentions an F1b name. That second direction is the point: "no
   manifest decides this" is the one claim a fixture table cannot make
   about itself, and M4c's one deciding cell and M4d's one were both
   found by asking the files. The three-valued verdict (`silent` /
   `decided` / `excluded`, M4d decision 2) is reused rather than
   re-derived, so this row gets the excluded-cell guard even though it
   holds no volatile of its own.
2. **The five lookups are `per_function_provenance`, and finding out
   why was this row's real discovery.** `propagateAndInvoke` scans
   *evaluated scalars*. A lookup's key slot is `.value_any` and a key
   normally arrives as a **reference**, which the scan cannot see at
   all — so under `.propagate`, `VLOOKUP(A5,table,A6)` answered `#N/A`:
   the scan found the third argument's scalar error and never saw the
   first argument's. §5.3c's declaration order was wrong *by
   construction*. The same slot shape carries the other half: an error
   inside a lookup TABLE must not propagate, because it is a value the
   lookup may return. Both are one rule in `lookupPropagate`, which
   walks the arguments in order with the table slots masked out — and
   `per_function_provenance` is precisely the class §5.3c provides for
   a function whose answer depends on **where** an error was found.
3. **MIN and MAX are the comparator's named exception, and the sharp
   fixture is MAX rather than MIN.** Under `collation_v1`'s cross-type
   ranking any text outranks every number, so a `MAX` that used the
   comparator would answer `"gamma"` over a column holding 1…4 and
   three words. It answers 4. §5.4b's three-part rule — a direct text
   argument coerces or is `#VALUE!`, text in a range is ignored, no
   numbers anywhere is 0 — is fixtured in all three positions. The fold
   deliberately does **not** use `@min`/`@max`: those are IEEE
   minNum/maxNum, which treat `-0` and `+0` as interchangeable and
   would have silently decided the one case N3 makes observable, in
   whichever direction the hardware felt like.
4. **The two rule tables reach this batch through the fold, not through
   a rounding argument.** M4d's divergence was `decimalView`; nothing
   in F1b rounds to a decimal place, so that seam is untouched. What is
   left is N2's **additive scope** — and an aggregate is a chain of
   additions. `SUM(0.1,0.2,-0.3)` snaps to 0 under `excel_fp_rules_v1`
   and stays `5.551115123125783e-17` under `ieee_fp_rules_v1`;
   `AVERAGE` and `SUMPRODUCT` inherit it through the same accumulator.
   N3's half needs a function that RETURNS an input rather than
   computing one, which in this batch is only MIN and MAX: an
   accumulator would have added `-0` to `+0` and lost the sign on the
   way. The converse test proves the rest of the batch agrees.
5. **§5.6g's frozen six are all here, so the 3D matrix is finally
   fixtured end to end.** M4b3 shipped it entirely spec-pinned and
   could only run the three the registry then held (its decision 14);
   `AVERAGE`, `MIN` and `MAX` arrived at this row. All six now
   aggregate a real span — six functions, six different right answers
   over the same two cells, which is also the proof they aggregate the
   span rather than one member — and the other **sixteen** refuse one
   *typed*, carrying `three_d_ineligible_function` rather than a
   generic failure. Neither list is retyped: the eligible one is read
   from `names.three_d_eligible` and the refusing one is derived from
   the inventory, so a seventh eligible name cannot ship without a span
   fixture and a twenty-third batch name cannot ship in neither
   direction.
6. **CHOOSE's laziness is proved by a draw count in every arm position,
   including a dead arm at exactly zero.** §5.3a assigns its fixtures
   here. A three-arm call draws **once** whichever arm is taken;
   `CHOOSE(1,7,RAND())` draws **none**, which is the statement a result
   cannot make — under a constant source a draw that happened and one
   that did not look identical. An out-of-range selector and an
   erroring selector each take no arm and draw nothing. An **array**
   selector switches the whole form to per-element masking, both arms
   evaluate and both draw, and a per-element `#VALUE!` stays in its own
   cell — the opposite proof, from the same instrument.
7. **`Grid` is a materialization, and that is a decision about which
   failure a lookup is allowed to have.** Aggregate walking is a sparse
   fold and needs no coordinates; a lookup is a search along an axis
   and does. So a reference reaching a lookup is materialized under
   §9's `max_matrix_cells`, which makes an absurd rectangle a **limit**
   rather than a run that never ends: `MATCH(1,A1:XFD100000,0)` refuses
   where `SUM` over the same range still answers. Making a
   whole-column lookup *fast* is M7b2's row; making an impossible one
   refuse is this one's.
8. **Lookup equality is the criteria matcher's, and the lookup VALUE is
   not a criterion.** §5.4b names lookup equality and criteria in one
   sentence; the matcher is already type-restricted the way a lookup is
   (a numeric key never matches a text cell), already folds under
   `collation_v1`, and already implements `*`/`?`/`~`. So the criterion
   is **built** here rather than parsed — `criteria.parse` would read
   `VLOOKUP("<5",…)` as "less than 5" where Excel looks for the literal
   text `<5`. Wildcards are switched per call, which is what makes
   `XMATCH("t*",…,0)` `#N/A` and `XMATCH("t*",…,2)` a hit.
9. **An ordered match never crosses a type boundary, and ties go to the
   last position in array order.** The cross-type ranking is a *total*
   order, so without a type restriction a text cell would qualify as
   "≥" a numeric key and `MATCH(5, a_column_of_words, 1)` would find
   one; with it, that is `#N/A`. Ties resolve by position rather than
   by scan direction, which is what makes `search_mode` change an exact
   match's answer and leave an ordered one alone — the way Excel's
   binary search does over the sorted input it documents as a
   requirement. Blank and error cells are not ordered candidates: the
   pass is looking for a value, and `#N/A` in a column is not an answer
   to "where is 5".
10. **XLOOKUP's `if_not_found` is masked with the arrays, because it is
    a value the call may RETURN.** It is not an array, but it is on the
    same side of decision 2's line for the same reason: propagating it
    would make a successful lookup fail over a fallback nobody reached.
    `XLOOKUP(20,…,A5)` is the hit and `XLOOKUP(99,…,A5)` is A5 — one
    cell, two outcomes, and the pair is the fixture. The reading that
    propagates it was tried first and failed on a shape rather than on
    an error: a perfectly ordinary **range** fallback made every hit
    `#VALUE!`, because a multi-cell reference reduces to `#VALUE!` in a
    scalar context. Laziness is still NOT the answer here — §5.3a
    enumerates the deferring forms and XLOOKUP is not among them, so
    the argument evaluates (and any volatile in it draws); it simply
    does not propagate. Declaration order is fixtured across slots 0
    and 4 instead, where both slots do propagate.
11. **INDEX returns a value, not a reference, and the deferral is
    stated.** `reference_producing` is `false` on it: Excel's
    `INDEX(…):INDEX(…)` range form needs the evaluator to carry a
    reference out of a call, and §7 lists exactly two reference-
    producing rows — `INDIRECT` and `OFFSET`, both M5a2's. What INDEX
    ships is the whole of its *value* form, including the zero-index
    slices that legitimately return an array.
12. **Every row in the registry is now pinned, and the gate is counted
    rather than read.** M3a2's seven borrowed framework subjects — SUM,
    COUNT, COUNTA, COUNTBLANK, COUNTIF, SUMIF, CHOOSE (M4c decision 12)
    — are held to the same five-field check, fixture-per-name coverage
    and evidence labelling as the fifteen this row writes, **without
    moving a table row**: relocating one to demonstrate that it belongs
    to a batch would demonstrate something about the file rather than
    about the table, exactly as M4d decided for SQRT and RAND. Core
    gate 59 is obtained by counting three milestones in the frozen TSV
    and asserting that all fifty-nine resolve through `registry.lookup`
    — so the ladder's figure and the implementation meet at the data.
13. **The exhaustive sweep pads to each name's own minimum arity,
    because otherwise it would prove nothing about half the batch.**
    M4d swept one and two arguments, which is every arity F1a-2 has.
    Ten F1b names have a minimum of three or more and refuse *before*
    reaching an implementation at one or two — so a sweep built the
    same way would have enumerated twenty-two names and exercised
    twelve. Padding by repeating the last shape keeps the enumeration
    the same size and puts every name in front of its own impl, in both
    rule tables, each input evaluated twice.

---

**M4f decisions (shipped 2026-08-04).** Twelve points, in
`unicode/casing.zig` + `unicode/tables/casing_data.zig` (new),
`src/formula/text.zig` (new), `src/formula/registry.zig` (nineteen rows,
their implementations, the batch gates), `criteria.zig` (the substring
matcher) and `eval.zig` (fixtures, plus the `&` cap). The row where a
count stopped being a property of a string.

1. **§5.4d's "five affected functions" is an undercount, and the row
    ships seven.** The prose names LEN/MID/FIND/SEARCH/REPLACE. `LEFT`
    and `RIGHT` take a count of characters in the same units, and Excel's
    own CV1 `LEFT` hands back half a surrogate pair exactly as `MID`
    does — so a build faithful to the list would have had `LEFT("😀",1)`
    return the whole emoji while `MID("😀",1,1)` refused, inside one
    workbook, for the same character. `cv_sensitive` is registry data
    checked against a list **in both directions**, so an eighth name
    cannot ship uncounted and a whole-string function cannot ship
    claiming to index.
2. **A CV1 index into an astral character is `ResultNotRepresentable`,
    not a fabricated value.** Excel answers with a lone surrogate, which
    its UTF-16 strings can hold and UTF-8 cannot. The alternatives were
    U+FFFD (a character the user never had) and rounding to a whole
    scalar (a length the caller never asked for); both invent. The
    refusal is typed, fixtured across `MID`/`LEFT`/`RIGHT`/`REPLACE`,
    and each fixture asserts that the SAME formula under CV2 is an
    ordinary answer — which is the sharpest available statement that the
    two versions are different semantics rather than a flag.
3. **Absent compatibility metadata is CV1, and the code said CV2.**
    `run_inputs.CalcState.text_compat` defaulted to `.cv2` — the right
    answer to "what does Excel write into a new workbook" and the wrong
    one to "what does a workbook that says nothing mean". §5.4d says
    absent = CV1; the files that say nothing are every pre-2024 workbook
    **and every file zlsx's own Writer emits** (`fresh_emit.zig` writes
    no metadata part). `calc.zig`'s parsed `TextCompat` already defaulted
    to `.v1`, so the two halves of the engine disagreed. Corrected, with
    the default asserted in a test rather than trusted.
4. **`casing_v1` is a second table, not a second use of the fold, and
    the generator says so in the file it writes.** `fold("ß")` is `"ss"`;
    `UPPER("ß")` is `"SS"`. One is a lowercase comparison key nobody
    displays, the other is a displayed value — no fold implements
    `UPPER`, and no one-to-one mapping does either, since one scalar
    becomes two. The generator's fourth mode layers SpecialCasing's 106
    unconditional rows over UnicodeData's simple mappings, ships
    Final_Sigma (the one conditional row with no language tag) with the
    `Cased`/`Case_Ignorable` intervals it needs, and **rejects the
    fifteen `tr`/`az`/`lt` rows by name, printing each one it drops**.
    There is no locale in `RunInputs` to select them with, so Turkish
    dotless-ı casing is recorded as a divergence rather than implemented.
5. **`UPPER("ß")` ships `spec_pinned` and FLAGGED.** §5.4b marks it
    oracle-pinned; the M1b Excel adapter that would pin it is parked on
    a modal dialog (§8.2). So the batch's oracle-row count is pinned at
    **zero** like M4c's, M4d's and M4e's, the three-valued checker
    (`silent`/`decided`/`excluded`) guards the label in both directions,
    and the test says in as many words which number has to move when
    that leg runs.
6. **`title` mappings ship now although nothing calls them.** M8b's
    `PROPER` is word segmentation *over this table*; shipping the
    character-level half here means M8b adds a segmenter rather than
    reopening the generator, and the generated file is one artifact with
    one pinned revision either way.
7. **`SEARCH` reuses the criteria matcher; a second one was written and
    thrown away.** The first attempt gave `casefold.zig` its own
    folded→original position map. M3b's `criteria.Folded` already had
    one, keyed by original code point — and two maps is precisely how
    `SEARCH("~*",…)` and `COUNTIF(…,"~*")` come to disagree about what a
    literal star is. The anchored matcher gained a **prefix mode and a
    start offset**, both parameters rather than a fork, and M4e's
    precedent (lookup equality is the criteria matcher's, decision 8)
    is the one being followed.
8. **Half an expansion is not a match, and that rule was inherited
    rather than chosen.** `SEARCH("ss","aßb")` is 2; `SEARCH("s","aßb")`
    is `#VALUE!`, because both sides must end on a code-point boundary —
    the same rule that makes the criterion `"?s"` fail against `"ß"`. A
    `SEARCH` answering 2 for the second would be claiming a position
    inside a character.
9. **The cell-text cap moved to the `&` operator, because M4f is the row
    that made it reachable.** §9 makes text past 32 767 code points an
    Excel-domain `#VALUE!` (a catchable error value) rather than a
    refusal. A string literal is bounded by the 8 192-byte formula
    length, so before `REPT` no formula could overflow a cell at all;
    `REPT("a",20000)&REPT("b",20000)` is where that stops being true.
    One rule, one constant, both sites — and `REPT` checks the COUNT
    rather than the result, so an absurd repetition is `#VALUE!` without
    a gigabyte on the way to it.
10. **`TEXTJOIN` is the first function in the ladder that can tell a
    blank from an absence, so it gets a dense walk.** §5.6a's sparse
    iteration is correct for every M4e aggregate — none of them counts a
    blank — but with `ignore_empty` false each blank is an empty field,
    and a skipped one drops a delimiter and shifts every position after
    it. The flag selects the walk, which is honest in both directions:
    with empties skipped a blank and an absence ARE the same thing and
    the sparse walk is both correct and cheaper. The dense walk is
    bounded by the accumulator rather than by the area, so
    `TEXTJOIN(",",FALSE,C:C)` stops at the cap instead of emitting a
    million delimiters.
11. **`VALUE` inherits §5.3b's three-way split rather than defining a
    fourth.** Numeric text is a number, non-numeric text is `#VALUE!`,
    and locale-flavoured text — `"1,5"`, `"50%"` — is a typed
    `LocaleSensitiveInput` refusal, because 1,5 is 1.5 in Germany and 15
    in the United States and zlsx has no locale to pick with. The
    consequence worth naming: `VALUE("50%")` refuses where Excel answers
    0.5. That follows from M3a1's classification, not from this row, and
    it is fixtured here so the next reader finds it stated.
12. **`CODE` answers 63 for anything outside the code page.** It reports
    a position in a 256-entry table, so a character with no position has
    no number to report and Excel substitutes `?`. `UNICODE` (M8c) is
    the function that answers the other question. `CHAR`'s five
    undefined windows-1252 slots (0x81, 0x8D, 0x8F, 0x90, 0x9D) map to
    the matching C1 control, which is what Windows itself does.

---

**M4g decisions (shipped 2026-08-04).** Nine points, in
`src/formula/serial_date.zig` (the clock conversion and the invariant
grammar), `registry.zig` (fifteen rows and their implementations) and
`eval.zig` (fixtures, and the clock reaching `Options`). The row where
the calendar and the serial turned out to disagree.

1. **`WEEKDAY` counts SERIALS, not calendar days, and under the 1900
    system those are different answers.** 1900-02-29 never happened but
    it occupies a serial, so serial arithmetic and the proleptic
    calendar drift apart by exactly one day everywhere below the gap.
    The first implementation read the calendar and was wrong: Excel
    answers **1 (Sunday) for `WEEKDAY(1)`** although 1900-01-01 was a
    Monday, and a workbook whose `dddd` format says Sunday must not
    have a `WEEKDAY` that says Monday. Above serial 60 the two agree
    again, which is why the divergence is invisible in every date
    anyone actually uses — and why it had to be found on purpose. The
    1904 system has no phantom day and therefore no drift; only its
    phase differs, because its serial 0 is a Friday.
2. **`DATEVALUE` refuses exactly when the locale would change the
    MEANING, which is M3a1's rule for numbers applied to dates.**
    `"1.5"` parses and `"1,5"` refuses; likewise `"1/15/2020"` parses —
    15 is not a month, so only one reading exists, in every locale — and
    `"1/2/2020"` refuses, because it is January 2nd in the United States
    and February 1st nearly everywhere else. ISO and named-month forms
    always parse: neither has a field order a locale gets to decide.
    The alternative considered and rejected was refusing every slash
    form, which is simpler to state and throws away dates that are not
    ambiguous at all.
3. **The refusal and the error are different kinds of answer, and the
    fixtures prove it with `IFERROR`.** `DATEVALUE("hello")` is
    `#VALUE!` — a value a formula can catch. `DATEVALUE("1/2/2020")` is
    `FormulaLocaleSensitiveInput` — a refusal that propagates past
    `IFERROR` to the caller. Collapsing them would either hide an
    ambiguity behind a catchable error or invent an error for text that
    is a perfectly good date.
4. **`TIMEVALUE` has no locale refusal, and that asymmetry is asserted
    rather than assumed.** A clock has one field order everywhere, so
    `LocaleOrdered` is unreachable from it; the test sweeps every hour
    against five minute values to say so.
5. **Two different two-digit-year rules ship, on purpose.**
    `DATE(20,1,1)` is **1920** — Excel's rule for a numeric year below
    1900 is `1900 + y`. `DATEVALUE("1/15/20")` is **2020** — the text
    grammar uses the documented 00–29 / 30–99 window. They are
    different rules for different inputs, both Excel's, and writing one
    of them twice would have been the easy mistake.
6. **`TODAY` and `NOW` are volatile for a different reason than `RAND`
    is.** A draw changes at every callsite; the clock is stable within a
    run and changes between recalculations. Both still re-key the M5a2
    schedule, so both carry `volatile_fn` — and the registry gate that
    said "RAND and RANDBETWEEN are the only volatile rows" was widened
    deliberately rather than deleted, keeping its both-directions shape.
7. **The instant is an INPUT, never a clock read.** `Options` carries
    `now_utc_ms` and `utc_offset_min` beside the epoch, so a recalc is
    reproducible from `RunInputs` alone and every volatile-date fixture
    pins an exact answer. The offset is a fixed civil offset and can
    move the DAY — 23:30 UTC plus 60 minutes is tomorrow — which is the
    case a UTC-only engine gets wrong and the fixtures cover in both
    directions. No OS timezone is consulted anywhere; TZif is M10+.
8. **`epoch_sensitive` is a second flag, not a reuse of
    `cv_sensitive`.** They are different workbook properties with
    different owners: one comes from `workbookPr@date1904` and changes
    what a serial MEANS, the other from the metadata part and changes
    what a character IS. Ten of the fifteen names carry it; `TIME`,
    `HOUR`, `MINUTE`, `SECOND` and `TIMEVALUE` do not, because a
    fraction of a day is the same fraction under either epoch.
9. **`EDATE` clamps and `DATE` overflows, and both are Excel.**
    `EDATE(2020-01-31, 1)` is February 29th — month arithmetic lands
    inside the month it names — while `DATE(2020,1,32)` is February 1st
    and `DATE(2020,13,1)` is January 2021, because a constructor's
    fields are a running total. One walk with two landings serves
    `EDATE` and `EOMONTH`, which is how Excel documents them and the
    only difference the code has.

---

**M5a1 decisions (shipped 2026-08-04).** Thirteen points, in
`src/formula/graph.zig` (new — the node model, the walk, the index, the
condensation, the seed table, the closure plan), `run_inputs.zig` (§9's
work counters), `symbols.zig` (table geometry, a public fold) and
`pkg/workbook.zig` (`evaluateClosure` and the bridge). The row where
the tokenizer's idea of where a reference ends turned out not to be the
reference.

1. **The left endpoint's qualifier governs the whole range operator, and
    the tokenizer's split is why that has to be said out loud.**
    `Sheet3!$C$5:$C$6` arrives as `qualified(Sheet3, C5)` *then* `:` then
    `C6` — the second endpoint sits **outside** the qualified node — so a
    walk that resolves each side as an independent reference puts `C6` on
    the referencing sheet and loses the area entirely. The identical
    shape governs 3D: `Sheet1:Sheet3!$B$1:$B$2` is one span over `B1:B2`,
    not a span over `B1` beside a stray `B2`. Both were found by the
    differential test and neither is visible in a fixture written from
    the spelling, which is the argument for the gate being what it is.
2. **Range nodes, not per-reader edges (§5.6b).** Two formulas reading
    `A1:A100` share one node, so the area is resolved against the
    producer index once instead of once per reader. That is the whole of
    "no O(F×R)": without it the edge count is readers × producers, with
    it it is readers + producers.
3. **The producer index is sorted twice, and the probe picks.**
    Row-major and column-major, and each area is probed through whichever
    of the two makes its *narrower* band the leading key. `SUM(A:A)`
    costs the cells stored in column A; `SUM(1:1)` costs the cells stored
    in row 1; neither costs a million coordinates. One order would have
    made exactly one of those two cheap. The instrument is
    `stats.index_probes` — a counter, so the assertion is "40 candidates
    examined", not "it felt fast".
4. **A reference to a cell holding no formula contributes no edge, and
    the rule is written in the module header rather than discovered from
    the code.** A constant cannot be recalculated, so it cannot constrain
    an order. It has to be stated because the brute-force builder has to
    share it, and a rule two implementations infer separately is a rule
    they will eventually infer differently.
5. **Kahn with a min-heap, not the reverse of Tarjan's output.** Both are
    deterministic; only one is **canonical**. Kahn emits, at every step,
    the ready component with the smallest node under `Key.order`, so the
    result depends on the graph and on nothing else — not on which node a
    depth-first search happened to start from. That is what makes "the
    same order across randomized insertion order" a property of the
    algorithm instead of a property of this seed.
6. **The seed table's malformed row is checked FIRST, before the
    array-anchor row.** An anchor with a declared shape never reads its
    cache, so the other ordering would let an anchor launder a malformed
    `<v>` into a zero-filled array — the exact substitution §5.6c
    forbids, arrived at by a path nobody would think to test.
7. **"No declared shape" and "not an array" are different facts, so
    `Anchor` is three-valued.** `.none` seeds from the cache, `.shape`
    seeds zero-filled, and only `.unknown` takes §5.6c's pre-iteration
    shape pass. A nullable shape would have collapsed the first and third
    into one, and the third is the only one the shape pass exists for.
8. **`max_eval_depth` counts cell hops, not node hops.** §9 says
    "cell → cell" and the counter has to mean it: the range, span, name
    and producer nodes model *how* one cell reaches another and consume
    no depth. Counting nodes would have made the effective limit depend
    on whether a dependency happened to be written as a cell or as a
    range.
9. **`max_total_cell_evals` is charged when the plan admits a cell, not
    when one is evaluated.** The plan is exact, so charging at admission
    refuses **before the first evaluation** rather than halfway through
    one — which makes a §9 refusal pre-mutation for free instead of by
    an unwind that has to be got right.
10. **The graph resolves structured references even though *evaluating*
    one is M7b's.** They are different obligations: an evaluator that
    refuses a construct returns an error value, while a graph that drops
    the edge writes a stale cache. Conservative wins, so `TableGeometry`
    and the row-band arithmetic ship here and `symbols.Table` gained
    `headerRowCount`/`totalsRowCount` to feed them.
11. **A position-dependent name body draws no edges, deliberately.**
    Referencing a relative name is refused at M4b3, so no cache can be
    computed from an edge the graph declines to draw; inventing a sheet
    for it would be the unsafe direction. The node still exists — a name
    is a node whether or not anything can be said about its body.
12. **A tail covered by two declared arrays belongs to the anchor that
    sorts first.** Overlapping anchors are an M7a obstruction and this
    row does not decide them — but it does decide that the answer is a
    *rule* rather than an accident, because "whichever cell arrived
    first" would make the node set depend on input order and break the
    determinism this row is gated on.
13. **A cycle refuses at M5a1 and the seed table ships anyway.** §5.6c's
    iteration-off rule is `FormulaCycle`, and M5a2 is where a workbook
    whose `calcPr` asks for iteration gets a schedule instead. The seeds
    are computed, fixtured per row, and unused — which is the point: the
    engine that will consume them arrives to a table that has already
    been tested against something.

**M5a2 decisions (shipped 2026-08-05).** Thirteen points, in
`src/formula/iterate.zig` (new — the multi-SCC schedule, convergence, the
two exhaustion outcomes, §5.6e's fixpoint), `src/formula/draws.zig` (new
— §5.6d's key and its memo), `eval.zig` (the draw seam, `INDIRECT` and
`OFFSET`), `registry.zig` (the two reference-producing rows),
`run_inputs.zig` (§9's two new bounds), `graph.zig` (the runtime-edge
seam and `PlanOptions`) and `pkg/workbook.zig` (the package's `Host`).
The row where the seed table stopped being a fixture and started being
load-bearing.

1.  **A zero `iterateDelta` is an unset attribute, not a request for
    exact equality.** §5.6c's rule is `abs(new − previous) < iterateDelta`,
    strictly, and at zero nothing satisfies it — not even a value that
    did not move. The transition table originally read the zero as
    "iterate until nothing changes at all", which is a coherent thing to
    want and is not what the spelling means; honouring it would need an
    exception *inside* the comparison, and a comparison with an
    exception in it is one two implementations will eventually disagree
    about. So zero joins the same row as negative, which is also the
    reasoning that already turns `iterateCount` zero into 100. Found by
    the signed-zero fixture, which converged at 0.001 and did not at 0.
2.  **The resource ceiling refuses only when it is STRICTLY lower.** At
    equality the caller permitted exactly what the file asked for and
    the file got it, so a refusal there would refuse a run nothing
    actually constrained. The three fixtures — ceiling above, equal,
    below — exist because only the middle one distinguishes this reading
    from the obvious `<=`.
3.  **Zero mutation is a rollback, not an absence of writing.**
    Gauss–Seidel visibility means a pass has to publish as it goes, so
    "the refusal wrote nothing" cannot be true by construction. Every
    publish is journalled and a refusal retracts them in reverse, and
    `Host.retract` is infallible for that reason: a rollback that could
    run out of memory would make the promise conditional on there being
    memory to keep it. The fixture asserts `publishes == retracts` as
    well as the empty state, because an engine that never wrote would
    pass the second assertion and fail the first.
4.  **The §5.6d key has four terms and deliberately no component.** The
    obvious spelling of "SCC-pass" is the component beside the pass
    number, and it is wrong exactly where §5.6e is: a changed SCC resets
    its pass counter and re-seeds, so the same cell re-runs pass 1 while
    belonging to a different component. With a component term that is a
    new key and a fresh draw — and §5.6e says in as many words that a
    discovery pass must not perturb a result. The path already names the
    cell.
5.  **`CallCtx.draw` became fallible, because a memo has to be stored
    somewhere.** The alternative — swallowing the allocation failure and
    drawing afresh — would let an out-of-memory condition silently
    change a result, which is the exact class of bug the memo exists to
    prevent.
6.  **The fixpoint test is on the graph, not on the edge set.** A read
    the static walk already found produces the same edge, so a workbook
    with no dynamic reference reports edges the graph already had;
    comparing edge sets would call that a change and cost every ordinary
    workbook a second pass. Terminating on graph identity is also sound
    rather than merely cheap: if rebuilding with what was actually read
    yields the same condensation, every dependency is already in the
    order the run used, and every cell downstream of a value that moved
    re-ran.
7.  **The captured edge set is per-pass, not cumulative** — and that is
    what makes §5.6e's "split" and "lost edges" possible at all. An
    accumulating set can never lose an edge, so a component could never
    split and a reference that stopped pointing somewhere would keep an
    edge nothing reads. An owner the pass skipped keeps what it last
    reported, because "not re-evaluated" and "reads nothing" are
    different facts.
8.  **A component's signature keeps its INTRA-component edges.**
    Filtering them out as "implied by the member list" is wrong for the
    case §5.6e is about: a one-node component is cyclic when the node
    depends on itself and acyclic when it does not, and those are the
    same member list. `INDIRECT` closing a self-reference is precisely
    that, and the filtered signature made the flip invisible to the
    comparison that has to see it.
9.  **A closure is held as ROOTS and re-derived every pass.** Held as
    component ids it freezes at whatever the discovery pass could see,
    and `A1=INDIRECT("C1")` then reorders a component it declines to
    run. Two consequences follow: `graph.plan` gained an `iterating`
    option (M5a1 hardcoded "cycles refuse" because there was no engine
    to hand one to) and a `charge_evals` option, since the engine
    re-plans per pass and charges per *evaluation* — the number §9
    actually bounds for an iterating run is passes times members.
10. **"Unchanged" is only a reason to skip a component that ran.** A
    dynamic edge can pull a component into the closure without changing
    anything about the component itself, and skipping it there leaves
    the cell the new edge points at unevaluated — the one state the
    rebuild existed to fix.
11. **`WorkbookEnv.formulaAt` was reading the merged view, and a second
    pass would have found no formula.** A computed-layer entry is a
    value a run produced; it shadows lower layers for `cellValue`
    because that is what the precedence means, and it carries no `<f>`
    because a run does not author formulas. M5a1 could never observe it
    — it evaluated each cell once and published after — and the first
    iterating fixture converged in two passes on a cell that had
    silently become blank. A body is now looked up across layers, which
    is what "the formula at this coordinate" always meant.
12. **`OFFSET` ships Microsoft's documented contract, and says so.**
    Height and width are documented as positive numbers, so zero and
    negative are `#VALUE!`. Excel 365 has since grown an undocumented
    reverse-extent behaviour for negatives; with the Excel oracle leg
    parked there is no evidence for it here, and implementing a
    behaviour no committed manifest records would be a claim about
    Excel this repo cannot back. `INDIRECT`'s R1C1 request is the
    mirror-image decision: it refuses as a *construct* rather than
    answering `#REF!`, because the text was not malformed — it was
    R1C1, which v1 refuses in written formulas too.
13. **A declared array range is bounded by the GRID, not by
    `max_matrix_cells`.** `A1:D1048576` is 4 194 304 cells against a
    4 000 000 limit, so §5.6c's "zero-filled that shape" seed can be
    handed a shape it cannot build — which made the obvious
    `unreachable` a workbook-reachable crash. It is a §9 refusal
    (`seed_shape_too_large` → `FormulaLimitExceeded`), with the
    one-column-narrower case fixtured beside it so the refusal is the
    limit rather than the construct. The same reasoning turned the
    engine's `assert(max_dynamic_passes >= 1)` into a refusal:
    `WorkLimits.validate` rejects a zero, but the engine is reachable
    without a validated set, and a caller-supplied number must not be
    able to trip an assertion.

---

## 8. Testing & oracles

### 8.1 Fuzz/property targets (wired per PR)

tokenizer (M1a) · parser + limits (M2) · eval no-panic/leak/non-finite (M3a) ·
criteria + PRNG KATs (M3b) · metadata (M4a) · decode/symbols (M4b1) ·
topology + translation (M4b2) · **defined-name attributes + 3D spans
(M4b3)** · draw KATs (M4d) · **per-version index units + casing tables
(M4f)** · SCC + stabilization +
rebuild-reuse (M5a) · patcher confinement + ResolvedSheet round-trip (M5b1) ·
transaction post-failure + calc-state round-trip + refusal purity (M5b2) ·
buffer equivalence (M5c) · determinism + scoped idempotence (M5d) · spill
obstruction/ownership (M7a) · packaging transitions (M7b/c) · numfmt (M8) ·
ABI invalid-input + narrowing + canary-tail (M9a1/2).
`checkAllAllocationFailures` incl. reports.

### 8.2 Oracles (all landed M1b)

Excel-for-Mac AppleScript (isolated Excel instance, target workbook only →
**CalculateFullRebuild** — full calc *with dependency-tree rebuild*;
`CalculateFull` alone does not rebuild dependencies, which matters for
rewritten/dynamic refs → completion wait → save → close → reopen-extract;
recorded calc mode; **two sentinels**: a stale-*value* workbook AND a
stale-*dependency* workbook — both must come back changed or the run is
invalid) → independent frozen extractor (own
ZIP/XML decode; versioned; self-tested) → **semantic manifests** (decoded
typed values; binary64 bits; normalized error spellings; signed-zero policy;
NaN = hard error). Provenance: Excel build, LO build, OS, locale, extractor
version, workbook digest. **Volatile formulas are excluded from every
external value oracle (Excel AND LibreOffice)** — covered by KATs, draw-count
invariants, type/range checks instead. CHAR/CODE 128–255 excluded from Mac
goldens (CP-1252 spec cases). **LibreOffice leg hardened**: pinned invocation + dedicated profile; **hard recalc forced** (calculateAll macro); a **volatile sentinel must come back changed or the LO run is rejected** (else a save merely echoes stale caches). Screened corpus (screen out + count:
manual calcMode, fullCalcOnLoad, fullPrecision=0, external links, volatiles,
unknown provenance) = consistency signal. Hand-derived spec suite anchors
divergence points. Precedence is **fidelity-specific**: `.excel` mode — Excel > hand-spec >
corpus > LO; `.ieee` mode — the normative `ieee_fp_rules_v1` table +
hand-derived bit goldens lead, with Excel retained only as a recorded
divergence witness. Conflicts recorded, never averaged. Regeneration = reviewed command + manifest diff.

**M1b decisions (shipped 2026-08-03).** The harness lives in
`tests/oracle/` (replay, pure Zig, no application needed) and
`scripts/oracle/` (recording: Python, AppleScript, shell). Fourteen
points the row left open or got wrong:

1. **Independence is structural.** The oracle imports nothing from
   `zlsx` — it carries its own ZIP reader (`tests/oracle/zip_reader.zig`,
   own central-directory walk + `std.compress.flate`) and its own XML
   scanner (`xml_scan.zig`). An oracle that read workbooks through
   `pkg/zip.zig` and `pkg/sheet_xml.zig` would have the same bug on both
   sides of every comparison and would confirm zlsx against itself.
   `scripts/oracle/build_inputs.py` hand-authors the input XML for the
   same reason — plus the plain fact that no sane writer will emit a
   cached value that contradicts its own formula, which is exactly what
   a sentinel is.
2. **Recording and replay are separate programs.** Recording needs
   macOS, Excel and LibreOffice; replay needs a committed JSON file.
   `zig build test-oracle` therefore gates CI on a Linux box with no
   spreadsheet application installed — evidence gathered once keeps
   gating everywhere.
3. **The Excel leg is PARKED — it needs a human.** Excel 16.111.2
   answers AppleScript property queries (`version`, `name of every
   workbook`, `calculate full rebuild`) but every window operation is a
   no-op or refused: `open workbook`, `open -a`, `close … saving no`,
   and `quit` (which returns "User canceled", −128). That signature is a
   modal dialog nobody has dismissed. It cannot be diagnosed from here
   either: `osascript` lacks assistive access (−1728), so the dialog
   cannot even be read. **Ask:** bring Excel to the front, dismiss
   whatever is on screen, close any open workbook, then re-run
   `scripts/oracle/regenerate.sh`. The driver
   (`scripts/oracle/record_excel_mac.sh`) is complete and exits 3 with
   this instruction; the LO and hand-spec legs shipped without it.
4. **Excel silently refuses cells that are out of column order within a
   row.** Rows ascending and cells ascending *within* each row are both
   mandatory; violate either and Excel declines to open the file with no
   error an automated caller can see. zlsx and LibreOffice both accept
   the unsorted form, so nothing else in the tree reveals this.
5. **Excel is sandboxed.** A scripted open of a path outside its
   container fails silently and reports success. Inputs are staged into
   `~/Library/Containers/com.microsoft.Excel/Data/Documents/` first.
6. **LibreOffice: replace only `Standard/Module1.xba`.** Writing
   `user/basic/script.xlc` or `Standard/script.xlb` de-registers the
   Basic library; the macro then does not exist and `soffice` reports
   that by exiting 0 having done nothing. The profile must be
   materialised by one `--terminate_after_init` run before the macro is
   installed.
7. **The stale-dependency sentinel's discriminating power is
   UNVERIFIED.** It is built as specified — a three-deep chain with a
   deliberately inverted `calcChain` — but whether `CalculateFull` and
   `CalculateFullRebuild` actually differ on it could not be measured
   with the Excel leg parked. It functions today as a second,
   independent stale-value proof on a different formula shape. Measure
   it when the Excel leg runs; do not cite the discrimination until then.
8. **Sentinels are checked against the raw extraction, never a
   manifest.** A manifest excludes volatile cells by design, so checking
   the volatile sentinel there would pass by absence — the exact failure
   it guards against.
9. **An empty sentinel set cannot vacuously accept a run**
   (`sentinel.hasProof`), and neither can a set of volatile draws alone:
   Excel redraws volatiles on load without a full calculation.
10. **`hand_spec` has no workbook**, so its provenance digest pins the
    *input* workbook its values answer, and its `app_build` names the
    authority the values were derived from. Cases whose answer cannot be
    defended from a documented rule are deliberately ABSENT — a
    hand-spec entry that is really a recollection of Excel outranks the
    corpus and LibreOffice under §8.2 precedence.
11. **The corpus admits zero workbooks today**, all 28 screened out for
    unknown provenance. That is the correct answer, and it is why the
    corpus is a witness in both fidelities rather than an authority.
12. **Application-leg digests change on every run** — Excel and
    LibreOffice embed timestamps on save — so `regenerate.sh` classifies
    a digest-only diff explicitly. Without that, reviewers learn to skim
    past a diff that is usually noise and one day is not.
13. **`std.json.parseFromSlice` needs `allocate = .alloc_always`.** The
    default borrows escape-free strings from the input buffer, so a
    manifest outlives its JSON bytes as a struct full of dangling
    slices that still compare and print as though fine.
14. **First recorded divergences** (LibreOffice 26.2.5.2 vs the
    hand-derived suite), pinned as named tests in `replay.zig`:
    `SQRT(-1)` → `#VALUE!` in LO where the spec says `#NUM!`; `0.1+0.2`
    → `0x3FD3333333333333` in LO where IEEE gives `…334`; `1/3` one ULP
    low; and consequently `(0.1+0.2)=0.3` → TRUE. LO **preserved**
    signed zero (`-0` → `0x8000000000000000`). Ten other hand-derived
    `.excel` cases — operator precedence, error taxonomy, coercion,
    percent, overflow-to-`#NUM!` — were confirmed independently.

### 8.3 Comparison rule

Bit-exact parsed binary64; per-case documented tolerances only; decoded
text/bool; normalized-then-exact errors.

### 8.4 Fidelity levels

No-formula no-op → whole-file identity via the **pre-mutation fast path
(ordered)**: namespace preflight → **mutation-journal check** (staged
`cell_deltas`/`row_appends`, `workbook.zig:2553-2637`, disqualify identity —
recalc no-ops but save proceeds normally; staged edits never discarded;
byte tests both cases) → **carrier census** → **identity when journal empty
AND census zero — signed workbooks included** (the permitted signed no-op);
actual mutation on a signed workbook refuses. Prefixed `<x:f>` refuses;
table-formula-without-cell-`<f>` refuses at zero cells. Real recalc → exact overridden-part
set; raw LFH/payload identity for untouched entries (`store.zig:695`);
approved-node confinement inside changed parts. Mode-differential fuzz;
metamorphic edit-then-recalc; refusal purity (staged-edit baseline).

---

## 9. Performance & limits

Lanes: hyperfine ReleaseFast (`bench_ci.sh:34`); RSS ReleaseSafe
(`build.zig:897`); whole-graph single-mode builds asserted (fixes the
`build.zig:1030` trap). Report-only CI; `compare_bench.py` stats.
**Workloads are committed artifacts** (M5d bench PR): a checked-in generator
script producing fixed-topology workbooks (dependency fan-in/out recorded,
formula mix per the F1 profile, shared-string density, archive size), pinned
warm-up policy, N repetitions, gating statistic (median; p95 reported), and
per-target baselines; measured phases (model / evaluate / stage /
prepare+swap / serialize) reported separately, ZIP+XML time included and
labeled. **250 ms / 100k cells aspirational until the M5d baseline**
(thresholds frozen after). Parse ≥100 MB/s target; 10k-vs-100k ≤15×;
full-column sparse-range benches (§5.6a). **Milestone-local perf gates after
M5d** (a tree walker can pass M5d and regress later): M7a adds large-spill AND SORT/UNIQUE benches (they ship there); M7b2 adds
whole-column criteria (`SUMIF(A:A,…,B:B)`) and multi-criteria benches; M8
adds a TEXT-heavy bench; M9d adds a mixed full-registry workload
— each report-only against its own recorded baseline, same methodology. **Release gating**: once the M5d baseline is frozen, the local `compare_bench` regression check becomes **blocking for release cuts** (CI stays report-only), and v1 carries **absolute ceilings on ONE named workload — the 100k-cell F1-mix fixture, digest recorded at M5d3**: evaluate ≤ 500 ms and end-to-end ≤ 1 s in **ReleaseFast** (**warm-cache; N=20 runs matching `bench_ci.sh:50`; comparison via `compare_bench.py` EXTENDED in M5d3 with a `--gate` mode** — median-based, nonzero exit on regression for release cuts (today it compares means and always exits 0, `compare_bench.py:88,146`); CI keeps report-only mode; both exit behaviors tested; distribution + baseline commit reported; thermal/load controls); peak RSS ≤ 3× **model bytes = sum of the fixture's decompressed part bytes**, **baseline-adjusted** (pre-open process RSS subtracted), in the **ReleaseSafe RSS lane**; FIRST recalc, zero retained generations; host recorded at baseline. Owner waiver required to exceed; M9d runs absolute + regression checks.

**Limits (each named, typed refusal, boundary-tested; units explicit)**:

| Limit | Value | Unit |
|---|---|---|
| `max_formula_chars` | 8 192 | Unicode code points (Excel-aligned; algorithm oracle-pinned) |
| `max_formula_utf8_bytes` | 32 768 | **conservative zlsx safety cap on UTF-8 source bytes** — Excel's "16 384 bytes" is its internal *compiled* representation, which zlsx does not model; no equivalence is claimed |
| `max_tokens` | 16 384 | tokens |
| `max_ast_nodes` | 16 384 | nodes |
| `max_parse_depth` | 256 | nesting of any grammar production |
| `max_fn_nesting` | 64 | function-call depth (Excel's documented limit) |
| `max_args` | 255 | per call |
| `max_operand_stack` | 1 024 | evaluator stack slots |
| cell-text 32 767 cp | **Excel-domain rule, NOT a refusal**: producing longer text yields `#VALUE!` (Excel's own REPT behavior — a successful error value) | oracle-pinned |
| `max_text_bytes_safety` | 1 MiB | per-string resource cap (typed refusal; distinct from the Excel-domain rule) |
| `max_matrix_cells` | 4 M | elements |
| `max_eval_depth` | 512 | dependency-closure recursion (cell → cell; M5a) |
| `max_expr_depth` | 1024 | **expression-tree walk** (AST nodes on the stack; M3a2, `eval.Options`). Distinct from both `max_parse_depth` (recursing grammar productions) and `max_eval_depth` above. Left-associative operator chains are folded iteratively, so in practice only parenthesis nesting reaches it |
| `max_output_archive_bytes` | **2³²−1 bytes exactly** (matches the ZIP32 sentinel bounds `pkg/store.zig:720-742`) | serialized output archives — `saveToOwnedBuffer`, `save_to_buffer`, `saveWithRecalc`; identical typed outcome at every layer |
| workbook materialization | **`max_workbook_compressed_bytes` 1 GiB; `max_workbook_decompressed_bytes` 4 GiB; `max_modeled_cells` 64M** — PartStore allocations (own arena, 512 MiB/part, `store.zig:105-129,1340-1384`) sit outside `max_run_arena_bytes`; early-refusal tests | pre-model refusal |
| retained generations | `max_retained_generations` 4; **`max_retained_generation_bytes` 2 GiB; `max_retained_fds` 16** — in resolved limits + fingerprints; projected retention **preflighted before allocating or swapping** | pre-swap refusal |
| aggregates — **bytes** (counted allocator) | `max_run_arena_bytes` **1 GiB**, live matrix cells 8M, string payload 256 MiB, retained ASTs 128 MiB, diagnostics 1 MiB — defaults; hard maxima 4× each; caller-adjustable via `ResourceLimits` (M3b, `src/formula/run_inputs.zig` — a **separate struct from `parser.Limits`**, which bounds parse shape); resolved values echoed + fingerprinted. An exhausted category is `FormulaLimitExceeded`, never a bare `OutOfMemory`: the budget records which one tripped. `matrix_cells` is charged as a **count**, so the limit does not depend on `@sizeOf(ScalarValue)` | byte accounting; below/at/above per category |
| aggregates — **work** (explicit checked counters; can burn CPU without allocating) | `max_total_cell_evals` **50M**, dependency edges 50M, `max_scc_iterations` = **caller RESOURCE ceiling only, default 32 767 (hard max 32 767)** — never conflated with the workbook's semantic `calcPr@iterateCount`: hitting `iterateCount` = success + non-converged, hitting a lower caller ceiling = `FormulaLimitExceeded` + zero mutation (§5.6c), `max_dynamic_passes` default 3 (caller-adjustable, hard max 10), sort/comparison ops 500M — defaults; hard maxima 4× unless stated; caller-adjustable **in Zig/C only — CLI and Python fix limits at defaults in v1 (declared, no flags)**; resolved values echoed + fingerprinted. **Three shapes of bound, classified (M5a2 `WorkCategory.kind`)**: a *total* only grows (`dependency_edges`, `total_cell_evals`), a *depth* unwinds (`eval_depth`), and a **per-scope** bound is re-counted from zero in every scope it governs (`scc_iterations`, `dynamic_passes`) — §5.6c gives each SCC its own pass counter, so accumulating passes into one running total would refuse a workbook whose every component iterated legally; `WorkCounters.charge` rejects a per-scope category, and the engine that owns the scope reads the limit instead | decrement sites named per counter; below/at/above boundary tests |

---

## 10. Refusal & error taxonomy

Plane 1: producible classic errors + preserved rich errors; **Excel error
values are successful results** (CLI exit 0). Plane 2 (merged into
`pkg/workbook.Error`; diag sink):

| Error | Fires |
|---|---|
| `FormulaUnsupportedFunction` | unregistered call |
| `FormulaUnsupportedConstruct` | **unsupported SEMANTICS only** (mark-eligible): external ref; R1C1; LAMBDA/LET/`_xlpm.`; ambiguous dialect; relative defined name; table formula w/o cell `<f>` (`aca` honored — non-refusal regression pinned) |
| `FormulaPrecisionAsDisplayed` | `fullPrecision="0"` — **always refuses, never mark-eligible** (keep-stale regression pins it) |
| `FormulaMalformedInput` | **malformed/unsafe input, NEVER mark-eligible**: malformed shared topology; namespace-prefixed parts; `bx="true"`/unknown `<f>` attrs; unparseable input |
| `FormulaLocaleSensitiveInput` | outside a pinned locale grammar |
| `FormulaDataTableUnsupported` | data-table formulas (`dt*`/`r1`/`r2`/`del*` incl.) |
| `FormulaSignedWorkbook` | signature parts present at model build — mutation invalidates signatures (embedding-arc precedent `embeddings-in-xlsx.md:646,847-855`); identity no-ops permitted; strip = M10+ |
| `FormulaStaleEmbeddings` | **any** coverage overlapping final staged replacements — independent of `include_formulas` (heads, caches, spill tails, ordinary values, appended rows; v1 refuse-only) |
| `FormulaAnchorRequired` | site-less site-dependent eval |
| `FormulaCycle` | cycle, iteration off |
| `FormulaDynamicRefUnstable` | fixpoint exhausted |
| `FormulaSpillPersistUnsupported` | outside the approved mutation set |
| `FormulaResultNotRepresentable` | top-level multi-area result (v1) |
| `FormulaLimitExceeded` | any §9 limit |

---

## 11. Interactions

`<xm:f>` route-through (post-M2, scheduled M10+); shared-strip hazard fixed
M5b/M5d; typed-overlay gaps M4b2; ResolvedSheet supersedes delta-write;
`zlsx_recalc` + `openBuffer` + `saveToOwnedBuffer` are general infrastructure
(also discharge the Databricks writer-half follow-on); literal-masking
correction M-1; calcChain plans superseded on the recalc path only;
embedding-arc names inert; rewriter unchanged for previously-recognized
constructs.

---

## 12. API surface

### 12.1 Zig

```zig
Workbook.recalculate(alloc, io: std.Io, run: RunInputs, opts, diag) RecalcError!RecalcReport
Workbook.saveWithRecalc(alloc, io, path, save_opts, run, opts, diag) SaveError!RecalcReport
Workbook.openBuffer(alloc, io: std.Io, bytes: []const u8) OpenError!Workbook  // borrow ends at return
Workbook.openBufferControlled(alloc, io, bytes, ctl: Control) OpenError!Workbook   // cancel/deadline-aware (§5.10)
Workbook.evaluate(alloc, current_sheet, site: ?EvalSite, formula, run, diag) EvalError!EvaluatedValue
Workbook.markRecalcOnLoad() Error!void
Writer.saveToOwnedBuffer(alloc, io: std.Io) SaveError![]u8                     // byte≡path-save; capped by max_output_archive_bytes
Writer.saveToOwnedBufferControlled(alloc, io, ctl: Control) SaveError![]u8     // identical bytes; polls the M5d1 seams
zlsx_recalc.writerSaveWithRecalc(alloc, io, writer, path, run, opts, diag) SaveError!RecalcReport  // threads Control into BOTH pre-recalc stages
```

CalcState workbook-derived; allocator-explicit; `deinit` everywhere; diag
pre-error.

### 12.2 CLI

Command syntax:

```
zlsx eval   <file.xlsx> --formula "<formula>" (--sheet N | --name NAME) [--anchor A1]
            [--dialect da|legacy] [--now ISO8601] [--utc-offset MIN] [--seed N]
            [--mode excel|ieee] [--profile windows_1252] [--deadline SECONDS]
zlsx recalc <file.xlsx> --out <out.xlsx> [--now ISO8601] [--utc-offset MIN]
            [--seed N] [--mode excel|ieee] [--profile …]
            [--on-unsupported refuse|keep-stale-and-mark] [--report]
            [--deadline SECONDS]
```

`--formula` carries the text (the shipped option parser rejects
leading-dash positionals, `cli.zig:613-645` — `--formula` makes `-A1`,
`--A1`, and flag-shaped text unambiguous; contract-tested). `--anchor`
supplies the `EvalSite` (validated A1; missing on a site-dependent formula →
exit 3 `FormulaAnchorRequired`); it and every resolved input appear in the
reproducibility metadata. NDJSON is a **stream
state machine** (versioned records; grammar normative — refusals and
cancellation can occur **before** any header exists):

```
eval stream   := refusal | cancelled
               | eval-header ( eval-cell* ) diagnostic* ( eval-complete | refusal | cancelled )
recalc stream := diagnostic* ( recalc-report | refusal | cancelled )   (only with --report)
```

Terminal states are explicit: success = `eval-complete`/`recalc-report`;
typed refusal = `refusal`; cancellation = a **`cancelled` record**
(`{"kind":"cancelled","v":1,"after":"…"}`) — **SIGPIPE exception**: a closed pipe cannot receive a terminal record — per the shipped contract (`cli.zig:1651-1666,1721-1729`, `docs/cli.md:224-227`) SIGPIPE ends the stream prefix-valid, no terminal, exit 0. Any OTHER EOF without a terminal = abnormal termination, distinguishable.

```jsonl
{"kind":"eval-header","v":1,"type":"number|text|bool|error|matrix","value":…,"rows":…,"cols":…,"resolved":{"now":"…","utcOffsetMin":0,"seed":"7","mode":"excel","profile":"windows_1252","dialect":"da","anchor":"B7"}}
{"kind":"eval-cell","v":1,"r":1,"c":1,"type":"number","value":1}
{"kind":"diagnostic","v":1,"severity":"note","message":"…"}
{"kind":"eval-complete","v":1,"cells":4}
{"kind":"refusal","v":1,"error":"FormulaCycle","cells":["Sheet1!A1"],"truncated":false}
{"kind":"recalc-report","v":1,"sheets":[…],"passes":1,"nonConverged":0,"dynamicPasses":1,"census":[],"resolved":{…}}
```

**`blank` is INTERNAL-ONLY (supersedes round-8)**: public results never
contain it — standalone eval coerces exactly as recalc publication (`=A1`
on empty A1 → 0 in BOTH APIs; fixtures: `=A1`, arithmetic, `&`, comparison,
COUNT, ISBLANK). NDJSON/C/Python `blank` variants removed. **`seed` serializes as a decimal string** (u64 exceeds
JSON/JS safe-integer range).

Scalar results carry `value` in the header and emit no `eval-cell` records;
matrix cells are row-major; completed records stand (prefix-valid) when a
`cancelled` terminal closes the stream (exit 130/143 or 3 for `--deadline`). Empty matrices cannot
occur (§5.3a — normalized to `#CALC!`). Exit codes:

| Code | `eval` | `recalc` |
|---|---|---|
| 0 | evaluated (incl. Excel error values) | recalced + written |
| 1 | usage | usage |
| 2 | open/parse | open/parse |
| 3 | typed refusal (**incl. `FormulaLimitExceeded`** — limits are refusals at every layer) / `--deadline` | refusal / cancellation — **destination untouched** (§5.7.9; "no file" only when initially absent) |
| 4 | allocation failure (genuine OOM only) | allocation failure |
| 5 | stdout write failure | output write/rename failure |
| 6 | **default-context acquisition failure** — the `std.Io` wall clock or random source could not be read while resolving an omitted `--now`/`--seed` (never reported as OOM; Python raises `ZlsxContextUnavailable`, a `ZlsxError` subclass; no C status is needed because the ABI has no defaults to resolve, §5.5) | same |
| 130/143 | SIGINT/SIGTERM (`docs/cli.md:268`) | same |

**Cancellation scope**: the no-output guarantee covers **workbook-file
mutation only** (recalc never leaves a partial file; the final poll sits
immediately before rename and the rename-through-swap span is
non-cancellable — §5.7.9). **Exit mapping is commit-aware**: a signal
arriving after rename reports success (0), never 130/143 — the current
override-at-exit behavior (`cli.zig:1718-1729`) is corrected for these
commands. `eval`'s NDJSON stream may be **prefix-valid** on cancellation
(§12.2 grammar) — matching shipped flush-on-signal behavior
(`docs/cli.md:224-227`). `--sheet N | --name NAME` mandatory for `eval`;
`--out` identity-checked.

### 12.3 C ABI + Python (M9a1/M9a2)

**`zlsx_status_v1` (frozen — NEW exports only; legacy exports keep shipped `0/-1`, `c_abi.zig:1927-1953`; mapping documented; old-binding-vs-new-library compat test)**: `0` OK · `-1` generic error
(message in the caller-supplied error buffer — existing convention,
`c_abi.zig:1927`) · `-2` typed refusal (structured diag populated — **this
includes `FormulaLimitExceeded`: limits are Plane-2 refusals at every layer**,
CLI exit 3, Python `ZlsxFormulaRefusal`; one normative outcome→layer mapping,
no per-layer reclassification) · `-3` allocation failure (genuine OOM only —
CLI exit 4) · `-5` cancelled (observed pre-commit). Every fallible
export takes `(char* errbuf, size_t errbuf_len)` AND (where applicable)
`zlsx_diag_v1*`. **`struct_size` rules**: readers use
`min(caller_size, known_size)`; writers touch only that prefix — **bytes
beyond the known prefix are never written** (canary-tail tests); minimum
accepted size = offset+size of the last v1 field (rejected below that with
`-1` + message); output params zero-initialized within the known prefix on
entry. Value descriptor: dims + element descriptors
`{uint8_t tag; double num; uint64_t payload_off, payload_len;}` + one payload
arena; references dereferenced before crossing; per-type release fns
(`zlsx_value_release`, `zlsx_recalc_report_release`, `zlsx_diag_release`,
`zlsx_buffer_release`); nullable pointer/length pairs documented per field;
feature probes per export (absent export = feature off). Field-width table +
boundary tests (negative, `UINT32_MAX`, `UINT64_MAX`, `INT32_MAX`,
`SIZE_MAX`, one-past-limit). Exact field layouts land in the **committed M9a1
design note** (a repo artifact, not a plan revision), reviewed before code.

**Cancellation across the ABI**:
`int32_t zlsx_cancel_token_new(zlsx_cancel_token_t** out, char* errbuf, size_t)`
(status-style like every fallible export — allocation can fail) /
`zlsx_cancel_token_trigger()` (thread-safe, any thread) /
`zlsx_cancel_token_free()`; fallible exports accept an optional token (NULL =
non-cancellable). **Semantics: cancellation is "observed before commit"** —
a trigger observed at a poll point before the commit point returns `-5` with
memory untouched; a trigger arriving after the commit point (§5.7.9) is
ignored for status: the call returns success with the report's
`committed=true` (+`cancelled_late` flag). C, Python, and CLI expose this
identically. **Polling bound (enforceable)**: the engine polls at least once
per 1 024 cell evaluations and per 65 536 matrix-element/sort/serialize work
units; §8 gains latency tests (pre-start, mid-SCC, mid-sort, mid-lookup,
mid-serialization, commit-race) with a deliberately long evaluation. Token
lifetime: caller-owned, outlives the call. **Build-mode reality (corrected)**: single-threaded is an *option*
(`build.zig:11`, forwarded at `:836-840`) and the `-Dsingle-threaded=true`
lane is CI-compiled — under it Zig lowers atomics to plain ops, so a
cross-thread token would be silently broken in exactly the supported
configuration. **M9a1 hard-sets the C ABI module to multi-threaded
regardless of the option** (compile-time assertion; the option narrows to
CLI-only; build help, `AGENTS.md`, and CI commentary updated); the CLI
binary uses the signal-safe flag kind of `CancelToken` (§5.5). Both
configurations tested.

**Python API (normative)**:

| Method | Signature | Returns / raises |
|---|---|---|
| `Editor.recalculate` | `(now=None, utc_offset_min=0, seed=None, mode="excel", profile="windows_1252", on_unsupported="refuse", timeout=None)` | `RecalcReport` dataclass; raises `ZlsxFormulaRefusal(error_name, cells, census)` on Plane-2; `TimeoutError` **only when cancellation is observed before commit** (post-commit trigger → normal return, `report.cancelled_late=True`) |
| `Editor.save_with_recalc` | `(path, now=None, …same ctx…, on_unsupported="refuse", timeout=None)` — the **atomic §5.7.9 transaction** via `zlsx_editor_save_with_recalc` (failure-injection tested) | `RecalcReport`; pre-commit failure ⇒ destination keeps prior bytes (or stays absent) + memory untouched; **`TimeoutError`/`KeyboardInterrupt` per shared mechanics — pre-commit only; post-commit ⇒ normal return + `cancelled_late`** |
| `Editor.evaluate` | `(formula, sheet, anchor=None, dialect="da", now=None, utc_offset_min=0, seed=None, mode="excel", profile="windows_1252", timeout=None)` | **`EvalResult`** dataclass. `.value` is a scalar (`float/str/bool` — blank never escapes, §12.2), an **`ExcelError(str)`** value wrapper for Excel error values (a *result*, not an exception — `ZlsxError` remains the binding's exception type, `__init__.py:195`), or a `Matrix` (`.rows`, `.cols`, list-of-lists; cells are float/str/bool/`ExcelError` — blanks publish as 0, §12.2). **`.resolved` echoes the exact resolved context** (now, utc_offset_min, seed, mode, profile, dialect, anchor) — the §5.5 echo requirement reaches Python *results*, not only reports, so a defaulted volatile evaluation is reproducible (test: two defaulted `evaluate()` calls differ; replaying either `.resolved` reproduces it exactly). Raises `ZlsxFormulaRefusal`; **`TimeoutError`/`KeyboardInterrupt` per shared mechanics (eval never commits)** |
| `Editor.save_to_buffer` | `()` | `bytes` |
| `Editor.mark_recalc_on_load` | `()` | `None` |
| `Writer.save` | `(path, recalculate=None)` — `recalculate=RecalcOptions(...)` composes via the orchestrator. **Save-generation tracking lives in CORE `Writer` (Zig)**: the retained artifact is keyed by **(writer_generation, EffectiveRunInputs, recalc options, engine fingerprint)** — a same-generation save with a DIFFERENT context re-recalculates, never reusing stale volatile caches (same-gen/different-context tests); every `Writer`/`SheetWriter` mutator invalidates; **all save paths route through it** — Zig path/buffer saves (`writer.zig:496`), the existing C export (`c_abi.zig:1927`), new exports, Python `save()`/auto-save (`__init__.py:2686-2697,2758`). Save-then-save, C-save-after-recalc, explicit-save-inside-`with`, mutate-then-save tests | `RecalcReport` or `None` |
| `SheetWriter.write_row_with_formulas` | existing signature (`__init__.py:1514`) + per-cell formula descriptors — a cell may be text (scalar) or `FormulaSpec(text, dialect, ref=None)`; a row-wide `dialect=` kwarg remains as sugar for uniform rows. **`.cse(ref)` legal ONLY on the anchor** — one `<f t="array" ref>` on the top-left, members carry `<v>`/no `<f>` (`xlsx.zig:2138-2142`); authored members are empty placeholders; the **CSE-rectangle state machine** refuses missing/duplicated/overlapping/formula-bearing members before finalization (the C v2 export mirrors this: per-cell descriptor {text ptr+len, dialect tag, optional ref ptr+len}, not parallel text arrays — the shipped per-cell text-array shape at `c_abi.zig:1850-1859` cannot encode a rectangle). Older-dylib probe keeps the legacy path | as today |

**Cancellation mechanics**: cancellable FFI calls run on a **worker thread**;
the main thread waits interruptibly and triggers the token on
`KeyboardInterrupt`/timeout (a Python signal handler cannot run while the
main thread is blocked inside a synchronous `ctypes` call); cancellation
latency is tested with a deliberately long evaluation.

`_ffi.py` feature-probes **every staged symbol group independently** —
`_HAS_RECALC` (recalculate), `_HAS_EVAL` (evaluate), `_HAS_SAVE_BUFFER`
(save_to_buffer/open_buffer), `_HAS_SAVE_WITH_RECALC` (save_with_recalc),
`_HAS_CANCEL` (token API), `_HAS_WRITER_RECALC` (writer save-with-recalc),
`_HAS_FORMULAS_V2` (dialect writer export) — each probe gates exactly the
Python methods named here, with the shipped older-dylib pattern; `test_basic.py`
skip-guards mirror it. Every call wraps release fns in `try/finally`
(tested); `Book` stays read-only.

### 12.4 Spark (batch-only)

`zlsx.recalc="true"` **activates**; source files are **never mutated**
(recalc is in-memory per open). Driver: open → recalc (resolved context) →
buffer → reopen → **schema inference and planning on the recalced snapshot**.
Partition inputs carry (path, **source digest**, resolved context, engine
version); executors **read source bytes ONCE, SHA-256 that buffer, compare to the
partition digest, and `open_buffer` THE SAME buffer** — verification and
parsing share one immutable snapshot (no verify-then-reopen window; driver
identical); refuse on drift; retries re-derive identically (tests incl. a
mutate-between-verify-and-open race fixture). Same wheel + same context + same bytes ⇒ identical results;
mixed-version fleets refused — "engine identity" is a **probed
fingerprint export** (`zlsx_engine_fingerprint()`: semantic version +
`excel_fp_rules`/`rng`/`collation` rule versions + target triple + build
hash; the bare package version at `c_abi.zig:51-54` is insufficient, and
`ZLSX_LIBRARY` can load arbitrary binaries — fingerprint mismatch refuses). The per-executor recalc cache is keyed by
**(workbook digest, every resolved run input — now, offset, seed, mode,
profile, limits — engine/build version, recalc options)**, never by digest
alone (two partitions with different contexts must never share a snapshot),
and is **byte-bounded**: `zlsx.recalcCacheMaxBytes` (default 512 MiB, `0` =
off) with LRU eviction and a per-entry cap of the workbook's buffer size —
long-running executors cannot retain unbounded archives; documented
optimization, not a correctness dependency. Spark's
default `zlsx.recalcUtcOffsetMin` is **0 (UTC)** like every other layer —
the driver's zone applies only via the explicit option, recorded in the
resolved context. Streaming + recalc refused at option validation.
Serverless-verified before documenting.

---

## 13. Documentation flips

M-1 ships this checklist (extended by an rg sweep at M-1; each entry
classified flip-at / historical-label):

| Doc | Claim | Action |
|---|---|---|
| `docs/plans/post-0.2.9-roadmap.md:144,203,208-211,487-488,596-607` | D1 deferred / exclusions | **M-1** flip |
| `goal_plan.md` D1 row + `:214-219` "Deliberately not on this list" (+ site rebuild per `goal_plan.md:3-7` at M-1 and M9d) | deferred | **M-1** flip |
| `goal_evol.md:48` scope line · `goal.md:219-222` follow-up 4 | out of scope / standalone | **M-1** flip |
| `docs/plans/formula-literal-masking.md:42-48` | misstates `<v>` preservation | **M-1** correction |
| `docs/plans/editor-rebase.md:359` · `docs/plans/workbook-overlay.md:287` · `docs/plans/writer-rebase.md:575` · `docs/plans/structural-edits.md:121` ("Excel will recompute") | no-evaluation statements | **M-1** historical labels |
| `README.md:452-455` | "Out (by design) … never computes" | **M5d/M6** flip |
| `docs/cli.md` | no eval/recalc | **M6** new sections |
| `docs/jq-for-excel.md:290-292` | "reader, not a spreadsheet engine" | **M6** historical label |
| `docs/vs_calamine.md:64,90` | claims zlsx skips `<f>` / cannot emit formulas — **already false** (`xlsx.zig:2070-2153`, `writer.zig:940-955`) | **M-1** correction |
| `docs/vs_calamine.md:5,130` | "no formula evaluation" (true until M5d2) | **M6** historical label |
| `docs/xlsx_test_corpus.md:27,56` | "don't need to evaluate" | **M6** flip |
| `docs/package-layer.md` | layer description | **M5d** update |
| `bindings/python/README.md:252` ("never") + **full new-API docs**: methods, `Matrix`, `ExcelError`, refusal/cancellation semantics | **M9a2** gate (docs land with the code PR) |
| `bindings/python/README.md:177-179` ("all batch options apply to streaming" — false once recalc refuses streaming) + **Spark option table (batch-only)** | **M9b** gate |
| `src/xlsx.zig:1-13` · `src/cli.zig:1,1141` · `pkg/workbook.zig:366,5477-5479` | in-source scope comments (incl. the "future evaluator (Tier D1)" promise at the emitCell branch) | with the code that changes them (**M5d/M6**) |
| ~~`src/formula/tokenizer.zig:566-575`~~ (scope note made false by the new token kinds) | tokenizer scope comment | **M1a — done**: module doc rewritten with the tokens; `rewriter.zig`'s matching "classifies these as `.unknown`" claim flipped too |
| `build.zig:891-893` ("zlsx and zlsx_pkg cannot coexist" — contradicted by `zlsx_recalc`) | module-graph comment | **M5c**; `build.zig` joins the release rg scan |
| `src/writer.zig:1645-1647` (claims the reader does not expose formula text — already false, `src/xlsx.zig:2070-2135`) | stale reader claim | **M-1** historical correction; sweep regex extended to read-side formula-text claims |
| `AGENTS.md` | add formula conventions + harness how-to | **M4c** |

**Release gate (M9d)**: `rg -in "formula evaluation|never computes|not
evaluated|out of scope|spreadsheet engine|auto-recalculate|don't need to
evaluate|recalcs on open|evaluator stays deferred|evaluat|interpret|recomput|recalculat"
README.md docs/ bindings/ goal*.md src/ include/ pkg/ build.zig` — every hit on a
**reviewed allowlist** (flips done, historical labels, this plan) or the
release blocks.

---

## 14. Risks

1. Excel folklore semantics → M1b harness-first + sentinel; bit-exact;
   per-rule fixtures.
2. Editor/PartStore/Writer/module-graph integration → named gates
   (ResolvedSheet, prepare/swap + pre-swap reports, §5.7.9 save ordering,
   `zlsx_recalc` consumer test, calcChain resolution, `saveToOwnedBuffer`
   equivalence).
3. Scope weight → 41 gated PRs; v1 = M9d; hard refusal gates.
4. DA/CSE persistence → approved mutation set + empirically-pinned cm/vm
   collections + proofs; explicit authoring dialect.
5. Dynamic/volatile oscillation → callsite-keyed memoized draws + bounded
   fixpoints + KATs.
6. Cross-target math variance → same-build/target contract; CI goldens;
   in-tree math on divergence.
7. Iterative statefulness → multi-SCC schedule + seed table + scoped
   idempotence + parity fixture.
8. Oracle platform skew → profiles; exclusions covered by spec cases;
   volatiles KAT-only everywhere external.
9. Performance → repo methodology; aspirational-until-baseline; sparse-range
   fast paths; escape hatches in-boundary.
10. ABI freeze → status contract inline; min-size rule; canary-tail tests;
    design-note-first.

---

## 15. Review log — Codex 5.6 challenge rounds

> Convergence gate: rounds continue until **ship-ready / 0 findings above
> NIT**. Dispositioned findings are not re-raised unless a fix is defective
> (defect must be cited).

| Round | Date | Findings (B/M/m) | Disposition |
|---|---|---|---|
| 1 | 2026-08-02 | 12/25/6 | All 43 → v2: precision model rebuilt; shared translation = new work; transactional commit + patcher; spill decision table; refusals over `#NAME?`; editor handle; Spark driver-side; v1 redefined; oracle re-rank; deterministic limits |
| 2 | 2026-08-02 | 7/21/5 | All 33 → v3: logical view; prepare/swap; EvalSite; fullPrecision; 15-digit ingress; finiteness; dialect classification; manifest goldens; ABI value descriptor; Spark `save_to_buffer` path |
| 3 | 2026-08-02 | 8/19/4 | All 31 → v4: ResolvedSheet; A1-never-shifts; decode boundary; saveWithRecalc; relative-name refusal; scoped idempotence; metadata reader; EvalEnv; ScalarValue; spill ownership; independent extractor; Spark batch-only |
| 4 | 2026-08-02 | 5/13/3 | All 21 → v5: Writer surface; std.Io; RunInputs/CalcState split; complete calc-state incl. sheetCalcPr; attribute-based slaves; executor recalc; M4a/M4b order; CSE at M7b; platform profile; coordinate bases; four-oracle M1b; RNG KATs; honest counts; exit tables |
| 5 | 2026-08-02 | 10/23/2 | All 35 → v6: namespace preflight; `zlsx_recalc` orchestrator + openBuffer; calcChain rel resolution; approved mutation set; CV text-compat; FormulaWrite; decoded symbols; locale refusal; utc offset; SCC seeds; draw schedule; ordered iteration; array constants; ST_Xstring; reports pre-swap; raw-entry identity; sentinel automation; status contract; NDJSON; digests; 29 PRs |
| 6 | 2026-08-02 | **6 B / 21 M / 1 m** | All 28 → v7 (this revision; **self-containment made structural — nothing normative references a prior revision**). Defective-fix corrections: **CSE rules restated in declared-range-vs-result terms** (truncate when D<R; `#N/A` fill when D>R; per-dimension broadcast) (B1); **volatile draws keyed by (cell, callsite ordinal, SCC-pass, element ordinal)** — `RAND()+RAND()` distinct (B2); **`fullPrecision="0"` refuses through all of v1**, support = M10+ (B3); `openBuffer(alloc, io, bytes)` with borrow contract per Book precedent (B4); **`Writer.saveToOwnedBuffer`** producer API — `fresh_emit` alone can't serve the orchestrator (B5); **empty matrices normalize to `#CALC!`** (B6). New: multi-SCC schedule normative + 32767 clamp (7); `zlsx_recalc` third public module + consumer test (8); shape/coercion table (9); `collation_v1` + registry flags (10); error-propagation order contract (11); Unicode identifiers + extensible error literals in M1a (12); CT_CellFormula attribute inventory — `ca` honored, `aca`/`bx`/unknown refused (13); **ABI never writes beyond the known prefix** (14); status contract inlined, layouts to a committed design note (15); FormulaWrite Zig-only at M7c, C export `…_v2` at M9a2 (16); cancellation scope split — file-atomic vs prefix-valid streams (17); Spark activation/read-only/inference-on-recalced-snapshot defined (18); **recalculate vs saveWithRecalc transaction ordering** — serialize from prepared state, rename, then swap (19); multi-area → typed refusal v1 (20); sparse iteration + `logicalBlankCount` (21); volatiles excluded from ALL external oracles (22); doc checklist fully inlined (23); self-containment (24); limits table de-conflated — 64 = fn nesting only (25); cm/vm collection resolution pinned empirically (26); M5b and M9a split → **32 PRs** (27); CV2 = surrogate-pair handling, not graphemes (28m) |
| 7 | 2026-08-02 | **2 B / 15 M / 6 m** | All 23 → v8 (this revision, applied as targeted edits). **Cancellation commit region** — final poll immediately before rename; rename-through-swap non-cancellable; commit-aware CLI exits (B1); **implicit intersection split by operand kind** — references intersect by row/col, arrays reduce to top-left, `@SEQUENCE(3)` → 1 (B2). IFS/SWITCH reclassified **eager** per Excel, per-form contracts (3); **`ieee_fp_rules_v1` made normative** + per-mode signed-zero policy (4+20m); **scalar coercion matrix** — provenance × context × fidelity; locale-flavored text in arithmetic refuses, `"abc"+1` → `#VALUE!` (5); **ordinary text comparisons case-insensitive** via the existing full fold, `EXACT` case-sensitive (6); oracle uses **CalculateFullRebuild** + isolated instance + stale-dependency sentinel (7); coordinate-order iteration declared an intentional divergence with order-sensitive fixtures (8); shape-pass seeding for authored dynamic-array cycles (9); **outer rebuild × inner SCC loop integration** — changed SCCs reset, unchanged keep state (10); **NDJSON stream state machine** — eval-header/eval-cell/diagnostic/eval-complete; absent terminal = cancelled (11); **cancellation primitive at every layer** — RunInputs.cancel, C token API, Python timeout/KeyboardInterrupt (12); **UTC default everywhere** — no OS tz under stdlib-only; TZif = M10+ (13); **CV2 implemented in v1** (default for new workbooks since 2026-04; refusal dropped; shared CV1/CV2 text layer in M4f) (14); **AtomicFile.finish gains fsync** + POSIX dir fsync + injected failures (15); M7b split into persistence/criteria-dynrefs/stats → **34 PRs** (16); **normative Python API table** + probe flags (17); committed bench workloads + phase labeling (18m); 16384 renamed to a zlsx UTF-8 safety cap — no compiled-representation equivalence claimed (19m); `workbook.zig:5477-5479` added to the flip checklist (21m); eval exposure sentence aligned to the ladder (22m); seeded RNG marked an intentional divergence in both modes (23m) |
| 8 | 2026-08-02 | **4 B / 13 M / 2 m** | All 19 → v9 (this revision, targeted edits). **Rename is the commit point** — swap immediately after; post-swap dir-fsync failure = committed-with-durability-warning, never a contradicting error (B1); **standalone eval takes an explicit `dialect` (default `da`) at every layer** — dialect is a stored-cell property standalone text doesn't have (B2); **Spark cache keyed by digest + every resolved input + engine version**, never digest alone (B3); **`@` single-cell exception** — single-item references pass through unchanged; 2-D spanning rule pinned (B4); `collation_v1` moved to M3a, before any collation-dependent fn (5); **`Editor.save_with_recalc` C export + Python method** — the atomic transaction reaches the bindings (6); NDJSON grammar admits **pre-header refusal/cancel** (7); **`blank` first-class in NDJSON/C/Python** (8); formula error values renamed **`ExcelError`** — `ZlsxError` stays the exception (`__init__.py:195`) (9); **worker-thread cancellation** — signal handlers can't run during a blocked ctypes call (10); **"observed before commit" semantics** — post-commit trigger returns success + `cancelled_late` (11); token constructor status-style (12); **CLI `--anchor` restored** + full command syntax (13); Spark UTC default aligned (14); CV2 removed from the M10 backlog (15); `write_row_with_formulas` dialect kwarg specified (16); **polling bound: ≥1/1 024 cell-evals, ≥1/65 536 work units** + latency tests (17); seed as decimal string in JSON (18m); ladder count corrected to **33** (19m) |
| 9 | 2026-08-02 | **2 B / 15 M / 4 m** | All 21 → v10 (this revision, targeted edits). **`aca` un-refused** — valid always-calculate-array semantics, parsed+preserved+honored; only `bx` stays refused (B1, round-6 fix defective); **`<dimension ref>` expansion joins the approved spill-mutation set** (B2); `dialect` lands in `RunInputs` + the Zig signature (3, round-8 fix incomplete); **collation = full non-Turkic fold everywhere** — the repo's existing algorithm; "simple fold" contradiction removed (4); **`alignedRangeIterator`** — sparse positional zipper for `SUMIF(A:A,"",B:B)`-class alignment (5); **locale-sensitive OUTPUT pinned invariant in v1** — no locale field; TEXT/FIXED/DOLLAR/NUMBERVALUE divergences fixtured (6); minimal INDIRECT/OFFSET moved into M5a2 so fixpoint fixtures precede their functions (7); **atomic-save promise corrected for existing/in-place destinations** — "destination bytes unchanged until commit" (8); **limits are Plane-2 refusals at every layer** — CLI 3, C −2, one mapping table (9); **Spark cache byte-bounded LRU** (10); Spark `evalDialect` removed — no standalone eval op there (11); **C ABI build flips multi-threaded at M9a1** — `-fsingle-threaded` atomics can't back a cross-thread token; CLI keeps sig_atomic_t (12); **dormant durability slot** — post-rename dir-fsync outcome needs no allocation (13); **full-row references** specified for evaluation (14); **single `parseDecimal` ingress primitive** with enumerated callers (15); M3a and M5a split → **35 PRs** (16); `saveToOwnedBuffer(allocator, io)` — allocator-first per AGENTS.md:194 (17); per-symbol-group probes (18m); `max_output_archive_bytes` named (19m); milestone-local perf gates after M5d (20m); structural-edits.md flip + `recomput|recalculat` rg terms (21m) |
| 10 | 2026-08-02 | **2 B / 16 M / 2 m** | All 20 → v11 (targeted edits). Blockers were v10 edit residues, fixed immediately: `aca` removed from the §10 refusal row (B1); residual "no file" phrasing replaced by prior-bytes-intact in §5.7.9 + Python row (B2). **Input cell-type contract** — unknown `t` refuses, `t="d"` pinned (3); **SUMIF top-left projection + N-way `*IFS` cursor** replace the pairwise zipper (4); **ST_Xstring decode moved to M4b1** (SST/inline consumers), encode at M5b1 (5); **`parseDecimal(fidelity, ingress)`** with a caller→ingress table (6); **one total preorder** — full-folded compare + raw tie-break for ALL ordering; only EXACT/CODE exempt (7); **borrow-lifetime preservation** — stable sheet slots, superseded arenas retained to deinit (8); **`CancelToken` union + `error.Cancelled`** across storage kinds (9); build premise corrected — single-threaded is an option; M9a1 hard-sets ABI multi-threaded + compile assertion (10); **chunked cancellation-aware deflate seam** at M5d (11); **byte vs work limits split** — counted allocator vs explicit counters (12); **`evaluate` purity contract** (13); INDIRECT/OFFSET completed at M5a2 — no half-functions behind M6 (14); **probed engine fingerprint** for Spark identity (15); **`--formula` flag** — parser rejects leading-dash positionals (16); explicit `cancelled` NDJSON terminal; EOF-no-terminal = crash (17); Python/Spark doc gates at M9a2/M9b incl. the stale streaming claim (18); dialect fingerprint projections per operation (19m); SORT/UNIQUE bench to M7a (20m). **Owner directive recorded: loop to SHIP-READY, then commit the plan via PR.** |
| 11 | 2026-08-02 | **3 B / 14 M / 2 m** | All 19 → v12. Second differently-wrapped "no file" residue excised — both destination cases now tested at every failure point (B1); **raw tie-break removed from the semantic preorder** — fold-equal IS equal; SORT stability via private source-position tie-break (B2); **operation-specific blank classes** — COUNTBLANK counts `=""`, ISBLANK doesn't (B3); borrow-lifetime preservation made normative in §5.7.4 — the v10 edit had silently failed to land, exactly as Codex inferred (4); M5c = importable shell, composition + consumer test at M5d (5); cancellation seam extended to PartStore preparation compression (6); `t="d"` deferred to M4b2 with `date1904` (7); **cache-state contract** — absent vs malformed `<v>` separated; malformed refuses, never zero-seeds (8); **cell-text 32 767 reclassified as Excel-domain `#VALUE!`** + separate 1 MiB safety refusal (9); N2 zero-snap scoped additive-only + counterexamples (10); **`EvalOwnerId`** — volatile determinism for names + standalone roots (11); SST-number ingress path removed (12); **numeric defaults for every aggregate limit**, echoed + fingerprinted (13); fingerprint export staged at M9a1 with probe (14); **cancel excluded from EffectiveRunInputs** (15); randomized differential graph test vs brute force (16); **`casing_v1`** — folding can't implement casing; UPPER("ß") oracle-pinned (17); tokenizer scope-comment flip at M1a (18m); `writer.zig:1645` stale claim + regex extension (19m) |
| 12 | 2026-08-02 | **3 B / 7 M / 2 m** | All 12 → v13. **Whole-generation retention** — sheet leaves borrow part bytes, so the old PartStore (source handle + arena) is retained, not just views (B1); **per-cell FormulaSpec descriptor + CSE-rectangle state machine** — parallel text arrays can't encode a rectangle (B2); **invocation-path volatile keys** — `N=RAND()` referenced twice draws twice, oracle-pinned (B3); `abs()` in convergence (4); `xml:space` preserved + `bx="0"` accepted (5); **`casing_v1` = full casing, SpecialCasing-backed** — simple mappings can't produce `ß`→`SS` (6); limits resolution completed — min(calcPr, caller, 32767), ZIP32-exact archive bound, CLI/Python fixed-at-defaults declared (7); cancellation chunking extended to decompression/raw-copies/XML-scans/temp-writes + documented `File.sync` SLA exception (8); **embedding stale-coverage policy** — recalc preflights `include_formulas` coverages, refuse-or-mark-stale, never silent hash drift (9); **perf becomes release-blocking post-baseline + M9d sign-off** (10); M5d and M8 split → **39 PRs** (11m); already-false vs_calamine claims → M-1 (12m) |
| 13 | 2026-08-02 | **5 B / 7 M / 1 m** | All 13 → v14 — and the apply process now asserts every replacement individually (three prior "fixes" had silently no-op'd on wrapping mismatches, exactly as Codex kept inferring). `abs()` convergence actually landed + array max-elementwise rule (B1); **Xstring decode for formula/name/table carriers + inverse encode** (B2); **`aca` = calculation granularity, not volatility** — whole-array under both values, circular `aca=false` refuses (B3); **reference-occurrence ordinals** — `N+N` draws twice; property-based volatile oracles admitted (B4); **Writer context-exit save-generation guard** (B5); embedding policy normative as §5.7.1e with transaction ordering + tests (6); **counted generation retention** — max 4, byte/fd accounting, pre-swap refusal (7); coercion matrix is an actual table incl. cross-type ordering and direct-vs-range asymmetry (8); **array-selector masking for lazy forms** (9); one comparator stated once (10); dir-fsync joins the SLA exceptions (11); **absolute release ceilings** — 500 ms eval / 1 s total / 3× RSS, owner-waiver only, wired into M9d (12); CSE authoring unified at M7c (13m) |
| 14 | 2026-08-02 | **3 B / 12 M / 3 m** | All 18 → v15. `.mark_stale` deferred (M10+) — `zlsxER1` has no status field; v1 refuses (B1); save-generation guard covers every save (B2); static-edges vs runtime-lanes split + v14 fragment repaired; CHOOSE at M4e (B3); blank publication → 0 (4); rich-string run concatenation (5); CSE anchor-only authoring (6); Unicode identifier grammar (7); MIN/MAX out of comparator (8); `FormulaSignedWorkbook` (9); CT_DefinedName inventory (10); volatile-oracle unified (11); pre-mutation zero-formula fast path (12); utc_offset i32 (13); pinned perf environment (14); function names frozen — **40 PRs** (15); `ca` = dirty hint (16m); closed profile enum + RFC 3339 (17m); `build.zig:891` flip (18m) |
| 15 | 2026-08-02 | **5 B / 11 M / 3 m / 1 nit** | All 20 → v16. **One candidate, swap LAST** (B1); **save-generation guard in core Writer**, all paths route through (B2); **preflights-first fast path** (B3); **embedding overlap post-staging on canonical payloads incl. DA tails** (B4+6); **casing_v1 in M4f** (B5); zero-cell path runs the carrier census (7); absent CV = CV1 (8); **retained byte/fd limits + preflight** (9); **XID tables from pinned Unicode 17** + License V3 notices + carve-out (10+16); **Spark one-buffer digest rule** (11); **3D list frozen to six** (12); **solver contract before M9c1** (13); measurable perf gate (14); **folded→original position map** (15); CT_DefinedName completed (17m); pointer stability scoped nonstructural (18m); `_xlpm.` fixed (19m); `ca` residue gone (20n) |
| 16 | 2026-08-02 | **6 B / 13 M / 3 m** | All 22 → v17. **ST_Xstring removed from formula carriers** (B1, corrects round-13); **fast path requires empty mutation journal** (B2); **embedding invalidation = any coverage overlap** (B3); **engine provenance — no flag clearing in v1** (B4); **3D context-legality gate** (B5); **licensing deliverable** — LICENSE carve-out (owner action) + notices in every artifact (B6); embedding step renumbered 3b (7); signed no-op via census-first (8); monotonic deadline (9); **blank internal-only** (10); typed per-element convergence (11); mark_recalc_on_load export (12); taxonomy scoped (13); WorkBudget (14); LO hard-recalc + sentinel (15); casing sources corrected (16); warm-cache (17); materialization caps (18); backslash start-only (19); deinit-and-reopen (20m); build.zig scanned (21m); post-swap invariant narrowed (22m) |
| 17 | 2026-08-02 | **5 B / 5 M** | All 10 → v18. Surviving inverse-ST_Xstring clause in M4b1 deleted — authored formulas XML-escape only (B1); **truthful producer state** — calcId set to "0" + fullCalcOnLoad set on every zlsx recalc; preserving Excel's ID would misattribute zlsx caches to Excel (supersedes both prior calc-flag policies) (B2); blank residues swept from ScalarValue comment/NDJSON union/Python Matrix — evaluator-internal, publishes as 0 (B3); **registry match policies** — FIND/SUBSTITUTE raw like EXACT; TEXTBEFORE-family arg-selected (B4); **wildcards are CV-independent** — CV touches exactly five functions' index units (B5); deadline joins cancel outside EffectiveRunInputs (6); **fidelity-specific oracle precedence** — ieee led by its rule table, Excel demoted to divergence witness (7); `t="d"` normative lexical table — two forms, no zones, epoch-checked (8); **`numfmt_v1` grammar + support matrix before M8a** (9); bench tooling aligned — N=20, `compare_bench.py --gate` mode added in M5d3 (10) |
| 18 | 2026-08-02 | **3 B / 6 M** | All 9 → v19. calcId residue deleted + byte-level assertion (B1); **implicit `c@r` reconstruction** (B2); **Writer artifact keyed by full context** (B3); **Published types + single blank→0 conversion** (4); **provenance-aware COUNT policies — COUNTA counts errors** (5); `deadline: ?std.Io.Timestamp` (6); **AST copy-translation matrix** (7); **SourceBacking before M5b2** (8); **FormulaMalformedInput split + mark-eligibility table** (9) |
| 19 | 2026-08-02 | **3 B / 4 M / 2 m** | All 9 → v20 — apply script hard-asserts every edit. **SourceBacking → M5b0** (B1); **COUNT provenance normative** (B2); **XLRICHVALUE/unknown metadata refuses** (B3); **FormulaPrecisionAsDisplayed always-refuse** (4); **b/e lexical tables** (5); **per-layer clock/seed defaults** (6); **SIGPIPE exception, exit 0** (7); Python cancellation unified (8m); vs_calamine split M-1/M6 (9m) |
| 20 | 2026-08-02 | **3 B / 5 M / 1 m** | All 9 → v21 — apply script hard-asserts every edit. **CSE broadcasts scalars ONLY** — 1×N/N×1 vectors are coordinate-placed with `#N/A` padding, not replicated (B1, corrects round-7); **semantic `iterateCount` split from the resource ceiling** — workbook exhaustion succeeds non-converged, a lower caller cap refuses with zero mutation (B2); **M5b0 physically moved ahead of M5b1/M5b2** + M5b2 gate re-runs its ownership tests (B3); **formula-plus-absent-`<v>` outranks `t`** — b/e lexical tables apply only when `<v>` is present (4); **Python `EvalResult{value, resolved}`** + explicit "Zig/C have no defaults" row (5); **entropy/clock acquisition failure = exit 6**, exit 4 stays OOM-only (6); **`rangeIterator` is a logical iterator** merging stored + staged + computed + virtual spill tails, with same-run grow/shrink tests (7); **control-aware `saveToOwnedBufferControlled`/`openBufferControlled`** so Writer-path pre-recalc stages meet the polling bound (8); ladder count corrected to **41** (9m) |
| 21 | — | pending | — |
