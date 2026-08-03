# Deferred: masking string literals inside formulas

**Status:** deferred, deliberately. Recorded so the gap is visible
rather than discovered later by someone auditing a "masked" workbook.

---

## The gap

zlsx can mask cell *values* (`Editor.setCell` / `setCells`) and, as of
Z3, document *metadata* (`Editor.stripDocProps`). It cannot mask a
string literal embedded in a formula:

```
A1:  ="Report for " & B2          <- "Report for " is untouched
A2:  =IF(C1="Jane Q. Fixture", 1, 0)   <- the name survives masking
```

A pipeline that pseudonymises `C1` but leaves `A2`'s formula intact has
leaked the very value it set out to hide, and the leak is invisible to
a reviewer looking at rendered cell values — Excel shows the *result*,
not the literal.

## Why it isn't done here

`src/formula/rewriter.zig` rewrites **references** (`B2` → `B3` when a
row is inserted). It is a reference rewriter, not a general formula
text rewriter: it does not need to understand string literals, operator
precedence, or function arity to do its job, and it deliberately
preserves everything it does not recognise byte-for-byte.

Masking literals needs strictly more:

- tokenise the formula (the tokenizer in `src/formula/tokenizer.zig`
  already does this, and is loss-preserving),
- identify string-literal tokens, including the `""` escape for an
  embedded quote,
- substitute replacements of *different length*, then re-emit,
- and re-emit without disturbing the cached `<v>` value, which Excel
  will otherwise show as stale until recalculation.

That last point is the sharp edge. Changing a formula invalidates its
cached result, and a literal-masking pass would have to either drop the
cached value (a visible change for every consumer that reads values
without recalculating) or knowingly leave a stale result that still
contains the unmasked string.

> **Corrected 2026-08-02 (M-1).** This section previously said "zlsx's
> writer emits formulas with no cached `<v>` so Excel recalculates, but
> the *editor* path preserves existing `<v>` elements." Both halves were
> wrong, and in opposite directions: the writer emits whatever cached
> value the caller hands it, and the editor path drops the cached value
> of the cell it rewrites (it preserves the `<v>` of cells it does not
> touch, which is a different claim). The policy per path:

### Cache policy per write path

| Path | API | Cached `<v>` on the written formula cell | Other formula cells |
|---|---|---|---|
| Fresh emit | `SheetWriter.writeRowWithFormulas` (`src/writer.zig:984-991`) | **Caller-supplied.** The accompanying `cells[i]` is emitted as the `<v>` Excel displays until recalculation; pass `.empty` for a formula cell with no cached value (Excel shows 0 initially). | n/a — new file |
| Overlay / editor mutation | `Workbook.setCell` with `.formula` (`pkg/workbook.zig:366-369`; emit at `:5476-5483`) | **None.** The cell is re-emitted as `<f>…</f>` with no `<v>`, so any prior cached result on that cell is dropped and Excel recalculates on open. | Preserved byte-for-byte — untouched `<c>` bytes are spliced through. |
| Read | `Book.formula(sheet, ref)` | Returned exactly as stored; zlsx does not compute it. | — |

So a masking pass on the overlay path already pays the "drop the cached
value" cost by construction — the stale-result branch is only reachable
if the pass rewrites `<f>` bytes in place rather than going through
`setCell`.

## Where it belongs for now

Caller-side, in the consuming pipeline (nemonym), which already owns
the mask dictionary and can decide the recalculation policy for its own
outputs. `Book.formula(sheet, ref)` exposes the formula text for
inspection, so a caller can detect and *report* affected cells today
even though zlsx will not rewrite them.

## What would change the calculus

- A `Workbook`-level "invalidate cached values on this sheet" primitive,
  which would make the stale-`<v>` question a policy the caller sets
  once rather than a per-formula hazard. **As of 2026-08-02 the D1
  ladder (`goal_formula.md`) supplies a stronger version of this**:
  `markRecalcOnLoad` (M5b2) and full recalculation on the save path
  (M5d2). Once either lands, a masking pass can mask and then either
  mark or recompute, instead of choosing between a dropped and a stale
  cache. This does not by itself unblock literal masking — the
  tokenise-substitute-re-emit work above is still unwritten.
- Evidence that formula literals are common in real masking corpora.
  Current read: they are rare in data-export workbooks (the dominant
  zlsx use case) and common in hand-built report templates.

Until then: **zlsx masks values and metadata, not formula text.** Any
documentation of the masking capability should say so plainly rather
than let a reader assume coverage it does not have.
