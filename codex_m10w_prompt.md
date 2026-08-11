# Review target — M10w (§9.1f of goal_formula.md), zlsx

You are reviewing one performance row on a long memory-reduction ladder,
plus the write-up that records it. The write-up's standard is that **every
number in it must be recomputable from the tables beside it**, and that
no sentence may claim more than the measurement it cites.

## What to read

1. `git diff 640e05b..HEAD` — the whole row (the branch is stacked on M10v, so `main` is two rows behind; the code diff and the new section are appended to this prompt in full) (Zig source, bench harness, and
   the new §9.1f section in `goal_formula.md`).
2. §9.1f in `goal_formula.md` (the new section at the end of §9.1).
3. For context only, the three preceding rows: §9.1c (M10t), §9.1d
   (M10u), §9.1e (M10v).

## What the row claims

- The resident maximum on the digest-gated `f1_mix_named` fixture was one
  `gpa` block in `graph.zig`'s `link` (10 210 051 B) standing beside
  20 734 840 B of builder scratch.
- `Allocator.alloc` memsets to `undefined` in Debug/ReleaseSafe, so the
  whole block became resident at its request — including six regions
  first read only after three builder lists are freed.
- Two changes: request the block with `rawAlloc` (mapped, not written;
  each `fa.alloc` still poisons its own region), and move `Index.build`
  below those three frees.
- Nothing was removed: same block size, same lifetime. Page *arrival*
  moved, and the envelope's gap paid 2 031 616 B of a 2 048 000 B budget.
- The resident maximum is now the decode era, which re-opens a candidate
  §9.1c had priced at zero pages.

## Specific things to attack

1. **Arithmetic.** Recompute every figure from the tables rather than
   reading the prose. Census terms must sum to `block_bytes`; the
   deferred set must equal `block_bytes − carved_before_last_read`; the
   gate deltas must equal the differences of the quoted adjusted figures;
   the ReleaseSafe/ReleaseFast comparison must be internally consistent.
   Name any figure that does not reproduce.
2. **Attribution.** The row says the cut's payoff is capped by the gap to
   the runner-up era, and that the maximum moved to decode. Is that
   *shown*, or is it inferred from era vectors that each arm segments
   independently? What could produce the measured −2 031 616 other than
   the stated cause?
3. **Lane honesty.** In ReleaseFast the memset does not happen. Does the
   row overstate what a user gets? Is the ReleaseSafe/ReleaseFast split
   measured or asserted? Is any per-cell or per-byte figure quoted
   without its lane?
4. **Safety of `rawAlloc`.** The row claims nothing a caller can read is
   less poisoned than before. Find a path where a caller reads a byte of
   that block that no `fa.alloc` handed out. Consider the refusal/error
   paths, `Graph.deinit`, the `probe_block_padding` slack, and alignment
   padding between carved regions.
5. **Correctness of the reorder.** `Index.build` moved below three
   `clearAndFree` calls. Find any reader of `idx` between its old and new
   positions, any dependency of `Index.build` on freed state, or any way
   the changed `FixedBufferAllocator` carve order can overflow the
   block's 128 B headroom on some other input shape (deep nesting, many
   names/producers, spill tails, cyclic components).
6. **The instruments.** `census_sink` and `probe_block_padding` are
   mutable globals in a library. Is the production cost claim ("one load
   and one add per graph build") true as written? Can a released build
   reach a state where padding is non-zero? Is the padding knob measuring
   what the row says it measures, given it is never carved?
7. **The tests.** Three new tests. Each is claimed to fail without its
   fix (verified: the stamping tests read 0xaa instead of 0x5a with
   `rawAlloc` reverted). Are they testing the invariant or an
   implementation detail that a legitimate refactor would break? Is the
   ReleaseFast weakening of the stamp tests stated?

## What I believe the evidence proves, so you can tell me what it is blind to

The padding sweep (0 / 1 / 2 / 4 MiB, same binary, three reps each, every
figure byte-identical) gives exactly 1.000 B of resident page per byte of
block *before* the cut. After the cut the same knob reads **+0 up to
1 343 488 B of padding and then 1.000 B/B**, with the absorbed amount
equal to 1 343 488 at both padded points — the same figure the era vector
reports as the new gap from the maximum to the runner-up, from an
instrument with a different baseline. I read the pair as the memset
mechanism plus the envelope's shape, measured on the gate's own fixture.
Tell me what else those readings are consistent with, and what they
cannot establish. In particular: is my attribution of the post-cut
padding cost to `Allocator.free`'s memset at graph teardown necessary, or
would something else produce the same hinge?

## Output format

For each finding:

```
[SEVERITY: BLOCKER | MAJOR | MINOR | NIT]
[KIND: arithmetic | attribution | overclaim | correctness | safety | test | style]
[WHERE: file:line or §9.1f subsection]
[CLAIM: the sentence or number you are disputing, quoted]
[WHY: what is wrong, with the recomputation if it is arithmetic]
[FIX: the smallest change that makes it correct]
```

End with exactly one line:

`SHIP-READY: yes | no — <one sentence>`
