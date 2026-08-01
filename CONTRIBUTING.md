# Contributing to zlsx

Thanks for considering a contribution. zlsx is **proprietary software**
whose repository is public for reading — see [LICENSE](LICENSE) — with
commercial licenses negotiated separately. Two things every contributor
needs to agree to before a PR can be merged.

## 1. Developer Certificate of Origin (DCO)

Every commit must include a `Signed-off-by:` line, which certifies that you
wrote the patch (or otherwise have the right to submit it under the project's
license). The full text of the DCO is at <https://developercertificate.org>.

Add the sign-off automatically with:

```bash
git commit -s -m "your message"
```

A commit signed off by you looks like:

```
your commit message

Signed-off-by: Your Name <your.email@example.com>
```

CI rejects PRs whose commits are missing the trailer.

## 2. Inbound Licensing Grant

By submitting a contribution to zlsx, you agree that:

1. Your contribution is licensed to the project under the terms in
   [LICENSE](LICENSE) — the same terms as the rest of zlsx.
2. You **additionally grant the project maintainer (Laurent Fabre) a
   perpetual, worldwide, non-exclusive, royalty-free, irrevocable license**
   to relicense your contribution — in whole or in part — under any other
   terms, including commercial terms, as part of zlsx.

This dual grant lets the project keep offering paid commercial licenses to
companies without having to track down every contributor for permission.
It does not take any rights away from you: you keep your copyright, and you
can still use, distribute, or relicense your own contribution however you
want.

By opening a pull request you confirm that you have read this section and
agree to both grants.

## Commit style

- Imperative subject ("Add foo" not "Added foo"), ≤ 72 characters.
- Body explaining the *why* for non-trivial changes.
- Conventional-commit-ish prefixes are common in the history (`feat(pkg):`,
  `fix(reader):`, `docs(plans):`, etc.) — match what's already there.

## Tests

```bash
zig build test                # unit + fuzz-smoke (~700 ms)
ZLSX_FUZZ_ITERS=1000000 zig build test    # deeper fuzz
zig build integration         # public-corpus integration suite
```

If a test would download or commit corpus fixtures, see
`scripts/fetch_test_corpus.sh`.

## Bench gate

Performance changes are gated by `scripts/bench_ci.sh` against the canonical
baselines in `docs/benchmarks.md`. A regression > 10% will fail CI.

## Questions

Open an issue at <https://github.com/laurentfabre/zlsx/issues>.

For commercial licensing inquiries, see [LICENSE](LICENSE) §
"Commercial use" or email `laurent.fabre@gmail.com`.
