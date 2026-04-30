# Enforcement plan — worktrees, subagents, TDD, merge guards, review

Status as of 2026-04-30. Single-author public repo (`laurentfabre/zlsx`), Zig 0.15.2.

## Status table

| Phase | Description | Status |
|---|---|---|
| 1 | Free wins: branch protection, repo merge settings, CODEOWNERS, PR template | **Done (2026-04-30)** |
| 2 | TDD CI gates: test-presence, C-ABI 3-file-transaction, monotonic test count | **Done (2026-04-30, advisory)** |
| 3 | Worktree + subagent conventions: helper script, commit-msg trailer, PR template fields | **Done (2026-04-30)** |
| 4 | Agent-as-reviewer CI job (codex-review-on-PR) | **Done (2026-04-30, requires `OPENAI_API_KEY` secret)** |
| 5 | Optional: coverage gate, TDAD map, mutation testing | Deferred |

---

## Baseline (what exists today)

| Layer | State |
|---|---|
| `main` branch protection | None. Direct push allowed. |
| Required status checks | None. CI runs but does not gate merges. |
| Required reviews | None. |
| CODEOWNERS / PR template / Dependabot | Absent. |
| Pre-commit hook | Installed via `scripts/githooks/install-hooks.sh`. Gates: merge markers, CLAUDE.md size cap (vestigial here), 21-pattern secret scan, `zig fmt --check`. Bypassable with `--no-verify` and opt-in per dev. |
| Pre-push hook | None. |
| CI workflow | `ci.yml`: fmt, build (Debug + ReleaseFast), unit + corpus tests on macOS-14 / Ubuntu-22.04 / Windows (`continue-on-error`), cross-compile to 4 targets, single-threaded compile. |
| Worktree convention | None. Single worktree on `main`. |
| Subagent attribution | None. |
| Merge styles | All three (squash, rebase, merge commit). `deleteBranchOnMerge: false`. |

**Net**: a careful solo dev can push directly to `main`, skip pre-commit, and merge a red PR. Only client-side discipline catches mistakes.

---

## Five gates — mechanism + recommendation

### 1. Worktree per PR

Solves: parallel agent sessions stomping each other; cache thrash from branch-switching.

| Layer | Enforceable? |
|---|---|
| Server-side | **No.** GitHub has no concept of which checkout you used. |
| Client-side | A `scripts/wt-new <branch>` helper that creates `../zlsx-<branch>` with isolated cache. Optionally a pre-commit warning when committing to a feature branch from the primary worktree. |
| Convention | Document in AGENTS.md so agents and humans default to it. |

**Recommendation**: convention + helper, not a hard gate. Cost of dropped enforcement is "cache rebuild," not data loss.

### 2. Subagent per PR

Solves: attribution, reproducibility, writer/reviewer hygiene.

| Mechanism | Enforceability |
|---|---|
| Commit-msg trailer (`Agent: <name>`) | Soft — forgeable, but useful as a default. |
| PR template field (Author agent / Reviewer agent) | Soft — relies on filling the template. |
| Writer/reviewer split | Workflow discipline; not enforceable in git/CI. |

**Recommendation**: PR template field + soft commit trailer. Don't gate.

### 3. TDD

| Mechanism | What it catches | Cost | Verdict |
|---|---|---|---|
| Test-file-changed-when-source-changed CI rule | Adding `src/foo.zig` without touching tests | Trivial — diff parse | Worth doing. ~10% false-positive rate on refactors; escape via PR label. |
| Coverage gate (`zig test --test-coverage` + kcov, Linux only) | Untested code paths | Tooling is rough | Skip for now. |
| Mutation testing | Tests that exist but don't test anything | No Zig tooling | Skip. |
| Monotonic test count | Stealth test deletion | One-line CI check | Worth doing. |
| TDAD (code-to-test dependency map) | Per AGENT-PRACTICES research: 70% regression reduction vs procedural TDD prompts | Map-gen script + agent integration | One-time experiment, not a hard gate. |
| C ABI 3-file-transaction check | The `c_abi.zig` ↔ `zlsx.h` ↔ `_ffi.py` rule | CI diff check | Worth doing — highest ROI for this codebase. |
| ctypes-narrowing audit | Python boundary integer assignments without range checks | Static lint | Small linter, low priority. |

**Recommendation**: build the C-ABI 3-file-transaction check first, then test-presence + monotonic count. Skip coverage and mutation tooling.

### 4. Merge guards

GitHub branch protection on `main`:

| Setting | Recommended | Why |
|---|---|---|
| Require pull request | Yes | Forces CI on the diff |
| Required approvals | 0 (solo) → 1 when collaborators land | Can't approve own PR; gating self blocks all work |
| Dismiss stale approvals on new commits | Yes | Approvals follow code, not time |
| Required status checks | Yes — `test (macos-14)`, `test (ubuntu-22.04)`, all `cross-compile/*`. Skip `windows-runtime` while it's `continue-on-error`. | |
| Require branches up-to-date before merging | Yes | Catches semantic conflicts |
| Require conversation resolution | Yes | PR-review hygiene |
| Signed commits | Optional | Worth it across multiple machines |
| Linear history | Yes | Forces squash or rebase; no merge commits |
| Block force pushes | Yes | |
| Block deletions | Yes | |
| Restrict who can push | Owner + admins | |
| Allow bypass | No, even for admins | Or "yes but log it" if you need an emergency lever |

Plus repo-level: disable merge-commit style, enable auto-delete on merge.

**Recommendation**: do this first. Highest ROI of the entire plan. Solo-author still benefits.

### 5. Review

| Mechanism | Cost | Verdict |
|---|---|---|
| Required approvals = 1 | Free, but blocks solo work | Skip while solo. Enable when second contributor lands. |
| `CODEOWNERS` for paths | One file | Worth adding even solo: `* @laurentfabre` for now. |
| Agent-as-reviewer (auto-run codex review on PRs) | Hook + CI artifact upload | The interesting option. Real value. Phase 4. |
| `dangerjs`-style PR linter | One config | Low-value at this scale. Skip. |
| PR template | One file | Worth it. Forces scope/test/ABI thinking. |

**Recommendation**: PR template + agent-as-reviewer CI job (Phase 4).

---

## Layered plan, in priority order

### Phase 1 — Free wins (~10 minutes)

1. Branch protection on `main`: required PR + required status checks + linear history + block force/delete + dismiss stale approvals + 0 required approvals.
2. Repo merge settings: disable merge-commit, enable auto-delete on merge.
3. `.github/CODEOWNERS` with `* @laurentfabre`.
4. `.github/pull_request_template.md` with: summary, scope, test plan, ABI impact, agent attribution.

### Phase 2 — TDD CI gates (~1-2 hours)

5. **Test-presence check**: CI fails if a `src/*.zig` non-import-only change ships with no `tests/` or `test "..."` change. Escape via PR label `no-test-needed`.
6. **C-ABI 3-file-transaction check**: CI fails if `src/c_abi.zig` changes without `include/zlsx.h` AND `bindings/python/zlsx/_ffi.py` changing.
7. **Monotonic test count**: PR cannot net-decrease `grep -c '^test "' src/**/*.zig` without label `delete-tests-ok`.

### Phase 3 — Worktree + subagent conventions (~30 minutes)

8. **`scripts/wt-new <branch>`**: helper that creates `../zlsx-<branch>` with isolated cache.
9. **Commit-msg hook**: require `Agent: <name>` trailer (soft warn, not error).
10. **PR template fields**: "Author agent" / "Reviewer agent."

### Phase 4 — Agent-as-reviewer (~half a day)

11. **GitHub Actions job: codex-review-on-PR**. Runs a constrained codex review on the diff, posts findings as a PR comment. Required to pass (or to post) before merge.

### Phase 5 — Optional, gate by pain

12. Coverage gate via `zig test --test-coverage` + kcov on Linux only.
13. TDAD code-to-test map generator.
14. Mutation testing (defer indefinitely; no Zig tooling).

---

## What NOT to do

- Require human review approval while solo — blocks all work.
- Gate merges on the `windows-runtime` job until `continue-on-error: false`.
- Gate merges on coverage thresholds — Zig tooling is rough; false-fails will exceed signal.
- Enforce "subagent per PR" via CI — trailers are forgeable; the value is in *how* you invoke agents.
- Add husky / Node-based hook framework — bash hooks already work.
- Enforce TDD via "test ratio" metrics — Goodhart's law; people pad with trivial tests.

---

## Phase 1 — implementation log

- [x] Branch protection applied to `main` (2026-04-30)
  - 6 required status checks: `test / macos-14`, `test / ubuntu-22.04`, `cross / {x86_64,aarch64}-linux-musl`, `cross / x86_64-windows-gnu`, `cross / aarch64-macos`
  - `windows-runtime` deliberately excluded (still `continue-on-error: true` in `ci.yml`); add once flipped to `false`
  - `strict: true` (PR must be up-to-date with `main` before merge)
  - `required_approving_review_count: 0` (solo-dev mode)
  - `dismiss_stale_reviews: true`, `required_conversation_resolution: true`
  - `required_linear_history: true`, `allow_force_pushes: false`, `allow_deletions: false`
  - `enforce_admins: true` — no bypass, even for the owner. Flip to `false` if an emergency lever is needed.
  - Restore current settings: `gh api repos/laurentfabre/zlsx/branches/main/protection`
- [x] Repo merge settings updated (2026-04-30)
  - `allow_merge_commit: false`, `allow_squash_merge: true`, `allow_rebase_merge: true`, `delete_branch_on_merge: true`
- [x] `.github/CODEOWNERS` created — `* @laurentfabre`
- [x] `.github/pull_request_template.md` created — summary, scope, test plan, C ABI 3-file checklist, roadmap link, agent attribution

## Phase 2 — implementation log

- [x] `.github/workflows/pr-gates.yml` created (2026-04-30)
- [x] `scripts/ci/test-presence-check.sh` (Gate 5)
- [x] `scripts/ci/abi-3file-check.sh` (Gate 6)
- [x] `scripts/ci/monotonic-test-count.sh` (Gate 7)
- [x] Escape labels created on GitHub: `no-test-needed`, `abi-no-3file`, `delete-tests-ok`
- [x] Local smoke-tested against historical commits:
  - ABI gate fails correctly on `ab2c1a2` (c_abi.zig only) — would have caught real historical lapse in `001af88` (c_abi.zig + header but no `_ffi.py`).
  - Test-presence passes on `ab2c1a2` because the commit added inline `test "..."` blocks to c_abi.zig (correct).
  - Monotonic gate fails correctly on a synthetic 480→0 fixture, passes with `delete-tests-ok` label.
- [x] **Done (2026-04-30)**: promoted gates to `required_status_checks`. Three contexts added: `Test-presence check`, `C ABI 3-file transaction`, `Monotonic test count`. Total now 9 required contexts (the original 6 + these 3). The gates have run green on every PR since #6 and the escape labels are exercised whenever needed; the advisory window worked as intended.

### Local invocation (debugging a gate)

```sh
BASE_SHA=$(git merge-base origin/main HEAD) HEAD_SHA=HEAD LABELS='[]' \
  bash scripts/ci/test-presence-check.sh

BASE_SHA=$(git merge-base origin/main HEAD) HEAD_SHA=HEAD LABELS='["abi-no-3file"]' \
  bash scripts/ci/abi-3file-check.sh
```

### Known limitations

- **Test-presence**: heuristic "non-trivial change" filter strips blank lines and `// ...` comments only. Multi-line `///` doc-comments still count as significant; tolerable since the escape label exists.
- **ABI gate**: only triggers when `src/c_abi.zig` is in the changed-file list. An ABI-affecting change made elsewhere (e.g., changing an extern struct's layout via a `pub const` in another file imported by `c_abi.zig`) won't be caught.
- **Monotonic**: counts `^test "` exactly — moving a test from `test "foo"` to `test "renamed"` is a no-op (count preserved), but reformatting a test header onto a multi-line form would silently drop it.

## Phase 3 — implementation log

- [x] `scripts/wt-new <branch> [base]` — creates a sibling worktree at `../<repo>-<branch>/` from `origin/<base>` (default `main`). Convention only; no enforcement.
- [x] `scripts/githooks/commit-msg` — soft-checks for an `Agent: <name>` trailer. Skips merge / squash / cherry-pick / fixup messages. Soft-warn only — commit proceeds.
- [x] PR template `Author agent` / `Reviewer agent` fields — already shipped in Phase 1 (`.github/pull_request_template.md`).
- [x] `AGENTS.md` updated with a "Workflow conventions" section documenting both.

### Notes

- Hooks are wired via `core.hooksPath = scripts/githooks` (set by `scripts/install-hooks.sh`). Adding a new hook is just dropping the file in `scripts/githooks/` with executable bit set.
- The `commit-msg` hook intentionally never blocks. If the convention sticks (i.e. trailers appear consistently), promote to a hard fail later.
- `wt-new` slug-escapes `/` in branch names (`feat/streaming-sst` → `zlsx-feat-streaming-sst/`) to avoid filesystem nesting surprises.

## Phase 4 — implementation log

- [x] `.github/workflows/codex-review.yml` created (2026-04-30)
- [x] `codex-review` PR label created on GitHub (forces a review run; also re-runs on draft PRs).
- [ ] **Deferred (2026-04-30)**: add `OPENAI_API_KEY` as a repository secret under Settings → Secrets → Actions. Until set, the workflow gracefully no-ops with a CI warning. Picking this back up is a one-time action; the workflow itself is already in place.
- [ ] **Deferred (2026-04-30)**: promoting `Codex review` to `required_status_checks`. Gating on a third-party AI review on every PR isn't worth the merge friction yet; the comment-posting flow gives the same reviewer signal without blocking. Promote later if/when both the secret is set and signal-to-noise has been observed across ~5+ real PRs.

### Workflow shape

- Triggers on `pull_request` (opened / synchronize / reopened / ready_for_review / labeled).
- Skips draft PRs unless they carry the `codex-review` label.
- Pins reasoning effort to `low`, hides agent-reasoning trace, disables MCP servers, and excludes secret-shaped env vars from any spawned shells.
- Posts the review as a single comment per PR — re-runs **update** that comment in place (via a hidden `<!-- codex-review-on-pr -->` marker) instead of stacking.
- Capped at 600 s (codex `timeout`) and 15 min (job).
- Truncates output above 60 KB to stay under GitHub's comment size cap; full review remains in the action log.

### Costs and trade-offs

- Every non-draft PR triggers one codex run. With `low` reasoning effort and `service_tier="fast"`, expect ~30-90 s per run.
- Review is **advisory** — not in `required_status_checks`. Promote later if the signal-to-noise ratio justifies it.
- Forgeable: a determined contributor could push code that exploits prompt-injection in comments / fixtures. Treat the review as "second opinion", not authoritative.
- The model still sees whatever it reads from files. Keep `.env`, real fixtures with secrets, and any sensitive content out of the diff.

### Tuning knobs

- Bump reasoning effort to `medium` if reviews miss obvious issues.
- Pin a specific codex CLI version (`@openai/codex@<version>`) once a stable release is known to work in CI; the current `@latest` is a moving target.
- Switch to `xhigh` reasoning effort for security-critical paths via a per-path matrix; current single-job design is intentionally simple.

## Pending items (live state)

| Item | Status | Notes |
|---|---|---|
| Phase 2 gates → `required_status_checks` | **Done (2026-04-30)** | 9 contexts now required: 6 build/test + 3 PR gates. |
| `windows-runtime` job — `zig fmt --check` failing | **Fixed (2026-04-30, PR #9)** | `.gitattributes` forces LF on all platforms. |
| `windows-runtime` job — `Fetch corpus` UnicodeEncodeError | **Fixed (2026-04-30, PR #10)** | Python heredoc reconfigures stdout to utf-8. |
| `windows-runtime` first fully-green run | **Achieved (2026-04-30, PR #10)** | All 14 steps passed: fmt, build × 2, unit + fuzz-smoke, fetch corpus, corpus integration, CLI smoke, dll staging, wheel install + reader smoke, pytest binding suite. Job total 4m 8s. |
| `windows-runtime` → `continue-on-error: false` | Pending | Wait for ≥3 consecutive green runs on `main` before flipping. |
| `windows-runtime` → required status check | Pending | Same condition as above; chain after the flip. |
| `bench` regression CI — pre-existing apt-hyperfine failure | **Fixed (2026-04-30, PR #5)** | Pinned hyperfine 1.18.0 via release `.deb`. |
| `OPENAI_API_KEY` repo secret | **Deferred (2026-04-30)** | Workflow no-ops with a CI warning until set. |
| `Codex review` → `required_status_checks` | **Deferred (2026-04-30)** | Comment-posting flow gives the reviewer signal without merge friction. |
| Phase 5 (coverage / TDAD / mutation) | **Deferred indefinitely** | Tooling rough or absent in Zig ecosystem. |

## Phase 1 — operational note

`main` is now PR-only. To land changes:

```sh
git switch -c <branch>
# ... commits ...
git push -u origin <branch>
gh pr create --fill
# wait for CI green, then:
gh pr merge --squash --delete-branch    # or --rebase
```

Direct `git push origin main` will be rejected. If `enforce_admins` proves too strict during emergencies, soften with:

```sh
gh api -X PATCH repos/laurentfabre/zlsx/branches/main/protection/enforce_admins -F enabled=false
```
