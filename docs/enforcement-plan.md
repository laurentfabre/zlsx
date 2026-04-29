# Enforcement plan — worktrees, subagents, TDD, merge guards, review

Status as of 2026-04-30. Single-author public repo (`laurentfabre/zlsx`), Zig 0.15.2.

## Status table

| Phase | Description | Status |
|---|---|---|
| 1 | Free wins: branch protection, repo merge settings, CODEOWNERS, PR template | **Done (2026-04-30)** |
| 2 | TDD CI gates: test-presence, C-ABI 3-file-transaction, monotonic test count | Not started |
| 3 | Worktree + subagent conventions: helper script, commit-msg trailer, PR template fields | Not started |
| 4 | Agent-as-reviewer CI job (codex-review-on-PR) | Not started |
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
