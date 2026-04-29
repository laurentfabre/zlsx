## Summary

<!-- 1-3 sentences: what changes and why. -->

## Scope

**In scope:**

-

**Out of scope:**

-

## Test plan

<!-- Concrete: which `zig build` step exercises the change, which test file(s), which fixtures. -->

- [ ] `zig build test` passes locally
- [ ] `zig build test-corpus` passes locally (if reader/package layer touched)
- [ ] `zig fmt --check src/ tests/ build.zig` clean

## C ABI impact

- [ ] No C ABI surface changed.
- [ ] **C ABI changed** — if checked, confirm the 3-file transaction:
  - [ ] `src/c_abi.zig` updated
  - [ ] `include/zlsx.h` updated (matching `pub fn` and `extern struct` order)
  - [ ] `bindings/python/zlsx/_ffi.py` updated
  - [ ] `bindings/python/zlsx/test_basic.py` exercises the new surface
  - [ ] Older-dylib feature-probe / skip in place
  - [ ] Python-side integer arguments are bounded before ctypes narrowing

## Roadmap link

<!-- Item from docs/plans/post-0.2.9-roadmap.md or per-phase plan. -->

## Agent attribution (optional)

- Author agent: _e.g. claude-opus-4-7 / human / codex-gpt-5_
- Reviewer agent: _e.g. /zig-defensive / codex review / human_
