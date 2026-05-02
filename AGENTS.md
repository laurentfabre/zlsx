# AGENTS.md — zlsx

Guide for LLM coding agents (Claude Code, Codex, Cursor) working on this repository. Dense on purpose: every rule below has caused a bug or burned an afternoon at least once.

> **Toolchain.** This project targets **Zig 0.15.2 only**. Verify with `zig version` before doing anything else — Zig's stdlib churns between minor releases and 0.16 changes are not backward compatible (`std.Thread.Mutex` → `std.Io.Mutex`, `std.time.nanoTimestamp` removed, `std.process.Child.run` signature change). If `zig version` doesn't print `0.15.2`, fix the PATH before debugging anything else.

---

## What this project is

zlsx is a fast, mostly-read `.xlsx` library written in Zig 0.15.2. It ships:

- **`zlsx`** — Zig module (reader + writer) for direct consumption from Zig.
- **`zlsx_pkg`** — package-layer module (PartStore, image/chart anchors) usable without the full reader/writer surface.
- **`zlsx`** — CLI binary that streams xlsx rows to stdout as JSONL/TSV/CSV.
- **`zlsx-extract-images`** — standalone CLI driving the package layer for image extraction.
- **`libzlsx`** — C ABI (shared `.dylib`/`.so`/`.dll` + static `.a`) consumed by language bindings.
- **`py-zlsx`** — Python binding over `libzlsx` via `ctypes`. Distributed as a wheel.

zlsx is **stdlib-only**: zero third-party Zig runtime dependencies. `build.zig.zon`'s `.dependencies` is empty by design; adding a dep needs explicit authorization.

---

## Project shape

```
zlsx/
├── build.zig                       # 5 build steps: default, test, test-corpus, fuzz, run
├── build.zig.zon                   # version is the single source of truth (propagated everywhere)
├── src/
│   ├── xlsx.zig                    # reader + the public Zig module root
│   ├── writer.zig                  # greenfield writer
│   ├── c_abi.zig                   # C ABI exports — paired with include/zlsx.h
│   ├── cli.zig                     # `zlsx` CLI: JSONL/TSV/CSV streaming
│   ├── extract_images_main.zig     # `zlsx-extract-images` standalone exe
│   ├── unicode/{casefold,nfc}.zig  # NFC + Unicode case-fold for sheet-name dedup
│   └── formula/tokenizer.zig       # formula tokenizer (loss-preserving printer)
├── pkg/                            # "package layer" — usable without the full reader/writer
│   ├── root.zig                    # @import("zlsx_pkg") root
│   ├── store.zig                   # PartStore + decodeXmlEntities + looksExternal (fuzzed)
│   └── drawings.zig                # image / chart anchor parsing
├── include/zlsx.h                  # PUBLIC C header — must track src/c_abi.zig
├── bindings/python/                # py-zlsx
│   ├── pyproject.toml              # version mirrors build.zig.zon
│   └── zlsx/{__init__.py,_ffi.py,test_basic.py}
├── tests/
│   ├── xlsx_corpus.zig             # reader integration tests (corpus-driven)
│   ├── package_corpus.zig          # package-layer integration tests
│   ├── corpus/                     # fixtures — populated by scripts/fetch_test_corpus.sh
│   ├── fuzz/                       # fuzz corpus seeds
│   └── bench/                      # bench harness vs other xlsx libraries
├── docs/plans/                     # roadmap + per-phase plans (authoritative work queue)
├── scripts/                        # corpus fetch, bench CI, Homebrew tap publish, unicode tables
└── packaging/                      # Homebrew formula tap
```

---

## Build steps

```sh
zig build                            # default — builds CLI, zlsx-extract-images, libzlsx (.dylib + .a)
zig build test                       # all unit + writer + unicode + nfc + formula + cli + c_abi + pkg + corpus
zig build test-corpus                # corpus integration only
zig build fuzz                       # coverage-guided fuzz (Linux x64 only — see notes below)
zig build run -- <args>              # build + run the CLI
zig build -Doptimize=ReleaseSafe     # production-shape build
zig build -Dtarget=aarch64-linux-musl  # cross-compile (compile-only — execution requires QEMU)
zig build -Dsingle-threaded=true     # CLI + C ABI: smp_allocator → page_allocator
zig fmt --check src/                 # formatting check
zig fmt src/                         # auto-format
```

Populate the corpus before `test-corpus` / `test`:

```sh
scripts/fetch_test_corpus.sh         # fetches tests/corpus/*.xlsx
```

---

## C ABI rules (zlsx-specific, load-bearing)

### 1. extern struct layout is forever

`src/c_abi.zig` exports `extern struct` types (`CCell`, `CMergeRange`, `CHyperlink`, `CComment`, `CDataValidation`, `CDateTime`, `CFont`, `CFill`, `CBorderSide`, `CCellBorder`, `CCellAlignment`, …). Once shipped:

- **Reordering or resizing fields is a breaking ABI change.** Add new fields at the end, never in the middle.
- A C ABI change is a **3-file transaction**: `src/c_abi.zig` + `include/zlsx.h` + `bindings/python/zlsx/_ffi.py`. Plus tests in `bindings/python/zlsx/test_basic.py`.

### 2. include/zlsx.h is hand-maintained

The public C header is not auto-generated. Convention:

- Every `pub fn zlsx_*` in `c_abi.zig` has a matching `extern` declaration in `zlsx.h`.
- Every `pub const C*` extern struct has a matching `typedef struct` block.
- Order in the header follows the order in `c_abi.zig` so diffs are reviewable.
- Verify after any ABI edit:

  ```sh
  grep -n "^pub fn zlsx_"   src/c_abi.zig
  grep -n "^pub const C"    src/c_abi.zig
  ```

  Each export should appear in `include/zlsx.h`.

### 3. Bound integers BEFORE ctypes narrowing

ctypes silently truncates on overflow into fixed-width fields → corrupt indices, undefined reads in Zig. At the Python boundary, **bound every integer that crosses into a fixed-width ctypes field before casting**:

```python
if row_idx > UINT32_MAX:
    raise ValueError(f"row_idx={row_idx} exceeds UINT32_MAX")
ffi.row_idx = ctypes.c_uint32(row_idx)   # safe
```

The same rule covers `INT32_MAX` (signed `local_sheet_id`), `409.5` (the row-height ceiling enforced at the ABI), pane / freeze counts, and any other fixed-width field. This pattern is recurring; treat unconditional ctypes assignments at the boundary as a bug.

### 4. Python feature-probe pattern

When new ABI exports land but older `libzlsx` versions are still installed, `_ffi.py` probes for the symbol and `test_basic.py` skips gracefully. **Do not fail-import** — the binding has to remain usable against an older dylib.

### 5. Mach-O install_name headerpad

`build.zig` sets `dylib.headerpad_max_install_names = true;`. This reserves Mach-O load-command space so packaging tools can rewrite the dylib's `install_name` post-install. No-op on Linux/Windows; **don't remove it** — it's not cosmetic.

---

## Build / module quirks to know

### Two fuzz binaries, Linux x64 only

`zig build fuzz` runs:

- `unit_fuzz_tests` — fuzz targets in `src/xlsx.zig` (reader / parser surface).
- `package_fuzz_tests` — fuzz targets in `pkg/store.zig` (`decodeXmlEntities`, `looksExternal`).

Both require **Linux x64**; macOS and Windows are upstream-broken in Zig 0.15.2. Don't claim "fuzzing works" from a macOS dev box — it doesn't.

### Three-module collision

`build.zig` documents a Zig 0.15.2 limitation: `cli_mod`, `zlsx_pkg`, and `writer` cannot coexist in one compilation, because every file that `@import("writer")`s ends up claimed by both writer's tree and zlsx_pkg's tree. That's why `zlsx-extract-images` is a separate exe rather than a CLI subcommand. Don't try to merge it back without re-checking the constraint.

### `-Dsingle-threaded` swaps the allocator

More than a thread-disabling toggle: `smp_allocator` (lock-striped) is swapped for `page_allocator` in the CLI and C ABI. Bench numbers will diverge. Don't blindly compare single-threaded vs default builds.

### Version is a single source of truth

```zig
const pkg_version: []const u8 = @import("build.zig.zon").version;
```

Propagated to:
- `zlsx_version_string()` C ABI export
- `build_options.version` for downstream Zig code
- Homebrew formula via `scripts/publish_homebrew_tap.sh`
- `bindings/python/pyproject.toml` (manual mirror — bump in lockstep)

A version bump touches multiple files, but only `build.zig.zon` is the source.

### Roadmap is the authoritative work queue

`docs/plans/post-0.2.9-roadmap.md` is what next-PR tooling reads. Per-phase plans live in `docs/plans/{streaming-sst,cell-mutate,load-modify-save,structural-edits}.md`. Don't pick "the next thing" from intuition — read the roadmap.

---

## Defensive-programming patterns

### Allocator discipline
- Single GPA in `main()`, threaded through every function that allocates. No globals.
- Functions that allocate take `allocator: std.mem.Allocator` as the first parameter (after `self` if a method).
- `defer` for success-path cleanup, `errdefer` for the error path. They're complements, not synonyms.
- For tests: `std.testing.allocator` (catches leaks at exit). For multi-allocation paths: wrap in `std.testing.checkAllAllocationFailures` so every allocation site is force-failed once and the recovery path is exercised.

### Error handling
- Explicit error sets in public APIs. `error{Foo, Bar}!T`, never `anyerror!T`.
- `try` for propagation. `catch |err| switch (err) { ... }` for exhaustive handling at trust boundaries.
- *Inside* the trust boundary, prefer `assert` / `unreachable` over error propagation. *At* the boundary, exhaustive switch.
- Never silently swallow errors. `_ = foo() catch {};` requires a comment justifying it.
- `catch unreachable` only when you've genuinely proven the error case can't happen.

### Type-driven design
- Tagged unions (`union(enum)`) for state machines with more than two valid configurations. Exhaustive `switch` makes adding a state a compile error at every transition site.
- Optionals (`?T`) for "may be absent." Never use a sentinel value. Hard unwrap (`.?`, `orelse unreachable`) only when the invariant is enforced *somewhere visible* — and name the invariant in a comment.
- Pointer flavors are load-bearing: `*T` (one), `[]T` (slice + length), `[*]T` (many, no length — C interop only), `[*:0]T` (sentinel). The type names the risk.
- Distinguish `index`, `count`, `size_bytes` in variable names.

### Arithmetic
- Default `+`, `-`, `*`, `<<` **trap** on overflow in Debug/ReleaseSafe; UB in ReleaseFast.
- Use wrapping (`+%`) only for legitimate hashing/checksumming. Saturating (`+|`) for clamping. `std.math.add` for `error.Overflow`-returning checked arithmetic.
- Use `@intCast` (traps) vs `@truncate` (intentional bit truncation) **explicitly**. Don't let the compiler infer.
- Use `@divExact`, `@divFloor`, `@divTrunc` instead of `/` whenever the rounding matters.

### Assertions
- ≥2 assertions per non-trivial function — preconditions, postconditions, invariants.
- Split compound assertions: `assert(a); assert(b);` not `assert(a and b);` so the failing line points at the exact clause.
- `unreachable` on cold branches turns dead code into a tripwire.

### Comments
- Comments only when the WHY is non-obvious. A hidden invariant, an external constraint, surprising behavior. Never restate what the code does.

---

## Testing

### Layer 1: inline `test` blocks
Most tests live in the same file as the code under test, as `test "<name>" { ... }` blocks. `zig build test` discovers and runs every test in the build graph.

`zig build test` only runs tests reachable from the build's test root. **A new file's tests run only if the file is imported by something already in the test graph.** When adding a new file, also `@import` it from a module already wired in `build.zig`, otherwise the TDD gate is silently bypassed.

### Layer 2: corpus integration
`tests/xlsx_corpus.zig` and `tests/package_corpus.zig` walk every fixture in `tests/corpus/`. Run with `zig build test-corpus`. The corpus must be populated first — see `scripts/fetch_test_corpus.sh`.

### Layer 3: fuzz
Built-in via `std.testing.fuzz(...)` inside a `test` block; run with `zig build fuzz`. Linux x64 only. Use on parsers, deserializers, anything that takes untrusted input.

### Layer 4: Python binding
`bindings/python/zlsx/test_basic.py` exercises the C ABI through ctypes. Add a feature-probe + skip when introducing a new ABI export.

### Layer 5: cross-target
For portable code, periodic `zig build -Dtarget=aarch64-linux-musl` (compile-only) keeps host-only assumptions honest.

---

## Stale-API guard — do NOT emit

These were correct in older Zig but are stale as of 0.15.x. If you're tempted to write them, don't:

| Stale | Replacement |
|---|---|
| `std.io.getStdOut().writer()` | `std.fs.File.stdout().deprecatedWriter()` (or new `std.Io.Writer`) |
| `std.io.getStdErr().writer()` | `std.fs.File.stderr().deprecatedWriter()` |
| `std.fmt.format("{}", .{x})` on a custom-`format` type | `{f}` is now required |
| `BoundedArray`, `LinearFifo` | `ArrayListUnmanaged` over a stack/static buffer |
| `async fn`, `await`, `suspend`, `resume` keywords | Removed; asynchrony via `std.Io` interface |
| `usingnamespace` | Removed |
| `std.ArrayList(T).init(allocator)` (managed) as default | `std.ArrayList(T).empty` + explicit allocator on `.append(allocator, x)` / `.deinit(allocator)` |
| `std.testing.fuzzInput()` | `std.testing.fuzz(ctx, testOne, .{})` |
| Lossy int→float coercion (`const x: f32 = 1_234_567_890;`) | Compile error in 0.15.x |
| Bare `{}` format on a type with a `format` method | Compile error; must use `{f}` |
| `std.heap.GeneralPurposeAllocator(.{})` deinit returning `bool` | Returns `Check{ .ok, .leak }` enum |

---

## Common gotchas

### Wrong zig binary
**Symptom**: cryptic stdlib errors, missing functions, signature mismatches that don't match what this guide describes.
**Cause**: shell PATH selects a different Zig than 0.15.2.
**Fix**: `zig version`. If it isn't `0.15.2`, fix the PATH before debugging anything else.

### Build cache lying
`rm -rf .zig-cache zig-out` clears the per-project cache. Only do this when build behavior diverges from `build.zig` reality — not as a reflex.

### Allocator leaks
**Symptom**: `general_purpose_allocator.deinit()` reports leaks at process exit.
**Cause**: an allocation path forgot its `defer allocator.free(...)` or `errdefer`.
**Fix**: read the leak's stack trace; trace back to the allocation site; pair it with cleanup. Never ignore — GPA in Debug mode is the cheapest leak detector you'll get.

### `errdefer` ordering surprises
`defer` always runs; `errdefer` runs only on error. Prefer the explicit pattern — `try + errdefer` for the happy-path-with-cleanup case. Less clever, more readable than mixing with manual `catch` blocks.

### Cross-target stdlib drift
Use `std.fs.path`, `@import("builtin").os.tag` for OS-conditional branches, and the abstractions in `std.fs.Dir`. If the abstraction doesn't exist, write a per-OS branch and TEST it.

### `build.zig.zon` hash drift
URL deps fail with `error: hash mismatch` when their tarball changes; replace `.hash = "..."` with the value Zig prints in the error message. `.path` deps have no `.hash` field — they're resolved at build time. zlsx currently has no URL deps; this is informational.

---

## Workflow conventions

### Worktrees, one per concurrent PR

`main` is PR-only. For any branch you'd run an agent against, prefer a sibling worktree with its own build cache:

```sh
scripts/wt-new feat/streaming-sst                # creates ../zlsx-feat-streaming-sst/
scripts/wt-new fix/header-pad origin/release-0.3 # custom base ref
```

Each worktree owns its own `.zig-cache/`, so parallel agent sessions never thrash the shared cache. Tear down with `git worktree remove <path>`. Convention only — no enforcement.

### Commit `Agent:` trailer

Append a soft trailer naming the agent that authored the commit:

```
Agent: human
Agent: claude-opus-4-7
Agent: codex-gpt-5
```

Installed as a `commit-msg` hook (run `scripts/install-hooks.sh` once after clone). Soft-check only — missing trailer warns but does not block.

### CI gates and branch protection

`main` is PR-only with 9 required status checks (`test/macos-14`, `test/ubuntu-22.04`, 4 `cross/*` targets, plus 3 PR gates: `Test-presence check`, `C ABI 3-file transaction`, `Monotonic test count`). Linear history, no force-push, no deletions, dismiss-stale-approvals on. The `tdad-map` job posts an updated-in-place comment on every PR listing inline tests + related corpus / fuzz / binding surfaces affected by the change. Rationale and per-gate trade-offs live in `docs/enforcement-plan.md`.

## Process when making changes

1. **Confirm the toolchain**: `zig version` should print `0.15.2`. If not, fix that before anything else.
2. **Read the roadmap**: `docs/plans/post-0.2.9-roadmap.md` and any per-phase plan in `docs/plans/`.
3. **Write the test first**: add a `test "<name>"` block, watch it fail with `zig build test`, then implement.
4. **Verify the file is in the test graph**: a new file's tests don't run unless the file is `@import`-ed from something already wired in `build.zig`.
5. **Allocator threading**: if your function allocates, take `allocator: std.mem.Allocator`. No globals.
6. **For C ABI changes**: update `src/c_abi.zig` + `include/zlsx.h` + `bindings/python/zlsx/_ffi.py` in the same change. Add a feature probe in `_ffi.py` and a skip in `test_basic.py` so older dylibs still pass.
7. **For Python boundary code**: bound every integer before ctypes narrowing.
8. **Format**: `zig fmt --check src/` before committing.
9. **Cross-target sanity**: for portable code, `zig build -Dtarget=aarch64-linux-musl` keeps host-only assumptions honest.

---

## Quick discovery

```sh
cat build.zig.zon                         # version, deps
zig build --help                          # auto-generated step list
git log --oneline -20                     # recent commit thread
grep -n "^pub fn zlsx_" src/c_abi.zig    # exported C ABI functions
grep -n "^pub const C"  src/c_abi.zig    # exported extern structs
ls docs/plans/                            # roadmap + phase plans
```

---

## Anti-patterns

- Using a non-0.15.2 zig — every error you see will be confusing.
- Adding a third-party Zig runtime dependency without authorization. zlsx is stdlib-only.
- Reordering or inserting fields into a published `extern struct` (breaks the C ABI).
- Updating `src/c_abi.zig` without updating `include/zlsx.h` and `bindings/python/zlsx/_ffi.py` in the same change.
- Assigning a Python integer to a fixed-width ctypes field without a range check (silent truncation).
- Failing to handle older dylibs in `_ffi.py` — the binding must remain usable when probing for newly-added symbols.
- Silently swallowing errors with `catch unreachable` when the error is recoverable.
- Removing `headerpad_max_install_names` from `build.zig` (breaks Homebrew dylib relocation).
- Running `rm -rf .zig-cache zig-out` reflexively. Diagnose first.
- Skipping `zig fmt`. The check is free; drift is annoying to clean up later.
- Trusting Zig examples without checking the version they target. Stdlib churns; what worked on 0.13 may not on 0.15, and 0.16 already breaks 0.15.
