# AGENTS.md — zlsx

Guide for LLM coding agents (Claude Code, Codex, Cursor) working on this repository. Dense on purpose: every rule below has caused a bug or burned an afternoon at least once.

> **Toolchain.** This project targets **Zig 0.16.0 only**. Verify with `zig version` before doing anything else — Zig's stdlib churns hard between minor releases, and 0.16 is a large break from 0.15: I/O is no longer ambient (`std.fs` moved under `std.Io` and every call takes an `io`), `std.time` lost all its functions, and `main` now receives a `std.process.Init`. See the stale-API guard below before writing any stdlib call. If `zig version` doesn't print `0.16.0`, fix the PATH before debugging anything else.

---

## What this project is

zlsx is a fast, mostly-read `.xlsx` library written in Zig 0.16.0. It ships:

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
│   ├── zlsx/{__init__.py,_ffi.py}
│   └── tests/{test_basic,test_embeddings,test_spark_core}.py
├── tests/
│   ├── xlsx_corpus.zig             # reader integration tests (corpus-driven)
│   ├── package_corpus.zig          # package-layer integration tests
│   ├── corpus/                     # fixtures — populated by scripts/fetch_test_corpus.sh
│   ├── fuzz/                       # fuzz corpus seeds
│   └── bench/                      # bench harness vs other xlsx libraries
├── docs/ROADMAP.md                 # the plan of record (authoritative work queue)
├── docs/plans/                     # live per-phase plans (archive/ holds shipped ones)
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
zig build -Dsingle-threaded=true     # CLI only: smp_allocator → page_allocator (C ABI stays multi-threaded, R9-12)
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
- A C ABI change is a **3-file transaction**: `src/c_abi.zig` + `include/zlsx.h` + `bindings/python/zlsx/_ffi.py`. Plus tests in `bindings/python/tests/test_basic.py`.

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

**Coverage-guided fuzzing works**, but only because `vendor/zig-test-runner/` exists.

Zig 0.16.0 ships two bugs that make `-ffuzz` builds unusable out of the box: its own `lib/compiler/test_runner.zig` fails to compile (`writeStackTrace` receives a `*builtin.StackTrace` where a `*const debug.StackTrace` is required), and even once that is fixed the runner calls the fuzzer ABI during the test-discovery pass, when `fuzzer.test_i` is still `undefined` — segfaulting on every target that supplies a seed corpus. `vendor/zig-test-runner/` is upstream's runner with those two hunks patched, wired into the two fuzz binaries only via `.test_runner` in `build.zig`. **Read `vendor/zig-test-runner/README.md` before touching it, and delete the whole directory as soon as a Zig release fuzzes without it.**

The correct invocation is:

```sh
zig build fuzz --fuzz --webui=127.0.0.1:0
```

`fuzz` alone is the *step name*; `--fuzz` is the *flag* that enables coverage-guided mode, and omitting it makes the build runner panic on a null `fuzz_context`. A healthy session runs until killed, so wrap it in `timeout` — exit 124/143 means "ran the full session", not failure.

Verified locally: a 100 s session completes with 226/226 tests and writes fuzzer state under `.zig-cache/v/`.

Linux x64 remains the only supported target; macOS (Mach-O `addEntryPoint`) and Windows (shared-memory + COFF/PE debug info) are separately broken upstream.

### Three-module collision

**Resolved on 0.16 — the note below is history, not a live constraint.**

Under Zig 0.15.2, `cli_mod`, `zlsx_pkg` and `writer` could not coexist in one compilation: every file that `@import("writer")`ed ended up claimed by both writer's tree and zlsx_pkg's tree. That is why `zlsx-extract-images` ships as a separate exe rather than a CLI subcommand.

The 0.16 migration retested it directly — adding `cli_mod.addImport("writer", writer_mod)` on top of the existing `zlsx` + `zlsx_pkg` imports builds clean and keeps 1029/1029 tests green. Merging `zlsx-extract-images` back into the CLI is therefore possible now. It has deliberately **not** been done: dropping a shipped binary is a user-visible packaging change and belongs to whoever owns that call, not to a build-graph tidy-up.

What downstream consumers actually depend on — importing the public modules together — is now a build gate, not a claim: `tests/consumer/` is a standalone package with a path dependency on the repo root that writes a workbook with `zlsx.Writer`, reads it with `zlsx.Book`, mutates it through `zlsx_pkg.Editor`, re-reads to verify, and then (M5d3) drives `zlsx_recalc.writerSaveWithRecalc` — all three public modules in one compilation. Run it with:

```sh
cd tests/consumer && zig build && ./zig-out/bin/consumer /tmp/in.xlsx /tmp/out.xlsx
```

The third module is the reason the gate lives outside the repo's own `build.zig`: `assertAcyclicModules` walks modules zlsx constructed, while a downstream package resolves them through `b.dependency(...)`, and the composition only compiles if the `zlsx` reached that way is the same module object `zlsx_pkg` and `zlsx_recalc` were built against.

### `-Dsingle-threaded` swaps the allocator — CLI only

More than a thread-disabling toggle: `smp_allocator` (lock-striped) is swapped for `page_allocator` in the CLI. Bench numbers will diverge. Don't blindly compare single-threaded vs default builds.

The C ABI is exempt since M9a1 (decision R9-12): `zlsx_cancel_token_trigger` is documented callable from any thread, and `-fsingle-threaded` lowers atomics to plain ops — which would silently break the token in exactly the supported configuration. `build.zig` hard-sets the module multi-threaded and `src/c_abi.zig` carries a comptime assertion refusing anything else; the `-Dsingle-threaded=true` CI lane still compiles both shapes from one invocation (CLI single-threaded, ABI multi-threaded). The CLI keeps its signal-safe `flag` token kind.

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

`docs/ROADMAP.md` is the plan of record and what next-PR tooling reads. It carries the status table, the dependency graph, and the candidate follow-ups. Live per-phase plans live in `docs/plans/`; plans for shipped work are in `docs/plans/archive/`, kept for per-PR traceability but **not** a work queue. Don't pick "the next thing" from intuition — read the roadmap.

`goal_formula.md` (repo root) is normative for the D1 formula engine and is cited by section number from 33 source files. It is shipped, not a queue.

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
`bindings/python/tests/test_basic.py` exercises the C ABI through ctypes. Add a feature-probe + skip when introducing a new ABI export.

### Layer 5: cross-target
For portable code, periodic `zig build -Dtarget=aarch64-linux-musl` (compile-only) keeps host-only assumptions honest.

---

## Stale-API guard — do NOT emit

These were correct in older Zig but are stale as of **0.16**. If you're tempted to write them, don't:

**The 0.16 headline: I/O is no longer ambient.** Every filesystem, stdio,
clock, and child-process call takes an `io: std.Io`. Programs receive one
(plus argv, a gpa and an arena) through `std.process.Init`, the parameter
`main` now takes. Library code should accept an `io` parameter rather
than manufacturing one; long-lived owners (`Book`, `PartStore`, the C ABI
handles) retain the `Io` they were opened with because they re-touch the
file after the opening call returns.

| Stale | Replacement |
|---|---|
| `std.fs.cwd()` | `std.Io.Dir.cwd()`, and every method takes `io` first: `.openFile(io, path, .{})` |
| `std.fs.File` / `std.fs.File.Reader` | `std.Io.File` / `std.Io.File.Reader` |
| `file.readAll(buf)` / `file.seekTo(n)` | Removed. Make a `file.reader(io, &.{})` and use `reader.seekTo(n)` + `reader.interface.readSliceAll(buf)` |
| `file.writeAll(bytes)` after `createFile` | `dir.writeFile(io, .{ .sub_path = p, .data = bytes })` |
| `dir.realpathAlloc(alloc, p)` | `dir.realPathFileAlloc(io, p, alloc)` — renamed **and** reordered |
| `std.fs.Dir.atomicFile` | Removed. Use `pkg/atomic_file.zig`'s `AtomicFile` |
| `std.fs.File.stdout().writer(&buf)` | `std.Io.File.stdout().writer(io, &buf)` |
| `std.process.argsAlloc(alloc)` / `argsFree` | `init.minimal.args.toSlice(init.arena.allocator())`; there is no free — the arena owns it |
| `std.process.getEnvVarOwned` | Gone. A test binary's environ is **empty**; read env in `build.zig` via `b.graph.environ_map` and pass values down as build options |
| `std.process.Child.init` + `child.spawn()` | `std.process.spawn(io, .{ .argv = …, .stdout = .pipe })`; `Term` tags are lowercase (`.exited`) |
| `std.time.nanoTimestamp` / `milliTimestamp` / `std.time.Timer` | All gone — `std.time` has no functions. Use `std.Io.Clock.now(.awake, io).nanoseconds` (an `i96`) |
| `std.crypto.random` | Gone. Prefer a design that needs no entropy (e.g. exclusive-create probing) |
| `std.mem.trimRight` / `trimLeft` | `std.mem.trimEnd` / `trimStart` |
| `std.meta.intToEnum` (error union) | `std.enums.fromInt` — returns an **optional**, so `orelse` not `catch` |
| `std.io.fixedBufferStream(&buf)` + `.getWritten()` | `std.Io.Writer.fixed(&buf)` + `.buffered()` |
| `std.heap.GeneralPurposeAllocator` | `std.heap.DebugAllocator`, or take `init.gpa` |
| `std.ArrayListUnmanaged(T) = .{}` / `T{}` | `= .empty` |
| `fn (context, input: []const u8)` fuzz target | `fn (context, smith: *std.testing.Smith)`; draw bytes with `smith.slice(&buf)` |
| POSIX signal handler typed `fn (i32)` | Derive it: `@typeInfo(@typeInfo(std.posix.Sigaction.handler_fn).pointer.child).@"fn".params[0].type.?` — 0.16 types it per platform |
| `std.io.getStdOut().writer()` | `std.Io.File.stdout().writer(io, &buf)` |
| `std.fmt.format("{}", .{x})` on a custom-`format` type | `{f}` is now required |
| `BoundedArray`, `LinearFifo` | `ArrayListUnmanaged` over a stack/static buffer |
| `async fn`, `await`, `suspend`, `resume` keywords | Removed; asynchrony via `std.Io` interface |
| `usingnamespace` | Removed |
| Lossy int→float coercion (`const x: f32 = 1_234_567_890;`) | Compile error |
| Bare `{}` format on a type with a `format` method | Compile error; must use `{f}` |

---

## Common gotchas

### Wrong zig binary
**Symptom**: cryptic stdlib errors, missing functions, signature mismatches that don't match what this guide describes.
**Cause**: shell PATH selects a different Zig than 0.16.0.
**Fix**: `zig version`. If it isn't `0.16.0`, fix the PATH before debugging anything else.

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

1. **Confirm the toolchain**: `zig version` should print `0.16.0`. If not, fix that before anything else.
2. **Read the roadmap**: `docs/ROADMAP.md` and any live per-phase plan in `docs/plans/`.
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
cat docs/ROADMAP.md                       # the plan of record
ls docs/plans/                            # live phase plans (archive/ = shipped)
```

---

## Anti-patterns

- Using a non-0.16.0 zig — every error you see will be confusing.
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
