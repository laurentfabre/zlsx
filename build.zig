const std = @import("std");

// Single source of truth for the package version. Everything else
// (C ABI `zlsx_version_string()`, release tarball names, Homebrew
// formula) derives from this.
const pkg_version: []const u8 = @import("build.zig.zon").version;

pub fn build(b: *std.Build) void {
    const target = b.standardTargetOptions(.{});
    const optimize = b.standardOptimizeOption(.{});
    const single_threaded = b.option(bool, "single-threaded", "Build the CLI and C ABI with -fsingle-threaded (smp_allocator is swapped for page_allocator)");

    // Options module exposes `version` to downstream Zig code via
    // `@import("build_options").version`. We share one instance across
    // the CLI and C ABI modules so both report the same string.
    const build_options = b.addOptions();
    build_options.addOption([]const u8, "version", pkg_version);
    const build_options_mod = build_options.createModule();

    // Public module. Consumers add zlsx to their build.zig.zon as a
    // path or git dependency, then `@import("zlsx")`.
    const zlsx_mod = b.addModule("zlsx", .{
        .root_source_file = b.path("src/xlsx.zig"),
        .target = target,
        .optimize = optimize,
    });

    // Unit tests (embedded in src/xlsx.zig, including the fuzz suite).
    const unit_mod = b.createModule(.{
        .root_source_file = b.path("src/xlsx.zig"),
        .target = target,
        .optimize = optimize,
    });
    const unit_tests = b.addTest(.{ .root_module = unit_mod });
    const test_step = b.step("test", "Run zlsx unit + fuzz-smoke tests");
    test_step.dependOn(&b.addRunArtifact(unit_tests).step);

    // B-fuzz: coverage-guided fuzz step. Runs the same test binary
    // as `zig build test` but with -ffuzz instrumentation, so any
    // test that calls `std.testing.fuzz(...)` becomes a
    // coverage-guided target. Linux x64 is the production platform
    // per the roadmap; macOS/Windows are gated on Zig upstream
    // fixes. Use `zig build fuzz` to launch.
    const unit_fuzz_mod = b.createModule(.{
        .root_source_file = b.path("src/xlsx.zig"),
        .target = target,
        .optimize = optimize,
        // The `fuzz: ?bool` flag on Module flips the compile to
        // emit -ffuzz instrumentation. `addTest` doesn't expose
        // this directly in Zig 0.15.2 — set it on the module.
        .fuzz = true,
    });
    const unit_fuzz_tests = b.addTest(.{ .root_module = unit_fuzz_mod });
    const fuzz_step = b.step("fuzz", "Run coverage-guided fuzz targets (Linux x64; macOS/Windows broken upstream)");
    fuzz_step.dependOn(&b.addRunArtifact(unit_fuzz_tests).step);

    // Package-layer fuzz module is wired further down, after
    // writer_mod is created — see the `package_fuzz_mod` block
    // near the package_mod / package_drawings_tests_mod section.

    // Integration tests: tests/xlsx_corpus.zig, fed by
    // scripts/fetch_test_corpus.sh into tests/corpus/.
    const corpus_mod = b.createModule(.{
        .root_source_file = b.path("tests/xlsx_corpus.zig"),
        .target = target,
        .optimize = optimize,
    });
    corpus_mod.addImport("zlsx", zlsx_mod);
    const corpus_tests = b.addTest(.{ .root_module = corpus_mod });
    const corpus_step = b.step("test-corpus", "Run integration tests against tests/corpus/*.xlsx");
    corpus_step.dependOn(&b.addRunArtifact(corpus_tests).step);

    // Package-layer corpus sweep: walks every fixture through
    // PartStore + imageAnchors + chartAnchors. Wired into both the
    // dedicated `test-corpus` step and the default `test` step so
    // the package layer's robustness is exercised on every CI run.
    // The module itself is created further down once `package_mod`
    // exists; the wiring is appended there.

    // CLI: `zlsx` binary, streams xlsx rows to stdout in JSONL / TSV / CSV.
    // `zig build` (default step) installs it at zig-out/bin/zlsx.
    const cli_mod = b.createModule(.{
        .root_source_file = b.path("src/cli.zig"),
        .target = target,
        .optimize = optimize,
        .single_threaded = single_threaded,
    });
    cli_mod.addImport("build_options", build_options_mod);
    // After the B2 iter-er-0 Editor relocation, cli reaches Editor
    // through zlsx_pkg and the rest of xlsx via the named `zlsx` dep
    // (no more relative `@import("xlsx.zig")` / `@import("writer.zig")`,
    // so the iter-wb-3 module-graph collision no longer fires).
    cli_mod.addImport("zlsx", zlsx_mod);
    cli_mod.addImport("zlsx_pkg", package_mod);
    const cli_exe = b.addExecutable(.{ .name = "zlsx", .root_module = cli_mod });
    b.installArtifact(cli_exe);

    const run_cli = b.addRunArtifact(cli_exe);
    if (b.args) |args| run_cli.addArgs(args);
    const run_step = b.step("run", "Build and run the zlsx CLI (args after --)");
    run_step.dependOn(&run_cli.step);

    // CLI unit tests (colLetter, JSON/CSV escapers, arg parser).
    const cli_tests = b.addTest(.{ .root_module = cli_mod });
    test_step.dependOn(&b.addRunArtifact(cli_tests).step);

    // Writer unit tests (MVP: round-trip via the reader).
    const writer_mod = b.createModule(.{
        .root_source_file = b.path("src/writer.zig"),
        .target = target,
        .optimize = optimize,
    });
    const writer_tests = b.addTest(.{ .root_module = writer_mod });
    test_step.dependOn(&b.addRunArtifact(writer_tests).step);

    // Unicode case-fold + NFC module tests (A1: Excel sheet-name
    // dedup, wired into validateSheetName + Editor.isSheetNameTaken).
    const unicode_mod = b.createModule(.{
        .root_source_file = b.path("src/unicode/casefold.zig"),
        .target = target,
        .optimize = optimize,
    });
    const unicode_tests = b.addTest(.{ .root_module = unicode_mod });
    test_step.dependOn(&b.addRunArtifact(unicode_tests).step);

    // A1 phase 3: standalone NFC tests (the casefold module imports
    // nfc.zig but has its own tests; this catches NFC bugs that
    // wouldn't surface through the casefold compositional path).
    const nfc_mod = b.createModule(.{
        .root_source_file = b.path("src/unicode/nfc.zig"),
        .target = target,
        .optimize = optimize,
    });
    const nfc_tests = b.addTest(.{ .root_module = nfc_mod });
    test_step.dependOn(&b.addRunArtifact(nfc_tests).step);

    // C1 milestone 1: formula tokenizer + loss-preserving printer.
    // Independent module (no cross-deps yet) — the rewriter that
    // depends on this lands in later C1 iterations.
    const formula_tokenizer_mod = b.createModule(.{
        .root_source_file = b.path("src/formula/tokenizer.zig"),
        .target = target,
        .optimize = optimize,
    });
    const formula_tokenizer_tests = b.addTest(.{ .root_module = formula_tokenizer_mod });
    test_step.dependOn(&b.addRunArtifact(formula_tokenizer_tests).step);

    // C1 milestone 2 (iter 1): pure-function A1 cell-formula
    // rewriter. Imports the M1 tokenizer via a relative path inside
    // src/formula/, so this module's package dir matches the
    // tokenizer's. Sibling of formula_tokenizer_mod with its own
    // test target — the rewriter has no other cross-deps yet
    // (Workbook.rewriteReferences wiring lands in a later iter).
    const formula_rewriter_mod = b.createModule(.{
        .root_source_file = b.path("src/formula/rewriter.zig"),
        .target = target,
        .optimize = optimize,
    });
    const formula_rewriter_tests = b.addTest(.{ .root_module = formula_rewriter_mod });
    test_step.dependOn(&b.addRunArtifact(formula_rewriter_tests).step);

    // ─── Package layer (B0 + C2a) ───────────────────────────────
    //
    // Public root: pkg/root.zig re-exports PartStore /
    // ImageAnchor / etc. Downstream consumers add this to their
    // build.zig.zon as `zlsx_pkg` and `@import("zlsx_pkg").PartStore`.
    // Lives alongside the existing `zlsx` module rather than under
    // it because the package layer is explicitly meant to be usable
    // WITHOUT pulling the full reader/writer surface (e.g. for
    // image-extraction tools that just want raw bytes).
    const package_mod = b.addModule("zlsx_pkg", .{
        .root_source_file = b.path("pkg/root.zig"),
        .target = target,
        .optimize = optimize,
    });
    package_mod.addImport("zlsx", zlsx_mod);

    // C2a: standalone `zlsx-extract-images` binary that drives the
    // package layer (PartStore + imageParts) without going through
    // Editor / Book. Shipped as a separate exe rather than a CLI
    // subcommand because cli_mod + zlsx_pkg + writer can't coexist
    // in one compilation under Zig 0.15.2 (every file that
    // `@import("writer")`s ends up claimed by both writer's tree
    // and zlsx_pkg's tree). The standalone binary's module sees
    // zlsx_pkg + writer in isolation, so no collision.
    const extract_images_mod = b.createModule(.{
        .root_source_file = b.path("src/extract_images_main.zig"),
        .target = target,
        .optimize = optimize,
    });
    extract_images_mod.addImport("zlsx_pkg", package_mod);
    const extract_images_exe = b.addExecutable(.{
        .name = "zlsx-extract-images",
        .root_module = extract_images_mod,
    });
    b.installArtifact(extract_images_exe);

    // Per-source-file test targets so each module gets its own test
    // binary (matches the rest of build.zig's pattern).
    const package_store_tests_mod = b.createModule(.{
        .root_source_file = b.path("pkg/store.zig"),
        .target = target,
        .optimize = optimize,
    });
    package_store_tests_mod.addImport("zlsx", zlsx_mod);
    const package_store_tests = b.addTest(.{ .root_module = package_store_tests_mod });
    test_step.dependOn(&b.addRunArtifact(package_store_tests).step);

    const package_drawings_tests_mod = b.createModule(.{
        .root_source_file = b.path("pkg/drawings.zig"),
        .target = target,
        .optimize = optimize,
    });
    package_drawings_tests_mod.addImport("zlsx", zlsx_mod);
    const package_drawings_tests = b.addTest(.{ .root_module = package_drawings_tests_mod });
    test_step.dependOn(&b.addRunArtifact(package_drawings_tests).step);

    // pkg/typed_parts/* — typed-overlay parsers for known OOXML parts
    // (B1 iter-wb-1). One test binary covers all five children via
    // `pkg/typed_parts/root.zig` re-exports. Stdlib-only, no writer
    // import (typed-overlay parsers don't need it).
    const package_typed_parts_tests_mod = b.createModule(.{
        .root_source_file = b.path("pkg/typed_parts/root.zig"),
        .target = target,
        .optimize = optimize,
    });
    const package_typed_parts_tests = b.addTest(.{ .root_module = package_typed_parts_tests_mod });
    test_step.dependOn(&b.addRunArtifact(package_typed_parts_tests).step);

    // pkg/workbook.zig — Workbook + Worksheet typed-overlay roots
    // (B1 iter-wb-2). Composes PartStore + typed_parts into a single
    // model surface; read-only in this iter, mutation lands iter-wb-4.
    // Inline tests open small fixtures directly, so this module needs
    // tests/corpus/* to exist at runtime — the corpus_step covers
    // missing-fixture skipping per scripts/fetch_test_corpus.sh.
    const package_workbook_tests_mod = b.createModule(.{
        .root_source_file = b.path("pkg/workbook.zig"),
        .target = target,
        .optimize = optimize,
    });
    package_workbook_tests_mod.addImport("zlsx", zlsx_mod);
    const package_workbook_tests = b.addTest(.{ .root_module = package_workbook_tests_mod });
    test_step.dependOn(&b.addRunArtifact(package_workbook_tests).step);

    // Package-layer fuzz module: pkg/store.zig hosts fuzz targets
    // for decodeXmlEntities + looksExternal. Same fuzz=true flag
    // as src/xlsx.zig's unit_fuzz_mod, separate binary because
    // the writer import + arena profile differs from the reader.
    const package_fuzz_mod = b.createModule(.{
        .root_source_file = b.path("pkg/store.zig"),
        .target = target,
        .optimize = optimize,
        .fuzz = true,
    });
    package_fuzz_mod.addImport("zlsx", zlsx_mod);
    const package_fuzz_tests = b.addTest(.{ .root_module = package_fuzz_mod });
    fuzz_step.dependOn(&b.addRunArtifact(package_fuzz_tests).step);

    // tests/package_corpus.zig — corpus-level integration test for
    // the package layer. Imports `zlsx_pkg` and walks every fixture
    // through PartStore.open + partNames + imageAnchors +
    // chartAnchors, asserting valid fixtures open clean and
    // adversarial fixtures error typed instead of crashing.
    const package_corpus_mod = b.createModule(.{
        .root_source_file = b.path("tests/package_corpus.zig"),
        .target = target,
        .optimize = optimize,
    });
    package_corpus_mod.addImport("zlsx_pkg", package_mod);
    const package_corpus_tests = b.addTest(.{ .root_module = package_corpus_mod });
    corpus_step.dependOn(&b.addRunArtifact(package_corpus_tests).step);
    test_step.dependOn(&b.addRunArtifact(package_corpus_tests).step);

    // tests/typed_parts_corpus.zig — corpus parity sweep for B1
    // iter-wb-1: every fixture's workbook.xml / sst / styles / theme
    // / first sheet runs through the matching pkg.typed_parts parser.
    const typed_parts_corpus_mod = b.createModule(.{
        .root_source_file = b.path("tests/typed_parts_corpus.zig"),
        .target = target,
        .optimize = optimize,
    });
    typed_parts_corpus_mod.addImport("zlsx_pkg", package_mod);
    const typed_parts_corpus_tests = b.addTest(.{ .root_module = typed_parts_corpus_mod });
    corpus_step.dependOn(&b.addRunArtifact(typed_parts_corpus_tests).step);
    test_step.dependOn(&b.addRunArtifact(typed_parts_corpus_tests).step);

    // tests/workbook_corpus.zig — corpus parity sweep for B1 iter-wb-2:
    // every fixture opens as Workbook, every Worksheet materialises
    // through the typed-overlay composition (PartStore → typed_parts
    // → Workbook → Worksheet).
    const workbook_corpus_mod = b.createModule(.{
        .root_source_file = b.path("tests/workbook_corpus.zig"),
        .target = target,
        .optimize = optimize,
    });
    workbook_corpus_mod.addImport("zlsx_pkg", package_mod);
    const workbook_corpus_tests = b.addTest(.{ .root_module = workbook_corpus_mod });
    corpus_step.dependOn(&b.addRunArtifact(workbook_corpus_tests).step);
    test_step.dependOn(&b.addRunArtifact(workbook_corpus_tests).step);

    // C ABI — both a shared library (for Python / cffi bindings) and a
    // static library (for language toolchains that prefer linking in).
    const c_abi_mod = b.createModule(.{
        .root_source_file = b.path("src/c_abi.zig"),
        .target = target,
        .optimize = optimize,
        .single_threaded = single_threaded,
    });
    c_abi_mod.addImport("build_options", build_options_mod);
    // After the B2 iter-er-0 Editor relocation, c_abi reaches Editor
    // through zlsx_pkg and xlsx via the named `zlsx` dep (no more
    // relative `@import("xlsx.zig")` / `@import("writer.zig")`, so
    // the iter-wb-3 module-graph collision no longer fires).
    c_abi_mod.addImport("zlsx", zlsx_mod);
    c_abi_mod.addImport("zlsx_pkg", package_mod);
    const dylib = b.addLibrary(.{
        .name = "zlsx",
        .linkage = .dynamic,
        .root_module = c_abi_mod,
    });
    // Reserve Mach-O load-command headerpad on macOS targets so
    // Homebrew's install_name_tool can rewrite the dylib's install_name
    // to an absolute cellar path after unpacking. Without this, brew's
    // `fix_install_linkage` fails (the default `@rpath/libzlsx.dylib`
    // install_name has no room to expand). No-op on non-Mach-O targets.
    dylib.headerpad_max_install_names = true;
    b.installArtifact(dylib);

    // On Unix-likes, static and dynamic artifacts have different
    // extensions (.a vs .dylib/.so) so both can share the basename `zlsx`
    // and conventional `-lzlsx` linking keeps working. On Windows both
    // resolve to .lib, so the static archive gets the distinct
    // `zlsx_static` name there to avoid clobbering the DLL's import lib.
    const static_name: []const u8 = if (target.result.os.tag == .windows) "zlsx_static" else "zlsx";
    const staticlib = b.addLibrary(.{
        .name = static_name,
        .linkage = .static,
        .root_module = c_abi_mod,
    });
    b.installArtifact(staticlib);

    // Unit tests for the ABI layer (version constant, CCell translation,
    // and a corpus-gated end-to-end lifecycle smoke test).
    const c_abi_tests = b.addTest(.{ .root_module = c_abi_mod });
    test_step.dependOn(&b.addRunArtifact(c_abi_tests).step);

    // ─── B1 iter-wb-6: RSS gate ─────────────────────────────────
    //
    // Off the default `test` step. The 100k × 10 fixture takes
    // 1-3 minutes to synthesise on first invocation (cached
    // afterwards), and RSS measurement is order-sensitive enough
    // that we don't want it co-resident with other tests in the
    // same binary. Run with `zig build bench-workbook-rss`.
    //
    // The orchestrator test (`tests/bench/workbook_rss.zig`) and
    // its three child probes are split into separate compilations
    // because `zlsx` and `zlsx_pkg` cannot coexist in one binary
    // (the same `pkg/store.zig` ends up claimed by both — see
    // `AGENTS.md` "Three-module collision"). The test spawns each
    // probe as a subprocess; each probe measures its own RSS delta
    // in isolation. The gate compares the two deltas as a ratio.
    //
    // ReleaseSafe: keeps reader / writer overflow checks active so
    // we measure code paths a production caller actually runs.
    const bench_optimize: std.builtin.OptimizeMode = .ReleaseSafe;

    // Shared rss + synth modules — created once, imported by each
    // probe that needs them.
    const bench_rss_mod = b.createModule(.{
        .root_source_file = b.path("tests/bench/rss.zig"),
        .target = target,
        .optimize = bench_optimize,
    });
    const bench_synth_mod = b.createModule(.{
        .root_source_file = b.path("tests/bench/synth_100k_x_10.zig"),
        .target = target,
        .optimize = bench_optimize,
    });
    bench_synth_mod.addImport("zlsx", zlsx_mod);

    // Probe 1 — synth. Pulls `zlsx` (writer) only; no pkg.
    const probe_synth_mod = b.createModule(.{
        .root_source_file = b.path("tests/bench/rss_probe_synth.zig"),
        .target = target,
        .optimize = bench_optimize,
    });
    probe_synth_mod.addImport("synth", bench_synth_mod);
    const probe_synth_exe = b.addExecutable(.{
        .name = "zlsx-bench-rss-synth",
        .root_module = probe_synth_mod,
    });
    const probe_synth_install = b.addInstallArtifact(probe_synth_exe, .{});

    // Probe 2 — book. Pulls `zlsx` only; no pkg.
    const probe_book_mod = b.createModule(.{
        .root_source_file = b.path("tests/bench/rss_probe_book.zig"),
        .target = target,
        .optimize = bench_optimize,
    });
    probe_book_mod.addImport("zlsx", zlsx_mod);
    probe_book_mod.addImport("rss", bench_rss_mod);
    const probe_book_exe = b.addExecutable(.{
        .name = "zlsx-bench-rss-book",
        .root_module = probe_book_mod,
    });
    const probe_book_install = b.addInstallArtifact(probe_book_exe, .{});

    // Probe 3 — workbook. Pulls `zlsx_pkg` only; no zlsx (the
    // collision avoidance is the whole reason this is a separate
    // binary).
    const probe_wb_mod = b.createModule(.{
        .root_source_file = b.path("tests/bench/rss_probe_workbook.zig"),
        .target = target,
        .optimize = bench_optimize,
    });
    probe_wb_mod.addImport("zlsx_pkg", package_mod);
    probe_wb_mod.addImport("rss", bench_rss_mod);
    const probe_wb_exe = b.addExecutable(.{
        .name = "zlsx-bench-rss-workbook",
        .root_module = probe_wb_mod,
    });
    const probe_wb_install = b.addInstallArtifact(probe_wb_exe, .{});

    // Orchestrator test. No dependency on `zlsx` or `zlsx_pkg` —
    // it spawns the probes as subprocesses.
    const bench_workbook_rss_mod = b.createModule(.{
        .root_source_file = b.path("tests/bench/workbook_rss.zig"),
        .target = target,
        .optimize = bench_optimize,
    });
    bench_workbook_rss_mod.addImport("rss", bench_rss_mod);
    const bench_workbook_rss_tests = b.addTest(.{ .root_module = bench_workbook_rss_mod });
    const bench_workbook_rss_run = b.addRunArtifact(bench_workbook_rss_tests);
    bench_workbook_rss_run.has_side_effects = true;
    bench_workbook_rss_run.step.dependOn(&probe_synth_install.step);
    bench_workbook_rss_run.step.dependOn(&probe_book_install.step);
    bench_workbook_rss_run.step.dependOn(&probe_wb_install.step);

    const bench_workbook_rss_step = b.step(
        "bench-workbook-rss",
        "B1 iter-wb-6 RSS gate: Workbook.openLazy ≤ 2× Book.openLazy on a 100k × 10 fixture (off the default test path)",
    );
    bench_workbook_rss_step.dependOn(&bench_workbook_rss_run.step);

    // Per-module unit tests for the bench helpers (rss + synth).
    // These DO go on the default `test` step — they're cheap and
    // exercise the platform-specific code paths.
    const bench_rss_tests = b.addTest(.{ .root_module = bench_rss_mod });
    test_step.dependOn(&b.addRunArtifact(bench_rss_tests).step);

    const bench_synth_tests = b.addTest(.{ .root_module = bench_synth_mod });
    test_step.dependOn(&b.addRunArtifact(bench_synth_tests).step);
}
