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

    // M5d1: the cancellation / deadline / 64 KiB-chunk seam (§5.5,
    // §5.10). Std-only and imported by name everywhere, which is what
    // lets it sit beneath BOTH public trees at once: `pkg/zip.zig` polls
    // it while staying stdlib-only, and `src/formula/run_inputs.zig`
    // re-exports its `CancelToken` so the evaluator and the archive
    // layer name one type rather than two structurally identical ones.
    //
    // Wired into every module that reaches either consumer — the formula
    // files that touch `run_inputs.zig` (eval, graph, iterate, registry,
    // rng, serial_date, symbols, text) and the whole `pkg/` + writer
    // side. An unused module dep costs nothing; a missing one is a
    // compile error naming the module, so the list below is checked by
    // the build itself.
    const control_mod = b.createModule(.{
        .root_source_file = b.path("pkg/control.zig"),
        .target = target,
        .optimize = optimize,
    });

    // B3 iter-wr-1: shared `SstExtensionPlan` substrate. Std-only by
    // design, so wiring it into the writer + xlsx test trees does not
    // form the cycle `writer → zlsx_pkg → workbook → zlsx → writer`
    // that the comment near `extract_images_mod` warns about.
    const sst_plan_mod = b.createModule(.{
        .root_source_file = b.path("pkg/sst_plan.zig"),
        .target = target,
        .optimize = optimize,
    });
    sst_plan_mod.addImport("zlsx_control", control_mod);

    // B3 iter-wr-2: shared `StylesPlan` substrate. Same architectural
    // role as `sst_plan_mod` for the styles axis — std-only,
    // cycle-free, hosts the `Style` / `Dxf` / `BorderSide` /
    // `BorderStyle` / `PatternType` / `HAlign` types plus the dedup +
    // fresh-emit registry. `xlsx.Writer` re-exports the type names so
    // the public writer API surface is unchanged.
    const styles_plan_mod = b.createModule(.{
        .root_source_file = b.path("pkg/styles_plan.zig"),
        .target = target,
        .optimize = optimize,
    });
    styles_plan_mod.addImport("zlsx_control", control_mod);

    // B3 iter-wr-3: shared workbook.xml fresh-emit plan substrate
    // (defined-name validator + storage + emitter). Same std-only,
    // cycle-free shape as `sst_plan_mod` — both Workbook and
    // `xlsx.Writer` import this without forming `writer → zlsx_pkg
    // → workbook → zlsx → writer`.
    // NOTE: `zlsx_refs` is wired into this module further down, right
    // after `refs_mod` is created (it cannot be referenced yet here).
    const workbook_xml_plan_mod = b.createModule(.{
        .root_source_file = b.path("pkg/workbook_xml_plan.zig"),
        .target = target,
        .optimize = optimize,
    });
    workbook_xml_plan_mod.addImport("zlsx_control", control_mod);

    // B3 iter-wr-5: shared ZIP archive emit substrate. Std-only by
    // the same rationale as `sst_plan_mod` — pkg/zip.zig hosts the
    // LFH+CDFH+EOCD layout consumed by both `xlsx.Writer.save`
    // (fresh-file producer) and (future) `pkg.PartStore.save`. The
    // module is purely structural; it takes a `DeflateFn` callback
    // from the caller so it never needs to depend on the deflate
    // implementation living in `src/writer.zig`.
    const zip_mod = b.createModule(.{
        .root_source_file = b.path("pkg/zip.zig"),
        .target = target,
        .optimize = optimize,
    });
    zip_mod.addImport("zlsx_control", control_mod);

    // B3 iter-wr-4: shared per-sheet fresh-emit plan substrate. Same
    // std-only, cycle-free shape as `sst_plan_mod` — hosts the
    // CT_Worksheet child-order emitter (`emitWorksheetXml`), per-sheet
    // rels emitter, comments + VML drawing emitters, and the row /
    // ref / range / xml-escape primitives shared by both fresh-emit
    // producers (Writer today; Workbook fresh-emit in future iters).
    const sheet_plan_mod = b.createModule(.{
        .root_source_file = b.path("pkg/sheet_plan.zig"),
        .target = target,
        .optimize = optimize,
    });
    sheet_plan_mod.addImport("zlsx_control", control_mod);

    // B3 iter-wr-7: fresh-archive emit substrate. Lifts the entire
    // archive orchestration (Content_Types.xml + rels + workbook.xml +
    // per-sheet sheet/rels/comments/vml + sst + styles + ZIP CD/EOCD)
    // into a std-only module both `xlsx.Writer.save` and
    // `pkg.Workbook.saveFreshEmit` consume. Same cycle-avoidance pattern
    // as the other plan modules (takes a `DeflateFn` callback so the
    // caller's deflate impl can live downstream).
    const fresh_emit_mod = b.createModule(.{
        .root_source_file = b.path("pkg/fresh_emit.zig"),
        .target = target,
        .optimize = optimize,
    });
    fresh_emit_mod.addImport("zlsx_control", control_mod);
    fresh_emit_mod.addImport("zlsx_sst_plan", sst_plan_mod);
    fresh_emit_mod.addImport("zlsx_styles_plan", styles_plan_mod);
    fresh_emit_mod.addImport("zlsx_workbook_xml_plan", workbook_xml_plan_mod);
    fresh_emit_mod.addImport("zlsx_sheet_plan", sheet_plan_mod);
    fresh_emit_mod.addImport("zlsx_zip", zip_mod);

    // PRNG fuzz-harness knobs.
    //
    // These used to be read at runtime via `std.process.getEnvVarOwned`.
    // Zig 0.16 removed that, and made the process environment
    // non-ambient: a test binary built with `Io.Threaded.init(gpa, .{})`
    // sees an EMPTY environ, so a test can no longer read
    // XLSX_FUZZ_ITERS itself (verified experimentally, not assumed).
    //
    // The build runner still has an environment, so read the same two
    // variables here and pass them down as build options. That keeps
    // the documented invocation working unchanged:
    //
    //     XLSX_FUZZ_ITERS=50_000 zig build test
    //     XLSX_FUZZ_SEED=12345   zig build test
    //
    // An explicit -Dfuzz-iters / -Dfuzz-seed wins over the environment
    // so a CI job can pin values regardless of ambient state.
    const fuzz_opts = b.addOptions();
    fuzz_opts.addOption(?usize, "iters_override", blk: {
        if (b.option(usize, "fuzz-iters", "PRNG fuzz iterations per target")) |v| break :blk v;
        break :blk parseUnderscoredInt(usize, b.graph.environ_map.get("XLSX_FUZZ_ITERS"));
    });
    fuzz_opts.addOption(?u64, "seed_override", blk: {
        if (b.option(u64, "fuzz-seed", "PRNG fuzz seed")) |v| break :blk v;
        break :blk parseUnderscoredInt(u64, b.graph.environ_map.get("XLSX_FUZZ_SEED"));
    });
    const fuzz_config_mod = fuzz_opts.createModule();

    // Unicode NFC normalizer, shared by writer/editor sheet-name
    // validation (via unicode/casefold.zig) and the embedding-part
    // canonical hash pipeline (pkg/embedding_part.zig).
    //
    // Rooted at top-level `unicode/` rather than under `src/`
    // deliberately: a file may belong to exactly one module's package
    // tree, and these two consumers sit in different trees (`src/` is
    // the `zlsx` module's package dir, `pkg/` is `zlsx_pkg`'s). Anywhere
    // under either of those and the compile fails with "file exists in
    // modules 'zlsx' and 'zlsx_nfc'".
    //
    // Re-verified experimentally against Zig 0.16.0 (2026-07-26) — the
    // constraint is unchanged from 0.15.2, so the relocation stands.
    const nfc_mod = b.createModule(.{
        .root_source_file = b.path("unicode/nfc.zig"),
        .target = target,
        .optimize = optimize,
    });
    nfc_mod.addImport("zlsx_control", control_mod);

    // M1a (tier D1): UAX #31 `XID_Start` / `XID_Continue` tables — the
    // Unicode half of the formula identifier grammar. Rooted next to
    // `nfc.zig` for the same one-file-one-module-tree reason; the
    // tokenizer lives in `src/`, so the tables cannot.
    const xid_mod = b.createModule(.{
        .root_source_file = b.path("unicode/xid.zig"),
        .target = target,
        .optimize = optimize,
    });
    xid_mod.addImport("zlsx_control", control_mod);

    // A1: full case folding (Excel sheet-name dedup, `collation_v1`).
    // M4f moved it out of `src/unicode/` and up here beside the other
    // three, which is where the one-file-one-module-tree rule always
    // said it belonged: inside `src/` it was part of the `zlsx` package
    // tree, so a compilation could not hold both `zlsx` and a named
    // module rooted on it, and `src/formula/value.zig` had to keep its
    // fold import confined to a test block to stay buildable. Now
    // `zlsx` imports it by name like everyone else.
    const unicode_mod = b.createModule(.{
        .root_source_file = b.path("unicode/casefold.zig"),
        .target = target,
        .optimize = optimize,
    });
    unicode_mod.addImport("zlsx_control", control_mod);
    unicode_mod.addImport("zlsx_nfc", nfc_mod);

    // M4f: `casing_v1` — full Unicode casing for UPPER/LOWER, and the
    // titlecase half M8b's PROPER will segment over. A separate module
    // from the fold beside it because they answer different questions:
    // `fold("ß")` is the comparison key `"ss"`, `UPPER("ß")` is the
    // displayed value `"SS"`.
    const casing_mod = b.createModule(.{
        .root_source_file = b.path("unicode/casing.zig"),
        .target = target,
        .optimize = optimize,
    });
    casing_mod.addImport("zlsx_control", control_mod);
    // For the version cross-check only: casing and folding must be
    // generated from one Unicode revision, and reaching the fold's
    // table file directly would put it in two module trees.
    casing_mod.addImport("zlsx_casefold", unicode_mod);

    // M0 (tier D1): typed coordinates — the single owner of A1 parse /
    // format and the grid bounds. Rooted at top-level `refs/` for the
    // same one-file-one-module-tree reason as `unicode/` above: its
    // consumers span `zlsx` (src/), `zlsx_pkg` (pkg/), and
    // `zlsx_sheet_plan` (pkg/sheet_plan.zig).
    const refs_mod = b.createModule(.{
        .root_source_file = b.path("refs/refs.zig"),
        .target = target,
        .optimize = optimize,
    });
    refs_mod.addImport("zlsx_control", control_mod);
    sheet_plan_mod.addImport("zlsx_refs", refs_mod);
    workbook_xml_plan_mod.addImport("zlsx_refs", refs_mod);

    // Public module. Consumers add zlsx to their build.zig.zon as a
    // path or git dependency, then `@import("zlsx")`.
    const zlsx_mod = b.addModule("zlsx", .{
        .root_source_file = b.path("src/xlsx.zig"),
        .target = target,
        .optimize = optimize,
    });
    zlsx_mod.addImport("zlsx_control", control_mod);
    zlsx_mod.addImport("zlsx_sst_plan", sst_plan_mod);
    zlsx_mod.addImport("zlsx_styles_plan", styles_plan_mod);
    zlsx_mod.addImport("zlsx_workbook_xml_plan", workbook_xml_plan_mod);
    zlsx_mod.addImport("zlsx_zip", zip_mod);
    zlsx_mod.addImport("zlsx_sheet_plan", sheet_plan_mod);
    zlsx_mod.addImport("zlsx_fresh_emit", fresh_emit_mod);
    zlsx_mod.addImport("fuzz_config", fuzz_config_mod);
    zlsx_mod.addImport("zlsx_nfc", nfc_mod);
    zlsx_mod.addImport("zlsx_refs", refs_mod);
    zlsx_mod.addImport("zlsx_xid", xid_mod);
    zlsx_mod.addImport("zlsx_casefold", unicode_mod);
    zlsx_mod.addImport("zlsx_casing", casing_mod);

    // Unit tests (embedded in src/xlsx.zig, including the fuzz suite).
    const unit_mod = b.createModule(.{
        .root_source_file = b.path("src/xlsx.zig"),
        .target = target,
        .optimize = optimize,
    });
    unit_mod.addImport("zlsx_control", control_mod);
    unit_mod.addImport("zlsx_sst_plan", sst_plan_mod);
    unit_mod.addImport("zlsx_styles_plan", styles_plan_mod);
    unit_mod.addImport("zlsx_workbook_xml_plan", workbook_xml_plan_mod);
    unit_mod.addImport("zlsx_zip", zip_mod);
    unit_mod.addImport("zlsx_sheet_plan", sheet_plan_mod);
    unit_mod.addImport("zlsx_fresh_emit", fresh_emit_mod);
    unit_mod.addImport("fuzz_config", fuzz_config_mod);
    unit_mod.addImport("zlsx_nfc", nfc_mod);
    unit_mod.addImport("zlsx_casefold", unicode_mod);
    unit_mod.addImport("zlsx_casing", casing_mod);
    unit_mod.addImport("zlsx_refs", refs_mod);
    unit_mod.addImport("zlsx_xid", xid_mod);
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
    unit_fuzz_mod.addImport("zlsx_control", control_mod);
    unit_fuzz_mod.addImport("zlsx_sst_plan", sst_plan_mod);
    unit_fuzz_mod.addImport("zlsx_styles_plan", styles_plan_mod);
    unit_fuzz_mod.addImport("zlsx_workbook_xml_plan", workbook_xml_plan_mod);
    unit_fuzz_mod.addImport("zlsx_zip", zip_mod);
    unit_fuzz_mod.addImport("zlsx_sheet_plan", sheet_plan_mod);
    unit_fuzz_mod.addImport("zlsx_fresh_emit", fresh_emit_mod);
    unit_fuzz_mod.addImport("fuzz_config", fuzz_config_mod);
    unit_fuzz_mod.addImport("zlsx_nfc", nfc_mod);
    unit_fuzz_mod.addImport("zlsx_casefold", unicode_mod);
    unit_fuzz_mod.addImport("zlsx_casing", casing_mod);
    unit_fuzz_mod.addImport("zlsx_refs", refs_mod);
    unit_fuzz_mod.addImport("zlsx_xid", xid_mod);
    // Zig 0.16.0 cannot compile its own test runner in `-ffuzz` mode
    // (`writeStackTrace` receives a `*builtin.StackTrace` where a
    // `*const debug.StackTrace` is required), which blocked coverage-
    // guided fuzzing entirely. `vendor/zig-test-runner/` is upstream's
    // runner with that one hunk fixed; see its README for the removal
    // condition once a Zig release carries the fix.
    //
    // Scoped to the two fuzz binaries only — every other test target
    // below uses the stock runner from the toolchain.
    const fuzz_test_runner: std.Build.Step.Compile.TestRunner = .{
        .path = b.path("vendor/zig-test-runner/test_runner.zig"),
        .mode = .server,
    };

    const unit_fuzz_tests = b.addTest(.{
        .root_module = unit_fuzz_mod,
        .test_runner = fuzz_test_runner,
    });
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
    corpus_mod.addImport("zlsx_control", control_mod);
    corpus_mod.addImport("zlsx", zlsx_mod);
    // corpus_mod also needs `zlsx_pkg` for `Editor` post B2 iter-er-0;
    // wired below after `package_mod` is declared.
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
    cli_mod.addImport("zlsx_control", control_mod);
    cli_mod.addImport("build_options", build_options_mod);
    cli_mod.addImport("fuzz_config", fuzz_config_mod);
    cli_mod.addImport("zlsx_refs", refs_mod);
    // `cli_mod.addImport("zlsx", zlsx_mod);` is wired below, after
    // `package_mod` is declared (cli also gains a `zlsx_pkg` dep
    // post B2 iter-er-0).
    const cli_exe = b.addExecutable(.{ .name = "zlsx", .root_module = cli_mod });
    b.installArtifact(cli_exe);

    const run_cli = b.addRunArtifact(cli_exe);
    if (b.args) |args| run_cli.addArgs(args);
    const run_step = b.step("run", "Build and run the zlsx CLI (args after --)");
    run_step.dependOn(&run_cli.step);

    // CLI unit tests (colLetter, JSON/CSV escapers, arg parser).
    const cli_tests = b.addTest(.{ .root_module = cli_mod });
    test_step.dependOn(&b.addRunArtifact(cli_tests).step);

    // Writer unit tests (MVP: round-trip via the reader). Pulls in the
    // SST plan module (B3 iter-wr-1) so writer.zig can stage strings
    // through `SstExtensionPlan`.
    const writer_mod = b.createModule(.{
        .root_source_file = b.path("src/writer.zig"),
        .target = target,
        .optimize = optimize,
    });
    writer_mod.addImport("zlsx_control", control_mod);
    writer_mod.addImport("zlsx_sst_plan", sst_plan_mod);
    writer_mod.addImport("zlsx_styles_plan", styles_plan_mod);
    writer_mod.addImport("zlsx_workbook_xml_plan", workbook_xml_plan_mod);
    writer_mod.addImport("zlsx_zip", zip_mod);
    writer_mod.addImport("zlsx_sheet_plan", sheet_plan_mod);
    writer_mod.addImport("zlsx_fresh_emit", fresh_emit_mod);
    writer_mod.addImport("fuzz_config", fuzz_config_mod);
    writer_mod.addImport("zlsx_nfc", nfc_mod);
    writer_mod.addImport("zlsx_casefold", unicode_mod);
    writer_mod.addImport("zlsx_casing", casing_mod);
    writer_mod.addImport("zlsx_refs", refs_mod);
    const writer_tests = b.addTest(.{ .root_module = writer_mod });
    test_step.dependOn(&b.addRunArtifact(writer_tests).step);

    // Standalone tests for the SST plan substrate (it's tiny but the
    // dedup invariants pin Writer's hot path; cover them in their own
    // step rather than rely on the in-tree Workbook tests touching
    // them).
    const sst_plan_tests_mod = b.createModule(.{
        .root_source_file = b.path("pkg/sst_plan.zig"),
        .target = target,
        .optimize = optimize,
    });
    sst_plan_tests_mod.addImport("zlsx_control", control_mod);
    const sst_plan_tests = b.addTest(.{ .root_module = sst_plan_tests_mod });
    test_step.dependOn(&b.addRunArtifact(sst_plan_tests).step);

    // Standalone tests for the styles plan substrate (B3 iter-wr-2).
    // Same per-module test pattern as `sst_plan_tests_mod` — the plan
    // is the byte-fragile axis, so its emit invariants get their own
    // test binary.
    const styles_plan_tests_mod = b.createModule(.{
        .root_source_file = b.path("pkg/styles_plan.zig"),
        .target = target,
        .optimize = optimize,
    });
    styles_plan_tests_mod.addImport("zlsx_control", control_mod);
    const styles_plan_tests = b.addTest(.{ .root_module = styles_plan_tests_mod });
    test_step.dependOn(&b.addRunArtifact(styles_plan_tests).step);
    // B3 iter-wr-3 standalone tests for the workbook.xml fresh-emit
    // plan substrate. Same shape as `sst_plan_tests_mod`: a tiny
    // module exercising the validator + plan + emit invariants
    // without the surrounding consumer wiring.
    const workbook_xml_plan_tests_mod = b.createModule(.{
        .root_source_file = b.path("pkg/workbook_xml_plan.zig"),
        .target = target,
        .optimize = optimize,
    });
    workbook_xml_plan_tests_mod.addImport("zlsx_control", control_mod);
    workbook_xml_plan_tests_mod.addImport("zlsx_refs", refs_mod);
    const workbook_xml_plan_tests = b.addTest(.{ .root_module = workbook_xml_plan_tests_mod });
    test_step.dependOn(&b.addRunArtifact(workbook_xml_plan_tests).step);
    // Standalone tests for the ZIP emit substrate (B3 iter-wr-5).
    // Same separation rationale as `sst_plan_tests_mod`.
    const zip_tests_mod = b.createModule(.{
        .root_source_file = b.path("pkg/zip.zig"),
        .target = target,
        .optimize = optimize,
    });
    zip_tests_mod.addImport("zlsx_control", control_mod);
    const zip_tests = b.addTest(.{ .root_module = zip_tests_mod });
    test_step.dependOn(&b.addRunArtifact(zip_tests).step);

    // B3 iter-wr-4 standalone tests for the per-sheet fresh-emit plan
    // substrate. Same separation rationale as `sst_plan_tests_mod` —
    // the plan owns the CT_Worksheet child-order emit + VML drawing,
    // both byte-fragile invariants worth their own test binary.
    const sheet_plan_tests_mod = b.createModule(.{
        .root_source_file = b.path("pkg/sheet_plan.zig"),
        .target = target,
        .optimize = optimize,
    });
    sheet_plan_tests_mod.addImport("zlsx_control", control_mod);
    sheet_plan_tests_mod.addImport("zlsx_refs", refs_mod);
    const sheet_plan_tests = b.addTest(.{ .root_module = sheet_plan_tests_mod });
    test_step.dependOn(&b.addRunArtifact(sheet_plan_tests).step);

    // M0: standalone tests for the typed-coordinate module. Its own
    // binary because every other suite reaches it only through an
    // adapter, so a bug in the primitive would otherwise surface as a
    // confusing failure three layers up.
    const refs_tests = b.addTest(.{ .root_module = refs_mod });
    test_step.dependOn(&b.addRunArtifact(refs_tests).step);

    // M0 import gate. Scans every `src/` and `pkg/` source for
    // bijective base-26 coordinate arithmetic — the fingerprint shared
    // by all six pre-M0 A1 parsers — and fails the build on a new one.
    // Wired into `test_step` so CI enforces it without a bespoke job.
    const import_gate_mod = b.createModule(.{
        .root_source_file = b.path("refs/import_gate.zig"),
        .target = b.graph.host,
        .optimize = .Debug,
    });
    import_gate_mod.addImport("zlsx_control", control_mod);
    const import_gate_exe = b.addExecutable(.{
        .name = "import-gate",
        .root_module = import_gate_mod,
    });
    const import_gate_run = b.addRunArtifact(import_gate_exe);
    // 0.16: filesystem calls take the build graph's `Io`.
    const gate_io = b.graph.io;
    for ([_][]const u8{ "src", "pkg" }) |root| {
        var dir = b.build_root.handle.openDir(gate_io, root, .{ .iterate = true }) catch
            @panic("import gate: cannot open source dir");
        defer dir.close(gate_io);
        var walker = dir.walk(b.allocator) catch @panic("import gate: walk failed");
        defer walker.deinit();
        while (walker.next(gate_io) catch @panic("import gate: walk failed")) |entry| {
            if (entry.kind != .file) continue;
            if (!std.mem.endsWith(u8, entry.basename, ".zig")) continue;
            import_gate_run.addFileArg(b.path(b.pathJoin(&.{ root, entry.path })));
        }
    }
    test_step.dependOn(&import_gate_run.step);

    const import_gate_tests = b.addTest(.{ .root_module = import_gate_mod });
    test_step.dependOn(&b.addRunArtifact(import_gate_tests).step);

    // B3 iter-wr-7: standalone tests for the fresh-archive emit
    // substrate. Same separation rationale as `sst_plan_tests_mod` —
    // the module owns the [Content_Types].xml + workbook.xml + per-sheet
    // archive orchestration, byte-fragile invariants worth their own
    // test binary.
    const fresh_emit_tests_mod = b.createModule(.{
        .root_source_file = b.path("pkg/fresh_emit.zig"),
        .target = target,
        .optimize = optimize,
    });
    fresh_emit_tests_mod.addImport("zlsx_control", control_mod);
    fresh_emit_tests_mod.addImport("zlsx_sst_plan", sst_plan_mod);
    fresh_emit_tests_mod.addImport("zlsx_styles_plan", styles_plan_mod);
    fresh_emit_tests_mod.addImport("zlsx_workbook_xml_plan", workbook_xml_plan_mod);
    fresh_emit_tests_mod.addImport("zlsx_sheet_plan", sheet_plan_mod);
    fresh_emit_tests_mod.addImport("zlsx_zip", zip_mod);
    const fresh_emit_tests = b.addTest(.{ .root_module = fresh_emit_tests_mod });
    test_step.dependOn(&b.addRunArtifact(fresh_emit_tests).step);

    // Unicode case-fold + NFC module tests (A1: Excel sheet-name
    // dedup, wired into validateSheetName + Editor.isSheetNameTaken).
    // The module itself is created up with `nfc_mod`, because `zlsx`
    // now imports it by name (M4f).
    const unicode_tests = b.addTest(.{ .root_module = unicode_mod });
    test_step.dependOn(&b.addRunArtifact(unicode_tests).step);

    // M4f: `casing_v1` tests — the length-changing mappings, the dotted
    // and dotless I, Final_Sigma in all three positions, and the table
    // invariants.
    const casing_tests = b.addTest(.{ .root_module = casing_mod });
    test_step.dependOn(&b.addRunArtifact(casing_tests).step);

    // A1 phase 3: standalone NFC tests (the casefold module imports
    // nfc.zig but has its own tests; this catches NFC bugs that
    // wouldn't surface through the casefold compositional path).
    // Reuses the shared `nfc_mod` created above — it can no longer be
    // a second module rooted on the same file.
    const nfc_tests = b.addTest(.{ .root_module = nfc_mod });
    test_step.dependOn(&b.addRunArtifact(nfc_tests).step);

    // M1a: XID interval-table tests (boundary cases + a full-space
    // sweep of the binary search against a linear scan).
    const xid_tests = b.addTest(.{ .root_module = xid_mod });
    test_step.dependOn(&b.addRunArtifact(xid_tests).step);

    // M1b: the oracle harness. Deliberately imports NOTHING from zlsx —
    // it carries its own ZIP and XML decoders (`tests/oracle/zip_reader.zig`,
    // `xml_scan.zig`) because an oracle that shared a decoder with the
    // implementation under test could not detect a bug in that decoder.
    // The replay half runs in `zig build test` with no spreadsheet
    // application installed; recording is a separate, explicit command.
    const oracle_mod = b.createModule(.{
        .root_source_file = b.path("tests/oracle/replay.zig"),
        .target = target,
        .optimize = optimize,
    });
    oracle_mod.addImport("zlsx_control", control_mod);
    // Read by `sentinel_set.zig`, which checks its planted values
    // against the builder script rather than trusting them to stay in
    // sync by convention.
    oracle_mod.addAnonymousImport("build_inputs_py", .{
        .root_source_file = b.path("scripts/oracle/build_inputs.py"),
    });
    const oracle_tests = b.addTest(.{ .root_module = oracle_mod });
    test_step.dependOn(&b.addRunArtifact(oracle_tests).step);

    // Also reachable on its own, so CI can show the oracle gate as a
    // named step and a developer can run it without the full suite.
    const oracle_step = b.step("test-oracle", "Replay the committed oracle manifests (no apps needed)");
    oracle_step.dependOn(&b.addRunArtifact(oracle_tests).step);

    // The recorder. Not wired into `test` — it drives nothing by
    // itself, but `scripts/oracle/regenerate.sh` needs the binary.
    const oracle_record_mod = b.createModule(.{
        .root_source_file = b.path("tests/oracle/record.zig"),
        .target = target,
        .optimize = optimize,
    });
    oracle_record_mod.addImport("zlsx_control", control_mod);
    const oracle_record_exe = b.addExecutable(.{
        .name = "zlsx-oracle-record",
        .root_module = oracle_record_mod,
    });
    b.installArtifact(oracle_record_exe);

    // C1 milestone 1: formula tokenizer + loss-preserving printer.
    // M1a widened it to Unicode identifiers (`zlsx_xid`), the
    // dynamic-array / structured-ref token kinds, and extensible error
    // literals; the parser that consumes those lands at M2.
    const formula_tokenizer_mod = b.createModule(.{
        .root_source_file = b.path("src/formula/tokenizer.zig"),
        .target = target,
        .optimize = optimize,
    });
    formula_tokenizer_mod.addImport("zlsx_control", control_mod);
    formula_tokenizer_mod.addImport("zlsx_refs", refs_mod);
    formula_tokenizer_mod.addImport("zlsx_xid", xid_mod);
    const formula_tokenizer_tests = b.addTest(.{ .root_module = formula_tokenizer_mod });
    test_step.dependOn(&b.addRunArtifact(formula_tokenizer_tests).step);

    // M1a: coverage-guided tokenizer fuzz target (round-trip identity,
    // no-panic, refusal/kind consistency). Same `.fuzz = true` +
    // vendored test-runner pattern as `unit_fuzz_mod`.
    const formula_tokenizer_fuzz_mod = b.createModule(.{
        .root_source_file = b.path("src/formula/tokenizer.zig"),
        .target = target,
        .optimize = optimize,
        .fuzz = true,
    });
    formula_tokenizer_fuzz_mod.addImport("zlsx_control", control_mod);
    formula_tokenizer_fuzz_mod.addImport("zlsx_refs", refs_mod);
    formula_tokenizer_fuzz_mod.addImport("zlsx_xid", xid_mod);
    const formula_tokenizer_fuzz_tests = b.addTest(.{
        .root_module = formula_tokenizer_fuzz_mod,
        .test_runner = fuzz_test_runner,
    });
    fuzz_step.dependOn(&b.addRunArtifact(formula_tokenizer_fuzz_tests).step);

    // M2: the formula parser — AST, canonical printer, typed refusals.
    // Sits in the same package dir as the tokenizer it consumes, so the
    // relative `@import("tokenizer.zig")` resolves and `@embedFile` can
    // reach the M1a fixture tables for the round-trip corpus gate.
    //
    // The oracle manifests come in as anonymous imports rather than a
    // relative `@embedFile`: `tests/oracle/` is outside this module's
    // package tree, and the parser must pin its precedence against the
    // *committed* manifests, not against a copy that could drift.
    const formula_parser_mod = b.createModule(.{
        .root_source_file = b.path("src/formula/parser.zig"),
        .target = target,
        .optimize = optimize,
    });
    formula_parser_mod.addImport("zlsx_control", control_mod);
    formula_parser_mod.addImport("zlsx_refs", refs_mod);
    formula_parser_mod.addImport("zlsx_xid", xid_mod);
    formula_parser_mod.addAnonymousImport("oracle_hand_spec_excel", .{
        .root_source_file = b.path("tests/oracle/fixtures/hand_spec_excel.json"),
    });
    formula_parser_mod.addAnonymousImport("oracle_hand_spec_ieee", .{
        .root_source_file = b.path("tests/oracle/fixtures/hand_spec_ieee.json"),
    });
    formula_parser_mod.addAnonymousImport("oracle_libreoffice_suite", .{
        .root_source_file = b.path("tests/oracle/fixtures/libreoffice_oracle_suite.json"),
    });
    const formula_parser_tests = b.addTest(.{ .root_module = formula_parser_mod });
    test_step.dependOn(&b.addRunArtifact(formula_parser_tests).step);

    // M2: coverage-guided parser fuzz target (no panic, no lost bytes,
    // print/re-parse structural equality). Same `.fuzz = true` +
    // vendored test-runner pattern as the tokenizer's.
    const formula_parser_fuzz_mod = b.createModule(.{
        .root_source_file = b.path("src/formula/parser.zig"),
        .target = target,
        .optimize = optimize,
        .fuzz = true,
    });
    formula_parser_fuzz_mod.addImport("zlsx_control", control_mod);
    formula_parser_fuzz_mod.addImport("zlsx_refs", refs_mod);
    formula_parser_fuzz_mod.addImport("zlsx_xid", xid_mod);
    formula_parser_fuzz_mod.addAnonymousImport("oracle_hand_spec_excel", .{
        .root_source_file = b.path("tests/oracle/fixtures/hand_spec_excel.json"),
    });
    formula_parser_fuzz_mod.addAnonymousImport("oracle_hand_spec_ieee", .{
        .root_source_file = b.path("tests/oracle/fixtures/hand_spec_ieee.json"),
    });
    formula_parser_fuzz_mod.addAnonymousImport("oracle_libreoffice_suite", .{
        .root_source_file = b.path("tests/oracle/fixtures/libreoffice_oracle_suite.json"),
    });
    const formula_parser_fuzz_tests = b.addTest(.{
        .root_module = formula_parser_fuzz_mod,
        .test_runner = fuzz_test_runner,
    });
    fuzz_step.dependOn(&b.addRunArtifact(formula_parser_fuzz_tests).step);

    // M3a1: the formula value model — ScalarValue/Matrix, the two
    // fidelity rule tables, `parseDecimal`, the §5.3b shape and
    // coercion tables, and `collation_v1`.
    //
    // `zlsx_casefold` is imported by the TEST section of value.zig and
    // nowhere else — a file-scope `const` referenced only from a `test`
    // block is not resolved in a non-test build (verified on 0.16.0),
    // so a consumer of value.zig can build it without declaring this
    // import. `collation_v1` takes the fold as a parameter precisely so
    // that stays true.
    //
    // Until M4f the confinement was also load-bearing: the fold lived
    // at `src/unicode/casefold.zig`, inside the `zlsx` package tree, so
    // a compilation holding both `zlsx` and a module rooted on that file
    // failed "file exists in modules 'zlsx' and 'zlsx_casefold'" — the
    // collision M0 hit with `refs/`. Moving the file up here dissolved
    // that half; the parameter stays because it is the better design.
    const formula_value_mod = b.createModule(.{
        .root_source_file = b.path("src/formula/value.zig"),
        .target = target,
        .optimize = optimize,
    });
    formula_value_mod.addImport("zlsx_control", control_mod);
    formula_value_mod.addImport("zlsx_casefold", unicode_mod);
    formula_value_mod.addImport("zlsx_casing", casing_mod);
    formula_value_mod.addAnonymousImport("oracle_hand_spec_excel", .{
        .root_source_file = b.path("tests/oracle/fixtures/hand_spec_excel.json"),
    });
    formula_value_mod.addAnonymousImport("oracle_hand_spec_ieee", .{
        .root_source_file = b.path("tests/oracle/fixtures/hand_spec_ieee.json"),
    });
    formula_value_mod.addAnonymousImport("oracle_libreoffice_suite", .{
        .root_source_file = b.path("tests/oracle/fixtures/libreoffice_oracle_suite.json"),
    });
    const formula_value_tests = b.addTest(.{ .root_module = formula_value_mod });
    test_step.dependOn(&b.addRunArtifact(formula_value_tests).step);

    // M3a1: coverage-guided value fuzz targets (decimal ingress and
    // arithmetic never panic and never go non-finite).
    const formula_value_fuzz_mod = b.createModule(.{
        .root_source_file = b.path("src/formula/value.zig"),
        .target = target,
        .optimize = optimize,
        .fuzz = true,
    });
    formula_value_fuzz_mod.addImport("zlsx_control", control_mod);
    formula_value_fuzz_mod.addImport("zlsx_casefold", unicode_mod);
    formula_value_fuzz_mod.addImport("zlsx_casing", casing_mod);
    formula_value_fuzz_mod.addAnonymousImport("oracle_hand_spec_excel", .{
        .root_source_file = b.path("tests/oracle/fixtures/hand_spec_excel.json"),
    });
    formula_value_fuzz_mod.addAnonymousImport("oracle_hand_spec_ieee", .{
        .root_source_file = b.path("tests/oracle/fixtures/hand_spec_ieee.json"),
    });
    formula_value_fuzz_mod.addAnonymousImport("oracle_libreoffice_suite", .{
        .root_source_file = b.path("tests/oracle/fixtures/libreoffice_oracle_suite.json"),
    });
    const formula_value_fuzz_tests = b.addTest(.{
        .root_module = formula_value_fuzz_mod,
        .test_runner = fuzz_test_runner,
    });
    fuzz_step.dependOn(&b.addRunArtifact(formula_value_fuzz_tests).step);

    // M3a2: the EvalEnv interface and its in-memory fake. Its own test
    // target as well as being a dependency of the evaluator's, because
    // env.zig must stay buildable WITHOUT eval.zig — the direction of
    // that dependency is what keeps `src/formula/` free of `pkg/`.
    //
    // The `zlsx_casefold` and oracle imports are here because a test
    // build analyses the tests of every file it imports, and env.zig
    // imports value.zig. They cost nothing in a non-test build.
    const formula_env_mod = b.createModule(.{
        .root_source_file = b.path("src/formula/env.zig"),
        .target = target,
        .optimize = optimize,
    });
    formula_env_mod.addImport("zlsx_control", control_mod);
    formula_env_mod.addImport("zlsx_refs", refs_mod);
    formula_env_mod.addImport("zlsx_casefold", unicode_mod);
    formula_env_mod.addImport("zlsx_casing", casing_mod);
    addOracleFixtures(b, formula_env_mod);
    const formula_env_tests = b.addTest(.{ .root_module = formula_env_mod });
    test_step.dependOn(&b.addRunArtifact(formula_env_tests).step);

    // M3a2: the evaluator core plus the registry framework and the
    // frozen inventory. One module: eval.zig and registry.zig import
    // each other (the table needs the evaluator's `Value`, the
    // dispatcher needs the table), which files in one module may do and
    // separate modules may not.
    const formula_eval_mod = b.createModule(.{
        .root_source_file = b.path("src/formula/eval.zig"),
        .target = target,
        .optimize = optimize,
    });
    formula_eval_mod.addImport("zlsx_control", control_mod);
    formula_eval_mod.addImport("zlsx_refs", refs_mod);
    formula_eval_mod.addImport("zlsx_xid", xid_mod);
    formula_eval_mod.addImport("zlsx_casefold", unicode_mod);
    formula_eval_mod.addImport("zlsx_casing", casing_mod);
    addOracleFixtures(b, formula_eval_mod);
    const formula_eval_tests = b.addTest(.{ .root_module = formula_eval_mod });
    test_step.dependOn(&b.addRunArtifact(formula_eval_tests).step);

    // M3b: run inputs + the §9 byte budget, and serial dates. Both are
    // leaves — they import `value.zig` and nothing else in the engine —
    // so each gets its own module and its own test target, which is what
    // keeps them buildable without the evaluator.
    const formula_run_inputs_mod = b.createModule(.{
        .root_source_file = b.path("src/formula/run_inputs.zig"),
        .target = target,
        .optimize = optimize,
    });
    formula_run_inputs_mod.addImport("zlsx_control", control_mod);
    formula_run_inputs_mod.addImport("zlsx_refs", refs_mod);
    formula_run_inputs_mod.addImport("zlsx_casefold", unicode_mod);
    formula_run_inputs_mod.addImport("zlsx_casing", casing_mod);
    addOracleFixtures(b, formula_run_inputs_mod);
    const formula_run_inputs_tests = b.addTest(.{ .root_module = formula_run_inputs_mod });
    test_step.dependOn(&b.addRunArtifact(formula_run_inputs_tests).step);

    const formula_serial_date_mod = b.createModule(.{
        .root_source_file = b.path("src/formula/serial_date.zig"),
        .target = target,
        .optimize = optimize,
    });
    formula_serial_date_mod.addImport("zlsx_control", control_mod);
    formula_serial_date_mod.addImport("zlsx_refs", refs_mod);
    formula_serial_date_mod.addImport("zlsx_casefold", unicode_mod);
    formula_serial_date_mod.addImport("zlsx_casing", casing_mod);
    addOracleFixtures(b, formula_serial_date_mod);
    const formula_serial_date_tests = b.addTest(.{ .root_module = formula_serial_date_mod });
    test_step.dependOn(&b.addRunArtifact(formula_serial_date_tests).step);

    // M3b: criteria. Its own test target as well as being reached from
    // the evaluator's, because criteria must stay buildable without the
    // evaluator — `criteria.scan` reads through `EvalEnv`, never through
    // an `Evaluator`.
    const formula_criteria_mod = b.createModule(.{
        .root_source_file = b.path("src/formula/criteria.zig"),
        .target = target,
        .optimize = optimize,
    });
    formula_criteria_mod.addImport("zlsx_control", control_mod);
    formula_criteria_mod.addImport("zlsx_refs", refs_mod);
    formula_criteria_mod.addImport("zlsx_casefold", unicode_mod);
    formula_criteria_mod.addImport("zlsx_casing", casing_mod);
    addOracleFixtures(b, formula_criteria_mod);
    const formula_criteria_tests = b.addTest(.{ .root_module = formula_criteria_mod });
    test_step.dependOn(&b.addRunArtifact(formula_criteria_tests).step);

    // M4f: §5.4d's shared text layer — the CV1/CV2 index unit and the
    // conversions the five affected functions need. It sits below the
    // evaluator like `criteria.zig` does, and imports neither it nor the
    // registry: what a character IS cannot depend on which function
    // asked.
    const formula_text_mod = b.createModule(.{
        .root_source_file = b.path("src/formula/text.zig"),
        .target = target,
        .optimize = optimize,
    });
    formula_text_mod.addImport("zlsx_control", control_mod);
    formula_text_mod.addImport("zlsx_refs", refs_mod);
    formula_text_mod.addImport("zlsx_casefold", unicode_mod);
    formula_text_mod.addImport("zlsx_casing", casing_mod);
    addOracleFixtures(b, formula_text_mod);
    const formula_text_tests = b.addTest(.{ .root_module = formula_text_mod });
    test_step.dependOn(&b.addRunArtifact(formula_text_tests).step);

    // M5a1: the dependency graph — nodes, edges, SCC condensation, the
    // deterministic order, the §5.6c seed table. Its own module and its
    // own test target, like the text layer above: `graph.zig` imports
    // `eval.zig` for `DependencyLog` and for the static walk's shape,
    // but nothing imports `graph.zig` back, so it stays a leaf of the
    // engine and can be built and tested without the package tree.
    const formula_graph_mod = b.createModule(.{
        .root_source_file = b.path("src/formula/graph.zig"),
        .target = target,
        .optimize = optimize,
    });
    formula_graph_mod.addImport("zlsx_control", control_mod);
    formula_graph_mod.addImport("zlsx_refs", refs_mod);
    formula_graph_mod.addImport("zlsx_xid", xid_mod);
    formula_graph_mod.addImport("zlsx_casefold", unicode_mod);
    formula_graph_mod.addImport("zlsx_casing", casing_mod);
    addOracleFixtures(b, formula_graph_mod);
    const formula_graph_tests = b.addTest(.{ .root_module = formula_graph_mod });
    test_step.dependOn(&b.addRunArtifact(formula_graph_tests).step);

    // M5a2: §5.6d's volatile draw schedule. A leaf below both the
    // evaluator and the iteration engine — the evaluator knows WHERE a
    // draw happens and the engine knows WHICH PASS it happens in, so a
    // file either of them owned would make the other import a cycle.
    const formula_draws_mod = b.createModule(.{
        .root_source_file = b.path("src/formula/draws.zig"),
        .target = target,
        .optimize = optimize,
    });
    formula_draws_mod.addImport("zlsx_control", control_mod);
    const formula_draws_tests = b.addTest(.{ .root_module = formula_draws_mod });
    test_step.dependOn(&b.addRunArtifact(formula_draws_tests).step);

    // M5a2: the iteration engine — the multi-SCC schedule, convergence,
    // the two exhaustion outcomes and §5.6e's dynamic-edge fixpoint. It
    // consumes `graph.zig`'s condensation rather than rebuilding one, so
    // it sits directly above the graph and below the package tree.
    const formula_iterate_mod = b.createModule(.{
        .root_source_file = b.path("src/formula/iterate.zig"),
        .target = target,
        .optimize = optimize,
    });
    formula_iterate_mod.addImport("zlsx_control", control_mod);
    formula_iterate_mod.addImport("zlsx_refs", refs_mod);
    formula_iterate_mod.addImport("zlsx_xid", xid_mod);
    formula_iterate_mod.addImport("zlsx_casefold", unicode_mod);
    formula_iterate_mod.addImport("zlsx_casing", casing_mod);
    addOracleFixtures(b, formula_iterate_mod);
    const formula_iterate_tests = b.addTest(.{ .root_module = formula_iterate_mod });
    test_step.dependOn(&b.addRunArtifact(formula_iterate_tests).step);

    // M5b1: §5.7.3 step 3 — the `ResolvedSheet` projection and the
    // byte-confined cached-value patcher. It sits above `decode.zig`
    // (which hands out the spans it writes into) and `calc.zig` (which
    // classifies the `<f>` its spill gate refuses on), and below the
    // package tree: nothing here opens a file.
    const formula_resolved_mod = b.createModule(.{
        .root_source_file = b.path("src/formula/resolved.zig"),
        .target = target,
        .optimize = optimize,
    });
    formula_resolved_mod.addImport("zlsx_control", control_mod);
    formula_resolved_mod.addImport("zlsx_refs", refs_mod);
    formula_resolved_mod.addImport("zlsx_xid", xid_mod);
    formula_resolved_mod.addImport("zlsx_casefold", unicode_mod);
    formula_resolved_mod.addImport("zlsx_casing", casing_mod);
    addOracleFixtures(b, formula_resolved_mod);
    const formula_resolved_tests = b.addTest(.{ .root_module = formula_resolved_mod });
    test_step.dependOn(&b.addRunArtifact(formula_resolved_tests).step);

    // M5b1: the patcher's two fuzz targets (§8.1) — a patch never writes
    // outside the ranges it declares, and every value it writes reads
    // back as itself.
    const formula_resolved_fuzz_mod = b.createModule(.{
        .root_source_file = b.path("src/formula/resolved.zig"),
        .target = target,
        .optimize = optimize,
        .fuzz = true,
    });
    formula_resolved_fuzz_mod.addImport("zlsx_control", control_mod);
    formula_resolved_fuzz_mod.addImport("zlsx_refs", refs_mod);
    formula_resolved_fuzz_mod.addImport("zlsx_xid", xid_mod);
    formula_resolved_fuzz_mod.addImport("zlsx_casefold", unicode_mod);
    formula_resolved_fuzz_mod.addImport("zlsx_casing", casing_mod);
    addOracleFixtures(b, formula_resolved_fuzz_mod);
    const formula_resolved_fuzz_tests = b.addTest(.{
        .root_module = formula_resolved_fuzz_mod,
        .test_runner = fuzz_test_runner,
    });
    fuzz_step.dependOn(&b.addRunArtifact(formula_resolved_fuzz_tests).step);

    // M5b2: §5.7.6's calc-state writes and §5.7.7's mark-only
    // eligibility. Sibling of `resolved.zig` — the same byte-confined
    // edit-list shape, over `xl/workbook.xml` instead of a worksheet —
    // and it borrows M5b1's `changedWindow`, so it roots the same tree.
    const formula_calc_patch_mod = b.createModule(.{
        .root_source_file = b.path("src/formula/calc_patch.zig"),
        .target = target,
        .optimize = optimize,
    });
    formula_calc_patch_mod.addImport("zlsx_control", control_mod);
    formula_calc_patch_mod.addImport("zlsx_refs", refs_mod);
    formula_calc_patch_mod.addImport("zlsx_xid", xid_mod);
    formula_calc_patch_mod.addImport("zlsx_casefold", unicode_mod);
    formula_calc_patch_mod.addImport("zlsx_casing", casing_mod);
    addOracleFixtures(b, formula_calc_patch_mod);
    const formula_calc_patch_tests = b.addTest(.{ .root_module = formula_calc_patch_mod });
    test_step.dependOn(&b.addRunArtifact(formula_calc_patch_tests).step);

    // M3b: the criteria fuzz target (§8.1) — no criterion string may
    // panic, leak, or match non-deterministically.
    const formula_criteria_fuzz_mod = b.createModule(.{
        .root_source_file = b.path("src/formula/criteria.zig"),
        .target = target,
        .optimize = optimize,
        .fuzz = true,
    });
    formula_criteria_fuzz_mod.addImport("zlsx_control", control_mod);
    formula_criteria_fuzz_mod.addImport("zlsx_refs", refs_mod);
    formula_criteria_fuzz_mod.addImport("zlsx_casefold", unicode_mod);
    formula_criteria_fuzz_mod.addImport("zlsx_casing", casing_mod);
    addOracleFixtures(b, formula_criteria_fuzz_mod);
    const formula_criteria_fuzz_tests = b.addTest(.{
        .root_module = formula_criteria_fuzz_mod,
        .test_runner = fuzz_test_runner,
    });
    fuzz_step.dependOn(&b.addRunArtifact(formula_criteria_fuzz_tests).step);

    // M3b: `rng_v1`. Rooted with the evaluator's imports because it
    // constructs the `DrawSource` the evaluator counts.
    const formula_rng_mod = b.createModule(.{
        .root_source_file = b.path("src/formula/rng.zig"),
        .target = target,
        .optimize = optimize,
    });
    formula_rng_mod.addImport("zlsx_control", control_mod);
    formula_rng_mod.addImport("zlsx_refs", refs_mod);
    formula_rng_mod.addImport("zlsx_xid", xid_mod);
    formula_rng_mod.addImport("zlsx_casefold", unicode_mod);
    formula_rng_mod.addImport("zlsx_casing", casing_mod);
    addOracleFixtures(b, formula_rng_mod);
    const formula_rng_tests = b.addTest(.{ .root_module = formula_rng_mod });
    test_step.dependOn(&b.addRunArtifact(formula_rng_tests).step);

    // M4a: `xl/metadata.xml` — typed reader, cm/vm resolution, dialect
    // primitives. It roots with the parser's imports because its
    // refusals are `parser.PlaneTwo` values: §10's taxonomy has one home
    // and this file refuses into it rather than beside it. `env.zig`
    // comes along because M4a binds the reader to `EvalEnv.dialectOf`
    // through `env.DialectResolver` — the dependency points this way, so
    // `env.zig` stays a leaf.
    const formula_metadata_mod = b.createModule(.{
        .root_source_file = b.path("src/formula/metadata.zig"),
        .target = target,
        .optimize = optimize,
    });
    formula_metadata_mod.addImport("zlsx_control", control_mod);
    formula_metadata_mod.addImport("zlsx_refs", refs_mod);
    formula_metadata_mod.addImport("zlsx_xid", xid_mod);
    formula_metadata_mod.addImport("zlsx_casefold", unicode_mod);
    formula_metadata_mod.addImport("zlsx_casing", casing_mod);
    addOracleFixtures(b, formula_metadata_mod);
    const formula_metadata_tests = b.addTest(.{ .root_module = formula_metadata_mod });
    test_step.dependOn(&b.addRunArtifact(formula_metadata_tests).step);

    // M4a: the metadata fuzz target (§8.1) — no part can panic, leak, or
    // leave a run's dialects partially resolved.
    const formula_metadata_fuzz_mod = b.createModule(.{
        .root_source_file = b.path("src/formula/metadata.zig"),
        .target = target,
        .optimize = optimize,
        .fuzz = true,
    });
    formula_metadata_fuzz_mod.addImport("zlsx_control", control_mod);
    formula_metadata_fuzz_mod.addImport("zlsx_refs", refs_mod);
    formula_metadata_fuzz_mod.addImport("zlsx_xid", xid_mod);
    formula_metadata_fuzz_mod.addImport("zlsx_casefold", unicode_mod);
    formula_metadata_fuzz_mod.addImport("zlsx_casing", casing_mod);
    addOracleFixtures(b, formula_metadata_fuzz_mod);
    const formula_metadata_fuzz_tests = b.addTest(.{
        .root_module = formula_metadata_fuzz_mod,
        .test_runner = fuzz_test_runner,
    });
    fuzz_step.dependOn(&b.addRunArtifact(formula_metadata_fuzz_tests).step);

    // M4b1: the XML decode boundary and the decoded symbol layer.
    //
    // One module rooted at `symbols.zig`, which imports `decode.zig`
    // relatively: the symbol layer decodes through the boundary, so the
    // two are a unit, and a `pkg/` adapter that reached them as two
    // named modules would get two incompatible copies of every type
    // they share with `env.zig`. `decode.zig` gets its own test target
    // as well, because the boundary must stay buildable — and provable —
    // without the symbol layer above it.
    const formula_decode_mod = b.createModule(.{
        .root_source_file = b.path("src/formula/decode.zig"),
        .target = target,
        .optimize = optimize,
    });
    formula_decode_mod.addImport("zlsx_control", control_mod);
    formula_decode_mod.addImport("zlsx_refs", refs_mod);
    formula_decode_mod.addImport("zlsx_xid", xid_mod);
    formula_decode_mod.addImport("zlsx_casefold", unicode_mod);
    formula_decode_mod.addImport("zlsx_casing", casing_mod);
    addOracleFixtures(b, formula_decode_mod);
    const formula_decode_tests = b.addTest(.{ .root_module = formula_decode_mod });
    test_step.dependOn(&b.addRunArtifact(formula_decode_tests).step);

    // M4b2: the `CT_CellFormula` attribute inventory, the shared
    // topology, the translation matrix, and the workbook's calc state.
    // Its own module for the same reason `decode.zig` has one — the
    // attribute table and the topology must stay provable without the
    // symbol layer above them — and it reaches `decode.zig` relatively,
    // so both end up in the one engine module `pkg/` imports.
    const formula_calc_mod = b.createModule(.{
        .root_source_file = b.path("src/formula/calc.zig"),
        .target = target,
        .optimize = optimize,
    });
    formula_calc_mod.addImport("zlsx_control", control_mod);
    formula_calc_mod.addImport("zlsx_refs", refs_mod);
    formula_calc_mod.addImport("zlsx_xid", xid_mod);
    formula_calc_mod.addImport("zlsx_casefold", unicode_mod);
    formula_calc_mod.addImport("zlsx_casing", casing_mod);
    addOracleFixtures(b, formula_calc_mod);
    const formula_calc_tests = b.addTest(.{ .root_module = formula_calc_mod });
    test_step.dependOn(&b.addRunArtifact(formula_calc_tests).step);

    // M4b2: the calc fuzz target (§8.1) — no `<f>` attribute string and
    // no `<calcPr>` may panic, leak, or classify two ways.
    const formula_calc_fuzz_mod = b.createModule(.{
        .root_source_file = b.path("src/formula/calc.zig"),
        .target = target,
        .optimize = optimize,
        .fuzz = true,
    });
    formula_calc_fuzz_mod.addImport("zlsx_control", control_mod);
    formula_calc_fuzz_mod.addImport("zlsx_refs", refs_mod);
    formula_calc_fuzz_mod.addImport("zlsx_xid", xid_mod);
    formula_calc_fuzz_mod.addImport("zlsx_casefold", unicode_mod);
    formula_calc_fuzz_mod.addImport("zlsx_casing", casing_mod);
    addOracleFixtures(b, formula_calc_fuzz_mod);
    const formula_calc_fuzz_tests = b.addTest(.{
        .root_module = formula_calc_fuzz_mod,
        .test_runner = fuzz_test_runner,
    });
    fuzz_step.dependOn(&b.addRunArtifact(formula_calc_fuzz_tests).step);

    // M4b3: the `CT_DefinedName` and `CT_TableFormula` inventories,
    // §5.9's resolution drivers, and the 3D reference matrix. Its own
    // module for the reason `calc.zig` has one — the tables and the
    // matrix must stay provable without the symbol layer that consumes
    // them, and this file is the one the symbol layer imports.
    const formula_names_mod = b.createModule(.{
        .root_source_file = b.path("src/formula/names.zig"),
        .target = target,
        .optimize = optimize,
    });
    formula_names_mod.addImport("zlsx_control", control_mod);
    formula_names_mod.addImport("zlsx_refs", refs_mod);
    formula_names_mod.addImport("zlsx_xid", xid_mod);
    formula_names_mod.addImport("zlsx_casefold", unicode_mod);
    formula_names_mod.addImport("zlsx_casing", casing_mod);
    addOracleFixtures(b, formula_names_mod);
    const formula_names_tests = b.addTest(.{ .root_module = formula_names_mod });
    test_step.dependOn(&b.addRunArtifact(formula_names_tests).step);

    // M4b3: the name/3D fuzz target (§8.1) — no defined-name attribute
    // string and no 3D span may panic, leak, or resolve two ways.
    const formula_names_fuzz_mod = b.createModule(.{
        .root_source_file = b.path("src/formula/names.zig"),
        .target = target,
        .optimize = optimize,
        .fuzz = true,
    });
    formula_names_fuzz_mod.addImport("zlsx_control", control_mod);
    formula_names_fuzz_mod.addImport("zlsx_refs", refs_mod);
    formula_names_fuzz_mod.addImport("zlsx_xid", xid_mod);
    formula_names_fuzz_mod.addImport("zlsx_casefold", unicode_mod);
    formula_names_fuzz_mod.addImport("zlsx_casing", casing_mod);
    addOracleFixtures(b, formula_names_fuzz_mod);
    const formula_names_fuzz_tests = b.addTest(.{
        .root_module = formula_names_fuzz_mod,
        .test_runner = fuzz_test_runner,
    });
    fuzz_step.dependOn(&b.addRunArtifact(formula_names_fuzz_tests).step);

    // M4b1: the decode fuzz target (§8.1) — no input may panic, leak,
    // or produce a decode that differs between two runs.
    const formula_decode_fuzz_mod = b.createModule(.{
        .root_source_file = b.path("src/formula/decode.zig"),
        .target = target,
        .optimize = optimize,
        .fuzz = true,
    });
    formula_decode_fuzz_mod.addImport("zlsx_control", control_mod);
    formula_decode_fuzz_mod.addImport("zlsx_refs", refs_mod);
    formula_decode_fuzz_mod.addImport("zlsx_xid", xid_mod);
    formula_decode_fuzz_mod.addImport("zlsx_casefold", unicode_mod);
    formula_decode_fuzz_mod.addImport("zlsx_casing", casing_mod);
    addOracleFixtures(b, formula_decode_fuzz_mod);
    const formula_decode_fuzz_tests = b.addTest(.{
        .root_module = formula_decode_fuzz_mod,
        .test_runner = fuzz_test_runner,
    });
    fuzz_step.dependOn(&b.addRunArtifact(formula_decode_fuzz_tests).step);

    // M4b1: the decoded symbol layer.
    const formula_symbols_mod = b.createModule(.{
        .root_source_file = b.path("src/formula/symbols.zig"),
        .target = target,
        .optimize = optimize,
    });
    formula_symbols_mod.addImport("zlsx_control", control_mod);
    formula_symbols_mod.addImport("zlsx_refs", refs_mod);
    formula_symbols_mod.addImport("zlsx_xid", xid_mod);
    formula_symbols_mod.addImport("zlsx_casefold", unicode_mod);
    formula_symbols_mod.addImport("zlsx_casing", casing_mod);
    addOracleFixtures(b, formula_symbols_mod);
    const formula_symbols_tests = b.addTest(.{ .root_module = formula_symbols_mod });
    test_step.dependOn(&b.addRunArtifact(formula_symbols_tests).step);

    // M4b1: the same root, as the module `pkg/` imports — the ONE place
    // the package layer and the engine meet (§5.6a). One module rather
    // than four because a file compiled into two modules is two
    // distinct types, and an adapter naming `env` and `symbols`
    // separately would build an `EvalEnv` the evaluator could not take.
    //
    // It deliberately does NOT declare the oracle fixtures: they are
    // imported by the TEST sections of `value.zig` and `symbols.zig` and
    // are needed only in a test build rooted there.
    const formula_pkg_mod = b.createModule(.{
        .root_source_file = b.path("src/formula/symbols.zig"),
        .target = target,
        .optimize = optimize,
    });
    formula_pkg_mod.addImport("zlsx_control", control_mod);
    formula_pkg_mod.addImport("zlsx_refs", refs_mod);
    formula_pkg_mod.addImport("zlsx_xid", xid_mod);
    // M4f: the package-facing formula module reaches `registry.zig`,
    // which now calls `casing_v1` for UPPER and LOWER. Unlike the fold,
    // this one cannot be confined to a test block — it is what two
    // shipped functions compute with.
    formula_pkg_mod.addImport("zlsx_casing", casing_mod);
    // `src/xlsx.zig` re-exports the rewriter from here rather than by
    // relative path: `rewriter.zig` and the engine module both contain
    // `tokenizer.zig`, and a file may belong to only one module. Wired
    // at the declaration site of the engine module because that is the
    // constraint's home — see `xlsx.zig`'s `formula_rewriter`.
    zlsx_mod.addImport("zlsx_formula", formula_pkg_mod);

    // M3a2: the non-finite escape fuzz target (§8.1). No evaluation of
    // any input may produce a non-finite number, a zero-dimension
    // matrix, a panic, or a leak.
    const formula_eval_fuzz_mod = b.createModule(.{
        .root_source_file = b.path("src/formula/eval.zig"),
        .target = target,
        .optimize = optimize,
        .fuzz = true,
    });
    formula_eval_fuzz_mod.addImport("zlsx_control", control_mod);
    formula_eval_fuzz_mod.addImport("zlsx_refs", refs_mod);
    formula_eval_fuzz_mod.addImport("zlsx_xid", xid_mod);
    formula_eval_fuzz_mod.addImport("zlsx_casefold", unicode_mod);
    formula_eval_fuzz_mod.addImport("zlsx_casing", casing_mod);
    addOracleFixtures(b, formula_eval_fuzz_mod);
    const formula_eval_fuzz_tests = b.addTest(.{
        .root_module = formula_eval_fuzz_mod,
        .test_runner = fuzz_test_runner,
    });
    fuzz_step.dependOn(&b.addRunArtifact(formula_eval_fuzz_tests).step);

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
    formula_rewriter_mod.addImport("zlsx_control", control_mod);
    formula_rewriter_mod.addImport("zlsx_refs", refs_mod);
    formula_rewriter_mod.addImport("zlsx_xid", xid_mod);
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
    package_mod.addImport("zlsx_control", control_mod);
    package_mod.addImport("zlsx", zlsx_mod);
    // B3 iter-wr-1: workbook.zig stages strings through the shared
    // SST plan substrate via a named import (`@import("zlsx_sst_plan")`)
    // so the file is owned by exactly one module — using a relative
    // `@import("sst_plan.zig")` would make pkg/sst_plan.zig a member
    // of every module that reaches workbook.zig, which collides with
    // the dedicated `zlsx_sst_plan` module declaration.
    package_mod.addImport("zlsx_sst_plan", sst_plan_mod);
    // B3 iter-wr-2: styles plan substrate. Workbook gains
    // `addStyle` / `addDxf` / `internNumFmt` thin pass-throughs that
    // delegate to a `StylesPlan`, mirroring the SST-plan wiring.
    package_mod.addImport("zlsx_styles_plan", styles_plan_mod);
    package_mod.addImport("zlsx_workbook_xml_plan", workbook_xml_plan_mod);
    package_mod.addImport("zlsx_zip", zip_mod);
    package_mod.addImport("zlsx_sheet_plan", sheet_plan_mod);
    package_mod.addImport("zlsx_fresh_emit", fresh_emit_mod);
    package_mod.addImport("zlsx_nfc", nfc_mod);
    package_mod.addImport("zlsx_refs", refs_mod);
    // M4b1: the EvalEnv adapter. `pkg/workbook.zig` reaches the whole
    // engine through this one module — see the import's comment there
    // for why it is one and not four.
    package_mod.addImport("zlsx_formula", formula_pkg_mod);

    // After the B2 iter-er-0 Editor relocation, `cli_mod` reaches
    // Editor through `zlsx_pkg` and xlsx via the named `zlsx` dep.
    // Wired here (post `package_mod` declaration) instead of next to
    // the `cli_mod` block above where `package_mod` isn't in scope yet.
    cli_mod.addImport("zlsx", zlsx_mod);
    cli_mod.addImport("zlsx_pkg", package_mod);
    corpus_mod.addImport("zlsx_pkg", package_mod);

    // ─── M5c: `zlsx_recalc`, the third public module (§5.10) ────
    //
    // Sits ABOVE both public modules and imports each by name. Putting
    // the composition in either of them would close a loop that today
    // runs one way (`zlsx_pkg → zlsx`) — and that direction is load
    // bearing: `pkg/zip.zig` and `pkg/fresh_emit.zig` are stdlib-only
    // and take deflate as a function pointer precisely to keep the
    // graph a DAG.
    //
    // Rooted at top-level `recalc/` for the same reason `unicode/` and
    // `refs/` are: a root inside `src/` or `pkg/` would put the file in
    // that tree's module too, and a file belongs to exactly one.
    //
    // M5c is the shell — the graph, and the buffer handoff it makes
    // possible. `writerSaveWithRecalc` and the `tests/consumer`
    // dependency test land at M5d3.
    const recalc_mod = b.addModule("zlsx_recalc", .{
        .root_source_file = b.path("recalc/recalc.zig"),
        .target = target,
        .optimize = optimize,
    });
    recalc_mod.addImport("zlsx_control", control_mod);
    recalc_mod.addImport("zlsx", zlsx_mod);
    recalc_mod.addImport("zlsx_pkg", package_mod);

    // The module-graph gate. Zig would eventually choke on a cyclic
    // `-M` graph, but "eventually" is a confusing compiler error deep in
    // a build command; this fails at graph-construction time and names
    // the edge. Wired for all three public modules, not just the new
    // one — the cycle a future edit introduces is as likely to be
    // `zlsx → zlsx_recalc` as the other way round.
    assertAcyclicModules(b, "zlsx", zlsx_mod);
    assertAcyclicModules(b, "zlsx_pkg", package_mod);
    assertAcyclicModules(b, "zlsx_recalc", recalc_mod);

    const recalc_tests_mod = b.createModule(.{
        .root_source_file = b.path("recalc/recalc.zig"),
        .target = target,
        .optimize = optimize,
    });
    recalc_tests_mod.addImport("zlsx_control", control_mod);
    recalc_tests_mod.addImport("zlsx", zlsx_mod);
    recalc_tests_mod.addImport("zlsx_pkg", package_mod);
    const recalc_tests = b.addTest(.{ .root_module = recalc_tests_mod });
    test_step.dependOn(&b.addRunArtifact(recalc_tests).step);

    // C2a: standalone `zlsx-extract-images` binary that drives the
    // package layer (PartStore + imageParts) without going through
    // Editor / Book.
    //
    // It was split out because `cli_mod` + `zlsx_pkg` + `writer` could
    // not coexist in one compilation under Zig 0.15.2 — every file that
    // `@import("writer")`ed ended up claimed by both writer's tree and
    // zlsx_pkg's tree.
    //
    // **That constraint no longer holds on 0.16** (verified by probe:
    // adding `cli_mod.addImport("writer", writer_mod)` on top of the
    // existing `zlsx` + `zlsx_pkg` imports builds clean and keeps
    // 1029/1029 tests green). So merging this back into the CLI as a
    // subcommand is now *possible*. It is deliberately NOT done here:
    // `zlsx-extract-images` is a shipped binary and removing it is a
    // user-visible packaging change, not a build-graph cleanup. Left as
    // an explicit product decision.
    //
    // What downstream consumers care about — importing `zlsx` and
    // `zlsx_pkg` together — is gated by `tests/consumer/`.
    const extract_images_mod = b.createModule(.{
        .root_source_file = b.path("src/extract_images_main.zig"),
        .target = target,
        .optimize = optimize,
    });
    extract_images_mod.addImport("zlsx_control", control_mod);
    extract_images_mod.addImport("zlsx_pkg", package_mod);
    const extract_images_exe = b.addExecutable(.{
        .name = "zlsx-extract-images",
        .root_module = extract_images_mod,
    });
    b.installArtifact(extract_images_exe);

    // emb-4 compat-matrix tooling. `emb4-fixture` writes a small
    // .xlsx exercising the full setEmbeddings surface; the user
    // hand-drives it through Excel mac, Excel Win, LibreOffice, and
    // Numbers per docs/plans/emb-4-compat-matrix.md, then
    // `emb4-verify` re-opens the round-tripped file and reports
    // whether the workbook→index rel + embedding parts survived.
    const emb4_fixture_mod = b.createModule(.{
        .root_source_file = b.path("tests/emb-4/fixture_gen.zig"),
        .target = target,
        .optimize = optimize,
    });
    emb4_fixture_mod.addImport("zlsx_control", control_mod);
    emb4_fixture_mod.addImport("zlsx_pkg", package_mod);
    emb4_fixture_mod.addImport("zlsx", zlsx_mod);
    const emb4_fixture_exe = b.addExecutable(.{
        .name = "zlsx-emb4-fixture",
        .root_module = emb4_fixture_mod,
    });
    const emb4_fixture_run = b.addRunArtifact(emb4_fixture_exe);
    if (b.args) |args| emb4_fixture_run.addArgs(args);
    const emb4_fixture_step = b.step(
        "emb4-fixture",
        "Write a small embeddings-bearing .xlsx for the emb-4 compat matrix (arg: out path)",
    );
    emb4_fixture_step.dependOn(&emb4_fixture_run.step);

    const emb4_verify_mod = b.createModule(.{
        .root_source_file = b.path("tests/emb-4/verify.zig"),
        .target = target,
        .optimize = optimize,
    });
    emb4_verify_mod.addImport("zlsx_control", control_mod);
    emb4_verify_mod.addImport("zlsx_pkg", package_mod);
    const emb4_verify_exe = b.addExecutable(.{
        .name = "zlsx-emb4-verify",
        .root_module = emb4_verify_mod,
    });
    const emb4_verify_run = b.addRunArtifact(emb4_verify_exe);
    if (b.args) |args| emb4_verify_run.addArgs(args);
    const emb4_verify_step = b.step(
        "emb4-verify",
        "Verify embedding parts survived a round-trip through an external tool (arg: file path)",
    );
    emb4_verify_step.dependOn(&emb4_verify_run.step);

    // `emb4-passive-save` is the matrix control leg: open → save unedited
    // through zlsx's own delta-on-bytes writer, confirming the part format
    // round-trips through zlsx itself (so a strip elsewhere is a tool property).
    const emb4_passive_mod = b.createModule(.{
        .root_source_file = b.path("tests/emb-4/passive_save.zig"),
        .target = target,
        .optimize = optimize,
    });
    emb4_passive_mod.addImport("zlsx_control", control_mod);
    emb4_passive_mod.addImport("zlsx_pkg", package_mod);
    const emb4_passive_exe = b.addExecutable(.{
        .name = "zlsx-emb4-passive-save",
        .root_module = emb4_passive_mod,
    });
    const emb4_passive_run = b.addRunArtifact(emb4_passive_exe);
    if (b.args) |args| emb4_passive_run.addArgs(args);
    const emb4_passive_step = b.step(
        "emb4-passive-save",
        "Control leg: open then save unedited via zlsx, preserving embeddings (args: in out)",
    );
    emb4_passive_step.dependOn(&emb4_passive_run.step);

    // The helpers are also installed into zig-out/bin so the matrix runner can
    // invoke them directly. Going through `zig build emb4-verify` collapses any
    // non-zero verifier exit into build-failure 1, which erases exactly the
    // distinction the matrix exists to record: 3 STRIPPED vs 4 PARTS-ONLY vs
    // 5 ORPHANED-REL are all different verdicts, and 3 is an *expected* result
    // for the archive-rebuilding tools.
    const emb4_tools_step = b.step(
        "emb4-tools",
        "Install the emb-4 matrix helpers into zig-out/bin (preserves their exit codes)",
    );
    emb4_tools_step.dependOn(&b.addInstallArtifact(emb4_fixture_exe, .{}).step);
    emb4_tools_step.dependOn(&b.addInstallArtifact(emb4_verify_exe, .{}).step);
    emb4_tools_step.dependOn(&b.addInstallArtifact(emb4_passive_exe, .{}).step);

    // emb-4B carrier-survival tooling. emb-4 established that the custom
    // OPC part does not survive Numbers or LibreOffice; emb-4B measures
    // whether any *other* place in the package does, so a small recovery
    // record can ride in a second carrier. See tests/emb-4b/carriers.zig.
    //
    // The carrier catalogue is its own named module because both the
    // generator and the verifier import it, and a file reachable from two
    // module package trees is a hard error ("file exists in modules ...").
    const emb4b_carriers_mod = b.createModule(.{
        .root_source_file = b.path("tests/emb-4b/carriers.zig"),
        .target = target,
        .optimize = optimize,
    });
    emb4b_carriers_mod.addImport("zlsx_control", control_mod);

    const emb4b_fixture_mod = b.createModule(.{
        .root_source_file = b.path("tests/emb-4b/carrier_gen.zig"),
        .target = target,
        .optimize = optimize,
    });
    emb4b_fixture_mod.addImport("zlsx_control", control_mod);
    emb4b_fixture_mod.addImport("zlsx_pkg", package_mod);
    emb4b_fixture_mod.addImport("zlsx", zlsx_mod);
    emb4b_fixture_mod.addImport("emb4b_carriers", emb4b_carriers_mod);
    const emb4b_fixture_exe = b.addExecutable(.{
        .name = "zlsx-emb4b-fixture",
        .root_module = emb4b_fixture_mod,
    });
    const emb4b_fixture_run = b.addRunArtifact(emb4b_fixture_exe);
    if (b.args) |args| emb4b_fixture_run.addArgs(args);
    const emb4b_fixture_step = b.step(
        "emb4b-fixture",
        "Write the emb-4B carrier fixture: one marker in six carriers (arg: out path)",
    );
    emb4b_fixture_step.dependOn(&emb4b_fixture_run.step);

    const emb4b_verify_mod = b.createModule(.{
        .root_source_file = b.path("tests/emb-4b/carrier_verify.zig"),
        .target = target,
        .optimize = optimize,
    });
    emb4b_verify_mod.addImport("zlsx_control", control_mod);
    emb4b_verify_mod.addImport("zlsx_pkg", package_mod);
    emb4b_verify_mod.addImport("emb4b_carriers", emb4b_carriers_mod);
    const emb4b_verify_exe = b.addExecutable(.{
        .name = "zlsx-emb4b-verify",
        .root_module = emb4b_verify_mod,
    });
    const emb4b_verify_run = b.addRunArtifact(emb4b_verify_exe);
    if (b.args) |args| emb4b_verify_run.addArgs(args);
    const emb4b_verify_step = b.step(
        "emb4b-verify",
        "Report which emb-4B carriers survived a round-trip (arg: file path)",
    );
    emb4b_verify_step.dependOn(&emb4b_verify_run.step);

    // Same exit-code reasoning as emb4-tools: the verifier's exit code IS
    // the carrier-loss count, and `zig build` would flatten it to 1.
    const emb4b_tools_step = b.step(
        "emb4b-tools",
        "Install the emb-4B carrier helpers into zig-out/bin (preserves their exit codes)",
    );
    emb4b_tools_step.dependOn(&b.addInstallArtifact(emb4b_fixture_exe, .{}).step);
    emb4b_tools_step.dependOn(&b.addInstallArtifact(emb4b_verify_exe, .{}).step);

    // The carrier catalogue and the workbook.xml patch helpers carry unit
    // tests; wire them into `zig build test` so a regression in the fixture
    // generator fails the gate rather than waiting for a matrix run.
    const emb4b_carriers_tests_mod = b.createModule(.{
        .root_source_file = b.path("tests/emb-4b/carriers.zig"),
        .target = target,
        .optimize = optimize,
    });
    emb4b_carriers_tests_mod.addImport("zlsx_control", control_mod);
    const emb4b_carriers_tests = b.addTest(.{ .root_module = emb4b_carriers_tests_mod });
    test_step.dependOn(&b.addRunArtifact(emb4b_carriers_tests).step);

    const emb4b_gen_tests_mod = b.createModule(.{
        .root_source_file = b.path("tests/emb-4b/carrier_gen.zig"),
        .target = target,
        .optimize = optimize,
    });
    emb4b_gen_tests_mod.addImport("zlsx_control", control_mod);
    emb4b_gen_tests_mod.addImport("zlsx_pkg", package_mod);
    emb4b_gen_tests_mod.addImport("zlsx", zlsx_mod);
    emb4b_gen_tests_mod.addImport("emb4b_carriers", emb4b_carriers_mod);
    const emb4b_gen_tests = b.addTest(.{ .root_module = emb4b_gen_tests_mod });
    test_step.dependOn(&b.addRunArtifact(emb4b_gen_tests).step);

    // Per-source-file test targets so each module gets its own test
    // binary (matches the rest of build.zig's pattern).
    const package_store_tests_mod = b.createModule(.{
        .root_source_file = b.path("pkg/store.zig"),
        .target = target,
        .optimize = optimize,
    });
    package_store_tests_mod.addImport("zlsx_control", control_mod);
    package_store_tests_mod.addImport("zlsx", zlsx_mod);
    const package_store_tests = b.addTest(.{ .root_module = package_store_tests_mod });
    test_step.dependOn(&b.addRunArtifact(package_store_tests).step);

    const package_drawings_tests_mod = b.createModule(.{
        .root_source_file = b.path("pkg/drawings.zig"),
        .target = target,
        .optimize = optimize,
    });
    package_drawings_tests_mod.addImport("zlsx_control", control_mod);
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
    package_typed_parts_tests_mod.addImport("zlsx_control", control_mod);
    const package_typed_parts_tests = b.addTest(.{ .root_module = package_typed_parts_tests_mod });
    test_step.dependOn(&b.addRunArtifact(package_typed_parts_tests).step);

    // pkg/embedding_part.zig — embedding wire-format primitives +
    // XML manifest emit/parse + text canonicalization (emb-1a/1b).
    // Stdlib + zlsx_nfc only; no store/workbook import, so it gets a
    // standalone test binary.
    const embedding_part_tests_mod = b.createModule(.{
        .root_source_file = b.path("pkg/embedding_part.zig"),
        .target = target,
        .optimize = optimize,
    });
    embedding_part_tests_mod.addImport("zlsx_control", control_mod);
    embedding_part_tests_mod.addImport("zlsx_nfc", nfc_mod);
    embedding_part_tests_mod.addImport("zlsx_refs", refs_mod);
    const embedding_part_tests = b.addTest(.{ .root_module = embedding_part_tests_mod });
    test_step.dependOn(&b.addRunArtifact(embedding_part_tests).step);

    // pkg/recovery_record.zig — ER recovery-record codec + carrier
    // extraction. Std-only, so it needs no imports; its own target
    // keeps a failure attributable to the codec rather than to a
    // workbook round-trip.
    const recovery_record_tests_mod = b.createModule(.{
        .root_source_file = b.path("pkg/recovery_record.zig"),
        .target = target,
        .optimize = optimize,
    });
    recovery_record_tests_mod.addImport("zlsx_control", control_mod);
    const recovery_record_tests = b.addTest(.{ .root_module = recovery_record_tests_mod });
    test_step.dependOn(&b.addRunArtifact(recovery_record_tests).step);

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
    package_workbook_tests_mod.addImport("zlsx_control", control_mod);
    package_workbook_tests_mod.addImport("zlsx", zlsx_mod);
    package_workbook_tests_mod.addImport("zlsx_sst_plan", sst_plan_mod);
    package_workbook_tests_mod.addImport("zlsx_styles_plan", styles_plan_mod);
    package_workbook_tests_mod.addImport("zlsx_workbook_xml_plan", workbook_xml_plan_mod);
    package_workbook_tests_mod.addImport("zlsx_zip", zip_mod);
    package_workbook_tests_mod.addImport("zlsx_sheet_plan", sheet_plan_mod);
    package_workbook_tests_mod.addImport("zlsx_fresh_emit", fresh_emit_mod);
    package_workbook_tests_mod.addImport("zlsx_nfc", nfc_mod);
    package_workbook_tests_mod.addImport("zlsx_refs", refs_mod);
    // M4b1: this target is where the adapter's own tests run — the
    // shared `EvalEnv` suite, the decode round-trips, and the corpus
    // sweep.
    package_workbook_tests_mod.addImport("zlsx_formula", formula_pkg_mod);
    const package_workbook_tests = b.addTest(.{ .root_module = package_workbook_tests_mod });
    test_step.dependOn(&b.addRunArtifact(package_workbook_tests).step);

    // pkg/recalc_txn.zig — M5b2's prepare/swap transaction. Its own test
    // binary rather than a section of the workbook target: nothing that is
    // already a test root reaches it (`workbook.zig` imports it, but only
    // through one method body, and a file no analysis reaches is a file
    // whose tests never run — the lesson `src/dbx.zig` cost).
    const package_recalc_txn_tests_mod = b.createModule(.{
        .root_source_file = b.path("pkg/recalc_txn.zig"),
        .target = target,
        .optimize = optimize,
    });
    package_recalc_txn_tests_mod.addImport("zlsx_control", control_mod);
    package_recalc_txn_tests_mod.addImport("zlsx", zlsx_mod);
    package_recalc_txn_tests_mod.addImport("zlsx_sst_plan", sst_plan_mod);
    package_recalc_txn_tests_mod.addImport("zlsx_styles_plan", styles_plan_mod);
    package_recalc_txn_tests_mod.addImport("zlsx_workbook_xml_plan", workbook_xml_plan_mod);
    package_recalc_txn_tests_mod.addImport("zlsx_zip", zip_mod);
    package_recalc_txn_tests_mod.addImport("zlsx_sheet_plan", sheet_plan_mod);
    package_recalc_txn_tests_mod.addImport("zlsx_fresh_emit", fresh_emit_mod);
    package_recalc_txn_tests_mod.addImport("zlsx_nfc", nfc_mod);
    package_recalc_txn_tests_mod.addImport("zlsx_refs", refs_mod);
    package_recalc_txn_tests_mod.addImport("zlsx_formula", formula_pkg_mod);
    const package_recalc_txn_tests = b.addTest(.{ .root_module = package_recalc_txn_tests_mod });
    test_step.dependOn(&b.addRunArtifact(package_recalc_txn_tests).step);

    // pkg/recalc_run.zig — M5d2's pipeline and §5.7.9's file
    // transaction. Its own root for the same reason `recalc_txn` has
    // one: `workbook.zig` reaches it through two forwarder bodies, and
    // a file no analysis reaches is a file whose tests never run.
    const package_recalc_run_tests_mod = b.createModule(.{
        .root_source_file = b.path("pkg/recalc_run.zig"),
        .target = target,
        .optimize = optimize,
    });
    package_recalc_run_tests_mod.addImport("zlsx_control", control_mod);
    package_recalc_run_tests_mod.addImport("zlsx", zlsx_mod);
    package_recalc_run_tests_mod.addImport("zlsx_sst_plan", sst_plan_mod);
    package_recalc_run_tests_mod.addImport("zlsx_styles_plan", styles_plan_mod);
    package_recalc_run_tests_mod.addImport("zlsx_workbook_xml_plan", workbook_xml_plan_mod);
    package_recalc_run_tests_mod.addImport("zlsx_zip", zip_mod);
    package_recalc_run_tests_mod.addImport("zlsx_sheet_plan", sheet_plan_mod);
    package_recalc_run_tests_mod.addImport("zlsx_fresh_emit", fresh_emit_mod);
    package_recalc_run_tests_mod.addImport("zlsx_nfc", nfc_mod);
    package_recalc_run_tests_mod.addImport("zlsx_refs", refs_mod);
    package_recalc_run_tests_mod.addImport("zlsx_formula", formula_pkg_mod);
    const package_recalc_run_tests = b.addTest(.{ .root_module = package_recalc_run_tests_mod });
    test_step.dependOn(&b.addRunArtifact(package_recalc_run_tests).step);

    // pkg/editor.zig had no test target at all, so its inline tests
    // were never collected: Zig gathers tests from the root file and
    // the files it imports *within the same module*, and nothing that
    // was already a test root reaches editor.zig (workbook.zig does not
    // import it — the dependency runs the other way). Same module wiring
    // as the workbook target above, since editor.zig pulls workbook.zig
    // in as a plain file import.
    const package_editor_tests_mod = b.createModule(.{
        .root_source_file = b.path("pkg/editor.zig"),
        .target = target,
        .optimize = optimize,
    });
    package_editor_tests_mod.addImport("zlsx_control", control_mod);
    package_editor_tests_mod.addImport("zlsx", zlsx_mod);
    package_editor_tests_mod.addImport("zlsx_sst_plan", sst_plan_mod);
    package_editor_tests_mod.addImport("zlsx_styles_plan", styles_plan_mod);
    package_editor_tests_mod.addImport("zlsx_workbook_xml_plan", workbook_xml_plan_mod);
    package_editor_tests_mod.addImport("zlsx_zip", zip_mod);
    package_editor_tests_mod.addImport("zlsx_sheet_plan", sheet_plan_mod);
    package_editor_tests_mod.addImport("zlsx_fresh_emit", fresh_emit_mod);
    package_editor_tests_mod.addImport("zlsx_nfc", nfc_mod);
    package_editor_tests_mod.addImport("zlsx_refs", refs_mod);
    package_editor_tests_mod.addImport("zlsx_formula", formula_pkg_mod);
    const package_editor_tests = b.addTest(.{ .root_module = package_editor_tests_mod });
    test_step.dependOn(&b.addRunArtifact(package_editor_tests).step);

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
    package_fuzz_mod.addImport("zlsx_control", control_mod);
    package_fuzz_mod.addImport("zlsx", zlsx_mod);
    const package_fuzz_tests = b.addTest(.{
        .root_module = package_fuzz_mod,
        .test_runner = fuzz_test_runner,
    });
    fuzz_step.dependOn(&b.addRunArtifact(package_fuzz_tests).step);

    // Byte-walker fuzz modules. These four files are ~3000 LOC of
    // hand-rolled byte-splicing over attacker-controlled XML and had no
    // fuzz coverage at all; #125 found four unhandled coordinate
    // elements in them by reading, and the first fuzz run found an
    // out-of-bounds read in `matchTagAt` within seconds.
    //
    // They need their own modules because nothing already wired into
    // `fuzz` reaches them: `package_fuzz_mod` is rooted at store.zig,
    // which does not import the edit walkers. A target the fuzz step
    // cannot reach is a target that never runs — the same trap
    // scripts/bench_ci.sh fell into.
    //
    // Each is std-only or std+zlsx, so the modules stay cheap.
    const walker_fuzz = [_]struct { name: []const u8, path: []const u8, needs_zlsx: bool }{
        .{ .name = "sheet_edit", .path = "pkg/sheet_edit.zig", .needs_zlsx = true },
        .{ .name = "table_edit", .path = "pkg/table_edit.zig", .needs_zlsx = true },
        .{ .name = "drawing_edit", .path = "pkg/drawing_edit.zig", .needs_zlsx = false },
        .{ .name = "vml_edit", .path = "pkg/vml_edit.zig", .needs_zlsx = false },
    };
    for (walker_fuzz) |w| {
        const mod = b.createModule(.{
            .root_source_file = b.path(w.path),
            .target = target,
            .optimize = optimize,
            .fuzz = true,
        });
        if (w.needs_zlsx) mod.addImport("zlsx", zlsx_mod);
        mod.addImport("zlsx_refs", refs_mod);
        const t = b.addTest(.{ .root_module = mod, .test_runner = fuzz_test_runner });
        fuzz_step.dependOn(&b.addRunArtifact(t).step);
    }

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
    package_corpus_mod.addImport("zlsx_control", control_mod);
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
    typed_parts_corpus_mod.addImport("zlsx_control", control_mod);
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
    workbook_corpus_mod.addImport("zlsx_control", control_mod);
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
    c_abi_mod.addImport("zlsx_control", control_mod);
    c_abi_mod.addImport("build_options", build_options_mod);
    c_abi_mod.addImport("fuzz_config", fuzz_config_mod);
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
    // because **each probe must measure its own RSS in its own
    // process** — a delta is only meaningful against a process that
    // did nothing else. The test spawns each probe as a subprocess
    // and compares the two deltas as a ratio.
    //
    // M5c correction: this comment used to say `zlsx` and `zlsx_pkg`
    // "cannot coexist in one binary". They can, and did even then —
    // `cli_mod`, `corpus_mod` and `package_mod` itself all import
    // both. What could not coexist under Zig 0.15.2 was a *file*
    // claimed by two module trees, which is a different statement and
    // is resolved on 0.16 (`AGENTS.md` "Three-module collision",
    // marked history there). `zlsx_recalc` imports both public modules
    // by name and is gated on the graph staying acyclic
    // (`assertAcyclicModules`), so the old claim is not merely stale —
    // it contradicts a shipped module.
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
    bench_rss_mod.addImport("zlsx_control", control_mod);
    const bench_synth_mod = b.createModule(.{
        .root_source_file = b.path("tests/bench/synth_100k_x_10.zig"),
        .target = target,
        .optimize = bench_optimize,
    });
    bench_synth_mod.addImport("zlsx_control", control_mod);
    bench_synth_mod.addImport("zlsx", zlsx_mod);

    // Probe 1 — synth. Pulls `zlsx` (writer) only; no pkg.
    const probe_synth_mod = b.createModule(.{
        .root_source_file = b.path("tests/bench/rss_probe_synth.zig"),
        .target = target,
        .optimize = bench_optimize,
    });
    probe_synth_mod.addImport("zlsx_control", control_mod);
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
    probe_book_mod.addImport("zlsx_control", control_mod);
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
    probe_wb_mod.addImport("zlsx_control", control_mod);
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
    bench_workbook_rss_mod.addImport("zlsx_control", control_mod);
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

    // ─── B2 iter-er-3: appendRows wall-clock baseline ──────────
    //
    // Off the default `test` step. Builds a 100k×5 fixture and
    // times Editor.open + appendRows + save. Run with:
    //   zig build bench-append-rows -- <tmpdir> [rows]
    //
    // iter-er-3's walk-away gate is "the rebased Editor.appendRows
    // (route through Worksheet.appendRows on the Workbook overlay)
    // stays within 1.10× of the legacy substring path on this
    // fixture." Capture baseline → implement rebase → re-run →
    // compare.
    const bench_append_rows_mod = b.createModule(.{
        .root_source_file = b.path("tests/bench/bench_append_rows.zig"),
        .target = target,
        .optimize = bench_optimize,
        .single_threaded = single_threaded,
    });
    bench_append_rows_mod.addImport("zlsx_control", control_mod);
    bench_append_rows_mod.addImport("zlsx", zlsx_mod);
    bench_append_rows_mod.addImport("zlsx_pkg", package_mod);
    const bench_append_rows_exe = b.addExecutable(.{
        .name = "zlsx-bench-append-rows",
        .root_module = bench_append_rows_mod,
    });
    const bench_append_rows_install = b.addInstallArtifact(bench_append_rows_exe, .{});
    const bench_append_rows_run = b.addRunArtifact(bench_append_rows_exe);
    if (b.args) |args| bench_append_rows_run.addArgs(args);
    const bench_append_rows_step = b.step(
        "bench-append-rows",
        "B2 iter-er-3 baseline: time Editor.appendRows on a 100k×5 fixture (off the default test path)",
    );
    bench_append_rows_step.dependOn(&bench_append_rows_install.step);
    bench_append_rows_step.dependOn(&bench_append_rows_run.step);

    // ── bench harness executables ────────────────────────────────────
    //
    // These two were previously built ONLY by `scripts/bench_ci.sh`'s
    // hand-rolled `build-exe` invocations, which restated the whole
    // module graph by hand. Two consequences, both of which actually
    // bit during the 0.16 migration:
    //
    //   * every new named module had to be added to the script as well
    //     as here, or the bench job broke on its own;
    //   * no `zig build` step compiled these files, so they silently
    //     kept a stale API through a tree-wide migration and only the
    //     nightly bench job noticed.
    //
    // Defining them here makes `zig build` the single source of truth
    // for the module graph, and hanging them off `test_step` means a
    // signature change that misses them is a local test failure rather
    // than a CI-only surprise. `bench-exes` installs them for
    // bench_ci.sh to run under hyperfine.
    //
    // **These take `optimize`, not `bench_optimize`** (M5d3). The RSS
    // probes above are pinned to ReleaseSafe on purpose — an RSS number
    // is only meaningful with the overflow checks a production caller
    // runs. The hyperfine lane is the opposite case: `bench_ci.sh:34`
    // passes `-Doptimize=ReleaseFast` and §9 labels the lane
    // "hyperfine ReleaseFast", and a hardcoded `bench_optimize` silently
    // discarded that flag — every wall-clock number this repo has
    // recorded was measured in ReleaseSafe under a ReleaseFast label.
    const bench_read_mod = b.createModule(.{
        .root_source_file = b.path("tests/bench/bench_zlsx.zig"),
        .target = target,
        .optimize = optimize,
        .single_threaded = single_threaded,
    });
    bench_read_mod.addImport("zlsx_control", control_mod);
    bench_read_mod.addImport("zlsx", zlsx_mod);
    const bench_read_exe = b.addExecutable(.{
        .name = "zlsx-bench-read",
        .root_module = bench_read_mod,
    });

    const bench_write_mod = b.createModule(.{
        .root_source_file = b.path("tests/bench/bench_write_zlsx.zig"),
        .target = target,
        .optimize = optimize,
        .single_threaded = single_threaded,
    });
    bench_write_mod.addImport("zlsx_control", control_mod);
    bench_write_mod.addImport("zlsx", zlsx_mod);
    const bench_write_exe = b.addExecutable(.{
        .name = "zlsx-bench-write",
        .root_module = bench_write_mod,
    });

    // ── M5d3: §9's named workload and the recalc bench over it ───────
    //
    // The generator is the committed artifact (§9): a fixed-topology
    // F1-mix workbook whose digest identifies the workload the absolute
    // ceilings bind to. It lives in its own module because two things
    // need it — the bench binary, and the unit tests that keep the
    // topology honest.
    const bench_f1_mix_mod = b.createModule(.{
        .root_source_file = b.path("tests/bench/synth_f1_mix.zig"),
        .target = target,
        .optimize = optimize,
        .single_threaded = single_threaded,
    });
    bench_f1_mix_mod.addImport("zlsx_control", control_mod);
    bench_f1_mix_mod.addImport("zlsx", zlsx_mod);

    // Imports `zlsx_recalc`, so the ReleaseFast bench lane compiles the
    // third public module too — a composition that only ever built in
    // Debug unit tests would be a surface nothing releases exercises.
    const bench_recalc_mod = b.createModule(.{
        .root_source_file = b.path("tests/bench/bench_recalc.zig"),
        .target = target,
        .optimize = optimize,
        .single_threaded = single_threaded,
    });
    bench_recalc_mod.addImport("zlsx_control", control_mod);
    bench_recalc_mod.addImport("zlsx_recalc", recalc_mod);
    bench_recalc_mod.addImport("synth_f1_mix", bench_f1_mix_mod);
    const bench_recalc_exe = b.addExecutable(.{
        .name = "zlsx-bench-recalc",
        .root_module = bench_recalc_mod,
    });

    const bench_exes_step = b.step(
        "bench-exes",
        "Build the read + write + recalc bench binaries (consumed by scripts/bench_ci.sh)",
    );
    bench_exes_step.dependOn(&b.addInstallArtifact(bench_read_exe, .{}).step);
    bench_exes_step.dependOn(&b.addInstallArtifact(bench_write_exe, .{}).step);
    bench_exes_step.dependOn(&b.addInstallArtifact(bench_recalc_exe, .{}).step);

    // Compile-only on the default test path: catches API drift without
    // paying for a ReleaseFast link on every `zig build test`.
    test_step.dependOn(&bench_read_exe.step);
    test_step.dependOn(&bench_write_exe.step);
    test_step.dependOn(&bench_recalc_exe.step);

    // Per-module unit tests for the bench helpers (rss + synth).
    // These DO go on the default `test` step — they're cheap and
    // exercise the platform-specific code paths.
    const bench_rss_tests = b.addTest(.{ .root_module = bench_rss_mod });
    test_step.dependOn(&b.addRunArtifact(bench_rss_tests).step);

    const bench_synth_tests = b.addTest(.{ .root_module = bench_synth_mod });
    test_step.dependOn(&b.addRunArtifact(bench_synth_tests).step);

    // The F1-mix generator's own tests: determinism and topology, at a
    // size the default path can afford. The named workload's digest is
    // gated by `zlsx-bench-recalc emit` in the ReleaseFast lane instead
    // — see that file's test-section comment.
    const bench_f1_mix_tests = b.addTest(.{ .root_module = bench_f1_mix_mod });
    test_step.dependOn(&b.addRunArtifact(bench_f1_mix_tests).step);
}

/// The three committed oracle manifests, as anonymous imports.
///
/// `tests/oracle/` is outside every formula module's package tree, so a
/// relative `@embedFile` cannot reach it — and reaching it is the point:
/// every layer that ties itself to the oracle must tie itself to the
/// *committed* manifests rather than to a copy that could drift.
fn addOracleFixtures(b: *std.Build, mod: *std.Build.Module) void {
    mod.addAnonymousImport("oracle_hand_spec_excel", .{
        .root_source_file = b.path("tests/oracle/fixtures/hand_spec_excel.json"),
    });
    mod.addAnonymousImport("oracle_hand_spec_ieee", .{
        .root_source_file = b.path("tests/oracle/fixtures/hand_spec_ieee.json"),
    });
    mod.addAnonymousImport("oracle_libreoffice_suite", .{
        .root_source_file = b.path("tests/oracle/fixtures/libreoffice_oracle_suite.json"),
    });
}

/// Fail the build if `root`'s module import graph contains a cycle
/// (M5c, §5.10's "no cycle" clause).
///
/// The one-way edge `zlsx_pkg → zlsx` is what lets `pkg/zip.zig` and
/// `pkg/fresh_emit.zig` stay stdlib-only and take deflate as a function
/// pointer. `zlsx_recalc` importing both is safe *because* it sits above
/// them; the moment something under them imports it back, that argument
/// is gone — and the symptom would be a compiler error inside a `-M`
/// command line rather than a sentence naming the edge.
///
/// Runs at graph-construction time, so `zig build` — any target — is the
/// gate. Not a test: a broken module graph is not something a test
/// binary gets far enough to run.
fn assertAcyclicModules(b: *std.Build, name: []const u8, root: *std.Build.Module) void {
    var on_path: std.ArrayListUnmanaged(*std.Build.Module) = .empty;
    defer on_path.deinit(b.allocator);
    // Fully-explored nodes. Without it the walk is exponential on a DAG
    // this wide — `zlsx` alone is reached by a dozen modules.
    var done: std.AutoHashMapUnmanaged(*std.Build.Module, void) = .empty;
    defer done.deinit(b.allocator);
    walkModules(b, name, root, &on_path, &done);
}

fn walkModules(
    b: *std.Build,
    name: []const u8,
    m: *std.Build.Module,
    on_path: *std.ArrayListUnmanaged(*std.Build.Module),
    done: *std.AutoHashMapUnmanaged(*std.Build.Module, void),
) void {
    if (done.contains(m)) return;
    for (on_path.items) |p| {
        if (p == m) {
            std.debug.panic(
                "module import cycle reached from `{s}`: a module imports itself " ++
                    "through {d} edge(s). §5.10 requires the public module graph to " ++
                    "stay acyclic.",
                .{ name, on_path.items.len },
            );
        }
    }
    on_path.append(b.allocator, m) catch @panic("OOM");
    for (m.import_table.values()) |dep| walkModules(b, name, dep, on_path, done);
    _ = on_path.pop();
    done.put(b.allocator, m, {}) catch @panic("OOM");
}

/// Parse an optional env-var string that may contain `_` digit
/// separators ("1_000_000"), the spelling the fuzz docs use. Returns
/// null for absent or unparseable input so the caller falls back to the
/// in-source default rather than failing the build on a typo'd shell
/// variable.
fn parseUnderscoredInt(comptime T: type, raw: ?[]const u8) ?T {
    const s = raw orelse return null;
    var digits: [32]u8 = undefined;
    var n: usize = 0;
    for (s) |c| {
        if (c == '_') continue;
        if (n == digits.len) return null;
        digits[n] = c;
        n += 1;
    }
    if (n == 0) return null;
    return std.fmt.parseInt(T, digits[0..n], 10) catch null;
}
