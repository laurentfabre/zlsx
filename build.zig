const std = @import("std");

pub fn build(b: *std.Build) void {
    const target = b.standardTargetOptions(.{});
    const optimize = b.standardOptimizeOption(.{});
    const single_threaded = b.option(bool, "single-threaded", "Build the CLI and C ABI with -fsingle-threaded (smp_allocator is swapped for page_allocator)");

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

    // CLI: `zlsx` binary, streams xlsx rows to stdout in JSONL / TSV / CSV.
    // `zig build` (default step) installs it at zig-out/bin/zlsx.
    const cli_mod = b.createModule(.{
        .root_source_file = b.path("src/cli.zig"),
        .target = target,
        .optimize = optimize,
        .single_threaded = single_threaded,
    });
    const cli_exe = b.addExecutable(.{ .name = "zlsx", .root_module = cli_mod });
    b.installArtifact(cli_exe);

    const run_cli = b.addRunArtifact(cli_exe);
    if (b.args) |args| run_cli.addArgs(args);
    const run_step = b.step("run", "Build and run the zlsx CLI (args after --)");
    run_step.dependOn(&run_cli.step);

    // CLI unit tests (colLetter, JSON/CSV escapers, arg parser).
    const cli_tests = b.addTest(.{ .root_module = cli_mod });
    test_step.dependOn(&b.addRunArtifact(cli_tests).step);

    // C ABI — both a shared library (for Python / cffi bindings) and a
    // static library (for language toolchains that prefer linking in).
    const c_abi_mod = b.createModule(.{
        .root_source_file = b.path("src/c_abi.zig"),
        .target = target,
        .optimize = optimize,
        .single_threaded = single_threaded,
    });
    const dylib = b.addLibrary(.{
        .name = "zlsx",
        .linkage = .dynamic,
        .root_module = c_abi_mod,
    });
    b.installArtifact(dylib);

    const staticlib = b.addLibrary(.{
        .name = "zlsx",
        .linkage = .static,
        .root_module = c_abi_mod,
    });
    b.installArtifact(staticlib);

    // Unit tests for the ABI layer (version constant, CCell translation,
    // and a corpus-gated end-to-end lifecycle smoke test).
    const c_abi_tests = b.addTest(.{ .root_module = c_abi_mod });
    test_step.dependOn(&b.addRunArtifact(c_abi_tests).step);
}
