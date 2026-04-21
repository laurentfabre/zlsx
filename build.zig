const std = @import("std");

pub fn build(b: *std.Build) void {
    const target = b.standardTargetOptions(.{});
    const optimize = b.standardOptimizeOption(.{});

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
}
