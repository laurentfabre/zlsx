//! Z2 gate: an external consumer importing BOTH public zlsx modules.
//!
//! `build.zig` has long carried a note that `cli_mod`, `zlsx_pkg` and
//! `writer` cannot coexist in one compilation, which is why
//! `zlsx-extract-images` is a separate binary. That constraint is about
//! zlsx's own internal graph. What actually matters to downstream users
//! — and to nemonym, which needs the reader *and* Editor — is whether a
//! consumer can `@import("zlsx")` and `@import("zlsx_pkg")` together.
//!
//! This package exists to answer that with a build rather than a claim.

const std = @import("std");

pub fn build(b: *std.Build) void {
    const target = b.standardTargetOptions(.{});
    const optimize = b.standardOptimizeOption(.{});

    const zlsx_dep = b.dependency("zlsx", .{
        .target = target,
        .optimize = optimize,
    });

    const mod = b.createModule(.{
        .root_source_file = b.path("src/main.zig"),
        .target = target,
        .optimize = optimize,
    });
    // Both public modules, in one compilation, from one consumer.
    mod.addImport("zlsx", zlsx_dep.module("zlsx"));
    mod.addImport("zlsx_pkg", zlsx_dep.module("zlsx_pkg"));

    const exe = b.addExecutable(.{
        .name = "consumer",
        .root_module = mod,
    });
    b.installArtifact(exe);

    const run = b.addRunArtifact(exe);
    if (b.args) |args| run.addArgs(args);
    b.step("run", "Run the co-import round-trip check").dependOn(&run.step);
}
