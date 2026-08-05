//! Z2 gate: an external consumer importing ALL THREE public zlsx
//! modules.
//!
//! `build.zig` has long carried a note that `cli_mod`, `zlsx_pkg` and
//! `writer` cannot coexist in one compilation, which is why
//! `zlsx-extract-images` is a separate binary. That constraint is about
//! zlsx's own internal graph. What actually matters to downstream users
//! — and to nemonym, which needs the reader *and* Editor — is whether a
//! consumer can `@import("zlsx")` and `@import("zlsx_pkg")` together.
//!
//! M5d3 adds the third: `zlsx_recalc`, which imports the other two.
//! Inside the repo the graph is gated by `assertAcyclicModules`, but
//! that gate runs over modules zlsx's own `build.zig` constructed. A
//! downstream package resolves them through `b.dependency(...)`, and the
//! composition only compiles if the `zlsx` reached that way is the same
//! module object `zlsx_pkg` and `zlsx_recalc` were built against — two
//! instances would be two structurally-identical `Cell` types. So the
//! §5.10 dependency test belongs here, in a build zlsx does not control,
//! rather than in a unit test that shares its graph.
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
    // All three public modules, in one compilation, from one consumer.
    mod.addImport("zlsx", zlsx_dep.module("zlsx"));
    mod.addImport("zlsx_pkg", zlsx_dep.module("zlsx_pkg"));
    mod.addImport("zlsx_recalc", zlsx_dep.module("zlsx_recalc"));

    const exe = b.addExecutable(.{
        .name = "consumer",
        .root_module = mod,
    });
    b.installArtifact(exe);

    const run = b.addRunArtifact(exe);
    if (b.args) |args| run.addArgs(args);
    b.step("run", "Run the co-import round-trip check").dependOn(&run.step);
}
