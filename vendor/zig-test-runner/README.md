# Vendored Zig test runner (fuzz builds only)

`test_runner.zig` here is a **verbatim copy of Zig 0.16.0's
`lib/compiler/test_runner.zig`** with exactly one hunk changed. It is
used only by the two `-ffuzz` test binaries; every other test target in
`build.zig` uses the stock runner that ships with the toolchain.

## Why this exists

Zig 0.16.0 cannot build **any** `-ffuzz` test binary. Its own test
runner fails to compile:

```
lib/compiler/test_runner.zig:566:55: error: expected type
  '*const debug.StackTrace', found '*builtin.StackTrace'
```

`@errorReturnTrace()` returns a `*std.builtin.StackTrace`
(`{ index, instruction_addresses }`), but `std.debug.writeStackTrace`
takes a `*const std.debug.StackTrace` (`{ return_addresses, skipped }`)
— two different structs. The offending line sits in the error branch of
`std.testing.fuzz`'s generated wrapper, which is only instantiated in
fuzz mode. That is why `zig build test` is completely unaffected and
only `zig build fuzz --fuzz` trips over it.

Upstream bug, not a zlsx one. There is no 0.16.x point release carrying
a fix at the time of writing (0.16.0 is the newest published release;
`master` is 0.17.0-dev).

## The patch

Convert rather than mis-pass. `builtin.StackTrace.index` counts how many
addresses were written and may exceed the buffer length when the trace
wrapped — which is precisely what `debug.StackTrace.skipped` encodes, so
the conversion is information-preserving:

```zig
const written = @min(trace.index, trace.instruction_addresses.len);
const converted: std.debug.StackTrace = .{
    .return_addresses = trace.instruction_addresses[0..written],
    .skipped = if (trace.index > trace.instruction_addresses.len)
        @enumFromInt(trace.index - trace.instruction_addresses.len)
    else
        .none,
};
std.debug.writeStackTrace(&converted, stderr) catch break :p;
```

## Removal condition

**Delete this directory and drop the `.test_runner` fields in
`build.zig` as soon as a Zig release builds `-ffuzz` test binaries
without it.** To check after a toolchain bump:

```sh
rm -rf vendor/zig-test-runner
# drop the two `.test_runner = fuzz_test_runner` lines in build.zig
zig build fuzz --fuzz --webui=127.0.0.1:0    # should get past compilation
```

If that works, the vendored copy is dead weight — a 600-line stdlib file
frozen at 0.16.0 is a real maintenance liability and should not outlive
its reason for existing.

## Keeping it honest

The file is otherwise **byte-identical to upstream 0.16.0**. Verify:

```sh
diff ~/.zvm/0.16.0/lib/compiler/test_runner.zig \
     vendor/zig-test-runner/test_runner.zig
```

Expect exactly one hunk, at the `writeStackTrace` call. If a toolchain
bump makes that diff larger, re-vendor from the new version rather than
hand-merging — and re-check whether the patch is still needed at all.
