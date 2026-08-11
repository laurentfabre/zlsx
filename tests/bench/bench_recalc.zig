//! §9's recalc bench harness (M5d3): the baseline v1's absolute
//! ceilings are measured against.
//!
//! One binary, eight modes, because hyperfine measures whole processes
//! and §9 wants two different numbers out of the same workload:
//!
//!   emit    write the fixture and print its SHA-256 (and, for the
//!           named geometry, refuse if it is not the recorded one)
//!   open    `Workbook.open` and stop — the fixed cost
//!   recalc  open + `Workbook.recalculate` — §9's *evaluate* lane
//!   save    open + `Workbook.saveWithRecalc` — §9's *end-to-end* lane
//!   phases  one instrumented run, phases reported separately
//!   heap    §9.1's RSS lane, attributed: the same first recalc under
//!           an allocator that names which call sites hold how many
//!           bytes at the moment of peak live footprint
//!   fill    M10q's four quantities uncontaminated — capacity, fill,
//!           churn and RSS with no profiler resident beside them
//!   pages   M10t's resident-page curve, sampled off-thread from the
//!           kernel: the only mode that measures pages rather than
//!           inferring them from bytes an allocator was asked for
//!
//! The last three are diagnostics, not lanes: `recalc` under
//! `/usr/bin/time -l` is what the §9.1b budget gates on.
//!
//! **Evaluate time is a difference, not a mode.** No process can
//! recalculate without first opening the archive, so `recalc − open` is
//! how the evaluate ceiling is read off two hyperfine means. That is
//! also why `open` exists at all, and why `bench_ci.sh` runs both sizes:
//! two points an order of magnitude apart separate the fixed cost from
//! the per-cell one, which one point cannot.
//!
//! `phases` decomposes as far as the *public* surface allows —
//! `recalc_run.prepare` is public (M5d2 exported it for exactly this
//! kind of caller), so model+graph+evaluate+stage+txn-prepare is one
//! measurable span and the swap is another, and serialize+commit falls
//! out as `saveWithRecalc − prepare`. Splitting model from evaluate from
//! stage needs timers *inside* `pkg/recalc_run.zig`, which is M5d2's
//! code and not this row's to touch.
//!
//! Allocator is `std.heap.smp_allocator` for the reason the other bench
//! binaries give: it is what a production caller plugs in, and
//! `DebugAllocator`'s per-allocation tracking would dominate a
//! measurement of the pipeline rather than describe it.

const std = @import("std");
const recalc = @import("zlsx_recalc");
const synth = @import("synth_f1_mix");
const crit = @import("synth_criteria_mix");
const text = @import("synth_text_mix");
const registry = @import("synth_registry_mix");

const pkg = recalc.pkg;
const Workbook = recalc.Workbook;

/// Fixed clock and fixed seed. §5.5 makes a run reproducible from its
/// inputs; a bench that read the wall clock would be timing a different
/// run on every invocation, and `TODAY()`-shaped drift is exactly the
/// kind of noise a baseline cannot absorb.
const RUN: recalc.RunInputs = .{
    .now_utc_ms = 1_700_000_000_000,
    .rng_seed = 0x5EED_5D3,
    .limits = .{},
};

const Mode = enum { emit, open, recalc, save, phases, heap, fill, pages };

/// Which generator `emit` runs. Only `emit` cares: the other four modes
/// take a fixture path and measure whatever workbook is behind it,
/// which is why the criteria lane (M7b2) needed a flag here and not a
/// second binary. `text` is M8c's lane and `registry` M9d's, same
/// reasoning.
const Workload = enum { f1, criteria, text, registry };

pub fn main(init: std.process.Init) !u8 {
    const io = init.io;
    const gpa = std.heap.smp_allocator;
    const args = try init.minimal.args.toSlice(init.arena.allocator());

    var out_file: std.Io.File.Writer = .init(std.Io.File.stdout(), io, &.{});
    const w = &out_file.interface;
    defer w.flush() catch {};

    if (args.len < 3) {
        try w.print(
            "usage: {s} <emit|open|recalc|save|phases|heap|fill|pages> <fixture.xlsx>" ++
                " [--workload f1|criteria|text|registry] [--size named|small|tiny] [--rows N] [--out PATH]\n",
            .{args[0]},
        );
        return 2;
    }

    const mode = std.meta.stringToEnum(Mode, args[1]) orelse {
        try w.print("unknown mode: {s}\n", .{args[1]});
        return 2;
    };
    const path = args[2];

    var workload: Workload = .f1;
    var geometry = synth.named;
    var named_size = true;
    var crit_geometry = crit.small;
    var crit_identity = true;
    var text_geometry = text.small;
    var text_identity = true;
    var registry_geometry = registry.small;
    var registry_identity = true;
    var out_path: ?[]const u8 = null;
    var i: usize = 3;
    while (i < args.len) : (i += 1) {
        if (std.mem.eql(u8, args[i], "--workload") and i + 1 < args.len) {
            i += 1;
            workload = std.meta.stringToEnum(Workload, args[i]) orelse {
                try w.print("unknown workload: {s}\n", .{args[i]});
                return 2;
            };
        } else if (std.mem.eql(u8, args[i], "--size") and i + 1 < args.len) {
            i += 1;
            if (std.mem.eql(u8, args[i], "small")) {
                geometry = synth.small;
                named_size = false;
                crit_geometry = crit.small;
                crit_identity = true;
                text_geometry = text.small;
                text_identity = true;
                registry_geometry = registry.small;
                registry_identity = true;
            } else if (std.mem.eql(u8, args[i], "tiny")) {
                geometry = synth.tiny;
                named_size = false;
                crit_geometry = crit.tiny;
                crit_identity = false;
                text_geometry = text.tiny;
                text_identity = false;
                registry_geometry = registry.tiny;
                registry_identity = false;
            } else if (!std.mem.eql(u8, args[i], "named")) {
                try w.print("unknown size: {s}\n", .{args[i]});
                return 2;
            }
        } else if (std.mem.eql(u8, args[i], "--rows") and i + 1 < args.len) {
            // Off the two named sizes: an arbitrary row count, so §9's
            // "10k-vs-100k ≤ 15×" scaling assertion can be swept rather
            // than inferred from two points. Never the named workload,
            // whatever number is passed — the digest gate below would
            // be meaningless otherwise.
            i += 1;
            geometry = .{ .data_rows = try std.fmt.parseInt(u32, args[i], 10) };
            named_size = false;
            crit_geometry = .{ .data_rows = geometry.data_rows };
            crit_identity = false;
            text_geometry = .{ .data_rows = geometry.data_rows };
            text_identity = false;
            registry_geometry = .{ .data_rows = geometry.data_rows };
            registry_identity = false;
        } else if (std.mem.eql(u8, args[i], "--out") and i + 1 < args.len) {
            i += 1;
            out_path = args[i];
        }
    }

    switch (mode) {
        .emit => switch (workload) {
            .f1 => return emitFixture(gpa, io, w, path, geometry, named_size),
            .criteria => {
                // The criteria workload's sizes are `tiny` and `small`
                // (its identity size); the F1 mix's `named` default has
                // no counterpart here, so asking for it is a mistake
                // rather than a mapping.
                if (named_size) {
                    try w.writeAll("the criteria workload needs --size small|tiny or --rows N\n");
                    return 2;
                }
                return emitCriteria(gpa, io, w, path, crit_geometry, crit_identity);
            },
            .text => {
                if (named_size) {
                    try w.writeAll("the text workload needs --size small|tiny or --rows N\n");
                    return 2;
                }
                return emitText(gpa, io, w, path, text_geometry, text_identity);
            },
            .registry => {
                if (named_size) {
                    try w.writeAll("the registry workload needs --size small|tiny or --rows N\n");
                    return 2;
                }
                return emitRegistry(gpa, io, w, path, registry_geometry, registry_identity);
            },
        },
        .open => {
            var wb = try Workbook.open(gpa, io, path);
            wb.deinit();
        },
        .recalc => {
            var wb = try Workbook.open(gpa, io, path);
            defer wb.deinit();
            var report = try wb.recalculate(gpa, io, RUN, .{});
            defer report.deinit(gpa);
            try w.print("cells_written={d}\n", .{report.cells_written});
            // §9.1d's density denominator. Printed after the run so the
            // timed span and the peak instant are both behind us — this
            // lane is measured by `/usr/bin/time -l` around the whole
            // process, and a print cannot move either number.
            try w.print("dependency_edges={d}\n", .{report.dependency_edges});
        },
        .save => {
            const dest = out_path orelse {
                try w.writeAll("save mode needs --out PATH\n");
                return 2;
            };
            var wb = try Workbook.open(gpa, io, path);
            defer wb.deinit();
            var report = try wb.saveWithRecalc(gpa, io, dest, RUN, .{});
            defer report.deinit(gpa);
            try w.print("cells_written={d}\n", .{report.cells_written});
        },
        .phases => return reportPhases(gpa, io, w, path, out_path),
        .heap => return reportHeap(io, w, path),
        .fill => return reportFill(io, w, path),
        .pages => return reportPages(io, w, path),
    }
    return 0;
}

/// §9.1's RSS lane, attributed: the same first recalc the
/// `/usr/bin/time -l` probe measures, run through an allocator that
/// records which call sites hold how many bytes at the moment of peak
/// live footprint. macOS heap tooling cannot see these allocations —
/// `smp_allocator` maps its own pools rather than going through libc
/// malloc — so the attribution has to happen at the `Allocator`
/// boundary, inside the process.
///
/// The numbers describe *live bytes at peak*, not churn: a site that
/// allocates gigabytes but frees promptly is invisible to an RSS
/// ceiling, and a site that allocates once and holds is the whole
/// problem. Churn is reported alongside because a fix that converts a
/// holder into a churner shows up there.
fn reportHeap(io: std.Io, w: *std.Io.Writer, path: []const u8) !u8 {
    var prof: HeapProfiler = .init(std.heap.smp_allocator);
    const gpa = prof.allocator();

    // §9.1 M10q: the profiler above sees an arena's chunk REQUESTS. These
    // tallies see what the arena hands out. The gap between them is
    // where M10p's 1.57 MB came from, and no era height can show it.
    const fill = pkg.fill_probe;
    var tallies: [std.meta.fields(fill.Owner).len]fill.Tally = @splat(.{});
    inline for (std.meta.fields(fill.Owner), 0..) |f, i| {
        _ = fill.install(@field(fill.Owner, f.name), &tallies[i]);
    }
    defer fill.clear();

    const rss_before = peakRssBytes();

    var wb = try Workbook.open(gpa, io, path);
    defer wb.deinit();
    const open_live = prof.live_total;
    const rss_after_open = peakRssBytes();

    var report = try wb.recalculate(gpa, io, RUN, .{});
    defer report.deinit(gpa);
    const rss_after_recalc = peakRssBytes();

    try w.print("live_after_open_bytes={d}\n", .{open_live});
    try prof.report(w);

    // Four quantities, named separately, because every §9.1 mispricing
    // from M10k onward was one of them wearing another's number:
    // capacity is what the trace sees, fill is what the pages are,
    // handed is churn, and RSS is the only one a user experiences.
    try w.print("\n--- arena fill (M10q) ---\n", .{});
    inline for (std.meta.fields(fill.Owner), 0..) |f, i| {
        const t = tallies[i];
        try w.print(
            "arena={s} capacity_end={d} fill_peak={d} unfilled={d} handed={d} calls={d}\n",
            .{ f.name, t.capacity_end, t.peak, t.capacity_end -| t.peak, t.handed, t.calls },
        );
    }
    // Labelled, not reported plainly: in THIS mode the profiler's own
    // site table and live map are resident too, so the figure is tens of
    // times the production one. `fill` mode is where RSS means what a
    // user would see. Printing it anyway because a reader comparing the
    // two modes deserves to be told why they disagree.
    var b0: [24]u8 = undefined;
    var b1: [24]u8 = undefined;
    var b2: [24]u8 = undefined;
    try w.print(
        "rss_peak_bytes_PROFILER_INFLATED start={s} after_open={s} after_recalc={s}\n",
        .{ fmtRss(rss_before, &b0), fmtRss(rss_after_open, &b1), fmtRss(rss_after_recalc, &b2) },
    );
    return 0;
}

/// §9.1 M10q's four quantities, uncontaminated.
///
/// The same first recalc as `recalc` mode, with fill tallies installed
/// but **no heap profiler** — so `capacity`, `fill`, `unfilled` and RSS
/// are all measured on the process a user would run. `heap` mode cannot
/// do this: its own bookkeeping is resident and moves RSS by an order of
/// magnitude. A separate mode rather than a flag on `recalc`, because
/// `recalc` is the hyperfine timing lane and counting every arena
/// allocation would perturb the evaluate ceiling it exists to measure.
fn reportFill(io: std.Io, w: *std.Io.Writer, path: []const u8) !u8 {
    const gpa = std.heap.smp_allocator;

    const fill = pkg.fill_probe;
    var tallies: [std.meta.fields(fill.Owner).len]fill.Tally = @splat(.{});
    inline for (std.meta.fields(fill.Owner), 0..) |f, i| {
        _ = fill.install(@field(fill.Owner, f.name), &tallies[i]);
    }
    defer fill.clear();

    const rss_start = peakRssBytes();
    var wb = try Workbook.open(gpa, io, path);
    defer wb.deinit();
    const rss_open = peakRssBytes();

    var report = try wb.recalculate(gpa, io, RUN, .{});
    defer report.deinit(gpa);
    const rss_recalc = peakRssBytes();

    var cap_total: usize = 0;
    var fill_total: usize = 0;
    inline for (std.meta.fields(fill.Owner), 0..) |f, i| {
        const t = tallies[i];
        cap_total += t.capacity_end;
        fill_total += t.peak;
        try w.print(
            "arena={s} capacity_end={d} fill_peak={d} unfilled={d} handed={d} calls={d}\n",
            .{ f.name, t.capacity_end, t.peak, t.capacity_end -| t.peak, t.handed, t.calls },
        );
    }
    try w.print(
        "arena_totals capacity_end={d} fill_peak={d} unfilled={d}\n",
        .{ cap_total, fill_total, cap_total -| fill_total },
    );
    var b0: [24]u8 = undefined;
    var b1: [24]u8 = undefined;
    var b2: [24]u8 = undefined;
    try w.print(
        "rss_peak_bytes start={s} after_open={s} after_recalc={s}\n",
        .{ fmtRss(rss_start, &b0), fmtRss(rss_open, &b1), fmtRss(rss_recalc, &b2) },
    );
    return 0;
}

/// Peak RSS so far, in bytes, or null where the platform has no
/// `getrusage`. `maxrss` is monotonic, which is what makes it useful at a
/// checkpoint: it says *which phase* set the high-water mark, which is
/// the question the era trace answers in the wrong currency. macOS
/// reports bytes; Linux reports kilobytes.
///
/// **Null on Windows, not zero.** `std.posix.rusage` is `void` there, and
/// a zero would read as "no memory resident" in a table whose whole
/// purpose is comparing residency against capacity. §9's RSS lane is
/// POSIX-only anyway; this keeps the binary compiling for the
/// windows-runtime lane, which builds every bench exe.
fn peakRssBytes() ?u64 {
    const os = @import("builtin").os.tag;
    if (os == .windows) return null;
    const ru = std.posix.getrusage(std.posix.rusage.SELF);
    const raw: u64 = @intCast(ru.maxrss);
    return if (os == .macos) raw else raw * 1024;
}

/// Render an optional RSS reading. `unavailable` rather than a number the
/// reader would have to know to distrust.
fn fmtRss(v: ?u64, buf: []u8) []const u8 {
    const n = v orelse return "unavailable";
    return std.fmt.bufPrint(buf, "{d}", .{n}) catch "overflow";
}

/// **Current** resident set, not the high-water mark — the quantity
/// `getrusage` cannot report and every §9.1 row so far had to infer.
///
/// `maxrss` is monotonic, so it answers "how high did this process ever
/// go" and nothing else; an era is a *local* maximum, which needs a
/// curve that can come back down. macOS carries it in
/// `mach_task_basic_info`, Linux in field 2 of `/proc/self/statm`
/// (pages). Null elsewhere, so the windows-runtime lane still builds
/// this binary — a zero would read as "nothing resident" in a table
/// whose whole subject is residency.
fn residentBytes() ?u64 {
    const os = @import("builtin").os.tag;
    switch (os) {
        .macos => {
            var info: std.c.mach_task_basic_info = undefined;
            var count: std.c.mach_msg_type_number_t = std.c.MACH.TASK.BASIC.INFO_COUNT;
            const kr = std.c.task_info(
                std.c.mach_task_self(),
                std.c.MACH.TASK.BASIC.INFO,
                @ptrCast(&info),
                &count,
            );
            if (kr != 0) return null;
            return info.resident_size;
        },
        .linux => {
            // Raw syscalls: 0.16's `std.posix` has neither `open` nor
            // `close`, and the file layer wants an `std.Io` this
            // sampling thread does not have and must not borrow from
            // the process it is measuring.
            const linux = std.os.linux;
            const rc = linux.openat(linux.AT.FDCWD, "/proc/self/statm", .{ .ACCMODE = .RDONLY }, 0);
            if (linux.errno(rc) != .SUCCESS) return null;
            const fd: i32 = @intCast(rc);
            defer _ = linux.close(fd);
            var buf: [128]u8 = undefined;
            const rn = linux.read(fd, &buf, buf.len);
            if (linux.errno(rn) != .SUCCESS) return null;
            const n: usize = @intCast(rn);
            var it = std.mem.tokenizeScalar(u8, buf[0..n], ' ');
            _ = it.next() orelse return null;
            const rss_pages = it.next() orelse return null;
            const pages = std.fmt.parseInt(u64, rss_pages, 10) catch return null;
            return pages * std.heap.pageSize();
        },
        else => return null,
    }
}

/// Sum of every installed fill tally's high-water mark — "touched fill"
/// in one number, so a sample can carry it beside the resident reading.
///
/// Read off `peak` rather than `live` deliberately: `peak` only ever
/// rises, so a sampling thread reading it while the pipeline writes it
/// sees a value that was true at some instant at or before the sample,
/// never a torn one. `live` would need synchronisation the measured
/// process should not be paying for.
fn fillPeakTotal() u64 {
    const fill = pkg.fill_probe;
    var total: u64 = 0;
    inline for (std.meta.fields(fill.Owner)) |f| {
        if (fill.sinks[@intFromEnum(@field(fill.Owner, f.name))]) |t| total += t.peak;
    }
    return total;
}

/// §9.1c M10t — the resident-page curve, sampled off-thread.
///
/// Every §9.1 figure from M10k to M10s is denominated in one of two
/// currencies, and the ladder mispriced a row each time it read one as
/// the other: traced **live bytes** (the heap profiler wraps the
/// *backing* allocator, so for an arena it sees a chunk REQUEST) and
/// arena **fill** (what the arena hands out). Neither is a page. This
/// samples the third quantity directly, from the kernel, under the
/// production allocator — no profiler, nothing to inflate it.
///
/// **Why a sampling thread and not checkpoints in the pipeline.** An era
/// is a *local maximum of a curve*, and a curve needs samples between
/// the phase boundaries rather than at them. The public seams are four
/// (`open`, `prepare`, `serialize`, `swap`); the era vector has
/// twenty-four. Sampling finds maxima no seam brackets, and it needs no
/// production hook at all — which also means this mode measures exactly
/// the binary the gate measures.
///
/// The sampler's own buffer is allocated and touched *before* the
/// baseline reading, so every delta below is net of it.
const Sampler = struct {
    samples: []Sample,
    n: std.atomic.Value(usize) = .init(0),
    stop: std.atomic.Value(bool) = .init(false),
    phase: std.atomic.Value(u8) = .init(0),
    polls: std.atomic.Value(u64) = .init(0),
    dropped: std.atomic.Value(u64) = .init(0),

    const Sample = struct { rss: u64, fill: u64, phase: u8 };

    /// Phase labels. Coarse by construction — they name which public
    /// span a sample fell in, and the era structure comes from the
    /// curve itself, not from these.
    const phase_names = [_][]const u8{ "pre-open", "open", "recalc", "post" };

    /// Record only when the reading MOVES. RSS steps a page at a time
    /// and then sits flat for millions of instructions; storing every
    /// poll would need a buffer big enough to distort the thing being
    /// measured. A phase change always records, so every span has at
    /// least one sample.
    fn run(self: *Sampler) void {
        var last_rss: u64 = std.math.maxInt(u64);
        var last_phase: u8 = 255;
        while (!self.stop.load(.acquire)) {
            self.polls.store(self.polls.load(.monotonic) + 1, .monotonic);
            const r = residentBytes() orelse return;
            const ph = self.phase.load(.monotonic);
            if (r == last_rss and ph == last_phase) {
                std.atomic.spinLoopHint();
                continue;
            }
            last_rss = r;
            last_phase = ph;
            const i = self.n.load(.monotonic);
            if (i >= self.samples.len) {
                self.dropped.store(self.dropped.load(.monotonic) + 1, .monotonic);
                continue;
            }
            self.samples[i] = .{ .rss = r, .fill = fillPeakTotal(), .phase = ph };
            self.n.store(i + 1, .release);
        }
    }
};

fn reportPages(io: std.Io, w: *std.Io.Writer, path: []const u8) !u8 {
    const gpa = std.heap.smp_allocator;

    if (residentBytes() == null) {
        try w.writeAll("pages mode needs a resident-set probe; this platform has none\n");
        return 2;
    }

    // 32 768 change-points is ~10× what a 55 MiB climb can produce at
    // 16 KiB page granularity, and 786 KiB of buffer — allocated and
    // written here so the baseline below already contains it.
    const cap = 32 * 1024;
    const buf = try std.heap.page_allocator.alloc(Sampler.Sample, cap);
    defer std.heap.page_allocator.free(buf);
    @memset(buf, .{ .rss = 0, .fill = 0, .phase = 0 });

    const fill = pkg.fill_probe;
    var tallies: [std.meta.fields(fill.Owner).len]fill.Tally = @splat(.{});
    inline for (std.meta.fields(fill.Owner), 0..) |f, i| {
        _ = fill.install(@field(fill.Owner, f.name), &tallies[i]);
    }
    defer fill.clear();

    var sampler: Sampler = .{ .samples = buf };
    const baseline = residentBytes().?;

    const thread = try std.Thread.spawn(.{}, Sampler.run, .{&sampler});

    sampler.phase.store(1, .monotonic);
    var wb = try Workbook.open(gpa, io, path);
    sampler.phase.store(2, .monotonic);
    var report = try wb.recalculate(gpa, io, RUN, .{});
    sampler.phase.store(3, .monotonic);

    const rss_end = residentBytes().?;
    sampler.stop.store(true, .release);
    thread.join();

    report.deinit(gpa);
    wb.deinit();

    const n = sampler.n.load(.acquire);
    const taken = buf[0..n];

    try w.print(
        "baseline_rss={d} end_rss={d} samples={d} polls={d} dropped={d} sampler_buffer_bytes={d}\n",
        .{ baseline, rss_end, n, sampler.polls.load(.monotonic), sampler.dropped.load(.monotonic), cap * @sizeOf(Sampler.Sample) },
    );

    var peak: u64 = 0;
    for (taken) |s| peak = @max(peak, s.rss);
    const maxrss = peakRssBytes() orelse 0;
    // The sampler's validity check, and it is not optional: if the thread
    // missed the instant the kernel's own monotonic counter caught, every
    // era height below is a floor rather than a height. Reported as a
    // ratio so a reader can see how close, not just whether.
    try w.print(
        "sampled_peak={d} getrusage_maxrss={d} coverage={d:.4}\n",
        .{ peak, maxrss, if (maxrss == 0) 0.0 else @as(f64, @floatFromInt(peak)) / @as(f64, @floatFromInt(maxrss)) },
    );

    try reportResidentEras(w, taken, baseline);

    try w.writeAll("\n--- arena fill at exit ---\n");
    var cap_total: usize = 0;
    var fill_total: usize = 0;
    inline for (std.meta.fields(fill.Owner), 0..) |f, i| {
        const t = tallies[i];
        cap_total += t.capacity_end;
        fill_total += t.peak;
        try w.print(
            "arena={s} capacity_end={d} fill_peak={d} unfilled={d} handed={d} calls={d}\n",
            .{ f.name, t.capacity_end, t.peak, t.capacity_end -| t.peak, t.handed, t.calls },
        );
    }
    try w.print(
        "arena_totals capacity_end={d} fill_peak={d} unfilled={d}\n",
        .{ cap_total, fill_total, cap_total -| fill_total },
    );
    return 0;
}

/// The same era rule the heap profiler applies to the live curve
/// (`era_dip` below a local maximum ends an era), applied to the
/// **resident** curve instead. Same threshold on purpose: the two
/// vectors are meant to be read side by side, and a different dip would
/// make a difference in era count an artefact of the instrument.
fn reportResidentEras(w: *std.Io.Writer, taken: []const Sampler.Sample, baseline: u64) !void {
    const era_dip = HeapProfiler.era_dip;

    try w.writeAll("\n--- resident eras (M10t) ---\n");
    try w.writeAll("era phase resident_bytes adjusted_bytes fill_at_era unfilled_vs_resident\n");

    var local_max: u64 = 0;
    var max_fill: u64 = 0;
    var max_phase: u8 = 0;
    var recorded = true;
    var era: usize = 0;

    for (taken) |s| {
        if (s.rss > local_max) {
            local_max = s.rss;
            max_fill = s.fill;
            max_phase = s.phase;
            recorded = false;
        } else if (!recorded and local_max - s.rss >= era_dip) {
            try printEra(w, era, max_phase, local_max, baseline, max_fill);
            era += 1;
            recorded = true;
            local_max = s.rss;
            max_fill = s.fill;
            max_phase = s.phase;
        }
    }
    if (!recorded) try printEra(w, era, max_phase, local_max, baseline, max_fill);
}

fn printEra(
    w: *std.Io.Writer,
    era: usize,
    phase: u8,
    resident: u64,
    baseline: u64,
    fill: u64,
) !void {
    const adjusted = resident -| baseline;
    try w.print(
        "era={d} phase={s} resident={d} adjusted={d} ({d:.2} MiB) fill={d} resident_minus_fill={d}\n",
        .{
            era,
            Sampler.phase_names[@min(phase, Sampler.phase_names.len - 1)],
            resident,
            adjusted,
            @as(f64, @floatFromInt(adjusted)) / (1024.0 * 1024.0),
            fill,
            @as(i64, @intCast(adjusted)) - @as(i64, @intCast(fill)),
        },
    );
}

/// Attributing allocator: wraps a backing allocator, keys every
/// allocation to a call-site stack, and snapshots each site's live
/// bytes whenever the total live footprint sets a new high-water mark.
/// Bookkeeping lives in `page_allocator` so the profiler never recurses
/// into the allocator it is measuring.
const HeapProfiler = struct {
    backing: std.mem.Allocator,
    lock: std.atomic.Value(u8) = .init(0),
    live: std.AutoHashMapUnmanaged(usize, LiveEntry) = .empty,
    site_index: std.AutoHashMapUnmanaged(u64, u32) = .empty,
    sites: std.ArrayList(Site) = .empty,
    live_total: usize = 0,
    peak_total: usize = 0,
    last_snapshot: usize = 0,
    /// Chronological local maxima of the live curve (M10k). The peak
    /// table names one instant; §9.1 M10j's lesson is that cutting it
    /// moves RSS only down to the *runner-up* instant, so the probe has
    /// to name every era's height, not just the winner's. A climb's
    /// high-water mark is recorded once the curve falls `era_dip` below
    /// it, then the tracker re-arms at the current level so a later,
    /// lower era still registers.
    era_local_max: usize = 0,
    era_recorded: bool = true,
    era_last_snapshot: usize = 0,
    eras: std.ArrayList(EraRecord) = .empty,

    /// Deep enough to see past the parser's recursive-descent chain
    /// (eleven frames of parseX before the caller appears); shallow
    /// enough that a million-event run stays seconds, not minutes.
    const max_depth = 32;
    /// Re-snapshot per-site live bytes when the peak has grown this
    /// much since the last snapshot. 1 MiB of drift against a >500 MiB
    /// peak bounds the attribution error at ~0.2 %.
    const snapshot_step = 1 * 1024 * 1024;
    /// A descent this deep ends an era. Deep enough that a scratch
    /// arena's per-formula churn cannot split an era in two; shallow
    /// enough that the pipeline's real phase frees all register.
    const era_dip = 2 * 1024 * 1024;

    const LiveEntry = struct { site: u32, size: usize };

    /// One era: its height and the sites that held it, snapshotted with
    /// the same ≤`snapshot_step` lag the global peak table carries.
    const EraRecord = struct {
        height: usize,
        top: [12]EraSite,
        n: u8,
    };
    const EraSite = struct { site: u32, bytes: usize };

    const Site = struct {
        frames: [max_depth]usize,
        depth: u8,
        live: usize = 0,
        at_peak: usize = 0,
        at_era: usize = 0,
        allocs: u64 = 0,
        total: u64 = 0,
    };

    const meta = std.heap.page_allocator;

    fn init(backing: std.mem.Allocator) HeapProfiler {
        return .{ .backing = backing };
    }

    fn allocator(self: *HeapProfiler) std.mem.Allocator {
        return .{ .ptr = self, .vtable = &vtable };
    }

    const vtable: std.mem.Allocator.VTable = .{
        .alloc = allocFn,
        .resize = resizeFn,
        .remap = remapFn,
        .free = freeFn,
    };

    fn acquire(self: *HeapProfiler) void {
        while (self.lock.cmpxchgWeak(0, 1, .acquire, .monotonic) != null) {
            std.atomic.spinLoopHint();
        }
    }

    fn release(self: *HeapProfiler) void {
        self.lock.store(0, .release);
    }

    /// The stack is captured with `first_address = ret_addr`, so frame 0
    /// is the caller of the allocator itself — `HashMap.grow`,
    /// `ArrayList.ensureTotalCapacity` — and the frames above it are
    /// what make the site nameable.
    fn siteOf(self: *HeapProfiler, ret_addr: usize) u32 {
        var buf: [max_depth]usize = undefined;
        const st = std.debug.captureCurrentStackTrace(.{ .first_address = ret_addr }, &buf);
        var frames = st.return_addresses;
        if (frames.len == 0) {
            // Unwind failed to reach the caller (tail call, missing
            // frame): key on the raw return address rather than losing
            // the event.
            buf[0] = ret_addr;
            frames = buf[0..1];
        }
        const key = std.hash.Wyhash.hash(0, std.mem.sliceAsBytes(frames));
        const gop = self.site_index.getOrPut(meta, key) catch @panic("heap profiler OOM");
        if (!gop.found_existing) {
            var site: Site = .{ .frames = undefined, .depth = @intCast(frames.len) };
            @memcpy(site.frames[0..frames.len], frames);
            self.sites.append(meta, site) catch @panic("heap profiler OOM");
            gop.value_ptr.* = @intCast(self.sites.items.len - 1);
        }
        return gop.value_ptr.*;
    }

    fn noteAlloc(self: *HeapProfiler, ptr: usize, len: usize, ret_addr: usize) void {
        const site_id = self.siteOf(ret_addr);
        self.live.put(meta, ptr, .{ .site = site_id, .size = len }) catch @panic("heap profiler OOM");
        const site = &self.sites.items[site_id];
        site.live += len;
        site.allocs += 1;
        site.total += len;
        self.live_total += len;
        self.maybeSnapshot();
    }

    fn noteResize(self: *HeapProfiler, ptr: usize, new_ptr: usize, new_len: usize) void {
        const entry = self.live.getPtr(ptr) orelse return;
        const site = &self.sites.items[entry.site];
        site.live = site.live - entry.size + new_len;
        if (new_len > entry.size) site.total += new_len - entry.size;
        self.live_total = self.live_total - entry.size + new_len;
        if (new_ptr == ptr) {
            entry.size = new_len;
        } else {
            const moved: LiveEntry = .{ .site = entry.site, .size = new_len };
            _ = self.live.remove(ptr);
            self.live.put(meta, new_ptr, moved) catch @panic("heap profiler OOM");
        }
        self.maybeSnapshot();
    }

    fn noteFree(self: *HeapProfiler, ptr: usize) void {
        const kv = self.live.fetchRemove(ptr) orelse return;
        self.sites.items[kv.value.site].live -= kv.value.size;
        self.live_total -= kv.value.size;
        self.noteEra();
    }

    fn maybeSnapshot(self: *HeapProfiler) void {
        self.noteEra();
        if (self.live_total <= self.peak_total) return;
        self.peak_total = self.live_total;
        if (self.peak_total - self.last_snapshot < snapshot_step) return;
        self.last_snapshot = self.peak_total;
        for (self.sites.items) |*site| site.at_peak = site.live;
    }

    fn noteEra(self: *HeapProfiler) void {
        if (self.live_total > self.era_local_max) {
            self.era_local_max = self.live_total;
            self.era_recorded = false;
            if (self.era_local_max - self.era_last_snapshot >= snapshot_step) {
                self.era_last_snapshot = self.era_local_max;
                for (self.sites.items) |*site| site.at_era = site.live;
            }
        } else if (!self.era_recorded and
            self.era_local_max - self.live_total >= era_dip)
        {
            self.eras.append(meta, self.eraRecord()) catch @panic("heap profiler OOM");
            self.era_recorded = true;
            self.era_local_max = self.live_total;
            self.era_last_snapshot = self.live_total;
        }
    }

    /// The era's top sites by their lagging snapshot, by selection into
    /// the record's fixed-width table.
    fn eraRecord(self: *HeapProfiler) EraRecord {
        var rec: EraRecord = .{ .height = self.era_local_max, .top = undefined, .n = 0 };
        for (self.sites.items, 0..) |site, id| {
            if (site.at_era == 0) continue;
            var candidate: EraSite = .{ .site = @intCast(id), .bytes = site.at_era };
            var i: usize = 0;
            while (i < rec.n) : (i += 1) {
                if (candidate.bytes > rec.top[i].bytes) {
                    std.mem.swap(EraSite, &candidate, &rec.top[i]);
                }
            }
            if (rec.n < rec.top.len) {
                rec.top[rec.n] = candidate;
                rec.n += 1;
            }
        }
        return rec;
    }

    fn allocFn(ctx: *anyopaque, len: usize, alignment: std.mem.Alignment, ret_addr: usize) ?[*]u8 {
        const self: *HeapProfiler = @ptrCast(@alignCast(ctx));
        const p = self.backing.rawAlloc(len, alignment, ret_addr) orelse return null;
        self.acquire();
        defer self.release();
        self.noteAlloc(@intFromPtr(p), len, ret_addr);
        return p;
    }

    fn resizeFn(ctx: *anyopaque, memory: []u8, alignment: std.mem.Alignment, new_len: usize, ret_addr: usize) bool {
        const self: *HeapProfiler = @ptrCast(@alignCast(ctx));
        if (!self.backing.rawResize(memory, alignment, new_len, ret_addr)) return false;
        self.acquire();
        defer self.release();
        self.noteResize(@intFromPtr(memory.ptr), @intFromPtr(memory.ptr), new_len);
        return true;
    }

    fn remapFn(ctx: *anyopaque, memory: []u8, alignment: std.mem.Alignment, new_len: usize, ret_addr: usize) ?[*]u8 {
        const self: *HeapProfiler = @ptrCast(@alignCast(ctx));
        const p = self.backing.rawRemap(memory, alignment, new_len, ret_addr) orelse return null;
        self.acquire();
        defer self.release();
        self.noteResize(@intFromPtr(memory.ptr), @intFromPtr(p), new_len);
        return p;
    }

    fn freeFn(ctx: *anyopaque, memory: []u8, alignment: std.mem.Alignment, ret_addr: usize) void {
        const self: *HeapProfiler = @ptrCast(@alignCast(ctx));
        self.acquire();
        self.noteFree(@intFromPtr(memory.ptr));
        self.release();
        self.backing.rawFree(memory, alignment, ret_addr);
    }

    /// Top sites by live-bytes-at-peak, symbolized in-process. Stacks
    /// go to stderr (`dumpStackTrace`'s stream); the table goes to
    /// stdout with a `site=` line each so the two can be joined.
    fn report(self: *HeapProfiler, w: *std.Io.Writer) !void {
        try w.print(
            "peak_live_bytes={d} ({d:.1} MiB) live_end_bytes={d} sites={d}\n",
            .{
                self.peak_total,
                @as(f64, @floatFromInt(self.peak_total)) / (1024.0 * 1024.0),
                self.live_total,
                self.sites.items.len,
            },
        );
        for (self.eras.items, 0..) |rec, i| {
            try w.print(
                "era={d} peak_bytes={d} ({d:.1} MiB) top=",
                .{ i, rec.height, @as(f64, @floatFromInt(rec.height)) / (1024.0 * 1024.0) },
            );
            for (rec.top[0..rec.n], 0..) |s, j| {
                try w.print("{s}id{d}:{d}", .{ if (j == 0) "" else ",", s.site, s.bytes });
            }
            try w.writeAll("\n");
        }
        if (!self.era_recorded) {
            try w.print(
                "era={d} peak_bytes={d} ({d:.1} MiB) (open at exit)\n",
                .{
                    self.eras.items.len,
                    self.era_local_max,
                    @as(f64, @floatFromInt(self.era_local_max)) / (1024.0 * 1024.0),
                },
            );
        }

        const order = meta.alloc(u32, self.sites.items.len) catch @panic("heap profiler OOM");
        defer meta.free(order);
        for (order, 0..) |*slot, idx| slot.* = @intCast(idx);
        std.mem.sort(u32, order, self.sites.items, siteGreater);

        const top = @min(order.len, 25);
        var printed: std.ArrayList(u32) = .empty;
        defer printed.deinit(meta);
        for (order[0..top], 0..) |site_id, rank| {
            const site = self.sites.items[site_id];
            if (site.at_peak == 0) break;
            printed.append(meta, site_id) catch @panic("heap profiler OOM");
            try w.print(
                "site={d} id={d} at_peak_bytes={d} ({d:.1} MiB, {d:.1}%) allocs={d} churn_bytes={d}\n",
                .{
                    rank,
                    site_id,
                    site.at_peak,
                    @as(f64, @floatFromInt(site.at_peak)) / (1024.0 * 1024.0),
                    @as(f64, @floatFromInt(site.at_peak)) * 100.0 / @as(f64, @floatFromInt(self.peak_total)),
                    site.allocs,
                    site.total,
                },
            );
            try w.flush();
            std.debug.print("--- site {d} ---\n", .{rank});
            var frames_buf: [max_depth]usize = undefined;
            @memcpy(frames_buf[0..site.depth], site.frames[0..site.depth]);
            const st: std.debug.StackTrace = .{
                .return_addresses = frames_buf[0..site.depth],
                .skipped = .none,
            };
            std.debug.dumpStackTrace(&st);
        }
        self.dumpEraOnlySites(printed.items);
    }

    fn siteGreater(sites: []const Site, a: u32, b: u32) bool {
        return sites[a].at_peak > sites[b].at_peak;
    }

    /// Stacks for sites the era tables name but the global top-25 does
    /// not — without these an `idN` in an era row is a number with no
    /// name.
    fn dumpEraOnlySites(self: *HeapProfiler, printed: []const u32) void {
        var dumped: std.ArrayList(u32) = .empty;
        defer dumped.deinit(meta);
        for (self.eras.items) |rec| {
            for (rec.top[0..rec.n]) |s| {
                const already = for (printed) |p| {
                    if (p == s.site) break true;
                } else for (dumped.items) |p| {
                    if (p == s.site) break true;
                } else false;
                if (already) continue;
                dumped.append(meta, s.site) catch @panic("heap profiler OOM");
                const site = self.sites.items[s.site];
                std.debug.print("--- id {d} (era-only) ---\n", .{s.site});
                var frames_buf: [max_depth]usize = undefined;
                @memcpy(frames_buf[0..site.depth], site.frames[0..site.depth]);
                const st: std.debug.StackTrace = .{
                    .return_addresses = frames_buf[0..site.depth],
                    .skipped = .none,
                };
                std.debug.dumpStackTrace(&st);
            }
        }
    }
};

fn emitFixture(
    gpa: std.mem.Allocator,
    io: std.Io,
    w: *std.Io.Writer,
    path: []const u8,
    geometry: synth.Geometry,
    named_size: bool,
) !u8 {
    const bytes = try synth.bytes(gpa, io, geometry);
    defer gpa.free(bytes);

    var hex: [synth.digest_len * 2]u8 = undefined;
    const digest = synth.digestHex(bytes, &hex);

    try std.Io.Dir.cwd().writeFile(io, .{ .sub_path = path, .data = bytes });
    try w.print(
        "fixture={s} rows={d} cells={d} formula_cells={d} archive_bytes={d} sha256={s}\n",
        .{ path, geometry.data_rows, geometry.cells(), geometry.formulaCells(), bytes.len, digest },
    );

    // The gate §9 asks for: the named workload is specific bytes, and a
    // baseline measured against different ones describes a different
    // workbook. Re-record both together or neither.
    if (named_size and !std.mem.eql(u8, digest, synth.named_digest_sha256)) {
        try w.print(
            "FAIL: named workload digest drifted\n  recorded {s}\n  emitted  {s}\n" ++
                "  The fixture IS the baseline's identity: re-measure §9's numbers and\n" ++
                "  update `named_digest_sha256` in the same commit.\n",
            .{ synth.named_digest_sha256, digest },
        );
        return 1;
    }
    return 0;
}

/// The criteria workload's `emit`, under the same digest discipline:
/// `small` is its identity size, and a baseline measured against
/// different bytes describes a different workbook.
fn emitCriteria(
    gpa: std.mem.Allocator,
    io: std.Io,
    w: *std.Io.Writer,
    path: []const u8,
    geometry: crit.Geometry,
    identity_size: bool,
) !u8 {
    const bytes = try crit.bytes(gpa, io, geometry);
    defer gpa.free(bytes);

    var hex: [crit.digest_len * 2]u8 = undefined;
    const digest = crit.digestHex(bytes, &hex);

    try std.Io.Dir.cwd().writeFile(io, .{ .sub_path = path, .data = bytes });
    try w.print(
        "fixture={s} rows={d} cells={d} formula_cells={d} archive_bytes={d} sha256={s}\n",
        .{ path, geometry.data_rows, geometry.cells(), geometry.formulaCells(), bytes.len, digest },
    );

    if (identity_size and !std.mem.eql(u8, digest, crit.small_digest_sha256)) {
        try w.print(
            "FAIL: criteria workload digest drifted\n  recorded {s}\n  emitted  {s}\n" ++
                "  The fixture IS the baseline's identity: re-measure the criteria\n" ++
                "  numbers and update `small_digest_sha256` in the same commit.\n",
            .{ crit.small_digest_sha256, digest },
        );
        return 1;
    }
    return 0;
}

/// The TEXT workload's `emit`, under the same digest discipline:
/// `small` is its identity size, and a baseline measured against
/// different bytes describes a different workbook.
fn emitText(
    gpa: std.mem.Allocator,
    io: std.Io,
    w: *std.Io.Writer,
    path: []const u8,
    geometry: text.Geometry,
    identity_size: bool,
) !u8 {
    const bytes = try text.bytes(gpa, io, geometry);
    defer gpa.free(bytes);

    var hex: [text.digest_len * 2]u8 = undefined;
    const digest = text.digestHex(bytes, &hex);

    try std.Io.Dir.cwd().writeFile(io, .{ .sub_path = path, .data = bytes });
    try w.print(
        "fixture={s} rows={d} cells={d} formula_cells={d} archive_bytes={d} sha256={s}\n",
        .{ path, geometry.data_rows, geometry.cells(), geometry.formulaCells(), bytes.len, digest },
    );

    if (identity_size and !std.mem.eql(u8, digest, text.small_digest_sha256)) {
        try w.print(
            "FAIL: text workload digest drifted\n  recorded {s}\n  emitted  {s}\n" ++
                "  The fixture IS the baseline's identity: re-measure the TEXT\n" ++
                "  numbers and update `small_digest_sha256` in the same commit.\n",
            .{ text.small_digest_sha256, digest },
        );
        return 1;
    }
    return 0;
}

/// The registry workload's `emit`, under the same digest discipline:
/// `small` is its identity size, and a baseline measured against
/// different bytes describes a different workbook.
fn emitRegistry(
    gpa: std.mem.Allocator,
    io: std.Io,
    w: *std.Io.Writer,
    path: []const u8,
    geometry: registry.Geometry,
    identity_size: bool,
) !u8 {
    const bytes = try registry.bytes(gpa, io, geometry);
    defer gpa.free(bytes);

    var hex: [registry.digest_len * 2]u8 = undefined;
    const digest = registry.digestHex(bytes, &hex);

    try std.Io.Dir.cwd().writeFile(io, .{ .sub_path = path, .data = bytes });
    try w.print(
        "fixture={s} rows={d} cells={d} formula_cells={d} archive_bytes={d} sha256={s}\n",
        .{ path, geometry.data_rows, geometry.cells(), geometry.formulaCells(), bytes.len, digest },
    );

    if (identity_size and !std.mem.eql(u8, digest, registry.small_digest_sha256)) {
        try w.print(
            "FAIL: registry workload digest drifted\n  recorded {s}\n  emitted  {s}\n" ++
                "  The fixture IS the baseline's identity: re-measure the registry\n" ++
                "  numbers and update `small_digest_sha256` in the same commit.\n",
            .{ registry.small_digest_sha256, digest },
        );
        return 1;
    }
    return 0;
}

/// §9's "measured phases reported separately", as far as the public
/// surface reaches — and every row is a directly measured span rather
/// than a difference of two runs, which at the named size would subtract
/// one ~2-minute measurement from another.
///
/// That works because `recalc_run.prepare` is public (M5d2 exported it
/// for exactly this kind of caller) and hands back the candidate
/// *unswapped*, which is the state §5.7.9 serializes from. So one
/// process can time the open, the prepare, the serialize+commit off that
/// candidate, and the swap, in the order the transaction performs them.
fn reportPhases(
    gpa: std.mem.Allocator,
    io: std.Io,
    w: *std.Io.Writer,
    path: []const u8,
    out_path: ?[]const u8,
) !u8 {
    const dest = out_path orelse {
        try w.writeAll("phases mode needs --out PATH\n");
        return 2;
    };

    var t = Stopwatch.start(io);
    var wb = try Workbook.open(gpa, io, path);
    const open_ns = t.lap();

    var prepared = try pkg.recalc_run.prepare(&wb, gpa, io, RUN, .{});
    const prepare_ns = t.lap();

    var serialize_ns: i128 = 0;
    var swap_ns: i128 = 0;
    switch (prepared) {
        .ok => |*candidate| {
            // §5.7.9 serializes from the prepared, *unswapped* candidate,
            // so this is the same store and the same bytes
            // `saveWithRecalc` writes. Only the swap's position differs —
            // it runs inside the commit region there and after the save
            // here — and where the swap happens does not change what
            // either span costs.
            try candidate.next.save(io, dest);
            serialize_ns = t.lap();

            candidate.swap(&wb);
            swap_ns = t.lap();
            var report = candidate.takeReport();
            report.deinit(gpa);
        },
        .none => |*r| r.deinit(gpa),
        .refused => {
            wb.deinit();
            try w.writeAll("FAIL: the fixture refused; a bench cannot measure a refusal\n");
            return 1;
        },
    }
    wb.deinit();

    // Every row below is a span this process measured. An earlier
    // revision derived serialize+commit as `saveWithRecalc − prepare`
    // across two independent runs, which at the named size is a
    // difference of two ~2-minute measurements — comfortably able to
    // come out negative on a thermally-drifting laptop, and reported as
    // a phase cost.
    try w.print("| Phase | ms |\n|---|---:|\n", .{});
    try w.print("| open (archive + structural parts) | {d:.2} |\n", .{ms(open_ns)});
    try w.print("| prepare (model + graph + evaluate + stage + txn) | {d:.2} |\n", .{ms(prepare_ns)});
    try w.print("| serialize + commit | {d:.2} |\n", .{ms(serialize_ns)});
    try w.print("| swap | {d:.3} |\n", .{ms(swap_ns)});
    try w.print(
        "| end-to-end (sum) | {d:.2} |\n",
        .{ms(open_ns + prepare_ns + serialize_ns + swap_ns)},
    );
    return 0;
}

/// `std.Io`'s monotonic clock — the same one §5.5's deadlines are read
/// against. 0.16 removed `std.time.nanoTimestamp`, and reaching for a
/// second clock would make the numbers here incomparable with the ones
/// the polling seam produces.
const Stopwatch = struct {
    io: std.Io,
    last: i128,

    fn start(io: std.Io) Stopwatch {
        return .{ .io = io, .last = std.Io.Timestamp.now(io, .awake).nanoseconds };
    }

    fn lap(self: *Stopwatch) i128 {
        const now = std.Io.Timestamp.now(self.io, .awake).nanoseconds;
        const delta = now - self.last;
        self.last = now;
        return delta;
    }
};

fn ms(ns: i128) f64 {
    return @as(f64, @floatFromInt(ns)) / std.time.ns_per_ms;
}
