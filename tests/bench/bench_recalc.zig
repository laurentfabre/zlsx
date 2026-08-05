//! §9's recalc bench harness (M5d3): the baseline v1's absolute
//! ceilings are measured against.
//!
//! One binary, five modes, because hyperfine measures whole processes
//! and §9 wants two different numbers out of the same workload:
//!
//!   emit    write the fixture and print its SHA-256 (and, for the
//!           named geometry, refuse if it is not the recorded one)
//!   open    `Workbook.open` and stop — the fixed cost
//!   recalc  open + `Workbook.recalculate` — §9's *evaluate* lane
//!   save    open + `Workbook.saveWithRecalc` — §9's *end-to-end* lane
//!   phases  one instrumented run, phases reported separately
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

const Mode = enum { emit, open, recalc, save, phases };

pub fn main(init: std.process.Init) !u8 {
    const io = init.io;
    const gpa = std.heap.smp_allocator;
    const args = try init.minimal.args.toSlice(init.arena.allocator());

    var out_file: std.Io.File.Writer = .init(std.Io.File.stdout(), io, &.{});
    const w = &out_file.interface;
    defer w.flush() catch {};

    if (args.len < 3) {
        try w.print(
            "usage: {s} <emit|open|recalc|save|phases> <fixture.xlsx>" ++
                " [--size named|small|tiny] [--rows N] [--out PATH]\n",
            .{args[0]},
        );
        return 2;
    }

    const mode = std.meta.stringToEnum(Mode, args[1]) orelse {
        try w.print("unknown mode: {s}\n", .{args[1]});
        return 2;
    };
    const path = args[2];

    var geometry = synth.named;
    var named_size = true;
    var out_path: ?[]const u8 = null;
    var i: usize = 3;
    while (i < args.len) : (i += 1) {
        if (std.mem.eql(u8, args[i], "--size") and i + 1 < args.len) {
            i += 1;
            if (std.mem.eql(u8, args[i], "small")) {
                geometry = synth.small;
                named_size = false;
            } else if (std.mem.eql(u8, args[i], "tiny")) {
                geometry = synth.tiny;
                named_size = false;
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
        } else if (std.mem.eql(u8, args[i], "--out") and i + 1 < args.len) {
            i += 1;
            out_path = args[i];
        }
    }

    switch (mode) {
        .emit => return emitFixture(gpa, io, w, path, geometry, named_size),
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
    }
    return 0;
}

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
