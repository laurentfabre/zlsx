//! §12.2 (M6): `zlsx eval` and `zlsx recalc` — the formula engine on the
//! command line.
//!
//! Like `dbx`, this family has its own argument grammar and delegates the
//! whole tail before `parseArgs`, so the shipped read/edit/embed commands
//! and their NDJSON row envelope are untouched by construction.
//!
//! Three contracts live here and nowhere else:
//!
//!  1. **The stream state machine.** Every record carries `"kind"` and
//!     `"v":1`. The grammars are normative (docs/cli.md) — refusal and
//!     cancellation can occur BEFORE any header exists:
//!
//!         eval   := refusal | cancelled
//!                 | eval-header (eval-cell*) diagnostic*
//!                   (eval-complete | refusal | cancelled)
//!         recalc := diagnostic* (recalc-report | refusal | cancelled)
//!                   (stdout records only with --report)
//!
//!     SIGPIPE is the one exception: a closed pipe cannot receive a
//!     terminal record, so the stream ends prefix-valid with no terminal
//!     and exit 0. Any OTHER EOF without a terminal exits nonzero, which
//!     is what makes the two distinguishable.
//!
//!  2. **The exit table** (§12.2): 0 success (Excel error values ARE
//!     results), 1 usage, 2 open/parse, 3 typed refusal / `--deadline`,
//!     4 genuine OOM, 5 output write failure, 6 default-context
//!     acquisition failure, 130/143 SIGINT/SIGTERM.
//!
//!  3. **Commit-aware exit mapping.** A signal that arrives after the
//!     rename reports success (0), never 130/143 — the rename is the
//!     commit, and the destination already holds the recalced bytes.
//!     `mapRecalcSignalExit` states the rule; `cli.zig`'s
//!     `signals.exitCode` defers to it via the `exit_is_final` latch
//!     instead of overriding at exit.
//!
//! Everything the process environment supplies is injected through
//! `Deps`, so every row of the exit table — including the two that need
//! a broken clock and a signal at the commit seam — is drivable from a
//! unit test without timing.

const std = @import("std");
const builtin = @import("builtin");
const zlsx_pkg = @import("zlsx_pkg");
const engine = @import("zlsx_formula");
const coords = @import("zlsx_refs");
const zlsx_control = @import("zlsx_control");
const Allocator = std.mem.Allocator;
const assert = std.debug.assert;

pub const stream_version = 1;

// ─── the exit table (§12.2) ──────────────────────────────────────

pub const exit_ok: u8 = 0;
pub const exit_usage: u8 = 1;
pub const exit_open: u8 = 2;
pub const exit_refusal: u8 = 3;
pub const exit_oom: u8 = 4;
pub const exit_write: u8 = 5;
/// The `std.Io` wall clock or random source could not be read while
/// resolving an omitted `--now` / `--seed`. Never conflated with OOM:
/// a run that cannot know what time it is has not run out of memory.
pub const exit_context: u8 = 6;
pub const exit_sigint: u8 = 130;
pub const exit_sigterm: u8 = 143;

// ─── injected process environment ────────────────────────────────

pub const ContextError = error{ContextUnavailable};

/// §5.5's default-context acquisition, as two injectable reads. The
/// library refuses to default `now`/`seed` (equal inputs ⇒ equal output
/// would silently break); the CLI is the layer that IS allowed to read
/// the environment, and this is the only place it does.
pub const ContextSource = struct {
    now_ms: *const fn (io: std.Io) ContextError!i64 = productionNowMs,
    seed: *const fn (io: std.Io) ContextError!u64 = productionSeed,
};

fn productionNowMs(io: std.Io) ContextError!i64 {
    // `Clock.now` itself cannot fail; an unsupported clock reveals
    // itself through a zero (or unqueryable) resolution, so that is
    // the acquisition check.
    const res = std.Io.Clock.resolution(.real, io) catch return error.ContextUnavailable;
    if (res.nanoseconds == 0) return error.ContextUnavailable;
    const ts = std.Io.Timestamp.now(io, .real);
    return @intCast(@divFloor(ts.nanoseconds, std.time.ns_per_ms));
}

fn productionSeed(io: std.Io) ContextError!u64 {
    var bytes: [8]u8 = undefined;
    // `randomSecure`, not `random`: the fallback path of the latter is
    // exactly the silent degradation exit 6 exists to refuse.
    io.randomSecure(&bytes) catch return error.ContextUnavailable;
    return std.mem.readInt(u64, &bytes, .little);
}

/// The CLI's signal flags, as this module reads them. Pointers rather
/// than values because the handlers in `cli.zig` keep writing while we
/// run; null pointers (the test default) read as "never fired".
pub const SignalState = struct {
    stop: ?*const std.atomic.Value(bool) = null,
    sigint: ?*const std.atomic.Value(bool) = null,
    sigterm: ?*const std.atomic.Value(bool) = null,
    sigpipe: ?*const std.atomic.Value(bool) = null,

    fn flag(p: ?*const std.atomic.Value(bool)) bool {
        return if (p) |a| a.load(.acquire) else false;
    }

    pub fn stopped(self: SignalState) bool {
        return flag(self.stop);
    }
    pub fn intFired(self: SignalState) bool {
        return flag(self.sigint);
    }
    pub fn termFired(self: SignalState) bool {
        return flag(self.sigterm);
    }
    pub fn pipeFired(self: SignalState) bool {
        return flag(self.sigpipe);
    }

    /// The cancel token the recalc transaction polls — the same flag
    /// the signal handlers set, so the §5.5 polling bound the archive
    /// layer honours is what makes Ctrl-C prompt.
    fn cancelToken(self: SignalState) ?zlsx_control.CancelToken {
        const p = self.stop orelse return null;
        return .{ .atomic = p };
    }
};

pub const Deps = struct {
    sig: SignalState = .{},
    ctx: ContextSource = .{},
    /// Test seam: called before each stdout record with how many records
    /// are already out. Production leaves it null. This is what makes
    /// "SIGPIPE after the second record" a statement a test can inject
    /// rather than a race it has to win.
    before_record: ?*const fn (state: ?*anyopaque, emitted: usize) void = null,
    before_record_state: ?*anyopaque = null,
};

// ─── §12.2's commit-aware exit mapping ───────────────────────────

/// The rule `cli.zig:1718-1729` used to override: map a run's outcome
/// and the signal flags to the shell status byte. `committed` means the
/// rename happened — §5.7.9's commit — and from that point the run IS a
/// success no matter what arrives afterwards; reporting 130/143 for a
/// file that durably holds the recalced bytes would tell a scripted
/// caller to retry work that finished.
pub fn mapRecalcSignalExit(committed: bool, sig: SignalState) u8 {
    if (committed) return exit_ok;
    if (sig.intFired()) return exit_sigint;
    if (sig.termFired()) return exit_sigterm;
    // Neither signal: the cancellation came from `--deadline`, which is
    // a refusal of the run, not a kill.
    return exit_refusal;
}

// ─── entry point ─────────────────────────────────────────────────

/// `argv` starts at the sub-command name (`eval` or `recalc`).
///
/// Never returns an error: every failure is mapped to the exit table
/// here, where the stream state is known — a propagated error would
/// reach `main`'s generic classifier, which cannot know whether a
/// terminal record was already written or a rename already happened.
pub fn run(
    gpa: Allocator,
    io: std.Io,
    argv: []const []const u8,
    out: *std.Io.Writer,
    err_w: *std.Io.Writer,
    deps: Deps,
) u8 {
    assert(argv.len >= 1);
    const code = if (std.mem.eql(u8, argv[0], "eval"))
        runEval(gpa, io, argv[1..], out, err_w, deps)
    else if (std.mem.eql(u8, argv[0], "recalc"))
        runRecalc(gpa, io, argv[1..], out, err_w, deps)
    else
        unreachable; // dispatch in cli.zig only routes these two names
    err_w.flush() catch {};
    return code;
}

// ─── argument grammars ───────────────────────────────────────────

const UsageError = error{Usage};

/// One token-taking step shared by both parsers: the token after a
/// value-bearing flag is its literal value, verbatim — this is what
/// makes `--formula "-A1"` and flag-shaped formula text unambiguous.
fn takeValue(argv: []const []const u8, i: *usize, err_w: *std.Io.Writer, flag_name: []const u8) UsageError![]const u8 {
    i.* += 1;
    if (i.* >= argv.len) {
        err_w.print("zlsx: {s} requires a value\n", .{flag_name}) catch {};
        return error.Usage;
    }
    return argv[i.*];
}

const EvalArgs = struct {
    file: []const u8 = "",
    formula: ?[]const u8 = null,
    sheet: ?u32 = null,
    sheet_name: ?[]const u8 = null,
    anchor: ?engine.eval.EvalSite = null,
    anchor_text: ?[]const u8 = null,
    dialect: engine.value.Dialect = .dynamic_array,
    now_ms: ?i64 = null,
    utc_offset_min: i32 = 0,
    seed: ?u64 = null,
    fidelity: engine.value.Fidelity = .excel,
    deadline_s: ?u64 = null,
    help: bool = false,
};

const RecalcArgs = struct {
    file: []const u8 = "",
    out_path: ?[]const u8 = null,
    now_ms: ?i64 = null,
    utc_offset_min: i32 = 0,
    seed: ?u64 = null,
    fidelity: engine.value.Fidelity = .excel,
    on_unsupported: zlsx_pkg.recalc_txn.OnUnsupported = .refuse,
    report: bool = false,
    deadline_s: ?u64 = null,
    help: bool = false,
};

fn parseCommonValueFlag(
    a: []const u8,
    argv: []const []const u8,
    i: *usize,
    err_w: *std.Io.Writer,
    now_ms: *?i64,
    utc_offset_min: *i32,
    seed: *?u64,
    fidelity: *engine.value.Fidelity,
    deadline_s: *?u64,
) UsageError!bool {
    if (std.mem.eql(u8, a, "--now")) {
        const v = try takeValue(argv, i, err_w, "--now");
        now_ms.* = parseIso8601UtcMs(v) catch {
            err_w.print("zlsx: --now expects ISO 8601 (e.g. 2026-08-05T12:00:00Z), got '{s}'\n", .{v}) catch {};
            return error.Usage;
        };
        return true;
    }
    if (std.mem.eql(u8, a, "--utc-offset")) {
        const v = try takeValue(argv, i, err_w, "--utc-offset");
        const n = std.fmt.parseInt(i32, v, 10) catch {
            err_w.print("zlsx: --utc-offset expects minutes, got '{s}'\n", .{v}) catch {};
            return error.Usage;
        };
        if (n < engine.run_inputs.utc_offset_min_min or n > engine.run_inputs.utc_offset_min_max) {
            err_w.print("zlsx: --utc-offset out of range [-1440, 1440]: {d}\n", .{n}) catch {};
            return error.Usage;
        }
        utc_offset_min.* = n;
        return true;
    }
    if (std.mem.eql(u8, a, "--seed")) {
        const v = try takeValue(argv, i, err_w, "--seed");
        seed.* = std.fmt.parseInt(u64, v, 10) catch {
            err_w.print("zlsx: --seed expects a decimal u64, got '{s}'\n", .{v}) catch {};
            return error.Usage;
        };
        return true;
    }
    if (std.mem.eql(u8, a, "--mode")) {
        const v = try takeValue(argv, i, err_w, "--mode");
        if (std.mem.eql(u8, v, "excel")) {
            fidelity.* = .excel;
        } else if (std.mem.eql(u8, v, "ieee")) {
            fidelity.* = .ieee;
        } else {
            err_w.print("zlsx: --mode expects excel|ieee, got '{s}'\n", .{v}) catch {};
            return error.Usage;
        }
        return true;
    }
    if (std.mem.eql(u8, a, "--profile")) {
        const v = try takeValue(argv, i, err_w, "--profile");
        // One profile exists in v1; the flag validates rather than
        // silently accepting a name §5.4b never defined.
        if (!std.mem.eql(u8, v, "windows_1252")) {
            err_w.print("zlsx: --profile expects windows_1252, got '{s}'\n", .{v}) catch {};
            return error.Usage;
        }
        return true;
    }
    if (std.mem.eql(u8, a, "--deadline")) {
        const v = try takeValue(argv, i, err_w, "--deadline");
        deadline_s.* = std.fmt.parseInt(u64, v, 10) catch {
            err_w.print("zlsx: --deadline expects whole seconds, got '{s}'\n", .{v}) catch {};
            return error.Usage;
        };
        return true;
    }
    return false;
}

fn parseEvalArgs(argv: []const []const u8, err_w: *std.Io.Writer) UsageError!EvalArgs {
    var out: EvalArgs = .{};
    var i: usize = 0;
    while (i < argv.len) : (i += 1) {
        const a = argv[i];
        if (std.mem.eql(u8, a, "-h") or std.mem.eql(u8, a, "--help")) {
            out.help = true;
            return out;
        } else if (std.mem.eql(u8, a, "--formula")) {
            out.formula = try takeValue(argv, &i, err_w, "--formula");
        } else if (std.mem.eql(u8, a, "--sheet")) {
            const v = try takeValue(argv, &i, err_w, "--sheet");
            out.sheet = std.fmt.parseInt(u32, v, 10) catch {
                err_w.print("zlsx: --sheet expects a 0-based index, got '{s}'\n", .{v}) catch {};
                return error.Usage;
            };
        } else if (std.mem.eql(u8, a, "--name")) {
            out.sheet_name = try takeValue(argv, &i, err_w, "--name");
        } else if (std.mem.eql(u8, a, "--anchor")) {
            const v = try takeValue(argv, &i, err_w, "--anchor");
            const cell = coords.parseCell(v, .{}) catch {
                err_w.print("zlsx: --anchor expects an A1 reference, got '{s}'\n", .{v}) catch {};
                return error.Usage;
            };
            out.anchor = .{ .row = cell.row, .col = cell.col };
            out.anchor_text = v;
        } else if (std.mem.eql(u8, a, "--dialect")) {
            const v = try takeValue(argv, &i, err_w, "--dialect");
            if (std.mem.eql(u8, v, "da")) {
                out.dialect = .dynamic_array;
            } else if (std.mem.eql(u8, v, "legacy")) {
                out.dialect = .legacy;
            } else {
                err_w.print("zlsx: --dialect expects da|legacy, got '{s}'\n", .{v}) catch {};
                return error.Usage;
            }
        } else if (try parseCommonValueFlag(a, argv, &i, err_w, &out.now_ms, &out.utc_offset_min, &out.seed, &out.fidelity, &out.deadline_s)) {
            // handled
        } else if (a.len >= 1 and a[0] == '-') {
            err_w.print("zlsx: unknown eval flag '{s}'\n", .{a}) catch {};
            return error.Usage;
        } else if (out.file.len == 0) {
            out.file = a;
        } else {
            err_w.print("zlsx: unexpected positional '{s}'\n", .{a}) catch {};
            return error.Usage;
        }
    }
    if (out.file.len == 0) {
        err_w.writeAll("zlsx: eval needs an input workbook\n") catch {};
        return error.Usage;
    }
    if (out.formula == null) {
        err_w.writeAll("zlsx: eval needs --formula\n") catch {};
        return error.Usage;
    }
    // `--sheet N | --name NAME` is mandatory and exclusive — a formula
    // is evaluated AGAINST a sheet, and guessing sheet 0 would make
    // `Sheet2!`-relative results silently wrong.
    if ((out.sheet == null) == (out.sheet_name == null)) {
        err_w.writeAll("zlsx: eval needs exactly one of --sheet N | --name NAME\n") catch {};
        return error.Usage;
    }
    return out;
}

fn parseRecalcArgs(argv: []const []const u8, err_w: *std.Io.Writer) UsageError!RecalcArgs {
    var out: RecalcArgs = .{};
    var i: usize = 0;
    while (i < argv.len) : (i += 1) {
        const a = argv[i];
        if (std.mem.eql(u8, a, "-h") or std.mem.eql(u8, a, "--help")) {
            out.help = true;
            return out;
        } else if (std.mem.eql(u8, a, "--out")) {
            out.out_path = try takeValue(argv, &i, err_w, "--out");
        } else if (std.mem.eql(u8, a, "--on-unsupported")) {
            const v = try takeValue(argv, &i, err_w, "--on-unsupported");
            if (std.mem.eql(u8, v, "refuse")) {
                out.on_unsupported = .refuse;
            } else if (std.mem.eql(u8, v, "keep-stale-and-mark")) {
                out.on_unsupported = .keep_stale_and_mark;
            } else {
                err_w.print("zlsx: --on-unsupported expects refuse|keep-stale-and-mark, got '{s}'\n", .{v}) catch {};
                return error.Usage;
            }
        } else if (std.mem.eql(u8, a, "--report")) {
            out.report = true;
        } else if (try parseCommonValueFlag(a, argv, &i, err_w, &out.now_ms, &out.utc_offset_min, &out.seed, &out.fidelity, &out.deadline_s)) {
            // handled
        } else if (a.len >= 1 and a[0] == '-') {
            err_w.print("zlsx: unknown recalc flag '{s}'\n", .{a}) catch {};
            return error.Usage;
        } else if (out.file.len == 0) {
            out.file = a;
        } else {
            err_w.print("zlsx: unexpected positional '{s}'\n", .{a}) catch {};
            return error.Usage;
        }
    }
    if (out.file.len == 0) {
        err_w.writeAll("zlsx: recalc needs an input workbook\n") catch {};
        return error.Usage;
    }
    if (out.out_path == null) {
        err_w.writeAll("zlsx: recalc needs --out\n") catch {};
        return error.Usage;
    }
    return out;
}

fn writeEvalHelp(w: *std.Io.Writer) !void {
    try w.writeAll(
        \\usage: zlsx eval <file.xlsx> --formula "<text>" (--sheet N | --name NAME)
        \\                 [--anchor A1] [--dialect da|legacy] [--now ISO8601]
        \\                 [--utc-offset MIN] [--seed N] [--mode excel|ieee]
        \\                 [--profile windows_1252] [--deadline SECONDS]
        \\
        \\Evaluate one formula against a workbook and stream the result as
        \\versioned NDJSON (every record carries "kind" and "v":1).
        \\
        \\  --formula TEXT    the formula (a leading '=' is accepted and ignored);
        \\                    the next token is taken verbatim, so flag-shaped
        \\                    text like "-A1" needs no escaping
        \\  --sheet N         0-based sheet the formula is evaluated against
        \\  --name NAME       select that sheet by name instead (exactly one of
        \\                    --sheet / --name is required)
        \\  --anchor A1       evaluation site for site-dependent constructs
        \\                    (ROW(), COLUMN(), @) — refused as
        \\                    FormulaAnchorRequired when needed and absent
        \\  --dialect D       da (dynamic-array, default) | legacy
        \\  --now ISO8601     the instant NOW()/TODAY() report; default: the
        \\                    system wall clock (unreadable clock = exit 6)
        \\  --utc-offset MIN  fixed civil offset in minutes [-1440, 1440]
        \\  --seed N          decimal u64 RNG seed; default: the system secure
        \\                    random source (unreadable source = exit 6)
        \\  --mode M          excel (default) | ieee — §5.4's fidelity switch
        \\  --profile P       platform code page; windows_1252 is v1's only one
        \\  --deadline SECS   refuse (exit 3, "cancelled" record) once this
        \\                    many seconds elapse
        \\
        \\Exit: 0 evaluated (Excel error values are results), 1 usage,
        \\      2 open/parse, 3 typed refusal or --deadline, 4 out of memory,
        \\      5 stdout write failure, 6 default-context acquisition failure,
        \\      130/143 SIGINT/SIGTERM. SIGPIPE ends the stream early with
        \\      exit 0 and no terminal record.
        \\
    );
}

fn writeRecalcHelp(w: *std.Io.Writer) !void {
    try w.writeAll(
        \\usage: zlsx recalc <file.xlsx> --out <out.xlsx> [--now ISO8601]
        \\                   [--utc-offset MIN] [--seed N] [--mode excel|ieee]
        \\                   [--profile windows_1252]
        \\                   [--on-unsupported refuse|keep-stale-and-mark]
        \\                   [--report] [--deadline SECONDS]
        \\
        \\Recalculate every formula cell and write the result to --out in one
        \\atomic transaction (§5.7.9): any failure or cancellation before the
        \\rename leaves the destination byte-identical to what it held —
        \\"no file" only when it was absent. A signal that lands after the
        \\rename reports success (0): the commit already happened.
        \\
        \\  --out PATH        destination workbook; refused (exit 1) when it
        \\                    is the input, before anything is opened
        \\  --now ISO8601     the instant NOW()/TODAY() report; default: the
        \\                    system wall clock (unreadable clock = exit 6)
        \\  --utc-offset MIN  fixed civil offset in minutes [-1440, 1440]
        \\  --seed N          decimal u64 RNG seed; default: the system secure
        \\                    random source (unreadable source = exit 6)
        \\  --mode M          excel (default) | ieee
        \\  --profile P       platform code page; windows_1252 is v1's only one
        \\  --on-unsupported  refuse (default) — exit 3, nothing written;
        \\                    keep-stale-and-mark — keep the workbook's caches,
        \\                    set fullCalcOnLoad="1", report the census
        \\  --report          emit the NDJSON report stream on stdout
        \\                    (diagnostic* then recalc-report | refusal |
        \\                    cancelled); without it stdout stays silent
        \\  --deadline SECS   cancel (exit 3, destination untouched) once this
        \\                    many seconds elapse
        \\
        \\Exit: 0 recalced + written, 1 usage, 2 open/parse, 3 refusal or
        \\      cancellation with the destination untouched, 4 out of memory,
        \\      5 output write/rename failure, 6 default-context acquisition
        \\      failure, 130/143 SIGINT/SIGTERM before the rename only.
        \\
    );
}

// ─── ISO 8601 (UTC milliseconds) ─────────────────────────────────

const IsoError = error{BadIso8601};

/// Strict profile: `YYYY-MM-DD[THH:MM[:SS[.fff]]][Z|±HH:MM]`. No week
/// dates, no ordinal dates, no lowercase `t`/`z` — a reproducibility
/// input wants one spelling per instant, not a liberal parser.
fn parseIso8601UtcMs(text: []const u8) IsoError!i64 {
    if (text.len < 10) return error.BadIso8601;
    const year = parseDigits(i32, text[0..4]) orelse return error.BadIso8601;
    if (text[4] != '-') return error.BadIso8601;
    const month = parseDigits(u32, text[5..7]) orelse return error.BadIso8601;
    if (text[7] != '-') return error.BadIso8601;
    const day = parseDigits(u32, text[8..10]) orelse return error.BadIso8601;
    if (month < 1 or month > 12) return error.BadIso8601;
    if (day < 1 or day > engine.serial_date.daysInMonth(year, @intCast(month))) return error.BadIso8601;

    var hour: i64 = 0;
    var minute: i64 = 0;
    var second: i64 = 0;
    var milli: i64 = 0;
    var offset_min: i64 = 0;
    var rest = text[10..];

    if (rest.len > 0 and rest[0] == 'T') {
        if (rest.len < 6) return error.BadIso8601;
        hour = parseDigits(i64, rest[1..3]) orelse return error.BadIso8601;
        if (rest[3] != ':') return error.BadIso8601;
        minute = parseDigits(i64, rest[4..6]) orelse return error.BadIso8601;
        rest = rest[6..];
        if (rest.len >= 3 and rest[0] == ':') {
            second = parseDigits(i64, rest[1..3]) orelse return error.BadIso8601;
            rest = rest[3..];
            if (rest.len >= 2 and rest[0] == '.') {
                var n: usize = 1;
                var frac: i64 = 0;
                var scale: i64 = 100;
                while (n < rest.len and rest[n] >= '0' and rest[n] <= '9') : (n += 1) {
                    if (scale > 0) {
                        frac += scale * (rest[n] - '0');
                        scale = @divTrunc(scale, 10);
                    }
                }
                if (n == 1) return error.BadIso8601;
                milli = frac;
                rest = rest[n..];
            }
        }
        if (hour > 23 or minute > 59 or second > 59) return error.BadIso8601;
    }

    if (rest.len > 0) {
        if (rest.len == 1 and rest[0] == 'Z') {
            // UTC, which is also the default.
        } else if (rest.len == 6 and (rest[0] == '+' or rest[0] == '-')) {
            const oh = parseDigits(i64, rest[1..3]) orelse return error.BadIso8601;
            if (rest[3] != ':') return error.BadIso8601;
            const om = parseDigits(i64, rest[4..6]) orelse return error.BadIso8601;
            if (oh > 23 or om > 59) return error.BadIso8601;
            offset_min = oh * 60 + om;
            if (rest[0] == '-') offset_min = -offset_min;
        } else {
            return error.BadIso8601;
        }
    }

    const days = engine.serial_date.daysFromCivil(year, month, day);
    const local_ms = days * std.time.ms_per_day +
        hour * std.time.ms_per_hour + minute * std.time.ms_per_min +
        second * std.time.ms_per_s + milli;
    return local_ms - offset_min * std.time.ms_per_min;
}

fn parseDigits(comptime T: type, s: []const u8) ?T {
    var v: T = 0;
    for (s) |c| {
        if (c < '0' or c > '9') return null;
        v = v * 10 + (c - '0');
    }
    return v;
}

/// The one spelling `resolved.now` uses: `YYYY-MM-DDTHH:MM:SS.mmmZ`.
fn writeIsoUtc(w: *std.Io.Writer, ms: i64) !void {
    const days = @divFloor(ms, std.time.ms_per_day);
    const in_day = ms - days * std.time.ms_per_day;
    const civil = engine.serial_date.civilFromDays(days);
    const s_total = @divTrunc(in_day, std.time.ms_per_s);
    const milli = in_day - s_total * std.time.ms_per_s;
    try w.print("{d:0>4}-{d:0>2}-{d:0>2}T{d:0>2}:{d:0>2}:{d:0>2}.{d:0>3}Z", .{
        @as(u32, @intCast(civil.year)),                       @as(u32, civil.month),
        @as(u32, civil.day),                                  @as(u64, @intCast(@divTrunc(s_total, 3600))),
        @as(u64, @intCast(@mod(@divTrunc(s_total, 60), 60))), @as(u64, @intCast(@mod(s_total, 60))),
        @as(u64, @intCast(milli)),
    });
}

// ─── the resolved run, shared by both commands ───────────────────

const Resolved = struct {
    now_ms: i64,
    utc_offset_min: i32,
    seed: u64,
    fidelity: engine.value.Fidelity,
    /// Absolute `.awake` deadline, armed from `--deadline`.
    deadline: ?std.Io.Timestamp,
};

/// Resolve the defaults the caller omitted. This is the only site exit 6
/// can originate from, and it is deliberately BEFORE the workbook opens:
/// a run that cannot know what time it is should not read a file first.
fn resolveDefaults(
    io: std.Io,
    deps: Deps,
    now_ms: ?i64,
    seed: ?u64,
    utc_offset_min: i32,
    fidelity: engine.value.Fidelity,
    deadline_s: ?u64,
    err_w: *std.Io.Writer,
    cmd: []const u8,
) error{Context}!Resolved {
    const now = now_ms orelse deps.ctx.now_ms(io) catch {
        err_w.print("zlsx {s}: cannot read the wall clock to default --now\n", .{cmd}) catch {};
        return error.Context;
    };
    const s = seed orelse deps.ctx.seed(io) catch {
        err_w.print("zlsx {s}: cannot read the random source to default --seed\n", .{cmd}) catch {};
        return error.Context;
    };
    const deadline: ?std.Io.Timestamp = if (deadline_s) |secs| blk: {
        const base = std.Io.Timestamp.now(io, .awake);
        break :blk base.addDuration(.{ .nanoseconds = @as(i96, @intCast(secs)) * std.time.ns_per_s });
    } else null;
    return .{
        .now_ms = now,
        .utc_offset_min = utc_offset_min,
        .seed = s,
        .fidelity = fidelity,
        .deadline = deadline,
    };
}

fn deadlinePassed(io: std.Io, deadline: ?std.Io.Timestamp) bool {
    const d = deadline orelse return false;
    return std.Io.Timestamp.now(io, .awake).nanoseconds >= d.nanoseconds;
}

// ─── the stream ──────────────────────────────────────────────────

const RecordKind = enum {
    none,
    @"eval-header",
    @"eval-cell",
    diagnostic,
    @"eval-complete",
    refusal,
    cancelled,
    @"recalc-report",
};

const EmitError = error{ WriteFailed, PipeClosed, Stopped };

/// The NDJSON stream state machine's writing half. One rule from the
/// shipped envelope carries over: poll BEFORE starting a record, so
/// every line on stdout is complete (`cli.zig`'s mid-record discard
/// contract). Terminal records ignore `Stopped` — they are how a stop is
/// reported — but not `PipeClosed`: a closed pipe gets nothing.
const Emitter = struct {
    w: *std.Io.Writer,
    sig: SignalState,
    deps: *const Deps,
    emitted: usize = 0,
    last: RecordKind = .none,

    /// Data records: refuse to start once a signal fired.
    fn beginData(self: *Emitter) EmitError!void {
        if (self.deps.before_record) |hook| hook(self.deps.before_record_state, self.emitted);
        if (self.sig.pipeFired()) return error.PipeClosed;
        if (self.sig.stopped()) return error.Stopped;
    }

    /// Terminal records: written even when stopping — unless the pipe
    /// is gone, which is the SIGPIPE exception.
    fn beginTerminal(self: *Emitter) EmitError!void {
        if (self.deps.before_record) |hook| hook(self.deps.before_record_state, self.emitted);
        if (self.sig.pipeFired()) return error.PipeClosed;
    }

    fn finishRecord(self: *Emitter, kind: RecordKind) EmitError!void {
        self.w.writeByte('\n') catch return self.writeFail();
        self.w.flush() catch return self.writeFail();
        self.emitted += 1;
        self.last = kind;
    }

    /// A write failure after SIGPIPE IS the pipe closing; anything else
    /// is exit 5's row.
    fn writeFail(self: *Emitter) EmitError {
        if (self.sig.pipeFired()) return error.PipeClosed;
        return error.WriteFailed;
    }

    fn head(self: *Emitter, kind: RecordKind) EmitError!void {
        self.w.print("{{\"kind\":\"{s}\",\"v\":{d}", .{ @tagName(kind), stream_version }) catch
            return self.writeFail();
    }

    fn cancelled(self: *Emitter) EmitError!void {
        try self.beginTerminal();
        try self.head(.cancelled);
        self.w.print(",\"after\":\"{s}\"}}", .{@tagName(self.last)}) catch return self.writeFail();
        try self.finishRecord(.cancelled);
    }

    fn refusal(self: *Emitter, error_name: []const u8, cell: ?[]const u8) EmitError!void {
        try self.beginTerminal();
        try self.head(.refusal);
        self.w.writeAll(",\"error\":") catch return self.writeFail();
        writeJsonString(self.w, error_name) catch return self.writeFail();
        self.w.writeAll(",\"cells\":[") catch return self.writeFail();
        if (cell) |c| writeJsonString(self.w, c) catch return self.writeFail();
        self.w.writeAll("],\"truncated\":false}") catch return self.writeFail();
        try self.finishRecord(.refusal);
    }

    fn diagnostic(self: *Emitter, severity: []const u8, message: []const u8) EmitError!void {
        try self.beginData();
        try self.head(.diagnostic);
        self.w.print(",\"severity\":\"{s}\",\"message\":", .{severity}) catch return self.writeFail();
        writeJsonString(self.w, message) catch return self.writeFail();
        self.w.writeByte('}') catch return self.writeFail();
        try self.finishRecord(.diagnostic);
    }
};

/// JSON string escaping. Same escapes the shipped envelope uses
/// (`cli.zig`'s private twin): quote, backslash, and C0 controls.
fn writeJsonString(w: *std.Io.Writer, s: []const u8) !void {
    try w.writeByte('"');
    for (s) |c| {
        switch (c) {
            '"' => try w.writeAll("\\\""),
            '\\' => try w.writeAll("\\\\"),
            '\n' => try w.writeAll("\\n"),
            '\r' => try w.writeAll("\\r"),
            '\t' => try w.writeAll("\\t"),
            0...8, 11, 12, 14...31 => try w.print("\\u{x:0>4}", .{c}),
            else => try w.writeByte(c),
        }
    }
    try w.writeByte('"');
}

fn writePublishedValue(w: *std.Io.Writer, p: engine.value.PublishedScalar) !void {
    switch (p) {
        .number => |n| try w.print("{d}", .{n}),
        .text => |t| try writeJsonString(w, t),
        .boolean => |b| try w.writeAll(if (b) "true" else "false"),
        .err => |e| try writeJsonString(w, e.spelling()),
    }
}

fn publishedTypeName(p: engine.value.PublishedScalar) []const u8 {
    return switch (p) {
        .number => "number",
        .text => "text",
        .boolean => "bool",
        .err => "error",
    };
}

const ResolvedEcho = struct {
    r: Resolved,
    /// `eval` only; recalc derives dialect per stored cell (§5.3b).
    dialect: ?engine.value.Dialect = null,
    anchor_text: ?[]const u8 = null,
};

/// §5.7.8's resolved-input echo. `seed` serializes as a DECIMAL STRING:
/// u64 exceeds the JSON/JS safe-integer range, and a seed that rounds
/// is a seed that cannot reproduce the run it labels.
fn writeResolvedEcho(w: *std.Io.Writer, echo: ResolvedEcho) !void {
    try w.writeAll("\"resolved\":{\"now\":\"");
    try writeIsoUtc(w, echo.r.now_ms);
    try w.print("\",\"utcOffsetMin\":{d},\"seed\":\"{d}\",\"mode\":\"{s}\",\"profile\":\"windows_1252\"", .{
        echo.r.utc_offset_min,
        echo.r.seed,
        @tagName(echo.r.fidelity),
    });
    if (echo.dialect) |d| {
        try w.print(",\"dialect\":\"{s}\"", .{switch (d) {
            .dynamic_array => "da",
            .legacy => "legacy",
        }});
    }
    if (echo.anchor_text) |a| {
        try w.writeAll(",\"anchor\":");
        try writeJsonString(w, a);
    }
    try w.writeByte('}');
}

// ─── eval ────────────────────────────────────────────────────────

fn runEval(
    gpa: Allocator,
    io: std.Io,
    argv: []const []const u8,
    out: *std.Io.Writer,
    err_w: *std.Io.Writer,
    deps: Deps,
) u8 {
    const args = parseEvalArgs(argv, err_w) catch return exit_usage;
    if (args.help) {
        writeEvalHelp(out) catch return exit_write;
        out.flush() catch return exit_write;
        return exit_ok;
    }

    const resolved = resolveDefaults(
        io,
        deps,
        args.now_ms,
        args.seed,
        args.utc_offset_min,
        args.fidelity,
        args.deadline_s,
        err_w,
        "eval",
    ) catch return exit_context;

    var em: Emitter = .{ .w = out, .sig = deps.sig, .deps = &deps };

    // Cancellation and the deadline can both fire before any header
    // exists — the grammar's first alternative.
    if (checkStop(io, &em, resolved.deadline)) |code| return code;

    var wb = zlsx_pkg.Workbook.open(gpa, io, args.file) catch |e| switch (e) {
        error.OutOfMemory => return exit_oom,
        else => {
            err_w.print("zlsx eval: cannot open '{s}': {s}\n", .{ args.file, @errorName(e) }) catch {};
            return exit_open;
        },
    };
    defer wb.deinit();

    // Sheet resolution failure is an input problem — the workbook does
    // not contain the requested evaluation context — not a §10 refusal
    // (the CLI maps refusals, it does not invent them), so it sits in
    // exit 2's row with open/parse.
    const sheet_index: u32 = if (args.sheet) |n| blk: {
        if (n >= wb.worksheets.len) {
            err_w.print("zlsx eval: sheet {d} out of range ({d} sheets)\n", .{ n, wb.worksheets.len }) catch {};
            return exit_open;
        }
        break :blk n;
    } else blk: {
        const want = args.sheet_name.?;
        for (wb.workbook.sheets, 0..) |s, idx| {
            if (std.mem.eql(u8, s.name, want)) break :blk @intCast(idx);
        }
        err_w.print("zlsx eval: no sheet named '{s}'\n", .{want}) catch {};
        return exit_open;
    };

    if (checkStop(io, &em, resolved.deadline)) |code| return code;

    // A leading '=' is the parser's business, not ours: M2 already
    // strips exactly one (`leading_eq: .optional`) and refuses `==` as
    // `double_equals`. A CLI-side strip on top would swallow that
    // refusal — `==1` would silently evaluate as `1`.
    var result = wb.evaluate(gpa, sheet_index, args.formula.?, .{
        .collation = zlsx_pkg.recalc_run.collation_v1,
        .fidelity = resolved.fidelity,
        .dialect = args.dialect,
        .site = args.anchor,
        .now_utc_ms = resolved.now_ms,
        .utc_offset_min = resolved.utc_offset_min,
    }) catch |e| switch (e) {
        error.OutOfMemory => return exit_oom,
        else => {
            err_w.print("zlsx eval: cannot model '{s}': {s}\n", .{ args.file, @errorName(e) }) catch {};
            return exit_open;
        },
    };
    defer result.deinit();

    if (checkStop(io, &em, resolved.deadline)) |code| return code;

    const echo: ResolvedEcho = .{
        .r = resolved,
        .dialect = args.dialect,
        .anchor_text = args.anchor_text,
    };

    switch (result) {
        .ok => |*evaluation| return emitEvaluation(io, &em, err_w, evaluation, resolved, echo),
        .refused => |r| return emitRefusal(&em, err_w, @tagName(r.planeTwo()), null),
        .parse_refused => |r| return emitRefusal(&em, err_w, @tagName(r.planeTwo()), null),
        .graph_refused => |r| {
            var cell_buf: [refusal_cell_buf_len]u8 = undefined;
            return emitRefusal(&em, err_w, @tagName(r.planeTwo()), refusalCell(&cell_buf, &wb, r.at));
        },
        .iteration_refused => |r| {
            var cell_buf: [refusal_cell_buf_len]u8 = undefined;
            return emitRefusal(&em, err_w, @tagName(r.planeTwo()), refusalCell(&cell_buf, &wb, r.at));
        },
        .eval_refused => |r| return emitRefusal(&em, err_w, @tagName(r.plane), null),
    }
}

/// Longest cell label: a 31-scalar sheet name (OOXML's cap, ×4 UTF-8
/// bytes), '!', 3 column letters, 7 row digits.
const refusal_cell_buf_len = 31 * 4 + 1 + 3 + 7;

/// `Sheet1!B7` for a refusal that names a cell node; null for every
/// other key kind. Stack-rendered — a refusal must not fail to be
/// reported because its label could not be allocated.
fn refusalCell(buf: *[refusal_cell_buf_len]u8, wb: *zlsx_pkg.Workbook, at: ?engine.graph.Key) ?[]const u8 {
    const key = at orelse return null;
    const cell = switch (key) {
        .cell, .spill_tail => |c| c,
        else => return null,
    };
    const sheet_idx = @intFromEnum(cell.sheet);
    if (sheet_idx >= wb.workbook.sheets.len) return null;
    // M0's one formatter — the import gate exists precisely so this
    // file does not grow the tree's seventh base-26 formatter.
    var col_buf: [coords.max_col_letters]u8 = undefined;
    const n = coords.writeColLetters(&col_buf, cell.col);
    return std.fmt.bufPrint(buf, "{s}!{s}{d}", .{
        wb.workbook.sheets[sheet_idx].name, col_buf[0..n], cell.row.oneBased(),
    }) catch null;
}

/// One place decides which exit code a stop is (deadline → 3, SIGINT →
/// 130, SIGTERM → 143, SIGPIPE → 0) and emits the `cancelled` terminal
/// where a terminal is still possible. Null means: keep going.
fn checkStop(io: std.Io, em: *Emitter, deadline: ?std.Io.Timestamp) ?u8 {
    if (em.sig.pipeFired()) return exit_ok;
    if (em.sig.intFired() or em.sig.termFired()) {
        const code: u8 = if (em.sig.intFired()) exit_sigint else exit_sigterm;
        em.cancelled() catch |e| switch (e) {
            error.PipeClosed => return exit_ok,
            error.WriteFailed => return exit_write,
            error.Stopped => unreachable, // terminals ignore the stop flag
        };
        return code;
    }
    if (deadlinePassed(io, deadline)) {
        em.cancelled() catch |e| switch (e) {
            error.PipeClosed => return exit_ok,
            error.WriteFailed => return exit_write,
            error.Stopped => unreachable,
        };
        return exit_refusal;
    }
    return null;
}

fn emitRefusal(em: *Emitter, err_w: *std.Io.Writer, error_name: []const u8, cell: ?[]const u8) u8 {
    err_w.print("zlsx: refused: {s}\n", .{error_name}) catch {};
    em.refusal(error_name, cell) catch |e| switch (e) {
        error.PipeClosed => return exit_ok,
        error.WriteFailed => return exit_write,
        error.Stopped => unreachable,
    };
    return exit_refusal;
}

fn emitEvaluation(
    io: std.Io,
    em: *Emitter,
    err_w: *std.Io.Writer,
    evaluation: *zlsx_pkg.Evaluation,
    resolved: Resolved,
    echo: ResolvedEcho,
) u8 {
    _ = err_w;
    emitEvaluationInner(io, em, evaluation, resolved, echo) catch |e| switch (e) {
        // The SIGPIPE exception: prefix-valid stream, no terminal,
        // success — `head -1` is not an error.
        error.PipeClosed => return exit_ok,
        error.WriteFailed => return exit_write,
        // SIGINT/SIGTERM between records: close with the `cancelled`
        // terminal and report the signal's code.
        error.Stopped => {
            em.cancelled() catch |e2| switch (e2) {
                error.PipeClosed => return exit_ok,
                error.WriteFailed => return exit_write,
                error.Stopped => unreachable,
            };
            return if (em.sig.intFired()) exit_sigint else exit_sigterm;
        },
    };
    return exit_ok;
}

fn emitEvaluationInner(
    io: std.Io,
    em: *Emitter,
    evaluation: *zlsx_pkg.Evaluation,
    resolved: Resolved,
    echo: ResolvedEcho,
) EmitError!void {
    _ = io;
    const fidelity = resolved.fidelity;
    switch (evaluation.value) {
        .array => |m| {
            // Matrix: shape in the header, cells as records, row-major.
            try em.beginData();
            try em.head(.@"eval-header");
            em.w.print(",\"type\":\"matrix\",\"rows\":{d},\"cols\":{d},", .{ m.rows, m.cols }) catch
                return em.writeFail();
            writeResolvedEcho(em.w, echo) catch return em.writeFail();
            em.w.writeByte('}') catch return em.writeFail();
            try em.finishRecord(.@"eval-header");

            var idx: usize = 0;
            var r: u32 = 1;
            while (r <= m.rows) : (r += 1) {
                var c: u32 = 1;
                while (c <= m.cols) : (c += 1) {
                    const p = engine.value.publish(m.cells[idx], fidelity);
                    idx += 1;
                    try em.beginData();
                    try em.head(.@"eval-cell");
                    em.w.print(",\"r\":{d},\"c\":{d},\"type\":\"{s}\",\"value\":", .{ r, c, publishedTypeName(p) }) catch
                        return em.writeFail();
                    writePublishedValue(em.w, p) catch return em.writeFail();
                    em.w.writeByte('}') catch return em.writeFail();
                    try em.finishRecord(.@"eval-cell");
                }
            }
            try emitComplete(em, m.cells.len);
        },
        else => {
            // Scalar: the value rides in the header and no eval-cell
            // records follow. `publish` is §5.3a's one mandatory
            // blank→0 conversion — `=A1` on an empty A1 emits 0 here,
            // and the word "blank" exists nowhere in the stream.
            const scalar: engine.value.ScalarValue = switch (evaluation.value) {
                .scalar => |s| s,
                .missing_arg => .blank,
                .array => unreachable,
                .reference => unreachable, // dereferenced before return (§5.3b)
            };
            const p = engine.value.publish(scalar, fidelity);
            try em.beginData();
            try em.head(.@"eval-header");
            em.w.print(",\"type\":\"{s}\",\"value\":", .{publishedTypeName(p)}) catch
                return em.writeFail();
            writePublishedValue(em.w, p) catch return em.writeFail();
            em.w.writeByte(',') catch return em.writeFail();
            writeResolvedEcho(em.w, echo) catch return em.writeFail();
            em.w.writeByte('}') catch return em.writeFail();
            try em.finishRecord(.@"eval-header");
            try emitComplete(em, 0);
        },
    }
}

/// `cells` counts the eval-cell records emitted: 0 for a scalar result.
fn emitComplete(em: *Emitter, cells: usize) EmitError!void {
    try em.beginTerminal();
    try em.head(.@"eval-complete");
    em.w.print(",\"cells\":{d}}}", .{cells}) catch return em.writeFail();
    try em.finishRecord(.@"eval-complete");
}

// ─── recalc ──────────────────────────────────────────────────────

fn runRecalc(
    gpa: Allocator,
    io: std.Io,
    argv: []const []const u8,
    out: *std.Io.Writer,
    err_w: *std.Io.Writer,
    deps: Deps,
) u8 {
    const args = parseRecalcArgs(argv, err_w) catch return exit_usage;
    if (args.help) {
        writeRecalcHelp(out) catch return exit_write;
        out.flush() catch return exit_write;
        return exit_ok;
    }
    const out_path = args.out_path.?;

    // `--out` identity is refused before ANY mutation — before the input
    // is even opened. Byte-equal paths are enough to refuse; different
    // spellings of one file are caught through the filesystem's answer,
    // and a destination that does not resolve is simply not the input.
    if (std.mem.eql(u8, args.file, out_path) or pathsAlias(gpa, io, args.file, out_path)) {
        err_w.print("zlsx recalc: --out must not be the input ('{s}')\n", .{out_path}) catch {};
        return exit_usage;
    }

    const resolved = resolveDefaults(
        io,
        deps,
        args.now_ms,
        args.seed,
        args.utc_offset_min,
        args.fidelity,
        args.deadline_s,
        err_w,
        "recalc",
    ) catch return exit_context;

    var em: Emitter = .{ .w = out, .sig = deps.sig, .deps = &deps };

    // The recalc stream only exists with --report; a silent Emitter
    // keeps one code path. Pre-open stop: destination trivially
    // untouched.
    if (stopBeforeCommit(io, &em, resolved.deadline, args.report)) |code| return code;

    var wb = zlsx_pkg.Workbook.open(gpa, io, args.file) catch |e| switch (e) {
        error.OutOfMemory => return exit_oom,
        else => {
            err_w.print("zlsx recalc: cannot open '{s}': {s}\n", .{ args.file, @errorName(e) }) catch {};
            return exit_open;
        },
    };
    defer wb.deinit();

    const run_inputs: zlsx_pkg.RunInputs = .{
        .now_utc_ms = resolved.now_ms,
        .rng_seed = resolved.seed,
        .limits = .{},
        .utc_offset_min = resolved.utc_offset_min,
        .fidelity = resolved.fidelity,
        .deadline = resolved.deadline,
        .cancel = deps.sig.cancelToken(),
    };

    var report = wb.saveWithRecalc(gpa, io, out_path, run_inputs, .{
        .on_unsupported = args.on_unsupported,
    }) catch |e| return mapRecalcError(e, &em, err_w, args.report);

    defer report.deinit(gpa);

    // From here the rename has happened: the destination durably holds
    // the recalced bytes and the outcome is success — §12.2's
    // commit-aware mapping. Signals that arrive now change nothing;
    // only the report stream can still fail (exit 5), and SIGPIPE on
    // that stream is still the shipped exception.
    if (args.report) {
        emitRecalcReport(&em, &wb, &report, .{ .r = resolved }) catch |e| switch (e) {
            error.PipeClosed => return exit_ok,
            error.WriteFailed => return exit_write,
            error.Stopped => unreachable, // report path uses terminal begins only
        };
    }
    if (report.durability.warning) {
        err_w.writeAll("zlsx recalc: post-commit directory fsync failed; contents are written, the directory entry may not be durable\n") catch {};
    }
    return mapRecalcSignalExit(true, deps.sig);
}

/// Pre-commit stop check. Emits `cancelled` on the stream only when
/// `--report` opened one.
fn stopBeforeCommit(io: std.Io, em: *Emitter, deadline: ?std.Io.Timestamp, report: bool) ?u8 {
    if (em.sig.pipeFired()) return if (report) exit_ok else null;
    const code: ?u8 = if (em.sig.intFired())
        exit_sigint
    else if (em.sig.termFired())
        exit_sigterm
    else if (deadlinePassed(io, deadline))
        exit_refusal
    else
        null;
    if (code) |c| {
        if (report) {
            em.cancelled() catch |e| switch (e) {
                error.PipeClosed => return exit_ok,
                error.WriteFailed => return exit_write,
                error.Stopped => unreachable,
            };
        }
        return c;
    }
    return null;
}

/// The recalc error → exit-code mapping. §5.7.9 guarantees the
/// destination's prior bytes on every one of these paths — the
/// transaction cannot fail after the rename.
fn mapRecalcError(e: anyerror, em: *Emitter, err_w: *std.Io.Writer, report: bool) u8 {
    switch (e) {
        error.OutOfMemory => return exit_oom,
        error.Cancelled => {
            if (report) {
                em.cancelled() catch |e2| switch (e2) {
                    error.PipeClosed => return exit_ok,
                    error.WriteFailed => return exit_write,
                    error.Stopped => unreachable,
                };
            }
            return mapRecalcSignalExit(false, em.sig);
        },
        else => {},
    }
    const name = @errorName(e);
    if (std.mem.startsWith(u8, name, "Formula")) {
        // §10's plane-2 namespace, verbatim — the CLI maps refusals, it
        // does not invent them.
        err_w.print("zlsx recalc: refused: {s}\n", .{name}) catch {};
        if (report) {
            em.refusal(name, null) catch |e2| switch (e2) {
                error.PipeClosed => return exit_ok,
                error.WriteFailed => return exit_write,
                error.Stopped => unreachable,
            };
        }
        return exit_refusal;
    }
    if (std.mem.startsWith(u8, name, "Malformed")) {
        // Input-side decode failures surfacing through the pipeline are
        // the same class as open failures.
        err_w.print("zlsx recalc: cannot model the workbook: {s}\n", .{name}) catch {};
        return exit_open;
    }
    err_w.print("zlsx recalc: cannot write the output: {s}\n", .{name}) catch {};
    return exit_write;
}

fn emitRecalcReport(
    em: *Emitter,
    wb: *zlsx_pkg.Workbook,
    report: *const zlsx_pkg.RecalcReport,
    echo: ResolvedEcho,
) EmitError!void {
    if (report.durability.warning) {
        // The one post-commit fact §5.7.9 demotes to a warning; ride it
        // as a diagnostic ahead of the terminal, where the grammar puts
        // diagnostics. `beginTerminal` semantics via the plain data
        // begin would drop it after a late signal — acceptable: it is
        // advisory, and the report itself still closes the stream.
        em.diagnostic("warning", "post-commit directory fsync failed; contents are written, the directory entry may not be durable") catch |e| switch (e) {
            error.Stopped => {}, // advisory; the terminal still goes out
            else => return e,
        };
    }
    try em.beginTerminal();
    try em.head(.@"recalc-report");
    em.w.print(",\"sheets\":{d},\"cells\":{d},\"passes\":{d},\"nonConverged\":{d},\"dynamicPasses\":{d},\"keptStale\":{},\"calcChainRemoved\":{},\"census\":[", .{
        report.sheets_patched,
        report.cells_written,
        report.passes,
        report.non_converged_cells,
        report.dynamic_passes,
        report.kept_stale,
        report.calc_chain_removed,
    }) catch return em.writeFail();
    for (report.census, 0..) |entry, i| {
        if (i > 0) em.w.writeByte(',') catch return em.writeFail();
        em.w.writeAll("{\"error\":") catch return em.writeFail();
        writeJsonString(em.w, @tagName(entry.plane)) catch return em.writeFail();
        if (entry.row != 0 and entry.sheet < wb.workbook.sheets.len) {
            // `Unsupported.col` is zero-based; M0's formatter takes the
            // 1-based number.
            var col_buf: [coords.max_col_letters]u8 = undefined;
            const n = coords.writeColNumberLetters(&col_buf, entry.col + 1) catch return em.writeFail();
            em.w.writeAll(",\"cell\":") catch return em.writeFail();
            var name_buf: [refusal_cell_buf_len]u8 = undefined;
            const label = std.fmt.bufPrint(&name_buf, "{s}!{s}{d}", .{
                wb.workbook.sheets[entry.sheet].name, col_buf[0..n], entry.row,
            }) catch return em.writeFail();
            writeJsonString(em.w, label) catch return em.writeFail();
        }
        em.w.writeByte('}') catch return em.writeFail();
    }
    em.w.print("],\"censusTruncated\":{},", .{report.census_truncated}) catch return em.writeFail();
    writeResolvedEcho(em.w, echo) catch return em.writeFail();
    em.w.writeByte('}') catch return em.writeFail();
    try em.finishRecord(.@"recalc-report");
}

/// Do two paths name the same file? Resolved through the filesystem so
/// `./out.xlsx` and `out.xlsx` cannot slip past the identity refusal.
/// A destination that does not exist yet resolves to nothing and
/// aliases nothing; failures fall back to "not identical" — the
/// byte-equality check above already caught the literal case, and open
/// errors get their own exit row.
fn pathsAlias(gpa: Allocator, io: std.Io, a_path: []const u8, b_path: []const u8) bool {
    const cwd = std.Io.Dir.cwd();
    const ra = cwd.realPathFileAlloc(io, a_path, gpa) catch return false;
    defer gpa.free(ra);
    const rb = cwd.realPathFileAlloc(io, b_path, gpa) catch return false;
    defer gpa.free(rb);
    return std.mem.eql(u8, ra, rb);
}

// ─────────────────────────────────────────────────────────────────
// M6 contract tests. Every production of both stream grammars, every
// row of the exit table, the SIGPIPE exception, and the commit seam —
// all driven through injected `Deps`, none through timing.
// ─────────────────────────────────────────────────────────────────

const testing = std.testing;

const t_ns_main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const t_ns_r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const t_ct_sheet = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
const t_ct_workbook = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
const t_ct_rels = "application/vnd.openxmlformats-package.relationships+xml";

/// A1=1, B1=2, A2=3, B2=4 — the matrix source and eval baseline.
const t_sheet_values =
    "<worksheet xmlns=\"" ++ t_ns_main ++ "\"><sheetData>" ++
    "<row r=\"1\"><c r=\"A1\"><v>1</v></c><c r=\"B1\"><v>2</v></c></row>" ++
    "<row r=\"2\"><c r=\"A2\"><v>3</v></c><c r=\"B2\"><v>4</v></c></row>" ++
    "</sheetData></worksheet>";

/// A1 deliberately empty; only B1 holds a value. `=A1` must publish 0.
const t_sheet_empty_a1 =
    "<worksheet xmlns=\"" ++ t_ns_main ++ "\"><sheetData><row r=\"1\">" ++
    "<c r=\"B1\"><v>5</v></c>" ++
    "</row></sheetData></worksheet>";

/// B1 = A1+1 with a stale cache — a recalc has something to change.
const t_sheet_stale =
    "<worksheet xmlns=\"" ++ t_ns_main ++ "\"><sheetData><row r=\"1\">" ++
    "<c r=\"A1\"><v>1</v></c><c r=\"B1\"><f>A1+1</f><v>999</v></c>" ++
    "</row></sheetData></worksheet>";

/// B1 calls a function no registry row implements.
const t_sheet_unsupported =
    "<worksheet xmlns=\"" ++ t_ns_main ++ "\"><sheetData><row r=\"1\">" ++
    "<c r=\"A1\"><v>1</v></c><c r=\"B1\"><f>NOTAFUNC(A1)</f><v>999</v></c>" ++
    "</row></sheetData></worksheet>";

fn writeTestFixture(gpa: Allocator, io: std.Io, dir: []const u8, name: []const u8, sheet_xml: []const u8) ![]u8 {
    const path = try std.fs.path.join(gpa, &.{ dir, name });
    errdefer gpa.free(path);

    var store = try zlsx_pkg.PartStore.fresh(gpa, io);
    defer store.deinit();

    try store.addPart("_rels/.rels", t_ct_rels, "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" ++
        "<Relationship Id=\"rId1\" Type=\"" ++ t_ns_r ++ "/officeDocument\" Target=\"xl/workbook.xml\"/>" ++
        "</Relationships>");
    try store.addPart("xl/workbook.xml", t_ct_workbook, "<workbook xmlns=\"" ++ t_ns_main ++ "\" xmlns:r=\"" ++ t_ns_r ++ "\">" ++
        "<sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets><calcPr calcId=\"191029\"/></workbook>");
    try store.addPart("xl/_rels/workbook.xml.rels", t_ct_rels, "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" ++
        "<Relationship Id=\"rId1\" Type=\"" ++ t_ns_r ++ "/worksheet\" Target=\"worksheets/sheet1.xml\"/>" ++
        "</Relationships>");
    try store.addPart("xl/worksheets/sheet1.xml", t_ct_sheet, sheet_xml);

    try store.save(io, path);
    return path;
}

fn testTmpPath(gpa: Allocator, io: std.Io, tmp: *testing.TmpDir) ![:0]u8 {
    return tmp.dir.realPathFileAlloc(io, ".", gpa);
}

const Driven = struct {
    code: u8,
    out: []u8,
    err: []u8,

    fn deinit(self: *Driven, gpa: Allocator) void {
        gpa.free(self.out);
        gpa.free(self.err);
        self.* = undefined;
    }
};

fn drive(gpa: Allocator, io: std.Io, argv: []const []const u8, deps: Deps) !Driven {
    var out_aw: std.Io.Writer.Allocating = .init(gpa);
    defer out_aw.deinit();
    var err_aw: std.Io.Writer.Allocating = .init(gpa);
    defer err_aw.deinit();
    const code = run(gpa, io, argv, &out_aw.writer, &err_aw.writer, deps);
    return .{
        .code = code,
        .out = try gpa.dupe(u8, out_aw.writer.buffered()),
        .err = try gpa.dupe(u8, err_aw.writer.buffered()),
    };
}

/// Every line parses as JSON and carries `"kind"` and `"v":1`; the
/// kinds are returned in order for grammar assertions.
fn streamKinds(gpa: Allocator, stream: []const u8) ![][]u8 {
    var kinds: std.ArrayList([]u8) = .empty;
    errdefer {
        for (kinds.items) |k| gpa.free(k);
        kinds.deinit(gpa);
    }
    var it = std.mem.splitScalar(u8, stream, '\n');
    while (it.next()) |line| {
        if (line.len == 0) continue;
        var parsed = try std.json.parseFromSlice(std.json.Value, gpa, line, .{});
        defer parsed.deinit();
        const obj = parsed.value.object;
        try testing.expectEqual(@as(i64, stream_version), obj.get("v").?.integer);
        try kinds.append(gpa, try gpa.dupe(u8, obj.get("kind").?.string));
    }
    return kinds.toOwnedSlice(gpa);
}

fn freeKinds(gpa: Allocator, kinds: [][]u8) void {
    for (kinds) |k| gpa.free(k);
    gpa.free(kinds);
}

fn expectKinds(gpa: Allocator, stream: []const u8, expected: []const []const u8) !void {
    const kinds = try streamKinds(gpa, stream);
    defer freeKinds(gpa, kinds);
    try testing.expectEqual(expected.len, kinds.len);
    for (expected, kinds) |want, got| try testing.expectEqualStrings(want, got);
}

const TestEnv = struct {
    threaded: std.Io.Threaded,
    tmp: testing.TmpDir,
    dir: [:0]u8,

    fn init(gpa: Allocator) !TestEnv {
        var threaded: std.Io.Threaded = .init(gpa, .{});
        errdefer threaded.deinit();
        var tmp = testing.tmpDir(.{});
        errdefer tmp.cleanup();
        const dir = try testTmpPath(gpa, threaded.io(), &tmp);
        return .{ .threaded = threaded, .tmp = tmp, .dir = dir };
    }

    fn io(self: *TestEnv) std.Io {
        return self.threaded.io();
    }

    fn deinit(self: *TestEnv, gpa: Allocator) void {
        gpa.free(self.dir);
        self.tmp.cleanup();
        self.threaded.deinit();
        self.* = undefined;
    }
};

/// Deterministic run inputs for every test that does not probe the
/// default-context path.
const t_now = "2026-08-05T12:00:00Z";
const t_seed = "7";

fn failingNow(io: std.Io) ContextError!i64 {
    _ = io;
    return error.ContextUnavailable;
}
fn failingSeed(io: std.Io) ContextError!u64 {
    _ = io;
    return error.ContextUnavailable;
}

/// Trips signal flags before the record whose index is `at`.
const Trip = struct {
    at: usize,
    flags: []const *std.atomic.Value(bool),

    fn hook(state: ?*anyopaque, emitted: usize) void {
        const self: *Trip = @ptrCast(@alignCast(state.?));
        if (emitted >= self.at) {
            for (self.flags) |f| f.store(true, .release);
        }
    }
};

// ─── done-when 2: help + mandatory sheet selection ───────────────

test "eval --help lists every §12.2 flag" {
    const a = testing.allocator;
    var env = try TestEnv.init(a);
    defer env.deinit(a);

    var r = try drive(a, env.io(), &.{ "eval", "--help" }, .{});
    defer r.deinit(a);
    try testing.expectEqual(exit_ok, r.code);
    for ([_][]const u8{
        "--formula",  "--sheet",      "--name", "--anchor", "--dialect",
        "--now",      "--utc-offset", "--seed", "--mode",   "--profile",
        "--deadline",
    }) |flag| {
        try testing.expect(std.mem.indexOf(u8, r.out, flag) != null);
    }
}

test "recalc --help lists every §12.2 flag" {
    const a = testing.allocator;
    var env = try TestEnv.init(a);
    defer env.deinit(a);

    var r = try drive(a, env.io(), &.{ "recalc", "--help" }, .{});
    defer r.deinit(a);
    try testing.expectEqual(exit_ok, r.code);
    for ([_][]const u8{
        "--out",     "--now",            "--utc-offset", "--seed",     "--mode",
        "--profile", "--on-unsupported", "--report",     "--deadline",
    }) |flag| {
        try testing.expect(std.mem.indexOf(u8, r.out, flag) != null);
    }
}

test "eval without --sheet/--name exits 1; with both exits 1" {
    const a = testing.allocator;
    var env = try TestEnv.init(a);
    defer env.deinit(a);
    const path = try writeTestFixture(a, env.io(), env.dir, "v.xlsx", t_sheet_values);
    defer a.free(path);

    var r1 = try drive(a, env.io(), &.{ "eval", path, "--formula", "1+2" }, .{});
    defer r1.deinit(a);
    try testing.expectEqual(exit_usage, r1.code);
    try testing.expectEqual(@as(usize, 0), r1.out.len);

    var r2 = try drive(a, env.io(), &.{ "eval", path, "--formula", "1+2", "--sheet", "0", "--name", "Sheet1" }, .{});
    defer r2.deinit(a);
    try testing.expectEqual(exit_usage, r2.code);
}

// ─── done-when 3: the grammars, production by production ─────────

test "eval scalar: eval-header then eval-complete, every record versioned" {
    const a = testing.allocator;
    var env = try TestEnv.init(a);
    defer env.deinit(a);
    const path = try writeTestFixture(a, env.io(), env.dir, "v.xlsx", t_sheet_values);
    defer a.free(path);

    var r = try drive(a, env.io(), &.{ "eval", path, "--formula", "A1+B1", "--sheet", "0", "--now", t_now, "--seed", t_seed }, .{});
    defer r.deinit(a);
    try testing.expectEqual(exit_ok, r.code);
    try expectKinds(a, r.out, &.{ "eval-header", "eval-complete" });
    try testing.expect(std.mem.indexOf(u8, r.out, "\"value\":3") != null);
    try testing.expect(std.mem.indexOf(u8, r.out, "\"cells\":0") != null);
}

test "eval matrix: header, row-major cells, complete with the cell count" {
    const a = testing.allocator;
    var env = try TestEnv.init(a);
    defer env.deinit(a);
    const path = try writeTestFixture(a, env.io(), env.dir, "v.xlsx", t_sheet_values);
    defer a.free(path);

    var r = try drive(a, env.io(), &.{ "eval", path, "--formula", "A1:B2", "--sheet", "0", "--now", t_now, "--seed", t_seed }, .{});
    defer r.deinit(a);
    try testing.expectEqual(exit_ok, r.code);
    try expectKinds(a, r.out, &.{ "eval-header", "eval-cell", "eval-cell", "eval-cell", "eval-cell", "eval-complete" });
    try testing.expect(std.mem.indexOf(u8, r.out, "\"type\":\"matrix\",\"rows\":2,\"cols\":2") != null);
    try testing.expect(std.mem.indexOf(u8, r.out, "\"cells\":4") != null);
    // Row-major: r1c1=1, r1c2=2, r2c1=3, r2c2=4.
    try testing.expect(std.mem.indexOf(u8, r.out, "\"r\":1,\"c\":2,\"type\":\"number\",\"value\":2") != null);
    try testing.expect(std.mem.indexOf(u8, r.out, "\"r\":2,\"c\":1,\"type\":\"number\",\"value\":3") != null);
}

test "eval refusal-before-header: the refusal IS the stream" {
    const a = testing.allocator;
    var env = try TestEnv.init(a);
    defer env.deinit(a);
    const path = try writeTestFixture(a, env.io(), env.dir, "v.xlsx", t_sheet_values);
    defer a.free(path);

    var r = try drive(a, env.io(), &.{ "eval", path, "--formula", "NOTAFUNC(1)", "--sheet", "0", "--now", t_now, "--seed", t_seed }, .{});
    defer r.deinit(a);
    try testing.expectEqual(exit_refusal, r.code);
    try expectKinds(a, r.out, &.{"refusal"});
    try testing.expect(std.mem.indexOf(u8, r.out, "\"error\":\"FormulaUnsupportedFunction\"") != null);
    try testing.expect(std.mem.indexOf(u8, r.out, "\"truncated\":false") != null);
}

test "eval cancelled-before-header: SIGINT before any record, exit 130" {
    const a = testing.allocator;
    var env = try TestEnv.init(a);
    defer env.deinit(a);
    const path = try writeTestFixture(a, env.io(), env.dir, "v.xlsx", t_sheet_values);
    defer a.free(path);

    var stop = std.atomic.Value(bool).init(true);
    var int_flag = std.atomic.Value(bool).init(true);
    var r = try drive(a, env.io(), &.{ "eval", path, "--formula", "1+2", "--sheet", "0", "--now", t_now, "--seed", t_seed }, .{
        .sig = .{ .stop = &stop, .sigint = &int_flag },
    });
    defer r.deinit(a);
    try testing.expectEqual(exit_sigint, r.code);
    try expectKinds(a, r.out, &.{"cancelled"});
    try testing.expect(std.mem.indexOf(u8, r.out, "\"after\":\"none\"") != null);
}

test "eval cancelled mid-stream: completed records stand, cancelled closes, exit 143" {
    const a = testing.allocator;
    var env = try TestEnv.init(a);
    defer env.deinit(a);
    const path = try writeTestFixture(a, env.io(), env.dir, "v.xlsx", t_sheet_values);
    defer a.free(path);

    var stop = std.atomic.Value(bool).init(false);
    var term_flag = std.atomic.Value(bool).init(false);
    var trip: Trip = .{ .at = 2, .flags = &.{ &stop, &term_flag } };
    var r = try drive(a, env.io(), &.{ "eval", path, "--formula", "A1:B2", "--sheet", "0", "--now", t_now, "--seed", t_seed }, .{
        .sig = .{ .stop = &stop, .sigterm = &term_flag },
        .before_record = Trip.hook,
        .before_record_state = &trip,
    });
    defer r.deinit(a);
    try testing.expectEqual(exit_sigterm, r.code);
    try expectKinds(a, r.out, &.{ "eval-header", "eval-cell", "cancelled" });
    try testing.expect(std.mem.indexOf(u8, r.out, "\"after\":\"eval-cell\"") != null);
}

test "eval --deadline 0: cancelled-before-header, exit 3" {
    const a = testing.allocator;
    var env = try TestEnv.init(a);
    defer env.deinit(a);
    const path = try writeTestFixture(a, env.io(), env.dir, "v.xlsx", t_sheet_values);
    defer a.free(path);

    var r = try drive(a, env.io(), &.{ "eval", path, "--formula", "1+2", "--sheet", "0", "--deadline", "0", "--now", t_now, "--seed", t_seed }, .{});
    defer r.deinit(a);
    try testing.expectEqual(exit_refusal, r.code);
    try expectKinds(a, r.out, &.{"cancelled"});
}

test "recalc --report success: the report is the stream's terminal" {
    const a = testing.allocator;
    var env = try TestEnv.init(a);
    defer env.deinit(a);
    const path = try writeTestFixture(a, env.io(), env.dir, "in.xlsx", t_sheet_stale);
    defer a.free(path);
    const out_path = try std.fs.path.join(a, &.{ env.dir, "out.xlsx" });
    defer a.free(out_path);

    var r = try drive(a, env.io(), &.{ "recalc", path, "--out", out_path, "--report", "--now", t_now, "--seed", t_seed }, .{});
    defer r.deinit(a);
    try testing.expectEqual(exit_ok, r.code);
    try expectKinds(a, r.out, &.{"recalc-report"});
    try testing.expect(std.mem.indexOf(u8, r.out, "\"sheets\":1") != null);
    try testing.expect(std.mem.indexOf(u8, r.out, "\"cells\":1") != null);
    try testing.expect(std.mem.indexOf(u8, r.out, "\"keptStale\":false") != null);

    // The destination holds the recalced cache: B1 = A1+1 = 2.
    var wb = try zlsx_pkg.Workbook.open(a, env.io(), out_path);
    defer wb.deinit();
    const view = try (try wb.sheet(0)).ensureParsed();
    var found = false;
    for (view.rows) |row| for (row.cells) |c| {
        if (std.mem.eql(u8, c.ref, "B1")) {
            try testing.expectEqualStrings("2", c.raw_value orelse "");
            found = true;
        }
    };
    try testing.expect(found);
}

test "recalc without --report: stdout silent on every outcome" {
    const a = testing.allocator;
    var env = try TestEnv.init(a);
    defer env.deinit(a);
    const path = try writeTestFixture(a, env.io(), env.dir, "in.xlsx", t_sheet_stale);
    defer a.free(path);
    const out_path = try std.fs.path.join(a, &.{ env.dir, "out.xlsx" });
    defer a.free(out_path);

    var r = try drive(a, env.io(), &.{ "recalc", path, "--out", out_path, "--now", t_now, "--seed", t_seed }, .{});
    defer r.deinit(a);
    try testing.expectEqual(exit_ok, r.code);
    try testing.expectEqual(@as(usize, 0), r.out.len);
}

test "recalc --report refusal: refusal terminal, exit 3, destination untouched" {
    const a = testing.allocator;
    var env = try TestEnv.init(a);
    defer env.deinit(a);
    const path = try writeTestFixture(a, env.io(), env.dir, "in.xlsx", t_sheet_unsupported);
    defer a.free(path);
    const out_path = try std.fs.path.join(a, &.{ env.dir, "out.xlsx" });
    defer a.free(out_path);

    // Pre-existing destination bytes must survive the refusal.
    {
        var f = try std.Io.Dir.cwd().createFile(env.io(), out_path, .{});
        defer f.close(env.io());
        var wbuf: [16]u8 = undefined;
        var w = f.writer(env.io(), &wbuf);
        try w.interface.writeAll("prior bytes");
        try w.interface.flush();
    }

    var r = try drive(a, env.io(), &.{ "recalc", path, "--out", out_path, "--report", "--now", t_now, "--seed", t_seed }, .{});
    defer r.deinit(a);
    try testing.expectEqual(exit_refusal, r.code);
    try expectKinds(a, r.out, &.{"refusal"});
    try testing.expect(std.mem.indexOf(u8, r.out, "\"error\":\"FormulaUnsupportedFunction\"") != null);

    var f = try std.Io.Dir.cwd().openFile(env.io(), out_path, .{});
    defer f.close(env.io());
    var rbuf: [64]u8 = undefined;
    var fr = f.reader(env.io(), &rbuf);
    const n = try fr.interface.readSliceShort(&rbuf);
    try testing.expectEqualStrings("prior bytes", rbuf[0..n]);
}

test "recalc refusal with absent destination: no file appears" {
    const a = testing.allocator;
    var env = try TestEnv.init(a);
    defer env.deinit(a);
    const path = try writeTestFixture(a, env.io(), env.dir, "in.xlsx", t_sheet_unsupported);
    defer a.free(path);
    const out_path = try std.fs.path.join(a, &.{ env.dir, "never.xlsx" });
    defer a.free(out_path);

    var r = try drive(a, env.io(), &.{ "recalc", path, "--out", out_path, "--now", t_now, "--seed", t_seed }, .{});
    defer r.deinit(a);
    try testing.expectEqual(exit_refusal, r.code);
    try testing.expectError(error.FileNotFound, std.Io.Dir.cwd().openFile(env.io(), out_path, .{}));
}

test "recalc --on-unsupported keep-stale-and-mark: success with a census" {
    const a = testing.allocator;
    var env = try TestEnv.init(a);
    defer env.deinit(a);
    const path = try writeTestFixture(a, env.io(), env.dir, "in.xlsx", t_sheet_unsupported);
    defer a.free(path);
    const out_path = try std.fs.path.join(a, &.{ env.dir, "out.xlsx" });
    defer a.free(out_path);

    var r = try drive(a, env.io(), &.{ "recalc", path, "--out", out_path, "--report", "--on-unsupported", "keep-stale-and-mark", "--now", t_now, "--seed", t_seed }, .{});
    defer r.deinit(a);
    try testing.expectEqual(exit_ok, r.code);
    try expectKinds(a, r.out, &.{"recalc-report"});
    try testing.expect(std.mem.indexOf(u8, r.out, "\"keptStale\":true") != null);
    try testing.expect(std.mem.indexOf(u8, r.out, "\"error\":\"FormulaUnsupportedFunction\",\"cell\":\"Sheet1!B1\"") != null);
}

test "recalc --report cancelled-before-header: SIGTERM pre-run, exit 143" {
    const a = testing.allocator;
    var env = try TestEnv.init(a);
    defer env.deinit(a);
    const path = try writeTestFixture(a, env.io(), env.dir, "in.xlsx", t_sheet_stale);
    defer a.free(path);
    const out_path = try std.fs.path.join(a, &.{ env.dir, "out.xlsx" });
    defer a.free(out_path);

    var stop = std.atomic.Value(bool).init(true);
    var term_flag = std.atomic.Value(bool).init(true);
    var r = try drive(a, env.io(), &.{ "recalc", path, "--out", out_path, "--report", "--now", t_now, "--seed", t_seed }, .{
        .sig = .{ .stop = &stop, .sigterm = &term_flag },
    });
    defer r.deinit(a);
    try testing.expectEqual(exit_sigterm, r.code);
    try expectKinds(a, r.out, &.{"cancelled"});
    try testing.expectError(error.FileNotFound, std.Io.Dir.cwd().openFile(env.io(), out_path, .{}));
}

test "recalc --deadline 0: exit 3, destination untouched" {
    const a = testing.allocator;
    var env = try TestEnv.init(a);
    defer env.deinit(a);
    const path = try writeTestFixture(a, env.io(), env.dir, "in.xlsx", t_sheet_stale);
    defer a.free(path);
    const out_path = try std.fs.path.join(a, &.{ env.dir, "out.xlsx" });
    defer a.free(out_path);

    var r = try drive(a, env.io(), &.{ "recalc", path, "--out", out_path, "--report", "--deadline", "0", "--now", t_now, "--seed", t_seed }, .{});
    defer r.deinit(a);
    try testing.expectEqual(exit_refusal, r.code);
    try expectKinds(a, r.out, &.{"cancelled"});
    try testing.expectError(error.FileNotFound, std.Io.Dir.cwd().openFile(env.io(), out_path, .{}));
}

// ─── done-when 4: the exit table, row by row ─────────────────────

test "exit 2: missing file and non-workbook bytes" {
    const a = testing.allocator;
    var env = try TestEnv.init(a);
    defer env.deinit(a);

    var r1 = try drive(a, env.io(), &.{ "eval", "/nonexistent/x.xlsx", "--formula", "1", "--sheet", "0", "--now", t_now, "--seed", t_seed }, .{});
    defer r1.deinit(a);
    try testing.expectEqual(exit_open, r1.code);

    const garbage = try std.fs.path.join(a, &.{ env.dir, "garbage.xlsx" });
    defer a.free(garbage);
    {
        var f = try std.Io.Dir.cwd().createFile(env.io(), garbage, .{});
        defer f.close(env.io());
        var wbuf: [32]u8 = undefined;
        var w = f.writer(env.io(), &wbuf);
        try w.interface.writeAll("this is not a zip archive");
        try w.interface.flush();
    }
    var r2 = try drive(a, env.io(), &.{ "recalc", garbage, "--out", "unused-out.xlsx", "--now", t_now, "--seed", t_seed }, .{});
    defer r2.deinit(a);
    try testing.expectEqual(exit_open, r2.code);
}

test "exit 3: FormulaLimitExceeded is a refusal at the CLI layer too" {
    const a = testing.allocator;
    var env = try TestEnv.init(a);
    defer env.deinit(a);
    const path = try writeTestFixture(a, env.io(), env.dir, "v.xlsx", t_sheet_values);
    defer a.free(path);

    // 300 nested groups beat max_parse_depth (256) without touching
    // the 8192-char ceiling — the refusal must name the limit plane.
    var formula: std.ArrayList(u8) = .empty;
    defer formula.deinit(a);
    for (0..300) |_| try formula.append(a, '(');
    try formula.append(a, '1');
    for (0..300) |_| try formula.append(a, ')');

    var r = try drive(a, env.io(), &.{ "eval", path, "--formula", formula.items, "--sheet", "0", "--now", t_now, "--seed", t_seed }, .{});
    defer r.deinit(a);
    try testing.expectEqual(exit_refusal, r.code);
    try expectKinds(a, r.out, &.{"refusal"});
    try testing.expect(std.mem.indexOf(u8, r.out, "\"error\":\"FormulaLimitExceeded\"") != null);
}

test "exit 4: genuine OOM is 4, and only OOM is 4" {
    const a = testing.allocator;
    var env = try TestEnv.init(a);
    defer env.deinit(a);
    const path = try writeTestFixture(a, env.io(), env.dir, "v.xlsx", t_sheet_values);
    defer a.free(path);

    var failing = testing.FailingAllocator.init(a, .{ .fail_index = 10 });
    var r = try drive(failing.allocator(), env.io(), &.{ "eval", path, "--formula", "1+2", "--sheet", "0", "--now", t_now, "--seed", t_seed }, .{});
    defer r.deinit(failing.allocator());
    try testing.expectEqual(exit_oom, r.code);
}

test "exit 5: stdout write failure without SIGPIPE" {
    const a = testing.allocator;
    var env = try TestEnv.init(a);
    defer env.deinit(a);
    const path = try writeTestFixture(a, env.io(), env.dir, "v.xlsx", t_sheet_values);
    defer a.free(path);

    var tiny: [8]u8 = undefined;
    var out_w = std.Io.Writer.fixed(&tiny);
    var err_aw: std.Io.Writer.Allocating = .init(a);
    defer err_aw.deinit();
    const code = run(a, env.io(), &.{ "eval", path, "--formula", "1+2", "--sheet", "0", "--now", t_now, "--seed", t_seed }, &out_w, &err_aw.writer, .{});
    try testing.expectEqual(exit_write, code);
}

test "exit 6: an unreadable clock or random source, never reported as OOM" {
    const a = testing.allocator;
    var env = try TestEnv.init(a);
    defer env.deinit(a);
    const path = try writeTestFixture(a, env.io(), env.dir, "v.xlsx", t_sheet_values);
    defer a.free(path);

    // Omitted --now, broken clock.
    var r1 = try drive(a, env.io(), &.{ "eval", path, "--formula", "1", "--sheet", "0", "--seed", t_seed }, .{
        .ctx = .{ .now_ms = failingNow },
    });
    defer r1.deinit(a);
    try testing.expectEqual(exit_context, r1.code);
    try testing.expectEqual(@as(usize, 0), r1.out.len);

    // Omitted --seed, broken random source — same row for recalc.
    var r2 = try drive(a, env.io(), &.{ "recalc", path, "--out", "unused.xlsx", "--now", t_now }, .{
        .ctx = .{ .seed = failingSeed },
    });
    defer r2.deinit(a);
    try testing.expectEqual(exit_context, r2.code);

    // Both stated: the broken sources are never consulted.
    var r3 = try drive(a, env.io(), &.{ "eval", path, "--formula", "1", "--sheet", "0", "--now", t_now, "--seed", t_seed }, .{
        .ctx = .{ .now_ms = failingNow, .seed = failingSeed },
    });
    defer r3.deinit(a);
    try testing.expectEqual(exit_ok, r3.code);
}

// ─── done-when 5: SIGPIPE vs every other truncated stream ────────

test "SIGPIPE: prefix-valid, no terminal record, exit 0" {
    const a = testing.allocator;
    var env = try TestEnv.init(a);
    defer env.deinit(a);
    const path = try writeTestFixture(a, env.io(), env.dir, "v.xlsx", t_sheet_values);
    defer a.free(path);

    var stop = std.atomic.Value(bool).init(false);
    var pipe_flag = std.atomic.Value(bool).init(false);
    var trip: Trip = .{ .at = 2, .flags = &.{ &stop, &pipe_flag } };
    var r = try drive(a, env.io(), &.{ "eval", path, "--formula", "A1:B2", "--sheet", "0", "--now", t_now, "--seed", t_seed }, .{
        .sig = .{ .stop = &stop, .sigpipe = &pipe_flag },
        .before_record = Trip.hook,
        .before_record_state = &trip,
    });
    defer r.deinit(a);
    try testing.expectEqual(exit_ok, r.code);
    // Prefix-valid and terminal-free: header and one cell, nothing else.
    try expectKinds(a, r.out, &.{ "eval-header", "eval-cell" });
}

test "abnormal EOF is distinguishable from SIGPIPE by exit code" {
    const a = testing.allocator;
    var env = try TestEnv.init(a);
    defer env.deinit(a);
    const path = try writeTestFixture(a, env.io(), env.dir, "v.xlsx", t_sheet_values);
    defer a.free(path);

    // The same truncated shape — records then silence — but from a
    // write failure with no SIGPIPE: exit 5, not 0. A consumer that
    // sees no terminal record asks the exit code which case it was.
    var buf: [220]u8 = undefined;
    var out_w = std.Io.Writer.fixed(&buf);
    var err_aw: std.Io.Writer.Allocating = .init(a);
    defer err_aw.deinit();
    const code = run(a, env.io(), &.{ "eval", path, "--formula", "A1:B2", "--sheet", "0", "--now", t_now, "--seed", t_seed }, &out_w, &err_aw.writer, .{});
    try testing.expectEqual(exit_write, code);
}

// ─── done-when 6: the commit seam ────────────────────────────────

/// The commit hook a production `saveWithRecalc` wires is the swap;
/// this one is the swap PLUS the signal — the injection §12.2's
/// commit-aware mapping is proven by. It runs where every commit hook
/// runs: between the rename and the directory fsync (§5.7.9).
const SeamInjection = struct {
    wb: *zlsx_pkg.Workbook,
    candidate: *zlsx_pkg.recalc_txn.Candidate,
    stop: *std.atomic.Value(bool),
    term_flag: *std.atomic.Value(bool),
    fired: bool = false,

    fn call(ctx: ?*anyopaque) void {
        const self: *SeamInjection = @ptrCast(@alignCast(ctx.?));
        self.candidate.swap(self.wb);
        // The signal, delivered after the rename and before process
        // exit — deterministically, not by timing.
        self.stop.store(true, .release);
        self.term_flag.store(true, .release);
        self.fired = true;
    }
};

test "a signal at the commit seam exits 0 and the destination holds the recalc" {
    const a = testing.allocator;
    var env = try TestEnv.init(a);
    defer env.deinit(a);
    const path = try writeTestFixture(a, env.io(), env.dir, "in.xlsx", t_sheet_stale);
    defer a.free(path);
    const out_path = try std.fs.path.join(a, &.{ env.dir, "out.xlsx" });
    defer a.free(out_path);

    var wb = try zlsx_pkg.Workbook.open(a, env.io(), path);
    defer wb.deinit();

    var stop = std.atomic.Value(bool).init(false);
    var term_flag = std.atomic.Value(bool).init(false);
    const run_inputs: zlsx_pkg.RunInputs = .{
        .now_utc_ms = 1_700_000_000_000,
        .rng_seed = 7,
        .limits = .{},
        .cancel = .{ .atomic = &stop },
    };

    var prepared = try zlsx_pkg.recalc_run.prepare(&wb, a, env.io(), run_inputs, .{});
    switch (prepared) {
        .ok => |*candidate| {
            var seam: SeamInjection = .{
                .wb = &wb,
                .candidate = candidate,
                .stop = &stop,
                .term_flag = &term_flag,
            };
            var watch: zlsx_control.Watch = .init(env.io(), .{ .cancel = .{ .atomic = &stop } });
            _ = try candidate.next.saveCommitted(env.io(), out_path, watch.poller(), .{
                .ctx = &seam,
                .call = SeamInjection.call,
            });
            try testing.expect(seam.fired);
            var report = candidate.takeReport();
            defer report.deinit(a);
        },
        else => return error.TestUnexpectedResult,
    }

    // The signal is now pending the way a real SIGTERM between rename
    // and exit is pending. Commit-aware mapping: success.
    const sig: SignalState = .{ .stop = &stop, .sigterm = &term_flag };
    try testing.expect(sig.termFired());
    try testing.expectEqual(exit_ok, mapRecalcSignalExit(true, sig));
    // And the same flags before the commit would have been 143 — the
    // distinction is the commit, not the flags.
    try testing.expectEqual(exit_sigterm, mapRecalcSignalExit(false, sig));

    // The destination holds the recalced cache: B1 = 2.
    var out_wb = try zlsx_pkg.Workbook.open(a, env.io(), out_path);
    defer out_wb.deinit();
    const view = try (try out_wb.sheet(0)).ensureParsed();
    var found = false;
    for (view.rows) |row| for (row.cells) |c| {
        if (std.mem.eql(u8, c.ref, "B1")) {
            try testing.expectEqualStrings("2", c.raw_value orelse "");
            found = true;
        }
    };
    try testing.expect(found);
}

test "signals that fire after a committed run do not re-map its exit" {
    const a = testing.allocator;
    var env = try TestEnv.init(a);
    defer env.deinit(a);
    const path = try writeTestFixture(a, env.io(), env.dir, "in.xlsx", t_sheet_stale);
    defer a.free(path);
    const out_path = try std.fs.path.join(a, &.{ env.dir, "out.xlsx" });
    defer a.free(out_path);

    // The full command path with signal flags raised after the
    // transaction returned: still 0. (The seam test above pins the
    // stronger in-transaction case; this pins the command wiring.)
    var stop = std.atomic.Value(bool).init(false);
    var int_flag = std.atomic.Value(bool).init(false);
    var r = try drive(a, env.io(), &.{ "recalc", path, "--out", out_path, "--now", t_now, "--seed", t_seed }, .{
        .sig = .{ .stop = &stop, .sigint = &int_flag },
    });
    defer r.deinit(a);
    try testing.expectEqual(exit_ok, r.code);
    stop.store(true, .release);
    int_flag.store(true, .release);
    try testing.expectEqual(exit_ok, mapRecalcSignalExit(true, .{ .stop = &stop, .sigint = &int_flag }));
}

// ─── done-when 7: --out identity ─────────────────────────────────

test "--out identity is refused before any mutation" {
    const a = testing.allocator;
    var env = try TestEnv.init(a);
    defer env.deinit(a);
    const path = try writeTestFixture(a, env.io(), env.dir, "in.xlsx", t_sheet_stale);
    defer a.free(path);

    const before = try readFileAlloc(a, env.io(), path);
    defer a.free(before);

    // Literal identity.
    var r1 = try drive(a, env.io(), &.{ "recalc", path, "--out", path, "--now", t_now, "--seed", t_seed }, .{});
    defer r1.deinit(a);
    try testing.expectEqual(exit_usage, r1.code);

    // Aliased spelling of the same file.
    const dotted = try std.fs.path.join(a, &.{ env.dir, ".", "in.xlsx" });
    defer a.free(dotted);
    var r2 = try drive(a, env.io(), &.{ "recalc", path, "--out", dotted, "--now", t_now, "--seed", t_seed }, .{});
    defer r2.deinit(a);
    try testing.expectEqual(exit_usage, r2.code);

    const after = try readFileAlloc(a, env.io(), path);
    defer a.free(after);
    try testing.expectEqualSlices(u8, before, after);
}

fn readFileAlloc(gpa: Allocator, io: std.Io, path: []const u8) ![]u8 {
    var f = try std.Io.Dir.cwd().openFile(io, path, .{});
    defer f.close(io);
    const size = (try f.stat(io)).size;
    const buf = try gpa.alloc(u8, @intCast(size));
    errdefer gpa.free(buf);
    var r = f.reader(io, &.{});
    try r.interface.readSliceAll(buf);
    return buf;
}

// ─── done-when 8: blank never crosses the boundary ───────────────

test "=A1 on an empty A1 emits 0 and the word blank appears nowhere" {
    const a = testing.allocator;
    var env = try TestEnv.init(a);
    defer env.deinit(a);
    const path = try writeTestFixture(a, env.io(), env.dir, "e.xlsx", t_sheet_empty_a1);
    defer a.free(path);

    var r = try drive(a, env.io(), &.{ "eval", path, "--formula", "=A1", "--sheet", "0", "--now", t_now, "--seed", t_seed }, .{});
    defer r.deinit(a);
    try testing.expectEqual(exit_ok, r.code);
    try testing.expect(std.mem.indexOf(u8, r.out, "\"type\":\"number\",\"value\":0") != null);
    try testing.expect(std.mem.indexOf(u8, r.out, "blank") == null);

    // The recalc stream keeps the same vocabulary.
    const stale = try writeTestFixture(a, env.io(), env.dir, "s.xlsx", t_sheet_stale);
    defer a.free(stale);
    const out_path = try std.fs.path.join(a, &.{ env.dir, "out.xlsx" });
    defer a.free(out_path);
    var r2 = try drive(a, env.io(), &.{ "recalc", stale, "--out", out_path, "--report", "--now", t_now, "--seed", t_seed }, .{});
    defer r2.deinit(a);
    try testing.expectEqual(exit_ok, r2.code);
    try testing.expect(std.mem.indexOf(u8, r2.out, "blank") == null);
}

// ─── done-when 9: seed as a decimal string ───────────────────────

test "seed above 2^53 round-trips as a decimal string" {
    const a = testing.allocator;
    var env = try TestEnv.init(a);
    defer env.deinit(a);
    const path = try writeTestFixture(a, env.io(), env.dir, "v.xlsx", t_sheet_values);
    defer a.free(path);

    // 2^53 + 3: a JSON number would silently round it.
    const big = "9007199254740995";
    var r = try drive(a, env.io(), &.{ "eval", path, "--formula", "1", "--sheet", "0", "--now", t_now, "--seed", big }, .{});
    defer r.deinit(a);
    try testing.expectEqual(exit_ok, r.code);

    var it = std.mem.splitScalar(u8, r.out, '\n');
    const header = it.next().?;
    var parsed = try std.json.parseFromSlice(std.json.Value, a, header, .{});
    defer parsed.deinit();
    const resolved_obj = parsed.value.object.get("resolved").?.object;
    const seed_val = resolved_obj.get("seed").?;
    try testing.expect(seed_val == .string);
    try testing.expectEqualStrings(big, seed_val.string);
    try testing.expectEqual(@as(u64, 9007199254740995), try std.fmt.parseInt(u64, seed_val.string, 10));
}

// ─── unit coverage: ISO 8601 both ways, column letters ───────────

test "parseIso8601UtcMs: epoch, offsets, and rejects" {
    try testing.expectEqual(@as(i64, 0), try parseIso8601UtcMs("1970-01-01"));
    try testing.expectEqual(@as(i64, std.time.ms_per_day), try parseIso8601UtcMs("1970-01-02T00:00:00Z"));
    try testing.expectEqual(
        try parseIso8601UtcMs("2026-08-05T10:00:00Z"),
        try parseIso8601UtcMs("2026-08-05T12:00:00+02:00"),
    );
    try testing.expectEqual(
        try parseIso8601UtcMs("2026-08-05T12:30:00Z"),
        try parseIso8601UtcMs("2026-08-05T07:00:00-05:30"),
    );
    try testing.expectEqual(@as(i64, 1500), try parseIso8601UtcMs("1970-01-01T00:00:01.5"));
    try testing.expectError(error.BadIso8601, parseIso8601UtcMs("2026-13-01"));
    try testing.expectError(error.BadIso8601, parseIso8601UtcMs("2026-02-30"));
    try testing.expectError(error.BadIso8601, parseIso8601UtcMs("20260805"));
    try testing.expectError(error.BadIso8601, parseIso8601UtcMs("2026-08-05T25:00"));
    try testing.expectError(error.BadIso8601, parseIso8601UtcMs("2026-08-05T12:00:00+2:00"));
}

test "writeIsoUtc round-trips through the parser" {
    var buf: [64]u8 = undefined;
    var w = std.Io.Writer.fixed(&buf);
    const ms = try parseIso8601UtcMs("2026-08-05T12:34:56.789Z");
    try writeIsoUtc(&w, ms);
    try testing.expectEqualStrings("2026-08-05T12:34:56.789Z", w.buffered());
    try testing.expectEqual(ms, try parseIso8601UtcMs(w.buffered()));
}

test "census cell labels go through M0's one formatter" {
    // `refs.zig` owns the base-26 tests; here we only pin the 1-based
    // crossing this file performs for census columns.
    var buf: [coords.max_col_letters]u8 = undefined;
    var n = try coords.writeColNumberLetters(&buf, 0 + 1);
    try testing.expectEqualStrings("A", buf[0..n]);
    n = try coords.writeColNumberLetters(&buf, 702 + 1);
    try testing.expectEqualStrings("AAA", buf[0..n]);
}

test "the stream machine can produce every remaining production shape" {
    // Two productions no v1 command path reaches — a refusal AFTER a
    // header (the result is computed before the header goes out) and a
    // diagnostic inside an eval stream — are still normative grammar,
    // and a consumer may rely on parsing them. Drive the emitter
    // directly so their shapes are pinned where they are defined.
    const a = testing.allocator;
    var aw: std.Io.Writer.Allocating = .init(a);
    defer aw.deinit();

    const deps: Deps = .{};
    var em: Emitter = .{ .w = &aw.writer, .sig = .{}, .deps = &deps };
    try em.beginData();
    try em.head(.@"eval-header");
    try em.w.print(",\"type\":\"number\",\"value\":1}}", .{});
    try em.finishRecord(.@"eval-header");
    try em.diagnostic("note", "a diagnostic between cells and the terminal");
    try em.refusal("FormulaCycle", "Sheet1!A1");

    try expectKinds(a, aw.writer.buffered(), &.{ "eval-header", "diagnostic", "refusal" });
    try testing.expect(std.mem.indexOf(u8, aw.writer.buffered(), "\"cells\":[\"Sheet1!A1\"]") != null);
}

test "==1 refuses (the CLI must not stack a second leading-= strip)" {
    const a = testing.allocator;
    var env = try TestEnv.init(a);
    defer env.deinit(a);
    const path = try writeTestFixture(a, env.io(), env.dir, "v.xlsx", t_sheet_values);
    defer a.free(path);

    var r = try drive(a, env.io(), &.{ "eval", path, "--formula", "==1", "--sheet", "0", "--now", t_now, "--seed", t_seed }, .{});
    defer r.deinit(a);
    try testing.expectEqual(exit_refusal, r.code);
    try expectKinds(a, r.out, &.{"refusal"});
    try testing.expect(std.mem.indexOf(u8, r.out, "\"error\":\"FormulaMalformedInput\"") != null);
}
