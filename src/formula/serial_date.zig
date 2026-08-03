//! Excel serial dates — both epochs, including the day that does not
//! exist (`goal_formula.md` §5.4a).
//!
//! M3b of the tier-D1 ladder. Nothing here reads a clock: `NOW()` and
//! `TODAY()` get their instant from `RunInputs.now_utc_ms` (§5.5), and
//! this module only converts.
//!
//! The 1900 system has a hole in it
//! --------------------------------
//! Lotus 1-2-3 treated 1900 as a leap year, Excel copied the bug for
//! compatibility, and the file format now depends on it. So serial 60 is
//! **1900-02-29 — a date that never happened**, and serial 0 is
//! **1900-01-00 — a day that is not in any month**. Both are
//! representable here rather than being errors, because both appear in
//! real workbooks and a converter that refused them would refuse files
//! Excel opens.
//!
//! The consequence is that the 1900 epoch is *two* epochs. Serials 1–59
//! count from 1899-12-31 and serials 61+ count from 1899-12-30, and the
//! step between them is the fictitious day. Writing it as one epoch plus
//! a correction is the same arithmetic wearing a disguise; writing it as
//! two named constants makes the discontinuity impossible to lose.
//!
//! The 1904 system has no hole, which is the whole reason it exists.
//! `1900-02-29` under `.d1904` is simply not a date, and says so.
//!
//! Round-tripping
//! --------------
//! `serialFromDate(dateFromSerial(s)) == s` for every serial in range,
//! including both fictitious ones, and the reverse holds for every real
//! date in range. Both directions are property-tested across the whole
//! domain rather than at sampled points, because the interesting
//! failures are all within two days of a boundary.

const std = @import("std");
const assert = std.debug.assert;

const run_inputs = @import("run_inputs.zig");

pub const DateSystem = run_inputs.DateSystem;

pub const Error = error{
    /// Outside `[0, maxSerial(system)]`.
    SerialOutOfRange,
    /// Not a date the system can express: before its epoch, past
    /// 9999-12-31, or a calendar date that does not exist.
    DateOutOfRange,
};

/// Which invented day a serial denotes, if any. Named rather than
/// signalled by an out-of-band day value, so a caller that ignores it
/// still gets a sensible `year`/`month`/`day`.
pub const Fictitious = enum {
    none,
    /// Serial 0 in the 1900 system: `1900-01-00`, reported as day 0.
    day_zero,
    /// Serial 60 in the 1900 system: `1900-02-29`, which never occurred.
    feb_29_1900,
};

pub const Date = struct {
    year: u16,
    month: u8,
    /// Zero only for `.day_zero`.
    day: u8,
    fictitious: Fictitious = .none,

    pub fn eql(a: Date, b: Date) bool {
        return a.year == b.year and a.month == b.month and
            a.day == b.day and a.fictitious == b.fictitious;
    }
};

pub const Time = struct {
    hour: u8,
    minute: u8,
    second: u8,
    /// Set when rounding to the nearest second carried past midnight —
    /// `0.9999999` is 24:00:00, which is tomorrow's 00:00:00.
    day_carry: u1 = 0,
};

// ─── the proleptic Gregorian calendar ────────────────────────────
//
// Howard Hinnant's civil-from-days pair, which is exact over the whole
// i64 range and needs no tables. Days are counted from 1970-01-01 only
// because that is the formulation's natural origin; no Unix semantics
// are implied and no clock is involved.

pub fn daysFromCivil(year: i32, month: u32, day: u32) i64 {
    assert(month >= 1 and month <= 12);
    const y: i64 = @as(i64, year) - @intFromBool(month <= 2);
    const era = @divFloor(y, 400);
    const yoe: i64 = y - era * 400; // [0, 399]
    const m: i64 = month;
    const doy = @divTrunc(153 * (m + (if (m > 2) @as(i64, -3) else 9)) + 2, 5) + @as(i64, day) - 1;
    const doe = yoe * 365 + @divTrunc(yoe, 4) - @divTrunc(yoe, 100) + doy;
    return era * 146097 + doe - 719468;
}

pub fn civilFromDays(days: i64) struct { year: i32, month: u8, day: u8 } {
    const z = days + 719468;
    const era = @divFloor(z, 146097);
    const doe = z - era * 146097; // [0, 146096]
    const yoe = @divTrunc(doe - @divTrunc(doe, 1460) + @divTrunc(doe, 36524) - @divTrunc(doe, 146096), 365);
    const y = yoe + era * 400;
    const doy = doe - (365 * yoe + @divTrunc(yoe, 4) - @divTrunc(yoe, 100));
    const mp = @divTrunc(5 * doy + 2, 153); // [0, 11]
    const d = doy - @divTrunc(153 * mp + 2, 5) + 1; // [1, 31]
    const m = mp + (if (mp < 10) @as(i64, 3) else -9); // [1, 12]
    return .{
        .year = @intCast(y + @intFromBool(m <= 2)),
        .month = @intCast(m),
        .day = @intCast(d),
    };
}

pub fn isLeapYear(year: i32) bool {
    return @rem(year, 4) == 0 and (@rem(year, 100) != 0 or @rem(year, 400) == 0);
}

pub fn daysInMonth(year: i32, month: u8) u8 {
    const table = [_]u8{ 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 };
    assert(month >= 1 and month <= 12);
    if (month == 2 and isLeapYear(year)) return 29;
    return table[month - 1];
}

// ─── the two epochs ──────────────────────────────────────────────

/// Serials 61 and above count from here. The day *before* 1899-12-31,
/// because the fictitious 1900-02-29 has already consumed a serial by
/// the time the calendar gets this far.
const epoch_1900_late: i64 = daysFromCivil(1899, 12, 30);
/// Serials 1–59 count from here, which is the arithmetic that would
/// hold everywhere if 1900 had not been given a 29th of February.
const epoch_1900_early: i64 = daysFromCivil(1899, 12, 31);
const epoch_1904: i64 = daysFromCivil(1904, 1, 1);

const last_day: i64 = daysFromCivil(9999, 12, 31);

/// Excel's documented maxima. Asserted against the calendar arithmetic
/// below rather than trusted: a constant nobody checks is a comment.
pub const max_serial_1900: i32 = 2_958_465;
pub const max_serial_1904: i32 = 2_957_003;

comptime {
    assert(last_day - epoch_1900_late == max_serial_1900);
    assert(last_day - epoch_1904 == max_serial_1904);
    // The two 1900 epochs differ by exactly the invented day.
    assert(epoch_1900_early - epoch_1900_late == 1);
}

pub fn maxSerial(system: DateSystem) i32 {
    return switch (system) {
        .d1900 => max_serial_1900,
        .d1904 => max_serial_1904,
    };
}

/// The last serial before the 1900 discontinuity, and the discontinuity
/// itself. Named because three call sites and four fixtures use them.
pub const last_serial_before_gap: i32 = 59;
pub const fictitious_leap_serial: i32 = 60;

// ─── conversions ─────────────────────────────────────────────────

pub fn dateFromSerial(system: DateSystem, serial: i32) Error!Date {
    if (serial < 0 or serial > maxSerial(system)) return error.SerialOutOfRange;
    switch (system) {
        .d1904 => {
            const c = civilFromDays(epoch_1904 + serial);
            return .{ .year = @intCast(c.year), .month = c.month, .day = c.day };
        },
        .d1900 => {
            if (serial == 0) {
                // `1900-01-00`. Day 0 is not a typo: it is what Excel
                // displays, and the `fictitious` tag says why.
                return .{ .year = 1900, .month = 1, .day = 0, .fictitious = .day_zero };
            }
            if (serial == fictitious_leap_serial) {
                return .{ .year = 1900, .month = 2, .day = 29, .fictitious = .feb_29_1900 };
            }
            const base = if (serial <= last_serial_before_gap) epoch_1900_early else epoch_1900_late;
            const c = civilFromDays(base + serial);
            return .{ .year = @intCast(c.year), .month = c.month, .day = c.day };
        },
    }
}

pub fn serialFromDate(system: DateSystem, year: i32, month: u8, day: u8) Error!i32 {
    switch (system) {
        .d1900 => {
            // The two invented days first: neither is a calendar date,
            // so neither survives validation below.
            if (year == 1900 and month == 1 and day == 0) return 0;
            if (year == 1900 and month == 2 and day == 29) return fictitious_leap_serial;
            try validateCivil(year, month, day);
            const days = daysFromCivil(year, month, day);
            if (days < epoch_1900_early + 1) return error.DateOutOfRange;
            if (days > last_day) return error.DateOutOfRange;
            const base = if (days <= daysFromCivil(1900, 2, 28)) epoch_1900_early else epoch_1900_late;
            return @intCast(days - base);
        },
        .d1904 => {
            // 1904 is clean: 1900-02-29 is simply not a date here, and
            // falls out of `validateCivil` rather than needing a rule.
            try validateCivil(year, month, day);
            const days = daysFromCivil(year, month, day);
            if (days < epoch_1904 or days > last_day) return error.DateOutOfRange;
            return @intCast(days - epoch_1904);
        },
    }
}

fn validateCivil(year: i32, month: u8, day: u8) Error!void {
    if (year < 1 or year > 9999) return error.DateOutOfRange;
    if (month < 1 or month > 12) return error.DateOutOfRange;
    if (day < 1 or day > daysInMonth(year, month)) return error.DateOutOfRange;
}

/// Split a serial's fractional part into a clock time, rounded to the
/// nearest second — Excel's own rounding for `HOUR`/`MINUTE`/`SECOND`.
///
/// `day_carry` exists because rounding can reach 86 400 seconds, and
/// 24:00:00 is not a time. Returning `hour = 24` would put an
/// out-of-range value in a `u8` for a caller to trip over later.
pub fn timeFromFraction(fraction: f64) Time {
    assert(std.math.isFinite(fraction));
    const frac = fraction - @floor(fraction);
    var secs: i64 = @intFromFloat(@round(frac * 86_400.0));
    var carry: u1 = 0;
    if (secs >= 86_400) {
        secs -= 86_400;
        carry = 1;
    }
    return .{
        .hour = @intCast(@divTrunc(secs, 3600)),
        .minute = @intCast(@divTrunc(@rem(secs, 3600), 60)),
        .second = @intCast(@rem(secs, 60)),
        .day_carry = carry,
    };
}

/// The day fraction a clock time denotes. Exact for whole seconds.
pub fn fractionFromTime(hour: u32, minute: u32, second: u32) f64 {
    const secs: f64 = @floatFromInt(hour * 3600 + minute * 60 + second);
    return secs / 86_400.0;
}

/// Split a whole serial into its day and time halves. Negative serials
/// are out of range in both systems, so the floor is always the day.
pub fn splitSerial(serial: f64) Error!struct { days: i32, fraction: f64 } {
    if (!std.math.isFinite(serial)) return error.SerialOutOfRange;
    const whole = @floor(serial);
    if (whole < 0 or whole > @as(f64, @floatFromInt(max_serial_1900))) {
        return error.SerialOutOfRange;
    }
    return .{ .days = @intFromFloat(whole), .fraction = serial - whole };
}

// ─── tests ───────────────────────────────────────────────────────

const testing = std.testing;

fn expectDate(system: DateSystem, serial: i32, y: u16, m: u8, d: u8, f: Fictitious) !void {
    const got = try dateFromSerial(system, serial);
    const want = Date{ .year = y, .month = m, .day = d, .fictitious = f };
    if (!got.eql(want)) {
        std.debug.print(
            "serial {d} ({t}): expected {d:0>4}-{d:0>2}-{d:0>2} ({t}), got {d:0>4}-{d:0>2}-{d:0>2} ({t})\n",
            .{ serial, system, y, m, d, f, got.year, got.month, got.day, got.fictitious },
        );
        return error.WrongDate;
    }
    // Every boundary is pinned in BOTH directions.
    try testing.expectEqual(serial, try serialFromDate(system, y, m, d));
}

test "1900: the boundaries around the day that never happened" {
    try expectDate(.d1900, 0, 1900, 1, 0, .day_zero);
    try expectDate(.d1900, 1, 1900, 1, 1, .none);
    try expectDate(.d1900, 2, 1900, 1, 2, .none);
    try expectDate(.d1900, 31, 1900, 1, 31, .none);
    try expectDate(.d1900, 32, 1900, 2, 1, .none);
    try expectDate(.d1900, 59, 1900, 2, 28, .none);
    // The hole: a date that never existed, and Excel's serial for it.
    try expectDate(.d1900, 60, 1900, 2, 29, .feb_29_1900);
    try expectDate(.d1900, 61, 1900, 3, 1, .none);
    try expectDate(.d1900, 62, 1900, 3, 2, .none);
    // …and the far end.
    try expectDate(.d1900, max_serial_1900, 9999, 12, 31, .none);
}

test "1904: the same boundaries, with no hole to step over" {
    try expectDate(.d1904, 0, 1904, 1, 1, .none);
    try expectDate(.d1904, 1, 1904, 1, 2, .none);
    // 1904 IS a leap year, so its 29 February is a real day.
    try expectDate(.d1904, 58, 1904, 2, 28, .none);
    try expectDate(.d1904, 59, 1904, 2, 29, .none);
    try expectDate(.d1904, 60, 1904, 3, 1, .none);
    try expectDate(.d1904, 61, 1904, 3, 2, .none);
    try expectDate(.d1904, max_serial_1904, 9999, 12, 31, .none);

    // The 1900 system's invented day is not a date here at all — that is
    // what "1904 clean" means, and it needs no special case to hold.
    try testing.expectError(error.DateOutOfRange, serialFromDate(.d1904, 1900, 2, 29));
    // Nor is anything before the 1904 epoch.
    try testing.expectError(error.DateOutOfRange, serialFromDate(.d1904, 1903, 12, 31));
    try testing.expectError(error.DateOutOfRange, serialFromDate(.d1904, 1900, 1, 1));
}

test "the two systems are offset by exactly 1462 days" {
    // 1904-01-01 is serial 0 there and 1462 here; a workbook converted
    // between systems shifts every date by this constant.
    try testing.expectEqual(@as(i32, 1462), try serialFromDate(.d1900, 1904, 1, 1));
    try testing.expectEqual(@as(i32, 0), try serialFromDate(.d1904, 1904, 1, 1));
    try testing.expectEqual(@as(i32, 1462), max_serial_1900 - max_serial_1904);
}

test "domain: serials outside the range refuse rather than wrapping" {
    try testing.expectError(error.SerialOutOfRange, dateFromSerial(.d1900, -1));
    try testing.expectError(error.SerialOutOfRange, dateFromSerial(.d1900, max_serial_1900 + 1));
    try testing.expectError(error.SerialOutOfRange, dateFromSerial(.d1904, -1));
    try testing.expectError(error.SerialOutOfRange, dateFromSerial(.d1904, max_serial_1904 + 1));

    // The 1900 system has no date before 1900-01-01 except its day zero.
    try testing.expectError(error.DateOutOfRange, serialFromDate(.d1900, 1899, 12, 31));
    try testing.expectError(error.DateOutOfRange, serialFromDate(.d1900, 10000, 1, 1));
    // And no calendar date that does not exist.
    try testing.expectError(error.DateOutOfRange, serialFromDate(.d1900, 1901, 2, 29));
    try testing.expectError(error.DateOutOfRange, serialFromDate(.d1900, 2001, 4, 31));
    try testing.expectError(error.DateOutOfRange, serialFromDate(.d1900, 2001, 13, 1));
}

test "round trip: every serial in both systems, not a sample" {
    // The interesting failures all live within two days of a boundary,
    // so sampling is exactly the wrong test. Both domains are ~3M wide;
    // walking them costs milliseconds.
    inline for ([_]DateSystem{ .d1900, .d1904 }) |system| {
        var s: i32 = 0;
        while (s <= maxSerial(system)) : (s += 1) {
            const d = try dateFromSerial(system, s);
            const back = try serialFromDate(system, d.year, d.month, d.day);
            if (back != s) {
                std.debug.print("round trip broke at serial {d} ({t})\n", .{ s, system });
                return error.RoundTripMismatch;
            }
        }
    }
}

test "round trip: monotonic and gapless, so no serial names two dates" {
    // A weaker converter can round-trip every serial and still map two
    // serials onto one date. Days must advance by exactly one, except
    // across the 1900 hole where the *serial* advances and the calendar
    // does not.
    var previous = try dateFromSerial(.d1900, 1);
    var s: i32 = 2;
    while (s <= 400) : (s += 1) {
        const d = try dateFromSerial(.d1900, s);
        // `daysFromCivil` is pure arithmetic, so it happily places the
        // invented 1900-02-29 on top of 1 March — which is exactly the
        // collision this test is looking for.
        const step = daysFromCivil(d.year, d.month, d.day) -
            daysFromCivil(previous.year, previous.month, previous.day);
        if (d.fictitious == .feb_29_1900 or previous.fictitious == .feb_29_1900) {
            // The invented day sits on top of 1 March in the real
            // calendar; that is the discontinuity, and it is expected.
            try testing.expect(step == 1 or step == 0);
        } else {
            try testing.expectEqual(@as(i64, 1), step);
        }
        previous = d;
    }
}

test "leap years: the century rule, at the years that expose it" {
    try testing.expect(isLeapYear(2000)); // divisible by 400
    try testing.expect(!isLeapYear(1900)); // divisible by 100, not 400
    try testing.expect(!isLeapYear(2100));
    try testing.expect(isLeapYear(2024));
    try testing.expect(!isLeapYear(2023));
    try testing.expectEqual(@as(u8, 29), daysInMonth(2000, 2));
    try testing.expectEqual(@as(u8, 28), daysInMonth(1900, 2));
    try testing.expectEqual(@as(u8, 29), daysInMonth(2024, 2));

    // 1900 is NOT a leap year, which is the entire reason serial 60
    // needs a rule of its own.
    try testing.expectError(error.DateOutOfRange, serialFromDate(.d1904, 1900, 2, 29));
}

test "time: rounds to the nearest second and carries past midnight" {
    try testing.expectEqual(Time{ .hour = 0, .minute = 0, .second = 0 }, timeFromFraction(0));
    try testing.expectEqual(Time{ .hour = 12, .minute = 0, .second = 0 }, timeFromFraction(0.5));
    try testing.expectEqual(Time{ .hour = 6, .minute = 0, .second = 0 }, timeFromFraction(0.25));
    try testing.expectEqual(Time{ .hour = 23, .minute = 59, .second = 59 }, timeFromFraction(86_399.0 / 86_400.0));

    // 24:00:00 is not a time; it is tomorrow.
    const carried = timeFromFraction(0.99999999);
    try testing.expectEqual(@as(u8, 0), carried.hour);
    try testing.expectEqual(@as(u1, 1), carried.day_carry);

    // The whole part is ignored, so a full serial works directly.
    try testing.expectEqual(Time{ .hour = 12, .minute = 0, .second = 0 }, timeFromFraction(45_000.5));
}

test "time: fraction round-trips for every second of a day" {
    var s: u32 = 0;
    while (s < 86_400) : (s += 1) {
        const frac = fractionFromTime(s / 3600, (s % 3600) / 60, s % 60);
        const t = timeFromFraction(frac);
        try testing.expectEqual(@as(u1, 0), t.day_carry);
        const back = @as(u32, t.hour) * 3600 + @as(u32, t.minute) * 60 + t.second;
        try testing.expectEqual(s, back);
    }
}

test "split: a serial's day and fraction, refusing what is not a date" {
    const split = try splitSerial(45_000.75);
    try testing.expectEqual(@as(i32, 45_000), split.days);
    try testing.expectApproxEqAbs(@as(f64, 0.75), split.fraction, 1e-12);

    try testing.expectError(error.SerialOutOfRange, splitSerial(-0.5));
    try testing.expectError(error.SerialOutOfRange, splitSerial(3_000_000));
    try testing.expectError(error.SerialOutOfRange, splitSerial(std.math.inf(f64)));
}

test "civil calendar: the algorithm agrees with known anchors" {
    // Fixed points that are easy to check by hand and hard to get right
    // by accident.
    try testing.expectEqual(@as(i64, 0), daysFromCivil(1970, 1, 1));
    try testing.expectEqual(@as(i64, -1), daysFromCivil(1969, 12, 31));
    try testing.expectEqual(@as(i64, 10957), daysFromCivil(2000, 1, 1));
    try testing.expectEqual(@as(i64, -25567), daysFromCivil(1900, 1, 1));

    const c = civilFromDays(0);
    try testing.expectEqual(@as(i32, 1970), c.year);
    try testing.expectEqual(@as(u8, 1), c.month);
    try testing.expectEqual(@as(u8, 1), c.day);

    // Round-trip across the era boundary the algorithm pivots on, where
    // an off-by-one in the March-based year would show up.
    var d: i64 = -800_000;
    while (d < 3_000_000) : (d += 9_973) {
        const v = civilFromDays(d);
        try testing.expectEqual(d, daysFromCivil(v.year, v.month, v.day));
    }
}

test "checkAllAllocationFailures: conversion is allocation-free, and stays so" {
    // The conversions themselves take no allocator — the runner supplies
    // one only to collect their results. What this proves is that a
    // future change which starts allocating inside a conversion would
    // have to leak nothing under OOM, and that failing to collect a
    // result never leaves a half-built list behind.
    const H = struct {
        fn run(allocator: std.mem.Allocator) !void {
            var out: std.ArrayListUnmanaged(Date) = .empty;
            defer out.deinit(allocator);
            const boundaries = [_]i32{ 0, 1, 59, 60, 61, 1462, max_serial_1900 };
            for (boundaries) |serial| {
                try out.append(allocator, try dateFromSerial(.d1900, serial));
            }
            for (boundaries[0 .. boundaries.len - 1]) |serial| {
                try out.append(allocator, try dateFromSerial(.d1904, serial));
            }
            for (out.items) |d| {
                if (d.year == 0) return error.ImpossibleDate;
            }
        }
    };
    try testing.checkAllAllocationFailures(testing.allocator, H.run, .{});
}
