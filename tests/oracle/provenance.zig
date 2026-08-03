//! Provenance records for oracle manifests (M1b, `goal_formula.md` §8.2).
//!
//! A recorded value is only evidence if you know what produced it. §8.2
//! names six facts — Excel build, LibreOffice build, OS, locale,
//! extractor version, workbook digest — and this module makes them
//! mandatory rather than optional: `validate` fails on any blank, and
//! the replay gate runs it over every committed manifest.
//!
//! Why the strictness. When Excel 16.112 disagrees with a golden
//! recorded under 16.111, the useful question is "which build?", and a
//! manifest that answered "unknown" would be a divergence nobody can
//! act on. Locale matters for the same reason: the argument separator,
//! the decimal point, and several function results are locale-sensitive
//! (§5.4), so a golden recorded under `fr_FR` is not comparable with
//! one recorded under `en_US` even from the same build.

const std = @import("std");

pub const Error = error{
    ProvenanceMissingField,
    ProvenanceBadDigest,
    ProvenanceUnknownAdapter,
};

/// Which oracle produced the manifest. §8.2's four legs.
pub const Adapter = enum {
    /// Excel for Mac driven over AppleScript.
    excel_mac,
    /// LibreOffice Calc, pinned invocation + dedicated profile.
    libreoffice,
    /// Hand-derived spec suite. Values come from a documented reading of
    /// the specification, not from running anything — which is what
    /// makes it the tie-breaker at a divergence point.
    hand_spec,
    /// Screened corpus workbooks: a consistency signal, never a
    /// primary authority (§8.2 precedence).
    corpus,

    pub fn parse(s: []const u8) Error!Adapter {
        return std.meta.stringToEnum(Adapter, s) orelse error.ProvenanceUnknownAdapter;
    }

    /// True when the adapter's values come from running a spreadsheet
    /// application. Volatile formulas are excluded from exactly these
    /// (§8.2) — a hand-derived value for `RAND()` is not a draw, it is
    /// a documented statement about the function's contract.
    pub fn isExternalApp(self: Adapter) bool {
        return switch (self) {
            .excel_mac, .libreoffice => true,
            .hand_spec, .corpus => false,
        };
    }
};

pub const Record = struct {
    adapter: []const u8,
    /// Application build string. For `hand_spec` this names the
    /// specification and section the values were derived from, which is
    /// the equivalent evidence.
    app_build: []const u8,
    /// Host OS and version.
    os: []const u8,
    /// Locale the recording ran under.
    locale: []const u8,
    /// `extractor.version` at recording time.
    extractor_version: []const u8,
    /// SHA-256 (lowercase hex, 64 chars) of the .xlsx the values came
    /// from — AFTER recalculation, so it identifies the exact bytes
    /// that were read.
    workbook_digest: []const u8,
    /// ISO 8601 date of the recording. Not a §8.2 requirement; carried
    /// because "when" is the first question asked about a stale golden.
    recorded: []const u8,

    pub fn validate(self: Record) Error!void {
        _ = try Adapter.parse(self.adapter);
        // Every field is load-bearing, so blankness is a failure rather
        // than a default. An "unknown" provenance is worse than no
        // manifest: it looks like evidence.
        inline for (.{
            self.app_build,       self.os,
            self.locale,          self.extractor_version,
            self.workbook_digest, self.recorded,
        }) |field| {
            if (field.len == 0) return error.ProvenanceMissingField;
        }
        if (self.workbook_digest.len != 64) return error.ProvenanceBadDigest;
        for (self.workbook_digest) |c| {
            const ok = (c >= '0' and c <= '9') or (c >= 'a' and c <= 'f');
            if (!ok) return error.ProvenanceBadDigest;
        }
    }

    pub fn adapterEnum(self: Record) Error!Adapter {
        return Adapter.parse(self.adapter);
    }
};

// ─── tests ───────────────────────────────────────────────────────

const testing = std.testing;

const good: Record = .{
    .adapter = "excel_mac",
    .app_build = "Microsoft Excel 16.111.2",
    .os = "macOS 26.5 (Darwin 25.5.0)",
    .locale = "en_US.UTF-8",
    .extractor_version = "oracle-extractor-1",
    .workbook_digest = "0" ** 64,
    .recorded = "2026-08-03",
};

test "a complete record validates" {
    try good.validate();
    try testing.expectEqual(Adapter.excel_mac, try good.adapterEnum());
}

test "every field is mandatory" {
    // Walk the struct: blanking ANY field must fail. Written as a loop
    // over field names so a field added later is covered without anyone
    // remembering to extend the test.
    inline for (std.meta.fields(Record)) |field| {
        var r = good;
        @field(r, field.name) = "";
        // A blank adapter fails the enum parse first; every other blank
        // reaches the emptiness check.
        try testing.expectError(
            if (std.mem.eql(u8, field.name, "adapter"))
                error.ProvenanceUnknownAdapter
            else
                error.ProvenanceMissingField,
            r.validate(),
        );
    }
}

test "digest must be 64 lowercase hex characters" {
    var short = good;
    short.workbook_digest = "abc";
    try testing.expectError(error.ProvenanceBadDigest, short.validate());

    var upper = good;
    upper.workbook_digest = "A" ** 64;
    try testing.expectError(error.ProvenanceBadDigest, upper.validate());

    var nonhex = good;
    nonhex.workbook_digest = "g" ** 64;
    try testing.expectError(error.ProvenanceBadDigest, nonhex.validate());

    var ok = good;
    ok.workbook_digest = "0123456789abcdef" ** 4;
    try ok.validate();
}

test "unknown adapter names are refused" {
    var r = good;
    r.adapter = "google_sheets";
    try testing.expectError(error.ProvenanceUnknownAdapter, r.validate());
}

test "external-app classification drives volatile exclusion" {
    try testing.expect(Adapter.excel_mac.isExternalApp());
    try testing.expect(Adapter.libreoffice.isExternalApp());
    try testing.expect(!Adapter.hand_spec.isExternalApp());
    try testing.expect(!Adapter.corpus.isExternalApp());
}

test "parses from JSON as written by the recording scripts" {
    const json =
        \\{
        \\  "adapter": "libreoffice",
        \\  "app_build": "LibreOffice 26.2.5.2",
        \\  "os": "macOS 26.5",
        \\  "locale": "en_US.UTF-8",
        \\  "extractor_version": "oracle-extractor-1",
        \\  "workbook_digest": "0123456789abcdef0123456789abcdef0123456789abcdef0123456789abcdef",
        \\  "recorded": "2026-08-03"
        \\}
    ;
    const parsed = try std.json.parseFromSlice(Record, testing.allocator, json, .{});
    defer parsed.deinit();
    try parsed.value.validate();
    try testing.expectEqual(Adapter.libreoffice, try parsed.value.adapterEnum());
}
