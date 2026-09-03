//! Corpus-level integration tests for the package layer (zlsx_pkg).
//!
//! For every committed-or-fetched .xlsx in tests/corpus/, walks
//! through `PartStore.open` + drains `partNames`, `imageParts`,
//! `imageAnchors`, `chartAnchors`. The contract is "no crash, no
//! hang, no unbounded allocation". Per-fixture content assertions
//! live in their dedicated unit tests inside pkg/store.zig and
//! pkg/drawings.zig — this file is a wide robustness sweep.
//!
//! Adversarial fixtures (truncated ZIPs, malformed CRCs, bogus
//! offsets) MUST surface a typed error rather than panic. Anything
//! that opens cleanly must return a non-empty parts list.

const std = @import("std");
const pkg = @import("zlsx_pkg");

const corpus_dir = "tests/corpus/";

const Verdict = enum {
    /// Container is well-formed and at least [Content_Types].xml is present.
    must_open_clean,
    /// Container itself is broken — EOCD missing, signature mangled, etc.
    /// PartStore.open MUST return a typed error rather than crash.
    must_error_typed,
    /// Container is structurally well-formed but its payloads or
    /// higher-level OOXML semantics are damaged (bad CRC, encrypted
    /// EncryptedPackage stream, truncation past EOCD). PartStore.open
    /// is structural-only by design, so these open cleanly — the
    /// damage only surfaces when a consumer actually decompresses.
    must_open_with_lazy_corruption,
};

const fixture_list = [_]struct { name: []const u8, verdict: Verdict }{
    // Valid OOXML (committed).
    .{ .name = "frictionless_2sheets.xlsx", .verdict = .must_open_clean },
    .{ .name = "openpyxl_guess_types.xlsx", .verdict = .must_open_clean },
    .{ .name = "openpyxl_chart.xlsx", .verdict = .must_open_clean },
    .{ .name = "phpoi_test1.xlsx", .verdict = .must_open_clean },
    .{ .name = "calamine_empty_s_attribute.xlsx", .verdict = .must_open_clean },
    .{ .name = "calamine_empty_shared_string.xlsx", .verdict = .must_open_clean },
    .{ .name = "calamine_encoded_entities.xlsx", .verdict = .must_open_clean },
    .{ .name = "calamine_non_monotonic_si.xlsx", .verdict = .must_open_clean },
    .{ .name = "openxlsx_loadExample.xlsx", .verdict = .must_open_clean },
    .{ .name = "phpsheet_3654c.xlsx", .verdict = .must_open_clean },
    .{ .name = "poi_57893_many_merges.xlsx", .verdict = .must_open_clean },
    .{ .name = "poi_58325_db.xlsx", .verdict = .must_open_clean },
    .{ .name = "poi_excel_with_trash_item.xlsx", .verdict = .must_open_clean },
    .{ .name = "poi_poc_shared_strings.xlsx", .verdict = .must_open_clean },

    // Valid (fetched larger fixtures).
    .{ .name = "ecdc_covid.xlsx", .verdict = .must_open_clean },
    .{ .name = "ons_cpi_detailed.xlsx", .verdict = .must_open_clean },
    .{ .name = "wdi_excel.xlsx", .verdict = .must_open_clean },
    .{ .name = "worldbank_catalog.xlsx", .verdict = .must_open_clean },

    // Adversarial / malformed.
    //
    // PartStore.open is structural-only: parses EOCD / CDFH / LFH but
    // does NOT decompress. The verdict reflects the boundary:
    //   must_error_typed              — structural ZIP is broken
    //   must_open_with_lazy_corruption — ZIP fine, payload/OOXML broken
    //
    // Probed empirically once and pinned here so a regression that
    // flips the boundary surfaces immediately.
    // Since the deferred-decompress refactor, CRC32 mismatch no
    // longer surfaces at open() for the BULK of parts — only at
    // first `part(name)` access. The structural eager-decompress
    // path (`[Content_Types].xml` and every `_rels/*.rels`) still
    // CRC-checks at open(), so a bad CRC on those parts errors at
    // open(); a bad CRC on a worksheet / image / other deferred
    // part errors only on materialization.
    .{ .name = "derived_bad_crc32.xlsx", .verdict = .must_error_typed }, // bad CRC on a STRUCTURAL part — eager check still fires at open()
    .{ .name = "derived_truncated_mid_payload.xlsx", .verdict = .must_error_typed },
    .{ .name = "derived_truncated_pre_eocd.xlsx", .verdict = .must_error_typed },
    .{ .name = "derived_truncated_signature.xlsx", .verdict = .must_error_typed },
    .{ .name = "poi_MalformedSSTCount.xlsx", .verdict = .must_open_clean }, // structural ZIP is valid
    .{ .name = "poi_clusterfuzz_xssf.xlsx", .verdict = .must_open_with_lazy_corruption }, // bad CRC on a deferred (worksheet) part — surfaces only on materialize
    .{ .name = "poi_crash_274d6342.xlsx", .verdict = .must_error_typed }, // BadZip — broken CDFH
    .{ .name = "poi_crash_9bf3cd4b.xlsx", .verdict = .must_error_typed }, // NotPkzip — no EOCD
    .{ .name = "poi_workbook_password_2013.xlsx", .verdict = .must_open_with_lazy_corruption }, // encrypted CFB inside valid ZIP
    .{ .name = "poi_xlsx_corrupted.xlsx", .verdict = .must_open_with_lazy_corruption }, // ZIP fine, payload damage
    .{ .name = "poi_xxe_in_schema.xlsx", .verdict = .must_open_clean },
};

test "PartStore corpus sweep — open + walk every fixture without crash" {
    var threaded: std.Io.Threaded = .init(std.testing.allocator, .{});
    defer threaded.deinit();
    const io = threaded.io();
    const alloc = std.testing.allocator;
    var any_seen: bool = false;
    for (fixture_list) |fix| {
        var path_buf: [256]u8 = undefined;
        const path = try std.fmt.bufPrint(&path_buf, "{s}{s}", .{ corpus_dir, fix.name });
        std.Io.Dir.cwd().access(io, path, .{}) catch continue;
        any_seen = true;

        const open_result = pkg.PartStore.open(alloc, io, path);
        switch (fix.verdict) {
            .must_open_clean => {
                var store = open_result catch |err| {
                    std.debug.print("\n  [unexpected error] {s} -> {s}\n", .{ fix.name, @errorName(err) });
                    return err;
                };
                defer store.deinit();
                // Walks must not crash; results may be empty.
                _ = try store.partNames();
                _ = try store.imageParts();
                const images = try pkg.imageAnchors(&store, alloc);
                defer alloc.free(images);
                const charts = try pkg.chartAnchors(&store, alloc);
                defer {
                    for (charts) |c| alloc.free(c.series_refs);
                    alloc.free(charts);
                }
                // At least the [Content_Types].xml part must be present
                // in any well-formed OOXML package.
                try std.testing.expect((try store.part("[Content_Types].xml")) != null);
            },
            .must_error_typed => {
                if (open_result) |*s| {
                    var store = s.*;
                    store.deinit();
                    std.debug.print("\n  [unexpected success] {s} opened cleanly but should have errored\n", .{fix.name});
                    return error.TestUnexpectedResult;
                } else |_| {
                    // Any typed error is acceptable. The contract is
                    // "no crash" — discriminating which error fired
                    // belongs in per-fixture unit tests.
                }
            },
            .must_open_with_lazy_corruption => {
                // PartStore.open is structural-only: it parses EOCD
                // / CDFH / LFH metadata but does NOT decompress
                // payloads. Fixtures whose corruption lives inside
                // payload bytes (bad CRC, truncation past EOCD,
                // encrypted EncryptedPackage stream) therefore open
                // cleanly here — the corruption surfaces when the
                // reader actually decompresses.
                var store = open_result catch |err| {
                    std.debug.print("\n  [unexpected error] {s} -> {s}\n", .{ fix.name, @errorName(err) });
                    return err;
                };
                defer store.deinit();
                _ = try store.partNames();
                // imageParts() materializes each image's payload to
                // expose decompressed bytes — that's where lazy
                // corruption (bad CRC etc.) surfaces. Either outcome
                // is acceptable on the lazy-corruption arm: a typed
                // error is the contract for damaged payloads, and a
                // clean return is the contract for fixtures whose
                // corruption sits in non-image parts.
                _ = store.imageParts() catch |err| switch (err) {
                    error.BadZip, error.UnsupportedCompression, error.OutOfMemory => {},
                    else => return err,
                };
                // Drawing walks are payload-driven; skip on lazy-corruption
                // fixtures since decompression is allowed to fail.
            },
        }
    }
    if (!any_seen) {
        std.debug.print("\n  [skip] no corpus fixtures present — run scripts/fetch_test_corpus.sh\n", .{});
        return error.SkipZigTest;
    }
}
