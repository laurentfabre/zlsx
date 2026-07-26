//! emb-4B carrier verifier.
//!
//! Opens a fixture that has been round-tripped through a consumer tool
//! and reports, per carrier, whether the recovery marker survived.
//!
//! Usage:
//!   zig build emb4b-verify -- <file.xlsx>
//!
//! Exit code is the number of carriers that did NOT survive (0–6), so
//! 0 means full preservation and the runner can record a count without
//! parsing prose. The per-carrier table is always printed — the table
//! is the deliverable, the exit code is just a summary. Exit 64 is
//! reserved for usage/IO failure so it can never be mistaken for a
//! carrier count.
//!
//! Detection is deliberately **byte-level** rather than structural: a
//! consumer that preserves the payload but relocates, reindents or
//! re-encodes the containing XML still counts as preserving it. The one
//! exception is `opc_part`, which is read back through
//! `Workbook.embeddings()` so this stays a like-for-like control
//! against emb-4's verdict.

const std = @import("std");
const pkg = @import("zlsx_pkg");
const carriers = @import("emb4b_carriers");

const Carrier = carriers.Carrier;

const USAGE_EXIT: u8 = 64;

pub fn main(init: std.process.Init) !u8 {
    const allocator = init.gpa;
    const io = init.io;

    const args = try init.minimal.args.toSlice(init.arena.allocator());

    var stdout_buf: [2048]u8 = undefined;
    var stdout_w = std.Io.File.stdout().writer(io, &stdout_buf);
    const stdout = &stdout_w.interface;
    defer stdout.flush() catch {};

    if (args.len < 2) {
        try stdout.print("usage: {s} <file.xlsx>\n", .{args[0]});
        return USAGE_EXIT;
    }
    const in_path = args[1];

    var wb = pkg.Workbook.open(allocator, io, in_path) catch |e| {
        try stdout.print("emb-4B verify: {s}\n  cannot open: error.{s}\n", .{ in_path, @errorName(e) });
        return USAGE_EXIT;
    };
    defer wb.deinit();

    try stdout.print("emb-4B verify: {s}\n", .{in_path});

    var lost: u8 = 0;
    for (carriers.ALL) |c| {
        var mbuf: [carriers.MARKER_MAX]u8 = undefined;
        const m = carriers.marker(&mbuf, c);
        const survived = try check(allocator, &wb, c, m);
        if (!survived) lost += 1;
        try stdout.print("  {s:<9} {s:<8} {s}\n", .{
            c.slug(),
            if (survived) "SURVIVED" else "STRIPPED",
            c.location(),
        });
    }

    // A hidden marker sheet that comes back visible is a partial loss
    // worth recording separately: the data survived but the carrier's
    // concealment did not, which is exactly the trade-off the
    // cell_data carrier is being evaluated on.
    if (try sheetHidden(&wb, carriers.CELL_SHEET)) |hidden| {
        try stdout.print("  {s:<9} {s:<8} {s}\n", .{
            "CELLHIDE",
            if (hidden) "SURVIVED" else "STRIPPED",
            "state=\"hidden\" on the marker sheet (informational)",
        });
    }

    try stdout.print("carriers lost: {d}/{d}\n", .{ lost, carriers.ALL.len });
    return lost;
}

fn check(
    allocator: std.mem.Allocator,
    wb: *pkg.Workbook,
    c: Carrier,
    marker: []const u8,
) !bool {
    return switch (c) {
        // Structural, to stay a like-for-like control against emb-4.
        .opc_part => blk: {
            const view = wb.embeddings() catch break :blk false;
            if (view == null) break :blk false;
            break :blk std.mem.eql(u8, view.?.index.model, marker);
        },
        .custom_xml => partContains(wb, "customXml/item1.xml", marker),
        .doc_props => partContains(wb, "docProps/custom.xml", marker),
        .defined_name, .ext_lst => partContains(wb, "xl/workbook.xml", marker),
        // The marker may land in the SST or inline in the sheet, and
        // the sheet may have been renumbered by the consumer, so scan
        // every candidate part rather than assuming a path.
        .cell_data => anyPartContains(allocator, wb, marker),
    };
}

fn partContains(wb: *pkg.Workbook, name: []const u8, needle: []const u8) !bool {
    const p = (try wb.store.part(name)) orelse return false;
    return std.mem.indexOf(u8, p.bytes, needle) != null;
}

/// Scan every part in the package. Used for `cell_data`, whose landing
/// site depends on the consumer's shared-string policy and sheet
/// numbering.
///
/// Goes back through `store.part(name)` for each entry rather than
/// reading `store.parts[i].bytes` directly: part bytes are lazily
/// materialized, so an untouched part reports `bytes.len == 0` and a
/// direct scan would silently miss the marker.
fn anyPartContains(
    allocator: std.mem.Allocator,
    wb: *pkg.Workbook,
    needle: []const u8,
) !bool {
    // `part()` mutates the store's cache, which invalidates the
    // `parts` slice's byte pointers but not the names — so snapshot
    // the names first.
    var names: std.ArrayListUnmanaged([]const u8) = .empty;
    defer names.deinit(allocator);
    for (wb.store.parts) |p| try names.append(allocator, p.name);

    for (names.items) |n| {
        const p = (try wb.store.part(n)) orelse continue;
        if (std.mem.indexOf(u8, p.bytes, needle) != null) return true;
    }
    return false;
}

/// Whether `xl/workbook.xml` still marks the named sheet hidden.
/// Null when the sheet is gone entirely.
fn sheetHidden(wb: *pkg.Workbook, name: []const u8) !?bool {
    const p = (try wb.store.part("xl/workbook.xml")) orelse return null;
    var needle_buf: [64]u8 = undefined;
    const needle = std.fmt.bufPrint(&needle_buf, "name=\"{s}\"", .{name}) catch return null;
    const at = std.mem.indexOf(u8, p.bytes, needle) orelse return null;
    const rest = p.bytes[at..];
    const end = std.mem.indexOfScalar(u8, rest, '>') orelse return null;
    return std.mem.indexOf(u8, rest[0..end], "state=\"hidden\"") != null;
}
