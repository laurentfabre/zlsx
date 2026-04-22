// Writer benchmark — 1000 rows × 10 cols of mixed content, one header
// row with a bold blue fill style, 3 number-format styles reused across
// columns. Save to the path passed as argv[1].
//
// Uses `std.heap.smp_allocator` because that's what a production
// writer would typically plug in — it's what Zig's own toolchain
// reaches for when throughput matters. DebugAllocator would add
// several-ms of per-alloc tracking overhead that isn't representative
// of what downstream users see.
const std = @import("std");
const xlsx = @import("zlsx");

pub fn main() !void {
    const alloc = std.heap.smp_allocator;

    const args = try std.process.argsAlloc(alloc);
    defer std.process.argsFree(alloc, args);
    if (args.len < 2) {
        std.debug.print("usage: {s} <out.xlsx>\n", .{args[0]});
        return;
    }

    var w = xlsx.Writer.init(alloc);
    defer w.deinit();

    const header = try w.addStyle(.{
        .font_bold = true,
        .font_color_argb = 0xFFFFFFFF,
        .fill_pattern = .solid,
        .fill_fg_argb = 0xFF1E3A8A,
        .alignment_horizontal = .center,
    });
    const money = try w.addStyle(.{ .number_format = "$#,##0.00" });
    const pct = try w.addStyle(.{ .number_format = "0.00%" });

    var sheet = try w.addSheet("Bench");
    try sheet.setColumnWidth(0, 20);
    sheet.freezePanes(1, 0);

    const header_cells = [_]xlsx.Cell{
        .{ .string = "Name" }, .{ .string = "Amount" }, .{ .string = "Share" },
        .{ .string = "Qty" },  .{ .string = "Active" }, .{ .string = "Code" },
        .{ .string = "A" },    .{ .string = "B" },      .{ .string = "C" },
        .{ .string = "D" },
    };
    const header_styles = [_]u32{ header, header, header, header, header, header, header, header, header, header };
    try sheet.writeRowStyled(&header_cells, &header_styles);

    var name_buf: [16]u8 = undefined;
    var code_buf: [16]u8 = undefined;
    var i: u32 = 0;
    while (i < 1000) : (i += 1) {
        const n = try std.fmt.bufPrint(&name_buf, "row_{d}", .{i});
        const c = try std.fmt.bufPrint(&code_buf, "CODE_{x}", .{i});
        const cells = [_]xlsx.Cell{
            .{ .string = n },
            .{ .number = 100.0 + @as(f64, @floatFromInt(i)) * 1.5 },
            .{ .number = @as(f64, @floatFromInt(i)) / 1000.0 },
            .{ .integer = @as(i64, i) },
            .{ .boolean = i % 2 == 0 },
            .{ .string = c },
            .{ .integer = @as(i64, i) * 7 },
            .{ .number = @as(f64, @floatFromInt(i)) * 0.1 },
            .{ .string = "x" },
            .empty,
        };
        const styles = [_]u32{ 0, money, pct, 0, 0, 0, 0, 0, 0, 0 };
        try sheet.writeRowStyled(&cells, &styles);
    }

    try w.save(args[1]);
}
