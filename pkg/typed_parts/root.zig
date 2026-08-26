//! Typed-overlay parsers for known OOXML parts (B1 iter-wb-1).
//!
//! Each child file is a stdlib-only, allocator-explicit, defensive
//! parser exposing a typed view over one OOXML part:
//!
//!   - `workbook_xml.zig` — `xl/workbook.xml`        (sheets list, defined names, calc props)
//!   - `sheet_xml.zig`    — `xl/worksheets/sheetN.xml` (rows, cells, merges, hyperlinks, validations, CF, freeze)
//!   - `sst_xml.zig`      — `xl/sharedStrings.xml`   (string entries, rich runs, on-demand entity decode)
//!   - `styles_xml.zig`   — `xl/styles.xml`          (numFmts, fonts, fills, borders, cellXfs, alignment)
//!   - `theme_xml.zig`    — `xl/theme/themeN.xml`    (color scheme: srgbClr + sysClr)
//!
//! Re-exported from `pkg/root.zig` under the `typed_parts` namespace
//! so callers reach them via:
//!
//!     const pkg = @import("zlsx_pkg");
//!     var view = try pkg.typed_parts.workbook_xml.parse(allocator, bytes);
//!     defer view.deinit(allocator);
//!
//! These parsers are read-only in iter-wb-1. Emit / mutate paths
//! land in iter-wb-4 (Worksheet.setCell + delta-emit save). See
//! `docs/plans/archive/workbook-overlay.md`.

pub const workbook_xml = @import("workbook_xml.zig");
pub const sheet_xml = @import("sheet_xml.zig");
pub const sst_xml = @import("sst_xml.zig");
pub const styles_xml = @import("styles_xml.zig");
pub const theme_xml = @import("theme_xml.zig");
pub const doc_props_xml = @import("doc_props_xml.zig");

test {
    _ = workbook_xml;
    _ = sheet_xml;
    _ = sst_xml;
    _ = styles_xml;
    _ = theme_xml;
    _ = doc_props_xml;
}
