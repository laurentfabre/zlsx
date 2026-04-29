//! Public root for the zlsx package layer (B0 + C2a).
//!
//! Downstream Zig consumers reach this via:
//!
//!     // in their build.zig.zon
//!     .dependencies = .{
//!         .zlsx_pkg = .{ .url = "..." },
//!     }
//!
//!     // in their code
//!     const pkg = @import("zlsx_pkg");
//!     var store = try pkg.PartStore.open(allocator, "in.xlsx");
//!     defer store.deinit();
//!     for (try pkg.imageAnchors(&store, allocator)) |a| { ... }
//!
//! Lives alongside the `zlsx` reader/writer module rather than
//! folded into it: the package layer is meant to be usable on its
//! own for callers (image-extraction utilities, OOXML inspection
//! tools) that don't need the full Book/Editor surface.

const store_mod = @import("store.zig");
const drawings_mod = @import("drawings.zig");

pub const PartStore = store_mod.PartStore;
pub const Part = store_mod.Part;
pub const Relationship = store_mod.Relationship;
pub const TargetMode = store_mod.TargetMode;
pub const Error = store_mod.Error;

pub const ImageAnchor = drawings_mod.ImageAnchor;
pub const CellAnchor = drawings_mod.CellAnchor;
pub const AbsoluteAnchor = drawings_mod.AbsoluteAnchor;
pub const imageAnchors = drawings_mod.imageAnchors;
pub const ChartAnchor = drawings_mod.ChartAnchor;
pub const ChartType = drawings_mod.ChartType;
pub const chartAnchors = drawings_mod.chartAnchors;
