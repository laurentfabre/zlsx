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

/// dbx-1: crash-safe whole-file replacement, exported for the CLI's
/// `dbx pull` (write-verify-rename). Same primitive PartStore.save uses.
pub const AtomicFile = @import("atomic_file.zig").AtomicFile;
pub const PartStore = store_mod.PartStore;
/// M5b0: the ref-counted source a `PartStore` reads through. Exported
/// because it is now part of the store's shape — several generations
/// may share one, and the last to `deinit` is the one that closes.
pub const SourceBacking = store_mod.SourceBacking;
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

/// Typed-overlay parsers for well-known OOXML parts (B1 iter-wb-1).
/// See `pkg/typed_parts/root.zig` for the per-part API.
pub const typed_parts = @import("typed_parts/root.zig");

/// Workbook + Worksheet typed-overlay roots (B1 iter-wb-2).
/// Read-only in this iter; mutation lands in iter-wb-4.
const workbook_mod = @import("workbook.zig");
pub const Workbook = workbook_mod.Workbook;
pub const Worksheet = workbook_mod.Worksheet;
pub const ResolvedStyle = workbook_mod.ResolvedStyle;
pub const WorkbookError = workbook_mod.Error;
pub const NumberFormatInfo = workbook_mod.NumberFormatInfo;
pub const EmbeddingView = workbook_mod.EmbeddingView;
pub const EmbeddingState = workbook_mod.EmbeddingState;
pub const EmbeddingCoverageView = workbook_mod.EmbeddingCoverageView;
pub const EmbeddingCoverageInput = workbook_mod.EmbeddingCoverageInput;

/// Shared-strings extension-plan substrate (B3 iter-wr-1). Re-exported
/// so `xlsx.Writer` (in `src/`) can stage strings + rich-text runs
/// through the same dedup/storage types Workbook uses.
pub const SstExtensionPlan = workbook_mod.SstExtensionPlan;
pub const RichRun = workbook_mod.RichRun;
pub const RichEntry = workbook_mod.RichEntry;

/// Greenfield drawing-emit surface (B1 iter-wb-c2b). Re-exported so
/// callers can spell `pkg.ImageMime`, `pkg.ImageCellAnchor` without
/// reaching into `drawing_emit` directly.
const drawing_emit_mod = @import("drawing_emit.zig");
pub const ImageMime = drawing_emit_mod.ImageMime;
pub const ImageCellAnchor = drawing_emit_mod.ImageCellAnchor;

/// Editor — load-modify-save overlay (relocated to pkg/ in B2 iter-er-0).
const editor_mod = @import("editor.zig");
pub const Editor = editor_mod.Editor;
/// The element type `Editor.setCells` takes. Without it the batch form is
/// unreachable from outside: a consumer cannot spell the slice it needs,
/// and falls back to N × `setCell` — which is what nemonym's xlsx pipeline
/// had to do.
pub const Edit = editor_mod.Edit;

/// Document-properties view + scrub mask, re-exported so consumers can
/// name the types without reaching into `typed_parts`.
pub const DocProps = workbook_mod.DocProps;
pub const DocPropsMask = workbook_mod.DocPropsMask;
pub const CellSpan = editor_mod.CellSpan;
pub const WorksheetSpans = editor_mod.WorksheetSpans;

/// Embedding-part wire-format primitives (emb-1). Pure functions over
/// byte slices for the on-disk index.xml / index.xml.rels / vec.bin /
/// hashes.bin parts under `xl/zlsxEmbeddings/`.
pub const embedding_part = @import("embedding_part.zig");
pub const recovery_record = @import("recovery_record.zig");
