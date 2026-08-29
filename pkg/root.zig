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
/// S1: the decompression limits every opener in this tree enforces
/// (`PartStore.open`, `Editor.open`, and the core reader behind both).
/// One home in `zlsx_control`; re-exported here and from `zlsx` so a
/// caller of either module can read the numbers without a third import.
pub const DecompressLimits = @import("zlsx_control").DecompressLimits;
pub const decompress_limits = @import("zlsx_control").decompress_limits;

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
/// §9.1 M10q's arena-fill probe. Public so the bench harness can install
/// tallies; inert in every build that does not (`pkg/fill.zig`).
pub const fill_probe = workbook_mod.fill_probe;
/// §9.1f M10w's graph-block census and page-rate knob, public on the same
/// terms and for the complementary reason: the build era's peak is one
/// `gpa` block carved by a `FixedBufferAllocator`, which is the one region
/// an arena tally cannot reach.
///
/// **Two symbols, not the module.** Re-exporting `zlsx_formula.graph`
/// would put `Graph` — and `Graph.block`, whose slack past the last carve
/// is unspecified — in this package's public API for the sake of a probe.
pub const graph_probe = struct {
    const graph = @import("zlsx_formula").graph;

    pub const Census = graph.Census;

    /// Install a census sink, or clear it with `null`. Not thread-safe;
    /// see `graph.census_sink`.
    pub fn install(sink: ?*Census) void {
        graph.census_sink = sink;
    }

    /// N bytes of never-carved slack on the retained block, so its cost
    /// in resident pages can be measured against its own unpadded run on
    /// the same binary. Zero is the working default.
    pub fn setBlockPadding(bytes: usize) void {
        graph.probe_block_padding = bytes;
    }
};
/// §9.1g M10x's census and the two padding knobs for the **decode** half
/// — the half §9.1f's cut left holding the resident maximum. Public on
/// `graph_probe`'s terms and as named symbols only, for the same reason:
/// re-exporting the modules to reach a probe would put the whole formula
/// engine in this package's public API.
pub const shared_probe = struct {
    const calc = @import("zlsx_formula").calc;
    const decode = @import("zlsx_formula").decode;

    pub const Census = calc.SharedCensus;

    /// Install a census sink, or clear it with `null`. Not thread-safe;
    /// see `calc.census_sink`.
    pub fn install(sink: ?*Census) void {
        calc.census_sink = sink;
    }

    /// N never-written bytes held for exactly one `scanSheet` — the span
    /// §9.1c's candidate 2 lives in.
    pub fn setScanPadding(bytes: usize) void {
        decode.probe_scan_padding = bytes;
    }

    /// N never-written bytes held for exactly one `classifySheet` — the
    /// span the resident maximum falls in.
    pub fn setClassifyPadding(bytes: usize) void {
        calc.probe_classify_padding = bytes;
    }
};
pub const Worksheet = workbook_mod.Worksheet;
/// S6: the read-only pivot graph behind `Workbook.pivotTables`. The
/// module is exported for its fixture builder and `collect`, which
/// walks a bare `PartStore` for callers that hold no `Workbook`.
pub const pivots = @import("pivots.zig");
pub const Pivots = pivots.Pivots;
pub const PivotTable = pivots.PivotTable;
pub const PivotCache = pivots.PivotCache;
pub const PivotSourceResolution = pivots.SourceResolution;
pub const PivotSourceBounds = pivots.Bounds;
pub const PivotSourceUnresolved = pivots.Unresolved;
pub const ResolvedStyle = workbook_mod.ResolvedStyle;
pub const WorkbookError = workbook_mod.Error;
/// M6 (§12.2): what `Workbook.evaluate` returns — the CLI is the first
/// consumer outside this package that holds one.
pub const Evaluation = workbook_mod.Evaluation;
pub const EvaluateResult = workbook_mod.EvaluateResult;
pub const EvaluateOptions = workbook_mod.EvaluateOptions;
/// M6: the commit region's one seam (§5.7.9). Exported so a contract
/// test can inject a signal exactly between the rename and the
/// directory fsync instead of racing a timer against the transaction.
pub const CommitHook = store_mod.CommitHook;
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
/// Range-anchor (`twoCellAnchor`) surface, for
/// `Workbook.addImageRange` / `Workbook.addImageAnchored`.
pub const ImageCellRange = drawing_emit_mod.ImageCellRange;
pub const ImageAnchorSpec = drawing_emit_mod.ImageAnchorSpec;
/// Native-size sizing, for callers that want to inspect or override
/// what `Workbook.addImage` derives from an image header.
pub const ImageExtent = drawing_emit_mod.ImageExtent;
pub const ImagePixelSize = drawing_emit_mod.ImagePixelSize;
pub const imagePixelSize = drawing_emit_mod.pixelSize;
pub const imageNativeExtent = drawing_emit_mod.nativeExtent;
pub const emuFromPixels = drawing_emit_mod.emuFromPixels;

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

/// M5b2: §5.7.3 step 4's prepare/swap transaction. Exported because it
/// is the shape `recalculate()` and `saveWithRecalc` (M5d2) are built on
/// — `saveWithRecalc` has to hold a prepared candidate across a file
/// rename, which means the candidate is part of the package's surface
/// and not an implementation detail of one method.
pub const recalc_txn = @import("recalc_txn.zig");

/// M5d2: §5.7's pipeline behind `Workbook.recalculate` /
/// `Workbook.saveWithRecalc`. Exported for its `Options` and its
/// `Report` — a caller cannot name what the two methods take and return
/// otherwise — and for `prepare`, which M5d3's `writerSaveWithRecalc`
/// composes across a different serialization.
pub const recalc_run = @import("recalc_run.zig");
pub const RecalcOptions = recalc_run.Options;
pub const RecalcReport = recalc_run.Report;
pub const RunInputs = recalc_run.RunInputs;
pub const FormulaWrite = recalc_run.FormulaWrite;

/// M5d1: §5.5's cancellation / deadline seam. Re-exported because
/// `Workbook.openBufferControlled` takes a `Control` and a consumer of
/// this module has to be able to spell it — the same reason `Edit` is
/// here. `CancelToken` comes with it: a caller supplying a control has to
/// construct one of the two token arms.
pub const control = @import("zlsx_control");
pub const Control = control.Control;
pub const CancelToken = control.CancelToken;
pub const RetainedGeneration = workbook_mod.RetainedGeneration;
