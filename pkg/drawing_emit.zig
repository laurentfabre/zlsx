//! Greenfield drawing emit (B1 iter-wb-c2b).
//!
//! Scope: images anchored either at one cell (`oneCellAnchor` —
//! top-left pinned, sized by an explicit extent) or across a cell
//! range (`twoCellAnchor` — the range itself sizes the image). Every
//! anchor marker carries a pixel offset into its cell, and
//! `Workbook.addImage` derives the one-cell extent from the image's
//! own header at 96 DPI. One image per call; a sheet that already
//! carries a `<drawing>` element gets each further image spliced into
//! that existing drawing part as one more anchor — with a fresh `rId`
//! in the drawing's rels and a fresh `cNvPr` id — never a second
//! drawing part.
//!
//! Per OOXML spec (Part 1, Annex A) emitting a single image into a
//! drawing-less sheet requires touching FIVE artifacts:
//!
//!   1. `xl/media/imageN.{ext}`           — image bytes (STORED).
//!   2. `xl/drawings/drawingN.xml`        — wsDr/anchor wrapper.
//!   3. `xl/drawings/_rels/drawingN.xml.rels` — drawing → image link.
//!   4. `xl/worksheets/sheetN.xml`        — append `<drawing r:id=...>`.
//!   5. `xl/worksheets/_rels/sheetN.xml.rels` — sheet → drawing link.
//!
//! Appending into an existing drawing touches only artifacts 1–3: the
//! sheet's `<drawing r:id>` element and its rels entry already exist
//! and stay byte-identical.
//!
//! `PartStore.addPart` registers content types via the
//! `<Override>` mechanism; the image content type (`image/png` etc.)
//! is added the same way (the Office consumers we care about accept
//! per-image overrides over the `<Default Extension="png">` form, so
//! we don't need to touch the Defaults block).

const std = @import("std");
const assert = std.debug.assert;

const store_mod = @import("store.zig");
const PartStore = store_mod.PartStore;
const coords = @import("zlsx_refs");

pub const ImageMime = enum {
    png,
    jpeg,
    gif,

    pub fn contentType(self: ImageMime) []const u8 {
        return switch (self) {
            .png => "image/png",
            .jpeg => "image/jpeg",
            .gif => "image/gif",
        };
    }

    pub fn extension(self: ImageMime) []const u8 {
        return switch (self) {
            .png => "png",
            .jpeg => "jpg",
            .gif => "gif",
        };
    }
};

pub const Error = error{
    /// `bytes` magic header didn't match `mime`.
    MimeMagicMismatch,
    /// An anchor cell carried a 0 row or 0 col (OOXML is 1-based).
    InvalidAnchor,
    /// A range anchor's bottom-right cell sat above or left of its
    /// top-left cell. Emitting it would produce a negative-extent
    /// `<xdr:twoCellAnchor>`, which Office renders as an invisible or
    /// mirrored image rather than refusing the file.
    InvertedAnchorRange,
    /// The declared MIME's magic matched, but the header that follows
    /// it is truncated or carries a nonsensical size — so the image's
    /// native pixel dimensions can't be read. Also covers a 0-pixel
    /// axis (invisible in Office) and a PNG axis above the spec's
    /// 2^31-1 cap.
    UndecodableImageHeader,
    /// A `one_cell` extent had a 0 axis, or an axis above OOXML's
    /// `ST_PositiveCoordinate` maximum.
    InvalidExtent,
    /// `bytes` was empty.
    EmptyImage,
    /// An id-allocation scan (media part numbers, `rId`s, `cNvPr`
    /// ids) found the u32 id space exhausted — some existing entry
    /// already uses maxInt(u32). Crafted input; refused typed before
    /// any mutation rather than emitting a duplicate id.
    IdSpaceExhausted,
};

/// EMU per pixel at the 96 DPI that Office assumes for images without
/// their own resolution metadata: 914 400 EMU/inch ÷ 96 px/inch.
pub const emu_per_pixel_96dpi: u64 = 9525;

/// Upper bound of OOXML's `ST_PositiveCoordinate` (ECMA-376 Part 1,
/// §20.1.10.42) — about 29 837 inches. Extents above it are outside
/// the schema's value space, so Office rejects the part outright
/// rather than clamping.
pub const st_positive_coordinate_max: u64 = 27273042316900;

/// Convert a pixel count to EMU at 96 DPI.
///
/// The product always fits `u64`, but it only stays inside
/// `st_positive_coordinate_max` for `px <= 2_863_310_478`; a full
/// `u32` overshoots the schema cap by about 1.5×. `pixelSize` caps
/// every axis at 2^31-1, which is comfortably below that, so extents
/// derived from a decoded header are always in range. Callers passing
/// their own pixel counts should check the result against
/// `st_positive_coordinate_max` — `validateAnchor` does it for them.
pub fn emuFromPixels(px: u32) u64 {
    return @as(u64, px) * emu_per_pixel_96dpi;
}

/// 1-based (col, row), matching OOXML A1 conventions. Distinct from
/// `Workbook.CellRef` (which is row-major) because the only call site
/// is image anchoring and the OOXML drawing wire format puts col
/// first.
pub const ImageCellAnchor = struct {
    col: u32,
    row: u32,
    /// Horizontal offset into the anchor cell, in EMU. `u32` rather
    /// than a signed coordinate on purpose: a negative offset is
    /// schema-legal but places the image outside the cell it names,
    /// which no caller of a create-an-image API means to do. Every
    /// `u32` is a valid `ST_PositiveCoordinate`, so no range check is
    /// needed. Use `emuFromPixels` to convert from pixels at 96 DPI.
    col_off_emu: u32 = 0,
    /// Vertical offset into the anchor cell, in EMU. See
    /// `col_off_emu`.
    row_off_emu: u32 = 0,

    /// Build an anchor whose offsets are given in pixels at 96 DPI,
    /// which is how callers usually think about nudging an image
    /// inside its cell.
    ///
    /// Pixels are `u16` so the conversion is total: 65 535 px is
    /// 624 220 875 EMU, comfortably inside `u32`, so this can neither
    /// overflow nor need a cast at the call site. Offsets beyond that
    /// exceed any real cell and can be set on the fields directly.
    pub fn atPixelOffset(col: u32, row: u32, col_px: u16, row_px: u16) ImageCellAnchor {
        return .{
            .col = col,
            .row = row,
            .col_off_emu = @intCast(emuFromPixels(col_px)),
            .row_off_emu = @intCast(emuFromPixels(row_px)),
        };
    }
};

/// Rendered size of a `oneCellAnchor` image, in EMU — the
/// `<xdr:ext cx=… cy=…>` pair. `twoCellAnchor` images carry no extent
/// because the cell range supplies it.
pub const ImageExtent = struct {
    cx_emu: u64,
    cy_emu: u64,
};

/// An image's intrinsic size in pixels, as recorded in its own header.
pub const ImagePixelSize = struct {
    width: u32,
    height: u32,

    /// This size as a `oneCellAnchor` extent at 96 DPI.
    pub fn extent(self: ImagePixelSize) ImageExtent {
        return .{
            .cx_emu = emuFromPixels(self.width),
            .cy_emu = emuFromPixels(self.height),
        };
    }
};

/// Inclusive 1-based cell range for a `twoCellAnchor` image: the
/// picture covers `from` through `to` in full, the way a spreadsheet
/// user reads "B2:D5". Both endpoints are `ImageCellAnchor`s, so a
/// 1×1 range (`from` == `to`) is legal and covers exactly that cell.
///
/// **Wire mapping.** OOXML anchor markers are 0-based *grid corners*,
/// not cell indices: `<xdr:from>` is the top-left corner of the first
/// covered cell (`from.col - 1`) and `<xdr:to>` is the top-left corner
/// of the cell *after* the last covered one (`to.col`, no decrement).
/// The asymmetry is what makes an inclusive range come out right — see
/// `buildDrawingXml`.
pub const ImageCellRange = struct {
    from: ImageCellAnchor,
    to: ImageCellAnchor,
};

/// Which OOXML anchor wrapper an image gets. `Workbook.addImage` and
/// `Workbook.addImageRange` are thin wrappers that each pick one
/// variant; `Workbook.addImageAnchored` takes the union directly.
pub const ImageAnchorSpec = union(enum) {
    /// `<xdr:oneCellAnchor>` — pinned at one cell at an explicit
    /// extent. The image does not resize when the anchor cell's
    /// row/col does.
    ///
    /// The extent lives in the payload rather than in a separate
    /// parameter so that the one form that needs it cannot be built
    /// without it, and the form that must not carry it cannot be
    /// handed one. `Workbook.addImage` fills `ext` from
    /// `nativeExtent`; callers wanting a deliberate size pass their
    /// own through `Workbook.addImageAnchored`.
    one_cell: struct {
        at: ImageCellAnchor,
        ext: ImageExtent,
    },
    /// `<xdr:twoCellAnchor>` — spans a cell range and resizes with it
    /// (the spec's default `editAs="twoCell"` behaviour, which we get
    /// by omitting the attribute).
    two_cell: ImageCellRange,
};

/// Reject an anchor cell the sheet grid can't hold: 0 row/col (the
/// wire format is 1-based on our side, 0-based on its own, so 0 has
/// no valid pre-image) or a coordinate past Excel's XFD / 1_048_576
/// grid caps. The grid caps also keep every emitted wire coordinate
/// inside `xsd:int` — the marker types' value space — with room to
/// spare (a `to` corner emits `col`, at most 16_384).
pub fn validateAnchorCell(at: ImageCellAnchor) Error!void {
    if (at.col == 0 or at.row == 0) return Error.InvalidAnchor;
    if (at.col > coords.max_col_1based or at.row > coords.max_row) {
        return Error.InvalidAnchor;
    }
}

/// Reject anchors OOXML can't express: cells outside the sheet grid
/// (see `validateAnchorCell`) and ranges whose bottom-right precedes
/// their top-left.
///
/// Callers must run this before `buildDrawingXml`, which asserts the
/// same invariants rather than re-checking them.
pub fn validateAnchor(spec: ImageAnchorSpec) Error!void {
    switch (spec) {
        .one_cell => |o| {
            try validateAnchorCell(o.at);
            // A 0 axis renders as an invisible picture; an axis past
            // the schema cap makes Office reject the whole part.
            if (o.ext.cx_emu == 0 or o.ext.cy_emu == 0) return Error.InvalidExtent;
            if (o.ext.cx_emu > st_positive_coordinate_max or
                o.ext.cy_emu > st_positive_coordinate_max)
            {
                return Error.InvalidExtent;
            }
        },
        .two_cell => |r| {
            try validateAnchorCell(r.from);
            try validateAnchorCell(r.to);
            if (r.to.col < r.from.col or r.to.row < r.from.row) {
                return Error.InvertedAnchorRange;
            }
        },
    }
}

/// Validate the magic header for `mime` against the leading bytes of
/// `bytes`. Defensive: callers pass a declared MIME plus a byte
/// buffer; if the two disagree, the resulting xlsx will look fine in
/// the package metadata but Office may render a broken-image
/// placeholder. Cheaper to reject up front.
pub fn validateMagic(mime: ImageMime, bytes: []const u8) Error!void {
    if (bytes.len == 0) return Error.EmptyImage;
    switch (mime) {
        .png => {
            // 89 50 4E 47 0D 0A 1A 0A
            const sig = [_]u8{ 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A };
            if (bytes.len < sig.len) return Error.MimeMagicMismatch;
            if (!std.mem.eql(u8, bytes[0..sig.len], &sig)) return Error.MimeMagicMismatch;
        },
        .jpeg => {
            // FF D8 FF
            if (bytes.len < 3) return Error.MimeMagicMismatch;
            if (bytes[0] != 0xFF or bytes[1] != 0xD8 or bytes[2] != 0xFF) {
                return Error.MimeMagicMismatch;
            }
        },
        .gif => {
            // 47 49 46 (GIF)
            if (bytes.len < 3) return Error.MimeMagicMismatch;
            if (bytes[0] != 'G' or bytes[1] != 'I' or bytes[2] != 'F') {
                return Error.MimeMagicMismatch;
            }
        },
    }
}

/// PNG caps each axis at 2^31-1 (spec §11.2.2). We reuse the bound
/// for every format: it is far above any real image, and holding to
/// it keeps `emuFromPixels` inside `st_positive_coordinate_max`.
const pixel_axis_max: u32 = std.math.maxInt(i32);

/// Shared tail of the three header readers: a size is only usable if
/// both axes are non-zero (a 0 axis renders as nothing) and within
/// `pixel_axis_max`.
fn checkedPixelSize(width: u32, height: u32) Error!ImagePixelSize {
    if (width == 0 or height == 0) return Error.UndecodableImageHeader;
    if (width > pixel_axis_max or height > pixel_axis_max) {
        return Error.UndecodableImageHeader;
    }
    return .{ .width = width, .height = height };
}

/// PNG: IHDR is mandated to be the first chunk, at a fixed offset —
/// 8-byte signature, 4-byte chunk length, 4-byte type, then two
/// big-endian u32 axes. No chunk walk needed.
fn pngPixelSize(bytes: []const u8) Error!ImagePixelSize {
    if (bytes.len < 24) return Error.UndecodableImageHeader;
    if (!std.mem.eql(u8, bytes[12..16], "IHDR")) return Error.UndecodableImageHeader;
    return checkedPixelSize(
        std.mem.readInt(u32, bytes[16..20], .big),
        std.mem.readInt(u32, bytes[20..24], .big),
    );
}

/// GIF: the logical screen descriptor follows the 6-byte header, and
/// its first four bytes are the canvas size as two little-endian u16s.
/// Frames may be smaller, but the canvas is what Office lays out.
fn gifPixelSize(bytes: []const u8) Error!ImagePixelSize {
    if (bytes.len < 10) return Error.UndecodableImageHeader;
    return checkedPixelSize(
        std.mem.readInt(u16, bytes[6..8], .little),
        std.mem.readInt(u16, bytes[8..10], .little),
    );
}

/// True for SOF0..SOF15 — the frame headers that carry a size —
/// excluding the three markers that share the range but do not:
/// `C4` DHT, `C8` JPG, `CC` DAC.
fn isStartOfFrame(marker: u8) bool {
    return switch (marker) {
        0xC0...0xC3, 0xC5...0xC7, 0xC9...0xCB, 0xCD...0xCF => true,
        else => false,
    };
}

/// JPEG: unlike PNG and GIF the size lives at no fixed offset, so walk
/// the marker chain from just past SOI until a frame header turns up.
/// Its payload is precision(1), height(2), width(2) — height first,
/// which is the classic place to get this backwards.
fn jpegPixelSize(bytes: []const u8) Error!ImagePixelSize {
    var i: usize = 2;
    while (i + 1 < bytes.len) {
        if (bytes[i] != 0xFF) return Error.UndecodableImageHeader;

        // A marker may be preceded by any number of 0xFF fill bytes.
        var marker = bytes[i + 1];
        i += 2;
        while (marker == 0xFF) {
            if (i >= bytes.len) return Error.UndecodableImageHeader;
            marker = bytes[i];
            i += 1;
        }

        switch (marker) {
            // Standalone: no length field, no payload.
            0x01, 0xD0...0xD7 => continue,
            // SOI cannot recur mid-stream, and reaching EOI or the
            // start of scan data means no frame header ever came.
            0xD8, 0xD9, 0xDA => return Error.UndecodableImageHeader,
            else => {},
        }

        if (i + 2 > bytes.len) return Error.UndecodableImageHeader;
        const seg_len = std.mem.readInt(u16, bytes[i..][0..2], .big);
        // The length field counts itself, so anything under 2 is junk.
        if (seg_len < 2) return Error.UndecodableImageHeader;
        const payload = bytes[i + 2 ..];
        const payload_len = seg_len - 2;
        if (payload.len < payload_len) return Error.UndecodableImageHeader;

        if (isStartOfFrame(marker)) {
            if (payload_len < 5) return Error.UndecodableImageHeader;
            return checkedPixelSize(
                std.mem.readInt(u16, payload[3..5], .big),
                std.mem.readInt(u16, payload[1..3], .big),
            );
        }
        i += seg_len;
    }
    return Error.UndecodableImageHeader;
}

/// Read an image's intrinsic pixel size from its own header. Reads
/// headers only — no pixel data is decoded, so this stays O(header)
/// and pulls in no codec.
///
/// The declared `mime` is checked against the magic first, so a
/// mismatch surfaces as `MimeMagicMismatch` rather than as a confusing
/// `UndecodableImageHeader` from the wrong parser.
pub fn pixelSize(mime: ImageMime, bytes: []const u8) Error!ImagePixelSize {
    try validateMagic(mime, bytes);
    return switch (mime) {
        .png => pngPixelSize(bytes),
        .jpeg => jpegPixelSize(bytes),
        .gif => gifPixelSize(bytes),
    };
}

/// The `oneCellAnchor` extent for `bytes` at its native size, mapping
/// pixels to EMU at 96 DPI.
///
/// Office's own "insert picture" does the same thing for images that
/// carry no resolution metadata. We do it unconditionally: PNG `pHYs`
/// and JPEG JFIF density fields exist but are unreliable in the wild,
/// and honouring them only sometimes would make the emitted size
/// depend on which encoder produced the file.
pub fn nativeExtent(mime: ImageMime, bytes: []const u8) Error!ImageExtent {
    return (try pixelSize(mime, bytes)).extent();
}

/// Walk every part name in the store and return the lowest unused
/// `<prefix>N<suffix>` integer >= 1. Used for `imageN.png`,
/// `drawingN.xml` etc. — the OOXML convention for autogenerated
/// part names. We don't try to be clever about gap-filling: bumping
/// past the highest seen N keeps the result stable across saves.
pub fn nextFreeNumber(
    store: *const PartStore,
    prefix: []const u8,
    suffix: []const u8,
) Error!u32 {
    assert(prefix.len > 0);
    var max_seen: u32 = 0;
    for (store.parts) |p| {
        if (!std.mem.startsWith(u8, p.name, prefix)) continue;
        if (!std.mem.endsWith(u8, p.name, suffix)) continue;
        const middle = p.name[prefix.len .. p.name.len - suffix.len];
        // Leading digits only: the image call passes an empty suffix
        // (Excel numbers media across extensions, so `.png` and `.jpg`
        // must share one number space), which leaves the extension in
        // `middle` ("1.png"). A middle with no leading digit isn't a
        // numbered slot at all.
        var digits: usize = 0;
        while (digits < middle.len and std.ascii.isDigit(middle[digits])) digits += 1;
        if (digits == 0) continue;
        const n = std.fmt.parseInt(u32, middle[0..digits], 10) catch continue;
        if (n > max_seen) max_seen = n;
    }
    // A part already numbered maxInt(u32) is crafted input; refuse
    // typed rather than trapping in safe builds or wrapping into a
    // duplicate name in fast ones.
    if (max_seen == std.math.maxInt(u32)) return Error.IdSpaceExhausted;
    return max_seen + 1;
}

/// Emit one `<xdr:from>` / `<xdr:to>` marker: the four scalar
/// children OOXML requires, in schema order. `col0` / `row0` are
/// already in wire (0-based) coordinates — the caller owns the
/// inclusive→corner conversion — while the offsets pass through as
/// given.
fn appendAnchorMarker(
    allocator: std.mem.Allocator,
    buf: *std.ArrayListUnmanaged(u8),
    tag: []const u8,
    col0: u32,
    row0: u32,
    anchor: ImageCellAnchor,
) !void {
    assert(tag.len > 0);

    var num_buf: [32]u8 = undefined;
    try buf.append(allocator, '<');
    try buf.appendSlice(allocator, tag);
    try buf.append(allocator, '>');
    try buf.appendSlice(allocator, "<xdr:col>");
    try buf.appendSlice(allocator, try std.fmt.bufPrint(&num_buf, "{d}", .{col0}));
    try buf.appendSlice(allocator, "</xdr:col><xdr:colOff>");
    try buf.appendSlice(allocator, try std.fmt.bufPrint(&num_buf, "{d}", .{anchor.col_off_emu}));
    try buf.appendSlice(allocator, "</xdr:colOff>");
    try buf.appendSlice(allocator, "<xdr:row>");
    try buf.appendSlice(allocator, try std.fmt.bufPrint(&num_buf, "{d}", .{row0}));
    try buf.appendSlice(allocator, "</xdr:row><xdr:rowOff>");
    try buf.appendSlice(allocator, try std.fmt.bufPrint(&num_buf, "{d}", .{anchor.row_off_emu}));
    try buf.appendSlice(allocator, "</xdr:rowOff>");
    try buf.appendSlice(allocator, "</");
    try buf.appendSlice(allocator, tag);
    try buf.append(allocator, '>');
}

/// OOXML namespace URIs for the three prefixes emitted anchors use.
/// Transitional (ECMA-376 1st edition) and Strict (2nd edition) are
/// the same logical names under different URIs; `pkg/drawings.zig`
/// accepts both on the read side, and the append path must emit
/// whichever dialect the host drawing part already speaks — a
/// Transitional child inside a Strict `wsDr` (or vice versa) is a
/// foreign-namespace element the host schema doesn't admit.
pub const ns_xdr_transitional = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
pub const ns_xdr_strict = "http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing";
pub const ns_a_transitional = "http://schemas.openxmlformats.org/drawingml/2006/main";
pub const ns_a_strict = "http://purl.oclc.org/ooxml/drawingml/main";
pub const ns_r_transitional = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
pub const ns_r_strict = "http://purl.oclc.org/ooxml/officeDocument/relationships";
pub const ns_main_transitional = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
pub const ns_main_strict = "http://purl.oclc.org/ooxml/spreadsheetml/main";
pub const image_rel_type_transitional = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
pub const image_rel_type_strict = "http://purl.oclc.org/ooxml/officeDocument/relationships/image";
pub const drawing_rel_type_transitional = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
pub const drawing_rel_type_strict = "http://purl.oclc.org/ooxml/officeDocument/relationships/drawing";

/// Which ECMA-376 conformance dialect a part speaks, read off a root
/// namespace declaration by `detectWsDrDialect` / `detectSheetDialect`.
pub const WsDrDialect = enum {
    transitional,
    strict,

    pub fn nsXdr(self: WsDrDialect) []const u8 {
        return switch (self) {
            .transitional => ns_xdr_transitional,
            .strict => ns_xdr_strict,
        };
    }

    pub fn nsA(self: WsDrDialect) []const u8 {
        return switch (self) {
            .transitional => ns_a_transitional,
            .strict => ns_a_strict,
        };
    }

    pub fn nsR(self: WsDrDialect) []const u8 {
        return switch (self) {
            .transitional => ns_r_transitional,
            .strict => ns_r_strict,
        };
    }

    pub fn imageRelType(self: WsDrDialect) []const u8 {
        return switch (self) {
            .transitional => image_rel_type_transitional,
            .strict => image_rel_type_strict,
        };
    }

    pub fn drawingRelType(self: WsDrDialect) []const u8 {
        return switch (self) {
            .transitional => drawing_rel_type_transitional,
            .strict => drawing_rel_type_strict,
        };
    }
};

/// Index of the document's root start-tag `<`, skipping the XML
/// declaration, processing instructions, comments, a doctype, and
/// whitespace. Null when no element follows the prolog. Keeps a
/// commented-out pseudo-root (`<!-- <xdr:wsDr …> -->`) from being
/// mistaken for the real one.
fn rootElementStart(xml: []const u8) ?usize {
    var i: usize = 0;
    while (i < xml.len) {
        if (std.ascii.isWhitespace(xml[i])) {
            i += 1;
            continue;
        }
        if (xml[i] != '<') return null;
        if (std.mem.startsWith(u8, xml[i..], "<?")) {
            const end = std.mem.indexOfPos(u8, xml, i, "?>") orelse return null;
            i = end + 2;
            continue;
        }
        if (std.mem.startsWith(u8, xml[i..], "<!--")) {
            const end = std.mem.indexOfPos(u8, xml, i, "-->") orelse return null;
            i = end + 3;
            continue;
        }
        if (std.mem.startsWith(u8, xml[i..], "<!")) {
            // Doctype — never present in OOXML parts; skipping to the
            // first `>` misses an internal subset, which is beyond
            // what this writer will chase.
            const end = std.mem.indexOfScalarPos(u8, xml, i, '>') orelse return null;
            i = end + 1;
            continue;
        }
        return i;
    }
    return null;
}

/// Read a host drawing's conformance dialect off its ROOT element's
/// `xmlns:xdr` declaration — the actual document root, located past
/// comments and PIs, so a commented-out pseudo-root can't spoof the
/// answer. Null when the root isn't `<xdr:wsDr>` or binds `xdr` to
/// neither known URI — the caller refuses rather than splicing
/// content the host schema can't admit.
pub fn detectWsDrDialect(drawing_xml: []const u8) ?WsDrDialect {
    const root = rootElementStart(drawing_xml) orelse return null;
    if (!std.mem.startsWith(u8, drawing_xml[root..], "<xdr:wsDr")) return null;
    const after_idx = root + "<xdr:wsDr".len;
    // A root using the prefix without declaring it (`>` right away)
    // is malformed; a longer name (`<xdr:wsDrX`) is another element.
    if (after_idx >= drawing_xml.len) return null;
    if (!std.ascii.isWhitespace(drawing_xml[after_idx])) return null;
    const el_end = store_mod.xmlStartTagEnd(drawing_xml, root) orelse return null;
    const uri = store_mod.xmlAttrValue(drawing_xml[root..el_end], "xmlns:xdr") orelse return null;
    if (std.mem.eql(u8, uri, ns_xdr_transitional)) return .transitional;
    if (std.mem.eql(u8, uri, ns_xdr_strict)) return .strict;
    return null;
}

/// Read a worksheet part's conformance dialect off its root's default
/// namespace declaration, so the fresh-image path can emit drawing
/// parts and relationship Types in the workbook's own dialect.
/// Defaults to `.transitional` when the root is unrecognized: the
/// sheet already round-tripped through zlsx's own parser to get here,
/// and Transitional is what every mainstream producer writes.
pub fn detectSheetDialect(sheet_xml: []const u8) WsDrDialect {
    const root = rootElementStart(sheet_xml) orelse return .transitional;
    if (!std.mem.startsWith(u8, sheet_xml[root..], "<worksheet")) return .transitional;
    const el_end = store_mod.xmlStartTagEnd(sheet_xml, root) orelse return .transitional;
    const uri = store_mod.xmlAttrValue(sheet_xml[root..el_end], "xmlns") orelse return .transitional;
    if (std.mem.eql(u8, uri, ns_main_strict)) return .strict;
    return .transitional;
}

/// Whether a built anchor declares its namespace bindings inline.
/// `.inherit` relies on the enclosing document's root — correct for
/// the fresh path, which writes that root itself. `.declare` pins all
/// three prefixes on the anchor element in the given dialect, so a
/// splice into a host that doesn't declare `a:`/`r:` at its root (or
/// binds them to other URIs) stays well-formed and schema-valid;
/// where the host already declares them, the local re-declaration is
/// redundant but conformant.
pub const NsDecl = union(enum) {
    inherit,
    declare: WsDrDialect,
};

/// Build one anchor element (`<xdr:oneCellAnchor>` /
/// `<xdr:twoCellAnchor>`) for a single image — the unit that both a
/// fresh drawing part and an append into an existing one are made of.
/// Every marker carries its anchor's pixel offsets.
///
///   - `.one_cell`: `<xdr:oneCellAnchor>` at `(col-1, row-1)` sized by
///     the payload's `ext`.
///   - `.two_cell`: `<xdr:twoCellAnchor>` from `(from.col-1,
///     from.row-1)` to `(to.col, to.row)`. The `to` marker is *not*
///     decremented: OOXML anchor markers name grid corners, so the
///     corner that closes an inclusive range is the top-left of the
///     next cell over. No `<xdr:ext>` — the range sizes the image.
///
/// `rel_id` is the image relationship in the OWNING DRAWING's rels
/// (`r:embed`), and `pic_id` the `cNvPr` id — unique per drawing part,
/// so appenders must allocate it with `nextFreeCnvprId`. The shape
/// name is derived as `Picture {pic_id}`, mirroring Excel.
///
/// Caller must have run `validateAnchor` first; the invariants are
/// asserted here, not re-checked.
pub fn buildPicAnchorXml(
    allocator: std.mem.Allocator,
    spec: ImageAnchorSpec,
    rel_id: []const u8,
    pic_id: u32,
    ns: NsDecl,
) ![]u8 {
    assert(rel_id.len > 0);
    assert(pic_id >= 1);
    switch (spec) {
        .one_cell => |o| {
            assert(o.at.col >= 1);
            assert(o.at.row >= 1);
            assert(o.ext.cx_emu >= 1);
            assert(o.ext.cy_emu >= 1);
            assert(o.ext.cx_emu <= st_positive_coordinate_max);
            assert(o.ext.cy_emu <= st_positive_coordinate_max);
        },
        .two_cell => |r| {
            assert(r.from.col >= 1);
            assert(r.from.row >= 1);
            assert(r.to.col >= r.from.col);
            assert(r.to.row >= r.from.row);
        },
    }

    var buf: std.ArrayListUnmanaged(u8) = .empty;
    defer buf.deinit(allocator);

    // `editAs` is omitted on the two-cell form: the spec's default is
    // `twoCell` (move AND size with the cells), which is the whole
    // point of anchoring to a range.
    const wrapper = switch (spec) {
        .one_cell => "xdr:oneCellAnchor",
        .two_cell => "xdr:twoCellAnchor",
    };
    try buf.append(allocator, '<');
    try buf.appendSlice(allocator, wrapper);
    switch (ns) {
        .inherit => {},
        .declare => |d| {
            try buf.appendSlice(allocator, " xmlns:xdr=\"");
            try buf.appendSlice(allocator, d.nsXdr());
            try buf.appendSlice(allocator, "\" xmlns:a=\"");
            try buf.appendSlice(allocator, d.nsA());
            try buf.appendSlice(allocator, "\" xmlns:r=\"");
            try buf.appendSlice(allocator, d.nsR());
            try buf.append(allocator, '"');
        },
    }
    try buf.append(allocator, '>');

    switch (spec) {
        .one_cell => |o| {
            try appendAnchorMarker(allocator, &buf, "xdr:from", o.at.col - 1, o.at.row - 1, o.at);
            var ext_buf: [64]u8 = undefined;
            try buf.appendSlice(allocator, try std.fmt.bufPrint(
                &ext_buf,
                "<xdr:ext cx=\"{d}\" cy=\"{d}\"/>",
                .{ o.ext.cx_emu, o.ext.cy_emu },
            ));
        },
        .two_cell => |r| {
            try appendAnchorMarker(allocator, &buf, "xdr:from", r.from.col - 1, r.from.row - 1, r.from);
            try appendAnchorMarker(allocator, &buf, "xdr:to", r.to.col, r.to.row, r.to);
        },
    }

    var id_buf: [16]u8 = undefined;
    const id_digits = try std.fmt.bufPrint(&id_buf, "{d}", .{pic_id});

    try buf.appendSlice(allocator, "<xdr:pic>");
    try buf.appendSlice(allocator, "<xdr:nvPicPr><xdr:cNvPr id=\"");
    try buf.appendSlice(allocator, id_digits);
    try buf.appendSlice(allocator, "\" name=\"Picture ");
    try buf.appendSlice(allocator, id_digits);
    try buf.appendSlice(allocator, "\"/><xdr:cNvPicPr/></xdr:nvPicPr>");
    try buf.appendSlice(allocator, "<xdr:blipFill><a:blip r:embed=\"");
    try buf.appendSlice(allocator, rel_id);
    try buf.appendSlice(allocator, "\"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>");
    try buf.appendSlice(
        allocator,
        "<xdr:spPr><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></xdr:spPr>",
    );
    try buf.appendSlice(allocator, "</xdr:pic>");
    try buf.appendSlice(allocator, "<xdr:clientData/>");
    try buf.appendSlice(allocator, "</");
    try buf.appendSlice(allocator, wrapper);
    try buf.append(allocator, '>');

    return try buf.toOwnedSlice(allocator);
}

/// Build a fresh drawingN.xml body holding a single image anchor. The
/// first image in a fresh part always gets `rId1` / `cNvPr id="1"` —
/// appends into the part allocate theirs with `nextFreeRelId` /
/// `nextFreeCnvprId` instead. The root declares all three prefixes in
/// `dialect`'s URIs (the fresh path derives it from the host sheet
/// via `detectSheetDialect`), and the anchor inherits them.
///
/// Caller must have run `validateAnchor` first.
pub fn buildDrawingXml(
    allocator: std.mem.Allocator,
    spec: ImageAnchorSpec,
    dialect: WsDrDialect,
) ![]u8 {
    const anchor_xml = try buildPicAnchorXml(allocator, spec, "rId1", 1, .inherit);
    defer allocator.free(anchor_xml);

    var buf: std.ArrayListUnmanaged(u8) = .empty;
    defer buf.deinit(allocator);

    try buf.appendSlice(allocator, "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
    try buf.appendSlice(allocator, "<xdr:wsDr xmlns:xdr=\"");
    try buf.appendSlice(allocator, dialect.nsXdr());
    try buf.appendSlice(allocator, "\" xmlns:a=\"");
    try buf.appendSlice(allocator, dialect.nsA());
    try buf.appendSlice(allocator, "\" xmlns:r=\"");
    try buf.appendSlice(allocator, dialect.nsR());
    try buf.appendSlice(allocator, "\">");
    try buf.appendSlice(allocator, anchor_xml);
    try buf.appendSlice(allocator, "</xdr:wsDr>");

    return try buf.toOwnedSlice(allocator);
}

/// Splice a pre-built anchor element into an existing drawing part,
/// immediately before its closing `</xdr:wsDr>`. Caller owns the
/// returned bytes.
///
/// `error.UnsupportedDrawingPart` when no `</xdr:wsDr>` exists — the
/// part is not a spreadsheet-drawing wrapper this writer knows how to
/// extend (a legacy VML part, a chartSpace root, or a wsDr bound to a
/// prefix other than `xdr`).
pub fn appendAnchorToDrawing(
    allocator: std.mem.Allocator,
    drawing_xml: []const u8,
    anchor_xml: []const u8,
) ![]u8 {
    assert(anchor_xml.len > 0);

    const close_tag = "</xdr:wsDr>";
    const close_pos = std.mem.lastIndexOf(u8, drawing_xml, close_tag) orelse
        return error.UnsupportedDrawingPart;

    var buf: std.ArrayListUnmanaged(u8) = .empty;
    defer buf.deinit(allocator);

    try buf.appendSlice(allocator, drawing_xml[0..close_pos]);
    try buf.appendSlice(allocator, anchor_xml);
    try buf.appendSlice(allocator, drawing_xml[close_pos..]);

    return try buf.toOwnedSlice(allocator);
}

/// Build the drawing-rels file linking `rId1` to `../media/imageN.{ext}`.
///
/// `image_type_uri` is the relationship Type — Transitional or Strict
/// to match the owning drawing's dialect (`WsDrDialect.imageRelType`).
/// The wrapping package-relationships namespace is OPC (ISO 29500
/// Part 2) and does not vary with the Part 1 dialect.
pub fn buildDrawingRels(
    allocator: std.mem.Allocator,
    image_part_basename: []const u8,
    image_type_uri: []const u8,
) ![]u8 {
    assert(image_part_basename.len > 0);
    assert(image_type_uri.len > 0);

    var buf: std.ArrayListUnmanaged(u8) = .empty;
    defer buf.deinit(allocator);

    try buf.appendSlice(
        allocator,
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" ++
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">",
    );
    try buf.appendSlice(allocator, "<Relationship Id=\"rId1\" Type=\"");
    try buf.appendSlice(allocator, image_type_uri);
    try buf.appendSlice(allocator, "\" Target=\"../media/");
    try buf.appendSlice(allocator, image_part_basename);
    try buf.appendSlice(allocator, "\"/>");
    try buf.appendSlice(allocator, "</Relationships>");

    return try buf.toOwnedSlice(allocator);
}

/// Build the bytes for a brand-new `xl/worksheets/_rels/sheetN.xml.rels`
/// containing exactly one Relationship to `../drawings/drawingN.xml`.
/// `drawing_type_uri` follows the sheet's dialect
/// (`WsDrDialect.drawingRelType`); the OPC wrapper namespace does not
/// vary with it.
pub fn buildFreshSheetRels(
    allocator: std.mem.Allocator,
    drawing_basename: []const u8,
    rel_id: []const u8,
    drawing_type_uri: []const u8,
) ![]u8 {
    assert(drawing_basename.len > 0);
    assert(rel_id.len > 0);
    assert(drawing_type_uri.len > 0);

    var buf: std.ArrayListUnmanaged(u8) = .empty;
    defer buf.deinit(allocator);

    try buf.appendSlice(
        allocator,
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" ++
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">",
    );
    try buf.appendSlice(allocator, "<Relationship Id=\"");
    try buf.appendSlice(allocator, rel_id);
    try buf.appendSlice(allocator, "\" Type=\"");
    try buf.appendSlice(allocator, drawing_type_uri);
    try buf.appendSlice(allocator, "\" Target=\"../drawings/");
    try buf.appendSlice(allocator, drawing_basename);
    try buf.appendSlice(allocator, "\"/>");
    try buf.appendSlice(allocator, "</Relationships>");

    return try buf.toOwnedSlice(allocator);
}

/// Splice a `<Relationship>` element into existing rels XML before
/// `</Relationships>`. `drawing_type_uri` follows the sheet's dialect
/// (see `buildFreshSheetRels`). Caller owns the returned bytes.
pub fn appendRelationship(
    allocator: std.mem.Allocator,
    existing_xml: []const u8,
    drawing_basename: []const u8,
    rel_id: []const u8,
    drawing_type_uri: []const u8,
) ![]u8 {
    assert(drawing_type_uri.len > 0);
    const close_tag = "</Relationships>";
    const close_pos = std.mem.lastIndexOf(u8, existing_xml, close_tag) orelse
        return error.MalformedSheetRels;

    var buf: std.ArrayListUnmanaged(u8) = .empty;
    defer buf.deinit(allocator);

    try buf.appendSlice(allocator, existing_xml[0..close_pos]);
    try buf.appendSlice(allocator, "<Relationship Id=\"");
    try buf.appendSlice(allocator, rel_id);
    try buf.appendSlice(allocator, "\" Type=\"");
    try buf.appendSlice(allocator, drawing_type_uri);
    try buf.appendSlice(allocator, "\" Target=\"../drawings/");
    try buf.appendSlice(allocator, drawing_basename);
    try buf.appendSlice(allocator, "\"/>");
    try buf.appendSlice(allocator, existing_xml[close_pos..]);

    return try buf.toOwnedSlice(allocator);
}

/// Splice an image `<Relationship>` into an existing drawing-rels file
/// before `</Relationships>` — the sibling of `appendRelationship` for
/// the drawing → image edge of the graph. `image_type_uri` follows the
/// owning drawing's dialect (see `buildDrawingRels`). Caller owns the
/// returned bytes.
pub fn appendImageRelationship(
    allocator: std.mem.Allocator,
    existing_xml: []const u8,
    image_basename: []const u8,
    rel_id: []const u8,
    image_type_uri: []const u8,
) ![]u8 {
    assert(image_basename.len > 0);
    assert(rel_id.len > 0);
    assert(image_type_uri.len > 0);

    const close_tag = "</Relationships>";
    const close_pos = std.mem.lastIndexOf(u8, existing_xml, close_tag) orelse
        return error.MalformedDrawingRels;

    var buf: std.ArrayListUnmanaged(u8) = .empty;
    defer buf.deinit(allocator);

    try buf.appendSlice(allocator, existing_xml[0..close_pos]);
    try buf.appendSlice(allocator, "<Relationship Id=\"");
    try buf.appendSlice(allocator, rel_id);
    try buf.appendSlice(allocator, "\" Type=\"");
    try buf.appendSlice(allocator, image_type_uri);
    try buf.appendSlice(allocator, "\" Target=\"../media/");
    try buf.appendSlice(allocator, image_basename);
    try buf.appendSlice(allocator, "\"/>");
    try buf.appendSlice(allocator, existing_xml[close_pos..]);

    return try buf.toOwnedSlice(allocator);
}

/// Splice `<drawing r:id="..."/>` into a sheet's XML, immediately
/// before `</worksheet>`. Caller owns the returned bytes.
pub fn appendDrawingElementToSheet(
    allocator: std.mem.Allocator,
    sheet_xml: []const u8,
    rel_id: []const u8,
) ![]u8 {
    assert(rel_id.len > 0);

    const close_tag = "</worksheet>";
    const close_pos = std.mem.lastIndexOf(u8, sheet_xml, close_tag) orelse
        return error.MalformedSheetXml;

    var buf: std.ArrayListUnmanaged(u8) = .empty;
    defer buf.deinit(allocator);

    try buf.appendSlice(allocator, sheet_xml[0..close_pos]);
    // The `r:` namespace prefix is bound in the worksheet root by
    // every OOXML producer we've round-tripped (xmlns:r=...
    // /relationships); we emit it without re-declaring.
    try buf.appendSlice(allocator, "<drawing r:id=\"");
    try buf.appendSlice(allocator, rel_id);
    try buf.appendSlice(allocator, "\"/>");
    try buf.appendSlice(allocator, sheet_xml[close_pos..]);

    return try buf.toOwnedSlice(allocator);
}

/// Pick the next free `rId<N>` for the given existing rels list.
/// Returns the integer N (caller formats `rId<N>`). Stable across
/// saves: bumps past max(N) rather than gap-filling. A crafted
/// `rId4294967295` exhausts the space — refused typed rather than
/// trapping or duplicating.
///
/// Callers holding raw rels-file bytes rather than the store's cached
/// view (the append path — a rels part created by `addPart` earlier
/// in the same session is absent from the cache, which only
/// `replacePart` refreshes) parse them first with
/// `store.parseRelationships`, the same quote-tolerant,
/// entity-decoding parser that fills the cache.
pub fn nextFreeRelId(rels: []const store_mod.Relationship) Error!u32 {
    var max_seen: u32 = 0;
    for (rels) |r| {
        if (!std.mem.startsWith(u8, r.id, "rId")) continue;
        const n = std.fmt.parseInt(u32, r.id[3..], 10) catch continue;
        if (n > max_seen) max_seen = n;
    }
    if (max_seen == std.math.maxInt(u32)) return Error.IdSpaceExhausted;
    return max_seen + 1;
}

/// Extract the `r:id` of a worksheet's `<drawing>` element, or null
/// when the sheet has none (or carries one without an `r:id`, which
/// is schema-invalid and unresolvable either way). Returns a slice
/// into `sheet_xml`. Attribute reads go through the shared
/// `store.xmlAttrValue` lexer (both quote styles, `Eq` whitespace,
/// `=`/`>`-in-value safe).
///
/// The name test requires XML whitespace after `<drawing` so that
/// `<drawingHF>` (header/footer images) and a bare `<drawing/>` never
/// match — only an attribute-carrying worksheet drawing element does.
pub fn extractDrawingRelId(sheet_xml: []const u8) ?[]const u8 {
    var pos: usize = 0;
    while (std.mem.indexOfPos(u8, sheet_xml, pos, "<drawing")) |i| {
        const after_idx = i + "<drawing".len;
        if (after_idx >= sheet_xml.len) return null;
        if (!std.ascii.isWhitespace(sheet_xml[after_idx])) {
            pos = after_idx;
            continue;
        }
        const el_end = store_mod.xmlStartTagEnd(sheet_xml, i) orelse return null;
        return store_mod.xmlAttrValue(sheet_xml[i..el_end], "r:id");
    }
    return null;
}

/// Walk every `cNvPr` element in a drawing part and return the lowest
/// id above the highest one seen (>= 1). ECMA-376 requires `cNvPr`
/// ids to be unique within their drawing part, and the id space is
/// shared across shape kinds — pictures, charts' graphicFrames,
/// shapes — so the scan matches the element name under any namespace
/// prefix rather than picture anchors only. A crafted
/// `id="4294967295"` exhausts the space — refused typed rather than
/// trapping or duplicating.
pub fn nextFreeCnvprId(drawing_xml: []const u8) Error!u32 {
    var max_seen: u32 = 0;
    var pos: usize = 0;
    while (std.mem.indexOfPos(u8, drawing_xml, pos, "cNvPr")) |i| {
        pos = i + "cNvPr".len;
        // Element-name boundary on both sides: `<cNvPr` (no prefix) or
        // `:cNvPr` (prefixed) before, whitespace after (`id` is a
        // required attribute, so a bare `<cNvPr>` names nothing). This
        // is what keeps `cNvPicPr` / `cNvSpPr` from matching.
        if (i == 0) continue;
        const before = drawing_xml[i - 1];
        if (before != '<' and before != ':') continue;
        if (pos >= drawing_xml.len) break;
        if (!std.ascii.isWhitespace(drawing_xml[pos])) continue;
        // Back up to the element's `<` so the shared lexer sees a full
        // start tag, cut quote-aware — a `>` inside a name="a>b" value
        // must not truncate the scan.
        const lt = std.mem.lastIndexOfScalar(u8, drawing_xml[0..i], '<') orelse continue;
        const tag_end = store_mod.xmlStartTagEnd(drawing_xml, lt) orelse break;
        const id_str = store_mod.xmlAttrValue(drawing_xml[lt..tag_end], "id") orelse continue;
        const n = std.fmt.parseInt(u32, id_str, 10) catch continue;
        if (n > max_seen) max_seen = n;
    }
    if (max_seen == std.math.maxInt(u32)) return Error.IdSpaceExhausted;
    return max_seen + 1;
}

// ─── Tests ────────────────────────────────────────────────────────────

test "validateMagic: PNG accepts canonical 8-byte signature" {
    const png = [_]u8{ 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00 };
    try validateMagic(.png, &png);
}

test "validateMagic: PNG declared with JPEG bytes errors typed" {
    const jpeg = [_]u8{ 0xFF, 0xD8, 0xFF, 0xE0 };
    try std.testing.expectError(Error.MimeMagicMismatch, validateMagic(.png, &jpeg));
}

test "validateMagic: empty bytes errors EmptyImage" {
    try std.testing.expectError(Error.EmptyImage, validateMagic(.png, &.{}));
}

test "validateMagic: JPEG accepts FFD8FF prefix" {
    const jpeg = [_]u8{ 0xFF, 0xD8, 0xFF, 0xE0, 0x00, 0x10 };
    try validateMagic(.jpeg, &jpeg);
}

test "validateMagic: GIF accepts 'GIF' prefix" {
    const gif = "GIF89a\x00\x00";
    try validateMagic(.gif, gif);
}

/// Assert `xml` contains the exact marker element these coordinates
/// and offsets imply. Built by formatting rather than spelled as a
/// literal so that a zero offset has no hard-coded spelling anywhere
/// in this file — the whole point of the offsets work.
fn expectMarker(
    xml: []const u8,
    tag: []const u8,
    col0: u32,
    col_off: u32,
    row0: u32,
    row_off: u32,
) !void {
    var buf: [256]u8 = undefined;
    const want = try std.fmt.bufPrint(
        &buf,
        "<{s}><xdr:col>{d}</xdr:col><xdr:colOff>{d}</xdr:colOff>" ++
            "<xdr:row>{d}</xdr:row><xdr:rowOff>{d}</xdr:rowOff></{s}>",
        .{ tag, col0, col_off, row0, row_off, tag },
    );
    if (std.mem.indexOf(u8, xml, want) == null) {
        std.debug.print("missing marker:\n  want: {s}\n  in:   {s}\n", .{ want, xml });
        return error.MarkerNotFound;
    }
}

/// A 1-inch square, the extent every pre-native-size caller got.
const one_inch_square: ImageExtent = .{ .cx_emu = 914400, .cy_emu = 914400 };

test "buildDrawingXml: anchors to (col-1, row-1) and emits oneCellAnchor" {
    const xml = try buildDrawingXml(std.testing.allocator, .{ .one_cell = .{
        .at = .{ .col = 3, .row = 5 },
        .ext = one_inch_square,
    } }, .transitional);
    defer std.testing.allocator.free(xml);
    try expectMarker(xml, "xdr:from", 2, 0, 4, 0);
    try std.testing.expect(std.mem.indexOf(u8, xml, "<xdr:oneCellAnchor>") != null);
    try std.testing.expect(std.mem.indexOf(u8, xml, "</xdr:oneCellAnchor>") != null);
    try std.testing.expect(std.mem.indexOf(u8, xml, "r:embed=\"rId1\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, xml, "<xdr:ext cx=\"914400\" cy=\"914400\"/>") != null);
    // The one-cell form carries no `<xdr:to>`.
    try std.testing.expect(std.mem.indexOf(u8, xml, "<xdr:to>") == null);
}

test "buildDrawingXml: oneCellAnchor extent is whatever the caller passed" {
    // Not a round inch: proves the value is plumbed, not defaulted.
    const xml = try buildDrawingXml(std.testing.allocator, .{ .one_cell = .{
        .at = .{ .col = 1, .row = 1 },
        .ext = .{ .cx_emu = 1_524_000, .cy_emu = 762_000 },
    } }, .transitional);
    defer std.testing.allocator.free(xml);
    try std.testing.expect(std.mem.indexOf(u8, xml, "<xdr:ext cx=\"1524000\" cy=\"762000\"/>") != null);
    try std.testing.expect(std.mem.indexOf(u8, xml, "914400") == null);
}

test "buildDrawingXml: oneCellAnchor carries its anchor's pixel offsets" {
    const xml = try buildDrawingXml(std.testing.allocator, .{ .one_cell = .{
        .at = .{ .col = 3, .row = 5, .col_off_emu = 190_500, .row_off_emu = 95_250 },
        .ext = one_inch_square,
    } }, .transitional);
    defer std.testing.allocator.free(xml);
    // 20 px and 10 px at 96 DPI, on the grid corner (2, 4).
    try expectMarker(xml, "xdr:from", 2, 190_500, 4, 95_250);
}

test "buildDrawingXml: twoCellAnchor emits from (col-1) and to (col) markers" {
    // B2:D5 inclusive → from corner (1,1), to corner (4,5).
    const xml = try buildDrawingXml(std.testing.allocator, .{ .two_cell = .{
        .from = .{ .col = 2, .row = 2 },
        .to = .{ .col = 4, .row = 5 },
    } }, .transitional);
    defer std.testing.allocator.free(xml);

    try std.testing.expect(std.mem.indexOf(u8, xml, "<xdr:twoCellAnchor>") != null);
    try std.testing.expect(std.mem.indexOf(u8, xml, "</xdr:twoCellAnchor>") != null);
    try std.testing.expect(std.mem.indexOf(u8, xml, "<xdr:oneCellAnchor>") == null);

    try expectMarker(xml, "xdr:from", 1, 0, 1, 0);
    try expectMarker(xml, "xdr:to", 4, 0, 5, 0);

    // The range sizes the picture, so no fixed extent element.
    try std.testing.expect(std.mem.indexOf(u8, xml, "<xdr:ext") == null);
    // Wire order: from, to, pic — a `to` after `<xdr:pic>` is invalid.
    const to_pos = std.mem.indexOf(u8, xml, "<xdr:to>").?;
    const pic_pos = std.mem.indexOf(u8, xml, "<xdr:pic>").?;
    try std.testing.expect(to_pos < pic_pos);
    try std.testing.expect(std.mem.indexOf(u8, xml, "r:embed=\"rId1\"") != null);
}

test "buildDrawingXml: twoCellAnchor offsets are per-endpoint, not shared" {
    // Distinct values on all four fields: a marker that reused the
    // wrong endpoint's offsets would still match a symmetric case.
    const xml = try buildDrawingXml(std.testing.allocator, .{ .two_cell = .{
        .from = .{ .col = 2, .row = 2, .col_off_emu = 9_525, .row_off_emu = 19_050 },
        .to = .{ .col = 4, .row = 5, .col_off_emu = 28_575, .row_off_emu = 38_100 },
    } }, .transitional);
    defer std.testing.allocator.free(xml);
    try expectMarker(xml, "xdr:from", 1, 9_525, 1, 19_050);
    try expectMarker(xml, "xdr:to", 4, 28_575, 5, 38_100);
}

test "buildDrawingXml: 1x1 twoCellAnchor range spans exactly its own cell" {
    // A1:A1 → from corner (0,0), to corner (1,1): one cell of area.
    const xml = try buildDrawingXml(std.testing.allocator, .{ .two_cell = .{
        .from = .{ .col = 1, .row = 1 },
        .to = .{ .col = 1, .row = 1 },
    } }, .transitional);
    defer std.testing.allocator.free(xml);
    try expectMarker(xml, "xdr:to", 1, 0, 1, 0);
}

test "validateAnchor: rejects zero col/row on both anchor forms" {
    try std.testing.expectError(Error.InvalidAnchor, validateAnchor(.{ .one_cell = .{
        .at = .{ .col = 0, .row = 1 },
        .ext = one_inch_square,
    } }));
    try std.testing.expectError(Error.InvalidAnchor, validateAnchor(.{ .one_cell = .{
        .at = .{ .col = 1, .row = 0 },
        .ext = one_inch_square,
    } }));
    try std.testing.expectError(Error.InvalidAnchor, validateAnchor(.{ .two_cell = .{
        .from = .{ .col = 0, .row = 1 },
        .to = .{ .col = 2, .row = 2 },
    } }));
    try std.testing.expectError(Error.InvalidAnchor, validateAnchor(.{ .two_cell = .{
        .from = .{ .col = 1, .row = 1 },
        .to = .{ .col = 2, .row = 0 },
    } }));
}

test "validateAnchor: rejects inverted range, accepts degenerate and forward" {
    try std.testing.expectError(Error.InvertedAnchorRange, validateAnchor(.{ .two_cell = .{
        .from = .{ .col = 4, .row = 2 },
        .to = .{ .col = 2, .row = 5 },
    } }));
    try std.testing.expectError(Error.InvertedAnchorRange, validateAnchor(.{ .two_cell = .{
        .from = .{ .col = 2, .row = 9 },
        .to = .{ .col = 4, .row = 5 },
    } }));
    // from == to is a legal 1×1 range, not an inversion.
    try validateAnchor(.{ .two_cell = .{
        .from = .{ .col = 3, .row = 3 },
        .to = .{ .col = 3, .row = 3 },
    } });
    try validateAnchor(.{ .two_cell = .{
        .from = .{ .col = 1, .row = 1 },
        .to = .{ .col = 8, .row = 12 },
    } });
    try validateAnchor(.{ .one_cell = .{
        .at = .{ .col = 1, .row = 1 },
        .ext = one_inch_square,
    } });
}

test "validateAnchor: rejects a zero or over-large one-cell extent" {
    try std.testing.expectError(Error.InvalidExtent, validateAnchor(.{ .one_cell = .{
        .at = .{ .col = 1, .row = 1 },
        .ext = .{ .cx_emu = 0, .cy_emu = 914400 },
    } }));
    try std.testing.expectError(Error.InvalidExtent, validateAnchor(.{ .one_cell = .{
        .at = .{ .col = 1, .row = 1 },
        .ext = .{ .cx_emu = 914400, .cy_emu = 0 },
    } }));
    try std.testing.expectError(Error.InvalidExtent, validateAnchor(.{ .one_cell = .{
        .at = .{ .col = 1, .row = 1 },
        .ext = .{ .cx_emu = st_positive_coordinate_max + 1, .cy_emu = 914400 },
    } }));
    try std.testing.expectError(Error.InvalidExtent, validateAnchor(.{ .one_cell = .{
        .at = .{ .col = 1, .row = 1 },
        .ext = .{ .cx_emu = 914400, .cy_emu = st_positive_coordinate_max + 1 },
    } }));
    // Exactly at the cap is legal.
    try validateAnchor(.{ .one_cell = .{
        .at = .{ .col = 1, .row = 1 },
        .ext = .{ .cx_emu = st_positive_coordinate_max, .cy_emu = st_positive_coordinate_max },
    } });
}

test "validateAnchor: a zero-row anchor is rejected before its extent" {
    // Both fields are bad; the anchor error is the more specific
    // diagnosis, so it must win.
    try std.testing.expectError(Error.InvalidAnchor, validateAnchor(.{ .one_cell = .{
        .at = .{ .col = 0, .row = 0 },
        .ext = .{ .cx_emu = 0, .cy_emu = 0 },
    } }));
}

// ─── Native-size header reads ─────────────────────────────────────────

/// Canonical 1×1 PNG — the smallest legal IHDR.
const png_1x1 = [_]u8{
    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
    0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
    0x08, 0x00, 0x00, 0x00, 0x00, 0x3A, 0x7E, 0x9B,
    0x55,
};

/// A PNG header claiming 640×480, to prove the axes are read
/// big-endian and in the right order.
const png_640x480 = [_]u8{
    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
    0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x02, 0x80, 0x00, 0x00, 0x01, 0xE0,
    0x08, 0x06, 0x00, 0x00, 0x00,
};

test "pixelSize: PNG reads IHDR width and height big-endian" {
    try std.testing.expectEqual(
        ImagePixelSize{ .width = 1, .height = 1 },
        try pixelSize(.png, &png_1x1),
    );
    try std.testing.expectEqual(
        ImagePixelSize{ .width = 640, .height = 480 },
        try pixelSize(.png, &png_640x480),
    );
}

test "pixelSize: PNG truncated inside IHDR is undecodable" {
    // Signature and chunk type present, axes cut short.
    try std.testing.expectError(
        Error.UndecodableImageHeader,
        pixelSize(.png, png_640x480[0..20]),
    );
    // Signature only.
    try std.testing.expectError(
        Error.UndecodableImageHeader,
        pixelSize(.png, png_640x480[0..8]),
    );
}

test "pixelSize: PNG whose first chunk isn't IHDR is undecodable" {
    var bad = png_640x480;
    // "IHDR" → "iHDR": a valid-looking signature over a wrong chunk.
    bad[12] = 'i';
    try std.testing.expectError(Error.UndecodableImageHeader, pixelSize(.png, &bad));
}

test "pixelSize: PNG with a zero or over-cap axis is undecodable" {
    var zero_w = png_640x480;
    zero_w[16] = 0;
    zero_w[17] = 0;
    zero_w[18] = 0;
    zero_w[19] = 0;
    try std.testing.expectError(Error.UndecodableImageHeader, pixelSize(.png, &zero_w));

    var huge_h = png_640x480;
    // 0x80000000 = 2^31, one past the spec's cap.
    huge_h[20] = 0x80;
    huge_h[21] = 0x00;
    huge_h[22] = 0x00;
    huge_h[23] = 0x00;
    try std.testing.expectError(Error.UndecodableImageHeader, pixelSize(.png, &huge_h));
}

test "pixelSize: GIF reads the logical screen descriptor little-endian" {
    // "GIF89a" then 640×480 as two LE u16s.
    const gif = [_]u8{ 'G', 'I', 'F', '8', '9', 'a', 0x80, 0x02, 0xE0, 0x01, 0x00 };
    try std.testing.expectEqual(
        ImagePixelSize{ .width = 640, .height = 480 },
        try pixelSize(.gif, &gif),
    );
}

test "pixelSize: GIF truncated before its descriptor is undecodable" {
    const gif = [_]u8{ 'G', 'I', 'F', '8', '9', 'a', 0x80, 0x02 };
    try std.testing.expectError(Error.UndecodableImageHeader, pixelSize(.gif, &gif));
}

test "pixelSize: JPEG walks segments to SOF0 and reads height before width" {
    const jpeg = [_]u8{
        0xFF, 0xD8, // SOI
        0xFF, 0xE0, 0x00, 0x04, 0x00, 0x00, // APP0, length 4 → 2 payload bytes
        0xFF, 0xDB, 0x00, 0x03, 0x00, // DQT, length 3 → 1 payload byte
        0xFF, 0xC0, 0x00, 0x0B, // SOF0, length 11
        0x08, // precision
        0x01, 0xE0, // height 480
        0x02, 0x80, // width 640
        0x01, 0x01,
        0x11, 0x00,
    };
    try std.testing.expectEqual(
        ImagePixelSize{ .width = 640, .height = 480 },
        try pixelSize(.jpeg, &jpeg),
    );
}

test "pixelSize: JPEG skips DHT, whose marker sits inside the SOF range" {
    // 0xC4 is DHT, not a frame header: reading its payload as a size
    // would yield 0x0102 × 0x0304 instead of walking on to SOF2.
    const jpeg = [_]u8{
        0xFF, 0xD8,
        0xFF, 0xC4, 0x00, 0x07, 0x00, 0x01, 0x02, 0x03, 0x04, // DHT
        0xFF, 0xC2, 0x00, 0x0B, // SOF2 (progressive)
        0x08,
        0x00, 0x64, // height 100
        0x00, 0xC8, // width 200
        0x01, 0x01,
        0x11, 0x00,
    };
    try std.testing.expectEqual(
        ImagePixelSize{ .width = 200, .height = 100 },
        try pixelSize(.jpeg, &jpeg),
    );
}

test "pixelSize: JPEG tolerates 0xFF fill bytes before a marker" {
    const jpeg = [_]u8{
        0xFF, 0xD8,
        0xFF, 0xFF, 0xFF, 0xC0, 0x00, 0x0B, // two fill bytes, then SOF0
        0x08,
        0x00, 0x10, // height 16
        0x00, 0x20, // width 32
        0x01, 0x01,
        0x11, 0x00,
    };
    try std.testing.expectEqual(
        ImagePixelSize{ .width = 32, .height = 16 },
        try pixelSize(.jpeg, &jpeg),
    );
}

test "pixelSize: JPEG reaching scan data without a frame header is undecodable" {
    const jpeg = [_]u8{
        0xFF, 0xD8,
        0xFF, 0xE0,
        0x00, 0x04,
        0x00, 0x00,
        0xFF, 0xDA, 0x00, 0x08, 0x01, 0x01, 0x00, 0x00, 0x3F, 0x00, // SOS
    };
    try std.testing.expectError(Error.UndecodableImageHeader, pixelSize(.jpeg, &jpeg));
}

test "pixelSize: JPEG with a segment length past the buffer is undecodable" {
    // Claims a 200-byte segment inside a 10-byte file.
    const jpeg = [_]u8{ 0xFF, 0xD8, 0xFF, 0xC0, 0x00, 0xC8, 0x08, 0x00, 0x10, 0x00 };
    try std.testing.expectError(Error.UndecodableImageHeader, pixelSize(.jpeg, &jpeg));
}

test "pixelSize: a mime/magic mismatch outranks an undecodable header" {
    // Truncated JPEG bytes declared as PNG: the caller's first mistake
    // is the declaration, so report that, not the failed parse.
    const jpeg = [_]u8{ 0xFF, 0xD8, 0xFF };
    try std.testing.expectError(Error.MimeMagicMismatch, pixelSize(.png, &jpeg));
    try std.testing.expectError(Error.EmptyImage, pixelSize(.png, &.{}));
}

test "nativeExtent: maps pixels to EMU at 96 DPI" {
    // 96 px = 1 inch = 914 400 EMU.
    const gif = [_]u8{ 'G', 'I', 'F', '8', '9', 'a', 0x60, 0x00, 0x30, 0x00, 0x00 };
    const ext = try nativeExtent(.gif, &gif);
    try std.testing.expectEqual(@as(u64, 914_400), ext.cx_emu);
    try std.testing.expectEqual(@as(u64, 457_200), ext.cy_emu);

    // A 1×1 PNG is one pixel, not one inch — the old fixed extent
    // would have been 96× too large on each axis.
    const tiny = try nativeExtent(.png, &png_1x1);
    try std.testing.expectEqual(@as(u64, 9_525), tiny.cx_emu);
    try std.testing.expectEqual(@as(u64, 9_525), tiny.cy_emu);
}

test "ImageCellAnchor.atPixelOffset: converts at 96 DPI without a cast" {
    const a = ImageCellAnchor.atPixelOffset(3, 5, 4, 7);
    try std.testing.expectEqual(@as(u32, 3), a.col);
    try std.testing.expectEqual(@as(u32, 5), a.row);
    try std.testing.expectEqual(@as(u32, 38_100), a.col_off_emu);
    try std.testing.expectEqual(@as(u32, 66_675), a.row_off_emu);

    // The widest pixel offset the helper accepts still fits `u32`,
    // which is what makes the internal `@intCast` total.
    const widest = ImageCellAnchor.atPixelOffset(1, 1, std.math.maxInt(u16), std.math.maxInt(u16));
    try std.testing.expectEqual(@as(u32, 624_220_875), widest.col_off_emu);
    try std.testing.expect(widest.col_off_emu < std.math.maxInt(u32));

    // Default is a flush-to-the-cell-corner anchor.
    const plain: ImageCellAnchor = .{ .col = 1, .row = 1 };
    try std.testing.expectEqual(@as(u32, 0), plain.col_off_emu);
    try std.testing.expectEqual(@as(u32, 0), plain.row_off_emu);
}

test "emuFromPixels: agrees with the inch constant and stays in range" {
    try std.testing.expectEqual(@as(u64, 0), emuFromPixels(0));
    try std.testing.expectEqual(@as(u64, 914_400), emuFromPixels(96));
    // The bound the doc comment claims for decoded headers.
    try std.testing.expect(emuFromPixels(pixel_axis_max) <= st_positive_coordinate_max);
    // And the bound it warns about for raw u32 input.
    try std.testing.expect(emuFromPixels(std.math.maxInt(u32)) > st_positive_coordinate_max);
}

test "buildDrawingRels: includes ../media/<basename> target and the given Type" {
    const rels = try buildDrawingRels(std.testing.allocator, "image1.png", image_rel_type_transitional);
    defer std.testing.allocator.free(rels);
    try std.testing.expect(std.mem.indexOf(u8, rels, "Target=\"../media/image1.png\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, rels, "Id=\"rId1\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, rels, image_rel_type_transitional) != null);

    // A Strict host gets the Strict Type; the OPC wrapper namespace
    // stays the Part 2 URI either way.
    const strict = try buildDrawingRels(std.testing.allocator, "image1.png", image_rel_type_strict);
    defer std.testing.allocator.free(strict);
    try std.testing.expect(std.mem.indexOf(u8, strict, image_rel_type_strict) != null);
    try std.testing.expect(std.mem.indexOf(
        u8,
        strict,
        "http://schemas.openxmlformats.org/package/2006/relationships",
    ) != null);
}

test "appendDrawingElementToSheet: splices before </worksheet>" {
    const sheet =
        "<?xml version=\"1.0\"?>" ++
        "<worksheet xmlns:r=\"...\"><sheetData/></worksheet>";
    const out = try appendDrawingElementToSheet(std.testing.allocator, sheet, "rId7");
    defer std.testing.allocator.free(out);
    try std.testing.expect(std.mem.indexOf(u8, out, "<drawing r:id=\"rId7\"/></worksheet>") != null);
}

test "appendRelationship: appends before </Relationships>" {
    const existing =
        "<?xml version=\"1.0\"?>" ++
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" ++
        "<Relationship Id=\"rId1\" Type=\"x\" Target=\"y\"/>" ++
        "</Relationships>";
    const out = try appendRelationship(
        std.testing.allocator,
        existing,
        "drawing1.xml",
        "rId2",
        drawing_rel_type_transitional,
    );
    defer std.testing.allocator.free(out);
    try std.testing.expect(std.mem.indexOf(u8, out, "Id=\"rId2\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "Target=\"../drawings/drawing1.xml\"") != null);
    // Original rel still present.
    try std.testing.expect(std.mem.indexOf(u8, out, "Id=\"rId1\"") != null);
}

// ─── Append-to-existing-drawing helpers ───────────────────────────────

test "buildPicAnchorXml: plumbs the rel id and cNvPr id/name" {
    const xml = try buildPicAnchorXml(std.testing.allocator, .{ .one_cell = .{
        .at = .{ .col = 2, .row = 3 },
        .ext = one_inch_square,
    } }, "rId7", 3, .inherit);
    defer std.testing.allocator.free(xml);
    try std.testing.expect(std.mem.indexOf(u8, xml, "r:embed=\"rId7\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, xml, "<xdr:cNvPr id=\"3\" name=\"Picture 3\"/>") != null);
    // A fragment, not a document: no wsDr wrapper of its own, and
    // `.inherit` declares nothing.
    try std.testing.expect(std.mem.indexOf(u8, xml, "wsDr") == null);
    try std.testing.expect(std.mem.indexOf(u8, xml, "xmlns") == null);
    try std.testing.expect(std.mem.startsWith(u8, xml, "<xdr:oneCellAnchor>"));
    try std.testing.expect(std.mem.endsWith(u8, xml, "</xdr:oneCellAnchor>"));
}

test "buildPicAnchorXml: .declare pins all three prefixes in the host's dialect" {
    const xml = try buildPicAnchorXml(std.testing.allocator, .{ .one_cell = .{
        .at = .{ .col = 1, .row = 1 },
        .ext = one_inch_square,
    } }, "rId2", 2, .{ .declare = .strict });
    defer std.testing.allocator.free(xml);
    try std.testing.expect(std.mem.indexOf(u8, xml, "xmlns:xdr=\"" ++ ns_xdr_strict ++ "\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, xml, "xmlns:a=\"" ++ ns_a_strict ++ "\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, xml, "xmlns:r=\"" ++ ns_r_strict ++ "\"") != null);
    // The declarations live on the anchor's open tag, before its `>`.
    const decl_pos = std.mem.indexOf(u8, xml, "xmlns:xdr").?;
    const first_close = std.mem.indexOfScalar(u8, xml, '>').?;
    try std.testing.expect(decl_pos < first_close);

    const trans = try buildPicAnchorXml(std.testing.allocator, .{ .one_cell = .{
        .at = .{ .col = 1, .row = 1 },
        .ext = one_inch_square,
    } }, "rId2", 2, .{ .declare = .transitional });
    defer std.testing.allocator.free(trans);
    try std.testing.expect(std.mem.indexOf(u8, trans, "xmlns:xdr=\"" ++ ns_xdr_transitional ++ "\"") != null);
}

test "detectWsDrDialect: reads the root xmlns:xdr binding" {
    const fresh = try buildDrawingXml(std.testing.allocator, .{ .one_cell = .{
        .at = .{ .col = 1, .row = 1 },
        .ext = one_inch_square,
    } }, .transitional);
    defer std.testing.allocator.free(fresh);
    try std.testing.expectEqual(@as(?WsDrDialect, .transitional), detectWsDrDialect(fresh));

    const strict =
        "<xdr:wsDr xmlns:xdr='" ++ ns_xdr_strict ++ "'></xdr:wsDr>";
    try std.testing.expectEqual(@as(?WsDrDialect, .strict), detectWsDrDialect(strict));

    // Unknown URI, foreign root, or a root that never declares the
    // prefix: no dialect — the caller refuses to splice.
    const bogus =
        "<xdr:wsDr xmlns:xdr=\"http://example.com/not-a-drawing\"></xdr:wsDr>";
    try std.testing.expect(detectWsDrDialect(bogus) == null);
    try std.testing.expect(detectWsDrDialect("<c:chartSpace></c:chartSpace>") == null);
    try std.testing.expect(detectWsDrDialect("<xdr:wsDr></xdr:wsDr>") == null);
}

test "detectWsDrDialect: a commented pseudo-root cannot spoof the dialect" {
    // The comment binds Strict; the REAL root is Transitional. Reading
    // the comment's URI would splice Strict content into a
    // Transitional part (Codex r2, MAJOR 1).
    const spoofed =
        "<?xml version=\"1.0\"?>" ++
        "<!-- <xdr:wsDr xmlns:xdr=\"" ++ ns_xdr_strict ++ "\"> -->" ++
        "<xdr:wsDr xmlns:xdr=\"" ++ ns_xdr_transitional ++ "\"></xdr:wsDr>";
    try std.testing.expectEqual(@as(?WsDrDialect, .transitional), detectWsDrDialect(spoofed));

    // A comment with no real root after it detects nothing.
    const only_comment =
        "<!-- <xdr:wsDr xmlns:xdr=\"" ++ ns_xdr_strict ++ "\"> -->";
    try std.testing.expect(detectWsDrDialect(only_comment) == null);
}

test "detectSheetDialect: reads the worksheet root's default namespace" {
    try std.testing.expectEqual(WsDrDialect.transitional, detectSheetDialect(
        "<?xml version=\"1.0\"?><worksheet xmlns=\"" ++ ns_main_transitional ++ "\"><sheetData/></worksheet>",
    ));
    try std.testing.expectEqual(WsDrDialect.strict, detectSheetDialect(
        "<worksheet xmlns=\"" ++ ns_main_strict ++ "\"><sheetData/></worksheet>",
    ));
    // Unrecognized roots fall back to Transitional — the dialect every
    // mainstream producer writes.
    try std.testing.expectEqual(WsDrDialect.transitional, detectSheetDialect("<chartsheet/>"));
    try std.testing.expectEqual(WsDrDialect.transitional, detectSheetDialect(""));
}

test "nextFreeCnvprId: a '>' inside an attribute value cannot truncate the scan" {
    // `>` is legal inside attribute values (XML 1.0 §2.4). Cutting the
    // tag at the first `>` would hide id 7 and hand out a duplicate
    // (Codex r2, MAJOR 2).
    try std.testing.expectEqual(@as(u32, 8), try nextFreeCnvprId(
        "<xdr:wsDr><xdr:cNvPr name=\"a>b\" id=\"7\"/></xdr:wsDr>",
    ));
}

test "appendAnchorToDrawing: splices before </xdr:wsDr>, keeping the first anchor" {
    const fresh = try buildDrawingXml(std.testing.allocator, .{ .one_cell = .{
        .at = .{ .col = 1, .row = 1 },
        .ext = one_inch_square,
    } }, .transitional);
    defer std.testing.allocator.free(fresh);

    const second = try buildPicAnchorXml(std.testing.allocator, .{ .two_cell = .{
        .from = .{ .col = 2, .row = 2 },
        .to = .{ .col = 4, .row = 5 },
    } }, "rId2", 2, .{ .declare = .transitional });
    defer std.testing.allocator.free(second);

    const out = try appendAnchorToDrawing(std.testing.allocator, fresh, second);
    defer std.testing.allocator.free(out);

    try std.testing.expect(std.mem.indexOf(u8, out, "<xdr:oneCellAnchor>") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "</xdr:twoCellAnchor>") != null);
    // The splice lands after the existing anchor, inside the wrapper.
    const one_pos = std.mem.indexOf(u8, out, "</xdr:oneCellAnchor>").?;
    const two_pos = std.mem.indexOf(u8, out, "<xdr:twoCellAnchor").?;
    const close_pos = std.mem.indexOf(u8, out, "</xdr:wsDr>").?;
    try std.testing.expect(one_pos < two_pos);
    try std.testing.expect(two_pos < close_pos);
    try std.testing.expect(std.mem.endsWith(u8, out, "</xdr:wsDr>"));
}

test "appendAnchorToDrawing: refuses a part without a wsDr wrapper" {
    const chart =
        "<?xml version=\"1.0\"?>" ++
        "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">" ++
        "</c:chartSpace>";
    try std.testing.expectError(
        error.UnsupportedDrawingPart,
        appendAnchorToDrawing(std.testing.allocator, chart, "<xdr:oneCellAnchor/>"),
    );
}

test "appendImageRelationship: appends an image rel before </Relationships>" {
    const existing =
        "<?xml version=\"1.0\"?>" ++
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" ++
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/image1.png\"/>" ++
        "</Relationships>";
    const out = try appendImageRelationship(
        std.testing.allocator,
        existing,
        "image2.gif",
        "rId2",
        image_rel_type_transitional,
    );
    defer std.testing.allocator.free(out);
    try std.testing.expect(std.mem.indexOf(u8, out, "Id=\"rId2\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "Target=\"../media/image2.gif\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, out, "Target=\"../media/image1.png\"") != null);
    try std.testing.expect(std.mem.endsWith(u8, out, "</Relationships>"));

    // The Type is whatever dialect the caller resolved.
    const strict_out = try appendImageRelationship(
        std.testing.allocator,
        existing,
        "image2.gif",
        "rId2",
        image_rel_type_strict,
    );
    defer std.testing.allocator.free(strict_out);
    try std.testing.expect(std.mem.indexOf(u8, strict_out, image_rel_type_strict) != null);

    try std.testing.expectError(
        error.MalformedDrawingRels,
        appendImageRelationship(
            std.testing.allocator,
            "<Relationships>",
            "image2.gif",
            "rId2",
            image_rel_type_transitional,
        ),
    );
}

test "extractDrawingRelId: pulls the r:id and skips non-drawing lookalikes" {
    const sheet =
        "<worksheet><sheetData/>" ++
        "<drawing r:id=\"rId3\"/>" ++
        "</worksheet>";
    try std.testing.expectEqualStrings("rId3", extractDrawingRelId(sheet).?);

    // No drawing element at all.
    try std.testing.expect(extractDrawingRelId("<worksheet><sheetData/></worksheet>") == null);
    // Header/footer images use <drawingHF>, a different element.
    try std.testing.expect(extractDrawingRelId("<worksheet><drawingHF r:id=\"rId2\"/></worksheet>") == null);
    // A bare <drawing> without attributes carries no r:id to chase.
    try std.testing.expect(extractDrawingRelId("<worksheet><drawing/></worksheet>") == null);
}

test "extractDrawingRelId: accepts single quotes and Eq whitespace" {
    // Both are legal XML 1.0 spellings some producers emit; missing
    // them turns a resolvable drawing into DrawingRelationshipNotFound
    // (Codex r1, MAJOR 2).
    try std.testing.expectEqualStrings(
        "rId9",
        extractDrawingRelId("<worksheet><drawing r:id='rId9'/></worksheet>").?,
    );
    try std.testing.expectEqualStrings(
        "rId4",
        extractDrawingRelId("<worksheet><drawing\n  r:id = \"rId4\" /></worksheet>").?,
    );
}

test "nextFreeCnvprId: shares the id space across shape kinds" {
    // A picture anchor (id 1) plus a chart graphicFrame (id 7): the
    // next picture must clear BOTH, and the cNvPicPr decoy must not
    // parse as a cNvPr.
    const drawing =
        "<xdr:wsDr>" ++
        "<xdr:oneCellAnchor><xdr:pic><xdr:nvPicPr>" ++
        "<xdr:cNvPr id=\"1\" name=\"Picture 1\"/><xdr:cNvPicPr/>" ++
        "</xdr:nvPicPr></xdr:pic></xdr:oneCellAnchor>" ++
        "<xdr:twoCellAnchor><xdr:graphicFrame><xdr:nvGraphicFramePr>" ++
        "<xdr:cNvPr id=\"7\" name=\"Chart 1\"/>" ++
        "</xdr:nvGraphicFramePr></xdr:graphicFrame></xdr:twoCellAnchor>" ++
        "</xdr:wsDr>";
    try std.testing.expectEqual(@as(u32, 8), try nextFreeCnvprId(drawing));

    // An empty wsDr starts at 1.
    try std.testing.expectEqual(@as(u32, 1), try nextFreeCnvprId("<xdr:wsDr></xdr:wsDr>"));
    // Attribute order: name before id still parses.
    try std.testing.expectEqual(@as(u32, 3), try nextFreeCnvprId(
        "<xdr:wsDr><xdr:cNvPr name=\"Picture 2\" id=\"2\"/></xdr:wsDr>",
    ));
    // Single quotes and Eq whitespace are legal spellings; missing
    // them would hand out a duplicate id (Codex r1, MAJOR 2).
    try std.testing.expectEqual(@as(u32, 10), try nextFreeCnvprId(
        "<xdr:wsDr><xdr:cNvPr id='9' name='P'/></xdr:wsDr>",
    ));
    try std.testing.expectEqual(@as(u32, 6), try nextFreeCnvprId(
        "<xdr:wsDr><xdr:cNvPr\n id = \"5\" name=\"P\"/></xdr:wsDr>",
    ));
    // A crafted maxInt id is typed exhaustion, not a trap or a
    // duplicate.
    try std.testing.expectError(Error.IdSpaceExhausted, nextFreeCnvprId(
        "<xdr:wsDr><xdr:cNvPr id=\"4294967295\" name=\"P\"/></xdr:wsDr>",
    ));
}

test "nextFreeRelId: a crafted maxInt rId is typed exhaustion" {
    const rels = [_]store_mod.Relationship{.{
        .id = "rId4294967295",
        .type = "x",
        .target = "y",
        .target_mode = .internal,
    }};
    try std.testing.expectError(Error.IdSpaceExhausted, nextFreeRelId(&rels));
}

test "validateAnchor: rejects anchors past the sheet grid caps" {
    // XFD / 1_048_576 are the last legal cells; one past either axis
    // is outside the grid (and would eventually leave the marker
    // types' xsd:int value space far behind at u32 extremes).
    try validateAnchor(.{ .one_cell = .{
        .at = .{ .col = 16_384, .row = 1_048_576 },
        .ext = one_inch_square,
    } });
    try std.testing.expectError(Error.InvalidAnchor, validateAnchor(.{ .one_cell = .{
        .at = .{ .col = 16_385, .row = 1 },
        .ext = one_inch_square,
    } }));
    try std.testing.expectError(Error.InvalidAnchor, validateAnchor(.{ .one_cell = .{
        .at = .{ .col = 1, .row = 1_048_577 },
        .ext = one_inch_square,
    } }));
    try std.testing.expectError(Error.InvalidAnchor, validateAnchor(.{ .two_cell = .{
        .from = .{ .col = 1, .row = 1 },
        .to = .{ .col = std.math.maxInt(u32), .row = 5 },
    } }));
}
