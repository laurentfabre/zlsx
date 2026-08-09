//! The XML decode boundary — how bytes in a package become values the
//! evaluator may read (`goal_formula.md` §5.7.1, M4b1).
//!
//! M4b1 of the tier-D1 ladder, and the half of it that does not touch
//! `pkg/`. Everything here is a pure function over part bytes, which is
//! what lets the boundary be fixtured, fuzzed, and proven before the
//! adapter that calls it exists.
//!
//! Decoding is split by CARRIER CLASS, and getting it wrong is silent
//! ---------------------------------------------------------------
//! Two decode passes exist — XML entities (`&amp;`, `&#10;`) and
//! ST_Xstring (`_xHHHH_`) — and which of them applies is a property of
//! *where the bytes came from*, not of what they look like:
//!
//!   * **FORMULA carriers** — `<f>` bodies, defined-name bodies, table
//!     calculated-column and totals-row formulas — get XML entity
//!     decoding **ONLY**. ST_Xstring does not apply, so a literal
//!     `_x0041_` written inside a formula survives byte-exact and is
//!     still `_x0041_` when the tokenizer sees it. Applying the second
//!     pass here would rewrite `IF(A1,"_x0041_","b")` into
//!     `IF(A1,"A","b")` — a changed formula that still parses, still
//!     evaluates, and is wrong.
//!   * **STRING carriers** — shared strings, inline strings, `t="str"`
//!     values, and the ST_Xstring-typed identifiers (sheet, defined-name,
//!     table, and table-column names) — get XML entities **then**
//!     ST_Xstring. Rich strings concatenate every visible `<r>/<t>` run
//!     in document order; `<rPh>` phonetic runs are excluded, because
//!     furigana is an annotation Excel shows beside the text and never
//!     part of the cell's value.
//!   * **LEXICAL carriers** — numeric, boolean and error `<v>` bodies —
//!     get XML entities only, like formulas, but are listed separately
//!     because the *authored* direction differs: this engine generates
//!     their spelling, so there is nothing to escape and nothing that can
//!     fail to encode.
//!
//! Authoring is the same split, inverted: a string ST_Xstring-encodes
//! and then XML-escapes; a formula XML-escapes and **nothing else**,
//! which is why a formula carrying a character XML cannot represent is a
//! refusal rather than an escape (there is no second pass to hide it in).
//!
//! Provenance
//! ----------
//! The carrier split, the rich-run rule, and the implicit-coordinate
//! rule are pinned by `goal_formula.md` §5 M4b1 and are already frozen in
//! the oracle extractor (`tests/oracle/extractor.zig:14-24,392-397`),
//! which records manifests this boundary must stay comparable with. The
//! decode of an encoded C0 control in a shared string is corpus-decided:
//! `tests/corpus/wdi_excel.xlsx` carries `_x000D_` in `xl/sharedStrings.xml`.
//! Everything else is spec-pinned and cited at its table: element and
//! attribute inventories to ECMA-376 (`CT_Worksheet`, `CT_Row`, `CT_Cell`,
//! `CT_Rst`, `CT_Table`), the cell-type mapping to `ST_CellType` plus
//! §5.7.2's input cell-type contract, the implicit coordinate to
//! MS-OE376 §2.1.624, and the namespace to the SpreadsheetML main
//! namespace URI.
//!
//! Allocation
//! ----------
//! The scanning entry points own an arena and hand it back inside their
//! result; a refusal frees it before returning, so a caller that ignores
//! the payload of a refusal leaks nothing. The primitives take an
//! allocator and return owned bytes. Nothing here holds a reference to
//! the input after it returns except through explicitly borrowed slices,
//! which are documented per field.

const std = @import("std");
const assert = std.debug.assert;

const coords = @import("zlsx_refs");
const value = @import("value.zig");
const parser = @import("parser.zig");
const tokenizer = @import("tokenizer.zig");
const serial_date = @import("serial_date.zig");

/// §10's plane-2 taxonomy has exactly one home (`metadata.zig:75-79`
/// pays the same import for the same reason).
pub const PlaneTwo = parser.PlaneTwo;

// ─── refusals (§10) ──────────────────────────────────────────────

pub const LimitKind = enum {
    part_bytes,
    modeled_cells,
    shared_strings,
    depth,
    skip_depth,
    namespace_bindings,
    rich_runs,
};

/// Where in the grid a refusal happened, when it happened at a cell.
/// Kept numeric rather than as an A1 string so constructing it cannot
/// itself allocate or fail.
pub const CellSite = struct {
    /// One-based, as OOXML writes it.
    row: u32,
    /// Zero-based, as `coords.Col` counts.
    col: u32,
};

pub const Refusal = struct {
    reason: Reason,
    /// Byte offset into the part where the refusal was detected.
    offset: u32 = 0,
    /// Set exactly when `reason == .limit_exceeded`.
    limit: ?LimitKind = null,
    /// Set when the refusal is about one cell.
    cell: ?CellSite = null,

    pub const Reason = enum {
        // ── the part as a document ──
        /// Not valid UTF-8.
        invalid_utf8,
        /// Tag soup: unterminated markup, a close that matches no open,
        /// text where only elements are allowed.
        malformed_xml,
        /// `<!DOCTYPE` / `<!ENTITY`. Refused outright: an
        /// entity-expanding sheet reader is a denial-of-service surface
        /// with no upside (`tests/corpus/poi_xxe_in_schema.xlsx`).
        doctype_declaration,
        /// `<![CDATA[…]]>`. Legal XML that SpreadsheetML never uses;
        /// its content is NOT entity-decoded, so silently treating it
        /// as text would decode one workbook two ways.
        unexpected_cdata,
        /// An element with no row in the inventory legal at this
        /// position.
        unexpected_element,
        /// An element from a namespace other than the main one, inside
        /// `<sheetData>`. Skipping it could drop cells; interpreting it
        /// would mean guessing at a vocabulary we do not know.
        foreign_element,
        /// An attribute the schema does not define on an interpreted
        /// element. Foreign-namespace attributes are exempt — that is
        /// what a namespace is for (`x14ac:dyDescent` is the common one).
        unexpected_attribute,
        /// A required attribute is missing.
        missing_attribute,
        /// An attribute whose value is not the lexical form its type
        /// demands.
        bad_attribute_value,
        /// Character data where the schema allows only elements.
        unexpected_text,
        /// §9-shaped bound; `limit` names it.
        limit_exceeded,

        // ── namespace preflight ──
        /// The root element is not the one this part kind must have.
        wrong_root_element,
        /// The root element is in no namespace at all.
        missing_namespace,
        /// A namespace this engine knows about and does not implement
        /// (the ISO Strict vocabulary).
        unsupported_namespace,
        /// A namespace nothing has classified. Refused before any
        /// mutation: a scanner that literal-matches `<sheetData>` reads
        /// a document in an unknown vocabulary as an empty sheet, and an
        /// empty sheet recalculates to a workbook full of zeroes.
        unknown_namespace,

        // ── text decoding ──
        /// An undeclared entity, or a character reference outside the
        /// XML character range.
        bad_entity,
        /// An `_xHHHH_` escape naming a surrogate or a non-scalar.
        bad_xstring_escape,
        /// An authored formula carries a character XML cannot represent
        /// and ST_Xstring is not available to escape it (see the header).
        unencodable_formula_char,

        // ── the input cell-type contract (§5.7.2) ──
        /// A `t` outside `ST_CellType`'s enumeration. Never a silent
        /// number.
        unknown_cell_type,
        /// A `t="d"` `<v>` outside §5.7.2's normative lexical table —
        /// not one of the two accepted ISO-8601 forms, carrying a
        /// timezone offset, or naming a date the active epoch cannot
        /// express (M4b2, `serial_date.serialFromLexical`).
        bad_date_cache,
        /// A `<v>` that is not a number under §5.4's invariant grammar,
        /// or is one that binary64 cannot hold.
        bad_number_cache,
        /// A `t="b"` `<v>` that is not exactly `0` or `1`.
        bad_boolean_cache,
        /// A `t="e"` `<v>` that is not an error spelling.
        bad_error_cache,
        /// A non-formula cell whose `t` demands a `<v>` that is absent.
        missing_cached_value,
        /// `t="s"` naming an entry the shared-string table does not have.
        shared_string_index_out_of_range,
        /// A `<c r="…">` that is not an A1 reference.
        malformed_cell_ref,
        /// An implicit coordinate would step off the edge of the grid —
        /// past column XFD, or past row 1 048 576.
        implicit_ref_out_of_grid,
        /// Two `<c>` elements at the same coordinate. Last-wins and
        /// first-wins are both defensible, which is why neither is
        /// silently chosen.
        duplicate_cell,

        // ── the decoded symbol layer (§5.9) ──
        /// Two sheets, names, or tables that fold to one spelling. A
        /// reference to it could not name one of them.
        duplicate_symbol,
        /// A referenced `_xlnm.` name nothing has classified.
        unknown_builtin_name,
        /// A referenced `_xlpm.` LAMBDA parameter. §5.9 refuses
        /// LAMBDA/LET through v1, and a parameter binding reached from
        /// outside its own body has no value to give.
        lambda_parameter_name,
        /// A referenced name carrying `function`, `vbProcedure`, or
        /// `xlm` (M4b3's `CT_DefinedName` inventory). The name is a
        /// macro entry point, not a value; resolving it as one would
        /// answer a different question than the formula asked. Carried
        /// rather than refused at read time — a macro name a formula
        /// never mentions is not a reason to refuse a workbook.
        macro_defined_name,
        /// A referenced name whose body contains a relative reference.
        /// What it denotes depends on where it is used, and v1 does not
        /// carry the use site into name expansion (§5.9; M10+ lifts it).
        relative_reference_name,
    };

    /// Exhaustive by construction — a new `Reason` fails to compile
    /// until it has a §10 plane.
    pub fn planeTwo(self: Refusal) PlaneTwo {
        return switch (self.reason) {
            .invalid_utf8,
            .malformed_xml,
            .unexpected_cdata,
            .unexpected_attribute,
            .missing_attribute,
            .bad_attribute_value,
            .unexpected_text,
            .wrong_root_element,
            .missing_namespace,
            .bad_entity,
            .bad_xstring_escape,
            .unknown_cell_type,
            .bad_number_cache,
            .bad_boolean_cache,
            .bad_error_cache,
            .bad_date_cache,
            .missing_cached_value,
            .shared_string_index_out_of_range,
            .malformed_cell_ref,
            .implicit_ref_out_of_grid,
            .duplicate_cell,
            .duplicate_symbol,
            => .FormulaMalformedInput,

            .doctype_declaration,
            .unexpected_element,
            .foreign_element,
            .unsupported_namespace,
            .unknown_namespace,
            .unencodable_formula_char,
            .unknown_builtin_name,
            .lambda_parameter_name,
            .macro_defined_name,
            .relative_reference_name,
            => .FormulaUnsupportedConstruct,

            .limit_exceeded => .FormulaLimitExceeded,
        };
    }
};

pub const Limits = struct {
    /// Every diagnostic offset is a `u32`, so a part that could not be
    /// addressed by one refuses before it is scanned rather than
    /// reporting a truncated lie about where the problem is.
    max_bytes: u32 = std.math.maxInt(u32),
    /// §9's "workbook materialization" row, by its name and its
    /// default. Charged per scan; `WorkbookEnv.build` passes the
    /// remaining workbook-wide budget to each sheet, so the bound is on
    /// the model as a whole rather than on whichever sheet is largest.
    max_modeled_cells: u32 = 64 << 20,
    max_shared_strings: u32 = 1 << 22,
    /// `worksheet > sheetData > row > c > is > r > t` is 7.
    max_depth: u32 = 32,
    /// Foreign content outside `<sheetData>` nests as deep as it likes;
    /// this is where "as deep as it likes" stops.
    max_skip_depth: u32 = 64,
    max_namespace_bindings: u32 = 32,
    max_rich_runs: u32 = 1 << 16,
};

pub const Options = struct {
    limits: Limits = .{},
    fidelity: value.Fidelity = .excel,
    /// Which epoch a `t="d"` `<v>` is read against (§5.7.2). Workbook-
    /// derived and never caller-writable at the public layer — the
    /// adapter reads it from `<workbookPr date1904>` — because the same
    /// eight bytes are a different serial under each system.
    date_system: serial_date.DateSystem = .d1900,
};

// ─── carrier classes ─────────────────────────────────────────────

/// Which decode passes a span of bytes gets. The whole point of the
/// file: this is a property of the SITE, never of the bytes.
pub const Carrier = enum {
    /// XML entities only. ST_Xstring does NOT apply.
    formula,
    /// XML entities, then ST_Xstring.
    string,
    /// XML entities only, on a token this engine also generates.
    lexical,

    pub fn appliesXstring(self: Carrier) bool {
        return self == .string;
    }
};

/// Every place the engine takes bytes out of a package. A site with no
/// row here has no decode rule, which is the same as saying it cannot be
/// read — `carrierOf` is a switch, so adding a site without classifying
/// it fails to compile.
pub const Site = enum {
    /// `<f>` body.
    cell_formula_body,
    /// `<definedName>` body.
    defined_name_body,
    /// `<calculatedColumnFormula>` body.
    table_calculated_column_formula,
    /// `<totalsRowFormula>` body.
    table_totals_row_formula,
    /// `<si>` — shared string, rich runs concatenated.
    shared_string,
    /// `<is>` — inline string, rich runs concatenated.
    inline_string,
    /// `t="str"` `<v>` — a formula's cached text result.
    formula_string_value,
    /// `CT_Sheet@name`.
    sheet_name,
    /// `CT_DefinedName@name`.
    defined_name_identifier,
    /// `CT_Table@name` / `@displayName`.
    table_name,
    /// `CT_TableColumn@name`.
    table_column_name,
    /// A numeric `<v>`.
    number_value,
    /// A `t="b"` `<v>`.
    boolean_value,
    /// A `t="e"` `<v>`.
    error_value,

    pub fn carrier(self: Site) Carrier {
        return switch (self) {
            .cell_formula_body,
            .defined_name_body,
            .table_calculated_column_formula,
            .table_totals_row_formula,
            => .formula,

            // ST_Xstring-typed in the schema, and the only class where
            // that typing is also what Office does.
            .shared_string,
            .inline_string,
            .formula_string_value,
            .sheet_name,
            .defined_name_identifier,
            .table_name,
            .table_column_name,
            => .string,

            .number_value,
            .boolean_value,
            .error_value,
            => .lexical,
        };
    }
};

/// The SpreadsheetML main namespace. Element matching is by (namespace,
/// local name): a document that binds it to a prefix — `<x:worksheet
/// xmlns:x="…">` — is as valid as one that makes it the default, and a
/// reader that literal-matches `<sheetData>` sees the first as empty.
pub const ns_main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
/// ISO 29500 Strict. Recognized so it refuses as *classified* rather
/// than as unknown: the vocabulary is nearly the same but the parts it
/// differs in (dates, some enumerations) are exactly the parts recalc
/// reads.
pub const ns_main_strict = "http://purl.oclc.org/ooxml/spreadsheetml/main";

pub const NamespaceTreatment = enum { accepted, unsupported };

pub const NamespaceRow = struct {
    uri: []const u8,
    treatment: NamespaceTreatment,
    note: []const u8,
};

/// Not documentation: `preflight` reads it, so a namespace absent from
/// the table is `unknown_namespace` by construction.
pub const namespace_inventory = [_]NamespaceRow{
    .{
        .uri = ns_main,
        .treatment = .accepted,
        .note = "ECMA-376 transitional SpreadsheetML; what Excel writes",
    },
    .{
        .uri = ns_main_strict,
        .treatment = .unsupported,
        .note = "ISO 29500 Strict; recognized so it refuses classified",
    },
};

/// Which part a preflight is looking at, and therefore which root
/// element it must find.
pub const PartKind = enum {
    worksheet,
    workbook,
    shared_strings,
    table,

    pub fn rootElement(self: PartKind) []const u8 {
        return switch (self) {
            .worksheet => "worksheet",
            .workbook => "workbook",
            .shared_strings => "sst",
            .table => "table",
        };
    }
};

// ─── text primitives ─────────────────────────────────────────────

pub const TextError = error{
    OutOfMemory,
    BadEntity,
    BadXstring,
};

/// Decode the five predefined XML entities and numeric character
/// references. This is the whole of a formula carrier's decoding.
///
/// An undeclared entity refuses. Passing it through would turn `&nbsp;`
/// into six literal bytes that no producer meant, and in a formula
/// carrier those bytes go straight to the tokenizer.
pub fn decodeEntities(allocator: std.mem.Allocator, raw: []const u8) TextError![]u8 {
    if (std.mem.indexOfScalar(u8, raw, '&') == null) return allocator.dupe(u8, raw);

    var out: std.ArrayListUnmanaged(u8) = .empty;
    errdefer out.deinit(allocator);
    try out.ensureTotalCapacity(allocator, raw.len);

    var i: usize = 0;
    while (i < raw.len) {
        if (raw[i] != '&') {
            try out.append(allocator, raw[i]);
            i += 1;
            continue;
        }
        const semi = std.mem.indexOfScalarPos(u8, raw, i, ';') orelse return error.BadEntity;
        const body = raw[i + 1 .. semi];
        if (body.len == 0) return error.BadEntity;

        if (std.mem.eql(u8, body, "amp")) {
            try out.append(allocator, '&');
        } else if (std.mem.eql(u8, body, "lt")) {
            try out.append(allocator, '<');
        } else if (std.mem.eql(u8, body, "gt")) {
            try out.append(allocator, '>');
        } else if (std.mem.eql(u8, body, "quot")) {
            try out.append(allocator, '"');
        } else if (std.mem.eql(u8, body, "apos")) {
            try out.append(allocator, '\'');
        } else if (body[0] == '#') {
            const cp = parseCharRef(body[1..]) orelse return error.BadEntity;
            var buf: [4]u8 = undefined;
            const n = std.unicode.utf8Encode(cp, &buf) catch return error.BadEntity;
            try out.appendSlice(allocator, buf[0..n]);
        } else {
            return error.BadEntity;
        }
        i = semi + 1;
    }
    return out.toOwnedSlice(allocator);
}

fn parseCharRef(digits: []const u8) ?u21 {
    if (digits.len == 0) return null;
    const v = if (digits[0] == 'x' or digits[0] == 'X')
        std.fmt.parseInt(u32, digits[1..], 16) catch return null
    else
        std.fmt.parseInt(u32, digits, 10) catch return null;
    if (v > 0x10FFFF) return null;
    if (v >= 0xD800 and v <= 0xDFFF) return null;
    return @intCast(v);
}

/// Decode ST_Xstring's `_xHHHH_` escapes — the SECOND pass, and only
/// for string carriers.
///
/// `_x005F_` is the escape for a literal underscore, which is what makes
/// the encoding invertible: `_x005F_x0041_` is the literal text
/// `_x0041_`, not the letter `A`. A sequence that only looks like an
/// escape (`_xZZZZ_`, `_x12_`) is literal text and passes through, which
/// is also what Excel does — the escape grammar is exact or it is
/// nothing.
pub fn decodeXstring(allocator: std.mem.Allocator, s: []const u8) TextError![]u8 {
    if (std.mem.indexOfScalar(u8, s, '_') == null) return allocator.dupe(u8, s);

    var out: std.ArrayListUnmanaged(u8) = .empty;
    errdefer out.deinit(allocator);
    try out.ensureTotalCapacity(allocator, s.len);

    var i: usize = 0;
    while (i < s.len) {
        if (xstringEscapeAt(s, i)) |cp| {
            var buf: [4]u8 = undefined;
            const n = std.unicode.utf8Encode(cp, &buf) catch return error.BadXstring;
            try out.appendSlice(allocator, buf[0..n]);
            i += escape_len;
            continue;
        }
        try out.append(allocator, s[i]);
        i += 1;
    }
    return out.toOwnedSlice(allocator);
}

/// `_xHHHH_` — underscore, `x`, four hex digits, underscore.
const escape_len: usize = 7;

/// The code point an escape at `i` denotes, or null when the bytes there
/// are not an escape. A surrogate is not a scalar value and cannot be
/// what the producer meant, so it is reported as an error by the caller
/// through `utf8Encode` failing.
fn xstringEscapeAt(s: []const u8, i: usize) ?u21 {
    if (i + escape_len > s.len) return null;
    if (s[i] != '_') return null;
    if (s[i + 1] != 'x' and s[i + 1] != 'X') return null;
    if (s[i + 6] != '_') return null;
    var v: u32 = 0;
    for (s[i + 2 .. i + 6]) |c| {
        const d = hexDigit(c) orelse return null;
        v = v * 16 + d;
    }
    return @intCast(v);
}

fn hexDigit(c: u8) ?u4 {
    return switch (c) {
        '0'...'9' => @intCast(c - '0'),
        'a'...'f' => @intCast(c - 'a' + 10),
        'A'...'F' => @intCast(c - 'A' + 10),
        else => null,
    };
}

/// Decode a span according to its site's carrier class. The one function
/// every reader should call, because it cannot be called without naming
/// the site.
pub fn decodeAt(allocator: std.mem.Allocator, site: Site, raw: []const u8) TextError![]u8 {
    return decodeCarrier(allocator, site.carrier(), raw);
}

pub fn decodeCarrier(allocator: std.mem.Allocator, carrier: Carrier, raw: []const u8) TextError![]u8 {
    const entities = try decodeEntities(allocator, raw);
    if (!carrier.appliesXstring()) return entities;
    defer allocator.free(entities);
    return decodeXstring(allocator, entities);
}

/// XML-escape `&`, `<`, `>` — the escaping every carrier's authored form
/// ends with.
///
/// `"` and `'` are escaped too, so one function serves element text and
/// attribute values; both spellings are legal in element content and
/// Excel writes `&quot;` inside attributes.
pub fn xmlEscape(allocator: std.mem.Allocator, s: []const u8) error{OutOfMemory}![]u8 {
    var out: std.ArrayListUnmanaged(u8) = .empty;
    errdefer out.deinit(allocator);
    try out.ensureTotalCapacity(allocator, s.len);
    for (s) |c| switch (c) {
        '&' => try out.appendSlice(allocator, "&amp;"),
        '<' => try out.appendSlice(allocator, "&lt;"),
        '>' => try out.appendSlice(allocator, "&gt;"),
        '"' => try out.appendSlice(allocator, "&quot;"),
        '\'' => try out.appendSlice(allocator, "&apos;"),
        else => try out.append(allocator, c),
    };
    return out.toOwnedSlice(allocator);
}

/// ST_Xstring-encode: escape what XML cannot carry, and escape a literal
/// `_` that would otherwise start an escape.
///
/// The second half is the whole reason the encoding is invertible. Text
/// containing the six characters `_x0041_` must come back as those six
/// characters, so it is written `_x005F_x0041_`; without that rule
/// `decodeXstring` would hand back `A` and a user's string would have
/// silently changed.
pub fn encodeXstring(allocator: std.mem.Allocator, s: []const u8) error{OutOfMemory}![]u8 {
    var out: std.ArrayListUnmanaged(u8) = .empty;
    errdefer out.deinit(allocator);
    try out.ensureTotalCapacity(allocator, s.len);

    var i: usize = 0;
    while (i < s.len) {
        const c = s[i];
        if (c == '_' and xstringEscapeAt(s, i) != null) {
            try out.appendSlice(allocator, "_x005F_");
            i += 1;
            continue;
        }
        if (needsXstringEscape(c)) {
            var buf: [7]u8 = undefined;
            const w = std.fmt.bufPrint(&buf, "_x{X:0>4}_", .{c}) catch unreachable;
            try out.appendSlice(allocator, w);
            i += 1;
            continue;
        }
        try out.append(allocator, c);
        i += 1;
    }
    return out.toOwnedSlice(allocator);
}

/// The C0 controls XML 1.0 cannot represent, plus carriage return.
///
/// Tab and line feed stay literal — XML carries them and Excel writes
/// them literally. CR is escaped: `tests/corpus/wdi_excel.xlsx` has
/// `_x000D_` in its shared strings, which is the corpus deciding this row
/// rather than the schema, and it has to be escaped anyway because XML
/// parsers normalize a literal CR to LF.
fn needsXstringEscape(c: u8) bool {
    if (c == '\t' or c == '\n') return false;
    return c < 0x20;
}

/// Author a STRING carrier: ST_Xstring first, then XML escaping.
pub fn encodeAuthoredString(allocator: std.mem.Allocator, text: []const u8) error{OutOfMemory}![]u8 {
    const xstr = try encodeXstring(allocator, text);
    defer allocator.free(xstr);
    return xmlEscape(allocator, xstr);
}

pub const AuthorError = error{ OutOfMemory, UnencodableChar };

/// Author a FORMULA carrier: XML escaping, and **no ST_Xstring stage**.
///
/// The asymmetry is load-bearing in both directions. Reading, it means
/// `_x0041_` in a formula stays six characters; writing, it means a
/// formula containing a character XML cannot represent has nowhere to
/// go — there is no second escape layer to put it in — so this refuses
/// instead of inventing one. Emitting `_x0001_` here would produce a
/// formula that reads back as those seven literal characters.
pub fn encodeAuthoredFormula(allocator: std.mem.Allocator, text: []const u8) AuthorError![]u8 {
    for (text) |c| {
        if (c == '\t' or c == '\n' or c == '\r') continue;
        if (c < 0x20) return error.UnencodableChar;
    }
    return xmlEscape(allocator, text);
}

// ─── the scanner ─────────────────────────────────────────────────

/// A namespace-aware pull scanner over one part.
///
/// Independent of `pkg/typed_parts/*`: those views drop the raw `t`
/// attribute, skip a `<c>` with no `r`, and match tags literally, and
/// each of those is a thing this row exists to get right.
pub const Event = union(enum) {
    open: Element,
    self_closing: Element,
    close: Name,
    /// Character data, still entity-encoded.
    text: []const u8,
    cdata: []const u8,
    doctype,
};

pub const Name = struct {
    /// The qualified name exactly as written.
    qname: []const u8,

    pub fn local(self: Name) []const u8 {
        if (std.mem.indexOfScalar(u8, self.qname, ':')) |c| return self.qname[c + 1 ..];
        return self.qname;
    }

    pub fn prefix(self: Name) []const u8 {
        if (std.mem.indexOfScalar(u8, self.qname, ':')) |c| return self.qname[0..c];
        return "";
    }
};

pub const Element = struct {
    name: Name,
    /// The raw attribute region between the name and the closing `>`.
    attrs: []const u8,
    /// Offset of the `<` that opened this element.
    offset: usize,

    pub fn local(self: Element) []const u8 {
        return self.name.local();
    }

    pub fn attr(self: Element, name: []const u8) ?[]const u8 {
        var it = self.attrIterator();
        while (it.next()) |a| {
            if (a.prefix().len == 0 and std.mem.eql(u8, a.local(), name)) return a.raw_value;
        }
        return null;
    }

    pub fn attrIterator(self: Element) AttrIterator {
        return .{ .attrs = self.attrs };
    }
};

pub const Attr = struct {
    qname: []const u8,
    raw_value: []const u8,

    pub fn local(self: Attr) []const u8 {
        if (std.mem.indexOfScalar(u8, self.qname, ':')) |c| return self.qname[c + 1 ..];
        return self.qname;
    }

    pub fn prefix(self: Attr) []const u8 {
        if (std.mem.indexOfScalar(u8, self.qname, ':')) |c| return self.qname[0..c];
        return "";
    }
};

pub const AttrIterator = struct {
    attrs: []const u8,
    i: usize = 0,

    pub fn next(self: *AttrIterator) ?Attr {
        const s = self.attrs;
        while (self.i < s.len and isSpace(s[self.i])) : (self.i += 1) {}
        if (self.i >= s.len) return null;

        const name_start = self.i;
        while (self.i < s.len and s[self.i] != '=' and !isSpace(s[self.i])) : (self.i += 1) {}
        const qname = s[name_start..self.i];
        while (self.i < s.len and isSpace(s[self.i])) : (self.i += 1) {}
        if (self.i >= s.len or s[self.i] != '=') return null;
        self.i += 1;
        while (self.i < s.len and isSpace(s[self.i])) : (self.i += 1) {}
        if (self.i >= s.len) return null;
        const quote = s[self.i];
        if (quote != '"' and quote != '\'') return null;
        self.i += 1;
        const value_start = self.i;
        while (self.i < s.len and s[self.i] != quote) : (self.i += 1) {}
        const raw = s[value_start..self.i];
        if (self.i < s.len) self.i += 1;
        return .{ .qname = qname, .raw_value = raw };
    }
};

pub const ScanError = error{Malformed};

pub const Scanner = struct {
    xml: []const u8,
    i: usize = 0,

    pub fn init(xml: []const u8) Scanner {
        // A UTF-8 BOM ahead of the declaration is legal and common.
        const start: usize = if (std.mem.startsWith(u8, xml, "\xEF\xBB\xBF")) 3 else 0;
        return .{ .xml = xml, .i = start };
    }

    pub fn next(self: *Scanner) ScanError!?Event {
        while (true) {
            if (self.i >= self.xml.len) return null;

            if (self.xml[self.i] != '<') {
                const start = self.i;
                while (self.i < self.xml.len and self.xml[self.i] != '<') : (self.i += 1) {}
                return .{ .text = self.xml[start..self.i] };
            }

            if (std.mem.startsWith(u8, self.xml[self.i..], "<!--")) {
                const end = std.mem.indexOfPos(u8, self.xml, self.i + 4, "-->") orelse
                    return error.Malformed;
                self.i = end + 3;
                continue;
            }
            if (std.mem.startsWith(u8, self.xml[self.i..], "<![CDATA[")) {
                const body = self.i + 9;
                const end = std.mem.indexOfPos(u8, self.xml, body, "]]>") orelse
                    return error.Malformed;
                const out = self.xml[body..end];
                self.i = end + 3;
                return .{ .cdata = out };
            }
            if (std.mem.startsWith(u8, self.xml[self.i..], "<!")) {
                const end = std.mem.indexOfScalarPos(u8, self.xml, self.i, '>') orelse
                    return error.Malformed;
                self.i = end + 1;
                return .doctype;
            }
            if (std.mem.startsWith(u8, self.xml[self.i..], "<?")) {
                const end = std.mem.indexOfPos(u8, self.xml, self.i + 2, "?>") orelse
                    return error.Malformed;
                self.i = end + 2;
                continue;
            }

            const open_at = self.i;
            const tag_end = indexOfTagEnd(self.xml, self.i) orelse return error.Malformed;
            const inner = self.xml[self.i + 1 .. tag_end];
            self.i = tag_end + 1;
            if (inner.len == 0) return error.Malformed;

            if (inner[0] == '/') {
                const name = trimSpace(inner[1..]);
                if (name.len == 0) return error.Malformed;
                return .{ .close = .{ .qname = name } };
            }

            const self_closing = inner[inner.len - 1] == '/';
            const body = if (self_closing) inner[0 .. inner.len - 1] else inner;
            var n: usize = 0;
            while (n < body.len and !isSpace(body[n])) : (n += 1) {}
            if (n == 0) return error.Malformed;
            const element: Element = .{
                .name = .{ .qname = body[0..n] },
                .attrs = body[n..],
                .offset = open_at,
            };
            return if (self_closing) .{ .self_closing = element } else .{ .open = element };
        }
    }
};

/// Find the `>` closing the tag at `start`, skipping any inside a quoted
/// attribute value. `<c r="A>1"/>` is pathological but legal, and a
/// naive `indexOfScalar` truncates it.
fn indexOfTagEnd(xml: []const u8, start: usize) ?usize {
    var i = start + 1;
    var quote: ?u8 = null;
    while (i < xml.len) : (i += 1) {
        const c = xml[i];
        if (quote) |q| {
            if (c == q) quote = null;
            continue;
        }
        switch (c) {
            '"', '\'' => quote = c,
            '>' => return i,
            else => {},
        }
    }
    return null;
}

inline fn isSpace(c: u8) bool {
    return c == ' ' or c == '\t' or c == '\n' or c == '\r';
}

fn trimSpace(s: []const u8) []const u8 {
    var a: usize = 0;
    var b: usize = s.len;
    while (a < b and isSpace(s[a])) : (a += 1) {}
    while (b > a and isSpace(s[b - 1])) : (b -= 1) {}
    return s[a..b];
}

fn isWhitespaceOnly(s: []const u8) bool {
    for (s) |c| {
        if (!isSpace(c)) return false;
    }
    return true;
}

// ─── namespace bindings ──────────────────────────────────────────

/// The prefix→URI bindings in scope, as a depth-tagged stack. Scoped
/// rather than root-only because a document may rebind a prefix on an
/// inner element, and a reader that assumed otherwise would classify the
/// rebound subtree as main-namespace when it is not.
const NsStack = struct {
    const Binding = struct { prefix: []const u8, uri: []const u8, depth: u32 };

    buf: []Binding,
    len: usize = 0,

    fn push(self: *NsStack, b: Binding) bool {
        if (self.len >= self.buf.len) return false;
        self.buf[self.len] = b;
        self.len += 1;
        return true;
    }

    fn popTo(self: *NsStack, depth: u32) void {
        while (self.len > 0 and self.buf[self.len - 1].depth > depth) self.len -= 1;
    }

    fn resolve(self: *const NsStack, prefix: []const u8) ?[]const u8 {
        var i = self.len;
        while (i > 0) {
            i -= 1;
            if (std.mem.eql(u8, self.buf[i].prefix, prefix)) return self.buf[i].uri;
        }
        // The `xml` prefix is bound by the XML spec itself, never
        // declared. Nothing here interprets it; naming it keeps
        // `xml:space` from looking like an unbound prefix.
        if (std.mem.eql(u8, prefix, "xml")) return "http://www.w3.org/XML/1998/namespace";
        return null;
    }

    /// Record every `xmlns` / `xmlns:p` declaration on one element.
    fn declare(self: *NsStack, el: Element, depth: u32) bool {
        var it = el.attrIterator();
        while (it.next()) |a| {
            if (std.mem.eql(u8, a.qname, "xmlns")) {
                if (!self.push(.{ .prefix = "", .uri = a.raw_value, .depth = depth })) return false;
            } else if (std.mem.eql(u8, a.prefix(), "xmlns")) {
                if (!self.push(.{ .prefix = a.local(), .uri = a.raw_value, .depth = depth })) return false;
            }
        }
        return true;
    }
};

/// True when an element is the named main-namespace element.
fn isMain(ns: *const NsStack, el: Element, local_name: []const u8) bool {
    if (!std.mem.eql(u8, el.local(), local_name)) return false;
    const uri = ns.resolve(el.name.prefix()) orelse return false;
    return std.mem.eql(u8, uri, ns_main);
}

fn inMainNs(ns: *const NsStack, name: Name) bool {
    const uri = ns.resolve(name.prefix()) orelse return false;
    return std.mem.eql(u8, uri, ns_main);
}

// ─── namespace preflight ─────────────────────────────────────────

/// Refuse an unreadable vocabulary **before** anything is decoded, let
/// alone mutated.
///
/// The failure this prevents is not a crash. A scanner that matches
/// `<sheetData>` literally reads `<x:sheetData>` as absent, reports a
/// sheet with no cells, and a recalc over "no cells" is a workbook of
/// zeroes written over someone's data. The preflight is the cheapest
/// possible place to notice, and it runs over every part the run will
/// read.
pub fn preflight(kind: PartKind, xml: []const u8, limits: Limits) ?Refusal {
    if (xml.len > limits.max_bytes) {
        return .{ .reason = .limit_exceeded, .limit = .part_bytes };
    }
    if (!std.unicode.utf8ValidateSlice(xml)) return .{ .reason = .invalid_utf8 };

    var ns_buf: [64]NsStack.Binding = undefined;
    const cap = @min(ns_buf.len, limits.max_namespace_bindings);
    var ns: NsStack = .{ .buf = ns_buf[0..cap] };

    var sc: Scanner = .init(xml);
    while (true) {
        const ev = sc.next() catch |err| switch (err) {
            error.Malformed => return .{ .reason = .malformed_xml, .offset = offsetOf(sc.i) },
        };
        const e = ev orelse return .{ .reason = .wrong_root_element };
        switch (e) {
            .doctype => return .{ .reason = .doctype_declaration, .offset = offsetOf(sc.i) },
            .cdata => return .{ .reason = .unexpected_cdata, .offset = offsetOf(sc.i) },
            .text => |t| {
                if (!isWhitespaceOnly(t)) {
                    return .{ .reason = .unexpected_text, .offset = offsetOf(sc.i) };
                }
            },
            .close => return .{ .reason = .malformed_xml, .offset = offsetOf(sc.i) },
            .open, .self_closing => |el| {
                if (!ns.declare(el, 0)) {
                    return .{ .reason = .limit_exceeded, .limit = .namespace_bindings };
                }
                if (!std.mem.eql(u8, el.local(), kind.rootElement())) {
                    return .{ .reason = .wrong_root_element, .offset = offsetOf(el.offset) };
                }
                const uri = ns.resolve(el.name.prefix()) orelse
                    return .{ .reason = .missing_namespace, .offset = offsetOf(el.offset) };
                for (namespace_inventory) |row| {
                    if (!std.mem.eql(u8, row.uri, uri)) continue;
                    return switch (row.treatment) {
                        .accepted => null,
                        .unsupported => .{
                            .reason = .unsupported_namespace,
                            .offset = offsetOf(el.offset),
                        },
                    };
                }
                return .{ .reason = .unknown_namespace, .offset = offsetOf(el.offset) };
            },
        }
    }
}

fn offsetOf(i: usize) u32 {
    return @intCast(@min(i, std.math.maxInt(u32)));
}

/// The span a slice the scanner handed out occupies in the part it came
/// from. Every such slice is a subslice of `xml` — the scanner never
/// copies — so the arithmetic is exact rather than a search.
///
/// Public since M5b2: `calc.zig` records `<calcPr>`'s coordinates the
/// same way the sheet walk records a `<c>`'s, and a second implementation
/// of this arithmetic is a second chance to be off by one.
pub fn spanOfSub(xml: []const u8, sub: []const u8) Span {
    assert(@intFromPtr(sub.ptr) >= @intFromPtr(xml.ptr));
    const start = @intFromPtr(sub.ptr) - @intFromPtr(xml.ptr);
    assert(start + sub.len <= xml.len);
    return .{ .start = offsetOf(start), .end = offsetOf(start + sub.len) };
}

/// The span of one unprefixed attribute inside a start tag's raw
/// attribute region — the first byte of its name through the closing
/// quote — or null when the element has no such attribute.
///
/// Public since M5b2 for the same reason as `spanOfSub`: the calc-state
/// patcher replaces a whole `calcId="…"` run, and the run it replaces has
/// to be the one the parser read.
pub fn attrSpanIn(xml: []const u8, attrs: []const u8, name: []const u8) ?Span {
    const region = spanOfSub(xml, attrs);
    var it: AttrIterator = .{ .attrs = attrs };
    while (it.next()) |a| {
        if (a.prefix().len != 0) continue;
        if (!std.mem.eql(u8, a.local(), name)) continue;
        const name_at = spanOfSub(xml, a.qname).start;
        // `raw_value` stops at the closing quote; the span includes it,
        // clamped because an unterminated value would otherwise address
        // a byte past the tag. `indexOfTagEnd` makes that unreachable in
        // a document the scanner accepted, and clamping is cheaper than
        // depending on it.
        const value_end = spanOfSub(xml, a.raw_value).end;
        return .{ .start = name_at, .end = @min(value_end + 1, region.end) };
    }
    return null;
}

/// The offset of the `<` that opens the close tag ending at `end`.
fn closeTagStart(xml: []const u8, end: usize) u32 {
    return offsetOf(std.mem.lastIndexOfScalar(u8, xml[0..end], '<') orelse end);
}

// ─── shared strings ──────────────────────────────────────────────

/// The decoded shared-string table: one string per `<si>`, in part
/// order, each already through both string-carrier passes.
pub const Strings = struct {
    arena: std.heap.ArenaAllocator,
    items: [][]const u8,

    pub fn deinit(self: *Strings) void {
        self.arena.deinit();
        self.* = undefined;
    }

    pub fn at(self: Strings, idx: usize) ?[]const u8 {
        if (idx >= self.items.len) return null;
        return self.items[idx];
    }
};

pub const StringsResult = union(enum) {
    ok: Strings,
    refused: Refusal,
};

/// An empty table, for a workbook with no `xl/sharedStrings.xml`.
pub fn emptyStrings(allocator: std.mem.Allocator) Strings {
    return .{ .arena = std.heap.ArenaAllocator.init(allocator), .items = &.{} };
}

/// Decode `xl/sharedStrings.xml`.
///
/// Each `<si>` is every visible `<t>` concatenated in document order.
/// `<rPh>` runs are excluded: phonetic guides are shown beside the text,
/// never part of it, and concatenating them would make `LEN` disagree
/// with Excel on every furigana-annotated cell.
pub fn decodeSharedStrings(
    gpa: std.mem.Allocator,
    xml: []const u8,
    opts: Options,
) error{OutOfMemory}!StringsResult {
    if (preflight(.shared_strings, xml, opts.limits)) |r| return .{ .refused = r };

    var arena = std.heap.ArenaAllocator.init(gpa);
    // A refusal is a normal return, not an error, so `errdefer` would
    // not fire — and a caller who is told "refused" must not also have
    // to free a payload it was never handed.
    var keep = false;
    defer if (!keep) arena.deinit();
    const a = arena.allocator();

    var items: std.ArrayListUnmanaged([]const u8) = .empty;
    var run: std.ArrayListUnmanaged(u8) = .empty;

    var ns_buf: [64]NsStack.Binding = undefined;
    const cap = @min(ns_buf.len, opts.limits.max_namespace_bindings);
    var ns: NsStack = .{ .buf = ns_buf[0..cap] };

    var depth: u32 = 0;
    var si_depth: ?u32 = null;
    var phonetic_depth: ?u32 = null;
    var in_t = false;
    var runs: u32 = 0;

    var sc: Scanner = .init(xml);
    while (true) {
        const ev = sc.next() catch return .{
            .refused = .{ .reason = .malformed_xml, .offset = offsetOf(sc.i) },
        };
        const e = ev orelse break;
        switch (e) {
            .doctype => return .{ .refused = .{ .reason = .doctype_declaration, .offset = offsetOf(sc.i) } },
            .cdata => return .{ .refused = .{ .reason = .unexpected_cdata, .offset = offsetOf(sc.i) } },
            .text => |t| {
                if (in_t and phonetic_depth == null) {
                    const decoded = decodeEntities(a, t) catch |err| return .{
                        .refused = textRefusal(err, sc.i) catch return error.OutOfMemory,
                    };
                    try run.appendSlice(a, decoded);
                }
            },
            .open, .self_closing => |el| {
                const closing = e == .self_closing;
                depth += 1;
                if (depth > opts.limits.max_depth) {
                    return .{ .refused = .{ .reason = .limit_exceeded, .limit = .depth } };
                }
                if (!ns.declare(el, depth)) {
                    return .{ .refused = .{ .reason = .limit_exceeded, .limit = .namespace_bindings } };
                }
                if (isMain(&ns, el, "si") and si_depth == null) {
                    if (items.items.len >= opts.limits.max_shared_strings) {
                        return .{ .refused = .{ .reason = .limit_exceeded, .limit = .shared_strings } };
                    }
                    si_depth = depth;
                    // Per string, not per part: the bound is on how many
                    // runs one `<si>` may concatenate. Counting across
                    // the whole table refused `tests/corpus/wdi_excel.xlsx`,
                    // which is simply a large workbook.
                    runs = 0;
                    run.clearRetainingCapacity();
                    if (closing) {
                        // `<si/>` is a legal empty entry.
                        try items.append(a, "");
                        si_depth = null;
                    }
                } else if (si_depth != null and isMain(&ns, el, "rPh")) {
                    if (phonetic_depth == null) phonetic_depth = depth;
                } else if (si_depth != null and isMain(&ns, el, "t")) {
                    runs += 1;
                    if (runs > opts.limits.max_rich_runs) {
                        return .{ .refused = .{ .reason = .limit_exceeded, .limit = .rich_runs } };
                    }
                    if (!closing) in_t = true;
                }
                if (closing) {
                    if (phonetic_depth == depth) phonetic_depth = null;
                    depth -= 1;
                }
            },
            .close => {
                if (in_t) in_t = false;
                if (phonetic_depth) |pd| {
                    if (pd == depth) phonetic_depth = null;
                }
                if (si_depth) |sd| {
                    if (sd == depth) {
                        const decoded = decodeXstring(a, run.items) catch |err| return .{
                            .refused = try textRefusal(err, sc.i),
                        };
                        try items.append(a, decoded);
                        si_depth = null;
                    }
                }
                if (depth == 0) {
                    return .{ .refused = .{ .reason = .malformed_xml, .offset = offsetOf(sc.i) } };
                }
                ns.popTo(depth - 1);
                depth -= 1;
            },
        }
    }

    const owned = try items.toOwnedSlice(a);
    keep = true;
    return .{ .ok = .{ .arena = arena, .items = owned } };
}

/// Turn a text-primitive error into the refusal that names it. The
/// `OutOfMemory` arm cannot be a refusal — running out of memory is not
/// a statement about the workbook — so it is re-raised.
fn textRefusal(err: TextError, offset: usize) error{OutOfMemory}!Refusal {
    return switch (err) {
        error.OutOfMemory => error.OutOfMemory,
        error.BadEntity => .{ .reason = .bad_entity, .offset = offsetOf(offset) },
        error.BadXstring => .{ .reason = .bad_xstring_escape, .offset = offsetOf(offset) },
    };
}

// ─── implicit coordinates (MS-OE376 §2.1.624) ────────────────────

/// Reconstructs the column of a `<c>` that carries no `r`.
///
/// Office's rule: such a cell sits in the column **after** its
/// predecessor in the same row, and the first cell of a row with no `r`
/// is column A. `pkg/typed_parts/sheet_xml.zig:507` skips these cells
/// entirely, which is invisible until a workbook written by a producer
/// that omits `r` recalculates as if those cells were empty.
///
/// Stepping past the last column is a refusal, not a wrap: there is no
/// column after XFD, and a clamped coordinate would silently overwrite
/// one that already exists.
pub const CellCursor = struct {
    /// The column of the previous `<c>` in this row, or null at the
    /// start of a row.
    previous: ?coords.Col = null,
    /// The row of the previous `<row>`, or null at the start of the
    /// sheet. The row axis has the same rule and the same producers:
    /// `tests/corpus/wdi_excel.xlsx` omits `r` on **both** `<row>` and
    /// `<c>` for all 287 MB of its sheet, and `src/xlsx.zig:1858-1864`
    /// already reconstructs it the same way.
    previous_row: ?u32 = null,

    pub fn startRow(self: *CellCursor) void {
        self.previous = null;
    }

    /// The row a `<row>` without `r` sits in: the one after its
    /// predecessor, and 1 for the first row of the sheet.
    pub fn implicitRow(self: *const CellCursor) Resolution.Row {
        const next = if (self.previous_row) |p| p + 1 else 1;
        if (next > coords.max_row) return .{ .refused = .implicit_ref_out_of_grid };
        return .{ .ok = next };
    }

    pub const Resolution = union(enum) {
        ok: coords.Col,
        refused: Refusal.Reason,

        pub const Row = union(enum) {
            ok: u32,
            refused: Refusal.Reason,
        };
    };

    /// Resolve a cell's column from its `r` attribute, or from the
    /// cursor when it has none. An explicit `r` also *sets* the cursor,
    /// so `<c r="C1"/><c/>` puts the second cell in D.
    pub fn resolve(self: *CellCursor, r_attr: ?[]const u8, row: u32) Resolution {
        if (r_attr) |ref| {
            const cell = coords.parseCell(ref, .{ .case = .insensitive }) catch
                return .{ .refused = .malformed_cell_ref };
            if (cell.row.oneBased() != row) return .{ .refused = .malformed_cell_ref };
            self.previous = cell.col;
            return .{ .ok = cell.col };
        }
        const next_zero: u32 = if (self.previous) |p| blk: {
            if (p.zeroBased() == coords.max_col_1based - 1) {
                return .{ .refused = .implicit_ref_out_of_grid };
            }
            break :blk p.zeroBased() + 1;
        } else 0;
        const col = coords.Col.fromZeroBased(next_zero) catch
            return .{ .refused = .implicit_ref_out_of_grid };
        self.previous = col;
        return .{ .ok = col };
    }
};

// ─── the input cell-type contract (§5.7.2) ───────────────────────

/// `ST_CellType`. A closed enumeration, matched byte-exactly: a
/// differently-cased or entity-escaped spelling is not this enumeration
/// and classifying it as one would be a guess about what the producer
/// meant (`metadata.zig`'s decision 15, same reasoning).
pub const CellType = enum {
    number,
    shared_string,
    boolean,
    formula_string,
    inline_string,
    error_value,
    date,

    pub fn fromAttr(raw: ?[]const u8) ?CellType {
        const t = raw orelse return .number;
        if (std.mem.eql(u8, t, "n")) return .number;
        if (std.mem.eql(u8, t, "s")) return .shared_string;
        if (std.mem.eql(u8, t, "b")) return .boolean;
        if (std.mem.eql(u8, t, "str")) return .formula_string;
        if (std.mem.eql(u8, t, "inlineStr")) return .inline_string;
        if (std.mem.eql(u8, t, "e")) return .error_value;
        if (std.mem.eql(u8, t, "d")) return .date;
        return null;
    }

    /// Whether `classifyCell` *retains* the `<v>` bytes it is handed
    /// rather than parsing a value out of them.
    ///
    /// The walk decodes each `<v>` into one scratch buffer that the next
    /// `<v>` clears, so a retained slice has to leave that buffer before
    /// the next cell starts. `t="str"` leaves it through `decodeXstring`,
    /// which allocates; `t="e"` had nothing to allocate and kept
    /// pointing at bytes the following cell overwrote — a rich error
    /// spelling read back as whatever came next.
    ///
    /// Exhaustive on purpose: a carrier cannot be added without
    /// answering the question.
    pub fn retainsCachedText(self: CellType) bool {
        return switch (self) {
            .formula_string, .error_value => true,
            .number, .shared_string, .boolean, .date, .inline_string => false,
        };
    }
};

/// What a stored cell contributes to the merged view.
pub const InputCell = union(enum) {
    /// No value: no `<v>`, no `<is>`, and no formula.
    blank,
    /// A formula cell with no cached value. §5.6c seeds it; §5.6f reads
    /// it as blank before closure evaluation exists.
    uncached,
    number: f64,
    /// Borrows: from the scan arena for inline and `t="str"` text, from
    /// the shared-string table for `t="s"`.
    text: []const u8,
    boolean: bool,
    err: value.ErrorValue,

    pub fn scalar(self: InputCell) value.ScalarValue {
        return switch (self) {
            .blank, .uncached => .blank,
            .number => |n| value.ScalarValue.fromNumber(n),
            .text => |t| .{ .text = t },
            .boolean => |b| .{ .boolean = b },
            .err => |e| .{ .err = e },
        };
    }
};

/// Everything the contract needs to know about one `<c>`, gathered by
/// the scanner and classified separately so the rule can be fixtured
/// without a document around it.
pub const RawCell = struct {
    /// The raw `t` attribute, or null when absent.
    type_attr: ?[]const u8 = null,
    has_formula: bool = false,
    /// Present `<v>`, already entity-decoded. Null when the element is
    /// absent; an empty slice is a present-but-empty `<v>`.
    v: ?[]const u8 = null,
    /// Present `<is>`, already fully string-decoded.
    is: ?[]const u8 = null,
};

pub const Classified = union(enum) {
    ok: InputCell,
    refused: Refusal.Reason,
};

/// The mapping the reader lacks (`sheet_xml.zig:540-549` turns every
/// unknown `t` into a number).
///
/// **Precedence is normative**: a formula cell with no `<v>` is
/// *uncached*, decided FIRST and whatever `t` says. A formula cell may
/// legitimately carry `t="b"` or `t="e"` with no cached value yet, and
/// the b/e lexical tables below apply only when `<v>` is present.
pub fn classifyCell(
    raw: RawCell,
    strings: []const []const u8,
    fidelity: value.Fidelity,
    date_system: serial_date.DateSystem,
) Classified {
    const t = CellType.fromAttr(raw.type_attr) orelse
        return .{ .refused = .unknown_cell_type };

    // Inline strings carry their value in `<is>`, never in `<v>`.
    if (t == .inline_string) {
        if (raw.is) |s| return .{ .ok = .{ .text = s } };
        if (raw.has_formula) return .{ .ok = .uncached };
        return .{ .ok = .blank };
    }

    const v = raw.v orelse {
        if (raw.has_formula) return .{ .ok = .uncached };
        return switch (t) {
            // Nothing to interpret, and nothing that says there should
            // have been: an empty styled cell.
            .number, .shared_string, .formula_string => .{ .ok = .blank },
            // A `t` that promises a value the cell does not have.
            // `d` joins `b` and `e`: a `t` that promises a value the
            // cell does not have.
            .boolean, .error_value, .date => .{ .refused = .missing_cached_value },
            .inline_string => unreachable,
        };
    };

    return switch (t) {
        .number => blk: {
            const n = value.parseDecimal(fidelity, .cache_import, v) catch
                break :blk .{ .refused = .bad_number_cache };
            if (!std.math.isFinite(n)) break :blk .{ .refused = .bad_number_cache };
            break :blk .{ .ok = .{ .number = n } };
        },
        .shared_string => blk: {
            const idx = std.fmt.parseInt(u32, v, 10) catch
                break :blk .{ .refused = .bad_number_cache };
            if (idx >= strings.len) break :blk .{ .refused = .shared_string_index_out_of_range };
            break :blk .{ .ok = .{ .text = strings[idx] } };
        },
        .formula_string => .{ .ok = .{ .text = v } },
        .boolean => blk: {
            // Exactly `0` or `1`, whitespace-free. The reader's lax rule
            // (`xlsx.zig:1981-1985`) accepts more; a reader may be
            // generous about what it shows, an engine that writes values
            // back may not.
            if (std.mem.eql(u8, v, "1")) break :blk .{ .ok = .{ .boolean = true } };
            if (std.mem.eql(u8, v, "0")) break :blk .{ .ok = .{ .boolean = false } };
            break :blk .{ .refused = .bad_boolean_cache };
        },
        .error_value => blk: {
            const e = errorFromSpelling(v) orelse break :blk .{ .refused = .bad_error_cache };
            break :blk .{ .ok = .{ .err = e } };
        },
        // §5.7.2's normative lexical table, under the active epoch.
        // Both failure shapes — a form outside the table and a date the
        // epoch cannot express — are one pre-mutation refusal here; the
        // two are distinguishable in `serial_date`, and nothing above
        // this line acts on the difference.
        .date => blk: {
            const serial = serial_date.serialFromLexical(date_system, v) catch
                break :blk .{ .refused = .bad_date_cache };
            break :blk .{ .ok = .{ .number = serial } };
        },
        .inline_string => unreachable,
    };
}

/// The error lexical table: the frozen ten, or the tokenizer's
/// extensible rule (`#` + `[A-Za-z0-9_/.]` + `!`/`?`, bounded). Rich
/// spellings are preserved byte-exact, never normalized — §5.3a's
/// "preserved, never produced".
pub fn errorFromSpelling(text: []const u8) ?value.ErrorValue {
    if (value.KnownError.fromSpelling(text)) |k| return .{ .known = k };
    if (!isExtensibleErrorSpelling(text)) return null;
    return .{ .rich = text };
}

fn isExtensibleErrorSpelling(s: []const u8) bool {
    if (s.len < 3 or s.len > tokenizer.max_error_literal_bytes) return false;
    if (s[0] != '#') return false;
    const last = s[s.len - 1];
    if (last != '!' and last != '?') return false;
    for (s[1 .. s.len - 1]) |c| {
        const ok = (c >= 'A' and c <= 'Z') or (c >= 'a' and c <= 'z') or
            (c >= '0' and c <= '9') or c == '_' or c == '/' or c == '.';
        if (!ok) return false;
    }
    return true;
}

// ─── the worksheet walk ──────────────────────────────────────────

/// A byte range in the part, `[start, end)`.
///
/// `u32` rather than `usize` because `Limits.max_bytes` already refuses
/// a part a `u32` could not address, and a span that could not be
/// compared against a `Refusal.offset` would be a second addressing
/// scheme for one document.
pub const Span = struct {
    start: u32 = 0,
    end: u32 = 0,

    pub fn slice(self: Span, xml: []const u8) []const u8 {
        assert(self.end >= self.start);
        return xml[self.start..self.end];
    }
};

/// Where one `<c>` and its interpreted children sit in the part.
///
/// Recorded by the walk rather than re-found afterwards. The cached-value
/// patcher (M5b1) writes into these ranges and nowhere else, and a
/// locator that disagreed with the classifier by one byte would put a
/// value in the wrong element — so there is one parser, and it hands out
/// the coordinates it used.
pub const CellSpans = struct {
    /// `<c` through the `>` that closes `</c>`, or through the `>` of a
    /// self-closing `<c/>`.
    cell: Span = .{},
    /// Just past the `>` of the start tag. Equal to `cell.end` exactly
    /// when the element is self-closing.
    open_end: u32 = 0,
    /// The raw attribute region of the start tag — between the element
    /// name and the `/` or `>` that ends it. An empty span is normal and
    /// still positional: `<c/>` has no attributes and a place to put one.
    attrs: Span = .{},
    /// `t="…"`, first byte of the name through the closing quote. Null
    /// when the attribute is absent.
    type_attr: ?Span = null,
    /// `<v>` … `</v>`, whole element. Null when absent.
    v: ?Span = null,
    /// What sits between `<v>` and `</v>`. Null when the element is
    /// self-closing — a `<v/>` has no content region, and giving it one
    /// would mean inventing a position between two bytes that are not
    /// adjacent to any content.
    v_content: ?Span = null,
    /// `<f>` … `</f>`, whole element. Null when absent.
    f: ?Span = null,
    /// The raw value of `<f ref="…">`, between the quotes — the one
    /// byte range §5.8b's anchor-ref mutation may address. Null when
    /// the formula has no `ref`. The single exception to "no edit
    /// addresses a byte inside `spans.f`": the exception is this span
    /// and nothing else in the element.
    f_ref: ?Span = null,
    /// `<is>` … `</is>`, whole element. Null when absent.
    is: ?Span = null,

    pub fn selfClosing(self: CellSpans) bool {
        return self.open_end == self.cell.end;
    }
};

/// Every `<c>` the part contains, whether or not it contributes a value.
///
/// `Sheet.cells` drops an empty styled cell, because blank is the
/// *absence* of a cell in the merged view. A writer cannot afford the
/// same simplification: a `<c r="A1" s="3"/>` the model does not carry is
/// still a run of bytes sitting exactly where a new `<c r="A1">` would
/// have to go.
pub const CellSlot = struct {
    row: coords.Row,
    col: coords.Col,
    spans: CellSpans,
};

/// Where one `<row>` sits in the part, in document order (M7b1).
///
/// A tail `<c>` this run creates has to land INSIDE an existing row
/// element, and the insertion point is derived from these coordinates
/// plus the row's slots — recorded by the same walk that classified the
/// cells, for the reason `CellSpans` is: a second locator would disagree
/// by one byte and the disagreement would be a corrupted sheet.
pub const RowSlot = struct {
    /// One-based, as resolved — a provisional row (no `r`) takes its
    /// first cell's number; with no cells it keeps the
    /// predecessor-plus-one guess, exactly as the model placed it.
    number: u32,
    /// `<row` through the `>` that closes `</row>`, or through the
    /// `/>` of a self-closing row.
    element: Span,
    /// Just past the `>` of the start tag; `== element.end` exactly
    /// when the row is self-closing.
    open_end: u32,
    /// The `<` of `</row>` — where content can be appended — or
    /// `open_end` for a self-closing row.
    content_end: u32,
    /// The raw `spans="…"` value, between the quotes. A created `<c>`
    /// outside the declared span range would leave the attribute
    /// describing a row that no longer exists; §5.8b's approved set has
    /// no spans maintenance, so that shape refuses at the patcher.
    spans_attr: ?Span,

    pub fn selfClosing(self: RowSlot) bool {
        return self.open_end == self.element.end;
    }
};

/// Where the worksheet `<dimension>` sits (M7b1). Only the `ref`
/// value's bytes are ever rewritten — §5.8b's used-range expansion —
/// and only through these coordinates. Recorded whether or not any
/// mutation will need it; interpretation waits for the patcher, because
/// a stale dimension is tolerated everywhere EXCEPT under a spill that
/// extends the used range (`docs/plans/structural-edits.md:100`).
pub const DimensionSpans = struct {
    /// The start tag, `<dimension` through its `>`.
    element: Span,
    /// The raw `ref` value, between the quotes. Null when the
    /// attribute is absent.
    ref: ?Span,
};

/// One `<f>`, decoded. The attribute inventory and the shared/array
/// classification are M4b2's; what M4b1 owes is the decoded body and the
/// raw attributes preserved for that row to classify.
pub const Formula = struct {
    /// Decoded body (FORMULA carrier: entities only). Empty for a
    /// slave's bodiless `<f t="shared" si="0"/>`.
    text: []const u8,
    /// Raw `t` — null, "shared", "array" or "dataTable". Classified at
    /// M4b2.
    kind: ?[]const u8,
    ref: ?[]const u8,
    si: ?[]const u8,
    /// The whole raw attribute region, kept so M4b2's complete
    /// `CT_CellFormula` inventory has the bytes to inventory.
    raw_attrs: []const u8,
};

pub const SheetCell = struct {
    row: coords.Row,
    col: coords.Col,
    input: InputCell,
    formula: ?Formula,
    /// `c@cm` / `c@vm`, one-based, `0` = absent (M4a resolves them).
    cm: u32 = 0,
    vm: u32 = 0,
};

pub const Sheet = struct {
    arena: std.heap.ArenaAllocator,
    /// Row-major, one entry per occupied coordinate.
    ///
    /// **Backed by the scan's `gpa`, not the arena** (§9.1 M10d): the
    /// records are the scan's biggest single block and the staging path
    /// is done with them the moment the projection has copied what it
    /// reads — `releaseCells` lets that caller drop them while the
    /// arena (which the projection's slices still borrow) lives on.
    cells: []SheetCell,
    /// Row-major, one entry per `<c>` element — a superset of `cells`,
    /// because an empty styled cell has a position without having a
    /// value. Document order is preserved among equal coordinates, so a
    /// consumer that refuses duplicates can name the first one.
    ///
    /// **Backed by the scan's `gpa` like `cells`** (§9.1 M10f), but for
    /// the other half of the reason: the slot dupe was the request that
    /// minted the arena's next half-again chunk — 16.4 MiB of chunk
    /// around a 9.2 MiB record. Unlike the cell records the slots live
    /// to `deinit`: the projection and the patcher borrow them.
    slots: []CellSlot,
    /// `<mergeCells>`, normalized, in document order. Interpreted
    /// outside `<sheetData>` since M7a: §5.8a's merge row needs the
    /// geometry, and a spill decided without it would land on a merged
    /// range.
    merges: []coords.Range,
    /// Every `<row>` element, in document order (M7b1) — the geometry a
    /// tail `<c>` insertion is confined by. `gpa`-backed like `slots`,
    /// same lifetime.
    rows: []RowSlot,
    /// The worksheet `<dimension>`, or null when the part has none —
    /// in which case there is nothing to maintain and nothing to
    /// expand (M7b1).
    dimension: ?DimensionSpans,

    pub fn deinit(self: *Sheet) void {
        self.releaseCells();
        const gpa = self.arena.child_allocator;
        if (self.slots.len != 0) gpa.free(self.slots);
        if (self.rows.len != 0) gpa.free(self.rows);
        self.arena.deinit();
        self.* = undefined;
    }

    /// Free the cell records ahead of the rest of the scan. Everything
    /// else — slots, rows, merges, every decoded string the records
    /// pointed into — stays valid until `deinit`; only the record
    /// array itself dies. Idempotent, and `deinit` calls it.
    pub fn releaseCells(self: *Sheet) void {
        if (self.cells.len != 0) self.arena.child_allocator.free(self.cells);
        self.cells = &.{};
    }
};

pub const SheetResult = union(enum) {
    ok: Sheet,
    refused: Refusal,
};

/// `CT_Cell`'s complete attribute inventory. An attribute outside it
/// refuses: three of the six change what the cell's value *is*, so an
/// unrecognized seventh may too, and the alternative is discarding it
/// the way the typed reader does.
const cell_attrs = [_][]const u8{ "r", "s", "t", "cm", "vm", "ph" };

/// `CT_Row`'s complete attribute inventory.
const row_attrs = [_][]const u8{
    "r",        "spans",        "s",            "customFormat", "ht",
    "hidden",   "customHeight", "outlineLevel", "collapsed",    "thickTop",
    "thickBot", "ph",
};

/// Decode one `xl/worksheets/sheet*.xml` into cells.
///
/// The scan is one pass and refuses at the first thing it cannot
/// classify, before any of it reaches a caller — which is what makes
/// every refusal in this file pre-mutation by construction.
pub fn scanSheet(
    gpa: std.mem.Allocator,
    xml: []const u8,
    strings: []const []const u8,
    opts: Options,
) error{OutOfMemory}!SheetResult {
    assert(opts.limits.max_depth > 0);
    if (preflight(.worksheet, xml, opts.limits)) |r| return .{ .refused = r };

    var arena = std.heap.ArenaAllocator.init(gpa);
    // See `decodeSharedStrings`: a refusal returns normally, so the
    // arena has to be released on that path too.
    var keep = false;
    defer if (!keep) arena.deinit();
    const a = arena.allocator();

    // `gpa`-backed while they grow, arena-owned once they are done
    // (the exact-size dupes below): the four lists interleave their
    // growth, so inside the arena at most one of them could ever
    // resize in place and every abandoned buffer stayed resident.
    var cells: std.ArrayListUnmanaged(SheetCell) = .empty;
    defer cells.deinit(gpa);
    var slots: std.ArrayListUnmanaged(CellSlot) = .empty;
    defer slots.deinit(gpa);
    var merges: std.ArrayListUnmanaged(coords.Range) = .empty;
    defer merges.deinit(gpa);
    var rows: std.ArrayListUnmanaged(RowSlot) = .empty;
    defer rows.deinit(gpa);

    var ns_buf: [64]NsStack.Binding = undefined;
    const ns_cap = @min(ns_buf.len, opts.limits.max_namespace_bindings);
    var ns: NsStack = .{ .buf = ns_buf[0..ns_cap] };

    var w: SheetWalk = .{
        .a = a,
        .gpa = gpa,
        .xml = xml,
        .opts = opts,
        .strings = strings,
        .cells = &cells,
        .slots = &slots,
        .merges = &merges,
        .rows = &rows,
    };

    var depth: u32 = 0;
    var sc: Scanner = .init(xml);
    while (true) {
        const ev = sc.next() catch return .{
            .refused = .{ .reason = .malformed_xml, .offset = offsetOf(sc.i) },
        };
        const e = ev orelse break;

        switch (e) {
            .doctype => return .{ .refused = .{ .reason = .doctype_declaration, .offset = offsetOf(sc.i) } },
            .cdata => return .{ .refused = .{ .reason = .unexpected_cdata, .offset = offsetOf(sc.i) } },
            .text => |t| try w.onText(t, sc.i),
            .open, .self_closing => |el| {
                depth += 1;
                if (depth > opts.limits.max_depth) {
                    return .{ .refused = .{ .reason = .limit_exceeded, .limit = .depth } };
                }
                if (!ns.declare(el, depth)) {
                    return .{ .refused = .{ .reason = .limit_exceeded, .limit = .namespace_bindings } };
                }
                try w.onOpen(&ns, el, depth, e == .self_closing, sc.i);
                if (e == .self_closing) {
                    try w.onClose(depth, sc.i);
                    ns.popTo(depth - 1);
                    depth -= 1;
                }
            },
            .close => {
                if (depth == 0) {
                    return .{ .refused = .{ .reason = .malformed_xml, .offset = offsetOf(sc.i) } };
                }
                try w.onClose(depth, sc.i);
                ns.popTo(depth - 1);
                depth -= 1;
            },
        }
        if (w.refusal) |r| return .{ .refused = r };
    }

    if (depth != 0) {
        return .{ .refused = .{ .reason = .malformed_xml, .offset = offsetOf(sc.i) } };
    }

    // Each growing list is dropped the moment its exact-size copy
    // exists: at the copy instant the biggest list would otherwise be
    // resident twice, and §9.1 measures that instant. The cell records,
    // slots and rows go to `gpa`, not the arena — `Sheet.releaseCells`
    // is the cells' point, and the two fixed dupes were the requests
    // that bought the arena's half-again chunks (§9.1 M10f); only the
    // decoded strings, whose total no count predicts, stay laddered.
    const items = try gpa.dupe(SheetCell, cells.items);
    // The duplicate-cell refusal below RETURNS normally, so the same
    // `keep` discipline that guards the arena guards this block too.
    defer if (!keep) gpa.free(items);
    cells.clearAndFree(gpa);
    std.mem.sortUnstable(SheetCell, items, {}, lessThanCell);
    // Stable, so two `<c>` claiming one coordinate stay in document
    // order for the consumer that has to name the first of them.
    const slot_items = try gpa.dupe(CellSlot, slots.items);
    defer if (!keep) gpa.free(slot_items);
    slots.clearAndFree(gpa);
    std.mem.sort(CellSlot, slot_items, {}, lessThanSlot);
    // Two `<c>` at one coordinate: last-wins and first-wins are both
    // defensible readings, so neither is chosen silently.
    var i: usize = 1;
    while (i < items.len) : (i += 1) {
        if (items[i - 1].row == items[i].row and items[i - 1].col == items[i].col) {
            return .{ .refused = .{
                .reason = .duplicate_cell,
                .cell = .{ .row = items[i].row.oneBased(), .col = items[i].col.zeroBased() },
            } };
        }
    }

    // Materialized BEFORE the result literal: `.arena = arena` copies
    // the arena's state first, so an allocation in a later field
    // initializer that opens a fresh chunk would be known only to the
    // local copy — and leak when the returned one deinits.
    const merge_items = try a.dupe(coords.Range, merges.items);
    merges.clearAndFree(gpa);
    const row_items = try gpa.dupe(RowSlot, rows.items);
    defer if (!keep) gpa.free(row_items);
    rows.clearAndFree(gpa);

    keep = true;
    return .{ .ok = .{
        .arena = arena,
        .cells = items,
        .slots = slot_items,
        .merges = merge_items,
        .rows = row_items,
        .dimension = w.dimension,
    } };
}

fn lessThanCell(_: void, x: SheetCell, y: SheetCell) bool {
    if (x.row.oneBased() != y.row.oneBased()) return x.row.oneBased() < y.row.oneBased();
    return x.col.zeroBased() < y.col.zeroBased();
}

fn lessThanSlot(_: void, x: CellSlot, y: CellSlot) bool {
    if (x.row.oneBased() != y.row.oneBased()) return x.row.oneBased() < y.row.oneBased();
    return x.col.zeroBased() < y.col.zeroBased();
}

/// The walk's state. Split out of `scanSheet` so the state machine is
/// one object with named transitions rather than a dozen locals that
/// every branch can quietly desynchronize.
const SheetWalk = struct {
    a: std.mem.Allocator,
    /// Backs the four *growing* lists, while `a` (the sheet arena)
    /// keeps the payload dupes. A list that grows inside an arena
    /// strands every backing buffer it abandons until the arena dies —
    /// on a 100k-cell sheet that held ~3× the scan's live bytes
    /// (§9.1's profile). The lists move to the arena as one exact-size
    /// copy each when the scan finishes.
    gpa: std.mem.Allocator,
    /// The part, kept so a slice handed out by the scanner can be turned
    /// back into the offsets it came from.
    xml: []const u8,
    opts: Options,
    strings: []const []const u8,
    cells: *std.ArrayListUnmanaged(SheetCell),
    slots: *std.ArrayListUnmanaged(CellSlot),

    refusal: ?Refusal = null,

    /// Depth of `<sheetData>` once inside it.
    sheet_data_depth: ?u32 = null,
    /// Depth of `<mergeCells>` once inside it (M7a, §5.8a merge row).
    merge_depth: ?u32 = null,
    merges: *std.ArrayListUnmanaged(coords.Range),
    rows: *std.ArrayListUnmanaged(RowSlot),
    /// The `<dimension>` once seen (M7b1). A second one refuses — the
    /// schema allows one, and a patcher with two candidate ranges to
    /// rewrite would have to pick one silently.
    dimension: ?DimensionSpans = null,
    /// Depth of a subtree being skipped wholesale (foreign content or
    /// an inert main-namespace element outside `<sheetData>`).
    skip_depth: ?u32 = null,

    row_depth: ?u32 = null,
    row_index: u32 = 0,
    /// The current `<row>`'s start-tag coordinates, pending its close
    /// (M7b1's geometry record).
    row_start: u32 = 0,
    row_open_end: u32 = 0,
    row_spans_attr: ?Span = null,
    /// True between a `<row>` with no `r` and its first `<c>`, while
    /// the row number is still the predecessor-plus-one guess.
    row_provisional: bool = false,
    cursor: CellCursor = .{},

    cell_depth: ?u32 = null,
    cell_row: u32 = 0,
    cell_col: coords.Col = @enumFromInt(0),
    cell_type: ?[]const u8 = null,
    cell_cm: u32 = 0,
    cell_vm: u32 = 0,

    v_depth: ?u32 = null,
    f_depth: ?u32 = null,
    is_depth: ?u32 = null,
    /// Depth of a `<t>` whose text counts toward the inline string.
    t_depth: ?u32 = null,
    phonetic_depth: ?u32 = null,

    has_v: bool = false,
    has_is: bool = false,
    has_f: bool = false,
    v_text: std.ArrayListUnmanaged(u8) = .empty,
    f_text: std.ArrayListUnmanaged(u8) = .empty,
    is_text: std.ArrayListUnmanaged(u8) = .empty,
    formula: ?Formula = null,

    /// The current `<c>`'s byte ranges, filled in as the walk passes
    /// each of them and handed to `slots` when the element closes.
    spans: CellSpans = .{},
    /// Just past the `>` of the open `<v>` tag, so the content region is
    /// known before the close tag arrives.
    v_open_end: u32 = 0,
    v_self_closing: bool = false,

    fn refuse(self: *SheetWalk, reason: Refusal.Reason, offset: usize) void {
        if (self.refusal != null) return;
        self.refusal = .{ .reason = reason, .offset = offsetOf(offset) };
    }

    fn refuseAtCell(self: *SheetWalk, reason: Refusal.Reason) void {
        if (self.refusal != null) return;
        self.refusal = .{ .reason = reason, .cell = .{
            .row = self.cell_row,
            .col = self.cell_col.zeroBased(),
        } };
    }

    fn onText(self: *SheetWalk, t: []const u8, offset: usize) error{OutOfMemory}!void {
        if (self.skip_depth != null) return;
        // A `<t>` inside `<rPh>` is furigana: real text that is
        // deliberately not part of the value. Dropping it is the rule;
        // refusing it would refuse every phonetically annotated
        // workbook.
        if (self.t_depth != null and self.phonetic_depth != null) return;
        const sink: ?*std.ArrayListUnmanaged(u8) = if (self.v_depth != null)
            &self.v_text
        else if (self.f_depth != null)
            &self.f_text
        else if (self.t_depth != null)
            &self.is_text
        else
            null;
        if (sink) |list| {
            const decoded = decodeEntities(self.a, t) catch |err| {
                self.refusal = try textRefusal(err, offset);
                return;
            };
            try list.appendSlice(self.a, decoded);
            return;
        }
        // Between elements, only whitespace is legal. Text anywhere else
        // means the document is not the shape the schema describes, and
        // guessing which element it belonged to is how a value lands in
        // the wrong cell.
        if (!isWhitespaceOnly(t)) self.refuse(.unexpected_text, offset);
    }

    fn onOpen(
        self: *SheetWalk,
        ns: *const NsStack,
        el: Element,
        depth: u32,
        self_closing: bool,
        /// Just past the `>` that ended this start tag.
        open_end: usize,
    ) error{OutOfMemory}!void {
        if (self.skip_depth != null) return;

        // Outside `<sheetData>` almost nothing is interpreted: the one
        // exception is `<mergeCells>` (M7a) — §5.8a's merge row needs
        // the geometry, so the two merge elements are read where
        // everything else is got past.
        const sd = self.sheet_data_depth orelse {
            if (self.merge_depth) |md| {
                if (inMainNs(ns, el.name) and depth == md + 1 and
                    std.mem.eql(u8, el.local(), "mergeCell"))
                {
                    try self.mergeCell(el);
                    // The schema gives `<mergeCell>` no children; skip
                    // to its close rather than guessing about any.
                    if (!self_closing) self.skip_depth = depth;
                } else if (!inMainNs(ns, el.name)) {
                    // Foreign content inside an interpreted element is
                    // skipped like foreign content anywhere outside
                    // `<sheetData>` (M4b1 decision 8).
                    if (!self_closing) self.skip_depth = depth;
                } else {
                    // A main-namespace element the schema does not put
                    // here could be wrapping merges; refuse rather than
                    // drop it (M4b1's discipline for interpreted
                    // elements).
                    self.refuse(.unexpected_element, el.offset);
                }
                return;
            }
            if (isMain(ns, el, "sheetData")) {
                self.sheet_data_depth = depth;
            } else if (isMain(ns, el, "dimension") and depth == 2) {
                // Interpreted since M7b1: §5.8b's used-range expansion
                // rewrites the `ref` value, so the walk records where it
                // is. `CT_SheetDimension` has exactly one attribute.
                if (self.unexpectedAttr(el, &.{"ref"})) return;
                if (self.dimension != null) {
                    self.refuse(.unexpected_element, el.offset);
                    return;
                }
                self.dimension = .{
                    .element = .{ .start = offsetOf(el.offset), .end = offsetOf(open_end) },
                    .ref = if (el.attr("ref")) |rv| spanOfSub(self.xml, rv) else null,
                };
                // The schema gives `<dimension>` no children; skip to
                // its close rather than guessing about any.
                if (!self_closing) self.skip_depth = depth;
            } else if (isMain(ns, el, "mergeCells") and depth == 2) {
                // `count` is recorded nowhere and enforced nowhere —
                // M4a decision 11's rule: producers miscount, and
                // resolution is by position.
                if (self.unexpectedAttr(el, &.{"count"})) return;
                if (!self_closing) self.merge_depth = depth;
            } else if (depth > 1 and !self_closing) {
                self.skip_depth = depth;
            }
            return;
        };

        if (!inMainNs(ns, el.name)) {
            // Inside `<sheetData>` a foreign element could be wrapping
            // cells (`mc:AlternateContent`); skipping it would drop
            // them and interpreting it would be a guess.
            self.refuse(.foreign_element, el.offset);
            return;
        }

        const local = el.local();
        // `<row>` — a child of `<sheetData>`.
        if (depth == sd + 1) {
            if (std.mem.eql(u8, local, "row")) {
                self.startRow(el, open_end);
            } else if (std.mem.eql(u8, local, "extLst")) {
                self.skip_depth = depth;
            } else {
                self.refuse(.unexpected_element, el.offset);
            }
            return;
        }

        // `<c>` — a child of `<row>`.
        if (self.row_depth != null and depth == self.row_depth.? + 1) {
            if (std.mem.eql(u8, local, "c")) {
                self.startCell(el, open_end);
            } else if (std.mem.eql(u8, local, "extLst")) {
                self.skip_depth = depth;
            } else {
                self.refuse(.unexpected_element, el.offset);
            }
            return;
        }

        // Children of `<c>`: `<v>`, `<f>`, `<is>`.
        if (self.cell_depth != null and depth == self.cell_depth.? + 1) {
            if (std.mem.eql(u8, local, "v")) {
                self.has_v = true;
                self.v_depth = depth;
                self.v_text.clearRetainingCapacity();
                self.spans.v = .{ .start = offsetOf(el.offset) };
                self.spans.v_content = null;
                self.v_open_end = offsetOf(open_end);
                self.v_self_closing = self_closing;
            } else if (std.mem.eql(u8, local, "f")) {
                self.has_f = true;
                self.f_depth = depth;
                self.f_text.clearRetainingCapacity();
                self.formula = .{
                    .text = "",
                    .kind = el.attr("t"),
                    .ref = el.attr("ref"),
                    .si = el.attr("si"),
                    .raw_attrs = el.attrs,
                };
                self.spans.f = .{ .start = offsetOf(el.offset) };
                self.spans.f_ref = if (el.attr("ref")) |rv| spanOfSub(self.xml, rv) else null;
            } else if (std.mem.eql(u8, local, "is")) {
                self.has_is = true;
                self.is_depth = depth;
                self.is_text.clearRetainingCapacity();
                self.spans.is = .{ .start = offsetOf(el.offset) };
            } else if (std.mem.eql(u8, local, "extLst")) {
                self.skip_depth = depth;
            } else {
                self.refuse(.unexpected_element, el.offset);
            }
            return;
        }

        // Inside `<is>`: `CT_Rst` — `t`, `r`, `rPh`, `phoneticPr`.
        if (self.is_depth != null and depth > self.is_depth.?) {
            if (std.mem.eql(u8, local, "t")) {
                self.t_depth = depth;
            } else if (std.mem.eql(u8, local, "rPh")) {
                if (self.phonetic_depth == null) self.phonetic_depth = depth;
            } else if (std.mem.eql(u8, local, "r")) {
                // A run contributes its `<t>`; the element itself
                // carries nothing.
            } else if (std.mem.eql(u8, local, "rPr") or
                std.mem.eql(u8, local, "phoneticPr"))
            {
                // Formatting. Skipped wholesale so a `<rFont val="…"/>`
                // cannot be mistaken for text.
                self.skip_depth = depth;
            } else {
                self.refuse(.unexpected_element, el.offset);
            }
            return;
        }

        // Anything else inside `<sheetData>` has no legal position.
        self.refuse(.unexpected_element, el.offset);
    }

    fn onClose(
        self: *SheetWalk,
        depth: u32,
        /// Just past the `>` that ended this element.
        end: usize,
    ) error{OutOfMemory}!void {
        if (self.skip_depth) |sd| {
            if (sd == depth) self.skip_depth = null;
            return;
        }

        if (self.t_depth == depth) self.t_depth = null;
        if (self.phonetic_depth == depth) self.phonetic_depth = null;
        if (self.v_depth == depth) {
            self.v_depth = null;
            if (self.spans.v) |*v| {
                v.end = offsetOf(end);
                // A `<v/>` has no content region. Everything else does,
                // and its far edge is the `<` of the close tag: a close
                // tag carries no attributes, so nothing can quote a `<`
                // between here and there.
                if (!self.v_self_closing) {
                    self.spans.v_content = .{
                        .start = self.v_open_end,
                        .end = closeTagStart(self.xml, end),
                    };
                }
            }
        }
        if (self.f_depth == depth) {
            self.f_depth = null;
            if (self.formula) |*f| f.text = try self.a.dupe(u8, self.f_text.items);
            if (self.spans.f) |*f| f.end = offsetOf(end);
        }
        if (self.is_depth == depth) {
            self.is_depth = null;
            if (self.spans.is) |*is| is.end = offsetOf(end);
        }
        if (self.cell_depth == depth) try self.finishCell(end);
        if (self.row_depth == depth) {
            self.row_depth = null;
            self.cursor.startRow();
            try self.finishRow(end);
        }
        if (self.sheet_data_depth == depth) self.sheet_data_depth = null;
        if (self.merge_depth == depth) self.merge_depth = null;
    }

    /// One `<mergeCell ref="…"/>` (M7a). The schema gives it exactly
    /// one attribute, and a range this grid cannot address is a
    /// malformed value — the same refusal a bad `<row r>` takes. A `$`
    /// is accepted because it is unambiguous: refusing one would refuse
    /// a file Excel opens.
    fn mergeCell(self: *SheetWalk, el: Element) error{OutOfMemory}!void {
        if (self.unexpectedAttr(el, &.{"ref"})) return;
        const raw = el.attr("ref") orelse {
            self.refuse(.bad_attribute_value, el.offset);
            return;
        };
        const range = coords.parseRange(raw, .{ .dollar = .accept }) catch {
            self.refuse(.bad_attribute_value, el.offset);
            return;
        };
        try self.merges.append(self.gpa, range.normalized());
    }

    fn startRow(self: *SheetWalk, el: Element, open_end: usize) void {
        if (self.unexpectedAttr(el, &row_attrs)) return;
        self.row_depth = self.sheet_data_depth.? + 1;
        self.cursor.startRow();
        self.row_start = offsetOf(el.offset);
        self.row_open_end = offsetOf(open_end);
        self.row_spans_attr = if (el.attr("spans")) |sv| spanOfSub(self.xml, sv) else null;

        if (el.attr("r")) |r| {
            const idx = std.fmt.parseInt(u32, r, 10) catch {
                self.refuse(.bad_attribute_value, el.offset);
                return;
            };
            // `r="0"` is out of the grid. The typed reader skips such a
            // row (`sheet_xml.zig:455`); dropping cells is exactly what
            // an engine that writes values back must not do, so this
            // refuses instead (`tests/corpus/poi_poc_shared_strings.xlsx`).
            if (idx == 0 or idx > coords.max_row) {
                self.refuse(.bad_attribute_value, el.offset);
                return;
            }
            self.row_index = idx;
            self.row_provisional = false;
            self.cursor.previous_row = idx;
            return;
        }

        // No `r`: the row after its predecessor. Provisional until the
        // first cell, because a producer that omits `r` on the row may
        // still put it on the cells — and then the cells are the
        // authority (`src/xlsx.zig:1863` recovers it the same way).
        switch (self.cursor.implicitRow()) {
            .refused => |reason| {
                self.refuse(reason, el.offset);
                return;
            },
            .ok => |idx| {
                self.row_index = idx;
                self.row_provisional = true;
                self.cursor.previous_row = idx;
            },
        }
    }

    fn startCell(self: *SheetWalk, el: Element, open_end: usize) void {
        if (self.unexpectedAttr(el, &cell_attrs)) return;
        self.spans = .{
            .cell = .{ .start = offsetOf(el.offset) },
            .open_end = offsetOf(open_end),
            .attrs = spanOfSub(self.xml, el.attrs),
            .type_attr = attrSpanIn(self.xml, el.attrs, "t"),
        };
        const r_attr = el.attr("r");
        if (self.row_provisional) {
            // The first cell fixes the row for the whole row element;
            // every later cell must agree with it, which is the check
            // `CellCursor.resolve` already makes.
            if (r_attr) |ref| {
                const parsed = coords.parseCell(ref, .{ .case = .insensitive }) catch {
                    self.refuse(.malformed_cell_ref, el.offset);
                    return;
                };
                self.row_index = parsed.row.oneBased();
                self.cursor.previous_row = self.row_index;
            }
            self.row_provisional = false;
        }
        switch (self.cursor.resolve(r_attr, self.row_index)) {
            .refused => |reason| {
                self.refusal = .{
                    .reason = reason,
                    .offset = offsetOf(el.offset),
                    .cell = .{ .row = self.row_index, .col = 0 },
                };
                return;
            },
            .ok => |col| self.cell_col = col,
        }
        self.cell_depth = self.row_depth.? + 1;
        self.cell_row = self.row_index;
        self.cell_type = el.attr("t");
        self.cell_cm = parseIndexAttr(el, "cm") orelse {
            self.refuse(.bad_attribute_value, el.offset);
            return;
        };
        self.cell_vm = parseIndexAttr(el, "vm") orelse {
            self.refuse(.bad_attribute_value, el.offset);
            return;
        };
        self.has_v = false;
        self.has_is = false;
        self.has_f = false;
        self.formula = null;
        self.v_text.clearRetainingCapacity();
        self.f_text.clearRetainingCapacity();
        self.is_text.clearRetainingCapacity();
    }

    /// Record the closed `<row>`'s geometry (M7b1). The number is the
    /// resolved one — a provisional row was fixed by its first cell
    /// before any slot was recorded, so the geometry and the slots
    /// agree on where row N is.
    fn finishRow(self: *SheetWalk, end: usize) error{OutOfMemory}!void {
        if (self.refusal != null) return;
        // Rows are bounded like cells: a part with more row elements
        // than the cell ceiling is past the shape bound either way.
        if (self.rows.items.len >= self.opts.limits.max_modeled_cells) {
            self.refusal = .{ .reason = .limit_exceeded, .limit = .modeled_cells };
            return;
        }
        const e = offsetOf(end);
        const self_closing = self.row_open_end == e;
        try self.rows.append(self.gpa, .{
            .number = self.row_index,
            .element = .{ .start = self.row_start, .end = e },
            .open_end = self.row_open_end,
            .content_end = if (self_closing) self.row_open_end else closeTagStart(self.xml, end),
            .spans_attr = self.row_spans_attr,
        });
    }

    fn finishCell(self: *SheetWalk, end: usize) error{OutOfMemory}!void {
        self.cell_depth = null;
        if (self.refusal != null) return;
        if (self.cells.items.len >= self.opts.limits.max_modeled_cells or
            self.slots.items.len >= self.opts.limits.max_modeled_cells)
        {
            self.refusal = .{ .reason = .limit_exceeded, .limit = .modeled_cells };
            return;
        }

        // The slot is recorded before the classification, because a
        // `<c>` occupies its bytes whatever the classification decides
        // — including the empty styled cell the merged view drops.
        self.spans.cell.end = offsetOf(end);
        try self.slots.append(self.gpa, .{
            .row = coords.Row.fromOneBased(self.cell_row) catch unreachable,
            .col = self.cell_col,
            .spans = self.spans,
        });

        // The inline string is a STRING carrier: its runs are already
        // entity-decoded, so what is left is the ST_Xstring pass over
        // the concatenation.
        const is_text: ?[]const u8 = if (self.has_is)
            decodeXstring(self.a, self.is_text.items) catch |err| {
                self.refusal = try textRefusal(err, 0);
                return;
            }
        else
            null;

        // `t="str"` is a STRING carrier too; every other `<v>` is
        // lexical and stops at entity decoding. What every arm has in
        // common is that a value the classification *keeps* must not
        // still be pointing into the scratch buffer when the next `<v>`
        // clears it — see `CellType.retainsCachedText`.
        const v_raw: ?[]const u8 = if (self.has_v) blk: {
            const t = CellType.fromAttr(self.cell_type) orelse break :blk self.v_text.items;
            if (t == .formula_string) {
                break :blk decodeXstring(self.a, self.v_text.items) catch |err| {
                    self.refusal = try textRefusal(err, 0);
                    return;
                };
            }
            if (!t.retainsCachedText()) break :blk self.v_text.items;
            break :blk try self.a.dupe(u8, self.v_text.items);
        } else null;

        const classified = classifyCell(.{
            .type_attr = self.cell_type,
            .has_formula = self.has_f,
            .v = v_raw,
            .is = is_text,
        }, self.strings, self.opts.fidelity, self.opts.date_system);

        const input = switch (classified) {
            .refused => |reason| {
                self.refuseAtCell(reason);
                return;
            },
            .ok => |cell| cell,
        };

        // A cell with nothing in it is not in the merged view at all:
        // blank is the *absence* of a cell (`env.zig:486-489`), and
        // storing one would give two spellings for one state.
        if (input == .blank and !self.has_f) return;

        try self.cells.append(self.gpa, .{
            .row = coords.Row.fromOneBased(self.cell_row) catch unreachable,
            .col = self.cell_col,
            .input = input,
            .formula = self.formula,
            .cm = self.cell_cm,
            .vm = self.cell_vm,
        });
    }

    /// True (and refuses) when the element carries a main-namespace
    /// attribute outside `allowed`. Prefixed attributes are exempt:
    /// `x14ac:dyDescent` is on most rows Excel writes, and refusing a
    /// namespaced extension would refuse the ordinary case.
    fn unexpectedAttr(self: *SheetWalk, el: Element, allowed: []const []const u8) bool {
        var it = el.attrIterator();
        while (it.next()) |a| {
            if (a.prefix().len != 0) continue;
            if (std.mem.eql(u8, a.qname, "xmlns")) continue;
            const name = a.local();
            var ok = false;
            for (allowed) |candidate| {
                if (std.mem.eql(u8, name, candidate)) {
                    ok = true;
                    break;
                }
            }
            if (!ok) {
                self.refuse(.unexpected_attribute, el.offset);
                return true;
            }
        }
        return false;
    }
};

/// An `xsd:unsignedInt` attribute with a schema default. Null means the
/// attribute was present and not that type — which for a table's row
/// counts is a geometry nobody can compute.
fn parseCount(raw: ?[]const u8, default: u32) ?u32 {
    const s = raw orelse return default;
    return std.fmt.parseInt(u32, s, 10) catch null;
}

/// `cm`/`vm` are `xsd:unsignedInt` with default `0`. Null means the
/// attribute was present and not that type.
fn parseIndexAttr(el: Element, name: []const u8) ?u32 {
    const raw = el.attr(name) orelse return 0;
    return std.fmt.parseInt(u32, raw, 10) catch null;
}

// ─── table parts ─────────────────────────────────────────────────

pub const TableColumn = struct {
    /// Decoded (STRING carrier).
    name: []const u8,
    /// Decoded (FORMULA carrier), when the column is calculated.
    calculated_formula: ?[]const u8,
    totals_formula: ?[]const u8,
    /// The raw attribute regions of the two producer elements, kept for
    /// the same reason `Formula.raw_attrs` is: M4b3's `CT_TableFormula`
    /// inventory needs the bytes to inventory, and a reader that
    /// dropped them would have nothing to refuse an unknown attribute
    /// with. Empty when the element is absent.
    calculated_attrs: []const u8 = "",
    totals_attrs: []const u8 = "",
};

pub const Table = struct {
    arena: std.heap.ArenaAllocator,
    /// Decoded `name`, and `displayName` when it differs.
    name: []const u8,
    display_name: []const u8,
    /// Raw `ref` (an A1 range; borrows from the input bytes).
    ref: []const u8,
    /// `CT_Table@headerRowCount` / `@totalsRowCount`, with the schema's
    /// defaults (1 and 0). They decide which rows a producer covers,
    /// which is the whole of M4b3's member check.
    header_rows: u32 = 1,
    totals_rows: u32 = 0,
    columns: []TableColumn,

    pub fn deinit(self: *Table) void {
        self.arena.deinit();
        self.* = undefined;
    }
};

pub const TableResult = union(enum) {
    ok: Table,
    refused: Refusal,
};

/// Decode one `xl/tables/table*.xml`. The symbol layer needs the name
/// and the column names; the two formula elements are here because they
/// are FORMULA carriers and the split has to be complete to be a split.
pub fn scanTable(
    gpa: std.mem.Allocator,
    xml: []const u8,
    opts: Options,
) error{OutOfMemory}!TableResult {
    if (preflight(.table, xml, opts.limits)) |r| return .{ .refused = r };

    var arena = std.heap.ArenaAllocator.init(gpa);
    // See `decodeSharedStrings`.
    var keep = false;
    defer if (!keep) arena.deinit();
    const a = arena.allocator();

    var ns_buf: [64]NsStack.Binding = undefined;
    const cap = @min(ns_buf.len, opts.limits.max_namespace_bindings);
    var ns: NsStack = .{ .buf = ns_buf[0..cap] };

    var columns: std.ArrayListUnmanaged(TableColumn) = .empty;
    var name: []const u8 = "";
    var display_name: []const u8 = "";
    var ref: []const u8 = "";
    var header_rows: u32 = 1;
    var totals_rows: u32 = 0;

    var depth: u32 = 0;
    var text_sink: ?*std.ArrayListUnmanaged(u8) = null;
    var calc: std.ArrayListUnmanaged(u8) = .empty;
    var totals: std.ArrayListUnmanaged(u8) = .empty;
    var pending: ?TableColumn = null;
    var column_depth: ?u32 = null;

    var sc: Scanner = .init(xml);
    while (true) {
        const ev = sc.next() catch return .{
            .refused = .{ .reason = .malformed_xml, .offset = offsetOf(sc.i) },
        };
        const e = ev orelse break;
        switch (e) {
            .doctype => return .{ .refused = .{ .reason = .doctype_declaration, .offset = offsetOf(sc.i) } },
            .cdata => return .{ .refused = .{ .reason = .unexpected_cdata, .offset = offsetOf(sc.i) } },
            .text => |t| {
                if (text_sink) |sink| {
                    const decoded = decodeEntities(a, t) catch |err| return .{
                        .refused = try textRefusal(err, sc.i),
                    };
                    try sink.appendSlice(a, decoded);
                }
            },
            .open, .self_closing => |el| {
                depth += 1;
                if (depth > opts.limits.max_depth) {
                    return .{ .refused = .{ .reason = .limit_exceeded, .limit = .depth } };
                }
                if (!ns.declare(el, depth)) {
                    return .{ .refused = .{ .reason = .limit_exceeded, .limit = .namespace_bindings } };
                }
                if (isMain(&ns, el, "table") and depth == 1) {
                    name = decodeAt(a, .table_name, el.attr("name") orelse "") catch |err| return .{
                        .refused = try textRefusal(err, el.offset),
                    };
                    display_name = decodeAt(a, .table_name, el.attr("displayName") orelse "") catch |err| return .{
                        .refused = try textRefusal(err, el.offset),
                    };
                    ref = el.attr("ref") orelse "";
                    // A count that is not `xsd:unsignedInt` is a table
                    // whose geometry nobody can compute, and computing
                    // it wrong places a producer over the header.
                    header_rows = parseCount(el.attr("headerRowCount"), 1) orelse return .{
                        .refused = .{ .reason = .bad_attribute_value, .offset = offsetOf(el.offset) },
                    };
                    totals_rows = parseCount(el.attr("totalsRowCount"), 0) orelse return .{
                        .refused = .{ .reason = .bad_attribute_value, .offset = offsetOf(el.offset) },
                    };
                } else if (isMain(&ns, el, "tableColumn")) {
                    const col_name = decodeAt(a, .table_column_name, el.attr("name") orelse "") catch |err| return .{
                        .refused = try textRefusal(err, el.offset),
                    };
                    pending = .{ .name = col_name, .calculated_formula = null, .totals_formula = null };
                    calc.clearRetainingCapacity();
                    totals.clearRetainingCapacity();
                    column_depth = depth;
                } else if (isMain(&ns, el, "calculatedColumnFormula")) {
                    text_sink = &calc;
                    if (pending) |*p| p.calculated_attrs = el.attrs;
                } else if (isMain(&ns, el, "totalsRowFormula")) {
                    text_sink = &totals;
                    if (pending) |*p| p.totals_attrs = el.attrs;
                }
                if (e == .self_closing) {
                    text_sink = null;
                    if (column_depth == depth) {
                        if (pending) |p| try columns.append(a, p);
                        pending = null;
                        column_depth = null;
                    }
                    ns.popTo(depth - 1);
                    depth -= 1;
                }
            },
            .close => {
                if (depth == 0) {
                    return .{ .refused = .{ .reason = .malformed_xml, .offset = offsetOf(sc.i) } };
                }
                text_sink = null;
                if (column_depth == depth) {
                    if (pending) |p| {
                        var c = p;
                        if (calc.items.len > 0) c.calculated_formula = try a.dupe(u8, calc.items);
                        if (totals.items.len > 0) c.totals_formula = try a.dupe(u8, totals.items);
                        try columns.append(a, c);
                    }
                    pending = null;
                    column_depth = null;
                }
                ns.popTo(depth - 1);
                depth -= 1;
            },
        }
    }

    const owned_columns = try columns.toOwnedSlice(a);
    keep = true;
    return .{ .ok = .{
        .arena = arena,
        .name = name,
        .display_name = if (display_name.len > 0) display_name else name,
        .ref = ref,
        .header_rows = header_rows,
        .totals_rows = totals_rows,
        .columns = owned_columns,
    } };
}

// ─── tests ───────────────────────────────────────────────────────

const testing = std.testing;

const ns_attr = " xmlns=\"" ++ ns_main ++ "\"";

fn sheetXml(comptime body: []const u8) []const u8 {
    return "<worksheet" ++ ns_attr ++ "><sheetData>" ++ body ++ "</sheetData></worksheet>";
}

fn scanOk(xml: []const u8, strings: []const []const u8) !Sheet {
    return switch (try scanSheet(testing.allocator, xml, strings, .{})) {
        .ok => |s| s,
        .refused => |ref| {
            std.debug.print("unexpected refusal: {any}\n", .{ref});
            return error.TestUnexpectedRefusal;
        },
    };
}

fn scanRefusal(xml: []const u8, strings: []const []const u8) !Refusal {
    return switch (try scanSheet(testing.allocator, xml, strings, .{})) {
        .refused => |ref| ref,
        .ok => |s| {
            var sheet = s;
            sheet.deinit();
            return error.TestExpectedRefusal;
        },
    };
}

// ─── carrier split ───────────────────────────────────────────────

test "carrier: every site is classified, and the two named classes hold" {
    // The split is the row's hardest correctness claim, so it is
    // asserted as a table rather than trusted to a switch nobody reads.
    inline for (@typeInfo(Site).@"enum".fields) |f| {
        const site: Site = @enumFromInt(f.value);
        _ = site.carrier();
    }
    try testing.expectEqual(Carrier.formula, Site.cell_formula_body.carrier());
    try testing.expectEqual(Carrier.formula, Site.defined_name_body.carrier());
    try testing.expectEqual(Carrier.formula, Site.table_calculated_column_formula.carrier());
    try testing.expectEqual(Carrier.formula, Site.table_totals_row_formula.carrier());
    try testing.expectEqual(Carrier.string, Site.shared_string.carrier());
    try testing.expectEqual(Carrier.string, Site.inline_string.carrier());
    try testing.expectEqual(Carrier.string, Site.formula_string_value.carrier());
    try testing.expectEqual(Carrier.lexical, Site.number_value.carrier());
    try testing.expect(!Site.cell_formula_body.carrier().appliesXstring());
    try testing.expect(Site.shared_string.carrier().appliesXstring());
}

test "decode: a literal _x0041_ survives a formula carrier and decodes in a string carrier" {
    // The asymmetry, at its smallest. Same seven bytes, two sites, two
    // answers — and the formula answer is the one a second decode pass
    // would silently destroy.
    const raw = "_x0041_";
    const as_formula = try decodeAt(testing.allocator, .cell_formula_body, raw);
    defer testing.allocator.free(as_formula);
    try testing.expectEqualStrings("_x0041_", as_formula);

    const as_string = try decodeAt(testing.allocator, .shared_string, raw);
    defer testing.allocator.free(as_string);
    try testing.expectEqualStrings("A", as_string);
}

test "decode: entities resolve in both carriers, and an undeclared one refuses" {
    const f = try decodeAt(testing.allocator, .cell_formula_body, "IF(A2&gt;1,&quot;x&quot;,&apos;y&apos;)");
    defer testing.allocator.free(f);
    try testing.expectEqualStrings("IF(A2>1,\"x\",'y')", f);

    const s = try decodeAt(testing.allocator, .shared_string, "R&amp;D");
    defer testing.allocator.free(s);
    try testing.expectEqualStrings("R&D", s);

    const n = try decodeAt(testing.allocator, .shared_string, "a&#65;&#x42;");
    defer testing.allocator.free(n);
    try testing.expectEqualStrings("aAB", n);

    try testing.expectError(error.BadEntity, decodeEntities(testing.allocator, "x&nbsp;y"));
    try testing.expectError(error.BadEntity, decodeEntities(testing.allocator, "x&y"));
    try testing.expectError(error.BadEntity, decodeEntities(testing.allocator, "&#xD800;"));
}

test "decode: _x005F_ is the escaped underscore, and it protects what follows" {
    // Getting this backwards is how a round-trip stops being one.
    const s = try decodeXstring(testing.allocator, "_x005F_x0041_");
    defer testing.allocator.free(s);
    try testing.expectEqualStrings("_x0041_", s);

    const plain = try decodeXstring(testing.allocator, "_x005F_");
    defer testing.allocator.free(plain);
    try testing.expectEqualStrings("_", plain);
}

test "decode: encoded C0 controls, including the corpus-attested CR" {
    // `tests/corpus/wdi_excel.xlsx` carries `_x000D_` in its shared
    // strings — the one row here the corpus decides rather than the
    // schema.
    const cr = try decodeAt(testing.allocator, .shared_string, "line_x000D_break");
    defer testing.allocator.free(cr);
    try testing.expectEqualStrings("line\rbreak", cr);

    const soh = try decodeAt(testing.allocator, .inline_string, "_x0001_");
    defer testing.allocator.free(soh);
    try testing.expectEqualStrings("\x01", soh);

    // A near-miss is literal text, not an escape.
    for ([_][]const u8{ "_xZZZZ_", "_x12_", "_x0041", "x0041_" }) |near| {
        const got = try decodeXstring(testing.allocator, near);
        defer testing.allocator.free(got);
        try testing.expectEqualStrings(near, got);
    }
}

test "decode: a surrogate escape refuses rather than producing invalid UTF-8" {
    try testing.expectError(error.BadXstring, decodeXstring(testing.allocator, "_xD800_"));
}

test "author: a formula XML-escapes only; a string ST_Xstring-encodes first" {
    // The authored direction of the same asymmetry. A formula
    // containing `_x0041_` must come back out as `_x0041_`, which means
    // the encoder must NOT escape the underscore — while a string with
    // the same bytes must.
    const f = try encodeAuthoredFormula(testing.allocator, "IF(A1>0,\"_x0041_\",\"b\")");
    defer testing.allocator.free(f);
    try testing.expectEqualStrings("IF(A1&gt;0,&quot;_x0041_&quot;,&quot;b&quot;)", f);

    const back = try decodeAt(testing.allocator, .cell_formula_body, f);
    defer testing.allocator.free(back);
    try testing.expectEqualStrings("IF(A1>0,\"_x0041_\",\"b\")", back);

    const s = try encodeAuthoredString(testing.allocator, "_x0041_");
    defer testing.allocator.free(s);
    try testing.expectEqualStrings("_x005F_x0041_", s);

    const s_back = try decodeAt(testing.allocator, .shared_string, s);
    defer testing.allocator.free(s_back);
    try testing.expectEqualStrings("_x0041_", s_back);
}

test "author: a formula that XML cannot carry refuses instead of inventing an escape" {
    try testing.expectError(
        error.UnencodableChar,
        encodeAuthoredFormula(testing.allocator, "A1&\x01"),
    );
    // …while a string has somewhere to put it.
    const s = try encodeAuthoredString(testing.allocator, "a\x01b");
    defer testing.allocator.free(s);
    try testing.expectEqualStrings("a_x0001_b", s);
}

test "author: round-trip over every carrier, including entity and control edges" {
    const cases = [_][]const u8{
        "plain",
        "R&D",
        "a<b>c",
        "_x0041_",
        "_x005F_",
        "__x0041__",
        "quote\"and'apos",
        "tab\tnewline\n",
    };
    for (cases) |c| {
        const encoded = try encodeAuthoredString(testing.allocator, c);
        defer testing.allocator.free(encoded);
        const back = try decodeAt(testing.allocator, .shared_string, encoded);
        defer testing.allocator.free(back);
        try testing.expectEqualStrings(c, back);

        const f = try encodeAuthoredFormula(testing.allocator, c);
        defer testing.allocator.free(f);
        const f_back = try decodeAt(testing.allocator, .cell_formula_body, f);
        defer testing.allocator.free(f_back);
        try testing.expectEqualStrings(c, f_back);
    }
}

// ─── namespace preflight ─────────────────────────────────────────

test "preflight: the main namespace is accepted, prefixed or default" {
    try testing.expectEqual(@as(?Refusal, null), preflight(
        .worksheet,
        "<worksheet xmlns=\"" ++ ns_main ++ "\"><sheetData/></worksheet>",
        .{},
    ));
    try testing.expectEqual(@as(?Refusal, null), preflight(
        .worksheet,
        "<x:worksheet xmlns:x=\"" ++ ns_main ++ "\"><x:sheetData/></x:worksheet>",
        .{},
    ));
}

test "preflight: an unknown namespace refuses before anything is decoded" {
    // The failure this prevents: a literal-matching scanner reads a
    // document in a vocabulary it does not know as an EMPTY sheet, and
    // recalculating "no cells" writes zeroes over real data.
    const bogus = "<worksheet xmlns=\"http://example.invalid/sheet\"><sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c></row></sheetData></worksheet>";
    const r = preflight(.worksheet, bogus, .{}).?;
    try testing.expectEqual(Refusal.Reason.unknown_namespace, r.reason);
    try testing.expectEqual(PlaneTwo.FormulaUnsupportedConstruct, r.planeTwo());

    // And the scan refuses too, without producing a single cell.
    try testing.expectEqual(
        Refusal.Reason.unknown_namespace,
        (try scanRefusal(bogus, &.{})).reason,
    );
}

test "preflight: strict OOXML refuses as classified, not as unknown" {
    const r = preflight(
        .worksheet,
        "<worksheet xmlns=\"" ++ ns_main_strict ++ "\"/>",
        .{},
    ).?;
    try testing.expectEqual(Refusal.Reason.unsupported_namespace, r.reason);
}

test "preflight: no namespace, wrong root, and a DOCTYPE each refuse" {
    try testing.expectEqual(
        Refusal.Reason.missing_namespace,
        preflight(.worksheet, "<worksheet><sheetData/></worksheet>", .{}).?.reason,
    );
    try testing.expectEqual(
        Refusal.Reason.wrong_root_element,
        preflight(.worksheet, "<workbook xmlns=\"" ++ ns_main ++ "\"/>", .{}).?.reason,
    );
    try testing.expectEqual(
        Refusal.Reason.doctype_declaration,
        preflight(.worksheet, "<!DOCTYPE w><worksheet xmlns=\"" ++ ns_main ++ "\"/>", .{}).?.reason,
    );
    try testing.expectEqual(
        Refusal.Reason.invalid_utf8,
        preflight(.worksheet, "<worksheet xmlns=\"\xff\xfe\"/>", .{}).?.reason,
    );
}

test "preflight: each part kind demands its own root element" {
    try testing.expectEqual(@as(?Refusal, null), preflight(
        .shared_strings,
        "<sst xmlns=\"" ++ ns_main ++ "\"/>",
        .{},
    ));
    try testing.expectEqual(@as(?Refusal, null), preflight(
        .table,
        "<table xmlns=\"" ++ ns_main ++ "\"/>",
        .{},
    ));
    try testing.expectEqual(
        Refusal.Reason.wrong_root_element,
        preflight(.table, "<sst xmlns=\"" ++ ns_main ++ "\"/>", .{}).?.reason,
    );
}

// ─── shared strings ──────────────────────────────────────────────

test "sst: plain, rich, and phonetic entries" {
    const xml = "<sst" ++ ns_attr ++ ">" ++
        "<si><t>plain</t></si>" ++
        "<si><r><rPr><b/></rPr><t>bo</t></r><r><t>ld</t></r></si>" ++
        "<si><t>kanji</t><rPh sb=\"0\" eb=\"2\"><t>furigana</t></rPh></si>" ++
        "<si/>" ++
        "<si><t>_x000D_</t></si>" ++
        "</sst>";
    var s = switch (try decodeSharedStrings(testing.allocator, xml, .{})) {
        .ok => |ok| ok,
        .refused => return error.TestUnexpectedRefusal,
    };
    defer s.deinit();

    try testing.expectEqual(@as(usize, 5), s.items.len);
    try testing.expectEqualStrings("plain", s.items[0]);
    // Rich runs concatenate in document order.
    try testing.expectEqualStrings("bold", s.items[1]);
    // …and the phonetic run is not part of the value.
    try testing.expectEqualStrings("kanji", s.items[2]);
    try testing.expectEqualStrings("", s.items[3]);
    try testing.expectEqualStrings("\r", s.items[4]);
}

test "sst: a string carrier applies both passes, in order" {
    const xml = "<sst" ++ ns_attr ++ "><si><t>R&amp;D _x0041_ &amp;#65;</t></si></sst>";
    var s = switch (try decodeSharedStrings(testing.allocator, xml, .{})) {
        .ok => |ok| ok,
        .refused => return error.TestUnexpectedRefusal,
    };
    defer s.deinit();
    // Entities first (`&amp;` → `&`), then ST_Xstring (`_x0041_` → `A`).
    // The `&amp;#65;` proves the order: entity decoding does not rerun,
    // so it stays the literal text `&#65;`.
    try testing.expectEqualStrings("R&D A &#65;", s.items[0]);
}

// ─── implicit coordinates ────────────────────────────────────────

test "implicit coords: the first r-less cell is column A, and the rest step" {
    var c: CellCursor = .{};
    try testing.expectEqual(@as(u32, 0), c.resolve(null, 1).ok.zeroBased());
    try testing.expectEqual(@as(u32, 1), c.resolve(null, 1).ok.zeroBased());
    try testing.expectEqual(@as(u32, 2), c.resolve(null, 1).ok.zeroBased());
    // An explicit `r` also sets the cursor.
    try testing.expectEqual(@as(u32, 9), c.resolve("J1", 1).ok.zeroBased());
    try testing.expectEqual(@as(u32, 10), c.resolve(null, 1).ok.zeroBased());
    // …and a new row starts over.
    c.startRow();
    try testing.expectEqual(@as(u32, 0), c.resolve(null, 2).ok.zeroBased());
}

test "implicit coords: stepping past the last column refuses, never wraps" {
    var c: CellCursor = .{};
    _ = c.resolve("XFD1", 1);
    try testing.expectEqual(
        Refusal.Reason.implicit_ref_out_of_grid,
        c.resolve(null, 1).refused,
    );
}

test "implicit coords: a ref naming another row is malformed" {
    var c: CellCursor = .{};
    try testing.expectEqual(
        Refusal.Reason.malformed_cell_ref,
        c.resolve("A2", 1).refused,
    );
    try testing.expectEqual(
        Refusal.Reason.malformed_cell_ref,
        c.resolve("not-a-ref", 1).refused,
    );
}

test "scan: r-less cells reconstruct, with gaps, formulas, and the grid edge" {
    // First-cell-no-r, then a gap made by an explicit ref, then an
    // r-less formula cell after it.
    var sheet = try scanOk(sheetXml(
        \\<row r="1"><c><v>1</v></c><c><v>2</v></c><c r="E1"><v>5</v></c><c><f>E1+1</f><v>6</v></c></row>
    ), &.{});
    defer sheet.deinit();

    try testing.expectEqual(@as(usize, 4), sheet.cells.len);
    try testing.expectEqual(@as(u32, 0), sheet.cells[0].col.zeroBased());
    try testing.expectEqual(@as(u32, 1), sheet.cells[1].col.zeroBased());
    try testing.expectEqual(@as(u32, 4), sheet.cells[2].col.zeroBased());
    try testing.expectEqual(@as(u32, 5), sheet.cells[3].col.zeroBased());
    try testing.expectEqualStrings("E1+1", sheet.cells[3].formula.?.text);

    // Out of grid: a cell after XFD has nowhere to go.
    const r = try scanRefusal(sheetXml(
        \\<row r="1"><c r="XFD1"><v>1</v></c><c><v>2</v></c></row>
    ), &.{});
    try testing.expectEqual(Refusal.Reason.implicit_ref_out_of_grid, r.reason);
}

// ─── the input cell-type contract ────────────────────────────────

test "contract: uncached wins over t, always" {
    // A formula cell with no `<v>` is uncached whatever `t` claims —
    // the precedence §5.7.2 calls normative. `t="b"` with no cached
    // value is a legitimate state for a formula cell and a refusal for
    // a stored one.
    var sheet = try scanOk(sheetXml(
        \\<row r="1"><c r="A1" t="b"><f>A2=1</f></c><c r="B1" t="e"><f>NA()</f></c></row>
    ), &.{});
    defer sheet.deinit();
    try testing.expectEqual(@as(usize, 2), sheet.cells.len);
    try testing.expect(sheet.cells[0].input == .uncached);
    try testing.expect(sheet.cells[1].input == .uncached);

    const r = try scanRefusal(sheetXml(
        \\<row r="1"><c r="A1" t="b"/></row>
    ), &.{});
    try testing.expectEqual(Refusal.Reason.missing_cached_value, r.reason);
}

test "contract: every t maps, and an unknown one refuses instead of becoming a number" {
    var sheet = try scanOk(sheetXml(
        \\<row r="1">
        \\<c r="A1"><v>1.5</v></c>
        \\<c r="B1" t="n"><v>2</v></c>
        \\<c r="C1" t="s"><v>0</v></c>
        \\<c r="D1" t="str"><f>X()</f><v>text</v></c>
        \\<c r="E1" t="inlineStr"><is><t>inline</t></is></c>
        \\<c r="F1" t="b"><v>1</v></c>
        \\<c r="G1" t="e"><v>#REF!</v></c>
        \\<c r="H1" s="3"/>
        \\</row>
    ), &.{"shared"});
    defer sheet.deinit();

    try testing.expectEqual(@as(f64, 1.5), sheet.cells[0].input.number);
    try testing.expectEqual(@as(f64, 2), sheet.cells[1].input.number);
    try testing.expectEqualStrings("shared", sheet.cells[2].input.text);
    try testing.expectEqualStrings("text", sheet.cells[3].input.text);
    try testing.expectEqualStrings("inline", sheet.cells[4].input.text);
    try testing.expectEqual(true, sheet.cells[5].input.boolean);
    try testing.expectEqual(value.KnownError.ref, sheet.cells[6].input.err.known);
    // `H1` is a styled empty cell: blank is the absence of a cell, so it
    // is not in the view at all.
    try testing.expectEqual(@as(usize, 7), sheet.cells.len);

    const unknown = try scanRefusal(sheetXml(
        \\<row r="1"><c r="A1" t="q"><v>1</v></c></row>
    ), &.{});
    try testing.expectEqual(Refusal.Reason.unknown_cell_type, unknown.reason);
    try testing.expectEqual(PlaneTwo.FormulaMalformedInput, unknown.planeTwo());
    // Case matters: `t="S"` is not `ST_CellType`'s `s`.
    const cased = try scanRefusal(sheetXml(
        \\<row r="1"><c r="A1" t="S"><v>0</v></c></row>
    ), &.{"x"});
    try testing.expectEqual(Refusal.Reason.unknown_cell_type, cased.reason);
}

test "contract: malformed caches refuse, and never seed a zero" {
    const cases = [_]struct { xml: []const u8, reason: Refusal.Reason }{
        .{ .xml = "<row r=\"1\"><c r=\"A1\"><v>abc</v></c></row>", .reason = .bad_number_cache },
        .{ .xml = "<row r=\"1\"><c r=\"A1\"><v></v></c></row>", .reason = .bad_number_cache },
        .{ .xml = "<row r=\"1\"><c r=\"A1\"><v>1e400</v></c></row>", .reason = .bad_number_cache },
        .{ .xml = "<row r=\"1\"><c r=\"A1\" t=\"b\"><v>2</v></c></row>", .reason = .bad_boolean_cache },
        .{ .xml = "<row r=\"1\"><c r=\"A1\" t=\"b\"><v> 1</v></c></row>", .reason = .bad_boolean_cache },
        .{ .xml = "<row r=\"1\"><c r=\"A1\" t=\"b\"><v></v></c></row>", .reason = .bad_boolean_cache },
        .{ .xml = "<row r=\"1\"><c r=\"A1\" t=\"e\"><v>#NOPE</v></c></row>", .reason = .bad_error_cache },
        .{ .xml = "<row r=\"1\"><c r=\"A1\" t=\"e\"><v></v></c></row>", .reason = .bad_error_cache },
        .{ .xml = "<row r=\"1\"><c r=\"A1\" t=\"e\"/></row>", .reason = .missing_cached_value },
        .{ .xml = "<row r=\"1\"><c r=\"A1\" t=\"s\"><v>7</v></c></row>", .reason = .shared_string_index_out_of_range },
        .{ .xml = "<row r=\"1\"><c r=\"A1\" t=\"d\"><v>2026-08-03T12:00:00Z</v></c></row>", .reason = .bad_date_cache },
    };
    for (cases) |c| {
        var buf: [512]u8 = undefined;
        const xml = try std.fmt.bufPrint(&buf, "<worksheet{s}><sheetData>{s}</sheetData></worksheet>", .{ ns_attr, c.xml });
        const r = try scanRefusal(xml, &.{});
        testing.expectEqual(c.reason, r.reason) catch |err| {
            std.debug.print("case: {s}\n", .{c.xml});
            return err;
        };
    }
}

test "contract: a rich error spelling is preserved byte-exact" {
    var sheet = try scanOk(sheetXml(
        \\<row r="1"><c r="A1" t="e"><v>#BLOCKED!</v></c></row>
    ), &.{});
    defer sheet.deinit();
    try testing.expectEqualStrings("#BLOCKED!", sheet.cells[0].input.err.rich);
}

test "contract: a retained cached value survives the cells that follow it" {
    // A one-cell fixture cannot see this. The walk decodes every `<v>`
    // into one scratch buffer and clears it per cell, so a value the
    // classification *keeps* — a rich error spelling, a `t="str"` body —
    // has to leave that buffer before the next `<c>` starts. It did not,
    // and a rich error read back as whatever the following cell cached
    // (found by M5b1's round-trip fuzz target).
    var sheet = try scanOk(sheetXml(
        \\<row r="1"><c r="A1" t="e"><v>#POWER_QUERY!</v></c>
    ++
        \\<c r="B1" t="str"><v>bbbbbbbbbbbbbbbbbb</v></c>
    ++
        \\<c r="C1" t="e"><v>#DIV/0!</v></c>
    ++
        \\<c r="D1"><v>12345678901234567890</v></c></row>
    ), &.{});
    defer sheet.deinit();

    try testing.expectEqualStrings("#POWER_QUERY!", sheet.cells[0].input.err.rich);
    try testing.expectEqualStrings("bbbbbbbbbbbbbbbbbb", sheet.cells[1].input.text);
    try testing.expectEqual(value.KnownError.div0, sheet.cells[2].input.err.known);
}

test "scan: a prefixed document reads exactly like a default-namespaced one" {
    const prefixed =
        "<x:worksheet xmlns:x=\"" ++ ns_main ++ "\"><x:sheetData>" ++
        "<x:row r=\"1\"><x:c r=\"A1\" t=\"s\"><x:v>0</x:v></x:c>" ++
        "<x:c r=\"B1\"><x:f>A1&amp;\"x\"</x:f><x:v>3</x:v></x:c></x:row>" ++
        "</x:sheetData></x:worksheet>";
    var sheet = try scanOk(prefixed, &.{"R&D"});
    defer sheet.deinit();
    try testing.expectEqual(@as(usize, 2), sheet.cells.len);
    try testing.expectEqualStrings("R&D", sheet.cells[0].input.text);
    try testing.expectEqualStrings("A1&\"x\"", sheet.cells[1].formula.?.text);
}

test "scan: the formula body is a formula carrier and its attributes survive for M4b2" {
    var sheet = try scanOk(sheetXml(
        \\<row r="1"><c r="A1"><f t="shared" si="0" ref="A1:A3" ca="1">SUM(B1:B3)&amp;"_x0041_"</f><v>1</v></c></row>
    ), &.{});
    defer sheet.deinit();
    const f = sheet.cells[0].formula.?;
    try testing.expectEqualStrings("SUM(B1:B3)&\"_x0041_\"", f.text);
    try testing.expectEqualStrings("shared", f.kind.?);
    try testing.expectEqualStrings("0", f.si.?);
    try testing.expectEqualStrings("A1:A3", f.ref.?);
    // The raw region is kept whole — M4b2 owns the inventory, and it
    // cannot inventory bytes M4b1 threw away.
    try testing.expect(std.mem.indexOf(u8, f.raw_attrs, "ca=\"1\"") != null);
}

test "scan: an inline string concatenates runs and drops phonetics" {
    var sheet = try scanOk(sheetXml(
        \\<row r="1"><c r="A1" t="inlineStr"><is><r><t>bo</t></r><r><t>ld</t></r><rPh sb="0" eb="2"><t>skip</t></rPh></is></c></row>
    ), &.{});
    defer sheet.deinit();
    try testing.expectEqualStrings("bold", sheet.cells[0].input.text);
}

test "scan: the three string carriers decode the same bytes the same way" {
    // One escape, three carriers, one answer — while the formula in the
    // same cell keeps its literal.
    var sheet = try scanOk(sheetXml(
        \\<row r="1">
        \\<c r="A1" t="s"><v>0</v></c>
        \\<c r="B1" t="inlineStr"><is><t>_x0041_&amp;_x005F_x0042_</t></is></c>
        \\<c r="C1" t="str"><f>"_x0041_"</f><v>_x0041_&amp;_x005F_x0042_</v></c>
        \\</row>
    ), &.{"A&_x0042_"});
    defer sheet.deinit();
    try testing.expectEqualStrings("A&_x0042_", sheet.cells[0].input.text);
    try testing.expectEqualStrings("A&_x0042_", sheet.cells[1].input.text);
    try testing.expectEqualStrings("A&_x0042_", sheet.cells[2].input.text);
    try testing.expectEqualStrings("\"_x0041_\"", sheet.cells[2].formula.?.text);
}

test "scan: structural refusals name what they refused" {
    const cases = [_]struct { xml: []const u8, reason: Refusal.Reason }{
        .{ .xml = "<row r=\"1\"><c r=\"A1\" zz=\"1\"><v>1</v></c></row>", .reason = .unexpected_attribute },
        .{ .xml = "<row r=\"1\"><c r=\"A1\"><nope/></c></row>", .reason = .unexpected_element },
        .{ .xml = "<row r=\"1\"><nope/></row>", .reason = .unexpected_element },
        .{ .xml = "<nope/>", .reason = .unexpected_element },
        .{ .xml = "<row r=\"1\"><c r=\"A1\"><v>1</v></c></row>text", .reason = .unexpected_text },
        .{ .xml = "<row r=\"1\"><c r=\"A1\"><v>1</v></c><c r=\"A1\"><v>2</v></c></row>", .reason = .duplicate_cell },
        .{ .xml = "<row r=\"0\"><c r=\"A1\"/></row>", .reason = .bad_attribute_value },
        .{ .xml = "<row r=\"1\"><c r=\"A1\"><v><![CDATA[1]]></v></c></row>", .reason = .unexpected_cdata },
    };
    for (cases) |c| {
        var buf: [512]u8 = undefined;
        const xml = try std.fmt.bufPrint(&buf, "<worksheet{s}><sheetData>{s}</sheetData></worksheet>", .{ ns_attr, c.xml });
        const r = try scanRefusal(xml, &.{});
        testing.expectEqual(c.reason, r.reason) catch |err| {
            std.debug.print("case: {s}\n", .{c.xml});
            return err;
        };
    }
}

test "scan: a foreign element inside sheetData refuses rather than dropping cells" {
    const xml = "<worksheet" ++ ns_attr ++ " xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\">" ++
        "<sheetData><mc:AlternateContent><row r=\"1\"><c r=\"A1\"><v>1</v></c></row></mc:AlternateContent></sheetData></worksheet>";
    const r = try scanRefusal(xml, &.{});
    try testing.expectEqual(Refusal.Reason.foreign_element, r.reason);
}

test "scan: namespaced extension attributes and out-of-sheetData parts are tolerated" {
    // `x14ac:dyDescent` is on nearly every row Excel writes; refusing a
    // namespaced extension would refuse the ordinary case. Everything
    // outside `<sheetData>` is skipped wholesale, foreign or not.
    const xml =
        "<worksheet" ++ ns_attr ++ " xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\">" ++
        "<sheetPr><tabColor rgb=\"FF00FF\"/></sheetPr>" ++
        "<dimension ref=\"A1:A1\"/>" ++
        "<sheetData><row r=\"1\" x14ac:dyDescent=\"0.25\"><c r=\"A1\"><v>1</v></c></row></sheetData>" ++
        "<extLst><ext uri=\"{x}\"><foo:bar xmlns:foo=\"http://example.invalid\"><deep/></foo:bar></ext></extLst>" ++
        "</worksheet>";
    var sheet = try scanOk(xml, &.{});
    defer sheet.deinit();
    try testing.expectEqual(@as(usize, 1), sheet.cells.len);
    try testing.expectEqual(@as(f64, 1), sheet.cells[0].input.number);
}

test "scan: mergeCells is the one interpreted element outside sheetData (M7a)" {
    const xml =
        "<worksheet" ++ ns_attr ++ ">" ++
        "<sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c></row></sheetData>" ++
        "<mergeCells count=\"99\">" ++ // miscounted on purpose: recorded nowhere, enforced nowhere
        "<mergeCell ref=\"B2:C3\"/>" ++
        "<mergeCell ref=\"$E$1:$F$2\"/>" ++ // `$` is unambiguous; refusing it would refuse a file Excel opens
        "<mergeCell ref=\"D9\"/>" ++ // a degenerate single-cell spelling
        "</mergeCells>" ++
        "</worksheet>";
    var sheet = try scanOk(xml, &.{});
    defer sheet.deinit();
    try testing.expectEqual(@as(usize, 3), sheet.merges.len);
    try testing.expectEqual(@as(u32, 2), sheet.merges[0].first.row.oneBased());
    try testing.expectEqual(@as(u32, 2), sheet.merges[0].last.col.zeroBased());
    try testing.expectEqual(@as(u32, 4), sheet.merges[1].first.col.zeroBased());
    try testing.expectEqual(@as(u32, 9), sheet.merges[2].first.row.oneBased());
}

test "scan: a malformed mergeCell refuses like any bad attribute" {
    const bad_ref =
        "<worksheet" ++ ns_attr ++ ">" ++
        "<sheetData/>" ++
        "<mergeCells><mergeCell ref=\"NOT A RANGE\"/></mergeCells>" ++
        "</worksheet>";
    const r = try scanRefusal(bad_ref, &.{});
    try testing.expectEqual(Refusal.Reason.bad_attribute_value, r.reason);

    const no_ref =
        "<worksheet" ++ ns_attr ++ ">" ++
        "<sheetData/>" ++
        "<mergeCells><mergeCell/></mergeCells>" ++
        "</worksheet>";
    const r2 = try scanRefusal(no_ref, &.{});
    try testing.expectEqual(Refusal.Reason.bad_attribute_value, r2.reason);

    // An unknown attribute on an interpreted element refuses (M4a
    // decision 12), and so does a main-namespace element the schema
    // does not put inside `<mergeCells>`.
    const bad_attr =
        "<worksheet" ++ ns_attr ++ ">" ++
        "<sheetData/>" ++
        "<mergeCells><mergeCell ref=\"A1:B2\" bogus=\"1\"/></mergeCells>" ++
        "</worksheet>";
    const r3 = try scanRefusal(bad_attr, &.{});
    try testing.expectEqual(Refusal.Reason.unexpected_attribute, r3.reason);

    const bad_child =
        "<worksheet" ++ ns_attr ++ ">" ++
        "<sheetData/>" ++
        "<mergeCells><row r=\"1\"/></mergeCells>" ++
        "</worksheet>";
    const r4 = try scanRefusal(bad_child, &.{});
    try testing.expectEqual(Refusal.Reason.unexpected_element, r4.reason);
}

test "scan: a sheet with no mergeCells answers an empty list" {
    var sheet = try scanOk(sheetXml(
        \\<row r="1"><c r="A1"><v>1</v></c></row>
    ), &.{});
    defer sheet.deinit();
    try testing.expectEqual(@as(usize, 0), sheet.merges.len);
}

test "scan: row geometry is recorded where the cells are (M7b1)" {
    const xml =
        "<worksheet" ++ ns_attr ++ "><sheetData>" ++
        "<row r=\"1\" spans=\"1:2\"><c r=\"A1\"><v>1</v></c><c r=\"B1\"><v>2</v></c></row>" ++
        "<row r=\"3\"/>" ++
        "<row><c r=\"A9\"><v>9</v></c></row>" ++
        "</sheetData></worksheet>";
    var sheet = try scanOk(xml, &.{});
    defer sheet.deinit();

    try testing.expectEqual(@as(usize, 3), sheet.rows.len);

    // Row 1: number, spans value, and an open/content pair that brackets
    // exactly its two cells.
    const r1 = sheet.rows[0];
    try testing.expectEqual(@as(u32, 1), r1.number);
    try testing.expectEqualStrings("1:2", r1.spans_attr.?.slice(xml));
    try testing.expect(!r1.selfClosing());
    try testing.expectEqualStrings(
        "<c r=\"A1\"><v>1</v></c><c r=\"B1\"><v>2</v></c>",
        xml[r1.open_end..r1.content_end],
    );
    try testing.expectEqual(@as(u8, '<'), xml[r1.element.start]);
    try testing.expectEqual(@as(u8, '>'), xml[r1.element.end - 1]);

    // Row 3: self-closing — no content region at all.
    const r3 = sheet.rows[1];
    try testing.expectEqual(@as(u32, 3), r3.number);
    try testing.expect(r3.selfClosing());
    try testing.expectEqual(r3.open_end, r3.content_end);
    try testing.expect(r3.spans_attr == null);

    // The bare row took its number from its first cell — the same
    // authority the model used, so the geometry cannot disagree with
    // the slots about where row 9 is.
    try testing.expectEqual(@as(u32, 9), sheet.rows[2].number);
}

test "scan: the dimension is recorded, not interpreted (M7b1)" {
    const xml =
        "<worksheet" ++ ns_attr ++ ">" ++
        "<dimension ref=\"A1:B2\"/>" ++
        "<sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c></row></sheetData>" ++
        "</worksheet>";
    var sheet = try scanOk(xml, &.{});
    defer sheet.deinit();
    const dim = sheet.dimension.?;
    try testing.expectEqualStrings("A1:B2", dim.ref.?.slice(xml));
    try testing.expectEqualStrings("<dimension ref=\"A1:B2\"/>", dim.element.slice(xml));

    // Absent: nothing to maintain, and nothing invented.
    var bare = try scanOk(sheetXml(
        \\<row r="1"><c r="A1"><v>1</v></c></row>
    ), &.{});
    defer bare.deinit();
    try testing.expect(bare.dimension == null);

    // A ref this grid cannot address is still RECORDED — dimension
    // staleness is tolerated everywhere except under a spill that
    // needs to expand it, so interpretation is the patcher's, at the
    // one moment it matters.
    const odd =
        "<worksheet" ++ ns_attr ++ ">" ++
        "<dimension ref=\"NOT A RANGE\"/>" ++
        "<sheetData/>" ++
        "</worksheet>";
    var tolerated = try scanOk(odd, &.{});
    defer tolerated.deinit();
    try testing.expectEqualStrings("NOT A RANGE", tolerated.dimension.?.ref.?.slice(odd));
}

test "scan: two dimensions refuse, and an unknown attribute on one refuses (M7b1)" {
    const twice =
        "<worksheet" ++ ns_attr ++ ">" ++
        "<dimension ref=\"A1\"/><dimension ref=\"A1:B2\"/>" ++
        "<sheetData/>" ++
        "</worksheet>";
    const r = try scanRefusal(twice, &.{});
    try testing.expectEqual(Refusal.Reason.unexpected_element, r.reason);

    const bad_attr =
        "<worksheet" ++ ns_attr ++ ">" ++
        "<dimension ref=\"A1\" bogus=\"1\"/>" ++
        "<sheetData/>" ++
        "</worksheet>";
    const r2 = try scanRefusal(bad_attr, &.{});
    try testing.expectEqual(Refusal.Reason.unexpected_attribute, r2.reason);
}

test "scan: the anchor ref value has its own span, inside `spans.f` (M7b1)" {
    const xml = sheetXml(
        \\<row r="1"><c r="A1"><f t="array" ref="A1:B3" aca="1">SEQUENCE(3,2)</f><v>1</v></c><c r="B1"><f>1+1</f><v>2</v></c></row>
    );
    var sheet = try scanOk(xml, &.{});
    defer sheet.deinit();

    const anchor = sheet.slots[0].spans;
    const ref = anchor.f_ref.?;
    try testing.expectEqualStrings("A1:B3", ref.slice(xml));
    // Inside the `<f>` element — the one byte range there §5.8b's
    // anchor-ref mutation may address.
    const f = anchor.f.?;
    try testing.expect(ref.start > f.start and ref.end < f.end);

    // A refless formula has no span to address.
    try testing.expect(sheet.slots[1].spans.f_ref == null);
}

test "scan: cells come back row-major regardless of document order" {
    var sheet = try scanOk(sheetXml(
        \\<row r="2"><c r="C2"><v>3</v></c><c r="A2"><v>1</v></c></row>
        \\<row r="1"><c r="B1"><v>2</v></c></row>
    ), &.{});
    defer sheet.deinit();
    try testing.expectEqual(@as(u32, 1), sheet.cells[0].row.oneBased());
    try testing.expectEqual(@as(u32, 2), sheet.cells[1].row.oneBased());
    try testing.expectEqual(@as(u32, 0), sheet.cells[1].col.zeroBased());
    try testing.expectEqual(@as(u32, 2), sheet.cells[2].col.zeroBased());
}

test "scan: cm and vm survive for M4a to resolve" {
    var sheet = try scanOk(sheetXml(
        \\<row r="1"><c r="A1" cm="1" vm="2"><v>1</v></c><c r="B1"><v>2</v></c></row>
    ), &.{});
    defer sheet.deinit();
    try testing.expectEqual(@as(u32, 1), sheet.cells[0].cm);
    try testing.expectEqual(@as(u32, 2), sheet.cells[0].vm);
    try testing.expectEqual(@as(u32, 0), sheet.cells[1].cm);
}

test "scan: limits refuse with the limit named" {
    const res = try scanSheet(testing.allocator, sheetXml(
        \\<row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row>
    ), &.{}, .{ .limits = .{ .max_modeled_cells = 1 } });
    switch (res) {
        .ok => |s| {
            var sheet = s;
            sheet.deinit();
            return error.TestExpectedRefusal;
        },
        .refused => |r| {
            try testing.expectEqual(LimitKind.modeled_cells, r.limit.?);
            try testing.expectEqual(PlaneTwo.FormulaLimitExceeded, r.planeTwo());
        },
    }
}

// ─── tables ──────────────────────────────────────────────────────

test "table: names are string carriers and column formulas are formula carriers" {
    const xml = "<table" ++ ns_attr ++ " id=\"1\" name=\"Sales_x0041_\" displayName=\"Sales_x0041_\" ref=\"A1:C4\">" ++
        "<tableColumns count=\"2\"><tableColumn id=\"1\" name=\"R&amp;D\"/>" ++
        "<tableColumn id=\"2\" name=\"Total\"><calculatedColumnFormula>SUM(Sales[@[R&amp;D]])&amp;\"_x0041_\"</calculatedColumnFormula></tableColumn>" ++
        "</tableColumns></table>";
    var t = switch (try scanTable(testing.allocator, xml, .{})) {
        .ok => |ok| ok,
        .refused => return error.TestUnexpectedRefusal,
    };
    defer t.deinit();

    // Identifiers are ST_Xstring-typed, so the escape decodes…
    try testing.expectEqualStrings("SalesA", t.name);
    try testing.expectEqualStrings("R&D", t.columns[0].name);
    // …while the calculated-column formula keeps its literal.
    try testing.expectEqualStrings("SUM(Sales[@[R&D]])&\"_x0041_\"", t.columns[1].calculated_formula.?);
    try testing.expectEqualStrings("A1:C4", t.ref);
}

// ─── allocation failure ──────────────────────────────────────────

test "checkAllAllocationFailures: decoding leaks nothing under OOM" {
    const H = struct {
        fn run(allocator: std.mem.Allocator) !void {
            const sst = "<sst" ++ ns_attr ++ "><si><t>a_x0041_</t></si><si><r><t>b</t></r></si></sst>";
            var s = switch (try decodeSharedStrings(allocator, sst, .{})) {
                .ok => |ok| ok,
                .refused => return error.TestUnexpectedRefusal,
            };
            defer s.deinit();

            var sheet = switch (try scanSheet(allocator, sheetXml(
                \\<row r="1"><c r="A1" t="s"><v>0</v></c><c><f>A1&amp;"x"</f><v>2</v></c>
                \\<c r="D1" t="inlineStr"><is><r><t>x</t></r></is></c></row>
            ), s.items, .{})) {
                .ok => |ok| ok,
                .refused => return error.TestUnexpectedRefusal,
            };
            defer sheet.deinit();
            if (sheet.cells.len != 3) return error.WrongCount;

            const tbl = "<table" ++ ns_attr ++ " name=\"T\" ref=\"A1:B2\"><tableColumns>" ++
                "<tableColumn name=\"c\"><calculatedColumnFormula>A1</calculatedColumnFormula></tableColumn>" ++
                "</tableColumns></table>";
            var t = switch (try scanTable(allocator, tbl, .{})) {
                .ok => |ok| ok,
                .refused => return error.TestUnexpectedRefusal,
            };
            defer t.deinit();

            const authored = try encodeAuthoredString(allocator, "_x0041_");
            defer allocator.free(authored);
            const back = try decodeAt(allocator, .shared_string, authored);
            defer allocator.free(back);
        }
    };
    try testing.checkAllAllocationFailures(testing.allocator, H.run, .{});
}

// ─── `t="d"` (§5.7.2's normative lexical table, M4b2) ────────────

fn scanWithEpoch(xml: []const u8, system: serial_date.DateSystem) !SheetResult {
    return scanSheet(testing.allocator, xml, &.{}, .{ .date_system = system });
}

test "t=\"d\": the two accepted forms become serials under the active epoch" {
    var sheet = switch (try scanWithEpoch(sheetXml(
        \\<row r="1"><c r="A1" t="d"><v>2026-08-03</v></c><c r="B1" t="d"><v>2026-08-03T12:00:00</v></c><c r="C1" t="d"><v>1900-02-29</v></c></row>
    ), .d1900)) {
        .ok => |s| s,
        .refused => return error.TestUnexpectedRefusal,
    };
    defer sheet.deinit();

    const day: f64 = @floatFromInt(try serial_date.serialFromDate(.d1900, 2026, 8, 3));
    try testing.expectEqual(day, sheet.cells[0].input.number);
    try testing.expectEqual(day + 0.5, sheet.cells[1].input.number);
    // The invented day is a serial like any other, and only in 1900.
    try testing.expectEqual(@as(f64, 60), sheet.cells[2].input.number);
}

test "t=\"d\": `date1904` reads the same bytes as a different serial" {
    // Which is exactly why the epoch is workbook-derived and not a
    // caller option at the public layer (§5.5).
    for ([_]serial_date.DateSystem{ .d1900, .d1904 }) |system| {
        var sheet = switch (try scanWithEpoch(sheetXml(
            \\<row r="1"><c r="A1" t="d"><v>2026-08-03</v></c></row>
        ), system)) {
            .ok => |s| s,
            .refused => return error.TestUnexpectedRefusal,
        };
        defer sheet.deinit();
        const want: f64 = @floatFromInt(try serial_date.serialFromDate(system, 2026, 8, 3));
        try testing.expectEqual(want, sheet.cells[0].input.number);
    }
}

test "t=\"d\": a timezone offset, a bad form, and a bad date all refuse" {
    for ([_][]const u8{
        "2026-08-03T12:00:00Z",
        "2026-08-03T12:00:00+01:00",
        "2026-08-03 12:00:00",
        "03/08/2026",
        "2026-02-30",
        "",
    }) |v| {
        var doc: std.ArrayListUnmanaged(u8) = .empty;
        defer doc.deinit(testing.allocator);
        try doc.appendSlice(testing.allocator, "<worksheet" ++ ns_attr ++
            "><sheetData><row r=\"1\"><c r=\"A1\" t=\"d\"><v>");
        try doc.appendSlice(testing.allocator, v);
        try doc.appendSlice(testing.allocator, "</v></c></row></sheetData></worksheet>");
        switch (try scanWithEpoch(doc.items, .d1900)) {
            .ok => |s| {
                var sheet = s;
                sheet.deinit();
                std.debug.print("accepted `{s}`\n", .{v});
                return error.TestExpectedRefusal;
            },
            .refused => |r| {
                try testing.expectEqual(Refusal.Reason.bad_date_cache, r.reason);
                try testing.expectEqual(PlaneTwo.FormulaMalformedInput, r.planeTwo());
            },
        }
    }
    // A date the 1904 epoch cannot express refuses under it and is a
    // serial under the other.
    switch (try scanWithEpoch(sheetXml(
        \\<row r="1"><c r="A1" t="d"><v>1900-02-29</v></c></row>
    ), .d1904)) {
        .ok => |s| {
            var sheet = s;
            sheet.deinit();
            return error.TestExpectedRefusal;
        },
        .refused => |r| try testing.expectEqual(Refusal.Reason.bad_date_cache, r.reason),
    }
}

test "t=\"d\": the uncached rule still wins, and a stored cell still needs a value" {
    // §5.7.2's normative precedence is unchanged by the lexical table:
    // a formula cell with no `<v>` is uncached whatever `t` says.
    var sheet = switch (try scanWithEpoch(sheetXml(
        \\<row r="1"><c r="A1" t="d"><f>TODAY()</f></c></row>
    ), .d1900)) {
        .ok => |s| s,
        .refused => return error.TestUnexpectedRefusal,
    };
    defer sheet.deinit();
    try testing.expectEqual(InputCell.uncached, sheet.cells[0].input);

    // A *stored* `t="d"` with no `<v>` has nothing to interpret, the
    // same way `t="b"` and `t="e"` do not.
    switch (try scanWithEpoch(sheetXml(
        \\<row r="1"><c r="A1" t="d"/></row>
    ), .d1900)) {
        .ok => |s| {
            var sh = s;
            sh.deinit();
            return error.TestExpectedRefusal;
        },
        .refused => |r| try testing.expectEqual(Refusal.Reason.missing_cached_value, r.reason),
    }
}

// ─── fuzz (§8.1: the decode boundary) ────────────────────────────

fn fuzzDecodeTarget(_: void, smith: *std.testing.Smith) anyerror!void {
    @disableInstrumentation();
    var buf: [1024]u8 = undefined;
    const input = buf[0..smith.slice(&buf)];
    const a = std.testing.allocator;

    // 1. Every primitive either answers or refuses, twice the same way.
    for ([_]Carrier{ .formula, .string, .lexical }) |carrier| {
        const first = decodeCarrier(a, carrier, input) catch |e1| {
            const second = decodeCarrier(a, carrier, input);
            try std.testing.expectError(e1, second);
            continue;
        };
        defer a.free(first);
        const second = try decodeCarrier(a, carrier, input);
        defer a.free(second);
        // Determinism: the same bytes decode to the same bytes.
        try std.testing.expectEqualSlices(u8, first, second);
        // A formula carrier never rewrites an ST_Xstring escape.
        if (carrier == .formula and std.mem.indexOfScalar(u8, input, '&') == null) {
            try std.testing.expectEqualSlices(u8, input, first);
        }
    }

    // 2. Authoring round-trips, which is what makes the encoding an
    //    encoding rather than a lossy rendering.
    if (std.unicode.utf8ValidateSlice(input)) {
        const encoded = try encodeAuthoredString(a, input);
        defer a.free(encoded);
        const back = try decodeCarrier(a, .string, encoded);
        defer a.free(back);
        try std.testing.expectEqualSlices(u8, input, back);
    }

    // 3. No part can panic or leak the scan arena, whatever it contains.
    for ([_][]const u8{ "<worksheet", "<sst", "<table" }) |root| {
        var doc: std.ArrayListUnmanaged(u8) = .empty;
        defer doc.deinit(a);
        try doc.appendSlice(a, root);
        try doc.appendSlice(a, ns_attr);
        try doc.append(a, '>');
        try doc.appendSlice(a, input);
        var sheet_res = try scanSheet(a, doc.items, &.{"s"}, .{});
        switch (sheet_res) {
            .ok => |*ok| {
                var o = ok.*;
                o.deinit();
            },
            .refused => |r| _ = r.planeTwo(),
        }
        sheet_res = undefined;

        var sst_res = try decodeSharedStrings(a, doc.items, .{});
        switch (sst_res) {
            .ok => |*ok| {
                var o = ok.*;
                o.deinit();
            },
            .refused => |r| _ = r.planeTwo(),
        }
        sst_res = undefined;

        var tbl_res = try scanTable(a, doc.items, .{});
        switch (tbl_res) {
            .ok => |*ok| {
                var o = ok.*;
                o.deinit();
            },
            .refused => |r| _ = r.planeTwo(),
        }
        tbl_res = undefined;
    }
}

test "fuzz: no input can panic, leak, or decode two ways" {
    try std.testing.fuzz({}, fuzzDecodeTarget, .{
        .corpus = &[_][]const u8{
            "_x0041_",
            "_x005F_x0041_",
            "R&amp;D",
            "&#65;&#x42;",
            "<sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c></row></sheetData>",
            "<sheetData><row r=\"1\"><c><f>A1</f></c></row></sheetData>",
            "<si><t>a</t></si>",
            "<si><r><t>a</t></r><rPh><t>b</t></rPh></si>",
            "<sheetData><row r=\"1\"><c r=\"A1\" t=\"e\"><v>#REF!</v></c></row></sheetData>",
            "<sheetData><row r=\"1\"><c r=\"A1\" t=\"q\"><v>1</v></c></row></sheetData>",
            "\xFF\xFE",
            "",
        },
    });
}
