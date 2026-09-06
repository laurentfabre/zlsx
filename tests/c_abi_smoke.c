/*
 * M9a1/M9a2/S3a header compile gate. Compiled (never linked, never run) as
 * part of `zig build test`, so the header side of the 3-file
 * transaction is verified by a C compiler rather than by eye:
 *   - the header parses as C;
 *   - every ZLSX_HAS_* feature macro is defined;
 *   - every M9a1/M9a2/S3a export has a prototype (their addresses
 *     are taken);
 *   - struct sizes match the layout the design note pins
 *     (docs/plans/c-abi-status-v1.md) on this target.
 */
#include <zlsx.h>

#if !defined(ZLSX_HAS_FINGERPRINT)
#error "ZLSX_HAS_FINGERPRINT missing"
#endif
#if !defined(ZLSX_HAS_MARK_RECALC)
#error "ZLSX_HAS_MARK_RECALC missing"
#endif
#if !defined(ZLSX_HAS_RECALC)
#error "ZLSX_HAS_RECALC missing"
#endif
#if !defined(ZLSX_HAS_EVAL)
#error "ZLSX_HAS_EVAL missing"
#endif
#if !defined(ZLSX_HAS_CANCEL)
#error "ZLSX_HAS_CANCEL missing"
#endif
#if !defined(ZLSX_HAS_SAVE_BUFFER)
#error "ZLSX_HAS_SAVE_BUFFER missing"
#endif
#if !defined(ZLSX_HAS_SAVE_WITH_RECALC)
#error "ZLSX_HAS_SAVE_WITH_RECALC missing"
#endif
#if !defined(ZLSX_HAS_WRITER_RECALC)
#error "ZLSX_HAS_WRITER_RECALC missing"
#endif
#if !defined(ZLSX_HAS_FORMULAS_V2)
#error "ZLSX_HAS_FORMULAS_V2 missing"
#endif
#if !defined(ZLSX_HAS_STRUCTURAL_EDITS)
#error "ZLSX_HAS_STRUCTURAL_EDITS missing"
#endif
#if !defined(ZLSX_HAS_PIVOTS)
#error "ZLSX_HAS_PIVOTS missing"
#endif
#if !defined(ZLSX_HAS_DEFINED_NAMES)
#error "ZLSX_HAS_DEFINED_NAMES missing"
#endif
#if !defined(ZLSX_HAS_CONDITIONAL_FORMATS)
#error "ZLSX_HAS_CONDITIONAL_FORMATS missing"
#endif
#if !defined(ZLSX_HAS_ANCHORS)
#error "ZLSX_HAS_ANCHORS missing"
#endif
#if !defined(ZLSX_HAS_SHEET_PROPS)
#error "ZLSX_HAS_SHEET_PROPS missing"
#endif
#if !defined(ZLSX_HAS_SHEET_STATE)
#error "ZLSX_HAS_SHEET_STATE missing"
#endif
#if !defined(ZLSX_HAS_ROWS_FORMULAS)
#error "ZLSX_HAS_ROWS_FORMULAS missing"
#endif
#if !defined(ZLSX_HAS_EMBEDDING_WRITE)
#error "ZLSX_HAS_EMBEDDING_WRITE missing"
#endif
#if !defined(ZLSX_HAS_EMBEDDABLE_ROWS)
#error "ZLSX_HAS_EMBEDDABLE_ROWS missing"
#endif
#if !defined(ZLSX_HAS_EMBEDDING_SWEEPS)
#error "ZLSX_HAS_EMBEDDING_SWEEPS missing"
#endif

#define ZLSX_STATIC_ASSERT(cond, name) typedef char name[(cond) ? 1 : -1]

ZLSX_STATIC_ASSERT(sizeof(zlsx_census_entry_v1) == 16, census_entry_size);
ZLSX_STATIC_ASSERT(sizeof(zlsx_diag_v1) == 96, diag_size);
ZLSX_STATIC_ASSERT(sizeof(zlsx_run_v1) == 104, run_size);
ZLSX_STATIC_ASSERT(sizeof(zlsx_resolved_v1) == 80, resolved_size);
ZLSX_STATIC_ASSERT(sizeof(zlsx_recalc_report_v1) == 168, report_size);
ZLSX_STATIC_ASSERT(sizeof(zlsx_value_elem_v1) == 32, value_elem_size);
ZLSX_STATIC_ASSERT(sizeof(zlsx_value_v1) == 56, value_size);
ZLSX_STATIC_ASSERT(sizeof(zlsx_formula_cell_v1) == 40, formula_cell_size);
/* S3b slice 10: the sheet-state codes are src/c_abi.zig's literals
 * (`ZLSX_SHEET_STATE_*`, pinned there too) — two hand-maintained
 * spellings, one value each. */
ZLSX_STATIC_ASSERT(ZLSX_SHEET_STATE_VISIBLE == 0, sheet_state_visible_code);
ZLSX_STATIC_ASSERT(ZLSX_SHEET_STATE_HIDDEN == 1, sheet_state_hidden_code);
ZLSX_STATIC_ASSERT(ZLSX_SHEET_STATE_VERY_HIDDEN == 2, sheet_state_very_hidden_code);

/* Taking each export's address forces the prototypes to exist and to
 * have function type; the array keeps the compiler from eliding them. */
static const void *const m9a1_exports[] = {
    (const void *)&zlsx_engine_fingerprint,
    (const void *)&zlsx_editor_mark_recalc_on_load,
    (const void *)&zlsx_editor_recalculate,
    (const void *)&zlsx_editor_evaluate,
    (const void *)&zlsx_cancel_token_new,
    (const void *)&zlsx_cancel_token_trigger,
    (const void *)&zlsx_cancel_token_free,
    (const void *)&zlsx_value_release,
    (const void *)&zlsx_recalc_report_release,
    (const void *)&zlsx_diag_release,
};

static const void *const m9a2_exports[] = {
    (const void *)&zlsx_buffer_release,
    (const void *)&zlsx_editor_save_to_buffer,
    (const void *)&zlsx_open_buffer,
    (const void *)&zlsx_editor_save_with_recalc,
    (const void *)&zlsx_writer_save_with_recalc,
    (const void *)&zlsx_sheet_writer_write_row_with_formulas_v2,
};

const void *zlsx_c_abi_smoke_anchor(void);
const void *zlsx_c_abi_smoke_anchor(void) { return m9a1_exports[0]; }
static const void *const s3a_exports[] = {
    (const void *)&zlsx_editor_insert_row,
    (const void *)&zlsx_editor_delete_row,
    (const void *)&zlsx_editor_insert_column,
    (const void *)&zlsx_editor_delete_column,
    (const void *)&zlsx_editor_add_sheet,
    (const void *)&zlsx_editor_rename_sheet,
    (const void *)&zlsx_editor_delete_sheet,
    (const void *)&zlsx_editor_rename_table_column,
    (const void *)&zlsx_editor_pivots_ndjson,
    /* S3b slice 2: the defined-names read rides the S3a gate. */
    (const void *)&zlsx_editor_defined_names_ndjson,
    /* S3b slice 6: the conditional-formats read rides it too. */
    (const void *)&zlsx_editor_conditional_formats_ndjson,
    /* S3b slice 7: the anchors read. */
    (const void *)&zlsx_editor_anchors_ndjson,
    /* S3b slice 9: the sheet-props and calc-props reads. */
    (const void *)&zlsx_editor_sheet_props_ndjson,
    (const void *)&zlsx_editor_calc_props_ndjson,
};

const void *zlsx_c_abi_smoke_anchor_m9a2(void);
const void *zlsx_c_abi_smoke_anchor_m9a2(void) { return m9a2_exports[0]; }
const void *zlsx_c_abi_smoke_anchor_s3a(void);
const void *zlsx_c_abi_smoke_anchor_s3a(void) { return s3a_exports[0]; }

/* S3b slice 10: sheet visibility on the reader handle. */
static const void *const s3b10_exports[] = {
    (const void *)&zlsx_sheet_state,
};
const void *zlsx_c_abi_smoke_anchor_s3b10(void);
const void *zlsx_c_abi_smoke_anchor_s3b10(void) { return s3b10_exports[0]; }

/* S3b slice 11: formula text and error tags on the row iterator. */
static const void *const s3b11_exports[] = {
    (const void *)&zlsx_rows_formula_at,
    (const void *)&zlsx_rows_formula_ref_at,
    (const void *)&zlsx_rows_error_at,
};
const void *zlsx_c_abi_smoke_anchor_s3b11(void);
const void *zlsx_c_abi_smoke_anchor_s3b11(void) { return s3b11_exports[0]; }

/* S3c slice 1: the embedding write on the editor handle. The coverage
 * descriptor is an array element with a frozen layout: pinned here on
 * the C side as the Zig side pins it. */
_Static_assert(sizeof(zlsx_emb_coverage_v1) == 88, "zlsx_emb_coverage_v1 is 88 bytes");
_Static_assert(offsetof(zlsx_emb_coverage_v1, vectors) == 48, "vectors at 48");
_Static_assert(offsetof(zlsx_emb_coverage_v1, hashes) == 64, "hashes at 64");
_Static_assert(offsetof(zlsx_emb_coverage_v1, sheet_idx) == 80, "sheet_idx at 80");
_Static_assert(offsetof(zlsx_emb_coverage_v1, include_formulas) == 84, "include_formulas at 84");
static const void *const s3c1_exports[] = {
    (const void *)&zlsx_editor_set_embeddings,
};
const void *zlsx_c_abi_smoke_anchor_s3c1(void);
const void *zlsx_c_abi_smoke_anchor_s3c1(void) { return s3c1_exports[0]; }

/* S3c slice 4: recovery_in_cells on the write — bit 0 of the same
 * export's flags word, under the same macro (0.9.0 ships both). */
#if !defined(ZLSX_EMB_WRITE_RECOVERY_IN_CELLS)
#error "ZLSX_EMB_WRITE_RECOVERY_IN_CELLS missing"
#endif
_Static_assert(ZLSX_EMB_WRITE_RECOVERY_IN_CELLS == 1u, "recovery_in_cells is flags bit 0");

/* S3c slice 2: the embeddable-rows read on the editor handle. */
static const void *const s3c2_exports[] = {
    (const void *)&zlsx_editor_embeddable_rows_ndjson,
};
const void *zlsx_c_abi_smoke_anchor_s3c2(void);
const void *zlsx_c_abi_smoke_anchor_s3c2(void) { return s3c2_exports[0]; }

/* S3c slice 3: the embedding sweeps on the editor handle. */
_Static_assert(sizeof(zlsx_prune_report_v1) == 40, "zlsx_prune_report_v1 is 40 bytes");
_Static_assert(offsetof(zlsx_prune_report_v1, redacted) == 8, "redacted at 8");
_Static_assert(offsetof(zlsx_prune_report_v1, stale) == 16, "stale at 16");
_Static_assert(offsetof(zlsx_prune_report_v1, fresh) == 24, "fresh at 24");
_Static_assert(offsetof(zlsx_prune_report_v1, valid_empty) == 32, "valid_empty at 32");
static const void *const s3c3_exports[] = {
    (const void *)&zlsx_editor_prune_embeddings,
    (const void *)&zlsx_editor_strip_embeddings,
};
const void *zlsx_c_abi_smoke_anchor_s3c3(void);
const void *zlsx_c_abi_smoke_anchor_s3c3(void) { return s3c3_exports[0]; }
