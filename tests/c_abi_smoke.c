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

#define ZLSX_STATIC_ASSERT(cond, name) typedef char name[(cond) ? 1 : -1]

ZLSX_STATIC_ASSERT(sizeof(zlsx_census_entry_v1) == 16, census_entry_size);
ZLSX_STATIC_ASSERT(sizeof(zlsx_diag_v1) == 96, diag_size);
ZLSX_STATIC_ASSERT(sizeof(zlsx_run_v1) == 104, run_size);
ZLSX_STATIC_ASSERT(sizeof(zlsx_resolved_v1) == 80, resolved_size);
ZLSX_STATIC_ASSERT(sizeof(zlsx_recalc_report_v1) == 168, report_size);
ZLSX_STATIC_ASSERT(sizeof(zlsx_value_elem_v1) == 32, value_elem_size);
ZLSX_STATIC_ASSERT(sizeof(zlsx_value_v1) == 56, value_size);
ZLSX_STATIC_ASSERT(sizeof(zlsx_formula_cell_v1) == 40, formula_cell_size);

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
};

const void *zlsx_c_abi_smoke_anchor_m9a2(void);
const void *zlsx_c_abi_smoke_anchor_m9a2(void) { return m9a2_exports[0]; }
const void *zlsx_c_abi_smoke_anchor_s3a(void);
const void *zlsx_c_abi_smoke_anchor_s3a(void) { return s3a_exports[0]; }
