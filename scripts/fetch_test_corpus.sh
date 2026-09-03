#!/usr/bin/env bash
# Materialize the xlsx integration-test corpus described in
# docs/xlsx_test_corpus.md. Idempotent: existing files are kept.
#
#   Usage:
#     scripts/fetch_test_corpus.sh [target-dir]
#
# Three groups:
#   (1) Small base corpus — committed to the repo. Round-tripped by
#       tests/xlsx_corpus.zig as the smoke-test set.
#   (2) Large fixtures — fetched, NOT committed. Drive perf + memory
#       characterisation; tests skip cleanly if absent.
#   (3) Adversarial fixtures — fetched, NOT committed. Each one
#       exercises a known-bad input class (truncated ZIP, malformed
#       SST count, encrypted entries, OSS-Fuzz minimisations, …);
#       tests assert zlsx fails *cleanly* (typed error, no panic).
#
# A short locally-derived suite is generated at the end (truncated
# variants + bad-CRC) so we don't need a network round-trip for the
# obvious ZIP edge cases.
#
# Group (1) base corpus is committed to the repo, so its `fetch`
# calls never hit the network in CI; if one ever did and went 404,
# we still fail loudly. Groups (2) and (3) are downloaded each run
# from third-party endpoints we don't control — World Bank, ECDC,
# ONS, GitHub-hosted POI / openxlsx mirrors. Per-vendor SSL cert
# expiry, transient TLS resets, and rate limits would otherwise
# break every PR's CI for reasons unrelated to zlsx, so those
# fetches go through `fetch_optional` which warns and continues
# instead of aborting. Tests that need a missing fixture skip
# cleanly (see `tests/*_corpus.zig`'s `openOrSkip` pattern and
# the per-name skip table in `tests/package_corpus.zig`).

set -euo pipefail

dir="${1:-tests/corpus}"
mkdir -p "$dir"

# (1) Small base corpus — committed.
#     `openpyxl_chart.xlsx` is committed too but generated, not fetched:
#     scripts/gen_openpyxl_chart_fixture.py (the chart-sweep witness).
declare -a base_files=(
  "frictionless_2sheets.xlsx|https://raw.githubusercontent.com/frictionlessdata/datasets/main/files/excel/sample-2-sheets.xlsx"
  "openpyxl_guess_types.xlsx|https://github.com/fluidware/openpyxl/raw/master/openpyxl/tests/test_data/genuine/guess_types.xlsx"
  "phpoi_test1.xlsx|https://github.com/phax/ph-poi/raw/master/src/test/resources/excel/test1.xlsx"
  "worldbank_catalog.xlsx|https://databankfiles.worldbank.org/public/ddpext_download/world_bank_data_catalog.xlsx"
)

# (2) Large fixtures — perf / memory / many-sheet / many-merges.
#     CC-BY-4.0 / OGL / MIT / Apache-2.0; redistributable but big.
declare -a large_files=(
  "ecdc_covid.xlsx|https://opendata.ecdc.europa.eu/covid19/casedistribution/xlsx/"
  "ons_cpi_detailed.xlsx|https://www.ons.gov.uk/file?uri=/economy/inflationandpriceindices/datasets/consumerpriceinflation/current/consumerpriceinflationdetailedreferencetables.xlsx"
  "phpsheet_3654c.xlsx|https://raw.githubusercontent.com/PHPOffice/PhpSpreadsheet/master/tests/data/Reader/XLSX/issue.3654c.xlsx"
  "poi_57893_many_merges.xlsx|https://raw.githubusercontent.com/apache/poi/trunk/test-data/spreadsheet/57893-many-merges.xlsx"
  "poi_58325_db.xlsx|https://raw.githubusercontent.com/apache/poi/trunk/test-data/spreadsheet/58325_db.xlsx"
  "openxlsx_loadExample.xlsx|https://raw.githubusercontent.com/ycphs/openxlsx/master/inst/extdata/loadExample.xlsx"
  "wdi_excel.zip|https://databankfiles.worldbank.org/public/ddpext_download/WDI_excel.zip"
)

# (3) Adversarial fixtures — POI hand-crafted + OSS-Fuzz minimisations
#     + calamine edge cases + bare-ZIP malformation suite.
declare -a broken_files=(
  "poi_MalformedSSTCount.xlsx|https://raw.githubusercontent.com/apache/poi/trunk/test-data/spreadsheet/MalformedSSTCount.xlsx"
  "poi_xlsx_corrupted.xlsx|https://raw.githubusercontent.com/apache/poi/trunk/test-data/spreadsheet/xlsx-corrupted.xlsx"
  "poi_xxe_in_schema.xlsx|https://raw.githubusercontent.com/apache/poi/trunk/test-data/spreadsheet/xxe_in_schema.xlsx"
  "poi_crash_274d6342.xlsx|https://raw.githubusercontent.com/apache/poi/trunk/test-data/spreadsheet/crash-274d6342e4842d61be0fb48eaadad6208ae767ae.xlsx"
  "poi_crash_9bf3cd4b.xlsx|https://raw.githubusercontent.com/apache/poi/trunk/test-data/spreadsheet/crash-9bf3cd4bd6f50a8a9339d363c2c7af14b536865c.xlsx"
  "poi_clusterfuzz_xssf.xlsx|https://raw.githubusercontent.com/apache/poi/trunk/test-data/spreadsheet/clusterfuzz-testcase-minimized-POIXSSFFuzzer-5937385319563264.xlsx"
  "poi_workbook_password_2013.xlsx|https://raw.githubusercontent.com/apache/poi/trunk/test-data/spreadsheet/workbookProtection-workbook_password-2013.xlsx"
  "poi_poc_shared_strings.xlsx|https://raw.githubusercontent.com/apache/poi/trunk/test-data/spreadsheet/poc-shared-strings.xlsx"
  "poi_excel_with_trash_item.xlsx|https://raw.githubusercontent.com/apache/poi/trunk/test-data/spreadsheet/Excel_file_with_trash_item.xlsx"
  "calamine_non_monotonic_si.xlsx|https://raw.githubusercontent.com/tafia/calamine/master/tests/non_monotonic_si.xlsx"
  "calamine_encoded_entities.xlsx|https://raw.githubusercontent.com/tafia/calamine/master/tests/encoded_entities.xlsx"
  "calamine_empty_shared_string.xlsx|https://raw.githubusercontent.com/tafia/calamine/master/tests/empty_shared_string.xlsx"
  "calamine_empty_s_attribute.xlsx|https://raw.githubusercontent.com/tafia/calamine/master/tests/empty_s_attribute.xlsx"
  "ziprs_invalid_offset.zip|https://raw.githubusercontent.com/zip-rs/zip2/master/tests/data/invalid_offset.zip"
  "ziprs_invalid_cde_files_greater.zip|https://raw.githubusercontent.com/zip-rs/zip2/master/tests/data/invalid_cde_number_of_files_allocation_greater_offset.zip"
  "ziprs_aes_archive.zip|https://raw.githubusercontent.com/zip-rs/zip2/master/tests/data/aes_archive.zip"
  "ziprs_data_descriptor.zip|https://raw.githubusercontent.com/zip-rs/zip2/master/tests/data/data_descriptor.zip"
  "ziprs_comment_garbage.zip|https://raw.githubusercontent.com/zip-rs/zip2/master/tests/data/comment_garbage.zip"
  "ziprs_extended_timestamp_bad.zip|https://raw.githubusercontent.com/zip-rs/zip2/master/tests/data/extended_timestamp_bad.zip"
  "ziprs_misaligned_comment.zip|https://raw.githubusercontent.com/zip-rs/zip2/master/tests/data/misaligned_comment.zip"
)

fetch() {
  local entry="$1"
  local name="${entry%%|*}"
  local url="${entry#*|}"
  local dest="$dir/$name"
  if [[ -f "$dest" ]]; then
    printf '  · %-44s (already present, %s)\n' "$name" "$(du -h "$dest" | cut -f1)"
    return
  fi
  printf '  ↓ %-44s %s\n' "$name" "$url"
  curl -sfL --max-time 120 -o "$dest.tmp" "$url"
  mv "$dest.tmp" "$dest"
}

# Same as fetch, but tolerates curl failure: drops the .tmp,
# prints a SKIP line, and continues. Used for groups (2) and (3)
# whose endpoints are out of our control. Tests that depend on
# the missing fixture skip cleanly.
fetch_optional() {
  local entry="$1"
  local name="${entry%%|*}"
  local url="${entry#*|}"
  local dest="$dir/$name"
  if [[ -f "$dest" ]]; then
    printf '  · %-44s (already present, %s)\n' "$name" "$(du -h "$dest" | cut -f1)"
    return
  fi
  printf '  ↓ %-44s %s\n' "$name" "$url"
  if curl -sfL --max-time 120 -o "$dest.tmp" "$url"; then
    mv "$dest.tmp" "$dest"
  else
    local rc=$?
    rm -f "$dest.tmp"
    printf '  ⚠ %-44s SKIP (curl exit %d) — dependent tests will skip\n' "$name" "$rc" >&2
  fi
}

echo "(1) base corpus —"
for entry in "${base_files[@]}"; do fetch "$entry"; done

echo
echo "(2) large fixtures —"
for entry in "${large_files[@]}"; do fetch_optional "$entry"; done

# WDI ships as a zip-of-xlsx; extract idempotently. Write to .tmp
# first so a failed `unzip` (e.g. if the zip's internal layout
# changes upstream) doesn't leave a half-baked file that future
# runs treat as cached.
if [[ -f "$dir/wdi_excel.zip" && ! -f "$dir/wdi_excel.xlsx" ]]; then
  printf '  ↻ %-44s (extracting from wdi_excel.zip)\n' "wdi_excel.xlsx"
  unzip -p "$dir/wdi_excel.zip" WDIEXCEL.xlsx > "$dir/wdi_excel.xlsx.tmp"
  mv "$dir/wdi_excel.xlsx.tmp" "$dir/wdi_excel.xlsx"
fi

echo
echo "(3) adversarial fixtures —"
for entry in "${broken_files[@]}"; do fetch_optional "$entry"; done

echo
echo "(4) locally-derived adversarial variants —"
# Derive deterministic ZIP-edge-case fixtures from worldbank_catalog.xlsx.
# Keeps these byte-exact, regenerable, and license-trivial (we wrote them).
src="$dir/worldbank_catalog.xlsx"
if [[ -f "$src" ]] && command -v python3 >/dev/null 2>&1; then
  python3 - "$src" "$dir" <<'PY'
import os, struct, sys
# Windows default cp1252 stdout can't encode the `✎` glyph used below;
# force UTF-8 so the script runs identically across platforms.
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except (AttributeError, OSError):
    pass
src_path, out_dir = sys.argv[1], sys.argv[2]
src = open(src_path, "rb").read()

def flip_first_cdfh_crc(blob):
    # Locate the EOCD by walking back from end (no trailing comment in
    # our source fixture). Read cd_offset from the EOCD, then flip a
    # byte at CDFH+16 (= CRC32 LSB of the first central-directory
    # entry). PartStore.open verifies decompressed bytes against the
    # CDFH CRC32 eagerly, so flipping it surfaces as BadZip.
    eocd_sig = b"PK\x05\x06"
    eocd_off = blob.rfind(eocd_sig)
    if eocd_off < 0:
        raise SystemExit("source fixture has no EOCD")
    cd_offset = struct.unpack_from("<I", blob, eocd_off + 16)[0]
    out = bytearray(blob)
    out[cd_offset + 16] ^= 0xFF
    return bytes(out)

variants = {
    # Cut last 22 bytes — removes the EOCD record entirely. Reader
    # must fail with a typed error, not crash.
    "derived_truncated_pre_eocd.xlsx": src[:-22],
    # Cut at byte 50 — only the first LFH header survives, no
    # payload, no CDFH, no EOCD.
    "derived_truncated_mid_payload.xlsx": src[:50],
    # Cut at byte 4 — incomplete LFH signature.
    "derived_truncated_signature.xlsx": src[:4],
    # Flip the CRC byte in the first central-directory entry. See
    # flip_first_cdfh_crc above for why CDFH and not LFH.
    "derived_bad_crc32.xlsx": flip_first_cdfh_crc(src),
}
for name, data in variants.items():
    dest = os.path.join(out_dir, name)
    if os.path.exists(dest):
        print(f"  · {name:<44} (already present)")
        continue
    open(dest, "wb").write(data)
    print(f"  ✎ {name:<44} ({len(data)} bytes)")
PY
else
  echo "  (skipped — python3 or worldbank_catalog.xlsx missing)"
fi

echo
echo "corpus contents:"
ls -la "$dir"
