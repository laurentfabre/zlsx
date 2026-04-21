#!/usr/bin/env bash
# Materialize the xlsx integration-test corpus described in
# docs/xlsx_test_corpus.md. Idempotent: existing files are kept.
#
#   Usage:
#     scripts/fetch_test_corpus.sh [target-dir]
#
# The downloads are small (~110 KB total) and stable public URLs
# (government data portal + long-lived library mirrors). If any URL
# goes 404, the script fails loudly with the URL that broke — don't
# silently skip, because a gap here means the corresponding test
# stops exercising its edge case without anyone noticing.

set -euo pipefail

dir="${1:-tests/corpus}"
mkdir -p "$dir"

declare -a files=(
  "frictionless_2sheets.xlsx|https://raw.githubusercontent.com/frictionlessdata/datasets/main/files/excel/sample-2-sheets.xlsx"
  "openpyxl_guess_types.xlsx|https://github.com/fluidware/openpyxl/raw/master/openpyxl/tests/test_data/genuine/guess_types.xlsx"
  "phpoi_test1.xlsx|https://github.com/phax/ph-poi/raw/master/src/test/resources/excel/test1.xlsx"
  "worldbank_catalog.xlsx|https://databankfiles.worldbank.org/public/ddpext_download/world_bank_data_catalog.xlsx"
)

for entry in "${files[@]}"; do
  name="${entry%%|*}"
  url="${entry#*|}"
  dest="$dir/$name"
  if [[ -f "$dest" ]]; then
    printf '  · %-40s (already present, %s)\n' "$name" "$(du -h "$dest" | cut -f1)"
    continue
  fi
  printf '  ↓ %-40s %s\n' "$name" "$url"
  curl -sfL --max-time 60 -o "$dest.tmp" "$url"
  mv "$dest.tmp" "$dest"
done

echo
echo "corpus contents:"
ls -la "$dir"
