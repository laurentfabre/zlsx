#!/usr/bin/env bash
# Regen gate for the vendored Unicode tables (M1a).
#
# Each generated file pins, in its own header, the Unicode version and
# the SHA-256 of every UCD input it was derived from. This script:
#
#   1. reads those pins;
#   2. downloads the inputs from the VERSIONED UCD path (never
#      `.../UCD/latest/...`, which moves under us);
#   3. verifies each download against the pinned digest;
#   4. re-runs `scripts/gen_unicode_tables.py`;
#   5. fails on any byte of difference.
#
# So a hand-edit of a "DO NOT EDIT" table, a silently-bumped Unicode
# revision, and a generator change that was never re-run all surface
# here rather than as a mystery tokenizer divergence months later.
#
# Network failures are tolerated (warn + skip), matching
# `fetch_test_corpus.sh` — an expired upstream certificate must not
# block every PR. A digest mismatch or a table diff is NEVER tolerated.
#
# Usage: scripts/ci/check_unicode_tables.sh

set -euo pipefail

repo_root=$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)
cd "$repo_root"

work=$(mktemp -d)
trap 'rm -rf "$work"' EXIT

skipped=0

sha256_of() {
  if command -v sha256sum >/dev/null 2>&1; then
    sha256sum "$1" | cut -d' ' -f1
  else
    shasum -a 256 "$1" | cut -d' ' -f1
  fi
}

# Pinned Unicode version, read from the generated file itself.
pinned_version() {
  sed -n 's/^pub const unicode_version: \[\]const u8 = "\(.*\)";$/\1/p' "$1" | head -1
}

# Pinned SHA-256 for one input. The casefold and xid generators emit
# `// SHA-256 of input: <hex>` (single input); the NFC generator names
# each file: `// SHA-256 of UnicodeData.txt: <hex>`.
pinned_digest() {
  local file="$1" input_name="$2" got
  got=$(sed -n "s|^// SHA-256 of ${input_name}: \([0-9a-f]*\)\$|\1|p" "$file" | head -1)
  if [ -z "$got" ]; then
    got=$(sed -n 's|^// SHA-256 of input: \([0-9a-f]*\)$|\1|p' "$file" | head -1)
  fi
  printf '%s' "$got"
}

# Download one UCD input and check it against the pin. Returns 1 if the
# download failed (tolerated), 2 if the digest mismatched (fatal).
fetch_input() {
  local version="$1" name="$2" expected="$3" dest="$4"
  local url="https://www.unicode.org/Public/${version}/ucd/${name}"

  if ! curl -fsSL --retry 2 --max-time 120 -o "$dest" "$url"; then
    echo "WARN: could not download ${url} — skipping regen check" >&2
    return 1
  fi

  local actual
  actual=$(sha256_of "$dest")
  if [ "$actual" != "$expected" ]; then
    cat >&2 <<EOF
ERROR: ${name} does not match the digest pinned in the generated table.

  url:      ${url}
  expected: ${expected}
  actual:   ${actual}

Either upstream republished the file at this version (report it), or the
table was regenerated from a different input than its header claims. Do
not "fix" this by editing the header.
EOF
    return 2
  fi
  return 0
}

# check <generated-file> <generator-mode> <input-name> [<flag>=<input-name>...]
#
# Every input after the first is given as the generator flag that carries
# it, so a mode with three inputs (casing) needs no special case here.
check() {
  local out_file="$1" mode="$2" primary="$3"
  shift 3
  local version
  version=$(pinned_version "$out_file")
  if [ -z "$version" ]; then
    echo "ERROR: ${out_file} has no pinned unicode_version" >&2
    exit 1
  fi
  echo "==> ${out_file} (Unicode ${version}, mode ${mode})"

  local primary_path="$work/$primary"
  local rc=0
  fetch_input "$version" "$primary" "$(pinned_digest "$out_file" "$primary")" "$primary_path" || rc=$?
  if [ "$rc" -eq 2 ]; then exit 1; fi
  if [ "$rc" -eq 1 ]; then skipped=1; return 0; fi

  local args=(--mode "$mode" --input "$primary_path" --output "$work/regenerated.zig")
  local spec flag name extra_path
  for spec in "$@"; do
    flag="${spec%%=*}"
    name="${spec#*=}"
    extra_path="$work/$name"
    rc=0
    fetch_input "$version" "$name" "$(pinned_digest "$out_file" "$name")" "$extra_path" || rc=$?
    if [ "$rc" -eq 2 ]; then exit 1; fi
    if [ "$rc" -eq 1 ]; then skipped=1; return 0; fi
    args+=("$flag" "$extra_path")
  done

  python3 scripts/gen_unicode_tables.py "${args[@]}" >/dev/null

  # Compare after normalising both sides through `zig fmt`. The
  # committed tables are formatted (CI enforces it) while the generator
  # emits plain rows, and `zig fmt` column-aligns long scalar lists — so
  # a raw diff reports whitespace churn on a table that is byte-correct.
  cp "$out_file" "$work/committed.zig"
  if command -v zig >/dev/null 2>&1; then
    zig fmt "$work/committed.zig" >/dev/null
    zig fmt "$work/regenerated.zig" >/dev/null
  fi

  if ! diff -u "$work/committed.zig" "$work/regenerated.zig" >"$work/table.diff"; then
    cat >&2 <<EOF
ERROR: ${out_file} is not what the generator produces from its pinned inputs.

$(head -40 "$work/table.diff")

Regenerate with:
  python3 scripts/gen_unicode_tables.py --mode ${mode} \\
      --input <${primary}> $(printf '%s <%s> ' "${@/=/ }")--output ${out_file}
EOF
    exit 1
  fi
  echo "    identical"
}

check unicode/tables/xid_data.zig xid DerivedCoreProperties.txt
check unicode/tables/casefold_data.zig casefold CaseFolding.txt
check unicode/tables/nfc_data.zig nfc UnicodeData.txt --excl=CompositionExclusions.txt
check unicode/tables/casing_data.zig casing UnicodeData.txt \
  --special=SpecialCasing.txt --props=DerivedCoreProperties.txt

if [ "$skipped" -eq 1 ]; then
  echo "unicode-tables: PARTIAL — at least one input could not be downloaded."
else
  echo "unicode-tables: OK — every table matches its pinned inputs."
fi
