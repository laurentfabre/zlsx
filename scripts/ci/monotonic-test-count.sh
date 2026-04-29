#!/usr/bin/env bash
# Gate 7: Monotonic test count.
#
# Fails the PR if the total count of `test "..."` blocks across
# tracked .zig files (under src/, pkg/, tests/) net-decreases.
#
# Rationale: makes stealth test deletion visible. Renaming or
# consolidating tests preserves the count; intentional deletion
# requires the escape label.
#
# Escape: PR label `delete-tests-ok`.
#
# Required env: BASE_SHA, HEAD_SHA, LABELS (JSON array of label names).

set -euo pipefail

: "${BASE_SHA:?BASE_SHA must be set}"
: "${HEAD_SHA:?HEAD_SHA must be set}"
: "${LABELS:=[]}"

if printf '%s' "$LABELS" | grep -q '"delete-tests-ok"'; then
  echo "Skipping: PR has the \`delete-tests-ok\` label."
  exit 0
fi

count_tests() {
  local ref="$1"
  local count=0
  while IFS= read -r f; do
    [ -z "$f" ] && continue
    n=$(git show "$ref:$f" 2>/dev/null | grep -c '^test "' || true)
    count=$((count + n))
  done < <(git ls-tree -r "$ref" --name-only \
    | grep -E '\.zig$' \
    | grep -E '^(src|pkg|tests)/')
  echo "$count"
}

base_count=$(count_tests "$BASE_SHA")
head_count=$(count_tests "$HEAD_SHA")

echo "base ($BASE_SHA) test count:  $base_count"
echo "head ($HEAD_SHA) test count:  $head_count"

if [ "$head_count" -lt "$base_count" ]; then
  cat >&2 <<EOF
ERROR: test count decreased: $base_count -> $head_count.

PRs may not net-delete tests. If you renamed or consolidated tests,
ensure the count is preserved.

Escape: add the \`delete-tests-ok\` label if you are intentionally
removing tests (e.g., as part of a deprecation).
EOF
  exit 1
fi

echo "OK: test count stable or growing."
