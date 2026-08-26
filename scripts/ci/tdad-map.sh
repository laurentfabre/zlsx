#!/usr/bin/env bash
# TDAD — Test-Dependency-Aware-Development map.
#
# For each .zig file changed in the PR diff, emits the inline
# `test "..."` blocks declared in that file plus related corpus /
# fuzz / binding test surfaces. Surfaces this as a PR comment so an
# author (human or agent) sees up-front which tests their patch will
# affect — empirical 70% regression reduction in TDAD studies vs
# procedural-TDD prompts alone.
#
# Output: Markdown to stdout. Empty (one-line note) when no .zig
# files changed. Always exits 0 — this is information, not a gate.
#
# Required env: BASE_SHA, HEAD_SHA.
# Local invocation:
#   BASE_SHA=$(git merge-base origin/main HEAD) HEAD_SHA=HEAD \
#     bash scripts/ci/tdad-map.sh

set -euo pipefail

: "${BASE_SHA:?BASE_SHA must be set}"
: "${HEAD_SHA:?HEAD_SHA must be set}"

changed=$(git diff --name-only "$BASE_SHA" "$HEAD_SHA" | grep -E '\.zig$' || true)

if [ -z "$changed" ]; then
  echo "_No \`.zig\` files changed — TDAD map empty._"
  exit 0
fi

echo "## TDAD — tests potentially affected by this PR"
echo ""
echo "_Mechanically-derived from changed \`.zig\` files. Use as a hint, not a verdict._"
echo ""

# 1. Per-file inline test blocks. Use HEAD_SHA's blob (not the
# working tree) so deleted files / renames don't blow up. Show line
# numbers from the head-side file when present.
emitted_inline=0
for f in $changed; do
  # Pull the head-side content via git to avoid filesystem assumptions.
  if ! git cat-file -e "$HEAD_SHA:$f" 2>/dev/null; then
    # File removed in HEAD — skip the inline-test scan for it.
    continue
  fi
  blob=$(git show "$HEAD_SHA:$f")
  tests=$(printf '%s\n' "$blob" | grep -nE '^test "[^"]+"' || true)
  if [ -n "$tests" ]; then
    if [ "$emitted_inline" -eq 0 ]; then
      echo "### Inline tests in changed files"
      echo ""
      emitted_inline=1
    fi
    echo "**\`$f\`**"
    echo ""
    while IFS= read -r line; do
      lineno=${line%%:*}
      rest=${line#*:}
      # Strip the leading `test ` and surrounding quotes for display.
      name=$(printf '%s' "$rest" | sed -E 's/^test "([^"]*)".*/\1/')
      echo "- L${lineno}: \`test \"${name}\"\`"
    done <<< "$tests"
    echo ""
  fi
done

if [ "$emitted_inline" -eq 0 ]; then
  echo "_No inline \`test \"...\"\` blocks in the changed files._"
  echo ""
fi

# 2. Related corpus / fuzz / binding surfaces that the change-set touches.
related=""
add() { related+="- $1"$'\n'; }

if printf '%s\n' "$changed" | grep -qE '^src/(xlsx|writer)\.zig$|^unicode/|^src/formula/'; then
  add "\`tests/xlsx_corpus.zig\` — reader integration over the fixture corpus"
  add "\`tests/package_corpus.zig\` — package-layer integration"
  add "\`zig build fuzz\` — coverage-guided fuzz (Linux x64; \`src/xlsx.zig\` + \`pkg/store.zig\`)"
fi

if printf '%s\n' "$changed" | grep -qE '^pkg/'; then
  add "\`tests/package_corpus.zig\` — package-layer integration over the fixture corpus"
  add "\`zig build fuzz\` — coverage-guided fuzz on \`pkg/store.zig\` (\`decodeXmlEntities\`, \`looksExternal\`)"
fi

if printf '%s\n' "$changed" | grep -qE '^src/c_abi\.zig$'; then
  add "\`bindings/python/tests/test_basic.py\` — Python binding pytest suite (exercises every C ABI export)"
  add "Verify \`include/zlsx.h\` and \`bindings/python/zlsx/_ffi.py\` are updated in lockstep (the C-ABI 3-file gate enforces this)"
fi

if printf '%s\n' "$changed" | grep -qE '^src/cli\.zig$|^src/extract_images_main\.zig$'; then
  add "\`windows-runtime\` CI job — CLI smoke (\`zlsx list-sheets / meta / cells\`)"
fi

if printf '%s\n' "$changed" | grep -qE '^bindings/python/'; then
  add "\`bindings/python/tests/test_basic.py\` — Python binding pytest suite"
  add "\`windows-runtime\` CI job — wheel install + import + reader smoke + pytest"
fi

if [ -n "$related" ]; then
  echo "### Related test surfaces"
  echo ""
  printf '%s' "$related"
  echo ""
fi

# 3. Diff summary (file count + line-counts) — useful at-a-glance.
echo "### Diff summary"
echo ""
file_count=$(printf '%s\n' "$changed" | wc -l | tr -d ' ')
added=$(git diff --numstat "$BASE_SHA" "$HEAD_SHA" -- '*.zig' | awk '{a+=$1} END {print a+0}')
removed=$(git diff --numstat "$BASE_SHA" "$HEAD_SHA" -- '*.zig' | awk '{r+=$2} END {print r+0}')
echo "- ${file_count} \`.zig\` file(s) changed — +${added} / -${removed}"
echo ""
