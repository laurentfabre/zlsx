#!/usr/bin/env bash
#
# publish_homebrew_tap.sh — push the zlsx Homebrew formula to the
# personal tap so `brew install laurentfabre/zlsx/zlsx` works.
#
# Usage:
#   scripts/publish_homebrew_tap.sh [VERSION] [--dry-run]
#
# Arguments:
#   VERSION   — version to publish (default: read from build.zig.zon).
#   --dry-run — fill the formula and print the diff, but don't push
#               (still clones the tap into a tempdir).
#
# Env vars:
#   ZLSX_TAP_REPO     — override tap repo (default: laurentfabre/homebrew-zlsx).
#   ZLSX_FORMULA_FILE — override local formula template (default:
#                       packaging/homebrew/zlsx.rb).
#
# Behaviour:
#   1. Resolves version (from arg or build.zig.zon).
#   2. Downloads SHA256SUMS from the GitHub release at that tag.
#   3. Rewrites the local formula with the release version + 4 sha256s.
#   4. Creates the tap repo if it doesn't exist (public, no description).
#   5. Clones it into a tempdir, writes Formula/zlsx.rb, commits, pushes.
#   6. Cleans up the tempdir.

set -euo pipefail

DRY_RUN=false
VERSION=""
for arg in "$@"; do
  case "$arg" in
    --dry-run) DRY_RUN=true ;;
    -h|--help)
      sed -n '2,30p' "$0" | sed 's/^# *//; s/^#$//'
      exit 0
      ;;
    -*)
      echo "unknown flag: $arg" >&2
      exit 2
      ;;
    *)
      if [ -z "$VERSION" ]; then VERSION="$arg"; else
        echo "unexpected extra argument: $arg" >&2
        exit 2
      fi
      ;;
  esac
done

# Default version from build.zig.zon.
if [ -z "$VERSION" ]; then
  VERSION=$(awk -F'"' '/^[[:space:]]*\.version[[:space:]]*=/{print $2; exit}' build.zig.zon)
  if [ -z "$VERSION" ]; then
    echo "error: could not read .version from build.zig.zon — pass VERSION explicitly" >&2
    exit 1
  fi
fi
TAG="v${VERSION}"
TAP_REPO="${ZLSX_TAP_REPO:-laurentfabre/homebrew-zlsx}"
FORMULA_SRC="${ZLSX_FORMULA_FILE:-packaging/homebrew/zlsx.rb}"

for cmd in gh git curl awk sed mktemp; do
  if ! command -v "$cmd" >/dev/null 2>&1; then
    echo "error: $cmd not found in PATH" >&2; exit 1
  fi
done
if [ ! -f "$FORMULA_SRC" ]; then
  echo "error: formula not found at $FORMULA_SRC" >&2; exit 1
fi

echo "publishing zlsx $TAG → tap $TAP_REPO (dry-run=$DRY_RUN)"

# ─── 1. Fetch SHA256SUMS from the release ──────────────────────────────
WORK=$(mktemp -d -t zlsx-publish-XXXXXX)
trap 'rm -rf "$WORK"' EXIT
SUMS="$WORK/SHA256SUMS"

echo "  fetching SHA256SUMS from release ${TAG}…"
gh release download "$TAG" -R "$(gh repo view --json nameWithOwner -q .nameWithOwner)" \
  -p SHA256SUMS -D "$WORK" >/dev/null

sha_for() {
  local asset="$1"
  local sha
  sha=$(awk -v a="$asset" '$2 == a { print $1 }' "$SUMS")
  if [ -z "$sha" ]; then
    echo "error: no sha256 entry for $asset in SHA256SUMS" >&2; exit 1
  fi
  printf "%s" "$sha"
}

SHA_MAC_ARM=$(sha_for "zlsx-${VERSION}-aarch64-apple-darwin.tar.gz")
SHA_MAC_X86=$(sha_for "zlsx-${VERSION}-x86_64-apple-darwin.tar.gz")
SHA_LNX_ARM=$(sha_for "zlsx-${VERSION}-aarch64-linux-musl.tar.gz")
SHA_LNX_X86=$(sha_for "zlsx-${VERSION}-x86_64-linux-musl.tar.gz")

# ─── 2. Rewrite the formula with $VERSION + 4 shas ────────────────────
GENERATED="$WORK/zlsx.rb"
# Position-dependent sha256 substitution: walk the formula and swap the
# 4 hex strings in order (arm/intel macOS, arm/intel linux). A simple
# `awk` pass tracks which sha256 occurrence we're on and rewrites it.
awk -v v="$VERSION" \
    -v s1="$SHA_MAC_ARM" -v s2="$SHA_MAC_X86" \
    -v s3="$SHA_LNX_ARM" -v s4="$SHA_LNX_X86" '
  BEGIN { n = 0 }
  /^[[:space:]]*version[[:space:]]+"/ {
    sub(/"[^"]*"/, "\"" v "\"")
    print; next
  }
  /^[[:space:]]*sha256[[:space:]]+"/ {
    n++
    if      (n == 1) sub(/"[^"]*"/, "\"" s1 "\"")
    else if (n == 2) sub(/"[^"]*"/, "\"" s2 "\"")
    else if (n == 3) sub(/"[^"]*"/, "\"" s3 "\"")
    else if (n == 4) sub(/"[^"]*"/, "\"" s4 "\"")
    print; next
  }
  { print }
' "$FORMULA_SRC" > "$GENERATED"

echo "  generated formula for $TAG:"
grep -E 'version |sha256 ' "$GENERATED" | sed 's/^/    /'

if [ "$DRY_RUN" = "true" ]; then
  echo "  --dry-run: would publish to $TAP_REPO (tempdir retained: $WORK)"
  trap - EXIT
  exit 0
fi

# ─── 3. Ensure the tap repo exists ────────────────────────────────────
if ! gh repo view "$TAP_REPO" >/dev/null 2>&1; then
  echo "  tap $TAP_REPO does not exist — creating it (public)…"
  gh repo create "$TAP_REPO" --public --description "Homebrew tap for zlsx" --confirm >/dev/null \
    || gh repo create "$TAP_REPO" --public --description "Homebrew tap for zlsx" >/dev/null
fi

# ─── 4. Clone, copy, commit, push ─────────────────────────────────────
CLONE="$WORK/tap"
git clone --quiet "https://github.com/$TAP_REPO.git" "$CLONE"
mkdir -p "$CLONE/Formula"
cp "$GENERATED" "$CLONE/Formula/zlsx.rb"

pushd "$CLONE" >/dev/null
if git diff --quiet; then
  echo "  formula unchanged — nothing to commit."
  popd >/dev/null
  exit 0
fi
git add Formula/zlsx.rb
git commit -m "zlsx $TAG" -m "Auto-published by scripts/publish_homebrew_tap.sh from $(gh repo view --json nameWithOwner -q .nameWithOwner)@$TAG." >/dev/null
git push --quiet origin HEAD
popd >/dev/null

echo "  published $TAG to https://github.com/$TAP_REPO"
echo "  users now run:"
echo "    brew tap ${TAP_REPO%/*}/${TAP_REPO#*/homebrew-}"
echo "    brew install ${TAP_REPO##*-}"
