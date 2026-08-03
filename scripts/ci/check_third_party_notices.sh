#!/usr/bin/env bash
# Artifact-content gate for THIRD_PARTY_NOTICES.md (M1a).
#
# The Unicode Character Database data compiled into every zlsx binary is
# licensed on the condition that its notice appears with all copies, or
# in associated documentation. That obligation is discharged by a file
# only if the file actually ships — so this gate asserts it is wired
# into every distribution channel, and (in `stage`/`archive` mode) that
# it is really inside the artifact about to be published.
#
# Usage:
#   check_third_party_notices.sh                 # repo wiring (default)
#   check_third_party_notices.sh repo
#   check_third_party_notices.sh stage <dir>     # staged release tree
#   check_third_party_notices.sh archive <file>  # .tar.gz / .zip / wheel / sdist

set -euo pipefail

repo_root=$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)
notices="THIRD_PARTY_NOTICES.md"
# The load-bearing line: if the license text is gone, the file is a
# stub and satisfies nothing.
marker="UNICODE LICENSE V3"

fail() {
  echo "ERROR: $*" >&2
  exit 1
}

require_grep() {
  local pattern="$1" file="$2" why="$3"
  grep -qF -- "$pattern" "$file" || fail "$why (expected '$pattern' in $file)"
}

check_repo() {
  cd "$repo_root"

  [ -f "$notices" ] || fail "$notices is missing from the repository root"
  require_grep "$marker" "$notices" "$notices does not carry the Unicode license text"

  # 1. Zig package: `zig fetch` copies only what `.paths` lists.
  require_grep "\"$notices\"" build.zig.zon \
    "build.zig.zon .paths does not include $notices, so the Zig package omits it"

  # 2. Release tarballs / zips.
  require_grep "$notices" .github/workflows/release.yml \
    "release.yml does not stage $notices into the release archives"

  # 3. Homebrew formula (installs into the keg's doc dir).
  require_grep "$notices" packaging/homebrew/zlsx.rb \
    "the Homebrew formula does not install $notices"

  # 4. Python wheel + sdist. `license-files` puts it in the wheel's
  #    dist-info; the sdist `include` list puts it in the tarball.
  local pyproject="bindings/python/pyproject.toml"
  require_grep "$notices" "$pyproject" \
    "$pyproject does not ship $notices"
  grep -q "^license-files = .*$notices" "$pyproject" ||
    fail "$pyproject: $notices missing from license-files (wheel would omit it)"
  awk '/^\[tool\.hatch\.build\.targets\.sdist\]/{f=1} f&&/THIRD_PARTY_NOTICES\.md/{found=1} END{exit !found}' \
    "$pyproject" || fail "$pyproject: $notices missing from the sdist include list"
  [ -e "bindings/python/$notices" ] ||
    fail "bindings/python/$notices does not resolve (the symlink to the root file is broken)"

  # 5. Every UCD-derived table points at the notices file. A new
  #    generated table that forgets the attribution header is the way
  #    this obligation quietly stops being met.
  local tables=(
    unicode/tables/xid_data.zig
    unicode/tables/nfc_data.zig
    src/unicode/tables/casefold_data.zig
  )
  local t
  for t in "${tables[@]}"; do
    [ -f "$t" ] || fail "expected generated table $t is missing"
    require_grep "$notices" "$t" "$t lacks the Unicode attribution header"
    require_grep "Unicode License v3" "$t" "$t lacks the Unicode attribution header"
  done

  echo "third-party-notices: OK — wired into zon paths, release staging, Homebrew, wheel and sdist."
}

check_stage() {
  local dir="${1:?usage: check_third_party_notices.sh stage <dir>}"
  [ -d "$dir" ] || fail "staging directory '$dir' does not exist"
  [ -f "$dir/$notices" ] || fail "$dir/$notices is missing — the archive would ship without it"
  require_grep "$marker" "$dir/$notices" "$dir/$notices is present but carries no license text"
  echo "third-party-notices: OK — $dir/$notices present."
}

check_archive() {
  local archive="${1:?usage: check_third_party_notices.sh archive <file>}"
  [ -f "$archive" ] || fail "archive '$archive' does not exist"

  local listing
  case "$archive" in
    *.tar.gz | *.tgz) listing=$(tar -tzf "$archive") ;;
    *.zip | *.whl) listing=$(unzip -Z1 "$archive") ;;
    *) fail "unsupported archive type: $archive" ;;
  esac

  printf '%s\n' "$listing" | grep -q "$notices" ||
    fail "$archive does not contain $notices"
  echo "third-party-notices: OK — $archive contains $notices."
}

mode="${1:-repo}"
case "$mode" in
  repo) check_repo ;;
  stage) check_stage "${2:-}" ;;
  archive) check_archive "${2:-}" ;;
  *) fail "unknown mode '$mode' (expected repo | stage | archive)" ;;
esac
