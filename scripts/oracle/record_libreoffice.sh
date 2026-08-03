#!/usr/bin/env bash
# LibreOffice oracle leg (M1b, goal_formula.md §8.2).
#
# §8.2 asks for a "hardened" LO leg, and each word of that is a specific
# failure this script exists to prevent:
#
#   pinned invocation   — the binary and its version are recorded, not
#                         whatever `soffice` happens to resolve to.
#   dedicated profile   — a fresh `-env:UserInstallation` per run, so a
#                         developer's autocorrect settings, locale, or
#                         cached calc options cannot leak into a golden.
#   hard recalc forced  — LibreOffice will happily re-save a document it
#                         only LOADED, echoing every cached value
#                         straight back. `calculateAll()` is the
#                         difference between a recalculation and a copy.
#   volatile sentinel   — and because "forced" is a claim, the run must
#                         still prove it: the sentinel check in
#                         tests/oracle/sentinel.zig rejects a run whose
#                         volatile cell came back at its planted value.
#
# Usage:
#   scripts/oracle/record_libreoffice.sh <input.xlsx> <output.xlsx>

set -euo pipefail

soffice_bin="${SOFFICE:-/Applications/LibreOffice.app/Contents/MacOS/soffice}"
input="${1:?usage: record_libreoffice.sh <input.xlsx> <output.xlsx>}"
output="${2:?usage: record_libreoffice.sh <input.xlsx> <output.xlsx>}"

if [ ! -x "$soffice_bin" ]; then
  echo "PARK: LibreOffice not found at $soffice_bin" >&2
  echo "      Install it, or set SOFFICE to the soffice binary." >&2
  exit 3
fi

input_abs=$(cd "$(dirname "$input")" && pwd)/$(basename "$input")
output_dir=$(cd "$(dirname "$output")" && pwd)
output_abs="$output_dir/$(basename "$output")"
[ -f "$input_abs" ] || { echo "ERROR: no such input: $input_abs" >&2; exit 1; }

profile=$(mktemp -d)
trap 'rm -rf "$profile"' EXIT

version=$("$soffice_bin" --version 2>/dev/null | head -1 || echo "unknown")

# Let LibreOffice materialise the profile before the macro is installed
# into it. Writing the Basic library into an empty directory does not
# work: the first real start rewrites `user/basic/script.xlc`, and the
# library registration written beforehand is discarded — the macro then
# silently does not exist and the run produces no output at all.
"$soffice_bin" \
  --headless --norestore --nolockcheck --nodefault --nofirststartwizard \
  --terminate_after_init \
  -env:UserInstallation="file://$profile" \
  >"$profile/init.log" 2>&1 || true

# The macro. `calculateAll()` is the load-bearing line: without it this
# is a format conversion, not an oracle run.
#
# ONLY `Standard/Module1.xba` is replaced. The initialised profile
# already contains a registered, empty `Standard` library — rewriting
# its `script.xlb` or `user/basic/script.xlc` de-registers it, and the
# macro then does not exist. LibreOffice reports that by doing nothing
# at all: exit status 0, no output file, no message.
macro_dir="$profile/user/basic/Standard"
if [ ! -f "$macro_dir/script.xlb" ]; then
  echo "ERROR: LibreOffice did not initialise a Basic library in $profile" >&2
  sed -n '1,20p' "$profile/init.log" >&2 || true
  exit 1
fi

cat >"$macro_dir/Module1.xba" <<'XBA'
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">
<script:module xmlns:script="http://openoffice.org/2000/script" script:name="Module1" script:language="StarBasic"><![CDATA[
Sub RecalcAndSave(inPath As String, outPath As String)
    Dim loadArgs(0) As New com.sun.star.beans.PropertyValue
    loadArgs(0).Name = "Hidden"
    loadArgs(0).Value = True

    Dim doc As Object
    doc = StarDesktop.loadComponentFromURL(ConvertToURL(inPath), "_blank", 0, loadArgs())

    ' The hard recalc. Not optional, and not the same as saving:
    ' storeToURL on a freshly loaded document writes the caches back
    ' exactly as they came in.
    doc.calculateAll()

    Dim saveArgs(0) As New com.sun.star.beans.PropertyValue
    saveArgs(0).Name = "FilterName"
    saveArgs(0).Value = "Calc MS Excel 2007 XML"
    doc.storeToURL(ConvertToURL(outPath), saveArgs())
    doc.close(False)
End Sub
]]></script:module>
XBA

rm -f "$output_abs"
set +e
"$soffice_bin" \
  --headless --norestore --nolockcheck --nodefault --nofirststartwizard \
  -env:UserInstallation="file://$profile" \
  "macro:///Standard.Module1.RecalcAndSave(\"$input_abs\",\"$output_abs\")" \
  >"$profile/soffice.log" 2>&1
rc=$?
set -e

if [ ! -f "$output_abs" ]; then
  echo "ERROR: LibreOffice produced no output (exit $rc)" >&2
  sed -n '1,40p' "$profile/soffice.log" >&2 || true
  exit 1
fi

# Provenance goes next to the artefact. The sentinel check that decides
# whether this run counts runs later, over the extracted workbook.
cat >"${output_abs%.xlsx}.provenance.json" <<JSON
{
  "adapter": "libreoffice",
  "app_build": "$version",
  "os": "$(sw_vers -productName 2>/dev/null || uname -s) $(sw_vers -productVersion 2>/dev/null || uname -r)",
  "locale": "${LANG:-en_US.UTF-8}",
  "extractor_version": "oracle-extractor-1",
  "workbook_digest": "$(shasum -a 256 "$output_abs" | cut -d' ' -f1)",
  "recorded": "$(date -u +%Y-%m-%d)"
}
JSON

echo "recorded: $output_abs"
echo "          $version"
