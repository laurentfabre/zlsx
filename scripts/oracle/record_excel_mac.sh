#!/usr/bin/env bash
# Excel-for-Mac oracle leg (M1b, goal_formula.md §8.2).
#
# Drives Excel over AppleScript: open the target workbook and nothing
# else, force a FULL calculation WITH a dependency-tree rebuild, wait
# for it to finish, save, close, and hand the file to the extractor.
#
# Three details that are not incidental:
#
#   CalculateFullRebuild, not CalculateFull. `CalculateFull` recomputes
#   every cell but walks the dependency tree it already has. Rebuilding
#   the tree first is what matters for rewritten and dynamic references,
#   which is exactly where a stale edge survives a "full" calculation.
#   The stale-dependency sentinel is what makes the difference visible.
#
#   Target workbook only. The script REFUSES to run while other
#   workbooks are open, rather than closing them: someone's unsaved work
#   is not ours to discard, and a stray add-in workbook recalculating
#   alongside ours would contaminate the timing and the result.
#
#   The file must live inside Excel's sandbox container. Excel for Mac
#   is sandboxed; a scripted open of a path it was not granted fails
#   SILENTLY — no error, no dialog, the workbook simply does not appear.
#   Staging into the container is what makes the open deterministic.
#
# Exit codes:
#   0  recorded
#   1  error
#   3  PARK — needs a human (see the message)
#
# Usage:
#   scripts/oracle/record_excel_mac.sh <input.xlsx> <output.xlsx>

set -euo pipefail

input="${1:?usage: record_excel_mac.sh <input.xlsx> <output.xlsx>}"
output="${2:?usage: record_excel_mac.sh <input.xlsx> <output.xlsx>}"

if [ "$(uname -s)" != "Darwin" ]; then
  echo "PARK: the Excel leg only runs on macOS." >&2
  exit 3
fi
if [ ! -d "/Applications/Microsoft Excel.app" ]; then
  echo "PARK: Microsoft Excel is not installed." >&2
  exit 3
fi

input_abs=$(cd "$(dirname "$input")" && pwd)/$(basename "$input")
output_dir=$(cd "$(dirname "$output")" && pwd)
output_abs="$output_dir/$(basename "$output")"
[ -f "$input_abs" ] || { echo "ERROR: no such input: $input_abs" >&2; exit 1; }

container="$HOME/Library/Containers/com.microsoft.Excel/Data/Documents/zlsx-oracle"
mkdir -p "$container"
staged="$container/$(basename "$output_abs")"
cp "$input_abs" "$staged"

# ─── preflight ───────────────────────────────────────────────────
#
# Every automation failure below is a human-action failure, so each one
# says exactly what the human has to do. Retrying them in a loop just
# burns time against a dialog nobody is looking at.

if ! probe=$(osascript -e 'tell application "Microsoft Excel" to return version' 2>&1); then
  cat >&2 <<EOF
PARK: Excel is not answering AppleScript.

  $probe

  Grant the terminal automation access to Excel:
    System Settings → Privacy & Security → Automation →
      <your terminal> → enable "Microsoft Excel"

  macOS asks for this with a dialog the first time; if it was dismissed,
  the permission has to be enabled by hand.
EOF
  exit 3
fi

open_count=$(osascript -e 'tell application "Microsoft Excel" to return count of workbooks' 2>/dev/null || echo "?")
if [ "$open_count" != "0" ]; then
  names=$(osascript -e 'tell application "Microsoft Excel" to return name of every workbook' 2>/dev/null || echo "?")
  cat >&2 <<EOF
PARK: Excel already has $open_count workbook(s) open: $names

  §8.2 requires an isolated instance holding the target workbook only.
  This script will not close them — unsaved work is not ours to discard,
  and another open workbook recalculates alongside ours.

  Close them in Excel (or quit Excel entirely) and re-run.
EOF
  exit 3
fi

# ─── drive ───────────────────────────────────────────────────────

script=$(mktemp -t zlsx-oracle-excel).scpt
trap 'rm -f "$script"' EXIT

cat >"$script" <<APPLESCRIPT
set stagedPath to "$staged"
tell application "Microsoft Excel"
    set display alerts to false

    -- LaunchServices rather than \`open workbook\`: the AppleScript verb
    -- fails silently for paths outside the sandbox, and returns success
    -- either way, so a failure is indistinguishable from a slow open.
    do shell script "open -a 'Microsoft Excel' " & quoted form of stagedPath
    set waited to 0
    repeat until (count of workbooks) > 0
        delay 1
        set waited to waited + 1
        if waited > 60 then
            set display alerts to true
            return "PARK:open-timeout"
        end if
    end repeat

    set wb to workbook 1
    set priorMode to (calculation as text)

    -- Automatic, so nothing is deferred, then the full rebuild.
    set calculation to calculation automatic
    calculate full rebuild

    -- Wait for completion. Excel calculates asynchronously; saving
    -- mid-calculation writes half-updated caches, which is a silently
    -- wrong oracle result rather than a visible failure.
    set spins to 0
    repeat until (calculation state is done)
        delay 1
        set spins to spins + 1
        if spins > 300 then
            set display alerts to true
            return "PARK:calculation-timeout"
        end if
    end repeat

    save wb
    close wb saving no
    set display alerts to true
    return "OK:" & priorMode
end tell
APPLESCRIPT

result=$(osascript "$script" 2>&1) || {
  echo "PARK: driving Excel failed: $result" >&2
  exit 3
}

case "$result" in
  PARK:*)
    cat >&2 <<EOF
PARK: Excel did not complete the run ($result).

  Bring Excel to the front and check for a dialog — a repair prompt, a
  sign-in prompt, or an autosave question will block every scripted open
  and close while still answering property queries, which is why this
  looks like silence rather than an error.
EOF
    exit 3
    ;;
  OK:*) prior_mode="${result#OK:}" ;;
  *) echo "ERROR: unexpected driver result: $result" >&2; exit 1 ;;
esac

[ -f "$staged" ] || { echo "ERROR: Excel did not leave a saved file" >&2; exit 1; }
cp "$staged" "$output_abs"

version=$(osascript -e 'tell application "Microsoft Excel" to return version' 2>/dev/null || echo unknown)
cat >"${output_abs%.xlsx}.provenance.json" <<JSON
{
  "adapter": "excel_mac",
  "app_build": "Microsoft Excel $version",
  "os": "$(sw_vers -productName) $(sw_vers -productVersion) (Darwin $(uname -r))",
  "locale": "${LANG:-en_US.UTF-8}",
  "extractor_version": "oracle-extractor-1",
  "workbook_digest": "$(shasum -a 256 "$output_abs" | cut -d' ' -f1)",
  "recorded": "$(date -u +%Y-%m-%d)"
}
JSON

echo "recorded: $output_abs"
echo "          Microsoft Excel $version (calculation mode on open: $prior_mode)"
