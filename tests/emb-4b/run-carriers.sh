#!/usr/bin/env bash
# emb-4B carrier-survival runner.
#
# emb-4 established that `xl/zlsxEmbeddings/*` does not survive Apple
# Numbers or LibreOffice Calc — both rebuild the archive and drop every
# unknown part. emb-4B asks whether anything else in the package does,
# so a small recovery record can ride in a second carrier.
#
# One fixture carries the same marker in six places (see
# tests/emb-4b/carriers.zig), so a single open → save round-trip per
# tool measures every carrier at once. That matters most for the
# GUI-only legs, where the alternative is six manual saves per tool.
#
# Usage:  tests/emb-4b/run-carriers.sh [workdir]   (default: /tmp/zlsx-emb4b)
set -uo pipefail
cd "$(git rev-parse --show-toplevel)"

WORK="${1:-/tmp/zlsx-emb4b}"
mkdir -p "$WORK"
FIXTURE="$WORK/fixture.xlsx"

# Same pin and same reasoning as tests/emb-4/run-matrix.sh: the helpers
# use 0.16's `std.process.Init` entry point and do not compile under 0.15.x.
ZIG_PIN=0.16.0
if [[ -z "${ZIG:-}" ]]; then
  if [[ -x "$HOME/.zvm/$ZIG_PIN/zig" ]]; then ZIG="$HOME/.zvm/$ZIG_PIN/zig"; else ZIG="$(command -v zig)"; fi
fi

note() { printf '\n\033[1;36m── %s ──\033[0m\n' "$*"; }
ok()   { printf '   \033[1;32m✔ %s\033[0m\n' "$*"; }
warn() { printf '   \033[1;33m⚠ %s\033[0m\n' "$*"; }

ZV="$("$ZIG" version 2>/dev/null)"
[[ "$ZV" == "$ZIG_PIN" ]] || warn "zig is $ZV, expected $ZIG_PIN ($ZIG) — set ZIG=… to override"

# Invoke the installed binaries directly. The verifier's exit code IS the
# carrier-loss count, and `zig build emb4b-verify` would flatten every
# non-zero count to build-failure 1.
note "building helpers (zig build emb4b-tools)"
"$ZIG" build emb4b-tools || exit 1
BIN=zig-out/bin

note "generate fixture"
"$BIN/zlsx-emb4b-fixture" "$FIXTURE" || exit 1

note "baseline (writer output — all six carriers must plant)"
if "$BIN/zlsx-emb4b-verify" "$FIXTURE"; then
  ok "baseline: 0 carriers lost"
else
  warn "baseline lost $? carriers — stop, the fixture generator regressed"
  exit 1
fi

# ---- automated leg: openpyxl ---------------------------------------------
if python3 -c 'import openpyxl' 2>/dev/null; then
  note "leg: openpyxl (rebuilds the archive)"
  python3 - "$FIXTURE" "$WORK/openpyxl.xlsx" <<'PY'
import sys, openpyxl
openpyxl.load_workbook(sys.argv[1]).save(sys.argv[2])
PY
  "$BIN/zlsx-emb4b-verify" "$WORK/openpyxl.xlsx"; warn "openpyxl lost $? carriers"
else
  warn "openpyxl not installed — skipping leg"
fi

# ---- automated leg: LibreOffice Calc headless ----------------------------
SOFFICE="$(command -v soffice || true)"
[[ -z "$SOFFICE" && -x /Applications/LibreOffice.app/Contents/MacOS/soffice ]] && SOFFICE=/Applications/LibreOffice.app/Contents/MacOS/soffice
if [[ -n "$SOFFICE" ]]; then
  note "leg: LibreOffice Calc (headless convert == save)"
  rm -rf "$WORK/lo"; mkdir -p "$WORK/lo"
  LO_OUT="$WORK/lo/$(basename "$FIXTURE")"
  # Watchdog: a cold `--convert-to` hangs on macOS until LibreOffice has
  # done its first-run setup once via the GUI (open -a LibreOffice, settle, quit).
  ( "$SOFFICE" --headless --convert-to xlsx --outdir "$WORK/lo" "$FIXTURE" >/dev/null 2>&1 ) &
  LOPID=$!; for _ in $(seq 1 25); do [[ -f "$LO_OUT" ]] && break; kill -0 $LOPID 2>/dev/null || break; sleep 3; done
  kill -9 $LOPID 2>/dev/null
  if [[ -f "$LO_OUT" ]]; then
    "$BIN/zlsx-emb4b-verify" "$LO_OUT"; warn "LibreOffice lost $? carriers"
  else
    warn "LibreOffice produced no output in ~75s — on macOS, launch the GUI once first, then re-run"
  fi
else
  warn "LibreOffice not installed (brew install --cask libreoffice to enable this leg)"
fi

# ---- GUI legs -------------------------------------------------------------
# Numbers and Excel are not driven from here. AppleScript `export` against
# Numbers hangs with -1712 (AppleEvent timed out) even with an explicit
# 400s timeout block, and diagnosing that needs assistive access that a
# headless run cannot assume. Staged for a manual pass instead.
note "GUI legs — stage copies (manual open → save → verify)"
for tool in excel-mac excel-win numbers; do
  cp "$FIXTURE" "$WORK/$tool.xlsx"
  printf '   • %-10s  open %s  →  save  →  close, then:\n' "$tool" "$WORK/$tool.xlsx"
  printf '                %s %s\n' "$BIN/zlsx-emb4b-verify" "$WORK/$tool.xlsx"
done
printf '   (Numbers writes xlsx only via File▸Export To▸Excel, and the target must be a\n'
printf '    non-TCC folder such as ~/ — exporting into ~/Documents fails silently.)\n'

note "summary"
printf '   verifier exit code = number of carriers lost (0–6); 64 = usage/IO error\n'
printf '   record verdicts in docs/plans/emb-4b-carrier-matrix.md\n'
