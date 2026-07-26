#!/usr/bin/env bash
# emb-4 compat-matrix runner.
#
# Builds the emb-4 helpers, generates the fixture, and runs every matrix leg
# that can be driven without a GUI on this machine (zlsx control, openpyxl,
# LibreOffice headless if installed). For the GUI-only tools (Excel mac,
# Excel Win, Numbers) it stages a per-tool copy and prints the exact open →
# save → verify steps.
#
# Helpers are built through the canonical `zig build emb4-*` steps on every
# platform. (Before the 0.16 migration, macOS needed a standalone
# `zig build-exe -target aarch64-macos-none` path because 0.15.2's bundled
# libSystem had no arm64-macos slice; 0.16.0 links the build runner natively,
# so the hand-rolled module wiring is gone.)
#
# Usage:  tests/emb-4/run-matrix.sh [workdir]    (default workdir: /tmp/zlsx-emb4)
set -uo pipefail
cd "$(git rev-parse --show-toplevel)"

WORK="${1:-/tmp/zlsx-emb4}"
mkdir -p "$WORK/bin"
FIXTURE="$WORK/zlsx-emb4.xlsx"
# zlsx pins Zig 0.16.0 (see .github/workflows/ci.yml). Prefer the pinned local
# binary over whatever `zig` happens to be on PATH — the emb-4 helpers use
# 0.16's `std.process.Init` entry point and will not compile under 0.15.x.
ZIG_PIN=0.16.0
if [[ -z "${ZIG:-}" ]]; then
  if [[ -x "$HOME/.zvm/$ZIG_PIN/zig" ]]; then ZIG="$HOME/.zvm/$ZIG_PIN/zig"; else ZIG="$(command -v zig)"; fi
fi
PASS=0; FAIL=0

note() { printf '\n\033[1;36m── %s ──\033[0m\n' "$*"; }
ok()   { printf '   \033[1;32m✔ %s\033[0m\n' "$*"; }
warn() { printf '   \033[1;33m⚠ %s\033[0m\n' "$*"; }

ZV="$("$ZIG" version 2>/dev/null)"
[[ "$ZV" == "$ZIG_PIN" ]] || warn "zig is $ZV, expected $ZIG_PIN ($ZIG) — set ZIG=… to override"

# ---- build helpers --------------------------------------------------------
note "building helpers (zig build emb4-tools)"
# Invoke the installed binaries directly rather than via `zig build emb4-verify`:
# the build runner collapses any non-zero child exit into failure code 1, and the
# whole matrix is about telling 3 STRIPPED from 4 PARTS-ONLY from 5 ORPHANED-REL.
"$ZIG" build emb4-tools || exit 1
BIN=zig-out/bin
emb4_gen()     { "$BIN/zlsx-emb4-fixture" "$1"; }
emb4_verify()  { "$BIN/zlsx-emb4-verify" "$1"; }
emb4_passive() { "$BIN/zlsx-emb4-passive-save" "$1" "$2"; }

# ---- fixture + baseline ---------------------------------------------------
note "generate fixture"
emb4_gen "$FIXTURE" || exit 1

note "baseline (writer output, must PASS)"
if emb4_verify "$FIXTURE" >/dev/null; then ok "baseline PASS (exit 0)"; PASS=$((PASS+1)); else warn "baseline FAILED — stop, the writer regressed"; exit 1; fi

# ---- automated leg: zlsx control (open → save → verify) -------------------
note "leg: zlsx (control, expect PASS)"
cp "$FIXTURE" "$WORK/zlsx.in.xlsx"
emb4_passive "$WORK/zlsx.in.xlsx" "$WORK/zlsx.xlsx" >/dev/null
emb4_verify "$WORK/zlsx.xlsx" >/dev/null; rc=$?
[[ $rc -eq 0 ]] && { ok "zlsx control PASS (exit 0)"; PASS=$((PASS+1)); } || { warn "zlsx control exit $rc"; FAIL=$((FAIL+1)); }

# ---- automated leg: openpyxl (informational, expect STRIPPED exit 3) ------
if python3 -c 'import openpyxl' 2>/dev/null; then
  note "leg: openpyxl (informational, expect STRIPPED)"
  python3 - "$FIXTURE" "$WORK/openpyxl.xlsx" <<'PY'
import sys, openpyxl
openpyxl.load_workbook(sys.argv[1]).save(sys.argv[2])
PY
  emb4_verify "$WORK/openpyxl.xlsx" >/dev/null; rc=$?
  [[ $rc -eq 3 ]] && ok "openpyxl STRIPPED (exit 3 — matches design expectation)" || warn "openpyxl exit $rc (expected 3)"
else
  warn "openpyxl not installed — skipping informational leg"
fi

# ---- automated leg: LibreOffice headless (if installed) -------------------
SOFFICE="$(command -v soffice || true)"
[[ -z "$SOFFICE" && -x /Applications/LibreOffice.app/Contents/MacOS/soffice ]] && SOFFICE=/Applications/LibreOffice.app/Contents/MacOS/soffice
if [[ -n "$SOFFICE" ]]; then
  note "leg: LibreOffice Calc (headless convert == save)"
  rm -rf "$WORK/lo"; mkdir -p "$WORK/lo"
  LO_OUT="$WORK/lo/$(basename "$FIXTURE")"
  # Watchdog: a cold `--convert-to` hangs on macOS 26.4 until LibreOffice has done
  # its first-run setup once via the GUI (open -a LibreOffice, settle, quit).
  ( "$SOFFICE" --headless --convert-to xlsx --outdir "$WORK/lo" "$FIXTURE" >/dev/null 2>&1 ) &
  LOPID=$!; for _ in $(seq 1 20); do [[ -f "$LO_OUT" ]] && break; kill -0 $LOPID 2>/dev/null || break; sleep 3; done
  kill -9 $LOPID 2>/dev/null
  if [[ -f "$LO_OUT" ]]; then
    emb4_verify "$LO_OUT" >/dev/null; rc=$?
    case $rc in
      0) ok "LibreOffice PASS (exit 0)"; PASS=$((PASS+1));;
      3) warn "LibreOffice STRIPPED (exit 3) — Calc rebuilds the archive, drops the parts";;
      *) warn "LibreOffice exit $rc — see $LO_OUT"; FAIL=$((FAIL+1));;
    esac
  else warn "LibreOffice produced no output in ~60s — on macOS, launch the GUI once first (open -a LibreOffice, let it settle, quit), then re-run"; fi
else
  warn "LibreOffice not installed (brew install --cask libreoffice to enable this leg)"
fi

# ---- GUI legs: stage copies + print manual procedure ----------------------
note "GUI legs — stage per-tool copies (manual open → File▸Save → close → verify)"
for tool in excel-mac excel-win numbers; do
  cp "$FIXTURE" "$WORK/$tool.xlsx"
  printf '   • %-10s  open %s  →  File▸Save (not Save As)  →  close, then:\n' "$tool" "$WORK/$tool.xlsx"
  printf '                %s %s\n' "$BIN/zlsx-emb4-verify" "$WORK/$tool.xlsx"
done
printf '   (Numbers preserves xlsx only via File▸Export To▸Excel — export over %s.)\n' "$WORK/numbers.xlsx"

note "summary"
printf '   automated legs passed: %d   failed/flagged: %d\n' "$PASS" "$FAIL"
printf '   verify exit codes: 0 PASS · 2 PARTIAL · 3 STRIPPED · 4 PARTS-ONLY · 5 ORPHANED-REL\n'
printf '   record verdicts in docs/plans/emb-4-compat-matrix.md\n'
