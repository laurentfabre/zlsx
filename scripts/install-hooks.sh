#!/usr/bin/env bash
# install-hooks.sh — one-time setup after `git clone`.
# Points git at the tracked hooks directory under scripts/githooks/.
# Survives clones and cross-machine work, unlike unversioned .git/hooks/.

set -euo pipefail

cd "$(git rev-parse --show-toplevel)"

git config core.hooksPath scripts/githooks
chmod +x scripts/githooks/*

echo "install-hooks: core.hooksPath -> scripts/githooks/"
echo "Active hooks:"
ls scripts/githooks/ | sed 's/^/  /'
