#!/usr/bin/env bash
set -euo pipefail

REPO_ROOT="$(git rev-parse --show-toplevel)"
cd "$REPO_ROOT"

git config --local alias.amendall '!git add -A && git commit --amend --no-edit'

echo "Configured local git alias in $REPO_ROOT/.git/config:"
echo "  amendall = !git add -A && git commit --amend --no-edit"
echo
echo "Check aliases:"
echo "  git config --show-origin --get-regexp '^alias\\.'"
