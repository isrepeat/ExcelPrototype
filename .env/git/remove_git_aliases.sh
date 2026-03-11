#!/usr/bin/env bash
set -euo pipefail

REPO_ROOT="$(git rev-parse --show-toplevel)"
cd "$REPO_ROOT"

if git config --local --get alias.amendall >/dev/null 2>&1; then
  git config --local --unset alias.amendall
  echo "Removed local git alias from $REPO_ROOT/.git/config:"
  echo "  amendall"
else
  echo "Local git alias amendall is not set in $REPO_ROOT/.git/config"
fi

echo
echo "Check aliases:"
echo "  git config --show-origin --get-regexp '^alias\\.'"
