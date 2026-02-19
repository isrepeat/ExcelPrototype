#!/usr/bin/env bash
set -euo pipefail

REPO_ROOT="$(git rev-parse --show-toplevel)"
ALIASES_FILE="$REPO_ROOT/Codex/aliases.sh"
BASHRC_FILE="${HOME}/.bashrc"
SOURCE_LINE="source \"$ALIASES_FILE\""

if [[ ! -f "$ALIASES_FILE" ]]; then
  echo "Error: aliases file not found: $ALIASES_FILE" >&2
  exit 1
fi

grep -qxF "$SOURCE_LINE" "$BASHRC_FILE" || echo "$SOURCE_LINE" >> "$BASHRC_FILE"
source "$ALIASES_FILE"

echo "codexlog alias is configured."
echo "Run: codexlog"
