#!/usr/bin/env bash
set -euo pipefail

XLSM_PATH="${1:-Prototype/Prototype.xlsm}"
OUT_DIR="${2:-Prototype/xlsm_unpacked}"

rm -rf "$OUT_DIR"
mkdir -p "$OUT_DIR"

unzip -q "$XLSM_PATH" -d "$OUT_DIR"

# Pretty-print XML/RELS for readability.
# xmllint preserves namespaces; formatting is for human inspection.
while IFS= read -r -d '' f; do
  xmllint --format "$f" -o "$f"
done < <(find "$OUT_DIR" -type f \( -name '*.xml' -o -name '*.rels' \) -print0)

printf 'Unpacked to %s\n' "$OUT_DIR"
