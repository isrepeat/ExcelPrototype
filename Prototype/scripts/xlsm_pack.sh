#!/usr/bin/env bash
set -euo pipefail

IN_DIR="${1:-Prototype/xlsm_unpacked}"
OUT_FILE="${2:-Prototype/Prototype_packed.xlsm}"

if [[ ! -d "$IN_DIR" ]]; then
  echo "Input dir not found: $IN_DIR" >&2
  exit 1
fi

TMP_DIR="$(mktemp -d)"
trap 'rm -rf "$TMP_DIR"' EXIT

cp -a "$IN_DIR/." "$TMP_DIR/"

# Minify XML/RELS back to compact form.
# We only collapse whitespace between tags to avoid changing text nodes.
while IFS= read -r -d '' f; do
  tmp_out="$f.tmp"
  xmllint --noblanks "$f" > "$tmp_out"
  # collapse whitespace between tags
  python - <<'PY' "$tmp_out" "$f"
import re, sys
src, dst = sys.argv[1], sys.argv[2]
text = open(src, 'r', encoding='utf-8').read()
text = re.sub(r'>\s+<', '><', text)
with open(dst, 'w', encoding='utf-8') as w:
    w.write(text)
PY
  rm -f "$tmp_out"
done < <(find "$TMP_DIR" -type f \( -name '*.xml' -o -name '*.rels' \) -print0)

# Build xlsm
rm -f "$OUT_FILE"
( cd "$TMP_DIR" && zip -qr "$OUT_FILE" . )

printf 'Packed to %s\n' "$OUT_FILE"
