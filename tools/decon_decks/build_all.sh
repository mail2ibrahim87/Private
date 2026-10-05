#!/bin/bash
# Usage: tools/decon_decks/build_all.sh <outdir> <logo.png> <pptx-skill-dir> <node_modules-dir>
set -e
OUT=$1; LOGO=$2; SK=$3; NM=$4
HERE=$(cd "$(dirname "$0")" && pwd)
mkdir -p "$OUT/_raw"
NODE_PATH=$NM PPTX_SKILL=$SK node "$HERE/build.js" "$OUT/_raw" "$LOGO"
for f in "$OUT"/_raw/*.pptx; do
  b=$(basename "$f")
  python3 "$HERE/animate.py" "$f" "${f%.pptx}.anim.json" "$OUT/$b"
  python3 "$SK/scripts/office/validate.py" "$OUT/$b" | tail -1
done
rm -rf "$OUT/_raw"
