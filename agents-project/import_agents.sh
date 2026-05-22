#!/usr/bin/env bash
set -euo pipefail

ZIP="agents.zip"
OUT_DIR="agents"

if [ ! -f "$ZIP" ]; then
  echo "Error: $ZIP not found in this directory. Place the provided agents.zip here and rerun."
  echo "Usage: sh import_agents.sh or npm run import"
  exit 1
fi

mkdir -p "$OUT_DIR"

if command -v unzip >/dev/null 2>&1; then
  unzip -o "$ZIP" -d "$OUT_DIR"
  echo "Extracted $ZIP -> $OUT_DIR/"
else
  echo "Error: 'unzip' command not found. Install it (e.g. 'brew install unzip') or extract $ZIP manually." >&2
  exit 2
fi
