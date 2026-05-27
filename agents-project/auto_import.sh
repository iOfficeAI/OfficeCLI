#!/usr/bin/env bash
set -euo pipefail

# Search locations to find an uploaded agents zip and copy it into this project.
# Locations searched (in order): current dir, workspace root, /tmp, ~/Downloads

CANDIDATES=(
  "./agents.zip"
  "$(pwd)/agents.zip"
  "$(dirname "$0")/agents.zip"
  "/tmp/agents.zip"
  "${HOME}/Downloads/agents.zip"
)

# Also allow pattern matches
PATTERNS=(
  "agents*.zip"
  "*agents*.zip"
)

DEST_DIR="$(dirname "$0")"
ZIP_DEST="$DEST_DIR/agents.zip"

found=""

for p in "${CANDIDATES[@]}"; do
  if [ -f "$p" ]; then
    found="$p"
    break
  fi
done

if [ -z "$found" ]; then
  # search common locations for matching patterns
  search_paths=("$(pwd)" "$(dirname "$0")" "/tmp" "${HOME}/Downloads")
  for sp in "${search_paths[@]}"; do
    for pat in "${PATTERNS[@]}"; do
      matches=("$sp"/$pat)
      for m in "${matches[@]}"; do
        if [ -f "$m" ]; then
          found="$m"
          break 3
        fi
      done
    done
  done
fi

if [ -z "$found" ]; then
  echo "No agents.zip found in common locations."
  echo "Searched: ${CANDIDATES[*]} and patterns ${PATTERNS[*]} in pwd, script dir, /tmp, and ~/Downloads."
  echo "You can also pass a path: ./auto_import.sh /path/to/agents.zip"
  exit 2
fi

# If a path was provided as argument, prefer it
if [ "$#" -ge 1 ]; then
  if [ -f "$1" ]; then
    found="$1"
  else
    echo "Provided path '$1' not found." >&2
    exit 3
  fi
fi

# Copy into project folder and run import
cp -f "$found" "$ZIP_DEST"
chmod 644 "$ZIP_DEST"

echo "Copied: $found -> $ZIP_DEST"

# Run import script
sh "$DEST_DIR/import_agents.sh"
