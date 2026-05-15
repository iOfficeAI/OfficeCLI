#!/bin/bash
set -euo pipefail

if [ "$#" -lt 1 ]; then
    echo "Usage: scripts/build-rhwp-sidecars.sh <output-dir> [rid] [Release|Debug]" >&2
    exit 2
fi

OUT_DIR="$1"
TARGET_RID="${2:-}"
CONFIG="${3:-Release}"
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
ROOT_DIR="$(cd "$SCRIPT_DIR/.." && pwd)"
BRIDGE_PROJECT="$ROOT_DIR/src/rhwp-officecli-bridge/rhwp-officecli-bridge.csproj"
API_MANIFEST="$ROOT_DIR/src/rhwp-field-bridge/Cargo.toml"

detect_local_rid() {
    local OS
    local ARCH
    local LIBC
    OS=$(uname -s | tr '[:upper:]' '[:lower:]')
    ARCH=$(uname -m)
    LIBC="gnu"
    if [ "$OS" = "linux" ]; then
        if command -v ldd >/dev/null 2>&1 && ldd --version 2>&1 | grep -qi musl; then
            LIBC="musl"
        elif [ -f /etc/alpine-release ]; then
            LIBC="musl"
        fi
    fi
    case "$OS" in
        darwin)
            case "$ARCH" in
                arm64) echo "osx-arm64" ;;
                x86_64) echo "osx-x64" ;;
            esac ;;
        linux)
            case "$ARCH" in
                x86_64)
                    if [ "$LIBC" = "musl" ]; then echo "linux-musl-x64"; else echo "linux-x64"; fi ;;
                aarch64|arm64)
                    if [ "$LIBC" = "musl" ]; then echo "linux-musl-arm64"; else echo "linux-arm64"; fi ;;
            esac ;;
    esac
}

LOCAL_RID="$(detect_local_rid)"
if [ -z "$LOCAL_RID" ]; then
    echo "Unsupported local platform for rhwp sidecars: $(uname -s) $(uname -m)" >&2
    exit 1
fi

if [ -n "$TARGET_RID" ] && [ "$TARGET_RID" != "$LOCAL_RID" ]; then
    echo "Skipping rhwp sidecars for $TARGET_RID; local Rust sidecar build is $LOCAL_RID."
    exit 0
fi

mkdir -p "$OUT_DIR"

BRIDGE_TMP="$(mktemp -d)"
cleanup() {
    rm -rf "$BRIDGE_TMP"
}
trap cleanup EXIT

echo "Building rhwp-officecli-bridge ($LOCAL_RID)..."
dotnet publish "$BRIDGE_PROJECT" \
    -c "$CONFIG" \
    -r "$LOCAL_RID" \
    -o "$BRIDGE_TMP" \
    --self-contained true \
    -p:PublishSingleFile=true \
    -p:PublishTrimmed=false \
    --nologo -v quiet

BRIDGE_OUT=""
for candidate in \
    "$BRIDGE_TMP/rhwp-officecli-bridge" \
    "$BRIDGE_TMP/rhwp-officecli-bridge.exe" \
    "$BRIDGE_TMP/rhwp-officecli-bridge.dll"
do
    if [ -f "$candidate" ]; then
        BRIDGE_OUT="$candidate"
        break
    fi
done

if [ -z "$BRIDGE_OUT" ]; then
    echo "rhwp-officecli-bridge publish completed but no bridge executable was found." >&2
    exit 1
fi

cp "$BRIDGE_OUT" "$OUT_DIR/$(basename "$BRIDGE_OUT")"
chmod +x "$OUT_DIR/$(basename "$BRIDGE_OUT")" 2>/dev/null || true

echo "Building rhwp-field-bridge ($LOCAL_RID)..."
CONFIG_LOWER="$(printf '%s' "$CONFIG" | tr '[:upper:]' '[:lower:]')"
API_FEATURES="${OFFICECLI_RHWP_API_FEATURES:-native-skia}"
if [ "$API_FEATURES" = "none" ]; then
    API_FEATURES=""
fi
FEATURE_ARGS=()
if [ -n "$API_FEATURES" ]; then
    FEATURE_ARGS=(--features "$API_FEATURES")
fi
if [ "$CONFIG_LOWER" = "release" ]; then
    cargo build --manifest-path "$API_MANIFEST" --release "${FEATURE_ARGS[@]}"
    API_BIN="$ROOT_DIR/src/rhwp-field-bridge/target/release/rhwp-field-bridge"
else
    cargo build --manifest-path "$API_MANIFEST" "${FEATURE_ARGS[@]}"
    API_BIN="$ROOT_DIR/src/rhwp-field-bridge/target/debug/rhwp-field-bridge"
fi

if [ ! -f "$API_BIN" ]; then
    echo "rhwp-field-bridge build completed but no executable was found at $API_BIN." >&2
    exit 1
fi

cp "$API_BIN" "$OUT_DIR/rhwp-field-bridge"
chmod +x "$OUT_DIR/rhwp-field-bridge"

if [ "$(uname -s)" = "Darwin" ]; then
    xattr -d com.apple.quarantine "$OUT_DIR/rhwp-officecli-bridge" 2>/dev/null || true
    xattr -d com.apple.quarantine "$OUT_DIR/rhwp-field-bridge" 2>/dev/null || true
    codesign -s - -f "$OUT_DIR/rhwp-officecli-bridge" 2>/dev/null || true
    codesign -s - -f "$OUT_DIR/rhwp-field-bridge" 2>/dev/null || true
fi

echo "rhwp sidecars copied to $OUT_DIR"
