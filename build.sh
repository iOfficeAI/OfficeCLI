#!/bin/bash
set -e

PROJECT="src/officecli/officecli.csproj"
ALL_TARGETS="osx-arm64:officecli-mac-arm64 osx-x64:officecli-mac-x64 linux-x64:officecli-linux-x64 linux-arm64:officecli-linux-arm64 linux-musl-x64:officecli-linux-alpine-x64 linux-musl-arm64:officecli-linux-alpine-arm64 win-x64:officecli-win-x64.exe win-arm64:officecli-win-arm64.exe"
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"

# Detect current platform RID
detect_local_rid() {
    local OS=$(uname -s | tr '[:upper:]' '[:lower:]')
    local ARCH=$(uname -m)
    local LIBC="gnu"
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

# Find target entry by RID
find_target() {
    local RID="$1"
    for target in $ALL_TARGETS; do
        if [ "${target%%:*}" = "$RID" ]; then
            echo "$target"
            return
        fi
    done
}

build_config() {
    local CONFIG="$1"
    local TARGETS="$2"
    local OUTPUT="bin/$(echo "$CONFIG" | tr '[:upper:]' '[:lower:]')"

    rm -rf "$OUTPUT"
    mkdir -p "$OUTPUT"

    for target in $TARGETS; do
        RID="${target%%:*}"
        NAME="${target##*:}"
        TMPDIR=$(mktemp -d)

        echo "[$CONFIG] Building $RID -> $NAME"
        dotnet publish "$PROJECT" -c "$CONFIG" -r "$RID" -o "$TMPDIR" --nologo -v quiet

        # Atomic replace: stage as .new alongside the target, sign there, then rename.
        # Overwriting the binary in place would trash the text segment of any
        # running officecli process that happens to be mmap'd on this path
        # (macOS does not block ETXTBSY), leaving it stuck in uninterruptible
        # `UE` state on the next code page fault.
        if [ -f "$TMPDIR/officecli.exe" ]; then
            cp "$TMPDIR/officecli.exe" "$OUTPUT/$NAME.new"
        else
            cp "$TMPDIR/officecli" "$OUTPUT/$NAME.new"
        fi

        # Ad-hoc codesign on macOS (required by AppleSystemPolicy).
        # Done on the staged .new copy so the live binary is never mutated in place.
        if [ "$(uname -s)" = "Darwin" ] && [[ "$RID" == osx-* ]]; then
            codesign -s - -f "$OUTPUT/$NAME.new" 2>/dev/null || true
        fi

        mv -f "$OUTPUT/$NAME.new" "$OUTPUT/$NAME"
        cp "$TMPDIR/officecli.pdb" "$OUTPUT/${NAME%.*}.pdb"

        "$SCRIPT_DIR/scripts/build-rhwp-sidecars.sh" "$OUTPUT" "$RID" "$CONFIG"
        copy_platform_sidecar_assets "$OUTPUT" "$NAME"

        rm -rf "$TMPDIR"
    done

    rm -rf src/officecli/bin src/officecli/obj

    echo ""
    echo "$CONFIG build complete:"
    ls -lh "$OUTPUT"
}

copy_platform_sidecar_assets() {
    local OUTPUT="$1"
    local NAME="$2"
    local ASSET_BASE="${NAME%.exe}"

    if [ -f "$OUTPUT/rhwp-field-bridge" ]; then
        cp "$OUTPUT/rhwp-field-bridge" "$OUTPUT/${ASSET_BASE}-rhwp-field-bridge"
        chmod +x "$OUTPUT/${ASSET_BASE}-rhwp-field-bridge" 2>/dev/null || true
    fi
    if [ -f "$OUTPUT/rhwp-officecli-bridge" ]; then
        cp "$OUTPUT/rhwp-officecli-bridge" "$OUTPUT/${ASSET_BASE}-rhwp-officecli-bridge"
        chmod +x "$OUTPUT/${ASSET_BASE}-rhwp-officecli-bridge" 2>/dev/null || true
    fi
    if [ -f "$OUTPUT/rhwp-field-bridge.exe" ]; then
        cp "$OUTPUT/rhwp-field-bridge.exe" "$OUTPUT/${ASSET_BASE}-rhwp-field-bridge.exe"
    fi
    if [ -f "$OUTPUT/rhwp-officecli-bridge.exe" ]; then
        cp "$OUTPUT/rhwp-officecli-bridge.exe" "$OUTPUT/${ASSET_BASE}-rhwp-officecli-bridge.exe"
    fi
}

CONFIG="${1:-release}"

case "$CONFIG" in
    release|Release)
        LOCAL_RID=$(detect_local_rid)
        TARGET=$(find_target "$LOCAL_RID")
        if [ -z "$TARGET" ]; then
            echo "Unsupported platform: $(uname -s) $(uname -m)"
            exit 1
        fi
        build_config "Release" "$TARGET"
        ;;
    debug|Debug)
        LOCAL_RID=$(detect_local_rid)
        TARGET=$(find_target "$LOCAL_RID")
        if [ -z "$TARGET" ]; then
            echo "Unsupported platform: $(uname -s) $(uname -m)"
            exit 1
        fi
        build_config "Debug" "$TARGET"
        ;;
    all)
        build_config "Release" "$ALL_TARGETS"
        ;;
    *)
        echo "Usage: ./build.sh [release|debug|all]"
        echo "  release  - Build Release for current platform (default)"
        echo "  debug    - Build Debug for current platform"
        echo "  all      - Build Release for all platforms"
        exit 1
        ;;
esac
