#!/bin/bash
set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PROJECT="$SCRIPT_DIR/src/officecli/officecli.csproj"
BINARY_NAME="officecli"

# Detect platform
OS=$(uname -s | tr '[:upper:]' '[:lower:]')
ARCH=$(uname -m)

case "$OS" in
    darwin)
        case "$ARCH" in
            arm64) RID="osx-arm64" ;;
            x86_64) RID="osx-x64" ;;
            *) echo "Unsupported architecture: $ARCH"; exit 1 ;;
        esac
        ;;
    linux)
        # Detect musl libc (Alpine, etc.)
        LIBC="gnu"
        if command -v ldd >/dev/null 2>&1 && ldd --version 2>&1 | grep -qi musl; then
            LIBC="musl"
        elif [ -f /etc/alpine-release ]; then
            LIBC="musl"
        fi
        case "$ARCH" in
            x86_64)
                if [ "$LIBC" = "musl" ]; then RID="linux-musl-x64"; else RID="linux-x64"; fi ;;
            aarch64|arm64)
                if [ "$LIBC" = "musl" ]; then RID="linux-musl-arm64"; else RID="linux-arm64"; fi ;;
            *) echo "Unsupported architecture: $ARCH"; exit 1 ;;
        esac
        ;;
    *)
        echo "Unsupported OS: $OS"
        exit 1
        ;;
esac

# Build
echo "Building officecli ($RID)..."
TMPDIR=$(mktemp -d)
dotnet publish "$PROJECT" -c Release -r "$RID" -o "$TMPDIR" --nologo -v quiet
echo "Build complete."

# Install
EXISTING=$(command -v "$BINARY_NAME" 2>/dev/null || true)
if [ -n "$EXISTING" ]; then
    INSTALL_DIR=$(dirname "$EXISTING")
    echo "Found existing installation at $EXISTING, upgrading..."
else
    INSTALL_DIR="$HOME/.local/bin"
fi

mkdir -p "$INSTALL_DIR"
# Atomic replace: stage as .new alongside the target, sign there, then rename.
# Overwriting the binary in place would trash the text segment of any
# running officecli process (macOS does not block ETXTBSY), leaving it
# stuck in uninterruptible `UE` state on the next code page fault.
cp "$TMPDIR/$BINARY_NAME" "$INSTALL_DIR/$BINARY_NAME.new"
chmod +x "$INSTALL_DIR/$BINARY_NAME.new"
rm -rf "$TMPDIR"

# macOS: remove quarantine flag and ad-hoc codesign (required by AppleSystemPolicy)
# Done on the staged .new copy so the live binary is never mutated in place.
if [ "$(uname -s)" = "Darwin" ]; then
    xattr -d com.apple.quarantine "$INSTALL_DIR/$BINARY_NAME.new" 2>/dev/null || true
    codesign -s - -f "$INSTALL_DIR/$BINARY_NAME.new" 2>/dev/null || true
fi

mv -f "$INSTALL_DIR/$BINARY_NAME.new" "$INSTALL_DIR/$BINARY_NAME"

# Hint if not in PATH
case ":$PATH:" in
    *":$INSTALL_DIR:"*) ;;
    *) echo "Add to PATH: export PATH=\"$INSTALL_DIR:\$PATH\""
       echo "Or add the line above to your ~/.zshrc or ~/.bashrc" ;;
esac

echo "OfficeCli installed successfully!"
echo "Run 'officecli --help' to get started."
