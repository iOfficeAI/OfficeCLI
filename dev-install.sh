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
        case "$ARCH" in
            x86_64) RID="linux-x64" ;;
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
cp "$TMPDIR/$BINARY_NAME" "$INSTALL_DIR/$BINARY_NAME"
chmod +x "$INSTALL_DIR/$BINARY_NAME"
rm -rf "$TMPDIR"

# Hint if not in PATH
case ":$PATH:" in
    *":$INSTALL_DIR:"*) ;;
    *) echo "Add to PATH: export PATH=\"$INSTALL_DIR:\$PATH\""
       echo "Or add the line above to your ~/.zshrc or ~/.bashrc" ;;
esac

echo "OfficeCli installed successfully!"
echo "Run 'officecli --help' to get started."
