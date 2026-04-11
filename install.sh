#!/bin/bash
set -e

REPO="iOfficeAI/OfficeCli"
BINARY_NAME="officecli"

# Detect platform
OS=$(uname -s | tr '[:upper:]' '[:lower:]')
ARCH=$(uname -m)

case "$OS" in
    darwin)
        case "$ARCH" in
            arm64) ASSET="officecli-mac-arm64" ;;
            x86_64) ASSET="officecli-mac-x64" ;;
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
                if [ "$LIBC" = "musl" ]; then
                    ASSET="officecli-linux-alpine-x64"
                else
                    ASSET="officecli-linux-x64"
                fi
                ;;
            aarch64|arm64)
                if [ "$LIBC" = "musl" ]; then
                    ASSET="officecli-linux-alpine-arm64"
                else
                    ASSET="officecli-linux-arm64"
                fi
                ;;
            *) echo "Unsupported architecture: $ARCH"; exit 1 ;;
        esac
        ;;
    *)
        echo "Unsupported OS: $OS"
        echo "For Windows, download from: https://github.com/$REPO/releases"
        exit 1
        ;;
esac

SOURCE=""

# Step 1: Try downloading from GitHub
DOWNLOAD_URL="https://github.com/$REPO/releases/latest/download/$ASSET"
CHECKSUM_URL="https://github.com/$REPO/releases/latest/download/SHA256SUMS"
echo "Downloading OfficeCli ($ASSET)..."
if curl -fsSL "$DOWNLOAD_URL" -o "/tmp/$BINARY_NAME" 2>/dev/null; then
    # Verify checksum if available
    CHECKSUM_OK=false
    if curl -fsSL "$CHECKSUM_URL" -o "/tmp/officecli-SHA256SUMS" 2>/dev/null; then
        EXPECTED=$(grep "$ASSET" "/tmp/officecli-SHA256SUMS" | awk '{print $1}')
        if [ -n "$EXPECTED" ]; then
            if command -v sha256sum >/dev/null 2>&1; then
                ACTUAL=$(sha256sum "/tmp/$BINARY_NAME" | awk '{print $1}')
            else
                ACTUAL=$(shasum -a 256 "/tmp/$BINARY_NAME" | awk '{print $1}')
            fi
            if [ "$EXPECTED" = "$ACTUAL" ]; then
                CHECKSUM_OK=true
                echo "Checksum verified."
            else
                echo "Checksum mismatch! Expected: $EXPECTED, Got: $ACTUAL"
                rm -f "/tmp/$BINARY_NAME" "/tmp/officecli-SHA256SUMS"
                exit 1
            fi
        fi
        rm -f "/tmp/officecli-SHA256SUMS"
    fi
    if [ "$CHECKSUM_OK" = false ]; then
        echo "Checksum file not available, skipping verification."
    fi
    chmod +x "/tmp/$BINARY_NAME"
    SOURCE="/tmp/$BINARY_NAME"
else
    echo "Download failed."
fi

# Step 2: Fallback to local files
if [ -z "$SOURCE" ]; then
    echo "Looking for local binary..."
    for candidate in "./$ASSET" "./$BINARY_NAME" "./bin/$ASSET" "./bin/$BINARY_NAME" "./bin/release/$ASSET" "./bin/release/$BINARY_NAME"; do
        if [ -f "$candidate" ]; then
            if [ ! -x "$candidate" ]; then
                chmod +x "$candidate"
            fi
            if "$candidate" --version >/dev/null 2>&1; then
                SOURCE="$candidate"
                echo "Found valid binary at $candidate"
                break
            fi
        fi
    done
fi

if [ -z "$SOURCE" ]; then
    echo "Error: Could not find a valid OfficeCli binary."
    echo "Download manually from: https://github.com/$REPO/releases"
    exit 1
fi

# Step 3: Install
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
cp "$SOURCE" "$INSTALL_DIR/$BINARY_NAME.new"
chmod +x "$INSTALL_DIR/$BINARY_NAME.new"

# macOS: remove quarantine flag and ad-hoc codesign (required by AppleSystemPolicy)
# Done on the staged .new copy so the live binary is never mutated in place.
if [ "$(uname -s)" = "Darwin" ]; then
    xattr -d com.apple.quarantine "$INSTALL_DIR/$BINARY_NAME.new" 2>/dev/null || true
    codesign -s - -f "$INSTALL_DIR/$BINARY_NAME.new" 2>/dev/null || true
fi

mv -f "$INSTALL_DIR/$BINARY_NAME.new" "$INSTALL_DIR/$BINARY_NAME"

# Auto-add to PATH if needed
case ":$PATH:" in
    *":$INSTALL_DIR:"*) ;;
    *)
        PATH_LINE="export PATH=\"$INSTALL_DIR:\$PATH\""
        if [ "$(uname -s)" = "Darwin" ]; then
            SHELL_RC="$HOME/.zshrc"
        elif [ -n "$ZSH_VERSION" ]; then
            SHELL_RC="$HOME/.zshrc"
        else
            SHELL_RC="$HOME/.bashrc"
        fi
        if ! grep -qF "$INSTALL_DIR" "$SHELL_RC" 2>/dev/null; then
            echo "" >> "$SHELL_RC"
            echo "$PATH_LINE" >> "$SHELL_RC"
            echo "Added $INSTALL_DIR to PATH in $SHELL_RC"
            echo "Run 'source $SHELL_RC' or restart your terminal to apply."
        fi
        ;;
esac

rm -f "/tmp/$BINARY_NAME"

# Step 4: Install AI agent skills (first install only)
SKILL_MARKER="$INSTALL_DIR/.officecli-skills-installed"
if [ ! -f "$SKILL_MARKER" ]; then
    SKILL_TARGETS=""
    for tool_dir in "$HOME/.claude:Claude Code" "$HOME/.copilot:GitHub Copilot" "$HOME/.agents:Codex CLI" "$HOME/.cursor:Cursor" "$HOME/.windsurf:Windsurf" "$HOME/.minimax:MiniMax CLI" "$HOME/.openclaw:OpenClaw" "$HOME/.nanobot/workspace:NanoBot" "$HOME/.zeroclaw/workspace:ZeroClaw"; do
        dir="${tool_dir%%:*}"
        name="${tool_dir##*:}"
        if [ -d "$dir" ]; then
            SKILL_TARGETS="$SKILL_TARGETS $dir/skills/officecli"
            echo "$name detected."
        fi
    done

    if [ -n "$SKILL_TARGETS" ]; then
        echo "Downloading officecli skill..."
        if curl -fsSL "https://raw.githubusercontent.com/$REPO/main/SKILL.md" -o "/tmp/officecli-skill.md" 2>/dev/null; then
            for target in $SKILL_TARGETS; do
                mkdir -p "$target"
                cp "/tmp/officecli-skill.md" "$target/SKILL.md"
                echo "  Installed: $target/SKILL.md"
            done
            rm -f "/tmp/officecli-skill.md"
        fi
    fi
    touch "$SKILL_MARKER"
fi

echo "OfficeCli installed successfully!"
echo "Run 'officecli --help' to get started."
