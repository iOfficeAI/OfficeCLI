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
        case "$ARCH" in
            x86_64) ASSET="officecli-linux-x64" ;;
            aarch64|arm64) ASSET="officecli-linux-arm64" ;;
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
echo "Downloading OfficeCli ($ASSET)..."
if curl -fsSL "$DOWNLOAD_URL" -o "/tmp/$BINARY_NAME" 2>/dev/null; then
    chmod +x "/tmp/$BINARY_NAME"
    if "/tmp/$BINARY_NAME" --version >/dev/null 2>&1; then
        SOURCE="/tmp/$BINARY_NAME"
        echo "Download verified."
    else
        echo "Downloaded file is not a valid OfficeCli binary."
        rm -f "/tmp/$BINARY_NAME"
    fi
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
cp "$SOURCE" "$INSTALL_DIR/$BINARY_NAME"
chmod +x "$INSTALL_DIR/$BINARY_NAME"

# macOS: remove quarantine flag and ad-hoc codesign (required by AppleSystemPolicy)
if [ "$(uname -s)" = "Darwin" ]; then
    xattr -d com.apple.quarantine "$INSTALL_DIR/$BINARY_NAME" 2>/dev/null || true
    codesign -s - -f "$INSTALL_DIR/$BINARY_NAME" 2>/dev/null || true
fi

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

# Step 4: Install AI agent skills for detected tools
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

echo "OfficeCli installed successfully!"
echo "Run 'officecli --help' to get started."
