#!/bin/bash
set -e

REPO="iOfficeAI/OfficeCLI"
BINARY_NAME="officecli"

# Mirror primary, github fallback. The mirror is exercised first so issues
# surface there fast; github is the final safety net.
MIRROR_BASE="https://d.officecli.ai"
GITHUB_RELEASE_BASE="https://github.com/$REPO/releases/latest/download"
GITHUB_RAW_BASE="https://raw.githubusercontent.com/$REPO/main"

# fetch_with_fallback <primary_url> <fallback_url> <output_path>
# Returns 0 if either source delivered the file, non-zero if both failed.
# Short connect-timeout on primary so a dead mirror doesn't add minutes
# of stall before falling through.
fetch_with_fallback() {
    local primary="$1" fallback="$2" out="$3"
    if curl -fsSL --max-time 300 --connect-timeout 5 "$primary" -o "$out" 2>/dev/null; then
        echo "  (via mirror)"
        return 0
    fi
    echo "  mirror unreachable, falling back to github..."
    curl -fsSL --max-time 300 "$fallback" -o "$out" 2>/dev/null
}

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

# Step 1: Try downloading (mirror first, github fallback)
echo "Downloading OfficeCLI ($ASSET)..."
if fetch_with_fallback \
        "$MIRROR_BASE/releases/latest/download/$ASSET" \
        "$GITHUB_RELEASE_BASE/$ASSET" \
        "/tmp/$BINARY_NAME"; then
    # Verify checksum if available
    CHECKSUM_OK=false
    if fetch_with_fallback \
            "$MIRROR_BASE/releases/latest/download/SHA256SUMS" \
            "$GITHUB_RELEASE_BASE/SHA256SUMS" \
            "/tmp/officecli-SHA256SUMS"; then
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
    echo "Error: Could not find a valid OfficeCLI binary."
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

# macOS: clear the quarantine flag, then ensure the staged copy carries a
# valid signature (Apple Silicon refuses to exec an unsigned Mach-O).
# Release binaries are Developer ID signed + notarized by CI; a forced
# ad-hoc re-sign would strip that signature and invalidate notarization,
# so only ad-hoc sign as a fallback when no valid signature is present.
# Done on the staged .new copy so the live binary is never mutated in place.
if [ "$(uname -s)" = "Darwin" ]; then
    xattr -d com.apple.quarantine "$INSTALL_DIR/$BINARY_NAME.new" 2>/dev/null || true
    if ! codesign -v --strict "$INSTALL_DIR/$BINARY_NAME.new" 2>/dev/null; then
        # The JIT entitlement is required for CoreCLR to allocate executable
        # memory under hardened runtime; without it the process fails with
        # "Failed to create CoreCLR, HRESULT: 0x80070008".
        ENTITLEMENTS_URL="https://raw.githubusercontent.com/$REPO/main/officecli.entitlements"
        ENTITLEMENTS_FILE="/tmp/officecli.entitlements"
        if curl -fsSL --max-time 10 "$ENTITLEMENTS_URL" -o "$ENTITLEMENTS_FILE" 2>/dev/null; then
            codesign -s - -f --entitlements "$ENTITLEMENTS_FILE" "$INSTALL_DIR/$BINARY_NAME.new" 2>/dev/null || true
            rm -f "$ENTITLEMENTS_FILE"
        else
            codesign -s - -f "$INSTALL_DIR/$BINARY_NAME.new" 2>/dev/null || true
        fi
    fi
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
    for tool_dir in "$HOME/.claude:Claude Code" "$HOME/.copilot:GitHub Copilot" "$HOME/.agents:Codex CLI" "$HOME/.cursor:Cursor" "$HOME/.windsurf:Windsurf" "$HOME/.minimax:MiniMax CLI" "$HOME/.openclaw:OpenClaw" "$HOME/.nanobot/workspace:NanoBot" "$HOME/.zeroclaw/workspace:ZeroClaw" "$HOME/.hermes:Hermes Agent"; do
        dir="${tool_dir%%:*}"
        name="${tool_dir##*:}"
        if [ -d "$dir" ]; then
            SKILL_TARGETS="$SKILL_TARGETS $dir/skills/officecli"
            echo "$name detected."
        fi
    done

    if [ -n "$SKILL_TARGETS" ]; then
        echo "Downloading officecli skill..."
        if fetch_with_fallback \
                "$MIRROR_BASE/SKILL.md" \
                "$GITHUB_RAW_BASE/SKILL.md" \
                "/tmp/officecli-skill.md"; then
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

echo "OfficeCLI installed successfully!"
echo "Run 'officecli --help' to get started."
