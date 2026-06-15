# AGENTS.md — OfficeCLI Fork (AionUI-optimized)

> **For AI agents** working with this repository. This is a **fork** of [iOfficeAI/OfficeCLI](https://github.com/iOfficeAI/OfficeCLI) optimized for [AionUI](https://github.com/iOfficeAI/AionUi) preview with enhanced Track Changes / revision mode support.

## Quick start (if user asked you to install this project)

```bash
# 1. Build from source
git clone https://github.com/NextDoorLaoHuang-HF/OfficeCLI.git && cd OfficeCLI
dotnet publish src/officecli/officecli.csproj -c Release -r osx-arm64 --self-contained -o out
# Other platforms: replace osx-arm64 with osx-x64, linux-x64, linux-arm64, win-x64, or win-arm64

# 2. Install to PATH
sudo cp out/officecli /usr/local/bin/   # macOS/Linux
# or: copy out/officecli.exe %LOCALAPPDATA%\OfficeCli\   (Windows)

# 3. Verify
officecli --version
```

## Install skills to AI agents

```bash
officecli install          # install base skill to all detected agents
officecli skills install word   # install docx-specific skill
```

## AionUI integration (Track Changes / 修订模式)

If the user uses **AionUI** and wants revision/track-changes capabilities:

```bash
officecli setup-aionui
```

This does two things:
1. Installs the `officecli-track-changes` skill into AionUI's skills directory
2. Registers the **Word 修订助手** assistant preset in AionUI's configuration

After running, tell the user to **restart AionUI** — the "Word 修订助手" will appear in the assistant list, pre-configured with both `officecli-docx` and `officecli-track-changes` skills.

The `setup-aionui` command is **idempotent** — safe to run multiple times.

### Track Changes workflow after setup

```bash
# Preferred: find + revision (one command)
officecli set "$FILE" /body --prop find="old" --prop replace="new" --prop revision.author=AI

# Verify
officecli query "$FILE" revision

# Accept/reject
officecli set "$FILE" /revision --prop revision.action=accept
```

## Key differences from upstream

| Feature | Upstream | This fork |
|---------|----------|-----------|
| Binary distribution | Prebuilt releases | Build from source |
| AionUI assistant | Word Creator (generic) | + Word 修订助手 (revision-focused) |
| HTML preview | Standard | Enhanced: 7 CSS classes for revision marks |
| Watch API | Standard | + /api/revision/accept\|reject\|count endpoints |
| `setup-aionui` command | None | ✅ One-shot AionUI integration |
| `officecli-track-changes` skill | None | ✅ Embedded in binary |

## Important files

- `SKILL.md` — base skill file (embedded in binary, installed by `officecli install`)
- `skills/officecli-docx/SKILL.md` — docx-specific skill
- `skills/officecli-track-changes/SKILL.md` — track changes / revision mode skill
- `src/officecli/Core/AionuiInstaller.cs` — `setup-aionui` implementation
- `src/officecli/Core/SkillInstaller.cs` — skill installation to AI agents
