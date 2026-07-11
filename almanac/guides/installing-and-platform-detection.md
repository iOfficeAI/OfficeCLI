---
title: "Installing And Platform Detection"
summary: "Installers choose the correct OfficeCLI asset, verify it when possible, place it in the user install location, and optionally install agent skills or MCP fallback configuration."
topics: [guides, install]
sources:
  - id: unix-installer
    type: file
    path: install.sh
  - id: windows-installer
    type: file
    path: install.ps1
  - id: npm-installer
    type: file
    path: npm/lib/install-binary.js
  - id: cli-installer
    type: file
    path: src/officecli/Core/Installer.cs
---

Installing OfficeCLI means selecting the right platform asset, proving the download when a checksum manifest is available, and putting the binary where future commands can find it. The shell installer covers macOS and Linux, the PowerShell installer covers Windows, npm postinstall vendors a binary inside the package, and the built-in `officecli install` command can copy a running self-contained binary into the user install location [@unix-installer] [@windows-installer] [@npm-installer] [@cli-installer].

## Pick The Asset

On macOS and Linux, `install.sh` derives the asset from `uname -s` and `uname -m`; Linux also checks `ldd --version` and `/etc/alpine-release` to distinguish glibc from musl, producing Alpine asset names for musl systems [@unix-installer]. On Windows, `install.ps1` currently targets `officecli-win-x64.exe` for the script download path [@windows-installer]. The npm installer performs the same selection in Node: `darwin`, `linux`, and `win32` map with `x64` or `arm64`, and Linux calls `isMusl()` before choosing a glibc or Alpine asset [@npm-installer].

If you are changing asset names, update all three surfaces together. The platform names in the install scripts must match the names emitted by the [self-contained binary build](../architecture/build/self-contained-binary-and-embedded-resources) and uploaded by the release workflow.

## Download And Verify

The shell and PowerShell installers try `https://d.officecli.ai` first and fall back to GitHub releases [@unix-installer] [@windows-installer]. Both scripts first resolve the latest tag and prefer `/releases/download/vX.Y.Z/...` over `/releases/latest/download/...` so a new release is not mixed with cached latest assets [@unix-installer] [@windows-installer]. If `SHA256SUMS` is available, both scripts match the exact filename column and compare the expected hash with the downloaded file's SHA-256 hash [@unix-installer] [@windows-installer].

The npm installer pins downloads to the package version. It derives tag `v<version>` from `npm/package.json`, strips prerelease or build suffixes for the binary tag, downloads mirror-first then GitHub, and verifies `SHA256SUMS` with exact filename matching [@npm-installer].

## Place The Binary

For a script install, an existing `officecli` on `PATH` decides the upgrade directory; otherwise Unix installs to `~/.local/bin` and Windows installs to `%LOCALAPPDATA%\OfficeCLI` [@unix-installer] [@windows-installer]. The Unix installer stages the binary as `officecli.new`, applies executable permissions, handles macOS quarantine/signing, and renames it into place [@unix-installer]. The PowerShell script copies to `%LOCALAPPDATA%\OfficeCLI\officecli.exe` and updates the user PATH when needed [@windows-installer].

The built-in installer uses the same canonical install locations and skips copying when the process is already running from the target path, from Homebrew-managed paths, or from a small framework-dependent development build [@cli-installer]. It can also install agent skills and MCP fallback configuration; specific MCP-only targets such as VS Code and LM Studio skip the skill phase [@cli-installer].

## Recovery Checks

After installer changes, verify the selected asset on each affected platform, then check that `officecli --version` runs from the final location. For npm, remove `npm/vendor/officecli` or `officecli.exe` before a local postinstall test so `ensureBinary()` cannot return early from an existing file [@npm-installer].
