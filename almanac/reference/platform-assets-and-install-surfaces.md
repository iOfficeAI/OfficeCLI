---
title: "Platform Assets And Install Surfaces"
summary: "Lookup reference for OfficeCLI platform asset names, installer locations, npm binary vendoring, and install-time side effects."
topics: [reference, install]
sources:
  - id: readme
    type: file
    path: README.md
  - id: unix-installer
    type: file
    path: install.sh
  - id: windows-installer
    type: file
    path: install.ps1
  - id: npm-package
    type: file
    path: npm/package.json
  - id: npm-installer
    type: file
    path: npm/lib/install-binary.js
---

OfficeCLI ships as a self-contained native binary selected by platform and installed through shell, PowerShell, npm, package-manager, manual-download, or self-install surfaces [@readme]. This reference lists the asset names, download rules, install locations, npm behavior, and skill side effects that maintainers must keep aligned with the [installing and platform detection](../guides/installing-and-platform-detection) and [release build and checksum flow](../guides/release-build-and-checksum-flow) guides.

## Platform assets

| Platform | Asset name |
|---|---|
| macOS arm64 | `officecli-mac-arm64` [@readme] |
| macOS x64 | `officecli-mac-x64` [@readme] |
| Linux x64 glibc | `officecli-linux-x64` [@unix-installer] |
| Linux arm64 glibc | `officecli-linux-arm64` [@unix-installer] |
| Linux x64 musl or Alpine | `officecli-linux-alpine-x64` [@unix-installer] |
| Linux arm64 musl or Alpine | `officecli-linux-alpine-arm64` [@unix-installer] |
| Windows x64 | `officecli-win-x64.exe` [@readme] |
| Windows arm64 | `officecli-win-arm64.exe` [@readme] [@npm-installer] |

`install.sh` detects macOS and Linux with `uname`, maps `x86_64` to x64, maps `aarch64` and `arm64` to arm64, and treats Linux as musl when `ldd --version` mentions musl or `/etc/alpine-release` exists [@unix-installer]. The npm installer repeats the same platform split using Node `process.platform`, `process.arch`, `process.report`, `/etc/alpine-release`, and `ldd --version` [@npm-installer]. The PowerShell installer download path currently targets `officecli-win-x64.exe` [@windows-installer].

## Download surfaces

| Surface | Selection and download behavior |
|---|---|
| `install.sh` | Resolves the latest release tag, downloads mirror-first from `https://d.officecli.ai`, falls back to GitHub releases, and prefers immutable `/releases/download/<tag>/...` URLs [@unix-installer] |
| `install.ps1` | Uses the same mirror-first and GitHub-fallback pattern, resolves a `vX.Y.Z` tag, and falls back to the mutable latest path only when tag resolution fails [@windows-installer] |
| npm postinstall | Derives the release tag from `npm/package.json` version, strips prerelease and build suffixes, downloads mirror-first then GitHub, and writes the native binary under `npm/vendor/` [@npm-installer] |
| Manual download | README lists GitHub Releases as the manual source for platform binaries [@readme] |

The shell, PowerShell, and npm installers all try to verify `SHA256SUMS` when available and match the exact filename column before comparing SHA-256 hashes [@unix-installer] [@windows-installer] [@npm-installer].

## Install locations

| Surface | Final location |
|---|---|
| Unix script with existing `officecli` | Directory containing the existing command on `PATH` [@unix-installer] |
| Unix script without existing command | `~/.local/bin/officecli` [@unix-installer] |
| Windows script with existing `officecli.exe` | Directory containing the existing command [@windows-installer] |
| Windows script without existing command | `%LOCALAPPDATA%\OfficeCLI\officecli.exe` [@windows-installer] |
| npm package | `npm/vendor/officecli` or `npm/vendor/officecli.exe`, reached through package `bin` shim `officecli.js` [@npm-package] [@npm-installer] |

The Unix installer stages the binary as `officecli.new`, applies executable permissions, handles macOS quarantine and signing on the staged copy, and then renames it into place [@unix-installer]. The Windows installer copies the downloaded binary to `officecli.exe` and adds the install directory to the user's `Path` if needed [@windows-installer].

## Package-manager surfaces

The README lists Homebrew, Scoop, npm, manual download, one-line shell or PowerShell install, and `officecli install` self-install as supported human install paths [@readme]. The npm package is named `@officecli/officecli`, supports `darwin`, `linux`, and `win32` on `x64` and `arm64`, requires Node `>=14`, and runs `node install.js` during `postinstall` [@npm-package].

## Skill side effects

On first script install, Unix and Windows installers detect supported AI-agent directories and download the umbrella `SKILL.md` into each detected `skills/officecli/SKILL.md` location [@unix-installer] [@windows-installer]. They create a `.officecli-skills-installed` marker in the install directory so that script path does not repeat the first-install skill download on every upgrade [@unix-installer] [@windows-installer].

The README also states that `officecli install` copies the binary to `PATH` and installs the OfficeCLI skill into detected AI coding agents [@readme].
