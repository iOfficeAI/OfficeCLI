---
title: "Self Contained Binary And Embedded Resources"
summary: "OfficeCLI is published as a self-contained single-file .NET binary that carries its help schemas, skills, and preview resources inside the executable."
topics: [architecture, build]
sources:
  - id: csproj
    type: file
    path: src/officecli/officecli.csproj
  - id: build-script
    type: file
    path: build.sh
  - id: npm-installer
    type: file
    path: npm/lib/install-binary.js
  - id: skill-installer
    type: file
    path: src/officecli/Core/SkillInstaller.cs
---

OfficeCLI's release artifact is meant to behave like a complete product, not a thin launcher. The project file publishes an executable named `officecli` as a self-contained, single-file, trimmed .NET application, then embeds help schemas, agent skills, preview assets, chart resources, and PowerPoint effect templates into the assembly [@csproj]. That choice lets installers and npm downloads place one platform binary and still expose the command help, [help schemas](../../concepts/help-schemas), and [officecli skills](../../concepts/officecli-skills) needed by users and agents.

## Binary Boundary

The build boundary is the .NET publish output for `src/officecli/officecli.csproj`. The project sets `PublishSingleFile`, `SelfContained`, and `PublishTrimmed`, so each runtime identifier produces a native executable that includes the .NET runtime instead of requiring a machine-wide runtime install [@csproj]. The local build script maps runtime identifiers to release asset names such as `officecli-mac-arm64`, `officecli-linux-x64`, `officecli-linux-alpine-arm64`, and `officecli-win-x64.exe` [@build-script].

The same script builds into a temporary publish directory, copies the result to a staged `.new` file, optionally ad-hoc signs macOS outputs, and then renames the staged file into place [@build-script]. This keeps the live binary from being mutated in place while another `officecli` process may still be mapped into memory [@build-script].

## Embedded Contract

The embedded resources are part of the runtime contract. `preview.css`, `preview.js`, watch scripts, chart style XML, chart gallery XML, and PowerPoint effect templates are declared as embedded resources in the project file [@csproj]. The help schema tree under `schemas/help/**/*.json` is also embedded with logical names under `schemas/help/`, so schema help can be loaded from the assembly without extracting files beside the executable [@csproj].

Skills use the same packaging model. The project embeds `../../skills/**/*` under `skills/...` logical names and normalizes recursive directory separators so resource prefix lookups work across platforms [@csproj]. `SkillInstaller` then reads embedded skill files, builds the skill catalog, serves `load_skill` content, lists bundled reference files, and rejects binary skill assets on text-only channels [@skill-installer].

## Installer Consequence

Because the binary is self-contained, the npm package does not ship native code in the package tarball. Its postinstall logic detects the current platform, downloads the matching release asset into `vendor/`, verifies `SHA256SUMS` when available, and exposes the binary path to callers [@npm-installer]. This is why the [release build and checksum flow](../../guides/release-build-and-checksum-flow) treats asset names, checksums, and platform detection as one system: the downloaded file must already contain the resources needed after installation.
