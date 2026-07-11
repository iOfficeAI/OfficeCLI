---
title: "Release Build And Checksum Flow"
summary: "A release builds every platform asset, signs and smoke-tests the runnable binaries, creates SHA256SUMS, and publishes npm only after the GitHub release is public."
topics: [guides, release, build]
sources:
  - id: build-script
    type: file
    path: build.sh
  - id: build-workflow
    type: file
    path: .github/workflows/build.yml
  - id: npm-workflow
    type: file
    path: .github/workflows/publish-npm.yml
  - id: npm-package
    type: file
    path: npm/package.json
---

The release flow turns the [self-contained binary](../architecture/build/self-contained-binary-and-embedded-resources) into named assets, verifies that representative binaries can create and mutate Office files, generates checksums, and only then publishes npm packages that download those public assets. The key invariant is order: npm publication happens after a GitHub Release is published, because postinstall needs the platform binaries and `SHA256SUMS` to already be reachable [@npm-workflow].

## Build The Assets

For local work, `./build.sh release` builds the current platform, `./build.sh debug` builds the current platform in Debug, and `./build.sh all` builds all release targets [@build-script]. The script maps runtime identifiers to asset names for macOS, Linux glibc, Linux musl, and Windows, then writes outputs under `bin/release` or `bin/debug` [@build-script].

CI uses the same asset set in a matrix. The build workflow publishes `osx-arm64`, `osx-x64`, `linux-x64`, `linux-arm64`, `linux-musl-x64`, `linux-musl-arm64`, `win-x64`, and `win-arm64`, then renames each publish output to the release asset name [@build-workflow].

## Sign And Smoke Test

macOS assets are Developer ID signed with hardened runtime and an entitlement file that keeps self-contained CoreCLR JIT execution working; the workflow then verifies the signature and checks that the `allow-jit` entitlement was embedded [@build-workflow]. macOS binaries are submitted to Apple notarization as zipped bare binaries because a bare Mach-O cannot be stapled [@build-workflow].

The workflow smoke-tests runnable host/asset combinations by creating a `.docx`, adding a paragraph, reading it back, and closing the file [@build-workflow]. It also runs a Linux `dotnet/runtime:8.0` container smoke test for the `linux-x64` self-contained binary and tests both the built-in `install` command and the public `install.sh` or `install.ps1` script on supported runners [@build-workflow].

## Publish Checksums And npm

On tag builds, the release job downloads all artifacts, flattens them into one directory, and runs `sha256sum officecli-* > SHA256SUMS` before creating a draft GitHub Release [@build-workflow]. The draft step matters: a human can inspect assets and checksums before publishing.

The npm workflow runs when a release is published or manually dispatched [@npm-workflow]. It sets the package name and version for both `@officecli/officecli` and `@aionui/officecli`, rewrites README package commands, and publishes with npm trusted publishing and provenance when the package already exists [@npm-workflow]. The npm package exposes an `officecli` bin shim and runs `node install.js` on postinstall [@npm-package].

## Maintainer Check

Before publishing npm, confirm that the GitHub Release is public and includes every asset plus `SHA256SUMS`. That check protects immediate `npm install` users, because npm postinstall downloads from the release tag rather than building a binary locally [@npm-workflow].
