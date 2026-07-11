---
title: "Atomic Binary Replacement"
summary: "Unix install and build flows stage OfficeCLI binaries beside the target and rename them into place instead of overwriting executable files in place."
topics: [decisions, install, build]
sources:
  - id: unix-installer
    type: file
    path: install.sh
  - id: dev-installer
    type: file
    path: dev-install.sh
  - id: build-script
    type: file
    path: build.sh
---

OfficeCLI's Unix installer, development installer, and local build script replace binaries atomically: they copy the new executable to a sibling `.new` file, apply permissions or signing to that staged file, then rename it over the target. The choice protects running `officecli` processes from in-place mutation and future install or build changes must keep live binaries immutable until the final rename [@unix-installer] [@dev-installer] [@build-script].

## Context

The installed CLI is a self-contained executable that may be running while an upgrade or local build writes a replacement. The Unix install script explicitly says that overwriting the binary in place can damage the text segment of a running process on macOS and leave it stuck on a later code-page fault [@unix-installer]. The same warning appears in the development installer and the build script, so the rule applies to local development, install upgrades, and release asset generation [@dev-installer] [@build-script].

This decision constrains both [installing and platform detection](../guides/installing-and-platform-detection) and the [self-contained binary build](../architecture/build/self-contained-binary-and-embedded-resources). The binary may be signed, made executable, or copied from a publish directory, but those operations happen before the target path is swapped into service.

## Decision

Unix replacement flows stage the new executable beside the final path with a `.new` suffix and then call `mv -f` to replace the target. `install.sh` copies the downloaded or local source to `$INSTALL_DIR/officecli.new`, marks it executable, performs macOS quarantine and signing work on the staged copy, and renames it to `$INSTALL_DIR/officecli` [@unix-installer].

`dev-install.sh` follows the same pattern after `dotnet publish`: it copies the published binary to `$INSTALL_DIR/officecli.new`, marks it executable, signs the staged copy on macOS, and only then renames it into place [@dev-installer]. `build.sh` uses the same staging rule for generated release or debug assets under `bin/release` or `bin/debug`; each target asset is copied as `<asset>.new`, optionally signed on macOS, and renamed to the final asset name [@build-script].

## Status

This decision is active for Unix install and build paths. The evidence is repeated in all three scripts, and each one performs the final replacement with a rename after preparation is complete [@unix-installer] [@dev-installer] [@build-script].

## Consequences

The benefit is that the live binary path is not mutated while preparation is still happening. Permissions, quarantine removal, and signing can fail or be retried against the staged file without partially changing the executable that users may currently be running [@unix-installer] [@dev-installer].

The required discipline is simple but important: future install or build code should prepare a sibling file on the same filesystem and rename it into place. Code that writes directly to the final executable path, signs the live binary in place, or streams a download into the installed path violates this decision.
