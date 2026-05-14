# OfficeCLI HWPX Rebase Report

## Overview

This document records the 2026-05-14 rebase of the OfficeCLI `feat/hwpx`
branch onto `upstream/main` 1.0.91.

## Repository State

- Repository: `/Users/jun/Developer/new/700_projects/cli-jaw/officecli`
- Branch: `feat/hwpx`
- Upstream base: `upstream/main` 1.0.91
- PATH binary: `/Users/jun/.local/bin/officecli`
- Registered target: `/Users/jun/Developer/new/700_projects/cli-jaw/officecli/build-local/officecli`
- Published OfficeCLI version: `1.0.91.0`

## Rebase Result

The rebase completed successfully after manual conflict resolution. The resolved
branch preserves upstream plugin, PDF, exporter, and minimal-document support
while keeping the fork's native HWPX handler and experimental HWP bridge work.

## Conflict Resolution Notes

- `BlankDocCreator.cs`: kept upstream plugin creation support and native `.hwpx`
  creation before plugin fallback.
- `CommandBuilder.Import.cs`: kept upstream `--minimal` and fork HWPX
  `--from-markdown` / `--align` import behavior.
- `CommandBuilder.cs`: registered both `BuildPluginsCommand` and
  `BuildCompareCommand`.
- `DocumentHandlerFactory.cs`: routed `.hwpx` to `HwpxHandler`, preserved `.hwp`
  bridge guidance, and kept plugin fallback for unknown extensions.
- `CommandBuilder.View.cs` and `ResidentServer.cs`: kept upstream `pdf` and
  plugin forms support together with HWP/HWPX modes such as `forms`, `tables`,
  `markdown`, `objects`, `styles`, `fields`, and `field`.
- `WordHandler.Add.Text.cs`: preserved upstream `sym=font:hex` handling to avoid
  duplicate symbol glyph text on dump/batch round trips.

## Verification

- No tracked Rust `target` artifacts: passed.
- `dotnet build officecli.slnx`: passed with warnings only.
- `cargo build --manifest-path src/rhwp-field-bridge/Cargo.toml`: passed.
- HWP bridge focused tests: 36 passed, 0 failed.
- Full OfficeCLI test project: 234 passed, 0 failed.
- `build-local/officecli` republished for `osx-arm64`.
- PATH `officecli` resolves to the republished `build-local/officecli` target.

## HWP Capability Boundary

Binary `.hwp` creation remains unsupported by OfficeCLI capabilities. The current
safe document creation path is `.hwpx`; binary HWP mutation remains
operation-gated and requires the experimental rhwp bridge environment.
