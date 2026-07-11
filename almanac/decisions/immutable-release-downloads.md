---
title: "Immutable Release Downloads"
summary: "Installers resolve or derive a concrete release tag and download versioned assets so a cached latest URL cannot install an older binary."
topics: [decisions, release, install]
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
---

OfficeCLI installers prefer immutable release downloads: they turn "latest" into a concrete tag, or derive the tag from the npm package version, then fetch assets from `/releases/download/vX.Y.Z/...`. This choice exists because mutable latest URLs can lag behind a just-published release, even when the downloaded checksum manifest is self-consistent; future installer work must preserve tag-pinned downloads before it changes mirrors, checksums, or package publishing [@unix-installer] [@windows-installer] [@npm-installer].

## Context

The shell and PowerShell installers still need to discover the current release for users who run a generic install command. They do that by following the `/releases/latest` redirect on the OfficeCLI mirror first, then GitHub, and extracting the final `vX.Y.Z` tag from the resolved URL [@unix-installer] [@windows-installer]. After that resolution step, both scripts build mirror and GitHub asset bases with `/releases/download/$VERSION` instead of downloading directly from `/releases/latest/download` [@unix-installer] [@windows-installer].

The npm installer has a different source of truth. Because npm publication sets `npm/package.json` to the release version, postinstall derives `TAG` from the package version and strips prerelease or build suffixes before forming versioned mirror and GitHub URLs [@npm-installer]. That makes npm install depend on the package version rather than a mutable latest redirect.

This decision belongs with [installing and platform detection](../guides/installing-and-platform-detection) and the [release build and checksum flow](../guides/release-build-and-checksum-flow), because asset names, release tags, and checksum manifests are one install contract.

## Decision

Installers must download release assets from immutable versioned paths whenever the release tag is known. The Unix and Windows scripts may fall back to latest paths only when tag resolution fails [@unix-installer] [@windows-installer]. The npm installer does not use latest paths; it downloads from the tag derived from the package version [@npm-installer].

The checksum lookup follows the same pinned asset base. The script installers download `SHA256SUMS` beside the selected asset and match the exact filename column before comparing hashes [@unix-installer] [@windows-installer]. The npm installer also fetches `SHA256SUMS` from the versioned tag and matches the exact asset name, allowing an optional leading binary-mode marker in the manifest row [@npm-installer].

## Status

This decision is active. `install.sh`, `install.ps1`, and `npm/lib/install-binary.js` all document the reason in code comments and implement versioned downloads with the OfficeCLI mirror as the first source and GitHub releases as fallback [@unix-installer] [@windows-installer] [@npm-installer].

## Consequences

The main benefit is release freshness. A user who installs immediately after a release should receive the asset for the resolved tag, not a stale object served from a cached latest URL [@unix-installer] [@windows-installer]. It also makes npm reproducible by tying the downloaded binary to the package version that triggered postinstall [@npm-installer].

The cost is that installer code must keep tag resolution healthy. If the generic script installers cannot resolve a tag, they deliberately fall back to latest paths so installation can still proceed, but that is the degraded path and carries the stale-download risk this decision avoids [@unix-installer] [@windows-installer].

Future changes must not reintroduce direct latest downloads as the normal path. Mirror rewrites, checksum changes, and package publishing changes should keep asset and checksum URLs on the same immutable release tag.
