---
title: "Plugin Discovery"
summary: "Exact lookup reference for how OfficeCLI resolves plugin executables, probes manifests, caches results, and hardens PATH lookup."
topics: [reference, plugins]
sources:
  - id: plugin-protocol
    type: file
    path: plugins/plugin-protocol.md
  - id: plugin-registry
    type: file
    path: src/officecli/Core/Plugins/PluginRegistry.cs
---

Plugin discovery is the fixed search process OfficeCLI uses when it needs a plugin for a `(kind, extension)` pair. The protocol defines four ordered locations, and `PluginRegistry.FindFor` implements that order, normalizes extensions, probes candidates with `--info`, applies the protocol-version gate, and caches both hits and misses for the current process [@plugin-protocol] [@plugin-registry]. For the broader extension model, see the [plugin system](../architecture/plugins/plugin-system).

## Lookup key

Discovery is keyed by plugin kind and file extension. The supported kind wire strings are `dump-reader`, `exporter`, and `format-handler` [@plugin-protocol]. The registry normalizes extensions to lowercase and adds a leading dot when needed, so `doc` and `.DOC` resolve as `.doc` [@plugin-registry].

## Search order

The first matching candidate wins [@plugin-protocol].

| Order | Surface | Candidate form |
|---:|---|---|
| 1 | Environment variable | `OFFICECLI_PLUGIN_<KIND>_<EXT>` with kind uppercased and `-` changed to `_`, such as `OFFICECLI_PLUGIN_DUMP_READER_DOC` [@plugin-registry] |
| 2 | User plugin directory | `~/.officecli/plugins/<kind>/<ext>/plugin` or `plugin.exe` on Windows [@plugin-registry] |
| 3 | Bundled plugin directory | `<app-base>/plugins/<kind>/<ext>/plugin` or `plugin.exe` on Windows [@plugin-registry] |
| 4 | `PATH` lookup | `officecli-<kind>-<ext>` first, then `officecli-<ext>` [@plugin-registry] |

The protocol uses extension names without the leading dot in paths and environment variable names, such as `doc`, `hwpx`, or `pdf` [@plugin-protocol]. Symlinks are allowed by the protocol, because the registry only tests whether the candidate file exists before probing it [@plugin-protocol] [@plugin-registry].

## Executable names

Directory-based plugins use `plugin` on Unix-like systems. On Windows, the registry checks `plugin.exe` and then `plugin` [@plugin-registry].

`PATH` plugins use two name variants. The specific form is `officecli-<kind>-<ext>`, such as `officecli-dump-reader-doc`; the fallback form is `officecli-<ext>`, such as `officecli-doc` [@plugin-protocol] [@plugin-registry]. On Windows, `PATH` lookup checks both the `.exe` and extensionless forms [@plugin-registry].

## Probe and match

Every candidate must answer `<plugin> --info` with one JSON manifest on stdout and exit `0` [@plugin-protocol]. The registry gives this probe 5 seconds, reads stdout and stderr asynchronously to avoid pipe deadlock, rejects empty or malformed JSON, rejects nonzero exit, and treats any probe failure as "not found" for that candidate [@plugin-registry].

After parsing the manifest, the registry rejects any plugin whose `protocol` is not `1` [@plugin-registry]. A parsed manifest must also contain the requested kind and requested normalized extension before it can resolve the lookup [@plugin-registry]. Manifest fields are listed in the [plugin manifest](plugin-manifest) reference.

## Cache and enumeration

`FindFor` caches the result for each `(kind, extension)` pair for the lifetime of the process. Negative results are cached too, so a missing or invalid plugin is not re-probed on every operation [@plugin-registry]. `InvalidateCache` clears this map for install flows that need re-discovery without restarting [@plugin-registry].

`EnumerateAll` is narrower than targeted lookup. It walks only the user and bundled plugin directory trees, looking for the two-level layout `<root>/<kind>/<ext>/plugin(.exe)` [@plugin-registry]. It does not enumerate environment-variable plugins or scan arbitrary `PATH` names [@plugin-registry].

## PATH hardening

`PATH` lookup ignores relative entries such as `.` or `src/bin` [@plugin-registry]. On Unix-like systems, it also skips world-writable directories by checking the `OtherWrite` mode bit [@plugin-registry]. Windows returns `false` from that mode-bit check because Windows uses ACLs rather than Unix mode bits [@plugin-registry].
