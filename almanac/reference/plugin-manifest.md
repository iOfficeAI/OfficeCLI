---
title: "Plugin Manifest"
summary: "Exact reference for OfficeCLI plugin manifest fields, kind names, target formats, idle-timeout rules, and format-handler vocabulary."
topics: [reference, plugins]
sources:
  - id: plugin-protocol
    type: file
    path: plugins/plugin-protocol.md
  - id: plugin-manifest
    type: file
    path: src/officecli/Core/Plugins/PluginManifest.cs
  - id: plugin-registry
    type: file
    path: src/officecli/Core/Plugins/PluginRegistry.cs
---

A plugin manifest is the JSON object printed by a plugin in response to `<plugin> --info`. OfficeCLI uses it to identify plugin kind, supported extensions, protocol compatibility, target native format, idle-timeout budgets, diagnostics metadata, and format-handler vocabulary [@plugin-protocol] [@plugin-manifest]. Discovery and probe behavior are covered by [plugin discovery](plugin-discovery).

## Protocol gate

The supported plugin protocol major version is `1` [@plugin-registry]. The registry rejects any parsed manifest whose `protocol` value is not `1`, writes a warning to stderr, and treats the candidate as unresolved [@plugin-registry].

## Kind values

| Wire value | Runtime enum | Role |
|---|---|---|
| `dump-reader` | `PluginKind.DumpReader` | Reads a foreign format and emits OfficeCLI batch commands for a native target [@plugin-protocol] [@plugin-manifest] |
| `exporter` | `PluginKind.Exporter` | Converts a native Office file to a foreign output file [@plugin-protocol] [@plugin-manifest] |
| `format-handler` | `PluginKind.FormatHandler` | Owns a foreign file through a long-lived handler session [@plugin-protocol] [@plugin-manifest] |

The parser accepts only these three wire strings as known plugin kinds [@plugin-manifest]. The protocol reserves `engine` and `transformer`, but v1 plugins must not declare them [@plugin-protocol].

## Manifest fields

| Field | Type | Required by protocol | Runtime behavior |
|---|---|---:|---|
| `name` | string | Yes | Stable plugin identifier stored as `Name` [@plugin-protocol] [@plugin-manifest] |
| `version` | string | Yes | Plugin version string stored as `Version` [@plugin-protocol] [@plugin-manifest] |
| `protocol` | integer | Yes | Must equal `1` to pass registry probing [@plugin-protocol] [@plugin-registry] |
| `kinds` | string array | Yes | Must include the requested kind to match a lookup [@plugin-registry] |
| `extensions` | string array | Yes | Entries include leading dots, and matching is case-insensitive after normalization [@plugin-protocol] [@plugin-registry] |
| `idle_timeout_seconds` | object | Yes | Missing or invalid default produces a warning and runtime fallback to 60 seconds [@plugin-protocol] [@plugin-manifest] |
| `runtime` | string | Yes | Diagnostic tag only; main does not branch on it [@plugin-protocol] [@plugin-manifest] |
| `target` | string | Required for `dump-reader` | Must resolve to `docx`, `xlsx`, or `pptx`; omitted values default to `docx` in runtime resolution [@plugin-protocol] [@plugin-manifest] |
| `vocabulary` | object | Required for `format-handler` | Missing vocabulary produces a warning; runtime handshake vocabulary can still be used [@plugin-protocol] [@plugin-manifest] |
| `description` | string | No | Human-readable description [@plugin-protocol] [@plugin-manifest] |
| `tier` | string | No | Free-form tier label [@plugin-protocol] [@plugin-manifest] |
| `supports` | string array | No | Capability tags for diagnostics and listing [@plugin-protocol] [@plugin-manifest] |
| `limits` | object | No | Plugin-defined limit metadata [@plugin-protocol] [@plugin-manifest] |
| `homepage` | string | No | Plugin homepage URL [@plugin-protocol] [@plugin-manifest] |
| `license` | string | No | License identifier [@plugin-protocol] [@plugin-manifest] |

## Target format

`target` names the native format produced by a `dump-reader`. Runtime resolution accepts only `docx`, `xlsx`, and `pptx` [@plugin-manifest]. If the field is omitted, `ResolveTargetFormat` falls back to `docx`; if it names any other format, runtime resolution throws an error naming the supported values [@plugin-manifest].

## Idle timeout

`idle_timeout_seconds` contains a mandatory positive `default` and optional positive per-verb overrides under `verbs` [@plugin-protocol]. `PluginIdleTimeout.For` uses the verb override when it is present and greater than zero, otherwise it uses the default, and finally a 60-second safe default if the manifest object is missing or invalid [@plugin-manifest].

Users can override all manifest budgets with `OFFICECLI_PLUGIN_IDLE_TIMEOUT_SECONDS=<n>` [@plugin-protocol]. The runtime accepts `0` in this environment variable to disable the watchdog for the invocation, but `0` is not allowed inside the manifest itself [@plugin-protocol] [@plugin-manifest].

## Format-handler vocabulary

`vocabulary` has three fields: `addable_types`, `settable_props`, and `path_segments` [@plugin-protocol] [@plugin-manifest].

| Field | Type | Meaning |
|---|---|---|
| `addable_types` | string array | Type names the plugin exposes for add operations [@plugin-protocol] |
| `settable_props` | object mapping type to string array | Property names accepted for each type [@plugin-protocol] [@plugin-manifest] |
| `path_segments` | string array | Document path patterns exposed by the plugin model [@plugin-protocol] [@plugin-manifest] |

Manifest vocabulary is used for discovery and help output, while the format-handler open handshake returns a runtime vocabulary snapshot that the host trusts for the session [@plugin-protocol] [@plugin-manifest].

## Warnings

`PluginManifestExtensions.Warnings` reports soft problems without failing the manifest probe. It warns for missing or nonpositive idle-timeout defaults, empty `kinds`, unknown kind strings, unsupported explicit dump-reader targets, and missing format-handler vocabulary [@plugin-manifest].
