---
title: "Plugin System"
summary: "OfficeCLI's plugin system discovers sidecar executables that add foreign-format reading, exporting, or full handler ownership without moving those implementations into the native handler code."
topics: [architecture, plugins]
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
  - id: plugin-process
    type: file
    path: src/officecli/Core/Plugins/PluginProcess.cs
  - id: plugin-command
    type: file
    path: src/officecli/CommandBuilder.Plugins.cs
  - id: handler-factory
    type: file
    path: src/officecli/Handlers/DocumentHandlerFactory.cs
---

OfficeCLI's plugin system extends the three native Office formats with sidecar executables. A plugin can migrate a foreign file into a native sibling, export a native file to a foreign target, or own a foreign file through the normal document-handler interface. The boundary matters because the main binary keeps native `.docx`, `.xlsx`, and `.pptx` support in-tree while heavier, licensed, regional, or proprietary implementations live outside the repository [@plugin-protocol].

## Shape

The protocol defines three plugin kinds: `dump-reader`, `exporter`, and `format-handler` [@plugin-protocol]. The same enum and wire strings are represented in `PluginKind`, so runtime selection uses the same names as manifests and the protocol document [@plugin-manifest].

A `dump-reader` reads a foreign source and emits OfficeCLI batch items for a native target. An `exporter` reads a native file and writes a foreign output file. A `format-handler` keeps a foreign file open and answers document operations over stdin and stdout [@plugin-protocol].

This design keeps the plugin boundary at process level. The main program discovers an executable, probes its manifest with `--info`, and then invokes the executable through the flow for its kind [@plugin-registry]. Short-lived plugin calls share the `PluginProcess` runner, which starts the process, reads stdout and stderr, applies an idle watchdog, and maps timeout behavior back to OfficeCLI errors [@plugin-process].

## Discovery And Manifests

Discovery is ordered. OfficeCLI first checks an environment variable, then the user plugin directory, then a bundled plugin directory beside the main executable, and then `PATH` names such as `officecli-dump-reader-doc` or `officecli-doc` [@plugin-protocol]. `PluginRegistry.CandidatePaths` implements that order and caches both positive and negative lookups for the current process [@plugin-registry].

Each candidate must answer `--info` with JSON. The registry parses that JSON as a `PluginManifest` and rejects a plugin whose protocol major version is not `1` [@plugin-registry]. The manifest model stores the name, version, protocol, kinds, extensions, target format, runtime tag, idle-timeout block, optional metadata, and format-handler vocabulary [@plugin-manifest].

The registry also hardens `PATH` lookup. It ignores relative `PATH` entries and Unix directories that are world-writable, which reduces accidental or malicious executable hijacking during plugin resolution [@plugin-registry].

## Runtime Integration

The document open path is where plugins enter normal OfficeCLI behavior. Native extensions still open built-in handlers. Other extensions first try a dump-reader and then a format-handler [@handler-factory]. If neither plugin kind resolves, OfficeCLI reports the file type as unsupported and points the user toward installed plugins and the protocol document [@handler-factory].

The plugin system therefore collaborates with the [document handler lifecycle](../handlers/document-handler-lifecycle) instead of replacing it. Dump-readers convert into a native file that built-in handlers can open. Format-handlers are wrapped as handlers through [format handler plugins](format-handler-plugins), so commands such as get, query, set, add, and validate still flow through the `IDocumentHandler` boundary [@handler-factory].

The command surface includes `plugins list`, `plugins info`, and `plugins lint`. Listing enumerates discoverable plugins and surfaces manifest warnings. Info re-runs `--info` and prints the raw manifest. Lint currently targets dump-readers: it runs `dump`, parses the JSONL batch stream, and checks emitted `add` and `set` properties against the target-format help schemas [@plugin-command].

## Constraints

The manifest is both a compatibility gate and a diagnostic surface. Protocol mismatch is fatal during probing, while softer problems such as missing idle-timeout defaults, unknown kind names, unsupported dump-reader targets, or missing format-handler vocabulary are surfaced as warnings [@plugin-manifest].

The idle watchdog is activity-based, not wall-clock based. Any stdout byte or heartbeat line can reset it, and a user can override the manifest budget with `OFFICECLI_PLUGIN_IDLE_TIMEOUT_SECONDS`, including `0` to disable the watchdog for debugging [@plugin-protocol]. This behavior is implemented in the shared process runner and is the basis for the plugin idle watchdog decision [@plugin-process].

The system's consequence is a narrow but useful extension point. Plugins can add format support without changing native handlers, but they must obey the manifest, discovery, stream, timeout, and ownership rules described by the plugin manifest and plugin discovery references [@plugin-protocol].
