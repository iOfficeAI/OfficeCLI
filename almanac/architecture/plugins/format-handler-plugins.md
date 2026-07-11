---
title: "Format Handler Plugins"
summary: "Format-handler plugins are long-lived sidecar sessions that let a non-native file behave like an OfficeCLI document handler."
topics: [architecture, plugins]
sources:
  - id: plugin-protocol
    type: file
    path: plugins/plugin-protocol.md
  - id: format-session
    type: file
    path: src/officecli/Core/Plugins/FormatHandlerSession.cs
  - id: format-proxy
    type: file
    path: src/officecli/Core/Plugins/FormatHandlerProxy.cs
  - id: handler-factory
    type: file
    path: src/officecli/Handlers/DocumentHandlerFactory.cs
  - id: plugin-manifest
    type: file
    path: src/officecli/Core/Plugins/PluginManifest.cs
---

Format-handler plugins make a foreign file act like a first-class OfficeCLI document. The main process starts a plugin with `open <file>`, completes an open handshake, caches the plugin's capabilities and vocabulary, and wraps the session in an `IDocumentHandler` proxy. That lets the rest of OfficeCLI use the same read, query, mutation, validation, and raw-access paths it uses for native handlers [@plugin-protocol] [@format-session] [@format-proxy].

## Session Boundary

A format-handler plugin owns the foreign file for the lifetime of the session. The protocol says the plugin keeps the file open, uses stdin for requests, stdout for JSONL replies, and stderr for diagnostics or heartbeat lines [@plugin-protocol]. `FormatHandlerSession` implements that boundary with a child process, UTF-8 no-BOM streams, an internal I/O lock, and a start-time open handshake [@format-session].

The session begins by sending an `open` envelope with protocol version, file path, and editability. The plugin must reply with `ok` and may include runtime capabilities and vocabulary [@plugin-protocol]. OfficeCLI deserializes that reply into `PluginSessionCapabilities`, which contains supported commands, optional features, and a vocabulary snapshot [@plugin-manifest] [@format-session].

The manifest vocabulary is useful for discovery and help, but the runtime snapshot is the session's live contract. The protocol allows the snapshot to differ from the manifest, and the session caches the open-handshake result for the session lifetime [@plugin-protocol] [@format-session].

## Proxy Behavior

`FormatHandlerProxy` is the adapter that makes the plugin look like an `IDocumentHandler`. It forwards text views, annotated views, outlines, stats, issues, get, query, validate, set, add, remove, move, copy, raw, raw-set, add-part, binary extraction, and save through the session [@format-proxy].

Most proxy methods build a JSON object of command arguments, convert user properties into a string map, and call `FormatHandlerSession.Send`. Reply shapes are strict where the protocol requires them. For example, `set` reads `unsupported_properties`, and `add` reads the created `path` plus unsupported properties [@format-proxy].

The proxy also provides plugin-side fallbacks for view modes that built-in handlers historically owned directly. SVG, HTML, and forms JSON are sent as `view` commands with plugin-specific modes, and `unsupported_command` is treated as absence of that optional feature [@format-proxy].

## Failure And Concurrency

The protocol permits only one request in flight per session, and the session enforces that with an internal lock [@plugin-protocol] [@format-session]. Before a command is sent, the session checks the cached capability list. If the plugin declared supported commands and the requested command is missing, OfficeCLI fails fast with `unsupported_command` without a round trip [@format-session].

Malformed stdout poisons the session. If a plugin writes non-JSON, returns a non-object reply, uses an unknown `msg_type`, closes stdout unexpectedly, or hits an I/O failure, the session enters a broken state and later sends fail with `plugin_stream_closed` [@format-session].

Long commands use the same idle-timeout model as the wider [plugin system](plugin-system). The session waits for a reply while polling activity. Stderr heartbeats reset the timer; if no activity arrives within the verb's timeout, OfficeCLI kills the process tree and raises `plugin_idle_timeout` [@format-session].

## Opening Flow

When a file extension is not native, the document handler factory tries plugins. Dump-readers are tried first. If no dump-reader resolves, the factory looks for a `format-handler`, starts a `FormatHandlerSession`, and returns a `FormatHandlerProxy` [@handler-factory].

This ordering gives migrations a chance to produce native sibling files before handing full ownership to a foreign-format process. When a format-handler is used, the rest of the system sees a handler object, not a special-case plugin branch [@handler-factory]. The plugin still defines its own document vocabulary through the protocol and manifest, which keeps format-specific semantics outside the main binary [@plugin-protocol].

The result is a controlled foreign-format integration. The main process owns command routing and handler boundaries; the plugin owns parsing, mutation, save durability, and any format-specific vocabulary [@plugin-protocol] [@format-proxy].
