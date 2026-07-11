---
title: "Dump Reader And Exporter Plugins"
summary: "Dump-reader and exporter plugins are one-shot sidecar flows for converting foreign inputs into native OfficeCLI files or rendering native files to foreign outputs."
topics: [architecture, plugins]
sources:
  - id: plugin-protocol
    type: file
    path: plugins/plugin-protocol.md
  - id: dump-invoker
    type: file
    path: src/officecli/Core/Plugins/DumpReaderInvoker.cs
  - id: exporter-invoker
    type: file
    path: src/officecli/Core/Plugins/ExporterInvoker.cs
  - id: handler-factory
    type: file
    path: src/officecli/Handlers/DocumentHandlerFactory.cs
  - id: plugin-process
    type: file
    path: src/officecli/Core/Plugins/PluginProcess.cs
---

Dump-reader and exporter plugins are OfficeCLI's short-lived plugin paths. A dump-reader converts a foreign source into a native OfficeCLI document by emitting batch items; an exporter converts a native document into a foreign output file. They share discovery and timeout rules with the [plugin system](plugin-system), but they do not keep an editing session open [@plugin-protocol].

## Dump Readers

A dump-reader is used when a non-native file should be migrated into `.docx`, `.xlsx`, or `.pptx`. The protocol requires the plugin to run as `<plugin> dump <source>` and write one JSON object per line to stdout, where each line is a batch item in the target format's vocabulary [@plugin-protocol].

`DumpReaderInvoker.Run` resolves a dump-reader for the source extension, creates a minimal blank native temp file based on the plugin manifest target, runs the plugin, and replays the emitted batch into that temp file [@dump-invoker]. The emitted target is native, so after replay the normal document-handler machinery can open it [@handler-factory].

The protocol says dump-reader output must stream. The current invoker still requires JSONL from the plugin, rejects top-level JSON arrays, and uses stdout activity for the idle watchdog, but it buffers all non-empty JSONL lines until the plugin exits before replaying them [@dump-invoker] [@plugin-process]. The code comment records why: replaying OpenXML mutations from the stdout reader thread hit non-thread-safe package behavior, so replay now happens synchronously on the caller thread [@dump-invoker]. That implementation detail is the core reason behind the dump reader buffered replay decision.

The document handler factory adds a sibling-file cache around dump-readers. When a foreign source such as `.doc` is opened, OfficeCLI checks for a sibling native file with the plugin target extension. If the sibling exists and is newer than the source, OfficeCLI opens it directly. Otherwise it runs the dump-reader, writes or reuses the sibling, and reports that the sibling will be reused later [@handler-factory].

Edits belong to the sibling native file, not to the original foreign source. The protocol states this ownership rule, and the factory implements it by opening the generated or cached native sibling after conversion [@plugin-protocol] [@handler-factory].

## Exporters

An exporter is used when a native Office document should be rendered to a foreign target such as PDF. The protocol runs it as `<plugin> export <source-file> --out <target-file>` and requires the plugin not to modify the source file [@plugin-protocol].

`ExporterInvoker.Run` resolves an exporter by target extension, filters it by source support when the manifest advertises `supports`, closes an active resident for the source if one exists, runs the plugin, and verifies that the output path was written [@exporter-invoker]. If no exporter matches, the error is `exporter_not_found`; non-zero plugin exits map to plugin error codes such as corrupt input, unsupported feature, license expiry, protocol mismatch, or generic plugin failure [@exporter-invoker].

The exporter path passes the source directly to the plugin. That is why the source-read-only rule is important: OfficeCLI does not snapshot the native file before export [@plugin-protocol] [@exporter-invoker].

## Shared Failure Model

Both flows use the shared short-lived process runner. The runner starts the executable with redirected stdout and stderr, sets `OFFICECLI_BIN` for plugins that need to call back into the current binary, and kills the process tree if no activity arrives within the resolved idle timeout [@plugin-process].

For dump-readers, any stdout line counts as activity. For exporters, long rendering work should use stderr heartbeat lines because exporter output is normally the target file, not stdout JSONL [@plugin-protocol] [@plugin-process].

These one-shot plugin types keep ownership clear. Dump-readers hand control back to native handlers after conversion. Exporters never enter the document-handler interface at all; they only consume a native file and produce an external artifact [@plugin-protocol] [@exporter-invoker].
