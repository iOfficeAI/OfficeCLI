---
title: Coverage Map
summary: Frozen page inventory for this first wiki build.
topics: [build, wiki, reference]
sources: []
---

## Page Inventory

### Root

- path: `almanac/getting-started.md`
  - slug: `getting-started`
  - purpose: Front door that routes contributors to the command, handler, agent integration, plugin, and release clusters.
  - planned links: `architecture/cli/command-dispatch`, `architecture/handlers/document-handler-lifecycle`, `architecture/runtime/resident-process-and-pipes`, `concepts/command-layers`
  - key evidence: `README.md`, `src/officecli/Program.cs`, `src/officecli/CommandBuilder.cs`

### `concepts/`

- path: `almanac/concepts/command-layers.md`
  - slug: `concepts/command-layers`
  - purpose: Explain the semantic view, document DOM, and raw XML layers that shape the public command model.
  - planned links: `architecture/cli/command-dispatch`, `architecture/handlers/document-handler-lifecycle`, `reference/command-surface`
  - key evidence: `README.md`, `src/officecli/Core/IDocumentHandler.cs`
- path: `almanac/concepts/document-node-and-paths.md`
  - slug: `concepts/document-node-and-paths`
  - purpose: Define the cross-format `DocumentNode` abstraction and 1-based path model used by get, query, mutations, and SDK calls.
  - planned links: `concepts/command-layers`, `concepts/selectors-vs-slash-paths`, `reference/batch-item-fields`
  - key evidence: `src/officecli/Core/DocumentNode.cs`, `src/officecli/Core/IDocumentHandler.cs`, `src/officecli/McpServer.cs`
- path: `almanac/concepts/batch-item-protocol.md`
  - slug: `concepts/batch-item-protocol`
  - purpose: Explain why batch items are the shared command object across CLI batch, resident RPC, MCP, SDKs, and in-process embedding.
  - planned links: `architecture/cli/batch-execution`, `reference/batch-item-fields`, `guides/using-batch-safely`
  - key evidence: `src/officecli/BatchTypes.cs`, `src/officecli/CommandBuilder.Batch.cs`, `src/officecli/Core/BatchExecutor.cs`
- path: `almanac/concepts/resident-sessions.md`
  - slug: `concepts/resident-sessions`
  - purpose: Define resident sessions as per-file in-memory handlers with deferred flush and serialized mutation delivery.
  - planned links: `architecture/runtime/resident-process-and-pipes`, `guides/resident-save-close-flush`, `decisions/resident-ping-liveness-invariant`
  - key evidence: `src/officecli/ResidentServer.cs`, `src/officecli/ResidentClient.cs`, `src/officecli/Core/ResidentFlushPolicy.cs`
- path: `almanac/concepts/help-schemas.md`
  - slug: `concepts/help-schemas`
  - purpose: Explain schemas as the embedded agent-facing capability contract, not narrative documentation.
  - planned links: `architecture/cli/help-schema-loader`, `reference/help-schema-format`, `reference/schema-crc`
  - key evidence: `schemas/README.md`, `src/officecli/Help/SchemaHelpLoader.cs`, `src/officecli/officecli.csproj`
- path: `almanac/concepts/officecli-skills.md`
  - slug: `concepts/officecli-skills`
  - purpose: Explain the base and specialized skills as lazy agent guidance that complements but does not replace help schemas.
  - planned links: `architecture/agent/mcp-and-skills`, `reference/bundled-skills`
  - key evidence: `SKILL.md`, `skills/officecli-docx/SKILL.md`, `src/officecli/Core/SkillInstaller.cs`
- path: `almanac/concepts/selectors-vs-slash-paths.md`
  - slug: `concepts/selectors-vs-slash-paths`
  - purpose: Clarify when OfficeCLI uses exact slash paths versus CSS-like selectors and how handlers bridge the two.
  - planned links: `concepts/document-node-and-paths`, `reference/selector-grammar-by-format`, `guides/adding-handler-mutations`
  - key evidence: `src/officecli/Handlers/Word/WordHandler.Selector.cs`, `src/officecli/Handlers/Excel/ExcelHandler.Selector.cs`, `src/officecli/Handlers/Pptx/PowerPointHandler.Selector.cs`
- path: `almanac/concepts/raw-xml-access.md`
  - slug: `concepts/raw-xml-access`
  - purpose: Explain raw XML and zip-URI access as the universal fallback below curated document operations.
  - planned links: `concepts/command-layers`, `architecture/handlers/document-handler-lifecycle`, `reference/command-surface`
  - key evidence: `src/officecli/CommandBuilder.Raw.cs`, `src/officecli/Core/RawXmlHelper.cs`, `src/officecli/Handlers/ExcelHandler.cs`, `src/officecli/Handlers/PowerPointHandler.cs`

### `architecture/cli/`

- path: `almanac/architecture/cli/command-dispatch.md`
  - slug: `architecture/cli/command-dispatch`
  - purpose: Map early dispatch, help rewriting, root command construction, command families, and shared `--json` behavior.
  - planned links: `concepts/command-layers`, `architecture/cli/batch-execution`, `architecture/runtime/mcp-shared-cli-root`, `reference/command-surface`
  - key evidence: `src/officecli/Program.cs`, `src/officecli/CommandBuilder.cs`, `src/officecli/CommandBuilder.IntegrationStubs.cs`
- path: `almanac/architecture/cli/batch-execution.md`
  - slug: `architecture/cli/batch-execution`
  - purpose: Explain input normalization, envelope unwrapping, continue-on-error semantics, shared replay, and output formatting for batches.
  - planned links: `concepts/batch-item-protocol`, `guides/using-batch-safely`, `reference/batch-item-fields`
  - key evidence: `src/officecli/BatchTypes.cs`, `src/officecli/CommandBuilder.Batch.cs`, `src/officecli/Core/BatchExecutor.cs`
- path: `almanac/architecture/cli/help-schema-loader.md`
  - slug: `architecture/cli/help-schema-loader`
  - purpose: Explain how embedded help schemas are located, alias-resolved, merged through `extends`, rendered, and fingerprinted.
  - planned links: `concepts/help-schemas`, `reference/help-schema-format`, `reference/schema-crc`
  - key evidence: `src/officecli/Help/SchemaHelpLoader.cs`, `src/officecli/Help/SchemaHelpRenderer.cs`, `src/officecli/Help/SchemaHelpFlatRenderer.cs`, `src/officecli/Help/SchemaCrc.cs`

### `architecture/runtime/`

- path: `almanac/architecture/runtime/resident-process-and-pipes.md`
  - slug: `architecture/runtime/resident-process-and-pipes`
  - purpose: Explain the resident server/client lifecycle, per-file pipe names, singleton lock, serialized command queue, autosave, and failure behavior.
  - planned links: `concepts/resident-sessions`, `architecture/agent/sdk-resident-clients`, `guides/resident-save-close-flush`, `decisions/resident-ping-liveness-invariant`
  - key evidence: `src/officecli/CommandBuilder.cs`, `src/officecli/ResidentServer.cs`, `src/officecli/ResidentClient.cs`, `src/officecli/Core/ResidentFlushPolicy.cs`
- path: `almanac/architecture/runtime/mcp-shared-cli-root.md`
  - slug: `architecture/runtime/mcp-shared-cli-root`
  - purpose: Explain the stdio JSON-RPC MCP server as a single-tool bridge into the same CLI root.
  - planned links: `architecture/cli/command-dispatch`, `architecture/agent/mcp-and-skills`, `decisions/mcp-single-tool-command-string`
  - key evidence: `src/officecli/McpServer.cs`, `src/officecli/McpInstaller.cs`, `src/officecli/Program.cs`

### `architecture/handlers/`

- path: `almanac/architecture/handlers/document-handler-lifecycle.md`
  - slug: `architecture/handlers/document-handler-lifecycle`
  - purpose: Explain file opening, safety guards, native handler dispatch, plugin fallback, and the shared `IDocumentHandler` contract.
  - planned links: `concepts/command-layers`, `concepts/document-node-and-paths`, `concepts/raw-xml-access`, `architecture/plugins/plugin-system`
  - key evidence: `src/officecli/Handlers/DocumentHandlerFactory.cs`, `src/officecli/Core/IDocumentHandler.cs`, `src/officecli/Handlers/WordHandler.cs`, `src/officecli/Handlers/ExcelHandler.cs`, `src/officecli/Handlers/PowerPointHandler.cs`
- path: `almanac/architecture/handlers/rendering-and-preview-stack.md`
  - slug: `architecture/handlers/rendering-and-preview-stack`
  - purpose: Explain the renderer registry, parsed-model render inputs, format-specific preview implementations, screenshots, and resources.
  - planned links: `architecture/handlers/document-handler-lifecycle`, `architecture/handlers/watch-preview-contract`, `decisions/renderers-receive-parsed-models`
  - key evidence: `src/officecli/Core/Rendering/`, `src/officecli/Handlers/Rendering/`, `src/officecli/CommandBuilder.View.cs`, `src/officecli/Resources/preview.css`, `src/officecli/Resources/preview.js`
- path: `almanac/architecture/handlers/watch-preview-contract.md`
  - slug: `architecture/handlers/watch-preview-contract`
  - purpose: Explain live preview watch/unwatch, marks and goto, SSE refresh, and why watch does not own document parsing.
  - planned links: `architecture/handlers/rendering-and-preview-stack`, `reference/command-surface`
  - key evidence: `src/officecli/CommandBuilder.Watch.cs`, `src/officecli/CommandBuilder.Mark.cs`, `src/officecli/Core/Watch/WatchServer.cs`, `src/officecli/Core/Watch/WatchNotifier.cs`
- path: `almanac/architecture/handlers/rich-object-helpers.md`
  - slug: `architecture/handlers/rich-object-helpers`
  - purpose: Explain shared helper subsystems for charts, diagrams, OLE, formulas, themes, and rich Office objects used across handlers.
  - planned links: `architecture/handlers/document-handler-lifecycle`, `reference/rich-object-helper-map`
  - key evidence: `src/officecli/Core/Chart/`, `src/officecli/Core/Diagram/`, `src/officecli/Core/OleHelper.cs`, `src/officecli/Core/Formula/`

### `architecture/plugins/`

- path: `almanac/architecture/plugins/plugin-system.md`
  - slug: `architecture/plugins/plugin-system`
  - purpose: Explain plugin kinds, discovery, manifest probing, and how plugins extend format support without entering the native handler code.
  - planned links: `architecture/handlers/document-handler-lifecycle`, `reference/plugin-discovery`, `reference/plugin-manifest`, `decisions/plugin-idle-watchdog`
  - key evidence: `plugins/plugin-protocol.md`, `src/officecli/Core/Plugins/PluginRegistry.cs`, `src/officecli/Core/Plugins/PluginManifest.cs`
- path: `almanac/architecture/plugins/format-handler-plugins.md`
  - slug: `architecture/plugins/format-handler-plugins`
  - purpose: Explain long-lived format-handler sessions, vocabulary snapshots, proxying, capabilities, and broken-state handling.
  - planned links: `architecture/plugins/plugin-system`, `reference/plugin-manifest`
  - key evidence: `plugins/plugin-protocol.md`, `src/officecli/Core/Plugins/FormatHandlerSession.cs`, `src/officecli/Core/Plugins/FormatHandlerProxy.cs`
- path: `almanac/architecture/plugins/dump-reader-and-exporter-plugins.md`
  - slug: `architecture/plugins/dump-reader-and-exporter-plugins`
  - purpose: Explain one-shot dump-reader and exporter flows, sibling native-file caching, output verification, and source ownership.
  - planned links: `architecture/plugins/plugin-system`, `decisions/dump-reader-buffered-replay`
  - key evidence: `plugins/plugin-protocol.md`, `src/officecli/Core/Plugins/DumpReaderInvoker.cs`, `src/officecli/Core/Plugins/ExporterInvoker.cs`

### `architecture/agent/`

- path: `almanac/architecture/agent/mcp-and-skills.md`
  - slug: `architecture/agent/mcp-and-skills`
  - purpose: Explain how MCP, `load_skill`, skill installation, and the root skill form the agent-facing workflow.
  - planned links: `concepts/officecli-skills`, `architecture/runtime/mcp-shared-cli-root`, `reference/bundled-skills`
  - key evidence: `src/officecli/McpServer.cs`, `src/officecli/Core/SkillInstaller.cs`, `src/officecli/McpInstaller.cs`, `SKILL.md`
- path: `almanac/architecture/agent/sdk-resident-clients.md`
  - slug: `architecture/agent/sdk-resident-clients`
  - purpose: Explain the Node and Python SDKs as thin resident-pipe clients with no second command vocabulary.
  - planned links: `concepts/batch-item-protocol`, `architecture/runtime/resident-process-and-pipes`, `guides/using-sdks`
  - key evidence: `sdk/node/index.js`, `sdk/node/index.d.ts`, `sdk/python/officecli.py`, `sdk/python/pyproject.toml`

### `architecture/build/`

- path: `almanac/architecture/build/self-contained-binary-and-embedded-resources.md`
  - slug: `architecture/build/self-contained-binary-and-embedded-resources`
  - purpose: Explain the .NET single-file binary, embedded schemas, embedded skills, preview assets, and trimming constraints.
  - planned links: `concepts/help-schemas`, `concepts/officecli-skills`, `guides/release-build-and-checksum-flow`
  - key evidence: `src/officecli/officecli.csproj`, `build.sh`, `npm/lib/install-binary.js`

### `guides/`

- path: `almanac/guides/using-batch-safely.md`
  - slug: `guides/using-batch-safely`
  - purpose: Show maintainers how to prepare, run, verify, and recover batch operations without command-shape mistakes.
  - planned links: `concepts/batch-item-protocol`, `architecture/cli/batch-execution`, `reference/batch-item-fields`
  - key evidence: `src/officecli/CommandBuilder.Batch.cs`, `src/officecli/BatchTypes.cs`, `README.md`
- path: `almanac/guides/resident-save-close-flush.md`
  - slug: `guides/resident-save-close-flush`
  - purpose: Show when to use `open`, `save`, `close`, and flush environment variables so external readers see current bytes.
  - planned links: `concepts/resident-sessions`, `architecture/runtime/resident-process-and-pipes`, `decisions/resident-ping-liveness-invariant`
  - key evidence: `src/officecli/CommandBuilder.cs`, `src/officecli/CommandBuilder.Save.cs`, `src/officecli/Core/ResidentFlushPolicy.cs`, `README.md`
- path: `almanac/guides/adding-handler-mutations.md`
  - slug: `guides/adding-handler-mutations`
  - purpose: Guide contributors through adding or changing a curated add/set/remove operation while keeping handlers, schemas, and warnings aligned.
  - planned links: `architecture/handlers/document-handler-lifecycle`, `concepts/selectors-vs-slash-paths`, `reference/help-schema-format`
  - key evidence: `src/officecli/Handlers/Word/WordHandler.Set.cs`, `src/officecli/Handlers/Excel/ExcelHandler.Set.cs`, `src/officecli/Handlers/Pptx/PowerPointHandler.Set.cs`, `schemas/README.md`
- path: `almanac/guides/debugging-preview-rendering.md`
  - slug: `guides/debugging-preview-rendering`
  - purpose: Guide maintainers through tracing a preview, screenshot, or watch rendering issue across registry, handler preview, and static resources.
  - planned links: `architecture/handlers/rendering-and-preview-stack`, `architecture/handlers/watch-preview-contract`
  - key evidence: `src/officecli/CommandBuilder.View.cs`, `src/officecli/Core/Rendering/`, `src/officecli/Handlers/Word/WordHandler.HtmlPreview.cs`, `src/officecli/Handlers/Excel/ExcelHandler.HtmlPreview.cs`, `src/officecli/Handlers/Pptx/PowerPointHandler.HtmlPreview.cs`
- path: `almanac/guides/installing-and-platform-detection.md`
  - slug: `guides/installing-and-platform-detection`
  - purpose: Guide maintainers through installer platform detection, checksum verification, binary placement, PATH updates, and skill/MCP side effects.
  - planned links: `architecture/build/self-contained-binary-and-embedded-resources`, `reference/platform-assets-and-install-surfaces`, `decisions/immutable-release-downloads`
  - key evidence: `install.sh`, `install.ps1`, `npm/lib/install-binary.js`, `src/officecli/Core/Installer.cs`
- path: `almanac/guides/release-build-and-checksum-flow.md`
  - slug: `guides/release-build-and-checksum-flow`
  - purpose: Guide maintainers through release build assets, signing/notarization, smoke checks, checksums, and npm publishing.
  - planned links: `architecture/build/self-contained-binary-and-embedded-resources`, `reference/platform-assets-and-install-surfaces`, `decisions/immutable-release-downloads`
  - key evidence: `build.sh`, `.github/workflows/build.yml`, `.github/workflows/publish-npm.yml`, `npm/package.json`
- path: `almanac/guides/using-sdks.md`
  - slug: `guides/using-sdks`
  - purpose: Guide maintainers and integrators through using the Node and Python SDKs safely with resident handles.
  - planned links: `architecture/agent/sdk-resident-clients`, `concepts/batch-item-protocol`, `guides/resident-save-close-flush`
  - key evidence: `sdk/node/README.md`, `sdk/node/index.js`, `sdk/python/README.md`, `sdk/python/officecli.py`
- path: `almanac/guides/examples-corpus.md`
  - slug: `guides/examples-corpus`
  - purpose: Guide contributors through using the examples directory as executable feature coverage and source material.
  - planned links: `reference/command-surface`, `reference/bundled-skills`
  - key evidence: `examples/README.md`, `examples/excel/`, `examples/ppt/`, `examples/word/`

### `decisions/`

- path: `almanac/decisions/mcp-single-tool-command-string.md`
  - slug: `decisions/mcp-single-tool-command-string`
  - purpose: Record the choice to expose one MCP tool that accepts CLI command strings or argv arrays instead of per-command MCP tools.
  - planned links: `architecture/runtime/mcp-shared-cli-root`, `architecture/agent/mcp-and-skills`
  - key evidence: `src/officecli/McpServer.cs`
- path: `almanac/decisions/resident-ping-liveness-invariant.md`
  - slug: `decisions/resident-ping-liveness-invariant`
  - purpose: Record the invariant that ping availability must imply handler ownership and why shutdown ordering preserves direct-mode safety.
  - planned links: `architecture/runtime/resident-process-and-pipes`, `guides/resident-save-close-flush`
  - key evidence: `src/officecli/ResidentServer.cs`, `src/officecli/ResidentClient.cs`
- path: `almanac/decisions/plugin-idle-watchdog.md`
  - slug: `decisions/plugin-idle-watchdog`
  - purpose: Record the activity-based plugin timeout model and user override behavior.
  - planned links: `architecture/plugins/plugin-system`, `reference/plugin-manifest`
  - key evidence: `plugins/plugin-protocol.md`, `src/officecli/Core/Plugins/PluginProcess.cs`, `src/officecli/Core/Plugins/PluginManifest.cs`
- path: `almanac/decisions/dump-reader-buffered-replay.md`
  - slug: `decisions/dump-reader-buffered-replay`
  - purpose: Record why dump-reader JSONL output is buffered before replay even though the protocol requires streaming output.
  - planned links: `architecture/plugins/dump-reader-and-exporter-plugins`, `architecture/plugins/plugin-system`
  - key evidence: `src/officecli/Core/Plugins/DumpReaderInvoker.cs`, `plugins/plugin-protocol.md`
- path: `almanac/decisions/immutable-release-downloads.md`
  - slug: `decisions/immutable-release-downloads`
  - purpose: Record why installers and npm downloads prefer immutable versioned release URLs over mutable latest URLs.
  - planned links: `guides/installing-and-platform-detection`, `guides/release-build-and-checksum-flow`
  - key evidence: `install.sh`, `install.ps1`, `npm/lib/install-binary.js`
- path: `almanac/decisions/atomic-binary-replacement.md`
  - slug: `decisions/atomic-binary-replacement`
  - purpose: Record why local installer and dev install flows use atomic replacement rather than in-place overwrite.
  - planned links: `guides/installing-and-platform-detection`, `architecture/build/self-contained-binary-and-embedded-resources`
  - key evidence: `install.sh`, `dev-install.sh`, `build.sh`
- path: `almanac/decisions/renderers-receive-parsed-models.md`
  - slug: `decisions/renderers-receive-parsed-models`
  - purpose: Record the renderer API choice that renderers receive parsed model hosts instead of reopening files by path.
  - planned links: `architecture/handlers/rendering-and-preview-stack`, `architecture/handlers/document-handler-lifecycle`
  - key evidence: `src/officecli/Core/Rendering/IRenderer.cs`, `src/officecli/Handlers/Rendering/HandlerRenderInput.cs`, `src/officecli/Handlers/Rendering/BasicRenderers.cs`

### `reference/`

- path: `almanac/reference/command-surface.md`
  - slug: `reference/command-surface`
  - purpose: Lookup reference for public command families, early-dispatched commands, hidden compatibility aliases, and command ownership files.
  - planned links: `architecture/cli/command-dispatch`, `concepts/command-layers`
  - key evidence: `src/officecli/Program.cs`, `src/officecli/CommandBuilder.cs`, `src/officecli/CommandBuilder.*.cs`
- path: `almanac/reference/batch-item-fields.md`
  - slug: `reference/batch-item-fields`
  - purpose: Exact field reference for `BatchItem`, lenient props, aliases, and resident request conversion.
  - planned links: `concepts/batch-item-protocol`, `architecture/cli/batch-execution`
  - key evidence: `src/officecli/BatchTypes.cs`, `src/officecli/CommandBuilder.Batch.cs`
- path: `almanac/reference/selector-grammar-by-format.md`
  - slug: `reference/selector-grammar-by-format`
  - purpose: Lookup reference comparing selector capabilities and limits for Word, Excel, and PowerPoint.
  - planned links: `concepts/selectors-vs-slash-paths`, `guides/adding-handler-mutations`
  - key evidence: `src/officecli/Handlers/Word/WordHandler.Selector.cs`, `src/officecli/Handlers/Excel/ExcelHandler.Selector.cs`, `src/officecli/Handlers/Pptx/PowerPointHandler.Selector.cs`
- path: `almanac/reference/help-schema-format.md`
  - slug: `reference/help-schema-format`
  - purpose: Exact reference for help schema files, including operations, paths, aliases, properties, examples, enforcement, and shared bases.
  - planned links: `concepts/help-schemas`, `architecture/cli/help-schema-loader`, `guides/adding-handler-mutations`
  - key evidence: `schemas/README.md`, `schemas/help/_schema.json`, `schemas/help/docx/paragraph.json`, `schemas/help/xlsx/cell.json`, `schemas/help/pptx/shape.json`
- path: `almanac/reference/schema-crc.md`
  - slug: `reference/schema-crc`
  - purpose: Lookup reference for the schema CRC command and what it does and does not fingerprint.
  - planned links: `concepts/help-schemas`, `architecture/cli/help-schema-loader`
  - key evidence: `src/officecli/Program.cs`, `src/officecli/Help/SchemaCrc.cs`
- path: `almanac/reference/plugin-discovery.md`
  - slug: `reference/plugin-discovery`
  - purpose: Exact plugin discovery order, executable names, cache behavior, and hardening rules.
  - planned links: `architecture/plugins/plugin-system`, `reference/plugin-manifest`
  - key evidence: `plugins/plugin-protocol.md`, `src/officecli/Core/Plugins/PluginRegistry.cs`
- path: `almanac/reference/plugin-manifest.md`
  - slug: `reference/plugin-manifest`
  - purpose: Exact plugin manifest fields, supported protocol version, kinds, target, idle timeout, and vocabulary fields.
  - planned links: `architecture/plugins/plugin-system`, `decisions/plugin-idle-watchdog`
  - key evidence: `plugins/plugin-protocol.md`, `src/officecli/Core/Plugins/PluginManifest.cs`
- path: `almanac/reference/platform-assets-and-install-surfaces.md`
  - slug: `reference/platform-assets-and-install-surfaces`
  - purpose: Lookup reference for platform binary names, install locations, npm package behavior, package-manager surfaces, and config side effects.
  - planned links: `guides/installing-and-platform-detection`, `guides/release-build-and-checksum-flow`
  - key evidence: `README.md`, `install.sh`, `install.ps1`, `npm/package.json`, `npm/lib/install-binary.js`
- path: `almanac/reference/bundled-skills.md`
  - slug: `reference/bundled-skills`
  - purpose: Lookup reference for bundled skill names, aliases, specialized use cases, and installation targets.
  - planned links: `concepts/officecli-skills`, `architecture/agent/mcp-and-skills`
  - key evidence: `src/officecli/Core/SkillInstaller.cs`, `SKILL.md`, `skills/`
- path: `almanac/reference/rich-object-helper-map.md`
  - slug: `reference/rich-object-helper-map`
  - purpose: Lookup reference for shared chart, diagram, OLE, formula, theme, and style helper directories and what each supports.
  - planned links: `architecture/handlers/rich-object-helpers`
  - key evidence: `src/officecli/Core/Chart/`, `src/officecli/Core/Diagram/`, `src/officecli/Core/OleHelper.cs`, `src/officecli/Core/Formula/`, `src/officecli/Core/TableStyles/`
- path: `almanac/reference/security-policy.md`
  - slug: `reference/security-policy`
  - purpose: Lookup reference for supported security versions, private reporting, and the security-sensitive file-opening guards in this repo.
  - planned links: `architecture/handlers/document-handler-lifecycle`, `architecture/plugins/plugin-system`
  - key evidence: `SECURITY.md`, `src/officecli/Handlers/DocumentHandlerFactory.cs`, `src/officecli/Core/SsrfGuard.cs`, `src/officecli/Core/DocumentLimits.cs`
