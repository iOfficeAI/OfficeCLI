---
title: "Document Node And Paths"
summary: "DocumentNode is OfficeCLI's cross-format tree object, and slash paths are the stable addresses used by get, query output, mutations, batches, and agent calls."
topics: [concepts, handlers, cli]
sources:
  - id: document-node
    type: file
    path: src/officecli/Core/DocumentNode.cs
  - id: handler-paths
    type: file
    path: src/officecli/Core/IDocumentHandler.cs
  - id: get-query-command
    type: file
    path: src/officecli/CommandBuilder.GetQuery.cs
  - id: mcp-command-line
    type: file
    path: src/officecli/McpServer.cs
---

`DocumentNode` is the common object OfficeCLI uses to describe Word, Excel, PowerPoint, and plugin-backed document structure. Its `path` is the durable address of a node, while fields such as `type`, `text`, `preview`, `style`, `childCount`, `format`, and `children` carry the visible data returned to callers [@document-node]. Paths matter because most read and mutation commands use them as their exact target language.

The shared handler contract makes paths central. `Get` takes a path and depth, while `Set`, `Remove`, `Move`, `CopyFrom`, and binary extraction all target paths directly [@handler-paths]. `Query` is different: it accepts a selector and returns matching `DocumentNode` values, which then carry paths that can be reused by later commands [@handler-paths].

## Slash Paths

Slash paths name positions in the document tree. The CLI describes `get` as reading a DOM path such as `/body/p[1]`, and it defaults to `/` when no path is supplied [@get-query-command]. Other examples in the agent-facing MCP guidance use the same 1-based bracketed form, such as `/slide[1]`, so agents can pass the same path string through the CLI tool instead of learning a separate transport model [@mcp-command-line].

The path string is exact. A command such as `set` or `remove` is expected to name one target node, not a broad query. This is why [selectors vs slash paths](selectors-vs-slash-paths) is a separate concept: selectors are for finding nodes, while paths are for addressing the chosen node.

## Node Shape

`DocumentNode` is deliberately small. User-facing JSON includes the address, type, text-like summaries, a `format` dictionary, and optional children [@document-node]. The class also has `InternalFormat`, but it is ignored by JSON serialization and is reserved for round-trip metadata such as verbatim OOXML fragments used by internal emitters [@document-node].

That split is important. The public node shape stays stable for CLI, MCP, batch, and SDK consumers, while handlers can still carry internal details needed for faithful document rewrites.

## How Paths Flow

The normal path flow is: use a semantic view or selector to locate content, read the node with `get`, then pass the node path into a mutation. `query` returns node lists, and `get --json` uses the same `{matches, results}` shape as query so consumers can parse both through one path [@get-query-command].

Batch items use the same address fields described in [batch item fields](../reference/batch-item-fields). That is why a path learned from `get`, query output, a watch selection, or MCP text can be carried into a later batch or resident request without translating it into a format-specific object id.
