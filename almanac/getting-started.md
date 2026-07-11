---
title: "Getting Started"
summary: "Start here to navigate the OfficeCLI wiki by command dispatch, handlers, resident runtime, command layers, skills, selectors, and raw XML."
topics: [guides, architecture, cli]
sources:
  - id: readme
    type: file
    path: README.md
  - id: program
    type: file
    path: src/officecli/Program.cs
  - id: command-builder
    type: file
    path: src/officecli/CommandBuilder.cs
---

OfficeCLI is a single-binary CLI for creating, reading, rendering, validating, and modifying `.docx`, `.xlsx`, and `.pptx` files without requiring Microsoft Office [@readme]. This wiki starts from the command surface, then routes into handlers, resident runtime behavior, agent integration, and lower-level concepts. Use this page as a map, not as a tutorial.

## Start With Commands

Begin with [command dispatch](architecture/cli/command-dispatch) if you need to understand how a terminal invocation becomes OfficeCLI behavior. Startup code rewrites `--help` into the schema-driven `help` command, handles early-dispatch commands such as `mcp`, `install`, `skills`, and `load_skill`, then builds the root command for normal document operations [@program].

For the public command vocabulary, read [command surface](reference/command-surface). The root command registers the main document verbs: `open`, `close`, `watch`, `view`, `get`, `query`, `set`, `add`, `remove`, `move`, `swap`, `raw`, `raw-set`, `add-part`, `validate`, `save`, `batch`, `dump`, `import`, `create`, `merge`, plugins, and help [@command-builder].

## Understand The Document Model

Read [command layers](concepts/command-layers) before changing command behavior. OfficeCLI exposes high-level document operations for creating, reading, analyzing, modifying, and reorganizing Office content, while still keeping lower-level commands available for exact nodes and raw package parts [@readme] [@command-builder].

Then read [document node and paths](concepts/document-node-and-paths) and [selectors vs slash paths](concepts/selectors-vs-slash-paths). Exact slash paths identify known nodes, while selectors search by type, content, or attributes. That distinction is central to `get`, `query`, `set`, `remove`, handler mutation code, and batch replay.

## Follow Handler And Runtime Flow

Use [document handler lifecycle](architecture/handlers/document-handler-lifecycle) when working on file opening, format dispatch, validation, or handler contracts. The root command delegates document verbs to handler-backed command builders, and those builders decide whether to use a resident process or open a handler directly [@command-builder].

Use [resident process and pipes](architecture/runtime/resident-process-and-pipes) when debugging performance, file locking, save/close behavior, or serialized mutation delivery. The root command has explicit `open` and `close` commands, starts a hidden `__resident-serve__` process, and uses a resident client probe before falling back to direct file handling [@command-builder].

## Use The Agent Cluster

Read [MCP and skills](architecture/agent/mcp-and-skills) to understand the agent-facing workflow. Program dispatch starts the MCP server with `officecli mcp`, registers or unregisters clients through `mcp <target>`, and exposes `load_skill` as the read-only way to fetch bundled skill guidance [@program].

The concept page [OfficeCLI skills](concepts/officecli-skills) explains why skills exist beside schemas. The reference page [bundled skills](reference/bundled-skills) is the planned lookup surface for the skill inventory.

## Use Lower-Level References When Needed

For broad searches and mutation targeting, use [selector grammar by format](reference/selector-grammar-by-format). For last-resort OpenXML work, use [raw XML access](concepts/raw-xml-access). For adding or changing curated mutations, use [adding handler mutations](guides/adding-handler-mutations). These pages sit downstream of the command and handler architecture, so they make the most sense after the routing and lifecycle pages.
