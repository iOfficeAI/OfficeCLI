---
title: "Help Schemas"
summary: "Help schemas are embedded JSON capability contracts that tell agents which OfficeCLI elements, verbs, properties, aliases, and examples are supported."
topics: [concepts, schemas, cli]
sources:
  - id: schemas-readme
    type: file
    path: schemas/README.md
  - id: schema-loader
    type: file
    path: src/officecli/Help/SchemaHelpLoader.cs
  - id: project-embedding
    type: file
    path: src/officecli/officecli.csproj
---

Help schemas are OfficeCLI's machine-readable capability contract. They live under `schemas/help/`, are embedded into the binary, and describe what each supported Word, Excel, and PowerPoint element can do through operations, properties, aliases, and examples [@schemas-readme]. They are not narrative documentation; they are the data source agents and tests use to stay aligned with the command surface.

The schema system matters because OfficeCLI has a large cross-format vocabulary. A caller should be able to ask what `docx`, `xlsx`, or `pptx` supports and receive a precise answer from the same binary that will execute the command.

## Embedded Contract

The project file embeds every JSON file under `schemas/help/**` as an assembly resource with a `schemas/help/...` logical name [@project-embedding]. `SchemaHelpLoader` builds an index from manifest resource names and opens schemas from the assembly rather than from the runtime filesystem [@schema-loader].

That design is important for the self-contained binary. The schema reference travels with the executable, so `help` can work without a checked-out source tree, extracted schema directory, or network access.

## Resolution Model

The loader normalizes format aliases: `word` maps to `docx`, `excel` to `xlsx`, and `ppt` or `powerpoint` to `pptx` [@schema-loader]. It also treats `/` as the root element alias for the relevant format, mapping it to workbook, document, or presentation [@schema-loader].

Element lookup first tries the schema filename and then scans `elementAliases` declared in schemas, so path-style names such as short element aliases can still resolve to their canonical schema file [@schema-loader]. A schema may also declare `extends`, and the loader composes shared bases before applying the element-specific override [@schema-loader].

## How They Differ From Wiki Pages

The schemas README draws a sharp boundary: schemas are the single source of truth for supported add, set, get, and readback behavior, while narrative tutorials and best practices belong in the wiki [@schemas-readme]. That is why [help schema loader](../architecture/cli/help-schema-loader), [help schema format](../reference/help-schema-format), and [schema CRC](../reference/schema-crc) are separate pages. The concept is the contract; the architecture explains loading; the references describe file shape and fingerprinting.

When command behavior changes, the matching schema is expected to change with it, and contract tests verify schema claims against real handler behavior [@schemas-readme]. This keeps the agent-facing help surface from becoming aspirational documentation.
