---
title: "Help Schema Loader"
summary: "The help schema loader turns embedded schema JSON files into the CLI's schema-aware help, validation, flat search dumps, and schema CRC fingerprint."
topics: [architecture, cli, schemas]
sources:
  - id: schema-loader
    type: file
    path: src/officecli/Help/SchemaHelpLoader.cs
  - id: schema-renderer
    type: file
    path: src/officecli/Help/SchemaHelpRenderer.cs
  - id: flat-renderer
    type: file
    path: src/officecli/Help/SchemaHelpFlatRenderer.cs
  - id: schema-crc
    type: file
    path: src/officecli/Help/SchemaCrc.cs
  - id: help-command
    type: file
    path: src/officecli/CommandBuilder.Help.cs
  - id: project-file
    type: file
    path: src/officecli/officecli.csproj
---

The help schema loader is the boundary between embedded schema files and the user-facing command reference. It indexes `schemas/help/**` resources from the assembly, normalizes format and element aliases, composes schemas that extend shared bases, and feeds renderers used by `officecli help`. It also supports property validation and the schema CRC fingerprint, so the same embedded schema corpus drives discovery, agent-facing references, and compatibility checks [@schema-loader] [@help-command] [@schema-crc].

## Embedded Resource Index

The project embeds every `schemas/help/**/*.json` file with logical names under `schemas/help/` so a single-file OfficeCLI binary can read schemas from the assembly without extracting files to disk [@project-file]. `SchemaHelpLoader` builds a lazy manifest index from assembly resource names, normalizing backslashes to forward slashes so Windows and Unix builds resolve the same canonical keys [@schema-loader].

Only the canonical formats `docx`, `xlsx`, and `pptx` are listed. User-facing aliases map `word` to `docx`, `excel` to `xlsx`, and `ppt` or `powerpoint` to `pptx` [@schema-loader]. Unknown formats raise an error with a closest-match suggestion when possible [@schema-loader].

## Element Resolution

Element lookup starts with a case-insensitive filename match. If that fails, the loader scans the format's schemas for `elementAliases` and builds a per-format alias cache, allowing path-style names such as `p` or `col` to resolve to schema filenames such as `paragraph` or `column` [@schema-loader].

The loader also treats `/` as the root element alias for each format: workbook for `xlsx`, document for `docx`, and presentation for `pptx`. This mirrors the document path vocabulary used by get, set, and query commands [@schema-loader].

When an element is unknown, the loader suggests the closest element name and truncates caller-provided strings in error messages so a bad request cannot amplify arbitrary input into a large response [@schema-loader].

## Schema Composition

Schemas may declare `extends` as a string or an array of shared base references. The loader reads bases from `schemas/help/<base>.json`, layers multiple bases in declaration order, and applies the concrete schema last [@schema-loader]. Merge behavior is shallow for top-level fields except for `properties`: override-declared properties come first, base-only properties are appended, and a same-name property in the override replaces the base entry atomically [@schema-loader].

The loader strips the synthetic `extends` and `shared_base` markers from the composed output. That means renderers and JSON callers see the resolved schema, not the inheritance mechanics used to author it [@schema-loader].

## Help Rendering

`officecli help` uses the loader to decide whether a request is root help, command help, a format listing, a verb-filtered element listing, a specific element schema, or a flat corpus dump [@help-command]. For element pages, `SchemaHelpRenderer` can print resolved raw JSON or a human-readable view with operations, usage examples, properties, parts, children, notes, and examples [@schema-renderer].

Verb-filtered help uses the operations block. If an element does not support the requested verb, the renderer says so instead of showing an empty property list [@schema-renderer]. When properties are shown under a verb filter, only properties that declare that verb are rendered [@schema-renderer].

For broad search, `SchemaHelpFlatRenderer` emits one record per element and property. It supports human-readable flat text, JSONL, and a JSON-array form for `help all` and `help <format> all` [@flat-renderer] [@help-command]. The flat rows include operation flags, paths, property names, types, aliases, enum values, descriptions, and examples where available [@flat-renderer].

## Validation And Fingerprint

The loader also validates property keys for a given format, element, and verb. It accepts declared property names, property aliases, generic add/set keys such as `from` and `text`, known dotted namespaces such as `font.` and `border.`, and selected indexed prefixes such as `series1.` [@schema-loader]. Unknown format or element resolution is treated leniently so new handlers are not blocked by a missing schema entry [@schema-loader].

`SchemaCrc.Compute()` calculates a CRC32 over every embedded `schemas/help/` resource, ordered by normalized resource name and including both canonical names and raw bytes [@schema-crc]. `Program.cs` exposes that value through `--output-schema-crc`, giving downstream automation a cheap way to detect schema-surface drift across binary upgrades [@schema-crc].

## Consequences

Because the schema loader is embedded and deterministic, the help surface does not depend on files beside the executable. The same schema corpus drives detailed help, grep-friendly discovery, validation leniency, and upgrade fingerprinting [@project-file] [@schema-loader] [@flat-renderer] [@schema-crc]. It is part of the same CLI dispatch surface described in [Command Dispatch](command-dispatch).
