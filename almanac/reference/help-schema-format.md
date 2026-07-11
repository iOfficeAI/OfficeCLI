---
title: "Help Schema Format"
summary: "Exact reference for OfficeCLI help schema files, including element metadata, paths, operations, properties, examples, inheritance, and runtime validation behavior."
topics: [reference, cli, schemas]
sources:
  - id: schemas-readme
    type: file
    path: schemas/README.md
  - id: schema-meta
    type: file
    path: schemas/help/_schema.json
  - id: schema-loader
    type: file
    path: src/officecli/Help/SchemaHelpLoader.cs
  - id: schema-renderer
    type: file
    path: src/officecli/Help/SchemaHelpRenderer.cs
  - id: flat-renderer
    type: file
    path: src/officecli/Help/SchemaHelpFlatRenderer.cs
  - id: help-command
    type: file
    path: src/officecli/CommandBuilder.Help.cs
  - id: docx-paragraph
    type: file
    path: schemas/help/docx/paragraph.json
  - id: xlsx-cell
    type: file
    path: schemas/help/xlsx/cell.json
  - id: pptx-shape
    type: file
    path: schemas/help/pptx/shape.json
---

OfficeCLI help schemas are JSON files under `schemas/help/` that describe one format and element pair: supported operations, paths, properties, aliases, readback, children, examples, and special raw parts. The repository treats these files as the agent-facing capability source of truth for add, set, get, and readback behavior, while the runtime loader embeds and resolves them for `officecli help` [@schemas-readme] [@schema-loader]. For the loader path, see [Help Schema Loader](../architecture/cli/help-schema-loader); for the concept, see [Help Schemas](../concepts/help-schemas); for the upgrade fingerprint, see [Schema CRC](schema-crc).

## File Layout

The schema tree has a meta-schema at `schemas/help/_schema.json`, shared bases under `schemas/help/_shared/`, and per-format element files under `schemas/help/docx/`, `schemas/help/xlsx/`, and `schemas/help/pptx/` [@schemas-readme]. The README says PRs that change `Add`, `Set`, or `Get` behavior for an element must update the matching schema file in the same PR, with contract tests checking schema claims against handlers [@schemas-readme].

## Required Top-Level Fields

The JSON Schema requires `format`, `element`, `operations`, and `properties` [@schema-meta].

| Field | Shape | Meaning |
| --- | --- | --- |
| `$schema` | string | Pointer to the meta-schema for tooling; ignored at runtime [@schema-meta]. |
| `format` | enum | One of `docx`, `xlsx`, or `pptx` [@schema-meta]. |
| `element` | string | CLI element name, such as `paragraph`, `cell`, or `shape` [@schema-meta]. |
| `parent` | string or array | Parent element name or names when the element exists only under another element [@schema-meta]. |
| `note` | string | Free-form clarification for caveats or invariants [@schema-meta]. |
| `container` | boolean | Marks read-only root or container entities that are navigated but not created or mutated [@schema-meta]. |
| `operations` | object | Top-level support flags for `add`, `set`, `get`, `query`, and `remove` [@schema-meta]. |
| `addParent` | string or array | Explicit Add parent path when it cannot be derived from the element path [@schema-meta]. |
| `paths` | object | `stable` and `positional` path examples for addressing the element [@schema-meta]. |
| `addressing` | object | Keyed child path form, with `key`, `pathForm`, and optional `keyValues` [@schema-meta]. |
| `properties` | object | Map of canonical property name to property definition [@schema-meta]. |
| `children` | array | Child element references with path segment and cardinality [@schema-meta]. |
| `parts` | array | Raw-part listing for synthetic `raw` schemas [@schema-meta]. |
| `examples` | array | Element-level examples, especially for raw schemas [@schema-meta]. |
| `description` | string | Element-level description [@schema-meta]. |
| `elementAliases` | array | Alternate element names that resolve to the same schema file [@schema-meta]. |

## Property Definition

Each property definition must have a `type` according to the meta-schema [@schema-meta]. Allowed property types are `string`, `bool`, `number`, `color`, `length`, `font-size`, and `enum` [@schema-meta].

| Field | Shape | Meaning |
| --- | --- | --- |
| `description` | string | Human-readable property meaning [@schema-meta]. |
| `aliases` | array or object | Input aliases. Array form is a plain alias list; object form maps alias to canonical value [@schema-meta]. |
| `values` | string array | Allowed values for `type: enum` [@schema-meta]. |
| `modifiers` | object | Modifier composition rules for enum-like properties [@schema-meta]. |
| `appliesWhen` | object | Conditional applicability; all listed state keys must match [@schema-meta]. |
| `requires` | string array | Other same-element properties required for a well-formed OOXML result [@schema-meta]. |
| `add`, `set`, `get` | boolean | Per-property verb participation [@schema-meta]. |
| `examples` | string array | Property-level command examples [@schema-meta]. |
| `readback` | string | Expected Get readback shape [@schema-meta]. |
| `enforcement` | enum | `strict` breaks CI on drift; `report` logs drift [@schema-meta]. |

Runtime property validation also accepts `propAliases` arrays in property definitions even though the meta-schema excerpt names `aliases`; the loader explicitly reads `propAliases` after `aliases` while building allowed keys [@schema-loader].

## Examples In Current Schemas

`schemas/help/docx/paragraph.json` declares `format: docx`, `element: paragraph`, alias `p`, all five operations, stable and positional body paragraph paths, `run` children, and `extends: _shared/paragraph` [@docx-paragraph]. `schemas/help/xlsx/cell.json` declares `parent: sheet`, all five operations, paths such as `/<sheetName>/<A1Ref>` and `/Sheet1/A1`, and cell-specific property aliases such as `bold` for `font.bold` [@xlsx-cell]. `schemas/help/pptx/shape.json` declares all five operations, stable id and positional shape paths, `extends: _shared/shape`, and shape-specific properties such as `geometry`, `opacity`, `isTitle`, and text-direction keys [@pptx-shape].

## Inheritance And Aliases

A schema may declare `extends` as a string or an array. The loader reads shared bases from `schemas/help/<base>.json`, applies bases in declaration order, then applies the concrete schema last [@schema-loader]. Top-level fields replace base values except `properties`: override-declared properties come first, base-only properties are appended, and same-name properties in the override replace the base property atomically [@schema-loader].

Format aliases are `word` to `docx`, `excel` to `xlsx`, and `ppt` or `powerpoint` to `pptx` [@schema-loader]. Element lookup first tries the filename stem, then scans `elementAliases` inside schemas, so `p`, `col`, or `img` can resolve to canonical schema filenames when declared [@schema-loader]. The root path `/` resolves to `workbook`, `document`, or `presentation` by format [@schema-loader].

## Help Rendering

`officecli help` lists formats, lists elements, filters elements by schema verb, renders element detail, renders raw JSON, and emits flat dumps through `help all`, `help <format> all`, `--json`, and `--jsonl` [@help-command]. Human rendering prints header, read-only container status, description, parent, paths, addressing, operations, synthesized usage, properties filtered by verb when needed, parts, children, notes, and examples [@schema-renderer]. Flat rendering emits `ELEM` and `PROP` rows, with an `ops` legend for `add`, `set`, `get`, `query`, and `remove` [@flat-renderer].

## Runtime Validation

`SchemaHelpLoader.ValidateProperties()` checks property dictionaries for a format, element, and verb, but it is intentionally lenient when a format or element is unknown [@schema-loader]. It accepts declared property names, array-form `aliases`, `propAliases`, generic add/set keys such as `from`, `copyFrom`, `path`, `positional`, and `text`, dotted namespaces such as `font.`, `border.`, and `series.`, and indexed prefixes such as `series1.` and `dataLabel3.` [@schema-loader].
