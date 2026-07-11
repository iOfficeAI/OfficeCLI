---
title: "Batch Item Fields"
summary: "Exact reference for the JSON fields accepted by OfficeCLI batch items and how those fields map to resident requests and per-item results."
topics: [reference, cli, batch]
sources:
  - id: batch-types
    type: file
    path: src/officecli/BatchTypes.cs
  - id: batch-command
    type: file
    path: src/officecli/CommandBuilder.Batch.cs
  - id: batch-dispatch
    type: file
    path: src/officecli/CommandBuilder.cs
---

A batch item is one JSON object that names a bare OfficeCLI verb and supplies that verb's arguments as sibling fields. The accepted item keys are fixed by `BatchItem.KnownFields`, while deserialization is lenient about `command` versus `op` and about `props` object versus `["key=value"]` array form [@batch-types]. The CLI `batch` command pre-validates unknown fields before deserializing and then replays the items through the shared batch dispatcher described in [Batch Execution](../architecture/cli/batch-execution) and [Batch Item Protocol](../concepts/batch-item-protocol) [@batch-command].

## Field Table

| Field | Type | Meaning |
| --- | --- | --- |
| `command` | string | Canonical verb field. Use a bare verb such as `add`, `set`, `remove`, `move`, `swap`, `get`, `query`, `view`, `raw`, `raw-set`, `add-part`, or `validate` [@batch-types] [@batch-dispatch]. |
| `op` | string | Alias for `command` during deserialization [@batch-types]. |
| `path` | string | Target path for `get`, `set`, `remove`, `move`, `raw` fallback defaults, and other path-oriented verbs; `query` also accepts it as a selector alias in the dispatcher [@batch-types] [@batch-dispatch]. |
| `parent` | string | Parent path for `add` and `add-part` [@batch-types] [@batch-dispatch]. |
| `type` | string | Element or part type for `add` and `add-part` [@batch-types] [@batch-dispatch]. |
| `from` | string | Source path for copy-style `add` operations [@batch-types]. |
| `index` | integer | Insert or move index; it is serialized to resident request args as a string [@batch-types]. |
| `after` | string | Relative insertion or move anchor after a path [@batch-types]. |
| `before` | string | Relative insertion or move anchor before a path [@batch-types]. |
| `to` | string | Move destination parent and legacy second path for `swap` [@batch-types]. |
| `path2` | string | Canonical second path for `swap`; both `path2` and legacy `to` are accepted [@batch-types]. |
| `props` | object or string array | Property dictionary. Object values may be string, number, boolean, or null; array form must contain `key=value` strings and splits on the first equals sign [@batch-types]. |
| `selector` | string | Selector for `query` and selector-based operations [@batch-types]. |
| `text` | string | Text argument carried into the resident request argument map [@batch-types]. |
| `mode` | string | View mode or command-specific mode value [@batch-types]. |
| `depth` | integer | Depth for node reads; it is serialized to resident request args as a string [@batch-types]. |
| `part` | string | Raw OOXML part path for `raw` and `raw-set` [@batch-types] [@batch-dispatch]. |
| `xpath` | string | XPath target for `raw-set` [@batch-types] [@batch-dispatch]. |
| `action` | string | Raw-set action such as append, prepend, replace, remove, or setattr [@batch-types] [@batch-dispatch]. |
| `xml` | string | XML fragment or attribute payload for `raw-set` [@batch-types] [@batch-dispatch]. |

## Known Fields And Unknown Fields

`BatchItem.KnownFields` contains exactly `command`, `op`, `path`, `parent`, `type`, `from`, `index`, `after`, `before`, `to`, `path2`, `props`, `selector`, `text`, `mode`, `depth`, `part`, `xpath`, `action`, and `xml` [@batch-types]. The converter itself skips unknown JSON properties, but the CLI batch command scans each object first and rejects unknown field names with the valid-field list [@batch-types] [@batch-command].

## Props Conversion

`props` may be an object or an array. Object values are converted into strings: strings stay strings, numbers become invariant string values, booleans become `true` or `false`, and null becomes an empty string [@batch-types]. Array form accepts only strings and splits `key=value` on the first equals sign; malformed array entries without a key are skipped the same way the single-command prop parser does [@batch-types].

## Resident Request Mapping

`BatchItem.ToResidentRequest()` copies `Command` to `ResidentRequest.Command`, copies scalar fields into `ResidentRequest.Args`, and assigns `Props` to `ResidentRequest.Props` [@batch-types]. Integer fields such as `index` and `depth` are converted to strings in the args map, matching the resident pipe's stringly command argument model [@batch-types].

## Result Fields

Each `BatchResult` has `index`, `success`, optional `output`, optional `error`, and optional `item` [@batch-types]. Failed results include the original `item` so a caller can inspect and retry the failed command [@batch-types]. When `output` is valid JSON whose root is an object or array, the result converter writes it as raw JSON instead of a double-encoded string [@batch-types].

## Input Envelope Rules

The CLI batch command accepts an array from `--commands`, `--input`, explicit `--input -`, or stdin, and `--commands` and `--input` are mutually exclusive [@batch-command]. It also unwraps a normal JSON envelope whose root has a `data` array, so `dump --json` output can feed batch directly [@batch-command]. Non-array roots fail with a message telling the caller to wrap a single item in an array, and explicit null items are rejected by index [@batch-command].
