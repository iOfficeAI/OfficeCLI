# rhwp Sidecar Contract

This contract defines the stable boundary between OfficeCLI and the experimental
rhwp provider. The current implementation still invokes concrete sidecar
commands, but new work must keep responses compatible with the request/response
schemas in `schemas/interfaces/`.

## Provider Identity

Every sidecar response must expose:

- `schemaVersion`
- `operation`
- `format`
- `engineVersion`
- `warnings`
- either `data` or `error`

The provider must include the rhwp version or pinned commit whenever available.

## Request Shape

Requests are described by:

```text
schemas/interfaces/rhwp-sidecar-request.v1.schema.json
```

Required fields:

- `schemaVersion`
- `operation`
- `format`
- `inputPath`

Mutating operations must also include `outputPath`.

## Response Shape

Responses are described by:

```text
schemas/interfaces/rhwp-sidecar-response.v1.schema.json
```

Required fields:

- `schemaVersion`
- `ok`
- `operation`
- `format`
- `engineVersion`

## Operation Policy

Supported HWP operations:

- `read-text`
- `render-svg`
- `list-fields`
- `read-field`
- `fill-field`
- `replace-text`
- `table-map`
- `set-table-cell`

Supported HWPX rhwp operations:

- `read-text`
- `render-svg`
- `list-fields`
- `read-field`
- `fill-field`
- `replace-text`

Blocked operations must return typed errors instead of silent fallback.

## Save Policy

The sidecar must not overwrite the source path. Mutating operations write to an
explicit output path until safe-save transactions are implemented.

Future in-place support must go through:

1. temp output in the source directory;
2. provider readback;
3. semantic delta validation;
4. SVG/visual validation where available;
5. backup creation;
6. atomic replace.

## Error Policy

Errors must include:

- `code`
- `message`
- `format`
- `operation`
- `engine`
- `nextCommand` when there is an obvious diagnostic command

Examples:

```text
bridge_not_enabled
bridge_missing
unsupported_operation
roundtrip_unverified
binary_hwp_write_forbidden
```
