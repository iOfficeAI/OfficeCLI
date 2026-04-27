# schemas/

Agent-facing capability schemas for officecli. Single source of truth for what the CLI supports, consumed in three places:

1. **`officecli <format> <op> <element> --help --json`** — runtime output for agents. Schemas are embedded into the binary at build time, so runtime does not depend on filesystem paths or network access.
2. **Contract tests** — every schema claim (`add`, `set`, `get`, `readback`) is verified against the real handler implementation. Properties marked `enforcement: strict` break CI on drift; `report` only log.
3. **Release-time wiki generation** (future) — wiki markdown is generated/diffed from schemas before publishing. During development, wiki is not touched; agents read schemas directly.

## Layout

```
schemas/
  help/
    _schema.json                ← JSON Schema (draft 2020-12) describing the format below
    docx/<element>.json         ← Word per-element capability
    pptx/<element>.json         ← PowerPoint per-element capability
    xlsx/<element>.json         ← Excel per-element capability
```

## Editing rule

Any PR that changes `Add`, `Set`, or `Get` behavior for an element **must** update the matching schema file in the same PR. CI contract tests will fail otherwise.

## Not here

- Narrative / tutorials / best practices → wiki (generated or hand-written at release time).
- Internal implementation notes → CLAUDE.md and code comments.
- Ephemeral release notes → CHANGELOG.
