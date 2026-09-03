# `officecli deck layout-query`

Capacity-aware layout ranking for CSBU WorkMate DeckSpec slides. Scores layouts from the embedded catalog using role + content hints. Original WorkMate implementation (no AGPL upstream code).

## Usage

```bash
officecli deck layout-query \
  --role metrics \
  --item-count 4 \
  --has-chart false \
  --needs-media false \
  --limit 5 \
  --json
```

### Options

| Flag | Description |
|------|-------------|
| `--role` | Semantic role filter (`cover`, `metrics`, `trend`, `process`, …) |
| `--item-count` / `--module-count` | Approximate modules / KPIs / cards / steps needed |
| `--has-chart` | `true`/`false` — need a chart-capable slot |
| `--needs-media` / `--has-image` | `true`/`false` — need an image slot |
| `--has-table` | `true`/`false` — need a table-capable slot |
| `--query` | Free-text match against layout id / label / role |
| `--limit` | Max results (1–50, default 8) |
| `--json` | JSON stdout (recommended for agents) |

### Example output

```json
{
  "query": { "role": "metrics", "itemCount": 4, "limit": 5 },
  "catalogVersion": "1.4.0",
  "resultCount": 5,
  "results": [
    {
      "layoutId": "metrics-row-4",
      "role": "metrics",
      "score": 31.1,
      "capacity": 4,
      "accepts": ["metric", "text"],
      "reasons": ["role_match", "capacity_exact", "has_alternatives"]
    }
  ]
}
```

## Agent workflow

1. Confirm outline / role per slide.
2. Call `layout-query` with role + item/chart/media hints.
3. Pick `results[0].layoutId` (or set `candidates` from top-k and keep one as `layoutId`).
4. Fill blocks; run `officecli deck validate`.

Optional DeckSpec field (P1.6 start): `slides[].candidates` — array of preferred layout ids validated against the catalog. Export still uses `layoutId` only.
