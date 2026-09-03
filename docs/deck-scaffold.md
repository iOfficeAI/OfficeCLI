# `officecli deck scaffold`

Outline a long WorkMate deck from goal / audience / page count with section transitions and role-mix heuristics. Original CSBU WorkMate flow — emits `stage: "outline"` slides with empty blocks and optional same-role `candidates[]`. Not a third-party goal-spec clone; no AGPL / HTML deck runtime.

## Usage

```bash
officecli deck scaffold \
  --goal "Q3 business review: growth, risks, asks" \
  --audience "exec leadership" \
  --pages 20 \
  --theme csbu-workmate \
  --language zh-CN \
  --seed q3-2026 \
  -o q3-outline.workmate-deck.json \
  --json
```

### Options

| Flag | Description |
|------|-------------|
| `--goal` | Deck goal / one-line intent (biases role mix) |
| `--audience` | Primary audience (biases role mix) |
| `--pages` / `--page-count` | Target slide count (4–60, default 12) |
| `--title` | Deck title (defaults from goal) |
| `--language` / `--lang` | BCP-47 tag (default `en-US`) |
| `--theme` | Catalog `theme.id` (default `csbu-workmate`) |
| `--seed` | Reproducibility seed (default: hash of goal\|audience\|pages) |
| `--output` / `-o` | Output `*.workmate-deck.json` (required) |
| `--write` | Write the outline DeckSpec (default true) |
| `--json` | JSON stdout |

Stdout includes `report.roleMix`, `report.sectionBreaks`, and `report.seed`. The same report is embedded at `extensions.deckScaffold`.

## Heuristics (WorkMate-original)

1. Always start with `cover`; end with `closing` (and `actions` when pages ≥ 10).
2. Long decks (pages ≥ 8) insert an agenda/`breakdown` near the front.
3. Middle pages draw from a role pool biased by goal/audience keywords (metrics, trend, comparison, process, risks, team, case, actions, …).
4. Decks ≥ 10 pages insert `transition` section breaks about every 5–7 content slides.
5. Same-role layout `candidates[]` are pinned when the catalog has alternatives — Studio chips / wireframe compare can switch without rewriting narrative.

## Agent / Studio workflow

1. Run `deck scaffold` (or follow the workmate-presentation skill long-deck outline) with confirmed goal/audience/page count.
2. Open the outline in Presentation Studio; confirm titles / theme via Outline Gate.
3. Optionally compare same-role candidates via Studio wireframe compare, then fill blocks.
4. `officecli deck validate` before export.
