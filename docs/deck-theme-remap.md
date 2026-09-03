# `officecli deck theme-remap`

Change a WorkMate DeckSpec `theme.id`, report slides whose layouts may need a same-role remap, and suggest alternatives via `layout-query`. Original CSBU WorkMate flow — token identity only (no AGPL / third-party theme grids).

## Usage

```bash
# Dry-run report (JSON)
officecli deck theme-remap deck.workmate-deck.json --to csbu-workmate --json

# Apply theme + embed report under extensions.themeRemap
officecli deck theme-remap deck.workmate-deck.json --to csbu-workmate-night --apply --json

# Apply to a copy
officecli deck theme-remap deck.workmate-deck.json --to boardroom-navy --apply -o remapped.workmate-deck.json --json
```

### Options

| Flag | Description |
|------|-------------|
| `--to` / `--theme` | Target catalog `theme.id` (required) |
| `--apply` | Write `theme.id` (and optional report) back to the spec |
| `--write-report` | When applying, embed report at `extensions.themeRemap` (default true) |
| `--set-mode` | Set `theme.mode` from target background luminance (default true) |
| `--limit` | Max same-role alternatives per slide (1–20, default 5) |
| `--output` / `-o` | Output path when applying (default: in-place) |
| `--json` | JSON stdout |

`needsRemap` is true when the layout is unknown, a clearly better same-role alternative exists (layout-query score delta ≥ 2), or a light↔dark mode shift leaves the current layout outside the top ranked set.

## Agent / Studio workflow

1. Prefer brand themes `csbu-workmate` / `csbu-workmate-night` for CSBU WorkMate token identity.
2. Dry-run `theme-remap --to <id>` and review `needsRemap` + alternatives.
3. Apply with `--apply`; optionally switch layouts / pin `candidates` from alternatives.
4. Validate with `officecli deck validate`.
