# Directional PPT Transitions

Three files work together:

- **transitions-directional.sh** — Build script.
- **transitions-directional.pptx** — 25-slide deck.
- **transitions-directional.md** — This file.

## Regenerate

```bash
cd examples/ppt/transitions
bash transitions-directional.sh
# → transitions-directional.pptx
```

## Family / direction matrix

Direction support is **not uniform**. Mismatching the family triggers
an error (`Invalid slide direction: 'leftup'. Valid values: left, right, up, down.`):

| Family | Direction set |
|---|---|
| `push` | `up` / `down` / `left` / `right` (4 cardinal) |
| `wipe` | `up` / `down` / `left` / `right` (4 cardinal) |
| `cover` | 4 cardinal **plus** `leftup` / `rightup` / `leftdown` / `rightdown` (8 total) |
| `uncover` | 8 directions, same set as `cover` |
| `pull` | Alias for `uncover` — same XML, same canonical readback |

## Combined-token shorthand

The combined input syntax is `TYPE[-DIR][-SPEED|DUR]`:

```bash
officecli set deck.pptx /slide[2] --prop transition=push-right
officecli set deck.pptx /slide[2] --prop transition=cover-leftdown
officecli set deck.pptx /slide[2] --prop transition=wipe-up-slow
officecli set deck.pptx /slide[2] --prop transition=push-right-1500
```

## Canonical readback

`Get` returns the canonical full-word form: `push-right`, not `push-r`.
Earlier versions of officecli leaked the raw OOXML single-letter dir
attribute for p14 transitions (e.g. `pan-up` → `pan-u`); the read path
now expands the abbreviation so both branches speak full words.

## Aliases (input only)

- `pull` accepts but reads back as `uncover`
- Single-letter dir abbreviations (`l`/`r`/`u`/`d`) accepted on input
