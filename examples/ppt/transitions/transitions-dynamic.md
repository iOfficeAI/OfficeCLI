# Dynamic / 3D PPT Transitions (Office 2010+ "Exciting")

Three files work together:

- **transitions-dynamic.sh** — Build script.
- **transitions-dynamic.pptx** — 24-slide deck.
- **transitions-dynamic.md** — This file.

## Regenerate

```bash
cd examples/ppt/transitions
bash transitions-dynamic.sh
# → transitions-dynamic.pptx
```

## Why these are special

These transitions ship in PowerPoint 2010 or later. officecli writes
each one with an inline fade fallback baked in, so a pre-2010
PowerPoint opening the same deck plays a plain fade in their place
instead of failing or showing nothing.

## Direction grouping

| Family | Direction set | Example |
|---|---|---|
| LeftRight | `left` / `right` | `switch-right`, `flip-right`, `ferris-right`, `gallery-right`, `conveyor-right`, `reveal-right` |
| InOut | `in` / `out` | `shred-out`, `flythrough-out`, `warp-out` |
| SlideDir (4 cardinal) | `up` / `down` / `left` / `right` | `vortex-up`, `glitter-right`, `pan-up`, `prism-right` |
| Orientation | `horizontal` / `vertical` | `doors-vertical`, `window-horizontal` |
| (direction-less) | — | `ripple`, `honeycomb` |

## Combined-token shorthand

```bash
officecli set deck.pptx /slide[N] --prop transition=switch-right
officecli set deck.pptx /slide[N] --prop transition=shred-out-slow
officecli set deck.pptx /slide[N] --prop transition=ferris-right-1500
```

## Recent fixes pinned by this trio

- `reveal-right`, `ferris-right`, `gallery-right`, `conveyor-right`,
  `shred-out`, `flythrough-out`, `warp-out` — direction was silently
  dropped at write time and lost on readback; each now round-trips.
- `pan-up` previously read back as the truncated `pan-u`; single-letter
  abbreviations now always expand to full words on readback.

## Related

- [transitions-basic.md](transitions-basic.md) — Office 97-era cut/fade/dissolve.
- [transitions-morph.md](transitions-morph.md) — Office 2016+ Morph (separate code path).
