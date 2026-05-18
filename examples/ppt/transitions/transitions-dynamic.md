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
each one inside an `mc:AlternateContent` wrapper with an inline
`<mc:Fallback><p:transition><p:fade/></p:transition></mc:Fallback>` —
older PowerPoint that doesn't recognize the p14 namespace will play a
plain fade instead, so the deck still opens.

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
  `shred-out`, `flythrough-out`, `warp-out` — earlier the parser called
  the SDK constructor with no direction, so the resulting
  `<p14:reveal/>` element had no `dir` attribute and read back as bare
  `reveal`. Each type now passes through `ParseLeftRightDir` or
  `ParseInOutDir` and round-trips correctly.
- `pan-up` previously read back as `pan-u` because the regex-based p14
  readback echoed the raw single-letter attribute. The readback now
  expands abbreviations to canonical full words.

## Related

- [transitions-basic.md](transitions-basic.md) — Office 97-era cut/fade/dissolve.
- [transitions-morph.md](transitions-morph.md) — Office 2016+ Morph (separate code path).
