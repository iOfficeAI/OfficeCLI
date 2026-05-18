# Basic PPT Transitions

Three files work together:

- **transitions-basic.sh** — Shell script that calls `officecli` to build the deck.
- **transitions-basic.pptx** — The generated 6-slide deck.
- **transitions-basic.md** — This file.

## Regenerate

```bash
cd examples/ppt/transitions
bash transitions-basic.sh
# → transitions-basic.pptx
```

## What this trio shows

The five "everyday" transition tokens, plus the canonical clear form.
Transitions on slide N animate the entry of slide N (replacing slide N-1).
Step through the deck in PowerPoint's Slide Show mode to see the actual
motion — the static .pptx looks identical for all of them.

| Slide | Transition | Effect in playback |
|---|---|---|
| 1 | (none — entry slide) | No transition; baseline |
| 2 | `cut` | Instant swap, no animation |
| 3 | `fade` | Pixel cross-fade |
| 4 | `dissolve` | Speckle-pattern blend |
| 5 | `flash` | Quick white flash through |
| 6 | `none` cleared | Was `fade`, then `transition=none` removed it |

## How `transition=none` clears

```bash
officecli set deck.pptx /slide[6] --prop transition=fade
officecli set deck.pptx /slide[6] --prop transition=none
```

After the second call, `Get` returns no `transition` key — the OOXML
`<p:transition>` element is removed entirely (including the
`mc:AlternateContent` wrapper that some Office 2010+ types use). Use
this to remove a stale transition rather than guessing at an empty
string.

## Related trios

- [transitions-directional.md](transitions-directional.md) — push / cover / wipe with direction
- [transitions-shapes.md](transitions-shapes.md) — circle / diamond / wheel / zoom
- [transitions-bands.md](transitions-bands.md) — blinds / strips / split / checker
- [transitions-dynamic.md](transitions-dynamic.md) — Office 2010+ "Exciting" gallery
- [transitions-random.md](transitions-random.md) — newsflash / random
- [transitions-timing.md](transitions-timing.md) — speed, duration, advance
- [transitions-morph.md](transitions-morph.md) — Morph (2016+)
