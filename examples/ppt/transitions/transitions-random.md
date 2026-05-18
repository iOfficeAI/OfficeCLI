# Random PPT Transitions

Three files work together:

- **transitions-random.sh** — Build script.
- **transitions-random.pptx** — 4-slide deck.
- **transitions-random.md** — This file.

## Regenerate

```bash
cd examples/ppt/transitions
bash transitions-random.sh
# → transitions-random.pptx
```

## What this trio shows

```bash
officecli set deck.pptx /slide[N] --prop transition=newsflash
officecli set deck.pptx /slide[N] --prop transition=random
```

- **`newsflash`** — a fixed legacy animation: the new slide spins
  inward newspaper-style. Pre-2010 OOXML element; ships in plain
  `<p:transition><p:newsflash/></p:transition>` (no mc:AlternateContent
  wrapper needed).
- **`random`** — at render time, PowerPoint picks a random transition
  from its available set. The .pptx only captures the *intent*; the
  motion you see in Slide Show mode will differ each time you enter
  presentation mode, even for the same slide.

To experience the difference: open `transitions-random.pptx` in
PowerPoint, run Slide Show, exit, run Slide Show again — slides 3 and
4 should animate differently each pass.

## Related

- [transitions-basic.md](transitions-basic.md) — for the deterministic counterparts (cut/fade/dissolve/flash).
