# Band-Pattern PPT Transitions

Three files work together:

- **transitions-bands.sh** — Build script.
- **transitions-bands.pptx** — 21-slide deck.
- **transitions-bands.md** — This file.

## Regenerate

```bash
cd examples/ppt/transitions
bash transitions-bands.sh
# → transitions-bands.pptx
```

## Three pattern flavors

### Orientation modifier (-horizontal / -vertical)

```bash
officecli set deck.pptx /slide[N] --prop transition=blinds-vertical
officecli set deck.pptx /slide[N] --prop transition=checker-horizontal
officecli set deck.pptx /slide[N] --prop transition=comb-vertical
officecli set deck.pptx /slide[N] --prop transition=bars-horizontal
```

The default is `horizontal`; explicit writes round-trip as
`blinds-horizontal` even when matching the default (the readback path
preserves the explicit form, matching the wipe/push convention).

### Corner direction (-leftup / -rightup / -leftdown / -rightdown)

```bash
officecli set deck.pptx /slide[N] --prop transition=strips-leftup
officecli set deck.pptx /slide[N] --prop transition=strips-rightdown   # default
```

### Orient × in/out (split needs BOTH)

```bash
officecli set deck.pptx /slide[N] --prop transition=split-vertical-in
officecli set deck.pptx /slide[N] --prop transition=split-horizontal-out
```

`split-vertical` (orient without in/out) defaults `dir=in`, so it
round-trips canonically as `split-vertical-in`. Bare `split` (no
orientation given) reads back as plain `split`.

## Aliases (input only, canonicalize on readback)

| Input alias | Canonical readback |
|---|---|
| `venetian` | `blinds` |
| `checkerboard` | `checker` |
| `randombar` | `bars` |
| `diagonal` | `strips` |

Pick whichever spelling reads best in your script; the OOXML is identical
either way.

## Related

- [transitions-shapes.md](transitions-shapes.md) — circle/diamond/wheel (the non-banded geometric family).
- [transitions-directional.md](transitions-directional.md) — push/cover/wipe (cardinal-direction transitions).
