# Shape-Mask PPT Transitions

Three files work together:

- **transitions-shapes.sh** — Build script.
- **transitions-shapes.pptx** — 12-slide deck.
- **transitions-shapes.md** — This file.

## Regenerate

```bash
cd examples/ppt/transitions
bash transitions-shapes.sh
# → transitions-shapes.pptx
```

## The three sub-families

Shape-mask transitions reveal the new slide through a growing geometric
mask. They split into three flavors by what modifies the shape:

### 1. Direction-less (no `-in`/`-out` suffix)

```bash
officecli set deck.pptx /slide[N] --prop transition=circle
officecli set deck.pptx /slide[N] --prop transition=diamond
officecli set deck.pptx /slide[N] --prop transition=plus
officecli set deck.pptx /slide[N] --prop transition=wedge
```

These four are direction-less. Passing `-in`/`-out` is rejected with a
clear error rather than silently dropped:

```
Error: Transition 'circle' does not accept a direction modifier (got '-in').
'circle' is a direction-less shape transition — drop the suffix and
use plain 'transition=circle'.
```

### 2. In / Out

```bash
officecli set deck.pptx /slide[N] --prop transition=zoom-in
officecli set deck.pptx /slide[N] --prop transition=zoom-out
officecli set deck.pptx /slide[N] --prop transition=box-in
officecli set deck.pptx /slide[N] --prop transition=box-out
```

The default is `-in`; bare `zoom` / `box` round-trip as `zoom` / `box`
(default collapses on readback), `zoom-out` / `box-out` round-trip with
the suffix intact.

**Box is a PowerPoint 2013+ "modern" transition.** officecli writes
it with an inline fade fallback baked in, so pre-2013 PowerPoint plays
a fade in its place instead of nothing.

### 3. Spoke count — `wheel-N`

```bash
officecli set deck.pptx /slide[N] --prop transition=wheel       # 4 spokes (default)
officecli set deck.pptx /slide[N] --prop transition=wheel-1     # 1 spoke
officecli set deck.pptx /slide[N] --prop transition=wheel-8     # 8 spokes
```

The integer suffix (1..32) is the spoke count, not a duration. To set
both spokes and duration: combine — `wheel-8-1500` writes 8 spokes +
1500 ms duration. Readback returns `wheel-N` for non-default counts;
`wheel-4` collapses to bare `wheel`.

## Related

- [transitions-bands.md](transitions-bands.md) for split-vertical-in / split-horizontal-out (similar in/out modifier).
- [transitions-basic.md](transitions-basic.md) for the no-shape baseline (cut/fade/dissolve).
