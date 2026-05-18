# Morph PPT Transition (Office 2016+)

Three files work together:

- **transitions-morph.sh** — Build script.
- **transitions-morph.pptx** — 4-slide deck.
- **transitions-morph.md** — This file.

## Regenerate

```bash
cd examples/ppt/transitions
bash transitions-morph.sh
# → transitions-morph.pptx
```

## What Morph does

Unlike every other transition (which animates the *replacement* of one
slide by the next), Morph **tweens** content between adjacent slides
that share named objects. A shape with the same `name=` on slide N and
slide N+1, but with different x/y/width/height/rotation/fill, smoothly
interpolates from the slide-N position to the slide-N+1 position when
the transition plays.

## Three granularities

```bash
officecli set deck.pptx /slide[N] --prop transition=morph             # default: byobject
officecli set deck.pptx /slide[N] --prop transition=morph-byword
officecli set deck.pptx /slide[N] --prop transition=morph-bychar
```

| Option | Tweens at the level of… |
|---|---|
| `byobject` (default) | Whole shape pairs (matched by name/id) |
| `byword` | Whole words within text bodies |
| `bychar` | Individual characters within text bodies |

`object`, `word`, `char`, `character` are accepted input aliases that
canonicalize to the OOXML-spelling on readback.

## How shape pairing works

In this trio the script creates a shape `name=morphBall` on every
slide. Same name → PowerPoint pairs them across slides and animates
the geometry change as continuous motion. Without matching names,
shapes are treated as independent and fade in/out instead of tweening.

```bash
# Slide 1 — small yellow ball, bottom-left
officecli add deck.pptx /slide[1] --type shape \
    --prop shape=ellipse --prop name=morphBall \
    --prop x=2cm --prop y=14cm --prop width=3cm --prop height=3cm

# Slide 2 — same name, larger and centered → ball grows and moves
officecli set deck.pptx /slide[2] --prop transition=morph
officecli add deck.pptx /slide[2] --type shape \
    --prop shape=ellipse --prop name=morphBall \
    --prop x=15cm --prop y=10cm --prop width=6cm --prop height=6cm
```

## OOXML representation

Morph is written inside an `mc:AlternateContent` wrapper using the
2015 p159 namespace:

```xml
<mc:AlternateContent>
  <mc:Choice Requires="p159">
    <p:transition>
      <p159:morph option="byObject"/>
    </p:transition>
  </mc:Choice>
  <mc:Fallback>
    <p:transition><p:fade/></p:transition>
  </mc:Fallback>
</mc:AlternateContent>
```

Older PowerPoint (pre-2016) plays the fallback fade — the deck remains
openable everywhere.

## See also

- `examples/product_launch_morph.pptx` in the repo root — a full
  product-launch deck built with morph as the primary motion.
- [transitions-dynamic.md](transitions-dynamic.md) — Office 2010+ "Exciting" gallery, which uses a similar mc:AlternateContent wrapper around p14 transitions.
