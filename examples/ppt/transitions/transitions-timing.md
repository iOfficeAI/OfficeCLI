# Transition Timing — Speed, Duration, and Advance Knobs

Three files work together:

- **transitions-timing.sh** — Build script.
- **transitions-timing.pptx** — 9-slide deck.
- **transitions-timing.md** — This file.

## Regenerate

```bash
cd examples/ppt/transitions
bash transitions-timing.sh
# → transitions-timing.pptx
```

## Four knobs

### 1. Legacy speed token (PowerPoint 97+)

```bash
officecli set deck.pptx /slide[N] --prop transition=fade-fast
officecli set deck.pptx /slide[N] --prop transition=fade-med    # or 'medium'
officecli set deck.pptx /slide[N] --prop transition=fade-slow
```

Stored as `<p:transition spd="fast|med|slow">`. `Get` surfaces it as
the read-only `transitionSpeed` format key.

### 2. Office 2010+ duration in milliseconds

```bash
officecli set deck.pptx /slide[N] --prop transition=fade-500     # 0.5 s
officecli set deck.pptx /slide[N] --prop transition=fade-1500    # 1.5 s
officecli set deck.pptx /slide[N] --prop transition=fade-3000    # 3.0 s
```

Stored as `<p:transition p14:dur="1500">` inside the
`mc:AlternateContent` wrapper. `Get` surfaces it as the read-only
`transitionDuration` format key.

Specifying both speed and duration is allowed — newer PowerPoint
honors `@dur`, older falls back to `@spd`.

### 3. Auto-advance after N milliseconds

```bash
officecli set deck.pptx /slide[N] --prop transition=fade --prop advanceTime=2000
# To clear later:
officecli set deck.pptx /slide[N] --prop advanceTime=none
```

When `advanceTime` is set, the slide moves on by itself after N ms in
Slide Show mode. The transition itself can be `cut` if you want an
instant auto-advance with no animation.

### 4. Disable click-to-advance

```bash
officecli set deck.pptx /slide[N] --prop transition=fade --prop advanceClick=false
```

`advanceClick` defaults to true (and the XML attribute is stripped
when true — only the explicit `false` survives a round-trip). With
`advanceClick=false`, the slide ignores click/Enter advance; the user
must use arrow keys or wait for `advanceTime`.

## Round-trip semantics

- `advanceClick=true` → XML attribute stripped → readback emits no
  `advanceClick` key (default is true).
- `advanceClick=false` → XML keeps `advClick="0"` → readback emits
  `advanceClick=false`.
- `advanceTime=none` → XML attribute removed → readback emits no
  `advanceTime` key.

## Related

- [transitions-basic.md](transitions-basic.md) — uses `transition=none` to remove the entire transition element.
