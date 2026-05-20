# Modern PPT Transitions (PowerPoint 2013+ "Exciting" Gallery)

Three files work together:

- **transitions-modern.sh** — Build script.
- **transitions-modern.pptx** — 19-slide deck.
- **transitions-modern.md** — This file.

## Regenerate

```bash
cd examples/ppt/transitions
bash transitions-modern.sh
# → transitions-modern.pptx
```

## What this trio shows

The 12 "Exciting" / "Dynamic Content" presets PowerPoint 2013 added to
the Transitions gallery. Each one has an inline fade fallback baked
in, so pre-2013 PowerPoint plays a graceful fade in their place
instead of nothing.

## The 12 presets

| CLI token | UI name | Notes |
|---|---|---|
| `fallOver` | Fall Over | direction-sensitive |
| `drape` | Drape | direction-sensitive |
| `curtains` | Curtains | symmetric |
| `wind` | Wind | direction-sensitive |
| `prestige` | Prestige | symmetric |
| `fracture` | Fracture | symmetric |
| `crush` | Crush | symmetric |
| `peelOff` | Peel Off | direction-sensitive |
| `pageCurlDouble` | Page Curl (double) | direction-sensitive |
| `pageCurlSingle` | Page Curl (single) | direction-sensitive |
| `airplane` | Airplane | direction-sensitive |
| `origami` | Origami | direction-sensitive |

Token spelling is lowerCamelCase. Input is case-insensitive
(`transition=PageCurlDouble` and `pagecurldouble` both work), but `Get`
returns the canonical lowerCamelCase form.

## Direction (-in / -out)

```bash
officecli set deck.pptx /slide[N] --prop transition=pageCurlDouble
officecli set deck.pptx /slide[N] --prop transition=pageCurlDouble-out
officecli set deck.pptx /slide[N] --prop transition=wind-out
```

- Default (no suffix) = `-in`.
- `-out` flips the Left/Right direction. Visually affects the
  direction-sensitive presets in the table above; symmetric presets
  (curtains, fracture, crush, prestige) parse the suffix but render
  unchanged.
- Any other direction is rejected:
  ```
  Error: Transition 'fallOver' only accepts -in or -out (got '-up').
  ```

## UI tiles backed by other tokens

A few PowerPoint UI tiles that look like they belong in this gallery
are actually wired through other CLI tokens — just write the right one:

| PowerPoint UI tile | CLI token |
|---|---|
| Cube (Exciting) | `prism` or `cube` |
| Rotate (Dynamic Content) | `rotate` |
| Orbit (Dynamic Content) | `orbit` |
| Clock (Exciting) | `wheel-1` or `clock` |

## See also

- [transitions-shapes.md](transitions-shapes.md) — Box lives there alongside circle/diamond/zoom.
- [transitions-dynamic.md](transitions-dynamic.md) — the older 2010 "Exciting" gallery (vortex / switch / flip / ferris / ... / prism / rotate / orbit).
- [transitions-morph.md](transitions-morph.md) — Morph (2016+).
