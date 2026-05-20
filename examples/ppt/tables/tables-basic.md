# Basic PPT Tables

Three files work together:

- **tables-basic.sh** — Shell script that calls `officecli` to build the deck.
- **tables-basic.pptx** — The generated 3-slide deck.
- **tables-basic.md** — This file.

## Regenerate

```bash
cd examples/ppt
bash tables-basic.sh
# → tables-basic.pptx
```

## Slides

### Slide 1 — Inline `data=` seed

Whole table populated in a single command with `data="H1,H2;R1C1,R1C2"`
(commas separate cells, semicolons separate rows).

```bash
officecli add file.pptx /slide[1] --type table \
  --prop x=0.5in --prop y=1.2in --prop width=12in --prop height=2in \
  --prop headerFill=4472C4 --prop bodyFill=DEEAF6 \
  --prop data="Region,Q1,Q2,Q3,Q4;North,120,135,142,168;South,98,110,121,140;East,165,178,190,205"
```

> ⚠ `headerFill` / `bodyFill` are a **per-cell stamp** applied at table
> creation, not a table-level property. If you later run `add row` or
> `add column`, the new cells will not auto-color — you have to set
> their `fill` explicitly. Want appended rows/columns to follow the
> coloring automatically? Use a theme style instead:
> `--prop style=medium2 --prop firstRow=true --prop bandedRows=true`.
> See [tables-rows-cols.md](tables-rows-cols.md) for the side-by-side
> comparison.

### Slide 2 — Empty grid + per-cell `set`

Reserve the grid with `rows`/`cols`, then set each cell. Useful when cell
values aren't known up-front, or different cells need different styling.

```bash
officecli add file.pptx /slide[2] --type table \
  --prop rows=4 --prop cols=3 --prop headerFill=2E75B6

officecli set file.pptx /slide[2]/table[1]/tr[1]/tc[1] \
  --prop text="Product" --prop bold=true --prop color=FFFFFF
```

### Slide 3 — Cell fill variations

`fill` (alias `background`/`shd`) accepts several forms:

| Form | Example |
|---|---|
| Solid hex | `fill=FF0000` or `fill=#FF0000` |
| Named color | `fill=red` |
| `rgb(...)` | `fill="rgb(255,0,0)"` |
| Theme color | `fill=accent1` (also `accent2..6`, `dk1`, `dk2`, `lt1`, `lt2`, `hyperlink`) |
| Gradient | `fill="FF0000-0000FF-90"` — `"COLOR1-COLOR2[-ANGLE]"`, angle in degrees |
| No fill | `fill=none` — transparent (page bg shows through) |

Theme colors follow the deck theme — recolor the deck and the table follows.
Hex/named colors are absolute.

**Features:** `data=` inline seed, `headerFill`/`bodyFill`, `rows`/`cols`,
per-cell `text`/`bold`/`color`/`fill` (solid/named/rgb/theme/gradient/none),
EMU-parseable dimensions (`0.5in`).
