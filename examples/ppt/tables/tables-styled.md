# Styled PPT Tables (Built-in Styles & Banding)

Three files work together:

- **tables-styled.sh** — Build script.
- **tables-styled.pptx** — 11-slide deck (9 styles + 1 banding combo slide + 1 `rowHeight`/`name=` slide).
- **tables-styled.md** — This file.

## Regenerate

```bash
cd examples/ppt
bash tables-styled.sh
# → tables-styled.pptx
```

## Slides

### Slides 1–9 — Each built-in style

PowerPoint ships 9 named theme styles. One slide per style:

| Slide | `--prop style=` |
|------:|-----------------|
| 1 | `medium1` |
| 2 | `medium2` (default) |
| 3 | `medium3` |
| 4 | `medium4` |
| 5 | `light1` |
| 6 | `light2` |
| 7 | `light3` |
| 8 | `dark1` |
| 9 | `dark2` |

```bash
officecli add file.pptx /slide[1] --type table \
  --prop style=medium2 \
  --prop firstRow=true --prop bandedRows=true \
  --prop data="Region,Q1,Q2,Q3,Q4;North,120,135,142,168;..."
```

### Slide 10 — Banding flag combinations

Four tables on one slide showing how the banding flags interact with a style:

- `firstRow=true --prop bandedRows=true` — emphasized header + zebra rows
- `firstCol=true --prop bandedCols=true` — emphasized first column + zebra columns
- `firstRow=true --prop lastRow=true` — emphasized header *and* totals row
- `style=none` — no theme; pair with explicit `border.all` for visible grid

### Slide 11 — `rowHeight` + `name=` addressing

Two table-level props for ergonomics:

- `rowHeight=1cm` — stamps every row with the same height at create time
  (otherwise officecli derives row height from `height / rows`).
- `name=SalesData` — sets the table's `NonVisualDrawingProperties Name`.
  After creation, the table can be addressed by name instead of by
  positional index, which survives slide reordering:

```bash
officecli add file.pptx /slide[11] --type table \
  --prop name=SalesData --prop rowHeight=1cm \
  --prop data="..."

# Stable path — works even if more tables are added before this one.
officecli set file.pptx '/slide[11]/table[@name=SalesData]/tr[2]/tc[2]' \
  --prop text="120 ▲" --prop bold=true --prop fill=C6E0B4
```

`@name=` / `@id=` addressing is also accepted on `get`, `query`, `remove`.

**Features:** `style=medium1..4|light1..3|dark1..2|none`,
`firstRow`, `lastRow`, `firstCol`, `lastCol`, `bandedRows`, `bandedCols`,
`rowHeight`, `name` + `/table[@name=…]` addressing.
