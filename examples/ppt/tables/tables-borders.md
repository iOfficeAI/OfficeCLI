# PPT Table Borders

Three files work together:

- **tables-borders.sh** — Build script.
- **tables-borders.pptx** — 3-slide deck.
- **tables-borders.md** — This file.

## Regenerate

```bash
cd examples/ppt
bash tables-borders.sh
# → tables-borders.pptx
```

## Border format

PPT accepts two equivalent formats. Pick whichever reads best for your case:

| Form | Example | Notes |
|---|---|---|
| **Space-separated** | `"1pt solid FF0000"` | `"WIDTH[ DASH][ COLOR]"`, WIDTH in pt |
| **Semicolon** | `"single;4;FF0000"` | `"STYLE;WIDTH;COLOR[;DASH]"`, WIDTH in 1/8 pt — matches docx |
| **Semicolon + dash** | `"single;8;0070C0;dash"` | optional dash on the end |
| **Clear** | `"none"` | removes the border |

`DASH ∈ solid | dot | dash | lgDash | dashDot | sysDot | sysDash`.
STYLE in the semicolon form is ignored by pptx (kept for docx compatibility).

## Slides

### Slide 1 — Shorthand & per-edge

Six small tables demonstrating:

- `border.all="1pt solid 808080"` — grey grid on every edge
- `border.all="2pt solid FF0000"` — thick red on every edge
- `border.all=none` — clear borders (table is invisible)
- `border.top` / `border.bottom` / `border.left` — outer edges only

### Slide 2 — Inside dividers & dash patterns

- `border.horizontal` (alias `border.insideH`) — between rows
- `border.vertical` (alias `border.insideV`) — between columns
- Dash variations: `lgDash`, `dashDot`, `sysDash`

```bash
officecli add file.pptx /slide[2] --type table \
  --prop border.horizontal="1pt solid CCCCCC" \
  --prop border.all="1pt solid 404040" \
  --prop data="A,B,C;1,2,3;..."
```

### Slide 3 — Diagonal borders

Per-cell `border.tl2br` / `border.tr2bl` for crossed-out headers, N/A cells,
or matrix corners:

```bash
# Header corner cell with a diagonal split
officecli set file.pptx /slide[3]/table[1]/tr[1]/tc[1] \
  --prop border.tl2br="1pt solid 808080"

# N/A cell with both diagonals (X)
officecli set file.pptx /slide[3]/table[2]/tr[1]/tc[1] \
  --prop text="N/A" \
  --prop border.tl2br="1pt solid C00000" \
  --prop border.tr2bl="1pt solid C00000"
```

Diagonal borders are **add/set only** — `get` does not surface them today.

**Features:** `border.all`, `border.top/right/bottom/left`,
`border.horizontal`/`border.vertical`, `border.tl2br`/`border.tr2bl`,
dash patterns.
