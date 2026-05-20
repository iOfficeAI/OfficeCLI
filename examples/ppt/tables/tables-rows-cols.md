# PPT Table Rows & Columns

Three files work together:

- **tables-rows-cols.sh** — Build script.
- **tables-rows-cols.pptx** — 4-slide deck.
- **tables-rows-cols.md** — This file.

Covers the `row` and `column` child elements of `table`, which the other
`tables-*` examples leave as one-line static `rows`/`cols` props.

## Regenerate

```bash
cd examples/ppt
bash tables-rows-cols.sh
# → tables-rows-cols.pptx
```

## Slides

### Slide 1 — Grow an existing table (theme vs per-cell stamp)

This slide shows **two tables side by side** that are populated identically
but colored two different ways — the contrast is the point of the slide.

```bash
# Append a row. Inherits column count; seed cells via c{N}=value.
officecli add file.pptx /slide[1]/table[1] --type row \
  --prop c1=Bob --prop c2=95 --prop c3=110

# Append a column. Inserts a cell in EVERY existing row.
officecli add file.pptx /slide[1]/table[1] --type column \
  --prop width=2cm --prop text="Q3"
```

`text=` on `add column` seeds the same value into every cell of the new
column (useful for adding a placeholder "—" column). Per-row body values
go in via `set` on the newly created `tc[N]` cells.

### The two coloring models — pick one deliberately

PowerPoint tables (and Word tables, and Excel tables) have **two
independent ways** to color cells. They behave differently when you
later `add row` / `add column`:

| Model | How you write it | Stored as | Appended row/col follows? |
|---|---|---|---|
| **Theme (recommended for "auto-follow")** | `--prop style=medium2` (+ `firstRow`/`bandedRows`/…) | Table-level `<a:tableStyleId>` reference. Renderer paints all cells in range, including ones added later. | ✓ **Yes — automatic.** Same model as Excel Table styles. |
| **Per-cell stamp** | `--prop headerFill=4472C4 --prop bodyFill=DEEAF6` *(or any later `set tc[N] fill=…`)* | `<a:solidFill>` fanned out onto **each existing cell** at that moment. | ✗ **No — never.** Per-cell fills are personal overrides; they don't spread. |

This is not a bug or a gap — it's the OOXML model speaking through
officecli. The same separation exists for `<a:tableStyleId>` vs cell
`<a:solidFill>` in PPT, `<w:tblStyle>` vs cell `<w:shd>` in Word, and
`<x:tableStyleInfo>` vs cell-level fills in Excel.

**Rule of thumb:**

- Want appended rows/columns to look the same as the original? → use a
  **theme style** (`style=medium2|light1|dark1|…`).
- Need a specific custom color that's not in any theme? → use
  `headerFill`/`bodyFill` (or per-cell `fill=`), and accept that you'll
  manually fill new cells after `add row`/`add column`. Table B on
  slide 1 of the demo deck shows exactly this top-up:

  ```bash
  HDR=4472C4; BODY=DEEAF6
  # Total header added via 'add column' — needs headerFill manually
  officecli set file.pptx /slide[1]/table[2]/tr[1]/tc[4] \
    --prop fill=$HDR --prop color=FFFFFF --prop bold=true

  # Newly appended Bob/Carol rows — need bodyFill on each cell
  for c in 1 2 3 4; do
    officecli set file.pptx /slide[1]/table[2]/tr[3]/tc[$c] --prop fill=$BODY
    officecli set file.pptx /slide[1]/table[2]/tr[4]/tc[$c] --prop fill=$BODY
  done
  ```

### Slide 2 — Per-row height + per-column width

After a table is created, each row and column can have its own size:

```bash
# Custom column widths (must sum to roughly the table width)
officecli set file.pptx /slide[2]/table[1]/col[1] --prop width=2in
officecli set file.pptx /slide[2]/table[1]/col[2] --prop width=1.5in
officecli set file.pptx /slide[2]/table[1]/col[3] --prop width=7in
officecli set file.pptx /slide[2]/table[1]/col[4] --prop width=1.5in

# Custom row heights — header thin, body increasing
officecli set file.pptx /slide[2]/table[1]/tr[1] --prop height=0.5in
officecli set file.pptx /slide[2]/table[1]/tr[2] --prop height=0.6in
officecli set file.pptx /slide[2]/table[1]/tr[3] --prop height=1in
officecli set file.pptx /slide[2]/table[1]/tr[4] --prop height=1.5in
```

### Slide 3 — Uniform `rowHeight` (table-level)

When every row should be the same height, set `rowHeight` once at
`add table` time instead of running `set tr[N] height=` N times:

```bash
officecli add file.pptx /slide[3] --type table \
  --prop rows=5 --prop cols=3 --prop rowHeight=0.8in \
  --prop data="Step,Action,Result;1,Init,OK;..."
```

### Slide 4 — Cell merging: `gridSpan` (horizontal) + `merge.down` (vertical)

OOXML fixes a table's column count at `<a:tblGrid>` and row count at
`<a:tr>` — no "narrower row" or "shorter column" exists. Visual merging
is done in-place on the full grid via `gridSpan` (horizontal) or
`merge.down` (vertical, wraps `rowSpan` + `vMerge` continuation cells).

**Top table — `gridSpan=N` (full-width footnote):**

```bash
# Append a normal 4-cell row, then horizontally merge tc[1] across all 4 cols.
officecli add file.pptx /slide[4]/table[1] --type row \
  --prop c1="Footnote: figures in thousands USD, unaudited."
officecli set file.pptx /slide[4]/table[1]/tr[3]/tc[1] \
  --prop gridSpan=4 --prop fill=F2F2F2 --prop bold=true
```

`gridSpan=N` on `tc[1]` flags the next N-1 cells as `hMerge=true` — they
keep their `tc[N]` slots but render as part of the wide cell. Don't set
text on the continuation cells.

**Bottom table — `merge.down=N` (grouped row labels):**

```bash
# Merge "North" cell down 3 rows total (anchor + 2 continuations).
officecli set file.pptx /slide[4]/table[2]/tr[2]/tc[1] \
  --prop merge.down=2 --prop bold=true --prop fill=DEEAF6 --prop valign=middle
```

`merge.down=N` sets `rowSpan=N+1` on the anchor cell and `vMerge=true`
on the N continuation cells directly below. Useful for grouped row
labels (region/category bands).

**Rule of thumb:** for full-width headers / footnotes / totals, use
`gridSpan` on a normal row. For grouped row labels spanning vertically,
use `merge.down`.

**Features:** `add row`, `add column`, `set row.height`, `set col.width`,
table-level `rowHeight`, `c{N}=value` cell seeding, column `text=` seed,
`gridSpan` (horizontal merge), `merge.down` (vertical merge).
