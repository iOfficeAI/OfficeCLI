# Real-World PPT Tables — Financial Review Deck

Three files work together:

- **tables-financial.sh** — Build script.
- **tables-financial.pptx** — 4-slide deck (title + 3 table slides).
- **tables-financial.md** — This file.

A realistic finance deck combining the techniques from the other
`tables-*` examples: a quarterly P&L with section headers, a risk register
with traffic-light status fills, and a KPI summary using a built-in theme
style.

## Regenerate

```bash
cd examples/ppt
bash tables-financial.sh
# → tables-financial.pptx
```

## Slides

### Slide 1 — Title

Plain text shapes, no table — establishes a navy/grey theme used throughout.

### Slide 2 — Quarterly P&L

11×6 table. Demonstrates:

- **Header row** — solid navy fill, white bold centered text.
- **Section bands** — `gridSpan=6` cells used as REVENUE / EXPENSES dividers.
- **Subtotal emphasis** — pale-blue row fill via per-cell `fill=$PALE`.
- **Net Income highlight** — green-fill row across all columns.
- **Right-aligned numbers** — `align=right` on numeric columns; bold on totals.

### Slide 3 — Risk Register (traffic-light fills)

Built with `style=medium2` + `firstRow + bandedRows`, then the Status
column is overridden per-cell with green / amber / red fills:

```bash
officecli set file.pptx /slide[3]/table[1]/tr[4]/tc[5] \
  --prop text="Critical" --prop fill=C00000 \
  --prop color=FFFFFF --prop bold=true --prop align=center
```

### Slide 4 — KPI Summary

Built-in `style=medium4` with `firstRow + firstCol + lastRow` — a compact
table where every visual element (header band, first-column emphasis,
totals row) comes from the theme; no per-cell styling needed beyond the
inline `data=` seed.

```bash
officecli add file.pptx /slide[4] --type table \
  --prop style=medium4 \
  --prop firstRow=true --prop firstCol=true --prop lastRow=true \
  --prop data="Metric,Target,Actual,Variance;Revenue (\$M),8.0,8.6,+7.5%;..."
```

**Features used:** all of `tables-basic`, `tables-styled`, `tables-merged`,
plus per-cell traffic-light fills.
