# Merged Cells in PPT Tables

Three files work together:

- **tables-merged.sh** — Build script.
- **tables-merged.pptx** — 2-slide deck.
- **tables-merged.md** — This file.

## Regenerate

```bash
cd examples/ppt
bash tables-merged.sh
# → tables-merged.pptx
```

## Merge axes — horizontal + vertical

PPT OOXML supports horizontal (`gridSpan` + `hMerge`) and vertical
(`rowSpan` + `vMerge`) merging. officecli exposes both on the write side:

- `gridSpan=N` on a cell — horizontal merge across N columns
- `merge.down=N` on a cell — vertical merge spanning N+1 rows total
  (anchor + N continuation rows below)

This file walks through `gridSpan` (the more common case). See
`tables-rows-cols.{md,sh}` slide 4 for a `merge.down` example.

## Slides

### Slide 1 — Two-level header (`gridSpan` on the super-header row)

A super-header row where two cells each span two columns:

```
| Department | 2024 Performance | 2025 Forecast |
|            | Revenue | Margin | Revenue | Margin |
| Eng        |  1.20M  |  18%   |  1.45M  |  22%   |
```

```bash
# Row 1: super-headers. gridSpan=2 stamps hMerge on the next cell.
officecli set file.pptx /slide[1]/table[1]/tr[1]/tc[2] \
  --prop text="2024 Performance" --prop gridSpan=2

# tc[3] is now a continuation cell — skip to tc[4]:
officecli set file.pptx /slide[1]/table[1]/tr[1]/tc[4] \
  --prop text="2025 Forecast" --prop gridSpan=2
```

**Important:** cell indices do **not** renumber after `gridSpan`. The merged
cells still occupy their original `tc[N]` slots; you just shouldn't set text
on them. Setting `gridSpan=2` on `tc[2]` doesn't make `tc[3]` go away — it
flags `tc[3]` as `hMerge=true`.

### Slide 2 — Full-width section headers

Span the entire table to create section dividers, then list items below:

```bash
officecli set file.pptx /slide[2]/table[1]/tr[2]/tc[1] \
  --prop text="◆ Phase 1 — Discovery" --prop bold=true \
  --prop fill=FFE699 --prop gridSpan=4
```

**Features:** `gridSpan`, per-cell `fill` for color-coding sections.
