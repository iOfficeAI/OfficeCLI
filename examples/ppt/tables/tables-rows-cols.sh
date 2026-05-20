#!/bin/bash
# PowerPoint table row & column operations.
# Demonstrates: add row / add column (grow an existing table),
# per-row height (set row.height), per-column width (set col.width),
# column seed text, gridSpan (horizontal merge), merge.down (vertical merge).

set -e

DIR="$(dirname "$0")"
PPTX="$DIR/tables-rows-cols.pptx"

rm -f "$PPTX"
officecli create "$PPTX"
officecli open "$PPTX"

# --- Slide 1: Grow a table by add row / add column ---
# Two side-by-side tables compare the two coloring models:
#   LEFT  A. style=medium2   → table-level theme, auto-follows new rows/columns.
#   RIGHT B. headerFill/bodyFill → per-cell stamp, does NOT follow; manual top-up needed.
# Placed side-by-side (each 6in wide, half the 13.33in widescreen slide) so the
# visual contrast is immediate.
officecli add "$PPTX" /presentation/slides --type slide
officecli add "$PPTX" '/slide[1]' --type shape \
    --prop text="Grow a Table — Theme vs Per-Cell Stamp" --prop size=28 --prop bold=true \
    --prop x=0.5in --prop y=0.3in --prop width=12in --prop height=0.6in

# === LEFT: Table A — style=medium2 (theme, auto-inherits) ===
officecli add "$PPTX" '/slide[1]' --type shape \
    --prop text="A) style=medium2  (auto-follows)" --prop size=14 --prop bold=true \
    --prop x=0.5in --prop y=1in --prop width=6in --prop height=0.4in

officecli add "$PPTX" '/slide[1]' --type table \
    --prop x=0.5in --prop y=1.5in --prop width=2in --prop height=1.5in \
    --prop style=medium2 --prop firstRow=true --prop bandedRows=true --prop lastCol=true \
    --prop data="Name,H1;Alice,220"

# Append 2 rows + 1 column — set NOTHING about fill. PowerPoint paints
# every new cell via the medium2 theme. H1 = first-half total (Q1+Q2),
# H2 = second-half total (Q3+Q4); appended as a derived summary column.
officecli add "$PPTX" '/slide[1]/table[1]' --type row --prop c1=Bob --prop c2=205
officecli add "$PPTX" '/slide[1]/table[1]' --type row --prop c1=Carol --prop c2=275
officecli add "$PPTX" '/slide[1]/table[1]' --type column \
    --prop width=1in --prop text="H2"
officecli set "$PPTX" '/slide[1]/table[1]/tr[2]/tc[3]' --prop text="245"
officecli set "$PPTX" '/slide[1]/table[1]/tr[3]/tc[3]' --prop text="225"
officecli set "$PPTX" '/slide[1]/table[1]/tr[4]/tc[3]' --prop text="335"
officecli add "$PPTX" '/slide[1]/table[1]' --type column \
    --prop width=1in --prop text="Total"
officecli set "$PPTX" '/slide[1]/table[1]/tr[2]/tc[4]' --prop text="465" --prop bold=true
officecli set "$PPTX" '/slide[1]/table[1]/tr[3]/tc[4]' --prop text="430" --prop bold=true
officecli set "$PPTX" '/slide[1]/table[1]/tr[4]/tc[4]' --prop text="610" --prop bold=true

# === RIGHT: Table B — headerFill/bodyFill (per-cell stamp, does NOT inherit) ===
officecli add "$PPTX" '/slide[1]' --type shape \
    --prop text="B) headerFill/bodyFill  (manual top-up)" --prop size=14 --prop bold=true \
    --prop x=7in --prop y=1in --prop width=6in --prop height=0.4in

officecli add "$PPTX" '/slide[1]' --type table \
    --prop x=7in --prop y=1.5in --prop width=2in --prop height=1.5in \
    --prop headerFill=4472C4 --prop bodyFill=DEEAF6 \
    --prop data="Name,H1;Alice,220"
officecli add "$PPTX" '/slide[1]/table[2]' --type row --prop c1=Bob --prop c2=205
officecli add "$PPTX" '/slide[1]/table[2]' --type row --prop c1=Carol --prop c2=275
officecli add "$PPTX" '/slide[1]/table[2]' --type column \
    --prop width=1in --prop text="H2"
officecli set "$PPTX" '/slide[1]/table[2]/tr[2]/tc[3]' --prop text="245"
officecli set "$PPTX" '/slide[1]/table[2]/tr[3]/tc[3]' --prop text="225"
officecli set "$PPTX" '/slide[1]/table[2]/tr[4]/tc[3]' --prop text="335"
officecli add "$PPTX" '/slide[1]/table[2]' --type column \
    --prop width=1in --prop text="Total"
officecli set "$PPTX" '/slide[1]/table[2]/tr[2]/tc[4]' --prop text="465"
officecli set "$PPTX" '/slide[1]/table[2]/tr[3]/tc[4]' --prop text="430"
officecli set "$PPTX" '/slide[1]/table[2]/tr[4]/tc[4]' --prop text="610"

# Manual top-up — headerFill/bodyFill are a one-shot stamp at add-table time.
# Every cell created later by add row / add column has no fill and must be
# styled explicitly. The Total column gets a darker fill (SUM) + bold so it
# reads as a totals band; table A gets the equivalent emphasis from medium2's
# last-column theme styling for free.
HDR=4472C4; BODY=DEEAF6; SUM=B4C7E7
# Bob, Carol body fill across the original 2 columns.
for c in 1 2; do
    officecli set "$PPTX" "/slide[1]/table[2]/tr[3]/tc[$c]" --prop fill=$BODY
    officecli set "$PPTX" "/slide[1]/table[2]/tr[4]/tc[$c]" --prop fill=$BODY
done
# H2 column — header HDR, body BODY for all 3 data rows.
officecli set "$PPTX" '/slide[1]/table[2]/tr[1]/tc[3]' --prop fill=$HDR --prop color=FFFFFF --prop bold=true
for r in 2 3 4; do
    officecli set "$PPTX" "/slide[1]/table[2]/tr[$r]/tc[3]" --prop fill=$BODY
done
# Total column — header HDR, body SUM (bold) for all 3 data rows.
officecli set "$PPTX" '/slide[1]/table[2]/tr[1]/tc[4]' --prop fill=$HDR --prop color=FFFFFF --prop bold=true
for r in 2 3 4; do
    officecli set "$PPTX" "/slide[1]/table[2]/tr[$r]/tc[4]" --prop fill=$SUM --prop bold=true
done

# --- Slide 2: Per-row heights & per-column widths ---
officecli add "$PPTX" /presentation/slides --type slide
officecli add "$PPTX" '/slide[2]' --type shape \
    --prop text="Per-Row Height + Per-Column Width" --prop size=28 --prop bold=true \
    --prop x=0.5in --prop y=0.3in --prop width=12in --prop height=0.6in

officecli add "$PPTX" '/slide[2]' --type table \
    --prop x=0.5in --prop y=1.2in --prop width=12in --prop height=4in \
    --prop rows=4 --prop cols=4 --prop headerFill=2E75B6

# Header
for c in 1 2 3 4; do
    officecli set "$PPTX" "/slide[2]/table[1]/tr[1]/tc[$c]" \
        --prop bold=true --prop color=FFFFFF --prop align=center
done
officecli set "$PPTX" '/slide[2]/table[1]/tr[1]/tc[1]' --prop text="Field" --prop bold=true --prop color=FFFFFF
officecli set "$PPTX" '/slide[2]/table[1]/tr[1]/tc[2]' --prop text="Short" --prop bold=true --prop color=FFFFFF
officecli set "$PPTX" '/slide[2]/table[1]/tr[1]/tc[3]' --prop text="Wide" --prop bold=true --prop color=FFFFFF
officecli set "$PPTX" '/slide[2]/table[1]/tr[1]/tc[4]' --prop text="Narrow" --prop bold=true --prop color=FFFFFF

# Custom per-column widths (the four columns total ~12in).
officecli set "$PPTX" '/slide[2]/table[1]/col[1]' --prop width=2in
officecli set "$PPTX" '/slide[2]/table[1]/col[2]' --prop width=1.5in
officecli set "$PPTX" '/slide[2]/table[1]/col[3]' --prop width=7in
officecli set "$PPTX" '/slide[2]/table[1]/col[4]' --prop width=1.5in

# Custom per-row heights — header thin, body increasing.
officecli set "$PPTX" '/slide[2]/table[1]/tr[1]' --prop height=0.5in
officecli set "$PPTX" '/slide[2]/table[1]/tr[2]' --prop height=0.6in
officecli set "$PPTX" '/slide[2]/table[1]/tr[3]' --prop height=1in
officecli set "$PPTX" '/slide[2]/table[1]/tr[4]' --prop height=1.5in

officecli set "$PPTX" '/slide[2]/table[1]/tr[2]/tc[1]' --prop text="Title"
officecli set "$PPTX" '/slide[2]/table[1]/tr[2]/tc[2]' --prop text="A"
officecli set "$PPTX" '/slide[2]/table[1]/tr[2]/tc[3]' --prop text="Standard row height (0.6in)"
officecli set "$PPTX" '/slide[2]/table[1]/tr[2]/tc[4]' --prop text="x"

officecli set "$PPTX" '/slide[2]/table[1]/tr[3]/tc[1]' --prop text="Body"
officecli set "$PPTX" '/slide[2]/table[1]/tr[3]/tc[2]' --prop text="B"
officecli set "$PPTX" '/slide[2]/table[1]/tr[3]/tc[3]' --prop text="Taller row (1in) for emphasis"
officecli set "$PPTX" '/slide[2]/table[1]/tr[3]/tc[4]' --prop text="y"

officecli set "$PPTX" '/slide[2]/table[1]/tr[4]/tc[1]' --prop text="Notes"
officecli set "$PPTX" '/slide[2]/table[1]/tr[4]/tc[2]' --prop text="C"
officecli set "$PPTX" '/slide[2]/table[1]/tr[4]/tc[3]' --prop text="Tallest row (1.5in) — multi-line content"
officecli set "$PPTX" '/slide[2]/table[1]/tr[4]/tc[4]' --prop text="z"

# --- Slide 3: Uniform row height via table-level rowHeight ---
officecli add "$PPTX" /presentation/slides --type slide
officecli add "$PPTX" '/slide[3]' --type shape \
    --prop text="Uniform rowHeight (table-level)" --prop size=28 --prop bold=true \
    --prop x=0.5in --prop y=0.3in --prop width=12in --prop height=0.6in

# Setting rowHeight at add-time stamps every row with the same height
# (no need to set each row individually).
officecli add "$PPTX" '/slide[3]' --type table \
    --prop x=0.5in --prop y=1.2in --prop width=12in \
    --prop rows=5 --prop cols=3 --prop rowHeight=0.8in \
    --prop headerFill=1F4E79 --prop bodyFill=F2F2F2 \
    --prop data="Step,Action,Result;1,Init,OK;2,Process,OK;3,Verify,OK;4,Commit,OK"

# --- Slide 4: Cell merging — gridSpan (horizontal) + vMerge (vertical) ---
# OOXML's table model fixes row width at <a:tblGrid> column count and row
# count at <a:tr> count — no "narrower row" or "shorter column" exists.
# Visual merging is done in-place via gridSpan (horizontal) or merge.down
# (vertical, wraps rowSpan + vMerge). Two tables on this slide show the
# two axes of merging.
officecli add "$PPTX" /presentation/slides --type slide
officecli add "$PPTX" '/slide[4]' --type shape \
    --prop text="Cell Merging — gridSpan (horizontal) + merge.down (vertical)" \
    --prop size=28 --prop bold=true \
    --prop x=0.5in --prop y=0.3in --prop width=12in --prop height=1in

# === Top table: gridSpan=N — full-width footnote ===
officecli add "$PPTX" '/slide[4]' --type shape \
    --prop text="1) gridSpan=N on first cell of a row — one wide cell across all N columns" \
    --prop size=14 --prop bold=true \
    --prop x=0.5in --prop y=1in --prop width=12in --prop height=0.4in
officecli add "$PPTX" '/slide[4]' --type table \
    --prop x=0.5in --prop y=1.5in --prop width=12in --prop height=1.5in \
    --prop headerFill=2E75B6 \
    --prop data="Q1,Q2,Q3,Q4;100,120,135,150"
# Append a normal 4-cell row, then horizontally merge via gridSpan on tc[1].
officecli add "$PPTX" '/slide[4]/table[1]' --type row \
    --prop c1="Footnote: figures in thousands USD, unaudited."
officecli set "$PPTX" '/slide[4]/table[1]/tr[3]/tc[1]' \
    --prop gridSpan=4 --prop fill=F2F2F2 --prop bold=true

# === Bottom table: merge.down=N — grouped row labels ===
officecli add "$PPTX" '/slide[4]' --type shape \
    --prop text="2) merge.down=N on a cell — one tall cell spanning N rows (vMerge + rowSpan)" \
    --prop size=14 --prop bold=true \
    --prop x=0.5in --prop y=3.3in --prop width=12in --prop height=0.4in
officecli add "$PPTX" '/slide[4]' --type table \
    --prop x=0.5in --prop y=3.8in --prop width=12in --prop height=3in \
    --prop headerFill=2E75B6 --prop rowHeight=0.5in \
    --prop data="Region,Month,Sales,Notes;North,Jan,120,;North,Feb,135,;North,Mar,142,;South,Jan,98,;South,Feb,110,"
# Merge "North" cell down 3 rows (rows 2..4); merge "South" cell down 2 rows
# (rows 5..6). merge.down=N spans the cell over N+1 rows total — the anchor
# row plus N continuation rows.
officecli set "$PPTX" '/slide[4]/table[2]/tr[2]/tc[1]' \
    --prop merge.down=2 --prop bold=true --prop fill=DEEAF6 --prop valign=middle
officecli set "$PPTX" '/slide[4]/table[2]/tr[5]/tc[1]' \
    --prop merge.down=1 --prop bold=true --prop fill=DEEAF6 --prop valign=middle

officecli close "$PPTX"
officecli validate "$PPTX"
echo "Created: $PPTX"
