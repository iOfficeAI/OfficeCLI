#!/bin/bash
# PowerPoint built-in table styles showcase.
# Demonstrates: style= (medium1..4, light1..3, dark1..2, none),
# firstRow/lastRow/firstCol/lastCol/bandedRows/bandedCols banding flags.

set -e

DIR="$(dirname "$0")"
PPTX="$DIR/tables-styled.pptx"

rm -f "$PPTX"
officecli create "$PPTX"
officecli open "$PPTX"

DATA="Region,Q1,Q2,Q3,Q4;North,120,135,142,168;South,98,110,121,140;East,165,178,190,205;West,140,155,168,182"

add_slide () {
    local idx="$1" style="$2" title="$3"
    officecli add "$PPTX" /presentation/slides --type slide
    officecli add "$PPTX" "/slide[$idx]" --type shape \
        --prop text="$title" --prop size=28 --prop bold=true \
        --prop x=0.5in --prop y=0.3in --prop width=12in --prop height=0.6in
    officecli add "$PPTX" "/slide[$idx]" --type table \
        --prop x=0.5in --prop y=1.2in --prop width=12in --prop height=3in \
        --prop style="$style" \
        --prop firstRow=true --prop bandedRows=true \
        --prop data="$DATA"
}

# 9 built-in styles, one per slide.
add_slide 1 medium1 "style=medium1"
add_slide 2 medium2 "style=medium2"
add_slide 3 medium3 "style=medium3"
add_slide 4 medium4 "style=medium4"
add_slide 5 light1  "style=light1"
add_slide 6 light2  "style=light2"
add_slide 7 light3  "style=light3"
add_slide 8 dark1   "style=dark1"
add_slide 9 dark2   "style=dark2"

# Slide 10: banding flag combinations on a single style.
officecli add "$PPTX" /presentation/slides --type slide
officecli add "$PPTX" '/slide[10]' --type shape \
    --prop text="Banding Flags (style=medium2)" --prop size=28 --prop bold=true \
    --prop x=0.5in --prop y=0.3in --prop width=12in --prop height=0.6in

officecli add "$PPTX" '/slide[10]' --type shape \
    --prop text="firstRow + bandedRows" --prop size=14 \
    --prop x=0.5in --prop y=1in --prop width=6in --prop height=0.4in
officecli add "$PPTX" '/slide[10]' --type table \
    --prop x=0.5in --prop y=1.4in --prop width=6in --prop height=2.5in \
    --prop style=medium2 --prop firstRow=true --prop bandedRows=true \
    --prop data="$DATA"

officecli add "$PPTX" '/slide[10]' --type shape \
    --prop text="firstCol + bandedCols" --prop size=14 \
    --prop x=7in --prop y=1in --prop width=6in --prop height=0.4in
officecli add "$PPTX" '/slide[10]' --type table \
    --prop x=7in --prop y=1.4in --prop width=6in --prop height=2.5in \
    --prop style=medium2 --prop firstCol=true --prop bandedCols=true \
    --prop data="$DATA"

officecli add "$PPTX" '/slide[10]' --type shape \
    --prop text="firstRow + lastRow (total row)" --prop size=14 \
    --prop x=0.5in --prop y=4.3in --prop width=6in --prop height=0.4in
officecli add "$PPTX" '/slide[10]' --type table \
    --prop x=0.5in --prop y=4.7in --prop width=6in --prop height=2.5in \
    --prop style=medium2 --prop firstRow=true --prop lastRow=true \
    --prop data="$DATA;Total,523,578,621,695"

officecli add "$PPTX" '/slide[10]' --type shape \
    --prop text="style=none (no theme)" --prop size=14 \
    --prop x=7in --prop y=4.3in --prop width=6in --prop height=0.4in
officecli add "$PPTX" '/slide[10]' --type table \
    --prop x=7in --prop y=4.7in --prop width=6in --prop height=2.5in \
    --prop style=none --prop border.all="1pt solid 808080" \
    --prop data="$DATA"

# --- Slide 11: rowHeight (uniform) + name (stable @name addressing) ---
officecli add "$PPTX" /presentation/slides --type slide
officecli add "$PPTX" '/slide[11]' --type shape \
    --prop text="rowHeight + name= addressing" --prop size=28 --prop bold=true \
    --prop x=0.5in --prop y=0.3in --prop width=12in --prop height=0.6in
officecli add "$PPTX" '/slide[11]' --type shape \
    --prop text="The table below was created with name=SalesData and rowHeight=1cm. After creation, it can be addressed as /slide[11]/table[@name=SalesData] instead of by positional index — handy when slides are reordered or tables added/removed." \
    --prop size=12 \
    --prop x=0.5in --prop y=0.95in --prop width=12in --prop height=0.8in

officecli add "$PPTX" '/slide[11]' --type table \
    --prop x=0.5in --prop y=2in --prop width=12in \
    --prop rows=5 --prop cols=4 --prop rowHeight=1cm \
    --prop name=SalesData --prop style=medium2 --prop firstRow=true \
    --prop data="Region,Q1,Q2,Q3;North,120,135,142;South,98,110,121;East,165,178,190;West,140,155,168"

# Demonstrate @name addressing — set a cell via the stable name path.
officecli set "$PPTX" '/slide[11]/table[@name=SalesData]/tr[2]/tc[2]' \
    --prop text="120 ▲" --prop bold=true --prop fill=C6E0B4

officecli close "$PPTX"
officecli validate "$PPTX"
echo "Created: $PPTX"
