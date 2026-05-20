#!/bin/bash
# Basic PowerPoint table — header row, body rows, fills, font sizing.
# Demonstrates: add table with inline `data=` CSV, headerFill/bodyFill,
# per-cell text override via set, table dimensions (x/y/width/height).

set -e

DIR="$(dirname "$0")"
PPTX="$DIR/tables-basic.pptx"

rm -f "$PPTX"
officecli create "$PPTX"
officecli open "$PPTX"

# --- Slide 1: minimal 3x3 table seeded inline ---
officecli add "$PPTX" /presentation/slides --type slide
officecli add "$PPTX" '/slide[1]' --type shape \
    --prop text="Basic Table — Inline Data" --prop size=28 --prop bold=true \
    --prop x=0.5in --prop y=0.3in --prop width=12in --prop height=0.6in

# 'data=' uses CSV (comma = cell sep, semicolon = row sep).
officecli add "$PPTX" '/slide[1]' --type table \
    --prop x=0.5in --prop y=1.2in --prop width=12in --prop height=2in \
    --prop headerFill=4472C4 --prop bodyFill=DEEAF6 \
    --prop data="Region,Q1,Q2,Q3,Q4;North,120,135,142,168;South,98,110,121,140;East,165,178,190,205"

# --- Slide 2: explicit rows/cols then per-cell text ---
officecli add "$PPTX" /presentation/slides --type slide
officecli add "$PPTX" '/slide[2]' --type shape \
    --prop text="Basic Table — Per-Cell Set" --prop size=28 --prop bold=true \
    --prop x=0.5in --prop y=0.3in --prop width=12in --prop height=0.6in

officecli add "$PPTX" '/slide[2]' --type table \
    --prop x=0.5in --prop y=1.2in --prop width=10in --prop height=2.5in \
    --prop rows=4 --prop cols=3 --prop headerFill=2E75B6

# Header
for entry in "1:Product" "2:Units" "3:Revenue"; do
    col="${entry%%:*}"; txt="${entry#*:}"
    officecli set "$PPTX" "/slide[2]/table[1]/tr[1]/tc[$col]" \
        --prop text="$txt" --prop bold=true --prop color=FFFFFF
done

# Body
officecli set "$PPTX" '/slide[2]/table[1]/tr[2]/tc[1]' --prop text="Widget"
officecli set "$PPTX" '/slide[2]/table[1]/tr[2]/tc[2]' --prop text="1,200"
officecli set "$PPTX" '/slide[2]/table[1]/tr[2]/tc[3]' --prop text="\$48,000"
officecli set "$PPTX" '/slide[2]/table[1]/tr[3]/tc[1]' --prop text="Gizmo"
officecli set "$PPTX" '/slide[2]/table[1]/tr[3]/tc[2]' --prop text="850"
officecli set "$PPTX" '/slide[2]/table[1]/tr[3]/tc[3]' --prop text="\$72,250"
officecli set "$PPTX" '/slide[2]/table[1]/tr[4]/tc[1]' --prop text="Sprocket"
officecli set "$PPTX" '/slide[2]/table[1]/tr[4]/tc[2]' --prop text="430"
officecli set "$PPTX" '/slide[2]/table[1]/tr[4]/tc[3]' --prop text="\$25,800"

# --- Slide 3: Cell fill variations (solid hex, theme color, gradient, none) ---
officecli add "$PPTX" /presentation/slides --type slide
officecli add "$PPTX" '/slide[3]' --type shape \
    --prop text="Cell Fill Variations" --prop size=28 --prop bold=true \
    --prop x=0.5in --prop y=0.3in --prop width=12in --prop height=0.6in

officecli add "$PPTX" '/slide[3]' --type table \
    --prop x=0.5in --prop y=1.2in --prop width=12in --prop height=4in \
    --prop rows=5 --prop cols=2 --prop style=none --prop border.all="1pt solid 808080"

officecli set "$PPTX" '/slide[3]/table[1]/tr[1]/tc[1]' --prop text="fill spec" --prop bold=true --prop fill=404040 --prop color=FFFFFF
officecli set "$PPTX" '/slide[3]/table[1]/tr[1]/tc[2]' --prop text="rendered" --prop bold=true --prop fill=404040 --prop color=FFFFFF

# Solid hex
officecli set "$PPTX" '/slide[3]/table[1]/tr[2]/tc[1]' --prop text='fill=FF0000  (solid hex)'
officecli set "$PPTX" '/slide[3]/table[1]/tr[2]/tc[2]' --prop fill=FF0000

# Named color
officecli set "$PPTX" '/slide[3]/table[1]/tr[3]/tc[1]' --prop text='fill=red  /  fill=rgb(255,0,0)  (named / rgb forms)'
officecli set "$PPTX" '/slide[3]/table[1]/tr[3]/tc[2]' --prop fill=red

# Theme color — accent1 follows the deck theme
officecli set "$PPTX" '/slide[3]/table[1]/tr[4]/tc[1]' --prop text='fill=accent1  (theme color, follows deck theme)'
officecli set "$PPTX" '/slide[3]/table[1]/tr[4]/tc[2]' --prop fill=accent1

# Gradient — "COLOR1-COLOR2[-ANGLE]"
officecli set "$PPTX" '/slide[3]/table[1]/tr[5]/tc[1]' --prop text='fill="FF0000-0000FF-90"  (gradient, 90° angle)'
officecli set "$PPTX" '/slide[3]/table[1]/tr[5]/tc[2]' --prop fill="FF0000-0000FF-90"

# fill=none demo (separate small table so 'none' is visible against page bg)
officecli add "$PPTX" '/slide[3]' --type shape \
    --prop text='fill=none  (explicit no-fill; cell becomes transparent):' --prop size=14 \
    --prop x=0.5in --prop y=5.4in --prop width=12in --prop height=0.4in
officecli add "$PPTX" '/slide[3]' --type table \
    --prop x=0.5in --prop y=5.9in --prop width=4in --prop height=0.8in \
    --prop rows=1 --prop cols=2 --prop style=none --prop border.all="1pt solid 000000"
officecli set "$PPTX" '/slide[3]/table[2]/tr[1]/tc[1]' --prop text="solid" --prop fill=FFE699
officecli set "$PPTX" '/slide[3]/table[2]/tr[1]/tc[2]' --prop text="none" --prop fill=none

officecli close "$PPTX"
officecli validate "$PPTX"
echo "Created: $PPTX"
