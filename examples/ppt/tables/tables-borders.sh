#!/bin/bash
# PowerPoint table border styling.
# Demonstrates: border.all shorthand, per-edge borders (top/right/bottom/left),
# inside dividers (horizontal/vertical), diagonal borders (tl2br/tr2bl),
# dash patterns (solid/dot/dash/lgDash/dashDot/sysDot/sysDash).

set -e

DIR="$(dirname "$0")"
PPTX="$DIR/tables-borders.pptx"

rm -f "$PPTX"
officecli create "$PPTX"
officecli open "$PPTX"

DATA="A,B,C;1,2,3;4,5,6;7,8,9"

add_table () {
    local slide="$1" x="$2" y="$3" label="$4"; shift 4
    local ty
    ty=$(awk -v a="$y" 'BEGIN{print a+0.35}')
    officecli add "$PPTX" "/slide[$slide]" --type shape \
        --prop text="$label" --prop size=12 --prop bold=true \
        --prop x="${x}in" --prop y="${y}in" --prop width=4in --prop height=0.3in
    officecli add "$PPTX" "/slide[$slide]" --type table \
        --prop x="${x}in" --prop y="${ty}in" \
        --prop width=3.5in --prop height=1.8in --prop style=none \
        --prop data="$DATA" "$@"
}

# --- Slide 1: border shorthand & per-edge ---
officecli add "$PPTX" /presentation/slides --type slide
officecli add "$PPTX" '/slide[1]' --type shape \
    --prop text="Borders: Shorthand & Per-Edge" --prop size=28 --prop bold=true \
    --prop x=0.5in --prop y=0.3in --prop width=12in --prop height=0.6in

add_table 1 0.5  1.0 "border.all=1pt solid 808080"     --prop border.all="1pt solid 808080"
add_table 1 5.0  1.0 "border.all=2pt solid FF0000"     --prop border.all="2pt solid FF0000"
add_table 1 9.5  1.0 "border.all=none"                 --prop border.all=none

add_table 1 0.5  3.5 "border.top=3pt solid 000000"     --prop border.top="3pt solid 000000"
add_table 1 5.0  3.5 "border.bottom=3pt solid 0070C0"  --prop border.bottom="3pt solid 0070C0"
add_table 1 9.5  3.5 "border.left=3pt solid 00B050"    --prop border.left="3pt solid 00B050"

# --- Slide 2: inside dividers & dash patterns ---
officecli add "$PPTX" /presentation/slides --type slide
officecli add "$PPTX" '/slide[2]' --type shape \
    --prop text="Borders: Inside Dividers & Dashes" --prop size=28 --prop bold=true \
    --prop x=0.5in --prop y=0.3in --prop width=12in --prop height=0.6in

add_table 2 0.5  1.0 "border.horizontal=1pt solid CCC" \
    --prop border.horizontal="1pt solid CCCCCC" --prop border.all="1pt solid 404040"
add_table 2 5.0  1.0 "border.vertical=1pt dash 0070C0" \
    --prop border.vertical="1pt dash 0070C0" --prop border.all="1pt solid 404040"
add_table 2 9.5  1.0 "horizontal+vertical=dot" \
    --prop border.horizontal="1pt dot 808080" --prop border.vertical="1pt dot 808080" \
    --prop border.all="2pt solid 000000"

add_table 2 0.5  3.5 "dash=lgDash"  --prop border.all="1.5pt lgDash FF0000"
add_table 2 5.0  3.5 "dash=dashDot" --prop border.all="1.5pt dashDot 0070C0"
add_table 2 9.5  3.5 "dash=sysDash" --prop border.all="1.5pt sysDash 00B050"

# --- Slide 3: diagonal borders (per-cell) ---
officecli add "$PPTX" /presentation/slides --type slide
officecli add "$PPTX" '/slide[3]' --type shape \
    --prop text="Diagonal Borders (per-cell, tl2br / tr2bl)" --prop size=28 --prop bold=true \
    --prop x=0.5in --prop y=0.3in --prop width=12in --prop height=0.6in
officecli add "$PPTX" '/slide[3]' --type shape \
    --prop text="Typical use: 'crossed out' header corner cell." --prop size=14 \
    --prop x=0.5in --prop y=0.95in --prop width=12in --prop height=0.4in

officecli add "$PPTX" '/slide[3]' --type table \
    --prop x=2in --prop y=1.6in --prop width=9in --prop height=3in \
    --prop rows=4 --prop cols=4 --prop border.all="1pt solid 808080"

# Top-left corner: diagonal split with 'Month' / 'Region' labels
officecli set "$PPTX" '/slide[3]/table[1]/tr[1]/tc[1]' \
    --prop text="" --prop fill=F2F2F2 \
    --prop border.tl2br="1pt solid 808080"

# Column headers
officecli set "$PPTX" '/slide[3]/table[1]/tr[1]/tc[2]' --prop text="Jan" --prop bold=true --prop align=center --prop fill=DEEAF6
officecli set "$PPTX" '/slide[3]/table[1]/tr[1]/tc[3]' --prop text="Feb" --prop bold=true --prop align=center --prop fill=DEEAF6
officecli set "$PPTX" '/slide[3]/table[1]/tr[1]/tc[4]' --prop text="Mar" --prop bold=true --prop align=center --prop fill=DEEAF6

# Row headers + data
officecli set "$PPTX" '/slide[3]/table[1]/tr[2]/tc[1]' --prop text="North" --prop bold=true --prop fill=DEEAF6
officecli set "$PPTX" '/slide[3]/table[1]/tr[2]/tc[2]' --prop text="120"
officecli set "$PPTX" '/slide[3]/table[1]/tr[2]/tc[3]' --prop text="135"
officecli set "$PPTX" '/slide[3]/table[1]/tr[2]/tc[4]' --prop text="142"
officecli set "$PPTX" '/slide[3]/table[1]/tr[3]/tc[1]' --prop text="South" --prop bold=true --prop fill=DEEAF6
officecli set "$PPTX" '/slide[3]/table[1]/tr[3]/tc[2]' --prop text="98"
officecli set "$PPTX" '/slide[3]/table[1]/tr[3]/tc[3]' --prop text="110"
officecli set "$PPTX" '/slide[3]/table[1]/tr[3]/tc[4]' --prop text="121"
officecli set "$PPTX" '/slide[3]/table[1]/tr[4]/tc[1]' --prop text="East" --prop bold=true --prop fill=DEEAF6
officecli set "$PPTX" '/slide[3]/table[1]/tr[4]/tc[2]' --prop text="165"
officecli set "$PPTX" '/slide[3]/table[1]/tr[4]/tc[3]' --prop text="178"
officecli set "$PPTX" '/slide[3]/table[1]/tr[4]/tc[4]' --prop text="190"

# A standalone cell with both diagonals (X pattern)
officecli add "$PPTX" '/slide[3]' --type shape \
    --prop text="Both diagonals on a single cell:" --prop size=14 \
    --prop x=0.5in --prop y=5.2in --prop width=12in --prop height=0.4in
officecli add "$PPTX" '/slide[3]' --type table \
    --prop x=5in --prop y=5.7in --prop width=3in --prop height=1.2in \
    --prop rows=1 --prop cols=1 --prop border.all="1pt solid 000000"
officecli set "$PPTX" '/slide[3]/table[2]/tr[1]/tc[1]' \
    --prop text="N/A" --prop align=center --prop fill=F2F2F2 \
    --prop border.tl2br="1pt solid C00000" \
    --prop border.tr2bl="1pt solid C00000"

officecli close "$PPTX"
officecli validate "$PPTX"
echo "Created: $PPTX"
