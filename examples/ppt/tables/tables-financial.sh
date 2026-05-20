#!/bin/bash
# Real-world PowerPoint table example — quarterly financial report deck.
# Combines: built-in style, header banding, per-cell fills for traffic-light
# status, gridSpan section headers, right-aligned numbers, totals row.

set -e

DIR="$(dirname "$0")"
PPTX="$DIR/tables-financial.pptx"

rm -f "$PPTX"
officecli create "$PPTX"
officecli open "$PPTX"

# Theme colors
NAVY=1F3864; STEEL=2E75B6; PALE=DEEAF6
GREEN=00B050; AMBER=FFC000; RED=C00000

# --- Slide 1: Title ---
officecli add "$PPTX" /presentation/slides --type slide
officecli add "$PPTX" '/slide[1]' --type shape \
    --prop text="Q4 2025 Financial Review" --prop size=44 --prop bold=true --prop color="$NAVY" \
    --prop x=1in --prop y=2.5in --prop width=11in --prop height=1.2in --prop align=center
officecli add "$PPTX" '/slide[1]' --type shape \
    --prop text="Revenue · Expenses · Margin · Forecast" --prop size=22 --prop color=595959 \
    --prop x=1in --prop y=4in --prop width=11in --prop height=0.8in --prop align=center

# --- Slide 2: Quarterly P&L (sections via gridSpan) ---
officecli add "$PPTX" /presentation/slides --type slide
officecli add "$PPTX" '/slide[2]' --type shape \
    --prop text="Quarterly P&L (USD, thousands)" --prop size=28 --prop bold=true --prop color="$NAVY" \
    --prop x=0.5in --prop y=0.3in --prop width=12in --prop height=0.6in

officecli add "$PPTX" '/slide[2]' --type table \
    --prop x=0.5in --prop y=1.2in --prop width=12in --prop height=5.5in \
    --prop rows=11 --prop cols=6

# Header
for entry in "1:Line Item" "2:Q1" "3:Q2" "4:Q3" "5:Q4" "6:FY Total"; do
    c="${entry%%:*}"; t="${entry#*:}"
    officecli set "$PPTX" "/slide[2]/table[1]/tr[1]/tc[$c]" \
        --prop text="$t" --prop bold=true --prop color=FFFFFF --prop fill="$NAVY" \
        --prop align=center
done

# Section: Revenue
officecli set "$PPTX" '/slide[2]/table[1]/tr[2]/tc[1]' \
    --prop text="REVENUE" --prop bold=true --prop fill="$STEEL" --prop color=FFFFFF \
    --prop gridSpan=6

set_row () {
    local r="$1" label="$2" q1="$3" q2="$4" q3="$5" q4="$6" tot="$7" emphasize="$8"
    local fill=""
    [ "$emphasize" = "1" ] && fill="--prop fill=$PALE"
    officecli set "$PPTX" "/slide[2]/table[1]/tr[$r]/tc[1]" --prop text="$label" $fill
    officecli set "$PPTX" "/slide[2]/table[1]/tr[$r]/tc[2]" --prop text="$q1" --prop align=right $fill
    officecli set "$PPTX" "/slide[2]/table[1]/tr[$r]/tc[3]" --prop text="$q2" --prop align=right $fill
    officecli set "$PPTX" "/slide[2]/table[1]/tr[$r]/tc[4]" --prop text="$q3" --prop align=right $fill
    officecli set "$PPTX" "/slide[2]/table[1]/tr[$r]/tc[5]" --prop text="$q4" --prop align=right $fill
    officecli set "$PPTX" "/slide[2]/table[1]/tr[$r]/tc[6]" --prop text="$tot" --prop align=right --prop bold=true $fill
}

set_row 3 "  Product Sales"      "1,200" "1,350" "1,480" "1,720" "5,750" 0
set_row 4 "  Services"             "480"   "520"   "590"   "640" "2,230" 0
set_row 5 "  Licensing"            "120"   "140"   "165"   "195"   "620" 0
set_row 6 "  Subtotal"           "1,800" "2,010" "2,235" "2,555" "8,600" 1

# Section: Expenses
officecli set "$PPTX" '/slide[2]/table[1]/tr[7]/tc[1]' \
    --prop text="EXPENSES" --prop bold=true --prop fill="$STEEL" --prop color=FFFFFF \
    --prop gridSpan=6

set_row 8  "  COGS"              "720" "810" "895" "1,025" "3,450" 0
set_row 9  "  Operating"         "380" "410" "445"   "490" "1,725" 0
set_row 10 "  Subtotal"        "1,100" "1,220" "1,340" "1,515" "5,175" 1

# Net row
officecli set "$PPTX" '/slide[2]/table[1]/tr[11]/tc[1]' \
    --prop text="NET INCOME" --prop bold=true --prop fill="$GREEN" --prop color=FFFFFF
for entry in "2:700" "3:790" "4:895" "5:1,040" "6:3,425"; do
    c="${entry%%:*}"; v="${entry#*:}"
    officecli set "$PPTX" "/slide[2]/table[1]/tr[11]/tc[$c]" \
        --prop text="$v" --prop align=right --prop bold=true \
        --prop fill="$GREEN" --prop color=FFFFFF
done

# --- Slide 3: Risk register (traffic-light fills) ---
officecli add "$PPTX" /presentation/slides --type slide
officecli add "$PPTX" '/slide[3]' --type shape \
    --prop text="Risk Register" --prop size=28 --prop bold=true --prop color="$NAVY" \
    --prop x=0.5in --prop y=0.3in --prop width=12in --prop height=0.6in

officecli add "$PPTX" '/slide[3]' --type table \
    --prop x=0.5in --prop y=1.2in --prop width=12in --prop height=4in \
    --prop style=medium2 --prop firstRow=true --prop bandedRows=true \
    --prop data="Risk,Impact,Likelihood,Owner,Status;FX volatility,High,Medium,CFO,At risk;Supply chain,Medium,Low,COO,On track;Talent attrition,High,High,CPO,Critical;Reg compliance,Medium,Medium,GC,On track;Cybersecurity,High,Low,CTO,On track"

# Color the Status column (col 5, rows 2..6)
officecli set "$PPTX" '/slide[3]/table[1]/tr[2]/tc[5]' \
    --prop text="At risk" --prop fill="$AMBER" --prop bold=true --prop align=center
officecli set "$PPTX" '/slide[3]/table[1]/tr[3]/tc[5]' \
    --prop text="On track" --prop fill="$GREEN" --prop color=FFFFFF --prop bold=true --prop align=center
officecli set "$PPTX" '/slide[3]/table[1]/tr[4]/tc[5]' \
    --prop text="Critical" --prop fill="$RED" --prop color=FFFFFF --prop bold=true --prop align=center
officecli set "$PPTX" '/slide[3]/table[1]/tr[5]/tc[5]' \
    --prop text="On track" --prop fill="$GREEN" --prop color=FFFFFF --prop bold=true --prop align=center
officecli set "$PPTX" '/slide[3]/table[1]/tr[6]/tc[5]' \
    --prop text="On track" --prop fill="$GREEN" --prop color=FFFFFF --prop bold=true --prop align=center

# --- Slide 4: KPI summary (small table) ---
officecli add "$PPTX" /presentation/slides --type slide
officecli add "$PPTX" '/slide[4]' --type shape \
    --prop text="KPI Summary" --prop size=28 --prop bold=true --prop color="$NAVY" \
    --prop x=0.5in --prop y=0.3in --prop width=12in --prop height=0.6in

officecli add "$PPTX" '/slide[4]' --type table \
    --prop x=2in --prop y=1.5in --prop width=9in --prop height=3.5in \
    --prop style=medium4 --prop firstRow=true --prop firstCol=true --prop lastRow=true \
    --prop data="Metric,Target,Actual,Variance;Revenue (\$M),8.0,8.6,+7.5%;Gross Margin,38%,40.1%,+2.1pp;Op Margin,18%,19.8%,+1.8pp;CAC Payback,14 mo,12 mo,-2 mo;Total,—,—,Beat"

officecli close "$PPTX"
officecli validate "$PPTX"
echo "Created: $PPTX"
