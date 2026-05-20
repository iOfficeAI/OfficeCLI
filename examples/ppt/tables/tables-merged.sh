#!/bin/bash
# PowerPoint table cell merging — horizontal merge via gridSpan.
# Demonstrates: multi-column header spans, section headers spanning the table,
# nested header hierarchy.
#
# Note: officecli supports both horizontal (gridSpan) and vertical (merge.down)
# write-side merging. This file walks through gridSpan; see tables-rows-cols.sh
# slide 4 for a merge.down example.

set -e

DIR="$(dirname "$0")"
PPTX="$DIR/tables-merged.pptx"

rm -f "$PPTX"
officecli create "$PPTX"
officecli open "$PPTX"

# --- Slide 1: 2-level header (gridSpan on row 1) ---
officecli add "$PPTX" /presentation/slides --type slide
officecli add "$PPTX" '/slide[1]' --type shape \
    --prop text="Two-Level Header (gridSpan)" --prop size=28 --prop bold=true \
    --prop x=0.5in --prop y=0.3in --prop width=12in --prop height=0.6in

officecli add "$PPTX" '/slide[1]' --type table \
    --prop x=0.5in --prop y=1.2in --prop width=12in --prop height=3.5in \
    --prop rows=6 --prop cols=5 --prop headerFill=2E75B6 --prop bodyFill=DEEAF6

# Row 1: super-headers
officecli set "$PPTX" '/slide[1]/table[1]/tr[1]/tc[1]' \
    --prop text="Department" --prop bold=true --prop color=FFFFFF --prop align=center
officecli set "$PPTX" '/slide[1]/table[1]/tr[1]/tc[2]' \
    --prop text="2024 Performance" --prop bold=true --prop color=FFFFFF --prop align=center \
    --prop gridSpan=2
# tc[3] is now a continuation cell from gridSpan=2 — skip directly to tc[4].
officecli set "$PPTX" '/slide[1]/table[1]/tr[1]/tc[4]' \
    --prop text="2025 Forecast" --prop bold=true --prop color=FFFFFF --prop align=center \
    --prop gridSpan=2

# Row 2: sub-headers (lighter shade)
officecli set "$PPTX" '/slide[1]/table[1]/tr[2]/tc[1]' \
    --prop text="" --prop fill=5B9BD5
officecli set "$PPTX" '/slide[1]/table[1]/tr[2]/tc[2]' \
    --prop text="Revenue" --prop bold=true --prop color=FFFFFF --prop align=center --prop fill=5B9BD5
officecli set "$PPTX" '/slide[1]/table[1]/tr[2]/tc[3]' \
    --prop text="Margin" --prop bold=true --prop color=FFFFFF --prop align=center --prop fill=5B9BD5
officecli set "$PPTX" '/slide[1]/table[1]/tr[2]/tc[4]' \
    --prop text="Revenue" --prop bold=true --prop color=FFFFFF --prop align=center --prop fill=5B9BD5
officecli set "$PPTX" '/slide[1]/table[1]/tr[2]/tc[5]' \
    --prop text="Margin" --prop bold=true --prop color=FFFFFF --prop align=center --prop fill=5B9BD5

# Body rows
for row in "3:Engineering:1.20M:18%:1.45M:22%" \
           "4:Sales:2.30M:12%:2.80M:15%" \
           "5:Marketing:0.95M:25%:1.10M:28%" \
           "6:Operations:0.78M:30%:0.85M:32%"; do
    IFS=':' read -r r d a b c e <<< "$row"
    officecli set "$PPTX" "/slide[1]/table[1]/tr[$r]/tc[1]" --prop text="$d" --prop bold=true
    officecli set "$PPTX" "/slide[1]/table[1]/tr[$r]/tc[2]" --prop text="$a" --prop align=right
    officecli set "$PPTX" "/slide[1]/table[1]/tr[$r]/tc[3]" --prop text="$b" --prop align=right
    officecli set "$PPTX" "/slide[1]/table[1]/tr[$r]/tc[4]" --prop text="$c" --prop align=right
    officecli set "$PPTX" "/slide[1]/table[1]/tr[$r]/tc[5]" --prop text="$e" --prop align=right
done

# --- Slide 2: Section header rows spanning the full table ---
officecli add "$PPTX" /presentation/slides --type slide
officecli add "$PPTX" '/slide[2]' --type shape \
    --prop text="Full-Width Section Headers" --prop size=28 --prop bold=true \
    --prop x=0.5in --prop y=0.3in --prop width=12in --prop height=0.6in

officecli add "$PPTX" '/slide[2]' --type table \
    --prop x=0.5in --prop y=1.2in --prop width=12in --prop height=4.5in \
    --prop rows=9 --prop cols=4 --prop headerFill=1F3864

# Header
for entry in "1:Item" "2:Owner" "3:Due" "4:Status"; do
    c="${entry%%:*}"; t="${entry#*:}"
    officecli set "$PPTX" "/slide[2]/table[1]/tr[1]/tc[$c]" \
        --prop text="$t" --prop bold=true --prop color=FFFFFF
done

# Section: "Phase 1" — spans all 4 columns
officecli set "$PPTX" '/slide[2]/table[1]/tr[2]/tc[1]' \
    --prop text="◆ Phase 1 — Discovery" --prop bold=true --prop fill=FFE699 \
    --prop gridSpan=4
officecli set "$PPTX" '/slide[2]/table[1]/tr[3]/tc[1]' --prop text="Stakeholder interviews"
officecli set "$PPTX" '/slide[2]/table[1]/tr[3]/tc[2]' --prop text="Alice"
officecli set "$PPTX" '/slide[2]/table[1]/tr[3]/tc[3]' --prop text="Mar 15"
officecli set "$PPTX" '/slide[2]/table[1]/tr[3]/tc[4]' --prop text="✓ Done" --prop color=00B050
officecli set "$PPTX" '/slide[2]/table[1]/tr[4]/tc[1]' --prop text="Market research"
officecli set "$PPTX" '/slide[2]/table[1]/tr[4]/tc[2]' --prop text="Bob"
officecli set "$PPTX" '/slide[2]/table[1]/tr[4]/tc[3]' --prop text="Mar 30"
officecli set "$PPTX" '/slide[2]/table[1]/tr[4]/tc[4]' --prop text="✓ Done" --prop color=00B050

# Section: "Phase 2"
officecli set "$PPTX" '/slide[2]/table[1]/tr[5]/tc[1]' \
    --prop text="◆ Phase 2 — Design" --prop bold=true --prop fill=C6E0B4 \
    --prop gridSpan=4
officecli set "$PPTX" '/slide[2]/table[1]/tr[6]/tc[1]' --prop text="Architecture spec"
officecli set "$PPTX" '/slide[2]/table[1]/tr[6]/tc[2]' --prop text="Carol"
officecli set "$PPTX" '/slide[2]/table[1]/tr[6]/tc[3]' --prop text="Apr 20"
officecli set "$PPTX" '/slide[2]/table[1]/tr[6]/tc[4]' --prop text="◐ WIP" --prop color=FFC000

# Section: "Phase 3"
officecli set "$PPTX" '/slide[2]/table[1]/tr[7]/tc[1]' \
    --prop text="◆ Phase 3 — Build" --prop bold=true --prop fill=F4B084 \
    --prop gridSpan=4
officecli set "$PPTX" '/slide[2]/table[1]/tr[8]/tc[1]' --prop text="Backend services"
officecli set "$PPTX" '/slide[2]/table[1]/tr[8]/tc[2]' --prop text="Dave"
officecli set "$PPTX" '/slide[2]/table[1]/tr[8]/tc[3]' --prop text="Jun 15"
officecli set "$PPTX" '/slide[2]/table[1]/tr[8]/tc[4]' --prop text="◯ Not started"
officecli set "$PPTX" '/slide[2]/table[1]/tr[9]/tc[1]' --prop text="Frontend UI"
officecli set "$PPTX" '/slide[2]/table[1]/tr[9]/tc[2]' --prop text="Eve"
officecli set "$PPTX" '/slide[2]/table[1]/tr[9]/tc[3]' --prop text="Jul 01"
officecli set "$PPTX" '/slide[2]/table[1]/tr[9]/tc[4]' --prop text="◯ Not started"

officecli close "$PPTX"
officecli validate "$PPTX"
echo "Created: $PPTX"
