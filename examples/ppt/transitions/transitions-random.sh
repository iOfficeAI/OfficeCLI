#!/bin/bash
# Random transitions — PowerPoint picks the actual animation at render
# time, so the .pptx only captures the intent, not the specific motion.
#
#   newsflash — newspaper-style spin-and-zoom (one fixed animation, but
#               grouped with the random family because it's pre-2010
#               legacy and rarely used outside this slot).
#   random    — PowerPoint chooses a random transition each time you
#               enter Slide Show mode.

set -e

DIR="$(dirname "$0")"
PPTX="$DIR/transitions-random.pptx"

rm -f "$PPTX"
officecli create "$PPTX"
officecli open "$PPTX"

n=0
add_demo_slide() {
    n=$((n+1))
    local trans=$1 title=$2 bg=$3
    officecli add "$PPTX" / --type slide
    officecli add "$PPTX" "/slide[$n]" --type shape \
        --prop x=0 --prop y=0 --prop width=33.87cm --prop height=19.05cm --prop fill="$bg"
    officecli add "$PPTX" "/slide[$n]" --type shape \
        --prop text="$title" --prop size=44 --prop bold=true --prop color=FFFFFF --prop align=center \
        --prop x=2cm --prop y=7cm --prop width=29.87cm --prop height=4cm
    if [ -n "$trans" ]; then officecli set "$PPTX" "/slide[$n]" --prop transition="$trans"; fi
}

add_demo_slide ""           "Random Transitions"  "1F3864"
add_demo_slide "newsflash"  "newsflash"           "C00000"
# Run Slide Show twice on this deck — slide 3 should animate differently each time.
add_demo_slide "random"     "random (re-rolls each play)" "2E75B6"
add_demo_slide "random"     "random (different again)"    "7030A0"

officecli close "$PPTX"
officecli validate "$PPTX"
echo "Created: $PPTX"
