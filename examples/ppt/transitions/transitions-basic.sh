#!/bin/bash
# Basic slide transitions — cut, fade, dissolve, flash, and the 'none' clear.
# These five tokens form the everyday transition vocabulary: cut = instant,
# fade = pixel cross-fade, dissolve = pixel-noise dissolve, flash = white
# flash-through, none = remove an existing transition.
#
# Each demo slide carries the transition that triggers AS THE SLIDE ENTERS
# (replacing the previous one). To experience them, open the .pptx and step
# through Slide Show mode — most differences only show in playback.

set -e

DIR="$(dirname "$0")"
PPTX="$DIR/transitions-basic.pptx"

rm -f "$PPTX"
officecli create "$PPTX"
officecli open "$PPTX"

# add_demo_slide N transition title-text bg-hex
add_demo_slide() {
    local n=$1 trans=$2 title=$3 bg=$4
    officecli add "$PPTX" / --type slide
    officecli add "$PPTX" "/slide[$n]" --type shape \
        --prop x=0 --prop y=0 --prop width=33.87cm --prop height=19.05cm \
        --prop fill="$bg"
    officecli add "$PPTX" "/slide[$n]" --type shape \
        --prop text="$title" --prop size=54 --prop bold=true --prop color=FFFFFF \
        --prop align=center \
        --prop x=2cm --prop y=7cm --prop width=29.87cm --prop height=4cm
    if [ -n "$trans" ]; then
        officecli set "$PPTX" "/slide[$n]" --prop transition="$trans"
    fi
}

add_demo_slide 1 ""         "Basic Transitions"         "1F3864"
add_demo_slide 2 "cut"      "cut — instant swap"        "C00000"
add_demo_slide 3 "fade"     "fade — pixel cross-fade"   "2E75B6"
add_demo_slide 4 "dissolve" "dissolve — speckle blend"  "7030A0"
add_demo_slide 5 "flash"    "flash — white flash-thru"  "BF8F00"

# Demonstrate the 'none' clear: slide 6 first gets fade, then we wipe it.
# Final readback on slide 6 should NOT have a transition key at all.
add_demo_slide 6 "fade"     "none — fade cleared"       "404040"
officecli set "$PPTX" "/slide[6]" --prop transition=none

officecli close "$PPTX"
officecli validate "$PPTX"
echo "Created: $PPTX"
