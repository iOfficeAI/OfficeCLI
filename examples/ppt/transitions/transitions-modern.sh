#!/bin/bash
# PowerPoint 2013+ "Exciting" / "Dynamic Content" transitions —
# the 12 presets stored as <p15:prstTrans prst="..."/> inside an
# mc:AlternateContent wrapper (PowerPoint <2013 plays the inline fade
# fallback). Token spelling matches the OOXML prst attribute
# (lowerCamelCase): fallOver, peelOff, pageCurlDouble, pageCurlSingle,
# etc. Box (covered in transitions-shapes) uses the same element.
#
# Direction modifier (-in / -out):
#   default is -in (no invX/invY attributes written)
#   -out sets invX="1" invY="1" — visually flips the transition axis on
#   presets with a directional component (wind, peelOff, pageCurl*,
#   airplane, origami, fallOver, drape). Symmetric presets (curtains,
#   fracture, crush, prestige) accept the suffix but render unchanged.

set -e

DIR="$(dirname "$0")"
PPTX="$DIR/transitions-modern.pptx"

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
        --prop text="$title" --prop size=40 --prop bold=true --prop color=FFFFFF --prop align=center \
        --prop x=2cm --prop y=7cm --prop width=29.87cm --prop height=4cm
    if [ -n "$trans" ]; then officecli set "$PPTX" "/slide[$n]" --prop transition="$trans"; fi
}

add_demo_slide ""                "Modern (p15) Transitions"  "1F3864"

# Each preset's bare form (= -in)
for t in fallOver drape curtains wind prestige fracture crush peelOff \
         pageCurlDouble pageCurlSingle airplane origami; do
    add_demo_slide "$t" "$t" "2E5C8A"
done

# A handful of -out variants showing the invX/invY flip on direction-sensitive presets
for t in wind peelOff pageCurlDouble airplane origami fallOver; do
    add_demo_slide "$t-out" "$t-out" "8A5A2B"
done

officecli close "$PPTX"
officecli validate "$PPTX"
echo "Created: $PPTX"
