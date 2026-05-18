#!/bin/bash
# 3D / dynamic transitions — PowerPoint 2010+ "Exciting" gallery. Each
# requires Office 2010+ to render; older PowerPoint silently falls back
# to fade (officecli emits an mc:AlternateContent wrapper with that
# fallback baked in).
#
# Direction families:
#   left/right (LeftRightDir):  switch, flip, ferris, gallery, conveyor, reveal
#   in/out     (InOutDir):      shred, flythrough, warp
#   up/down/left/right          (SlideDir):  vortex, glitter, pan, prism
#   horizontal/vertical:        doors, window
#   no direction:               ripple, honeycomb

set -e

DIR="$(dirname "$0")"
PPTX="$DIR/transitions-dynamic.pptx"

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

add_demo_slide ""               "Dynamic Transitions"  "1F3864"

# LeftRight family
for t in switch flip ferris gallery conveyor reveal; do
    add_demo_slide "$t-right" "$t-right" "2E5C8A"
done

# InOut family
for t in shred flythrough warp; do
    add_demo_slide "$t-out" "$t-out" "4F7C3A"
done

# SlideDir family — 4 cardinal
for t in vortex glitter pan prism; do
    for d in up right; do
        add_demo_slide "$t-$d" "$t-$d" "8A5A2B"
    done
done

# Horizontal/vertical orientation
for t in doors window; do
    for d in horizontal vertical; do
        add_demo_slide "$t-$d" "$t-$d" "7030A0"
    done
done

# Direction-less
for t in ripple honeycomb; do
    add_demo_slide "$t" "$t" "C00000"
done

officecli close "$PPTX"
officecli validate "$PPTX"
echo "Created: $PPTX"
