#!/bin/bash
# Shape-mask transitions — the new slide reveals through a growing geometric
# mask cut into the old slide. OOXML calls these CT_OptionalBlackTransition.
#
# Direction-less (NO -in/-out suffix; officecli will reject one):
#   circle, diamond, plus, wedge
#
# Direction-ful (-in / -out):
#   zoom
#
# Spoke-count (wheel-N where N = number of spokes, 1..8 typical):
#   wheel-1, wheel-2, wheel-3, wheel-4 (default), wheel-8
#
# Known officecli limitation: `transition=box` writes `<p:box>` which is
# not a valid `<p:transition>` child per OOXML — schema validator rejects
# the file. The token is accepted by the parser but the resulting .pptx
# fails `officecli validate`. Avoided here; tracked separately.

set -e

DIR="$(dirname "$0")"
PPTX="$DIR/transitions-shapes.pptx"

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

add_demo_slide ""           "Shape Transitions"  "1F3864"

# Direction-less geometric masks
for t in circle diamond plus wedge; do
    add_demo_slide "$t" "$t" "C00000"
done

# In/out direction masks
for combo in zoom-in zoom-out; do
    add_demo_slide "$combo" "$combo" "2E75B6"
done

# Wheel spokes: same shape, different spoke count
for n_spokes in 1 2 3 4 8; do
    add_demo_slide "wheel-$n_spokes" "wheel-$n_spokes ($n_spokes spokes)" "7030A0"
done

officecli close "$PPTX"
officecli validate "$PPTX"
echo "Created: $PPTX"
