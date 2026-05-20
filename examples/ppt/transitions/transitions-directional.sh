#!/bin/bash
# Directional transitions — push, cover, uncover/pull, wipe.
#
# Direction support is NOT uniform across the family:
#   - push:           up/down/left/right                       (4 dirs)
#   - wipe:           up/down/left/right                       (4 dirs)
#   - cover/uncover:  up/down/left/right + leftup/rightup/
#                     leftdown/rightdown                       (8 dirs)
#   - pull:           alias for uncover                        (same 8 dirs)
#
# Mismatching the family/direction triggers an OOXML schema-level error
# from officecli (e.g. 'push-leftup' is rejected — push only supports the
# four cardinal directions). See transitions-basic.sh for the no-direction
# transitions (cut/fade/dissolve/flash).

set -e

DIR="$(dirname "$0")"
PPTX="$DIR/transitions-directional.pptx"

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

add_demo_slide ""            "Directional Transitions"  "1F3864"

# push: 4 cardinal directions
for d in up down left right; do
    add_demo_slide "push-$d" "push-$d" "2E5C8A"
done

# wipe: 4 cardinal directions
for d in up down left right; do
    add_demo_slide "wipe-$d" "wipe-$d" "4F7C3A"
done

# cover: 8 directions (4 cardinal + 4 diagonal corner)
for d in up down left right leftup rightup leftdown rightdown; do
    add_demo_slide "cover-$d" "cover-$d" "8A5A2B"
done

# uncover (a.k.a. pull): 8 directions
for d in up down left right leftup rightup leftdown rightdown; do
    add_demo_slide "uncover-$d" "uncover-$d" "7030A0"
done

officecli close "$PPTX"
officecli validate "$PPTX"
echo "Created: $PPTX"
