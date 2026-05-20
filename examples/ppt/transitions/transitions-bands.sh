#!/bin/bash
# Band / strip transitions — the slide reveals through parallel bands,
# checkerboard squares, or diagonal strips.
#
# Orientation modifier (-horizontal / -vertical):
#   blinds, venetian (alias for blinds), checker, checkerboard (alias),
#   comb, bars, randombar (alias for bars)
#
# Corner direction (-leftup / -rightup / -leftdown / -rightdown):
#   strips, diagonal (alias for strips)
#
# Orient + in/out (BOTH must be specified for explicit form):
#   split-vertical-in, split-horizontal-out, ...
#   Bare 'split' rounds back as 'split'; 'split-vertical' alone now
#   defaults dir=in (canonical readback: split-vertical-in).
#
# Aliases land on the canonical token at readback time:
#   checkerboard → checker, randombar → bars, diagonal → strips,
#   venetian → blinds.

set -e

DIR="$(dirname "$0")"
PPTX="$DIR/transitions-bands.pptx"

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

add_demo_slide ""                 "Band Transitions"             "1F3864"

# Orientation: vertical vs horizontal
for combo in blinds-horizontal blinds-vertical \
             checker-horizontal checker-vertical \
             comb-horizontal comb-vertical \
             bars-horizontal bars-vertical; do
    add_demo_slide "$combo" "$combo" "2E5C8A"
done

# Strips: 4 corner directions
for d in leftup rightup leftdown rightdown; do
    add_demo_slide "strips-$d" "strips-$d" "4F7C3A"
done

# Split: orient × in/out matrix
for orient in horizontal vertical; do
    for io in in out; do
        add_demo_slide "split-$orient-$io" "split-$orient-$io" "8A5A2B"
    done
done

# Alias demo — same XML, different input spelling
add_demo_slide "venetian-vertical"     "venetian-vertical (alias → blinds)"     "7030A0"
add_demo_slide "checkerboard-vertical" "checkerboard-vertical (alias → checker)" "7030A0"
add_demo_slide "randombar-vertical"    "randombar-vertical (alias → bars)"       "7030A0"
add_demo_slide "diagonal-leftdown"     "diagonal-leftdown (alias → strips)"     "7030A0"

officecli close "$PPTX"
officecli validate "$PPTX"
echo "Created: $PPTX"
