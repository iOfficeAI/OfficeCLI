#!/bin/bash
# Morph transition — PowerPoint 2016+ smooth tweening between two slides
# that share named objects. Same shape on adjacent slides with different
# x/y/width/height/rotation is interpolated as continuous motion.
#
# Morph option (-byobject / -byword / -bychar):
#   byobject (default) — shapes with matching IDs are paired and tweened
#   byword             — text body is morphed word-by-word
#   bychar             — text body is morphed character-by-character
#
# This trio is a starter demo. For a fuller scene-level showcase using
# this skill: see examples/product_launch_morph.pptx in the repo root.

set -e

DIR="$(dirname "$0")"
PPTX="$DIR/transitions-morph.pptx"

rm -f "$PPTX"
officecli create "$PPTX"
officecli open "$PPTX"

# Slide 1: starting state (no transition — this is the entry point).
officecli add "$PPTX" / --type slide
officecli add "$PPTX" '/slide[1]' --type shape \
    --prop x=0 --prop y=0 --prop width=33.87cm --prop height=19.05cm --prop fill=1F3864
officecli add "$PPTX" '/slide[1]' --type shape \
    --prop text="Morph" --prop size=72 --prop bold=true --prop color=FFFFFF --prop align=center \
    --prop x=2cm --prop y=7cm --prop width=29.87cm --prop height=4cm
# A named circle that will tween across slides 2/3/4.
officecli add "$PPTX" '/slide[1]' --type shape \
    --prop shape=ellipse --prop fill=FFC000 --prop name=morphBall \
    --prop x=2cm --prop y=14cm --prop width=3cm --prop height=3cm

# Slide 2: morph-byobject — same-named ball moves right and grows.
officecli add "$PPTX" / --type slide
officecli set "$PPTX" '/slide[2]' --prop transition=morph
officecli add "$PPTX" '/slide[2]' --type shape \
    --prop x=0 --prop y=0 --prop width=33.87cm --prop height=19.05cm --prop fill=2E5C8A
officecli add "$PPTX" '/slide[2]' --type shape \
    --prop text="morph (byobject — default)" --prop size=44 --prop bold=true \
    --prop color=FFFFFF --prop align=center \
    --prop x=2cm --prop y=2cm --prop width=29.87cm --prop height=3cm
officecli add "$PPTX" '/slide[2]' --type shape \
    --prop shape=ellipse --prop fill=FFC000 --prop name=morphBall \
    --prop x=15cm --prop y=10cm --prop width=6cm --prop height=6cm

# Slide 3: morph-byword — title text recomposes word-by-word.
officecli add "$PPTX" / --type slide
officecli set "$PPTX" '/slide[3]' --prop transition=morph-byword
officecli add "$PPTX" '/slide[3]' --type shape \
    --prop x=0 --prop y=0 --prop width=33.87cm --prop height=19.05cm --prop fill=4F7C3A
officecli add "$PPTX" '/slide[3]' --type shape \
    --prop text="morph byword tweens words" --prop size=44 --prop bold=true \
    --prop color=FFFFFF --prop align=center \
    --prop x=2cm --prop y=2cm --prop width=29.87cm --prop height=3cm
officecli add "$PPTX" '/slide[3]' --type shape \
    --prop shape=ellipse --prop fill=FFC000 --prop name=morphBall \
    --prop x=27cm --prop y=14cm --prop width=3cm --prop height=3cm

# Slide 4: morph-bychar — recomposes letter-by-letter.
officecli add "$PPTX" / --type slide
officecli set "$PPTX" '/slide[4]' --prop transition=morph-bychar
officecli add "$PPTX" '/slide[4]' --type shape \
    --prop x=0 --prop y=0 --prop width=33.87cm --prop height=19.05cm --prop fill=8A5A2B
officecli add "$PPTX" '/slide[4]' --type shape \
    --prop text="bychar tweens letters" --prop size=44 --prop bold=true \
    --prop color=FFFFFF --prop align=center \
    --prop x=2cm --prop y=2cm --prop width=29.87cm --prop height=3cm
officecli add "$PPTX" '/slide[4]' --type shape \
    --prop shape=ellipse --prop fill=FFC000 --prop name=morphBall \
    --prop x=14cm --prop y=14cm --prop width=4cm --prop height=4cm

officecli close "$PPTX"
officecli validate "$PPTX"
echo "Created: $PPTX"
