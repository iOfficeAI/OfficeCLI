#!/bin/bash
# Transition timing — speed token vs millisecond duration, plus the
# auto-advance / click-to-advance knobs.
#
# Speed token (legacy CT_SlideTransition @spd, PowerPoint 97+):
#   fast / medium|med / slow      (e.g. fade-slow)
#
# Duration in ms (CT_TransitionStartSoundAction @dur extLst, Office 2010+):
#   integer ms                    (e.g. fade-1500 → 1.5 seconds)
#
# Specifying both in the same combined token is allowed — the integer
# wins for new PowerPoint, the speed for legacy. Speed-only also writes
# transitionSpeed; duration-only writes transitionDuration on readback.
#
# Auto-advance:
#   advanceTime=<ms>     auto-advance after N milliseconds (or 'none' to clear)
#   advanceClick=false   disable click-to-advance (default true, stripped from XML when true)

set -e

DIR="$(dirname "$0")"
PPTX="$DIR/transitions-timing.pptx"

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

add_demo_slide ""           "Transition Timing"           "1F3864"

# Legacy speed tokens
add_demo_slide "fade-fast"  "fade-fast (legacy @spd)"     "C00000"
add_demo_slide "fade-med"   "fade-med  (legacy @spd)"     "2E75B6"
add_demo_slide "fade-slow"  "fade-slow (legacy @spd)"     "7030A0"

# Office 2010+ duration in ms
add_demo_slide "fade-500"   "fade-500ms"                  "4F7C3A"
add_demo_slide "fade-1500"  "fade-1500ms"                 "8A5A2B"
add_demo_slide "fade-3000"  "fade-3000ms"                 "404040"

# Auto-advance: slide stays for 2 seconds then advances on its own
n=$((n+1))
officecli add "$PPTX" / --type slide
officecli add "$PPTX" "/slide[$n]" --type shape \
    --prop x=0 --prop y=0 --prop width=33.87cm --prop height=19.05cm --prop fill=BF8F00
officecli add "$PPTX" "/slide[$n]" --type shape \
    --prop text="advanceTime=2000  (auto-advance after 2s)" --prop size=36 \
    --prop bold=true --prop color=FFFFFF --prop align=center \
    --prop x=2cm --prop y=7cm --prop width=29.87cm --prop height=4cm
officecli set "$PPTX" "/slide[$n]" --prop transition=fade --prop advanceTime=2000

# Disable click-to-advance: this slide will only advance via auto-time or arrow keys
n=$((n+1))
officecli add "$PPTX" / --type slide
officecli add "$PPTX" "/slide[$n]" --type shape \
    --prop x=0 --prop y=0 --prop width=33.87cm --prop height=19.05cm --prop fill=2E5C8A
officecli add "$PPTX" "/slide[$n]" --type shape \
    --prop text="advanceClick=false  (no click advance)" --prop size=36 \
    --prop bold=true --prop color=FFFFFF --prop align=center \
    --prop x=2cm --prop y=7cm --prop width=29.87cm --prop height=4cm
officecli set "$PPTX" "/slide[$n]" --prop transition=fade --prop advanceClick=false

officecli close "$PPTX"
officecli validate "$PPTX"
echo "Created: $PPTX"
