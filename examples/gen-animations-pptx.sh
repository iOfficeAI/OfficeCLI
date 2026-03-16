#!/bin/bash
# Generate a presentation showcasing all animation effects
# Each slide demonstrates a different category of animations
set -e

OUT="$(dirname "$0")/gen-animations-pptx.pptx"
rm -f "$OUT"
officecli create "$OUT"
officecli open "$OUT"

###############################################################################
# SLIDE 1 — Title
###############################################################################
echo "  -> Slide 1: Title"
officecli add "$OUT" / --type slide --prop layout=title
officecli set "$OUT" /slide[1] --prop background=radial:0D1B2A-1B4F72-bl
officecli set "$OUT" '/slide[1]/placeholder[centertitle]' \
  --prop text="Animation Showcase" --prop color=FFFFFF --prop size=48
officecli set "$OUT" '/slide[1]/placeholder[subtitle]' \
  --prop text="Every animation effect in officecli" --prop color=85C1E9 --prop size=22
officecli set "$OUT" /slide[1] --prop transition=fade

###############################################################################
# SLIDE 2 — Entrance Animations
###############################################################################
echo "  -> Slide 2: Entrance Animations"
officecli add "$OUT" / --type slide --prop title="Entrance Effects"
officecli set "$OUT" /slide[2] --prop background=1B2838
officecli set "$OUT" '/slide[2]/shape[1]' --prop color=FFFFFF --prop size=28

# appear
officecli add "$OUT" '/slide[2]' --type shape \
  --prop text="appear" --prop font=Consolas --prop size=14 --prop color=FFFFFF \
  --prop fill=2E86C1 --prop preset=roundRect \
  --prop x=1cm --prop y=4cm --prop width=5cm --prop height=2cm
officecli set "$OUT" '/slide[2]/shape[2]' --prop animation=appear-entrance-500

# fade
officecli add "$OUT" '/slide[2]' --type shape \
  --prop text="fade" --prop font=Consolas --prop size=14 --prop color=FFFFFF \
  --prop fill=27AE60 --prop preset=roundRect \
  --prop x=7cm --prop y=4cm --prop width=5cm --prop height=2cm
officecli set "$OUT" '/slide[2]/shape[3]' --prop animation=fade-entrance-800

# fly
officecli add "$OUT" '/slide[2]' --type shape \
  --prop text="fly" --prop font=Consolas --prop size=14 --prop color=FFFFFF \
  --prop fill=E74C3C --prop preset=roundRect \
  --prop x=13cm --prop y=4cm --prop width=5cm --prop height=2cm
officecli set "$OUT" '/slide[2]/shape[4]' --prop animation=fly-entrance-600

# zoom
officecli add "$OUT" '/slide[2]' --type shape \
  --prop text="zoom" --prop font=Consolas --prop size=14 --prop color=FFFFFF \
  --prop fill=8E44AD --prop preset=roundRect \
  --prop x=19cm --prop y=4cm --prop width=5cm --prop height=2cm
officecli set "$OUT" '/slide[2]/shape[5]' --prop animation=zoom-entrance-700

# wipe
officecli add "$OUT" '/slide[2]' --type shape \
  --prop text="wipe" --prop font=Consolas --prop size=14 --prop color=FFFFFF \
  --prop fill=F39C12 --prop preset=roundRect \
  --prop x=1cm --prop y=7.5cm --prop width=5cm --prop height=2cm
officecli set "$OUT" '/slide[2]/shape[6]' --prop animation=wipe-entrance-600

# bounce
officecli add "$OUT" '/slide[2]' --type shape \
  --prop text="bounce" --prop font=Consolas --prop size=14 --prop color=FFFFFF \
  --prop fill=1ABC9C --prop preset=roundRect \
  --prop x=7cm --prop y=7.5cm --prop width=5cm --prop height=2cm
officecli set "$OUT" '/slide[2]/shape[7]' --prop animation=bounce-entrance-800

# float
officecli add "$OUT" '/slide[2]' --type shape \
  --prop text="float" --prop font=Consolas --prop size=14 --prop color=FFFFFF \
  --prop fill=E67E22 --prop preset=roundRect \
  --prop x=13cm --prop y=7.5cm --prop width=5cm --prop height=2cm
officecli set "$OUT" '/slide[2]/shape[8]' --prop animation=float-entrance-700

# split
officecli add "$OUT" '/slide[2]' --type shape \
  --prop text="split" --prop font=Consolas --prop size=14 --prop color=FFFFFF \
  --prop fill=2980B9 --prop preset=roundRect \
  --prop x=19cm --prop y=7.5cm --prop width=5cm --prop height=2cm
officecli set "$OUT" '/slide[2]/shape[9]' --prop animation=split-entrance-600

# wheel
officecli add "$OUT" '/slide[2]' --type shape \
  --prop text="wheel" --prop font=Consolas --prop size=14 --prop color=FFFFFF \
  --prop fill=C0392B --prop preset=roundRect \
  --prop x=1cm --prop y=11cm --prop width=5cm --prop height=2cm
officecli set "$OUT" '/slide[2]/shape[10]' --prop animation=wheel-entrance-800

# swivel
officecli add "$OUT" '/slide[2]' --type shape \
  --prop text="swivel" --prop font=Consolas --prop size=14 --prop color=FFFFFF \
  --prop fill=16A085 --prop preset=roundRect \
  --prop x=7cm --prop y=11cm --prop width=5cm --prop height=2cm
officecli set "$OUT" '/slide[2]/shape[11]' --prop animation=swivel-entrance-700

# checkerboard
officecli add "$OUT" '/slide[2]' --type shape \
  --prop text="checkerboard" --prop font=Consolas --prop size=12 --prop color=FFFFFF \
  --prop fill=D35400 --prop preset=roundRect \
  --prop x=13cm --prop y=11cm --prop width=5cm --prop height=2cm
officecli set "$OUT" '/slide[2]/shape[12]' --prop animation=checkerboard-entrance-600

# blinds
officecli add "$OUT" '/slide[2]' --type shape \
  --prop text="blinds" --prop font=Consolas --prop size=14 --prop color=FFFFFF \
  --prop fill=7D3C98 --prop preset=roundRect \
  --prop x=19cm --prop y=11cm --prop width=5cm --prop height=2cm
officecli set "$OUT" '/slide[2]/shape[13]' --prop animation=blinds-entrance-600

officecli set "$OUT" /slide[2] --prop transition=wipe

###############################################################################
# SLIDE 3 — Exit Animations
###############################################################################
echo "  -> Slide 3: Exit Animations"
officecli add "$OUT" / --type slide --prop title="Exit Effects"
officecli set "$OUT" /slide[3] --prop background=1B2838
officecli set "$OUT" '/slide[3]/shape[1]' --prop color=FFFFFF --prop size=28

# fade exit
officecli add "$OUT" '/slide[3]' --type shape \
  --prop text="fade out" --prop font=Consolas --prop size=14 --prop color=FFFFFF \
  --prop fill=E74C3C --prop preset=roundRect \
  --prop x=1cm --prop y=4cm --prop width=7cm --prop height=2.5cm
officecli set "$OUT" '/slide[3]/shape[2]' --prop animation=fade-exit-800

# fly exit
officecli add "$OUT" '/slide[3]' --type shape \
  --prop text="fly out" --prop font=Consolas --prop size=14 --prop color=FFFFFF \
  --prop fill=2E86C1 --prop preset=roundRect \
  --prop x=9cm --prop y=4cm --prop width=7cm --prop height=2.5cm
officecli set "$OUT" '/slide[3]/shape[3]' --prop animation=fly-exit-600

# zoom exit
officecli add "$OUT" '/slide[3]' --type shape \
  --prop text="zoom out" --prop font=Consolas --prop size=14 --prop color=FFFFFF \
  --prop fill=27AE60 --prop preset=roundRect \
  --prop x=17cm --prop y=4cm --prop width=7cm --prop height=2.5cm
officecli set "$OUT" '/slide[3]/shape[4]' --prop animation=zoom-exit-700

# dissolve exit
officecli add "$OUT" '/slide[3]' --type shape \
  --prop text="dissolve out" --prop font=Consolas --prop size=14 --prop color=FFFFFF \
  --prop fill=8E44AD --prop preset=roundRect \
  --prop x=1cm --prop y=8cm --prop width=7cm --prop height=2.5cm
officecli set "$OUT" '/slide[3]/shape[5]' --prop animation=dissolve-exit-600

# wipe exit
officecli add "$OUT" '/slide[3]' --type shape \
  --prop text="wipe out" --prop font=Consolas --prop size=14 --prop color=FFFFFF \
  --prop fill=F39C12 --prop preset=roundRect \
  --prop x=9cm --prop y=8cm --prop width=7cm --prop height=2.5cm
officecli set "$OUT" '/slide[3]/shape[6]' --prop animation=wipe-exit-600

# flash exit
officecli add "$OUT" '/slide[3]' --type shape \
  --prop text="flash out" --prop font=Consolas --prop size=14 --prop color=FFFFFF \
  --prop fill=1ABC9C --prop preset=roundRect \
  --prop x=17cm --prop y=8cm --prop width=7cm --prop height=2.5cm
officecli set "$OUT" '/slide[3]/shape[7]' --prop animation=flash-exit-500

officecli set "$OUT" /slide[3] --prop transition=push

###############################################################################
# SLIDE 4 — Emphasis Animations
###############################################################################
echo "  -> Slide 4: Emphasis Animations"
officecli add "$OUT" / --type slide --prop title="Emphasis Effects"
officecli set "$OUT" /slide[4] --prop background=1B2838
officecli set "$OUT" '/slide[4]/shape[1]' --prop color=FFFFFF --prop size=28

# spin
officecli add "$OUT" '/slide[4]' --type shape \
  --prop text="spin" --prop font=Consolas --prop size=16 --prop color=FFFFFF \
  --prop fill=E74C3C --prop preset=ellipse \
  --prop x=2cm --prop y=4.5cm --prop width=4.5cm --prop height=4.5cm
officecli set "$OUT" '/slide[4]/shape[2]' --prop animation=spin-emphasis-1000

# grow
officecli add "$OUT" '/slide[4]' --type shape \
  --prop text="grow" --prop font=Consolas --prop size=16 --prop color=FFFFFF \
  --prop fill=2E86C1 --prop preset=ellipse \
  --prop x=8cm --prop y=4.5cm --prop width=4.5cm --prop height=4.5cm
officecli set "$OUT" '/slide[4]/shape[3]' --prop animation=grow-emphasis-800

# wave
officecli add "$OUT" '/slide[4]' --type shape \
  --prop text="wave" --prop font=Consolas --prop size=16 --prop color=FFFFFF \
  --prop fill=27AE60 --prop preset=ellipse \
  --prop x=14cm --prop y=4.5cm --prop width=4.5cm --prop height=4.5cm
officecli set "$OUT" '/slide[4]/shape[4]' --prop animation=wave-emphasis-700

# bold flash
officecli add "$OUT" '/slide[4]' --type shape \
  --prop text="bold" --prop font=Consolas --prop size=16 --prop color=FFFFFF \
  --prop fill=8E44AD --prop preset=ellipse \
  --prop x=20cm --prop y=4.5cm --prop width=4.5cm --prop height=4.5cm
officecli set "$OUT" '/slide[4]/shape[5]' --prop animation=bold-emphasis-500

officecli set "$OUT" /slide[4] --prop transition=zoom

###############################################################################
# SLIDE 5 — Slide Transitions Gallery
###############################################################################
echo "  -> Slide 5: Transitions Gallery"
officecli add "$OUT" / --type slide --prop title="Slide Transitions"
officecli set "$OUT" /slide[5] --prop background=0D1B2A
officecli set "$OUT" '/slide[5]/shape[1]' --prop color=FFFFFF --prop size=28

TRANSITIONS="fade wipe push split zoom wheel cover reveal dissolve random blinds checker strips"
X=1
Y=4
COL=0
for TR in $TRANSITIONS; do
  PX=$(echo "$X + $COL * 6" | bc)cm
  officecli add "$OUT" '/slide[5]' --type shape \
    --prop text="$TR" --prop font=Consolas --prop size=12 --prop color=FFFFFF \
    --prop fill=2C3E50 --prop preset=roundRect --prop line=5DADE2 --prop linewidth=0.5pt \
    --prop x=${PX} --prop y=${Y}cm --prop width=5cm --prop height=1.8cm
  COL=$((COL + 1))
  if [ $COL -ge 4 ]; then
    COL=0
    Y=$(echo "$Y + 2.5" | bc)
  fi
done

officecli set "$OUT" /slide[5] --prop transition=split

###############################################################################
# SLIDE 6 — Timing & Triggers
###############################################################################
echo "  -> Slide 6: Timing & Triggers"
officecli add "$OUT" / --type slide --prop title="Timing & Triggers"
officecli set "$OUT" /slide[6] --prop background=1B2838
officecli set "$OUT" '/slide[6]/shape[1]' --prop color=FFFFFF --prop size=28

# Click trigger (default)
officecli add "$OUT" '/slide[6]' --type shape \
  --prop text="Click to animate\n(default trigger)" --prop font=Consolas --prop size=13 --prop color=FFFFFF \
  --prop fill=2E86C1 --prop preset=roundRect \
  --prop x=1cm --prop y=4cm --prop width=7cm --prop height=3cm
officecli set "$OUT" '/slide[6]/shape[2]' --prop animation=fade-entrance-500

# After previous
officecli add "$OUT" '/slide[6]' --type shape \
  --prop text="After previous\n(auto-follows)" --prop font=Consolas --prop size=13 --prop color=FFFFFF \
  --prop fill=27AE60 --prop preset=roundRect \
  --prop x=9cm --prop y=4cm --prop width=7cm --prop height=3cm
officecli set "$OUT" '/slide[6]/shape[3]' --prop animation=fly-entrance-500-after

# With previous
officecli add "$OUT" '/slide[6]' --type shape \
  --prop text="With previous\n(simultaneous)" --prop font=Consolas --prop size=13 --prop color=FFFFFF \
  --prop fill=E74C3C --prop preset=roundRect \
  --prop x=17cm --prop y=4cm --prop width=7cm --prop height=3cm
officecli set "$OUT" '/slide[6]/shape[4]' --prop animation=zoom-entrance-500-with

# Slow vs Fast
officecli add "$OUT" '/slide[6]' --type shape \
  --prop text="Slow (2000ms)" --prop font=Consolas --prop size=13 --prop color=FFFFFF \
  --prop fill=8E44AD --prop preset=roundRect \
  --prop x=1cm --prop y=9cm --prop width=11cm --prop height=3cm
officecli set "$OUT" '/slide[6]/shape[5]' --prop animation=wipe-entrance-2000

officecli add "$OUT" '/slide[6]' --type shape \
  --prop text="Fast (200ms)" --prop font=Consolas --prop size=13 --prop color=FFFFFF \
  --prop fill=F39C12 --prop preset=roundRect \
  --prop x=13cm --prop y=9cm --prop width=11cm --prop height=3cm
officecli set "$OUT" '/slide[6]/shape[6]' --prop animation=wipe-entrance-200

officecli set "$OUT" /slide[6] --prop transition=reveal
officecli set "$OUT" /slide[6] --prop advanceTime=5000

###############################################################################
# Done
###############################################################################
officecli close "$OUT"
echo ""
echo "Done! Output: $OUT"
echo "Open with: open \"$OUT\""
