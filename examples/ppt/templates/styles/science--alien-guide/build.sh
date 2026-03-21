#!/bin/bash
set +H
set -e

F="Alien_Guide.pptx"
echo "Building $F..."
rm -f "$F"
officecli create "$F"

BG="0B0C10"
CYAN="66FCF1"
TEAL="45A29E"
WHITE="FFFFFF"
GRAY="C5C6C7"
DARK="1F2833"

a() { officecli add "$F" "$1" --type shape "${@:2}"; }
sl() { officecli add "$F" / --type slide "$@"; }

# Helper for scene actors to maintain consistency across slides for Morph animation
scene_actors() {
  local s="$1"
  local c_x="$2" c_y="$3" c_w="$4" c_o="$5"
  local r_x="$6" r_y="$7" r_w="$8" r_h="$9" r_o="${10}"
  local a1_x="${11}" a1_y="${12}"
  local a2_x="${13}" a2_y="${14}"
  local lt_x="${15}" lt_y="${16}" lt_w="${17}"
  local lb_x="${18}" lb_y="${19}" lb_w="${20}"

  a "$s" --prop name="!!bg-circ" --prop preset=ellipse --prop x="${c_x}cm" --prop y="${c_y}cm" --prop width="${c_w}cm" --prop height="${c_w}cm" --prop fill=$DARK --prop line=none --prop opacity="${c_o}"
  a "$s" --prop name="!!bg-rect" --prop preset=roundRect --prop x="${r_x}cm" --prop y="${r_y}cm" --prop width="${r_w}cm" --prop height="${r_h}cm" --prop fill=$TEAL --prop line=none --prop opacity="${r_o}"
  a "$s" --prop name="!!accent-1" --prop preset=ellipse --prop x="${a1_x}cm" --prop y="${a1_y}cm" --prop width=0.8cm --prop height=0.8cm --prop fill=$CYAN --prop line=none
  a "$s" --prop name="!!accent-2" --prop preset=ellipse --prop x="${a2_x}cm" --prop y="${a2_y}cm" --prop width=1.2cm --prop height=1.2cm --prop fill=$CYAN --prop line=none
  a "$s" --prop name="!!line-top" --prop preset=rect --prop x="${lt_x}cm" --prop y="${lt_y}cm" --prop width="${lt_w}cm" --prop height=0.2cm --prop fill=$CYAN --prop line=none
  a "$s" --prop name="!!line-bot" --prop preset=rect --prop x="${lb_x}cm" --prop y="${lb_y}cm" --prop width="${lb_w}cm" --prop height=0.2cm --prop fill=$TEAL --prop line=none
}

# Slide 1: Hero
echo "  S1..."
sl --prop background=$BG
scene_actors '/slide[1]' 20 4 15 0.5   1 2 12 12 0.1   5 15   30 2   2 1 8   24 18 8

a '/slide[1]' --prop text="外星人地球" --prop x=2cm --prop y=4cm --prop width=18cm --prop height=3cm --prop size=64 --prop bold=true --prop color=$WHITE --prop fill=none --prop line=none
a '/slide[1]' --prop text="生存指南" --prop x=2cm --prop y=7.5cm --prop width=18cm --prop height=3cm --prop size=64 --prop bold=true --prop color=$CYAN --prop fill=none --prop line=none

a '/slide[1]' --prop text="从伪装到精通 (An Alien's Guide to Earth)" --prop x=2.2cm --prop y=11.5cm --prop width=20cm --prop height=1.5cm --prop size=24 --prop color=$GRAY --prop fill=none --prop line=none
a '/slide[1]' --prop text="本安全手册专为刚抵达银河系猎户旋臂第三行星的访客编写，
帮助你完美融入人类社会。" --prop x=2.2cm --prop y=13.5cm --prop width=18cm --prop height=3cm --prop size=16 --prop color=$GRAY --prop fill=none --prop line=none --prop lineSpacing=1.5

# Slide 2: Statement
echo "  S2..."
sl --prop background=$BG --prop transition=morph
scene_actors '/slide[2]' 2 2 18 0.3   22 5 8 14 0.1   15 3   2 16   10 1 4   2 18 12

a '/slide[2]' --prop text="RULE NO.1" --prop x=18cm --prop y=4cm --prop width=12cm --prop height=1.5cm --prop size=20 --prop bold=true --prop color=$CYAN --prop fill=none --prop line=none
a '/slide[2]' --prop text="第一法则" --prop x=18cm --prop y=5.5cm --prop width=12cm --prop height=2cm --prop size=48 --prop bold=true --prop color=$WHITE --prop fill=none --prop line=none

a '/slide[2]' --prop text="永远不要试图与猫讲道理。" --prop x=6cm --prop y=9cm --prop width=24cm --prop height=4cm --prop size=54 --prop bold=true --prop color=$CYAN --prop fill=none --prop line=none --prop align=center

a '/slide[2]' --prop text="数据表明，它们才是这颗星球真正的统治者，
人类只是它们的“铲屎官”。" --prop x=6cm --prop y=14cm --prop width=24cm --prop height=3cm --prop size=18 --prop color=$GRAY --prop fill=none --prop line=none --prop lineSpacing=1.5 --prop align=center

# Slide 3: Pillars
echo "  S3..."
sl --prop background=$BG --prop transition=morph
scene_actors '/slide[3]' 10 8 14 0.6   2 2 30 6 0.05   2 2   31 16   14 1 6   14 18 6

a '/slide[3]' --prop text="人类三大迷惑行为" --prop x=2cm --prop y=2.5cm --prop width=20cm --prop height=2cm --prop size=40 --prop bold=true --prop color=$WHITE --prop fill=none --prop line=none

# Pillar 1
a '/slide[3]' --prop preset=roundRect --prop x=2cm --prop y=6cm --prop width=8cm --prop height=10cm --prop fill=$DARK --prop line=none
a '/slide[3]' --prop text="01" --prop x=3cm --prop y=7cm --prop width=3cm --prop height=1.5cm --prop size=28 --prop bold=true --prop color=$CYAN --prop fill=none --prop line=none
a '/slide[3]' --prop text="排队 (Queueing)" --prop x=3cm --prop y=9cm --prop width=6cm --prop height=1.5cm --prop size=20 --prop bold=true --prop color=$WHITE --prop fill=none --prop line=none
a '/slide[3]' --prop text="人类极其喜欢排成一条直线，这种奇特的几何排列会给他们带来莫名的安全感。" --prop x=3cm --prop y=11.5cm --prop width=6cm --prop height=4cm --prop size=14 --prop color=$GRAY --prop fill=none --prop line=none --prop lineSpacing=1.5

# Pillar 2
a '/slide[3]' --prop preset=roundRect --prop x=12.5cm --prop y=6cm --prop width=8cm --prop height=10cm --prop fill=$DARK --prop line=none
a '/slide[3]' --prop text="02" --prop x=13.5cm --prop y=7cm --prop width=3cm --prop height=1.5cm --prop size=28 --prop bold=true --prop color=$CYAN --prop fill=none --prop line=none
a '/slide[3]' --prop text="密码 (Passwords)" --prop x=13.5cm --prop y=9cm --prop width=6cm --prop height=1.5cm --prop size=20 --prop bold=true --prop color=$WHITE --prop fill=none --prop line=none
a '/slide[3]' --prop text="他们总是忘记自己设置的安全验证码，然后被迫重置成一模一样的。" --prop x=13.5cm --prop y=11.5cm --prop width=6cm --prop height=4cm --prop size=14 --prop color=$GRAY --prop fill=none --prop line=none --prop lineSpacing=1.5

# Pillar 3
a '/slide[3]' --prop preset=roundRect --prop x=23cm --prop y=6cm --prop width=8cm --prop height=10cm --prop fill=$DARK --prop line=none
a '/slide[3]' --prop text="03" --prop x=24cm --prop y=7cm --prop width=3cm --prop height=1.5cm --prop size=28 --prop bold=true --prop color=$CYAN --prop fill=none --prop line=none
a '/slide[3]' --prop text="“好的收到”" --prop x=24cm --prop y=9cm --prop width=6cm --prop height=1.5cm --prop size=20 --prop bold=true --prop color=$WHITE --prop fill=none --prop line=none
a '/slide[3]' --prop text="人类常用此短语结束在线对话，但实际上有 80% 的概率并未接收任何实质信息。" --prop x=24cm --prop y=11.5cm --prop width=6cm --prop height=4cm --prop size=14 --prop color=$GRAY --prop fill=none --prop line=none --prop lineSpacing=1.5

# Slide 4: Evidence
echo "  S4..."
sl --prop background=$BG --prop transition=morph
scene_actors '/slide[4]' 4 4 12 0.8   18 4 12 12 0.1   16 10   8 16   2 4 2   18 16 12

a '/slide[4]' --prop text="99.9%" --prop x=4cm --prop y=7cm --prop width=12cm --prop height=5cm --prop size=80 --prop bold=true --prop color=$CYAN --prop fill=none --prop line=none --prop align=center

a '/slide[4]' --prop text="能量来源分析" --prop x=18cm --prop y=5cm --prop width=12cm --prop height=2cm --prop size=36 --prop bold=true --prop color=$WHITE --prop fill=none --prop line=none
a '/slide[4]' --prop text="早晨系统启动所需咖啡因比例" --prop x=18cm --prop y=8.5cm --prop width=12cm --prop height=1.5cm --prop size=18 --prop color=$CYAN --prop fill=none --prop line=none
a '/slide[4]' --prop text="警告！如果没有摄入这种被称为“咖啡”的黑色苦味液体，地球人的核心系统在早晨极易发生崩溃。" --prop x=18cm --prop y=11cm --prop width=12cm --prop height=4cm --prop size=16 --prop color=$GRAY --prop fill=none --prop line=none --prop lineSpacing=1.5

# Slide 5: CTA
echo "  S5..."
sl --prop background=$BG --prop transition=morph
scene_actors '/slide[5]' 14 5 20 0.4   8 6 18 8 0.1   10 5   24 14   13 4 8   13 16 8

a '/slide[5]' --prop text="祝你在地球好运！" --prop x=6cm --prop y=7cm --prop width=22cm --prop height=3cm --prop size=54 --prop bold=true --prop color=$WHITE --prop fill=none --prop line=none --prop align=center

a '/slide[5]' --prop text="切记收好你的触角，保持双足行走，并随时保持尴尬又不失礼貌的微笑。" --prop x=6cm --prop y=11cm --prop width=22cm --prop height=2cm --prop size=16 --prop color=$GRAY --prop fill=none --prop line=none --prop align=center

a '/slide[5]' --prop preset=roundRect --prop x=12.5cm --prop y=14cm --prop width=9cm --prop height=1.5cm --prop fill=$CYAN --prop line=none
a '/slide[5]' --prop text="启动伪装程序 [ ENGAGE ]" --prop x=12.5cm --prop y=14cm --prop width=9cm --prop height=1.5cm --prop size=14 --prop bold=true --prop color=$DARK --prop fill=none --prop line=none --prop align=center --prop valign=center

echo "Done!"
