#!/bin/bash
set -e

FILE="Feline_Report.pptx"

echo "Creating Feline_Report.pptx..."
rm -f "$FILE"
officecli create "$FILE"

echo "Setting 16:9 aspect ratio..."
officecli set "$FILE" / --prop slideSize=16:9

echo "--- Slide 1: Cover ---"
officecli add "$FILE" / --type slide --prop layout=blank --prop background=1E1E1E --prop transition=morph
officecli add "$FILE" /slide[1] --type shape --prop preset=rect --prop text="猫星人地球潜伏观察报告" --prop x=1.9cm --prop y=3cm --prop width=30cm --prop height=4cm --prop color=FFFFFF --prop size=44 --prop bold=true --prop align=center --prop fill=none --prop line=none --prop name="TitleText"
officecli add "$FILE" /slide[1] --type shape --prop preset=rect --prop text="绝密资料 / 阶段性成果汇报" --prop x=1.9cm --prop y=6.5cm --prop width=30cm --prop height=2cm --prop color=CCCCCC --prop size=24 --prop align=center --prop fill=none --prop line=none --prop name="SubText"
officecli add "$FILE" /slide[1] --type shape --prop preset=ellipse --prop x=11.9cm --prop y=9cm --prop width=10cm --prop height=10cm --prop fill=FFD700 --prop line=none --prop name="!!TargetCircle"

echo "--- Slide 2: Observation 1 ---"
officecli add "$FILE" / --type slide --prop layout=blank --prop background=1E1E1E --prop transition=morph
officecli add "$FILE" /slide[2] --type shape --prop preset=ellipse --prop x=3cm --prop y=7.5cm --prop width=4cm --prop height=4cm --prop fill=FFD700 --prop line=none --prop name="!!TargetCircle"
officecli add "$FILE" /slide[2] --type shape --prop preset=rect --prop text="战术 01：键盘物理覆盖" --prop x=9cm --prop y=6cm --prop width=22cm --prop height=3cm --prop color=FFD700 --prop size=36 --prop bold=true --prop align=left --prop fill=none --prop line=none --prop name="Obs1Title"
officecli add "$FILE" /slide[2] --type shape --prop preset=rect --prop text="通过阻断人类的输入设备，成功降低地球人 45% 的工作效率。人类依然以为我们在'撒娇'。" --prop x=9cm --prop y=9.5cm --prop width=22cm --prop height=4cm --prop color=FFFFFF --prop size=24 --prop align=left --prop fill=none --prop line=none --prop name="Obs1Text"

echo "--- Slide 3: Observation 2 ---"
officecli add "$FILE" / --type slide --prop layout=blank --prop background=1E1E1E --prop transition=morph
officecli add "$FILE" /slide[3] --type shape --prop preset=ellipse --prop x=28cm --prop y=2cm --prop width=1cm --prop height=1cm --prop fill=FF004D --prop line=none --prop name="!!TargetCircle"
officecli add "$FILE" /slide[3] --type shape --prop preset=rect --prop text="战术 02：红点追逐伪装" --prop x=9cm --prop y=6cm --prop width=22cm --prop height=3cm --prop color=FF004D --prop size=36 --prop bold=true --prop align=left --prop fill=none --prop line=none --prop name="Obs2Title"
officecli add "$FILE" /slide[3] --type shape --prop preset=rect --prop text="假装被红色激光点吸引，实则在测试地球人的智力底线与耐心。实验证明：人类比我们更执着于红点。" --prop x=9cm --prop y=9.5cm --prop width=22cm --prop height=4cm --prop color=FFFFFF --prop size=24 --prop align=left --prop fill=none --prop line=none --prop name="Obs2Text"

echo "--- Slide 4: Conclusion ---"
officecli add "$FILE" / --type slide --prop layout=blank --prop background=1E1E1E --prop transition=morph
officecli add "$FILE" /slide[4] --type shape --prop preset=ellipse --prop x=0cm --prop y=0cm --prop width=15cm --prop height=15cm --prop fill=FF004D --prop line=none --prop name="!!TargetCircle"
officecli add "$FILE" /slide[4] --type shape --prop preset=rect --prop text="结论：同化完成度 99%" --prop x=18cm --prop y=6cm --prop width=14cm --prop height=3cm --prop color=FFFFFF --prop size=36 --prop bold=true --prop align=left --prop fill=none --prop line=none --prop name="ConcTitle"
officecli add "$FILE" /slide[4] --type shape --prop preset=rect --prop text="人类已自愿成为“铲屎官”。\n地球占领计划基本达成。\n下一步：控制罐头生产线。" --prop x=18cm --prop y=9.5cm --prop width=14cm --prop height=6cm --prop color=FFFFFF --prop size=24 --prop align=left --prop fill=none --prop line=none --prop name="ConcText"

echo "Presentation created successfully: $FILE"
