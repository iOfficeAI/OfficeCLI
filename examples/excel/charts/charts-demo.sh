#!/bin/bash
# Generate an Excel chart showcase document
# Contains 6 chart types: clustered bar, smooth line, pie, stacked area, radar, doughnut
# Demonstrates officecli's Excel chart generation capabilities

set -e

XLSX="$(dirname "$0")/charts-demo.xlsx"
echo ""
echo "=========================================="
echo "Generating Excel chart showcase: $XLSX"
echo "=========================================="

rm -f "$XLSX"
officecli create "$XLSX"
officecli open "$XLSX"

###############################################################################
# 1. Populate data
###############################################################################
echo "  -> Populating sales data"

# Header (different colors per region)
officecli set "$XLSX" '/Sheet1/A1' --prop value="Month"  --prop font.bold=true --prop fill=2F5496 --prop font.color=FFFFFF --prop alignment.horizontal=center
officecli set "$XLSX" '/Sheet1/B1' --prop value="East Region" --prop font.bold=true --prop fill=4472C4 --prop font.color=FFFFFF --prop alignment.horizontal=center
officecli set "$XLSX" '/Sheet1/C1' --prop value="South Region" --prop font.bold=true --prop fill=5B9BD5 --prop font.color=FFFFFF --prop alignment.horizontal=center
officecli set "$XLSX" '/Sheet1/D1' --prop value="North Region" --prop font.bold=true --prop fill=70AD47 --prop font.color=FFFFFF --prop alignment.horizontal=center
officecli set "$XLSX" '/Sheet1/E1' --prop value="West Region" --prop font.bold=true --prop fill=FFC000 --prop font.color=000000 --prop alignment.horizontal=center

# 12 months of data
declare -a MONTHS=("Jan" "Feb" "Mar" "Apr" "May" "Jun" "Jul" "Aug" "Sep" "Oct" "Nov" "Dec")
declare -a EAST=(120 135 148 162 155 178 195 210 188 172 165 198)
declare -a SOUTH=(95 108 115 128 142 155 168 175 160 148 135 158)
declare -a NORTH=(88 92 105 118 125 138 145 152 140 130 122 142)
declare -a WEST=(72 78 85 95 102 115 125 132 120 110 98 118)

for i in $(seq 0 11); do
    row=$((i + 2))
    officecli set "$XLSX" "/Sheet1/A${row}" --prop "value=${MONTHS[$i]}" --prop alignment.horizontal=center
    officecli set "$XLSX" "/Sheet1/B${row}" --prop "value=${EAST[$i]}"  --prop 'numFmt=#,##0' --prop alignment.horizontal=center
    officecli set "$XLSX" "/Sheet1/C${row}" --prop "value=${SOUTH[$i]}" --prop 'numFmt=#,##0' --prop alignment.horizontal=center
    officecli set "$XLSX" "/Sheet1/D${row}" --prop "value=${NORTH[$i]}" --prop 'numFmt=#,##0' --prop alignment.horizontal=center
    officecli set "$XLSX" "/Sheet1/E${row}" --prop "value=${WEST[$i]}"  --prop 'numFmt=#,##0' --prop alignment.horizontal=center
done

echo "  Done: Data populated"

###############################################################################
# 2. Clustered bar chart
###############################################################################
echo "  -> Chart 1: Clustered bar chart"

CHART1_REL=$(officecli add-part "$XLSX" /Sheet1 --type chart 2>&1 | grep -o 'relId=[^ ]*' | cut -d= -f2)

officecli raw-set "$XLSX" '/Sheet1/chart[1]' --xpath "/c:chartSpace" --action replace --xml '
<c:chartSpace>
  <c:chart>
    <c:title>
      <c:tx><c:rich><a:bodyPr /><a:lstStyle />
        <a:p><a:pPr><a:defRPr sz="1400" b="1"><a:solidFill><a:srgbClr val="333333" /></a:solidFill></a:defRPr></a:pPr>
        <a:r><a:rPr lang="en-US" sz="1400" b="1" /><a:t>2025 Monthly Sales by Region (10K)</a:t></a:r></a:p>
      </c:rich></c:tx>
      <c:overlay val="0" />
    </c:title>
    <c:plotArea>
      <c:layout />
      <c:barChart>
        <c:barDir val="col" /><c:grouping val="clustered" /><c:varyColors val="0" />
        <c:ser>
          <c:idx val="0" /><c:order val="0" />
          <c:tx><c:strRef><c:f>Sheet1!$B$1</c:f></c:strRef></c:tx>
          <c:spPr><a:solidFill><a:srgbClr val="4472C4" /></a:solidFill><a:ln w="0"><a:noFill /></a:ln></c:spPr>
          <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$13</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$B$2:$B$13</c:f></c:numRef></c:val>
        </c:ser>
        <c:ser>
          <c:idx val="1" /><c:order val="1" />
          <c:tx><c:strRef><c:f>Sheet1!$C$1</c:f></c:strRef></c:tx>
          <c:spPr><a:solidFill><a:srgbClr val="ED7D31" /></a:solidFill><a:ln w="0"><a:noFill /></a:ln></c:spPr>
          <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$13</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$C$2:$C$13</c:f></c:numRef></c:val>
        </c:ser>
        <c:ser>
          <c:idx val="2" /><c:order val="2" />
          <c:tx><c:strRef><c:f>Sheet1!$D$1</c:f></c:strRef></c:tx>
          <c:spPr><a:solidFill><a:srgbClr val="70AD47" /></a:solidFill><a:ln w="0"><a:noFill /></a:ln></c:spPr>
          <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$13</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$D$2:$D$13</c:f></c:numRef></c:val>
        </c:ser>
        <c:ser>
          <c:idx val="3" /><c:order val="3" />
          <c:tx><c:strRef><c:f>Sheet1!$E$1</c:f></c:strRef></c:tx>
          <c:spPr><a:solidFill><a:srgbClr val="FFC000" /></a:solidFill><a:ln w="0"><a:noFill /></a:ln></c:spPr>
          <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$13</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$E$2:$E$13</c:f></c:numRef></c:val>
        </c:ser>
        <c:axId val="1" /><c:axId val="2" />
      </c:barChart>
      <c:catAx><c:axId val="1" /><c:scaling><c:orientation val="minMax" /></c:scaling><c:delete val="0" /><c:axPos val="b" /><c:crossAx val="2" /></c:catAx>
      <c:valAx><c:axId val="2" /><c:scaling><c:orientation val="minMax" /></c:scaling><c:delete val="0" /><c:axPos val="l" /><c:numFmt formatCode="#,##0" sourceLinked="0" /><c:crossAx val="1" /></c:valAx>
    </c:plotArea>
    <c:legend><c:legendPos val="b" /><c:overlay val="0" /></c:legend>
    <c:plotVisOnly val="1" />
  </c:chart>
</c:chartSpace>'

officecli raw-set "$XLSX" '/Sheet1/drawing' --xpath "//xdr:wsDr" --action append --xml "
<xdr:twoCellAnchor>
  <xdr:from><xdr:col>6</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
  <xdr:to><xdr:col>15</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>15</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
  <xdr:graphicFrame macro=\"\">
    <xdr:nvGraphicFramePr><xdr:cNvPr id=\"2\" name=\"Chart 1\" /><xdr:cNvGraphicFramePr /></xdr:nvGraphicFramePr>
    <xdr:xfrm><a:off x=\"0\" y=\"0\" /><a:ext cx=\"0\" cy=\"0\" /></xdr:xfrm>
    <a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:chart r:id=\"${CHART1_REL}\" /></a:graphicData></a:graphic>
  </xdr:graphicFrame>
  <xdr:clientData />
</xdr:twoCellAnchor>"

echo "  Done: Clustered bar chart"

###############################################################################
# 3. Smooth line chart (with data markers)
###############################################################################
echo "  -> Chart 2: Smooth line chart"

CHART2_REL=$(officecli add-part "$XLSX" /Sheet1 --type chart 2>&1 | grep -o 'relId=[^ ]*' | cut -d= -f2)

officecli raw-set "$XLSX" '/Sheet1/chart[2]' --xpath "/c:chartSpace" --action replace --xml '
<c:chartSpace>
  <c:chart>
    <c:title>
      <c:tx><c:rich><a:bodyPr /><a:lstStyle />
        <a:p><a:pPr><a:defRPr sz="1400" b="1"><a:solidFill><a:srgbClr val="333333" /></a:solidFill></a:defRPr></a:pPr>
        <a:r><a:rPr lang="en-US" sz="1400" b="1" /><a:t>Sales Trend Line Chart</a:t></a:r></a:p>
      </c:rich></c:tx>
      <c:overlay val="0" />
    </c:title>
    <c:plotArea>
      <c:layout />
      <c:lineChart>
        <c:grouping val="standard" /><c:varyColors val="0" />
        <c:ser>
          <c:idx val="0" /><c:order val="0" />
          <c:tx><c:strRef><c:f>Sheet1!$B$1</c:f></c:strRef></c:tx>
          <c:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:srgbClr val="4472C4" /></a:solidFill><a:round /></a:ln></c:spPr>
          <c:marker><c:symbol val="circle" /><c:size val="6" /><c:spPr><a:solidFill><a:srgbClr val="4472C4" /></a:solidFill></c:spPr></c:marker>
          <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$13</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$B$2:$B$13</c:f></c:numRef></c:val>
          <c:smooth val="1" />
        </c:ser>
        <c:ser>
          <c:idx val="1" /><c:order val="1" />
          <c:tx><c:strRef><c:f>Sheet1!$C$1</c:f></c:strRef></c:tx>
          <c:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:srgbClr val="ED7D31" /></a:solidFill><a:round /></a:ln></c:spPr>
          <c:marker><c:symbol val="diamond" /><c:size val="6" /><c:spPr><a:solidFill><a:srgbClr val="ED7D31" /></a:solidFill></c:spPr></c:marker>
          <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$13</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$C$2:$C$13</c:f></c:numRef></c:val>
          <c:smooth val="1" />
        </c:ser>
        <c:ser>
          <c:idx val="2" /><c:order val="2" />
          <c:tx><c:strRef><c:f>Sheet1!$D$1</c:f></c:strRef></c:tx>
          <c:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:srgbClr val="70AD47" /></a:solidFill><a:round /></a:ln></c:spPr>
          <c:marker><c:symbol val="triangle" /><c:size val="6" /><c:spPr><a:solidFill><a:srgbClr val="70AD47" /></a:solidFill></c:spPr></c:marker>
          <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$13</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$D$2:$D$13</c:f></c:numRef></c:val>
          <c:smooth val="1" />
        </c:ser>
        <c:ser>
          <c:idx val="3" /><c:order val="3" />
          <c:tx><c:strRef><c:f>Sheet1!$E$1</c:f></c:strRef></c:tx>
          <c:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:srgbClr val="FFC000" /></a:solidFill><a:round /></a:ln></c:spPr>
          <c:marker><c:symbol val="square" /><c:size val="6" /><c:spPr><a:solidFill><a:srgbClr val="FFC000" /></a:solidFill></c:spPr></c:marker>
          <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$13</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$E$2:$E$13</c:f></c:numRef></c:val>
          <c:smooth val="1" />
        </c:ser>
        <c:marker val="1" />
        <c:axId val="10" /><c:axId val="20" />
      </c:lineChart>
      <c:catAx><c:axId val="10" /><c:scaling><c:orientation val="minMax" /></c:scaling><c:delete val="0" /><c:axPos val="b" /><c:crossAx val="20" /></c:catAx>
      <c:valAx><c:axId val="20" /><c:scaling><c:orientation val="minMax" /></c:scaling><c:delete val="0" /><c:axPos val="l" /><c:numFmt formatCode="#,##0" sourceLinked="0" /><c:crossAx val="10" /></c:valAx>
    </c:plotArea>
    <c:legend><c:legendPos val="b" /><c:overlay val="0" /></c:legend>
    <c:plotVisOnly val="1" />
  </c:chart>
</c:chartSpace>'

officecli raw-set "$XLSX" '/Sheet1/drawing' --xpath "//xdr:wsDr" --action append --xml "
<xdr:twoCellAnchor>
  <xdr:from><xdr:col>6</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>16</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
  <xdr:to><xdr:col>15</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>31</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
  <xdr:graphicFrame macro=\"\">
    <xdr:nvGraphicFramePr><xdr:cNvPr id=\"3\" name=\"Chart 2\" /><xdr:cNvGraphicFramePr /></xdr:nvGraphicFramePr>
    <xdr:xfrm><a:off x=\"0\" y=\"0\" /><a:ext cx=\"0\" cy=\"0\" /></xdr:xfrm>
    <a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:chart r:id=\"${CHART2_REL}\" /></a:graphicData></a:graphic>
  </xdr:graphicFrame>
  <xdr:clientData />
</xdr:twoCellAnchor>"

echo "  Done: Line chart"

###############################################################################
# 4. Pie chart
###############################################################################
echo "  -> Chart 3: Pie chart"

CHART3_REL=$(officecli add-part "$XLSX" /Sheet1 --type chart 2>&1 | grep -o 'relId=[^ ]*' | cut -d= -f2)

officecli raw-set "$XLSX" '/Sheet1/chart[3]' --xpath "/c:chartSpace" --action replace --xml '
<c:chartSpace>
  <c:chart>
    <c:title>
      <c:tx><c:rich><a:bodyPr /><a:lstStyle />
        <a:p><a:pPr><a:defRPr sz="1400" b="1" /></a:pPr>
        <a:r><a:rPr lang="en-US" sz="1400" b="1" /><a:t>Annual Regional Sales Share</a:t></a:r></a:p>
      </c:rich></c:tx>
      <c:overlay val="0" />
    </c:title>
    <c:plotArea>
      <c:layout />
      <c:pieChart>
        <c:varyColors val="1" />
        <c:ser>
          <c:idx val="0" /><c:order val="0" />
          <c:dPt><c:idx val="0" /><c:spPr><a:solidFill><a:srgbClr val="4472C4" /></a:solidFill></c:spPr></c:dPt>
          <c:dPt><c:idx val="1" /><c:spPr><a:solidFill><a:srgbClr val="ED7D31" /></a:solidFill></c:spPr></c:dPt>
          <c:dPt><c:idx val="2" /><c:spPr><a:solidFill><a:srgbClr val="70AD47" /></a:solidFill></c:spPr></c:dPt>
          <c:dPt><c:idx val="3" /><c:spPr><a:solidFill><a:srgbClr val="FFC000" /></a:solidFill></c:spPr></c:dPt>
          <c:dLbls>
            <c:showLegendKey val="0" /><c:showVal val="0" /><c:showCatName val="1" /><c:showSerName val="0" /><c:showPercent val="1" />
          </c:dLbls>
          <c:cat><c:strRef><c:f>Sheet1!$B$1:$E$1</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$B$2:$E$2</c:f></c:numRef></c:val>
        </c:ser>
      </c:pieChart>
    </c:plotArea>
    <c:legend><c:legendPos val="b" /><c:overlay val="0" /></c:legend>
  </c:chart>
</c:chartSpace>'

officecli raw-set "$XLSX" '/Sheet1/drawing' --xpath "//xdr:wsDr" --action append --xml "
<xdr:twoCellAnchor>
  <xdr:from><xdr:col>6</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>32</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
  <xdr:to><xdr:col>13</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>47</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
  <xdr:graphicFrame macro=\"\">
    <xdr:nvGraphicFramePr><xdr:cNvPr id=\"4\" name=\"Chart 3\" /><xdr:cNvGraphicFramePr /></xdr:nvGraphicFramePr>
    <xdr:xfrm><a:off x=\"0\" y=\"0\" /><a:ext cx=\"0\" cy=\"0\" /></xdr:xfrm>
    <a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:chart r:id=\"${CHART3_REL}\" /></a:graphicData></a:graphic>
  </xdr:graphicFrame>
  <xdr:clientData />
</xdr:twoCellAnchor>"

echo "  Done: Pie chart"

###############################################################################
# 5. Stacked area chart
###############################################################################
echo "  -> Chart 4: Stacked area chart"

CHART4_REL=$(officecli add-part "$XLSX" /Sheet1 --type chart 2>&1 | grep -o 'relId=[^ ]*' | cut -d= -f2)

officecli raw-set "$XLSX" '/Sheet1/chart[4]' --xpath "/c:chartSpace" --action replace --xml '
<c:chartSpace>
  <c:chart>
    <c:title>
      <c:tx><c:rich><a:bodyPr /><a:lstStyle />
        <a:p><a:pPr><a:defRPr sz="1400" b="1" /></a:pPr>
        <a:r><a:rPr lang="en-US" sz="1400" b="1" /><a:t>Stacked Area - Sales Composition</a:t></a:r></a:p>
      </c:rich></c:tx>
      <c:overlay val="0" />
    </c:title>
    <c:plotArea>
      <c:layout />
      <c:areaChart>
        <c:grouping val="stacked" /><c:varyColors val="0" />
        <c:ser>
          <c:idx val="0" /><c:order val="0" />
          <c:tx><c:strRef><c:f>Sheet1!$B$1</c:f></c:strRef></c:tx>
          <c:spPr><a:solidFill><a:srgbClr val="4472C4"><a:alpha val="80000" /></a:srgbClr></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="4472C4" /></a:solidFill></a:ln></c:spPr>
          <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$13</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$B$2:$B$13</c:f></c:numRef></c:val>
        </c:ser>
        <c:ser>
          <c:idx val="1" /><c:order val="1" />
          <c:tx><c:strRef><c:f>Sheet1!$C$1</c:f></c:strRef></c:tx>
          <c:spPr><a:solidFill><a:srgbClr val="ED7D31"><a:alpha val="80000" /></a:srgbClr></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="ED7D31" /></a:solidFill></a:ln></c:spPr>
          <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$13</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$C$2:$C$13</c:f></c:numRef></c:val>
        </c:ser>
        <c:ser>
          <c:idx val="2" /><c:order val="2" />
          <c:tx><c:strRef><c:f>Sheet1!$D$1</c:f></c:strRef></c:tx>
          <c:spPr><a:solidFill><a:srgbClr val="70AD47"><a:alpha val="80000" /></a:srgbClr></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="70AD47" /></a:solidFill></a:ln></c:spPr>
          <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$13</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$D$2:$D$13</c:f></c:numRef></c:val>
        </c:ser>
        <c:ser>
          <c:idx val="3" /><c:order val="3" />
          <c:tx><c:strRef><c:f>Sheet1!$E$1</c:f></c:strRef></c:tx>
          <c:spPr><a:solidFill><a:srgbClr val="FFC000"><a:alpha val="80000" /></a:srgbClr></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="FFC000" /></a:solidFill></a:ln></c:spPr>
          <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$13</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$E$2:$E$13</c:f></c:numRef></c:val>
        </c:ser>
        <c:axId val="30" /><c:axId val="40" />
      </c:areaChart>
      <c:catAx><c:axId val="30" /><c:scaling><c:orientation val="minMax" /></c:scaling><c:delete val="0" /><c:axPos val="b" /><c:crossAx val="40" /></c:catAx>
      <c:valAx><c:axId val="40" /><c:scaling><c:orientation val="minMax" /></c:scaling><c:delete val="0" /><c:axPos val="l" /><c:numFmt formatCode="#,##0" sourceLinked="0" /><c:crossAx val="30" /></c:valAx>
    </c:plotArea>
    <c:legend><c:legendPos val="b" /><c:overlay val="0" /></c:legend>
    <c:plotVisOnly val="1" />
  </c:chart>
</c:chartSpace>'

officecli raw-set "$XLSX" '/Sheet1/drawing' --xpath "//xdr:wsDr" --action append --xml "
<xdr:twoCellAnchor>
  <xdr:from><xdr:col>6</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>48</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
  <xdr:to><xdr:col>15</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>63</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
  <xdr:graphicFrame macro=\"\">
    <xdr:nvGraphicFramePr><xdr:cNvPr id=\"5\" name=\"Chart 4\" /><xdr:cNvGraphicFramePr /></xdr:nvGraphicFramePr>
    <xdr:xfrm><a:off x=\"0\" y=\"0\" /><a:ext cx=\"0\" cy=\"0\" /></xdr:xfrm>
    <a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:chart r:id=\"${CHART4_REL}\" /></a:graphicData></a:graphic>
  </xdr:graphicFrame>
  <xdr:clientData />
</xdr:twoCellAnchor>"

echo "  Done: Stacked area chart"

###############################################################################
# 6. Radar chart
###############################################################################
echo "  -> Chart 5: Radar chart"

CHART5_REL=$(officecli add-part "$XLSX" /Sheet1 --type chart 2>&1 | grep -o 'relId=[^ ]*' | cut -d= -f2)

officecli raw-set "$XLSX" '/Sheet1/chart[5]' --xpath "/c:chartSpace" --action replace --xml '
<c:chartSpace>
  <c:chart>
    <c:title>
      <c:tx><c:rich><a:bodyPr /><a:lstStyle />
        <a:p><a:pPr><a:defRPr sz="1400" b="1" /></a:pPr>
        <a:r><a:rPr lang="en-US" sz="1400" b="1" /><a:t>Regional Capability Radar (Q3)</a:t></a:r></a:p>
      </c:rich></c:tx>
      <c:overlay val="0" />
    </c:title>
    <c:plotArea>
      <c:layout />
      <c:radarChart>
        <c:radarStyle val="marker" /><c:varyColors val="0" />
        <c:ser>
          <c:idx val="0" /><c:order val="0" />
          <c:tx><c:strRef><c:f>Sheet1!$B$1</c:f></c:strRef></c:tx>
          <c:spPr><a:ln w="28575"><a:solidFill><a:srgbClr val="4472C4" /></a:solidFill></a:ln></c:spPr>
          <c:cat><c:strRef><c:f>Sheet1!$A$8:$A$10</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$B$8:$B$10</c:f></c:numRef></c:val>
        </c:ser>
        <c:ser>
          <c:idx val="1" /><c:order val="1" />
          <c:tx><c:strRef><c:f>Sheet1!$C$1</c:f></c:strRef></c:tx>
          <c:spPr><a:ln w="28575"><a:solidFill><a:srgbClr val="ED7D31" /></a:solidFill></a:ln></c:spPr>
          <c:cat><c:strRef><c:f>Sheet1!$A$8:$A$10</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$C$8:$C$10</c:f></c:numRef></c:val>
        </c:ser>
        <c:ser>
          <c:idx val="2" /><c:order val="2" />
          <c:tx><c:strRef><c:f>Sheet1!$D$1</c:f></c:strRef></c:tx>
          <c:spPr><a:ln w="28575"><a:solidFill><a:srgbClr val="70AD47" /></a:solidFill></a:ln></c:spPr>
          <c:cat><c:strRef><c:f>Sheet1!$A$8:$A$10</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$D$8:$D$10</c:f></c:numRef></c:val>
        </c:ser>
        <c:ser>
          <c:idx val="3" /><c:order val="3" />
          <c:tx><c:strRef><c:f>Sheet1!$E$1</c:f></c:strRef></c:tx>
          <c:spPr><a:ln w="28575"><a:solidFill><a:srgbClr val="FFC000" /></a:solidFill></a:ln></c:spPr>
          <c:cat><c:strRef><c:f>Sheet1!$A$8:$A$10</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$E$8:$E$10</c:f></c:numRef></c:val>
        </c:ser>
        <c:axId val="50" /><c:axId val="60" />
      </c:radarChart>
      <c:catAx><c:axId val="50" /><c:scaling><c:orientation val="minMax" /></c:scaling><c:delete val="0" /><c:axPos val="b" /><c:crossAx val="60" /></c:catAx>
      <c:valAx><c:axId val="60" /><c:scaling><c:orientation val="minMax" /></c:scaling><c:delete val="0" /><c:axPos val="l" /><c:crossAx val="50" /></c:valAx>
    </c:plotArea>
    <c:legend><c:legendPos val="b" /><c:overlay val="0" /></c:legend>
  </c:chart>
</c:chartSpace>'

officecli raw-set "$XLSX" '/Sheet1/drawing' --xpath "//xdr:wsDr" --action append --xml "
<xdr:twoCellAnchor>
  <xdr:from><xdr:col>6</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>64</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
  <xdr:to><xdr:col>13</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>79</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
  <xdr:graphicFrame macro=\"\">
    <xdr:nvGraphicFramePr><xdr:cNvPr id=\"6\" name=\"Chart 5\" /><xdr:cNvGraphicFramePr /></xdr:nvGraphicFramePr>
    <xdr:xfrm><a:off x=\"0\" y=\"0\" /><a:ext cx=\"0\" cy=\"0\" /></xdr:xfrm>
    <a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:chart r:id=\"${CHART5_REL}\" /></a:graphicData></a:graphic>
  </xdr:graphicFrame>
  <xdr:clientData />
</xdr:twoCellAnchor>"

echo "  Done: Radar chart"

###############################################################################
# 7. Doughnut chart
###############################################################################
echo "  -> Chart 6: Doughnut chart"

CHART6_REL=$(officecli add-part "$XLSX" /Sheet1 --type chart 2>&1 | grep -o 'relId=[^ ]*' | cut -d= -f2)

officecli raw-set "$XLSX" '/Sheet1/chart[6]' --xpath "/c:chartSpace" --action replace --xml '
<c:chartSpace>
  <c:chart>
    <c:title>
      <c:tx><c:rich><a:bodyPr /><a:lstStyle />
        <a:p><a:pPr><a:defRPr sz="1400" b="1" /></a:pPr>
        <a:r><a:rPr lang="en-US" sz="1400" b="1" /><a:t>Q4 Regional Sales Doughnut</a:t></a:r></a:p>
      </c:rich></c:tx>
      <c:overlay val="0" />
    </c:title>
    <c:plotArea>
      <c:layout />
      <c:doughnutChart>
        <c:varyColors val="1" />
        <c:ser>
          <c:idx val="0" /><c:order val="0" />
          <c:dPt><c:idx val="0" /><c:spPr><a:solidFill><a:srgbClr val="4472C4" /></a:solidFill></c:spPr></c:dPt>
          <c:dPt><c:idx val="1" /><c:spPr><a:solidFill><a:srgbClr val="ED7D31" /></a:solidFill></c:spPr></c:dPt>
          <c:dPt><c:idx val="2" /><c:spPr><a:solidFill><a:srgbClr val="70AD47" /></a:solidFill></c:spPr></c:dPt>
          <c:dPt><c:idx val="3" /><c:spPr><a:solidFill><a:srgbClr val="FFC000" /></a:solidFill></c:spPr></c:dPt>
          <c:dLbls>
            <c:showLegendKey val="0" /><c:showVal val="0" /><c:showCatName val="1" /><c:showSerName val="0" /><c:showPercent val="1" />
          </c:dLbls>
          <c:cat><c:strRef><c:f>Sheet1!$B$1:$E$1</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$B$13:$E$13</c:f></c:numRef></c:val>
        </c:ser>
        <c:holeSize val="50" />
      </c:doughnutChart>
    </c:plotArea>
    <c:legend><c:legendPos val="b" /><c:overlay val="0" /></c:legend>
  </c:chart>
</c:chartSpace>'

officecli raw-set "$XLSX" '/Sheet1/drawing' --xpath "//xdr:wsDr" --action append --xml "
<xdr:twoCellAnchor>
  <xdr:from><xdr:col>14</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>32</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
  <xdr:to><xdr:col>21</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>47</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
  <xdr:graphicFrame macro=\"\">
    <xdr:nvGraphicFramePr><xdr:cNvPr id=\"7\" name=\"Chart 6\" /><xdr:cNvGraphicFramePr /></xdr:nvGraphicFramePr>
    <xdr:xfrm><a:off x=\"0\" y=\"0\" /><a:ext cx=\"0\" cy=\"0\" /></xdr:xfrm>
    <a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:chart r:id=\"${CHART6_REL}\" /></a:graphicData></a:graphic>
  </xdr:graphicFrame>
  <xdr:clientData />
</xdr:twoCellAnchor>"

echo "  Done: Doughnut chart"

###############################################################################
# Validation
###############################################################################
officecli close "$XLSX"

echo ""
echo "=========================================="
echo "Validating file"
echo "=========================================="
officecli validate "$XLSX"
officecli view "$XLSX" outline
echo ""
ls -lh "$XLSX"
echo ""
echo "All done!"
