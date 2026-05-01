#!/bin/bash
# Generate complex textbox test document
# Includes 10 textbox scenarios for testing officecli compatibility with complex textbox cases

set -e

OUT="complex-textbox-test.docx"

echo "Using CLI: officecli"
echo "Output file: $OUT"

# ==================== Create base document ====================
officecli create "$OUT"
officecli add "$OUT" /body --type paragraph --prop text="Complex Textbox Examples" --prop style=Heading1 --prop align=center
officecli add "$OUT" /body --type paragraph --prop text="The following contains multiple complex textbox scenarios for testing textbox behavior under various conditions."

# ==================== Scenario 1: Basic Textbox (with border and background + VML Fallback) ====================
officecli add "$OUT" /body --type paragraph --prop text="Scenario 1: Basic Textbox (with border and background)" --prop style=Heading2

officecli raw-set "$OUT" /document --xpath "//w:body" --action append --xml '
<w:p>
  <w:r>
    <w:rPr><w:noProof/></w:rPr>
    <mc:AlternateContent>
      <mc:Choice Requires="wps">
        <w:drawing>
          <wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
            <wp:simplePos x="0" y="0"/>
            <wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>
            <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
            <wp:extent cx="5400000" cy="1200000"/>
            <wp:effectExtent l="0" t="0" r="0" b="0"/>
            <wp:wrapTopAndBottom/>
            <wp:docPr id="1" name="TextBox 1"/>
            <a:graphic>
              <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                <wps:wsp>
                  <wps:cNvSpPr txBox="1"/>
                  <wps:spPr>
                    <a:xfrm><a:off x="0" y="0"/><a:ext cx="5400000" cy="1200000"/></a:xfrm>
                    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                    <a:solidFill><a:srgbClr val="E6F3FF"/></a:solidFill>
                    <a:ln w="25400"><a:solidFill><a:srgbClr val="0070C0"/></a:solidFill></a:ln>
                  </wps:spPr>
                  <wps:txbx>
                    <w:txbxContent>
                      <w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="28"/><w:color w:val="0070C0"/></w:rPr><w:t>Basic Textbox</w:t></w:r></w:p>
                      <w:p><w:r><w:t>This is a textbox with a blue border and light blue background. It contains a centered title and a normal paragraph.</w:t></w:r></w:p>
                    </w:txbxContent>
                  </wps:txbx>
                  <wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t"/>
                </wps:wsp>
              </a:graphicData>
            </a:graphic>
          </wp:anchor>
        </w:drawing>
      </mc:Choice>
      <mc:Fallback>
        <w:pict>
          <v:shapetype id="_x0000_t202" coordsize="21600,21600" path="m,l,21600r21600,l21600,xe">
            <v:stroke joinstyle="miter"/>
            <v:path gradientshapeok="t"/>
          </v:shapetype>
          <v:shape id="TextBox1" type="#_x0000_t202" style="width:425.2pt;height:94.5pt;mso-wrap-style:square" fillcolor="#E6F3FF" strokecolor="#0070C0" strokeweight="2pt">
            <v:textbox><w:txbxContent>
              <w:p><w:r><w:t>Basic Textbox (VML fallback)</w:t></w:r></w:p>
            </w:txbxContent></v:textbox>
          </v:shape>
        </w:pict>
      </mc:Fallback>
    </mc:AlternateContent>
  </w:r>
</w:p>'

echo "Done: Scenario 1: Basic Textbox"

# ==================== Scenario 2: Multi-paragraph Rich Text Textbox ====================
officecli add "$OUT" /body --type paragraph --prop text="Scenario 2: Multi-paragraph Rich Text Textbox" --prop style=Heading2

officecli raw-set "$OUT" /document --xpath "//w:body" --action append --xml '
<w:p>
  <w:r>
    <w:rPr><w:noProof/></w:rPr>
    <mc:AlternateContent>
      <mc:Choice Requires="wps">
        <w:drawing>
          <wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251660288" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
            <wp:simplePos x="0" y="0"/>
            <wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>
            <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
            <wp:extent cx="5400000" cy="2400000"/>
            <wp:effectExtent l="0" t="0" r="0" b="0"/>
            <wp:wrapTopAndBottom/>
            <wp:docPr id="2" name="TextBox 2"/>
            <a:graphic>
              <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                <wps:wsp>
                  <wps:cNvSpPr txBox="1"/>
                  <wps:spPr>
                    <a:xfrm><a:off x="0" y="0"/><a:ext cx="5400000" cy="2400000"/></a:xfrm>
                    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                    <a:solidFill><a:srgbClr val="FFFDE7"/></a:solidFill>
                    <a:ln w="19050"><a:solidFill><a:srgbClr val="FF8C00"/></a:solidFill><a:prstDash val="dash"/></a:ln>
                  </wps:spPr>
                  <wps:txbx>
                    <w:txbxContent>
                      <w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="32"/><w:color w:val="FF8C00"/></w:rPr><w:t>Rich Text Content</w:t></w:r></w:p>
                      <w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Bold</w:t></w:r><w:r><w:t xml:space="preserve"> | </w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>Italic</w:t></w:r><w:r><w:t xml:space="preserve"> | </w:t></w:r><w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>Underline</w:t></w:r><w:r><w:t xml:space="preserve"> | </w:t></w:r><w:r><w:rPr><w:strike/></w:rPr><w:t>Strikethrough</w:t></w:r></w:p>
                      <w:p><w:r><w:rPr><w:color w:val="FF0000"/><w:sz w:val="20"/></w:rPr><w:t>Red small</w:t></w:r><w:r><w:t xml:space="preserve"> </w:t></w:r><w:r><w:rPr><w:color w:val="00B050"/><w:sz w:val="36"/></w:rPr><w:t>Green large</w:t></w:r><w:r><w:t xml:space="preserve"> </w:t></w:r><w:r><w:rPr><w:color w:val="0000FF"/><w:sz w:val="28"/><w:b/><w:i/></w:rPr><w:t>Blue bold italic</w:t></w:r></w:p>
                      <w:p><w:r><w:rPr><w:highlight w:val="yellow"/></w:rPr><w:t>Yellow highlight</w:t></w:r><w:r><w:t xml:space="preserve"> </w:t></w:r><w:r><w:rPr><w:highlight w:val="green"/><w:color w:val="FFFFFF"/></w:rPr><w:t>Green highlight white</w:t></w:r></w:p>
                      <w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:i/><w:sz w:val="22"/></w:rPr><w:t>-- Right-aligned quote</w:t></w:r></w:p>
                    </w:txbxContent>
                  </wps:txbx>
                  <wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t"/>
                </wps:wsp>
              </a:graphicData>
            </a:graphic>
          </wp:anchor>
        </w:drawing>
      </mc:Choice>
    </mc:AlternateContent>
  </w:r>
</w:p>'

echo "Done: Scenario 2: Rich Text Textbox"

# ==================== Scenario 3: Textbox with Nested Table ====================
officecli add "$OUT" /body --type paragraph --prop text="Scenario 3: Textbox with Nested Table" --prop style=Heading2

officecli raw-set "$OUT" /document --xpath "//w:body" --action append --xml '
<w:p>
  <w:r>
    <w:rPr><w:noProof/></w:rPr>
    <mc:AlternateContent>
      <mc:Choice Requires="wps">
        <w:drawing>
          <wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251661312" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
            <wp:simplePos x="0" y="0"/>
            <wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>
            <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
            <wp:extent cx="5400000" cy="2000000"/>
            <wp:effectExtent l="0" t="0" r="0" b="0"/>
            <wp:wrapTopAndBottom/>
            <wp:docPr id="3" name="TextBox 3"/>
            <a:graphic>
              <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                <wps:wsp>
                  <wps:cNvSpPr txBox="1"/>
                  <wps:spPr>
                    <a:xfrm><a:off x="0" y="0"/><a:ext cx="5400000" cy="2000000"/></a:xfrm>
                    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                    <a:solidFill><a:srgbClr val="F5F5F5"/></a:solidFill>
                    <a:ln w="12700"><a:solidFill><a:srgbClr val="333333"/></a:solidFill></a:ln>
                  </wps:spPr>
                  <wps:txbx>
                    <w:txbxContent>
                      <w:p><w:r><w:rPr><w:b/><w:sz w:val="24"/></w:rPr><w:t>Table inside textbox:</w:t></w:r></w:p>
                      <w:tbl>
                        <w:tblPr>
                          <w:tblStyle w:val="TableGrid"/>
                          <w:tblW w:w="5000" w:type="pct"/>
                          <w:tblBorders>
                            <w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>
                            <w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>
                            <w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
                            <w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>
                            <w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>
                            <w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>
                          </w:tblBorders>
                        </w:tblPr>
                        <w:tblGrid><w:gridCol w:w="1800"/><w:gridCol w:w="1800"/><w:gridCol w:w="1800"/></w:tblGrid>
                        <w:tr>
                          <w:tc><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="4472C4"/></w:tcPr><w:p><w:r><w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr><w:t>Name</w:t></w:r></w:p></w:tc>
                          <w:tc><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="4472C4"/></w:tcPr><w:p><w:r><w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr><w:t>Department</w:t></w:r></w:p></w:tc>
                          <w:tc><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="4472C4"/></w:tcPr><w:p><w:r><w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr><w:t>Score</w:t></w:r></w:p></w:tc>
                        </w:tr>
                        <w:tr>
                          <w:tc><w:p><w:r><w:t>John</w:t></w:r></w:p></w:tc>
                          <w:tc><w:p><w:r><w:t>Engineering</w:t></w:r></w:p></w:tc>
                          <w:tc><w:p><w:r><w:rPr><w:color w:val="00B050"/><w:b/></w:rPr><w:t>95</w:t></w:r></w:p></w:tc>
                        </w:tr>
                        <w:tr>
                          <w:tc><w:p><w:r><w:t>Sarah</w:t></w:r></w:p></w:tc>
                          <w:tc><w:p><w:r><w:t>Marketing</w:t></w:r></w:p></w:tc>
                          <w:tc><w:p><w:r><w:rPr><w:color w:val="FF0000"/><w:b/></w:rPr><w:t>78</w:t></w:r></w:p></w:tc>
                        </w:tr>
                      </w:tbl>
                      <w:p><w:r><w:rPr><w:i/><w:sz w:val="18"/><w:color w:val="888888"/></w:rPr><w:t>* Table nested inside a textbox</w:t></w:r></w:p>
                    </w:txbxContent>
                  </wps:txbx>
                  <wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t"/>
                </wps:wsp>
              </a:graphicData>
            </a:graphic>
          </wp:anchor>
        </w:drawing>
      </mc:Choice>
    </mc:AlternateContent>
  </w:r>
</w:p>'

echo "Done: Scenario 3: Nested Table"

# ==================== Scenario 4: Rotated Textbox (45 degrees + gradient background) ====================
officecli add "$OUT" /body --type paragraph --prop text="Scenario 4: Rotated Textbox (45 degrees)" --prop style=Heading2

officecli raw-set "$OUT" /document --xpath "//w:body" --action append --xml '
<w:p>
  <w:r>
    <w:rPr><w:noProof/></w:rPr>
    <mc:AlternateContent>
      <mc:Choice Requires="wps">
        <w:drawing>
          <wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251662336" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
            <wp:simplePos x="0" y="0"/>
            <wp:positionH relativeFrom="column"><wp:posOffset>1500000</wp:posOffset></wp:positionH>
            <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
            <wp:extent cx="2400000" cy="1200000"/>
            <wp:effectExtent l="300000" t="300000" r="300000" b="300000"/>
            <wp:wrapTopAndBottom/>
            <wp:docPr id="4" name="TextBox 4"/>
            <a:graphic>
              <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                <wps:wsp>
                  <wps:cNvSpPr txBox="1"/>
                  <wps:spPr>
                    <a:xfrm rot="2700000"><a:off x="0" y="0"/><a:ext cx="2400000" cy="1200000"/></a:xfrm>
                    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                    <a:gradFill>
                      <a:gsLst>
                        <a:gs pos="0"><a:srgbClr val="FF6B6B"/></a:gs>
                        <a:gs pos="100000"><a:srgbClr val="FFE66D"/></a:gs>
                      </a:gsLst>
                    </a:gradFill>
                    <a:ln w="19050"><a:solidFill><a:srgbClr val="C0392B"/></a:solidFill></a:ln>
                  </wps:spPr>
                  <wps:txbx>
                    <w:txbxContent>
                      <w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="28"/><w:color w:val="FFFFFF"/></w:rPr><w:t>Rotated 45</w:t></w:r></w:p>
                      <w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:color w:val="FFFFFF"/></w:rPr><w:t>Gradient + Rotation</w:t></w:r></w:p>
                    </w:txbxContent>
                  </wps:txbx>
                  <wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="ctr"/>
                </wps:wsp>
              </a:graphicData>
            </a:graphic>
          </wp:anchor>
        </w:drawing>
      </mc:Choice>
    </mc:AlternateContent>
  </w:r>
</w:p>'

echo "Done: Scenario 4: Rotated Textbox"

# ==================== Scenario 5: Vertical Text Textbox ====================
officecli add "$OUT" /body --type paragraph --prop text="Scenario 5: Vertical Text Textbox" --prop style=Heading2

officecli raw-set "$OUT" /document --xpath "//w:body" --action append --xml '
<w:p>
  <w:r>
    <w:rPr><w:noProof/></w:rPr>
    <mc:AlternateContent>
      <mc:Choice Requires="wps">
        <w:drawing>
          <wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251663360" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
            <wp:simplePos x="0" y="0"/>
            <wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>
            <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
            <wp:extent cx="800000" cy="2400000"/>
            <wp:effectExtent l="0" t="0" r="0" b="0"/>
            <wp:wrapTopAndBottom/>
            <wp:docPr id="5" name="TextBox 5"/>
            <a:graphic>
              <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                <wps:wsp>
                  <wps:cNvSpPr txBox="1"/>
                  <wps:spPr>
                    <a:xfrm><a:off x="0" y="0"/><a:ext cx="800000" cy="2400000"/></a:xfrm>
                    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                    <a:solidFill><a:srgbClr val="FFF0F5"/></a:solidFill>
                    <a:ln w="12700"><a:solidFill><a:srgbClr val="8B0000"/></a:solidFill></a:ln>
                  </wps:spPr>
                  <wps:txbx>
                    <w:txbxContent>
                      <w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="36"/><w:color w:val="8B0000"/></w:rPr><w:t>Vertical text content</w:t></w:r></w:p>
                    </w:txbxContent>
                  </wps:txbx>
                  <wps:bodyPr rot="0" vert="eaVert" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t"/>
                </wps:wsp>
              </a:graphicData>
            </a:graphic>
          </wp:anchor>
        </w:drawing>
      </mc:Choice>
    </mc:AlternateContent>
  </w:r>
</w:p>'

echo "Done: Scenario 5: Vertical Textbox"

# ==================== Scenario 6: Rounded Rectangle + Shadow ====================
officecli add "$OUT" /body --type paragraph --prop text="Scenario 6: Rounded Rectangle Textbox" --prop style=Heading2

officecli raw-set "$OUT" /document --xpath "//w:body" --action append --xml '
<w:p>
  <w:r>
    <w:rPr><w:noProof/></w:rPr>
    <mc:AlternateContent>
      <mc:Choice Requires="wps">
        <w:drawing>
          <wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251664384" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
            <wp:simplePos x="0" y="0"/>
            <wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>
            <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
            <wp:extent cx="5400000" cy="1500000"/>
            <wp:effectExtent l="0" t="0" r="0" b="0"/>
            <wp:wrapTopAndBottom/>
            <wp:docPr id="6" name="TextBox 6"/>
            <a:graphic>
              <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                <wps:wsp>
                  <wps:cNvSpPr txBox="1"/>
                  <wps:spPr>
                    <a:xfrm><a:off x="0" y="0"/><a:ext cx="5400000" cy="1500000"/></a:xfrm>
                    <a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 16667"/></a:avLst></a:prstGeom>
                    <a:solidFill><a:srgbClr val="E8F5E9"/></a:solidFill>
                    <a:ln w="28575"><a:solidFill><a:srgbClr val="2E7D32"/></a:solidFill></a:ln>
                    <a:effectLst>
                      <a:outerShdw blurRad="50800" dist="38100" dir="5400000" algn="t" rotWithShape="0">
                        <a:srgbClr val="000000"><a:alpha val="40000"/></a:srgbClr>
                      </a:outerShdw>
                    </a:effectLst>
                  </wps:spPr>
                  <wps:txbx>
                    <w:txbxContent>
                      <w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="30"/><w:color w:val="2E7D32"/></w:rPr><w:t>Rounded Rectangle + Shadow</w:t></w:r></w:p>
                      <w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>This is a rounded rectangle textbox with an outer shadow effect.</w:t></w:r></w:p>
                      <w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:i/><w:color w:val="666666"/></w:rPr><w:t>Uses prstGeom=roundRect for rounded corners</w:t></w:r></w:p>
                    </w:txbxContent>
                  </wps:txbx>
                  <wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="ctr"/>
                </wps:wsp>
              </a:graphicData>
            </a:graphic>
          </wp:anchor>
        </w:drawing>
      </mc:Choice>
    </mc:AlternateContent>
  </w:r>
</w:p>'

echo "Done: Scenario 6: Rounded Rectangle"

# ==================== Scenario 7: Side-by-side Textboxes (Card Layout) ====================
officecli add "$OUT" /body --type paragraph --prop text="Scenario 7: Side-by-side Textboxes (Card Layout)" --prop style=Heading2

officecli raw-set "$OUT" /document --xpath "//w:body" --action append --xml '
<w:p>
  <w:r>
    <w:rPr><w:noProof/></w:rPr>
    <mc:AlternateContent>
      <mc:Choice Requires="wps">
        <w:drawing>
          <wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251665408" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
            <wp:simplePos x="0" y="0"/>
            <wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>
            <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
            <wp:extent cx="1700000" cy="1400000"/>
            <wp:effectExtent l="0" t="0" r="0" b="0"/>
            <wp:wrapNone/>
            <wp:docPr id="7" name="Card1"/>
            <a:graphic>
              <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                <wps:wsp>
                  <wps:cNvSpPr txBox="1"/>
                  <wps:spPr>
                    <a:xfrm><a:off x="0" y="0"/><a:ext cx="1700000" cy="1400000"/></a:xfrm>
                    <a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>
                    <a:solidFill><a:srgbClr val="E3F2FD"/></a:solidFill>
                    <a:ln w="12700"><a:solidFill><a:srgbClr val="1565C0"/></a:solidFill></a:ln>
                  </wps:spPr>
                  <wps:txbx>
                    <w:txbxContent>
                      <w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="28"/><w:color w:val="1565C0"/></w:rPr><w:t>Card A</w:t></w:r></w:p>
                      <w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:sz w:val="48"/></w:rPr><w:t>128</w:t></w:r></w:p>
                      <w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:color w:val="888888"/><w:sz w:val="18"/></w:rPr><w:t>Daily Visits</w:t></w:r></w:p>
                    </w:txbxContent>
                  </wps:txbx>
                  <wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="ctr"/>
                </wps:wsp>
              </a:graphicData>
            </a:graphic>
          </wp:anchor>
        </w:drawing>
      </mc:Choice>
    </mc:AlternateContent>
  </w:r>
  <w:r>
    <w:rPr><w:noProof/></w:rPr>
    <mc:AlternateContent>
      <mc:Choice Requires="wps">
        <w:drawing>
          <wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251666432" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
            <wp:simplePos x="0" y="0"/>
            <wp:positionH relativeFrom="column"><wp:posOffset>1900000</wp:posOffset></wp:positionH>
            <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
            <wp:extent cx="1700000" cy="1400000"/>
            <wp:effectExtent l="0" t="0" r="0" b="0"/>
            <wp:wrapNone/>
            <wp:docPr id="8" name="Card2"/>
            <a:graphic>
              <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                <wps:wsp>
                  <wps:cNvSpPr txBox="1"/>
                  <wps:spPr>
                    <a:xfrm><a:off x="0" y="0"/><a:ext cx="1700000" cy="1400000"/></a:xfrm>
                    <a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>
                    <a:solidFill><a:srgbClr val="FFF3E0"/></a:solidFill>
                    <a:ln w="12700"><a:solidFill><a:srgbClr val="E65100"/></a:solidFill></a:ln>
                  </wps:spPr>
                  <wps:txbx>
                    <w:txbxContent>
                      <w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="28"/><w:color w:val="E65100"/></w:rPr><w:t>Card B</w:t></w:r></w:p>
                      <w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:sz w:val="48"/></w:rPr><w:t>56</w:t></w:r></w:p>
                      <w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:color w:val="888888"/><w:sz w:val="18"/></w:rPr><w:t>New Orders</w:t></w:r></w:p>
                    </w:txbxContent>
                  </wps:txbx>
                  <wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="ctr"/>
                </wps:wsp>
              </a:graphicData>
            </a:graphic>
          </wp:anchor>
        </w:drawing>
      </mc:Choice>
    </mc:AlternateContent>
  </w:r>
  <w:r>
    <w:rPr><w:noProof/></w:rPr>
    <mc:AlternateContent>
      <mc:Choice Requires="wps">
        <w:drawing>
          <wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251667456" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
            <wp:simplePos x="0" y="0"/>
            <wp:positionH relativeFrom="column"><wp:posOffset>3800000</wp:posOffset></wp:positionH>
            <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
            <wp:extent cx="1700000" cy="1400000"/>
            <wp:effectExtent l="0" t="0" r="0" b="0"/>
            <wp:wrapNone/>
            <wp:docPr id="9" name="Card3"/>
            <a:graphic>
              <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                <wps:wsp>
                  <wps:cNvSpPr txBox="1"/>
                  <wps:spPr>
                    <a:xfrm><a:off x="0" y="0"/><a:ext cx="1700000" cy="1400000"/></a:xfrm>
                    <a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>
                    <a:solidFill><a:srgbClr val="E8F5E9"/></a:solidFill>
                    <a:ln w="12700"><a:solidFill><a:srgbClr val="2E7D32"/></a:solidFill></a:ln>
                  </wps:spPr>
                  <wps:txbx>
                    <w:txbxContent>
                      <w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="28"/><w:color w:val="2E7D32"/></w:rPr><w:t>Card C</w:t></w:r></w:p>
                      <w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:sz w:val="48"/></w:rPr><w:t>99.8%</w:t></w:r></w:p>
                      <w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:color w:val="888888"/><w:sz w:val="18"/></w:rPr><w:t>Uptime</w:t></w:r></w:p>
                    </w:txbxContent>
                  </wps:txbx>
                  <wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="ctr"/>
                </wps:wsp>
              </a:graphicData>
            </a:graphic>
          </wp:anchor>
        </w:drawing>
      </mc:Choice>
    </mc:AlternateContent>
  </w:r>
</w:p>'

echo "Done: Scenario 7: Side-by-side Cards"

# ==================== Scenario 8: Borderless Transparent Textbox ====================
officecli add "$OUT" /body --type paragraph --prop text="Scenario 8: Borderless Transparent Textbox" --prop style=Heading2

officecli raw-set "$OUT" /document --xpath "//w:body" --action append --xml '
<w:p>
  <w:r>
    <w:rPr><w:noProof/></w:rPr>
    <mc:AlternateContent>
      <mc:Choice Requires="wps">
        <w:drawing>
          <wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251668480" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
            <wp:simplePos x="0" y="0"/>
            <wp:positionH relativeFrom="column"><wp:posOffset>500000</wp:posOffset></wp:positionH>
            <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
            <wp:extent cx="4000000" cy="800000"/>
            <wp:effectExtent l="0" t="0" r="0" b="0"/>
            <wp:wrapTopAndBottom/>
            <wp:docPr id="10" name="TextBox 10"/>
            <a:graphic>
              <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                <wps:wsp>
                  <wps:cNvSpPr txBox="1"/>
                  <wps:spPr>
                    <a:xfrm><a:off x="0" y="0"/><a:ext cx="4000000" cy="800000"/></a:xfrm>
                    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                    <a:noFill/>
                    <a:ln><a:noFill/></a:ln>
                  </wps:spPr>
                  <wps:txbx>
                    <w:txbxContent>
                      <w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:sz w:val="44"/><w:color w:val="AAAAAA"/><w:i/></w:rPr><w:t>Borderless transparent text</w:t></w:r></w:p>
                    </w:txbxContent>
                  </wps:txbx>
                  <wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="ctr"/>
                </wps:wsp>
              </a:graphicData>
            </a:graphic>
          </wp:anchor>
        </w:drawing>
      </mc:Choice>
    </mc:AlternateContent>
  </w:r>
</w:p>'

echo "Done: Scenario 8: Transparent Textbox"

# ==================== Scenario 9: Text Overflow Textbox ====================
officecli add "$OUT" /body --type paragraph --prop text="Scenario 9: Text Overflow Textbox" --prop style=Heading2

officecli raw-set "$OUT" /document --xpath "//w:body" --action append --xml '
<w:p>
  <w:r>
    <w:rPr><w:noProof/></w:rPr>
    <mc:AlternateContent>
      <mc:Choice Requires="wps">
        <w:drawing>
          <wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251669504" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
            <wp:simplePos x="0" y="0"/>
            <wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>
            <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
            <wp:extent cx="5400000" cy="600000"/>
            <wp:effectExtent l="0" t="0" r="0" b="0"/>
            <wp:wrapTopAndBottom/>
            <wp:docPr id="11" name="TextBox 11"/>
            <a:graphic>
              <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                <wps:wsp>
                  <wps:cNvSpPr txBox="1"/>
                  <wps:spPr>
                    <a:xfrm><a:off x="0" y="0"/><a:ext cx="5400000" cy="600000"/></a:xfrm>
                    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                    <a:solidFill><a:srgbClr val="FCE4EC"/></a:solidFill>
                    <a:ln w="12700"><a:solidFill><a:srgbClr val="C62828"/></a:solidFill></a:ln>
                  </wps:spPr>
                  <wps:txbx>
                    <w:txbxContent>
                      <w:p><w:r><w:rPr><w:b/><w:color w:val="C62828"/></w:rPr><w:t>Line 1: This is a fixed-height textbox with too much text to test overflow behavior.</w:t></w:r></w:p>
                      <w:p><w:r><w:t>Line 2: In real usage, the textbox height is limited but content can be long.</w:t></w:r></w:p>
                      <w:p><w:r><w:t>Line 3: Word usually auto-expands the textbox height, but fixed height may truncate.</w:t></w:r></w:p>
                      <w:p><w:r><w:t>Line 4: This line may be truncated or overflow, depending on bodyPr settings.</w:t></w:r></w:p>
                      <w:p><w:r><w:t>Line 5: Continuing to test more overflow content...</w:t></w:r></w:p>
                      <w:p><w:r><w:t>Line 6: Final overflow line.</w:t></w:r></w:p>
                    </w:txbxContent>
                  </wps:txbx>
                  <wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t"/>
                </wps:wsp>
              </a:graphicData>
            </a:graphic>
          </wp:anchor>
        </w:drawing>
      </mc:Choice>
    </mc:AlternateContent>
  </w:r>
</w:p>'

echo "Done: Scenario 9: Overflow Textbox"

# ==================== Scenario 10: Textbox Stacking (Z-order) ====================
officecli add "$OUT" /body --type paragraph --prop text="Scenario 10: Textbox Stacking (Z-order)" --prop style=Heading2

officecli raw-set "$OUT" /document --xpath "//w:body" --action append --xml '
<w:p>
  <w:r>
    <w:rPr><w:noProof/></w:rPr>
    <mc:AlternateContent>
      <mc:Choice Requires="wps">
        <w:drawing>
          <wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251670528" behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1">
            <wp:simplePos x="0" y="0"/>
            <wp:positionH relativeFrom="column"><wp:posOffset>200000</wp:posOffset></wp:positionH>
            <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
            <wp:extent cx="3000000" cy="1500000"/>
            <wp:effectExtent l="0" t="0" r="0" b="0"/>
            <wp:wrapNone/>
            <wp:docPr id="12" name="Bottom layer"/>
            <a:graphic>
              <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                <wps:wsp>
                  <wps:cNvSpPr txBox="1"/>
                  <wps:spPr>
                    <a:xfrm><a:off x="0" y="0"/><a:ext cx="3000000" cy="1500000"/></a:xfrm>
                    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                    <a:solidFill><a:srgbClr val="BBDEFB"/></a:solidFill>
                    <a:ln w="19050"><a:solidFill><a:srgbClr val="1565C0"/></a:solidFill></a:ln>
                  </wps:spPr>
                  <wps:txbx>
                    <w:txbxContent>
                      <w:p><w:r><w:rPr><w:b/><w:sz w:val="28"/><w:color w:val="1565C0"/></w:rPr><w:t>Bottom layer (behindDoc)</w:t></w:r></w:p>
                      <w:p><w:r><w:t>This textbox is behind the document content.</w:t></w:r></w:p>
                      <w:p><w:r><w:t>It should be partially obscured by the top layer textbox.</w:t></w:r></w:p>
                    </w:txbxContent>
                  </wps:txbx>
                  <wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t"/>
                </wps:wsp>
              </a:graphicData>
            </a:graphic>
          </wp:anchor>
        </w:drawing>
      </mc:Choice>
    </mc:AlternateContent>
  </w:r>
  <w:r>
    <w:rPr><w:noProof/></w:rPr>
    <mc:AlternateContent>
      <mc:Choice Requires="wps">
        <w:drawing>
          <wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251671552" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
            <wp:simplePos x="0" y="0"/>
            <wp:positionH relativeFrom="column"><wp:posOffset>1200000</wp:posOffset></wp:positionH>
            <wp:positionV relativeFrom="paragraph"><wp:posOffset>400000</wp:posOffset></wp:positionV>
            <wp:extent cx="3000000" cy="1200000"/>
            <wp:effectExtent l="0" t="0" r="0" b="0"/>
            <wp:wrapTopAndBottom/>
            <wp:docPr id="13" name="Top layer"/>
            <a:graphic>
              <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                <wps:wsp>
                  <wps:cNvSpPr txBox="1"/>
                  <wps:spPr>
                    <a:xfrm><a:off x="0" y="0"/><a:ext cx="3000000" cy="1200000"/></a:xfrm>
                    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                    <a:solidFill><a:srgbClr val="FFCDD2"><a:alpha val="80000"/></a:srgbClr></a:solidFill>
                    <a:ln w="19050"><a:solidFill><a:srgbClr val="C62828"/></a:solidFill></a:ln>
                  </wps:spPr>
                  <wps:txbx>
                    <w:txbxContent>
                      <w:p><w:r><w:rPr><w:b/><w:sz w:val="28"/><w:color w:val="C62828"/></w:rPr><w:t>Top layer (translucent)</w:t></w:r></w:p>
                      <w:p><w:r><w:t>This textbox is on top, 80% opacity.</w:t></w:r></w:p>
                      <w:p><w:r><w:t>It partially obscures the bottom blue textbox.</w:t></w:r></w:p>
                    </w:txbxContent>
                  </wps:txbx>
                  <wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t"/>
                </wps:wsp>
              </a:graphicData>
            </a:graphic>
          </wp:anchor>
        </w:drawing>
      </mc:Choice>
    </mc:AlternateContent>
  </w:r>
</w:p>'

echo "Done: Scenario 10: Z-order Stacking"

# ==================== Verification ====================
echo ""
echo "=========================================="
echo "Document generated: $OUT"
echo "=========================================="
officecli view "$OUT" outline
echo ""
officecli validate "$OUT"
