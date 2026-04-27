// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler
{
    private List<DocumentNode> GetSlideChildNodes(SlidePart slidePart, int slideNum, int depth)
    {
        var children = new List<DocumentNode>();
        var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree;
        if (shapeTree == null) return children;

        int shapeIdx = 0;
        foreach (var shape in shapeTree.Elements<Shape>())
        {
            children.Add(ShapeToNode(shape, slideNum, shapeIdx + 1, depth, slidePart));
            shapeIdx++;
        }

        int tblIdx = 0;
        int chartIdx = 0;
        foreach (var gf in shapeTree.Elements<GraphicFrame>())
        {
            if (gf.Descendants<Drawing.Table>().Any())
            {
                tblIdx++;
                children.Add(TableToNode(gf, slideNum, tblIdx, depth));
            }
            else if (gf.Descendants<C.ChartReference>().Any() || IsExtendedChartFrame(gf))
            {
                chartIdx++;
                children.Add(ChartToNode(gf, slidePart, slideNum, chartIdx, depth));
            }
        }

        int picIdx = 0;
        foreach (var pic in shapeTree.Elements<Picture>())
        {
            children.Add(PictureToNode(pic, slideNum, picIdx + 1, slidePart));
            picIdx++;
        }

        var contentElements = shapeTree.ChildElements
            .Where(e => e is Shape or Picture or GraphicFrame or GroupShape or ConnectionShape).ToList();

        int grpIdx = 0;
        foreach (var grp in shapeTree.Elements<GroupShape>())
        {
            grpIdx++;
            var grpName = grp.NonVisualGroupShapeProperties?.NonVisualDrawingProperties?.Name?.Value ?? "Group";
            var grpPathSeg = BuildElementPathSegment("group", grp, grpIdx);
            var grpNode = new DocumentNode
            {
                Path = $"/slide[{slideNum}]/{grpPathSeg}",
                Type = "group",
                Preview = grpName,
                ChildCount = grp.Elements<Shape>().Count() + grp.Elements<Picture>().Count()
                    + grp.Elements<GraphicFrame>().Count() + grp.Elements<ConnectionShape>().Count()
                    + grp.Elements<GroupShape>().Count()
            };
            grpNode.Format["name"] = grpName;
            var grpXfrm = grp.GroupShapeProperties?.TransformGroup;
            if (grpXfrm?.Offset?.X != null) grpNode.Format["x"] = FormatEmu(grpXfrm.Offset.X.Value);
            if (grpXfrm?.Offset?.Y != null) grpNode.Format["y"] = FormatEmu(grpXfrm.Offset.Y.Value);
            if (grpXfrm?.Extents?.Cx != null) grpNode.Format["width"] = FormatEmu(grpXfrm.Extents.Cx.Value);
            if (grpXfrm?.Extents?.Cy != null) grpNode.Format["height"] = FormatEmu(grpXfrm.Extents.Cy.Value);
            if (grpXfrm?.Rotation != null && grpXfrm.Rotation.Value != 0)
                grpNode.Format["rotation"] = $"{grpXfrm.Rotation.Value / 60000.0:0.##}";
            var grpFillColor = ReadColorFromFill(grp.GroupShapeProperties?.GetFirstChild<Drawing.SolidFill>());
            if (grpFillColor != null) grpNode.Format["fill"] = grpFillColor;
            else if (grp.GroupShapeProperties?.GetFirstChild<Drawing.NoFill>() != null) grpNode.Format["fill"] = "none";
            else if (grp.GroupShapeProperties?.GetFirstChild<Drawing.GradientFill>() != null) grpNode.Format["fill"] = "gradient";
            var grpZIdx = contentElements.IndexOf(grp);
            if (grpZIdx >= 0) grpNode.Format["zorder"] = grpZIdx + 1;
            children.Add(grpNode);
        }

        int cxnIdx = 0;
        foreach (var cxn in shapeTree.Elements<ConnectionShape>())
        {
            cxnIdx++;
            children.Add(ConnectorToNode(cxn, slideNum, cxnIdx));
        }

        var zoomElements = GetZoomElements(shapeTree);
        int zmIdx = 0;
        foreach (var zmEl in zoomElements)
        {
            zmIdx++;
            children.Add(ZoomToNode(zmEl, slideNum, zmIdx));
        }

        var model3dElements = GetModel3DElements(shapeTree);
        int m3dIdx = 0;
        foreach (var m3dEl in model3dElements)
        {
            m3dIdx++;
            children.Add(Model3DToNode(m3dEl, slideNum, m3dIdx));
        }

        return children;
    }

    private static DocumentNode TableToNode(GraphicFrame gf, int slideNum, int tblIdx, int depth)
    {
        var table = gf.Descendants<Drawing.Table>().First();
        var rows = table.Elements<Drawing.TableRow>().ToList();
        var cols = rows.FirstOrDefault()?.Elements<Drawing.TableCell>().Count() ?? 0;
        var name = gf.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Name?.Value ?? "Table";

        var tblPathSeg = BuildElementPathSegment("table", gf, tblIdx);
        var node = new DocumentNode
        {
            Path = $"/slide[{slideNum}]/{tblPathSeg}",
            Type = "table",
            Preview = $"{name} ({rows.Count}x{cols})",
            ChildCount = rows.Count
        };

        node.Format["name"] = name;
        var tblId = GetCNvPrId(gf);
        if (tblId.HasValue) node.Format["id"] = tblId.Value;
        node.Format["rows"] = rows.Count;
        node.Format["cols"] = cols;

        // Table style
        var tblPr = table.GetFirstChild<Drawing.TableProperties>();
        var tableStyleId = tblPr?.GetFirstChild<Drawing.TableStyleId>()?.InnerText;
        if (!string.IsNullOrEmpty(tableStyleId))
        {
            var styleName = TableStyleGuidToName(tableStyleId);
            node.Format["tableStyleId"] = styleName ?? tableStyleId;
            if (styleName != null)
            {
                node.Format["tableStyle"] = styleName;
                node.Format["style"] = styleName;
            }
            else
            {
                node.Format["tableStyle"] = tableStyleId;
            }
        }

        // TableLook flags
        if (tblPr != null)
        {
            if (tblPr.FirstRow is not null) node.Format["firstRow"] = tblPr.FirstRow.Value;
            if (tblPr.LastRow is not null) node.Format["lastRow"] = tblPr.LastRow.Value;
            if (tblPr.FirstColumn is not null) node.Format["firstCol"] = tblPr.FirstColumn.Value;
            if (tblPr.LastColumn is not null) node.Format["lastCol"] = tblPr.LastColumn.Value;
            if (tblPr.BandRow is not null) node.Format["bandedRows"] = tblPr.BandRow.Value;
            if (tblPr.BandColumn is not null) node.Format["bandedCols"] = tblPr.BandColumn.Value;
        }

        // Position
        var offset = gf.Transform?.Offset;
        if (offset != null)
        {
            if (offset.X is not null) node.Format["x"] = FormatEmu(offset.X!);
            if (offset.Y is not null) node.Format["y"] = FormatEmu(offset.Y!);
        }
        var extents = gf.Transform?.Extents;
        if (extents != null)
        {
            if (extents.Cx is not null) node.Format["width"] = FormatEmu(extents.Cx!);
            if (extents.Cy is not null) node.Format["height"] = FormatEmu(extents.Cy!);
        }

        if (depth > 0)
        {
            int rIdx = 0;
            foreach (var row in rows)
            {
                rIdx++;
                var rowNode = new DocumentNode
                {
                    Path = $"/slide[{slideNum}]/{tblPathSeg}/tr[{rIdx}]",
                    Type = "tr",
                    ChildCount = row.Elements<Drawing.TableCell>().Count()
                };

                // Row height
                if (row.Height?.HasValue == true)
                    rowNode.Format["height"] = FormatEmu(row.Height.Value);

                if (depth > 1)
                {
                    int cIdx = 0;
                    foreach (var cell in row.Elements<Drawing.TableCell>())
                    {
                        cIdx++;
                        var cellText = cell.TextBody?.InnerText ?? "";
                        var cellNode = new DocumentNode
                        {
                            Path = $"/slide[{slideNum}]/{tblPathSeg}/tr[{rIdx}]/tc[{cIdx}]",
                            Type = "tc",
                            Text = cellText
                        };

                        // Cell fill (blip, gradient, or solid)
                        var tcPr = cell.TableCellProperties ?? cell.GetFirstChild<Drawing.TableCellProperties>();
                        var cellBlipFill = tcPr?.GetFirstChild<Drawing.BlipFill>();
                        if (cellBlipFill != null)
                        {
                            var blipEmbed = cellBlipFill.GetFirstChild<Drawing.Blip>()?.Embed?.Value;
                            cellNode.Format["fill"] = "image";
                            if (blipEmbed != null) cellNode.Format["image.relId"] = blipEmbed;
                        }
                        else if (tcPr?.GetFirstChild<Drawing.GradientFill>() is { } gradFill)
                        {
                            // Preserve all stops (including intermediate ones) via the shared helper.
                            cellNode.Format["gradient"] = ReadGradientString(gradFill);
                            cellNode.Format["fill"] = "gradient";
                        }
                        else
                        {
                            var cellFillHex = tcPr?.GetFirstChild<Drawing.SolidFill>()?.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value;
                            if (cellFillHex != null) cellNode.Format["fill"] = ParseHelpers.FormatHexColor(cellFillHex);
                        }

                        // Cell borders (including diagonal tl2br/tr2bl)
                        if (tcPr != null) ReadTableCellBorders(tcPr, cellNode);

                        // Cell vertical alignment
                        if (tcPr?.Anchor?.HasValue == true)
                        {
                            var av = tcPr.Anchor.Value;
                            if (av == Drawing.TextAnchoringTypeValues.Top) cellNode.Format["valign"] = "top";
                            else if (av == Drawing.TextAnchoringTypeValues.Center) cellNode.Format["valign"] = "center";
                            else if (av == Drawing.TextAnchoringTypeValues.Bottom) cellNode.Format["valign"] = "bottom";
                            else cellNode.Format["valign"] = tcPr.Anchor.InnerText switch
                            {
                                "ctr" => "center",
                                _ => tcPr.Anchor.InnerText
                            };
                        }

                        // Cell run-level formatting (font, size, bold, italic, underline, strike, color)
                        var cellFirstRun = cell.Descendants<Drawing.Run>().FirstOrDefault();
                        if (cellFirstRun?.RunProperties != null)
                        {
                            var rp = cellFirstRun.RunProperties;
                            var cellFont = rp.GetFirstChild<Drawing.LatinFont>()?.Typeface?.Value
                                ?? rp.GetFirstChild<Drawing.EastAsianFont>()?.Typeface?.Value;
                            if (cellFont != null) cellNode.Format["font"] = cellFont;

                            if (rp.FontSize?.HasValue == true)
                                cellNode.Format["size"] = $"{rp.FontSize.Value / 100.0:0.##}pt";

                            if (rp.Bold?.HasValue == true) cellNode.Format["bold"] = rp.Bold.Value;
                            if (rp.Italic?.HasValue == true) cellNode.Format["italic"] = rp.Italic.Value;

                            if (rp.Underline?.HasValue == true && rp.Underline.Value != Drawing.TextUnderlineValues.None)
                            {
                                cellNode.Format["underline"] = rp.Underline.InnerText switch
                                {
                                    "sng" => "single",
                                    "dbl" => "double",
                                    _ => rp.Underline.InnerText
                                };
                            }
                            if (rp.Strike?.HasValue == true && rp.Strike.Value != Drawing.TextStrikeValues.NoStrike)
                            {
                                cellNode.Format["strike"] = rp.Strike.Value == Drawing.TextStrikeValues.DoubleStrike ? "double" : "single";
                            }

                            var cellRunColor = ReadColorFromFill(rp.GetFirstChild<Drawing.SolidFill>());
                            if (cellRunColor != null) cellNode.Format["color"] = cellRunColor;

                            if (rp.Spacing?.HasValue == true)
                                cellNode.Format["spacing"] = $"{rp.Spacing.Value / 100.0:0.##}";
                            if (rp.Baseline?.HasValue == true && rp.Baseline.Value != 0)
                                cellNode.Format["baseline"] = $"{rp.Baseline.Value / 1000.0:0.##}";
                        }

                        // Cell paragraph alignment
                        var cellFirstPara = cell.TextBody?.Elements<Drawing.Paragraph>().FirstOrDefault();
                        if (cellFirstPara?.ParagraphProperties?.Alignment?.HasValue == true)
                        {
                            var alv = cellFirstPara.ParagraphProperties.Alignment.Value;
                            var align = cellFirstPara.ParagraphProperties.Alignment.InnerText;
                            if (alv == Drawing.TextAlignmentTypeValues.Left) align = "left";
                            else if (alv == Drawing.TextAlignmentTypeValues.Center) align = "center";
                            else if (alv == Drawing.TextAlignmentTypeValues.Right) align = "right";
                            else if (alv == Drawing.TextAlignmentTypeValues.Justified) align = "justify";
                            else if (align == "ctr") align = "center";
                            cellNode.Format["align"] = align;
                        }

                        rowNode.Children.Add(cellNode);
                    }
                }
                node.Children.Add(rowNode);
            }
        }

        return node;
    }

    private static DocumentNode ShapeToNode(Shape shape, int slideNum, int shapeIdx, int depth, OpenXmlPart? part = null)
    {
        var text = GetShapeText(shape);
        var name = GetShapeName(shape);
        var isTitle = IsTitle(shape);
        var isEquation = !isTitle && shape.TextBody != null
            && shape.TextBody.Descendants().Any(e => e.LocalName == "oMath" || e.LocalName == "oMathPara"
                || (e.LocalName == "m" && e.NamespaceUri == "http://schemas.microsoft.com/office/drawing/2010/main"));

        var shapePathSeg = BuildElementPathSegment("shape", shape, shapeIdx);
        var node = new DocumentNode
        {
            Path = $"/slide[{slideNum}]/{shapePathSeg}",
            Type = isTitle ? "title" : isEquation ? "equation" : "textbox",
            Text = text,
            Preview = string.IsNullOrEmpty(text) ? name : (text.Length > 50 ? text[..50] + "..." : text)
        };

        node.Format["name"] = name;
        var shapeId = GetCNvPrId(shape);
        if (shapeId.HasValue) node.Format["id"] = shapeId.Value;
        if (isTitle) node.Format["isTitle"] = true;

        // Position and size
        var xfrm = shape.ShapeProperties?.Transform2D;
        if (xfrm != null)
        {
            if (xfrm.Offset != null)
            {
                if (xfrm.Offset.X is not null) node.Format["x"] = FormatEmu(xfrm.Offset.X!);
                if (xfrm.Offset.Y is not null) node.Format["y"] = FormatEmu(xfrm.Offset.Y!);
            }
            if (xfrm.Extents != null)
            {
                if (xfrm.Extents.Cx is not null) node.Format["width"] = FormatEmu(xfrm.Extents.Cx!);
                if (xfrm.Extents.Cy is not null) node.Format["height"] = FormatEmu(xfrm.Extents.Cy!);
            }
        }

        // Shape fill
        var shapeFill = shape.ShapeProperties?.GetFirstChild<Drawing.SolidFill>();
        var shapeFillColor = ReadColorFromFill(shapeFill);
        if (shapeFillColor != null) node.Format["fill"] = shapeFillColor;
        // Gradient fill on shape
        var shapeGradFill = shape.ShapeProperties?.GetFirstChild<Drawing.GradientFill>();
        if (shapeGradFill != null)
        {
            var stops = shapeGradFill.GradientStopList?.Elements<Drawing.GradientStop>().ToList();
            if (stops != null && stops.Count >= 2)
            {
                var gc1 = ParseHelpers.FormatHexColor(stops[0].GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value ?? "");
                var gc2 = ParseHelpers.FormatHexColor(stops[^1].GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value ?? "");
                var lin = shapeGradFill.GetFirstChild<Drawing.LinearGradientFill>();
                var deg = lin?.Angle?.Value != null ? lin.Angle.Value / 60000.0 : 0.0;

                // Gradient opacity (from first stop's alpha)
                var gradAlpha = stops[0].GetFirstChild<Drawing.RgbColorModelHex>()?.GetFirstChild<Drawing.Alpha>()?.Val?.Value
                    ?? stops[0].GetFirstChild<Drawing.SchemeColor>()?.GetFirstChild<Drawing.Alpha>()?.Val?.Value;
                if (gradAlpha.HasValue) node.Format["opacity"] = $"{gradAlpha.Value / 100000.0:0.##}";
            }
        }
        if (shape.ShapeProperties?.GetFirstChild<Drawing.NoFill>() != null) node.Format["fill"] = "none";

        // Opacity (Alpha on SolidFill color element)
        var fillColorEl = shapeFill?.GetFirstChild<Drawing.RgbColorModelHex>() as OpenXmlElement
            ?? shapeFill?.GetFirstChild<Drawing.SchemeColor>();
        var alphaVal = fillColorEl?.GetFirstChild<Drawing.Alpha>()?.Val?.Value;
        if (alphaVal.HasValue) node.Format["opacity"] = $"{alphaVal.Value / 100000.0:0.##}";

        // Shape preset/geometry
        var presetGeom = shape.ShapeProperties?.GetFirstChild<Drawing.PresetGeometry>();
        if (presetGeom?.Preset?.HasValue == true)
        {
            node.Format["preset"] = presetGeom.Preset.InnerText;
            node.Format["geometry"] = presetGeom.Preset.InnerText;
        }
        else
        {
            var custGeom = shape.ShapeProperties?.GetFirstChild<Drawing.CustomGeometry>();
            if (custGeom != null)
            {
                node.Format["preset"] = "custom";
                // Reconstruct SVG-like path string from the custom geometry path list
                var pathData = ReconstructCustomGeometryPath(custGeom);
                node.Format["geometry"] = !string.IsNullOrEmpty(pathData) ? pathData : "custom";
            }
        }

        // Gradient fill
        var gradFill = shape.ShapeProperties?.GetFirstChild<Drawing.GradientFill>();
        if (gradFill != null)
        {
            node.Format["gradient"] = ReadGradientString(gradFill);
            if (!node.Format.ContainsKey("fill"))
                node.Format["fill"] = "gradient";
        }

        // Image (blip) fill on shape
        var blipFill = shape.ShapeProperties?.GetFirstChild<Drawing.BlipFill>();
        if (blipFill != null) node.Format["image"] = "true";

        // List style (from first paragraph)
        var firstParaBullet = shape.TextBody?.Elements<Drawing.Paragraph>().FirstOrDefault()?.ParagraphProperties;
        if (firstParaBullet != null)
        {
            var charBullet = firstParaBullet.GetFirstChild<Drawing.CharacterBullet>();
            var autoBullet = firstParaBullet.GetFirstChild<Drawing.AutoNumberedBullet>();
            if (charBullet != null)
            {
                var charVal = charBullet.Char?.Value ?? "•";
                node.Format["list"] = charVal switch
                {
                    "•" or "●" or "○" => "bullet",
                    "–" or "—" or "-" => "dash",
                    "►" or "▶" or "▸" or "➤" => "arrow",
                    "✓" or "✔" => "check",
                    "★" or "☆" or "⭐" => "star",
                    _ => charVal
                };
            }
            else if (autoBullet?.Type?.HasValue == true)
            {
                var autoVal = autoBullet.Type.InnerText;
                node.Format["list"] = autoVal switch
                {
                    "arabicPeriod" or "arabicParenR" or "arabicPlain" or "arabicParenBoth" => "numbered",
                    "romanLcPeriod" or "romanLcParenR" or "romanLcParenBoth" => "romanLc",
                    "romanUcPeriod" or "romanUcParenR" or "romanUcParenBoth" => "romanUc",
                    "alphaLcPeriod" or "alphaLcParenR" or "alphaLcParenBoth" => "alphaLc",
                    "alphaUcPeriod" or "alphaUcParenR" or "alphaUcParenBoth" => "alphaUc",
                    _ => autoVal
                };
            }
        }

        // Collect font info
        var firstRun = shape.TextBody?.Descendants<Drawing.Run>().FirstOrDefault();
        if (firstRun?.RunProperties != null)
        {
            var font = firstRun.RunProperties.GetFirstChild<Drawing.LatinFont>()?.Typeface?.Value
                ?? firstRun.RunProperties.GetFirstChild<Drawing.EastAsianFont>()?.Typeface?.Value;
            if (font != null) node.Format["font"] = font;

            var fontSize = firstRun.RunProperties.FontSize?.Value;
            if (fontSize.HasValue) node.Format["size"] = $"{fontSize.Value / 100.0:0.##}pt";

            if (firstRun.RunProperties.Bold?.HasValue == true) node.Format["bold"] = firstRun.RunProperties.Bold.Value;
            if (firstRun.RunProperties.Italic?.HasValue == true) node.Format["italic"] = firstRun.RunProperties.Italic.Value;
            // CONSISTENCY(rPr-cap): mirror cap attribute readback so shape-level
            // Get matches Set's allCaps/cap input (Set writes rPr cap="all"/"small").
            if (firstRun.RunProperties.Capital?.HasValue == true && firstRun.RunProperties.Capital.Value != Drawing.TextCapsValues.None)
                node.Format["cap"] = firstRun.RunProperties.Capital.InnerText;
            if (firstRun.RunProperties.Underline?.HasValue == true && firstRun.RunProperties.Underline.Value != Drawing.TextUnderlineValues.None)
            {
                var ulInner = firstRun.RunProperties.Underline.InnerText;
                node.Format["underline"] = ulInner switch
                {
                    "sng" => "single",
                    "dbl" => "double",
                    _ => ulInner
                };
            }
            if (firstRun.RunProperties.Strike?.HasValue == true && firstRun.RunProperties.Strike.Value != Drawing.TextStrikeValues.NoStrike)
            {
                node.Format["strike"] = firstRun.RunProperties.Strike.Value == Drawing.TextStrikeValues.DoubleStrike ? "double" : "single";
            }

            // Character spacing on first run
            if (firstRun.RunProperties.Spacing?.HasValue == true)
                node.Format["spacing"] = $"{firstRun.RunProperties.Spacing.Value / 100.0:0.##}";
            // Baseline (superscript/subscript)
            if (firstRun.RunProperties.Baseline?.HasValue == true && firstRun.RunProperties.Baseline.Value != 0)
                node.Format["baseline"] = $"{firstRun.RunProperties.Baseline.Value / 1000.0:0.##}";

            // Text color (from first run) — solid or gradient
            var runColor = ReadColorFromFill(firstRun.RunProperties.GetFirstChild<Drawing.SolidFill>());
            if (runColor != null) node.Format["color"] = runColor;
            var runGradFill = firstRun.RunProperties.GetFirstChild<Drawing.GradientFill>();
            if (runGradFill != null)
                node.Format["textFill"] = ReadGradientString(runGradFill);

            // Hyperlink on first run
            if (part != null)
            {
                var linkUrl = ReadRunHyperlinkUrl(firstRun, part);
                if (linkUrl != null) node.Format["link"] = linkUrl;
            }

            // CONSISTENCY(rpr-attr-fallback / R21-fuzzer-1+2): surface long-tail
            // rPr attributes (lang, kern, kumimoji, normalizeH, ...) at shape
            // level too, mirroring BuildRunNode. Without this, shape-level Add
            // can write `lang` to first-run rPr but shape-level Get cannot
            // surface it unless the user descends to /shape[N]/r[1] explicitly.
            FillUnknownRunProps(firstRun.RunProperties, node);
        }

        // Shape-level hyperlink (on NonVisualDrawingProperties)
        if (part != null && !node.Format.ContainsKey("link"))
        {
            var nvDp = shape.NonVisualShapeProperties?.NonVisualDrawingProperties;
            var hlClick = nvDp?.GetFirstChild<Drawing.HyperlinkOnClick>();
            var hlId = hlClick?.Id?.Value;
            if (hlId != null)
            {
                try
                {
                    var rel = part.HyperlinkRelationships.FirstOrDefault(r => r.Id == hlId);
                    if (rel?.Uri != null) node.Format["link"] = rel.Uri.ToString();
                }
                catch { }
            }
        }

        // Line/border
        var outline = shape.ShapeProperties?.GetFirstChild<Drawing.Outline>();
        if (outline != null)
        {
            var lineSolidFill = outline.GetFirstChild<Drawing.SolidFill>();
            var lineColor = ReadColorFromFill(lineSolidFill);
            if (lineColor != null) node.Format["line"] = lineColor;
            var lineIsNone = outline.GetFirstChild<Drawing.NoFill>() != null;
            if (lineIsNone) node.Format["line"] = "none";
            // When line=none, suppress the residual width readback so users don't
            // see a stale lineWidth from a prior color-set assignment.
            if (!lineIsNone && outline.Width?.HasValue == true) node.Format["lineWidth"] = FormatLineWidth(outline.Width.Value);
            var dash = outline.GetFirstChild<Drawing.PresetDash>();
            if (dash?.Val?.HasValue == true)
            {
                var dashValue = dash.Val.InnerText ?? "";
                node.Format["lineDash"] = dashValue switch
                {
                    "solid" => "solid",
                    "dot" => "dot",
                    "dash" => "dash",
                    "dashDot" => "dashdot",
                    "lgDash" => "longdash",
                    "lgDashDot" => "longdashdot",
                    "sysDot" => "sysdot",
                    "sysDash" => "sysdash",
                    "sysDashDot" => "sysdashdot",
                    "sysDashDotDot" => "sysdashdotdot",
                    _ => dashValue.ToLowerInvariant()
                };
            }
            var lineColorEl = lineSolidFill?.GetFirstChild<Drawing.RgbColorModelHex>() as OpenXmlElement
                ?? lineSolidFill?.GetFirstChild<Drawing.SchemeColor>();
            var lineAlpha = lineColorEl?.GetFirstChild<Drawing.Alpha>()?.Val?.Value;
            if (lineAlpha.HasValue) node.Format["lineOpacity"] = $"{lineAlpha.Value / 100000.0:0.##}";
        }

        // Effects (shadow, glow, reflection) — check shape-level first, then text run-level
        var effectList = shape.ShapeProperties?.GetFirstChild<Drawing.EffectList>();
        // Fall back to first text run's effectLst (used for fill=none shapes)
        var textEffectList = effectList == null || (!effectList.HasChildren)
            ? shape.TextBody?.Descendants<Drawing.RunProperties>()
                .Select(rp => rp.GetFirstChild<Drawing.EffectList>())
                .FirstOrDefault(el => el != null)
            : null;
        var activeEffectList = effectList?.HasChildren == true ? effectList : textEffectList;
        if (activeEffectList != null)
        {
            var outerShadow = activeEffectList.GetFirstChild<Drawing.OuterShadow>();
            if (outerShadow != null)
            {
                var shadowColor = ReadColorFromElement(outerShadow) ?? "000000";
                var blurPt = outerShadow.BlurRadius?.HasValue == true ? $"{outerShadow.BlurRadius.Value / 12700.0:0.##}" : "4";
                var angleDeg = outerShadow.Direction?.HasValue == true ? $"{outerShadow.Direction.Value / 60000.0:0.##}" : "45";
                var distPt = outerShadow.Distance?.HasValue == true ? $"{outerShadow.Distance.Value / 12700.0:0.##}" : "3";
                var alphaEl = outerShadow.Descendants<Drawing.Alpha>().FirstOrDefault();
                var opacity = alphaEl?.Val?.HasValue == true ? $"{alphaEl.Val.Value / 1000.0:0.##}" : "40";
                node.Format["shadow"] = $"{shadowColor}-{blurPt}-{angleDeg}-{distPt}-{opacity}";
            }
            var glow = activeEffectList.GetFirstChild<Drawing.Glow>();
            if (glow != null)
            {
                var glowColor = ReadColorFromElement(glow) ?? "000000";
                var radiusPt = glow.Radius?.HasValue == true ? $"{glow.Radius.Value / 12700.0:0.##}" : "8";
                var glowAlpha = glow.Descendants<Drawing.Alpha>().FirstOrDefault();
                var glowOpacity = glowAlpha?.Val?.HasValue == true ? $"{glowAlpha.Val.Value / 1000.0:0.##}" : "75";
                node.Format["glow"] = $"{glowColor}-{radiusPt}-{glowOpacity}";
            }
            var reflEl = activeEffectList.GetFirstChild<Drawing.Reflection>();
            if (reflEl != null)
            {
                // Map endPosition back to type: tight=55000, half=90000, full=100000
                var endPos = reflEl.EndPosition?.Value ?? 0;
                if (endPos >= 95000) node.Format["reflection"] = "full";
                else if (endPos >= 70000) node.Format["reflection"] = "half";
                else node.Format["reflection"] = "tight";
            }
            var softEdge = activeEffectList.GetFirstChild<Drawing.SoftEdge>();
            if (softEdge?.Radius?.HasValue == true)
                node.Format["softEdge"] = $"{softEdge.Radius.Value / 12700.0:0.##}";
        }

        // 3D rotation (scene3d)
        var scene3d = shape.ShapeProperties?.GetFirstChild<Drawing.Scene3DType>();
        if (scene3d != null)
        {
            var cam = scene3d.Camera;
            var rot3d = cam?.Rotation;
            if (rot3d != null)
            {
                var rx = rot3d.Latitude?.Value ?? 0;
                var ry = rot3d.Longitude?.Value ?? 0;
                var rz = rot3d.Revolution?.Value ?? 0;
                if (rx != 0 || ry != 0 || rz != 0)
                    node.Format["rot3d"] = $"{rx / 60000.0:0.##},{ry / 60000.0:0.##},{rz / 60000.0:0.##}";
            }
            var lightRig = scene3d.LightRig;
            if (lightRig?.Rig?.HasValue == true) node.Format["lighting"] = lightRig.Rig.InnerText;
        }

        // 3D format (sp3d)
        var sp3d = shape.ShapeProperties?.GetFirstChild<Drawing.Shape3DType>();
        if (sp3d != null)
        {
            if (sp3d.ExtrusionHeight?.HasValue == true && sp3d.ExtrusionHeight.Value != 0)
                node.Format["depth"] = $"{sp3d.ExtrusionHeight.Value / 12700.0:0.##}";
            if (sp3d.PresetMaterial?.HasValue == true)
                node.Format["material"] = sp3d.PresetMaterial.InnerText;
            var bevelT = sp3d.BevelTop;
            if (bevelT != null) node.Format["bevel"] = FormatBevel(bevelT);
            var bevelB = sp3d.BevelBottom;
            if (bevelB != null) node.Format["bevelBottom"] = FormatBevel(bevelB);
        }

        // Flip
        if (xfrm?.HorizontalFlip?.Value == true) node.Format["flipH"] = true;
        if (xfrm?.VerticalFlip?.Value == true) node.Format["flipV"] = true;

        // Z-order (1-based position among content elements: 1 = back, N = front)
        if (shape.Parent is ShapeTree zTree)
        {
            var contentEls = zTree.ChildElements
                .Where(e => e is Shape or Picture or GraphicFrame or GroupShape or ConnectionShape)
                .ToList();
            var zIdx = contentEls.IndexOf(shape);
            if (zIdx >= 0) node.Format["zorder"] = zIdx + 1;
        }

        // Rotation (plain number in degrees, no suffix, so Set can consume the value directly)
        if (xfrm?.Rotation != null && xfrm.Rotation.Value != 0)
            node.Format["rotation"] = $"{xfrm.Rotation.Value / 60000.0:0.##}";

        // Text margin
        var bodyPr = shape.TextBody?.Elements<Drawing.BodyProperties>().FirstOrDefault();
        if (bodyPr != null)
        {
            var lIns = bodyPr.LeftInset;
            var tIns = bodyPr.TopInset;
            var rIns = bodyPr.RightInset;
            var bIns = bodyPr.BottomInset;
            if (lIns != null || tIns != null || rIns != null || bIns != null)
            {
                // If all four are the same, show as single value
                if (lIns == tIns && tIns == rIns && rIns == bIns && lIns != null)
                    node.Format["margin"] = FormatEmu(lIns.Value);
                else
                    node.Format["margin"] = $"{FormatEmu(lIns ?? 91440)},{FormatEmu(tIns ?? 45720)},{FormatEmu(rIns ?? 91440)},{FormatEmu(bIns ?? 45720)}";
            }

            // Vertical alignment — map XML enum to user-friendly name (like POI TextAlign)
            if (bodyPr.Anchor?.HasValue == true)
            {
                var vaInner = bodyPr.Anchor.InnerText;
                node.Format["valign"] = vaInner switch
                {
                    "t" => "top",
                    "ctr" => "center",
                    "b" => "bottom",
                    _ => vaInner
                };
            }

            // TextWarp (WordArt)
            var prstTxWarp = bodyPr.GetFirstChild<Drawing.PresetTextWarp>();
            if (prstTxWarp?.Preset?.HasValue == true)
                node.Format["textWarp"] = prstTxWarp.Preset.InnerText;

            // AutoFit
            if (bodyPr.GetFirstChild<Drawing.NormalAutoFit>() != null) node.Format["autoFit"] = "normal";
            else if (bodyPr.GetFirstChild<Drawing.ShapeAutoFit>() != null) node.Format["autoFit"] = "shape";
            else node.Format["autoFit"] = "none";
        }

        // Text alignment (from first paragraph)
        var firstPara = shape.TextBody?.Elements<Drawing.Paragraph>().FirstOrDefault();
        if (firstPara?.ParagraphProperties?.Alignment?.HasValue == true)
        {
            var alInner = firstPara.ParagraphProperties.Alignment.InnerText;
            node.Format["align"] = alInner switch
            {
                "l" => "left",
                "ctr" => "center",
                "r" => "right",
                "just" => "justify",
                _ => alInner
            };
        }
        else if (shape.TextBody != null)
        {
            node.Format["align"] = "left";
        }

        // Paragraph spacing and indent (from first paragraph)
        var pProps = firstPara?.ParagraphProperties;
        if (pProps != null)
        {
            var lsPct = pProps.GetFirstChild<Drawing.LineSpacing>()?.GetFirstChild<Drawing.SpacingPercent>()?.Val?.Value;
            if (lsPct.HasValue) node.Format["lineSpacing"] = SpacingConverter.FormatPptLineSpacingPercent(lsPct.Value);
            var lsPts = pProps.GetFirstChild<Drawing.LineSpacing>()?.GetFirstChild<Drawing.SpacingPoints>()?.Val?.Value;
            if (lsPts.HasValue) node.Format["lineSpacing"] = SpacingConverter.FormatPptLineSpacingPoints(lsPts.Value);
            var sb = pProps.GetFirstChild<Drawing.SpaceBefore>()?.GetFirstChild<Drawing.SpacingPoints>()?.Val?.Value;
            if (sb.HasValue) node.Format["spaceBefore"] = SpacingConverter.FormatPptSpacing(sb.Value);
            var sa = pProps.GetFirstChild<Drawing.SpaceAfter>()?.GetFirstChild<Drawing.SpacingPoints>()?.Val?.Value;
            if (sa.HasValue) node.Format["spaceAfter"] = SpacingConverter.FormatPptSpacing(sa.Value);
            if (pProps.Indent?.HasValue == true) node.Format["indent"] = FormatEmu(pProps.Indent.Value);
            if (pProps.LeftMargin?.HasValue == true) node.Format["marginLeft"] = FormatEmu(pProps.LeftMargin.Value);
            if (pProps.RightMargin?.HasValue == true) node.Format["marginRight"] = FormatEmu(pProps.RightMargin.Value);
        }

        // Count paragraphs regardless of depth
        if (shape.TextBody != null)
        {
            var paragraphs = shape.TextBody.Elements<Drawing.Paragraph>().ToList();
            node.ChildCount = paragraphs.Count;

            // Include paragraph and run hierarchy at depth > 0
            if (depth > 0)
            {
                int paraIdx = 0;
                foreach (var para in paragraphs)
                {
                    var paraText = string.Join("", para.Elements<Drawing.Run>()
                        .Select(r => r.Text?.Text ?? ""));
                    var paraRuns = para.Elements<Drawing.Run>().ToList();

                    var paraNode = new DocumentNode
                    {
                        Path = $"/slide[{slideNum}]/{shapePathSeg}/paragraph[{paraIdx + 1}]",
                        Type = "paragraph",
                        Text = paraText,
                        ChildCount = paraRuns.Count
                    };

                    // Add paragraph formatting info
                    var paraPProps = para.ParagraphProperties;
                    if (paraPProps?.Alignment?.HasValue == true)
                    {
                        var paraAlignVal = paraPProps.Alignment.Value;
                        paraNode.Format["align"] = paraAlignVal == Drawing.TextAlignmentTypeValues.Center ? "center"
                            : paraAlignVal == Drawing.TextAlignmentTypeValues.Right ? "right"
                            : paraAlignVal == Drawing.TextAlignmentTypeValues.Justified ? "justify"
                            : "left";
                    }
                    if (paraPProps?.Level?.HasValue == true) paraNode.Format["level"] = paraPProps.Level.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
                    if (paraPProps?.Indent?.HasValue == true) paraNode.Format["indent"] = FormatEmu(paraPProps.Indent.Value);
                    if (paraPProps?.LeftMargin?.HasValue == true) paraNode.Format["marginLeft"] = FormatEmu(paraPProps.LeftMargin.Value);
                    if (paraPProps?.RightMargin?.HasValue == true) paraNode.Format["marginRight"] = FormatEmu(paraPProps.RightMargin.Value);
                    var pLsPct = paraPProps?.GetFirstChild<Drawing.LineSpacing>()?.GetFirstChild<Drawing.SpacingPercent>()?.Val?.Value;
                    if (pLsPct.HasValue) paraNode.Format["lineSpacing"] = SpacingConverter.FormatPptLineSpacingPercent(pLsPct.Value);
                    var pLsPts = paraPProps?.GetFirstChild<Drawing.LineSpacing>()?.GetFirstChild<Drawing.SpacingPoints>()?.Val?.Value;
                    if (pLsPts.HasValue) paraNode.Format["lineSpacing"] = SpacingConverter.FormatPptLineSpacingPoints(pLsPts.Value);
                    var pSb = paraPProps?.GetFirstChild<Drawing.SpaceBefore>()?.GetFirstChild<Drawing.SpacingPoints>()?.Val?.Value;
                    if (pSb.HasValue) paraNode.Format["spaceBefore"] = SpacingConverter.FormatPptSpacing(pSb.Value);
                    var pSa = paraPProps?.GetFirstChild<Drawing.SpaceAfter>()?.GetFirstChild<Drawing.SpacingPoints>()?.Val?.Value;
                    if (pSa.HasValue) paraNode.Format["spaceAfter"] = SpacingConverter.FormatPptSpacing(pSa.Value);

                    // Include runs at depth > 1
                    if (depth > 1)
                    {
                        int runIdx = 0;
                        foreach (var run in paraRuns)
                        {
                            paraNode.Children.Add(RunToNode(run,
                                $"/slide[{slideNum}]/{shapePathSeg}/paragraph[{paraIdx + 1}]/run[{runIdx + 1}]", part));
                            runIdx++;
                        }
                    }

                    node.Children.Add(paraNode);
                    paraIdx++;
                }
            }
        }

        // Animation (requires SlidePart to access Timing tree)
        if (part is SlidePart animSlidePart)
            ReadShapeAnimation(animSlidePart, shape, node);

        // Populate effective.* properties from slide layout/master inheritance
        PopulateEffectiveShapeProperties(node, shape, part);

        return node;
    }

    private static DocumentNode RunToNode(Drawing.Run run, string path, OpenXmlPart? part = null)
    {
        var node = new DocumentNode
        {
            Path = path,
            Type = "run",
            Text = run.Text?.Text ?? ""
        };

        if (run.RunProperties != null)
        {
            var f = run.RunProperties.GetFirstChild<Drawing.LatinFont>()?.Typeface?.Value
                ?? run.RunProperties.GetFirstChild<Drawing.EastAsianFont>()?.Typeface?.Value;
            if (f != null) node.Format["font"] = f;
            var fs = run.RunProperties.FontSize?.Value;
            if (fs.HasValue) node.Format["size"] = $"{fs.Value / 100.0:0.##}pt";
            if (run.RunProperties.Bold?.Value == true) node.Format["bold"] = true;
            if (run.RunProperties.Italic?.Value == true) node.Format["italic"] = true;
            if (run.RunProperties.Underline?.HasValue == true && run.RunProperties.Underline.Value != Drawing.TextUnderlineValues.None)
            {
                node.Format["underline"] = run.RunProperties.Underline.InnerText switch
                {
                    "sng" => "single",
                    "dbl" => "double",
                    _ => run.RunProperties.Underline.InnerText
                };
            }
            if (run.RunProperties.Strike?.HasValue == true && run.RunProperties.Strike.Value != Drawing.TextStrikeValues.NoStrike)
            {
                node.Format["strike"] = run.RunProperties.Strike.Value == Drawing.TextStrikeValues.DoubleStrike ? "double" : "single";
            }
            if (run.RunProperties.Spacing?.HasValue == true)
                node.Format["spacing"] = $"{run.RunProperties.Spacing.Value / 100.0:0.##}";
            if (run.RunProperties.Baseline?.HasValue == true && run.RunProperties.Baseline.Value != 0)
                node.Format["baseline"] = $"{run.RunProperties.Baseline.Value / 1000.0:0.##}";
            // Color (solid or gradient)
            var runFillColor = ReadColorFromFill(run.RunProperties.GetFirstChild<Drawing.SolidFill>());
            if (runFillColor != null) node.Format["color"] = runFillColor;
            var runGrad = run.RunProperties.GetFirstChild<Drawing.GradientFill>();
            if (runGrad != null) node.Format["textFill"] = ReadGradientString(runGrad);
            // Hyperlink
            if (part != null)
            {
                var linkUrl = ReadRunHyperlinkUrl(run, part);
                if (linkUrl != null) node.Format["link"] = linkUrl;
            }

            // Long-tail OOXML fallback. drawingML rPr carries most properties
            // as attributes on rPr itself (kern, spc, lang, dirty, smtClean,
            // normalizeH, baseline, ...), with sub-elements for fills/fonts/
            // hyperlinks. Symmetric with the run-context Set fallback in
            // SetRunOrShapeProperties.
            FillUnknownRunProps(run.RunProperties, node);
        }

        // Populate effective.* properties from slide layout/master inheritance
        PopulateEffectiveRunProperties(node, run, part);

        return node;
    }

    // OOXML attribute names already mapped to canonical Format keys by the
    // curated run reader. Skip these in the long-tail fallback so we don't
    // emit `b: "1"` alongside `bold: true`, `sz: "2400"` alongside `size: "24pt"`.
    private static readonly System.Collections.Generic.HashSet<string> CuratedRunAttrs =
        new(System.StringComparer.Ordinal)
    {
        "b", "i", "u", "strike", "sz", "spc", "baseline",
    };

    private static readonly System.Collections.Generic.HashSet<string> CuratedRunChildren =
        new(System.StringComparer.Ordinal)
    {
        "latin", "ea", "cs", "solidFill", "gradFill", "hlinkClick",
    };

    private static void FillUnknownRunProps(Drawing.RunProperties? rPr, DocumentNode node)
    {
        if (rPr == null) return;

        // Walk attributes on rPr itself.
        foreach (var attr in rPr.GetAttributes())
        {
            var name = attr.LocalName;
            if (string.IsNullOrEmpty(name)) continue;
            if (CuratedRunAttrs.Contains(name)) continue;
            if (node.Format.ContainsKey(name)) continue;
            node.Format[name] = attr.Value;
        }

        // Walk leaf children that match the OOXML "child-with-val" or "toggle"
        // pattern symmetric with TryCreateTypedChild's accepted shapes.
        foreach (var child in rPr.ChildElements)
        {
            var name = child.LocalName;
            if (string.IsNullOrEmpty(name)) continue;
            if (CuratedRunChildren.Contains(name)) continue;
            if (node.Format.ContainsKey(name)) continue;
            if (child.ChildElements.Count > 0) continue;

            string? valAttr = null;
            int attrCount = 0;
            foreach (var a in child.GetAttributes())
            {
                attrCount++;
                if (a.LocalName.Equals("val", System.StringComparison.OrdinalIgnoreCase))
                    valAttr = a.Value;
            }
            if (valAttr != null) node.Format[name] = valAttr;
            else if (attrCount == 0) node.Format[name] = true;
        }
    }

    private static DocumentNode PictureToNode(Picture pic, int slideNum, int picIdx, SlidePart? slidePart = null)
    {
        var name = pic.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value ?? "Picture";
        var alt = pic.NonVisualPictureProperties?.NonVisualDrawingProperties?.Description?.Value;

        // Detect video/audio
        var nvPr = pic.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties;
        var isVideo = nvPr?.GetFirstChild<Drawing.VideoFromFile>() != null;
        var isAudio = nvPr?.GetFirstChild<Drawing.AudioFromFile>() != null;
        var mediaType = isVideo ? "video" : isAudio ? "audio" : "picture";

        var picPathSeg = BuildElementPathSegment("picture", pic, picIdx);
        var node = new DocumentNode
        {
            Path = $"/slide[{slideNum}]/{picPathSeg}",
            Type = mediaType,
            Preview = name
        };

        node.Format["name"] = name;
        var picId = GetCNvPrId(pic);
        if (picId.HasValue) node.Format["id"] = picId.Value;
        if (!isVideo && !isAudio)
        {
            if (!string.IsNullOrEmpty(alt)) node.Format["alt"] = alt;
            else node.Format["alt"] = "(missing)";
        }

        // Read media timing (volume, autoplay) from slide Timing tree
        if ((isVideo || isAudio) && slidePart != null)
        {
            var shapeId = pic.NonVisualPictureProperties?.NonVisualDrawingProperties?.Id?.Value;
            if (shapeId != null)
                ReadMediaTimingProperties(slidePart, shapeId.Value, node);

            // p14:trim
            var p14Media = nvPr?.Descendants<DocumentFormat.OpenXml.Office2010.PowerPoint.Media>().FirstOrDefault();
            var trim = p14Media?.MediaTrim;
            if (trim != null)
            {
                if (trim.Start?.Value != null) node.Format["trimStart"] = trim.Start.Value;
                if (trim.End?.Value != null) node.Format["trimEnd"] = trim.End.Value;
            }
        }

        // Position and size
        var picXfrm = pic.ShapeProperties?.Transform2D;
        if (picXfrm?.Offset != null)
        {
            if (picXfrm.Offset.X is not null) node.Format["x"] = FormatEmu(picXfrm.Offset.X!);
            if (picXfrm.Offset.Y is not null) node.Format["y"] = FormatEmu(picXfrm.Offset.Y!);
        }
        if (picXfrm?.Extents != null)
        {
            if (picXfrm.Extents.Cx is not null) node.Format["width"] = FormatEmu(picXfrm.Extents.Cx!);
            if (picXfrm.Extents.Cy is not null) node.Format["height"] = FormatEmu(picXfrm.Extents.Cy!);
        }
        if (picXfrm?.Rotation != null && picXfrm.Rotation.Value != 0)
            node.Format["rotation"] = $"{picXfrm.Rotation.Value / 60000.0:0.##}";

        // Opacity (via AlphaModulateFixedEffect on blip)
        var picBlip = pic.BlipFill?.GetFirstChild<Drawing.Blip>();
        var alphaModFix = picBlip?.GetFirstChild<Drawing.AlphaModulationFixed>();
        if (alphaModFix?.Amount?.HasValue == true)
            node.Format["opacity"] = $"{alphaModFix.Amount.Value / 100000.0:0.##}";

        // Crop
        var srcRect = pic.BlipFill?.GetFirstChild<Drawing.SourceRectangle>();
        if (srcRect != null)
        {
            var cl = srcRect.Left?.Value ?? 0;
            var ct = srcRect.Top?.Value ?? 0;
            var cr = srcRect.Right?.Value ?? 0;
            var cb = srcRect.Bottom?.Value ?? 0;
            if (cl != 0 || ct != 0 || cr != 0 || cb != 0)
                node.Format["crop"] = $"{cl / 1000.0:0.##},{ct / 1000.0:0.##},{cr / 1000.0:0.##},{cb / 1000.0:0.##}";
        }

        return node;
    }

    /// <summary>
    /// Read volume and autoplay from the slide timing tree for a media shape.
    /// </summary>
    private static void ReadMediaTimingProperties(SlidePart slidePart, uint shapeId, DocumentNode node)
    {
        var timing = slidePart.Slide?.GetFirstChild<Timing>();
        if (timing == null) return;

        var shapeIdStr = shapeId.ToString();

        // Read volume from p:video/p:audio → cMediaNode
        foreach (var mediaNode in timing.Descendants<CommonMediaNode>())
        {
            var target = mediaNode.TargetElement?.GetFirstChild<ShapeTarget>();
            if (target?.ShapeId?.Value != shapeIdStr) continue;

            if (mediaNode.Volume?.HasValue == true)
                node.Format["volume"] = (int)(mediaNode.Volume.Value / 1000.0);
            break;
        }

        // Read autoplay from main sequence: look for cmd="playFrom(0)" targeting this shape
        // with nodeType="afterEffect" (autoplay) vs "clickEffect" (click-to-play)
        foreach (var cmd in timing.Descendants<Command>())
        {
            if (cmd.CommandName?.Value != "playFrom(0)") continue;
            var cmdTarget = cmd.CommonBehavior?.TargetElement?.GetFirstChild<ShapeTarget>();
            if (cmdTarget?.ShapeId?.Value != shapeIdStr) continue;

            // Found the playback command — check its parent cTn for nodeType
            var parentCTn = cmd.Parent as CommonTimeNode
                ?? cmd.Ancestors<CommonTimeNode>().FirstOrDefault();
            if (parentCTn?.NodeType?.Value == TimeNodeValues.AfterEffect)
                node.Format["autoplay"] = true;
            break;
        }
    }

    private static Shape CreateTextShape(uint id, string name, string text, bool isTitle)
    {
        var shape = new Shape();
        var appNvPr = new ApplicationNonVisualDrawingProperties();
        if (isTitle)
            appNvPr.AppendChild(new PlaceholderShape { Type = PlaceholderValues.Title });
        shape.NonVisualShapeProperties = new NonVisualShapeProperties(
            new NonVisualDrawingProperties { Id = id, Name = name },
            new NonVisualShapeDrawingProperties(),
            appNvPr
        );
        var spPr = new ShapeProperties();
        if (isTitle)
        {
            // Default title position: top-center area of standard 16:9 slide
            spPr.Transform2D = new Drawing.Transform2D
            {
                Offset = new Drawing.Offset { X = 838200, Y = 365125 },    // ~2.33cm, ~1.01cm
                Extents = new Drawing.Extents { Cx = 10515600, Cy = 1325563 } // ~29.21cm, ~3.68cm
            };
        }
        else
        {
            // Default body/content position: below title
            spPr.Transform2D = new Drawing.Transform2D
            {
                Offset = new Drawing.Offset { X = 838200, Y = 1825625 },   // ~2.33cm, ~5.07cm
                Extents = new Drawing.Extents { Cx = 10515600, Cy = 4351338 } // ~29.21cm, ~12.09cm
            };
        }
        shape.ShapeProperties = spPr;
        var body = new TextBody(
            new Drawing.BodyProperties(),
            new Drawing.ListStyle()
        );
        var lines = text.Replace("\\n", "\n").Split('\n');
        foreach (var line in lines)
        {
            body.AppendChild(new Drawing.Paragraph(
                new Drawing.Run(
                    new Drawing.RunProperties { Language = "en-US" },
                    new Drawing.Text { Text = line }
                )
            ));
        }
        shape.TextBody = body;
        return shape;
    }

    private static DocumentNode ConnectorToNode(ConnectionShape cxn, int slideNum, int cxnIdx)
    {
        var name = cxn.NonVisualConnectionShapeProperties?.NonVisualDrawingProperties?.Name?.Value ?? "Connector";
        var cxnPathSeg = BuildElementPathSegment("connector", cxn, cxnIdx);
        var node = new DocumentNode
        {
            Path = $"/slide[{slideNum}]/{cxnPathSeg}",
            Type = "connector",
            Preview = name
        };
        node.Format["name"] = name;
        var cxnId = GetCNvPrId(cxn);
        if (cxnId.HasValue) node.Format["id"] = cxnId.Value;

        var spPr = cxn.ShapeProperties;
        var xfrm = spPr?.GetFirstChild<Drawing.Transform2D>();
        if (xfrm != null)
        {
            if (xfrm.Offset?.X != null) node.Format["x"] = FormatEmu(xfrm.Offset.X!);
            if (xfrm.Offset?.Y != null) node.Format["y"] = FormatEmu(xfrm.Offset.Y!);
            if (xfrm.Extents?.Cx != null) node.Format["width"] = FormatEmu(xfrm.Extents.Cx!);
            if (xfrm.Extents?.Cy != null) node.Format["height"] = FormatEmu(xfrm.Extents.Cy!);
        }

        // Fill (solid fill on the connector shape itself, not on the outline)
        var cxnFill = ReadColorFromFill(spPr?.GetFirstChild<Drawing.SolidFill>());
        if (cxnFill != null) node.Format["fill"] = cxnFill;
        if (spPr?.GetFirstChild<Drawing.NoFill>() != null) node.Format["fill"] = "none";

        var geom = spPr?.GetFirstChild<Drawing.PresetGeometry>();
        if (geom?.Preset?.HasValue == true)
            // CONSISTENCY(canonical-key): canonical 'shape'; 'preset' was legacy key.
            node.Format["shape"] = geom.Preset.InnerText;

        var ln = spPr?.GetFirstChild<Drawing.Outline>();
        var lnIsNone = ln?.GetFirstChild<Drawing.NoFill>() != null;
        if (!lnIsNone && ln?.Width?.HasValue == true)
            node.Format["lineWidth"] = FormatLineWidth(ln.Width.Value);
        var cxnDash = ln?.GetFirstChild<Drawing.PresetDash>();
        if (cxnDash?.Val?.HasValue == true)
        {
            var dashValue = cxnDash.Val.InnerText ?? "";
            node.Format["lineDash"] = dashValue switch
            {
                "solid" => "solid",
                "dot" => "dot",
                "dash" => "dash",
                "dashDot" => "dashdot",
                "lgDash" => "longdash",
                "lgDashDot" => "longdashdot",
                "sysDot" => "sysdot",
                "sysDash" => "sysdash",
                _ => dashValue.ToLowerInvariant()
            };
        }
        var solidFill = ln?.GetFirstChild<Drawing.SolidFill>();
        var rgb = solidFill?.GetFirstChild<Drawing.RgbColorModelHex>();
        if (rgb?.Val?.HasValue == true)
            // CONSISTENCY(canonical-key): canonical 'color'; 'lineColor' was legacy key.
            node.Format["color"] = ParseHelpers.FormatHexColor(rgb.Val.Value!);

        // Line opacity
        var cxnColorEl = rgb as OpenXmlElement ?? solidFill?.GetFirstChild<Drawing.SchemeColor>();
        var cxnAlpha = cxnColorEl?.GetFirstChild<Drawing.Alpha>()?.Val?.Value;
        if (cxnAlpha.HasValue) node.Format["lineOpacity"] = $"{cxnAlpha.Value / 100000.0:0.##}";

        // Head/tail end arrows
        var headEnd = ln?.GetFirstChild<Drawing.HeadEnd>();
        if (headEnd?.Type?.HasValue == true)
            node.Format["headEnd"] = headEnd.Type.InnerText;
        var tailEnd = ln?.GetFirstChild<Drawing.TailEnd>();
        if (tailEnd?.Type?.HasValue == true)
            node.Format["tailEnd"] = tailEnd.Type.InnerText;

        // Rotation
        if (xfrm?.Rotation?.HasValue == true && xfrm.Rotation.Value != 0)
            node.Format["rotation"] = $"{xfrm.Rotation.Value / 60000.0:0.##}";

        // Connection info (startShape/endShape)
        var cxnDrawProps = cxn.NonVisualConnectionShapeProperties?.NonVisualConnectorShapeDrawingProperties;
        var startCxn = cxnDrawProps?.StartConnection;
        if (startCxn?.Id?.HasValue == true)
        {
            node.Format["startShape"] = startCxn.Id.Value;
            if (startCxn.Index?.HasValue == true)
                node.Format["startIdx"] = startCxn.Index.Value;
        }
        var endCxn = cxnDrawProps?.EndConnection;
        if (endCxn?.Id?.HasValue == true)
        {
            node.Format["endShape"] = endCxn.Id.Value;
            if (endCxn.Index?.HasValue == true)
                node.Format["endIdx"] = endCxn.Index.Value;
        }

        return node;
    }

    /// <summary>
    /// Reconstruct an SVG-like path string from a CustomGeometry element's path list.
    /// </summary>
    private static string ReconstructCustomGeometryPath(Drawing.CustomGeometry custGeom)
    {
        var sb = new StringBuilder();
        var pathList = custGeom.GetFirstChild<Drawing.PathList>();
        if (pathList == null) return "custom";

        foreach (var path in pathList.Elements<Drawing.Path>())
        {
            foreach (var child in path.ChildElements)
            {
                switch (child)
                {
                    case Drawing.MoveTo mt:
                        var mPt = mt.GetFirstChild<Drawing.Point>();
                        if (mPt != null)
                            sb.Append($"M{mPt.X?.Value ?? "0"},{mPt.Y?.Value ?? "0"} ");
                        break;
                    case Drawing.LineTo lt:
                        var lPt = lt.GetFirstChild<Drawing.Point>();
                        if (lPt != null)
                            sb.Append($"L{lPt.X?.Value ?? "0"},{lPt.Y?.Value ?? "0"} ");
                        break;
                    case Drawing.CubicBezierCurveTo cb:
                        var pts = cb.Elements<Drawing.Point>().ToList();
                        if (pts.Count >= 3)
                            sb.Append($"C{pts[0].X?.Value ?? "0"},{pts[0].Y?.Value ?? "0"} {pts[1].X?.Value ?? "0"},{pts[1].Y?.Value ?? "0"} {pts[2].X?.Value ?? "0"},{pts[2].Y?.Value ?? "0"} ");
                        break;
                    case Drawing.QuadraticBezierCurveTo qb:
                        var qPts = qb.Elements<Drawing.Point>().ToList();
                        if (qPts.Count >= 2)
                            sb.Append($"Q{qPts[0].X?.Value ?? "0"},{qPts[0].Y?.Value ?? "0"} {qPts[1].X?.Value ?? "0"},{qPts[1].Y?.Value ?? "0"} ");
                        break;
                    case Drawing.ArcTo at:
                        sb.Append($"A{at.WidthRadius?.Value ?? "0"},{at.HeightRadius?.Value ?? "0"} ");
                        break;
                    case Drawing.CloseShapePath:
                        sb.Append("Z ");
                        break;
                }
            }
        }

        return sb.ToString().Trim();
    }

    private static readonly Dictionary<string, string> _tableStyleGuidToName = new(StringComparer.OrdinalIgnoreCase)
    {
        ["{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}"] = "medium1",
        ["{F5AB1C69-6EDB-4FF4-983F-18BD219EF322}"] = "medium2",
        ["{3B4B98B0-60AC-42C2-AFA5-B58CD77FA1E5}"] = "medium3",
        ["{D7AC3CCA-C797-4891-BE02-D94E43425B78}"] = "medium4",
        ["{9D7B26C5-4107-4FEC-AEDC-1716B250A1EF}"] = "light1",
        ["{ED083AE6-46FA-4A59-8FB0-9F97EB10719F}"] = "light2",
        ["{C083E6E3-FA7D-4D7B-A595-EF9225AFEA82}"] = "light3",
        ["{E8034E78-7F5D-4C2E-B375-FC64B27BC917}"] = "dark1",
        ["{125E5076-3810-47DD-B79F-674D7AD40C01}"] = "dark2",
        ["{2D5ABB26-0587-4C30-8999-92F81FD0307C}"] = "none",
    };

    private static string? TableStyleGuidToName(string guid)
    {
        return _tableStyleGuidToName.TryGetValue(guid, out var name) ? name : null;
    }

    // ==================== Effective Properties Resolution (PPT) ====================

    /// <summary>
    /// Populates effective.* format keys on a shape node for font properties not explicitly set.
    /// Resolves from: shape placeholder → layout → master text styles → presentation defaults → theme.
    /// </summary>
    private static void PopulateEffectiveShapeProperties(DocumentNode node, Shape shape, OpenXmlPart? part)
    {
        if (part is not SlidePart slidePart) return;

        // Determine placeholder info for style resolution
        var ph = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
            ?.GetFirstChild<PlaceholderShape>();
        var phType = ph?.Type?.HasValue == true ? ph.Type.Value : PlaceholderValues.Body;
        bool isTitle = phType == PlaceholderValues.Title || phType == PlaceholderValues.CenteredTitle;
        bool isSubTitle = phType == PlaceholderValues.SubTitle;

        // Resolve effective font size
        if (!node.Format.ContainsKey("size"))
        {
            var effSize = ResolveEffectiveFontSize(shape, slidePart, ph, isTitle, isSubTitle, phType);
            if (effSize.HasValue)
                node.Format["effective.size"] = $"{effSize.Value / 100.0:0.##}pt";
        }

        // Resolve effective font name from theme
        if (!node.Format.ContainsKey("font"))
        {
            var effFont = ResolveEffectiveFont(shape, slidePart, ph, isTitle);
            if (effFont != null)
                node.Format["effective.font"] = effFont;
        }

        // Resolve effective color
        if (!node.Format.ContainsKey("color"))
        {
            var effColor = ResolveEffectiveColor(shape, slidePart, ph, isTitle, isSubTitle, phType);
            if (effColor != null)
                node.Format["effective.color"] = effColor;
        }

        // Resolve effective bold
        if (!node.Format.ContainsKey("bold"))
        {
            var effBold = ResolveEffectiveBold(shape, slidePart, ph, isTitle, isSubTitle, phType);
            if (effBold == true)
                node.Format["effective.bold"] = true;
        }
    }

    /// <summary>
    /// Populates effective.* format keys on a run node for properties not explicitly set.
    /// </summary>
    private static void PopulateEffectiveRunProperties(DocumentNode node, Drawing.Run run, OpenXmlPart? part)
    {
        if (part is not SlidePart slidePart) return;

        // Walk up to find the containing shape
        var shape = run.Ancestors<Shape>().FirstOrDefault();
        if (shape == null) return;

        var ph = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
            ?.GetFirstChild<PlaceholderShape>();
        var phType = ph?.Type?.HasValue == true ? ph.Type.Value : PlaceholderValues.Body;
        bool isTitle = phType == PlaceholderValues.Title || phType == PlaceholderValues.CenteredTitle;
        bool isSubTitle = phType == PlaceholderValues.SubTitle;

        // Determine the paragraph level for this run
        var para = run.Ancestors<Drawing.Paragraph>().FirstOrDefault();
        int level = para?.ParagraphProperties?.Level?.Value ?? 0;

        if (!node.Format.ContainsKey("size"))
        {
            var effSize = ResolveEffectiveFontSize(shape, slidePart, ph, isTitle, isSubTitle, phType, level);
            if (effSize.HasValue)
                node.Format["effective.size"] = $"{effSize.Value / 100.0:0.##}pt";
        }

        if (!node.Format.ContainsKey("font"))
        {
            var effFont = ResolveEffectiveFont(shape, slidePart, ph, isTitle);
            if (effFont != null)
                node.Format["effective.font"] = effFont;
        }

        if (!node.Format.ContainsKey("color"))
        {
            var effColor = ResolveEffectiveColor(shape, slidePart, ph, isTitle, isSubTitle, phType, level);
            if (effColor != null)
                node.Format["effective.color"] = effColor;
        }

        if (!node.Format.ContainsKey("bold"))
        {
            var effBold = ResolveEffectiveBold(shape, slidePart, ph, isTitle, isSubTitle, phType, level);
            if (effBold == true)
                node.Format["effective.bold"] = true;
        }
    }

    /// <summary>
    /// Resolves font size from: shape lstStyle → layout/master placeholder → master text styles → presentation defaults.
    /// </summary>
    private static int? ResolveEffectiveFontSize(Shape shape, SlidePart slidePart,
        PlaceholderShape? ph, bool isTitle, bool isSubTitle, PlaceholderValues phType, int level = 0)
    {
        // 1. Shape's own list style
        var lstStyle = shape.TextBody?.GetFirstChild<Drawing.ListStyle>();
        var defRp = GetLevelDefRp(lstStyle, level);
        if (defRp?.FontSize?.HasValue == true)
            return defRp.FontSize.Value;

        // 2. Layout/master placeholder matching
        if (ph != null)
        {
            var layoutTree = slidePart.SlideLayoutPart?.SlideLayout?.CommonSlideData?.ShapeTree;
            var masterTree = slidePart.SlideLayoutPart?.SlideMasterPart?.SlideMaster?.CommonSlideData?.ShapeTree;
            foreach (var tree in new[] { layoutTree, masterTree })
            {
                if (tree == null) continue;
                foreach (var candidate in tree.Elements<Shape>())
                {
                    var cPh = candidate.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
                        ?.GetFirstChild<PlaceholderShape>();
                    if (cPh == null) continue;
                    if (!PlaceholderMatches(ph, cPh)) continue;
                    var cLstStyle = candidate.TextBody?.GetFirstChild<Drawing.ListStyle>();
                    var cDefRp = GetLevelDefRp(cLstStyle, level);
                    if (cDefRp?.FontSize?.HasValue == true)
                        return cDefRp.FontSize.Value;
                }
            }
        }

        // 3. Master text styles
        var masterTxStyles = slidePart.SlideLayoutPart?.SlideMasterPart?.SlideMaster?.TextStyles;
        if (masterTxStyles != null)
        {
            OpenXmlCompositeElement? styleList = isTitle ? masterTxStyles.TitleStyle
                : (isSubTitle || phType == PlaceholderValues.Body || phType == PlaceholderValues.Object)
                    ? masterTxStyles.BodyStyle : masterTxStyles.OtherStyle;
            if (styleList != null)
            {
                var sDefRp = GetLevelDefRp(styleList, level);
                if (sDefRp?.FontSize?.HasValue == true) return sDefRp.FontSize.Value;
            }
        }

        // 4. Presentation-level defaultTextStyle
        var presStyle = GetPresentationDefaultTextStyle(slidePart);
        if (presStyle != null)
        {
            var pDefRp = GetLevelDefRp(presStyle, level);
            if (pDefRp?.FontSize?.HasValue == true) return pDefRp.FontSize.Value;
        }

        return null;
    }

    /// <summary>
    /// Resolves font name from: theme fonts (major for titles, minor for body).
    /// </summary>
    private static string? ResolveEffectiveFont(Shape shape, SlidePart slidePart,
        PlaceholderShape? ph, bool isTitle)
    {
        // Check layout/master placeholder for explicit font
        if (ph != null)
        {
            var layoutTree = slidePart.SlideLayoutPart?.SlideLayout?.CommonSlideData?.ShapeTree;
            var masterTree = slidePart.SlideLayoutPart?.SlideMasterPart?.SlideMaster?.CommonSlideData?.ShapeTree;
            foreach (var tree in new[] { layoutTree, masterTree })
            {
                if (tree == null) continue;
                foreach (var candidate in tree.Elements<Shape>())
                {
                    var cPh = candidate.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
                        ?.GetFirstChild<PlaceholderShape>();
                    if (cPh == null || !PlaceholderMatches(ph, cPh)) continue;
                    var cRun = candidate.TextBody?.Descendants<Drawing.Run>().FirstOrDefault();
                    var cFont = cRun?.RunProperties?.GetFirstChild<Drawing.LatinFont>()?.Typeface?.Value
                        ?? cRun?.RunProperties?.GetFirstChild<Drawing.EastAsianFont>()?.Typeface?.Value;
                    if (cFont != null && !cFont.StartsWith("+", StringComparison.Ordinal))
                        return cFont;
                }
            }
        }

        // Theme fonts
        var theme = slidePart.SlideLayoutPart?.SlideMasterPart?.ThemePart?.Theme;
        var fontScheme = theme?.ThemeElements?.FontScheme;
        if (fontScheme == null) return null;

        if (isTitle)
        {
            return fontScheme.MajorFont?.GetFirstChild<Drawing.LatinFont>()?.Typeface?.Value
                ?? fontScheme.MajorFont?.GetFirstChild<Drawing.EastAsianFont>()?.Typeface?.Value;
        }
        else
        {
            return fontScheme.MinorFont?.GetFirstChild<Drawing.LatinFont>()?.Typeface?.Value
                ?? fontScheme.MinorFont?.GetFirstChild<Drawing.EastAsianFont>()?.Typeface?.Value;
        }
    }

    /// <summary>
    /// Resolves text color from master text styles and presentation defaults.
    /// </summary>
    private static string? ResolveEffectiveColor(Shape shape, SlidePart slidePart,
        PlaceholderShape? ph, bool isTitle, bool isSubTitle, PlaceholderValues phType, int level = 0)
    {
        // 1. Layout/master placeholder
        if (ph != null)
        {
            var layoutTree = slidePart.SlideLayoutPart?.SlideLayout?.CommonSlideData?.ShapeTree;
            var masterTree = slidePart.SlideLayoutPart?.SlideMasterPart?.SlideMaster?.CommonSlideData?.ShapeTree;
            foreach (var tree in new[] { layoutTree, masterTree })
            {
                if (tree == null) continue;
                foreach (var candidate in tree.Elements<Shape>())
                {
                    var cPh = candidate.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
                        ?.GetFirstChild<PlaceholderShape>();
                    if (cPh == null || !PlaceholderMatches(ph, cPh)) continue;
                    var cLstStyle = candidate.TextBody?.GetFirstChild<Drawing.ListStyle>();
                    var cDefRp = GetLevelDefRp(cLstStyle, level);
                    var cColor = ReadColorFromFill(cDefRp?.GetFirstChild<Drawing.SolidFill>());
                    if (cColor != null) return cColor;
                }
            }
        }

        // 2. Master text styles
        var masterTxStyles = slidePart.SlideLayoutPart?.SlideMasterPart?.SlideMaster?.TextStyles;
        if (masterTxStyles != null)
        {
            OpenXmlCompositeElement? styleList = isTitle ? masterTxStyles.TitleStyle
                : (isSubTitle || phType == PlaceholderValues.Body || phType == PlaceholderValues.Object)
                    ? masterTxStyles.BodyStyle : masterTxStyles.OtherStyle;
            if (styleList != null)
            {
                var sDefRp = GetLevelDefRp(styleList, level);
                var sColor = ReadColorFromFill(sDefRp?.GetFirstChild<Drawing.SolidFill>());
                if (sColor != null) return sColor;
            }
        }

        return null;
    }

    /// <summary>
    /// Resolves bold from master text styles.
    /// </summary>
    private static bool? ResolveEffectiveBold(Shape shape, SlidePart slidePart,
        PlaceholderShape? ph, bool isTitle, bool isSubTitle, PlaceholderValues phType, int level = 0)
    {
        // Master text styles
        var masterTxStyles = slidePart.SlideLayoutPart?.SlideMasterPart?.SlideMaster?.TextStyles;
        if (masterTxStyles != null)
        {
            OpenXmlCompositeElement? styleList = isTitle ? masterTxStyles.TitleStyle
                : (isSubTitle || phType == PlaceholderValues.Body || phType == PlaceholderValues.Object)
                    ? masterTxStyles.BodyStyle : masterTxStyles.OtherStyle;
            if (styleList != null)
            {
                var sDefRp = GetLevelDefRp(styleList, level);
                if (sDefRp?.Bold?.HasValue == true) return sDefRp.Bold.Value;
            }
        }

        return null;
    }

    /// <summary>
    /// Gets the presentation-level DefaultTextStyle by navigating from a SlidePart.
    /// </summary>
    private static OpenXmlCompositeElement? GetPresentationDefaultTextStyle(SlidePart slidePart)
    {
        // Navigate: SlidePart → SlideLayoutPart → SlideMasterPart → PresentationPart → Presentation
        var masterPart = slidePart.SlideLayoutPart?.SlideMasterPart;
        if (masterPart == null) return null;

        // The SlideMasterPart's parent relationships include the PresentationPart
        // We can access the Presentation through the package
        foreach (var rel in masterPart.Parts)
        {
            if (rel.OpenXmlPart is PresentationPart presPart)
                return presPart.Presentation?.DefaultTextStyle;
        }

        return null;
    }
}
