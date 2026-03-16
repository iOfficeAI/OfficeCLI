// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
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
            else if (gf.Descendants<C.ChartReference>().Any())
            {
                chartIdx++;
                children.Add(ChartToNode(gf, slidePart, slideNum, chartIdx, depth));
            }
        }

        int picIdx = 0;
        foreach (var pic in shapeTree.Elements<Picture>())
        {
            children.Add(PictureToNode(pic, slideNum, picIdx + 1));
            picIdx++;
        }

        return children;
    }

    private static DocumentNode TableToNode(GraphicFrame gf, int slideNum, int tblIdx, int depth)
    {
        var table = gf.Descendants<Drawing.Table>().First();
        var rows = table.Elements<Drawing.TableRow>().ToList();
        var cols = rows.FirstOrDefault()?.Elements<Drawing.TableCell>().Count() ?? 0;
        var name = gf.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Name?.Value ?? "Table";

        var node = new DocumentNode
        {
            Path = $"/slide[{slideNum}]/table[{tblIdx}]",
            Type = "table",
            Preview = $"{name} ({rows.Count}x{cols})",
            ChildCount = rows.Count
        };

        node.Format["name"] = name;
        node.Format["rows"] = rows.Count;
        node.Format["cols"] = cols;

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
                    Path = $"/slide[{slideNum}]/table[{tblIdx}]/tr[{rIdx}]",
                    Type = "tr",
                    ChildCount = row.Elements<Drawing.TableCell>().Count()
                };

                if (depth > 1)
                {
                    int cIdx = 0;
                    foreach (var cell in row.Elements<Drawing.TableCell>())
                    {
                        cIdx++;
                        var cellText = cell.TextBody?.InnerText ?? "";
                        var cellNode = new DocumentNode
                        {
                            Path = $"/slide[{slideNum}]/table[{tblIdx}]/tr[{rIdx}]/tc[{cIdx}]",
                            Type = "tc",
                            Text = cellText
                        };

                        // Cell fill
                        var tcPr = cell.TableCellProperties ?? cell.GetFirstChild<Drawing.TableCellProperties>();
                        var cellFillHex = tcPr?.GetFirstChild<Drawing.SolidFill>()?.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value;
                        if (cellFillHex != null) cellNode.Format["fill"] = cellFillHex;

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

        var node = new DocumentNode
        {
            Path = $"/slide[{slideNum}]/shape[{shapeIdx}]",
            Type = isTitle ? "title" : "textbox",
            Text = text,
            Preview = string.IsNullOrEmpty(text) ? name : (text.Length > 50 ? text[..50] + "..." : text)
        };

        node.Format["name"] = name;
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
        var shapeFillHex = shapeFill?.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value;
        if (shapeFillHex != null) node.Format["fill"] = shapeFillHex;
        if (shape.ShapeProperties?.GetFirstChild<Drawing.NoFill>() != null) node.Format["fill"] = "none";

        // Opacity (Alpha on SolidFill)
        var alphaVal = shapeFill?.GetFirstChild<Drawing.RgbColorModelHex>()
            ?.GetFirstChild<Drawing.Alpha>()?.Val?.Value;
        if (alphaVal.HasValue) node.Format["opacity"] = $"{alphaVal.Value / 100000.0:0.##}";

        // Shape preset
        var presetGeom = shape.ShapeProperties?.GetFirstChild<Drawing.PresetGeometry>();
        if (presetGeom?.Preset?.HasValue == true)
            node.Format["preset"] = presetGeom.Preset.InnerText;

        // Gradient fill
        var gradFill = shape.ShapeProperties?.GetFirstChild<Drawing.GradientFill>();
        if (gradFill != null)
        {
            var stops = gradFill.GradientStopList?.Elements<Drawing.GradientStop>()
                .Select(gs => gs.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value ?? "?")
                .ToList();
            if (stops?.Count > 0)
            {
                var pathGrad = gradFill.GetFirstChild<Drawing.PathGradientFill>();
                if (pathGrad != null)
                {
                    // Radial/path gradient — decode focus point from FillToRectangle
                    var fillRect = pathGrad.GetFirstChild<Drawing.FillToRectangle>();
                    var focus = "center";
                    if (fillRect != null)
                    {
                        var fl = fillRect.Left?.Value ?? 50000;
                        var ft = fillRect.Top?.Value ?? 50000;
                        focus = (fl, ft) switch
                        {
                            (0, 0) => "tl",
                            ( >= 100000, 0) => "tr",
                            (0, >= 100000) => "bl",
                            ( >= 100000, >= 100000) => "br",
                            _ => "center"
                        };
                    }
                    node.Format["gradient"] = $"radial:{string.Join("-", stops)}-{focus}";
                }
                else
                {
                    var gradStr = string.Join("-", stops);
                    var linear = gradFill.GetFirstChild<Drawing.LinearGradientFill>();
                    if (linear?.Angle?.HasValue == true)
                        gradStr += $"-{linear.Angle.Value / 60000}";
                    node.Format["gradient"] = gradStr;
                }
            }
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
            if (charBullet != null) node.Format["list"] = charBullet.Char?.Value ?? "•";
            else if (autoBullet?.Type?.HasValue == true) node.Format["list"] = autoBullet.Type.InnerText;
        }

        // Collect font info
        var firstRun = shape.TextBody?.Descendants<Drawing.Run>().FirstOrDefault();
        if (firstRun?.RunProperties != null)
        {
            var font = firstRun.RunProperties.GetFirstChild<Drawing.LatinFont>()?.Typeface
                ?? firstRun.RunProperties.GetFirstChild<Drawing.EastAsianFont>()?.Typeface;
            if (font != null) node.Format["font"] = font;

            var fontSize = firstRun.RunProperties.FontSize?.Value;
            if (fontSize.HasValue) node.Format["size"] = $"{fontSize.Value / 100}pt";

            if (firstRun.RunProperties.Bold?.Value == true) node.Format["bold"] = true;
            if (firstRun.RunProperties.Italic?.Value == true) node.Format["italic"] = true;
            if (firstRun.RunProperties.Underline?.HasValue == true && firstRun.RunProperties.Underline.Value != Drawing.TextUnderlineValues.None)
                node.Format["underline"] = firstRun.RunProperties.Underline.InnerText;
            if (firstRun.RunProperties.Strike?.HasValue == true && firstRun.RunProperties.Strike.Value != Drawing.TextStrikeValues.NoStrike)
                node.Format["strikethrough"] = firstRun.RunProperties.Strike.InnerText;

            // Text color (from first run)
            var runColorHex = firstRun.RunProperties.GetFirstChild<Drawing.SolidFill>()
                ?.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value;
            if (runColorHex != null) node.Format["color"] = runColorHex;

            // Hyperlink on first run
            if (part != null)
            {
                var linkUrl = ReadRunHyperlinkUrl(firstRun, part);
                if (linkUrl != null) node.Format["link"] = linkUrl;
            }
        }

        // Line/border
        var outline = shape.ShapeProperties?.GetFirstChild<Drawing.Outline>();
        if (outline != null)
        {
            var lineColorNode = outline.GetFirstChild<Drawing.SolidFill>()?.GetFirstChild<Drawing.RgbColorModelHex>();
            var lineFill = lineColorNode?.Val?.Value;
            if (lineFill != null) node.Format["line"] = lineFill;
            if (outline.GetFirstChild<Drawing.NoFill>() != null) node.Format["line"] = "none";
            if (outline.Width?.HasValue == true) node.Format["lineWidth"] = FormatEmu(outline.Width.Value);
            var dash = outline.GetFirstChild<Drawing.PresetDash>();
            if (dash?.Val?.HasValue == true) node.Format["lineDash"] = dash.Val.InnerText.ToLowerInvariant();
            var lineAlpha = lineColorNode?.GetFirstChild<Drawing.Alpha>()?.Val?.Value;
            if (lineAlpha.HasValue) node.Format["lineOpacity"] = $"{lineAlpha.Value / 100000.0:0.##}";
        }

        // Effects (shadow, glow, reflection)
        var effectList = shape.ShapeProperties?.GetFirstChild<Drawing.EffectList>();
        if (effectList != null)
        {
            var outerShadow = effectList.GetFirstChild<Drawing.OuterShadow>();
            if (outerShadow != null)
            {
                var colorHex = outerShadow.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value ?? "000000";
                node.Format["shadow"] = colorHex;
            }
            var glow = effectList.GetFirstChild<Drawing.Glow>();
            if (glow != null)
            {
                var colorHex = glow.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value ?? "000000";
                node.Format["glow"] = colorHex;
            }
            if (effectList.GetFirstChild<Drawing.Reflection>() != null)
                node.Format["reflection"] = "true";
        }

        // Rotation
        if (xfrm?.Rotation != null && xfrm.Rotation.Value != 0)
            node.Format["rotation"] = $"{xfrm.Rotation.Value / 60000.0}°";

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

            // Vertical alignment
            if (bodyPr.Anchor?.HasValue == true)
                node.Format["valign"] = bodyPr.Anchor.InnerText;

            // AutoFit
            if (bodyPr.GetFirstChild<Drawing.NormalAutoFit>() != null) node.Format["autoFit"] = "normal";
            else if (bodyPr.GetFirstChild<Drawing.ShapeAutoFit>() != null) node.Format["autoFit"] = "shape";
            else if (bodyPr.GetFirstChild<Drawing.NoAutoFit>() != null) node.Format["autoFit"] = "none";
        }

        // Text alignment (from first paragraph)
        var firstPara = shape.TextBody?.Elements<Drawing.Paragraph>().FirstOrDefault();
        if (firstPara?.ParagraphProperties?.Alignment?.HasValue == true)
            node.Format["align"] = firstPara.ParagraphProperties.Alignment.InnerText;

        // Paragraph spacing (from first paragraph)
        var pProps = firstPara?.ParagraphProperties;
        if (pProps != null)
        {
            var ls = pProps.GetFirstChild<Drawing.LineSpacing>()?.GetFirstChild<Drawing.SpacingPercent>()?.Val?.Value;
            if (ls.HasValue) node.Format["lineSpacing"] = $"{ls.Value / 1000.0:0.##}";
            var sb = pProps.GetFirstChild<Drawing.SpaceBefore>()?.GetFirstChild<Drawing.SpacingPoints>()?.Val?.Value;
            if (sb.HasValue) node.Format["spaceBefore"] = $"{sb.Value / 100.0:0.##}";
            var sa = pProps.GetFirstChild<Drawing.SpaceAfter>()?.GetFirstChild<Drawing.SpacingPoints>()?.Val?.Value;
            if (sa.HasValue) node.Format["spaceAfter"] = $"{sa.Value / 100.0:0.##}";
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
                        Path = $"/slide[{slideNum}]/shape[{shapeIdx}]/paragraph[{paraIdx + 1}]",
                        Type = "paragraph",
                        Text = paraText,
                        ChildCount = paraRuns.Count
                    };

                    // Add alignment info
                    var align = para.ParagraphProperties?.Alignment;
                    if (align != null && align.HasValue) paraNode.Format["align"] = align.InnerText;

                    // Include runs at depth > 1
                    if (depth > 1)
                    {
                        int runIdx = 0;
                        foreach (var run in paraRuns)
                        {
                            paraNode.Children.Add(RunToNode(run,
                                $"/slide[{slideNum}]/shape[{shapeIdx}]/paragraph[{paraIdx + 1}]/run[{runIdx + 1}]", part));
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
            var f = run.RunProperties.GetFirstChild<Drawing.LatinFont>()?.Typeface
                ?? run.RunProperties.GetFirstChild<Drawing.EastAsianFont>()?.Typeface;
            if (f != null) node.Format["font"] = f;
            var fs = run.RunProperties.FontSize?.Value;
            if (fs.HasValue) node.Format["size"] = $"{fs.Value / 100}pt";
            if (run.RunProperties.Bold?.Value == true) node.Format["bold"] = true;
            if (run.RunProperties.Italic?.Value == true) node.Format["italic"] = true;
            // Color
            var solidFill = run.RunProperties.GetFirstChild<Drawing.SolidFill>();
            var rgbHex = solidFill?.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value;
            if (rgbHex != null) node.Format["color"] = rgbHex;
            // Hyperlink
            if (part != null)
            {
                var linkUrl = ReadRunHyperlinkUrl(run, part);
                if (linkUrl != null) node.Format["link"] = linkUrl;
            }
        }

        return node;
    }

    private static DocumentNode PictureToNode(Picture pic, int slideNum, int picIdx)
    {
        var name = pic.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value ?? "Picture";
        var alt = pic.NonVisualPictureProperties?.NonVisualDrawingProperties?.Description?.Value;

        var node = new DocumentNode
        {
            Path = $"/slide[{slideNum}]/picture[{picIdx}]",
            Type = "picture",
            Preview = name
        };

        node.Format["name"] = name;
        if (!string.IsNullOrEmpty(alt)) node.Format["alt"] = alt;
        else node.Format["alt"] = "(missing)";

        return node;
    }

    private static Shape CreateTextShape(uint id, string name, string text, bool isTitle)
    {
        var shape = new Shape();
        shape.NonVisualShapeProperties = new NonVisualShapeProperties(
            new NonVisualDrawingProperties { Id = id, Name = name },
            new NonVisualShapeDrawingProperties(),
            new ApplicationNonVisualDrawingProperties(
                isTitle ? new PlaceholderShape { Type = PlaceholderValues.Title } : new PlaceholderShape()
            )
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
                    new Drawing.RunProperties { Language = "zh-CN" },
                    new Drawing.Text(line)
                )
            ));
        }
        shape.TextBody = body;
        return shape;
    }
}
