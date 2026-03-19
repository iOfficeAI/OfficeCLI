// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler
{
    private static bool IsTruthy(string value) =>
        ParseHelpers.IsTruthy(value);

    /// <summary>
    /// Find existing Transition element or create one, avoiding duplicates with unknown-element transitions.
    /// </summary>
    private static Transition FindOrCreateTransition(Slide slide)
    {
        var typed = slide.GetFirstChild<Transition>();
        if (typed != null) return typed;

        // Check for unknown-element transitions (injected as raw XML to survive SDK serialization)
        var unknown = slide.ChildElements.FirstOrDefault(c => c.LocalName == "transition" && c is not Transition);
        if (unknown != null)
        {
            // Replace with a typed Transition so we can set properties
            var trans = new Transition();
            foreach (var attr in unknown.GetAttributes()) trans.SetAttribute(attr);
            trans.InnerXml = unknown.InnerXml;
            unknown.InsertAfterSelf(trans);
            unknown.Remove();
            return trans;
        }

        return slide.AppendChild(new Transition());
    }

    private static double ParseFontSize(string value) =>
        ParseHelpers.ParseFontSize(value);

    /// <summary>
    /// Read table cell border properties following POI's getBorderWidth/getBorderColor pattern.
    /// Maps a:lnL/lnR/lnT/lnB → border.left, border.right, border.top, border.bottom in Format.
    /// </summary>
    private static void ReadTableCellBorders(Drawing.TableCellProperties tcPr, DocumentNode node)
    {
        ReadBorderLine(tcPr.LeftBorderLineProperties, "border.left", node);
        ReadBorderLine(tcPr.RightBorderLineProperties, "border.right", node);
        ReadBorderLine(tcPr.TopBorderLineProperties, "border.top", node);
        ReadBorderLine(tcPr.BottomBorderLineProperties, "border.bottom", node);
        ReadBorderLine(tcPr.TopLeftToBottomRightBorderLineProperties, "border.tl2br", node);
        ReadBorderLine(tcPr.BottomLeftToTopRightBorderLineProperties, "border.tr2bl", node);
    }

    /// <summary>
    /// Read a single border line's properties (color, width, dash) following POI's pattern:
    /// - Returns nothing if line is null, has NoFill, or lacks SolidFill
    /// - Reads width from w attribute, color from SolidFill, dash from PresetDash
    /// </summary>
    private static void ReadBorderLine(OpenXmlCompositeElement? lineProps, string prefix, DocumentNode node)
    {
        if (lineProps == null) return;
        // POI: if NoFill is set, the border is invisible — skip
        if (lineProps.GetFirstChild<Drawing.NoFill>() != null) return;
        var solidFill = lineProps.GetFirstChild<Drawing.SolidFill>();
        if (solidFill == null) return; // POI: !isSetSolidFill → null

        var color = ReadColorFromFill(solidFill);
        if (color != null) node.Format[$"{prefix}.color"] = color;

        // Width from "w" attribute (EMU) — POI: Units.toPoints(ln.getW())
        var wAttr = lineProps.GetAttributes().FirstOrDefault(a => a.LocalName == "w");
        if (!string.IsNullOrEmpty(wAttr.Value) && long.TryParse(wAttr.Value, out var wEmu) && wEmu > 0)
            node.Format[$"{prefix}.width"] = FormatEmu(wEmu);

        // Dash style from PresetDash — POI: ln.getPrstDash().getVal()
        var dash = lineProps.GetFirstChild<Drawing.PresetDash>();
        if (dash?.Val?.HasValue == true)
            node.Format[$"{prefix}.dash"] = dash.Val.InnerText;

        // Summary key: "1pt solid FF0000" format for convenience
        var parts = new List<string>();
        if (!string.IsNullOrEmpty(wAttr.Value) && long.TryParse(wAttr.Value, out var wEmu2) && wEmu2 > 0)
            parts.Add(FormatEmu(wEmu2));
        if (dash?.Val?.HasValue == true) parts.Add(dash.Val.InnerText!);
        else parts.Add("solid");
        if (color is not null) parts.Add(color);
        if (parts.Count > 0) node.Format[prefix] = string.Join(" ", parts);
    }

    private static string GetShapeText(Shape shape)
    {
        var textBody = shape.TextBody;
        if (textBody == null) return "";

        var sb = new StringBuilder();
        var first = true;
        foreach (var para in textBody.Elements<Drawing.Paragraph>())
        {
            if (!first) sb.Append('\n');
            first = false;
            foreach (var child in para.ChildElements)
            {
                if (child is Drawing.Run run)
                    sb.Append(run.Text?.Text ?? "");
                else if (HasMathContent(child))
                    sb.Append(FormulaParser.ToReadableText(GetMathElement(child)));
            }
        }
        return sb.ToString();
    }

    /// <summary>
    /// Find all OMML math elements inside a shape's text body.
    /// </summary>
    private static List<OpenXmlElement> FindShapeMathElements(Shape shape)
    {
        var results = new List<OpenXmlElement>();
        var textBody = shape.TextBody;
        if (textBody == null) return results;

        foreach (var para in textBody.Elements<Drawing.Paragraph>())
        {
            foreach (var child in para.ChildElements)
            {
                if (HasMathContent(child))
                    results.Add(GetMathElement(child));
            }
        }
        return results;
    }

    /// <summary>
    /// Check if an element contains math content (a14:m or mc:AlternateContent with math).
    /// </summary>
    private static bool HasMathContent(OpenXmlElement element)
    {
        if (element.LocalName == "m" && element.NamespaceUri == "http://schemas.microsoft.com/office/drawing/2010/main")
            return true;
        if (element is AlternateContent || element.LocalName == "AlternateContent")
        {
            if (element.Descendants().Any(e => e.LocalName == "oMath" || e.LocalName == "oMathPara"))
                return true;
            return element.InnerXml.Contains("oMath");
        }
        return false;
    }

    /// <summary>
    /// Extract the OMML math element from an a14:m or mc:AlternateContent wrapper.
    /// </summary>
    private static OpenXmlElement GetMathElement(OpenXmlElement element)
    {
        if (element.LocalName == "m" && element.NamespaceUri == "http://schemas.microsoft.com/office/drawing/2010/main")
        {
            var child = element.ChildElements.FirstOrDefault(e => e.LocalName == "oMathPara" || e.LocalName == "oMath");
            if (child != null) return child;

            var desc = element.Descendants().FirstOrDefault(e => e.LocalName == "oMathPara" || e.LocalName == "oMath");
            if (desc != null) return desc;

            var innerXml = element.InnerXml;
            if (!string.IsNullOrEmpty(innerXml) && innerXml.Contains("oMath"))
                return ReparseFromXml(innerXml) ?? element;

            return element;
        }
        if (element is AlternateContent || element.LocalName == "AlternateContent")
        {
            var choice = element.ChildElements.FirstOrDefault(e => e is AlternateContentChoice || e.LocalName == "Choice");
            if (choice != null)
            {
                var a14m = choice.ChildElements.FirstOrDefault(e =>
                    e.LocalName == "m" && e.NamespaceUri == "http://schemas.microsoft.com/office/drawing/2010/main");
                if (a14m != null)
                    return GetMathElement(a14m);

                var mathDesc = choice.Descendants().FirstOrDefault(e => e.LocalName == "oMathPara" || e.LocalName == "oMath");
                if (mathDesc != null)
                    return mathDesc;
            }

            var innerXml = element.InnerXml;
            if (!string.IsNullOrEmpty(innerXml) && innerXml.Contains("oMath"))
                return ReparseFromXml(innerXml) ?? element;
        }
        return element;
    }

    /// <summary>
    /// Re-parse OMML XML string into an OpenXmlElement with navigable children.
    /// </summary>
    private static OpenXmlElement? ReparseFromXml(string innerXml)
    {
        try
        {
            var xml = innerXml.Trim();
            if (xml.Contains("oMathPara"))
            {
                var startIdx = xml.IndexOf("<m:oMathPara", StringComparison.Ordinal);
                if (startIdx < 0) startIdx = xml.IndexOf("<oMathPara", StringComparison.Ordinal);
                if (startIdx >= 0)
                {
                    var endTag = xml.Contains("</m:oMathPara>") ? "</m:oMathPara>" : "</oMathPara>";
                    var endIdx = xml.IndexOf(endTag, StringComparison.Ordinal);
                    if (endIdx >= 0)
                    {
                        var oMathParaXml = xml[startIdx..(endIdx + endTag.Length)];
                        if (!oMathParaXml.Contains("xmlns:m="))
                            oMathParaXml = oMathParaXml.Replace("<m:oMathPara", "<m:oMathPara xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"");
                        var wrapper = new OpenXmlUnknownElement("m", "oMathPara", "http://schemas.openxmlformats.org/officeDocument/2006/math");
                        var innerStart = oMathParaXml.IndexOf('>') + 1;
                        var innerEnd = oMathParaXml.LastIndexOf('<');
                        if (innerStart > 0 && innerEnd > innerStart)
                            wrapper.InnerXml = oMathParaXml[innerStart..innerEnd];
                        return wrapper;
                    }
                }
            }
        }
        catch { }
        return null;
    }

    private static bool IsTitle(Shape shape)
    {
        var ph = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
            ?.GetFirstChild<PlaceholderShape>();
        if (ph == null) return false;
        var type = ph.Type?.Value;
        return type == PlaceholderValues.Title || type == PlaceholderValues.CenteredTitle;
    }

    private static string GetShapeName(Shape shape) =>
        shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value ?? "?";

    private static long ParseEmu(string value) => Core.EmuConverter.ParseEmu(value);

    private static string FormatEmu(long emu) => Core.EmuConverter.FormatEmu(emu);

    /// <summary>
    /// Read a GradientFill element and return a string representation (C1-C2[-angle] or radial:C1-C2[-focus]).
    /// </summary>
    internal static string ReadGradientString(Drawing.GradientFill gradFill)
    {
        var stopEls = gradFill.GradientStopList?.Elements<Drawing.GradientStop>().ToList();
        if (stopEls == null || stopEls.Count == 0) return "gradient";

        var stopData = stopEls.Select(gs => (
            color: gs.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value ?? "?",
            pos: gs.Position?.Value
        )).ToList();

        // Check if positions deviate >1% from even distribution (1000 units)
        bool hasCustomPos = false;
        int n = stopData.Count;
        for (int i = 0; i < n; i++)
        {
            var expectedPos = n == 1 ? 0 : (int)((long)i * 100000 / (n - 1));
            var actualPos = (int)(stopData[i].pos ?? 0);
            if (Math.Abs(actualPos - expectedPos) > 1000) { hasCustomPos = true; break; }
        }

        var stopStrs = stopData.Select((s, i) =>
            hasCustomPos && s.pos.HasValue
                ? $"{s.color}@{s.pos.Value / 1000}"
                : s.color
        ).ToList();

        var pathGrad = gradFill.GetFirstChild<Drawing.PathGradientFill>();
        if (pathGrad != null)
        {
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
            return $"radial:{string.Join("-", stopStrs)}-{focus}";
        }

        var linear = gradFill.GetFirstChild<Drawing.LinearGradientFill>();
        int deg = linear?.Angle?.HasValue == true ? linear.Angle.Value / 60000 : 0;
        return $"linear;{string.Join(";", stopStrs)};{deg}";
    }

    /// <summary>
    /// Parse SVG-like path syntax into a Drawing.CustomGeometry element.
    /// Format: "M x,y L x,y C x1,y1 x2,y2 x,y Q x1,y1 x,y Z"
    ///   M = moveTo, L = lineTo, C = cubicBezTo, Q = quadBezTo, A = arcTo, Z = close
    /// Coordinates are integers (EMU-scale, typically matching the shape's width/height).
    /// Example: "M 0,0 L 100,0 L 100,100 L 0,100 Z" (rectangle in 100x100 space)
    /// </summary>
    private static Drawing.CustomGeometry ParseCustomGeometry(string value)
    {
        var path = new Drawing.Path();

        // Parse SVG-like commands
        var tokens = value.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
        long maxX = 0, maxY = 0;
        int i = 0;

        while (i < tokens.Length)
        {
            var cmd = tokens[i].ToUpperInvariant();
            i++;

            switch (cmd)
            {
                case "M":
                {
                    var (x, y) = ParsePointToken(tokens[i++]);
                    path.AppendChild(new Drawing.MoveTo(new Drawing.Point { X = x.ToString(), Y = y.ToString() }));
                    TrackMax(ref maxX, ref maxY, x, y);
                    break;
                }
                case "L":
                {
                    var (x, y) = ParsePointToken(tokens[i++]);
                    path.AppendChild(new Drawing.LineTo(new Drawing.Point { X = x.ToString(), Y = y.ToString() }));
                    TrackMax(ref maxX, ref maxY, x, y);
                    break;
                }
                case "C":
                {
                    // Cubic bezier: 3 points (control1, control2, end)
                    var (x1, y1) = ParsePointToken(tokens[i++]);
                    var (x2, y2) = ParsePointToken(tokens[i++]);
                    var (x3, y3) = ParsePointToken(tokens[i++]);
                    path.AppendChild(new Drawing.CubicBezierCurveTo(
                        new Drawing.Point { X = x1.ToString(), Y = y1.ToString() },
                        new Drawing.Point { X = x2.ToString(), Y = y2.ToString() },
                        new Drawing.Point { X = x3.ToString(), Y = y3.ToString() }
                    ));
                    TrackMax(ref maxX, ref maxY, x3, y3);
                    break;
                }
                case "Q":
                {
                    // Quadratic bezier: 2 points (control, end)
                    var (x1, y1) = ParsePointToken(tokens[i++]);
                    var (x2, y2) = ParsePointToken(tokens[i++]);
                    path.AppendChild(new Drawing.QuadraticBezierCurveTo(
                        new Drawing.Point { X = x1.ToString(), Y = y1.ToString() },
                        new Drawing.Point { X = x2.ToString(), Y = y2.ToString() }
                    ));
                    TrackMax(ref maxX, ref maxY, x2, y2);
                    break;
                }
                case "Z":
                    path.AppendChild(new Drawing.CloseShapePath());
                    break;
                default:
                    // Skip unknown tokens
                    break;
            }
        }

        // Set path dimensions to bounding box
        if (maxX > 0) path.Width = maxX;
        if (maxY > 0) path.Height = maxY;

        return new Drawing.CustomGeometry(
            new Drawing.AdjustValueList(),
            new Drawing.ShapeGuideList(),
            new Drawing.AdjustHandleList(),
            new Drawing.ConnectionSiteList(),
            new Drawing.Rectangle { Left = "0", Top = "0", Right = "r", Bottom = "b" },
            new Drawing.PathList(path)
        );
    }

    private static (long x, long y) ParsePointToken(string token)
    {
        var parts = token.Split(',');
        if (parts.Length < 2)
            throw new ArgumentException($"Invalid coordinate '{token}'. Expected 'x,y' format (e.g. '100,200').");
        if (!long.TryParse(parts[0].Trim(), out var x))
            throw new ArgumentException($"Invalid x coordinate '{parts[0].Trim()}' in '{token}'. Expected a number.");
        if (!long.TryParse(parts[1].Trim(), out var y))
            throw new ArgumentException($"Invalid y coordinate '{parts[1].Trim()}' in '{token}'. Expected a number.");
        return (x, y);
    }

    private static void TrackMax(ref long maxX, ref long maxY, long x, long y)
    {
        if (x > maxX) maxX = x;
        if (y > maxY) maxY = y;
    }

    /// <summary>
    /// Change the z-order of a shape within the ShapeTree.
    /// Values: "front" (topmost), "back" (bottommost), "forward" (+1), "backward" (-1),
    ///         or an integer for absolute position (1-based, 1 = back, N = front).
    /// </summary>
    private static void ApplyZOrder(DocumentFormat.OpenXml.Packaging.SlidePart slidePart, Shape shape, string value)
    {
        var shapeTree = shape.Parent as ShapeTree
            ?? throw new InvalidOperationException("Shape is not in a ShapeTree");

        // Get all content elements (Shape, Picture, GraphicFrame, GroupShape, ConnectionShape)
        // that participate in z-order (skip structural elements like nvGrpSpPr, grpSpPr)
        var contentElements = shapeTree.ChildElements
            .Where(e => e is Shape or Picture or GraphicFrame or GroupShape or ConnectionShape)
            .ToList();
        var currentIndex = contentElements.IndexOf(shape);
        if (currentIndex < 0) return;

        int targetIndex;
        switch (value.ToLowerInvariant())
        {
            case "front" or "top" or "bringtofront":
                targetIndex = contentElements.Count - 1;
                break;
            case "back" or "bottom" or "sendtoback":
                targetIndex = 0;
                break;
            case "forward" or "bringforward" or "+1":
                targetIndex = Math.Min(currentIndex + 1, contentElements.Count - 1);
                break;
            case "backward" or "sendbackward" or "-1":
                targetIndex = Math.Max(currentIndex - 1, 0);
                break;
            default:
                // Absolute position (1-based: 1 = back, N = front)
                if (int.TryParse(value, out var pos))
                    targetIndex = Math.Clamp(pos - 1, 0, contentElements.Count - 1);
                else
                    throw new ArgumentException($"Invalid z-order value: {value}. Use front/back/forward/backward or a number.");
                break;
        }

        if (targetIndex == currentIndex) return;

        // Remove shape from its current position
        shape.Remove();

        // Insert at new position
        if (targetIndex >= contentElements.Count - 1)
        {
            // Front: append after last content element (or at end of tree)
            shapeTree.AppendChild(shape);
        }
        else if (targetIndex <= 0)
        {
            // Back: insert before the first content element
            var firstContent = shapeTree.ChildElements
                .FirstOrDefault(e => e is Shape or Picture or GraphicFrame or GroupShape or ConnectionShape);
            if (firstContent != null)
                firstContent.InsertBeforeSelf(shape);
            else
                shapeTree.AppendChild(shape);
        }
        else
        {
            // Refresh content list after removal
            var updatedContent = shapeTree.ChildElements
                .Where(e => e is Shape or Picture or GraphicFrame or GroupShape or ConnectionShape)
                .ToList();
            if (targetIndex < updatedContent.Count)
                updatedContent[targetIndex].InsertBeforeSelf(shape);
            else
                shapeTree.AppendChild(shape);
        }
    }
}
