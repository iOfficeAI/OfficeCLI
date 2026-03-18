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
        if (dash?.Val?.HasValue == true) parts.Add(dash.Val.InnerText);
        else parts.Add("solid");
        if (color != null) parts.Add(color);
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
}
