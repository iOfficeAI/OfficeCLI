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

    /// <summary>
    /// Set advanceTime on a slide, handling morph AlternateContent correctly.
    /// </summary>
    internal static void SetAdvanceTime(Slide slide, string value)
    {
        var acMorph = slide.ChildElements.FirstOrDefault(c =>
            c.LocalName == "AlternateContent" && c.InnerXml.Contains("morph"));
        if (acMorph != null)
        {
            // Set advTm directly on transitions inside AlternateContent
            foreach (var trans in acMorph.Descendants().Where(d => d.LocalName == "transition"))
                trans.SetAttribute(new OpenXmlAttribute("", "advTm", null!, value));
        }
        else
        {
            FindOrCreateTransition(slide).AdvanceAfterTime = value;
        }
    }

    /// <summary>
    /// Set advanceOnClick on a slide, handling morph AlternateContent correctly.
    /// </summary>
    internal static void SetAdvanceClick(Slide slide, bool value)
    {
        var acMorph = slide.ChildElements.FirstOrDefault(c =>
            c.LocalName == "AlternateContent" && c.InnerXml.Contains("morph"));
        if (acMorph != null)
        {
            foreach (var trans in acMorph.Descendants().Where(d => d.LocalName == "transition"))
                trans.SetAttribute(new OpenXmlAttribute("", "advClick", null!, value ? "1" : "0"));
        }
        else
        {
            FindOrCreateTransition(slide).AdvanceOnClick = value;
        }
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
    /// Normalize DrawingML alignment abbreviations to human-readable values.
    /// OOXML stores "l", "r", "ctr", "just" etc. — we return "left", "right", "center", "justify".
    /// </summary>
    private static string NormalizeAlignment(string innerText) => innerText switch
    {
        "l" => "left",
        "r" => "right",
        "ctr" => "center",
        "just" => "justify",
        "dist" => "distributed",
        _ => innerText
    };

    /// <summary>
    /// Generate a minimal 1x1 light-gray PNG for use as a zoom placeholder.
    /// PowerPoint regenerates the actual slide thumbnail when the file is opened.
    /// </summary>
    private static byte[] GenerateZoomPlaceholderPng()
    {
        // Minimal valid 1x1 PNG (RGBA: light gray #D0D0D0, fully opaque)
        using var ms = new MemoryStream();
        var bw = new BinaryWriter(ms);

        // PNG signature
        bw.Write(new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A });

        // IHDR chunk: 1x1, 8-bit RGBA
        WriteChunk(bw, "IHDR", new byte[] {
            0, 0, 0, 1, // width = 1
            0, 0, 0, 1, // height = 1
            8,           // bit depth
            6,           // color type = RGBA
            0, 0, 0      // compression, filter, interlace
        });

        // IDAT chunk: zlib-compressed pixel data (filter=0, R=0xD0, G=0xD0, B=0xD0, A=0xFF)
        // Pre-computed deflate of [0x00, 0xD0, 0xD0, 0xD0, 0xFF]
        WriteChunk(bw, "IDAT", new byte[] {
            0x78, 0x01, 0x62, 0x60, 0x60, 0x28, 0x61, 0x28,
            0x61, 0x68, 0xF8, 0x0F, 0x00, 0x01, 0x45, 0x00, 0xC5
        });

        // IEND chunk
        WriteChunk(bw, "IEND", Array.Empty<byte>());

        return ms.ToArray();
    }

    private static void WriteChunk(BinaryWriter bw, string type, byte[] data)
    {
        // Length (big-endian)
        var lenBytes = BitConverter.GetBytes(data.Length);
        if (BitConverter.IsLittleEndian) Array.Reverse(lenBytes);
        bw.Write(lenBytes);

        // Type
        var typeBytes = System.Text.Encoding.ASCII.GetBytes(type);
        bw.Write(typeBytes);

        // Data
        bw.Write(data);

        // CRC32 over type + data
        var crcData = new byte[4 + data.Length];
        Array.Copy(typeBytes, 0, crcData, 0, 4);
        Array.Copy(data, 0, crcData, 4, data.Length);
        var crc = Crc32(crcData);
        var crcBytes = BitConverter.GetBytes(crc);
        if (BitConverter.IsLittleEndian) Array.Reverse(crcBytes);
        bw.Write(crcBytes);
    }

    private static uint Crc32(byte[] data)
    {
        uint crc = 0xFFFFFFFF;
        foreach (var b in data)
        {
            crc ^= b;
            for (int i = 0; i < 8; i++)
                crc = (crc >> 1) ^ (crc & 1) * 0xEDB88320;
        }
        return ~crc;
    }

    /// <summary>
    /// Find all zoom AlternateContent elements in a shape tree.
    /// </summary>
    private static List<OpenXmlElement> GetZoomElements(ShapeTree shapeTree)
    {
        return shapeTree.ChildElements
            .Where(e => e.LocalName == "AlternateContent" &&
                   e.Descendants().Any(d => d.LocalName == "sldZm"))
            .ToList();
    }

    /// <summary>
    /// Find all 3D model AlternateContent elements in a shape tree.
    /// </summary>
    private static List<OpenXmlElement> GetModel3DElements(ShapeTree shapeTree)
    {
        return shapeTree.ChildElements
            .Where(e => e.LocalName == "AlternateContent" &&
                   e.Descendants().Any(d => d.LocalName == "model3d"))
            .ToList();
    }

    /// <summary>
    /// Build a DocumentNode from a 3D model AlternateContent element.
    /// </summary>
    private DocumentNode Model3DToNode(OpenXmlElement acElement, int slideNum, int modelIdx)
    {
        var node = new DocumentNode
        {
            Path = $"/slide[{slideNum}]/model3d[{modelIdx}]",
            Type = "model3d"
        };

        // Navigate: mc:Choice > p:graphicFrame (or p:sp for legacy)
        var choice = acElement.ChildElements.FirstOrDefault(e => e.LocalName == "Choice");
        var gf = choice?.ChildElements.FirstOrDefault(e => e.LocalName == "graphicFrame")
              ?? choice?.ChildElements.FirstOrDefault(e => e.LocalName == "sp");

        // Name from cNvPr
        var nvGfPr = gf?.ChildElements.FirstOrDefault(e => e.LocalName == "nvGraphicFramePr")
                  ?? gf?.ChildElements.FirstOrDefault(e => e.LocalName == "nvSpPr");
        var cNvPr = nvGfPr?.ChildElements.FirstOrDefault(e => e.LocalName == "cNvPr");
        if (cNvPr != null)
        {
            var nameAttr = cNvPr.GetAttribute("name", "");
            if (!string.IsNullOrEmpty(nameAttr.Value))
                node.Format["name"] = nameAttr.Value;
        }

        // Position/size from xfrm (graphicFrame level) or spPr > xfrm
        var xfrm = gf?.ChildElements.FirstOrDefault(e => e.LocalName == "xfrm");
        if (xfrm == null)
        {
            var spPr = gf?.ChildElements.FirstOrDefault(e => e.LocalName == "spPr");
            xfrm = spPr?.ChildElements.FirstOrDefault(e => e.LocalName == "xfrm");
        }
        if (xfrm != null)
        {
            var off = xfrm.ChildElements.FirstOrDefault(e => e.LocalName == "off");
            var ext = xfrm.ChildElements.FirstOrDefault(e => e.LocalName == "ext");
            if (off != null)
            {
                var xAttr = off.GetAttribute("x", "");
                var yAttr = off.GetAttribute("y", "");
                if (!string.IsNullOrEmpty(xAttr.Value) && long.TryParse(xAttr.Value, out var xVal))
                    node.Format["x"] = FormatEmu(xVal);
                if (!string.IsNullOrEmpty(yAttr.Value) && long.TryParse(yAttr.Value, out var yVal))
                    node.Format["y"] = FormatEmu(yVal);
            }
            if (ext != null)
            {
                var cxAttr = ext.GetAttribute("cx", "");
                var cyAttr = ext.GetAttribute("cy", "");
                if (!string.IsNullOrEmpty(cxAttr.Value) && long.TryParse(cxAttr.Value, out var cxVal))
                    node.Format["width"] = FormatEmu(cxVal);
                if (!string.IsNullOrEmpty(cyAttr.Value) && long.TryParse(cyAttr.Value, out var cyVal))
                    node.Format["height"] = FormatEmu(cyVal);
            }
        }

        // Model3D-specific properties
        var model3d = acElement.Descendants().FirstOrDefault(d => d.LocalName == "model3d");
        if (model3d != null)
        {
            // Model rotation
            var rot = model3d.Descendants().FirstOrDefault(d => d.LocalName == "rot");
            if (rot != null)
            {
                var ax = rot.GetAttribute("ax", "").Value ?? "";
                var ay = rot.GetAttribute("ay", "").Value ?? "";
                var az = rot.GetAttribute("az", "").Value ?? "";
                if (!string.IsNullOrEmpty(ax) || !string.IsNullOrEmpty(ay) || !string.IsNullOrEmpty(az))
                {
                    static string ToDeg(string val) =>
                        !string.IsNullOrEmpty(val) && int.TryParse(val, out var v) ? (v / 60000.0).ToString("0.##") : "0";
                    node.Format["rotation"] = $"{ToDeg(ax)},{ToDeg(ay)},{ToDeg(az)}";
                }
            }
        }

        return node;
    }

    /// <summary>
    /// Convert a SlideId value to 1-based slide number.
    /// </summary>
    private int SlideIdToNumber(uint sldId)
    {
        var slideIds = _doc.PresentationPart?.Presentation?.GetFirstChild<SlideIdList>()
            ?.Elements<SlideId>().ToList();
        if (slideIds == null) return -1;
        for (int i = 0; i < slideIds.Count; i++)
            if (slideIds[i].Id?.Value == sldId) return i + 1;
        return -1;
    }

    /// <summary>
    /// Build a DocumentNode from a zoom AlternateContent element.
    /// </summary>
    private DocumentNode ZoomToNode(OpenXmlElement acElement, int slideNum, int zoomIdx)
    {
        var node = new DocumentNode
        {
            Path = $"/slide[{slideNum}]/zoom[{zoomIdx}]",
            Type = "zoom"
        };

        // Navigate: mc:Choice > p:graphicFrame
        var choice = acElement.ChildElements.FirstOrDefault(e => e.LocalName == "Choice");
        var gf = choice?.ChildElements.FirstOrDefault(e => e.LocalName == "graphicFrame");

        // Name from cNvPr
        var nvGfPr = gf?.ChildElements.FirstOrDefault(e => e.LocalName == "nvGraphicFramePr");
        var cNvPr = nvGfPr?.ChildElements.FirstOrDefault(e => e.LocalName == "cNvPr");
        if (cNvPr != null)
        {
            var nameAttr = cNvPr.GetAttribute("name", "");
            if (!string.IsNullOrEmpty(nameAttr.Value))
                node.Format["name"] = nameAttr.Value;
        }

        // Position from xfrm
        var xfrm = gf?.ChildElements.FirstOrDefault(e => e.LocalName == "xfrm");
        if (xfrm != null)
        {
            var off = xfrm.ChildElements.FirstOrDefault(e => e.LocalName == "off");
            var ext = xfrm.ChildElements.FirstOrDefault(e => e.LocalName == "ext");
            if (off != null)
            {
                var xAttr = off.GetAttribute("x", "");
                var yAttr = off.GetAttribute("y", "");
                if (!string.IsNullOrEmpty(xAttr.Value) && long.TryParse(xAttr.Value, out var x))
                    node.Format["x"] = FormatEmu(x);
                if (!string.IsNullOrEmpty(yAttr.Value) && long.TryParse(yAttr.Value, out var y))
                    node.Format["y"] = FormatEmu(y);
            }
            if (ext != null)
            {
                var cxAttr = ext.GetAttribute("cx", "");
                var cyAttr = ext.GetAttribute("cy", "");
                if (!string.IsNullOrEmpty(cxAttr.Value) && long.TryParse(cxAttr.Value, out var cx))
                    node.Format["width"] = FormatEmu(cx);
                if (!string.IsNullOrEmpty(cyAttr.Value) && long.TryParse(cyAttr.Value, out var cy))
                    node.Format["height"] = FormatEmu(cy);
            }
        }

        // Zoom properties from sldZmObj / zmPr
        var sldZmObj = acElement.Descendants().FirstOrDefault(d => d.LocalName == "sldZmObj");
        if (sldZmObj != null)
        {
            var sldIdAttr = sldZmObj.GetAttribute("sldId", "");
            if (!string.IsNullOrEmpty(sldIdAttr.Value) && uint.TryParse(sldIdAttr.Value, out var sldId))
            {
                var targetNum = SlideIdToNumber(sldId);
                if (targetNum > 0) node.Format["target"] = targetNum;
            }
        }

        var zmPr = acElement.Descendants().FirstOrDefault(d => d.LocalName == "zmPr");
        if (zmPr != null)
        {
            var rtpAttr = zmPr.GetAttribute("returnToParent", "");
            if (!string.IsNullOrEmpty(rtpAttr.Value))
                node.Format["returnToParent"] = rtpAttr.Value;
            var tdAttr = zmPr.GetAttribute("transitionDur", "");
            if (!string.IsNullOrEmpty(tdAttr.Value))
                node.Format["transitionDur"] = tdAttr.Value;
        }

        return node;
    }

    /// <summary>
    /// Read a GradientFill element and return a string representation (C1-C2[-angle] or radial:C1-C2[-focus]).
    /// </summary>
    internal static string ReadGradientString(Drawing.GradientFill gradFill)
    {
        var stopEls = gradFill.GradientStopList?.Elements<Drawing.GradientStop>().ToList();
        if (stopEls == null || stopEls.Count == 0) return "gradient";

        var stopData = stopEls.Select(gs => (
            color: ParseHelpers.FormatHexColor(gs.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value ?? "?"),
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
        var deg = linear?.Angle?.HasValue == true ? linear.Angle.Value / 60000.0 : 0.0;
        var degStr = deg % 1 == 0 ? $"{(int)deg}" : $"{deg:0.##}";
        return $"linear;{string.Join(";", stopStrs)};{degStr}";
    }

    /// <summary>
    /// Parse SVG-like path syntax into a Drawing.CustomGeometry element.
    /// Format: "M x,y L x,y C x1,y1 x2,y2 x,y Q x1,y1 x,y Z"
    ///   M = moveTo, L = lineTo, C = cubicBezTo, Q = quadBezTo, A = arcTo, Z = close
    /// Coordinates use 0-100 relative space, internally scaled ×1000 to OOXML standard 0-100000.
    /// Example: "M 0,0 L 100,0 L 100,100 L 0,100 Z" (rectangle in 0-100 space)
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

    /// <summary>
    /// Parse "x,y" coordinate token and scale ×1000 to OOXML standard 0-100000 range.
    /// Input coordinates are 0-100 relative space.
    /// </summary>
    private static (long x, long y) ParsePointToken(string token)
    {
        var parts = token.Split(',');
        if (parts.Length < 2)
            throw new ArgumentException($"Invalid coordinate '{token}'. Expected 'x,y' format (e.g. '100,200').");
        if (!long.TryParse(parts[0].Trim(), out var x))
            throw new ArgumentException($"Invalid x coordinate '{parts[0].Trim()}' in '{token}'. Expected a number.");
        if (!long.TryParse(parts[1].Trim(), out var y))
            throw new ArgumentException($"Invalid y coordinate '{parts[1].Trim()}' in '{token}'. Expected a number.");
        // Scale from user space (0-100) to OOXML standard (0-100000)
        return (x * 1000, y * 1000);
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

    /// <summary>
    /// Apply a position/size property (x, y, width, height) to offset and extents.
    /// Returns true if the key was handled.
    /// </summary>
    private static bool TryApplyPositionSize(string key, string value, Drawing.Offset offset, Drawing.Extents extents)
    {
        var emu = ParseEmu(value);
        switch (key)
        {
            case "x": offset.X = emu; return true;
            case "y": offset.Y = emu; return true;
            case "width":
                if (emu < 0) throw new ArgumentException($"Negative width is not allowed: '{value}'.");
                extents.Cx = emu; return true;
            case "height":
                if (emu < 0) throw new ArgumentException($"Negative height is not allowed: '{value}'.");
                extents.Cy = emu; return true;
            default: return false;
        }
    }
}
