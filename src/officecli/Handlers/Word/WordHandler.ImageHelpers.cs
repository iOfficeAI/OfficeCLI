// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    // ==================== Image Helpers ====================

    private static long ParseEmu(string value) => Core.EmuConverter.ParseEmu(value);

    private uint NextDocPropId()
    {
        uint maxId = 0;
        var body = _doc.MainDocumentPart?.Document?.Body;
        if (body != null)
        {
            foreach (var dp in body.Descendants<DW.DocProperties>())
            {
                if (dp.Id?.HasValue == true && dp.Id.Value > maxId)
                    maxId = dp.Id.Value;
            }
        }
        return maxId + 1;
    }

    private static Run CreateImageRun(string relationshipId, long cx, long cy, string altText, uint docPropId)
    {
        var inline = new DW.Inline(
            new DW.Extent { Cx = cx, Cy = cy },
            new DW.EffectExtent { LeftEdge = 0, TopEdge = 0, RightEdge = 0, BottomEdge = 0 },
            new DW.DocProperties { Id = docPropId, Name = altText, Description = altText },
            new DW.NonVisualGraphicFrameDrawingProperties(
                new A.GraphicFrameLocks { NoChangeAspect = true }
            ),
            new A.Graphic(
                new A.GraphicData(
                    new PIC.Picture(
                        new PIC.NonVisualPictureProperties(
                            new PIC.NonVisualDrawingProperties { Id = docPropId, Name = altText },
                            new PIC.NonVisualPictureDrawingProperties()
                        ),
                        new PIC.BlipFill(
                            new A.Blip { Embed = relationshipId, CompressionState = A.BlipCompressionValues.Print },
                            new A.Stretch(new A.FillRectangle())
                        ),
                        new PIC.ShapeProperties(
                            new A.Transform2D(
                                new A.Offset { X = 0L, Y = 0L },
                                new A.Extents { Cx = cx, Cy = cy }
                            ),
                            new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
                        )
                    )
                ) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
            )
        )
        {
            DistanceFromTop = 0U,
            DistanceFromBottom = 0U,
            DistanceFromLeft = 0U,
            DistanceFromRight = 0U
        };

        return new Run(new Drawing(inline));
    }

    private static Run CreateAnchorImageRun(string relationshipId, long cx, long cy, string altText,
        string wrap, long hPos, long vPos,
        DW.HorizontalRelativePositionValues hRel, DW.VerticalRelativePositionValues vRel,
        bool behindText, uint docPropId)
    {
        OpenXmlElement wrapElement = wrap.ToLowerInvariant() switch
        {
            "square" => new DW.WrapSquare { WrapText = DW.WrapTextValues.BothSides },
            "tight" => new DW.WrapTight(new DW.WrapPolygon(
                new DW.StartPoint { X = 0, Y = 0 },
                new DW.LineTo { X = 21600, Y = 0 },
                new DW.LineTo { X = 21600, Y = 21600 },
                new DW.LineTo { X = 0, Y = 21600 },
                new DW.LineTo { X = 0, Y = 0 }
            ) { Edited = false }),
            "through" => new DW.WrapThrough(new DW.WrapPolygon(
                new DW.StartPoint { X = 0, Y = 0 },
                new DW.LineTo { X = 21600, Y = 0 },
                new DW.LineTo { X = 21600, Y = 21600 },
                new DW.LineTo { X = 0, Y = 21600 },
                new DW.LineTo { X = 0, Y = 0 }
            ) { Edited = false }),
            "topandbottom" or "topbottom" => new DW.WrapTopBottom(),
            "none" => new DW.WrapNone() as OpenXmlElement,
            _ => throw new ArgumentException($"Invalid wrap value: '{wrap}'. Valid values: none, square, tight, through, topandbottom.")
        };

        var anchorDocPropId = docPropId;
        var anchor = new DW.Anchor(
            new DW.SimplePosition { X = 0, Y = 0 },
            new DW.HorizontalPosition(new DW.PositionOffset(hPos.ToString()))
                { RelativeFrom = hRel },
            new DW.VerticalPosition(new DW.PositionOffset(vPos.ToString()))
                { RelativeFrom = vRel },
            new DW.Extent { Cx = cx, Cy = cy },
            new DW.EffectExtent { LeftEdge = 0, TopEdge = 0, RightEdge = 0, BottomEdge = 0 },
            wrapElement,
            new DW.DocProperties { Id = anchorDocPropId, Name = altText, Description = altText },
            new DW.NonVisualGraphicFrameDrawingProperties(
                new A.GraphicFrameLocks { NoChangeAspect = true }),
            new A.Graphic(
                new A.GraphicData(
                    new PIC.Picture(
                        new PIC.NonVisualPictureProperties(
                            new PIC.NonVisualDrawingProperties { Id = anchorDocPropId, Name = altText },
                            new PIC.NonVisualPictureDrawingProperties()),
                        new PIC.BlipFill(
                            new A.Blip { Embed = relationshipId, CompressionState = A.BlipCompressionValues.Print },
                            new A.Stretch(new A.FillRectangle())),
                        new PIC.ShapeProperties(
                            new A.Transform2D(
                                new A.Offset { X = 0L, Y = 0L },
                                new A.Extents { Cx = cx, Cy = cy }),
                            new A.PresetGeometry(new A.AdjustValueList())
                                { Preset = A.ShapeTypeValues.Rectangle })
                    )
                ) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
            )
        )
        {
            BehindDoc = behindText,
            DistanceFromTop = 0U,
            DistanceFromBottom = 0U,
            DistanceFromLeft = 114300U,
            DistanceFromRight = 114300U,
            SimplePos = false,
            RelativeHeight = 1U,
            AllowOverlap = true,
            LayoutInCell = true,
            Locked = false
        };

        return new Run(new Drawing(anchor));
    }

    private static DW.HorizontalRelativePositionValues ParseHorizontalRelative(string value) =>
        value.ToLowerInvariant() switch
        {
            "page" => DW.HorizontalRelativePositionValues.Page,
            "column" => DW.HorizontalRelativePositionValues.Column,
            "character" => DW.HorizontalRelativePositionValues.Character,
            "margin" => DW.HorizontalRelativePositionValues.Margin,
            _ => throw new ArgumentException($"Invalid horizontal relative position: '{value}'. Valid values: margin, page, column, character.")
        };

    private static DW.VerticalRelativePositionValues ParseVerticalRelative(string value) =>
        value.ToLowerInvariant() switch
        {
            "page" => DW.VerticalRelativePositionValues.Page,
            "paragraph" => DW.VerticalRelativePositionValues.Paragraph,
            "line" => DW.VerticalRelativePositionValues.Line,
            "margin" => DW.VerticalRelativePositionValues.Margin,
            _ => throw new ArgumentException($"Invalid vertical relative position: '{value}'. Valid values: margin, page, paragraph, line.")
        };

    private static string GetDrawingInfo(Drawing drawing)
    {
        var docProps = drawing.Descendants<DW.DocProperties>().FirstOrDefault();
        var extent = drawing.Descendants<DW.Extent>().FirstOrDefault();

        var parts = new List<string>();
        if (docProps?.Description?.Value is string desc && !string.IsNullOrEmpty(desc))
            parts.Add($"alt=\"{desc}\"");
        else if (docProps?.Name?.Value is string name && !string.IsNullOrEmpty(name))
            parts.Add($"name=\"{name}\"");
        if (extent != null)
        {
            var wCm = extent.Cx != null ? $"{extent.Cx.Value / 360000.0:F1}cm" : "?";
            var hCm = extent.Cy != null ? $"{extent.Cy.Value / 360000.0:F1}cm" : "?";
            parts.Add($"{wCm}×{hCm}");
        }
        return parts.Count > 0 ? string.Join(", ", parts) : "unknown";
    }

    private static DocumentNode CreateImageNode(Drawing drawing, Run run, string path)
    {
        var docProps = drawing.Descendants<DW.DocProperties>().FirstOrDefault();
        var extent = drawing.Descendants<DW.Extent>().FirstOrDefault();

        var node = new DocumentNode
        {
            Path = path,
            Type = "picture",
            Text = docProps?.Description?.Value ?? docProps?.Name?.Value ?? ""
        };
        if (docProps?.Id?.HasValue == true) node.Format["id"] = docProps.Id.Value;
        if (docProps?.Name?.Value != null) node.Format["name"] = docProps.Name.Value;
        if (extent?.Cx != null) node.Format["width"] = $"{extent.Cx.Value / 360000.0:F1}cm";
        if (extent?.Cy != null) node.Format["height"] = $"{extent.Cy.Value / 360000.0:F1}cm";
        if (docProps?.Description?.Value != null) node.Format["alt"] = docProps.Description.Value;

        // Surface the backing image part rel id so `get --save <path>`
        // and other downstream consumers can locate the payload without
        // re-walking the Drawing tree.
        var imgBlip = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().FirstOrDefault();
        if (imgBlip?.Embed?.Value != null)
            node.Format["relId"] = imgBlip.Embed.Value;

        // Distinguish inline from floating (anchor) and, for anchors, expose
        // the wrap mode, position offsets, and behind-text flag so callers
        // can inspect how the image is laid out.
        var inlineEl = drawing.GetFirstChild<DW.Inline>();
        var anchorEl = drawing.GetFirstChild<DW.Anchor>();
        if (inlineEl != null)
        {
            node.Format["wrap"] = "inline";
        }
        else if (anchorEl != null)
        {
            node.Format["wrap"] = DetectWrapType(anchorEl);
            if (anchorEl.BehindDoc?.Value == true)
                node.Format["behindText"] = true;

            var hPos = anchorEl.GetFirstChild<DW.HorizontalPosition>();
            if (hPos != null)
            {
                var offset = hPos.GetFirstChild<DW.PositionOffset>();
                if (offset != null && long.TryParse(offset.Text, out var hEmu))
                    node.Format["hPosition"] = $"{hEmu / 360000.0:F1}cm";
                if (hPos.RelativeFrom?.HasValue == true)
                    node.Format["hRelative"] = hPos.RelativeFrom.InnerText;
            }

            var vPos = anchorEl.GetFirstChild<DW.VerticalPosition>();
            if (vPos != null)
            {
                var offset = vPos.GetFirstChild<DW.PositionOffset>();
                if (offset != null && long.TryParse(offset.Text, out var vEmu))
                    node.Format["vPosition"] = $"{vEmu / 360000.0:F1}cm";
                if (vPos.RelativeFrom?.HasValue == true)
                    node.Format["vRelative"] = vPos.RelativeFrom.InnerText;
            }
        }

        return node;
    }

    private static string DetectWrapType(DW.Anchor anchor)
    {
        if (anchor.GetFirstChild<DW.WrapNone>() != null) return "none";
        if (anchor.GetFirstChild<DW.WrapSquare>() != null) return "square";
        if (anchor.GetFirstChild<DW.WrapTight>() != null) return "tight";
        if (anchor.GetFirstChild<DW.WrapThrough>() != null) return "through";
        if (anchor.GetFirstChild<DW.WrapTopBottom>() != null) return "topandbottom";
        return "none";
    }

    private static void ReplaceWrapElement(DW.Anchor anchor, string wrapType)
    {
        // Remove any existing wrap element first — at most one is allowed.
        anchor.GetFirstChild<DW.WrapNone>()?.Remove();
        anchor.GetFirstChild<DW.WrapSquare>()?.Remove();
        anchor.GetFirstChild<DW.WrapTight>()?.Remove();
        anchor.GetFirstChild<DW.WrapThrough>()?.Remove();
        anchor.GetFirstChild<DW.WrapTopBottom>()?.Remove();

        OpenXmlElement newWrap = wrapType.ToLowerInvariant() switch
        {
            "square" => new DW.WrapSquare { WrapText = DW.WrapTextValues.BothSides },
            "tight" => new DW.WrapTight(new DW.WrapPolygon(
                new DW.StartPoint { X = 0, Y = 0 },
                new DW.LineTo { X = 21600, Y = 0 },
                new DW.LineTo { X = 21600, Y = 21600 },
                new DW.LineTo { X = 0, Y = 21600 },
                new DW.LineTo { X = 0, Y = 0 }
            ) { Edited = false }),
            "through" => new DW.WrapThrough(new DW.WrapPolygon(
                new DW.StartPoint { X = 0, Y = 0 },
                new DW.LineTo { X = 21600, Y = 0 },
                new DW.LineTo { X = 21600, Y = 21600 },
                new DW.LineTo { X = 0, Y = 21600 },
                new DW.LineTo { X = 0, Y = 0 }
            ) { Edited = false }),
            "topandbottom" or "topbottom" => new DW.WrapTopBottom(),
            "none" => new DW.WrapNone(),
            _ => throw new ArgumentException(
                $"Invalid wrap value: '{wrapType}'. Valid values: none, square, tight, through, topandbottom.")
        };

        // Insert after EffectExtent (standard OOXML child order for
        // CT_Anchor — PowerPoint and Word silently drop wrap elements
        // placed out of schema order).
        var effectExtent = anchor.GetFirstChild<DW.EffectExtent>();
        if (effectExtent != null)
            effectExtent.InsertAfterSelf(newWrap);
        else
            anchor.PrependChild(newWrap);
    }

    /// <summary>
    /// Resolve a run to its top-level Drawing + Anchor, if the run wraps a
    /// floating picture. Used by Set.cs wrap/position cases so the six
    /// wrap/position properties share one lookup instead of each case
    /// re-running the same GetFirstChild chain.
    /// </summary>
    private static DW.Anchor? ResolveRunAnchor(Run run)
    {
        var drawing = run.GetFirstChild<Drawing>();
        return drawing?.GetFirstChild<DW.Anchor>();
    }

    // ==================== OLE Object Reading ====================
    //
    // Embedded OLE objects live inside <w:object> (EmbeddedObject). A VML
    // <v:shape> child carries the display box ("style=width:Xpt;height:Ypt")
    // and an <o:OLEObject> child carries the ProgID. These elements come
    // through as OpenXmlUnknownElement because they are not strongly typed
    // in the core wordprocessing namespace, so we walk descendants by
    // LocalName rather than by CLR type.

    private DocumentNode CreateOleNode(EmbeddedObject oleObj, Run run, string path)
        => CreateOleNode(oleObj, run, path, _doc.MainDocumentPart);

    // BUG-R10-02: OLE inside HeaderPart/FooterPart stores its relationship
    // on the header/footer part itself — not on MainDocumentPart. When we
    // tried to resolve the rel id against MainDocumentPart, GetPartById
    // threw and the node was marked orphan (no contentType/fileSize).
    // Callers in header/footer iteration must pass the enclosing HeaderPart
    // or FooterPart so the lookup succeeds.
    private DocumentNode CreateOleNode(EmbeddedObject oleObj, Run run, string path, OpenXmlPart? hostPart)
    {
        var node = new DocumentNode
        {
            Path = path,
            Type = "ole",
            Text = ""
        };
        node.Format["objectType"] = "ole";

        // ProgID + backing part rel id live on the nested o:OLEObject element.
        // The rel id ("r:id") points to the EmbeddedObjectPart / EmbeddedPackagePart
        // that holds the binary payload — follow it so we can surface content
        // type and byte length in the node, matching how media/image nodes are
        // enriched elsewhere in this handler.
        var oleElement = oleObj.Descendants().FirstOrDefault(e => e.LocalName == "OLEObject");
        string? progId = null;
        string? relId = null;
        string? drawAspect = null;
        if (oleElement != null)
        {
            foreach (var attr in oleElement.GetAttributes())
            {
                if (attr.LocalName == "ProgID")
                    progId = attr.Value;
                else if (attr.LocalName == "DrawAspect")
                    drawAspect = attr.Value;
                else if (attr.LocalName == "id"
                    && attr.NamespaceUri == "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
                    relId = attr.Value;
            }
        }
        // CONSISTENCY(ole-name): PPT OLE Get surfaces oleObj.Name as
        // Format["name"]. Word has no equivalent attribute on o:OLEObject
        // (VML CT_OleObject has no Name), so AddOle/Set store the friendly
        // name on the surrounding v:shape@alt attribute. Read it back from
        // the same place so Add → Get → Format["name"] round-trips.
        var shapeForName = oleObj.Descendants().FirstOrDefault(e => e.LocalName == "shape");
        if (shapeForName != null)
        {
            var altAttr = shapeForName.GetAttributes().FirstOrDefault(a => a.LocalName == "alt");
            if (!string.IsNullOrEmpty(altAttr.Value))
                node.Format["name"] = altAttr.Value;
        }
        // CONSISTENCY(ole-display): PPT OLE Get returns display=icon when the
        // object is shown as an icon; Word stores the same bit in the
        // o:OLEObject DrawAspect attribute ("Icon" vs "Content"). Normalize
        // to the same lowercase "icon"/"content" vocabulary.
        if (!string.IsNullOrEmpty(drawAspect))
        {
            node.Format["display"] = drawAspect.Equals("Content", StringComparison.OrdinalIgnoreCase)
                ? "content"
                : "icon";
        }
        if (!string.IsNullOrEmpty(progId))
        {
            node.Format["progId"] = progId;
            node.Text = progId;
        }
        if (!string.IsNullOrEmpty(relId))
        {
            node.Format["relId"] = relId;
            // GetPartById throws ArgumentOutOfRangeException when the rel id
            // is not present in the part's relationships — this can happen
            // if the document was hand-edited or partially corrupted. Degrade
            // gracefully by marking the node orphan and skipping enrichment,
            // rather than propagating the crash up through Query.
            try
            {
                var part = hostPart?.GetPartById(relId);
                if (part != null)
                    OfficeCli.Core.OleHelper.PopulateFromPart(node, part, progId);
                else
                    node.Format["orphan"] = true;
            }
            catch (ArgumentOutOfRangeException)
            {
                node.Format["orphan"] = true;
            }
            catch (KeyNotFoundException)
            {
                node.Format["orphan"] = true;
            }
        }

        // Display size lives on the VML v:shape element's style string.
        var shape = oleObj.Descendants().FirstOrDefault(e => e.LocalName == "shape");
        if (shape != null)
        {
            var styleAttr = shape.GetAttributes().FirstOrDefault(a => a.LocalName == "style");
            if (!string.IsNullOrEmpty(styleAttr.Value))
                ParseVmlStyle(styleAttr.Value, node);
        }

        return node;
    }

    /// <summary>
    /// Replace a single dimension (width|height) in a VML v:shape style
    /// string, preserving all other key:value pairs. If the key is not
    /// present, it's appended. Output is the re-joined "k1:v1;k2:v2" form.
    /// </summary>
    internal static string ReplaceVmlStyleDimension(string style, string dimKey, string newValue)
    {
        var parts = (style ?? "").Split(';', StringSplitOptions.RemoveEmptyEntries);
        var rebuilt = new List<string>();
        var replaced = false;
        foreach (var part in parts)
        {
            var kv = part.Split(':', 2);
            if (kv.Length == 2 && kv[0].Trim().Equals(dimKey, StringComparison.OrdinalIgnoreCase))
            {
                rebuilt.Add($"{kv[0].Trim()}:{newValue}");
                replaced = true;
            }
            else
            {
                rebuilt.Add(part.Trim());
            }
        }
        if (!replaced) rebuilt.Add($"{dimKey}:{newValue}");
        return string.Join(";", rebuilt);
    }

    private static void ParseVmlStyle(string style, DocumentNode node)
    {
        foreach (var part in style.Split(';', StringSplitOptions.RemoveEmptyEntries))
        {
            var kv = part.Split(':', 2);
            if (kv.Length != 2) continue;
            var k = kv[0].Trim().ToLowerInvariant();
            var v = kv[1].Trim();
            if (k == "width") node.Format["width"] = ConvertVmlLengthToCm(v);
            else if (k == "height") node.Format["height"] = ConvertVmlLengthToCm(v);
        }
    }

    private static readonly System.Text.RegularExpressions.Regex _vmlLengthRegex =
        new(@"^\s*([+-]?\d+(?:\.\d+)?)\s*(pt|in|cm|mm|px)?\s*$",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);

    /// <summary>
    /// Convert a VML length literal (e.g. "385.45pt", "2in", "5cm") into
    /// a "Xcm" string matching the picture width/height format. Uses a
    /// regex to split number from unit so that values containing the
    /// substring "in" (like "line:") inside larger tokens can never be
    /// mangled by naive string.Replace calls.
    /// </summary>
    private static string ConvertVmlLengthToCm(string length)
    {
        var m = _vmlLengthRegex.Match(length);
        if (!m.Success) return length;

        if (!double.TryParse(m.Groups[1].Value,
            System.Globalization.NumberStyles.Float,
            System.Globalization.CultureInfo.InvariantCulture,
            out var value))
            return length;

        var unit = m.Groups[2].Success ? m.Groups[2].Value.ToLowerInvariant() : "pt";
        double cm = unit switch
        {
            "pt" => value * 2.54 / 72.0,
            "in" => value * 2.54,
            "cm" => value,
            "mm" => value / 10.0,
            "px" => value * 2.54 / 96.0,
            _ => value * 2.54 / 72.0,
        };
        return $"{cm:0.##}cm";
    }
}
