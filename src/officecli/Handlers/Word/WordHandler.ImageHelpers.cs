// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
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

    private static uint _nextImageId = 1;
    private static uint NextImageId() => _nextImageId++;

    private static Run CreateImageRun(string relationshipId, long cx, long cy, string altText)
    {
        var docPropId = NextImageId();
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
        bool behindText)
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

        var anchorDocPropId = NextImageId();
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
        if (extent?.Cx != null) node.Format["width"] = $"{extent.Cx.Value / 360000.0:F1}cm";
        if (extent?.Cy != null) node.Format["height"] = $"{extent.Cy.Value / 360000.0:F1}cm";
        if (docProps?.Description?.Value != null) node.Format["alt"] = docProps.Description.Value;

        return node;
    }
}
