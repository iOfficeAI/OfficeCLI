// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0
//
// Render xlsx shapes (xdr:sp) and textboxes as absolutely-positioned SVG/HTML
// overlays on top of the sheet grid, mirroring how CollectSheetCharts handles
// chart anchors. Ports the preset-geometry SVG logic from WordHandler.
// Pictures (xdr:pic) and graphic frames (charts) are handled elsewhere.

using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Drawing = DocumentFormat.OpenXml.Drawing;
using XDR = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class ExcelHandler
{
    /// <summary>
    /// Pre-render all xdr:sp shapes / textboxes and return them with their
    /// anchor row/col positions (same tuple shape as CollectSheetCharts so the
    /// existing overlay positioning code can consume the result).
    /// </summary>
    private List<(int fromRow, int toRow, int fromCol, int toCol, int widthPtHint, string html)> CollectSheetShapes(WorksheetPart worksheetPart)
    {
        var result = new List<(int fromRow, int toRow, int fromCol, int toCol, int widthPtHint, string html)>();
        var drawingsPart = worksheetPart.DrawingsPart;
        if (drawingsPart?.WorksheetDrawing == null) return result;

        foreach (var anchor in drawingsPart.WorksheetDrawing.ChildElements)
        {
            // Collect shape child (ignore pics, graphicFrames, groupShapes)
            var shape = anchor.Elements<XDR.Shape>().FirstOrDefault();
            if (shape == null) continue;

            int fromRow = 0, toRow = 0, fromCol = 0, toCol = 0;
            if (anchor is XDR.TwoCellAnchor tca)
            {
                int.TryParse(tca.FromMarker?.RowId?.Text, out fromRow);
                int.TryParse(tca.ToMarker?.RowId?.Text, out toRow);
                int.TryParse(tca.FromMarker?.ColumnId?.Text, out fromCol);
                int.TryParse(tca.ToMarker?.ColumnId?.Text, out toCol);
            }
            else if (anchor is XDR.OneCellAnchor oca)
            {
                int.TryParse(oca.FromMarker?.RowId?.Text, out fromRow);
                int.TryParse(oca.FromMarker?.ColumnId?.Text, out fromCol);
                // Approximate to-row/col from ext (EMU) — used only for sizing
                var cx = oca.Extent?.Cx?.Value ?? 0;
                var cy = oca.Extent?.Cy?.Value ?? 0;
                toCol = fromCol + Math.Max(1, (int)(cx / EmuConverter.EmuPerInchF * 8)); // rough
                toRow = fromRow + Math.Max(1, (int)(cy / EmuConverter.EmuPerInchF * 6));
            }
            else
            {
                // AbsoluteAnchor or unsupported — skip
                continue;
            }

            var sb = new StringBuilder();
            RenderShape(sb, shape);
            // 0 = no width hint; the overlay loop falls back to its column-sum.
            result.Add((fromRow, toRow, fromCol, toCol, 0, sb.ToString()));
        }

        return result;
    }

    /// <summary>
    /// Render a single xdr:sp element as an SVG (for preset geometry) plus
    /// optional text body as an overlaid HTML flex-div.
    /// </summary>
    private static void RenderShape(StringBuilder sb, XDR.Shape shape)
    {
        var spPr = shape.ShapeProperties;
        var prstGeom = spPr?.GetFirstChild<Drawing.PresetGeometry>();
        // Preset token — Shape.Preset enum value serializes to the OOXML token
        // (e.g. "rect", "roundRect", "ellipse"). Fall back to "rect".
        var prst = prstGeom?.Preset?.Value.ToString() ?? "rect";

        // Fill
        var fillHex = TryReadSolidFillHex(spPr) ?? "#FFFFFF";
        var hasNoFill = spPr?.GetFirstChild<Drawing.NoFill>() != null;
        if (hasNoFill) fillHex = "transparent";

        // Line/stroke
        var ln = spPr?.GetFirstChild<Drawing.Outline>();
        var strokeHex = ln != null ? (TryReadSolidFillHex(ln) ?? "#000000") : "#000000";
        var strokeWidthPx = 1.0;
        if (ln?.Width?.Value is int lw) strokeWidthPx = Math.Max(0.5, lw / EmuConverter.EmuPerPointF); // EMU→pt≈px
        var hasNoLine = ln?.GetFirstChild<Drawing.NoFill>() != null;

        // Outer div fills the overlay parent.
        sb.Append("<div class=\"xlsx-shape\" style=\"position:absolute;inset:0;display:flex;align-items:center;justify-content:center;overflow:visible\">");

        // Inline SVG overlay for the geometry.
        sb.Append("<svg style=\"position:absolute;inset:0;width:100%;height:100%;overflow:visible\" viewBox=\"0 0 100 100\" preserveAspectRatio=\"none\" xmlns=\"http://www.w3.org/2000/svg\">");
        RenderPrstGeomSvgExcel(sb, prst, fillHex, hasNoLine ? "none" : strokeHex, strokeWidthPx);
        sb.Append("</svg>");

        // Text body overlay as HTML (positioned above SVG via relative stacking)
        var txBody = shape.TextBody;
        if (txBody != null)
        {
            RenderShapeTextBody(sb, txBody);
        }

        sb.Append("</div>");
    }

    /// <summary>
    /// Extract the first solidFill's hex color from the given element (or its
    /// outline child). Returns #-prefixed hex or null.
    /// </summary>
    private static string? TryReadSolidFillHex(OpenXmlElement? el)
    {
        if (el == null) return null;
        var solid = el.GetFirstChild<Drawing.SolidFill>();
        if (solid == null) return null;
        var srgb = solid.GetFirstChild<Drawing.RgbColorModelHex>();
        if (srgb?.Val?.Value is string hex && hex.Length >= 6)
        {
            var v = hex.Length > 6 ? hex[^6..] : hex;
            return "#" + v.ToUpperInvariant();
        }
        var scheme = solid.GetFirstChild<Drawing.SchemeColor>();
        if (scheme?.Val != null)
        {
            // Leave scheme references unresolved here; callers treat null as fallback.
            return null;
        }
        return null;
    }

    /// <summary>
    /// Render a shape's a:txBody as stacked &lt;div&gt; lines centered in the
    /// host container. Honors run-level size/bold/italic/color and paragraph
    /// alignment.
    /// </summary>
    private static void RenderShapeTextBody(StringBuilder sb, XDR.TextBody txBody)
    {
        sb.Append("<div style=\"position:relative;z-index:1;width:100%;padding:4px;text-align:center;pointer-events:none\">");
        foreach (var para in txBody.Elements<Drawing.Paragraph>())
        {
            var pPr = para.GetFirstChild<Drawing.ParagraphProperties>();
            var align = pPr?.Alignment?.Value.ToString() switch
            {
                "ctr" => "center",
                "r" => "right",
                "l" => "left",
                _ => "center"
            };
            sb.Append($"<div style=\"text-align:{align}\">");
            foreach (var run in para.Elements<Drawing.Run>())
            {
                var rPr = run.RunProperties;
                var style = new StringBuilder();
                if (rPr?.FontSize?.Value is int fs) style.Append($"font-size:{fs / 100.0:0.##}pt;");
                if (rPr?.Bold?.Value == true) style.Append("font-weight:bold;");
                if (rPr?.Italic?.Value == true) style.Append("font-style:italic;");
                var colorHex = TryReadSolidFillHex(rPr);
                if (colorHex != null) style.Append($"color:{colorHex};");
                var text = run.Text?.Text ?? "";
                sb.Append($"<span style=\"{style}\">{HtmlEncode(text)}</span>");
            }
            sb.Append("</div>");
        }
        sb.Append("</div>");
    }

    /// <summary>
    /// Emit SVG content for the given preset geometry inside a 0..100 viewBox.
    /// Mirrors WordHandler.RenderPrstGeomSvg with the addition of rect /
    /// roundRect / ellipse / triangle / diamond / parallelogram that xlsx
    /// shapes most commonly use. Unknown presets fall back to a plain rect.
    /// </summary>
    private static void RenderPrstGeomSvgExcel(
        StringBuilder sb, string prst, string fill, string stroke, double strokeW)
    {
        var sw = strokeW.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture);
        var strokeAttrs = stroke == "none"
            ? "stroke=\"none\""
            : $"stroke=\"{stroke}\" stroke-width=\"{sw}\" vector-effect=\"non-scaling-stroke\"";
        switch (prst)
        {
            case "rect":
                sb.Append($"<rect x=\"0\" y=\"0\" width=\"100\" height=\"100\" fill=\"{fill}\" {strokeAttrs}/>");
                break;
            case "roundRect":
                // Default adjustment ~0.1 of shorter side; viewBox is 100 so rx=10.
                sb.Append($"<rect x=\"0\" y=\"0\" width=\"100\" height=\"100\" rx=\"10\" ry=\"10\" fill=\"{fill}\" {strokeAttrs}/>");
                break;
            case "ellipse":
            case "oval":
                sb.Append($"<ellipse cx=\"50\" cy=\"50\" rx=\"50\" ry=\"50\" fill=\"{fill}\" {strokeAttrs}/>");
                break;
            case "triangle":
                sb.Append($"<polygon points=\"50,0 100,100 0,100\" fill=\"{fill}\" {strokeAttrs}/>");
                break;
            case "rtTriangle":
                sb.Append($"<polygon points=\"0,0 0,100 100,100\" fill=\"{fill}\" {strokeAttrs}/>");
                break;
            case "diamond":
                sb.Append($"<polygon points=\"50,0 100,50 50,100 0,50\" fill=\"{fill}\" {strokeAttrs}/>");
                break;
            case "parallelogram":
                sb.Append($"<polygon points=\"20,0 100,0 80,100 0,100\" fill=\"{fill}\" {strokeAttrs}/>");
                break;
            case "trapezoid":
                sb.Append($"<polygon points=\"20,0 80,0 100,100 0,100\" fill=\"{fill}\" {strokeAttrs}/>");
                break;
            case "pentagon":
                sb.Append($"<polygon points=\"50,0 100,38 81,100 19,100 0,38\" fill=\"{fill}\" {strokeAttrs}/>");
                break;
            case "hexagon":
                sb.Append($"<polygon points=\"25,0 75,0 100,50 75,100 25,100 0,50\" fill=\"{fill}\" {strokeAttrs}/>");
                break;
            case "octagon":
                sb.Append($"<polygon points=\"30,0 70,0 100,30 100,70 70,100 30,100 0,70 0,30\" fill=\"{fill}\" {strokeAttrs}/>");
                break;
            case "line":
            case "straightConnector1":
                sb.Append($"<line x1=\"0\" y1=\"0\" x2=\"100\" y2=\"100\" {(stroke == "none" ? $"stroke=\"#000\" stroke-width=\"{sw}\"" : strokeAttrs)}/>");
                break;
            case "rightArrow":
                sb.Append($"<polygon points=\"0,30 70,30 70,10 100,50 70,90 70,70 0,70\" fill=\"{fill}\" {strokeAttrs}/>");
                break;
            case "leftArrow":
                sb.Append($"<polygon points=\"100,30 30,30 30,10 0,50 30,90 30,70 100,70\" fill=\"{fill}\" {strokeAttrs}/>");
                break;
            case "upArrow":
                sb.Append($"<polygon points=\"30,100 70,100 70,30 90,30 50,0 10,30 30,30\" fill=\"{fill}\" {strokeAttrs}/>");
                break;
            case "downArrow":
                sb.Append($"<polygon points=\"30,0 70,0 70,70 90,70 50,100 10,70 30,70\" fill=\"{fill}\" {strokeAttrs}/>");
                break;
            default:
                // Unknown preset — fall back to a plain rect so the shape is at
                // least visible at its anchored position (better than blank).
                sb.Append($"<rect x=\"0\" y=\"0\" width=\"100\" height=\"100\" fill=\"{fill}\" {strokeAttrs}/>");
                break;
        }
    }
}
