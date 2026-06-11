// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0
//
// Render embedded xlsx pictures (xdr:pic in the worksheet drawing) as
// absolutely-positioned <img> overlays on top of the sheet grid, mirroring
// how CollectSheetCharts / CollectSheetShapes handle their anchors.

using System.Text;
using DocumentFormat.OpenXml.Packaging;
using Drawing = DocumentFormat.OpenXml.Drawing;
using XDR = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class ExcelHandler
{
    /// <summary>
    /// Pre-render all xdr:pic pictures and return them with their anchor
    /// row/col positions (same tuple shape as CollectSheetCharts so the
    /// existing overlay positioning code can consume the result). The image
    /// part is inlined as a data: URI.
    /// </summary>
    private List<(int fromRow, int toRow, int fromCol, int toCol, int widthPtHint, string html)> CollectSheetPictures(WorksheetPart worksheetPart)
    {
        var result = new List<(int fromRow, int toRow, int fromCol, int toCol, int widthPtHint, string html)>();
        var drawingsPart = worksheetPart.DrawingsPart;
        if (drawingsPart?.WorksheetDrawing == null) return result;

        foreach (var anchor in drawingsPart.WorksheetDrawing.ChildElements)
        {
            var pic = anchor.Elements<XDR.Picture>().FirstOrDefault();
            if (pic == null) continue;

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
                var cx = oca.Extent?.Cx?.Value ?? 0;
                var cy = oca.Extent?.Cy?.Value ?? 0;
                toCol = fromCol + Math.Max(1, (int)(cx / EmuConverter.EmuPerInchF * 8));
                toRow = fromRow + Math.Max(1, (int)(cy / EmuConverter.EmuPerInchF * 6));
            }
            else
            {
                continue; // AbsoluteAnchor / unsupported
            }

            var dataUri = TryReadPictureDataUri(drawingsPart, pic);
            if (dataUri == null) continue;

            var sb = new StringBuilder();
            sb.Append($"<img class=\"xlsx-picture\" src=\"{dataUri}\" style=\"position:absolute;inset:0;width:100%;height:100%;object-fit:contain;pointer-events:none\"/>");
            // 0 = no width hint; the overlay loop falls back to its column-sum.
            result.Add((fromRow, toRow, fromCol, toCol, 0, sb.ToString()));
        }

        return result;
    }

    /// <summary>
    /// Resolve the picture's blip embed relationship to a data: URI, or null
    /// if the part can't be read.
    /// </summary>
    private static string? TryReadPictureDataUri(DrawingsPart drawingsPart, XDR.Picture pic)
    {
        var blip = pic.BlipFill?.GetFirstChild<Drawing.Blip>();
        if (blip?.Embed?.HasValue != true) return null;
        try
        {
            var imgPart = drawingsPart.GetPartById(blip.Embed!.Value!);
            using var stream = imgPart.GetStream();
            using var ms = new MemoryStream();
            stream.CopyTo(ms);
            var base64 = Convert.ToBase64String(ms.ToArray());
            var contentType = imgPart.ContentType ?? "image/png";
            return $"data:{contentType};base64,{base64}";
        }
        catch
        {
            return null;
        }
    }
}
