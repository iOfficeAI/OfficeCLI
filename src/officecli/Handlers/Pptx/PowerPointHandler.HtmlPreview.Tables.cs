// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler
{
    // ==================== Table Rendering ====================

    private static void RenderTable(StringBuilder sb, GraphicFrame gf, Dictionary<string, string> themeColors)
    {
        var table = gf.Descendants<Drawing.Table>().FirstOrDefault();
        if (table == null) return;

        var offset = gf.Transform?.Offset;
        var extents = gf.Transform?.Extents;
        if (offset == null || extents == null) return;

        var x = offset.X?.Value ?? 0;
        var y = offset.Y?.Value ?? 0;
        var cx = extents.Cx?.Value ?? 0;
        var cy = extents.Cy?.Value ?? 0;

        sb.AppendLine($"    <div class=\"table-container\" style=\"left:{EmuToCm(x)}cm;top:{EmuToCm(y)}cm;width:{EmuToCm(cx)}cm;height:{EmuToCm(cy)}cm\">");
        sb.AppendLine("      <table class=\"slide-table\">");

        // Column widths
        var gridCols = table.TableGrid?.Elements<Drawing.GridColumn>().ToList();
        if (gridCols != null && gridCols.Count > 0)
        {
            sb.Append("        <colgroup>");
            long totalWidth = gridCols.Sum(gc => gc.Width?.Value ?? 0);
            foreach (var gc in gridCols)
            {
                var w = gc.Width?.Value ?? 0;
                var pct = totalWidth > 0 ? (w * 100.0 / totalWidth) : (100.0 / gridCols.Count);
                sb.Append($"<col style=\"width:{pct:0.##}%\">");
            }
            sb.AppendLine("</colgroup>");
        }

        foreach (var row in table.Elements<Drawing.TableRow>())
        {
            sb.AppendLine("        <tr>");
            int skipCols = 0;
            foreach (var cell in row.Elements<Drawing.TableCell>())
            {
                var cellStyles = new List<string>();

                // Cell fill
                var tcPr = cell.TableCellProperties ?? cell.GetFirstChild<Drawing.TableCellProperties>();
                var cellSolid = tcPr?.GetFirstChild<Drawing.SolidFill>();
                var cellColor = ResolveFillColor(cellSolid, themeColors);
                if (cellColor != null)
                    cellStyles.Add($"background:{cellColor}");

                var cellGrad = tcPr?.GetFirstChild<Drawing.GradientFill>();
                if (cellGrad != null)
                    cellStyles.Add($"background:{GradientToCss(cellGrad, themeColors)}");

                // Vertical alignment
                if (tcPr?.Anchor?.HasValue == true)
                {
                    var va = tcPr.Anchor.InnerText switch
                    {
                        "ctr" => "middle",
                        "b" => "bottom",
                        _ => "top"
                    };
                    cellStyles.Add($"vertical-align:{va}");
                }

                // Cell text formatting
                var firstRun = cell.Descendants<Drawing.Run>().FirstOrDefault();
                if (firstRun?.RunProperties != null)
                {
                    var rp = firstRun.RunProperties;
                    if (rp.FontSize?.HasValue == true)
                        cellStyles.Add($"font-size:{rp.FontSize.Value / 100.0:0.##}pt");
                    if (rp.Bold?.Value == true)
                        cellStyles.Add("font-weight:bold");
                    var fontVal = rp.GetFirstChild<Drawing.LatinFont>()?.Typeface?.Value
                        ?? rp.GetFirstChild<Drawing.EastAsianFont>()?.Typeface?.Value;
                    if (fontVal != null && !fontVal.StartsWith("+", StringComparison.Ordinal))
                        cellStyles.Add(CssFontFamilyWithFallback(fontVal));
                    var runColor = ResolveFillColor(rp.GetFirstChild<Drawing.SolidFill>(), themeColors);
                    if (runColor != null)
                        cellStyles.Add($"color:{runColor}");
                }

                // Paragraph alignment
                var firstPara = cell.TextBody?.Elements<Drawing.Paragraph>().FirstOrDefault();
                if (firstPara?.ParagraphProperties?.Alignment?.HasValue == true)
                {
                    var align = firstPara.ParagraphProperties.Alignment.InnerText switch
                    {
                        "ctr" => "center",
                        "r" => "right",
                        "just" => "justify",
                        _ => "left"
                    };
                    cellStyles.Add($"text-align:{align}");
                }

                var cellText = cell.TextBody?.InnerText ?? "";
                var styleStr = cellStyles.Count > 0 ? $" style=\"{string.Join(";", cellStyles)}\"" : "";

                // Column/row span (GridSpan and RowSpan are on the TableCell, not TableCellProperties)
                var gridSpan = cell.GridSpan?.Value;
                var rowSpan = cell.RowSpan?.Value;
                var spanAttrs = "";
                if (gridSpan > 1) spanAttrs += $" colspan=\"{gridSpan}\"";
                if (rowSpan > 1) spanAttrs += $" rowspan=\"{rowSpan}\"";

                // Skip merged continuation cells
                if (cell.HorizontalMerge?.Value == true || cell.VerticalMerge?.Value == true)
                    continue;

                // Skip cells covered by previous gridSpan
                if (skipCols > 0)
                {
                    skipCols--;
                    continue;
                }

                if (gridSpan > 1) skipCols = (int)gridSpan - 1;

                sb.AppendLine($"          <td{spanAttrs}{styleStr}>{HtmlEncode(cellText)}</td>");
            }
            sb.AppendLine("        </tr>");
        }

        sb.AppendLine("      </table>");
        sb.AppendLine("    </div>");
    }
}
