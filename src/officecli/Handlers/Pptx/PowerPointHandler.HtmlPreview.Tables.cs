// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler
{
    // ==================== Table Rendering ====================

    private static void RenderTable(StringBuilder sb, GraphicFrame gf, Dictionary<string, string> themeColors, string? dataPath = null)
    {
        var dataPathAttr = string.IsNullOrEmpty(dataPath) ? "" : $" data-path=\"{HtmlEncode(dataPath)}\"";
        var table = gf.Descendants<Drawing.Table>().FirstOrDefault();
        if (table == null) return;

        var offset = gf.Transform?.Offset;
        var extents = gf.Transform?.Extents;
        if (offset == null || extents == null) return;

        var x = offset.X?.Value ?? 0;
        var y = offset.Y?.Value ?? 0;
        var cx = extents.Cx?.Value ?? 0;
        var cy = extents.Cy?.Value ?? 0;

        // Detect table style for style-based coloring
        var tblPr = table.GetFirstChild<Drawing.TableProperties>();
        var tableStyleId = tblPr?.GetFirstChild<Drawing.TableStyleId>()?.InnerText;
        var tableStyleName = tableStyleId != null && _tableStyleGuidToName.TryGetValue(tableStyleId, out var sn) ? sn : null;
        bool hasFirstRow = tblPr?.FirstRow?.Value == true;
        bool hasBandRow = tblPr?.BandRow?.Value == true;

        sb.AppendLine($"    <div class=\"table-container\"{dataPathAttr} style=\"left:{Units.EmuToPt(x)}pt;top:{Units.EmuToPt(y)}pt;width:{Units.EmuToPt(cx)}pt;height:{Units.EmuToPt(cy)}pt\">");
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

        int rowIndex = 0;
        foreach (var row in table.Elements<Drawing.TableRow>())
        {
            sb.AppendLine("        <tr>");
            int skipCols = 0;
            bool isHeaderRow = hasFirstRow && rowIndex == 0;
            bool isBandedOdd = hasBandRow && (!hasFirstRow ? rowIndex % 2 == 0 : rowIndex > 0 && (rowIndex - 1) % 2 == 0);

            foreach (var cell in row.Elements<Drawing.TableCell>())
            {
                var cellStyles = new List<string>();

                // Cell fill
                var tcPr = cell.TableCellProperties ?? cell.GetFirstChild<Drawing.TableCellProperties>();
                var cellSolid = tcPr?.GetFirstChild<Drawing.SolidFill>();
                var cellColor = ResolveFillColor(cellSolid, themeColors);
                bool hasExplicitFill = cellColor != null;
                if (cellColor != null)
                    cellStyles.Add($"background:{cellColor}");

                var cellGrad = tcPr?.GetFirstChild<Drawing.GradientFill>();
                if (cellGrad != null)
                {
                    cellStyles.Add($"background:{GradientToCss(cellGrad, themeColors)}");
                    hasExplicitFill = true;
                }

                // Apply table-style-based colors when no explicit cell fill
                if (!hasExplicitFill && tableStyleName != null)
                {
                    var (bg, fg) = GetTableStyleColors(tableStyleName, isHeaderRow, isBandedOdd, themeColors);
                    if (bg != null) cellStyles.Add($"background:{bg}");
                    if (fg != null) cellStyles.Add($"color:{fg}");
                }

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
                    // else: inherit from table style / slideMaster (no hardcoded default)
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

                // Cell borders (per-edge). When the edge is absent from tcPr,
                // fall back to Office's implicit default: 1pt solid black hairline.
                // An explicit <a:lnL>/<a:lnR>/<a:lnT>/<a:lnB> with <a:noFill/> still
                // yields "none" via TableBorderToCss and is preserved as-is.
                // CONSISTENCY(table-borders): matches the `Npt solid #color` idiom
                // already produced by TableBorderToCss.
                const string defaultBorder = "1pt solid #000000";
                var borderLeft = tcPr?.GetFirstChild<Drawing.LeftBorderLineProperties>();
                var borderRight = tcPr?.GetFirstChild<Drawing.RightBorderLineProperties>();
                var borderTop = tcPr?.GetFirstChild<Drawing.TopBorderLineProperties>();
                var borderBottom = tcPr?.GetFirstChild<Drawing.BottomBorderLineProperties>();
                var bl = TableBorderToCss(borderLeft, themeColors) ?? defaultBorder;
                var br = TableBorderToCss(borderRight, themeColors) ?? defaultBorder;
                var bt = TableBorderToCss(borderTop, themeColors) ?? defaultBorder;
                var bb = TableBorderToCss(borderBottom, themeColors) ?? defaultBorder;
                cellStyles.Add($"border-left:{bl}");
                cellStyles.Add($"border-right:{br}");
                cellStyles.Add($"border-top:{bt}");
                cellStyles.Add($"border-bottom:{bb}");

                // Diagonal borders (<a:lnTlToBr> / <a:lnBlToTr>) — HTML has no
                // native diagonal-border; emit an absolute-positioned inline
                // SVG overlay inside the <td>. The <td> becomes position:relative
                // only when diagonals are actually present to minimize CSS
                // regression surface.
                var borderTlBr = tcPr?.GetFirstChild<Drawing.TopLeftToBottomRightBorderLineProperties>();
                var borderBlTr = tcPr?.GetFirstChild<Drawing.BottomLeftToTopRightBorderLineProperties>();
                var tlBrCss = TableBorderToCss(borderTlBr, themeColors);
                var blTrCss = TableBorderToCss(borderBlTr, themeColors);
                bool hasDiag = (tlBrCss != null && tlBrCss != "none")
                            || (blTrCss != null && blTrCss != "none");
                if (hasDiag)
                    cellStyles.Add("position:relative");

                // Cell margins/padding
                var marL = tcPr?.LeftMargin?.Value;
                var marR = tcPr?.RightMargin?.Value;
                var marT = tcPr?.TopMargin?.Value;
                var marB = tcPr?.BottomMargin?.Value;
                if (marL.HasValue || marR.HasValue || marT.HasValue || marB.HasValue)
                {
                    var pT = Units.EmuToPt(marT ?? 45720);
                    var pR = Units.EmuToPt(marR ?? 91440);
                    var pB = Units.EmuToPt(marB ?? 45720);
                    var pL = Units.EmuToPt(marL ?? 91440);
                    cellStyles.Add($"padding:{pT}pt {pR}pt {pB}pt {pL}pt");
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

                var diagOverlay = "";
                if (hasDiag)
                {
                    var diagLines = new StringBuilder();
                    if (tlBrCss != null && tlBrCss != "none")
                    {
                        var (stroke, widthPt) = ParseBorderCssForSvg(tlBrCss);
                        diagLines.Append($"<line x1=\"0\" y1=\"0\" x2=\"100%\" y2=\"100%\" stroke=\"{stroke}\" stroke-width=\"{widthPt:0.##}\"/>");
                    }
                    if (blTrCss != null && blTrCss != "none")
                    {
                        var (stroke, widthPt) = ParseBorderCssForSvg(blTrCss);
                        diagLines.Append($"<line x1=\"0\" y1=\"100%\" x2=\"100%\" y2=\"0\" stroke=\"{stroke}\" stroke-width=\"{widthPt:0.##}\"/>");
                    }
                    diagOverlay = $"<svg class=\"cell-diag\" width=\"100%\" height=\"100%\" style=\"position:absolute;inset:0;pointer-events:none;overflow:visible\" preserveAspectRatio=\"none\">{diagLines}</svg>";
                }

                sb.AppendLine($"          <td{spanAttrs}{styleStr}>{diagOverlay}{HtmlEncode(cellText)}</td>");
            }
            sb.AppendLine("        </tr>");
            rowIndex++;
        }

        sb.AppendLine("      </table>");
        sb.AppendLine("    </div>");
    }

    /// <summary>
    /// Convert a table cell border line properties element to a CSS border value.
    /// Returns null if the border has NoFill or is absent.
    /// </summary>
    private static string? TableBorderToCss(OpenXmlCompositeElement? borderProps, Dictionary<string, string> themeColors)
    {
        if (borderProps == null) return null;
        if (borderProps.GetFirstChild<Drawing.NoFill>() != null) return "none";

        var solidFill = borderProps.GetFirstChild<Drawing.SolidFill>();
        var color = ResolveFillColor(solidFill, themeColors) ?? "#000000";

        // Width attribute is on the element itself (w attr in EMU)
        double widthPt = 1.0;
        if (borderProps is Drawing.LeftBorderLineProperties lb && lb.Width?.HasValue == true)
            widthPt = lb.Width.Value / 12700.0;
        else if (borderProps is Drawing.RightBorderLineProperties rb && rb.Width?.HasValue == true)
            widthPt = rb.Width.Value / 12700.0;
        else if (borderProps is Drawing.TopBorderLineProperties tb && tb.Width?.HasValue == true)
            widthPt = tb.Width.Value / 12700.0;
        else if (borderProps is Drawing.BottomBorderLineProperties bb && bb.Width?.HasValue == true)
            widthPt = bb.Width.Value / 12700.0;
        else if (borderProps is Drawing.TopLeftToBottomRightBorderLineProperties tlbr && tlbr.Width?.HasValue == true)
            widthPt = tlbr.Width.Value / 12700.0;
        else if (borderProps is Drawing.BottomLeftToTopRightBorderLineProperties bltr && bltr.Width?.HasValue == true)
            widthPt = bltr.Width.Value / 12700.0;

        if (widthPt < 0.5) widthPt = 0.5;

        var dash = borderProps.GetFirstChild<Drawing.PresetDash>();
        var style = "solid";
        if (dash?.Val?.HasValue == true)
        {
            // CONSISTENCY(dash-pattern): map mixed dash-dot patterns to "dashed" (CSS has no native dashDot).
            // Previously fell through to "solid", which silently dropped the dash pattern.
            style = dash.Val.InnerText switch
            {
                "dash" or "lgDash" or "sysDash" => "dashed",
                "dot" or "sysDot" => "dotted",
                "dashDot" or "lgDashDot" or "lgDashDotDot"
                    or "sysDashDot" or "sysDashDotDot" => "dashed",
                _ => "solid"
            };
        }

        return $"{widthPt:0.##}pt {style} {color}";
    }

    /// <summary>
    /// Parse the "Npt style #color" shorthand produced by TableBorderToCss
    /// back into (stroke-color, stroke-width-in-pt) for SVG diagonal lines.
    /// Format is deterministic: "{w:0.##}pt {solid|dashed|dotted} {color}".
    /// </summary>
    private static (string stroke, double widthPt) ParseBorderCssForSvg(string css)
    {
        var parts = css.Split(' ', 3, StringSplitOptions.RemoveEmptyEntries);
        double widthPt = 1.0;
        string stroke = "#000000";
        if (parts.Length >= 1)
        {
            var w = parts[0];
            if (w.EndsWith("pt", StringComparison.OrdinalIgnoreCase))
                w = w[..^2];
            double.TryParse(w, System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out widthPt);
        }
        if (parts.Length >= 3)
            stroke = parts[2];
        return (stroke, widthPt);
    }

    /// <summary>
    /// Returns (background, foreground) CSS colors for a table style based on row position.
    /// Colors are derived from theme colors with lumMod/lumOff transforms matching PowerPoint's
    /// built-in table style definitions (OOXML spec).
    /// </summary>
    private static (string?, string?) GetTableStyleColors(string styleName, bool isHeader, bool isBandedOdd,
        Dictionary<string, string> themeColors)
    {
        // Helper: resolve a theme color key to hex, defaulting if missing
        static string ThemeHex(Dictionary<string, string> tc, string key, string fallback)
            => tc.TryGetValue(key, out var v) ? v : fallback;

        var dk1 = ThemeHex(themeColors, "dk1", OfficeDefaultThemeColors.Dark1);
        var accent1 = ThemeHex(themeColors, "accent1", OfficeDefaultThemeColors.Accent1);

        return styleName switch
        {
            // Medium Style 2: header=dk1 lumMod50% lumOff50%, band1=dk1 lumMod20% lumOff80%, band2=dk1 lumMod10% lumOff90%
            "medium2" => isHeader ? (ApplyLumModOff(dk1, 50000, 50000), (string?)"#FFFFFF")
                       : isBandedOdd ? (ApplyLumModOff(dk1, 20000, 80000), null)
                       : (ApplyLumModOff(dk1, 10000, 90000), null),

            // Medium Style 1: header=dk1, band1=dk1 tint25%, band2=none (uses dk1 base, not accent)
            "medium1" => isHeader ? ((string?)$"#{dk1}", (string?)"#FFFFFF")
                       : isBandedOdd ? (ApplyLumModOff(dk1, 25000, 75000), null)
                       : (null, null),

            // Medium Style 3: header border lines (accent1), band1=accent1 tint20%
            "medium3" => isBandedOdd ? (ApplyLumModOff(accent1, 20000, 80000), null)
                       : (null, null),

            // Medium Style 4: no header fill, band1=dk1 tint15%, band2=dk1 tint5%
            "medium4" => isBandedOdd ? (ApplyLumModOff(dk1, 15000, 85000), null)
                       : (ApplyLumModOff(dk1, 5000, 95000), null),

            // Dark Style 1: header=dk1 (raw), band1=dk1 tint25% (lumMod=25 lumOff=75), band2=dk1 tint15% (lumMod=15 lumOff=85)
            "dark1" => isHeader ? ($"#{dk1}", "#FFFFFF")
                     : isBandedOdd ? (ApplyLumModOff(dk1, 25000, 75000), "#FFFFFF")
                     : (ApplyLumModOff(dk1, 15000, 85000), "#FFFFFF"),

            // Dark Style 2 - Accent 1: header=dk1, band1=accent1 (raw), band2=accent1 lumMod75%
            "dark2" => isHeader ? ($"#{dk1}", "#FFFFFF")
                     : isBandedOdd ? ((string?)$"#{accent1}", "#FFFFFF")
                     : (ApplyLumModOff(accent1, 75000, 0), "#FFFFFF"),

            // Light Style 1: no fill, but banded rows get dk1 tint10%
            "light1" => isBandedOdd ? (ApplyLumModOff(dk1, 10000, 90000), null) : (null, null),
            // Light Style 2/3: band1=accent1 lumMod20% lumOff80%
            "light2" => isBandedOdd ? (ApplyLumModOff(accent1, 20000, 80000), null) : (null, null),
            "light3" => isBandedOdd ? (ApplyLumModOff(accent1, 20000, 80000), null) : (null, null),
            _ => (null, null),
        };
    }

    /// <summary>
    /// Apply OOXML lumMod/lumOff color transform in HSL space.
    /// Delegates to shared ColorMath.ApplyLumModOff.
    /// </summary>
    private static string ApplyLumModOff(string hex, int lumMod, int lumOff)
        => ColorMath.ApplyLumModOff(hex, lumMod, lumOff);
}
