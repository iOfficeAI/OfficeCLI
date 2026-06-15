// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    // ==================== Table Rendering ====================

    private void RenderTableHtml(StringBuilder sb, Table table, string? dataPath = null, int depth = 0)
    {
        // CONSISTENCY(dos-hardening): nested-table recursion has no structural
        // bound; a crafted deeply-nested table would overflow the stack
        // (uncatchable crash) during `view html`. See DocumentLimits.
        DocumentLimits.EnsureDepth(depth);

        // Check table-level borders to determine if this is a borderless layout table
        // First try direct table borders, then fall back to table style borders
        var tblPr = table.GetFirstChild<TableProperties>();
        var tblBorders = tblPr?.TableBorders;
        var styleId = tblPr?.TableStyle?.Val?.Value;
        if (tblBorders == null && styleId != null)
            tblBorders = ResolveTableStyleBorders(styleId);
        bool tableBordersNone = IsTableBorderless(tblBorders);

        // Parse tblLook bitmask for conditional formatting
        var tblLook = ParseTableLook(tblPr);

        // Resolve conditional formatting from table style
        var condFormats = styleId != null ? ResolveTableStyleConditionalFormats(styleId) : null;

        // Check for floating table (tblpPr = text wrapping)
        var tblpPr = tblPr?.GetFirstChild<TablePositionProperties>();
        var tableStyles = new List<string>();
        if (tblpPr != null)
        {
            // #2: Float the table with approximate positioning. Horizontal
            // anchor + tblpX/tblpY translated into float + margin. Coverage
            // is ~40% of Word's 2D flow (horzAnchor=margin + vertAnchor=text);
            // vertAnchor=page/margin would need absolute positioning which
            // doesn't interact with text flow.
            var hAnchor = tblpPr.HorizontalAnchor?.InnerText;
            var vAnchor = tblpPr.VerticalAnchor?.InnerText;
            var tblpX = tblpPr.TablePositionX?.Value ?? 0;
            var tblpY = tblpPr.TablePositionY?.Value ?? 0;
            var xAlign = tblpPr.TablePositionXAlignment?.InnerText;
            var floatDir = xAlign == "right" || (hAnchor == "page" && tblpX > 5000)
                ? "right"
                : xAlign == "left" ? "left" : "left";
            tableStyles.Add($"float:{floatDir}");
            // Margins from text distance (dist…FromText).
            var rightDist = tblpPr.RightFromText?.Value ?? 0;
            var bottomDist = tblpPr.BottomFromText?.Value ?? 0;
            var leftDist = tblpPr.LeftFromText?.Value ?? 0;
            var topDist = tblpPr.TopFromText?.Value ?? 0;
            // Fold tblpX into margin-left (or margin-right for float:right)
            // when the anchor is margin-relative so the column offset shows.
            var horzShiftPt = hAnchor == "margin" ? tblpX / 20.0 : 0;
            if (floatDir == "left")
            {
                var leftMargin = leftDist / 20.0 + horzShiftPt;
                if (leftMargin > 0) tableStyles.Add($"margin-left:{leftMargin:0.#}pt");
                if (rightDist > 0) tableStyles.Add($"margin-right:{rightDist / 20.0:0.#}pt");
            }
            else
            {
                var rightMargin = rightDist / 20.0 + horzShiftPt;
                if (rightMargin > 0) tableStyles.Add($"margin-right:{rightMargin:0.#}pt");
                if (leftDist > 0) tableStyles.Add($"margin-left:{leftDist / 20.0:0.#}pt");
            }
            // Vertical offset: only honor vertAnchor=text (default); other
            // anchors would need absolute positioning, which breaks text
            // flow and is better left to a future pass.
            var vertShiftPt = (vAnchor == null || vAnchor == "text") ? tblpY / 20.0 : 0;
            var topMargin = topDist / 20.0 + vertShiftPt;
            if (topMargin > 0) tableStyles.Add($"margin-top:{topMargin:0.#}pt");
            if (bottomDist > 0) tableStyles.Add($"margin-bottom:{bottomDist / 20.0:0.#}pt");
        }

        // Table horizontal alignment on page (jc = center/right)
        var tblJc = tblPr?.TableJustification?.Val?.InnerText;
        if (tblJc == "center")
            tableStyles.Add("margin-left:auto;margin-right:auto");
        else if (tblJc == "right")
            tableStyles.Add("margin-left:auto;margin-right:0");
        else if (tblpPr == null)
        {
            // Table left indent (w:tblInd, dxa): indent the whole table from
            // the left text margin. Only for left-aligned, non-floating tables
            // (center/right use auto margins; floating handles its own offset).
            // twips -> pt = w / 20. pct/auto types skipped (dxa is the common case).
            var tblInd = tblPr?.TableIndentation;
            if (tblInd?.Type?.InnerText is null or "dxa"
                && tblInd?.Width?.Value is int indW && indW > 0)
                tableStyles.Add($"margin-left:{indW / 20.0:0.#}pt");
        }

        // Apply base table style rPr (font-size, color, alignment) to the <table>
        if (styleId != null)
        {
            var baseStyle = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles
                ?.Elements<Style>().FirstOrDefault(s => s.StyleId?.Value == styleId);
            var baseRPr = baseStyle?.StyleRunProperties;
            if (baseRPr?.FontSize?.Val?.Value is string bsz && int.TryParse(bsz, out var bhp))
                tableStyles.Add($"font-size:{bhp / 2.0:0.##}pt");
            var baseColor = ResolveRunColor(baseRPr?.Color);
            if (baseColor != null) tableStyles.Add($"color:{baseColor}");
            var basePPr = baseStyle?.StyleParagraphProperties;
            if (basePPr?.Justification?.Val?.InnerText is string bjc)
            {
                var align = bjc switch { "center" => "center", "right" => "right", _ => (string?)null };
                if (align != null) tableStyles.Add($"text-align:{align}");
            }
        }

        // Table width: explicit tblW → use it; pct → percentage; otherwise sum gridCol widths
        var tblW = tblPr?.TableWidth;
        var tblWType = tblW?.Type?.InnerText;
        if (tblWType == "dxa" && int.TryParse(tblW!.Width?.Value, out var twW) && twW > 0)
        {
            tableStyles.Add($"width:{twW / 20.0:0.##}pt");
        }
        else if (tblWType == "pct" && int.TryParse(tblW!.Width?.Value, out var pctW) && pctW > 0)
        {
            // pct values are in 1/50th of a percent (5000 = 100%)
            tableStyles.Add($"width:{pctW / 50.0:0.##}%");
        }
        else
        {
            // No explicit tblW or type=auto: use gridCol sum as max-width (Word auto-fit behavior)
            // auto layout tables in Word shrink to content; max-width lets browser do the same
            var isFixed = tblPr?.TableLayout?.Type?.InnerText == "fixed";
            var grid = table.GetFirstChild<TableGrid>();
            var gridCols = grid?.Elements<GridColumn>().ToList();
            if (gridCols != null && gridCols.Count > 0)
            {
                int totalTwips = 0;
                bool allValid = true;
                foreach (var gc in gridCols)
                {
                    if (gc.Width?.Value is string gw && int.TryParse(gw, out var gwVal))
                        totalTwips += gwVal;
                    else
                        allValid = false;
                }
                if (allValid && totalTwips > 0)
                {
                    var prop = isFixed ? "width" : "max-width";
                    tableStyles.Add($"{prop}:{totalTwips / 20.0:0.##}pt");
                }
            }
            // else: no grid info — browser auto-fits to content
        }

        // tblCellSpacing (w:tblCellSpacing w:w=twips w:type=dxa): Word draws each
        // cell as a separate box with gaps between them. The global table CSS uses
        // border-collapse:collapse (no gaps); override to separate + border-spacing
        // only for tables that actually declare cell spacing.
        var tblCellSpacing = tblPr?.TableCellSpacing;
        if (tblCellSpacing?.Type?.InnerText is null or "dxa"
            && int.TryParse(tblCellSpacing?.Width?.Value, out var csTwips) && csTwips > 0)
        {
            tableStyles.Add("border-collapse:separate");
            tableStyles.Add($"border-spacing:{csTwips / 20.0:0.##}pt");
        }

        var tableClass = tableBordersNone ? "borderless" : "";

        // Tracked table format change (tblPrChange) — yellow highlight with author attribution
        var tblPrChange = tblPr?.GetFirstChild<TablePropertiesChange>();
        var tblFmtAuthorAttr = "";
        if (tblPrChange != null)
        {
            tableClass = string.IsNullOrEmpty(tableClass) ? "track-format" : tableClass + " track-format";
            tableStyles.Add("background:#FFF9C4;border-left:4px solid #FFC107");
            var tblAuthor = tblPrChange.Author?.Value ?? "";
            if (!string.IsNullOrEmpty(tblAuthor))
                tblFmtAuthorAttr = $" title=\"Format changed by {HtmlEncodeAttr(tblAuthor)}\"";
        }

        var tableStyleAttr = tableStyles.Count > 0 ? $" style=\"{string.Join(";", tableStyles)}\"" : "";
        var dataPathAttr = !string.IsNullOrEmpty(dataPath) ? $" data-path=\"{dataPath}\"" : "";
        if (!string.IsNullOrEmpty(tableClass))
            sb.AppendLine($"<table class=\"{tableClass}\"{dataPathAttr}{tblFmtAuthorAttr}{tableStyleAttr}>");
        else
            sb.AppendLine($"<table{dataPathAttr}{tblFmtAuthorAttr}{tableStyleAttr}>");

        // Get column widths from grid
        // tblLayout=fixed → use fixed col widths; auto/missing → let browser auto-fit by content
        var isFixedLayout = tblPr?.TableLayout?.Type?.InnerText == "fixed";
        var tblGrid = table.GetFirstChild<TableGrid>();
        if (tblGrid != null)
        {
            sb.Append("<colgroup>");
            // BUG-R1-P3-13: autofit tables previously emitted bare <col> with
            // no width hint, dropping the proportions encoded in tblGrid.
            // Now emit proportional column widths (% of total) for autofit
            // *as well as* fixed pt widths for fixed-layout tables. Browser
            // honours pct in autofit mode without overriding content sizing.
            var twipsByCol = tblGrid.Elements<GridColumn>()
                .Select(c => double.TryParse(c.Width?.Value, System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture, out var v) ? v : 0.0)
                .ToList();
            double colTotal = twipsByCol.Sum();
            int colCount = twipsByCol.Count;

            // Per-cell pct widths (w:tcW type="pct") are authoritative over the
            // tblGrid (Word stores equal gridCols even when cells carry explicit
            // percentages, e.g. a 30/40/30 table whose only width-bearing row is
            // not row 1). Scan every row and, per column, capture the first
            // explicit pct/dxa tcW so the colgroup reflects the real proportions
            // instead of equal gridCol distribution. dxa wins over pct only if
            // pct is absent for that column. Columns with no explicit cell width
            // fall back to the gridCol-derived value below.
            var pctByCol = new double?[colCount];
            foreach (var r in table.Elements<TableRow>())
            {
                int ci = 0;
                foreach (var tc in r.Elements<TableCell>())
                {
                    if (ci >= colCount) break;
                    var tcW = tc.TableCellProperties?.TableCellWidth;
                    if (tcW?.Type?.InnerText == "pct" && pctByCol[ci] == null
                        && int.TryParse(tcW.Width?.Value, out var pctVal) && pctVal > 0)
                    {
                        // pct units are 1/50th of a percent (5000 = 100%)
                        pctByCol[ci] = pctVal / 50.0;
                    }
                    // gridSpan-aware advance so column index stays aligned
                    var span = tc.TableCellProperties?.GridSpan?.Val?.Value ?? 1;
                    ci += span < 1 ? 1 : span;
                }
            }

            int colIdx = 0;
            foreach (var col in tblGrid.Elements<GridColumn>())
            {
                var w = col.Width?.Value;
                if (colIdx < colCount && pctByCol[colIdx] is double explicitPct)
                {
                    // Explicit per-cell percentage drives the column width; this
                    // overrides both fixed-pt and gridCol-proportion paths so the
                    // browser renders the authored 30/40/30-style proportions.
                    var twipsAttr = w != null ? $" data-col-twips=\"{w}\"" : "";
                    sb.Append($"<col style=\"width:{explicitPct:0.##}%\"{twipsAttr}>");
                }
                else if (w != null && isFixedLayout)
                {
                    var pt = double.Parse(w, System.Globalization.CultureInfo.InvariantCulture) / 20.0; // twips to pt
                    sb.Append($"<col style=\"width:{pt:0.##}pt\" data-col-twips=\"{w}\">");
                }
                else if (w != null && colTotal > 0 && twipsByCol[colIdx] > 0)
                {
                    // Autofit: emit percentage so the browser respects gridCol
                    // proportions while still allowing content to expand cells.
                    // The raw twip count is also exposed via data-col-twips for
                    // round-trip / verification tooling.
                    var pct = twipsByCol[colIdx] / colTotal * 100.0;
                    sb.Append($"<col style=\"width:{pct:0.##}%;--col-twips:{w}\" data-col-twips=\"{w}\">");
                }
                else
                {
                    sb.Append("<col>");
                }
                colIdx++;
            }
            sb.AppendLine("</colgroup>");
        }

        var rows = table.Elements<TableRow>().ToList();
        // Also collect rows from revision wrappers (ins/del/moveTo/moveFrom)
        var revisionRowFlags = new Dictionary<TableRow, (string? revisionType, string? author)>();
        foreach (var child in table.ChildElements)
        {
            if (child.LocalName is "ins" or "del" or "moveTo" or "moveFrom")
            {
                var revType = child.LocalName;
                var revAuthor = child.GetAttributes().FirstOrDefault(a => a.LocalName == "author").Value;
                foreach (var nestedRow in child.Descendants<TableRow>())
                {
                    if (!rows.Contains(nestedRow))
                    {
                        rows.Add(nestedRow);
                        revisionRowFlags[nestedRow] = (revType, revAuthor);
                    }
                }
            }
        }
        var totalRows = rows.Count;
        var totalCols = tblGrid?.Elements<GridColumn>().Count() ?? rows.FirstOrDefault()?.Elements<TableCell>().Count() ?? 0;

        for (int rowIdx = 0; rowIdx < totalRows; rowIdx++)
        {
            var row = rows[rowIdx];
            var isHeader = row.TableRowProperties?.GetFirstChild<TableHeader>() != null;
            // Row height. trHeight has hRule = auto / atLeast / exact. CSS treats
            // tr.height as min-height (atLeast semantics), so for hRule="exact"
            // we additionally constrain the cell with max-height + overflow:hidden
            // to match Word's content-clipping behavior.
            var trHeight = row.TableRowProperties?.GetFirstChild<TableRowHeight>();
            var trStyle = "";
            double? exactRowHeightPt = null;
            if (trHeight?.Val?.Value is uint hVal && hVal > 0)
            {
                var heightPt = hVal / 20.0;
                trStyle = $" style=\"height:{heightPt:0.#}pt\"";
                if (trHeight.HeightType?.Value == HeightRuleValues.Exact)
                    exactRowHeightPt = heightPt;
            }
            // #7b00: mark tblHeader rows so the JS paginator can clone them
            // onto every continuation page when a long table spans pages.
            var hdrMarker = isHeader ? " data-tbl-header=\"1\"" : "";
            // Row data-path for goto/mark navigation. Skipped for nested tables
            // (dataPath is only set for top-level tables — see RenderTableHtml
            // call sites in HtmlPreview.cs:1906) because nested tables don't
            // have a stable /body/table[N] index.
            var rowDataPath = !string.IsNullOrEmpty(dataPath) ? $"{dataPath}/tr[{rowIdx + 1}]" : null;
            var rowDataPathAttr = rowDataPath != null ? $" data-path=\"{rowDataPath}\"" : "";
            // Tracked row revision (wrapped in ins/del/moveTo/moveFrom) — visual marker
            var rowRevAttr = "";
            var rowRevSuffix = "";
            if (revisionRowFlags.TryGetValue(row, out var revInfo))
            {
                var (revType, revAuthor) = revInfo;
                if (revType == "ins" || revType == "moveTo")
                {
                    rowRevAttr = string.IsNullOrEmpty(revAuthor)
                        ? " title=\"Inserted row\""
                        : $" title=\"Inserted row by {HtmlEncodeAttr(revAuthor)}\"";
                    rowRevSuffix = " track-ins-row";
                }
                else // del or moveFrom
                {
                    rowRevAttr = string.IsNullOrEmpty(revAuthor)
                        ? " title=\"Deleted row\""
                        : $" title=\"Deleted row by {HtmlEncodeAttr(revAuthor)}\"";
                    rowRevSuffix = " track-del-row";
                }
            }
            sb.AppendLine(isHeader
                ? $"<tr class=\"header-row{rowRevSuffix}\"{hdrMarker}{rowRevAttr}{rowDataPathAttr}{trStyle}>"
                : string.IsNullOrEmpty(rowRevSuffix)
                    ? $"<tr{rowRevAttr}{rowDataPathAttr}{trStyle}>"
                    : $"<tr class=\"{rowRevSuffix.Trim()}\"{rowRevAttr}{rowDataPathAttr}{trStyle}>");

            int colIdx = 0;
            foreach (var cell in row.Elements<TableCell>())
            {
                var tag = isHeader ? "th" : "td";
                var condTypes = GetConditionalTypes(tblLook, rowIdx, colIdx, totalRows, totalCols);
                var cellStyle = GetTableCellInlineCss(cell, tableBordersNone, tblBorders, condFormats, condTypes,
                    rowIdx, colIdx, totalRows, totalCols, exactRowHeightPt);

                // Check if conditional format overrides font-size (needs class for CSS override)
                bool hasTsf = cellStyle.Contains("__TSF__");
                cellStyle = cellStyle.Replace(";__TSF__", "").Replace("__TSF__", "");

                // Merge attributes
                var attrs = new StringBuilder();
                if (hasTsf) attrs.Append(" class=\"tsf\"");
                var gridSpan = cell.TableCellProperties?.GridSpan?.Val?.Value;
                if (gridSpan > 1) attrs.Append($" colspan=\"{gridSpan}\"");

                var vMerge = cell.TableCellProperties?.VerticalMerge;
                if (vMerge != null && vMerge.Val?.Value == MergedCellValues.Restart)
                {
                    // Count rowspan
                    var rowspan = CountRowSpan(table, row, cell);
                    if (rowspan > 1) attrs.Append($" rowspan=\"{rowspan}\"");
                }
                else if (vMerge != null && (vMerge.Val == null || vMerge.Val.Value == MergedCellValues.Continue))
                {
                    colIdx += gridSpan ?? 1;
                    continue; // Skip merged continuation cells
                }

                if (!string.IsNullOrEmpty(cellStyle))
                    attrs.Append($" style=\"{cellStyle}\"");

                // Cell data-path uses the OOXML positional cell index (colIdx+1)
                // rather than the visual grid column, to match the handler's
                // /body/table[N]/tr[R]/tc[C] addressing.
                if (rowDataPath != null)
                    attrs.Append($" data-path=\"{rowDataPath}/tc[{colIdx + 1}]\"");

                sb.Append($"<{tag}{attrs}>");

                // Diagonal cell borders (w:tl2br / w:tr2bl) — emit the SVG
                // overlay as the first child of the cell so it paints over the
                // content. The <td> already carries position:relative (added in
                // GetTableCellInlineCss when a diagonal is present). Mirrors the
                // Excel/PPTX cell-diag idiom.
                var diagSvg = TryBuildCellDiagonalSvg(cell);
                if (diagSvg != null) sb.Append(diagSvg);

                // hRule="exact": browsers ignore max-height on <td> (table layout
                // forces cells to contain their content), so wrap content in an
                // inner div with fixed height + overflow:hidden. The wrap also
                // takes over vertical alignment via flex (the td's vertical-align
                // applies to the wrap as a whole, not to content within it).
                bool exactWrap = exactRowHeightPt.HasValue;
                if (exactWrap)
                {
                    var vAlign = cell.TableCellProperties?.TableCellVerticalAlignment?.Val?.Value;
                    string justify;
                    if (vAlign == TableVerticalAlignmentValues.Center) justify = "center";
                    else if (vAlign == TableVerticalAlignmentValues.Bottom) justify = "flex-end";
                    else justify = "flex-start";
                    sb.Append($"<div style=\"height:{exactRowHeightPt:0.#}pt;max-height:{exactRowHeightPt:0.#}pt;overflow:hidden;display:flex;flex-direction:column;justify-content:{justify}\">");
                }

                // Render cell content in XML order. OOXML lets paragraphs and
                // nested tables interleave in a cell (typically: <w:tbl> then
                // a trailing <w:p/> — required by spec for cells ending with a
                // table). Iterating Paragraphs first then Tables would push the
                // trailing empty paragraph above the nested table, displacing
                // it ~one line down. Walk ChildElements directly to preserve
                // document order. Every paragraph (including empty) goes
                // through the same path as body paragraphs: <div> wrapper with
                // inline pPr CSS plus an &nbsp; placeholder for empties so the
                // line box forms and renders the resolved line-height.
                // List grouping inside the cell mirrors the body path: a run of
                // ListBullet/numbered paragraphs becomes <ul>/<ol> with <li>
                // children (single-level — the common in-cell case) instead of
                // plain <div>s, so bullets/numbers render. A non-list paragraph
                // or a nested table closes the open list.
                string? cellListTag = null; // "ul" | "ol" when a list is open
                void CloseCellList()
                {
                    if (cellListTag != null) { sb.Append($"</{cellListTag}>"); cellListTag = null; }
                }

                foreach (var child in cell.ChildElements)
                {
                    if (child is Paragraph cellPara)
                    {
                        var listStyle = GetParagraphListStyle(cellPara);
                        if (listStyle != null)
                        {
                            RenderCellListItem(sb, cellPara, listStyle, ref cellListTag);
                            continue;
                        }
                        CloseCellList();
                        var text = GetParagraphText(cellPara);
                        var runs = GetAllRuns(cellPara);
                        var pCss = GetParagraphInlineCss(cellPara);
                        sb.Append("<div");
                        if (!string.IsNullOrEmpty(pCss))
                            sb.Append($" style=\"{pCss}\"");
                        sb.Append(">");
                        bool hasVisibleContent = runs.Count > 0 || !string.IsNullOrWhiteSpace(text);
                        RenderParagraphContentHtml(sb, cellPara);
                        if (!hasVisibleContent) sb.Append("&nbsp;");
                        sb.Append("</div>");
                    }
                    else if (child is Table nestedTable)
                    {
                        CloseCellList();
                        RenderTableHtml(sb, nestedTable, depth: depth + 1);
                    }
                }
                CloseCellList();

                if (exactWrap) sb.Append("</div>");
                sb.AppendLine($"</{tag}>");
                colIdx += gridSpan ?? 1;
            }

            sb.AppendLine("</tr>");
        }

        sb.AppendLine("</table>");
    }

    private static bool IsTableBorderless(TableBorders? borders)
    {
        if (borders == null) return false;
        // Check if all borders are none/nil
        return IsBorderNone(borders.TopBorder)
            && IsBorderNone(borders.BottomBorder)
            && IsBorderNone(borders.LeftBorder)
            && IsBorderNone(borders.RightBorder)
            && IsBorderNone(borders.InsideHorizontalBorder)
            && IsBorderNone(borders.InsideVerticalBorder);
    }

    private static bool IsBorderNone(OpenXmlElement? border)
    {
        if (border == null) return true;
        var val = border.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
        return val is null or "nil" or "none";
    }

    /// <summary>Apply or clear a conditional format border edge.</summary>
    private void ApplyCondBorder(List<string> parts, OpenXmlElement? border, string cssProperty)
    {
        if (border == null) return;
        parts.RemoveAll(p => p.StartsWith(cssProperty + ":"));
        if (!IsBorderNone(border))
            RenderBorderCss(parts, border, cssProperty);
        // If val=nil/none, the RemoveAll already cleared it — border is removed
    }

    /// <summary>Resolve TableBorders from a table style (walking basedOn chain).</summary>
    private TableBorders? ResolveTableStyleBorders(string styleId)
    {
        var visited = new HashSet<string>();
        var currentId = styleId;
        while (currentId != null && visited.Add(currentId))
        {
            var style = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles
                ?.Elements<Style>().FirstOrDefault(s => s.StyleId?.Value == currentId);
            if (style == null) break;
            var borders = style.StyleTableProperties?.TableBorders;
            if (borders != null) return borders;
            currentId = style.BasedOn?.Val?.Value;
        }
        return null;
    }

    // ==================== Table Look / Conditional Formatting ====================

    [Flags]
    private enum TableLookFlags
    {
        None = 0,
        FirstRow = 0x0020,
        LastRow = 0x0040,
        FirstColumn = 0x0080,
        LastColumn = 0x0100,
        NoHBand = 0x0200,
        NoVBand = 0x0400,
    }

    /// <summary>Parse tblLook from table properties. Start from the legacy
    /// val hex bitmask (if present) and let each authored individual attr
    /// override only the bit it names — per ECMA-376 §17.7.6.7, individual
    /// attrs are independent overrides of val, not a full replacement.</summary>
    private static TableLookFlags ParseTableLook(TableProperties? tblPr)
    {
        var tblLook = tblPr?.GetFirstChild<TableLook>();
        if (tblLook == null) return TableLookFlags.None;

        var flags = TableLookFlags.None;
        var val = tblLook.Val?.Value;
        if (val != null && int.TryParse(val, System.Globalization.NumberStyles.HexNumber, null, out var hex))
            flags = (TableLookFlags)hex;

        // Each authored attr (regardless of true/false) overrides its bit.
        if (tblLook.FirstRow != null)
            flags = tblLook.FirstRow.Value == true ? flags | TableLookFlags.FirstRow : flags & ~TableLookFlags.FirstRow;
        if (tblLook.LastRow != null)
            flags = tblLook.LastRow.Value == true ? flags | TableLookFlags.LastRow : flags & ~TableLookFlags.LastRow;
        if (tblLook.FirstColumn != null)
            flags = tblLook.FirstColumn.Value == true ? flags | TableLookFlags.FirstColumn : flags & ~TableLookFlags.FirstColumn;
        if (tblLook.LastColumn != null)
            flags = tblLook.LastColumn.Value == true ? flags | TableLookFlags.LastColumn : flags & ~TableLookFlags.LastColumn;
        if (tblLook.NoHorizontalBand != null)
            flags = tblLook.NoHorizontalBand.Value == true ? flags | TableLookFlags.NoHBand : flags & ~TableLookFlags.NoHBand;
        if (tblLook.NoVerticalBand != null)
            flags = tblLook.NoVerticalBand.Value == true ? flags | TableLookFlags.NoVBand : flags & ~TableLookFlags.NoVBand;

        return flags;
    }

    /// <summary>Cached conditional format data from a table style.</summary>
    private class TableConditionalFormat
    {
        public Shading? Shading { get; set; }
        public TableCellBorders? Borders { get; set; }
        public RunPropertiesBaseStyle? RunProperties { get; set; }
    }

    /// <summary>Resolve all tblStylePr conditional formatting from a table style (walking basedOn chain).</summary>
    private Dictionary<string, TableConditionalFormat>? ResolveTableStyleConditionalFormats(string styleId)
    {
        var result = new Dictionary<string, TableConditionalFormat>(StringComparer.OrdinalIgnoreCase);
        var visited = new HashSet<string>();
        var currentId = styleId;

        // Walk basedOn chain, collecting conditional formats (child style overrides parent)
        var chainStyles = new List<Style>();
        while (currentId != null && visited.Add(currentId))
        {
            var style = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles
                ?.Elements<Style>().FirstOrDefault(s => s.StyleId?.Value == currentId);
            if (style == null) break;
            chainStyles.Add(style);
            currentId = style.BasedOn?.Val?.Value;
        }

        // Process in reverse (base first, derived last — derived wins)
        chainStyles.Reverse();
        foreach (var style in chainStyles)
        {
            foreach (var tsp in style.Elements<TableStyleProperties>())
            {
                var type = tsp.Type;
                if (type == null) continue;
                // Use the XML serialized value (e.g. "firstRow", "band1Horz") for consistent lookup
                var typeName = type.InnerText;

                var fmt = new TableConditionalFormat();
                // Try SDK-typed property first, then fall back to generic child lookup
                var tcPr = tsp.GetFirstChild<TableStyleConditionalFormattingTableCellProperties>();
                if (tcPr != null)
                {
                    fmt.Shading = tcPr.GetFirstChild<Shading>();
                    fmt.Borders = tcPr.GetFirstChild<TableCellBorders>();
                }
                fmt.RunProperties = tsp.GetFirstChild<RunPropertiesBaseStyle>();

                if (typeName != null)
                    result[typeName] = fmt;
            }
        }

        return result.Count > 0 ? result : null;
    }

    /// <summary>Get the list of conditional format type names that apply to a cell at the given position.</summary>
    private static List<string> GetConditionalTypes(TableLookFlags look, int rowIdx, int colIdx, int totalRows, int totalCols)
    {
        var types = new List<string>();

        // Banded rows (applied first, lowest priority)
        if ((look & TableLookFlags.NoHBand) == 0)
        {
            // Banding skips first/last row if those flags are set
            int bandRowIdx = rowIdx;
            if ((look & TableLookFlags.FirstRow) != 0 && rowIdx > 0) bandRowIdx = rowIdx - 1;
            else if ((look & TableLookFlags.FirstRow) != 0 && rowIdx == 0) bandRowIdx = -1; // first row, skip banding

            if (bandRowIdx >= 0)
                types.Add(bandRowIdx % 2 == 0 ? "band1Horz" : "band2Horz");
        }

        // Banded columns
        if ((look & TableLookFlags.NoVBand) == 0)
        {
            int bandColIdx = colIdx;
            if ((look & TableLookFlags.FirstColumn) != 0 && colIdx > 0) bandColIdx = colIdx - 1;
            else if ((look & TableLookFlags.FirstColumn) != 0 && colIdx == 0) bandColIdx = -1;

            if (bandColIdx >= 0)
                types.Add(bandColIdx % 2 == 0 ? "band1Vert" : "band2Vert");
        }

        // First/last column (higher priority than banding)
        if ((look & TableLookFlags.FirstColumn) != 0 && colIdx == 0)
            types.Add("firstCol");
        if ((look & TableLookFlags.LastColumn) != 0 && colIdx == totalCols - 1)
            types.Add("lastCol");

        // First/last row (highest priority)
        if ((look & TableLookFlags.FirstRow) != 0 && rowIdx == 0)
            types.Add("firstRow");
        if ((look & TableLookFlags.LastRow) != 0 && rowIdx == totalRows - 1)
            types.Add("lastRow");

        return types;
    }

    /// <summary>Calculate the grid column index for a cell, accounting for gridSpan in preceding cells.</summary>
    private static int GetGridColumn(TableRow row, TableCell cell)
    {
        int gridCol = 0;
        foreach (var c in row.Elements<TableCell>())
        {
            if (c == cell) return gridCol;
            gridCol += c.TableCellProperties?.GridSpan?.Val?.Value ?? 1;
        }
        return gridCol;
    }

    /// <summary>Find the cell at a given grid column in a row, accounting for gridSpan.</summary>
    private static TableCell? GetCellAtGridColumn(TableRow row, int targetGridCol)
    {
        int gridCol = 0;
        foreach (var cell in row.Elements<TableCell>())
        {
            if (gridCol == targetGridCol) return cell;
            gridCol += cell.TableCellProperties?.GridSpan?.Val?.Value ?? 1;
            if (gridCol > targetGridCol) return null; // target is inside a spanned cell
        }
        return null;
    }

    /// <summary>
    /// Get all TableRow elements from a table, including those nested inside
    /// revision wrappers (w:ins, w:del, w:moveTo, w:moveFrom).
    /// </summary>
    private static List<TableRow> GetFlattenedTableRows(Table table)
    {
        var rows = table.Elements<TableRow>().ToList();
        foreach (var child in table.ChildElements)
        {
            if (child.LocalName is "ins" or "del" or "moveTo" or "moveFrom")
            {
                foreach (var nestedRow in child.Descendants<TableRow>())
                {
                    if (!rows.Contains(nestedRow))
                        rows.Add(nestedRow);
                }
            }
        }
        return rows;
    }

    private static int CountRowSpan(Table table, TableRow startRow, TableCell startCell)
    {
        var rows = GetFlattenedTableRows(table);
        var startRowIdx = rows.IndexOf(startRow);
        if (startRowIdx < 0) return 1;

        // Use grid column position instead of cell index
        var gridCol = GetGridColumn(startRow, startCell);

        int span = 1;
        for (int i = startRowIdx + 1; i < rows.Count; i++)
        {
            var cell = GetCellAtGridColumn(rows[i], gridCol);
            if (cell == null) break;

            var vm = cell.TableCellProperties?.VerticalMerge;
            if (vm != null && (vm.Val == null || vm.Val.Value == MergedCellValues.Continue))
                span++;
            else
                break;
        }
        return span;
    }

    /// <summary>
    /// Render one list paragraph inside a table cell as an &lt;li&gt;, opening
    /// the &lt;ul&gt;/&lt;ol&gt; when needed. Single-level only — the common
    /// in-cell case — but uses the same marker classes / ordered-marker spans
    /// as the body path so bullets and numbers render identically. Multi-level
    /// nesting inside a cell collapses to one level (a known simplification;
    /// body-level lists remain the full-fidelity path).
    /// </summary>
    private void RenderCellListItem(StringBuilder sb, Paragraph para, string listStyle, ref string? cellListTag)
    {
        var resolvedNumPr = ResolveNumPrFromStyle(para);
        var ilvl = resolvedNumPr?.Ilvl ?? 0;
        var numId = resolvedNumPr?.NumId ?? 0;
        if (ilvl < 0) ilvl = 0; else if (ilvl > 8) ilvl = 8;
        var lvlText = GetLevelText(numId, ilvl);
        var picBulletUri = listStyle == "bullet" ? GetPicBulletDataUri(numId, ilvl) : null;
        var tag = listStyle == "bullet" ? "ul" : "ol";

        // Swap the open list if the type changed (ul ↔ ol).
        if (cellListTag != null && cellListTag != tag)
        {
            sb.Append($"</{cellListTag}>");
            cellListTag = null;
        }

        var (lvlLeft, lvlHanging) = GetListLevelIndentFull(numId, ilvl);
        var indentPt = lvlLeft / 20.0;
        if (indentPt < 18) indentPt = 18;
        var hangingPt = lvlHanging / 20.0;
        var listStyleParts = $"padding-left:{indentPt:0.#}pt;margin:0";
        if (tag == "ol") listStyleParts += ";list-style-type:none";
        if (picBulletUri != null)
            listStyleParts += $";list-style-image:url('{picBulletUri}')";
        else if (tag == "ul")
        {
            listStyleParts += ";list-style-image:none";
            var bulletType = lvlText switch
            {
                "o" => "circle",
                "◦" => "circle",
                "" or "▪" => "square",
                _ => "disc"
            };
            listStyleParts += $";list-style-type:{bulletType}";
        }

        if (cellListTag == null)
        {
            sb.Append($"<{tag} style=\"{listStyleParts}\">");
            cellListTag = tag;
        }

        // Ordered lists render the marker via an inline span (same as body),
        // since list-style-type:none suppresses the native ::marker. Build it
        // first so the <li> can carry the hanging text-indent (mirrors the
        // body-level path) — without it the marker box would push the cell
        // list text right of the bullet text.
        string? olMarkerSpan = null;
        var paraStyle = GetParagraphInlineCss(para, isListItem: true);
        if (tag == "ol")
        {
            var template = string.IsNullOrEmpty(lvlText) ? $"%{ilvl + 1}" : lvlText!;
            var counter = (GetStartValue(numId, ilvl) ?? 1);
            var marker = System.Text.RegularExpressions.Regex.Replace(template, @"%(\d)", m =>
            {
                var lvlFmt = GetNumberingFormat(numId, ilvl);
                return OfficeCli.Core.WordNumFmtRenderer.Render(counter, lvlFmt);
            });
            var suff = GetLevelSuffix(numId, ilvl);
            var jc = GetLevelJustification(numId, ilvl);
            var markerWidth = hangingPt > 0 ? $"{hangingPt:0.#}pt" : "3em";
            var markerPadding = suff switch { "nothing" => "0", "space" => "0.25em", _ => "0.5em" };
            // Default/left → right-align so the number hugs the text like the
            // bullet ::marker (ul ignores lvlJc); explicit center kept distinct.
            var align = jc switch { "center" => "center", _ => "right" };
            var inlineMarkerCss = GetMarkerInlineCss(numId, ilvl, para);
            var markerStyle = $"display:inline-block;min-width:{markerWidth};padding-right:{markerPadding};text-align:{align}";
            if (!string.IsNullOrEmpty(inlineMarkerCss))
                markerStyle = inlineMarkerCss + ";" + markerStyle;
            olMarkerSpan = $"<span style=\"{markerStyle}\">{HtmlEncode(marker)}</span>";
            var hangCss = $"text-indent:calc(-{markerWidth} - {markerPadding})";
            paraStyle = string.IsNullOrEmpty(paraStyle) ? hangCss : paraStyle + ";" + hangCss;
        }

        sb.Append("<li");
        sb.Append($" class=\"marker-{numId}-{ilvl}\"");
        if (!string.IsNullOrEmpty(paraStyle))
            sb.Append($" style=\"{paraStyle}\"");
        sb.Append(">");
        if (olMarkerSpan != null)
            sb.Append(olMarkerSpan);

        RenderParagraphContentHtml(sb, para);
        sb.Append("</li>");
    }
}
