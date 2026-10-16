// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
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

    // olState threads the body walk's shared ordered-list counter into cell
    // list items so a cell's <ol> continues the document-flow numbering instead
    // of restarting at the level start. Null at isolated content roots
    // (header/footer/footnote/textbox) — those pass a fresh per-table state so
    // multi-item cell lists still advance 1./2./3. within the root, without
    // crossing into the body counter. (CONSISTENCY(list-marker))
    private void RenderTableHtml(StringBuilder sb, Table table, string? dataPath = null, int depth = 0, OrderedListNumberingState? olState = null)
    {
        olState ??= new OrderedListNumberingState();
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
        // Table-level default cell margin (<w:tblCellMar>), with style fallback
        // mirroring tblBorders. Cells consult this when they lack their own
        // tcMar; an explicit value (incl. 0) overrides the hardcoded 5.4pt L/R
        // td-padding default so a tblCellMar=0 invoice table doesn't lose
        // content width to phantom padding under table-layout:fixed.
        var tblCellMar = tblPr?.TableCellMarginDefault;
        if (tblCellMar == null && styleId != null)
            tblCellMar = ResolveTableStyleCellMargin(styleId);

        // Parse tblLook bitmask for conditional formatting
        var tblLook = ParseTableLook(tblPr);

        // Resolve conditional formatting from table style
        var condFormats = styleId != null ? ResolveTableStyleConditionalFormats(styleId) : null;

        // Resolve the table-style base run properties (<w:style><w:rPr>) — the
        // whole-table run formatting (e.g. white caps text). Per-cell these are
        // merged into the run cascade just above docDefaults, with the matching
        // conditional-format (firstRow/band…) rPr layered on top. See
        // ResolveEffectiveRunPropertiesCore step 1b.
        var tableStyleBaseRunProps = styleId != null ? ResolveTableStyleBaseRunProps(styleId) : null;

        // Resolve the table-style base cell shading (<w:style><w:tcPr><w:shd>) —
        // the whole-table cell fill. Used as the lowest-priority cell background
        // fallback so a dark-list table's blue fill shows behind its white run
        // color (which is applied on the <table>). Conditional-format shd and
        // direct cell shd both override it.
        var tableStyleCellFill = styleId != null
            ? ResolveShadingFill(ResolveTableStyleCellShading(styleId))
            : null;

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
                && LenientDxa(tblInd?.Width) is int indW && indW > 0)
                tableStyles.Add($"margin-left:{indW / 20.0:0.#}pt");
        }

        // Apply base table style rPr (font-size, color, alignment) to the <table>
        if (styleId != null)
        {
            var baseStyle = FindStyleById(styleId);
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

        // Set when an auto-layout (tblW=auto / no tblW) table carries a complete tblGrid: it then
        // gets a definite width + table-layout:fixed so its colgroup proportions become hard column
        // widths (prevents pure-text columns collapsing to 1-char vertical when a list-bearing cell
        // would otherwise eat the row's width). See the auto-fit comment in the width block below.
        bool autoGridFixable = false;

        // Set when a pct-width table (w:tblW type="pct", e.g. width:100%) carries a complete tblGrid:
        // like autoGridFixable, it gets table-layout:fixed (below) so the colgroup proportions become
        // hard column widths. Without it the browser's auto algorithm lets a wide/unbreakable cell
        // override the declared col widths and squeeze pure-text columns into a 1-char vertical strip
        // ("O/w/n/e/r"). R31/R32's fixed pin only covered tblW=dxa / auto+grid tables; pct-width tables
        // (common in templates) fell through. The percentage width itself stays on the table.
        bool pctGridFixable = false;

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
            // A complete tblGrid lets the colgroup percentages act as hard column proportions under
            // table-layout:fixed (pinned below), matching Word's column sizing instead of letting the
            // browser's content-driven auto algorithm collapse text columns.
            var pctGrid = table.GetFirstChild<TableGrid>();
            var pctGridCols = pctGrid?.Elements<GridColumn>().ToList();
            if (pctGridCols != null && pctGridCols.Count > 0)
            {
                pctGridFixable = pctGridCols.All(gc =>
                    gc.Width?.Value is string gw && int.TryParse(gw, out var v) && v > 0);
            }
        }
        else
        {
            // No explicit tblW or type=auto: use gridCol sum for the table width (Word auto-fit behavior).
            // Word auto-fit does NOT let columns grow past their grid widths to the point of starving a
            // neighbour — it fixes the column widths from the grid and wraps content inside each cell.
            // A browser auto table-layout, in contrast, distributes width by content min/max, and with the
            // page-body `overflow-wrap:anywhere` rule the min-content of every text column collapses to one
            // character. A row whose middle/last cell carries long list content then steals the whole width,
            // squeezing pure-text columns into a 1-char vertical strip ("P/a/g/e/St/r/u/c…") while long cells
            // overflow the page edge. When the grid is complete we therefore pin a *definite* width plus
            // table-layout:fixed (below) so the colgroup proportions become hard column widths and content
            // wraps inside its column, matching Word. Tables with a partial/absent grid keep content sizing.
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
                    // fixed layout already uses a definite width; auto layout with a complete grid now
                    // also gets a definite width so the table-layout:fixed pin (below) can resolve the
                    // colgroup percentages instead of falling back to content distribution.
                    autoGridFixable = !isFixed;
                    var prop = (isFixed || autoGridFixable) ? "width" : "max-width";
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

        // Table-level RTL (w:bidiVisual on tblPr): Word mirrors the column order
        // so the first logical cell sits at the right edge (COL-C | COL-B | COL-A).
        // CSS direction:rtl on the table reverses the table-cell layout order,
        // reproducing that mirror. get already surfaces direction=rtl; this was a
        // pure-render gap.
        var tblBidiVisual = tblPr?.GetFirstChild<BiDiVisual>();
        if (tblBidiVisual != null)
        {
            // CT_OnOff: no val (or a truthy val) is ON; an explicit falsey val
            // is OFF. Read the raw attribute text, mirroring the Get-side
            // readback in Navigation.cs so the render matches direction=rtl.
            var bidiRaw = tblBidiVisual.Val?.InnerText;
            if (bidiRaw is null || !(bidiRaw is "0" or "false" or "off"))
                tableStyles.Add("direction:rtl");
        }

        // Fixed-layout tables (w:tblLayout type="fixed") encode hard per-column
        // widths in tblGrid. Word treats those widths as upper bounds and wraps
        // long cell content within the column. Without CSS table-layout:fixed the
        // browser treats <col> widths as *minimums* and lets unbreakable content
        // (esp. long header text) expand the column past its declared width, which
        // overflows the page right edge. Pin table-layout:fixed so the colgroup
        // widths become hard caps and over-long cell text wraps inside the column,
        // matching Word. Only applied when an explicit tblGrid is present (autofit /
        // no-grid tables keep their content-driven sizing).
        var isTableFixedLayout = tblPr?.TableLayout?.Type?.InnerText == "fixed";
        if ((isTableFixedLayout && table.GetFirstChild<TableGrid>()?.Elements<GridColumn>().Any() == true)
            || autoGridFixable || pctGridFixable)
            tableStyles.Add("table-layout:fixed");

        var tableClass = tableBordersNone ? "borderless" : "";
        var tableStyleAttr = tableStyles.Count > 0 ? $" style=\"{string.Join(";", tableStyles)}\"" : "";
        var dataPathAttr = !string.IsNullOrEmpty(dataPath) ? $" data-path=\"{dataPath}\"" : "";
        if (!string.IsNullOrEmpty(tableClass))
            sb.AppendLine($"<table class=\"{tableClass}\"{dataPathAttr}{tableStyleAttr}>");
        else
            sb.AppendLine($"<table{dataPathAttr}{tableStyleAttr}>");

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
                    var span = tc.TableCellProperties?.GridSpan?.Val?.Value ?? 1;
                    if (span < 1) span = 1;
                    var tcW = tc.TableCellProperties?.TableCellWidth;
                    // Only a single-column cell's pct can be attributed to one grid column. A cell
                    // that spans N columns carries the COMBINED pct for all N columns (whole-table
                    // units, 5000 = 100%); stamping that combined value onto the span's first column
                    // (and leaving the rest to gridCol fallback) produced wildly wrong widths whose
                    // sum exceeded 100% (e.g. a 3-col span's 71% landing on one column). For spanning
                    // pct cells we skip the per-cell path entirely and let every column it covers use
                    // the accurate gridCol proportions below (which already sum to ~100%).
                    if (span == 1 && tcW?.Type?.InnerText == "pct" && pctByCol[ci] == null
                        && int.TryParse(tcW.Width?.Value, out var pctVal) && pctVal > 0)
                    {
                        // pct units are 1/50th of a percent (5000 = 100%)
                        pctByCol[ci] = pctVal / 50.0;
                    }
                    // gridSpan-aware advance so column index stays aligned
                    ci += span;
                }
            }

            // R126: per-column max single-span dxa tcW. Word lets a cell's explicit
            // tcW drive its column when tcW > gridCol; under table-layout:fixed the
            // <col> width pins the cell instead, so a label cell with tcW=810 over a
            // gridCol=270 column collapses to ~13.5pt (vertical one-char-per-line).
            // Capture the widest single-span dxa tcW per column to correct the gridCol
            // below — but ONLY from rows whose total gridSpan equals colCount. Such a
            // row has exactly one cell per grid column, so cell index reliably maps to
            // column index AND each cell occupies exactly one column (its tcW is that
            // column's intended width). Rows that don't cover the grid carry cells that
            // effectively span multiple columns with no gridSpan element (e.g. a single
            // tcW=10350 full-width cell, or tcW boundaries that don't align with gridCol
            // boundaries); attributing their tcW to one column would mis-widen it — that
            // ambiguous geometry is the deferred R89 class, left on the gridCol path.
            var dxaMaxByCol = new double?[colCount];
            foreach (var r in table.Elements<TableRow>())
            {
                var rowCells = r.Elements<TableCell>().ToList();
                int rowSpanTotal = rowCells.Sum(c => { var s = c.TableCellProperties?.GridSpan?.Val?.Value ?? 1; return s < 1 ? 1 : s; });
                if (rowSpanTotal != colCount) continue; // only full-grid rows map cell→column reliably
                int ci = 0;
                foreach (var tc in rowCells)
                {
                    if (ci >= colCount) break;
                    var span = tc.TableCellProperties?.GridSpan?.Val?.Value ?? 1;
                    if (span < 1) span = 1;
                    var tcW = tc.TableCellProperties?.TableCellWidth;
                    if (span == 1 && tcW?.Type?.InnerText == "dxa"
                        && double.TryParse(tcW.Width?.Value, System.Globalization.NumberStyles.Float,
                            System.Globalization.CultureInfo.InvariantCulture, out var dxaVal) && dxaVal > 0
                        && (dxaMaxByCol[ci] == null || dxaVal > dxaMaxByCol[ci]))
                    {
                        dxaMaxByCol[ci] = dxaVal;
                    }
                    ci += span;
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
                    var twips = double.Parse(w, System.Globalization.CultureInfo.InvariantCulture);
                    // R126: if a single-span cell in this column carries an explicit dxa
                    // tcW wider than the gridCol, the cell's width drives the column (Word
                    // behavior). Widen the <col> to that tcW so fixed layout doesn't pin a
                    // label cell to a too-narrow gridCol and collapse its text vertically.
                    if (colIdx < colCount && dxaMaxByCol[colIdx] is double dxaW && dxaW > twips)
                        twips = dxaW;
                    var pt = twips / 20.0; // twips to pt
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
        var totalRows = rows.Count;
        var totalCols = tblGrid?.Elements<GridColumn>().Count()
            ?? (rows.FirstOrDefault() is { } firstRow ? GetRowCellsFlattened(firstRow).Count : 0);

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
            sb.AppendLine(isHeader ? $"<tr class=\"header-row\"{hdrMarker}{rowDataPathAttr}{trStyle}>" : $"<tr{rowDataPathAttr}{trStyle}>");

            int colIdx = 0;
            // Cell-level content controls (<w:sdt> wrapping a <w:tc> as a direct
            // <w:tr> child — Word's dropdown-bound / placeholder form-field cell
            // shape) are NOT direct TableCell children of the row, so
            // Elements<TableCell>() dropped them entirely, leaving the form cell
            // blank (e.g. the staff-evaluation form's "Click or tap here to
            // enter text." placeholder cells under <w:showingPlcHdr>). Flatten
            // through GetRowCellsFlattened so every cell — wrapped or not —
            // renders, mirroring the Get/Query/dump cell-flatten contract.
            foreach (var cell in GetRowCellsFlattened(row))
            {
                var tag = isHeader ? "th" : "td";
                var condTypes = GetConditionalTypes(tblLook, rowIdx, colIdx, totalRows, totalCols);
                var cellStyle = GetTableCellInlineCss(cell, tableBordersNone, tblBorders, condFormats, condTypes,
                    rowIdx, colIdx, totalRows, totalCols, exactRowHeightPt, tblCellMar, tableStyleCellFill);

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

                // A cell that anchors wrapNone shapes positioned relative to the
                // column/paragraph (e.g. checkbox rectangles floated over a label
                // list) becomes the position:relative containing block for those
                // absolutely positioned shapes — applied on the <td> rather than
                // the inner paragraph div so the label text keeps its normal flow
                // height (a relative div whose only in-flow content is wrapped
                // text inside a table cell collapses the row to zero otherwise).
                // The shapes' left/top posOffsets are measured from the cell's
                // content box, which coincides with the column/paragraph origin.
                if (CellAnchorsSubParagraphShape(cell))
                    cellStyle = string.IsNullOrEmpty(cellStyle) ? "position:relative" : cellStyle + ";position:relative";

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

                // hRule="exact": wrap content in an inner div whose flex column
                // takes over vertical alignment (the td's vertical-align applies
                // to the wrap as a whole, not to content within it). Use
                // min-height as a floor rather than a fixed height + max-height +
                // overflow:hidden — content taller than the exact value (e.g. a
                // label plus several stacked checkbox SDTs) would otherwise be
                // clipped to a single centered line, silently dropping the rest.
                // Priority: content stays visible over honoring the exact height
                // strictly (the R49/R31 don't-clip-content rule); Word shows the
                // content. Content at/under the exact value keeps that height.
                bool exactWrap = exactRowHeightPt.HasValue;
                if (exactWrap)
                {
                    var vAlign = cell.TableCellProperties?.TableCellVerticalAlignment?.Val?.Value;
                    string justify;
                    if (vAlign == TableVerticalAlignmentValues.Center) justify = "center";
                    else if (vAlign == TableVerticalAlignmentValues.Bottom) justify = "flex-end";
                    else justify = "flex-start";
                    sb.Append($"<div style=\"min-height:{exactRowHeightPt:0.#}pt;display:flex;flex-direction:column;justify-content:{justify}\">");
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

                // Walk cell children; block-level SDTs (content controls)
                // wrap real paragraphs/tables, so recurse into the SDT content
                // rather than dropping it. Placeholder/data-bound text inside a
                // text-box layout table (e.g. the newsletter sidebar's
                // "A Recent Success" block) lives under <w:sdt> here — handling
                // only Paragraph/Table lost every nested run.
                void RenderCellChild(OpenXmlElement child)
                {
                    if (TryEmitContainerBookmarkAnchor(sb, child)) return;
                    // OOXML allows w:altChunk directly under w:tc; the body
                    // loop already renders it, the cell walk dropped it.
                    if (child is AltChunk cellAltChunk)
                    {
                        CloseCellList();
                        RenderAltChunkHtml(sb, cellAltChunk);
                        return;
                    }
                    if (child is Paragraph cellPara)
                    {
                        // VML horizontal rule inside a cell — same pre-dispatch
                        // check as the body loop; the generic run walk skips
                        // w:pict, so without this the rule disappeared.
                        if (IsVmlHorizontalRule(cellPara))
                        {
                            CloseCellList();
                            RenderVmlHorizontalRule(sb, cellPara);
                            return;
                        }
                        // Display equation inside a cell: a <w:p> whose content is
                        // an <m:oMathPara>/<m:oMath> wrapper. Body paragraphs route
                        // these to a katex-formula span (HtmlPreview.cs ~line 2362);
                        // the cell path historically lacked the branch, so in-cell
                        // formulas (e.g. §7 Indicadores fractions) rendered blank.
                        // Mirror the body emit so cell math surfaces identically.
                        var cellOMath = cellPara.ChildElements.FirstOrDefault(e => e.LocalName == "oMathPara" || e.LocalName == "oMath" || e is M.Paragraph || e is M.OfficeMath);
                        if (cellOMath != null)
                        {
                            CloseCellList();
                            var mathLatex = FormulaParser.ToLatex(cellOMath);
                            sb.Append($"<div class=\"equation\"><span class=\"katex-formula\" data-formula=\"{HtmlEncodeAttr(mathLatex)}\" data-display=\"true\"></span></div>");
                            return;
                        }
                        var listStyle = GetParagraphListStyle(cellPara);
                        if (listStyle != null)
                        {
                            RenderCellListItem(sb, cellPara, listStyle, ref cellListTag, olState);
                            return;
                        }
                        CloseCellList();
                        var text = GetParagraphText(cellPara);
                        var pCss = GetParagraphInlineCss(cellPara);
                        sb.Append("<div");
                        if (!string.IsNullOrEmpty(pCss))
                            sb.Append($" style=\"{pCss}\"");
                        sb.Append(">");
                        // Emptiness must be judged by *visible* content, not raw
                        // run count: a paragraph holding only a textless run (e.g.
                        // an empty form-fill answer box whose <w:r> carries just an
                        // rPr) has runs.Count > 0 yet renders nothing, so the div
                        // collapsed to ~0 height and the row floor disappeared.
                        // Word still draws a one-line-tall empty box (line height =
                        // paragraph-mark run font). Treat a run as visible only when
                        // it carries text/drawing/break/tab/symbol/picture/field, so
                        // such empties get the &nbsp; placeholder and the line box
                        // forms at the resolved line-height (pCss carries it).
                        bool hasVisibleContent = !string.IsNullOrWhiteSpace(text)
                            || GetAllRuns(cellPara).Any(r =>
                                r.Descendants<Text>().Any(t => !string.IsNullOrEmpty(t.Text))
                                || r.Descendants<Drawing>().Any()
                                || r.Descendants<Break>().Any()
                                || r.Descendants<TabChar>().Any()
                                || r.Descendants<SymbolChar>().Any()
                                || r.GetFirstChild<FieldChar>() != null
                                || r.ChildElements.Any(c => c.LocalName == "pict"));
                        RenderParagraphContentHtml(sb, cellPara);
                        if (!hasVisibleContent) sb.Append("&nbsp;");
                        sb.Append("</div>");
                    }
                    else if (child is Table nestedTable)
                    {
                        CloseCellList();
                        RenderTableHtml(sb, nestedTable, depth: depth + 1, olState: olState);
                    }
                    else if (child is SdtBlock sdt && sdt.SdtContentBlock is { } sdtContent)
                    {
                        foreach (var sdtChild in sdtContent.ChildElements)
                            RenderCellChild(sdtChild);
                    }
                }

                // Assemble this cell's table-style run-property layers (base rPr
                // then matching conditional-format rPr, lowest→highest priority —
                // condTypes is already ordered band → firstCol/lastCol →
                // firstRow/lastRow) and stash them on _ctx so the run cascade
                // (ResolveEffectiveRunPropertiesCore step 1b) picks up white-caps
                // / band run formatting. Saved/restored like ImageHostPart.
                List<OpenXmlElement>? cellRunPropLayers = null;
                if (tableStyleBaseRunProps != null)
                    (cellRunPropLayers = new List<OpenXmlElement>()).AddRange(tableStyleBaseRunProps);
                if (condFormats != null)
                {
                    foreach (var ct in condTypes)
                    {
                        if (condFormats.TryGetValue(ct, out var cf) && cf.RunProperties != null)
                            (cellRunPropLayers ??= new List<OpenXmlElement>()).Add(cf.RunProperties);
                    }
                }
                var savedCellRunProps = _ctx.CurrentCellTableStyleRunProps;
                _ctx.CurrentCellTableStyleRunProps = cellRunPropLayers;

                foreach (var child in cell.ChildElements)
                    RenderCellChild(child);
                CloseCellList();

                _ctx.CurrentCellTableStyleRunProps = savedCellRunProps;

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
            var style = FindStyleById(currentId);
            if (style == null) break;
            var borders = style.StyleTableProperties?.TableBorders;
            if (borders != null) return borders;
            currentId = style.BasedOn?.Val?.Value;
        }
        return null;
    }

    /// <summary>Resolve the default cell margin (&lt;w:tblCellMar&gt;) from a
    /// table style (walking the basedOn chain). Mirrors ResolveTableStyleBorders
    /// — the first style in the chain that declares a tblCellMar wins.</summary>
    private TableCellMarginDefault? ResolveTableStyleCellMargin(string styleId)
    {
        // Word merges w:tblCellMar PER SIDE across the basedOn chain (verified
        // against real Word — unlike w:tblBorders, which replaces wholesale):
        // a child style declaring only w:top keeps the parent's left/right/
        // bottom margins. Walk derived→base, most-derived side wins.
        TableCellMarginDefault? merged = null;
        var visited = new HashSet<string>();
        var currentId = styleId;
        while (currentId != null && visited.Add(currentId))
        {
            var style = FindStyleById(currentId);
            if (style == null) break;
            var cm = style.StyleTableProperties?.TableCellMarginDefault;
            if (cm != null)
            {
                if (merged == null)
                    merged = (TableCellMarginDefault)cm.CloneNode(true);
                else
                    foreach (var side in cm.ChildElements)
                        if (!merged.ChildElements.Any(c => c.LocalName == side.LocalName))
                            merged.AppendChild(side.CloneNode(true));
            }
            currentId = style.BasedOn?.Val?.Value;
        }
        return merged;
    }

    /// <summary>Resolve the base cell shading (&lt;w:style&gt;&lt;w:tcPr&gt;&lt;w:shd&gt;)
    /// from a table style (walking the basedOn chain). Mirrors
    /// ResolveTableStyleBorders — the first style in the chain that declares a
    /// base tcPr/shd wins. This is the table-style whole-table cell fill (e.g. a
    /// dark-list table whose base tcPr paints every cell blue and base rPr writes
    /// white text); without it the white run color lands on the default white cell
    /// and the labels render as an invisible empty frame. Conditional-format
    /// (tblStylePr firstRow/band…) shd and direct cell shd both override this.</summary>
    private Shading? ResolveTableStyleCellShading(string styleId)
    {
        var visited = new HashSet<string>();
        var currentId = styleId;
        while (currentId != null && visited.Add(currentId))
        {
            var style = FindStyleById(currentId);
            if (style == null) break;
            var shd = style.StyleTableCellProperties?.GetFirstChild<Shading>();
            if (shd != null) return shd;
            currentId = style.BasedOn?.Val?.Value;
        }
        return null;
    }

    /// <summary>Resolve the table-style base run properties
    /// (&lt;w:style&gt;&lt;w:rPr&gt;) walking the basedOn chain, returned
    /// base→derived (lowest→highest priority) as a layer list. This is the
    /// whole-table run formatting (e.g. the Invoice table's base rPr that writes
    /// white caps text). It sits below the per-cell conditional-format
    /// (tblStylePr) rPr and below paragraph/character styles in the cascade —
    /// see ResolveEffectiveRunPropertiesCore step 1b. The caller appends the
    /// matching tblStylePr conditional rPr after these so the band wins.</summary>
    private List<OpenXmlElement>? ResolveTableStyleBaseRunProps(string styleId)
    {
        // Collect the basedOn chain (derived→base), then return base→derived.
        var visited = new HashSet<string>();
        var chain = new List<Style>();
        var currentId = styleId;
        while (currentId != null && visited.Add(currentId))
        {
            var style = FindStyleById(currentId);
            if (style == null) break;
            chain.Add(style);
            currentId = style.BasedOn?.Val?.Value;
        }
        chain.Reverse(); // base first

        List<OpenXmlElement>? layers = null;
        foreach (var style in chain)
        {
            var rPr = style.StyleRunProperties;
            if (rPr == null) continue;
            (layers ??= new List<OpenXmlElement>()).Add(rPr);
        }
        return layers;
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
            var style = FindStyleById(currentId);
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

    private static int CountRowSpan(Table table, TableRow startRow, TableCell startCell)
    {
        var rows = table.Elements<TableRow>().ToList();
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
    private void RenderCellListItem(StringBuilder sb, Paragraph para, string listStyle, ref string? cellListTag, OrderedListNumberingState olState)
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

        // BUG-R105: paragraph-direct <w:ind> overrides the numbering-level
        // indentation (same as the body path).
        var (lvlLeft, lvlHanging) = ResolveListIndent(para, numId, ilvl);
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
            // CONSISTENCY(bullet-glyph-map): shared with body path and
            // GetCustomListStyleString; default disc. Symbol-font bullets
            // resolve to the custom glyph string so an inline keyword doesn't
            // override the ::marker font-family.
            var bulletType = GetUlListStyleTypeCss(numId, ilvl, lvlText);
            listStyleParts += $";list-style-type:{bulletType}";
            // CONSISTENCY(bullet-text-indent-reset): mirror the body path — a
            // bullet's native ::marker must not inherit an ordered ancestor's
            // hanging text-indent, which would pull the first line over the disc.
            listStyleParts += ";text-indent:0";
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
            // Advance + render through the shared ordered-list engine (same as
            // the body walk) so a cell's list continues document-flow numbering
            // 1./2./3. instead of restarting at the level start on every item.
            // A numId of its OWN breaks that run first (GitHub #512): without
            // it the cell inherited the body walk's running count and numbered
            // its first item 100 instead of the level's own <w:start>. Same
            // numId stays a no-op, so a genuine continuation is preserved.
            BreakListRunOnNumIdChange(olState, numId);
            var seedAbsId = GetAbstractNumId(numId);
            AdvanceOrderedCounter(olState, numId, seedAbsId, ilvl);
            var marker = RenderOrderedMarker(olState, numId, ilvl, lvlText);
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

        // Deeper levels: the <ol>/<ul> element carries only the FIRST item's
        // indent (set when the list was opened), so a level-1+ item rendered
        // flat at the same x as its parent — the cell/textbox/header/footnote
        // "multi-level collapses to one level" gap. Indent each item by its
        // level's indent delta over the same list's level-0 indent (stateless:
        // no per-container base tracking needed; level 0 gets delta 0).
        double levelDeltaPt = 0;
        if (ilvl > 0)
        {
            var (baseLeft, _) = GetListLevelIndentFull(numId, 0);
            var basePt = baseLeft / 20.0;
            if (basePt < 18) basePt = 18;
            if (indentPt > basePt) levelDeltaPt = indentPt - basePt;
        }
        sb.Append("<li");
        sb.Append($" class=\"marker-{numId}-{ilvl}\"");
        if (levelDeltaPt > 0)
            paraStyle = string.IsNullOrEmpty(paraStyle)
                ? $"margin-left:{levelDeltaPt:0.#}pt"
                : paraStyle + $";margin-left:{levelDeltaPt:0.#}pt";
        if (!string.IsNullOrEmpty(paraStyle))
            sb.Append($" style=\"{paraStyle}\"");
        sb.Append(">");
        if (olMarkerSpan != null)
            sb.Append(olMarkerSpan);

        RenderParagraphContentHtml(sb, para);
        sb.Append("</li>");
    }
}
