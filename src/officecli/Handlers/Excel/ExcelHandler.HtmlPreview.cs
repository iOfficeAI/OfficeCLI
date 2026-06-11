// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeCli.Core;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
namespace OfficeCli.Handlers;

public partial class ExcelHandler
{
    // Wire the formula evaluator's TEXT() up to the cell number-format engine so
    // TEXT(value, format) applies date/time/percent/currency codes identically to
    // a cell carrying that numFmt. Core cannot reference Handlers, so the hook is
    // a static delegate it invokes only when registered (here, on first use).
    static ExcelHandler()
    {
        Core.FormulaEvaluator.NumberFormatProvider = ApplyNumberFormat;
    }

    // Theme color map (lazy-initialized from theme1.xml)
    private Dictionary<string, string>? _excelThemeColors;
    // Indexed color palette (default 64 + custom overrides from styles.xml)
    private string[]? _resolvedIndexedColors;

    private Dictionary<string, string> GetExcelThemeColors()
    {
        if (_excelThemeColors != null) return _excelThemeColors;
        var colorScheme = _doc.WorkbookPart?.ThemePart?.Theme?.ThemeElements?.ColorScheme;
        _excelThemeColors = Core.ThemeColorResolver.BuildColorMap(colorScheme);
        // Blank workbooks (BlankDocCreator) carry no ThemePart, so the map is
        // empty and theme-indexed colors (e.g. <color theme="4"/> = accent1)
        // would render black. Real Excel falls back to the default Office theme;
        // backfill any missing scheme name with the standard Office palette.
        foreach (var (name, hex) in DefaultOfficeThemeColors)
            if (!_excelThemeColors.ContainsKey(name)) _excelThemeColors[name] = hex;
        return _excelThemeColors;
    }

    /// <summary>
    /// Excel theme color index mapping:
    /// 0=lt1, 1=dk1, 2=lt2, 3=dk2, 4=accent1, 5=accent2, 6=accent3, 7=accent4, 8=accent5, 9=accent6
    /// </summary>
    private static readonly string[] ThemeIndexToName =
        ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6"];

    // Standard Office (default Excel) theme palette — used as a fallback when the
    // workbook has no ThemePart (e.g. blank docs created by BlankDocCreator).
    private static readonly Dictionary<string, string> DefaultOfficeThemeColors =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ["dk1"] = "000000", ["lt1"] = "FFFFFF", ["dk2"] = "44546A", ["lt2"] = "E7E6E6",
            ["accent1"] = "4472C4", ["accent2"] = "ED7D31", ["accent3"] = "A5A5A5",
            ["accent4"] = "FFC000", ["accent5"] = "5B9BD5", ["accent6"] = "70AD47",
            ["hlink"] = "0563C1", ["folHlink"] = "954F72",
        };

    private string? ResolveThemeColor(uint themeIndex, double? tintValue = null)
    {
        if (themeIndex >= (uint)ThemeIndexToName.Length) return null;
        var themeColors = GetExcelThemeColors();
        if (!themeColors.TryGetValue(ThemeIndexToName[themeIndex], out var hex)) return null;

        if (tintValue.HasValue && Math.Abs(tintValue.Value) > 0.001)
        {
            // Excel tint: positive = tint toward white, negative = shade toward black
            // Convert to OOXML 0-100000 range
            var t = tintValue.Value;
            if (t > 0)
                return Core.ColorMath.ApplyTransforms(hex, tint: (int)((1 - t) * 100000));
            else
                return Core.ColorMath.ApplyTransforms(hex, shade: (int)((1 + t) * 100000));
        }

        return $"#{hex}";
    }

    private string[] GetResolvedIndexedColors()
    {
        if (_resolvedIndexedColors != null) return _resolvedIndexedColors;

        // Start with default palette
        _resolvedIndexedColors = (string[])DefaultIndexedColors.Clone();

        // Check for custom overrides in styles.xml
        var stylesheet = _doc.WorkbookPart?.WorkbookStylesPart?.Stylesheet;
        var colors = stylesheet?.GetFirstChild<Colors>();
        var indexedColors = colors?.GetFirstChild<IndexedColors>();
        if (indexedColors != null)
        {
            int idx = 0;
            foreach (var rgbColor in indexedColors.Elements<RgbColor>())
            {
                if (idx < _resolvedIndexedColors.Length && rgbColor.Rgb?.Value != null)
                {
                    var raw = rgbColor.Rgb.Value;
                    _resolvedIndexedColors[idx] = FormatColorForCss(raw);
                }
                idx++;
            }
        }
        return _resolvedIndexedColors;
    }

    /// <summary>
    /// Generate a self-contained HTML file that previews all sheets as spreadsheet tables.
    /// Supports cell formatting (font, fill, borders, alignment), merged cells,
    /// column widths, row heights, frozen panes, and sheet tab switching.
    /// </summary>
    public string ViewAsHtml()
    {
        using var _cul = InvariantCultureScope.Enter();
        var sb = new StringBuilder();
        var sheets = GetWorksheets();
        // Real Excel omits hidden / very-hidden sheets from the tab strip and the
        // content slider. Filter them out up front so data-sheet indices and the
        // active-sheet selection (first VISIBLE sheet) stay consistent across both
        // loops below.
        sheets = sheets.Where(s => !IsSheetHidden(s.Name)).ToList();
        if (sheets.Count == 0) sheets = GetWorksheets();
        var wbStylesPart = _doc.WorkbookPart?.WorkbookStylesPart;
        var stylesheet = wbStylesPart?.Stylesheet;

        // If any sheet has a pivot table, build an editable in-memory copy so
        // we can re-materialize cells from the pivot cache without mutating
        // the live _doc. The copy's WorksheetParts replace the originals for
        // rendering; styles/theme come from _doc (identical).
        //
        // CONSISTENCY(pivot-clone-in-memory): we clone _doc directly instead of
        // re-opening _filePath from disk. The earlier "read the file back via
        // FileStream(FileShare.ReadWrite)" approach races the handler's still-
        // held editable handle on macOS and throws IOException despite the
        // share-mode hint — the error surfaces as a trailing "process cannot
        // access" stderr after every add pivot/slicer command, and worse, on
        // every SUBSEQUENT command once the file has a pivot part at all (the
        // `sheets.Any(...PivotTableParts...)` branch fires on every ViewAsHtml
        // from the NotifyWatch path). SpreadsheetDocument.Clone(Stream, bool)
        // serialises the already-loaded package into the MemoryStream without
        // touching disk, so there is no second file handle to race.
        MemoryStream? pivotMs = null;
        SpreadsheetDocument? pivotDoc = null;
        List<(string Name, WorksheetPart Part)>? pivotSheets = null;
        if (sheets.Any(s => s.Part.PivotTableParts.Any()))
        {
            pivotMs = new MemoryStream();
            pivotDoc = (SpreadsheetDocument)_doc.Clone(pivotMs, isEditable: true);
            pivotSheets = GetWorksheets(pivotDoc);

            foreach (var (_, wsPart) in pivotSheets)
            {
                if (wsPart.PivotTableParts.Any())
                    OfficeCli.Core.PivotTableHelper.RefreshPivotCellsForView(wsPart);
            }

            // Use the copy's stylesheet so new indent styles created by the
            // pivot refresh are visible to the HTML renderer.
            stylesheet = pivotDoc.WorkbookPart?.WorkbookStylesPart?.Stylesheet;
        }

        sb.AppendLine("<!DOCTYPE html>");
        sb.AppendLine("<html>");
        sb.AppendLine("<head>");
        sb.AppendLine("<meta charset=\"UTF-8\">");
        sb.AppendLine("<meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">");
        sb.AppendLine($"<title>{HtmlEncode(Path.GetFileName(_filePath))}</title>");
        sb.AppendLine("<style>");
        sb.AppendLine(GenerateExcelCss());
        sb.AppendLine("</style>");
        sb.AppendLine("</head>");
        sb.AppendLine("<body>");

        // File title
        sb.AppendLine($"<div class=\"file-title\">{HtmlEncode(Path.GetFileName(_filePath))}</div>");

        // Sheet content areas (tabs moved to bottom)
        sb.AppendLine("<div class=\"sheet-slider\">");
        for (int sheetIdx = 0; sheetIdx < sheets.Count; sheetIdx++)
        {
            var (sheetName, worksheetPart) = sheets[sheetIdx];
            // Use the pivot-refreshed copy's WorksheetPart when available
            var renderPart = pivotSheets != null && sheetIdx < pivotSheets.Count
                ? pivotSheets[sheetIdx].Part : worksheetPart;
            var activeClass = sheetIdx == 0 ? " active" : "";
            // Check if sheet is RTL
            var sheetView = GetSheet(renderPart).GetFirstChild<SheetViews>()?.GetFirstChild<SheetView>();
            var isRtl = sheetView?.RightToLeft?.Value == true;
            // ShowGridLines defaults to true; only false when explicitly set false.
            var showGridLines = sheetView?.ShowGridLines?.Value != false;
            var dirAttr = isRtl ? " dir=\"rtl\"" : "";
            sb.AppendLine($"<div class=\"sheet-content{activeClass}\" data-sheet=\"{sheetIdx}\"{dirAttr}>");
            var charts = CollectSheetCharts(worksheetPart, sheetName);
            // Shapes and textboxes (xdr:sp). Reuses the chart overlay
            // positioning pipeline — same (fromRow,toRow,fromCol,toCol,html)
            // tuple is consumed by RenderSheetTable to emit an absolutely-
            // positioned overlay over the sheet grid.
            var shapes = CollectSheetShapes(worksheetPart);
            if (shapes.Count > 0)
                charts.AddRange(shapes);
            // Embedded pictures (xdr:pic) — same overlay positioning pipeline.
            var pictures = CollectSheetPictures(worksheetPart);
            if (pictures.Count > 0)
                charts.AddRange(pictures);
            RenderSheetTable(sb, sheetName, renderPart, stylesheet, charts, sheetIdx, showGridLines);
            sb.AppendLine("</div>");
        }
        sb.AppendLine("</div>");

        // Sheet tabs at bottom (like real Excel)
        sb.AppendLine("<div class=\"sheet-tabs\" role=\"tablist\">");
        for (int i = 0; i < sheets.Count; i++)
        {
            var activeClass = i == 0 ? " active" : "";
            var tabColorStyle = "";
            var sheetProps = GetSheet(sheets[i].Part).GetFirstChild<SheetProperties>();
            var tabColorEl = sheetProps?.TabColor;
            if (tabColorEl?.Rgb?.Value != null)
            {
                var rgb = tabColorEl.Rgb.Value;
                if (rgb.Length > 6) rgb = rgb[^6..];
                // Hex-gate before inline style interpolation — unchecked
                // raw value would break out of the style attribute.
                if (rgb.Length == 6
                    && rgb.All(c => (c >= '0' && c <= '9') || (c >= 'A' && c <= 'F') || (c >= 'a' && c <= 'f')))
                    tabColorStyle = $" style=\"--tab-color:#{rgb}\"";
            }
            sb.AppendLine($"  <div class=\"sheet-tab{activeClass}\"{tabColorStyle} data-sheet=\"{i}\" role=\"tab\" tabindex=\"0\" onclick=\"switchSheet({i})\" onkeydown=\"if(event.key==='Enter'||event.key===' ')switchSheet({i})\">{HtmlEncode(sheets[i].Name)}</div>");
        }
        sb.AppendLine("</div>");

        // Sheet switching JavaScript
        sb.AppendLine("<script>");
        sb.AppendLine(GenerateExcelJs());
        sb.AppendLine("</script>");
        // CONSISTENCY(excel-virt): private virt script injected after standard overlay.
        // Open-source GetVirtScript() returns empty; private override loads watch-overlay-virt.js.
        var virtScript = GetVirtScript();
        if (virtScript.Length > 0)
        {
            sb.AppendLine("<script>");
            sb.AppendLine(virtScript);
            sb.AppendLine("</script>");
        }

        sb.AppendLine("</body>");
        sb.AppendLine("</html>");

        pivotDoc?.Dispose();
        pivotMs?.Dispose();

        return sb.ToString();
    }

    /// <summary>
    /// Get the number of sheets (for watch notifications).
    /// </summary>
    public int GetSheetCount() => GetWorksheets().Count;

    /// <summary>Get the 0-based index of a sheet by name, or -1 if not found.</summary>
    public int GetSheetIndex(string sheetName)
    {
        var sheets = GetWorksheets();
        for (int i = 0; i < sheets.Count; i++)
            if (string.Equals(sheets[i].Name, sheetName, System.StringComparison.OrdinalIgnoreCase))
                return i;
        return -1;
    }

    // ==================== Sheet Rendering ====================

    private void RenderSheetTable(StringBuilder sb, string sheetName, WorksheetPart worksheetPart, Stylesheet? stylesheet,
        List<(int fromRow, int toRow, int fromCol, int toCol, int widthPtHint, string html)>? charts = null, int sheetIdx = 0,
        bool showGridLines = true)
    {
        var ws = GetSheet(worksheetPart);
        var sheetData = ws.GetFirstChild<SheetData>();
        if (sheetData == null && (charts == null || charts.Count == 0))
        {
            if (worksheetPart.DrawingsPart?.WorksheetDrawing == null)
                sb.AppendLine("<div class=\"empty-sheet\">Empty sheet</div>");
            return;
        }

        // Read default dimensions from sheetFormatPr
        var sheetFmtPr = ws.GetFirstChild<SheetFormatProperties>();
        // Excel column width → pixels: chars * 7.0017 (DEFAULT_CHARACTER_WIDTH for Calibri 11)
        // pt = px * 0.75
        var defaultColWidthPt = sheetFmtPr?.DefaultColumnWidth?.Value != null
            ? sheetFmtPr.DefaultColumnWidth.Value * 7.0017 * 0.75 : 8.43 * 7.0017 * 0.75;
        var defaultRowHeightPt = sheetFmtPr?.DefaultRowHeight?.Value ?? 15.0;

        // Read default font size from stylesheet
        var defaultFontPt = 11.0;
        if (stylesheet?.Fonts != null && stylesheet.Fonts.Elements<Font>().Any())
        {
            var defFont = stylesheet.Fonts.Elements<Font>().First();
            defaultFontPt = defFont.FontSize?.Val?.Value ?? 11.0;
        }

        // Create formula evaluator for this sheet to compute uncached formula values
        var evaluator = sheetData != null ? new Core.FormulaEvaluator(sheetData, _doc.WorkbookPart) : null;

        // Collect merge info
        var mergeMap = BuildMergeMap(ws);

        // Build conditional formatting CSS overrides (skip if no cell data)
        var cfMap = sheetData != null ? BuildConditionalFormatMap(ws, stylesheet, sheetData, _doc.WorkbookPart) : new Dictionary<string, string>();
        var dataBarMap = sheetData != null ? BuildDataBarMap(ws, sheetData) : new Dictionary<string, string>();
        var iconSetMap = sheetData != null ? BuildIconSetMap(ws, sheetData) : new Dictionary<string, string>();
        // R12a: sparklines live in cells that often have no CellValue. Build the
        // host-cell → SVG map now; maxCol/maxRow are extended to cover those cells
        // below, after they're computed from cell data.
        var sparklineMap = BuildSparklineMap(ws);

        // AutoFilter header cells: every cell in the top row of an AutoFilter
        // range (sheet-level <autoFilter> and each table's own <autoFilter>)
        // gets a dropdown indicator, matching Excel's filter-button affordance.
        var autoFilterCells = BuildAutoFilterHeaderCells(ws, worksheetPart);

        // Table (ListObject) built-in style banding: header fill + alternating
        // row stripes derived from the workbook theme. Explicit cell fills win,
        // so this is applied only where the cell has no fill of its own.
        var tableStyleMap = BuildTableStyleMap(worksheetPart);

        // Collect column widths
        var colWidths = GetColumnWidths(ws);

        // Detect frozen panes
        var (frozenRows, frozenCols) = GetFrozenPanes(ws);

        // Compute cumulative left offsets for frozen columns (for sticky positioning)
        // Index 0 = row header width (30pt), index 1 = col 1 left offset, etc.
        var frozenLeftOffsets = new Dictionary<int, double>();
        if (frozenCols > 0)
        {
            double cumLeft = 30; // row header width in pt
            for (int fc = 1; fc <= frozenCols; fc++)
            {
                frozenLeftOffsets[fc] = cumLeft;
                cumLeft += colWidths.TryGetValue(fc, out var w) ? w : defaultColWidthPt;
            }
        }

        // Determine grid dimensions. Count all cells that exist in SheetData —
        // every Cell element with a CellReference contributes to maxRow/maxCol,
        // even if the cell is empty (no value, no formula). Empty cells are
        // explicitly created by the user or by Excel; either way they should
        // render so the grid matches the actual data range.
        var rows = sheetData?.Elements<Row>().ToList() ?? new List<Row>();
        int maxCol = 0;
        int maxRow = 0;
        foreach (var row in rows)
        {
            var rowIdx = (int)(row.RowIndex?.Value ?? 0);
            bool rowHasCells = false;
            foreach (var cell in row.Elements<Cell>())
            {
                var cellRef = cell.CellReference?.Value;
                if (cellRef == null) continue;
                var (colName, _) = ParseCellReference(cellRef);
                var colIdx = ColumnNameToIndex(colName);
                if (colIdx > maxCol) maxCol = colIdx;
                rowHasCells = true;
            }
            if (rowHasCells && rowIdx > maxRow) maxRow = rowIdx;
        }

        // Extend maxRow/maxCol from chart anchors even when no cell data
        if (charts != null)
        {
            foreach (var (fromRow, toRow, fromCol, toCol, _, _) in charts)
            {
                if (toRow > maxRow) maxRow = toRow;
                if (toCol > maxCol) maxCol = toCol;
            }
        }

        // Extend maxCol to cover columns that carry an explicit <col> width but no
        // cell data. Excel renders those columns (the user sized them on purpose),
        // and they are valid spill targets for a long text cell to their left.
        foreach (var widthCol in colWidths.Keys)
            if (widthCol > maxCol) maxCol = widthCol;

        // R12a: extend maxRow/maxCol to cover sparkline host cells (which often
        // have no CellValue and would otherwise be cropped out of the grid).
        foreach (var hostRef in sparklineMap.Keys)
        {
            var (hc, hr) = ParseCellReference(hostRef);
            var hcIdx = ColumnNameToIndex(hc);
            if (hcIdx > maxCol) maxCol = hcIdx;
            if (hr > maxRow) maxRow = hr;
        }

        // Extend maxRow/maxCol to cover conditional-formatting ranges: blank
        // in-range cells (e.g. containsBlanks fill on D1:D5 with only D1
        // populated) must exist in the grid for their CF style to display.
        // The CF contribution is clamped by the same render caps applied
        // below, so a whole-column sqref cannot inflate the grid (or the
        // truncation warning) past what would render anyway.
        var (cfMaxRow, cfMaxCol) = CfRangeExtents(ws.Elements<ConditionalFormatting>());
        cfMaxRow = Math.Min(cfMaxRow, GetHtmlRowCap());
        cfMaxCol = Math.Min(cfMaxCol, 200);
        if (cfMaxRow > maxRow) maxRow = cfMaxRow;
        if (cfMaxCol > maxCol) maxCol = cfMaxCol;

        // Empty sheet (no cells and no charts)
        if (maxRow == 0 || maxCol == 0)
        {
            if (worksheetPart.DrawingsPart?.WorksheetDrawing == null)
                sb.AppendLine("<div class=\"empty-sheet\">Empty sheet</div>");
            return;
        }

        // Extend maxRow/maxCol to include chart anchor ranges
        if (charts != null)
            foreach (var (_, toRow, fromCol, toCol, _, _) in charts)
            {
                if (toCol > maxCol) maxCol = toCol;
                if (toRow > maxRow) maxRow = toRow;
            }

        // Column cap: >200 cols is unusable in a browser table regardless of rendering mode.
        // Row cap: default 5000; overridable via OnGetHtmlRowCap when the rendering backend
        // keeps DOM node count bounded independently of sheet size.
        var actualRow = maxRow;
        var actualCol = maxCol;
        maxRow = Math.Min(maxRow, GetHtmlRowCap());
        maxCol = Math.Min(maxCol, 200);
        var truncated = actualRow > maxRow || actualCol > maxCol;

        // Build cell lookup: (row, col) → Cell
        var cellMap = new Dictionary<(int row, int col), Cell>();
        foreach (var row in rows)
        {
            var rowIdx = (int)(row.RowIndex?.Value ?? 0);
            if (rowIdx > maxRow) break;
            foreach (var cell in row.Elements<Cell>())
            {
                var cellRef = cell.CellReference?.Value;
                if (cellRef == null) continue;
                var (colName, _) = ParseCellReference(cellRef);
                var colIdx = ColumnNameToIndex(colName);
                if (colIdx <= maxCol)
                    cellMap[(rowIdx, colIdx)] = cell;
            }
        }

        // Row height and hidden row lookup
        var rowHeights = new Dictionary<int, double>();
        var hiddenRows = new HashSet<int>();
        foreach (var row in rows)
        {
            var rowIdx = (int)(row.RowIndex?.Value ?? 0);
            if (row.CustomHeight?.Value == true && row.Height?.Value != null)
                rowHeights[rowIdx] = row.Height.Value;
            // A row with height 0 is a hidden row in real Excel (mirrors the
            // hidden-column treatment, which drops width<=0 columns). Emit it
            // display:none rather than as a ~16px gap. The original row numbers
            // are preserved (subsequent rows keep their numbers; no renumber).
            if (row.Hidden?.Value == true
                || (row.Height?.Value != null && row.Height.Value == 0))
                hiddenRows.Add(rowIdx);
        }

        // Rotated-text rows auto-grow in real Excel so the vertical string is
        // visible. The HTML <td> only carries transform:rotate, which keeps the
        // glyph box at its un-rotated width — the row stays at default height and
        // clips. Bump the row's min-height to the rotated text extent (approx
        // text-length × font-size for ~90°), consistent with the spill/width
        // estimation heuristics elsewhere in this renderer.
        foreach (var ((r, _), cell) in cellMap)
        {
            var extent = EstimateRotatedCellHeightPt(cell, stylesheet, defaultFontPt);
            if (extent <= 0) continue;
            if (!rowHeights.TryGetValue(r, out var existing) || existing < extent)
                rowHeights[r] = extent;
        }

        // Compute cumulative top offsets for frozen rows (for sticky positioning)
        // Includes thead height (~24pt for column headers)
        var frozenTopOffsets = new Dictionary<int, double>();
        if (frozenRows > 0)
        {
            double cumTop = 24; // approximate thead (column header) height
            for (int fr = 1; fr <= frozenRows; fr++)
            {
                frozenTopOffsets[fr] = cumTop;
                if (rowHeights.TryGetValue(fr, out var rh))
                    cumTop += rh;
                else
                {
                    // Estimate row height from max font size in the row's cells
                    double maxFontPt = defaultFontPt;
                    foreach (var cell in cellMap.Where(kv => kv.Key.row == fr).Select(kv => kv.Value))
                    {
                        var si = cell.StyleIndex?.Value ?? 0;
                        if (stylesheet?.CellFormats != null && si < (uint)stylesheet.CellFormats.Elements<CellFormat>().Count())
                        {
                            var xf = stylesheet.CellFormats.Elements<CellFormat>().ElementAt((int)si);
                            var fontId = xf.FontId?.Value ?? 0;
                            if (stylesheet.Fonts != null && fontId < (uint)stylesheet.Fonts.Elements<Font>().Count())
                            {
                                var font = stylesheet.Fonts.Elements<Font>().ElementAt((int)fontId);
                                var sz = font.FontSize?.Val?.Value ?? defaultFontPt;
                                if (sz > maxFontPt) maxFontPt = sz;
                            }
                        }
                    }
                    cumTop += maxFontPt * 1.4 + 4; // font height + padding
                }
            }
        }

        // Collect hidden columns
        var hiddenCols = new HashSet<int>();
        foreach (var (colIdx, widthPx) in colWidths)
        {
            if (widthPx <= 0) hiddenCols.Add(colIdx);
        }

        // Columns without an explicit OOXML <col> width fall through to
        // defaultColWidthPt (Excel's default ~8.43 chars ≈ 44pt). We do NOT
        // auto-fit to content width: real Excel keeps the default column width
        // and lets long text spill into empty right-neighbour cells (see the
        // spill handling in the cell render below). Auto-fitting would grow the
        // column to hold the text, defeating both Excel fidelity and spill.

        // Build chart lookup: fromRow → chart info for inline insertion
        var chartAtRow = new Dictionary<int, (int toRow, int fromCol, int toCol, string html)>();
        if (charts != null)
            foreach (var (fromRow, toRow, fromCol, toCol, _, html) in charts)
                chartAtRow[fromRow] = (toRow, fromCol, toCol, html);

        // Compute total table width so the table sizes to its content (not the wrapper).
        // Without an explicit width, table-layout:fixed inside a flex wrapper shrinks columns
        // proportionally to fit the viewport, ignoring declared col widths.
        double totalTableWidthPt = 30; // row-header-col width
        for (int c = 1; c <= maxCol; c++)
        {
            if (hiddenCols.Contains(c)) continue;
            totalTableWidthPt += colWidths.TryGetValue(c, out var cw) ? cw : defaultColWidthPt;
        }

        // Start table (position:relative for chart overlays)
        sb.AppendLine("<div class=\"table-wrapper\" style=\"position:relative\">");
        var noGridClass = showGridLines ? "" : " class=\"no-grid\"";
        sb.AppendLine($"<table{noGridClass} style=\"width:{totalTableWidthPt:0.##}pt\">");
        sb.AppendLine($"<caption class=\"sr-only\">{HtmlEncode(sheetName)}</caption>");

        // Colgroup for column widths + header column (skip hidden columns to match td count)
        sb.Append("<colgroup><col class=\"row-header-col\">");
        for (int c = 1; c <= maxCol; c++)
        {
            if (hiddenCols.Contains(c)) continue; // skip hidden cols — tds are also skipped
            var width = colWidths.TryGetValue(c, out var w) ? w : defaultColWidthPt;
            sb.Append($"<col style=\"width:{width:0.##}pt\">");
        }
        sb.AppendLine("</colgroup>");

        // Column header row
        sb.Append("<thead><tr><th class=\"corner-cell\"");
        if (frozenRows > 0 || frozenCols > 0) sb.Append(" style=\"position:sticky;top:0;left:0;z-index:4\"");
        sb.Append("></th>");
        for (int c = 1; c <= maxCol; c++)
        {
            if (hiddenCols.Contains(c)) continue;
            var colName = IndexToColumnName(c);
            var isFrozenColHeader = frozenCols > 0 && c <= frozenCols;
            string stickyStyle;
            if (frozenRows > 0 && isFrozenColHeader)
            {
                var leftPt = frozenLeftOffsets.TryGetValue(c, out var lf) ? lf : 0;
                stickyStyle = $" style=\"position:sticky;top:0;left:{leftPt:0.##}pt;z-index:4\"";
            }
            else if (frozenRows > 0)
                stickyStyle = " style=\"position:sticky;top:0;z-index:3\"";
            else if (isFrozenColHeader)
            {
                var leftPt = frozenLeftOffsets.TryGetValue(c, out var lf2) ? lf2 : 0;
                stickyStyle = $" style=\"position:sticky;left:{leftPt:0.##}pt;z-index:3\"";
            }
            else
                stickyStyle = "";
            sb.Append($"<th class=\"col-header\" data-path=\"/{HtmlEncode(sheetName)}/col[{colName}]\"{stickyStyle}>{colName}</th>");
        }
        sb.AppendLine("</tr></thead>");

        // chartAtRow and sideCharts already built above

        // Visible column count for chart colspan
        var visibleColCount = Enumerable.Range(1, maxCol).Count(c => !hiddenCols.Contains(c));

        // CONSISTENCY(excel-virt): Extension point — private override in
        // ExcelHandler.HtmlPreview.Virt.cs replaces the full static tbody with a
        // JSON-data tbody + JS virtual renderer. BuildRowInnerHtml is shared for
        // cell rendering; open-source RenderTbody emits static <tr> elements.
        var ctx = new SheetRenderContext(sheetName, sheetIdx, cellMap, maxRow, maxCol,
            rowHeights, hiddenRows, hiddenCols, mergeMap, frozenRows, frozenCols,
            frozenLeftOffsets, frozenTopOffsets, cfMap, dataBarMap, iconSetMap, sparklineMap,
            tableStyleMap, autoFilterCells, stylesheet, evaluator, defaultColWidthPt, defaultRowHeightPt, colWidths);
        RenderTbody(sb, ctx);
        sb.AppendLine("</table>");

        // Render charts as absolute-positioned overlays on top of the table grid.
        // Position is computed from anchor row/col using column widths and row heights.
        if (charts != null)
        {
            var rowHeaderWidthPt = 30.0; // matches .row-header-col CSS
            foreach (var (fromRow, toRow, fromCol, toCol, widthPtHint, html) in charts)
            {
                // Compute left position: sum of column widths from col 1 to fromCol + row header
                double leftPt = rowHeaderWidthPt;
                for (int c = 1; c <= fromCol && c <= maxCol; c++)
                {
                    if (hiddenCols.Contains(c)) continue;
                    leftPt += colWidths.TryGetValue(c, out var cw) ? cw : defaultColWidthPt;
                }
                // Compute top position: sum of row heights from row 1 to fromRow + header row (~24px)
                double topPt = 24.0 * 0.75; // header row height in pt
                for (int r = 1; r <= fromRow && r <= maxRow; r++)
                {
                    if (hiddenRows.Contains(r)) continue;
                    topPt += rowHeights.TryGetValue(r, out var rh) ? rh : defaultRowHeightPt;
                }
                // Compute width/height from anchor span
                double widthPt = 0;
                for (int c = fromCol + 1; c <= toCol && c <= maxCol; c++)
                {
                    if (hiddenCols.Contains(c)) continue;
                    widthPt += colWidths.TryGetValue(c, out var cw2) ? cw2 : defaultColWidthPt;
                }
                double heightPt = 0;
                for (int r = fromRow + 1; r <= toRow && r <= maxRow; r++)
                {
                    if (hiddenRows.Contains(r)) continue;
                    heightPt += rowHeights.TryGetValue(r, out var rh2) ? rh2 : defaultRowHeightPt;
                }
                // Chart-container min-size fallback. Pictures (xdr:pic) must
                // reflect their true anchor span, not the chart default, so a
                // small-span picture isn't ballooned to 400x250.
                bool isPicture = html.Contains("xlsx-picture");
                if (!isPicture)
                {
                    if (widthPt < 100) widthPt = 400; // fallback min size
                    if (heightPt < 50) heightPt = 250;
                }
                else
                {
                    // Picture floor so a zero-span anchor still renders visibly.
                    if (widthPt < 1) widthPt = defaultColWidthPt;
                    if (heightPt < 1) heightPt = defaultRowHeightPt;
                }
                // Charts supply their EstimateChartSize width (which includes the
                // partial-column EMU offset this column-sum drops); honour it so the
                // card matches Excel's width instead of being clamped a column short.
                // Pictures/shapes pass 0, so this only affects charts.
                if (widthPtHint > 0) widthPt = widthPtHint;

                sb.AppendLine($"<div style=\"position:absolute;left:{leftPt:0.##}pt;top:{topPt:0.##}pt;width:{widthPt:0.##}pt;height:{heightPt:0.##}pt;z-index:10;pointer-events:auto;display:flex\" data-from-col=\"{fromCol}\" data-from-row=\"{fromRow}\">");
                sb.Append(html);
                sb.AppendLine("</div>");
            }
        }

        // Truncation warning
        if (truncated)
            sb.AppendLine($"<div class=\"truncation-warning\">Showing {maxRow} of {actualRow} rows, {maxCol} of {actualCol} columns</div>");
        sb.AppendLine("</div>"); // close table-wrapper
    }

    // ==================== Merge Map ====================

    internal record struct MergeInfo(bool IsAnchor, int RowSpan, int ColSpan);

    // CONSISTENCY(excel-virt): Packages all sheet-level computed data needed to render
    // tbody rows. Passed to RenderTbody so the private virt override can serialise all
    // cell HTML to JSON without re-running the data-collection logic.
    internal record SheetRenderContext(
        string SheetName,
        int SheetIdx,
        Dictionary<(int row, int col), Cell> CellMap,
        int MaxRow, int MaxCol,
        Dictionary<int, double> RowHeights,
        HashSet<int> HiddenRows,
        HashSet<int> HiddenCols,
        Dictionary<string, MergeInfo> MergeMap,
        int FrozenRows, int FrozenCols,
        Dictionary<int, double> FrozenLeftOffsets,
        Dictionary<int, double> FrozenTopOffsets,
        Dictionary<string, string> CfMap,
        Dictionary<string, string> DataBarMap,
        Dictionary<string, string> IconSetMap,
        Dictionary<string, string> SparklineMap,
        Dictionary<string, string> TableStyleMap,
        HashSet<string> AutoFilterCells,
        Stylesheet? Stylesheet,
        Core.FormulaEvaluator? Evaluator,
        double DefaultColWidthPt,
        double DefaultRowHeightPt,
        Dictionary<int, double> ColWidths);

    // CONSISTENCY(excel-virt): Private ExcelHandler.HtmlPreview.Virt.cs implements
    // OnRenderTbody to emit virtualised rows (JSON data + empty tbody) and sets
    // handled=true to skip the default. When no private implementation exists the
    // partial call is removed by the compiler and the default static rendering runs.
    partial void OnRenderTbody(StringBuilder sb, SheetRenderContext ctx, ref bool handled);

    // CONSISTENCY(excel-virt): default 5000-row cap for HTML preview; backend can
    // override via OnGetHtmlRowCap when DOM node count is bounded independently.
    partial void OnGetHtmlRowCap(ref int cap);
    internal int GetHtmlRowCap()
    {
        var cap = 5000;
        OnGetHtmlRowCap(ref cap);
        return cap;
    }

    internal void RenderTbody(StringBuilder sb, SheetRenderContext ctx)
    {
        bool handled = false;
        OnRenderTbody(sb, ctx, ref handled);
        if (handled) return;
        // Default: render all rows as static <tr> elements.
        sb.AppendLine("<tbody>");
        for (int r = 1; r <= ctx.MaxRow; r++)
        {
            if (ctx.HiddenRows.Contains(r)) { sb.AppendLine($"<tr data-row=\"{ctx.SheetIdx}-{r}\" style=\"display:none\"></tr>"); continue; }
            bool isRowFrozen = ctx.FrozenRows > 0 && r <= ctx.FrozenRows;
            var rowStyles = new List<string>();
            // Every row gets a height (explicit, else the sheet default ~15pt) so
            // empty rows don't collapse — matching Excel and keeping the grid's row
            // positions consistent with the chart anchor math (which uses the same).
            var rh = ctx.RowHeights.TryGetValue(r, out var explicitRh) ? explicitRh : ctx.DefaultRowHeightPt;
            rowStyles.Add($"height:{rh:0.##}pt");
            if (isRowFrozen) rowStyles.Add("background:#fff");
            var rowStyle = rowStyles.Count > 0 ? $" style=\"{string.Join(";", rowStyles)}\"" : "";
            var frozenAttr = isRowFrozen ? " data-frozen=\"1\"" : "";
            sb.Append($"<tr data-row=\"{ctx.SheetIdx}-{r}\"{rowStyle}{frozenAttr}>");
            sb.Append(BuildRowInnerHtml(ctx, r, isRowFrozen));
            sb.AppendLine("</tr>");
        }
        sb.AppendLine("</tbody>");
    }

    // CONSISTENCY(excel-virt): Shared row-cell renderer used by RenderTbody (open-source
    // static rendering) and ExcelHandler.HtmlPreview.Virt.cs (JSON serialisation).
    // Returns the <tr> inner content: row-header <th> + all cell <td> elements,
    // without the <tr> wrapper.
    // AutoFilter dropdown indicator: a small right-aligned filter-arrow button
    // box matching Excel's filter-button look. Rendered on each header cell in
    // an AutoFilter range's top row.
    private const string AutoFilterIndicatorHtml =
        "<span class=\"autofilter-btn\" style=\"display:inline-block;float:right;margin-left:3px;" +
        "padding:0 2px;border:1px solid #b0b0b0;border-radius:2px;background:#f3f3f3;" +
        "font-size:0.7em;line-height:1.2;color:#444\">▼</span>";

    internal string BuildRowInnerHtml(SheetRenderContext ctx, int r, bool isRowFrozen)
    {
        var rowSb = new StringBuilder();
        string rowHeaderStyle;
        if (isRowFrozen)
            rowHeaderStyle = " style=\"position:sticky;top:0;left:0;z-index:3\"";
        else if (ctx.FrozenCols > 0)
            rowHeaderStyle = " style=\"position:sticky;left:0;z-index:2\"";
        else
            rowHeaderStyle = "";
        rowSb.Append($"<th class=\"row-header\" data-path=\"/{HtmlEncode(ctx.SheetName)}/row[{r}]\"{rowHeaderStyle}>{r}</th>");

        for (int c = 1; c <= ctx.MaxCol; c++)
        {
            if (ctx.HiddenCols.Contains(c)) continue;
            var cellRef = $"{IndexToColumnName(c)}{r}";
            if (ctx.MergeMap.TryGetValue(cellRef, out var mergeInfo))
            {
                if (!mergeInfo.IsAnchor) continue;
                var cell = ctx.CellMap.TryGetValue((r, c), out var mc) ? mc : null;
                // Merged-region perimeter borders come from the perimeter member cells:
                // right edge from the right-column member, bottom edge from the bottom-row
                // member. The anchor's own right/bottom edges are interior to the merge.
                var rightMember = ctx.CellMap.TryGetValue((r, c + mergeInfo.ColSpan - 1), out var rmc) ? rmc : null;
                var bottomMember = ctx.CellMap.TryGetValue((r + mergeInfo.RowSpan - 1, c), out var bmc) ? bmc : null;
                var style = GetCellStyleCss(cell, ctx.Stylesheet, ctx.FrozenRows, ctx.FrozenCols, r, c, ctx.FrozenLeftOffsets, ctx.FrozenTopOffsets, ctx.CfMap, ctx.DataBarMap, ctx.IconSetMap, ctx.TableStyleMap, mergePerimeter: true, rightBorderCell: rightMember, bottomBorderCell: bottomMember);
                var value = cell != null ? GetFormattedCellValue(cell, ctx.Stylesheet, ctx.Evaluator) : "";
                var richHtml = cell != null ? TryBuildRichTextHtml(cell) : null;
                var adjColSpan = mergeInfo.ColSpan;
                if (adjColSpan > 1 && ctx.HiddenCols.Count > 0)
                    for (int hc = c + 1; hc < c + mergeInfo.ColSpan; hc++)
                        if (ctx.HiddenCols.Contains(hc)) adjColSpan--;
                var spanAttrs = "";
                if (adjColSpan > 1) spanAttrs += $" colspan=\"{adjColSpan}\"";
                if (mergeInfo.RowSpan > 1) spanAttrs += $" rowspan=\"{mergeInfo.RowSpan}\"";
                // Rich-text runs render as pre-built spans (already encoded); the
                // bar/icon overlay path is mutually exclusive with rich text here.
                var hlinkHtml = TryBuildHyperlinkFormulaHtml(cell, value);
                var content = hlinkHtml != null && !ctx.DataBarMap.ContainsKey(cellRef) && !ctx.IconSetMap.ContainsKey(cellRef)
                    ? hlinkHtml
                    : richHtml != null && !ctx.DataBarMap.ContainsKey(cellRef) && !ctx.IconSetMap.ContainsKey(cellRef)
                    ? richHtml
                    : BuildCellContent(cellRef, value, ctx.DataBarMap, ctx.IconSetMap);
                content = WrapVerticalAlign(content, GetCellVerticalAlign(cell, ctx.Stylesheet), richHtml);
                if (ctx.SparklineMap.TryGetValue(cellRef, out var spkSvg)) content = spkSvg + content;
                var diagSvg = TryBuildCellDiagonalSvg(cell, ctx.Stylesheet) ?? "";
                if (ctx.AutoFilterCells.Contains(cellRef)) content += AutoFilterIndicatorHtml;
                rowSb.Append($"<td data-path=\"/{HtmlEncode(ctx.SheetName)}/{cellRef}\"{GetFormulaAttr(cell)}{spanAttrs}{style}>{diagSvg}{content}</td>");
            }
            else
            {
                var cell = ctx.CellMap.TryGetValue((r, c), out var nc) ? nc : null;
                var style = GetCellStyleCss(cell, ctx.Stylesheet, ctx.FrozenRows, ctx.FrozenCols, r, c, ctx.FrozenLeftOffsets, ctx.FrozenTopOffsets, ctx.CfMap, ctx.DataBarMap, ctx.IconSetMap, ctx.TableStyleMap);
                var value = cell != null ? GetFormattedCellValue(cell, ctx.Stylesheet, ctx.Evaluator) : "";
                var richHtml = cell != null ? TryBuildRichTextHtml(cell) : null;
                var hlinkHtml = TryBuildHyperlinkFormulaHtml(cell, value);
                var content = hlinkHtml != null && !ctx.DataBarMap.ContainsKey(cellRef) && !ctx.IconSetMap.ContainsKey(cellRef)
                    ? hlinkHtml
                    : richHtml != null && !ctx.DataBarMap.ContainsKey(cellRef) && !ctx.IconSetMap.ContainsKey(cellRef)
                    ? richHtml
                    : BuildCellContent(cellRef, value, ctx.DataBarMap, ctx.IconSetMap);
                content = WrapVerticalAlign(content, GetCellVerticalAlign(cell, ctx.Stylesheet), richHtml);
                if (ctx.SparklineMap.TryGetValue(cellRef, out var spkSvg)) content = spkSvg + content;
                var diagSvg = TryBuildCellDiagonalSvg(cell, ctx.Stylesheet) ?? "";
                // Text-spill emulation (Excel-fidelity): a non-wrapped left/general
                // aligned text cell with empty right-neighbours paints its overflow
                // across those neighbours, clipping at the first occupied cell. The
                // <td> stays 1 column wide (preserving borders/gridlines/merges); the
                // text lives in an inline span that overflows visibly up to the summed
                // empty-neighbour width.
                // Shrink-to-fit (Excel-fidelity): a shrinkToFit cell with overflowing
                // text compresses the font so the WHOLE string fits — it never spills
                // and never truncates. Branch before the spill/ellipsis logic.
                string spillClass = "";
                var shrinkPt = GetShrinkToFitFontPt(ctx, cell, value, c);
                if (shrinkPt > 0)
                {
                    content = $"<span class=\"shrink-fit\" style=\"font-size:{shrinkPt:0.##}pt;white-space:nowrap\">{content}</span>";
                }
                else
                {
                    var spillWidth = GetSpillWidthPt(ctx, cell, value, r, c);
                    if (spillWidth > 0)
                    {
                        spillClass = " class=\"spill\"";
                        // text-decoration set on the <td> does not paint under the
                        // text that bleeds past the td via overflow:visible, so copy
                        // the cell's text-decoration* declarations onto the span that
                        // actually renders the spilled glyphs (underline / double /
                        // strikethrough). Matches real Excel, which underlines the
                        // whole spilled string.
                        var spanDecor = ExtractTextDecorationCss(style);
                        content = $"<span class=\"spill-text\" style=\"max-width:{spillWidth:0.##}pt{spanDecor}\">{content}</span>";
                    }
                }
                if (ctx.AutoFilterCells.Contains(cellRef)) content += AutoFilterIndicatorHtml;
                rowSb.Append($"<td data-path=\"/{HtmlEncode(ctx.SheetName)}/{cellRef}\"{GetFormulaAttr(cell)}{spillClass}{style}>{diagSvg}{content}</td>");
            }
        }
        return rowSb.ToString();
    }

    // Pull the text-decoration / text-decoration-style declarations out of a
    // cell's ` style="..."` attribute string so they can be re-applied to the
    // inner .spill-text span (the element that actually paints the overflowing
    // glyphs). Returns a ";"-prefixed CSS fragment, or "" when none present.
    private static string ExtractTextDecorationCss(string tdStyleAttr)
    {
        if (string.IsNullOrEmpty(tdStyleAttr)) return "";
        var sb = new System.Text.StringBuilder();
        foreach (var decl in tdStyleAttr.Split(';'))
        {
            var d = decl.Trim();
            // strip the leading ` style="` from the first declaration
            var eq = d.IndexOf("style=\"", StringComparison.Ordinal);
            if (eq >= 0) d = d.Substring(eq + 7).Trim();
            d = d.TrimEnd('"').Trim();
            if (d.StartsWith("text-decoration:", StringComparison.Ordinal)
                || d.StartsWith("text-decoration-style:", StringComparison.Ordinal))
                sb.Append(';').Append(d);
        }
        return sb.ToString();
    }

    // ==================== Text spill (Excel overflow into empty neighbours) ====================
    //
    // Real Excel renders a long text cell's overflow across adjacent empty cells to
    // the right (for left/general alignment), clipping only at the first occupied
    // right-neighbour or the sheet edge. Returns the EXTRA pt budget (own column NOT
    // included — that is the td's normal width) the inline span may overflow into, or
    // 0 when the cell must clip at its own boundary (number, wrapText, occupied
    // neighbour, non-left alignment, etc.).
    // Estimate the row min-height (pt) a cell needs because its text is rotated to
    // (near-)vertical. Real Excel auto-expands the row so the full string shows; the
    // CSS transform:rotate alone does not. Returns 0 when the cell isn't (near-)
    // vertically rotated or has no text. Approximation only — like the spill/width
    // heuristics, the goal is "not clipped", matching Excel's auto-expand.
    private double EstimateRotatedCellHeightPt(Cell? cell, Stylesheet? stylesheet, double defaultFontPt)
    {
        if (cell == null || stylesheet?.CellFormats == null) return 0;
        var si = (int)(cell.StyleIndex?.Value ?? 0);
        if (si >= stylesheet.CellFormats.Elements<CellFormat>().Count()) return 0;
        var xf = stylesheet.CellFormats.Elements<CellFormat>().ElementAt(si);
        var rot = xf.Alignment?.TextRotation?.Value;
        // Only steep rotations (near-vertical) materially grow the row. Excel:
        // 1–90 = CCW, 91–180 = CW, 255 = stacked vertical. Treat >=75° / 165–180 /
        // 255 as vertical enough to need height.
        bool vertical = rot.HasValue &&
            (rot.Value == 255 || (rot.Value >= 75 && rot.Value <= 105) || rot.Value >= 165);
        if (!vertical) return 0;

        // Cell text length: shared-string / inline-string / raw value.
        string text = GetFormattedCellValue(cell, stylesheet);
        if (string.IsNullOrEmpty(text)) return 0;

        // Font size for this cell.
        double fontPt = defaultFontPt;
        var fontId = xf.FontId?.Value ?? 0;
        if (stylesheet.Fonts != null && fontId < (uint)stylesheet.Fonts.Elements<Font>().Count())
            fontPt = stylesheet.Fonts.Elements<Font>().ElementAt((int)fontId).FontSize?.Val?.Value ?? defaultFontPt;

        // Vertical text stacks glyphs along the column: extent ≈ chars × glyph advance.
        // ~0.62em per glyph advance matches the spill width heuristic's char model;
        // clamp so a single huge string doesn't blow up the layout.
        var extent = text.Length * fontPt * 0.62 + fontPt;
        return Math.Min(extent, 600.0);
    }

    private double GetSpillWidthPt(SheetRenderContext ctx, Cell? cell, string value, int r, int c)
    {
        if (cell == null || string.IsNullOrEmpty(value)) return 0;

        // (a) Must be text/general — NOT a number and NOT a numeric formula result.
        // Excel right-aligns numbers and never spills them.
        var dt = cell.DataType?.Value;
        bool isText = dt == CellValues.SharedString || dt == CellValues.InlineString || dt == CellValues.String;
        bool isBoolOrError = dt == CellValues.Boolean || dt == CellValues.Error;
        // Formula or general number: if it has a CellValue and isn't a string type,
        // treat as numeric (Excel does not spill numbers). Booleans/errors centre and
        // also don't spill.
        if (!isText || isBoolOrError) return 0;

        // (b) Resolve alignment + wrapText from the cell's xf.
        bool wrapText = false;
        string? hAlign = null;
        if (ctx.Stylesheet?.CellFormats != null)
        {
            var si = (int)(cell.StyleIndex?.Value ?? 0);
            var xfs = ctx.Stylesheet.CellFormats.Elements<CellFormat>().ToList();
            if (si >= 0 && si < xfs.Count)
            {
                var al = xfs[si].Alignment;
                wrapText = al?.WrapText?.Value == true;
                if (al?.Horizontal?.HasValue == true) hAlign = al.Horizontal.InnerText;
            }
        }
        if (wrapText) return 0; // wrapped cells clip/wrap, never spill

        // (c) Only left/general aligned text spills to the right (the common case).
        // right/center/justify/fill keep clipping (handled as follow-up).
        if (hAlign != null && hAlign != "left" && hAlign != "general" && hAlign != "fill")
            return 0;

        // (d) Sum widths of the run of EMPTY right-neighbours, stopping at the first
        // occupied cell, a merged region, a hidden column, or the sheet edge.
        double extra = 0;
        for (int nc = c + 1; nc <= ctx.MaxCol; nc++)
        {
            if (ctx.HiddenCols.Contains(nc)) break;
            var neighbourRef = $"{IndexToColumnName(nc)}{r}";
            if (ctx.MergeMap.ContainsKey(neighbourRef)) break;
            bool occupied = ctx.CellMap.TryGetValue((r, nc), out var ncell)
                            && ncell != null
                            && !string.IsNullOrEmpty(GetCellDisplayValue(ncell));
            if (occupied) break;
            extra += ctx.ColWidths.TryGetValue(nc, out var w) ? w : ctx.DefaultColWidthPt;
        }
        if (extra <= 0) return 0;

        // Own column width + the empty-neighbour run = the span's max overflow budget.
        double ownWidth = ctx.ColWidths.TryGetValue(c, out var ow) ? ow : ctx.DefaultColWidthPt;
        return ownWidth + extra;
    }

    // ==================== Shrink-to-fit (Excel font compression) ====================
    //
    // A cell with shrinkToFit alignment whose text overflows the column is rendered
    // by Excel with the font SHRUNK so the WHOLE string fits — never truncated. We
    // approximate the shrink: estimate text width = chars × fontPt × 0.62 (the same
    // glyph-advance model as the spill heuristic), then scale = colWidthPt / textWidth
    // (clamped to ≤1 and a sane floor). Returns the reduced font-size in pt, or 0 when
    // the cell isn't shrinkToFit or already fits (no shrink needed).
    private double GetShrinkToFitFontPt(SheetRenderContext ctx, Cell? cell, string value, int c)
    {
        if (cell == null || string.IsNullOrEmpty(value)) return 0;
        if (ctx.Stylesheet?.CellFormats == null) return 0;

        var si = (int)(cell.StyleIndex?.Value ?? 0);
        var xfs = ctx.Stylesheet.CellFormats.Elements<CellFormat>().ToList();
        if (si < 0 || si >= xfs.Count) return 0;
        var xf = xfs[si];
        if (xf.Alignment?.ShrinkToFit?.Value != true) return 0;
        if (xf.Alignment?.WrapText?.Value == true) return 0; // wrap takes precedence

        // Base font size for this cell.
        double basePt = 11.0;
        var fontId = xf.FontId?.Value ?? 0;
        if (ctx.Stylesheet.Fonts != null && fontId < (uint)ctx.Stylesheet.Fonts.Elements<Font>().Count())
            basePt = ctx.Stylesheet.Fonts.Elements<Font>().ElementAt((int)fontId).FontSize?.Val?.Value ?? 11.0;

        double colWidthPt = ctx.ColWidths.TryGetValue(c, out var w) ? w : ctx.DefaultColWidthPt;
        if (colWidthPt <= 0) return 0;

        // Estimated text width at the base font, minus a little cell padding (~3pt).
        double textWidthPt = value.Length * basePt * 0.62;
        double avail = Math.Max(colWidthPt - 3.0, 1.0);
        if (textWidthPt <= avail) return 0; // already fits — no shrink

        double scale = avail / textWidthPt;
        double shrunk = basePt * scale;
        return Math.Max(shrunk, 3.0); // floor so it stays legible
    }

    // CONSISTENCY(excel-virt): Private ExcelHandler.HtmlPreview.Virt.cs implements
    // OnGetVirtScript to load watch-overlay-virt.js from embedded resources.
    // When no private implementation exists the partial call is removed and result
    // stays empty (no virtualisation script injected).
    partial void OnGetVirtScript(ref string result);

    internal string GetVirtScript()
    {
        var result = string.Empty;
        OnGetVirtScript(ref result);
        return result;
    }

    private Dictionary<string, MergeInfo> BuildMergeMap(Worksheet ws)
    {
        var map = new Dictionary<string, MergeInfo>(StringComparer.OrdinalIgnoreCase);
        var mergeCells = ws.GetFirstChild<MergeCells>();
        if (mergeCells == null) return map;

        foreach (var mc in mergeCells.Elements<MergeCell>())
        {
            var rangeRef = mc.Reference?.Value;
            if (string.IsNullOrEmpty(rangeRef) || !rangeRef.Contains(':')) continue;

            var parts = rangeRef.Split(':');
            var (startCol, startRow) = ParseCellReference(parts[0]);
            var (endCol, endRow) = ParseCellReference(parts[1]);
            var startColIdx = ColumnNameToIndex(startCol);
            var endColIdx = ColumnNameToIndex(endCol);
            // Clamp merge range to rendering limits to prevent memory explosion
            var clampedEndRow = Math.Min(endRow, 5000);
            var clampedEndCol = Math.Min(endColIdx, 200);
            var rowSpan = clampedEndRow - startRow + 1;
            var colSpan = clampedEndCol - startColIdx + 1;

            for (int r = startRow; r <= clampedEndRow; r++)
            {
                for (int ci = startColIdx; ci <= clampedEndCol; ci++)
                {
                    var cellRef = $"{IndexToColumnName(ci)}{r}";
                    bool isAnchor = (r == startRow && ci == startColIdx);
                    map[cellRef] = new MergeInfo(isAnchor, isAnchor ? rowSpan : 0, isAnchor ? colSpan : 0);
                }
            }
        }

        return map;
    }

    // ==================== Column Widths ====================

    private static Dictionary<int, double> GetColumnWidths(Worksheet ws)
    {
        var result = new Dictionary<int, double>();
        var columns = ws.GetFirstChild<Columns>();
        if (columns == null) return result;

        foreach (var col in columns.Elements<Column>())
        {
            if (col.Width?.Value == null) continue;
            var min = (int)(col.Min?.Value ?? 1u);
            var max = (int)(col.Max?.Value ?? (uint)min);
            // Hidden columns get width 0
            // Excel column width → pixels: chars * 7.0017; pt = px * 0.75
            var widthPt = col.Hidden?.Value == true ? 0 : (col.Width.Value == 0 ? 0 : col.Width.Value * 7.0017 * 0.75);
            for (int c = min; c <= max; c++)
                result[c] = widthPt;
        }

        return result;
    }

    // ==================== Frozen Panes ====================

    private static (int frozenRows, int frozenCols) GetFrozenPanes(Worksheet ws)
    {
        var sheetViews = ws.GetFirstChild<SheetViews>();
        var sheetView = sheetViews?.GetFirstChild<SheetView>();
        var pane = sheetView?.GetFirstChild<Pane>();
        if (pane == null) return (0, 0);

        // Only handle frozen panes (not split panes)
        if (pane.State?.Value != PaneStateValues.Frozen && pane.State?.Value != PaneStateValues.FrozenSplit)
            return (0, 0);

        var frozenRows = (int)(pane.VerticalSplit?.Value ?? 0);
        var frozenCols = (int)(pane.HorizontalSplit?.Value ?? 0);
        return (frozenRows, frozenCols);
    }

    // ==================== Conditional Formatting ====================

    /// <summary>
    /// Evaluate conditional formatting rules and return CSS overrides per cell.
    /// </summary>
    private Dictionary<string, string> BuildConditionalFormatMap(
        Worksheet ws, Stylesheet? stylesheet, SheetData sheetData, WorkbookPart? workbookPart)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        var cfElements = ws.Elements<ConditionalFormatting>().ToList();
        if (cfElements.Count == 0) return result;

        // Color-scale rules carry inline stop colors (no dxfId, no stylesheet);
        // render them in a separate pass so they work even when the workbook has
        // no <dxfs> or no stylesheet at all.
        AddColorScaleBackgrounds(cfElements, sheetData, result);

        // The remaining dxf-indexed rules need the stylesheet's <dxfs> catalogue.
        var dxfs = stylesheet?.DifferentialFormats?.Elements<DifferentialFormat>().ToArray();
        if (dxfs == null || dxfs.Length == 0) return result;

        var evaluator = new Core.FormulaEvaluator(sheetData, workbookPart);
        var (cfBoundRow, cfBoundCol) = UsedBounds(sheetData);
        // Blank-sensitive rules (containsBlanks / notContains*) must evaluate
        // cells beyond the used data extent — those cells render too (the grid
        // dimensions are extended the same way). Clamp the CF contribution with
        // the same caps UsedBounds applies so a whole-column sqref stays bounded.
        var (cfExtRow, cfExtCol) = CfRangeExtents(cfElements);
        cfBoundRow = Math.Max(cfBoundRow, Math.Min(cfExtRow, GetHtmlRowCap()));
        cfBoundCol = Math.Max(cfBoundCol, Math.Min(cfExtCol, 200));

        foreach (var cf in cfElements)
        {
            var sqref = cf.SequenceOfReferences?.Items?.ToList();
            if (sqref == null || sqref.Count == 0) continue;

            foreach (var rule in cf.Elements<ConditionalFormattingRule>())
            {
                var dxfId = rule.FormatId?.Value;
                if (dxfId == null || dxfId >= dxfs.Length) continue;
                var dxf = dxfs[(int)dxfId];

                // Extract CSS from dxf
                var cssParts = new List<string>();
                var fill = dxf.Fill?.PatternFill;
                if (fill != null)
                {
                    var bgColor = fill.BackgroundColor?.Rgb?.Value ?? fill.ForegroundColor?.Rgb?.Value;
                    if (bgColor != null)
                    {
                        if (bgColor.Length > 6) bgColor = bgColor[^6..];
                        cssParts.Add($"background:#{bgColor}");
                    }
                }
                var font = dxf.Font;
                if (font != null)
                {
                    var fontColor = font.Color?.Rgb?.Value;
                    if (fontColor != null)
                    {
                        if (fontColor.Length > 6) fontColor = fontColor[^6..];
                        cssParts.Add($"color:#{fontColor}");
                    }
                    // A dxf font may also carry bold/italic/underline/strike — Excel
                    // applies these on top of the color/fill for the matched cell.
                    var bEl = font.GetFirstChild<Bold>();
                    if (bEl != null && (bEl.Val == null || bEl.Val.Value))
                        cssParts.Add("font-weight:bold");
                    var iEl = font.GetFirstChild<Italic>();
                    if (iEl != null && (iEl.Val == null || iEl.Val.Value))
                        cssParts.Add("font-style:italic");
                    var uEl = font.GetFirstChild<Underline>();
                    bool hasUnderline = uEl != null && uEl.Val?.Value != UnderlineValues.None;
                    var sEl = font.GetFirstChild<Strike>();
                    bool hasStrike = sEl != null && (sEl.Val == null || sEl.Val.Value);
                    if (hasUnderline && hasStrike) cssParts.Add("text-decoration:underline line-through");
                    else if (hasUnderline) cssParts.Add("text-decoration:underline");
                    else if (hasStrike) cssParts.Add("text-decoration:line-through");
                }
                if (cssParts.Count == 0) continue;
                var cssOverride = string.Join(";", cssParts);

                // Expand sqref and evaluate each cell
                foreach (var rangeStr in sqref)
                {
                    var cells = ExpandSqref(rangeStr.Value ?? "", cfBoundRow, cfBoundCol);
                    foreach (var (cellRef, row, col) in cells)
                    {
                        if (result.ContainsKey(cellRef)) continue; // first matching rule wins

                        bool matches = EvaluateCfRule(rule, cellRef, row, col, sheetData, evaluator);
                        if (matches)
                            result[cellRef] = cssOverride;
                    }
                }
            }
        }
        return result;
    }

    /// <summary>
    /// Color-scale CF rules (type="colorScale") carry inline &lt;cfvo&gt; stops and a
    /// matching list of &lt;color&gt; children; they never use a dxfId. Interpolate a
    /// per-cell background between the stop colors (2-stop min→max, or 3-stop
    /// min→mid→max) based on each cell's value, and write "background:#RRGGBB" into
    /// the shared CF map so the existing per-cell apply path picks it up.
    /// First matching rule wins, consistent with the dxf loop.
    /// </summary>
    private void AddColorScaleBackgrounds(
        List<ConditionalFormatting> cfElements, SheetData sheetData, Dictionary<string, string> result)
    {
        var (cfBoundRow, cfBoundCol) = UsedBounds(sheetData);
        foreach (var cf in cfElements)
        {
            var sqref = cf.SequenceOfReferences?.Items?.ToList();
            if (sqref == null || sqref.Count == 0) continue;

            foreach (var rule in cf.Elements<ConditionalFormattingRule>())
            {
                var colorScale = rule.GetFirstChild<ColorScale>();
                if (colorScale == null) continue;

                var stops = colorScale.Elements<DocumentFormat.OpenXml.Spreadsheet.Color>()
                    .Select(c => NormalizeScaleColor(c.Rgb?.Value))
                    .ToList();
                if (stops.Count < 2) continue;

                // Collect numeric cell values in range to derive min/max anchors.
                var cells = new List<(string cellRef, double value)>();
                foreach (var rangeStr in sqref)
                {
                    foreach (var (cellRef, row, col) in ExpandSqref(rangeStr.Value ?? "", cfBoundRow, cfBoundCol))
                    {
                        var cell = sheetData.Descendants<Cell>()
                            .FirstOrDefault(c => string.Equals(c.CellReference?.Value, cellRef, StringComparison.OrdinalIgnoreCase));
                        if (cell?.CellValue != null && double.TryParse(cell.CellValue.Text,
                            System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var v))
                            cells.Add((cellRef, v));
                    }
                }
                if (cells.Count == 0) continue;

                double minVal = cells.Min(c => c.value);
                double maxVal = cells.Max(c => c.value);
                if (maxVal <= minVal) maxVal = minVal + 1;

                foreach (var (cellRef, value) in cells)
                {
                    if (result.ContainsKey(cellRef)) continue; // first matching rule wins
                    var t = (value - minVal) / (maxVal - minVal);
                    var rgb = InterpolateColorScale(stops, t);
                    result[cellRef] = $"background:#{rgb}";
                }
            }
        }
    }

    /// <summary>Strip an optional 8-hex ARGB prefix down to 6-hex RRGGBB; default black.</summary>
    private static string NormalizeScaleColor(string? argb)
    {
        if (string.IsNullOrEmpty(argb)) return "000000";
        return argb.Length > 6 ? argb[^6..] : argb;
    }

    /// <summary>
    /// Linearly interpolate across an ordered list of RRGGBB stops by fraction
    /// t in [0,1]. Two stops → single segment; three stops → min/mid(0.5)/max.
    /// </summary>
    private static string InterpolateColorScale(List<string> stops, double t)
    {
        t = Math.Max(0, Math.Min(1, t));
        // Map t onto the segment between stop[i] and stop[i+1].
        int segCount = stops.Count - 1;
        double scaled = t * segCount;
        int lo = Math.Min((int)scaled, segCount - 1);
        double frac = scaled - lo;
        var (r1, g1, b1) = HexToRgb(stops[lo]);
        var (r2, g2, b2) = HexToRgb(stops[lo + 1]);
        int r = (int)Math.Round(r1 + (r2 - r1) * frac);
        int g = (int)Math.Round(g1 + (g2 - g1) * frac);
        int b = (int)Math.Round(b1 + (b2 - b1) * frac);
        return $"{r:X2}{g:X2}{b:X2}";
    }

    private static (int r, int g, int b) HexToRgb(string hex)
    {
        int r = Convert.ToInt32(hex.Substring(0, 2), 16);
        int g = Convert.ToInt32(hex.Substring(2, 2), 16);
        int b = Convert.ToInt32(hex.Substring(4, 2), 16);
        return (r, g, b);
    }

    /// <summary>
    /// Build data bar info per cell: returns HTML for the bar overlay.
    /// </summary>
    private Dictionary<string, string> BuildDataBarMap(Worksheet ws, SheetData sheetData)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var (cfBoundRow, cfBoundCol) = UsedBounds(sheetData);
        foreach (var cf in ws.Elements<ConditionalFormatting>())
        {
            foreach (var rule in cf.Elements<ConditionalFormattingRule>())
            {
                var dataBar = rule.GetFirstChild<DataBar>();
                if (dataBar == null) continue;

                var sqref = cf.SequenceOfReferences?.Items?.ToList();
                if (sqref == null || sqref.Count == 0) continue;

                // Get bar color
                var barColorEl = dataBar.GetFirstChild<Color>();
                var barColor = barColorEl?.Rgb?.Value ?? "FF4472C4";
                if (barColor.Length > 6) barColor = barColor[^6..];

                // Collect all cell values in range
                var cells = new List<(string cellRef, double value)>();
                foreach (var rangeStr in sqref)
                {
                    foreach (var (cellRef, row, col) in ExpandSqref(rangeStr.Value ?? "", cfBoundRow, cfBoundCol))
                    {
                        var cell = sheetData.Descendants<Cell>()
                            .FirstOrDefault(c => string.Equals(c.CellReference?.Value, cellRef, StringComparison.OrdinalIgnoreCase));
                        if (cell?.CellValue != null && double.TryParse(cell.CellValue.Text,
                            System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var v))
                            cells.Add((cellRef, v));
                    }
                }
                if (cells.Count == 0) continue;

                // Determine min/max from cfvo elements or from data
                var cfvos = dataBar.Elements<ConditionalFormatValueObject>().ToList();
                double minVal, maxVal;
                var dataMin = cells.Min(c => c.value);
                if (cfvos.Count >= 2 && cfvos[0].Type?.Value == ConditionalFormatValueObjectValues.Number
                    && double.TryParse(cfvos[0].Val?.Value, System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture, out var explicitMin))
                    minVal = explicitMin;
                else
                    // Excel's automatic data-bar scaling anchors the LOW end at
                    // the data minimum (not 0): the smallest value gets a tiny sliver
                    // and the largest gets a full bar, linear in (v-min)/(max-min).
                    // The earlier "anchor at 0" model over-inflated small values
                    // (e.g. 20 of 20..100 rendered 26% instead of a few %).
                    // When negatives are present the floor is still the data minimum,
                    // and the zero-axis split (below) draws left/right bars from 0.
                    minVal = dataMin;

                if (cfvos.Count >= 2 && cfvos[1].Type?.Value == ConditionalFormatValueObjectValues.Number
                    && double.TryParse(cfvos[1].Val?.Value, System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture, out var explicitMax))
                    maxVal = explicitMax;
                else
                    maxVal = cells.Max(c => c.value);

                if (maxVal <= minVal) maxVal = minVal + 1;

                // Read bar length bounds. When the data bar carries no explicit
                // min/max length, the smallest value should show only a sliver and
                // the largest should fill the cell (matching real Excel automatic
                // data bars), so default to a tiny floor and a full-width cap rather
                // than the old 10%/90% band that inflated the minimum.
                var minLength = dataBar.MinLength?.Value ?? 3U;
                var maxLength = dataBar.MaxLength?.Value ?? 100U;
                var showValue = dataBar.ShowValue?.Value ?? true;

                // R17a: when the range straddles zero, draw a zero-axis and split
                // bars left/right of it. zeroPct is the axis position (0–100%).
                bool hasNegative = minVal < 0;
                var zeroPct = hasNegative ? (0 - minVal) / (maxVal - minVal) * 100 : 0;

                foreach (var (cellRef, value) in cells)
                {
                    string barDiv;
                    if (hasNegative)
                    {
                        // Width proportional to |value| over the full span; positive
                        // bars extend right from the zero-axis, negative bars left.
                        var wPct = Math.Min(100, Math.Abs(value) / (maxVal - minVal) * 100);
                        if (value >= 0)
                            barDiv = $"<div style=\"position:absolute;left:{zeroPct:0.#}%;top:1px;bottom:1px;width:{wPct:0.#}%;background:linear-gradient(to right,#{barColor},#{barColor}40);border-radius:1px\"></div>";
                        else
                            barDiv = $"<div style=\"position:absolute;left:{Math.Max(0, zeroPct - wPct):0.#}%;top:1px;bottom:1px;width:{wPct:0.#}%;background:linear-gradient(to left,#{barColor},#{barColor}40);border-radius:1px\"></div>";
                        // Zero-axis marker (thin line) — drawn once-style per cell is fine.
                        barDiv += $"<div style=\"position:absolute;left:{zeroPct:0.#}%;top:0;bottom:0;border-left:1px dashed #c0504d\"></div>";
                    }
                    else
                    {
                        var rawPct = (value - minVal) / (maxVal - minVal) * 100;
                        // Scale to minLength..maxLength range
                        var pct = Math.Max(0, Math.Min(100, minLength + rawPct / 100 * (maxLength - minLength)));
                        barDiv = $"<div style=\"position:absolute;left:0;top:1px;bottom:1px;width:{pct:0.#}%;background:linear-gradient(to right,#{barColor},#{barColor}40);border-radius:1px\"></div>";
                    }
                    // Store bar HTML + showValue flag (prefixed with "0|" or "1|")
                    result[cellRef] = $"{(showValue ? "1" : "0")}|{barDiv}";
                }
            }
        }
        return result;
    }

    /// <summary>
    /// Build icon set info per cell: returns HTML for the icon.
    /// </summary>
    private Dictionary<string, string> BuildIconSetMap(Worksheet ws, SheetData sheetData)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var (cfBoundRow, cfBoundCol) = UsedBounds(sheetData);
        foreach (var cf in ws.Elements<ConditionalFormatting>())
        {
            foreach (var rule in cf.Elements<ConditionalFormattingRule>())
            {
                var iconSet = rule.GetFirstChild<IconSet>();
                if (iconSet == null) continue;

                var sqref = cf.SequenceOfReferences?.Items?.ToList();
                if (sqref == null || sqref.Count == 0) continue;

                var iconSetName = iconSet.IconSetValue?.Value ?? IconSetValues.ThreeTrafficLights1;
                var showValue = iconSet.ShowValue?.Value ?? true;
                var reverse = iconSet.Reverse?.Value ?? false;

                // Collect all cell values in range
                var cells = new List<(string cellRef, double value)>();
                foreach (var rangeStr in sqref)
                {
                    foreach (var (cellRef, row, col) in ExpandSqref(rangeStr.Value ?? "", cfBoundRow, cfBoundCol))
                    {
                        var cell = sheetData.Descendants<Cell>()
                            .FirstOrDefault(c => string.Equals(c.CellReference?.Value, cellRef, StringComparison.OrdinalIgnoreCase));
                        if (cell?.CellValue != null && double.TryParse(cell.CellValue.Text,
                            System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var v))
                            cells.Add((cellRef, v));
                    }
                }
                if (cells.Count == 0) continue;

                // Parse cfvo thresholds
                var cfvos = iconSet.Elements<ConditionalFormatValueObject>().ToList();
                var allValues = cells.Select(c => c.value).OrderBy(v => v).ToList();
                double minVal = allValues.First(), maxVal = allValues.Last();
                var range = maxVal - minVal;
                if (range == 0) range = 1;

                // Resolve thresholds (skip first cfvo which is the base)
                var thresholds = new List<double>();
                for (int i = 1; i < cfvos.Count; i++)
                {
                    var cfvo = cfvos[i];
                    var type = cfvo.Type?.Value ?? ConditionalFormatValueObjectValues.Percent;
                    double.TryParse(cfvo.Val?.Value, System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture, out var tv);
                    if (type == ConditionalFormatValueObjectValues.Number)
                        thresholds.Add(tv);
                    else if (type == ConditionalFormatValueObjectValues.Percent)
                        thresholds.Add(minVal + range * tv / 100);
                    else if (type == ConditionalFormatValueObjectValues.Percentile)
                    {
                        var idx = (int)Math.Round(tv / 100.0 * (allValues.Count - 1));
                        thresholds.Add(allValues[Math.Clamp(idx, 0, allValues.Count - 1)]);
                    }
                    else
                        thresholds.Add(minVal + range * tv / 100);
                }

                foreach (var (cellRef, value) in cells)
                {
                    // Determine which bucket the value falls into
                    int bucket = 0;
                    for (int i = 0; i < thresholds.Count; i++)
                    {
                        if (value >= thresholds[i]) bucket = i + 1;
                    }
                    if (reverse) bucket = cfvos.Count - 1 - bucket;
                    var icon = GetIconHtml(iconSetName, bucket, cfvos.Count);
                    // Prefix with showValue flag: "0|" = hide value, "1|" = show value
                    result[cellRef] = $"{(showValue ? "1" : "0")}|{icon}";
                }
            }
        }
        return result;
    }

    /// <summary>
    /// R12a: build a cellRef → inline-SVG map for sparklines. Sparkline groups
    /// live in the worksheet extension list (x14:sparklineGroups); each
    /// x14:sparkline carries a Formula (data range) and a ReferenceSequence
    /// (the host cell). We render a small SVG into that host cell — column =
    /// proportional bars, line/stacked = polyline — so the cell is no longer
    /// blank. Lightweight approximation of the native in-cell mini chart.
    /// </summary>
    private Dictionary<string, string> BuildSparklineMap(Worksheet ws)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var extList = ws.GetFirstChild<WorksheetExtensionList>();
        if (extList == null) return result;
        foreach (var group in extList.Descendants<X14.SparklineGroup>())
        {
            var type = group.Type?.Value;
            var kind = type == X14.SparklineTypeValues.Column ? "column"
                     : type == X14.SparklineTypeValues.Stacked ? "stacked"
                     : "line";
            // Series color: mirror the reader (Helpers.Node) — fall back to
            // default blue only when no <x14:colorSeries rgb=...> is stored.
            var seriesRgb = group.SeriesColor?.Rgb?.Value;
            var seriesColor = seriesRgb != null
                ? ParseHelpers.FormatHexColor(seriesRgb)
                : "#4472C4";
            foreach (var spk in group.Descendants<X14.Sparkline>())
            {
                var dataRange = spk.Formula?.Text;
                var hostCell = spk.ReferenceSequence?.Text;
                if (string.IsNullOrEmpty(dataRange) || string.IsNullOrEmpty(hostCell)) continue;
                var values = ReadCellRangeAsDoubles(dataRange);
                if (values == null || values.Length == 0) continue;
                // host cell may be a range like "F1:F1" — take the first ref.
                var host = hostCell.Contains(':') ? hostCell.Split(':')[0] : hostCell;
                host = host.Contains('!') ? host.Split('!')[1] : host;
                result[host] = BuildSparklineSvg(values, kind, seriesColor);
            }
        }
        return result;
    }

    /// <summary>Render a sparkline's values as a small inline SVG (~80x20px).</summary>
    private static string BuildSparklineSvg(double[] values, string kind, string seriesColor)
    {
        const double w = 80, h = 18;
        var min = values.Min();
        var max = values.Max();
        var range = max - min;
        if (range == 0) range = 1;
        // Baseline at zero when data straddles it, else at the value floor.
        double zeroFloor = Math.Min(min, 0);
        double zeroRange = Math.Max(max, 0) - zeroFloor;
        if (zeroRange == 0) zeroRange = 1;
        var sb = new StringBuilder();
        sb.Append($"<svg class=\"sparkline\" width=\"{w:0}\" height=\"{h:0}\" viewBox=\"0 0 {w:0} {h:0}\" preserveAspectRatio=\"none\" style=\"vertical-align:middle\">");
        if (kind == "column" || kind == "stacked")
        {
            int n = values.Length;
            double bw = w / n;
            for (int i = 0; i < n; i++)
            {
                var v = values[i];
                // bar from the zero line; positive up, negative down.
                var zeroY = h - (0 - zeroFloor) / zeroRange * h;
                var valY = h - (v - zeroFloor) / zeroRange * h;
                var top = Math.Min(zeroY, valY);
                var bh = Math.Max(1, Math.Abs(valY - zeroY));
                var color = v < 0 ? "#C0504D" : seriesColor;
                sb.Append($"<rect x=\"{i * bw + 0.5:0.#}\" y=\"{top:0.#}\" width=\"{Math.Max(1, bw - 1):0.#}\" height=\"{bh:0.#}\" fill=\"{color}\"/>");
            }
        }
        else
        {
            int n = values.Length;
            var pts = new List<string>();
            for (int i = 0; i < n; i++)
            {
                var x = n > 1 ? (double)i / (n - 1) * w : w / 2;
                var y = h - (values[i] - min) / range * h;
                pts.Add($"{x:0.#},{y:0.#}");
            }
            sb.Append($"<polyline points=\"{string.Join(" ", pts)}\" fill=\"none\" stroke=\"{seriesColor}\" stroke-width=\"1\"/>");
        }
        sb.Append("</svg>");
        return sb.ToString();
    }

    private static string GetIconHtml(IconSetValues iconSetName, int bucket, int totalBuckets)
    {
        // Traffic lights: red=0, yellow=1, green=2
        if (iconSetName == IconSetValues.ThreeTrafficLights1 || iconSetName == IconSetValues.ThreeTrafficLights2)
        {
            var color = bucket switch { 0 => "#C00000", 1 => "#FFC000", _ => "#00B050" };
            return $"<span style=\"display:inline-block;width:10px;height:10px;border-radius:50%;background:{color};margin-right:4px;vertical-align:middle\"></span>";
        }
        // Arrows
        if (iconSetName == IconSetValues.ThreeArrows || iconSetName == IconSetValues.ThreeArrowsGray)
        {
            return bucket switch
            {
                0 => "<span style=\"color:#C00000;margin-right:4px;vertical-align:middle\">&#x25BC;</span>",
                1 => "<span style=\"color:#FFC000;margin-right:4px;vertical-align:middle\">&#x25B6;</span>",
                _ => "<span style=\"color:#00B050;margin-right:4px;vertical-align:middle\">&#x25B2;</span>",
            };
        }
        // R17b: directional arrow icon sets (4/5 arrows). Native renders
        // graduated arrows ↓↘→↗↑; previously these fell through to the default
        // colored circle. Glyphs: ↓ U+2193, ↘ U+2198, → U+2192, ↗ U+2197, ↑ U+2191.
        if (iconSetName == IconSetValues.FiveArrows || iconSetName == IconSetValues.FiveArrowsGray)
        {
            var gray = iconSetName == IconSetValues.FiveArrowsGray;
            var glyph = bucket switch { 0 => "&#x2193;", 1 => "&#x2198;", 2 => "&#x2192;", 3 => "&#x2197;", _ => "&#x2191;" };
            var color = gray ? "#808080" : bucket switch { 0 => "#C00000", 1 => "#E08000", 2 => "#FFC000", 3 => "#92D050", _ => "#00B050" };
            return $"<span style=\"color:{color};margin-right:4px;vertical-align:middle\">{glyph}</span>";
        }
        if (iconSetName == IconSetValues.FourArrows || iconSetName == IconSetValues.FourArrowsGray)
        {
            var gray = iconSetName == IconSetValues.FourArrowsGray;
            var glyph = bucket switch { 0 => "&#x2193;", 1 => "&#x2198;", 2 => "&#x2197;", _ => "&#x2191;" };
            var color = gray ? "#808080" : bucket switch { 0 => "#C00000", 1 => "#E08000", 2 => "#92D050", _ => "#00B050" };
            return $"<span style=\"color:{color};margin-right:4px;vertical-align:middle\">{glyph}</span>";
        }
        // 4-icon traffic lights
        if (iconSetName == IconSetValues.FourTrafficLights)
        {
            var color = bucket switch { 0 => "#C00000", 1 => "#FFC000", 2 => "#92D050", _ => "#00B050" };
            return $"<span style=\"display:inline-block;width:10px;height:10px;border-radius:50%;background:{color};margin-right:4px;vertical-align:middle\"></span>";
        }
        // Default: colored circles
        if (totalBuckets <= 3)
        {
            var color = bucket switch { 0 => "#C00000", 1 => "#FFC000", _ => "#00B050" };
            return $"<span style=\"display:inline-block;width:10px;height:10px;border-radius:50%;background:{color};margin-right:4px;vertical-align:middle\"></span>";
        }
        else
        {
            var pct = totalBuckets > 1 ? (double)bucket / (totalBuckets - 1) : 1;
            var r = (int)(0xC0 * (1 - pct));
            var g = (int)(0xB0 * pct);
            var color = $"#{r:X2}{g:X2}00";
            return $"<span style=\"display:inline-block;width:10px;height:10px;border-radius:50%;background:{color};margin-right:4px;vertical-align:middle\"></span>";
        }
    }

    /// <summary>Evaluate whether a conditional formatting rule matches a specific cell.</summary>
    private bool EvaluateCfRule(ConditionalFormattingRule rule, string cellRef, int row, int col,
        SheetData sheetData, Core.FormulaEvaluator evaluator)
    {
        var ruleType = rule.Type?.Value;

        // Get cell value for comparison
        double? cellValue = null;
        var cell = sheetData.Descendants<Cell>()
            .FirstOrDefault(c => string.Equals(c.CellReference?.Value, cellRef, StringComparison.OrdinalIgnoreCase));
        if (cell != null)
        {
            if (double.TryParse(cell.CellValue?.Text, System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out var v))
                cellValue = v;
        }

        if (ruleType == ConditionalFormatValues.Expression)
        {
            // Formula-based rule: evaluate with cell reference adjustment
            var formula = rule.Elements<Formula>().FirstOrDefault()?.Text;
            if (string.IsNullOrEmpty(formula)) return false;

            // Adjust formula references relative to the first cell in sqref
            // The formula is written for the top-left cell; adjust for current cell
            var adjusted = AdjustCfFormula(formula, row, col, rule);
            var result = evaluator.TryEvaluateFull(adjusted);
            return result?.BoolValue == true || (result?.NumericValue != null && result.NumericValue != 0);
        }

        if (ruleType == ConditionalFormatValues.CellIs)
        {
            var op = rule.Operator?.Value;
            var f1 = rule.Elements<Formula>().FirstOrDefault()?.Text;
            var f2 = rule.Elements<Formula>().Skip(1).FirstOrDefault()?.Text;
            double? v1 = f1 != null ? evaluator.TryEvaluate(f1) ?? (double.TryParse(f1, out var p1) ? p1 : null) : null;
            double? v2 = f2 != null ? evaluator.TryEvaluate(f2) ?? (double.TryParse(f2, out var p2) ? p2 : null) : null;

            // String comparison for equal/notEqual when the cell value and/or the
            // rule operand are non-numeric (Excel stores string operands as "Apple").
            if ((op == ConditionalFormattingOperatorValues.Equal || op == ConditionalFormattingOperatorValues.NotEqual)
                && (!cellValue.HasValue || v1 == null))
            {
                var hay = cell != null ? GetCellDisplayValue(cell) : "";
                var needle = (f1 ?? "").Trim();
                if (needle.Length >= 2 && needle.StartsWith("\"") && needle.EndsWith("\""))
                    needle = needle.Substring(1, needle.Length - 2);
                bool eq = string.Equals(hay, needle, StringComparison.OrdinalIgnoreCase);
                return op == ConditionalFormattingOperatorValues.Equal ? eq : !eq;
            }

            if (!cellValue.HasValue || v1 == null) return false;
            if (op == ConditionalFormattingOperatorValues.GreaterThan) return cellValue > v1;
            if (op == ConditionalFormattingOperatorValues.LessThan) return cellValue < v1;
            if (op == ConditionalFormattingOperatorValues.GreaterThanOrEqual) return cellValue >= v1;
            if (op == ConditionalFormattingOperatorValues.LessThanOrEqual) return cellValue <= v1;
            if (op == ConditionalFormattingOperatorValues.Equal) return cellValue == v1;
            if (op == ConditionalFormattingOperatorValues.NotEqual) return cellValue != v1;
            if (op == ConditionalFormattingOperatorValues.Between) return v2.HasValue && cellValue >= v1 && cellValue <= v2;
            if (op == ConditionalFormattingOperatorValues.NotBetween) return v2.HasValue && (cellValue < v1 || cellValue > v2);
            return false;
        }

        if (ruleType == ConditionalFormatValues.Top10 && cellValue.HasValue)
        {
            // Top/bottom N (or N%) of the numeric values in the rule's range.
            // Mirrors Excel: rank=N, percent=true means N% of cells, bottom=true
            // ranks ascending. A cell matches if it falls within that slice.
            var nums = CollectCfRangeNumbers(rule, sheetData);
            if (nums.Count == 0) return false;
            var rank = (int)(rule.Rank?.Value ?? 10);
            var bottom = rule.Bottom?.Value == true;
            var percent = rule.Percent?.Value == true;
            var count = percent
                ? (int)Math.Ceiling(nums.Count * (rank / 100.0))
                : rank;
            if (count < 1) count = 1;
            if (count > nums.Count) count = nums.Count;
            var ordered = bottom
                ? nums.OrderBy(n => n).ToList()
                : nums.OrderByDescending(n => n).ToList();
            // Threshold = the value at the Nth position; include ties (Excel
            // colors all cells at/beyond the cutoff value, like rank ties).
            var threshold = ordered[count - 1];
            return bottom ? cellValue <= threshold : cellValue >= threshold;
        }

        if (ruleType == ConditionalFormatValues.AboveAverage && cellValue.HasValue)
        {
            // Color cells above (or below) the numeric average of the rule's range.
            // The aboveAverage attribute defaults to true (omitted == above);
            // stdDev shifts the threshold by N standard deviations; equalAverage
            // includes values equal to the threshold.
            var nums = CollectCfRangeNumbers(rule, sheetData);
            if (nums.Count == 0) return false;
            var avg = nums.Average();
            var above = rule.AboveAverage?.Value ?? true;
            var equal = rule.EqualAverage?.Value ?? false;
            var threshold = avg;
            if (rule.StdDev?.Value is int sd && sd != 0)
            {
                var variance = nums.Select(n => (n - avg) * (n - avg)).Sum() / nums.Count;
                var stdDev = Math.Sqrt(variance);
                threshold = above ? avg + sd * stdDev : avg - sd * stdDev;
            }
            if (above) return equal ? cellValue >= threshold : cellValue > threshold;
            return equal ? cellValue <= threshold : cellValue < threshold;
        }

        if (ruleType == ConditionalFormatValues.ContainsText
            || ruleType == ConditionalFormatValues.NotContainsText
            || ruleType == ConditionalFormatValues.BeginsWith
            || ruleType == ConditionalFormatValues.EndsWith)
        {
            // Text-operator rules: compare the cell's displayed text against the
            // rule's <text> attribute. Case-insensitive, matching Excel.
            var needle = rule.Text?.Value ?? "";
            if (needle.Length == 0) return false;
            var hay = cell != null ? GetCellDisplayValue(cell) : "";
            var cmp = StringComparison.OrdinalIgnoreCase;
            if (ruleType == ConditionalFormatValues.ContainsText) return hay.Contains(needle, cmp);
            if (ruleType == ConditionalFormatValues.NotContainsText) return !hay.Contains(needle, cmp);
            if (ruleType == ConditionalFormatValues.BeginsWith) return hay.StartsWith(needle, cmp);
            if (ruleType == ConditionalFormatValues.EndsWith) return hay.EndsWith(needle, cmp);
            return false;
        }

        if (ruleType == ConditionalFormatValues.ContainsBlanks
            || ruleType == ConditionalFormatValues.NotContainsBlanks)
        {
            // A cell is "blank" when it does not exist, has no value, or its
            // displayed text is empty/whitespace (matching Excel's LEN(TRIM(...))=0).
            var disp = cell != null ? GetCellDisplayValue(cell) : "";
            bool isBlank = string.IsNullOrWhiteSpace(disp);
            return ruleType == ConditionalFormatValues.ContainsBlanks ? isBlank : !isBlank;
        }

        if (ruleType == ConditionalFormatValues.ContainsErrors
            || ruleType == ConditionalFormatValues.NotContainsErrors)
        {
            // A cell "contains an error" when its data type is error, or its
            // displayed/evaluated value is one of Excel's error literals.
            bool isError = cell?.DataType?.Value == CellValues.Error;
            if (!isError && cell != null)
                isError = IsExcelErrorText(GetFormattedCellValue(cell, null, evaluator));
            return ruleType == ConditionalFormatValues.ContainsErrors ? isError : !isError;
        }

        if (ruleType == ConditionalFormatValues.DuplicateValues || ruleType == ConditionalFormatValues.UniqueValues)
        {
            // Color cells whose value appears more than once (duplicateValues)
            // or exactly once (uniqueValues) within the rule's range. Compare
            // on the cell's display text (via GetCellDisplayValue) so inline
            // strings (<is><t>) and shared strings resolve too — not just the
            // <v> form; otherwise inlineStr cells read empty and never match.
            var thisText = cell != null ? GetCellDisplayValue(cell) : null;
            if (string.IsNullOrEmpty(thisText)) return false;
            var counts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (var t in CollectCfRangeTexts(rule, sheetData))
                counts[t] = counts.TryGetValue(t, out var c) ? c + 1 : 1;
            var occurrences = counts.TryGetValue(thisText, out var n) ? n : 0;
            return ruleType == ConditionalFormatValues.DuplicateValues
                ? occurrences > 1
                : occurrences == 1;
        }

        return false;
    }

    /// <summary>True when the text is one of Excel's seven error literals.</summary>
    private static bool IsExcelErrorText(string? text)
    {
        if (string.IsNullOrEmpty(text)) return false;
        return text is "#DIV/0!" or "#N/A" or "#VALUE!" or "#REF!"
            or "#NAME?" or "#NULL!" or "#NUM!";
    }

    /// <summary>Collect the numeric values of every cell in a CF rule's sqref range.</summary>
    private List<double> CollectCfRangeNumbers(ConditionalFormattingRule rule, SheetData sheetData)
    {
        var result = new List<double>();
        foreach (var c in CollectCfRangeCells(rule, sheetData))
        {
            if (double.TryParse(c.CellValue?.Text, System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out var v))
                result.Add(v);
        }
        return result;
    }

    /// <summary>Collect the raw text of every non-empty cell in a CF rule's sqref range.</summary>
    private List<string> CollectCfRangeTexts(ConditionalFormattingRule rule, SheetData sheetData)
    {
        var result = new List<string>();
        foreach (var c in CollectCfRangeCells(rule, sheetData))
        {
            // Use the display text (resolves inlineStr / shared strings), matching
            // the per-cell read in EvaluateCfRule so the two stay consistent.
            var t = GetCellDisplayValue(c);
            if (!string.IsNullOrEmpty(t)) result.Add(t);
        }
        return result;
    }

    /// <summary>Enumerate the cells that exist within a CF rule's sqref range(s).</summary>
    private IEnumerable<Cell> CollectCfRangeCells(ConditionalFormattingRule rule, SheetData sheetData)
    {
        var cf = rule.Parent as ConditionalFormatting;
        var sqrefs = cf?.SequenceOfReferences?.Items;
        if (sqrefs == null) yield break;
        foreach (var sqItem in sqrefs)
        {
            var sqref = sqItem?.Value;
            if (string.IsNullOrEmpty(sqref)) continue;
            var start = sqref.Contains(':') ? sqref.Split(':')[0] : sqref;
            var end = sqref.Contains(':') ? sqref.Split(':')[1] : sqref;
            var (startColName, startRow) = ParseCellReference(start);
            var (endColName, endRow) = ParseCellReference(end);
            int c1 = ColumnNameToIndex(startColName), c2 = ColumnNameToIndex(endColName);
            int r1 = Math.Min(startRow, endRow), r2 = Math.Max(startRow, endRow);
            int cMin = Math.Min(c1, c2), cMax = Math.Max(c1, c2);
            foreach (var cell in sheetData.Descendants<Cell>())
            {
                var refv = cell.CellReference?.Value;
                if (string.IsNullOrEmpty(refv)) continue;
                var (colName, rowNum) = ParseCellReference(refv);
                var colIdx = ColumnNameToIndex(colName);
                if (rowNum >= r1 && rowNum <= r2 && colIdx >= cMin && colIdx <= cMax)
                    yield return cell;
            }
        }
    }

    /// <summary>Adjust a CF formula's cell references from the anchor cell to the target cell.</summary>
    private string AdjustCfFormula(string formula, int targetRow, int targetCol, ConditionalFormattingRule rule)
    {
        // Find the anchor cell from the parent ConditionalFormatting sqref
        var cf = rule.Parent as ConditionalFormatting;
        var sqref = cf?.SequenceOfReferences?.Items?.FirstOrDefault()?.Value;
        if (string.IsNullOrEmpty(sqref)) return formula;

        // Extract anchor from sqref (e.g. "E7:E21" → anchor is E7)
        var anchorRef = sqref.Contains(':') ? sqref.Split(':')[0] : sqref;
        var (anchorColName, anchorRow) = ParseCellReference(anchorRef);
        var anchorCol = ColumnNameToIndex(anchorColName);

        // Argless ROW()/COLUMN() in a CF formula refer to the current cell.
        // The evaluator dereferences a bare single-cell ref to its value before
        // ROW()/COLUMN() can read the ref's origin, so substitute the resolved
        // row/column index directly (this also sidesteps the ref-delta shift
        // below, since a bare integer carries no A1-style reference).
        formula = Regex.Replace(formula, @"\bROW\s*\(\s*\)", targetRow.ToString(), RegexOptions.IgnoreCase);
        formula = Regex.Replace(formula, @"\bCOLUMN\s*\(\s*\)", targetCol.ToString(), RegexOptions.IgnoreCase);

        var rowDelta = targetRow - anchorRow;
        var colDelta = targetCol - anchorCol;
        if (rowDelta == 0 && colDelta == 0) return formula;

        // Replace cell references in formula, adjusting by delta
        return Regex.Replace(formula, @"(\$?)([A-Z]+)(\$?)(\d+)", m =>
        {
            var colAbsolute = m.Groups[1].Value == "$";
            var rowAbsolute = m.Groups[3].Value == "$";
            var refCol = ColumnNameToIndex(m.Groups[2].Value);
            var refRow = int.Parse(m.Groups[4].Value);

            var newCol = colAbsolute ? refCol : refCol + colDelta;
            var newRow = rowAbsolute ? refRow : refRow + rowDelta;
            if (newCol < 1) newCol = 1;
            if (newRow < 1) newRow = 1;
            return $"{(colAbsolute ? "$" : "")}{IndexToColumnName(newCol)}{(rowAbsolute ? "$" : "")}{newRow}";
        });
    }

    /// <summary>Expand a sqref string like "E7:E21" into individual cell references.</summary>
    /// <summary>
    /// Collect every header cellRef (top row of an AutoFilter range) that should
    /// carry a filter-dropdown indicator. Covers the sheet-level &lt;autoFilter&gt;
    /// and each table's own &lt;autoFilter&gt;.
    /// </summary>
    private HashSet<string> BuildAutoFilterHeaderCells(Worksheet ws, WorksheetPart worksheetPart)
    {
        var cells = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        void AddRange(string? rangeRef)
        {
            if (string.IsNullOrEmpty(rangeRef)) return;
            var range = rangeRef.Replace("$", "");
            var sides = range.Split(':');
            var (startColName, startRow) = ParseCellReference(sides[0]);
            var startCol = ColumnNameToIndex(startColName);
            int endCol = startCol;
            if (sides.Length > 1)
            {
                var (endColName, _) = ParseCellReference(sides[1]);
                endCol = ColumnNameToIndex(endColName);
            }
            for (int c = startCol; c <= endCol; c++)
                cells.Add($"{IndexToColumnName(c)}{startRow}");
        }

        AddRange(ws.GetFirstChild<AutoFilter>()?.Reference?.Value);
        foreach (var tdp in worksheetPart.TableDefinitionParts)
            AddRange(tdp.Table?.AutoFilter?.Reference?.Value);

        return cells;
    }

    /// <summary>
    /// Build a cellRef → background-CSS map for every table (ListObject) carrying a
    /// built-in <c>&lt;tableStyleInfo&gt;</c>. Built-in table styles aren't stored in
    /// styles.xml — their colors derive from the workbook theme. We resolve a base
    /// accent from the style-name family and paint:
    ///   header row  = accent solid
    ///   band rows   = accent @ light tint on odd data rows (when showRowStripes)
    /// Explicit cell fills win at the apply site (GetCellStyleCss), so a styled cell
    /// is never overwritten by the band.
    /// </summary>
    private Dictionary<string, string> BuildTableStyleMap(WorksheetPart worksheetPart)
    {
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        foreach (var tdp in worksheetPart.TableDefinitionParts)
        {
            var table = tdp.Table;
            var tsi = table?.TableStyleInfo;
            var styleName = tsi?.Name?.Value;
            if (table?.Reference?.Value == null || string.IsNullOrEmpty(styleName)) continue;

            var (headerHex, bandHex) = ResolveTableStyleColors(styleName);
            if (headerHex == null) continue;

            // Parse the table region A1:D7
            var range = table.Reference.Value!.Replace("$", "");
            var sides = range.Split(':');
            var (startColName, startRow) = ParseCellReference(sides[0]);
            int startCol = ColumnNameToIndex(startColName);
            int endRow = startRow, endCol = startCol;
            if (sides.Length > 1)
            {
                var (endColName, er) = ParseCellReference(sides[1]);
                endCol = ColumnNameToIndex(endColName);
                endRow = er;
            }

            int headerRows = (int)(table.HeaderRowCount?.Value ?? 1);
            int totalsRows = (int)(table.TotalsRowCount?.Value ?? 0);
            bool rowStripes = tsi?.ShowRowStripes?.Value == true;

            int firstDataRow = startRow + headerRows;
            int lastDataRow = endRow - totalsRows;

            for (int r = startRow; r <= endRow; r++)
            {
                string? css = null;
                if (r < firstDataRow)              // header band
                    css = $"background:{headerHex};color:#FFFFFF;font-weight:bold";
                else if (r > lastDataRow)          // totals band
                    css = $"background:{bandHex}";
                else if (rowStripes)
                {
                    // Stripe odd data rows (1st data row = unstriped/white) to match
                    // Excel's first-row-light convention for Medium styles.
                    int dataIdx = r - firstDataRow;
                    if (dataIdx % 2 == 1) css = $"background:{bandHex}";
                }
                if (css == null) continue;
                for (int c = startCol; c <= endCol; c++)
                    map[$"{IndexToColumnName(c)}{r}"] = css;
            }
        }

        return map;
    }

    /// <summary>
    /// Map a built-in table-style name (e.g. TableStyleMedium9) to a header fill and
    /// a band fill, sourced from the workbook theme accent palette. Light/Medium/Dark
    /// families share the accent-cycling mapping Excel uses (MediumN → accent((N-2)%6)),
    /// with band = accent at a light tint. Dark family inverts header text handled at
    /// the call site. Returns (null, null) for unrecognized names.
    /// </summary>
    private (string? header, string? band) ResolveTableStyleColors(string styleName)
    {
        var m = Regex.Match(styleName, @"TableStyle(Light|Medium|Dark)(\d+)", RegexOptions.IgnoreCase);
        if (!m.Success) return (null, null);

        int n = int.TryParse(m.Groups[2].Value, out var parsed) ? parsed : 1;
        var theme = GetExcelThemeColors();

        // Excel groups the colored families in blocks of 7: one neutral (gray) +
        // six accents. Style index 1 is the neutral; 2..7 map to accent1..accent6;
        // 8 neutral again; 9..14 accent1..accent6; etc. (n-1)%7 gives the slot:
        // slot 0 = neutral, slots 1..6 = accent1..accent6. So Medium9 → accent1 (blue),
        // matching native Excel.
        int slot = ((Math.Max(n, 1) - 1) % 7 + 7) % 7; // 0..6
        string accentHex;
        if (slot == 0)
            accentHex = "808080"; // neutral gray family
        else if (!theme.TryGetValue($"accent{slot}", out accentHex!))
            accentHex = "4472C4";

        // header = solid accent; band = accent tinted ~80% toward white (light blue).
        var header = $"#{accentHex}";
        var band = Core.ColorMath.ApplyTransforms(accentHex, tint: 20000); // ≈ 20% accent / 80% white
        return (header, band);
    }

    // Used extent of populated cells (capped by the html render caps). The CF
    // passes only need cells that actually render; expanding a full-column or
    // full-sheet sqref past the data would cost ~1M+ cells and hang the render.
    private (int maxRow, int maxCol) UsedBounds(SheetData sheetData)
    {
        int maxRow = 0, maxCol = 0;
        foreach (var row in sheetData.Elements<Row>())
            foreach (var cell in row.Elements<Cell>())
            {
                var cr = cell.CellReference?.Value;
                if (string.IsNullOrEmpty(cr)) continue;
                var (colName, r) = ParseCellReference(cr);
                if (r > maxRow) maxRow = r;
                var ci = ColumnNameToIndex(colName);
                if (ci > maxCol) maxCol = ci;
            }
        return (Math.Min(maxRow, GetHtmlRowCap()), Math.Min(maxCol, 200));
    }

    // Raw row/col extents of conditional-formatting sqrefs. Blank in-range
    // cells (e.g. a containsBlanks fill on D1:D5 with only D1 populated) live
    // beyond the used data extent; both the rendered grid and the CF passes
    // must cover them or the fill can never display. Callers merge these
    // extents into their bounds CLAMPED by the same html render caps as
    // UsedBounds (GetHtmlRowCap()/200), so a whole-column sqref
    // (A1:A1048576) cannot explode the grid. Refs that don't parse as plain
    // A1-style cells (defensive: malformed files) are skipped.
    private static (int maxRow, int maxCol) CfRangeExtents(IEnumerable<ConditionalFormatting> cfElements)
    {
        int maxRow = 0, maxCol = 0;
        foreach (var cf in cfElements)
        {
            var items = cf.SequenceOfReferences?.Items;
            if (items == null) continue;
            foreach (var item in items)
                foreach (var part in (item.Value ?? "").Split(' ', StringSplitOptions.RemoveEmptyEntries))
                {
                    var end = part.Contains(':') ? part.Split(':')[1] : part;
                    try
                    {
                        var (colName, row) = ParseCellReference(end);
                        if (row > maxRow) maxRow = row;
                        var ci = ColumnNameToIndex(colName);
                        if (ci > maxCol) maxCol = ci;
                    }
                    catch (ArgumentException) { /* skip non-A1 refs (e.g. whole-column "A:A") */ }
                }
        }
        return (maxRow, maxCol);
    }

    private List<(string cellRef, int row, int col)> ExpandSqref(string sqref, int maxRow, int maxCol)
    {
        var result = new List<(string, int, int)>();
        foreach (var part in sqref.Split(' '))
        {
            if (part.Contains(':'))
            {
                var sides = part.Split(':');
                var (startColName, startRow) = ParseCellReference(sides[0]);
                var (endColName, endRow) = ParseCellReference(sides[1]);
                // Clamp to the sheet's used extent (both dims). CF cells beyond
                // the rendered grid never display, and a full-column/full-sheet
                // sqref would otherwise expand to ~1M+ cells and hang the render.
                endRow = Math.Min(endRow, maxRow);
                var startCol = ColumnNameToIndex(startColName);
                var endCol = Math.Min(ColumnNameToIndex(endColName), maxCol);
                for (int r = startRow; r <= endRow; r++)
                    for (int c = startCol; c <= endCol; c++)
                        result.Add(($"{IndexToColumnName(c)}{r}", r, c));
            }
            else
            {
                var (colName, row) = ParseCellReference(part);
                result.Add((part, row, ColumnNameToIndex(colName)));
            }
        }
        return result;
    }

    // ==================== Cell Style to CSS ====================

    // CONSISTENCY(excel-merge-border): for a merged anchor td, the right/bottom CSS
    // edges of the td are the merged region PERIMETER, which Excel sources from the
    // perimeter member cells — not the anchor (whose right/bottom are interior to the
    // merge). Callers pass rightBorderCell (right-column member) and bottomBorderCell
    // (bottom-row member); default null = same as the anchor (non-merge path unchanged).
    private string GetCellStyleCss(Cell? cell, Stylesheet? stylesheet, int frozenRows, int frozenCols, int row, int col,
        Dictionary<int, double>? frozenLeftOffsets = null, Dictionary<int, double>? frozenTopOffsets = null,
        Dictionary<string, string>? cfMap = null, Dictionary<string, string>? dataBarMap = null,
        Dictionary<string, string>? iconSetMap = null,
        Dictionary<string, string>? tableStyleMap = null,
        bool mergePerimeter = false, Cell? rightBorderCell = null, Cell? bottomBorderCell = null)
    {
        var styles = new List<string>();

        // Frozen pane sticky positioning
        bool isFrozenRow = frozenRows > 0 && row <= frozenRows;
        bool isFrozenCol = frozenCols > 0 && col <= frozenCols;
        // z-index layering: corner-cell=4, col-header=3, frozen-row+col=2, frozen-col=1
        var frozenLeft = frozenLeftOffsets?.TryGetValue(col, out var fl) == true ? fl : 0;
        var frozenTop = frozenTopOffsets?.TryGetValue(row, out var ft) == true ? ft : 0;
        if (isFrozenRow && isFrozenCol)
            styles.Add($"position:sticky;top:0;left:{frozenLeft:0.##}pt;z-index:2");
        else if (isFrozenRow)
            styles.Add("position:sticky;top:0;z-index:1");
        else if (isFrozenCol)
            styles.Add($"position:sticky;left:{frozenLeft:0.##}pt;z-index:1");

        if (cell == null || stylesheet == null)
        {
            // CF color-scale backgrounds carry inline colors and don't need the
            // stylesheet — apply them here too so a workbook with no <dxfs>/styles
            // (e.g. a freshly-created file) still renders the gradient.
            var cfRefEarly = $"{IndexToColumnName(col)}{row}";
            if (cfMap != null && cfMap.TryGetValue(cfRefEarly, out var cfCssEarly))
            {
                foreach (var cfPart in cfCssEarly.Split(';'))
                    styles.RemoveAll(s => s.StartsWith(cfPart.Split(':')[0].Trim() + ":"));
                styles.Add(cfCssEarly);
            }
            // Data bar / icon set need position:relative on the TD so their inner
            // absolutely-positioned div anchors to the cell, not the sheet wrapper.
            // This must run here too — a freshly-created xlsx has no stylesheet and
            // every cell hits this early-return branch.
            if ((dataBarMap != null && dataBarMap.ContainsKey(cfRefEarly)) ||
                (iconSetMap != null && iconSetMap.ContainsKey(cfRefEarly)))
            {
                styles.Add("position:relative");
            }
            // Table built-in style banding (no explicit cell fill to override here).
            if (tableStyleMap != null && tableStyleMap.TryGetValue(cfRefEarly, out var tblCssEarly)
                && !styles.Any(s => s.StartsWith("background")))
                styles.Add(tblCssEarly);
            // Frozen rows need opaque background so scrolling content doesn't show through
            // Use actual cell fill if available; fallback to white for cells with no explicit fill
            if (isFrozenRow && !styles.Any(s => s.StartsWith("background")))
                styles.Add("background:#fff");
            // Default General alignment still applies when there is no stylesheet
            // (freshly-created workbooks have no styles.xml): numbers and error
            // values are right-aligned, text is left-aligned.
            if (cell != null) AddDefaultGeneralAlign(cell, styles);
            return styles.Count > 0 ? $" style=\"{string.Join(";", styles)}\"" : "";
        }

        var styleIndex = cell.StyleIndex?.Value ?? 0;

        {
            var cellFormats = stylesheet.CellFormats;
            if (cellFormats != null && styleIndex < (uint)cellFormats.Elements<CellFormat>().Count())
            {
                var xf = cellFormats.Elements<CellFormat>().ElementAt((int)styleIndex);
                BuildFontCss(xf, stylesheet, styles);
                BuildFillCss(xf, stylesheet, styles);
                BuildBorderCss(xf, stylesheet, styles, mergePerimeter, rightBorderCell, bottomBorderCell);
                BuildAlignmentCss(xf, styles, cell);

                // Number-format [Color] section (e.g. "$#,##0.00;[Red](...)" colors
                // negatives red). Applies to numeric cells only; the section is
                // chosen by the cell's value sign. Overrides the font color.
                var numFmtColor = GetCellNumberFormatColor(cell, xf, stylesheet);
                if (numFmtColor != null)
                {
                    styles.RemoveAll(s => s.StartsWith("color:"));
                    styles.Add($"color:{numFmtColor}");
                }

                // Diagonal border needs the TD to be a positioning context for the
                // inline SVG overlay emitted into the cell content (BuildRowInnerHtml).
                if (TryBuildCellDiagonalSvg(cell, stylesheet) != null
                    && !styles.Any(s => s.StartsWith("position:")))
                    styles.Add("position:relative");
            }
        }

        var cfCellRef = $"{IndexToColumnName(col)}{row}";

        // Table built-in style banding: header + alternating row fills derived
        // from the theme. Explicit cell fills win, so apply only when the cell
        // contributed no background of its own. CF (below) still overrides this.
        if (tableStyleMap != null && tableStyleMap.TryGetValue(cfCellRef, out var tblCss)
            && !styles.Any(s => s.StartsWith("background")))
            styles.Add(tblCss);

        // Conditional formatting overrides (background, color)
        if (cfMap != null && cfMap.TryGetValue(cfCellRef, out var cfCss))
        {
            // CF overrides existing background/color — remove conflicting base styles
            foreach (var cfPart in cfCss.Split(';'))
            {
                var prop = cfPart.Split(':')[0].Trim();
                styles.RemoveAll(s => s.StartsWith(prop + ":"));
            }
            styles.Add(cfCss);
        }

        // Data bar or icon set: add position:relative so inner elements can be absolutely positioned
        if ((dataBarMap != null && dataBarMap.ContainsKey(cfCellRef)) ||
            (iconSetMap != null && iconSetMap.ContainsKey(cfCellRef)))
        {
            styles.Add("position:relative");
        }

        // Frozen rows need opaque background so scrolling content doesn't show through
        if (isFrozenRow && !styles.Any(s => s.StartsWith("background:")))
            styles.Add("background:#fff");

        return styles.Count > 0 ? $" style=\"{string.Join(";", styles)}\"" : "";
    }

    private void BuildFontCss(CellFormat xf, Stylesheet stylesheet, List<string> styles)
    {
        var fontId = xf.FontId?.Value ?? 0;
        var fonts = stylesheet.Fonts;
        if (fonts == null || fontId >= (uint)fonts.Elements<Font>().Count()) return;

        var font = fonts.Elements<Font>().ElementAt((int)fontId);

        if (font.Bold != null && font.Bold.Val?.Value != false) styles.Add("font-weight:bold");
        if (font.Italic != null && font.Italic.Val?.Value != false) styles.Add("font-style:italic");
        if (font.Strike != null && font.Strike.Val?.Value != false) styles.Add("text-decoration:line-through");
        if (font.Underline != null)
        {
            var existing = styles.FindIndex(s => s.StartsWith("text-decoration:"));
            if (existing >= 0)
                styles[existing] = styles[existing] + " underline";
            else
                styles.Add("text-decoration:underline");
            // Render double / doubleAccounting as a true double underline.
            var ulVal = font.Underline.Val?.Value;
            if (ulVal == UnderlineValues.Double || ulVal == UnderlineValues.DoubleAccounting)
                styles.Add("text-decoration-style:double");
        }

        // Superscript/Subscript: handled by wrapping cell content in <sup>/<sub>
        // (see GetCellVerticalAlign + the <td> content path). vertical-align on a
        // <td> only controls cell-content block alignment (top/middle/bottom) — it
        // does NOT raise/lower the baseline of inline text, and the font-size:smaller
        // it paired with was overwritten by the full font-size:Npt added below.
        // R10a: emit nothing here; the <sup>/<sub> wrapper provides both the
        // raised baseline and the size reduction.

        if (font.FontSize?.Val?.Value != null)
            styles.Add($"font-size:{font.FontSize.Val.Value:0.##}pt");

        if (font.FontName?.Val?.Value != null)
            styles.Add($"font-family:'{CssSanitize(font.FontName.Val.Value)}'");

        var color = ResolveFontColor(font);
        if (color != null) styles.Add($"color:{color}");
    }

    /// <summary>
    /// R10a: returns "super" / "sub" / null for a cell's font vertical alignment.
    /// Used to wrap cell content in a <sup>/<sub> inline element — vertical-align
    /// on the <td> itself has no baseline-shifting effect on cell content.
    /// </summary>
    private static string? GetCellVerticalAlign(Cell? cell, Stylesheet? stylesheet)
    {
        if (cell == null || stylesheet?.CellFormats == null || stylesheet.Fonts == null) return null;
        var styleIndex = cell.StyleIndex?.Value ?? 0;
        if (styleIndex >= (uint)stylesheet.CellFormats.Elements<CellFormat>().Count()) return null;
        var xf = stylesheet.CellFormats.Elements<CellFormat>().ElementAt((int)styleIndex);
        var fontId = xf.FontId?.Value ?? 0;
        if (fontId >= (uint)stylesheet.Fonts.Elements<Font>().Count()) return null;
        var font = stylesheet.Fonts.Elements<Font>().ElementAt((int)fontId);
        var v = font.GetFirstChild<VerticalTextAlignment>()?.Val?.Value;
        if (v == VerticalAlignmentRunValues.Superscript) return "super";
        if (v == VerticalAlignmentRunValues.Subscript) return "sub";
        return null;
    }

    /// <summary>
    /// R10a: wrap cell content in a <sup>/<sub> element when the cell font is
    /// super/subscript. Skipped for rich text (its runs carry their own
    /// formatting) and when there's no content. <sup>/<sub> give both the
    /// raised/lowered baseline and the ~0.83em size reduction natively.
    /// </summary>
    private static string WrapVerticalAlign(string content, string? vAlign, string? richHtml)
    {
        if (vAlign == null || richHtml != null || string.IsNullOrEmpty(content)) return content;
        var tag = vAlign == "super" ? "sup" : "sub";
        return $"<{tag}>{content}</{tag}>";
    }

    private void BuildFillCss(CellFormat xf, Stylesheet stylesheet, List<string> styles)
    {
        var fillId = xf.FillId?.Value ?? 0;
        if (fillId <= 1) return; // 0=none, 1=gray125 pattern (default)

        var fills = stylesheet.Fills;
        if (fills == null || fillId >= (uint)fills.Elements<Fill>().Count()) return;

        var fill = fills.Elements<Fill>().ElementAt((int)fillId);

        // Gradient fill
        var gf = fill.GetFirstChild<GradientFill>();
        if (gf != null)
        {
            var stops = gf.Elements<GradientStop>().ToList();
            if (stops.Count >= 2)
            {
                var colors = stops
                    .Select(s => ResolveColorRgb(s.Color))
                    .Where(c => c != null)
                    .ToList();
                if (colors.Count >= 2)
                {
                    var deg = (int)(gf.Degree?.Value ?? 0);
                    styles.Add($"background:linear-gradient({(deg + 90) % 360}deg,{string.Join(",", colors)})");
                    return;
                }
            }
        }

        // Pattern fill
        var pf = fill.PatternFill;
        if (pf != null)
        {
            var bgColor = ResolveColorRgb(pf.ForegroundColor);
            if (bgColor != null) styles.Add($"background:{bgColor}");
        }
    }

    private void BuildBorderCss(CellFormat xf, Stylesheet stylesheet, List<string> styles,
        bool mergePerimeter = false, Cell? rightBorderCell = null, Cell? bottomBorderCell = null)
    {
        var borderId = xf.BorderId?.Value ?? 0;
        var borders = stylesheet.Borders;
        Border? border = (borderId != 0 && borders != null && borderId < (uint)borders.Elements<Border>().Count())
            ? borders.Elements<Border>().ElementAt((int)borderId)
            : null;

        // top/left always come from the anchor cell (top-left member of the merge).
        AddBorderSideCss(border?.TopBorder, "top", styles);
        AddBorderSideCss(border?.LeftBorder, "left", styles);

        // right/bottom: for a merged anchor the td edge is the region PERIMETER, which
        // Excel sources from the perimeter member cell — not the (interior) anchor edge.
        // If the member cell is absent (or carries no border) the region has NO border on
        // that edge. Non-merge path (mergePerimeter=false) keeps the anchor's own edges.
        var rightBorder = mergePerimeter ? (rightBorderCell != null ? GetCellBorder(rightBorderCell, stylesheet) : null) : border;
        var bottomBorder = mergePerimeter ? (bottomBorderCell != null ? GetCellBorder(bottomBorderCell, stylesheet) : null) : border;
        AddBorderSideCss(rightBorder?.RightBorder, "right", styles);
        AddBorderSideCss(bottomBorder?.BottomBorder, "bottom", styles);
    }

    // Resolve a cell's <border> element via its style index, or null if none.
    private Border? GetCellBorder(Cell cell, Stylesheet stylesheet)
    {
        var styleIndex = cell.StyleIndex?.Value ?? 0;
        var cellFormats = stylesheet.CellFormats;
        if (cellFormats == null || styleIndex >= (uint)cellFormats.Elements<CellFormat>().Count()) return null;
        var xf = cellFormats.Elements<CellFormat>().ElementAt((int)styleIndex);
        var borderId = xf.BorderId?.Value ?? 0;
        var borders = stylesheet.Borders;
        if (borderId == 0 || borders == null || borderId >= (uint)borders.Elements<Border>().Count()) return null;
        return borders.Elements<Border>().ElementAt((int)borderId);
    }

    /// <summary>
    /// Look up the cell's &lt;border&gt; and, if it carries a diagonal
    /// (diagonalDown / diagonalUp with a styled &lt;diagonal&gt; child), return an
    /// absolutely-positioned inline SVG that draws the diagonal line(s) inside the
    /// TD. Mirrors the PPTX table diagonal-overlay idiom (cell-diag SVG). Returns
    /// null when the cell has no diagonal border. The TD must be position:relative
    /// for the overlay to anchor to the cell (added in GetCellStyleCss).
    /// </summary>
    private string? TryBuildCellDiagonalSvg(Cell? cell, Stylesheet? stylesheet)
    {
        if (cell == null || stylesheet == null) return null;
        var styleIndex = cell.StyleIndex?.Value ?? 0;
        var cellFormats = stylesheet.CellFormats;
        if (cellFormats == null || styleIndex >= (uint)cellFormats.Elements<CellFormat>().Count())
            return null;
        var xf = cellFormats.Elements<CellFormat>().ElementAt((int)styleIndex);
        var borderId = xf.BorderId?.Value ?? 0;
        var borders = stylesheet.Borders;
        if (borders == null || borderId == 0 || borderId >= (uint)borders.Elements<Border>().Count())
            return null;
        var border = borders.Elements<Border>().ElementAt((int)borderId);

        bool down = border.DiagonalDown?.Value == true;
        bool up = border.DiagonalUp?.Value == true;
        if (!down && !up) return null;

        var diag = border.DiagonalBorder;
        if (diag?.Style?.Value == null || diag.Style.Value == BorderStyleValues.None) return null;

        var bsv = diag.Style.Value;
        double widthPx = bsv == BorderStyleValues.Thick ? 3 : bsv == BorderStyleValues.Medium ? 2 : 1;
        var color = ResolveColorRgb(diag.Color) ?? "#000";

        var lines = new StringBuilder();
        // diagonalDown = top-left → bottom-right; diagonalUp = bottom-left → top-right.
        if (down)
            lines.Append($"<line x1=\"0\" y1=\"0\" x2=\"100%\" y2=\"100%\" stroke=\"{color}\" stroke-width=\"{widthPx:0.##}\"/>");
        if (up)
            lines.Append($"<line x1=\"0\" y1=\"100%\" x2=\"100%\" y2=\"0\" stroke=\"{color}\" stroke-width=\"{widthPx:0.##}\"/>");

        return $"<svg class=\"cell-diag\" width=\"100%\" height=\"100%\" style=\"position:absolute;inset:0;pointer-events:none;overflow:visible\" preserveAspectRatio=\"none\">{lines}</svg>";
    }

    private void AddBorderSideCss(BorderPropertiesType? bp, string side, List<string> styles)
    {
        if (bp?.Style?.Value == null || bp.Style.Value == BorderStyleValues.None) return;

        var bsv = bp.Style.Value;
        // Medium-weight families render 2px; thick 3px; everything else 1px.
        var width = "1px";
        if (bsv == BorderStyleValues.Medium || bsv == BorderStyleValues.MediumDashed ||
            bsv == BorderStyleValues.MediumDashDot || bsv == BorderStyleValues.MediumDashDotDot)
            width = "2px";
        else if (bsv == BorderStyleValues.Thick) width = "3px";
        else if (bsv == BorderStyleValues.Double) width = "3px";

        // CSS border-style can't draw exact dash-dot patterns, so map the whole
        // dash/dash-dot family to the closest representable "dashed".
        var cssStyle = "solid";
        if (bsv == BorderStyleValues.Dashed || bsv == BorderStyleValues.MediumDashed ||
            bsv == BorderStyleValues.DashDot || bsv == BorderStyleValues.DashDotDot ||
            bsv == BorderStyleValues.SlantDashDot || bsv == BorderStyleValues.MediumDashDot ||
            bsv == BorderStyleValues.MediumDashDotDot) cssStyle = "dashed";
        else if (bsv == BorderStyleValues.Dotted) cssStyle = "dotted";
        else if (bsv == BorderStyleValues.Double) cssStyle = "double";

        var color = ResolveColorRgb(bp.Color);
        color ??= "#000";

        styles.Add($"border-{side}:{width} {cssStyle} {color}");
    }

    /// <summary>
    /// Apply Excel's General (default) horizontal alignment for a cell with no
    /// explicit alignment: numbers and error values right-aligned, text left
    /// (the CSS default, so nothing emitted). Error cells (t="e", e.g. #DIV/0!,
    /// #NAME?) align like numbers. No-ops if a text-align is already present.
    /// </summary>
    private static void AddDefaultGeneralAlign(Cell cell, List<string> styles)
    {
        if (styles.Any(s => s.StartsWith("text-align:"))) return;
        var dt = cell.DataType?.Value;
        bool isText = dt == CellValues.SharedString || dt == CellValues.InlineString || dt == CellValues.String;
        if (isText) return;
        // Boolean values (TRUE/FALSE, or a formula returning a boolean) are
        // center-aligned by default in Excel — not right-aligned like numbers.
        if (dt == CellValues.Boolean)
        {
            styles.Add("text-align:center");
            return;
        }
        // Error cells right-align even if the value lives only in the formula's
        // cached <v> (or is an error pattern), matching real Excel.
        bool isError = dt == CellValues.Error;
        if ((!isText && cell.CellValue != null) || isError)
            styles.Add("text-align:right");
    }

    private void BuildAlignmentCss(CellFormat xf, List<string> styles, Cell? cell = null)
    {
        var alignment = xf.Alignment;
        bool hasExplicitHAlign = alignment?.Horizontal?.HasValue == true;

        if (hasExplicitHAlign)
        {
            var h = alignment!.Horizontal!.InnerText;
            var cssAlign = h switch
            {
                "center" => "center",
                "right" => "right",
                "left" => "left",
                "justify" => "justify",
                "fill" => "left",
                "general" => (string?)null, // fall through to auto-detect
                _ => null
            };
            if (cssAlign != null) { styles.Add($"text-align:{cssAlign}"); hasExplicitHAlign = true; }
            else hasExplicitHAlign = false;
        }

        // Excel default: numbers and errors right-aligned, text left-aligned.
        if (!hasExplicitHAlign && cell != null)
            AddDefaultGeneralAlign(cell, styles);

        if (alignment == null) return;

        if (alignment.Vertical?.HasValue == true)
        {
            var v = alignment.Vertical.InnerText;
            var cssVAlign = v switch
            {
                "top" => "top",
                "center" => "middle",
                "bottom" => "bottom",
                _ => null
            };
            if (cssVAlign != null) styles.Add($"vertical-align:{cssVAlign}");
        }

        if (alignment.WrapText?.Value == true)
            styles.Add("white-space:pre-wrap;word-wrap:break-word");

        if (alignment.TextRotation?.HasValue == true && alignment.TextRotation.Value != 0)
        {
            var rot = alignment.TextRotation.Value;
            if (rot == 255)
            {
                // 255 = stacked vertical text (each char on its own line)
                styles.Add("writing-mode:vertical-rl;text-orientation:upright;letter-spacing:-2px");
            }
            else
            {
                // Excel: 0-90 = counter-clockwise, 91-180 = clockwise (91=1°CW, 180=90°CW)
                // Excel: 1-90 = CCW (CSS negative), 91-180 = CW (CSS positive, 91=1°, 180=90°)
                int cssDeg = rot <= 90 ? -(int)rot : (int)rot - 90;
                // The td's default rule clips its content (overflow:hidden +
                // text-overflow:ellipsis + max-width:500px) to the un-rotated column
                // width, so after the rotate the string truncates ("Rotat…") even
                // though the row was grown tall enough. Override those for rotated
                // cells: keep the box at column width but let the rotated text run to
                // its full length (vertically, within the expanded row height).
                styles.Add($"transform:rotate({cssDeg}deg)");
                styles.Add("white-space:nowrap");
                styles.Add("overflow:visible");
                styles.Add("text-overflow:clip");
                styles.Add("max-width:none");
            }
        }

        if (alignment.Indent?.HasValue == true && alignment.Indent.Value > 0)
        {
            // 1 indent level ≈ width of "0" in default font ≈ fontSize × 0.6
            var defFontSz = _doc.WorkbookPart?.WorkbookStylesPart?.Stylesheet
                ?.Fonts?.Elements<Font>().FirstOrDefault()?.FontSize?.Val?.Value ?? 11.0;
            var indentPt = alignment.Indent.Value * defFontSz * 0.6;
            styles.Add($"padding-left:{indentPt:0.#}pt");
        }

        // Reading order: 1=LTR, 2=RTL (for mixed-direction content)
        if (alignment.ReadingOrder?.HasValue == true)
        {
            var ro = alignment.ReadingOrder.Value;
            if (ro == 2) styles.Add("direction:rtl;unicode-bidi:embed");
            else if (ro == 1) styles.Add("direction:ltr;unicode-bidi:embed");
        }
    }

    // ==================== Color Resolution ====================

    private string? ResolveFontColor(Font font)
    {
        if (font.Color?.Rgb?.Value != null)
        {
            var raw = font.Color.Rgb.Value;
            return FormatColorForCss(raw);
        }
        if (font.Color?.Theme?.Value != null)
        {
            var tint = font.Color.Tint?.Value;
            return ResolveThemeColor(font.Color.Theme.Value, tint);
        }
        return null;
    }

    // Standard Excel indexed color palette (first 64 colors) — can be overridden by styles.xml
    private static readonly string[] DefaultIndexedColors = [
        "#000000","#FFFFFF","#FF0000","#00FF00","#0000FF","#FFFF00","#FF00FF","#00FFFF",
        "#000000","#FFFFFF","#FF0000","#00FF00","#0000FF","#FFFF00","#FF00FF","#00FFFF",
        "#800000","#008000","#000080","#808000","#800080","#008080","#C0C0C0","#808080",
        "#9999FF","#993366","#FFFFCC","#CCFFFF","#660066","#FF8080","#0066CC","#CCCCFF",
        "#000080","#FF00FF","#FFFF00","#00FFFF","#800080","#800000","#008080","#0000FF",
        "#00CCFF","#CCFFFF","#CCFFCC","#FFFF99","#99CCFF","#FF99CC","#CC99FF","#FFCC99",
        "#3366FF","#33CCCC","#99CC00","#FFCC00","#FF9900","#FF6600","#666699","#969696",
        "#003366","#339966","#003300","#333300","#993300","#993366","#333399","#333333"
    ];

    private string? ResolveColorRgb(ColorType? color)
    {
        if (color?.Rgb?.Value != null)
            return FormatColorForCss(color.Rgb.Value);
        if (color?.Indexed?.Value != null)
        {
            var idx = (int)color.Indexed.Value;
            var palette = GetResolvedIndexedColors();
            if (idx >= 0 && idx < palette.Length)
                return palette[idx];
            if (idx == 64) return null; // system foreground (context dependent)
            if (idx == 65) return null; // system background
        }
        if (color?.Theme?.Value != null)
        {
            var tint = color.Tint?.Value;
            return ResolveThemeColor(color.Theme.Value, tint);
        }
        return null;
    }

    private static string FormatColorForCss(string raw)
    {
        // Reject non-hex raw values before interpolating into inline CSS —
        // styles.xml / indexedColors attrs are attacker-controlled, and an
        // unvalidated raw flows into `color:#{raw}` / `background:#{raw}`
        // as an XSS sink.
        static bool isHex(string s) =>
            s.All(c => (c >= '0' && c <= '9') || (c >= 'A' && c <= 'F') || (c >= 'a' && c <= 'f'));
        if (raw.Length == 8 && isHex(raw)) return "#" + raw[2..];
        if (raw.Length is 6 or 3 && isHex(raw)) return "#" + raw;
        return "#000";
    }

    // ==================== Formatted Cell Value ====================

    /// <summary>
    /// Get cell display value with number formatting applied for HTML preview.
    /// Handles common formats: percentage, thousands separator, decimal places, dates.
    /// </summary>
    private string GetFormattedCellValue(Cell cell, Stylesheet? stylesheet, Core.FormulaEvaluator? evaluator = null)
    {
        var rawValue = GetCellDisplayValue(cell);

        // If the cell has a formula, always try to evaluate (cached values may be stale)
        if (cell.CellFormula?.Text != null && evaluator != null)
        {
            var result = evaluator.TryEvaluateFull(cell.CellFormula.Text);
            if (result != null)
            {
                if (result.IsError) return result.ErrorValue!;
                rawValue = result.ToCellValueText();
                if (result.IsString) return rawValue;
                if (result.IsBool) return result.BoolValue!.Value ? "TRUE" : "FALSE";
            }
            // If evaluation fails (null), fall through to use cached value / raw display
        }

        // The internal #OCLI_NOTEVAL! sentinel is an implementation detail and must
        // never reach user-facing HTML. Prefer the cached <v> if it carries a real
        // value; otherwise surface Excel's generic #NAME? error (unknown function /
        // unevaluable formula) so the cell reads like real Excel.
        if (rawValue == "#OCLI_NOTEVAL!")
        {
            var cached = cell.CellValue?.Text;
            rawValue = !string.IsNullOrEmpty(cached) && cached != "#OCLI_NOTEVAL!"
                ? cached
                : "#NAME?";
        }

        if (string.IsNullOrEmpty(rawValue)) return rawValue;

        // Boolean: convert 1/0 to TRUE/FALSE
        if (cell.DataType?.Value == CellValues.Boolean)
            return rawValue == "1" ? "TRUE" : "FALSE";

        // Only format numeric values (not strings, shared strings, etc.)
        if (cell.DataType?.Value == CellValues.SharedString ||
            cell.DataType?.Value == CellValues.InlineString ||
            cell.DataType?.Value == CellValues.String ||
            cell.DataType?.Value == CellValues.Error)
        {
            // Text-format codes (containing the '@' placeholder) wrap the cell text
            // with quoted literals, e.g. "Hello, "@ → "Hello, World". Excel applies
            // the @-section of the format only to text values.
            if (cell.DataType?.Value != CellValues.Error)
            {
                var textFmt = ResolveCellFormatCode(cell, stylesheet);
                if (textFmt != null && ContainsCharOutsideQuotes(textFmt, '@'))
                    return ApplyTextFormat(rawValue, textFmt);
            }
            return rawValue;
        }

        if (!double.TryParse(rawValue, System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out var numVal))
        {
            // GetCellDisplayValue pre-formats date serials to ISO (e.g. "2023-03-15"),
            // which fails double.TryParse. For numeric cells with a date format code,
            // recover the raw serial and let ApplyNumberFormat honour the date code
            // (HTML preview only — get/query keeps the canonical ISO form).
            var serialText = cell.CellValue?.Text;
            if (cell.CellFormula?.Text == null && serialText != null &&
                double.TryParse(serialText, System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture, out var serial))
            {
                var dateFmt = ResolveCellFormatCode(cell, stylesheet);
                if (dateFmt != null && ContainsDateTokenOutsideQuotes(dateFmt))
                    return ApplyNumberFormat(serial, dateFmt);
            }
            return rawValue;
        }

        // Normalize negative zero to positive zero — real Excel renders -0 as "0".
        if (numVal == 0) { numVal = 0.0; rawValue = "0"; }

        // Clean up floating point artifacts for display (e.g. 25300000.000000004 → 25300000)
        var cleanVal = numVal;
        var rounded = Math.Round(numVal, 10);
        if (Math.Abs(rounded - Math.Round(rounded)) < 1e-9)
            cleanVal = Math.Round(rounded);
        rawValue = cleanVal == numVal ? rawValue
            : cleanVal.ToString(System.Globalization.CultureInfo.InvariantCulture);

        // Look up number format
        var fmtCode = ResolveCellFormatCode(cell, stylesheet);
        if (fmtCode == null) return FormatGeneralNumber(numVal, rawValue);

        return ApplyNumberFormat(numVal, fmtCode);
    }

    /// <summary>
    /// Format a General-formatted numeric cell. Excel's General format falls back
    /// to scientific notation when a number's magnitude needs more than ~11
    /// significant digits to display (very large integers or very small
    /// fractions). Normal-magnitude numbers pass through their plain text.
    /// </summary>
    private static string FormatGeneralNumber(double value, string rawValue)
    {
        if (value == 0 || double.IsNaN(value) || double.IsInfinity(value))
            return rawValue;

        double abs = Math.Abs(value);
        int exp = (int)Math.Floor(Math.Log10(abs));

        // Excel General switches to scientific when the plain decimal would need
        // more than 11 significant digits / character columns: large magnitudes
        // (exp >= 11) and very small magnitudes (exp <= -5 — values below 1e-4).
        bool useScientific = exp >= 11 || exp <= -5;
        if (!useScientific)
        {
            // Excel's General format rounds the number to ~15 significant digits
            // for display: it never surfaces IEEE-754 float noise such as
            // 99.98999999999999 (shows 99.99) or 8.880000000000001 (shows 8.88).
            // The shortest round-trippable rawValue keeps that noise for computed/
            // imported/round-tripped <v> values, so re-render through G15 (which
            // stays fixed-point for this magnitude range). Already-clean values and
            // integers round-trip unchanged.
            return value.ToString("G15", System.Globalization.CultureInfo.InvariantCulture);
        }

        // Mantissa with up to 5 fractional digits (Excel General caps at ~6
        // significant figures in scientific), trailing zeros trimmed.
        //
        // Derive mantissa+exponent from the value's own shortest round-trippable
        // string (.NET Core default ToString) rather than value / 10^exp, which
        // overflows to Infinity for extreme exponents (e.g. the min positive
        // subnormal 5E-324, where Math.Pow(10, -324) underflows toward 0). The
        // shortest form gives Excel's "5E-324" instead of the exact-bits
        // "4.94066E-324", and for normal magnitudes still rounds to ~6 sig figs.
        var (m, e) = NormalizeScientific(value);
        exp = e;
        // Round the mantissa to the displayed precision (5 fractional digits)
        // BEFORE emitting. Rounding can push the mantissa to >= 10 (e.g.
        // 9.99999999999999e14 -> 10), which is non-canonical: carry the extra
        // power of ten into the exponent so 10E+14 renders as Excel's 1E+15.
        var roundedM = Math.Round(m, 5, MidpointRounding.AwayFromZero);
        if (Math.Abs(roundedM) >= 10.0) { roundedM /= 10.0; exp++; }
        var mantStr = roundedM.ToString("0.#####", System.Globalization.CultureInfo.InvariantCulture);
        var expStr = exp >= 0
            ? $"+{exp.ToString("00", System.Globalization.CultureInfo.InvariantCulture)}"
            : $"-{Math.Abs(exp).ToString("00", System.Globalization.CultureInfo.InvariantCulture)}";
        return $"{mantStr}E{expStr}";
    }

    /// <summary>
    /// Decompose a non-zero finite double into (mantissa, base-10 exponent) where
    /// 1 &lt;= |mantissa| &lt; 10, using the shortest round-trippable decimal string
    /// (.NET Core default ToString) so extreme magnitudes never overflow and the
    /// shortest sane mantissa is used (5E-324, not 4.94066E-324). Never divides by
    /// a power of 10, so it is overflow/underflow safe.
    /// </summary>
    private static (double mantissa, int exp) NormalizeScientific(double value)
    {
        // "R"/default round-trip in scientific form, e.g. "5E-324", "1.2345E+20".
        var s = value.ToString("R", System.Globalization.CultureInfo.InvariantCulture);
        var ePos = s.IndexOfAny(new[] { 'E', 'e' });
        double mant;
        int exp;
        if (ePos >= 0)
        {
            mant = double.Parse(s.Substring(0, ePos), System.Globalization.CultureInfo.InvariantCulture);
            exp = int.Parse(s.Substring(ePos + 1), System.Globalization.CultureInfo.InvariantCulture);
        }
        else
        {
            mant = double.Parse(s, System.Globalization.CultureInfo.InvariantCulture);
            exp = 0;
        }
        // Renormalize so 1 <= |mant| < 10 (the round-trip mantissa may be e.g. 12.3).
        while (Math.Abs(mant) >= 10.0) { mant /= 10.0; exp++; }
        while (mant != 0 && Math.Abs(mant) < 1.0) { mant *= 10.0; exp--; }
        return (mant, exp);
    }

    /// <summary>
    /// Resolve a cell's number format code (custom &lt;numFmt&gt; first, then built-in).
    /// Returns null when the cell has no explicit (non-General) format.
    /// </summary>
    private static string? ResolveCellFormatCode(Cell cell, Stylesheet? stylesheet)
    {
        var styleIndex = cell.StyleIndex?.Value ?? 0;
        if (styleIndex == 0 || stylesheet == null) return null;

        var cellFormats = stylesheet.CellFormats;
        if (cellFormats == null || styleIndex >= (uint)cellFormats.Elements<CellFormat>().Count())
            return null;

        var xf = cellFormats.Elements<CellFormat>().ElementAt((int)styleIndex);
        var numFmtId = xf.NumberFormatId?.Value ?? 0;
        if (numFmtId == 0) return null;

        var customFmt = stylesheet.NumberingFormats?.Elements<NumberingFormat>()
            .FirstOrDefault(nf => nf.NumberFormatId?.Value == numFmtId);
        if (customFmt?.FormatCode?.Value != null)
            return customFmt.FormatCode.Value;
        return ResolveBuiltInFormat(numFmtId);
    }

    /// <summary>
    /// Rich-text cells (shared-string items with multiple &lt;r&gt; runs carrying
    /// per-run &lt;rPr&gt;) flatten to plain text via GetCellDisplayValue's InnerText.
    /// Build per-run &lt;span&gt; HTML instead so color/bold/italic/font survive.
    /// Returns pre-encoded HTML (run text already HtmlEncoded) when the cell is a
    /// shared string with at least one run that has run-properties; otherwise null
    /// so the caller falls back to the flat CellHtml path.
    /// </summary>
    private string? TryBuildRichTextHtml(Cell cell)
    {
        if (cell.DataType?.Value != CellValues.SharedString) return null;
        var value = cell.CellValue?.Text;
        if (value == null || !int.TryParse(value, out int idx)) return null;

        var sst = _doc.WorkbookPart?.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
        var item = sst?.SharedStringTable?.Elements<SharedStringItem>().ElementAtOrDefault(idx);
        if (item == null) return null;

        var runs = item.Elements<Run>().ToList();
        // Only worth wrapping when at least one run carries explicit run-properties;
        // a single plain run is identical to flat text — let the normal path handle it.
        if (runs.Count == 0 || !runs.Any(r => r.RunProperties != null)) return null;

        var sb = new StringBuilder();
        foreach (var run in runs)
        {
            var rPr = run.RunProperties;
            var style = new StringBuilder();
            if (rPr != null)
            {
                // Shared-string run properties expose children only via GetFirstChild;
                // <b/>/<i/>/<u/> are presence-flags (a Val=false would disable, mirror that).
                var bold = rPr.GetFirstChild<Bold>();
                if (bold != null && bold.Val?.Value != false)
                    style.Append("font-weight:bold;");
                var italic = rPr.GetFirstChild<Italic>();
                if (italic != null && italic.Val?.Value != false)
                    style.Append("font-style:italic;");
                var underline = rPr.GetFirstChild<Underline>();
                if (underline != null && underline.Val?.Value != UnderlineValues.None)
                    style.Append("text-decoration:underline;");
                var colorHex = ResolveRunColorHex(rPr.GetFirstChild<Color>());
                if (colorHex != null) style.Append($"color:{colorHex};");
                if (rPr.GetFirstChild<FontSize>()?.Val?.Value is double fs)
                    style.Append($"font-size:{fs:0.##}pt;");
                var fontName = rPr.GetFirstChild<RunFont>()?.Val?.Value;
                if (!string.IsNullOrEmpty(fontName))
                    style.Append($"font-family:'{fontName}';");
            }
            var text = HtmlEncode(run.Text?.Text ?? "");
            var span = style.Length > 0 ? $"<span style=\"{style}\">{text}</span>" : $"<span>{text}</span>";
            // Mirror the cell-level vertical-align path (GetCellVerticalAlign /
            // WrapVerticalAlign): wrap in semantic <sup>/<sub> which gives both
            // the baseline shift and the ~0.83em size reduction, while the inner
            // <span> keeps the run's font/color.
            var vAlign = rPr?.GetFirstChild<VerticalTextAlignment>()?.Val?.Value;
            if (vAlign == VerticalAlignmentRunValues.Superscript) span = $"<sup>{span}</sup>";
            else if (vAlign == VerticalAlignmentRunValues.Subscript) span = $"<sub>{span}</sub>";
            sb.Append(span);
        }
        return sb.ToString();
    }

    /// <summary>Resolve a shared-string run &lt;color&gt; to a #RRGGBB hex (rgb or theme), else null.</summary>
    private string? ResolveRunColorHex(Color? color)
    {
        if (color == null) return null;
        var rgb = color.Rgb?.Value;
        if (!string.IsNullOrEmpty(rgb))
        {
            if (rgb.Length > 6) rgb = rgb[^6..];   // strip ARGB alpha
            return $"#{rgb}";
        }
        if (color.Theme?.Value is uint themeIdx)
        {
            var tint = color.Tint?.Value;
            return ResolveThemeColor(themeIdx, tint);
        }
        return null;
    }

    private static string? ResolveBuiltInFormat(uint numFmtId) => numFmtId switch
    {
        1 => "0",
        2 => "0.00",
        3 => "#,##0",
        4 => "#,##0.00",
        12 => "# ?/?",
        13 => "# ??/??",
        9 => "0%",
        10 => "0.00%",
        11 => "0.00E+00",
        14 => "m/d/yy",
        15 => "d-mmm-yy",
        16 => "d-mmm",
        17 => "mmm-yy",
        18 => "h:mm AM/PM",
        19 => "h:mm:ss AM/PM",
        20 => "h:mm",
        21 => "h:mm:ss",
        22 => "m/d/yy h:mm",
        37 => "#,##0 ;(#,##0)",
        38 => "#,##0 ;(#,##0)",
        39 => "#,##0.00;(#,##0.00)",
        40 => "#,##0.00;(#,##0.00)",
        49 => "@",
        _ => null
    };

    // Excel number-format [Color] names → CSS hex (the named palette only; the
    // [Color N] indexed form maps to the indexed-color table, omitted here as a
    // known limitation — named colors cover the common $;[Red](…) negative case).
    private static readonly Dictionary<string, string> NumFmtColorNames =
        new(StringComparer.OrdinalIgnoreCase)
        {
            // Excel's named format colors are the legacy VGA palette, NOT the
            // CSS color names — e.g. [Green] is lime #00FF00, not CSS #008000.
            ["Black"] = "#000000", ["White"] = "#FFFFFF", ["Red"] = "#FF0000",
            ["Green"] = "#00FF00", ["Blue"] = "#0000FF", ["Yellow"] = "#FFFF00",
            ["Magenta"] = "#FF00FF", ["Cyan"] = "#00FFFF",
        };

    /// <summary>
    /// Resolve the CSS color implied by the number format's [Color] tag for the
    /// section that applies to <paramref name="value"/> (positive;negative;zero).
    /// Returns null when the active section carries no [Color] marker.
    /// </summary>
    private static string? GetNumberFormatColor(double value, string fmtCode)
    {
        string section;
        if (fmtCode.Contains(';'))
        {
            var sections = fmtCode.Split(';');

            // Explicit [condition] sections (e.g. [Blue][>50]0;[Red]0). Excel
            // evaluates the bracketed conditions IN DECLARATION ORDER; the first
            // satisfied section applies, and the last unconditioned section is the
            // "else". This overrides the positional positive/negative/zero rule
            // and must match ApplyNumberFormat's section selection so the color
            // comes from the same section as the formatted value.
            if (System.Text.RegularExpressions.Regex.IsMatch(fmtCode, @"\[[<>=]=?\d"))
            {
                section = SelectConditionalSection(value, sections);
            }
            else if (value < 0 && sections.Length >= 2) section = sections[1];
            else if (value == 0 && sections.Length >= 3) section = sections[2];
            else section = sections[0];
        }
        else
        {
            section = fmtCode;
        }

        return ParseSectionColor(section);
    }

    /// <summary>
    /// Pick the section that applies to <paramref name="value"/> when at least one
    /// section carries a bracketed [comparison] condition. Conditions are evaluated
    /// left-to-right; the first satisfied condition wins, an unconditioned section
    /// is the "else". Falls back to the first section when nothing matches. Mirrors
    /// the selection in ApplyNumberFormat.
    /// </summary>
    private static string SelectConditionalSection(double value, string[] sections)
    {
        foreach (var raw in sections)
        {
            var sec = raw.Trim();
            var condMatch = System.Text.RegularExpressions.Regex.Match(
                sec, @"\[(<=|>=|<>|<|>|=)(-?\d+\.?\d*)\]");
            if (condMatch.Success)
            {
                var op = condMatch.Groups[1].Value;
                var cmp = double.Parse(condMatch.Groups[2].Value,
                    System.Globalization.CultureInfo.InvariantCulture);
                bool satisfied = op switch
                {
                    "<" => value < cmp,
                    "<=" => value <= cmp,
                    ">" => value > cmp,
                    ">=" => value >= cmp,
                    "=" => value == cmp,
                    "<>" => value != cmp,
                    _ => false
                };
                if (satisfied) return sec;
            }
            else
            {
                return sec; // unconditioned else
            }
        }
        return sections[0].Trim();
    }

    /// <summary>
    /// Resolve the CSS color implied by the number format's text (@) section
    /// for a string cell. Mirrors ApplyTextFormat's section selection: the 4th
    /// section in a multi-section code, else whichever section carries '@'.
    /// </summary>
    private static string? GetTextSectionColor(string fmtCode)
    {
        string section;
        if (fmtCode.Contains(';'))
        {
            var sections = fmtCode.Split(';');
            section = sections.Length >= 4 ? sections[3]
                : sections.FirstOrDefault(s => ContainsCharOutsideQuotes(s, '@')) ?? fmtCode;
        }
        else
        {
            section = fmtCode;
        }
        return ParseSectionColor(section);
    }

    /// <summary>Extract the [Color] named token from a single format section.</summary>
    private static string? ParseSectionColor(string section)
    {
        var m = System.Text.RegularExpressions.Regex.Match(
            section, @"\[(Black|White|Red|Green|Blue|Yellow|Magenta|Cyan)\]",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        return m.Success && NumFmtColorNames.TryGetValue(m.Groups[1].Value, out var hex) ? hex : null;
    }

    /// <summary>
    /// Resolve the number-format [Color] CSS for a numeric cell, or null. Mirrors
    /// the format-code lookup in GetFormattedCellValue (custom &lt;numFmt&gt; then
    /// built-in id), then delegates section/value selection to GetNumberFormatColor.
    /// </summary>
    private string? GetCellNumberFormatColor(Cell cell, CellFormat xf, Stylesheet stylesheet)
    {
        var numFmtId = xf.NumberFormatId?.Value ?? 0;
        if (numFmtId == 0) return null;

        var customFmt = stylesheet.NumberingFormats?.Elements<NumberingFormat>()
            .FirstOrDefault(nf => nf.NumberFormatId?.Value == numFmtId);
        var fmtCode = customFmt?.FormatCode?.Value ?? ResolveBuiltInFormat(numFmtId);
        if (fmtCode == null) return null;

        // Text cells take the text (@) section's [Color]; the positive/negative/
        // zero value-driven sections do not apply to them.
        var dt = cell.DataType?.Value;
        if (dt == CellValues.SharedString || dt == CellValues.InlineString
            || dt == CellValues.String || dt == CellValues.Boolean || dt == CellValues.Error)
            return GetTextSectionColor(fmtCode);

        if (!double.TryParse(cell.CellValue?.Text, System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out var numVal))
            return null;

        return GetNumberFormatColor(numVal, fmtCode);
    }

    /// <summary>
    /// Apply the text-format (@) section of a number-format code to a string
    /// value. The '@' placeholder is replaced by the cell text; quoted literals
    /// and escaped chars around it are emitted verbatim. Other format markers
    /// ([$..], [Color], _x fill placeholders) are stripped. Multi-section codes
    /// (pos;neg;zero;text) use the 4th (text) section when present.
    /// </summary>
    private static string ApplyTextFormat(string text, string fmtCode)
    {
        // The text section is the 4th in a multi-section code; if fewer sections
        // exist, use whichever section actually carries the '@' placeholder.
        if (fmtCode.Contains(';'))
        {
            var sections = fmtCode.Split(';');
            var sec = sections.Length >= 4 ? sections[3]
                : sections.FirstOrDefault(s => ContainsCharOutsideQuotes(s, '@'));
            if (sec == null) return text;
            fmtCode = sec;
        }

        // Strip non-literal markers that carry no displayable text.
        fmtCode = System.Text.RegularExpressions.Regex.Replace(fmtCode, @"\[[^\]]*\]", "");
        fmtCode = System.Text.RegularExpressions.Regex.Replace(fmtCode, @"_.", "");
        fmtCode = System.Text.RegularExpressions.Regex.Replace(fmtCode, @"\*.", "");

        var sb = new StringBuilder();
        bool inQuote = false;
        for (int i = 0; i < fmtCode.Length; i++)
        {
            var ch = fmtCode[i];
            if (ch == '"') { inQuote = !inQuote; continue; }
            if (inQuote) { sb.Append(ch); continue; }
            if (ch == '\\') { if (i + 1 < fmtCode.Length) sb.Append(fmtCode[++i]); continue; }
            if (ch == '@') { sb.Append(text); continue; }
            sb.Append(ch);
        }
        return sb.ToString();
    }

    private static string ApplyNumberFormat(double value, string fmtCode)
        => ApplyNumberFormat(value, fmtCode, autoMinus: false);

    // autoMinus: caller has already taken Math.Abs(value) and wants the
    // automatic negative sign PREPENDED to the whole section output (Excel's
    // single-section rule: a format with no explicit negative section is used
    // for negatives too, with a leading '-' added before everything). Literal
    // format chars (e.g. a '-' or '$' written in the code) are emitted verbatim
    // in position; this auto-minus is layered on top of that.
    private static string ApplyNumberFormat(double value, string fmtCode, bool autoMinus)
    {
        // Single-section negative: no ';' and no explicit [condition]. Excel
        // reuses the (positive) section for negatives, prepending an automatic
        // minus to the whole formatted result. Recurse on the magnitude with
        // autoMinus set so literal chars stay in place and the sign leads.
        if (value < 0 && !autoMinus && !fmtCode.Contains(';') &&
            !System.Text.RegularExpressions.Regex.IsMatch(fmtCode, @"\[[<>=]=?\d"))
        {
            return "-" + ApplyNumberFormat(Math.Abs(value), fmtCode, autoMinus: true);
        }

        // Handle multi-section format codes: positive;negative;zero
        if (fmtCode.Contains(';'))
        {
            var sections = fmtCode.Split(';');

            // Explicit [condition] sections (e.g. [>=100]"High: "0;[<0]"Low: "0;"Mid: "0).
            // Excel: when any section carries a [<op><num>] bracket, evaluate the
            // conditions IN DECLARATION ORDER; the first satisfied section applies,
            // and the last unconditioned section is the "else". This overrides the
            // positional positive/negative/zero convention below.
            if (System.Text.RegularExpressions.Regex.IsMatch(
                    fmtCode, @"\[[<>=]=?\d") )
            {
                for (int i = 0; i < sections.Length; i++)
                {
                    var sec = sections[i].Trim();
                    var condMatch = System.Text.RegularExpressions.Regex.Match(
                        sec, @"^\[(<=|>=|<>|<|>|=)(-?\d+\.?\d*)\]");
                    if (condMatch.Success)
                    {
                        var op = condMatch.Groups[1].Value;
                        var cmp = double.Parse(condMatch.Groups[2].Value,
                            System.Globalization.CultureInfo.InvariantCulture);
                        bool satisfied = op switch
                        {
                            "<" => value < cmp,
                            "<=" => value <= cmp,
                            ">" => value > cmp,
                            ">=" => value >= cmp,
                            "=" => value == cmp,
                            "<>" => value != cmp,
                            _ => false
                        };
                        if (satisfied)
                            return ApplyNumberFormat(Math.Abs(value), sec);
                    }
                    else
                    {
                        // Unconditioned (else) section — applies when no prior
                        // conditioned section matched.
                        return ApplyNumberFormat(Math.Abs(value), sec);
                    }
                }
                // No section matched and no else clause: fall back to first section.
                fmtCode = sections[0].Trim();
            }
            else
            {
                if (value < 0 && sections.Length >= 2)
                {
                    var negFmt = sections[1].Trim();
                    // If format already handles negative (has parens or minus), don't add extra minus
                    return ApplyNumberFormat(Math.Abs(value), negFmt);
                }
                if (value == 0 && sections.Length >= 3)
                {
                    var zeroFmt = sections[2].Trim();
                    // Quoted literal for zero section: "zero" → zero
                    if (zeroFmt.StartsWith('"') && zeroFmt.EndsWith('"'))
                        return zeroFmt[1..^1];
                    return ApplyNumberFormat(value, zeroFmt);
                }
                fmtCode = sections[0].Trim();
            }
        }

        // Strip [Color] markers: [Red], [Blue], [Green], [Color N], etc.
        fmtCode = System.Text.RegularExpressions.Regex.Replace(fmtCode, @"\[(Red|Blue|Green|Yellow|White|Black|Cyan|Magenta|Color\s*\d+)\]", "", System.Text.RegularExpressions.RegexOptions.IgnoreCase).Trim();

        // Strip [DBNumN] CJK-numeral modifiers cleanly (full numeral conversion is
        // out of scope; render the plain number rather than mangled bracket text).
        fmtCode = System.Text.RegularExpressions.Regex.Replace(fmtCode, @"\[DBNum\d+\]", "", System.Text.RegularExpressions.RegexOptions.IgnoreCase).Trim();

        // [$...] locale/currency specifiers. The [$<symbol>-<lcid>] form carries a
        // currency symbol (the text before the '-', e.g. "USD", "€", "¥") that real
        // Excel emits as a literal; preserve it as a quoted literal in place so the
        // downstream prefix/suffix extraction picks it up. The bare [$-409] form
        // (locale only, no symbol) is dropped entirely.
        fmtCode = System.Text.RegularExpressions.Regex.Replace(fmtCode, @"\[\$([^\]]*)\]", m =>
        {
            var inner = m.Groups[1].Value;
            var dash = inner.IndexOf('-');
            var sym = dash >= 0 ? inner[..dash] : inner;
            return string.IsNullOrEmpty(sym) ? "" : "\"" + sym + "\"";
        }).Trim();

        // Strip Excel numfmt special characters:
        // _X = space placeholder, *X = fill character, \X = literal character escape
        fmtCode = System.Text.RegularExpressions.Regex.Replace(fmtCode, @"_.", "").Trim();
        fmtCode = System.Text.RegularExpressions.Regex.Replace(fmtCode, @"\*.", "").Trim();
        fmtCode = System.Text.RegularExpressions.Regex.Replace(fmtCode, @"\\(.)", "$1").Trim();

        // Strip condition markers: [>100], [<=0], etc.
        fmtCode = System.Text.RegularExpressions.Regex.Replace(fmtCode, @"\[[<>=!]+\d+\.?\d*\]", "").Trim();

        // Handle parenthesis wrapping: ($#,##0.00) → prefix="(" suffix=")"
        if (fmtCode.StartsWith('(') && fmtCode.EndsWith(')'))
        {
            var inner = fmtCode[1..^1];
            return "(" + ApplyNumberFormat(value, inner) + ")";
        }

        var fmt = fmtCode.ToLowerInvariant();

        // Date/time formats may contain quoted literals (e.g. "D"d"D").
        // Skip prefix/suffix extraction for these — the date handler in
        // ApplyNumberFormatCore processes quotes via NormalizeDateFormatCase.
        if (ContainsDateTokenOutsideQuotes(fmtCode))
            return ApplyNumberFormatCore(value, fmtCode);

        // Extract currency/text prefix and suffix (e.g. "$", "€", "¥", or quoted strings like "USD ")
        var prefix = "";
        var suffix = "";
        var cleanFmt = fmtCode;
        // A literal leading sign ('-' or '+') written in the format string is a
        // verbatim character that must be emitted BEFORE the currency symbol
        // (Excel: "-$#,##0.00" → "-$2.00", not "$-2.00"). Pull it into the
        // prefix first so the currency symbol appends after it, preserving order.
        cleanFmt = cleanFmt.TrimStart();
        if (cleanFmt.StartsWith('-')) { prefix = "-"; cleanFmt = cleanFmt[1..]; }
        else if (cleanFmt.StartsWith('+')) { prefix = "+"; cleanFmt = cleanFmt[1..]; }
        // Handle literal characters: $, ¥, €, £
        foreach (var sym in new[] { "$", "¥", "€", "£", "₹" })
        {
            if (cleanFmt.Contains(sym))
            {
                var idx = cleanFmt.IndexOf(sym);
                var hashIdx = cleanFmt.IndexOf('#');
                var zeroIdx = cleanFmt.IndexOf('0');
                var firstDigit = (hashIdx >= 0 && zeroIdx >= 0) ? Math.Min(hashIdx, zeroIdx)
                    : Math.Max(hashIdx, zeroIdx);
                if (firstDigit < 0 || idx <= firstDigit)
                    prefix += sym;
                else
                    suffix = sym;
                cleanFmt = cleanFmt.Replace(sym, "");
            }
        }
        // Currency-symbol extraction can leave an EMPTY quote remnant ("") where a
        // quoted symbol used to be (accounting "$"-prefixed sections). Drop those so
        // the downstream paren / quoted-literal checks see the real leading token
        // (e.g. "" ( #,##0.00 ) -> ( #,##0.00 )).
        cleanFmt = cleanFmt.Replace("\"\"", "").Trim();
        // Handle quoted prefix/suffix: "USD ". A literal space immediately after a
        // quoted prefix (e.g. [$USD-409] #,##0.00 → "USD" #,##0.00) is part of the
        // displayed prefix and must survive the later cleanFmt.Trim().
        var quoteMatch = System.Text.RegularExpressions.Regex.Match(cleanFmt, "^\"([^\"]+)\"( *)");
        if (quoteMatch.Success) { prefix += quoteMatch.Groups[1].Value + quoteMatch.Groups[2].Value; cleanFmt = cleanFmt[quoteMatch.Length..]; }
        var quoteSuffix = System.Text.RegularExpressions.Regex.Match(cleanFmt, "\"([^\"]+)\"$");
        if (quoteSuffix.Success) { suffix = quoteSuffix.Groups[1].Value + suffix; cleanFmt = cleanFmt[..^quoteSuffix.Length]; }

        // Re-check for a paren-wrapped numeric pattern after the _X/*X/\X strips and
        // currency-symbol extraction. Accounting negative sections such as
        // "_($* (#,##0.00_)" lose their "_(" wrappers to the strip passes, leaving a
        // bare leading "(" (and possibly trailing ")") that the early line-2225 check
        // (which ran on the pre-strip fmtCode) never saw. Native Excel keeps the paren.
        cleanFmt = cleanFmt.Trim();
        var parenOpen = false;
        var parenClose = false;
        if (cleanFmt.StartsWith('('))
        {
            parenOpen = true;
            cleanFmt = cleanFmt[1..];
            if (cleanFmt.EndsWith(')')) { parenClose = true; cleanFmt = cleanFmt[..^1]; }
        }

        // Handle +/- prefix in format (e.g. "+0.0%", "-#,##0")
        cleanFmt = cleanFmt.Trim();
        if (cleanFmt.StartsWith('+'))
        { prefix += "+"; cleanFmt = cleanFmt[1..]; }
        else if (cleanFmt.StartsWith('-'))
        { prefix += "-"; cleanFmt = cleanFmt[1..]; }

        // Pure text format (only quoted prefix/suffix, no numeric pattern)
        if (string.IsNullOrEmpty(cleanFmt.Trim()))
            return prefix + suffix;

        var formatted = ApplyNumberFormatCore(value, cleanFmt.Trim());
        // Accounting paren wraps the numeric core (and trailing currency), keeping
        // "(1,234.56" contiguous even when a left-aligned "$" prefix is present.
        if (parenOpen) formatted = "(" + formatted + suffix + (parenClose ? ")" : "");
        else formatted += suffix;
        // For single-section formats with currency prefix, negative sign goes before the prefix
        if (value < 0 && prefix.Length > 0 && formatted.StartsWith('-'))
            return "-" + prefix + formatted[1..];
        return prefix + formatted;
    }

    private static int Gcd(int a, int b)
    {
        a = Math.Abs(a); b = Math.Abs(b);
        while (b != 0) { var t = b; b = a % b; a = t; }
        return a == 0 ? 1 : a;
    }

    private static string ApplyNumberFormatCore(double value, string fmtCode)
    {
        var fmt = fmtCode.ToLowerInvariant();

        // Percentage formats
        if (fmt.Contains('%'))
        {
            var pctVal = value * 100;
            var decimals = CountDecimalPlaces(fmtCode);
            return pctVal.ToString($"F{decimals}") + "%";
        }

        // Fraction formats: "# ?/?", "# ??/??", "?/?" etc.
        // Denominator-digit count = number of '?' after the slash → max denominator
        // (1→9, 2→99, 3→999). Find the best rational approximation within that limit.
        // Fixed-denominator fraction: a literal integer denominator after the
        // slash (e.g. "# ?/8", "# ??/16", "?/100"). Round the fractional part to
        // the nearest numerator/N and show numerator/N (numerator NOT reduced).
        // The whole-part token before the space ('#'/'0') governs the integer split.
        var fixedFracMatch = System.Text.RegularExpressions.Regex.Match(
            fmtCode, @"(?:^|\s)([#0]*)\s*[?#]*\s*/\s*(\d+)\s*$");
        if (fixedFracMatch.Success)
        {
            int fixedDen = int.Parse(fixedFracMatch.Groups[2].Value);
            if (fixedDen >= 1)
            {
                bool hasWholePart = fmtCode.TrimStart().Contains(' ');
                bool fneg = value < 0;
                double fabs = Math.Abs(value);
                long fwhole = (long)Math.Floor(fabs);
                double ffrac = fabs - fwhole;
                long num = (long)Math.Round(ffrac * fixedDen);
                if (num == fixedDen) { fwhole += 1; num = 0; }
                if (!hasWholePart)
                {
                    // No integer split: fold the whole part into the numerator.
                    num += fwhole * fixedDen;
                    fwhole = 0;
                }

                var fsb = new StringBuilder();
                if (fneg) fsb.Append('-');
                if (num == 0)
                    fsb.Append(fwhole.ToString(System.Globalization.CultureInfo.InvariantCulture));
                else if (!hasWholePart || fwhole == 0)
                    fsb.Append($"{num}/{fixedDen}");
                else
                    fsb.Append($"{fwhole} {num}/{fixedDen}");
                return fsb.ToString();
            }
        }

        // Variable-denominator fraction. Excel treats '#', '?' and '0' all as
        // valid fraction digit placeholders, so detect a slash with one or more
        // placeholders on BOTH sides (e.g. "# ??/??", "# #/#", "#/#", "?/100"'s
        // numerator side, "0/0"). Date formats (m/d/yyyy) never match — their
        // tokens are letters, not placeholders. The number of denominator-side
        // placeholders bounds the max denominator (1→9, 2→99, …).
        var fracMatch = System.Text.RegularExpressions.Regex.Match(fmtCode, @"[?#0]+\s*/\s*([?#0]+)");
        if (fracMatch.Success)
        {
            int denomDigits = fracMatch.Groups[1].Value.Count(c => c == '?' || c == '#' || c == '0');
            int maxDenom = (int)Math.Pow(10, denomDigits) - 1;
            if (maxDenom < 1) maxDenom = 9;
            bool neg = value < 0;
            double abs = Math.Abs(value);
            long whole = (long)Math.Floor(abs);
            double frac = abs - whole;

            // Continued-fraction convergents give the simplest fraction within the
            // denominator limit (matches Excel, e.g. 0.14159 → 1/7 not 14/99).
            long h0 = 0, h1 = 1, k0 = 1, k1 = 0;
            double b = frac;
            for (int guard = 0; guard < 64; guard++)
            {
                long a = (long)Math.Floor(b);
                long h2 = a * h1 + h0, k2 = a * k1 + k0;
                if (k2 > maxDenom || k2 == 0) break;
                h0 = h1; h1 = h2; k0 = k1; k1 = k2;
                double rem = b - a;
                if (rem < 1e-12) break;
                b = 1 / rem;
            }
            int bestNum = (int)h1, bestDen = (int)Math.Max(k1, 1);
            // Reduce the chosen fraction
            if (bestNum != 0)
            {
                int g = Gcd(bestNum, bestDen);
                bestNum /= g; bestDen /= g;
            }
            // Numerator rounded up to a whole unit (e.g. 0.999 → 1/1) folds into whole part
            if (bestNum == bestDen && bestNum != 0) { whole += 1; bestNum = 0; }

            var sb = new StringBuilder();
            if (neg) sb.Append('-');
            if (bestNum == 0)
                sb.Append(whole.ToString(System.Globalization.CultureInfo.InvariantCulture));
            else if (whole == 0)
                sb.Append($"{bestNum}/{bestDen}");
            else
                sb.Append($"{whole} {bestNum}/{bestDen}");
            return sb.ToString();
        }

        // Digit-placeholder-only pattern (only '?' / spaces, no '0' or '#'). Excel's
        // '?' shows a space for an insignificant digit, so a section like "??" formats
        // any value (notably 0 in an accounting zero section) as blanks — no digits.
        var phTrim = fmtCode.Trim();
        if (phTrim.Length > 0 && phTrim.All(c => c == '?' || c == ' '))
            return new string(' ', phTrim.Length);

        // Elapsed-time formats with a bracketed lead token: [h]/[hh] (total hours),
        // [m]/[mm] (total minutes), [s]/[ss] (total seconds) — none clock-wrapped.
        // The leading bracket token carries the TOTAL elapsed unit; any following
        // non-bracketed :mm / :ss tokens are the remainder within the next-smaller
        // unit. Must run before the generic date path, which would mangle "[mm]".
        var elapsedMatch = System.Text.RegularExpressions.Regex.Match(
            fmtCode, @"^\[(h+|m+|s+)\](.*)$", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (elapsedMatch.Success)
        {
            long totalSeconds = (long)Math.Round(value * 86400);
            var unit = char.ToLowerInvariant(elapsedMatch.Groups[1].Value[0]);
            var rest = elapsedMatch.Groups[2].Value;
            long lead = unit switch
            {
                'h' => totalSeconds / 3600,
                'm' => totalSeconds / 60,
                _ => totalSeconds,
            };
            var parts = new List<string> { lead.ToString(System.Globalization.CultureInfo.InvariantCulture) };
            // Remainder tokens: each subsequent m/mm or s/ss after the bracket.
            // [h] → following mm is minutes-of-hour, ss is seconds-of-minute.
            // [mm] → following ss is seconds-of-minute.
            foreach (System.Text.RegularExpressions.Match tok in
                     System.Text.RegularExpressions.Regex.Matches(rest, "(mm?|ss?)",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase))
            {
                var t = char.ToLowerInvariant(tok.Value[0]);
                if (t == 'm' && unit == 'h') parts.Add(((totalSeconds / 60) % 60).ToString("D2"));
                else if (t == 's') parts.Add((totalSeconds % 60).ToString("D2"));
            }
            return string.Join(":", parts);
        }

        // Date formats (serial number → DateTime)
        if (fmt.Contains('y') || fmt.Contains('m') || fmt.Contains('d') || fmt.Contains('h'))
        {
            try
            {
                var dt = DateTime.FromOADate(value);
                // Context-sensitive m/mm: after h → minute, otherwise → month
                // Strategy: mark minute 'm' as '\x01' placeholder, then convert remaining m→M
                var dotnetFmt = NormalizeDateFormatCase(fmtCode);
                // Step 1: Replace h:mm and h:m patterns → mark minutes as placeholder
                dotnetFmt = System.Text.RegularExpressions.Regex.Replace(dotnetFmt, @"([hH]+)([:.])(mm?)", m =>
                    m.Groups[1].Value + m.Groups[2].Value + new string('\x01', m.Groups[3].Value.Length));
                // Also handle mm:ss (mm before ss is also minutes)
                dotnetFmt = System.Text.RegularExpressions.Regex.Replace(dotnetFmt, @"(mm?)([:.])(ss?)", m =>
                    new string('\x01', m.Groups[1].Value.Length) + m.Groups[2].Value + m.Groups[3].Value);
                // Step 2: Convert remaining m/mm to M/MM (month)
                dotnetFmt = dotnetFmt.Replace("mmmm", "MMMM").Replace("mmm", "MMM")
                    .Replace("mm", "MM").Replace("m", "M");
                // Step 3: Restore minute placeholders
                dotnetFmt = dotnetFmt.Replace("\x01\x01", "mm").Replace("\x01", "m");
                // Step 4: Other conversions
                // If AM/PM format (has 't' outside quotes), use h (12h); otherwise use H (24h)
                if (!ContainsCharOutsideQuotes(dotnetFmt, 't'))
                    dotnetFmt = dotnetFmt.Replace("hh", "HH").Replace("h", "H");
                dotnetFmt = dotnetFmt.Replace("dddd", "dddd").Replace("ddd", "ddd").Replace("dd", "dd");
                return dt.ToString(dotnetFmt, System.Globalization.CultureInfo.InvariantCulture);
            }
            catch { return value.ToString(); }
        }

        // Scientific notation
        if (fmt.Contains("e+") || fmt.Contains("e-"))
        {
            var decimals = CountDecimalPlaces(fmtCode);
            if (value == 0) return decimals > 0 ? $"0.{new string('0', decimals)}E+00" : "0E+00";
            var eIdx = fmt.IndexOf("e+", StringComparison.Ordinal);
            if (eIdx < 0) eIdx = fmt.IndexOf("e-", StringComparison.Ordinal);
            var expDigits = eIdx >= 0 ? fmtCode[(eIdx + 2)..].Count(c => c == '0') : 2;
            var exp = (int)Math.Floor(Math.Log10(Math.Abs(value)));

            // Engineering notation: when the integer-mantissa width N (# / 0
            // placeholders before the 'E') is > 1, Excel rounds the exponent DOWN
            // to a multiple of N (e.g. ##0.0E+0 → exponent multiple of 3, mantissa
            // 1-999). Standard 0.00E+00 (N=1) keeps the per-decade exponent.
            var mantSpec = eIdx >= 0 ? fmtCode[..eIdx] : "";
            var mantDot = mantSpec.IndexOf('.');
            var mantInt = mantDot >= 0 ? mantSpec[..mantDot] : mantSpec;
            var mantWidth = mantInt.Count(c => c == '#' || c == '0');
            if (mantWidth < 1) mantWidth = 1;
            if (mantWidth > 1)
                exp = (int)(Math.Floor((double)exp / mantWidth) * mantWidth);

            var mantissa = value / Math.Pow(10, exp);
            var expStr = exp >= 0 ? $"+{exp.ToString().PadLeft(expDigits, '0')}" : $"-{Math.Abs(exp).ToString().PadLeft(expDigits, '0')}";
            return $"{mantissa.ToString($"F{decimals}")}E{expStr}";
        }

        // Trailing comma scaling: each trailing comma divides value by 1000
        // e.g. "#," = ÷1000, "#,," = ÷1000000, "#,##0," = thousands + ÷1000
        var trailingCommas = 0;
        var fmtTrimmed = fmtCode.TrimEnd();
        while (fmtTrimmed.EndsWith(',')) { trailingCommas++; fmtTrimmed = fmtTrimmed[..^1]; }
        if (trailingCommas > 0)
        {
            value /= Math.Pow(1000, trailingCommas);
            fmtCode = fmtTrimmed;
        }

        // Digit-group integer formats with embedded literal separators, e.g. phone
        // "###-####" → "555-1234", SSN "000-00-0000" → "123-45-6789". An integer-only
        // format (no '.', no thousands ',') whose '#'/'0' placeholders are interleaved
        // with literal chars maps the number's digits RIGHT-TO-LEFT into the
        // placeholder positions, emitting the literals in place. '0' placeholders with
        // no remaining digit emit '0'; surplus '#' emit nothing.
        if (!fmtCode.Contains('.') && !fmtCode.Contains(',')
            && (fmtCode.Contains('#') || fmtCode.Contains('0'))
            && fmtCode.Contains('-'))
        {
            var digits = ((long)Math.Round(Math.Abs(value)))
                .ToString(System.Globalization.CultureInfo.InvariantCulture);
            var outChars = new List<char>();
            int di = digits.Length - 1;
            for (int i = fmtCode.Length - 1; i >= 0; i--)
            {
                var c = fmtCode[i];
                if (c == '#' || c == '0')
                {
                    if (di >= 0) { outChars.Add(digits[di]); di--; }
                    else if (c == '0') outChars.Add('0');
                }
                else outChars.Add(c);
            }
            // Any leftover leading digits (more digits than placeholders) prepend to
            // the front, matching Excel which never truncates the number.
            while (di >= 0) { outChars.Add(digits[di]); di--; }
            outChars.Reverse();
            var s = new string(outChars.ToArray());
            return value < 0 ? "-" + s : s;
        }

        // Numeric with thousands separator and/or decimals
        bool hasThousands = fmtCode.Contains(',') && (fmtCode.Contains('#') || fmtCode.Contains('0'));
        var numDecimals = CountDecimalPlaces(fmtCode);

        // Leading-zero placeholders in the integer portion (e.g. "00000" → pad
        // the integer part to 5 digits: 42 → "00042"). Excel zero-pads the
        // integer part to the count of '0' placeholders before any decimal point.
        int intZeroPad = CountIntegerZeroPlaceholders(fmtCode);

        if (hasThousands)
            return TrimOptionalDecimals(PadIntegerPart(value.ToString($"N{numDecimals}", System.Globalization.CultureInfo.InvariantCulture), intZeroPad), fmtCode);
        if (numDecimals > 0)
            return TrimOptionalDecimals(PadIntegerPart(value.ToString($"F{numDecimals}", System.Globalization.CultureInfo.InvariantCulture), intZeroPad), fmtCode);

        // @ = text format — return raw
        if (fmt == "@") return value.ToString();

        // Integer placeholder format ("0", "00000", …)
        if (intZeroPad > 0)
            return PadIntegerPart(((long)Math.Round(value)).ToString(System.Globalization.CultureInfo.InvariantCulture), intZeroPad);

        return value.ToString();
    }

    /// <summary>
    /// Count '0' digit placeholders in the integer portion (before the first
    /// '.') of an Excel number-format code. Thousands-separator commas are not
    /// counted as digits.
    /// </summary>
    private static int CountIntegerZeroPlaceholders(string fmtCode)
    {
        int dotIdx = fmtCode.IndexOf('.');
        var intPart = dotIdx >= 0 ? fmtCode[..dotIdx] : fmtCode;
        return intPart.Count(c => c == '0');
    }

    /// <summary>
    /// Zero-pad the integer part of an already-formatted numeric string to at
    /// least <paramref name="minDigits"/> digits, preserving any sign, thousands
    /// separators in the original are left intact only when no padding is needed.
    /// </summary>
    private static string PadIntegerPart(string formatted, int minDigits)
    {
        if (minDigits <= 0) return formatted;
        bool neg = formatted.StartsWith('-');
        var body = neg ? formatted[1..] : formatted;
        int dotIdx = body.IndexOf('.');
        var intPart = dotIdx >= 0 ? body[..dotIdx] : body;
        var rest = dotIdx >= 0 ? body[dotIdx..] : "";
        // Count only digit characters when padding (ignore thousands separators).
        int digitCount = intPart.Count(char.IsDigit);
        if (digitCount < minDigits)
            intPart = intPart.PadLeft(intPart.Length + (minDigits - digitCount), '0');
        return (neg ? "-" : "") + intPart + rest;
    }

    private static int CountDecimalPlaces(string fmtCode)
    {
        var dotIdx = fmtCode.IndexOf('.');
        if (dotIdx < 0) return 0;
        int count = 0;
        for (int i = dotIdx + 1; i < fmtCode.Length; i++)
        {
            // '?' is a digit placeholder like '0'/'#' (it pads with a space rather
            // than a zero, but still counts toward the decimal-place count).
            if (fmtCode[i] == '0' || fmtCode[i] == '#' || fmtCode[i] == '?') count++;
            else break;
        }
        return count;
    }

    /// <summary>
    /// Apply optional fractional-digit-placeholder semantics to an already
    /// fully-zero-padded numeric string. The fractional part of an Excel number
    /// format mixes three placeholder kinds, read left-to-right after the '.':
    ///   '0' = required digit (always shown, kept as a zero),
    ///   '#' = optional digit (trailing zero suppressed — dropped entirely),
    ///   '?' = optional digit (trailing zero suppressed but padded with a
    ///         non-breaking space so decimal columns align).
    /// The incoming <paramref name="formatted"/> was produced via "F{n}"/"N{n}"
    /// (n = total placeholder count) so it is already rounded and fully padded.
    /// We walk the placeholders right-to-left, trimming trailing zero digits whose
    /// placeholder is '#' or '?', stopping at the first '0' placeholder or any
    /// significant (non-zero) digit. If every fractional digit is dropped, the
    /// decimal point is also removed (the trailing '?' alignment spaces remain).
    /// </summary>
    private static string TrimOptionalDecimals(string formatted, string fmtCode)
    {
        int dotIdx = fmtCode.IndexOf('.');
        if (dotIdx < 0) return formatted;
        // Collect the fractional placeholder kinds in order.
        var placeholders = new List<char>();
        for (int i = dotIdx + 1; i < fmtCode.Length; i++)
        {
            var c = fmtCode[i];
            if (c == '0' || c == '#' || c == '?') placeholders.Add(c);
            else break;
        }
        if (placeholders.Count == 0) return formatted;
        // No optional placeholders → nothing to trim (all required '0').
        if (!placeholders.Contains('#') && !placeholders.Contains('?')) return formatted;

        int fmtDot = formatted.LastIndexOf('.');
        if (fmtDot < 0) return formatted;
        var intPart = formatted[..fmtDot];
        var fracDigits = formatted[(fmtDot + 1)..].ToCharArray();
        // fracDigits length should equal placeholders.Count (both = numDecimals).
        var trailing = new StringBuilder(); // '?' alignment spaces, emitted after.
        int p = Math.Min(placeholders.Count, fracDigits.Length) - 1;
        for (; p >= 0; p--)
        {
            var kind = placeholders[p];
            if (kind == '0') break;                 // required digit — stop trimming
            if (fracDigits[p] != '0') break;        // significant digit — stop trimming
            if (kind == '?') trailing.Insert(0, '\u00A0'); // '?' pads with a non-breaking space for column alignment
            fracDigits[p] = '\0';
        }
        var kept = new string(fracDigits).Replace("\0", "");
        if (kept.Length == 0)
            return intPart + trailing.ToString();   // all fractional digits gone → drop '.'
        return intPart + "." + kept + trailing.ToString();
    }

    /// <summary>
    /// Returns true if fmtCode contains date/time tokens (y, m, d, h, s) outside
    /// double-quoted strings. Used to route date formats past prefix/suffix extraction.
    /// </summary>
    private static bool ContainsDateTokenOutsideQuotes(string fmtCode)
    {
        bool inQuote = false;
        foreach (var ch in fmtCode)
        {
            if (ch == '"') { inQuote = !inQuote; continue; }
            if (!inQuote)
            {
                var lower = char.ToLowerInvariant(ch);
                if (lower is 'y' or 'm' or 'd' or 'h' or 's') return true;
            }
        }
        return false;
    }

    /// <summary>
    /// Returns true if ch appears outside double-quoted strings in fmtCode.
    /// </summary>
    private static bool ContainsCharOutsideQuotes(string fmtCode, char target)
    {
        bool inQuote = false;
        foreach (var ch in fmtCode)
        {
            if (ch == '"') { inQuote = !inQuote; continue; }
            if (!inQuote && ch == target) return true;
        }
        return false;
    }

    /// <summary>
    /// Normalize Excel date/time format specifiers to .NET-compatible case
    /// and replace AM/PM → tt, A/P → t outside quoted strings.
    /// </summary>
    private static string NormalizeDateFormatCase(string fmtCode)
    {
        var sb = new StringBuilder(fmtCode.Length);
        bool inQuote = false;
        for (int i = 0; i < fmtCode.Length; i++)
        {
            var ch = fmtCode[i];
            if (ch == '"') { inQuote = !inQuote; sb.Append(ch); continue; }
            if (inQuote) { sb.Append(ch); continue; }
            // AM/PM → tt (check before single-char A/P)
            if ((ch == 'A' || ch == 'a') && i + 4 < fmtCode.Length
                && (fmtCode[i + 1] == 'M' || fmtCode[i + 1] == 'm')
                && fmtCode[i + 2] == '/'
                && (fmtCode[i + 3] == 'P' || fmtCode[i + 3] == 'p')
                && (fmtCode[i + 4] == 'M' || fmtCode[i + 4] == 'm'))
            {
                sb.Append("tt"); i += 4; continue;
            }
            // A/P → t
            if ((ch == 'A' || ch == 'a') && i + 2 < fmtCode.Length
                && fmtCode[i + 1] == '/'
                && (fmtCode[i + 2] == 'P' || fmtCode[i + 2] == 'p'))
            {
                sb.Append('t'); i += 2; continue;
            }
            sb.Append(ch switch { 'Y' => 'y', 'D' => 'd', 'S' => 's', 'M' => 'm', 'H' => 'h', _ => ch });
        }
        return sb.ToString();
    }

    // ==================== CSS ====================

    private string GenerateExcelCss()
    {
        // Read default font from workbook styles (font index 0)
        var defFontName = OfficeDefaultFonts.MinorLatin;
        var defFontSize = OfficeDefaultFonts.ExcelBodySizePt;
        var stylesheet = _doc.WorkbookPart?.WorkbookStylesPart?.Stylesheet;
        if (stylesheet?.Fonts != null && stylesheet.Fonts.Elements<Font>().Any())
        {
            var f0 = stylesheet.Fonts.Elements<Font>().First();
            if (f0.FontName?.Val?.Value != null) defFontName = CssSanitize(f0.FontName.Val.Value);
            if (f0.FontSize?.Val?.Value != null) defFontSize = f0.FontSize.Val.Value.ToString("0.##");
        }
        return $$"""
        * { margin: 0; padding: 0; box-sizing: border-box; }
        html, body { height: 100%; }
        body {
            font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif;
            background: #f0f0f0;
            color: #333;
            display: flex;
            flex-direction: column;
            min-height: 100vh;
        }
        .file-title {
            padding: 12px 20px;
            font-size: 14px;
            font-weight: 600;
            background: #217346;
            color: #fff;
        }
        .sheet-tabs {
            display: flex;
            background: #e0e0e0;
            border-top: 1px solid #ccc;
            overflow-x: auto;
            padding: 0 8px;
            flex-shrink: 0;
            position: sticky;
            bottom: 0;
            z-index: 10;
        }
        .sheet-tab {
            --tab-color: #e8e8e8;
            padding: 8px 16px;
            font-size: 12px;
            cursor: pointer;
            border: 1px solid #bbb;
            border-top: none;
            background: var(--tab-color);
            color: #fff;
            margin-bottom: 0;
            border-radius: 0 0 3px 3px;
            white-space: nowrap;
            user-select: none;
            position: relative;
            transition: background 0.15s, color 0.15s;
        }
        .sheet-tab[style*="--tab-color:#e8e8e8"], .sheet-tab:not([style*="--tab-color"]) {
            color: #333;
        }
        .sheet-tab:hover { opacity: 0.85; }
        .sheet-tab.active {
            background: linear-gradient(to bottom, #fff 60%, color-mix(in srgb, var(--tab-color) 30%, #fff)) !important;
            color: #333 !important;
            border-color: #aaa;
            border-bottom: 3px solid var(--tab-color);
            font-weight: 600;
        }
        .sheet-slider { flex: 1; position: relative; overflow: hidden; display: flex; flex-direction: column; min-height: 0; }
        .sheet-content { background: #fff; display: none; flex: 1; min-height: 0; }
        .sheet-content.active { display: flex; flex-direction: column; }
        .table-wrapper {
            flex: 1;
            overflow: auto;
            min-height: 0;
            background: #fff;
        }
        table {
            border-collapse: collapse;
            font-size: {{defFontSize}}px;
            font-family: '{{defFontName}}', 'Segoe UI', sans-serif;
            table-layout: fixed;
        }
        .row-header-col { width: 30pt; }
        th {
            background: #f8f8f8;
            border: 1px solid #e0e0e0;
            font-weight: normal;
            color: #666;
            font-size: 10px;
            text-align: center;
            padding: 2px 4px;
        }
        .corner-cell { background: #f0f0f0; z-index: 4; }
        .col-header {
            position: sticky;
            top: 0;
            z-index: 3;
            background: #f8f8f8;
            min-width: 50px;
            cursor: s-resize;
        }
        .row-header {
            position: sticky;
            left: 0;
            z-index: 2;
            background: #f8f8f8;
            min-width: 40px;
            cursor: e-resize;
            /* Drop right border so the data cell's own (often darker) left border shows through.
               Otherwise, with border-collapse, the row-header's light grey right border can win
               the collapse contest and erase the merged-cell left border (rowspan cells especially). */
            border-right: none;
        }
        td {
            /* Default gridlines are painted with inset box-shadow instead of
               border, so they do NOT participate in border-collapse tie-breaking.
               Explicit OOXML borders (rendered as inline border styles on cells
               with an OOXML style) always win at cell boundaries; missing cells
               / style-0 cells no longer erase neighbours' black borders via the
               CSS position-based tie-break. Right+bottom gridlines are owned by
               each cell; first-row top and first-col left gridlines are added
               via the :first-child rules below. Scoped to table:not(.no-grid) so
               sheets with showGridLines=false suppress the default gridlines while
               still honouring explicit OOXML cell borders (inline styles). */
            padding: 1px 4px;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
            vertical-align: bottom;
            max-width: 500px;
            word-break: break-all; /* CJK text wrapping support */
        }
        table:not(.no-grid) td { box-shadow: inset -1px -1px 0 #e0e0e0; }
        /* Text spill: a left/general text cell with empty right-neighbours paints
           its overflow across them (Excel fidelity). The td stays 1 column wide so
           borders/gridlines/merges are unaffected; overflow:visible lets the inner
           span bleed into the (empty) neighbour cells. The span clips at max-width =
           own width + summed empty-neighbour widths, stopping before the first
           occupied cell — matching real Excel. */
        td.spill { overflow: visible; }
        td.spill .spill-text {
            display: inline-block;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: clip;
            vertical-align: bottom;
        }
        /* Shrink-to-fit: the font is compressed (inline font-size) so the whole
           string fits the column. Keep it inline-block + clip (never ellipsis) so
           the full text shows; the td's ellipsis rule must not re-truncate it. */
        .shrink-fit {
            display: inline-block;
            white-space: nowrap;
            text-overflow: clip;
            overflow: visible;
        }
        table:not(.no-grid) tbody tr:first-child td { box-shadow: inset -1px -1px 0 #e0e0e0, inset 0 1px 0 #e0e0e0; }
        table:not(.no-grid) tr td:first-of-type { box-shadow: inset -1px -1px 0 #e0e0e0, inset 1px 0 0 #e0e0e0; }
        table:not(.no-grid) tbody tr:first-child td:first-of-type { box-shadow: inset -1px -1px 0 #e0e0e0, inset 1px 1px 0 #e0e0e0; }
        .empty-sheet {
            padding: 40px;
            text-align: center;
            color: #999;
            font-size: 14px;
        }
        /* Chart containers */
        .chart-container {
            margin: 16px auto;
            background: #fff;
            border: 1px solid #e0e0e0;
            border-radius: 6px;
            padding: 12px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.08);
        }
        .chart-container svg { display: block; }
        /* Truncation warning */
        .truncation-warning {
            padding: 8px 16px;
            background: #FFF3CD;
            color: #856404;
            border: 1px solid #FFEEBA;
            font-size: 12px;
            text-align: center;
            margin: 4px 0;
        }
        /* Screen reader only */
        .sr-only { position:absolute; clip:rect(0 0 0 0); width:1px; height:1px; overflow:hidden; }
        /* Print styles */
        @media print {
            .file-title, .sheet-tabs { display: none !important; }
            .table-wrapper { max-height: none !important; overflow: visible !important; flex: none !important; }
            body { background: #fff !important; min-height: auto !important; }
            .sheet-content { display: block !important; flex: none !important; }
            td { max-width: none !important; white-space: normal !important; overflow: visible !important; }
        }
        """;
    }

    // ==================== JavaScript ====================

    private static string GenerateExcelJs() => """
        function switchSheet(idx) {
            document.querySelectorAll('.sheet-tab').forEach(function(t) {
                t.classList.toggle('active', parseInt(t.getAttribute('data-sheet')) === idx);
            });
            document.querySelectorAll('.sheet-content').forEach(function(c) {
                c.classList.toggle('active', parseInt(c.getAttribute('data-sheet')) === idx);
            });
            window.scrollTo(0, 0);
        }
        // Fix frozen row sticky top values using actual rendered heights
        document.querySelectorAll('.table-wrapper table').forEach(function(table) {
            var thead = table.querySelector('thead');
            if (!thead) return;
            var theadH = thead.offsetHeight;
            var cumTop = theadH;
            var frozen = table.querySelectorAll('tr[data-frozen]');
            frozen.forEach(function(tr) {
                tr.querySelectorAll('th, td').forEach(function(cell) {
                    if (cell.style.position === 'sticky') cell.style.top = cumTop + 'px';
                });
                cumTop += tr.offsetHeight;
            });
        });
        """;

    // ==================== Utility ====================

    // CONSISTENCY(html-encode): shared plain entity-encoder lives in Core/HtmlPreviewHelper.
    private static string HtmlEncode(string text) => HtmlPreviewHelper.HtmlEncode(text);

    /// <summary>HtmlEncode + convert newlines to br for cell display</summary>
    private static string CellHtml(string text)
    {
        var encoded = HtmlEncode(text);
        if (encoded.Contains('\n')) encoded = encoded.Replace("\n", "<br>");
        // Browsers collapse runs of literal spaces; real Excel preserves them.
        // Encode leading spaces and runs of 2+ spaces as &nbsp; so the cell text
        // keeps its original spacing (single interior spaces left as-is to allow
        // normal wrapping). Cheap fast-path: only touch strings that actually
        // contain a double space or start with one.
        if (encoded.StartsWith(" ") || encoded.Contains("  "))
        {
            // Any leading-space run, or any interior run of 2+ spaces, becomes
            // that many &nbsp;. A lone interior space stays a normal space.
            encoded = Regex.Replace(encoded, @"^ +|  +",
                m => string.Concat(System.Linq.Enumerable.Repeat("&nbsp;", m.Length)));
        }
        return encoded;
    }

    /// <summary>
    /// When a cell's formula is a HYPERLINK(url, [friendly]) call, build a blue,
    /// underlined &lt;a href&gt; matching how real Excel renders it (the friendly
    /// text, or the url when no friendly arg). <paramref name="display"/> is the
    /// already-formatted cell value (Excel evaluates HYPERLINK to its friendly
    /// text, so it usually equals the friendly arg). Returns null when the cell
    /// carries no HYPERLINK formula. Pre-encoded HTML.
    /// </summary>
    private static string? TryBuildHyperlinkFormulaHtml(Cell? cell, string display)
    {
        var formula = cell?.CellFormula?.Text;
        if (string.IsNullOrEmpty(formula)) return null;
        // Match HYPERLINK( "url" [, "friendly"] ) — case-insensitive, optional
        // leading '='. Args are double-quoted string literals here (the common
        // authored form); non-literal arg expressions are out of scope.
        var m = System.Text.RegularExpressions.Regex.Match(
            formula,
            "^=?\\s*HYPERLINK\\s*\\(\\s*\"([^\"]*)\"\\s*(?:,\\s*\"([^\"]*)\"\\s*)?\\)\\s*$",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (!m.Success) return null;

        var url = m.Groups[1].Value;
        // Reject anything but a safe scheme so the url never becomes an href XSS sink.
        if (!Core.HyperlinkUriValidator.IsSafeScheme(url)) return null;

        // Friendly text: explicit 2nd arg, else the evaluated display, else the url.
        var friendly = m.Groups[2].Success && m.Groups[2].Value.Length > 0
            ? m.Groups[2].Value
            : (!string.IsNullOrEmpty(display) ? display : url);
        return $"<a href=\"{HtmlEncode(url)}\" style=\"color:#0563C1;text-decoration:underline\">{HtmlEncode(friendly)}</a>";
    }

    /// <summary>Get data-formula attribute for cells with formulas (for inline editing).</summary>
    private static string GetFormulaAttr(Cell? cell)
    {
        var formula = cell?.CellFormula?.Text;
        if (string.IsNullOrEmpty(formula)) return "";
        return $" data-formula=\"={HtmlEncode(formula)}\"";
    }

    private static string BuildCellContent(string cellRef, string value,
        Dictionary<string, string> dataBarMap, Dictionary<string, string> iconSetMap)
    {
        var hasBar = dataBarMap.TryGetValue(cellRef, out var barEntry);
        var hasIcon = iconSetMap.TryGetValue(cellRef, out var iconEntry);
        if (!hasBar && !hasIcon) return CellHtml(value);

        // Parse "showValue|html" format
        var barShowValue = true;
        var barHtml = "";
        if (hasBar && barEntry != null)
        {
            var sep = barEntry.IndexOf('|');
            barShowValue = sep < 0 || barEntry[0] != '0';
            barHtml = sep >= 0 ? barEntry[(sep + 1)..] : barEntry;
        }
        var iconShowValue = true;
        var iconHtml = "";
        if (hasIcon && iconEntry != null)
        {
            var sep = iconEntry.IndexOf('|');
            iconShowValue = sep < 0 || iconEntry[0] != '0';
            iconHtml = sep >= 0 ? iconEntry[(sep + 1)..] : iconEntry;
        }
        var showValue = barShowValue && iconShowValue;

        var sb = new StringBuilder();
        if (hasBar) sb.Append(barHtml);
        if (hasIcon) sb.Append($"<span style=\"position:absolute;left:4px;top:50%;transform:translateY(-50%);z-index:1\">{iconHtml}</span>");
        if (showValue)
            sb.Append($"<span style=\"position:relative;z-index:1\">{CellHtml(value)}</span>");
        return sb.ToString();
    }

    private static string CssSanitize(string value)
    {
        // Strip characters that could break CSS context
        return Regex.Replace(value, @"[;:{}()\\""']", "");
    }

}
