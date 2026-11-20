// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeCli.Core;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Drawing = DocumentFormat.OpenXml.Drawing;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using XDR = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeCli.Handlers;

public partial class ExcelHandler
{

    /// <summary>
    /// Apply `x` / `y` / `width` / `height` to the N-th chart's
    /// <see cref="XDR.TwoCellAnchor"/> in a drawings part. Accepts the same
    /// value grammar as OLE objects and chart Add: integer cell counts, or
    /// unit-qualified EMU strings ("6cm", "2in", "720pt", raw EMU).
    ///
    /// Returns any keys from the input dict that couldn't be applied (parse
    /// failures, missing anchor, ...). Keys present but successfully applied
    /// are NOT returned — the caller is expected to strip them before
    /// forwarding to the chart content setter.
    ///
    /// CONSISTENCY(chart-position-set): mirrors the PPTX
    /// PowerPointHandler.Set.cs chart path — same vocabulary, same units —
    /// so one prop grammar covers chart position across all three document
    /// types. The mutation mechanic differs because Excel charts are pinned
    /// to cells via TwoCellAnchor.
    /// </summary>
    // BUG-R11-04: read the N-th chart's TwoCellAnchor as a "B2:F7" cell range
    // for chart Get. Mirrors ApplyChartPositionSet's GraphicFrame lookup so the
    // index semantics match. Returns null if the chart has no TwoCellAnchor
    // (e.g. absolute-anchored), in which case the caller omits the field.
    private static string? GetChartAnchorRange(DrawingsPart drawingsPart, int chartIdx)
    {
        if (drawingsPart.WorksheetDrawing == null) return null;
        var chartFrames = drawingsPart.WorksheetDrawing
            .Descendants<XDR.GraphicFrame>()
            .Where(gf => gf.Descendants<C.ChartReference>().Any() || IsExtendedChartFrame(gf))
            .ToList();
        if (chartIdx < 1 || chartIdx > chartFrames.Count) return null;
        var gf = chartFrames[chartIdx - 1];
        if (gf.Parent is not XDR.TwoCellAnchor anchor) return null;
        var fromM = anchor.FromMarker;
        var toM = anchor.ToMarker;
        if (fromM == null || toM == null) return null;
        if (!int.TryParse(fromM.GetFirstChild<XDR.ColumnId>()?.Text ?? "0", out var fc)) return null;
        if (!int.TryParse(fromM.GetFirstChild<XDR.RowId>()?.Text ?? "0", out var fr)) return null;
        if (!int.TryParse(toM.GetFirstChild<XDR.ColumnId>()?.Text ?? "0", out var tc)) return null;
        if (!int.TryParse(toM.GetFirstChild<XDR.RowId>()?.Text ?? "0", out var tr)) return null;
        // XDR col/row are 0-based; IndexToColumnName expects 1-based.
        return $"{IndexToColumnName(fc + 1)}{fr + 1}:{IndexToColumnName(tc + 1)}{tr + 1}";
    }

    /// <summary>
    /// Read the N-th chart's TwoCellAnchor into FormatEmu strings for the
    /// caller's Format dict (x / y / width / height in cm). Mirrors the
    /// OLE/picture readback so add/set/get round-trip in the same vocabulary
    /// as the schema doc. CONSISTENCY(ole-width-units).
    /// </summary>
    private static void PopulateChartPositionFormat(
        DrawingsPart drawingsPart, int chartIdx, DocumentNode chartNode)
    {
        if (drawingsPart.WorksheetDrawing == null) return;
        var chartFrames = drawingsPart.WorksheetDrawing
            .Descendants<XDR.GraphicFrame>()
            .Where(gf => gf.Descendants<C.ChartReference>().Any() || IsExtendedChartFrame(gf))
            .ToList();
        if (chartIdx < 1 || chartIdx > chartFrames.Count) return;
        var gf = chartFrames[chartIdx - 1];
        // Externally-authored charts (e.g. openpyxl) hang off oneCellAnchor /
        // absoluteAnchor. Bailing out here dropped x/y/width/height from Get,
        // so dump emitted no position and replay parked the chart at the
        // default origin — covering the sheet data. Read the from+ext (or
        // pos+ext) rectangle instead; replay normalizes to a twoCell anchor
        // at the same rectangle, which preserves the rendered layout.
        if (gf.Parent is XDR.OneCellAnchor oneAnchor)
        {
            var oFrom = oneAnchor.FromMarker;
            var oExt = oneAnchor.GetFirstChild<XDR.Extent>();
            if (oFrom == null || oExt?.Cx?.HasValue != true || oExt.Cy?.HasValue != true) return;
            int.TryParse(oFrom.GetFirstChild<XDR.ColumnId>()?.Text ?? "0", out var oCol);
            int.TryParse(oFrom.GetFirstChild<XDR.RowId>()?.Text ?? "0", out var oRow);
            long.TryParse(oFrom.GetFirstChild<XDR.ColumnOffset>()?.Text ?? "0", out var oColOff);
            long.TryParse(oFrom.GetFirstChild<XDR.RowOffset>()?.Text ?? "0", out var oRowOff);
            chartNode.Format["x"] = OfficeCli.Core.EmuConverter.FormatEmu((long)oCol * EmuPerColApprox + oColOff);
            chartNode.Format["y"] = OfficeCli.Core.EmuConverter.FormatEmu((long)oRow * EmuPerRowApprox + oRowOff);
            chartNode.Format["width"] = OfficeCli.Core.EmuConverter.FormatEmu(oExt.Cx.Value);
            chartNode.Format["height"] = OfficeCli.Core.EmuConverter.FormatEmu(oExt.Cy.Value);
            return;
        }
        if (gf.Parent is XDR.AbsoluteAnchor absAnchor)
        {
            var aPos = absAnchor.GetFirstChild<XDR.Position>();
            var aExt = absAnchor.GetFirstChild<XDR.Extent>();
            if (aExt?.Cx?.HasValue != true || aExt.Cy?.HasValue != true) return;
            chartNode.Format["x"] = OfficeCli.Core.EmuConverter.FormatEmu(aPos?.X?.Value ?? 0);
            chartNode.Format["y"] = OfficeCli.Core.EmuConverter.FormatEmu(aPos?.Y?.Value ?? 0);
            chartNode.Format["width"] = OfficeCli.Core.EmuConverter.FormatEmu(aExt.Cx.Value);
            chartNode.Format["height"] = OfficeCli.Core.EmuConverter.FormatEmu(aExt.Cy.Value);
            return;
        }
        if (gf.Parent is not XDR.TwoCellAnchor anchor) return;
        var fromM = anchor.FromMarker;
        var toM = anchor.ToMarker;
        if (fromM == null || toM == null) return;
        int fromCol = 0, fromRow = 0, toCol = 0, toRow = 0;
        long fromColOff = 0, fromRowOff = 0, toColOff = 0, toRowOff = 0;
        int.TryParse(fromM.GetFirstChild<XDR.ColumnId>()?.Text ?? "0", out fromCol);
        int.TryParse(fromM.GetFirstChild<XDR.RowId>()?.Text ?? "0", out fromRow);
        int.TryParse(toM.GetFirstChild<XDR.ColumnId>()?.Text ?? "0", out toCol);
        int.TryParse(toM.GetFirstChild<XDR.RowId>()?.Text ?? "0", out toRow);
        long.TryParse(fromM.GetFirstChild<XDR.ColumnOffset>()?.Text ?? "0", out fromColOff);
        long.TryParse(fromM.GetFirstChild<XDR.RowOffset>()?.Text ?? "0", out fromRowOff);
        long.TryParse(toM.GetFirstChild<XDR.ColumnOffset>()?.Text ?? "0", out toColOff);
        long.TryParse(toM.GetFirstChild<XDR.RowOffset>()?.Text ?? "0", out toRowOff);

        long xEmu = (long)fromCol * EmuPerColApprox + fromColOff;
        long yEmu = (long)fromRow * EmuPerRowApprox + fromRowOff;
        long widthEmu = Math.Max(0, (long)(toCol - fromCol)) * EmuPerColApprox + (toColOff - fromColOff);
        long heightEmu = Math.Max(0, (long)(toRow - fromRow)) * EmuPerRowApprox + (toRowOff - fromRowOff);
        if (xEmu < 0) xEmu = 0;
        if (yEmu < 0) yEmu = 0;
        if (widthEmu < 0) widthEmu = 0;
        if (heightEmu < 0) heightEmu = 0;
        chartNode.Format["x"] = OfficeCli.Core.EmuConverter.FormatEmu(xEmu);
        chartNode.Format["y"] = OfficeCli.Core.EmuConverter.FormatEmu(yEmu);
        chartNode.Format["width"] = OfficeCli.Core.EmuConverter.FormatEmu(widthEmu);
        chartNode.Format["height"] = OfficeCli.Core.EmuConverter.FormatEmu(heightEmu);
    }

    private static List<string> ApplyChartPositionSet(
        DrawingsPart drawingsPart, int chartIdx, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();
        if (drawingsPart.WorksheetDrawing == null) return unsupported;

        // Find the N-th chart frame (same order as GetExcelCharts).
        var chartFrames = drawingsPart.WorksheetDrawing
            .Descendants<XDR.GraphicFrame>()
            .Where(gf => gf.Descendants<C.ChartReference>().Any() || IsExtendedChartFrame(gf))
            .ToList();
        if (chartIdx < 1 || chartIdx > chartFrames.Count) return unsupported;
        var gf = chartFrames[chartIdx - 1];
        var anchor = gf.Parent as XDR.TwoCellAnchor;
        if (anchor?.FromMarker == null || anchor.ToMarker == null)
        {
            foreach (var k in new[] { "x", "y", "width", "height" })
                if (properties.ContainsKey(k)) unsupported.Add(k);
            return unsupported;
        }

        var fromM = anchor.FromMarker;
        var toM = anchor.ToMarker;

        // ---- Full rectangle (anchor=D2:K20) → set both markers ----
        // Lets a chart be moved AND resized in one prop after create, mirroring
        // chart Add's `anchor=<range>`. Takes precedence over x/y/width/height
        // (anchor defines the whole rectangle), so those are dropped if combined.
        static void SetMarkerCell(XDR.MarkerType marker, int col, int row)
        {
            if (marker.GetFirstChild<XDR.ColumnId>() is { } ci) ci.Text = col.ToString();
            if (marker.GetFirstChild<XDR.RowId>() is { } ri) ri.Text = row.ToString();
            if (marker.GetFirstChild<XDR.ColumnOffset>() is { } co) co.Text = "0";
            if (marker.GetFirstChild<XDR.RowOffset>() is { } ro) ro.Text = "0";
        }
        if (properties.TryGetValue("anchor", out var anchorStr) && !string.IsNullOrWhiteSpace(anchorStr))
        {
            if (TryParseCellRangeAnchor(anchorStr, out var aFromCol, out var aFromRow, out var aToCol, out var aToRow))
            {
                SetMarkerCell(fromM, aFromCol, aFromRow);
                SetMarkerCell(toM, aToCol, aToRow);
                // anchor defines the whole rectangle — skip x/y/width/height
                // (mirrors chart Add, which ignores them when anchor is given).
                return unsupported;
            }
            unsupported.Add("anchor");
        }

        // ---- Position (x, y) → FromMarker cell indices ----
        // `x` = column index (0-based), `y` = row index (0-based). Integer
        // only — sub-cell offset is not supported here (matches chart Add).
        // CONSISTENCY(ole-width-units): accept cm/in/pt/EMU via ParseAnchorOrigin
        // (mirrors chart Add). Plain int stays cell-count.
        if (properties.TryGetValue("x", out var xStr))
        {
            int newFromCol = -1;
            try { newFromCol = ParseAnchorOrigin(xStr, "x"); }
            catch { /* fall through to unsupported */ }
            if (newFromCol >= 0)
            {
                var fromColChild = fromM.GetFirstChild<XDR.ColumnId>();
                var oldFromCol = int.TryParse(fromColChild?.Text ?? "0", out var ofc) ? ofc : 0;
                if (fromColChild != null) fromColChild.Text = newFromCol.ToString();
                // Shift ToMarker column by the same delta to preserve width.
                var toColChild = toM.GetFirstChild<XDR.ColumnId>();
                if (toColChild != null && int.TryParse(toColChild.Text ?? "0", out var oldToCol))
                    toColChild.Text = (oldToCol + (newFromCol - oldFromCol)).ToString();
                // Reset fromCol offset to 0 (align to cell boundary).
                var fromColOffChild = fromM.GetFirstChild<XDR.ColumnOffset>();
                if (fromColOffChild != null) fromColOffChild.Text = "0";
            }
            else unsupported.Add("x");
        }

        if (properties.TryGetValue("y", out var yStr))
        {
            int newFromRow = -1;
            try { newFromRow = ParseAnchorOrigin(yStr, "y"); }
            catch { /* fall through to unsupported */ }
            if (newFromRow >= 0)
            {
                var fromRowChild = fromM.GetFirstChild<XDR.RowId>();
                var oldFromRow = int.TryParse(fromRowChild?.Text ?? "0", out var ofr) ? ofr : 0;
                if (fromRowChild != null) fromRowChild.Text = newFromRow.ToString();
                var toRowChild = toM.GetFirstChild<XDR.RowId>();
                if (toRowChild != null && int.TryParse(toRowChild.Text ?? "0", out var oldToRow))
                    toRowChild.Text = (oldToRow + (newFromRow - oldFromRow)).ToString();
                var fromRowOffChild = fromM.GetFirstChild<XDR.RowOffset>();
                if (fromRowOffChild != null) fromRowOffChild.Text = "0";
            }
            else unsupported.Add("y");
        }

        // ---- Dimensions (width, height) → rebuild ToMarker from FromMarker ----
        // Reuses the OLE-object path's EMU math (EmuPerColApprox / EmuPerRowApprox
        // approximation, sub-cell offset preserves precision).
        if (properties.TryGetValue("width", out var wStr))
        {
            long emuTotal;
            try { emuTotal = ParseAnchorDimensionEmu(wStr, "width"); }
            catch { unsupported.Add("width"); emuTotal = -1; }
            if (emuTotal >= 0)
            {
                int.TryParse(fromM.GetFirstChild<XDR.ColumnId>()?.Text ?? "0", out var fromCol);
                long.TryParse(fromM.GetFirstChild<XDR.ColumnOffset>()?.Text ?? "0", out var fromColOff);
                long wholeCols = emuTotal / EmuPerColApprox;
                long remCols = emuTotal % EmuPerColApprox;
                var toColChild = toM.GetFirstChild<XDR.ColumnId>();
                if (toColChild != null) toColChild.Text = (fromCol + (int)wholeCols).ToString();
                var toColOffChild = toM.GetFirstChild<XDR.ColumnOffset>();
                if (toColOffChild != null) toColOffChild.Text = (fromColOff + remCols).ToString();
            }
        }

        if (properties.TryGetValue("height", out var hStr))
        {
            long emuTotal;
            try { emuTotal = ParseAnchorDimensionEmu(hStr, "height"); }
            catch { unsupported.Add("height"); emuTotal = -1; }
            if (emuTotal >= 0)
            {
                int.TryParse(fromM.GetFirstChild<XDR.RowId>()?.Text ?? "0", out var fromRow);
                long.TryParse(fromM.GetFirstChild<XDR.RowOffset>()?.Text ?? "0", out var fromRowOff);
                long wholeRows = emuTotal / EmuPerRowApprox;
                long remRows = emuTotal % EmuPerRowApprox;
                var toRowChild = toM.GetFirstChild<XDR.RowId>();
                if (toRowChild != null) toRowChild.Text = (fromRow + (int)wholeRows).ToString();
                var toRowOffChild = toM.GetFirstChild<XDR.RowOffset>();
                if (toRowOffChild != null) toRowOffChild.Text = (fromRowOff + remRows).ToString();
            }
        }

        drawingsPart.WorksheetDrawing.Save();
        return unsupported;
    }

    // ==================== Extended Chart Helpers ====================

    private const string ExcelChartExUri = "http://schemas.microsoft.com/office/drawing/2014/chartex";

    /// <summary>
    /// Load a chartEx sidecar resource (style / colors XML) bundled as an
    /// embedded resource. Files are copied verbatim from an Excel reference
    /// treemap and reused for every chartEx type — they carry default
    /// style/palette content that has no dependency on chart layout or data.
    /// See the chartex-sidecars CONSISTENCY note in ExcelHandler.Add.cs for
    /// why these sidecars are load-bearing (Excel deletes the whole drawing
    /// if they are missing from the relationships).
    /// </summary>
    private static Stream LoadChartExResource(string fileName)
    {
        var assembly = typeof(ExcelHandler).Assembly;
        var resourceName = $"OfficeCli.Resources.{fileName}";
        var stream = assembly.GetManifestResourceStream(resourceName)
            ?? throw new InvalidOperationException(
                $"Embedded resource not found: {resourceName}. Ensure it is declared in officecli.csproj.");
        return stream;
    }

    /// <summary>
    /// Check if an XDR.GraphicFrame contains an extended chart (cx:chart).
    /// </summary>
    private static bool IsExtendedChartFrame(XDR.GraphicFrame gf)
    {
        return gf.Descendants<Drawing.GraphicData>()
            .Any(gd => gd.Uri == ExcelChartExUri);
    }

    /// <summary>
    /// Get the relationship ID from an extended chart GraphicFrame.
    /// </summary>
    private static string? GetExtendedChartRelId(XDR.GraphicFrame gf)
    {
        var gd = gf.Descendants<Drawing.GraphicData>().FirstOrDefault(g => g.Uri == ExcelChartExUri);
        if (gd == null) return null;
        var typed = gd.Descendants<DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing.RelId>().FirstOrDefault();
        if (typed?.Id?.Value != null) return typed.Id.Value;
        foreach (var child in gd.ChildElements)
        {
            var rId = child.GetAttributes().FirstOrDefault(a =>
                a.LocalName == "id" && a.NamespaceUri == "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            if (rId.Value != null) return rId.Value;
        }
        return null;
    }

    /// <summary>
    /// Count all charts (both standard ChartPart and ExtendedChartPart) in a DrawingsPart.
    /// </summary>
    private static int CountExcelCharts(DrawingsPart drawingsPart)
    {
        if (drawingsPart.WorksheetDrawing == null) return 0;
        return drawingsPart.WorksheetDrawing.Descendants<XDR.GraphicFrame>()
            .Count(gf => gf.Descendants<C.ChartReference>().Any() || IsExtendedChartFrame(gf));
    }

    /// <summary>
    /// Represents a chart in Excel that could be either a standard ChartPart or an ExtendedChartPart.
    /// </summary>
    private class ExcelChartInfo
    {
        public ChartPart? StandardPart { get; set; }
        public ExtendedChartPart? ExtendedPart { get; set; }
        public bool IsExtended => ExtendedPart != null;
    }

    /// <summary>
    /// Get all chart parts (standard + extended) in document order by walking GraphicFrame elements.
    /// </summary>
    private static List<ExcelChartInfo> GetExcelCharts(DrawingsPart drawingsPart)
    {
        var result = new List<ExcelChartInfo>();
        if (drawingsPart.WorksheetDrawing == null) return result;

        foreach (var gf in drawingsPart.WorksheetDrawing.Descendants<XDR.GraphicFrame>())
        {
            var chartRef = gf.Descendants<C.ChartReference>().FirstOrDefault();
            if (chartRef?.Id?.Value != null)
            {
                try
                {
                    var chartPart = (ChartPart)drawingsPart.GetPartById(chartRef.Id.Value);
                    result.Add(new ExcelChartInfo { StandardPart = chartPart });
                }
                catch { /* skip invalid references */ }
            }
            else if (IsExtendedChartFrame(gf))
            {
                var relId = GetExtendedChartRelId(gf);
                if (relId == null) continue;
                try
                {
                    var extPart = (ExtendedChartPart)drawingsPart.GetPartById(relId);
                    result.Add(new ExcelChartInfo { ExtendedPart = extPart });
                }
                catch { /* skip invalid references */ }
            }
        }

        return result;
    }

    /// <summary>
    /// Parse a dataRange (e.g. "Sheet1!A1:D5" or "A1:B3") and read cell data from the worksheet.
    /// Returns series data and populates properties with cell references for chart building.
    /// First row = category labels + series names, remaining rows = data.
    /// </summary>
    /// <summary>
    /// Host-range geometry of a parsed dataRange, for consumers that must emit
    /// their own cell formulas against the HOST workbook (the chartEx path —
    /// issue #176: ChartExBuilder hardcodes the embedded-xlsx Sheet1!$A$2 layout,
    /// which is correct for PPT/Word but must be remapped to these coordinates
    /// on an xlsx host). Column/row values are 1-based.
    /// </summary>
    internal sealed record ChartRangeGeometry(
        string SheetName,
        int CatColIdx,
        int FirstSeriesColIdx,
        int HeaderRow,
        int DataStartRow,
        int EndRow,
        bool HasHeaderRow,
        bool HasExplicitCategories,
        string? ExplicitCategoriesRef);

    private (List<(string name, double[] values)> seriesData, string[]? categories, ChartRangeGeometry geometry) ParseDataRangeForChart(
        string dataRange, string defaultSheetName, Dictionary<string, string> properties)
    {
        // CONSISTENCY(defined-name-range): if dataRange has no '!' and no ':' and
        // looks like a workbook-defined name, resolve it to its referent range
        // (e.g. "MyData" -> "Sheet1!$A$1:$B$3"). Excel charts accept defined-name
        // references as a data source, so do the same here.
        var trimmedInput = dataRange.Trim();
        if (!trimmedInput.Contains('!') && !trimmedInput.Contains(':') &&
            System.Text.RegularExpressions.Regex.IsMatch(trimmedInput, @"^[A-Za-z_][A-Za-z0-9_\.]*$"))
        {
            var workbook = _doc.WorkbookPart?.Workbook;
            var defNames = workbook?.GetFirstChild<DefinedNames>();
            var match = defNames?.Elements<DefinedName>()
                .FirstOrDefault(dn => string.Equals(dn.Name?.Value, trimmedInput, StringComparison.OrdinalIgnoreCase));
            if (match == null || string.IsNullOrEmpty(match.Text))
                throw new ArgumentException($"DefinedName '{trimmedInput}' not found");
            dataRange = match.Text!;
        }

        // Parse sheet name and range
        string rangeSheetName = defaultSheetName;
        string rangePart = dataRange.Trim();
        var bangIdx = rangePart.LastIndexOf('!');
        if (bangIdx >= 0)
        {
            rangeSheetName = UnquoteSheetName(rangePart[..bangIdx]);
            rangePart = rangePart[(bangIdx + 1)..];
        }

        // Strip any $ signs for parsing
        var cleanRange = rangePart.Replace("$", "");
        var rangeParts = cleanRange.Split(':');
        if (rangeParts.Length != 2)
            throw new ArgumentException($"Invalid dataRange: '{dataRange}'. Expected format: 'Sheet1!A1:D5', 'A1:B3', or a defined-name");

        var (startCol, startRow) = ParseCellReference(rangeParts[0]);
        var (endCol, endRow) = ParseCellReference(rangeParts[1]);
        var startColIdx = ColumnNameToIndex(startCol);
        var endColIdx = ColumnNameToIndex(endCol);

        // Find the worksheet and read cells
        var ws = FindWorksheet(rangeSheetName)
            ?? throw new ArgumentException($"Sheet not found: {rangeSheetName}");
        var sheetData = GetSheet(ws).GetFirstChild<SheetData>();
        if (sheetData == null)
            throw new ArgumentException($"Sheet '{rangeSheetName}' has no data");

        // Build cell lookup. Track value, the originating Cell (for DataType),
        // and a "is blank" flag for cells that exist but carry no value.
        // R20-03: blank-vs-zero distinction is needed for dispBlanksAs=gap.
        // R20-04: DataType drives header detection — only string-typed
        // first-row cells are treated as series names.
        var cellLookup = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var cellTypeLookup = new Dictionary<string, Cell>(StringComparer.OrdinalIgnoreCase);
        var cellPresent = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var row in sheetData.Elements<Row>())
        {
            var rowIdx = (int)(row.RowIndex?.Value ?? 0);
            if (rowIdx < startRow || rowIdx > endRow) continue;
            foreach (var cell in row.Elements<Cell>())
            {
                if (cell.CellReference?.Value != null)
                {
                    cellLookup[cell.CellReference.Value] = GetCellDisplayValue(cell);
                    cellTypeLookup[cell.CellReference.Value] = cell;
                    cellPresent.Add(cell.CellReference.Value);
                }
            }
        }

        // R20-04: a first-row cell counts as a header only when its DataType
        // is string-like (SharedString / InlineString / String). Numeric or
        // missing first-row cells mean "no header" — series starts at row 1.
        bool IsStringTypedHeader(string cellRef)
        {
            if (!cellTypeLookup.TryGetValue(cellRef, out var c)) return false;
            var dt = c.DataType?.Value;
            return dt == CellValues.SharedString
                || dt == CellValues.InlineString
                || dt == CellValues.String;
        }

        // Decide globally: if ANY non-corner cell in the first row is string-typed,
        // treat row 1 as headers; otherwise treat all rows as data and synthesize
        // series names. Picking globally keeps a single header convention
        // across columns (mixed string/number headers would be ambiguous).
        // NOTE: the authoritative header/dataStartRow decision is (re)computed
        // below once firstSeriesCol is known (it depends on whether an explicit
        // categories arg shifts the series window). These two locals are seeded
        // here only so the legacy variable names exist for the explicit-cat block.
        bool hasHeaderRow = false;
        int dataStartRow = startRow;

        // CONSISTENCY(chart-explicit-categories): when the caller supplies an
        // explicit `categories=` arg alongside dataRange, the dataRange's first
        // column is NO LONGER hijacked as category labels — it becomes a data
        // series like every other column. Categories come solely from the
        // explicit arg (a cell range read from the sheet, or an inline literal
        // list). Without an explicit arg, the legacy first-column-as-categories
        // behavior is preserved verbatim.
        bool hasExplicitCategories =
            (properties.TryGetValue("categories", out var explicitCatVal)
             && !string.IsNullOrWhiteSpace(explicitCatVal));

        // Series start one column later only when the first column is being
        // consumed as categories (i.e. no explicit categories arg).
        int firstSeriesCol = hasExplicitCategories ? startColIdx : startColIdx + 1;

        // Header detection must scan the same column window the series-building
        // loop uses, otherwise a 2-column dataRange with explicit categories
        // (firstSeriesCol == startColIdx) would never see its first header.
        hasHeaderRow = false;
        for (int c = firstSeriesCol; c <= endColIdx; c++)
        {
            var headerRef = $"{IndexToColumnName(c)}{startRow}";
            if (IsStringTypedHeader(headerRef)) { hasHeaderRow = true; break; }
        }
        dataStartRow = hasHeaderRow ? startRow + 1 : startRow;

        // Resolve category labels + (optional) categoriesRef.
        var categories = new List<string>();
        string? explicitCategoriesRef = null;   // non-null only for explicit RANGE form
        if (hasExplicitCategories)
        {
            var (explicitCats, explicitRef) =
                ResolveExplicitCategories(explicitCatVal!, rangeSheetName, properties);
            categories.AddRange(explicitCats);
            explicitCategoriesRef = explicitRef;
        }
        else
        {
            // Legacy path: first column (excluding header row if present) = category labels.
            for (int r = dataStartRow; r <= endRow; r++)
            {
                var cellRef = $"{startCol}{r}";
                cellLookup.TryGetValue(cellRef, out var catVal);
                categories.Add(catVal ?? "");
            }
        }

        var seriesData = new List<(string name, double[] values)>();
        int seriesIdx = 1;
        for (int c = firstSeriesCol; c <= endColIdx; c++)
        {
            var colName = IndexToColumnName(c);
            string seriesName;
            if (hasHeaderRow)
            {
                var headerRef = $"{colName}{startRow}";
                cellLookup.TryGetValue(headerRef, out var sn);
                seriesName = sn ?? $"Series {seriesIdx}";
            }
            else
            {
                seriesName = $"Series {seriesIdx}";
            }

            // Series values + per-index blank tracking. R20-03: under
            // dispBlanksAs=gap, blank source cells must be omitted from the
            // numCache; we forward the blank-index list via properties so
            // ApplySeriesReferences/numCache builder can honor it.
            var values = new List<double>();
            var blankIndexes = new List<int>();
            int idx = 0;
            for (int r = dataStartRow; r <= endRow; r++)
            {
                var cellRef = $"{colName}{r}";
                bool isBlank = !cellPresent.Contains(cellRef)
                    || string.IsNullOrEmpty(cellLookup.GetValueOrDefault(cellRef));
                cellLookup.TryGetValue(cellRef, out var valStr);
                if (double.TryParse(valStr, System.Globalization.CultureInfo.InvariantCulture, out var num))
                    values.Add(num);
                else
                    values.Add(0);
                if (isBlank) blankIndexes.Add(idx);
                idx++;
            }

            // Set up cell references in properties for ApplySeriesReferences
            var valuesRef = $"{rangeSheetName}!${colName}${dataStartRow}:${colName}${endRow}";
            properties[$"series{seriesIdx}.name"] = seriesName;
            properties[$"series{seriesIdx}.values"] = valuesRef;
            // categoriesRef:
            //  - legacy (no explicit arg): first dataRange column.
            //  - explicit RANGE arg: the resolved explicit ref.
            //  - explicit INLINE literal: no ref (mirror the no-dataRange inline
            //    path which never sets series{i}.categories).
            if (hasExplicitCategories)
            {
                if (explicitCategoriesRef != null)
                    properties[$"series{seriesIdx}.categories"] = explicitCategoriesRef;
            }
            else
            {
                var categoriesRef = $"{rangeSheetName}!${startCol}${dataStartRow}:${startCol}${endRow}";
                properties[$"series{seriesIdx}.categories"] = categoriesRef;
            }
            if (blankIndexes.Count > 0)
                properties[$"series{seriesIdx}._blankIndexes"] = string.Join(",", blankIndexes);

            seriesData.Add((seriesName, values.ToArray()));
            seriesIdx++;
        }

        var geometry = new ChartRangeGeometry(
            rangeSheetName, startColIdx, firstSeriesCol, startRow, dataStartRow,
            endRow, hasHeaderRow, hasExplicitCategories, explicitCategoriesRef);
        return (seriesData, categories.ToArray(), geometry);
    }

    /// <summary>
    /// Resolve an explicit `categories=` value supplied alongside a dataRange.
    /// Two forms:
    ///   - cell RANGE ("Sheet1!A2:A3" / "A2:A3", contains ':'): read the cells
    ///     directly from the target worksheet to produce label strings, and
    ///     compute a categoriesRef in the same `Sheet!$A$2:$A$3` form the
    ///     dataRange path emits for valuesRef/categoriesRef.
    ///   - inline literal list ("North,EMEA,APAC"): parse via
    ///     ChartHelper.ParseCategories; no ref (returns null) so the caller
    ///     does not set a bogus series{i}.categories — mirroring the
    ///     no-dataRange inline path.
    /// The range cells may span rows outside the dataRange's loaded window, so
    /// read them straight from the worksheet rather than the dataRange cellLookup.
    /// </summary>
    private (string[] categories, string? categoriesRef) ResolveExplicitCategories(
        string explicitValue, string defaultSheetName, Dictionary<string, string> properties)
    {
        var trimmed = explicitValue.Trim();

        // Inline literal list (no ':') → no ref.
        if (!trimmed.Contains(':'))
        {
            var inline = ChartHelper.ParseCategories(properties) ?? Array.Empty<string>();
            return (inline, null);
        }

        // Cell range form. Strip a leading 'Sheet'! prefix (same approach as dataRange).
        string catSheetName = defaultSheetName;
        string rangePart = trimmed;
        var bangIdx = rangePart.LastIndexOf('!');
        if (bangIdx >= 0)
        {
            catSheetName = UnquoteSheetName(rangePart[..bangIdx]);
            rangePart = rangePart[(bangIdx + 1)..];
        }

        var cleanRange = rangePart.Replace("$", "");
        var rangeParts = cleanRange.Split(':');
        if (rangeParts.Length != 2)
            throw new ArgumentException(
                $"Invalid categories range: '{explicitValue}'. Expected format: 'Sheet1!A2:A3' or 'A2:A3'.");

        var (startCol, startRow) = ParseCellReference(rangeParts[0]);
        var (endCol, endRow) = ParseCellReference(rangeParts[1]);
        var startColIdx = ColumnNameToIndex(startCol);
        var endColIdx = ColumnNameToIndex(endCol);

        var ws = FindWorksheet(catSheetName)
            ?? throw new ArgumentException($"Sheet not found: {catSheetName}");
        var sheetData = GetSheet(ws).GetFirstChild<SheetData>()
            ?? throw new ArgumentException($"Sheet '{catSheetName}' has no data");

        // Read the explicit-range cells directly (rows may differ from dataRange).
        var lookup = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var row in sheetData.Elements<Row>())
        {
            var rowIdx = (int)(row.RowIndex?.Value ?? 0);
            if (rowIdx < startRow || rowIdx > endRow) continue;
            foreach (var cell in row.Elements<Cell>())
                if (cell.CellReference?.Value != null)
                    lookup[cell.CellReference.Value] = GetCellDisplayValue(cell);
        }

        var labels = new List<string>();
        for (int r = startRow; r <= endRow; r++)
            for (int c = startColIdx; c <= endColIdx; c++)
            {
                var cellRef = $"{IndexToColumnName(c)}{r}";
                lookup.TryGetValue(cellRef, out var v);
                labels.Add(v ?? "");
            }

        var categoriesRef =
            $"{catSheetName}!${IndexToColumnName(startColIdx)}${startRow}:${IndexToColumnName(endColIdx)}${endRow}";
        return (labels.ToArray(), categoriesRef);
    }

    /// <summary>
    /// CONSISTENCY(chart-series-rangeref-cache): backfill literal series values
    /// and category labels from cell-RANGE references (series{N}=B1:B4,
    /// series{N}.values=Sheet1!B1:B4, categories=A1:A4) so the chart builder
    /// seeds a numCache/strCache — exactly like the dataRange= path. The range
    /// is still emitted as a numRef/strRef formula by ApplySeriesReferences;
    /// this only fills the cached snapshot the HTML preview (and real Excel,
    /// pre-recalc) renders from. Cells that are empty/non-numeric resolve to 0
    /// for values and "" for labels, matching ParseDataRangeForChart.
    /// </summary>
    private void BackfillSeriesRangeValues(
        ref List<(string name, double[] values)> seriesData,
        ref string[]? categories,
        string defaultSheetName,
        Dictionary<string, string> properties)
    {
        for (int i = 1; i <= seriesData.Count; i++)
        {
            // Already has literal values (e.g. series1=10,20,30) — nothing to do.
            if (seriesData[i - 1].values.Length > 0) continue;

            string? rangeRef = null;
            if (properties.TryGetValue($"series{i}.values", out var dotVal)
                && ChartHelper.IsRangeReference(dotVal))
                rangeRef = dotVal;
            else if (properties.TryGetValue($"series{i}", out var legacyVal)
                && ChartHelper.IsRangeReference(legacyVal))
                rangeRef = legacyVal;
            if (rangeRef == null) continue;

            var cells = ResolveRangeToCellValues(rangeRef, defaultSheetName);
            if (cells == null) continue;
            var values = cells.Select(v =>
                double.TryParse(v, System.Globalization.CultureInfo.InvariantCulture, out var n) ? n : 0)
                .ToArray();
            seriesData[i - 1] = (seriesData[i - 1].name, values);
        }

        // categories=<range>: ParseCategories returns null for a range, so the
        // builder seeds no strLit (=> empty strCache). Resolve the labels here.
        // Also probe categoriesRef= and the per-series dotted form
        // (series1.categories=) — the batch emitter replays charts through
        // those keys, and without this the strCache degraded to ordinal
        // placeholders ("1","2") on every dump→replay: axis labels silently
        // lost their cell text.
        if (categories == null || categories.Length == 0)
        {
            string? catRange = null;
            if (properties.TryGetValue("categories", out var catStr)
                && ChartHelper.IsRangeReference(catStr))
                catRange = catStr;
            else if (properties.TryGetValue("categoriesRef", out var catRefStr)
                && ChartHelper.IsRangeReference(catRefStr))
                catRange = catRefStr;
            else if (properties.TryGetValue("series1.categories", out var s1Cat)
                && ChartHelper.IsRangeReference(s1Cat))
                catRange = s1Cat;
            if (catRange != null)
            {
                var labels = ResolveRangeToCellValues(catRange, defaultSheetName);
                if (labels != null && labels.Count > 0)
                    categories = labels.ToArray();
            }
        }
    }

    /// <summary>
    /// Reads a single-column (or row) cell range reference into an ordered list
    /// of display-value strings. Returns null if the reference can't be parsed
    /// or the sheet is missing/empty. Missing cells in the range yield "".
    /// </summary>
    private List<string>? ResolveRangeToCellValues(string rangeRef, string defaultSheetName)
    {
        string sheetName = defaultSheetName;
        string rangePart = rangeRef.Trim();
        var bangIdx = rangePart.LastIndexOf('!');
        if (bangIdx >= 0)
        {
            sheetName = UnquoteSheetName(rangePart[..bangIdx]);
            rangePart = rangePart[(bangIdx + 1)..];
        }

        var cleanRange = rangePart.Replace("$", "");
        var rangeParts = cleanRange.Split(':');
        if (rangeParts.Length != 2) return null;

        var (startCol, startRow) = ParseCellReference(rangeParts[0]);
        var (endCol, endRow) = ParseCellReference(rangeParts[1]);
        var startColIdx = ColumnNameToIndex(startCol);
        var endColIdx = ColumnNameToIndex(endCol);

        var ws = FindWorksheet(sheetName);
        if (ws == null) return null;
        var sheetData = GetSheet(ws).GetFirstChild<SheetData>();
        if (sheetData == null) return null;

        var lookup = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var row in sheetData.Elements<Row>())
        {
            var rowIdx = (int)(row.RowIndex?.Value ?? 0);
            if (rowIdx < startRow || rowIdx > endRow) continue;
            foreach (var cell in row.Elements<Cell>())
                if (cell.CellReference?.Value != null)
                    lookup[cell.CellReference.Value] = GetCellDisplayValue(cell);
        }

        var result = new List<string>();
        for (int r = startRow; r <= endRow; r++)
            for (int c = startColIdx; c <= endColIdx; c++)
            {
                lookup.TryGetValue($"{IndexToColumnName(c)}{r}", out var v);
                result.Add(v ?? "");
            }
        return result;
    }
}
