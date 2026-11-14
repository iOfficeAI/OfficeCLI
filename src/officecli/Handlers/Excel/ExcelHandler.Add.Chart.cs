// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using OfficeCli.Core;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Drawing = DocumentFormat.OpenXml.Drawing;
using SpreadsheetDrawing = DocumentFormat.OpenXml.Spreadsheet.Drawing;
using XDR = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using CX = DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;

namespace OfficeCli.Handlers;

// Per-element-type Add helpers for chart paths and the generic-XML default fallback. Mechanically extracted from the Add() god-method.
public partial class ExcelHandler
{
    private string AddChart(string parentPath, string type, InsertPosition? position, Dictionary<string, string> properties)
    {
        var index = position?.Index;
        var chartSegments = parentPath.TrimStart('/').Split('/', 2);
        var chartSheetName = chartSegments[0];
        var chartWorksheet = FindWorksheet(chartSheetName)
            ?? throw new ArgumentException($"Sheet not found: {chartSheetName}");

        // Parse chart data. Use TryGetValue(case-insensitive) so reads
        // are recorded by TrackingPropertyDictionary in handler-as-truth path.
        string chartType = "column";
        if (properties.TryGetValue("charttype", out var ctVal) || properties.TryGetValue("type", out ctVal))
            chartType = ctVal;
        var chartTitle = properties.GetValueOrDefault("title");

        // Support dataRange: read cell data from worksheet and build series with cell references
        string[]? categories;
        List<(string name, double[] values)> seriesData;
        string? dataRangeStr = null;
        if (properties.TryGetValue("datarange", out var dr) || properties.TryGetValue("range", out dr))
            dataRangeStr = dr;
        ChartRangeGeometry? dataRangeGeometry = null;
        if (!string.IsNullOrEmpty(dataRangeStr))
        {
            (seriesData, categories, dataRangeGeometry) = ParseDataRangeForChart(dataRangeStr, chartSheetName, properties);
        }
        else
        {
            // Qualify bare range references (A1:A3 → Sheet1!A1:A3) before the
            // shared parser sees them. An unqualified <c:f>$B$1:$B$3</c:f> is
            // unresolvable inside a chart part — real Excel ignores it and
            // falls back to ordinal category labels / the stale cache.
            foreach (var refKey in EnumerateChartRangeKeys(properties))
            {
                if (properties.TryGetValue(refKey, out var refVal)
                    && ChartHelper.IsRangeReference(refVal)
                    && !refVal.Contains('!'))
                    properties[refKey] = $"{Core.ModernFunctionQualifier.QuoteSheetNameForRef(chartSheetName)}!{refVal.Trim()}";
            }
            categories = ChartHelper.ParseCategories(properties);
            seriesData = ChartHelper.ParseSeriesData(properties);
            // CONSISTENCY(chart-series-rangeref-cache): when a series value or
            // categories arg is given as a cell RANGE (series1=B1:B4,
            // categories=A1:A4) rather than a literal list, ParseSeriesData/
            // ParseCategories leave the literal values empty (the range is only
            // emitted as a numRef/strRef formula). Real Excel — and the
            // `dataRange=` path here — also snapshot the referenced cells into
            // a numCache/strCache so the chart renders before the workbook is
            // re-evaluated (the HTML preview plots only from that cache).
            // Resolve the ranges against the worksheet now to backfill the
            // literal values, mirroring ParseDataRangeForChart.
            BackfillSeriesRangeValues(ref seriesData, ref categories, chartSheetName, properties);
        }

        if (seriesData.Count == 0)
        {
            // A supplied-but-consumed dataRange must not get the generic
            // "requires a data property" message: with a single-column range
            // and no explicit categories=, the sole column is reserved as
            // the category column, leaving zero series — say so.
            if (properties.ContainsKey("dataRange") || properties.ContainsKey("datarange"))
                throw new ArgumentException(
                    "dataRange resolved to 0 series columns: a single-column range is consumed as the " +
                    "category column by default. Pass categories= explicitly (e.g. categories=Sheet1!A1:A5) " +
                    "to plot that column as a series, or widen the dataRange to include a values column.");
            throw new ArgumentException("Chart requires a 'data' property. Use: data=\"Series1:1,2,3;Series2:4,5,6\" " +
                "or dataRange=\"Sheet1!A1:D5\" or series1=\"Revenue:100,200,300\"");
        }

        // Validate the chart type BEFORE any part is created: an unknown type
        // used to throw inside the builder AFTER the DrawingsPart and its
        // sheet relationship were attached, leaving an orphaned empty
        // <xdr:wsDr/> part behind on every failed attempt. Extended (cx)
        // types — funnel/treemap/… — route through ChartExBuilder below and
        // must not be run through the classic-type parser.
        if (!ChartExBuilder.IsExtendedChartType(chartType))
            ChartHelper.ParseChartType(chartType);

        // Create DrawingsPart if needed
        var drawingsPart = chartWorksheet.DrawingsPart
            ?? chartWorksheet.AddNewPart<DrawingsPart>();

        if (drawingsPart.WorksheetDrawing == null)
        {
            drawingsPart.WorksheetDrawing = new XDR.WorksheetDrawing();
            drawingsPart.WorksheetDrawing.Save();

            if (GetSheet(chartWorksheet).GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Drawing>() == null)
            {
                var drawingRelId = chartWorksheet.GetIdOfPart(drawingsPart);
                GetSheet(chartWorksheet).Append(
                    new DocumentFormat.OpenXml.Spreadsheet.Drawing { Id = drawingRelId });
                SaveWorksheet(chartWorksheet);
            }
        }

        // Position via TwoCellAnchor (shared by both standard and extended charts)
        // CONSISTENCY(ole-width-units): accept `anchor=D2:J18` as a cell
        // range (same grammar as OLE, shape, picture). When both
        // `anchor=<range>` and `x/y/width/height` are supplied, anchor
        // wins with a warning — matches shape/picture/OLE convention.
        int fromCol, fromRow, toCol, toRow;
        if (properties.TryGetValue("anchor", out var chartAnchorStr) && !string.IsNullOrWhiteSpace(chartAnchorStr))
        {
            // Non-short-circuit | on purpose: each ContainsKey marks the key
            // as handler-read in TrackingPropertyDictionary. With ||, finding
            // `width` skipped the x/y/height probes and they surfaced as a
            // false "UNSUPPORTED props" warning alongside this explicit one.
            if (properties.ContainsKey("width") | properties.ContainsKey("height")
                | properties.ContainsKey("x") | properties.ContainsKey("y"))
                Console.Error.WriteLine(
                    "Warning: 'x'/'y'/'width'/'height' are ignored when 'anchor' is provided (anchor defines the full rectangle).");
            if (!TryParseCellRangeAnchor(chartAnchorStr, out var cxFrom, out var cyFrom, out var cxTo, out var cyTo))
                throw new ArgumentException($"Invalid anchor: '{chartAnchorStr}'. Expected e.g. 'D2' or 'D2:J18'.");
            fromCol = cxFrom;
            fromRow = cyFrom;
            if (cxTo < 0) { cxTo = fromCol + 8; cyTo = fromRow + 15; }
            toCol = cxTo;
            toRow = cyTo;
        }
        else
        {
            // CONSISTENCY(ole-width-units): accept cm/in/pt/EMU on chart x/y/width/height
            // (matches schema doc + OLE/picture/shape Add). Plain ints stay cell-count.
            fromCol = properties.TryGetValue("x", out var xStr) ? ParseAnchorOrigin(xStr, "x") : 0;
            fromRow = properties.TryGetValue("y", out var yStr) ? ParseAnchorOrigin(yStr, "y") : 0;
            toCol = properties.TryGetValue("width", out var wStr) ? fromCol + ParseAnchorDimension(wStr, "width") : fromCol + 8;
            toRow = properties.TryGetValue("height", out var hStr) ? fromRow + ParseAnchorDimension(hStr, "height") : fromRow + 15;
        }

        // Extended chart types (cx:chart) — funnel, treemap, sunburst, boxWhisker, histogram
        if (ChartExBuilder.IsExtendedChartType(chartType))
        {
            // Excel chartEx pulls data directly from the host workbook via
            // cx:f references, not from an embedded xlsx. When the caller
            // provided inline categories+values (no dataRange), persist
            // them into the chart's host sheet at A1..B(N+1) so the cx:f
            // formulas resolve. Skip when dataRange is given — those cx:f
            // already point at user-owned cells.
            if (string.IsNullOrEmpty(dataRangeStr))
                WriteChartExInlineDataToSheet(chartWorksheet, categories, seriesData);

            var cxChartSpace = ChartExBuilder.BuildExtendedChartSpace(
                chartType, chartTitle, categories, seriesData, properties);
            // Excel chartEx references the host workbook directly via cx:f
            // formulas (no embedded xlsx sidecar). Strip the externalData
            // element the shared builder emits for PPT/Word, otherwise Excel
            // tries to resolve rId1 against this chart's rels and errors out.
            var extData = cxChartSpace.Descendants<CX.ExternalData>().FirstOrDefault();
            extData?.Remove();
            // Rewrite cx:f references to the actual host cells.
            // BuildExtendedChartSpace hardcodes the embedded-xlsx layout
            // (Sheet1!$A$2:… — correct for PPT/Word, whose chart data lives in
            // an embedded workbook whose sheet really is Sheet1). On an xlsx
            // host the formulas must reference the HOST workbook instead:
            //  - dataRange given: remap sheet name AND coordinates to the
            //    user's range (issue #176 — the old code skipped this case
            //    entirely, so a renamed sheet produced Sheet1! references
            //    that Excel resolves as an external-workbook link: "cannot
            //    update links" prompt + #REF! category axis + empty plot;
            //    a non-A1-anchored range additionally pointed at the wrong
            //    rows even on a sheet that happened to be named Sheet1).
            //  - inline data (no dataRange): cells were just written to the
            //    host sheet in the embedded A1-anchored layout, so only the
            //    sheet name needs replacing.
            if (dataRangeGeometry != null)
            {
                RemapChartExFormulasToHostRange(cxChartSpace, dataRangeGeometry);
            }
            else if (chartSheetName != "Sheet1")
            {
                var quoted = Core.ModernFunctionQualifier.QuoteSheetNameForRef(chartSheetName);
                foreach (var f in cxChartSpace.Descendants<CX.Formula>())
                {
                    if (f.Text.StartsWith("Sheet1!", StringComparison.Ordinal))
                        f.Text = quoted + f.Text.Substring("Sheet1".Length);
                }
            }
            var extChartPart = drawingsPart.AddNewPart<ExtendedChartPart>();
            extChartPart.ChartSpace = cxChartSpace;
            extChartPart.ChartSpace.Save();

            // CONSISTENCY(chartex-sidecars): every Office-canonical
            // chartEx part requires two sidecar parts linked via
            // relationships: a ChartStylePart (chs:chartStyle) and a
            // ChartColorStylePart (chs:colorStyle). Excel rejects
            // files that have the chartEx body but lack these
            // sidecars (silent "We found a problem" repair that
            // DELETES the entire drawing containing the chart —
            // slicers and all other anchors get collateral-damaged).
            // The SDK validator doesn't flag this because each part
            // is independently schema-valid; it's only the absence
            // of the sidecar relationship that Excel trips on.
            //
            // chartStyle is built by ChartExStyleBuilder; an
            // optional chartStyle=N prop on the caller picks a
            // numbered style variant, default = 0.
            var styleVariant = properties.GetValueOrDefault("chartStyle")
                            ?? properties.GetValueOrDefault("chartstyle")
                            ?? "default";
            var stylePart = extChartPart.AddNewPart<ChartStylePart>();
            using (var styleStream = ChartExStyleBuilder.BuildChartStyleXml(chartType, styleVariant))
                stylePart.FeedData(styleStream);
            var colorStylePart = extChartPart.AddNewPart<ChartColorStylePart>();
            using (var colorStream = LoadChartExResource("chartex-colors.xml"))
                colorStylePart.FeedData(colorStream);

            var cxRelId = drawingsPart.GetIdOfPart(extChartPart);
            var cxAnchor = new XDR.TwoCellAnchor();
            cxAnchor.Append(new XDR.FromMarker(
                new XDR.ColumnId(fromCol.ToString()),
                new XDR.ColumnOffset("0"),
                new XDR.RowId(fromRow.ToString()),
                new XDR.RowOffset("0")));
            cxAnchor.Append(new XDR.ToMarker(
                new XDR.ColumnId(toCol.ToString()),
                new XDR.ColumnOffset("0"),
                new XDR.RowId(toRow.ToString()),
                new XDR.RowOffset("0")));

            var cxGraphicFrame = new XDR.GraphicFrame();
            var cxExistingIds = drawingsPart.WorksheetDrawing.Descendants<XDR.NonVisualDrawingProperties>()
                .Select(p => (uint?)p.Id?.Value ?? 0u)
                .DefaultIfEmpty(1u)
                .Max();
            var cxFrameId = cxExistingIds + 1;
            cxGraphicFrame.NonVisualGraphicFrameProperties = new XDR.NonVisualGraphicFrameProperties(
                new XDR.NonVisualDrawingProperties
                {
                    Id = cxFrameId,
                    // CONSISTENCY(drawing-name): honor `name=` like
                    // sheet/namedrange/picture/shape. Fall back to
                    // chartTitle for back-compat, then "Chart".
                    Name = properties.GetValueOrDefault("name") ?? chartTitle ?? "Chart"
                },
                new XDR.NonVisualGraphicFrameDrawingProperties()
            );
            cxGraphicFrame.Transform = new XDR.Transform(
                new Drawing.Offset { X = 0, Y = 0 },
                new Drawing.Extents { Cx = 0, Cy = 0 }
            );

            var cxChartRef = new DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing.RelId { Id = cxRelId };
            cxGraphicFrame.Append(new Drawing.Graphic(
                new Drawing.GraphicData(cxChartRef)
                {
                    Uri = "http://schemas.microsoft.com/office/drawing/2014/chartex"
                }
            ));

            cxAnchor.Append(cxGraphicFrame);
            cxAnchor.Append(new XDR.ClientData());
            drawingsPart.WorksheetDrawing.Append(cxAnchor);
            drawingsPart.WorksheetDrawing.Save();

            // Count all charts (both regular and extended)
            var totalCharts = CountExcelCharts(drawingsPart);
            return $"/{chartSheetName}/chart[{totalCharts}]";
        }

        // Build chart content BEFORE adding part (invalid type throws, must not leave empty part)
        var chartSpace = ChartHelper.BuildChartSpace(chartType, chartTitle, categories, seriesData, properties);
        var chartPart = drawingsPart.AddNewPart<ChartPart>();
        chartPart.ChartSpace = chartSpace;
        chartPart.ChartSpace.Save();

        // Apply deferred properties (axisTitle, dataLabels, etc.) via SetChartProperties.
        // CONSISTENCY(tracking-deferred-filter): see PowerPointHandler.Add.Media.cs —
        // .Where() over TrackingPropertyDictionary marks every key consumed and
        // silently swallows real typos. Iterate Keys + TryGetValue per match instead.
        var deferredProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var dk in properties.Keys.ToList())
        {
            if (ChartHelper.IsDeferredKey(dk) && properties.TryGetValue(dk, out var dv))
                deferredProps[dk] = dv;
        }
        if (deferredProps.Count > 0)
        {
            // Atomicity: a bad deferred prop (e.g. holeSize=500) throws here,
            // AFTER the ChartPart was created. Roll the part back on failure so
            // a rejected Add leaves no orphaned/partial chart part behind — the
            // reported error (exit 1) must mean nothing was added.
            try
            {
                ChartHelper.SetChartProperties(chartPart, deferredProps);
            }
            catch
            {
                drawingsPart.DeletePart(chartPart);
                throw;
            }
        }

        var anchor = new XDR.TwoCellAnchor();
        anchor.Append(new XDR.FromMarker(
            new XDR.ColumnId(fromCol.ToString()),
            new XDR.ColumnOffset("0"),
            new XDR.RowId(fromRow.ToString()),
            new XDR.RowOffset("0")));
        anchor.Append(new XDR.ToMarker(
            new XDR.ColumnId(toCol.ToString()),
            new XDR.ColumnOffset("0"),
            new XDR.RowId(toRow.ToString()),
            new XDR.RowOffset("0")));

        var chartRelId = drawingsPart.GetIdOfPart(chartPart);
        var graphicFrame = new XDR.GraphicFrame();
        // Compute a unique cNvPr ID: use max existing ID + 1 to avoid duplicates after deletion
        var existingIds = drawingsPart.WorksheetDrawing.Descendants<XDR.NonVisualDrawingProperties>()
            .Select(p => (uint?)p.Id?.Value ?? 0u)
            .DefaultIfEmpty(1u)
            .Max();
        var chartFrameId = existingIds + 1;
        graphicFrame.NonVisualGraphicFrameProperties = new XDR.NonVisualGraphicFrameProperties(
            new XDR.NonVisualDrawingProperties
            {
                Id = chartFrameId,
                // CONSISTENCY(drawing-name): honor `name=` like
                // sheet/namedrange/picture/shape. Fall back to
                // chartTitle for back-compat, then "Chart".
                Name = properties.GetValueOrDefault("name") ?? chartTitle ?? "Chart"
            },
            new XDR.NonVisualGraphicFrameDrawingProperties()
        );
        graphicFrame.Transform = new XDR.Transform(
            new Drawing.Offset { X = 0, Y = 0 },
            new Drawing.Extents { Cx = 0, Cy = 0 }
        );

        var chartRef = new C.ChartReference { Id = chartRelId };
        graphicFrame.Append(new Drawing.Graphic(
            new Drawing.GraphicData(chartRef)
            {
                Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart"
            }
        ));

        anchor.Append(graphicFrame);
        anchor.Append(new XDR.ClientData());
        drawingsPart.WorksheetDrawing.Append(anchor);
        drawingsPart.WorksheetDrawing.Save();

        // Legend is already handled inside BuildChartSpace

        var chartIdx = CountExcelCharts(drawingsPart);
        return $"/{chartSheetName}/chart[{chartIdx}]";
    }

    // BUG-002: `add /SheetName/chart[N] --type chart-series` — append a data
    // series to an existing chart. Mirrors PowerPointHandler.AddChartSeries
    // (R22-1); additionally resolves xlsx cell-range values/categories into
    // numRef/strRef + cached snapshot, matching what chart Add emits for
    // range-referenced series (CONSISTENCY(chart-series-rangeref-cache)).
    private string AddChartSeries(string parentPath, Dictionary<string, string> properties)
    {
        var m = Regex.Match(parentPath, @"^/([^/]+)/chart\[(\d+)\]$");
        if (!m.Success)
            throw new ArgumentException(
                "series must be added to a chart parent: /SheetName/chart[N]");
        var sheetName = m.Groups[1].Value;
        var chartIdx = int.Parse(m.Groups[2].Value);
        var worksheet = FindWorksheet(sheetName)
            ?? throw new ArgumentException($"Sheet not found: {sheetName}");
        var drawingsPart = worksheet.DrawingsPart
            ?? throw new ArgumentException("Sheet has no drawings/charts");
        var excelCharts = GetExcelCharts(drawingsPart);
        if (chartIdx < 1 || chartIdx > excelCharts.Count)
            throw new ArgumentException($"Chart {chartIdx} not found (total: {excelCharts.Count})");
        var chartInfo = excelCharts[chartIdx - 1];
        if (chartInfo.StandardPart == null)
            throw new ArgumentException(
                $"Chart at {parentPath} is not a standard chart (extended cx charts do not support add series).");
        var chartPart = chartInfo.StandardPart;

        // Resolve range-reference values/categories against the workbook so
        // AddSeries seeds literal data (which becomes the cached snapshot).
        // Mutate `properties` in place instead of copying: enumerating a
        // TrackingPropertyDictionary into a fresh dict marks EVERY key as
        // consumed (see TrackingPropertyDictionary.GetEnumerator), which
        // silently swallows unsupported_property reporting for unknown keys.
        // TryGetValue / indexer / Remove are the tracked access routes.
        string? valuesRef = null, categoriesRef = null;
        List<string>? cachedCats = null;
        if (properties.TryGetValue("values", out var valRaw) && ChartHelper.IsRangeReference(valRaw))
        {
            valuesRef = ChartHelper.NormalizeRangeReference(valRaw, sheetName);
            var cells = ResolveRangeToCellValues(valRaw, sheetName);
            if (cells != null)
                properties["values"] = string.Join(",", cells.Select(v =>
                    double.TryParse(v, System.Globalization.CultureInfo.InvariantCulture, out var n) ? n : 0));
            else
                properties.Remove("values");
        }
        if (properties.TryGetValue("categories", out var catRaw))
        {
            properties.Remove("categories"); // ChartHelper.AddSeries doesn't consume it
            if (ChartHelper.IsRangeReference(catRaw))
            {
                categoriesRef = ChartHelper.NormalizeRangeReference(catRaw, sheetName);
                cachedCats = ResolveRangeToCellValues(catRaw, sheetName);
            }
        }

        var newIdx = ChartHelper.AddSeries(chartPart, properties);
        if (newIdx == 0)
            throw new ArgumentException(
                "Cannot add a series: the chart has no existing series to derive structure from. Recreate the chart with the desired series instead.");
        ChartHelper.ApplySeriesRangeRefs(chartPart, newIdx, valuesRef, categoriesRef, cachedCats);
        return $"/{sheetName}/chart[{chartIdx}]/series[{newIdx}]";
    }

    private string AddDefault(string parentPath, string type, InsertPosition? position, Dictionary<string, string> properties)
    {
        var index = position?.Index;
        // Generic fallback: create typed element via SDK schema validation
        // Parse parentPath: /<SheetName>/xmlPath...
        var fbSegments = parentPath.TrimStart('/').Split('/', 2);
        var fbSheetName = fbSegments[0];
        var fbWorksheet = FindWorksheet(fbSheetName);
        if (fbWorksheet == null)
            throw new ArgumentException($"Sheet not found: {fbSheetName}");

        OpenXmlElement fbParent = GetSheet(fbWorksheet);
        if (fbSegments.Length > 1 && !string.IsNullOrEmpty(fbSegments[1]))
        {
            var xmlSegments = GenericXmlQuery.ParsePathSegments(fbSegments[1]);
            fbParent = GenericXmlQuery.NavigateByPath(fbParent!, xmlSegments)
                ?? throw new ArgumentException($"Parent element not found: {parentPath}");
        }

        var created = GenericXmlQuery.TryCreateTypedElement(fbParent!, type, properties, index);
        if (created == null)
            throw new ArgumentException(
                $"Unknown element type '{type}' for {parentPath}. " +
                "Valid types: sheet, row, cell, shape, chart, ole (object, embed), autofilter, databar, colorscale, iconset, formulacf, comment, namedrange, table, picture, validation, pivottable. " +
                "Use 'officecli xlsx add' for details.");

        SaveWorksheet(fbWorksheet);

        var siblings = fbParent.ChildElements.Where(e => e.LocalName == created.LocalName).ToList();
        var createdIdx = PathIndex.FromArrayIndex(siblings.IndexOf(created));
        return $"{parentPath}/{created.LocalName}[{createdIdx}]";
    }

    // Write inline chartEx categories/values into the host sheet at A1..B(N+1).
    // cx:f formulas in BuildExtendedChartSpace assume:
    //   row 1     = headers (A1 empty, B1+ = series names)
    //   rows 2..  = data (col A = categories, col B+ = series values)
    private void WriteChartExInlineDataToSheet(
        WorksheetPart worksheetPart,
        string[]? categories,
        List<(string name, double[] values)> seriesData)
    {
        if (seriesData.Count == 0) return;
        var sheet = worksheetPart.Worksheet
            ?? throw new InvalidOperationException("WorksheetPart has no Worksheet element.");
        var sheetData = sheet.GetFirstChild<SheetData>() ?? sheet.AppendChild(new SheetData());

        // Header row: B1, C1, ... = series names
        for (int s = 0; s < seriesData.Count; s++)
        {
            var col = ColumnIndexToName(2 + s); // B=2, C=3, ...
            var cell = FindOrCreateCell(sheetData, $"{col}1");
            cell.DataType = CellValues.String;
            cell.CellValue = new CellValue(seriesData[s].name);
        }
        // Data rows: A = category, B/C/... = series values
        var rowCount = categories?.Length ?? seriesData.Max(s => s.values.Length);
        for (int r = 0; r < rowCount; r++)
        {
            if (categories != null && r < categories.Length)
            {
                var aCell = FindOrCreateCell(sheetData, $"A{r + 2}");
                aCell.DataType = CellValues.String;
                aCell.CellValue = new CellValue(categories[r]);
            }
            for (int s = 0; s < seriesData.Count; s++)
            {
                if (r >= seriesData[s].values.Length) continue;
                var col = ColumnIndexToName(2 + s);
                var vCell = FindOrCreateCell(sheetData, $"{col}{r + 2}");
                vCell.DataType = CellValues.Number;
                vCell.CellValue = new CellValue(
                    seriesData[s].values[r].ToString("G", System.Globalization.CultureInfo.InvariantCulture));
            }
        }
        SaveWorksheet(worksheetPart);
    }

    private static string ColumnIndexToName(int idx)
    {
        // 1-indexed: 1→A, 2→B, ..., 26→Z, 27→AA
        var sb = new System.Text.StringBuilder();
        while (idx > 0) { idx--; sb.Insert(0, (char)('A' + idx % 26)); idx /= 26; }
        return sb.ToString();
    }

    /// <summary>
    /// Remap the embedded-xlsx cx:f formulas ChartExBuilder emits
    /// (Sheet1!$A$2:$A$N cats / Sheet1!$B$2… values / Sheet1!$B$1 series name)
    /// onto the HOST worksheet cells named by dataRange — issue #176. Column
    /// translation: embedded col A = the dataRange's category column, embedded
    /// col B+i = the i-th series column. Rows translate to the dataRange's
    /// data window (header row excluded). A series-name formula is dropped
    /// (cached literal kept) when the range has no header row; a category
    /// formula follows the explicit `categories=` ref when one was given,
    /// and is dropped for an inline literal list.
    /// </summary>
    private static void RemapChartExFormulasToHostRange(
        CX.ChartSpace cxChartSpace, ChartRangeGeometry geo)
    {
        var sheet = Core.ModernFunctionQualifier.QuoteSheetNameForRef(geo.SheetName);
        var pattern = new System.Text.RegularExpressions.Regex(
            @"^Sheet1!\$([A-Z]+)\$(\d+)(?::\$([A-Z]+)\$(\d+))?$");
        foreach (var f in cxChartSpace.Descendants<CX.Formula>().ToList())
        {
            var m = pattern.Match(f.Text ?? "");
            if (!m.Success) continue;
            var embCol = ColumnNameToIndex(m.Groups[1].Value);   // 1-based: A=1
            bool isRange = m.Groups[3].Success;

            if (!isRange && m.Groups[2].Value == "1")
            {
                // Series-name header cell (embedded row 1, col B+i).
                if (!geo.HasHeaderRow) { f.Remove(); continue; }
                var col = ColumnIndexToName(geo.FirstSeriesColIdx + embCol - 2);
                f.Text = $"{sheet}!${col}${geo.HeaderRow}";
                continue;
            }
            if (embCol == 1)
            {
                // Category column (embedded col A).
                if (geo.HasExplicitCategories)
                {
                    if (geo.ExplicitCategoriesRef is string catRef)
                    {
                        var bang = catRef.LastIndexOf('!');
                        f.Text = bang > 0
                            ? Core.ModernFunctionQualifier.QuoteSheetNameForRef(catRef[..bang].Trim('\'')) + catRef[bang..]
                            : catRef;
                    }
                    else
                    {
                        f.Remove();   // inline literal list — cached labels only
                    }
                    continue;
                }
                var catCol = ColumnIndexToName(geo.CatColIdx);
                f.Text = $"{sheet}!${catCol}${geo.DataStartRow}:${catCol}${geo.EndRow}";
                continue;
            }
            // Series values (embedded col B+i).
            var valCol = ColumnIndexToName(geo.FirstSeriesColIdx + embCol - 2);
            f.Text = $"{sheet}!${valCol}${geo.DataStartRow}:${valCol}${geo.EndRow}";
        }
    }

    /// <summary>Every chart prop key whose value may be a cell-range
    /// reference that ends up inside a chart-part c:f formula (and therefore
    /// must carry a sheet qualifier).</summary>
    private static IEnumerable<string> EnumerateChartRangeKeys(Dictionary<string, string> properties)
    {
        yield return "categories";
        yield return "categoriesRef";
        for (int i = 1; i <= 40; i++)
        {
            yield return $"series{i}";
            yield return $"series{i}.values";
            yield return $"series{i}.valuesRef";
            yield return $"series{i}.categories";
            yield return $"series{i}.bubbleSize";
            yield return $"series{i}.bubbleSizeRef";
        }
    }
}
