// Copyright 2025 OfficeCli (officecli.ai)
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

        // Parse chart data
        var chartType = properties.FirstOrDefault(kv =>
            kv.Key.Equals("charttype", StringComparison.OrdinalIgnoreCase)
            || kv.Key.Equals("type", StringComparison.OrdinalIgnoreCase)).Value
            ?? "column";
        var chartTitle = properties.GetValueOrDefault("title");

        // Support dataRange: read cell data from worksheet and build series with cell references
        string[]? categories;
        List<(string name, double[] values)> seriesData;
        var dataRangeStr = properties.FirstOrDefault(kv =>
            kv.Key.Equals("datarange", StringComparison.OrdinalIgnoreCase)
            || kv.Key.Equals("dataRange", StringComparison.OrdinalIgnoreCase)
            || kv.Key.Equals("range", StringComparison.OrdinalIgnoreCase)).Value;
        if (!string.IsNullOrEmpty(dataRangeStr))
        {
            (seriesData, categories) = ParseDataRangeForChart(dataRangeStr, chartSheetName, properties);
        }
        else
        {
            categories = ChartHelper.ParseCategories(properties);
            seriesData = ChartHelper.ParseSeriesData(properties);
        }

        if (seriesData.Count == 0)
            throw new ArgumentException("Chart requires data. Use: data=\"Series1:1,2,3;Series2:4,5,6\" " +
                "or dataRange=\"Sheet1!A1:D5\" or series1=\"Revenue:100,200,300\"");

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
            if (properties.ContainsKey("width") || properties.ContainsKey("height")
                || properties.ContainsKey("x") || properties.ContainsKey("y"))
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
            fromCol = properties.TryGetValue("x", out var xStr) ? ParseHelpers.SafeParseInt(xStr, "x") : 0;
            fromRow = properties.TryGetValue("y", out var yStr) ? ParseHelpers.SafeParseInt(yStr, "y") : 0;
            toCol = properties.TryGetValue("width", out var wStr) ? fromCol + ParseHelpers.SafeParseInt(wStr, "width") : fromCol + 8;
            toRow = properties.TryGetValue("height", out var hStr) ? fromRow + ParseHelpers.SafeParseInt(hStr, "height") : fromRow + 15;
        }

        // Extended chart types (cx:chart) — funnel, treemap, sunburst, boxWhisker, histogram
        if (ChartExBuilder.IsExtendedChartType(chartType))
        {
            var cxChartSpace = ChartExBuilder.BuildExtendedChartSpace(
                chartType, chartTitle, categories, seriesData, properties);
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

        // Apply deferred properties (axisTitle, dataLabels, etc.) via SetChartProperties
        var deferredProps = properties
            .Where(kv => ChartHelper.IsDeferredKey(kv.Key))
            .ToDictionary(kv => kv.Key, kv => kv.Value);
        if (deferredProps.Count > 0)
            ChartHelper.SetChartProperties(chartPart, deferredProps);

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
        var createdIdx = siblings.IndexOf(created) + 1;
        return $"{parentPath}/{created.LocalName}[{createdIdx}]";
    }

}
