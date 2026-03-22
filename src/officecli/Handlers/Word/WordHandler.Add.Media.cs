// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    private string AddChart(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        var chartMainPart = _doc.MainDocumentPart!;

        // Parse chart data
        var chartType = properties.FirstOrDefault(kv =>
            kv.Key.Equals("charttype", StringComparison.OrdinalIgnoreCase)
            || kv.Key.Equals("type", StringComparison.OrdinalIgnoreCase)).Value
            ?? "column";
        var chartTitle = properties.GetValueOrDefault("title");
        var categories = Core.ChartHelper.ParseCategories(properties);
        var seriesData = Core.ChartHelper.ParseSeriesData(properties);

        if (seriesData.Count == 0)
            throw new ArgumentException("Chart requires data. Use: data=\"Series1:1,2,3;Series2:4,5,6\" " +
                "or series1=\"Revenue:100,200,300\"");

        // Create ChartPart and build chart
        var chartPart = chartMainPart.AddNewPart<ChartPart>();
        chartPart.ChartSpace = Core.ChartHelper.BuildChartSpace(chartType, chartTitle, categories, seriesData, properties);

        // Apply deferred properties (axisTitle, dataLabels, etc.) via SetChartProperties
        // Must be called BEFORE Save() so the in-memory DOM is still available
        var deferredProps = properties
            .Where(kv => Core.ChartHelper.DeferredAddKeys.Contains(kv.Key))
            .ToDictionary(kv => kv.Key, kv => kv.Value);
        if (deferredProps.Count > 0)
            Core.ChartHelper.SetChartProperties(chartPart, deferredProps);
        else
            chartPart.ChartSpace.Save();

        var chartRelId = chartMainPart.GetIdOfPart(chartPart);

        // Dimensions (default: 15cm x 10cm)
        long chartCx = properties.TryGetValue("width", out var chartWStr) ? ParseEmu(chartWStr) : 5400000;
        long chartCy = properties.TryGetValue("height", out var chStr) ? ParseEmu(chStr) : 3600000;

        var docPropId = NextImageId();
        var chartName = chartTitle ?? $"Chart {docPropId}";

        // Build Drawing/Inline with ChartReference
        var inline = new DW.Inline(
            new DW.Extent { Cx = chartCx, Cy = chartCy },
            new DW.EffectExtent { LeftEdge = 0, TopEdge = 0, RightEdge = 0, BottomEdge = 0 },
            new DW.DocProperties { Id = docPropId, Name = chartName },
            new DW.NonVisualGraphicFrameDrawingProperties(),
            new A.Graphic(
                new A.GraphicData(
                    new DocumentFormat.OpenXml.Drawing.Charts.ChartReference { Id = chartRelId }
                )
                { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" }
            )
        )
        {
            DistanceFromTop = 0U,
            DistanceFromBottom = 0U,
            DistanceFromLeft = 0U,
            DistanceFromRight = 0U
        };

        var chartRun = new Run(new Drawing(inline));
        Paragraph chartPara;
        if (parent is Paragraph existingChartPara)
        {
            existingChartPara.AppendChild(chartRun);
            chartPara = existingChartPara;
        }
        else
        {
            chartPara = new Paragraph(chartRun);
            AppendToParent(parent, chartPara);
        }

        var chartIdx = chartMainPart.ChartParts.ToList().IndexOf(chartPart) + 1;
        return $"/chart[{chartIdx}]";
    }

    private string AddPicture(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        if (!properties.TryGetValue("path", out var imgPath) && !properties.TryGetValue("src", out imgPath))
            throw new ArgumentException("'path' (or 'src') property is required for picture type");

        var (imgStream, imgPartType) = OfficeCli.Core.ImageSource.Resolve(imgPath);
        using var imgStreamDispose = imgStream;

        var mainPart = _doc.MainDocumentPart!;
        var imagePart = mainPart.AddImagePart(imgPartType);
        imagePart.FeedData(imgStream);
        var relId = mainPart.GetIdOfPart(imagePart);

        // Determine dimensions (default: 6 inches wide, auto height)
        long cxEmu = 5486400; // 6 inches in EMUs (914400 * 6)
        long cyEmu = 3657600; // 4 inches default
        if (properties.TryGetValue("width", out var widthStr))
            cxEmu = ParseEmu(widthStr);
        if (properties.TryGetValue("height", out var heightStr))
            cyEmu = ParseEmu(heightStr);

        var altText = properties.GetValueOrDefault("alt", Path.GetFileName(imgPath));

        Run imgRun;
        if (properties.TryGetValue("anchor", out var anchorVal) && IsTruthy(anchorVal))
        {
            var wrapType = properties.GetValueOrDefault("wrap", "none");
            long hPos = properties.TryGetValue("hposition", out var hPosStr) ? ParseEmu(hPosStr) : 0;
            long vPos = properties.TryGetValue("vposition", out var vPosStr) ? ParseEmu(vPosStr) : 0;
            var hRel = properties.TryGetValue("hrelative", out var hRelStr)
                ? ParseHorizontalRelative(hRelStr)
                : DW.HorizontalRelativePositionValues.Margin;
            var vRel = properties.TryGetValue("vrelative", out var vRelStr)
                ? ParseVerticalRelative(vRelStr)
                : DW.VerticalRelativePositionValues.Margin;
            var behind = properties.TryGetValue("behindtext", out var behindStr) && IsTruthy(behindStr);
            imgRun = CreateAnchorImageRun(relId, cxEmu, cyEmu, altText, wrapType, hPos, vPos, hRel, vRel, behind);
        }
        else
        {
            imgRun = CreateImageRun(relId, cxEmu, cyEmu, altText);
        }

        string resultPath;
        Paragraph imgPara;
        if (parent is Paragraph existingPara)
        {
            existingPara.AppendChild(imgRun);
            imgPara = existingPara;
            var imgRunCount = existingPara.Elements<Run>().Count();
            resultPath = $"{parentPath}/r[{imgRunCount}]";
        }
        else if (parent is TableCell imgCell)
        {
            // Insert image into existing first paragraph if empty, otherwise create new paragraph
            var firstCellPara = imgCell.Elements<Paragraph>().FirstOrDefault();
            if (firstCellPara != null && !firstCellPara.Elements<Run>().Any())
            {
                firstCellPara.AppendChild(imgRun);
                imgPara = firstCellPara;
            }
            else
            {
                imgPara = new Paragraph(imgRun);
                imgCell.AppendChild(imgPara);
            }
            var imgPIdx = imgCell.Elements<Paragraph>().ToList().IndexOf(imgPara) + 1;
            resultPath = $"{parentPath}/p[{imgPIdx}]";
        }
        else
        {
            imgPara = new Paragraph(imgRun);
            var imgParaCount = parent.Elements<Paragraph>().Count();
            if (index.HasValue && index.Value < imgParaCount)
            {
                var refPara = parent.Elements<Paragraph>().ElementAt(index.Value);
                parent.InsertBefore(imgPara, refPara);
                resultPath = $"{parentPath}/p[{index.Value + 1}]";
            }
            else
            {
                AppendToParent(parent, imgPara);
                resultPath = $"{parentPath}/p[{imgParaCount + 1}]";
            }
        }
        return resultPath;
    }
}
