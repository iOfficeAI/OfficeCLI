// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Round 9 user interview tests — three schema ordering and readback bugs.
///
/// Bug 1: Scatter marker schema order — ApplySeriesMarker inserted marker before xVal/yVal,
///   violating CT_ScatterSer schema. Fixed: marker now inserted before data references.
///
/// Bug 2: Bubble bubbleScale schema order — AppendChild placed bubbleScale after axId.
///   Fixed: InsertBefore axId.
///
/// Bug 3a: View3D schema order — PrependChild placed view3D before title.
///   Fixed: InsertBefore plotArea.
///
/// Bug 3b: View3D readback — ChartReader did not read view3D properties.
///   Fixed: added RotateX, RotateY, Perspective readback.
/// </summary>
public class UserInterviewRound9Tests : IDisposable
{
    private readonly string _xlsxPath;
    private ExcelHandler _excel;

    public UserInterviewRound9Tests()
    {
        _xlsxPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.xlsx");
        BlankDocCreator.Create(_xlsxPath);
        _excel = new ExcelHandler(_xlsxPath, editable: true);
    }

    public void Dispose()
    {
        _excel?.Dispose();
        if (File.Exists(_xlsxPath)) File.Delete(_xlsxPath);
    }

    private void Reopen()
    {
        _excel?.Dispose();
        _excel = new ExcelHandler(_xlsxPath, editable: true);
    }

    private string AddChart(string chartType = "column", Dictionary<string, string>? extra = null)
    {
        var props = new Dictionary<string, string>
        {
            ["chartType"] = chartType,
            ["title"] = "Test",
            ["data"] = "S1:10,20,30;S2:15,25,35",
            ["categories"] = "A,B,C",
            ["legend"] = "bottom"
        };
        if (extra != null) foreach (var kv in extra) props[kv.Key] = kv.Value;
        return _excel.Add("/Sheet1", "chart", null, props);
    }

    // ==================== Bug 1: Scatter marker schema order ====================

    [Fact]
    public void Set_ScatterSeries_Marker_SchemaOrder_MarkerBeforeXVal()
    {
        // Create scatter chart, then set marker on series
        var path = AddChart("scatter");
        _excel.Set(path, new() { ["series1.marker"] = "circle:8" });

        // Verify the file is valid by reopening
        Reopen();
        var node = _excel.Get(path, depth: 1);
        node.Should().NotBeNull();

        // Dispose handler to release file lock before direct OpenXml access
        _excel.Dispose();
        using var doc = SpreadsheetDocument.Open(_xlsxPath, false);
        var chartPart = doc.WorkbookPart!.GetPartsOfType<WorksheetPart>()
            .SelectMany(wp => wp.DrawingsPart?.ChartParts ?? Enumerable.Empty<ChartPart>())
            .First();
        var scatterSer = chartPart.ChartSpace
            .Descendants<C.ScatterChartSeries>().First();
        var children = scatterSer.ChildElements.Select(e => e.LocalName).ToList();

        var markerIdx = children.IndexOf("marker");
        var xValIdx = children.IndexOf("xVal");
        var yValIdx = children.IndexOf("yVal");

        markerIdx.Should().BeGreaterOrEqualTo(0, "marker should exist");
        if (xValIdx >= 0) markerIdx.Should().BeLessThan(xValIdx, "marker must come before xVal in CT_ScatterSer schema");
        if (yValIdx >= 0) markerIdx.Should().BeLessThan(yValIdx, "marker must come before yVal in CT_ScatterSer schema");
    }

    [Fact]
    public void Set_ScatterSeries_Marker_PersistsAfterReopen()
    {
        var path = AddChart("scatter");
        _excel.Set(path, new() { ["series1.marker"] = "diamond:6:FF0000" });

        Reopen();
        var node = _excel.Get(path, depth: 1);
        node.Should().NotBeNull();

        // Verify series child has marker info
        var series = node.Children?.FirstOrDefault(c => c.Type == "series");
        series.Should().NotBeNull();
        series!.Format.Should().ContainKey("marker");
    }

    // ==================== Bug 2: Bubble bubbleScale schema order ====================

    [Fact]
    public void Set_BubbleChart_BubbleScale_SchemaOrder_BeforeAxId()
    {
        var path = AddChart("bubble");
        _excel.Set(path, new() { ["bubbleScale"] = "150" });

        Reopen();
        var node = _excel.Get(path, depth: 0);
        node.Format["bubbleScale"].Should().Be(150);

        // Dispose handler to release file lock before direct OpenXml access
        _excel.Dispose();
        using var doc = SpreadsheetDocument.Open(_xlsxPath, false);
        var chartPart = doc.WorkbookPart!.GetPartsOfType<WorksheetPart>()
            .SelectMany(wp => wp.DrawingsPart?.ChartParts ?? Enumerable.Empty<ChartPart>())
            .First();
        var bubbleChart = chartPart.ChartSpace
            .Descendants<C.BubbleChart>().First();
        var children = bubbleChart.ChildElements.Select(e => e.LocalName).ToList();

        var scaleIdx = children.IndexOf("bubbleScale");
        var axIdIdx = children.IndexOf("axId");

        scaleIdx.Should().BeGreaterOrEqualTo(0, "bubbleScale should exist");
        axIdIdx.Should().BeGreaterOrEqualTo(0, "axId should exist");
        scaleIdx.Should().BeLessThan(axIdIdx, "bubbleScale must come before axId in CT_BubbleChart schema");
    }

    [Fact]
    public void Set_BubbleChart_BubbleScale_ReadBack()
    {
        var path = AddChart("bubble");
        _excel.Set(path, new() { ["bubbleScale"] = "200" });

        var node = _excel.Get(path, depth: 0);
        node.Format["bubbleScale"].Should().Be(200);

        // Verify persistence
        Reopen();
        node = _excel.Get(path, depth: 0);
        node.Format["bubbleScale"].Should().Be(200);
    }

    // ==================== Bug 3a: View3D schema order ====================

    [Fact]
    public void Set_View3D_SchemaOrder_BeforePlotArea()
    {
        var path = AddChart("column");
        _excel.Set(path, new() { ["view3d"] = "15,20,30" });

        Reopen();

        // Dispose handler to release file lock before direct OpenXml access
        _excel.Dispose();
        using var doc = SpreadsheetDocument.Open(_xlsxPath, false);
        var chartPart = doc.WorkbookPart!.GetPartsOfType<WorksheetPart>()
            .SelectMany(wp => wp.DrawingsPart?.ChartParts ?? Enumerable.Empty<ChartPart>())
            .First();
        var chart = chartPart.ChartSpace.GetFirstChild<C.Chart>()!;
        var children = chart.ChildElements.Select(e => e.LocalName).ToList();

        var view3dIdx = children.IndexOf("view3D");
        var plotAreaIdx = children.IndexOf("plotArea");

        view3dIdx.Should().BeGreaterOrEqualTo(0, "view3D should exist");
        plotAreaIdx.Should().BeGreaterOrEqualTo(0, "plotArea should exist");
        view3dIdx.Should().BeLessThan(plotAreaIdx, "view3D must come before plotArea in CT_Chart schema");

        // Also verify it doesn't come before title if title exists
        var titleIdx = children.IndexOf("title");
        if (titleIdx >= 0)
            view3dIdx.Should().BeGreaterThan(titleIdx, "view3D must come after title in CT_Chart schema");
    }

    // ==================== Bug 3b: View3D readback ====================

    [Fact]
    public void Set_View3D_ReadBack_AllProperties()
    {
        var path = AddChart("column");
        _excel.Set(path, new() { ["view3d"] = "15,20,30" });

        var node = _excel.Get(path, depth: 0);
        node.Format.Should().ContainKey("view3d");
        node.Format["view3d"].Should().Be("15,20,30");
        node.Format["view3d.rotateX"].Should().Be(15);
        node.Format["view3d.rotateY"].Should().Be(20);
        node.Format["view3d.perspective"].Should().Be(30);
    }

    [Fact]
    public void Set_View3D_ReadBack_PersistsAfterReopen()
    {
        var path = AddChart("column");
        _excel.Set(path, new() { ["view3d"] = "10,25,40" });

        Reopen();
        var node = _excel.Get(path, depth: 0);
        node.Format.Should().ContainKey("view3d");
        node.Format["view3d"].Should().Be("10,25,40");
        node.Format["view3d.rotateX"].Should().Be(10);
        node.Format["view3d.rotateY"].Should().Be(25);
        node.Format["view3d.perspective"].Should().Be(40);
    }

    [Fact]
    public void Set_View3D_SingleValue_PerspectiveOnly()
    {
        var path = AddChart("column");
        _excel.Set(path, new() { ["view3d"] = "30" });

        var node = _excel.Get(path, depth: 0);
        node.Format.Should().ContainKey("view3d");
        node.Format["view3d.perspective"].Should().Be(30);
    }
}
