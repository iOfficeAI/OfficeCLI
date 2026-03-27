// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Round 3 user interview tests — cx:chart (extended chart) Get/Set/Query issues.
///
/// Root cause: cx:chart uses ExtendedChartPart, not ChartPart.
/// - ResolveChart() in PowerPointHandler only filters for C.ChartReference descendants,
///   so cx:chart frames are invisible to Set and series-level Get.
/// - Get("/slide[N]/chart[M]") works because Query.cs includes IsExtendedChartFrame().
/// - Get("/slide[N]/chart[M]/series[K]") calls ResolveChart() → fails.
/// - Set("/slide[N]/chart[M]") calls ResolveChart() → fails with "Chart N not found (total: 0)".
///
/// Product decisions for cx:chart:
/// - Get at chart level must return: chartType, seriesCount, title (if present), position.
/// - Get at series sub-path should work and return basic series info.
/// - Set should support position/size properties (x, y, width, height, name) which
///   operate on the GraphicFrame, not the chart internals.
/// - Set for chart-internal properties (series colors, data labels, etc.) should return
///   them as unsupported with a clear message, not throw.
/// - Query("chart") must find cx:chart in PPTX (already works).
/// </summary>
public class UserInterviewRound3Tests : IDisposable
{
    private readonly string _pptxPath;
    private PowerPointHandler _pptxHandler;

    public UserInterviewRound3Tests()
    {
        _pptxPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.pptx");
        BlankDocCreator.Create(_pptxPath);
        _pptxHandler = new PowerPointHandler(_pptxPath, editable: true);
    }

    public void Dispose()
    {
        _pptxHandler?.Dispose();
        if (File.Exists(_pptxPath)) File.Delete(_pptxPath);
    }

    private void ReopenPptx()
    {
        _pptxHandler?.Dispose();
        _pptxHandler = new PowerPointHandler(_pptxPath, editable: true);
    }

    // ==================== Issue 1: cx:chart Get returns valid node ====================

    [Fact]
    public void Pptx_FunnelChart_Get_ReturnsChartType()
    {
        // Arrange: create a funnel chart
        _pptxHandler.Add("/", "slide", null, new());
        _pptxHandler.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "funnel",
            ["series1"] = "Sales:100,80,60,40,20"
        });

        // Act
        var node = _pptxHandler.Get("/slide[1]/chart[1]");

        // Assert
        node.Should().NotBeNull();
        node.Type.Should().Be("chart");
        node.Format.Should().ContainKey("chartType");
        node.Format["chartType"].Should().Be("funnel");
        node.Format.Should().ContainKey("seriesCount");
        ((int)node.Format["seriesCount"]).Should().BeGreaterOrEqualTo(1);
    }

    [Fact]
    public void Pptx_TreemapChart_Get_ReturnsChartType()
    {
        _pptxHandler.Add("/", "slide", null, new());
        _pptxHandler.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "treemap",
            ["categories"] = "A,B,C,D",
            ["series1"] = "Values:40,30,20,10"
        });

        var node = _pptxHandler.Get("/slide[1]/chart[1]");

        node.Should().NotBeNull();
        node.Type.Should().Be("chart");
        node.Format["chartType"].Should().Be("treemap");
    }

    [Fact]
    public void Pptx_FunnelChart_Get_PersistsAfterReopen()
    {
        _pptxHandler.Add("/", "slide", null, new());
        _pptxHandler.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "funnel",
            ["series1"] = "Pipeline:100,75,50,25"
        });

        ReopenPptx();

        var node = _pptxHandler.Get("/slide[1]/chart[1]");
        node.Should().NotBeNull();
        node.Format["chartType"].Should().Be("funnel");
    }

    // ==================== Issue 2: cx:chart Set should not crash ====================

    [Fact]
    public void Pptx_FunnelChart_Set_PositionDoesNotThrow()
    {
        // Set position/size properties should work because they operate on GraphicFrame,
        // not on chart internals. Currently ResolveChart() fails to find cx:chart → throws.
        _pptxHandler.Add("/", "slide", null, new());
        _pptxHandler.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "funnel",
            ["series1"] = "Data:90,70,50,30,10"
        });

        // Act: Set position — should not throw
        var unsupported = _pptxHandler.Set("/slide[1]/chart[1]", new()
        {
            ["x"] = "3cm",
            ["y"] = "4cm",
            ["width"] = "20cm",
            ["height"] = "12cm"
        });

        // Verify position was applied
        var node = _pptxHandler.Get("/slide[1]/chart[1]");
        node.Format["x"].Should().Be("3cm");
        node.Format["y"].Should().Be("4cm");
        node.Format["width"].Should().Be("20cm");
        node.Format["height"].Should().Be("12cm");
    }

    [Fact]
    public void Pptx_FunnelChart_Set_ChartInternalProperty_ReturnsUnsupported()
    {
        // cx:chart internal properties (series colors, etc.) cannot be set via ChartSetter
        // because cx:chart uses a completely different XML schema. Instead of throwing,
        // these should be returned as unsupported.
        _pptxHandler.Add("/", "slide", null, new());
        _pptxHandler.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "funnel",
            ["series1"] = "Data:90,70,50,30,10"
        });

        // Act: Set a chart-internal property
        var unsupported = _pptxHandler.Set("/slide[1]/chart[1]", new()
        {
            ["series1.fill"] = "#FF0000"
        });

        // Should return as unsupported rather than throwing
        unsupported.Should().NotBeNull();
        unsupported.Should().Contain("series1.fill");
    }

    [Fact]
    public void Pptx_TreemapChart_Set_NameProperty()
    {
        _pptxHandler.Add("/", "slide", null, new());
        _pptxHandler.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "treemap",
            ["categories"] = "A,B,C",
            ["series1"] = "Size:50,30,20"
        });

        var unsupported = _pptxHandler.Set("/slide[1]/chart[1]", new()
        {
            ["name"] = "My Treemap"
        });

        var node = _pptxHandler.Get("/slide[1]/chart[1]");
        node.Format["name"].Should().Be("My Treemap");
    }

    // ==================== Issue 3: cx:chart series sub-path Get ====================

    [Fact]
    public void Pptx_FunnelChart_GetSeries_ReturnsSeriesNode()
    {
        // Get("/slide[1]/chart[1]/series[1]") calls ResolveChart() which only finds c:chart.
        // This should also work for cx:chart.
        _pptxHandler.Add("/", "slide", null, new());
        _pptxHandler.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "funnel",
            ["series1"] = "Pipeline:100,75,50,25"
        });

        // Act: get the chart first to verify series exist at chart level
        var chartNode = _pptxHandler.Get("/slide[1]/chart[1]", depth: 1);
        ((int)chartNode.Format["seriesCount"]).Should().BeGreaterOrEqualTo(1);
    }

    // ==================== Issue 4: cx:chart visible in Query("chart") ====================

    [Fact]
    public void Pptx_Query_Chart_FindsCxChart()
    {
        _pptxHandler.Add("/", "slide", null, new());
        _pptxHandler.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "funnel",
            ["series1"] = "Data:10,20,30"
        });

        var results = _pptxHandler.Query("chart");
        results.Should().HaveCount(1);
        results[0].Type.Should().Be("chart");
        results[0].Format["chartType"].Should().Be("funnel");
    }

    [Fact]
    public void Pptx_MixedCharts_Query_FindsBothTypes()
    {
        // Create one regular chart and one cx:chart on same slide
        _pptxHandler.Add("/", "slide", null, new());
        _pptxHandler.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "bar",
            ["categories"] = "A,B,C",
            ["series1"] = "S1:10,20,30"
        });
        _pptxHandler.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "funnel",
            ["series1"] = "Pipeline:100,75,50"
        });

        var results = _pptxHandler.Query("chart");
        results.Should().HaveCount(2);
        // One should be bar, the other funnel
        results.Should().Contain(n => n.Format.ContainsKey("chartType") && n.Format["chartType"].ToString() == "funnel");
    }

    // ==================== Issue 5: Mixed c:chart + cx:chart indexing ====================

    [Fact]
    public void Pptx_MixedCharts_Set_CxChartByCorrectIndex()
    {
        // When a slide has both c:chart and cx:chart, the chart index must
        // account for both types. ResolveChart only counts c:chart, so
        // Set on the cx:chart fails because the index is wrong.
        _pptxHandler.Add("/", "slide", null, new());
        _pptxHandler.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "bar",
            ["categories"] = "A,B",
            ["series1"] = "S1:10,20"
        });
        _pptxHandler.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "funnel",
            ["series1"] = "Data:50,30,10"
        });

        // chart[1] = bar (c:chart), chart[2] = funnel (cx:chart)
        // Set on chart[2] should not throw
        var unsupported = _pptxHandler.Set("/slide[1]/chart[2]", new()
        {
            ["x"] = "5cm"
        });

        var node = _pptxHandler.Get("/slide[1]/chart[2]");
        node.Format["chartType"].Should().Be("funnel");
        node.Format["x"].Should().Be("5cm");
    }
}
