// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Tests for cx:chart extended chart types (funnel, treemap, sunburst, boxWhisker, histogram)
/// and chart style presets (minimal, dark, corporate, magazine, dashboard, colorful, monochrome).
/// </summary>
public class ChartExAndPresetTests : IDisposable
{
    private readonly string _pptxPath;
    private readonly string _xlsxPath;
    private PowerPointHandler _pptx;
    private ExcelHandler _excel;

    public ChartExAndPresetTests()
    {
        _pptxPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.pptx");
        _xlsxPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.xlsx");
        BlankDocCreator.Create(_pptxPath);
        BlankDocCreator.Create(_xlsxPath);
        _pptx = new PowerPointHandler(_pptxPath, editable: true);
        _excel = new ExcelHandler(_xlsxPath, editable: true);
    }

    public void Dispose()
    {
        _pptx.Dispose();
        _excel.Dispose();
        if (File.Exists(_pptxPath)) File.Delete(_pptxPath);
        if (File.Exists(_xlsxPath)) File.Delete(_xlsxPath);
    }

    private void ReopenPptx() { _pptx.Dispose(); _pptx = new PowerPointHandler(_pptxPath, editable: true); }
    private void ReopenExcel() { _excel.Dispose(); _excel = new ExcelHandler(_xlsxPath, editable: true); }

    private string AddSlideAndExtChart(string chartType, string data, string categories, string title = "Test")
    {
        _pptx.Add("/", "slide", null, new());
        return _pptx.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = chartType,
            ["title"] = title,
            ["data"] = data,
            ["categories"] = categories
        });
    }

    private string AddExcelColumnChart()
    {
        return _excel.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "column",
            ["title"] = "Excel Test",
            ["data"] = "S1:10,20,30",
            ["categories"] = "A,B,C"
        });
    }

    // ==================== cx:chart Extended Types ====================

    [Fact]
    public void CxChart_Funnel_ChartType_And_SeriesCount()
    {
        var path = AddSlideAndExtChart(
            "funnel",
            "Pipeline:1200,900,600,300,150",
            "Leads,Qualified,Proposal,Negotiation,Won",
            "Sales Funnel");

        var node = _pptx.Get(path, depth: 0);
        node.Should().NotBeNull();
        node.Format["chartType"].Should().Be("funnel");
        node.Format["seriesCount"].Should().Be(1);
    }

    [Fact]
    public void CxChart_Funnel_Persists_After_Reopen()
    {
        var path = AddSlideAndExtChart(
            "funnel",
            "Pipeline:1200,900,600,300,150",
            "Leads,Qualified,Proposal,Negotiation,Won");

        ReopenPptx();
        var node = _pptx.Get(path, depth: 0);
        node.Format["chartType"].Should().Be("funnel");
        node.Format["seriesCount"].Should().Be(1);
    }

    [Fact]
    public void CxChart_Treemap_ChartType_And_SeriesCount()
    {
        var path = AddSlideAndExtChart(
            "treemap",
            "Sales:500,300,200,100,80,60",
            "Electronics,Clothing,Food,Books,Toys,Sports");

        var node = _pptx.Get(path, depth: 0);
        node.Format["chartType"].Should().Be("treemap");
        node.Format["seriesCount"].Should().Be(1);
    }

    [Fact]
    public void CxChart_Treemap_Persists_After_Reopen()
    {
        var path = AddSlideAndExtChart(
            "treemap",
            "Sales:500,300,200,100",
            "A,B,C,D");

        ReopenPptx();
        var node = _pptx.Get(path, depth: 0);
        node.Format["chartType"].Should().Be("treemap");
    }

    [Fact]
    public void CxChart_Sunburst_ChartType_And_SeriesCount()
    {
        var path = AddSlideAndExtChart(
            "sunburst",
            "Revenue:400,250,150,100,50",
            "North,South,East,West,Central");

        var node = _pptx.Get(path, depth: 0);
        node.Format["chartType"].Should().Be("sunburst");
        node.Format["seriesCount"].Should().Be(1);
    }

    [Fact]
    public void CxChart_Sunburst_Persists_After_Reopen()
    {
        var path = AddSlideAndExtChart(
            "sunburst",
            "Data:100,80,60,40",
            "Q1,Q2,Q3,Q4");

        ReopenPptx();
        var node = _pptx.Get(path, depth: 0);
        node.Format["chartType"].Should().Be("sunburst");
    }

    [Fact]
    public void CxChart_BoxWhisker_ChartType_And_SeriesCount()
    {
        var path = AddSlideAndExtChart(
            "boxWhisker",
            "Scores:55,60,65,70,75,80,85,90,95",
            "Jan,Feb,Mar,Apr,May,Jun,Jul,Aug,Sep");

        var node = _pptx.Get(path, depth: 0);
        node.Format["chartType"].Should().Be("boxWhisker");
        node.Format["seriesCount"].Should().Be(1);
    }

    [Fact]
    public void CxChart_BoxWhisker_Persists_After_Reopen()
    {
        var path = AddSlideAndExtChart(
            "boxWhisker",
            "Scores:55,60,70,80,90",
            "A,B,C,D,E");

        ReopenPptx();
        var node = _pptx.Get(path, depth: 0);
        node.Format["chartType"].Should().Be("boxWhisker");
    }

    [Fact]
    public void CxChart_Histogram_ChartType_And_SeriesCount()
    {
        var path = AddSlideAndExtChart(
            "histogram",
            "Values:12,15,18,22,25,28,30,32,35,38,40,42",
            "");

        var node = _pptx.Get(path, depth: 0);
        node.Format["chartType"].Should().Be("histogram");
        node.Format["seriesCount"].Should().Be(1);
    }

    [Fact]
    public void CxChart_Histogram_Persists_After_Reopen()
    {
        var path = AddSlideAndExtChart(
            "histogram",
            "Data:10,20,30,40,50,60",
            "");

        ReopenPptx();
        var node = _pptx.Get(path, depth: 0);
        node.Format["chartType"].Should().Be("histogram");
    }

    [Fact]
    public void CxChart_InvalidExtendedTypeName_ThrowsArgumentException()
    {
        // Unknown type is not in ExtendedChartTypes, falls through to ParseChartType which throws
        _pptx.Add("/", "slide", null, new());
        Action act = () => _pptx.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "not_a_real_type_xyz",
            ["data"] = "S1:10,20,30",
            ["categories"] = "A,B,C"
        });
        act.Should().Throw<ArgumentException>()
            .WithMessage("*not_a_real_type_xyz*");
    }

    // ==================== Preset Tests (PPTX) ====================

    [Fact]
    public void Preset_Minimal_AppliesWithoutError()
    {
        _pptx.Add("/", "slide", null, new());
        var chartPath = _pptx.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "column",
            ["title"] = "Minimal Test",
            ["data"] = "S1:10,20,30",
            ["categories"] = "A,B,C"
        });

        Action act = () => _pptx.Set(chartPath, new() { ["preset"] = "minimal" });
        act.Should().NotThrow();

        // Minimal preset sets gridlines — verify node is readable after preset
        var node = _pptx.Get(chartPath, depth: 0);
        node.Should().NotBeNull();
    }

    [Fact]
    public void Preset_Dark_SetsChartFill()
    {
        _pptx.Add("/", "slide", null, new());
        var chartPath = _pptx.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "bar",
            ["title"] = "Dark Test",
            ["data"] = "S1:5,10,15",
            ["categories"] = "X,Y,Z"
        });

        Action act = () => _pptx.Set(chartPath, new() { ["preset"] = "dark" });
        act.Should().NotThrow();

        var node = _pptx.Get(chartPath, depth: 0);
        node.Should().NotBeNull();
        // dark preset includes chartFill = 1E1E1E
        node.Format.Should().ContainKey("chartFill");
        node.Format["chartFill"].Should().Be("#1E1E1E");
    }

    [Fact]
    public void Preset_Corporate_AppliesWithoutError()
    {
        _pptx.Add("/", "slide", null, new());
        var chartPath = _pptx.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "line",
            ["title"] = "Corporate",
            ["data"] = "S1:1,2,3",
            ["categories"] = "A,B,C"
        });

        Action act = () => _pptx.Set(chartPath, new() { ["preset"] = "corporate" });
        act.Should().NotThrow();

        var node = _pptx.Get(chartPath, depth: 0);
        node.Should().NotBeNull();
    }

    [Fact]
    public void Preset_Magazine_SetsDataLabels()
    {
        _pptx.Add("/", "slide", null, new());
        var chartPath = _pptx.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "column",
            ["title"] = "Magazine",
            ["data"] = "S1:10,20,30",
            ["categories"] = "A,B,C"
        });

        Action act = () => _pptx.Set(chartPath, new() { ["preset"] = "magazine" });
        act.Should().NotThrow();

        var node = _pptx.Get(chartPath, depth: 0);
        node.Should().NotBeNull();
        // magazine preset enables datalabels
        node.Format.Should().ContainKey("dataLabels");
    }

    [Fact]
    public void Preset_Dashboard_AppliesWithoutError()
    {
        _pptx.Add("/", "slide", null, new());
        var chartPath = _pptx.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "column",
            ["title"] = "Dashboard",
            ["data"] = "S1:1,2,3",
            ["categories"] = "A,B,C"
        });

        Action act = () => _pptx.Set(chartPath, new() { ["preset"] = "dashboard" });
        act.Should().NotThrow();
    }

    [Fact]
    public void Preset_Colorful_AppliesWithoutError()
    {
        _pptx.Add("/", "slide", null, new());
        var chartPath = _pptx.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "pie",
            ["title"] = "Colorful",
            ["data"] = "S1:30,25,20,15,10",
            ["categories"] = "A,B,C,D,E"
        });

        Action act = () => _pptx.Set(chartPath, new() { ["preset"] = "colorful" });
        act.Should().NotThrow();
    }

    [Fact]
    public void Preset_Monochrome_AppliesWithoutError()
    {
        _pptx.Add("/", "slide", null, new());
        var chartPath = _pptx.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "area",
            ["title"] = "Mono",
            ["data"] = "S1:10,20,30",
            ["categories"] = "A,B,C"
        });

        Action act = () => _pptx.Set(chartPath, new() { ["preset"] = "monochrome" });
        act.Should().NotThrow();
    }

    [Fact]
    public void Preset_InvalidName_ThrowsArgumentException()
    {
        _pptx.Add("/", "slide", null, new());
        var chartPath = _pptx.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "column",
            ["data"] = "S1:1,2,3",
            ["categories"] = "A,B,C"
        });

        Action act = () => _pptx.Set(chartPath, new() { ["preset"] = "nonexistent_preset" });
        act.Should().Throw<ArgumentException>()
            .WithMessage("*nonexistent_preset*");
    }

    [Fact]
    public void Preset_StylePresetKey_Works()
    {
        _pptx.Add("/", "slide", null, new());
        var chartPath = _pptx.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "column",
            ["data"] = "S1:1,2,3",
            ["categories"] = "A,B,C"
        });

        Action act = () => _pptx.Set(chartPath, new() { ["style.preset"] = "minimal" });
        act.Should().NotThrow();
    }

    [Fact]
    public void Preset_ThemeKey_Works()
    {
        _pptx.Add("/", "slide", null, new());
        var chartPath = _pptx.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "column",
            ["data"] = "S1:1,2,3",
            ["categories"] = "A,B,C"
        });

        Action act = () => _pptx.Set(chartPath, new() { ["theme"] = "dark" });
        act.Should().NotThrow();
    }

    [Fact]
    public void Preset_ThenOverrideProperty_Succeeds()
    {
        _pptx.Add("/", "slide", null, new());
        var chartPath = _pptx.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "column",
            ["title"] = "Override Test",
            ["data"] = "S1:10,20,30",
            ["categories"] = "A,B,C"
        });

        // Apply dark preset first (sets chartFill = 1E1E1E)
        _pptx.Set(chartPath, new() { ["preset"] = "dark" });

        // Override chartFill after preset
        _pptx.Set(chartPath, new() { ["chartFill"] = "FF0000" });

        var node = _pptx.Get(chartPath, depth: 0);
        node.Format["chartFill"].Should().Be("#FF0000");
    }

    [Fact]
    public void Preset_Persists_After_Reopen()
    {
        _pptx.Add("/", "slide", null, new());
        var chartPath = _pptx.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "column",
            ["title"] = "Persist Preset",
            ["data"] = "S1:10,20,30",
            ["categories"] = "A,B,C"
        });

        _pptx.Set(chartPath, new() { ["preset"] = "dark" });

        ReopenPptx();
        var node = _pptx.Get(chartPath, depth: 0);
        node.Format.Should().ContainKey("chartFill");
        node.Format["chartFill"].Should().Be("#1E1E1E");
    }

    // ==================== Preset Tests (Excel) ====================

    [Fact]
    public void Excel_Preset_Minimal_AppliesWithoutError()
    {
        var chartPath = AddExcelColumnChart();
        Action act = () => _excel.Set(chartPath, new() { ["preset"] = "minimal" });
        act.Should().NotThrow();

        var node = _excel.Get(chartPath, depth: 0);
        node.Should().NotBeNull();
    }

    [Fact]
    public void Excel_Preset_Dark_SetsChartFill()
    {
        var chartPath = AddExcelColumnChart();
        _excel.Set(chartPath, new() { ["preset"] = "dark" });

        var node = _excel.Get(chartPath, depth: 0);
        node.Format.Should().ContainKey("chartFill");
        node.Format["chartFill"].Should().Be("#1E1E1E");
    }

    [Fact]
    public void Excel_Preset_InvalidName_ThrowsArgumentException()
    {
        var chartPath = AddExcelColumnChart();
        Action act = () => _excel.Set(chartPath, new() { ["preset"] = "does_not_exist" });
        act.Should().Throw<ArgumentException>()
            .WithMessage("*does_not_exist*");
    }

    [Fact]
    public void Excel_Preset_Persists_After_Reopen()
    {
        var chartPath = AddExcelColumnChart();
        _excel.Set(chartPath, new() { ["preset"] = "dark" });

        ReopenExcel();
        var node = _excel.Get(chartPath, depth: 0);
        node.Format.Should().ContainKey("chartFill");
        node.Format["chartFill"].Should().Be("#1E1E1E");
    }
}
