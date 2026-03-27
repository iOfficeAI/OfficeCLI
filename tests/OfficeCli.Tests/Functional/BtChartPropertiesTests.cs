// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Black-box tests for chart properties from a user perspective.
/// Covers combinations, edge cases, cross-handler behavior, and persistence.
/// </summary>
public class BtChartPropertiesTests : IDisposable
{
    private readonly string _xlsxPath;
    private readonly string _pptxPath;
    private ExcelHandler _excel;
    private PowerPointHandler _pptx;

    public BtChartPropertiesTests()
    {
        _xlsxPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.xlsx");
        _pptxPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.pptx");
        BlankDocCreator.Create(_xlsxPath);
        BlankDocCreator.Create(_pptxPath);
        _excel = new ExcelHandler(_xlsxPath, editable: true);
        _pptx = new PowerPointHandler(_pptxPath, editable: true);
    }

    public void Dispose()
    {
        _excel.Dispose();
        _pptx.Dispose();
        if (File.Exists(_xlsxPath)) File.Delete(_xlsxPath);
        if (File.Exists(_pptxPath)) File.Delete(_pptxPath);
    }

    private void ReopenExcel() { _excel.Dispose(); _excel = new ExcelHandler(_xlsxPath, editable: true); }
    private void ReopenPptx() { _pptx.Dispose(); _pptx = new PowerPointHandler(_pptxPath, editable: true); }

    private string AddExcelChart(string chartType = "column", string data = "S1:10,20,30;S2:15,25,35", string categories = "A,B,C")
    {
        return _excel.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = chartType,
            ["title"] = "Test Chart",
            ["data"] = data,
            ["categories"] = categories
        });
    }

    private string AddPptxChart(string chartType = "line", string data = "S1:10,20,30;S2:15,25,35", string categories = "A,B,C")
    {
        _pptx.Add("/", "slide", null, new());
        return _pptx.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = chartType,
            ["title"] = "Test Chart",
            ["data"] = data,
            ["categories"] = categories
        });
    }

    // ==================== Scenario 1: Combined Set with 5+ categories at once ====================

    [Fact]
    public void Excel_CombinedSet_FivePlusCategories_AllApplied()
    {
        var chartPath = AddExcelChart("line");

        _excel.Set(chartPath, new()
        {
            ["title"] = "Combined",
            ["legend"] = "top",
            ["dataLabels"] = "value",
            ["gridlines"] = "CCCCCC:0.5",
            ["smooth"] = "true",
            ["majorTickMark"] = "in",
            ["style"] = "5"
        });

        var node = _excel.Get(chartPath, depth: 0);
        ((string)node.Format["title"]).Should().Be("Combined");
        ((string)node.Format["legend"]).Should().BeOneOf("t", "top");
        node.Format["dataLabels"].Should().Be("value");
        node.Format["gridlines"].Should().Be("true");
        node.Format["smooth"].Should().Be("true");
        node.Format["majorTickMark"].Should().NotBeNull();
        node.Format["style"].Should().Be((byte)5);
    }

    [Fact]
    public void Pptx_CombinedSet_FivePlusCategories_AllApplied()
    {
        var chartPath = AddPptxChart();

        _pptx.Set(chartPath, new()
        {
            ["title"] = "PPTX Combined",
            ["legend"] = "bottom",
            ["dataLabels"] = "value,category",
            ["labelPos"] = "top",
            ["gridlines"] = "DDDDDD:0.3",
            ["lineWidth"] = "2",
            ["style"] = "8"
        });

        var node = _pptx.Get(chartPath, depth: 0);
        ((string)node.Format["title"]).Should().Be("PPTX Combined");
        node.Format["dataLabels"].Should().NotBeNull();
        node.Format["gridlines"].Should().Be("true");
        node.Format["style"].Should().Be((byte)8);
    }

    // ==================== Scenario 2: Excel chart supports new properties ====================

    [Fact]
    public void Excel_Trendline_Set_And_Readable()
    {
        var chartPath = AddExcelChart("line");
        _excel.Set(chartPath, new() { ["trendline"] = "linear" });

        var node = _excel.Get(chartPath, depth: 1);
        node.Children.Should().NotBeEmpty();
        node.Children[0].Format["trendline"].Should().Be("linear");
    }

    [Fact]
    public void Excel_DataTable_Set_And_Readable()
    {
        var chartPath = AddExcelChart();
        _excel.Set(chartPath, new() { ["dataTable"] = "true" });

        var node = _excel.Get(chartPath, depth: 0);
        node.Format["dataTable"].Should().Be("true");
    }

    [Fact]
    public void Excel_ChartBorder_Set_NoException()
    {
        var chartPath = AddExcelChart();
        var act = () => _excel.Set(chartPath, new() { ["chartBorder"] = "000000:1" });
        act.Should().NotThrow();
    }

    [Fact]
    public void Excel_PlotBorder_Set_NoException()
    {
        var chartPath = AddExcelChart();
        var act = () => _excel.Set(chartPath, new() { ["plotArea.border"] = "CCCCCC:0.5" });
        act.Should().NotThrow();
    }

    [Fact]
    public void Excel_MajorTickMark_Set_And_Readable()
    {
        var chartPath = AddExcelChart();
        _excel.Set(chartPath, new() { ["majorTickMark"] = "in" });

        var node = _excel.Get(chartPath, depth: 0);
        node.Format["majorTickMark"].Should().NotBeNull();
    }

    [Fact]
    public void Excel_MinorTickMark_Set_And_Readable()
    {
        var chartPath = AddExcelChart();
        _excel.Set(chartPath, new() { ["minorTickMark"] = "out" });

        var node = _excel.Get(chartPath, depth: 0);
        node.Format["minorTickMark"].Should().NotBeNull();
    }

    [Fact]
    public void Excel_DispUnits_Set_And_Readable()
    {
        var chartPath = AddExcelChart();
        _excel.Set(chartPath, new() { ["dispUnits"] = "thousands" });

        var node = _excel.Get(chartPath, depth: 0);
        node.Format["dispUnits"].Should().NotBeNull();
    }

    // ==================== Scenario 3: Trendline preserved after series data Set ====================

    [Fact]
    public void Excel_Trendline_PreservedAfterSeriesDataUpdate()
    {
        var chartPath = AddExcelChart("line");
        _excel.Set(chartPath, new() { ["trendline"] = "linear" });

        // Verify trendline is set
        var before = _excel.Get(chartPath, depth: 1);
        before.Children[0].Format["trendline"].Should().Be("linear");

        // Update series data
        _excel.Set(chartPath, new() { ["series1"] = "S1:100,200,300" });

        // Trendline should still be present on series 1
        var after = _excel.Get(chartPath, depth: 1);
        after.Children[0].Format["trendline"].Should().Be("linear",
            "trendline should be preserved when series data is updated");
    }

    [Fact]
    public void Pptx_Trendline_PreservedAfterSeriesDataUpdate()
    {
        var chartPath = AddPptxChart();
        _pptx.Set(chartPath, new() { ["trendline"] = "linear" });

        var before = _pptx.Get(chartPath, depth: 1);
        before.Children[0].Format["trendline"].Should().Be("linear");

        _pptx.Set(chartPath, new() { ["series1"] = "S1:100,200,300" });

        var after = _pptx.Get(chartPath, depth: 1);
        after.Children[0].Format["trendline"].Should().Be("linear",
            "trendline should be preserved after series data update in PPTX");
    }

    // ==================== Scenario 4: dataTable + legend combination ====================

    [Fact]
    public void Excel_DataTableAndLegend_BothApplied()
    {
        var chartPath = AddExcelChart();
        _excel.Set(chartPath, new() { ["dataTable"] = "true", ["legend"] = "top" });

        var node = _excel.Get(chartPath, depth: 0);
        node.Format["dataTable"].Should().Be("true");
        ((string)node.Format["legend"]).Should().BeOneOf("t", "top");
    }

    [Fact]
    public void Pptx_DataTableAndLegend_BothApplied()
    {
        var chartPath = AddPptxChart();
        _pptx.Set(chartPath, new() { ["dataTable"] = "true", ["legend"] = "bottom" });

        var node = _pptx.Get(chartPath, depth: 0);
        node.Format["dataTable"].Should().Be("true");
        node.Format["legend"].Should().NotBeNull();
    }

    // ==================== Scenario 5: smooth on bar chart — should not crash ====================

    [Fact]
    public void Excel_Smooth_OnBarChart_DoesNotCrash()
    {
        var chartPath = AddExcelChart("bar");
        var act = () => _excel.Set(chartPath, new() { ["smooth"] = "true" });
        act.Should().NotThrow("smooth on bar chart should be silently ignored or handled gracefully");
    }

    [Fact]
    public void Pptx_Smooth_OnBarChart_DoesNotCrash()
    {
        _pptx.Add("/", "slide", null, new());
        var chartPath = _pptx.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "bar",
            ["data"] = "S1:1,2,3",
            ["categories"] = "A,B,C"
        });
        var act = () => _pptx.Set(chartPath, new() { ["smooth"] = "true" });
        act.Should().NotThrow("smooth on bar chart should be silently ignored or handled gracefully");
    }

    // ==================== Scenario 6: Invalid enum values do not crash ====================

    [Fact]
    public void Excel_InvalidMajorTickMark_DoesNotCrash()
    {
        var chartPath = AddExcelChart();
        var act = () => _excel.Set(chartPath, new() { ["majorTickMark"] = "invalid_value" });
        act.Should().NotThrow("invalid majorTickMark enum should not cause an exception");
    }

    [Fact]
    public void Excel_InvalidLabelPos_DoesNotCrash()
    {
        var chartPath = AddExcelChart();
        var act = () => _excel.Set(chartPath, new() { ["dataLabels"] = "value", ["labelPos"] = "not_a_real_position" });
        act.Should().NotThrow("invalid labelPos should not crash");
    }

    [Fact]
    public void Excel_InvalidDispUnits_DoesNotCrash()
    {
        var chartPath = AddExcelChart();
        var act = () => _excel.Set(chartPath, new() { ["dispUnits"] = "not_a_unit" });
        act.Should().NotThrow("invalid dispUnits falls back gracefully");
    }

    [Fact]
    public void Excel_InvalidMinorTickMark_DoesNotCrash()
    {
        var chartPath = AddExcelChart();
        var act = () => _excel.Set(chartPath, new() { ["minorTickMark"] = "GARBAGE" });
        act.Should().NotThrow("invalid minorTickMark should not crash");
    }

    // ==================== Scenario 7: dispUnits + displayUnitsLabel position ====================

    [Fact]
    public void Excel_DispUnits_ThenDisplayUnitsLabel_Position()
    {
        var chartPath = AddExcelChart();
        _excel.Set(chartPath, new() { ["dispUnits"] = "thousands" });

        // Now set the label position
        var act = () => _excel.Set(chartPath, new()
        {
            ["displayUnitsLabel.x"] = "0.1",
            ["displayUnitsLabel.y"] = "0.05"
        });
        act.Should().NotThrow("displayUnitsLabel.x/y should work after dispUnits is set");
    }

    [Fact]
    public void Pptx_DispUnits_ThenDisplayUnitsLabel_Position()
    {
        var chartPath = AddPptxChart();
        _pptx.Set(chartPath, new() { ["dispUnits"] = "millions" });

        var act = () => _pptx.Set(chartPath, new()
        {
            ["displayUnitsLabel.x"] = "0.1",
            ["displayUnitsLabel.y"] = "0.1"
        });
        act.Should().NotThrow("displayUnitsLabel positioning should not crash in PPTX");
    }

    [Fact]
    public void Excel_DisplayUnitsLabel_WithoutDispUnits_IsHandledGracefully()
    {
        var chartPath = AddExcelChart();
        // No dispUnits set — displayUnitsLabel won't exist, should not crash
        var act = () => _excel.Set(chartPath, new()
        {
            ["displayUnitsLabel.x"] = "0.2",
            ["displayUnitsLabel.y"] = "0.1"
        });
        act.Should().NotThrow("displayUnitsLabel position without prior dispUnits should not throw");
    }

    // ==================== Scenario 8: per-series combinations ====================

    [Fact]
    public void Excel_PerSeries_ColorAndTrendlineAndSmooth_LineChart()
    {
        var chartPath = AddExcelChart("line");

        _excel.Set(chartPath, new()
        {
            ["series1.color"] = "FF0000",
            ["series1.trendline"] = "linear",
            ["series1.smooth"] = "true"
        });

        var node = _excel.Get(chartPath, depth: 1);
        node.Children.Should().NotBeEmpty();
        var s1 = node.Children[0];
        s1.Format.Should().ContainKey("color");
        ((string)s1.Format["color"]).Should().BeOneOf("#FF0000", "FF0000");
        s1.Format["trendline"].Should().Be("linear");
        s1.Format["smooth"].Should().Be("true");
    }

    [Fact]
    public void Pptx_PerSeries_ColorAndTrendlineAndSmooth_LineChart()
    {
        var chartPath = AddPptxChart("line");

        _pptx.Set(chartPath, new()
        {
            ["series1.color"] = "4472C4",
            ["series1.trendline"] = "linear",
            ["series1.smooth"] = "true"
        });

        var node = _pptx.Get(chartPath, depth: 1);
        node.Children.Should().NotBeEmpty();
        var s1 = node.Children[0];
        s1.Format.Should().ContainKey("color");
        s1.Format["trendline"].Should().Be("linear");
        s1.Format["smooth"].Should().Be("true");
    }

    // ==================== Scenario 9: Multi-series — only modify one ====================

    [Fact]
    public void Excel_MultiSeries_ModifyOnlySeriesTwo_SeriesOneUnchanged()
    {
        var chartPath = AddExcelChart("line");

        // Set series 1 color first
        _excel.Set(chartPath, new() { ["series1.color"] = "FF0000" });

        // Now only modify series 2
        _excel.Set(chartPath, new() { ["series2.color"] = "0000FF" });

        var node = _excel.Get(chartPath, depth: 1);
        node.Children.Count.Should().BeGreaterOrEqualTo(2);
        // Series 1 should still have the original color
        ((string)node.Children[0].Format["color"]).Should().BeOneOf("#FF0000", "FF0000",
            "series1 color should be unaffected when only series2 is modified");
        ((string)node.Children[1].Format["color"]).Should().BeOneOf("#0000FF", "0000FF");
    }

    [Fact]
    public void Pptx_MultiSeries_ModifyOnlySeriesTwo_SeriesOneUnchanged()
    {
        var chartPath = AddPptxChart("line");

        _pptx.Set(chartPath, new() { ["series1.color"] = "FF0000" });
        _pptx.Set(chartPath, new() { ["series2.color"] = "00FF00" });

        var node = _pptx.Get(chartPath, depth: 1);
        node.Children.Count.Should().BeGreaterOrEqualTo(2);
        ((string)node.Children[0].Format["color"]).Should().BeOneOf("#FF0000", "FF0000",
            "series1 color should remain after only series2 was modified");
        ((string)node.Children[1].Format["color"]).Should().BeOneOf("#00FF00", "00FF00");
    }

    [Fact]
    public void Excel_MultiSeries_GlobalTrendline_ThenPerSeriesColor_BothApplied()
    {
        var chartPath = AddExcelChart("line");

        _excel.Set(chartPath, new() { ["trendline"] = "linear" });
        _excel.Set(chartPath, new() { ["series2.color"] = "FFAA00" });

        var node = _excel.Get(chartPath, depth: 1);
        // Both series should have trendlines (global set)
        node.Children[0].Format["trendline"].Should().Be("linear",
            "global trendline should be on series1 after per-series color change");
        node.Children[1].Format["trendline"].Should().Be("linear",
            "global trendline should be on series2 after per-series color change");
        // Series 2 should have the color
        ((string)node.Children[1].Format["color"]).Should().BeOneOf("#FFAA00", "FFAA00");
    }

    // ==================== Scenario 10: Reopen persistence for complex properties ====================

    [Fact]
    public void Excel_Reopen_Trendline_Persists()
    {
        var chartPath = AddExcelChart("line");
        _excel.Set(chartPath, new() { ["trendline"] = "linear" });

        ReopenExcel();
        var node = _excel.Get(chartPath, depth: 1);
        node.Children[0].Format["trendline"].Should().Be("linear",
            "trendline should persist after reopen");
    }

    [Fact]
    public void Excel_Reopen_DataTable_Persists()
    {
        var chartPath = AddExcelChart();
        _excel.Set(chartPath, new() { ["dataTable"] = "true" });

        ReopenExcel();
        var node = _excel.Get(chartPath, depth: 0);
        node.Format["dataTable"].Should().Be("true",
            "dataTable should persist after reopen");
    }

    [Fact]
    public void Excel_Reopen_ChartBorder_Persists()
    {
        var chartPath = AddExcelChart();
        _excel.Set(chartPath, new() { ["chartBorder"] = "4472C4:1.5" });

        ReopenExcel();
        var node = _excel.Get(chartPath, depth: 0);
        node.Format.Should().ContainKey("chartArea.border.color",
            "chartArea border should persist after reopen");
    }

    [Fact]
    public void Excel_Reopen_PlotBorder_Persists()
    {
        var chartPath = AddExcelChart();
        _excel.Set(chartPath, new() { ["plotArea.border"] = "888888:1" });

        ReopenExcel();
        var node = _excel.Get(chartPath, depth: 0);
        node.Format.Should().ContainKey("plotArea.border.color",
            "plotArea border should persist after reopen");
    }

    [Fact]
    public void Excel_Reopen_MajorTickMark_Persists()
    {
        var chartPath = AddExcelChart();
        _excel.Set(chartPath, new() { ["majorTickMark"] = "in" });

        ReopenExcel();
        var node = _excel.Get(chartPath, depth: 0);
        node.Format["majorTickMark"].Should().Be("in",
            "majorTickMark should persist after reopen");
    }

    [Fact]
    public void Pptx_Reopen_TrendlineAndDataTable_Persist()
    {
        var chartPath = AddPptxChart("line");
        _pptx.Set(chartPath, new()
        {
            ["trendline"] = "linear",
            ["dataTable"] = "true"
        });

        ReopenPptx();
        var node = _pptx.Get(chartPath, depth: 1);
        node.Format["dataTable"].Should().Be("true",
            "dataTable should persist after PPTX reopen");
        node.Children[0].Format["trendline"].Should().Be("linear",
            "trendline should persist after PPTX reopen");
    }

    [Fact]
    public void Excel_Reopen_PerSeriesProperties_Persist()
    {
        var chartPath = AddExcelChart("line");
        _excel.Set(chartPath, new()
        {
            ["series1.color"] = "FF0000",
            ["series1.trendline"] = "polynomial",
            ["series2.color"] = "0000FF"
        });

        ReopenExcel();
        var node = _excel.Get(chartPath, depth: 1);
        ((string)node.Children[0].Format["color"]).Should().BeOneOf("#FF0000", "FF0000");
        node.Children[0].Format["trendline"].Should().Be("poly",
            "per-series trendline type should persist (stored as 'poly')");
        ((string)node.Children[1].Format["color"]).Should().BeOneOf("#0000FF", "0000FF");
    }

    [Fact]
    public void Excel_Reopen_DispUnits_Persists()
    {
        var chartPath = AddExcelChart();
        _excel.Set(chartPath, new() { ["dispUnits"] = "thousands" });

        ReopenExcel();
        var node = _excel.Get(chartPath, depth: 0);
        node.Format["dispUnits"].Should().NotBeNull("dispUnits should persist after reopen");
    }
}
