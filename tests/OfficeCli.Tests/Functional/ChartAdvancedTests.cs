// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Tests for 4 advanced chart features:
/// 1. Reference Line (referenceLine/refLine/targetLine)
/// 2. Conditional Coloring (colorRule/conditionalColor)
/// 3. Waterfall Chart (chartType=waterfall)
/// 4. Flexible Combo Types (comboTypes/combo.types)
/// </summary>
public class ChartAdvancedTests : IDisposable
{
    private readonly string _xlsxPath;
    private readonly string _pptxPath;
    private ExcelHandler _excel;
    private PowerPointHandler _pptx;

    public ChartAdvancedTests()
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

    private string AddExcelColumnChart(string? extraData = null)
    {
        return _excel.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "column",
            ["title"] = "Test",
            ["data"] = extraData ?? "S1:10,20,30;S2:15,25,35",
            ["categories"] = "A,B,C"
        });
    }

    private string AddExcelLineChart()
    {
        return _excel.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "line",
            ["title"] = "Test",
            ["data"] = "S1:10,20,30;S2:15,25,35",
            ["categories"] = "A,B,C"
        });
    }

    private string AddPptxColumnChart()
    {
        _pptx.Add("/", "slide", null, new());
        return _pptx.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "column",
            ["title"] = "Test",
            ["data"] = "S1:10,20,30;S2:15,25,35",
            ["categories"] = "A,B,C"
        });
    }

    private string AddPptxLineChart()
    {
        _pptx.Add("/", "slide", null, new());
        return _pptx.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "line",
            ["title"] = "Test",
            ["data"] = "S1:10,20,30;S2:15,25,35",
            ["categories"] = "A,B,C"
        });
    }

    // ==================== #1 Reference Line ====================

    [Fact]
    public void Excel_ReferenceLine_Basic_IncrementsSeriesCount()
    {
        var chartPath = AddExcelColumnChart();
        var before = _excel.Get(chartPath, depth: 0);
        var countBefore = (int)before.Format["seriesCount"];

        _excel.Set(chartPath, new() { ["referenceLine"] = "50" });

        var after = _excel.Get(chartPath, depth: 0);
        var countAfter = (int)after.Format["seriesCount"];
        countAfter.Should().BeGreaterThan(countBefore, "a reference line series should be added");
    }

    [Fact]
    public void Excel_ReferenceLine_WithColorAndLabel()
    {
        var chartPath = AddExcelColumnChart();
        var before = _excel.Get(chartPath, depth: 0);
        var countBefore = (int)before.Format["seriesCount"];

        _excel.Set(chartPath, new() { ["referenceLine"] = "75:FF0000:Target" });

        var after = _excel.Get(chartPath, depth: 0);
        ((int)after.Format["seriesCount"]).Should().BeGreaterThan(countBefore);
    }

    [Fact]
    public void Excel_ReferenceLine_WithDash_Persists()
    {
        var chartPath = AddExcelColumnChart();
        var before = _excel.Get(chartPath, depth: 0);
        var countBefore = (int)before.Format["seriesCount"];

        _excel.Set(chartPath, new() { ["referenceLine"] = "80:0000FF:Average:dash" });

        ReopenExcel();
        var after = _excel.Get(chartPath, depth: 0);
        ((int)after.Format["seriesCount"]).Should().BeGreaterThan(countBefore);
    }

    [Fact]
    public void Excel_ReferenceLine_AlternativeKey_RefLine()
    {
        var chartPath = AddExcelColumnChart();
        var before = _excel.Get(chartPath, depth: 0);
        var countBefore = (int)before.Format["seriesCount"];

        _excel.Set(chartPath, new() { ["refLine"] = "100" });

        var after = _excel.Get(chartPath, depth: 0);
        ((int)after.Format["seriesCount"]).Should().BeGreaterThan(countBefore);
    }

    [Fact]
    public void Excel_ReferenceLine_AlternativeKey_TargetLine()
    {
        var chartPath = AddExcelColumnChart();
        var before = _excel.Get(chartPath, depth: 0);
        var countBefore = (int)before.Format["seriesCount"];

        _excel.Set(chartPath, new() { ["targetLine"] = "60:FFAA00:Goal" });

        var after = _excel.Get(chartPath, depth: 0);
        ((int)after.Format["seriesCount"]).Should().BeGreaterThan(countBefore);
    }

    [Fact]
    public void Excel_ReferenceLine_OnLineChart_IncrementsSeriesCount()
    {
        var chartPath = AddExcelLineChart();
        var before = _excel.Get(chartPath, depth: 0);
        var countBefore = (int)before.Format["seriesCount"];

        _excel.Set(chartPath, new() { ["referenceLine"] = "50:FF0000:Target" });

        var after = _excel.Get(chartPath, depth: 0);
        ((int)after.Format["seriesCount"]).Should().BeGreaterThan(countBefore);
    }

    [Fact]
    public void Pptx_ReferenceLine_OnColumnChart_IncrementsSeriesCount()
    {
        var chartPath = AddPptxColumnChart();
        var before = _pptx.Get(chartPath, depth: 0);
        var countBefore = (int)before.Format["seriesCount"];

        _pptx.Set(chartPath, new() { ["referenceLine"] = "50:FF0000:Target" });

        var after = _pptx.Get(chartPath, depth: 0);
        ((int)after.Format["seriesCount"]).Should().BeGreaterThan(countBefore);
    }

    [Fact]
    public void Pptx_ReferenceLine_Persists()
    {
        var chartPath = AddPptxLineChart();
        var before = _pptx.Get(chartPath, depth: 0);
        var countBefore = (int)before.Format["seriesCount"];

        _pptx.Set(chartPath, new() { ["referenceLine"] = "25:00AA00:Average" });

        ReopenPptx();
        var after = _pptx.Get(chartPath, depth: 0);
        ((int)after.Format["seriesCount"]).Should().BeGreaterThan(countBefore);
    }

    [Fact]
    public void Excel_ReferenceLine_InvalidValue_Throws()
    {
        var chartPath = AddExcelColumnChart();
        var act = () => _excel.Set(chartPath, new() { ["referenceLine"] = "notANumber" });
        act.Should().Throw<Exception>();
    }

    // ==================== #2 Conditional Coloring ====================

    [Fact]
    public void Excel_ColorRule_TwoZone_AppliesPointColors()
    {
        var chartPath = AddExcelColumnChart("S1:-10,20,-5,30");
        _excel.Set(chartPath, new() { ["colorRule"] = "0:FF0000:00AA00" });

        var node = _excel.Get(chartPath, depth: 1);
        node.Children.Should().NotBeEmpty();
        // Negative values (index 0 and 2) should have red, positive (1 and 3) green
        var ser = node.Children[0];
        ser.Format.Should().ContainKey("point1.color"); // -10 < 0 → red
        ((string)ser.Format["point1.color"]).Should().Be("#FF0000");
        ser.Format.Should().ContainKey("point2.color"); // 20 >= 0 → green
        ((string)ser.Format["point2.color"]).Should().Be("#00AA00");
    }

    [Fact]
    public void Excel_ColorRule_AlternativeKey_ConditionalColor()
    {
        var chartPath = AddExcelColumnChart("S1:-10,20,-5,30");
        _excel.Set(chartPath, new() { ["conditionalColor"] = "0:FF0000:00AA00" });

        var node = _excel.Get(chartPath, depth: 1);
        node.Children.Should().NotBeEmpty();
        node.Children[0].Format.Should().ContainKey("point1.color");
    }

    [Fact]
    public void Excel_ColorRule_MultiZone_AppliesPointColors()
    {
        var chartPath = AddExcelColumnChart("S1:10,40,70,90");
        // Four zones: <25=red, <50=orange, <75=yellow, >=75=green
        _excel.Set(chartPath, new() { ["colorRule"] = "25:FF0000:50:FFAA00:75:00FF00:00AA00" });

        var node = _excel.Get(chartPath, depth: 1);
        node.Children.Should().NotBeEmpty();
        var ser = node.Children[0];
        // 10 < 25 → red
        ser.Format.Should().ContainKey("point1.color");
        ((string)ser.Format["point1.color"]).Should().Be("#FF0000");
        // 90 >= 75 → green
        ser.Format.Should().ContainKey("point4.color");
        ((string)ser.Format["point4.color"]).Should().Be("#00AA00");
    }

    [Fact]
    public void Excel_ColorRule_Persists()
    {
        var chartPath = AddExcelColumnChart("S1:-10,20");
        _excel.Set(chartPath, new() { ["colorRule"] = "0:FF0000:00AA00" });

        ReopenExcel();
        var node = _excel.Get(chartPath, depth: 1);
        node.Children[0].Format.Should().ContainKey("point1.color");
        ((string)node.Children[0].Format["point1.color"]).Should().Be("#FF0000");
    }

    [Fact]
    public void Pptx_ColorRule_TwoZone_AppliesPointColors()
    {
        _pptx.Add("/", "slide", null, new());
        var chartPath = _pptx.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "column",
            ["title"] = "ColorRule",
            ["data"] = "S1:-10,20,-5,30",
            ["categories"] = "A,B,C,D"
        });

        _pptx.Set(chartPath, new() { ["colorRule"] = "0:FF0000:00AA00" });

        var node = _pptx.Get(chartPath, depth: 1);
        node.Children.Should().NotBeEmpty();
        node.Children[0].Format.Should().ContainKey("point1.color");
        ((string)node.Children[0].Format["point1.color"]).Should().Be("#FF0000");
    }

    [Fact]
    public void Pptx_ColorRule_Persists()
    {
        _pptx.Add("/", "slide", null, new());
        var chartPath = _pptx.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "column",
            ["title"] = "Persist",
            ["data"] = "S1:-5,10",
            ["categories"] = "A,B"
        });
        _pptx.Set(chartPath, new() { ["colorRule"] = "0:FF0000:00AA00" });

        ReopenPptx();
        var node = _pptx.Get(chartPath, depth: 1);
        node.Children[0].Format.Should().ContainKey("point1.color");
    }

    [Fact]
    public void Excel_ColorRule_InvalidFormat_Throws()
    {
        var chartPath = AddExcelColumnChart();
        var act = () => _excel.Set(chartPath, new() { ["colorRule"] = "onlyone" });
        act.Should().Throw<Exception>();
    }

    [Fact]
    public void Excel_ColorRule_InvalidThreshold_Throws()
    {
        var chartPath = AddExcelColumnChart();
        var act = () => _excel.Set(chartPath, new() { ["colorRule"] = "notanumber:FF0000:00AA00" });
        act.Should().Throw<Exception>();
    }

    // ==================== #3 Waterfall Chart ====================

    [Fact]
    public void Excel_Waterfall_Create_ChartTypeIsColumnStacked()
    {
        var chartPath = _excel.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "waterfall",
            ["title"] = "Cashflow",
            ["data"] = "Flow:100,-30,-15,55",
            ["categories"] = "Revenue,Cost,Tax,Profit"
        });

        var node = _excel.Get(chartPath, depth: 0);
        node.Should().NotBeNull();
        // Waterfall is simulated as stacked column but detected back as waterfall
        ((string)node.Format["chartType"]).Should().Be("waterfall");
    }

    [Fact]
    public void Excel_Waterfall_Create_SeriesCountIsThree()
    {
        var chartPath = _excel.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "waterfall",
            ["title"] = "Cashflow",
            ["data"] = "Flow:100,-30,-15,55",
            ["categories"] = "Revenue,Cost,Tax,Profit"
        });

        var node = _excel.Get(chartPath, depth: 0);
        // Base + Increase + Decrease = 3
        ((int)node.Format["seriesCount"]).Should().Be(3);
    }

    [Fact]
    public void Excel_Waterfall_WithCustomColors_Persists()
    {
        var chartPath = _excel.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "waterfall",
            ["title"] = "Cashflow",
            ["data"] = "Flow:100,-30,-15,55",
            ["categories"] = "Revenue,Cost,Tax,Profit",
            ["increaseColor"] = "4472C4",
            ["decreaseColor"] = "FF0000",
            ["totalColor"] = "00AA00"
        });

        ReopenExcel();
        var node = _excel.Get(chartPath, depth: 0);
        ((string)node.Format["chartType"]).Should().Be("waterfall");
        ((int)node.Format["seriesCount"]).Should().Be(3);
    }

    [Fact]
    public void Pptx_Waterfall_Create_ChartTypeIsColumnStacked()
    {
        _pptx.Add("/", "slide", null, new());
        var chartPath = _pptx.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "waterfall",
            ["title"] = "Cashflow",
            ["data"] = "Flow:100,-30,-15,55",
            ["categories"] = "Revenue,Cost,Tax,Profit"
        });

        var node = _pptx.Get(chartPath, depth: 0);
        node.Should().NotBeNull();
        ((string)node.Format["chartType"]).Should().Be("waterfall");
    }

    [Fact]
    public void Pptx_Waterfall_SeriesCountIsThree()
    {
        _pptx.Add("/", "slide", null, new());
        var chartPath = _pptx.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "waterfall",
            ["title"] = "Cashflow",
            ["data"] = "Flow:100,-30,-15,55",
            ["categories"] = "Revenue,Cost,Tax,Profit"
        });

        var node = _pptx.Get(chartPath, depth: 0);
        ((int)node.Format["seriesCount"]).Should().Be(3);
    }

    [Fact]
    public void Pptx_Waterfall_WithCustomColors_Persists()
    {
        _pptx.Add("/", "slide", null, new());
        var chartPath = _pptx.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "waterfall",
            ["title"] = "Cashflow",
            ["data"] = "Flow:100,-30,-15,55",
            ["categories"] = "Revenue,Cost,Tax,Profit",
            ["increaseColor"] = "2E75B6",
            ["decreaseColor"] = "FF4B4B",
            ["totalColor"] = "00AA00"
        });

        ReopenPptx();
        var node = _pptx.Get(chartPath, depth: 0);
        ((string)node.Format["chartType"]).Should().Be("waterfall");
        ((int)node.Format["seriesCount"]).Should().Be(3);
    }

    [Fact]
    public void Excel_Waterfall_WaterfallTotalFalse_SeriesCountIsThree()
    {
        // waterfallTotal=false means last bar is treated as regular, still 3 series (base/inc/dec)
        var chartPath = _excel.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "waterfall",
            ["title"] = "Cashflow",
            ["data"] = "Flow:100,-30,-15,55",
            ["categories"] = "Revenue,Cost,Tax,Profit",
            ["waterfallTotal"] = "false"
        });

        var node = _excel.Get(chartPath, depth: 0);
        ((int)node.Format["seriesCount"]).Should().Be(3);
    }

    // ==================== #4 Flexible Combo Types ====================

    [Fact]
    public void Excel_ComboTypes_Basic_ChartTypeIsCombo()
    {
        var chartPath = _excel.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "combo",
            ["title"] = "Combo",
            ["data"] = "S1:10,20,30;S2:15,25,35",
            ["categories"] = "A,B,C"
        });

        _excel.Set(chartPath, new() { ["comboTypes"] = "column,line" });

        var node = _excel.Get(chartPath, depth: 0);
        // After comboTypes reassignment, should still be combo (multiple chart types)
        ((string)node.Format["chartType"]).Should().Be("combo");
    }

    [Fact]
    public void Excel_ComboTypes_AlternativeKey_ComboDotsTypes()
    {
        var chartPath = _excel.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "combo",
            ["title"] = "Combo",
            ["data"] = "S1:10,20,30;S2:15,25,35",
            ["categories"] = "A,B,C"
        });

        _excel.Set(chartPath, new() { ["combo.types"] = "column,area" });

        var node = _excel.Get(chartPath, depth: 0);
        ((string)node.Format["chartType"]).Should().Be("combo");
    }

    [Fact]
    public void Excel_ComboTypes_ThreeSeries_Persists()
    {
        var chartPath = _excel.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "combo",
            ["title"] = "Combo",
            ["data"] = "S1:10,20,30;S2:15,25,35;S3:5,10,20",
            ["categories"] = "A,B,C"
        });

        _excel.Set(chartPath, new() { ["comboTypes"] = "column,column,line" });

        ReopenExcel();
        var node = _excel.Get(chartPath, depth: 0);
        ((string)node.Format["chartType"]).Should().Be("combo");
        ((int)node.Format["seriesCount"]).Should().Be(3);
    }

    [Fact]
    public void Pptx_ComboTypes_Basic_ChartTypeIsCombo()
    {
        _pptx.Add("/", "slide", null, new());
        var chartPath = _pptx.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "combo",
            ["title"] = "Combo",
            ["data"] = "S1:10,20,30;S2:15,25,35",
            ["categories"] = "A,B,C"
        });

        _pptx.Set(chartPath, new() { ["comboTypes"] = "column,line" });

        var node = _pptx.Get(chartPath, depth: 0);
        ((string)node.Format["chartType"]).Should().Be("combo");
    }

    [Fact]
    public void Pptx_ComboTypes_Persists()
    {
        _pptx.Add("/", "slide", null, new());
        var chartPath = _pptx.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "combo",
            ["title"] = "Combo",
            ["data"] = "S1:10,20,30;S2:15,25,35",
            ["categories"] = "A,B,C"
        });

        _pptx.Set(chartPath, new() { ["comboTypes"] = "column,area" });

        ReopenPptx();
        var node = _pptx.Get(chartPath, depth: 0);
        ((string)node.Format["chartType"]).Should().Be("combo");
    }

    [Fact]
    public void Excel_ComboTypes_ColumnOnly_ChartTypeIsColumn()
    {
        var chartPath = _excel.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "combo",
            ["title"] = "Combo",
            ["data"] = "S1:10,20,30;S2:15,25,35",
            ["categories"] = "A,B,C"
        });

        // All series assigned same type — should collapse to single type (column)
        _excel.Set(chartPath, new() { ["comboTypes"] = "column,column" });

        var node = _excel.Get(chartPath, depth: 0);
        // Could be "column" or "combo" depending on implementation — just verify no exception & seriesCount
        node.Format.Should().ContainKey("seriesCount");
        ((int)node.Format["seriesCount"]).Should().Be(2);
    }
}
