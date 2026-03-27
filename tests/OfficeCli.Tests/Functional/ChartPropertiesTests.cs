// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Functional tests for new chart properties: axis, chart-level, series-level,
/// data labels, borders, data table, chart-type-specific, and legend enhancements.
/// </summary>
public class ChartPropertiesTests : IDisposable
{
    private readonly string _xlsxPath;
    private ExcelHandler _excel;

    public ChartPropertiesTests()
    {
        _xlsxPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.xlsx");
        BlankDocCreator.Create(_xlsxPath);
        _excel = new ExcelHandler(_xlsxPath, editable: true);
    }

    public void Dispose()
    {
        _excel.Dispose();
        if (File.Exists(_xlsxPath)) File.Delete(_xlsxPath);
    }

    // ─── helpers ───────────────────────────────────────────────────────────

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

    // ==================== Axis: axisVisible ====================

    [Fact]
    public void Set_AxisVisible_False_HidesAxis()
    {
        var path = AddChart();
        _excel.Set(path, new() { ["axisVisible"] = "false" });

        var node = _excel.Get(path, depth: 0);
        node.Format["valAxisVisible"].Should().Be("false");
        node.Format["catAxisVisible"].Should().Be("false");
    }

    [Fact]
    public void Set_AxisVisible_True_ShowsAxis()
    {
        var path = AddChart();
        _excel.Set(path, new() { ["axisVisible"] = "false" });
        _excel.Set(path, new() { ["axisVisible"] = "true" });

        var node = _excel.Get(path, depth: 0);
        node.Format.Should().NotContainKey("valAxisVisible");
        node.Format.Should().NotContainKey("catAxisVisible");
    }

    // ==================== Axis: majorTickMark ====================

    [Fact]
    public void Set_MajorTickMark_Out_IsReadBack()
    {
        var path = AddChart();
        _excel.Set(path, new() { ["majorTickMark"] = "out" });

        var node = _excel.Get(path, depth: 0);
        node.Format["majorTickMark"].Should().Be("out");
    }

    [Fact]
    public void Set_MajorTickMark_Cross_IsReadBack()
    {
        var path = AddChart();
        _excel.Set(path, new() { ["majorTickMark"] = "cross" });

        var node = _excel.Get(path, depth: 0);
        node.Format["majorTickMark"].Should().Be("cross");
    }

    [Fact]
    public void Set_MajorTickMark_None_IsReadBack()
    {
        var path = AddChart();
        _excel.Set(path, new() { ["majorTickMark"] = "none" });

        var node = _excel.Get(path, depth: 0);
        node.Format["majorTickMark"].Should().Be("none");
    }

    // ==================== Axis: minorTickMark ====================

    [Fact]
    public void Set_MinorTickMark_In_IsReadBack()
    {
        var path = AddChart();
        _excel.Set(path, new() { ["minorTickMark"] = "in" });

        var node = _excel.Get(path, depth: 0);
        node.Format["minorTickMark"].Should().Be("in");
    }

    // ==================== Axis: tickLabelPos ====================

    [Fact]
    public void Set_TickLabelPos_High_IsReadBack()
    {
        var path = AddChart();
        _excel.Set(path, new() { ["tickLabelPos"] = "high" });

        var node = _excel.Get(path, depth: 0);
        node.Format["tickLabelPos"].Should().Be("high");
    }

    [Fact]
    public void Set_TickLabelPos_None_IsReadBack()
    {
        var path = AddChart();
        _excel.Set(path, new() { ["tickLabelPos"] = "none" });

        var node = _excel.Get(path, depth: 0);
        node.Format["tickLabelPos"].Should().Be("none");
    }

    [Fact]
    public void Set_TickLabelPos_NextTo_IsReadBack()
    {
        var path = AddChart();
        _excel.Set(path, new() { ["tickLabelPos"] = "nextTo" });

        var node = _excel.Get(path, depth: 0);
        node.Format["tickLabelPos"].Should().Be("nextTo");
    }

    // ==================== Axis: crosses ====================

    [Fact]
    public void Set_Crosses_Max_IsReadBack()
    {
        var path = AddChart();
        _excel.Set(path, new() { ["crosses"] = "max" });

        var node = _excel.Get(path, depth: 0);
        node.Format["crosses"].Should().Be("max");
    }

    [Fact]
    public void Set_Crosses_Min_IsReadBack()
    {
        var path = AddChart();
        _excel.Set(path, new() { ["crosses"] = "min" });

        var node = _excel.Get(path, depth: 0);
        node.Format["crosses"].Should().Be("min");
    }

    [Fact]
    public void Set_Crosses_AutoZero_IsReadBack()
    {
        var path = AddChart();
        _excel.Set(path, new() { ["crosses"] = "autoZero" });

        var node = _excel.Get(path, depth: 0);
        node.Format["crosses"].Should().Be("autoZero");
    }

    // ==================== Axis: crossBetween ====================

    [Fact]
    public void Set_CrossBetween_Between_IsReadBack()
    {
        var path = AddChart();
        _excel.Set(path, new() { ["crossBetween"] = "between" });

        var node = _excel.Get(path, depth: 0);
        node.Format["crossBetween"].Should().Be("between");
    }

    [Fact]
    public void Set_CrossBetween_MidCat_IsReadBack()
    {
        var path = AddChart();
        _excel.Set(path, new() { ["crossBetween"] = "midcat" });

        var node = _excel.Get(path, depth: 0);
        node.Format["crossBetween"].Should().Be("midCat");
    }

    // ==================== Axis: axisOrientation ====================

    [Fact]
    public void Set_AxisOrientation_True_ReturnsMaxMin()
    {
        var path = AddChart();
        _excel.Set(path, new() { ["axisOrientation"] = "true" });

        var node = _excel.Get(path, depth: 0);
        node.Format["axisOrientation"].Should().Be("maxMin");
    }

    [Fact]
    public void Set_AxisOrientation_MaxMin_String_IsReadBack()
    {
        var path = AddChart();
        _excel.Set(path, new() { ["axisOrientation"] = "maxMin" });

        var node = _excel.Get(path, depth: 0);
        node.Format["axisOrientation"].Should().Be("maxMin");
    }

    // ==================== Axis: logBase ====================

    [Fact]
    public void Set_LogBase_10_IsReadBack()
    {
        var path = AddChart();
        _excel.Set(path, new() { ["logBase"] = "10" });

        var node = _excel.Get(path, depth: 0);
        node.Format["logBase"].Should().Be(10.0);
    }

    // ==================== Axis: dispUnits ====================

    [Fact]
    public void Set_DispUnits_Thousands_IsReadBack()
    {
        var path = AddChart();
        _excel.Set(path, new() { ["dispUnits"] = "thousands" });

        var node = _excel.Get(path, depth: 0);
        node.Format["dispUnits"].Should().Be("thousands");
    }

    [Fact]
    public void Set_DispUnits_Millions_IsReadBack()
    {
        var path = AddChart();
        _excel.Set(path, new() { ["dispUnits"] = "millions" });

        var node = _excel.Get(path, depth: 0);
        node.Format["dispUnits"].Should().Be("millions");
    }

    [Fact]
    public void Set_DispUnits_Billions_IsReadBack()
    {
        var path = AddChart();
        _excel.Set(path, new() { ["dispUnits"] = "billions" });

        var node = _excel.Get(path, depth: 0);
        node.Format["dispUnits"].Should().Be("billions");
    }

    // ==================== Axis: labelOffset ====================

    [Fact]
    public void Set_LabelOffset_200_IsReadBack()
    {
        var path = AddChart();
        _excel.Set(path, new() { ["labelOffset"] = "200" });

        var node = _excel.Get(path, depth: 0);
        node.Format["labelOffset"].Should().Be((ushort)200);
    }

    // ==================== Axis: tickLabelSkip ====================

    [Fact]
    public void Set_TickLabelSkip_2_IsReadBack()
    {
        var path = AddChart();
        _excel.Set(path, new() { ["tickLabelSkip"] = "2" });

        var node = _excel.Get(path, depth: 0);
        node.Format["tickLabelSkip"].Should().Be(2);
    }

    // ==================== Chart-level: smooth ====================

    [Fact]
    public void Set_Smooth_True_IsReadBack_LineChart()
    {
        var path = AddChart("line");
        _excel.Set(path, new() { ["smooth"] = "true" });

        var node = _excel.Get(path, depth: 0);
        node.Format["smooth"].Should().Be("true");
    }

    [Fact]
    public void Set_Smooth_False_IsReadBack_LineChart()
    {
        var path = AddChart("line");
        _excel.Set(path, new() { ["smooth"] = "false" });

        var node = _excel.Get(path, depth: 0);
        node.Format["smooth"].Should().Be("false");
    }

    // ==================== Chart-level: showMarker ====================

    [Fact]
    public void Set_ShowMarker_True_IsReadBack()
    {
        var path = AddChart("line");
        _excel.Set(path, new() { ["showMarker"] = "true" });

        var node = _excel.Get(path, depth: 0);
        node.Format["showMarker"].Should().Be("true");
    }

    [Fact]
    public void Set_ShowMarker_False_IsReadBack()
    {
        var path = AddChart("line");
        _excel.Set(path, new() { ["showMarker"] = "false" });

        var node = _excel.Get(path, depth: 0);
        node.Format["showMarker"].Should().Be("false");
    }

    // ==================== Chart-level: scatterStyle ====================

    [Fact]
    public void Set_ScatterStyle_Line_IsReadBack()
    {
        var path = AddChart("scatter");
        _excel.Set(path, new() { ["scatterStyle"] = "line" });

        var node = _excel.Get(path, depth: 0);
        node.Format["scatterStyle"].Should().Be("line");
    }

    [Fact]
    public void Set_ScatterStyle_LineMarker_IsReadBack()
    {
        var path = AddChart("scatter");
        _excel.Set(path, new() { ["scatterStyle"] = "lineMarker" });

        var node = _excel.Get(path, depth: 0);
        node.Format["scatterStyle"].Should().Be("lineMarker");
    }

    [Fact]
    public void Set_ScatterStyle_SmoothMarker_IsReadBack()
    {
        var path = AddChart("scatter");
        _excel.Set(path, new() { ["scatterStyle"] = "smoothMarker" });

        var node = _excel.Get(path, depth: 0);
        node.Format["scatterStyle"].Should().Be("smoothMarker");
    }

    // ==================== Chart-level: varyColors ====================

    [Fact]
    public void Set_VaryColors_True_IsReadBack()
    {
        var path = AddChart();
        _excel.Set(path, new() { ["varyColors"] = "true" });

        var node = _excel.Get(path, depth: 1);
        // varyColors affects series rendering; verify no exception and chart is intact
        node.Should().NotBeNull();
    }

    // ==================== Chart-level: dispBlanksAs ====================

    [Fact]
    public void Set_DispBlanksAs_Zero_IsReadBack()
    {
        var path = AddChart();
        _excel.Set(path, new() { ["dispBlanksAs"] = "zero" });

        var node = _excel.Get(path, depth: 0);
        node.Format["dispBlanksAs"].Should().Be("zero");
    }

    [Fact]
    public void Set_DispBlanksAs_Span_IsReadBack()
    {
        var path = AddChart();
        _excel.Set(path, new() { ["dispBlanksAs"] = "span" });

        var node = _excel.Get(path, depth: 0);
        node.Format["dispBlanksAs"].Should().Be("span");
    }

    [Fact]
    public void Set_DispBlanksAs_Gap_IsReadBack()
    {
        var path = AddChart();
        _excel.Set(path, new() { ["dispBlanksAs"] = "gap" });

        var node = _excel.Get(path, depth: 0);
        node.Format["dispBlanksAs"].Should().Be("gap");
    }

    // ==================== Chart-level: roundedCorners ====================

    [Fact]
    public void Set_RoundedCorners_True_IsReadBack()
    {
        var path = AddChart();
        _excel.Set(path, new() { ["roundedCorners"] = "true" });

        var node = _excel.Get(path, depth: 0);
        node.Format["roundedCorners"].Should().Be("true");
    }

    [Fact]
    public void Set_RoundedCorners_False_IsReadBack()
    {
        var path = AddChart();
        _excel.Set(path, new() { ["roundedCorners"] = "true" });
        _excel.Set(path, new() { ["roundedCorners"] = "false" });

        var node = _excel.Get(path, depth: 0);
        node.Format["roundedCorners"].Should().Be("false");
    }

    // ==================== Chart-level: plotVisOnly ====================

    [Fact]
    public void Set_PlotVisOnly_True_IsSet()
    {
        var path = AddChart();
        // plotVisOnly is set-only (no readback key defined), verify no exception
        var act = () => _excel.Set(path, new() { ["plotVisOnly"] = "true" });
        act.Should().NotThrow();
    }

    // ==================== Series: series1.smooth ====================

    [Fact]
    public void Set_Series1_Smooth_True_IsReadBack()
    {
        var path = AddChart("line");
        _excel.Set(path, new() { ["series1.smooth"] = "true" });

        var node = _excel.Get(path, depth: 1);
        node.Children[0].Format["smooth"].Should().Be("true");
    }

    [Fact]
    public void Set_Series1_Smooth_False_IsReadBack()
    {
        var path = AddChart("line");
        _excel.Set(path, new() { ["series1.smooth"] = "false" });

        var node = _excel.Get(path, depth: 1);
        node.Children[0].Format["smooth"].Should().Be("false");
    }

    // ==================== Series: series1.trendline ====================

    [Fact]
    public void Set_Series1_Trendline_Linear_IsReadBack()
    {
        var path = AddChart("line");
        _excel.Set(path, new() { ["series1.trendline"] = "linear" });

        var node = _excel.Get(path, depth: 1);
        node.Children[0].Format["trendline"].Should().Be("linear");
    }

    [Fact]
    public void Set_Series1_Trendline_Exp_IsReadBack()
    {
        var path = AddChart("line");
        _excel.Set(path, new() { ["series1.trendline"] = "exp" });

        var node = _excel.Get(path, depth: 1);
        node.Children[0].Format["trendline"].Should().Be("exp");
    }

    [Fact]
    public void Set_Series1_Trendline_Log_IsReadBack()
    {
        var path = AddChart("line");
        _excel.Set(path, new() { ["series1.trendline"] = "log" });

        var node = _excel.Get(path, depth: 1);
        node.Children[0].Format["trendline"].Should().Be("log");
    }

    // ==================== Series: series1.trendline.dispRSqr ====================

    [Fact]
    public void Set_Series1_Trendline_DispRSqr_True_IsReadBack()
    {
        var path = AddChart("line");
        _excel.Set(path, new() { ["series1.trendline"] = "linear", ["series1.trendline.dispRSqr"] = "true" });

        var node = _excel.Get(path, depth: 1);
        node.Children[0].Format["trendline.dispRSqr"].Should().Be("true");
    }

    // ==================== Series: series1.trendline.dispEq ====================

    [Fact]
    public void Set_Series1_Trendline_DispEq_True_IsReadBack()
    {
        var path = AddChart("line");
        _excel.Set(path, new() { ["series1.trendline"] = "linear", ["series1.trendline.dispEq"] = "true" });

        var node = _excel.Get(path, depth: 1);
        node.Children[0].Format["trendline.dispEq"].Should().Be("true");
    }

    // ==================== Series: series1.point1.color ====================

    [Fact]
    public void Set_Series1_Point1_Color_IsReadBack()
    {
        var path = AddChart();
        _excel.Set(path, new() { ["series1.point1.color"] = "#FF0000" });

        var node = _excel.Get(path, depth: 1);
        node.Children[0].Format["point1.color"].Should().Be("#FF0000");
    }

    [Fact]
    public void Set_Series1_Point2_Color_WithoutHash_IsReadBack()
    {
        var path = AddChart();
        _excel.Set(path, new() { ["series1.point2.color"] = "0000FF" });

        var node = _excel.Get(path, depth: 1);
        node.Children[0].Format["point2.color"].Should().Be("#0000FF");
    }

    // ==================== Series: series1.errBars ====================

    [Fact]
    public void Set_Series1_ErrBars_Fixed_IsReadBack()
    {
        var path = AddChart();
        _excel.Set(path, new() { ["series1.errBars"] = "fixed:5" });

        var node = _excel.Get(path, depth: 1);
        node.Children[0].Format["errBars"].Should().NotBeNull();
    }

    [Fact]
    public void Set_Series1_ErrBars_Percent_IsReadBack()
    {
        var path = AddChart();
        _excel.Set(path, new() { ["series1.errBars"] = "percent:10" });

        var node = _excel.Get(path, depth: 1);
        node.Children[0].Format["errBars"].Should().NotBeNull();
    }

    // ==================== Series: series1.invertIfNeg ====================

    [Fact]
    public void Set_Series1_InvertIfNeg_True_IsReadBack()
    {
        var path = AddChart();
        _excel.Set(path, new() { ["series1.invertIfNeg"] = "true" });

        var node = _excel.Get(path, depth: 1);
        node.Children[0].Format["invertIfNeg"].Should().Be("true");
    }

    // ==================== Series: series1.explosion ====================

    [Fact]
    public void Set_Series1_Explosion_25_IsReadBack()
    {
        var path = AddChart("pie", new() { ["data"] = "S1:10,20,30", ["categories"] = "A,B,C" });
        _excel.Set(path, new() { ["series1.explosion"] = "25" });

        var node = _excel.Get(path, depth: 1);
        node.Children[0].Format["explosion"].Should().Be((uint)25);
    }

    // ==================== DataLabels: separator ====================

    [Fact]
    public void Set_DataLabels_Separator_Comma_IsReadBack()
    {
        var path = AddChart();
        _excel.Set(path, new() { ["dataLabels"] = "value,category", ["dataLabels.separator"] = ", " });

        var node = _excel.Get(path, depth: 0);
        node.Format["dataLabels.separator"].Should().Be(", ");
    }

    [Fact]
    public void Set_DataLabels_Separator_Newline_IsReadBack()
    {
        var path = AddChart();
        _excel.Set(path, new() { ["dataLabels"] = "value,category", ["dataLabels.separator"] = "\\n" });

        var node = _excel.Get(path, depth: 0);
        node.Format["dataLabels.separator"].Should().Be("\n");
    }

    // ==================== DataLabels: numFmt ====================

    [Fact]
    public void Set_DataLabels_NumFmt_IsReadBack()
    {
        var path = AddChart();
        _excel.Set(path, new() { ["dataLabels"] = "value", ["dataLabels.numFmt"] = "0.00" });

        var node = _excel.Get(path, depth: 0);
        node.Format["dataLabels.numFmt"].Should().Be("0.00");
    }

    // ==================== DataLabels: leaderLines ====================

    [Fact]
    public void Set_LeaderLines_True_NoException()
    {
        var path = AddChart();
        _excel.Set(path, new() { ["dataLabels"] = "value", ["leaderLines"] = "true" });

        var node = _excel.Get(path, depth: 0);
        node.Should().NotBeNull();
    }

    // ==================== Border: plotArea.border ====================

    [Fact]
    public void Set_PlotAreaBorder_IsReadBack()
    {
        var path = AddChart();
        _excel.Set(path, new() { ["plotArea.border"] = "000000:1" });

        var node = _excel.Get(path, depth: 0);
        node.Format.Should().ContainKey("plotArea.border.color");
        node.Format["plotArea.border.color"].Should().Be("#000000");
        node.Format["plotArea.border.width"].Should().Be(1.0);
    }

    // ==================== Border: chartArea.border ====================

    [Fact]
    public void Set_ChartAreaBorder_IsReadBack()
    {
        var path = AddChart();
        _excel.Set(path, new() { ["chartArea.border"] = "333333:0.5" });

        var node = _excel.Get(path, depth: 0);
        node.Format.Should().ContainKey("chartArea.border.color");
        node.Format["chartArea.border.color"].Should().Be("#333333");
        node.Format["chartArea.border.width"].Should().Be(0.5);
    }

    // ==================== Data table ====================

    [Fact]
    public void Set_DataTable_True_IsReadBack()
    {
        var path = AddChart();
        _excel.Set(path, new() { ["dataTable"] = "true" });

        var node = _excel.Get(path, depth: 0);
        node.Format["dataTable"].Should().Be("true");
    }

    [Fact]
    public void Set_DataTable_False_Removes_DataTable()
    {
        var path = AddChart();
        _excel.Set(path, new() { ["dataTable"] = "true" });
        _excel.Set(path, new() { ["dataTable"] = "false" });

        var node = _excel.Get(path, depth: 0);
        node.Format.Should().NotContainKey("dataTable");
    }

    // ==================== firstSliceAngle (pie) ====================

    [Fact]
    public void Set_FirstSliceAngle_90_IsReadBack()
    {
        var path = AddChart("pie", new() { ["data"] = "S1:10,20,30", ["categories"] = "A,B,C" });
        _excel.Set(path, new() { ["firstSliceAngle"] = "90" });

        var node = _excel.Get(path, depth: 0);
        node.Format["firstSliceAngle"].Should().Be((ushort)90);
    }

    [Fact]
    public void Set_FirstSliceAngle_270_IsReadBack()
    {
        var path = AddChart("pie", new() { ["data"] = "S1:10,20,30", ["categories"] = "A,B,C" });
        _excel.Set(path, new() { ["firstSliceAngle"] = "270" });

        var node = _excel.Get(path, depth: 0);
        node.Format["firstSliceAngle"].Should().Be((ushort)270);
    }

    // ==================== holeSize (doughnut) ====================

    [Fact]
    public void Set_HoleSize_50_IsReadBack()
    {
        var path = AddChart("doughnut", new() { ["data"] = "S1:10,20,30", ["categories"] = "A,B,C" });
        _excel.Set(path, new() { ["holeSize"] = "50" });

        var node = _excel.Get(path, depth: 0);
        node.Format["holeSize"].Should().Be((byte)50);
    }

    [Fact]
    public void Set_HoleSize_75_IsReadBack()
    {
        var path = AddChart("doughnut", new() { ["data"] = "S1:10,20,30", ["categories"] = "A,B,C" });
        _excel.Set(path, new() { ["holeSize"] = "75" });

        var node = _excel.Get(path, depth: 0);
        node.Format["holeSize"].Should().Be((byte)75);
    }

    // ==================== radarStyle ====================

    [Fact]
    public void Set_RadarStyle_Filled_IsReadBack()
    {
        var path = AddChart("radar");
        _excel.Set(path, new() { ["radarStyle"] = "filled" });

        var node = _excel.Get(path, depth: 0);
        node.Format["radarStyle"].Should().Be("filled");
    }

    [Fact]
    public void Set_RadarStyle_Marker_IsReadBack()
    {
        var path = AddChart("radar");
        _excel.Set(path, new() { ["radarStyle"] = "marker" });

        var node = _excel.Get(path, depth: 0);
        node.Format["radarStyle"].Should().Be("marker");
    }

    [Fact]
    public void Set_RadarStyle_Standard_IsReadBack()
    {
        var path = AddChart("radar");
        _excel.Set(path, new() { ["radarStyle"] = "standard" });

        var node = _excel.Get(path, depth: 0);
        node.Format["radarStyle"].Should().Be("standard");
    }

    // ==================== legend.overlay ====================

    [Fact]
    public void Set_LegendOverlay_True_IsReadBack()
    {
        var path = AddChart();
        _excel.Set(path, new() { ["legend.overlay"] = "true" });

        var node = _excel.Get(path, depth: 0);
        node.Format["legend.overlay"].Should().Be("true");
    }

    [Fact]
    public void Set_LegendOverlay_False_IsNotPresent()
    {
        var path = AddChart();
        _excel.Set(path, new() { ["legend.overlay"] = "true" });
        _excel.Set(path, new() { ["legend.overlay"] = "false" });

        var node = _excel.Get(path, depth: 0);
        // When overlay=false, reader only writes the key when val is true
        node.Format.Should().NotContainKey("legend.overlay");
    }

    // ==================== legendEntry1.delete ====================

    [Fact]
    public void Set_LegendEntry1_Delete_True_IsHandled()
    {
        var path = AddChart();
        // Should not throw; legend must exist
        var act = () => _excel.Set(path, new() { ["legendEntry1.delete"] = "true" });
        act.Should().NotThrow();
    }
}
