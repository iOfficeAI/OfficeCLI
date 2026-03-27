// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Readback gap tests: properties that are write-only (Set succeeds but Get doesn't return them).
///
/// Covers the highest-impact gaps in ChartReader.cs:
///   - title.font, title.size, title.color, title.bold
///   - gridlines detail (color:width:dash)
///   - chartFill (chart area fill)
///   - plotFill detail
///   - gapwidth / overlap
///   - legendFont, axisFont (label font)
///   - series.shadow, series.outline
///
/// Also covers the paragraph alignment normalization gap in PowerPointHandler.Query.cs:
///   DrawingML stores "r" for right, "l" for left, "just" for justify —
///   Get must normalize these to human-readable values.
///
/// All tests follow the lifecycle:
///   Create blank file → Add chart → Set property → Get → Assert readback present
/// </summary>
public class PptxChartReadbackTests : IDisposable
{
    private readonly string _pptxPath;
    private PowerPointHandler _pptx;

    public PptxChartReadbackTests()
    {
        _pptxPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.pptx");
        BlankDocCreator.Create(_pptxPath);
        _pptx = new PowerPointHandler(_pptxPath, editable: true);
    }

    public void Dispose()
    {
        _pptx.Dispose();
        if (File.Exists(_pptxPath)) File.Delete(_pptxPath);
    }

    private void Reopen()
    {
        _pptx.Dispose();
        _pptx = new PowerPointHandler(_pptxPath, editable: true);
    }

    private string AddBarChart(string title = "Sales")
    {
        _pptx.Add("/", "slide", null, new() { ["title"] = "S" });
        return _pptx.Add("/slide[1]", "chart", null, new()
        {
            ["chartType"] = "column",
            ["title"] = title,
            ["data"] = "Q1:10,20,30;Q2:15,25,35",
            ["categories"] = "Jan,Feb,Mar"
        });
    }

    // ==================== Title Formatting Readback ====================

    /// <summary>
    /// BUG: title.font is write-only — Set("title.font", "Arial") does not appear in Get.
    /// ChartReader reads titleText via Descendants<Drawing.Text>() but never reads RunProperties.
    /// </summary>
    [Fact]
    public void Chart_TitleFont_IsReadBack()
    {
        AddBarChart();
        _pptx.Set("/slide[1]/chart[1]", new() { ["title.font"] = "Arial" });

        var node = _pptx.Get("/slide[1]/chart[1]");
        node.Format.Should().ContainKey("title.font");
        ((string)node.Format["title.font"]).Should().Be("Arial");
    }

    /// <summary>
    /// BUG: title.size is write-only — Set("title.size", "18") does not appear in Get.
    /// </summary>
    [Fact]
    public void Chart_TitleSize_IsReadBack()
    {
        AddBarChart();
        _pptx.Set("/slide[1]/chart[1]", new() { ["title.size"] = "18" });

        var node = _pptx.Get("/slide[1]/chart[1]");
        node.Format.Should().ContainKey("title.size");
        ((string)node.Format["title.size"]).Should().Be("18pt");
    }

    /// <summary>
    /// BUG: title.color is write-only — Set("title.color", "FF0000") does not appear in Get.
    /// </summary>
    [Fact]
    public void Chart_TitleColor_IsReadBack()
    {
        AddBarChart();
        _pptx.Set("/slide[1]/chart[1]", new() { ["title.color"] = "FF0000" });

        var node = _pptx.Get("/slide[1]/chart[1]");
        node.Format.Should().ContainKey("title.color");
        ((string)node.Format["title.color"]).Should().Be("#FF0000");
    }

    /// <summary>
    /// BUG: title.bold is write-only — Set("title.bold", "true") does not appear in Get.
    /// </summary>
    [Fact]
    public void Chart_TitleBold_IsReadBack()
    {
        AddBarChart();
        _pptx.Set("/slide[1]/chart[1]", new() { ["title.bold"] = "true" });

        var node = _pptx.Get("/slide[1]/chart[1]");
        node.Format.Should().ContainKey("title.bold");
        ((string)node.Format["title.bold"]).Should().Be("true");
    }

    // ==================== Title Formatting — Persistence After Reopen ====================

    [Fact]
    public void Chart_TitleFormatting_PersistsAfterReopen()
    {
        AddBarChart();
        _pptx.Set("/slide[1]/chart[1]", new()
        {
            ["title.font"] = "Calibri",
            ["title.size"] = "14",
            ["title.color"] = "4472C4",
            ["title.bold"] = "true"
        });

        Reopen();

        var node = _pptx.Get("/slide[1]/chart[1]");
        node.Format.Should().ContainKey("title.font");
        ((string)node.Format["title.font"]).Should().Be("Calibri");
        node.Format.Should().ContainKey("title.size");
        ((string)node.Format["title.size"]).Should().Be("14pt");
        node.Format.Should().ContainKey("title.color");
        ((string)node.Format["title.color"]).Should().Be("#4472C4");
        node.Format.Should().ContainKey("title.bold");
        ((string)node.Format["title.bold"]).Should().Be("true");
    }

    // ==================== Gridlines Detail Readback ====================

    /// <summary>
    /// BUG: gridlines returns only "true" when present — Set("gridlines", "CCCCCC:0.5:dash") writes
    /// color/width/dash to the spPr of the MajorGridlines element, but Get returns just "true".
    /// </summary>
    [Fact]
    public void Chart_GridlinesWithDetail_ColorIsReadBack()
    {
        AddBarChart();
        // Format: "COLOR:WIDTH:DASH"
        _pptx.Set("/slide[1]/chart[1]", new() { ["gridlines"] = "CCCCCC:0.5:dash" });

        var node = _pptx.Get("/slide[1]/chart[1]");
        // Current behavior: Format["gridlines"] == "true" (color/width/dash lost)
        // Expected behavior: gridlines detail accessible, e.g. Format["gridlines"] contains color info
        node.Format.Should().ContainKey("gridlines");
        node.Format["gridlines"]?.ToString().Should().Be("true");
        // Detail accessible via separate keys
        node.Format.Should().ContainKey("gridlineColor");
        node.Format["gridlineColor"]?.ToString().Should().Be("#CCCCCC");
    }

    /// <summary>
    /// Even a simple "true" gridlines should round-trip.
    /// </summary>
    [Fact]
    public void Chart_GridlinesSimple_IsReadBack()
    {
        AddBarChart();
        _pptx.Set("/slide[1]/chart[1]", new() { ["gridlines"] = "true" });

        var node = _pptx.Get("/slide[1]/chart[1]");
        node.Format.Should().ContainKey("gridlines");
    }

    [Fact]
    public void Chart_MinorGridlines_IsReadBack()
    {
        AddBarChart();
        _pptx.Set("/slide[1]/chart[1]", new() { ["minorGridlines"] = "true" });

        var node = _pptx.Get("/slide[1]/chart[1]");
        node.Format.Should().ContainKey("minorGridlines");
    }

    // ==================== Chart Area Fill Readback ====================

    /// <summary>
    /// BUG: chartFill is write-only — Set("chartFill", "1F2937") writes to ChartShapeProperties
    /// on chartSpace, but ChartReader only reads plotArea's ShapeProperties (plotFill), never chartSpace.
    /// </summary>
    [Fact]
    public void Chart_ChartFill_IsReadBack()
    {
        AddBarChart();
        _pptx.Set("/slide[1]/chart[1]", new() { ["chartFill"] = "1F2937" });

        var node = _pptx.Get("/slide[1]/chart[1]");
        node.Format.Should().ContainKey("chartFill");
        ((string)node.Format["chartFill"]).Should().Be("#1F2937");
    }

    [Fact]
    public void Chart_ChartFill_PersistsAfterReopen()
    {
        AddBarChart();
        _pptx.Set("/slide[1]/chart[1]", new() { ["chartFill"] = "0D1117" });

        Reopen();

        var node = _pptx.Get("/slide[1]/chart[1]");
        node.Format.Should().ContainKey("chartFill");
        ((string)node.Format["chartFill"]).Should().Be("#0D1117");
    }

    // ==================== Plot Area Fill Readback ====================

    [Fact]
    public void Chart_PlotFill_IsReadBack()
    {
        AddBarChart();
        _pptx.Set("/slide[1]/chart[1]", new() { ["plotFill"] = "E5E7EB" });

        var node = _pptx.Get("/slide[1]/chart[1]");
        node.Format.Should().ContainKey("plotFill");
        ((string)node.Format["plotFill"]).Should().Be("#E5E7EB");
    }

    [Fact]
    public void Chart_PlotFill_PersistsAfterReopen()
    {
        AddBarChart();
        _pptx.Set("/slide[1]/chart[1]", new() { ["plotFill"] = "F3F4F6" });

        Reopen();

        var node = _pptx.Get("/slide[1]/chart[1]");
        node.Format.Should().ContainKey("plotFill");
        ((string)node.Format["plotFill"]).Should().Be("#F3F4F6");
    }

    // ==================== GapWidth / Overlap Readback ====================

    /// <summary>
    /// BUG: gapwidth is write-only — Set("gapwidth", "150") sets GapWidth on bar chart element,
    /// but ChartReader never reads GapWidth back into Format.
    /// </summary>
    [Fact]
    public void Chart_GapWidth_IsReadBack()
    {
        AddBarChart();
        _pptx.Set("/slide[1]/chart[1]", new() { ["gapwidth"] = "150" });

        var node = _pptx.Get("/slide[1]/chart[1]");
        node.Format.Should().ContainKey("gapwidth");
        ((string)node.Format["gapwidth"].ToString()).Should().Be("150");
    }

    [Fact]
    public void Chart_GapWidth_PersistsAfterReopen()
    {
        AddBarChart();
        _pptx.Set("/slide[1]/chart[1]", new() { ["gapwidth"] = "80" });

        Reopen();

        var node = _pptx.Get("/slide[1]/chart[1]");
        node.Format.Should().ContainKey("gapwidth");
        node.Format["gapwidth"].ToString().Should().Be("80");
    }

    /// <summary>
    /// BUG: overlap is write-only — Set("overlap", "50") sets Overlap on bar chart element,
    /// but ChartReader never reads Overlap back into Format.
    /// </summary>
    [Fact]
    public void Chart_Overlap_IsReadBack()
    {
        AddBarChart();
        _pptx.Set("/slide[1]/chart[1]", new() { ["overlap"] = "50" });

        var node = _pptx.Get("/slide[1]/chart[1]");
        node.Format.Should().ContainKey("overlap");
        node.Format["overlap"].ToString().Should().Be("50");
    }

    [Fact]
    public void Chart_Overlap_PersistsAfterReopen()
    {
        AddBarChart();
        _pptx.Set("/slide[1]/chart[1]", new() { ["overlap"] = "-20" });

        Reopen();

        var node = _pptx.Get("/slide[1]/chart[1]");
        node.Format.Should().ContainKey("overlap");
        node.Format["overlap"].ToString().Should().Be("-20");
    }

    // ==================== Legend Font Readback ====================

    /// <summary>
    /// BUG: legendFont is write-only — Set("legendFont", "10:CCCCCC:Arial") writes TextProperties
    /// to the legend element, but ChartReader only reads the legend position, never font details.
    /// </summary>
    [Fact]
    public void Chart_LegendFont_SizeIsReadBack()
    {
        AddBarChart();
        _pptx.Set("/slide[1]/chart[1]", new() { ["legendFont"] = "10:CCCCCC:Arial" });

        var node = _pptx.Get("/slide[1]/chart[1]");
        node.Format.Should().ContainKey("legendFont");
    }

    // ==================== Axis Font Readback ====================

    /// <summary>
    /// BUG: axisFont is write-only — Set("axisFont", "10:8B949E:Arial") applies TextProperties
    /// to all axis elements, but ChartReader never reads axis TextProperties.
    /// </summary>
    [Fact]
    public void Chart_AxisFont_IsReadBack()
    {
        AddBarChart();
        _pptx.Set("/slide[1]/chart[1]", new() { ["axisFont"] = "10:8B949E:Arial" });

        var node = _pptx.Get("/slide[1]/chart[1]");
        node.Format.Should().ContainKey("axisFont");
    }

    // ==================== Series Shadow Readback ====================

    /// <summary>
    /// BUG: series.shadow is write-only — Set writes OuterShadow to each series' spPr EffectList,
    /// but series readback in ChartReader only reads color, alpha, gradient, lineWidth, lineDash, marker.
    /// </summary>
    [Fact]
    public void Chart_SeriesShadow_IsReadBack()
    {
        AddBarChart();
        _pptx.Set("/slide[1]/chart[1]", new() { ["series.shadow"] = "000000-6-315-4-40" });

        var node = _pptx.Get("/slide[1]/chart[1]", depth: 1);
        node.Children.Should().NotBeEmpty();
        var series1 = node.Children.FirstOrDefault(c => c.Type == "series");
        series1.Should().NotBeNull("at least one series node must be returned");
        series1!.Format.Should().ContainKey("shadow",
            "series.shadow was set but is not read back into the series node's Format");
    }

    // ==================== Series Outline Readback ====================

    /// <summary>
    /// BUG: series.outline is write-only — Set writes Outline with SolidFill to each series' spPr,
    /// but ChartReader only reads lineWidth (from outline.Width) — the outline color is lost.
    /// </summary>
    [Fact]
    public void Chart_SeriesOutline_ColorIsReadBack()
    {
        AddBarChart();
        _pptx.Set("/slide[1]/chart[1]", new() { ["series.outline"] = "FFFFFF-1" });

        var node = _pptx.Get("/slide[1]/chart[1]", depth: 1);
        node.Children.Should().NotBeEmpty();
        var series1 = node.Children.FirstOrDefault(c => c.Type == "series");
        series1.Should().NotBeNull();
        // lineWidth may already be read back; the gap is the color
        series1!.Format.Should().ContainKey("outlineColor",
            "series.outline sets a color but it is never read back — only lineWidth is returned");
    }

    // ==================== Paragraph Alignment Normalization ====================

    /// <summary>
    /// BUG: DrawingML stores paragraph alignment as abbreviated codes:
    ///   "r" = right, "l" = left, "ctr" = center, "just" = justify
    /// PowerPointHandler.Query.cs line 143 returns .InnerText directly, so Get returns "r" not "right".
    /// The CLAUDE.md canonical value for align is a readable string like "right".
    /// </summary>
    [Fact]
    public void Paragraph_AlignRight_NormalizesToRight()
    {
        _pptx.Add("/", "slide", null, new() { ["title"] = "Test" });
        _pptx.Add("/slide[1]", "shape", null, new() { ["text"] = "Hello" });
        _pptx.Set("/slide[1]/shape[1]/paragraph[1]", new() { ["align"] = "right" });

        var para = _pptx.Get("/slide[1]/shape[1]/paragraph[1]");
        para.Format.Should().ContainKey("align");
        var alignVal = (string)para.Format["align"];
        alignVal.Should().Be("right",
            "DrawingML InnerText returns 'r' but Get should normalize to 'right'");
    }

    [Fact]
    public void Paragraph_AlignLeft_NormalizesToLeft()
    {
        _pptx.Add("/", "slide", null, new() { ["title"] = "Test" });
        _pptx.Add("/slide[1]", "shape", null, new() { ["text"] = "Hello" });
        _pptx.Set("/slide[1]/shape[1]/paragraph[1]", new() { ["align"] = "left" });

        var para = _pptx.Get("/slide[1]/shape[1]/paragraph[1]");
        para.Format.Should().ContainKey("align");
        var alignVal = (string)para.Format["align"];
        alignVal.Should().Be("left",
            "DrawingML InnerText may return 'l' but Get should normalize to 'left'");
    }

    [Fact]
    public void Paragraph_AlignCenter_NormalizesToCenter()
    {
        _pptx.Add("/", "slide", null, new() { ["title"] = "Test" });
        _pptx.Add("/slide[1]", "shape", null, new() { ["text"] = "Hello" });
        _pptx.Set("/slide[1]/shape[1]/paragraph[1]", new() { ["align"] = "center" });

        var para = _pptx.Get("/slide[1]/shape[1]/paragraph[1]");
        para.Format.Should().ContainKey("align");
        var alignVal = (string)para.Format["align"];
        alignVal.Should().Be("center",
            "DrawingML InnerText returns 'ctr' but Get should normalize to 'center'");
    }

    [Fact]
    public void Paragraph_AlignJustify_NormalizesToJustify()
    {
        _pptx.Add("/", "slide", null, new() { ["title"] = "Test" });
        _pptx.Add("/slide[1]", "shape", null, new() { ["text"] = "Hello" });
        _pptx.Set("/slide[1]/shape[1]/paragraph[1]", new() { ["align"] = "justify" });

        var para = _pptx.Get("/slide[1]/shape[1]/paragraph[1]");
        para.Format.Should().ContainKey("align");
        var alignVal = (string)para.Format["align"];
        alignVal.Should().Be("justify",
            "DrawingML InnerText returns 'just' but Get should normalize to 'justify'");
    }

    /// <summary>
    /// Verify that alignment normalization persists after file reopen.
    /// </summary>
    [Fact]
    public void Paragraph_AlignRight_PersistsAfterReopen()
    {
        _pptx.Add("/", "slide", null, new() { ["title"] = "Test" });
        _pptx.Add("/slide[1]", "shape", null, new() { ["text"] = "Hello" });
        _pptx.Set("/slide[1]/shape[1]/paragraph[1]", new() { ["align"] = "right" });

        Reopen();

        var para = _pptx.Get("/slide[1]/shape[1]/paragraph[1]");
        ((string)para.Format["align"]).Should().Be("right");
    }

    // ==================== Table Cell Alignment Normalization ====================

    /// <summary>
    /// Same normalization gap exists in table cell alignment readback (Query.cs line 352-358):
    /// Only "ctr" → "center" is mapped; "r" → "right" is not.
    /// </summary>
    [Fact]
    public void TableCell_AlignRight_NormalizesToRight()
    {
        _pptx.Add("/", "slide", null, new() { ["title"] = "T" });
        _pptx.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "2", ["cols"] = "2"
        });
        _pptx.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["align"] = "right" });

        var cell = _pptx.Get("/slide[1]/table[1]/tr[1]/tc[1]");
        cell.Format.Should().ContainKey("align");
        ((string)cell.Format["align"]).Should().Be("right",
            "table cell alignment should normalize 'r' to 'right'");
    }

    // ==================== PlotArea Layout ====================

    [Fact]
    public void Chart_PlotAreaLayout_IsReadBack()
    {
        AddBarChart();
        _pptx.Set("/slide[1]/chart[1]", new()
        {
            ["plotArea.x"] = "0.12",
            ["plotArea.y"] = "0.08",
            ["plotArea.w"] = "0.82",
            ["plotArea.h"] = "0.78"
        });

        var node = _pptx.Get("/slide[1]/chart[1]");
        ((string)node.Format["plotArea.x"]).Should().Be("0.12");
        ((string)node.Format["plotArea.y"]).Should().Be("0.08");
        ((string)node.Format["plotArea.w"]).Should().Be("0.82");
        ((string)node.Format["plotArea.h"]).Should().Be("0.78");
    }

    [Fact]
    public void Chart_PlotAreaLayout_PersistsAcrossReopen()
    {
        AddBarChart();
        _pptx.Set("/slide[1]/chart[1]", new()
        {
            ["plotArea.x"] = "0.15",
            ["plotArea.y"] = "0.10",
            ["plotArea.w"] = "0.80",
            ["plotArea.h"] = "0.75"
        });

        Reopen();

        var node = _pptx.Get("/slide[1]/chart[1]");
        ((string)node.Format["plotArea.x"]).Should().Be("0.15");
        ((string)node.Format["plotArea.y"]).Should().Be("0.1");
        ((string)node.Format["plotArea.w"]).Should().Be("0.8");
        ((string)node.Format["plotArea.h"]).Should().Be("0.75");
    }

    // ==================== Title Layout ====================

    [Fact]
    public void Chart_TitleLayout_IsReadBack()
    {
        AddBarChart();
        _pptx.Set("/slide[1]/chart[1]", new()
        {
            ["title.x"] = "0.3",
            ["title.y"] = "0.02",
            ["title.w"] = "0.4",
            ["title.h"] = "0.08"
        });

        var node = _pptx.Get("/slide[1]/chart[1]");
        ((string)node.Format["title.x"]).Should().Be("0.3");
        ((string)node.Format["title.y"]).Should().Be("0.02");
        ((string)node.Format["title.w"]).Should().Be("0.4");
        ((string)node.Format["title.h"]).Should().Be("0.08");
    }

    [Fact]
    public void Chart_TitleLayout_PersistsAcrossReopen()
    {
        AddBarChart();
        _pptx.Set("/slide[1]/chart[1]", new()
        {
            ["title.x"] = "0.25",
            ["title.y"] = "0.01",
            ["title.w"] = "0.5",
            ["title.h"] = "0.06"
        });

        Reopen();

        var node = _pptx.Get("/slide[1]/chart[1]");
        ((string)node.Format["title.x"]).Should().Be("0.25");
        ((string)node.Format["title.y"]).Should().Be("0.01");
        ((string)node.Format["title.w"]).Should().Be("0.5");
        ((string)node.Format["title.h"]).Should().Be("0.06");
    }

    [Fact]
    public void Chart_TitleLayout_NoTitleReturnsUnsupported()
    {
        AddBarChart();
        // Remove title via Set
        _pptx.Set("/slide[1]/chart[1]", new() { ["title"] = "none" });

        // Setting title layout on chart with no title should not crash
        _pptx.Set("/slide[1]/chart[1]", new() { ["title.x"] = "0.1" });

        var node = _pptx.Get("/slide[1]/chart[1]");
        node.Format.Should().NotContainKey("title.x");
    }

    // ==================== Legend Layout ====================

    [Fact]
    public void Chart_LegendLayout_IsReadBack()
    {
        AddBarChart();
        _pptx.Set("/slide[1]/chart[1]", new()
        {
            ["legend.x"] = "0.7",
            ["legend.y"] = "0.3",
            ["legend.w"] = "0.25",
            ["legend.h"] = "0.4"
        });

        var node = _pptx.Get("/slide[1]/chart[1]");
        ((string)node.Format["legend.x"]).Should().Be("0.7");
        ((string)node.Format["legend.y"]).Should().Be("0.3");
        ((string)node.Format["legend.w"]).Should().Be("0.25");
        ((string)node.Format["legend.h"]).Should().Be("0.4");
    }

    [Fact]
    public void Chart_LegendLayout_PersistsAcrossReopen()
    {
        AddBarChart();
        _pptx.Set("/slide[1]/chart[1]", new()
        {
            ["legend.x"] = "0.75",
            ["legend.y"] = "0.25",
            ["legend.w"] = "0.2",
            ["legend.h"] = "0.5"
        });

        Reopen();

        var node = _pptx.Get("/slide[1]/chart[1]");
        ((string)node.Format["legend.x"]).Should().Be("0.75");
        ((string)node.Format["legend.y"]).Should().Be("0.25");
        ((string)node.Format["legend.w"]).Should().Be("0.2");
        ((string)node.Format["legend.h"]).Should().Be("0.5");
    }

    [Fact]
    public void Chart_LegendLayout_NoLegendReturnsUnsupported()
    {
        AddBarChart();
        // Remove legend
        _pptx.Set("/slide[1]/chart[1]", new() { ["legend"] = "none" });

        // Setting legend layout on chart with no legend should not crash
        _pptx.Set("/slide[1]/chart[1]", new() { ["legend.x"] = "0.1" });

        var node = _pptx.Get("/slide[1]/chart[1]");
        node.Format.Should().NotContainKey("legend.x");
    }

    // ==================== DataLabel Layout ====================

    [Fact]
    public void Chart_DataLabelLayout_IsReadBack()
    {
        AddBarChart();
        _pptx.Set("/slide[1]/chart[1]", new() { ["dataLabels"] = "value" });
        _pptx.Set("/slide[1]/chart[1]", new()
        {
            ["dataLabel1.x"] = "0.05",
            ["dataLabel1.y"] = "0.02",
            ["dataLabel1.w"] = "0.1",
            ["dataLabel1.h"] = "0.06"
        });

        var node = _pptx.Get("/slide[1]/chart[1]");
        ((string)node.Format["dataLabel1.x"]).Should().Be("0.05");
        ((string)node.Format["dataLabel1.y"]).Should().Be("0.02");
        ((string)node.Format["dataLabel1.w"]).Should().Be("0.1");
        ((string)node.Format["dataLabel1.h"]).Should().Be("0.06");
    }

    [Fact]
    public void Chart_DataLabelLayout_MultiplePoints()
    {
        AddBarChart();
        _pptx.Set("/slide[1]/chart[1]", new()
        {
            ["dataLabel1.x"] = "0.1",
            ["dataLabel1.y"] = "0.2",
            ["dataLabel2.x"] = "0.3",
            ["dataLabel2.y"] = "0.4",
            ["dataLabel3.x"] = "0.5",
            ["dataLabel3.y"] = "0.6"
        });

        var node = _pptx.Get("/slide[1]/chart[1]");
        ((string)node.Format["dataLabel1.x"]).Should().Be("0.1");
        ((string)node.Format["dataLabel2.x"]).Should().Be("0.3");
        ((string)node.Format["dataLabel3.x"]).Should().Be("0.5");
    }

    [Fact]
    public void Chart_DataLabelLayout_PersistsAcrossReopen()
    {
        AddBarChart();
        _pptx.Set("/slide[1]/chart[1]", new()
        {
            ["dataLabel1.x"] = "0.12",
            ["dataLabel1.y"] = "0.08"
        });

        Reopen();

        var node = _pptx.Get("/slide[1]/chart[1]");
        ((string)node.Format["dataLabel1.x"]).Should().Be("0.12");
        ((string)node.Format["dataLabel1.y"]).Should().Be("0.08");
    }

    [Fact]
    public void Chart_DataLabelLayout_AutoCreatesDLbls()
    {
        // Chart without explicit dataLabels — Set should auto-create dLbls container
        AddBarChart();
        _pptx.Set("/slide[1]/chart[1]", new() { ["dataLabel1.x"] = "0.15" });

        var node = _pptx.Get("/slide[1]/chart[1]");
        ((string)node.Format["dataLabel1.x"]).Should().Be("0.15");
    }
}
