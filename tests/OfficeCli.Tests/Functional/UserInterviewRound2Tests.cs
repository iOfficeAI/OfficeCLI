// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Round 2 user interview tests — proving 3 partially-passed issues from Round 1,
/// plus proactively discovered issues in Excel/Word chart handling.
///
/// Issue 1: preset=corporate on chart without title reports unsupported for title.bold/size/color
/// Issue 2: waterfall chart readback shows column_stacked instead of waterfall
/// Issue 3: Get('/slide[1]/chart[1]/series[1]') throws instead of returning series node
///
/// Proactive findings:
/// Issue 4: Excel chart Set does not support series sub-path
/// Issue 5: Excel chart preset (corporate) may fail on charts without title
/// Issue 6: Word chart Set does not support series sub-path
/// </summary>
public class UserInterviewRound2Tests : IDisposable
{
    private readonly string _pptxPath;
    private readonly string _xlsxPath;
    private readonly string _docxPath;
    private PowerPointHandler _pptxHandler;
    private ExcelHandler _excelHandler;
    private WordHandler _wordHandler;

    public UserInterviewRound2Tests()
    {
        _pptxPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.pptx");
        _xlsxPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.xlsx");
        _docxPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.docx");

        BlankDocCreator.Create(_pptxPath);
        BlankDocCreator.Create(_xlsxPath);
        BlankDocCreator.Create(_docxPath);

        _pptxHandler = new PowerPointHandler(_pptxPath, editable: true);
        _excelHandler = new ExcelHandler(_xlsxPath, editable: true);
        _wordHandler = new WordHandler(_docxPath, editable: true);
    }

    public void Dispose()
    {
        _pptxHandler?.Dispose();
        _excelHandler?.Dispose();
        _wordHandler?.Dispose();
        if (File.Exists(_pptxPath)) File.Delete(_pptxPath);
        if (File.Exists(_xlsxPath)) File.Delete(_xlsxPath);
        if (File.Exists(_docxPath)) File.Delete(_docxPath);
    }

    private void ReopenPptx()
    {
        _pptxHandler?.Dispose();
        _pptxHandler = new PowerPointHandler(_pptxPath, editable: true);
    }

    // ==================== Issue 1: preset=corporate on chart without title ====================

    /// <summary>
    /// When a chart has no title (title=none), applying preset=corporate should not
    /// report title.bold, title.size, title.color as unsupported. The preset should
    /// silently skip properties that don't apply.
    /// </summary>
    [Fact]
    public void Issue1_PresetCorporate_NoTitle_ShouldNotReportUnsupported()
    {
        // Arrange: create chart without title
        _pptxHandler.Add("/", "slide", null, new() { ["title"] = "Preset Test" });
        _pptxHandler.Add("/slide[1]", "chart", null, new()
        {
            ["type"] = "column",
            ["data"] = "Sales:10,20,30",
            ["categories"] = "Q1,Q2,Q3"
        });

        // Remove title if present by default
        _pptxHandler.Set("/slide[1]/chart[1]", new() { ["title"] = "none" });

        // Act: apply corporate preset (which includes title.bold, title.size, title.color)
        var unsupported = _pptxHandler.Set("/slide[1]/chart[1]", new() { ["preset"] = "corporate" });

        // Assert: title-related properties should NOT appear in unsupported list
        unsupported.Should().NotContain(k => k.Contains("title", StringComparison.OrdinalIgnoreCase),
            "preset should silently skip title properties when chart has no title, not report them as unsupported");
    }

    // ==================== Issue 2: waterfall readback shows column_stacked ====================

    /// <summary>
    /// Creating a waterfall chart and reading it back should show chartType=waterfall,
    /// not column_stacked. The internal stacked bar simulation should be detected.
    /// </summary>
    [Fact]
    public void Issue2_WaterfallChart_ReadbackShouldShowWaterfall()
    {
        // Arrange: create waterfall chart
        _pptxHandler.Add("/", "slide", null, new() { ["title"] = "Waterfall Test" });
        _pptxHandler.Add("/slide[1]", "chart", null, new()
        {
            ["type"] = "waterfall",
            ["data"] = "Values:100,-20,30,-10,50",
            ["categories"] = "Start,Loss1,Gain1,Loss2,End"
        });

        // Act: read back chart
        var chartNode = _pptxHandler.Get("/slide[1]/chart[1]");

        // Assert: chartType should be "waterfall", not "column_stacked"
        chartNode.Should().NotBeNull();
        chartNode.Format.Should().ContainKey("chartType");
        chartNode.Format["chartType"].ToString().Should().Be("waterfall",
            "waterfall chart should be detected as 'waterfall', not as the underlying 'column_stacked' implementation");
    }

    /// <summary>
    /// Waterfall readback should survive persistence (reopen).
    /// </summary>
    [Fact]
    public void Issue2_WaterfallChart_ReadbackAfterReopen_ShouldShowWaterfall()
    {
        _pptxHandler.Add("/", "slide", null, new() { ["title"] = "Waterfall Persist" });
        _pptxHandler.Add("/slide[1]", "chart", null, new()
        {
            ["type"] = "waterfall",
            ["data"] = "Values:100,-20,30,-10,50",
            ["categories"] = "Start,Loss1,Gain1,Loss2,End"
        });

        ReopenPptx();

        var chartNode = _pptxHandler.Get("/slide[1]/chart[1]");
        chartNode.Should().NotBeNull();
        chartNode.Format.Should().ContainKey("chartType");
        chartNode.Format["chartType"].ToString().Should().Be("waterfall");
    }

    // ==================== Issue 3: Get series sub-path ====================

    /// <summary>
    /// Get('/slide[1]/chart[1]/series[1]') should return a series node.
    /// Currently it throws "Element not found" because the path regex doesn't support 3 segments.
    /// </summary>
    [Fact]
    public void Issue3_GetSeriesPath_ShouldReturnSeriesNode()
    {
        // Arrange
        _pptxHandler.Add("/", "slide", null, new() { ["title"] = "Series Get Test" });
        _pptxHandler.Add("/slide[1]", "chart", null, new()
        {
            ["type"] = "column",
            ["data"] = "Sales:10,20,30",
            ["categories"] = "Q1,Q2,Q3"
        });

        // Act: Get series by direct path
        var act = () => _pptxHandler.Get("/slide[1]/chart[1]/series[1]");

        // Assert: should not throw, should return a valid node
        var node = act.Should().NotThrow("Get should support /slide[N]/chart[M]/series[K] paths")
            .Which;
        node.Should().NotBeNull();
        node.Type.Should().Be("series");
        node.Path.Should().Contain("series[1]");
    }

    /// <summary>
    /// Get series path should return format properties (color, name, etc.)
    /// </summary>
    [Fact]
    public void Issue3_GetSeriesPath_ShouldIncludeFormatProperties()
    {
        _pptxHandler.Add("/", "slide", null, new() { ["title"] = "Series Props" });
        _pptxHandler.Add("/slide[1]", "chart", null, new()
        {
            ["type"] = "column",
            ["data"] = "Revenue:100,200,300;Cost:50,80,120",
            ["categories"] = "Q1,Q2,Q3"
        });

        // Set a color on series 1
        _pptxHandler.Set("/slide[1]/chart[1]/series[1]", new() { ["color"] = "FF0000" });

        // Get the series directly
        var node = _pptxHandler.Get("/slide[1]/chart[1]/series[1]");
        node.Should().NotBeNull();
        node.Format.Should().ContainKey("color",
            "series node from Get should include format properties like color");
    }

    /// <summary>
    /// Get/Set symmetry: if Set works on a path, Get should also work on it.
    /// </summary>
    [Fact]
    public void Issue3_GetSetSymmetry_SeriesPath()
    {
        _pptxHandler.Add("/", "slide", null, new() { ["title"] = "Symmetry Test" });
        _pptxHandler.Add("/slide[1]", "chart", null, new()
        {
            ["type"] = "column",
            ["data"] = "Sales:10,20,30",
            ["categories"] = "Q1,Q2,Q3"
        });

        // Set should work (already proven)
        var setAct = () => _pptxHandler.Set("/slide[1]/chart[1]/series[1]", new() { ["color"] = "00FF00" });
        setAct.Should().NotThrow();

        // Get should also work on the same path
        var getAct = () => _pptxHandler.Get("/slide[1]/chart[1]/series[1]");
        getAct.Should().NotThrow("Get and Set should support the same paths — API symmetry");
    }

    // ==================== Proactive Issue 4: Excel chart series sub-path ====================

    /// <summary>
    /// Excel Set on chart series sub-path (/Sheet1/chart[1]/series[1]) should work,
    /// similar to PPTX series sub-path support.
    /// </summary>
    [Fact]
    public void Proactive4_Excel_SetSeriesSubPath_ShouldWork()
    {
        // Arrange: create chart in Excel
        _excelHandler.Add("/Sheet1", "chart", null, new()
        {
            ["type"] = "column",
            ["data"] = "Sales:10,20,30",
            ["categories"] = "Q1,Q2,Q3"
        });

        // Act: Set color on series sub-path
        var act = () => _excelHandler.Set("/Sheet1/chart[1]/series[1]", new() { ["color"] = "FF0000" });

        // Assert
        act.Should().NotThrow("Excel Set should support chart series sub-paths like PPTX does");
    }

    /// <summary>
    /// Excel Get on chart series sub-path should also work.
    /// </summary>
    [Fact]
    public void Proactive4_Excel_GetSeriesSubPath_ShouldWork()
    {
        _excelHandler.Add("/Sheet1", "chart", null, new()
        {
            ["type"] = "column",
            ["data"] = "Sales:10,20,30",
            ["categories"] = "Q1,Q2,Q3"
        });

        // Act: Get series by sub-path
        var act = () => _excelHandler.Get("/Sheet1/chart[1]/series[1]");

        act.Should().NotThrow("Excel Get should support chart series sub-paths");
        var node = _excelHandler.Get("/Sheet1/chart[1]/series[1]");
        node.Should().NotBeNull();
        node.Type.Should().Be("series");
    }

    // ==================== Proactive Issue 5: Excel chart preset with no title ====================

    /// <summary>
    /// Excel chart with preset=corporate and no title should not report unsupported.
    /// Same issue as PPTX Issue 1 but on Excel side.
    /// </summary>
    [Fact]
    public void Proactive5_Excel_PresetCorporate_NoTitle_ShouldNotReportUnsupported()
    {
        // Arrange
        _excelHandler.Add("/Sheet1", "chart", null, new()
        {
            ["type"] = "column",
            ["data"] = "Sales:10,20,30",
            ["categories"] = "Q1,Q2,Q3"
        });

        // Remove title
        _excelHandler.Set("/Sheet1/chart[1]", new() { ["title"] = "none" });

        // Act: apply corporate preset
        var unsupported = _excelHandler.Set("/Sheet1/chart[1]", new() { ["preset"] = "corporate" });

        // Assert
        unsupported.Should().NotContain(k => k.Contains("title", StringComparison.OrdinalIgnoreCase),
            "Excel preset should also silently skip title properties when chart has no title");
    }

    // ==================== Proactive Issue 6: Word chart series sub-path ====================

    /// <summary>
    /// Word chart Set on series sub-path (/chart[1]/series[1]) should work.
    /// </summary>
    [Fact]
    public void Proactive6_Word_SetSeriesSubPath_ShouldWork()
    {
        // Arrange: create chart in Word
        _wordHandler.Add("/", "chart", null, new()
        {
            ["type"] = "column",
            ["data"] = "Sales:10,20,30",
            ["categories"] = "Q1,Q2,Q3"
        });

        // Act: Set color on series sub-path
        var act = () => _wordHandler.Set("/chart[1]/series[1]", new() { ["color"] = "FF0000" });

        // Assert
        act.Should().NotThrow("Word Set should support chart series sub-paths");
    }

    /// <summary>
    /// Word chart Get on series sub-path should also work.
    /// </summary>
    [Fact]
    public void Proactive6_Word_GetSeriesSubPath_ShouldWork()
    {
        _wordHandler.Add("/", "chart", null, new()
        {
            ["type"] = "column",
            ["data"] = "Sales:10,20,30",
            ["categories"] = "Q1,Q2,Q3"
        });

        // Act
        var act = () => _wordHandler.Get("/chart[1]/series[1]");

        act.Should().NotThrow("Word Get should support chart series sub-paths");
        var node = _wordHandler.Get("/chart[1]/series[1]");
        node.Should().NotBeNull();
        node.Type.Should().Be("series");
    }
}
