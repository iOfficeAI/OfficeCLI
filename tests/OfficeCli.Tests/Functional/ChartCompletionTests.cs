// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Tests for four new chart features:
/// #1 dataLabel{N}.text — custom per-point label text
/// #2 axisLine / catAxisLine / valAxisLine — axis line styling
/// #3 Excel chart Remove
/// #4 Word chart Remove
/// </summary>
public class ChartCompletionTests : IDisposable
{
    private readonly string _xlsxPath;
    private readonly string _pptxPath;
    private readonly string _docxPath;

    private ExcelHandler _excel;
    private PowerPointHandler _pptx;
    private WordHandler _word;

    public ChartCompletionTests()
    {
        _xlsxPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.xlsx");
        _pptxPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.pptx");
        _docxPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.docx");

        BlankDocCreator.Create(_xlsxPath);
        BlankDocCreator.Create(_pptxPath);
        BlankDocCreator.Create(_docxPath);

        _excel = new ExcelHandler(_xlsxPath, editable: true);
        _pptx = new PowerPointHandler(_pptxPath, editable: true);
        _word = new WordHandler(_docxPath, editable: true);
    }

    public void Dispose()
    {
        _excel.Dispose();
        _pptx.Dispose();
        _word.Dispose();
        if (File.Exists(_xlsxPath)) File.Delete(_xlsxPath);
        if (File.Exists(_pptxPath)) File.Delete(_pptxPath);
        if (File.Exists(_docxPath)) File.Delete(_docxPath);
    }

    private void ReopenExcel() { _excel.Dispose(); _excel = new ExcelHandler(_xlsxPath, editable: true); }
    private void ReopenPptx() { _pptx.Dispose(); _pptx = new PowerPointHandler(_pptxPath, editable: true); }
    private void ReopenWord() { _word.Dispose(); _word = new WordHandler(_docxPath, editable: true); }

    private string AddExcelChart(Dictionary<string, string>? extraProps = null)
    {
        var props = new Dictionary<string, string>
        {
            ["chartType"] = "column",
            ["title"] = "Test",
            ["data"] = "S1:10,20,30;S2:15,25,35",
            ["categories"] = "A,B,C"
        };
        if (extraProps != null) foreach (var kv in extraProps) props[kv.Key] = kv.Value;
        return _excel.Add("/Sheet1", "chart", null, props);
    }

    private string AddPptxChart(Dictionary<string, string>? extraProps = null)
    {
        _pptx.Add("/", "slide", null, new());
        var props = new Dictionary<string, string>
        {
            ["chartType"] = "column",
            ["title"] = "Test",
            ["data"] = "S1:10,20,30;S2:15,25,35",
            ["categories"] = "A,B,C"
        };
        if (extraProps != null) foreach (var kv in extraProps) props[kv.Key] = kv.Value;
        return _pptx.Add("/slide[1]", "chart", null, props);
    }

    private string AddWordChart(Dictionary<string, string>? extraProps = null)
    {
        var props = new Dictionary<string, string>
        {
            ["chartType"] = "column",
            ["title"] = "WordTest",
            ["data"] = "S1:5,10,15",
            ["categories"] = "X,Y,Z"
        };
        if (extraProps != null) foreach (var kv in extraProps) props[kv.Key] = kv.Value;
        return _word.Add("/", "chart", null, props);
    }

    // ==================== #1 dataLabel{N}.text ====================

    [Fact]
    public void Excel_Set_DataLabel1_Text_ReadBack()
    {
        var chartPath = AddExcelChart();
        // Must enable dataLabels first so dLbls element exists
        _excel.Set(chartPath, new() { ["dataLabels"] = "value" });
        _excel.Set(chartPath, new() { ["dataLabel1.text"] = "Custom A" });

        var node = _excel.Get(chartPath, depth: 0);
        node.Format["dataLabel1.text"].Should().Be("Custom A");
    }

    [Fact]
    public void Excel_Set_DataLabel1_Text_Persistence()
    {
        var chartPath = AddExcelChart();
        _excel.Set(chartPath, new() { ["dataLabels"] = "value" });
        _excel.Set(chartPath, new() { ["dataLabel1.text"] = "Persist Me" });

        ReopenExcel();
        var node = _excel.Get(chartPath, depth: 0);
        node.Format["dataLabel1.text"].Should().Be("Persist Me");
    }

    [Fact]
    public void Excel_Set_Multiple_DataLabel_Texts()
    {
        var chartPath = AddExcelChart();
        _excel.Set(chartPath, new() { ["dataLabels"] = "value" });
        _excel.Set(chartPath, new()
        {
            ["dataLabel1.text"] = "First",
            ["dataLabel2.text"] = "Second"
        });

        var node = _excel.Get(chartPath, depth: 0);
        node.Format["dataLabel1.text"].Should().Be("First");
        node.Format["dataLabel2.text"].Should().Be("Second");
    }

    [Fact]
    public void Pptx_Set_DataLabel1_Text_ReadBack()
    {
        var chartPath = AddPptxChart();
        _pptx.Set(chartPath, new() { ["dataLabels"] = "value" });
        _pptx.Set(chartPath, new() { ["dataLabel1.text"] = "PPTX Label" });

        var node = _pptx.Get(chartPath, depth: 0);
        node.Format["dataLabel1.text"].Should().Be("PPTX Label");
    }

    [Fact]
    public void Pptx_Set_DataLabel1_Text_Persistence()
    {
        var chartPath = AddPptxChart();
        _pptx.Set(chartPath, new() { ["dataLabels"] = "value" });
        _pptx.Set(chartPath, new() { ["dataLabel1.text"] = "Saved Label" });

        ReopenPptx();
        var node = _pptx.Get(chartPath, depth: 0);
        node.Format["dataLabel1.text"].Should().Be("Saved Label");
    }

    // ==================== #2 axisLine / catAxisLine / valAxisLine ====================

    [Fact]
    public void Excel_Set_ValAxisLine_ColorAndWidth()
    {
        var chartPath = AddExcelChart();
        _excel.Set(chartPath, new() { ["valAxisLine"] = "333333:1.5" });

        var node = _excel.Get(chartPath, depth: 0);
        node.Format["valAxisLine.color"].Should().Be("#333333");
        node.Format["valAxisLine.width"].Should().Be(1.5);
    }

    [Fact]
    public void Excel_Set_CatAxisLine_ColorAndWidth()
    {
        var chartPath = AddExcelChart();
        _excel.Set(chartPath, new() { ["catAxisLine"] = "000000:1" });

        var node = _excel.Get(chartPath, depth: 0);
        node.Format["catAxisLine.color"].Should().Be("#000000");
        node.Format["catAxisLine.width"].Should().Be(1.0);
    }

    [Fact]
    public void Excel_Set_AxisLine_BothAxes()
    {
        var chartPath = AddExcelChart();
        _excel.Set(chartPath, new() { ["axisLine"] = "AAAAAA:0.5" });

        var node = _excel.Get(chartPath, depth: 0);
        node.Format["valAxisLine.color"].Should().Be("#AAAAAA");
        node.Format["catAxisLine.color"].Should().Be("#AAAAAA");
    }

    [Fact]
    public void Excel_Set_AxisLine_None_HidesLine()
    {
        var chartPath = AddExcelChart();
        // First set a visible line, then hide it
        _excel.Set(chartPath, new() { ["valAxisLine"] = "000000:1" });
        _excel.Set(chartPath, new() { ["valAxisLine"] = "none" });

        var node = _excel.Get(chartPath, depth: 0);
        node.Format.Should().NotContainKey("valAxisLine.color");
    }

    [Fact]
    public void Excel_Set_AxisLine_WithDash()
    {
        var chartPath = AddExcelChart();
        _excel.Set(chartPath, new() { ["valAxisLine"] = "FF0000:1:dash" });

        var node = _excel.Get(chartPath, depth: 0);
        node.Format["valAxisLine.color"].Should().Be("#FF0000");
        node.Format["valAxisLine.width"].Should().Be(1.0);
        node.Format["valAxisLine.dash"].Should().NotBeNull();
    }

    [Fact]
    public void Excel_Set_ValAxisLine_Persistence()
    {
        var chartPath = AddExcelChart();
        _excel.Set(chartPath, new() { ["valAxisLine"] = "4472C4:2" });

        ReopenExcel();
        var node = _excel.Get(chartPath, depth: 0);
        node.Format["valAxisLine.color"].Should().Be("#4472C4");
        node.Format["valAxisLine.width"].Should().Be(2.0);
    }

    [Fact]
    public void Pptx_Set_ValAxisLine_ColorAndWidth()
    {
        var chartPath = AddPptxChart();
        _pptx.Set(chartPath, new() { ["valAxisLine"] = "FF6600:1.5" });

        var node = _pptx.Get(chartPath, depth: 0);
        node.Format["valAxisLine.color"].Should().Be("#FF6600");
        node.Format["valAxisLine.width"].Should().Be(1.5);
    }

    [Fact]
    public void Pptx_Set_CatAxisLine_Separate()
    {
        var chartPath = AddPptxChart();
        _pptx.Set(chartPath, new() { ["catAxisLine"] = "0070C0:0.75" });
        _pptx.Set(chartPath, new() { ["valAxisLine"] = "FF0000:2" });

        var node = _pptx.Get(chartPath, depth: 0);
        node.Format["catAxisLine.color"].Should().Be("#0070C0");
        node.Format["valAxisLine.color"].Should().Be("#FF0000");
    }

    // ==================== #3 Excel chart Remove ====================

    [Fact]
    public void Excel_Remove_Chart_DisappearsFromQuery()
    {
        // Add chart and verify it exists
        AddExcelChart();
        var before = _excel.Query("chart");
        before.Should().NotBeEmpty();

        // Remove and verify it's gone
        _excel.Remove("/Sheet1/chart[1]");
        var after = _excel.Query("chart");
        after.Should().BeEmpty();
    }

    [Fact]
    public void Excel_Remove_Chart_OnlyRemovesTarget()
    {
        AddExcelChart();
        AddExcelChart();
        var before = _excel.Query("chart");
        before.Should().HaveCount(2);

        _excel.Remove("/Sheet1/chart[1]");
        var after = _excel.Query("chart");
        after.Should().HaveCount(1);
    }

    [Fact]
    public void Excel_Remove_Chart_Persistence()
    {
        AddExcelChart();
        _excel.Remove("/Sheet1/chart[1]");

        ReopenExcel();
        var charts = _excel.Query("chart");
        charts.Should().BeEmpty();
    }

    // ==================== #4 Word chart Remove ====================

    [Fact]
    public void Word_Remove_Chart_DisappearsFromQuery()
    {
        AddWordChart();
        var before = _word.Query("chart");
        before.Should().NotBeEmpty();

        _word.Remove("/chart[1]");
        var after = _word.Query("chart");
        after.Should().BeEmpty();
    }

    [Fact]
    public void Word_Remove_Chart_OnlyRemovesTarget()
    {
        AddWordChart();
        AddWordChart();
        var before = _word.Query("chart");
        before.Should().HaveCount(2);

        _word.Remove("/chart[1]");
        var after = _word.Query("chart");
        after.Should().HaveCount(1);
    }

    [Fact]
    public void Word_Remove_Chart_Persistence()
    {
        AddWordChart();
        _word.Remove("/chart[1]");

        ReopenWord();
        var charts = _word.Query("chart");
        charts.Should().BeEmpty();
    }
}
