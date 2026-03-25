// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Functional tests for Excel sparklines: Create → Add → Get → Set → Remove lifecycle.
/// </summary>
public class ExcelSparklineTests : IDisposable
{
    private readonly string _path;
    private ExcelHandler _handler;

    public ExcelSparklineTests()
    {
        _path = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.xlsx");
        BlankDocCreator.Create(_path);
        _handler = new ExcelHandler(_path, editable: true);

        // Add some data for sparklines to reference
        _handler.Set("/Sheet1/A1", new() { ["value"] = "10" });
        _handler.Set("/Sheet1/B1", new() { ["value"] = "20" });
        _handler.Set("/Sheet1/C1", new() { ["value"] = "30" });
        _handler.Set("/Sheet1/D1", new() { ["value"] = "15" });
        _handler.Set("/Sheet1/E1", new() { ["value"] = "25" });
    }

    public void Dispose()
    {
        _handler.Dispose();
        if (File.Exists(_path)) File.Delete(_path);
    }

    private ExcelHandler Reopen()
    {
        _handler.Dispose();
        _handler = new ExcelHandler(_path, editable: true);
        return _handler;
    }

    // ==================== Add + Get ====================

    [Fact]
    public void Add_LineSparkline_ReturnsPath()
    {
        var path = _handler.Add("/Sheet1", "sparkline", null, new()
        {
            ["cell"] = "F1",
            ["range"] = "A1:E1",
            ["type"] = "line"
        });

        path.Should().Be("/Sheet1/sparkline[1]");
    }

    [Fact]
    public void Add_LineSparkline_Get_ReturnsCorrectProperties()
    {
        _handler.Add("/Sheet1", "sparkline", null, new()
        {
            ["cell"] = "F1",
            ["range"] = "A1:E1",
            ["type"] = "line"
        });

        var node = _handler.Get("/Sheet1/sparkline[1]");
        node.Should().NotBeNull();
        node.Type.Should().Be("sparkline");
        node.Format["type"].Should().Be("line");
        node.Format["cell"].Should().Be("F1");
        node.Format["range"].Should().Be("A1:E1");
        node.Format.Should().ContainKey("color");
    }

    [Fact]
    public void Add_ColumnSparkline_Get_ReturnsColumnType()
    {
        _handler.Add("/Sheet1", "sparkline", null, new()
        {
            ["cell"] = "F1",
            ["range"] = "A1:E1",
            ["type"] = "column"
        });

        var node = _handler.Get("/Sheet1/sparkline[1]");
        node.Format["type"].Should().Be("column");
    }

    [Fact]
    public void Add_StackedSparkline_Get_ReturnsStackedType()
    {
        _handler.Add("/Sheet1", "sparkline", null, new()
        {
            ["cell"] = "F1",
            ["range"] = "A1:E1",
            ["type"] = "stacked"
        });

        var node = _handler.Get("/Sheet1/sparkline[1]");
        node.Format["type"].Should().Be("stacked");
    }

    [Fact]
    public void Add_WithColor_Get_ReturnsColor()
    {
        _handler.Add("/Sheet1", "sparkline", null, new()
        {
            ["cell"] = "F1",
            ["range"] = "A1:E1",
            ["color"] = "#FF0000"
        });

        var node = _handler.Get("/Sheet1/sparkline[1]");
        node.Format["color"].Should().Be("#FF0000");
    }

    [Fact]
    public void Add_WithMarkers_Get_ReturnsMarkers()
    {
        _handler.Add("/Sheet1", "sparkline", null, new()
        {
            ["cell"] = "F1",
            ["range"] = "A1:E1",
            ["markers"] = "true",
            ["highPoint"] = "true",
            ["lowPoint"] = "true"
        });

        var node = _handler.Get("/Sheet1/sparkline[1]");
        node.Format["markers"].Should().Be(true);
        node.Format["highPoint"].Should().Be(true);
        node.Format["lowPoint"].Should().Be(true);
    }

    [Fact]
    public void Add_WithFirstLastPoints_Get_ReturnsFlags()
    {
        _handler.Add("/Sheet1", "sparkline", null, new()
        {
            ["cell"] = "F1",
            ["range"] = "A1:E1",
            ["firstPoint"] = "true",
            ["lastPoint"] = "true"
        });

        var node = _handler.Get("/Sheet1/sparkline[1]");
        node.Format["firstPoint"].Should().Be(true);
        node.Format["lastPoint"].Should().Be(true);
    }

    [Fact]
    public void Add_WithNegativeColor_Get_ReturnsNegativeColor()
    {
        _handler.Add("/Sheet1", "sparkline", null, new()
        {
            ["cell"] = "F1",
            ["range"] = "A1:E1",
            ["negative"] = "true",
            ["negativeColor"] = "FF0000"
        });

        var node = _handler.Get("/Sheet1/sparkline[1]");
        node.Format["negative"].Should().Be(true);
        node.Format["negativeColor"].Should().Be("#FF0000");
    }

    [Fact]
    public void Add_WithLineWeight_Get_ReturnsLineWeight()
    {
        _handler.Add("/Sheet1", "sparkline", null, new()
        {
            ["cell"] = "F1",
            ["range"] = "A1:E1",
            ["lineWeight"] = "2.25"
        });

        var node = _handler.Get("/Sheet1/sparkline[1]");
        node.Format["lineWeight"].Should().Be(2.25);
    }

    [Fact]
    public void Add_MissingCell_Throws()
    {
        var act = () => _handler.Add("/Sheet1", "sparkline", null, new()
        {
            ["range"] = "A1:E1"
        });

        act.Should().Throw<ArgumentException>().WithMessage("*cell*");
    }

    [Fact]
    public void Add_MissingRange_Throws()
    {
        var act = () => _handler.Add("/Sheet1", "sparkline", null, new()
        {
            ["cell"] = "F1"
        });

        act.Should().Throw<ArgumentException>().WithMessage("*range*");
    }

    // ==================== Multiple Sparklines ====================

    [Fact]
    public void Add_MultipleSparklines_GetEach()
    {
        _handler.Add("/Sheet1", "sparkline", null, new()
        {
            ["cell"] = "F1",
            ["range"] = "A1:E1",
            ["type"] = "line"
        });

        // Add second row of data
        _handler.Set("/Sheet1/A2", new() { ["value"] = "5" });
        _handler.Set("/Sheet1/E2", new() { ["value"] = "35" });

        _handler.Add("/Sheet1", "sparkline", null, new()
        {
            ["cell"] = "F2",
            ["range"] = "A2:E2",
            ["type"] = "column"
        });

        var node1 = _handler.Get("/Sheet1/sparkline[1]");
        node1.Format["type"].Should().Be("line");
        node1.Format["cell"].Should().Be("F1");

        var node2 = _handler.Get("/Sheet1/sparkline[2]");
        node2.Format["type"].Should().Be("column");
        node2.Format["cell"].Should().Be("F2");
    }

    // ==================== Query ====================

    [Fact]
    public void Query_Sparkline_ReturnsAllSparklines()
    {
        _handler.Add("/Sheet1", "sparkline", null, new()
        {
            ["cell"] = "F1",
            ["range"] = "A1:E1",
            ["type"] = "line"
        });
        _handler.Add("/Sheet1", "sparkline", null, new()
        {
            ["cell"] = "F2",
            ["range"] = "A2:E2",
            ["type"] = "column"
        });

        var results = _handler.Query("sparkline");
        results.Should().HaveCount(2);
        results[0].Type.Should().Be("sparkline");
        results[1].Type.Should().Be("sparkline");
    }

    // ==================== Set ====================

    [Fact]
    public void Set_Color_IsUpdated()
    {
        _handler.Add("/Sheet1", "sparkline", null, new()
        {
            ["cell"] = "F1",
            ["range"] = "A1:E1",
            ["color"] = "4472C4"
        });

        _handler.Set("/Sheet1/sparkline[1]", new() { ["color"] = "#00FF00" });

        var node = _handler.Get("/Sheet1/sparkline[1]");
        node.Format["color"].Should().Be("#00FF00");
    }

    [Fact]
    public void Set_Type_IsUpdated()
    {
        _handler.Add("/Sheet1", "sparkline", null, new()
        {
            ["cell"] = "F1",
            ["range"] = "A1:E1",
            ["type"] = "line"
        });

        _handler.Set("/Sheet1/sparkline[1]", new() { ["type"] = "column" });

        var node = _handler.Get("/Sheet1/sparkline[1]");
        node.Format["type"].Should().Be("column");
    }

    [Fact]
    public void Set_Markers_IsUpdated()
    {
        _handler.Add("/Sheet1", "sparkline", null, new()
        {
            ["cell"] = "F1",
            ["range"] = "A1:E1"
        });

        _handler.Set("/Sheet1/sparkline[1]", new() { ["markers"] = "true" });

        var node = _handler.Get("/Sheet1/sparkline[1]");
        node.Format["markers"].Should().Be(true);
    }

    [Fact]
    public void Set_NegativeColor_IsUpdated()
    {
        _handler.Add("/Sheet1", "sparkline", null, new()
        {
            ["cell"] = "F1",
            ["range"] = "A1:E1"
        });

        _handler.Set("/Sheet1/sparkline[1]", new()
        {
            ["negative"] = "true",
            ["negativeColor"] = "FF0000"
        });

        var node = _handler.Get("/Sheet1/sparkline[1]");
        node.Format["negative"].Should().Be(true);
        node.Format["negativeColor"].Should().Be("#FF0000");
    }

    [Fact]
    public void Set_HighLowPoints_IsUpdated()
    {
        _handler.Add("/Sheet1", "sparkline", null, new()
        {
            ["cell"] = "F1",
            ["range"] = "A1:E1"
        });

        _handler.Set("/Sheet1/sparkline[1]", new()
        {
            ["highPoint"] = "true",
            ["lowPoint"] = "true"
        });

        var node = _handler.Get("/Sheet1/sparkline[1]");
        node.Format["highPoint"].Should().Be(true);
        node.Format["lowPoint"].Should().Be(true);
    }

    [Fact]
    public void Set_LineWeight_IsUpdated()
    {
        _handler.Add("/Sheet1", "sparkline", null, new()
        {
            ["cell"] = "F1",
            ["range"] = "A1:E1"
        });

        _handler.Set("/Sheet1/sparkline[1]", new() { ["lineWeight"] = "3.5" });

        var node = _handler.Get("/Sheet1/sparkline[1]");
        node.Format["lineWeight"].Should().Be(3.5);
    }

    // ==================== Remove ====================

    [Fact]
    public void Remove_Sparkline_IsDeleted()
    {
        _handler.Add("/Sheet1", "sparkline", null, new()
        {
            ["cell"] = "F1",
            ["range"] = "A1:E1"
        });

        var node = _handler.Get("/Sheet1/sparkline[1]");
        node.Should().NotBeNull();

        _handler.Remove("/Sheet1/sparkline[1]");

        var act = () => _handler.Get("/Sheet1/sparkline[1]");
        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void Remove_FirstSparkline_SecondBecomesFirst()
    {
        _handler.Add("/Sheet1", "sparkline", null, new()
        {
            ["cell"] = "F1",
            ["range"] = "A1:E1",
            ["type"] = "line"
        });
        _handler.Add("/Sheet1", "sparkline", null, new()
        {
            ["cell"] = "F2",
            ["range"] = "A2:E2",
            ["type"] = "column"
        });

        _handler.Remove("/Sheet1/sparkline[1]");

        var node = _handler.Get("/Sheet1/sparkline[1]");
        node.Format["type"].Should().Be("column");
        node.Format["cell"].Should().Be("F2");
    }

    // ==================== Persistence ====================

    [Fact]
    public void Add_Sparkline_PersistsAfterReopen()
    {
        _handler.Add("/Sheet1", "sparkline", null, new()
        {
            ["cell"] = "F1",
            ["range"] = "A1:E1",
            ["type"] = "column",
            ["color"] = "#FF5733",
            ["markers"] = "true",
            ["highPoint"] = "true"
        });

        Reopen();

        var node = _handler.Get("/Sheet1/sparkline[1]");
        node.Type.Should().Be("sparkline");
        node.Format["type"].Should().Be("column");
        node.Format["cell"].Should().Be("F1");
        node.Format["range"].Should().Be("A1:E1");
        node.Format["color"].Should().Be("#FF5733");
        node.Format["markers"].Should().Be(true);
        node.Format["highPoint"].Should().Be(true);
    }

    [Fact]
    public void Set_Sparkline_PersistsAfterReopen()
    {
        _handler.Add("/Sheet1", "sparkline", null, new()
        {
            ["cell"] = "F1",
            ["range"] = "A1:E1",
            ["type"] = "line"
        });

        _handler.Set("/Sheet1/sparkline[1]", new()
        {
            ["type"] = "column",
            ["color"] = "#00AA00"
        });

        Reopen();

        var node = _handler.Get("/Sheet1/sparkline[1]");
        node.Format["type"].Should().Be("column");
        node.Format["color"].Should().Be("#00AA00");
    }

    [Fact]
    public void Remove_Sparkline_PersistsAfterReopen()
    {
        _handler.Add("/Sheet1", "sparkline", null, new()
        {
            ["cell"] = "F1",
            ["range"] = "A1:E1"
        });

        _handler.Remove("/Sheet1/sparkline[1]");

        Reopen();

        var results = _handler.Query("sparkline");
        results.Should().BeEmpty();
    }

    // ==================== Default Type ====================

    [Fact]
    public void Add_DefaultType_IsLine()
    {
        _handler.Add("/Sheet1", "sparkline", null, new()
        {
            ["cell"] = "F1",
            ["range"] = "A1:E1"
            // no type specified
        });

        var node = _handler.Get("/Sheet1/sparkline[1]");
        node.Format["type"].Should().Be("line");
    }

    // ==================== Default Color ====================

    [Fact]
    public void Add_DefaultColor_IsBlue()
    {
        _handler.Add("/Sheet1", "sparkline", null, new()
        {
            ["cell"] = "F1",
            ["range"] = "A1:E1"
        });

        var node = _handler.Get("/Sheet1/sparkline[1]");
        node.Format["color"].Should().Be("#4472C4");
    }
}
