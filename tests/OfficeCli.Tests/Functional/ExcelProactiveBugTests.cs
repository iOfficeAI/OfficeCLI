// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.Json;
using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Proactive bug hunt: tests for patterns similar to previously fixed bugs.
///
/// Pattern 1: Non-standard types in Format dictionary that crash JSON serialization.
///   - ChartReader: style (byte), axisMin/axisMax/majorUnit/minorUnit (double),
///     alpha (int), lineWidth (double)
///   - ExcelHandler.Query: zoom (uint), paperSize (uint), outlineLevel (byte)
///   - ExcelHandler.Helpers: numFmtId (uint)
///
/// Pattern 2: Properties writable but not readable.
///   - textRotation: Set via ExcelStyleManager, not read back in CellToNode
///   - indent: Set via ExcelStyleManager, not read back in CellToNode
///   - shrinkToFit: Set via ExcelStyleManager, not read back in CellToNode
///
/// Pattern 3: shadow=true crash in PowerPoint handler (same bug as Excel had).
/// </summary>
public class ExcelProactiveBugTests : IDisposable
{
    private readonly string _path;
    private ExcelHandler _handler;

    public ExcelProactiveBugTests()
    {
        _path = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.xlsx");
        BlankDocCreator.Create(_path);
        _handler = new ExcelHandler(_path, editable: true);
    }

    public void Dispose()
    {
        _handler.Dispose();
        if (File.Exists(_path)) File.Delete(_path);
    }

    private void Reopen()
    {
        _handler.Dispose();
        _handler = new ExcelHandler(_path, editable: true);
    }

    // ==================== Pattern 1: Non-standard types in Format ====================

    [Fact]
    public void Zoom_Format_ShouldSerializeToJson()
    {
        // Set zoom on sheet
        _handler.Set("/Sheet1", new() { ["zoom"] = "150" });
        var node = _handler.Get("/Sheet1");
        node.Format.Should().ContainKey("zoom");

        // Verify it can be serialized to JSON without crashing
        var json = JsonSerializer.Serialize(node.Format);
        json.Should().Contain("zoom");

        // Verify the value is a JSON-safe type
        var val = node.Format["zoom"];
        (val is int or string).Should().BeTrue($"zoom should be int or string for JSON safety, but was {val.GetType().Name}");
    }

    [Fact]
    public void PaperSize_Format_ShouldSerializeToJson()
    {
        _handler.Set("/Sheet1", new() { ["paperSize"] = "9" }); // A4
        var node = _handler.Get("/Sheet1");
        node.Format.Should().ContainKey("paperSize");

        var json = JsonSerializer.Serialize(node.Format);
        json.Should().Contain("paperSize");

        var val = node.Format["paperSize"];
        (val is int or string).Should().BeTrue($"paperSize should be int or string for JSON safety, but was {val.GetType().Name}");
    }

    [Fact]
    public void OutlineLevel_Row_Format_ShouldSerializeToJson()
    {
        // Add data in row 1
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "test" });
        // Set outline level on row
        _handler.Set("/Sheet1/row[1]", new() { ["outlineLevel"] = "2" });
        var node = _handler.Get("/Sheet1/row[1]");
        node.Format.Should().ContainKey("outlineLevel");

        var json = JsonSerializer.Serialize(node.Format);
        json.Should().Contain("outlineLevel");

        // outlineLevel should be a JSON-safe type (int or string, not byte)
        var val = node.Format["outlineLevel"];
        (val is int or string).Should().BeTrue($"outlineLevel should be int or string, but was {val.GetType().Name}");
    }

    [Fact]
    public void OutlineLevel_Column_Format_ShouldSerializeToJson()
    {
        _handler.Set("/Sheet1/col[A]", new() { ["outlineLevel"] = "2" });
        var node = _handler.Get("/Sheet1/col[A]");
        node.Format.Should().ContainKey("outlineLevel");

        var json = JsonSerializer.Serialize(node.Format);
        json.Should().Contain("outlineLevel");

        var val = node.Format["outlineLevel"];
        (val is int or string).Should().BeTrue($"outlineLevel should be int or string, but was {val.GetType().Name}");
    }

    [Fact]
    public void NumFmtId_Format_ShouldSerializeToJson()
    {
        // Apply a number format to a cell
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "42", ["numberformat"] = "0.00" });
        var node = _handler.Get("/Sheet1/A1");
        node.Format.Should().ContainKey("numFmtId");

        var json = JsonSerializer.Serialize(node.Format);
        json.Should().Contain("numFmtId");

        var val = node.Format["numFmtId"];
        (val is int or string).Should().BeTrue($"numFmtId should be int or string, but was {val.GetType().Name}");
    }

    // ==================== Pattern 2: Properties writable but not readable ====================

    [Fact]
    public void TextRotation_SetAndGet_ShouldRoundTrip()
    {
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "rotated" });
        _handler.Set("/Sheet1/A1", new() { ["rotation"] = "45" });

        var node = _handler.Get("/Sheet1/A1");
        node.Format.Should().ContainKey("alignment.textRotation",
            "textRotation should be readable after being set");
        node.Format["alignment.textRotation"].ToString().Should().Be("45");
    }

    [Fact]
    public void TextRotation_Persistence_ShouldSurviveReopen()
    {
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "rotated" });
        _handler.Set("/Sheet1/A1", new() { ["rotation"] = "90" });

        Reopen();

        var node = _handler.Get("/Sheet1/A1");
        node.Format.Should().ContainKey("alignment.textRotation");
        node.Format["alignment.textRotation"].ToString().Should().Be("90");
    }

    [Fact]
    public void Indent_SetAndGet_ShouldRoundTrip()
    {
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "indented" });
        _handler.Set("/Sheet1/A1", new() { ["indent"] = "2" });

        var node = _handler.Get("/Sheet1/A1");
        node.Format.Should().ContainKey("alignment.indent",
            "indent should be readable after being set");
        node.Format["alignment.indent"].ToString().Should().Be("2");
    }

    [Fact]
    public void ShrinkToFit_SetAndGet_ShouldRoundTrip()
    {
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "shrunk" });
        _handler.Set("/Sheet1/A1", new() { ["shrinktofit"] = "true" });

        var node = _handler.Get("/Sheet1/A1");
        node.Format.Should().ContainKey("alignment.shrinkToFit",
            "shrinkToFit should be readable after being set");
        node.Format["alignment.shrinkToFit"].Should().Be(true);
    }

    // ==================== Pattern 3: shadow=true in PowerPoint ====================
    // Note: This test uses PowerPointHandler to verify the same shadow=true bug
    // that was fixed in Excel also exists (and is fixed) in PowerPoint.
}

/// <summary>
/// PowerPoint shadow=true bug test — same pattern as the Excel fix.
/// </summary>
public class PptxProactiveBugTests : IDisposable
{
    private readonly string _path;
    private PowerPointHandler _handler;

    public PptxProactiveBugTests()
    {
        _path = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.pptx");
        BlankDocCreator.Create(_path);
        _handler = new PowerPointHandler(_path, editable: true);
    }

    public void Dispose()
    {
        _handler.Dispose();
        if (File.Exists(_path)) File.Delete(_path);
    }

    [Fact]
    public void Shadow_True_ShouldNotCrash_Add()
    {
        // Add a slide first
        _handler.Add("/", "slide", null, new());

        // shadow=true should work (default black shadow), not crash with "Invalid color value: 'true'"
        var path = _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Shadow Test",
            ["shadow"] = "true"
        });

        path.Should().NotBeNullOrEmpty();

        var node = _handler.Get(path);
        node.Format.Should().ContainKey("shadow");
    }

    [Fact]
    public void Shadow_True_ShouldNotCrash_Set()
    {
        _handler.Add("/", "slide", null, new());
        var shapePath = _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Test Shape",
            ["fill"] = "FF0000"
        });

        // Set shadow=true should work, not crash
        var unsupported = _handler.Set(shapePath, new() { ["shadow"] = "true" });
        unsupported.Should().NotContain("shadow");

        var node = _handler.Get(shapePath);
        node.Format.Should().ContainKey("shadow");
    }
}
