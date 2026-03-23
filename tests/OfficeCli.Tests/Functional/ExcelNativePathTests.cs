// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Functional tests for Excel-native path notation (Sheet1!A1) and range Set.
/// Each test follows the full Create → Add/Set → Get → Verify → Reopen → Verify lifecycle.
/// </summary>
public class ExcelNativePathTests : IDisposable
{
    private readonly string _path;
    private ExcelHandler _handler;

    public ExcelNativePathTests()
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

    private ExcelHandler Reopen()
    {
        _handler.Dispose();
        _handler = new ExcelHandler(_path, editable: true);
        return _handler;
    }

    // ==================== Native path: Get ====================

    [Fact]
    public void Get_NativePath_SingleCell_ReturnsNode()
    {
        // Arrange: set a value via DOM path
        _handler.Set("/Sheet1/B3", new() { ["value"] = "hello" });

        // Act: get via native path
        var node = _handler.Get("Sheet1!B3");

        // Assert
        node.Should().NotBeNull();
        node.Text.Should().Be("hello");
        node.Type.Should().Be("cell");
    }

    [Fact]
    public void Get_NativePath_EquivalentToDomPath()
    {
        _handler.Set("/Sheet1/A1", new() { ["value"] = "test" });

        var nativeNode = _handler.Get("Sheet1!A1");
        var domNode = _handler.Get("/Sheet1/A1");

        nativeNode.Text.Should().Be(domNode.Text);
        nativeNode.Type.Should().Be(domNode.Type);
    }

    // ==================== Native path: Set single cell ====================

    [Fact]
    public void Set_NativePath_SingleCell_ValuePersists()
    {
        // Set via native path
        _handler.Set("Sheet1!C5", new() { ["value"] = "42" });

        // Verify via DOM path
        var node = _handler.Get("/Sheet1/C5");
        node.Text.Should().Be("42");

        // Reopen and verify persistence
        Reopen();
        var nodeAfter = _handler.Get("Sheet1!C5");
        nodeAfter.Text.Should().Be("42");
    }

    [Fact]
    public void Set_NativePath_SingleCell_StylePersists()
    {
        _handler.Set("Sheet1!A2", new() { ["font.bold"] = "true", ["fill"] = "FF0000" });

        var node = _handler.Get("Sheet1!A2");
        node.Format["font.bold"].Should().Be(true);
        node.Format["fill"].Should().Be("#FF0000");

        Reopen();
        var nodeAfter = _handler.Get("Sheet1!A2");
        nodeAfter.Format["font.bold"].Should().Be(true);
        nodeAfter.Format["fill"].Should().Be("#FF0000");
    }

    // ==================== Range Set ====================

    [Fact]
    public void Set_Range_AppliesStyleToAllCells()
    {
        // Act: set bold + fill on a 2x3 range
        _handler.Set("/Sheet1/A1:B3", new() { ["font.bold"] = "true", ["fill"] = "4472C4" });

        // Verify all 6 cells have the style
        var cells = new[] { "A1", "A2", "A3", "B1", "B2", "B3" };
        foreach (var cellRef in cells)
        {
            var node = _handler.Get($"/Sheet1/{cellRef}");
            node.Format["font.bold"].Should().Be(true, $"{cellRef} should be bold");
            node.Format["fill"].Should().Be("#4472C4", $"{cellRef} fill mismatch");
        }
    }

    [Fact]
    public void Set_Range_NativePath_AppliesStyleToAllCells()
    {
        // Act: same as above but via native path
        _handler.Set("Sheet1!A1:C2", new() { ["font.italic"] = "true" });

        var cells = new[] { "A1", "B1", "C1", "A2", "B2", "C2" };
        foreach (var cellRef in cells)
        {
            var node = _handler.Get($"Sheet1!{cellRef}");
            node.Format["font.italic"].Should().Be(true, $"{cellRef} should be italic");
        }
    }

    [Fact]
    public void Set_Range_StylePersistsAfterReopen()
    {
        _handler.Set("Sheet1!A1:B2", new() { ["font.bold"] = "true", ["fill"] = "FFFF00" });

        Reopen();

        foreach (var cellRef in new[] { "A1", "A2", "B1", "B2" })
        {
            var node = _handler.Get($"Sheet1!{cellRef}");
            node.Format["font.bold"].Should().Be(true, $"{cellRef} bold not persisted");
            node.Format["fill"].Should().Be("#FFFF00", $"{cellRef} fill not persisted");
        }
    }

    [Fact]
    public void Set_Range_MergeStillWorks()
    {
        // Cell must exist for merge info to be readable via Get
        _handler.Set("/Sheet1/A1", new() { ["value"] = "header" });

        _handler.Set("/Sheet1/A1:C1", new() { ["merge"] = "true" });

        // Merge is visible as "merge" key on the top-left cell of the range
        var cell = _handler.Get("/Sheet1/A1");
        cell.Format.Should().ContainKey("merge");
        cell.Format["merge"].ToString().Should().Be("A1:C1");
    }

    [Fact]
    public void Set_Range_MixedMergeAndStyle()
    {
        // merge + bold in one call — bold creates the cell elements, so merge is readable
        _handler.Set("Sheet1!A1:C1", new() { ["merge"] = "true", ["font.bold"] = "true" });

        // merge applied — visible on top-left cell (cell was created by the bold pass)
        var a1 = _handler.Get("Sheet1!A1");
        a1.Format.Should().ContainKey("merge");
        a1.Format["merge"].ToString().Should().Be("A1:C1");

        // bold applied to all cells
        foreach (var cellRef in new[] { "A1", "B1", "C1" })
        {
            var node = _handler.Get($"Sheet1!{cellRef}");
            node.Format["font.bold"].Should().Be(true, $"{cellRef} should be bold");
        }
    }

    [Fact]
    public void Set_Range_ValueAppliedToAllCells()
    {
        _handler.Set("Sheet1!A1:B2", new() { ["value"] = "X" });

        foreach (var cellRef in new[] { "A1", "B1", "A2", "B2" })
        {
            var node = _handler.Get($"Sheet1!{cellRef}");
            node.Text.Should().Be("X", $"{cellRef} value mismatch");
        }

        Reopen();

        foreach (var cellRef in new[] { "A1", "B1", "A2", "B2" })
        {
            var node = _handler.Get($"Sheet1!{cellRef}");
            node.Text.Should().Be("X", $"{cellRef} value not persisted");
        }
    }

    // ==================== Native path: Query ====================

    [Fact]
    public void Query_NativeCellRef_ReturnsSingleNode()
    {
        _handler.Set("/Sheet1/D4", new() { ["value"] = "queryMe" });

        var results = _handler.Query("Sheet1!D4");

        results.Should().HaveCount(1);
        results[0].Text.Should().Be("queryMe");
    }

    [Fact]
    public void Query_NativeCellRef_Range_ReturnsRangeNode()
    {
        _handler.Set("/Sheet1/A1", new() { ["value"] = "1" });
        _handler.Set("/Sheet1/B1", new() { ["value"] = "2" });

        var results = _handler.Query("Sheet1!A1:B1");

        results.Should().HaveCount(1);
        results[0].Type.Should().Be("range");
        results[0].Children.Should().HaveCount(2);
    }

    // ==================== Native path: Remove ====================

    [Fact]
    public void Remove_NativePath_RemovesCell()
    {
        _handler.Set("/Sheet1/E5", new() { ["value"] = "toDelete" });
        _handler.Get("/Sheet1/E5").Text.Should().Be("toDelete");

        _handler.Remove("Sheet1!E5");

        // After removal the cell element is gone — Get returns a stub node with "(empty)" text
        var after = _handler.Get("/Sheet1/E5");
        after.Text.Should().Be("(empty)");
    }
}
