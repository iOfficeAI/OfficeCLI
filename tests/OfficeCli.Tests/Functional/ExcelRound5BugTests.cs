// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Tests for Round 5 bugs reported by Agent A (Sonnet).
///
/// CONFIRMED BUG (failing):
///
///   Bug 1 (HIGH) — remove /Sheet1/shape[1] fails — shapes cannot be deleted
///            Root cause: ExcelHandler.Remove in ExcelHandler.Remove.cs has no handler
///            for shape[N] or picture[N] path segments. After dispatching row/column/cell/
///            break paths, the remaining segment "shape[1]" is treated as a raw cell
///            reference. FindCell returns null, and Remove throws:
///            "Cell shape[1] not found"
///            Fix: Add a shape[N] / picture[N] dispatch block before the cell fallback
///            that locates the TwoCellAnchor containing the shape/picture in the
///            worksheet's DrawingsPart and calls .Remove() on it.
///
/// NOT REPRODUCED (tests pass — agent reports may have been incorrect or stale):
///
///   Bug 2 (MEDIUM) — alignment.textRotation set succeeds but get doesn't return value
///            Investigation: ExcelStyleManager correctly writes TextRotation (line 195-196,
///            key "textrotation" after stripping "alignment." prefix). CellToNode reads
///            it back because it falls through ExcelStyleManager on Add and via Set.
///            Tests pass in the current codebase. Keeping as regression guard.
///
///   Bug 3 (MEDIUM) — font.size on cell add doesn't take effect
///            Investigation: ExcelStyleManager.IsStyleKey recognizes "font.size"
///            (StartsWith "font."), ApplyStyle extracts it with the "font." prefix
///            stripped to "size", and GetOrCreateFont handles it. CellToNode reads it
///            back at line 349-350. Tests pass in the current codebase. Keeping as
///            regression guard.
/// </summary>
public class ExcelRound5BugTests : IDisposable
{
    private readonly string _path;
    private ExcelHandler _handler;

    public ExcelRound5BugTests()
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

    // ─────────────────────────────────────────────────────────────────────────
    // Bug 1 (FAILING) — remove /Sheet1/shape[1] fails — shapes cannot be deleted
    // ─────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Remove must accept a shape[N] path without throwing.
    /// Currently Remove falls through to FindCell("shape[1]") which returns null,
    /// causing: System.ArgumentException: Cell shape[1] not found
    /// </summary>
    [Fact]
    public void RemoveShape_DoesNotThrow()
    {
        // Arrange: add a shape so there is something to remove
        _handler.Add("/Sheet1", "shape", null, new()
        {
            ["text"] = "DeleteMe",
            ["x"] = "1", ["y"] = "1", ["width"] = "3", ["height"] = "2"
        });

        // Verify the shape is there before removal
        var before = _handler.Get("/Sheet1/shape[1]");
        before.Should().NotBeNull("shape must exist before removal");

        // Act — this currently throws "Cell shape[1] not found"
        var act = () => _handler.Remove("/Sheet1/shape[1]");
        act.Should().NotThrow("Remove must accept a shape[N] path");
    }

    /// <summary>
    /// After removing a shape, Get("/Sheet1/shape[1]") must not find it.
    /// </summary>
    [Fact]
    public void RemoveShape_ShapeIsGoneAfterRemoval()
    {
        _handler.Add("/Sheet1", "shape", null, new()
        {
            ["text"] = "ToRemove",
            ["x"] = "0", ["y"] = "0", ["width"] = "3", ["height"] = "3"
        });

        _handler.Remove("/Sheet1/shape[1]");

        // After removal, shape[1] should not be found
        var after = _handler.Get("/Sheet1/shape[1]");
        after.Should().BeNull("shape should not exist after removal");
    }

    /// <summary>
    /// When two shapes exist and only the first is removed, the second
    /// must remain and become shape[1].
    /// </summary>
    [Fact]
    public void RemoveShape_RemainingShapeIsReindexed()
    {
        _handler.Add("/Sheet1", "shape", null, new()
        {
            ["text"] = "First",
            ["x"] = "0", ["y"] = "0", ["width"] = "2", ["height"] = "2"
        });
        _handler.Add("/Sheet1", "shape", null, new()
        {
            ["text"] = "Second",
            ["x"] = "3", ["y"] = "0", ["width"] = "2", ["height"] = "2"
        });

        _handler.Remove("/Sheet1/shape[1]");

        // The remaining shape (originally shape[2]) must now be shape[1]
        var remaining = _handler.Get("/Sheet1/shape[1]");
        remaining.Should().NotBeNull("the second shape must survive removal of the first");
        remaining!.Text.Should().Be("Second", "the remaining shape must be the one that was not removed");
    }

    /// <summary>
    /// Removal of a shape must persist across file reopen.
    /// </summary>
    [Fact]
    public void RemoveShape_PersistsAfterReopen()
    {
        _handler.Add("/Sheet1", "shape", null, new()
        {
            ["text"] = "EphemeralShape",
            ["x"] = "0", ["y"] = "0", ["width"] = "3", ["height"] = "3"
        });

        _handler.Remove("/Sheet1/shape[1]");
        Reopen();

        var after = _handler.Get("/Sheet1/shape[1]");
        after.Should().BeNull("removed shape must not reappear after reopen");
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Bug 2 (regression guard — currently passing)
    // alignment.textRotation Set + Get round-trip
    // ─────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// alignment.textRotation set via Set must be returned by Get.
    /// This is a regression guard — the feature works today and must stay working.
    /// </summary>
    [Fact]
    public void SetTextRotation_IsReturnedByGet()
    {
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Rotated" });
        _handler.Set("/Sheet1/A1", new() { ["alignment.textRotation"] = "45" });

        var node = _handler.Get("/Sheet1/A1");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("alignment.textRotation",
            "Get must expose textRotation that was written by Set");
        node.Format["alignment.textRotation"].ToString().Should().Be("45",
            "the rotation angle must match what was set");
    }

    /// <summary>
    /// alignment.textRotation must survive a file reopen.
    /// </summary>
    [Fact]
    public void SetTextRotation_PersistsAfterReopen()
    {
        _handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "B2", ["value"] = "Angled" });
        _handler.Set("/Sheet1/B2", new() { ["alignment.textRotation"] = "90" });
        Reopen();

        var node = _handler.Get("/Sheet1/B2");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("alignment.textRotation",
            "textRotation must survive a file reopen");
        node.Format["alignment.textRotation"].ToString().Should().Be("90");
    }

    /// <summary>
    /// alignment.textRotation can also be specified during Add.
    /// </summary>
    [Fact]
    public void AddCell_WithTextRotation_IsReturnedByGet()
    {
        _handler.Add("/Sheet1", "cell", null, new()
        {
            ["ref"] = "C3",
            ["value"] = "Tilted",
            ["alignment.textRotation"] = "30"
        });

        var node = _handler.Get("/Sheet1/C3");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("alignment.textRotation",
            "textRotation specified in Add must be readable back via Get");
        node.Format["alignment.textRotation"].ToString().Should().Be("30");
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Bug 3 (regression guard — currently passing)
    // font.size on cell Add
    // ─────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// font.size set during Add must be readable back via Get immediately.
    /// Regression guard — works today via ExcelStyleManager path.
    /// </summary>
    [Fact]
    public void AddCell_WithFontSize_IsReturnedByGet()
    {
        _handler.Add("/Sheet1", "cell", null, new()
        {
            ["ref"] = "A1",
            ["value"] = "BigText",
            ["font.size"] = "18"
        });

        var node = _handler.Get("/Sheet1/A1");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("font.size",
            "font.size specified during Add must be readable back immediately via Get");
        node.Format["font.size"].ToString().Should().Be("18pt",
            "font.size must be returned as pt-suffixed string matching the value set during Add");
    }

    /// <summary>
    /// font.size specified during Add must survive a file reopen.
    /// </summary>
    [Fact]
    public void AddCell_WithFontSize_PersistsAfterReopen()
    {
        _handler.Add("/Sheet1", "cell", null, new()
        {
            ["ref"] = "D4",
            ["value"] = "Persistent",
            ["font.size"] = "24"
        });
        Reopen();

        var node = _handler.Get("/Sheet1/D4");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("font.size",
            "font.size applied during Add must persist after reopen");
        node.Format["font.size"].ToString().Should().Be("24pt");
    }

    /// <summary>
    /// font.size specified during Add must produce the same result as Set after Add.
    /// </summary>
    [Fact]
    public void AddCell_WithFontSize_EquivalentToSetAfterAdd()
    {
        _handler.Add("/Sheet1", "cell", null, new()
        {
            ["ref"] = "A1", ["value"] = "ViaAdd", ["font.size"] = "16"
        });
        _handler.Add("/Sheet1", "cell", null, new()
        {
            ["ref"] = "B1", ["value"] = "ViaSet"
        });
        _handler.Set("/Sheet1/B1", new() { ["font.size"] = "16" });

        var nodeA = _handler.Get("/Sheet1/A1");
        var nodeB = _handler.Get("/Sheet1/B1");

        nodeA.Should().NotBeNull();
        nodeB.Should().NotBeNull();
        nodeA!.Format.Should().ContainKey("font.size",
            "Add with font.size must produce the same result as Set with font.size");
        nodeB!.Format.Should().ContainKey("font.size");
        nodeA.Format["font.size"].Should().Be(nodeB.Format["font.size"],
            "font.size via Add must be identical to font.size via Set");
    }
}
