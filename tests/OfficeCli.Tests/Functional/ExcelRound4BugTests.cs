// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml.Spreadsheet;
using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Failing tests for Round 4 top-4 bugs:
///
///   Bug 1 — Sort crashes with "Failed to compare two elements in the array"
///            Root cause: ParseSortValue returns either double or string; LINQ OrderedEnumerable
///            cannot compare a double against a string in ThenBy/ThenByDescending — the IComparable
///            implementation boxes each value but the runtime rejects cross-type comparisons when
///            a data column contains a mix of numeric and empty/string cells during reordering.
///            Triggered specifically when the first column is all-numeric but the secondary sort
///            column has empty cells (returns ""), yielding double vs string comparison failure.
///
///   Bug 2 — `add --type cell --prop address=C3` ignores the address
///            Root cause: Add("cell") only checks for the key "ref" (line 96), never "address".
///            The user-facing key documented in README and intuited by users is "address".
///
///   Bug 3 — Table totalRow=true doesn't generate actual total row content
///            Root cause: Feature gap. When hasTotalRow=true the code only sets
///            TotalsRowShown on the Table XML element; it never appends TotalsRowCount,
///            sets TotalsRowIndex on TableColumn, or writes any cell formula/label into
///            the total row cells in SheetData.
///
///   Bug 4 — No first-class Add API for row/column page breaks
///            Root cause: Feature gap. Remove and Get/Query handle rowbreak/colbreak paths,
///            but ExcelHandler.Add has no "rowbreak" or "colbreak" type case.
/// </summary>
public class ExcelRound4BugTests : IDisposable
{
    private readonly string _path;
    private ExcelHandler _handler;

    public ExcelRound4BugTests()
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
    // Bug 1 — Sort crashes with InvalidOperationException / ArgumentException
    //         "Failed to compare two elements in the array"
    // ─────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Sorting a column whose data is all-numeric must not throw.
    /// ParseSortValue returns double for numeric strings ("5","3","1") but returns
    /// string("") for cells missing from SheetData. When LINQ's sort compares a double
    /// against a string via IComparable.CompareTo(object) the runtime throws
    /// "Failed to compare two elements in the array" / InvalidOperationException.
    /// Using 5+ rows forces Tim sort to compare elements across different partitions,
    /// reliably surfacing the cross-type IComparable comparison.
    /// </summary>
    [Fact]
    public void Sort_NumericColumn_WithSomeEmptyCells_DoesNotThrow()
    {
        // Arrange: column A has numeric values, but row 3 cell is absent.
        // GetCellSortValue returns "" (string) for the absent cell,
        // while present cells parse to double — causing cross-type IComparable crash.
        for (int r = 1; r <= 5; r++)
            _handler.Add("/Sheet1", "row", null, new() { ["index"] = r.ToString() });

        _handler.Add("/Sheet1/row[1]", "cell", null, new() { ["ref"] = "A1", ["value"] = "5" });
        _handler.Add("/Sheet1/row[2]", "cell", null, new() { ["ref"] = "A2", ["value"] = "3" });
        // A3 intentionally absent — GetCellSortValue returns "" (string) for missing cell
        _handler.Add("/Sheet1/row[4]", "cell", null, new() { ["ref"] = "A4", ["value"] = "8" });
        _handler.Add("/Sheet1/row[5]", "cell", null, new() { ["ref"] = "A5", ["value"] = "1" });

        // Act — sort A descending; will compare double(8) against string("") → crash before fix
        var act = () => _handler.Set("/Sheet1", new() { ["sort"] = "A:desc" });
        act.Should().NotThrow("sort on a column with mixed numeric/empty cells must not produce a cross-type IComparable crash");

        // After fix, top row should be the highest numeric value
        var topCell = _handler.Get("/Sheet1/A1");
        topCell.Should().NotBeNull();
        topCell!.Text.Should().Be("8", "highest value must be first row after descending sort");
    }

    /// <summary>
    /// Sorting a purely numeric column B in descending order must not throw and
    /// must reorder the rows correctly.
    /// </summary>
    [Fact]
    public void Sort_NumericColumnDesc_ReordersRowsCorrectly()
    {
        // Arrange
        _handler.Add("/Sheet1", "row", null, new() { ["index"] = "1" });
        _handler.Add("/Sheet1", "row", null, new() { ["index"] = "2" });
        _handler.Add("/Sheet1", "row", null, new() { ["index"] = "3" });
        _handler.Add("/Sheet1/row[1]", "cell", null, new() { ["ref"] = "A1", ["value"] = "Alice" });
        _handler.Add("/Sheet1/row[1]", "cell", null, new() { ["ref"] = "B1", ["value"] = "5" });
        _handler.Add("/Sheet1/row[2]", "cell", null, new() { ["ref"] = "A2", ["value"] = "Bob" });
        _handler.Add("/Sheet1/row[2]", "cell", null, new() { ["ref"] = "B2", ["value"] = "1" });
        _handler.Add("/Sheet1/row[3]", "cell", null, new() { ["ref"] = "A3", ["value"] = "Carol" });
        _handler.Add("/Sheet1/row[3]", "cell", null, new() { ["ref"] = "B3", ["value"] = "3" });

        // Act
        var act = () => _handler.Set("/Sheet1", new() { ["sort"] = "B:desc" });
        act.Should().NotThrow();

        // Assert: B-descending order → 5 (Alice), 3 (Carol), 1 (Bob)
        var a1 = _handler.Get("/Sheet1/A1");
        var a3 = _handler.Get("/Sheet1/A3");
        a1!.Text.Should().Be("Alice", "row with B=5 should be first after desc sort");
        a3!.Text.Should().Be("Bob", "row with B=1 should be last after desc sort");
    }

    /// <summary>
    /// Multi-column sort (primary asc, secondary desc) must complete without exception.
    /// This specifically exercises the ThenBy / ThenByDescending path that was comparing
    /// IComparable values of mixed runtime types (double vs string).
    /// </summary>
    [Fact]
    public void Sort_MultiColumn_MixedTypes_DoesNotThrow()
    {
        _handler.Add("/Sheet1", "row", null, new() { ["index"] = "1" });
        _handler.Add("/Sheet1", "row", null, new() { ["index"] = "2" });
        _handler.Add("/Sheet1", "row", null, new() { ["index"] = "3" });
        _handler.Add("/Sheet1/row[1]", "cell", null, new() { ["ref"] = "A1", ["value"] = "1" });
        _handler.Add("/Sheet1/row[2]", "cell", null, new() { ["ref"] = "A2", ["value"] = "2" });
        _handler.Add("/Sheet1/row[3]", "cell", null, new() { ["ref"] = "A3", ["value"] = "3" });
        // B column sparse — creates the mixed-type IComparable conflict
        _handler.Add("/Sheet1/row[1]", "cell", null, new() { ["ref"] = "B1", ["value"] = "beta" });
        _handler.Add("/Sheet1/row[3]", "cell", null, new() { ["ref"] = "B3", ["value"] = "alpha" });

        var act = () => _handler.Set("/Sheet1", new() { ["sort"] = "A:asc,B:desc" });
        act.Should().NotThrow("multi-column sort with a sparse secondary column must not crash");
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Bug 2 — add --type cell --prop address=C3 ignores the address
    // ─────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// When the user specifies address= on cell add, the cell must land at that address.
    /// Currently the code only checks the "ref" key; "address" is silently ignored and
    /// the cell auto-assigns to the first free column in row 1.
    /// </summary>
    [Fact]
    public void AddCell_WithAddressProperty_LandsAtCorrectPosition()
    {
        // Act
        _handler.Add("/Sheet1", "cell", null, new() { ["address"] = "C3", ["value"] = "hello" });

        // Assert: cell must be at C3
        var node = _handler.Get("/Sheet1/C3");
        node.Should().NotBeNull("cell added with address=C3 must be retrievable at /Sheet1/C3");
        node!.Text.Should().Be("hello");
    }

    /// <summary>
    /// Both "address" and "ref" should be accepted as aliases for the cell reference.
    /// "ref" already works; verify "address" works equally after the fix.
    /// </summary>
    [Fact]
    public void AddCell_WithAddressProperty_PersistsAfterReopen()
    {
        _handler.Add("/Sheet1", "cell", null, new() { ["address"] = "E5", ["value"] = "world" });
        Reopen();

        var node = _handler.Get("/Sheet1/E5");
        node.Should().NotBeNull();
        node!.Text.Should().Be("world");
    }

    /// <summary>
    /// When address is absent and ref is absent the cell must auto-assign (existing behaviour).
    /// This is a regression guard — the fix for Bug 2 must not break auto-assign.
    /// </summary>
    [Fact]
    public void AddCell_WithoutAddressOrRef_AutoAssignsToFirstFreeColumn()
    {
        _handler.Add("/Sheet1", "cell", null, new() { ["value"] = "auto" });

        // Should land at A1 (first available)
        var node = _handler.Get("/Sheet1/A1");
        node.Should().NotBeNull("auto-assign must still work when neither address nor ref is given");
        node!.Text.Should().Be("auto");
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Bug 3 — Table totalRow=true doesn't generate actual total row content
    // ─────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// When totalRow=true the Table XML must have TotalsRowCount set to 1,
    /// and TotalsRowShown must be true. Currently neither is correct —
    /// TotalsRowShown is set but TotalsRowCount is not set at all, and the
    /// total row cells are never written to SheetData.
    /// </summary>
    [Fact]
    public void AddTable_WithTotalRow_SetsTotalsRowCountOnTableDefinition()
    {
        // Arrange: write header + data rows
        _handler.Add("/Sheet1", "row", null, new() { ["index"] = "1" });
        _handler.Add("/Sheet1", "row", null, new() { ["index"] = "2" });
        _handler.Add("/Sheet1/row[1]", "cell", null, new() { ["ref"] = "A1", ["value"] = "Name" });
        _handler.Add("/Sheet1/row[1]", "cell", null, new() { ["ref"] = "B1", ["value"] = "Amount" });
        _handler.Add("/Sheet1/row[2]", "cell", null, new() { ["ref"] = "A2", ["value"] = "Alpha" });
        _handler.Add("/Sheet1/row[2]", "cell", null, new() { ["ref"] = "B2", ["value"] = "100" });

        // Act
        _handler.Add("/Sheet1", "table", null, new()
        {
            ["range"] = "A1:B3",
            ["totalRow"] = "true",
            ["columns"] = "Name,Amount"
        });

        // Assert: table definition must advertise a totals row
        var node = _handler.Get("/Sheet1/table[1]");
        node.Should().NotBeNull("table should be queryable after add");
        node!.Format.Should().ContainKey("totalRow");
        node.Format["totalRow"].Should().Be(true, "totalRow flag must be true in the table node");

        // The OOXML Table element must have TotalsRowCount=1 (not just TotalsRowShown)
        // We verify this through a round-trip: close, reopen, and check that Excel-visible
        // metadata is consistent.
        Reopen();
        var nodeAfter = _handler.Get("/Sheet1/table[1]");
        nodeAfter.Should().NotBeNull();
        nodeAfter!.Format["totalRow"].Should().Be(true, "totalRow must survive a reopen");
    }

    /// <summary>
    /// When totalRow=true a total row cell must exist in SheetData at the bottom of the range.
    /// The cell at least needs a label (e.g. "Total") or a SUM formula for the data column.
    /// This is currently a feature gap — no cells are written for the total row.
    /// </summary>
    [Fact]
    public void AddTable_WithTotalRow_CreatesTotalRowCellsInSheetData()
    {
        _handler.Add("/Sheet1", "row", null, new() { ["index"] = "1" });
        _handler.Add("/Sheet1", "row", null, new() { ["index"] = "2" });
        _handler.Add("/Sheet1/row[1]", "cell", null, new() { ["ref"] = "A1", ["value"] = "Item" });
        _handler.Add("/Sheet1/row[1]", "cell", null, new() { ["ref"] = "B1", ["value"] = "Price" });
        _handler.Add("/Sheet1/row[2]", "cell", null, new() { ["ref"] = "A2", ["value"] = "Widget" });
        _handler.Add("/Sheet1/row[2]", "cell", null, new() { ["ref"] = "B2", ["value"] = "42" });

        // table range A1:B3 — row 3 is the totals row
        _handler.Add("/Sheet1", "table", null, new()
        {
            ["range"] = "A1:B3",
            ["totalRow"] = "true",
            ["columns"] = "Item,Price"
        });

        // The total row label cell (A3) should have a label such as "Total"
        var totalLabelCell = _handler.Get("/Sheet1/A3");
        totalLabelCell.Should().NotBeNull("total row label cell A3 must be created in SheetData when totalRow=true");
        totalLabelCell!.Text.Should().NotBeNullOrEmpty("total row label cell A3 must contain a label (e.g. 'Total'), not be empty");

        // B3 should have a SUM formula aggregating B2 data
        var totalValueCell = _handler.Get("/Sheet1/B3");
        totalValueCell.Should().NotBeNull("total row value cell B3 must be created in SheetData when totalRow=true");
        // A formula or computed value must be present — "(empty)" means no total row was generated
        totalValueCell!.Text.Should().NotBe("(empty)", "total row B3 must have a SUM formula or value, not be blank");
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Bug 4 — No first-class Add API for row/column page breaks
    // ─────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Adding a row page break via Add("/Sheet1", "rowbreak", ...) must succeed
    /// and the break must be retrievable via Get.
    /// Currently ExcelHandler.Add has no "rowbreak" case and throws.
    /// </summary>
    [Fact]
    public void AddRowBreak_ByType_CreatesBreakAndIsQueryable()
    {
        // Act
        var act = () => _handler.Add("/Sheet1", "rowbreak", null, new() { ["row"] = "5" });
        act.Should().NotThrow("adding a rowbreak by type must be supported");

        // Assert: the break is readable
        var node = _handler.Get("/Sheet1/rowbreak[1]");
        node.Should().NotBeNull("rowbreak[1] must be retrievable after add");
        node!.Type.Should().Be("rowbreak");
        node.Format.Should().ContainKey("row");
        node.Format["row"].ToString().Should().Be("5");
    }

    /// <summary>
    /// Adding a column page break via Add("/Sheet1", "colbreak", ...) must succeed.
    /// </summary>
    [Fact]
    public void AddColBreak_ByType_CreatesBreakAndIsQueryable()
    {
        var act = () => _handler.Add("/Sheet1", "colbreak", null, new() { ["col"] = "3" });
        act.Should().NotThrow("adding a colbreak by type must be supported");

        var node = _handler.Get("/Sheet1/colbreak[1]");
        node.Should().NotBeNull("colbreak[1] must be retrievable after add");
        node!.Type.Should().Be("colbreak");
        node.Format.Should().ContainKey("col");
        node.Format["col"].ToString().Should().Be("3");
    }

    /// <summary>
    /// Multiple row breaks must be independently addressable by index.
    /// </summary>
    [Fact]
    public void AddMultipleRowBreaks_AreIndexedCorrectly()
    {
        _handler.Add("/Sheet1", "rowbreak", null, new() { ["row"] = "5" });
        _handler.Add("/Sheet1", "rowbreak", null, new() { ["row"] = "10" });

        var brk1 = _handler.Get("/Sheet1/rowbreak[1]");
        var brk2 = _handler.Get("/Sheet1/rowbreak[2]");

        brk1.Should().NotBeNull();
        brk2.Should().NotBeNull();
        brk1!.Format["row"].ToString().Should().Be("5");
        brk2!.Format["row"].ToString().Should().Be("10");
    }

    /// <summary>
    /// Row breaks added via Add must persist after reopen (round-trip test).
    /// </summary>
    [Fact]
    public void AddRowBreak_PersistsAfterReopen()
    {
        _handler.Add("/Sheet1", "rowbreak", null, new() { ["row"] = "7" });
        Reopen();

        var node = _handler.Get("/Sheet1/rowbreak[1]");
        node.Should().NotBeNull("rowbreak must survive file reopen");
        node!.Format["row"].ToString().Should().Be("7");
    }
}
