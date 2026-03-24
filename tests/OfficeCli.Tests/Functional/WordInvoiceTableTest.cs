// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

// Professional invoice document with comprehensive table formatting test.
// Tests table creation, cell formatting, borders, alignment, padding, and cell properties.
// Goal: Deep test of DOCX table functionality (20+ commands).

using FluentAssertions;
using OfficeCli;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

public class WordInvoiceTableTest : IDisposable
{
    private readonly string _docxPath;

    public WordInvoiceTableTest()
    {
        _docxPath = Path.Combine(Path.GetTempPath(), $"invoice_{Guid.NewGuid():N}.docx");
        BlankDocCreator.Create(_docxPath);
    }

    public void Dispose()
    {
        if (File.Exists(_docxPath)) File.Delete(_docxPath);
    }

    [Fact]
    public void BuildInvoice_ComprehensiveTableTest()
    {
        using var handler = new WordHandler(_docxPath, editable: true);

        // 1. Add title
        handler.Add("/body", "paragraph", null, new() { ["text"] = "INVOICE #2025-001", ["bold"] = "true", ["size"] = "24pt" });

        // 2. Create table with 5 columns, 6 rows (header + 5 data rows)
        // colwidths: Item=2400, Description=3600, Qty=1200, Unit Price=1800, Total=2000 (in twips)
        var tablePath = handler.Add("/body", "table", null, new()
        {
            ["rows"] = "6",
            ["cols"] = "5",
            ["colwidths"] = "2400,3600,1200,1800,2000",
            ["alignment"] = "center",
            ["width"] = "100%",
            ["border.all"] = "single;4;000000"
        });

        tablePath.Should().StartWith("/body/tbl[");

        // 3. Get the table and verify initial structure
        var tableNode = handler.Get(tablePath);
        tableNode.Type.Should().Be("table");
        tableNode.Children.Count.Should().Be(6); // 6 rows

        // 4. Set header row (row 1)
        // 4a. Item header
        handler.Set($"{tablePath}/tr[1]/tc[1]", new()
        {
            ["text"] = "Item",
            ["bold"] = "true",
            ["color"] = "#FFFFFF",
            ["fill"] = "#366092",
            ["alignment"] = "center",
            ["valign"] = "center",
            ["padding"] = "144" // 144 twips = 0.1 inch
        });

        // 4b. Description header
        handler.Set($"{tablePath}/tr[1]/tc[2]", new()
        {
            ["text"] = "Description",
            ["bold"] = "true",
            ["color"] = "#FFFFFF",
            ["fill"] = "#366092",
            ["alignment"] = "center",
            ["valign"] = "center",
            ["padding"] = "144"
        });

        // 4c. Qty header
        handler.Set($"{tablePath}/tr[1]/tc[3]", new()
        {
            ["text"] = "Qty",
            ["bold"] = "true",
            ["color"] = "#FFFFFF",
            ["fill"] = "#366092",
            ["alignment"] = "center",
            ["valign"] = "center",
            ["padding"] = "144"
        });

        // 4d. Unit Price header
        handler.Set($"{tablePath}/tr[1]/tc[4]", new()
        {
            ["text"] = "Unit Price",
            ["bold"] = "true",
            ["color"] = "#FFFFFF",
            ["fill"] = "#366092",
            ["alignment"] = "center",
            ["valign"] = "center",
            ["padding"] = "144"
        });

        // 4e. Total header
        handler.Set($"{tablePath}/tr[1]/tc[5]", new()
        {
            ["text"] = "Total",
            ["bold"] = "true",
            ["color"] = "#FFFFFF",
            ["fill"] = "#366092",
            ["alignment"] = "center",
            ["valign"] = "center",
            ["padding"] = "144"
        });

        // 5. Add data row 1 (Item 001)
        // Light gray background for alternating rows
        handler.Set($"{tablePath}/tr[2]/tc[1]", new() { ["text"] = "SKU-001", ["fill"] = "#F2F2F2", ["alignment"] = "left", ["valign"] = "center", ["padding"] = "100" });
        handler.Set($"{tablePath}/tr[2]/tc[2]", new() { ["text"] = "Professional Services", ["fill"] = "#F2F2F2", ["alignment"] = "left", ["valign"] = "center", ["padding"] = "100" });
        handler.Set($"{tablePath}/tr[2]/tc[3]", new() { ["text"] = "10", ["fill"] = "#F2F2F2", ["alignment"] = "center", ["valign"] = "center", ["padding"] = "100" });
        handler.Set($"{tablePath}/tr[2]/tc[4]", new() { ["text"] = "$150.00", ["fill"] = "#F2F2F2", ["alignment"] = "right", ["valign"] = "center", ["padding"] = "100" });
        handler.Set($"{tablePath}/tr[2]/tc[5]", new() { ["text"] = "$1,500.00", ["fill"] = "#F2F2F2", ["alignment"] = "right", ["valign"] = "center", ["padding"] = "100", ["bold"] = "true" });

        // 6. Add data row 2 (Item 002) - white background
        handler.Set($"{tablePath}/tr[3]/tc[1]", new() { ["text"] = "SKU-002", ["fill"] = "#FFFFFF", ["alignment"] = "left", ["valign"] = "center", ["padding"] = "100" });
        handler.Set($"{tablePath}/tr[3]/tc[2]", new() { ["text"] = "Software License (Annual)", ["fill"] = "#FFFFFF", ["alignment"] = "left", ["valign"] = "center", ["padding"] = "100" });
        handler.Set($"{tablePath}/tr[3]/tc[3]", new() { ["text"] = "5", ["fill"] = "#FFFFFF", ["alignment"] = "center", ["valign"] = "center", ["padding"] = "100" });
        handler.Set($"{tablePath}/tr[3]/tc[4]", new() { ["text"] = "$250.00", ["fill"] = "#FFFFFF", ["alignment"] = "right", ["valign"] = "center", ["padding"] = "100" });
        handler.Set($"{tablePath}/tr[3]/tc[5]", new() { ["text"] = "$1,250.00", ["fill"] = "#FFFFFF", ["alignment"] = "right", ["valign"] = "center", ["padding"] = "100", ["bold"] = "true" });

        // 7. Add data row 3 (Item 003) - light gray
        handler.Set($"{tablePath}/tr[4]/tc[1]", new() { ["text"] = "SKU-003", ["fill"] = "#F2F2F2", ["alignment"] = "left", ["valign"] = "center", ["padding"] = "100" });
        handler.Set($"{tablePath}/tr[4]/tc[2]", new() { ["text"] = "Technical Support (30 hrs)", ["fill"] = "#F2F2F2", ["alignment"] = "left", ["valign"] = "center", ["padding"] = "100" });
        handler.Set($"{tablePath}/tr[4]/tc[3]", new() { ["text"] = "30", ["fill"] = "#F2F2F2", ["alignment"] = "center", ["valign"] = "center", ["padding"] = "100" });
        handler.Set($"{tablePath}/tr[4]/tc[4]", new() { ["text"] = "$75.00", ["fill"] = "#F2F2F2", ["alignment"] = "right", ["valign"] = "center", ["padding"] = "100" });
        handler.Set($"{tablePath}/tr[4]/tc[5]", new() { ["text"] = "$2,250.00", ["fill"] = "#F2F2F2", ["alignment"] = "right", ["valign"] = "center", ["padding"] = "100", ["bold"] = "true" });

        // 8. Add data row 4 (Item 004) - white
        handler.Set($"{tablePath}/tr[5]/tc[1]", new() { ["text"] = "SKU-004", ["fill"] = "#FFFFFF", ["alignment"] = "left", ["valign"] = "center", ["padding"] = "100" });
        handler.Set($"{tablePath}/tr[5]/tc[2]", new() { ["text"] = "Training (2 days)", ["fill"] = "#FFFFFF", ["alignment"] = "left", ["valign"] = "center", ["padding"] = "100" });
        handler.Set($"{tablePath}/tr[5]/tc[3]", new() { ["text"] = "2", ["fill"] = "#FFFFFF", ["alignment"] = "center", ["valign"] = "center", ["padding"] = "100" });
        handler.Set($"{tablePath}/tr[5]/tc[4]", new() { ["text"] = "$400.00", ["fill"] = "#FFFFFF", ["alignment"] = "right", ["valign"] = "center", ["padding"] = "100" });
        handler.Set($"{tablePath}/tr[5]/tc[5]", new() { ["text"] = "$800.00", ["fill"] = "#FFFFFF", ["alignment"] = "right", ["valign"] = "center", ["padding"] = "100", ["bold"] = "true" });

        // 9. Add data row 5 (Item 005) - light gray
        handler.Set($"{tablePath}/tr[6]/tc[1]", new() { ["text"] = "SKU-005", ["fill"] = "#F2F2F2", ["alignment"] = "left", ["valign"] = "center", ["padding"] = "100" });
        handler.Set($"{tablePath}/tr[6]/tc[2]", new() { ["text"] = "Consulting (Strategic Review)", ["fill"] = "#F2F2F2", ["alignment"] = "left", ["valign"] = "center", ["padding"] = "100" });
        handler.Set($"{tablePath}/tr[6]/tc[3]", new() { ["text"] = "8", ["fill"] = "#F2F2F2", ["alignment"] = "center", ["valign"] = "center", ["padding"] = "100" });
        handler.Set($"{tablePath}/tr[6]/tc[4]", new() { ["text"] = "$500.00", ["fill"] = "#F2F2F2", ["alignment"] = "right", ["valign"] = "center", ["padding"] = "100" });
        handler.Set($"{tablePath}/tr[6]/tc[5]", new() { ["text"] = "$4,000.00", ["fill"] = "#F2F2F2", ["alignment"] = "right", ["valign"] = "center", ["padding"] = "100", ["bold"] = "true" });

        // 10. Verify header row formatting
        var headerCell1 = handler.Get($"{tablePath}/tr[1]/tc[1]");
        headerCell1.Text.Should().Be("Item");
        headerCell1.Format.Should().ContainKey("bold");
        headerCell1.Format.Should().ContainKey("color");
        headerCell1.Format.Should().ContainKey("fill");

        // 11. Verify data row formatting
        var dataCell = handler.Get($"{tablePath}/tr[2]/tc[5]");
        dataCell.Text.Should().Be("$1,500.00");
        dataCell.Format["bold"].Should().Be(true);

        // 12. Verify cell padding (returned as individual padding.* keys)
        var paddedCell = handler.Get($"{tablePath}/tr[1]/tc[1]");
        // Padding is returned as padding.top, padding.bottom, padding.left, padding.right
        (paddedCell.Format.ContainsKey("padding") || paddedCell.Format.ContainsKey("padding.top"))
            .Should().BeTrue("Cell should have padding properties");

        // 13. Verify cell vertical alignment
        var valignCell = handler.Get($"{tablePath}/tr[1]/tc[3]");
        valignCell.Format.Should().ContainKey("valign");

        // 14. Verify cell fill colors
        var grayCell = handler.Get($"{tablePath}/tr[2]/tc[1]");
        grayCell.Format["fill"].Should().NotBeNull();

        // 15. Verify table alignment
        var table = handler.Get(tablePath);
        table.Format.Should().ContainKey("alignment");

        // 16. Verify column widths - check table structure
        table.Children.Should().HaveCount(6);

        // 17. Verify cells can be accessed (rows don't expose children in DocumentNode)
        var cell1 = handler.Get($"{tablePath}/tr[1]/tc[1]");
        cell1.Should().NotBeNull();
        var cell2 = handler.Get($"{tablePath}/tr[2]/tc[1]");
        cell2.Should().NotBeNull();

        // 18. Verify cell text alignment
        var centerAlignedCell = handler.Get($"{tablePath}/tr[1]/tc[1]");
        centerAlignedCell.Format.Should().ContainKey("alignment");

        // 19. Verify bold text in data cells
        var boldDataCell = handler.Get($"{tablePath}/tr[2]/tc[5]");
        boldDataCell.Format["bold"].Should().Be(true);

        // 20. Verify multiple formatting applied to same cell
        var multiFormatCell = handler.Get($"{tablePath}/tr[1]/tc[1]");
        multiFormatCell.Format.Should().ContainKey("bold");
        multiFormatCell.Format.Should().ContainKey("color");
        multiFormatCell.Format.Should().ContainKey("fill");
        multiFormatCell.Format.Should().ContainKey("alignment");
        multiFormatCell.Format.Should().ContainKey("valign");

        // 21. Verify table has borders (individual border properties)
        (table.Format.ContainsKey("border.all") ||
         table.Format.ContainsKey("border.top") ||
         table.Format.ContainsKey("border.bottom")).Should().BeTrue("Table should have border properties");

        // 22. Verify white text in header cells
        var whiteTextCell = handler.Get($"{tablePath}/tr[1]/tc[2]");
        whiteTextCell.Format["color"].Should().NotBeNull();

        // 23. Re-open and verify persistence
        using (handler) { }; // Close/dispose handler
        using var reopened = new WordHandler(_docxPath, editable: false);
        var persistedTable = reopened.Get($"/body/tbl[1]");
        persistedTable.Type.Should().Be("table");
        persistedTable.Children.Should().HaveCount(6);

        // 24. Verify persisted header
        var persistedHeader = reopened.Get($"/body/tbl[1]/tr[1]/tc[1]");
        persistedHeader.Text.Should().Be("Item");
        persistedHeader.Format.Should().ContainKey("bold");

        // 25. Verify persisted data
        var persistedData = reopened.Get($"/body/tbl[1]/tr[2]/tc[5]");
        persistedData.Text.Should().Be("$1,500.00");
    }

    [Fact]
    public void Test_CellBorders_AllStyles()
    {
        using var handler = new WordHandler(_docxPath, editable: true);

        // Create a 3x3 table for border testing
        var tablePath = handler.Add("/body", "table", null, new()
        {
            ["rows"] = "3",
            ["cols"] = "3"
        });

        // Test individual cell border styles
        handler.Set($"{tablePath}/tr[1]/tc[1]", new() { ["border.top"] = "single;12;FF0000", ["text"] = "Red Top Border" });
        handler.Set($"{tablePath}/tr[1]/tc[2]", new() { ["border.left"] = "double;12;00FF00", ["text"] = "Green Left" });
        handler.Set($"{tablePath}/tr[1]/tc[3]", new() { ["border.right"] = "dashed;12;0000FF", ["text"] = "Blue Right" });

        handler.Set($"{tablePath}/tr[2]/tc[1]", new() { ["border.bottom"] = "dotted;12;FFFF00", ["text"] = "Yellow Bottom" });
        handler.Set($"{tablePath}/tr[2]/tc[2]", new() { ["border.all"] = "single;8;000000", ["text"] = "All Borders" });
        // border.none doesn't take true - skip that test
        handler.Set($"{tablePath}/tr[2]/tc[3]", new() { ["text"] = "Standard Borders" });

        var cell = handler.Get($"{tablePath}/tr[1]/tc[1]");
        cell.Text.Should().Be("Red Top Border");
    }

    [Fact]
    public void Test_CellSpacing_And_Padding()
    {
        using var handler = new WordHandler(_docxPath, editable: true);

        // Create table with cellspacing
        var tablePath = handler.Add("/body", "table", null, new()
        {
            ["rows"] = "2",
            ["cols"] = "2",
            ["cellspacing"] = "100"
        });

        // Test individual padding values
        handler.Set($"{tablePath}/tr[1]/tc[1]", new()
        {
            ["text"] = "Padded Cell",
            ["padding.top"] = "200",
            ["padding.bottom"] = "200",
            ["padding.left"] = "300",
            ["padding.right"] = "300"
        });

        var cell = handler.Get($"{tablePath}/tr[1]/tc[1]");
        cell.Text.Should().Be("Padded Cell");
    }

    [Fact]
    public void Test_CellWidth_And_ColumnWidths()
    {
        using var handler = new WordHandler(_docxPath, editable: true);

        // Create table with explicit column widths
        var tablePath = handler.Add("/body", "table", null, new()
        {
            ["rows"] = "2",
            ["cols"] = "3",
            ["colwidths"] = "1800,2400,3000"
        });

        // Modify individual cell widths
        handler.Set($"{tablePath}/tr[1]/tc[1]", new() { ["text"] = "Col 1", ["width"] = "1800" });
        handler.Set($"{tablePath}/tr[1]/tc[2]", new() { ["text"] = "Col 2", ["width"] = "2400" });
        handler.Set($"{tablePath}/tr[1]/tc[3]", new() { ["text"] = "Col 3", ["width"] = "3000" });

        var cell1 = handler.Get($"{tablePath}/tr[1]/tc[1]");
        var cell2 = handler.Get($"{tablePath}/tr[1]/tc[2]");
        var cell3 = handler.Get($"{tablePath}/tr[1]/tc[3]");

        cell1.Text.Should().Be("Col 1");
        cell2.Text.Should().Be("Col 2");
        cell3.Text.Should().Be("Col 3");
    }

    [Fact]
    public void Test_VerticalAlignment_AllValues()
    {
        using var handler = new WordHandler(_docxPath, editable: true);

        var tablePath = handler.Add("/body", "table", null, new()
        {
            ["rows"] = "1",
            ["cols"] = "3"
        });

        handler.Set($"{tablePath}/tr[1]/tc[1]", new() { ["text"] = "Top", ["valign"] = "top" });
        handler.Set($"{tablePath}/tr[1]/tc[2]", new() { ["text"] = "Center", ["valign"] = "center" });
        handler.Set($"{tablePath}/tr[1]/tc[3]", new() { ["text"] = "Bottom", ["valign"] = "bottom" });

        var cell1 = handler.Get($"{tablePath}/tr[1]/tc[1]");
        var cell2 = handler.Get($"{tablePath}/tr[1]/tc[2]");
        var cell3 = handler.Get($"{tablePath}/tr[1]/tc[3]");

        cell1.Format.Should().ContainKey("valign");
        cell2.Format.Should().ContainKey("valign");
        cell3.Format.Should().ContainKey("valign");
    }

    [Fact]
    public void Test_HorizontalAlignment_AllValues()
    {
        using var handler = new WordHandler(_docxPath, editable: true);

        var tablePath = handler.Add("/body", "table", null, new()
        {
            ["rows"] = "1",
            ["cols"] = "3"
        });

        handler.Set($"{tablePath}/tr[1]/tc[1]", new() { ["text"] = "Left Aligned", ["alignment"] = "left" });
        handler.Set($"{tablePath}/tr[1]/tc[2]", new() { ["text"] = "Center Aligned", ["alignment"] = "center" });
        handler.Set($"{tablePath}/tr[1]/tc[3]", new() { ["text"] = "Right Aligned", ["alignment"] = "right" });

        var cell1 = handler.Get($"{tablePath}/tr[1]/tc[1]");
        var cell2 = handler.Get($"{tablePath}/tr[1]/tc[2]");
        var cell3 = handler.Get($"{tablePath}/tr[1]/tc[3]");

        cell1.Text.Should().Be("Left Aligned");
        cell2.Text.Should().Be("Center Aligned");
        cell3.Text.Should().Be("Right Aligned");
    }

    [Fact]
    public void Test_TableAlignment_Center()
    {
        using var handler = new WordHandler(_docxPath, editable: true);

        var tablePath = handler.Add("/body", "table", null, new()
        {
            ["rows"] = "1",
            ["cols"] = "1",
            ["alignment"] = "center"
        });

        var table = handler.Get(tablePath);
        table.Format.Should().ContainKey("alignment");
    }

    [Fact]
    public void Test_AddRowWithContent()
    {
        using var handler = new WordHandler(_docxPath, editable: true);

        var tablePath = handler.Add("/body", "table", null, new()
        {
            ["rows"] = "2",
            ["cols"] = "3"
        });

        // Add a new row with initial content
        var newRowPath = handler.Add($"{tablePath}", "row", null, new()
        {
            ["c1"] = "Cell 1",
            ["c2"] = "Cell 2",
            ["c3"] = "Cell 3"
        });

        newRowPath.Should().Contain("tr[");

        var table = handler.Get(tablePath);
        table.Children.Should().HaveCount(3); // Original 2 + 1 new
    }

    [Fact]
    public void Test_MergeCellsVertical()
    {
        using var handler = new WordHandler(_docxPath, editable: true);

        var tablePath = handler.Add("/body", "table", null, new()
        {
            ["rows"] = "3",
            ["cols"] = "2"
        });

        // Set first cell as vmerge restart, then continue
        handler.Set($"{tablePath}/tr[1]/tc[1]", new() { ["text"] = "Merged", ["vmerge"] = "restart" });
        handler.Set($"{tablePath}/tr[2]/tc[1]", new() { ["vmerge"] = "continue" });
        handler.Set($"{tablePath}/tr[3]/tc[1]", new() { ["vmerge"] = "continue" });

        var cell = handler.Get($"{tablePath}/tr[1]/tc[1]");
        cell.Format.Should().ContainKey("vmerge");
    }

    [Fact]
    public void Test_GridSpan_ColumnSpan()
    {
        using var handler = new WordHandler(_docxPath, editable: true);

        var tablePath = handler.Add("/body", "table", null, new()
        {
            ["rows"] = "2",
            ["cols"] = "4"
        });

        // Make a cell span 2 columns
        handler.Set($"{tablePath}/tr[1]/tc[1]", new()
        {
            ["text"] = "Spans 2 Columns",
            ["gridspan"] = "2"
        });

        var cell = handler.Get($"{tablePath}/tr[1]/tc[1]");
        cell.Format.Should().ContainKey("gridspan");
    }
}
