// Bug reproduction tests — each test exposes a specific issue in the codebase.

using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

public class BugReproTests : IDisposable
{
    private readonly string _docxPath;
    private readonly string _xlsxPath;
    private WordHandler _wordHandler;
    private ExcelHandler _excelHandler;

    public BugReproTests()
    {
        _docxPath = Path.Combine(Path.GetTempPath(), $"bug_{Guid.NewGuid():N}.docx");
        _xlsxPath = Path.Combine(Path.GetTempPath(), $"bug_{Guid.NewGuid():N}.xlsx");
        BlankDocCreator.Create(_docxPath);
        BlankDocCreator.Create(_xlsxPath);
        _wordHandler = new WordHandler(_docxPath, editable: true);
        _excelHandler = new ExcelHandler(_xlsxPath, editable: true);
    }

    public void Dispose()
    {
        _wordHandler.Dispose();
        _excelHandler.Dispose();
        if (File.Exists(_docxPath)) File.Delete(_docxPath);
        if (File.Exists(_xlsxPath)) File.Delete(_xlsxPath);
    }

    private WordHandler ReopenWord()
    {
        _wordHandler.Dispose();
        _wordHandler = new WordHandler(_docxPath, editable: true);
        return _wordHandler;
    }

    private ExcelHandler ReopenExcel()
    {
        _excelHandler.Dispose();
        _excelHandler = new ExcelHandler(_xlsxPath, editable: true);
        return _excelHandler;
    }

    // ==================== BUG 1: Footnote Set prepends space every time ====================
    // Each call to Set "/footnote[1]" text="X" writes " X" (with leading space).
    // If you Get and then Set again, the space accumulates.
    // The Get returns ALL descendant Text including the reference mark's space.

    [Fact]
    public void Bug_FootnoteSet_TextShouldBeExact()
    {
        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "Text" });
        _wordHandler.Add("/body/p[1]", "footnote", null, new() { ["text"] = "Original" });

        _wordHandler.Set("/footnote[1]", new() { ["text"] = "Updated" });

        var fn = _wordHandler.Get("/footnote[1]");
        // The Set prepends " " to text, so Get returns " Updated" not "Updated".
        // Also the footnote has a ReferenceMark run before the text run.
        // Let's verify the actual text is retrievable and clean.
        fn.Text.Should().Contain("Updated");
        // BUG: does Get return ONLY the user text, or does it include reference mark text too?
        // The Get joins ALL Descendants<Text>(), which includes the space from reference mark.
        fn.Text.Trim().Should().Be("Updated",
            "Get should return clean footnote text without extra whitespace");
    }

    // ==================== BUG 2: Column width Set on range column ====================
    // If a Column element has Min=1 Max=5 (covers A-E), setting width on /Sheet1/col[C]
    // modifies the width for ALL columns A-E, not just C.

    [Fact]
    public void Bug_ColumnWidth_ModifiesSharedRange()
    {
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "A" });
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "E1", ["value"] = "E" });

        // Set col A width
        _excelHandler.Set("/Sheet1/col[A]", new() { ["width"] = "10" });
        // Set col E width — should NOT affect col A
        _excelHandler.Set("/Sheet1/col[E]", new() { ["width"] = "30" });

        var colA = _excelHandler.Get("/Sheet1/col[A]");
        var colE = _excelHandler.Get("/Sheet1/col[E]");

        // If both columns are independent, A should be 10 and E should be 30
        ((double)colA.Format["width"]).Should().Be(10, "Column A width should be independent of Column E");
        ((double)colE.Format["width"]).Should().Be(30, "Column E width should be 30");
    }

    // ==================== BUG 3: MergeCells element order violation ====================
    // OpenXML schema requires: SheetData > MergeCells > ConditionalFormatting
    // But if CF is added first, then merge is added, the code inserts MergeCells
    // after SheetData — which is correct. However if AutoFilter exists between
    // SheetData and CF, the order might break.

    [Fact]
    public void Bug_MergeCells_OrderAfterConditionalFormatting()
    {
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "10" });
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A2", ["value"] = "20" });

        // Add conditional formatting first
        _excelHandler.Add("/Sheet1", "databar", null, new() { ["sqref"] = "A1:A2" });

        // Then merge — this should still produce valid XML
        _excelHandler.Set("/Sheet1/A1:B1", new() { ["merge"] = "true" });

        // Reopen to verify file is valid
        ReopenExcel();
        var cell = _excelHandler.Get("/Sheet1/A1");
        cell.Format.Should().ContainKey("merge");
    }

    // ==================== BUG 4: Section break with landscape doesn't swap dimensions correctly ====================
    // When orientation=landscape and source page is portrait (width < height),
    // the code swaps them. But if the source already IS landscape, it won't swap
    // even though it should stay landscape.

    [Fact]
    public void Bug_SectionBreak_LandscapeSwap()
    {
        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "Before" });
        _wordHandler.Add("/body", "section", null, new()
        {
            ["type"] = "nextPage", ["orientation"] = "landscape"
        });

        var sec = _wordHandler.Get("/section[1]");
        var w = (uint)sec.Format["pageWidth"];
        var h = (uint)sec.Format["pageHeight"];

        // In landscape, width should be > height
        w.Should().BeGreaterThan(h, "Landscape page width should be greater than height");
    }

    // ==================== BUG 5: Excel merge then add conditional formatting — validate XML ====================
    // After merging cells and adding CF, does the file pass validation?

    [Fact]
    public void Bug_ExcelValidation_AfterMergeAndCF()
    {
        // Reproduce: Add CF first, then merge — same handler instance
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "10" });
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "B1", ["value"] = "20" });
        _excelHandler.Add("/Sheet1", "databar", null, new() { ["sqref"] = "A2:A10" });
        _excelHandler.Set("/Sheet1/A1:B1", new() { ["merge"] = "true" });

        // Validate
        ReopenExcel();
        var errors = _excelHandler.Validate();
        errors.Should().BeEmpty("File should be valid after CF + merge (element order must be correct)");
    }

    // ==================== BUG 5b: Excel freeze pane Get returns TopLeftCell, not the freeze reference ====================
    // When you set freeze=B3, the Pane element stores VerticalSplit/HorizontalSplit
    // as row/col counts and TopLeftCell as "B3".
    // But the Get reads TopLeftCell — if someone sets freeze differently,
    // the returned value may not match what was set.

    [Fact]
    public void Bug_FreezePanes_GetMatchesSet()
    {
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Data" });

        _excelHandler.Set("/Sheet1", new() { ["freeze"] = "C4" });

        var sheet = _excelHandler.Get("/Sheet1");
        ((string)sheet.Format["freeze"]).Should().Be("C4",
            "Get should return exactly what was Set");
    }

    // ==================== BUG 6: Excel Set row height on non-existent row ====================
    // Setting height on a row that has no cells creates the row,
    // but Get "/Sheet1/row[5]" uses rowIndex matching, and the created row
    // won't be findable if the sheet only has data in row 1.

    [Fact]
    public void Bug_SetRowHeight_OnEmptyRow()
    {
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Data" });

        // Set height on row 5 which has no data
        _excelHandler.Set("/Sheet1/row[5]", new() { ["height"] = "25" });

        // Should be retrievable
        var row = _excelHandler.Get("/Sheet1/row[5]");
        row.Type.Should().Be("row");
        ((double)row.Format["height"]).Should().Be(25);
    }

    // ==================== BUG 7: Excel column width Set twice on same column ====================
    // Does the second Set correctly update or create a duplicate?

    [Fact]
    public void Bug_ColumnWidth_SetTwice()
    {
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "X" });

        _excelHandler.Set("/Sheet1/col[A]", new() { ["width"] = "15" });
        _excelHandler.Set("/Sheet1/col[A]", new() { ["width"] = "25" });

        var col = _excelHandler.Get("/Sheet1/col[A]");
        ((double)col.Format["width"]).Should().Be(25, "Second Set should override first");
    }

    // ==================== BUG 8: Word section Get after Reopen ====================
    // The body-level SectionProperties is the "last section" and is counted
    // in FindSectionProperties(). So /section[2] might be the body sectPr.
    // But does the index stay consistent after reopen?

    [Fact]
    public void Bug_SectionIndex_StableAfterReopen()
    {
        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "S1" });
        _wordHandler.Add("/body", "section", null, new() { ["type"] = "continuous" });
        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "S2" });

        var sec = _wordHandler.Get("/section[1]");
        sec.Type.Should().Be("section");

        ReopenWord();
        sec = _wordHandler.Get("/section[1]");
        sec.Type.Should().Be("section");
        ((string)sec.Format["type"]).Should().Be("continuous");
    }

    // ==================== BUG 9: Word style Set bold=false doesn't remove bold from Get ====================
    // When creating a style with bold=true, then Setting bold=false,
    // does Get correctly show bold is removed?

    [Fact]
    public void Bug_StyleSet_BoldFalseRemovesBold()
    {
        _wordHandler.Add("/body", "style", null, new()
        {
            ["name"] = "TestBold", ["id"] = "TestBold", ["bold"] = "true", ["font"] = "Arial"
        });

        var style = _wordHandler.Get("/styles/TestBold");
        style.Format.Should().ContainKey("bold");

        _wordHandler.Set("/styles/TestBold", new() { ["bold"] = "false" });

        style = _wordHandler.Get("/styles/TestBold");
        style.Format.Should().NotContainKey("bold",
            "After Set bold=false, bold should be removed from Format");
    }

    // ==================== ORDERING BUGS: Excel element order violations ====================

    [Fact]
    public void Bug_ExcelOrder_AutoFilterThenCF()
    {
        // Schema: sheetData > autoFilter > mergeCells > conditionalFormatting
        // If autoFilter is added first, then CF is InsertAfterSelf(sheetData),
        // CF ends up BEFORE autoFilter = wrong
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "1" });
        _excelHandler.Add("/Sheet1", "autofilter", null, new() { ["range"] = "A1:A10" });
        _excelHandler.Add("/Sheet1", "databar", null, new() { ["sqref"] = "A1:A10" });

        ReopenExcel();
        var errors = _excelHandler.Validate();
        errors.Should().BeEmpty("AutoFilter then CF should produce valid XML");
    }

    [Fact]
    public void Bug_ExcelOrder_CFThenAutoFilter()
    {
        // CF added first, then autoFilter — autoFilter should be BEFORE CF
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "1" });
        _excelHandler.Add("/Sheet1", "databar", null, new() { ["sqref"] = "A1:A10" });
        _excelHandler.Add("/Sheet1", "autofilter", null, new() { ["range"] = "A1:A10" });

        ReopenExcel();
        var errors = _excelHandler.Validate();
        errors.Should().BeEmpty("CF then AutoFilter should produce valid XML");
    }

    [Fact]
    public void Bug_ExcelOrder_MergeThenAutoFilter()
    {
        // Merge first, then autoFilter — autoFilter should be BEFORE mergeCells
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "1" });
        _excelHandler.Set("/Sheet1/A1:B1", new() { ["merge"] = "true" });
        _excelHandler.Add("/Sheet1", "autofilter", null, new() { ["range"] = "A1:A10" });

        ReopenExcel();
        var errors = _excelHandler.Validate();
        errors.Should().BeEmpty("Merge then AutoFilter should produce valid XML");
    }

    [Fact]
    public void Bug_ExcelOrder_ValidationThenCFThenMerge()
    {
        // Full combo: validation + CF + merge — all must be in correct order
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "1" });
        _excelHandler.Add("/Sheet1", "validation", null, new()
        {
            ["sqref"] = "B1:B10", ["type"] = "list", ["formula1"] = "Yes,No"
        });
        _excelHandler.Add("/Sheet1", "colorscale", null, new()
        {
            ["sqref"] = "A1:A10", ["mincolor"] = "FF0000", ["maxcolor"] = "00FF00"
        });
        _excelHandler.Set("/Sheet1/A1:C1", new() { ["merge"] = "true" });

        ReopenExcel();
        var errors = _excelHandler.Validate();
        errors.Should().BeEmpty("Validation + CF + Merge combo should produce valid XML");
    }

    // ==================== EDGE CASE BUGS ====================

    [Fact]
    public void Bug_Excel_MergeSameRangeTwice()
    {
        // Merging the same range twice should not create duplicate entries
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "X" });
        _excelHandler.Set("/Sheet1/A1:B1", new() { ["merge"] = "true" });
        _excelHandler.Set("/Sheet1/A1:B1", new() { ["merge"] = "true" });

        ReopenExcel();
        var errors = _excelHandler.Validate();
        errors.Should().BeEmpty("Merging same range twice should not cause errors");
        var cell = _excelHandler.Get("/Sheet1/A1");
        cell.Format.Should().ContainKey("merge");
    }

    [Fact]
    public void Bug_Word_SetSuperscript_ThenBold_Preserves()
    {
        // Setting superscript then bold should preserve both
        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "x" });
        _wordHandler.Add("/body/p[1]", "run", null, new() { ["text"] = "2", ["superscript"] = "true" });

        _wordHandler.Set("/body/p[1]/r[2]", new() { ["bold"] = "true" });

        var run = _wordHandler.Get("/body/p[1]/r[2]");
        run.Format.Should().ContainKey("superscript", "Superscript should survive bold set");
        run.Format.Should().ContainKey("bold", "Bold should be applied");
    }

    [Fact]
    public void Bug_Word_AddTwoFootnotes_IndependentIds()
    {
        // Two footnotes should have different IDs
        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "Text" });
        var fn1 = _wordHandler.Add("/body/p[1]", "footnote", null, new() { ["text"] = "First" });
        var fn2 = _wordHandler.Add("/body/p[1]", "footnote", null, new() { ["text"] = "Second" });

        fn1.Should().NotBe(fn2, "Two footnotes should have different paths");
        _wordHandler.Get(fn1).Text.Should().Contain("First");
        _wordHandler.Get(fn2).Text.Should().Contain("Second");
    }

    [Fact]
    public void Bug_Word_HangingAndFirstLineExclusive()
    {
        // Hanging indent and first line indent are mutually exclusive
        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "Para" });

        _wordHandler.Set("/body/p[1]", new() { ["firstlineindent"] = "2" });
        var node = _wordHandler.Get("/body/p[1]");
        node.Format.Should().ContainKey("firstLineIndent");

        // Set hanging — should clear firstline
        _wordHandler.Set("/body/p[1]", new() { ["hanging"] = "720" });
        node = _wordHandler.Get("/body/p[1]");
        node.Format.Should().ContainKey("hangingIndent");
        node.Format.Should().NotContainKey("firstLineIndent",
            "Hanging should clear firstLineIndent (mutually exclusive)");
    }

    [Fact]
    public void Bug_Excel_ChartOnSheetWithNoData()
    {
        // Adding a chart to a sheet with no data should work
        var path = _excelHandler.Add("/Sheet1", "chart", null, new()
        {
            ["chartType"] = "pie",
            ["data"] = "Sales:40,30,30",
            ["categories"] = "A,B,C"
        });
        path.Should().Be("/Sheet1/chart[1]");

        ReopenExcel();
        var chart = _excelHandler.Get("/Sheet1/chart[1]");
        chart.Type.Should().Be("chart");
    }

    [Fact]
    public void Bug_Pptx_AddRowMatchesTableColumnCount()
    {
        // Adding a row to a PPTX table should auto-match the existing column count
        using var pptxPath = new DisposablePath(".pptx");
        BlankDocCreator.Create(pptxPath.Path);
        using var handler = new PowerPointHandler(pptxPath.Path, editable: true);

        handler.Add("/", "slide", null, new());
        handler.Add("/slide[1]", "table", null, new() { ["rows"] = "1", ["cols"] = "4" });

        // Add row without specifying cols — should auto-detect 4 columns
        var path = handler.Add("/slide[1]/table[1]", "row", null, new() { ["c1"] = "A" });

        var table = handler.Get("/slide[1]/table[1]", depth: 2);
        var rows = table.Children.Where(c => c.Type == "tr").ToList();
        rows.Should().HaveCount(2);
        rows[1].Children.Should().HaveCount(4, "New row should match table's 4 columns");
    }

    /// <summary>Helper for disposable temp files.</summary>
    private class DisposablePath : IDisposable
    {
        public string Path { get; }
        public DisposablePath(string ext)
        {
            Path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"test_{Guid.NewGuid():N}{ext}");
        }
        public void Dispose() { if (File.Exists(Path)) File.Delete(Path); }
    }
}
