// Table Enhancement Tests — Full lifecycle tests for 7 new table features:
// 1. Cell gradient background (Word + PPTX)
// 2. Cell inline image (Word)
// 3. Precise row height with units (Word)
// 4. Multi-paragraph rich text in cells (Word + PPTX)
// 5. Built-in table styles (Word Add + PPTX Add)
// 6. Repeat header row (Word only)
// 7. Diagonal cell border (Word + PPTX)
//
// Every test follows: Create → Add → Get → Verify → Set(modify) → Get → Verify → Reopen → Verify

using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

public class TableEnhancementTests : IDisposable
{
    private readonly string _docxPath;
    private readonly string _pptxPath;
    private readonly string _xlsxPath;
    private WordHandler _wordHandler;
    private PowerPointHandler _pptxHandler;
    private ExcelHandler _excelHandler;

    public TableEnhancementTests()
    {
        _docxPath = Path.Combine(Path.GetTempPath(), $"tbl_enhance_{Guid.NewGuid():N}.docx");
        _pptxPath = Path.Combine(Path.GetTempPath(), $"tbl_enhance_{Guid.NewGuid():N}.pptx");
        _xlsxPath = Path.Combine(Path.GetTempPath(), $"tbl_enhance_{Guid.NewGuid():N}.xlsx");
        BlankDocCreator.Create(_docxPath);
        BlankDocCreator.Create(_pptxPath);
        BlankDocCreator.Create(_xlsxPath);
        _wordHandler = new WordHandler(_docxPath, editable: true);
        _pptxHandler = new PowerPointHandler(_pptxPath, editable: true);
        _excelHandler = new ExcelHandler(_xlsxPath, editable: true);
        _pptxHandler.Add("/", "slide", null, new());
    }

    public void Dispose()
    {
        _wordHandler.Dispose();
        _pptxHandler.Dispose();
        _excelHandler.Dispose();
        if (File.Exists(_docxPath)) File.Delete(_docxPath);
        if (File.Exists(_pptxPath)) File.Delete(_pptxPath);
        if (File.Exists(_xlsxPath)) File.Delete(_xlsxPath);
    }

    private void ReopenWord()
    {
        _wordHandler.Dispose();
        _wordHandler = new WordHandler(_docxPath, editable: true);
    }

    private void ReopenPptx()
    {
        _pptxHandler.Dispose();
        _pptxHandler = new PowerPointHandler(_pptxPath, editable: true);
    }

    private void ReopenExcel()
    {
        _excelHandler.Dispose();
        _excelHandler = new ExcelHandler(_xlsxPath, editable: true);
    }

    // ==================== Feature 1: Cell Gradient Background ====================

    [Fact]
    public void Word_CellGradient_FullLifecycle()
    {
        // 1. Create + Add
        _wordHandler.Add("/body", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });
        _wordHandler.Set("/body/tbl[1]/tr[1]/tc[1]", new() { ["text"] = "Gradient Cell" });

        // 2. Get + Verify (no gradient yet)
        var node = _wordHandler.Get("/body/tbl[1]/tr[1]/tc[1]");
        node.Text.Should().Be("Gradient Cell");
        node.Format.Should().NotContainKey("shd");

        // 3. Set gradient
        _wordHandler.Set("/body/tbl[1]/tr[1]/tc[1]", new() { ["shd"] = "gradient;FF0000;0000FF;90" });

        // 4. Get + Verify gradient applied
        node = _wordHandler.Get("/body/tbl[1]/tr[1]/tc[1]");
        node.Format.Should().ContainKey("shd");
        var shd = node.Format["shd"].ToString()!;
        shd.Should().Contain("gradient");
        shd.Should().Contain("FF0000");
        shd.Should().Contain("0000FF");
        shd.Should().Contain("90");

        // 5. Modify: change gradient to different colors
        _wordHandler.Set("/body/tbl[1]/tr[1]/tc[1]", new() { ["shd"] = "gradient;00FF00;FF00FF;180" });

        // 6. Get + Verify modification
        node = _wordHandler.Get("/body/tbl[1]/tr[1]/tc[1]");
        shd = node.Format["shd"].ToString()!;
        shd.Should().Contain("00FF00");
        shd.Should().Contain("FF00FF");
        shd.Should().Contain("180");

        // 7. Reopen + Verify persistence
        ReopenWord();
        node = _wordHandler.Get("/body/tbl[1]/tr[1]/tc[1]");
        node.Text.Should().Be("Gradient Cell");
        shd = node.Format["shd"].ToString()!;
        shd.Should().Contain("gradient");
        shd.Should().Contain("00FF00");
        shd.Should().Contain("FF00FF");

        // 8. Modify: override gradient with solid fill
        _wordHandler.Set("/body/tbl[1]/tr[1]/tc[1]", new() { ["shd"] = "solid;AABBCC" });
        node = _wordHandler.Get("/body/tbl[1]/tr[1]/tc[1]");
        node.Format["shd"].ToString().Should().Be("AABBCC");

        // 9. Verify solid persists
        ReopenWord();
        node = _wordHandler.Get("/body/tbl[1]/tr[1]/tc[1]");
        node.Format["shd"].ToString().Should().Be("AABBCC");
    }

    [Fact]
    public void Pptx_CellGradient_FullLifecycle()
    {
        // 1. Create + Add
        _pptxHandler.Add("/slide[1]", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });
        _pptxHandler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["text"] = "Grad" });

        // 2. Get + Verify (no fill yet)
        var node = _pptxHandler.Get("/slide[1]/table[1]/tr[1]/tc[1]");
        node.Text.Should().Be("Grad");

        // 3. Set gradient fill
        _pptxHandler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["fill"] = "FF0000-0000FF-90" });

        // 4. Get + Verify gradient (format: "COLOR1-COLOR2[-angle]")
        node = _pptxHandler.Get("/slide[1]/table[1]/tr[1]/tc[1]");
        node.Format.Should().ContainKey("fill");
        var fill = node.Format["fill"].ToString()!;
        fill.Should().Contain("FF0000");
        fill.Should().Contain("0000FF");

        // 5. Modify: change gradient colors
        _pptxHandler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["fill"] = "00FF00-FF00FF" });

        // 6. Get + Verify modification
        node = _pptxHandler.Get("/slide[1]/table[1]/tr[1]/tc[1]");
        fill = node.Format["fill"].ToString()!;
        fill.Should().Contain("00FF00");
        fill.Should().Contain("FF00FF");

        // 7. Reopen + Verify persistence
        ReopenPptx();
        node = _pptxHandler.Get("/slide[1]/table[1]/tr[1]/tc[1]");
        node.Format["fill"].ToString().Should().Contain("-"); // gradient format: COLOR1-COLOR2

        // 8. Modify: override with solid fill
        _pptxHandler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["fill"] = "AABBCC" });
        node = _pptxHandler.Get("/slide[1]/table[1]/tr[1]/tc[1]");
        node.Format["fill"].ToString().Should().Be("AABBCC");

        // 9. Verify solid persists
        ReopenPptx();
        node = _pptxHandler.Get("/slide[1]/table[1]/tr[1]/tc[1]");
        node.Format["fill"].ToString().Should().Be("AABBCC");
    }

    // ==================== Feature 2: Cell Inline Image ====================

    [Fact]
    public void Word_CellImage_FullLifecycle()
    {
        var imgPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.png");
        CreateTinyPng(imgPath);
        try
        {
            // 1. Create + Add table with text
            _wordHandler.Add("/body", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });
            _wordHandler.Set("/body/tbl[1]/tr[1]/tc[1]", new() { ["text"] = "Before Image" });

            // 2. Get + Verify text
            var node = _wordHandler.Get("/body/tbl[1]/tr[1]/tc[1]");
            node.Text.Should().Be("Before Image");

            // 3. Add picture to cell
            var result = _wordHandler.Add("/body/tbl[1]/tr[1]/tc[2]", "picture", null,
                new() { ["path"] = imgPath, ["width"] = "2cm", ["height"] = "2cm" });
            result.Should().Contain("p[");

            // 4. Verify cell still accessible
            node = _wordHandler.Get("/body/tbl[1]/tr[1]/tc[2]");
            node.Should().NotBeNull();

            // 5. Add picture to cell that already has text
            _wordHandler.Set("/body/tbl[1]/tr[1]/tc[1]", new() { ["text"] = "Has Text" });
            var result2 = _wordHandler.Add("/body/tbl[1]/tr[1]/tc[1]", "picture", null,
                new() { ["path"] = imgPath, ["width"] = "1cm", ["height"] = "1cm" });
            result2.Should().Contain("p[");

            // 6. Verify cell still has text
            node = _wordHandler.Get("/body/tbl[1]/tr[1]/tc[1]");
            node.Text.Should().Contain("Has Text");

            // 7. Reopen + Verify persistence
            ReopenWord();
            node = _wordHandler.Get("/body/tbl[1]/tr[1]/tc[1]");
            node.Should().NotBeNull();
            node = _wordHandler.Get("/body/tbl[1]/tr[1]/tc[2]");
            node.Should().NotBeNull();
        }
        finally
        {
            if (File.Exists(imgPath)) File.Delete(imgPath);
        }
    }

    // ==================== Feature 3: Precise Row Height ====================

    [Fact]
    public void Word_RowHeight_FullLifecycle_CmPtIn()
    {
        // 1. Create + Add table
        _wordHandler.Add("/body", "table", null, new() { ["rows"] = "3", ["cols"] = "2" });

        // 2. Get + Verify default (no explicit height)
        var node = _wordHandler.Get("/body/tbl[1]/tr[1]");
        node.Format.Should().NotContainKey("height");

        // 3. Set cm unit height
        _wordHandler.Set("/body/tbl[1]/tr[1]", new() { ["height"] = "1.5cm" });

        // 4. Get + Verify (1.5cm ≈ 851 twips)
        node = _wordHandler.Get("/body/tbl[1]/tr[1]");
        Convert.ToUInt32(node.Format["height"]).Should().BeInRange(849, 852);

        // 5. Modify: change to pt unit
        _wordHandler.Set("/body/tbl[1]/tr[1]", new() { ["height"] = "36pt" });

        // 6. Get + Verify (36pt = 720 twips)
        node = _wordHandler.Get("/body/tbl[1]/tr[1]");
        Convert.ToUInt32(node.Format["height"]).Should().Be(720);

        // 7. Set inch unit on row 2
        _wordHandler.Set("/body/tbl[1]/tr[2]", new() { ["height"] = "0.5in" });
        node = _wordHandler.Get("/body/tbl[1]/tr[2]");
        Convert.ToUInt32(node.Format["height"]).Should().Be(720);

        // 8. Set exact height on row 3
        _wordHandler.Set("/body/tbl[1]/tr[3]", new() { ["height.exact"] = "1cm" });
        node = _wordHandler.Get("/body/tbl[1]/tr[3]");
        Convert.ToUInt32(node.Format["height"]).Should().Be(567);
        node.Format["height.rule"].ToString().Should().Be("exact");

        // 9. Reopen + Verify all three rows persist
        ReopenWord();
        node = _wordHandler.Get("/body/tbl[1]/tr[1]");
        Convert.ToUInt32(node.Format["height"]).Should().Be(720);

        node = _wordHandler.Get("/body/tbl[1]/tr[2]");
        Convert.ToUInt32(node.Format["height"]).Should().Be(720);

        node = _wordHandler.Get("/body/tbl[1]/tr[3]");
        Convert.ToUInt32(node.Format["height"]).Should().Be(567);
        node.Format["height.rule"].ToString().Should().Be("exact");
    }

    [Fact]
    public void Word_RowHeight_AddWithUnits_FullLifecycle()
    {
        // 1. Create table
        _wordHandler.Add("/body", "table", null, new() { ["rows"] = "1", ["cols"] = "2" });

        // 2. Add row with cm height
        _wordHandler.Add("/body/tbl[1]", "row", null, new() { ["height"] = "2cm", ["c1"] = "A", ["c2"] = "B" });

        // 3. Get + Verify row height and cell text
        var node = _wordHandler.Get("/body/tbl[1]/tr[2]");
        Convert.ToUInt32(node.Format["height"]).Should().Be(1134);
        var cell = _wordHandler.Get("/body/tbl[1]/tr[2]/tc[1]");
        cell.Text.Should().Be("A");

        // 4. Add row with exact height
        _wordHandler.Add("/body/tbl[1]", "row", null, new() { ["height.exact"] = "36pt", ["c1"] = "X" });

        // 5. Get + Verify exact row
        node = _wordHandler.Get("/body/tbl[1]/tr[3]");
        Convert.ToUInt32(node.Format["height"]).Should().Be(720);
        node.Format["height.rule"].ToString().Should().Be("exact");

        // 6. Modify the row's height
        _wordHandler.Set("/body/tbl[1]/tr[2]", new() { ["height"] = "3cm" });
        node = _wordHandler.Get("/body/tbl[1]/tr[2]");
        Convert.ToUInt32(node.Format["height"]).Should().Be(1701);

        // 7. Reopen + Verify persistence
        ReopenWord();
        node = _wordHandler.Get("/body/tbl[1]/tr[2]");
        Convert.ToUInt32(node.Format["height"]).Should().Be(1701);
        cell = _wordHandler.Get("/body/tbl[1]/tr[2]/tc[1]");
        cell.Text.Should().Be("A");
    }

    // ==================== Feature 4: Multi-Paragraph in Cells ====================

    [Fact]
    public void Word_MultiParagraph_FullLifecycle()
    {
        // 1. Create + Add table with text in cell
        _wordHandler.Add("/body", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });
        _wordHandler.Set("/body/tbl[1]/tr[1]/tc[1]", new() { ["text"] = "Title" });

        // 2. Get + Verify single paragraph
        var node = _wordHandler.Get("/body/tbl[1]/tr[1]/tc[1]");
        node.Text.Should().Be("Title");

        // 3. Add second paragraph with different formatting
        _wordHandler.Add("/body/tbl[1]/tr[1]/tc[1]", "paragraph", null,
            new() { ["text"] = "Subtitle", ["bold"] = "true", ["color"] = "FF0000" });

        // 4. Get + Verify two paragraphs exist
        node = _wordHandler.Get("/body/tbl[1]/tr[1]/tc[1]");
        node.Children.Count.Should().BeGreaterThanOrEqualTo(2);

        // 5. Add a third paragraph
        _wordHandler.Add("/body/tbl[1]/tr[1]/tc[1]", "paragraph", null,
            new() { ["text"] = "Description", ["italic"] = "true", ["size"] = "10" });

        // 6. Get + Verify three paragraphs
        node = _wordHandler.Get("/body/tbl[1]/tr[1]/tc[1]");
        node.Children.Count.Should().BeGreaterThanOrEqualTo(3);

        // 7. Reopen + Verify persistence
        ReopenWord();
        node = _wordHandler.Get("/body/tbl[1]/tr[1]/tc[1]");
        node.Children.Count.Should().BeGreaterThanOrEqualTo(3);
        // Cell text should concatenate all paragraphs
        node.Text.Should().Contain("Title");
        node.Text.Should().Contain("Subtitle");
        node.Text.Should().Contain("Description");
    }

    [Fact]
    public void Pptx_MultiParagraph_FullLifecycle()
    {
        // 1. Create + Add table
        _pptxHandler.Add("/slide[1]", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });

        // 2. Get + Verify empty cell
        var node = _pptxHandler.Get("/slide[1]/table[1]/tr[1]/tc[1]");
        node.Text.Should().BeEmpty();

        // 3. Set multiline text (PPTX uses \n for multiple paragraphs)
        _pptxHandler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["text"] = "Line 1\\nLine 2\\nLine 3" });

        // 4. Get + Verify multiline
        node = _pptxHandler.Get("/slide[1]/table[1]/tr[1]/tc[1]");
        node.Text.Should().Contain("Line 1");
        node.Text.Should().Contain("Line 2");
        node.Text.Should().Contain("Line 3");

        // 5. Modify text to fewer lines
        _pptxHandler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["text"] = "Only One" });

        // 6. Get + Verify single line
        node = _pptxHandler.Get("/slide[1]/table[1]/tr[1]/tc[1]");
        node.Text.Should().Be("Only One");

        // 7. Reopen + Verify persistence
        ReopenPptx();
        node = _pptxHandler.Get("/slide[1]/table[1]/tr[1]/tc[1]");
        node.Text.Should().Be("Only One");
    }

    // ==================== Feature 5: Built-in Table Styles ====================

    [Fact]
    public void Word_TableStyle_FullLifecycle()
    {
        // 1. Create table with style
        _wordHandler.Add("/body", "table", null,
            new() { ["rows"] = "3", ["cols"] = "3", ["style"] = "GridTable4-Accent1" });

        // 2. Get + Verify table exists
        var node = _wordHandler.Get("/body/tbl[1]");
        node.Should().NotBeNull();
        node.Type.Should().Be("table");

        // 3. Add some cell content
        _wordHandler.Set("/body/tbl[1]/tr[1]/tc[1]", new() { ["text"] = "Header" });

        // 4. Get + Verify cell text
        var cell = _wordHandler.Get("/body/tbl[1]/tr[1]/tc[1]");
        cell.Text.Should().Be("Header");

        // 5. Modify: change style
        _wordHandler.Set("/body/tbl[1]", new() { ["style"] = "LightShading-Accent2" });

        // 6. Verify table still accessible after style change
        node = _wordHandler.Get("/body/tbl[1]");
        node.Should().NotBeNull();
        cell = _wordHandler.Get("/body/tbl[1]/tr[1]/tc[1]");
        cell.Text.Should().Be("Header");

        // 7. Reopen + Verify persistence
        ReopenWord();
        node = _wordHandler.Get("/body/tbl[1]");
        node.Should().NotBeNull();
        cell = _wordHandler.Get("/body/tbl[1]/tr[1]/tc[1]");
        cell.Text.Should().Be("Header");
    }

    [Fact]
    public void Pptx_TableStyle_FullLifecycle()
    {
        // 1. Create table with named style
        _pptxHandler.Add("/slide[1]", "table", null,
            new() { ["rows"] = "3", ["cols"] = "3", ["style"] = "medium2" });

        // 2. Get + Verify
        var node = _pptxHandler.Get("/slide[1]/table[1]");
        node.Should().NotBeNull();
        node.Type.Should().Be("table");

        // 3. Add cell content
        _pptxHandler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["text"] = "Styled" });

        // 4. Get + Verify cell
        var cell = _pptxHandler.Get("/slide[1]/table[1]/tr[1]/tc[1]");
        cell.Text.Should().Be("Styled");

        // 5. Modify: change style
        _pptxHandler.Set("/slide[1]/table[1]", new() { ["style"] = "dark1" });

        // 6. Verify table & cell still valid
        cell = _pptxHandler.Get("/slide[1]/table[1]/tr[1]/tc[1]");
        cell.Text.Should().Be("Styled");

        // 7. Reopen + Verify persistence
        ReopenPptx();
        node = _pptxHandler.Get("/slide[1]/table[1]");
        node.Should().NotBeNull();
        cell = _pptxHandler.Get("/slide[1]/table[1]/tr[1]/tc[1]");
        cell.Text.Should().Be("Styled");
    }

    // ==================== Feature 6: Repeat Header Row ====================

    [Fact]
    public void Word_RepeatHeader_FullLifecycle()
    {
        // 1. Create + Add table
        _wordHandler.Add("/body", "table", null, new() { ["rows"] = "3", ["cols"] = "2" });
        _wordHandler.Set("/body/tbl[1]/tr[1]/tc[1]", new() { ["text"] = "Name" });
        _wordHandler.Set("/body/tbl[1]/tr[1]/tc[2]", new() { ["text"] = "Value" });

        // 2. Get + Verify no header yet
        var node = _wordHandler.Get("/body/tbl[1]/tr[1]");
        node.Format.Should().NotContainKey("header");

        // 3. Set header = true
        _wordHandler.Set("/body/tbl[1]/tr[1]", new() { ["header"] = "true" });

        // 4. Get + Verify header is set
        node = _wordHandler.Get("/body/tbl[1]/tr[1]");
        node.Format.Should().ContainKey("header");
        node.Format["header"].Should().Be(true);

        // 5. Verify cell text preserved
        var cell = _wordHandler.Get("/body/tbl[1]/tr[1]/tc[1]");
        cell.Text.Should().Be("Name");

        // 6. Reopen + Verify persistence
        ReopenWord();
        node = _wordHandler.Get("/body/tbl[1]/tr[1]");
        node.Format.Should().ContainKey("header");
        node.Format["header"].Should().Be(true);
        cell = _wordHandler.Get("/body/tbl[1]/tr[1]/tc[1]");
        cell.Text.Should().Be("Name");

        // 7. Modify: remove header
        _wordHandler.Set("/body/tbl[1]/tr[1]", new() { ["header"] = "false" });

        // 8. Get + Verify header removed
        node = _wordHandler.Get("/body/tbl[1]/tr[1]");
        node.Format.Should().NotContainKey("header");

        // 9. Reopen + Verify removal persists
        ReopenWord();
        node = _wordHandler.Get("/body/tbl[1]/tr[1]");
        node.Format.Should().NotContainKey("header");
    }

    [Fact]
    public void Word_RepeatHeader_AddRow_FullLifecycle()
    {
        // 1. Create table
        _wordHandler.Add("/body", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });

        // 2. Add header row at index 0 (becomes first row)
        _wordHandler.Add("/body/tbl[1]", "row", 0,
            new() { ["header"] = "true", ["c1"] = "Col A", ["c2"] = "Col B" });

        // 3. Get + Verify header row
        var node = _wordHandler.Get("/body/tbl[1]/tr[1]");
        node.Format.Should().ContainKey("header");
        node.Format["header"].Should().Be(true);
        var cell = _wordHandler.Get("/body/tbl[1]/tr[1]/tc[1]");
        cell.Text.Should().Be("Col A");

        // 4. Verify table now has 3 rows
        var tbl = _wordHandler.Get("/body/tbl[1]");
        tbl.Children.Count.Should().Be(3);

        // 5. Reopen + Verify
        ReopenWord();
        node = _wordHandler.Get("/body/tbl[1]/tr[1]");
        node.Format.Should().ContainKey("header");
        node.Format["header"].Should().Be(true);
        cell = _wordHandler.Get("/body/tbl[1]/tr[1]/tc[1]");
        cell.Text.Should().Be("Col A");
    }

    // ==================== Feature 7: Diagonal Cell Border ====================

    [Fact]
    public void Word_DiagonalBorder_FullLifecycle()
    {
        // 1. Create + Add table with text
        _wordHandler.Add("/body", "table", null, new() { ["rows"] = "3", ["cols"] = "3" });
        _wordHandler.Set("/body/tbl[1]/tr[1]/tc[1]", new() { ["text"] = "Diagonal" });

        // 2. Get + Verify no diagonal border
        var node = _wordHandler.Get("/body/tbl[1]/tr[1]/tc[1]");
        node.Text.Should().Be("Diagonal");
        node.Format.Should().NotContainKey("border.tl2br");
        node.Format.Should().NotContainKey("border.tr2bl");

        // 3. Set tl2br diagonal
        _wordHandler.Set("/body/tbl[1]/tr[1]/tc[1]", new() { ["border.tl2br"] = "single;4;000000" });

        // 4. Get + Verify tl2br set
        node = _wordHandler.Get("/body/tbl[1]/tr[1]/tc[1]");
        node.Format.Should().ContainKey("border.tl2br");
        node.Text.Should().Be("Diagonal");

        // 5. Set tr2bl on another cell
        _wordHandler.Set("/body/tbl[1]/tr[1]/tc[2]", new() { ["border.tr2bl"] = "single;4;FF0000" });

        // 6. Get + Verify tr2bl
        node = _wordHandler.Get("/body/tbl[1]/tr[1]/tc[2]");
        node.Format.Should().ContainKey("border.tr2bl");

        // 7. Modify: change tl2br style
        _wordHandler.Set("/body/tbl[1]/tr[1]/tc[1]", new() { ["border.tl2br"] = "double;8;0000FF" });

        // 8. Get + Verify modification
        node = _wordHandler.Get("/body/tbl[1]/tr[1]/tc[1]");
        node.Format.Should().ContainKey("border.tl2br");

        // 9. Reopen + Verify persistence of both
        ReopenWord();
        node = _wordHandler.Get("/body/tbl[1]/tr[1]/tc[1]");
        node.Format.Should().ContainKey("border.tl2br");
        node.Text.Should().Be("Diagonal");

        node = _wordHandler.Get("/body/tbl[1]/tr[1]/tc[2]");
        node.Format.Should().ContainKey("border.tr2bl");
    }

    [Fact]
    public void Pptx_DiagonalBorder_FullLifecycle()
    {
        // 1. Create + Add table with text
        _pptxHandler.Add("/slide[1]", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });
        _pptxHandler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["text"] = "Diag" });

        // 2. Get + Verify no diagonal
        var node = _pptxHandler.Get("/slide[1]/table[1]/tr[1]/tc[1]");
        node.Text.Should().Be("Diag");
        node.Format.Should().NotContainKey("border.tl2br");

        // 3. Set tl2br diagonal
        _pptxHandler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["border.tl2br"] = "1pt solid FF0000" });

        // 4. Get + Verify
        node = _pptxHandler.Get("/slide[1]/table[1]/tr[1]/tc[1]");
        node.Format.Should().ContainKey("border.tl2br");
        node.Text.Should().Be("Diag");

        // 5. Modify: change border
        _pptxHandler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["border.tl2br"] = "2pt dash 0000FF" });

        // 6. Get + Verify modification
        node = _pptxHandler.Get("/slide[1]/table[1]/tr[1]/tc[1]");
        node.Format.Should().ContainKey("border.tl2br");

        // 7. Reopen + Verify persistence
        ReopenPptx();
        node = _pptxHandler.Get("/slide[1]/table[1]/tr[1]/tc[1]");
        node.Format.Should().ContainKey("border.tl2br");
        node.Text.Should().Be("Diag");
    }

    // ==================== Feature 8: Excel Gradient Fill ====================

    [Fact]
    public void Excel_CellGradient_FullLifecycle()
    {
        // 1. Add cell with text
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Gradient" });

        // 2. Get + Verify no fill
        var node = _excelHandler.Get("/Sheet1/A1");
        node.Text.Should().Be("Gradient");
        node.Format.Should().NotContainKey("fill");

        // 3. Set gradient fill
        _excelHandler.Set("/Sheet1/A1", new() { ["fill"] = "FF0000-0000FF-90" });

        // 4. Get + Verify gradient
        node = _excelHandler.Get("/Sheet1/A1");
        node.Format.Should().ContainKey("fill");
        var fill = node.Format["fill"].ToString()!;
        fill.Should().Contain("gradient");
        fill.Should().Contain("FF0000");
        fill.Should().Contain("0000FF");

        // 5. Modify: change gradient
        _excelHandler.Set("/Sheet1/A1", new() { ["fill"] = "00FF00-FF00FF-180" });

        // 6. Get + Verify modified
        node = _excelHandler.Get("/Sheet1/A1");
        fill = node.Format["fill"].ToString()!;
        fill.Should().Contain("00FF00");
        fill.Should().Contain("FF00FF");

        // 7. Reopen + Verify persistence
        ReopenExcel();
        node = _excelHandler.Get("/Sheet1/A1");
        node.Text.Should().Be("Gradient");
        fill = node.Format["fill"].ToString()!;
        fill.Should().Contain("gradient");
        fill.Should().Contain("00FF00");

        // 8. Override with solid fill
        _excelHandler.Set("/Sheet1/A1", new() { ["fill"] = "AABBCC" });
        node = _excelHandler.Get("/Sheet1/A1");
        node.Format["fill"].ToString().Should().Be("AABBCC");

        // 9. Verify solid persists
        ReopenExcel();
        node = _excelHandler.Get("/Sheet1/A1");
        node.Format["fill"].ToString().Should().Be("AABBCC");
    }

    // ==================== Feature 9: Excel Diagonal Border ====================

    [Fact]
    public void Excel_DiagonalBorder_FullLifecycle()
    {
        // 1. Add cell with text
        _excelHandler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Diagonal" });

        // 2. Get + Verify no diagonal border
        var node = _excelHandler.Get("/Sheet1/A1");
        node.Text.Should().Be("Diagonal");
        node.Format.Should().NotContainKey("border.diagonal");

        // 3. Set diagonal border (tl2br = diagonalDown)
        _excelHandler.Set("/Sheet1/A1", new()
        {
            ["border.diagonal"] = "thin",
            ["border.diagonal.color"] = "FF0000",
            ["border.diagonalDown"] = "true"
        });

        // 4. Get + Verify diagonal applied
        node = _excelHandler.Get("/Sheet1/A1");
        node.Format.Should().ContainKey("border.diagonal");
        node.Format["border.diagonal"].ToString().Should().Be("thin");
        node.Format.Should().ContainKey("border.diagonalDown");
        node.Format["border.diagonalDown"].Should().Be(true);

        // 5. Modify: add diagonalUp too (cross pattern)
        _excelHandler.Set("/Sheet1/A1", new()
        {
            ["border.diagonal"] = "medium",
            ["border.diagonal.color"] = "0000FF",
            ["border.diagonalUp"] = "true",
            ["border.diagonalDown"] = "true"
        });

        // 6. Get + Verify both diagonals
        node = _excelHandler.Get("/Sheet1/A1");
        node.Format["border.diagonal"].ToString().Should().Be("medium");
        node.Format.Should().ContainKey("border.diagonalUp");
        node.Format.Should().ContainKey("border.diagonalDown");

        // 7. Reopen + Verify persistence
        ReopenExcel();
        node = _excelHandler.Get("/Sheet1/A1");
        node.Text.Should().Be("Diagonal");
        node.Format.Should().ContainKey("border.diagonal");
        node.Format.Should().ContainKey("border.diagonalUp");
        node.Format.Should().ContainKey("border.diagonalDown");
    }

    // ==================== Feature 10: PPTX Cell Image ====================

    [Fact]
    public void Pptx_CellImage_FullLifecycle()
    {
        var imgPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.png");
        CreateTinyPng(imgPath);
        try
        {
            // 1. Create table with text
            _pptxHandler.Add("/slide[1]", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });
            _pptxHandler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["text"] = "ImgCell" });

            // 2. Get + Verify no image
            var node = _pptxHandler.Get("/slide[1]/table[1]/tr[1]/tc[1]");
            node.Text.Should().Be("ImgCell");
            node.Format.Should().NotContainKey("image.relId");

            // 3. Set image fill
            _pptxHandler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["image"] = imgPath });

            // 4. Get + Verify image fill applied
            node = _pptxHandler.Get("/slide[1]/table[1]/tr[1]/tc[1]");
            node.Format["fill"].ToString().Should().Be("image");
            node.Format.Should().ContainKey("image.relId");

            // 5. Text should still be accessible
            node.Text.Should().Be("ImgCell");

            // 6. Modify: override with solid fill
            _pptxHandler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["fill"] = "AABBCC" });
            node = _pptxHandler.Get("/slide[1]/table[1]/tr[1]/tc[1]");
            node.Format["fill"].ToString().Should().Be("AABBCC");

            // 7. Set image again
            _pptxHandler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["image"] = imgPath });

            // 8. Reopen + Verify persistence
            ReopenPptx();
            node = _pptxHandler.Get("/slide[1]/table[1]/tr[1]/tc[1]");
            node.Format["fill"].ToString().Should().Be("image");
            node.Format.Should().ContainKey("image.relId");
            node.Text.Should().Be("ImgCell");
        }
        finally
        {
            if (File.Exists(imgPath)) File.Delete(imgPath);
        }
    }

    // ==================== Helper ====================

    private static void CreateTinyPng(string path)
    {
        // Minimal valid 1x1 white PNG
        byte[] png = [
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
            0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
            0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
            0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
            0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41,
            0x54, 0x08, 0xD7, 0x63, 0xF8, 0xCF, 0xC0, 0x00,
            0x00, 0x00, 0x02, 0x00, 0x01, 0xE2, 0x21, 0xBC,
            0x33, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E,
            0x44, 0xAE, 0x42, 0x60, 0x82
        ];
        File.WriteAllBytes(path, png);
    }
}
