// Bug hunt Part 31 — Word handler and Excel handler confirmed bugs:
// 1. Word Find/Replace doesn't work on text inside hyperlinks
// 2. Word Set paragraph "text" not supported (no handler)
// 3. Excel picture shadow removes ALL effects (EffectList clobbered)
// 4. Word paragraph bold=false doesn't remove bold (IsTruthy returns false)
// 5. PPTX table cell "strike" with "single" value creates invalid XML
// 6. Word Run Set underline with "single" creates UnderlineValues("single")
// 7. Excel Set cell border color readback doesn't strip ARGB prefix
// 8. PPTX shape text color with scheme color readback
// 9. Word section set columns with "2" doesn't persist after reopen

using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

public class BugHuntPart31 : IDisposable
{
    private readonly string _docxPath;
    private readonly string _xlsxPath;
    private readonly string _pptxPath;
    private WordHandler _wordHandler;
    private ExcelHandler _excelHandler;
    private PowerPointHandler _pptxHandler;

    public BugHuntPart31()
    {
        _docxPath = Path.Combine(Path.GetTempPath(), $"bughunt31_{Guid.NewGuid():N}.docx");
        _xlsxPath = Path.Combine(Path.GetTempPath(), $"bughunt31_{Guid.NewGuid():N}.xlsx");
        _pptxPath = Path.Combine(Path.GetTempPath(), $"bughunt31_{Guid.NewGuid():N}.pptx");
        BlankDocCreator.Create(_docxPath);
        BlankDocCreator.Create(_xlsxPath);
        BlankDocCreator.Create(_pptxPath);
        _wordHandler = new WordHandler(_docxPath, editable: true);
        _excelHandler = new ExcelHandler(_xlsxPath, editable: true);
        _pptxHandler = new PowerPointHandler(_pptxPath, editable: true);
    }

    public void Dispose()
    {
        _wordHandler.Dispose();
        _excelHandler.Dispose();
        _pptxHandler.Dispose();
        if (File.Exists(_docxPath)) File.Delete(_docxPath);
        if (File.Exists(_xlsxPath)) File.Delete(_xlsxPath);
        if (File.Exists(_pptxPath)) File.Delete(_pptxPath);
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

    private PowerPointHandler ReopenPptx()
    {
        _pptxHandler.Dispose();
        _pptxHandler = new PowerPointHandler(_pptxPath, editable: true);
        return _pptxHandler;
    }

    // CONFIRMED BUG: Word Find/Replace doesn't work on text inside
    [Fact]
    public void Bug_Word_FindReplace_DoesntWorkInHyperlinks()
    {
        // 1. Add elements
        _wordHandler.Add("/", "paragraph", null, new() { ["text"] = "Click here to visit" });
        _wordHandler.Add("/body/p[1]", "run", null, new()
        {
            ["text"] = "our website",
            ["link"] = "https://example.com"
        });

        // 2. Get + Verify initial state
        var node1 = _wordHandler.Get("/body/p[1]");
        var fullText = node1.Text ?? "";
        fullText.Should().Contain("our website");

        // 3. Set (modify via find/replace)
        _wordHandler.Set("/", new()
        {
            ["find"] = "our website",
            ["replace"] = "the site"
        });

        // 4. Get + Verify modification
        var node2 = _wordHandler.Get("/body/p[1]");
        var newText = node2.Text ?? "";

        // BUG: The text inside the hyperlink is NOT replaced because
        // ReplaceInParagraph uses para.Elements<Run>() instead of Descendants<Run>()
        newText.Should().Contain("the site",
            "Find/Replace should work on text inside hyperlinks, " +
            "but ReplaceInParagraph uses para.Elements<Run>() instead of Descendants<Run>()");

        // 5. Reopen + Verify persistence
        ReopenWord();
        var node3 = _wordHandler.Get("/body/p[1]");
        (node3.Text ?? "").Should().Contain("the site",
            "Find/Replace result should persist after reopen");
    }

    [Fact]
    public void Bug_Word_FindReplace_WorksOnRegularText_ForComparison()
    {
        // 1. Add element
        _wordHandler.Add("/", "paragraph", null, new() { ["text"] = "Hello world, this is a test" });

        // 2. Get + Verify initial state
        var node1 = _wordHandler.Get("/body/p[1]");
        node1.Text.Should().Contain("world");

        // 3. Set (modify via find/replace)
        _wordHandler.Set("/", new() { ["find"] = "world", ["replace"] = "earth" });

        // 4. Get + Verify modification
        var node2 = _wordHandler.Get("/body/p[1]");
        node2.Text.Should().Contain("earth");
        node2.Text.Should().NotContain("world");

        // 5. Reopen + Verify persistence
        ReopenWord();
        var node3 = _wordHandler.Get("/body/p[1]");
        node3.Text.Should().Contain("earth");
    }

    // CONFIRMED BUG: Word paragraph-level bold=false doesn't actually
    [Fact]
    public void Bug_Word_Paragraph_BoldFalse_DoesntClearRunBold()
    {
        // 1. Add a bold paragraph
        _wordHandler.Add("/", "paragraph", null, new()
        {
            ["text"] = "Bold text here",
            ["bold"] = "true"
        });

        // 2. Get + Verify initial state (bold)
        var raw1 = _wordHandler.Raw("/document");
        raw1.Should().Contain("<w:b", "text should be bold initially");

        // 3. Set bold=false on the paragraph
        _wordHandler.Set("/body/p[1]", new() { ["bold"] = "false" });

        // 4. Get + Verify modification
        var raw2 = _wordHandler.Raw("/document");
        // BUG: The run's bold element from the Add is still there
        raw2.Should().NotContain("<w:b/>",
            "setting bold=false on paragraph should remove bold from all runs, " +
            "but ApplyRunFormatting only adds bold when truthy, doesn't remove existing");

        // 5. Reopen + Verify persistence
        ReopenWord();
        var raw3 = _wordHandler.Raw("/document");
        raw3.Should().NotContain("<w:b/>",
            "bold=false should persist after reopen");
    }

    // CONFIRMED BUG: Word Run underline with value "true" creates
    [Fact]
    public void Bug_Word_Run_Underline_True_MayCreateInvalidValue()
    {
        // 1. Add element
        _wordHandler.Add("/", "paragraph", null, new() { ["text"] = "Underlined text" });

        // 2. Get + Verify initial state
        var node1 = _wordHandler.Get("/body/p[1]", 1);
        node1.Children.Should().HaveCountGreaterThan(0);

        // 3. Set underline "true" on the run
        _wordHandler.Set("/body/p[1]/r[1]", new() { ["underline"] = "true" });

        // 4. Get + Verify modification
        var raw = _wordHandler.Raw("/document");
        // BUG: "true" is passed as-is to UnderlineValues constructor.
        // Should map "true" -> "single" like PPTX does.
        raw.Should().Contain("w:val=\"single\"",
            "underline='true' should produce w:val='single', " +
            "but Word handler passes 'true' directly to UnderlineValues constructor");

        // 5. Reopen + Verify persistence
        ReopenWord();
        var raw2 = _wordHandler.Raw("/document");
        raw2.Should().Contain("w:val=\"single\"",
            "underline value should persist after reopen");
    }

    [Fact]
    public void Bug_Word_Run_Underline_Single_Works_ForComparison()
    {
        // 1. Add element
        _wordHandler.Add("/", "paragraph", null, new() { ["text"] = "Underlined text" });

        // 2. Get + Verify initial state
        var node1 = _wordHandler.Get("/body/p[1]", 1);
        node1.Children.Should().HaveCountGreaterThan(0);

        // 3. Set underline "single"
        _wordHandler.Set("/body/p[1]/r[1]", new() { ["underline"] = "single" });

        // 4. Get + Verify modification
        var raw = _wordHandler.Raw("/document");
        raw.Should().Contain("w:val=\"single\"");

        // 5. Reopen + Verify persistence
        ReopenWord();
        var raw2 = _wordHandler.Raw("/document");
        raw2.Should().Contain("w:val=\"single\"");
    }

    // EDGE CASE: Word section properties persistence after reopen.
    [Fact]
    public void Edge_Word_Section_Columns_Persist()
    {
        // 1. Add element (section exists by default, add paragraph for content)
        _wordHandler.Add("/", "paragraph", null, new() { ["text"] = "Column content" });

        // 2. Get + Verify initial state
        var node1 = _wordHandler.Get("/section[1]");
        node1.Should().NotBeNull();

        // 3. Set columns
        _wordHandler.Set("/section[1]", new() { ["columns"] = "2" });

        // 4. Get + Verify modification
        var node2 = _wordHandler.Get("/section[1]");
        node2.Should().NotBeNull();
        node2.Format.Should().ContainKey("columns");
        node2.Format["columns"].ToString().Should().Be("2");

        // 5. Reopen + Verify persistence
        ReopenWord();
        var node3 = _wordHandler.Get("/section[1]");
        node3.Should().NotBeNull();
        node3.Format.Should().ContainKey("columns");
        node3.Format["columns"].ToString().Should().Be("2");
    }

    // EDGE CASE: Word paragraph styles round-trip.
    [Fact]
    public void Edge_Word_Paragraph_Style_RoundTrip()
    {
        // 1. Add element with style
        _wordHandler.Add("/", "paragraph", null, new()
        {
            ["text"] = "Heading text",
            ["style"] = "Heading1"
        });

        // 2. Get + Verify initial state
        var node1 = _wordHandler.Get("/body/p[1]");
        node1.Format.Should().ContainKey("style");
        node1.Format["style"].ToString()!.Should().Contain("Heading");

        // 3. Set (modify style)
        _wordHandler.Set("/body/p[1]", new() { ["style"] = "Heading2" });

        // 4. Get + Verify modification
        var node2 = _wordHandler.Get("/body/p[1]");
        node2.Format.Should().ContainKey("style");
        node2.Format["style"].ToString()!.Should().Contain("Heading");

        // 5. Reopen + Verify persistence
        ReopenWord();
        var node3 = _wordHandler.Get("/body/p[1]");
        node3.Format.Should().ContainKey("style");
    }

    // EDGE CASE: Word paragraph indentation round-trip.
    [Fact]
    public void Edge_Word_Paragraph_Indent_RoundTrip()
    {
        // 1. Add element with indentation
        _wordHandler.Add("/", "paragraph", null, new()
        {
            ["text"] = "Indented text",
            ["leftindent"] = "720"
        });

        // 2. Get + Verify initial state
        var node1 = _wordHandler.Get("/body/p[1]");
        node1.Format.Should().ContainKey("leftindent");
        node1.Format["leftindent"].ToString().Should().Be("720");

        // 3. Set (modify indentation)
        _wordHandler.Set("/body/p[1]", new() { ["leftindent"] = "1440" });

        // 4. Get + Verify modification
        var node2 = _wordHandler.Get("/body/p[1]");
        node2.Format.Should().ContainKey("leftindent");
        node2.Format["leftindent"].ToString().Should().Be("1440");

        // 5. Reopen + Verify persistence
        ReopenWord();
        var node3 = _wordHandler.Get("/body/p[1]");
        node3.Format.Should().ContainKey("leftindent");
        node3.Format["leftindent"].ToString().Should().Be("1440");
    }

    // EDGE CASE: Word paragraph spacing round-trip.
    [Fact]
    public void Edge_Word_Paragraph_Spacing_RoundTrip()
    {
        // 1. Add element with spacing
        _wordHandler.Add("/", "paragraph", null, new()
        {
            ["text"] = "Spaced text",
            ["spacebefore"] = "240",
            ["spaceafter"] = "120"
        });

        // 2. Get + Verify initial state
        var node1 = _wordHandler.Get("/body/p[1]");
        node1.Format.Should().ContainKey("spacebefore");
        node1.Format.Should().ContainKey("spaceafter");
        node1.Format["spacebefore"].ToString().Should().Be("240");
        node1.Format["spaceafter"].ToString().Should().Be("120");

        // 3. Set (modify spacing)
        _wordHandler.Set("/body/p[1]", new() { ["spacebefore"] = "480" });

        // 4. Get + Verify modification
        var node2 = _wordHandler.Get("/body/p[1]");
        node2.Format.Should().ContainKey("spacebefore");
        node2.Format["spacebefore"].ToString().Should().Be("480");

        // 5. Reopen + Verify persistence
        ReopenWord();
        var node3 = _wordHandler.Get("/body/p[1]");
        node3.Format.Should().ContainKey("spacebefore");
        node3.Format["spacebefore"].ToString().Should().Be("480");
    }

    // EDGE CASE: Word multiple runs in one paragraph.
    [Fact]
    public void Edge_Word_MultipleRuns_DifferentFormatting()
    {
        // 1. Add elements
        _wordHandler.Add("/", "paragraph", null, new() { ["text"] = "Normal text" });
        _wordHandler.Add("/body/p[1]", "run", null, new()
        {
            ["text"] = " bold text",
            ["bold"] = "true"
        });
        _wordHandler.Add("/body/p[1]", "run", null, new()
        {
            ["text"] = " italic text",
            ["italic"] = "true"
        });

        // 2. Get + Verify initial state
        var node1 = _wordHandler.Get("/body/p[1]", 1);
        node1.Text.Should().Contain("Normal text");
        node1.Text.Should().Contain("bold text");
        node1.Text.Should().Contain("italic text");
        node1.Children.Should().HaveCountGreaterThanOrEqualTo(3);

        // 3. Set (modify a run)
        _wordHandler.Set("/body/p[1]/r[2]", new() { ["underline"] = "single" });

        // 4. Get + Verify modification
        var node2 = _wordHandler.Get("/body/p[1]", 1);
        node2.Text.Should().Contain("bold text");
        node2.Children.Should().HaveCountGreaterThanOrEqualTo(3);

        // 5. Reopen + Verify persistence
        ReopenWord();
        var node3 = _wordHandler.Get("/body/p[1]", 1);
        node3.Text.Should().Contain("Normal text");
        node3.Text.Should().Contain("bold text");
        node3.Text.Should().Contain("italic text");
    }

    // CONFIRMED BUG: Excel picture shadow clobbers other effects.
    [Fact]
    public void Bug_Excel_Picture_Shadow_ClobbersGlow()
    {
        var imgPath = Path.Combine(Path.GetTempPath(), $"test_img_{Guid.NewGuid():N}.png");
        CreateMinimalPng(imgPath);

        try
        {
            // 1. Add element
            _excelHandler.Add("/Sheet1", "picture", null, new()
            {
                ["path"] = imgPath,
                ["x"] = "0",
                ["y"] = "0"
            });

            // 2. Get + Verify initial state
            var node1 = _excelHandler.Get("/Sheet1/picture[1]");
            node1.Should().NotBeNull();

            // 3. Set glow first
            _excelHandler.Set("/Sheet1/picture[1]", new() { ["glow"] = "FF0000:8" });

            // Verify glow was added
            var raw1 = _excelHandler.Raw("/Sheet1/drawing");
            raw1.Should().Contain("glow", "glow should be set initially");

            // 4. Set shadow (this should NOT remove glow)
            _excelHandler.Set("/Sheet1/picture[1]", new() { ["shadow"] = "000000:4:3:45" });

            // 5. Get + Verify - BUG: shadow handler removes entire EffectList
            var raw2 = _excelHandler.Raw("/Sheet1/drawing");
            raw2.Should().Contain("glow",
                "adding shadow to Excel picture should not remove existing glow, " +
                "but ExcelHandler.Set removes entire EffectList instead of just OuterShadow");
        }
        finally
        {
            if (File.Exists(imgPath)) File.Delete(imgPath);
        }
    }

    // EDGE CASE: Excel cell formatting persistence after reopen.
    [Fact]
    public void Edge_Excel_CellFormatting_Persists()
    {
        // 1. Add element (set cell value and formatting)
        _excelHandler.Set("/Sheet1/A1", new()
        {
            ["value"] = "Styled cell",
            ["bold"] = "true",
            ["fill"] = "FFFF00"
        });

        // 2. Get + Verify initial state
        var node1 = _excelHandler.Get("/Sheet1/A1");
        node1.Text.Should().Be("Styled cell");
        node1.Format.Should().ContainKey("font.bold");
        node1.Format["font.bold"].Should().Be(true);
        node1.Format.Should().ContainKey("fill");
        node1.Format["fill"].ToString().Should().Be("FFFF00");

        // 3. Set (modify fill color)
        _excelHandler.Set("/Sheet1/A1", new() { ["fill"] = "00FF00" });

        // 4. Get + Verify modification
        var node2 = _excelHandler.Get("/Sheet1/A1");
        node2.Format.Should().ContainKey("fill");
        node2.Format["fill"].ToString().Should().Be("00FF00");

        // 5. Reopen + Verify persistence
        ReopenExcel();
        var node3 = _excelHandler.Get("/Sheet1/A1");
        node3.Text.Should().Be("Styled cell");
        node3.Format.Should().ContainKey("font.bold");
        node3.Format["font.bold"].Should().Be(true);
        node3.Format.Should().ContainKey("fill");
        node3.Format["fill"].ToString().Should().Be("00FF00");
    }

    // EDGE CASE: Excel cell merge and readback.
    [Fact]
    public void Edge_Excel_MergeCell_RoundTrip()
    {
        // 1. Add element (set cell value)
        _excelHandler.Set("/Sheet1/A1", new() { ["value"] = "Merged" });

        // 2. Get + Verify initial state
        var node1 = _excelHandler.Get("/Sheet1/A1");
        node1.Text.Should().Be("Merged");

        // 3. Set (merge at sheet level)
        _excelHandler.Set("/Sheet1", new() { ["merge"] = "A1:C1" });

        // 4. Get + Verify modification
        var node2 = _excelHandler.Get("/Sheet1/A1");
        node2.Format.Should().ContainKey("merge");
        node2.Format["merge"].ToString().Should().Be("A1:C1");

        // 5. Reopen + Verify persistence
        ReopenExcel();
        var node3 = _excelHandler.Get("/Sheet1/A1");
        node3.Format.Should().ContainKey("merge");
        node3.Format["merge"].ToString().Should().Be("A1:C1");
    }

    // EDGE CASE: Excel cell formula round-trip.
    [Fact]
    public void Edge_Excel_CellFormula_RoundTrip()
    {
        // 1. Add elements (set cell values)
        _excelHandler.Set("/Sheet1/A1", new() { ["value"] = "10" });
        _excelHandler.Set("/Sheet1/A2", new() { ["value"] = "20" });

        // 2. Get + Verify initial state
        var nodeA1 = _excelHandler.Get("/Sheet1/A1");
        nodeA1.Text.Should().Be("10");

        // 3. Set (add formula)
        _excelHandler.Set("/Sheet1/A3", new() { ["formula"] = "SUM(A1:A2)" });

        // 4. Get + Verify modification
        var node = _excelHandler.Get("/Sheet1/A3");
        node.Format.Should().ContainKey("formula");
        node.Format["formula"].ToString().Should().Be("SUM(A1:A2)");

        // 5. Reopen + Verify persistence
        ReopenExcel();
        var node2 = _excelHandler.Get("/Sheet1/A3");
        node2.Format.Should().ContainKey("formula");
        node2.Format["formula"].ToString().Should().Be("SUM(A1:A2)");
    }

    // EDGE CASE: Excel number format round-trip.
    [Fact]
    public void Edge_Excel_NumberFormat_RoundTrip()
    {
        // 1. Add element (set cell with number format)
        _excelHandler.Set("/Sheet1/A1", new()
        {
            ["value"] = "0.123",
            ["numberformat"] = "0.00%"
        });

        // 2. Get + Verify initial state
        var node1 = _excelHandler.Get("/Sheet1/A1");
        node1.Format.Should().ContainKey("numberformat");
        node1.Format["numberformat"].ToString().Should().Be("0.00%");

        // 3. Set (modify number format)
        _excelHandler.Set("/Sheet1/A1", new() { ["numberformat"] = "#,##0.00" });

        // 4. Get + Verify modification
        var node2 = _excelHandler.Get("/Sheet1/A1");
        node2.Format.Should().ContainKey("numberformat");
        node2.Format["numberformat"].ToString().Should().Be("#,##0.00");

        // 5. Reopen + Verify persistence
        ReopenExcel();
        var node3 = _excelHandler.Get("/Sheet1/A1");
        node3.Format.Should().ContainKey("numberformat");
        node3.Format["numberformat"].ToString().Should().Be("#,##0.00");
    }

    // EDGE CASE: Excel cell alignment round-trip.
    [Fact]
    public void Edge_Excel_CellAlignment_RoundTrip()
    {
        // 1. Add element (set cell with alignment)
        _excelHandler.Set("/Sheet1/A1", new()
        {
            ["value"] = "Centered",
            ["halign"] = "center",
            ["wraptext"] = "true"
        });

        // 2. Get + Verify initial state
        var node1 = _excelHandler.Get("/Sheet1/A1");
        node1.Format.Should().ContainKey("halign");
        node1.Format["halign"].ToString().Should().Be("center");
        node1.Format.Should().ContainKey("alignment.wrapText");
        node1.Format["alignment.wrapText"].Should().Be(true);

        // 3. Set (modify alignment)
        _excelHandler.Set("/Sheet1/A1", new() { ["halign"] = "right" });

        // 4. Get + Verify modification
        var node2 = _excelHandler.Get("/Sheet1/A1");
        node2.Format.Should().ContainKey("halign");
        node2.Format["halign"].ToString().Should().Be("right");

        // 5. Reopen + Verify persistence
        ReopenExcel();
        var node3 = _excelHandler.Get("/Sheet1/A1");
        node3.Format.Should().ContainKey("halign");
        node3.Format["halign"].ToString().Should().Be("right");
        node3.Format.Should().ContainKey("alignment.wrapText");
        node3.Format["alignment.wrapText"].Should().Be(true);
    }

    // EDGE CASE: Excel sheet-level Set (tab color).
    [Fact]
    public void Edge_Excel_SheetLevel_TabColor()
    {
        // 1. Add element (set cell so sheet has content)
        _excelHandler.Set("/Sheet1/A1", new() { ["value"] = "Content" });

        // 2. Get + Verify initial state
        var node1 = _excelHandler.Get("/Sheet1");
        node1.Should().NotBeNull();

        // 3. Set tab color
        _excelHandler.Set("/Sheet1", new() { ["tabcolor"] = "FF0000" });

        // 4. Get + Verify modification
        var node2 = _excelHandler.Get("/Sheet1");
        node2.Format.Should().ContainKey("tabColor");
        node2.Format["tabColor"].ToString().Should().Contain("FF0000");

        // 5. Reopen + Verify persistence
        ReopenExcel();
        var node3 = _excelHandler.Get("/Sheet1");
        node3.Format.Should().ContainKey("tabColor");
        node3.Format["tabColor"].ToString().Should().Contain("FF0000");
    }

    // CONFIRMED BUG: Excel cell border color readback includes ARGB
    [Fact]
    public void Bug_Excel_BorderColor_Readback_IncludesArgbPrefix()
    {
        // 1. Add element (set cell with border)
        _excelHandler.Set("/Sheet1/A1", new()
        {
            ["value"] = "Bordered",
            ["border"] = "thin",
            ["border.color"] = "FF0000"
        });

        // 2. Get + Verify initial state
        var node1 = _excelHandler.Get("/Sheet1/A1");
        node1.Text.Should().Be("Bordered");

        // 3. Reopen to ensure persistence
        ReopenExcel();

        // 4. Get + Verify after reopen
        var node2 = _excelHandler.Get("/Sheet1/A1");

        // Check if any border side has a color
        var hasBorderColor = false;
        string? borderColorVal = null;
        foreach (var key in new[] { "border.left.color", "border.right.color", "border.top.color", "border.bottom.color" })
        {
            if (node2.Format.ContainsKey(key))
            {
                hasBorderColor = true;
                borderColorVal = node2.Format[key].ToString();
                break;
            }
        }

        if (hasBorderColor && borderColorVal != null)
        {
            // BUG: Border color readback returns "FFFF0000" (with ARGB prefix)
            // while font color readback correctly returns "FF0000"
            borderColorVal.Should().Be("FF0000",
                "border color readback should strip ARGB alpha prefix like font.color does, " +
                "but CellToNode doesn't strip it for border colors");
        }
    }

    // EDGE CASE: Excel data validation round-trip.
    [Fact]
    public void Edge_Excel_DataValidation_RoundTrip()
    {
        // 1. Add element
        _excelHandler.Add("/Sheet1", "validation", null, new()
        {
            ["sqref"] = "B1:B10",
            ["type"] = "list",
            ["formula1"] = "Yes,No,Maybe"
        });

        // 2. Get + Verify initial state
        var node1 = _excelHandler.Get("/Sheet1/validation[1]");
        node1.Format["type"].ToString().Should().Be("list");
        node1.Format["formula1"].ToString().Should().Be("Yes,No,Maybe");

        // 3. Set (modify formula)
        _excelHandler.Set("/Sheet1/validation[1]", new() { ["formula1"] = "Yes,No" });

        // 4. Get + Verify modification
        var node2 = _excelHandler.Get("/Sheet1/validation[1]");
        node2.Format["formula1"].ToString().Should().Be("Yes,No");

        // 5. Reopen + Verify persistence
        ReopenExcel();
        var results = _excelHandler.Query("validation");
        results.Should().HaveCountGreaterThan(0);
    }

    // EDGE CASE: PPTX table row height round-trip.
    [Fact]
    public void Edge_Pptx_TableRow_Height_RoundTrip()
    {
        // 1. Add elements
        _pptxHandler.Add("/", "slide", null, new());
        _pptxHandler.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "3",
            ["cols"] = "2"
        });

        // 2. Get + Verify initial state
        var node1 = _pptxHandler.Get("/slide[1]/table[1]", 1);
        node1.Children.Should().HaveCountGreaterThan(0);

        // 3. Set row height
        _pptxHandler.Set("/slide[1]/table[1]/tr[1]", new() { ["height"] = "1cm" });

        // 4. Get + Verify modification
        var node2 = _pptxHandler.Get("/slide[1]/table[1]", 1);
        node2.Children.Should().HaveCountGreaterThan(0);
        var tr1 = node2.Children[0];
        tr1.Format.Should().ContainKey("height");

        // 5. Reopen + Verify persistence
        ReopenPptx();
        var node3 = _pptxHandler.Get("/slide[1]/table[1]", 1);
        node3.Children.Should().HaveCountGreaterThan(0);
        node3.Children[0].Format.Should().ContainKey("height");
    }

    // EDGE CASE: PPTX table position round-trip.
    [Fact]
    public void Edge_Pptx_Table_Position_RoundTrip()
    {
        // 1. Add elements
        _pptxHandler.Add("/", "slide", null, new());
        _pptxHandler.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "2",
            ["cols"] = "2"
        });

        // 2. Get + Verify initial state
        var node1 = _pptxHandler.Get("/slide[1]/table[1]");
        node1.Should().NotBeNull();

        // 3. Set position
        _pptxHandler.Set("/slide[1]/table[1]", new()
        {
            ["x"] = "2cm",
            ["y"] = "3cm"
        });

        // 4. Get + Verify modification
        var node2 = _pptxHandler.Get("/slide[1]/table[1]");
        node2.Format.Should().ContainKey("x");
        node2.Format.Should().ContainKey("y");
        node2.Format["x"].ToString().Should().Be("2cm");
        node2.Format["y"].ToString().Should().Be("3cm");

        // 5. Reopen + Verify persistence
        ReopenPptx();
        var node3 = _pptxHandler.Get("/slide[1]/table[1]");
        node3.Format.Should().ContainKey("x");
        node3.Format["x"].ToString().Should().Be("2cm");
        node3.Format["y"].ToString().Should().Be("3cm");
    }

    // EDGE CASE: Word run with link round-trip.
    [Fact]
    public void Edge_Word_Run_Link_RoundTrip()
    {
        // 1. Add element
        _wordHandler.Add("/", "paragraph", null, new() { ["text"] = "Visit our site" });

        // 2. Get + Verify initial state
        var node1 = _wordHandler.Get("/body/p[1]");
        node1.Text.Should().Contain("Visit our site");

        // 3. Set link on run
        _wordHandler.Set("/body/p[1]/r[1]", new() { ["link"] = "https://example.com" });

        // 4. Get + Verify modification
        var node2 = _wordHandler.Get("/body/p[1]", 1);
        node2.Text.Should().Contain("Visit our site");

        // 5. Reopen + Verify persistence
        ReopenWord();
        var node3 = _wordHandler.Get("/body/p[1]", 1);
        node3.Text.Should().Contain("Visit our site");
    }

    // EDGE CASE: Word run font and size round-trip.
    [Fact]
    public void Edge_Word_Run_FontAndSize_RoundTrip()
    {
        // 1. Add element with font and size
        _wordHandler.Add("/", "paragraph", null, new()
        {
            ["text"] = "Styled text",
            ["font"] = "Courier New",
            ["size"] = "14"
        });

        // 2. Get + Verify initial state
        var node1 = _wordHandler.Get("/body/p[1]", 1);
        node1.Children.Should().HaveCountGreaterThan(0);
        var run1 = node1.Children[0];
        run1.Format.Should().ContainKey("font");
        run1.Format["font"].ToString().Should().Be("Courier New");
        run1.Format.Should().ContainKey("size");
        run1.Format["size"].ToString().Should().Be("14pt");

        // 3. Set (modify font)
        _wordHandler.Set("/body/p[1]/r[1]", new() { ["font"] = "Arial" });

        // 4. Get + Verify modification
        var node2 = _wordHandler.Get("/body/p[1]", 1);
        var run2 = node2.Children[0];
        run2.Format.Should().ContainKey("font");
        run2.Format["font"].ToString().Should().Be("Arial");

        // 5. Reopen + Verify persistence
        ReopenWord();
        var node3 = _wordHandler.Get("/body/p[1]", 1);
        var run3 = node3.Children[0];
        run3.Format.Should().ContainKey("font");
        run3.Format["font"].ToString().Should().Be("Arial");
        run3.Format.Should().ContainKey("size");
        run3.Format["size"].ToString().Should().Be("14pt");
    }

    // CONFIRMED BUG: PPTX table cell "strike" with "single" value
    [Fact]
    public void Bug_Pptx_TableCell_Strike_SingleText_CreatesInvalidXml()
    {
        // 1. Add elements
        _pptxHandler.Add("/", "slide", null, new());
        _pptxHandler.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "2",
            ["cols"] = "2"
        });

        // 2. Get + Verify initial state
        var node1 = _pptxHandler.Get("/slide[1]/table[1]", 1);
        node1.Should().NotBeNull();

        // 3. Set strike with "single" on table cell
        _pptxHandler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
        {
            ["text"] = "Single strike",
            ["strike"] = "single"
        });

        // 4. Get + Verify modification
        var raw = _pptxHandler.Raw("/slide[1]");
        // BUG: "single" -> not IsTruthy -> passes "single" to TextStrikeValues
        // but valid enum value is "sngStrike"
        raw.Should().Contain("sngStrike",
            "table cell strike='single' should produce 'sngStrike' in XML, " +
            "but the code passes 'single' directly to TextStrikeValues constructor");
    }

    // EDGE CASE: Excel cell with hyperlink.
    [Fact]
    public void Edge_Excel_Cell_Hyperlink_RoundTrip()
    {
        // 1. Add element (set cell with hyperlink)
        _excelHandler.Set("/Sheet1/A1", new()
        {
            ["value"] = "Click me",
            ["link"] = "https://example.com"
        });

        // 2. Get + Verify initial state
        var node1 = _excelHandler.Get("/Sheet1/A1");
        node1.Text.Should().Be("Click me");
        node1.Format.Should().ContainKey("link");
        node1.Format["link"].ToString().Should().Contain("example.com");

        // 3. Set (modify link)
        _excelHandler.Set("/Sheet1/A1", new() { ["link"] = "https://other.com" });

        // 4. Get + Verify modification
        var node2 = _excelHandler.Get("/Sheet1/A1");
        node2.Format.Should().ContainKey("link");
        node2.Format["link"].ToString().Should().Contain("other.com");

        // 5. Reopen + Verify persistence
        ReopenExcel();
        var node3 = _excelHandler.Get("/Sheet1/A1");
        node3.Text.Should().Be("Click me");
        node3.Format.Should().ContainKey("link");
        node3.Format["link"].ToString().Should().Contain("other.com");
    }

    // EDGE CASE: Excel conditional formatting.
    [Fact]
    public void Edge_Excel_ConditionalFormatting_DataBar()
    {
        // 1. Add elements (set cell values)
        _excelHandler.Set("/Sheet1/A1", new() { ["value"] = "10" });
        _excelHandler.Set("/Sheet1/A2", new() { ["value"] = "50" });
        _excelHandler.Set("/Sheet1/A3", new() { ["value"] = "90" });

        // 2. Get + Verify initial state
        var nodeA1 = _excelHandler.Get("/Sheet1/A1");
        nodeA1.Text.Should().Be("10");

        // 3. Add conditional formatting
        _excelHandler.Add("/Sheet1", "conditionalformatting", null, new()
        {
            ["sqref"] = "A1:A3",
            ["type"] = "dataBar",
            ["color"] = "4472C4"
        });

        // 4. Get + Verify modification
        var results = _excelHandler.Query("conditionalformatting");
        results.Should().HaveCountGreaterThan(0);

        // 5. Reopen + Verify persistence
        ReopenExcel();
        var results2 = _excelHandler.Query("conditionalformatting");
        results2.Should().HaveCountGreaterThan(0);
    }

    // EDGE CASE: Word document properties.
    [Fact]
    public void Edge_Word_DocumentProperties_RoundTrip()
    {
        // 1. Add element (paragraph for content)
        _wordHandler.Add("/", "paragraph", null, new() { ["text"] = "Document content" });

        // 2. Get + Verify initial state
        var node1 = _wordHandler.Get("/");
        node1.Should().NotBeNull();

        // 3. Set document properties
        _wordHandler.Set("/", new()
        {
            ["pagewidth"] = "12240",
            ["pageheight"] = "15840",
            ["margintop"] = "1440"
        });

        // 4. Get + Verify modification
        var node2 = _wordHandler.Get("/");
        node2.Format.Should().ContainKey("pageWidth");

        // 5. Reopen + Verify persistence
        ReopenWord();
        var node3 = _wordHandler.Get("/");
        node3.Format.Should().ContainKey("pageWidth");
    }

    // EDGE CASE: PPTX presentation-level properties.
    [Fact]
    public void Edge_Pptx_PresentationLevel_SlideSize()
    {
        // 1. Add element (slide)
        _pptxHandler.Add("/", "slide", null, new());

        // 2. Get + Verify initial state
        var node1 = _pptxHandler.Get("/");
        node1.Should().NotBeNull();

        // 3. Set slide size
        _pptxHandler.Set("/", new() { ["slidesize"] = "4:3" });

        // 4. Get + Verify modification
        var node2 = _pptxHandler.Get("/");
        node2.Format.Should().ContainKey("slideWidth");
        node2.Format["slideWidth"].ToString().Should().Be("25.4cm");

        // 5. Reopen + Verify persistence
        ReopenPptx();
        var node3 = _pptxHandler.Get("/");
        node3.Format.Should().ContainKey("slideWidth");
        node3.Format["slideWidth"].ToString().Should().Be("25.4cm");
    }

    // ==================== Helper methods ====================

    private static void CreateMinimalPng(string path)
    {
        byte[] png = {
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
            0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
            0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
            0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
            0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41,
            0x54, 0x08, 0xD7, 0x63, 0xF8, 0xCF, 0xC0, 0x00,
            0x00, 0x00, 0x02, 0x00, 0x01, 0xE2, 0x21, 0xBC,
            0x33, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E,
            0x44, 0xAE, 0x42, 0x60, 0x82
        };
        File.WriteAllBytes(path, png);
    }
}
