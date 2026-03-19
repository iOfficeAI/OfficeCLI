// Bug hunt Part 32 — Cross-handler bugs and edge cases:
// 1. PPTX shape hyperlink with special characters in URL
// 2. PPTX chart position Set with EMU units
// 3. Word paragraph keepNext/keepLines persistence
// 4. Excel autofilter persistence
// 5. PPTX slide removal and re-indexing
// 6. Word table cell Set properties
// 7. PPTX shape flip round-trip
// 8. Excel frozen panes persistence
// 9. PPTX table cell valign readback
// 10. Word bookmark text Set
// 11. Excel named range scope round-trip
// 12. PPTX shape preset geometry round-trip
// 13. Word run color with # prefix round-trip
// 14. Excel cell border diagonal round-trip
// 15. PPTX Remove shape and re-index

using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

public class BugHuntPart32 : IDisposable
{
    private readonly string _docxPath;
    private readonly string _xlsxPath;
    private readonly string _pptxPath;
    private WordHandler _wordHandler;
    private ExcelHandler _excelHandler;
    private PowerPointHandler _pptxHandler;

    public BugHuntPart32()
    {
        _docxPath = Path.Combine(Path.GetTempPath(), $"bughunt32_{Guid.NewGuid():N}.docx");
        _xlsxPath = Path.Combine(Path.GetTempPath(), $"bughunt32_{Guid.NewGuid():N}.xlsx");
        _pptxPath = Path.Combine(Path.GetTempPath(), $"bughunt32_{Guid.NewGuid():N}.pptx");
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

    // =================================================================
    // EDGE CASE: PPTX shape hyperlink Set and Get round-trip.
    // =================================================================

    [Fact]
    public void Edge_Pptx_Shape_Hyperlink_RoundTrip()
    {
        // 1. Add
        _pptxHandler.Add("/", "slide", null, new());
        _pptxHandler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Click here",
            ["link"] = "https://example.com/page?q=test&lang=en"
        });

        // 2. Get + Verify initial state
        var node = _pptxHandler.Get("/slide[1]/shape[1]");
        node.Should().NotBeNull();
        node.Text.Should().Be("Click here");
        node.Format.Should().ContainKey("link");
        node.Format["link"].ToString().Should().Contain("example.com");

        // 3. Set (modify — remove hyperlink)
        _pptxHandler.Set("/slide[1]/shape[1]", new() { ["link"] = "none" });

        // 4. Get + Verify modification
        var node2 = _pptxHandler.Get("/slide[1]/shape[1]");
        node2.Format.Should().NotContainKey("link",
            "setting link=none should remove the hyperlink");

        // 5. Reopen + Verify persistence
        ReopenPptx();
        var node3 = _pptxHandler.Get("/slide[1]/shape[1]");
        node3.Text.Should().Be("Click here");
        node3.Format.Should().NotContainKey("link");
    }

    // =================================================================
    // EDGE CASE: PPTX chart position Set with EMU units.
    // =================================================================

    [Fact]
    public void Edge_Pptx_Chart_Position_RoundTrip()
    {
        // 1. Add
        _pptxHandler.Add("/", "slide", null, new());
        _pptxHandler.Add("/slide[1]", "chart", null, new()
        {
            ["type"] = "bar",
            ["data"] = "Sales:10,20,30"
        });

        // 2. Get + Verify initial state
        var initial = _pptxHandler.Get("/slide[1]/chart[1]");
        initial.Should().NotBeNull();

        // 3. Set (modify position)
        _pptxHandler.Set("/slide[1]/chart[1]", new()
        {
            ["x"] = "2cm",
            ["y"] = "3cm",
            ["width"] = "10cm",
            ["height"] = "8cm"
        });

        // 4. Get + Verify modification
        var node = _pptxHandler.Get("/slide[1]/chart[1]");
        node.Format.Should().ContainKey("x");
        node.Format.Should().ContainKey("y");
        node.Format["x"].ToString().Should().Be("2cm");
        node.Format["y"].ToString().Should().Be("3cm");

        // 5. Reopen + Verify persistence
        ReopenPptx();
        var persisted = _pptxHandler.Get("/slide[1]/chart[1]");
        persisted.Format["x"].ToString().Should().Be("2cm");
        persisted.Format["y"].ToString().Should().Be("3cm");
    }

    // =================================================================
    // CONFIRMED BUG: PPTX slide Remove and shape re-indexing.
    // After removing a slide, the remaining slides should be
    // re-indexed. But Get on slide[2] after removing slide[1]
    // should now access what was previously slide[2].
    // =================================================================

    [Fact]
    public void Edge_Pptx_SlideRemove_ReIndexes()
    {
        // 1. Add
        _pptxHandler.Add("/", "slide", null, new() { ["title"] = "Slide A" });
        _pptxHandler.Add("/", "slide", null, new() { ["title"] = "Slide B" });
        _pptxHandler.Add("/", "slide", null, new() { ["title"] = "Slide C" });

        // 2. Get + Verify initial state
        var a = _pptxHandler.Get("/slide[1]");
        a.Children.Should().Contain(c => c.Text != null && c.Text.Contains("Slide A"));
        var b = _pptxHandler.Get("/slide[2]");
        b.Children.Should().Contain(c => c.Text != null && c.Text.Contains("Slide B"));

        // 3. Set (modify — remove first slide)
        _pptxHandler.Remove("/slide[1]");

        // 4. Get + Verify re-indexing
        var newFirst = _pptxHandler.Get("/slide[1]");
        newFirst.Children.Should().Contain(c => c.Text != null && c.Text.Contains("Slide B"),
            "after removing slide[1], the new slide[1] should be 'Slide B'");

        // 5. Reopen + Verify persistence
        ReopenPptx();
        var persisted = _pptxHandler.Get("/slide[1]");
        persisted.Children.Should().Contain(c => c.Text != null && c.Text.Contains("Slide B"));
    }

    // =================================================================
    // EDGE CASE: PPTX shape flip round-trip.
    // =================================================================

    [Fact]
    public void Edge_Pptx_Shape_Flip_RoundTrip()
    {
        // 1. Add
        _pptxHandler.Add("/", "slide", null, new());
        _pptxHandler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Flippable"
        });

        // 2. Get + Verify initial state
        var initial = _pptxHandler.Get("/slide[1]/shape[1]");
        initial.Should().NotBeNull();
        initial.Text.Should().Be("Flippable");

        // 3. Set (modify — flip horizontally)
        _pptxHandler.Set("/slide[1]/shape[1]", new() { ["fliph"] = "true" });

        // 4. Get + Verify modification
        var node = _pptxHandler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("flipH");
        node.Format["flipH"].Should().Be(true);

        // 5. Set again (unset flip)
        _pptxHandler.Set("/slide[1]/shape[1]", new() { ["fliph"] = "false" });

        // 6. Get + Verify
        var node2 = _pptxHandler.Get("/slide[1]/shape[1]");
        if (node2.Format.ContainsKey("flipH"))
            node2.Format["flipH"].Should().Be(false);
    }

    // =================================================================
    // EDGE CASE: PPTX shape preset geometry change and readback.
    // =================================================================

    [Fact]
    public void Edge_Pptx_Shape_PresetGeometry_Change()
    {
        // 1. Add with preset
        _pptxHandler.Add("/", "slide", null, new());
        _pptxHandler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Default rect",
            ["preset"] = "ellipse"
        });

        // 2. Get + Verify initial state
        var node1 = _pptxHandler.Get("/slide[1]/shape[1]");
        node1.Format.Should().ContainKey("preset");
        node1.Format["preset"].ToString().Should().Be("ellipse");

        // 3. Set (modify — change to triangle)
        _pptxHandler.Set("/slide[1]/shape[1]", new() { ["preset"] = "triangle" });

        // 4. Get + Verify modification
        var node2 = _pptxHandler.Get("/slide[1]/shape[1]");
        node2.Format["preset"].ToString().Should().Be("triangle");

        // 5. Reopen + Verify persistence
        ReopenPptx();
        var persisted = _pptxHandler.Get("/slide[1]/shape[1]");
        persisted.Format["preset"].ToString().Should().Be("triangle");
    }

    // =================================================================
    // EDGE CASE: Word run color with # prefix should be handled.
    // The code strips # prefix: value.TrimStart('#').ToUpperInvariant()
    // =================================================================

    [Fact]
    public void Edge_Word_Run_Color_WithHashPrefix()
    {
        // 1. Add with # prefix color
        _wordHandler.Add("/", "paragraph", null, new()
        {
            ["text"] = "Colored text",
            ["color"] = "#FF0000"
        });

        // 2. Get + Verify initial state (color stored without #)
        var node = _wordHandler.Get("/body/p[1]", 1);
        node.Children.Should().HaveCountGreaterThan(0);
        var run = node.Children[0];
        run.Format.Should().ContainKey("color");
        run.Format["color"].ToString().Should().Be("FF0000");

        // 3. Set (modify color to a different value with #)
        _wordHandler.Set("/body/p[1]", new() { ["color"] = "#00FF00" });

        // 4. Get + Verify modification
        var node2 = _wordHandler.Get("/body/p[1]", 1);
        var run2 = node2.Children[0];
        run2.Format["color"].ToString().Should().Be("00FF00");

        // 5. Reopen + Verify persistence
        ReopenWord();
        var node3 = _wordHandler.Get("/body/p[1]", 1);
        var run3 = node3.Children[0];
        run3.Format["color"].ToString().Should().Be("00FF00");
    }

    // =================================================================
    // EDGE CASE: Excel frozen panes.
    // =================================================================

    [Fact]
    public void Edge_Excel_FrozenPanes_RoundTrip()
    {
        // 1. Add data
        _excelHandler.Set("/Sheet1/A1", new() { ["value"] = "Header" });

        // 2. Get + Verify initial state (no freeze yet)
        var initial = _excelHandler.Get("/Sheet1");
        initial.Should().NotBeNull();

        // 3. Set (modify — freeze panes)
        _excelHandler.Set("/Sheet1", new() { ["freeze"] = "A2" });

        // 4. Get + Verify modification
        var node = _excelHandler.Get("/Sheet1");
        node.Format.Should().ContainKey("freeze");

        // 5. Reopen + Verify persistence
        ReopenExcel();
        var persisted = _excelHandler.Get("/Sheet1");
        persisted.Format.Should().ContainKey("freeze");
    }

    // =================================================================
    // CONFIRMED BUG: Word paragraph Set with "text" property on a
    // paragraph path doesn't have a direct text handler. Setting
    // { ["text"] = "new text" } on a paragraph path falls through
    // to the default case and gets added as unsupported.
    //
    // The paragraph switch handles: formula, liststyle, start,
    // size/font/bold/italic/color/highlight/underline/strike
    // But NOT "text" directly.
    // =================================================================

    [Fact]
    public void Bug_Word_Paragraph_Set_Text_NotSupported()
    {
        // 1. Add
        _wordHandler.Add("/", "paragraph", null, new()
        {
            ["text"] = "Original text"
        });

        // 2. Get + Verify initial state
        var initial = _wordHandler.Get("/body/p[1]");
        initial.Text.Should().Contain("Original text");

        // 3. Set (modify — try to set text on paragraph)
        var unsupported = _wordHandler.Set("/body/p[1]", new()
        {
            ["text"] = "Updated text"
        });

        // 4. Verify — BUG: "text" is returned as unsupported because
        // the paragraph Set handler doesn't have a "text" case.
        unsupported.Should().NotContain("text",
            "setting 'text' on a paragraph path should be supported, " +
            "but the paragraph Set handler doesn't have a 'text' case");

        // 5. Get + Verify modification (if supported)
        var node = _wordHandler.Get("/body/p[1]");
        node.Text.Should().Contain("Updated text");
    }

    // =================================================================
    // EDGE CASE: Excel named range round-trip.
    // =================================================================

    [Fact]
    public void Edge_Excel_NamedRange_RoundTrip()
    {
        // 1. Add
        _excelHandler.Add("/", "namedrange", null, new()
        {
            ["name"] = "TestRange",
            ["ref"] = "Sheet1!$A$1:$C$10"
        });

        // 2. Get + Verify initial state
        var node = _excelHandler.Get("/namedrange[1]");
        node.Should().NotBeNull();
        node.Format.Should().ContainKey("name");
        node.Format["name"].ToString().Should().Be("TestRange");

        // 3. Set data in the range area
        _excelHandler.Set("/Sheet1/A1", new() { ["value"] = "InRange" });

        // 4. Get + Verify range still valid
        var node2 = _excelHandler.Get("/namedrange[1]");
        node2.Format["name"].ToString().Should().Be("TestRange");

        // 5. Reopen + Verify persistence
        ReopenExcel();
        var persisted = _excelHandler.Get("/namedrange[1]");
        persisted.Should().NotBeNull();
        persisted.Format["name"].ToString().Should().Be("TestRange");
    }

    // =================================================================
    // EDGE CASE: Excel autofilter Set and Get.
    // =================================================================

    [Fact]
    public void Edge_Excel_AutoFilter_RoundTrip()
    {
        // 1. Add data
        _excelHandler.Set("/Sheet1/A1", new() { ["value"] = "Name" });
        _excelHandler.Set("/Sheet1/B1", new() { ["value"] = "Age" });
        _excelHandler.Set("/Sheet1/A2", new() { ["value"] = "Alice" });
        _excelHandler.Set("/Sheet1/B2", new() { ["value"] = "30" });

        // 2. Get + Verify initial state (no autofilter yet)
        var initial = _excelHandler.Get("/Sheet1");
        initial.Should().NotBeNull();

        // 3. Set (modify — add autofilter)
        _excelHandler.Set("/Sheet1", new() { ["autofilter"] = "A1:B2" });

        // 4. Get + Verify modification
        var node = _excelHandler.Get("/Sheet1");
        node.Format.Should().ContainKey("autofilter");
        node.Format["autofilter"].ToString().Should().Be("A1:B2");

        // 5. Reopen + Verify persistence
        ReopenExcel();
        var persisted = _excelHandler.Get("/Sheet1");
        persisted.Format.Should().ContainKey("autofilter");
        persisted.Format["autofilter"].ToString().Should().Be("A1:B2");
    }

    // =================================================================
    // EDGE CASE: PPTX table cell valign readback.
    // =================================================================

    [Fact]
    public void Edge_Pptx_TableCell_Valign_RoundTrip()
    {
        // 1. Add
        _pptxHandler.Add("/", "slide", null, new());
        _pptxHandler.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "2",
            ["cols"] = "2"
        });

        // 2. Get + Verify initial state
        var table = _pptxHandler.Get("/slide[1]/table[1]");
        table.Should().NotBeNull();

        // 3. Set (modify — valign on cell)
        _pptxHandler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
        {
            ["text"] = "Centered",
            ["valign"] = "middle"
        });

        // 4. Get + Verify via raw XML
        var raw = _pptxHandler.Raw("/slide[1]");
        raw.Should().Contain("anchor=\"ctr\"",
            "valign=middle should produce anchor=ctr in the XML");

        // 5. Reopen + Verify persistence
        ReopenPptx();
        var rawPersisted = _pptxHandler.Raw("/slide[1]");
        rawPersisted.Should().Contain("anchor=\"ctr\"");
    }

    // =================================================================
    // CONFIRMED BUG: PPTX shape Add with both "text" multiline and
    // "align" loses alignment on some paragraphs.
    // =================================================================

    [Fact]
    public void Edge_Pptx_Add_Shape_MultilineText_WithAlign()
    {
        // 1. Add with multiline text and alignment
        _pptxHandler.Add("/", "slide", null, new());
        _pptxHandler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Line 1\\nLine 2\\nLine 3",
            ["align"] = "right"
        });

        // 2. Get + Verify initial state
        var node = _pptxHandler.Get("/slide[1]/shape[1]");
        node.Should().NotBeNull();

        // 3. Verify all paragraphs have right alignment via raw XML
        var raw = _pptxHandler.Raw("/slide[1]");
        var rightAlignCount = System.Text.RegularExpressions.Regex.Matches(raw, @"algn=""r""").Count;
        rightAlignCount.Should().BeGreaterThanOrEqualTo(3,
            "all paragraphs in a multiline shape should get the alignment");

        // 4. Set (modify — change alignment)
        _pptxHandler.Set("/slide[1]/shape[1]", new() { ["align"] = "center" });

        // 5. Get + Verify modification
        var raw2 = _pptxHandler.Raw("/slide[1]");
        var centerAlignCount = System.Text.RegularExpressions.Regex.Matches(raw2, @"algn=""ctr""").Count;
        centerAlignCount.Should().BeGreaterThanOrEqualTo(1);
    }

    // =================================================================
    // EDGE CASE: PPTX shape 3D bevel round-trip.
    // =================================================================

    [Fact]
    public void Edge_Pptx_Shape_Bevel_RoundTrip()
    {
        // 1. Add
        _pptxHandler.Add("/", "slide", null, new());
        _pptxHandler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Beveled shape",
            ["fill"] = "4472C4"
        });

        // 2. Get + Verify initial state
        var initial = _pptxHandler.Get("/slide[1]/shape[1]");
        initial.Should().NotBeNull();
        initial.Text.Should().Be("Beveled shape");
        initial.Format["fill"].ToString().Should().Be("4472C4");

        // 3. Set (modify — add bevel)
        _pptxHandler.Set("/slide[1]/shape[1]", new()
        {
            ["bevel"] = "circle-8-6"
        });

        // 4. Get + Verify modification
        var node = _pptxHandler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("bevel");
        var bevelStr = node.Format["bevel"].ToString()!;
        bevelStr.Should().Contain("circle");

        // 5. Reopen + Verify persistence
        ReopenPptx();
        var persisted = _pptxHandler.Get("/slide[1]/shape[1]");
        persisted.Format.Should().ContainKey("bevel");
        persisted.Format["bevel"].ToString().Should().Contain("circle");
    }

    // =================================================================
    // EDGE CASE: PPTX shape 3D depth and material round-trip.
    // =================================================================

    [Fact]
    public void Edge_Pptx_Shape_3D_Depth_Material_RoundTrip()
    {
        // 1. Add
        _pptxHandler.Add("/", "slide", null, new());
        _pptxHandler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "3D shape",
            ["fill"] = "ED7D31"
        });

        // 2. Get + Verify initial state
        var initial = _pptxHandler.Get("/slide[1]/shape[1]");
        initial.Should().NotBeNull();
        initial.Text.Should().Be("3D shape");

        // 3. Set (modify — add 3D properties)
        _pptxHandler.Set("/slide[1]/shape[1]", new()
        {
            ["depth"] = "10",
            ["material"] = "metal"
        });

        // 4. Get + Verify modification
        var node = _pptxHandler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("depth");
        node.Format["depth"].ToString().Should().Be("10");
        node.Format.Should().ContainKey("material");
        node.Format["material"].ToString().Should().Be("metal");

        // 5. Reopen + Verify persistence
        ReopenPptx();
        var persisted = _pptxHandler.Get("/slide[1]/shape[1]");
        persisted.Format["depth"].ToString().Should().Be("10");
        persisted.Format["material"].ToString().Should().Be("metal");
    }

    // =================================================================
    // EDGE CASE: Word bookmark text Set and Get.
    // =================================================================

    [Fact]
    public void Edge_Word_Bookmark_RoundTrip()
    {
        // 1. Add paragraph and bookmark
        _wordHandler.Add("/", "paragraph", null, new()
        {
            ["text"] = "Before bookmark text After"
        });
        _wordHandler.Add("/body/p[1]", "bookmark", null, new()
        {
            ["name"] = "MyBookmark"
        });

        // 2. Get + Verify initial state
        var bookmarks = _wordHandler.Query("bookmark");
        bookmarks.Should().HaveCountGreaterThan(0);
        bookmarks[0].Format["name"].ToString().Should().Be("MyBookmark");

        // 3. Reopen + Verify persistence
        ReopenWord();
        var persisted = _wordHandler.Query("bookmark");
        persisted.Should().HaveCountGreaterThan(0);
        persisted[0].Format["name"].ToString().Should().Be("MyBookmark");
    }

    // =================================================================
    // EDGE CASE: PPTX slide notes Set and Get.
    // =================================================================

    [Fact]
    public void Edge_Pptx_SlideNotes_RoundTrip()
    {
        // 1. Add
        _pptxHandler.Add("/", "slide", null, new() { ["title"] = "Noted slide" });

        // 2. Get + Verify initial state (slide exists)
        var slide = _pptxHandler.Get("/slide[1]");
        slide.Should().NotBeNull();

        // 3. Set (modify — add notes)
        _pptxHandler.Set("/slide[1]", new() { ["notes"] = "Speaker notes here" });

        // 4. Get + Verify modification
        var node = _pptxHandler.Get("/slide[1]/notes");
        node.Should().NotBeNull();
        node.Text.Should().Contain("Speaker notes here");

        // 5. Set again (update notes)
        _pptxHandler.Set("/slide[1]/notes", new() { ["text"] = "Updated notes" });

        // 6. Get + Verify second modification
        var node2 = _pptxHandler.Get("/slide[1]/notes");
        node2.Text.Should().Contain("Updated notes");

        // 7. Reopen + Verify persistence
        ReopenPptx();
        var persisted = _pptxHandler.Get("/slide[1]/notes");
        persisted.Text.Should().Contain("Updated notes");
    }

    // =================================================================
    // CONFIRMED BUG: PPTX shape opacity during Add handles SchemeColor
    // correctly, but the readback in NodeBuilder only checks
    // RgbColorModelHex for alpha, not SchemeColor. So opacity might
    // not be read back for scheme colors even though it was set.
    // =================================================================

    [Fact]
    public void Bug_Pptx_Opacity_Readback_SchemeColor()
    {
        // 1. Add with scheme color and opacity
        _pptxHandler.Add("/", "slide", null, new());
        _pptxHandler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Semi-transparent",
            ["fill"] = "accent3",
            ["opacity"] = "0.7"
        });

        // 2. Get + Verify initial state
        var node = _pptxHandler.Get("/slide[1]/shape[1]");
        node.Should().NotBeNull();
        node.Text.Should().Be("Semi-transparent");

        // 3. Verify opacity readback — BUG: NodeBuilder only checks
        // RgbColorModelHex for alpha, not SchemeColor
        node.Format.Should().ContainKey("opacity",
            "opacity readback should work for SchemeColor fills, " +
            "but NodeBuilder only checks RgbColorModelHex for alpha element");

        if (node.Format.ContainsKey("opacity"))
        {
            node.Format["opacity"].ToString().Should().Be("0.7");

            // 4. Set (modify opacity)
            _pptxHandler.Set("/slide[1]/shape[1]", new() { ["opacity"] = "0.5" });

            // 5. Get + Verify modification
            var modified = _pptxHandler.Get("/slide[1]/shape[1]");
            modified.Format["opacity"].ToString().Should().Be("0.5");
        }
    }

    // =================================================================
    // EDGE CASE: PPTX slide transition persistence.
    // =================================================================

    [Fact]
    public void Edge_Pptx_SlideTransition_Persistence()
    {
        // 1. Add
        _pptxHandler.Add("/", "slide", null, new());

        // 2. Get + Verify initial state
        var slide = _pptxHandler.Get("/slide[1]");
        slide.Should().NotBeNull();

        // 3. Set (modify — add transition)
        _pptxHandler.Set("/slide[1]", new() { ["transition"] = "fade" });

        // 4. Get + Verify via raw XML
        var raw = _pptxHandler.Raw("/slide[1]");
        raw.Should().Contain("fade",
            "slide transition should be present in raw XML");

        // 5. Reopen + Verify persistence
        ReopenPptx();
        var rawPersisted = _pptxHandler.Raw("/slide[1]");
        rawPersisted.Should().Contain("fade",
            "slide transition should persist after reopen");
    }

    // =================================================================
    // EDGE CASE: Excel data validation with between operator.
    // =================================================================

    [Fact]
    public void Edge_Excel_Validation_Between_RoundTrip()
    {
        // 1. Add validation
        _excelHandler.Add("/Sheet1", "validation", null, new()
        {
            ["sqref"] = "C1:C10",
            ["type"] = "whole",
            ["operator"] = "between",
            ["formula1"] = "1",
            ["formula2"] = "100"
        });

        // 2. Get + Verify initial state
        var node = _excelHandler.Get("/Sheet1/validation[1]");
        node.Should().NotBeNull();
        node.Format["type"].ToString().Should().Be("whole");
        node.Format["formula1"].ToString().Should().Be("1");
        node.Format["formula2"].ToString().Should().Be("100");

        // 3. Reopen + Verify persistence
        ReopenExcel();
        var persisted = _excelHandler.Get("/Sheet1/validation[1]");
        persisted.Format["type"].ToString().Should().Be("whole");
        persisted.Format["formula1"].ToString().Should().Be("1");
        persisted.Format["formula2"].ToString().Should().Be("100");
    }

    // =================================================================
    // EDGE CASE: PPTX shape z-order manipulation.
    // =================================================================

    [Fact]
    public void Edge_Pptx_Shape_ZOrder_Manipulation()
    {
        // 1. Add
        _pptxHandler.Add("/", "slide", null, new());
        _pptxHandler.Add("/slide[1]", "shape", null, new() { ["text"] = "Back" });
        _pptxHandler.Add("/slide[1]", "shape", null, new() { ["text"] = "Middle" });
        _pptxHandler.Add("/slide[1]", "shape", null, new() { ["text"] = "Front" });

        // 2. Get + Verify initial state
        var node1 = _pptxHandler.Get("/slide[1]/shape[1]");
        node1.Format.Should().ContainKey("zorder");
        var z1 = int.Parse(node1.Format["zorder"].ToString()!);

        var node3 = _pptxHandler.Get("/slide[1]/shape[3]");
        var z3 = int.Parse(node3.Format["zorder"].ToString()!);
        z3.Should().BeGreaterThan(z1, "shape[3] should have higher z-order than shape[1]");

        // 3. Set (modify — bring shape[1] to front)
        _pptxHandler.Set("/slide[1]/shape[1]", new() { ["zorder"] = "front" });

        // 4. Get + Verify modification
        var node1After = _pptxHandler.Get("/slide[1]/shape[3]");
        node1After.Text.Should().Be("Back",
            "after bringing shape[1] (Back) to front, it should now be shape[3]");

        // 5. Reopen + Verify persistence
        ReopenPptx();
        var persisted = _pptxHandler.Get("/slide[1]/shape[3]");
        persisted.Text.Should().Be("Back");
    }

    // =================================================================
    // EDGE CASE: Word table Add and query.
    // =================================================================

    [Fact]
    public void Edge_Word_Table_Add_Query()
    {
        // 1. Add
        _wordHandler.Add("/", "table", null, new()
        {
            ["rows"] = "3",
            ["cols"] = "2"
        });

        // 2. Get + Verify initial state
        var tables = _wordHandler.Query("table");
        tables.Should().HaveCountGreaterThan(0);
        tables[0].Format.Should().ContainKey("cols");

        // 3. Set (modify — set cell text)
        _wordHandler.Set("/body/tbl[1]/tr[1]/tc[1]", new() { ["text"] = "Cell A1" });

        // 4. Get + Verify modification
        var cell = _wordHandler.Get("/body/tbl[1]/tr[1]/tc[1]");
        cell.Text.Should().Contain("Cell A1");

        // 5. Reopen + Verify persistence
        ReopenWord();
        var persisted = _wordHandler.Get("/body/tbl[1]/tr[1]/tc[1]");
        persisted.Text.Should().Contain("Cell A1");
    }

    // =================================================================
    // EDGE CASE: Word paragraph with page break before.
    // =================================================================

    [Fact]
    public void Edge_Word_Paragraph_PageBreakBefore()
    {
        // 1. Add
        _wordHandler.Add("/", "paragraph", null, new()
        {
            ["text"] = "First page"
        });
        _wordHandler.Add("/", "paragraph", null, new()
        {
            ["text"] = "New page",
            ["pagebreakbefore"] = "true"
        });

        // 2. Get + Verify initial state
        var node = _wordHandler.Get("/body/p[2]");
        node.Text.Should().Be("New page");
        node.Format.Should().ContainKey("pagebreakbefore");

        // 3. Set (modify — remove page break before)
        _wordHandler.Set("/body/p[2]", new() { ["pagebreakbefore"] = "false" });

        // 4. Get + Verify modification
        var modified = _wordHandler.Get("/body/p[2]");
        modified.Text.Should().Be("New page");

        // 5. Reopen + Verify persistence
        ReopenWord();
        var persisted = _wordHandler.Get("/body/p[2]");
        persisted.Text.Should().Be("New page");
    }

    // =================================================================
    // EDGE CASE: Excel sheet data persistence.
    // =================================================================

    [Fact]
    public void Edge_Excel_SheetRename_RoundTrip()
    {
        // 1. Add data to sheet
        _excelHandler.Set("/Sheet1/A1", new() { ["value"] = "Test data" });
        _excelHandler.Set("/Sheet1/B1", new() { ["value"] = "More data" });

        // 2. Get + Verify initial state
        var nodeA = _excelHandler.Get("/Sheet1/A1");
        nodeA.Text.Should().Be("Test data");
        var nodeB = _excelHandler.Get("/Sheet1/B1");
        nodeB.Text.Should().Be("More data");

        // 3. Set (modify — update a cell)
        _excelHandler.Set("/Sheet1/A1", new() { ["value"] = "Updated data" });

        // 4. Get + Verify modification
        var modified = _excelHandler.Get("/Sheet1/A1");
        modified.Text.Should().Be("Updated data");

        // 5. Reopen + Verify persistence
        ReopenExcel();
        var persisted = _excelHandler.Get("/Sheet1/A1");
        persisted.Text.Should().Be("Updated data");
        var persistedB = _excelHandler.Get("/Sheet1/B1");
        persistedB.Text.Should().Be("More data");
    }

    // =================================================================
    // EDGE CASE: PPTX multiple slides with different backgrounds.
    // =================================================================

    [Fact]
    public void Edge_Pptx_MultipleSlides_DifferentBackgrounds()
    {
        // 1. Add slides with different backgrounds
        _pptxHandler.Add("/", "slide", null, new() { ["background"] = "FF0000" });
        _pptxHandler.Add("/", "slide", null, new() { ["background"] = "00FF00" });
        _pptxHandler.Add("/", "slide", null, new() { ["background"] = "0000FF" });

        // 2. Get + Verify initial state
        var node1 = _pptxHandler.Get("/slide[1]");
        var node2 = _pptxHandler.Get("/slide[2]");
        var node3 = _pptxHandler.Get("/slide[3]");
        node1.Format["background"].ToString().Should().Contain("FF0000");
        node2.Format["background"].ToString().Should().Contain("00FF00");
        node3.Format["background"].ToString().Should().Contain("0000FF");

        // 3. Set (modify — change slide[2] background)
        _pptxHandler.Set("/slide[2]", new() { ["background"] = "FFFF00" });

        // 4. Get + Verify modification
        var modified = _pptxHandler.Get("/slide[2]");
        modified.Format["background"].ToString().Should().Contain("FFFF00");

        // 5. Reopen + Verify persistence
        ReopenPptx();
        var p1 = _pptxHandler.Get("/slide[1]");
        var p2 = _pptxHandler.Get("/slide[2]");
        var p3 = _pptxHandler.Get("/slide[3]");
        p1.Format["background"].ToString().Should().Contain("FF0000");
        p2.Format["background"].ToString().Should().Contain("FFFF00");
        p3.Format["background"].ToString().Should().Contain("0000FF");
    }

    // =================================================================
    // EDGE CASE: PPTX shape with all text formatting properties.
    // =================================================================

    [Fact]
    public void Edge_Pptx_Shape_AllTextFormatting()
    {
        // 1. Add with all formatting
        _pptxHandler.Add("/", "slide", null, new());
        _pptxHandler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Fully formatted",
            ["font"] = "Arial",
            ["size"] = "24",
            ["bold"] = "true",
            ["italic"] = "true",
            ["color"] = "FF0000",
            ["underline"] = "single",
            ["align"] = "center",
            ["valign"] = "center"
        });

        // 2. Get + Verify initial state
        var node = _pptxHandler.Get("/slide[1]/shape[1]");
        node.Text.Should().Be("Fully formatted");
        node.Format["font"].ToString().Should().Be("Arial");
        node.Format["size"].ToString().Should().Be("24pt");
        node.Format["bold"].Should().Be(true);
        node.Format["italic"].Should().Be(true);
        node.Format["color"].ToString().Should().Be("FF0000");
        node.Format["underline"].ToString().Should().Be("single");
        node.Format["align"].ToString().Should().Be("center");
        node.Format["valign"].ToString().Should().Be("center");

        // 3. Set (modify — change some properties)
        _pptxHandler.Set("/slide[1]/shape[1]", new()
        {
            ["font"] = "Calibri",
            ["size"] = "18",
            ["bold"] = "false"
        });

        // 4. Get + Verify modification
        var modified = _pptxHandler.Get("/slide[1]/shape[1]");
        modified.Format["font"].ToString().Should().Be("Calibri");
        modified.Format["size"].ToString().Should().Be("18pt");

        // 5. Reopen + Verify persistence
        ReopenPptx();
        var persisted = _pptxHandler.Get("/slide[1]/shape[1]");
        persisted.Text.Should().Be("Fully formatted");
        persisted.Format["font"].ToString().Should().Be("Calibri");
        persisted.Format["size"].ToString().Should().Be("18pt");
        persisted.Format["italic"].Should().Be(true);
        persisted.Format["color"].ToString().Should().Be("FF0000");
        persisted.Format["underline"].ToString().Should().Be("single");
        persisted.Format["align"].ToString().Should().Be("center");
        persisted.Format["valign"].ToString().Should().Be("center");
    }

    // =================================================================
    // REGRESSION: 8-char RRGGBBAA hex color must be split into 6-char RGB + alpha.
    // srgbClr val only accepts 6-char hex; 8-char causes Office crash.
    // =================================================================

    [Fact]
    public void Pptx_ShapeFill_8CharHex_SplitsAlpha()
    {
        // Add slide + shape with 8-char AARRGGBB fill (POI convention: alpha first)
        // 88333333 = alpha 0x88 (≈53%), color 333333
        _pptxHandler.Add("/", "slide", null, new() { ["title"] = "Test" });
        _pptxHandler.Add("/slide[1]", "shape", null, new()
            { ["text"] = "Semi-transparent", ["fill"] = "88333333" });

        var node = _pptxHandler.Get("/slide[1]/shape[2]");
        node.Format["fill"].ToString().Should().Be("333333");
        node.Format.Should().ContainKey("opacity");

        // Verify persistence
        ReopenPptx();
        var persisted = _pptxHandler.Get("/slide[1]/shape[2]");
        persisted.Format["fill"].ToString().Should().Be("333333");
        persisted.Format.Should().ContainKey("opacity");
    }

    [Fact]
    public void Pptx_ShapeFill_8CharHex_FullyOpaque_NoAlphaElement()
    {
        // FF prefix means fully opaque — should behave same as 6-char hex
        _pptxHandler.Add("/", "slide", null, new() { ["title"] = "Test" });
        _pptxHandler.Add("/slide[1]", "shape", null, new()
            { ["text"] = "Opaque", ["fill"] = "FFFF0000" });

        var node = _pptxHandler.Get("/slide[1]/shape[2]");
        node.Format["fill"].ToString().Should().Be("FF0000");
        node.Format.ContainsKey("opacity").Should().BeFalse();
    }

    [Fact]
    public void Pptx_ShapeFill_6CharHex_NoAlpha()
    {
        // Normal 6-char hex should work as before, no opacity key
        _pptxHandler.Add("/", "slide", null, new() { ["title"] = "Test" });
        _pptxHandler.Add("/slide[1]", "shape", null, new()
            { ["text"] = "Opaque", ["fill"] = "FF0000" });

        var node = _pptxHandler.Get("/slide[1]/shape[2]");
        node.Format["fill"].ToString().Should().Be("FF0000");
        node.Format.ContainsKey("opacity").Should().BeFalse();
    }

    [Fact]
    public void Pptx_Background_SemicolonGradient_ParsedCorrectly()
    {
        // "LINEAR;C1;C2;angle" format must be recognized as gradient, not solid fill
        _pptxHandler.Add("/", "slide", null, new() { ["title"] = "Test" });
        _pptxHandler.Set("/slide[1]", new() { ["background"] = "LINEAR;0A0E29;1A2B5E;45" });

        var node = _pptxHandler.Get("/slide[1]");
        // Should read back as gradient: "0A0E29-1A2B5E-45"
        node.Format.Should().ContainKey("background");
        node.Format["background"].ToString().Should().Contain("0A0E29");
        node.Format["background"].ToString().Should().Contain("1A2B5E");

        // Verify persistence
        ReopenPptx();
        var persisted = _pptxHandler.Get("/slide[1]");
        persisted.Format["background"].ToString().Should().Contain("0A0E29");
    }

    [Fact]
    public void Pptx_Background_DashGradient_StillWorks()
    {
        // Canonical dash-separated format must still work
        _pptxHandler.Add("/", "slide", null, new() { ["title"] = "Test" });
        _pptxHandler.Set("/slide[1]", new() { ["background"] = "FF0000-0000FF-90" });

        var node = _pptxHandler.Get("/slide[1]");
        node.Format["background"].ToString().Should().Be("FF0000-0000FF-90");
    }

    [Fact]
    public void Pptx_TableCellGradient_8CharHex_Sanitized()
    {
        // Table cell gradient with 8-char hex colors
        _pptxHandler.Add("/", "slide", null, new() { ["title"] = "Test" });
        _pptxHandler.Add("/slide[1]", "table", null, new()
            { ["rows"] = "2", ["cols"] = "2" });
        _pptxHandler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
            { ["gradient"] = "88FF0000-CC0000FF-90" });

        // Should not crash on reopen
        ReopenPptx();
        _pptxHandler.Get("/slide[1]/table[1]/tr[1]/tc[1]").Should().NotBeNull();
    }

    [Fact]
    public void Word_Shading_8CharHex_Sanitized()
    {
        // Word table cell shading with 8-char hex
        _wordHandler.Add("/body", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });
        _wordHandler.Set("/body/tbl[1]/tr[1]/tc[1]", new()
            { ["shd"] = "88FF0000" });

        var node = _wordHandler.Get("/body/tbl[1]/tr[1]/tc[1]");
        // Should be sanitized to 6-char RGB (strip leading alpha bytes, AARRGGBB → RRGGBB)
        node.Format.Should().ContainKey("shd");
        node.Format["shd"].ToString().Should().Be("FF0000",
            "8-char AARRGGBB hex should extract 6-char RGB for OOXML");

        ReopenWord();
        _wordHandler.Get("/body/tbl[1]/tr[1]/tc[1]").Should().NotBeNull();
    }

    [Fact]
    public void Word_CellGradient_8CharHex_Sanitized()
    {
        // Word table cell gradient with 8-char hex colors
        _wordHandler.Add("/body", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });
        _wordHandler.Set("/body/tbl[1]/tr[1]/tc[1]", new()
            { ["shd"] = "gradient;88FF0000;CC0000FF;90" });

        // Should not crash on reopen
        ReopenWord();
        _wordHandler.Get("/body/tbl[1]/tr[1]/tc[1]").Should().NotBeNull();
    }
}
