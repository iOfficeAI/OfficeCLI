// Bug hunt Part 30 — PPTX handler bugs (full lifecycle tests):
// 1. Opacity not applied on SchemeColor fills → Set+Get+Reopen
// 2. LineOpacity not applied on SchemeColor lines → Set+Get+Reopen
// 3. Table cell strikethrough "double" mapping → Add+Get+Set+Get+Reopen
// 4. Table cell underline "true"/"double" mapping → Add+Get+Set+Get+Reopen
// 5. Table cell gradient fill readback format → Add+Get+Set+Get+Reopen
// 6. Table cell align all paragraphs → Add+Get+Set+Get+Reopen
// 7. Table cell border "none" handling → Add+Get+Set+Get+Set+Get+Reopen
// 8. Shape multiline text preserves paragraph props → Add+Get+Set+Get+Reopen
// 9. Shape preset geometry round-trip → Add+Get+Set+Get+Reopen
// 10. Shape rotation round-trip → Add+Get+Set+Get+Reopen
// 11. Shape text warp round-trip → Add+Get+Set+Get+Reopen
// 12. Multiple effects coexist → Add+Get+Set+Get+Set+Get+Reopen
// 13. Shape autofit round-trip → Add+Get+Set+Get+Reopen
// 14. Table cell font properties → Add+Get+Set+Get+Reopen
// 15. Shape name Set → Add+Get+Set+Get+Reopen
// 16. Paragraph-level alignment → Add+Get+Set+Get
// 17. Shape list style round-trip → Add+Get+Set+Get+Reopen
// 18. Slide background gradient persistence → Add+Get+Reopen+Get
// 19. Shape line dash round-trip → Add+Get+Set+Get+Reopen
// 20. Shape character spacing round-trip → Add+Get+Set+Get+Reopen
// 21. Shape baseline superscript/subscript → Add+Get+Set+Get+Set+Get+Reopen
// 22. Shape hyperlink round-trip → Add+Get+Set+Get+Set+Get

using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

public class BugHuntPart30 : IDisposable
{
    private readonly string _pptxPath;
    private PowerPointHandler _handler;

    public BugHuntPart30()
    {
        _pptxPath = Path.Combine(Path.GetTempPath(), $"bughunt30_{Guid.NewGuid():N}.pptx");
        BlankDocCreator.Create(_pptxPath);
        _handler = new PowerPointHandler(_pptxPath, editable: true);
    }

    public void Dispose()
    {
        _handler.Dispose();
        if (File.Exists(_pptxPath)) File.Delete(_pptxPath);
    }

    private PowerPointHandler Reopen()
    {
        _handler.Dispose();
        _handler = new PowerPointHandler(_pptxPath, editable: true);
        return _handler;
    }

    // ===================== Bug 1: Opacity on SchemeColor fill =====================

    [Fact]
    public void Bug_Pptx_Opacity_SchemeColor_FullLifecycle()
    {
        // 1. Create + Add shape with theme color fill
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Theme colored",
            ["fill"] = "accent1"
        });

        // 2. Get + Verify initial state (no opacity)
        var node1 = _handler.Get("/slide[1]/shape[1]");
        node1.Should().NotBeNull();
        node1.Text.Should().Be("Theme colored");
        node1.Format.Should().NotContainKey("opacity");

        // 3. Set opacity
        _handler.Set("/slide[1]/shape[1]", new() { ["opacity"] = "0.5" });

        // 4. Get + Verify opacity applied
        var node2 = _handler.Get("/slide[1]/shape[1]");
        node2.Format.Should().ContainKey("opacity",
            "opacity should be applied to shapes with scheme color fills");
        node2.Format["opacity"].ToString().Should().Be("0.5");

        // 5. Reopen + Verify persistence
        Reopen();
        var node3 = _handler.Get("/slide[1]/shape[1]");
        node3.Format.Should().ContainKey("opacity");
        node3.Format["opacity"].ToString().Should().Be("0.5");
    }

    [Fact]
    public void Bug_Pptx_Opacity_RgbColor_FullLifecycle()
    {
        // 1. Create + Add shape with RGB fill
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "RGB colored",
            ["fill"] = "FF0000"
        });

        // 2. Get + Verify initial state
        var node1 = _handler.Get("/slide[1]/shape[1]");
        node1.Format["fill"].ToString().Should().Be("FF0000");
        node1.Format.Should().NotContainKey("opacity");

        // 3. Set opacity
        _handler.Set("/slide[1]/shape[1]", new() { ["opacity"] = "0.5" });

        // 4. Get + Verify
        var node2 = _handler.Get("/slide[1]/shape[1]");
        node2.Format.Should().ContainKey("opacity");
        node2.Format["opacity"].ToString().Should().Be("0.5");

        // 5. Reopen + Verify
        Reopen();
        var node3 = _handler.Get("/slide[1]/shape[1]");
        node3.Format["opacity"].ToString().Should().Be("0.5");
    }

    // ===================== Bug 2: LineOpacity on SchemeColor line =====================

    [Fact]
    public void Bug_Pptx_LineOpacity_SchemeColor_FullLifecycle()
    {
        // 1. Create + Add
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Shape with theme line",
            ["line"] = "accent2"
        });

        // 2. Get + Verify initial state
        var node1 = _handler.Get("/slide[1]/shape[1]");
        node1.Format.Should().NotContainKey("lineOpacity");

        // 3. Set line opacity
        _handler.Set("/slide[1]/shape[1]", new() { ["lineopacity"] = "0.5" });

        // 4. Get + Verify
        var node2 = _handler.Get("/slide[1]/shape[1]");
        node2.Format.Should().ContainKey("lineOpacity",
            "lineOpacity should be applied to lines with scheme color");
        node2.Format["lineOpacity"].ToString().Should().Be("0.5");

        // 5. Reopen + Verify
        Reopen();
        var node3 = _handler.Get("/slide[1]/shape[1]");
        node3.Format.Should().ContainKey("lineOpacity");
        node3.Format["lineOpacity"].ToString().Should().Be("0.5");
    }

    [Fact]
    public void Bug_Pptx_LineOpacity_RgbColor_FullLifecycle()
    {
        // 1. Create + Add
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "RGB line",
            ["line"] = "0000FF"
        });

        // 2. Get + Verify
        var node1 = _handler.Get("/slide[1]/shape[1]");
        node1.Format.Should().ContainKey("line");

        // 3. Set + 4. Get + Verify
        _handler.Set("/slide[1]/shape[1]", new() { ["lineopacity"] = "0.5" });
        var node2 = _handler.Get("/slide[1]/shape[1]");
        node2.Format["lineOpacity"].ToString().Should().Be("0.5");

        // 5. Reopen + Verify
        Reopen();
        var node3 = _handler.Get("/slide[1]/shape[1]");
        node3.Format["lineOpacity"].ToString().Should().Be("0.5");
    }

    // ===================== Bug 3: Table cell strike "double" =====================

    [Fact]
    public void Bug_Pptx_TableCell_Strike_FullLifecycle()
    {
        // 1. Create + Add table
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });

        // 2. Add text to cell + Get + Verify
        _handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["text"] = "Strike me" });
        var node1 = _handler.Get("/slide[1]/table[1]/tr[1]/tc[1]", 0);
        node1.Text.Should().Be("Strike me");

        // 3. Set double strikethrough
        _handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["strike"] = "double" });

        // 4. Get + Verify via raw XML
        var raw = _handler.Raw("/slide[1]");
        raw.Should().Contain("dblStrike",
            "table cell strike='double' should produce dblStrike in XML");

        // 5. Set single strikethrough
        _handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["strike"] = "true" });

        // 6. Get + Verify
        raw = _handler.Raw("/slide[1]");
        raw.Should().Contain("sngStrike",
            "table cell strike='true' should produce sngStrike");

        // 7. Reopen + Verify
        Reopen();
        raw = _handler.Raw("/slide[1]");
        raw.Should().Contain("sngStrike");
    }

    // ===================== Bug 4: Table cell underline mapping =====================

    [Fact]
    public void Bug_Pptx_TableCell_Underline_FullLifecycle()
    {
        // 1. Create + Add
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });
        _handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["text"] = "Underline me" });

        // 2. Get + Verify no underline
        var node1 = _handler.Get("/slide[1]/table[1]/tr[1]/tc[1]", 0);
        node1.Text.Should().Be("Underline me");

        // 3. Set underline "true" → should map to sng
        _handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["underline"] = "true" });
        var raw1 = _handler.Raw("/slide[1]");
        raw1.Should().Contain("u=\"sng\"",
            "underline='true' should map to u='sng' (single underline)");

        // 4. Set underline "double" → should map to dbl
        _handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["underline"] = "double" });
        var raw2 = _handler.Raw("/slide[1]");
        raw2.Should().Contain("u=\"dbl\"",
            "underline='double' should map to u='dbl' (double underline)");

        // 5. Reopen + Verify
        Reopen();
        var raw3 = _handler.Raw("/slide[1]");
        raw3.Should().Contain("u=\"dbl\"");
    }

    // ===================== Bug 5: Table cell gradient fill readback =====================

    [Fact]
    public void Bug_Pptx_TableCell_GradientFill_FullLifecycle()
    {
        // 1. Create + Add
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });
        _handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["text"] = "Gradient cell" });

        // 2. Get + Verify initial (no fill)
        var node1 = _handler.Get("/slide[1]/table[1]/tr[1]/tc[1]", 0);
        node1.Text.Should().Be("Gradient cell");

        // 3. Set gradient fill
        _handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["fill"] = "FF0000-0000FF-90" });

        // 4. Get + Verify — readback format should match Set format
        var node2 = _handler.Get("/slide[1]/table[1]/tr[1]/tc[1]", 0);
        var fill = node2.Format["fill"].ToString()!;
        fill.Should().NotStartWith("gradient;",
            "readback should use hyphen format matching Set input");
        fill.Should().Contain("FF0000");
        fill.Should().Contain("0000FF");

        // 5. Set new gradient (round-trip from readback)
        _handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["fill"] = "00FF00-FF00FF" });

        // 6. Get + Verify
        var node3 = _handler.Get("/slide[1]/table[1]/tr[1]/tc[1]", 0);
        node3.Format["fill"].ToString().Should().Contain("00FF00");
        node3.Format["fill"].ToString().Should().Contain("FF00FF");

        // 7. Reopen + Verify
        Reopen();
        var node4 = _handler.Get("/slide[1]/table[1]/tr[1]/tc[1]", 0);
        node4.Format["fill"].ToString().Should().Contain("00FF00");
    }

    // ===================== Bug 6: Table cell align all paragraphs =====================

    [Fact]
    public void Bug_Pptx_TableCell_Align_FullLifecycle()
    {
        // 1. Create + Add
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });

        // 2. Set multi-line text
        _handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
        {
            ["text"] = "Line 1\\nLine 2\\nLine 3"
        });

        // 3. Get + Verify text
        var node1 = _handler.Get("/slide[1]/table[1]/tr[1]/tc[1]", 0);
        node1.Text.Should().Contain("Line 1");

        // 4. Set alignment
        _handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["align"] = "center" });

        // 5. Get + Verify all paragraphs aligned
        var raw = _handler.Raw("/slide[1]");
        var centerCount = System.Text.RegularExpressions.Regex.Matches(raw, @"algn=""ctr""").Count;
        centerCount.Should().BeGreaterThanOrEqualTo(3,
            "all paragraphs in a multi-line table cell should get alignment");

        // 6. Reopen + Verify
        Reopen();
        raw = _handler.Raw("/slide[1]");
        System.Text.RegularExpressions.Regex.Matches(raw, @"algn=""ctr""").Count.Should().BeGreaterThanOrEqualTo(3);
    }

    // ===================== Bug 7: Table cell border "none" =====================

    [Fact]
    public void Bug_Pptx_TableCell_Border_None_FullLifecycle()
    {
        // 1. Create + Add
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });

        // 2. Set border
        _handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
        {
            ["border.left"] = "2pt solid FF0000"
        });

        // 3. Get + Verify border set
        var node1 = _handler.Get("/slide[1]/table[1]/tr[1]/tc[1]", 0);
        node1.Format.Should().ContainKey("border.left");

        // 4. Set border to "none"
        _handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["border.left"] = "none" });

        // 5. Get + Verify border removed
        var node2 = _handler.Get("/slide[1]/table[1]/tr[1]/tc[1]", 0);
        node2.Format.Should().NotContainKey("border.left",
            "border.left should be removed after setting 'none'");

        // 6. Reopen + Verify
        Reopen();
        var node3 = _handler.Get("/slide[1]/table[1]/tr[1]/tc[1]", 0);
        node3.Format.Should().NotContainKey("border.left");
    }

    // ===================== Bug 8: Multiline text preserves paragraph props =====================

    [Fact]
    public void Bug_Pptx_SetText_Multiline_PreservesParagraphProps()
    {
        // 1. Create + Add shape with alignment
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Initial text",
            ["align"] = "center"
        });

        // 2. Get + Verify initial state
        var node1 = _handler.Get("/slide[1]/shape[1]");
        node1.Text.Should().Be("Initial text");
        node1.Format["align"].ToString().Should().Be("center");

        // 3. Set multiline text (triggers multi-paragraph replace)
        _handler.Set("/slide[1]/shape[1]", new() { ["text"] = "Line 1\\nLine 2" });

        // 4. Get + Verify alignment preserved
        var node2 = _handler.Get("/slide[1]/shape[1]");
        node2.Text.Should().Be("Line 1\nLine 2");
        node2.Format.Should().ContainKey("align",
            "paragraph alignment should be preserved when replacing text with multiline");
        node2.Format["align"].ToString().Should().Be("center");

        // 5. Reopen + Verify
        Reopen();
        var node3 = _handler.Get("/slide[1]/shape[1]");
        node3.Text.Should().Be("Line 1\nLine 2");
        node3.Format["align"].ToString().Should().Be("center");
    }

    // ===================== 9: Shape preset geometry round-trip =====================

    [Fact]
    public void Edge_Pptx_Shape_PresetGeometry_FullLifecycle()
    {
        // 1. Create + Add
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Ellipse shape",
            ["preset"] = "ellipse"
        });

        // 2. Get + Verify
        var node1 = _handler.Get("/slide[1]/shape[1]");
        node1.Format["preset"].ToString().Should().Be("ellipse");

        // 3. Set new preset
        _handler.Set("/slide[1]/shape[1]", new() { ["preset"] = "triangle" });

        // 4. Get + Verify
        var node2 = _handler.Get("/slide[1]/shape[1]");
        node2.Format["preset"].ToString().Should().Be("triangle");

        // 5. Reopen + Verify
        Reopen();
        var node3 = _handler.Get("/slide[1]/shape[1]");
        node3.Format["preset"].ToString().Should().Be("triangle");
    }

    // ===================== 10: Shape rotation round-trip =====================

    [Fact]
    public void Edge_Pptx_Shape_Rotation_FullLifecycle()
    {
        // 1. Create + Add
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Rotated",
            ["rotation"] = "45"
        });

        // 2. Get + Verify
        var node1 = _handler.Get("/slide[1]/shape[1]");
        node1.Format["rotation"].ToString().Should().Be("45");

        // 3. Set new rotation
        _handler.Set("/slide[1]/shape[1]", new() { ["rotation"] = "90" });

        // 4. Get + Verify
        var node2 = _handler.Get("/slide[1]/shape[1]");
        node2.Format["rotation"].ToString().Should().Be("90");

        // 5. Reopen + Verify
        Reopen();
        var node3 = _handler.Get("/slide[1]/shape[1]");
        node3.Format["rotation"].ToString().Should().Be("90");
    }

    // ===================== 11: Shape text warp round-trip =====================

    [Fact]
    public void Edge_Pptx_Shape_TextWarp_FullLifecycle()
    {
        // 1. Create + Add
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Warped text" });

        // 2. Get + Verify (no warp initially)
        var node1 = _handler.Get("/slide[1]/shape[1]");
        node1.Format.Should().NotContainKey("textWarp");

        // 3. Set text warp
        _handler.Set("/slide[1]/shape[1]", new() { ["textwarp"] = "textWave1" });

        // 4. Get + Verify
        var node2 = _handler.Get("/slide[1]/shape[1]");
        node2.Format["textWarp"].ToString().Should().Be("textWave1");

        // 5. Set to none
        _handler.Set("/slide[1]/shape[1]", new() { ["textwarp"] = "none" });

        // 6. Get + Verify removed
        var node3 = _handler.Get("/slide[1]/shape[1]");
        node3.Format.Should().NotContainKey("textWarp");

        // 7. Reopen + Verify still gone
        Reopen();
        var node4 = _handler.Get("/slide[1]/shape[1]");
        node4.Format.Should().NotContainKey("textWarp");
    }

    // ===================== 12: Multiple effects coexist =====================

    [Fact]
    public void Edge_Pptx_Shape_MultipleEffects_FullLifecycle()
    {
        // 1. Create + Add
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Effects shape",
            ["fill"] = "4472C4"
        });

        // 2. Get + Verify no effects
        var node1 = _handler.Get("/slide[1]/shape[1]");
        node1.Format.Should().NotContainKey("shadow");

        // 3. Set shadow + glow
        _handler.Set("/slide[1]/shape[1]", new()
        {
            ["shadow"] = "000000-4-315-3-50",
            ["glow"] = "FF0000-8-60"
        });

        // 4. Get + Verify both present
        var node2 = _handler.Get("/slide[1]/shape[1]");
        node2.Format.Should().ContainKey("shadow");
        node2.Format.Should().ContainKey("glow");

        // 5. Set reflection — should NOT remove shadow/glow
        _handler.Set("/slide[1]/shape[1]", new() { ["reflection"] = "half" });

        // 6. Get + Verify all three coexist
        var node3 = _handler.Get("/slide[1]/shape[1]");
        node3.Format.Should().ContainKey("shadow");
        node3.Format.Should().ContainKey("glow");
        node3.Format.Should().ContainKey("reflection");

        // 7. Reopen + Verify
        Reopen();
        var node4 = _handler.Get("/slide[1]/shape[1]");
        node4.Format.Should().ContainKey("shadow");
        node4.Format.Should().ContainKey("glow");
        node4.Format.Should().ContainKey("reflection");
    }

    // ===================== 13: Shape autofit round-trip =====================

    [Fact]
    public void Edge_Pptx_Shape_AutoFit_FullLifecycle()
    {
        // 1. Create + Add with autofit
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Auto fit",
            ["autofit"] = "normal"
        });

        // 2. Get + Verify
        var node1 = _handler.Get("/slide[1]/shape[1]");
        node1.Format["autoFit"].ToString().Should().Be("normal");

        // 3. Set to shape
        _handler.Set("/slide[1]/shape[1]", new() { ["autofit"] = "shape" });

        // 4. Get + Verify
        var node2 = _handler.Get("/slide[1]/shape[1]");
        node2.Format["autoFit"].ToString().Should().Be("shape");

        // 5. Set to none
        _handler.Set("/slide[1]/shape[1]", new() { ["autofit"] = "none" });

        // 6. Get + Verify
        var node3 = _handler.Get("/slide[1]/shape[1]");
        node3.Format["autoFit"].ToString().Should().Be("none");

        // 7. Reopen + Verify
        Reopen();
        var node4 = _handler.Get("/slide[1]/shape[1]");
        node4.Format["autoFit"].ToString().Should().Be("none");
    }

    // ===================== 14: Table cell font properties =====================

    [Fact]
    public void Edge_Pptx_TableCell_FontProps_FullLifecycle()
    {
        // 1. Create + Add
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });

        // 2. Set text + formatting
        _handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
        {
            ["text"] = "Formatted cell",
            ["bold"] = "true",
            ["italic"] = "true",
            ["font"] = "Arial",
            ["size"] = "14"
        });

        // 3. Get + Verify
        var raw1 = _handler.Raw("/slide[1]");
        raw1.Should().Contain("b=\"1\"");
        raw1.Should().Contain("i=\"1\"");
        raw1.Should().Contain("Arial");

        // 4. Reopen + Verify persistence
        Reopen();
        var raw2 = _handler.Raw("/slide[1]");
        raw2.Should().Contain("b=\"1\"");
        raw2.Should().Contain("Arial");
    }

    // ===================== 15: Shape name Set =====================

    [Fact]
    public void Edge_Pptx_Shape_Name_FullLifecycle()
    {
        // 1. Create + Add
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Named shape",
            ["name"] = "MyShape"
        });

        // 2. Get + Verify
        var node1 = _handler.Get("/slide[1]/shape[1]");
        node1.Format["name"].ToString().Should().Be("MyShape");

        // 3. Set new name
        _handler.Set("/slide[1]/shape[1]", new() { ["name"] = "RenamedShape" });

        // 4. Get + Verify
        var node2 = _handler.Get("/slide[1]/shape[1]");
        node2.Format["name"].ToString().Should().Be("RenamedShape");

        // 5. Reopen + Verify
        Reopen();
        var node3 = _handler.Get("/slide[1]/shape[1]");
        node3.Format["name"].ToString().Should().Be("RenamedShape");
    }

    // ===================== 16: Paragraph-level alignment =====================

    [Fact]
    public void Edge_Pptx_ParagraphLevel_Alignment_FullLifecycle()
    {
        // 1. Create + Add multiline shape
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Line 1\\nLine 2\\nLine 3"
        });

        // 2. Get + Verify initial (3 paragraphs)
        var node1 = _handler.Get("/slide[1]/shape[1]", 2);
        node1.Children.Should().HaveCount(3);

        // 3. Set alignment only on paragraph 2
        _handler.Set("/slide[1]/shape[1]/paragraph[2]", new() { ["align"] = "right" });

        // 4. Get + Verify only para 2 has right
        var node2 = _handler.Get("/slide[1]/shape[1]", 2);
        var para2 = node2.Children[1];
        para2.Format.Should().ContainKey("align");
        para2.Format["align"].ToString().Should().Be("r");

        // Para 1 should NOT be "r"
        var para1 = node2.Children[0];
        if (para1.Format.ContainsKey("align"))
            para1.Format["align"].ToString().Should().NotBe("r");
    }

    // ===================== 17: Shape list style =====================

    [Fact]
    public void Edge_Pptx_Shape_ListStyle_FullLifecycle()
    {
        // 1. Create + Add with bullet list
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Item 1\\nItem 2\\nItem 3",
            ["list"] = "bullet"
        });

        // 2. Get + Verify
        var node1 = _handler.Get("/slide[1]/shape[1]");
        node1.Format.Should().ContainKey("list");

        // 3. Set to numbered
        _handler.Set("/slide[1]/shape[1]", new() { ["list"] = "numbered" });

        // 4. Get + Verify
        var node2 = _handler.Get("/slide[1]/shape[1]");
        node2.Format.Should().ContainKey("list");

        // 5. Set to none
        _handler.Set("/slide[1]/shape[1]", new() { ["list"] = "none" });

        // 6. Get + Verify
        var node3 = _handler.Get("/slide[1]/shape[1]");
        node3.Format.Should().NotContainKey("list");

        // 7. Reopen + Verify
        Reopen();
        var node4 = _handler.Get("/slide[1]/shape[1]");
        node4.Format.Should().NotContainKey("list");
    }

    // ===================== 18: Slide background gradient persistence =====================

    [Fact]
    public void Edge_Pptx_SlideBackground_FullLifecycle()
    {
        // 1. Create + Add slide with gradient bg
        _handler.Add("/", "slide", null, new() { ["background"] = "FF0000-0000FF-90" });

        // 2. Get + Verify
        var node1 = _handler.Get("/slide[1]");
        node1.Format.Should().ContainKey("background");
        node1.Format["background"].ToString().Should().Contain("FF0000");

        // 3. Reopen + Verify persistence
        Reopen();
        var node2 = _handler.Get("/slide[1]");
        node2.Format.Should().ContainKey("background");
        node2.Format["background"].ToString().Should().Contain("FF0000");
        node2.Format["background"].ToString().Should().Contain("0000FF");
    }

    // ===================== 19: Shape line dash =====================

    [Fact]
    public void Edge_Pptx_Shape_LineDash_FullLifecycle()
    {
        // 1. Create + Add
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Dashed border",
            ["line"] = "000000",
            ["linewidth"] = "2pt"
        });

        // 2. Get + Verify line set
        var node1 = _handler.Get("/slide[1]/shape[1]");
        node1.Format.Should().ContainKey("line");

        // 3. Set dash
        _handler.Set("/slide[1]/shape[1]", new() { ["linedash"] = "dash" });

        // 4. Get + Verify
        var node2 = _handler.Get("/slide[1]/shape[1]");
        node2.Format["lineDash"].ToString().Should().Be("dash");

        // 5. Set to dot
        _handler.Set("/slide[1]/shape[1]", new() { ["linedash"] = "dot" });

        // 6. Get + Verify
        var node3 = _handler.Get("/slide[1]/shape[1]");
        node3.Format["lineDash"].ToString().Should().Be("dot");

        // 7. Reopen + Verify
        Reopen();
        var node4 = _handler.Get("/slide[1]/shape[1]");
        node4.Format["lineDash"].ToString().Should().Be("dot");
    }

    // ===================== 20: Character spacing =====================

    [Fact]
    public void Edge_Pptx_Shape_Spacing_FullLifecycle()
    {
        // 1. Create + Add
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Spaced text" });

        // 2. Get + Verify no spacing
        var node1 = _handler.Get("/slide[1]/shape[1]");
        node1.Format.Should().NotContainKey("spacing");

        // 3. Set spacing
        _handler.Set("/slide[1]/shape[1]", new() { ["spacing"] = "3" });

        // 4. Get + Verify
        var node2 = _handler.Get("/slide[1]/shape[1]");
        node2.Format["spacing"].ToString().Should().Be("3");

        // 5. Reopen + Verify
        Reopen();
        var node3 = _handler.Get("/slide[1]/shape[1]");
        node3.Format["spacing"].ToString().Should().Be("3");
    }

    // ===================== 21: Baseline superscript/subscript =====================

    [Fact]
    public void Edge_Pptx_Shape_Baseline_FullLifecycle()
    {
        // 1. Create + Add
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Normal text" });

        // 2. Get + Verify no baseline
        var node1 = _handler.Get("/slide[1]/shape[1]");
        node1.Format.Should().NotContainKey("baseline");

        // 3. Set superscript
        _handler.Set("/slide[1]/shape[1]", new() { ["superscript"] = "true" });

        // 4. Get + Verify positive baseline
        var node2 = _handler.Get("/slide[1]/shape[1]");
        node2.Format.Should().ContainKey("baseline");
        double.Parse(node2.Format["baseline"].ToString()!).Should().BeGreaterThan(0);

        // 5. Set subscript
        _handler.Set("/slide[1]/shape[1]", new() { ["subscript"] = "true" });

        // 6. Get + Verify negative baseline
        var node3 = _handler.Get("/slide[1]/shape[1]");
        double.Parse(node3.Format["baseline"].ToString()!).Should().BeLessThan(0);

        // 7. Reopen + Verify
        Reopen();
        var node4 = _handler.Get("/slide[1]/shape[1]");
        double.Parse(node4.Format["baseline"].ToString()!).Should().BeLessThan(0);
    }

    // ===================== 22: Shape hyperlink =====================

    [Fact]
    public void Edge_Pptx_Shape_Hyperlink_FullLifecycle()
    {
        // 1. Create + Add with link
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Click here",
            ["link"] = "https://example.com"
        });

        // 2. Get + Verify
        var node1 = _handler.Get("/slide[1]/shape[1]");
        node1.Format.Should().ContainKey("link");
        node1.Format["link"].ToString().Should().Contain("example.com");

        // 3. Remove hyperlink
        _handler.Set("/slide[1]/shape[1]", new() { ["link"] = "none" });

        // 4. Get + Verify removed
        var node2 = _handler.Get("/slide[1]/shape[1]");
        node2.Format.Should().NotContainKey("link");

        // 5. Re-add a new link
        _handler.Set("/slide[1]/shape[1]", new() { ["link"] = "https://test.org" });

        // 6. Get + Verify
        var node3 = _handler.Get("/slide[1]/shape[1]");
        node3.Format.Should().ContainKey("link");
        node3.Format["link"].ToString().Should().Contain("test.org");
    }
}
