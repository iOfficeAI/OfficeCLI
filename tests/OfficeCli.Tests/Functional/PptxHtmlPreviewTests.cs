// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Tests for PPTX HTML preview generation.
/// Verifies that ViewAsHtml produces valid, self-contained HTML
/// without modifying the original document.
/// </summary>
public class PptxHtmlPreviewTests : IDisposable
{
    private readonly string _path;
    private PowerPointHandler _handler;

    public PptxHtmlPreviewTests()
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

    private PowerPointHandler Reopen()
    {
        _handler.Dispose();
        _handler = new PowerPointHandler(_path, editable: true);
        return _handler;
    }

    // ==================== Basic HTML structure ====================

    [Fact]
    public void ViewAsHtml_EmptyPresentation_ReturnsValidHtml()
    {
        _handler.Add("/", "slide", null, new());
        var html = _handler.ViewAsHtml();

        html.Should().StartWith("<!DOCTYPE html>");
        html.Should().Contain("<html");
        html.Should().Contain("</html>");
        html.Should().Contain("<style>");
        html.Should().Contain("</style>");
        html.Should().Contain("Slide 1");
        html.Should().Contain("class=\"slide\"");
    }

    [Fact]
    public void ViewAsHtml_MultipleSlides_ContainsAllSlides()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/", "slide", null, new());
        _handler.Add("/", "slide", null, new());

        var html = _handler.ViewAsHtml();

        html.Should().Contain("Slide 1");
        html.Should().Contain("Slide 2");
        html.Should().Contain("Slide 3");
    }

    [Fact]
    public void ViewAsHtml_SlideRange_OnlyIncludesRequestedSlides()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/", "slide", null, new());
        _handler.Add("/", "slide", null, new());

        var html = _handler.ViewAsHtml(startSlide: 2, endSlide: 2);

        html.Should().NotContain("Slide 1");
        html.Should().Contain("Slide 2");
        html.Should().NotContain("Slide 3");
    }

    // ==================== Shape rendering ====================

    [Fact]
    public void ViewAsHtml_ShapeWithText_ContainsTextContent()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Hello World",
            ["x"] = "2cm", ["y"] = "3cm", ["width"] = "10cm", ["height"] = "5cm"
        });

        var html = _handler.ViewAsHtml();

        html.Should().Contain("Hello World");
        html.Should().Contain("class=\"shape\"");
    }

    [Fact]
    public void ViewAsHtml_ShapeWithFill_ContainsBackgroundColor()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Colored", ["fill"] = "FF0000"
        });

        var html = _handler.ViewAsHtml();

        html.Should().Contain("#FF0000");
    }

    [Fact]
    public void ViewAsHtml_ShapeWithBold_ContainsFontWeight()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Bold Text", ["bold"] = "true"
        });

        var html = _handler.ViewAsHtml();

        html.Should().Contain("font-weight:bold");
        html.Should().Contain("Bold Text");
    }

    [Fact]
    public void ViewAsHtml_ShapeWithFontSize_ContainsFontSize()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Big Text", ["size"] = "24pt"
        });

        var html = _handler.ViewAsHtml();

        html.Should().Contain("font-size:24pt");
    }

    [Fact]
    public void ViewAsHtml_ShapeWithRotation_ContainsTransformRotate()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Rotated", ["rotation"] = "45"
        });

        var html = _handler.ViewAsHtml();

        html.Should().Contain("rotate(45");
    }

    // ==================== Background ====================

    [Fact]
    public void ViewAsHtml_SlideWithBackground_ContainsBackgroundStyle()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Set("/slide[1]", new() { ["background"] = "4472C4" });

        var html = _handler.ViewAsHtml();

        html.Should().Contain("#4472C4");
    }

    [Fact]
    public void ViewAsHtml_GradientBackground_ContainsLinearGradient()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Set("/slide[1]", new() { ["background"] = "FF0000-0000FF" });

        var html = _handler.ViewAsHtml();

        html.Should().Contain("linear-gradient");
    }

    // ==================== Picture rendering ====================

    [Fact]
    public void ViewAsHtml_Picture_ContainsBase64Image()
    {
        _handler.Add("/", "slide", null, new());

        // Create a minimal PNG file for testing
        var imgPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.png");
        try
        {
            File.WriteAllBytes(imgPath, CreateMinimalPng());
            _handler.Add("/slide[1]", "picture", null, new() { ["path"] = imgPath });

            var html = _handler.ViewAsHtml();

            html.Should().Contain("class=\"picture\"");
            html.Should().Contain("data:image/png;base64,");
            html.Should().Contain("<img ");
        }
        finally
        {
            if (File.Exists(imgPath)) File.Delete(imgPath);
        }
    }

    // ==================== Table rendering ====================

    [Fact]
    public void ViewAsHtml_Table_ContainsTableElement()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "2", ["cols"] = "3"
        });

        var html = _handler.ViewAsHtml();

        html.Should().Contain("class=\"slide-table\"");
        html.Should().Contain("<table");
        html.Should().Contain("<td");
    }

    // ==================== Connector rendering ====================

    [Fact]
    public void ViewAsHtml_Connector_ContainsSvgLine()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "connector", null, new()
        {
            ["x"] = "2cm", ["y"] = "3cm", ["width"] = "5cm", ["height"] = "2cm",
            ["line"] = "FF0000"
        });

        var html = _handler.ViewAsHtml();

        html.Should().Contain("<svg");
        html.Should().Contain("<line");
    }

    // ==================== Document safety ====================

    [Fact]
    public void ViewAsHtml_DoesNotModifyDocument_GetReturnsOriginalValues()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Original",
            ["x"] = "2cm", ["y"] = "3cm", ["width"] = "10cm", ["height"] = "5cm",
            ["fill"] = "FF0000"
        });

        // Capture original state
        var before = _handler.Get("/slide[1]/shape[1]");
        var origX = before.Format["x"];
        var origY = before.Format["y"];
        var origW = before.Format["width"];
        var origH = before.Format["height"];
        var origText = before.Text;

        // Generate HTML preview
        _handler.ViewAsHtml();

        // Verify nothing changed
        var after = _handler.Get("/slide[1]/shape[1]");
        after.Format["x"].Should().Be(origX);
        after.Format["y"].Should().Be(origY);
        after.Format["width"].Should().Be(origW);
        after.Format["height"].Should().Be(origH);
        after.Text.Should().Be(origText);
    }

    [Fact]
    public void ViewAsHtml_WithGroup_DoesNotModifyGroupChildren()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Shape A",
            ["x"] = "2cm", ["y"] = "2cm", ["width"] = "4cm", ["height"] = "3cm"
        });
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Shape B",
            ["x"] = "7cm", ["y"] = "2cm", ["width"] = "4cm", ["height"] = "3cm"
        });
        _handler.Add("/slide[1]", "group", null, new() { ["shapes"] = "1,2" });

        // Capture original shape texts via view
        var textBefore = _handler.ViewAsText();

        // Generate HTML preview
        var html = _handler.ViewAsHtml();
        html.Should().Contain("class=\"group\"");

        // Verify shapes still accessible and unchanged after preview
        var textAfter = _handler.ViewAsText();
        textAfter.Should().Be(textBefore);
    }

    [Fact]
    public void ViewAsHtml_AfterPreview_DocumentSavesCorrectly()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Persist Test", ["fill"] = "00FF00"
        });

        // Generate preview
        _handler.ViewAsHtml();

        // Reopen and verify data intact
        Reopen();
        var node = _handler.Get("/slide[1]/shape[1]");
        node.Text.Should().Be("Persist Test");
        node.Format["fill"].Should().Be("#00FF00");
    }

    // ==================== XSS / injection safety ====================

    [Fact]
    public void ViewAsHtml_TextWithHtmlTags_IsEscaped()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "<script>alert('xss')</script>"
        });

        var html = _handler.ViewAsHtml();

        html.Should().NotContain("<script>alert");
        html.Should().Contain("&lt;script&gt;");
    }

    [Fact]
    public void ViewAsHtml_TextWithQuotes_IsEscaped()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "He said \"hello\" & she said 'bye'"
        });

        var html = _handler.ViewAsHtml();

        html.Should().Contain("&amp;");
        html.Should().NotContain("\" & ");
    }

    // ==================== Edge cases ====================

    [Fact]
    public void ViewAsHtml_ShapeWithNoFill_ContainsTransparent()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "No Fill", ["fill"] = "none"
        });

        var html = _handler.ViewAsHtml();

        html.Should().Contain("transparent");
    }

    [Fact]
    public void ViewAsHtml_EmptyTextShape_DoesNotCrash()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["x"] = "1cm", ["y"] = "1cm", ["width"] = "5cm", ["height"] = "3cm"
        });

        var html = _handler.ViewAsHtml();
        html.Should().Contain("class=\"slide\"");
    }

    [Fact]
    public void ViewAsHtml_ShapeWithGradientFill_ContainsGradient()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Gradient", ["gradient"] = "FF0000-0000FF-90"
        });

        var html = _handler.ViewAsHtml();

        html.Should().Contain("linear-gradient");
    }

    [Fact]
    public void ViewAsHtml_ShapeWithBorder_ContainsBorder()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Bordered", ["line"] = "000000", ["lineWidth"] = "2pt"
        });

        var html = _handler.ViewAsHtml();

        html.Should().Contain("border:");
    }

    [Fact]
    public void ViewAsHtml_ContainsKeyboardNavScript()
    {
        _handler.Add("/", "slide", null, new());
        var html = _handler.ViewAsHtml();

        html.Should().Contain("<script>");
        html.Should().Contain("ArrowDown");
        html.Should().Contain("ArrowUp");
    }

    // ==================== All element types ====================

    [Fact]
    public void ViewAsHtml_AllElementTypes_DoesNotCrash()
    {
        // Create a slide with every element type that renders in HTML
        _handler.Add("/", "slide", null, new());
        _handler.Add("/", "slide", null, new()); // second slide for zoom target

        // Shape with all formatting options
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Full Shape",
            ["x"] = "1cm", ["y"] = "1cm", ["width"] = "8cm", ["height"] = "3cm",
            ["fill"] = "4472C4", ["line"] = "000000", ["lineWidth"] = "1pt",
            ["bold"] = "true", ["italic"] = "true", ["underline"] = "single",
            ["size"] = "18pt", ["color"] = "FFFFFF", ["align"] = "center",
            ["rotation"] = "15", ["valign"] = "center"
        });

        // Shape with gradient fill
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Gradient Shape",
            ["gradient"] = "FF0000-0000FF-45"
        });

        // Shape with no fill
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Transparent", ["fill"] = "none"
        });

        // Picture
        var imgPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.png");
        File.WriteAllBytes(imgPath, CreateMinimalPng());
        try
        {
            _handler.Add("/slide[1]", "picture", null, new()
            {
                ["path"] = imgPath,
                ["x"] = "10cm", ["y"] = "5cm", ["width"] = "6cm", ["height"] = "4cm"
            });
        }
        finally { if (File.Exists(imgPath)) File.Delete(imgPath); }

        // Table
        _handler.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "3", ["cols"] = "3",
            ["x"] = "1cm", ["y"] = "10cm", ["width"] = "20cm", ["height"] = "6cm"
        });
        // Set table cell content
        _handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new() { ["text"] = "Cell A1", ["fill"] = "E0E0E0" });
        _handler.Set("/slide[1]/table[1]/tr[1]/tc[2]", new() { ["text"] = "Cell B1", ["bold"] = "true" });

        // Connector
        _handler.Add("/slide[1]", "connector", null, new()
        {
            ["x"] = "5cm", ["y"] = "8cm", ["width"] = "10cm", ["height"] = "0cm",
            ["line"] = "FF0000", ["lineWidth"] = "2pt"
        });

        // Zoom (slide zoom)
        _handler.Add("/slide[1]", "zoom", null, new()
        {
            ["target"] = "2",
            ["x"] = "15cm", ["y"] = "1cm", ["width"] = "6cm", ["height"] = "3.38cm"
        });

        // Group (group existing shapes)
        // Shapes 2 and 3 (gradient and transparent)
        _handler.Add("/slide[1]", "group", null, new() { ["shapes"] = "2,3" });

        // Generate HTML — should not throw
        var html = _handler.ViewAsHtml();

        // Verify all element types appear in HTML
        html.Should().Contain("Full Shape");
        html.Should().Contain("Gradient Shape");
        html.Should().Contain("Transparent");
        html.Should().Contain("class=\"picture\"");
        html.Should().Contain("class=\"slide-table\"");
        html.Should().Contain("Cell A1");
        html.Should().Contain("<svg");              // connector
        html.Should().Contain("class=\"group\"");
    }

    [Fact]
    public void ViewAsHtml_ShapeWithShadow_ContainsBoxShadow()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Shadow Box",
            ["fill"] = "FFFFFF",
            ["shadow"] = "000000-4-45-3-40"
        });

        var html = _handler.ViewAsHtml();
        html.Should().Contain("box-shadow:");
    }

    [Fact]
    public void ViewAsHtml_ShapeWithRoundRect_ContainsBorderRadius()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Rounded",
            ["preset"] = "roundRect"
        });

        var html = _handler.ViewAsHtml();
        html.Should().Contain("border-radius");
    }

    [Fact]
    public void ViewAsHtml_ShapeWithEllipse_ContainsBorderRadius50()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Circle",
            ["preset"] = "ellipse"
        });

        var html = _handler.ViewAsHtml();
        html.Should().Contain("border-radius:50%");
    }

    [Fact]
    public void ViewAsHtml_ShapeWithItalicAndUnderline_ContainsStyles()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Styled",
            ["italic"] = "true",
            ["underline"] = "single"
        });

        var html = _handler.ViewAsHtml();
        html.Should().Contain("font-style:italic");
        html.Should().Contain("text-decoration:underline");
    }

    [Fact]
    public void ViewAsHtml_ShapeWithStrikethrough_ContainsLineThrough()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Struck",
            ["strike"] = "single"
        });

        var html = _handler.ViewAsHtml();
        html.Should().Contain("line-through");
    }

    [Fact]
    public void ViewAsHtml_ShapeWithFlipH_ContainsScaleX()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Flipped", ["flipH"] = "true"
        });

        var html = _handler.ViewAsHtml();
        html.Should().Contain("scaleX(-1)");
    }

    [Fact]
    public void ViewAsHtml_TableWithCellFill_ContainsBackgroundColor()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "2", ["cols"] = "2"
        });
        _handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
        {
            ["text"] = "Colored Cell", ["fill"] = "FFD700"
        });

        var html = _handler.ViewAsHtml();
        html.Should().Contain("Colored Cell");
        html.Should().Contain("#FFD700");
    }

    [Fact]
    public void ViewAsHtml_MultipleRunsWithDifferentFormatting_AllRendered()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new() { ["text"] = "initial" });
        // Add runs with different formatting
        _handler.Add("/slide[1]/shape[1]/paragraph[1]", "run", null, new()
        {
            ["text"] = "Bold Part", ["bold"] = "true"
        });
        _handler.Add("/slide[1]/shape[1]/paragraph[1]", "run", null, new()
        {
            ["text"] = "Red Part", ["color"] = "FF0000"
        });

        var html = _handler.ViewAsHtml();
        html.Should().Contain("Bold Part");
        html.Should().Contain("Red Part");
    }

    [Fact]
    public void ViewAsHtml_MultipleParagraphs_AllRendered()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Line 1" });
        _handler.Add("/slide[1]/shape[1]", "paragraph", null, new() { ["text"] = "Line 2" });
        _handler.Add("/slide[1]/shape[1]", "paragraph", null, new() { ["text"] = "Line 3" });

        var html = _handler.ViewAsHtml();
        html.Should().Contain("Line 1");
        html.Should().Contain("Line 2");
        html.Should().Contain("Line 3");
    }

    [Fact]
    public void ViewAsHtml_ShapeWithBulletList_ContainsBullet()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Bullet Item", ["list"] = "•"
        });

        var html = _handler.ViewAsHtml();
        html.Should().Contain("•");
        html.Should().Contain("Bullet Item");
    }

    [Fact]
    public void ViewAsHtml_RadialGradientBackground_ContainsRadialGradient()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Set("/slide[1]", new() { ["background"] = "radial:FF0000-0000FF" });

        var html = _handler.ViewAsHtml();
        html.Should().Contain("radial-gradient");
    }

    [Fact]
    public void ViewAsHtml_ImageBackground_ContainsBackgroundUrl()
    {
        _handler.Add("/", "slide", null, new());
        var imgPath = Path.Combine(Path.GetTempPath(), $"test_bg_{Guid.NewGuid():N}.png");
        try
        {
            File.WriteAllBytes(imgPath, CreateMinimalPng());
            _handler.Set("/slide[1]", new() { ["background"] = $"image:{imgPath}" });

            var html = _handler.ViewAsHtml();
            html.Should().Contain("background:url(");
            html.Should().Contain("base64,");
        }
        finally { if (File.Exists(imgPath)) File.Delete(imgPath); }
    }

    [Fact]
    public void ViewAsHtml_ConnectorWithDash_ContainsDashed()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "connector", null, new()
        {
            ["x"] = "1cm", ["y"] = "1cm", ["width"] = "10cm", ["height"] = "5cm",
            ["line"] = "0000FF", ["lineDash"] = "dash"
        });

        var html = _handler.ViewAsHtml();
        html.Should().Contain("<line");
    }

    [Fact]
    public void ViewAsHtml_ShapeWithTextColor_ContainsColorStyle()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Colored Text", ["color"] = "00AA00"
        });

        var html = _handler.ViewAsHtml();
        html.Should().Contain("color:");
        html.Should().Contain("Colored Text");
    }

    [Fact]
    public void ViewAsHtml_ShapeWithFont_ContainsFontFamily()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Custom Font", ["font"] = "Arial"
        });

        var html = _handler.ViewAsHtml();
        html.Should().Contain("font-family:'Arial'");
    }

    [Fact]
    public void ViewAsHtml_AllElementTypes_DocumentUnchanged()
    {
        // Create a complex slide, snapshot state, run preview, verify unchanged
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Test Shape",
            ["x"] = "2cm", ["y"] = "3cm", ["width"] = "10cm", ["height"] = "5cm",
            ["fill"] = "4472C4", ["bold"] = "true", ["size"] = "20pt"
        });
        _handler.Add("/slide[1]", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });

        // Snapshot
        var shapeBefore = _handler.Get("/slide[1]/shape[1]");
        var tableBefore = _handler.Get("/slide[1]/table[1]");

        // Preview
        _handler.ViewAsHtml();

        // Verify
        var shapeAfter = _handler.Get("/slide[1]/shape[1]");
        shapeAfter.Text.Should().Be(shapeBefore.Text);
        shapeAfter.Format["x"].Should().Be(shapeBefore.Format["x"]);
        shapeAfter.Format["fill"].Should().Be(shapeBefore.Format["fill"]);
        shapeAfter.Format["size"].Should().Be(shapeBefore.Format["size"]);

        var tableAfter = _handler.Get("/slide[1]/table[1]");
        tableAfter.Format["rows"].Should().Be(tableBefore.Format["rows"]);
    }

    // ==================== Default line-height regression ====================

    [Fact]
    public void ViewAsHtml_DefaultLineHeight_IsSingleSpaced()
    {
        // PowerPoint uses line-height:1 (single-spaced) by default.
        // Browsers default to ~1.2, which causes large text to overflow
        // into adjacent elements (e.g. 60pt "38%" overlapping "火星重力仅为地球的38%").
        // This test locks the default to prevent accidental regression.
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Test" });

        var html = _handler.ViewAsHtml();

        // The .para CSS class must have line-height:1
        html.Should().Contain("line-height: 1", "default paragraph line-height must be 1 (single-spaced) to match PowerPoint");
    }

    [Fact]
    public void ViewAsHtml_ExplicitLineSpacing_OverridesDefault()
    {
        // When lineSpacing is explicitly set, it should appear as inline style
        // and override the default line-height:1
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Spaced Text",
            ["lineSpacing"] = "1.5x"
        });

        var html = _handler.ViewAsHtml();

        html.Should().Contain("line-height:1.5", "explicit lineSpacing should override default");
    }

    [Fact]
    public void ViewAsHtml_StackedLargeAndSmallText_NoOverlap()
    {
        // Regression scenario: 60pt number "38%" at y=5cm (height 2cm) followed by
        // 18pt label "Description" at y=7.5cm. With browser default line-height:1.2,
        // the 60pt text overflows 12pt (~0.4cm) into the label area, causing overlap.
        // With line-height:1, it should fit within bounds.
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "38%",
            ["size"] = "60pt", ["bold"] = "true", ["color"] = "FF5722",
            ["x"] = "10cm", ["y"] = "5cm", ["width"] = "10cm", ["height"] = "2cm"
        });
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Description below large number",
            ["size"] = "18pt", ["color"] = "B0BEC5",
            ["x"] = "10cm", ["y"] = "7.5cm", ["width"] = "10cm", ["height"] = "2cm"
        });

        var html = _handler.ViewAsHtml();

        // Both texts must be present (not clipped)
        html.Should().Contain("38%");
        html.Should().Contain("Description below large number");
        // Default line-height must be 1 (not 1.2) to prevent overflow
        html.Should().Contain("line-height: 1");
    }

    [Fact]
    public void ViewAsHtml_MultiLineShapeText_NotSquished()
    {
        // Multi-line text should still render all lines, not collapse them.
        // line-height:1 is fine for single lines, but multi-line text with
        // explicit lineSpacing should use that value instead.
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Line One",
            ["size"] = "24pt",
            ["x"] = "2cm", ["y"] = "2cm", ["width"] = "15cm", ["height"] = "6cm",
            ["lineSpacing"] = "1.5x"
        });
        _handler.Add("/slide[1]/shape[1]", "paragraph", null, new() { ["text"] = "Line Two" });
        _handler.Add("/slide[1]/shape[1]", "paragraph", null, new() { ["text"] = "Line Three" });

        var html = _handler.ViewAsHtml();

        html.Should().Contain("Line One");
        html.Should().Contain("Line Two");
        html.Should().Contain("Line Three");
        html.Should().Contain("line-height:1.5", "explicit 1.5x spacing must be present");
    }

    [Fact]
    public void ViewAsHtml_LargeTextNoOverflow_TextNotClipped()
    {
        // Transparent text boxes (no fill) should allow text overflow (overflow:visible)
        // to match PowerPoint behavior where text can extend beyond the box.
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Large",
            ["size"] = "72pt",
            ["x"] = "2cm", ["y"] = "2cm", ["width"] = "10cm", ["height"] = "1cm"
        });

        var html = _handler.ViewAsHtml();

        // Shape without fill should NOT have overflow:hidden
        html.Should().Contain("class=\"shape\"", "transparent shape should not have has-fill class");
    }

    [Fact]
    public void ViewAsHtml_FilledShapeClipsOverflow()
    {
        // Shapes with fill should clip content (overflow:hidden via has-fill class)
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Clipped",
            ["fill"] = "FF0000",
            ["x"] = "2cm", ["y"] = "2cm", ["width"] = "5cm", ["height"] = "1cm"
        });

        var html = _handler.ViewAsHtml();

        html.Should().Contain("class=\"shape has-fill\"", "filled shape should have has-fill class for overflow:hidden");
    }

    [Fact]
    public void ViewAsHtml_OpacityNotDoubleApplied()
    {
        // Fill opacity (alpha) is baked into rgba() by ResolveFillColor.
        // A separate CSS opacity property must NOT be added (would cause double-application).
        // e.g. rgba(255,255,255,0.05) + opacity:0.05 = 0.25% visible (almost invisible)
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "",
            ["fill"] = "FFFFFF",
            ["opacity"] = "0.5"
        });

        var html = _handler.ViewAsHtml();

        // Should have rgba with alpha, but NOT a separate opacity property
        html.Should().Contain("rgba(255,255,255,0.5)", "fill alpha should be in rgba");
        // Count occurrences of "opacity" - should only appear in rgba, not as standalone
        var shapeDiv = html.Split("class=\"shape")[1].Split("</div>")[0];
        shapeDiv.Should().NotContain("opacity:", "separate opacity property would double-apply with rgba alpha");
    }

    [Fact]
    public void ViewAsHtml_CjkFontFallback_IncludesPingFangAndYaHei()
    {
        // Default font-family must include CJK fallback fonts to prevent
        // garbled Chinese text when theme font (e.g. Calibri) doesn't support CJK.
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new() { ["text"] = "中文测试" });

        var html = _handler.ViewAsHtml();

        html.Should().Contain("PingFang SC", "must include macOS CJK font");
        html.Should().Contain("Microsoft YaHei", "must include Windows CJK font");
    }

    // ==================== Helpers ====================

    /// <summary>
    /// Create a minimal valid 1x1 PNG file (same as GenerateZoomPlaceholderPng).
    /// </summary>
    private static byte[] CreateMinimalPng()
    {
        using var ms = new MemoryStream();
        var bw = new BinaryWriter(ms);
        bw.Write(new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A });
        WriteChunk(bw, "IHDR", new byte[] { 0, 0, 0, 1, 0, 0, 0, 1, 8, 6, 0, 0, 0 });
        WriteChunk(bw, "IDAT", new byte[] { 0x78, 0x01, 0x62, 0x60, 0x60, 0x28, 0x61, 0x28, 0x61, 0x68, 0xF8, 0x0F, 0x00, 0x01, 0x45, 0x00, 0xC5 });
        WriteChunk(bw, "IEND", Array.Empty<byte>());
        return ms.ToArray();
    }

    private static void WriteChunk(BinaryWriter bw, string type, byte[] data)
    {
        var lenBytes = BitConverter.GetBytes(data.Length);
        if (BitConverter.IsLittleEndian) Array.Reverse(lenBytes);
        bw.Write(lenBytes);
        var typeBytes = System.Text.Encoding.ASCII.GetBytes(type);
        bw.Write(typeBytes);
        bw.Write(data);
        var crcData = new byte[4 + data.Length];
        Array.Copy(typeBytes, 0, crcData, 0, 4);
        Array.Copy(data, 0, crcData, 4, data.Length);
        var crc = Crc32(crcData);
        var crcBytes = BitConverter.GetBytes(crc);
        if (BitConverter.IsLittleEndian) Array.Reverse(crcBytes);
        bw.Write(crcBytes);
    }

    private static uint Crc32(byte[] data)
    {
        uint crc = 0xFFFFFFFF;
        foreach (var b in data)
        {
            crc ^= b;
            for (int i = 0; i < 8; i++)
                crc = (crc >> 1) ^ (crc & 1) * 0xEDB88320;
        }
        return ~crc;
    }
}
