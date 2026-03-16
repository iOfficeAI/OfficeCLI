// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Functional tests for PPTX: each test creates a blank file, adds elements,
/// queries them, and modifies them — exercising the full Create→Add→Get→Set lifecycle.
/// </summary>
public class PptxFunctionalTests : IDisposable
{
    private readonly string _path;
    private PowerPointHandler _handler;

    public PptxFunctionalTests()
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

    // Reopen the file to verify persistence
    private PowerPointHandler Reopen()
    {
        _handler.Dispose();
        _handler = new PowerPointHandler(_path, editable: true);
        return _handler;
    }

    // ==================== Slide lifecycle ====================

    [Fact]
    public void AddSlide_ReturnsPath_Slide1()
    {
        var path = _handler.Add("/", "slide", null, new Dictionary<string, string>());
        path.Should().Be("/slide[1]");
    }

    [Fact]
    public void AddSlide_Get_ReturnsSlideType()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        var node = _handler.Get("/slide[1]");
        node.Type.Should().Be("slide");
    }

    [Fact]
    public void AddSlide_Multiple_PathIncrements()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        var path3 = _handler.Add("/", "slide", null, new Dictionary<string, string>());
        path3.Should().Be("/slide[3]");
    }

    [Fact]
    public void AddSlide_WithTitle_TitleVisibleInText()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string> { ["title"] = "Hello World" });
        var node = _handler.Get("/slide[1]", depth: 2);
        var allText = node.Children.SelectMany(c => c.Children).Select(c => c.Text).Concat(
                      node.Children.Select(c => c.Text))
                      .Where(t => t != null).ToList();
        allText.Any(t => t!.Contains("Hello World")).Should().BeTrue();
    }

    // ==================== Shape lifecycle ====================

    [Fact]
    public void AddShape_ReturnsPath_Shape1()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        var shapePath = _handler.Add("/slide[1]", "shape", null,
            new Dictionary<string, string> { ["text"] = "Test" });
        shapePath.Should().Be("/slide[1]/shape[1]");
    }

    [Fact]
    public void AddShape_WithText_TextIsReadBack()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null,
            new Dictionary<string, string> { ["text"] = "Hello Shape" });

        var node = _handler.Get("/slide[1]/shape[1]");
        node.Text.Should().Be("Hello Shape");
    }

    [Fact]
    public void AddShape_WithFill_FillIsReadBack()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null,
            new Dictionary<string, string> { ["text"] = "Filled", ["fill"] = "FF0000" });

        var node = _handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("fill");
        node.Format["fill"].Should().Be("FF0000");
    }

    [Fact]
    public void AddShape_WithPosition_PositionIsReadBack()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["text"] = "Positioned",
            ["x"] = "2cm",
            ["y"] = "3cm",
            ["width"] = "5cm",
            ["height"] = "2cm"
        });

        var node = _handler.Get("/slide[1]/shape[1]");
        node.Format["x"].Should().Be("2cm");
        node.Format["y"].Should().Be("3cm");
        node.Format["width"].Should().Be("5cm");
        node.Format["height"].Should().Be("2cm");
    }

    [Fact]
    public void AddShape_WithName_NameIsReadBack()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null,
            new Dictionary<string, string> { ["text"] = "Named", ["name"] = "MyBox" });

        var node = _handler.Get("/slide[1]/shape[1]");
        node.Format["name"].Should().Be("MyBox");
    }

    // ==================== Set: modify shape properties ====================

    [Fact]
    public void SetShape_Bold_BoldIsReadBack()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null,
            new Dictionary<string, string> { ["text"] = "Normal" });

        _handler.Set("/slide[1]/shape[1]", new Dictionary<string, string> { ["bold"] = "true" });

        var node = _handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("bold");
        node.Format["bold"].Should().Be(true);
    }

    [Fact]
    public void SetShape_Italic_ItalicIsReadBack()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null,
            new Dictionary<string, string> { ["text"] = "Normal" });

        _handler.Set("/slide[1]/shape[1]", new Dictionary<string, string> { ["italic"] = "true" });

        var node = _handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("italic");
        node.Format["italic"].Should().Be(true);
    }

    [Fact]
    public void SetShape_Fill_FillIsUpdated()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null,
            new Dictionary<string, string> { ["text"] = "A", ["fill"] = "0000FF" });

        // Change fill from blue to red
        _handler.Set("/slide[1]/shape[1]", new Dictionary<string, string> { ["fill"] = "FF0000" });

        var node = _handler.Get("/slide[1]/shape[1]");
        node.Format["fill"].Should().Be("FF0000");
    }

    [Fact]
    public void SetShape_FontSize_SizeIsReadBack()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null,
            new Dictionary<string, string> { ["text"] = "Big Text" });

        // size property accepts a raw point number (stored as pt*100 internally)
        _handler.Set("/slide[1]/shape[1]", new Dictionary<string, string> { ["size"] = "24" });

        var node = _handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("size");
        node.Format["size"].Should().Be("24pt");
    }

    [Fact]
    public void SetShape_Position_PositionIsUpdated()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null, new Dictionary<string, string> { ["text"] = "A" });

        _handler.Set("/slide[1]/shape[1]", new Dictionary<string, string>
        {
            ["x"] = "4cm",
            ["y"] = "5cm"
        });

        var node = _handler.Get("/slide[1]/shape[1]");
        node.Format["x"].Should().Be("4cm");
        node.Format["y"].Should().Be("5cm");
    }

    // ==================== Query ====================

    [Fact]
    public void QueryShapes_AfterAddTwo_ReturnsBoth()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null, new Dictionary<string, string> { ["text"] = "A" });
        _handler.Add("/slide[1]", "shape", null, new Dictionary<string, string> { ["text"] = "B" });

        var nodes = _handler.Query("shape");
        nodes.Should().HaveCountGreaterThanOrEqualTo(2);
    }

    [Fact]
    public void GetRoot_AfterAddThreeSlides_HasThreeChildren()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/", "slide", null, new Dictionary<string, string>());

        var root = _handler.Get("/");
        root.Children.Should().HaveCount(3);
        root.Children.Should().AllSatisfy(c => c.Type.Should().Be("slide"));
    }

    // ==================== Table lifecycle ====================

    [Fact]
    public void AddTable_ReturnsTablePath()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        var path = _handler.Add("/slide[1]", "table", null,
            new Dictionary<string, string> { ["rows"] = "2", ["cols"] = "3" });
        path.Should().Be("/slide[1]/table[1]");
    }

    [Fact]
    public void AddTable_Get_HasCorrectDimensions()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "table", null,
            new Dictionary<string, string> { ["rows"] = "3", ["cols"] = "4" });

        var node = _handler.Get("/slide[1]/table[1]");
        node.Type.Should().Be("table");
        node.Format["rows"].Should().Be(3);
        node.Format["cols"].Should().Be(4);
    }

    [Fact]
    public void SetTableCell_TextIsReadBack()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "table", null,
            new Dictionary<string, string> { ["rows"] = "2", ["cols"] = "2" });

        _handler.Set("/slide[1]/table[1]/tr[1]/tc[1]",
            new Dictionary<string, string> { ["text"] = "Cell A1" });

        var table = _handler.Get("/slide[1]/table[1]", depth: 2);
        var cell = table.Children
            .FirstOrDefault(r => r.Type == "tr")
            ?.Children.FirstOrDefault(c => c.Type == "tc");
        cell.Should().NotBeNull();
        cell!.Text.Should().Be("Cell A1");
    }

    // ==================== Slide background ====================

    [Fact]
    public void AddSlide_WithBackground_BackgroundIsReadBack()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string> { ["background"] = "FF0000" });

        var node = _handler.Get("/slide[1]");
        node.Format.Should().ContainKey("background");
        node.Format["background"].Should().Be("FF0000");
    }

    [Fact]
    public void AddSlide_WithGradientBackground_BackgroundIsReadBack()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string> { ["background"] = "FF0000-0000FF" });

        var node = _handler.Get("/slide[1]");
        node.Format.Should().ContainKey("background");
        var bg = node.Format["background"]?.ToString();
        bg.Should().NotBeNull();
        bg!.Should().Contain("FF0000");
        bg.Should().Contain("0000FF");
    }

    [Fact]
    public void SetSlideBackground_SolidColor_IsReadBack()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Set("/slide[1]", new Dictionary<string, string> { ["background"] = "FF0000" });

        var node = _handler.Get("/slide[1]");
        node.Format.Should().ContainKey("background");
        node.Format["background"].Should().Be("FF0000");
    }

    [Fact]
    public void SetSlideBackground_UpdateColor_NewColorIsReadBack()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string> { ["background"] = "FF0000" });

        _handler.Set("/slide[1]", new Dictionary<string, string> { ["background"] = "0000FF" });

        var node = _handler.Get("/slide[1]");
        node.Format["background"].Should().Be("0000FF");
    }

    [Fact]
    public void SetSlideBackground_None_RemovesBackground()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string> { ["background"] = "FF0000" });
        _handler.Set("/slide[1]", new Dictionary<string, string> { ["background"] = "none" });

        var node = _handler.Get("/slide[1]");
        node.Format.Should().NotContainKey("background");
    }

    [Fact]
    public void AddSlide_WithBackground_Persist_SurvivesReopenFile()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string> { ["background"] = "00FF00" });

        var handler2 = Reopen();
        var node = handler2.Get("/slide[1]");
        node.Format.Should().ContainKey("background");
        node.Format["background"].Should().Be("00FF00");
    }

    // ==================== Persistence ====================

    [Fact]
    public void AddShape_Persist_SurvivesReopenFile()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null,
            new Dictionary<string, string> { ["text"] = "Persistent" });

        var handler2 = Reopen();
        var node = handler2.Get("/slide[1]/shape[1]");
        node.Text.Should().Be("Persistent");
    }

    // ==================== Speaker Notes ====================

    [Fact]
    public void Notes_Lifecycle()
    {
        // 1. Add slide + notes
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        var path = _handler.Add("/slide[1]", "notes", null, new Dictionary<string, string> { ["text"] = "Original note" });
        path.Should().Be("/slide[1]/notes");

        // 2. Get + Verify
        var node = _handler.Get("/slide[1]/notes");
        node.Type.Should().Be("notes");
        node.Text.Should().Be("Original note");

        // 3. Query + Verify
        var results = _handler.Query("notes");
        results.Should().Contain(n => n.Type == "notes" && n.Text == "Original note");

        // 4. Set + Verify
        _handler.Set("/slide[1]/notes", new Dictionary<string, string> { ["text"] = "Updated note" });
        node = _handler.Get("/slide[1]/notes");
        node.Text.Should().Be("Updated note");

        // 5. Query reflects update
        results = _handler.Query("notes");
        results.Should().Contain(n => n.Text == "Updated note");
    }

    [Fact]
    public void Notes_Persist_SurvivesReopenFile()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "notes", null, new Dictionary<string, string> { ["text"] = "Persist me" });
        _handler.Set("/slide[1]/notes", new Dictionary<string, string> { ["text"] = "Persisted note" });

        var handler2 = Reopen();
        var node = handler2.Get("/slide[1]/notes");
        node.Text.Should().Be("Persisted note");
    }

    // ==================== PPTX Hyperlinks ====================

    [Fact]
    public void ShapeLink_Lifecycle()
    {
        // 1. Add slide + shape with link
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["text"] = "Click me",
            ["link"] = "https://first.com"
        });

        // 2. Get + Verify
        var node = _handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("link");
        ((string)node.Format["link"]).Should().StartWith("https://first.com");

        // 3. Set new link + Verify
        _handler.Set("/slide[1]/shape[1]", new Dictionary<string, string> { ["link"] = "https://updated.com" });
        node = _handler.Get("/slide[1]/shape[1]");
        ((string)node.Format["link"]).Should().StartWith("https://updated.com");

        // 4. Remove link + Verify
        _handler.Set("/slide[1]/shape[1]", new Dictionary<string, string> { ["link"] = "none" });
        node = _handler.Get("/slide[1]/shape[1]");
        node.Format.Should().NotContainKey("link");
    }

    [Fact]
    public void ShapeLink_Persist_SurvivesReopenFile()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["text"] = "Persistent link",
            ["link"] = "https://persist.com"
        });

        var handler2 = Reopen();
        var node = handler2.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("link");
        ((string)node.Format["link"]).Should().StartWith("https://persist.com");
    }

    // ==================== PPTX lineDash ====================

    [Fact]
    public void ShapeLineDash_Lifecycle()
    {
        // 1. Add shape with lineDash
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["text"] = "dashed border",
            ["line"] = "FF0000",
            ["lineDash"] = "dash"
        });

        // 2. Get + Verify
        var node = _handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("lineDash");
        node.Format["lineDash"].Should().Be("dash");

        // 3. Set new dash style + Verify
        _handler.Set("/slide[1]/shape[1]", new Dictionary<string, string> { ["lineDash"] = "dot" });
        node = _handler.Get("/slide[1]/shape[1]");
        node.Format["lineDash"].Should().Be("dot");

        // 4. Set solid (remove dash) + Verify
        _handler.Set("/slide[1]/shape[1]", new Dictionary<string, string> { ["lineDash"] = "solid" });
        node = _handler.Get("/slide[1]/shape[1]");
        node.Format["lineDash"].Should().Be("solid");
    }

    [Fact]
    public void ShapeLineDash_Persist_SurvivesReopenFile()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["text"] = "persist dash",
            ["line"] = "0000FF",
            ["lineDash"] = "dashdot"
        });

        Reopen();
        var node = _handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("lineDash");
        node.Format["lineDash"].Should().Be("dashdot");
    }

    // ==================== PPTX Effects (shadow / glow / reflection) ====================

    [Fact]
    public void ShapeShadow_Lifecycle()
    {
        // 1. Add shape with shadow
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["text"] = "shadowed",
            ["shadow"] = "000000"
        });

        // 2. Get + Verify
        var node = _handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("shadow");
        node.Format["shadow"].Should().Be("000000");

        // 3. Set new shadow color + Verify
        _handler.Set("/slide[1]/shape[1]", new Dictionary<string, string> { ["shadow"] = "404040" });
        node = _handler.Get("/slide[1]/shape[1]");
        node.Format["shadow"].Should().Be("404040");

        // 4. Remove shadow + Verify
        _handler.Set("/slide[1]/shape[1]", new Dictionary<string, string> { ["shadow"] = "none" });
        node = _handler.Get("/slide[1]/shape[1]");
        node.Format.Should().NotContainKey("shadow");
    }

    [Fact]
    public void ShapeGlow_Lifecycle()
    {
        // 1. Add shape with glow
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["text"] = "glowing",
            ["glow"] = "0070FF"
        });

        // 2. Get + Verify
        var node = _handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("glow");
        node.Format["glow"].Should().Be("0070FF");

        // 3. Set new glow + Verify
        _handler.Set("/slide[1]/shape[1]", new Dictionary<string, string> { ["glow"] = "FF0000-10" });
        node = _handler.Get("/slide[1]/shape[1]");
        node.Format["glow"].Should().Be("FF0000");

        // 4. Remove glow + Verify
        _handler.Set("/slide[1]/shape[1]", new Dictionary<string, string> { ["glow"] = "none" });
        node = _handler.Get("/slide[1]/shape[1]");
        node.Format.Should().NotContainKey("glow");
    }

    [Fact]
    public void ShapeReflection_Lifecycle()
    {
        // 1. Add shape with reflection
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["text"] = "reflected",
            ["reflection"] = "half"
        });

        // 2. Get + Verify
        var node = _handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("reflection");
        node.Format["reflection"].Should().Be("true");

        // 3. Remove reflection + Verify
        _handler.Set("/slide[1]/shape[1]", new Dictionary<string, string> { ["reflection"] = "none" });
        node = _handler.Get("/slide[1]/shape[1]");
        node.Format.Should().NotContainKey("reflection");

        // 4. Re-add via Set + Verify
        _handler.Set("/slide[1]/shape[1]", new Dictionary<string, string> { ["reflection"] = "full" });
        node = _handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("reflection");
    }

    // ==================== PPTX Animation ====================

    [Fact]
    public void ShapeAnimation_Lifecycle()
    {
        // 1. Add shape with animation
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["text"] = "animated",
            ["animation"] = "fade"
        });

        // 2. Get shape — shape itself is accessible
        var node = _handler.Get("/slide[1]/shape[1]");
        node.Text.Should().Be("animated");

        // 3. Set different animation + Verify shape still accessible
        _handler.Set("/slide[1]/shape[1]", new Dictionary<string, string> { ["animation"] = "fly" });
        node = _handler.Get("/slide[1]/shape[1]");
        node.Text.Should().Be("animated");

        // 4. Remove animation + Verify
        _handler.Set("/slide[1]/shape[1]", new Dictionary<string, string> { ["animation"] = "none" });
        node = _handler.Get("/slide[1]/shape[1]");
        node.Text.Should().Be("animated");
    }

    [Fact]
    public void ShapeAnimation_Persist_SurvivesReopenFile()
    {
        _handler.Add("/", "slide", null, new Dictionary<string, string>());
        _handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
        {
            ["text"] = "persist anim",
            ["animation"] = "fade-entrance-500"
        });

        Reopen();
        var node = _handler.Get("/slide[1]/shape[1]");
        node.Text.Should().Be("persist anim");
    }

    // ==================== Radial Gradient ====================

    [Fact]
    public void RadialGradient_Shape_Lifecycle()
    {
        // 1. Add slide + shape
        _handler.Add("/", "slide", null, new() { ["title"] = "Gradient Test" });
        _handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Radial" });

        // 2. Set radial gradient (blue→purple, top-right focus)
        _handler.Set("/slide[1]/shape[1]", new() { ["gradient"] = "radial:1E90FF-4B0082-tr" });

        // 3. Get + Verify
        var node = _handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("gradient");
        var grad = (string)node.Format["gradient"];
        grad.Should().StartWith("radial:");
        grad.Should().Contain("1E90FF");
        grad.Should().Contain("4B0082");
        grad.Should().EndWith("tr");

        // 4. Change to center focus
        _handler.Set("/slide[1]/shape[1]", new() { ["gradient"] = "radial:FF0000-FFFF00-center" });
        node = _handler.Get("/slide[1]/shape[1]");
        ((string)node.Format["gradient"]).Should().Contain("center");

        // 5. Persist + Verify
        Reopen();
        node = _handler.Get("/slide[1]/shape[1]");
        ((string)node.Format["gradient"]).Should().StartWith("radial:");
    }

    [Fact]
    public void RadialGradient_Background_Lifecycle()
    {
        // 1. Add slide
        _handler.Add("/", "slide", null, new() { ["title"] = "BG Gradient" });

        // 2. Set radial gradient as background
        _handler.Set("/slide[1]", new() { ["background"] = "radial:4B0082-1E90FF-bl" });

        // 3. Get + Verify
        var node = _handler.Get("/slide[1]");
        node.Format.Should().ContainKey("background");
        var bg = (string)node.Format["background"];
        bg.Should().StartWith("radial:");
        bg.Should().Contain("bl");

        // 4. Persist + Verify
        Reopen();
        node = _handler.Get("/slide[1]");
        ((string)node.Format["background"]).Should().StartWith("radial:");
    }

    // ==================== Line Opacity ====================

    [Fact]
    public void LineOpacity_Lifecycle()
    {
        // 1. Add slide + shape with line
        _handler.Add("/", "slide", null, new() { ["title"] = "Line Test" });
        _handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Bordered" });

        // 2. Set line color + opacity
        _handler.Set("/slide[1]/shape[1]", new()
        {
            ["line"] = "FFFFFF",
            ["linewidth"] = "2pt",
            ["lineopacity"] = "0.5"
        });

        // 3. Get + Verify
        var node = _handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("line");
        ((string)node.Format["line"]).Should().Be("FFFFFF");
        node.Format.Should().ContainKey("lineOpacity");
        ((string)node.Format["lineOpacity"]).Should().Be("0.5");

        // 4. Change opacity
        _handler.Set("/slide[1]/shape[1]", new() { ["lineopacity"] = "0.3" });
        node = _handler.Get("/slide[1]/shape[1]");
        ((string)node.Format["lineOpacity"]).Should().Be("0.3");

        // 5. Persist + Verify
        Reopen();
        node = _handler.Get("/slide[1]/shape[1]");
        ((string)node.Format["lineOpacity"]).Should().Be("0.3");
    }

    // ==================== Shape Image Fill ====================

    [Fact]
    public void ShapeImageFill_Lifecycle()
    {
        // 1. Create a tiny test image
        var imgPath = Path.Combine(Path.GetTempPath(), $"test_img_{Guid.NewGuid():N}.png");
        try
        {
            // Write a minimal 1x1 PNG
            var pngBytes = new byte[]
            {
                0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A, // PNG signature
                0x00,0x00,0x00,0x0D,0x49,0x48,0x44,0x52,  // IHDR chunk
                0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x01,0x08,0x02,0x00,0x00,0x00,0x90,0x77,0x53,0xDE,
                0x00,0x00,0x00,0x0C,0x49,0x44,0x41,0x54,  // IDAT chunk
                0x08,0xD7,0x63,0xF8,0xCF,0xC0,0x00,0x00,0x00,0x02,0x00,0x01,0xE2,0x21,0xBC,0x33,
                0x00,0x00,0x00,0x00,0x49,0x45,0x4E,0x44,0xAE,0x42,0x60,0x82 // IEND chunk
            };
            File.WriteAllBytes(imgPath, pngBytes);

            // 2. Add slide + shape
            _handler.Add("/", "slide", null, new() { ["title"] = "Image Fill" });
            _handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Filled" });

            // 3. Set image fill
            _handler.Set("/slide[1]/shape[1]", new() { ["image"] = imgPath });

            // 4. Get + Verify
            var node = _handler.Get("/slide[1]/shape[1]");
            node.Format.Should().ContainKey("image");
            ((string)node.Format["image"]).Should().Be("true");

            // 5. Persist + Verify
            Reopen();
            node = _handler.Get("/slide[1]/shape[1]");
            node.Format.Should().ContainKey("image");
        }
        finally
        {
            if (File.Exists(imgPath)) File.Delete(imgPath);
        }
    }

    // ==================== Slide Layout ====================

    [Fact]
    public void SlideLayout_Default_HasLayoutInfo()
    {
        // 1. Add slide without specifying layout
        _handler.Add("/", "slide", null, new() { ["title"] = "Default Layout" });

        // 2. Get + Verify layout info is returned
        var node = _handler.Get("/slide[1]");
        node.Format.Should().ContainKey("layout");
        ((string)node.Format["layout"]).Should().NotBeNullOrEmpty();

        // 3. Persist + Verify
        Reopen();
        node = _handler.Get("/slide[1]");
        node.Format.Should().ContainKey("layout");
    }

    [Fact]
    public void SlideLayout_ByIndex_SelectsDifferentLayout()
    {
        // 1. Add two slides with different layout indices
        _handler.Add("/", "slide", null, new() { ["title"] = "Layout 1", ["layout"] = "1" });
        _handler.Add("/", "slide", null, new() { ["title"] = "Layout 2", ["layout"] = "2" });

        // 2. Get + Verify they have different layouts
        var node1 = _handler.Get("/slide[1]");
        var node2 = _handler.Get("/slide[2]");
        node1.Format.Should().ContainKey("layout");
        node2.Format.Should().ContainKey("layout");

        // Layout names should be different (assuming blank doc has >1 layout)
        var layout1 = (string)node1.Format["layout"];
        var layout2 = (string)node2.Format["layout"];
        layout1.Should().NotBeNullOrEmpty();
        layout2.Should().NotBeNullOrEmpty();
        layout1.Should().NotBe(layout2, "different layout indices should yield different layouts");

        // 3. Persist + Verify
        Reopen();
        node1 = _handler.Get("/slide[1]");
        node2 = _handler.Get("/slide[2]");
        ((string)node1.Format["layout"]).Should().Be(layout1);
        ((string)node2.Format["layout"]).Should().Be(layout2);
    }

    [Fact]
    public void SlideLayout_ByType_Blank()
    {
        // 1. Add slide with blank layout type
        _handler.Add("/", "slide", null, new() { ["layout"] = "blank" });

        // 2. Get + Verify
        var node = _handler.Get("/slide[1]");
        node.Format.Should().ContainKey("layoutType");
        ((string)node.Format["layoutType"]).Should().Be("blank");

        // 3. Persist + Verify
        Reopen();
        node = _handler.Get("/slide[1]");
        ((string)node.Format["layoutType"]).Should().Be("blank");
    }

    [Fact]
    public void SlideLayout_ByName_MatchesLayoutName()
    {
        // 1. Get the name of layout #1 to use as a name lookup
        _handler.Add("/", "slide", null, new() { ["layout"] = "1" });
        var node = _handler.Get("/slide[1]");
        var layoutName = (string)node.Format["layout"];

        // 2. Add another slide using that layout name
        _handler.Add("/", "slide", null, new() { ["title"] = "By Name", ["layout"] = layoutName });
        var node2 = _handler.Get("/slide[2]");
        ((string)node2.Format["layout"]).Should().Be(layoutName);
    }

    [Fact]
    public void SlideLayout_RootGet_ShowsLayoutPerSlide()
    {
        // 1. Add slides with different layouts
        _handler.Add("/", "slide", null, new() { ["title"] = "Slide A", ["layout"] = "1" });
        _handler.Add("/", "slide", null, new() { ["title"] = "Slide B", ["layout"] = "blank" });

        // 2. Get root
        var root = _handler.Get("/");
        root.Children.Should().HaveCount(2);
        root.Children[0].Format.Should().ContainKey("layout");
        root.Children[1].Format.Should().ContainKey("layout");
    }

    // ==================== Charts ====================

    [Fact]
    public void Chart_Column_Lifecycle()
    {
        // 1. Add slide + column chart
        _handler.Add("/", "slide", null, new() { ["title"] = "Chart Test" });
        var path = _handler.Add("/slide[1]", "chart", null, new()
        {
            ["charttype"] = "column",
            ["title"] = "Q1 Sales",
            ["categories"] = "Jan,Feb,Mar",
            ["series1"] = "Revenue:100,200,300",
            ["series2"] = "Cost:80,150,250"
        });
        path.Should().Be("/slide[1]/chart[1]");

        // 2. Get + Verify
        var node = _handler.Get("/slide[1]/chart[1]");
        node.Type.Should().Be("chart");
        node.Format.Should().ContainKey("chartType");
        ((string)node.Format["chartType"]).Should().Be("column");
        ((string)node.Format["title"]).Should().Be("Q1 Sales");
        ((int)node.Format["seriesCount"]).Should().Be(2);
        ((string)node.Format["categories"]).Should().Be("Jan,Feb,Mar");

        // 3. Set — change title
        _handler.Set("/slide[1]/chart[1]", new() { ["title"] = "Updated Sales" });

        // 4. Get + Verify title changed
        node = _handler.Get("/slide[1]/chart[1]");
        ((string)node.Format["title"]).Should().Be("Updated Sales");

        // 5. Persist + Verify
        Reopen();
        node = _handler.Get("/slide[1]/chart[1]");
        ((string)node.Format["chartType"]).Should().Be("column");
        ((string)node.Format["title"]).Should().Be("Updated Sales");
        ((int)node.Format["seriesCount"]).Should().Be(2);
    }

    [Fact]
    public void Chart_Bar_Lifecycle()
    {
        _handler.Add("/", "slide", null, new() { ["title"] = "Bar Chart" });
        _handler.Add("/slide[1]", "chart", null, new()
        {
            ["charttype"] = "bar",
            ["title"] = "Horizontal",
            ["categories"] = "A,B,C",
            ["series1"] = "Data:10,20,30"
        });

        var node = _handler.Get("/slide[1]/chart[1]");
        ((string)node.Format["chartType"]).Should().Be("bar");

        Reopen();
        node = _handler.Get("/slide[1]/chart[1]");
        ((string)node.Format["chartType"]).Should().Be("bar");
    }

    [Fact]
    public void Chart_Line_Lifecycle()
    {
        _handler.Add("/", "slide", null, new() { ["title"] = "Line Chart" });
        _handler.Add("/slide[1]", "chart", null, new()
        {
            ["charttype"] = "line",
            ["title"] = "Trend",
            ["categories"] = "Q1,Q2,Q3,Q4",
            ["data"] = "Sales:10,25,30,45;Profit:5,12,18,30"
        });

        var node = _handler.Get("/slide[1]/chart[1]");
        ((string)node.Format["chartType"]).Should().Be("line");
        ((int)node.Format["seriesCount"]).Should().Be(2);

        Reopen();
        node = _handler.Get("/slide[1]/chart[1]");
        ((string)node.Format["chartType"]).Should().Be("line");
    }

    [Fact]
    public void Chart_Pie_Lifecycle()
    {
        _handler.Add("/", "slide", null, new() { ["title"] = "Pie Chart" });
        _handler.Add("/slide[1]", "chart", null, new()
        {
            ["charttype"] = "pie",
            ["title"] = "Market Share",
            ["categories"] = "Apple,Google,Microsoft",
            ["series1"] = "Share:40,30,30"
        });

        var node = _handler.Get("/slide[1]/chart[1]");
        ((string)node.Format["chartType"]).Should().Be("pie");
        ((string)node.Format["categories"]).Should().Be("Apple,Google,Microsoft");

        Reopen();
        node = _handler.Get("/slide[1]/chart[1]");
        ((string)node.Format["chartType"]).Should().Be("pie");
    }

    [Fact]
    public void Chart_SeriesData_ReadbackAtDepth()
    {
        _handler.Add("/", "slide", null, new() { ["title"] = "Data Test" });
        _handler.Add("/slide[1]", "chart", null, new()
        {
            ["charttype"] = "column",
            ["categories"] = "A,B",
            ["series1"] = "S1:10,20",
            ["series2"] = "S2:30,40"
        });

        // depth=0: just seriesCount
        var node0 = _handler.Get("/slide[1]/chart[1]", 0);
        ((int)node0.Format["seriesCount"]).Should().Be(2);
        node0.Children.Should().BeEmpty();

        // depth=1: series children with values
        var node1 = _handler.Get("/slide[1]/chart[1]", 1);
        node1.Children.Should().HaveCount(2);
        node1.Children[0].Type.Should().Be("series");
        node1.Children[0].Text.Should().Be("S1");
        ((string)node1.Children[0].Format["values"]).Should().Be("10,20");
        node1.Children[1].Text.Should().Be("S2");
        ((string)node1.Children[1].Format["values"]).Should().Be("30,40");
    }

    [Fact]
    public void Chart_Query_FindsCharts()
    {
        _handler.Add("/", "slide", null, new() { ["title"] = "Slide 1" });
        _handler.Add("/slide[1]", "chart", null, new()
        {
            ["charttype"] = "column",
            ["title"] = "Revenue",
            ["series1"] = "Data:1,2,3"
        });
        _handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Not a chart" });

        var charts = _handler.Query("chart");
        charts.Should().HaveCount(1);
        charts[0].Type.Should().Be("chart");
        ((string)charts[0].Format["title"]).Should().Be("Revenue");
    }

    [Fact]
    public void Chart_SetLegend_Lifecycle()
    {
        _handler.Add("/", "slide", null, new() { ["title"] = "Legend Test" });
        _handler.Add("/slide[1]", "chart", null, new()
        {
            ["charttype"] = "column",
            ["series1"] = "A:1,2",
            ["legend"] = "top"
        });

        var node = _handler.Get("/slide[1]/chart[1]");
        node.Format.Should().ContainKey("legend");
        ((string)node.Format["legend"]).Should().Be("t");

        // Change legend to none
        _handler.Set("/slide[1]/chart[1]", new() { ["legend"] = "none" });
        node = _handler.Get("/slide[1]/chart[1]");
        node.Format.Should().NotContainKey("legend");

        // Set legend back
        _handler.Set("/slide[1]/chart[1]", new() { ["legend"] = "right" });
        node = _handler.Get("/slide[1]/chart[1]");
        ((string)node.Format["legend"]).Should().Be("r");

        Reopen();
        node = _handler.Get("/slide[1]/chart[1]");
        ((string)node.Format["legend"]).Should().Be("r");
    }

    [Fact]
    public void Chart_Doughnut_Lifecycle()
    {
        _handler.Add("/", "slide", null, new() { ["title"] = "Donut" });
        _handler.Add("/slide[1]", "chart", null, new()
        {
            ["charttype"] = "doughnut",
            ["title"] = "Budget",
            ["categories"] = "Rent,Food,Transport",
            ["series1"] = "Spending:1200,800,400"
        });

        var node = _handler.Get("/slide[1]/chart[1]");
        ((string)node.Format["chartType"]).Should().Be("doughnut");

        Reopen();
        node = _handler.Get("/slide[1]/chart[1]");
        ((string)node.Format["chartType"]).Should().Be("doughnut");
    }

    [Fact]
    public void Chart_SlideChildNodes_IncludesChart()
    {
        _handler.Add("/", "slide", null, new() { ["title"] = "Mixed" });
        _handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Hello" });
        _handler.Add("/slide[1]", "chart", null, new()
        {
            ["charttype"] = "line",
            ["series1"] = "Data:1,2,3"
        });

        var slide = _handler.Get("/slide[1]");
        slide.Children.Should().Contain(c => c.Type == "chart");
        slide.Children.Should().Contain(c => c.Type == "textbox" || c.Type == "title");
    }

    // ==================== Theme Colors ====================

    [Fact]
    public void ThemeColor_Fill_Lifecycle()
    {
        // 1. Add slide + shape with theme color fill
        _handler.Add("/", "slide", null, new() { ["title"] = "Theme Test" });
        _handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Accent", ["fill"] = "accent1" });

        // 2. Get + Verify
        var node = _handler.Get("/slide[1]/shape[2]");
        ((string)node.Format["fill"]).Should().Be("accent1");

        // 3. Set to different theme color
        _handler.Set("/slide[1]/shape[2]", new() { ["fill"] = "accent3" });
        node = _handler.Get("/slide[1]/shape[2]");
        ((string)node.Format["fill"]).Should().Be("accent3");

        // 4. Persist + Verify
        Reopen();
        node = _handler.Get("/slide[1]/shape[2]");
        ((string)node.Format["fill"]).Should().Be("accent3");
    }

    [Fact]
    public void ThemeColor_TextAndLine_Lifecycle()
    {
        _handler.Add("/", "slide", null, new() { ["title"] = "Theme" });
        _handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Styled" });

        // Set theme colors for text and line
        _handler.Set("/slide[1]/shape[2]", new()
        {
            ["color"] = "dk1",
            ["line"] = "accent2"
        });

        var node = _handler.Get("/slide[1]/shape[2]");
        ((string)node.Format["color"]).Should().Be("dk1");
        ((string)node.Format["line"]).Should().Be("accent2");

        Reopen();
        node = _handler.Get("/slide[1]/shape[2]");
        ((string)node.Format["color"]).Should().Be("dk1");
    }

    // ==================== Connectors ====================

    [Fact]
    public void Connector_Lifecycle()
    {
        // 1. Add slide + connector
        _handler.Add("/", "slide", null, new() { ["title"] = "Flow" });
        var path = _handler.Add("/slide[1]", "connector", null, new()
        {
            ["x"] = "2cm",
            ["y"] = "5cm",
            ["width"] = "6cm",
            ["height"] = "0cm",
            ["line"] = "000000",
            ["linewidth"] = "2pt"
        });
        path.Should().StartWith("/slide[1]/connector[");

        // 2. Persist + Verify (connector should survive reopen)
        Reopen();
        // Connectors are GraphicFrame-less, they're direct children
        var slide = _handler.Get("/slide[1]");
        slide.Should().NotBeNull();
    }

    [Fact]
    public void Connector_Elbow_Lifecycle()
    {
        _handler.Add("/", "slide", null, new() { ["title"] = "Elbow" });
        _handler.Add("/slide[1]", "connector", null, new()
        {
            ["preset"] = "elbow",
            ["x"] = "1cm",
            ["y"] = "2cm",
            ["width"] = "5cm",
            ["height"] = "3cm",
            ["line"] = "accent1"
        });

        Reopen();
        var slide = _handler.Get("/slide[1]");
        slide.Should().NotBeNull();
    }

    // ==================== Group Shapes ====================

    [Fact]
    public void Group_Lifecycle()
    {
        // 1. Add slide + 3 shapes
        _handler.Add("/", "slide", null, new() { ["title"] = "Group Test" });
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "A", ["x"] = "1cm", ["y"] = "1cm", ["width"] = "3cm", ["height"] = "2cm"
        });
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "B", ["x"] = "5cm", ["y"] = "1cm", ["width"] = "3cm", ["height"] = "2cm"
        });

        // Shapes 2 and 3 (title is shape 1)
        var slide = _handler.Get("/slide[1]");
        var shapesBeforeGroup = slide.Children.Count(c => c.Type == "textbox" || c.Type == "title");
        shapesBeforeGroup.Should().BeGreaterThanOrEqualTo(3);

        // 2. Group shapes 2 and 3
        var groupPath = _handler.Add("/slide[1]", "group", null, new() { ["shapes"] = "2,3" });
        groupPath.Should().StartWith("/slide[1]/group[");

        // 3. Persist + Verify
        Reopen();
        var root = _handler.Get("/slide[1]");
        root.Should().NotBeNull();
    }

    // ==================== Chart Data Modification ====================

    [Fact]
    public void Chart_SetData_Lifecycle()
    {
        // 1. Add chart
        _handler.Add("/", "slide", null, new() { ["title"] = "Data Mod" });
        _handler.Add("/slide[1]", "chart", null, new()
        {
            ["charttype"] = "column",
            ["categories"] = "A,B",
            ["series1"] = "S1:10,20",
            ["series2"] = "S2:30,40"
        });

        // 2. Verify original data
        var node = _handler.Get("/slide[1]/chart[1]", 1);
        ((string)node.Children[0].Format["values"]).Should().Be("10,20");

        // 3. Set — modify series1 data
        _handler.Set("/slide[1]/chart[1]", new() { ["series1"] = "S1:100,200" });

        // 4. Get + Verify
        node = _handler.Get("/slide[1]/chart[1]", 1);
        ((string)node.Children[0].Format["values"]).Should().Be("100,200");
        node.Children[0].Text.Should().Be("S1");

        // 5. Set — modify all data at once via data property
        _handler.Set("/slide[1]/chart[1]", new() { ["data"] = "X:1,2;Y:3,4" });
        node = _handler.Get("/slide[1]/chart[1]", 1);
        ((string)node.Children[0].Format["values"]).Should().Be("1,2");
        ((string)node.Children[1].Format["values"]).Should().Be("3,4");

        // 6. Persist + Verify
        Reopen();
        node = _handler.Get("/slide[1]/chart[1]", 1);
        ((string)node.Children[0].Format["values"]).Should().Be("1,2");
    }

    [Fact]
    public void Chart_SetCategories_Lifecycle()
    {
        _handler.Add("/", "slide", null, new() { ["title"] = "Cat Mod" });
        _handler.Add("/slide[1]", "chart", null, new()
        {
            ["charttype"] = "column",
            ["categories"] = "A,B,C",
            ["series1"] = "D:1,2,3"
        });

        // Change categories
        _handler.Set("/slide[1]/chart[1]", new() { ["categories"] = "X,Y,Z" });

        var node = _handler.Get("/slide[1]/chart[1]");
        ((string)node.Format["categories"]).Should().Be("X,Y,Z");

        Reopen();
        node = _handler.Get("/slide[1]/chart[1]");
        ((string)node.Format["categories"]).Should().Be("X,Y,Z");
    }

    // ==================== Slide Size ====================

    [Fact]
    public void SlideSize_GetDefault()
    {
        // Blank doc should have default slide size
        var root = _handler.Get("/");
        root.Format.Should().ContainKey("slideWidth");
        root.Format.Should().ContainKey("slideHeight");
    }

    [Fact]
    public void SlideSize_SetPreset_Lifecycle()
    {
        // 1. Set to 4:3
        _handler.Set("/", new() { ["slidesize"] = "4:3" });

        // 2. Get + Verify
        var root = _handler.Get("/");
        ((string)root.Format["slideSize"]).Should().Be("screen4x3");

        // 3. Set to 16:9
        _handler.Set("/", new() { ["slidesize"] = "16:9" });
        root = _handler.Get("/");
        ((string)root.Format["slideSize"]).Should().Be("screen16x9");

        // 4. Persist + Verify
        Reopen();
        root = _handler.Get("/");
        ((string)root.Format["slideSize"]).Should().Be("screen16x9");
    }

    [Fact]
    public void SlideSize_SetCustom_Lifecycle()
    {
        // Set custom dimensions
        _handler.Set("/", new() { ["slidewidth"] = "30cm", ["slideheight"] = "20cm" });

        var root = _handler.Get("/");
        ((string)root.Format["slideWidth"]).Should().Contain("cm");
        ((string)root.Format["slideSize"]).Should().Be("custom");

        Reopen();
        root = _handler.Get("/");
        ((string)root.Format["slideSize"]).Should().Be("custom");
    }
}
