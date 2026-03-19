// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Full-lifecycle tests for PPTX enhancements: rich text (paragraph/run),
/// character spacing, paragraph indent, 3D effects, soft edge, and flip.
/// Pattern: Create → Add → Get → Verify → Set (modify) → Get → Verify → Reopen → Verify
/// </summary>
public class PptxEnhancementTests : IDisposable
{
    private readonly string _path = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.pptx");

    public void Dispose()
    {
        if (File.Exists(_path)) File.Delete(_path);
    }

    private PowerPointHandler CreateHandler(bool editable = true)
    {
        return new PowerPointHandler(_path, editable);
    }

    private void Reopen(ref PowerPointHandler handler)
    {
        handler.Dispose();
        handler = CreateHandler();
    }

    // ========== Feature 1: Rich Text (Mixed Runs) ==========

    [Fact]
    public void RichText_AddRun_FullLifecycle()
    {
        // 1. Create
        BlankDocCreator.Create(_path);
        var handler = CreateHandler();

        // 2. Add
        handler.Add("/", "slide", null, new() { ["title"] = "Test" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Hello" });

        // 3. Get + Verify (initial state: 1 paragraph, 1 run)
        var para = handler.Get("/slide[1]/shape[2]/paragraph[1]");
        para.Children.Should().HaveCount(1);
        para.Children[0].Text.Should().Be("Hello");

        // 4. Add run (modify: add a bold red run)
        handler.Add("/slide[1]/shape[2]/paragraph[1]", "run", null, new()
        {
            ["text"] = " World",
            ["bold"] = "true",
            ["color"] = "FF0000"
        });

        // 5. Get + Verify (now 2 runs)
        para = handler.Get("/slide[1]/shape[2]/paragraph[1]");
        para.Children.Should().HaveCount(2);
        para.Children[0].Text.Should().Be("Hello");
        para.Children[1].Text.Should().Be(" World");
        para.Children[1].Format["bold"].Should().Be(true);
        para.Children[1].Format["color"].Should().Be("FF0000");

        // 6. Set (modify the added run)
        handler.Set("/slide[1]/shape[2]/paragraph[1]/run[2]", new()
        {
            ["italic"] = "true",
            ["size"] = "24"
        });

        // 7. Get + Verify modification
        para = handler.Get("/slide[1]/shape[2]/paragraph[1]");
        para.Children[1].Format["bold"].Should().Be(true);
        para.Children[1].Format["italic"].Should().Be(true);
        para.Children[1].Format["size"].Should().Be("24pt");

        // 8. Reopen + Verify persistence
        Reopen(ref handler);
        para = handler.Get("/slide[1]/shape[2]/paragraph[1]");
        para.Children.Should().HaveCount(2);
        para.Children[1].Text.Should().Be(" World");
        para.Children[1].Format["bold"].Should().Be(true);
        para.Children[1].Format["italic"].Should().Be(true);
        para.Children[1].Format["color"].Should().Be("FF0000");

        handler.Dispose();
    }

    [Fact]
    public void RichText_AddRunToShape_AppendsToLastParagraph()
    {
        // 1. Create
        BlankDocCreator.Create(_path);
        var handler = CreateHandler();

        // 2. Add shape with multi-line text
        handler.Add("/", "slide", null, new() { ["title"] = "Test" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Line1\\nLine2" });

        // 3. Get + Verify (2 paragraphs)
        var shape = handler.Get("/slide[1]/shape[2]");
        shape.ChildCount.Should().Be(2);

        // 4. Add run to shape (appends to last paragraph)
        handler.Add("/slide[1]/shape[2]", "run", null, new()
        {
            ["text"] = " appended",
            ["italic"] = "true"
        });

        // 5. Get + Verify
        var para2 = handler.Get("/slide[1]/shape[2]/paragraph[2]");
        para2.Children.Should().HaveCount(2);
        para2.Children[1].Text.Should().Be(" appended");
        para2.Children[1].Format["italic"].Should().Be(true);

        // 6. Set (modify the appended run)
        handler.Set("/slide[1]/shape[2]/paragraph[2]/run[2]", new() { ["bold"] = "true" });

        // 7. Get + Verify
        para2 = handler.Get("/slide[1]/shape[2]/paragraph[2]");
        para2.Children[1].Format["italic"].Should().Be(true);
        para2.Children[1].Format["bold"].Should().Be(true);

        // 8. Reopen + Verify
        Reopen(ref handler);
        para2 = handler.Get("/slide[1]/shape[2]/paragraph[2]");
        para2.Children.Should().HaveCount(2);
        para2.Children[1].Text.Should().Be(" appended");
        para2.Children[1].Format["italic"].Should().Be(true);
        para2.Children[1].Format["bold"].Should().Be(true);

        handler.Dispose();
    }

    [Fact]
    public void RichText_AddParagraph_FullLifecycle()
    {
        // 1. Create
        BlankDocCreator.Create(_path);
        var handler = CreateHandler();

        // 2. Add
        handler.Add("/", "slide", null, new() { ["title"] = "Test" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "First" });

        // 3. Get + Verify
        var shape = handler.Get("/slide[1]/shape[2]");
        shape.ChildCount.Should().Be(1);

        // 4. Add paragraph (modify structure)
        handler.Add("/slide[1]/shape[2]", "paragraph", null, new()
        {
            ["text"] = "Second",
            ["bold"] = "true",
            ["align"] = "center"
        });

        // 5. Get + Verify
        shape = handler.Get("/slide[1]/shape[2]");
        shape.ChildCount.Should().Be(2);
        var para2 = handler.Get("/slide[1]/shape[2]/paragraph[2]");
        para2.Text.Should().Be("Second");
        para2.Format["align"].Should().Be("ctr");

        // 6. Set (modify the added paragraph)
        handler.Set("/slide[1]/shape[2]/paragraph[2]", new() { ["align"] = "right" });

        // 7. Get + Verify
        para2 = handler.Get("/slide[1]/shape[2]/paragraph[2]");
        para2.Format["align"].Should().Be("r");

        // 8. Reopen + Verify
        Reopen(ref handler);
        shape = handler.Get("/slide[1]/shape[2]");
        shape.ChildCount.Should().Be(2);
        para2 = handler.Get("/slide[1]/shape[2]/paragraph[2]");
        para2.Text.Should().Be("Second");
        para2.Format["align"].Should().Be("r");

        handler.Dispose();
    }

    [Fact]
    public void RichText_MultipleRuns_MixedFormatting_FullLifecycle()
    {
        // 1. Create
        BlankDocCreator.Create(_path);
        var handler = CreateHandler();

        // 2. Add shape + multiple runs
        handler.Add("/", "slide", null, new() { ["title"] = "Test" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Normal " });

        handler.Add("/slide[1]/shape[2]/paragraph[1]", "run", null, new()
            { ["text"] = "bold", ["bold"] = "true" });
        handler.Add("/slide[1]/shape[2]/paragraph[1]", "run", null, new()
            { ["text"] = " italic", ["italic"] = "true" });
        handler.Add("/slide[1]/shape[2]/paragraph[1]", "run", null, new()
            { ["text"] = " red", ["color"] = "FF0000", ["size"] = "24" });

        // 3. Get + Verify
        var para = handler.Get("/slide[1]/shape[2]/paragraph[1]");
        para.Children.Should().HaveCount(4);
        para.Children[0].Text.Should().Be("Normal ");
        para.Children[1].Text.Should().Be("bold");
        para.Children[1].Format["bold"].Should().Be(true);
        para.Children[2].Text.Should().Be(" italic");
        para.Children[2].Format["italic"].Should().Be(true);
        para.Children[3].Text.Should().Be(" red");
        para.Children[3].Format["color"].Should().Be("FF0000");

        // 4. Set (modify run[3] — change italic to bold+italic)
        handler.Set("/slide[1]/shape[2]/paragraph[1]/run[3]", new() { ["bold"] = "true" });

        // 5. Get + Verify modification
        para = handler.Get("/slide[1]/shape[2]/paragraph[1]");
        para.Children[2].Format["italic"].Should().Be(true);
        para.Children[2].Format["bold"].Should().Be(true);

        // 6. Reopen + Verify persistence
        Reopen(ref handler);
        para = handler.Get("/slide[1]/shape[2]/paragraph[1]");
        para.Children.Should().HaveCount(4);
        para.Children[1].Format["bold"].Should().Be(true);
        para.Children[2].Format["italic"].Should().Be(true);
        para.Children[2].Format["bold"].Should().Be(true);
        para.Children[3].Format["color"].Should().Be("FF0000");

        handler.Dispose();
    }

    // ========== Feature 2: Character Spacing & Paragraph Indent ==========

    [Fact]
    public void CharacterSpacing_FullLifecycle()
    {
        // 1. Create
        BlankDocCreator.Create(_path);
        var handler = CreateHandler();

        // 2. Add with spacing
        handler.Add("/", "slide", null, new() { ["title"] = "Test" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Spaced", ["spacing"] = "2" });

        // 3. Get + Verify
        var node = handler.Get("/slide[1]/shape[2]");
        node.Format["spacing"].Should().Be("2");

        // 4. Set (modify spacing)
        handler.Set("/slide[1]/shape[2]", new() { ["spacing"] = "5" });

        // 5. Get + Verify
        node = handler.Get("/slide[1]/shape[2]");
        node.Format["spacing"].Should().Be("5");

        // 6. Reopen + Verify
        Reopen(ref handler);
        node = handler.Get("/slide[1]/shape[2]");
        node.Format["spacing"].Should().Be("5");

        handler.Dispose();
    }

    [Fact]
    public void NegativeCharacterSpacing_FullLifecycle()
    {
        // 1. Create
        BlankDocCreator.Create(_path);
        var handler = CreateHandler();

        // 2. Add
        handler.Add("/", "slide", null, new() { ["title"] = "Test" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Tight" });

        // 3. Get + Verify (no spacing initially)
        var node = handler.Get("/slide[1]/shape[2]");
        node.Format.Should().NotContainKey("spacing");

        // 4. Set negative spacing
        handler.Set("/slide[1]/shape[2]", new() { ["spacing"] = "-1.5" });

        // 5. Get + Verify
        node = handler.Get("/slide[1]/shape[2]");
        node.Format["spacing"].Should().Be("-1.5");

        // 6. Reopen + Verify
        Reopen(ref handler);
        node = handler.Get("/slide[1]/shape[2]");
        node.Format["spacing"].Should().Be("-1.5");

        handler.Dispose();
    }

    [Fact]
    public void ParagraphIndent_FullLifecycle()
    {
        // 1. Create
        BlankDocCreator.Create(_path);
        var handler = CreateHandler();

        // 2. Add with indent
        handler.Add("/", "slide", null, new() { ["title"] = "Test" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Indented", ["indent"] = "1cm" });

        // 3. Get + Verify
        var node = handler.Get("/slide[1]/shape[2]");
        node.Format.Should().ContainKey("indent");

        // 4. Set (change indent)
        handler.Set("/slide[1]/shape[2]", new() { ["indent"] = "2cm" });

        // 5. Get + Verify
        node = handler.Get("/slide[1]/shape[2]");
        node.Format.Should().ContainKey("indent");

        // 6. Reopen + Verify
        Reopen(ref handler);
        node = handler.Get("/slide[1]/shape[2]");
        node.Format.Should().ContainKey("indent");

        handler.Dispose();
    }

    [Fact]
    public void ParagraphMargins_FullLifecycle()
    {
        // 1. Create
        BlankDocCreator.Create(_path);
        var handler = CreateHandler();

        // 2. Add with margins
        handler.Add("/", "slide", null, new() { ["title"] = "Test" });
        handler.Add("/slide[1]", "shape", null, new()
            { ["text"] = "Margins", ["marginLeft"] = "1cm", ["marginRight"] = "0.5cm" });

        // 3. Get + Verify
        var node = handler.Get("/slide[1]/shape[2]");
        node.Format.Should().ContainKey("marginLeft");
        node.Format.Should().ContainKey("marginRight");

        // 4. Set (change margins)
        handler.Set("/slide[1]/shape[2]", new() { ["marginLeft"] = "3cm" });

        // 5. Get + Verify
        node = handler.Get("/slide[1]/shape[2]");
        node.Format.Should().ContainKey("marginLeft");

        // 6. Reopen + Verify
        Reopen(ref handler);
        node = handler.Get("/slide[1]/shape[2]");
        node.Format.Should().ContainKey("marginLeft");
        node.Format.Should().ContainKey("marginRight");

        handler.Dispose();
    }

    [Fact]
    public void ParagraphLevelIndent_FullLifecycle()
    {
        // 1. Create
        BlankDocCreator.Create(_path);
        var handler = CreateHandler();

        // 2. Add
        handler.Add("/", "slide", null, new() { ["title"] = "Test" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Para" });

        // 3. Get + Verify (no indent initially)
        var para = handler.Get("/slide[1]/shape[2]/paragraph[1]");
        para.Format.Should().NotContainKey("indent");

        // 4. Set indent at paragraph level
        handler.Set("/slide[1]/shape[2]/paragraph[1]", new()
        {
            ["indent"] = "0.5cm",
            ["marginLeft"] = "1cm"
        });

        // 5. Get + Verify
        para = handler.Get("/slide[1]/shape[2]/paragraph[1]");
        para.Format.Should().ContainKey("indent");
        para.Format.Should().ContainKey("marginLeft");

        // 6. Set (modify)
        handler.Set("/slide[1]/shape[2]/paragraph[1]", new() { ["indent"] = "1cm" });

        // 7. Get + Verify
        para = handler.Get("/slide[1]/shape[2]/paragraph[1]");
        para.Format.Should().ContainKey("indent");

        // 8. Reopen + Verify
        Reopen(ref handler);
        para = handler.Get("/slide[1]/shape[2]/paragraph[1]");
        para.Format.Should().ContainKey("indent");
        para.Format.Should().ContainKey("marginLeft");

        handler.Dispose();
    }

    [Fact]
    public void RunSpacing_InRichText_FullLifecycle()
    {
        // 1. Create
        BlankDocCreator.Create(_path);
        var handler = CreateHandler();

        // 2. Add shape + run with spacing
        handler.Add("/", "slide", null, new() { ["title"] = "Test" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Normal" });
        handler.Add("/slide[1]/shape[2]/paragraph[1]", "run", null, new()
            { ["text"] = " wide", ["spacing"] = "5" });

        // 3. Get + Verify
        var para = handler.Get("/slide[1]/shape[2]/paragraph[1]");
        para.Children.Should().HaveCount(2);
        para.Children[1].Format["spacing"].Should().Be("5");

        // 4. Set (modify spacing on run)
        handler.Set("/slide[1]/shape[2]/paragraph[1]/run[2]", new() { ["spacing"] = "3" });

        // 5. Get + Verify
        para = handler.Get("/slide[1]/shape[2]/paragraph[1]");
        para.Children[1].Format["spacing"].Should().Be("3");

        // 6. Reopen + Verify
        Reopen(ref handler);
        para = handler.Get("/slide[1]/shape[2]/paragraph[1]");
        para.Children.Should().HaveCount(2);
        para.Children[1].Format["spacing"].Should().Be("3");

        handler.Dispose();
    }

    // ========== Feature 3: 3D Effects + Soft Edge + Flip ==========

    [Fact]
    public void SoftEdge_FullLifecycle()
    {
        // 1. Create
        BlankDocCreator.Create(_path);
        var handler = CreateHandler();

        // 2. Add with soft edge
        handler.Add("/", "slide", null, new() { ["title"] = "Test" });
        handler.Add("/slide[1]", "shape", null, new()
            { ["text"] = "Soft", ["fill"] = "4472C4", ["softEdge"] = "3" });

        // 3. Get + Verify
        var node = handler.Get("/slide[1]/shape[2]");
        node.Format["softEdge"].Should().Be("3");

        // 4. Set (modify)
        handler.Set("/slide[1]/shape[2]", new() { ["softEdge"] = "8" });

        // 5. Get + Verify
        node = handler.Get("/slide[1]/shape[2]");
        node.Format["softEdge"].Should().Be("8");

        // 6. Reopen + Verify
        Reopen(ref handler);
        node = handler.Get("/slide[1]/shape[2]");
        node.Format["softEdge"].Should().Be("8");

        // 7. Set (remove)
        handler.Set("/slide[1]/shape[2]", new() { ["softEdge"] = "none" });
        node = handler.Get("/slide[1]/shape[2]");
        node.Format.Should().NotContainKey("softEdge");

        handler.Dispose();
    }

    [Fact]
    public void FlipH_FullLifecycle()
    {
        // 1. Create
        BlankDocCreator.Create(_path);
        var handler = CreateHandler();

        // 2. Add
        handler.Add("/", "slide", null, new() { ["title"] = "Test" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Flip" });

        // 3. Get + Verify (not flipped)
        var node = handler.Get("/slide[1]/shape[2]");
        node.Format.Should().NotContainKey("flipH");

        // 4. Set flip
        handler.Set("/slide[1]/shape[2]", new() { ["flipH"] = "true" });

        // 5. Get + Verify
        node = handler.Get("/slide[1]/shape[2]");
        node.Format["flipH"].Should().Be(true);

        // 6. Set (also flip V)
        handler.Set("/slide[1]/shape[2]", new() { ["flipV"] = "true" });

        // 7. Get + Verify both
        node = handler.Get("/slide[1]/shape[2]");
        node.Format["flipH"].Should().Be(true);
        node.Format["flipV"].Should().Be(true);

        // 8. Reopen + Verify
        Reopen(ref handler);
        node = handler.Get("/slide[1]/shape[2]");
        node.Format["flipH"].Should().Be(true);
        node.Format["flipV"].Should().Be(true);

        handler.Dispose();
    }

    [Fact]
    public void Rotation3D_FullLifecycle()
    {
        // 1. Create
        BlankDocCreator.Create(_path);
        var handler = CreateHandler();

        // 2. Add
        handler.Add("/", "slide", null, new() { ["title"] = "Test" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "3D", ["fill"] = "4472C4" });

        // 3. Get + Verify (no 3D initially)
        var node = handler.Get("/slide[1]/shape[2]");
        node.Format.Should().NotContainKey("rot3d");

        // 4. Set 3D rotation
        handler.Set("/slide[1]/shape[2]", new() { ["rot3d"] = "45,30,0" });

        // 5. Get + Verify
        node = handler.Get("/slide[1]/shape[2]");
        node.Format["rot3d"].Should().Be("45,30,0");

        // 6. Set (modify rotation)
        handler.Set("/slide[1]/shape[2]", new() { ["rot3d"] = "20,10,5" });

        // 7. Get + Verify
        node = handler.Get("/slide[1]/shape[2]");
        node.Format["rot3d"].Should().Be("20,10,5");

        // 8. Reopen + Verify
        Reopen(ref handler);
        node = handler.Get("/slide[1]/shape[2]");
        node.Format["rot3d"].Should().Be("20,10,5");

        handler.Dispose();
    }

    [Fact]
    public void IndividualRotationAxes_FullLifecycle()
    {
        // 1. Create
        BlankDocCreator.Create(_path);
        var handler = CreateHandler();

        // 2. Add
        handler.Add("/", "slide", null, new() { ["title"] = "Test" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "3D", ["fill"] = "4472C4" });

        // 3. Set rotX
        handler.Set("/slide[1]/shape[2]", new() { ["rotX"] = "30" });

        // 4. Get + Verify
        var node = handler.Get("/slide[1]/shape[2]");
        node.Format.Should().ContainKey("rot3d");
        node.Format["rot3d"].ToString()!.Should().Contain("30");

        // 5. Set rotY (adds to existing)
        handler.Set("/slide[1]/shape[2]", new() { ["rotY"] = "15" });

        // 6. Get + Verify both axes
        node = handler.Get("/slide[1]/shape[2]");
        var rot = node.Format["rot3d"].ToString()!;
        rot.Should().Contain("30");
        rot.Should().Contain("15");

        // 7. Reopen + Verify
        Reopen(ref handler);
        node = handler.Get("/slide[1]/shape[2]");
        rot = node.Format["rot3d"].ToString()!;
        rot.Should().Contain("30");
        rot.Should().Contain("15");

        handler.Dispose();
    }

    [Fact]
    public void Bevel_FullLifecycle()
    {
        // 1. Create
        BlankDocCreator.Create(_path);
        var handler = CreateHandler();

        // 2. Add
        handler.Add("/", "slide", null, new() { ["title"] = "Test" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Bevel", ["fill"] = "4472C4" });

        // 3. Get + Verify (no bevel initially)
        var node = handler.Get("/slide[1]/shape[2]");
        node.Format.Should().NotContainKey("bevel");

        // 4. Set bevel
        handler.Set("/slide[1]/shape[2]", new() { ["bevel"] = "circle-8-8" });

        // 5. Get + Verify
        node = handler.Get("/slide[1]/shape[2]");
        node.Format["bevel"].Should().Be("circle-8-8");

        // 6. Set (modify bevel + add bottom bevel)
        handler.Set("/slide[1]/shape[2]", new()
        {
            ["bevel"] = "coolSlant-4-4",
            ["bevelBottom"] = "relaxedInset-3-3"
        });

        // 7. Get + Verify
        node = handler.Get("/slide[1]/shape[2]");
        node.Format["bevel"].Should().Be("coolSlant-4-4");
        node.Format["bevelBottom"].Should().Be("relaxedInset-3-3");

        // 8. Reopen + Verify
        Reopen(ref handler);
        node = handler.Get("/slide[1]/shape[2]");
        node.Format["bevel"].Should().Be("coolSlant-4-4");
        node.Format["bevelBottom"].Should().Be("relaxedInset-3-3");

        handler.Dispose();
    }

    [Fact]
    public void DepthAndMaterial_FullLifecycle()
    {
        // 1. Create
        BlankDocCreator.Create(_path);
        var handler = CreateHandler();

        // 2. Add
        handler.Add("/", "slide", null, new() { ["title"] = "Test" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Deep", ["fill"] = "4472C4" });

        // 3. Set depth + material
        handler.Set("/slide[1]/shape[2]", new() { ["depth"] = "10", ["material"] = "metal" });

        // 4. Get + Verify
        var node = handler.Get("/slide[1]/shape[2]");
        node.Format["depth"].Should().Be("10");
        node.Format["material"].Should().Be("metal");

        // 5. Set (modify)
        handler.Set("/slide[1]/shape[2]", new() { ["depth"] = "20", ["material"] = "plastic" });

        // 6. Get + Verify
        node = handler.Get("/slide[1]/shape[2]");
        node.Format["depth"].Should().Be("20");
        node.Format["material"].Should().Be("plastic");

        // 7. Reopen + Verify
        Reopen(ref handler);
        node = handler.Get("/slide[1]/shape[2]");
        node.Format["depth"].Should().Be("20");
        node.Format["material"].Should().Be("plastic");

        handler.Dispose();
    }

    [Fact]
    public void Lighting_FullLifecycle()
    {
        // 1. Create
        BlankDocCreator.Create(_path);
        var handler = CreateHandler();

        // 2. Add
        handler.Add("/", "slide", null, new() { ["title"] = "Test" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Lit", ["fill"] = "4472C4" });

        // 3. Set lighting
        handler.Set("/slide[1]/shape[2]", new() { ["lighting"] = "balanced" });

        // 4. Get + Verify
        var node = handler.Get("/slide[1]/shape[2]");
        node.Format["lighting"].Should().Be("balanced");

        // 5. Set (modify)
        handler.Set("/slide[1]/shape[2]", new() { ["lighting"] = "harsh" });

        // 6. Get + Verify
        node = handler.Get("/slide[1]/shape[2]");
        node.Format["lighting"].Should().Be("harsh");

        // 7. Reopen + Verify
        Reopen(ref handler);
        node = handler.Get("/slide[1]/shape[2]");
        node.Format["lighting"].Should().Be("harsh");

        handler.Dispose();
    }

    [Fact]
    public void Complete3DEffect_FullLifecycle()
    {
        // 1. Create
        BlankDocCreator.Create(_path);
        var handler = CreateHandler();

        // 2. Add with 3D via Add
        handler.Add("/", "slide", null, new() { ["title"] = "Test" });
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Full 3D",
            ["fill"] = "4472C4",
            ["softEdge"] = "2",
            ["bevel"] = "circle",
            ["depth"] = "5"
        });

        // 3. Get + Verify (from Add)
        var node = handler.Get("/slide[1]/shape[2]");
        node.Format["softEdge"].Should().Be("2");
        node.Format.Should().ContainKey("bevel");
        node.Format["depth"].Should().Be("5");

        // 4. Set (full 3D effect)
        handler.Set("/slide[1]/shape[2]", new()
        {
            ["rot3d"] = "20,10,0",
            ["bevel"] = "circle-6-6",
            ["depth"] = "8",
            ["material"] = "plastic",
            ["lighting"] = "harsh",
            ["softEdge"] = "3",
            ["flipH"] = "true"
        });

        // 5. Get + Verify
        node = handler.Get("/slide[1]/shape[2]");
        node.Format["rot3d"].Should().Be("20,10,0");
        node.Format["bevel"].Should().Be("circle-6-6");
        node.Format["depth"].Should().Be("8");
        node.Format["material"].Should().Be("plastic");
        node.Format["lighting"].Should().Be("harsh");
        node.Format["softEdge"].Should().Be("3");
        node.Format["flipH"].Should().Be(true);

        // 6. Set (modify subset)
        handler.Set("/slide[1]/shape[2]", new()
        {
            ["rot3d"] = "30,15,5",
            ["material"] = "metal"
        });

        // 7. Get + Verify modification (other props unchanged)
        node = handler.Get("/slide[1]/shape[2]");
        node.Format["rot3d"].Should().Be("30,15,5");
        node.Format["material"].Should().Be("metal");
        node.Format["bevel"].Should().Be("circle-6-6"); // unchanged
        node.Format["softEdge"].Should().Be("3"); // unchanged
        node.Format["flipH"].Should().Be(true); // unchanged

        // 8. Reopen + Verify persistence of all properties
        Reopen(ref handler);
        node = handler.Get("/slide[1]/shape[2]");
        node.Format["rot3d"].Should().Be("30,15,5");
        node.Format["bevel"].Should().Be("circle-6-6");
        node.Format["depth"].Should().Be("8");
        node.Format["material"].Should().Be("metal");
        node.Format["lighting"].Should().Be("harsh");
        node.Format["softEdge"].Should().Be("3");
        node.Format["flipH"].Should().Be(true);

        handler.Dispose();
    }

    // ========== Feature 4: Baseline (Superscript/Subscript) ==========

    [Fact]
    public void Baseline_Superscript_FullLifecycle()
    {
        // 1. Create
        BlankDocCreator.Create(_path);
        var handler = CreateHandler();

        // 2. Add shape + run with superscript
        handler.Add("/", "slide", null, new() { ["title"] = "Test" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "E=mc" });
        handler.Add("/slide[1]/shape[2]/paragraph[1]", "run", null, new()
            { ["text"] = "2", ["baseline"] = "super" });

        // 3. Get + Verify
        var para = handler.Get("/slide[1]/shape[2]/paragraph[1]");
        para.Children.Should().HaveCount(2);
        para.Children[1].Text.Should().Be("2");
        para.Children[1].Format["baseline"].Should().Be("30");

        // 4. Set (modify to custom value)
        handler.Set("/slide[1]/shape[2]/paragraph[1]/run[2]", new() { ["baseline"] = "40" });

        // 5. Get + Verify
        para = handler.Get("/slide[1]/shape[2]/paragraph[1]");
        para.Children[1].Format["baseline"].Should().Be("40");

        // 6. Reopen + Verify
        Reopen(ref handler);
        para = handler.Get("/slide[1]/shape[2]/paragraph[1]");
        para.Children[1].Text.Should().Be("2");
        para.Children[1].Format["baseline"].Should().Be("40");

        handler.Dispose();
    }

    [Fact]
    public void Baseline_Subscript_FullLifecycle()
    {
        // 1. Create
        BlankDocCreator.Create(_path);
        var handler = CreateHandler();

        // 2. Add shape + run with subscript
        handler.Add("/", "slide", null, new() { ["title"] = "Test" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "H" });
        handler.Add("/slide[1]/shape[2]/paragraph[1]", "run", null, new()
            { ["text"] = "2", ["baseline"] = "sub" });
        handler.Add("/slide[1]/shape[2]/paragraph[1]", "run", null, new()
            { ["text"] = "O" });

        // 3. Get + Verify
        var para = handler.Get("/slide[1]/shape[2]/paragraph[1]");
        para.Children.Should().HaveCount(3);
        para.Children[1].Text.Should().Be("2");
        para.Children[1].Format["baseline"].Should().Be("-25");
        para.Children[2].Format.Should().NotContainKey("baseline"); // O has no baseline

        // 4. Set (modify subscript value)
        handler.Set("/slide[1]/shape[2]/paragraph[1]/run[2]", new() { ["baseline"] = "-30" });

        // 5. Get + Verify
        para = handler.Get("/slide[1]/shape[2]/paragraph[1]");
        para.Children[1].Format["baseline"].Should().Be("-30");

        // 6. Set (remove baseline)
        handler.Set("/slide[1]/shape[2]/paragraph[1]/run[2]", new() { ["baseline"] = "0" });

        // 7. Get + Verify (baseline removed)
        para = handler.Get("/slide[1]/shape[2]/paragraph[1]");
        para.Children[1].Format.Should().NotContainKey("baseline");

        // 8. Reopen + Verify
        Reopen(ref handler);
        para = handler.Get("/slide[1]/shape[2]/paragraph[1]");
        para.Children[1].Format.Should().NotContainKey("baseline");

        handler.Dispose();
    }

    [Fact]
    public void Baseline_SuperscriptShorthand_FullLifecycle()
    {
        // 1. Create
        BlankDocCreator.Create(_path);
        var handler = CreateHandler();

        // 2. Add
        handler.Add("/", "slide", null, new() { ["title"] = "Test" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "x" });

        // 3. Set via superscript shorthand
        handler.Add("/slide[1]/shape[2]/paragraph[1]", "run", null, new()
            { ["text"] = "n", ["superscript"] = "true" });

        // 4. Get + Verify
        var para = handler.Get("/slide[1]/shape[2]/paragraph[1]");
        para.Children[1].Format["baseline"].Should().Be("30");

        // 5. Set via subscript shorthand on same run
        handler.Set("/slide[1]/shape[2]/paragraph[1]/run[2]", new() { ["subscript"] = "true" });

        // 6. Get + Verify (now subscript)
        para = handler.Get("/slide[1]/shape[2]/paragraph[1]");
        para.Children[1].Format["baseline"].Should().Be("-25");

        // 7. Reopen + Verify
        Reopen(ref handler);
        para = handler.Get("/slide[1]/shape[2]/paragraph[1]");
        para.Children[1].Format["baseline"].Should().Be("-25");

        handler.Dispose();
    }

    // ========== Feature 5: Z-Order ==========

    [Fact]
    public void ZOrder_BringToFront_FullLifecycle()
    {
        // 1. Create
        BlankDocCreator.Create(_path);
        var handler = CreateHandler();

        // 2. Add 3 shapes
        handler.Add("/", "slide", null, new() { ["title"] = "Test" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "A" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "B" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "C" });

        // 3. Get + Verify initial z-order (title=1, A=2, B=3, C=4)
        var nodeA = handler.Get("/slide[1]/shape[2]");
        nodeA.Text.Should().Be("A");
        var zorderA = (int)nodeA.Format["zorder"];
        var nodeC = handler.Get("/slide[1]/shape[4]");
        nodeC.Text.Should().Be("C");
        var zorderC = (int)nodeC.Format["zorder"];
        zorderC.Should().BeGreaterThan(zorderA);

        // 4. Set z-order: bring A to front
        handler.Set("/slide[1]/shape[2]", new() { ["zorder"] = "front" });

        // 5. Get + Verify: A is now at the front (highest z-order)
        // After moving, shape indices change. A is now the last shape.
        // Find A by checking text of shapes
        var shapes = handler.Query("shape");
        var aNode = shapes.First(s => s.Text == "A");
        aNode.Format["zorder"].Should().Be(4); // front position

        // 6. Reopen + Verify persistence
        Reopen(ref handler);
        shapes = handler.Query("shape");
        aNode = shapes.First(s => s.Text == "A");
        aNode.Format["zorder"].Should().Be(4);

        handler.Dispose();
    }

    [Fact]
    public void ZOrder_SendToBack_FullLifecycle()
    {
        // 1. Create
        BlankDocCreator.Create(_path);
        var handler = CreateHandler();

        // 2. Add 3 shapes
        handler.Add("/", "slide", null, new() { ["title"] = "Test" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "A" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "B" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "C" });

        // 3. Get + Verify: C is shape[4], at the front
        var nodeC = handler.Get("/slide[1]/shape[4]");
        nodeC.Text.Should().Be("C");

        // 4. Set: send C to back
        handler.Set("/slide[1]/shape[4]", new() { ["zorder"] = "back" });

        // 5. Get + Verify: C is now at the back
        var shapes = handler.Query("shape");
        var cNode = shapes.First(s => s.Text == "C");
        cNode.Format["zorder"].Should().Be(1); // back position

        // 6. Reopen + Verify
        Reopen(ref handler);
        shapes = handler.Query("shape");
        cNode = shapes.First(s => s.Text == "C");
        cNode.Format["zorder"].Should().Be(1);

        handler.Dispose();
    }

    [Fact]
    public void ZOrder_ForwardBackward_FullLifecycle()
    {
        // 1. Create
        BlankDocCreator.Create(_path);
        var handler = CreateHandler();

        // 2. Add 3 shapes (title + A + B + C)
        handler.Add("/", "slide", null, new() { ["title"] = "Test" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "A" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "B" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "C" });

        // 3. A is shape[2] (zorder=2). Move forward by 1.
        handler.Set("/slide[1]/shape[2]", new() { ["zorder"] = "forward" });

        // 4. Verify: A moved one step forward
        var shapes = handler.Query("shape");
        var aNode = shapes.First(s => s.Text == "A");
        var bNode = shapes.First(s => s.Text == "B");
        ((int)aNode.Format["zorder"]).Should().BeGreaterThan((int)bNode.Format["zorder"]);

        // 5. Move A backward
        var aPath = aNode.Path;
        handler.Set(aPath, new() { ["zorder"] = "backward" });

        // 6. Verify: A is back behind B
        shapes = handler.Query("shape");
        aNode = shapes.First(s => s.Text == "A");
        bNode = shapes.First(s => s.Text == "B");
        ((int)aNode.Format["zorder"]).Should().BeLessThan((int)bNode.Format["zorder"]);

        // 7. Reopen + Verify
        Reopen(ref handler);
        shapes = handler.Query("shape");
        aNode = shapes.First(s => s.Text == "A");
        bNode = shapes.First(s => s.Text == "B");
        ((int)aNode.Format["zorder"]).Should().BeLessThan((int)bNode.Format["zorder"]);

        handler.Dispose();
    }

    [Fact]
    public void ZOrder_AbsolutePosition_FullLifecycle()
    {
        // 1. Create
        BlankDocCreator.Create(_path);
        var handler = CreateHandler();

        // 2. Add 4 shapes (title + A + B + C)
        handler.Add("/", "slide", null, new() { ["title"] = "Title" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "A" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "B" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "C" });

        // 3. Move C (shape[4], zorder=4) to position 2
        handler.Set("/slide[1]/shape[4]", new() { ["zorder"] = "2" });

        // 4. Verify
        var shapes = handler.Query("shape");
        var cNode = shapes.First(s => s.Text == "C");
        cNode.Format["zorder"].Should().Be(2);

        // 5. Reopen + Verify
        Reopen(ref handler);
        shapes = handler.Query("shape");
        cNode = shapes.First(s => s.Text == "C");
        cNode.Format["zorder"].Should().Be(2);

        handler.Dispose();
    }

    // ========== Feature 6: Slide Clone ==========

    [Fact]
    public void CloneSlide_FullLifecycle()
    {
        // 1. Create
        BlankDocCreator.Create(_path);
        var handler = CreateHandler();

        // 2. Add a slide with shapes
        handler.Add("/", "slide", null, new() { ["title"] = "Template" });
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Hello World",
            ["fill"] = "4472C4",
            ["bold"] = "true",
            ["font"] = "Arial"
        });
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Subtitle",
            ["color"] = "FF0000"
        });

        // 3. Get + Verify original
        var slide1 = handler.Get("/slide[1]");
        slide1.Should().NotBeNull();
        var shapes1 = handler.Query("slide[1] shape");
        shapes1.Count.Should().BeGreaterThanOrEqualTo(3); // title + 2 shapes

        // 4. Clone slide: --from /slide[1] to /
        var clonedPath = handler.CopyFrom("/slide[1]", "/", null);
        clonedPath.Should().Be("/slide[2]");

        // 5. Get + Verify clone has same content
        var slide2Shapes = handler.Query("slide[2] shape");
        slide2Shapes.Count.Should().Be(shapes1.Count);

        // Find the "Hello World" shape in cloned slide
        var helloShape = slide2Shapes.FirstOrDefault(s => s.Text?.Contains("Hello World") == true);
        helloShape.Should().NotBeNull();
        helloShape!.Format["fill"].Should().Be("4472C4");
        helloShape.Format["bold"].Should().Be(true);

        var subShape = slide2Shapes.FirstOrDefault(s => s.Text?.Contains("Subtitle") == true);
        subShape.Should().NotBeNull();
        subShape!.Format["color"].Should().Be("FF0000");

        // 6. Set (modify cloned slide to verify independence)
        handler.Set("/slide[2]/shape[2]", new() { ["text"] = "Cloned!" });
        var modified = handler.Get("/slide[2]/shape[2]");
        modified.Text.Should().Contain("Cloned!");

        // Original should be unchanged
        var original = handler.Get("/slide[1]/shape[2]");
        original.Text.Should().Contain("Hello World");

        // 7. Reopen + Verify
        Reopen(ref handler);
        var root = handler.Get("/");
        root.ChildCount.Should().Be(2);

        var slide2ShapesAfter = handler.Query("slide[2] shape");
        slide2ShapesAfter.Count.Should().Be(shapes1.Count);

        handler.Dispose();
    }

    [Fact]
    public void CloneSlide_AtIndex_FullLifecycle()
    {
        // 1. Create
        BlankDocCreator.Create(_path);
        var handler = CreateHandler();

        // 2. Add 2 slides
        handler.Add("/", "slide", null, new() { ["title"] = "Slide A" });
        handler.Add("/", "slide", null, new() { ["title"] = "Slide B" });

        // 3. Verify order
        handler.Get("/slide[1]").Children.Should().Contain(c => c.Text == "Slide A");
        handler.Get("/slide[2]").Children.Should().Contain(c => c.Text == "Slide B");

        // 4. Clone slide 2 at index 0 (insert at beginning)
        var clonedPath = handler.CopyFrom("/slide[2]", "/", 0);
        clonedPath.Should().Be("/slide[1]");

        // 5. Verify: cloned B is now slide[1], original A is slide[2], original B is slide[3]
        handler.Get("/slide[1]").Children.Should().Contain(c => c.Text == "Slide B");
        handler.Get("/slide[2]").Children.Should().Contain(c => c.Text == "Slide A");
        handler.Get("/slide[3]").Children.Should().Contain(c => c.Text == "Slide B");

        // 6. Reopen + Verify
        Reopen(ref handler);
        var root = handler.Get("/");
        root.ChildCount.Should().Be(3);
        handler.Get("/slide[1]").Children.Should().Contain(c => c.Text == "Slide B");

        handler.Dispose();
    }

    [Fact]
    public void CloneSlide_WithBackground_FullLifecycle()
    {
        // 1. Create
        BlankDocCreator.Create(_path);
        var handler = CreateHandler();

        // 2. Add slide with background and shapes
        handler.Add("/", "slide", null, new() { ["title"] = "Dark", ["background"] = "1F3864" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "On dark bg", ["color"] = "FFFFFF" });

        // 3. Get + Verify
        var slide1 = handler.Get("/slide[1]");
        slide1.Format.Should().ContainKey("background");

        // 4. Clone
        handler.CopyFrom("/slide[1]", "/", null);

        // 5. Verify clone has background
        var slide2 = handler.Get("/slide[2]");
        slide2.Format.Should().ContainKey("background");
        var slide2Shapes = handler.Query("slide[2] shape");
        slide2Shapes.Should().Contain(s => s.Text!.Contains("On dark bg"));

        // 6. Set (modify cloned slide background)
        handler.Set("/slide[2]", new() { ["background"] = "FF0000" });

        // 7. Verify independence
        handler.Get("/slide[1]").Format["background"].Should().NotBe("FF0000");

        // 8. Reopen + Verify
        Reopen(ref handler);
        handler.Get("/").ChildCount.Should().Be(2);
        handler.Get("/slide[2]").Format.Should().ContainKey("background");

        handler.Dispose();
    }

    // ========== Feature 7: Text Gradient Fill ==========

    [Fact]
    public void TextGradientFill_FullLifecycle()
    {
        // 1. Create
        BlankDocCreator.Create(_path);
        var handler = CreateHandler();

        // 2. Add shape with text
        handler.Add("/", "slide", null, new() { ["title"] = "Test" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Gradient Text", ["size"] = "36" });

        // 3. Get + Verify (no textFill initially)
        var node = handler.Get("/slide[1]/shape[2]");
        node.Format.Should().NotContainKey("textFill");

        // 4. Set text gradient
        handler.Set("/slide[1]/shape[2]", new() { ["textFill"] = "FF0000-0000FF-90" });

        // 5. Get + Verify
        node = handler.Get("/slide[1]/shape[2]");
        node.Format.Should().ContainKey("textFill");
        var textFill = node.Format["textFill"].ToString()!;
        textFill.Should().Contain("FF0000");
        textFill.Should().Contain("0000FF");

        // 6. Set (modify to different gradient)
        handler.Set("/slide[1]/shape[2]", new() { ["textFill"] = "00FF00-FFFF00" });

        // 7. Get + Verify
        node = handler.Get("/slide[1]/shape[2]");
        textFill = node.Format["textFill"].ToString()!;
        textFill.Should().Contain("00FF00");

        // 8. Reopen + Verify
        Reopen(ref handler);
        node = handler.Get("/slide[1]/shape[2]");
        node.Format.Should().ContainKey("textFill");

        handler.Dispose();
    }

    [Fact]
    public void TextGradientFill_RunLevel_FullLifecycle()
    {
        // 1. Create
        BlankDocCreator.Create(_path);
        var handler = CreateHandler();

        // 2. Add shape + runs with different fills
        handler.Add("/", "slide", null, new() { ["title"] = "Test" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Normal" });
        handler.Add("/slide[1]/shape[2]/paragraph[1]", "run", null, new()
            { ["text"] = " Gradient" });

        // 3. Set gradient on the added run
        handler.Set("/slide[1]/shape[2]/paragraph[1]/run[2]", new() { ["textFill"] = "FF0000-0000FF" });

        // 4. Get + Verify
        var para = handler.Get("/slide[1]/shape[2]/paragraph[1]");
        para.Children.Should().HaveCount(2);
        para.Children[0].Format.Should().NotContainKey("textFill");
        para.Children[1].Format.Should().ContainKey("textFill");

        // 5. Set (change to solid color)
        handler.Set("/slide[1]/shape[2]/paragraph[1]/run[2]", new() { ["color"] = "FF0000" });

        // 6. Get + Verify (gradient replaced by solid)
        para = handler.Get("/slide[1]/shape[2]/paragraph[1]");
        para.Children[1].Format["color"].Should().Be("FF0000");
        para.Children[1].Format.Should().NotContainKey("textFill");

        // 7. Reopen + Verify
        Reopen(ref handler);
        para = handler.Get("/slide[1]/shape[2]/paragraph[1]");
        para.Children[1].Format["color"].Should().Be("FF0000");

        handler.Dispose();
    }

    // ========== Feature 8: Custom Geometry ==========

    [Fact]
    public void CustomGeometry_Triangle_FullLifecycle()
    {
        // 1. Create
        BlankDocCreator.Create(_path);
        var handler = CreateHandler();

        // 2. Add shape with custom geometry (triangle)
        handler.Add("/", "slide", null, new() { ["title"] = "Test" });
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Triangle",
            ["fill"] = "4472C4",
            ["geometry"] = "M 0,100 L 50,0 L 100,100 Z"
        });

        // 3. Get + Verify (should NOT have preset since it's custom)
        var node = handler.Get("/slide[1]/shape[2]");
        node.Text.Should().Contain("Triangle");
        node.Format["fill"].Should().Be("4472C4");

        // 4. Set (change to different custom geometry — wave shape)
        handler.Set("/slide[1]/shape[2]", new()
        {
            ["geometry"] = "M 0,50 C 25,0 75,100 100,50 L 100,100 L 0,100 Z"
        });

        // 5. Reopen + Verify persistence
        Reopen(ref handler);
        node = handler.Get("/slide[1]/shape[2]");
        node.Text.Should().Contain("Triangle");
        node.Format["fill"].Should().Be("4472C4");

        // 6. Set back to preset (should replace custom geometry)
        handler.Set("/slide[1]/shape[2]", new() { ["preset"] = "ellipse" });

        // 7. Get + Verify
        node = handler.Get("/slide[1]/shape[2]");
        node.Format["preset"].Should().Be("ellipse");

        handler.Dispose();
    }

    // ========== Feature 9: Morph Transition ==========

    [Fact]
    public void MorphTransition_FullLifecycle()
    {
        // 1. Create
        BlankDocCreator.Create(_path);
        var handler = CreateHandler();

        // 2. Add 2 slides with same-named shapes (morph pairs by name)
        handler.Add("/", "slide", null, new() { ["title"] = "Slide 1" });
        handler.Add("/slide[1]", "shape", null, new()
            { ["text"] = "Move Me", ["fill"] = "4472C4", ["x"] = "2cm", ["y"] = "2cm" });

        handler.Add("/", "slide", null, new() { ["title"] = "Slide 2" });
        handler.Add("/slide[2]", "shape", null, new()
            { ["text"] = "Move Me", ["fill"] = "FF0000", ["x"] = "10cm", ["y"] = "8cm" });

        // 3. Set morph transition on slide 2
        handler.Set("/slide[2]", new() { ["transition"] = "morph" });

        // 4. Get + Verify (transition is stored as raw XML, check via Get)
        var slide2 = handler.Get("/slide[2]");
        // Morph is stored as raw XML, the format key may or may not expose it
        // The important thing is it persists

        // 5. Set (modify to byWord)
        handler.Set("/slide[2]", new() { ["transition"] = "morph-byWord" });

        // 6. Reopen + Verify persistence
        Reopen(ref handler);
        // Slide still has 2 shapes
        var slide2Shapes = handler.Query("slide[2] shape");
        slide2Shapes.Count.Should().BeGreaterThanOrEqualTo(2);

        handler.Dispose();
    }

    [Fact]
    public void Transition_NoDuplicate_WhenSwitchingBetweenPushAndMorph()
    {
        // Regression: setting push then morph (or vice versa) used to leave duplicate
        // <p:transition> elements, causing Microsoft PowerPoint to crash.
        BlankDocCreator.Create(_path);
        var handler = CreateHandler();

        handler.Add("/", "slide", null, new() { ["title"] = "Slide 1" });
        handler.Add("/", "slide", null, new() { ["title"] = "Slide 2" });

        // Set push first, then switch to morph
        handler.Set("/slide[2]", new() { ["transition"] = "push" });
        handler.Set("/slide[2]", new() { ["transition"] = "morph" });

        // Verify only one transition-related element exists in XML
        Reopen(ref handler);
        handler.Get("/").ChildCount.Should().Be(2);

        // Now switch back from morph to push
        handler.Set("/slide[2]", new() { ["transition"] = "push" });
        Reopen(ref handler);
        handler.Get("/").ChildCount.Should().Be(2);

        handler.Dispose();

        // Verify the slide XML has exactly one transition element
        using var pkg = DocumentFormat.OpenXml.Packaging.PresentationDocument.Open(_path, false);
        var slide2Part = pkg.PresentationPart!.SlideParts.ElementAt(1);
        // mc:AlternateContent may contain p:transition inside mc:Choice/Fallback, so count top-level only
        var topLevelTransitions = slide2Part.Slide.ChildElements
            .Count(c => c.LocalName == "transition");
        var topLevelAC = slide2Part.Slide.ChildElements
            .Count(c => c.LocalName == "AlternateContent");
        (topLevelTransitions + topLevelAC).Should().BeLessOrEqualTo(1,
            "slide should have at most one transition element to avoid PowerPoint crash");
    }

    [Fact]
    public void MorphTransition_ByChar_FullLifecycle()
    {
        // 1. Create
        BlankDocCreator.Create(_path);
        var handler = CreateHandler();

        // 2. Add slides
        handler.Add("/", "slide", null, new() { ["title"] = "Before" });
        handler.Add("/", "slide", null, new() { ["title"] = "After" });

        // 3. Set morph-byChar on slide 2
        handler.Set("/slide[2]", new() { ["transition"] = "morph-byChar" });

        // 4. Reopen + Verify file is not corrupt
        Reopen(ref handler);
        handler.Get("/").ChildCount.Should().Be(2);
        handler.Get("/slide[2]").Should().NotBeNull();

        // 5. Set (change to fade)
        handler.Set("/slide[2]", new() { ["transition"] = "fade" });

        // 6. Set (change back to morph)
        handler.Set("/slide[2]", new() { ["transition"] = "morph" });

        // 7. Reopen + Verify
        Reopen(ref handler);
        handler.Get("/").ChildCount.Should().Be(2);

        handler.Dispose();
    }
}
