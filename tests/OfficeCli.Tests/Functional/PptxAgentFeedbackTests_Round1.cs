// Agent Feedback Round 1 — Bug tests discovered by Agent A's PPT testing
// Tests for schema ordering violations and animation bugs.

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Tests.Functional;

public class PptxAgentFeedbackTests_Round1 : IDisposable
{
    private readonly string _path;
    private PowerPointHandler _handler;

    public PptxAgentFeedbackTests_Round1()
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

    // ==================== Reflection Helpers ====================

    private PresentationDocument GetDoc()
    {
        return (PresentationDocument)_handler.GetType()
            .GetField("_doc", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)!
            .GetValue(_handler)!;
    }

    private Shape GetFirstShape()
    {
        var doc = GetDoc();
        var slidePart = doc.PresentationPart!.SlideParts.First();
        return slidePart.Slide.Descendants<Shape>().First();
    }

    private Picture GetFirstPicture()
    {
        var doc = GetDoc();
        var slidePart = doc.PresentationPart!.SlideParts.First();
        return slidePart.Slide.Descendants<Picture>().First();
    }

    private static int ChildIndex<T>(OpenXmlElement parent) where T : OpenXmlElement
    {
        int i = 0;
        foreach (var child in parent.ChildElements)
        {
            if (child is T) return i;
            i++;
        }
        return -1;
    }

    private static List<string> ChildLocalNames(OpenXmlElement parent)
    {
        return parent.ChildElements.Select(c => c.LocalName).ToList();
    }

    // ==================== Bug 1: solidFill before prstGeom after opacity then gradient ====================
    // When setting opacity (which touches solidFill) then switching to gradient,
    // the old solidFill is not removed and ends up before prstGeom, violating
    // CT_ShapeProperties schema: xfrm → custGeom/prstGeom → fill → ln → effectLst

    [Fact]
    public void Bug1_SetOpacity_ThenGradient_SolidFillShouldBeRemoved()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Test",
            ["fill"] = "FF0000"
        });

        // Set opacity on the solid fill
        _handler.Set("/slide[1]/shape[1]", new() { ["opacity"] = "0.5" });

        // Now switch to gradient — should remove old solidFill
        _handler.Set("/slide[1]/shape[1]", new() { ["gradient"] = "FF0000-0000FF-90" });

        var shape = GetFirstShape();
        var spPr = shape.ShapeProperties!;

        // After setting gradient, solidFill should be gone
        spPr.GetFirstChild<Drawing.SolidFill>().Should().BeNull(
            "solidFill must be removed when gradient is applied");

        // Gradient fill should exist
        spPr.GetFirstChild<Drawing.GradientFill>().Should().NotBeNull(
            "gradientFill should be present after setting gradient");
    }

    [Fact]
    public void Bug1_SetOpacity_ThenGradient_FillAfterPrstGeom()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Test",
            ["fill"] = "FF0000"
        });

        _handler.Set("/slide[1]/shape[1]", new() { ["opacity"] = "0.5" });
        _handler.Set("/slide[1]/shape[1]", new() { ["gradient"] = "FF0000-0000FF-90" });

        var shape = GetFirstShape();
        var spPr = shape.ShapeProperties!;
        var names = ChildLocalNames(spPr);

        // Schema order: xfrm → prstGeom → fill (gradFill) → ln → effectLst
        var prstGeomIdx = names.IndexOf("prstGeom");
        var gradFillIdx = names.IndexOf("gradFill");
        var solidFillIdx = names.IndexOf("solidFill");

        // solidFill should not exist at all
        solidFillIdx.Should().Be(-1,
            "solidFill should be fully removed after gradient is applied");

        // gradFill must come after prstGeom
        if (prstGeomIdx >= 0 && gradFillIdx >= 0)
        {
            gradFillIdx.Should().BeGreaterThan(prstGeomIdx,
                "gradFill must come after prstGeom in CT_ShapeProperties schema order");
        }
    }

    // ==================== Bug 2: ln child solidFill after prstDash ====================
    // CT_LineProperties schema order: fill (solidFill/noFill/gradFill/pattFill)
    //   → prstDash/custDash → round/bevel/miter → headEnd → tailEnd
    // When setting lineDash and line color separately, solidFill ends up after prstDash.

    [Fact]
    public void Bug2_LineDash_ThenLineColor_SolidFillBeforePrstDash()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Test",
            ["line"] = "FF0000",
            ["lineWidth"] = "2pt"
        });

        // Set dash first, then line color
        _handler.Set("/slide[1]/shape[1]", new() { ["lineDash"] = "dash" });
        _handler.Set("/slide[1]/shape[1]", new() { ["line"] = "0000FF" });

        var shape = GetFirstShape();
        var outline = shape.ShapeProperties?.GetFirstChild<Drawing.Outline>();
        outline.Should().NotBeNull("outline should exist");

        var names = ChildLocalNames(outline!);
        var solidFillIdx = names.IndexOf("solidFill");
        var prstDashIdx = names.IndexOf("prstDash");

        solidFillIdx.Should().BeGreaterOrEqualTo(0, "solidFill should exist in outline");
        prstDashIdx.Should().BeGreaterOrEqualTo(0, "prstDash should exist in outline");

        solidFillIdx.Should().BeLessThan(prstDashIdx,
            "solidFill must come before prstDash in CT_LineProperties schema order");
    }

    [Fact]
    public void Bug2_LineColor_ThenLineDash_SchemaOrderPreserved()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Test",
            ["line"] = "FF0000",
            ["lineWidth"] = "2pt"
        });

        // Set color first, then dash — order should still be correct
        _handler.Set("/slide[1]/shape[1]", new() { ["lineDash"] = "dot" });

        var shape = GetFirstShape();
        var outline = shape.ShapeProperties?.GetFirstChild<Drawing.Outline>();
        outline.Should().NotBeNull();

        var names = ChildLocalNames(outline!);
        var solidFillIdx = names.IndexOf("solidFill");
        var prstDashIdx = names.IndexOf("prstDash");

        if (solidFillIdx >= 0 && prstDashIdx >= 0)
        {
            solidFillIdx.Should().BeLessThan(prstDashIdx,
                "solidFill must come before prstDash in CT_LineProperties schema order");
        }
    }

    // ==================== Bug 4: pPr lnSpc after spcBef/spcAft ====================
    // CT_TextParagraphProperties schema: lnSpc → spcBef → spcAft → buClr → ...
    // When setting spaceBefore first, then lineSpacing, lnSpc ends up after spcBef.

    [Fact]
    public void Bug4_SpaceBefore_ThenLineSpacing_SchemaOrder()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Test" });

        // Set spaceBefore first
        _handler.Set("/slide[1]/shape[1]", new() { ["spaceBefore"] = "12pt" });
        // Then set lineSpacing — lnSpc should be inserted before spcBef
        _handler.Set("/slide[1]/shape[1]", new() { ["lineSpacing"] = "1.5x" });

        var shape = GetFirstShape();
        var para = shape.TextBody!.Elements<Drawing.Paragraph>().First();
        var pPr = para.ParagraphProperties!;
        var names = ChildLocalNames(pPr);

        var lnSpcIdx = names.IndexOf("lnSpc");
        var spcBefIdx = names.IndexOf("spcBef");

        lnSpcIdx.Should().BeGreaterOrEqualTo(0, "lnSpc should exist");
        spcBefIdx.Should().BeGreaterOrEqualTo(0, "spcBef should exist");
        lnSpcIdx.Should().BeLessThan(spcBefIdx,
            "lnSpc must come before spcBef in CT_TextParagraphProperties schema order");
    }

    [Fact]
    public void Bug4_SpaceAfter_ThenLineSpacing_SchemaOrder()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Test" });

        _handler.Set("/slide[1]/shape[1]", new() { ["spaceAfter"] = "6pt" });
        _handler.Set("/slide[1]/shape[1]", new() { ["lineSpacing"] = "2x" });

        var shape = GetFirstShape();
        var para = shape.TextBody!.Elements<Drawing.Paragraph>().First();
        var pPr = para.ParagraphProperties!;
        var names = ChildLocalNames(pPr);

        var lnSpcIdx = names.IndexOf("lnSpc");
        var spcAftIdx = names.IndexOf("spcAft");

        lnSpcIdx.Should().BeGreaterOrEqualTo(0, "lnSpc should exist");
        spcAftIdx.Should().BeGreaterOrEqualTo(0, "spcAft should exist");
        lnSpcIdx.Should().BeLessThan(spcAftIdx,
            "lnSpc must come before spcAft in CT_TextParagraphProperties schema order");
    }

    [Fact]
    public void Bug4_AllSpacingSet_CorrectOrder()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Test" });

        // Set all three in worst-case order (spaceAfter, spaceBefore, lineSpacing)
        _handler.Set("/slide[1]/shape[1]", new() { ["spaceAfter"] = "6pt" });
        _handler.Set("/slide[1]/shape[1]", new() { ["spaceBefore"] = "12pt" });
        _handler.Set("/slide[1]/shape[1]", new() { ["lineSpacing"] = "1.5x" });

        var shape = GetFirstShape();
        var para = shape.TextBody!.Elements<Drawing.Paragraph>().First();
        var pPr = para.ParagraphProperties!;
        var names = ChildLocalNames(pPr);

        var lnSpcIdx = names.IndexOf("lnSpc");
        var spcBefIdx = names.IndexOf("spcBef");
        var spcAftIdx = names.IndexOf("spcAft");

        lnSpcIdx.Should().BeGreaterOrEqualTo(0, "lnSpc should exist");
        spcBefIdx.Should().BeGreaterOrEqualTo(0, "spcBef should exist");
        spcAftIdx.Should().BeGreaterOrEqualTo(0, "spcAft should exist");

        // Schema order: lnSpc < spcBef < spcAft
        lnSpcIdx.Should().BeLessThan(spcBefIdx,
            "lnSpc must come before spcBef in schema order");
        spcBefIdx.Should().BeLessThan(spcAftIdx,
            "spcBef must come before spcAft in schema order");
    }

    // ==================== Bug 5: animation Set accumulates instead of replacing ====================
    // Calling Set with animation multiple times should replace the existing animation,
    // not append additional animation entries (animation, animation2, animation3...).

    [Fact]
    public void Bug5_SetAnimation_MultipleTimes_ShouldReplace()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Test" });

        // Set animation three times
        _handler.Set("/slide[1]/shape[1]", new() { ["animation"] = "fade-entrance-500" });
        _handler.Set("/slide[1]/shape[1]", new() { ["animation"] = "fly-entrance-400" });
        _handler.Set("/slide[1]/shape[1]", new() { ["animation"] = "zoom-entrance-600" });

        var node = _handler.Get("/slide[1]/shape[1]");

        // Should have exactly one animation (the last one set)
        node.Format.Should().ContainKey("animation",
            "the shape should have an animation");
        node.Format["animation"]!.ToString().Should().Contain("zoom",
            "the animation should be the last one set (zoom)");

        // There should NOT be animation2 or animation3
        node.Format.Should().NotContainKey("animation2",
            "previous animations should be replaced, not accumulated");
        node.Format.Should().NotContainKey("animation3",
            "previous animations should be replaced, not accumulated");
    }

    [Fact]
    public void Bug5_SetAnimation_ThenNone_ThenNew_OnlyOneAnimation()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Test" });

        _handler.Set("/slide[1]/shape[1]", new() { ["animation"] = "fade-entrance-500" });
        _handler.Set("/slide[1]/shape[1]", new() { ["animation"] = "none" });
        _handler.Set("/slide[1]/shape[1]", new() { ["animation"] = "fly-entrance-400" });

        var node = _handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("animation");
        node.Format["animation"]!.ToString().Should().Contain("fly");
        node.Format.Should().NotContainKey("animation2",
            "after removing and re-adding, there should be only one animation");
    }

    // ==================== Bug 6: animation direction lost on readback ====================
    // Setting "fly-entrance-left-500" reads back as "fly-entrance-500" — direction is lost.

    [Fact]
    public void Bug6_FlyAnimation_DirectionPreservedOnReadback()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Test" });

        _handler.Set("/slide[1]/shape[1]", new() { ["animation"] = "fly-entrance-left-500" });

        var node = _handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("animation");

        // The readback should include the direction
        var animValue = node.Format["animation"]!.ToString()!;
        animValue.Should().Contain("left",
            "fly animation direction 'left' should be preserved on readback; " +
            $"got '{animValue}'");
    }

    [Fact]
    public void Bug6_WipeAnimation_DirectionPreservedOnReadback()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Test" });

        _handler.Set("/slide[1]/shape[1]", new() { ["animation"] = "wipe-entrance-right-600" });

        var node = _handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("animation");

        var animValue = node.Format["animation"]!.ToString()!;
        // Direction should be part of the readback
        animValue.Should().Contain("right",
            "wipe animation direction 'right' should be preserved on readback; " +
            $"got '{animValue}'");
    }

    [Fact]
    public void Bug6_FlyAnimation_UpDirection_PreservedOnReadback()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Test" });

        _handler.Set("/slide[1]/shape[1]", new() { ["animation"] = "fly-entrance-up-400" });

        var node = _handler.Get("/slide[1]/shape[1]");
        node.Format.Should().ContainKey("animation");

        var animValue = node.Format["animation"]!.ToString()!;
        animValue.Should().Contain("up",
            "fly animation direction 'up' should be preserved on readback; " +
            $"got '{animValue}'");
    }
}
