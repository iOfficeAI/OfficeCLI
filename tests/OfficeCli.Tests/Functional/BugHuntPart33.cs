// Bug hunt Part 33 — Effects element ordering and separator bugs:
// 1. Glow with semicolon separator ("COLOR;RADIUS") should parse correctly
// 2. Glow with semicolon separator produces valid XML (no schema errors)
// 3. Glow with dash separator still works as before
// 4. Scene3D must come before sp3d in schema order when both bevel and rot3d are set
// 5. Bevel + rot3d combination produces no validation errors
// 6. Bevel + rot3d values persist after reopen
// 7. Shadow with semicolon separator should parse correctly (similar to glow bug)
// 8. Bevel with semicolon separator should parse correctly (similar to glow bug)
// 9. EffectLst ordering: adding glow/shadow AFTER bevel+rot3d should not violate schema
// 10. Outline ordering: setting line AFTER bevel+rot3d should not violate schema

using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

public class BugHuntPart33 : IDisposable
{
    private readonly string _pptxPath;
    private PowerPointHandler _pptxHandler;

    public BugHuntPart33()
    {
        _pptxPath = Path.Combine(Path.GetTempPath(), $"bughunt33_{Guid.NewGuid():N}.pptx");
        BlankDocCreator.Create(_pptxPath);
        _pptxHandler = new PowerPointHandler(_pptxPath, editable: true);
    }

    public void Dispose()
    {
        _pptxHandler.Dispose();
        if (File.Exists(_pptxPath)) File.Delete(_pptxPath);
    }

    private void Reopen()
    {
        _pptxHandler.Dispose();
        _pptxHandler = new PowerPointHandler(_pptxPath, editable: true);
    }

    // ── Bug 1: Glow semicolon separator ──────────────────────────────

    [Fact]
    public void Add_Shape_WithGlowSemicolonSeparator_GlowIsReadBack()
    {
        // "00FFFF;15" should parse as color=00FFFF, radius=15pt, opacity=75 (default)
        _pptxHandler.Add("/", "slide", null, new() { ["title"] = "GlowTest" });
        _pptxHandler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Glow Test",
            ["glow"] = "00FFFF;15"
        });

        var node = _pptxHandler.Get("/slide[1]/shape[2]");
        node.Format.Should().ContainKey("glow");
        // Readback format: "COLOR-RADIUS-OPACITY"
        node.Format["glow"].Should().Be("00FFFF-15-75");
    }

    [Fact]
    public void Add_Shape_WithGlowSemicolonSeparator_ProducesNoValidationErrors()
    {
        _pptxHandler.Add("/", "slide", null, new() { ["title"] = "GlowValidate" });
        _pptxHandler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Validate",
            ["glow"] = "FF0000;20"
        });

        var errors = _pptxHandler.Validate();
        errors.Should().BeEmpty("glow with semicolon separator should produce valid XML");
    }

    [Fact]
    public void Set_Shape_GlowSemicolonSeparator_IsUpdated()
    {
        _pptxHandler.Add("/", "slide", null, new() { ["title"] = "GlowSet" });
        _pptxHandler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Before",
            ["glow"] = "FF0000-8"
        });

        _pptxHandler.Set("/slide[1]/shape[2]", new() { ["glow"] = "00FF00;12" });

        var node = _pptxHandler.Get("/slide[1]/shape[2]");
        node.Format["glow"].Should().Be("00FF00-12-75");
    }

    [Fact]
    public void Add_Shape_WithGlowDashSeparator_StillWorks()
    {
        // Dash separator should continue to work as before
        _pptxHandler.Add("/", "slide", null, new() { ["title"] = "GlowDash" });
        _pptxHandler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Dash",
            ["glow"] = "0000FF-10-60"
        });

        var node = _pptxHandler.Get("/slide[1]/shape[2]");
        node.Format["glow"].Should().Be("0000FF-10-60");
    }

    [Fact]
    public void GlowSemicolon_Persist_SurvivesReopenFile()
    {
        _pptxHandler.Add("/", "slide", null, new() { ["title"] = "GlowPersist" });
        _pptxHandler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Persist",
            ["glow"] = "00FFFF;15"
        });

        Reopen();

        var node = _pptxHandler.Get("/slide[1]/shape[2]");
        node.Format["glow"].Should().Be("00FFFF-15-75");

        var errors = _pptxHandler.Validate();
        errors.Should().BeEmpty();
    }

    // ── Bug 2: Scene3D element ordering (bevel + rot3d) ─────────────

    [Fact]
    public void Add_Shape_WithBevelAndRot3d_ProducesNoValidationErrors()
    {
        // When both bevel (sp3d) and rot3d (scene3d) are applied,
        // scene3d must come before sp3d in schema order
        _pptxHandler.Add("/", "slide", null, new() { ["title"] = "3DTest" });
        _pptxHandler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "3D Shape",
            ["bevel"] = "circle-6-6",
            ["rot3d"] = "20,340,0"
        });

        var errors = _pptxHandler.Validate();
        errors.Should().BeEmpty("scene3d should be ordered before sp3d in spPr children");
    }

    [Fact]
    public void Add_Shape_WithBevelAndRot3d_BothAreReadBack()
    {
        _pptxHandler.Add("/", "slide", null, new() { ["title"] = "3DReadback" });
        _pptxHandler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "3D Shape",
            ["bevel"] = "circle-6-6",
            ["rot3d"] = "20,340,0"
        });

        var node = _pptxHandler.Get("/slide[1]/shape[2]");
        node.Format.Should().ContainKey("bevel");
        node.Format["bevel"].ToString().Should().Contain("circle");
        node.Format.Should().ContainKey("rot3d");
        node.Format["rot3d"].Should().Be("20,340,0");
    }

    [Fact]
    public void Set_Shape_BevelThenRot3d_ProducesNoValidationErrors()
    {
        _pptxHandler.Add("/", "slide", null, new() { ["title"] = "3DSet" });
        _pptxHandler.Add("/slide[1]", "shape", null, new() { ["text"] = "Plain" });

        // Set bevel first, then rot3d — this was the order that triggered the bug
        _pptxHandler.Set("/slide[1]/shape[2]", new() { ["bevel"] = "softRound-8-8" });
        _pptxHandler.Set("/slide[1]/shape[2]", new() { ["rot3d"] = "15,330,0" });

        var errors = _pptxHandler.Validate();
        errors.Should().BeEmpty("setting bevel then rot3d should maintain correct schema order");

        var node = _pptxHandler.Get("/slide[1]/shape[2]");
        node.Format["rot3d"].Should().Be("15,330,0");
    }

    [Fact]
    public void BevelAndRot3d_Persist_SurvivesReopenFile()
    {
        _pptxHandler.Add("/", "slide", null, new() { ["title"] = "3DPersist" });
        _pptxHandler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Persist 3D",
            ["bevel"] = "circle-6-6",
            ["rot3d"] = "20,340,0"
        });

        Reopen();

        var node = _pptxHandler.Get("/slide[1]/shape[2]");
        node.Format["bevel"].ToString().Should().Contain("circle");
        node.Format["rot3d"].Should().Be("20,340,0");

        var errors = _pptxHandler.Validate();
        errors.Should().BeEmpty();
    }

    // ── Combined: all effects together ───────────────────────────────

    [Fact]
    public void Add_Shape_WithGlowBevelRot3d_AllValid()
    {
        // Combine all three effects — glow (effectLst), bevel (sp3d), rot3d (scene3d)
        _pptxHandler.Add("/", "slide", null, new() { ["title"] = "Combined" });
        _pptxHandler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "All Effects",
            ["glow"] = "00FFFF;15",
            ["bevel"] = "circle-6-6",
            ["rot3d"] = "20,340,0"
        });

        var errors = _pptxHandler.Validate();
        errors.Should().BeEmpty("glow + bevel + rot3d combined should produce valid XML");

        var node = _pptxHandler.Get("/slide[1]/shape[2]");
        node.Format["glow"].Should().Be("00FFFF-15-75");
        node.Format["bevel"].ToString().Should().Contain("circle");
        node.Format["rot3d"].Should().Be("20,340,0");
    }

    // ── Bug 3: Shadow semicolon separator ────────────────────────────

    [Fact]
    public void Add_Shape_WithShadowSemicolonSeparator_ShadowIsReadBack()
    {
        // "000000;6;315;4;50" should parse as color=000000, blur=6, angle=315, dist=4, opacity=50
        _pptxHandler.Add("/", "slide", null, new() { ["title"] = "ShadowTest" });
        _pptxHandler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Shadow Test",
            ["shadow"] = "000000;6;315;4;50"
        });

        var node = _pptxHandler.Get("/slide[1]/shape[2]");
        node.Format.Should().ContainKey("shadow");
        // Readback format: "COLOR-BLUR-ANGLE-DIST-OPACITY"
        node.Format["shadow"].Should().Be("000000-6-315-4-50");
    }

    [Fact]
    public void Add_Shape_WithShadowSemicolonSeparator_ProducesNoValidationErrors()
    {
        _pptxHandler.Add("/", "slide", null, new() { ["title"] = "ShadowValidate" });
        _pptxHandler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Validate",
            ["shadow"] = "333333;8;45;3;60"
        });

        var errors = _pptxHandler.Validate();
        errors.Should().BeEmpty("shadow with semicolon separator should produce valid XML");
    }

    [Fact]
    public void Set_Shape_ShadowSemicolonSeparator_IsUpdated()
    {
        _pptxHandler.Add("/", "slide", null, new() { ["title"] = "ShadowSet" });
        _pptxHandler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Before",
            ["shadow"] = "000000-4-45-3-40"
        });

        _pptxHandler.Set("/slide[1]/shape[2]", new() { ["shadow"] = "FF0000;10;90;5;80" });

        var node = _pptxHandler.Get("/slide[1]/shape[2]");
        node.Format["shadow"].Should().Be("FF0000-10-90-5-80");
    }

    [Fact]
    public void ShadowSemicolon_Persist_SurvivesReopenFile()
    {
        _pptxHandler.Add("/", "slide", null, new() { ["title"] = "ShadowPersist" });
        _pptxHandler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Persist",
            ["shadow"] = "000000;6;315;4;50"
        });

        Reopen();

        var node = _pptxHandler.Get("/slide[1]/shape[2]");
        node.Format["shadow"].Should().Be("000000-6-315-4-50");

        var errors = _pptxHandler.Validate();
        errors.Should().BeEmpty();
    }

    // ── Bug 4: Bevel semicolon separator ─────────────────────────────

    [Fact]
    public void Add_Shape_WithBevelSemicolonSeparator_BevelIsReadBack()
    {
        // "circle;6;6" should parse as preset=circle, width=6pt, height=6pt
        _pptxHandler.Add("/", "slide", null, new() { ["title"] = "BevelTest" });
        _pptxHandler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Bevel Test",
            ["bevel"] = "circle;6;6"
        });

        var node = _pptxHandler.Get("/slide[1]/shape[2]");
        node.Format.Should().ContainKey("bevel");
        node.Format["bevel"].ToString().Should().Contain("circle");
    }

    [Fact]
    public void Add_Shape_WithBevelSemicolonSeparator_ProducesNoValidationErrors()
    {
        _pptxHandler.Add("/", "slide", null, new() { ["title"] = "BevelValidate" });
        _pptxHandler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Validate",
            ["bevel"] = "softRound;8;8"
        });

        var errors = _pptxHandler.Validate();
        errors.Should().BeEmpty("bevel with semicolon separator should produce valid XML");
    }

    [Fact]
    public void Set_Shape_BevelSemicolonSeparator_IsUpdated()
    {
        _pptxHandler.Add("/", "slide", null, new() { ["title"] = "BevelSet" });
        _pptxHandler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Before",
            ["bevel"] = "circle-6-6"
        });

        _pptxHandler.Set("/slide[1]/shape[2]", new() { ["bevel"] = "angle;10;10" });

        var node = _pptxHandler.Get("/slide[1]/shape[2]");
        node.Format["bevel"].ToString().Should().Contain("angle");
    }

    [Fact]
    public void BevelSemicolon_Persist_SurvivesReopenFile()
    {
        _pptxHandler.Add("/", "slide", null, new() { ["title"] = "BevelPersist" });
        _pptxHandler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Persist",
            ["bevel"] = "circle;6;6"
        });

        Reopen();

        var node = _pptxHandler.Get("/slide[1]/shape[2]");
        node.Format["bevel"].ToString().Should().Contain("circle");

        var errors = _pptxHandler.Validate();
        errors.Should().BeEmpty();
    }

    // ── Bug 5: EffectLst ordering — adding effects AFTER 3D elements ─

    [Fact]
    public void Set_Shape_GlowAfterBevelRot3d_ProducesNoValidationErrors()
    {
        // Create shape with bevel+rot3d first (creates sp3d + scene3d),
        // then add glow (creates effectLst) — effectLst must come BEFORE scene3d in schema
        _pptxHandler.Add("/", "slide", null, new() { ["title"] = "EffectOrder" });
        _pptxHandler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "3D first",
            ["bevel"] = "circle-6-6",
            ["rot3d"] = "20,340,0"
        });

        // Now add glow AFTER 3D elements already exist
        _pptxHandler.Set("/slide[1]/shape[2]", new() { ["glow"] = "00FFFF-15" });

        var errors = _pptxHandler.Validate();
        errors.Should().BeEmpty("effectLst should be ordered before scene3d/sp3d in spPr");

        var node = _pptxHandler.Get("/slide[1]/shape[2]");
        node.Format["glow"].Should().Be("00FFFF-15-75");
        node.Format["bevel"].ToString().Should().Contain("circle");
        node.Format["rot3d"].Should().Be("20,340,0");
    }

    [Fact]
    public void Set_Shape_ShadowAfterBevelRot3d_ProducesNoValidationErrors()
    {
        _pptxHandler.Add("/", "slide", null, new() { ["title"] = "ShadowOrder" });
        _pptxHandler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "3D first",
            ["bevel"] = "circle-6-6",
            ["rot3d"] = "10,350,0"
        });

        _pptxHandler.Set("/slide[1]/shape[2]", new() { ["shadow"] = "000000-6-45-3-50" });

        var errors = _pptxHandler.Validate();
        errors.Should().BeEmpty("effectLst (shadow) should be ordered before scene3d/sp3d");
    }

    [Fact]
    public void Set_Shape_ReflectionAfterBevelRot3d_ProducesNoValidationErrors()
    {
        _pptxHandler.Add("/", "slide", null, new() { ["title"] = "ReflOrder" });
        _pptxHandler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "3D first",
            ["bevel"] = "circle-6-6",
            ["rot3d"] = "10,350,0"
        });

        _pptxHandler.Set("/slide[1]/shape[2]", new() { ["reflection"] = "half" });

        var errors = _pptxHandler.Validate();
        errors.Should().BeEmpty("effectLst (reflection) should be ordered before scene3d/sp3d");
    }

    // ── Bug 6: Outline (ln) ordering — setting line AFTER 3D elements ─

    [Fact]
    public void Set_Shape_LineAfterBevelRot3d_ProducesNoValidationErrors()
    {
        // Schema order: fill → ln → effectLst → scene3d → sp3d
        // If line is added after scene3d/sp3d exist, it must be inserted correctly
        _pptxHandler.Add("/", "slide", null, new() { ["title"] = "LineOrder" });
        _pptxHandler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "3D first",
            ["bevel"] = "circle-6-6",
            ["rot3d"] = "10,350,0"
        });

        _pptxHandler.Set("/slide[1]/shape[2]", new()
        {
            ["line"] = "FF0000",
            ["lineWidth"] = "2pt"
        });

        var errors = _pptxHandler.Validate();
        errors.Should().BeEmpty("outline (ln) should be ordered before effectLst/scene3d/sp3d");
    }

    [Fact]
    public void Set_Shape_LineAndGlowAfterBevel_ProducesNoValidationErrors()
    {
        _pptxHandler.Add("/", "slide", null, new() { ["title"] = "AllOrder" });
        _pptxHandler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "3D only",
            ["bevel"] = "circle-6-6"
        });

        // Add line, glow, and rot3d incrementally
        _pptxHandler.Set("/slide[1]/shape[2]", new() { ["line"] = "0000FF" });
        _pptxHandler.Set("/slide[1]/shape[2]", new() { ["glow"] = "FF0000-10" });
        _pptxHandler.Set("/slide[1]/shape[2]", new() { ["rot3d"] = "15,330,0" });

        var errors = _pptxHandler.Validate();
        errors.Should().BeEmpty("incrementally added ln, effectLst, scene3d should all be in schema order");
    }

    [Fact]
    public void EffectLstOrdering_Persist_SurvivesReopenFile()
    {
        _pptxHandler.Add("/", "slide", null, new() { ["title"] = "OrderPersist" });
        _pptxHandler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "3D first",
            ["bevel"] = "circle-6-6",
            ["rot3d"] = "20,340,0"
        });
        _pptxHandler.Set("/slide[1]/shape[2]", new() { ["glow"] = "00FFFF-15" });

        Reopen();

        var errors = _pptxHandler.Validate();
        errors.Should().BeEmpty();

        var node = _pptxHandler.Get("/slide[1]/shape[2]");
        node.Format["glow"].Should().Be("00FFFF-15-75");
        node.Format["bevel"].ToString().Should().Contain("circle");
        node.Format["rot3d"].Should().Be("20,340,0");
    }
}
