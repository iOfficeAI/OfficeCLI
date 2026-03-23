// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Bug hunt rounds 36-55: Four targeted bugs found via white-box code review.
/// All tests are expected to FAIL until the bugs are fixed.
/// </summary>
public class PptxRound36Tests : IDisposable
{
    private readonly List<string> _tempFiles = new();

    private string CreateTemp(string ext = ".pptx")
    {
        var path = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}{ext}");
        _tempFiles.Add(path);
        return path;
    }

    public void Dispose()
    {
        foreach (var f in _tempFiles)
            try { File.Delete(f); } catch { }
    }

    // =========================================================================
    // Bug 3 (HIGHEST PRIORITY) — `shape[size<24pt]` strict less-than returns 0
    //
    // Root cause: AttributeFilter.Parse() regex only matches ~=, >=, <=, !=, =
    // The strict less-than (<) and strict greater-than (>) operators are never
    // parsed. FilterOp enum has no LessThan / GreaterThan values.
    // In PowerPointHandler.Selector.cs, ParseShapeSelector() uses a similar
    // regex that also omits < and >.
    // Fix: add FilterOp.LessThan / FilterOp.GreaterThan, update the regex
    // in AttributeFilter.cs to include ">(?!=)" and "<(?!=)", and handle
    // them in MatchOne() with CompareNumeric() < 0 / > 0 respectively.
    // =========================================================================

    [Fact]
    public void Bug3_AttributeFilter_Parse_LessThanOperator_IsRecognized()
    {
        // Currently throws or silently drops the < condition.
        // After fix: should parse as FilterOp.LessThan.
        var conditions = AttributeFilter.Parse("shape[size<24pt]");

        conditions.Should().HaveCount(1);
        conditions[0].Key.Should().Be("size");
        conditions[0].Value.Should().Be("24pt");
        // The op should be a strict less-than (not LessOrEqual)
        // We verify it is NOT LessOrEqual — a node with size==24pt must NOT match
        var nodeEqual = new DocumentNode { Format = { ["size"] = "24pt" } };
        AttributeFilter.MatchAll(nodeEqual, conditions).Should().BeFalse(
            "size=24pt should NOT match size<24pt (strict less-than)");
    }

    [Fact]
    public void Bug3_AttributeFilter_Parse_GreaterThanOperator_IsRecognized()
    {
        // Currently throws or silently drops the > condition.
        // After fix: should parse as FilterOp.GreaterThan.
        var conditions = AttributeFilter.Parse("shape[size>14pt]");

        conditions.Should().HaveCount(1);
        conditions[0].Key.Should().Be("size");
        conditions[0].Value.Should().Be("14pt");
        // A node with size==14pt must NOT match size>14pt (strict greater-than)
        var nodeEqual = new DocumentNode { Format = { ["size"] = "14pt" } };
        AttributeFilter.MatchAll(nodeEqual, conditions).Should().BeFalse(
            "size=14pt should NOT match size>14pt (strict greater-than)");
    }

    [Fact]
    public void Bug3_AttributeFilter_LessThan_MatchAll_ReturnsCorrectSubset()
    {
        // Node with size=12pt should match size<24pt
        var nodeSmall = new DocumentNode { Format = { ["size"] = "12pt" } };
        // Node with size=24pt should NOT match size<24pt
        var nodeEqual = new DocumentNode { Format = { ["size"] = "24pt" } };
        // Node with size=36pt should NOT match size<24pt
        var nodeLarge = new DocumentNode { Format = { ["size"] = "36pt" } };

        var conditions = AttributeFilter.Parse("shape[size<24pt]");

        AttributeFilter.MatchAll(nodeSmall, conditions).Should().BeTrue(
            "12pt < 24pt should match");
        AttributeFilter.MatchAll(nodeEqual, conditions).Should().BeFalse(
            "24pt is not strictly less than 24pt");
        AttributeFilter.MatchAll(nodeLarge, conditions).Should().BeFalse(
            "36pt is not less than 24pt");
    }

    [Fact]
    public void Bug3_AttributeFilter_GreaterThan_MatchAll_ReturnsCorrectSubset()
    {
        var nodeSmall = new DocumentNode { Format = { ["size"] = "12pt" } };
        var nodeEqual = new DocumentNode { Format = { ["size"] = "24pt" } };
        var nodeLarge = new DocumentNode { Format = { ["size"] = "36pt" } };

        var conditions = AttributeFilter.Parse("shape[size>24pt]");

        AttributeFilter.MatchAll(nodeSmall, conditions).Should().BeFalse(
            "12pt is not greater than 24pt");
        AttributeFilter.MatchAll(nodeEqual, conditions).Should().BeFalse(
            "24pt is not strictly greater than 24pt");
        AttributeFilter.MatchAll(nodeLarge, conditions).Should().BeTrue(
            "36pt > 24pt should match");
    }

    [Fact]
    public void Bug3_Pptx_Query_StrictLessThan_FiltersCorrectly()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new() { ["title"] = "Test" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Small", ["size"] = "12" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Threshold", ["size"] = "24" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Large", ["size"] = "36" });

        // Query all shapes, then post-filter with strict less-than
        var allShapes = handler.Query("shape");
        allShapes.Should().HaveCountGreaterOrEqualTo(3);

        var conditions = AttributeFilter.Parse("shape[size<24pt]");
        var filtered = AttributeFilter.Apply(allShapes, conditions);

        // Only "Small" (12pt) should match; "Threshold" (24pt) must NOT be included
        filtered.Should().NotBeEmpty("there should be shapes with size < 24pt");
        filtered.All(n => {
            var sz = n.Format.ContainsKey("size") ? n.Format["size"]?.ToString() : null;
            if (sz == null) return true; // no size key — skip
            var num = decimal.Parse(sz.Replace("pt", "").Trim());
            return num < 24m;
        }).Should().BeTrue("all returned nodes must have size strictly less than 24pt");

        // Verify the threshold shape (24pt) is excluded
        filtered.Should().NotContain(n => n.Text == "Threshold",
            "shape with size exactly 24pt must not match size<24pt");
    }

    [Fact]
    public void Bug3_Pptx_Query_StrictGreaterThan_FiltersCorrectly()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new() { ["title"] = "Test" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Small", ["size"] = "12" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Threshold", ["size"] = "18" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Large", ["size"] = "36" });

        var allShapes = handler.Query("shape");
        var conditions = AttributeFilter.Parse("shape[size>18pt]");
        var filtered = AttributeFilter.Apply(allShapes, conditions);

        // Only "Large" (36pt) should match; "Threshold" (18pt) must NOT be included
        filtered.Should().NotContain(n => n.Text == "Threshold",
            "shape with size exactly 18pt must not match size>18pt");
        filtered.Should().Contain(n => n.Text == "Large",
            "shape with size 36pt should match size>18pt");
    }

    [Fact]
    public void Bug3_AttributeFilter_LessThan_Boundary_StrictlyExcludesEqual()
    {
        // This verifies the critical distinction: < vs <=
        var nodeExact = new DocumentNode { Format = { ["size"] = "24pt" } };

        var lessThan = AttributeFilter.Parse("shape[size<24pt]");
        var lessOrEqual = AttributeFilter.Parse("shape[size<=24pt]");

        AttributeFilter.MatchAll(nodeExact, lessThan).Should().BeFalse(
            "strict < must exclude the boundary value");
        AttributeFilter.MatchAll(nodeExact, lessOrEqual).Should().BeTrue(
            "<= must include the boundary value");
    }

    [Fact]
    public void Bug3_AttributeFilter_GreaterThan_Boundary_StrictlyExcludesEqual()
    {
        var nodeExact = new DocumentNode { Format = { ["size"] = "24pt" } };

        var greaterThan = AttributeFilter.Parse("shape[size>24pt]");
        var greaterOrEqual = AttributeFilter.Parse("shape[size>=24pt]");

        AttributeFilter.MatchAll(nodeExact, greaterThan).Should().BeFalse(
            "strict > must exclude the boundary value");
        AttributeFilter.MatchAll(nodeExact, greaterOrEqual).Should().BeTrue(
            ">= must include the boundary value");
    }

    // =========================================================================
    // Bug 1 — Transition readback loses direction and speed
    //
    // Root cause: ReadSlideTransition(Slide, DocumentNode) in
    // PowerPointHandler.Animations.cs (line ~1240) only stores the transition
    // type name, never the direction attribute from the child element
    // (e.g. WipeTransition.Direction) or the speed attribute from the parent
    // Transition element. So "wipe-left" is stored correctly in XML but read
    // back as just "wipe", and "morph-byWord" comes back as just "morph".
    //
    // Fix: after determining typeName, also read the direction subtype from
    // the child element (transElem) and the speed from trans.Speed, then
    // append them to build the full compound name e.g. "wipe-left",
    // "push-right", "morph-byWord".
    // =========================================================================

    [Fact]
    public void Bug1_Transition_WipeLeft_ReadbackIncludesDirection()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new() { ["title"] = "Slide 1" });
        handler.Set("/slide[1]", new() { ["transition"] = "wipe-left" });

        var node = handler.Get("/slide[1]");

        // Currently fails: readback is "wipe" without the direction
        node.Format.Should().ContainKey("transition");
        var transition = node.Format["transition"]?.ToString();
        transition.Should().NotBe("wipe",
            "direction 'left' should be included in transition readback");
        transition.Should().Be("wipe-left",
            "wipe-left should round-trip through write→read");
    }

    [Fact]
    public void Bug1_Transition_WipeRight_ReadbackIncludesDirection()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new() { ["title"] = "Slide 1" });
        handler.Set("/slide[1]", new() { ["transition"] = "wipe-right" });

        var node = handler.Get("/slide[1]");

        node.Format["transition"]?.ToString().Should().Be("wipe-right",
            "wipe-right should round-trip through write→read");
    }

    [Fact]
    public void Bug1_Transition_PushRight_ReadbackIncludesDirection()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new() { ["title"] = "Slide 1" });
        handler.Set("/slide[1]", new() { ["transition"] = "push-right" });

        var node = handler.Get("/slide[1]");

        node.Format["transition"]?.ToString().Should().Be("push-right",
            "push-right direction should survive round-trip");
    }

    [Fact]
    public void Bug1_Transition_MorphByWord_ReadbackIncludesOption()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new() { ["title"] = "Slide 1" });
        handler.Add("/", "slide", null, new() { ["title"] = "Slide 2" });
        // Morph transitions are applied to slide 2 (the destination slide)
        handler.Set("/slide[2]", new() { ["transition"] = "morph-byWord" });

        var node = handler.Get("/slide[2]");

        node.Format.Should().ContainKey("transition");
        var transition = node.Format["transition"]?.ToString();
        transition.Should().NotBe("morph",
            "morph option 'byWord' should be included in transition readback");
        // The exact format should be "morph-byWord" (or "morph-byword")
        transition?.ToLowerInvariant().Should().Contain("byword",
            "morph-byWord readback should include the byWord option");
    }

    [Fact]
    public void Bug1_Transition_MorphByChar_ReadbackIncludesOption()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new() { ["title"] = "Slide 1" });
        handler.Add("/", "slide", null, new() { ["title"] = "Slide 2" });
        handler.Set("/slide[2]", new() { ["transition"] = "morph-byChar" });

        var node = handler.Get("/slide[2]");

        var transition = node.Format["transition"]?.ToString();
        transition?.ToLowerInvariant().Should().Contain("bychar",
            "morph-byChar readback should include the byChar option");
    }

    [Fact]
    public void Bug1_Transition_WipeLeft_Persistence_DirectionSurvivesReopen()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);

        using (var handler = new PowerPointHandler(path, editable: true))
        {
            handler.Add("/", "slide", null, new() { ["title"] = "Slide" });
            handler.Set("/slide[1]", new() { ["transition"] = "wipe-left" });
        }

        // Reopen and verify
        using var handler2 = new PowerPointHandler(path, editable: false);
        var node = handler2.Get("/slide[1]");
        node.Format["transition"]?.ToString().Should().Be("wipe-left",
            "direction must survive file save and reopen");
    }

    // =========================================================================
    // Bug 2 — Animation delay/easing absent from Query readback
    //
    // Root cause: the Get path (/slide[N]/shape[M]/animation[A]) reads delay,
    // easein, easeout correctly in Query.cs lines 249-261. However the tree
    // walk for delay uses: effectCTn.Parent?.Parent?.Parent as CommonTimeNode
    // This assumes a fixed 3-level nesting: effectCTn < par < seq/par < midCTn
    // but the actual OOXML tree for AfterPrevious animations has different
    // nesting. As a result, midCTn is often null and delay is never written.
    // Similarly, easing attributes (Acceleration/Deceleration) may not be
    // read correctly for all animation classes.
    //
    // Fix: walk the timing tree more robustly to find the Condition that holds
    // the delay value, rather than assuming fixed parent depth.
    // =========================================================================

    [Fact]
    public void Bug2_Animation_Delay_IsReadBack()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new() { ["title"] = "Slide 1" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Animated Shape" });
        // Add animation with 1000ms delay
        handler.Add("/slide[1]/shape[1]", "animation", null,
            new() { ["effect"] = "fly", ["delay"] = "1000" });

        var animNode = handler.Get("/slide[1]/shape[1]/animation[1]");

        animNode.Should().NotBeNull();
        animNode.Format.Should().ContainKey("delay",
            "delay property must appear in animation node Format");
        animNode.Format["delay"]?.ToString().Should().Be("1000",
            "delay value 1000ms must round-trip");
    }

    [Fact]
    public void Bug2_Animation_EaseIn_IsReadBack()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new() { ["title"] = "Slide 1" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Eased Shape" });
        // Add animation with easein=50 (50% acceleration)
        handler.Add("/slide[1]/shape[1]", "animation", null,
            new() { ["effect"] = "fade", ["easein"] = "50" });

        var animNode = handler.Get("/slide[1]/shape[1]/animation[1]");

        animNode.Format.Should().ContainKey("easein",
            "easein must be returned in animation Format");
        animNode.Format["easein"]?.ToString().Should().Be("50",
            "easein=50 must round-trip");
    }

    [Fact]
    public void Bug2_Animation_EaseOut_IsReadBack()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new() { ["title"] = "Slide 1" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Eased Shape" });
        handler.Add("/slide[1]/shape[1]", "animation", null,
            new() { ["effect"] = "fly", ["easeout"] = "30" });

        var animNode = handler.Get("/slide[1]/shape[1]/animation[1]");

        animNode.Format.Should().ContainKey("easeout",
            "easeout must be returned in animation Format");
        animNode.Format["easeout"]?.ToString().Should().Be("30");
    }

    [Fact]
    public void Bug2_Animation_DelayAndEasing_TogetherAreReadBack()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new() { ["title"] = "Slide" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Shape" });
        handler.Add("/slide[1]/shape[1]", "animation", null,
            new() { ["effect"] = "fly", ["delay"] = "500", ["easein"] = "40", ["easeout"] = "40" });

        var animNode = handler.Get("/slide[1]/shape[1]/animation[1]");

        animNode.Format.Should().ContainKey("delay");
        animNode.Format.Should().ContainKey("easein");
        animNode.Format.Should().ContainKey("easeout");
        animNode.Format["delay"]?.ToString().Should().Be("500");
        animNode.Format["easein"]?.ToString().Should().Be("40");
        animNode.Format["easeout"]?.ToString().Should().Be("40");
    }

    [Fact]
    public void Bug2_Animation_Delay_Persistence_SurvivesReopen()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);

        using (var handler = new PowerPointHandler(path, editable: true))
        {
            handler.Add("/", "slide", null, new() { ["title"] = "Slide" });
            handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Animated" });
            handler.Add("/slide[1]/shape[1]", "animation", null,
                new() { ["effect"] = "appear", ["delay"] = "2000" });
        }

        using var handler2 = new PowerPointHandler(path, editable: false);
        var animNode = handler2.Get("/slide[1]/shape[1]/animation[1]");
        animNode.Format.Should().ContainKey("delay");
        animNode.Format["delay"]?.ToString().Should().Be("2000");
    }

    // =========================================================================
    // Bug 4 — Custom freeform geometry not readable back
    //
    // Root cause: ShapeToNode() in PowerPointHandler.NodeBuilder.cs (line ~349)
    // reads geometry only from PresetGeometry:
    //   var presetGeom = shape.ShapeProperties?.GetFirstChild<Drawing.PresetGeometry>();
    //   if (presetGeom?.Preset?.HasValue == true)
    //   {
    //       node.Format["preset"] = presetGeom.Preset.InnerText;
    //       node.Format["geometry"] = presetGeom.Preset.InnerText;
    //   }
    // When a freeform path is set (creating a CustomGeometry element instead),
    // neither "geometry" nor "preset" keys are written to Format.
    //
    // Fix: also check for custGeom:
    //   var custGeom = shape.ShapeProperties?.GetFirstChild<Drawing.CustomGeometry>();
    //   if (custGeom != null)
    //   {
    //       // Reconstruct SVG-like path from pathLst, or use a sentinel like "custom"
    //       node.Format["geometry"] = ReconstructCustomGeometryPath(custGeom);
    //       node.Format["preset"] = "custom";
    //   }
    // =========================================================================

    [Fact]
    public void Bug4_CustomGeometry_GeometryKeyPresentAfterSet()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new() { ["title"] = "Slide" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Freeform" });

        // Set a freeform path (replaces PresetGeometry with CustomGeometry in XML)
        handler.Set("/slide[1]/shape[1]", new() { ["geometry"] = "M0,0 L100,0 L100,100 Z" });

        var node = handler.Get("/slide[1]/shape[1]");

        // Currently fails: both keys are absent after setting custGeom
        node.Format.Should().ContainKey("geometry",
            "geometry key must be present after setting a custom freeform path");
    }

    [Fact]
    public void Bug4_CustomGeometry_PresetKeyPresentAfterSet()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new() { ["title"] = "Slide" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Freeform" });
        handler.Set("/slide[1]/shape[1]", new() { ["geometry"] = "M0,0 L100,0 L50,100 Z" });

        var node = handler.Get("/slide[1]/shape[1]");

        // preset key should reflect that this is custom geometry
        node.Format.Should().ContainKey("preset",
            "preset key must be present for shapes with custom geometry");
    }

    [Fact]
    public void Bug4_CustomGeometry_GeometryValue_ContainsPathData()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new() { ["title"] = "Slide" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Freeform" });
        handler.Set("/slide[1]/shape[1]", new() { ["geometry"] = "M0,0 L100,0 L100,100 Z" });

        var node = handler.Get("/slide[1]/shape[1]");

        node.Format.Should().ContainKey("geometry");
        var geomVal = node.Format["geometry"]?.ToString();

        // The geometry value should either:
        // (a) contain the original path data (full reconstruction), or
        // (b) be a sentinel like "custom" indicating custom geometry is present
        // Either is acceptable — what matters is that it's NOT null/empty
        geomVal.Should().NotBeNullOrWhiteSpace(
            "geometry value must not be empty for a custom freeform shape");
    }

    [Fact]
    public void Bug4_CustomGeometry_Persistence_GeometryKeyAfterReopen()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);

        using (var handler = new PowerPointHandler(path, editable: true))
        {
            handler.Add("/", "slide", null, new() { ["title"] = "Slide" });
            handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Shape" });
            handler.Set("/slide[1]/shape[1]", new() { ["geometry"] = "M0,0 L200,0 L200,100 L0,100 Z" });
        }

        using var handler2 = new PowerPointHandler(path, editable: false);
        var node = handler2.Get("/slide[1]/shape[1]");

        node.Format.Should().ContainKey("geometry",
            "geometry key must survive file save and reopen");
    }

    [Fact]
    public void Bug4_PresetGeometry_StillReadBack_AfterCustomGeometryCodeAdded()
    {
        // Regression guard: adding custGeom readback must not break preset readback
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new() { ["title"] = "Slide" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Preset Shape" });
        // Leave geometry as default (rect) — preset geometry should still be readable
        // Or explicitly set a named preset:
        handler.Set("/slide[1]/shape[1]", new() { ["geometry"] = "roundRect" });

        var node = handler.Get("/slide[1]/shape[1]");

        node.Format.Should().ContainKey("geometry");
        node.Format.Should().ContainKey("preset");
        node.Format["geometry"]?.ToString().Should().Be("roundRect");
        node.Format["preset"]?.ToString().Should().Be("roundRect");
    }
}
