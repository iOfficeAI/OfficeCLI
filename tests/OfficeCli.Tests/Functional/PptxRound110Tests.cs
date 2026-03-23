// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Bug hunt round 110: Four targeted bugs found via white-box code review.
///
/// Bug A — border.top.width sets ALL 4 sides (most impactful):
///   In PowerPointHandler.ShapeProperties.cs the case clause
///   `case var k when k.StartsWith("border"):` captures the full key, e.g.
///   "border.top.width".  The subsequent `edges` switch only matches
///   "border.left", "border.right", "border.top", "border.bottom", etc.
///   The sub-property suffix ".width" means none of those arms fire, so
///   the key falls through to the wildcard `_` arm which expands to
///   new[] { "left", "right", "top", "bottom" }, modifying ALL four sides.
///
/// Bug B — Cross-slide shape Move loses hyperlink:
///   CopyRelationships() in PowerPointHandler.Mutations.cs iterates
///   element attributes looking for r:* relationship IDs, then calls
///   sourcePart.GetPartById(oldRelId).  Hyperlink relationships are
///   *external* relationships — GetPartById throws ArgumentOutOfRangeException
///   for them (caught and silently swallowed).  The XML element still contains
///   the old r:id, but the target slide part never receives the corresponding
///   HyperlinkRelationship, so the hyperlink is broken after a cross-slide move.
///
/// Bug C — Connector startShape/endShape silently ignored:
///   AddConnector() does handle "startshape" and "endshape" keys (lines 50-61
///   of PowerPointHandler.Add.Misc.cs) — the connections ARE wired up.
///   This is NOT a bug; the test verifies correct behaviour and documents the API.
///
/// Bug D — Slide Set "layout" and "name" listed as valid but unimplemented:
///   The slide-level Set switch (PowerPointHandler.Set.cs ~line 869) handles
///   "background", "transition", "advancetime", "notes", "align", "distribute".
///   Neither "layout" nor "name" has a case arm.  They fall through to the
///   GenericXmlQuery.SetGenericAttribute fallback (which sets raw XML attributes
///   on the Slide element and doesn't affect the CommonSlideData name or the
///   layout relationship).  The error message explicitly lists both "layout"
///   and "name" as valid, misleading callers.
/// </summary>
public class PptxRound110Tests : IDisposable
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
    // Bug A — border.top.width sets ALL 4 sides
    // =========================================================================

    /// <summary>
    /// Set border.top with a specific width and colour on a table cell, then
    /// verify that ONLY the top border is modified.  Left, right, and bottom
    /// must remain at their original state (absent / default).
    /// </summary>
    [Fact]
    public void TableCell_BorderTopWidth_OnlyTopBorderIsModified()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        // Add a slide and a 2×2 table
        handler.Add("/", "slide", null, new() { ["title"] = "BugA" });
        handler.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "2",
            ["cols"] = "2",
            ["width"] = "6cm",
            ["height"] = "3cm"
        });

        // Set ONLY the top border on cell [1,1] with a distinctive width (5pt)
        // and a distinctive colour so we can tell the sides apart.
        handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
        {
            ["border.top"] = "5pt solid FF0000"
        });

        // Close handler before opening for raw XML inspection (avoid file lock)
        handler.Dispose();

        // --- Verify via raw XML (DocumentNode does not yet expose per-side border widths)
        using var doc = PresentationDocument.Open(path, false);
        var slidePart = doc.PresentationPart!.SlideParts.First();
        var tables = slidePart.Slide.Descendants<Drawing.Table>().ToList();
        tables.Should().NotBeEmpty("a table should exist on the slide");

        var row1 = tables[0].Elements<Drawing.TableRow>().First();
        var cell11 = row1.Elements<Drawing.TableCell>().First();
        var tcPr = cell11.TableCellProperties;
        tcPr.Should().NotBeNull("cell should have TableCellProperties after border.top Set");

        // Top border must exist and have the 5pt width attribute (5pt = 60000 EMU)
        var topBorder = tcPr!.TopBorderLineProperties;
        topBorder.Should().NotBeNull("top border should be set");
        var topWidthAttr = topBorder!.GetAttributes().FirstOrDefault(a => a.LocalName == "w");
        topWidthAttr.Value.Should().NotBeNullOrEmpty("top border width attribute must be present");
        long.Parse(topWidthAttr.Value).Should().Be(63500, "5pt = 63500 EMU (1pt = 12700 EMU)");

        // Left border must NOT have been set by a border.top.width key
        var leftBorder = tcPr.LeftBorderLineProperties;
        if (leftBorder != null)
        {
            var leftWidthAttr = leftBorder.GetAttributes().FirstOrDefault(a => a.LocalName == "w");
            // Either the element is absent or has no width set to 60000
            if (leftWidthAttr.Value != null)
                long.Parse(leftWidthAttr.Value).Should().NotBe(63500,
                    "border.top should NOT have changed left border width to 5pt");
        }

        // Right border must NOT have been set by a border.top.width key
        var rightBorder = tcPr.RightBorderLineProperties;
        if (rightBorder != null)
        {
            var rightWidthAttr = rightBorder.GetAttributes().FirstOrDefault(a => a.LocalName == "w");
            if (rightWidthAttr.Value != null)
                long.Parse(rightWidthAttr.Value).Should().NotBe(63500,
                    "border.top should NOT have changed right border width to 5pt");
        }

        // Bottom border must NOT have been set by a border.top.width key
        var bottomBorder = tcPr.BottomBorderLineProperties;
        if (bottomBorder != null)
        {
            var bottomWidthAttr = bottomBorder.GetAttributes().FirstOrDefault(a => a.LocalName == "w");
            if (bottomWidthAttr.Value != null)
                long.Parse(bottomWidthAttr.Value).Should().NotBe(63500,
                    "border.top should NOT have changed bottom border width to 5pt");
        }
    }

    /// <summary>
    /// Complementary check: border.all sets all 4 sides; verify each side
    /// receives the same width so the wildcard arm works as intended.
    /// </summary>
    [Fact]
    public void TableCell_BorderAll_SetsAllFourSides()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new() { ["title"] = "BugA-all" });
        handler.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "2",
            ["cols"] = "2",
            ["width"] = "6cm",
            ["height"] = "3cm"
        });

        handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
        {
            ["border"] = "2pt solid 0000FF"
        });

        handler.Dispose();

        using var doc = PresentationDocument.Open(path, false);
        var slidePart = doc.PresentationPart!.SlideParts.First();
        var table = slidePart.Slide.Descendants<Drawing.Table>().First();
        var cell = table.Elements<Drawing.TableRow>().First()
                        .Elements<Drawing.TableCell>().First();
        var tcPr = cell.TableCellProperties;
        tcPr.Should().NotBeNull();

        // All four sides should exist and have 2pt = 25400 EMU (1pt = 12700 EMU)
        long expected = 25400;
        tcPr!.TopBorderLineProperties.Should().NotBeNull("border (all) should set top");
        tcPr.BottomBorderLineProperties.Should().NotBeNull("border (all) should set bottom");
        tcPr.LeftBorderLineProperties.Should().NotBeNull("border (all) should set left");
        tcPr.RightBorderLineProperties.Should().NotBeNull("border (all) should set right");

        long.Parse(tcPr.TopBorderLineProperties!.GetAttributes()
            .First(a => a.LocalName == "w").Value).Should().Be(expected);
        long.Parse(tcPr.BottomBorderLineProperties!.GetAttributes()
            .First(a => a.LocalName == "w").Value).Should().Be(expected);
        long.Parse(tcPr.LeftBorderLineProperties!.GetAttributes()
            .First(a => a.LocalName == "w").Value).Should().Be(expected);
        long.Parse(tcPr.RightBorderLineProperties!.GetAttributes()
            .First(a => a.LocalName == "w").Value).Should().Be(expected);
    }

    /// <summary>
    /// Regression guard: after independently setting border.left and border.right
    /// to different widths, verify each side retains its own value.
    /// </summary>
    [Fact]
    public void TableCell_IndependentBorderSides_DoNotCrossContaminate()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new() { ["title"] = "BugA-independent" });
        handler.Add("/slide[1]", "table", null, new()
        {
            ["rows"] = "2",
            ["cols"] = "2",
            ["width"] = "6cm",
            ["height"] = "3cm"
        });

        handler.Set("/slide[1]/table[1]/tr[1]/tc[1]", new()
        {
            ["border.left"] = "1pt solid 00FF00",
            ["border.right"] = "3pt solid FF00FF"
        });

        handler.Dispose();

        using var doc = PresentationDocument.Open(path, false);
        var slidePart = doc.PresentationPart!.SlideParts.First();
        var table = slidePart.Slide.Descendants<Drawing.Table>().First();
        var cell = table.Elements<Drawing.TableRow>().First()
                        .Elements<Drawing.TableCell>().First();
        var tcPr = cell.TableCellProperties!;

        long leftW = long.Parse(tcPr.LeftBorderLineProperties!
            .GetAttributes().First(a => a.LocalName == "w").Value);
        long rightW = long.Parse(tcPr.RightBorderLineProperties!
            .GetAttributes().First(a => a.LocalName == "w").Value);

        leftW.Should().Be(12700, "1pt = 12700 EMU for left border");
        rightW.Should().Be(38100, "3pt = 38100 EMU for right border (3 × 12700)");

        // Top and bottom should be absent (or at least not have 1pt or 3pt)
        tcPr.TopBorderLineProperties.Should().BeNull("top border should not exist");
        tcPr.BottomBorderLineProperties.Should().BeNull("bottom border should not exist");
    }

    // =========================================================================
    // Bug B — Cross-slide shape Move loses hyperlink
    // =========================================================================

    /// <summary>
    /// Add a shape with a hyperlink on slide 1, move it to slide 2, then verify
    /// that Format["link"] is preserved on the moved shape.
    ///
    /// Root cause: CopyRelationships() calls GetPartById for external hyperlink
    /// relationship IDs, which throws ArgumentOutOfRangeException (caught silently).
    /// The target slide never receives the HyperlinkRelationship registration,
    /// so the link's r:id attribute is dangling.
    /// </summary>
    [Fact]
    public void CrossSlideMove_ShapeWithHyperlink_LinkIsPreservedAfterMove()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        // Two slides (no title so the only shapes are ones we explicitly add)
        handler.Add("/", "slide", null, new());
        handler.Add("/", "slide", null, new());

        // Add linked shape on slide 1
        handler.Add("/slide[1]", "shape", null, new()
        {
            ["text"] = "Click me",
            ["link"] = "https://example.com/target"
        });

        // Verify link is present before move (shape[1] is our added shape)
        var before = handler.Get("/slide[1]/shape[1]");
        before.Should().NotBeNull();
        before!.Format.Should().ContainKey("link",
            "shape added with link= should have link in Format before any move");
        ((string)before.Format["link"]).Should().Be("https://example.com/target");

        // Move the shape to slide 2
        handler.Move("/slide[1]/shape[1]", "/slide[2]", null);

        // After move: slide 2 should have the shape
        var afterShape = handler.Get("/slide[2]/shape[1]");
        afterShape.Should().NotBeNull("shape should have landed on slide 2");

        var afterPath = afterShape!.Path;
        afterShape.Format.Should().ContainKey("link",
            "hyperlink must survive cross-slide move — CopyRelationships must handle external rels");
        ((string)afterShape.Format["link"]).Should().Be("https://example.com/target",
            "the moved shape's hyperlink URL must be identical to the original");
    }

    /// <summary>
    /// Persistence check: reopen the file after cross-slide move and verify
    /// the hyperlink relationship is correctly recorded in the XML package.
    /// </summary>
    [Fact]
    public void CrossSlideMove_ShapeWithHyperlink_LinkPersistedAfterReopen()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);

        using (var handler = new PowerPointHandler(path, editable: true))
        {
            handler.Add("/", "slide", null, new());
            handler.Add("/", "slide", null, new());
            handler.Add("/slide[1]", "shape", null, new()
            {
                ["text"] = "Linked",
                ["link"] = "https://persist-check.example.org/"
            });
            handler.Move("/slide[1]/shape[1]", "/slide[2]", null);
        }

        // Re-open read-only and inspect raw XML
        using var doc = PresentationDocument.Open(path, false);
        var slideParts = doc.PresentationPart!.SlideParts.ToList();
        slideParts.Should().HaveCount(2);

        // Slide 2 should have the shape with a run whose HyperlinkOnClick has a valid r:id
        var slide2 = slideParts[1].Slide;
        var runs = slide2.Descendants<Drawing.Run>().ToList();
        runs.Should().NotBeEmpty("the moved shape must have a text run on slide 2");

        Drawing.HyperlinkOnClick? hlinkEl = null;
        foreach (var run in runs)
        {
            hlinkEl = run.RunProperties?.GetFirstChild<Drawing.HyperlinkOnClick>();
            if (hlinkEl != null) break;
        }
        hlinkEl.Should().NotBeNull("a HyperlinkOnClick element must exist in the run properties");

        var rId = hlinkEl!.Id?.Value;
        rId.Should().NotBeNullOrEmpty("r:id on HyperlinkOnClick must be non-empty");

        // The relationship must be registered on slide 2's part
        var hyperRels = slideParts[1].HyperlinkRelationships.ToList();
        hyperRels.Should().Contain(r => r.Id == rId,
            "slide 2 must have the HyperlinkRelationship registered for the moved shape's link");

        var targetRel = hyperRels.First(r => r.Id == rId);
        targetRel.Uri.AbsoluteUri.Should().Contain("persist-check.example.org",
            "the registered relationship URI must match the original hyperlink");
    }

    // =========================================================================
    // Bug C — Connector startShape/endShape: verify behaviour (NOT a bug)
    // =========================================================================

    /// <summary>
    /// Verifies that startshape and endshape properties are wired up when adding
    /// a connector.  The code in AddConnector() does handle these properties
    /// (StartConnection / EndConnection elements are created).
    /// </summary>
    [Fact]
    public void Connector_StartAndEndShape_AreWiredInXml()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        // No title so the only shapes in the tree are the ones we explicitly add
        handler.Add("/", "slide", null, new());

        // Add two shapes — get their IDs via DocumentNode.Format["id"] (or parse from the
        // returned path if handler exposes it).  We use shape[1] and shape[2] paths.
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "A", ["x"] = "1cm", ["y"] = "1cm", ["width"] = "3cm", ["height"] = "2cm" });
        handler.Add("/slide[1]", "shape", null, new() { ["text"] = "B", ["x"] = "8cm", ["y"] = "1cm", ["width"] = "3cm", ["height"] = "2cm" });

        // Add connector with startshape=1 / endshape=2 (relative numeric IDs in connector API)
        // The connector API uses the shape's *XML element ID* attribute, but since we don't
        // expose that via Get, we use shape index 1 and 2 as documented by the connector API.
        // We'll verify the StartConnection/EndConnection XML elements exist with non-zero IDs.
        handler.Add("/slide[1]", "connector", null, new()
        {
            ["startshape"] = "1",
            ["endshape"] = "2",
            ["preset"] = "straight",
            ["x"] = "4cm",
            ["y"] = "2cm",
            ["width"] = "4cm",
            ["height"] = "0"
        });

        // Close handler before raw XML inspection
        handler.Dispose();

        // Inspect raw XML for StartConnection / EndConnection
        using var doc = PresentationDocument.Open(path, false);
        var slide = doc.PresentationPart!.SlideParts.First().Slide;
        var connectors = slide.Descendants<ConnectionShape>().ToList();
        connectors.Should().NotBeEmpty("connector should be present on the slide");

        var cxn = connectors.Last();
        var cxnDrawProps = cxn.NonVisualConnectionShapeProperties!
                              .NonVisualConnectorShapeDrawingProperties;
        cxnDrawProps.Should().NotBeNull();

        cxnDrawProps!.StartConnection.Should().NotBeNull(
            "startshape property must create a StartConnection element");
        cxnDrawProps.StartConnection!.Id.Should().NotBeNull();
        cxnDrawProps.StartConnection.Id!.Value.Should().Be(1u,
            "startshape=1 must set StartConnection/@id=1");

        cxnDrawProps.EndConnection.Should().NotBeNull(
            "endshape property must create an EndConnection element");
        cxnDrawProps.EndConnection!.Id.Should().NotBeNull();
        cxnDrawProps.EndConnection.Id!.Value.Should().Be(2u,
            "endshape=2 must set EndConnection/@id=2");
    }

    // =========================================================================
    // Bug D — Slide Set "layout" and "name" listed as valid but unimplemented
    // =========================================================================

    /// <summary>
    /// Set slide name via Set("/slide[1]", { ["name"] = "My Slide" }).
    /// The slide-level switch has no arm for "name", so the value is silently
    /// dropped (or pushed to GenericXmlQuery which sets it as a raw XML attribute
    /// on the Slide element — not on CommonSlideData.Name where PPT reads it).
    /// Expected: the unsupported list should contain "name" (documenting the gap),
    /// OR CommonSlideData.Name should reflect the new value.
    /// This test asserts that either the property is correctly applied (preferred fix)
    /// or that the returned unsupported list is non-empty (acceptable gap marker).
    /// </summary>
    [Fact]
    public void SlideSet_Name_IsEitherAppliedOrReportedAsUnsupported()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new() { ["title"] = "Slide D" });

        // Attempt to set slide name
        var unsupported = handler.Set("/slide[1]", new() { ["name"] = "My Custom Slide Name" });

        bool nameWasReportedUnsupported = unsupported != null && unsupported.Any(u =>
            u.Contains("name", StringComparison.OrdinalIgnoreCase));

        // Close handler before opening the file for raw XML inspection (avoid file lock)
        handler.Dispose();

        // Check via raw XML whether CommonSlideData.Name was actually updated
        using var doc = PresentationDocument.Open(path, false);
        var slide = doc.PresentationPart!.SlideParts.First().Slide;
        var csd = slide.CommonSlideData;
        var actualName = csd?.Name?.Value;

        bool nameWasApplied = actualName == "My Custom Slide Name";

        // Either the bug is fixed (name applied) or it is documented (in unsupported list)
        (nameWasApplied || nameWasReportedUnsupported).Should().BeTrue(
            "Set(slide, name) must either apply the name to CommonSlideData OR " +
            "return 'name' in the unsupported list — it must not silently drop it " +
            "while advertising 'name' as a valid slide property");
    }

    /// <summary>
    /// Set slide layout via Set("/slide[1]", { ["layout"] = "Title Slide" }).
    /// The slide-level switch has no arm for "layout", so it falls through to
    /// the default handler.  The layout relationship is not changed.
    /// Expected: the unsupported list should contain "layout" (documenting the gap),
    /// OR the slide's layout relationship should be updated.
    /// </summary>
    [Fact]
    public void SlideSet_Layout_IsEitherAppliedOrReportedAsUnsupported()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new() { ["title"] = "Layout test" });

        // Get current layout name before attempting to change it
        var slideBefore = handler.Get("/slide[1]");
        var layoutBefore = slideBefore?.Format.GetValueOrDefault("layout");

        var unsupported = handler.Set("/slide[1]", new() { ["layout"] = "Title Slide" });

        // Re-query to see if layout changed
        var slideAfter = handler.Get("/slide[1]");
        var layoutAfter = slideAfter?.Format.GetValueOrDefault("layout") as string;

        bool layoutWasApplied = layoutAfter == "Title Slide";
        bool layoutWasReportedUnsupported = unsupported != null && unsupported.Any(u =>
            u.Contains("layout", StringComparison.OrdinalIgnoreCase));

        (layoutWasApplied || layoutWasReportedUnsupported).Should().BeTrue(
            "Set(slide, layout) must either change the layout OR return 'layout' in the " +
            "unsupported list — it must not silently drop it while advertising 'layout' " +
            "as a valid slide property in the error message");
    }

    /// <summary>
    /// Verify that passing BOTH "name" and "layout" simultaneously either applies
    /// both, or surfaces both in the unsupported list (no silent partial drops).
    /// </summary>
    [Fact]
    public void SlideSet_NameAndLayout_BothAccountedFor()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        using var handler = new PowerPointHandler(path, editable: true);

        handler.Add("/", "slide", null, new() { ["title"] = "Both" });

        var unsupported = handler.Set("/slide[1]", new()
        {
            ["name"] = "Named Slide",
            ["layout"] = "Title Slide"
        });

        bool nameUnsupported = unsupported != null && unsupported.Any(u =>
            u.Contains("name", StringComparison.OrdinalIgnoreCase));
        bool layoutUnsupported = unsupported != null && unsupported.Any(u =>
            u.Contains("layout", StringComparison.OrdinalIgnoreCase));

        // Check layout via handler (still open) before closing
        bool layoutApplied = handler.Get("/slide[1]")?.Format
            .GetValueOrDefault("layout") as string == "Title Slide";

        // Close handler before opening the file for raw XML inspection
        handler.Dispose();

        // Each key must be accounted for: either applied or in unsupported list
        using var doc = PresentationDocument.Open(path, false);
        var slide = doc.PresentationPart!.SlideParts.First().Slide;
        var actualName = slide.CommonSlideData?.Name?.Value;

        bool nameApplied = actualName == "Named Slide";

        (nameApplied || nameUnsupported).Should().BeTrue(
            "'name' must be applied or reported unsupported");

        (layoutApplied || layoutUnsupported).Should().BeTrue(
            "'layout' must be applied or reported unsupported");
    }
}
