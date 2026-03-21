// Bug hunt tests Part 6: Bugs #461-490
// Focus: PPTX lineSpacing multiplier, fill type conflict (BlipFill leak),
//        PPTX/Word font-size integer division truncation,
//        EmuConverter "rem" unit error message, SetRange double reorder,
//        Word GetRunFont EastAsia-first priority

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Wordprocessing;
using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Tests.Functional;

public partial class BugHuntTests
{
    // ==================== BUG #461 (CRITICAL — FIXED): PPTX lineSpacing multiplier was ×1000, now ×100000 ====================
    // Help docs say "lineSpacing  Line spacing multiplier (e.g. 1.5 for 150%)"
    // Was: new SpacingPercent { Val = (int)(double.Parse(value) * 1000) }
    //   → "1.5" → 1.5 × 1000 = 1500 (= 1.5%, not 150%)
    // Fixed: 1.5 × 100000 = 150000 (= 150%)
    // Confirmed via Apache POI XDDFSpacingPercent: 100% = 100000 (1/1000th of a percent unit)
    //
    // Location: PowerPointHandler.ShapeProperties.cs — case "linespacing" or "line.spacing"

    [Fact]
    public void Bug461_Pptx_LineSpacing_MultiplierStoredAs1Pct_NotAs150Pct()
    {
        // 1. Create
        var path = Path.Combine(Path.GetTempPath(), $"bug461_{Guid.NewGuid():N}.pptx");
        try
        {
            BlankDocCreator.Create(path);
            using var handler = new PowerPointHandler(path, editable: true);

            // 2. Add slide + shape
            handler.Add("/", "slide", null, new() { ["title"] = "Title" });
            handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Hello" });

            // 3. Set lineSpacing to 1.5 (documentation says this means 150%)
            handler.Set("/slide[1]/shape[2]", new() { ["lineSpacing"] = "1.5" });

            // 4. Get — readback looks correct due to reciprocal division (/1000)
            var node = handler.Get("/slide[1]/shape[2]");
            node.Format.Should().ContainKey("lineSpacing");
            node.Format["lineSpacing"]?.ToString().Should().Be("1.5x");

            // 5. Verify actual XML value — this exposes the bug
            var rawXml = handler.Raw("/slide[1]");

            // BUG: The stored val is 1500 (representing 1.5%), not 150000 (representing 150%)
            // Correct: val="150000" for 1.5× line spacing
            // Actual: val="1500" for 1.5% line spacing
            rawXml.Should().Contain("val=\"150000\"",
                "lineSpacing=1.5 should store SpacingPercent val=150000 (150%), " +
                "but the code does value×1000 = 1500 instead of value×100000 = 150000");
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    // ==================== BUG #462 (HIGH): PPTX lineSpacing read-back roundtrip hides bug ====================
    // The readback also divides by 1000, so the roundtrip "1.5" → stores 1500 → reads "1.5" is consistent
    // but the actual OOXML semantics are wrong.  A correct implementation would:
    //   Set: value × 100000 (e.g. 1.5 → 150000)
    //   Read: value / 100000 (e.g. 150000 → 1.5)
    //
    // This test confirms the symmetrical bug by checking that a roundtrip of "2.0" (double-spacing)
    // actually results in recognizably wrong XML (val="2000" = 2%, not "200000" = 200%).

    [Fact]
    public void Bug462_Pptx_LineSpacing_DoubleSpacingStoredAs2Pct()
    {
        var path = Path.Combine(Path.GetTempPath(), $"bug462_{Guid.NewGuid():N}.pptx");
        try
        {
            BlankDocCreator.Create(path);
            using var handler = new PowerPointHandler(path, editable: true);
            handler.Add("/", "slide", null, new() { ["title"] = "Title" });
            handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Double" });

            // "2.0" = 200% = double spacing per documentation
            handler.Set("/slide[1]/shape[2]", new() { ["lineSpacing"] = "2.0" });

            var rawXml = handler.Raw("/slide[1]");
            // BUG: val="2000" (2% spacing), should be val="200000" (200% spacing)
            rawXml.Should().Contain("val=\"200000\"",
                "lineSpacing=2.0 should store 200000 (200% = double spacing), " +
                "not 2000 (2% = nearly zero spacing)");
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    // ==================== BUG #463 (HIGH): PPTX font size readback truncates fractional points ====================
    // When reading back font size in ShapeToNode (NodeBuilder.cs):
    //   node.Format["size"] = $"{fontSize.Value / 100}pt"
    // This uses integer division! So 11.5pt (stored as 1150 hundredths-of-pt) → 1150/100 = 11 → "11pt"
    //
    // Location: PowerPointHandler.NodeBuilder.cs — ShapeToNode and RunToNode

    [Fact]
    public void Bug463_Pptx_FontSize_FractionalPtTruncatedOnReadback()
    {
        var path = Path.Combine(Path.GetTempPath(), $"bug463_{Guid.NewGuid():N}.pptx");
        try
        {
            BlankDocCreator.Create(path);
            using var handler = new PowerPointHandler(path, editable: true);
            handler.Add("/", "slide", null, new() { ["title"] = "Title" });
            handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Test", ["size"] = "11.5" });

            var node = handler.Get("/slide[1]/shape[2]");
            // BUG: returns "11pt" due to integer division (1150 / 100 = 11)
            // Expected: "11.5pt"
            node.Format["size"]?.ToString().Should().Be("11.5pt",
                "font size 11.5pt should round-trip correctly, " +
                "but integer division truncates it to 11pt");
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    // ==================== BUG #464 (HIGH): Word font size readback truncates half-points ====================
    // In WordHandler.Helpers.cs GetRunFontSize:
    //   return $"{int.Parse(size) / 2}pt"; // stored as half-points
    // Integer division: 23 half-points (= 11.5pt) / 2 = 11 → "11pt" (wrong!)
    //
    // Location: WordHandler.Helpers.cs — GetRunFontSize

    [Fact]
    public void Bug464_Word_FontSize_HalfPointTruncatedOnReadback()
    {
        // 1. Create
        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "Test" });

        // 2. Set font size to 11.5pt (stored as 23 half-points)
        _wordHandler.Set("/body/p[1]", new() { ["size"] = "11.5" });

        // 3. Get + verify
        // The run should report "11.5pt" but due to integer division it returns "11pt"
        var node = _wordHandler.Get("/body/p[1]");
        node.Children.Should().NotBeEmpty();
        var runNode = node.Children[0];
        // BUG: returns "11pt" because int.Parse("23") / 2 = 11 (integer truncation)
        // Expected: "11.5pt"
        runNode.Format["size"]?.ToString().Should().Be("11.5pt",
            "font size 11.5pt should round-trip correctly, " +
            "but GetRunFontSize uses integer division truncating 23/2=11 instead of 11.5");
    }

    // ==================== BUG #465 (HIGH): ApplyShapeFill does NOT remove BlipFill (image fill) ====================
    // When a shape has an image fill (BlipFill) and Set(fill="FF0000") is called,
    // ApplyShapeFill removes SolidFill/NoFill/GradientFill/PatternFill but NOT BlipFill.
    // This leaves the shape with both BlipFill AND SolidFill, causing visual inconsistency.
    //
    // Location: PowerPointHandler.Fill.cs — ApplyShapeFill

    [Fact]
    public void Bug465_Pptx_ApplyShapeFill_DoesNotRemoveBlipFill()
    {
        var path = Path.Combine(Path.GetTempPath(), $"bug465_{Guid.NewGuid():N}.pptx");
        var imgPath = Path.Combine(Path.GetTempPath(), $"bug465_{Guid.NewGuid():N}.png");
        try
        {
            // Create a minimal 1×1 white PNG
            var pngBytes = new byte[]
            {
                0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A, // PNG signature
                0x00,0x00,0x00,0x0D,                     // IHDR length=13
                0x49,0x48,0x44,0x52,                     // "IHDR"
                0x00,0x00,0x00,0x01,                     // width=1
                0x00,0x00,0x00,0x01,                     // height=1
                0x08,0x02,0x00,0x00,0x00,               // 8-bit RGB, no interlace
                0x90,0x77,0x53,0xDE,                     // CRC
                0x00,0x00,0x00,0x0C,                     // IDAT length=12
                0x49,0x44,0x41,0x54,                     // "IDAT"
                0x08,0xD7,0x63,0xF8,0xFF,0xFF,0x3F,0x00,0x05,0xFE,0x02,0xFE, // compressed data
                0xDC,0xCC,0x59,0xE7,                     // CRC
                0x00,0x00,0x00,0x00,                     // IEND length=0
                0x49,0x45,0x4E,0x44,                     // "IEND"
                0xAE,0x42,0x60,0x82                      // CRC
            };
            File.WriteAllBytes(imgPath, pngBytes);

            BlankDocCreator.Create(path);
            using var handler = new PowerPointHandler(path, editable: true);
            handler.Add("/", "slide", null, new() { ["title"] = "Title" });
            handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Test" });

            // 1. Apply image fill
            handler.Set("/slide[1]/shape[2]", new() { ["image"] = imgPath });

            // Verify image fill was set
            var rawAfterImage = handler.Raw("/slide[1]");
            rawAfterImage.Should().Contain("blipFill", "image fill should be applied");

            // 2. Now apply a solid fill — this should REPLACE the image fill
            handler.Set("/slide[1]/shape[2]", new() { ["fill"] = "FF0000" });

            // 3. Check raw XML — BUG: blipFill is NOT removed by ApplyShapeFill
            var rawAfterSolid = handler.Raw("/slide[1]");

            // BUG: both blipFill and solidFill exist simultaneously
            rawAfterSolid.Should().NotContain("blipFill",
                "ApplyShapeFill should remove BlipFill when replacing with solid fill, " +
                "but it only removes SolidFill/NoFill/GradientFill/PatternFill, leaving BlipFill intact");
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
            if (File.Exists(imgPath)) File.Delete(imgPath);
        }
    }

    // ==================== BUG #466 (HIGH): ApplyGradientFill does NOT remove BlipFill ====================
    // Same root cause as Bug #465 but for gradient fill.
    // ApplyGradientFill removes SolidFill/NoFill/GradientFill but NOT BlipFill.
    //
    // Location: PowerPointHandler.Fill.cs — ApplyGradientFill

    [Fact]
    public void Bug466_Pptx_ApplyGradientFill_DoesNotRemoveBlipFill()
    {
        var path = Path.Combine(Path.GetTempPath(), $"bug466_{Guid.NewGuid():N}.pptx");
        var imgPath = Path.Combine(Path.GetTempPath(), $"bug466_{Guid.NewGuid():N}.png");
        try
        {
            var pngBytes = new byte[]
            {
                0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A,
                0x00,0x00,0x00,0x0D,0x49,0x48,0x44,0x52,
                0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x01,
                0x08,0x02,0x00,0x00,0x00,0x90,0x77,0x53,0xDE,
                0x00,0x00,0x00,0x0C,0x49,0x44,0x41,0x54,
                0x08,0xD7,0x63,0xF8,0xFF,0xFF,0x3F,0x00,0x05,0xFE,0x02,0xFE,
                0xDC,0xCC,0x59,0xE7,
                0x00,0x00,0x00,0x00,0x49,0x45,0x4E,0x44,0xAE,0x42,0x60,0x82
            };
            File.WriteAllBytes(imgPath, pngBytes);

            BlankDocCreator.Create(path);
            using var handler = new PowerPointHandler(path, editable: true);
            handler.Add("/", "slide", null, new() { ["title"] = "Title" });
            handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Test" });

            // 1. Apply image fill
            handler.Set("/slide[1]/shape[2]", new() { ["image"] = imgPath });

            // 2. Apply gradient fill — should replace image fill
            handler.Set("/slide[1]/shape[2]", new() { ["gradient"] = "FF0000-0000FF" });

            // 3. Verify blipFill is removed
            var rawXml = handler.Raw("/slide[1]");

            // BUG: blipFill persists alongside gradFill
            rawXml.Should().NotContain("blipFill",
                "ApplyGradientFill should remove BlipFill when replacing with gradient fill, " +
                "but it only removes SolidFill/NoFill/GradientFill, not BlipFill");
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
            if (File.Exists(imgPath)) File.Delete(imgPath);
        }
    }

    // ==================== BUG #467 (LOW): EmuConverter "rem" unit error says "em" not "rem" ====================
    // In EmuConverter.HasKnownUnitSuffix, "5rem".EndsWith("em") == true,
    // so unit is reported as "em" even though input is "rem".
    // Error: "Unsupported unit 'em' in dimension value '5rem'." (should say 'rem')
    //
    // Location: EmuConverter.cs — HasKnownUnitSuffix

    [Fact]
    public void Bug467_EmuConverter_RemUnitReportsEmInErrorMessage()
    {
        var ex = Assert.Throws<ArgumentException>(() => EmuConverter.ParseEmu("5rem"));
        // BUG: error message says 'em' not 'rem' because "5rem".EndsWith("em") == true
        ex.Message.Should().Contain("rem",
            "error message for '5rem' should mention 'rem', not 'em', " +
            "but HasKnownUnitSuffix matches EndsWith('em') which is a suffix of 'rem'");
        ex.Message.Should().NotContain("'em'",
            "the error should say 'rem' not 'em'");
    }

    // ==================== BUG #468 (LOW): SetRange calls ReorderWorksheetChildren twice ====================
    // In ExcelHandler.Set.cs SetRange method, line 707-708:
    //   ReorderWorksheetChildren(ws);
    //   ReorderWorksheetChildren(ws); ws.Save();
    // The second call is redundant (the method is idempotent but unnecessary).
    // This is a code error / dead code that wastes time on every range Set operation.
    //
    // Location: ExcelHandler.Set.cs — SetRange (lines 707-708)
    // Note: This test verifies the functionality works correctly (no exceptions)
    //       but cannot directly observe the double call from outside.

    [Fact]
    public void Bug468_Excel_SetRange_DoubleReorderCallDoesNotCorruptFile()
    {
        // 1. Set values in a range
        _excelHandler.Set("/Sheet1/A1", new() { ["value"] = "1" });
        _excelHandler.Set("/Sheet1/A2", new() { ["value"] = "2" });
        _excelHandler.Set("/Sheet1/B1", new() { ["value"] = "3" });
        _excelHandler.Set("/Sheet1/B2", new() { ["value"] = "4" });

        // 2. Merge range — this triggers SetRange → double ReorderWorksheetChildren
        _excelHandler.Set("/Sheet1/A1:B2", new() { ["merge"] = "true" });

        // 3. Reopen and verify merge is preserved and file is not corrupted
        ReopenExcel();
        var cell = _excelHandler.Get("/Sheet1/A1");
        cell.Should().NotBeNull("file should be valid after double ReorderWorksheetChildren");

        // Note: The bug is the redundant call in SetRange (lines 707-708):
        //   ReorderWorksheetChildren(ws);       ← call 1
        //   ReorderWorksheetChildren(ws); ws.Save();  ← call 2 (redundant!)
        // A simple source-level verification:
        var setRangeSource = System.IO.File.ReadAllText(
            Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "..",
            "src", "officecli", "Handlers", "Excel", "ExcelHandler.Set.cs"));
        // Count occurrences in the SetRange method region (rough heuristic):
        var doubleCallPattern = "ReorderWorksheetChildren(ws);\r\n        ReorderWorksheetChildren(ws);";
        var doubleCallPatternUnix = "ReorderWorksheetChildren(ws);\n        ReorderWorksheetChildren(ws);";
        var hasDoubleCall = setRangeSource.Contains(doubleCallPattern) ||
                            setRangeSource.Contains(doubleCallPatternUnix);
        hasDoubleCall.Should().BeFalse(
            "SetRange should not call ReorderWorksheetChildren twice — the second call at line 708 is redundant");
    }

    // ==================== BUG #469 (MEDIUM): Word GetRunFont returns EastAsia font before Ascii ====================
    // In WordHandler.Helpers.cs GetRunFont (line 126):
    //   return fonts?.EastAsia?.Value ?? fonts?.Ascii?.Value ?? fonts?.HighAnsi?.Value;
    // EastAsia is checked FIRST. For a Western document where Ascii="Arial" but
    // EastAsia="Microsoft YaHei", the function incorrectly returns "Microsoft YaHei".
    //
    // Location: WordHandler.Helpers.cs — GetRunFont

    [Fact]
    public void Bug469_Word_GetRunFont_ReturnsEastAsiaBeforeAscii()
    {
        // 1. Add paragraph + run with font set via public API
        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "Test" });
        _wordHandler.Set("/body/p[1]", new() { ["font"] = "Arial" });

        // At this point: RunFonts.Ascii = "Arial", RunFonts.EastAsia = "Arial", RunFonts.HighAnsi = "Arial"
        // So GetRunFont returns "Arial" from EastAsia (correct result, but for wrong reason)

        // 2. Now use RawSet to change only the EastAsia font to a different value
        // This simulates a document created by Word where Latin and CJK fonts differ
        _wordHandler.RawSet("/document",
            "//w:r[1]/w:rPr/w:rFonts",
            "setattr",
            "w:eastAsia=Microsoft YaHei");

        // 3. Read back the font — should return Ascii ("Arial"), not EastAsia ("Microsoft YaHei")
        ReopenWord();
        var node = _wordHandler.Get("/body/p[1]");
        node.Children.Should().NotBeEmpty();
        var runNode = node.Children[0];

        // BUG: returns "Microsoft YaHei" (EastAsia) instead of "Arial" (Ascii)
        // because GetRunFont checks EastAsia first
        runNode.Format["font"]?.ToString().Should().Be("Arial",
            "GetRunFont should return Ascii font for Western text, " +
            "not EastAsia font, but it checks EastAsia before Ascii");
    }

    // ==================== BUG #470 (MEDIUM): Picture Set(path) leaks old ImagePart ====================
    // In PowerPointHandler.Set.cs, case "path" or "src":
    //   var newImgPart = slidePart.AddImagePart(imgType);
    //   using (var stream = File.OpenRead(value)) newImgPart.FeedData(stream);
    //   blip.Embed = slidePart.GetIdOfPart(newImgPart);
    // The old ImagePart is NEVER removed! Every image source replacement adds a new part
    // without cleaning up the previous one, causing file size bloat.
    //
    // Location: PowerPointHandler.Set.cs — case "path" or "src"

    [Fact]
    public void Bug470_Pptx_PictureSetPath_OrphansOldImagePart()
    {
        var path = Path.Combine(Path.GetTempPath(), $"bug470_{Guid.NewGuid():N}.pptx");
        var img1 = Path.Combine(Path.GetTempPath(), $"bug470_img1_{Guid.NewGuid():N}.png");
        var img2 = Path.Combine(Path.GetTempPath(), $"bug470_img2_{Guid.NewGuid():N}.png");
        try
        {
            // Create two different minimal PNGs
            var png1 = new byte[]
            {
                0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A,
                0x00,0x00,0x00,0x0D,0x49,0x48,0x44,0x52,
                0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x01,
                0x08,0x02,0x00,0x00,0x00,0x90,0x77,0x53,0xDE,
                0x00,0x00,0x00,0x0C,0x49,0x44,0x41,0x54,
                0x08,0xD7,0x63,0xF8,0xFF,0xFF,0x3F,0x00,0x05,0xFE,0x02,0xFE,
                0xDC,0xCC,0x59,0xE7,
                0x00,0x00,0x00,0x00,0x49,0x45,0x4E,0x44,0xAE,0x42,0x60,0x82
            };
            // Make img2 slightly different (different content)
            var png2 = (byte[])png1.Clone();
            png2[^5] = 0x00; // slight modification

            File.WriteAllBytes(img1, png1);
            File.WriteAllBytes(img2, png2);

            BlankDocCreator.Create(path);
            using (var handler = new PowerPointHandler(path, editable: true))
            {
                handler.Add("/", "slide", null, new());
                // Add a picture
                handler.Add("/slide[1]", "picture", null, new() { ["path"] = img1 });
            }

            // Count image parts before replacement
            int imagePartsBeforeReplacement;
            using (var pkg = DocumentFormat.OpenXml.Packaging.PresentationDocument.Open(path, false))
            {
                var slideParts = pkg.PresentationPart!.SlideParts.ToList();
                imagePartsBeforeReplacement = slideParts[0].ImageParts.Count();
            }

            // Replace picture source
            using (var handler = new PowerPointHandler(path, editable: true))
            {
                handler.Set("/slide[1]/picture[1]", new() { ["path"] = img2 });
            }

            // Count image parts after replacement — should still be 1
            int imagePartsAfterReplacement;
            using (var pkg = DocumentFormat.OpenXml.Packaging.PresentationDocument.Open(path, false))
            {
                var slideParts = pkg.PresentationPart!.SlideParts.ToList();
                imagePartsAfterReplacement = slideParts[0].ImageParts.Count();
            }

            // BUG: imagePartsAfterReplacement == 2 (old + new), should be 1
            imagePartsAfterReplacement.Should().Be(imagePartsBeforeReplacement,
                "replacing a picture source should remove the old ImagePart, " +
                "but Set(path/src) adds a new ImagePart without cleaning up the old one (resource leak)");
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
            if (File.Exists(img1)) File.Delete(img1);
            if (File.Exists(img2)) File.Delete(img2);
        }
    }

    // ==================== BUG #471 (MEDIUM): PPTX lineSpacing round-trip consistency masks bug ====================
    // Extend Bug #461: verify that Set("1.5") → Get("lineSpacing") round-trips to "1.5"
    // but then a fresh open and manual XML inspection shows the actual stored val is wrong.
    // The read/write are using the same wrong scale (×1000/÷1000), so they agree with each other,
    // but both are wrong vs. the OOXML spec (should be 100000 units per 100%).

    [Fact]
    public void Bug471_Pptx_LineSpacing_RoundTripHidesBug_StoredValIs100xTooSmall()
    {
        var path = Path.Combine(Path.GetTempPath(), $"bug471_{Guid.NewGuid():N}.pptx");
        try
        {
            BlankDocCreator.Create(path);
            using var handler = new PowerPointHandler(path, editable: true);
            handler.Add("/", "slide", null, new() { ["title"] = "Title" });
            handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Test" });

            // Standard spacing
            handler.Set("/slide[1]/shape[2]", new() { ["lineSpacing"] = "1.0" });

            var rawXml = handler.Raw("/slide[1]");

            // BUG: val="1000" (1% spacing), should be val="100000" (100% = single spacing)
            rawXml.Should().Contain("val=\"100000\"",
                "lineSpacing=1.0 should store 100000 (100% = single line spacing), " +
                "not 1000 (1% = nearly invisible spacing). " +
                "Root cause: multiplier is ×1000 but should be ×100000");
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    // ==================== BUG #472 (HIGH): Word color Set() does not strip '#' prefix ====================
    // In WordHandler.Set.cs, color handling at multiple locations:
    //   EnsureRunProperties(run).Color = new Color { Val = value.ToUpperInvariant() };
    // If user provides "#FF0000", it stores "#FF0000" in the XML (invalid — should be "FF0000").
    // PPTX handler correctly uses .TrimStart('#') before storing but Word does not.
    //
    // Location: WordHandler.Set.cs — all "case 'color':" branches

    [Fact]
    public void Bug472_Word_Color_HashPrefixStoredInXml()
    {
        // 1. Add paragraph + run
        _wordHandler.Add("/body", "paragraph", null, new() { ["text"] = "Hello" });

        // 2. Set color with # prefix (as many UI tools provide)
        _wordHandler.Set("/body/p[1]", new() { ["color"] = "#FF0000" });

        // 3. Check raw XML — should NOT contain the # character in the color val attribute
        var rawXml = _wordHandler.Raw("/document");

        // BUG: rawXml contains w:val="#FF0000" instead of w:val="FF0000"
        rawXml.Should().NotContain("\"#FF0000\"",
            "Word color set should strip '#' prefix before storing (like PPTX does), " +
            "but current code uses value.ToUpperInvariant() without TrimStart('#')");

        rawXml.Should().Contain("\"FF0000\"",
            "Color should be stored as 'FF0000' (no # prefix) per OpenXML spec");
    }

    // ==================== BUG #473 (MEDIUM): SplitTransition Direction is hardcoded to 'in' ====================
    // In PowerPointHandler.Animations.cs ApplyTransition(), the "split" case:
    //   "split" => new SplitTransition {
    //       Orientation = ParseOrientation(direction ?? "horizontal"),
    //       Direction = ParseInOutDir("in")  // ← HARDCODED! Never uses user's direction
    //   },
    // So "split-out" doesn't work — Direction is always In, never Out.
    //
    // Location: PowerPointHandler.Animations.cs — ApplyTransition, "split" case

    [Fact]
    public void Bug473_Pptx_SplitTransition_DirectionHardcodedToIn()
    {
        var path = Path.Combine(Path.GetTempPath(), $"bug473_{Guid.NewGuid():N}.pptx");
        try
        {
            BlankDocCreator.Create(path);
            using var handler = new PowerPointHandler(path, editable: true);
            handler.Add("/", "slide", null, new());

            // Set "split-out" transition (should create Direction="out")
            handler.Set("/slide[1]", new() { ["transition"] = "split-out" });

            var rawXml = handler.Raw("/slide[1]");

            // BUG: Direction is hardcoded to "in" regardless of user input
            // The raw XML will show splt with dir="in" not dir="out"
            rawXml.Should().Contain("dir=\"out\"",
                "split-out transition should set Direction=Out, " +
                "but the code hardcodes Direction=ParseInOutDir('in')");
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    // ==================== BUG #474 (LOW): GenericXmlQuery.ParsePathSegments throws FormatException for non-numeric bracket ====================
    // In GenericXmlQuery.cs ParsePathSegments():
    //   var indexStr = part[(bracketIdx + 1)..^1];
    //   segments.Add((name, int.Parse(indexStr)));   // ← throws FormatException, not ArgumentException
    // A path like "/slide[1]/shape[abc]" throws FormatException("Input string was not in a correct format")
    // instead of an informative ArgumentException.
    //
    // Location: GenericXmlQuery.cs — ParsePathSegments (line 231)

    [Fact]
    public void Bug474_GenericXmlQuery_ParsePathSegments_ThrowsFormatExceptionForNonNumericIndex()
    {
        // A path with a non-numeric bracket index should throw a helpful error
        // Currently throws FormatException instead of ArgumentException
        var pptxPath = Path.Combine(Path.GetTempPath(), $"bug474_{Guid.NewGuid():N}.pptx");
        try
        {
            BlankDocCreator.Create(pptxPath);
            using var handler = new PowerPointHandler(pptxPath, editable: true);
            handler.Add("/", "slide", null, new());

            // This should throw a clear ArgumentException about invalid path
            var ex = Assert.ThrowsAny<Exception>(() =>
                handler.Get("/slide[1]/shape[invalid]"));

            // BUG: throws FormatException("Input string was not in a correct format")
            // Expected: ArgumentException with clear message about invalid path index
            ex.Should().BeOfType<ArgumentException>(
                "an invalid path index like 'shape[invalid]' should throw ArgumentException, " +
                "not FormatException which doesn't communicate the root cause");
        }
        finally
        {
            if (File.Exists(pptxPath)) File.Delete(pptxPath);
        }
    }

    // ==================== BUG #476 (HIGH): Excel conditional formatting colors don't strip '#' prefix ====================
    // In ExcelHandler.Add.cs, all conditional formatting color normalization patterns:
    //   (value.Length == 6 ? "FF" : "") + value.ToUpperInvariant()
    // do NOT call TrimStart('#'). So "#FF0000" gets stored as "#FF0000" (length=7, not 6),
    // missing the FF alpha prefix AND retaining the invalid # character.
    // ExcelStyleManager.NormalizeColor correctly uses TrimStart('#') but this path doesn't.
    //
    // Affected: databar, colorscale (min/mid/max), formulacf (font.color, fill)
    // Location: ExcelHandler.Add.cs lines 386, 438-449, 555, 565

    [Fact]
    public void Bug476_Excel_ConditionalFormatting_HashPrefixNotStripped()
    {
        // 1. Add a databar conditional formatting with # prefix color
        _excelHandler.Add("/Sheet1", "databar", null, new()
        {
            ["sqref"] = "A1:A10",
            ["color"] = "#4472C4"  // User provides # prefix
        });

        // 2. Read back raw XML to check what was stored
        var rawXml = _excelHandler.Raw("/Sheet1");

        // BUG: stores "#4472C4" without stripping '#' and without ARGB "FF" prefix
        // Expected: "FF4472C4" (ARGB format)
        rawXml.Should().NotContain("\"#4472C4\"",
            "Excel CF color normalization should strip '#' before storing, " +
            "but uses .Length == 6 check which fails for 7-char '#FF0000' inputs");

        rawXml.Should().Contain("FF4472C4",
            "Excel CF color should be stored as ARGB 'FF4472C4' not '#4472C4'");
    }

    // ==================== BUG #477 (HIGH): Excel colorscale conditional formatting colors don't strip '#' ====================
    // Same root cause as Bug #476 but in colorscale add/set path
    // Location: ExcelHandler.Add.cs lines 438-449

    [Fact]
    public void Bug477_Excel_ColorScale_HashPrefixNotStripped()
    {
        // 1. Add colorscale CF with # prefix colors
        _excelHandler.Add("/Sheet1", "colorscale", null, new()
        {
            ["sqref"] = "B1:B10",
            ["mincolor"] = "#FF0000",
            ["maxcolor"] = "#00FF00"
        });

        // 2. Check raw XML
        var rawXml = _excelHandler.Raw("/Sheet1");

        // BUG: "#FF0000" stored as-is (or "##FF0000" if the "#" gets prefixed wrong)
        // The condition `"#FF0000".Length == 6` is false (length=7), so no "FF" prefix added
        // Result: just "#FF0000" stored instead of "FFFF0000"
        rawXml.Should().NotContain("\"#FF0000\"",
            "ColorScale min/max color should strip '#' and add ARGB prefix, " +
            "but the normalization pattern checks Length==6 before stripping '#'");
    }

    // ==================== BUG #478 (HIGH): Excel CF Set color also missing TrimStart('#') ====================
    // Same root cause as Bug #476/477 but in ExcelHandler.Set.cs for conditional formatting updates.
    // Lines 335, 341, 347: (value.Length == 6 ? "FF" : "") + value.ToUpperInvariant()
    // Location: ExcelHandler.Set.cs — Set /SheetName/cf[N] color/mincolor/maxcolor

    [Fact]
    public void Bug478_Excel_SetConditionalFormatColor_HashPrefixNotStripped()
    {
        // 1. Add a colorscale CF first (with valid colors)
        _excelHandler.Add("/Sheet1", "colorscale", null, new()
        {
            ["sqref"] = "C1:C10",
            ["mincolor"] = "FF0000",
            ["maxcolor"] = "00FF00"
        });

        // 2. Now Set the color via Set() with # prefix
        _excelHandler.Set("/Sheet1/cf[1]", new() { ["mincolor"] = "#0000FF" });

        // 3. Check raw XML
        var rawXml = _excelHandler.Raw("/Sheet1");

        // BUG: "#0000FF" stored as "#0000FF" (length=7, so no "FF" prefix, # not stripped)
        rawXml.Should().NotContain("\"#0000FF\"",
            "Set conditional format color should strip '#' prefix like ExcelStyleManager.NormalizeColor does");

        rawXml.Should().Contain("FF0000FF",
            "Set CF color '#0000FF' should be stored as ARGB 'FF0000FF'");
    }

    // ==================== BUG #479 (HIGH): Excel formulacf font.color missing TrimStart('#') ====================
    // Same root cause, in formulacf (formula-based CF) path.
    // Location: ExcelHandler.Add.cs line 555

    [Fact]
    public void Bug479_Excel_FormulaCF_FontColorHashPrefixNotStripped()
    {
        // 1. Add formula-based CF with # prefix font color
        _excelHandler.Add("/Sheet1", "formulacf", null, new()
        {
            ["sqref"] = "D1:D10",
            ["formula"] = "$D1>100",
            ["font.color"] = "#FF0000"
        });

        // 2. Check styles raw XML (formulacf color goes into dxf in stylesheet)
        var stylesXml = _excelHandler.Raw("/styles");

        // BUG: "#FF0000" stored without # stripped and without FF prefix
        stylesXml.Should().NotContain("\"#FF0000\"",
            "formulacf font.color '#FF0000' should strip '#' and add ARGB prefix");

        stylesXml.Should().Contain("FFFF0000",
            "formulacf font color should be stored as ARGB 'FFFF0000'");
    }

    // ==================== BUG #480 (HIGH): Word table row height uses AppendChild — duplicates on second Set ====================
    // In WordHandler.Set.cs (table row handling):
    //   trPr.AppendChild(new TableRowHeight { Val = uint.Parse(value), ... });
    // AppendChild is used unconditionally — no check for existing TableRowHeight.
    // Setting height twice creates TWO w:trHeight elements in w:trPr, resulting in invalid XML.
    //
    // Location: WordHandler.Set.cs — "height" case in TableRow handling (~line 859)

    [Fact]
    public void Bug480_Word_TableRowHeight_AppendChildDuplicatesOnSecondSet()
    {
        // 1. Add table
        _wordHandler.Add("/body", "table", null, new() { ["rows"] = "2", ["cols"] = "2" });

        // 2. Set row height first time
        _wordHandler.Set("/body/tbl[1]/tr[1]", new() { ["height"] = "500" });

        // 3. Set row height SECOND time with different value
        _wordHandler.Set("/body/tbl[1]/tr[1]", new() { ["height"] = "1000" });

        // 4. Check raw XML for duplicate trHeight elements
        var rawXml = _wordHandler.Raw("/document");

        // BUG: TWO w:trHeight elements exist in the row's trPr
        // Count occurrences of trHeight in the XML
        var count = System.Text.RegularExpressions.Regex.Matches(rawXml, "trHeight").Count;
        count.Should().Be(1,
            "Setting table row height twice should UPDATE the existing w:trHeight, " +
            "not append a second one. Found " + count + " w:trHeight elements — " +
            "AppendChild is used without first removing the existing element");
    }

    // ==================== BUG #481 (MEDIUM): PPTX chart with empty series data throws unhelpful FormatException ====================
    // In PowerPointHandler.Chart.cs ParseSeriesData, if series data is empty after the colon:
    //   "SeriesName:" → Split(',') → [""] → double.Parse("") → FormatException
    // The user gets an unhelpful error: "Input string was not in a correct format"
    // instead of a clear ArgumentException about invalid series data.
    // Same bug exists in ExcelHandler.Helpers.cs ExcelChartParseSeriesData.
    //
    // Location: PowerPointHandler.Chart.cs lines 64-65, ExcelHandler.Helpers.cs lines 476-477

    [Fact]
    public void Bug481_Pptx_Chart_EmptySeriesDataThrowsFormatException()
    {
        var path = Path.Combine(Path.GetTempPath(), $"bug481_{Guid.NewGuid():N}.pptx");
        try
        {
            BlankDocCreator.Create(path);
            using var handler = new PowerPointHandler(path, editable: true);
            handler.Add("/", "slide", null, new());

            // BUG: "SeriesName:" with empty values causes FormatException (unhelpful)
            var ex = Assert.ThrowsAny<Exception>(() =>
                handler.Add("/slide[1]", "chart", null, new()
                {
                    ["type"] = "bar",
                    ["data"] = "Sales:"  // empty values - should give clear error
                }));

            // BUG: throws FormatException instead of ArgumentException
            ex.Should().BeOfType<ArgumentException>(
                "empty series data 'Sales:' should throw ArgumentException with helpful message, " +
                "but double.Parse(\"\") throws FormatException");
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    // ==================== BUG #482 (MEDIUM): Excel chart with empty series data throws unhelpful FormatException ====================
    // Same root cause as Bug #481 but in ExcelHandler.Helpers.cs ExcelChartParseSeriesData.
    // Location: ExcelHandler.Helpers.cs lines 476-477

    [Fact]
    public void Bug482_Excel_Chart_EmptySeriesDataThrowsFormatException()
    {
        // BUG: empty values after colon causes unhelpful FormatException
        var ex = Assert.ThrowsAny<Exception>(() =>
            _excelHandler.Add("/Sheet1", "chart", null, new()
            {
                ["type"] = "column",
                ["data"] = "Revenue:"  // empty values
            }));

        ex.Should().BeOfType<ArgumentException>(
            "empty series data 'Revenue:' should throw ArgumentException with helpful message, " +
            "not FormatException from double.Parse(\"\")");
    }

    // ==================== BUG #483 (HIGH): Excel Set chart title silently fails when chart has no existing title ====================
    // In ExcelHandler.Set.cs, case "title" for chart:
    //   var titleEl = chart.Title;
    //   if (titleEl != null) { /* update */ }
    //   break;  ← exits without adding to unsupported if title is null!
    // When a chart was created without a title, Set(chart[1], {["title"] = "..."})
    // silently does nothing — title is NOT created, and "title" is NOT added to unsupported.
    //
    // Location: ExcelHandler.Set.cs — case "title" in chart Set handler

    [Fact]
    public void Bug483_Excel_SetChartTitle_SilentlyFailsWhenChartHasNoTitle()
    {
        // 1. Add chart WITHOUT title
        _excelHandler.Add("/Sheet1", "chart", null, new()
        {
            ["type"] = "column",
            ["data"] = "Sales:10,20,30"
            // Note: no "title" property — chart is created without a title element
        });

        // 2. Now try to Set the title
        var unsupported = _excelHandler.Set("/Sheet1/chart[1]", new() { ["title"] = "My Chart" });

        // The unsupported list should be empty (title is "supported")
        unsupported.Should().NotContain("title", "title is a supported property");

        // 3. Verify title was actually SET in the chart XML
        var chartXml = _excelHandler.Raw("/Sheet1/chart[1]");

        // BUG: chartXml does NOT contain "My Chart" because the Set silently fails
        // when chart.Title is null (the if block is skipped, no title created)
        chartXml.Should().Contain("My Chart",
            "Setting chart title should create the title element if it doesn't exist, " +
            "but ExcelHandler.Set.cs only updates existing title: 'if (titleEl != null) {...} break'");
    }

    // ==================== BUG #484 (MEDIUM): Excel SetColumn adds Column elements out of schema order ====================
    // In ExcelHandler.Set.cs SetColumn(), when a new Column element is created:
    //   col = new Column { Min = colIdx, Max = colIdx, ... };
    //   columns.AppendChild(col);  ← always appends to end, no sorting
    // Per OOXML spec, Column elements within Columns must be sorted by Min attribute.
    // Setting column B width then column A width creates: [B, A] ordering (wrong).
    // This may cause validation failures or incorrect rendering in strict OOXML readers.
    //
    // Location: ExcelHandler.Set.cs — SetColumn method

    [Fact]
    public void Bug484_Excel_SetColumn_OutOfOrderColumnElements()
    {
        // 1. Set column B width first, then column A
        _excelHandler.Set("/Sheet1/col[B]", new() { ["width"] = "20" });
        _excelHandler.Set("/Sheet1/col[A]", new() { ["width"] = "15" });

        // 2. Check raw XML — columns should be in Min order (A before B)
        var rawXml = _excelHandler.Raw("/Sheet1");

        // BUG: column B (min=2) appears BEFORE column A (min=1) in the XML
        // because SetColumn always uses AppendChild without sorting
        var colAPos = rawXml.IndexOf("min=\"1\"", StringComparison.OrdinalIgnoreCase);
        var colBPos = rawXml.IndexOf("min=\"2\"", StringComparison.OrdinalIgnoreCase);

        colAPos.Should().BeGreaterThan(-1, "column A should be in the XML");
        colBPos.Should().BeGreaterThan(-1, "column B should be in the XML");

        // OOXML spec requires columns sorted by Min: A (min=1) should come before B (min=2)
        colAPos.Should().BeLessThan(colBPos,
            "Column A (min=1) should appear before Column B (min=2) per OOXML spec, " +
            "but SetColumn uses AppendChild which adds B before A when set in reverse order");
    }

    // ==================== BUG #485 (HIGH): PPTX slide Add with title creates duplicate NonVisualDrawingProperties ID=1 ====================
    // In PowerPointHandler.Add.cs, when adding a slide with a title:
    //   ShapeTree's NonVisualGroupShapeProperties → NonVisualDrawingProperties { Id = 1, Name = "" }
    //   CreateTextShape(1, "Title", ...) → NonVisualDrawingProperties { Id = 1, Name = "Title" }
    // BOTH the ShapeTree group AND the title shape have Id=1!
    // Per OOXML spec, each element in a slide must have a unique drawing ID.
    //
    // Location: PowerPointHandler.Add.cs — CreateTextShape(1, ...) called with same id as ShapeTree group

    [Fact]
    public void Bug485_Pptx_AddSlideWithTitle_DuplicateNvPrId()
    {
        var path = Path.Combine(Path.GetTempPath(), $"bug485_{Guid.NewGuid():N}.pptx");
        try
        {
            BlankDocCreator.Create(path);
            using var handler = new PowerPointHandler(path, editable: true);

            // Add slide with title — creates both ShapeTree group (id=1) and title shape (id=1)
            handler.Add("/", "slide", null, new() { ["title"] = "My Title" });

            var rawXml = handler.Raw("/slide[1]");

            // Count occurrences of cNvPr id="1" in the XML
            var count = System.Text.RegularExpressions.Regex.Matches(rawXml, @"id=""1""").Count;

            // BUG: there should be exactly 1 element with id="1" (the ShapeTree group)
            // but actually there are 2: the group AND the title shape
            count.Should().Be(1,
                "Only the ShapeTree group should have id=1. " +
                "CreateTextShape(1, 'Title', ...) creates a title shape with id=1 which duplicates " +
                "the ShapeTree group's id=1. Found " + count + " elements with id=1");
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    // ==================== BUG #475 (HIGH): PPTX RunToNode font size uses integer division ====================
    // In PowerPointHandler.NodeBuilder.cs RunToNode():
    //   node.Format["size"] = $"{fs.Value / 100}pt";
    // C# integer division: 1050 / 100 = 10 (not 10.5), 1150 / 100 = 11 (not 11.5)
    // So any fractional pt font size is truncated on readback.
    // Same issue exists in ShapeToNode for the shape-level font summary.
    //
    // Location: PowerPointHandler.NodeBuilder.cs — RunToNode (and ShapeToNode)

    [Fact]
    public void Bug475_Pptx_RunNode_FontSize_IntDivisionTruncates()
    {
        var path = Path.Combine(Path.GetTempPath(), $"bug475_{Guid.NewGuid():N}.pptx");
        try
        {
            BlankDocCreator.Create(path);
            using var handler = new PowerPointHandler(path, editable: true);
            handler.Add("/", "slide", null, new() { ["title"] = "Title" });
            handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Test" });

            // Set 10.5pt font size
            handler.Set("/slide[1]/shape[2]", new() { ["size"] = "10.5" });

            var node = handler.Get("/slide[1]/shape[2]", depth: 2);
            node.Children.Should().NotBeEmpty("shape should have paragraph children");
            var para = node.Children[0];
            para.Children.Should().NotBeEmpty("paragraph should have run children");
            var run = para.Children[0];

            // BUG: RunToNode uses fs.Value / 100 (integer) → 1050 / 100 = 10 → "10pt" (wrong!)
            run.Format["size"]?.ToString().Should().Be("10.5pt",
                "font size 10.5pt (stored as 1050 hundredths-of-pt) should read back as '10.5pt', " +
                "but integer division 1050/100 = 10 truncates it to '10pt'");
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }
}
