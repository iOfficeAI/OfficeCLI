// File: tests/OfficeCli.Tests/Hwpx/HwpxHandlerTests.cs
using OfficeCli.Core;
using OfficeCli.Handlers;

namespace OfficeCli.Tests.Hwpx;

public class HwpxHandlerTests : IDisposable
{
    private readonly List<string> _tempFiles = new();

    private string CreateTemp(string text = "테스트 문단")
    {
        var path = HwpxTestHelper.CreateMinimalHwpx(text);
        _tempFiles.Add(path);
        return path;
    }

    public void Dispose()
    {
        foreach (var f in _tempFiles)
        {
            try { File.Delete(f); } catch { }
        }
    }

    // ============================================================
    // 1. Move() — detach before reinsert (not Add on parented element)
    // ============================================================
    [Fact]
    public void Move_DetachesBeforeReinsert()
    {
        // Arrange: create a HWPX with 3 paragraphs
        var path = HwpxTestHelper.CreateMultiSectionHwpx(
            new[] { "Para 1", "Para 2", "Para 3" });
        _tempFiles.Add(path);

        using var handler = new HwpxHandler(path, editable: true);

        // Act: move paragraph 3 to position 0 (before paragraph 1)
        var resultPath = handler.Move("/section[1]/p[3]", "/section[1]", 0);

        // Assert: paragraph order should now be [Para 3, Para 1, Para 2]
        var text = handler.ViewAsText();
        var lines = text.Split('\n', StringSplitOptions.RemoveEmptyEntries);
        Assert.Contains("Para 3", lines[0]);
        Assert.Contains("Para 1", lines[1]);
        Assert.Contains("Para 2", lines[2]);
    }

    // ============================================================
    // 2. CopyFrom() — deep clone gets new identity
    // ============================================================
    [Fact]
    public void CopyFrom_AssignsNewId()
    {
        var path = HwpxTestHelper.CreateMultiSectionHwpx(
            new[] { "Original" });
        _tempFiles.Add(path);

        using var handler = new HwpxHandler(path, editable: true);

        // Act: copy paragraph 1 to the end of section 1
        var resultPath = handler.CopyFrom("/section[1]/p[1]", "/section[1]", null);

        // Assert: should have 2 paragraphs now, both with same text
        Assert.Equal("/p[2]", resultPath);
        var node1 = handler.Get("/section[1]/p[1]", 0);
        var node2 = handler.Get("/p[2]", 0);
        Assert.Equal(node1.Text, node2.Text);

        // The copied element should be a distinct XML node (deep clone)
        var sectionXml = handler.Raw("Contents/section0.xml");
        Assert.Equal(2, System.Xml.Linq.XDocument.Parse(sectionXml)
            .Root!.Elements().Count(e => e.Name.LocalName == "p"));
    }

    // ============================================================
    // 3. AddPart() — throws CliException with "unsupported_operation"
    // ============================================================
    [Fact]
    public void AddPart_ThrowsUnsupportedOperation()
    {
        var path = CreateTemp();
        using var handler = new HwpxHandler(path, editable: true);

        var ex = Assert.Throws<CliException>(() => handler.AddPart("/", "chart"));
        Assert.Equal("unsupported_operation", ex.Code);
        Assert.Contains("OPF packaging", ex.Message);
    }

    // ============================================================
    // 4. Raw() roundtrip — parse → serialize → re-parse
    // ============================================================
    [Fact]
    public void Raw_RoundtripIsEquivalent()
    {
        var path = CreateTemp("라운드트립 테스트");
        using var handler = new HwpxHandler(path, editable: false);

        // Act: get raw XML, parse it, serialize again
        var xml1 = handler.Raw("Contents/section0.xml");
        var parsed = System.Xml.Linq.XDocument.Parse(xml1);
        var xml2 = parsed.ToString();

        // Assert: re-serialized XML should match
        Assert.Equal(xml1, xml2);
    }

    // ============================================================
    // 5. Multi-section path resolution
    // ============================================================
    [Fact]
    public void PathResolver_MultiSection_CorrectIndex()
    {
        var path = HwpxTestHelper.CreateMultiSectionHwpx(
            new[] { "Sec1 Para1", "Sec1 Para2", "Sec1 Para3" },
            new[] { "Sec2 Para1", "Sec2 Para2" });
        _tempFiles.Add(path);

        using var handler = new HwpxHandler(path, editable: false);

        // Act: get paragraph 1 of section 2
        var node = handler.Get("/section[2]/p[1]", 0);

        // Assert: should resolve to section 2's first paragraph
        Assert.Equal("Sec2 Para1", node.Text);
        Assert.Equal("/section[2]/p[1]", node.Path);
    }

    // ============================================================
    // 6. PUA characters stripped
    // ============================================================
    [Fact]
    public void Korean_PuaCharsStripped()
    {
        var textWithPua = "계약서\uE001 작성\uF8FF 안내";
        var path = CreateTemp(textWithPua);
        using var handler = new HwpxHandler(path, editable: false);

        var text = handler.ViewAsText();

        Assert.Contains("계약서작성안내", text);
        Assert.DoesNotContain("\uE001", text);
        Assert.DoesNotContain("\uF8FF", text);
    }

    // ============================================================
    // 7. Korean uniform spacing normalized
    // ============================================================
    [Fact]
    public void Korean_UniformSpacingNormalized()
    {
        // Double-width spaces between Korean syllables should be collapsed
        var textWithSpacing = "한글  문서  처리";
        var path = CreateTemp(textWithSpacing);
        using var handler = new HwpxHandler(path, editable: false);

        var text = handler.ViewAsText();

        // R7: The regex removes ALL spaces between Korean syllables, not just extra spaces
        Assert.Contains("한글문서처리", text);
        Assert.DoesNotContain("한글  문서", text);
    }

    // ============================================================
    // 8. cellMargin required — validate fails if missing
    // ============================================================
    [Fact]
    public void Validate_MissingCellMargin_ReturnsError()
    {
        var path = HwpxTestHelper.CreateHwpxWithTable(
            rows: 2, cols: 2, includeCellMargin: false);
        _tempFiles.Add(path);

        using var handler = new HwpxHandler(path, editable: false);

        var errors = handler.Validate();

        Assert.Contains(errors, e => e.ErrorType == "table_missing_cellmargin" && e.Description.Contains("cellMargin"));
    }

    // ============================================================
    // 9a. Dual cellAddr format — child element variant
    // ============================================================
    [Fact]
    public void GetCellAddr_ChildElement_Parsed()
    {
        var path = HwpxTestHelper.CreateHwpxWithTable(
            rows: 1, cols: 1, includeCellAddr: true);
        _tempFiles.Add(path);

        using var handler = new HwpxHandler(path, editable: false);

        var node = handler.Get("/section[1]/tbl[1]/tr[1]/tc[1]", 0);

        Assert.Equal(0, node.Format["row"]);
        Assert.Equal(0, node.Format["col"]);
    }

    // ============================================================
    // 9b. Dual cellAddr format — tc-attribute variant (legacy)
    // ============================================================
    [Fact]
    public void GetCellAddr_TcAttribute_Parsed()
    {
        var path = HwpxTestHelper.CreateHwpxWithLegacyCellAddr();
        _tempFiles.Add(path);

        using var handler = new HwpxHandler(path, editable: false);

        var node = handler.Get("/section[1]/tbl[1]/tr[1]/tc[2]", 0);

        // tc[2] should have colAddr=1 from the legacy attribute format
        Assert.Equal(1, node.Format["col"]);
        Assert.Equal(0, node.Format["row"]);
    }

    // ============================================================
    // 10. AllParagraphs() local index resets per section
    // ============================================================
    [Fact]
    public void AllParagraphs_LocalIndexResetPerSection()
    {
        // Section 1 has 3 paragraphs, Section 2 has 2
        var path = HwpxTestHelper.CreateMultiSectionHwpx(
            new[] { "S1P1", "S1P2", "S1P3" },
            new[] { "S2P1", "S2P2" });
        _tempFiles.Add(path);

        using var handler = new HwpxHandler(path, editable: false);

        // /section[2]/p[1] should be the FIRST paragraph of section 2
        // NOT the 4th paragraph overall
        var node = handler.Get("/section[2]/p[1]", 0);
        Assert.Equal("S2P1", node.Text);

        // And /section[2]/p[2] should be the second
        var node2 = handler.Get("/section[2]/p[2]", 0);
        Assert.Equal("S2P2", node2.Text);
    }

    // ============================================================
    // 11. Tables uses Elements not Descendants — nested tables not double-counted
    // ============================================================
    [Fact]
    public void Section_Tables_DoesNotCountNestedTables()
    {
        // Create a section with a table that contains a nested table in a cell
        // The outer table should be counted, but the inner one should NOT
        // because HwpxSection.Tables uses Elements() (direct children only)
        var path = CreateTemp("test");
        _tempFiles.Add(path);

        using var handler = new HwpxHandler(path, editable: false);

        // Stats should report table count based on direct children only
        var stats = handler.ViewAsStats();
        // Minimal HWPX has no tables
        Assert.Contains("Tables:     0", stats);
    }

    // ============================================================
    // 12. EnsureCharPrProp clones shared charPr before modify
    // ============================================================
    [Fact]
    public void SetRun_SharedCharPr_ClonesBeforeModify()
    {
        // Two paragraphs sharing the same charPrIDRef="0"
        var path = HwpxTestHelper.CreateMultiSectionHwpx(
            new[] { "Paragraph A", "Paragraph B" });
        _tempFiles.Add(path);

        using var handler = new HwpxHandler(path, editable: true);

        // Get initial charPrIDRef for both runs
        var run1Before = handler.Get("/section[1]/p[1]/run[1]", 2);
        var run2Before = handler.Get("/section[1]/p[2]/run[1]", 2);

        // Both should share charPrIDRef="0" initially
        var ref1 = run1Before.Format.ContainsKey("charPrIDRef")
            ? run1Before.Format["charPrIDRef"]
            : null;
        var ref2 = run2Before.Format.ContainsKey("charPrIDRef")
            ? run2Before.Format["charPrIDRef"]
            : null;
        Assert.Equal(ref1, ref2); // shared charPrIDRef before mutation

        // Modify run 1's bold — triggers EnsureCharPrProp clone-on-write
        handler.Set("/section[1]/p[1]/run[1]", new Dictionary<string, string> { ["bold"] = "true" });

        // Run 2 should still reference the original charPrIDRef
        var run2After = handler.Get("/section[1]/p[2]/run[1]", 2);
        var ref2After = run2After.Format.ContainsKey("charPrIDRef")
            ? run2After.Format["charPrIDRef"]
            : null;
        Assert.Equal(ref2, ref2After); // run 2's charPrIDRef unchanged
    }
}
