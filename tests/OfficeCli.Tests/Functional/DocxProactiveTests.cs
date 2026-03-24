// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Proactive tests for DOCX bugs found via pattern scanning.
/// </summary>
public class DocxProactiveTests : IDisposable
{
    private readonly List<string> _tempFiles = new();

    private (string path, WordHandler handler) CreateDoc()
    {
        var path = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.docx");
        _tempFiles.Add(path);
        BlankDocCreator.Create(path);
        return (path, new WordHandler(path, editable: true));
    }

    public void Dispose()
    {
        foreach (var f in _tempFiles)
            try { File.Delete(f); } catch { }
    }

    // ────────────────────────────────────────────────────────────────────────
    // Pattern 2: highlight is Set-able on paragraph but not read back
    //
    // Set(paragraphPath, { ["highlight"] = "yellow" }) applies highlight to
    // all runs, but Get(paragraphPath).Format does not include "highlight"
    // from the first run. This is inconsistent with bold/italic/color etc.
    // ────────────────────────────────────────────────────────────────────────
    [Fact]
    public void ParagraphGet_HighlightFromFirstRun_IsReadBack()
    {
        var (_, handler) = CreateDoc();
        using var h = handler;

        h.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Highlighted text",
            ["highlight"] = "yellow"
        });

        var node = h.Get("/body/p[1]");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("highlight",
            "paragraph Get should read back highlight from first run");
        node.Format["highlight"].Should().Be("yellow");
    }

    // ────────────────────────────────────────────────────────────────────────
    // Pattern 2: Set highlight on paragraph, then Get readback
    // Verifies Set + Get roundtrip at paragraph level.
    // ────────────────────────────────────────────────────────────────────────
    [Fact]
    public void ParagraphSet_Highlight_IsReadBackOnGet()
    {
        var (_, handler) = CreateDoc();
        using var h = handler;

        h.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Normal text"
        });

        h.Set("/body/p[1]", new Dictionary<string, string> { ["highlight"] = "green" });

        var node = h.Get("/body/p[1]");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("highlight");
        node.Format["highlight"].Should().Be("green");
    }

    // ────────────────────────────────────────────────────────────────────────
    // Pattern 1 (Query): paragraph[bold=true] should match bold paragraphs
    // This is the same as Bug8 but from the proactive scan perspective,
    // verifying italic, font, size, color also work in paragraph queries.
    // ────────────────────────────────────────────────────────────────────────
    [Fact]
    public void QueryParagraph_ByItalic_ReturnsMatches()
    {
        var (_, handler) = CreateDoc();
        using var h = handler;

        h.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Italic paragraph",
            ["italic"] = "true"
        });
        h.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Normal paragraph"
        });

        var results = h.Query("paragraph[italic=true]");
        results.Should().HaveCount(1);
        results[0].Text.Should().Be("Italic paragraph");
    }

    [Fact]
    public void QueryParagraph_ByFont_ReturnsMatches()
    {
        var (_, handler) = CreateDoc();
        using var h = handler;

        h.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Courier paragraph",
            ["font"] = "Courier New"
        });
        h.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Default paragraph"
        });

        var results = h.Query("paragraph[font=Courier New]");
        results.Should().HaveCount(1);
        results[0].Text.Should().Be("Courier paragraph");
    }

    [Fact]
    public void QueryParagraph_BySize_ReturnsMatches()
    {
        var (_, handler) = CreateDoc();
        using var h = handler;

        h.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Big paragraph",
            ["size"] = "24pt"
        });
        h.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Default paragraph"
        });

        var results = h.Query("paragraph[size=24pt]");
        results.Should().HaveCount(1);
        results[0].Text.Should().Be("Big paragraph");
    }

    [Fact]
    public void QueryParagraph_ByColor_ReturnsMatches()
    {
        var (_, handler) = CreateDoc();
        using var h = handler;

        h.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Red paragraph",
            ["color"] = "#FF0000"
        });
        h.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Default paragraph"
        });

        var results = h.Query("paragraph[color=#FF0000]");
        results.Should().HaveCount(1);
        results[0].Text.Should().Be("Red paragraph");
    }

    [Fact]
    public void QueryParagraph_ByHighlight_ReturnsMatches()
    {
        var (_, handler) = CreateDoc();
        using var h = handler;

        h.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Highlighted paragraph",
            ["highlight"] = "yellow"
        });
        h.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Default paragraph"
        });

        var results = h.Query("paragraph[highlight=yellow]");
        results.Should().HaveCount(1);
        results[0].Text.Should().Be("Highlighted paragraph");
    }

    [Fact]
    public void QueryParagraph_BoldFalse_ReturnsNonBoldParagraphs()
    {
        var (_, handler) = CreateDoc();
        using var h = handler;

        h.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Bold paragraph",
            ["bold"] = "true"
        });
        h.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Normal paragraph"
        });

        var results = h.Query("paragraph[bold=false]");
        results.Should().HaveCount(1);
        results[0].Text.Should().Be("Normal paragraph");
    }

    // ────────────────────────────────────────────────────────────────────────
    // Persistence test: highlight on paragraph survives reopen
    // ────────────────────────────────────────────────────────────────────────
    [Fact]
    public void ParagraphHighlight_PersistsAfterReopen()
    {
        var (path, handler) = CreateDoc();
        using var h = handler;

        h.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "Highlighted text",
            ["highlight"] = "cyan"
        });

        // Reopen
        h.Dispose();
        using var h2 = new WordHandler(path, editable: false);

        var node = h2.Get("/body/p[1]");
        node.Should().NotBeNull();
        node!.Format.Should().ContainKey("highlight");
        node.Format["highlight"].Should().Be("cyan");
    }
}
