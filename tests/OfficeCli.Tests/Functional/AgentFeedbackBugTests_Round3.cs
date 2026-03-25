// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Tests that prove bugs found during Agent A Round 3 testing.
/// These tests are expected to FAIL until the bugs are fixed.
///
/// BUG-1: listLevel property is ignored during Add — all list items are created at level 0
///        even when listLevel=1 or listLevel=2 is specified. ApplyListStyle hardcodes Val=0.
///
/// BUG-3: SuggestProperty returns the exact same property name as the suggestion when the
///        input exactly matches a known property. For footnotes, Set only supports "text",
///        so "bold" is reported as unsupported with suggestion "did you mean: bold?" — self-referencing.
///
/// BUG-6: Get('/section[N]') with an out-of-bounds index returns a DocumentNode with Type="error"
///        but the CLI wraps it with success:true, which is misleading. The handler should throw
///        or the envelope should detect error-typed nodes.
///
/// BUG-7: Margin keys in Get('/section[N]') are returned as lowercase (margintop, marginbottom, etc.)
///        while Get('/') returns them in camelCase (marginTop, marginBottom, etc.). This violates
///        the canonical key naming convention.
/// </summary>
public class AgentFeedbackBugTests_Round3 : IDisposable
{
    private readonly string _path;
    private WordHandler _handler;

    public AgentFeedbackBugTests_Round3()
    {
        _path = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.docx");
        BlankDocCreator.Create(_path);
        _handler = new WordHandler(_path, editable: true);
    }

    public void Dispose()
    {
        _handler.Dispose();
        if (File.Exists(_path)) File.Delete(_path);
    }

    private WordHandler Reopen()
    {
        _handler.Dispose();
        _handler = new WordHandler(_path, editable: true);
        return _handler;
    }

    // ==================== BUG-1: listLevel ignored during Add ====================

    /// <summary>
    /// BUG-1: When creating a paragraph with listStyle=bullet and listLevel=1,
    /// the list level should be set to 1 (indented sub-item). Instead, ApplyListStyle
    /// always hardcodes NumberingLevelReference.Val = 0, ignoring the listLevel property.
    ///
    /// Note: Add uses "listLevel" or "numlevel" for the level, and Get returns "numlevel".
    /// </summary>
    [Fact]
    public void Add_Paragraph_WithListLevel_ShouldRespectLevel()
    {
        // Create a level-0 bullet item first
        _handler.Add("/", "paragraph", null, new()
        {
            ["text"] = "Level 0 item",
            ["listStyle"] = "bullet"
        });

        // Create a level-1 bullet item (sub-item) — use start=1 to force new numbering
        // (avoiding continuation which also ignores the level)
        var path2 = _handler.Add("/", "paragraph", null, new()
        {
            ["text"] = "Level 1 item",
            ["listStyle"] = "bullet",
            ["listLevel"] = "1",
            ["start"] = "1"
        });

        // Get the second paragraph and verify the list level is 1, not 0
        // Add returns /body/p[N] paths
        var node = _handler.Get(path2);
        node.Should().NotBeNull();

        // The numbering level (numlevel / ilvl) should be 1 for the sub-item.
        // Currently fails because ApplyListStyle always hardcodes Val = 0.
        node.Format.Should().ContainKey("numlevel");
        node.Format["numlevel"].ToString().Should().Be("1");
    }

    /// <summary>
    /// BUG-1 (variant): Even with listLevel=2, the paragraph is created at level 0.
    /// Uses numlevel (the canonical Get key) for verification.
    /// </summary>
    [Fact]
    public void Add_Paragraph_WithListLevel2_ShouldCreateLevel2()
    {
        var path = _handler.Add("/", "paragraph", null, new()
        {
            ["text"] = "Deep nested item",
            ["listStyle"] = "ordered",
            ["listLevel"] = "2"
        });

        // Reopen to verify persistence
        Reopen();

        var node = _handler.Get(path);
        node.Should().NotBeNull();
        // The item should be at level 2, producing a different numbering format (e.g. roman)
        // Currently fails: always level 0
        node.Format.Should().ContainKey("numlevel");
        node.Format["numlevel"].ToString().Should().Be("2");
    }

    // ==================== BUG-3: SuggestProperty self-reference ====================

    /// <summary>
    /// BUG-3: When setting a property on a footnote that only supports "text",
    /// other valid property names like "bold" are reported as unsupported.
    /// SuggestProperty("bold") finds "bold" in knownProps with distance 0 and
    /// suggests it, producing: "Unsupported property: bold (did you mean: bold?)"
    ///
    /// The suggestion should either be null (no suggestion when the input is already
    /// a known property) or suggest something actually different.
    /// </summary>
    [Fact]
    public void SuggestProperty_ExactMatch_ShouldNotSuggestSameProperty()
    {
        // SuggestProperty is internal static on CommandBuilder
        var suggestion = CommandBuilder.SuggestProperty("bold");

        // If there IS a suggestion, it must not be the same as the input.
        // Currently fails: suggestion == "bold" (distance 0 passes the threshold check).
        if (suggestion != null)
        {
            suggestion.Should().NotBe("bold",
                "suggesting the exact same property name is confusing — " +
                "'did you mean: bold?' when the user typed 'bold'");
        }
    }

    /// <summary>
    /// BUG-3 (variant): Same issue for other exact-match property names.
    /// "italic", "font", "color" etc. should not self-suggest.
    /// </summary>
    [Theory]
    [InlineData("italic")]
    [InlineData("font")]
    [InlineData("color")]
    [InlineData("underline")]
    [InlineData("size")]
    public void SuggestProperty_ExactMatchVariants_ShouldNotSuggestSameProperty(string propName)
    {
        var suggestion = CommandBuilder.SuggestProperty(propName);

        if (suggestion != null)
        {
            suggestion.Should().NotBe(propName,
                $"suggesting '{propName}' when user typed '{propName}' is a self-referencing error");
        }
    }

    /// <summary>
    /// BUG-3: End-to-end test — Set bold on a footnote should not produce a
    /// self-referencing "did you mean" message. The unsupported list should contain
    /// just "bold" without a misleading suggestion.
    /// </summary>
    [Fact]
    public void Set_FootnoteBold_ShouldNotProduceSelfReferencingSuggestion()
    {
        // Add a paragraph with a footnote
        _handler.Add("/", "paragraph", null, new() { ["text"] = "Text with footnote" });
        _handler.Add("/body/p[1]", "footnote", null, new() { ["text"] = "Footnote text" });

        // Try to set bold on the footnote — only "text" is supported
        var unsupported = _handler.Set("/footnote[1]", new() { ["bold"] = "true" });

        // "bold" should be in unsupported list
        unsupported.Should().NotBeEmpty();
        // The unsupported entry should NOT contain a self-referencing suggestion
        // like "bold (valid run props: ... bold ...)" that leads to "did you mean: bold?"
        unsupported[0].Should().NotContain("did you mean");
    }

    // ==================== BUG-6: Section out-of-bounds returns success with type=error ====================

    /// <summary>
    /// BUG-6: Get('/section[99]') on a document with 1 section returns a DocumentNode
    /// with Type="error". This error node is problematic because the CLI envelope
    /// wraps it with success:true. The handler should throw an exception for
    /// out-of-bounds access instead of silently returning an error-typed node.
    /// </summary>
    [Fact]
    public void Get_SectionOutOfBounds_ShouldThrowOrReturnClearError()
    {
        // A blank document has 1 section
        // Requesting section[99] should clearly indicate failure
        // The fix is to throw — so either the exception proves the bug is fixed,
        // or if it somehow returns a node, the Type must not be "error".
        DocumentNode? node = null;
        try
        {
            node = _handler.Get("/section[99]");
        }
        catch (ArgumentException)
        {
            // Throwing is the preferred fix — test passes
            return;
        }

        // If it didn't throw, the node type should NOT be "error" wrapped in a success response.
        node!.Type.Should().NotBe("error",
            "returning Type='error' inside a success:true envelope is confusing; " +
            "the handler should throw ArgumentException for out-of-bounds section index");
    }

    /// <summary>
    /// BUG-6 (alternative assertion): If the handler returns an error node,
    /// at minimum the node should have enough information for the caller to detect failure.
    /// The ideal fix is to throw, matching the pattern used by other handlers.
    /// </summary>
    [Fact]
    public void Get_SectionOutOfBounds_ShouldThrow()
    {
        // This is the preferred fix: throw an exception like other out-of-bounds cases
        var act = () => _handler.Get("/section[99]");
        act.Should().Throw<Exception>(
            "out-of-bounds section access should throw, not return a silent error node");
    }

    // ==================== BUG-7: Section margin keys returned as lowercase ====================

    /// <summary>
    /// BUG-7: Get('/section[1]') returns margin keys as "margintop", "marginbottom",
    /// "marginleft", "marginright" (all lowercase). But Get('/') returns them as
    /// "marginTop", "marginBottom", "marginLeft", "marginRight" (camelCase).
    /// Per canonical key rules, the keys should be consistent and use camelCase.
    /// </summary>
    [Fact]
    public void Get_Section_MarginKeys_ShouldBeCamelCase()
    {
        // Set margins via the document root so we know they exist
        _handler.Set("/", new()
        {
            ["marginTop"] = "1440",
            ["marginBottom"] = "1440",
            ["marginLeft"] = "1440",
            ["marginRight"] = "1440"
        });

        var sectionNode = _handler.Get("/section[1]");
        sectionNode.Should().NotBeNull();
        sectionNode.Type.Should().Be("section");

        // Verify keys use canonical camelCase, not lowercase
        // Currently fails: returns "margintop" instead of "marginTop"
        sectionNode.Format.Keys.Should().Contain("marginTop",
            "section margin keys should use camelCase 'marginTop', not 'margintop'");
        sectionNode.Format.Keys.Should().Contain("marginBottom",
            "section margin keys should use camelCase 'marginBottom', not 'marginbottom'");
        sectionNode.Format.Keys.Should().Contain("marginLeft",
            "section margin keys should use camelCase 'marginLeft', not 'marginleft'");
        sectionNode.Format.Keys.Should().Contain("marginRight",
            "section margin keys should use camelCase 'marginRight', not 'marginright'");

        // Should NOT contain lowercase versions
        sectionNode.Format.Keys.Should().NotContain("margintop");
        sectionNode.Format.Keys.Should().NotContain("marginbottom");
        sectionNode.Format.Keys.Should().NotContain("marginleft");
        sectionNode.Format.Keys.Should().NotContain("marginright");
    }

    /// <summary>
    /// BUG-7: Consistency check — root document Get and section Get should return
    /// the same key names for the same margin properties.
    /// </summary>
    [Fact]
    public void Get_Root_And_Section_MarginKeys_ShouldBeConsistent()
    {
        _handler.Set("/", new()
        {
            ["marginTop"] = "1440",
            ["marginBottom"] = "1440"
        });

        var rootNode = _handler.Get("/");
        var sectionNode = _handler.Get("/section[1]");

        // Root uses camelCase — section should match
        rootNode.Format.Keys.Should().Contain("marginTop");
        sectionNode.Format.Keys.Should().Contain("marginTop",
            "section Get should use the same camelCase keys as root Get");

        // The values should also be consistent for the same margins
        rootNode.Format["marginTop"].ToString()
            .Should().Be(sectionNode.Format.ContainsKey("marginTop")
                ? sectionNode.Format["marginTop"].ToString()
                : sectionNode.Format["margintop"].ToString(),
            "root and section should report the same margin values");
    }
}
