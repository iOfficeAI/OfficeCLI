// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0
//
// CJK Helper unit tests — validates character detection, font chains,
// language tags, kinsoku rules, and WordML/DrawingML integration.

using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Tests;

public class CjkHelperTests
{
    // ── ContainsCjk ─────────────────────────────────────────────────

    [Theory]
    [InlineData("한글 테스트", true)]
    [InlineData("日本語テスト", true)]
    [InlineData("中文测试", true)]
    [InlineData("Hello World", false)]
    [InlineData("", false)]
    [InlineData(null, false)]
    [InlineData("Mixed 한글 text", true)]
    [InlineData("1234567890", false)]
    public void ContainsCjk_DetectsCorrectly(string? text, bool expected)
    {
        Assert.Equal(expected, CjkHelper.ContainsCjk(text));
    }

    // ── DetectScript ────────────────────────────────────────────────

    [Theory]
    [InlineData("한글 테스트", CjkScript.Korean)]
    [InlineData("日本語テスト", CjkScript.Japanese)]
    [InlineData("中文测试", CjkScript.Chinese)]
    [InlineData("Hello World", CjkScript.None)]
    [InlineData("", CjkScript.None)]
    [InlineData(null, CjkScript.None)]
    [InlineData("한글과 漢字 mixed", CjkScript.Korean)]
    public void DetectScript_IdentifiesCorrectScript(string? text, CjkScript expected)
    {
        Assert.Equal(expected, CjkHelper.DetectScript(text));
    }

    [Fact]
    public void SegmentText_SplitsMixedLatinAndCjk()
    {
        var segments = CjkHelper.SegmentText("Hello 한글 mixed 日本語テスト test");

        Assert.Collection(segments,
            segment =>
            {
                Assert.Equal("Hello ", segment.text);
                Assert.Equal(CjkScript.None, segment.script);
            },
            segment =>
            {
                Assert.Equal("한글", segment.text);
                Assert.Equal(CjkScript.Korean, segment.script);
            },
            segment =>
            {
                Assert.Equal(" mixed ", segment.text);
                Assert.Equal(CjkScript.None, segment.script);
            },
            segment =>
            {
                Assert.Equal("日本語テスト", segment.text);
                Assert.Equal(CjkScript.Japanese, segment.script);
            },
            segment =>
            {
                Assert.Equal(" test", segment.text);
                Assert.Equal(CjkScript.None, segment.script);
            });
    }

    // ── GetFontChain ────────────────────────────────────────────────

    [Fact]
    public void GetFontChain_Korean_ReturnsMalgunGothic()
    {
        var (primary, _, _) = CjkHelper.GetFontChain(CjkScript.Korean);
        Assert.Equal("Malgun Gothic", primary);
    }

    [Fact]
    public void GetFontChain_Japanese_ReturnsYuGothic()
    {
        var (primary, _, _) = CjkHelper.GetFontChain(CjkScript.Japanese);
        Assert.Equal("Yu Gothic", primary);
    }

    [Fact]
    public void GetFontChain_Chinese_ReturnsYaHei()
    {
        var (primary, _, _) = CjkHelper.GetFontChain(CjkScript.Chinese);
        Assert.Equal("Microsoft YaHei", primary);
    }

    [Fact]
    public void GetFontChain_None_ReturnsEmpty()
    {
        var (primary, _, _) = CjkHelper.GetFontChain(CjkScript.None);
        Assert.Equal("", primary);
    }

    // ── GetLanguageTag ──────────────────────────────────────────────

    [Theory]
    [InlineData(CjkScript.Korean, "ko-KR")]
    [InlineData(CjkScript.Japanese, "ja-JP")]
    [InlineData(CjkScript.Chinese, "zh-CN")]
    [InlineData(CjkScript.None, "")]
    public void GetLanguageTag_ReturnsCorrectTag(CjkScript script, string expected)
    {
        Assert.Equal(expected, CjkHelper.GetLanguageTag(script));
    }

    // ── Kinsoku ─────────────────────────────────────────────────────

    [Theory]
    [InlineData('。', true)]
    [InlineData('、', true)]
    [InlineData('）', true)]
    [InlineData('A', false)]
    public void IsKinsokuStart_DetectsCorrectly(char c, bool expected)
    {
        Assert.Equal(expected, CjkHelper.IsKinsokuStart(c));
    }

    [Theory]
    [InlineData('「', true)]
    [InlineData('（', true)]
    [InlineData('A', false)]
    public void IsKinsokuEnd_DetectsCorrectly(char c, bool expected)
    {
        Assert.Equal(expected, CjkHelper.IsKinsokuEnd(c));
    }

    // ── WordML integration ──────────────────────────────────────────

    [Fact]
    public void ApplyToWordRun_Korean_SetsEastAsiaFont()
    {
        var rPr = new RunProperties();
        CjkHelper.ApplyToWordRun(rPr, CjkScript.Korean);

        var rFonts = rPr.GetFirstChild<RunFonts>();
        Assert.NotNull(rFonts);
        Assert.Equal("Malgun Gothic", rFonts!.EastAsia?.Value);
    }

    [Fact]
    public void ApplyToWordRun_Korean_SetsLanguageTag()
    {
        var rPr = new RunProperties();
        CjkHelper.ApplyToWordRun(rPr, CjkScript.Korean);

        var lang = rPr.GetFirstChild<Languages>();
        Assert.NotNull(lang);
        Assert.Equal("ko-KR", lang!.EastAsia?.Value);
    }

    [Fact]
    public void ApplyToWordRunIfCjk_CreatesRunProperties()
    {
        var run = new Run(new Text("한글 테스트"));
        CjkHelper.ApplyToWordRunIfCjk(run, "한글 테스트");

        var rPr = run.GetFirstChild<RunProperties>();
        Assert.NotNull(rPr);
        Assert.Equal("Malgun Gothic", rPr!.GetFirstChild<RunFonts>()?.EastAsia?.Value);
    }

    [Fact]
    public void ApplyToWordRunIfCjk_NoOp_ForLatinText()
    {
        var run = new Run(new Text("Hello"));
        CjkHelper.ApplyToWordRunIfCjk(run, "Hello");

        Assert.Null(run.GetFirstChild<RunProperties>());
    }

    [Fact]
    public void ApplyToWordRunIfCjk_ClearsExistingCjkMetadata_ForLatinText()
    {
        var run = new Run(new RunProperties(), new Text("한글"));
        CjkHelper.ApplyToWordRunIfCjk(run, "한글");
        CjkHelper.ApplyToWordRunIfCjk(run, "Hello");

        var rPr = run.GetFirstChild<RunProperties>();
        Assert.NotNull(rPr);
        Assert.Null(rPr!.GetFirstChild<RunFonts>()?.EastAsia?.Value);
        Assert.Null(rPr.GetFirstChild<Languages>()?.EastAsia?.Value);
    }

    // ── DrawingML integration ───────────────────────────────────────

    [Fact]
    public void ApplyToDrawingRun_Korean_SetsEastAsianFont()
    {
        var rPr = new A.RunProperties();
        CjkHelper.ApplyToDrawingRun(rPr, CjkScript.Korean);

        var eaFont = rPr.GetFirstChild<A.EastAsianFont>();
        Assert.NotNull(eaFont);
        Assert.Equal("Malgun Gothic", eaFont!.Typeface?.Value);
    }

    [Fact]
    public void ApplyToDrawingRun_Korean_SetsLanguage()
    {
        var rPr = new A.RunProperties();
        CjkHelper.ApplyToDrawingRun(rPr, CjkScript.Korean);

        Assert.Equal("ko-KR", rPr.Language?.Value);
    }

    [Fact]
    public void ApplyToDrawingRunIfCjk_FallsBackToEnUS()
    {
        var rPr = new A.RunProperties();
        CjkHelper.ApplyToDrawingRunIfCjk(rPr, "Hello World");

        Assert.Equal("en-US", rPr.Language?.Value);
        Assert.Null(rPr.GetFirstChild<A.EastAsianFont>());
    }

    [Fact]
    public void ApplyToDrawingRunIfCjk_Chinese_SetsYaHei()
    {
        var rPr = new A.RunProperties();
        CjkHelper.ApplyToDrawingRunIfCjk(rPr, "中文测试");

        Assert.Equal("zh-CN", rPr.Language?.Value);
        Assert.Equal("Microsoft YaHei", rPr.GetFirstChild<A.EastAsianFont>()?.Typeface?.Value);
    }

    [Fact]
    public void ApplyToDrawingRunIfCjk_ClearsExistingCjkMetadata_ForLatinText()
    {
        var rPr = new A.RunProperties();
        CjkHelper.ApplyToDrawingRunIfCjk(rPr, "한글");
        CjkHelper.ApplyToDrawingRunIfCjk(rPr, "Hello World");

        Assert.Equal("en-US", rPr.Language?.Value);
        Assert.Null(rPr.GetFirstChild<A.EastAsianFont>());
    }
}
