// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;
using Vml = DocumentFormat.OpenXml.Vml;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeCli.Handlers;

public partial class WordHandler
{

    // CONSISTENCY(style-dual-key): resolve a style display name to its
    // OOXML styleId by scanning the styles part. Returns null when no
    // matching style is found, letting callers fall back to using the
    // value verbatim (lenient input). Used by paragraph-level Set on
    // styleName so users can write back the canonical readback key.
    private string? ResolveStyleIdFromName(string displayName)
    {
        var stylesPart = _doc.MainDocumentPart?.StyleDefinitionsPart;
        if (stylesPart?.Styles == null || string.IsNullOrEmpty(displayName)) return null;
        var match = stylesPart.Styles.Elements<Style>()
            .FirstOrDefault(s => string.Equals(s.StyleName?.Val?.Value, displayName, StringComparison.Ordinal));
        return match?.StyleId?.Value;
    }

    /// <summary>
    /// Issue #366: `style=`/`styleId=` target the OOXML styleId. Passing a DISPLAY
    /// NAME ("heading 2" where the styleId is "Heading2" — exactly what a Word or
    /// Google Docs export looks like) stores a dangling pStyle that Word renders as
    /// body text and drops on save. The generic "not found" advisory left the caller
    /// with nowhere to go, so when the value IS an existing style's display name,
    /// name the channel that resolves it and the id it resolves to.
    /// </summary>
    private string StyleNotFoundWarning(string value)
    {
        var resolved = ResolveStyleIdFromName(value);
        return resolved != null
            ? $"style '{value}' is a display name, not a styleId — will be referenced as-is. "
                + $"Use styleName='{value}' to resolve it (styleId '{resolved}')."
            : $"style '{value}' not found in styles part — will be referenced as-is";
    }

    /// <summary>
    /// Word's built-in paragraph styles that a fresh document does not define
    /// (its styles.xml carries only Normal). Word keys a built-in by its
    /// DISPLAY NAME: a <w:pStyle w:val="Heading1"/> with no such style in the
    /// part is body text to Word — no outline level, no Navigation pane entry,
    /// nothing for a screen reader to jump to — while `view outline` happily
    /// reported a heading. Referencing one of these ids now materializes the
    /// definition Word itself would write (name, outline level, the default
    /// look), so the file means what the caller asked for.
    /// Values follow Word's Normal template: colour accent1 (2F5496) / its
    /// darker shade (1F3763) / near-black (272727); sizes in half-points.
    /// </summary>
    private static readonly Dictionary<string, (string Name, int? OutlineLvl, int Sz, string? Color, bool Italic, int Before)>
        BuiltInParagraphStyles = new(StringComparer.Ordinal)
    {
        ["Heading1"] = ("heading 1", 0, 32, "2F5496", false, 240),
        ["Heading2"] = ("heading 2", 1, 26, "2F5496", false, 40),
        ["Heading3"] = ("heading 3", 2, 24, "1F3763", false, 40),
        ["Heading4"] = ("heading 4", 3, 22, "2F5496", true, 40),
        ["Heading5"] = ("heading 5", 4, 22, "2F5496", false, 40),
        ["Heading6"] = ("heading 6", 5, 22, "1F3763", false, 40),
        ["Heading7"] = ("heading 7", 6, 22, "1F3763", true, 40),
        ["Heading8"] = ("heading 8", 7, 21, "272727", false, 40),
        ["Heading9"] = ("heading 9", 8, 21, "272727", true, 40),
        ["Title"] = ("Title", null, 56, null, false, 0),
        ["Subtitle"] = ("Subtitle", null, 22, "5A5A5A", false, 0),
    };

    /// <summary>Word's display name for a built-in styleId, or null.</summary>
    internal static string? BuiltInStyleName(string? styleId)
        => styleId != null && BuiltInParagraphStyles.TryGetValue(styleId, out var d) ? d.Name : null;

    /// <summary>
    /// Define <paramref name="styleId"/> in the styles part when it is one of
    /// Word's built-ins and the part does not carry it yet. Returns true when a
    /// definition was added.
    /// </summary>
    internal bool TryMaterializeBuiltInStyle(string? styleId)
    {
        if (styleId == null || StyleIdExists(styleId)
            || !BuiltInParagraphStyles.TryGetValue(styleId, out var d))
            return false;
        var mainPart = _doc.MainDocumentPart;
        if (mainPart == null) return false;
        var stylesPart = mainPart.StyleDefinitionsPart ?? mainPart.AddNewPart<StyleDefinitionsPart>();
        stylesPart.Styles ??= new Styles();

        var style = new Style { Type = StyleValues.Paragraph, StyleId = styleId };
        style.AppendChild(new StyleName { Val = d.Name });
        style.AppendChild(new BasedOn { Val = "Normal" });
        style.AppendChild(new NextParagraphStyle { Val = "Normal" });
        style.AppendChild(new UIPriority { Val = d.OutlineLvl.HasValue ? 9 : 10 });
        if (d.OutlineLvl is > 0) style.AppendChild(new SemiHidden());
        if (d.OutlineLvl is > 0) style.AppendChild(new UnhideWhenUsed());
        style.AppendChild(new PrimaryStyle());

        var pPr = new StyleParagraphProperties();
        if (d.OutlineLvl.HasValue)
        {
            pPr.AppendChild(new KeepNext());
            pPr.AppendChild(new KeepLines());
            pPr.AppendChild(new SpacingBetweenLines { Before = d.Before.ToString(), After = "0" });
            pPr.AppendChild(new OutlineLevel { Val = d.OutlineLvl.Value });
        }
        else
        {
            pPr.AppendChild(new SpacingBetweenLines { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto });
            pPr.AppendChild(new ContextualSpacing());
        }
        style.AppendChild(pPr);

        var rPr = new StyleRunProperties();
        if (d.OutlineLvl is null or < 3 || styleId == "Title")
            rPr.AppendChild(new RunFonts { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi });
        if (d.Italic) { rPr.AppendChild(new Italic()); rPr.AppendChild(new ItalicComplexScript()); }
        if (d.Color != null) rPr.AppendChild(new Color { Val = d.Color });
        if (styleId == "Title") { rPr.AppendChild(new Spacing { Val = -10 }); rPr.AppendChild(new Kern { Val = 28U }); }
        rPr.AppendChild(new FontSize { Val = d.Sz.ToString() });
        rPr.AppendChild(new FontSizeComplexScript { Val = d.Sz.ToString() });
        style.AppendChild(rPr);

        stylesPart.Styles.AppendChild(style);
        return true;
    }

    /// <summary>
    /// Returns true if a style with the given styleId exists in the Styles part.
    /// "Normal" is implicit in OOXML and considered to exist even when the
    /// blank-document StyleDefinitionsPart is empty/absent — matches Word's
    /// own behaviour where every doc has Normal as the default paragraph style.
    /// </summary>
    internal bool StyleIdExists(string? styleId)
    {
        if (string.IsNullOrEmpty(styleId)) return false;
        if (string.Equals(styleId, "Normal", StringComparison.Ordinal)) return true;
        var stylesPart = _doc.MainDocumentPart?.StyleDefinitionsPart;
        if (stylesPart?.Styles == null) return false;
        return stylesPart.Styles.Elements<Style>()
            .Any(s => string.Equals(s.StyleId?.Value, styleId, StringComparison.Ordinal));
    }

    private string GetStyleName(Paragraph para)
    {
        var styleId = para.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        if (styleId == null) return "Normal";

        // Try to resolve display name from styles part
        var stylesPart = _doc.MainDocumentPart?.StyleDefinitionsPart;
        if (stylesPart?.Styles != null)
        {
            var style = stylesPart.Styles.Elements<Style>()
                .FirstOrDefault(s => s.StyleId?.Value == styleId);
            if (style?.StyleName?.Val?.Value != null)
                return style.StyleName.Val.Value;
        }

        return styleId;
    }

    private static int GetHeadingLevel(string styleName)
    {
        // Heading 1, Heading 2, heading1, 标题 1, etc.
        foreach (var ch in styleName)
        {
            if (char.IsDigit(ch))
                return ch - '0';
        }
        if (styleName == "Title") return 0;
        if (styleName == "Subtitle") return 1;
        return 1;
    }

    // CONSISTENCY(outline-resolution): build a styleId -> outline level map
    // from each style's own w:outlineLvl. OOXML §17.3.1.20: w:val is
    // ST_DecimalNumber 0-8; surface it as a 1-based level (val + 1). Heading
    // styles carry this, so a paragraph can resolve to an outline level even
    // when its styleId is numeric or localized and its display name contains
    // no recognizable "Heading"/"标题" token.
    private Dictionary<string, int> BuildStyleOutlineLevels()
    {
        var map = new Dictionary<string, int>(StringComparer.Ordinal);
        var styles = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles;
        if (styles == null) return map;
        foreach (var s in styles.Elements<Style>())
        {
            var id = s.StyleId?.Value;
            if (string.IsNullOrEmpty(id)) continue;
            var lvl = s.StyleParagraphProperties?.OutlineLevel?.Val?.Value;
            if (lvl.HasValue && lvl.Value >= 0 && lvl.Value <= 8)
                map[id] = lvl.Value + 1;
        }
        return map;
    }

    // Resolve a paragraph's outline level (0 = Title, 1-9 = outline depth),
    // or -1 when the paragraph is not part of the document outline. Signals,
    // in priority order, mirror TOC generation so `view outline` and TOC agree:
    //   1. direct paragraph w:outlineLvl (OOXML §17.3.1.20, val 0-8 -> level 1-9)
    //   2. the paragraph style's own w:outlineLvl (via BuildStyleOutlineLevels)
    //   3. style display name (Heading N / 标题 N / Title / Subtitle)
    private int GetParagraphOutlineLevel(Paragraph para, Dictionary<string, int> styleLevels, out string styleName)
    {
        styleName = GetStyleName(para);

        var direct = para.ParagraphProperties?.OutlineLevel?.Val?.Value;
        if (direct.HasValue && direct.Value >= 0 && direct.Value <= 8)
            return direct.Value + 1;

        var styleId = para.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        if (styleId != null && styleLevels.TryGetValue(styleId, out var sl))
            return sl;

        if (styleName.Contains("Heading") || styleName.Contains("标题")
            || styleName.StartsWith("heading", StringComparison.OrdinalIgnoreCase))
            return GetHeadingLevel(styleName);
        if (styleName == "Title") return 0;
        if (styleName == "Subtitle") return 1;
        return -1;
    }

    private static bool IsNormalStyle(string styleName)
    {
        return styleName is "Normal" or "正文" or "Body Text" or "Body" or "a"
            || styleName.StartsWith("Normal");
    }

    /// <summary>
    /// Apply the <c>typography.preset=zh-body</c> compound preset to a
    /// paragraph: set an East Asian body face (SimSun/宋体) on every run and on
    /// the paragraph mark so CJK text stops falling back to an unknown default,
    /// and set the paragraph rhythm to a CJK-friendly single-spaced body
    /// (line rule atLeast, single-line, no space after). Run-level — lives here
    /// because it is easiest to read next to the East-asian helpers, and it is
    /// invoked from <see cref="WordHandler.Set.Element"/> where the paragraph is
    /// in hand.
    /// </summary>
    private static void ApplyTypography_Body_Zh(
        Paragraph para, ParagraphProperties pProps, List<string>? warnings)
    {
        const string eaFont = "SimSun";
        // Runs keep their existing explicit western face; only the East Asia slot
        // is pinned to a Chinese body face so latin tokens stay on the doc's font
        // while CJK glyphs resolve deterministically.
        foreach (var run in para.Elements<Run>())
        {
            var rProps = run.RunProperties ?? run.PrependChild(new RunProperties());
            var rf = rProps.GetFirstChild<RunFonts>();
            if (rf == null)
            {
                rf = new RunFonts();
                // CT_RPr schema order: rFonts must come first (before b/i/sz…).
                rProps.PrependChild(rf);
            }
            rf.EastAsia = eaFont;
        }
        // Paragraph-mark rPr needs the eastAsia face too, or the line rhythm
        // cursor can still fall back.
        var markRPr = pProps.ParagraphMarkRunProperties;
        if (markRPr != null)
        {
            var mrF = markRPr.GetFirstChild<RunFonts>();
            if (mrF == null)
            {
                mrF = new RunFonts();
                markRPr.PrependChild(mrF);
            }
            mrF.EastAsia = eaFont;
        }

        // Rhythm: CJK body conventionally lineRule=atLeast single (so a line is
        // never compressed below one full ascent of the letters), spaceAfter 0.
        var spacing = pProps.SpacingBetweenLines ?? (pProps.SpacingBetweenLines = new SpacingBetweenLines());
        spacing.Line = "240";                                          // single
        spacing.LineRule = LineSpacingRuleValues.AtLeast;              // never compress
        spacing.After = "0";
        warnings?.Add($"zh-body preset: eastAsia font set to '{eaFont}' on all runs of this paragraph");
    }
}
