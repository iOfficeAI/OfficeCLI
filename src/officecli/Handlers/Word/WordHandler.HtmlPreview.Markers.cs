// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    /// <summary>
    /// Walk every list-item paragraph in the body, collect the (numId, ilvl)
    /// pairs in use (resolving through pStyle for style-borne numbering), and
    /// emit a CSS block that styles each list marker per the abstractNum level's
    /// rPr (color, font, size, bold, italic) plus, for ul, the actual lvlText
    /// glyph as <c>list-style-type: '&lt;char&gt; '</c>.
    ///
    /// Class names used: <c>marker-{numId}-{ilvl}</c> on each &lt;li&gt;.
    /// Both ::marker (for ul) and the inline ol marker &lt;span&gt; pick up the
    /// styling — ol's path also reads the same fields inline at render time
    /// via <see cref="GetMarkerInlineCss"/>.
    /// </summary>
    private string BuildListMarkerCss(Body body)
    {
        var seen = new HashSet<(int numId, int ilvl)>();
        foreach (var para in body.Descendants<Paragraph>())
        {
            if (IsNumberingSuppressed(para)) continue;
            var resolved = ResolveNumPrFromStyle(para);
            if (resolved == null) continue;
            var (numId, ilvl) = resolved.Value;
            if (numId == 0) continue;
            if (ilvl < 0) ilvl = 0; else if (ilvl > 8) ilvl = 8;
            seen.Add((numId, ilvl));
        }
        if (seen.Count == 0) return "";

        var sb = new StringBuilder();
        foreach (var (numId, ilvl) in seen)
        {
            var lvl = GetLevel(numId, ilvl);
            if (lvl == null) continue;
            var rpr = lvl.NumberingSymbolRunProperties;
            var listStyleStr = GetCustomListStyleString(numId, ilvl);

            // When the marker is a CSS keyword (disc/circle/square) the browser
            // draws the glyph itself — font-family doesn't change the glyph but
            // its metrics still inflate the line box (Symbol's ascent > SimSun's
            // → ~0.75pt/line drift). Strip font-family from ::marker for keyword
            // markers; keep it for custom-string markers (★/▶/etc.) where the
            // font is what actually renders the glyph.
            var markerProps = BuildMarkerCssProperties(rpr, includeFontFamily: listStyleStr != null);
            // Skip when there is nothing to say — keeps the emitted CSS minimal.
            if (markerProps.Length == 0 && listStyleStr == null) continue;

            // ul: use ::marker and (when applicable) a custom list-style-type string.
            // CSS list-style-type accepts '<string> ' since CSS Counter Styles L3
            // (broad browser support), so we can render exact Word glyphs ★/▶/●
            // instead of falling back to disc/circle/square.
            if (listStyleStr != null)
            {
                sb.AppendLine($"li.marker-{numId}-{ilvl} {{ list-style-type: {listStyleStr}; }}");
            }
            if (markerProps.Length > 0)
            {
                sb.AppendLine($"li.marker-{numId}-{ilvl}::marker {{ {markerProps} }}");
            }
        }
        return sb.ToString();
    }

    /// <summary>
    /// Build a semicolon-separated CSS property string from a level's
    /// NumberingSymbolRunProperties (color, font, size, bold, italic).
    /// Empty string means no styled marker — caller skips emission.
    /// Used for both ::marker (ul) and the inline ol marker &lt;span&gt;.
    ///
    /// <paramref name="includeFontFamily"/> controls whether font-family is
    /// emitted. Pass false when the marker is a CSS keyword (disc/circle/
    /// square) — the keyword glyph is drawn by the browser regardless of font,
    /// but the font's metrics still inflate the ::marker line box. Pass true
    /// for custom-string markers and the ol inline span where the font does
    /// render the glyph.
    /// </summary>
    private static string BuildMarkerCssProperties(NumberingSymbolRunProperties? rpr, bool includeFontFamily = true)
    {
        if (rpr == null) return "";
        var parts = new List<string>();
        var clr = rpr.GetFirstChild<Color>();
        if (clr?.Val?.Value != null && !string.IsNullOrEmpty(clr.Val.Value) && clr.Val.Value != "auto")
            parts.Add($"color:#{clr.Val.Value}");
        var rf = rpr.GetFirstChild<RunFonts>();
        var fontName = rf?.Ascii?.Value ?? rf?.HighAnsi?.Value ?? rf?.EastAsia?.Value;
        if (includeFontFamily && !string.IsNullOrEmpty(fontName))
            parts.Add($"font-family:'{fontName}'");
        var fs = rpr.GetFirstChild<FontSize>();
        if (fs?.Val?.Value != null && int.TryParse(fs.Val.Value, out var halfPt))
        {
            parts.Add($"font-size:{halfPt / 2.0:0.##}pt");
            // Pin the marker's line-height to the font's natural ratio so the
            // marker doesn't inherit the parent body multiplier — keeps an
            // oversized marker from inflating the line box past its glyph
            // height.
            var ratio = OfficeCli.Core.FontMetricsReader.GetRatio(fontName ?? "Calibri");
            if (ratio > 0)
                parts.Add($"line-height:{ratio:0.####}");
        }
        if (rpr.GetFirstChild<Bold>() != null)
            parts.Add("font-weight:bold");
        if (rpr.GetFirstChild<Italic>() != null)
            parts.Add("font-style:italic");
        return string.Join(";", parts);
    }

    /// <summary>
    /// Public-to-class accessor for the inline marker CSS used by the ol
    /// marker &lt;span&gt; rendering path. Resolves the level by (numId, ilvl)
    /// and returns its rPr-derived CSS string, or empty if unstyled.
    /// </summary>
    private string GetMarkerInlineCss(int numId, int ilvl)
    {
        var lvl = GetLevel(numId, ilvl);
        return BuildMarkerCssProperties(lvl?.NumberingSymbolRunProperties);
    }

    /// <summary>
    /// Inline marker CSS that takes the host paragraph into account. Replaces
    /// the ratio-only line-height that <see cref="BuildMarkerCssProperties"/>
    /// emits with one driven by a per-paragraph layout formula:
    /// <code>
    ///   final = body_mlh × line_multiplier
    ///         + max(0, marker_ascent_pt − body_ascent_pt)
    /// </code>
    /// where ascent percentages come from <see cref="Core.FontMetricsReader.GetSplitAscDscOverride"/>
    /// and the multiplier is read from spacing.line (auto rule). For markers
    /// that are smaller than or equal to body content, the formula collapses
    /// to <c>body_mlh × multiplier</c>, matching plain-paragraph layout.
    /// Falls back to the ratio-based output when marker font-size is absent or
    /// font metrics aren't readable.
    /// </summary>
    private string GetMarkerInlineCss(int numId, int ilvl, Paragraph para)
    {
        var basic = GetMarkerInlineCss(numId, ilvl);
        if (string.IsNullOrEmpty(basic)) return basic;

        var lvl = GetLevel(numId, ilvl);
        var rpr = lvl?.NumberingSymbolRunProperties;

        var (bodySize, bodyFont, lineMulti) = ResolveBodyMetricsForMarker(para);
        var (bodyAscPct, bodyDscPct) = Core.FontMetricsReader.GetSplitAscDscOverride(bodyFont);
        if (bodyAscPct <= 0) return basic;
        var bodyAscPt = bodySize * bodyAscPct / 100.0;
        var bodyDscPt = bodySize * bodyDscPct / 100.0;

        var fs = rpr?.GetFirstChild<FontSize>();
        double markerSize = fs?.Val?.Value != null
                            && int.TryParse(fs.Val.Value, out var halfPt)
                            && halfPt > 0
            ? halfPt / 2.0
            : bodySize;

        var rf = rpr?.GetFirstChild<RunFonts>();
        var markerFont = rf?.Ascii?.Value ?? rf?.HighAnsi?.Value ?? rf?.EastAsia?.Value ?? "Calibri";
        var (markerAscPct, _) = Core.FontMetricsReader.GetSplitAscDscOverride(markerFont);
        if (markerAscPct <= 0) return basic;

        var lvlText = GetLevelText(numId, ilvl);
        if (!string.IsNullOrEmpty(lvlText)
            && lvlText.Any(c => c >= 0x2600)
            && !Core.FontMetricsReader.HasGlyphsForChars(markerFont, lvlText))
            markerAscPct = Math.Max(markerAscPct, 108.0);
        var markerAscPt = markerSize * markerAscPct / 100.0;

        var bodyExtraPt = (bodyAscPt + bodyDscPt) * (lineMulti - 1);
        var finalPt = Math.Max(bodyAscPt, markerAscPt) + bodyDscPt + bodyExtraPt;
        var lineHeight = finalPt / markerSize;

        var rx = new System.Text.RegularExpressions.Regex(@"line-height:[^;]+");
        var replacement = $"line-height:{lineHeight:0.####}";
        return rx.IsMatch(basic) ? rx.Replace(basic, replacement) : basic + ";" + replacement;
    }

    /// <summary>
    /// Absolute line height (pt) for a list item's &lt;li&gt; when the marker's
    /// ascent exceeds the body's. Returns null when the body lane already
    /// dominates (marker is smaller or absent). Returned as absolute pt rather
    /// than unitless multiplier so the &lt;li&gt; doesn't inherit a wrong body
    /// size — wild-bullet (TNR docDefaults, no run-level sz) showed the
    /// inherited 11pt default, not the actual 10pt body, would apply the
    /// multiplier and overshoot the intended height.
    /// </summary>
    private double? GetListItemLineHeightOverride(int numId, int ilvl, Paragraph para)
    {
        var lvl = GetLevel(numId, ilvl);
        var rpr = lvl?.NumberingSymbolRunProperties;

        var (bodySize, bodyFont, lineMulti) = ResolveBodyMetricsForMarker(para);
        var (bodyAscPct, bodyDscPct) = Core.FontMetricsReader.GetSplitAscDscOverride(bodyFont);
        if (bodyAscPct <= 0) return null;
        var bodyAscPt = bodySize * bodyAscPct / 100.0;
        var bodyDscPt = bodySize * bodyDscPct / 100.0;

        // Marker font-size: explicit <w:sz> in the lvl rPr if present,
        // otherwise inherit body size.
        var fs = rpr?.GetFirstChild<FontSize>();
        double markerSize = fs?.Val?.Value != null
                            && int.TryParse(fs.Val.Value, out var halfPt)
                            && halfPt > 0
            ? halfPt / 2.0
            : bodySize;

        var rf = rpr?.GetFirstChild<RunFonts>();
        var markerFont = rf?.Ascii?.Value ?? rf?.HighAnsi?.Value ?? rf?.EastAsia?.Value ?? "Calibri";
        var (markerAscPct, _) = Core.FontMetricsReader.GetSplitAscDscOverride(markerFont);
        if (markerAscPct <= 0) return null;

        // When the marker font's cmap doesn't cover lvlText, the renderer
        // falls back to a wider face whose effective ascent/em is ~108%.
        // Fallback-detection is gated on codepoints in the Misc Symbols /
        // Dingbats range (U+2600+) that Latin/symbol-encoded fonts
        // typically don't ship native glyphs for. Common bullets below
        // that range — • U+2022, ▪ U+25AA, ▫ U+25AB, ◦ U+25E6 — render
        // natively in most fonts (or via Symbol's PUA remap), so they
        // skip the bump.
        var lvlText = GetLevelText(numId, ilvl);
        if (!string.IsNullOrEmpty(lvlText)
            && lvlText.Any(c => c >= 0x2600)
            && !Core.FontMetricsReader.HasGlyphsForChars(markerFont, lvlText))
            markerAscPct = Math.Max(markerAscPct, 108.0);
        var markerAscPt = markerSize * markerAscPct / 100.0;

        if (markerAscPt <= bodyAscPt) return null;

        var bodyExtraPt = (bodyAscPt + bodyDscPt) * (lineMulti - 1);
        return markerAscPt + bodyDscPt + bodyExtraPt;
    }

    /// <summary>
    /// Resolve the body run's font/size and the paragraph's line multiplier
    /// for use in the marker line-height formula. Resolution order for size
    /// and font: explicit run rPr → docDefaults rPrDefault → OOXML implicit
    /// (10pt body, Calibri).
    /// </summary>
    private (double size, string font, double multi) ResolveBodyMetricsForMarker(Paragraph para)
    {
        double size = 0;
        string font = "";
        foreach (var r in para.Elements<Run>())
        {
            var rprBody = r.RunProperties;
            if (size == 0)
            {
                var sz = rprBody?.FontSize?.Val?.Value;
                if (sz != null && int.TryParse(sz, out var halfPt) && halfPt > 0)
                    size = halfPt / 2.0;
            }
            if (string.IsNullOrEmpty(font))
            {
                var f = rprBody?.RunFonts;
                font = f?.Ascii?.Value ?? f?.HighAnsi?.Value ?? f?.EastAsia?.Value ?? "";
            }
            if (size > 0 && !string.IsNullOrEmpty(font)) break;
        }
        if (size == 0 || string.IsNullOrEmpty(font))
        {
            var rPrDefault = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles?
                .DocDefaults?.RunPropertiesDefault?.RunPropertiesBaseStyle;
            if (size == 0)
            {
                var sz = rPrDefault?.FontSize?.Val?.Value;
                if (sz != null && int.TryParse(sz, out var halfPt) && halfPt > 0)
                    size = halfPt / 2.0;
            }
            if (string.IsNullOrEmpty(font))
            {
                var f = rPrDefault?.RunFonts;
                font = f?.Ascii?.Value ?? f?.HighAnsi?.Value ?? f?.EastAsia?.Value ?? "";
            }
        }
        if (size == 0) size = 10.0;
        if (string.IsNullOrEmpty(font)) font = "Calibri";

        double multi = 1.0;
        var pPr = para.ParagraphProperties;
        var spacing = pPr?.SpacingBetweenLines
                      ?? ResolveSpacingFromStyle(pPr?.ParagraphStyleId?.Val?.Value);
        if (spacing?.Line?.Value is string lv && int.TryParse(lv, out var twips))
        {
            var rule = spacing.LineRule?.InnerText;
            if (rule == "auto" || rule == null)
                multi = twips / 240.0;
        }
        return (size, font, multi);
    }

    /// <summary>
    /// Look up the abstractNumId that a num instance points at. Returns null
    /// if the num isn't found. Used to key the cross-num running counter so
    /// "continue" sibling lists (no startOverride) share a counter with the
    /// list that ran before them on the same abstractNum.
    /// </summary>
    private int? GetAbstractNumId(int numId)
    {
        var numbering = _doc.MainDocumentPart?.NumberingDefinitionsPart?.Numbering;
        var inst = numbering?.Elements<NumberingInstance>()
            .FirstOrDefault(n => n.NumberID?.Value == numId);
        return inst?.AbstractNumId?.Val?.Value;
    }

    /// <summary>
    /// Read the startOverride value (if any) for one level of a num instance.
    /// Returns null when the num lacks a &lt;w:lvlOverride w:ilvl=N&gt; with a
    /// &lt;w:startOverride/&gt; child for the requested level — i.e. "continue
    /// counting" semantics applies.
    /// </summary>
    private int? GetNumStartOverride(int numId, int ilvl)
    {
        var numbering = _doc.MainDocumentPart?.NumberingDefinitionsPart?.Numbering;
        var inst = numbering?.Elements<NumberingInstance>()
            .FirstOrDefault(n => n.NumberID?.Value == numId);
        if (inst == null) return null;
        var ovr = inst.Elements<LevelOverride>()
            .FirstOrDefault(o => o.LevelIndex?.Value == ilvl);
        return ovr?.StartOverrideNumberingValue?.Val?.Value;
    }

    /// <summary>
    /// For ul lists, when the lvlText is a single non-standard glyph (★/▶/etc.)
    /// the existing disc/circle/square mapping silently downgrades to •.
    /// Return a CSS string literal like <c>'★ '</c> that <c>list-style-type</c>
    /// accepts directly, so the rendered bullet matches the Word source.
    /// Returns null if the standard CSS mapping is sufficient.
    /// </summary>
    private string? GetCustomListStyleString(int numId, int ilvl)
    {
        var fmt = GetNumberingFormat(numId, ilvl);
        if (!fmt.Equals("bullet", StringComparison.OrdinalIgnoreCase)) return null;
        var text = GetLevelText(numId, ilvl);
        if (string.IsNullOrEmpty(text)) return null;
        // Already covered by the existing disc/circle/square switch in the
        // main render path — don't override those.
        if (text == "•" || text == "o" || text == "▪"
            || text == "◦" /* ◦ */ || text == "▪" /* ▪ */
            || text == "" /* Wingdings square */)
            return null;
        // Escape ' and \ for CSS string literal.
        var escaped = text!.Replace("\\", "\\\\").Replace("'", "\\'");
        return $"'{escaped} '";
    }
}
