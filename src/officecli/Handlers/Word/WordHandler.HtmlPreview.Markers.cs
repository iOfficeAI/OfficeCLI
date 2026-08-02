// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
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
        // Bold/Italic are OnOff toggles: ON when the element is present with
        // no val or a true val, OFF when val is "0"/false. The marker level rPr
        // commonly carries <w:i w:val="0"/> to turn OFF an italic inherited from
        // the surrounding context — treating the element's mere presence as ON
        // wrongly italicized an upright numbered marker. Mirror the run path.
        var markerBold = rpr.GetFirstChild<Bold>();
        if (markerBold != null && (markerBold.Val == null || markerBold.Val.Value))
            parts.Add("font-weight:bold");
        var markerItalic = rpr.GetFirstChild<Italic>();
        if (markerItalic != null && (markerItalic.Val == null || markerItalic.Val.Value))
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
    /// the ratio-only line-height with the per-paragraph layout formula:
    /// <code>
    ///   final = max(body_asc, marker_asc) + body_dsc
    ///         + max(body_mlh, marker_mlh) × (line_multiplier − 1)
    /// </code>
    /// per OOXML §17.3.1.33 auto-rule: extra leading (multiplier > 1) is
    /// driven by the taller of body / marker mlh and added below the baseline.
    /// Ascent / descent ratios come from
    /// <see cref="Core.FontMetricsReader.GetSplitAscDscOverride"/>; the
    /// multiplier is read from spacing.line. For markers smaller than or equal
    /// to body content the formula collapses to <c>body_mlh × multiplier</c>.
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
        var (markerAscPct, markerDscPct) = Core.FontMetricsReader.GetSplitAscDscOverride(markerFont);
        if (markerAscPct <= 0) return basic;

        var lvlText = GetLevelText(numId, ilvl);
        if (!string.IsNullOrEmpty(lvlText)
            && lvlText.Any(c => c >= 0x2600)
            && !Core.FontMetricsReader.HasGlyphsForChars(markerFont, lvlText))
            markerAscPct = Math.Max(markerAscPct, 108.0);
        var markerAscPt = markerSize * markerAscPct / 100.0;
        var markerDscPt = markerSize * markerDscPct / 100.0;

        var bodyMlhPt = bodyAscPt + bodyDscPt;
        var markerMlhPt = markerAscPt + markerDscPt;
        var bodyExtraPt = Math.Max(bodyMlhPt, markerMlhPt) * (lineMulti - 1);
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
        var (markerAscPct, markerDscPct) = Core.FontMetricsReader.GetSplitAscDscOverride(markerFont);
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
        var markerDscPt = markerSize * markerDscPct / 100.0;

        if (markerAscPt <= bodyAscPt) return null;

        // Auto-rule (OOXML §17.3.1.33) extra leading is driven by the taller
        // of body / marker mlh, added below the baseline.
        var bodyMlhPt = bodyAscPt + bodyDscPt;
        var markerMlhPt = markerAscPt + markerDscPt;
        var bodyExtraPt = Math.Max(bodyMlhPt, markerMlhPt) * (lineMulti - 1);
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
        // ECMA-376 §17.9.7: a lvlOverride that embeds a full <w:lvl> replaces
        // the entire level definition (including its own <w:start>), and the
        // startOverride is ignored. Defer to the embedded level's start (read
        // via GetStartValue) by reporting "no override" here.
        if (ovr?.GetFirstChild<Level>() != null) return null;
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
        // A LOW (non-PUA) code point under a SYMBOL font (Wingdings / Symbol /
        // Webdings) must render through the custom list-style-type string +
        // ::marker font-family, so the browser draws the symbol font's glyph at
        // that slot. The disc/circle/square keyword switch maps low ASCII by its
        // LATIN meaning (e.g. "o" → circle), which is wrong for a symbol font
        // where the 'o' slot is a checkbox ☐. PUA bullets that HAVE a keyword
        // entry (U+F0B7 • → disc, U+F0A7 ▪ → square) keep that mapping: routing
        // them to disc/square is intentional (cleaner marker metrics). PUA
        // checkbox/checkmark slots with no keyword (U+F0A8/F0FE/F0FD/F0FC/F0FB)
        // fall through to TranslateSymbolPuaGlyph below and become portable
        // Unicode ☐/☒/☑/✓/✗. So only skip the keyword early-return for a
        // low-code symbol-font bullet.
        var symbolLowCode = text![0] < 0xF000
                            && IsSymbolBulletFont(GetBulletFontName(numId, ilvl));
        if (!symbolLowCode)
        {
            // Already covered by the standard disc/circle/square switch in the
            // main render path — don't override those.
            if (BulletGlyphToCssKeyword(text!) != null) return null;
        }
        // Translate Symbol/Wingdings private-use code points that map to a
        // real Unicode glyph but have no CSS list-style-type keyword (e.g.
        // Symbol 0x2D minus → en-dash bullet). Otherwise the raw PUA char
        // lands in the CSS string literal and renders as tofu (□).
        text = TranslateSymbolPuaGlyph(text!);
        // Escape ' and \ for CSS string literal.
        var escaped = text!.Replace("\\", "\\\\").Replace("'", "\\'");
        return $"'{escaped} '";
    }

    /// <summary>
    /// Resolve a numbering level's bullet font name from its
    /// NumberingSymbolRunProperties rFonts (ascii → hAnsi → eastAsia), or null
    /// when the level has no symbol run properties / no font.
    /// </summary>
    private string? GetBulletFontName(int numId, int ilvl)
    {
        var rf = GetLevel(numId, ilvl)?.NumberingSymbolRunProperties?.GetFirstChild<RunFonts>();
        return rf?.Ascii?.Value ?? rf?.HighAnsi?.Value ?? rf?.EastAsia?.Value;
    }

    // CONSISTENCY(bullet-glyph-map): a "symbol font" is one whose bullet glyph
    // is selected purely by code point in the font's private encoding (the
    // letter "o" in Wingdings is a checkbox ☐, not the Latin o). When the
    // level's bullet font is one of these, the lvlText code point must be
    // rendered with that font on the marker — at ANY code point, high PUA
    // (U+F0xx) or low ASCII (U+006F) — never downgraded to a disc/circle/square
    // keyword. Single source of truth for the symbol-font test, shared by
    // GetCustomListStyleString (CSS ::marker path) and GetUlListStyleTypeCss
    // (inline body/table list-style-type).
    private static bool IsSymbolBulletFont(string? fontName)
    {
        if (string.IsNullOrEmpty(fontName)) return false;
        return fontName!.StartsWith("Wingdings", StringComparison.OrdinalIgnoreCase)
            || fontName.Equals("Symbol", StringComparison.OrdinalIgnoreCase)
            || fontName.Equals("Webdings", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// CONSISTENCY(bullet-glyph-map): compute the inline <c>list-style-type</c>
    /// value for a <c>ul</c> bullet, shared by the body and table-cell render
    /// paths. For a symbol-font bullet (Wingdings/Symbol/Webdings) returns the
    /// custom string literal (e.g. <c>'o '</c>) so it matches the
    /// <c>li.marker-N-K</c> CSS class's list-style-type and the browser draws
    /// the symbol glyph via the ::marker font-family — an inline keyword
    /// (disc/circle) would otherwise win over the class and drop the symbol.
    /// For non-symbol fonts returns the disc/circle/square keyword (default
    /// disc), preserving existing behaviour for ordinary bullets.
    /// </summary>
    private string GetUlListStyleTypeCss(int numId, int ilvl, string? lvlText)
    {
        // Mirror GetCustomListStyleString: a low-code symbol-font bullet renders
        // via the custom glyph string so the inline list-style-type matches the
        // li.marker-N-K class (whose ::marker font-family draws the symbol). PUA
        // and non-symbol bullets keep the disc/circle/square keyword.
        if (!string.IsNullOrEmpty(lvlText) && lvlText![0] < 0xF000
            && IsSymbolBulletFont(GetBulletFontName(numId, ilvl)))
        {
            var custom = GetCustomListStyleString(numId, ilvl);
            if (custom != null) return custom;
        }
        return BulletGlyphToCssKeyword(lvlText ?? "") ?? "disc";
    }

    // CONSISTENCY(bullet-glyph-map): single source of truth mapping a Word
    // bullet lvlText glyph to a CSS list-style-type keyword. Called from all
    // three ul render paths so they never diverge:
    //   - GetCustomListStyleString (above): non-null => glyph is "standard",
    //     so skip the custom 'X ' list-style-type string override.
    //   - WordHandler.HtmlPreview.cs (body ul switch)
    //   - WordHandler.HtmlPreview.Tables.cs (table-cell ul switch)
    // Returns null when the glyph is not a recognized standard bullet; the
    // body/table callers then default to "disc". Covers Word's default
    // Wingdings round bullet U+F0B7 (the most common default) -> disc.
    private static string? BulletGlyphToCssKeyword(string lvlText) => lvlText switch
    {
        "•" => "disc",        // • BULLET
        "" => "disc",        // Wingdings round bullet (Word default)
        "o" => "circle",
        "◦" => "circle",      // ◦ WHITE BULLET (Word outline level 1)
        "" => "square",      // Wingdings square
        "▪" => "square",      // ▪ BLACK SMALL SQUARE
        _ => null
    };

    // CONSISTENCY(bullet-glyph-map): Symbol/Wingdings private-use bullet code
    // points that resolve to a real Unicode glyph with NO CSS list-style-type
    // keyword (so they can't go through BulletGlyphToCssKeyword). Word renders
    // the font's glyph at that slot; the HTML preview must substitute the
    // matching Unicode char or the raw PUA code point lands in the CSS string
    // literal / plain text and renders as tofu (□) on any host without the
    // symbol font installed (Wingdings/Symbol are not web-safe).
    //
    // Mappings follow the published Wingdings/Symbol -> Unicode glyph
    // tables. The PUA slots (U+F0xx) are the F000-shifted form of the font's
    // 8-bit code (e.g. Wingdings 0xFE -> U+F0FE); the low-code slot is the bare
    // 8-bit code under a symbol font (Wingdings 'o' = U+006F = empty box).
    //   - U+F02D = Symbol font slot 0x2D (minus/hyphen) -> en-dash (U+2013).
    //   - Wingdings checkbox/checkmark slots used by "to-do" list markers, so a
    //     host with no Wingdings font still shows a real box/check, not tofu:
    //       0xA8 / U+F0A8 -> U+2610 empty ballot box
    //       'o'  / U+006F -> U+2610 empty ballot box (Wingdings 'o' slot)
    //       0xFE / U+F0FE -> U+2611 ballot box with check
    //       0xFD / U+F0FD -> U+2612 ballot box with X
    //       0xFC / U+F0FC -> U+2713 check mark
    //       0xFB / U+F0FB -> U+2717 ballot X
    // (Wingdings 0xA7 / U+F0A7 square and 0xB7 / U+F0B7 round bullet already
    //  resolve to the square/disc CSS keyword via BulletGlyphToCssKeyword, so
    //  they never reach here and need no entry.)
    // Single source of truth, applied by GetCustomListStyleString (HTML ::marker
    // string literal) and BulletGlyphForText (plain-text walker). Real Unicode
    // equivalents are preferred (most portable); a slot with no equivalent keeps
    // the raw char and the caller adds font-family:'Wingdings' as fallback.
    private static string TranslateSymbolPuaGlyph(string lvlText) => lvlText switch
    {
        "\uf02d" => "\u2013", // Symbol 0x2D minus -> en-dash bullet
        "\uf0a8" => "\u2610", // Wingdings 0xA8 -> empty ballot box
        "o" => "\u2610",      // Wingdings 'o' (low-code) -> empty ballot box
        "\uf0fe" => "\u2611", // Wingdings 0xFE -> ballot box with check
        "\uf0fd" => "\u2612", // Wingdings 0xFD -> ballot box with X
        "\uf0fc" => "\u2713", // Wingdings 0xFC -> check mark
        "\uf0fb" => "\u2717", // Wingdings 0xFB -> ballot X
        _ => lvlText
    };

    // CONSISTENCY(bullet-glyph-map): plain-text counterpart of
    // BulletGlyphToCssKeyword. `view text` must show the SAME bullet Word
    // renders, which for a custom lvlText glyph (★ ▶ ● …) is the glyph itself —
    // the old code collapsed every bullet to "•", so custom glyphs vanished and
    // text disagreed with both the HTML preview (which passes them through via
    // list-style-type) and Word. Map the recognized standard bullets to their
    // visible glyph, pass real Unicode glyphs through verbatim, and fall back to
    // "•" for an empty lvlText or an unmapped private-use (Wingdings/Symbol)
    // code point that would render as tofu in plain text.
    private static string BulletGlyphForText(string? lvlText)
    {
        switch (BulletGlyphToCssKeyword(lvlText ?? ""))
        {
            case "disc": return "•";
            case "circle": return "◦";
            case "square": return "▪";
        }
        if (string.IsNullOrEmpty(lvlText)) return "•";
        // Symbol/Wingdings PUA slots with a known real-glyph equivalent
        // (e.g. F02D → en-dash) translate before the generic-disc fallback so
        // plain text matches Word and the HTML ::marker string.
        var translated = TranslateSymbolPuaGlyph(lvlText!);
        if (!ReferenceEquals(translated, lvlText)) return translated;
        var c = lvlText![0];
        if (c >= 0xF000 && c <= 0xF0FF) return "•"; // unmapped PUA -> generic disc
        return lvlText;
    }
}
