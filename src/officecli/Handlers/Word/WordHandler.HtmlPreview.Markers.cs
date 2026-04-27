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

            var markerProps = BuildMarkerCssProperties(rpr);
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
    /// </summary>
    private static string BuildMarkerCssProperties(NumberingSymbolRunProperties? rpr)
    {
        if (rpr == null) return "";
        var parts = new List<string>();
        var clr = rpr.GetFirstChild<Color>();
        if (clr?.Val?.Value != null && !string.IsNullOrEmpty(clr.Val.Value) && clr.Val.Value != "auto")
            parts.Add($"color:#{clr.Val.Value}");
        var rf = rpr.GetFirstChild<RunFonts>();
        var fontName = rf?.Ascii?.Value ?? rf?.HighAnsi?.Value ?? rf?.EastAsia?.Value;
        if (!string.IsNullOrEmpty(fontName))
            parts.Add($"font-family:'{fontName}'");
        var fs = rpr.GetFirstChild<FontSize>();
        if (fs?.Val?.Value != null && int.TryParse(fs.Val.Value, out var halfPt))
            parts.Add($"font-size:{halfPt / 2.0:0.##}pt");
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
