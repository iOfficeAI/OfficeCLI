// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    // ==================== Style Inheritance ====================

    private RunProperties ResolveEffectiveRunProperties(Run run, Paragraph para)
    {
        var effective = new RunProperties();

        // 1. Start with docDefaults rPr
        var docDefaults = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles?.DocDefaults;
        var defaultRPr = docDefaults?.RunPropertiesDefault?.RunPropertiesBaseStyle;
        if (defaultRPr != null)
            MergeRunProperties(effective, defaultRPr);

        // 2. Walk paragraph style basedOn chain (collect in order, apply from base to derived)
        var styleId = para.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        if (styleId != null)
        {
            var chain = new List<Style>();
            var visited = new HashSet<string>();
            var currentStyleId = styleId;
            while (currentStyleId != null && visited.Add(currentStyleId))
            {
                var style = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles
                    ?.Elements<Style>().FirstOrDefault(s => s.StyleId?.Value == currentStyleId);
                if (style == null) break;
                chain.Add(style);
                currentStyleId = style.BasedOn?.Val?.Value;
            }
            // Apply from base to derived (reverse order)
            for (int i = chain.Count - 1; i >= 0; i--)
            {
                var styleRPr = chain[i].StyleRunProperties;
                if (styleRPr != null)
                    MergeRunProperties(effective, styleRPr);
            }
        }

        // 3. Resolve character style (rStyle) from the run's rPr
        var rStyleId = run.RunProperties?.GetFirstChild<RunStyle>()?.Val?.Value;
        if (rStyleId != null)
        {
            var rStyleChain = new List<Style>();
            var rVisited = new HashSet<string>();
            var curRStyleId = rStyleId;
            while (curRStyleId != null && rVisited.Add(curRStyleId))
            {
                var rStyle = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles
                    ?.Elements<Style>().FirstOrDefault(s => s.StyleId?.Value == curRStyleId);
                if (rStyle == null) break;
                rStyleChain.Add(rStyle);
                curRStyleId = rStyle.BasedOn?.Val?.Value;
            }
            for (int i = rStyleChain.Count - 1; i >= 0; i--)
            {
                var sRPr = rStyleChain[i].StyleRunProperties;
                if (sRPr != null)
                    MergeRunProperties(effective, sRPr);
            }
        }

        // 4. Apply run's own direct rPr (highest priority, excluding rStyle which was resolved above)
        if (run.RunProperties != null)
            MergeRunProperties(effective, run.RunProperties);

        return effective;
    }

    private static void MergeRunProperties(RunProperties target, OpenXmlElement source)
    {
        // Merge each known property: source overwrites target
        var srcFonts = source.GetFirstChild<RunFonts>();
        if (srcFonts != null)
            target.RunFonts = srcFonts.CloneNode(true) as RunFonts;

        var srcSize = source.GetFirstChild<FontSize>();
        if (srcSize != null)
            target.FontSize = srcSize.CloneNode(true) as FontSize;

        var srcBold = source.GetFirstChild<Bold>();
        if (srcBold != null)
            target.Bold = srcBold.CloneNode(true) as Bold;

        var srcItalic = source.GetFirstChild<Italic>();
        if (srcItalic != null)
            target.Italic = srcItalic.CloneNode(true) as Italic;

        var srcUnderline = source.GetFirstChild<Underline>();
        if (srcUnderline != null)
            target.Underline = srcUnderline.CloneNode(true) as Underline;

        var srcStrike = source.GetFirstChild<Strike>();
        if (srcStrike != null)
            target.Strike = srcStrike.CloneNode(true) as Strike;

        var srcDStrike = source.GetFirstChild<DoubleStrike>();
        if (srcDStrike != null)
            target.DoubleStrike = srcDStrike.CloneNode(true) as DoubleStrike;

        var srcColor = source.GetFirstChild<Color>();
        if (srcColor != null)
            target.Color = srcColor.CloneNode(true) as Color;

        var srcHighlight = source.GetFirstChild<Highlight>();
        if (srcHighlight != null)
            target.Highlight = srcHighlight.CloneNode(true) as Highlight;

        var srcVertAlign = source.GetFirstChild<VerticalTextAlignment>();
        if (srcVertAlign != null)
            target.VerticalTextAlignment = srcVertAlign.CloneNode(true) as VerticalTextAlignment;

        var srcSmallCaps = source.GetFirstChild<SmallCaps>();
        if (srcSmallCaps != null)
            target.SmallCaps = srcSmallCaps.CloneNode(true) as SmallCaps;

        var srcCaps = source.GetFirstChild<Caps>();
        if (srcCaps != null)
            target.Caps = srcCaps.CloneNode(true) as Caps;

        var srcRtl = source.GetFirstChild<RightToLeftText>();
        if (srcRtl != null)
            target.RightToLeftText = srcRtl.CloneNode(true) as RightToLeftText;

        var srcShd = source.GetFirstChild<Shading>();
        if (srcShd != null)
            target.Shading = srcShd.CloneNode(true) as Shading;

        // Character spacing (w:spacing val in twips) — letter-spacing CSS equivalent
        var srcSpacing = source.GetFirstChild<Spacing>();
        if (srcSpacing != null)
            target.Spacing = srcSpacing.CloneNode(true) as Spacing;

        // Character scale (w:w horizontal stretch percentage)
        var srcCharScale = source.GetFirstChild<CharacterScale>();
        if (srcCharScale != null)
            target.CharacterScale = srcCharScale.CloneNode(true) as CharacterScale;

        // East Asian emphasis mark (w:em)
        var srcEm = source.GetFirstChild<Emphasis>();
        if (srcEm != null)
            target.Emphasis = srcEm.CloneNode(true) as Emphasis;

        // Rendering effects: outline, shadow, emboss, imprint
        var srcOutline = source.GetFirstChild<Outline>();
        if (srcOutline != null)
            target.Outline = srcOutline.CloneNode(true) as Outline;

        var srcShadow = source.GetFirstChild<Shadow>();
        if (srcShadow != null)
            target.Shadow = srcShadow.CloneNode(true) as Shadow;

        var srcEmboss = source.GetFirstChild<Emboss>();
        if (srcEmboss != null)
            target.Emboss = srcEmboss.CloneNode(true) as Emboss;

        var srcImprint = source.GetFirstChild<Imprint>();
        if (srcImprint != null)
            target.Imprint = srcImprint.CloneNode(true) as Imprint;

        var srcVanish = source.GetFirstChild<Vanish>();
        if (srcVanish != null)
            target.Vanish = srcVanish.CloneNode(true) as Vanish;

        var srcNoProof = source.GetFirstChild<NoProof>();
        if (srcNoProof != null)
            target.NoProof = srcNoProof.CloneNode(true) as NoProof;

        var srcBdr = source.GetFirstChild<Border>();
        if (srcBdr != null)
        {
            target.RemoveAllChildren<Border>();
            target.AppendChild(srcBdr.CloneNode(true));
        }

        // w14 text effects (textFill, textOutline, glow, shadow, reflection)
        foreach (var child in source.ChildElements)
        {
            if (child.NamespaceUri != "http://schemas.microsoft.com/office/word/2010/wordml") continue;
            // Remove existing w14 element with same local name, then add the new one
            var existing = target.ChildElements.FirstOrDefault(
                e => e.NamespaceUri == child.NamespaceUri && e.LocalName == child.LocalName);
            if (existing != null) target.RemoveChild(existing);
            target.AppendChild(child.CloneNode(true));
        }
    }

    private static string? GetFontFromProperties(RunProperties? rProps)
    {
        if (rProps == null) return null;
        var fonts = rProps.RunFonts;
        return fonts?.EastAsia?.Value ?? fonts?.Ascii?.Value ?? fonts?.HighAnsi?.Value;
    }

    private static string? GetSizeFromProperties(RunProperties? rProps)
    {
        if (rProps == null) return null;
        var size = rProps.FontSize?.Val?.Value;
        if (size == null) return null;
        return $"{int.Parse(size) / 2}pt";
    }

    // ==================== Effective Properties Resolution ====================

    /// <summary>
    /// Populates effective.* format keys on a paragraph node for properties not explicitly set.
    /// Resolves from: paragraph style chain → document defaults.
    /// </summary>
    private void PopulateEffectiveParagraphProperties(DocumentNode node, Paragraph para)
    {
        // Resolve effective run properties from the first run (or an empty run for style-only resolution)
        var firstRun = para.Elements<Run>().FirstOrDefault(r => r.GetFirstChild<Text>() != null)
            ?? new Run();
        var effective = ResolveEffectiveRunProperties(firstRun, para);

        // font.size
        if (!node.Format.ContainsKey("size") && effective.FontSize?.Val?.Value != null)
        {
            var sz = int.Parse(effective.FontSize.Val.Value) / 2.0;
            node.Format["effective.size"] = $"{sz:0.##}pt";
        }

        // font.name — CONSISTENCY(canonical-keys): check per-script slot keys, not legacy "font".
        if (!node.Format.ContainsKey("font.ascii") && !node.Format.ContainsKey("font.eastAsia")
            && !node.Format.ContainsKey("font.hAnsi") && !node.Format.ContainsKey("font.cs")
            && !node.Format.ContainsKey("font"))
        {
            var font = effective.RunFonts?.Ascii?.Value ?? effective.RunFonts?.HighAnsi?.Value
                ?? effective.RunFonts?.EastAsia?.Value;
            if (font != null)
                node.Format["effective.font"] = font;
        }

        // bold
        if (!node.Format.ContainsKey("bold") && effective.Bold != null)
            node.Format["effective.bold"] = true;

        // italic
        if (!node.Format.ContainsKey("italic") && effective.Italic != null)
            node.Format["effective.italic"] = true;

        // color
        if (!node.Format.ContainsKey("color"))
        {
            if (effective.Color?.Val?.Value != null)
                node.Format["effective.color"] = ParseHelpers.FormatHexColor(effective.Color.Val.Value);
            else if (effective.Color?.ThemeColor?.HasValue == true)
                node.Format["effective.color"] = effective.Color.ThemeColor.InnerText;
        }

        // underline
        if (!node.Format.ContainsKey("underline") && effective.Underline?.Val != null)
            node.Format["effective.underline"] = effective.Underline.Val.InnerText;

        // Resolve effective paragraph properties from style chain
        ResolveEffectiveParagraphStyleProperties(node, para);
    }

    /// <summary>
    /// Populates effective.* format keys on a run node for properties not explicitly set.
    /// </summary>
    private void PopulateEffectiveRunProperties(DocumentNode node, Run run, Paragraph para)
    {
        var effective = ResolveEffectiveRunProperties(run, para);

        if (!node.Format.ContainsKey("size") && effective.FontSize?.Val?.Value != null)
        {
            var sz = int.Parse(effective.FontSize.Val.Value) / 2.0;
            node.Format["effective.size"] = $"{sz:0.##}pt";
        }

        // CONSISTENCY(canonical-keys): check per-script slot keys, not legacy "font".
        if (!node.Format.ContainsKey("font.ascii") && !node.Format.ContainsKey("font.eastAsia")
            && !node.Format.ContainsKey("font.hAnsi") && !node.Format.ContainsKey("font.cs")
            && !node.Format.ContainsKey("font"))
        {
            var font = effective.RunFonts?.Ascii?.Value ?? effective.RunFonts?.HighAnsi?.Value
                ?? effective.RunFonts?.EastAsia?.Value;
            if (font != null)
                node.Format["effective.font"] = font;
        }

        if (!node.Format.ContainsKey("bold") && effective.Bold != null)
            node.Format["effective.bold"] = true;

        if (!node.Format.ContainsKey("italic") && effective.Italic != null)
            node.Format["effective.italic"] = true;

        if (!node.Format.ContainsKey("color"))
        {
            if (effective.Color?.Val?.Value != null)
                node.Format["effective.color"] = ParseHelpers.FormatHexColor(effective.Color.Val.Value);
            else if (effective.Color?.ThemeColor?.HasValue == true)
                node.Format["effective.color"] = effective.Color.ThemeColor.InnerText;
        }

        if (!node.Format.ContainsKey("underline") && effective.Underline?.Val != null)
            node.Format["effective.underline"] = effective.Underline.Val.InnerText;

        if (!node.Format.ContainsKey("strike") && effective.Strike != null)
            node.Format["effective.strike"] = true;

        if (!node.Format.ContainsKey("highlight") && effective.Highlight?.Val != null)
            node.Format["effective.highlight"] = effective.Highlight.Val.InnerText;
    }

    /// <summary>
    /// Resolves paragraph-level properties (alignment, spacing) from the paragraph style chain.
    /// </summary>
    private void ResolveEffectiveParagraphStyleProperties(DocumentNode node, Paragraph para)
    {
        var styleId = para.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        if (styleId == null) return;

        var chain = new List<Style>();
        var visited = new HashSet<string>();
        var currentStyleId = styleId;
        while (currentStyleId != null && visited.Add(currentStyleId))
        {
            var style = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles
                ?.Elements<Style>().FirstOrDefault(s => s.StyleId?.Value == currentStyleId);
            if (style == null) break;
            chain.Add(style);
            currentStyleId = style.BasedOn?.Val?.Value;
        }

        // Apply from base to derived (reverse order), collecting effective paragraph properties
        string? alignment = null;
        string? spaceBefore = null;
        string? spaceAfter = null;
        string? lineSpacing = null;

        for (int i = chain.Count - 1; i >= 0; i--)
        {
            var ppr = chain[i].StyleParagraphProperties;
            if (ppr == null) continue;

            if (ppr.Justification?.Val != null)
            {
                var txt = ppr.Justification.Val.InnerText;
                alignment = txt == "both" ? "justify" : txt;
            }
            if (ppr.SpacingBetweenLines?.Before?.Value != null)
                spaceBefore = SpacingConverter.FormatWordSpacing(ppr.SpacingBetweenLines.Before.Value);
            if (ppr.SpacingBetweenLines?.After?.Value != null)
                spaceAfter = SpacingConverter.FormatWordSpacing(ppr.SpacingBetweenLines.After.Value);
            if (ppr.SpacingBetweenLines?.Line?.Value != null)
                lineSpacing = SpacingConverter.FormatWordLineSpacing(
                    ppr.SpacingBetweenLines.Line.Value,
                    ppr.SpacingBetweenLines.LineRule?.InnerText);
        }

        if (!node.Format.ContainsKey("alignment") && alignment != null)
            node.Format["effective.alignment"] = alignment;
        if (!node.Format.ContainsKey("spaceBefore") && spaceBefore != null)
            node.Format["effective.spaceBefore"] = spaceBefore;
        if (!node.Format.ContainsKey("spaceAfter") && spaceAfter != null)
            node.Format["effective.spaceAfter"] = spaceAfter;
        if (!node.Format.ContainsKey("lineSpacing") && lineSpacing != null)
            node.Format["effective.lineSpacing"] = lineSpacing;
    }

    // ==================== List / Numbering ====================

    /// <summary>
    /// Resolve (numId, ilvl) from a paragraph by first checking its direct
    /// numPr and then walking up the linked paragraph style chain. Used by
    /// heading auto-numbering, which must honour style-defined numPr even
    /// when the paragraph itself has no NumberingProperties.
    /// </summary>
    /// <summary>
    /// True iff the paragraph explicitly suppresses numbering via a direct
    /// <c>&lt;w:numPr&gt;&lt;w:numId w:val="0"/&gt;&lt;/w:numPr&gt;</c>.
    /// This intentionally ignores the style chain — callers that want the
    /// effective numPr use <see cref="ResolveNumPrFromStyle"/> separately.
    /// </summary>
    private static bool IsNumberingSuppressed(Paragraph para)
    {
        var numProps = para.ParagraphProperties?.NumberingProperties;
        if (numProps == null) return false;
        var nid = numProps.NumberingId?.Val?.Value;
        return nid == 0;
    }

    private (int NumId, int Ilvl)? ResolveNumPrFromStyle(Paragraph para)
    {
        // 1. Direct numPr on the paragraph wins.
        var numProps = para.ParagraphProperties?.NumberingProperties;
        if (numProps != null)
        {
            var nid = numProps.NumberingId?.Val?.Value;
            if (nid != null && nid != 0)
                return (nid.Value, numProps.NumberingLevelReference?.Val?.Value ?? 0);
        }

        // 2. Walk the style chain through BasedOn references.
        var styleId = para.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        if (styleId == null) return null;

        var stylesPart = _doc.MainDocumentPart?.StyleDefinitionsPart;
        if (stylesPart?.Styles == null) return null;

        var visited = new HashSet<string>();
        while (styleId != null && visited.Add(styleId))
        {
            var style = stylesPart.Styles.Elements<Style>()
                .FirstOrDefault(s => s.StyleId?.Value == styleId);
            if (style == null) break;

            var styleNumPr = style.StyleParagraphProperties?.NumberingProperties;
            if (styleNumPr != null)
            {
                var nid = styleNumPr.NumberingId?.Val?.Value;
                if (nid != null && nid != 0)
                    return (nid.Value, styleNumPr.NumberingLevelReference?.Val?.Value ?? 0);
            }

            styleId = style.BasedOn?.Val?.Value;
        }

        return null;
    }

    private string? GetParagraphListStyle(Paragraph para)
    {
        if (IsNumberingSuppressed(para)) return null;

        // Direct numPr always wins — paragraph is a list item.
        var directNumPr = para.ParagraphProperties?.NumberingProperties;
        var directNid = directNumPr?.NumberingId?.Val?.Value;
        if (directNid != null && directNid != 0)
        {
            var ilvl = directNumPr!.NumberingLevelReference?.Val?.Value ?? 0;
            var numFmt = GetNumberingFormat(directNid.Value, ilvl);
            return numFmt.ToLowerInvariant() == "bullet" ? "bullet" : "ordered";
        }

        // Style-inherited numPr: skip when the paragraph is itself a heading
        // (Heading1..9 / Title / Subtitle). Headings with style-borne numPr
        // render via the heading path with a heading-num span (existing
        // behavior); treating them as <li> would double-count and break the
        // expected <h1>/<h2> output.
        var styleName = GetStyleName(para);
        if (!string.IsNullOrEmpty(styleName))
        {
            if (styleName.Contains("Heading") || styleName.Contains("标题")
                || styleName.StartsWith("heading", StringComparison.OrdinalIgnoreCase)
                || styleName == "Title" || styleName == "Subtitle")
                return null;
        }
        var resolved = ResolveNumPrFromStyle(para);
        if (resolved == null) return null;
        var (numId, ilvlR) = resolved.Value;
        if (numId == 0) return null;
        var numFmtR = GetNumberingFormat(numId, ilvlR);
        return numFmtR.ToLowerInvariant() == "bullet" ? "bullet" : "ordered";
    }

    private string GetListPrefix(Paragraph para)
    {
        var numProps = para.ParagraphProperties?.NumberingProperties;
        if (numProps == null) return "";

        var numId = numProps.NumberingId?.Val?.Value;
        var ilvl = numProps.NumberingLevelReference?.Val?.Value ?? 0;
        if (numId == null || numId == 0) return "";

        var indent = new string(' ', ilvl * 2);
        var numFmt = GetNumberingFormat(numId.Value, ilvl);

        return numFmt.ToLowerInvariant() switch
        {
            "bullet" => $"{indent}• ",
            "decimal" => $"{indent}1. ",
            "lowerletter" => $"{indent}a. ",
            "upperletter" => $"{indent}A. ",
            "lowerroman" => $"{indent}i. ",
            "upperroman" => $"{indent}I. ",
            _ => $"{indent}• "
        };
    }

    private string GetNumberingFormat(int numId, int ilvl)
    {
        var level = GetLevel(numId, ilvl);
        var numFmt = level?.NumberingFormat?.Val;
        if (numFmt == null || !numFmt.HasValue) return "bullet";
        return numFmt.InnerText ?? "bullet";
    }

    /// <summary>Get picture bullet data URI for a numbering level (if lvlPicBulletId is set).</summary>
    private string? GetPicBulletDataUri(int numId, int ilvl)
    {
        var numPart = _doc.MainDocumentPart?.NumberingDefinitionsPart;
        var numbering = numPart?.Numbering;
        if (numbering == null) return null;

        var numInstance = numbering.Elements<NumberingInstance>()
            .FirstOrDefault(n => n.NumberID?.Value == numId);
        var abstractNumId = numInstance?.AbstractNumId?.Val?.Value;
        if (abstractNumId == null) return null;
        var abstractNum = numbering.Elements<AbstractNum>()
            .FirstOrDefault(a => a.AbstractNumberId?.Value == abstractNumId);
        var level = abstractNum?.Elements<Level>()
            .FirstOrDefault(l => l.LevelIndex?.Value == ilvl);

        // Check for lvlPicBulletId
        var picBulletIdAttr = level?.GetAttributes().FirstOrDefault(a => a.LocalName == "lvlPicBulletId");
        if (picBulletIdAttr is not { } attr || attr.Value == null) return null;

        // Find the matching numPicBullet element
        var picBulletEl = level?.Descendants().FirstOrDefault(e => e.LocalName == "lvlPicBulletId");
        if (picBulletEl == null) return null;
        var picBulletIdStr = picBulletEl.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;
        if (picBulletIdStr == null || !int.TryParse(picBulletIdStr, out var picBulletId)) return null;

        // Find numPicBullet with this ID in numbering.xml
        var numPicBullet = numbering.Descendants().FirstOrDefault(e =>
            e.LocalName == "numPicBullet" &&
            e.GetAttributes().Any(a => a.LocalName == "numPicBulletId" && a.Value == picBulletIdStr));
        if (numPicBullet == null) return null;

        // Extract image from VML imagedata r:id reference
        var imageData = numPicBullet.Descendants().FirstOrDefault(e => e.LocalName == "imagedata");
        var rId = imageData?.GetAttributes().FirstOrDefault(a => a.LocalName == "id").Value;
        if (rId == null) return null;

        try
        {
            var imgPart = numPart!.GetPartById(rId);
            if (imgPart == null) return null;
            using var stream = imgPart.GetStream();
            using var ms = new System.IO.MemoryStream();
            stream.CopyTo(ms);
            var bytes = ms.ToArray();
            var mime = imgPart.ContentType ?? "image/png";
            return $"data:{mime};base64,{Convert.ToBase64String(bytes)}";
        }
        catch { return null; }
    }

    private string? GetLevelText(int numId, int ilvl)
        => GetLevel(numId, ilvl)?.LevelText?.Val?.Value;

    /// <summary>Get the LevelSuffix (tab/space/nothing) for a numbering level. Defaults to "tab".</summary>
    private string GetLevelSuffix(int numId, int ilvl)
    {
        var level = GetLevel(numId, ilvl);
        var suff = level?.LevelSuffix?.Val;
        if (suff?.HasValue != true) return "tab";
        return suff.InnerText ?? "tab";
    }

    /// <summary>Get the LevelJustification (left/center/right) for a numbering level. Defaults to "left".</summary>
    private string GetLevelJustification(int numId, int ilvl)
    {
        var level = GetLevel(numId, ilvl);
        var jc = level?.LevelJustification?.Val;
        if (jc?.HasValue != true) return "left";
        return jc.InnerText ?? "left";
    }

    private Level? GetLevel(int numId, int ilvl)
    {
        var numbering = _doc.MainDocumentPart?.NumberingDefinitionsPart?.Numbering;
        if (numbering == null) return null;
        var numInstance = numbering.Elements<NumberingInstance>()
            .FirstOrDefault(n => n.NumberID?.Value == numId);
        if (numInstance == null) return null;

        // A `<w:lvlOverride>` on the NumberingInstance can embed an entire
        // `<w:lvl>` replacing the abstractNum's level definition (not just
        // the startOverride number). Honor that before falling back.
        var lvlOverride = numInstance.Elements<LevelOverride>()
            .FirstOrDefault(o => o.LevelIndex?.Value == ilvl);
        var overrideLevel = lvlOverride?.GetFirstChild<Level>();
        if (overrideLevel != null) return overrideLevel;

        var abstractNumId = numInstance.AbstractNumId?.Val?.Value;
        if (abstractNumId == null) return null;
        var abstractNum = numbering.Elements<AbstractNum>()
            .FirstOrDefault(a => a.AbstractNumberId?.Value == abstractNumId);
        return abstractNum?.Elements<Level>()
            .FirstOrDefault(l => l.LevelIndex?.Value == ilvl);
    }

    private int? GetStartValue(int numId, int ilvl)
    {
        var numbering = _doc.MainDocumentPart?.NumberingDefinitionsPart?.Numbering;
        if (numbering == null) return null;

        var numInstance = numbering.Elements<NumberingInstance>()
            .FirstOrDefault(n => n.NumberID?.Value == numId);
        if (numInstance == null) return null;

        // Check level override first
        var lvlOverride = numInstance.Elements<LevelOverride>()
            .FirstOrDefault(o => o.LevelIndex?.Value == ilvl);
        if (lvlOverride?.StartOverrideNumberingValue?.Val?.Value is int overrideStart)
            return overrideStart;

        var abstractNumId = numInstance.AbstractNumId?.Val?.Value;
        if (abstractNumId == null) return null;

        var abstractNum = numbering.Elements<AbstractNum>()
            .FirstOrDefault(a => a.AbstractNumberId?.Value == abstractNumId);
        var level = abstractNum?.Elements<Level>()
            .FirstOrDefault(l => l.LevelIndex?.Value == ilvl);

        return level?.StartNumberingValue?.Val?.Value;
    }

    /// <summary>
    /// Removes numbering from a paragraph.
    /// </summary>
    private static void RemoveListStyle(Paragraph para)
    {
        var pProps = para.ParagraphProperties;
        if (pProps?.NumberingProperties != null)
        {
            pProps.NumberingProperties.Remove();
        }
    }

    /// <summary>
    /// Finds an existing NumberingInstance that uses the same list type (bullet vs ordered),
    /// scanning the last paragraph in the body to support list continuation.
    /// </summary>
    private int? FindContinuationNumId(bool isBullet)
    {
        var body = _doc.MainDocumentPart?.Document?.Body;
        if (body == null) return null;

        // Check the last paragraph in the body
        var lastPara = body.Elements<Paragraph>().LastOrDefault();
        if (lastPara == null) return null;

        var numProps = lastPara.ParagraphProperties?.NumberingProperties;
        var prevNumId = numProps?.NumberingId?.Val?.Value;
        if (prevNumId == null || prevNumId == 0) return null;

        var fmt = GetNumberingFormat(prevNumId.Value, 0);
        var prevIsBullet = fmt.ToLowerInvariant() == "bullet";
        if (prevIsBullet == isBullet)
            return prevNumId.Value;

        return null;
    }

    private void ApplyListStyle(Paragraph para, string listStyleValue, int? startValue = null, int? listLevel = null)
    {
        // Handle "none" — remove numbering
        if (listStyleValue.ToLowerInvariant() is "none" or "remove" or "clear")
        {
            RemoveListStyle(para);
            return;
        }

        var isBullet = listStyleValue.ToLowerInvariant() is "bullet" or "unordered" or "ul";

        // Try to continue from a preceding list of the same type
        var continuationNumId = FindContinuationNumId(isBullet);
        if (continuationNumId != null && startValue == null)
        {
            var pProps = para.ParagraphProperties ?? para.PrependChild(new ParagraphProperties());
            var ilvl = listLevel ?? para.ParagraphProperties?.NumberingProperties?.NumberingLevelReference?.Val?.Value ?? 0;
            pProps.NumberingProperties = new NumberingProperties
            {
                NumberingId = new NumberingId { Val = continuationNumId.Value },
                NumberingLevelReference = new NumberingLevelReference { Val = ilvl }
            };
            return;
        }

        var mainPart = _doc.MainDocumentPart!;
        var numberingPart = mainPart.NumberingDefinitionsPart;
        if (numberingPart == null)
        {
            numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>();
            numberingPart.Numbering = new Numbering();
        }
        var numbering = numberingPart.Numbering
            ?? throw new InvalidOperationException("Corrupt file: numbering data missing");

        // Determine the next available IDs
        var maxAbstractId = numbering.Elements<AbstractNum>()
            .Select(a => a.AbstractNumberId?.Value ?? 0).DefaultIfEmpty(-1).Max() + 1;
        var maxNumId = numbering.Elements<NumberingInstance>()
            .Select(n => n.NumberID?.Value ?? 0).DefaultIfEmpty(0).Max() + 1;

        // Create abstract numbering definition with 9 levels
        var abstractNum = new AbstractNum { AbstractNumberId = maxAbstractId };
        abstractNum.AppendChild(new MultiLevelType { Val = MultiLevelValues.HybridMultilevel });

        var bulletChars = new[] { "\u2022", "\u25E6", "\u25AA" }; // •, ◦, ▪

        for (int lvl = 0; lvl < 9; lvl++)
        {
            var level = new Level { LevelIndex = lvl };
            level.AppendChild(new StartNumberingValue { Val = (lvl == 0 && startValue.HasValue) ? startValue.Value : 1 });

            if (isBullet)
            {
                level.AppendChild(new NumberingFormat { Val = NumberFormatValues.Bullet });
                level.AppendChild(new LevelText { Val = bulletChars[lvl % bulletChars.Length] });
            }
            else
            {
                var fmt = (lvl % 3) switch
                {
                    0 => NumberFormatValues.Decimal,
                    1 => NumberFormatValues.LowerLetter,
                    _ => NumberFormatValues.LowerRoman
                };
                level.AppendChild(new NumberingFormat { Val = fmt });
                level.AppendChild(new LevelText { Val = $"%{lvl + 1}." });
            }

            level.AppendChild(new LevelJustification { Val = LevelJustificationValues.Left });
            level.AppendChild(new PreviousParagraphProperties(
                new Indentation { Left = ((lvl + 1) * 720).ToString(), Hanging = "360" }
            ));
            abstractNum.AppendChild(level);
        }

        // Insert AbstractNum before any NumberingInstance elements
        var firstNumInstance = numbering.GetFirstChild<NumberingInstance>();
        if (firstNumInstance != null)
            numbering.InsertBefore(abstractNum, firstNumInstance);
        else
            numbering.AppendChild(abstractNum);

        // Create numbering instance
        var numInstance = new NumberingInstance { NumberID = maxNumId };
        numInstance.AppendChild(new AbstractNumId { Val = maxAbstractId });
        numbering.AppendChild(numInstance);

        numbering.Save();

        // Apply to paragraph
        var pProps2 = para.ParagraphProperties ?? para.PrependChild(new ParagraphProperties());
        pProps2.NumberingProperties = new NumberingProperties
        {
            NumberingId = new NumberingId { Val = maxNumId },
            NumberingLevelReference = new NumberingLevelReference { Val = listLevel ?? 0 }
        };
    }

    /// <summary>
    /// Sets the start value override for a paragraph's numbering instance.
    /// </summary>
    private void SetListStartValue(Paragraph para, int startValue)
    {
        var numProps = para.ParagraphProperties?.NumberingProperties;
        var numId = numProps?.NumberingId?.Val?.Value;
        if (numId == null || numId == 0) return;

        var ilvl = numProps?.NumberingLevelReference?.Val?.Value ?? 0;
        var numbering = _doc.MainDocumentPart?.NumberingDefinitionsPart?.Numbering;
        if (numbering == null) return;

        var numInstance = numbering.Elements<NumberingInstance>()
            .FirstOrDefault(n => n.NumberID?.Value == numId);
        if (numInstance == null) return;

        // Find or create LevelOverride for this ilvl
        var lvlOverride = numInstance.Elements<LevelOverride>()
            .FirstOrDefault(o => o.LevelIndex?.Value == ilvl);
        if (lvlOverride == null)
        {
            lvlOverride = new LevelOverride { LevelIndex = ilvl };
            numInstance.AppendChild(lvlOverride);
        }
        lvlOverride.StartOverrideNumberingValue = new StartOverrideNumberingValue { Val = startValue };

        numbering.Save();
    }
}
