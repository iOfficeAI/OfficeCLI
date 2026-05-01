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
        => ResolveEffectiveRunPropertiesCore(run, para, sources: null);

    /// <summary>
    /// Same as <see cref="ResolveEffectiveRunProperties"/> but also returns
    /// a per-property provenance map: key = property name (e.g. "size",
    /// "font.eastAsia", "color"), value = path-form layer label
    /// ("/docDefaults", "/styles/Heading1", "/direct"). The "/direct" source
    /// is recorded for completeness; PopulateEffectiveRunProperties suppresses
    /// effective.* keys when the base key is set, so direct never surfaces.
    /// </summary>
    private (RunProperties Effective, Dictionary<string, string> Sources)
        ResolveEffectiveRunPropertiesWithSources(Run run, Paragraph para)
    {
        var sources = new Dictionary<string, string>();
        var effective = ResolveEffectiveRunPropertiesCore(run, para, sources);
        return (effective, sources);
    }

    private RunProperties ResolveEffectiveRunPropertiesCore(
        Run run, Paragraph para, Dictionary<string, string>? sources)
    {
        var effective = new RunProperties();

        // 1. Start with docDefaults rPr
        var docDefaults = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles?.DocDefaults;
        var defaultRPr = docDefaults?.RunPropertiesDefault?.RunPropertiesBaseStyle;
        if (defaultRPr != null)
            MergeRunProperties(effective, defaultRPr, "/docDefaults", sources);

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
            // Apply from base to derived (reverse order). Source label is the
            // styleId that actually wrote the property — not the chain top —
            // so agents can jump straight to the writer instead of walking
            // basedOn themselves.
            for (int i = chain.Count - 1; i >= 0; i--)
            {
                var styleRPr = chain[i].StyleRunProperties;
                if (styleRPr != null)
                    MergeRunProperties(effective, styleRPr,
                        $"/styles/{chain[i].StyleId?.Value}", sources);

                // CONSISTENCY(rtl-cascade): paragraph-style direction lives
                // ONLY on style pPr (<w:bidi/>) — we do not stamp <w:rtl/> on
                // styleRPr because CT_RPr requires <w:rFonts> as the first
                // child and a bare <w:rtl/> trips the validator. Lift the
                // pPr/bidi flag into the effective run's RightToLeftText so
                // runs inheriting the style still resolve effective.rtl.
                var stylePPr = chain[i].StyleParagraphProperties;
                var styleBiDi = stylePPr?.GetFirstChild<BiDi>();
                if (styleBiDi != null)
                {
                    var biVal = styleBiDi.Val;
                    bool on = biVal == null
                        || biVal.InnerText == "1"
                        || biVal.InnerText == "true"
                        || (biVal.HasValue && biVal.Value);
                    effective.RightToLeftText = on
                        ? new RightToLeftText()
                        : new RightToLeftText { Val = DocumentFormat.OpenXml.OnOffValue.FromBoolean(false) };
                    if (sources != null)
                        sources["effective.rtl"] = $"/styles/{chain[i].StyleId?.Value}";
                }
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
                    MergeRunProperties(effective, sRPr,
                        $"/styles/{rStyleChain[i].StyleId?.Value}", sources);
            }
        }

        // 3b. Lift direct pPr/<w:bidi/> into effective RightToLeftText.
        // CONSISTENCY(rtl-cascade): mirrors step-2 paragraph-style pPr/bidi
        // lift, but for the paragraph's own direct pPr (not its style).
        // Without this, a run inside a hyperlink wrapper inherits no
        // effective.rtl when the cascade only stamped <w:rtl/> on bare
        // <w:r> children — hyperlink runs are added via a path that
        // historically skipped the rtl stamp, leaving the resolver blind
        // to paragraph direction (R16-bt-3).
        var directBiDi = para.ParagraphProperties?.BiDi;
        if (directBiDi != null)
        {
            var dBiVal = directBiDi.Val;
            bool dOn = dBiVal == null
                || dBiVal.InnerText == "1"
                || dBiVal.InnerText == "true"
                || (dBiVal.HasValue && dBiVal.Value);
            effective.RightToLeftText = dOn
                ? new RightToLeftText()
                : new RightToLeftText { Val = DocumentFormat.OpenXml.OnOffValue.FromBoolean(false) };
            if (sources != null)
                sources["effective.rtl"] = "/direct";
        }

        // 4. Apply run's own direct rPr (highest priority, excluding rStyle which was resolved above)
        if (run.RunProperties != null)
            MergeRunProperties(effective, run.RunProperties, "/direct", sources);

        return effective;
    }

    private static void MergeRunProperties(
        RunProperties target,
        OpenXmlElement source,
        string? layer = null,
        Dictionary<string, string>? sources = null)
    {
        // Helper: record provenance only when both layer + sources provided.
        void Tag(string prop)
        {
            if (layer != null && sources != null) sources[prop] = layer;
        }

        // RunFonts is an attribute container — OOXML spec semantics is
        // per-slot inheritance, NOT whole-element overwrite. Previously we
        // cloned the whole rFonts element which silently dropped slots set
        // by lower-priority layers. Common Chinese-doc breakage:
        // docDefaults sets eastAsia=宋体, Heading1 only sets ascii=Calibri,
        // and the eastAsia slot would vanish from the effective merge.
        var srcFonts = source.GetFirstChild<RunFonts>();
        if (srcFonts != null)
        {
            target.RunFonts ??= new RunFonts();
            if (srcFonts.Ascii?.Value != null)
            {
                target.RunFonts.Ascii = srcFonts.Ascii.Value;
                Tag("font.ascii");
            }
            if (srcFonts.EastAsia?.Value != null)
            {
                target.RunFonts.EastAsia = srcFonts.EastAsia.Value;
                Tag("font.eastAsia");
            }
            if (srcFonts.HighAnsi?.Value != null)
            {
                target.RunFonts.HighAnsi = srcFonts.HighAnsi.Value;
                Tag("font.hAnsi");
            }
            if (srcFonts.ComplexScript?.Value != null)
            {
                target.RunFonts.ComplexScript = srcFonts.ComplexScript.Value;
                Tag("font.cs");
            }
            // Theme variants and hint propagate alongside their slot but are
            // not currently exposed in Get output, so they get no source tag.
            if (srcFonts.AsciiTheme?.HasValue == true)
                target.RunFonts.AsciiTheme = srcFonts.AsciiTheme.Value;
            if (srcFonts.EastAsiaTheme?.HasValue == true)
                target.RunFonts.EastAsiaTheme = srcFonts.EastAsiaTheme.Value;
            if (srcFonts.HighAnsiTheme?.HasValue == true)
                target.RunFonts.HighAnsiTheme = srcFonts.HighAnsiTheme.Value;
            if (srcFonts.ComplexScriptTheme?.HasValue == true)
                target.RunFonts.ComplexScriptTheme = srcFonts.ComplexScriptTheme.Value;
            if (srcFonts.Hint?.HasValue == true)
                target.RunFonts.Hint = srcFonts.Hint.Value;
        }

        var srcSize = source.GetFirstChild<FontSize>();
        if (srcSize != null)
        {
            target.FontSize = srcSize.CloneNode(true) as FontSize;
            Tag("size");
        }

        var srcBold = source.GetFirstChild<Bold>();
        if (srcBold != null)
        {
            target.Bold = srcBold.CloneNode(true) as Bold;
            Tag("bold");
        }

        var srcItalic = source.GetFirstChild<Italic>();
        if (srcItalic != null)
        {
            target.Italic = srcItalic.CloneNode(true) as Italic;
            Tag("italic");
        }

        var srcUnderline = source.GetFirstChild<Underline>();
        if (srcUnderline != null)
        {
            target.Underline = srcUnderline.CloneNode(true) as Underline;
            Tag("underline");
        }

        var srcStrike = source.GetFirstChild<Strike>();
        if (srcStrike != null)
        {
            target.Strike = srcStrike.CloneNode(true) as Strike;
            Tag("strike");
        }

        var srcDStrike = source.GetFirstChild<DoubleStrike>();
        if (srcDStrike != null)
            target.DoubleStrike = srcDStrike.CloneNode(true) as DoubleStrike;

        var srcColor = source.GetFirstChild<Color>();
        if (srcColor != null)
        {
            target.Color = srcColor.CloneNode(true) as Color;
            Tag("color");
        }

        var srcHighlight = source.GetFirstChild<Highlight>();
        if (srcHighlight != null)
        {
            target.Highlight = srcHighlight.CloneNode(true) as Highlight;
            Tag("highlight");
        }

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
        var (effective, sources) = ResolveEffectiveRunPropertiesWithSources(firstRun, para);
        EmitEffectiveRunProperties(node, effective, sources);
        // Resolve effective paragraph properties from style chain
        ResolveEffectiveParagraphStyleProperties(node, para);
    }

    /// <summary>
    /// Populates effective.* format keys on a run node for properties not explicitly set.
    /// </summary>
    private void PopulateEffectiveRunProperties(DocumentNode node, Run run, Paragraph para)
    {
        var (effective, sources) = ResolveEffectiveRunPropertiesWithSources(run, para);
        EmitEffectiveRunProperties(node, effective, sources);
    }

    /// <summary>
    /// Shared emit logic for run-level effective.* properties. Each property
    /// is suppressed when the corresponding base key is already set (run
    /// owns it directly). When emitted, also writes effective.X.src pointing
    /// to the path of the writing layer (e.g. "/styles/Heading1",
    /// "/docDefaults"). Per-slot RunFonts surface as effective.font.ascii /
    /// .eastAsia / .hAnsi / .cs — each independently sourced.
    /// </summary>
    private static void EmitEffectiveRunProperties(
        DocumentNode node,
        RunProperties effective,
        Dictionary<string, string> sources)
    {
        void EmitSrc(string effectiveKey, string sourceKey)
        {
            if (sources.TryGetValue(sourceKey, out var src) && src != "/direct")
                node.Format[effectiveKey + ".src"] = src;
        }

        // size
        if (!node.Format.ContainsKey("size") && effective.FontSize?.Val?.Value != null)
        {
            var sz = int.Parse(effective.FontSize.Val.Value) / 2.0;
            node.Format["effective.size"] = $"{sz:0.##}pt";
            EmitSrc("effective.size", "size");
        }

        // Per-slot font: each slot independently honors style cascade and
        // is suppressed only when that specific slot is set on the run.
        // CONSISTENCY(canonical-keys): mirrors the 4-slot direct readback in
        // Navigation.cs:1186-1192.
        if (!node.Format.ContainsKey("font.ascii") && !node.Format.ContainsKey("font")
            && effective.RunFonts?.Ascii?.Value != null)
        {
            node.Format["effective.font.ascii"] = effective.RunFonts.Ascii.Value;
            EmitSrc("effective.font.ascii", "font.ascii");
        }
        if (!node.Format.ContainsKey("font.eastAsia") && !node.Format.ContainsKey("font")
            && effective.RunFonts?.EastAsia?.Value != null)
        {
            node.Format["effective.font.eastAsia"] = effective.RunFonts.EastAsia.Value;
            EmitSrc("effective.font.eastAsia", "font.eastAsia");
        }
        if (!node.Format.ContainsKey("font.hAnsi") && !node.Format.ContainsKey("font")
            && effective.RunFonts?.HighAnsi?.Value != null)
        {
            node.Format["effective.font.hAnsi"] = effective.RunFonts.HighAnsi.Value;
            EmitSrc("effective.font.hAnsi", "font.hAnsi");
        }
        if (!node.Format.ContainsKey("font.cs") && !node.Format.ContainsKey("font")
            && effective.RunFonts?.ComplexScript?.Value != null)
        {
            node.Format["effective.font.cs"] = effective.RunFonts.ComplexScript.Value;
            EmitSrc("effective.font.cs", "font.cs");
        }

        if (!node.Format.ContainsKey("bold") && effective.Bold != null)
        {
            node.Format["effective.bold"] = true;
            EmitSrc("effective.bold", "bold");
        }

        if (!node.Format.ContainsKey("italic") && effective.Italic != null)
        {
            node.Format["effective.italic"] = true;
            EmitSrc("effective.italic", "italic");
        }

        if (effective.RightToLeftText != null)
        {
            // Honor explicit <w:rtl w:val="0"/> off-override. RightToLeftText is
            // an OnOff element: missing Val means true, Val="0"/"false" means
            // explicit off (used to defeat an inherited docDefaults rtl=true).
            // Emitted even when direct `rtl` is also present so callers can see
            // both the direct value and the cascade-resolved effective state —
            // matters for RTL because docDefaults.rtl is the common inheritance
            // path that callers want to verify against the per-run override.
            var rtlVal = effective.RightToLeftText.Val;
            node.Format["effective.rtl"] = rtlVal == null ? true : rtlVal.Value;
            EmitSrc("effective.rtl", "rtl");
        }

        if (!node.Format.ContainsKey("color"))
        {
            if (effective.Color?.Val?.Value != null)
            {
                node.Format["effective.color"] = ParseHelpers.FormatHexColor(effective.Color.Val.Value);
                EmitSrc("effective.color", "color");
            }
            else if (effective.Color?.ThemeColor?.HasValue == true)
            {
                node.Format["effective.color"] = effective.Color.ThemeColor.InnerText;
                EmitSrc("effective.color", "color");
            }
        }

        if (!node.Format.ContainsKey("underline") && effective.Underline?.Val != null)
        {
            node.Format["effective.underline"] = effective.Underline.Val.InnerText;
            EmitSrc("effective.underline", "underline");
        }

        if (!node.Format.ContainsKey("strike") && effective.Strike != null)
        {
            node.Format["effective.strike"] = true;
            EmitSrc("effective.strike", "strike");
        }

        if (!node.Format.ContainsKey("highlight") && effective.Highlight?.Val != null)
        {
            node.Format["effective.highlight"] = effective.Highlight.Val.InnerText;
            EmitSrc("effective.highlight", "highlight");
        }
    }

    /// <summary>
    /// Resolves paragraph-level properties (alignment, spacing) from the paragraph style chain.
    /// </summary>
    private void ResolveEffectiveParagraphStyleProperties(DocumentNode node, Paragraph para)
    {
        // R9-1: do NOT early-return when the paragraph has no style. Numbering
        // lvl pPr.bidi is a separate cascade layer that applies even when the
        // paragraph is style-less, and table/docDefaults fallbacks downstream
        // also apply unconditionally.
        var styleId = para.ParagraphProperties?.ParagraphStyleId?.Val?.Value;

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

        // Apply from base to derived (reverse order), collecting effective
        // paragraph properties + provenance. Source label is the styleId
        // that actually wrote the property (the most-derived layer that
        // touched it), not the chain top.
        string? alignment = null, alignSrc = null;
        string? spaceBefore = null, spaceBeforeSrc = null;
        string? spaceAfter = null, spaceAfterSrc = null;
        string? lineSpacing = null, lineSpacingSrc = null;
        string? direction = null, directionSrc = null;

        for (int i = chain.Count - 1; i >= 0; i--)
        {
            var ppr = chain[i].StyleParagraphProperties;
            if (ppr == null) continue;
            var layer = $"/styles/{chain[i].StyleId?.Value}";

            if (ppr.Justification?.Val != null)
            {
                var txt = ppr.Justification.Val.InnerText;
                alignment = txt == "both" ? "justify" : txt;
                alignSrc = layer;
            }
            if (ppr.SpacingBetweenLines?.Before?.Value != null)
            {
                spaceBefore = SpacingConverter.FormatWordSpacing(ppr.SpacingBetweenLines.Before.Value);
                spaceBeforeSrc = layer;
            }
            if (ppr.SpacingBetweenLines?.After?.Value != null)
            {
                spaceAfter = SpacingConverter.FormatWordSpacing(ppr.SpacingBetweenLines.After.Value);
                spaceAfterSrc = layer;
            }
            if (ppr.SpacingBetweenLines?.Line?.Value != null)
            {
                lineSpacing = SpacingConverter.FormatWordLineSpacing(
                    ppr.SpacingBetweenLines.Line.Value,
                    ppr.SpacingBetweenLines.LineRule?.InnerText);
                lineSpacingSrc = layer;
            }
            // R8-1: paragraph-scope effective.direction. Mirrors the
            // run-level effective.rtl pattern but reads <w:bidi/> from the
            // style-chain pPr. TryReadOnOff defends against the malformed
            // attribute case (R8-fuzz-5).
            var styleBidi = ppr.GetFirstChild<BiDi>();
            if (styleBidi != null)
            {
                var on = TryReadOnOff(styleBidi.Val);
                if (on.HasValue)
                {
                    direction = on.Value ? "rtl" : "ltr";
                    directionSrc = layer;
                }
            }
        }

        if (!node.Format.ContainsKey("align") && !node.Format.ContainsKey("alignment") && alignment != null)
        {
            node.Format["effective.alignment"] = alignment;
            if (alignSrc != null) node.Format["effective.alignment.src"] = alignSrc;
        }
        if (!node.Format.ContainsKey("spaceBefore") && spaceBefore != null)
        {
            node.Format["effective.spaceBefore"] = spaceBefore;
            if (spaceBeforeSrc != null) node.Format["effective.spaceBefore.src"] = spaceBeforeSrc;
        }
        if (!node.Format.ContainsKey("spaceAfter") && spaceAfter != null)
        {
            node.Format["effective.spaceAfter"] = spaceAfter;
            if (spaceAfterSrc != null) node.Format["effective.spaceAfter.src"] = spaceAfterSrc;
        }
        if (!node.Format.ContainsKey("lineSpacing") && lineSpacing != null)
        {
            node.Format["effective.lineSpacing"] = lineSpacing;
            if (lineSpacingSrc != null) node.Format["effective.lineSpacing.src"] = lineSpacingSrc;
        }
        // R9-1: numbering lvl pPr.bidi layer. A list-bound paragraph that
        // does not have a direct or style-chain bidi must still inherit
        // pPr.bidi from its abstractNum.lvl[ilvl]. This sits between the
        // style chain and the table-style fallback because Word's
        // numbering definition layers between paragraph style and the
        // enclosing table — see CT_PPr semantics.
        if (!node.Format.ContainsKey("direction") && direction == null)
        {
            var resolved = ResolveNumPrFromStyle(para);
            if (resolved != null)
            {
                var (numId, ilvl) = resolved.Value;
                var numbering = _doc.MainDocumentPart?.NumberingDefinitionsPart?.Numbering;
                var inst = numbering?.Elements<NumberingInstance>()
                    .FirstOrDefault(n => n.NumberID?.Value == numId);
                var absId = inst?.AbstractNumId?.Val?.Value;
                var abs = absId != null
                    ? numbering!.Elements<AbstractNum>()
                        .FirstOrDefault(a => a.AbstractNumberId?.Value == absId.Value)
                    : null;
                var lvl = abs?.Elements<Level>()
                    .FirstOrDefault(l => l.LevelIndex?.Value == ilvl);
                var lvlBidi = lvl?.PreviousParagraphProperties?.GetFirstChild<BiDi>();
                if (lvlBidi != null)
                {
                    var on = TryReadOnOff(lvlBidi.Val);
                    if (on.HasValue)
                    {
                        direction = on.Value ? "rtl" : "ltr";
                        directionSrc = $"/numbering/abstractNum[@id={absId}]/level[{ilvl}]";
                    }
                }
            }
        }
        // R8-1: paragraph-scope effective.direction. After the paragraph-style
        // chain, fall back to the enclosing table style's pPr.bidi (paragraphs
        // inside a table cell inherit from tblPr-style.pPr) and finally to
        // docDefaults pPrDefault.bidi. PPT has had this since R5.
        if (!node.Format.ContainsKey("direction") && direction == null)
        {
            // Enclosing table style
            var tbl = para.Ancestors<Table>().FirstOrDefault();
            var tblStyleId = tbl?.GetFirstChild<TableProperties>()?.TableStyle?.Val?.Value;
            if (tblStyleId != null)
            {
                var tblStyle = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles
                    ?.Elements<Style>().FirstOrDefault(s => s.StyleId?.Value == tblStyleId);
                var tblPpr = tblStyle?.StyleParagraphProperties;
                var tblBidi = tblPpr?.GetFirstChild<BiDi>();
                if (tblBidi != null)
                {
                    var on = TryReadOnOff(tblBidi.Val);
                    if (on.HasValue)
                    {
                        direction = on.Value ? "rtl" : "ltr";
                        directionSrc = $"/styles/{tblStyleId}";
                    }
                }
            }
        }
        // R20-bt-2: enclosing table's own tblPr/<w:bidiVisual/> cascades to
        // every paragraph in every cell — independent of the table-style
        // layer above (a table can carry direct bidiVisual without referencing
        // any RTL table style). Sits between the table-style layer and the
        // section layer so direct table bidiVisual beats sectPr bidi but is
        // beaten by an explicit pPr.bidi or a paragraph-style bidi.
        if (!node.Format.ContainsKey("direction") && direction == null)
        {
            var ownTbl = para.Ancestors<Table>().FirstOrDefault();
            if (ownTbl?.GetFirstChild<TableProperties>()?.GetFirstChild<BiDiVisual>() != null)
            {
                direction = "rtl";
                // Locate 1-based table index in document order for src.
                var tbls = _doc.MainDocumentPart?.Document?.Body?.Descendants<Table>().ToList();
                var tblIdx = tbls?.FindIndex(t => ReferenceEquals(t, ownTbl)) ?? -1;
                directionSrc = tblIdx >= 0 ? $"/body/tbl[{tblIdx + 1}]" : "/body/tbl";
            }
        }
        if (!node.Format.ContainsKey("direction") && direction == null)
        {
            // R15-bt-3: enclosing section's <w:bidi/> on sectPr cascades
            // to every paragraph in the section. The section that owns a
            // paragraph is the first paragraph-level sectPr that comes
            // after it in document order, falling back to the body-level
            // (final) sectPr if none does.
            var owningSect = FindOwningSectionProperties(para);
            if (owningSect != null && owningSect.GetFirstChild<BiDi>() != null)
            {
                // sectPr <w:bidi/> has no Val attribute defaulting to true
                // (CT_OnOff default-true). Honor explicit Val=false too.
                var on = TryReadOnOff(owningSect.GetFirstChild<BiDi>()?.Val);
                if (on != true) on = on ?? true;
                direction = on.Value ? "rtl" : "ltr";
                // Locate the section's 1-based document-order index for src.
                var sects = FindSectionProperties();
                var idx = sects.FindIndex(s => ReferenceEquals(s, owningSect));
                directionSrc = idx >= 0
                    ? $"/section[{idx + 1}]"
                    : "/body/sectPr[1]";
            }
        }
        if (!node.Format.ContainsKey("direction") && direction == null)
        {
            // docDefaults pPrDefault.bidi
            var docDefaults = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles?.DocDefaults;
            var pPrDefault = docDefaults?.ParagraphPropertiesDefault?.ParagraphPropertiesBaseStyle;
            var ddBidi = pPrDefault?.GetFirstChild<BiDi>();
            if (ddBidi != null)
            {
                var on = TryReadOnOff(ddBidi.Val);
                if (on.HasValue)
                {
                    direction = on.Value ? "rtl" : "ltr";
                    directionSrc = "/docDefaults";
                }
            }
        }
        if (!node.Format.ContainsKey("direction") && direction != null)
        {
            node.Format["effective.direction"] = direction;
            if (directionSrc != null) node.Format["effective.direction.src"] = directionSrc;
            // R21-bt-1 + R21-bt-2: cascade-uniform effective.rtl. The
            // style-chain path (ResolveEffectiveRunPropertiesCore) already
            // lifts pPr.bidi into effective.rtl on style-style cascades.
            // Section / table-bidiVisual / table-style / docDefaults /
            // numbering layers were missing that lift, so paragraphs
            // inheriting RTL from any of these emitted only effective.direction.
            // Emit effective.rtl alongside effective.direction so callers see
            // the same surface regardless of the originating cascade layer.
            if (!node.Format.ContainsKey("effective.rtl"))
            {
                node.Format["effective.rtl"] = direction == "rtl";
                if (directionSrc != null) node.Format["effective.rtl.src"] = directionSrc;
            }
        }
        // R21-fuzz-2: paragraph carries its own pPr.bidi. Emit
        // effective.direction + .src=self for cascade-uniform readback so
        // downstream consumers always have an effective.direction key
        // regardless of whether the resolved direction came from the
        // paragraph itself or an inherited cascade layer.
        else if (node.Format.ContainsKey("direction"))
        {
            var ownBidi = para.ParagraphProperties?.GetFirstChild<BiDi>();
            if (ownBidi != null)
            {
                var on = TryReadOnOff(ownBidi.Val);
                if (on.HasValue)
                {
                    node.Format["effective.direction"] = on.Value ? "rtl" : "ltr";
                    var bodyParas = _doc.MainDocumentPart?.Document?.Body?
                        .Descendants<Paragraph>().ToList();
                    var pIdx = bodyParas?.FindIndex(p => ReferenceEquals(p, para)) ?? -1;
                    node.Format["effective.direction.src"] = pIdx >= 0
                        ? $"/body/p[{pIdx + 1}]"
                        : "/body/p";
                }
            }
        }
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
    /// scanning the last paragraph in the same container (body / header / footer) as the
    /// paragraph being styled. Header/footer paragraphs were previously falling through to
    /// the body scan, which always missed (body has no list paras when adding to a header)
    /// and a fresh numId was minted per paragraph.
    /// </summary>
    private int? FindContinuationNumId(bool isBullet, Paragraph? targetPara = null, OpenXmlElement? containerHint = null)
    {
        // Resolution order for the scan container:
        //   1. explicit hint from caller (Add path passes the still-detached para's
        //      parent — the para hasn't been appended yet so ancestor walk fails)
        //   2. ancestor walk on targetPara (Set path or already-inserted paras)
        //   3. body fallback
        OpenXmlElement? container = containerHint;
        if (container == null && targetPara != null)
        {
            container = targetPara.Ancestors<Header>().FirstOrDefault()
                ?? targetPara.Ancestors<Footer>().FirstOrDefault()
                ?? (OpenXmlElement?)_doc.MainDocumentPart?.Document?.Body;
        }
        container ??= _doc.MainDocumentPart?.Document?.Body;
        if (container == null) return null;

        var lastPara = container.Elements<Paragraph>().LastOrDefault(p => !ReferenceEquals(p, targetPara));
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

    private void ApplyListStyle(Paragraph para, string listStyleValue, int? startValue = null, int? listLevel = null, OpenXmlElement? containerHint = null)
    {
        // Handle "none" — remove numbering
        if (listStyleValue.ToLowerInvariant() is "none" or "remove" or "clear")
        {
            RemoveListStyle(para);
            return;
        }

        var isBullet = listStyleValue.ToLowerInvariant() is "bullet" or "unordered" or "ul";

        // Try to continue from a preceding list of the same type — pass the target
        // paragraph so the scan walks the right container (body / header / footer).
        // The Add path supplies containerHint because the para is still detached
        // when ApplyListStyle runs (insertion happens after).
        var continuationNumId = FindContinuationNumId(isBullet, para, containerHint);
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
