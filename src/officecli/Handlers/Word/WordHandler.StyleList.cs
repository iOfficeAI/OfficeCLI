// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
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

        // 1b. Table-style run properties (base rPr + matching conditional-format
        // rPr) for the cell the run lives in. Per ECMA-376 §17.7.2 the run-
        // property cascade places table styles directly above docDefaults and
        // below paragraph/character styles, so merge the layers (lowest→highest
        // priority, already ordered by the renderer) here. The HTML renderer
        // stashes them on _ctx around RenderCellChild; null/empty on the body
        // path. Without this a firstRow/band cell's <w:caps/> + white <w:color/>
        // never reach the run and it falls back to docDefaults (e.g. the Invoice
        // "PAYMENT OPTIONS" header rendered lowercase grey instead of white caps).
        var cellTableStyleRPr = _ctx?.CurrentCellTableStyleRunProps;
        if (cellTableStyleRPr != null)
        {
            foreach (var layerRPr in cellTableStyleRPr)
                MergeRunProperties(effective, layerRPr, "/tableStyle", sources);
        }

        // 2. Walk paragraph style basedOn chain (collect in order, apply from base to derived)
        var styleId = para.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        // OOXML §17.7.4.17: a paragraph without an explicit <w:pStyle>
        // implicitly inherits the document's default paragraph style
        // (the one carrying w:default="1" on w:type="paragraph"). Resolve
        // it here so the run properties from that style still cascade.
        if (styleId == null)
        {
            var defaultStyle = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles
                ?.Elements<Style>().FirstOrDefault(s =>
                    s.Type?.Value == StyleValues.Paragraph
                    && s.Default?.Value == true);
            styleId = defaultStyle?.StyleId?.Value;
        }
        if (styleId != null)
        {
            var chain = new List<Style>();
            var visited = new HashSet<string>();
            var currentStyleId = styleId;
            while (currentStyleId != null && visited.Add(currentStyleId))
            {
                var style = FindStyleById(currentStyleId);
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
                var rStyle = FindStyleById(curRStyleId);
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

        // Vertical run position (w:position, half-points; raised/lowered text)
        var srcPosition = source.GetFirstChild<Position>();
        if (srcPosition != null)
            target.Position = srcPosition.CloneNode(true) as Position;

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
        return $"{int.Parse(size) / 2.0:0.##}pt"; // half-points; keep .5 (e.g. sz=21 -> 10.5pt)
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
            var style = FindStyleById(currentStyleId);
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
                spaceBefore = SpacingConverter.FormatWordSpacingNonNegative(ppr.SpacingBetweenLines.Before.Value);
                spaceBeforeSrc = layer;
            }
            if (ppr.SpacingBetweenLines?.After?.Value != null)
            {
                spaceAfter = SpacingConverter.FormatWordSpacingNonNegative(ppr.SpacingBetweenLines.After.Value);
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
                var abs = FindAbstractNum(numbering, absId);
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
                var tblStyle = FindStyleById(tblStyleId);
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
        // numId and ilvl inherit PER-PROPERTY through the pPr cascade (issue
        // #180): Word's multilevel-list-linked heading styles routinely declare
        // only <w:ilvl> on a child style and inherit <w:numId> from its basedOn
        // parent. Track the two independently — the first-seen (closest) value
        // of each wins, and they may come from different links in the chain.
        int? ilvl = null;

        // 1. Direct numPr on the paragraph wins.
        var numProps = para.ParagraphProperties?.NumberingProperties;
        if (numProps != null)
        {
            var nid = numProps.NumberingId?.Val?.Value;
            // numId=0 is OOXML's "no numbering" sentinel — it explicitly removes
            // any inherited list reference, so stop here rather than walking the
            // style chain (which would resurrect the based-on style's numbering).
            if (nid == 0) return null;
            ilvl = numProps.NumberingLevelReference?.Val?.Value;
            if (nid != null)
                return (nid.Value, ilvl ?? 0);
        }

        // 2. Walk the style chain through BasedOn references.
        var styleId = para.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        if (styleId == null) return null;

        if (_doc.MainDocumentPart?.StyleDefinitionsPart?.Styles == null) return null;

        var visited = new HashSet<string>();
        while (styleId != null && visited.Add(styleId))
        {
            var style = FindStyleById(styleId);
            if (style == null) break;

            var styleNumPr = style.StyleParagraphProperties?.NumberingProperties;
            if (styleNumPr != null)
            {
                var nid = styleNumPr.NumberingId?.Val?.Value;
                // A style that sets numId=0 explicitly cancels numbering for
                // everything based on it (e.g. a title style based on Heading1
                // that opts out of the heading's auto-number). Stop the walk so
                // the based-on style's numbering is not inherited.
                if (nid == 0) return null;
                ilvl ??= styleNumPr.NumberingLevelReference?.Val?.Value;
                if (nid != null)
                    return (nid.Value, ilvl ?? 0);
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
            // CONSISTENCY(ilvl-clamp): same guard as GetListPrefix.
            if (ilvl < 0) ilvl = 0; else if (ilvl > 8) ilvl = 8;
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
        // CONSISTENCY(ilvl-clamp): same guard as GetListPrefix.
        var ilvlRc = ilvlR < 0 ? 0 : ilvlR > 8 ? 8 : ilvlR;
        var numFmtR = GetNumberingFormat(numId, ilvlRc);
        return numFmtR.ToLowerInvariant() == "bullet" ? "bullet" : "ordered";
    }

    // ---- Shared ordered-list numbering (text + HTML preview) -------------
    //
    // CONSISTENCY(list-marker): the plain-text walker (WordHandler.View.cs) and
    // the HTML preview (WordHandler.HtmlPreview.cs) compute ordered-list markers
    // through this ONE state machine, so the two views can never diverge on
    // auto-numbering (the bug behind GitHub #161, where text fell back to "• ").
    // The HTML path additionally owns its <ol>/<ul> nesting + CSS; only the
    // counter seeding / advancing / glyph rendering lives here.

    /// <summary>
    /// Running ordered-list counter state for one render walk. Holds the same
    /// three dictionaries the HTML preview has always used:
    ///   • <see cref="OlCountPerLevel"/> — running item count per ilvl in the
    ///     active list run (cleared when the list breaks to a different numId).
    ///   • <see cref="MultiLevelCounters"/> — the values fed into the lvlText
    ///     template ("%1.%2.") when rendering a marker.
    ///   • <see cref="AbsNumLevelCounters"/> — per-abstractNum running counts
    ///     that survive a numId change, so sibling num instances sharing one
    ///     abstractNum continue numbering (Word's default list continuation).
    /// <see cref="CurrentNumId"/> drives the text path's break detection; the
    /// HTML path tracks its own currentNumId alongside its nesting state.
    /// </summary>
    private sealed class OrderedListNumberingState
    {
        public int? CurrentNumId;
        public readonly Dictionary<int, int> OlCountPerLevel = new();
        public readonly Dictionary<int, int> MultiLevelCounters = new();
        public readonly Dictionary<int, Dictionary<int, int>> AbsNumLevelCounters = new();
        // Levels of the active num that carry an un-consumed per-instance
        // <w:startOverride> (or override-embedded <w:lvl><w:start>). Re-armed on
        // every numId switch; a level fires its override the FIRST time it is
        // directly advanced under that num, then is removed. This lets the
        // override win over abstractNum continuation even when a deeper item
        // appeared first and carried the level over (ECMA-376 §17.9.7).
        public readonly HashSet<int> PendingInstanceOverride = new();
    }

    /// <summary>
    /// Seed value (one BELOW the first rendered number) for an ordered level,
    /// using the four-way precedence Word follows: in-run running count →
    /// explicit startOverride → abstractNum continuation → level start value.
    /// </summary>
    // INT_MIN guard: seed - 1 would wrap (unchecked) to int.MaxValue, which then
    // permanently triggers the saturation branch in AdvanceOrderedCounter and
    // freezes every item at int.MaxValue. Seed at INT_MIN instead so the counter
    // still advances upward (graceful for an absurd start).
    private static int SeedBelow(int v) => v == int.MinValue ? int.MinValue : v - 1;

    private int SeedOrderedStart(OrderedListNumberingState st, int numId, int? absId, int forIlvl, bool autoInit = false)
    {
        if (st.OlCountPerLevel.TryGetValue(forIlvl, out var prev) && prev > 0)
            return prev;
        // A per-num-instance level override (a <w:lvlOverride> carrying a
        // <w:startOverride> OR an embedded <w:lvl> with its own <w:start>) means
        // "this num restarts its numbering at that value" and takes precedence
        // over continuation from a sibling num sharing the abstractNum. Without
        // checking it BEFORE the continuation branch below, an overridden num
        // sharing an abstractNum continued the sibling's count (e.g. 4,5,6)
        // instead of starting fresh at the override (5,6,7). This is a restart
        // marker, so it applies only on explicit visitation — autoInit (a deeper
        // level implicitly initializing this skipped level) ignores it and uses
        // the level's intrinsic <w:start>, matching Word.
        int? continuedRun = null;
        if (absId.HasValue
            && st.AbsNumLevelCounters.TryGetValue(absId.Value, out var byIlvl)
            && byIlvl.TryGetValue(forIlvl, out var running) && running > 0)
            continuedRun = running;
        if (!autoInit)
        {
            var ovr = GetNumInstanceOverrideStart(numId, forIlvl);
            // A bare startOverride restarts the level to that value — but only on
            // the level's INITIAL start or an lvlRestart-triggered restart
            // (ECMA-376 §17.9.7). A level that is CONTINUED from a sibling num
            // sharing the abstractNum AND never restarts (lvlRestart="0") meets
            // neither condition, so it keeps its running count rather than jumping
            // to the override value (matches Word). A fresh level (not continued)
            // still takes the override as its initial start.
            if (ovr.HasValue
                && !(continuedRun.HasValue && GetEffectiveLvlRestart(numId, forIlvl) == 0))
                return SeedBelow(ovr.Value);
        }
        if (continuedRun.HasValue)
            return continuedRun.Value;
        var rawStart = autoInit
            ? GetLevelDefinedStart(numId, forIlvl)
            : (GetStartValue(numId, forIlvl) ?? 1);
        return SeedBelow(rawStart);
    }

    // The level's intrinsic start value, EXCLUDING <w:startOverride> (which is a
    // restart marker, not part of the level definition). An embedded <w:lvl>
    // override replaces the level definition, so its own <w:start> governs;
    // otherwise the abstractNum level's <w:start>. Default 1.
    private int GetLevelDefinedStart(int numId, int ilvl)
    {
        var numbering = _doc.MainDocumentPart?.NumberingDefinitionsPart?.Numbering;
        var inst = numbering?.Elements<NumberingInstance>()
            .FirstOrDefault(n => n.NumberID?.Value == numId);
        var lvlOverride = inst?.Elements<LevelOverride>()
            .FirstOrDefault(o => o.LevelIndex?.Value == ilvl);
        if (lvlOverride?.GetFirstChild<Level>() is Level emb)
            return emb.StartNumberingValue?.Val?.Value ?? 1;
        var absId = inst?.AbstractNumId?.Val?.Value;
        var abs = FindAbstractNum(numbering, absId);
        var lvl = abs?.Elements<Level>()
            .FirstOrDefault(l => l.LevelIndex?.Value == ilvl);
        return lvl?.StartNumberingValue?.Val?.Value ?? 1;
    }

    // The per-num-instance override start for a level, or null when this num has
    // no <w:lvlOverride> for it. An embedded <w:lvl> replaces the level (its own
    // <w:start> governs, ECMA-376 §17.9.7 — startOverride is then ignored);
    // otherwise the bare <w:startOverride> value. Non-null signals "this num
    // restarts here", which outranks shared-abstractNum continuation.
    private int? GetNumInstanceOverrideStart(int numId, int ilvl)
    {
        var numbering = _doc.MainDocumentPart?.NumberingDefinitionsPart?.Numbering;
        var inst = numbering?.Elements<NumberingInstance>()
            .FirstOrDefault(n => n.NumberID?.Value == numId);
        var ovr = inst?.Elements<LevelOverride>()
            .FirstOrDefault(o => o.LevelIndex?.Value == ilvl);
        if (ovr == null) return null;
        // Only a bare <w:startOverride> forces a RESTART at that value
        // (ECMA-376 §17.9.7). An embedded <w:lvl> override merely REDEFINES the
        // level (its format + its own <w:start> for future lvlRestart events);
        // the counter CONTINUES from the shared abstractNum, it does not restart
        // (Word: a num sharing abstractNum that ends at 7 then carries an
        // embedded-lvl override renders the next item as 8 in the new format,
        // not the override's start). The embedded start is still honored for a
        // fresh (non-continuation) list via normal start resolution — GetLevel
        // prefers the override lvl — so it is not lost; it just must not trigger
        // an override-restart here.
        return ovr.StartOverrideNumberingValue?.Val?.Value;
    }

    /// <summary>
    /// Advance the counter for an ordered item at (numId, ilvl): increment this
    /// level, then reset deeper levels per each deeper level's
    /// <c>&lt;w:lvlRestart&gt;</c> (ECMA-376 §17.9.6), and mirror the running
    /// count into the per-abstractNum store so a later sibling num can continue
    /// from it.
    /// </summary>
    private void AdvanceOrderedCounter(OrderedListNumberingState st, int numId, int? absId, int ilvl)
    {
        // ECMA-376 §17.9: jumping to a deeper level auto-initializes every
        // shallower level that was never visited to its start value. Word
        // renders a skipped intermediate level as its start (0->2 yields
        // "1.1.1.", not "1.0.1.") AND treats the implicit init as a visit, so a
        // later explicit visit to that level increments from start ("1.2.").
        // Seed each un-visited shallower level before advancing the current one.
        for (int j = 0; j < ilvl; j++)
        {
            if (st.OlCountPerLevel.ContainsKey(j)) continue;
            int sj;
            // A shallower level missing from OlCountPerLevel is one of two things:
            //   (a) never visited at all (jump-to-deeper within a list) → it must
            //       APPEAR as its start value, so seed-below + 1 = start.
            //   (b) already visited under a sibling num sharing this abstractNum,
            //       then dropped when the numId switch cleared the in-run counters
            //       (GetListPrefix). Its running count survives in
            //       AbsNumLevelCounters; the parent did NOT advance merely because
            //       a deeper item appeared first under the new num, so RESTORE that
            //       value verbatim (no +1). Adding +1 here over-counted every
            //       subsequent marker (parent shown one level too high) — e.g. a
            //       0,0,0,1 / switch / 1,0,0,1,0 walk rendered the post-switch
            //       parent as d,e,f instead of Word's c,d,e.
            if (absId.HasValue
                && st.AbsNumLevelCounters.TryGetValue(absId.Value, out var carried)
                && carried.TryGetValue(j, out var run) && run > 0)
            {
                sj = run;
            }
            else
            {
                sj = SeedOrderedStart(st, numId, absId, j, autoInit: true) + 1;
            }
            st.OlCountPerLevel[j] = sj;
            st.MultiLevelCounters[j] = sj;
            if (absId.HasValue)
            {
                if (!st.AbsNumLevelCounters.TryGetValue(absId.Value, out var bi))
                {
                    bi = new Dictionary<int, int>();
                    st.AbsNumLevelCounters[absId.Value] = bi;
                }
                bi[j] = sj;
            }
        }
        var seed = SeedOrderedStart(st, numId, absId, ilvl);
        // A per-instance startOverride fires on the FIRST direct advance of its
        // level under the new num — taking the override value verbatim (a
        // restart), not continuing from the carried-over count. The override is
        // consumed once; later advances of the same level continue normally.
        if (st.PendingInstanceOverride.Remove(ilvl)
            && GetNumInstanceOverrideStart(numId, ilvl) is int ovrStart)
        {
            st.OlCountPerLevel[ilvl] = ovrStart;
        }
        else
        {
            var prev = st.OlCountPerLevel.GetValueOrDefault(ilvl, seed);
            // Saturate at int.MaxValue to avoid overflow when w:start is INT_MAX.
            st.OlCountPerLevel[ilvl] = prev == int.MaxValue ? int.MaxValue : prev + 1;
        }
        st.MultiLevelCounters[ilvl] = st.OlCountPerLevel[ilvl];
        for (int lk = ilvl + 1; lk <= 8; lk++)
        {
            if (!ShouldRestartDeeperLevel(numId, ilvl, lk)) continue;
            // Restart by REMOVING the counter, not zeroing it: the next advance
            // then re-seeds via SeedOrderedStart, which honors the level's
            // <w:start> / <w:startOverride>. Zeroing made GetValueOrDefault
            // return the stored 0, so the level always restarted at 1 — wrong
            // for any level whose start != 1 (e.g. start=5 restarted to 1, not
            // 5). Verified vs real Word.
            st.OlCountPerLevel.Remove(lk);
            st.MultiLevelCounters.Remove(lk);
        }
        if (absId.HasValue)
        {
            if (!st.AbsNumLevelCounters.TryGetValue(absId.Value, out var byIlvl))
            {
                byIlvl = new Dictionary<int, int>();
                st.AbsNumLevelCounters[absId.Value] = byIlvl;
            }
            byIlvl[ilvl] = st.OlCountPerLevel[ilvl];
            for (int lk = ilvl + 1; lk <= 8; lk++)
            {
                if (!ShouldRestartDeeperLevel(numId, ilvl, lk)) continue;
                byIlvl.Remove(lk); // mirror the OlCountPerLevel restart: re-seed, don't zero
            }
        }
    }

    /// <summary>
    /// ECMA-376 §17.9.6 &lt;w:lvlRestart&gt;: the current level restarts whenever the
    /// level numbered R (1-based) <em>or any level above it (lower index = shallower)</em>
    /// is used. So when the level at <paramref name="ilvl"/> (0-based) increments,
    /// deeper level <paramref name="deeperIlvl"/> restarts iff its lvlRestart value
    /// R satisfies (ilvl+1) &lt;= R — i.e. the incrementing level is R or shallower.
    ///   • R absent → default R = deeperIlvl (its own 0-based index): restarts on
    ///     every STRICTLY shallower level (the standard outline default). This is
    ///     mathematically identical to the historic "(ilvl+1) &gt;= 1" default, so
    ///     plain multi-level lists are unaffected; only explicit lvlRestart values
    ///     change behavior.
    ///   • R == 0   → never restart.
    /// (lvlRestart="1" therefore restarts ONLY when the top level (ilvl=0) ticks,
    /// NOT when an intermediate level ticks — matching real Word.)
    /// </summary>
    private bool ShouldRestartDeeperLevel(int numId, int ilvl, int deeperIlvl)
    {
        var restart = GetEffectiveLvlRestart(numId, deeperIlvl);
        if (restart == 0) return false;          // val="0": never restart
        var r = restart ?? deeperIlvl;           // default: restart on any strictly shallower level
        return (ilvl + 1) <= r;
    }

    /// <summary>
    /// lvlRestart for a level, honoring the embedded-&lt;w:lvl&gt;-override MERGE
    /// model: an embedded override redefines the properties it specifies but
    /// INHERITS those it omits from the base abstractNum level. GetLevel returns
    /// the override lvl (which may omit lvlRestart), so fall back to the base
    /// abstractNum level's lvlRestart when the override does not set one — Word
    /// keeps the original restart rule (e.g. lvlRestart="1") rather than reverting
    /// to the default "restart on any parent tick".
    /// </summary>
    private int? GetEffectiveLvlRestart(int numId, int ilvl)
    {
        var r = GetLevel(numId, ilvl)?.GetFirstChild<LevelRestart>()?.Val?.Value;
        if (r.HasValue) return r;
        return GetAbstractLevel(numId, ilvl)?.GetFirstChild<LevelRestart>()?.Val?.Value;
    }

    /// <summary>
    /// The level definition straight from the abstractNum, IGNORING any
    /// per-instance lvlOverride (the opposite of <see cref="GetLevel"/>, which
    /// prefers an embedded override lvl). Used to recover properties the override
    /// omits under the merge model.
    /// </summary>
    private Level? GetAbstractLevel(int numId, int ilvl)
    {
        var numbering = _doc.MainDocumentPart?.NumberingDefinitionsPart?.Numbering;
        var inst = numbering?.Elements<NumberingInstance>()
            .FirstOrDefault(n => n.NumberID?.Value == numId);
        var absId = inst?.AbstractNumId?.Val?.Value;
        if (absId == null) return null;
        var abs = FindAbstractNum(numbering, absId);
        return abs?.Elements<Level>().FirstOrDefault(l => l.LevelIndex?.Value == ilvl);
    }

    /// <summary>
    /// Expand a level's lvlText template ("%1、", "%1.%2.") against the current
    /// multi-level counters, rendering each %N through WordNumFmtRenderer so
    /// every numFmt (decimal, chineseCounting, decimalEnclosedCircle, …) maps to
    /// the same glyphs in both views. An empty template defaults to "%N".
    /// </summary>
    private string RenderOrderedMarker(OrderedListNumberingState st, int numId, int ilvl, string? lvlText)
    {
        // Distinguish "no lvlText element" (null → legit %N fallback) from a
        // present-but-empty <w:lvlText w:val=""/> (intentional no-marker; Word
        // renders nothing). Only null defaults to the %N template; an explicit
        // "" yields an empty marker. GetLevelText returns null when the level
        // or its lvlText is absent and "" when val="" is present, so the
        // distinction is clean at this layer.
        var template = lvlText ?? $"%{ilvl + 1}";
        // isLgl (ECMA-376 §17.9.10 w:isLgl) on the CURRENT level forces every
        // placeholder in this level's lvlText to render as decimal (arabic),
        // regardless of each referenced level's own numFmt — e.g. an
        // upperRoman level 0 referenced as "%1" inside an isLgl level renders
        // "1" not "I". Read off the current (ilvl) level definition.
        // EXCEPTION: decimalZero is already a decimal format; isLgl only changes
        // the number SYSTEM, so Word preserves decimalZero's zero-padding ("05"
        // stays "05", not "5"). Only non-decimal systems collapse to plain
        // decimal.
        var legal = GetLevel(numId, ilvl)?.GetFirstChild<IsLegalNumberingStyle>();
        var isLgl = legal != null &&
            (legal.Val == null || legal.Val.Value); // CT_OnOff default-true
        // Only %1..%9 are valid Word level placeholders (CT_LevelText / lvlText
        // references 1-based level indices 1-9). %0 is not a placeholder — Word
        // leaves it literal — and matching it here produced k=-1 →
        // GetNumberingFormat(numId,-1) → no level → "bullet" → a • glyph
        // emitted inside an ordered marker. Restrict to %1-%9 so any other %x
        // stays literal in both views.
        return System.Text.RegularExpressions.Regex.Replace(template, @"%([1-9])", m =>
        {
            var k = int.Parse(m.Groups[1].Value) - 1;
            var actualFmt = GetNumberingFormat(numId, k);
            var lvlFmt = isLgl
                ? (actualFmt.Equals("decimalZero", StringComparison.OrdinalIgnoreCase)
                    ? "decimalZero" : "decimal")
                : actualFmt;
            // A placeholder referencing an UNDEFINED level (or one whose fmt
            // resolves to "bullet") must emit nothing for an ordered marker —
            // Word renders an unstarted/undefined level as empty, NOT a bullet
            // glyph. Without this, lvlText "%1.%2.%3." with only level 0 defined
            // produced "1.•.•." (bullet pollution). isLgl forces decimal so it
            // never hits this branch.
            if (!isLgl && lvlFmt.Equals("bullet", StringComparison.OrdinalIgnoreCase))
                return "";
            var counter = st.MultiLevelCounters.GetValueOrDefault(k, 0);
            return OfficeCli.Core.WordNumFmtRenderer.Render(counter, lvlFmt);
        });
    }

    private string GetListPrefix(Paragraph para, OrderedListNumberingState? state = null)
    {
        // A direct <w:numPr><w:numId val="0"/> explicitly suppresses numbering
        // (e.g. a heading that opts out of its style's list). Honor it before
        // resolving the style chain so suppressed paragraphs stay marker-less.
        if (IsNumberingSuppressed(para)) return "";

        // Direct numPr wins; otherwise fall back to the paragraph/character
        // style chain so STYLE-inherited numbering (custom list styles AND
        // Heading1..9) gets a text marker — matching the HTML preview, which
        // already resolves via ResolveNumPrFromStyle. Reading only direct
        // numPr left these paragraphs with no marker in `view text` while
        // `view html` rendered "1.", "1.1.", etc.
        var numProps = para.ParagraphProperties?.NumberingProperties;
        int numIdVal;
        int ilvl;
        var directNumId = numProps?.NumberingId?.Val?.Value;
        if (directNumId != null && directNumId != 0)
        {
            numIdVal = directNumId.Value;
            ilvl = numProps!.NumberingLevelReference?.Val?.Value ?? 0;
        }
        else
        {
            var resolved = ResolveNumPrFromStyle(para);
            if (resolved == null) return "";
            (numIdVal, ilvl) = resolved.Value;
            if (numIdVal == 0) return "";
        }
        int? numId = numIdVal;

        // Clamp ilvl to the OOXML-legal range [0, 8] ONCE before any use.
        // Malformed docs (raw-zip fuzz) carry ilvl negative (→ new string(' ',
        // neg) throws) or huge (~1e9 → ilvl*2 int overflow → negative → throws;
        // 10000 → a 20000-space line). Matches the HTML preview path
        // (WordHandler.HtmlPreview.cs, "ilvl > 8") so text == html marker.
        if (ilvl < 0) ilvl = 0;
        else if (ilvl > 8) ilvl = 8;

        var indent = new string(' ', ilvl * 2);
        var numFmt = GetNumberingFormat(numId.Value, ilvl);

        // Bullet lists render their lvlText glyph in text mode — a custom glyph
        // (★ ▶ ● …) passes through verbatim to match Word and the HTML preview;
        // standard/Wingdings bullets map to •/◦/▪. See BulletGlyphForText.
        if (numFmt.Equals("bullet", StringComparison.OrdinalIgnoreCase))
            return $"{indent}{BulletGlyphForText(GetLevelText(numId.Value, ilvl))} ";

        // Stateless fallback: no walk context → render from the level start.
        var st = state ?? new OrderedListNumberingState();

        // A different numId breaks the active run: clear the in-run counters
        // (but NOT AbsNumLevelCounters — that carries continuation across
        // sibling nums sharing an abstractNum). Mirrors the HTML reset at
        // WordHandler.HtmlPreview.cs. Same numId continues across intervening
        // non-list paragraphs, exactly as Word does.
        if (st.CurrentNumId != numId.Value)
        {
            st.OlCountPerLevel.Clear();
            st.MultiLevelCounters.Clear();
            st.CurrentNumId = numId.Value;
            // Re-arm per-instance startOverrides for the new num: each overridden
            // level fires its override the next time it is DIRECTLY advanced,
            // even if a deeper item carried the level over first. EXCEPTION: a
            // level with lvlRestart="0" (never restart) that is already running
            // does NOT fire — startOverride applies only on a level's initial
            // start or an lvlRestart-triggered restart (ECMA-376 §17.9.7), and a
            // never-restart continued level meets neither, so it keeps counting
            // (Word continues such a level across the numId switch instead of
            // jumping to the override value).
            st.PendingInstanceOverride.Clear();
            for (int lv = 0; lv <= 8; lv++)
                if (GetNumInstanceOverrideStart(numId.Value, lv).HasValue
                    && GetEffectiveLvlRestart(numId.Value, lv) != 0)
                    st.PendingInstanceOverride.Add(lv);
        }

        var absId = GetAbstractNumId(numId.Value);
        AdvanceOrderedCounter(st, numId.Value, absId, ilvl);
        var marker = RenderOrderedMarker(st, numId.Value, ilvl, GetLevelText(numId.Value, ilvl));

        // Separator after the marker: honor <w:suff> (nothing → none; tab/space
        // both collapse to a single space in plain text).
        var sep = GetLevelSuffix(numId.Value, ilvl) == "nothing" ? "" : " ";
        return $"{indent}{marker}{sep}";
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
        var abstractNum = FindAbstractNum(numbering, abstractNumId);
        var level = abstractNum?.Elements<Level>()
            .FirstOrDefault(l => l.LevelIndex?.Value == ilvl);

        // <w:lvlPicBulletId w:val="N"/> is a CHILD element of <w:lvl>, not an
        // attribute — the old GetAttributes() guard here never matched and made
        // this method always return null (picture bullets fell back to a plain
        // circle in the HTML preview). Read the child element directly.
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
    // Returns the level's explicit lvlJc value, or null when the level
    // does not specify one. Callers that need a default substitute their
    // own (the HTML marker uses "right" so a hanging-indent number hugs the
    // text like the bullet ::marker, rather than floating to the far-left of
    // its hanging box).
    private string? GetLevelJustification(int numId, int ilvl)
    {
        var level = GetLevel(numId, ilvl);
        var jc = level?.LevelJustification?.Val;
        if (jc?.HasValue != true) return null;
        return jc.InnerText;
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
        var abstractNum = FindAbstractNum(numbering, abstractNumId);
        return abstractNum?.Elements<Level>()
            .FirstOrDefault(l => l.LevelIndex?.Value == ilvl);
    }

    // Resolve an abstractNumId to its AbstractNum, following a <w:numStyleLink>
    // indirection. An abstractNum that only links to a numbering style
    // (numStyleLink="X") has no level definitions of its own — it borrows them
    // from the abstractNum carrying the matching <w:styleLink w:val="X"/>. Word
    // resolves list-gallery "list styles" this way; without following the link
    // every marker for such a list fell back to a bullet glyph.
    private static AbstractNum? FindAbstractNum(Numbering? numbering, int? abstractNumId)
    {
        if (numbering == null || abstractNumId == null) return null;
        var abs = numbering.Elements<AbstractNum>()
            .FirstOrDefault(a => a.AbstractNumberId?.Value == abstractNumId);
        if (abs == null) return null;
        var link = abs.GetFirstChild<NumberingStyleLink>()?.Val?.Value;
        if (string.IsNullOrEmpty(link)) return abs;
        return numbering.Elements<AbstractNum>()
            .FirstOrDefault(a => a.GetFirstChild<StyleLink>()?.Val?.Value == link) ?? abs;
    }

    private int? GetStartValue(int numId, int ilvl)
    {
        var numbering = _doc.MainDocumentPart?.NumberingDefinitionsPart?.Numbering;
        if (numbering == null) return null;

        var numInstance = numbering.Elements<NumberingInstance>()
            .FirstOrDefault(n => n.NumberID?.Value == numId);
        if (numInstance == null) return null;

        // Check level override first. ECMA-376 §17.9.7: when the lvlOverride
        // embeds a full <w:lvl>, that replacement's own <w:start> governs and
        // the sibling startOverride is ignored — mirror GetLevel's precedence.
        var lvlOverride = numInstance.Elements<LevelOverride>()
            .FirstOrDefault(o => o.LevelIndex?.Value == ilvl);
        if (lvlOverride?.GetFirstChild<Level>() is Level embeddedLevel)
            return embeddedLevel.StartNumberingValue?.Val?.Value;
        if (lvlOverride?.StartOverrideNumberingValue?.Val?.Value is int overrideStart)
            return overrideStart;

        var abstractNumId = numInstance.AbstractNumId?.Val?.Value;
        if (abstractNumId == null) return null;

        var abstractNum = FindAbstractNum(numbering, abstractNumId);
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
    private int? FindContinuationNumId(bool isBullet, Paragraph? targetPara = null, OpenXmlElement? containerHint = null, int targetLevel = 0)
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

        // Continue from the preceding list paragraph at the SAME nesting level
        // (the documented "immediately preceding list" rule, extended to survive
        // a nested sub-list of a different type). When the target is already in
        // the tree (Set path) the predecessors are its earlier siblings; when it
        // is still detached (Add path, inserted only after ApplyListStyle runs)
        // every container element is a predecessor.
        //
        // Walking strictly the immediately preceding paragraph broke ordered
        // lists interrupted by a deeper bullet sub-list: after
        //   1. item1 / (bullet) nested a / 2. item2
        // the predecessor of "item2" (ilvl 0) is the bullet "nested a" (ilvl 1),
        // whose type doesn't match, so a fresh numId was minted and Word
        // restarted numbering at 1. We instead skip deeper-level paragraphs
        // (ilvl > target) and compare against the nearest paragraph at the
        // target level or shallower. A non-list paragraph or a same/shallower-
        // level paragraph of a different list type terminates the scan (no
        // continuation). Only paragraphs are considered — a trailing body-level
        // sectPr (Add path, target still detached before the sectPr) or other
        // non-paragraph siblings are skipped, matching the previous
        // OfType<Paragraph> predecessor selection.
        // Walk predecessors backward via sibling links (NOT LINQ .Reverse(),
        // which buffers every preceding element first — O(n) per call, so an
        // n-item list minting one numId per item was O(n²): ~35s for 8000
        // bullets. Backward sibling walk exits on the first same/shallower-level
        // paragraph, i.e. O(1) for a flat list and O(depth) for a nested one).
        // Set path (attached): start at the previous sibling; Add path
        // (detached, appended later): start at the container's last child.
        OpenXmlElement? cur = targetPara?.Parent != null
            ? targetPara.PreviousSibling()
            : container.LastChild;

        for (; cur != null; cur = cur.PreviousSibling())
        {
            if (cur is not Paragraph prevPara || ReferenceEquals(prevPara, targetPara))
                continue; // skip tables / sectPr / the target itself

            var numProps = prevPara.ParagraphProperties?.NumberingProperties;
            var prevNumId = numProps?.NumberingId?.Val?.Value;
            if (prevNumId == null || prevNumId == 0)
                return null; // non-list paragraph terminates the run

            var prevLevel = numProps?.NumberingLevelReference?.Val?.Value ?? 0;
            if (prevLevel > targetLevel)
                continue; // deeper nested item — skip and keep looking

            var fmt = GetNumberingFormat(prevNumId.Value, 0);
            var prevIsBullet = fmt.ToLowerInvariant() == "bullet";
            return prevIsBullet == isBullet ? prevNumId.Value : (int?)null;
        }

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
        var targetIlvl = listLevel ?? para.ParagraphProperties?.NumberingProperties?.NumberingLevelReference?.Val?.Value ?? 0;
        var continuationNumId = FindContinuationNumId(isBullet, para, containerHint, targetIlvl);
        if (continuationNumId != null && startValue == null)
        {
            var pProps = para.ParagraphProperties ?? para.PrependChild(new ParagraphProperties());
            var ilvl = targetIlvl;
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
            // Apply the caller's start override to the level the paragraph
            // actually binds to (its ilvl), not hardcoded level 0 — a list bound
            // to level 1+ (e.g. a nested ordered markdown list, or a plain
            // `--prop level=1 start=5`) otherwise silently lost its start.
            level.AppendChild(new StartNumberingValue { Val = (lvl == (listLevel ?? 0) && startValue.HasValue) ? startValue.Value : 1 });

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

        // Create numbering instance. CT_Numbering schema order is
        // numPicBullet*, abstractNum*, num*, numIdMacAtCleanup? — so <w:num>
        // must sit AFTER every abstractNum/num but BEFORE a trailing
        // <w:numIdMacAtCleanup>. Appending unconditionally placed the new num
        // after numIdMacAtCleanup when the source carried one, producing a
        // schema-invalid "unexpected w:num child" on validate/open. Insert
        // before numIdMacAtCleanup when present; else append.
        var numInstance = new NumberingInstance { NumberID = maxNumId };
        numInstance.AppendChild(new AbstractNumId { Val = maxAbstractId });
        var macCleanup = numbering.GetFirstChild<NumberingIdMacAtCleanup>();
        if (macCleanup != null)
            numbering.InsertBefore(numInstance, macCleanup);
        else
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
