// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler
{
    // ==================== Style Inheritance (PPT) ====================
    //
    // CONSISTENCY(effective-X-mirror): see WordHandler.StyleList.cs.
    // Mirrors the docx PopulateEffectiveRunProperties / EmitEffectiveRunProperties
    // pattern: single core resolver walks the 8-layer cascade collecting per-key
    // provenance, then a shared emitter writes effective.* + .src onto the
    // DocumentNode, applying the two docx-aligned suppression rules:
    //   1. direct key on node → do NOT emit effective.X
    //   2. source layer is "/direct" → do NOT emit effective.X.src
    //
    // Cascade order (high → low priority, applied in reverse so high wins):
    //   /theme/{majorFont|minorFont|clrScheme/accent1}     (lowest)
    //   /presentation/defaultTextStyle                     (8)
    //   /master[N]/{titleStyle|bodyStyle|otherStyle}       (7)
    //   /master[N]/ph[@type=...][@idx=...]/lvlNpPr         (6)
    //   /slide[N]/layout/ph[@type=...][@idx=...]/lvlNpPr   (5)
    //   /slide[N]/shape[K]/lstStyle/lvlNpPr                (4)
    //   /direct (= para.pPr.defRPr or run.rPr / pPr itself) (highest)

    /// <summary>
    /// Per-property cascade result: value (typed) + source path of the
    /// writing layer. Slot-aware for fonts (latin/ea/cs each independently
    /// sourced, matching docx per-slot rFonts semantics).
    /// </summary>
    private sealed class ResolvedEffective
    {
        public int? Size;
        public string? Color; // hex or scheme token
        public bool? Bold;
        public bool? Italic;
        public string? Underline;
        public string? Strike;
        public string? FontLatin;
        public string? FontEa;
        public string? FontCs;

        public string? Align;
        public string? LineSpacing;
        public string? SpaceBefore;
        public string? SpaceAfter;

        public readonly Dictionary<string, string> Sources = new();
    }

    /// <summary>
    /// Resolve effective run + paragraph properties for a (shape, level) pair.
    /// `directRpr` and `directPpr` are the highest-priority layer (the run's
    /// rPr / paragraph's pPr); passing nulls means "shape-level resolution"
    /// (no run/paragraph direct layer). Layered low → high so later writes
    /// win, mirroring WordHandler.ResolveEffectiveRunPropertiesCore.
    /// </summary>
    private static ResolvedEffective ResolveEffectiveDefRpCore(
        Shape shape,
        SlidePart slidePart,
        int level,
        Drawing.RunProperties? directRpr,
        Drawing.ParagraphProperties? directPpr)
    {
        var r = new ResolvedEffective();

        var ph = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
            ?.GetFirstChild<PlaceholderShape>();
        var phType = ph?.Type?.HasValue == true ? ph.Type.Value : PlaceholderValues.Body;
        bool isTitle = phType == PlaceholderValues.Title || phType == PlaceholderValues.CenteredTitle;
        bool isSubTitle = phType == PlaceholderValues.SubTitle;

        // Layer 1 (lowest): theme — major font for titles, minor for body.
        var theme = slidePart.SlideLayoutPart?.SlideMasterPart?.ThemePart?.Theme;
        var fontScheme = theme?.ThemeElements?.FontScheme;
        if (fontScheme != null)
        {
            OpenXmlCompositeElement? themeFont = isTitle ? fontScheme.MajorFont : (OpenXmlCompositeElement?)fontScheme.MinorFont;
            if (themeFont != null)
            {
                var tLatin = themeFont.GetFirstChild<Drawing.LatinFont>()?.Typeface?.Value;
                if (tLatin != null && !tLatin.StartsWith("+", StringComparison.Ordinal))
                {
                    r.FontLatin = tLatin;
                    r.Sources["font.latin"] = isTitle ? "/theme/majorFont" : "/theme/minorFont";
                }
                var tEa = themeFont.GetFirstChild<Drawing.EastAsianFont>()?.Typeface?.Value;
                if (!string.IsNullOrEmpty(tEa) && !tEa.StartsWith("+", StringComparison.Ordinal))
                {
                    r.FontEa = tEa;
                    r.Sources["font.ea"] = isTitle ? "/theme/majorFont" : "/theme/minorFont";
                }
                var tCs = themeFont.GetFirstChild<Drawing.ComplexScriptFont>()?.Typeface?.Value;
                if (!string.IsNullOrEmpty(tCs) && !tCs.StartsWith("+", StringComparison.Ordinal))
                {
                    r.FontCs = tCs;
                    r.Sources["font.cs"] = isTitle ? "/theme/majorFont" : "/theme/minorFont";
                }
            }
        }

        // Layer 2: presentation defaultTextStyle
        var presStyle = GetPresentationDefaultTextStyle(slidePart);
        ApplyLevelPpr(r, presStyle, level, "/presentation/defaultTextStyle");

        // Layer 3: master txStyles (title/body/other)
        var masterIdx = GetMasterIndex(slidePart);
        var masterTxStyles = slidePart.SlideLayoutPart?.SlideMasterPart?.SlideMaster?.TextStyles;
        if (masterTxStyles != null)
        {
            OpenXmlCompositeElement? styleList;
            string styleLabel;
            if (isTitle) { styleList = masterTxStyles.TitleStyle; styleLabel = "titleStyle"; }
            else if (isSubTitle || phType == PlaceholderValues.Body || phType == PlaceholderValues.Object)
            { styleList = masterTxStyles.BodyStyle; styleLabel = "bodyStyle"; }
            else { styleList = masterTxStyles.OtherStyle; styleLabel = "otherStyle"; }
            if (styleList != null)
                ApplyLevelPpr(r, styleList, level, $"/master[{masterIdx}]/{styleLabel}/lvl{level + 1}pPr");
        }

        // Layer 4: master placeholder (matching ph type/idx) lstStyle
        if (ph != null)
        {
            var masterTree = slidePart.SlideLayoutPart?.SlideMasterPart?.SlideMaster?.CommonSlideData?.ShapeTree;
            ApplyPhTreeLayer(r, masterTree, ph, level, $"/master[{masterIdx}]");
        }

        // Layer 5: layout placeholder lstStyle
        if (ph != null)
        {
            var slideIdx = GetSlideIndex(slidePart);
            var layoutTree = slidePart.SlideLayoutPart?.SlideLayout?.CommonSlideData?.ShapeTree;
            ApplyPhTreeLayer(r, layoutTree, ph, level, $"/slide[{slideIdx}]/layout");
        }

        // Layer 6: shape's own lstStyle/lvlNpPr defRPr
        var lstStyle = shape.TextBody?.GetFirstChild<Drawing.ListStyle>();
        ApplyLevelPpr(r, lstStyle, level, "/shape/lstStyle");

        // Layer 7 (highest): paragraph's own pPr (defRPr + pPr attrs) and run rPr
        if (directPpr != null)
        {
            ApplyParaPprDirect(r, directPpr);
            // The paragraph's defRPr is also part of "direct" for runs in that para.
            var defRp = directPpr.GetFirstChild<Drawing.DefaultRunProperties>();
            if (defRp != null)
                ApplyDefRp(r, defRp, "/direct");
        }
        if (directRpr != null)
            ApplyRunRpr(r, directRpr);

        return r;
    }

    private static void ApplyPhTreeLayer(
        ResolvedEffective r,
        OpenXmlCompositeElement? tree,
        PlaceholderShape ph,
        int level,
        string layerRoot)
    {
        if (tree == null) return;
        foreach (var candidate in tree.Elements<Shape>())
        {
            var cPh = candidate.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
                ?.GetFirstChild<PlaceholderShape>();
            if (cPh == null || !PlaceholderMatches(ph, cPh)) continue;

            var typeAttr = cPh.Type?.HasValue == true ? cPh.Type.InnerText : null;
            var idxAttr = cPh.Index?.HasValue == true ? cPh.Index.Value.ToString() : null;
            var phPath = BuildPhPath(layerRoot, typeAttr, idxAttr, level);

            var cLstStyle = candidate.TextBody?.GetFirstChild<Drawing.ListStyle>();
            ApplyLevelPpr(r, cLstStyle, level, phPath);
            break;
        }
    }

    /// <summary>
    /// Builds a placeholder source path of the form
    /// "/<root>/ph[@type=body][@idx=1]/lvl1pPr". Both type and idx are
    /// chain-bracketed with `@` per project convention.
    /// </summary>
    private static string BuildPhPath(string root, string? type, string? idx, int level)
    {
        var sb = new System.Text.StringBuilder(root);
        sb.Append("/ph");
        if (type != null) sb.Append("[@type=").Append(type).Append(']');
        if (idx != null) sb.Append("[@idx=").Append(idx).Append(']');
        sb.Append("/lvl").Append(level + 1).Append("pPr");
        return sb.ToString();
    }

    /// <summary>
    /// Apply a level-N pPr block (extracted from a styleList) to the
    /// resolved cascade. styleList may be a master textStyles container
    /// (titleStyle/bodyStyle/otherStyle), a shape lstStyle, or the
    /// presentation defaultTextStyle — they all carry lvlNpPr children.
    /// </summary>
    private static void ApplyLevelPpr(
        ResolvedEffective r,
        OpenXmlCompositeElement? styleList,
        int level,
        string layer)
    {
        if (styleList == null) return;
        OpenXmlElement? lvlPpr = level switch
        {
            0 => styleList.GetFirstChild<Drawing.Level1ParagraphProperties>(),
            1 => styleList.GetFirstChild<Drawing.Level2ParagraphProperties>(),
            2 => styleList.GetFirstChild<Drawing.Level3ParagraphProperties>(),
            3 => styleList.GetFirstChild<Drawing.Level4ParagraphProperties>(),
            4 => styleList.GetFirstChild<Drawing.Level5ParagraphProperties>(),
            5 => styleList.GetFirstChild<Drawing.Level6ParagraphProperties>(),
            6 => styleList.GetFirstChild<Drawing.Level7ParagraphProperties>(),
            7 => styleList.GetFirstChild<Drawing.Level8ParagraphProperties>(),
            8 => styleList.GetFirstChild<Drawing.Level9ParagraphProperties>(),
            _ => styleList.GetFirstChild<Drawing.Level1ParagraphProperties>(),
        };
        if (lvlPpr == null) return;

        ApplyPprAttrs(r, lvlPpr, layer);

        var defRp = lvlPpr.GetFirstChild<Drawing.DefaultRunProperties>();
        if (defRp != null)
            ApplyDefRp(r, defRp, layer);
    }

    /// <summary>
    /// Read paragraph-level attributes (align) and child line/space elements
    /// off a pPr-shaped element. Used for both lvlNpPr layers and the
    /// paragraph's own direct pPr.
    /// </summary>
    private static void ApplyPprAttrs(ResolvedEffective r, OpenXmlElement pPr, string layer)
    {
        // alignment lives as an attribute on lvlNpPr / pPr ("algn")
        var algnAttr = pPr.GetAttributes().FirstOrDefault(a => a.LocalName == "algn");
        if (algnAttr.Value != null)
        {
            r.Align = algnAttr.Value switch
            {
                "l" => "left",
                "ctr" => "center",
                "r" => "right",
                "just" => "justify",
                _ => algnAttr.Value
            };
            r.Sources["align"] = layer;
        }

        var lineSp = pPr.GetFirstChild<Drawing.LineSpacing>();
        if (lineSp != null)
        {
            var pct = lineSp.GetFirstChild<Drawing.SpacingPercent>()?.Val?.Value;
            if (pct.HasValue)
            {
                r.LineSpacing = SpacingConverter.FormatPptLineSpacingPercent(pct.Value);
                r.Sources["lineSpacing"] = layer;
            }
            else
            {
                var pts = lineSp.GetFirstChild<Drawing.SpacingPoints>()?.Val?.Value;
                if (pts.HasValue)
                {
                    r.LineSpacing = SpacingConverter.FormatPptLineSpacingPoints(pts.Value);
                    r.Sources["lineSpacing"] = layer;
                }
            }
        }
        var sb = pPr.GetFirstChild<Drawing.SpaceBefore>()?.GetFirstChild<Drawing.SpacingPoints>()?.Val?.Value;
        if (sb.HasValue)
        {
            r.SpaceBefore = SpacingConverter.FormatPptSpacing(sb.Value);
            r.Sources["spaceBefore"] = layer;
        }
        var sa = pPr.GetFirstChild<Drawing.SpaceAfter>()?.GetFirstChild<Drawing.SpacingPoints>()?.Val?.Value;
        if (sa.HasValue)
        {
            r.SpaceAfter = SpacingConverter.FormatPptSpacing(sa.Value);
            r.Sources["spaceAfter"] = layer;
        }
    }

    private static void ApplyParaPprDirect(ResolvedEffective r, Drawing.ParagraphProperties pPr)
        => ApplyPprAttrs(r, pPr, "/direct");

    /// <summary>Apply a defRPr layer (style cascade).</summary>
    private static void ApplyDefRp(ResolvedEffective r, Drawing.DefaultRunProperties defRp, string layer)
    {
        ApplyRprLike(r, defRp, layer);
    }

    /// <summary>Apply a run's direct rPr (highest priority).</summary>
    private static void ApplyRunRpr(ResolvedEffective r, Drawing.RunProperties rPr)
    {
        ApplyRprLike(r, rPr, "/direct");
    }

    /// <summary>
    /// Shared rPr/defRPr property extraction. Both share the same OOXML
    /// shape (CT_TextCharacterProperties): size attr (sz), bold (b), italic
    /// (i), underline (u), strike, latin/ea/cs children, solidFill child.
    /// </summary>
    private static void ApplyRprLike(ResolvedEffective r, OpenXmlElement rprLike, string layer)
    {
        foreach (var attr in rprLike.GetAttributes())
        {
            switch (attr.LocalName)
            {
                case "sz":
                    if (int.TryParse(attr.Value, out var sz)) { r.Size = sz; r.Sources["size"] = layer; }
                    break;
                case "b":
                    r.Bold = attr.Value == "1" || attr.Value?.Equals("true", StringComparison.OrdinalIgnoreCase) == true;
                    r.Sources["bold"] = layer;
                    break;
                case "i":
                    r.Italic = attr.Value == "1" || attr.Value?.Equals("true", StringComparison.OrdinalIgnoreCase) == true;
                    r.Sources["italic"] = layer;
                    break;
                case "u":
                    if (attr.Value != null && attr.Value != "none")
                    {
                        r.Underline = attr.Value switch { "sng" => "single", "dbl" => "double", _ => attr.Value };
                        r.Sources["underline"] = layer;
                    }
                    break;
                case "strike":
                    if (attr.Value != null)
                    {
                        r.Strike = attr.Value switch
                        {
                            "dblStrike" => "double",
                            "noStrike" => "none",
                            _ => "single",
                        };
                        r.Sources["strike"] = layer;
                    }
                    break;
            }
        }

        var latin = rprLike.GetFirstChild<Drawing.LatinFont>()?.Typeface?.Value;
        if (!string.IsNullOrEmpty(latin) && !latin.StartsWith("+", StringComparison.Ordinal))
        {
            r.FontLatin = latin;
            r.Sources["font.latin"] = layer;
        }
        var ea = rprLike.GetFirstChild<Drawing.EastAsianFont>()?.Typeface?.Value;
        if (!string.IsNullOrEmpty(ea) && !ea.StartsWith("+", StringComparison.Ordinal))
        {
            r.FontEa = ea;
            r.Sources["font.ea"] = layer;
        }
        var cs = rprLike.GetFirstChild<Drawing.ComplexScriptFont>()?.Typeface?.Value;
        if (!string.IsNullOrEmpty(cs) && !cs.StartsWith("+", StringComparison.Ordinal))
        {
            r.FontCs = cs;
            r.Sources["font.cs"] = layer;
        }

        var fill = rprLike.GetFirstChild<Drawing.SolidFill>();
        var color = ReadColorFromFill(fill);
        if (color != null)
        {
            r.Color = color;
            r.Sources["color"] = layer;
        }
    }

    // ==================== Emit (mirror docx) ====================

    /// <summary>
    /// Populates effective.* keys on a shape or run node. Mirrors docx
    /// EmitEffectiveRunProperties: suppress effective.X when direct X exists;
    /// suppress .src when source is /direct.
    /// </summary>
    private static void EmitEffectiveFromResolved(DocumentNode node, ResolvedEffective r)
    {
        void EmitSrc(string effectiveKey, string sourceKey)
        {
            if (r.Sources.TryGetValue(sourceKey, out var src) && src != "/direct")
                node.Format[effectiveKey + ".src"] = src;
        }

        // size
        if (!node.Format.ContainsKey("size") && r.Size.HasValue)
        {
            node.Format["effective.size"] = $"{r.Size.Value / 100.0:0.##}pt";
            EmitSrc("effective.size", "size");
        }

        // Per-slot font — each slot independently honors the cascade and is
        // suppressed only when that specific slot or bare `font` is set.
        // CONSISTENCY(effective-X-mirror): see WordHandler.StyleList.cs
        // EmitEffectiveRunProperties per-slot fonts.
        bool hasBareFont = node.Format.ContainsKey("font");
        if (!hasBareFont && !node.Format.ContainsKey("font.latin") && r.FontLatin != null)
        {
            node.Format["effective.font.latin"] = r.FontLatin;
            EmitSrc("effective.font.latin", "font.latin");
        }
        if (!hasBareFont && !node.Format.ContainsKey("font.ea") && r.FontEa != null)
        {
            node.Format["effective.font.ea"] = r.FontEa;
            EmitSrc("effective.font.ea", "font.ea");
        }
        if (!hasBareFont && !node.Format.ContainsKey("font.cs") && r.FontCs != null)
        {
            node.Format["effective.font.cs"] = r.FontCs;
            EmitSrc("effective.font.cs", "font.cs");
        }
        // Bare effective.font convenience: emitted only when none of the
        // per-slot effective.font.* keys collide with direct slot keys, and
        // a Latin slot resolved (matches the docx-era bare-font readback
        // contract for round-trip with PptxTextboxR9Tests).
        if (!hasBareFont && r.FontLatin != null)
        {
            node.Format["effective.font"] = r.FontLatin;
            EmitSrc("effective.font", "font.latin");
        }

        if (!node.Format.ContainsKey("bold") && r.Bold == true)
        {
            node.Format["effective.bold"] = true;
            EmitSrc("effective.bold", "bold");
        }
        if (!node.Format.ContainsKey("italic") && r.Italic == true)
        {
            node.Format["effective.italic"] = true;
            EmitSrc("effective.italic", "italic");
        }
        if (!node.Format.ContainsKey("underline") && r.Underline != null)
        {
            node.Format["effective.underline"] = r.Underline;
            EmitSrc("effective.underline", "underline");
        }
        if (!node.Format.ContainsKey("strike") && r.Strike != null)
        {
            node.Format["effective.strike"] = r.Strike;
            EmitSrc("effective.strike", "strike");
        }
        if (!node.Format.ContainsKey("color") && r.Color != null)
        {
            node.Format["effective.color"] = r.Color;
            EmitSrc("effective.color", "color");
        }
    }

    private static void EmitEffectiveParagraphProperties(DocumentNode node, ResolvedEffective r)
    {
        void EmitSrc(string effectiveKey, string sourceKey)
        {
            if (r.Sources.TryGetValue(sourceKey, out var src) && src != "/direct")
                node.Format[effectiveKey + ".src"] = src;
        }

        if (!node.Format.ContainsKey("align") && r.Align != null)
        {
            node.Format["effective.align"] = r.Align;
            EmitSrc("effective.align", "align");
        }
        if (!node.Format.ContainsKey("lineSpacing") && r.LineSpacing != null)
        {
            node.Format["effective.lineSpacing"] = r.LineSpacing;
            EmitSrc("effective.lineSpacing", "lineSpacing");
        }
        if (!node.Format.ContainsKey("spaceBefore") && r.SpaceBefore != null)
        {
            node.Format["effective.spaceBefore"] = r.SpaceBefore;
            EmitSrc("effective.spaceBefore", "spaceBefore");
        }
        if (!node.Format.ContainsKey("spaceAfter") && r.SpaceAfter != null)
        {
            node.Format["effective.spaceAfter"] = r.SpaceAfter;
            EmitSrc("effective.spaceAfter", "spaceAfter");
        }
    }

    // ==================== Top-level entry points (called from NodeBuilder) ====================

    /// <summary>
    /// Shape-level effective.* — resolves run cascade for level 0 (no direct
    /// run/para layer; the shape doesn't have a single direct rPr).
    /// </summary>
    private static void PopulateEffectiveShapeProperties(DocumentNode node, Shape shape, OpenXmlPart? part)
    {
        if (part is not SlidePart slidePart) return;
        var firstPara = shape.TextBody?.Elements<Drawing.Paragraph>().FirstOrDefault();
        int level = firstPara?.ParagraphProperties?.Level?.Value ?? 0;
        var r = ResolveEffectiveDefRpCore(shape, slidePart, level, directRpr: null, directPpr: firstPara?.ParagraphProperties);
        EmitEffectiveFromResolved(node, r);
        EmitEffectiveParagraphProperties(node, r);
    }

    /// <summary>
    /// Run-level effective.* — resolves the cascade with the run's own rPr
    /// as the highest-priority layer plus the enclosing paragraph's pPr.
    /// </summary>
    private static void PopulateEffectiveRunProperties(DocumentNode node, Drawing.Run run, OpenXmlPart? part)
    {
        if (part is not SlidePart slidePart) return;
        var shape = run.Ancestors<Shape>().FirstOrDefault();
        if (shape == null) return;
        var para = run.Ancestors<Drawing.Paragraph>().FirstOrDefault();
        int level = para?.ParagraphProperties?.Level?.Value ?? 0;
        var r = ResolveEffectiveDefRpCore(shape, slidePart, level, run.RunProperties, para?.ParagraphProperties);
        EmitEffectiveFromResolved(node, r);
    }

    /// <summary>
    /// Paragraph-level effective.* — paragraph-only keys (align/lineSpacing/
    /// spaceBefore/spaceAfter) sourced through the same cascade.
    /// </summary>
    private static void PopulateEffectiveParagraphPropertiesPpt(
        DocumentNode node, Shape shape, Drawing.Paragraph para, OpenXmlPart? part)
    {
        if (part is not SlidePart slidePart) return;
        int level = para.ParagraphProperties?.Level?.Value ?? 0;
        var r = ResolveEffectiveDefRpCore(shape, slidePart, level, directRpr: null, directPpr: para.ParagraphProperties);
        EmitEffectiveParagraphProperties(node, r);
    }

    // ==================== Helpers for source path indices ====================

    private static int GetMasterIndex(SlidePart slidePart)
    {
        var masterPart = slidePart.SlideLayoutPart?.SlideMasterPart;
        if (masterPart == null) return 1;
        // Walk PresentationPart's master list to find the 1-based index.
        foreach (var rel in masterPart.GetParentParts())
        {
            if (rel is PresentationPart pp)
            {
                var masters = pp.SlideMasterParts.ToList();
                var idx = masters.FindIndex(m => ReferenceEquals(m, masterPart));
                if (idx >= 0) return idx + 1;
            }
        }
        return 1;
    }

    private static int GetSlideIndex(SlidePart slidePart)
    {
        foreach (var rel in slidePart.GetParentParts())
        {
            if (rel is PresentationPart pp)
            {
                var slides = pp.SlideParts.ToList();
                var idx = slides.FindIndex(s => ReferenceEquals(s, slidePart));
                if (idx >= 0) return idx + 1;
            }
        }
        return 1;
    }
}
