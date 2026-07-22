// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler
{
    // ==================== Text Rendering ====================

    private static void RenderTextBody(StringBuilder sb, OpenXmlElement textBody, Dictionary<string, string> themeColors,
        Shape? placeholderShape = null, OpenXmlPart? placeholderPart = null, string? fontRefDefaultColor = null,
        int? slideNumber = null)
    {
        // Per-textbody auto-number counters, keyed by scheme type + paragraph level.
        // Resets when switching type/level. Paragraphs aren't wrapped in <ol>, so
        // we count manually and emit the numeric glyph inline.
        var autoNumCounters = new Dictionary<string, int>();
        string? lastAutoKey = null;
        int? lastAutoLevel = null;
        // Resolve the theme font that runs in this textbody should inherit when
        // they carry no explicit Latin typeface (or carry the theme reference
        // "+mj-lt" / "+mn-lt"). Title placeholders inherit the major (heading)
        // typeface; everything else inherits minor (body). GetTextDefaults emits
        // only the body font at slide scope, so without this fallback title
        // placeholders silently render in the body face in HTML preview while
        // PowerPoint renders them in the heading face.
        bool isTitle = IsTitlePlaceholder(placeholderShape);
        string? themeFontFallback = ResolveThemeFontTypeface(placeholderPart, isTitle ? "major" : "minor");
        // Bug #8(B): honor <a:normAutofit fontScale="N"> (1/1000% → ratio). A real
        // PowerPoint deck that shrank text to fit stores the computed scale here;
        // multiply run font sizes by it. Absent/100% (CLI-created files) = no scale.
        // textBody is a p:txBody (Presentation.TextBody), and its a:bodyPr is a
        // Drawing.BodyProperties child — read it generically, not via a cast to
        // Drawing.TextBody (which a p:txBody is not).
        var bodyPr = textBody.GetFirstChild<Drawing.BodyProperties>();
        // R35: default tab interval (<a:bodyPr defTabSz="EMU">), used for tabs
        // beyond the last defined stop. Default 914400 EMU (1in = 72pt) per OOXML.
        // SDK 3.x does not surface defTabSz as a typed property on BodyProperties,
        // so read the raw attribute.
        double defTabPt = 72.0;
        var defTabAttr = bodyPr?.GetAttributes().FirstOrDefault(a => a.LocalName == "defTabSz").Value;
        if (long.TryParse(defTabAttr, out var defTabEmu) && defTabEmu > 0)
            defTabPt = Units.EmuToPt(defTabEmu);
        var naf = bodyPr?
            .GetFirstChild<Drawing.NormalAutoFit>();
        var nafScale = naf?.FontScale?.Value;
        double fontScale = (nafScale.HasValue && nafScale.Value > 0) ? nafScale.Value / 100000.0 : 1.0;
        // R10-1: honor <a:normAutofit lnSpcReduction="N"> (1/1000% → ratio). When a
        // PowerPoint deck shrank line spacing to fit, it stores the reduction here;
        // multiply paragraph line-heights by (1 - reduction). Absent/0 = no change.
        var lnRed = naf?.LineSpaceReduction?.Value;
        double lnSpcFactor = (lnRed.HasValue && lnRed.Value > 0) ? (1 - lnRed.Value / 100000.0) : 1.0;
        bool isFirstPara = true;
        // PowerPoint ignores spaceAfter on the LAST paragraph of a text body (the
        // gap would fall outside the body and has no effect on layout/centering).
        var lastParaRef = textBody.Elements<Drawing.Paragraph>().LastOrDefault();
        foreach (var para in textBody.Elements<Drawing.Paragraph>())
        {
            bool isLastPara = ReferenceEquals(para, lastParaRef);
            // Resolve per-paragraph font size based on paragraph level
            int? defaultFontSizeHundredths = null;
            // R7-2: inherited default run color from the placeholder/master cascade.
            string? defaultRunColor = null;
            // R7-3: inherited default line-spacing CSS fragment from the cascade.
            string? inheritedLineSpacing = null;
            // R26: inherited level-paragraph-properties (lvlNpPr) from the
            // placeholder/master cascade. Used to fill bullet (R26-1), paragraph
            // spacing (R26-2), bold/italic (R26-3), and alignment (R26-4) when the
            // slide paragraph itself declares none. The GET path already reports
            // these as effective.* — mirror it in the RENDER path.
            OpenXmlElement? inheritedLvlPpr = null;
            Drawing.DefaultRunProperties? inheritedDefRp = null;
            Drawing.DefaultRunProperties? inheritedCapsRp = null;
            Drawing.DefaultRunProperties? inheritedUnderlineRp = null;
            Drawing.DefaultRunProperties? inheritedStrikeRp = null;
            Drawing.DefaultRunProperties? inheritedSpacingRp = null;
            if (placeholderShape != null && placeholderPart != null)
            {
                int level = para.ParagraphProperties?.Level?.Value ?? 0;
                defaultFontSizeHundredths = ResolvePlaceholderFontSize(placeholderShape, placeholderPart, level);
                defaultRunColor = ResolvePlaceholderDefaultColor(placeholderShape, placeholderPart, themeColors, level);
                inheritedLineSpacing = ResolvePlaceholderLineSpacing(placeholderShape, placeholderPart, level);
                // Any inherited lvlNpPr (for alignment/spacing/bullet); take the
                // first level pPr that carries ANY content in the chain.
                inheritedLvlPpr = ResolvePlaceholderLevelPpr(placeholderShape, placeholderPart, level, _ => true);
                inheritedDefRp = ResolvePlaceholderDefRp(placeholderShape, placeholderPart, level,
                    dr => dr.Bold?.HasValue == true || dr.Italic?.HasValue == true);
                // Caps (cap="all"/"small") is resolved with its OWN predicate, not folded
                // into the bold/italic lookup: many themes apply all-caps to title/body
                // placeholders via the master/layout defRPr with NO bold or italic, so the
                // bold/italic predicate never matched and inherited caps was silently
                // dropped (PowerPoint renders uppercase; the preview rendered mixed case).
                // A dedicated lookup also avoids cross-level contamination (caps on one
                // inheritance level, bold on another).
                inheritedCapsRp = ResolvePlaceholderDefRp(placeholderShape, placeholderPart, level,
                    dr => dr.Capital?.HasValue == true);
                // Underline / strike from a master/layout placeholder defRPr get their
                // OWN dedicated lookups too (same reasoning as caps): a theme may set
                // u="sng"/strike on a placeholder level's defRPr with no bold/italic, so
                // the bold/italic predicate never matched and the decoration was dropped
                // (PowerPoint renders the underline; the preview rendered plain text).
                inheritedUnderlineRp = ResolvePlaceholderDefRp(placeholderShape, placeholderPart, level,
                    dr => dr.Underline?.HasValue == true);
                inheritedStrikeRp = ResolvePlaceholderDefRp(placeholderShape, placeholderPart, level,
                    dr => dr.Strike?.HasValue == true);
                // Character spacing (spc) gets its own dedicated lookup too: a theme may
                // set spc on a placeholder level's defRPr with no bold/italic, so the
                // bold/italic predicate never matched and the inherited tracking was
                // dropped (PowerPoint renders wide tracking; the preview rendered normal).
                inheritedSpacingRp = ResolvePlaceholderDefRp(placeholderShape, placeholderPart, level,
                    dr => dr.Spacing?.HasValue == true);
            }
            // R11-3: style-matrix fontRef schemeClr is the FINAL fallback run color
            // when no explicit run color and no inherited placeholder color is found.
            defaultRunColor ??= fontRefDefaultColor;
            var paraStyles = new List<string>();

            var pProps = para.ParagraphProperties;
            // R26-4: alignment — explicit slide pPr@algn wins; otherwise inherit
            // the master/layout lvlNpPr@algn via the placeholder cascade.
            var algnInner = pProps?.Alignment?.HasValue == true
                ? pProps.Alignment.InnerText
                : (inheritedLvlPpr as Drawing.TextParagraphPropertiesType)?.Alignment?.HasValue == true
                    ? ((Drawing.TextParagraphPropertiesType)inheritedLvlPpr!).Alignment!.InnerText
                    : null;
            if (algnInner != null)
            {
                var align = algnInner switch
                {
                    "l" => "left",
                    "ctr" => "center",
                    "r" => "right",
                    "just" => "justify",
                    "justLow" => "justify",   // Justify Low (Arabic kashida): like just
                    "dist" => "justify",
                    "thaiDist" => "justify",  // Thai Distributed: like dist
                    _ => "left"
                };
                paraStyles.Add($"text-align:{align}");
                // dist / thaiDist stretch EVERY line — including the last — to the
                // full text-box width (inter-word, not inter-character for Latin).
                if (algnInner == "dist" || algnInner == "thaiDist")
                    paraStyles.Add("text-align-last:justify");
            }

            // Paragraph spacing. PowerPoint ignores spcBef on the FIRST paragraph
            // of a text body (the line sits flush at the body's top inset), so we
            // suppress the margin-top contribution for that paragraph only;
            // subsequent paragraphs keep their spaceBefore.
            // Effective font size (hundredths of a point) for percent-based
            // spacing: spcPct expresses the gap as a percentage of the line's
            // font size (300000 = 300%). Mirror the bullet-size resolution
            // chain (first run size > placeholder/default > 18pt body default).
            var spcFontHundredths = para.Elements<Drawing.Run>().FirstOrDefault()?.RunProperties?.FontSize?.Value
                ?? defaultFontSizeHundredths ?? 1800;

            // Issue #236: the paragraph div carries the line-height, so it must
            // also carry a font-size — otherwise the CSS strut resolves against
            // the inherited root size and every line box collapses to the root
            // font size regardless of the runs' actual sizes (6.5pt text got an
            // 18pt line box; 24pt text got a 1.0× box). Size the strut to the
            // tallest run in the paragraph, mirroring PowerPoint's line box.
            int paraMaxFontHundredths = 0;
            foreach (var r0 in para.Elements<Drawing.Run>())
            {
                int sz = r0.RunProperties?.FontSize?.Value ?? defaultFontSizeHundredths ?? 1800;
                if (sz > paraMaxFontHundredths) paraMaxFontHundredths = sz;
            }
            if (paraMaxFontHundredths == 0) paraMaxFontHundredths = defaultFontSizeHundredths ?? 1800;
            paraStyles.Add($"font-size:{paraMaxFontHundredths / 100.0 * fontScale:0.##}pt");
            // PowerPoint single spacing is the font's line pitch (hhea/OS2
            // ascent+descent+lineGap ≈ 1.2× for Latin faces, ~1.32× for CJK
            // faces like Microsoft YaHei), not 1.0× the font size. Resolve the
            // ratio from the paragraph's first run typeface; unresolvable fonts
            // fall back to the 1.2 Latin norm (GetRatio returns 1.0 on failure,
            // which is never a real line pitch).
            var firstRp = para.Elements<Drawing.Run>().FirstOrDefault()?.RunProperties;
            var paraTypeface = firstRp?.GetFirstChild<Drawing.LatinFont>()?.Typeface?.Value
                ?? firstRp?.GetFirstChild<Drawing.EastAsianFont>()?.Typeface?.Value;
            if (paraTypeface != null && paraTypeface.StartsWith("+", StringComparison.Ordinal))
                paraTypeface = ResolveThemeFontToken(placeholderPart, paraTypeface) ?? themeFontFallback;
            paraTypeface ??= themeFontFallback;
            double singleRatio = SingleSpacingPitch(paraTypeface);
            // R26-2: spaceBefore — explicit slide pPr wins; otherwise inherit the
            // master/layout lvlNpPr spcBef via the placeholder cascade.
            var sbElem = pProps?.GetFirstChild<Drawing.SpaceBefore>()
                ?? inheritedLvlPpr?.GetFirstChild<Drawing.SpaceBefore>();
            var sbPts = sbElem?.GetFirstChild<Drawing.SpacingPoints>()?.Val?.Value;
            if (sbPts.HasValue && !isFirstPara) paraStyles.Add($"margin-top:{sbPts.Value / 100.0:0.##}pt");
            else
            {
                // SpacingPercent fallback (parity with lineSpacing below). pct
                // is /100000; multiply by the effective font pt to get the gap.
                var sbPct = sbElem?.GetFirstChild<Drawing.SpacingPercent>().PercentVal();
                if (sbPct.HasValue && !isFirstPara)
                    paraStyles.Add($"margin-top:{sbPct.Value / 100000.0 * (spcFontHundredths / 100.0):0.##}pt");
            }
            var saElem = pProps?.GetFirstChild<Drawing.SpaceAfter>()
                ?? inheritedLvlPpr?.GetFirstChild<Drawing.SpaceAfter>();
            var saPts = saElem?.GetFirstChild<Drawing.SpacingPoints>()?.Val?.Value;
            if (saPts.HasValue && !isLastPara) paraStyles.Add($"margin-bottom:{saPts.Value / 100.0:0.##}pt");
            else if (!isLastPara)
            {
                var saPct = saElem?.GetFirstChild<Drawing.SpacingPercent>().PercentVal();
                if (saPct.HasValue)
                    paraStyles.Add($"margin-bottom:{saPct.Value / 100000.0 * (spcFontHundredths / 100.0):0.##}pt");
            }

            // Line spacing. R10-1: scale percent/default line-heights by the
            // normAutofit lnSpcReduction factor (fixed-pt line-heights are absolute
            // and unaffected, matching PowerPoint).
            // Issue #236: percent spacing (and the single-spacing default) is
            // relative to the font's line pitch, so every unitless multiplier is
            // scaled by singleRatio; fixed-pt spacing stays absolute.
            var lsPct = pProps?.GetFirstChild<Drawing.LineSpacing>()?.GetFirstChild<Drawing.SpacingPercent>().PercentVal();
            if (lsPct.HasValue) paraStyles.Add($"line-height:{lsPct.Value / 100000.0 * lnSpcFactor * singleRatio:0.##}");
            var lsPts = pProps?.GetFirstChild<Drawing.LineSpacing>()?.GetFirstChild<Drawing.SpacingPoints>()?.Val?.Value;
            if (lsPts.HasValue) paraStyles.Add($"line-height:{lsPts.Value / 100.0:0.##}pt");
            // R7-3: no explicit lnSpc on the paragraph — inherit the master/layout
            // bodyStyle lvl lnSpc resolved via the placeholder cascade. The
            // fragment is "line-height:N" (unitless multiplier — scale it) or
            // "line-height:Npt" (fixed — keep absolute).
            else if (!lsPct.HasValue && inheritedLineSpacing != null)
            {
                if (!inheritedLineSpacing.EndsWith("pt", StringComparison.Ordinal)
                    && double.TryParse(inheritedLineSpacing["line-height:".Length..],
                        System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var inhMul))
                    paraStyles.Add($"line-height:{inhMul * singleRatio:0.##}");
                else
                    paraStyles.Add(inheritedLineSpacing);
            }
            // Default single spacing (incl. R10-1 lnSpcReduction): emit the
            // metric-derived multiplier — the bare `.para { line-height: 1 }`
            // stylesheet default understates every line box (issue #236).
            else if (!lsPct.HasValue && inheritedLineSpacing == null)
                paraStyles.Add($"line-height:{lnSpcFactor * singleRatio:0.##}");

            // Indent / left margin. OOXML hanging-indent idiom (bullet outside, text inside)
            // is marL>=0 paired with indent<0 (|indent|==marL). We translate marL to CSS
            // padding-left (text starts at marL inside the shape content). The negative
            // indent is realised on the bullet span itself (margin-left:-|indent|), NOT
            // via text-indent on the para — text-indent would shift the line into the
            // shape's outer padding box and route bulletless paragraphs (line-spacing only)
            // right-flush via overflow interactions with width:100%.
            // CONSISTENCY(pptx-hanging-indent): bullet pulled left via its own margin.
            bool hasBullet0 = pProps?.GetFirstChild<Drawing.CharacterBullet>() != null
                              || pProps?.GetFirstChild<Drawing.AutoNumberedBullet>() != null;
            // Bulletless hanging indent: marL>0 paired with indent<0 hangs the
            // first line |indent| left of the marL margin (real PowerPoint renders
            // this even without a bullet — real-PowerPoint-confirmed). Emit the negative
            // text-indent with padding-left=marL so the first line hangs while
            // wrapped lines align to the margin. Guard the bleed the prior code
            // worried about: only emit the negative shift when marL >= |indent|,
            // so the first line stays inside the shape content box (the hanging-
            // indent idiom always satisfies this). Otherwise clamp to 0.
            if (pProps?.Indent?.HasValue == true && !hasBullet0)
            {
                var indentPt = Units.EmuToPt(pProps.Indent.Value);
                if (indentPt < 0)
                {
                    var marLPt = pProps?.LeftMargin?.HasValue == true
                        ? Units.EmuToPt(pProps.LeftMargin.Value) : 0;
                    if (marLPt < -indentPt) indentPt = 0; // would bleed past content edge
                }
                paraStyles.Add($"text-indent:{indentPt}pt");
            }
            // marL fallback chain: explicit slide marL wins; else inherited
            // master/layout bodyStyle lvlNpPr marL (R45 — paragraph relies on the
            // master's marL for indentation); else the level*36pt default cascade.
            long? inheritedMarL = (inheritedLvlPpr as Drawing.TextParagraphPropertiesType)?.LeftMargin?.Value;
            if (pProps?.LeftMargin?.HasValue == true)
                paraStyles.Add($"padding-left:{Units.EmuToPt(pProps.LeftMargin.Value)}pt");
            else if (inheritedMarL.HasValue)
                // R45: marL inherited from master bodyStyle lvlNpPr. PowerPoint reads
                // this so the bullet hangs and text is indented to the inherited marL.
                paraStyles.Add($"padding-left:{Units.EmuToPt(inheritedMarL.Value)}pt");
            else if ((pProps?.Level?.Value ?? 0) > 0)
                // R10b: no explicit marL but lvl>0 — approximate the default
                // lstStyle cascade. PowerPoint's built-in body text styles use
                // marL ≈ 0.5in (36pt) per indent level; reproduce that so leveled
                // paragraphs aren't flush-left in the preview. Explicit marL (above)
                // always wins; this only fills the inherited-default gap.
                paraStyles.Add($"padding-left:{(pProps!.Level!.Value) * 36}pt");

            // marR (right margin): text wraps before the right edge of the text body
            // by this amount. PowerPoint honors it (a paragraph with a large marR wraps
            // early, occupying only the left portion of the shape); mirror marL with
            // padding-right so the wrap boundary matches.
            long? inheritedMarR = (inheritedLvlPpr as Drawing.TextParagraphPropertiesType)?.RightMargin?.Value;
            if (pProps?.RightMargin?.HasValue == true)
                paraStyles.Add($"padding-right:{Units.EmuToPt(pProps.RightMargin.Value)}pt");
            else if (inheritedMarR.HasValue)
                paraStyles.Add($"padding-right:{Units.EmuToPt(inheritedMarR.Value)}pt");

            // RTL paragraph (Arabic / Hebrew). <a:pPr rtl="1"/> reverses
            // character order; emit CSS so the browser does the same. Without
            // this, Arabic PPT slides rendered visually mirrored in HTML
            // preview compared to PowerPoint itself.
            if (pProps?.RightToLeft?.Value == true)
                paraStyles.Add("direction:rtl;unicode-bidi:embed");

            // Bullet. R26-1: the slide paragraph may declare no bullet element at
            // all, in which case the bullet (buChar / buAutoNum / buNone / buBlip
            // with buFont / buSzPct / buClr) is INHERITED from the master/layout
            // bodyStyle lvlNpPr via the placeholder cascade. An explicit slide
            // bullet element (including <a:buNone/>) always wins. Determine the
            // source element to read all bullet sub-properties from.
            // Detect bullet child elements by LocalName so raw-injected elements
            // (buNone / buBlip with no typed SDK class in this context) also count.
            bool slideHasExplicitBullet = pProps?.ChildElements.Any(e =>
                e.LocalName is "buChar" or "buAutoNum" or "buNone" or "buBlip") == true;
            // The element we read bullet sub-properties from: the slide pPr if it
            // declared any bullet element, otherwise the inherited lvlNpPr.
            var bulletSource = slideHasExplicitBullet ? (OpenXmlElement?)pProps : inheritedLvlPpr;
            // Respect an explicit <a:buNone/> at the chosen source level (slide
            // override OR inherited) — it suppresses any bullet.
            bool buNone = bulletSource?.ChildElements.Any(e => e.LocalName == "buNone") == true;
            var bulletChar = buNone ? null : bulletSource?.GetFirstChild<Drawing.CharacterBullet>()?.Char?.Value;
            var bulletAuto = buNone ? null : bulletSource?.GetFirstChild<Drawing.AutoNumberedBullet>();
            // Image bullet (<a:buBlip>). SDK 3.x has no typed BulletBlip in this
            // context; detect by local name so a raw-injected buBlip also counts.
            var bulletBlip = buNone ? null : bulletSource?.ChildElements
                .FirstOrDefault(e => e.LocalName == "buBlip");
            var hasBullet = bulletChar != null || bulletAuto != null || bulletBlip != null;

            // R39: last-resort default bullet. When the slide→layout→master
            // bullet-resolution chain yields NO bullet (no buChar/buAutoNum/
            // buBlip) AND no explicit <a:buNone> anywhere in the chain, real
            // PowerPoint falls back to its built-in application default text
            // style, which bullets BODY/CONTENT outline placeholders (lvl1 = "•").
            // The officecli blank deck has no <p:bodyStyle> and no
            // <p:defaultTextStyle>, so the chain is empty and the bullet
            // silently vanished in HTML preview while PowerPoint shows it.
            //
            // Scoping (regression guardrails): apply ONLY to body/content-family
            // placeholders (IsBodyContentPlaceholder). Title/ctrTitle/subTitle,
            // plain text boxes, and any paragraph that already resolved a bullet
            // or buNone above are untouched — so R26 and every real template with
            // a defined master bodyStyle bullet render byte-identically.
            //
            // To avoid firing when the chain DID define a buNone on a level pPr
            // that the first-content `inheritedLvlPpr` happened to skip, do a
            // dedicated bullet-specific walk: search the chain for the first
            // level pPr that declares ANY bullet element. If that finds nothing,
            // the chain is genuinely silent and the default applies.
            if (!hasBullet && !buNone && !slideHasExplicitBullet
                && placeholderShape != null && placeholderPart != null
                && IsBodyContentPlaceholder(placeholderShape))
            {
                int lvl = pProps?.Level?.Value ?? 0;
                var bulletDefiningPpr = ResolvePlaceholderLevelPpr(placeholderShape, placeholderPart, lvl,
                    p => p.ChildElements.Any(e =>
                        e.LocalName is "buChar" or "buAutoNum" or "buNone" or "buBlip"));
                if (bulletDefiningPpr == null)
                {
                    // Genuinely nothing defined anywhere → synthesize PowerPoint's
                    // built-in default outline bullet for this level.
                    bulletChar = DefaultOutlineBulletChar(lvl);
                    hasBullet = true;
                }
            }

            // Determine paragraph emptiness BEFORE resolving bullets/auto-numbers.
            // Real PowerPoint suppresses the bullet/number glyph on a paragraph that
            // has no run text and no visible field/math, AND does NOT advance the
            // auto-number counter for it (empty paras are skipped in numbering, so
            // subsequent real items number continuously: 1.,2. not 3.,4.). The same
            // hasVisibleText/hasVisibleField notion is reused below for the &nbsp;
            // empty-line placeholder.
            var hasMath = para.OuterXml.Contains("oMath");
            var runs = para.Elements<Drawing.Run>().ToList();
            bool hasVisibleField = para.Elements<Drawing.Field>()
                .Any(f => !string.IsNullOrEmpty(ResolveFieldText(f, slideNumber)));
            bool hasVisibleText = runs.Any(r => !string.IsNullOrEmpty(r.Text?.Text)) || hasVisibleField;
            bool isEmptyPara = !hasVisibleText && !hasMath;

            // Resolve auto-numbered glyph (e.g. "1.", "a.", "iv.") and track per-scheme counter.
            string? autoNumGlyph = null;
            if (bulletAuto != null && !isEmptyPara)
            {
                int paraLevel = pProps?.Level?.Value ?? 0;
                // When a shallower-level paragraph interrupts a deeper numbered list,
                // PowerPoint resets the deeper levels' counters so they restart from
                // startAt the next time they appear (parent continues, child resets).
                if (lastAutoLevel.HasValue && paraLevel < lastAutoLevel.Value)
                {
                    var staleKeys = autoNumCounters.Keys.Where(k =>
                    {
                        var at = k.LastIndexOf('@');
                        return at >= 0 && int.TryParse(k[(at + 1)..], out var kl) && kl > paraLevel;
                    }).ToList();
                    foreach (var k in staleKeys) autoNumCounters.Remove(k);
                }
                lastAutoLevel = paraLevel;
                string schemeKey = (bulletAuto.Type?.HasValue == true && !string.IsNullOrEmpty(bulletAuto.Type.InnerText)
                    ? bulletAuto.Type.InnerText : "arabicPeriod") + "@" + paraLevel;
                if (lastAutoKey != schemeKey)
                {
                    // Each (scheme,level) keeps its own counter that persists across
                    // the text body. Only initialize when the key is genuinely new —
                    // returning to a parent level after a deeper sub-level must NOT
                    // reset the parent's count (PowerPoint continues 1.,2.,...).
                    if (!autoNumCounters.ContainsKey(schemeKey))
                        autoNumCounters[schemeKey] = 0;
                    lastAutoKey = schemeKey;
                }
                int startAt = bulletAuto.StartAt?.Value ?? 1;
                int n = autoNumCounters.TryGetValue(schemeKey, out var c) ? c : 0;
                int index = (n == 0 ? startAt : startAt + n);
                autoNumCounters[schemeKey] = n + 1;
                // Use the OOXML *value* (e.g. "alphaLcPeriod"), not the C# enum
                // member name ("AlphaLowerCharacterPeriod") — they differ, and the
                // glyph mapping keys off the OOXML token.
                string schemeToken = bulletAuto.Type?.HasValue == true && !string.IsNullOrEmpty(bulletAuto.Type.InnerText)
                    ? bulletAuto.Type.InnerText
                    : "arabicPeriod";
                autoNumGlyph = FormatAutoNumberGlyph(schemeToken, index);
            }
            else if (bulletAuto == null)
            {
                // A genuinely non-auto-numbered paragraph interrupts the list —
                // PowerPoint RESTARTS the numbering at startAt for the next numbered
                // paragraph (verified: 1.,2.,<plain>,1.). Clearing the counters (not
                // just nulling lastAutoKey) makes the next numbered para re-initialize
                // from startAt instead of continuing (...,3.). A pure level return
                // between consecutive numbered paras never enters this branch, so the
                // continue-parent-count behavior is unaffected.
                if (autoNumCounters.Count > 0) autoNumCounters.Clear();
                lastAutoKey = null;
            }
            // else (bulletAuto != null && isEmptyPara): an EMPTY auto-numbered
            // paragraph. PowerPoint skips it in numbering — it neither consumes a
            // number nor resets the counters, so following real items continue
            // (1.,<empty>,2.). Leave the counter state untouched (no clear, no
            // advance).

            sb.Append($"<div class=\"para\" style=\"{string.Join(";", paraStyles)}\">");

            if (hasBullet && !isEmptyPara)
            {
                // Image bullets have no glyph; use a generic marker so the
                // bullet span is non-empty (the source blip relationship is not
                // resolved into an <img> here \u2014 fallback marker only).
                var bullet = autoNumGlyph ?? bulletChar ?? (bulletBlip != null ? "\u25a0" : "\u2022");
                var buStyles = new List<string>();

                // Bullet font (<a:buFont typeface="..."/>) \u2014 apply font-family so a
                // symbol-font glyph (e.g. Wingdings "l") renders with the right face.
                var buFontTypeface = bulletSource?.GetFirstChild<Drawing.BulletFont>()?.Typeface?.Value;
                if (!string.IsNullOrEmpty(buFontTypeface))
                {
                    // A "+mn-lt"/"+mj-lt"/… typeface is a THEME font token, not a literal
                    // family — resolve it to the theme typeface (as runs do) instead of
                    // emitting the invalid CSS `font-family:+mn-lt` the browser ignores.
                    var resolvedBuFont = buFontTypeface.StartsWith("+", StringComparison.Ordinal)
                        ? (ResolveThemeFontToken(placeholderPart, buFontTypeface) ?? themeFontFallback)
                        : buFontTypeface;
                    if (!string.IsNullOrEmpty(resolvedBuFont))
                        buStyles.Add(CssFontFamilyWithFallback(resolvedBuFont));
                }

                // Bullet color: explicit buClr > first run color > default (inherit).
                // <a:buClr> is a CT_Color whose color child (srgbClr/schemeClr) sits
                // directly inside it — NOT wrapped in <a:solidFill> — so resolve the
                // child element directly rather than via ResolveFillColor (which
                // expects a solidFill wrapper). <a:buClrTx/> means "use text color"
                // and is the implicit default, so absence of buClr falls through.
                var buClrEl = bulletSource?.GetFirstChild<Drawing.BulletColor>();
                var bulletColor = ResolveBulletColor(buClrEl, themeColors);
                if (bulletColor == null)
                {
                    // Follow first run text color
                    var firstRun = para.Elements<Drawing.Run>().FirstOrDefault();
                    var firstRunFill = firstRun?.RunProperties?.GetFirstChild<Drawing.SolidFill>();
                    bulletColor = ResolveFillColor(firstRunFill, themeColors);
                }
                if (bulletColor != null) buStyles.Add($"color:{bulletColor}");

                // Bullet size: explicit buSzPts/buSzPct > first run size > default size
                var buSzPts = bulletSource?.GetFirstChild<Drawing.BulletSizePoints>();
                var buSzPct = bulletSource?.GetFirstChild<Drawing.BulletSizePercentage>();
                // normAutofit fontScale scales every run's font-size; the bullet glyph
                // must scale with it too, else an auto-shrunk paragraph shows an
                // oversized bullet next to small text (PowerPoint scales both).
                if (buSzPts?.Val?.HasValue == true)
                {
                    buStyles.Add($"font-size:{buSzPts.Val.Value / 100.0 * fontScale:0.##}pt");
                }
                else
                {
                    // Determine base font size from first run or default
                    var firstRun = para.Elements<Drawing.Run>().FirstOrDefault();
                    // Terminal 1800 (18pt) fallback mirrors RenderRun: when no size is
                    // set anywhere in the chain, the run renders at the 18pt spec default,
                    // so the bullet must match it (else the bullet shrinks to browser default).
                    var baseSizeHundredths = firstRun?.RunProperties?.FontSize?.Value ?? defaultFontSizeHundredths ?? 1800;
                    {
                        var pct = buSzPct?.Val?.HasValue == true ? buSzPct.Val.Value / 100000.0 : 1.0;
                        buStyles.Add($"font-size:{baseSizeHundredths / 100.0 * pct * fontScale:0.##}pt");
                    }
                }

                // Hanging-indent tab gap: size bullet span to match the negative
                // indent so text starts at marL regardless of bullet glyph width.
                // OOXML marL (e.g. 457200 EMU = 0.5in = 36pt) paired with indent
                // = -marL creates the hanging layout; we mirror it in CSS by
                // sizing the bullet to |indent| AND pulling it left with margin-left
                // by the same amount, so the bullet sits at the para outer edge
                // (shape content-left + 0) while text continues at marL inside.
                // We do NOT use text-indent here — text-indent on the para offsets
                // the line into the shape's outer padding box, putting the bullet
                // physically outside the shape's content area.
                long indentEmu = pProps?.Indent?.Value
                    ?? (inheritedLvlPpr as Drawing.TextParagraphPropertiesType)?.Indent?.Value
                    ?? 0;
                if (indentEmu < 0)
                {
                    var gapPt = Units.EmuToPt(-indentEmu);
                    buStyles.Add($"display:inline-block");
                    buStyles.Add($"width:{gapPt}pt");
                    buStyles.Add($"margin-left:-{gapPt}pt");
                }
                var buStyle = buStyles.Count > 0 ? $" style=\"{string.Join(";", buStyles)}\"" : "";
                sb.Append($"<span class=\"bullet\"{buStyle}>{HtmlEncode(bullet)}</span>");
            }

            // Check for OfficeMath (a14:m inside mc:AlternateContent) in paragraph XML
            var paraXml = para.OuterXml;
            if (paraXml.Contains("oMath"))
            {
                // AlternateContent is opaque to Descendants() — parse from XML.
                // A paragraph can interleave several a14:m equations with its text
                // runs; emit every match, not just the first (#228 dropped the
                // second equation of a two-equation paragraph).
                foreach (System.Text.RegularExpressions.Match mathMatch in System.Text.RegularExpressions.Regex.Matches(paraXml,
                    @"<m:oMathPara[^>]*>.*?</m:oMathPara>|<m:oMath[^>]*>.*?</m:oMath>",
                    System.Text.RegularExpressions.RegexOptions.Singleline))
                {
                    try
                    {
                        var wrapper = new OpenXmlUnknownElement("wrapper");
                        wrapper.InnerXml = mathMatch.Value;
                        var oMath = wrapper.Descendants().FirstOrDefault(e => e.LocalName == "oMathPara" || e.LocalName == "oMath");
                        if (oMath != null)
                        {
                            var latex = FormulaParser.ToLatex(oMath);
                            sb.Append($"<span class=\"katex-formula\" data-formula=\"{HtmlEncode(latex)}\"></span>");
                        }
                    }
                    catch { }
                }
            }

            // R35: per-paragraph tab-stop context. Read the explicit <a:tabLst>
            // (defined stops with absolute positions + alignment); fall back to
            // the body's default tab interval for tabs beyond the last stop.
            // RenderRun consumes this to advance each \t to its DEFINED column
            // instead of the old hardcoded 0.5in spacer.
            var tabCtx = new TabContext(pProps?.GetFirstChild<Drawing.TabStopList>(), defTabPt);

            // A paragraph is visually empty when it has no runs OR all its runs
            // carry empty <a:t> (RenderRun emits nothing for empty text). Real
            // PowerPoint still reserves a full line of vertical space for such a
            // paragraph, sized to its effective font size (the run's/endParaRPr's
            // sz, default 18pt). Without a placeholder, an empty-text run div
            // collapses to zero height and the blank line disappears. Emit a
            // sized &nbsp; so the line occupies its proper height. (hasVisibleText/
            // hasVisibleField/isEmptyPara computed above, before bullet resolution.)
            if (isEmptyPara)
            {
                // Empty paragraph (blank line) — size the &nbsp; to the effective
                // font size so it isn't zero-height. Precedence: first run rPr sz
                // > endParaRPr sz > inherited placeholder default > 18pt.
                var emptySzHundredths = runs.FirstOrDefault()?.RunProperties?.FontSize?.Value
                    ?? para.GetFirstChild<Drawing.EndParagraphRunProperties>()?.FontSize?.Value
                    ?? defaultFontSizeHundredths
                    ?? 1800;
                sb.Append($"<span style=\"font-size:{emptySzHundredths / 100.0 * fontScale:0.##}pt\">&nbsp;</span>");
            }
            else
            {
                // Paragraph-level default run properties (a:pPr/a:defRPr) supply a
                // fallback for every run in the paragraph that omits the property. This
                // applies to ALL shapes (placeholder or not), and sits ABOVE the
                // master/layout placeholder inheritance but BELOW an explicit run rPr.
                // Previously dropped entirely (e.g. <a:pPr><a:defRPr u="sng"/></a:pPr>
                // rendered without the underline PowerPoint applies).
                var paraDefRp = pProps?.GetFirstChild<Drawing.DefaultRunProperties>();
                // R26-3: inherited default bold/italic — paragraph defRPr wins over the
                // master/layout placeholder defRPr; applied when the run sets neither.
                bool? inhBold = paraDefRp?.Bold?.HasValue == true ? paraDefRp.Bold.Value
                    : inheritedDefRp?.Bold?.HasValue == true ? inheritedDefRp.Bold.Value : null;
                bool? inhItalic = paraDefRp?.Italic?.HasValue == true ? paraDefRp.Italic.Value
                    : inheritedDefRp?.Italic?.HasValue == true ? inheritedDefRp.Italic.Value : null;
                // Inherited caps (e.g. layout/master title defRPr cap="all", or a
                // paragraph-local defRPr): applied as a fallback when the run sets no cap.
                Drawing.TextCapsValues? inhCap = paraDefRp?.Capital?.HasValue == true ? paraDefRp.Capital.Value
                    : inheritedCapsRp?.Capital?.HasValue == true ? inheritedCapsRp.Capital.Value : null;
                // Inherited underline / strike: paragraph defRPr wins, else the
                // master/layout placeholder defRPr (mirrors inhBold/inhItalic — was
                // previously read from paraDefRp only, dropping master/layout u/strike).
                Drawing.TextUnderlineValues? inhU = paraDefRp?.Underline?.HasValue == true ? paraDefRp.Underline.Value
                    : inheritedUnderlineRp?.Underline?.HasValue == true ? inheritedUnderlineRp.Underline.Value : null;
                Drawing.TextStrikeValues? inhStrike = paraDefRp?.Strike?.HasValue == true ? paraDefRp.Strike.Value
                    : inheritedStrikeRp?.Strike?.HasValue == true ? inheritedStrikeRp.Strike.Value : null;
                // Inherited character spacing (spc): paragraph defRPr wins, else the
                // master/layout placeholder defRPr (mirrors inhCap/inhU/inhStrike — the
                // run-level rPr/spc consumer below had no inherited fallback, so a
                // defRPr-only spc was dropped).
                int? inhSpc = paraDefRp?.Spacing?.HasValue == true ? paraDefRp.Spacing.Value
                    : inheritedSpacingRp?.Spacing?.HasValue == true ? inheritedSpacingRp.Spacing.Value : null;
                // Paragraph defRPr font size / color override the placeholder defaults.
                int? paraSize = paraDefRp?.FontSize?.HasValue == true ? paraDefRp.FontSize.Value : defaultFontSizeHundredths;
                var paraColor = ResolveFillColor(paraDefRp?.GetFirstChild<Drawing.SolidFill>(), themeColors) ?? defaultRunColor;
                // R63: walk the paragraph's children IN DOCUMENT ORDER so that a
                // soft line break (<a:br>, Drawing.Break) interleaved between runs
                // emits its <br> at the right position. Previously runs were all
                // rendered first and breaks dumped at the paragraph's end, so a
                // "run / br / run" sequence collapsed onto one line. Each <br>
                // also resets the tab column (the next line's tabs measure from
                // the line's left edge again).
                foreach (var child in para.Elements())
                {
                    if (child is Drawing.Run run)
                    {
                        RenderRun(sb, run, themeColors, paraSize, placeholderPart, themeFontFallback, fontScale, paraColor, inhBold, inhItalic, tabCtx, inhCap, inhU, inhStrike, inhSpc);
                    }
                    else if (child is Drawing.Break)
                    {
                        sb.Append("<br>");
                        tabCtx.ResetColumn();
                    }
                    else if (child is Drawing.Field fld)
                    {
                        // <a:fld> (slide number, date, …). PowerPoint paints the
                        // field's text using the field's own run properties, so
                        // build a throwaway Drawing.Run from the field's <a:rPr>
                        // (cloned) + the EFFECTIVE field text and route it through
                        // RenderRun — the field then inherits the field's font /
                        // size / color exactly like a real run. slidenum is
                        // recomputed to the actual 1-based slide position when known
                        // (matching PowerPoint); other fields emit their cached <a:t>.
                        var fldText = ResolveFieldText(fld, slideNumber);
                        if (string.IsNullOrEmpty(fldText)) continue;
                        var fldRun = new Drawing.Run();
                        var fldRpr = fld.GetFirstChild<Drawing.RunProperties>();
                        if (fldRpr != null)
                            fldRun.RunProperties = (Drawing.RunProperties)fldRpr.CloneNode(true);
                        fldRun.Text = new Drawing.Text(fldText);
                        RenderRun(sb, fldRun, themeColors, paraSize, placeholderPart, themeFontFallback, fontScale, paraColor, inhBold, inhItalic, tabCtx, inhCap, inhU, inhStrike, inhSpc);
                    }
                }
            }

            sb.AppendLine("</div>");
            isFirstPara = false;
        }
    }

    // Effective display text for an <a:fld> in the HTML preview. For
    // type="slidenum" we prefer the ACTUAL 1-based slide number when the
    // renderer threaded it (PowerPoint recomputes slidenum to the real slide
    // position), falling back to the cached <a:t> when unknown. All other field
    // types (datetime*, etc.) emit their cached <a:t> verbatim. Mirrors
    // GetShapeText's field handling for the non-render extractor.
    private static string ResolveFieldText(Drawing.Field fld, int? slideNumber)
    {
        var cached = string.Concat(fld.Elements<Drawing.Text>().Select(t => t.Text));
        var fldType = fld.Type?.Value ?? "";
        if (fldType == "slidenum" && slideNumber.HasValue)
            return slideNumber.Value.ToString();
        return cached;
    }

    // R35: per-paragraph tab-stop tracking for the HTML preview. Holds the
    // paragraph's declared stops (absolute position in pt, sorted) plus the
    // body default tab interval, and a running horizontal column position that
    // RenderRun advances as it emits text/tabs within the paragraph line.
    private sealed class TabContext
    {
        // Each declared stop keeps its alignment (l|ctr|r|dec) — center/right tabs
        // anchor the following text segment at the stop column, not just advance to it.
        private readonly List<(double pos, char algn)> _stops = new();
        private readonly double _defTabPt;
        // Current horizontal pen position (pt) measured from the line's left edge.
        public double ColumnPt;

        public TabContext(Drawing.TabStopList? tabLst, double defTabPt)
        {
            _defTabPt = defTabPt > 0 ? defTabPt : 72.0;
            if (tabLst != null)
            {
                foreach (var t in tabLst.Elements<Drawing.TabStop>())
                {
                    if (t.Position?.HasValue == true)
                    {
                        var a = t.Alignment?.HasValue == true ? t.Alignment.InnerText : "l";
                        var algn = a switch { "ctr" => 'c', "r" => 'r', "dec" => 'd', _ => 'l' };
                        _stops.Add((Units.EmuToPt(t.Position.Value), algn));
                    }
                }
                _stops.Sort((x, y) => x.pos.CompareTo(y.pos));
            }
        }

        public void ResetColumn() => ColumnPt = 0;

        // Approximate text width (pt) using the same average-char-width heuristic
        // as the autofit line-wrap estimator (0.55*fontPt; CJK/full-width = 1 em).
        public static double MeasureText(string text, double fontPt)
        {
            double w = 0;
            foreach (var ch in text)
                w += ParseHelpers.IsCjkOrFullWidth(ch) ? fontPt : fontPt * 0.55;
            return w;
        }

        // Advance the column past printed text.
        public void AdvanceText(string text, double fontPt) => ColumnPt += MeasureText(text, fontPt);

        // Resolve the next tab stop strictly beyond the current column (declared
        // stop first, else the next default-interval multiple) and its alignment.
        public (double target, char algn) NextTab()
        {
            foreach (var s in _stops)
                if (s.pos > ColumnPt + 0.01) return s;
            int n = (int)Math.Floor(ColumnPt / _defTabPt) + 1;
            return (n * _defTabPt, 'l');
        }
    }

    // R35: encode run text to HTML, replacing each literal tab with a
    // stop-aware inline-block spacer (see the call site for the alignment
    // approximation note). When no tab context is available, fall back to the
    // legacy fixed 0.5in spacer so behavior is unchanged outside text bodies.
    private static string EncodeTextWithTabs(string text, TabContext? tabCtx, double fontPt)
    {
        if (text.IndexOf('\t') < 0)
        {
            tabCtx?.AdvanceText(text, fontPt);
            return HtmlEncode(text);
        }
        if (tabCtx == null)
            return HtmlEncode(text).Replace("\t",
                "<span class=\"tab-spacer\" style=\"display:inline-block;width:0.5in\"></span>");

        var sb = new StringBuilder();
        int start = 0;
        for (int i = 0; i < text.Length; i++)
        {
            if (text[i] != '\t') continue;
            string seg = text.Substring(start, i - start);
            if (seg.Length > 0)
            {
                tabCtx.AdvanceText(seg, fontPt);
                sb.Append(HtmlEncode(seg));
            }
            var (target, algn) = tabCtx.NextTab();
            // For center/right tabs, the text segment that FOLLOWS this tab (up to
            // the next tab or end of this run) must be anchored at the stop column,
            // not merely advanced to it. Look ahead within this run to measure it and
            // place its START so the segment is centered on / ends at the column.
            // (Cross-run segments degrade to a left tab — segW measures 0 here.)
            double anchorStart = target;
            if (algn is 'c' or 'r')
            {
                int next = text.IndexOf('\t', i + 1);
                string nextSeg = text.Substring(i + 1, (next < 0 ? text.Length : next) - (i + 1));
                double segW = TabContext.MeasureText(nextSeg, fontPt);
                anchorStart = algn == 'c' ? target - segW / 2 : target - segW;
            }
            double advancePt = anchorStart - tabCtx.ColumnPt;
            if (advancePt < 0) advancePt = 0;
            sb.Append($"<span class=\"tab-spacer\" style=\"display:inline-block;width:{advancePt:0.##}pt\"></span>");
            // Move the pen to where the following segment now begins so subsequent
            // text/tabs measure from the right place.
            tabCtx.ColumnPt += advancePt;
            start = i + 1;
        }
        string tail = text.Substring(start);
        if (tail.Length > 0)
        {
            tabCtx.AdvanceText(tail, fontPt);
            sb.Append(HtmlEncode(tail));
        }
        return sb.ToString();
    }

    private static void RenderRun(StringBuilder sb, Drawing.Run run, Dictionary<string, string> themeColors,
        int? defaultFontSizeHundredths = null, OpenXmlPart? part = null, string? themeFontFallback = null,
        double fontScale = 1.0, string? defaultRunColor = null,
        bool? inheritedBold = null, bool? inheritedItalic = null, TabContext? tabCtx = null, Drawing.TextCapsValues? inheritedCap = null,
        Drawing.TextUnderlineValues? inheritedUnderline = null, Drawing.TextStrikeValues? inheritedStrike = null,
        int? inheritedSpacing = null)
    {
        var text = run.Text?.Text ?? "";
        if (string.IsNullOrEmpty(text)) return;

        var styles = new List<string>();
        var rp = run.RunProperties;

        // Hyperlink resolution (RUN-level only; shape-level deferred).
        // Read <a:hlinkClick> from run.RunProperties, resolve relationship ID
        // via containing part's HyperlinkRelationships to an external URI.
        string? hyperlinkUrl = null;
        bool hasExplicitColor = rp?.GetFirstChild<Drawing.SolidFill>() != null
            || rp?.GetFirstChild<Drawing.GradientFill>() != null
            || rp?.GetFirstChild<Drawing.SchemeColor>() != null
            || rp?.GetFirstChild<Drawing.RgbColorModelHex>() != null
            // R51: an explicit <a:noFill/> means the glyph fill is transparent
            // (hollow text). It counts as an explicit fill so the inherited
            // cascade color fallback below does NOT bleed through.
            || rp?.GetFirstChild<Drawing.NoFill>() != null;
        bool hasExplicitUnderline = rp?.Underline?.HasValue == true;
        var hlinkClick = rp?.GetFirstChild<Drawing.HyperlinkOnClick>();
        if (hlinkClick?.Id?.Value is string relId && part != null)
        {
            try
            {
                var rel = part.HyperlinkRelationships.FirstOrDefault(r => r.Id == relId);
                // Reject javascript:/vbscript:/data: etc. — OOXML hyperlink
                // relationships are attacker-controlled and HtmlEncode does not
                // neutralize a dangerous scheme. Mirrors the Word/Excel previews.
                if (rel?.Uri != null && Core.HyperlinkUriValidator.IsSafeScheme(rel.Uri.ToString()))
                    hyperlinkUrl = rel.Uri.ToString();
            }
            catch { }
        }

        // Font. Theme references (typeface starts with "+") are resolved to
        // their concrete major/minor face via the textbody-supplied fallback;
        // runs with no <a:rPr> at all (common on auto-generated title text)
        // also pick up the fallback so a /theme bodyFont / headingFont change
        // is visible in HTML preview.
        var runFont = rp?.GetFirstChild<Drawing.LatinFont>()?.Typeface?.Value
            ?? rp?.GetFirstChild<Drawing.EastAsianFont>()?.Typeface?.Value
            ?? rp?.GetFirstChild<Drawing.ComplexScriptFont>()?.Typeface?.Value;
        string? resolvedRunFont;
        if (runFont == null)
        {
            // No explicit font child at all — inherit the placeholder default
            // (title→major, body→minor) resolved once per textbody.
            resolvedRunFont = themeFontFallback;
        }
        else if (runFont.StartsWith("+", StringComparison.Ordinal))
        {
            // Theme reference token: resolve the EXACT slot. The prefix selects
            // major (+mj) vs minor (+mn); the suffix selects the script slot
            // (-lt Latin, -ea EastAsian, -cs ComplexScript). Mapping each token
            // to its own theme font keeps the major-vs-minor and Latin-vs-EA-vs-CS
            // distinctions PowerPoint honors (TF-01/02/03), instead of collapsing
            // every "+" token onto the single placeholder fallback.
            var slotFont = ResolveThemeFontToken(part, runFont);
            // Latin tokens (-lt) may fall back to the placeholder default when the
            // theme slot is unreadable. EastAsian/ComplexScript tokens (-ea/-cs)
            // must NOT fall back to the Latin themeFontFallback — when their slot
            // is empty (the common blank-deck case) we emit no explicit override
            // and let the document-level CJK fallback chain apply.
            resolvedRunFont = slotFont
                ?? (runFont.EndsWith("-lt", StringComparison.Ordinal) ? themeFontFallback : null);
        }
        else
        {
            // Literal font name — emit it verbatim.
            resolvedRunFont = runFont;
        }
        if (!string.IsNullOrEmpty(resolvedRunFont))
            styles.Add(CssFontFamilyWithFallback(resolvedRunFont));

        // Size — use explicit run size, fall back to inherited placeholder
        // default, else the real-PowerPoint plain-textbox default of 18pt.
        // Decided independently of whether an <a:rPr> element exists: a plain
        // textbox created with no run properties (rp == null) still gets the
        // 18pt default so the browser doesn't fall back to its ~12pt default.
        // Placeholders keep their layout/master size via defaultFontSizeHundredths.
        // Bug #8(B): multiply by the textbody's normAutofit fontScale (1.0 = none).
        // Effective font size in pt (used both for the CSS and the R35 tab-column
        // text-advance heuristic below).
        double effFontPt = rp?.FontSize?.HasValue == true
            ? rp.FontSize.Value / 100.0 * fontScale
            : defaultFontSizeHundredths.HasValue
                ? defaultFontSizeHundredths.Value / 100.0 * fontScale
                : 18 * fontScale;
        if (rp?.FontSize?.HasValue == true)
            styles.Add($"font-size:{rp.FontSize.Value / 100.0 * fontScale:0.##}pt");
        else if (defaultFontSizeHundredths.HasValue)
            styles.Add($"font-size:{defaultFontSizeHundredths.Value / 100.0 * fontScale:0.##}pt");
        else
            styles.Add($"font-size:{18 * fontScale:0.##}pt");

        if (rp != null)
        {

            // Bold — explicit run bold wins; else inherit master/layout defRPr b.
            if (rp.Bold?.HasValue == true)
            {
                if (rp.Bold.Value) styles.Add("font-weight:bold");
            }
            else if (inheritedBold == true)
                styles.Add("font-weight:bold");

            // Italic — explicit run italic wins; else inherit defRPr i.
            if (rp.Italic?.HasValue == true)
            {
                if (rp.Italic.Value) styles.Add("font-style:italic");
            }
            else if (inheritedItalic == true)
                styles.Add("font-style:italic");

            // Underline + Strikethrough — both map to CSS text-decoration, which
            // is a single property: emitting it twice makes the cascade keep only
            // the last (dropping the first line). Build the combined line keyword
            // set ("underline", "line-through") plus an optional decoration STYLE
            // (double/wavy/dotted/dashed) and a thickness, then emit ONE
            // `text-decoration` declaration. text-decoration-style applies to all
            // lines, so a strike+underline mix where one wants double/wavy is a
            // known CSS limitation — the common single-line strike+underline case
            // renders both lines correctly.
            var decoLines = new List<string>();
            string? decoStyle = null;
            string? decoThickness = null;
            // Effective underline/strike: explicit run rPr wins; otherwise fall back to
            // the inherited value (paragraph a:pPr/a:defRPr or master/layout defRPr). An
            // explicit u="none"/strike="noStrike" on the run is respected (overrides the
            // inherited value), since rp.Underline.HasValue is true in that case.
            var effU = rp.Underline?.HasValue == true ? rp.Underline.Value : inheritedUnderline;
            if (effU != null && effU != Drawing.TextUnderlineValues.None)
            {
                decoLines.Add("underline");
                var u = effU.Value;
                if (u == Drawing.TextUnderlineValues.Double)
                {
                    // CONSISTENCY(underline-variants): mirrors WordHandler's
                    // emitter. Chromium renders this as two distinct lines at
                    // common font sizes (verified via Word HTML preview at 18pt).
                    decoStyle = "double";
                }
                else if (u == Drawing.TextUnderlineValues.Wavy)
                {
                    decoStyle = "wavy";
                }
                else if (u == Drawing.TextUnderlineValues.WavyHeavy
                    || u == Drawing.TextUnderlineValues.WavyDouble)
                {
                    // best-effort: CSS has no wavy+double; emit wavy thicker.
                    decoStyle = "wavy";
                    decoThickness = "2px";
                }
                else if (u == Drawing.TextUnderlineValues.Dotted)
                {
                    decoStyle = "dotted";
                }
                else if (u == Drawing.TextUnderlineValues.HeavyDotted)
                {
                    decoStyle = "dotted";
                    decoThickness = "2px";
                }
                else if (u == Drawing.TextUnderlineValues.Dash
                    || u == Drawing.TextUnderlineValues.DashLong)
                {
                    decoStyle = "dashed";
                }
                else if (u == Drawing.TextUnderlineValues.DashHeavy
                    || u == Drawing.TextUnderlineValues.DashLongHeavy
                    || u == Drawing.TextUnderlineValues.DotDashHeavy
                    || u == Drawing.TextUnderlineValues.DotDotDashHeavy)
                {
                    decoStyle = "dashed";
                    decoThickness = "2px";
                }
                else if (u == Drawing.TextUnderlineValues.DotDash
                    || u == Drawing.TextUnderlineValues.DotDotDash)
                {
                    // TODO CONSISTENCY(underline-variants): CSS has no dot-dash
                    // pattern; approximate with dashed.
                    decoStyle = "dashed";
                }
                else if (u == Drawing.TextUnderlineValues.Heavy)
                {
                    decoStyle = "solid";
                    decoThickness = "2px";
                }
                // else: exotic combos (Words, HeavyWords, etc.) fall back to plain underline.
            }

            var effStrike = rp.Strike?.HasValue == true ? rp.Strike.Value : inheritedStrike;
            if (effStrike != null && effStrike != Drawing.TextStrikeValues.NoStrike)
            {
                decoLines.Add("line-through");
                if (effStrike == Drawing.TextStrikeValues.DoubleStrike && decoStyle == null)
                {
                    // CONSISTENCY(underline-variants): like underline `double`,
                    // `line-through double` may render visually identical to
                    // single at typical font sizes in Chromium. Known limitation.
                    decoStyle = "double";
                }
            }

            // A hyperlinked run is underlined by default (unless it sets u= explicitly).
            // Fold that into decoLines so it joins any strikethrough in ONE
            // text-decoration declaration — a separate `text-decoration:underline` later
            // would, by CSS last-one-wins, silently drop a line-through emitted here.
            if (hlinkClick != null && !hasExplicitUnderline && !decoLines.Contains("underline"))
                decoLines.Add("underline");
            if (decoLines.Count > 0)
            {
                styles.Add($"text-decoration:{string.Join(" ", decoLines)}");
                if (decoStyle != null) styles.Add($"text-decoration-style:{decoStyle}");
                if (decoThickness != null) styles.Add($"text-decoration-thickness:{decoThickness}");

                // Underline color + width. PowerPoint exposes underline appearance
                // through TWO sibling slots: <a:uFill> (fill-only color) and <a:uLn>
                // (full line properties — its own <a:solidFill> color AND a @w width).
                // We honored uFill but dropped uLn, so a colored/thick underline
                // authored via uLn (the common slot when a custom WIDTH is wanted,
                // since uFill has no @w) rendered in the text color at default
                // thickness. Mirror NodeBuilder's Get path: uFill color wins; uLn is
                // the color fallback and the sole source of the line width.
                var uFill = rp.GetFirstChild<Drawing.UnderlineFill>();
                var uColor = uFill != null
                    ? ResolveFillColor(uFill.GetFirstChild<Drawing.SolidFill>(), themeColors) : null;
                var uLn = rp.GetFirstChild<Drawing.Underline>();
                if (uLn != null)
                {
                    uColor ??= ResolveFillColor(uLn.GetFirstChild<Drawing.SolidFill>(), themeColors);
                    if (uLn.Width?.HasValue == true)
                        // Explicit uLn @w (EMU) overrides any heuristic decoThickness
                        // emitted above (CSS last-wins within the inline style).
                        styles.Add($"text-decoration-thickness:{Units.EmuToPt(uLn.Width.Value):0.##}pt");
                }
                if (uColor != null)
                {
                    var textColor = ResolveFillColor(rp.GetFirstChild<Drawing.SolidFill>(), themeColors);
                    if (!string.Equals(uColor, textColor, StringComparison.OrdinalIgnoreCase))
                        styles.Add($"text-decoration-color:{uColor}");
                }
            }

            // Caps (rPr/@cap). all → text-transform:uppercase; small → font-variant-caps:small-caps
            // (browsers fall back to synthetic small-caps when the font lacks the SC variant).
            // Effective cap: the run's own cap wins; otherwise inherit from the
            // layout/master defRPr (e.g. an all-caps title style) so the text
            // renders uppercase like PowerPoint.
            var effCap = rp.Capital?.HasValue == true ? rp.Capital.Value : inheritedCap;
            if (effCap.HasValue && effCap.Value != Drawing.TextCapsValues.None)
            {
                if (effCap.Value == Drawing.TextCapsValues.All)
                    styles.Add("text-transform:uppercase");
                else if (effCap.Value == Drawing.TextCapsValues.Small)
                    // small caps: lowercase letters render as smaller capitals, but
                    // ORIGINALLY-uppercase letters stay full size. That is CSS
                    // `small-caps` — NOT `all-small-caps` (which would also shrink the
                    // already-uppercase letters, diverging from PowerPoint).
                    styles.Add("font-variant-caps:small-caps");
            }

            // Color
            var solidFill = rp.GetFirstChild<Drawing.SolidFill>();
            var color = ResolveFillColor(solidFill, themeColors);
            if (color != null)
                styles.Add($"color:{color}");
            // R7-2: no explicit run color and no gradient/scheme fill — inherit the
            // default color resolved from the master/layout placeholder cascade.
            else if (!hasExplicitColor && hlinkClick == null && defaultRunColor != null)
                styles.Add($"color:{defaultRunColor}");

            // Text highlight (a:highlight). Authored only in real PowerPoint /
            // via raw-set (no officecli prop), but `view` renders arbitrary
            // files so honor it. The highlight's color child has the same shape
            // as a solidFill's (srgbClr / schemeClr), so wrap it in a throwaway
            // SolidFill to reuse ResolveFillColor (theme + transforms).
            var highlight = rp.GetFirstChild<Drawing.Highlight>();
            var hlColorChild = highlight?.GetFirstChild<Drawing.RgbColorModelHex>()
                ?? (OpenXmlElement?)highlight?.GetFirstChild<Drawing.SchemeColor>();
            if (hlColorChild != null)
            {
                var hlFill = new Drawing.SolidFill(hlColorChild.CloneNode(true));
                var hlColor = ResolveFillColor(hlFill, themeColors);
                if (hlColor != null)
                    styles.Add($"background-color:{hlColor}");
            }

            // R51: explicit <a:noFill/> on the run — the glyph interior is
            // transparent (hollow text). Emit both color:transparent and
            // -webkit-text-fill-color:transparent so the fill is suppressed even
            // when a glyph outline (<a:ln> below) sets up paint-order:stroke fill.
            // noFill + outline → hollow glyphs (outline visible, interior see-through);
            // noFill + no outline → fully invisible (matches PowerPoint).
            if (rp.GetFirstChild<Drawing.NoFill>() != null)
            {
                styles.Add("color:transparent");
                styles.Add("-webkit-text-fill-color:transparent");
            }

            // Gradient text fill
            var gradFill = rp.GetFirstChild<Drawing.GradientFill>();
            if (gradFill != null)
            {
                var gradCss = GradientToCss(gradFill, themeColors);
                if (!string.IsNullOrEmpty(gradCss))
                {
                    styles.Add($"background:{gradCss}");
                    styles.Add("-webkit-background-clip:text");
                    styles.Add("background-clip:text");
                    styles.Add("-webkit-text-fill-color:transparent");
                }
            }

            // R9-3: image (blip) text fill — clip the image to the glyphs, mirroring
            // the gradFill text-fill approach above.
            var runBlipFill = rp.GetFirstChild<Drawing.BlipFill>();
            if (runBlipFill != null && part != null)
            {
                var dataUri = BlipToDataUri(runBlipFill, part);
                if (!string.IsNullOrEmpty(dataUri))
                {
                    styles.Add($"background-image:url('{dataUri}')");
                    styles.Add("background-size:cover");
                    styles.Add("-webkit-background-clip:text");
                    styles.Add("background-clip:text");
                    styles.Add("color:transparent");
                    styles.Add("-webkit-text-fill-color:transparent");
                }
            }

            // Run-level text outline (a:rPr/a:ln). PowerPoint strokes each glyph
            // edge; Chromium renders this via -webkit-text-stroke. Width is the
            // a:ln @w in EMU (12700 EMU = 1pt); convert to px (1pt = 4/3 px) so a
            // 3pt outline reads as a ~4px stroke. Color comes from the a:ln's
            // solidFill child (default black when absent). paint-order:stroke fill
            // keeps the fill painted on top so the stroke hugs the glyph outside.
            var runOutline = rp.GetFirstChild<Drawing.Outline>();
            if (runOutline != null)
            {
                double strokePx = runOutline.Width?.HasValue == true
                    ? Units.EmuToPt(runOutline.Width.Value) * 4.0 / 3.0
                    : 1.0;
                var strokeColor = ResolveFillColor(runOutline.GetFirstChild<Drawing.SolidFill>(), themeColors)
                    ?? "#000000";
                styles.Add($"-webkit-text-stroke:{strokePx:0.##}px {strokeColor}");
                styles.Add("paint-order:stroke fill");
            }

            // Character spacing
            var effSpc = rp.Spacing?.HasValue == true ? rp.Spacing.Value : inheritedSpacing;
            if (effSpc.HasValue)
                styles.Add($"letter-spacing:{effSpc.Value / 100.0:0.##}pt");

            // Run-level text shadow (<a:rPr><a:effectLst><a:outerShdw>). The same
            // helper the shape renderer uses produces filter:drop-shadow(...), which
            // on a <span> renders as a per-glyph shadow — matching PowerPoint's
            // shadowed text. Absent effectLst => no shadow (unchanged).
            var runEffects = rp.GetFirstChild<Drawing.EffectList>();
            if (runEffects != null)
            {
                var runShadowCss = EffectListToShadowCss(runEffects, themeColors);
                if (!string.IsNullOrEmpty(runShadowCss)) styles.Add(runShadowCss);
                // Run-level glow halo (<a:glow>) — same drop-shadow-based filter the
                // shape renderer uses; merge into any existing filter from the shadow.
                var runGlowCss = EffectListToGlowCss(runEffects, themeColors);
                if (!string.IsNullOrEmpty(runGlowCss))
                {
                    int fIdx = styles.FindIndex(s => s.StartsWith("filter:"));
                    if (fIdx >= 0) styles[fIdx] += " " + runGlowCss["filter:".Length..];
                    else styles.Add(runGlowCss);
                }
            }

            // Superscript/subscript. OOXML baseline is a raw integer where
            // 1000 == 1% (so super preset 30000 == 30%, sub preset -25000 == -25%).
            // Real PowerPoint shifts the glyph by baseline% × the ORIGINAL (pre-shrink)
            // font size as an ABSOLUTE distance — positive raises, negative lowers.
            // CSS vertical-align:<percent>% is a percentage of line-height (not
            // font-size) and the span's size is already reduced to 65% below, so a
            // percentage value renders ~1.7× too small. Emit an absolute pt offset
            // computed against effFontPt instead (vertical-align in pt is independent
            // of the span's own font-size) so distinct baselines render at the same
            // distinct heights real PowerPoint uses.
            if (rp.Baseline?.HasValue == true && rp.Baseline.Value != 0)
            {
                // PowerPoint auto-scales any baseline-shifted run to ~65% of its
                // computed size (Office super/subscript convention), regardless of
                // whether sz= was explicit. Replace the full-size font-size emitted
                // above with the reduced pt value so view-html matches real PPT.
                int fsIdx = styles.FindIndex(s => s.StartsWith("font-size:"));
                string reducedFontSize = $"font-size:{effFontPt * 0.65:0.##}pt";
                if (fsIdx >= 0) styles[fsIdx] = reducedFontSize;
                else styles.Add(reducedFontSize);
                double shiftPt = (rp.Baseline.Value / 100000.0) * effFontPt;  // baseline% × original font size
                styles.Add($"vertical-align:{shiftPt:0.##}pt");
            }
        }
        // R7-2: run with no <a:rPr> at all — still inherit the cascade default color.
        else
        {
            if (hlinkClick == null && defaultRunColor != null)
                styles.Add($"color:{defaultRunColor}");
            // R26-3: inherit master/layout defRPr bold/italic for a run with no rPr.
            if (inheritedBold == true) styles.Add("font-weight:bold");
            if (inheritedItalic == true) styles.Add("font-style:italic");
            // Inherited caps / underline / strike / spacing for a run with NO rPr.
            // The rp != null branch above consumes these (inheritedCap/Underline/
            // Strike/Spacing), but this branch was never updated as each was added,
            // so a bare run dropped an inherited all-caps / underline / strike /
            // tracking that PowerPoint applies (e.g. an all-caps title style or a
            // defRPr with spc on a run that carries no rPr of its own).
            if (inheritedCap == Drawing.TextCapsValues.All) styles.Add("text-transform:uppercase");
            else if (inheritedCap == Drawing.TextCapsValues.Small) styles.Add("font-variant-caps:small-caps");
            var inhDeco = new List<string>();
            if (inheritedUnderline != null && inheritedUnderline != Drawing.TextUnderlineValues.None)
                inhDeco.Add("underline");
            if (inheritedStrike != null && inheritedStrike != Drawing.TextStrikeValues.NoStrike)
                inhDeco.Add("line-through");
            if (inhDeco.Count > 0) styles.Add($"text-decoration:{string.Join(" ", inhDeco)}");
            if (inheritedSpacing.HasValue)
                styles.Add($"letter-spacing:{inheritedSpacing.Value / 100.0:0.##}pt");
        }

        // Auto-style hyperlink runs that lack explicit color/underline. Uses
        // theme-less fallback #0563C1 (PowerPoint default hyperlink color).
        // Shape-level hyperlinks are deferred (R14-supplemental).
        if (hlinkClick != null)
        {
            if (!hasExplicitColor)
                styles.Add(themeColors.TryGetValue("hlink", out var hlinkHex) && !string.IsNullOrEmpty(hlinkHex)
                    ? $"color:#{hlinkHex}"
                    : "color:#0563C1");
            // (underline is folded into the text-decoration declaration above so it
            // composites with any strikethrough instead of overriding it)
        }

        // Tab chars (literal U+0009 inside <a:t>, the form the Add path writes)
        // would collapse to a single space in HTML. Replace each with an
        // inline-block spacer whose WIDTH advances the running column to the
        // paragraph's next defined <a:tabLst> stop (R35), matching how
        // PowerPoint advances to the next tab stop. Falls back to the body's
        // default tab interval for tabs beyond the last declared stop.
        // The column is tracked across runs of a paragraph (see TabContext),
        // with printed text advancing the column via an average-char-width
        // heuristic. ctr/r/dec stops are APPROXIMATED as left-at-stop: a static
        // HTML renderer can't measure the following text width to center/right-
        // align it, so we advance to the stop and let text start there. This is
        // still exact column placement (a huge improvement over the old fixed
        // 0.5in) and is acceptable per the static-render limitation.
        string encoded = EncodeTextWithTabs(text, tabCtx, effFontPt);

        string inner = styles.Count > 0
            ? $"<span style=\"{string.Join(";", styles)}\">{encoded}</span>"
            : encoded;

        if (!string.IsNullOrEmpty(hyperlinkUrl))
        {
            sb.Append($"<a href=\"{HtmlEncode(hyperlinkUrl)}\" rel=\"noopener\">{inner}</a>");
        }
        else
        {
            sb.Append(inner);
        }
    }

    // Format an auto-numbered bullet glyph (e.g. "1.", "(a)", "iv)") for a given
    // OOXML scheme and 1-based index. Covers the common schemes emitted by
    // ApplyListStyle; unsupported schemes fall back to "N." arabic-period.
    private static string FormatAutoNumberGlyph(string key, int n)
    {
        // `key` is the OOXML buAutoNum value, of form
        // "{alphaLc|alphaUc|romanLc|romanUc|arabic|...}{Period|ParenBoth|ParenR|Plain|Minus}".
        // Match on this token directly — NOT the C# enum member name
        // (TextAutoNumberSchemeValues.AlphaLowerCharacterPeriod.ToString() yields
        // "AlphaLowerCharacterPeriod", which never matched "alphaLc" and silently
        // fell through to decimal for every non-arabic scheme).
        string body;
        if (key.StartsWith("alphaLc", StringComparison.OrdinalIgnoreCase) || key.StartsWith("AlphaLc", StringComparison.OrdinalIgnoreCase))
            body = ToAlpha(n, upper: false);
        else if (key.StartsWith("alphaUc", StringComparison.OrdinalIgnoreCase) || key.StartsWith("AlphaUc", StringComparison.OrdinalIgnoreCase))
            body = ToAlpha(n, upper: true);
        else if (key.StartsWith("romanLc", StringComparison.OrdinalIgnoreCase) || key.StartsWith("RomanLc", StringComparison.OrdinalIgnoreCase))
            body = ToRoman(n).ToLowerInvariant();
        else if (key.StartsWith("romanUc", StringComparison.OrdinalIgnoreCase) || key.StartsWith("RomanUc", StringComparison.OrdinalIgnoreCase))
            body = ToRoman(n);
        else
            body = n.ToString();

        if (key.EndsWith("Period", StringComparison.OrdinalIgnoreCase)) return body + ".";
        if (key.EndsWith("ParenBoth", StringComparison.OrdinalIgnoreCase)) return "(" + body + ")";
        if (key.EndsWith("ParenR", StringComparison.OrdinalIgnoreCase)) return body + ")";
        if (key.EndsWith("Minus", StringComparison.OrdinalIgnoreCase)) return "- " + body + " -";
        if (key.EndsWith("Plain", StringComparison.OrdinalIgnoreCase)) return body;
        return body + ".";
    }

    private static string ToAlpha(int n, bool upper)
    {
        if (n <= 0) n = 1;
        var sb = new StringBuilder();
        while (n > 0)
        {
            n--;
            sb.Insert(0, (char)((upper ? 'A' : 'a') + (n % 26)));
            n /= 26;
        }
        return sb.ToString();
    }

    // True when the shape is a title-class placeholder (title or centeredTitle).
    // Title placeholders inherit the theme major (heading) face; everything else
    // inherits minor (body). SubTitle uses the minor face in PowerPoint, so it
    // is intentionally NOT included here.
    private static bool IsTitlePlaceholder(Shape? shape)
    {
        var ph = shape?.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
            ?.GetFirstChild<PlaceholderShape>();
        if (ph?.Type?.HasValue != true) return false;
        var t = ph.Type.Value;
        return t == PlaceholderValues.Title || t == PlaceholderValues.CenteredTitle;
    }

    // R39: True when the shape is a BODY/CONTENT-family placeholder — the only
    // placeholder class PowerPoint bullets by its built-in application default
    // text style. The body/content family is: explicit type="body", type="obj"
    // (content placeholder), or a placeholder element present but with NO type
    // attribute (OOXML defaults a typeless <p:ph> to type=body). Title /
    // ctrTitle / subTitle placeholders have a buNone default style and must
    // stay bulletless; plain shapes (no <p:ph> at all) are not placeholders and
    // get no default bullet. When unsure, this returns false (conservative).
    private static bool IsBodyContentPlaceholder(Shape? shape)
    {
        var ph = shape?.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
            ?.GetFirstChild<PlaceholderShape>();
        if (ph == null) return false; // not a placeholder → no default bullet
        if (ph.Type?.HasValue != true) return true; // typeless <p:ph> defaults to body
        var t = ph.Type.Value;
        return t == PlaceholderValues.Body || t == PlaceholderValues.Object;
    }

    // R39: PowerPoint's built-in application default text style bullets each
    // outline level. lvl0/lvl1 is the critical "•". We mirror PowerPoint's
    // alternating default glyph set for deeper levels (•, –, •, –, …) with the
    // bullet font kept implicit (Arial-class) so the glyph renders cleanly.
    private static string DefaultOutlineBulletChar(int level)
        => (level % 2 == 0) ? "•" : "–"; // • for even levels, – for odd

    // Resolve the theme major/minor Latin typeface for the slide owning `part`.
    // Returns null when no theme is reachable (e.g. orphan text body, or a part
    // whose master->theme chain is incomplete). kind is "major" or "minor".
    private static string? ResolveThemeFontTypeface(OpenXmlPart? part, string kind)
        => ResolveThemeFont(part, major: kind == "major", slot: "lt");

    // Resolve a theme font reference token of the form "+{mj|mn}-{lt|ea|cs}"
    // (e.g. "+mj-lt", "+mn-ea") to its concrete typeface. The prefix picks
    // major vs minor; the suffix picks the script slot (Latin / EastAsian /
    // ComplexScript). Unrecognized tokens return null.
    private static string? ResolveThemeFontToken(OpenXmlPart? part, string token)
    {
        if (token.Length < 6 || token[0] != '+') return null;
        bool major;
        if (token.StartsWith("+mj-", StringComparison.Ordinal)) major = true;
        else if (token.StartsWith("+mn-", StringComparison.Ordinal)) major = false;
        else return null;
        var slot = token.Substring(4); // "lt" | "ea" | "cs"
        if (slot != "lt" && slot != "ea" && slot != "cs") return null;
        return ResolveThemeFont(part, major, slot);
    }

    // Read a single script slot (lt/ea/cs) of the major or minor theme font,
    // walking the SlidePart→Layout→Master→Theme chain. Returns null when no
    // theme is reachable, the slot is empty, or the slot is self-referential
    // (a theme declaring "+mj-lt" as its own typeface).
    private static string? ResolveThemeFont(OpenXmlPart? part, bool major, string slot)
    {
        var theme = part switch
        {
            SlidePart sp => sp.SlideLayoutPart?.SlideMasterPart?.ThemePart?.Theme,
            SlideLayoutPart lp => lp.SlideMasterPart?.ThemePart?.Theme,
            SlideMasterPart mp => mp.ThemePart?.Theme,
            _ => null,
        };
        var fontScheme = theme?.ThemeElements?.FontScheme;
        var fontGroup = major ? (OpenXmlElement?)fontScheme?.MajorFont : fontScheme?.MinorFont;
        var typeface = slot switch
        {
            "ea" => fontGroup?.GetFirstChild<Drawing.EastAsianFont>()?.Typeface?.Value,
            "cs" => fontGroup?.GetFirstChild<Drawing.ComplexScriptFont>()?.Typeface?.Value,
            _ => fontGroup?.GetFirstChild<Drawing.LatinFont>()?.Typeface?.Value,
        };
        if (string.IsNullOrEmpty(typeface)) return null;
        // Theme entries are sometimes self-referential ("+mj-lt"); skip those.
        if (typeface.StartsWith("+", StringComparison.Ordinal)) return null;
        return typeface;
    }

    private static string ToRoman(int n)
    {
        if (n <= 0) return n.ToString();
        int[] values = { 1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1 };
        string[] numerals = { "M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I" };
        var sb = new StringBuilder();
        for (int i = 0; i < values.Length; i++)
        {
            while (n >= values[i]) { sb.Append(numerals[i]); n -= values[i]; }
        }
        return sb.ToString();
    }
}
