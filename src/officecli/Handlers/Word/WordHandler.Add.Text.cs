// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    private string AddParagraph(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        string resultPath;
        var para = new Paragraph();
        AssignParaId(para);
        var pProps = new ParagraphProperties();

        // CONSISTENCY(style-dual-key): mirror SetParagraph and AddStyle —
        // accept canonical readback aliases (styleId, styleName) so a
        // get→add clone of a paragraph round-trips its style intact.
        // styleName resolves the display name through the styles part;
        // falls back to verbatim if no match (lenient-input pattern).
        if (properties.TryGetValue("style", out var style)
            || properties.TryGetValue("styleId", out style)
            || properties.TryGetValue("styleid", out style))
            pProps.ParagraphStyleId = new ParagraphStyleId { Val = style };
        else if (properties.TryGetValue("styleName", out var styleName)
            || properties.TryGetValue("stylename", out styleName))
            pProps.ParagraphStyleId = new ParagraphStyleId { Val = ResolveStyleIdFromName(styleName) ?? styleName };
        if (properties.TryGetValue("align", out var alignment) || properties.TryGetValue("alignment", out alignment))
            pProps.Justification = new Justification { Val = ParseJustification(alignment) };
        // Reading direction (Arabic / Hebrew). 'rtl' enables <w:bidi/> AND
        // writes <w:rtl/> on the paragraph mark (so any later runs added
        // via Set inherit the run-level direction without a separate flag).
        // CONSISTENCY(rtl-cascade): mirrors SetElementParagraph — direction
        // is a paragraph-scope shorthand for "this paragraph is fully RTL".
        bool? paraRtl = null;
        if (properties.TryGetValue("direction", out var dirRaw)
            || properties.TryGetValue("dir", out dirRaw)
            || properties.TryGetValue("bidi", out dirRaw))
        {
            paraRtl = ParseDirectionRtl(dirRaw);
            if (paraRtl.Value)
            {
                pProps.BiDi = new BiDi();
                var markRPr = pProps.ParagraphMarkRunProperties ?? pProps.AppendChild(new ParagraphMarkRunProperties());
                ApplyRunFormatting(markRPr, "rtl", "true");
            }
            else
            {
                // Clear semantics: direction=ltr removes any prior bidi marker.
                // R19-fuzz-1/2 + R20-fuzz-11: if ANY inherited source carries
                // bidi=true (style chain, enclosing section, docDefaults, or
                // numbering lvl), simply clearing pPr.bidi re-inherits RTL —
                // the user's explicit ltr override would silently disappear.
                // Emit <w:bidi w:val="0"/> to cancel. Style-chain check happens
                // here (no parent context needed); section / docDefaults /
                // numbering checks are deferred until after the paragraph is
                // inserted into the tree (see post-insert HasInheritedBidi
                // pass below). Mirrors paragraph Set/ApplyDirectionCascade.
                pProps.RemoveAllChildren<BiDi>();
                var addStyleId = pProps.ParagraphStyleId?.Val?.Value;
                if (addStyleId != null && StyleChainHasBidi(addStyleId))
                {
                    pProps.BiDi = new BiDi { Val = new DocumentFormat.OpenXml.OnOffValue(false) };
                }
                var markRPr = pProps.ParagraphMarkRunProperties;
                markRPr?.RemoveAllChildren<RightToLeftText>();
            }
        }
        // CONSISTENCY(rtl-cascade): `rtl=true` on a paragraph add should
        // mirror direction=rtl — write <w:bidi/> on pPr AND <w:rtl/> on
        // the paragraph mark so the paragraph is fully RTL (not just any
        // text run). Without this, `add p --prop rtl=true` left the
        // paragraph LTR and only flagged individual runs.
        if (paraRtl == null && properties.TryGetValue("rtl", out var paraRtlRaw) && IsTruthy(paraRtlRaw))
        {
            paraRtl = true;
            pProps.BiDi = new BiDi();
            var markRPr = pProps.ParagraphMarkRunProperties ?? pProps.AppendChild(new ParagraphMarkRunProperties());
            ApplyRunFormatting(markRPr, "rtl", "true");
        }
        // Complex-script run flags (bCs/iCs/szCs) hoisted above the text
        // block so an `add p --prop bold.cs=true` without explicit text
        // still records the flag on the paragraph mark rPr — matches how
        // bare bold round-trips via the generic TypedAttributeFallback
        // path. Without this, schema-strict round-trip tests for
        // bold.cs/italic.cs/size.cs lose the flag (no run carrier exists
        // when text is absent, and TypedAttributeFallback can't synthesise
        // <w:bCs/> / <w:iCs/> / <w:szCs/> child elements from a key).
        if ((properties.TryGetValue("bold.cs", out var paraBoldCs)
                || properties.TryGetValue("font.bold.cs", out paraBoldCs)))
        {
            var markRPr = pProps.ParagraphMarkRunProperties ?? pProps.AppendChild(new ParagraphMarkRunProperties());
            ApplyRunFormatting(markRPr, "bold.cs", paraBoldCs);
        }
        if ((properties.TryGetValue("italic.cs", out var paraItalicCs)
                || properties.TryGetValue("font.italic.cs", out paraItalicCs)))
        {
            var markRPr = pProps.ParagraphMarkRunProperties ?? pProps.AppendChild(new ParagraphMarkRunProperties());
            ApplyRunFormatting(markRPr, "italic.cs", paraItalicCs);
        }
        if (properties.TryGetValue("size.cs", out var paraSizeCs)
            || properties.TryGetValue("font.size.cs", out paraSizeCs))
        {
            var markRPr = pProps.ParagraphMarkRunProperties ?? pProps.AppendChild(new ParagraphMarkRunProperties());
            ApplyRunFormatting(markRPr, "size.cs", paraSizeCs);
        }
        if (properties.TryGetValue("firstlineindent", out var indent) || properties.TryGetValue("firstLineIndent", out indent))
        {
            // Lenient input: accept "2cm", "0.5in", "18pt", or bare twips (backward compat).
            // SpacingConverter.ParseWordSpacing treats bare numbers as twips.
            var indentTwips = SpacingConverter.ParseWordSpacing(indent);
            if (indentTwips > 31680)
                throw new OverflowException($"First line indent value out of range (0-31680 twips): {indent}");
            pProps.Indentation = new Indentation
            {
                FirstLine = indentTwips.ToString()  // raw twips, consistent with Set and Get
            };
        }
        if (properties.TryGetValue("spacebefore", out var sb4) || properties.TryGetValue("spaceBefore", out sb4))
        {
            var spacing = pProps.SpacingBetweenLines ?? (pProps.SpacingBetweenLines = new SpacingBetweenLines());
            spacing.Before = SpacingConverter.ParseWordSpacing(sb4).ToString();
        }
        if (properties.TryGetValue("spaceafter", out var sa4) || properties.TryGetValue("spaceAfter", out sa4))
        {
            var spacing = pProps.SpacingBetweenLines ?? (pProps.SpacingBetweenLines = new SpacingBetweenLines());
            spacing.After = SpacingConverter.ParseWordSpacing(sa4).ToString();
        }
        if (properties.TryGetValue("linespacing", out var ls4) || properties.TryGetValue("lineSpacing", out ls4))
        {
            var spacing = pProps.SpacingBetweenLines ?? (pProps.SpacingBetweenLines = new SpacingBetweenLines());
            var (twips, isMultiplier) = SpacingConverter.ParseWordLineSpacing(ls4);
            spacing.Line = twips.ToString();
            spacing.LineRule = isMultiplier ? LineSpacingRuleValues.Auto : LineSpacingRuleValues.Exact;
        }
        // Numbering properties. Parallel branches so `ilvl` alone still
        // emits <w:ilvl> (matching `set --prop ilvl=N` behaviour); both
        // inputs are range-checked so schema-invalid values never reach XML.
        if (properties.TryGetValue("numid", out var numId)
            || properties.TryGetValue("numId", out numId)
            || properties.TryGetValue("listId", out numId)
            || properties.TryGetValue("listid", out numId))
        {
            var numIdVal = ParseHelpers.SafeParseInt(numId, "numid");
            if (numIdVal < 0)
                throw new ArgumentException($"numId must be >= 0 (got {numIdVal}).");
            // numId=0 is OOXML's way of saying "remove numbering" (no-list sentinel).
            // Positive numIds must reference an existing <w:num> to avoid silent dangling
            // references — Word renders such paragraphs without any list marker.
            if (numIdVal > 0)
            {
                var numbering = _doc.MainDocumentPart?.NumberingDefinitionsPart?.Numbering;
                var numExists = numbering?.Elements<NumberingInstance>()
                    .Any(n => n.NumberID?.Value == numIdVal) ?? false;
                if (!numExists)
                    throw new ArgumentException(
                        $"numId={numIdVal} not found in /numbering. " +
                        "Create the num first (add /numbering --type num), or use numId=0 to remove numbering.");
            }
            var numPr = pProps.NumberingProperties ?? (pProps.NumberingProperties = new NumberingProperties());
            numPr.NumberingId = new NumberingId { Val = numIdVal };
        }
        // Accept both "numlevel" and "ilvl" (the OOXML name); works with or
        // without numId to stay in sync with `set --prop ilvl=N`.
        if (properties.TryGetValue("numlevel", out var numLevel)
            || properties.TryGetValue("ilvl", out numLevel)
            || properties.TryGetValue("listLevel", out numLevel)
            || properties.TryGetValue("listlevel", out numLevel))
        {
            var ilvlVal = ParseHelpers.SafeParseInt(numLevel, "ilvl");
            if (ilvlVal < 0 || ilvlVal > 8)
                throw new ArgumentException($"ilvl must be in range 0..8 (got {ilvlVal}).");
            var numPr = pProps.NumberingProperties ?? (pProps.NumberingProperties = new NumberingProperties());
            numPr.NumberingLevelReference = new NumberingLevelReference { Val = ilvlVal };
        }
        if (properties.TryGetValue("shd", out var pShdVal) || properties.TryGetValue("shading", out pShdVal))
        {
            var shdParts = pShdVal.Split(';');
            var shd = new Shading();
            if (shdParts.Length == 1)
            {
                shd.Val = ShadingPatternValues.Clear;
                shd.Fill = SanitizeHex(shdParts[0]);
            }
            else if (shdParts.Length >= 2)
            {
                // Check if the pattern/color order is reversed (hex color in pattern position)
                var patternPart = shdParts[0].TrimStart('#');
                if (patternPart.Length >= 6 && patternPart.All(char.IsAsciiHexDigit))
                {
                    // Auto-swap: treat as "clear;COLOR" (user put color first)
                    Console.Error.WriteLine($"Warning: '{shdParts[0]}' looks like a color in the pattern position. Auto-swapping to: clear;{shdParts[0]}");
                    shd.Val = ShadingPatternValues.Clear;
                    shd.Fill = SanitizeHex(shdParts[0]);
                }
                else
                {
                    WarnIfShadingOrderWrong(shdParts[0]); shd.Val = new ShadingPatternValues(shdParts[0]);
                    shd.Fill = SanitizeHex(shdParts[1]);
                    if (shdParts.Length >= 3) shd.Color = SanitizeHex(shdParts[2]);
                }
            }
            pProps.Shading = shd;
        }
        if (properties.TryGetValue("leftindent", out var addLI) || properties.TryGetValue("leftIndent", out addLI) || properties.TryGetValue("indentleft", out addLI) || properties.TryGetValue("indent", out addLI))
        {
            var ind = pProps.Indentation ?? (pProps.Indentation = new Indentation());
            // CONSISTENCY(lenient-spacing): route through SpacingConverter so indent accepts
            // "2cm"/"0.5in"/"24pt"/bare twips — parity with spaceBefore/spaceAfter/lineSpacing.
            ind.Left = SpacingConverter.ParseWordSpacing(addLI).ToString();
        }
        if (properties.TryGetValue("rightindent", out var addRI) || properties.TryGetValue("rightIndent", out addRI) || properties.TryGetValue("indentright", out addRI))
        {
            var ind = pProps.Indentation ?? (pProps.Indentation = new Indentation());
            // CONSISTENCY(lenient-spacing): see leftindent above.
            ind.Right = SpacingConverter.ParseWordSpacing(addRI).ToString();
        }
        if (properties.TryGetValue("hangingindent", out var addHI) || properties.TryGetValue("hangingIndent", out addHI) || properties.TryGetValue("hanging", out addHI))
        {
            var ind = pProps.Indentation ?? (pProps.Indentation = new Indentation());
            // CONSISTENCY(lenient-spacing): see leftindent above.
            ind.Hanging = SpacingConverter.ParseWordSpacing(addHI).ToString();
            ind.FirstLine = null;
        }
        // firstlineindent already handled above (line ~66-74) with × 480 conversion
        if ((properties.TryGetValue("keepnext", out var addKN) || properties.TryGetValue("keepNext", out addKN)) && IsTruthy(addKN))
            pProps.KeepNext = new KeepNext();
        if ((properties.TryGetValue("keeplines", out var addKL) || properties.TryGetValue("keeptogether", out addKL) || properties.TryGetValue("keepLines", out addKL) || properties.TryGetValue("keepTogether", out addKL)) && IsTruthy(addKL))
            pProps.KeepLines = new KeepLines();
        if ((properties.TryGetValue("pagebreakbefore", out var addPBB) || properties.TryGetValue("pageBreakBefore", out addPBB)) && IsTruthy(addPBB))
            pProps.PageBreakBefore = new PageBreakBefore();
        // fuzz-2: paragraph-context `break=newPage` alias → pageBreakBefore=true.
        // Mirrors Set-side handling in WordHandler.Set.cs (case "break").
        if (properties.TryGetValue("break", out var addBrk))
        {
            bool pbb = addBrk?.ToLowerInvariant() switch
            {
                "newpage" or "page" or "nextpage" or "pagebreak" => true,
                "none" or "" or null => false,
                _ => IsTruthy(addBrk)
            };
            if (pbb) pProps.PageBreakBefore = new PageBreakBefore();
        }
        if (properties.TryGetValue("widowcontrol", out var addWC) || properties.TryGetValue("widowControl", out addWC))
        {
            if (IsTruthy(addWC))
                pProps.WidowControl = new WidowControl();
            else
                pProps.WidowControl = new WidowControl { Val = false };
        }
        // CONSISTENCY(add-set-symmetry): Set accepts wordWrap via the toggle
        // fallback in WordHandler.Set.cs; Add mirrors it so callers can build
        // CJK right-aligned paragraphs (which need wordWrap=false to preserve
        // trailing whitespace on right-aligned lines) in one call.
        if (properties.TryGetValue("wordwrap", out var addWW) || properties.TryGetValue("wordWrap", out addWW))
        {
            pProps.WordWrap = IsTruthy(addWW)
                ? new WordWrap()
                : new WordWrap { Val = false };
        }
        // CONSISTENCY(add-set-symmetry): Set supports contextualSpacing (WordHandler.Set.cs:529);
        // Add must accept the same prop so the "Add then Get" lifecycle test pattern works
        // without falling back to a separate Set call. Both true and false write an
        // explicit element — `false` is meaningful when a parent style sets
        // contextualSpacing=true, since omitting the element would inherit the
        // style's `true`. Setting `Val=false` explicitly overrides.
        if (properties.TryGetValue("contextualspacing", out var addCS) || properties.TryGetValue("contextualSpacing", out addCS))
            pProps.ContextualSpacing = IsTruthy(addCS)
                ? new ContextualSpacing()
                : new ContextualSpacing { Val = false };
        foreach (var (pk, pv) in properties)
        {
            // CONSISTENCY(add-set-symmetry): Set accepts border.top/bottom/left/right/between/bar
            // (and bare "border"/"border.all"); Add must accept the same vocabulary so the
            // Add → Get → verify lifecycle works without a follow-up Set call.
            if (pk.StartsWith("pbdr", StringComparison.OrdinalIgnoreCase)
                || pk.StartsWith("border", StringComparison.OrdinalIgnoreCase))
                ApplyParagraphBorders(pProps, pk, pv);
        }
        if (properties.TryGetValue("liststyle", out var listStyle) || properties.TryGetValue("listStyle", out listStyle))
        {
            para.AppendChild(pProps);
            int? startVal = null;
            if (properties.TryGetValue("start", out var sv))
                startVal = ParseHelpers.SafeParseInt(sv, "start");
            int? levelVal = null;
            if (properties.TryGetValue("listLevel", out var ll) || properties.TryGetValue("listlevel", out ll) || properties.TryGetValue("level", out ll) || properties.TryGetValue("numlevel", out ll))
            {
                levelVal = ParseHelpers.SafeParseInt(ll, "listLevel");
                // OOXML ST_DecimalNumber ilvl is bound to 0..8 (ECMA-376
                // §17.9.3) — Word silently drops out-of-range values, so
                // reject up-front to keep round-trip lossless.
                if (levelVal < 0 || levelVal > 8)
                    throw new ArgumentException($"listLevel must be in range 0..8 (got {levelVal}).");
            }
            ApplyListStyle(para, listStyle, startVal, levelVal, containerHint: parent);
            // pProps already appended, skip the append below
            goto paragraphPropsApplied;
        }

        para.AppendChild(pProps);
        paragraphPropsApplied:

        if (properties.TryGetValue("text", out var text))
        {
            var rProps = new RunProperties();
            // Per-script font slots (font.latin / font.ea / font.cs) write
            // to ascii+hAnsi / eastAsia / cs respectively. Bare 'font'
            // populates ascii+hAnsi+eastAsia for backward compatibility.
            // Build a single RunFonts so per-slot values compose cleanly
            // when the user supplies more than one (e.g. font.latin=Calibri
            // + font.cs=Arabic Typesetting on the same run).
            string? rfAscii = null, rfHAnsi = null, rfEa = null, rfCs = null;
            if (properties.TryGetValue("font", out var font) || properties.TryGetValue("font.name", out font))
            {
                rfAscii = font; rfHAnsi = font; rfEa = font;
            }
            if (properties.TryGetValue("font.latin", out var fLatin))
            {
                rfAscii = fLatin; rfHAnsi = fLatin;
            }
            if (properties.TryGetValue("font.ea", out var fEa)
                || properties.TryGetValue("font.eastasia", out fEa)
                || properties.TryGetValue("font.eastasian", out fEa))
            {
                rfEa = fEa;
            }
            if (properties.TryGetValue("font.cs", out var fCs)
                || properties.TryGetValue("font.complexscript", out fCs)
                || properties.TryGetValue("font.complex", out fCs))
            {
                rfCs = fCs;
            }
            if (rfAscii != null || rfHAnsi != null || rfEa != null || rfCs != null)
            {
                var rFonts = new RunFonts();
                if (rfAscii != null) rFonts.Ascii = rfAscii;
                if (rfHAnsi != null) rFonts.HighAnsi = rfHAnsi;
                if (rfEa != null) rFonts.EastAsia = rfEa;
                if (rfCs != null) rFonts.ComplexScript = rfCs;
                rProps.AppendChild(rFonts);
            }
            if (properties.TryGetValue("size", out var size) || properties.TryGetValue("font.size", out size) || properties.TryGetValue("fontsize", out size))
            {
                rProps.AppendChild(new FontSize { Val = ((int)Math.Round(ParseFontSize(size) * 2, MidpointRounding.AwayFromZero)).ToString() });
            }
            if ((properties.TryGetValue("bold", out var bold) || properties.TryGetValue("font.bold", out bold)) && IsTruthy(bold))
                rProps.Bold = new Bold();
            if ((properties.TryGetValue("bold.cs", out var boldCs)
                    || properties.TryGetValue("font.bold.cs", out boldCs))
                && IsTruthy(boldCs))
                rProps.BoldComplexScript = new BoldComplexScript();
            if ((properties.TryGetValue("italic", out var pItalic) || properties.TryGetValue("font.italic", out pItalic)) && IsTruthy(pItalic))
                rProps.Italic = new Italic();
            if ((properties.TryGetValue("italic.cs", out var italicCs)
                    || properties.TryGetValue("font.italic.cs", out italicCs))
                && IsTruthy(italicCs))
                rProps.ItalicComplexScript = new ItalicComplexScript();
            if (properties.TryGetValue("size.cs", out var sizeCs)
                || properties.TryGetValue("font.size.cs", out sizeCs))
            {
                rProps.FontSizeComplexScript = new FontSizeComplexScript
                {
                    Val = ((int)Math.Round(ParseFontSize(sizeCs) * 2, MidpointRounding.AwayFromZero)).ToString()
                };
            }
            if (properties.TryGetValue("color", out var pColor) || properties.TryGetValue("font.color", out pColor))
            {
                // CONSISTENCY(theme-color): Add paragraph color must accept
                // scheme color names (accent1, dark2, hyperlink, …) the same
                // way ApplyRunFormatting (Set path) does — otherwise
                // Add(.., {color=accent1}) would call SanitizeHex on the
                // scheme name and produce garbage hex.
                var pSchemeName = OfficeCli.Core.ParseHelpers.NormalizeSchemeColorName(pColor);
                if (pSchemeName != null)
                    rProps.Color = new Color { Val = "auto", ThemeColor = new EnumValue<ThemeColorValues>(new ThemeColorValues(pSchemeName)) };
                else
                    rProps.Color = new Color { Val = SanitizeHex(pColor) };
            }
            if (properties.TryGetValue("underline", out var pUnderline) || properties.TryGetValue("font.underline", out pUnderline))
            {
                var ulVal = NormalizeUnderlineValue(pUnderline);
                rProps.Underline = new Underline { Val = new UnderlineValues(ulVal) };
            }
            if ((properties.TryGetValue("strike", out var pStrike)
                    || properties.TryGetValue("strikethrough", out pStrike)
                    || properties.TryGetValue("font.strike", out pStrike)
                    || properties.TryGetValue("font.strikethrough", out pStrike))
                && IsTruthy(pStrike))
                rProps.Strike = new Strike();
            if (properties.TryGetValue("highlight", out var pHighlight))
                rProps.Highlight = new Highlight { Val = ParseHighlightColor(pHighlight) };
            if ((properties.TryGetValue("caps", out var pCaps)
                    || properties.TryGetValue("allcaps", out pCaps)
                    || properties.TryGetValue("allCaps", out pCaps))
                && IsTruthy(pCaps))
                rProps.Caps = new Caps();
            if (properties.TryGetValue("smallcaps", out var pSmallCaps) || properties.TryGetValue("smallCaps", out pSmallCaps))
            {
                if (IsTruthy(pSmallCaps)) rProps.SmallCaps = new SmallCaps();
            }
            if (properties.TryGetValue("dstrike", out var pDstrike) && IsTruthy(pDstrike))
                rProps.DoubleStrike = new DoubleStrike();
            if (properties.TryGetValue("vanish", out var pVanish) && IsTruthy(pVanish))
                rProps.Vanish = new Vanish();
            if (properties.TryGetValue("outline", out var pOutline) && IsTruthy(pOutline))
                rProps.Outline = new Outline();
            if (properties.TryGetValue("shadow", out var pShadow) && IsTruthy(pShadow))
                rProps.Shadow = new Shadow();
            if (properties.TryGetValue("emboss", out var pEmboss) && IsTruthy(pEmboss))
                rProps.Emboss = new Emboss();
            if (properties.TryGetValue("imprint", out var pImprint) && IsTruthy(pImprint))
                rProps.Imprint = new Imprint();
            if (properties.TryGetValue("noproof", out var pNoProof) && IsTruthy(pNoProof))
                rProps.NoProof = new NoProof();
            // Run-level rtl: explicit `rtl=true` OR cascaded from paragraph
            // direction=rtl above. Skipping the cascade would leave Latin
            // character order inside an RTL paragraph (broken Arabic).
            // Routes through ApplyRunFormatting so schema order matches
            // direct Set path. See WordHandler.I18n.cs.
            if ((properties.TryGetValue("rtl", out var pRtl) && IsTruthy(pRtl))
                || paraRtl == true)
                ApplyRunFormatting(rProps, "rtl", "true");
            if (properties.TryGetValue("vertAlign", out var pVertAlign) || properties.TryGetValue("vertalign", out pVertAlign))
            {
                rProps.VerticalTextAlignment = new VerticalTextAlignment
                {
                    Val = pVertAlign.ToLowerInvariant() switch
                    {
                        "superscript" or "super" => VerticalPositionValues.Superscript,
                        "subscript" or "sub" => VerticalPositionValues.Subscript,
                        _ => VerticalPositionValues.Baseline
                    }
                };
            }
            if (properties.TryGetValue("superscript", out var pSup) && IsTruthy(pSup))
                rProps.VerticalTextAlignment = new VerticalTextAlignment { Val = VerticalPositionValues.Superscript };
            if (properties.TryGetValue("subscript", out var pSub) && IsTruthy(pSub))
                rProps.VerticalTextAlignment = new VerticalTextAlignment { Val = VerticalPositionValues.Subscript };
            if (properties.TryGetValue("charspacing", out var pCharSp) || properties.TryGetValue("charSpacing", out pCharSp)
                || properties.TryGetValue("letterspacing", out pCharSp) || properties.TryGetValue("letterSpacing", out pCharSp))
            {
                var csPt = pCharSp.EndsWith("pt", StringComparison.OrdinalIgnoreCase)
                    ? ParseHelpers.SafeParseDouble(pCharSp[..^2], "charspacing")
                    : ParseHelpers.SafeParseDouble(pCharSp, "charspacing");
                rProps.Spacing = new Spacing { Val = (int)Math.Round(csPt * 20, MidpointRounding.AwayFromZero) };
            }
            if (properties.TryGetValue("shd", out var pShd) || properties.TryGetValue("shading", out pShd))
            {
                var shdParts = pShd.Split(';');
                var shd = new Shading();
                if (shdParts.Length == 1)
                {
                    shd.Val = ShadingPatternValues.Clear;
                    shd.Fill = SanitizeHex(shdParts[0]);
                }
                else if (shdParts.Length >= 2)
                {
                    var rPatternPart = shdParts[0].TrimStart('#');
                    if (rPatternPart.Length >= 6 && rPatternPart.All(char.IsAsciiHexDigit))
                    {
                        Console.Error.WriteLine($"Warning: '{shdParts[0]}' looks like a color in the pattern position. Auto-swapping to: clear;{shdParts[0]}");
                        shd.Val = ShadingPatternValues.Clear;
                        shd.Fill = SanitizeHex(shdParts[0]);
                    }
                    else
                    {
                        WarnIfShadingOrderWrong(shdParts[0]); shd.Val = new ShadingPatternValues(shdParts[0]);
                        shd.Fill = SanitizeHex(shdParts[1]);
                        if (shdParts.Length >= 3) shd.Color = SanitizeHex(shdParts[2]);
                    }
                }
                rProps.Shading = shd;
            }

            foreach (var segmentedRun in BuildSegmentedRuns(text, rProps))
                para.AppendChild(segmentedRun);
        }

        // Dotted-key fallback: any "element.attr=value" prop the hand-rolled
        // blocks above did not consume goes through the same generic helper
        // wired into Set. Pre-existing dotted prefixes already handled
        // upstream (pbdr.*) are skipped to avoid double application.
        // Anything still unconsumed is recorded as silent-drop so the CLI
        // layer can surface a WARNING. CONSISTENCY(add-set-symmetry).
        var rPropsForFallback = para.Descendants<RunProperties>().FirstOrDefault();
        // Set of bare (no-dot) keys that the curated text/run block above has
        // already consumed. Anything else bare is run-level (lang, bidi,
        // kern, …) and must reach ApplyRunFormatting / TypedAttributeFallback
        // — otherwise paragraph-add silently drops them while run-level Set /
        // Add accept them, breaking add/set symmetry.
        // CONSISTENCY(add-set-symmetry).
        var bareConsumed = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "type", "text", "html", "anchor", "anchorId", "anchorid",
            "style", "styleid", "stylename",
            "align", "alignment", "direction", "dir", "bidi",
            "firstlineindent", "leftindent", "indentleft", "indent",
            "rightindent", "indentright", "hangingindent", "hanging",
            "spacebefore", "spaceafter", "linespacing", "lineSpacing",
            "keepnext", "keepwithnext", "keeplines", "keeptogether",
            "pagebreakbefore", "break",
            "numid", "numId", "ilvl", "numlevel", "numLevel",
            "liststyle", "listStyle", "start", "level", "listLevel", "listlevel",
            "outlinelevel", "outlineLevel",
            "tabs", "tabstops",
            "border", "borders", "shd", "shading",
            "font", "size", "bold", "italic", "color", "highlight",
            "underline", "strike", "strikethrough", "doublestrike", "dstrike",
            "vanish", "outline", "shadow", "emboss", "imprint", "noproof",
            "rtl", "vertAlign", "vertalign", "superscript", "subscript",
            "charspacing", "charSpacing", "letterspacing", "letterSpacing",
            "caps", "smallcaps",
            "boldcs", "italiccs", "sizecs",
            "field", "formula", "ref", "id",
        };
        foreach (var (key, value) in properties)
        {
            if (key.StartsWith("pbdr", StringComparison.OrdinalIgnoreCase)) continue;
            if (!key.Contains('.') && bareConsumed.Contains(key)) continue;
            if (!key.Contains('.'))
            {
                // Bare run-level key (lang, bidi, kern, …) — try
                // ApplyRunFormatting on the existing run rPr first, then on
                // the paragraph mark rPr (so it survives even with no text
                // run). Falls through to TypedAttributeFallback below.
                if (rPropsForFallback != null
                    && ApplyRunFormatting(rPropsForFallback, key, value)) continue;
                var bareMarkRPr = pProps.GetFirstChild<ParagraphMarkRunProperties>()
                    ?? pProps.AppendChild(new ParagraphMarkRunProperties());
                if (ApplyRunFormatting(bareMarkRPr, key, value)) continue;
                if (bareMarkRPr.ChildElements.Count == 0) bareMarkRPr.Remove();
            }
            // CONSISTENCY(font-dotted-alias): same skip-list as run-add.
            switch (key.ToLowerInvariant())
            {
                case "font.name":
                case "font.size":
                case "font.bold":
                case "font.italic":
                case "font.color":
                case "font.underline":
                case "font.strike":
                case "font.strikethrough":
                // Per-script font slots and CS toggles are already consumed
                // by the curated text/run block above; skip the typed-attr
                // fallback so they are not re-flagged as UNSUPPORTED.
                case "font.latin":
                case "font.ea":
                case "font.eastasia":
                case "font.eastasian":
                case "font.cs":
                case "font.complexscript":
                case "font.complex":
                // CS run flags (<w:bCs/> / <w:iCs/> / <w:szCs/>) — the
                // hoisted block at line 57-74 writes them to the paragraph
                // mark rPr; the dotted-fallback below would re-flag them
                // here because TypedAttributeFallback can't resolve the
                // dotted-name into the OpenXml element type.
                case "bold.cs":
                case "italic.cs":
                case "size.cs":
                case "font.bold.cs":
                case "font.italic.cs":
                case "font.size.cs":
                case "boldcs":
                case "italiccs":
                case "sizecs":
                    continue;
            }
            // CONSISTENCY(add-set-symmetry / bcp47-validation): route lang.*
            // through ApplyRunFormatting (Set's path) so the validator runs
            // on Add too. Target the existing run rPr if present, else the
            // paragraph mark rPr.
            switch (key.ToLowerInvariant())
            {
                case "lang.latin":
                case "lang.val":
                case "lang.ea":
                case "lang.eastasia":
                case "lang.eastasian":
                case "lang.cs":
                case "lang.complexscript":
                case "lang.bidi":
                {
                    if (rPropsForFallback != null
                        && ApplyRunFormatting(rPropsForFallback, key, value)) continue;
                    var langMarkRPr = pProps.GetFirstChild<ParagraphMarkRunProperties>()
                        ?? pProps.AppendChild(new ParagraphMarkRunProperties());
                    if (ApplyRunFormatting(langMarkRPr, key, value)) continue;
                    break;
                }
            }
            if (Core.TypedAttributeFallback.TrySet(pProps, key, value)) continue;
            if (rPropsForFallback != null
                && Core.TypedAttributeFallback.TrySet(rPropsForFallback, key, value)) continue;
            // No text run on this paragraph yet; route run-level attrs to
            // the paragraph mark rPr (where they apply to the paragraph
            // mark glyph + inherited by future runs).
            var paraMarkRPr = pProps.GetFirstChild<ParagraphMarkRunProperties>()
                ?? pProps.AppendChild(new ParagraphMarkRunProperties());
            if (Core.TypedAttributeFallback.TrySet(paraMarkRPr, key, value)) continue;
            if (paraMarkRPr.ChildElements.Count == 0) paraMarkRPr.Remove();
            LastAddUnsupportedProps.Add(key);
        }

        // Use ChildElements for index lookup so that tables and sectPr
        // siblings do not shift the effective insertion position. This
        // matches ResolveAnchorPosition, which computes anchor indices
        // against ChildElements.
        var allChildren = parent.ChildElements.ToList();
        if (index.HasValue && index.Value < allChildren.Count)
        {
            var refElement = allChildren[index.Value];
            parent.InsertBefore(para, refElement);
            var paraPosIdx = parent.Elements<Paragraph>().ToList().IndexOf(para) + 1;
            resultPath = $"{parentPath}/{BuildParaPathSegment(para, paraPosIdx)}";
        }
        else
        {
            AppendToParent(parent, para);
            var paraCount = parent.Elements<Paragraph>().Count();
            resultPath = $"{parentPath}/{BuildParaPathSegment(para, paraCount)}";
        }
        // R20-fuzz-11: post-insert evaluation of inherited RTL for direction=ltr.
        // Only the style-chain layer can be evaluated before insertion; the
        // enclosing section, docDefaults, and numbering lvl all need the
        // paragraph to be parented. Mirror the Set path's HasInheritedBidi
        // helper and emit <w:bidi w:val="0"/> when any layer would otherwise
        // re-inherit RTL.
        if (paraRtl == false && pProps.GetFirstChild<BiDi>() == null && HasInheritedBidi(para))
        {
            pProps.BiDi = new BiDi { Val = new DocumentFormat.OpenXml.OnOffValue(false) };
        }
        return resultPath;
    }

    private string AddEquation(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        string resultPath;
        OpenXmlElement? newElement;
        if (!properties.TryGetValue("formula", out var formula) && !properties.TryGetValue("text", out formula))
            throw new ArgumentException("'formula' (or 'text') property is required for equation type");

        var mode = properties.GetValueOrDefault("mode", "display");

        if (mode == "inline" && parent is Paragraph inlinePara)
        {
            // Insert inline math into existing paragraph
            var mathElement = FormulaParser.Parse(formula);
            if (mathElement is M.OfficeMath oMathInline)
                inlinePara.AppendChild(oMathInline);
            else
                inlinePara.AppendChild(new M.OfficeMath(mathElement.CloneNode(true)));
            var mathCount = inlinePara.Elements<M.OfficeMath>().Count();
            resultPath = $"{parentPath}/oMath[{mathCount}]";
            newElement = inlinePara;
        }
        else if (mode == "inline" && (parent is Body || parent is SdtBlock))
        {
            // Inline math under Body: wrap in a w:p (Body cannot host m:oMath directly)
            // but emit a bare m:oMath instead of m:oMathPara so the math renders as
            // inline-with-text rather than as a centered display equation.
            var mathElement = FormulaParser.Parse(formula);
            M.OfficeMath inlineOMath = mathElement is M.OfficeMath direct
                ? direct
                : new M.OfficeMath(mathElement.CloneNode(true));
            var hostPara = new Paragraph(inlineOMath);
            AssignParaId(hostPara);
            if (index.HasValue)
            {
                var children = parent.ChildElements.ToList();
                if (index.Value < children.Count)
                    parent.InsertBefore(hostPara, children[index.Value]);
                else
                    AppendToParent(parent, hostPara);
            }
            else
            {
                AppendToParent(parent, hostPara);
            }
            var pIdx = parent.Elements<Paragraph>().Count();
            resultPath = $"{parentPath}/{BuildParaPathSegment(hostPara, pIdx)}/oMath[1]";
            newElement = hostPara;
        }
        else
        {
            // Display mode: create m:oMathPara
            var mathContent = FormulaParser.Parse(formula);
            M.OfficeMath oMath;
            if (mathContent is M.OfficeMath directMath)
                oMath = directMath;
            else
                oMath = new M.OfficeMath(mathContent.CloneNode(true));

            var mathPara = new M.Paragraph(oMath);

            // Display equation must be a direct child of Body (wrapped in w:p).
            // If parent is a Paragraph, insert after that paragraph as a sibling.
            var insertTarget = parent;
            OpenXmlElement? insertAfter = null;
            if (parent is Paragraph parentPara)
            {
                insertTarget = parentPara.Parent ?? parent;
                insertAfter = parentPara;
            }

            if (insertTarget is Body || insertTarget is SdtBlock)
            {
                // Wrap m:oMathPara in w:p for schema validity
                var wrapPara = new Paragraph(mathPara);
                AssignParaId(wrapPara);

                // CONSISTENCY(rtl-cascade): inherit pPr/bidi and paragraph-mark
                // rPr/rtl from the host paragraph so the wrapper preserves the
                // surrounding RTL flow. Without this, an equation inserted
                // into an Arabic paragraph silently breaks document direction
                // (mark anchors LTR, page side flips).
                if (parent is Paragraph parentParaForBidi
                    && parentParaForBidi.ParagraphProperties is { } parentPPr)
                {
                    var parentBidi = parentPPr.GetFirstChild<BiDi>();
                    var parentMarkRtl = parentPPr.ParagraphMarkRunProperties?
                        .GetFirstChild<RightToLeftText>();
                    if (parentBidi != null || parentMarkRtl != null)
                    {
                        var wrapPPr = wrapPara.ParagraphProperties ??= new ParagraphProperties();
                        if (parentBidi != null && wrapPPr.GetFirstChild<BiDi>() == null)
                            wrapPPr.PrependChild(new BiDi());
                        if (parentMarkRtl != null)
                        {
                            var markRPr = wrapPPr.ParagraphMarkRunProperties
                                ?? wrapPPr.AppendChild(new ParagraphMarkRunProperties());
                            if (markRPr.GetFirstChild<RightToLeftText>() == null)
                                markRPr.AppendChild(new RightToLeftText());
                        }
                    }
                }
                if (insertAfter != null)
                {
                    insertTarget.InsertAfter(wrapPara, insertAfter);
                }
                else if (index.HasValue)
                {
                    var children = insertTarget.ChildElements.ToList();
                    if (index.Value < children.Count)
                        insertTarget.InsertBefore(wrapPara, children[index.Value]);
                    else
                        AppendToParent(insertTarget, wrapPara);
                }
                else
                {
                    AppendToParent(insertTarget, wrapPara);
                }
                // Compute doc-order index matching NavigateToElement's /body/oMathPara[N]
                // resolution: enumerate bare M.Paragraph and pure oMathPara wrapper w:p's.
                var oMathParaOrdinal = 0;
                var found = 0;
                foreach (var el in insertTarget.ChildElements)
                {
                    if (el is M.Paragraph)
                    {
                        oMathParaOrdinal++;
                        if (ReferenceEquals(el, mathPara)) { found = oMathParaOrdinal; break; }
                    }
                    else if (el is Paragraph wp && IsOMathParaWrapperParagraph(wp))
                    {
                        oMathParaOrdinal++;
                        if (ReferenceEquals(el, wrapPara)) { found = oMathParaOrdinal; break; }
                    }
                }
                if (found == 0) found = oMathParaOrdinal; // fallback
                var bodyPath = insertAfter != null ? parentPath.Substring(0, parentPath.LastIndexOf('/')) : parentPath;
                resultPath = $"{bodyPath}/oMathPara[{found}]";
            }
            else
            {
                AppendToParent(parent, mathPara);
                resultPath = $"{parentPath}/oMathPara[1]";
            }
            newElement = mathPara;
        }

        return resultPath;
    }

    private string AddRun(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        string resultPath;
        if (parent is not Paragraph targetPara)
            throw new ArgumentException("Runs can only be added to paragraphs");

        var newRun = new Run();
        var newRProps = new RunProperties();
        // Per-script font slots (font.latin/ea/cs) compose with bare 'font'.
        // Mirrors AddParagraph's run-creation block.
        string? nrAscii = null, nrHAnsi = null, nrEa = null, nrCs = null;
        if (properties.TryGetValue("font", out var rFont) || properties.TryGetValue("font.name", out rFont))
        { nrAscii = rFont; nrHAnsi = rFont; nrEa = rFont; }
        if (properties.TryGetValue("font.latin", out var rfLatin))
        { nrAscii = rfLatin; nrHAnsi = rfLatin; }
        if (properties.TryGetValue("font.ea", out var rfEa)
            || properties.TryGetValue("font.eastasia", out rfEa)
            || properties.TryGetValue("font.eastasian", out rfEa))
        { nrEa = rfEa; }
        if (properties.TryGetValue("font.cs", out var rfCs)
            || properties.TryGetValue("font.complexscript", out rfCs)
            || properties.TryGetValue("font.complex", out rfCs))
        { nrCs = rfCs; }
        if (nrAscii != null || nrHAnsi != null || nrEa != null || nrCs != null)
        {
            var nrFonts = new RunFonts();
            if (nrAscii != null) nrFonts.Ascii = nrAscii;
            if (nrHAnsi != null) nrFonts.HighAnsi = nrHAnsi;
            if (nrEa != null) nrFonts.EastAsia = nrEa;
            if (nrCs != null) nrFonts.ComplexScript = nrCs;
            newRProps.AppendChild(nrFonts);
        }
        if (properties.TryGetValue("size", out var rSize) || properties.TryGetValue("font.size", out rSize) || properties.TryGetValue("fontsize", out rSize))
            newRProps.AppendChild(new FontSize { Val = ((int)Math.Round(ParseFontSize(rSize) * 2, MidpointRounding.AwayFromZero)).ToString() });
        if ((properties.TryGetValue("bold", out var rBold) || properties.TryGetValue("font.bold", out rBold)) && IsTruthy(rBold))
            newRProps.Bold = new Bold();
        if ((properties.TryGetValue("bold.cs", out var rBoldCs) || properties.TryGetValue("font.bold.cs", out rBoldCs))
            && IsTruthy(rBoldCs))
            newRProps.BoldComplexScript = new BoldComplexScript();
        if ((properties.TryGetValue("italic", out var rItalic) || properties.TryGetValue("font.italic", out rItalic)) && IsTruthy(rItalic))
            newRProps.Italic = new Italic();
        if ((properties.TryGetValue("italic.cs", out var rItalicCs) || properties.TryGetValue("font.italic.cs", out rItalicCs))
            && IsTruthy(rItalicCs))
            newRProps.ItalicComplexScript = new ItalicComplexScript();
        if (properties.TryGetValue("size.cs", out var rSizeCs) || properties.TryGetValue("font.size.cs", out rSizeCs))
        {
            newRProps.FontSizeComplexScript = new FontSizeComplexScript
            {
                Val = ((int)Math.Round(ParseFontSize(rSizeCs) * 2, MidpointRounding.AwayFromZero)).ToString()
            };
        }
        if (properties.TryGetValue("color", out var rColor) || properties.TryGetValue("font.color", out rColor))
        {
            // CONSISTENCY(theme-color): Add run color accepts scheme color
            // names (accent1, dark2, hyperlink, …); same logic as
            // ApplyRunFormatting in WordHandler.Helpers.cs.
            var rSchemeName = OfficeCli.Core.ParseHelpers.NormalizeSchemeColorName(rColor);
            if (rSchemeName != null)
                newRProps.Color = new Color { Val = "auto", ThemeColor = new EnumValue<ThemeColorValues>(new ThemeColorValues(rSchemeName)) };
            else
                newRProps.Color = new Color { Val = SanitizeHex(rColor) };
        }
        if (properties.TryGetValue("underline", out var rUnderline) || properties.TryGetValue("font.underline", out rUnderline))
        {
            var ulVal = NormalizeUnderlineValue(rUnderline);
            newRProps.Underline = new Underline { Val = new UnderlineValues(ulVal) };
        }
        if ((properties.TryGetValue("strike", out var rStrike)
                || properties.TryGetValue("strikethrough", out rStrike)
                || properties.TryGetValue("font.strike", out rStrike)
                || properties.TryGetValue("font.strikethrough", out rStrike))
            && IsTruthy(rStrike))
            newRProps.Strike = new Strike();
        if (properties.TryGetValue("highlight", out var rHighlight))
            newRProps.Highlight = new Highlight { Val = ParseHighlightColor(rHighlight) };
        if ((properties.TryGetValue("caps", out var rCaps)
                || properties.TryGetValue("allcaps", out rCaps)
                || properties.TryGetValue("allCaps", out rCaps))
            && IsTruthy(rCaps))
            newRProps.Caps = new Caps();
        if (properties.TryGetValue("smallcaps", out var rSmallCaps) || properties.TryGetValue("smallCaps", out rSmallCaps))
        {
            if (IsTruthy(rSmallCaps)) newRProps.SmallCaps = new SmallCaps();
        }
        if (properties.TryGetValue("dstrike", out var rDstrike) && IsTruthy(rDstrike))
            newRProps.DoubleStrike = new DoubleStrike();
        if (properties.TryGetValue("vanish", out var rVanish) && IsTruthy(rVanish))
            newRProps.Vanish = new Vanish();
        if (properties.TryGetValue("outline", out var rOutline) && IsTruthy(rOutline))
            newRProps.Outline = new Outline();
        if (properties.TryGetValue("shadow", out var rShadow) && IsTruthy(rShadow))
            newRProps.Shadow = new Shadow();
        if (properties.TryGetValue("emboss", out var rEmboss) && IsTruthy(rEmboss))
            newRProps.Emboss = new Emboss();
        if (properties.TryGetValue("imprint", out var rImprint) && IsTruthy(rImprint))
            newRProps.Imprint = new Imprint();
        if (properties.TryGetValue("noproof", out var rNoProof) && IsTruthy(rNoProof))
            newRProps.NoProof = new NoProof();
        if (properties.TryGetValue("rtl", out var rRtl) && IsTruthy(rRtl))
            ApplyRunFormatting(newRProps, "rtl", "true");
        // CONSISTENCY(canonical-key): accept "direction"=rtl|ltr as the
        // canonical alias for run-level rtl, matching paragraph/section
        // input vocabulary and the symmetric Get readback (R16-bt-1).
        else if (properties.TryGetValue("direction", out var rDir)
            || properties.TryGetValue("dir", out rDir))
        {
            var v = rDir?.Trim().ToLowerInvariant();
            if (v == "rtl") ApplyRunFormatting(newRProps, "rtl", "true");
            else if (v == "ltr") ApplyRunFormatting(newRProps, "rtl", "false");
        }
        if (properties.TryGetValue("vertAlign", out var rVertAlign) || properties.TryGetValue("vertalign", out rVertAlign))
        {
            newRProps.VerticalTextAlignment = new VerticalTextAlignment
            {
                Val = rVertAlign.ToLowerInvariant() switch
                {
                    "superscript" or "super" => VerticalPositionValues.Superscript,
                    "subscript" or "sub" => VerticalPositionValues.Subscript,
                    _ => VerticalPositionValues.Baseline
                }
            };
        }
        if (properties.TryGetValue("superscript", out var rSup) && IsTruthy(rSup))
            newRProps.VerticalTextAlignment = new VerticalTextAlignment { Val = VerticalPositionValues.Superscript };
        if (properties.TryGetValue("subscript", out var rSub) && IsTruthy(rSub))
            newRProps.VerticalTextAlignment = new VerticalTextAlignment { Val = VerticalPositionValues.Subscript };
        if (properties.TryGetValue("charspacing", out var rCharSp) || properties.TryGetValue("charSpacing", out rCharSp)
            || properties.TryGetValue("letterspacing", out rCharSp) || properties.TryGetValue("letterSpacing", out rCharSp))
        {
            var csPt = rCharSp.EndsWith("pt", StringComparison.OrdinalIgnoreCase)
                ? ParseHelpers.SafeParseDouble(rCharSp[..^2], "charspacing")
                : ParseHelpers.SafeParseDouble(rCharSp, "charspacing");
            newRProps.Spacing = new Spacing { Val = (int)Math.Round(csPt * 20, MidpointRounding.AwayFromZero) };
        }
        if (properties.TryGetValue("shd", out var rShd) || properties.TryGetValue("shading", out rShd))
        {
            var shdParts = rShd.Split(';');
            var shd = new Shading();
            if (shdParts.Length == 1)
            {
                shd.Val = ShadingPatternValues.Clear;
                shd.Fill = SanitizeHex(shdParts[0]);
            }
            else if (shdParts.Length >= 2)
            {
                var addRunPatternPart = shdParts[0].TrimStart('#');
                if (addRunPatternPart.Length >= 6 && addRunPatternPart.All(char.IsAsciiHexDigit))
                {
                    Console.Error.WriteLine($"Warning: '{shdParts[0]}' looks like a color in the pattern position. Auto-swapping to: clear;{shdParts[0]}");
                    shd.Val = ShadingPatternValues.Clear;
                    shd.Fill = SanitizeHex(shdParts[0]);
                }
                else
                {
                    WarnIfShadingOrderWrong(shdParts[0]); shd.Val = new ShadingPatternValues(shdParts[0]);
                    shd.Fill = SanitizeHex(shdParts[1]);
                    if (shdParts.Length >= 3) shd.Color = SanitizeHex(shdParts[2]);
                }
            }
            newRProps.Shading = shd;
        }

        // w14 text effects
        var tempRun = new Run();
        tempRun.PrependChild(newRProps);
        if (properties.TryGetValue("textOutline", out var toVal) || properties.TryGetValue("textoutline", out toVal))
            ApplyW14TextEffect(tempRun, "textOutline", toVal, BuildW14TextOutline);
        if (properties.TryGetValue("textFill", out var tfVal) || properties.TryGetValue("textfill", out tfVal))
            ApplyW14TextEffect(tempRun, "textFill", tfVal, BuildW14TextFill);
        if (properties.TryGetValue("w14shadow", out var w14sVal))
            ApplyW14TextEffect(tempRun, "shadow", w14sVal, BuildW14Shadow);
        if (properties.TryGetValue("w14glow", out var w14gVal))
            ApplyW14TextEffect(tempRun, "glow", w14gVal, BuildW14Glow);
        if (properties.TryGetValue("w14reflection", out var w14rVal))
            ApplyW14TextEffect(tempRun, "reflection", w14rVal, BuildW14Reflection);
        // Detach rPr from temp run for re-attachment to actual run
        newRProps.Remove();

        // Inherit default formatting from paragraph mark run properties
        var markRProps = targetPara.ParagraphProperties?.ParagraphMarkRunProperties;
        if (markRProps != null)
        {
            foreach (var child in markRProps.ChildElements)
            {
                var childType = child.GetType();
                if (newRProps.Elements().All(e => e.GetType() != childType))
                    newRProps.AppendChild(child.CloneNode(true));
            }
        }

        var runText = properties.GetValueOrDefault("text", "");
        newRun.AppendChild(newRProps);
        AppendTextWithBreaks(newRun, runText);

        // Dotted-key fallback: same generic helper as Set's run path.
        // Anything still unconsumed after the hand-rolled blocks above
        // gets routed through TypedAttributeFallback; failures land in
        // LastAddUnsupportedProps so the CLI surfaces a WARNING instead
        // of silently dropping. CONSISTENCY(add-set-symmetry).
        foreach (var (key, value) in properties)
        {
            if (!key.Contains('.')) continue;
            // CONSISTENCY(font-dotted-alias): font.name/font.bold/font.size/
            // font.italic/font.color/font.underline/font.strike are consumed
            // above by the curated alias blocks; skip the typed-attr fallback
            // so they don't get re-flagged as UNSUPPORTED.
            switch (key.ToLowerInvariant())
            {
                case "font.name":
                case "font.size":
                case "font.bold":
                case "font.italic":
                case "font.color":
                case "font.underline":
                case "font.strike":
                case "font.strikethrough":
                // Per-script slots and CS toggles already consumed above.
                case "font.latin":
                case "font.ea":
                case "font.eastasia":
                case "font.eastasian":
                case "font.cs":
                case "font.complexscript":
                case "font.complex":
                // CS run flags (<w:bCs/> / <w:iCs/> / <w:szCs/>) — the
                // run-add block above writes them through ApplyRunFormatting;
                // dotted-fallback can't resolve the dotted name into the
                // OpenXml element type.
                case "bold.cs":
                case "italic.cs":
                case "size.cs":
                case "font.bold.cs":
                case "font.italic.cs":
                case "font.size.cs":
                case "boldcs":
                case "italiccs":
                case "sizecs":
                    continue;
            }
            // CONSISTENCY(add-set-symmetry / bcp47-validation): route lang.*
            // through ApplyRunFormatting so the BCP-47 validator that Set
            // applies also runs on Add (without this, malformed lang values
            // like "-" silently became <w:lang w:val="-"/>).
            switch (key.ToLowerInvariant())
            {
                case "lang.latin":
                case "lang.val":
                case "lang.ea":
                case "lang.eastasia":
                case "lang.eastasian":
                case "lang.cs":
                case "lang.complexscript":
                case "lang.bidi":
                    if (ApplyRunFormatting(newRProps, key, value)) continue;
                    break;
            }
            if (Core.TypedAttributeFallback.TrySet(newRProps, key, value)) continue;
            LastAddUnsupportedProps.Add(key);
        }

        // Use ChildElements for index lookup so ResolveAnchorPosition's
        // childElement-indexed result lines up. If index points at
        // ParagraphProperties, clamp forward so pPr stays first.
        var allChildren = targetPara.ChildElements.ToList();
        if (index.HasValue && index.Value < allChildren.Count)
        {
            var refElement = allChildren[index.Value];
            if (refElement is ParagraphProperties)
            {
                // insert after pPr — i.e. before whatever sits at index+1, else append
                if (index.Value + 1 < allChildren.Count)
                    targetPara.InsertBefore(newRun, allChildren[index.Value + 1]);
                else
                    targetPara.AppendChild(newRun);
            }
            else
            {
                targetPara.InsertBefore(newRun, refElement);
            }
            // CONSISTENCY(run-path-index): match navigation's r[N] enumeration
            // (Descendants<Run>() minus comment-reference runs) via GetAllRuns.
            var runPosIdx = GetAllRuns(targetPara).IndexOf(newRun) + 1;
            // CONSISTENCY(para-path-canonical): canonicalize to paraId-form.
            resultPath = $"{ReplaceTrailingParaSegment(parentPath, targetPara)}/r[{runPosIdx}]";
        }
        else
        {
            targetPara.AppendChild(newRun);
            var runCount = GetAllRuns(targetPara).IndexOf(newRun) + 1;
            resultPath = $"{ReplaceTrailingParaSegment(parentPath, targetPara)}/r[{runCount}]";
        }

        // Refresh textId since paragraph content changed
        targetPara.TextId = GenerateParaId();

        return resultPath;
    }

    /// <summary>
    /// Append <paramref name="text"/> to <paramref name="run"/>, tokenizing on
    /// '\n' (w:br) and '\t' (w:tab) so the user-visible line breaks and tabs
    /// round-trip through Word instead of being collapsed to a single space.
    /// CRLF/CR are normalized to LF first.
    /// </summary>
    internal static void AppendTextWithBreaks(Run run, string text)
    {
        if (string.IsNullOrEmpty(text))
        {
            run.AppendChild(new Text("") { Space = SpaceProcessingModeValues.Preserve });
            return;
        }
        // CONSISTENCY(xml-text-validation): mirror Set's text= path — reject XML 1.0
        // illegal control chars before constructing Text nodes. Without this, the
        // resident process saves a corrupt DOM and surfaces "save failed — data may
        // be lost" only on close, costing the user their edits.
        Core.ParseHelpers.ValidateXmlText(text, "text");
        // CONSISTENCY(escape-sequences): cross-handler convention — `\n` / `\t`
        // two-char escapes in --prop text= are interpreted as real newline /
        // tab. Mirrors PPTX shape-text and Excel cell-value handling. CRLF/CR
        // collapsed afterwards so all break forms route through <w:br/>.
        var s = text.Replace("\\n", "\n").Replace("\\t", "\t");
        s = s.Replace("\r\n", "\n").Replace("\r", "\n");
        int start = 0;
        for (int i = 0; i < s.Length; i++)
        {
            char c = s[i];
            if (c == '\n' || c == '\t')
            {
                if (i > start)
                    run.AppendChild(new Text(s.Substring(start, i - start)) { Space = SpaceProcessingModeValues.Preserve });
                if (c == '\n') run.AppendChild(new Break());
                else run.AppendChild(new TabChar());
                start = i + 1;
            }
        }
        if (start < s.Length)
            run.AppendChild(new Text(s.Substring(start)) { Space = SpaceProcessingModeValues.Preserve });
        else if (start == 0)
            run.AppendChild(new Text("") { Space = SpaceProcessingModeValues.Preserve });
    }

    // Add a tab stop. Parent must be a Paragraph or a paragraph/table-typed
    // Style; the helper finds or creates the pPr/Tabs container and appends
    // a TabStop. `pos` is required (twips, or any unit accepted by
    // SpacingConverter.ParseWordSpacing). `val` defaults to "left";
    // `leader` is optional. Returns the new tab's path under the
    // conventional /<parent>/tab[N] form — Navigation descends through
    // pPr/tabs (paragraph) or StyleParagraphProperties/tabs (style)
    // transparently for this segment shape.
    private string AddTab(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        if (!properties.TryGetValue("pos", out var posStr) || string.IsNullOrWhiteSpace(posStr))
            throw new ArgumentException("tab requires 'pos' property (e.g. --prop pos=9360 or --prop pos=6cm)");

        var posTwips = (int)SpacingConverter.ParseWordSpacing(posStr);

        var tabStop = new TabStop { Position = posTwips };
        if (properties.TryGetValue("val", out var valStr) && !string.IsNullOrEmpty(valStr))
        {
            var tabValNorm = valStr.ToLowerInvariant();
            // Validate before constructing the enum — an invalid string throws
            // ArgumentOutOfRangeException which the outer dispatcher catches and
            // surfaces as a misleading "Invalid index or anchor" error.
            var knownTabVals = new[] { "left", "center", "right", "decimal", "bar", "clear", "num", "start", "end" };
            if (!knownTabVals.Contains(tabValNorm))
                throw new ArgumentException($"Invalid tab val '{valStr}'. Valid: {string.Join(", ", knownTabVals)}.");
            tabStop.Val = new EnumValue<TabStopValues>(new TabStopValues(tabValNorm));
        }
        else
            tabStop.Val = TabStopValues.Left;
        if (properties.TryGetValue("leader", out var leaderStr) && !string.IsNullOrEmpty(leaderStr))
        {
            var leaderNorm = leaderStr.ToLowerInvariant();
            var knownLeaders = new[] { "none", "dot", "heavy", "hyphen", "middledot", "underscore" };
            if (!knownLeaders.Contains(leaderNorm))
                throw new ArgumentException($"Invalid tab leader '{leaderStr}'. Valid: {string.Join(", ", knownLeaders)}.");
            tabStop.Leader = new EnumValue<TabStopLeaderCharValues>(new TabStopLeaderCharValues(leaderNorm));
        }

        // pPr children have schema order; Tabs sits early. PrependChild
        // is conservative — Word accepts Tabs at the start of pPr and
        // we don't want to interleave with later siblings (numPr, ind, ...)
        // that have stricter ordering constraints.
        Tabs tabs;
        if (parent is Paragraph para)
        {
            // pPr must come first inside <w:p> per CT_P schema
            var pProps = para.ParagraphProperties ?? para.PrependChild(new ParagraphProperties());
            tabs = pProps.GetFirstChild<Tabs>() ?? pProps.PrependChild(new Tabs());
        }
        else if (parent is Style style)
        {
            // Type guard already enforced in Add.cs (paragraph/table only).
            // EnsureStyleParagraphProperties handles schema-correct insertion
            // before StyleRunProperties.
            var spProps = style.StyleParagraphProperties ?? EnsureStyleParagraphProperties(style);
            tabs = spProps.GetFirstChild<Tabs>() ?? spProps.PrependChild(new Tabs());
        }
        else
        {
            throw new ArgumentException(
                $"Cannot add 'tab' under {parentPath}: tab stops belong inside a paragraph or a paragraph-typed style.");
        }

        var existing = tabs.Elements<TabStop>().ToList();
        if (index.HasValue && index.Value >= 0 && index.Value < existing.Count)
            tabs.InsertBefore(tabStop, existing[index.Value]);
        else
            tabs.AppendChild(tabStop);

        var newIdx = tabs.Elements<TabStop>().ToList().IndexOf(tabStop) + 1;
        return $"{parentPath}/tab[{newIdx}]";
    }

    // CONSISTENCY(run-special-content): inline `<w:ptab>` (positional tab,
    // Word 2007+) wrapped in `<w:r>`. Used in headers/footers to anchor
    // left/center/right alignment regions. Mirrors AddBreak's "wrap an
    // inline structure in a Run, insert into paragraph" pattern.
    private string AddPtab(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        // Validate parent first (more fundamental than property contents) so
        // a misrouted call surfaces the real failure ("must be a paragraph")
        // instead of pushing the user through alignment/leader/relativeTo
        // diagnostics that wouldn't matter at the right path.
        if (parent is not Paragraph para)
            throw new ArgumentException("ptab parent must be a paragraph (got " + parent.GetType().Name + ").");

        if (!(properties.TryGetValue("align", out var alignment) || properties.TryGetValue("alignment", out alignment)) || string.IsNullOrWhiteSpace(alignment))
            throw new ArgumentException("ptab requires 'alignment' property (left, center, or right).");

        var ptab = new PositionalTab { Alignment = ParsePtabAlignment(alignment) };
        // CONSISTENCY(empty-prop-as-default): three optional ptab props use
        // matching IsNullOrWhiteSpace guards so empty-string is uniformly
        // treated as "unset / use default" — previously relativeTo passed
        // "" straight to ParsePtabRelativeTo, raising "Invalid relativeTo
        // ''" while leader silently defaulted, an asymmetry that bit
        // scripted callers building param dicts.
        if ((properties.TryGetValue("relativeTo", out var relTo)
             || properties.TryGetValue("relativeto", out relTo))
            && !string.IsNullOrWhiteSpace(relTo))
            ptab.RelativeTo = ParsePtabRelativeTo(relTo);
        else
            ptab.RelativeTo = AbsolutePositionTabPositioningBaseValues.Margin;
        if (properties.TryGetValue("leader", out var leader) && !string.IsNullOrWhiteSpace(leader))
            ptab.Leader = ParsePtabLeader(leader);
        else
            ptab.Leader = AbsolutePositionTabLeaderCharValues.None;

        var ptabRun = new Run(ptab);
        InsertIntoParagraph(para, ptabRun, index);
        // CONSISTENCY(paraid-textid-refresh): paragraph contents changed,
        // so textId must regenerate to mark the paragraph as modified for
        // revision-tracking and diff tooling. Mirrors AddRun's behavior.
        para.TextId = GenerateParaId();
        var runIdx = GetAllRuns(para).IndexOf(ptabRun) + 1;
        // CONSISTENCY(para-path-canonical): when parent is itself a
        // paragraph, parentPath already points at it — appending another
        // /p[N] would yield an illegal /p[1]/p[1]/r[N] path. Replace the
        // trailing /p[...] segment with paraId-form so the returned
        // path round-trips through Get unchanged.
        var canonicalParaPath = ReplaceTrailingParaSegment(parentPath, para);
        return $"{canonicalParaPath}/r[{runIdx}]";
    }
}
