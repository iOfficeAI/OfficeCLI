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

        if (properties.TryGetValue("style", out var style))
            pProps.ParagraphStyleId = new ParagraphStyleId { Val = style };
        if (properties.TryGetValue("alignment", out var alignment) || properties.TryGetValue("align", out alignment))
            pProps.Justification = new Justification { Val = ParseJustification(alignment) };
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
        if (properties.TryGetValue("numid", out var numId) || properties.TryGetValue("numId", out numId))
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
            || properties.TryGetValue("ilvl", out numLevel))
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
        if (properties.TryGetValue("widowcontrol", out var addWC) || properties.TryGetValue("widowControl", out addWC))
        {
            if (IsTruthy(addWC))
                pProps.WidowControl = new WidowControl();
            else
                pProps.WidowControl = new WidowControl { Val = false };
        }
        // CONSISTENCY(add-set-symmetry): Set supports contextualSpacing (WordHandler.Set.cs:529);
        // Add must accept the same prop so the "Add then Get" lifecycle test pattern works
        // without falling back to a separate Set call. Mirrors widowControl above.
        if (properties.TryGetValue("contextualspacing", out var addCS) || properties.TryGetValue("contextualSpacing", out addCS))
        {
            if (IsTruthy(addCS))
                pProps.ContextualSpacing = new ContextualSpacing();
            else
                pProps.ContextualSpacing = new ContextualSpacing { Val = false };
        }
        foreach (var (pk, pv) in properties)
        {
            if (pk.StartsWith("pbdr", StringComparison.OrdinalIgnoreCase))
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
                levelVal = ParseHelpers.SafeParseInt(ll, "listLevel");
            ApplyListStyle(para, listStyle, startVal, levelVal);
            // pProps already appended, skip the append below
            goto paragraphPropsApplied;
        }

        para.AppendChild(pProps);
        paragraphPropsApplied:

        if (properties.TryGetValue("text", out var text))
        {
            var run = new Run();
            var rProps = new RunProperties();
            if (properties.TryGetValue("font", out var font))
            {
                rProps.AppendChild(new RunFonts { Ascii = font, HighAnsi = font, EastAsia = font });
            }
            if (properties.TryGetValue("size", out var size))
            {
                rProps.AppendChild(new FontSize { Val = ((int)Math.Round(ParseFontSize(size) * 2, MidpointRounding.AwayFromZero)).ToString() });
            }
            if (properties.TryGetValue("bold", out var bold) && IsTruthy(bold))
                rProps.Bold = new Bold();
            if (properties.TryGetValue("italic", out var pItalic) && IsTruthy(pItalic))
                rProps.Italic = new Italic();
            if (properties.TryGetValue("color", out var pColor))
                rProps.Color = new Color { Val = SanitizeHex(pColor) };
            if (properties.TryGetValue("underline", out var pUnderline))
            {
                var ulVal = NormalizeUnderlineValue(pUnderline);
                rProps.Underline = new Underline { Val = new UnderlineValues(ulVal) };
            }
            if ((properties.TryGetValue("strike", out var pStrike) || properties.TryGetValue("strikethrough", out pStrike)) && IsTruthy(pStrike))
                rProps.Strike = new Strike();
            if (properties.TryGetValue("highlight", out var pHighlight))
                rProps.Highlight = new Highlight { Val = ParseHighlightColor(pHighlight) };
            if (properties.TryGetValue("caps", out var pCaps) && IsTruthy(pCaps))
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
            if (properties.TryGetValue("rtl", out var pRtl) && IsTruthy(pRtl))
                rProps.RightToLeftText = new RightToLeftText();
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

            run.AppendChild(rProps);
            AppendTextWithBreaks(run, text);
            para.AppendChild(run);
        }

        // Dotted-key fallback: any "element.attr=value" prop the hand-rolled
        // blocks above did not consume goes through the same generic helper
        // wired into Set. Pre-existing dotted prefixes already handled
        // upstream (pbdr.*) are skipped to avoid double application.
        // Anything still unconsumed is recorded as silent-drop so the CLI
        // layer can surface a WARNING. CONSISTENCY(add-set-symmetry).
        var rPropsForFallback = para.Descendants<RunProperties>().FirstOrDefault();
        foreach (var (key, value) in properties)
        {
            if (!key.Contains('.')) continue;
            if (key.StartsWith("pbdr", StringComparison.OrdinalIgnoreCase)) continue;
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
        if (properties.TryGetValue("font", out var rFont))
            newRProps.AppendChild(new RunFonts { Ascii = rFont, HighAnsi = rFont, EastAsia = rFont });
        if (properties.TryGetValue("size", out var rSize))
            newRProps.AppendChild(new FontSize { Val = ((int)Math.Round(ParseFontSize(rSize) * 2, MidpointRounding.AwayFromZero)).ToString() });
        if (properties.TryGetValue("bold", out var rBold) && IsTruthy(rBold))
            newRProps.Bold = new Bold();
        if (properties.TryGetValue("italic", out var rItalic) && IsTruthy(rItalic))
            newRProps.Italic = new Italic();
        if (properties.TryGetValue("color", out var rColor))
            newRProps.Color = new Color { Val = SanitizeHex(rColor) };
        if (properties.TryGetValue("underline", out var rUnderline))
        {
            var ulVal = NormalizeUnderlineValue(rUnderline);
            newRProps.Underline = new Underline { Val = new UnderlineValues(ulVal) };
        }
        if ((properties.TryGetValue("strike", out var rStrike) || properties.TryGetValue("strikethrough", out rStrike)) && IsTruthy(rStrike))
            newRProps.Strike = new Strike();
        if (properties.TryGetValue("highlight", out var rHighlight))
            newRProps.Highlight = new Highlight { Val = ParseHighlightColor(rHighlight) };
        if (properties.TryGetValue("caps", out var rCaps) && IsTruthy(rCaps))
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
            newRProps.RightToLeftText = new RightToLeftText();
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

        newRun.AppendChild(newRProps);
        var runText = properties.GetValueOrDefault("text", "");
        AppendTextWithBreaks(newRun, runText);

        // Dotted-key fallback: same generic helper as Set's run path.
        // Anything still unconsumed after the hand-rolled blocks above
        // gets routed through TypedAttributeFallback; failures land in
        // LastAddUnsupportedProps so the CLI surfaces a WARNING instead
        // of silently dropping. CONSISTENCY(add-set-symmetry).
        foreach (var (key, value) in properties)
        {
            if (!key.Contains('.')) continue;
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
            var runPosIdx = targetPara.Elements<Run>().ToList().IndexOf(newRun) + 1;
            resultPath = $"{parentPath}/r[{runPosIdx}]";
        }
        else
        {
            targetPara.AppendChild(newRun);
            var runCount = targetPara.Elements<Run>().Count();
            resultPath = $"{parentPath}/r[{runCount}]";
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
        var s = text.Replace("\r\n", "\n").Replace("\r", "\n");
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
}
