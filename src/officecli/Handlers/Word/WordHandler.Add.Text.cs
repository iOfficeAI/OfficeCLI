// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    private string AddParagraph(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        string resultPath;
        var para = new Paragraph();
        var pProps = new ParagraphProperties();

        if (properties.TryGetValue("style", out var style))
            pProps.ParagraphStyleId = new ParagraphStyleId { Val = style };
        if (properties.TryGetValue("alignment", out var alignment) || properties.TryGetValue("align", out alignment))
            pProps.Justification = new Justification { Val = ParseJustification(alignment) };
        if (properties.TryGetValue("firstlineindent", out var indent) || properties.TryGetValue("firstLineIndent", out indent))
        {
            // Validate range — OOXML stores as StringValue but must fit within reasonable twip range
            if (long.TryParse(indent, out var indentLong) && (indentLong < 0 || indentLong > 31680))
                throw new OverflowException($"First line indent value out of range (0-31680 twips): {indent}");
            pProps.Indentation = new Indentation
            {
                FirstLine = indent  // raw twips, consistent with Set and Get
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
        if (properties.TryGetValue("numid", out var numId))
        {
            var numPr = pProps.NumberingProperties ?? (pProps.NumberingProperties = new NumberingProperties());
            numPr.NumberingId = new NumberingId { Val = ParseHelpers.SafeParseInt(numId, "numid") };
            if (properties.TryGetValue("numlevel", out var numLevel))
            {
                numPr.NumberingLevelReference = new NumberingLevelReference { Val = ParseHelpers.SafeParseInt(numLevel, "numlevel") };
            }
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
                WarnIfShadingOrderWrong(shdParts[0]); shd.Val = new ShadingPatternValues(shdParts[0]);
                shd.Fill = SanitizeHex(shdParts[1]);
                if (shdParts.Length >= 3) shd.Color = SanitizeHex(shdParts[2]);
            }
            pProps.Shading = shd;
        }
        if (properties.TryGetValue("leftindent", out var addLI) || properties.TryGetValue("leftIndent", out addLI) || properties.TryGetValue("indentleft", out addLI) || properties.TryGetValue("indent", out addLI))
        {
            var ind = pProps.Indentation ?? (pProps.Indentation = new Indentation());
            ind.Left = ParseHelpers.SafeParseUint(addLI, "leftindent").ToString();
        }
        if (properties.TryGetValue("rightindent", out var addRI) || properties.TryGetValue("rightIndent", out addRI) || properties.TryGetValue("indentright", out addRI))
        {
            var ind = pProps.Indentation ?? (pProps.Indentation = new Indentation());
            ind.Right = ParseHelpers.SafeParseUint(addRI, "rightindent").ToString();
        }
        if (properties.TryGetValue("hangingindent", out var addHI) || properties.TryGetValue("hangingIndent", out addHI) || properties.TryGetValue("hanging", out addHI))
        {
            var ind = pProps.Indentation ?? (pProps.Indentation = new Indentation());
            ind.Hanging = ParseHelpers.SafeParseUint(addHI, "hangingindent").ToString();
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
            if (properties.TryGetValue("listLevel", out var ll) || properties.TryGetValue("listlevel", out ll) || properties.TryGetValue("level", out ll))
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
                var ulVal = pUnderline.ToLowerInvariant() switch { "true" => "single", "false" or "none" => "none", _ => pUnderline };
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
                    WarnIfShadingOrderWrong(shdParts[0]); shd.Val = new ShadingPatternValues(shdParts[0]);
                    shd.Fill = SanitizeHex(shdParts[1]);
                    if (shdParts.Length >= 3) shd.Color = SanitizeHex(shdParts[2]);
                }
                rProps.Shading = shd;
            }

            run.AppendChild(rProps);
            run.AppendChild(new Text { Text = text, Space = SpaceProcessingModeValues.Preserve });
            para.AppendChild(run);
        }

        var paraCount = parent.Elements<Paragraph>().Count();
        if (index.HasValue && index.Value < paraCount)
        {
            var refElement = parent.Elements<Paragraph>().ElementAt(index.Value);
            parent.InsertBefore(para, refElement);
            resultPath = $"{parentPath}/p[{index.Value + 1}]";
        }
        else
        {
            AppendToParent(parent, para);
            resultPath = $"{parentPath}/p[{paraCount + 1}]";
        }
        return resultPath;
    }

    private string AddEquation(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        string resultPath;
        OpenXmlElement? newElement;
        if (!properties.TryGetValue("formula", out var formula))
            throw new ArgumentException("'formula' property is required for equation type");

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

            if (parent is Body || parent is SdtBlock)
            {
                // Wrap m:oMathPara in w:p for schema validity
                var wrapPara = new Paragraph(mathPara);
                var mathParaCount = parent.Descendants<M.Paragraph>().Count();
                if (index.HasValue)
                {
                    var children = parent.ChildElements.ToList();
                    if (index.Value < children.Count)
                        parent.InsertBefore(wrapPara, children[index.Value]);
                    else
                        AppendToParent(parent, wrapPara);
                }
                else
                {
                    AppendToParent(parent, wrapPara);
                }
                resultPath = $"{parentPath}/oMathPara[{mathParaCount + 1}]";
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
            var ulVal = rUnderline.ToLowerInvariant() switch { "true" => "single", "false" or "none" => "none", _ => rUnderline };
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
                WarnIfShadingOrderWrong(shdParts[0]); shd.Val = new ShadingPatternValues(shdParts[0]);
                shd.Fill = SanitizeHex(shdParts[1]);
                if (shdParts.Length >= 3) shd.Color = SanitizeHex(shdParts[2]);
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
        newRun.AppendChild(new Text { Text = runText, Space = SpaceProcessingModeValues.Preserve });

        var runCount = targetPara.Elements<Run>().Count();
        if (index.HasValue && index.Value < runCount)
        {
            var refRun = targetPara.Elements<Run>().ElementAt(index.Value);
            targetPara.InsertBefore(newRun, refRun);
            resultPath = $"{parentPath}/r[{index.Value + 1}]";
        }
        else
        {
            targetPara.AppendChild(newRun);
            resultPath = $"{parentPath}/r[{runCount + 1}]";
        }

        return resultPath;
    }
}
