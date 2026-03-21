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
    public string Add(string parentPath, string type, int? index, Dictionary<string, string> properties)
    {
        var body = _doc.MainDocumentPart?.Document?.Body
            ?? throw new InvalidOperationException("Document body not found");

        OpenXmlElement parent;
        if (parentPath is "/" or "" or "/body")
        {
            parent = body;
            parentPath = "/body"; // Normalize so result paths are usable (e.g. /body/p[1] not //p[1])
        }
        else if (parentPath == "/styles")
        {
            // Ensure styles part exists for style operations
            var stylesPart = _doc.MainDocumentPart!.StyleDefinitionsPart
                ?? _doc.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles ??= new Styles();
            parent = stylesPart.Styles;
        }
        else
        {
            var parts = ParsePath(parentPath);
            parent = NavigateToElement(parts)
                ?? throw new ArgumentException($"Path not found: {parentPath}");
        }

        OpenXmlElement newElement;
        string resultPath;

        switch (type.ToLowerInvariant())
        {
            case "paragraph" or "p":
                var para = new Paragraph();
                var pProps = new ParagraphProperties();

                if (properties.TryGetValue("style", out var style))
                    pProps.ParagraphStyleId = new ParagraphStyleId { Val = style };
                if (properties.TryGetValue("alignment", out var alignment) || properties.TryGetValue("align", out alignment))
                    pProps.Justification = new Justification { Val = ParseJustification(alignment) };
                if (properties.TryGetValue("firstlineindent", out var indent))
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
                if (properties.TryGetValue("leftindent", out var addLI) || properties.TryGetValue("indentleft", out addLI))
                {
                    var ind = pProps.Indentation ?? (pProps.Indentation = new Indentation());
                    ind.Left = ParseHelpers.SafeParseUint(addLI, "leftindent").ToString();
                }
                if (properties.TryGetValue("rightindent", out var addRI) || properties.TryGetValue("indentright", out addRI))
                {
                    var ind = pProps.Indentation ?? (pProps.Indentation = new Indentation());
                    ind.Right = ParseHelpers.SafeParseUint(addRI, "rightindent").ToString();
                }
                if (properties.TryGetValue("hangingindent", out var addHI) || properties.TryGetValue("hanging", out addHI))
                {
                    var ind = pProps.Indentation ?? (pProps.Indentation = new Indentation());
                    ind.Hanging = ParseHelpers.SafeParseUint(addHI, "hangingindent").ToString();
                    ind.FirstLine = null;
                }
                // firstlineindent already handled above (line ~66-74) with × 480 conversion
                if (properties.TryGetValue("keepnext", out var addKN) && IsTruthy(addKN))
                    pProps.KeepNext = new KeepNext();
                if ((properties.TryGetValue("keeplines", out var addKL) || properties.TryGetValue("keeptogether", out addKL)) && IsTruthy(addKL))
                    pProps.KeepLines = new KeepLines();
                if (properties.TryGetValue("pagebreakbefore", out var addPBB) && IsTruthy(addPBB))
                    pProps.PageBreakBefore = new PageBreakBefore();
                if (properties.TryGetValue("widowcontrol", out var addWC) && IsTruthy(addWC))
                    pProps.WidowControl = new WidowControl();
                foreach (var (pk, pv) in properties)
                {
                    if (pk.StartsWith("pbdr", StringComparison.OrdinalIgnoreCase))
                        ApplyParagraphBorders(pProps, pk, pv);
                }
                if (properties.TryGetValue("liststyle", out var listStyle))
                {
                    para.AppendChild(pProps);
                    int? startVal = null;
                    if (properties.TryGetValue("start", out var sv))
                        startVal = ParseHelpers.SafeParseInt(sv, "start");
                    ApplyListStyle(para, listStyle, startVal);
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
                        rProps.AppendChild(new FontSize { Val = ((int)(ParseFontSize(size) * 2)).ToString() });
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
                    if (properties.TryGetValue("strike", out var pStrike) && IsTruthy(pStrike))
                        rProps.Strike = new Strike();
                    if (properties.TryGetValue("highlight", out var pHighlight))
                        rProps.Highlight = new Highlight { Val = ParseHighlightColor(pHighlight) };
                    if (properties.TryGetValue("caps", out var pCaps) && IsTruthy(pCaps))
                        rProps.Caps = new Caps();
                    if (properties.TryGetValue("smallcaps", out var pSmallCaps) && IsTruthy(pSmallCaps))
                        rProps.SmallCaps = new SmallCaps();
                    if (properties.TryGetValue("dstrike", out var pDstrike) && IsTruthy(pDstrike))
                        rProps.DoubleStrike = new DoubleStrike();
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
                    run.AppendChild(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
                    para.AppendChild(run);
                }

                newElement = para;
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
                break;

            case "equation" or "formula" or "math":
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

                _doc.MainDocumentPart?.Document?.Save();
                return resultPath;

            case "run" or "r":
                if (parent is not Paragraph targetPara)
                    throw new ArgumentException("Runs can only be added to paragraphs");

                var newRun = new Run();
                var newRProps = new RunProperties();
                if (properties.TryGetValue("font", out var rFont))
                    newRProps.AppendChild(new RunFonts { Ascii = rFont, HighAnsi = rFont, EastAsia = rFont });
                if (properties.TryGetValue("size", out var rSize))
                    newRProps.AppendChild(new FontSize { Val = ((int)(ParseFontSize(rSize) * 2)).ToString() });
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
                if (properties.TryGetValue("strike", out var rStrike) && IsTruthy(rStrike))
                    newRProps.Strike = new Strike();
                if (properties.TryGetValue("highlight", out var rHighlight))
                    newRProps.Highlight = new Highlight { Val = ParseHighlightColor(rHighlight) };
                if (properties.TryGetValue("caps", out var rCaps) && IsTruthy(rCaps))
                    newRProps.Caps = new Caps();
                if (properties.TryGetValue("smallcaps", out var rSmallCaps) && IsTruthy(rSmallCaps))
                    newRProps.SmallCaps = new SmallCaps();
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
                newRun.AppendChild(new Text(runText) { Space = SpaceProcessingModeValues.Preserve });

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

                newElement = newRun;
                break;

            case "table" or "tbl":
                var table = new Table();
                var tblProps = new TableProperties(
                    new TableBorders(
                        new TopBorder { Val = BorderValues.Single, Size = 4 },
                        new LeftBorder { Val = BorderValues.Single, Size = 4 },
                        new BottomBorder { Val = BorderValues.Single, Size = 4 },
                        new RightBorder { Val = BorderValues.Single, Size = 4 },
                        new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                        new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 }
                    )
                );
                table.AppendChild(tblProps);

                // Apply border properties from Add parameters
                foreach (var (bk, bv) in properties)
                {
                    if (bk.StartsWith("border", StringComparison.OrdinalIgnoreCase))
                        ApplyTableBorders(tblProps, bk, bv);
                }

                int rows = 1;
                if (properties.TryGetValue("rows", out var rowsStr))
                {
                    if (!int.TryParse(rowsStr, out rows))
                        throw new ArgumentException($"Invalid 'rows' value: '{rowsStr}'. Expected a positive integer.");
                    if (rows <= 0)
                        throw new ArgumentException($"Invalid 'rows' value: '{rowsStr}'. Must be a positive integer (> 0).");
                }
                int cols = 1;
                if (properties.TryGetValue("cols", out var colsStr))
                {
                    cols = ParseHelpers.SafeParseInt(colsStr, "cols");
                    if (cols <= 0)
                        throw new ArgumentException($"Invalid 'cols' value: '{colsStr}'. Must be a positive integer (> 0).");
                }

                // Parse per-column widths: colWidths="3000,2000,5000"
                int[]? colWidthArr = null;
                if (properties.TryGetValue("colwidths", out var cwStr))
                {
                    var parts = cwStr.Split(',');
                    colWidthArr = new int[parts.Length];
                    for (int ci = 0; ci < parts.Length; ci++)
                    {
                        if (!int.TryParse(parts[ci].Trim(), out colWidthArr[ci]))
                            throw new ArgumentException($"Invalid 'colwidths' value: '{parts[ci].Trim()}'. Each column width must be a positive integer (in twips). Example: colwidths=3000,2000,5000");
                    }
                }

                // Add table grid
                var tblGrid = new TableGrid();
                for (int gc = 0; gc < cols; gc++)
                {
                    var w = colWidthArr != null && gc < colWidthArr.Length ? colWidthArr[gc].ToString() : "2400";
                    tblGrid.AppendChild(new GridColumn { Width = w });
                }
                table.AppendChild(tblGrid);

                // Apply table-level properties from Add parameters
                foreach (var (tk, tv) in properties)
                {
                    var tkl = tk.ToLowerInvariant();
                    if (tkl is "rows" or "cols" or "colwidths" || tkl.StartsWith("border")) continue;
                    switch (tkl)
                    {
                        case "alignment":
                            tblProps.TableJustification = new TableJustification
                            {
                                Val = tv.ToLowerInvariant() switch
                                {
                                    "center" => TableRowAlignmentValues.Center,
                                    "right" => TableRowAlignmentValues.Right,
                                    "left" => TableRowAlignmentValues.Left,
                                    _ => throw new ArgumentException($"Invalid table alignment value: '{tv}'. Valid values: left, center, right.")
                                }
                            };
                            break;
                        case "width":
                            if (tv.EndsWith('%'))
                            {
                                var pct = ParseHelpers.SafeParseInt(tv.TrimEnd('%'), "width") * 50;
                                tblProps.TableWidth = new TableWidth { Width = pct.ToString(), Type = TableWidthUnitValues.Pct };
                            }
                            else
                            {
                                tblProps.TableWidth = new TableWidth { Width = ParseHelpers.SafeParseUint(tv, "width").ToString(), Type = TableWidthUnitValues.Dxa };
                            }
                            break;
                        case "indent":
                            tblProps.TableIndentation = new TableIndentation { Width = ParseHelpers.SafeParseInt(tv, "indent"), Type = TableWidthUnitValues.Dxa };
                            break;
                        case "cellspacing":
                            tblProps.TableCellSpacing = new TableCellSpacing { Width = ParseHelpers.SafeParseUint(tv, "cellspacing").ToString(), Type = TableWidthUnitValues.Dxa };
                            break;
                        case "layout":
                            tblProps.TableLayout = new TableLayout
                            {
                                Type = tv.ToLowerInvariant() == "fixed" ? TableLayoutValues.Fixed : TableLayoutValues.Autofit
                            };
                            break;
                        case "padding":
                            var cm = tblProps.TableCellMarginDefault ?? tblProps.AppendChild(new TableCellMarginDefault());
                            var paddingVal = ParseHelpers.SafeParseInt(tv, "padding");
                            cm.TopMargin = new TopMargin { Width = tv, Type = TableWidthUnitValues.Dxa };
                            cm.TableCellLeftMargin = new TableCellLeftMargin { Width = (short)Math.Min(paddingVal, short.MaxValue), Type = TableWidthValues.Dxa };
                            cm.BottomMargin = new BottomMargin { Width = tv, Type = TableWidthUnitValues.Dxa };
                            cm.TableCellRightMargin = new TableCellRightMargin { Width = (short)Math.Min(paddingVal, short.MaxValue), Type = TableWidthValues.Dxa };
                            break;
                        case "style":
                            tblProps.TableStyle = new TableStyle { Val = tv };
                            // Add TableLook so built-in styles apply banding correctly
                            tblProps.RemoveAllChildren<TableLook>();
                            tblProps.AppendChild(new TableLook
                            {
                                Val = "04A0",
                                FirstRow = true,
                                LastRow = false,
                                FirstColumn = true,
                                LastColumn = false,
                                NoHorizontalBand = false,
                                NoVerticalBand = true
                            });
                            break;
                    }
                }

                for (int r = 0; r < rows; r++)
                {
                    var row = new TableRow();
                    for (int c = 0; c < cols; c++)
                    {
                        var cellPara = new Paragraph(new ParagraphProperties(
                            new SpacingBetweenLines { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto }));
                        var cell = new TableCell(cellPara);
                        if (colWidthArr != null && c < colWidthArr.Length)
                            cell.PrependChild(new TableCellProperties(new TableCellWidth { Width = colWidthArr[c].ToString(), Type = TableWidthUnitValues.Dxa }));
                        row.AppendChild(cell);
                    }
                    table.AppendChild(row);
                }

                AppendToParent(parent, table);
                var tblCount = parent.Elements<Table>().Count();
                resultPath = $"{parentPath}/tbl[{tblCount}]";
                newElement = table;
                break;

            case "row" or "tr":
            {
                if (parent is not Table targetTable)
                    throw new ArgumentException("Rows can only be added to a table: /body/tbl[N]");

                var existingCols = targetTable.Elements<TableGrid>().FirstOrDefault()
                    ?.Elements<GridColumn>().Count() ?? 1;
                int newCols = existingCols;
                if (properties.TryGetValue("cols", out var colsVal))
                    newCols = ParseHelpers.SafeParseInt(colsVal, "cols");

                var newRow = new TableRow();
                TableRowProperties? newRowProps = null;
                if (properties.TryGetValue("height", out var rowHeight))
                {
                    newRowProps ??= newRow.AppendChild(new TableRowProperties());
                    newRowProps.AppendChild(new TableRowHeight { Val = ParseTwips(rowHeight), HeightType = HeightRuleValues.AtLeast });
                }
                if (properties.TryGetValue("height.exact", out var rowHeightExact))
                {
                    newRowProps ??= newRow.AppendChild(new TableRowProperties());
                    newRowProps.GetFirstChild<TableRowHeight>()?.Remove();
                    newRowProps.AppendChild(new TableRowHeight { Val = ParseTwips(rowHeightExact), HeightType = HeightRuleValues.Exact });
                }
                if (properties.TryGetValue("header", out var headerVal) && IsTruthy(headerVal))
                {
                    newRowProps ??= newRow.AppendChild(new TableRowProperties());
                    newRowProps.AppendChild(new TableHeader());
                }

                for (int c = 0; c < newCols; c++)
                {
                    var cellText = properties.TryGetValue($"c{c + 1}", out var ct) ? ct : "";
                    var cellPara = new Paragraph();
                    if (!string.IsNullOrEmpty(cellText))
                        cellPara.AppendChild(new Run(new Text(cellText) { Space = SpaceProcessingModeValues.Preserve }));
                    newRow.AppendChild(new TableCell(cellPara));
                }

                if (index.HasValue)
                {
                    var existingRows = targetTable.Elements<TableRow>().ToList();
                    if (index.Value < existingRows.Count)
                        targetTable.InsertBefore(newRow, existingRows[index.Value]);
                    else
                        targetTable.AppendChild(newRow);
                }
                else
                {
                    targetTable.AppendChild(newRow);
                }

                var rowIdx = targetTable.Elements<TableRow>().ToList().IndexOf(newRow) + 1;
                resultPath = $"{parentPath}/tr[{rowIdx}]";
                newElement = newRow;
                break;
            }

            case "cell" or "tc":
            {
                if (parent is not TableRow targetRow)
                    throw new ArgumentException("Cells can only be added to a table row: /body/tbl[N]/tr[M]");

                var cellParagraph = new Paragraph();
                if (properties.TryGetValue("text", out var cellTxt))
                    cellParagraph.AppendChild(new Run(new Text(cellTxt) { Space = SpaceProcessingModeValues.Preserve }));

                var newCell = new TableCell(cellParagraph);

                if (properties.TryGetValue("width", out var cellWidth))
                    newCell.PrependChild(new TableCellProperties(new TableCellWidth { Width = cellWidth, Type = TableWidthUnitValues.Dxa }));

                if (index.HasValue)
                {
                    var cells = targetRow.Elements<TableCell>().ToList();
                    if (index.Value < cells.Count)
                        targetRow.InsertBefore(newCell, cells[index.Value]);
                    else
                        targetRow.AppendChild(newCell);
                }
                else
                {
                    targetRow.AppendChild(newCell);
                }

                var cellIdx = targetRow.Elements<TableCell>().ToList().IndexOf(newCell) + 1;
                resultPath = $"{parentPath}/tc[{cellIdx}]";
                newElement = newCell;
                break;
            }

            case "chart":
            {
                var chartMainPart = _doc.MainDocumentPart!;

                // Parse chart data
                var chartType = properties.FirstOrDefault(kv =>
                    kv.Key.Equals("charttype", StringComparison.OrdinalIgnoreCase)
                    || kv.Key.Equals("type", StringComparison.OrdinalIgnoreCase)).Value
                    ?? "column";
                var chartTitle = properties.GetValueOrDefault("title");
                var categories = Core.ChartHelper.ParseCategories(properties);
                var seriesData = Core.ChartHelper.ParseSeriesData(properties);

                if (seriesData.Count == 0)
                    throw new ArgumentException("Chart requires data. Use: data=\"Series1:1,2,3;Series2:4,5,6\" " +
                        "or series1=\"Revenue:100,200,300\"");

                // Create ChartPart and build chart
                var chartPart = chartMainPart.AddNewPart<ChartPart>();
                chartPart.ChartSpace = Core.ChartHelper.BuildChartSpace(chartType, chartTitle, categories, seriesData, properties);

                // Apply deferred properties (axisTitle, dataLabels, etc.) via SetChartProperties
                // Must be called BEFORE Save() so the in-memory DOM is still available
                var deferredProps = properties
                    .Where(kv => Core.ChartHelper.DeferredAddKeys.Contains(kv.Key))
                    .ToDictionary(kv => kv.Key, kv => kv.Value);
                if (deferredProps.Count > 0)
                    Core.ChartHelper.SetChartProperties(chartPart, deferredProps);
                else
                    chartPart.ChartSpace.Save();

                var chartRelId = chartMainPart.GetIdOfPart(chartPart);

                // Dimensions (default: 15cm x 10cm)
                long chartCx = properties.TryGetValue("width", out var chartWStr) ? ParseEmu(chartWStr) : 5400000;
                long chartCy = properties.TryGetValue("height", out var chStr) ? ParseEmu(chStr) : 3600000;

                var docPropId = NextImageId();
                var chartName = chartTitle ?? $"Chart {docPropId}";

                // Build Drawing/Inline with ChartReference
                var inline = new DW.Inline(
                    new DW.Extent { Cx = chartCx, Cy = chartCy },
                    new DW.EffectExtent { LeftEdge = 0, TopEdge = 0, RightEdge = 0, BottomEdge = 0 },
                    new DW.DocProperties { Id = docPropId, Name = chartName },
                    new DW.NonVisualGraphicFrameDrawingProperties(),
                    new A.Graphic(
                        new A.GraphicData(
                            new DocumentFormat.OpenXml.Drawing.Charts.ChartReference { Id = chartRelId }
                        )
                        { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" }
                    )
                )
                {
                    DistanceFromTop = 0U,
                    DistanceFromBottom = 0U,
                    DistanceFromLeft = 0U,
                    DistanceFromRight = 0U
                };

                var chartRun = new Run(new Drawing(inline));
                Paragraph chartPara;
                if (parent is Paragraph existingChartPara)
                {
                    existingChartPara.AppendChild(chartRun);
                    chartPara = existingChartPara;
                }
                else
                {
                    chartPara = new Paragraph(chartRun);
                    AppendToParent(parent, chartPara);
                }
                newElement = chartPara;

                var chartIdx = chartMainPart.ChartParts.ToList().IndexOf(chartPart) + 1;
                resultPath = $"/chart[{chartIdx}]";
                break;
            }

            case "picture" or "image" or "img":
                if (!properties.TryGetValue("path", out var imgPath) && !properties.TryGetValue("src", out imgPath))
                    throw new ArgumentException("'path' (or 'src') property is required for picture type");
                if (!File.Exists(imgPath))
                    throw new FileNotFoundException($"Image file not found: {imgPath}");

                var imgExtension = Path.GetExtension(imgPath).ToLowerInvariant();
                var imgPartType = imgExtension switch
                {
                    ".png" => ImagePartType.Png,
                    ".jpg" or ".jpeg" => ImagePartType.Jpeg,
                    ".gif" => ImagePartType.Gif,
                    ".bmp" => ImagePartType.Bmp,
                    ".tif" or ".tiff" => ImagePartType.Tiff,
                    ".emf" => ImagePartType.Emf,
                    ".wmf" => ImagePartType.Wmf,
                    ".svg" => ImagePartType.Svg,
                    _ => throw new ArgumentException($"Unsupported image format: {imgExtension}")
                };

                var mainPart = _doc.MainDocumentPart!;
                var imagePart = mainPart.AddImagePart(imgPartType);
                using (var stream = File.OpenRead(imgPath))
                    imagePart.FeedData(stream);
                var relId = mainPart.GetIdOfPart(imagePart);

                // Determine dimensions (default: 6 inches wide, auto height)
                long cxEmu = 5486400; // 6 inches in EMUs (914400 * 6)
                long cyEmu = 3657600; // 4 inches default
                if (properties.TryGetValue("width", out var widthStr))
                    cxEmu = ParseEmu(widthStr);
                if (properties.TryGetValue("height", out var heightStr))
                    cyEmu = ParseEmu(heightStr);

                var altText = properties.GetValueOrDefault("alt", Path.GetFileName(imgPath));

                Run imgRun;
                if (properties.TryGetValue("anchor", out var anchorVal) && IsTruthy(anchorVal))
                {
                    var wrapType = properties.GetValueOrDefault("wrap", "none");
                    long hPos = properties.TryGetValue("hposition", out var hPosStr) ? ParseEmu(hPosStr) : 0;
                    long vPos = properties.TryGetValue("vposition", out var vPosStr) ? ParseEmu(vPosStr) : 0;
                    var hRel = properties.TryGetValue("hrelative", out var hRelStr)
                        ? ParseHorizontalRelative(hRelStr)
                        : DW.HorizontalRelativePositionValues.Margin;
                    var vRel = properties.TryGetValue("vrelative", out var vRelStr)
                        ? ParseVerticalRelative(vRelStr)
                        : DW.VerticalRelativePositionValues.Margin;
                    var behind = properties.TryGetValue("behindtext", out var behindStr) && IsTruthy(behindStr);
                    imgRun = CreateAnchorImageRun(relId, cxEmu, cyEmu, altText, wrapType, hPos, vPos, hRel, vRel, behind);
                }
                else
                {
                    imgRun = CreateImageRun(relId, cxEmu, cyEmu, altText);
                }

                Paragraph imgPara;
                if (parent is Paragraph existingPara)
                {
                    existingPara.AppendChild(imgRun);
                    imgPara = existingPara;
                    var imgRunCount = existingPara.Elements<Run>().Count();
                    resultPath = $"{parentPath}/r[{imgRunCount}]";
                }
                else if (parent is TableCell imgCell)
                {
                    // Insert image into existing first paragraph if empty, otherwise create new paragraph
                    var firstCellPara = imgCell.Elements<Paragraph>().FirstOrDefault();
                    if (firstCellPara != null && !firstCellPara.Elements<Run>().Any())
                    {
                        firstCellPara.AppendChild(imgRun);
                        imgPara = firstCellPara;
                    }
                    else
                    {
                        imgPara = new Paragraph(imgRun);
                        imgCell.AppendChild(imgPara);
                    }
                    var imgPIdx = imgCell.Elements<Paragraph>().ToList().IndexOf(imgPara) + 1;
                    resultPath = $"{parentPath}/p[{imgPIdx}]";
                }
                else
                {
                    imgPara = new Paragraph(imgRun);
                    var imgParaCount = parent.Elements<Paragraph>().Count();
                    if (index.HasValue && index.Value < imgParaCount)
                    {
                        var refPara = parent.Elements<Paragraph>().ElementAt(index.Value);
                        parent.InsertBefore(imgPara, refPara);
                        resultPath = $"{parentPath}/p[{index.Value + 1}]";
                    }
                    else
                    {
                        AppendToParent(parent, imgPara);
                        resultPath = $"{parentPath}/p[{imgParaCount + 1}]";
                    }
                }
                newElement = imgPara;
                break;

            case "comment":
            {
                if (!properties.TryGetValue("text", out var commentText))
                    throw new ArgumentException("'text' property is required for comment type");

                var commentRun = parent as Run;
                var commentPara = commentRun?.Parent as Paragraph ?? parent as Paragraph
                    ?? throw new ArgumentException("Comments must be added to a paragraph or run: /body/p[N] or /body/p[N]/r[M]");

                var author = properties.GetValueOrDefault("author", "officecli");
                var initials = properties.GetValueOrDefault("initials", author[..1]);
                var commentsPart = _doc.MainDocumentPart!.WordprocessingCommentsPart
                    ?? _doc.MainDocumentPart.AddNewPart<WordprocessingCommentsPart>();
                commentsPart.Comments ??= new Comments();

                var commentId = (commentsPart.Comments.Elements<Comment>()
                    .Select(c => int.TryParse(c.Id?.Value, out var id) ? id : 0)
                    .DefaultIfEmpty(0).Max() + 1).ToString();

                commentsPart.Comments.AppendChild(new Comment(
                    new Paragraph(new Run(new Text(commentText) { Space = SpaceProcessingModeValues.Preserve })))
                {
                    Id = commentId, Author = author, Initials = initials,
                    Date = properties.TryGetValue("date", out var ds) ? DateTime.Parse(ds) : DateTime.UtcNow
                });
                commentsPart.Comments.Save();

                var rangeStart = new CommentRangeStart { Id = commentId };
                var rangeEnd = new CommentRangeEnd { Id = commentId };
                var refRun = new Run(new CommentReference { Id = commentId });

                if (commentRun != null)
                {
                    commentRun.InsertBeforeSelf(rangeStart);
                    commentRun.InsertAfterSelf(rangeEnd);
                    rangeEnd.InsertAfterSelf(refRun);
                }
                else
                {
                    var after = commentPara.ParagraphProperties as OpenXmlElement;
                    if (after != null) after.InsertAfterSelf(rangeStart);
                    else commentPara.InsertAt(rangeStart, 0);
                    commentPara.AppendChild(rangeEnd);
                    commentPara.AppendChild(refRun);
                }

                newElement = rangeStart;
                resultPath = $"{parentPath}/comment[{commentId}]";
                break;
            }

            case "bookmark":
            {
                var bkName = properties.GetValueOrDefault("name", "");
                if (string.IsNullOrEmpty(bkName))
                    throw new ArgumentException("'name' property is required for bookmark");

                var existingIds = body.Descendants<BookmarkStart>()
                    .Select(b => int.TryParse(b.Id?.Value, out var id) ? id : 0);
                var bkId = (existingIds.Any() ? existingIds.Max() + 1 : 1).ToString();

                var bookmarkStart = new BookmarkStart { Id = bkId, Name = bkName };
                var bookmarkEnd = new BookmarkEnd { Id = bkId };

                if (properties.TryGetValue("text", out var bkText))
                {
                    parent.AppendChild(bookmarkStart);
                    parent.AppendChild(new Run(new Text(bkText) { Space = SpaceProcessingModeValues.Preserve }));
                    parent.AppendChild(bookmarkEnd);
                }
                else
                {
                    parent.AppendChild(bookmarkStart);
                    parent.AppendChild(bookmarkEnd);
                }

                newElement = bookmarkStart;
                resultPath = $"{parentPath}/bookmark[{bkName}]";
                break;
            }

            case "hyperlink" or "link":
            {
                if (!properties.TryGetValue("url", out var hlUrl) && !properties.TryGetValue("href", out hlUrl))
                    throw new ArgumentException("'url' property is required for hyperlink type");

                if (parent is not Paragraph hlPara)
                    throw new ArgumentException("Hyperlinks can only be added to paragraphs: /body/p[N]");

                var mainDocPart = _doc.MainDocumentPart!;
                if (!Uri.TryCreate(hlUrl, UriKind.Absolute, out var hlUri))
                    throw new ArgumentException($"Invalid hyperlink URL '{hlUrl}'. Expected a valid absolute URI (e.g. 'https://example.com').");
                var hlRelId = mainDocPart.AddHyperlinkRelationship(hlUri, isExternal: true).Id;

                var hlRProps = new RunProperties();
                if (properties.TryGetValue("color", out var hlColor))
                    hlRProps.Color = new Color { Val = SanitizeHex(hlColor) };
                else
                    hlRProps.Color = new Color { Val = "0563C1" };
                hlRProps.Underline = new Underline { Val = UnderlineValues.Single };
                if (properties.TryGetValue("font", out var hlFont))
                    hlRProps.RunFonts = new RunFonts { Ascii = hlFont, HighAnsi = hlFont };
                if (properties.TryGetValue("size", out var hlSize))
                    hlRProps.FontSize = new FontSize { Val = ((int)(ParseFontSize(hlSize) * 2)).ToString() };
                if (properties.TryGetValue("bold", out var hlBold) && IsTruthy(hlBold))
                    hlRProps.Bold = new Bold();
                if (properties.TryGetValue("italic", out var hlItalic) && IsTruthy(hlItalic))
                    hlRProps.Italic = new Italic();

                var hlRun = new Run(hlRProps);
                var hlText = properties.GetValueOrDefault("text", hlUrl);
                hlRun.AppendChild(new Text(hlText) { Space = SpaceProcessingModeValues.Preserve });

                var hyperlink = new Hyperlink(hlRun) { Id = hlRelId };
                if (index.HasValue)
                    hlPara.InsertAt(hyperlink, index.Value);
                else
                    hlPara.AppendChild(hyperlink);

                var hlCount = hlPara.Elements<Hyperlink>().Count();
                resultPath = $"{parentPath}/hyperlink[{hlCount}]";
                newElement = hyperlink;
                break;
            }

            case "section" or "sectionbreak":
            {
                // Section break: adds SectionProperties to the last paragraph before the break point
                var breakType = properties.GetValueOrDefault("type", "nextPage").ToLowerInvariant();
                var sectType = breakType switch
                {
                    "nextpage" or "next" => SectionMarkValues.NextPage,
                    "continuous" => SectionMarkValues.Continuous,
                    "evenpage" or "even" => SectionMarkValues.EvenPage,
                    "oddpage" or "odd" => SectionMarkValues.OddPage,
                    _ => throw new ArgumentException($"Invalid section break type: '{breakType}'. Valid values: nextPage, continuous, evenPage, oddPage.")
                };

                // Create a paragraph with section properties to mark the break
                var sectPara = new Paragraph();
                var sectPProps = new ParagraphProperties();
                var sectPr = new SectionProperties();
                sectPr.AppendChild(new SectionType { Val = sectType });

                // Copy page size/margins from document section, or use A4 defaults
                var bodySectPr = body.GetFirstChild<SectionProperties>();
                var srcPageSize = bodySectPr?.GetFirstChild<PageSize>();
                sectPr.AppendChild(new PageSize
                {
                    Width = srcPageSize?.Width ?? 11906,   // A4 width
                    Height = srcPageSize?.Height ?? 16838,  // A4 height
                    Orient = srcPageSize?.Orient
                });
                var srcMargin = bodySectPr?.GetFirstChild<PageMargin>();
                sectPr.AppendChild(new PageMargin
                {
                    Top = srcMargin?.Top ?? 1440,
                    Bottom = srcMargin?.Bottom ?? 1440,
                    Left = srcMargin?.Left ?? 1800,
                    Right = srcMargin?.Right ?? 1800
                });

                // Allow per-section overrides
                if (properties.TryGetValue("pagewidth", out var sw) || properties.TryGetValue("pageWidth", out sw) || properties.TryGetValue("width", out sw))
                {
                    (sectPr.GetFirstChild<PageSize>() ?? sectPr.AppendChild(new PageSize())).Width = ParseHelpers.SafeParseUint(sw, "pagewidth");
                }
                if (properties.TryGetValue("pageheight", out var sh) || properties.TryGetValue("pageHeight", out sh) || properties.TryGetValue("height", out sh))
                {
                    (sectPr.GetFirstChild<PageSize>() ?? sectPr.AppendChild(new PageSize())).Height = ParseHelpers.SafeParseUint(sh, "pageheight");
                }
                if (properties.TryGetValue("orientation", out var orient))
                {
                    var ps = sectPr.GetFirstChild<PageSize>() ?? sectPr.AppendChild(new PageSize());
                    ps.Orient = orient.ToLowerInvariant() == "landscape"
                        ? PageOrientationValues.Landscape
                        : PageOrientationValues.Portrait;
                    // Swap width/height for landscape if needed
                    if (ps.Orient == PageOrientationValues.Landscape && ps.Width < ps.Height)
                        (ps.Width!.Value, ps.Height!.Value) = (ps.Height.Value, ps.Width.Value);
                }

                sectPProps.AppendChild(sectPr);
                sectPara.AppendChild(sectPProps);
                AppendToParent(parent, sectPara);

                // Count section properties in document
                var secCount = body.Elements<Paragraph>()
                    .Count(p => p.ParagraphProperties?.GetFirstChild<SectionProperties>() != null);
                resultPath = $"/section[{secCount}]";
                newElement = sectPara;
                break;
            }

            case "footnote":
            {
                if (!properties.TryGetValue("text", out var fnText))
                    throw new ArgumentException("'text' property is required for footnote type");

                if (parent is not Paragraph fnPara)
                    throw new ArgumentException("Footnotes must be added to a paragraph: /body/p[N]");

                var mainPart2 = _doc.MainDocumentPart!;
                var fnPart = mainPart2.FootnotesPart ?? mainPart2.AddNewPart<FootnotesPart>();
                fnPart.Footnotes ??= new Footnotes(
                    new Footnote(new Paragraph(new Run(new Text("")))) { Type = FootnoteEndnoteValues.Separator, Id = -1 },
                    new Footnote(new Paragraph(new Run(new Text("")))) { Type = FootnoteEndnoteValues.ContinuationSeparator, Id = 0 }
                );

                var fnId = (fnPart.Footnotes.Elements<Footnote>()
                    .Where(f => f.Id?.Value > 0)
                    .Select(f => f.Id!.Value)
                    .DefaultIfEmpty(0).Max() + 1);

                var footnote = new Footnote { Id = fnId };
                var fnContentPara = new Paragraph(
                    new ParagraphProperties(new ParagraphStyleId { Val = "FootnoteText" }),
                    new Run(
                        new RunProperties(new VerticalTextAlignment { Val = VerticalPositionValues.Superscript }),
                        new FootnoteReferenceMark()),
                    new Run(new Text(" " + fnText) { Space = SpaceProcessingModeValues.Preserve })
                );
                footnote.AppendChild(fnContentPara);
                fnPart.Footnotes.AppendChild(footnote);
                fnPart.Footnotes.Save();

                // Insert reference in document body
                var fnRefRun = new Run(
                    new RunProperties(new RunStyle { Val = "FootnoteReference" }),
                    new FootnoteReference { Id = fnId }
                );
                fnPara.AppendChild(fnRefRun);

                resultPath = $"/footnote[{fnId}]";
                newElement = fnRefRun;
                break;
            }

            case "endnote":
            {
                if (!properties.TryGetValue("text", out var enText))
                    throw new ArgumentException("'text' property is required for endnote type");

                if (parent is not Paragraph enPara)
                    throw new ArgumentException("Endnotes must be added to a paragraph: /body/p[N]");

                var mainPart3 = _doc.MainDocumentPart!;
                var enPart = mainPart3.EndnotesPart ?? mainPart3.AddNewPart<EndnotesPart>();
                enPart.Endnotes ??= new Endnotes(
                    new Endnote(new Paragraph(new Run(new Text("")))) { Type = FootnoteEndnoteValues.Separator, Id = -1 },
                    new Endnote(new Paragraph(new Run(new Text("")))) { Type = FootnoteEndnoteValues.ContinuationSeparator, Id = 0 }
                );

                var enId = (enPart.Endnotes.Elements<Endnote>()
                    .Where(e => e.Id?.Value > 0)
                    .Select(e => e.Id!.Value)
                    .DefaultIfEmpty(0).Max() + 1);

                var endnote = new Endnote { Id = enId };
                var enContentPara = new Paragraph(
                    new ParagraphProperties(new ParagraphStyleId { Val = "EndnoteText" }),
                    new Run(
                        new RunProperties(new VerticalTextAlignment { Val = VerticalPositionValues.Superscript }),
                        new EndnoteReferenceMark()),
                    new Run(new Text(" " + enText) { Space = SpaceProcessingModeValues.Preserve })
                );
                endnote.AppendChild(enContentPara);
                enPart.Endnotes.AppendChild(endnote);
                enPart.Endnotes.Save();

                // Insert reference in document body
                var enRefRun = new Run(
                    new RunProperties(new RunStyle { Val = "EndnoteReference" }),
                    new EndnoteReference { Id = enId }
                );
                enPara.AppendChild(enRefRun);

                resultPath = $"/endnote[{enId}]";
                newElement = enRefRun;
                break;
            }

            case "toc" or "tableofcontents":
            {
                // Table of Contents field code
                var levels = properties.GetValueOrDefault("levels", "1-3");
                var tocTitle = properties.GetValueOrDefault("title", "");
                var hyperlinks = !properties.TryGetValue("hyperlinks", out var hlVal) || IsTruthy(hlVal);
                var pageNumbers = !properties.TryGetValue("pagenumbers", out var pnVal) || IsTruthy(pnVal);

                // Build field code instruction
                var instrBuilder = new StringBuilder($" TOC \\o \"{levels}\"");
                if (hyperlinks) instrBuilder.Append(" \\h");
                if (!pageNumbers) instrBuilder.Append(" \\z");
                instrBuilder.Append(" \\u ");

                var tocPara = new Paragraph();

                // Optional title
                if (!string.IsNullOrEmpty(tocTitle))
                {
                    var titlePara = new Paragraph(
                        new ParagraphProperties(new ParagraphStyleId { Val = "TOCHeading" }),
                        new Run(new Text(tocTitle))
                    );
                    AppendToParent(parent, titlePara);
                }

                // Field begin
                tocPara.AppendChild(new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }));
                // Field code
                tocPara.AppendChild(new Run(new FieldCode(instrBuilder.ToString()) { Space = SpaceProcessingModeValues.Preserve }));
                // Field separate
                tocPara.AppendChild(new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }));
                // Placeholder text
                tocPara.AppendChild(new Run(new Text("Update field to see table of contents") { Space = SpaceProcessingModeValues.Preserve }));
                // Field end
                tocPara.AppendChild(new Run(new FieldChar { FieldCharType = FieldCharValues.End }));

                AppendToParent(parent, tocPara);

                // Add UpdateFieldsOnOpen setting
                var settingsPart2 = _doc.MainDocumentPart!.DocumentSettingsPart
                    ?? _doc.MainDocumentPart.AddNewPart<DocumentSettingsPart>();
                settingsPart2.Settings ??= new Settings();
                if (settingsPart2.Settings.GetFirstChild<UpdateFieldsOnOpen>() == null)
                    settingsPart2.Settings.AppendChild(new UpdateFieldsOnOpen { Val = true });
                settingsPart2.Settings.Save();

                // Count TOC fields in document to determine index
                var tocCount = body.Elements<Paragraph>()
                    .Count(p => p.Descendants<FieldCode>().Any(fc =>
                        fc.Text != null && fc.Text.TrimStart().StartsWith("TOC", StringComparison.OrdinalIgnoreCase)));
                resultPath = $"/toc[{tocCount}]";
                newElement = tocPara;
                break;
            }

            case "style":
            {
                // Create a new style in the styles part
                var stylesPart = _doc.MainDocumentPart!.StyleDefinitionsPart
                    ?? _doc.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
                stylesPart.Styles ??= new Styles();

                var styleId = properties.GetValueOrDefault("id", properties.GetValueOrDefault("name", "CustomStyle"));
                var styleName = properties.GetValueOrDefault("name", styleId);
                var styleType = properties.GetValueOrDefault("type", "paragraph").ToLowerInvariant() switch
                {
                    "character" or "char" => StyleValues.Character,
                    "table" => StyleValues.Table,
                    "numbering" => StyleValues.Numbering,
                    "paragraph" or "para" => StyleValues.Paragraph,
                    _ => throw new ArgumentException($"Invalid style type: '{properties.GetValueOrDefault("type", "paragraph")}'. Valid values: paragraph, character, table, numbering.")
                };

                var newStyle = new Style
                {
                    Type = styleType,
                    StyleId = styleId,
                    CustomStyle = true
                };
                newStyle.AppendChild(new StyleName { Val = styleName });

                if (properties.TryGetValue("basedon", out var basedOn) && !string.IsNullOrEmpty(basedOn))
                    newStyle.AppendChild(new BasedOn { Val = basedOn });
                if (properties.TryGetValue("next", out var nextStyle))
                    newStyle.AppendChild(new NextParagraphStyle { Val = nextStyle });

                // Style paragraph properties
                var stylePPr = new StyleParagraphProperties();
                bool hasPPr = false;
                if (properties.TryGetValue("alignment", out var sAlign) || properties.TryGetValue("align", out sAlign))
                {
                    stylePPr.Justification = new Justification { Val = ParseJustification(sAlign) };
                    hasPPr = true;
                }
                if (properties.TryGetValue("spacebefore", out var sSBefore) || properties.TryGetValue("spaceBefore", out sSBefore))
                {
                    var sp = stylePPr.SpacingBetweenLines ?? (stylePPr.SpacingBetweenLines = new SpacingBetweenLines());
                    sp.Before = SpacingConverter.ParseWordSpacing(sSBefore).ToString();
                    hasPPr = true;
                }
                if (properties.TryGetValue("spaceafter", out var sSAfter) || properties.TryGetValue("spaceAfter", out sSAfter))
                {
                    var sp = stylePPr.SpacingBetweenLines ?? (stylePPr.SpacingBetweenLines = new SpacingBetweenLines());
                    sp.After = SpacingConverter.ParseWordSpacing(sSAfter).ToString();
                    hasPPr = true;
                }
                if (hasPPr) newStyle.AppendChild(stylePPr);

                // Style run properties
                var styleRPr = new StyleRunProperties();
                bool hasRPr = false;
                if (properties.TryGetValue("font", out var sFont))
                {
                    styleRPr.RunFonts = new RunFonts { Ascii = sFont, HighAnsi = sFont, EastAsia = sFont };
                    hasRPr = true;
                }
                if (properties.TryGetValue("size", out var sSize))
                {
                    styleRPr.FontSize = new FontSize { Val = ((int)(ParseFontSize(sSize) * 2)).ToString() };
                    hasRPr = true;
                }
                if (properties.TryGetValue("bold", out var sBold) && IsTruthy(sBold))
                {
                    styleRPr.Bold = new Bold();
                    hasRPr = true;
                }
                if (properties.TryGetValue("italic", out var sItalic) && IsTruthy(sItalic))
                {
                    styleRPr.Italic = new Italic();
                    hasRPr = true;
                }
                if (properties.TryGetValue("color", out var sColor))
                {
                    styleRPr.Color = new Color { Val = SanitizeHex(sColor) };
                    hasRPr = true;
                }
                if (hasRPr) newStyle.AppendChild(styleRPr);

                stylesPart.Styles.AppendChild(newStyle);
                stylesPart.Styles.Save();

                resultPath = $"/styles/{styleId}";
                newElement = newStyle;
                break;
            }

            case "header":
            {
                var mainPartH = _doc.MainDocumentPart!;
                var headerPart = mainPartH.AddNewPart<HeaderPart>();

                var hPara = new Paragraph();
                var hPProps = new ParagraphProperties();

                if (properties.TryGetValue("alignment", out var hAlign))
                    hPProps.Justification = new Justification { Val = ParseJustification(hAlign) };
                hPara.AppendChild(hPProps);

                if (properties.TryGetValue("text", out var hText))
                {
                    var hRun = new Run();
                    var hRProps = new RunProperties();
                    if (properties.TryGetValue("font", out var hFont))
                        hRProps.AppendChild(new RunFonts { Ascii = hFont, HighAnsi = hFont, EastAsia = hFont });
                    if (properties.TryGetValue("size", out var hSize))
                        hRProps.AppendChild(new FontSize { Val = ((int)(ParseFontSize(hSize) * 2)).ToString() });
                    if (properties.TryGetValue("bold", out var hBold) && IsTruthy(hBold))
                        hRProps.Bold = new Bold();
                    if (properties.TryGetValue("italic", out var hItalic) && IsTruthy(hItalic))
                        hRProps.Italic = new Italic();
                    if (properties.TryGetValue("color", out var hColor))
                        hRProps.Color = new Color { Val = SanitizeHex(hColor) };
                    hRun.AppendChild(hRProps);
                    hRun.AppendChild(new Text(hText) { Space = SpaceProcessingModeValues.Preserve });
                    hPara.AppendChild(hRun);
                }

                headerPart.Header = new Header(hPara);
                headerPart.Header.Save();

                var hBody = mainPartH.Document!.Body!;
                var hSectPr = hBody.Elements<SectionProperties>().LastOrDefault()
                    ?? hBody.AppendChild(new SectionProperties());

                var headerType = HeaderFooterValues.Default;
                if (properties.TryGetValue("type", out var hTypeStr))
                {
                    headerType = hTypeStr.ToLowerInvariant() switch
                    {
                        "first" => HeaderFooterValues.First,
                        "even" => HeaderFooterValues.Even,
                        "default" => HeaderFooterValues.Default,
                        _ => throw new ArgumentException($"Invalid header type: '{hTypeStr}'. Valid values: default, first, even.")
                    };
                }

                var headerRef = new HeaderReference
                {
                    Id = mainPartH.GetIdOfPart(headerPart),
                    Type = headerType
                };
                hSectPr.PrependChild(headerRef);

                if (headerType == HeaderFooterValues.First)
                {
                    var settingsPart = mainPartH.DocumentSettingsPart
                        ?? mainPartH.AddNewPart<DocumentSettingsPart>();
                    settingsPart.Settings ??= new Settings();
                    if (settingsPart.Settings.GetFirstChild<TitlePage>() == null)
                        settingsPart.Settings.AppendChild(new TitlePage());
                    settingsPart.Settings.Save();
                }

                mainPartH.Document.Save();
                var hIdx = mainPartH.HeaderParts.ToList().IndexOf(headerPart);
                return $"/header[{hIdx + 1}]";
            }

            case "footer":
            {
                var mainPartF = _doc.MainDocumentPart!;
                var footerPart = mainPartF.AddNewPart<FooterPart>();

                var fPara = new Paragraph();
                var fPProps = new ParagraphProperties();

                if (properties.TryGetValue("alignment", out var fAlign))
                    fPProps.Justification = new Justification { Val = ParseJustification(fAlign) };
                fPara.AppendChild(fPProps);

                if (properties.TryGetValue("text", out var fText))
                {
                    var fRun = new Run();
                    var fRProps = new RunProperties();
                    if (properties.TryGetValue("font", out var fFont))
                        fRProps.AppendChild(new RunFonts { Ascii = fFont, HighAnsi = fFont, EastAsia = fFont });
                    if (properties.TryGetValue("size", out var fSize))
                        fRProps.AppendChild(new FontSize { Val = ((int)(ParseFontSize(fSize) * 2)).ToString() });
                    if (properties.TryGetValue("bold", out var fBold) && IsTruthy(fBold))
                        fRProps.Bold = new Bold();
                    if (properties.TryGetValue("italic", out var fItalic) && IsTruthy(fItalic))
                        fRProps.Italic = new Italic();
                    if (properties.TryGetValue("color", out var fColor))
                        fRProps.Color = new Color { Val = SanitizeHex(fColor) };
                    fRun.AppendChild(fRProps);
                    fRun.AppendChild(new Text(fText) { Space = SpaceProcessingModeValues.Preserve });
                    fPara.AppendChild(fRun);
                }

                footerPart.Footer = new Footer(fPara);
                footerPart.Footer.Save();

                var fBody = mainPartF.Document!.Body!;
                var fSectPr = fBody.Elements<SectionProperties>().LastOrDefault()
                    ?? fBody.AppendChild(new SectionProperties());

                var footerType = HeaderFooterValues.Default;
                if (properties.TryGetValue("type", out var fTypeStr))
                {
                    footerType = fTypeStr.ToLowerInvariant() switch
                    {
                        "first" => HeaderFooterValues.First,
                        "even" => HeaderFooterValues.Even,
                        "default" => HeaderFooterValues.Default,
                        _ => throw new ArgumentException($"Invalid footer type: '{fTypeStr}'. Valid values: default, first, even.")
                    };
                }

                var footerRef = new FooterReference
                {
                    Id = mainPartF.GetIdOfPart(footerPart),
                    Type = footerType
                };
                fSectPr.PrependChild(footerRef);

                if (footerType == HeaderFooterValues.First)
                {
                    var settingsPart = mainPartF.DocumentSettingsPart
                        ?? mainPartF.AddNewPart<DocumentSettingsPart>();
                    settingsPart.Settings ??= new Settings();
                    if (settingsPart.Settings.GetFirstChild<TitlePage>() == null)
                        settingsPart.Settings.AppendChild(new TitlePage());
                    settingsPart.Settings.Save();
                }

                mainPartF.Document.Save();
                var fIdx = mainPartF.FooterParts.ToList().IndexOf(footerPart);
                return $"/footer[{fIdx + 1}]";
            }

            case "field" or "pagenum" or "pagenumber" or "numpages" or "date":
            {
                // Insert a field code (PAGE, NUMPAGES, DATE, etc.) as a run
                // Determines field instruction from type or "field" property
                var fieldInstr = type.ToLowerInvariant() switch
                {
                    "pagenum" or "pagenumber" => " PAGE ",
                    "numpages" => " NUMPAGES ",
                    "date" => " DATE \\@ \"yyyy-MM-dd\" ",
                    _ => properties.GetValueOrDefault("instruction", " PAGE ")
                };
                // Allow override via property
                if (properties.TryGetValue("instruction", out var instr))
                    fieldInstr = instr.StartsWith(" ") ? instr : $" {instr} ";

                var fieldPlaceholder = properties.GetValueOrDefault("text", "1");

                // Build complex field: fldChar(begin) + instrText + fldChar(separate) + result + fldChar(end)
                var fieldRunBegin = new Run(new FieldChar { FieldCharType = FieldCharValues.Begin });
                var fieldRunInstr = new Run(new FieldCode(fieldInstr) { Space = SpaceProcessingModeValues.Preserve });
                var fieldRunSep = new Run(new FieldChar { FieldCharType = FieldCharValues.Separate });
                var fieldRunResult = new Run(new Text(fieldPlaceholder) { Space = SpaceProcessingModeValues.Preserve });
                var fieldRunEnd = new Run(new FieldChar { FieldCharType = FieldCharValues.End });

                // Apply optional run formatting to all runs
                RunProperties? fieldRProps = null;
                if (properties.TryGetValue("font", out var fFont) || properties.TryGetValue("size", out _) ||
                    properties.TryGetValue("bold", out _) || properties.TryGetValue("color", out _))
                {
                    fieldRProps = new RunProperties();
                    if (properties.TryGetValue("font", out var ff))
                        fieldRProps.AppendChild(new RunFonts { Ascii = ff, HighAnsi = ff, EastAsia = ff });
                    if (properties.TryGetValue("size", out var fs))
                        fieldRProps.AppendChild(new FontSize { Val = ((int)(ParseFontSize(fs) * 2)).ToString() });
                    if (properties.TryGetValue("bold", out var fb) && IsTruthy(fb))
                        fieldRProps.AppendChild(new Bold());
                    if (properties.TryGetValue("color", out var fc))
                        fieldRProps.AppendChild(new Color { Val = SanitizeHex(fc) });
                }

                if (fieldRProps != null)
                {
                    fieldRunBegin.PrependChild(fieldRProps.CloneNode(true));
                    fieldRunInstr.PrependChild(fieldRProps.CloneNode(true));
                    fieldRunSep.PrependChild(fieldRProps.CloneNode(true));
                    fieldRunResult.PrependChild(fieldRProps.CloneNode(true));
                    fieldRunEnd.PrependChild(fieldRProps.CloneNode(true));
                }

                if (parent is Paragraph fieldPara)
                {
                    fieldPara.AppendChild(fieldRunBegin);
                    fieldPara.AppendChild(fieldRunInstr);
                    fieldPara.AppendChild(fieldRunSep);
                    fieldPara.AppendChild(fieldRunResult);
                    fieldPara.AppendChild(fieldRunEnd);
                    newElement = fieldRunBegin;
                    var fParaIdx = body.Elements<Paragraph>().TakeWhile(p => p != fieldPara).Count();
                    resultPath = $"/body/p[{fParaIdx + 1}]/r[{GetAllRuns(fieldPara).Count - 4}]";
                }
                else
                {
                    // Create a new paragraph containing the field
                    var fNewPara = new Paragraph();
                    var fPProps = new ParagraphProperties();
                    if (properties.TryGetValue("alignment", out var fAlign))
                        fPProps.Justification = new Justification { Val = ParseJustification(fAlign) };
                    fNewPara.AppendChild(fPProps);
                    fNewPara.AppendChild(fieldRunBegin);
                    fNewPara.AppendChild(fieldRunInstr);
                    fNewPara.AppendChild(fieldRunSep);
                    fNewPara.AppendChild(fieldRunResult);
                    fNewPara.AppendChild(fieldRunEnd);
                    AppendToParent(parent, fNewPara);
                    newElement = fNewPara;
                    var fIdx2 = body.Elements<Paragraph>().TakeWhile(p => p != fNewPara).Count();
                    resultPath = $"/body/p[{fIdx2 + 1}]";
                }
                break;
            }

            case "pagebreak" or "columnbreak" or "break":
            {
                // Insert an explicit page break, column break, or line break
                var breakType = type.ToLowerInvariant() switch
                {
                    "columnbreak" => BreakValues.Column,
                    _ => BreakValues.Page
                };
                // Allow override via property
                if (properties.TryGetValue("type", out var brType))
                {
                    breakType = brType.ToLowerInvariant() switch
                    {
                        "page" => BreakValues.Page,
                        "column" => BreakValues.Column,
                        "textwrapping" or "line" => BreakValues.TextWrapping,
                        _ => throw new ArgumentException($"Invalid break type: '{brType}'. Valid values: page, column, line, textwrapping.")
                    };
                }

                var brk = new Break { Type = breakType };
                var brkRun = new Run(brk);

                if (parent is Paragraph brkPara)
                {
                    brkPara.AppendChild(brkRun);
                    newElement = brkRun;
                    var brkParaIdx = body.Elements<Paragraph>().TakeWhile(p => p != brkPara).Count();
                    resultPath = $"/body/p[{brkParaIdx + 1}]/r[{GetAllRuns(brkPara).Count}]";
                }
                else
                {
                    // Create a new empty paragraph with the break
                    var brkNewPara = new Paragraph(brkRun);
                    AppendToParent(parent, brkNewPara);
                    newElement = brkNewPara;
                    var brkIdx = body.Elements<Paragraph>().TakeWhile(p => p != brkNewPara).Count();
                    resultPath = $"/body/p[{brkIdx + 1}]";
                }
                break;
            }

            case "sdt" or "contentcontrol":
            {
                // Add a Structured Document Tag (Content Control)
                var sdtType = properties.GetValueOrDefault("sdttype", properties.GetValueOrDefault("controltype", "text")).ToLowerInvariant();
                var alias = properties.GetValueOrDefault("alias", properties.GetValueOrDefault("name", ""));
                var tag = properties.GetValueOrDefault("tag", "");
                var lockVal = properties.GetValueOrDefault("lock", "");
                var sdtText = properties.GetValueOrDefault("text", "");

                // Determine block-level vs inline
                bool isInline = parent is Paragraph;

                if (isInline)
                {
                    // Inline SDT (SdtRun) inside a paragraph
                    var sdtRun = new SdtRun();
                    var sdtProps = new SdtProperties();

                    // ID
                    sdtProps.AppendChild(new SdtId { Val = (int)(DateTimeOffset.UtcNow.ToUnixTimeMilliseconds() % int.MaxValue) });

                    if (!string.IsNullOrEmpty(alias))
                        sdtProps.AppendChild(new SdtAlias { Val = alias });
                    if (!string.IsNullOrEmpty(tag))
                        sdtProps.AppendChild(new Tag { Val = tag });
                    if (!string.IsNullOrEmpty(lockVal))
                    {
                        sdtProps.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Lock
                        {
                            Val = lockVal.ToLowerInvariant() switch
                            {
                                "contentlocked" or "content" => LockingValues.ContentLocked,
                                "sdtlocked" or "sdt" => LockingValues.SdtLocked,
                                "sdtcontentlocked" or "both" => LockingValues.SdtContentLocked,
                                "unlocked" or "none" => LockingValues.Unlocked,
                                _ => throw new ArgumentException($"Invalid lock value: '{lockVal}'. Valid values: unlocked, contentLocked, sdtLocked, sdtContentLocked.")
                            }
                        });
                    }

                    // Content type definition
                    switch (sdtType)
                    {
                        case "dropdown" or "dropdownlist":
                        {
                            var ddl = new SdtContentDropDownList();
                            if (properties.TryGetValue("items", out var items))
                            {
                                foreach (var item in items.Split(','))
                                {
                                    var trimmed = item.Trim();
                                    ddl.AppendChild(new ListItem { DisplayText = trimmed, Value = trimmed });
                                }
                            }
                            sdtProps.AppendChild(ddl);
                            break;
                        }
                        case "combobox" or "combo":
                        {
                            var cb = new SdtContentComboBox();
                            if (properties.TryGetValue("items", out var items))
                            {
                                foreach (var item in items.Split(','))
                                {
                                    var trimmed = item.Trim();
                                    cb.AppendChild(new ListItem { DisplayText = trimmed, Value = trimmed });
                                }
                            }
                            sdtProps.AppendChild(cb);
                            break;
                        }
                        case "date" or "datepicker":
                            var datePr = new SdtContentDate();
                            if (properties.TryGetValue("format", out var dateFmt))
                                datePr.DateFormat = new DateFormat { Val = dateFmt };
                            else
                                datePr.DateFormat = new DateFormat { Val = "yyyy-MM-dd" };
                            sdtProps.AppendChild(datePr);
                            break;
                        case "richtext" or "rich":
                            // Rich text has no specific type element (absence of w:text means rich text)
                            break;
                        default: // "text" or "plaintext"
                            sdtProps.AppendChild(new SdtContentText());
                            break;
                    }

                    sdtRun.AppendChild(sdtProps);
                    var sdtContent = new SdtContentRun();
                    var contentRun = new Run(new Text(sdtText) { Space = SpaceProcessingModeValues.Preserve });
                    sdtContent.AppendChild(contentRun);
                    sdtRun.AppendChild(sdtContent);

                    ((Paragraph)parent).AppendChild(sdtRun);
                    newElement = sdtRun;
                    var sdtParaIdx = body.Elements<Paragraph>().TakeWhile(p => p != parent).Count();
                    resultPath = $"/body/p[{sdtParaIdx + 1}]/sdt[{((Paragraph)parent).Elements<SdtRun>().Count()}]";
                }
                else
                {
                    // Block-level SDT (SdtBlock)
                    var sdtBlock = new SdtBlock();
                    var sdtProps = new SdtProperties();

                    sdtProps.AppendChild(new SdtId { Val = (int)(DateTimeOffset.UtcNow.ToUnixTimeMilliseconds() % int.MaxValue) });

                    if (!string.IsNullOrEmpty(alias))
                        sdtProps.AppendChild(new SdtAlias { Val = alias });
                    if (!string.IsNullOrEmpty(tag))
                        sdtProps.AppendChild(new Tag { Val = tag });
                    if (!string.IsNullOrEmpty(lockVal))
                    {
                        sdtProps.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Lock
                        {
                            Val = lockVal.ToLowerInvariant() switch
                            {
                                "contentlocked" or "content" => LockingValues.ContentLocked,
                                "sdtlocked" or "sdt" => LockingValues.SdtLocked,
                                "sdtcontentlocked" or "both" => LockingValues.SdtContentLocked,
                                "unlocked" or "none" => LockingValues.Unlocked,
                                _ => throw new ArgumentException($"Invalid lock value: '{lockVal}'. Valid values: unlocked, contentLocked, sdtLocked, sdtContentLocked.")
                            }
                        });
                    }

                    switch (sdtType)
                    {
                        case "dropdown" or "dropdownlist":
                        {
                            var ddl = new SdtContentDropDownList();
                            if (properties.TryGetValue("items", out var items))
                            {
                                foreach (var item in items.Split(','))
                                {
                                    var trimmed = item.Trim();
                                    ddl.AppendChild(new ListItem { DisplayText = trimmed, Value = trimmed });
                                }
                            }
                            sdtProps.AppendChild(ddl);
                            break;
                        }
                        case "combobox" or "combo":
                        {
                            var cb = new SdtContentComboBox();
                            if (properties.TryGetValue("items", out var items))
                            {
                                foreach (var item in items.Split(','))
                                {
                                    var trimmed = item.Trim();
                                    cb.AppendChild(new ListItem { DisplayText = trimmed, Value = trimmed });
                                }
                            }
                            sdtProps.AppendChild(cb);
                            break;
                        }
                        case "date" or "datepicker":
                            var datePr = new SdtContentDate();
                            if (properties.TryGetValue("format", out var dateFmt))
                                datePr.DateFormat = new DateFormat { Val = dateFmt };
                            else
                                datePr.DateFormat = new DateFormat { Val = "yyyy-MM-dd" };
                            sdtProps.AppendChild(datePr);
                            break;
                        case "richtext" or "rich":
                            break;
                        default:
                            sdtProps.AppendChild(new SdtContentText());
                            break;
                    }

                    sdtBlock.AppendChild(sdtProps);
                    var sdtContent = new SdtContentBlock();
                    var contentPara = new Paragraph(new Run(new Text(sdtText) { Space = SpaceProcessingModeValues.Preserve }));
                    sdtContent.AppendChild(contentPara);
                    sdtBlock.AppendChild(sdtContent);

                    AppendToParent(parent, sdtBlock);
                    newElement = sdtBlock;
                    var sdtCount = body.Elements<SdtBlock>().Count();
                    resultPath = $"/body/sdt[{sdtCount}]";
                }
                break;
            }

            case "watermark":
            {
                var wmText = properties.GetValueOrDefault("text", "DRAFT");
                // VML watermarks accept named colors (silver, red, etc.) or hex — don't sanitize
                var wmColor = properties.TryGetValue("color", out var wmcVal)
                    ? wmcVal.TrimStart('#') : "silver";
                var wmFont = properties.GetValueOrDefault("font", "Calibri");
                var wmSize = properties.GetValueOrDefault("size", "1pt");
                if (!wmSize.EndsWith("pt")) wmSize += "pt";
                var wmRotation = properties.GetValueOrDefault("rotation", "315");
                var wmOpacity = properties.TryGetValue("opacity", out var wmoVal) ? wmoVal : ".5";
                var wmWidth = properties.GetValueOrDefault("width", "415pt");
                var wmHeight = properties.GetValueOrDefault("height", "207.5pt");

                var mainPartWM = _doc.MainDocumentPart!;

                // Remove existing watermarks first
                RemoveWatermarkHeaders();

                // Create 3 headers (default, first, even) — same as POI's createWatermark()
                var headerTypes = new[] {
                    HeaderFooterValues.Default,
                    HeaderFooterValues.First,
                    HeaderFooterValues.Even
                };

                for (int wi = 0; wi < 3; wi++)
                {
                    var wmHeaderPart = mainPartWM.AddNewPart<HeaderPart>();
                    var wmIdx = wi + 1;

                    // Build VML watermark XML (follows POI's getWatermarkParagraph template)
                    var vmlXml = $@"<v:shapetype id=""_x0000_t136"" coordsize=""1600,21600"" o:spt=""136"" adj=""10800"" path=""m@7,0l@8,0m@5,21600l@6,21600e"" xmlns:v=""urn:schemas-microsoft-com:vml"" xmlns:o=""urn:schemas-microsoft-com:office:office"">
  <v:formulas>
    <v:f eqn=""sum #0 0 10800""/><v:f eqn=""prod #0 2 1""/><v:f eqn=""sum 21600 0 @1""/>
    <v:f eqn=""sum 0 0 @2""/><v:f eqn=""sum 21600 0 @3""/><v:f eqn=""if @0 @3 0""/>
    <v:f eqn=""if @0 21600 @1""/><v:f eqn=""if @0 0 @2""/><v:f eqn=""if @0 @4 21600""/>
    <v:f eqn=""mid @5 @6""/><v:f eqn=""mid @8 @5""/><v:f eqn=""mid @7 @8""/>
    <v:f eqn=""mid @6 @7""/><v:f eqn=""sum @6 0 @5""/>
  </v:formulas>
  <v:path textpathok=""t"" o:connecttype=""custom"" o:connectlocs=""@9,0;@10,10800;@11,21600;@12,10800"" o:connectangles=""270,180,90,0""/>
  <v:textpath on=""t"" fitshape=""t""/>
  <v:handles><v:h position=""#0,bottomRight"" xrange=""6629,14971""/></v:handles>
  <o:lock v:ext=""edit"" text=""t"" shapetype=""t""/>
</v:shapetype>
<v:shape id=""PowerPlusWaterMarkObject{wmIdx}"" o:spid=""_x0000_s102{4 + wmIdx}"" type=""#_x0000_t136"" style=""position:absolute;margin-left:0;margin-top:0;width:{wmWidth};height:{wmHeight};rotation:{wmRotation};z-index:-251654144;mso-wrap-edited:f;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin"" o:allowincell=""f"" fillcolor=""{wmColor}"" stroked=""f"" xmlns:v=""urn:schemas-microsoft-com:vml"" xmlns:o=""urn:schemas-microsoft-com:office:office"">
  <v:fill opacity=""{wmOpacity}""/>
  <v:textpath style=""font-family:&quot;{System.Security.SecurityElement.Escape(wmFont)}&quot;;font-size:{wmSize}"" string=""{System.Security.SecurityElement.Escape(wmText)}""/>
</v:shape>";

                    // Build header XML with SDT wrapper (docPartGallery=Watermarks)
                    var headerXml = $@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<w:hdr xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main""
       xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""
       xmlns:w10=""urn:schemas-microsoft-com:office:word"">
  <w:sdt>
    <w:sdtPr>
      <w:id w:val=""{-1000 - wmIdx}""/>
      <w:docPartObj>
        <w:docPartGallery w:val=""Watermarks""/>
        <w:docPartUnique/>
      </w:docPartObj>
    </w:sdtPr>
    <w:sdtContent>
      <w:p>
        <w:pPr><w:pStyle w:val=""Header""/></w:pPr>
        <w:r>
          <w:rPr><w:noProof/></w:rPr>
          <w:pict>{vmlXml}</w:pict>
        </w:r>
      </w:p>
    </w:sdtContent>
  </w:sdt>
</w:hdr>";

                    using (var stream = wmHeaderPart.GetStream(System.IO.FileMode.Create))
                    using (var writer = new System.IO.StreamWriter(stream, System.Text.Encoding.UTF8))
                        writer.Write(headerXml);

                    // Link header to section properties
                    var wmBody = mainPartWM.Document!.Body!;
                    var wmSectPr = wmBody.Elements<SectionProperties>().LastOrDefault()
                        ?? wmBody.AppendChild(new SectionProperties());

                    // Remove existing header reference of same type
                    var existingRef = wmSectPr.Elements<HeaderReference>()
                        .FirstOrDefault(r => r.Type?.Value == headerTypes[wi]);
                    existingRef?.Remove();

                    wmSectPr.PrependChild(new HeaderReference
                    {
                        Id = mainPartWM.GetIdOfPart(wmHeaderPart),
                        Type = headerTypes[wi]
                    });
                }

                // Enable even/odd page headers and title page
                var wmSettingsPart = mainPartWM.DocumentSettingsPart
                    ?? mainPartWM.AddNewPart<DocumentSettingsPart>();
                wmSettingsPart.Settings ??= new Settings();
                if (wmSettingsPart.Settings.GetFirstChild<EvenAndOddHeaders>() == null)
                    wmSettingsPart.Settings.AppendChild(new EvenAndOddHeaders());
                if (wmSettingsPart.Settings.GetFirstChild<TitlePage>() == null)
                    wmSettingsPart.Settings.AppendChild(new TitlePage());
                wmSettingsPart.Settings.Save();

                mainPartWM.Document.Save();
                return "/watermark";
            }

            default:
            {
                // Generic fallback: create typed element via SDK schema validation
                var created = GenericXmlQuery.TryCreateTypedElement(parent, type, properties, index);
                if (created == null)
                    throw new ArgumentException($"Schema-invalid element type '{type}' for parent '{parentPath}'. " +
                        "Use raw-set --action append with explicit XML instead.");

                newElement = created;
                var siblings = parent.ChildElements.Where(e => e.LocalName == created.LocalName).ToList();
                var createdIdx = siblings.IndexOf(created) + 1;
                resultPath = $"{parentPath}/{created.LocalName}[{createdIdx}]";
                break;
            }
        }

        _doc.MainDocumentPart?.Document?.Save();
        return resultPath;
    }

    public (string RelId, string PartPath) AddPart(string parentPartPath, string partType, Dictionary<string, string>? properties = null)
    {
        var mainPart = _doc.MainDocumentPart!;

        switch (partType.ToLowerInvariant())
        {
            case "chart":
                var chartPart = mainPart.AddNewPart<ChartPart>();
                var relId = mainPart.GetIdOfPart(chartPart);
                // Initialize with minimal valid ChartSpace
                chartPart.ChartSpace = new C.ChartSpace(
                    new C.Chart(new C.PlotArea(new C.Layout()))
                );
                chartPart.ChartSpace.Save();
                var chartIdx = mainPart.ChartParts.ToList().IndexOf(chartPart);
                return (relId, $"/chart[{chartIdx + 1}]");

            case "header":
                var headerPart = mainPart.AddNewPart<HeaderPart>();
                var hRelId = mainPart.GetIdOfPart(headerPart);
                headerPart.Header = new Header(new Paragraph());
                headerPart.Header.Save();
                var hIdx = mainPart.HeaderParts.ToList().IndexOf(headerPart);
                return (hRelId, $"/header[{hIdx + 1}]");

            case "footer":
                var footerPart = mainPart.AddNewPart<FooterPart>();
                var fRelId = mainPart.GetIdOfPart(footerPart);
                footerPart.Footer = new Footer(new Paragraph());
                footerPart.Footer.Save();
                var fIdx = mainPart.FooterParts.ToList().IndexOf(footerPart);
                return (fRelId, $"/footer[{fIdx + 1}]");

            default:
                throw new ArgumentException(
                    $"Unknown part type: {partType}. Supported: chart, header, footer");
        }
    }

    public void Remove(string path)
    {
        // Handle /watermark removal
        if (path.Equals("/watermark", StringComparison.OrdinalIgnoreCase))
        {
            RemoveWatermarkHeaders();
            _doc.MainDocumentPart?.Document?.Save();
            return;
        }

        var parts = ParsePath(path);
        var element = NavigateToElement(parts)
            ?? throw new ArgumentException($"Path not found: {path}");

        element.Remove();
        _doc.MainDocumentPart?.Document?.Save();
    }

    public string Move(string sourcePath, string? targetParentPath, int? index)
    {
        var srcParts = ParsePath(sourcePath);
        var element = NavigateToElement(srcParts)
            ?? throw new ArgumentException($"Source not found: {sourcePath}");

        // Determine target parent
        string effectiveParentPath;
        OpenXmlElement targetParent;
        if (string.IsNullOrEmpty(targetParentPath))
        {
            // Reorder within current parent
            targetParent = element.Parent
                ?? throw new InvalidOperationException("Element has no parent");
            // Compute parent path by removing last segment
            var lastSlash = sourcePath.LastIndexOf('/');
            effectiveParentPath = lastSlash > 0 ? sourcePath[..lastSlash] : "/body";
        }
        else
        {
            effectiveParentPath = targetParentPath;
            if (targetParentPath is "/" or "" or "/body")
                targetParent = _doc.MainDocumentPart!.Document!.Body!;
            else
            {
                var tgtParts = ParsePath(targetParentPath);
                targetParent = NavigateToElement(tgtParts)
                    ?? throw new ArgumentException($"Target parent not found: {targetParentPath}");
            }
        }

        element.Remove();

        // Insert at the specified position among same-type siblings (0-based index)
        if (index.HasValue)
        {
            var sameTypeSiblings = targetParent.ChildElements
                .Where(e => e.LocalName == element.LocalName).ToList();
            if (index.Value >= 0 && index.Value < sameTypeSiblings.Count)
                sameTypeSiblings[index.Value].InsertBeforeSelf(element);
            else
                AppendToParent(targetParent, element);
        }
        else
        {
            targetParent.AppendChild(element);
        }

        _doc.MainDocumentPart?.Document?.Save();

        var siblings = targetParent.ChildElements.Where(e => e.LocalName == element.LocalName).ToList();
        var newIdx = siblings.IndexOf(element) + 1;
        return $"{effectiveParentPath}/{element.LocalName}[{newIdx}]";
    }

    public (string NewPath1, string NewPath2) Swap(string path1, string path2)
    {
        var parts1 = ParsePath(path1);
        var elem1 = NavigateToElement(parts1)
            ?? throw new ArgumentException($"Element not found: {path1}");
        var parts2 = ParsePath(path2);
        var elem2 = NavigateToElement(parts2)
            ?? throw new ArgumentException($"Element not found: {path2}");

        if (elem1.Parent != elem2.Parent)
            throw new ArgumentException("Cannot swap elements with different parents");

        PowerPointHandler.SwapXmlElements(elem1, elem2);
        _doc.MainDocumentPart?.Document?.Save();

        // Recompute paths
        var parent = elem1.Parent!;
        var lastSlash = path1.LastIndexOf('/');
        var parentPath = lastSlash > 0 ? path1[..lastSlash] : "/body";

        var siblings1 = parent.ChildElements.Where(e => e.LocalName == elem1.LocalName).ToList();
        var newIdx1 = siblings1.IndexOf(elem1) + 1;
        var siblings2 = parent.ChildElements.Where(e => e.LocalName == elem2.LocalName).ToList();
        var newIdx2 = siblings2.IndexOf(elem2) + 1;
        return ($"{parentPath}/{elem1.LocalName}[{newIdx1}]", $"{parentPath}/{elem2.LocalName}[{newIdx2}]");
    }

    public string CopyFrom(string sourcePath, string targetParentPath, int? index)
    {
        var srcParts = ParsePath(sourcePath);
        var element = NavigateToElement(srcParts)
            ?? throw new ArgumentException($"Source not found: {sourcePath}");

        var clone = element.CloneNode(true);

        OpenXmlElement targetParent;
        if (targetParentPath is "/" or "" or "/body")
            targetParent = _doc.MainDocumentPart!.Document!.Body!;
        else
        {
            var tgtParts = ParsePath(targetParentPath);
            targetParent = NavigateToElement(tgtParts)
                ?? throw new ArgumentException($"Target parent not found: {targetParentPath}");
        }

        InsertAtPosition(targetParent, clone, index);

        _doc.MainDocumentPart?.Document?.Save();

        var siblings = targetParent.ChildElements.Where(e => e.LocalName == clone.LocalName).ToList();
        var newIdx = siblings.IndexOf(clone) + 1;
        return $"{targetParentPath}/{clone.LocalName}[{newIdx}]";
    }

    private static void InsertAtPosition(OpenXmlElement parent, OpenXmlElement element, int? index)
    {
        if (index.HasValue)
        {
            var children = parent.ChildElements.ToList();
            if (index.Value >= 0 && index.Value < children.Count)
                children[index.Value].InsertBeforeSelf(element);
            else
                parent.AppendChild(element);
        }
        else
        {
            parent.AppendChild(element);
        }
    }

    private void SetDocumentProperties(Dictionary<string, string> properties)
    {
        var doc = _doc.MainDocumentPart?.Document
            ?? throw new InvalidOperationException("Document not found");

        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "pagebackground" or "background":
                    doc.DocumentBackground = new DocumentBackground { Color = value };
                    // Enable background display in settings
                    var settingsPart = _doc.MainDocumentPart!.DocumentSettingsPart
                        ?? _doc.MainDocumentPart.AddNewPart<DocumentSettingsPart>();
                    settingsPart.Settings ??= new Settings();
                    if (settingsPart.Settings.GetFirstChild<DisplayBackgroundShape>() == null)
                        settingsPart.Settings.AppendChild(new DisplayBackgroundShape());
                    settingsPart.Settings.Save();
                    break;

                case "defaultfont":
                    var stylesPart = _doc.MainDocumentPart!.StyleDefinitionsPart;
                    if (stylesPart?.Styles != null)
                    {
                        var defaultRunProps = stylesPart.Styles.DocDefaults?.RunPropertiesDefault?.RunPropertiesBaseStyle;
                        if (defaultRunProps != null)
                        {
                            var fonts = defaultRunProps.GetFirstChild<RunFonts>()
                                ?? defaultRunProps.AppendChild(new RunFonts());
                            fonts.Ascii = value;
                            fonts.HighAnsi = value;
                            fonts.EastAsia = value;
                            stylesPart.Styles.Save();
                        }
                    }
                    break;

                case "pagewidth":
                    EnsureSectionProperties().GetFirstChild<PageSize>()!.Width = ParseHelpers.SafeParseUint(value, "pagewidth");
                    break;
                case "pageheight":
                    EnsureSectionProperties().GetFirstChild<PageSize>()!.Height = ParseHelpers.SafeParseUint(value, "pageheight");
                    break;
                case "margintop":
                    EnsurePageMargin().Top = ParseHelpers.SafeParseInt(value, "margintop");
                    break;
                case "marginbottom":
                    EnsurePageMargin().Bottom = ParseHelpers.SafeParseInt(value, "marginbottom");
                    break;
                case "marginleft":
                    EnsurePageMargin().Left = ParseHelpers.SafeParseUint(value, "marginleft");
                    break;
                case "marginright":
                    EnsurePageMargin().Right = ParseHelpers.SafeParseUint(value, "marginright");
                    break;

                // Core document properties
                case "title":
                    _doc.PackageProperties.Title = value;
                    break;
                case "author" or "creator":
                    _doc.PackageProperties.Creator = value;
                    break;
                case "subject":
                    _doc.PackageProperties.Subject = value;
                    break;
                case "keywords":
                    _doc.PackageProperties.Keywords = value;
                    break;
                case "description":
                    _doc.PackageProperties.Description = value;
                    break;
                case "category":
                    _doc.PackageProperties.Category = value;
                    break;
                case "lastmodifiedby":
                    _doc.PackageProperties.LastModifiedBy = value;
                    break;
                case "revision":
                    _doc.PackageProperties.Revision = value;
                    break;
            }
        }
    }

    private SectionProperties EnsureSectionProperties()
    {
        var body = _doc.MainDocumentPart!.Document!.Body!;
        var sectPr = body.GetFirstChild<SectionProperties>();
        if (sectPr == null)
        {
            sectPr = new SectionProperties();
            body.AppendChild(sectPr);
        }
        if (sectPr.GetFirstChild<PageSize>() == null)
            sectPr.AppendChild(new PageSize { Width = 11906, Height = 16838 }); // A4 default
        return sectPr;
    }

    private PageMargin EnsurePageMargin()
    {
        var sectPr = EnsureSectionProperties();
        var margin = sectPr.GetFirstChild<PageMargin>();
        if (margin == null)
        {
            margin = new PageMargin { Top = 1440, Bottom = 1440, Left = 1800, Right = 1800 };
            sectPr.AppendChild(margin);
        }
        return margin;
    }
}
