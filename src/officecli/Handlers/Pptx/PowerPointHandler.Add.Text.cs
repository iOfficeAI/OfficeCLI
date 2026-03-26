// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler
{
    private string AddEquation(string parentPath, int? index, Dictionary<string, string> properties)
    {
                if (!properties.TryGetValue("formula", out var eqFormula))
                    throw new ArgumentException("'formula' property is required for equation type");

                var eqSlideMatch = Regex.Match(parentPath, @"^/slide\[(\d+)\]$");
                if (!eqSlideMatch.Success)
                    throw new ArgumentException($"Equations must be added to a slide: /slide[N]");

                var eqSlideIdx = int.Parse(eqSlideMatch.Groups[1].Value);
                var eqSlideParts = GetSlideParts().ToList();
                if (eqSlideIdx < 1 || eqSlideIdx > eqSlideParts.Count)
                    throw new ArgumentException($"Slide {eqSlideIdx} not found (total: {eqSlideParts.Count})");

                var eqSlidePart = eqSlideParts[eqSlideIdx - 1];
                var eqShapeTree = GetSlide(eqSlidePart).CommonSlideData?.ShapeTree
                    ?? throw new InvalidOperationException("Slide has no shape tree");

                var eqShapeId = eqShapeTree.ChildElements
                    .Select(e => e.Descendants<NonVisualDrawingProperties>().FirstOrDefault()?.Id?.Value ?? 0)
                    .DefaultIfEmpty(1U)
                    .Max() + 1;
                var eqShapeName = properties.GetValueOrDefault("name", $"Equation {eqShapeId}");

                // Parse formula to OMML
                var mathContent = FormulaParser.Parse(eqFormula);
                M.OfficeMath oMath;
                if (mathContent is M.OfficeMath directMath)
                    oMath = directMath;
                else
                    oMath = new M.OfficeMath(mathContent.CloneNode(true));

                // Build the a14:m wrapper element via raw XML
                // PPT equations are embedded as: a:p > a14:m > m:oMathPara > m:oMath
                var mathPara = new M.Paragraph(oMath);
                var a14mXml = $"<a14:m xmlns:a14=\"http://schemas.microsoft.com/office/drawing/2010/main\">{mathPara.OuterXml}</a14:m>";

                // Create shape with equation paragraph
                var eqShape = new Shape();
                eqShape.NonVisualShapeProperties = new NonVisualShapeProperties(
                    new NonVisualDrawingProperties { Id = eqShapeId, Name = eqShapeName },
                    new NonVisualShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()
                );
                var eqSpPr = new ShapeProperties();
                {
                    long eqX = 838200, eqY = 2743200;        // default: ~2.33cm, ~7.62cm
                    long eqCx = 10515600, eqCy = 2743200;    // default: ~29.21cm, ~7.62cm
                    if (properties.TryGetValue("x", out var exStr)) eqX = ParseEmu(exStr);
                    if (properties.TryGetValue("y", out var eyStr)) eqY = ParseEmu(eyStr);
                    if (properties.TryGetValue("width", out var ewStr)) eqCx = ParseEmu(ewStr);
                    if (properties.TryGetValue("height", out var ehStr)) eqCy = ParseEmu(ehStr);
                    eqSpPr.Transform2D = new Drawing.Transform2D
                    {
                        Offset = new Drawing.Offset { X = eqX, Y = eqY },
                        Extents = new Drawing.Extents { Cx = eqCx, Cy = eqCy }
                    };
                }
                eqShape.ShapeProperties = eqSpPr;

                // Create text body with math paragraph
                var bodyProps = new Drawing.BodyProperties();
                var listStyle = new Drawing.ListStyle();
                var drawingPara = new Drawing.Paragraph();

                // Build mc:AlternateContent > mc:Choice(Requires="a14") > a14:m > m:oMathPara
                var a14mElement = new OpenXmlUnknownElement("a14", "m", "http://schemas.microsoft.com/office/drawing/2010/main");
                a14mElement.AppendChild(mathPara.CloneNode(true));

                var choice = new AlternateContentChoice();
                choice.Requires = "a14";
                choice.AppendChild(a14mElement);

                // Fallback: readable text for older versions
                var fallback = new AlternateContentFallback();
                var fallbackRun = new Drawing.Run(
                    new Drawing.RunProperties { Language = "en-US" },
                    new Drawing.Text(FormulaParser.ToReadableText(mathPara))
                );
                fallback.AppendChild(fallbackRun);

                var altContent = new AlternateContent();
                altContent.AppendChild(choice);
                altContent.AppendChild(fallback);
                drawingPara.AppendChild(altContent);

                eqShape.TextBody = new TextBody(bodyProps, listStyle, drawingPara);
                eqShapeTree.AppendChild(eqShape);

                // Ensure slide root has xmlns:a14 and mc:Ignorable="a14" so PowerPoint accepts the equation
                var eqSlide = GetSlide(eqSlidePart);
                if (eqSlide.LookupNamespace("a14") == null)
                    eqSlide.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");
                if (eqSlide.LookupNamespace("mc") == null)
                    eqSlide.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
                var currentIgnorable = eqSlide.MCAttributes?.Ignorable?.Value ?? "";
                if (!currentIgnorable.Contains("a14"))
                {
                    var newVal = string.IsNullOrEmpty(currentIgnorable) ? "a14" : $"{currentIgnorable} a14";
                    eqSlide.MCAttributes = new MarkupCompatibilityAttributes { Ignorable = newVal };
                }
                eqSlide.Save();

                var eqShapeCount = eqShapeTree.Elements<Shape>().Count();
                return $"/slide[{eqSlideIdx}]/shape[{eqShapeCount}]";
    }


    private string AddNotes(string parentPath, int? index, Dictionary<string, string> properties)
    {
                var notesSlideMatch = Regex.Match(parentPath, @"^/slide\[(\d+)\]$");
                if (!notesSlideMatch.Success)
                    throw new ArgumentException("Notes must be added to a slide: /slide[N]");
                var notesSlideIdx = int.Parse(notesSlideMatch.Groups[1].Value);
                var notesSlideParts = GetSlideParts().ToList();
                if (notesSlideIdx < 1 || notesSlideIdx > notesSlideParts.Count)
                    throw new ArgumentException($"Slide {notesSlideIdx} not found (total: {notesSlideParts.Count})");
                var notesSlidePart = EnsureNotesSlidePart(notesSlideParts[notesSlideIdx - 1]);
                if (properties.TryGetValue("text", out var notesText))
                    SetNotesText(notesSlidePart, notesText);
                return $"/slide[{notesSlideIdx}]/notes";
    }


    private string AddParagraph(string parentPath, int? index, Dictionary<string, string> properties)
    {
                // Add a paragraph to an existing shape: /slide[N]/shape[M]
                var paraParentMatch = Regex.Match(parentPath, @"^/slide\[(\d+)\]/shape\[(\d+)\]$");
                if (!paraParentMatch.Success)
                    throw new ArgumentException("Paragraphs must be added to a shape: /slide[N]/shape[M]");

                var paraSlideIdx = int.Parse(paraParentMatch.Groups[1].Value);
                var paraShapeIdx = int.Parse(paraParentMatch.Groups[2].Value);
                var (paraSlidePart, paraShape) = ResolveShape(paraSlideIdx, paraShapeIdx);

                var textBody = paraShape.TextBody
                    ?? throw new InvalidOperationException("Shape has no text body");

                var newPara = new Drawing.Paragraph();
                var pProps = new Drawing.ParagraphProperties();

                // Paragraph-level properties
                if (properties.TryGetValue("align", out var pAlign))
                    pProps.Alignment = ParseTextAlignment(pAlign);
                if (properties.TryGetValue("indent", out var pIndent))
                    pProps.Indent = (int)ParseEmu(pIndent);
                if (properties.TryGetValue("marginLeft", out var pMarL) || properties.TryGetValue("marl", out pMarL))
                    pProps.LeftMargin = (int)ParseEmu(pMarL);
                if (properties.TryGetValue("marginRight", out var pMarR) || properties.TryGetValue("marr", out pMarR))
                    pProps.RightMargin = (int)ParseEmu(pMarR);
                if (properties.TryGetValue("list", out var pList) || properties.TryGetValue("liststyle", out pList))
                    ApplyListStyle(pProps, pList);

                newPara.ParagraphProperties = pProps;

                // Create initial run with text and run-level properties
                var paraText = properties.GetValueOrDefault("text", "");
                var newRun = new Drawing.Run();
                var rProps = new Drawing.RunProperties { Language = "en-US" };

                if (properties.TryGetValue("size", out var pSize))
                    rProps.FontSize = (int)Math.Round(ParseFontSize(pSize) * 100);
                if (properties.TryGetValue("bold", out var pBold))
                    rProps.Bold = IsTruthy(pBold);
                if (properties.TryGetValue("italic", out var pItalic))
                    rProps.Italic = IsTruthy(pItalic);
                // Schema order: solidFill before latin/ea
                if (properties.TryGetValue("color", out var pColor))
                    rProps.AppendChild(BuildSolidFill(pColor));
                if (properties.TryGetValue("font", out var pFont))
                {
                    rProps.Append(new Drawing.LatinFont { Typeface = pFont });
                    rProps.Append(new Drawing.EastAsianFont { Typeface = pFont });
                }
                if (properties.TryGetValue("spacing", out var pSpacing) || properties.TryGetValue("charspacing", out pSpacing))
                {
                    if (!double.TryParse(pSpacing, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var pSpcVal))
                        throw new ArgumentException($"Invalid 'spacing' value: '{pSpacing}'. Expected a number in points.");
                    rProps.Spacing = (int)(pSpcVal * 100);
                }
                if (properties.TryGetValue("baseline", out var pBaseline))
                {
                    rProps.Baseline = pBaseline.ToLowerInvariant() switch
                    {
                        "super" or "true" => 30000,
                        "sub" => -25000,
                        _ => double.TryParse(pBaseline, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var pBlVal) && !double.IsNaN(pBlVal) && !double.IsInfinity(pBlVal)
                            ? (int)(pBlVal * 1000)
                            : throw new ArgumentException($"Invalid 'baseline' value: '{pBaseline}'. Expected 'super', 'sub', or a percentage.")
                    };
                }

                newRun.RunProperties = rProps;
                newRun.Text = new Drawing.Text(paraText.Replace("\\n", "\n"));
                newPara.Append(newRun);

                if (index.HasValue && index.Value >= 0)
                {
                    var existingParas = textBody.Elements<Drawing.Paragraph>().ToList();
                    if (index.Value < existingParas.Count)
                        textBody.InsertBefore(newPara, existingParas[index.Value]);
                    else
                        textBody.Append(newPara);
                }
                else
                {
                    textBody.Append(newPara);
                }

                var paraCount = textBody.Elements<Drawing.Paragraph>().Count();
                GetSlide(paraSlidePart).Save();
                return $"/slide[{paraSlideIdx}]/shape[{paraShapeIdx}]/paragraph[{paraCount}]";
    }


    private string AddRun(string parentPath, int? index, Dictionary<string, string> properties)
    {
                // Add a run to a paragraph: /slide[N]/shape[M]/paragraph[P] or /slide[N]/shape[M]
                var runParaMatch = Regex.Match(parentPath, @"^/slide\[(\d+)\]/shape\[(\d+)\](?:/paragraph\[(\d+)\])?$");
                if (!runParaMatch.Success)
                    throw new ArgumentException("Runs must be added to a shape or paragraph: /slide[N]/shape[M] or /slide[N]/shape[M]/paragraph[P]");

                var runSlideIdx = int.Parse(runParaMatch.Groups[1].Value);
                var runShapeIdx = int.Parse(runParaMatch.Groups[2].Value);
                var (runSlidePart, runShape) = ResolveShape(runSlideIdx, runShapeIdx);

                var runTextBody = runShape.TextBody
                    ?? throw new InvalidOperationException("Shape has no text body");

                Drawing.Paragraph targetPara;
                int targetParaIdx;
                if (runParaMatch.Groups[3].Success)
                {
                    targetParaIdx = int.Parse(runParaMatch.Groups[3].Value);
                    var paras = runTextBody.Elements<Drawing.Paragraph>().ToList();
                    if (targetParaIdx < 1 || targetParaIdx > paras.Count)
                        throw new ArgumentException($"Paragraph {targetParaIdx} not found");
                    targetPara = paras[targetParaIdx - 1];
                }
                else
                {
                    // Append to last paragraph
                    var paras = runTextBody.Elements<Drawing.Paragraph>().ToList();
                    targetPara = paras.LastOrDefault()
                        ?? throw new InvalidOperationException("Shape has no paragraphs");
                    targetParaIdx = paras.Count;
                }

                var runText = properties.GetValueOrDefault("text", "");
                var newRun = new Drawing.Run();
                var rProps = new Drawing.RunProperties { Language = "en-US" };

                if (properties.TryGetValue("size", out var rSize))
                    rProps.FontSize = (int)Math.Round(ParseFontSize(rSize) * 100);
                if (properties.TryGetValue("bold", out var rBold))
                    rProps.Bold = IsTruthy(rBold);
                if (properties.TryGetValue("italic", out var rItalic))
                    rProps.Italic = IsTruthy(rItalic);
                if (properties.TryGetValue("underline", out var rUnderline))
                    rProps.Underline = rUnderline.ToLowerInvariant() switch
                    {
                        "true" or "single" or "sng" => Drawing.TextUnderlineValues.Single,
                        "double" or "dbl" => Drawing.TextUnderlineValues.Double,
                        "heavy" => Drawing.TextUnderlineValues.Heavy,
                        "dotted" => Drawing.TextUnderlineValues.Dotted,
                        "dash" => Drawing.TextUnderlineValues.Dash,
                        "wavy" => Drawing.TextUnderlineValues.Wavy,
                        "false" or "none" => Drawing.TextUnderlineValues.None,
                        _ => throw new ArgumentException($"Invalid underline value: '{rUnderline}'. Valid values: single, double, heavy, dotted, dash, wavy, none.")
                    };
                if (properties.TryGetValue("strikethrough", out var rStrike) || properties.TryGetValue("strike", out rStrike))
                    rProps.Strike = rStrike.ToLowerInvariant() switch
                    {
                        "true" or "single" => Drawing.TextStrikeValues.SingleStrike,
                        "double" => Drawing.TextStrikeValues.DoubleStrike,
                        "false" or "none" => Drawing.TextStrikeValues.NoStrike,
                        _ => throw new ArgumentException($"Invalid strikethrough value: '{rStrike}'. Valid values: single, double, none.")
                    };
                // Schema order: solidFill before latin/ea
                if (properties.TryGetValue("color", out var rColor))
                    rProps.AppendChild(BuildSolidFill(rColor));
                if (properties.TryGetValue("font", out var rFont))
                {
                    rProps.Append(new Drawing.LatinFont { Typeface = rFont });
                    rProps.Append(new Drawing.EastAsianFont { Typeface = rFont });
                }
                if (properties.TryGetValue("spacing", out var rSpacing) || properties.TryGetValue("charspacing", out rSpacing))
                    rProps.Spacing = (int)(ParseHelpers.SafeParseDouble(rSpacing, "charspacing") * 100);
                if (properties.TryGetValue("baseline", out var rBaseline))
                {
                    rProps.Baseline = rBaseline.ToLowerInvariant() switch
                    {
                        "super" or "true" => 30000,
                        "sub" => -25000,
                        "none" or "false" or "0" => 0,
                        _ => (int)(ParseHelpers.SafeParseDouble(rBaseline, "baseline") * 1000)
                    };
                }
                else if (properties.TryGetValue("superscript", out var rSuper))
                    rProps.Baseline = IsTruthy(rSuper) ? 30000 : 0;
                else if (properties.TryGetValue("subscript", out var rSub))
                    rProps.Baseline = IsTruthy(rSub) ? -25000 : 0;

                newRun.RunProperties = rProps;
                newRun.Text = new Drawing.Text(runText.Replace("\\n", "\n"));

                // Append run to paragraph (before EndParagraphRunProperties if present)
                var endParaRun = targetPara.GetFirstChild<Drawing.EndParagraphRunProperties>();
                if (endParaRun != null)
                    targetPara.InsertBefore(newRun, endParaRun);
                else
                    targetPara.Append(newRun);

                var runCount = targetPara.Elements<Drawing.Run>().Count();
                GetSlide(runSlidePart).Save();
                return $"/slide[{runSlideIdx}]/shape[{runShapeIdx}]/paragraph[{targetParaIdx}]/run[{runCount}]";
    }


}
