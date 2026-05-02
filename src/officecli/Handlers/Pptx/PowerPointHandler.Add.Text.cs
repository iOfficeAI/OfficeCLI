// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
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
                if (!properties.TryGetValue("formula", out var eqFormula) && !properties.TryGetValue("text", out eqFormula))
                    throw new ArgumentException("'formula' (or 'text') property is required for equation type");

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

                var eqShapeId = AcquireShapeId(eqShapeTree, properties);
                var eqShapeName = properties.GetValueOrDefault("name", $"Equation {eqShapeTree.Elements<Shape>().Count() + 1}");

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
                    new Drawing.Text { Text = FormulaParser.ToReadableText(mathPara) }
                );
                fallback.AppendChild(fallbackRun);

                var altContent = new AlternateContent();
                altContent.AppendChild(choice);
                altContent.AppendChild(fallback);
                drawingPara.AppendChild(altContent);

                eqShape.TextBody = new TextBody(bodyProps, listStyle, drawingPara);
                InsertAtPosition(eqShapeTree, eqShape, index);

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

                return $"/slide[{eqSlideIdx}]/{BuildElementPathSegment("shape", eqShape, eqShapeTree.Elements<Shape>().Count())}";
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
                {
                    XmlTextValidator.ValidateOrThrow(notesText, "text");
                    SetNotesText(notesSlidePart, notesText);
                }
                // Reading direction (Arabic / Hebrew speaker notes). Mirrors
                // the AddShape direction handling — must run after SetNotesText
                // so the paragraphs it creates pick up rtl=1.
                if (properties.TryGetValue("direction", out var notesDir)
                    || properties.TryGetValue("dir", out notesDir)
                    || properties.TryGetValue("rtl", out notesDir))
                {
                    ApplyNotesDirection(notesSlidePart, notesDir);
                    notesSlidePart.NotesSlide!.Save();
                }
                // CONSISTENCY(add-set-symmetry): notes Set accepts lang=
                // (routes through SetRunOrShapeProperties on the notes
                // body). Add must accept the same key — without this,
                // `add /slide[N] --type notes --prop lang=ar-SA` reported
                // UNSUPPORTED while Set succeeded.
                if (properties.TryGetValue("lang", out var notesLang))
                {
                    Shape? notesBody = null;
                    var notesShapeTree = notesSlidePart.NotesSlide?.CommonSlideData?.ShapeTree;
                    if (notesShapeTree != null)
                    {
                        foreach (var sh in notesShapeTree.Elements<Shape>())
                        {
                            var ph = sh.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.GetFirstChild<PlaceholderShape>();
                            if (ph?.Index?.Value == 1) { notesBody = sh; break; }
                        }
                    }
                    if (notesBody != null)
                    {
                        var notesRuns = notesBody.Descendants<Drawing.Run>().ToList();
                        SetRunOrShapeProperties(new Dictionary<string, string> { ["lang"] = notesLang }, notesRuns, notesBody);
                        notesSlidePart.NotesSlide!.Save();
                    }
                }
                return $"/slide[{notesSlideIdx}]/notes";
    }


    private string AddParagraph(string parentPath, int? index, Dictionary<string, string> properties)
    {
                // Add a paragraph to an existing shape or placeholder:
                //   /slide[N]/shape[M] or /slide[N]/placeholder[X]
                // CONSISTENCY(placeholder-paragraph-path): same dual-route the
                // Set side ships at PowerPointHandler.Set.Shape.cs, so dump
                // emit can target either form via positional ordinals.
                var paraParentMatch = Regex.Match(parentPath, @"^/slide\[(\d+)\]/shape\[(\d+)\]$");
                var paraPhMatch = paraParentMatch.Success ? null : Regex.Match(parentPath, @"^/slide\[(\d+)\]/placeholder\[(\w+)\]$");
                if (!paraParentMatch.Success && (paraPhMatch == null || !paraPhMatch.Success))
                    throw new ArgumentException("Paragraphs must be added to a shape or placeholder: /slide[N]/shape[M] or /slide[N]/placeholder[X]");

                SlidePart paraSlidePart;
                Shape paraShape;
                int paraSlideIdx;
                int paraShapeIdx;
                if (paraParentMatch.Success)
                {
                    paraSlideIdx = int.Parse(paraParentMatch.Groups[1].Value);
                    paraShapeIdx = int.Parse(paraParentMatch.Groups[2].Value);
                    (paraSlidePart, paraShape) = ResolveShape(paraSlideIdx, paraShapeIdx);
                }
                else
                {
                    paraSlideIdx = int.Parse(paraPhMatch!.Groups[1].Value);
                    var phToken = paraPhMatch.Groups[2].Value;
                    var slideParts = GetSlideParts().ToList();
                    if (paraSlideIdx < 1 || paraSlideIdx > slideParts.Count)
                        throw new ArgumentException($"Slide {paraSlideIdx} not found (total: {slideParts.Count})");
                    paraSlidePart = slideParts[paraSlideIdx - 1];
                    paraShape = ResolvePlaceholderShape(paraSlidePart, phToken);
                    // Synthetic index for return path — placeholder positional
                    // lookup happens by Set's path resolver, this is informational.
                    paraShapeIdx = 1;
                }

                var textBody = paraShape.TextBody
                    ?? throw new InvalidOperationException("Shape has no text body");

                var newPara = new Drawing.Paragraph();
                var pProps = new Drawing.ParagraphProperties();

                // Paragraph-level properties
                if (properties.TryGetValue("align", out var pAlign))
                    pProps.Alignment = ParseTextAlignment(pAlign);
                if (properties.TryGetValue("indent", out var pIndent))
                {
                    // CONSISTENCY(pptx-bare-as-points): paragraph-level
                    // length inputs treat bare numbers as points (see
                    // spaceBefore/spaceAfter via SpacingConverter.ParsePptSpacing).
                    // ParseEmu("1") would return 1 raw EMU ≈ 0mm, useless.
                    // Bare "1" → 1pt → 12700 EMU; unit-qualified inputs
                    // ("0.5cm", "12pt") still go through ParseEmu.
                    pProps.Indent = (int)Math.Round(SpacingConverter.ParsePointsSigned(pIndent) * 12700.0);
                }
                if (properties.TryGetValue("marginLeft", out var pMarL) || properties.TryGetValue("marl", out pMarL))
                    pProps.LeftMargin = (int)Math.Round(SpacingConverter.ParsePointsSigned(pMarL) * 12700.0);
                if (properties.TryGetValue("marginRight", out var pMarR) || properties.TryGetValue("marr", out pMarR))
                    pProps.RightMargin = (int)Math.Round(SpacingConverter.ParsePointsSigned(pMarR) * 12700.0);
                if (properties.TryGetValue("list", out var pList) || properties.TryGetValue("liststyle", out pList))
                    ApplyListStyle(pProps, pList);
                if (properties.TryGetValue("level", out var pLevelStr))
                {
                    if (!int.TryParse(pLevelStr, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out var pLevelVal) || pLevelVal < 0 || pLevelVal > 8)
                        throw new ArgumentException($"Invalid 'level' value: '{pLevelStr}'. Expected an integer between 0 and 8 (OOXML a:pPr/@lvl).");
                    pProps.Level = pLevelVal;
                }
                // Line spacing (CONSISTENCY(lineSpacing): same idiom as AddShape:~180)
                if (properties.TryGetValue("lineSpacing", out var pLsVal) || properties.TryGetValue("linespacing", out pLsVal))
                {
                    var (pLsInternal, pLsIsPercent) = SpacingConverter.ParsePptLineSpacing(pLsVal);
                    pProps.RemoveAllChildren<Drawing.LineSpacing>();
                    if (pLsIsPercent)
                        pProps.AppendChild(new Drawing.LineSpacing(
                            new Drawing.SpacingPercent { Val = pLsInternal }));
                    else
                        pProps.AppendChild(new Drawing.LineSpacing(
                            new Drawing.SpacingPoints { Val = pLsInternal }));
                }
                if (properties.TryGetValue("spaceBefore", out var pSbVal) || properties.TryGetValue("spacebefore", out pSbVal))
                {
                    pProps.RemoveAllChildren<Drawing.SpaceBefore>();
                    pProps.AppendChild(new Drawing.SpaceBefore(new Drawing.SpacingPoints { Val = SpacingConverter.ParsePptSpacing(pSbVal) }));
                }
                if (properties.TryGetValue("spaceAfter", out var pSaVal) || properties.TryGetValue("spaceafter", out pSaVal))
                {
                    pProps.RemoveAllChildren<Drawing.SpaceAfter>();
                    pProps.AppendChild(new Drawing.SpaceAfter(new Drawing.SpacingPoints { Val = SpacingConverter.ParsePptSpacing(pSaVal) }));
                }

                newPara.ParagraphProperties = pProps;

                // Create initial run with text and run-level properties
                var paraText = properties.GetValueOrDefault("text", "");
                XmlTextValidator.ValidateOrThrow(paraText, "text");
                var newRun = new Drawing.Run();
                var rProps = new Drawing.RunProperties { Language = "en-US" };
                if (properties.TryGetValue("lang", out var pLang) && !string.IsNullOrEmpty(pLang))
                    rProps.Language = pLang;
                if (properties.TryGetValue("altLang", out var pAltLang) && !string.IsNullOrEmpty(pAltLang))
                    rProps.AlternativeLanguage = pAltLang;

                if (properties.TryGetValue("size", out var pSize)
                    || properties.TryGetValue("font.size", out pSize)
                    || properties.TryGetValue("fontsize", out pSize))
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
                // CONSISTENCY(font-4-slot): Set fans out font.latin/ea/cs to
                // the matching OOXML child elements; Add must mirror so the
                // CJK/complex slots round-trip through dump-replay instead of
                // silently collapsing to the bare `font` value (or being lost).
                if (properties.TryGetValue("font.latin", out var pFontLatin))
                {
                    rProps.RemoveAllChildren<Drawing.LatinFont>();
                    rProps.Append(new Drawing.LatinFont { Typeface = pFontLatin });
                }
                if (properties.TryGetValue("font.ea", out var pFontEa)
                    || properties.TryGetValue("font.eastasia", out pFontEa)
                    || properties.TryGetValue("font.eastasian", out pFontEa))
                {
                    rProps.RemoveAllChildren<Drawing.EastAsianFont>();
                    rProps.Append(new Drawing.EastAsianFont { Typeface = pFontEa });
                }
                if (properties.TryGetValue("font.cs", out var pFontCs)
                    || properties.TryGetValue("font.complexscript", out pFontCs)
                    || properties.TryGetValue("font.complex", out pFontCs))
                {
                    rProps.RemoveAllChildren<Drawing.ComplexScriptFont>();
                    rProps.Append(new Drawing.ComplexScriptFont { Typeface = pFontCs });
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

                foreach (var segmentedRun in BuildSegmentedRuns(paraText, rProps))
                    newPara.Append(segmentedRun);

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
                return $"/slide[{paraSlideIdx}]/{BuildElementPathSegment("shape", paraShape, paraShapeIdx)}/paragraph[{paraCount}]";
    }


    /// <summary>
    /// `add --type linebreak /slide[N]/shape[M]/paragraph[K]` (also /placeholder[X]) —
    /// insert an &lt;a:br/&gt; element into the target paragraph. Mirrors AddRun's path
    /// resolution shape so /paragraph[K] suffix and /placeholder[X] alias both work.
    /// </summary>
    private string AddLineBreak(string parentPath, int? index, Dictionary<string, string> properties)
    {
        var brParaMatch = Regex.Match(parentPath, @"^/slide\[(\d+)\]/shape\[(\d+)\](?:/(?:paragraph|p)\[(\d+)\])?$");
        var brPhMatch = brParaMatch.Success ? null : Regex.Match(parentPath, @"^/slide\[(\d+)\]/placeholder\[(\w+)\](?:/(?:paragraph|p)\[(\d+)\])?$");
        if (!brParaMatch.Success && (brPhMatch == null || !brPhMatch.Success))
            throw new ArgumentException(
                "Line breaks must be added to a shape/placeholder or paragraph: " +
                "/slide[N]/shape[M], /slide[N]/placeholder[X], /slide[N]/shape[M]/paragraph[K], or /slide[N]/placeholder[X]/paragraph[K]");

        Shape brShape;
        System.Text.RegularExpressions.Group brParaGroup;
        if (brParaMatch.Success)
        {
            var slideIdx = int.Parse(brParaMatch.Groups[1].Value);
            var shapeIdx = int.Parse(brParaMatch.Groups[2].Value);
            (_, brShape) = ResolveShape(slideIdx, shapeIdx);
            brParaGroup = brParaMatch.Groups[3];
        }
        else
        {
            var slideIdx = int.Parse(brPhMatch!.Groups[1].Value);
            var phToken = brPhMatch.Groups[2].Value;
            var slideParts = GetSlideParts().ToList();
            if (slideIdx < 1 || slideIdx > slideParts.Count)
                throw new ArgumentException($"Slide {slideIdx} not found (total: {slideParts.Count})");
            brShape = ResolvePlaceholderShape(slideParts[slideIdx - 1], phToken);
            brParaGroup = brPhMatch.Groups[3];
        }

        var brTextBody = brShape.TextBody
            ?? throw new InvalidOperationException("Shape has no text body");

        Drawing.Paragraph targetPara;
        int targetParaIdx;
        var paras = brTextBody.Elements<Drawing.Paragraph>().ToList();
        if (brParaGroup.Success)
        {
            targetParaIdx = int.Parse(brParaGroup.Value);
            if (targetParaIdx < 1 || targetParaIdx > paras.Count)
                throw new ArgumentException($"Paragraph {targetParaIdx} not found");
            targetPara = paras[targetParaIdx - 1];
        }
        else
        {
            targetPara = paras.LastOrDefault()
                ?? throw new InvalidOperationException("Shape has no paragraphs");
            targetParaIdx = paras.Count;
        }

        var br = new Drawing.Break();
        if (index.HasValue)
        {
            var children = targetPara.ChildElements.ToList();
            var insertAt = Math.Max(0, Math.Min(index.Value, children.Count));
            if (insertAt >= children.Count) targetPara.AppendChild(br);
            else children[insertAt].InsertBeforeSelf(br);
        }
        else
        {
            targetPara.AppendChild(br);
        }

        var brIdx = targetPara.Elements<Drawing.Break>().ToList().FindIndex(b => ReferenceEquals(b, br)) + 1;
        return $"/slide[{(brParaMatch.Success ? brParaMatch.Groups[1].Value : brPhMatch!.Groups[1].Value)}]" +
               (brParaMatch.Success
                    ? $"/shape[{brParaMatch.Groups[2].Value}]"
                    : $"/placeholder[{brPhMatch!.Groups[2].Value}]") +
               $"/paragraph[{targetParaIdx}]/br[{brIdx}]";
    }

    private string AddRun(string parentPath, int? index, Dictionary<string, string> properties)
    {
                // Add a run to a paragraph: /slide[N]/shape[M]/paragraph[P] or /slide[N]/shape[M]
                //   also: /slide[N]/placeholder[X]/paragraph[P] or /slide[N]/placeholder[X]
                // CONSISTENCY(path-aliases): accept short-form `/p[N]` alongside `/paragraph[N]`.
                // CONSISTENCY(placeholder-paragraph-path): mirror the dual route that
                // AddParagraph and SetParagraph already accept.
                var runParaMatch = Regex.Match(parentPath, @"^/slide\[(\d+)\]/shape\[(\d+)\](?:/(?:paragraph|p)\[(\d+)\])?$");
                var runPhMatch = runParaMatch.Success ? null : Regex.Match(parentPath, @"^/slide\[(\d+)\]/placeholder\[(\w+)\](?:/(?:paragraph|p)\[(\d+)\])?$");
                if (!runParaMatch.Success && (runPhMatch == null || !runPhMatch.Success))
                    throw new ArgumentException("Runs must be added to a shape/placeholder or paragraph: /slide[N]/shape[M], /slide[N]/placeholder[X], /slide[N]/shape[M]/paragraph[P], or /slide[N]/placeholder[X]/paragraph[P]");

                SlidePart runSlidePart;
                Shape runShape;
                int runSlideIdx;
                int runShapeIdx;
                System.Text.RegularExpressions.Group paraGroup;
                if (runParaMatch.Success)
                {
                    runSlideIdx = int.Parse(runParaMatch.Groups[1].Value);
                    runShapeIdx = int.Parse(runParaMatch.Groups[2].Value);
                    (runSlidePart, runShape) = ResolveShape(runSlideIdx, runShapeIdx);
                    paraGroup = runParaMatch.Groups[3];
                }
                else
                {
                    runSlideIdx = int.Parse(runPhMatch!.Groups[1].Value);
                    var phToken = runPhMatch.Groups[2].Value;
                    var slideParts = GetSlideParts().ToList();
                    if (runSlideIdx < 1 || runSlideIdx > slideParts.Count)
                        throw new ArgumentException($"Slide {runSlideIdx} not found (total: {slideParts.Count})");
                    runSlidePart = slideParts[runSlideIdx - 1];
                    runShape = ResolvePlaceholderShape(runSlidePart, phToken);
                    runShapeIdx = 1;
                    paraGroup = runPhMatch.Groups[3];
                }

                var runTextBody = runShape.TextBody
                    ?? throw new InvalidOperationException("Shape has no text body");

                Drawing.Paragraph targetPara;
                int targetParaIdx;
                if (paraGroup.Success)
                {
                    targetParaIdx = int.Parse(paraGroup.Value);
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
                XmlTextValidator.ValidateOrThrow(runText, "text");
                var newRun = new Drawing.Run();
                var rProps = new Drawing.RunProperties { Language = "en-US" };
                if (properties.TryGetValue("lang", out var rLang) && !string.IsNullOrEmpty(rLang))
                    rProps.Language = rLang;
                if (properties.TryGetValue("altLang", out var rAltLang) && !string.IsNullOrEmpty(rAltLang))
                    rProps.AlternativeLanguage = rAltLang;

                if (properties.TryGetValue("size", out var rSize)
                    || properties.TryGetValue("font.size", out rSize)
                    || properties.TryGetValue("fontsize", out rSize))
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
                // cap on run rPr (a:rPr/@cap). Schema declares add:true; symmetric
                // with the run-context Set branch in ShapeProperties.cs. Aliases
                // allCaps / smallCaps mirror Set's normalization — boolean-truthy
                // → all/small respectively; explicit "none"/"false" clears.
                string? rCapKey =
                    properties.ContainsKey("cap") ? "cap" :
                    properties.Keys.FirstOrDefault(k =>
                        k.Equals("allCaps", StringComparison.OrdinalIgnoreCase)
                        || k.Equals("allcaps", StringComparison.OrdinalIgnoreCase)
                        || k.Equals("smallCaps", StringComparison.OrdinalIgnoreCase)
                        || k.Equals("smallcaps", StringComparison.OrdinalIgnoreCase));
                if (rCapKey != null)
                {
                    var rCapRaw = properties[rCapKey];
                    string capNorm;
                    if (rCapKey.Equals("cap", StringComparison.Ordinal))
                        capNorm = rCapRaw.ToLowerInvariant();
                    else if (rCapKey.StartsWith("allCaps", StringComparison.OrdinalIgnoreCase)
                          || rCapKey.StartsWith("allcaps", StringComparison.OrdinalIgnoreCase))
                        capNorm = (rCapRaw is "0" or "false" or "False" or "none") ? "none" : "all";
                    else
                        capNorm = (rCapRaw is "0" or "false" or "False" or "none") ? "none" : "small";

                    rProps.Capital = capNorm switch
                    {
                        "all" => Drawing.TextCapsValues.All,
                        "small" => Drawing.TextCapsValues.Small,
                        "none" => Drawing.TextCapsValues.None,
                        _ => throw new ArgumentException($"Invalid cap value: '{rCapRaw}'. Valid values: none, small, all.")
                    };
                }
                // Schema order: solidFill before latin/ea
                if (properties.TryGetValue("color", out var rColor))
                    rProps.AppendChild(BuildSolidFill(rColor));
                if (properties.TryGetValue("font", out var rFont))
                {
                    rProps.Append(new Drawing.LatinFont { Typeface = rFont });
                    rProps.Append(new Drawing.EastAsianFont { Typeface = rFont });
                }
                // CONSISTENCY(font-4-slot): mirror AddParagraph and Set for the
                // per-script font slots (font.latin / font.ea / font.cs).
                if (properties.TryGetValue("font.latin", out var rFontLatin))
                {
                    rProps.RemoveAllChildren<Drawing.LatinFont>();
                    rProps.Append(new Drawing.LatinFont { Typeface = rFontLatin });
                }
                if (properties.TryGetValue("font.ea", out var rFontEa)
                    || properties.TryGetValue("font.eastasia", out rFontEa)
                    || properties.TryGetValue("font.eastasian", out rFontEa))
                {
                    rProps.RemoveAllChildren<Drawing.EastAsianFont>();
                    rProps.Append(new Drawing.EastAsianFont { Typeface = rFontEa });
                }
                if (properties.TryGetValue("font.cs", out var rFontCs)
                    || properties.TryGetValue("font.complexscript", out rFontCs)
                    || properties.TryGetValue("font.complex", out rFontCs))
                {
                    rProps.RemoveAllChildren<Drawing.ComplexScriptFont>();
                    rProps.Append(new Drawing.ComplexScriptFont { Typeface = rFontCs });
                }
                if (properties.TryGetValue("spacing", out var rSpacing) || properties.TryGetValue("charspacing", out rSpacing))
                    rProps.Spacing = (int)(ParseHelpers.SafeParseDouble(rSpacing, "charspacing") * 100);
                // kern: raw OOXML hundredths-of-a-point integer, matches Set
                // path semantics. Validated against ST_TextNonNegativePoint
                // range [0, 400000]; symmetric with the Set rPr attribute
                // branch (PowerPointHandler.ShapeProperties.cs) so that
                // `add run kern=N` no longer reports UNSUPPORTED.
                if (properties.TryGetValue("kern", out var rKern))
                {
                    if (!int.TryParse(rKern, out var kv) || kv < 0 || kv > 400000)
                        throw new ArgumentException(
                            $"Invalid kern '{rKern}': OOXML ST_TextNonNegativePoint requires an integer in [0, 400000] (hundredths of a point).");
                    rProps.Kerning = kv;
                }
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

                // CONSISTENCY(addrun-rpr-attrs): AddShape routes these through
                // SetRunOrShapeProperties / effectKeys but AddRun has its own
                // narrower property loop. Mirror the bool/int attribute set
                // here so `--type run` round-trips the same OOXML rPr surface
                // (matches DrawingRunBoolAttrs / DrawingRunIntAttrs in
                // PowerPointHandler.ShapeProperties.cs).
                foreach (var (boolKey, attrName) in new[] {
                    ("noProof", "noProof"), ("dirty", "dirty"),
                    ("err", "err"), ("smtClean", "smtClean") })
                {
                    if (properties.TryGetValue(boolKey, out var bv))
                        rProps.SetAttribute(new DocumentFormat.OpenXml.OpenXmlAttribute(
                            "", attrName, "", IsTruthy(bv) ? "1" : "0"));
                }
                if (properties.TryGetValue("smtId", out var smtIdRaw))
                {
                    var smtIdVal = OfficeCli.Core.ParseHelpers.SafeParseInt(smtIdRaw, "smtId");
                    if (smtIdVal < 0)
                        throw new ArgumentException($"Invalid smtId '{smtIdRaw}' (must be non-negative).");
                    rProps.SetAttribute(new DocumentFormat.OpenXml.OpenXmlAttribute(
                        "", "smtId", "", smtIdVal.ToString(System.Globalization.CultureInfo.InvariantCulture)));
                }

                var insertedRuns = BuildSegmentedRuns(runText, rProps);
                if (properties.TryGetValue("link", out var rLink))
                {
                    var firstRun = insertedRuns.FirstOrDefault();
                    if (firstRun != null)
                        ApplyRunHyperlink(runSlidePart, firstRun, rLink, properties.GetValueOrDefault("tooltip"));
                }

                // Insert runs at specified index, or append
                if (index.HasValue)
                {
                    var existingRuns = targetPara.Elements<Drawing.Run>().ToList();
                    if (index.Value >= 0 && index.Value < existingRuns.Count)
                    {
                        var insertRef = existingRuns[index.Value];
                        foreach (var segmentedRun in insertedRuns)
                            insertRef.InsertBeforeSelf(segmentedRun);
                    }
                    else
                    {
                        var endParaRun2 = targetPara.GetFirstChild<Drawing.EndParagraphRunProperties>();
                        foreach (var segmentedRun in insertedRuns)
                        {
                            if (endParaRun2 != null)
                                targetPara.InsertBefore(segmentedRun, endParaRun2);
                            else
                                targetPara.Append(segmentedRun);
                        }
                    }
                }
                else
                {
                    var endParaRun = targetPara.GetFirstChild<Drawing.EndParagraphRunProperties>();
                    foreach (var segmentedRun in insertedRuns)
                    {
                        if (endParaRun != null)
                            targetPara.InsertBefore(segmentedRun, endParaRun);
                        else
                            targetPara.Append(segmentedRun);
                    }
                }

                var runCount = targetPara.Elements<Drawing.Run>().Count();
                GetSlide(runSlidePart).Save();
                return $"/slide[{runSlideIdx}]/{BuildElementPathSegment("shape", runShape, runShapeIdx)}/paragraph[{targetParaIdx}]/run[{runCount - insertedRuns.Count + 1}]";
    }

    // CONSISTENCY(escape-sequences): cross-handler convention — \t in paragraph
    // text becomes a literal U+0009 inside an <a:r><a:t> run, matching what
    // PowerPoint itself writes. An earlier implementation emitted an
    // <a:tab/> sibling via OpenXmlUnknownElement, but CT_TextParagraph in
    // the DrawingML schema does not allow <a:tab/> as a direct child of
    // <a:p> — the SDK validator and `view issues` both flagged the file.
    // Caller has already split on real '\n' chars; this helper handles real
    // '\t' chars within a single line by joining segments with a tab character
    // and emitting a single run per segment.
    internal static void AppendLineWithTabs(
        Drawing.Paragraph paragraph,
        string line,
        Func<string, Drawing.Run> runFactory)
    {
        var segments = line.Split('\t');
        for (int i = 0; i < segments.Length; i++)
        {
            if (i > 0)
            {
                // Emit the tab as its own run so the surrounding segment runs
                // keep their independent rPr (formatting on either side of the
                // tab is preserved). PowerPoint accepts a literal U+0009 inside
                // <a:t> and renders it as a tab.
                paragraph.AppendChild(runFactory("\t"));
            }
            // Always emit a run per segment (including empty) so run formatting
            // is preserved on both sides of the tab. PowerPoint tolerates empty
            // <a:r><a:t/></a:r>.
            paragraph.AppendChild(runFactory(segments[i]));
        }
    }
}
