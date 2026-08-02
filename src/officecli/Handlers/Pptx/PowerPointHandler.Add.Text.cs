// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
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

                // Parse formula to OMML. R3-fuzz-1: lenient — a too-deep/
                // unparseable formula records a warning (exit 2) and writes a
                // placeholder instead of throwing (exit 1 / whole-batch failure).
                var mathContent = FormulaParser.ParseLenient(eqFormula, LastUnrecognizedLatex);
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
                    XmlTextValidator.ValidateOrThrow(notesText, "text", allowSoftBreakChar: true);
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
                // Grouped shape parent: /slide[N]/group[K](/group[L])*/shape[M].
                // dump emits paragraph adds into shapes that live inside (possibly
                // nested) groups; without this they hit the slide-only matcher and
                // failed "Element not found". Mirrors the Set-side nested-group
                // shape resolution.
                var paraGroupMatch = paraParentMatch.Success
                    ? null
                    : Regex.Match(parentPath, @"^/slide\[(\d+)\]((?:/group\[\d+\])+)/shape\[(\d+)\]$");
                var paraPhMatch = (paraParentMatch.Success || (paraGroupMatch?.Success == true)) ? null : Regex.Match(parentPath, @"^/slide\[(\d+)\]/placeholder\[(\w+)\]$");
                // R57 bt-4: accept connector parents so dump→replay round-trips
                // multi-paragraph / multi-run connector labels (the inline
                // `text=` prop on AddConnector handles only the single-run
                // case). The connector's <p:txBody> is not declared by the
                // p:cxnSp schema — see ConnectorEnsureTextBody.
                var paraCxnMatch = (paraParentMatch.Success || (paraGroupMatch?.Success == true) || (paraPhMatch?.Success == true))
                    ? null
                    : Regex.Match(parentPath, @"^/slide\[(\d+)\]/connector\[([^\]]+)\]$");
                if (!paraParentMatch.Success && (paraGroupMatch == null || !paraGroupMatch.Success) && (paraPhMatch == null || !paraPhMatch.Success) && (paraCxnMatch == null || !paraCxnMatch.Success))
                    throw new ArgumentException("Paragraphs must be added to a shape, placeholder, or connector: /slide[N]/shape[M], /slide[N]/group[K]/.../shape[M], /slide[N]/placeholder[X], or /slide[N]/connector[K]");

                SlidePart paraSlidePart;
                Shape? paraShape = null;
                ConnectionShape? paraCxn = null;
                int paraSlideIdx;
                int paraShapeIdx;
                string paraReturnPathHead;
                if (paraParentMatch.Success)
                {
                    paraSlideIdx = int.Parse(paraParentMatch.Groups[1].Value);
                    paraShapeIdx = int.Parse(paraParentMatch.Groups[2].Value);
                    (paraSlidePart, paraShape) = ResolveShape(paraSlideIdx, paraShapeIdx);
                    paraReturnPathHead = $"/slide[{paraSlideIdx}]/{BuildElementPathSegment("shape", paraShape, paraShapeIdx)}";
                }
                else if (paraGroupMatch != null && paraGroupMatch.Success)
                {
                    paraSlideIdx = int.Parse(paraGroupMatch.Groups[1].Value);
                    var groupSegs = paraGroupMatch.Groups[2].Value;
                    paraShapeIdx = int.Parse(paraGroupMatch.Groups[3].Value);
                    (paraSlidePart, paraShape) = ResolveGroupInnerShapeBySegments(paraSlideIdx, groupSegs, paraShapeIdx);
                    paraReturnPathHead = $"/slide[{paraSlideIdx}]{groupSegs}/{BuildElementPathSegment("shape", paraShape, paraShapeIdx)}";
                }
                else if (paraPhMatch != null && paraPhMatch.Success)
                {
                    paraSlideIdx = int.Parse(paraPhMatch.Groups[1].Value);
                    var phToken = paraPhMatch.Groups[2].Value;
                    var slideParts = GetSlideParts().ToList();
                    if (paraSlideIdx < 1 || paraSlideIdx > slideParts.Count)
                        throw new ArgumentException($"Slide {paraSlideIdx} not found (total: {slideParts.Count})");
                    paraSlidePart = slideParts[paraSlideIdx - 1];
                    paraShape = ResolvePlaceholderShape(paraSlidePart, phToken);
                    paraShapeIdx = 1;
                    paraReturnPathHead = $"/slide[{paraSlideIdx}]/{BuildElementPathSegment("shape", paraShape, paraShapeIdx)}";
                }
                else
                {
                    paraSlideIdx = int.Parse(paraCxnMatch!.Groups[1].Value);
                    var cxnTok = paraCxnMatch.Groups[2].Value;
                    var slideParts = GetSlideParts().ToList();
                    if (paraSlideIdx < 1 || paraSlideIdx > slideParts.Count)
                        throw new ArgumentException($"Slide {paraSlideIdx} not found (total: {slideParts.Count})");
                    paraSlidePart = slideParts[paraSlideIdx - 1];
                    paraCxn = ResolveConnectorByToken(paraSlidePart, cxnTok);
                    paraShapeIdx = ConnectorPositionalIndex(paraSlidePart, paraCxn);
                    paraReturnPathHead = $"/slide[{paraSlideIdx}]/{BuildElementPathSegment("connector", paraCxn, paraShapeIdx)}";
                }

                var textBody = paraShape != null
                    ? (paraShape.TextBody
                        ?? throw new InvalidOperationException("Shape has no text body"))
                    : ConnectorEnsureTextBody(paraCxn!);

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
                    pProps.Indent = (int)Math.Round(SpacingConverter.ParsePointsSigned(pIndent) * EmuConverter.EmuPerPointF);
                }
                if (properties.TryGetValue("marginLeft", out var pMarL) || properties.TryGetValue("marl", out pMarL))
                    pProps.LeftMargin = (int)Math.Round(SpacingConverter.ParsePointsSigned(pMarL) * EmuConverter.EmuPerPointF);
                if (properties.TryGetValue("marginRight", out var pMarR) || properties.TryGetValue("marr", out pMarR))
                    pProps.RightMargin = (int)Math.Round(SpacingConverter.ParsePointsSigned(pMarR) * EmuConverter.EmuPerPointF);
                // bulletRaw (full bullet group) takes precedence over the lossy
                // `list` keyword when both are present. Probe both
                // unconditionally: dump emits list AND bulletRaw side by side,
                // and an else-if short-circuit left `list` unread — flagged as
                // a false unsupported_property on every numbered-list replay.
                var hasPBulletRaw = properties.TryGetValue("bulletRaw", out var pBulletRaw) || properties.TryGetValue("bulletraw", out pBulletRaw);
                var hasPList = properties.TryGetValue("list", out var pList) || properties.TryGetValue("liststyle", out pList) || properties.TryGetValue("bullet", out pList);
                if (hasPBulletRaw)
                    ApplyBulletRaw(pProps, pBulletRaw!);
                else if (hasPList)
                    ApplyListStyle(pProps, pList!, preserveIndent: properties.ContainsKey("indent") || properties.ContainsKey("marginLeft") || properties.ContainsKey("marginleft") || properties.ContainsKey("marL") || properties.ContainsKey("marl"));
                // Paragraph-level default run properties (verbatim). Bare runs
                // inherit size/bold/font from here; see ApplyDefRPrRaw.
                if (properties.TryGetValue("defRPrRaw", out var pDefRPrRaw) || properties.TryGetValue("defrprraw", out pDefRPrRaw))
                    ApplyDefRPrRaw(pProps, pDefRPrRaw);
                if (properties.TryGetValue("level", out var pLevelStr))
                {
                    if (!int.TryParse(pLevelStr, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out var pLevelVal) || pLevelVal < 0 || pLevelVal > 8)
                        throw new ArgumentException($"Invalid 'level' value: '{pLevelStr}'. Expected an integer between 0 and 8 (OOXML a:pPr/@lvl).");
                    pProps.Level = pLevelVal;
                }
                // CJK / line-break pPr attributes (eaLnBrk / latinLnBrk /
                // hangingPunct / fontAlgn / defTabSz). Mirror NodeBuilder readback.
                foreach (var pBreakKey in new[] { "eaLnBrk", "latinLnBrk", "fontAlgn", "defTabSz" })
                    if (properties.TryGetValue(pBreakKey, out var pBreakVal))
                        ApplyParagraphBreakProp(pProps, pBreakKey, pBreakVal);
                // Line spacing (CONSISTENCY(lineSpacing): same idiom as AddShape:~180)
                if (properties.TryGetValue("lineSpacing", out var pLsVal) || properties.TryGetValue("linespacing", out pLsVal))
                {
                    var (pLsInternal, pLsIsPercent) = SpacingConverter.ParsePptLineSpacing(pLsVal);
                    pProps.RemoveAllChildren<Drawing.LineSpacing>();
                    if (pLsIsPercent)
                        InsertPPrChild(pProps, new Drawing.LineSpacing(
                            new Drawing.SpacingPercent { Val = pLsInternal }));
                    else
                        InsertPPrChild(pProps, new Drawing.LineSpacing(
                            new Drawing.SpacingPoints { Val = pLsInternal }));
                }
                if (properties.TryGetValue("spaceBefore", out var pSbVal) || properties.TryGetValue("spacebefore", out pSbVal))
                {
                    pProps.RemoveAllChildren<Drawing.SpaceBefore>();
                    InsertPPrChild(pProps, new Drawing.SpaceBefore(new Drawing.SpacingPoints { Val = SpacingConverter.ParsePptSpacing(pSbVal) }));
                }
                if (properties.TryGetValue("spaceAfter", out var pSaVal) || properties.TryGetValue("spaceafter", out pSaVal))
                {
                    pProps.RemoveAllChildren<Drawing.SpaceAfter>();
                    InsertPPrChild(pProps, new Drawing.SpaceAfter(new Drawing.SpacingPoints { Val = SpacingConverter.ParsePptSpacing(pSaVal) }));
                }
                // R65 bt-2: <a:tabLst>/<a:tab pos algn/> — accept the compact
                // compound form emitted by NodeBuilder so dump→replay restores
                // custom tab stops. Schema-order is handled by AppendChild here
                // because AddParagraph builds pPr top-to-bottom in declaration
                // order (tabLst rank > spcBef/spcAft/list); SetParagraph routes
                // through InsertPPrChild for the reverse-order case.
                if (properties.TryGetValue("tabs", out var pTabsVal) || properties.TryGetValue("tablist", out pTabsVal))
                {
                    var pTabList = ParseTabStopList(pTabsVal);
                    if (pTabList != null)
                    {
                        pProps.RemoveAllChildren<Drawing.TabStopList>();
                        pProps.AppendChild(pTabList);
                    }
                }

                // CONSISTENCY(pptx-no-empty-ppr): only attach paragraph
                // properties when at least one was set. Empty <a:pPr/>
                // is a real OOXML node — it doesn't affect rendering but
                // bloats every paragraph after the first on dump→replay
                // (the seeded first paragraph already has no pPr by
                // default, so the bloat is one xml element per added
                // paragraph). Skip when pProps has no attribute and no
                // child element.
                if (pProps.HasAttributes || pProps.HasChildren)
                    newPara.ParagraphProperties = pProps;

                // Create initial run with text and run-level properties
                var paraText = properties.GetValueOrDefault("text", "");
                XmlTextValidator.ValidateOrThrow(paraText, "text", allowSoftBreakChar: true);
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
                    // R56 bt-3: accept the canonical `33%` form emitted by Get
                    // alongside the legacy bare `33` (both = 33% superscript).
                    var pBlNorm = pBaseline.Trim().TrimEnd('%').Trim();
                    rProps.Baseline = pBlNorm.ToLowerInvariant() switch
                    {
                        "super" or "true" => 30000,
                        "sub" => -25000,
                        _ => double.TryParse(pBlNorm, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var pBlVal) && !double.IsNaN(pBlVal) && !double.IsInfinity(pBlVal)
                            ? (int)(pBlVal * 1000)
                            : throw new ArgumentException($"Invalid 'baseline' value: '{pBaseline}'. Expected 'super', 'sub', or a percentage (e.g. 30 or 30%).")
                    };
                }

                // CONSISTENCY(escape-sequences): \n still routes as raw newline
                // inside a single <a:t> (paragraph-level only adds one paragraph
                // here), but \t expands to <a:tab/> siblings between text runs
                // so tabular text round-trips through PowerPoint.
                if (paraText.Contains('\t'))
                {
                    AppendLineWithTabs(newPara, paraText, seg => new Drawing.Run
                    {
                        RunProperties = (Drawing.RunProperties)rProps.CloneNode(true),
                        Text = MakePreservingText(seg)
                    });
                }
                else if (paraText.Length == 0)
                {
                    // Empty paragraph (spacer line): real PowerPoint computes
                    // an empty line's height from <a:endParaRPr>, NOT from an
                    // empty run's rPr — a `<a:r><a:rPr sz="800"/><a:t/></a:r>`
                    // still renders at the inherited body size, inflating
                    // designer spacer paragraphs (8pt gap → 22pt gap). Write
                    // the collected run properties as endParaRPr instead.
                    var endPr = new Drawing.EndParagraphRunProperties();
                    foreach (var attr in rProps.GetAttributes())
                        endPr.SetAttribute(attr);
                    foreach (var child in rProps.ChildElements)
                        endPr.AppendChild(child.CloneNode(true));
                    newPara.Append(endPr);
                }
                else
                {
                    newRun.RunProperties = rProps;
                    newRun.Text = MakePreservingText(paraText);
                    newPara.Append(newRun);
                }

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
                return $"{paraReturnPathHead}/paragraph[{paraCount}]";
    }


    /// <summary>
    /// `add --type linebreak /slide[N]/shape[M]/paragraph[K]` (also /placeholder[X]) —
    /// insert an &lt;a:br/&gt; element into the target paragraph. Mirrors AddRun's path
    /// resolution shape so /paragraph[K] suffix and /placeholder[X] alias both work.
    /// </summary>
    private string AddLineBreak(string parentPath, int? index, Dictionary<string, string> properties)
    {
        // CONSISTENCY(pptx-group-flatten): accept an optional /group[G]/.../group[M]
        // chain between the slide and the shape so a line break can be added to a
        // textbox sitting inside a group — exactly mirroring AddRun / AddParagraph,
        // which already walk the same chain. Without it, dump→replay of an <a:br>
        // inside a grouped shape reported "Line breaks must be added to a
        // shape/placeholder or paragraph" even though Get exposed the path verbatim.
        var brParaMatch = Regex.Match(parentPath, @"^/slide\[(\d+)\]((?:/group\[\d+\])*)/shape\[(\d+)\](?:/(?:paragraph|p)\[(\d+)\])?$");
        var brPhMatch = brParaMatch.Success ? null : Regex.Match(parentPath, @"^/slide\[(\d+)\]/placeholder\[(\w+)\](?:/(?:paragraph|p)\[(\d+)\])?$");
        if (!brParaMatch.Success && (brPhMatch == null || !brPhMatch.Success))
            throw new ArgumentException(
                "Line breaks must be added to a shape/placeholder or paragraph: " +
                "/slide[N]/shape[M], /slide[N]/placeholder[X], /slide[N]/shape[M]/paragraph[K], or /slide[N]/placeholder[X]/paragraph[K]");

        Shape brShape;
        System.Text.RegularExpressions.Group brParaGroup;
        string brReturnPathHead;
        if (brParaMatch.Success)
        {
            var slideIdx = int.Parse(brParaMatch.Groups[1].Value);
            var grpChain = brParaMatch.Groups[2].Value;
            var shapeIdx = int.Parse(brParaMatch.Groups[3].Value);
            if (string.IsNullOrEmpty(grpChain))
            {
                (_, brShape) = ResolveShape(slideIdx, shapeIdx);
            }
            else
            {
                // Walk the /group[N]/.../group[M]/shape[K] chain, filtering each
                // scope to content elements — same resolution AddRun uses.
                var sps = GetSlideParts().ToList();
                if (slideIdx < 1 || slideIdx > sps.Count)
                    throw new ArgumentException($"Slide {slideIdx} not found (total: {sps.Count})");
                OpenXmlCompositeElement scope = GetSlide(sps[PathIndex.ToArrayIndex(slideIdx)]).CommonSlideData?.ShapeTree
                    ?? throw new ArgumentException($"Slide {slideIdx} has no shapes");
                foreach (Match gm in Regex.Matches(grpChain, @"/group\[(\d+)\]"))
                {
                    var gIdx = int.Parse(gm.Groups[1].Value);
                    var groupsHere = scope.Elements<GroupShape>().ToList();
                    if (gIdx < 1 || gIdx > groupsHere.Count)
                        throw new ArgumentException($"Group {gIdx} not found in scope (have {groupsHere.Count})");
                    scope = groupsHere[gIdx - 1];
                }
                var shapesInScope = scope.Elements<Shape>().ToList();
                if (shapeIdx < 1 || shapeIdx > shapesInScope.Count)
                    throw new ArgumentException($"Shape {shapeIdx} not found in group scope (have {shapesInScope.Count})");
                brShape = shapesInScope[PathIndex.ToArrayIndex(shapeIdx)];
            }
            brParaGroup = brParaMatch.Groups[4];
            brReturnPathHead = $"/slide[{slideIdx}]{grpChain}/shape[{shapeIdx}]";
        }
        else
        {
            var slideIdx = int.Parse(brPhMatch!.Groups[1].Value);
            var phToken = brPhMatch.Groups[2].Value;
            var slideParts = GetSlideParts().ToList();
            if (slideIdx < 1 || slideIdx > slideParts.Count)
                throw new ArgumentException($"Slide {slideIdx} not found (total: {slideParts.Count})");
            brShape = ResolvePlaceholderShape(slideParts[PathIndex.ToArrayIndex(slideIdx)], phToken);
            brParaGroup = brPhMatch.Groups[3];
            brReturnPathHead = $"/slide[{slideIdx}]/placeholder[{phToken}]";
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
        // Verbatim <a:rPr> on the break — controls the empty line's height
        // (a bare <a:br/> inherits the paragraph size instead). Emitted by
        // the dump as rPrRaw when the source break carries one.
        if (properties != null
            && (properties.TryGetValue("rPrRaw", out var brRPrRaw) || properties.TryGetValue("rprraw", out brRPrRaw))
            && !string.IsNullOrWhiteSpace(brRPrRaw))
        {
            br.AppendChild(new Drawing.RunProperties(brRPrRaw));
        }
        if (index.HasValue)
        {
            var children = targetPara.ChildElements.ToList();
            var insertAt = Math.Max(0, Math.Min(index.Value, children.Count));
            if (insertAt >= children.Count) targetPara.AppendChild(br);
            else children[insertAt].InsertBeforeSelf(br);
        }
        else
        {
            // OOXML schema: <a:br> must precede <a:endParaRPr> inside <a:p>.
            // AddParagraph seeds an empty paragraph that already carries an
            // <a:endParaRPr> (and the seeded shape factory adds it whenever
            // a body-text placeholder is materialised). Appending the break
            // unconditionally drops it AFTER endParaRPr and PowerPoint refuses
            // the file (0x80070570) because the schema forbids that order.
            // Insert before the first endParaRPr when present so the break
            // lands in the only legal slot — between runs/breaks and the
            // closing endParaRPr. Falls back to append when the paragraph
            // has no endParaRPr seed.
            var endParaRPr = targetPara.GetFirstChild<Drawing.EndParagraphRunProperties>();
            if (endParaRPr != null)
                targetPara.InsertBefore(br, endParaRPr);
            else
                targetPara.AppendChild(br);
        }

        var brIdx = targetPara.Elements<Drawing.Break>().ToList().FindIndex(b => ReferenceEquals(b, br)) + 1;
        return $"{brReturnPathHead}/paragraph[{targetParaIdx}]/br[{brIdx}]";
    }

    private string AddRun(string parentPath, int? index, Dictionary<string, string> properties)
    {
                // Add a run to a paragraph: /slide[N]/shape[M]/paragraph[P] or /slide[N]/shape[M]
                //   also: /slide[N]/placeholder[X]/paragraph[P] or /slide[N]/placeholder[X]
                //   also (group-nested): /slide[N]/group[G]/.../shape[M][/paragraph[P]] with
                //   one or more intervening group[] segments (PowerPoint allows arbitrarily
                //   nested group trees; AddParagraph / SetParagraph already accept the same
                //   shape — without this, dump→replay of a textbox sitting inside a group
                //   reported "Runs must be added to a shape/placeholder or paragraph" for
                //   every per-run set op, even though Get exposed the path verbatim).
                // CONSISTENCY(path-aliases): accept short-form `/p[N]` alongside `/paragraph[N]`.
                // CONSISTENCY(placeholder-paragraph-path): mirror the dual route that
                // AddParagraph and SetParagraph already accept.
                var runParaMatch = Regex.Match(parentPath, @"^/slide\[(\d+)\]((?:/group\[\d+\])*)/shape\[(\d+)\](?:/(?:paragraph|p)\[(\d+)\])?$");
                var runPhMatch = runParaMatch.Success ? null : Regex.Match(parentPath, @"^/slide\[(\d+)\]((?:/group\[\d+\])*)/placeholder\[(\w+)\](?:/(?:paragraph|p)\[(\d+)\])?$");
                // R57 bt-4: accept connector parents (slide-root only; group-
                // nested connectors fall back to positional resolution under
                // the group walker, deferred to a follow-up if encountered).
                var runCxnMatch = (runParaMatch.Success || (runPhMatch?.Success == true))
                    ? null
                    : Regex.Match(parentPath, @"^/slide\[(\d+)\]/connector\[([^\]]+)\](?:/(?:paragraph|p)\[(\d+)\])?$");
                if (!runParaMatch.Success && (runPhMatch == null || !runPhMatch.Success) && (runCxnMatch == null || !runCxnMatch.Success))
                    throw new ArgumentException("Runs must be added to a shape/placeholder/connector or paragraph: /slide[N]/shape[M], /slide[N]/placeholder[X], /slide[N]/connector[K], or one of those with /paragraph[P] suffix");

                SlidePart runSlidePart;
                Shape? runShape = null;
                ConnectionShape? runCxn = null;
                int runSlideIdx;
                int runShapeIdx;
                string runReturnPathHead;
                System.Text.RegularExpressions.Group paraGroup;
                if (runParaMatch.Success)
                {
                    runSlideIdx = int.Parse(runParaMatch.Groups[1].Value);
                    var grpChain = runParaMatch.Groups[2].Value;
                    runShapeIdx = int.Parse(runParaMatch.Groups[3].Value);
                    if (string.IsNullOrEmpty(grpChain))
                    {
                        (runSlidePart, runShape) = ResolveShape(runSlideIdx, runShapeIdx);
                        runReturnPathHead = $"/slide[{runSlideIdx}]/{BuildElementPathSegment("shape", runShape, runShapeIdx)}";
                    }
                    else
                    {
                        // CONSISTENCY(pptx-group-flatten): walk down the
                        // /group[N]/.../group[M]/shape[K] chain so AddRun can
                        // target a textbox sitting inside a group. Path
                        // semantics mirror InsertAtPosition (Helpers.Path.cs)
                        // — group children are filtered to content elements
                        // (skip NonVisualGroupShapeProperties / GroupShapeProperties).
                        var sps = GetSlideParts().ToList();
                        if (runSlideIdx < 1 || runSlideIdx > sps.Count)
                            throw new ArgumentException($"Slide {runSlideIdx} not found (total: {sps.Count})");
                        runSlidePart = sps[runSlideIdx - 1];
                        OpenXmlCompositeElement scope = GetSlide(runSlidePart).CommonSlideData?.ShapeTree
                            ?? throw new ArgumentException($"Slide {runSlideIdx} has no shapes");
                        foreach (Match gm in Regex.Matches(grpChain, @"/group\[(\d+)\]"))
                        {
                            var gIdx = int.Parse(gm.Groups[1].Value);
                            var groupsHere = scope.Elements<GroupShape>().ToList();
                            if (gIdx < 1 || gIdx > groupsHere.Count)
                                throw new ArgumentException($"Group {gIdx} not found in scope (have {groupsHere.Count})");
                            scope = groupsHere[gIdx - 1];
                        }
                        var shapesInScope = scope.Elements<Shape>().ToList();
                        if (runShapeIdx < 1 || runShapeIdx > shapesInScope.Count)
                            throw new ArgumentException($"Shape {runShapeIdx} not found in group scope (have {shapesInScope.Count})");
                        runShape = shapesInScope[runShapeIdx - 1];
                        runReturnPathHead = $"/slide[{runSlideIdx}]{grpChain}/{BuildElementPathSegment("shape", runShape, runShapeIdx)}";
                    }
                    paraGroup = runParaMatch.Groups[4];
                }
                else if (runPhMatch != null && runPhMatch.Success)
                {
                    runSlideIdx = int.Parse(runPhMatch.Groups[1].Value);
                    // Placeholder paths nested inside a group are rare in
                    // PowerPoint (placeholders typically live at slide-level),
                    // but accept the syntax for symmetry with the shape branch
                    // so the regex shape (slide / group-chain / placeholder /
                    // paragraph) is uniform. Placeholder resolution stays at
                    // the slide level — phType / phIndex matching scans the
                    // entire slide; group nesting is a no-op for the lookup.
                    var phToken = runPhMatch.Groups[3].Value;
                    var slideParts = GetSlideParts().ToList();
                    if (runSlideIdx < 1 || runSlideIdx > slideParts.Count)
                        throw new ArgumentException($"Slide {runSlideIdx} not found (total: {slideParts.Count})");
                    runSlidePart = slideParts[runSlideIdx - 1];
                    runShape = ResolvePlaceholderShape(runSlidePart, phToken);
                    runShapeIdx = 1;
                    paraGroup = runPhMatch.Groups[4];
                    runReturnPathHead = $"/slide[{runSlideIdx}]/{BuildElementPathSegment("shape", runShape, runShapeIdx)}";
                }
                else
                {
                    runSlideIdx = int.Parse(runCxnMatch!.Groups[1].Value);
                    var cxnTok = runCxnMatch.Groups[2].Value;
                    var slideParts = GetSlideParts().ToList();
                    if (runSlideIdx < 1 || runSlideIdx > slideParts.Count)
                        throw new ArgumentException($"Slide {runSlideIdx} not found (total: {slideParts.Count})");
                    runSlidePart = slideParts[runSlideIdx - 1];
                    runCxn = ResolveConnectorByToken(runSlidePart, cxnTok);
                    runShapeIdx = ConnectorPositionalIndex(runSlidePart, runCxn);
                    paraGroup = runCxnMatch.Groups[3];
                    runReturnPathHead = $"/slide[{runSlideIdx}]/{BuildElementPathSegment("connector", runCxn, runShapeIdx)}";
                }

                var runTextBody = runShape != null
                    ? (runShape.TextBody
                        ?? throw new InvalidOperationException("Shape has no text body"))
                    : ConnectorEnsureTextBody(runCxn!);

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
                    if (paras.Count == 0)
                    {
                        // R57 bt-4: connector txBody starts empty (no seeded
                        // <a:p>) so `add run` against a fresh connector path
                        // would throw before even creating the run. Seed one
                        // paragraph on demand — mirrors AddShape's auto-empty
                        // seed pattern at the txBody level.
                        var seeded = new Drawing.Paragraph();
                        runTextBody.Append(seeded);
                        targetPara = seeded;
                        targetParaIdx = 1;
                    }
                    else
                    {
                        targetPara = paras[^1];
                        targetParaIdx = paras.Count;
                    }
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
                if (properties.TryGetValue("bold", out var rBold)
                    || properties.TryGetValue("font.bold", out rBold))
                    rProps.Bold = IsTruthy(rBold);
                if (properties.TryGetValue("italic", out var rItalic)
                    || properties.TryGetValue("font.italic", out rItalic))
                    rProps.Italic = IsTruthy(rItalic);
                if (properties.TryGetValue("underline", out var rUnderline)
                    || properties.TryGetValue("font.underline", out rUnderline))
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
                // R61 bt-1: AddRun honors textOutline (and its width/color
                // split keys) so a single-run-collapse dump emits `add run
                // textOutline=…` and re-add round-trips the <a:ln> on rPr.
                // Symmetric with the Set branch in ShapeProperties.cs. Schema
                // order is enforced by ReorderDrawingRunProperties at the end
                // of this method (already invoked for endParaRPr inheritance).
                // Verbatim <a:blipFill> glyph fill (WordArt picture fill).
                if ((properties.TryGetValue("textFillRaw", out var rTfRaw)
                        || properties.TryGetValue("textfillraw", out rTfRaw))
                    && !string.IsNullOrWhiteSpace(rTfRaw))
                {
                    rProps.RemoveAllChildren<Drawing.SolidFill>();
                    rProps.RemoveAllChildren<Drawing.GradientFill>();
                    rProps.RemoveAllChildren<Drawing.BlipFill>();
                    rProps.RemoveAllChildren<Drawing.PatternFill>();
                    InsertFillInRunProperties(rProps, new Drawing.BlipFill(rTfRaw));
                }

                // Verbatim <a:ln> (dash pattern / gradient stroke / cap-join —
                // everything the width:color compound can't express). Wins over
                // the semantic keys; the emitter suppresses them when present.
                if ((properties.TryGetValue("textOutlineRaw", out var rToRaw)
                        || properties.TryGetValue("textoutlineraw", out rToRaw))
                    && !string.IsNullOrWhiteSpace(rToRaw))
                {
                    rProps.RemoveAllChildren<Drawing.Outline>();
                    rProps.PrependChild(new Drawing.Outline(rToRaw));
                }
                else if (properties.TryGetValue("textOutline", out var rTextOutline)
                    || properties.TryGetValue("textoutline", out rTextOutline))
                {
                    if (!rTextOutline.Equals("none", StringComparison.OrdinalIgnoreCase)
                        && !rTextOutline.Equals("false", StringComparison.OrdinalIgnoreCase))
                    {
                        // Compound is width:color (Get emit form). See Set
                        // branch in ShapeProperties.cs for the name-shadowing
                        // rationale on SplitCompoundLineValue's positional tuple.
                        var (toWidthPart, toColorPart, _) = SplitCompoundLineValue(rTextOutline);
                        long? widthEmu = null;
                        // Carry the full color string (incl. +shade/+alpha/+lumMod
                        // transform chain and #RRGGBBAA alpha) through BuildSolidFill
                        // so Get's emit form ("#4F81BD11+shade2") replays. Mirrors
                        // the Set branch in ShapeProperties.cs.
                        string? colorValue = null;
                        if (toColorPart != null)
                        {
                            widthEmu = Core.EmuConverter.ParseLineWidth(toWidthPart);
                            colorValue = toColorPart.Equals("none", StringComparison.OrdinalIgnoreCase)
                                ? null : toColorPart;
                        }
                        else
                        {
                            try { widthEmu = Core.EmuConverter.ParseLineWidth(rTextOutline); }
                            catch { widthEmu = null; }
                            if (widthEmu == null && !rTextOutline.Equals("true", StringComparison.OrdinalIgnoreCase))
                                colorValue = rTextOutline;
                        }
                        var ln = new Drawing.Outline();
                        if (widthEmu.HasValue) ln.Width = (int)widthEmu.Value;
                        if (colorValue != null)
                            ln.AppendChild(BuildSolidFill(colorValue));
                        rProps.PrependChild(ln);
                    }
                }
                if (properties.TryGetValue("textOutline.width", out var rToWidth)
                    || properties.TryGetValue("textoutline.width", out rToWidth))
                {
                    var widthEmu = Core.EmuConverter.ParseLineWidth(rToWidth);
                    var ln = rProps.GetFirstChild<Drawing.Outline>();
                    if (ln == null)
                    {
                        ln = new Drawing.Outline();
                        rProps.PrependChild(ln);
                    }
                    ln.Width = (int)widthEmu;
                }
                if (properties.TryGetValue("textOutline.color", out var rToColor)
                    || properties.TryGetValue("textoutline.color", out rToColor))
                {
                    var ln = rProps.GetFirstChild<Drawing.Outline>();
                    if (ln == null)
                    {
                        ln = new Drawing.Outline();
                        rProps.PrependChild(ln);
                    }
                    ln.RemoveAllChildren<Drawing.SolidFill>();
                    // BuildSolidFill carries the transform chain / alpha Get emits.
                    ln.AppendChild(BuildSolidFill(rToColor));
                }
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
                if (properties.TryGetValue("color", out var rColor)
                    || properties.TryGetValue("font.color", out rColor))
                    rProps.AppendChild(BuildSolidFill(rColor));
                // CONSISTENCY(highlight): a:highlight slot sits between the fill
                // and latin/ea in CT_TextCharacterProperties — appending here
                // (after solidFill, before the font branches below) lands it in
                // schema position. Same write as the Set case in
                // ShapeProperties.cs / ApplyPptRunFormatting.
                if (properties.TryGetValue("highlight", out var rHighlight)
                    && !rHighlight.Equals("none", StringComparison.OrdinalIgnoreCase)
                    && !rHighlight.Equals("false", StringComparison.OrdinalIgnoreCase))
                {
                    var rHl = new Drawing.Highlight();
                    rHl.AppendChild(BuildColorElement(rHighlight));
                    rProps.AppendChild(rHl);
                }
                if (properties.TryGetValue("font", out var rFont)
                    || properties.TryGetValue("font.name", out rFont))
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
                    // R56 bt-3: accept canonical `33%` (Get emit form) and bare `33`.
                    var rBlNorm = rBaseline.Trim().TrimEnd('%').Trim();
                    rProps.Baseline = rBlNorm.ToLowerInvariant() switch
                    {
                        "super" or "true" => 30000,
                        "sub" => -25000,
                        "none" or "false" or "0" => 0,
                        _ => (int)(ParseHelpers.SafeParseDouble(rBlNorm, "baseline") * 1000)
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
                    // ST_UnsignedInt-ish — accept any integer string; let
                    // ParseHelpers normalize (matches the Set int-attr path).
                    var smtIdVal = OfficeCli.Core.ParseHelpers.SafeParseInt(smtIdRaw, "smtId");
                    if (smtIdVal < 0)
                        throw new ArgumentException($"Invalid smtId '{smtIdRaw}' (must be non-negative).");
                    rProps.SetAttribute(new DocumentFormat.OpenXml.OpenXmlAttribute(
                        "", "smtId", "", smtIdVal.ToString(System.Globalization.CultureInfo.InvariantCulture)));
                }

                // R59 tester-1: inherit paragraph-level <a:endParaRPr> children
                // onto the new run's <a:rPr>. When `set …/paragraph[N] font.*`
                // lands before any run exists, font.* falls back to endParaRPr
                // (the only rPr-shaped sink on an empty paragraph). A subsequent
                // `add run` would then create a bare <a:rPr lang="…"/> and the
                // endParaRPr typeface/color/bold/etc. would silently disappear
                // from the rendered run. POI's XSLFTextRun resolves missing
                // CTRunProperties slots up the inheritance chain (run → para's
                // endParaRPr → lvl*pPr/defRPr → master) at read time; we apply
                // the lowest step at write time so dump→batch→re-dump is
                // byte-stable. Only copy children whose type isn't already
                // present on the new rPr (explicit run props win).
                var srcEndPr = targetPara.GetFirstChild<Drawing.EndParagraphRunProperties>();
                if (srcEndPr != null)
                {
                    foreach (var srcChild in srcEndPr.ChildElements)
                    {
                        var t = srcChild.GetType();
                        bool already = false;
                        foreach (var existing in rProps.ChildElements)
                            if (existing.GetType() == t) { already = true; break; }
                        if (!already)
                            rProps.AppendChild((OpenXmlElement)srcChild.CloneNode(true));
                    }
                    // Mirror endParaRPr scalar attributes (b, i, u, sz, kern,
                    // baseline, strike, cap, spc, dirty, …) onto rPr when the
                    // run didn't set them. Skip `lang`/`altLang` — AddRun
                    // already initialized rProps.Language="en-US" by default
                    // and respects explicit `lang=` overrides, matching the
                    // historical run-add contract.
                    foreach (var srcAttr in srcEndPr.GetAttributes())
                    {
                        if (srcAttr.LocalName == "lang" || srcAttr.LocalName == "altLang")
                            continue;
                        bool hasAttr = false;
                        foreach (var existing in rProps.GetAttributes())
                            if (existing.LocalName == srcAttr.LocalName
                                && existing.NamespaceUri == srcAttr.NamespaceUri)
                            { hasAttr = true; break; }
                        if (!hasAttr)
                            rProps.SetAttribute(new DocumentFormat.OpenXml.OpenXmlAttribute(
                                srcAttr.Prefix, srcAttr.LocalName, srcAttr.NamespaceUri, srcAttr.Value));
                    }
                    ReorderDrawingRunProperties(rProps);
                }

                // R62 bt-5: fillOverlayRaw — verbatim <a:fillOverlay…/> install
                // on the new run's <a:rPr><a:effectLst>. AddShape funnels this
                // through SetRunOrShapeProperties.effectKeys (so the shape-spPr
                // and run-rPr branches both fire); AddRun has its own narrower
                // prop loop and would otherwise drop the key as UNSUPPORTED on
                // dump→replay of a run carrying a fillOverlay. Mirror the
                // shape-Add routing here so `add run fillOverlayRaw=…` lands
                // on the new run instead of the silent-drop path.
                if (properties.TryGetValue("fillOverlayRaw", out var rFillOv)
                    || properties.TryGetValue("filloverlayraw", out rFillOv))
                {
                    // Run is not yet attached; bind rProps first so ApplyRunFillOverlayRaw
                    // can call run.RunProperties ?? new RunProperties() safely.
                    newRun.RunProperties = rProps;
                    ApplyRunFillOverlayRaw(newRun, rFillOv);
                    // RunProperties may have been mutated; reassign for the
                    // subsequent newRun.RunProperties = rProps no-op below.
                    rProps = newRun.RunProperties!;
                }

                newRun.RunProperties = rProps;
                // Hyperlink on the new run. Schema declares link.add=true with
                // parent "shape|run" — without this branch the shape-level Add
                // path accepts link= but `add ... --type run --prop link=...`
                // reports UNSUPPORTED, forcing callers into a second Set call.
                // Tooltip is paired with link (matches the AddShape / AddPicture
                // / AddGroup pattern).
                if (properties.TryGetValue("link", out var rLink))
                    ApplyRunHyperlink(runSlidePart, newRun, rLink, properties.GetValueOrDefault("tooltip"));
                // CONSISTENCY(escape-sequences): match shape-text path (\n and \t
                // two-char escapes resolved). Run-add stays single-element, so
                // tabs land as raw chars inside <a:t> rather than <a:tab/>;
                // higher-level shape-text Add/Set splits on \t into separate
                // runs with <a:tab/> siblings.
                newRun.Text = MakePreservingText(runText);

                // Insert run at specified index, or append
                if (index.HasValue)
                {
                    var existingRuns = targetPara.Elements<Drawing.Run>().ToList();
                    if (index.Value >= 0 && index.Value < existingRuns.Count)
                        existingRuns[index.Value].InsertBeforeSelf(newRun);
                    else
                    {
                        var endParaRun2 = targetPara.GetFirstChild<Drawing.EndParagraphRunProperties>();
                        if (endParaRun2 != null)
                            targetPara.InsertBefore(newRun, endParaRun2);
                        else
                            targetPara.Append(newRun);
                    }
                }
                else
                {
                    var endParaRun = targetPara.GetFirstChild<Drawing.EndParagraphRunProperties>();
                    if (endParaRun != null)
                        targetPara.InsertBefore(newRun, endParaRun);
                    else
                        targetPara.Append(newRun);
                }

                var runCount = targetPara.Elements<Drawing.Run>().Count();
                GetSlide(runSlidePart).Save();
                return $"{runReturnPathHead}/paragraph[{targetParaIdx}]/run[{runCount}]";
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
        // NEWLINE-SEMANTICS-V2: '\v' is the cross-handler soft-line-break
        // char. In DrawingML that is <a:br/> — a line break INSIDE the
        // paragraph, as opposed to '\n' which the caller has already split
        // into separate <a:p> paragraphs. Emit each '\v'-separated chunk,
        // joined by <a:br/> elements ('\v' is XML-illegal and must never
        // reach <a:t> as a literal char).
        var vChunks = line.Split('\v');
        for (int vi = 0; vi < vChunks.Length; vi++)
        {
            if (vi > 0) paragraph.AppendChild(new Drawing.Break());
            AppendChunkWithTabs(paragraph, vChunks[vi], runFactory);
        }
    }

    private static void AppendChunkWithTabs(
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
