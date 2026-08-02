// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler
{
    private string AddShape(string parentPath, int? index, Dictionary<string, string> properties,
                            string? elementTypeHint = null)
    {
                // CONSISTENCY(master-layout-shape-edit): a shape parent may be a
                // slide (/slide[N]) or a master/layout part. Master/layout shapes
                // power branding workflows (logo on every slide, watermarks).
                // Resolve once here; downstream code branches on slidePart != null
                // for the few slide-only features (morph rename, animation,
                // hyperlink relationship). Path forms accepted:
                //   /slide[N]
                //   /slidemaster[N]
                //   /slidelayout[N]
                //   /slidemaster[N]/slidelayout[L]
                SlidePart? slidePart = null;
                List<SlidePart>? slideParts = null;
                int slideIdx = 0;
                ShapeTree shapeTree;
                OpenXmlPart ownerPart;
                OpenXmlPartRootElement ownerRoot;
                string returnPathPrefix;
                // CONSISTENCY(group-inner-shape-add): when the parent is a group,
                // newShape is appended to the GroupShape rather than the slide's
                // ShapeTree. shapeTree still points at the slide root for helpers
                // that need slide-wide context (shape-id allocation, query for
                // morph naming); insertContainer is what InsertAtPosition writes to.
                OpenXmlCompositeElement? insertContainer = null;
                string? groupResultPathPrefix = null;

                var masterOrLayout = TryResolveMasterOrLayoutShapeParent(parentPath);
                if (masterOrLayout is not null)
                {
                    var ml = masterOrLayout.Value;
                    shapeTree = ml.shapeTree;
                    ownerPart = ml.part;
                    ownerRoot = ml.root;
                    returnPathPrefix = ml.canonicalPrefix;
                }
                else
                {
                    // /slide[N]/group[K]/group[L]/… — add the shape inside the
                    // (possibly nested) group, not at slide root. Required by
                    // dump-replay: empty groups are emitted first, then
                    // per-child `add shape parent=/slide/group[K]/…`. Groups
                    // can nest arbitrarily deep; walk every group[K] segment
                    // to resolve the final insertion container.
                    // CONSISTENCY(group-nest): mirrors EmitGroup recursion.
                    var groupParentMatch = Regex.Match(parentPath, @"^/slide\[(\d+)\]((?:/group\[\d+\])+)$");
                    if (groupParentMatch.Success)
                    {
                        slideIdx = int.Parse(groupParentMatch.Groups[1].Value);
                        slideParts = GetSlideParts().ToList();
                        if (slideIdx < 1 || slideIdx > slideParts.Count)
                            throw new ArgumentException($"Slide {slideIdx} not found (total: {slideParts.Count})");
                        slidePart = slideParts[PathIndex.ToArrayIndex(slideIdx)];
                        var slideG = GetSlide(slidePart);
                        shapeTree = slideG.CommonSlideData?.ShapeTree
                            ?? throw new InvalidOperationException("Slide has no shape tree");
                        OpenXmlCompositeElement currentContainer = shapeTree;
                        var grpSegments = Regex.Matches(groupParentMatch.Groups[2].Value, @"/group\[(\d+)\]");
                        foreach (Match seg in grpSegments)
                        {
                            var grpIdx = int.Parse(seg.Groups[1].Value);
                            var groups = currentContainer.Elements<GroupShape>().ToList();
                            if (grpIdx < 1 || grpIdx > groups.Count)
                                throw new ArgumentException(
                                    $"Group {grpIdx} not found under {currentContainer.LocalName} on slide {slideIdx} (total: {groups.Count})");
                            currentContainer = groups[grpIdx - 1];
                        }
                        insertContainer = currentContainer;
                        ownerPart = slidePart;
                        ownerRoot = slideG;
                        returnPathPrefix = $"/slide[{slideIdx}]";
                        groupResultPathPrefix = $"/slide[{slideIdx}]{groupParentMatch.Groups[2].Value}";
                    }
                    else
                    {
                        var slideMatch = Regex.Match(parentPath, @"^/slide\[(\d+)\]$");
                        if (!slideMatch.Success)
                        {
                            var hint = parentPath.StartsWith("-")
                                ? " The parent is a positional argument, not a flag — run 'add <file> /slide[N] --type shape ...'."
                                : "";
                            throw new ArgumentException(
                                $"Shapes must be added to a slide, master, layout, or group, but got parent '{parentPath}'. Expected /slide[N], /slide[N]/group[K], /slidemaster[N], /slidelayout[N], or /slidemaster[N]/slidelayout[L].{hint}");
                        }

                        slideIdx = int.Parse(slideMatch.Groups[1].Value);
                        slideParts = GetSlideParts().ToList();
                        if (slideIdx < 1 || slideIdx > slideParts.Count)
                            throw new ArgumentException($"Slide {slideIdx} not found (total: {slideParts.Count})");

                        slidePart = slideParts[PathIndex.ToArrayIndex(slideIdx)];
                        var slide = GetSlide(slidePart);
                        shapeTree = slide.CommonSlideData?.ShapeTree
                            ?? throw new InvalidOperationException("Slide has no shape tree");
                        ownerPart = slidePart;
                        ownerRoot = slide;
                        returnPathPrefix = $"/slide[{slideIdx}]";
                    }
                }

                var text = properties.GetValueOrDefault("text", "");
                XmlTextValidator.ValidateOrThrow(text, "text", allowSoftBreakChar: true);
                var shapeId = AcquireShapeId(shapeTree, properties);
                var shapeName = properties.GetValueOrDefault("name", $"TextBox {shapeTree.Elements<Shape>().Count() + 1}");

                // Auto-add !! prefix if the slide (or the next slide) has a morph transition
                // (slide-only — master/layout shapes are not part of slide-to-slide
                // morph continuity).
                if (slidePart != null && slideParts != null
                    && !shapeName.StartsWith("!!") && !shapeName.StartsWith("TextBox ") && !shapeName.StartsWith("Content ") && shapeName != "")
                {
                    if (SlideHasMorphContext(slidePart, slideParts))
                        shapeName = "!!" + shapeName;
                }

                // Classify: explicit `--type textbox` always produces a textbox
                // (writes <p:cNvSpPr txBox="1"/>). For `--type shape` (and the
                // legacy default where the dispatch doesn't pass a hint), fall
                // back to a heuristic: explicit geometry (shape=/preset=/
                // geometry=/customGeometryXml=) → real shape; bare `text=` →
                // textbox shorthand; otherwise → real shape (matches the
                // `--type shape` direct intuition).
                var hasGeometryProp = properties.ContainsKey("shape")
                    || properties.ContainsKey("preset")
                    || properties.ContainsKey("geometry")
                    || properties.ContainsKey("customGeometryXml");
                var hasTextProp = properties.ContainsKey("text");
                // `--type textbox` is the explicit textbox path and always
                // wins (covers the dump-emitter replay case where text is
                // split into separate paragraph/run adds, so hasTextProp here
                // is false). `--type shape` keeps the legacy "text= shorthand"
                // — bare text with no geometry is still classified as
                // textbox, since users (and earlier tests) treat that as the
                // textbox shortcut.
                bool isTextBoxFlavor = elementTypeHint == "textbox"
                    || (!hasGeometryProp && hasTextProp);
                var newShape = CreateTextShape(shapeId, shapeName, text, false, isTextBoxFlavor);

                // CONSISTENCY(splocks-round-trip): preserve <a:spLocks noGrp="1"/>
                // when dump captured it. Regular shapes default to no lock
                // (PowerPoint authors them without spLocks); only emit when
                // the prop is explicitly truthy.
                if ((properties.TryGetValue("noGrp", out var shNoGrpStr)
                        || properties.TryGetValue("nogrp", out shNoGrpStr))
                    && IsTruthy(shNoGrpStr))
                {
                    var cNvSp = newShape.NonVisualShapeProperties?.NonVisualShapeDrawingProperties;
                    if (cNvSp != null)
                    {
                        cNvSp.RemoveAllChildren<Drawing.ShapeLocks>();
                        cNvSp.AppendChild(new Drawing.ShapeLocks { NoGrouping = true });
                    }
                }

                // CONSISTENCY(font-dotted-alias): mirror Set's font.<attr> aliases
                // (commit 80fb739e). Without these, `add shape --prop font.name=Arial`
                // silently dropped while `set --prop font.name=Arial` succeeded.
                //
                // Shape-level run defaults must also survive when the shape was
                // created without inline text. CreateTextShape emits a runless
                // `<a:p/>` for empty-text shapes, so the foreach-Run loops below
                // had no targets and silently dropped bold / size / color / font
                // on `add shape --prop bold=false --prop size=68pt …`. Fall
                // back to a seeded endParaRPr (same CT_TextCharacterProperties
                // base type as RunProperties — FontSize / Bold / Italic /
                // SolidFill child / LatinFont / EastAsianFont apply uniformly)
                // on the first paragraph so the props round-trip through
                // dump→replay.
                List<OpenXmlCompositeElement> RunPropTargets()
                {
                    var targets = new List<OpenXmlCompositeElement>();
                    var runs = newShape.Descendants<Drawing.Run>().ToList();
                    if (runs.Count > 0)
                    {
                        foreach (var r in runs)
                            targets.Add(r.RunProperties ?? (r.RunProperties = new Drawing.RunProperties()));
                        return targets;
                    }
                    var firstPara = newShape.TextBody?.Elements<Drawing.Paragraph>().FirstOrDefault();
                    if (firstPara == null) return targets;
                    var endRPr = firstPara.GetFirstChild<Drawing.EndParagraphRunProperties>();
                    if (endRPr == null)
                    {
                        endRPr = new Drawing.EndParagraphRunProperties { Language = "en-US" };
                        firstPara.AppendChild(endRPr);
                    }
                    targets.Add(endRPr);
                    return targets;
                }
                static void SetFontSize(OpenXmlCompositeElement rPr, int sizeVal)
                {
                    if (rPr is Drawing.RunProperties r) r.FontSize = sizeVal;
                    else if (rPr is Drawing.EndParagraphRunProperties e) e.FontSize = sizeVal;
                }
                static void SetBold(OpenXmlCompositeElement rPr, bool b)
                {
                    if (rPr is Drawing.RunProperties r) r.Bold = b;
                    else if (rPr is Drawing.EndParagraphRunProperties e) e.Bold = b;
                }
                static void SetItalic(OpenXmlCompositeElement rPr, bool i)
                {
                    if (rPr is Drawing.RunProperties r) r.Italic = i;
                    else if (rPr is Drawing.EndParagraphRunProperties e) e.Italic = i;
                }
                if (properties.TryGetValue("size", out var sizeStr)
                    || properties.TryGetValue("fontSize", out sizeStr)
                    || properties.TryGetValue("fontsize", out sizeStr)
                    || properties.TryGetValue("font.size", out sizeStr))
                {
                    var sizeVal = (int)Math.Round(ParseFontSize(sizeStr) * 100);
                    foreach (var rProps in RunPropTargets()) SetFontSize(rProps, sizeVal);
                }
                if (properties.TryGetValue("bold", out var boldStr)
                    || properties.TryGetValue("font.bold", out boldStr))
                {
                    var isBold = IsTruthy(boldStr);
                    foreach (var rProps in RunPropTargets()) SetBold(rProps, isBold);
                }
                if (properties.TryGetValue("italic", out var italicStr)
                    || properties.TryGetValue("font.italic", out italicStr))
                {
                    var isItalic = IsTruthy(italicStr);
                    foreach (var rProps in RunPropTargets()) SetItalic(rProps, isItalic);
                }
                if (properties.TryGetValue("color", out var colorVal)
                    || properties.TryGetValue("font.color", out colorVal))
                {
                    foreach (var rProps in RunPropTargets())
                    {
                        rProps.RemoveAllChildren<Drawing.SolidFill>();
                        var solidFill = BuildSolidFill(colorVal);
                        if (!rProps.AddChild(solidFill, throwOnError: false))
                            rProps.AppendChild(solidFill);
                    }
                }

                // Schema order: font (latin/ea) after fill
                if (properties.TryGetValue("font", out var font)
                    || properties.TryGetValue("font.name", out font))
                {
                    foreach (var rProps in RunPropTargets())
                    {
                        rProps.Append(new Drawing.LatinFont { Typeface = font });
                        rProps.Append(new Drawing.EastAsianFont { Typeface = font });
                        ReorderDrawingRunProperties(rProps);
                    }
                }
                // Per-script font slots — used for Japanese/Korean/Arabic when
                // the bare 'font' would clobber an existing scheme. Schema
                // order is enforced below via ReorderDrawingRunProperties.
                if (properties.TryGetValue("font.latin", out var fontLatin))
                {
                    foreach (var rProps in RunPropTargets())
                    {
                        rProps.RemoveAllChildren<Drawing.LatinFont>();
                        rProps.Append(new Drawing.LatinFont { Typeface = fontLatin });
                        ReorderDrawingRunProperties(rProps);
                    }
                }
                if (properties.TryGetValue("font.ea", out var fontEa)
                    || properties.TryGetValue("font.eastasia", out fontEa)
                    || properties.TryGetValue("font.eastasian", out fontEa))
                {
                    foreach (var rProps in RunPropTargets())
                    {
                        rProps.RemoveAllChildren<Drawing.EastAsianFont>();
                        rProps.Append(new Drawing.EastAsianFont { Typeface = fontEa });
                        ReorderDrawingRunProperties(rProps);
                    }
                }
                if (properties.TryGetValue("font.cs", out var fontCs)
                    || properties.TryGetValue("font.complexscript", out fontCs)
                    || properties.TryGetValue("font.complex", out fontCs))
                {
                    foreach (var rProps in RunPropTargets())
                    {
                        rProps.RemoveAllChildren<Drawing.ComplexScriptFont>();
                        rProps.Append(new Drawing.ComplexScriptFont { Typeface = fontCs });
                        ReorderDrawingRunProperties(rProps);
                    }
                }
                // Reading direction (Arabic/Hebrew). Sets BOTH <a:pPr rtl="1"/>
                // (per-paragraph character order) AND <a:bodyPr rtlCol="1"/>
                // (textbox column direction) so a fresh shape created with
                // direction=rtl is fully RTL-correct end to end.
                if (properties.TryGetValue("direction", out var dirVal)
                    || properties.TryGetValue("dir", out dirVal)
                    || properties.TryGetValue("rtl", out dirVal))
                {
                    bool rtl = ParsePptDirectionRtl(dirVal);
                    foreach (var para in newShape.TextBody?.Elements<Drawing.Paragraph>() ?? Enumerable.Empty<Drawing.Paragraph>())
                    {
                        var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                        // Clear semantics: direction=ltr strips the rtl attribute
                        // rather than writing rtl="0" on every fresh paragraph.
                        if (rtl) pProps.RightToLeft = true;
                        else pProps.RightToLeft = null;
                    }
                    var dirBodyPr = newShape.TextBody?.Elements<Drawing.BodyProperties>().FirstOrDefault();
                    // For ltr (schema default), strip the attribute rather
                    // than writing rtlCol="0" — keeps the XML free of
                    // explicit-default noise on rtl→ltr toggles.
                    if (dirBodyPr != null)
                    {
                        if (rtl)
                            dirBodyPr.SetAttribute(new DocumentFormat.OpenXml.OpenXmlAttribute("", "rtlCol", "", "1"));
                        else
                            dirBodyPr.RemoveAttribute("rtlCol", "");
                    }
                }

                // Text margin (padding inside shape)
                if (properties.TryGetValue("margin", out var marginVal))
                {
                    var bodyPr = newShape.TextBody?.Elements<Drawing.BodyProperties>().FirstOrDefault();
                    if (bodyPr != null)
                        ApplyTextMargin(bodyPr, marginVal);
                }

                // Verbatim shape-level <a:lstStyle> re-injection. The default
                // txBody carries an empty <a:lstStyle/> stub; the source's
                // per-level lnSpc/defTabSz/algn/fonts live only in the captured
                // OuterXml (NodeBuilder lstStyleRaw). Replace the stub with the
                // parsed verbatim element so text reflow off the source metrics
                // is preserved. CT_TextBody order: bodyPr, lstStyle, p+ — replace
                // in place keeps lstStyle after bodyPr and before the first p.
                if (properties.TryGetValue("lstStyleRaw", out var lstStyleRawVal)
                    && !string.IsNullOrWhiteSpace(lstStyleRawVal)
                    && newShape.TextBody != null)
                {
                    var newLstStyle = new Drawing.ListStyle(lstStyleRawVal);
                    var existingLstStyle = newShape.TextBody.GetFirstChild<Drawing.ListStyle>();
                    if (existingLstStyle != null)
                    {
                        existingLstStyle.InsertAfterSelf(newLstStyle);
                        existingLstStyle.Remove();
                    }
                    else
                    {
                        var anchorBodyPr = newShape.TextBody.GetFirstChild<Drawing.BodyProperties>();
                        if (anchorBodyPr != null) anchorBodyPr.InsertAfterSelf(newLstStyle);
                        else newShape.TextBody.InsertAt(newLstStyle, 0);
                    }
                }

                // Text alignment (horizontal)
                if (properties.TryGetValue("align", out var alignVal))
                {
                    var alignment = ParseTextAlignment(alignVal);
                    foreach (var para in newShape.TextBody?.Elements<Drawing.Paragraph>() ?? Enumerable.Empty<Drawing.Paragraph>())
                    {
                        var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                        pProps.Alignment = alignment;
                    }
                }

                // CJK / line-break pPr attributes on the shape's first-paragraph
                // (eaLnBrk / latinLnBrk / fontAlgn / defTabSz). Shapes/textboxes
                // built with an inline text= seed their first paragraph here, not
                // through AddParagraph — without this the dropped attributes
                // rewrapped CJK text on round-trip. Apply to every paragraph.
                foreach (var pBreakKey in new[] { "eaLnBrk", "latinLnBrk", "fontAlgn", "defTabSz" })
                {
                    if (!properties.TryGetValue(pBreakKey, out var pBreakVal)) continue;
                    foreach (var para in newShape.TextBody?.Elements<Drawing.Paragraph>() ?? Enumerable.Empty<Drawing.Paragraph>())
                    {
                        var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                        ApplyParagraphBreakProp(pProps, pBreakKey, pBreakVal);
                    }
                }

                // Vertical alignment
                if (properties.TryGetValue("valign", out var valignVal))
                {
                    var bodyPr = newShape.TextBody?.Elements<Drawing.BodyProperties>().FirstOrDefault();
                    if (bodyPr != null)
                    {
                        bodyPr.Anchor = valignVal.ToLowerInvariant() switch
                        {
                            "top" or "t" => Drawing.TextAnchoringTypeValues.Top,
                            "center" or "middle" or "c" or "m" => Drawing.TextAnchoringTypeValues.Center,
                            "bottom" or "b" => Drawing.TextAnchoringTypeValues.Bottom,
                            _ => throw new ArgumentException($"Invalid valign: {valignVal}. Use top/center/bottom")
                        };
                    }
                }

                // Rotation
                if (properties.TryGetValue("rotation", out var rotStr) || properties.TryGetValue("rotate", out rotStr))
                {
                    // Will be set on Transform2D below
                }

                // Underline
                if (properties.TryGetValue("underline", out var ulVal)
                    || properties.TryGetValue("font.underline", out ulVal))
                {
                    var ulEnum = ulVal.ToLowerInvariant() switch
                    {
                        "true" or "single" or "sng" => Drawing.TextUnderlineValues.Single,
                        "double" or "dbl" => Drawing.TextUnderlineValues.Double,
                        "heavy" => Drawing.TextUnderlineValues.Heavy,
                        "dotted" => Drawing.TextUnderlineValues.Dotted,
                        "dash" => Drawing.TextUnderlineValues.Dash,
                        "wavy" => Drawing.TextUnderlineValues.Wavy,
                        "false" or "none" => Drawing.TextUnderlineValues.None,
                        _ => throw new ArgumentException($"Invalid underline value: '{ulVal}'. Valid values: single, double, heavy, dotted, dash, wavy, none.")
                    };
                    foreach (var rProps in RunPropTargets())
                    {
                        if (rProps is Drawing.RunProperties rp) rp.Underline = ulEnum;
                        else if (rProps is Drawing.EndParagraphRunProperties ep) ep.Underline = ulEnum;
                    }
                }

                // Underline color — mirrors Set ShapeProperties.cs:436. Writes
                // <a:uFill><a:solidFill><a:srgbClr val="…"/></a:solidFill></a:uFill>
                // on every run; ReorderDrawingRunProperties keeps the schema
                // order (uLn before uFill before latin).
                if (properties.TryGetValue("underline.color", out var ulColorVal)
                    || properties.TryGetValue("underlineColor", out ulColorVal)
                    || properties.TryGetValue("underlinecolor", out ulColorVal)
                    || properties.TryGetValue("font.underline.color", out ulColorVal))
                {
                    var ulHex = OfficeCli.Core.ParseHelpers.SanitizeColorForOoxml(ulColorVal).Rgb;
                    foreach (var rProps in RunPropTargets())
                    {
                        rProps.RemoveAllChildren<Drawing.UnderlineFill>();
                        rProps.RemoveAllChildren<Drawing.UnderlineFillText>();
                        rProps.AppendChild(new Drawing.UnderlineFill(
                            new Drawing.SolidFill(new Drawing.RgbColorModelHex { Val = ulHex })));
                        ReorderDrawingRunProperties(rProps);
                    }
                }

                // Strikethrough
                if (properties.TryGetValue("strikethrough", out var stVal)
                    || properties.TryGetValue("strike", out stVal)
                    || properties.TryGetValue("font.strike", out stVal)
                    || properties.TryGetValue("font.strikethrough", out stVal))
                {
                    var stEnum = stVal.ToLowerInvariant() switch
                    {
                        "true" or "single" => Drawing.TextStrikeValues.SingleStrike,
                        "double" => Drawing.TextStrikeValues.DoubleStrike,
                        "false" or "none" => Drawing.TextStrikeValues.NoStrike,
                        _ => throw new ArgumentException($"Invalid strikethrough value: '{stVal}'. Valid values: single, double, none.")
                    };
                    foreach (var rProps in RunPropTargets())
                    {
                        if (rProps is Drawing.RunProperties rp) rp.Strike = stEnum;
                        else if (rProps is Drawing.EndParagraphRunProperties ep) ep.Strike = stEnum;
                    }
                }

                // Caps (allCaps / smallCaps / cap=all|small|none)
                // CONSISTENCY(allcaps-alias): mirror Word commit ccaed17a;
                // accept allCaps/allcaps/smallCaps/smallcaps as run-level rPr cap.
                {
                    string? capValue = null;
                    if (properties.TryGetValue("cap", out var rawCap)) capValue = rawCap;
                    else if (properties.TryGetValue("allCaps", out var allCaps)
                          || properties.TryGetValue("allcaps", out allCaps))
                        capValue = (allCaps is "0" or "false" or "False" or "none") ? "none" : "all";
                    else if (properties.TryGetValue("smallCaps", out var smallCaps)
                          || properties.TryGetValue("smallcaps", out smallCaps))
                        capValue = (smallCaps is "0" or "false" or "False" or "none") ? "none" : "small";

                    if (capValue != null)
                    {
                        // ST_TextCapsType enum is lowercase {none, small, all}.
                        // Mixed-case input ("SMALL", "ALL") written verbatim
                        // produces schema-invalid OOXML — PowerPoint then
                        // refuses to open the file. Normalize on write.
                        capValue = capValue.ToLowerInvariant();
                        if (capValue is not ("none" or "small" or "all"))
                            throw new ArgumentException($"Invalid cap value: '{capValue}'. Valid values: none, small, all.");
                        foreach (var run in newShape.Descendants<Drawing.Run>())
                        {
                            var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                            rProps.SetAttribute(new OpenXmlAttribute("", "cap", "", capValue));
                        }
                    }
                }

                // Line spacing
                if (properties.TryGetValue("lineSpacing", out var lsVal) || properties.TryGetValue("linespacing", out lsVal))
                {
                    var (lsInternal, lsIsPercent) = SpacingConverter.ParsePptLineSpacing(lsVal);
                    foreach (var para in newShape.TextBody?.Elements<Drawing.Paragraph>() ?? Enumerable.Empty<Drawing.Paragraph>())
                    {
                        var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                        pProps.RemoveAllChildren<Drawing.LineSpacing>();
                        if (lsIsPercent)
                            pProps.AppendChild(new Drawing.LineSpacing(
                                new Drawing.SpacingPercent { Val = lsInternal }));
                        else
                            pProps.AppendChild(new Drawing.LineSpacing(
                                new Drawing.SpacingPoints { Val = lsInternal }));
                    }
                }

                // Space before/after
                if (properties.TryGetValue("spaceBefore", out var sbVal) || properties.TryGetValue("spacebefore", out sbVal))
                {
                    var sbInternal = SpacingConverter.ParsePptSpacing(sbVal);
                    foreach (var para in newShape.TextBody?.Elements<Drawing.Paragraph>() ?? Enumerable.Empty<Drawing.Paragraph>())
                    {
                        var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                        pProps.RemoveAllChildren<Drawing.SpaceBefore>();
                        pProps.AppendChild(new Drawing.SpaceBefore(new Drawing.SpacingPoints { Val = sbInternal }));
                    }
                }
                if (properties.TryGetValue("spaceAfter", out var saVal) || properties.TryGetValue("spaceafter", out saVal))
                {
                    var saInternal = SpacingConverter.ParsePptSpacing(saVal);
                    foreach (var para in newShape.TextBody?.Elements<Drawing.Paragraph>() ?? Enumerable.Empty<Drawing.Paragraph>())
                    {
                        var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                        pProps.RemoveAllChildren<Drawing.SpaceAfter>();
                        pProps.AppendChild(new Drawing.SpaceAfter(new Drawing.SpacingPoints { Val = saInternal }));
                    }
                }

                // AutoFit
                if (properties.TryGetValue("autofit", out var afVal))
                {
                    var bodyPr = newShape.TextBody?.Elements<Drawing.BodyProperties>().FirstOrDefault();
                    if (bodyPr != null)
                    {
                        switch (afVal.ToLowerInvariant())
                        {
                            // R10-4: 'shrink'/'true' alias normAutofit (PowerPoint's
                            // shrink-text-on-overflow mode).
                            case "true" or "shrink" or "normal": bodyPr.AppendChild(ApplyNormalAutoFitScale(new Drawing.NormalAutoFit(), properties)); break;
                            case "shape": bodyPr.AppendChild(new Drawing.ShapeAutoFit()); break;
                            case "false" or "none": bodyPr.AppendChild(new Drawing.NoAutoFit()); break;
                        }
                    }
                }

                // Position and size (in EMU, 1cm = 360000 EMU; or parse as cm/in)
                {
                    long xEmu = 0, yEmu = 0;
                    long cxEmu = 3600000, cyEmu = 1800000; // default: 10cm x 5cm (avoid full-slide overlap when width unspecified)
                    // Unified bounds check: PowerPoint truncates EMU coordinates
                    // past INT32_MAX (cx/cy schema-typed as int32 in practice).
                    // Error prefix "Invalid" so OutputFormatter routes the
                    // ArgumentException to invalid_value, mirroring Set's path.
                    static long ParseEmuBounded(string raw, string field, bool allowNegative)
                    {
                        var v = ParseEmu(raw);
                        if (!allowNegative && v < 0)
                            throw new ArgumentException($"Invalid {field} '{raw}': negative values are not allowed.");
                        if (v > int.MaxValue)
                            throw new ArgumentException($"Invalid {field} '{raw}': exceeds the maximum supported shape coordinate (INT32_MAX EMU).");
                        if (allowNegative && v < int.MinValue)
                            throw new ArgumentException($"Invalid {field} '{raw}': below the minimum supported shape coordinate (INT32_MIN EMU).");
                        return v;
                    }
                    if (properties.TryGetValue("x", out var xStr) || properties.TryGetValue("left", out xStr)) xEmu = ParseEmuBounded(xStr, "x", allowNegative: true);
                    if (properties.TryGetValue("y", out var yStr) || properties.TryGetValue("top", out yStr)) yEmu = ParseEmuBounded(yStr, "y", allowNegative: true);
                    if (properties.TryGetValue("width", out var wStr) || properties.TryGetValue("w", out wStr))
                    {
                        // Zero is legitimate (PowerPoint hides 0×0 shapes — used for invisible
                        // decorative lines). Dump-replay must round-trip zero-sized shapes that
                        // were authored that way; the prior strict reject broke real-world
                        // round-trips.
                        cxEmu = ParseEmuBounded(wStr, "width", allowNegative: false);
                    }
                    if (properties.TryGetValue("height", out var hStr) || properties.TryGetValue("h", out hStr))
                    {
                        cyEmu = ParseEmuBounded(hStr, "height", allowNegative: false);
                    }

                    var xfrm = new Drawing.Transform2D
                    {
                        Offset = new Drawing.Offset { X = xEmu, Y = yEmu },
                        Extents = new Drawing.Extents { Cx = cxEmu, Cy = cyEmu }
                    };
                    if (properties.TryGetValue("rotation", out var rotVal) || properties.TryGetValue("rotate", out rotVal))
                    {
                        var rotDbl = ParseHelpers.SafeParseRotationDegrees(rotVal!, "rotation");
                        xfrm.Rotation = (int)(rotDbl * 60000);
                    }
                    newShape.ShapeProperties!.Transform2D = xfrm;

                    // Custom geometry takes precedence — replay of dump-emitted
                    // custom-shape paths comes through as customGeometryXml (raw
                    // OOXML <a:custGeom>) which we splice in verbatim. Fallback
                    // to preset name when no custom-geometry signal is present.
                    if (properties.TryGetValue("customGeometryXml", out var custXml) && custXml.Length > 0)
                    {
                        // Parse the raw <a:custGeom> string and load it through
                        // OpenXmlReader so child namespaces resolve against the
                        // root scope ONCE rather than being re-declared per
                        // child. The earlier InnerXml-copy approach retained
                        // a redundant xmlns:a on every direct child (avLst,
                        // gdLst, ahLst, cxnLst, rect, pathLst) — the SDK
                        // then wrote the bloated attributes back to slide
                        // XML and every dump→replay round trip doubled the
                        // byte count of the custGeom block (test-samples/
                        // a.pptx slide5: 1.2KB → 2.4KB → 4.8KB across three
                        // round trips).
                        var parsed = new Drawing.CustomGeometry(custXml);
                        newShape.ShapeProperties.AppendChild(parsed);
                    }
                    else
                    {
                        var presetName = properties.TryGetValue("preset", out var pn) ? pn
                            : properties.TryGetValue("geometry", out pn) ? pn
                            : properties.GetValueOrDefault("shape", "rect");
                        // "custom" is a Get-side marker for "this was custGeom but we
                        // couldn't round-trip the path" — degrade to rect rather than
                        // erroring out (we'd lose the shape entirely otherwise).
                        if (presetName.Equals("custom", StringComparison.OrdinalIgnoreCase))
                            presetName = "rect";
                        // Validate the preset name so an unknown geometry
                        // surfaces unsupported_property instead of silently
                        // degrading to rect. Mirrors the Set path
                        // (ShapeProperties.cs uses TryParsePresetShape for
                        // the same reason).
                        if (!TryParsePresetShape(presetName, out var presetGeom))
                            throw new ArgumentException($"Unknown shape geometry: '{presetName}'");
                        var avLstForAdd = new Drawing.AdjustValueList();
                        // CONSISTENCY(preset-adj-handles): NodeBuilder surfaces
                        // <a:avLst><a:gd name="adj" fmla="val N"/>...</a:avLst>
                        // as the canonical `adj=name:fmla,...` Format key (the
                        // form "adj=adj1:val 6000,adj2:val 12000" round-trips
                        // through dump→replay). Repopulate <a:gd> children
                        // when the prop bag carries adj=...; without this,
                        // every preset reverts to its OOXML default
                        // proportions even though dump captured the values.
                        if (properties.TryGetValue("adj", out var adjSpec)
                            && !string.IsNullOrWhiteSpace(adjSpec))
                        {
                            ApplyAdjustHandles(avLstForAdd, adjSpec, presetGeom);
                        }
                        newShape.ShapeProperties.AppendChild(
                            new Drawing.PresetGeometry(avLstForAdd) { Preset = presetGeom }
                        );
                    }
                }

                // Shape fill (after xfrm and prstGeom to maintain schema order).
                // 'background' is an alias for 'fill' (symmetric with the Set
                // branch in ShapeProperties.cs case "background"), unless a
                // pattern= is present — in that case it is folded into the
                // pattFill bg below, matching Set's pattern precedence.
                if (properties.TryGetValue("fill", out var fillVal)
                    || (!properties.ContainsKey("pattern")
                        && properties.TryGetValue("background", out fillVal)))
                {
                    ApplyShapeFill(newShape.ShapeProperties!, fillVal);
                }

                // Gradient fill
                if (properties.TryGetValue("gradient", out var gradVal))
                {
                    ApplyGradientFill(newShape.ShapeProperties!, gradVal);
                }

                // bt-7: AddShape consumes gradientRaw inline so the verbatim
                // <a:gradFill flip=…><a:tileRect/></a:gradFill> emitted by
                // NodeBuilder (when HasGradientNonSemanticTuning returns true)
                // re-installs at create time. Without this branch the key
                // round-tripped through the prop bag, AddShape silently
                // ignored it, and replay decks lost flip / tileRect attrs the
                // semantic gradient= form can't carry.
                if (properties.TryGetValue("gradientRaw", out var gradRawVal)
                    || properties.TryGetValue("gradientraw", out gradRawVal))
                {
                    ApplyGradientRaw(newShape.ShapeProperties!, gradRawVal);
                }

                // Pattern fill (mutually exclusive with fill/gradient — last one wins, following fill/gradient convention)
                if (properties.TryGetValue("pattern", out var patternVal))
                {
                    // R9b: when pattern= is given alongside fill= / background=,
                    // fold those into the pattFill fg/bg (preset:fg:bg form) so a
                    // separate fill= SolidFill isn't silently overwritten and lost.
                    // Explicit colors inside the pattern value (preset:fg:bg) win.
                    if (!patternVal.Contains(':'))
                    {
                        var fgProp = properties.TryGetValue("fill", out var fgv) && !string.IsNullOrWhiteSpace(fgv) ? fgv.Trim() : null;
                        var bgProp = properties.TryGetValue("background", out var bgv) && !string.IsNullOrWhiteSpace(bgv) ? bgv.Trim() : null;
                        if (fgProp != null || bgProp != null)
                            patternVal = $"{patternVal}:{fgProp ?? "000000"}:{bgProp ?? "FFFFFF"}";
                    }
                    ApplyPatternFill(newShape.ShapeProperties!, patternVal);
                }

                // Opacity (alpha on fill) — uses <a:alpha val="N"/>
                // Must come after gradient so it can apply to gradient stops too.
                // Alpha must attach to a color element inside a fill carrier; if
                // the caller gave 'opacity' without any fill/gradient/pattern,
                // the value has nothing to bind to. Per schemas/help/pptx/shape.json
                // 'opacity.requires: ["fill"]', reject rather than silently drop.
                if (properties.TryGetValue("opacity", out var opacityVal))
                {
                    var hasFillCarrier =
                        properties.ContainsKey("fill") ||
                        properties.ContainsKey("gradient") ||
                        properties.ContainsKey("pattern") ||
                        (newShape.ShapeProperties?.GetFirstChild<Drawing.SolidFill>() != null) ||
                        (newShape.ShapeProperties?.GetFirstChild<Drawing.GradientFill>() != null) ||
                        (newShape.ShapeProperties?.GetFirstChild<Drawing.PatternFill>() != null);
                    if (!hasFillCarrier)
                        throw new ArgumentException(
                            $"'opacity'='{opacityVal}' requires a fill carrier. Provide one of 'fill' / 'gradient' / 'pattern' " +
                            "so the alpha value has a color element to attach to.");
                    if (double.TryParse(opacityVal, System.Globalization.CultureInfo.InvariantCulture, out var alphaNum))
                    {
                        // CONSISTENCY(opacity-clamp): (1, 2) ambiguous; see
                        // the shape Set path. Reject before the /100.
                        if (alphaNum > 1.0 && alphaNum < 2.0)
                            throw new ArgumentException(
                                $"Invalid 'opacity' value: '{opacityVal}'. Expected 0.0-1.0 as decimal or 2-100 as percent (values in (1, 2) are ambiguous).");
                        if (alphaNum > 1.0) alphaNum /= 100.0; // treat >=2 as percentage (e.g. 30 → 0.30)
                        if (alphaNum < 0.0 || alphaNum > 1.0)
                            throw new ArgumentException(
                                $"Invalid 'opacity' value: '{opacityVal}'. Expected 0.0-1.0 (or 0-100 as percent).");
                        var alphaPct = (int)(alphaNum * 100000);
                        var solidFill = newShape.ShapeProperties?.GetFirstChild<Drawing.SolidFill>();
                        if (solidFill != null)
                        {
                            var colorEl = solidFill.GetFirstChild<Drawing.RgbColorModelHex>() as OpenXmlElement
                                ?? solidFill.GetFirstChild<Drawing.SchemeColor>();
                            if (colorEl != null)
                            {
                                colorEl.RemoveAllChildren<Drawing.Alpha>();
                                colorEl.AppendChild(new Drawing.Alpha { Val = alphaPct });
                            }
                        }
                        var gradientFill = newShape.ShapeProperties?.GetFirstChild<Drawing.GradientFill>();
                        if (gradientFill != null)
                        {
                            foreach (var stop in gradientFill.Descendants<Drawing.GradientStop>())
                            {
                                var stopColor = stop.GetFirstChild<Drawing.RgbColorModelHex>() as OpenXmlElement
                                    ?? stop.GetFirstChild<Drawing.SchemeColor>();
                                if (stopColor != null)
                                {
                                    stopColor.RemoveAllChildren<Drawing.Alpha>();
                                    stopColor.AppendChild(new Drawing.Alpha { Val = alphaPct });
                                }
                            }
                        }
                    }
                }

                // Line/border (after fill per schema: xfrm → prstGeom → fill → ln)
                // Schema documents compound form 'color[:width[:style]]'
                // (schemas/help/_shared/shape.json) — split here so the
                // single-part code paths handle each component uniformly.
                string? compoundLineWidth = null;
                string? compoundLineDash = null;
                if (properties.TryGetValue("line", out var lineColor) || properties.TryGetValue("linecolor", out lineColor) || properties.TryGetValue("lineColor", out lineColor) || properties.TryGetValue("line.color", out lineColor) || properties.TryGetValue("border", out lineColor) || properties.TryGetValue("border.color", out lineColor))
                {
                    (lineColor, compoundLineWidth, compoundLineDash) = SplitCompoundLineValue(lineColor);
                    var outline = EnsureOutline(newShape.ShapeProperties!);
                    if (lineColor.Equals("none", StringComparison.OrdinalIgnoreCase))
                        outline.AppendChild(new Drawing.NoFill());
                    else
                        outline.AppendChild(BuildSolidFill(lineColor));
                }
                // styledLine: the dump signals a <p:style> raw-set follows and
                // no explicit line colour was captured — the stroke colour comes
                // from lnRef, so the default-black fill injection must be
                // skipped (stress013's theme-tinted borders replayed black).
                bool shStyledLine = IsTruthy(properties.GetValueOrDefault("styledLine"))
                                    || IsTruthy(properties.GetValueOrDefault("styledline"));
                if (properties.TryGetValue("linewidth", out var lwStr) || properties.TryGetValue("lineWidth", out lwStr) || properties.TryGetValue("line.width", out lwStr) || properties.TryGetValue("border.width", out lwStr))
                {
                    var outline = EnsureOutline(newShape.ShapeProperties!);
                    outline.Width = Core.EmuConverter.ParseLineWidth(lwStr);
                    if (!shStyledLine) EnsureOutlineHasFill(outline);
                }
                else if (compoundLineWidth != null)
                {
                    var outline = EnsureOutline(newShape.ShapeProperties!);
                    outline.Width = Core.EmuConverter.ParseLineWidth(compoundLineWidth);
                    if (!shStyledLine) EnsureOutlineHasFill(outline);
                }
                // Stash the compound dash so the lineDash branch in
                // SetRunOrShapeProperties below picks it up via the
                // shared effectProps dispatch.
                if (compoundLineDash != null
                    && !properties.ContainsKey("linedash")
                    && !properties.ContainsKey("lineDash")
                    && !properties.ContainsKey("line.dash"))
                {
                    properties["lineDash"] = compoundLineDash;
                }

                // Outline policy: "user didn't ask = we don't write". Earlier the
                // handler auto-injected a 0.75pt #595959 outline whenever the caller
                // picked a geometry and gave no fill+line, mimicking PowerPoint's
                // "Insert Shape" UI default. That phantom border survived through
                // dump→replay: NodeBuilder reported lineColor=595959, the batch
                // emitter forwarded it, and every round-trip grew a darker border on
                // a shape the user never asked to outline. The visibility regression
                // (presets render with no stroke) is the lesser harm; defer the
                // default-outline UX to a caller-driven `line=default`/UI layer.

                // List style (bullet/numbered). bulletRaw (full bullet group)
                // wins over the lossy `list` keyword when both are present.
                // Probe both unconditionally (tracking): an else-if left `list`
                // unread when bulletRaw was present, warning a false
                // unsupported_property on dump replays that emit both.
                var hasShBulletRaw = properties.TryGetValue("bulletRaw", out var shBulletRaw) || properties.TryGetValue("bulletraw", out shBulletRaw);
                var hasShList = properties.TryGetValue("list", out var listVal) || properties.TryGetValue("liststyle", out listVal) || properties.TryGetValue("bullet", out listVal);
                if (hasShBulletRaw)
                {
                    foreach (var para in newShape.TextBody?.Elements<Drawing.Paragraph>() ?? Enumerable.Empty<Drawing.Paragraph>())
                    {
                        var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                        ApplyBulletRaw(pProps, shBulletRaw!);
                    }
                }
                else if (hasShList)
                {
                    foreach (var para in newShape.TextBody?.Elements<Drawing.Paragraph>() ?? Enumerable.Empty<Drawing.Paragraph>())
                    {
                        var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                        ApplyListStyle(pProps, listVal!, preserveIndent: properties.ContainsKey("indent") || properties.ContainsKey("marginLeft") || properties.ContainsKey("marginleft") || properties.ContainsKey("marL") || properties.ContainsKey("marl"));
                    }
                }

                if (insertContainer != null)
                {
                    // Group container — InsertAtPosition over a GroupShape: the
                    // existing helper is shape-tree generic so it works here too.
                    InsertAtPosition(insertContainer, newShape, index);
                }
                else
                {
                    InsertAtPosition(shapeTree, newShape, index);
                }

                // Hyperlink on shape — slide-only. ApplyShapeHyperlink uses
                // SlidePart.AddHyperlinkRelationship; master/layout owner parts
                // would need a parallel relationship API. Out of scope for the
                // initial master/layout shape support.
                if (properties.TryGetValue("link", out var linkVal))
                {
                    if (slidePart == null)
                        throw new ArgumentException(
                            "'link' is not yet supported on master/layout shapes — set it on slide-level shapes instead.");
                    var tooltipVal = properties.GetValueOrDefault("tooltip");
                    ApplyShapeHyperlink(slidePart, newShape, linkVal, tooltipVal);
                }

                // lineDash, effects, 3D, flip — delegate to SetRunOrShapeProperties
                var effectKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                    { "linedash", "line.dash", "linedashraw", "line.dashraw", "shadow", "shadowraw", "innershadow", "innershadowraw", "glow",
                      "reflection", "reflectionraw", "filloverlayraw", "effectdagraw", "effectsraw",
                      "softedge", "blur", "fliph", "flipv", "rot3d", "rotation3d",
                      "rotx", "roty", "rotz", "bevel", "beveltop", "bevelbottom",
                      "depth", "extrusion", "material", "lighting", "lightrig",
                      "lightingdir", "lightrigdir", "lightingrot", "lightrigrot",
                      "camera", "camerapreset", "cameraprst",
                      "extrusioncolor", "extrusionclr",
                      "contourcolor", "contourclr",
                      "spacing", "charspacing", "letterspacing",
                      "indent", "marginleft", "marl", "marginright", "marr",
                      "textfill", "textgradient", "geometry",
                      "baseline", "superscript", "subscript",
                      "textwarp", "wordart", "autofit",
                      // WordArt raw forms — verbatim prstTxWarp (keeps avLst)
                      // and the bodyPr-level 3D-text scene3d/sp3d pair.
                      "textwarpraw", "textWarpRaw",
                      "textscene3draw", "textScene3dRaw",
                      "textsp3draw", "textSp3dRaw",
                      // shrink-on-overflow scale — consumed alongside autofit=normal
                      "fontScale", "fontscale", "lnSpcReduction", "lnspcreduction",
                      "wrap", "wordwrap", "anchorCtr", "anchorctr", "upright",
                      "columns", "numcol", "columnSpacing", "columnspacing",
                      "vertOverflow", "vertoverflow", "horzOverflow", "horzoverflow",
                      "lineopacity", "line.opacity",
                      "linegradient", "line.gradient",
                      // previously dropped silently — route through Set
                      // so OOXML attributes actually get emitted.
                      "linecap", "lineCap", "line.cap",
                      "linejoin", "lineJoin", "line.join",
                      "cmpd", "compoundline", "compoundLine", "line.compound",
                      "linealign", "lineAlign", "line.align",
                      "headend", "headEnd", "arrowstart", "arrowStart",
                      "tailend", "tailEnd", "arrowend", "arrowEnd",
                      "image", "imagefill",
                      // blip-fill framing — consumed alongside image= so the
                      // stretch insets / crop round-trip with the image fill.
                      "fillRect", "fillrect", "srcRect", "srcrect",
                      // CONSISTENCY(rpr-attr-fallback / R21-fuzzer-1+2): drawingML
                      // run-property attributes must reach SetRunOrShapeProperties
                      // so the long-tail rPr-attribute branch routes them to the
                      // first run instead of dropping them on the <p:sp> element.
                      "lang", "lang.latin", "altLang", "altlang", "spc", "kern", "cap",
                      "kumimoji", "normalizeH", "normalizeh", "noProof", "noproof",
                      "dirty", "smtClean", "smtclean", "smtId", "smtid", "err",
                      // BUG1: text direction — Set handles in SetRunOrShapeProperties
                      "textdirection", "textdir",
                      // BUG2: text outline — Set handles all three key variants
                      "textOutline", "textoutline", "textOutlineRaw", "textoutlineraw", "textFillRaw", "textfillraw",
                      "textOutline.width", "textoutline.width",
                      "textOutline.color", "textoutline.color",
                      // CONSISTENCY(highlight): a:highlight — Set's curated case
                      // in SetRunOrShapeProperties writes it; route Add through
                      // the same fan-out so both paths support it (root
                      // the project conventions Feature Implementation Checklist).
                      "highlight" };
                // CONSISTENCY(tracking-prop): explicit TryGetValue per known
                // key instead of `.Where(...)` iteration. Foreach over the
                // TrackingPropertyDictionary marks every entry as consumed
                // (see Core/TrackingPropertyDictionary.cs), which would
                // silently swallow user typos (xyzNeverExisted, anchor, …)
                // and make Add asymmetric with Set on the unsupported_property
                // contract. TryGetValue records only the keys we actually
                // looked up, leaving genuine typos visible to CommandBuilder.
                var effectProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                foreach (var ek in effectKeys)
                {
                    if (properties.TryGetValue(ek, out var ev))
                        effectProps[ek] = ev;
                }
                if (effectProps.Count > 0)
                    SetRunOrShapeProperties(effectProps, GetAllRuns(newShape), newShape, ownerPart);

                // Animation — slide-only (timing lives on <p:sld>/timing).
                if (properties.TryGetValue("animation", out var animVal) ||
                    properties.TryGetValue("animate", out animVal))
                {
                    if (slidePart == null)
                        throw new ArgumentException(
                            "'animation' is not supported on master/layout shapes — slide timing trees live on /slide[N].");
                    ApplyShapeAnimation(slidePart, newShape, animVal);
                }

                // Z-order — slide-only. NodeBuilder emits the 1-based position
                // among content elements; without consuming it here every Add
                // appended at the end of the shape tree and dump-replay lost
                // the original stacking order.
                if (slidePart != null && (
                    properties.TryGetValue("zorder", out var zVal)
                    || properties.TryGetValue("z-order", out zVal)
                    || properties.TryGetValue("order", out zVal)))
                {
                    ApplyZOrder(slidePart, newShape, zVal);
                }

                ownerRoot.Save();
                if (groupResultPathPrefix != null && insertContainer != null)
                {
                    // Positional within the group container: 1-based shape index
                    // among Shape children of the GroupShape, matching the format
                    // emitted by Get for /slide/group/shape paths.
                    var inGroupIdx = insertContainer.Elements<Shape>().Count();
                    return $"{groupResultPathPrefix}/{BuildElementPathSegment("shape", newShape, inGroupIdx)}";
                }
                return $"{returnPathPrefix}/{BuildElementPathSegment("shape", newShape, shapeTree.Elements<Shape>().Count())}";
    }


}
