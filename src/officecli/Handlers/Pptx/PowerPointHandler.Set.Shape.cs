// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Handlers;

// Per-element-type Set helpers for shape / paragraph / run / placeholder /
// group / connector paths. Mechanically extracted from the original god-method
// Set(); each helper owns one path-pattern's full handling. No behavior change.
public partial class PowerPointHandler
{
    private List<string> SetShapeRunByPath(Match runMatch, Dictionary<string, string> properties)
    {
        var slideIdx = int.Parse(runMatch.Groups[1].Value);
        var shapeIdx = int.Parse(runMatch.Groups[2].Value);
        var runIdx = int.Parse(runMatch.Groups[3].Value);

        var (slidePart, shape) = ResolveShape(slideIdx, shapeIdx);
        var allRuns = GetAllRuns(shape);
        if (runIdx < 1 || runIdx > allRuns.Count)
            throw new ArgumentException($"Run {runIdx} not found (shape has {allRuns.Count} runs)");

        var targetRun = allRuns[runIdx - 1];
        var linkValRun = properties.GetValueOrDefault("link");
        var tooltipValRun = properties.GetValueOrDefault("tooltip");
        var runOnlyProps = properties
            .Where(kv => !kv.Key.Equals("link", StringComparison.OrdinalIgnoreCase)
                      && !kv.Key.Equals("tooltip", StringComparison.OrdinalIgnoreCase))
            .ToDictionary(kv => kv.Key, kv => kv.Value);
        var unsupported = SetRunOrShapeProperties(runOnlyProps, new List<Drawing.Run> { targetRun }, shape, slidePart, runContext: true, unsupportedContextHint: RunPropsHint);
        if (linkValRun != null) ApplyRunHyperlink(slidePart, targetRun, linkValRun, tooltipValRun);
        GetSlide(slidePart).Save();
        return unsupported;
    }

    // Context labels used by SetRunOrShapeProperties so paragraph/run paths
    // report paragraph-/run-valid props instead of the broader shape list.
    private const string RunPropsHint =
        "valid run props: text, bold, italic, underline, strike, color, fill, size, font, font.latin, font.ea, font.cs, link, tooltip, baseline, spacing, cap";
    private const string ParagraphPropsHint =
        "valid paragraph props: align, indent, level, marginLeft, marginRight, lineSpacing, spaceBefore, spaceAfter, link, tooltip — plus any run prop (applied to all runs in the paragraph)";

    private List<string> SetParagraphRunByPath(Match paraRunMatch, Dictionary<string, string> properties)
    {
        var slideIdx = int.Parse(paraRunMatch.Groups[1].Value);
        var shapeIdx = int.Parse(paraRunMatch.Groups[2].Value);
        var paraIdx = int.Parse(paraRunMatch.Groups[3].Value);
        var runIdx = int.Parse(paraRunMatch.Groups[4].Value);

        var (slidePart, shape) = ResolveShape(slideIdx, shapeIdx);
        return SetParagraphRunOnShape(slidePart, shape, paraIdx, runIdx, properties);
    }


    // CONSISTENCY(placeholder-paragraph-path): /slide[N]/placeholder[X]/paragraph[K]
    // shares the same paragraph/run setter as the /shape[M]/paragraph[K] form.
    // ResolvePlaceholderShape materializes layout-inherited placeholders so
    // the slide-level <p:sp> exists before we navigate into its txBody.
    private List<string> SetPlaceholderParagraphByPath(Match phParaMatch, Dictionary<string, string> properties)
    {
        var slideIdx = int.Parse(phParaMatch.Groups[1].Value);
        var phId = phParaMatch.Groups[2].Value;
        var paraIdx = int.Parse(phParaMatch.Groups[3].Value);

        var slideParts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > slideParts.Count)
            throw new ArgumentException($"Slide {slideIdx} not found (total: {slideParts.Count})");
        var slidePart = slideParts[slideIdx - 1];
        var shape = ResolvePlaceholderShape(slidePart, phId);
        return SetParagraphOnShape(slidePart, shape, paraIdx, properties);
    }

    private List<string> SetPlaceholderParagraphRunByPath(Match phParaRunMatch, Dictionary<string, string> properties)
    {
        var slideIdx = int.Parse(phParaRunMatch.Groups[1].Value);
        var phId = phParaRunMatch.Groups[2].Value;
        var paraIdx = int.Parse(phParaRunMatch.Groups[3].Value);
        var runIdx = int.Parse(phParaRunMatch.Groups[4].Value);

        var slideParts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > slideParts.Count)
            throw new ArgumentException($"Slide {slideIdx} not found (total: {slideParts.Count})");
        var slidePart = slideParts[slideIdx - 1];
        var shape = ResolvePlaceholderShape(slidePart, phId);
        return SetParagraphRunOnShape(slidePart, shape, paraIdx, runIdx, properties);
    }


    private List<string> SetParagraphByPath(Match paraMatch, Dictionary<string, string> properties)
    {
        var slideIdx = int.Parse(paraMatch.Groups[1].Value);
        var shapeIdx = int.Parse(paraMatch.Groups[2].Value);
        var paraIdx = int.Parse(paraMatch.Groups[3].Value);

        var (slidePart, shape) = ResolveShape(slideIdx, shapeIdx);
        return SetParagraphOnShape(slidePart, shape, paraIdx, properties);
    }

    private List<string> SetParagraphRunOnShape(SlidePart slidePart, Shape shape, int paraIdx, int runIdx, Dictionary<string, string> properties)
    {
        var paragraphs = shape.TextBody?.Elements<Drawing.Paragraph>().ToList()
            ?? throw new ArgumentException("Shape has no text body");
        if (paraIdx < 1 || paraIdx > paragraphs.Count)
            throw new ArgumentException($"Paragraph {paraIdx} not found (shape has {paragraphs.Count} paragraphs)");
        var para = paragraphs[paraIdx - 1];
        var paraRuns = para.Elements<Drawing.Run>().ToList();
        if (runIdx < 1 || runIdx > paraRuns.Count)
            throw new ArgumentException($"Run {runIdx} not found (paragraph has {paraRuns.Count} runs)");

        var targetRun = paraRuns[runIdx - 1];
        var linkVal = properties.GetValueOrDefault("link");
        var tooltipVal = properties.GetValueOrDefault("tooltip");
        var runOnlyProps = properties
            .Where(kv => !kv.Key.Equals("link", StringComparison.OrdinalIgnoreCase)
                      && !kv.Key.Equals("tooltip", StringComparison.OrdinalIgnoreCase))
            .ToDictionary(kv => kv.Key, kv => kv.Value);
        var unsupported = SetRunOrShapeProperties(runOnlyProps, new List<Drawing.Run> { targetRun }, shape, slidePart, runContext: true, unsupportedContextHint: RunPropsHint);
        if (linkVal != null) ApplyRunHyperlink(slidePart, targetRun, linkVal, tooltipVal);
        GetSlide(slidePart).Save();
        return unsupported;
    }

    private List<string> SetParagraphOnShape(SlidePart slidePart, Shape shape, int paraIdx, Dictionary<string, string> properties)
    {
        var paragraphs = shape.TextBody?.Elements<Drawing.Paragraph>().ToList()
            ?? throw new ArgumentException("Shape has no text body");
        if (paraIdx < 1 || paraIdx > paragraphs.Count)
            throw new ArgumentException($"Paragraph {paraIdx} not found (shape has {paragraphs.Count} paragraphs)");

        var para = paragraphs[paraIdx - 1];
        var paraRuns = para.Elements<Drawing.Run>().ToList();
        var unsupported = new List<string>();

        // Order keys so `text` is processed BEFORE run-style props (size /
        // color / font.* / bold / italic / ...). The text branch in
        // SetRunOrShapeProperties rebuilds the shape's paragraphs from
        // scratch whenever the original run count was 0 (empty placeholder)
        // or the new text contains newlines / tabs — every Drawing.Run
        // already mutated by an earlier key in the iteration order is then
        // detached from the tree, and the styling silently disappears. The
        // post-text refresh below repairs the case where text comes first or
        // is interleaved with later keys; reordering up-front guarantees the
        // common dump→replay pattern ("set paragraph text=X, size=48pt,
        // color=#FFFFFF") lands the styling on the new runs.
        var orderedKeys = properties.Keys
            .OrderBy(k => k.Equals("text", StringComparison.OrdinalIgnoreCase) ? 0 : 1)
            .ToList();

        foreach (var key in orderedKeys)
        {
            var value = properties[key];
            switch (key.ToLowerInvariant())
            {
                // Schema declares aliases: [alignment, halign] for paragraph.align.
                // CONSISTENCY(canonical-keys): accept the documented aliases here so
                // they don't drop through to SetRunOrShapeProperties (which would
                // surface them as UNSUPPORTED, since shape's `align` is text body
                // alignment with a different code path).
                case "align" or "alignment" or "halign":
                {
                    var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                    pProps.Alignment = ParseTextAlignment(value);
                    break;
                }
                case "indent":
                {
                    var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                    pProps.Indent = (int)ParseEmu(value);
                    break;
                }
                case "level":
                {
                    var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                    if (!int.TryParse(value, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out var lvl) || lvl < 0 || lvl > 8)
                        throw new ArgumentException($"Invalid 'level' value: '{value}'. Expected an integer between 0 and 8 (OOXML a:pPr/@lvl).");
                    pProps.Level = lvl;
                    break;
                }
                case "marginleft" or "marl":
                {
                    var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                    pProps.LeftMargin = (int)ParseEmu(value);
                    break;
                }
                case "marginright" or "marr":
                {
                    var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                    pProps.RightMargin = (int)ParseEmu(value);
                    break;
                }
                case "linespacing" or "line.spacing":
                {
                    var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                    pProps.RemoveAllChildren<Drawing.LineSpacing>();
                    var (lsVal2, lsIsPercent) = SpacingConverter.ParsePptLineSpacing(value);
                    var lnSpc = lsIsPercent
                        ? new Drawing.LineSpacing(new Drawing.SpacingPercent { Val = lsVal2 })
                        : new Drawing.LineSpacing(new Drawing.SpacingPoints { Val = lsVal2 });
                    // CONSISTENCY(schema-order-pptx): pPr children must follow
                    // CT_TextParagraphProperties order or PowerPoint silently
                    // drops them. See PowerPointHandler.Helpers.cs.
                    InsertPPrChild(pProps, lnSpc);
                    break;
                }
                case "spacebefore" or "space.before":
                {
                    var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                    pProps.RemoveAllChildren<Drawing.SpaceBefore>();
                    InsertPPrChild(pProps, new Drawing.SpaceBefore(new Drawing.SpacingPoints { Val = SpacingConverter.ParsePptSpacing(value) }));
                    break;
                }
                case "spaceafter" or "space.after":
                {
                    var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                    pProps.RemoveAllChildren<Drawing.SpaceAfter>();
                    InsertPPrChild(pProps, new Drawing.SpaceAfter(new Drawing.SpacingPoints { Val = SpacingConverter.ParsePptSpacing(value) }));
                    break;
                }
                case "link":
                {
                    var paraTooltip = properties.GetValueOrDefault("tooltip");
                    foreach (var r in paraRuns)
                        ApplyRunHyperlink(slidePart, r, value, paraTooltip);
                    break;
                }
                case "tooltip":
                    // handled in tandem with "link"; standalone tooltip change is not supported here
                    break;
                default:
                    // Apply run-level properties to all runs in this paragraph
                    var runUnsup = SetRunOrShapeProperties(
                        new Dictionary<string, string> { { key, value } }, paraRuns, shape, slidePart, runContext: true,
                        unsupportedContextHint: ParagraphPropsHint);
                    unsupported.AddRange(runUnsup);
                    // The `text` case in SetRunOrShapeProperties rebuilds the
                    // shape's paragraphs from scratch when the original run
                    // count was 0 (empty placeholder) or when the new text
                    // spans multiple lines / contains tabs — in either case
                    // every Drawing.Run captured in paraRuns is detached from
                    // the tree and any subsequent property write lands on
                    // orphaned XML. Refresh paraRuns against the live shape
                    // so the very next key (size / color / font.latin / ...)
                    // hits the new run instances. Re-resolve via paraIdx so
                    // dump→replay of an empty title placeholder ("set para
                    // text=X, size=48pt, color=#FFF") keeps the rPr on the
                    // run instead of dropping every styling key after text.
                    if (key.Equals("text", StringComparison.OrdinalIgnoreCase))
                    {
                        var refreshed = shape.TextBody?.Elements<Drawing.Paragraph>().ToList()
                            ?? new List<Drawing.Paragraph>();
                        if (paraIdx >= 1 && paraIdx <= refreshed.Count)
                        {
                            para = refreshed[paraIdx - 1];
                            paraRuns = para.Elements<Drawing.Run>().ToList();
                        }
                    }
                    break;
            }
        }

        GetSlide(slidePart).Save();
        return unsupported;
    }



    private List<string> SetPlaceholderByPath(Match phMatch, Dictionary<string, string> properties)
    {
        var slideIdx = int.Parse(phMatch.Groups[1].Value);
        var phId = phMatch.Groups[2].Value;

        var slideParts2 = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > slideParts2.Count)
            throw new ArgumentException($"Slide {slideIdx} not found (total: {slideParts2.Count})");
        var slidePart = slideParts2[slideIdx - 1];
        var shape = ResolvePlaceholderShape(slidePart, phId);

        var allRuns = shape.Descendants<Drawing.Run>().ToList();
        var unsupported = SetRunOrShapeProperties(properties, allRuns, shape, slidePart);
        GetSlide(slidePart).Save();
        return unsupported;
    }

    private List<string> SetGroupByPath(Match grpMatch, Dictionary<string, string> properties)
    {
        var slideIdx = int.Parse(grpMatch.Groups[1].Value);
        var grpIdx = int.Parse(grpMatch.Groups[2].Value);

        var slideParts6 = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > slideParts6.Count)
            throw new ArgumentException($"Slide {slideIdx} not found (total: {slideParts6.Count})");

        var slidePart = slideParts6[slideIdx - 1];
        var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree
            ?? throw new ArgumentException("Slide has no shape tree");
        var groups = shapeTree.Elements<GroupShape>().ToList();
        if (grpIdx < 1 || grpIdx > groups.Count)
            throw new ArgumentException($"Group {grpIdx} not found (total: {groups.Count})");

        var grp = groups[grpIdx - 1];
        // Pull link/tooltip up front so the tooltip is applied alongside link
        // even when only one of them is also in properties — same pairing as
        // ApplyShapeHyperlink at shape level.
        var grpLinkValue = properties.GetValueOrDefault("link");
        var grpTooltipValue = properties.GetValueOrDefault("tooltip");
        var unsupported = new List<string>();
        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "name":
                    var nvGrpPr = grp.NonVisualGroupShapeProperties?.NonVisualDrawingProperties;
                    if (nvGrpPr != null)
                    {
                        Core.XmlTextValidator.ValidateOrThrow(value, "name");
                        nvGrpPr.Name = value;
                    }
                    break;
                case "link":
                    ApplyGroupHyperlink(slidePart, grp, value, grpTooltipValue);
                    break;
                case "tooltip":
                    // Paired with "link" above. When the user sets tooltip
                    // without link on a group that already has a hyperlink,
                    // update only the tooltip attribute in place.
                    if (grpLinkValue == null)
                    {
                        var existing = grp.NonVisualGroupShapeProperties?.NonVisualDrawingProperties
                            ?.GetFirstChild<Drawing.HyperlinkOnClick>();
                        if (existing != null)
                        {
                            Core.XmlTextValidator.ValidateOrThrow(value, "tooltip");
                            existing.Tooltip = value;
                        }
                    }
                    break;
                case "x" or "y" or "width" or "height":
                {
                    var grpSpPr = grp.GroupShapeProperties ?? (grp.GroupShapeProperties = new GroupShapeProperties());
                    var xfrm = grpSpPr.TransformGroup ?? (grpSpPr.TransformGroup = new Drawing.TransformGroup());
                    var off = xfrm.Offset ?? (xfrm.Offset = new Drawing.Offset());
                    var ext = xfrm.Extents ?? (xfrm.Extents = new Drawing.Extents());
                    var keyLower = key.ToLowerInvariant();
                    // CONSISTENCY(group-scale-baseline): group scaling needs <a:chOff>/<a:chExt>
                    // as a child-coordinate baseline. Before we mutate ext/off, snapshot the
                    // current ext/off into chExt/chOff if they aren't already present — that
                    // way the first Set of width/height captures the "before" as the logical
                    // child coordinate space, so shrinking ext shrinks the rendered children.
                    if (keyLower is "x" or "y")
                    {
                        if (xfrm.ChildOffset == null)
                            xfrm.ChildOffset = new Drawing.ChildOffset { X = off.X ?? 0, Y = off.Y ?? 0 };
                    }
                    else // width or height
                    {
                        if (xfrm.ChildExtents == null)
                            xfrm.ChildExtents = new Drawing.ChildExtents { Cx = ext.Cx ?? 0, Cy = ext.Cy ?? 0 };
                    }
                    TryApplyPositionSize(keyLower, value, off, ext);
                    break;
                }
                case "rotation" or "rotate":
                {
                    var grpSpPr = grp.GroupShapeProperties ?? (grp.GroupShapeProperties = new GroupShapeProperties());
                    var xfrm = grpSpPr.TransformGroup ?? (grpSpPr.TransformGroup = new Drawing.TransformGroup());
                    xfrm.Rotation = (int)(ParseHelpers.SafeParseRotationDegrees(value, "rotation") * 60000);
                    break;
                }
                case "fill":
                {
                    var grpSpPr = grp.GroupShapeProperties ?? (grp.GroupShapeProperties = new GroupShapeProperties());
                    grpSpPr.RemoveAllChildren<Drawing.SolidFill>();
                    grpSpPr.RemoveAllChildren<Drawing.NoFill>();
                    grpSpPr.RemoveAllChildren<Drawing.GradientFill>();
                    if (value.Equals("none", StringComparison.OrdinalIgnoreCase))
                        grpSpPr.AppendChild(new Drawing.NoFill());
                    else
                        grpSpPr.AppendChild(BuildSolidFill(value));
                    break;
                }
                default:
                    if (!GenericXmlQuery.SetGenericAttribute(grp, key, value))
                    {
                        if (unsupported.Count == 0)
                            unsupported.Add($"{key} (valid group props: x, y, width, height, rotation, name, fill, link, tooltip)");
                        else
                            unsupported.Add(key);
                    }
                    break;
            }
        }
        GetSlide(slidePart).Save();
        return unsupported;
    }

    private List<string> SetConnectorByPath(Match cxnMatch, Dictionary<string, string> properties)
    {
        var slideIdx = int.Parse(cxnMatch.Groups[1].Value);
        var cxnIdx = int.Parse(cxnMatch.Groups[2].Value);

        var slideParts5 = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > slideParts5.Count)
            throw new ArgumentException($"Slide {slideIdx} not found (total: {slideParts5.Count})");

        var slidePart = slideParts5[slideIdx - 1];
        var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree
            ?? throw new ArgumentException("Slide has no shape tree");
        var connectors = shapeTree.Elements<ConnectionShape>().ToList();
        if (cxnIdx < 1 || cxnIdx > connectors.Count)
            throw new ArgumentException($"Connector {cxnIdx} not found (total: {connectors.Count})");

        var cxn = connectors[cxnIdx - 1];
        var unsupported = new List<string>();
        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "name":
                    var nvCxnPr = cxn.NonVisualConnectionShapeProperties?.NonVisualDrawingProperties;
                    if (nvCxnPr != null)
                    {
                        Core.XmlTextValidator.ValidateOrThrow(value, "name");
                        nvCxnPr.Name = value;
                    }
                    break;
                case "x" or "y" or "width" or "height":
                {
                    var spPr = cxn.ShapeProperties ?? (cxn.ShapeProperties = new ShapeProperties());
                    var xfrm = spPr.Transform2D ?? (spPr.Transform2D = new Drawing.Transform2D());
                    TryApplyPositionSize(key.ToLowerInvariant(), value,
                        xfrm.Offset ?? (xfrm.Offset = new Drawing.Offset()),
                        xfrm.Extents ?? (xfrm.Extents = new Drawing.Extents()));
                    break;
                }
                case "linewidth" or "line.width":
                {
                    var spPr = cxn.ShapeProperties ?? (cxn.ShapeProperties = new ShapeProperties());
                    var outline = EnsureOutline(spPr);
                    outline.Width = Core.EmuConverter.ParseLineWidth(value);
                    break;
                }
                case "linecolor" or "line.color" or "line" or "color":
                {
                    // Schema documents compound 'color[:width[:style]]'
                    // for shape line=; mirror the same surface on connector
                    // so the documented form works uniformly.
                    var (lineColorPart, lineWidthPart, lineDashPart) = SplitCompoundLineValue(value);
                    var spPr = cxn.ShapeProperties ?? (cxn.ShapeProperties = new ShapeProperties());
                    var outline = EnsureOutline(spPr);
                    outline.RemoveAllChildren<Drawing.SolidFill>();
                    if (lineWidthPart != null)
                        outline.Width = Core.EmuConverter.ParseLineWidth(lineWidthPart);
                    if (lineDashPart != null)
                    {
                        outline.RemoveAllChildren<Drawing.PresetDash>();
                        outline.AppendChild(new Drawing.PresetDash { Val = ParseLineDashValue(lineDashPart) });
                    }
                    // CONSISTENCY(color-input-scheme): shape line= already routes
                    // through BuildSolidFill which accepts scheme names (accent1,
                    // dark1, …) and hex equally; mirror the same surface here so
                    // a connector accepts the documented vocabulary instead of
                    // rejecting scheme colors at SanitizeColorForOoxml.
                    var newFill = BuildSolidFill(lineColorPart);
                    // CT_LineProperties schema: fill → prstDash → ... → headEnd → tailEnd
                    var prstDash = outline.GetFirstChild<Drawing.PresetDash>();
                    if (prstDash != null)
                        outline.InsertBefore(newFill, prstDash);
                    else
                    {
                        var headEnd = outline.GetFirstChild<Drawing.HeadEnd>();
                        if (headEnd != null)
                            outline.InsertBefore(newFill, headEnd);
                        else
                        {
                            var tailEnd = outline.GetFirstChild<Drawing.TailEnd>();
                            if (tailEnd != null)
                                outline.InsertBefore(newFill, tailEnd);
                            else
                                outline.AppendChild(newFill);
                        }
                    }
                    break;
                }
                case "fill":
                {
                    var spPr = cxn.ShapeProperties ?? (cxn.ShapeProperties = new ShapeProperties());
                    ApplyShapeFill(spPr, value);
                    break;
                }
                case "line.gradient" or "linegradient":
                {
                    var spPr = cxn.ShapeProperties ?? (cxn.ShapeProperties = new ShapeProperties());
                    var outline = EnsureOutline(spPr);
                    outline.RemoveAllChildren<Drawing.SolidFill>();
                    outline.RemoveAllChildren<Drawing.NoFill>();
                    outline.RemoveAllChildren<Drawing.GradientFill>();
                    var cxnGrad = BuildGradientFill(value);
                    var cxnPrstDash = outline.GetFirstChild<Drawing.PresetDash>();
                    if (cxnPrstDash != null)
                        outline.InsertBefore(cxnGrad, cxnPrstDash);
                    else
                        outline.PrependChild(cxnGrad);
                    break;
                }
                case "linedash" or "line.dash":
                {
                    var spPr = cxn.ShapeProperties ?? (cxn.ShapeProperties = new ShapeProperties());
                    var outline = EnsureOutline(spPr);
                    outline.RemoveAllChildren<Drawing.PresetDash>();
                    var newDash = new Drawing.PresetDash { Val = ParseLineDashValue(value) };
                    // CT_LineProperties schema: fill → prstDash → ... → headEnd → tailEnd
                    var headEnd = outline.GetFirstChild<Drawing.HeadEnd>();
                    if (headEnd != null)
                        outline.InsertBefore(newDash, headEnd);
                    else
                    {
                        var tailEnd = outline.GetFirstChild<Drawing.TailEnd>();
                        if (tailEnd != null)
                            outline.InsertBefore(newDash, tailEnd);
                        else
                            outline.AppendChild(newDash);
                    }
                    break;
                }
                case "lineopacity" or "line.opacity":
                {
                    var spPr = cxn.ShapeProperties ?? (cxn.ShapeProperties = new ShapeProperties());
                    if (!double.TryParse(value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var lnOpacity)
                        || double.IsNaN(lnOpacity) || double.IsInfinity(lnOpacity))
                        throw new ArgumentException($"Invalid 'lineOpacity' value: '{value}'. Expected a finite decimal 0.0-1.0.");
                    var outline = EnsureOutline(spPr);
                    var solidFill = outline.GetFirstChild<Drawing.SolidFill>();
                    if (solidFill == null)
                    {
                        // Auto-create a black line fill (matching Apache POI behavior)
                        // CT_LineProperties schema: fill → prstDash → ... → headEnd → tailEnd
                        solidFill = new Drawing.SolidFill(new Drawing.RgbColorModelHex { Val = "000000" });
                        var prstDashEl = outline.GetFirstChild<Drawing.PresetDash>();
                        if (prstDashEl != null)
                            outline.InsertBefore(solidFill, prstDashEl);
                        else
                        {
                            var headEndEl = outline.GetFirstChild<Drawing.HeadEnd>();
                            if (headEndEl != null)
                                outline.InsertBefore(solidFill, headEndEl);
                            else
                            {
                                var tailEndEl = outline.GetFirstChild<Drawing.TailEnd>();
                                if (tailEndEl != null)
                                    outline.InsertBefore(solidFill, tailEndEl);
                                else
                                    outline.AppendChild(solidFill);
                            }
                        }
                    }
                    {
                        var colorEl = solidFill.GetFirstChild<Drawing.RgbColorModelHex>() as OpenXmlElement
                            ?? solidFill.GetFirstChild<Drawing.SchemeColor>();
                        if (colorEl != null)
                        {
                            colorEl.RemoveAllChildren<Drawing.Alpha>();
                            colorEl.AppendChild(new Drawing.Alpha { Val = (int)(lnOpacity * 100000) });
                        }
                    }
                    break;
                }
                case "rotation" or "rotate":
                {
                    var spPr = cxn.ShapeProperties ?? (cxn.ShapeProperties = new ShapeProperties());
                    var xfrm = spPr.Transform2D ?? (spPr.Transform2D = new Drawing.Transform2D());
                    xfrm.Rotation = (int)(ParseHelpers.SafeParseRotationDegrees(value, "rotation") * 60000);
                    break;
                }
                case "preset" or "prstgeom" or "shape":
                {
                    // CONSISTENCY(canonical-key): schema canonical is 'shape';
                    // 'preset'/'prstgeom' retained as legacy aliases.
                    var spPr = cxn.ShapeProperties ?? (cxn.ShapeProperties = new ShapeProperties());
                    var prstGeom = EnsurePresetGeometry(spPr);
                    // CONSISTENCY(connector-shape-aliases): mirror Add.Misc.cs —
                    // accept short canonical names (straight/elbow/curve) plus
                    // OOXML full names (incl. 2-segment forms which fold to 3-segment).
                    var resolvedShape = value.ToLowerInvariant() switch
                    {
                        "straight" or "straightconnector1" or "line" => Drawing.ShapeTypeValues.StraightConnector1,
                        "elbow" or "bentconnector3" or "bentconnector2" => Drawing.ShapeTypeValues.BentConnector3,
                        "curve" or "curvedconnector3" or "curvedconnector2" => Drawing.ShapeTypeValues.CurvedConnector3,
                        _ => new Drawing.ShapeTypeValues(value),
                    };
                    prstGeom.Preset = resolvedShape;
                    break;
                }
                case "headend" or "headEnd":
                {
                    var spPr = cxn.ShapeProperties ?? (cxn.ShapeProperties = new ShapeProperties());
                    var outline = EnsureOutline(spPr);
                    outline.RemoveAllChildren<Drawing.HeadEnd>();
                    var newHeadEnd = new Drawing.HeadEnd { Type = ParseLineEndType(value) };
                    // CT_LineProperties: ... → headEnd → tailEnd (headEnd before tailEnd)
                    var existingTailEnd = outline.GetFirstChild<Drawing.TailEnd>();
                    if (existingTailEnd != null)
                        outline.InsertBefore(newHeadEnd, existingTailEnd);
                    else
                        outline.AppendChild(newHeadEnd);
                    break;
                }
                case "tailend" or "tailEnd":
                {
                    var spPr = cxn.ShapeProperties ?? (cxn.ShapeProperties = new ShapeProperties());
                    var outline = EnsureOutline(spPr);
                    outline.RemoveAllChildren<Drawing.TailEnd>();
                    // CT_LineProperties: tailEnd is last — always append
                    outline.AppendChild(new Drawing.TailEnd { Type = ParseLineEndType(value) });
                    break;
                }
                case "from" or "startshape":
                case "to" or "endshape":
                {
                    // CONSISTENCY(connector-endpoints): mirror Add.Misc.cs's
                    // from/to wiring. Schema declares set:true for from/to;
                    // previously the Set path had no case so updates were
                    // rejected as unsupported_property. Replace any existing
                    // StartConnection/EndConnection rather than append (XML
                    // schema allows only one of each on a connector).
                    var endpointId = ResolveShapeId(value, shapeTree);
                    var cxnDrawProps = cxn.NonVisualConnectionShapeProperties
                        ?.GetFirstChild<NonVisualConnectorShapeDrawingProperties>();
                    if (cxnDrawProps == null) { unsupported.Add(key); break; }
                    bool isStart = key.Equals("from", StringComparison.OrdinalIgnoreCase)
                        || key.Equals("startshape", StringComparison.OrdinalIgnoreCase);
                    if (isStart)
                    {
                        cxnDrawProps.RemoveAllChildren<Drawing.StartConnection>();
                        cxnDrawProps.AppendChild(new Drawing.StartConnection { Id = endpointId, Index = 0 });
                    }
                    else
                    {
                        cxnDrawProps.RemoveAllChildren<Drawing.EndConnection>();
                        cxnDrawProps.AppendChild(new Drawing.EndConnection { Id = endpointId, Index = 0 });
                    }
                    break;
                }
                default:
                    if (!GenericXmlQuery.SetGenericAttribute(cxn, key, value))
                    {
                        if (unsupported.Count == 0)
                            unsupported.Add($"{key} (valid connector props: line, color, fill, x, y, width, height, rotation, name, headEnd, tailEnd, geometry, from, to)");
                        else
                            unsupported.Add(key);
                    }
                    break;
            }
        }
        GetSlide(slidePart).Save();
        return unsupported;
    }

    private List<string> SetShapeByPath(Match match, Dictionary<string, string> properties)
    {
        var slideIdx = int.Parse(match.Groups[1].Value);
        var shapeIdx = int.Parse(match.Groups[2].Value);

        var (slidePart, shape) = ResolveShape(slideIdx, shapeIdx);
        return ApplyShapePropsCore(slidePart, shape, properties);
    }

    /// <summary>
    /// Resolve a shape nested inside a group: /slide[N]/group[M]/shape[K].
    /// CONSISTENCY(group-inner-shape): Get already supports this path via the
    /// generic XML fallback; Set previously had no dispatch entry, leading to
    /// "Element not found" even though Get could read the same path.
    /// </summary>
    private List<string> SetGroupInnerShapeByPath(Match match, Dictionary<string, string> properties)
    {
        var slideIdx = int.Parse(match.Groups[1].Value);
        var grpIdx = int.Parse(match.Groups[2].Value);
        var shapeIdx = int.Parse(match.Groups[3].Value);

        var slideParts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > slideParts.Count)
            throw new ArgumentException($"Slide {slideIdx} not found (total: {slideParts.Count})");
        var slidePart = slideParts[slideIdx - 1];
        var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree
            ?? throw new ArgumentException("Slide has no shape tree");
        var groups = shapeTree.Elements<GroupShape>().ToList();
        if (grpIdx < 1 || grpIdx > groups.Count)
            throw new ArgumentException($"Group {grpIdx} not found (total: {groups.Count})");
        var grp = groups[grpIdx - 1];
        var innerShapes = grp.Elements<Shape>().ToList();
        if (shapeIdx < 1 || shapeIdx > innerShapes.Count)
            throw new ArgumentException($"Shape {shapeIdx} not found in group {grpIdx} (total: {innerShapes.Count})");
        return ApplyShapePropsCore(slidePart, innerShapes[shapeIdx - 1], properties);
    }

    /// <summary>
    /// Resolve a Shape nested inside a group on a slide and return it
    /// alongside the owning SlidePart. Used by group-paragraph / group-run
    /// setters so the dispatch tier can pass control to the existing
    /// SetParagraphOnShape / SetParagraphRunOnShape helpers without
    /// duplicating navigation logic.
    /// </summary>
    private (SlidePart slidePart, Shape shape) ResolveGroupInnerShape(int slideIdx, int grpIdx, int shapeIdx)
    {
        var slideParts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > slideParts.Count)
            throw new ArgumentException($"Slide {slideIdx} not found (total: {slideParts.Count})");
        var slidePart = slideParts[slideIdx - 1];
        var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree
            ?? throw new ArgumentException("Slide has no shape tree");
        var groups = shapeTree.Elements<GroupShape>().ToList();
        if (grpIdx < 1 || grpIdx > groups.Count)
            throw new ArgumentException($"Group {grpIdx} not found (total: {groups.Count})");
        var grp = groups[grpIdx - 1];
        var innerShapes = grp.Elements<Shape>().ToList();
        if (shapeIdx < 1 || shapeIdx > innerShapes.Count)
            throw new ArgumentException($"Shape {shapeIdx} not found in group {grpIdx} (total: {innerShapes.Count})");
        return (slidePart, innerShapes[shapeIdx - 1]);
    }

    /// <summary>
    /// /slide[N]/group[M]/shape[K]/paragraph[P] — mirrors SetParagraphByPath
    /// but navigates into a group first.
    /// </summary>
    private List<string> SetGroupParagraphByPath(Match m, Dictionary<string, string> properties)
    {
        var slideIdx = int.Parse(m.Groups[1].Value);
        var grpIdx = int.Parse(m.Groups[2].Value);
        var shapeIdx = int.Parse(m.Groups[3].Value);
        var paraIdx = int.Parse(m.Groups[4].Value);
        var (slidePart, shape) = ResolveGroupInnerShape(slideIdx, grpIdx, shapeIdx);
        return SetParagraphOnShape(slidePart, shape, paraIdx, properties);
    }

    /// <summary>
    /// /slide[N]/group[M]/shape[K]/paragraph[P]/run[R] — mirrors
    /// SetParagraphRunByPath but navigates into a group first.
    /// </summary>
    private List<string> SetGroupParagraphRunByPath(Match m, Dictionary<string, string> properties)
    {
        var slideIdx = int.Parse(m.Groups[1].Value);
        var grpIdx = int.Parse(m.Groups[2].Value);
        var shapeIdx = int.Parse(m.Groups[3].Value);
        var paraIdx = int.Parse(m.Groups[4].Value);
        var runIdx = int.Parse(m.Groups[5].Value);
        var (slidePart, shape) = ResolveGroupInnerShape(slideIdx, grpIdx, shapeIdx);
        return SetParagraphRunOnShape(slidePart, shape, paraIdx, runIdx, properties);
    }

    private List<string> ApplyShapePropsCore(SlidePart slidePart, Shape shape, Dictionary<string, string> properties)
    {
        // Handle z-order first (changes shape position in tree)
        var zOrderValue = properties.GetValueOrDefault("zorder")
            ?? properties.GetValueOrDefault("z-order")
            ?? properties.GetValueOrDefault("order");
        if (zOrderValue != null)
        {
            ApplyZOrder(slidePart, shape, zOrderValue);
        }

        // Clone shape for rollback on failure (atomic: no partial modifications)
        var shapeBackup = shape.CloneNode(true);

        try
        {
            var allRuns = shape.Descendants<Drawing.Run>().ToList();

            // Separate animation, motionPath, link, and z-order from other shape properties
            var animValue = properties.GetValueOrDefault("animation")
                ?? properties.GetValueOrDefault("animate");
            var motionPathValue = properties.GetValueOrDefault("motionpath")
                ?? properties.GetValueOrDefault("motionPath");
            var linkValue = properties.GetValueOrDefault("link");
            var tooltipValue = properties.GetValueOrDefault("tooltip");
            var excludeKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                { "animation", "animate", "motionpath", "motionPath", "link", "tooltip", "zorder", "z-order", "order" };
            var shapeProps = properties
                .Where(kv => !excludeKeys.Contains(kv.Key))
                .ToDictionary(kv => kv.Key, kv => kv.Value);

            var unsupported = SetRunOrShapeProperties(shapeProps, allRuns, shape, slidePart);

            if (animValue != null)
            {
                // Remove existing animations before applying new one (replace, not accumulate)
                var shapeId = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value;
                if (shapeId.HasValue)
                    RemoveShapeAnimations(slidePart.Slide!, shapeId.Value);
                ApplyShapeAnimation(slidePart, shape, animValue);
            }
            if (motionPathValue != null)
                ApplyMotionPathAnimation(slidePart, shape, motionPathValue);
            if (linkValue != null)
                ApplyShapeHyperlink(slidePart, shape, linkValue, tooltipValue);
            else if (tooltipValue != null)
            {
                // Standalone tooltip update — apply in place to the existing
                // hlinkClick on shape and all runs. Previously this was a silent
                // no-op: set returned "Updated" but the tooltip slot was untouched.
                // If no hyperlink exists, reject so callers don't believe a
                // tooltip without a link was stored.
                Core.XmlTextValidator.ValidateOrThrow(tooltipValue, "tooltip");
                var shapeHl = shape.NonVisualShapeProperties?.NonVisualDrawingProperties
                    ?.GetFirstChild<Drawing.HyperlinkOnClick>();
                var runHls = shape.Descendants<Drawing.Run>()
                    .Select(r => r.RunProperties?.GetFirstChild<Drawing.HyperlinkOnClick>())
                    .Where(h => h != null)
                    .ToList();
                if (shapeHl == null && runHls.Count == 0)
                    throw new ArgumentException(
                        "tooltip requires an existing hyperlink — set 'link' in the same call (e.g. --prop link=https://example.com --prop tooltip=…) " +
                        "or apply 'link' first, then update 'tooltip' on its own.");
                if (shapeHl != null) shapeHl.Tooltip = tooltipValue;
                foreach (var rh in runHls) rh!.Tooltip = tooltipValue;
            }

            GetSlide(slidePart).Save();
            return unsupported;
        }
        catch
        {
            // Rollback: restore shape to pre-modification state
            shape.Parent?.ReplaceChild(shapeBackup, shape);
            throw;
        }
    }

    private List<string> SetShapeAnimationByPath(Match match, Dictionary<string, string> properties)
    {
        var slideIdx = int.Parse(match.Groups[1].Value);
        var shapeIdx = int.Parse(match.Groups[2].Value);
        var animIdx = int.Parse(match.Groups[3].Value);

        var (slidePart, shape) = ResolveShape(slideIdx, shapeIdx);
        var ctns = EnumerateShapeAnimationCTns(slidePart, shape);
        if (animIdx < 1 || animIdx > ctns.Count)
            throw new ArgumentException(
                $"Animation {animIdx} not found on shape {shapeIdx} (total: {ctns.Count})");

        // Reject schema set:false keys up front. Without this, the merge
        // loop silently dropped them and Set returned success with the
        // value unchanged. Schema: presetId / easein / easeout / motionPath
        // are read-only (Get-only).
        var readOnlyAnimKeys = new[] { "presetId", "presetid", "easein", "easeout", "motionPath", "motionpath" };
        foreach (var k in readOnlyAnimKeys)
        {
            if (properties.ContainsKey(k))
                throw new ArgumentException(
                    $"Animation property '{k}' is read-only (Get-only per schema). " +
                    "It is derived from the effect preset and cannot be set directly.");
        }

        // L3 sub-A: chain-preserving Set. Snapshot every existing animation on
        // the shape into a (props) dict via PopulateAnimationNode, mutate the
        // K-th dict with the caller's overrides, then rebuild the whole chain
        // in original order. Previously this method removed ALL animations and
        // re-added one, destroying the chain on any indexed Set call.
        // CONSISTENCY(animation-chain): rebuild model also used implicitly by
        // /animation[K] Remove (RemoveSingleShapeAnimation), so Add/Get/Set/Remove
        // share one indexing contract.
        var snapshots = new List<Dictionary<string, string>>(ctns.Count);
        for (int i = 0; i < ctns.Count; i++)
        {
            var n = new DocumentNode { Path = "" };
            PopulateAnimationNode(n, ctns[i]);
            var d = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var (k, v) in n.Format)
            {
                if (v == null) continue;
                var s = v.ToString() ?? "";
                if (s.Length == 0) continue;
                // presetId is derived, not a re-applicable input.
                if (k.Equals("presetId", StringComparison.OrdinalIgnoreCase)) continue;
                d[k] = s;
            }
            snapshots.Add(d);
        }

        // Merge caller overrides onto the target index. CONSISTENCY(animation-
        // class-suffix): if the user overrides `effect` with a suffixed form
        // (fly-out, fade-exit, …) and did not also pass an explicit `class`,
        // drop the snapshot's class so the suffix's class wins. Mirrors
        // AddAnimation's behaviour where suffixCls overrides the entrance
        // default unless an explicit class= was supplied.
        var target = snapshots[animIdx - 1];
        var userOverridesEffect = properties.ContainsKey("effect");
        var userPassesClass = properties.ContainsKey("class");
        foreach (var (k, v) in properties)
        {
            target[k] = v;
        }
        if (userOverridesEffect && !userPassesClass)
        {
            var (_, suffixCls) = ParseEffectClassSuffix(properties["effect"]);
            if (suffixCls != null) target["class"] = suffixCls;
        }

        // Validate the target snapshot. Unsupplied snapshot values were
        // already validated on the original Add path, but the caller's
        // overrides may be junk — re-validate everything that touches the
        // schema-typed slots so the surface is symmetric with AddAnimation.
        if (target.TryGetValue("class", out var tCls)) ValidateAnimationClass(tCls);
        if (target.TryGetValue("duration", out var tDur)) ValidateAnimationDuration(tDur);
        if (target.TryGetValue("dur", out var tDur2)) ValidateAnimationDuration(tDur2);
        if (target.TryGetValue("delay", out var tDel)) ValidateAnimationDelay(tDel);
        if (target.TryGetValue("repeat", out var tRep)) ValidateAnimationRepeat(tRep);
        if (target.TryGetValue("restart", out var tRes)) ValidateAnimationRestart(tRes);
        if (target.TryGetValue("autoReverse", out var tAr)) ValidateAnimationAutoReverse(tAr);
        else if (target.TryGetValue("autoreverse", out tAr)) ValidateAnimationAutoReverse(tAr);

        // Wipe all animations on the shape, then re-apply each snapshot in order.
        // CONSISTENCY(animation-chain): motion-class snapshots route through
        // AppendMotionPathAnimation; preset (entrance/exit/emphasis) snapshots
        // route through ApplyShapeAnimation. Both append to the MainSequence
        // ChildTimeNodeList in original order so animation[K] indexing holds.
        var shapeId = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value;
        if (shapeId.HasValue)
        {
            RemoveShapeAnimations(slidePart.Slide!, shapeId.Value);
            // RemoveShapeAnimations targets MainSequence groups that contain
            // a matching ShapeTarget; motion-path groups land in the same list
            // so they're removed too. Belt-and-suspenders: also drop motion
            // path animations explicitly in case the writer changes.
            RemoveAllMotionPathAnimationsForShape(slidePart.Slide!, shapeId.Value);
        }
        for (int i = 0; i < snapshots.Count; i++)
        {
            var snap = snapshots[i];
            if (snap.TryGetValue("class", out var snapCls)
                && snapCls.Equals("motion", StringComparison.OrdinalIgnoreCase))
            {
                ReapplyMotionFromSnapshot(slidePart, shape, snap);
            }
            else
            {
                var animValue = BuildAnimValueFromProps(snap);
                ApplyShapeAnimation(slidePart, shape, animValue);
            }
        }
        GetSlide(slidePart).Save();
        return [];
    }

    /// <summary>
    /// Re-emit a motion-path animation from a snapshot dict produced by
    /// PopulateAnimationNode. Resolves preset+direction back to a path string
    /// (falling back to d= for path=custom) and appends via AppendMotionPathAnimation.
    /// CONSISTENCY(animation-motion-presets).
    /// </summary>
    private static void ReapplyMotionFromSnapshot(
        SlidePart slidePart, Shape shape, Dictionary<string, string> snap)
    {
        string pathString;
        var preset = snap.GetValueOrDefault("path");
        if (string.IsNullOrEmpty(preset)
            || preset.Equals("custom", StringComparison.OrdinalIgnoreCase))
        {
            // Custom path: prefer d= override, else stored motionPath string.
            pathString = snap.GetValueOrDefault("d")
                ?? snap.GetValueOrDefault("motionPath")
                ?? "M 0 0 L 0 0 E";
        }
        else
        {
            var dir = snap.GetValueOrDefault("direction");
            pathString = GetMotionPresetPath(preset, dir)
                ?? snap.GetValueOrDefault("motionPath")
                ?? "M 0 0 L 0 0 E";
        }
        var duration = int.TryParse(snap.GetValueOrDefault("duration", "2000"),
            out var dv) ? dv : 2000;
        var trigger = snap.GetValueOrDefault("trigger", "onClick").ToLowerInvariant() switch
        {
            "afterprevious" => PowerPointHandler.AnimTrigger.AfterPrevious,
            "withprevious"  => PowerPointHandler.AnimTrigger.WithPrevious,
            _                => PowerPointHandler.AnimTrigger.OnClick
        };
        var delayMs = int.TryParse(snap.GetValueOrDefault("delay", "0"), out var dvL) ? dvL : 0;
        var easin   = int.TryParse(snap.GetValueOrDefault("easein", "0"), out var ein) ? ein * 1000 : 0;
        var easout  = int.TryParse(snap.GetValueOrDefault("easeout", "0"), out var eout) ? eout * 1000 : 0;
        AppendMotionPathAnimation(slidePart, shape, pathString, duration,
            trigger, delayMs, easin, easout);
    }

    /// <summary>
    /// Drop every motion-path animation group on the slide that targets the
    /// given shape. Mirrors RemoveShapeAnimations' walk-up to the MainSequence
    /// click-group par, narrowed to ctns carrying presetClass="motion".
    /// </summary>
    private static void RemoveAllMotionPathAnimationsForShape(Slide slide, uint shapeId)
    {
        var timing = slide.GetFirstChild<Timing>();
        if (timing == null) return;
        var spIdStr = shapeId.ToString();
        var toRemove = timing.Descendants<ShapeTarget>()
            .Where(st => st.ShapeId?.Value == spIdStr)
            .Select(st =>
            {
                OpenXmlElement? node = st;
                while (node?.Parent != null)
                {
                    if (node.Parent is ChildTimeNodeList ctl
                        && ctl.Parent is CommonTimeNode ctn
                        && ctn.NodeType?.Value == TimeNodeValues.MainSequence)
                        return node;
                    node = node.Parent;
                }
                return null;
            })
            .Where(n => n != null
                && n.Descendants<CommonTimeNode>().Any(c =>
                    c.GetAttributes().Any(a => a.LocalName == "presetClass" && a.Value == "motion")))
            .Distinct()
            .ToList();
        foreach (var n in toRemove) n!.Remove();
    }

    /// <summary>
    /// Render a property dictionary (as produced by PopulateAnimationNode + user
    /// overrides) into the composite animValue string parsed by ApplyShapeAnimation.
    /// CONSISTENCY(animation-set): mirrors AddAnimation's animValue assembly.
    /// </summary>
    private static string BuildAnimValueFromProps(Dictionary<string, string> p)
    {
        var effect = p.TryGetValue("effect", out var e) ? e : "fade";
        var (effectStripped, suffixCls) = ParseEffectClassSuffix(effect);
        effect = effectStripped;
        var cls = p.TryGetValue("class", out var c) ? c : (suffixCls ?? "entrance");
        var duration = p.TryGetValue("duration", out var d) ? d
            : p.TryGetValue("dur", out var d2) ? d2 : "500";
        var trigger = p.TryGetValue("trigger", out var t) ? t : "onclick";
        var triggerPart = trigger.ToLowerInvariant() switch
        {
            "onclick" or "click" => "click",
            "after" or "afterprevious" => "after",
            "with" or "withprevious" => "with",
            _ => throw new ArgumentException(
                $"Invalid animation trigger: '{trigger}'. Valid values: onclick, click, after, afterprevious, with, withprevious.")
        };
        var animValue = $"{effect}-{cls}-{duration}-{triggerPart}";
        if (p.TryGetValue("delay", out var del) && !string.IsNullOrEmpty(del))
            animValue += $"-delay={del}";
        if (p.TryGetValue("easein", out var ein) && !string.IsNullOrEmpty(ein))
            animValue += $"-easein={ein}";
        if (p.TryGetValue("easeout", out var eout) && !string.IsNullOrEmpty(eout))
            animValue += $"-easeout={eout}";
        if (p.TryGetValue("easing", out var eas) && !string.IsNullOrEmpty(eas))
            animValue += $"-easing={eas}";
        if (p.TryGetValue("direction", out var dir) && !string.IsNullOrEmpty(dir))
            animValue += $"-{dir}";
        if (p.TryGetValue("repeat", out var rep) && !string.IsNullOrEmpty(rep))
            animValue += $"-repeat={rep}";
        if (p.TryGetValue("restart", out var res) && !string.IsNullOrEmpty(res))
            animValue += $"-restart={res}";
        var arKey = p.TryGetValue("autoReverse", out var ar) ? ar
            : p.TryGetValue("autoreverse", out var ar2) ? ar2 : null;
        if (!string.IsNullOrEmpty(arKey))
            animValue += $"-autoReverse={arKey}";
        return animValue;
    }
}
