// Copyright 2025 OfficeCli (officecli.ai)
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
        var unsupported = SetRunOrShapeProperties(runOnlyProps, new List<Drawing.Run> { targetRun }, shape, slidePart, runContext: true);
        if (linkValRun != null) ApplyRunHyperlink(slidePart, targetRun, linkValRun, tooltipValRun);
        GetSlide(slidePart).Save();
        return unsupported;
    }

    private List<string> SetParagraphRunByPath(Match paraRunMatch, Dictionary<string, string> properties)
    {
        var slideIdx = int.Parse(paraRunMatch.Groups[1].Value);
        var shapeIdx = int.Parse(paraRunMatch.Groups[2].Value);
        var paraIdx = int.Parse(paraRunMatch.Groups[3].Value);
        var runIdx = int.Parse(paraRunMatch.Groups[4].Value);

        var (slidePart, shape) = ResolveShape(slideIdx, shapeIdx);
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
        var unsupported = SetRunOrShapeProperties(runOnlyProps, new List<Drawing.Run> { targetRun }, shape, slidePart, runContext: true);
        if (linkVal != null) ApplyRunHyperlink(slidePart, targetRun, linkVal, tooltipVal);
        GetSlide(slidePart).Save();
        return unsupported;
    }


    private List<string> SetParagraphByPath(Match paraMatch, Dictionary<string, string> properties)
    {
        var slideIdx = int.Parse(paraMatch.Groups[1].Value);
        var shapeIdx = int.Parse(paraMatch.Groups[2].Value);
        var paraIdx = int.Parse(paraMatch.Groups[3].Value);

        var (slidePart, shape) = ResolveShape(slideIdx, shapeIdx);
        var paragraphs = shape.TextBody?.Elements<Drawing.Paragraph>().ToList()
            ?? throw new ArgumentException("Shape has no text body");
        if (paraIdx < 1 || paraIdx > paragraphs.Count)
            throw new ArgumentException($"Paragraph {paraIdx} not found (shape has {paragraphs.Count} paragraphs)");

        var para = paragraphs[paraIdx - 1];
        var paraRuns = para.Elements<Drawing.Run>().ToList();
        var unsupported = new List<string>();

        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "align":
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
                    if (lsIsPercent)
                        pProps.AppendChild(new Drawing.LineSpacing(
                            new Drawing.SpacingPercent { Val = lsVal2 }));
                    else
                        pProps.AppendChild(new Drawing.LineSpacing(
                            new Drawing.SpacingPoints { Val = lsVal2 }));
                    break;
                }
                case "spacebefore" or "space.before":
                {
                    var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                    pProps.RemoveAllChildren<Drawing.SpaceBefore>();
                    pProps.AppendChild(new Drawing.SpaceBefore(new Drawing.SpacingPoints { Val = SpacingConverter.ParsePptSpacing(value) }));
                    break;
                }
                case "spaceafter" or "space.after":
                {
                    var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                    pProps.RemoveAllChildren<Drawing.SpaceAfter>();
                    pProps.AppendChild(new Drawing.SpaceAfter(new Drawing.SpacingPoints { Val = SpacingConverter.ParsePptSpacing(value) }));
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
                        new Dictionary<string, string> { { key, value } }, paraRuns, shape, slidePart, runContext: true);
                    unsupported.AddRange(runUnsup);
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
        var unsupported = new List<string>();
        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "name":
                    var nvGrpPr = grp.NonVisualGroupShapeProperties?.NonVisualDrawingProperties;
                    if (nvGrpPr != null) nvGrpPr.Name = value;
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
                    xfrm.Rotation = (int)(ParseHelpers.SafeParseDouble(value, "rotation") * 60000);
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
                            unsupported.Add($"{key} (valid group props: x, y, width, height, rotation, name, fill)");
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
                    if (nvCxnPr != null) nvCxnPr.Name = value;
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
                    var outline = spPr.GetFirstChild<Drawing.Outline>()
                        ?? spPr.AppendChild(new Drawing.Outline());
                    outline.Width = Core.EmuConverter.ParseLineWidth(value);
                    break;
                }
                case "linecolor" or "line.color" or "line" or "color":
                {
                    var spPr = cxn.ShapeProperties ?? (cxn.ShapeProperties = new ShapeProperties());
                    var outline = spPr.GetFirstChild<Drawing.Outline>()
                        ?? spPr.AppendChild(new Drawing.Outline());
                    var (rgb, _) = ParseHelpers.SanitizeColorForOoxml(value);
                    outline.RemoveAllChildren<Drawing.SolidFill>();
                    var newFill = new Drawing.SolidFill(
                        new Drawing.RgbColorModelHex { Val = rgb });
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
                case "linedash" or "line.dash":
                {
                    var spPr = cxn.ShapeProperties ?? (cxn.ShapeProperties = new ShapeProperties());
                    var outline = spPr.GetFirstChild<Drawing.Outline>()
                        ?? spPr.AppendChild(new Drawing.Outline());
                    outline.RemoveAllChildren<Drawing.PresetDash>();
                    var newDash = new Drawing.PresetDash { Val = value.ToLowerInvariant() switch
                    {
                        "solid" => Drawing.PresetLineDashValues.Solid,
                        "dot" => Drawing.PresetLineDashValues.Dot,
                        "dash" => Drawing.PresetLineDashValues.Dash,
                        "dashdot" or "dash_dot" => Drawing.PresetLineDashValues.DashDot,
                        "longdash" or "lgdash" or "lg_dash" => Drawing.PresetLineDashValues.LargeDash,
                        "longdashdot" or "lgdashdot" or "lg_dash_dot" => Drawing.PresetLineDashValues.LargeDashDot,
                        _ => throw new ArgumentException($"Invalid 'lineDash' value: '{value}'. Valid values: solid, dot, dash, dashdot, longdash, longdashdot.")
                    }};
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
                    var outline = spPr.GetFirstChild<Drawing.Outline>()
                        ?? spPr.AppendChild(new Drawing.Outline());
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
                    xfrm.Rotation = (int)(ParseHelpers.SafeParseDouble(value, "rotation") * 60000);
                    break;
                }
                case "preset" or "prstgeom" or "shape":
                {
                    // CONSISTENCY(canonical-key): schema canonical is 'shape';
                    // 'preset'/'prstgeom' retained as legacy aliases.
                    var spPr = cxn.ShapeProperties ?? (cxn.ShapeProperties = new ShapeProperties());
                    var prstGeom = spPr.GetFirstChild<Drawing.PresetGeometry>()
                        ?? spPr.AppendChild(new Drawing.PresetGeometry());
                    prstGeom.Preset = new Drawing.ShapeTypeValues(value);
                    break;
                }
                case "headend" or "headEnd":
                {
                    var spPr = cxn.ShapeProperties ?? (cxn.ShapeProperties = new ShapeProperties());
                    var outline = spPr.GetFirstChild<Drawing.Outline>()
                        ?? spPr.AppendChild(new Drawing.Outline());
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
                    var outline = spPr.GetFirstChild<Drawing.Outline>()
                        ?? spPr.AppendChild(new Drawing.Outline());
                    outline.RemoveAllChildren<Drawing.TailEnd>();
                    // CT_LineProperties: tailEnd is last — always append
                    outline.AppendChild(new Drawing.TailEnd { Type = ParseLineEndType(value) });
                    break;
                }
                default:
                    if (!GenericXmlQuery.SetGenericAttribute(cxn, key, value))
                    {
                        if (unsupported.Count == 0)
                            unsupported.Add($"{key} (valid connector props: line, color, fill, x, y, width, height, rotation, name, headEnd, tailEnd, geometry)");
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

        // Read current animation properties via PopulateAnimationNode, then merge
        // with user-provided overrides, then re-apply via the standard pipeline.
        // Limitation: like Set on /slide/shape with animation=, this replaces ALL
        // animations on the shape (the apply pipeline only knows how to add one).
        // CONSISTENCY(animation-set): mirrors Add's animValue string assembly.
        var existing = new DocumentNode { Path = "" };
        PopulateAnimationNode(existing, ctns[animIdx - 1]);

        string Get(string key, string? fallback = null)
            => properties.TryGetValue(key, out var v)
                ? v
                : (existing.Format.TryGetValue(key, out var ev) ? ev?.ToString() ?? fallback ?? "" : fallback ?? "");

        var effect = Get("effect", "fade");
        var cls = Get("class", "entrance");
        var duration = properties.TryGetValue("duration", out var dv) ? dv
            : properties.TryGetValue("dur", out var dv2) ? dv2
            : (existing.Format.TryGetValue("duration", out var ed) ? ed?.ToString() ?? "500" : "500");
        var trigger = Get("trigger", "onclick");
        var triggerPart = trigger.ToLowerInvariant() switch
        {
            "onclick" or "click" => "click",
            "after" or "afterprevious" => "after",
            "with" or "withprevious" => "with",
            _ => throw new ArgumentException(
                $"Invalid animation trigger: '{trigger}'. Valid values: onclick, click, after, afterprevious, with, withprevious.")
        };

        var animValue = $"{effect}-{cls}-{duration}-{triggerPart}";
        string? Resolve(string key)
            => properties.TryGetValue(key, out var pv) ? pv
             : (existing.Format.TryGetValue(key, out var ev) ? ev?.ToString() : null);
        var delayVal = Resolve("delay");
        if (!string.IsNullOrEmpty(delayVal)) animValue += $"-delay={delayVal}";
        var einVal = Resolve("easein");
        if (!string.IsNullOrEmpty(einVal)) animValue += $"-easein={einVal}";
        var eoutVal = Resolve("easeout");
        if (!string.IsNullOrEmpty(eoutVal)) animValue += $"-easeout={eoutVal}";
        if (properties.TryGetValue("easing", out var easing))
            animValue += $"-easing={easing}";
        if (properties.TryGetValue("direction", out var dir))
            animValue += $"-{dir}";

        var shapeId = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value;
        if (shapeId.HasValue)
            RemoveShapeAnimations(slidePart.Slide!, shapeId.Value);
        ApplyShapeAnimation(slidePart, shape, animValue);
        GetSlide(slidePart).Save();
        return [];
    }
}
