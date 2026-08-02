// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
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

        var targetRun = allRuns[PathIndex.ToArrayIndex(runIdx)];
        var linkValRun = properties.GetValueOrDefault("link");
        var tooltipValRun = properties.GetValueOrDefault("tooltip");
        var runOnlyProps = properties
            .Where(kv => !kv.Key.Equals("link", StringComparison.OrdinalIgnoreCase)
                      && !kv.Key.Equals("tooltip", StringComparison.OrdinalIgnoreCase))
            .ToDictionary(kv => kv.Key, kv => kv.Value);
        var unsupported = SetRunOrShapeProperties(runOnlyProps, new List<Drawing.Run> { targetRun }, shape, slidePart, runContext: true, unsupportedContextHint: RunPropsHint);
        if (linkValRun != null) ApplyRunHyperlink(slidePart, targetRun, linkValRun, tooltipValRun);
        // R7-8: tooltip-only Set on a run that already has a hyperlink — update the
        // existing hlinkClick's Tooltip attribute instead of silently no-oping (the
        // link= branch above is the only place tooltips were applied previously).
        else if (tooltipValRun != null)
        {
            var existingHlink = targetRun.RunProperties?.GetFirstChild<Drawing.HyperlinkOnClick>();
            if (existingHlink != null) existingHlink.Tooltip = tooltipValRun;
        }
        GetSlide(slidePart).Save();
        return unsupported;
    }

    // Context labels used by SetRunOrShapeProperties so paragraph/run paths
    // report paragraph-/run-valid props instead of the broader shape list.
    private const string RunPropsHint =
        "valid run props: text, bold, italic, underline, strike, color, fill, size, font, font.latin, font.ea, font.cs, link, tooltip, baseline, spacing, cap";
    private const string ParagraphPropsHint =
        "valid paragraph props: align, indent, level, marginLeft, marginRight, lineSpacing, spaceBefore, spaceAfter, tabs, link, tooltip — plus any run prop (applied to all runs in the paragraph)";

    // Run-level (character) property keys that, set on a RUNLESS paragraph, must
    // seed an empty run to land on (see SetParagraphOnShape). Paragraph-level
    // keys (align/indent/lineSpacing/spaceBefore/…) write to pPr and need no run.
    private static readonly HashSet<string> RunStyleParagraphKeys = new(StringComparer.OrdinalIgnoreCase)
    {
        "size", "font.size", "fontsize", "sz",
        "color", "fill",
        "font", "font.latin", "font.ea", "font.eastasia", "font.eastasian",
        "font.cs", "font.complexscript", "font.complex",
        "bold", "b", "italic", "i", "underline", "u", "strike",
        "spacing", "charspacing", "letterspacing", "spc",
        "baseline", "cap", "allcaps", "smallcaps", "kern",
    };

    // Copy every attribute and child (deep clone) from one
    // CT_TextCharacterProperties element (a:rPr / a:endParaRPr / a:defRPr) to
    // another. Used to round-trip run-style props through a scratch a:rPr when
    // the only sink is an empty paragraph's a:endParaRPr. Clones nodes rather
    // than round-tripping OuterXml so inherited xmlns prefixes stay resolvable.
    private static void CopyCharacterProps(OpenXmlElement source, OpenXmlElement target)
    {
        target.RemoveAllChildren();
        foreach (var attr in source.GetAttributes())
            target.SetAttribute(attr);
        foreach (var child in source.Elements())
            target.AppendChild(child.CloneNode(true));
    }

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
        var slidePart = slideParts[PathIndex.ToArrayIndex(slideIdx)];
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
        var slidePart = slideParts[PathIndex.ToArrayIndex(slideIdx)];
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
        var para = paragraphs[PathIndex.ToArrayIndex(paraIdx)];
        var paraRuns = para.Elements<Drawing.Run>().ToList();
        if (runIdx < 1 || runIdx > paraRuns.Count)
            throw new ArgumentException($"Run {runIdx} not found (paragraph has {paraRuns.Count} runs)");

        var targetRun = paraRuns[PathIndex.ToArrayIndex(runIdx)];
        var linkVal = properties.GetValueOrDefault("link");
        var tooltipVal = properties.GetValueOrDefault("tooltip");
        var runOnlyProps = properties
            .Where(kv => !kv.Key.Equals("link", StringComparison.OrdinalIgnoreCase)
                      && !kv.Key.Equals("tooltip", StringComparison.OrdinalIgnoreCase))
            .ToDictionary(kv => kv.Key, kv => kv.Value);
        var unsupported = SetRunOrShapeProperties(runOnlyProps, new List<Drawing.Run> { targetRun }, shape, slidePart, runContext: true, unsupportedContextHint: RunPropsHint);
        if (linkVal != null) ApplyRunHyperlink(slidePart, targetRun, linkVal, tooltipVal);
        // R7-8: tooltip-only Set on a run that already has a hyperlink — update the
        // existing hlinkClick's Tooltip attribute instead of silently no-oping.
        else if (tooltipVal != null)
        {
            var existingHlink = targetRun.RunProperties?.GetFirstChild<Drawing.HyperlinkOnClick>();
            if (existingHlink != null) existingHlink.Tooltip = tooltipVal;
        }
        GetSlide(slidePart).Save();
        return unsupported;
    }

    private List<string> SetParagraphOnShape(SlidePart slidePart, Shape shape, int paraIdx, Dictionary<string, string> properties)
    {
        var paragraphs = shape.TextBody?.Elements<Drawing.Paragraph>().ToList()
            ?? throw new ArgumentException("Shape has no text body");
        if (paraIdx < 1 || paraIdx > paragraphs.Count)
            throw new ArgumentException($"Paragraph {paraIdx} not found (shape has {paragraphs.Count} paragraphs)");

        var para = paragraphs[PathIndex.ToArrayIndex(paraIdx)];
        var paraRuns = para.Elements<Drawing.Run>().ToList();
        var unsupported = new List<string>();

        // Every property case below vivifies <a:pPr> on demand before its own
        // value validation. When a value is rejected on a paragraph that had
        // no pPr, the freshly-created empty <a:pPr/> would otherwise survive
        // the throw and get autosaved on Dispose — the error path must leave
        // the document untouched. One guard here covers all 14 vivify sites.
        var pPrExistedBefore = para.ParagraphProperties != null;
        try
        {
            return SetParagraphOnShapeCore(slidePart, shape, para, paraIdx, paraRuns, unsupported, properties);
        }
        catch
        {
            var vivified = para.ParagraphProperties;
            if (!pPrExistedBefore && vivified != null
                && !vivified.HasChildren && !vivified.GetAttributes().Any())
                vivified.Remove();
            throw;
        }
    }

    private List<string> SetParagraphOnShapeCore(SlidePart slidePart, Shape shape, Drawing.Paragraph para,
        int paraIdx, List<Drawing.Run> paraRuns, List<string> unsupported, Dictionary<string, string> properties)
    {

        // Empty (runless) paragraph carrying run-style props: route size / color
        // / font.* / bold / ... onto the paragraph's endParaRPr. Without a run the
        // default branch below calls SetRunOrShapeProperties with an empty run
        // list and silently drops them — only `cap` had an endParaRPr fallback.
        // Designers use a runless paragraph with a small endParaRPr font size as
        // a vertical spacer; PowerPoint sizes the empty line from the endParaRPr,
        // not from a seeded empty run. The dump configures the shape's
        // auto-seeded first paragraph via Set (not Add), so `set paragraph[1]
        // size=1pt` left the spacer's endParaRPr at the inherited body size
        // (16pt+), inflating its height and pushing all following body text down
        // on every text slide. Apply the props to a detached run (reusing the
        // full run-property logic), then transfer its rPr children onto the
        // endParaRPr. Skip when `text` is present (it builds its own runs).
        if (paraRuns.Count == 0 && !properties.ContainsKey("text")
            && properties.Keys.Any(k => RunStyleParagraphKeys.Contains(k)))
        {
            var endRPr = para.GetFirstChild<Drawing.EndParagraphRunProperties>();
            // rPr and endParaRPr share the CT_TextCharacterProperties content
            // model. Seed a scratch run's rPr from the existing endParaRPr (clone
            // attributes + children — not OuterXml, whose xmlns prefixes wouldn't
            // re-parse), mutate it through the full run-property path, then mirror
            // the result back onto the endParaRPr.
            var scratchRPr = new Drawing.RunProperties();
            if (endRPr != null) CopyCharacterProps(endRPr, scratchRPr);
            var scratchRun = new Drawing.Run { RunProperties = scratchRPr };
            var runStyleProps = properties.Where(kv => RunStyleParagraphKeys.Contains(kv.Key))
                .ToDictionary(kv => kv.Key, kv => kv.Value, StringComparer.OrdinalIgnoreCase);
            unsupported.AddRange(SetRunOrShapeProperties(runStyleProps, new List<Drawing.Run> { scratchRun },
                shape, slidePart, runContext: true, unsupportedContextHint: ParagraphPropsHint));
            var newEnd = new Drawing.EndParagraphRunProperties();
            CopyCharacterProps(scratchRun.RunProperties ?? scratchRPr, newEnd);
            if (endRPr != null) para.ReplaceChild(newEnd, endRPr);
            else para.AppendChild(newEnd);
            // Drop consumed keys so the default branch below doesn't re-process
            // them against the (still empty) run list.
            properties = properties.Where(kv => !RunStyleParagraphKeys.Contains(kv.Key))
                .ToDictionary(kv => kv.Key, kv => kv.Value, StringComparer.OrdinalIgnoreCase);
        }

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
                    // CONSISTENCY(pptx-bare-as-points): mirror AddParagraph.
                    pProps.Indent = (int)Math.Round(SpacingConverter.ParsePointsSigned(value) * EmuConverter.EmuPerPointF);
                    break;
                }
                case "liststyle" or "list" or "bullet":
                {
                    // Handle here (not via the shape-level fall-through) for two
                    // reasons: the fall-through delegates a SINGLE-key dict, so
                    // ApplyListStyle's list=none convenience (clear the hanging
                    // indent) couldn't see a sibling indent=/marginLeft= in the
                    // same Set call and erased it (sample19, replayed indent="0"
                    // vanished); and shape-level scope would restyle every
                    // paragraph instead of just this one.
                    var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                    ApplyListStyle(pProps, value, preserveIndent:
                        properties.ContainsKey("indent") || properties.ContainsKey("marginLeft")
                        || properties.ContainsKey("marginleft") || properties.ContainsKey("marL")
                        || properties.ContainsKey("marl"));
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
                case "bulletraw" or "bulletRaw":
                {
                    // Full bullet group (buClr/buFont/buSzPct/buChar/…) verbatim —
                    // round-trips colored/sized/Wingdings bullets the `list`
                    // keyword cannot represent.
                    var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                    ApplyBulletRaw(pProps, value);
                    break;
                }
                case "defrprraw" or "defRPrRaw":
                {
                    // Paragraph-level <a:defRPr> verbatim — bare runs inherit
                    // size/bold/font/color from it; see ApplyDefRPrRaw.
                    var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                    ApplyDefRPrRaw(pProps, value);
                    break;
                }
                case "marginleft" or "marl":
                {
                    var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                    // CONSISTENCY(pptx-bare-as-points): mirror AddParagraph.
                    pProps.LeftMargin = (int)Math.Round(SpacingConverter.ParsePointsSigned(value) * EmuConverter.EmuPerPointF);
                    break;
                }
                case "marginright" or "marr":
                {
                    var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                    pProps.RightMargin = (int)Math.Round(SpacingConverter.ParsePointsSigned(value) * EmuConverter.EmuPerPointF);
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
                // R65 bt-2: custom tab stops — same compact compound form
                // emitted by NodeBuilder, parsed via ParseTabStopList. Routed
                // through InsertPPrChild so a subsequent `set` that adds an
                // earlier-ranked sibling (e.g. spcAft) doesn't push tabLst
                // out of schema order (CONSISTENCY(schema-order-pptx)).
                case "tabs" or "tablist":
                {
                    var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                    // Validate-before-mutate: ParseTabStopList throws on invalid
                    // alignment tokens. Parse FIRST so a failed Set leaves the
                    // existing <a:tabLst> intact (no data loss).
                    var tabList = ParseTabStopList(value);
                    pProps.RemoveAllChildren<Drawing.TabStopList>();
                    if (tabList != null)
                        InsertPPrChild(pProps, tabList);
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
                case "direction" or "dir" or "rtl":
                {
                    // CONSISTENCY(canonical-keys): paragraph-path direction must
                    // ONLY touch pPr.RightToLeft. Falling through to the run
                    // helper (runContext:true) would also write <a:rPr rtl="1"/>
                    // on every run, which NodeBuilder then surfaces as the
                    // legacy alias Format["rtl"] on the run — and the batch
                    // emitter's single-run collapse merges that back into the
                    // paragraph set bag, producing {direction:rtl, rtl:true}
                    // and violating "Get returns the canonical key only"
                    // (the project conventions). The textbox-column-flow side effect
                    // (<a:bodyPr rtlCol="1"/>) belongs to shape-level Set,
                    // not paragraph-level. Run-level rtl remains reachable
                    // via /paragraph[K]/run[R] direction=rtl.
                    //
                    // R64 bt-2: write the attribute explicitly on BOTH rtl and
                    // ltr instead of clearing for ltr. An explicit Set on the
                    // paragraph is the caller's "override inheritance" signal —
                    // a paragraph inside a master/layout with rtl=1 silently
                    // inherits RTL when we strip the attribute on ltr, so the
                    // Set looks Updated but the persisted XML reads as a no-op
                    // (`<a:pPr/>` with no rtl attr). rtl="0" pins ltr regardless
                    // of inherited cascade. (Add path keeps strip-on-ltr so a
                    // freshly built ltr shape stays free of explicit-default
                    // noise on every paragraph.)
                    var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                    bool rtl = key.Equals("rtl", StringComparison.OrdinalIgnoreCase)
                        ? IsTruthy(value)
                        : ParsePptDirectionRtl(value);
                    pProps.RightToLeft = rtl;
                    break;
                }
                case "ealnbrk" or "ealinebreak"
                  or "latinlnbrk" or "latinlinebreak"
                  or "fontalgn" or "fontalignment"
                  or "deftabsz" or "defaulttabsize":
                {
                    // CJK / line-break pPr attributes — mirror AddParagraph + readback.
                    var pProps = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                    ApplyParagraphBreakProp(pProps, key, value);
                    break;
                }
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
                            para = refreshed[PathIndex.ToArrayIndex(paraIdx)];
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
        var slidePart = slideParts2[PathIndex.ToArrayIndex(slideIdx)];
        var shape = ResolvePlaceholderShape(slidePart, phId);

        // CONSISTENCY(placeholder-materialize-run): ResolvePlaceholderShape clones
        // a layout placeholder onto the slide with an empty paragraph (no run)
        // when materializing for the first time. Run-level properties (font /
        // size / bold / color / ...) iterate over `runs`, so an empty placeholder
        // would silently drop them. Seed a single empty run on the first
        // paragraph so the run-level Set has a target to write to — mirrors how
        // `set text=...` materializes runs by rebuilding the paragraph tree.
        var allRuns = shape.Descendants<Drawing.Run>().ToList();
        if (allRuns.Count == 0 && shape.TextBody != null && HasRunLevelProperty(properties))
        {
            var firstPara = shape.TextBody.Elements<Drawing.Paragraph>().FirstOrDefault();
            if (firstPara == null)
            {
                firstPara = new Drawing.Paragraph();
                shape.TextBody.Append(firstPara);
            }
            var seededRun = new Drawing.Run(
                new Drawing.RunProperties { Language = "en-US" },
                new Drawing.Text { Text = "" });
            var endParaRPr = firstPara.GetFirstChild<Drawing.EndParagraphRunProperties>();
            if (endParaRPr != null)
                firstPara.InsertBefore(seededRun, endParaRPr);
            else
                firstPara.Append(seededRun);
            allRuns = new List<Drawing.Run> { seededRun };
        }
        var unsupported = SetRunOrShapeProperties(properties, allRuns, shape, slidePart, unrecognizedLatex: LastUnrecognizedLatex);
        GetSlide(slidePart).Save();
        return unsupported;
    }

    private List<string> SetTitleByPath(Match titleMatch, Dictionary<string, string> properties)
    {
        var slideIdx = int.Parse(titleMatch.Groups[1].Value);
        var titleIdx = int.Parse(titleMatch.Groups[2].Value);

        var slideParts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > slideParts.Count)
            throw new ArgumentException($"Slide {slideIdx} not found (total: {slideParts.Count})");
        var slidePart = slideParts[PathIndex.ToArrayIndex(slideIdx)];

        Shape? shape = null;
        if (titleIdx == 1)
        {
            // Reuse placeholder resolution so layout-only titles are materialized.
            try { shape = ResolvePlaceholderShape(slidePart, "title"); }
            catch (ArgumentException) { shape = null; }
        }
        if (shape == null)
        {
            var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree
                ?? throw new ArgumentException($"Slide {slideIdx} has no shape tree");
            var titles = shapeTree.Elements<Shape>().Where(IsTitle).ToList();
            if (titleIdx < 1 || titleIdx > titles.Count)
                throw new ArgumentException($"Title {titleIdx} not found on slide {slideIdx} (slide has {titles.Count} title shape(s))");
            shape = titles[titleIdx - 1];
        }

        var allRuns = shape.Descendants<Drawing.Run>().ToList();
        if (allRuns.Count == 0 && shape.TextBody != null && HasRunLevelProperty(properties))
        {
            var firstPara = shape.TextBody.Elements<Drawing.Paragraph>().FirstOrDefault();
            if (firstPara == null)
            {
                firstPara = new Drawing.Paragraph();
                shape.TextBody.Append(firstPara);
            }
            var seededRun = new Drawing.Run(
                new Drawing.RunProperties { Language = "en-US" },
                new Drawing.Text { Text = "" });
            var endParaRPr = firstPara.GetFirstChild<Drawing.EndParagraphRunProperties>();
            if (endParaRPr != null)
                firstPara.InsertBefore(seededRun, endParaRPr);
            else
                firstPara.Append(seededRun);
            allRuns = new List<Drawing.Run> { seededRun };
        }
        var unsupported = SetRunOrShapeProperties(properties, allRuns, shape, slidePart, unrecognizedLatex: LastUnrecognizedLatex);
        GetSlide(slidePart).Save();
        return unsupported;
    }

    private static bool HasRunLevelProperty(Dictionary<string, string> properties)
    {
        foreach (var key in properties.Keys)
        {
            var k = key.ToLowerInvariant();
            if (k is "font" or "font.name" or "font.latin" or "font.ea" or "font.eastasia"
                or "font.eastasian" or "font.cs" or "font.complexscript" or "font.complex"
                or "size" or "fontsize" or "font.size"
                or "bold" or "font.bold" or "italic" or "font.italic"
                or "underline" or "strike" or "color" or "highlight"
                or "spacing" or "baseline" or "kern" or "cap" or "allcaps" or "smallcaps"
                or "lang" or "lang.latin")
                return true;
        }
        return false;
    }

    private List<string> SetGroupByPath(Match grpMatch, Dictionary<string, string> properties)
    {
        var slideIdx = int.Parse(grpMatch.Groups[1].Value);
        var grpIdx = int.Parse(grpMatch.Groups[2].Value);

        var slideParts6 = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > slideParts6.Count)
            throw new ArgumentException($"Slide {slideIdx} not found (total: {slideParts6.Count})");

        var slidePart = slideParts6[PathIndex.ToArrayIndex(slideIdx)];
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
        // Snapshot the group's size BEFORE this command so child font sizes can
        // be scaled by the NET resize exactly once (mirrors PowerPoint's
        // interactive group resize, which re-bakes font size; opening a file
        // whose group ext was edited directly does NOT re-bake it, so the CLI
        // must). Per-command, not per width/height iteration — otherwise a
        // combined `width+height` set would square the ratio.
        var preXfrm = grp.GroupShapeProperties?.TransformGroup;
        long preCx = preXfrm?.Extents?.Cx ?? 0;
        long preCy = preXfrm?.Extents?.Cy ?? 0;
        var unsupported = new List<string>();
        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "ungroup":
                    // Structural + terminal: dissolve the group, promoting members
                    // back onto the slide with transforms flattened to absolute
                    // coordinates. Inverse of AddGroup; ignores any sibling props.
                    if (IsTruthy(value))
                        return UngroupGroup(grp, shapeTree, slidePart);
                    break;
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
                case "x" or "y" or "left" or "top" or "width" or "height":
                {
                    var grpSpPr = grp.GroupShapeProperties ?? (grp.GroupShapeProperties = new GroupShapeProperties());
                    var xfrm = grpSpPr.TransformGroup ?? (grpSpPr.TransformGroup = new Drawing.TransformGroup());
                    var off = xfrm.Offset ?? (xfrm.Offset = new Drawing.Offset());
                    var ext = xfrm.Extents ?? (xfrm.Extents = new Drawing.Extents());
                    var keyLower = key.ToLowerInvariant() switch { "left" => "x", "top" => "y", var k => k };
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
                case "keepaspect":
                    // Consumed by the post-loop resize step below.
                    break;
                case "fill":
                {
                    // OOXML CT_GroupShapeProperties (p:grpSpPr) does NOT allow
                    // fill children — only xfrm/scene3d/extLst. PowerPoint
                    // silently ignores any solidFill/noFill/gradFill written
                    // here, so accepting the prop and writing the OOXML was
                    // misleading (the file rendered unchanged). Reject as
                    // unsupported so the caller learns the operation isn't
                    // supported instead of seeing a no-op succeed.
                    unsupported.Add($"{key} (p:grpSpPr does not support fill in OOXML; PowerPoint ignores any fill on a group — apply fill to the child shapes instead)");
                    break;
                }
                default:
                    if (!GenericXmlQuery.SetGenericAttribute(grp, key, value))
                    {
                        if (unsupported.Count == 0)
                            unsupported.Add($"{key} (valid group props: x, y, width, height, keepAspect, rotation, name, link, tooltip, ungroup)");
                        else
                            unsupported.Add(key);
                    }
                    break;
            }
        }
        var postExt = grp.GroupShapeProperties?.TransformGroup?.Extents;
        if (postExt != null && preCx > 0 && preCy > 0)
        {
            bool hasW = properties.Keys.Any(k => k.Equals("width", StringComparison.OrdinalIgnoreCase));
            bool hasH = properties.Keys.Any(k => k.Equals("height", StringComparison.OrdinalIgnoreCase));
            // keepAspect=true + a single dimension → scale the OTHER
            // proportionally so a diagram group stays aspect-correct (a lone
            // width/height stretch squashes a flowchart into slivers). Default
            // honors exactly the axes the caller passed, matching `set width`
            // on a plain shape (GitHub #237 — the lock must be opt-in).
            bool keepAspect = properties.Any(kv =>
                kv.Key.Equals("keepAspect", StringComparison.OrdinalIgnoreCase) && IsTruthy(kv.Value));
            if (keepAspect && hasW && !hasH) postExt.Cy = (long)Math.Round(preCy * ((postExt.Cx ?? preCx) / (double)preCx));
            else if (keepAspect && hasH && !hasW) postExt.Cx = (long)Math.Round(preCx * ((postExt.Cy ?? preCy) / (double)preCy));

            if ((postExt.Cx ?? 0) <= 0 || (postExt.Cy ?? 0) <= 0)
                throw new ArgumentException("Invalid group size: width and height must be positive.");

            // Re-bake child font sizes to match the net resize. fontRatio =
            // min(width-ratio, height-ratio): a uniform (or keepAspect) resize
            // keeps text exactly proportional; a single-axis shrink still scales
            // text down so it keeps fitting, and a single-axis grow leaves it.
            double fontRatio = Math.Min((postExt.Cx ?? preCx) / (double)preCx,
                                        (postExt.Cy ?? preCy) / (double)preCy);
            if (Math.Abs(fontRatio - 1.0) > 1e-6)
                ScaleGroupFontSizes(grp, fontRatio);
        }
        GetSlide(slidePart).Save();
        return unsupported;
    }

    // Dissolve a group: flatten each member's transform from the group's child
    // coordinate space into slide-absolute coordinates, promote the members onto
    // the slide (preserving z-order, just before where the group sat), then drop
    // the empty group. Inverse of AddGroup. Groups this handler creates use an
    // identity child space (chOff==off, chExt==ext), so members keep their exact
    // coordinates; groups scaled/moved after creation flatten through the affine.
    private List<string> UngroupGroup(GroupShape grp, ShapeTree tree, SlidePart slidePart)
    {
        var warnings = new List<string>();
        var xf = grp.GroupShapeProperties?.TransformGroup;
        long ox = xf?.Offset?.X ?? 0, oy = xf?.Offset?.Y ?? 0;
        long ecx = xf?.Extents?.Cx ?? 0, ecy = xf?.Extents?.Cy ?? 0;
        long cox = xf?.ChildOffset?.X ?? ox, coy = xf?.ChildOffset?.Y ?? oy;
        long ccx = xf?.ChildExtents?.Cx ?? ecx, ccy = xf?.ChildExtents?.Cy ?? ecy;
        double sx = ccx != 0 ? (double)ecx / ccx : 1.0;
        double sy = ccy != 0 ? (double)ecy / ccy : 1.0;
        if ((xf?.Rotation?.Value ?? 0) != 0)
            warnings.Add("ungroup (group has rotation; member positions flattened without rotation compensation)");

        var members = grp.Elements()
            .Where(e => e is Shape or Picture or ConnectionShape or GraphicFrame or GroupShape)
            .ToList();
        foreach (var member in members)
        {
            FlattenChildTransform(member, ox, oy, cox, coy, sx, sy);
            member.Remove();
            grp.InsertBeforeSelf(member); // promote in order, keeping the group's z-slot
        }
        grp.Remove();
        GetSlide(slidePart).Save();
        return warnings;
    }

    // Map a group member's (offset, extents) from the group's child coordinate
    // space to slide-absolute: abs = off + (child − chOff) · scale, size · scale.
    private static void FlattenChildTransform(OpenXmlElement member,
        long ox, long oy, long cox, long coy, double sx, double sy)
    {
        void Apply(Drawing.Offset? off, Drawing.Extents? ext)
        {
            if (off == null || ext == null) return;
            long cx = off.X ?? 0, cy = off.Y ?? 0, cw = ext.Cx ?? 0, ch = ext.Cy ?? 0;
            off.X = ox + (long)Math.Round((cx - cox) * sx);
            off.Y = oy + (long)Math.Round((cy - coy) * sy);
            ext.Cx = (long)Math.Round(cw * sx);
            ext.Cy = (long)Math.Round(ch * sy);
        }
        switch (member)
        {
            case Shape s: Apply(s.ShapeProperties?.Transform2D?.Offset, s.ShapeProperties?.Transform2D?.Extents); break;
            case Picture p: Apply(p.ShapeProperties?.Transform2D?.Offset, p.ShapeProperties?.Transform2D?.Extents); break;
            case ConnectionShape c: Apply(c.ShapeProperties?.Transform2D?.Offset, c.ShapeProperties?.Transform2D?.Extents); break;
            case GraphicFrame gf: Apply(gf.Transform?.Offset, gf.Transform?.Extents); break;
            case GroupShape g: Apply(g.GroupShapeProperties?.TransformGroup?.Offset, g.GroupShapeProperties?.TransformGroup?.Extents); break;
        }
    }

    // Multiply every descendant run's font size by <paramref name="ratio"/>
    // (floor 1pt). Covers explicit run props, the end-of-paragraph mark, and
    // list-style defaults so no text escapes the group's resize.
    private static void ScaleGroupFontSizes(GroupShape grp, double ratio)
    {
        void Scale(Int32Value? fs, Action<int> set)
        {
            if (fs != null) set(Math.Max(100, (int)Math.Round(fs.Value * ratio)));
        }
        foreach (var rp in grp.Descendants<Drawing.RunProperties>())
            Scale(rp.FontSize, v => rp.FontSize = v);
        foreach (var ep in grp.Descendants<Drawing.EndParagraphRunProperties>())
            Scale(ep.FontSize, v => ep.FontSize = v);
        foreach (var dp in grp.Descendants<Drawing.DefaultRunProperties>())
            Scale(dp.FontSize, v => dp.FontSize = v);
    }

    private List<string> SetConnectorByPath(Match cxnMatch, Dictionary<string, string> properties)
    {
        var slideIdx = int.Parse(cxnMatch.Groups[1].Value);
        var cxnIdx = int.Parse(cxnMatch.Groups[2].Value);

        var slideParts5 = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > slideParts5.Count)
            throw new ArgumentException($"Slide {slideIdx} not found (total: {slideParts5.Count})");

        var slidePart = slideParts5[PathIndex.ToArrayIndex(slideIdx)];
        var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree
            ?? throw new ArgumentException("Slide has no shape tree");
        var connectors = shapeTree.Elements<ConnectionShape>().ToList();
        if (cxnIdx < 1 || cxnIdx > connectors.Count)
            throw new ArgumentException($"Connector {cxnIdx} not found (total: {connectors.Count})");

        var cxn = connectors[cxnIdx - 1];
        return ApplyConnectorProps(slidePart, cxn, properties);
    }

    /// <summary>
    /// Apply connector property mutations to an already-resolved ConnectionShape.
    /// Shared by SetConnectorByPath (slide-level) and the group-inner connector
    /// Set route so a connector nested in a group gets the same property surface.
    /// </summary>
    private List<string> ApplyConnectorProps(SlidePart slidePart, ConnectionShape cxn, Dictionary<string, string> properties)
    {
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
                case "x" or "y" or "left" or "top" or "width" or "height":
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
                    EnsureOutlineHasFill(outline);
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
                    var cxnGrad = BuildGradientFill(NormalizeLineGradientSpec(value));
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
                    outline.RemoveAllChildren<Drawing.CustomDash>();
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
                // R64 bt-3: lineDashRaw — verbatim <a:custDash> passthrough on
                // connector Set. Mirrors lineDash schema-order install; clears
                // any preset/custom dash already on the outline since
                // CT_LineProperties choice EG_LineDashProperties accepts only
                // one. Empty value removes the dash entirely.
                case "linedashraw" or "line.dashraw":
                {
                    var spPr = cxn.ShapeProperties ?? (cxn.ShapeProperties = new ShapeProperties());
                    var outline = EnsureOutline(spPr);
                    outline.RemoveAllChildren<Drawing.PresetDash>();
                    outline.RemoveAllChildren<Drawing.CustomDash>();
                    if (!string.IsNullOrWhiteSpace(value))
                    {
                        var newCustDash = BuildCustomDashFromRaw(value);
                        var headEnd = outline.GetFirstChild<Drawing.HeadEnd>();
                        if (headEnd != null)
                            outline.InsertBefore(newCustDash, headEnd);
                        else
                        {
                            var tailEnd = outline.GetFirstChild<Drawing.TailEnd>();
                            if (tailEnd != null)
                                outline.InsertBefore(newCustDash, tailEnd);
                            else
                                outline.AppendChild(newCustDash);
                        }
                    }
                    break;
                }
                // R58 bt-3: lineCap (<a:ln cap="..."/>) — outline attribute,
                // previously dropped on connector Set. Mirror shape ShapeProperties
                // handling (rnd→round/sq→square aliases).
                case "linecap" or "line.cap":
                {
                    var spPr = cxn.ShapeProperties ?? (cxn.ShapeProperties = new ShapeProperties());
                    var outline = EnsureOutline(spPr);
                    outline.CapType = value.ToLowerInvariant() switch
                    {
                        "round" or "rnd" => Drawing.LineCapValues.Round,
                        "flat" => Drawing.LineCapValues.Flat,
                        "square" or "sq" => Drawing.LineCapValues.Square,
                        _ => throw new ArgumentException($"Invalid 'lineCap' value: '{value}'. Valid values: round, flat, square.")
                    };
                    break;
                }
                // R58 bt-3: cmpd (<a:ln cmpd="..."/>) — outline attribute,
                // previously dropped on connector Set. Mirror shape handling
                // (sng/dbl/thickThin/thinThick/tri tokens with single/double aliases).
                case "cmpd" or "compoundline" or "line.compound":
                {
                    var spPr = cxn.ShapeProperties ?? (cxn.ShapeProperties = new ShapeProperties());
                    var outline = EnsureOutline(spPr);
                    outline.CompoundLineType = value switch
                    {
                        var s when s.Equals("sng", StringComparison.OrdinalIgnoreCase) || s.Equals("single", StringComparison.OrdinalIgnoreCase)
                            => Drawing.CompoundLineValues.Single,
                        var s when s.Equals("dbl", StringComparison.OrdinalIgnoreCase) || s.Equals("double", StringComparison.OrdinalIgnoreCase)
                            => Drawing.CompoundLineValues.Double,
                        var s when s.Equals("thickThin", StringComparison.OrdinalIgnoreCase)
                            => Drawing.CompoundLineValues.ThickThin,
                        var s when s.Equals("thinThick", StringComparison.OrdinalIgnoreCase)
                            => Drawing.CompoundLineValues.ThinThick,
                        var s when s.Equals("tri", StringComparison.OrdinalIgnoreCase) || s.Equals("triple", StringComparison.OrdinalIgnoreCase)
                            => Drawing.CompoundLineValues.Triple,
                        _ => throw new ArgumentException($"Invalid 'cmpd' value: '{value}'. Valid values: sng, dbl, thickThin, thinThick, tri.")
                    };
                    break;
                }
                // R61 bt-2: lineJoin (<a:round/>|<a:bevel/>|<a:miter/>) — outline child,
                // previously dropped on connector Set. Mirror shape ShapeProperties
                // handling. Accept compound "miter:<lim>" so a single key carries
                // both the join token and the miter limit.
                case "linejoin" or "line.join":
                {
                    var spPr = cxn.ShapeProperties ?? (cxn.ShapeProperties = new ShapeProperties());
                    var outline = EnsureOutline(spPr);
                    outline.RemoveAllChildren<Drawing.Round>();
                    outline.RemoveAllChildren<Drawing.LineJoinBevel>();
                    outline.RemoveAllChildren<Drawing.Miter>();
                    var joinVal = value;
                    int? joinMiterLim = null;
                    var joinColon = value.IndexOf(':');
                    if (joinColon > 0)
                    {
                        joinVal = value.Substring(0, joinColon);
                        var limTok = value.Substring(joinColon + 1).Trim();
                        if (!int.TryParse(limTok, System.Globalization.NumberStyles.Integer,
                                System.Globalization.CultureInfo.InvariantCulture, out var limParsed))
                            throw new ArgumentException($"Invalid 'lineJoin' miter limit token: '{limTok}'. Expected integer (1000ths of a percent, e.g. 800000 = 800%).");
                        joinMiterLim = limParsed;
                    }
                    OpenXmlElement joinEl = joinVal.ToLowerInvariant() switch
                    {
                        "round" => new Drawing.Round(),
                        "bevel" => new Drawing.LineJoinBevel(),
                        "miter" => joinMiterLim.HasValue
                            ? new Drawing.Miter { Limit = joinMiterLim.Value }
                            : new Drawing.Miter(),
                        _ => throw new ArgumentException($"Invalid 'lineJoin' value: '{joinVal}'. Valid values: round, bevel, miter.")
                    };
                    // CT_LineProperties schema: ... → prstDash → (round|bevel|miter) → headEnd → tailEnd
                    var joinHeadEnd = outline.GetFirstChild<Drawing.HeadEnd>();
                    if (joinHeadEnd != null) outline.InsertBefore(joinEl, joinHeadEnd);
                    else
                    {
                        var joinTailEnd = outline.GetFirstChild<Drawing.TailEnd>();
                        if (joinTailEnd != null) outline.InsertBefore(joinEl, joinTailEnd);
                        else outline.AppendChild(joinEl);
                    }
                    break;
                }
                // R61 bt-2: miterLimit (<a:miter lim="N"/>) — extends an existing
                // <a:miter/> with the lim attribute, or auto-creates the miter join
                // if none was set. Value is OOXML 1000ths-of-a-percent.
                case "miterlimit" or "miter.limit" or "line.miterlimit":
                {
                    var spPr = cxn.ShapeProperties ?? (cxn.ShapeProperties = new ShapeProperties());
                    if (!int.TryParse(value, System.Globalization.NumberStyles.Integer,
                            System.Globalization.CultureInfo.InvariantCulture, out var limVal))
                        throw new ArgumentException($"Invalid 'miterLimit' value: '{value}'. Expected integer (1000ths of a percent, e.g. 800000 = 800%).");
                    var outline = EnsureOutline(spPr);
                    var miterEl = outline.GetFirstChild<Drawing.Miter>();
                    if (miterEl == null)
                    {
                        outline.RemoveAllChildren<Drawing.Round>();
                        outline.RemoveAllChildren<Drawing.LineJoinBevel>();
                        miterEl = new Drawing.Miter { Limit = limVal };
                        var mlHeadEnd = outline.GetFirstChild<Drawing.HeadEnd>();
                        if (mlHeadEnd != null) outline.InsertBefore(miterEl, mlHeadEnd);
                        else
                        {
                            var mlTailEnd = outline.GetFirstChild<Drawing.TailEnd>();
                            if (mlTailEnd != null) outline.InsertBefore(miterEl, mlTailEnd);
                            else outline.AppendChild(miterEl);
                        }
                    }
                    else
                    {
                        miterEl.Limit = limVal;
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
                        // Auto-create a black line fill.
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
                // R4-5: arrowhead size. @w (width) and @len (length) take the
                // sm/med/lg enum (CT_LineEndProperties). Set both to the same
                // size token so the marker scales uniformly. Targets the existing
                // headEnd/tailEnd (or creates one if none yet, defaulting Type=none
                // so size-only Set on a bare line still produces a sized marker).
                case "headend.size" or "headEnd.size":
                case "tailend.size" or "tailEnd.size":
                {
                    if (!TryParseLineEndSize(value, out var wEnum, out var lEnum))
                    { unsupported.Add($"{key} (value '{value}' must be one of: small/sm, medium/med, large/lg)"); break; }
                    var spPr = cxn.ShapeProperties ?? (cxn.ShapeProperties = new ShapeProperties());
                    var outline = EnsureOutline(spPr);
                    bool isHead = key.StartsWith("head", StringComparison.OrdinalIgnoreCase);
                    if (isHead)
                    {
                        var headEnd = outline.GetFirstChild<Drawing.HeadEnd>();
                        if (headEnd == null)
                        {
                            headEnd = new Drawing.HeadEnd { Type = Drawing.LineEndValues.None };
                            var existingTailEnd = outline.GetFirstChild<Drawing.TailEnd>();
                            if (existingTailEnd != null) outline.InsertBefore(headEnd, existingTailEnd);
                            else outline.AppendChild(headEnd);
                        }
                        headEnd.Width = wEnum;
                        headEnd.Length = lEnum;
                    }
                    else
                    {
                        var tailEnd = outline.GetFirstChild<Drawing.TailEnd>();
                        if (tailEnd == null)
                        {
                            tailEnd = new Drawing.TailEnd { Type = Drawing.LineEndValues.None };
                            outline.AppendChild(tailEnd);
                        }
                        tailEnd.Width = wEnum;
                        tailEnd.Length = lEnum;
                    }
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
                    var endpointShapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree
                        ?? throw new ArgumentException("Slide has no shape tree");
                    // Connectors are slide-local: reject a from/to path naming a
                    // different slide, which ResolveShapeId would otherwise silently
                    // bind to the same-index shape on this slide (mirror Add path).
                    var cxnSlideNum = GetSlideIndex(slidePart);
                    var xSlide = Regex.Match(value, @"^/slide\[(\d+)\]/");
                    if (xSlide.Success && int.Parse(xSlide.Groups[1].Value) != cxnSlideNum)
                        throw new ArgumentException(
                            $"Connector '{key}={value}' references a different slide; a connector can "
                            + $"only attach to shapes on its own slide (slide {cxnSlideNum}).");
                    var endpointId = ResolveShapeId(value, endpointShapeTree);
                    // Reject an unresolvable reference (garbage id, out-of-range
                    // index, or a shape nested in a group) instead of writing a
                    // dangling stCxn/endCxn — mirror the Add path's RequireReachableFrame.
                    if (FrameBoundsById(endpointShapeTree, endpointId) == null)
                        throw new ArgumentException(
                            $"Connector '{key}={value}' does not resolve to a shape on this slide "
                            + "(no top-level frame with that id/index exists; note a shape nested inside a group cannot be a connector endpoint).");
                    var cxnDrawProps = cxn.NonVisualConnectionShapeProperties
                        ?.GetFirstChild<NonVisualConnectorShapeDrawingProperties>();
                    if (cxnDrawProps == null) { unsupported.Add(key); break; }
                    bool isStart = key.Equals("from", StringComparison.OrdinalIgnoreCase)
                        || key.Equals("startshape", StringComparison.OrdinalIgnoreCase);
                    // Assign via the typed property (not RemoveAllChildren+AppendChild):
                    // the SDK setter places stCxn/endCxn in canonical schema order
                    // (cxnSpLocks, stCxn, endCxn). Appending after an existing endCxn
                    // when setting 'from' later would emit stCxn AFTER endCxn, which is
                    // schema-invalid — PowerPoint silently drops the out-of-order child
                    // and loses the start attachment.
                    if (isStart)
                        cxnDrawProps.StartConnection = new Drawing.StartConnection
                        { Id = endpointId, Index = ParseConnectorIdx(properties, "fromIdx", "fromidx", "startIdx", "startidx") };
                    else
                        cxnDrawProps.EndConnection = new Drawing.EndConnection
                        { Id = endpointId, Index = ParseConnectorIdx(properties, "toIdx", "toidx", "endIdx", "endidx") };

                    // R14-4: recompute the connector's xfrm bounding box from the
                    // (possibly new) endpoint centers — mirror Add.Misc.cs connector
                    // wiring. Without this the <a:xfrm> stays pinned to the old
                    // endpoints, so a reconnect to a far-away shape leaves a stale
                    // (often near-zero) connector box. Only adjust when no explicit
                    // x/y/width/height is being set in the same call.
                    bool hasExplicitBox = properties.ContainsKey("x") || properties.ContainsKey("left")
                        || properties.ContainsKey("y") || properties.ContainsKey("top")
                        || properties.ContainsKey("width") || properties.ContainsKey("height");
                    if (!hasExplicitBox)
                        RecomputeConnectorBox(cxn, endpointShapeTree,
                            NormalizeConnectorSide(properties, "fromSide", "fromside", "startSide", "startside"),
                            NormalizeConnectorSide(properties, "toSide", "toside", "endSide", "endside"));
                    break;
                }
                // Edge anchoring on an existing connector without changing endpoints:
                // `set connector --prop fromSide=right --prop toSide=left`. Recompute
                // the xfrm from the current connected shapes using the requested edges
                // so the drawn line moves to those edges (PowerPoint renders our
                // offset/extent, not stCxn/endCxn). Only meaningful when both endpoints
                // resolve to frames; a no-op box recompute is harmless otherwise.
                case "fromside" or "startside":
                case "toside" or "endside":
                {
                    // Handled once per Set call — skip if the sibling side key already
                    // triggered the recompute, and skip when from/to is also changing
                    // (that case recomputes with sides above).
                    if (key is "toside" or "endside"
                        && (properties.ContainsKey("fromSide") || properties.ContainsKey("fromside")
                            || properties.ContainsKey("startSide") || properties.ContainsKey("startside")))
                        break;
                    if (properties.ContainsKey("from") || properties.ContainsKey("startshape")
                        || properties.ContainsKey("to") || properties.ContainsKey("endshape"))
                        break;
                    bool hasBox = properties.ContainsKey("x") || properties.ContainsKey("left")
                        || properties.ContainsKey("y") || properties.ContainsKey("top")
                        || properties.ContainsKey("width") || properties.ContainsKey("height");
                    if (hasBox) break;
                    var sideShapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree
                        ?? throw new ArgumentException("Slide has no shape tree");
                    // Compose sequential single-side sets: sides aren't persisted as
                    // connector state, but on a pure `set` (no shape moved) the stored
                    // geometry still sits on the last-chosen edges, so re-derive the
                    // side that isn't being changed from the current endpoint position
                    // instead of resetting it to auto. This makes
                    //   set fromSide=top ; set toSide=bottom
                    // end up top→bottom rather than clobbering fromSide back to auto.
                    var (rFrom, rTo) = ResolveSidesFromGeometry(cxn, sideShapeTree,
                        NormalizeConnectorSide(properties, "fromSide", "fromside", "startSide", "startside"),
                        NormalizeConnectorSide(properties, "toSide", "toside", "endSide", "endside"));
                    RecomputeConnectorBox(cxn, sideShapeTree, rFrom, rTo);
                    break;
                }
                // Low-level connection-site index on an existing connector. Mutates
                // the <a:stCxn/endCxn idx> on the current endpoint (no-op if the
                // endpoint isn't connected). Does not move the drawn line — that is
                // offset/extent, driven by fromSide/toSide.
                case "fromidx" or "startidx":
                {
                    var d = cxn.NonVisualConnectionShapeProperties
                        ?.GetFirstChild<NonVisualConnectorShapeDrawingProperties>();
                    var st = d?.GetFirstChild<Drawing.StartConnection>();
                    if (st != null) st.Index = ParseConnectorIdx(properties, key);
                    else unsupported.Add($"{key} (connector has no start connection; set 'from' first)");
                    break;
                }
                case "toidx" or "endidx":
                {
                    var d = cxn.NonVisualConnectionShapeProperties
                        ?.GetFirstChild<NonVisualConnectorShapeDrawingProperties>();
                    var en = d?.GetFirstChild<Drawing.EndConnection>();
                    if (en != null) en.Index = ParseConnectorIdx(properties, key);
                    else unsupported.Add($"{key} (connector has no end connection; set 'to' first)");
                    break;
                }
                // R15-4: connector text label. Mirrors the Add.Misc txBody write
                // path so `set connector --prop text=...` is accepted symmetrically.
                // Replaces any existing txBody (typed or the SDK-unknown reparse form)
                // with a fresh single-paragraph single-run label.
                case "text":
                {
                    Core.XmlTextValidator.ValidateOrThrow(value, "text", allowSoftBreakChar: true);
                    cxn.RemoveAllChildren<DocumentFormat.OpenXml.Presentation.TextBody>();
                    foreach (var unk in cxn.ChildElements.OfType<OpenXmlUnknownElement>()
                                 .Where(e => e.LocalName == "txBody").ToList())
                        unk.Remove();
                    var cxnRunProps = new Drawing.RunProperties { Language = "en-US" };
                    var cxnPara = new Drawing.Paragraph(new Drawing.Run(cxnRunProps,
                        MakePreservingText(value)));
                    var cxnTxBody = new DocumentFormat.OpenXml.Presentation.TextBody(
                        new Drawing.BodyProperties(),
                        new Drawing.ListStyle(),
                        cxnPara);
                    cxn.AppendChild(cxnTxBody);
                    break;
                }
                default:
                    if (!GenericXmlQuery.SetGenericAttribute(cxn, key, value))
                    {
                        if (unsupported.Count == 0)
                            unsupported.Add($"{key} (valid connector props: line, color, fill, x, y, width, height, rotation, name, headEnd, tailEnd, geometry, from, to, text)");
                        else
                            unsupported.Add(key);
                    }
                    break;
            }
        }
        GetSlide(slidePart).Save();
        return unsupported;
    }

    /// <summary>
    /// Recompute a connector's &lt;a:xfrm&gt; bounding box from its connected
    /// shapes' current centers — the same derivation Add.Misc.cs uses when a
    /// connector is authored with from=/to=. PowerPoint trusts the stored
    /// offset/extent at render time (it does not recompute from stCxn/endCxn),
    /// so whenever an endpoint's geometry changes the box must be refreshed or
    /// the line detaches from the shape it is glued to. No-op unless the
    /// connector has ShapeProperties and both endpoints resolve to frames.
    /// </summary>
    /// <summary>
    /// Parse a connector connection-site index (fromIdx/toIdx and legacy
    /// startIdx/endIdx aliases) from the property bag; defaults to 0 when absent
    /// or non-numeric, matching Add.Misc.cs's connector idx handling.
    /// </summary>
    private static uint ParseConnectorIdx(Dictionary<string, string> p, params string[] keys)
    {
        foreach (var k in keys)
        {
            if (!p.TryGetValue(k, out var raw)) continue;
            var t = raw?.Trim();
            if (string.IsNullOrEmpty(t)) continue;
            // Reject garbage instead of defaulting to 0 (mirror the Add path and
            // NormalizeConnectorSide's strict validation).
            if (uint.TryParse(t, System.Globalization.NumberStyles.Integer,
                    System.Globalization.CultureInfo.InvariantCulture, out var v))
                return v;
            throw new ArgumentException(
                $"Invalid connector index '{raw}' for '{k}': expected a non-negative integer.");
        }
        return 0;
    }

    /// <summary>
    /// Given explicit fromSide/toSide (either may be null), fill any null side by
    /// classifying the connector's current stored endpoint against its connected
    /// shape's edges. Lets two sequential single-side `set` calls compose without
    /// persisting side state — valid only on a pure `set` where the stored xfrm
    /// still reflects the last edge choice (not after a shape move, which is why
    /// the move-reroute path keeps using auto). Falls back to the explicit values
    /// (i.e. auto) when endpoints/geometry can't be resolved.
    /// </summary>
    private static (string? from, string? to) ResolveSidesFromGeometry(
        ConnectionShape cxn, ShapeTree tree, string? explicitFrom, string? explicitTo)
    {
        if (explicitFrom != null && explicitTo != null) return (explicitFrom, explicitTo);
        var draw = cxn.NonVisualConnectionShapeProperties
            ?.GetFirstChild<NonVisualConnectorShapeDrawingProperties>();
        var startId = draw?.GetFirstChild<Drawing.StartConnection>()?.Id?.Value;
        var endId = draw?.GetFirstChild<Drawing.EndConnection>()?.Id?.Value;
        var xf = cxn.ShapeProperties?.Transform2D;
        if (xf == null || !startId.HasValue || !endId.HasValue) return (explicitFrom, explicitTo);
        var sBox = FrameBoundsById(tree, startId.Value);
        var eBox = FrameBoundsById(tree, endId.Value);
        if (!sBox.HasValue || !eBox.HasValue) return (explicitFrom, explicitTo);
        long ox = xf.Offset?.X?.Value ?? 0, oy = xf.Offset?.Y?.Value ?? 0;
        long cx = xf.Extents?.Cx?.Value ?? 0, cy = xf.Extents?.Cy?.Value ?? 0;
        bool fh = xf.HorizontalFlip?.Value == true, fv = xf.VerticalFlip?.Value == true;
        long p1x = fh ? ox + cx : ox, p2x = fh ? ox : ox + cx;
        long p1y = fv ? oy + cy : oy, p2y = fv ? oy : oy + cy;
        return (explicitFrom ?? NearestConnectorSide(p1x, p1y, sBox.Value),
                explicitTo ?? NearestConnectorSide(p2x, p2y, eBox.Value));
    }

    /// <summary>Classify a point to the nearest of a box's five anchor points
    /// (left/right/top/bottom edge midpoints, or center).</summary>
    private static string NearestConnectorSide(long px, long py, (long x, long y, long cx, long cy) b)
    {
        (string s, long x, long y)[] c =
        {
            ("left", b.x, b.y + b.cy / 2), ("right", b.x + b.cx, b.y + b.cy / 2),
            ("top", b.x + b.cx / 2, b.y), ("bottom", b.x + b.cx / 2, b.y + b.cy),
            ("center", b.x + b.cx / 2, b.y + b.cy / 2),
        };
        string best = "center"; double bd = double.MaxValue;
        foreach (var k in c)
        {
            double dx = px - k.x, dy = py - k.y, d = dx * dx + dy * dy;
            if (d < bd) { bd = d; best = k.s; }
        }
        return best;
    }

    private static void RecomputeConnectorBox(ConnectionShape cxn, ShapeTree tree,
        string? fromSide = null, string? toSide = null)
    {
        var draw = cxn.NonVisualConnectionShapeProperties
            ?.GetFirstChild<NonVisualConnectorShapeDrawingProperties>();
        var startId = draw?.GetFirstChild<Drawing.StartConnection>()?.Id?.Value;
        var endId = draw?.GetFirstChild<Drawing.EndConnection>()?.Id?.Value;
        // A derived box needs BOTH endpoints anchored to real shapes. If only one
        // is connected (or a box can't be resolved), leave the existing geometry
        // alone rather than collapsing both ends onto the single shape — that
        // produced a diagonal corner-clip when setting a side on a one-ended
        // connector, and detaches nothing useful on a shape move.
        if (!startId.HasValue || !endId.HasValue) return;
        var startBox = FrameBoundsById(tree, startId.Value);
        var endBox = FrameBoundsById(tree, endId.Value);
        if (!startBox.HasValue || !endBox.HasValue) return;
        var pStart = startBox;
        var pEnd = endBox;
        var (p1x, p1y, p2x, p2y) = ComputeConnectorEndpoints(
            pStart.Value, pEnd.Value, fromSide, toSide);
        var spPr = cxn.ShapeProperties;
        if (spPr == null) return;
        var xfrm = spPr.Transform2D;
        if (xfrm == null) { xfrm = new Drawing.Transform2D(); spPr.InsertAt(xfrm, 0); }
        xfrm.Offset ??= new Drawing.Offset();
        xfrm.Extents ??= new Drawing.Extents();
        xfrm.Offset.X = Math.Min(p1x, p2x);
        xfrm.Offset.Y = Math.Min(p1y, p2y);
        xfrm.Extents.Cx = Math.Abs(p2x - p1x);
        xfrm.Extents.Cy = Math.Abs(p2y - p1y);
        xfrm.HorizontalFlip = p2x < p1x;
        xfrm.VerticalFlip = p2y < p1y;
    }

    /// <summary>
    /// After a shape's geometry (x/y/width/height) changes, re-glue every
    /// top-level connector whose start/end connection references it so the
    /// connector tracks the moved shape — PowerPoint's interactive
    /// "connector follows shape" behavior. Without this the connector's stored
    /// xfrm stays pinned to the shape's old position and the line visibly
    /// detaches.
    /// CONSISTENCY(connector-reroute): wired from the plain-shape Set path,
    /// which is what the editor drives on drag/resize; picture/chart/group
    /// endpoint moves are a future extension through this same helper.
    /// </summary>
    private void RerouteConnectorsForShape(SlidePart slidePart, uint movedShapeId)
    {
        var tree = GetSlide(slidePart).CommonSlideData?.ShapeTree;
        if (tree == null) return;
        foreach (var cxn in tree.Elements<ConnectionShape>())
        {
            var draw = cxn.NonVisualConnectionShapeProperties
                ?.GetFirstChild<NonVisualConnectorShapeDrawingProperties>();
            var startId = draw?.GetFirstChild<Drawing.StartConnection>()?.Id?.Value;
            var endId = draw?.GetFirstChild<Drawing.EndConnection>()?.Id?.Value;
            if (startId != movedShapeId && endId != movedShapeId) continue;
            RecomputeConnectorBox(cxn, tree);
        }
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
        var slidePart = slideParts[PathIndex.ToArrayIndex(slideIdx)];
        var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree
            ?? throw new ArgumentException("Slide has no shape tree");
        var groups = shapeTree.Elements<GroupShape>().ToList();
        if (grpIdx < 1 || grpIdx > groups.Count)
            throw new ArgumentException($"Group {grpIdx} not found (total: {groups.Count})");
        var grp = groups[grpIdx - 1];
        var innerShapes = grp.Elements<Shape>().ToList();
        if (shapeIdx < 1 || shapeIdx > innerShapes.Count)
            throw new ArgumentException($"Shape {shapeIdx} not found in group {grpIdx} (total: {innerShapes.Count})");
        return ApplyShapePropsCore(slidePart, innerShapes[PathIndex.ToArrayIndex(shapeIdx)], properties);
    }

    /// <summary>
    /// R14-5: Set on a table/chart GraphicFrame nested in a group:
    /// /slide[N]/group[M](/group[L])*/(table|chart)[K]. Walks the group chain the
    /// same way Query.cs nestedGroupMatch does, resolves the leaf GraphicFrame, then
    /// reuses the table prop core (or the chart relationship + ChartHelper) so a
    /// grouped table/chart accepts the same props as a top-level one.
    /// </summary>
    private List<string> SetGroupInnerFrameByPath(Match match, Dictionary<string, string> properties)
    {
        var slideIdx = int.Parse(match.Groups[1].Value);
        var rootGrpIdx = int.Parse(match.Groups[2].Value);
        var nestedSegs = match.Groups[3].Value;
        var leafType = match.Groups[4].Value.ToLowerInvariant();
        var leafIdx = int.Parse(match.Groups[5].Value);

        var slideParts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > slideParts.Count)
            throw new ArgumentException($"Slide {slideIdx} not found (total: {slideParts.Count})");
        var slidePart = slideParts[PathIndex.ToArrayIndex(slideIdx)];
        var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree
            ?? throw new ArgumentException("Slide has no shape tree");
        var rootGroups = shapeTree.Elements<GroupShape>().ToList();
        if (rootGrpIdx < 1 || rootGrpIdx > rootGroups.Count)
            throw new ArgumentException($"Group {rootGrpIdx} not found (total: {rootGroups.Count})");
        var current = rootGroups[rootGrpIdx - 1];
        foreach (Match seg in Regex.Matches(nestedSegs, @"/group\[(\d+)\]"))
        {
            var subIdx = int.Parse(seg.Groups[1].Value);
            var subs = current.Elements<GroupShape>().ToList();
            if (subIdx < 1 || subIdx > subs.Count)
                throw new ArgumentException($"Nested group {subIdx} not found (total: {subs.Count})");
            current = subs[subIdx - 1];
        }

        if (leafType == "table")
        {
            var inner = current.Elements<GraphicFrame>()
                .Where(gf => gf.Descendants<Drawing.Table>().Any()).ToList();
            if (leafIdx < 1 || leafIdx > inner.Count)
                throw new ArgumentException($"Table {leafIdx} not found in group (total: {inner.Count})");
            return SetTablePropsCore(slidePart, inner[leafIdx - 1], properties);
        }
        else // chart
        {
            var inner = current.Elements<GraphicFrame>()
                .Where(gf => gf.Descendants<DocumentFormat.OpenXml.Drawing.Charts.ChartReference>().Any()
                    || IsExtendedChartFrame(gf)).ToList();
            if (leafIdx < 1 || leafIdx > inner.Count)
                throw new ArgumentException($"Chart {leafIdx} not found in group (total: {inner.Count})");
            var chartGf = inner[leafIdx - 1];
            var unsupported = new List<string>();
            var chartProps = new Dictionary<string, string>();
            foreach (var (key, value) in properties)
            {
                switch (key.ToLowerInvariant())
                {
                    case "x" or "y" or "left" or "top" or "width" or "height":
                    {
                        var xfrm = chartGf.Transform ?? (chartGf.Transform = new Transform());
                        TryApplyPositionSize(key.ToLowerInvariant(), value,
                            xfrm.Offset ?? (xfrm.Offset = new Drawing.Offset()),
                            xfrm.Extents ?? (xfrm.Extents = new Drawing.Extents()));
                        break;
                    }
                    case "name":
                        var nvPr = chartGf.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties;
                        if (nvPr != null) { Core.XmlTextValidator.ValidateOrThrow(value, "name"); nvPr.Name = value; }
                        break;
                    default:
                        chartProps[key] = value;
                        break;
                }
            }
            if (chartProps.Count > 0)
            {
                var chartRef = chartGf.Descendants<DocumentFormat.OpenXml.Drawing.Charts.ChartReference>().FirstOrDefault();
                if (chartRef?.Id?.Value != null && slidePart.GetPartById(chartRef.Id.Value) is ChartPart cp)
                    unsupported.AddRange(ChartHelper.SetChartProperties(cp, chartProps));
                else
                    unsupported.AddRange(chartProps.Keys);
            }
            GetSlide(slidePart).Save();
            return unsupported;
        }
    }

    /// <summary>
    /// CONSISTENCY(group-inner-shape): arbitrary-depth Set on
    /// /slide[N]/group[M](/group[L])+/shape[K]. Mirrors Query.cs:836
    /// nestedGroupMatch's walk — descend each /group[L] segment in
    /// order, then resolve shape[K] inside the final group.
    /// </summary>
    private List<string> SetNestedGroupInnerShapeByPath(Match match, Dictionary<string, string> properties)
    {
        var slideIdx = int.Parse(match.Groups[1].Value);
        var rootGrpIdx = int.Parse(match.Groups[2].Value);
        var nestedSegs = match.Groups[3].Value;
        var shapeIdx = int.Parse(match.Groups[4].Value);

        var slideParts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > slideParts.Count)
            throw new ArgumentException($"Slide {slideIdx} not found (total: {slideParts.Count})");
        var slidePart = slideParts[PathIndex.ToArrayIndex(slideIdx)];
        var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree
            ?? throw new ArgumentException("Slide has no shape tree");
        var rootGroups = shapeTree.Elements<GroupShape>().ToList();
        if (rootGrpIdx < 1 || rootGrpIdx > rootGroups.Count)
            throw new ArgumentException($"Group {rootGrpIdx} not found (total: {rootGroups.Count})");
        var current = rootGroups[rootGrpIdx - 1];
        var depth = 1;
        foreach (Match seg in Regex.Matches(nestedSegs, @"/group\[(\d+)\]"))
        {
            depth++;
            var subIdx = int.Parse(seg.Groups[1].Value);
            var subs = current.Elements<GroupShape>().ToList();
            if (subIdx < 1 || subIdx > subs.Count)
                throw new ArgumentException($"Nested group {subIdx} not found at depth {depth} (total: {subs.Count})");
            current = subs[subIdx - 1];
        }
        var innerShapes = current.Elements<Shape>().ToList();
        if (shapeIdx < 1 || shapeIdx > innerShapes.Count)
            throw new ArgumentException($"Shape {shapeIdx} not found in nested group (total: {innerShapes.Count})");
        return ApplyShapePropsCore(slidePart, innerShapes[PathIndex.ToArrayIndex(shapeIdx)], properties);
    }

    /// <summary>
    /// Resolve a Shape nested inside a group on a slide and return it
    /// alongside the owning SlidePart. Used by group-paragraph / group-run
    /// setters so the dispatch tier can pass control to the existing
    /// SetParagraphOnShape / SetParagraphRunOnShape helpers without
    /// duplicating navigation logic.
    /// </summary>
    /// <summary>
    /// Resolve a Shape inside a (possibly nested) group from the raw
    /// "/group[K]/group[L]…" segment string and a final shape index. Walks each
    /// group segment in order, then picks shape[shapeIdx] inside the deepest
    /// group. Shared by AddParagraph's grouped-shape parent route.
    /// </summary>
    /// <summary>
    /// Resolve a ConnectionShape inside a (possibly nested) group. Each group
    /// segment and the final connector segment accept either a positional index
    /// ([N]) or an @id selector ([@id=K]) — dump addresses grouped connectors by
    /// @id. Walks every group segment, then locates the connector in the deepest
    /// group.
    /// </summary>
    private (SlidePart slidePart, ConnectionShape cxn) ResolveGroupInnerConnector(
        int slideIdx, string groupSegs, string connectorToken)
    {
        var slideParts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > slideParts.Count)
            throw new ArgumentException($"Slide {slideIdx} not found (total: {slideParts.Count})");
        var slidePart = slideParts[PathIndex.ToArrayIndex(slideIdx)];
        var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree
            ?? throw new ArgumentException("Slide has no shape tree");

        OpenXmlCompositeElement current = shapeTree;
        int depth = 0;
        foreach (Match seg in Regex.Matches(groupSegs, @"/group\[([^\]]+)\]"))
        {
            depth++;
            var groups = current.Elements<GroupShape>().ToList();
            var token = seg.Groups[1].Value;
            int grpIdx = ResolveContainerChildIndex(token, groups.Count,
                i => groups[i].NonVisualGroupShapeProperties?.NonVisualDrawingProperties);
            current = groups[grpIdx - 1];
        }
        var connectors = current.Elements<ConnectionShape>().ToList();
        int cxnIdx = ResolveContainerChildIndex(connectorToken, connectors.Count,
            i => connectors[i].NonVisualConnectionShapeProperties?.NonVisualDrawingProperties);
        return (slidePart, connectors[cxnIdx - 1]);
    }

    /// <summary>
    /// Resolve a child selector token — "N" (1-based positional) or "@id=K" — to
    /// a 1-based index within a container's typed child list. nvAt returns the
    /// NonVisualDrawingProperties for the i-th (0-based) child so @id matching can
    /// read the cNvPr id.
    /// </summary>
    private static int ResolveContainerChildIndex(string token, int count,
        Func<int, NonVisualDrawingProperties?> nvAt)
    {
        var idMatch = Regex.Match(token, @"^@id=(\d+)$");
        if (idMatch.Success)
        {
            var wantId = idMatch.Groups[1].Value;
            for (int i = 0; i < count; i++)
                if (nvAt(i)?.Id?.Value.ToString() == wantId)
                    return i + 1;
            throw new ArgumentException($"No child found with @id={wantId} (total: {count})");
        }
        if (int.TryParse(token, out var pos))
        {
            if (pos < 1 || pos > count)
                throw new ArgumentException($"Index {pos} out of range (total: {count})");
            return pos;
        }
        throw new ArgumentException($"Unsupported selector token '{token}': expected an index or @id=K.");
    }

    private (SlidePart slidePart, Shape shape) ResolveGroupInnerShapeBySegments(int slideIdx, string groupSegs, int shapeIdx)
    {
        var slideParts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > slideParts.Count)
            throw new ArgumentException($"Slide {slideIdx} not found (total: {slideParts.Count})");
        var slidePart = slideParts[PathIndex.ToArrayIndex(slideIdx)];
        var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree
            ?? throw new ArgumentException("Slide has no shape tree");
        OpenXmlCompositeElement current = shapeTree;
        int depth = 0;
        foreach (Match seg in Regex.Matches(groupSegs, @"/group\[(\d+)\]"))
        {
            depth++;
            var grpIdx = int.Parse(seg.Groups[1].Value);
            var groups = current.Elements<GroupShape>().ToList();
            if (grpIdx < 1 || grpIdx > groups.Count)
                throw new ArgumentException($"Group {grpIdx} not found at depth {depth} (total: {groups.Count})");
            current = groups[grpIdx - 1];
        }
        var innerShapes = current.Elements<Shape>().ToList();
        if (shapeIdx < 1 || shapeIdx > innerShapes.Count)
            throw new ArgumentException($"Shape {shapeIdx} not found in group (total: {innerShapes.Count})");
        return (slidePart, innerShapes[PathIndex.ToArrayIndex(shapeIdx)]);
    }

    /// <summary>
    /// /slide[N]/group[M](/group[L])*/picture[K] — resolve a Picture nested in a
    /// group at ANY depth and apply picture-set props via the shared core.
    /// Match group 2 is the full /group[..] chain. Mirrors
    /// ResolveGroupInnerShapeBySegments but for Picture children.
    /// </summary>
    private List<string> SetGroupInnerPictureByPath(Match m, Dictionary<string, string> properties)
    {
        var slideIdx = int.Parse(m.Groups[1].Value);
        var groupSegs = m.Groups[2].Value;
        var picIdx = int.Parse(m.Groups[3].Value);
        var slideParts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > slideParts.Count)
            throw new ArgumentException($"Slide {slideIdx} not found (total: {slideParts.Count})");
        var slidePart = slideParts[PathIndex.ToArrayIndex(slideIdx)];
        OpenXmlCompositeElement current = GetSlide(slidePart).CommonSlideData?.ShapeTree
            ?? throw new ArgumentException("Slide has no shape tree");
        int depth = 0;
        foreach (Match seg in Regex.Matches(groupSegs, @"/group\[(\d+)\]"))
        {
            depth++;
            var gi = int.Parse(seg.Groups[1].Value);
            var groups = current.Elements<GroupShape>().ToList();
            if (gi < 1 || gi > groups.Count)
                throw new ArgumentException($"Group {gi} not found at depth {depth} (total: {groups.Count})");
            current = groups[gi - 1];
        }
        var pics = current.Elements<Picture>().ToList();
        if (picIdx < 1 || picIdx > pics.Count)
            throw new ArgumentException($"Picture {picIdx} not found in group (total: {pics.Count})");
        return ApplyPicturePropertiesCore(slidePart, pics[picIdx - 1], properties);
    }

    private (SlidePart slidePart, Shape shape) ResolveGroupInnerShape(int slideIdx, int grpIdx, int shapeIdx)
    {
        var slideParts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > slideParts.Count)
            throw new ArgumentException($"Slide {slideIdx} not found (total: {slideParts.Count})");
        var slidePart = slideParts[PathIndex.ToArrayIndex(slideIdx)];
        var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree
            ?? throw new ArgumentException("Slide has no shape tree");
        var groups = shapeTree.Elements<GroupShape>().ToList();
        if (grpIdx < 1 || grpIdx > groups.Count)
            throw new ArgumentException($"Group {grpIdx} not found (total: {groups.Count})");
        var grp = groups[grpIdx - 1];
        var innerShapes = grp.Elements<Shape>().ToList();
        if (shapeIdx < 1 || shapeIdx > innerShapes.Count)
            throw new ArgumentException($"Shape {shapeIdx} not found in group {grpIdx} (total: {innerShapes.Count})");
        return (slidePart, innerShapes[PathIndex.ToArrayIndex(shapeIdx)]);
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

            var unsupported = SetRunOrShapeProperties(shapeProps, allRuns, shape, slidePart, unrecognizedLatex: LastUnrecognizedLatex);

            if (animValue != null)
            {
                // Replace, not accumulate — removal happens INSIDE
                // ApplyShapeAnimation after the effect name is validated, so a
                // bad effect can't strip the existing chain and leave empty
                // timing containers (CONSISTENCY(validate-before-mutate)).
                ApplyShapeAnimation(slidePart, shape, animValue, replaceExisting: true);
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

            // If this set moved or resized the shape, re-glue any connectors bound
            // to it so the line tracks the shape (PowerPoint interactive behavior).
            // The stored connector xfrm is authoritative at render time, so a stale
            // box would leave the connector detached from the shape it links.
            bool movedGeom = properties.Keys.Any(k =>
                k.Equals("x", StringComparison.OrdinalIgnoreCase) || k.Equals("left", StringComparison.OrdinalIgnoreCase)
                || k.Equals("y", StringComparison.OrdinalIgnoreCase) || k.Equals("top", StringComparison.OrdinalIgnoreCase)
                || k.Equals("width", StringComparison.OrdinalIgnoreCase) || k.Equals("height", StringComparison.OrdinalIgnoreCase));
            if (movedGeom
                && shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value is { } movedId)
                RerouteConnectorsForShape(slidePart, movedId);

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
        // New regex captures 4 groups (slide, kind, idx, animIdx); old 3-group
        // call sites still work because Groups[3] returns the empty group when
        // the regex doesn't capture it — but every live call site uses the
        // 4-group form now.
        var kindOrIdx = match.Groups[2].Value;
        SlidePart slidePart;
        OpenXmlElement targetEl;
        int elemIdx;
        int animIdx;
        bool isChart;
        if (kindOrIdx is "shape" or "chart")
        {
            isChart = kindOrIdx == "chart";
            elemIdx = int.Parse(match.Groups[3].Value);
            animIdx = int.Parse(match.Groups[4].Value);
            if (isChart)
            {
                var (sp, gf, _, _) = ResolveChart(slideIdx, elemIdx);
                slidePart = sp; targetEl = gf;
            }
            else
            {
                var (sp, sh) = ResolveShape(slideIdx, elemIdx);
                slidePart = sp; targetEl = sh;
            }
        }
        else
        {
            // Legacy 3-group capture (shape implicit) — kept for safety.
            isChart = false;
            elemIdx = int.Parse(kindOrIdx);
            animIdx = int.Parse(match.Groups[3].Value);
            var (sp, sh) = ResolveShape(slideIdx, elemIdx);
            slidePart = sp; targetEl = sh;
        }
        var ctns = EnumerateShapeAnimationCTns(slidePart, targetEl);
        if (animIdx < 1 || animIdx > ctns.Count)
            throw new ArgumentException(
                $"Animation {animIdx} not found on {(isChart ? "chart" : "shape")} {elemIdx} (total: {ctns.Count})");
        if (!isChart && (properties.ContainsKey("chartBuild") || properties.ContainsKey("chartbuild")))
            throw new ArgumentException(
                "chartBuild only applies to chart targets. Use /slide[N]/chart[M]/animation[K].");

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

        // For chart targets, seed every snapshot with the current chartBuild
        // value pulled from the slide's <p:bldGraphic>. chartBuild is chart-wide
        // (one bldGraphic per spid), so all snapshots share the same value;
        // user override on the target index propagates to every snapshot below
        // so the replay loop's last-write-wins lands on the user-intended value.
        string? currentChartBuild = null;
        if (isChart)
        {
            var spIdStr = GetAnimationTargetSpId(targetEl)?.ToString();
            if (spIdStr != null)
            {
                var bldGraphic = slidePart.Slide?.GetFirstChild<Timing>()?.BuildList?
                    .Elements<BuildGraphics>().FirstOrDefault(b => b.ShapeId?.Value == spIdStr);
                if (bldGraphic != null)
                    currentChartBuild = bldGraphic.BuildSubElement?.BuildChart?.Build?.Value ?? "asWhole";
            }
            if (currentChartBuild != null)
                foreach (var snap in snapshots) snap["chartBuild"] = currentChartBuild;
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
        // chartBuild is chart-wide; if the user overrode it on the target index,
        // validate and propagate to all snapshots so last-write-wins in the
        // replay loop reflects the user's choice (not the seeded old value).
        if (isChart && target.TryGetValue("chartBuild", out var tCb))
        {
            ValidateAnimationChartBuild(tCb);
            foreach (var snap in snapshots) snap["chartBuild"] = tCb;
        }

        // Wipe all animations on the shape, then re-apply each snapshot in order.
        // CONSISTENCY(animation-chain): motion-class snapshots route through
        // AppendMotionPathAnimation; preset (entrance/exit/emphasis) snapshots
        // route through ApplyShapeAnimation. Both append to the MainSequence
        // ChildTimeNodeList in original order so animation[K] indexing holds.
        // CONSISTENCY(validate-before-mutate): a bad edit (effect=badname) used
        // to surface only when re-applying snapshot K — AFTER the wipe below —
        // destroying every animation on the shape and leaving a half-built
        // timing tree. Pre-validate all non-motion snapshots' effect names.
        foreach (var preSnap in snapshots)
        {
            if (preSnap.TryGetValue("class", out var preCls)
                && preCls.Equals("motion", StringComparison.OrdinalIgnoreCase)) continue;
            ValidateAnimEffectName(BuildAnimValueFromProps(preSnap));
        }
        var shapeId = GetAnimationTargetSpId(targetEl);
        if (shapeId.HasValue)
        {
            RemoveShapeAnimations(slidePart.Slide!, shapeId.Value);
            // RemoveShapeAnimations targets MainSequence groups that contain
            // a matching ShapeTarget; motion-path groups land in the same list
            // so they're removed too. Belt-and-suspenders: also drop motion
            // path animations explicitly in case the writer changes.
            // (Charts don't carry motion-path animations — rejected on Add —
            // so this is a no-op for chart targets.)
            RemoveAllMotionPathAnimationsForShape(slidePart.Slide!, shapeId.Value);
        }
        for (int i = 0; i < snapshots.Count; i++)
        {
            var snap = snapshots[i];
            if (snap.TryGetValue("class", out var snapCls)
                && snapCls.Equals("motion", StringComparison.OrdinalIgnoreCase))
            {
                // Motion snapshots only occur on shape targets — chart Add
                // hard-rejects class=motion, so the snapshot can't carry it.
                ReapplyMotionFromSnapshot(slidePart, (Shape)targetEl, snap);
            }
            else
            {
                var animValue = BuildAnimValueFromProps(snap);
                ApplyShapeAnimation(slidePart, targetEl, animValue);
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
                    c.GetAttributes().Any(a => a.LocalName == "presetClass"
                        && a.Value is "path" or "motion")
                    && c.Descendants<AnimateMotion>().Any()))
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
        // chartBuild rides the same composite string so chart-target snapshots
        // re-emit the bldGraphic/bldChart wrapper on replay. Plain shape
        // snapshots never carry this key (Add hard-rejects it on shapes).
        if (p.TryGetValue("chartBuild", out var cbVal) && !string.IsNullOrEmpty(cbVal))
            animValue += $"-chartBuild={cbVal}";
        else if (p.TryGetValue("chartbuild", out var cbVal2) && !string.IsNullOrEmpty(cbVal2))
            animValue += $"-chartBuild={cbVal2}";
        return animValue;
    }
}
