// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler
{
    // ==================== Query Layer ====================

    public DocumentNode Get(string path, int depth = 1)
    {
        if (string.IsNullOrEmpty(path))
            throw new ArgumentException("Path cannot be empty.");
        path = NormalizePptxPathSegmentCasing(path);
        path = NormalizeCellPath(path);
        path = ResolveIdPath(path);
        path = ResolveLastPredicates(path);
        if (path == "/")
        {
            var node = new DocumentNode { Path = "/", Type = "presentation" };

            // Slide size
            var sldSz = _doc.PresentationPart?.Presentation?.GetFirstChild<SlideSize>();
            if (sldSz != null)
            {
                if (sldSz.Cx?.HasValue == true) node.Format["slideWidth"] = FormatEmu(sldSz.Cx.Value);
                if (sldSz.Cy?.HasValue == true) node.Format["slideHeight"] = FormatEmu(sldSz.Cy.Value);
                if (sldSz.Type is { HasValue: true } sldType) node.Format["slideSize"] = sldType.InnerText!.ToLowerInvariant() switch
                {
                    "screen16x9" => "widescreen",
                    "screen4x3" => "standard",
                    "screen16x10" => "16:10",
                    "a4" => "a4",
                    "a3" => "a3",
                    "letter" => "letter",
                    "b4iso" => "b4",
                    "b5iso" => "b5",
                    "35mm" => "35mm",
                    "overhead" => "overhead",
                    "banner" => "banner",
                    "ledger" => "ledger",
                    "custom" => "custom",
                    var other => other
                };
            }

            // Default font from theme
            var masterPart = _doc.PresentationPart?.SlideMasterParts?.FirstOrDefault();
            var fontScheme = masterPart?.ThemePart?.Theme?.ThemeElements?.FontScheme;
            if (fontScheme?.MinorFont?.LatinFont?.Typeface != null)
                node.Format["defaultFont"] = fontScheme.MinorFont.LatinFont.Typeface!.Value!;

            // Core document properties
            var props = _doc.PackageProperties;
            if (props.Title != null) node.Format["title"] = props.Title;
            if (props.Creator != null) node.Format["author"] = props.Creator;
            if (props.Subject != null) node.Format["subject"] = props.Subject;
            if (props.Keywords != null) node.Format["keywords"] = props.Keywords;
            if (props.Description != null) node.Format["description"] = props.Description;
            if (props.Category != null) node.Format["category"] = props.Category;
            if (props.LastModifiedBy != null) node.Format["lastModifiedBy"] = props.LastModifiedBy;
            if (props.Revision != null) node.Format["revision"] = props.Revision;
            if (props.Created != null) node.Format["created"] = props.Created.Value.ToString("o");
            if (props.Modified != null) node.Format["modified"] = props.Modified.Value.ToString("o");

            int slideNum = 0;
            foreach (var slidePart in GetSlideParts())
            {
                slideNum++;
                var title = GetSlide(slidePart).CommonSlideData?.ShapeTree?.Elements<Shape>()
                    .Where(IsTitle).Select(GetShapeText).FirstOrDefault() ?? "(untitled)";

                var slideNode = new DocumentNode
                {
                    Path = $"/slide[{slideNum}]",
                    Type = "slide",
                    Preview = title
                };
                var lName = GetSlideLayoutName(slidePart);
                if (lName != null) slideNode.Format["layout"] = lName;
                if (GetSlide(slidePart).Show?.Value == false)
                    slideNode.Format["hidden"] = true;
                ReadSlideBackground(GetSlide(slidePart), slideNode);
                ReadSlideTransition(slidePart, slideNode);

                if (depth > 0)
                {
                    slideNode.Children = GetSlideChildNodes(slidePart, slideNum, depth - 1);
                    slideNode.ChildCount = slideNode.Children.Count;
                }
                else
                {
                    var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree;
                    slideNode.ChildCount = (shapeTree?.Elements<Shape>().Count() ?? 0)
                        + (shapeTree?.Elements<Picture>().Count() ?? 0)
                        + (shapeTree?.Elements<GraphicFrame>().Count() ?? 0)
                        + (shapeTree?.Elements<ConnectionShape>().Count() ?? 0)
                        + (shapeTree?.Elements<GroupShape>().Count() ?? 0)
                        + (shapeTree != null ? GetZoomElements(shapeTree).Count : 0);
                }

                node.Children.Add(slideNode);
            }
            // Presentation-level settings
            PopulatePresentationSettings(node);
            Core.ThemeHandler.PopulateTheme(
                _doc.PresentationPart?.SlideMasterParts?.FirstOrDefault()?.ThemePart, node);
            Core.ExtendedPropertiesHandler.PopulateExtendedProperties(_doc.ExtendedFilePropertiesPart, node);

            node.ChildCount = node.Children.Count;
            return node;
        }

        if (path.Equals("/theme", StringComparison.OrdinalIgnoreCase))
            return GetThemeNode();

        if (path.Equals("/morph-check", StringComparison.OrdinalIgnoreCase))
            return GetMorphCheckNode();

        // Try slidemaster path: /slidemaster[N]
        var masterGetMatch = Regex.Match(path, @"^/slidemaster\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (masterGetMatch.Success)
        {
            var masterIdx = int.Parse(masterGetMatch.Groups[1].Value);
            var masters = _doc.PresentationPart?.SlideMasterParts?.ToList() ?? [];
            if (masterIdx < 1 || masterIdx > masters.Count)
                throw new ArgumentException($"Slide master {masterIdx} not found (total: {masters.Count})");
            var mp = masters[masterIdx - 1];
            var masterNode = new DocumentNode { Path = $"/slidemaster[{masterIdx}]", Type = "slidemaster" };
            var masterName = mp.SlideMaster?.CommonSlideData?.Name?.Value ?? "(unnamed)";
            masterNode.Preview = masterName;
            masterNode.Format["name"] = masterName;
            masterNode.Format["layoutCount"] = mp.SlideLayoutParts?.Count() ?? 0;
            var themePart = mp.ThemePart;
            if (themePart?.Theme?.Name?.Value != null)
                masterNode.Format["theme"] = themePart.Theme.Name.Value;
            var shapeTree = mp.SlideMaster?.CommonSlideData?.ShapeTree;
            var shapeCount = (shapeTree?.Elements<Shape>().Count() ?? 0)
                + (shapeTree?.Elements<Picture>().Count() ?? 0);
            masterNode.Format["shapeCount"] = shapeCount;
            ReadBackground(mp.SlideMaster?.CommonSlideData, masterNode);
            // CONSISTENCY(master-direction): Set persists rtl into the master's
            // <p:txStyles>/bodyStyle/lvl1pPr@rtl. Mirror it back on Get so users
            // can verify their own write (was previously set-only — Get omitted
            // the key entirely).
            var smTxStyles = mp.SlideMaster?.TextStyles;
            var rtlVal = smTxStyles?.GetFirstChild<BodyStyle>()?.GetFirstChild<Drawing.Level1ParagraphProperties>()?.RightToLeft?.Value
                ?? smTxStyles?.GetFirstChild<TitleStyle>()?.GetFirstChild<Drawing.Level1ParagraphProperties>()?.RightToLeft?.Value
                ?? smTxStyles?.GetFirstChild<OtherStyle>()?.GetFirstChild<Drawing.Level1ParagraphProperties>()?.RightToLeft?.Value;
            if (rtlVal == true)
                masterNode.Format["direction"] = "rtl";
            // Add layout children
            int lIdx = 0;
            foreach (var lp in mp.SlideLayoutParts ?? Enumerable.Empty<SlideLayoutPart>())
            {
                lIdx++;
                var lNode = new DocumentNode
                {
                    Path = $"/slidemaster[{masterIdx}]/slidelayout[{lIdx}]",
                    Type = "slidelayout",
                    Preview = lp.SlideLayout?.CommonSlideData?.Name?.Value ?? "(unnamed)"
                };
                lNode.Format["name"] = lNode.Preview;
                if (lp.SlideLayout?.Type?.HasValue == true)
                    lNode.Format["type"] = lp.SlideLayout.Type.InnerText;
                masterNode.Children.Add(lNode);
            }
            masterNode.ChildCount = masterNode.Children.Count;
            return masterNode;
        }

        // Try slidelayout path: /slidelayout[N] or /slidemaster[N]/slidelayout[M]
        var nestedLayoutMatch = Regex.Match(path, @"^/slidemaster\[(\d+)\]/slidelayout\[(\d+)\]$", RegexOptions.IgnoreCase);
        var layoutGetMatch = Regex.Match(path, @"^/slidelayout\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (nestedLayoutMatch.Success || layoutGetMatch.Success)
        {
            SlideLayoutPart lp;
            string resolvedPath;
            if (nestedLayoutMatch.Success)
            {
                var mIdx = int.Parse(nestedLayoutMatch.Groups[1].Value);
                var lIdx = int.Parse(nestedLayoutMatch.Groups[2].Value);
                var masters = _doc.PresentationPart?.SlideMasterParts?.ToList() ?? [];
                if (mIdx < 1 || mIdx > masters.Count)
                    throw new ArgumentException($"Slide master {mIdx} not found (total: {masters.Count})");
                var layouts = masters[mIdx - 1].SlideLayoutParts?.ToList() ?? [];
                if (lIdx < 1 || lIdx > layouts.Count)
                    throw new ArgumentException($"Slide layout {lIdx} not found under master {mIdx} (total: {layouts.Count})");
                lp = layouts[lIdx - 1];
                resolvedPath = $"/slidemaster[{mIdx}]/slidelayout[{lIdx}]";
            }
            else
            {
                var layoutIdx = int.Parse(layoutGetMatch.Groups[1].Value);
                var allLayouts = (_doc.PresentationPart?.SlideMasterParts ?? Enumerable.Empty<SlideMasterPart>())
                    .SelectMany(m => m.SlideLayoutParts ?? Enumerable.Empty<SlideLayoutPart>()).ToList();
                if (layoutIdx < 1 || layoutIdx > allLayouts.Count)
                    throw new ArgumentException($"Slide layout {layoutIdx} not found (total: {allLayouts.Count})");
                lp = allLayouts[layoutIdx - 1];
                resolvedPath = $"/slidelayout[{layoutIdx}]";
            }
            var layoutNode = new DocumentNode { Path = resolvedPath, Type = "slidelayout" };
            var layoutName = lp.SlideLayout?.CommonSlideData?.Name?.Value ?? "(unnamed)";
            layoutNode.Preview = layoutName;
            layoutNode.Format["name"] = layoutName;
            if (lp.SlideLayout?.Type?.HasValue == true)
                layoutNode.Format["type"] = lp.SlideLayout.Type.InnerText;
            ReadBackground(lp.SlideLayout?.CommonSlideData, layoutNode);

            // Populate child shapes — mirror what the slide Get branch does so
            // a layout-rooted Get exposes the same shape tree visible at
            // /slidelayout[N]/shape[K]. Previously childCount was always 0,
            // making layout edits non-discoverable via tree walk even though
            // the direct shape path worked.
            var layoutShapeTree = lp.SlideLayout?.CommonSlideData?.ShapeTree;
            if (layoutShapeTree != null)
            {
                int sIdx = 0;
                foreach (var sh in layoutShapeTree.Elements<Shape>())
                {
                    sIdx++;
                    var childNode = ShapeToNode(sh, slideNum: 0, sIdx,
                        depth: 0, part: null, parentPathPrefix: resolvedPath);
                    layoutNode.Children.Add(childNode);
                }
                layoutNode.ChildCount = layoutNode.Children.Count;
            }
            return layoutNode;
        }

        // CONSISTENCY(master-layout-shape-edit): Get on a master/layout shape path.
        // Add returns `/slidemaster[N]/shape[@id=K]` (and `/slidelayout[N]/shape[K]`,
        // `/slidemaster[N]/slidelayout[L]/shape[K]`); without this branch the path
        // fell through to the slide-only fallback and emitted the misleading
        // "Path must start with /slide[N], ..." error so Add output was
        // non-round-trippable.
        var nestedMasterShapeGetMatch = Regex.Match(path,
            @"^/slidemaster\[(\d+)\]/slidelayout\[(\d+)\]/shape\[(\d+)\]$", RegexOptions.IgnoreCase);
        var masterShapeGetMatch = Regex.Match(path,
            @"^/(slidemaster|slidelayout)\[(\d+)\]/shape\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (nestedMasterShapeGetMatch.Success || masterShapeGetMatch.Success)
        {
            ShapeTree? mlShapeTree;
            string mlPathPrefix;
            if (nestedMasterShapeGetMatch.Success)
            {
                var mIdx = int.Parse(nestedMasterShapeGetMatch.Groups[1].Value);
                var lIdx = int.Parse(nestedMasterShapeGetMatch.Groups[2].Value);
                var masters = _doc.PresentationPart?.SlideMasterParts?.ToList() ?? [];
                if (mIdx < 1 || mIdx > masters.Count)
                    throw new ArgumentException($"Slide master {mIdx} not found (total: {masters.Count})");
                var layouts = masters[mIdx - 1].SlideLayoutParts?.ToList() ?? [];
                if (lIdx < 1 || lIdx > layouts.Count)
                    throw new ArgumentException($"Slide layout {lIdx} not found under master {mIdx} (total: {layouts.Count})");
                mlShapeTree = layouts[lIdx - 1].SlideLayout?.CommonSlideData?.ShapeTree;
                mlPathPrefix = $"/slidemaster[{mIdx}]/slidelayout[{lIdx}]";
                var shapeIdx = int.Parse(nestedMasterShapeGetMatch.Groups[3].Value);
                return GetMasterOrLayoutShapeNode(mlShapeTree, shapeIdx, mlPathPrefix, depth);
            }
            else
            {
                var kind = masterShapeGetMatch.Groups[1].Value.ToLowerInvariant();
                var pIdx = int.Parse(masterShapeGetMatch.Groups[2].Value);
                var shapeIdx = int.Parse(masterShapeGetMatch.Groups[3].Value);
                if (kind == "slidemaster")
                {
                    var masters = _doc.PresentationPart?.SlideMasterParts?.ToList() ?? [];
                    if (pIdx < 1 || pIdx > masters.Count)
                        throw new ArgumentException($"Slide master {pIdx} not found (total: {masters.Count})");
                    mlShapeTree = masters[pIdx - 1].SlideMaster?.CommonSlideData?.ShapeTree;
                    mlPathPrefix = $"/slidemaster[{pIdx}]";
                }
                else
                {
                    var allLayouts = (_doc.PresentationPart?.SlideMasterParts ?? Enumerable.Empty<SlideMasterPart>())
                        .SelectMany(m => m.SlideLayoutParts ?? Enumerable.Empty<SlideLayoutPart>()).ToList();
                    if (pIdx < 1 || pIdx > allLayouts.Count)
                        throw new ArgumentException($"Slide layout {pIdx} not found (total: {allLayouts.Count})");
                    mlShapeTree = allLayouts[pIdx - 1].SlideLayout?.CommonSlideData?.ShapeTree;
                    mlPathPrefix = $"/slidelayout[{pIdx}]";
                }
                return GetMasterOrLayoutShapeNode(mlShapeTree, shapeIdx, mlPathPrefix, depth);
            }
        }

        // Try OLE path: /slide[N]/ole[M]
        // CONSISTENCY(ole-alias): "oleobject" mirrors Add's case switch
        var oleGetMatch = Regex.Match(path, @"^/slide\[(\d+)\]/(?:ole|oleobject|object|embed)\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (oleGetMatch.Success)
        {
            var oleSlideIdx = int.Parse(oleGetMatch.Groups[1].Value);
            var oleNodeIdx = int.Parse(oleGetMatch.Groups[2].Value);
            var slidePartsO = GetSlideParts().ToList();
            if (oleSlideIdx < 1 || oleSlideIdx > slidePartsO.Count)
                throw new ArgumentException($"Slide {oleSlideIdx} not found (total: {slidePartsO.Count})");
            var oleNodes = CollectOleNodesForSlide(oleSlideIdx, slidePartsO[oleSlideIdx - 1]);
            if (oleNodeIdx < 1 || oleNodeIdx > oleNodes.Count)
                throw new ArgumentException($"OLE object {oleNodeIdx} not found at /slide[{oleSlideIdx}] (available: {oleNodes.Count}).");
            return oleNodes[oleNodeIdx - 1];
        }

        // Modern p188 comment reply: /slide[N]/moderncomment[K]/reply[R].
        // (Top-level /slide[N]/moderncomment[K] is matched by the generic
        // /slide[N]/<type>[K] branch below via elementType == "moderncomment".)
        var mcReplyGetMatch = Regex.Match(path,
            @"^/slide\[(\d+)\]/moderncomment\[(\d+)\]/reply\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (mcReplyGetMatch.Success)
        {
            var rr = ResolveModernCommentReply(path)
                ?? throw new ArgumentException($"Modern comment reply not found: {path}");
            return ModernCommentReplyToNode(rr.slideIdx, rr.parentIdx, rr.reply, rr.replyIdx);
        }

        // Try notes path: /slide[N]/notes
        var notesGetMatch = Regex.Match(path, @"^/slide\[(\d+)\]/notes$");
        if (notesGetMatch.Success)
        {
            var notesSlideIdx = int.Parse(notesGetMatch.Groups[1].Value);
            var slidePartsN = GetSlideParts().ToList();
            if (notesSlideIdx < 1 || notesSlideIdx > slidePartsN.Count)
                throw new ArgumentException($"Slide {notesSlideIdx} not found (total: {slidePartsN.Count})");
            var slidePartN = slidePartsN[notesSlideIdx - 1];
            if (slidePartN.NotesSlidePart == null)
                // CONSISTENCY(not-found-uniformity): missing notes on a valid
                // slide is the same shape as out-of-range slide ("entity not
                // present at a valid path") — surface as not_found, not as
                // an error-typed DocumentNode (which the envelope formatter
                // coerces to internal_error).
                throw new ArgumentException($"Notes not found at /slide[{notesSlideIdx}]/notes (slide has no speaker notes)");
            var notesText = GetNotesText(slidePartN.NotesSlidePart);
            var notesNode = new DocumentNode { Path = path, Type = "notes", Text = notesText };
            // Schema declares text get=true; mirror node.Text into Format for parity.
            notesNode.Format["text"] = notesText ?? "";
            // Walk the notes body shape's first run to expose font formatting
            // (bold/italic/underline/color/size/font/lang/spacing/...). Without
            // this the Set→Get round-trip silently loses every formatting key
            // accepted by Set --prop. Mirrors the curated reader in RunToNode.
            PopulateNotesFormat(slidePartN.NotesSlidePart, notesNode);
            return notesNode;
        }

        // Try paragraph/run paths: /slide[N]/shape[M]/paragraph[P] or .../run[K] or .../paragraph[P]/run[K]
        // CONSISTENCY(path-aliases): see PowerPointHandler.Set.cs runMatch — PPT
        // accepts Word-style `/r[N]` / `/p[N]` short forms in addition to the
        // canonical `/run[N]` / `/paragraph[N]`.
        var runPathMatch = Regex.Match(path, @"^/slide\[(\d+)\]/shape\[(\d+)\]/(?:run|r)\[(\d+)\]$");
        if (runPathMatch.Success)
        {
            var sIdx = int.Parse(runPathMatch.Groups[1].Value);
            var shIdx = int.Parse(runPathMatch.Groups[2].Value);
            var rIdx = int.Parse(runPathMatch.Groups[3].Value);
            var (runSlidePart, shape) = ResolveShape(sIdx, shIdx);
            var shapePathSeg = BuildElementPathSegment("shape", shape, shIdx);
            var allRuns = GetAllRuns(shape);
            if (rIdx < 1 || rIdx > allRuns.Count)
                throw new ArgumentException($"Run {rIdx} not found (shape has {allRuns.Count} runs)");
            return RunToNode(allRuns[rIdx - 1], $"/slide[{sIdx}]/{shapePathSeg}/run[{rIdx}]", runSlidePart);
        }

        var paraPathMatch = Regex.Match(path, @"^/slide\[(\d+)\]/shape\[(\d+)\]/(?:paragraph|p)\[(\d+)\](?:/(?:run|r)\[(\d+)\])?$");
        if (paraPathMatch.Success)
        {
            var sIdx = int.Parse(paraPathMatch.Groups[1].Value);
            var shIdx = int.Parse(paraPathMatch.Groups[2].Value);
            var pIdx = int.Parse(paraPathMatch.Groups[3].Value);
            var (paraSlidePart, shape) = ResolveShape(sIdx, shIdx);
            var shapePathSeg = BuildElementPathSegment("shape", shape, shIdx);
            var paragraphs = shape.TextBody?.Elements<Drawing.Paragraph>().ToList()
                ?? throw new ArgumentException("Shape has no text body");
            if (pIdx < 1 || pIdx > paragraphs.Count)
                throw new ArgumentException($"Paragraph {pIdx} not found (shape has {paragraphs.Count} paragraphs)");

            var para = paragraphs[pIdx - 1];

            if (paraPathMatch.Groups[4].Success)
            {
                // /slide[N]/shape[@id=X]/paragraph[P]/run[K]
                var rIdx = int.Parse(paraPathMatch.Groups[4].Value);
                var paraRuns = para.Elements<Drawing.Run>().ToList();
                if (rIdx < 1 || rIdx > paraRuns.Count)
                    throw new ArgumentException($"Run {rIdx} not found (paragraph has {paraRuns.Count} runs)");
                return RunToNode(paraRuns[rIdx - 1],
                    $"/slide[{sIdx}]/{shapePathSeg}/paragraph[{pIdx}]/run[{rIdx}]", paraSlidePart);
            }

            // /slide[N]/shape[@id=X]/paragraph[P]
            var paraText = string.Join("", para.Elements<Drawing.Run>().Select(r => r.Text?.Text ?? ""));
            var paraNode = new DocumentNode
            {
                Path = $"/slide[{sIdx}]/{shapePathSeg}/paragraph[{pIdx}]",
                Type = "paragraph",
                Text = paraText
            };
            var qParaPProps = para.ParagraphProperties;
            if (qParaPProps?.Alignment?.HasValue == true) paraNode.Format["align"] = NormalizeAlignment(qParaPProps.Alignment.InnerText!);
            if (qParaPProps?.Level?.HasValue == true) paraNode.Format["level"] = qParaPProps.Level.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
            if (qParaPProps?.Indent?.HasValue == true) paraNode.Format["indent"] = FormatEmu(qParaPProps.Indent.Value);
            if (qParaPProps?.LeftMargin?.HasValue == true) paraNode.Format["marginLeft"] = FormatEmu(qParaPProps.LeftMargin.Value);
            if (qParaPProps?.RightMargin?.HasValue == true) paraNode.Format["marginRight"] = FormatEmu(qParaPProps.RightMargin.Value);
            var qLsPct = qParaPProps?.GetFirstChild<Drawing.LineSpacing>()?.GetFirstChild<Drawing.SpacingPercent>()?.Val?.Value;
            if (qLsPct.HasValue) paraNode.Format["lineSpacing"] = SpacingConverter.FormatPptLineSpacingPercent(qLsPct.Value);
            var qLsPts = qParaPProps?.GetFirstChild<Drawing.LineSpacing>()?.GetFirstChild<Drawing.SpacingPoints>()?.Val?.Value;
            if (qLsPts.HasValue) paraNode.Format["lineSpacing"] = SpacingConverter.FormatPptLineSpacingPoints(qLsPts.Value);
            var qSb = qParaPProps?.GetFirstChild<Drawing.SpaceBefore>()?.GetFirstChild<Drawing.SpacingPoints>()?.Val?.Value;
            if (qSb.HasValue) paraNode.Format["spaceBefore"] = SpacingConverter.FormatPptSpacing(qSb.Value);
            var qSa = qParaPProps?.GetFirstChild<Drawing.SpaceAfter>()?.GetFirstChild<Drawing.SpacingPoints>()?.Val?.Value;
            if (qSa.HasValue) paraNode.Format["spaceAfter"] = SpacingConverter.FormatPptSpacing(qSa.Value);
            // Reading direction (a:pPr rtl). Mirror NodeBuilder.ParaToNode so
            // direct paragraph Get matches shape-child-iteration Get.
            if (qParaPProps?.RightToLeft?.HasValue == true)
                paraNode.Format["direction"] = qParaPProps.RightToLeft.Value ? "rtl" : "ltr";

            var runs = para.Elements<Drawing.Run>().ToList();
            paraNode.ChildCount = runs.Count;
            if (depth > 0)
            {
                int runIdx = 0;
                foreach (var run in runs)
                {
                    paraNode.Children.Add(RunToNode(run,
                        $"/slide[{sIdx}]/{shapePathSeg}/paragraph[{pIdx}]/run[{runIdx + 1}]", paraSlidePart));
                    runIdx++;
                }
            }
            return paraNode;
        }

        // Try zoom path: /slide[N]/zoom[M]
        var zoomGetMatch = Regex.Match(path, @"^/slide\[(\d+)\]/zoom\[(\d+)\]$");
        if (zoomGetMatch.Success)
        {
            var sIdx = int.Parse(zoomGetMatch.Groups[1].Value);
            var zmIdx = int.Parse(zoomGetMatch.Groups[2].Value);
            var zmSlideParts = GetSlideParts().ToList();
            if (sIdx < 1 || sIdx > zmSlideParts.Count)
                throw new ArgumentException($"Slide {sIdx} not found (total: {zmSlideParts.Count})");
            var zmSlidePart = zmSlideParts[sIdx - 1];
            var zmShapeTree = GetSlide(zmSlidePart).CommonSlideData?.ShapeTree
                ?? throw new ArgumentException($"Slide {sIdx} has no shapes");
            var zoomElements = GetZoomElements(zmShapeTree);
            if (zmIdx < 1 || zmIdx > zoomElements.Count)
                throw new ArgumentException($"Zoom {zmIdx} not found (total: {zoomElements.Count})");
            return ZoomToNode(zoomElements[zmIdx - 1], sIdx, zmIdx);
        }

        // Try model3d path: /slide[N]/model3d[M]
        // model3d sits inside mc:AlternateContent, so the generic spTree
        // ChildElements fallback can't find it. Mirror the zoom branch.
        var m3dGetMatch = Regex.Match(path, @"^/slide\[(\d+)\]/model3d\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (m3dGetMatch.Success)
        {
            var sIdx = int.Parse(m3dGetMatch.Groups[1].Value);
            var mIdx = int.Parse(m3dGetMatch.Groups[2].Value);
            var m3dSlideParts = GetSlideParts().ToList();
            if (sIdx < 1 || sIdx > m3dSlideParts.Count)
                throw new ArgumentException($"Slide {sIdx} not found (total: {m3dSlideParts.Count})");
            var m3dSlidePart = m3dSlideParts[sIdx - 1];
            var m3dShapeTree = GetSlide(m3dSlidePart).CommonSlideData?.ShapeTree
                ?? throw new ArgumentException($"Slide {sIdx} has no shapes");
            var m3dElements = GetModel3DElements(m3dShapeTree);
            if (mIdx < 1 || mIdx > m3dElements.Count)
                throw new ArgumentException($"3D model {mIdx} not found at /slide[{sIdx}] (available: {m3dElements.Count}).");
            return Model3DToNode(m3dElements[mIdx - 1], sIdx, mIdx);
        }

        // Try animation path: /slide[N]/shape[M]/animation[A]
        var animPathMatch = Regex.Match(path, @"^/slide\[(\d+)\]/shape\[(\d+)\]/animation\[(\d+)\]$");
        if (animPathMatch.Success)
        {
            var sIdx = int.Parse(animPathMatch.Groups[1].Value);
            var shIdx = int.Parse(animPathMatch.Groups[2].Value);
            var aIdx = int.Parse(animPathMatch.Groups[3].Value);
            var (animSlidePart, animShape) = ResolveShape(sIdx, shIdx);
            var animShapePathSeg = BuildElementPathSegment("shape", animShape, shIdx);

            var effectCTns = EnumerateShapeAnimationCTns(animSlidePart, animShape);
            if (aIdx < 1 || aIdx > effectCTns.Count)
                return new DocumentNode { Path = path, Type = "error", Text = $"animation[{aIdx}] not found (shape has {effectCTns.Count} animation(s))" };
            var animNode = new DocumentNode { Path = $"/slide[{sIdx}]/{animShapePathSeg}/animation[{aIdx}]", Type = "animation" };
            PopulateAnimationNode(animNode, effectCTns[aIdx - 1]);
            return animNode;
        }

        // Try table cell path: /slide[N]/table[M]/tr[R]/tc[C]
        var tblCellGetMatch = Regex.Match(path, @"^/slide\[(\d+)\]/table\[(\d+)\]/tr\[(\d+)\]/tc\[(\d+)\]$");
        if (tblCellGetMatch.Success)
        {
            var sIdx = int.Parse(tblCellGetMatch.Groups[1].Value);
            var tIdx = int.Parse(tblCellGetMatch.Groups[2].Value);
            var rIdx = int.Parse(tblCellGetMatch.Groups[3].Value);
            var cIdx = int.Parse(tblCellGetMatch.Groups[4].Value);

            var (slidePart2, table) = ResolveTable(sIdx, tIdx);
            var tblGf = table.Ancestors<GraphicFrame>().FirstOrDefault();
            var tblPathSeg = tblGf != null ? BuildElementPathSegment("table", tblGf, tIdx) : $"table[{tIdx}]";
            var tableRows = table.Elements<Drawing.TableRow>().ToList();
            if (rIdx < 1 || rIdx > tableRows.Count)
                throw new ArgumentException($"Row {rIdx} not found (table has {tableRows.Count} rows)");
            var cells = tableRows[rIdx - 1].Elements<Drawing.TableCell>().ToList();
            if (cIdx < 1 || cIdx > cells.Count)
                throw new ArgumentException($"Cell {cIdx} not found (row has {cells.Count} cells)");

            var cell = cells[cIdx - 1];
            var cellText = GetCellTextWithParagraphBreaks(cell);
            var cellNode = new DocumentNode
            {
                Path = $"/slide[{sIdx}]/{tblPathSeg}/tr[{rIdx}]/tc[{cIdx}]",
                Type = "tc",
                Text = cellText
            };

            // BUG-R4-07: emit canonical 'colspan'/'rowspan' (matches docx),
            // not OOXML-internal 'gridSpan'/'rowSpan'. Set still accepts the
            // OOXML-internal aliases.
            if (cell.GridSpan?.HasValue == true && cell.GridSpan.Value > 1)
                cellNode.Format["colspan"] = cell.GridSpan.Value;
            if (cell.RowSpan?.HasValue == true && cell.RowSpan.Value > 1)
                cellNode.Format["rowspan"] = cell.RowSpan.Value;
            if (cell.HorizontalMerge?.HasValue == true && cell.HorizontalMerge.Value)
                cellNode.Format["hmerge"] = true;
            if (cell.VerticalMerge?.HasValue == true && cell.VerticalMerge.Value)
                cellNode.Format["vmerge"] = true;

            // Cell fill (blip, gradient, or solid)
            var tcPr = cell.TableCellProperties ?? cell.GetFirstChild<Drawing.TableCellProperties>();
            var cellBlipFill = tcPr?.GetFirstChild<Drawing.BlipFill>();
            if (cellBlipFill != null)
            {
                var blipEmbed = cellBlipFill.GetFirstChild<Drawing.Blip>()?.Embed?.Value;
                cellNode.Format["fill"] = "image";
                if (blipEmbed != null) cellNode.Format["image.relId"] = blipEmbed;
            }
            else if (tcPr?.GetFirstChild<Drawing.GradientFill>() is { } gradFill)
            {
                // BUG-R6-A: emit canonical fill="gradient" + Format["gradient"]=detail
                // (matches NodeBuilder cell path — was inconsistent before).
                cellNode.Format["fill"] = "gradient";
                cellNode.Format["gradient"] = ReadGradientString(gradFill);
            }
            else
            {
                // BUG-R6-A: read scheme color in addition to RgbColorModelHex.
                var cellFillSolid = tcPr?.GetFirstChild<Drawing.SolidFill>();
                var cellFillColor = ReadColorFromFill(cellFillSolid);
                if (cellFillColor != null) cellNode.Format["fill"] = cellFillColor;
            }

            // Cell borders — following POI's getBorderWidth/getBorderColor pattern
            if (tcPr != null)
                ReadTableCellBorders(tcPr, cellNode);

            // Vertical alignment
            if (tcPr?.Anchor?.HasValue == true)
            {
                cellNode.Format["valign"] = tcPr.Anchor.InnerText switch
                {
                    "ctr" => "center",
                    "t" => "top",
                    "b" => "bottom",
                    _ => tcPr.Anchor.InnerText
                };
            }

            // Cell text direction (a:tcPr @vert). Canonical readback mirrors the
            // Set vocabulary (horizontal / vertical270 / vertical90 / stacked).
            if (tcPr?.Vertical?.HasValue == true)
            {
                cellNode.Format["textdirection"] = tcPr.Vertical.InnerText switch
                {
                    "horz" => "horizontal",
                    "vert" => "vertical90",
                    "vert270" => "vertical270",
                    "wordArtVert" => "stacked",
                    _ => tcPr.Vertical.InnerText
                };
            }

            // Cell text wrap (a:tcPr/a:txBody/a:bodyPr @wrap). Set writes
            // square|none on the cell's BodyProperties; mirror back as bool.
            var qCellBodyPr = cell.TextBody?.GetFirstChild<Drawing.BodyProperties>();
            if (qCellBodyPr?.Wrap?.HasValue == true)
            {
                cellNode.Format["wrap"] = qCellBodyPr.Wrap.Value != Drawing.TextWrappingValues.None;
            }

            // BUG-R4-D9: padding.* readback (Set already wrote LeftMargin/etc;
            // Get was missing). Use FormatEmu to mirror cross-handler width/EMU
            // value formatting (e.g. "0.13cm").
            if (tcPr?.LeftMargin?.HasValue == true)
                cellNode.Format["padding.left"] = FormatEmu(tcPr.LeftMargin.Value);
            if (tcPr?.RightMargin?.HasValue == true)
                cellNode.Format["padding.right"] = FormatEmu(tcPr.RightMargin.Value);
            if (tcPr?.TopMargin?.HasValue == true)
                cellNode.Format["padding.top"] = FormatEmu(tcPr.TopMargin.Value);
            if (tcPr?.BottomMargin?.HasValue == true)
                cellNode.Format["padding.bottom"] = FormatEmu(tcPr.BottomMargin.Value);

            // Alignment from first paragraph
            var cellFirstPara = cell.TextBody?.Elements<Drawing.Paragraph>().FirstOrDefault();
            var cellParaAlign = cellFirstPara?.ParagraphProperties?.Alignment;
            if (cellParaAlign?.HasValue == true)
            {
                var align = NormalizeAlignment(cellParaAlign.InnerText!);
                // CONSISTENCY(canonical-format-keys): PPT canonical key for text
                // alignment is "align" (not "alignment"). Do not emit both.
                cellNode.Format["align"] = align;
            }

            // Direction from first paragraph (mirrors shape/textbox readback).
            // ltr is the schema default — only emit when explicitly set.
            if (cellFirstPara?.ParagraphProperties?.RightToLeft?.HasValue == true)
                cellNode.Format["direction"] = cellFirstPara.ParagraphProperties.RightToLeft.Value ? "rtl" : "ltr";

            // BUG-R6-A: cell-level lineSpacing/spaceBefore/spaceAfter readback
            // from first paragraph (Set writes to all paragraphs in cell;
            // Get returns the first one's value, mirroring shape paragraph aggregation).
            var qCellFirstPProps = cellFirstPara?.ParagraphProperties;
            if (qCellFirstPProps != null)
            {
                var qLsPct = qCellFirstPProps.GetFirstChild<Drawing.LineSpacing>()?.GetFirstChild<Drawing.SpacingPercent>()?.Val?.Value;
                if (qLsPct.HasValue) cellNode.Format["lineSpacing"] = OfficeCli.Core.SpacingConverter.FormatPptLineSpacingPercent(qLsPct.Value);
                var qLsPts = qCellFirstPProps.GetFirstChild<Drawing.LineSpacing>()?.GetFirstChild<Drawing.SpacingPoints>()?.Val?.Value;
                if (qLsPts.HasValue) cellNode.Format["lineSpacing"] = OfficeCli.Core.SpacingConverter.FormatPptLineSpacingPoints(qLsPts.Value);
                var qSb = qCellFirstPProps.GetFirstChild<Drawing.SpaceBefore>()?.GetFirstChild<Drawing.SpacingPoints>()?.Val?.Value;
                if (qSb.HasValue) cellNode.Format["spaceBefore"] = OfficeCli.Core.SpacingConverter.FormatPptSpacing(qSb.Value);
                var qSa = qCellFirstPProps.GetFirstChild<Drawing.SpaceAfter>()?.GetFirstChild<Drawing.SpacingPoints>()?.Val?.Value;
                if (qSa.HasValue) cellNode.Format["spaceAfter"] = OfficeCli.Core.SpacingConverter.FormatPptSpacing(qSa.Value);
            }

            // Font info from first run
            var firstRun = cell.Descendants<Drawing.Run>().FirstOrDefault();
            if (firstRun?.RunProperties != null)
            {
                var f = firstRun.RunProperties.GetFirstChild<Drawing.LatinFont>()?.Typeface?.Value
                    ?? firstRun.RunProperties.GetFirstChild<Drawing.EastAsianFont>()?.Typeface?.Value;
                if (f != null) cellNode.Format["font"] = f;
                var fs = firstRun.RunProperties.FontSize?.Value;
                if (fs.HasValue) cellNode.Format["size"] = $"{fs.Value / 100.0:0.##}pt";
                if (firstRun.RunProperties.Bold?.Value == true) cellNode.Format["bold"] = true;
                if (firstRun.RunProperties.Italic?.Value == true) cellNode.Format["italic"] = true;
                if (firstRun.RunProperties.Underline?.HasValue == true && firstRun.RunProperties.Underline.Value != Drawing.TextUnderlineValues.None)
                {
                    cellNode.Format["underline"] = firstRun.RunProperties.Underline.InnerText switch
                    {
                        "sng" => "single",
                        "dbl" => "double",
                        _ => firstRun.RunProperties.Underline.InnerText
                    };
                }
                if (firstRun.RunProperties.Strike?.HasValue == true)
                {
                    cellNode.Format["strike"] = firstRun.RunProperties.Strike.Value switch
                    {
                        var v when v == Drawing.TextStrikeValues.DoubleStrike => "double",
                        var v when v == Drawing.TextStrikeValues.NoStrike => "none",
                        _ => "single",
                    };
                }
                var colorHex = firstRun.RunProperties.GetFirstChild<Drawing.SolidFill>()
                    ?.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value;
                if (colorHex != null) cellNode.Format["color"] = ParseHelpers.FormatHexColor(colorHex);
            }

            return cellNode;
        }

        // Try placeholder path with type name: /slide[N]/placeholder[title]
        var phGetMatch = Regex.Match(path, @"^/slide\[(\d+)\]/placeholder\[(\w+)\]$");
        if (phGetMatch.Success && !Regex.IsMatch(path, @"^/slide\[\d+\](?:/\w+\[\d+\])?$"))
        {
            var phSlideIdx = int.Parse(phGetMatch.Groups[1].Value);
            var phId = phGetMatch.Groups[2].Value;

            var phSlideParts = GetSlideParts().ToList();
            if (phSlideIdx < 1 || phSlideIdx > phSlideParts.Count)
                throw new ArgumentException($"Slide {phSlideIdx} not found (total: {phSlideParts.Count})");

            var phSlidePart = phSlideParts[phSlideIdx - 1];

            // If numeric, delegate to GetPlaceholderNode
            if (int.TryParse(phId, out var phNumIdx))
                return GetPlaceholderNode(phSlidePart, phSlideIdx, phNumIdx, depth);

            // By type name: resolve the shape and return its node
            var phShape = ResolvePlaceholderShape(phSlidePart, phId);
            var ph = phShape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
                ?.GetFirstChild<PlaceholderShape>();
            var shapeTree = GetSlide(phSlidePart).CommonSlideData?.ShapeTree;
            var shapeIdx = shapeTree?.Elements<Shape>().ToList().IndexOf(phShape) ?? 0;
            var node = ShapeToNode(phShape, phSlideIdx, shapeIdx + 1, depth, phSlidePart);
            node.Path = path;
            node.Type = "placeholder";
            if (ph?.Type?.HasValue == true) node.Format["phType"] = ph.Type.InnerText;
            if (ph?.Index?.HasValue == true) node.Format["phIndex"] = ph.Index.Value;
            return node;
        }

        // Handle table sub-paths: /slide[N]/table[M]/tr[R] or /slide[N]/table[M]/tr[R]/tc[C]
        // Must come before generic XML fallback to use proper Format keys and unit formatting
        var tableSubMatch = Regex.Match(path, @"^/slide\[(\d+)\]/table\[(\d+)\]/(\w+)\[(\d+)\](?:/(\w+)\[(\d+)\])?$");
        if (tableSubMatch.Success)
        {
            var tSlideIdx = int.Parse(tableSubMatch.Groups[1].Value);
            var tTableIdx = int.Parse(tableSubMatch.Groups[2].Value);
            var tSubType = tableSubMatch.Groups[3].Value;  // "tr"
            var tSubIdx = int.Parse(tableSubMatch.Groups[4].Value);

            var tSlideParts = GetSlideParts().ToList();
            if (tSlideIdx < 1 || tSlideIdx > tSlideParts.Count)
                throw new ArgumentException($"Slide {tSlideIdx} not found (total: {tSlideParts.Count})");

            var tShapeTree = GetSlide(tSlideParts[tSlideIdx - 1]).CommonSlideData?.ShapeTree;
            if (tShapeTree == null) throw new ArgumentException($"Slide {tSlideIdx} has no shapes");

            var tTables = tShapeTree.Elements<GraphicFrame>()
                .Where(gf => gf.Descendants<Drawing.Table>().Any()).ToList();
            if (tTableIdx < 1 || tTableIdx > tTables.Count)
                throw new ArgumentException($"Table {tTableIdx} not found (total: {tTables.Count})");

            // Build table node with sufficient depth to include rows and cells
            var tableNode = TableToNode(tTables[tTableIdx - 1], tSlideIdx, tTableIdx, 2);

            // Find the row
            if (tSubType.Equals("tr", StringComparison.OrdinalIgnoreCase))
            {
                var rowNodes = tableNode.Children.Where(c => c.Type == "tr").ToList();
                if (tSubIdx < 1 || tSubIdx > rowNodes.Count)
                    throw new ArgumentException($"Row {tSubIdx} not found (total: {rowNodes.Count})");
                var rowNode = rowNodes[tSubIdx - 1];

                // If there's a further sub-path (e.g., /tc[C])
                if (tableSubMatch.Groups[5].Success)
                {
                    var tcType = tableSubMatch.Groups[5].Value;  // "tc"
                    var tcIdx = int.Parse(tableSubMatch.Groups[6].Value);
                    if (tcType.Equals("tc", StringComparison.OrdinalIgnoreCase))
                    {
                        var cellNodes = rowNode.Children.Where(c => c.Type == "tc").ToList();
                        if (tcIdx < 1 || tcIdx > cellNodes.Count)
                            throw new ArgumentException($"Cell {tcIdx} not found (total: {cellNodes.Count})");
                        return cellNodes[tcIdx - 1];
                    }
                }

                return rowNode;
            }

            // CONSISTENCY(table-col-get): mirror xlsx `get col[A]` — pptx
            // GridColumn carries Width directly, surface it as a unit-qualified
            // length. Schema: schemas/help/pptx/table-column.json declares
            // get: true; this implements it (was previously throwing).
            if (tSubType.Equals("col", StringComparison.OrdinalIgnoreCase))
            {
                var tbl = tTables[tTableIdx - 1].Descendants<Drawing.Table>().First();
                var gridCols = tbl.TableGrid?.Elements<Drawing.GridColumn>().ToList()
                    ?? new List<Drawing.GridColumn>();
                if (tSubIdx < 1 || tSubIdx > gridCols.Count)
                    throw new ArgumentException($"Column {tSubIdx} not found (total: {gridCols.Count})");
                var colNode = new DocumentNode { Path = path, Type = "col" };
                var gc = gridCols[tSubIdx - 1];
                if (gc.Width?.HasValue == true)
                    colNode.Format["width"] = FormatEmu(gc.Width.Value);
                return colNode;
            }

            throw new ArgumentException($"Unknown table sub-element: {tSubType}");
        }

        // Try chart series sub-path: /slide[N]/chart[M]/series[K]
        var chartSeriesGetMatch = Regex.Match(path, @"^/slide\[(\d+)\]/chart\[(\d+)\]/series\[(\d+)\]$");
        if (chartSeriesGetMatch.Success)
        {
            var csSlideIdx = int.Parse(chartSeriesGetMatch.Groups[1].Value);
            var csChartIdx = int.Parse(chartSeriesGetMatch.Groups[2].Value);
            var csSeriesIdx = int.Parse(chartSeriesGetMatch.Groups[3].Value);

            var (csSlidePart, csChartGf, csChartPart, _) = ResolveChart(csSlideIdx, csChartIdx);
            // Get the chart node with depth=1 to populate series children
            var chartNode = ChartToNode(csChartGf, csSlidePart, csSlideIdx, csChartIdx, 1);
            var seriesChildren = chartNode.Children.Where(c => c.Type == "series").ToList();
            if (csSeriesIdx < 1 || csSeriesIdx > seriesChildren.Count)
                throw new ArgumentException($"Series {csSeriesIdx} not found (total: {seriesChildren.Count})");
            var seriesNode = seriesChildren[csSeriesIdx - 1];
            seriesNode.Path = path;
            return seriesNode;
        }

        // Try chart axis-by-role sub-path: /slide[N]/chart[M]/axis[@role=ROLE]
        // Per schemas/help/pptx/chart-axis.json.
        var chartAxisGetMatch = Regex.Match(path,
            @"^/slide\[(\d+)\]/chart\[(\d+)\]/axis\[@role=([a-zA-Z0-9_]+)\]$");
        if (chartAxisGetMatch.Success)
        {
            var caSlideIdx = int.Parse(chartAxisGetMatch.Groups[1].Value);
            var caChartIdx = int.Parse(chartAxisGetMatch.Groups[2].Value);
            var caRole = chartAxisGetMatch.Groups[3].Value;

            var (_, _, caChartPart, _) = ResolveChart(caSlideIdx, caChartIdx);
            if (caChartPart?.ChartSpace == null)
                throw new ArgumentException($"Axis not found on chart {caChartIdx}: extended charts not supported.");
            var axisNode = Core.ChartHelper.BuildAxisNode(caChartPart.ChartSpace, caRole, path);
            if (axisNode == null)
                throw new ArgumentException($"Axis with role '{caRole}' not found on chart {caChartIdx}.");
            return axisNode;
        }

        // Try resolving logical paths with deeper segments (e.g. /slide[1]/placeholder[1]/...)
        // Only for paths not handled by dedicated handlers above
        if (Regex.IsMatch(path, @"^/slide\[\d+\]/placeholder\[\w+\]/"))
        {
            var logicalResolved = ResolveLogicalPath(path);
            if (logicalResolved.HasValue)
                return GenericXmlQuery.ElementToNode(logicalResolved.Value.element, path, depth);
        }

        // Try arbitrary-depth group descent: /slide[N]/group[M](/group[L])*/<leaf>[K]
        // CONSISTENCY(group-inner-shape): Get must traverse nested groups the
        // same way Query already does. Without this branch, paths like
        // /slide[1]/group[1]/group[1] or /slide[1]/group[1]/group[2]/shape[3]
        // fall through to the generic XML fallback, which mis-detects
        // GroupShape (LocalName="grpSp") and emits "Element not found".
        // Leaf may be a nested group itself or any non-group inner type (shape,
        // picture, table, chart, connector). Leaf-of-type-group returns the
        // group node; other leaves delegate to the matching ToNode builder so
        // the returned DocumentNode carries the full Format payload.
        var nestedGroupMatch = Regex.Match(path,
            @"^/slide\[(\d+)\]/group\[(\d+)\]((?:/group\[\d+\])*)(?:/(shape|picture|pic|table|chart|connector|connection)\[(\d+)\])?$");
        if (nestedGroupMatch.Success && (nestedGroupMatch.Groups[3].Length > 0 || nestedGroupMatch.Groups[4].Success))
        {
            var ngSlideIdx = int.Parse(nestedGroupMatch.Groups[1].Value);
            var ngRootGrpIdx = int.Parse(nestedGroupMatch.Groups[2].Value);
            var ngSlideParts = GetSlideParts().ToList();
            if (ngSlideIdx < 1 || ngSlideIdx > ngSlideParts.Count)
                throw new ArgumentException($"Slide {ngSlideIdx} not found (total: {ngSlideParts.Count})");
            var ngSlidePart = ngSlideParts[ngSlideIdx - 1];
            var ngShapeTree = GetSlide(ngSlidePart).CommonSlideData?.ShapeTree
                ?? throw new ArgumentException("Slide has no shape tree");
            var ngRootGroups = ngShapeTree.Elements<GroupShape>().ToList();
            if (ngRootGrpIdx < 1 || ngRootGrpIdx > ngRootGroups.Count)
                throw new ArgumentException($"Group {ngRootGrpIdx} not found (total: {ngRootGroups.Count})");
            GroupShape ngCurrent = ngRootGroups[ngRootGrpIdx - 1];
            var ngPathPrefix = $"/slide[{ngSlideIdx}]/{BuildElementPathSegment("group", ngCurrent, ngRootGrpIdx)}";
            // Walk nested /group[L] segments
            foreach (Match seg in Regex.Matches(nestedGroupMatch.Groups[3].Value, @"/group\[(\d+)\]"))
            {
                var subIdx = int.Parse(seg.Groups[1].Value);
                var subs = ngCurrent.Elements<GroupShape>().ToList();
                if (subIdx < 1 || subIdx > subs.Count)
                    throw new ArgumentException($"Nested group {subIdx} not found at {ngPathPrefix} (total: {subs.Count})");
                ngCurrent = subs[subIdx - 1];
                ngPathPrefix = $"{ngPathPrefix}/{BuildElementPathSegment("group", ngCurrent, subIdx)}";
            }
            // No leaf -> return the (possibly nested) group itself via the
            // shared NodeBuilder so children/Format stay in sync with the
            // top-level group branch.
            if (!nestedGroupMatch.Groups[4].Success)
            {
                var ngGrpName = ngCurrent.NonVisualGroupShapeProperties?.NonVisualDrawingProperties?.Name?.Value ?? "Group";
                var ngGrpNode = new DocumentNode
                {
                    Path = ngPathPrefix,
                    Type = "group",
                    Preview = ngGrpName,
                    ChildCount = ngCurrent.Elements<Shape>().Count() + ngCurrent.Elements<Picture>().Count()
                        + ngCurrent.Elements<GraphicFrame>().Count() + ngCurrent.Elements<ConnectionShape>().Count()
                        + ngCurrent.Elements<GroupShape>().Count()
                };
                ngGrpNode.Format["name"] = ngGrpName;
                var ngXfrm = ngCurrent.GroupShapeProperties?.TransformGroup;
                if (ngXfrm?.Offset?.X != null) ngGrpNode.Format["x"] = FormatEmu(ngXfrm.Offset.X.Value);
                if (ngXfrm?.Offset?.Y != null) ngGrpNode.Format["y"] = FormatEmu(ngXfrm.Offset.Y.Value);
                if (ngXfrm?.Extents?.Cx != null) ngGrpNode.Format["width"] = FormatEmu(ngXfrm.Extents.Cx.Value);
                if (ngXfrm?.Extents?.Cy != null) ngGrpNode.Format["height"] = FormatEmu(ngXfrm.Extents.Cy.Value);
                if (ngXfrm?.Rotation != null && ngXfrm.Rotation.Value != 0)
                    ngGrpNode.Format["rotation"] = $"{ngXfrm.Rotation.Value / 60000.0:0.##}";
                if (depth > 0)
                    BuildChildNodesIntoContainer(
                        ngGrpNode.Children, ngCurrent, ngSlidePart, ngSlideIdx, depth - 1,
                        ngPathPrefix, isSlideRoot: false);
                return ngGrpNode;
            }
            // Leaf is a non-group inner type
            var ngLeafType = nestedGroupMatch.Groups[4].Value.ToLowerInvariant();
            var ngLeafIdx = int.Parse(nestedGroupMatch.Groups[5].Value);
            switch (ngLeafType)
            {
                case "shape":
                {
                    var inner = ngCurrent.Elements<Shape>().ToList();
                    if (ngLeafIdx < 1 || ngLeafIdx > inner.Count)
                        throw new ArgumentException($"Shape {ngLeafIdx} not found in group {ngPathPrefix} (total: {inner.Count})");
                    var node = ShapeToNode(inner[ngLeafIdx - 1], ngSlideIdx, ngLeafIdx, depth, ngSlidePart, ngPathPrefix);
                    node.Path = $"{ngPathPrefix}/{BuildElementPathSegment("shape", inner[ngLeafIdx - 1], ngLeafIdx)}";
                    return node;
                }
                case "picture":
                case "pic":
                {
                    var inner = ngCurrent.Elements<Picture>().ToList();
                    if (ngLeafIdx < 1 || ngLeafIdx > inner.Count)
                        throw new ArgumentException($"Picture {ngLeafIdx} not found in group {ngPathPrefix} (total: {inner.Count})");
                    var node = PictureToNode(inner[ngLeafIdx - 1], ngSlideIdx, ngLeafIdx, ngSlidePart, ngPathPrefix);
                    node.Path = $"{ngPathPrefix}/{BuildElementPathSegment("picture", inner[ngLeafIdx - 1], ngLeafIdx)}";
                    return node;
                }
                case "connector":
                case "connection":
                {
                    var inner = ngCurrent.Elements<ConnectionShape>().ToList();
                    if (ngLeafIdx < 1 || ngLeafIdx > inner.Count)
                        throw new ArgumentException($"Connector {ngLeafIdx} not found in group {ngPathPrefix} (total: {inner.Count})");
                    return ConnectorToNode(inner[ngLeafIdx - 1], ngSlideIdx, ngLeafIdx, ngPathPrefix);
                }
                case "table":
                {
                    var inner = ngCurrent.Elements<GraphicFrame>()
                        .Where(gf => gf.Descendants<Drawing.Table>().Any()).ToList();
                    if (ngLeafIdx < 1 || ngLeafIdx > inner.Count)
                        throw new ArgumentException($"Table {ngLeafIdx} not found in group {ngPathPrefix} (total: {inner.Count})");
                    return TableToNode(inner[ngLeafIdx - 1], ngSlideIdx, ngLeafIdx, depth, ngPathPrefix);
                }
                case "chart":
                {
                    var inner = ngCurrent.Elements<GraphicFrame>()
                        .Where(gf => gf.Descendants<C.ChartReference>().Any() || IsExtendedChartFrame(gf)).ToList();
                    if (ngLeafIdx < 1 || ngLeafIdx > inner.Count)
                        throw new ArgumentException($"Chart {ngLeafIdx} not found in group {ngPathPrefix} (total: {inner.Count})");
                    return ChartToNode(inner[ngLeafIdx - 1], ngSlidePart, ngSlideIdx, ngLeafIdx, depth, ngPathPrefix);
                }
            }
        }

        // Try group inner shape path: /slide[N]/group[M]/shape[K]
        // CONSISTENCY(group-inner-shape): Set supports this; Get must too.
        // Previously fell through to the generic XML fallback, which mis-detected
        // GroupShape (LocalName="grpSp") as a shape and threw "No shape found".
        var grpInnerGetMatch = Regex.Match(path, @"^/slide\[(\d+)\]/group\[(\d+)\]/shape\[(\d+)\]$");
        if (grpInnerGetMatch.Success)
        {
            var giSlideIdx = int.Parse(grpInnerGetMatch.Groups[1].Value);
            var giGrpIdx = int.Parse(grpInnerGetMatch.Groups[2].Value);
            var giShapeIdx = int.Parse(grpInnerGetMatch.Groups[3].Value);
            var giSlideParts = GetSlideParts().ToList();
            if (giSlideIdx < 1 || giSlideIdx > giSlideParts.Count)
                throw new ArgumentException($"Slide {giSlideIdx} not found (total: {giSlideParts.Count})");
            var giSlidePart = giSlideParts[giSlideIdx - 1];
            var giShapeTree = GetSlide(giSlidePart).CommonSlideData?.ShapeTree
                ?? throw new ArgumentException("Slide has no shape tree");
            var giGroups = giShapeTree.Elements<GroupShape>().ToList();
            if (giGrpIdx < 1 || giGrpIdx > giGroups.Count)
                throw new ArgumentException($"Group {giGrpIdx} not found (total: {giGroups.Count})");
            var giInnerShapes = giGroups[giGrpIdx - 1].Elements<Shape>().ToList();
            if (giShapeIdx < 1 || giShapeIdx > giInnerShapes.Count)
                throw new ArgumentException($"Shape {giShapeIdx} not found in group {giGrpIdx} (total: {giInnerShapes.Count})");
            var giParentPrefix = $"/slide[{giSlideIdx}]/group[{giGrpIdx}]";
            var giNode = ShapeToNode(giInnerShapes[giShapeIdx - 1], giSlideIdx, giShapeIdx, depth, giSlidePart, giParentPrefix);
            giNode.Path = $"{giParentPrefix}/{BuildElementPathSegment("shape", giInnerShapes[giShapeIdx - 1], giShapeIdx)}";
            return giNode;
        }

        // Parse /slide[N] or /slide[N]/shape[M]
        var match = Regex.Match(path, @"^/slide\[(\d+)\](?:/(\w+)\[(\d+)\])?$");
        if (!match.Success)
        {
            // Generic XML fallback: navigate by element localName
            var allSegments = GenericXmlQuery.ParsePathSegments(path);
            if (allSegments.Count == 0 || !allSegments[0].Name.Equals("slide", StringComparison.OrdinalIgnoreCase) || !allSegments[0].Index.HasValue)
                throw new CliException($"Path must start with /slide[N], /slidemaster[N], or /slidelayout[N]: {path}")
                    { Code = "invalid_path" };

            var fbSlideIdx = allSegments[0].Index!.Value;
            var fbSlideParts = GetSlideParts().ToList();
            if (fbSlideIdx < 1 || fbSlideIdx > fbSlideParts.Count)
                throw new CliException($"Slide {fbSlideIdx} not found (total: {fbSlideParts.Count})")
                    { Code = "path_not_found" };

            OpenXmlElement fbCurrent = GetSlide(fbSlideParts[fbSlideIdx - 1]);
            var remaining = allSegments.Skip(1).ToList();
            if (remaining.Count > 0)
            {
                var target = GenericXmlQuery.NavigateByPath(fbCurrent, remaining);
                if (target == null)
                    throw new ArgumentException($"Element not found: {path}");
                return GenericXmlQuery.ElementToNode(target, path, depth);
            }
            return GenericXmlQuery.ElementToNode(fbCurrent, path, depth);
        }

        // BUG-R36-02 fix: int.Parse throws OverflowException for values > int.MaxValue.
        // Convert to ArgumentException to match the style of other handlers (Word/Excel).
        if (!int.TryParse(match.Groups[1].Value, out var slideIdx))
            throw new ArgumentException($"Invalid slide index '{match.Groups[1].Value}'. Must be a positive integer.");
        var slideParts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > slideParts.Count)
            throw new ArgumentException($"Slide {slideIdx} not found (total: {slideParts.Count})");

        var targetSlidePart = slideParts[slideIdx - 1];

        if (!match.Groups[2].Success)
        {
            // Return slide node
            var slide = GetSlide(targetSlidePart);
            var slideNode = new DocumentNode
            {
                Path = path,
                Type = "slide",
                Preview = slide.CommonSlideData?.ShapeTree?.Elements<Shape>()
                    .Where(IsTitle).Select(GetShapeText).FirstOrDefault() ?? "(untitled)"
            };
            var layoutName = GetSlideLayoutName(targetSlidePart);
            if (layoutName != null) slideNode.Format["layout"] = layoutName;
            var layoutType = GetSlideLayoutType(targetSlidePart);
            if (layoutType != null) slideNode.Format["layoutType"] = layoutType;
            if (slide.Show?.Value == false)
                slideNode.Format["hidden"] = true;
            ReadSlideBackground(slide, slideNode);
            ReadSlideTransition(targetSlidePart, slideNode);
            if (targetSlidePart.NotesSlidePart != null)
            {
                var notesText = GetNotesText(targetSlidePart.NotesSlidePart);
                if (!string.IsNullOrEmpty(notesText))
                    slideNode.Format["notes"] = notesText;
            }
            slideNode.Children = GetSlideChildNodes(targetSlidePart, slideIdx, depth);
            slideNode.ChildCount = slideNode.Children.Count;
            return slideNode;
        }

        // Shape or picture
        var elementType = match.Groups[2].Value;
        var elementIdx = int.Parse(match.Groups[3].Value);
        var shapeTreeEl = GetSlide(targetSlidePart).CommonSlideData?.ShapeTree;
        if (shapeTreeEl == null)
            throw new ArgumentException($"Slide {slideIdx} has no shapes");

        // BUG-R36-B11: comments live in the SlideCommentsPart, not the shape tree.
        if (elementType == "comment")
        {
            var commentsPart = targetSlidePart.SlideCommentsPart;
            var comments = commentsPart?.CommentList?
                .Elements<DocumentFormat.OpenXml.Presentation.Comment>().ToList()
                ?? new List<DocumentFormat.OpenXml.Presentation.Comment>();
            if (elementIdx < 1 || elementIdx > comments.Count)
                throw new ArgumentException($"Comment {elementIdx} not found (total: {comments.Count})");
            return CommentToNode(targetSlidePart, slideIdx, comments[elementIdx - 1], elementIdx);
        }

        // Modern p188 threaded comments live in PowerPointCommentPart(s).
        if (elementType == "moderncomment")
        {
            var top = ResolveModernComment($"/slide[{slideIdx}]/moderncomment[{elementIdx}]")
                ?? throw new ArgumentException($"Modern comment {elementIdx} not found on slide {slideIdx}");
            return ModernCommentToNode(slideIdx, top.comment, elementIdx);
        }

        if (elementType == "shape")
        {
            var shapes = shapeTreeEl.Elements<Shape>().ToList();
            if (elementIdx < 1 || elementIdx > shapes.Count)
                throw new ArgumentException($"Shape {elementIdx} not found (total: {shapes.Count})");
            return ShapeToNode(shapes[elementIdx - 1], slideIdx, elementIdx, depth, targetSlidePart);
        }
        else if (elementType == "table")
        {
            var tables = shapeTreeEl.Elements<GraphicFrame>()
                .Where(gf => gf.Descendants<Drawing.Table>().Any()).ToList();
            if (elementIdx < 1 || elementIdx > tables.Count)
                throw new ArgumentException($"Table {elementIdx} not found (total: {tables.Count})");
            return TableToNode(tables[elementIdx - 1], slideIdx, elementIdx, depth);
        }
        else if (elementType == "placeholder")
        {
            return GetPlaceholderNode(targetSlidePart, slideIdx, elementIdx, depth);
        }
        else if (elementType == "chart")
        {
            var charts = shapeTreeEl.Elements<GraphicFrame>()
                .Where(gf => gf.Descendants<C.ChartReference>().Any()
                    || IsExtendedChartFrame(gf)).ToList();
            if (elementIdx < 1 || elementIdx > charts.Count)
                throw new ArgumentException($"Chart {elementIdx} not found (total: {charts.Count})");
            return ChartToNode(charts[elementIdx - 1], targetSlidePart, slideIdx, elementIdx, depth);
        }
        else if (elementType == "picture" || elementType == "pic")
        {
            var pics = shapeTreeEl.Elements<Picture>().ToList();
            if (elementIdx < 1 || elementIdx > pics.Count)
                throw new ArgumentException($"Picture {elementIdx} not found (total: {pics.Count})");
            return PictureToNode(pics[elementIdx - 1], slideIdx, elementIdx, targetSlidePart);
        }
        else if (elementType == "audio" || elementType == "video")
        {
            var isVideoType = elementType == "video";
            var mediaList = shapeTreeEl.Elements<Picture>().Where(p =>
            {
                var nvPr = p.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties;
                return isVideoType
                    ? nvPr?.GetFirstChild<Drawing.VideoFromFile>() != null
                    : nvPr?.GetFirstChild<Drawing.AudioFromFile>() != null;
            }).ToList();
            if (elementIdx < 1 || elementIdx > mediaList.Count)
                throw new ArgumentException($"{elementType} {elementIdx} not found (total: {mediaList.Count}). " +
                    $"Slide {slideIdx} contains: {shapeTreeEl.Elements<Picture>().Count()} picture(s)");
            var mediaPic = mediaList[elementIdx - 1];
            // Find the picture's index among all pictures for PictureToNode
            var allPics = shapeTreeEl.Elements<Picture>().ToList();
            var picIdx = allPics.IndexOf(mediaPic) + 1;
            var node = PictureToNode(mediaPic, slideIdx, picIdx, targetSlidePart);
            // Override the path to use the media-type-specific path
            node.Path = $"/slide[{slideIdx}]/{BuildElementPathSegment(elementType, mediaPic, elementIdx)}";
            return node;
        }
        else if (elementType == "connector" || elementType == "connection")
        {
            var connectors = shapeTreeEl.Elements<ConnectionShape>().ToList();
            if (elementIdx < 1 || elementIdx > connectors.Count)
                throw new ArgumentException($"Connector {elementIdx} not found (total: {connectors.Count})");
            return ConnectorToNode(connectors[elementIdx - 1], slideIdx, elementIdx);
        }
        else if (elementType == "group")
        {
            var groups = shapeTreeEl.Elements<GroupShape>().ToList();
            if (elementIdx < 1 || elementIdx > groups.Count)
                throw new ArgumentException($"Group {elementIdx} not found (total: {groups.Count})");
            var grp = groups[elementIdx - 1];
            var grpName = grp.NonVisualGroupShapeProperties?.NonVisualDrawingProperties?.Name?.Value ?? "Group";
            var grpPathSeg = BuildElementPathSegment("group", grp, elementIdx);
            var grpPath = $"/slide[{slideIdx}]/{grpPathSeg}";
            var grpNode = new DocumentNode
            {
                Path = grpPath,
                Type = "group",
                Preview = grpName,
                ChildCount = grp.Elements<Shape>().Count() + grp.Elements<Picture>().Count()
                    + grp.Elements<GraphicFrame>().Count() + grp.Elements<ConnectionShape>().Count()
                    + grp.Elements<GroupShape>().Count()
            };
            grpNode.Format["name"] = grpName;
            // Bug 8 fix: read position/size from TransformGroup
            var grpXfrm = grp.GroupShapeProperties?.TransformGroup;
            if (grpXfrm?.Offset?.X != null) grpNode.Format["x"] = FormatEmu(grpXfrm.Offset.X.Value);
            if (grpXfrm?.Offset?.Y != null) grpNode.Format["y"] = FormatEmu(grpXfrm.Offset.Y.Value);
            if (grpXfrm?.Extents?.Cx != null) grpNode.Format["width"] = FormatEmu(grpXfrm.Extents.Cx.Value);
            if (grpXfrm?.Extents?.Cy != null) grpNode.Format["height"] = FormatEmu(grpXfrm.Extents.Cy.Value);
            if (grpXfrm?.Rotation != null && grpXfrm.Rotation.Value != 0)
                grpNode.Format["rotation"] = $"{grpXfrm.Rotation.Value / 60000.0:0.##}";
            var grpFillColor = ReadColorFromFill(grp.GroupShapeProperties?.GetFirstChild<Drawing.SolidFill>());
            if (grpFillColor != null) grpNode.Format["fill"] = grpFillColor;
            else if (grp.GroupShapeProperties?.GetFirstChild<Drawing.NoFill>() != null) grpNode.Format["fill"] = "none";
            else if (grp.GroupShapeProperties?.GetFirstChild<Drawing.GradientFill>() != null) grpNode.Format["fill"] = "gradient";
            // Hyperlink (nvGrpSpPr/cNvPr/a:hlinkClick) — mirrors the NodeBuilder
            // emit so round-trip Set link → reopen → Get returns the URL.
            var grpHl = grp.NonVisualGroupShapeProperties?.NonVisualDrawingProperties?
                .GetFirstChild<Drawing.HyperlinkOnClick>();
            var grpLinkUrl = ReadHyperlinkOnClickUrl(grpHl, targetSlidePart);
            if (grpLinkUrl != null) grpNode.Format["link"] = grpLinkUrl;
            var grpTip = grpHl?.Tooltip?.Value;
            if (!string.IsNullOrEmpty(grpTip)) grpNode.Format["tooltip"] = grpTip!;
            // CONSISTENCY(pptx-group-flatten): delegate to the shared walker so
            // group children include nested groups / tables / charts /
            // connectors — not just shapes and pictures.
            if (depth > 0)
            {
                BuildChildNodesIntoContainer(
                    grpNode.Children, grp, targetSlidePart, slideIdx, depth - 1,
                    grpPath, isSlideRoot: false);
            }
            return grpNode;
        }

        // Generic fallback for unknown element types
        {
            var shapes2 = shapeTreeEl.ChildElements
                .Where(e => e.LocalName.Equals(elementType, StringComparison.OrdinalIgnoreCase)).ToList();
            if (elementIdx < 1 || elementIdx > shapes2.Count)
                throw new ArgumentException($"{elementType} {elementIdx} not found (total: {shapes2.Count}). Slide {slideIdx} contains: {DescribeSlideInventory(shapeTreeEl)}");
            return GenericXmlQuery.ElementToNode(shapes2[elementIdx - 1], path, depth);
        }
    }

    // CONSISTENCY(master-layout-shape-edit): render a Shape that lives under a
    // slideMaster or slideLayout. Reuses ShapeToNode for property emission so
    // master/layout shapes Get back the same Format keys as slide shapes
    // (name, x/y/w/h, fill/stroke, runs, ...). slideNum=0 is a sentinel —
    // ShapeToNode honours parentPathPrefix when provided, so the slide index
    // is never consulted for path construction.
    private static DocumentNode GetMasterOrLayoutShapeNode(ShapeTree? shapeTree, int shapeIdx, string parentPathPrefix, int depth)
    {
        if (shapeTree == null)
            throw new ArgumentException($"No shape tree found at {parentPathPrefix}");
        var shapes = shapeTree.Elements<Shape>().ToList();
        if (shapeIdx < 1 || shapeIdx > shapes.Count)
            throw new ArgumentException($"Shape {shapeIdx} not found at {parentPathPrefix} (total: {shapes.Count})");
        return ShapeToNode(shapes[shapeIdx - 1], slideNum: 0, shapeIdx, depth, part: null, parentPathPrefix: parentPathPrefix);
    }

    public List<DocumentNode> Query(string selector)
    {
        // CONSISTENCY(query-selector-vs-path): ParseShapeSelector's regex
        // `^(\w+)` can't match a leading '/', so a path-style selector like
        // "/slide" produced elementType=null, isKnownType=true, and returned
        // ALL shapes — a false positive far worse than an empty result. Reject
        // any leading '/' selector that is NOT the supported `/slide[N]/...`
        // scoping form (handled by CONSISTENCY(query-slide-prefix) below).
        if (!string.IsNullOrEmpty(selector)
            && selector.StartsWith("/")
            && !Regex.IsMatch(selector, @"^\s*/slide\[\d+\]", RegexOptions.IgnoreCase))
            throw new ArgumentException(
                $"Invalid selector '{selector}': path-style selectors starting with '/' are not allowed in query. Use the element name (e.g. 'shape', 'slide') or a typed selector (e.g. 'shape[text=Hello]').");

        var results = new List<DocumentNode>();
        var parsed = ParseShapeSelector(selector);
        bool isEquationSelector = parsed.ElementType is "equation" or "math" or "formula";

        // Scheme B: generic XML fallback for unrecognized element types
        // Check if selector has a type that ParseShapeSelector didn't recognize
        // Extract raw element type for generic XML fallback check
        // Strip pseudo-selectors (:contains, :empty, :no-alt) and shorthand :text before checking
        var selectorForType = Regex.Replace(selector, @":(contains\([^)]*\)|empty|no-alt)", "");
        // Also strip shorthand ":text" syntax so "shape:Find me" → "shape"
        selectorForType = Regex.Replace(selectorForType, @":(?![\[\(]).*$", "");
        // Extract raw element type. If the selector starts with a slide
        // prefix ("slide[1]>shape"), strip it first; otherwise parse from
        // the beginning. Using Split(']').Last() on a selector that ENDS
        // with ']' (e.g. "ole[progId=Excel.Sheet.12]") yields an empty
        // string and the regex fails to capture — breaking the ole branch
        // dispatch and silently returning empty results.
        var typeSource = selectorForType;
        // CONSISTENCY(query-slide-prefix): strip the optional leading '/'
        // and the slide[N] prefix (with either '>' or '/' separator) so that
        // both "slide[1]>ole" and "/slide[1]/ole" resolve rawType correctly.
        var slidePrefixMatch = Regex.Match(typeSource, @"^\s*/?slide\[\d+\]\s*[>/]?\s*");
        if (slidePrefixMatch.Success)
            typeSource = typeSource.Substring(slidePrefixMatch.Length);
        else
        {
            // CONSISTENCY(query-slide-prefix): also strip unindexed `slide >` prefix
            // so `slide > shape` resolves rawType to "shape" (not "slide").
            var unindexedPrefix = Regex.Match(typeSource, @"^\s*slide\s*>\s*", RegexOptions.IgnoreCase);
            if (unindexedPrefix.Success)
                typeSource = typeSource.Substring(unindexedPrefix.Length);
        }
        var typeMatch = Regex.Match(typeSource, @"^([\w]+)");
        var rawType = typeMatch.Success ? typeMatch.Groups[1].Value.ToLowerInvariant() : "";
        bool isKnownType = string.IsNullOrEmpty(rawType)
            || rawType is "slide"
                or "shape" or "textbox" or "title" or "picture" or "pic"
                or "video" or "audio"
                or "equation" or "math" or "formula"
                or "table" or "chart" or "placeholder" or "notes"
                or "connector" or "connection"
                or "group" or "zoom"
                or "slidemaster" or "slidelayout"
                or "theme"
                or "media" or "image"
                // CONSISTENCY(ole-alias): "oleobject" mirrors Add's case switch
                or "ole" or "oleobject" or "object" or "embed"
                or "animation" or "animate"
                or "tc" or "cell" or "tr" or "row"
                // BUG-R36-B11: query("comment") enumerates all slide comments.
                or "comment"
                // Modern p188 threaded comments.
                or "moderncomment" or "modern-comment" or "thread" or "threadedcomment"
                // R8-8: paragraph/run as root selectors — walk every shape's
                // text body and emit one node per paragraph or run, matching
                // the docx surface where query("run") returns all body runs.
                or "paragraph" or "p" or "run" or "r";
        if (!isKnownType)
        {
            var genericParsed = GenericXmlQuery.ParseSelector(selector);
            foreach (var slidePart in GetSlideParts())
            {
                results.AddRange(GenericXmlQuery.Query(
                    GetSlide(slidePart), genericParsed.element, genericParsed.attrs, genericParsed.containsText));
            }
            return results;
        }

        // Theme query — schema advertises query=true; reuse Get("/theme").
        // CONSISTENCY(query-selector-vs-path): path format `/theme` (no index)
        // mirrors the Get path; PPTX has a single active theme.
        if (rawType == "theme")
        {
            var themeNode = GetThemeNode();
            if (themeNode != null)
                results.Add(themeNode);
            return results;
        }

        // BUG-R34-01: top-level slide query — `query slide` previously fell into the
        // generic XML fallback (rawType "slide" wasn't in isKnownType) and returned 0.
        // Emit one node per slide using the same shape as Get("/slide[N]") without
        // children (depth=0) so callers get a flat list of slide handles.
        if (rawType == "slide")
        {
            int qSlideNum = 0;
            foreach (var sp in GetSlideParts())
            {
                qSlideNum++;
                if (parsed.SlideNum.HasValue && parsed.SlideNum.Value != qSlideNum) continue;
                var sld = GetSlide(sp);
                var slideNode = new DocumentNode
                {
                    Path = $"/slide[{qSlideNum}]",
                    Type = "slide",
                    Preview = sld.CommonSlideData?.ShapeTree?.Elements<Shape>()
                        .Where(IsTitle).Select(GetShapeText).FirstOrDefault() ?? "(untitled)"
                };
                var lName = GetSlideLayoutName(sp);
                if (lName != null) slideNode.Format["layout"] = lName;
                var lType = GetSlideLayoutType(sp);
                if (lType != null) slideNode.Format["layoutType"] = lType;
                if (sld.Show?.Value == false)
                    slideNode.Format["hidden"] = true;
                ReadSlideBackground(sld, slideNode);
                ReadSlideTransition(sp, slideNode);
                if (sp.NotesSlidePart != null)
                {
                    var notesText = GetNotesText(sp.NotesSlidePart);
                    if (!string.IsNullOrEmpty(notesText))
                        slideNode.Format["notes"] = notesText;
                }
                var shapeTree = sld.CommonSlideData?.ShapeTree;
                slideNode.ChildCount = (shapeTree?.Elements<Shape>().Count() ?? 0)
                    + (shapeTree?.Elements<Picture>().Count() ?? 0)
                    + (shapeTree?.Elements<GraphicFrame>().Count() ?? 0)
                    + (shapeTree?.Elements<ConnectionShape>().Count() ?? 0)
                    + (shapeTree?.Elements<GroupShape>().Count() ?? 0);

                if (parsed.TextContains != null)
                {
                    var allText = string.Concat((shapeTree?.Descendants<Drawing.Text>() ?? Enumerable.Empty<Drawing.Text>()).Select(t => t.Text));
                    if (!allText.Contains(parsed.TextContains, StringComparison.OrdinalIgnoreCase))
                        continue;
                }
                if (MatchesGenericAttributes(slideNode, parsed.Attributes))
                    results.Add(slideNode);
            }
            return results;
        }

        // R8-8: top-level paragraph/run query — walk every slide's shape tree
        // and surface the paragraph/run sub-nodes that ShapeToNode emits.
        // Without this, the docx vocabulary `query "run"` returned 0 in
        // PowerPoint because rawType fell into the generic XML fallback (which
        // matched no a:r elements at the slide-XML root and bailed).
        if (rawType is "paragraph" or "p" or "run" or "r")
        {
            bool wantRun = rawType is "run" or "r";
            int qSlideNum = 0;
            foreach (var sp in GetSlideParts())
            {
                qSlideNum++;
                if (parsed.SlideNum.HasValue && parsed.SlideNum.Value != qSlideNum) continue;
                var slideNode = Get($"/slide[{qSlideNum}]", depth: 3);
                CollectParagraphsOrRuns(slideNode, wantRun, parsed.TextContains, parsed.Attributes, results);
            }
            return results;
        }

        // BUG-R36-B11: comment query — enumerate per-slide comments.
        if (rawType == "comment")
        {
            var slideFilter = parsed.SlideNum;
            var commentNodes = EnumerateComments(slideFilter);
            foreach (var n in commentNodes)
            {
                if (parsed.TextContains != null
                    && !(n.Text ?? "").Contains(parsed.TextContains, StringComparison.OrdinalIgnoreCase))
                    continue;
                if (MatchesGenericAttributes(n, parsed.Attributes))
                    results.Add(n);
            }
            return results;
        }

        // Modern p188 threaded comments — enumerate top-level threads
        // (replies live as children of each top-level node).
        if (rawType is "moderncomment" or "modern-comment" or "thread" or "threadedcomment")
        {
            var slideFilter = parsed.SlideNum;
            var mcNodes = EnumerateModernComments(slideFilter);
            foreach (var n in mcNodes)
            {
                if (parsed.TextContains != null
                    && !(n.Text ?? "").Contains(parsed.TextContains, StringComparison.OrdinalIgnoreCase))
                    continue;
                if (MatchesGenericAttributes(n, parsed.Attributes))
                    results.Add(n);
            }
            return results;
        }

        // Slide master query
        if (rawType == "slidemaster")
        {
            var masters = _doc.PresentationPart?.SlideMasterParts?.ToList() ?? [];
            for (int mi = 0; mi < masters.Count; mi++)
            {
                var mp = masters[mi];
                var masterNode = new DocumentNode
                {
                    Path = $"/slidemaster[{mi + 1}]",
                    Type = "slidemaster",
                    Preview = mp.SlideMaster?.CommonSlideData?.Name?.Value ?? "(unnamed)"
                };
                masterNode.Format["name"] = masterNode.Preview;
                masterNode.Format["layoutCount"] = mp.SlideLayoutParts?.Count() ?? 0;
                if (mp.ThemePart?.Theme?.Name?.Value != null)
                    masterNode.Format["theme"] = mp.ThemePart.Theme.Name.Value;
                // CONSISTENCY(slidemaster-emit): Get emits shapeCount; Query must
                // surface the same canonical keys so both code paths agree.
                var msShapeTree = mp.SlideMaster?.CommonSlideData?.ShapeTree;
                masterNode.Format["shapeCount"] = (msShapeTree?.Elements<Shape>().Count() ?? 0)
                    + (msShapeTree?.Elements<Picture>().Count() ?? 0);
                results.Add(masterNode);
            }
            return results;
        }

        // Slide layout query
        if (rawType == "slidelayout")
        {
            int globalIdx = 0;
            foreach (var mp in _doc.PresentationPart?.SlideMasterParts ?? Enumerable.Empty<SlideMasterPart>())
            {
                foreach (var lp in mp.SlideLayoutParts ?? Enumerable.Empty<SlideLayoutPart>())
                {
                    globalIdx++;
                    var lNode = new DocumentNode
                    {
                        Path = $"/slidelayout[{globalIdx}]",
                        Type = "slidelayout",
                        Preview = lp.SlideLayout?.CommonSlideData?.Name?.Value ?? "(unnamed)"
                    };
                    lNode.Format["name"] = lNode.Preview;
                    if (lp.SlideLayout?.Type?.HasValue == true)
                        lNode.Format["type"] = lp.SlideLayout.Type.InnerText;
                    results.Add(lNode);
                }
            }
            return results;
        }

        // Media/image query
        if (rawType is "media" or "image")
        {
            int mediaSlideNum = 0;
            foreach (var slidePart in GetSlideParts())
            {
                mediaSlideNum++;
                var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree;
                if (shapeTree == null) continue;
                int picIdx = 0;
                foreach (var pic in shapeTree.Elements<Picture>())
                {
                    picIdx++;
                    var picNvPr = pic.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties;
                    var isVideo = picNvPr?.GetFirstChild<Drawing.VideoFromFile>() != null;
                    var isAudio = picNvPr?.GetFirstChild<Drawing.AudioFromFile>() != null;
                    var mediaType = isVideo ? "video" : isAudio ? "audio" : "picture";
                    // For "image" selector, skip video/audio
                    if (rawType == "image" && mediaType != "picture") continue;
                    var picNode = PictureToNode(pic, mediaSlideNum, picIdx, slidePart);
                    picNode.Format["mediaType"] = mediaType;
                    // Add content type from image part
                    var blipFill = pic.BlipFill;
                    var blip = blipFill?.Blip;
                    if (blip?.Embed?.Value != null)
                    {
                        var part = slidePart.GetPartById(blip.Embed.Value);
                        if (part != null)
                        {
                            picNode.Format["contentType"] = part.ContentType;
                            picNode.Format["fileSize"] = part.GetStream().Length;
                        }
                    }
                    results.Add(picNode);
                }
            }
            return results;
        }

        // OLE object query. In PPTX, embedded OLE lives inside a
        // <p:graphicFrame> whose <a:graphicData uri="...ole"> contains a
        // <p:oleObj> element naming the progId + backing rel id. We also
        // surface any orphan embedded parts the slide may have — same
        // rationale as the Excel reader: forensics + zero silent loss.
        // CONSISTENCY(ole-alias): "oleobject" mirrors Add's case switch
        if (rawType is "ole" or "oleobject" or "object" or "embed")
        {
            int oleSlideNum = 0;
            foreach (var slidePart in GetSlideParts())
            {
                oleSlideNum++;
                // CONSISTENCY(query-slide-scope): match the shape/picture/table
                // branch below — apply parsed.SlideNum so that `slide[2]>ole`
                // returns only slide 2's OLE objects instead of leaking all
                // slides' results.
                if (parsed.SlideNum.HasValue && parsed.SlideNum.Value != oleSlideNum)
                    continue;
                var nodes = CollectOleNodesForSlide(oleSlideNum, slidePart);
                foreach (var n in nodes)
                {
                    // CONSISTENCY(query-attr-filter): match Word/Excel OLE query
                    // and the non-OLE PPT shape branch — apply generic attribute
                    // filter (e.g. progId=...) so users can narrow OLE results.
                    if (MatchesGenericAttributes(n, parsed.Attributes))
                        results.Add(n);
                }
            }
            return results;
        }

        // Notes query (notes live outside the shape tree in NotesSlidePart)
        if (rawType == "notes")
        {
            int notesSlideNum = 0;
            foreach (var sp in GetSlideParts())
            {
                notesSlideNum++;
                if (sp.NotesSlidePart == null) continue;
                var notesText = GetNotesText(sp.NotesSlidePart);
                if (string.IsNullOrEmpty(notesText)) continue;
                if (parsed.TextContains != null && !notesText.Contains(parsed.TextContains, StringComparison.OrdinalIgnoreCase))
                    continue;
                var notesQueryNode = new DocumentNode
                {
                    Path = $"/slide[{notesSlideNum}]/notes",
                    Type = "notes",
                    Text = notesText
                };
                notesQueryNode.Format["text"] = notesText;
                results.Add(notesQueryNode);
            }
            return results;
        }

        // Animation query: /slide[N]?/shape[M]?/animation (+ optional [attr=val] filter)
        // Enumerates every entrance/exit/emphasis effect on every shape across all slides.
        // Motion-path animations are excluded (handled separately).
        if (rawType is "animation" or "animate")
        {
            int animSlideNum = 0;
            foreach (var slidePart in GetSlideParts())
            {
                animSlideNum++;
                if (parsed.SlideNum.HasValue && parsed.SlideNum.Value != animSlideNum) continue;
                var animShapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree;
                if (animShapeTree == null) continue;

                int animShapeIdx = 0;
                foreach (var animShape in animShapeTree.Elements<Shape>())
                {
                    animShapeIdx++;
                    var effectCTns = EnumerateShapeAnimationCTns(slidePart, animShape);
                    if (effectCTns.Count == 0) continue;
                    var shapePathSeg = BuildElementPathSegment("shape", animShape, animShapeIdx);
                    for (int ai = 0; ai < effectCTns.Count; ai++)
                    {
                        var node = new DocumentNode
                        {
                            Path = $"/slide[{animSlideNum}]/{shapePathSeg}/animation[{ai + 1}]",
                            Type = "animation"
                        };
                        PopulateAnimationNode(node, effectCTns[ai]);
                        if (MatchesGenericAttributes(node, parsed.Attributes))
                            results.Add(node);
                    }
                }
            }
            return results;
        }

        int slideNum = 0;

        foreach (var slidePart in GetSlideParts())
        {
            slideNum++;

            // Slide filter
            if (parsed.SlideNum.HasValue && parsed.SlideNum.Value != slideNum)
                continue;

            var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree;
            if (shapeTree == null) continue;

            // CONSISTENCY(pptx-group-flatten): one recursive walk per slide,
            // cached as list so each type block filters without re-walking.
            // Walker descends into GroupShape, so y.ParentPath carries the
            // group ancestor chain — passed to *ToNode helpers so emitted
            // paths are honest (`/slide[1]/group[2]/shape[3]`).
            var slideRoot = $"/slide[{slideNum}]";
            var allRenderables = EnumerateRenderableElements(shapeTree, slideRoot).ToList();

            foreach (var y in allRenderables)
            {
                if (y.TypeName != "shape") continue;
                var shape = (Shape)y.Element;
                if (isEquationSelector)
                {
                    var mathElements = FindShapeMathElements(shape);
                    foreach (var mathElem in mathElements)
                    {
                        var latex = FormulaParser.ToLatex(mathElem);
                        if (parsed.TextContains == null || latex.Contains(parsed.TextContains))
                        {
                            results.Add(new DocumentNode
                            {
                                Path = $"{y.ParentPath}/{BuildElementPathSegment("shape", shape, y.IndexInParent)}",
                                Type = "equation",
                                Text = latex,
                                Format = { ["mode"] = "display" }
                            });
                        }
                    }
                }
                else if (MatchesShapeSelector(shape, parsed))
                {
                    var node = ShapeToNode(shape, slideNum, y.IndexInParent, 0, slidePart, y.ParentPath);
                    if (MatchesGenericAttributes(node, parsed.Attributes))
                        results.Add(node);
                }
            }

            if (parsed.ElementType is "picture" or "pic" or "video" or "audio" or null)
            {
                foreach (var y in allRenderables)
                {
                    if (y.TypeName != "picture") continue;
                    var pic = (Picture)y.Element;
                    var picNvPr = pic.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties;
                    var picIsVideo = picNvPr?.GetFirstChild<Drawing.VideoFromFile>() != null;
                    var picIsAudio = picNvPr?.GetFirstChild<Drawing.AudioFromFile>() != null;

                    // Filter by media type
                    if (parsed.ElementType == "video" && !picIsVideo) continue;
                    if (parsed.ElementType == "audio" && !picIsAudio) continue;
                    if (parsed.ElementType is "picture" or "pic" && (picIsVideo || picIsAudio)) continue;

                    if (MatchesPictureSelector(pic, parsed))
                    {
                        var picNode = PictureToNode(pic, slideNum, y.IndexInParent, slidePart, y.ParentPath);
                        if (MatchesGenericAttributes(picNode, parsed.Attributes))
                            results.Add(picNode);
                    }
                }
            }

            if (parsed.ElementType == "table" || (parsed.ElementType == null && !isEquationSelector))
            {
                foreach (var y in allRenderables)
                {
                    if (y.TypeName != "table") continue;
                    var gf = (GraphicFrame)y.Element;
                    var tblNode = TableToNode(gf, slideNum, y.IndexInParent, 0, y.ParentPath);
                    if (parsed.TextContains != null)
                    {
                        // GraphicData children may be opaque when loaded from disk,
                        // so extract text from all <a:t> elements via OuterXml
                        var xml = gf.OuterXml;
                        var textMatches = Regex.Matches(xml, @"<a:t[^>]*>([^<]*)</a:t>");
                        var allText = string.Concat(textMatches.Select(m => m.Groups[1].Value));
                        if (!allText.Contains(parsed.TextContains, StringComparison.OrdinalIgnoreCase))
                            continue;
                    }
                    results.Add(tblNode);
                }
            }

            // Table cell (tc/cell) and row (tr/row) query — returns friendly paths
            if (parsed.ElementType is "tc" or "cell" or "tr" or "row")
            {
                foreach (var y in allRenderables)
                {
                    if (y.TypeName != "table") continue;
                    var gf = (GraphicFrame)y.Element;
                    var tbl = gf.Descendants<Drawing.Table>().FirstOrDefault();
                    if (tbl == null) continue;
                    var tblPathSeg2 = BuildElementPathSegment("table", gf, y.IndexInParent);
                    var tblPath2 = $"{y.ParentPath}/{tblPathSeg2}";
                    int rIdx = 0;
                    foreach (var row in tbl.Elements<Drawing.TableRow>())
                    {
                        rIdx++;
                        if (parsed.ElementType is "tr" or "row")
                        {
                            var rowText = string.Join(" | ", row.Elements<Drawing.TableCell>().Select(c => c.TextBody?.InnerText ?? ""));
                            var rowNode = new DocumentNode
                            {
                                Path = $"{tblPath2}/tr[{rIdx}]",
                                Type = "tr",
                                Text = rowText,
                                ChildCount = row.Elements<Drawing.TableCell>().Count()
                            };
                            if (parsed.TextContains == null || rowText.Contains(parsed.TextContains, StringComparison.OrdinalIgnoreCase))
                            {
                                if (MatchesGenericAttributes(rowNode, parsed.Attributes))
                                    results.Add(rowNode);
                            }
                        }
                        else
                        {
                            int cIdx = 0;
                            foreach (var cell in row.Elements<Drawing.TableCell>())
                            {
                                cIdx++;
                                var cellText = cell.TextBody?.InnerText ?? "";
                                var cellNode = new DocumentNode
                                {
                                    Path = $"{tblPath2}/tr[{rIdx}]/tc[{cIdx}]",
                                    Type = "tc",
                                    Text = cellText
                                };
                                if (parsed.TextContains == null || cellText.Contains(parsed.TextContains, StringComparison.OrdinalIgnoreCase))
                                {
                                    if (MatchesGenericAttributes(cellNode, parsed.Attributes))
                                        results.Add(cellNode);
                                }
                            }
                        }
                    }
                }
            }

            if (parsed.ElementType == "chart" || (parsed.ElementType == null && !isEquationSelector))
            {
                foreach (var y in allRenderables)
                {
                    if (y.TypeName != "chart") continue;
                    var gf = (GraphicFrame)y.Element;
                    var chartNode = ChartToNode(gf, slidePart, slideNum, y.IndexInParent, 0, y.ParentPath);
                    if (parsed.TextContains != null)
                    {
                        var titleVal = chartNode.Format.ContainsKey("title") ? chartNode.Format["title"]?.ToString() ?? "" : "";
                        if (!titleVal.Contains(parsed.TextContains!, StringComparison.OrdinalIgnoreCase))
                            continue;
                    }
                    results.Add(chartNode);
                }
            }

            if (parsed.ElementType is "connector" or "connection" || (parsed.ElementType == null && !isEquationSelector))
            {
                foreach (var y in allRenderables)
                {
                    if (y.TypeName != "connector") continue;
                    var cxn = (ConnectionShape)y.Element;
                    var cxnNode = ConnectorToNode(cxn, slideNum, y.IndexInParent, y.ParentPath);
                    if (MatchesGenericAttributes(cxnNode, parsed.Attributes))
                        results.Add(cxnNode);
                }
            }

            if (parsed.ElementType == "group" || (parsed.ElementType == null && !isEquationSelector))
            {
                foreach (var y in allRenderables)
                {
                    if (y.TypeName != "group") continue;
                    var grp = (GroupShape)y.Element;
                    var grpName = grp.NonVisualGroupShapeProperties?.NonVisualDrawingProperties?.Name?.Value ?? "Group";
                    var grpNode = new DocumentNode
                    {
                        Path = $"{y.ParentPath}/{BuildElementPathSegment("group", grp, y.IndexInParent)}",
                        Type = "group",
                        Preview = grpName,
                        ChildCount = grp.Elements<Shape>().Count() + grp.Elements<Picture>().Count()
                        + grp.Elements<GraphicFrame>().Count() + grp.Elements<ConnectionShape>().Count()
                        + grp.Elements<GroupShape>().Count()
                    };
                    grpNode.Format["name"] = grpName;
                    if (MatchesGenericAttributes(grpNode, parsed.Attributes))
                        results.Add(grpNode);
                }
            }

            if (parsed.ElementType == "zoom" || (parsed.ElementType == null && !isEquationSelector))
            {
                var zoomElements = GetZoomElements(shapeTree);
                int zmIdx = 0;
                foreach (var zmEl in zoomElements)
                {
                    zmIdx++;
                    var zmNode = ZoomToNode(zmEl, slideNum, zmIdx);
                    if (parsed.TextContains != null)
                    {
                        var zmName = zmNode.Format.ContainsKey("name") ? zmNode.Format["name"]?.ToString() ?? "" : "";
                        if (!zmName.Contains(parsed.TextContains, StringComparison.OrdinalIgnoreCase))
                            continue;
                    }
                    if (MatchesGenericAttributes(zmNode, parsed.Attributes))
                        results.Add(zmNode);
                }
            }

            if (parsed.ElementType == "placeholder")
            {
                int phIdx = 0;
                // Track placeholder identity (type+idx pair) so we can skip
                // layout-inherited entries already materialized on the slide.
                var seenSlidePh = new HashSet<string>();
                foreach (var shape in shapeTree.Elements<Shape>())
                {
                    var ph = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
                        ?.GetFirstChild<PlaceholderShape>();
                    if (ph == null) continue;
                    phIdx++;
                    seenSlidePh.Add($"{ph.Type?.InnerText ?? ""}|{ph.Index?.Value.ToString() ?? ""}");

                    if (parsed.TextContains != null)
                    {
                        var shapeText = GetShapeText(shape);
                        if (!shapeText.Contains(parsed.TextContains, StringComparison.OrdinalIgnoreCase))
                            continue;
                    }

                    var node = ShapeToNode(shape, slideNum, phIdx, 0, slidePart);
                    node.Path = $"/slide[{slideNum}]/placeholder[{phIdx}]";
                    node.Type = "placeholder";
                    if (ph.Type?.HasValue == true) node.Format["phType"] = ph.Type.InnerText;
                    if (ph.Index?.HasValue == true) node.Format["phIndex"] = ph.Index.Value;
                    results.Add(node);
                }

                // Surface layout-inherited placeholders the slide hasn't
                // overridden — query previously skipped them entirely
                // because they live in the layout's shapeTree, not the
                // slide's. set/get of a layout-inherited placeholder
                // materializes a slide shape on demand (see
                // ResolvePlaceholderShape), so callers need a way to
                // discover them through query.
                var layoutShapeTree = slidePart.SlideLayoutPart?.SlideLayout?.CommonSlideData?.ShapeTree;
                if (layoutShapeTree != null)
                {
                    foreach (var layoutShape in layoutShapeTree.Elements<Shape>())
                    {
                        var lph = layoutShape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
                            ?.GetFirstChild<PlaceholderShape>();
                        if (lph == null) continue;
                        var key = $"{lph.Type?.InnerText ?? ""}|{lph.Index?.Value.ToString() ?? ""}";
                        if (seenSlidePh.Contains(key)) continue;
                        phIdx++;

                        if (parsed.TextContains != null)
                        {
                            var lShapeText = GetShapeText(layoutShape);
                            if (!lShapeText.Contains(parsed.TextContains, StringComparison.OrdinalIgnoreCase))
                                continue;
                        }

                        var lNode = ShapeToNode(layoutShape, slideNum, phIdx, 0, slidePart);
                        // Stable selector: type-name path resolves through
                        // ResolvePlaceholderShape's layout fallback at get/set.
                        var phTypeName = lph.Type?.InnerText;
                        lNode.Path = !string.IsNullOrEmpty(phTypeName)
                            ? $"/slide[{slideNum}]/placeholder[{phTypeName}]"
                            : $"/slide[{slideNum}]/placeholder[{phIdx}]";
                        lNode.Type = "placeholder";
                        if (lph.Type?.HasValue == true) lNode.Format["phType"] = lph.Type.InnerText;
                        if (lph.Index?.HasValue == true) lNode.Format["phIndex"] = lph.Index.Value;
                        lNode.Format["inheritedFrom"] = "layout";
                        results.Add(lNode);
                    }
                }
            }
        }

        return results;
    }

    // ==================== Animation helpers ====================

    /// <summary>
    /// Returns the ordered list of effect CommonTimeNodes for the given shape,
    /// including entrance/exit/emphasis presets (PresetClass set) and motion-path
    /// effects (presetClass="motion" raw attribute, no SDK enum). L3 sub-B
    /// promoted motion paths into this list so animation[K] indexing covers
    /// every animation surface on the shape.
    /// CONSISTENCY(animation-chain): Add/Set/Get/Remove all rely on this single
    /// enumeration order — keep the predicate in sync with what each writer
    /// emits (ApplyShapeAnimation + AppendMotionPathAnimation).
    /// </summary>
    private List<CommonTimeNode> EnumerateShapeAnimationCTns(SlidePart slidePart, Shape shape)
    {
        var shapeId = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value;
        if (shapeId == null) return [];
        var timing = GetSlide(slidePart).GetFirstChild<Timing>();
        if (timing == null) return [];
        var shapeIdStr = shapeId.Value.ToString();
        return timing.Descendants<CommonTimeNode>()
            .Where(ctn =>
            {
                if (!ctn.Descendants<ShapeTarget>().Any(st => st.ShapeId?.Value == shapeIdStr))
                    return false;
                // Regular entrance/exit/emphasis: SDK PresetClass set + PresetId set.
                if (ctn.PresetClass != null && ctn.PresetId != null) return true;
                // Motion path: SDK has no enum for "motion", so it's stored as a
                // raw attribute and PresetClass parses to null. Match via the
                // raw attribute + AnimateMotion descendant.
                var rawCls = ctn.GetAttributes()
                    .FirstOrDefault(a => a.LocalName == "presetClass").Value;
                if (rawCls == "motion" && ctn.Descendants<AnimateMotion>().Any())
                    return true;
                return false;
            })
            .ToList();
    }

    /// <summary>
    /// Populates a DocumentNode's Format with effect/class/presetId/duration/easing/delay fields
    /// from the given animation CommonTimeNode. Mirrors the single-Get implementation.
    /// </summary>
    private static void PopulateAnimationNode(DocumentNode animNode, CommonTimeNode effectCTn)
    {
        // L3 sub-B: motion-path effects are stored under presetClass="motion"
        // (a raw attribute the SDK doesn't model), with a child p:animMotion.
        // Surface as class=motion + path=<preset|"custom"> + d=<raw path> so
        // round-trip through Set/Get/Remove uses the same path= vocabulary as
        // AddMotionAnimation. CONSISTENCY(animation-motion-presets).
        var rawClsAttr = effectCTn.GetAttributes()
            .FirstOrDefault(a => a.LocalName == "presetClass").Value;
        if (rawClsAttr == "motion")
        {
            var animMotion = effectCTn.Descendants<AnimateMotion>().FirstOrDefault();
            var pathStr = animMotion?.Path?.Value ?? "";
            animNode.Format["class"] = "motion";
            animNode.Format["effect"] = "motion";
            var (preset, dir) = ResolveMotionPreset(pathStr);
            if (preset != null)
            {
                animNode.Format["path"] = preset;
                if (dir != null) animNode.Format["direction"] = dir;
            }
            else
            {
                animNode.Format["path"] = "custom";
                animNode.Format["d"] = pathStr;
            }
            animNode.Format["motionPath"] = pathStr;
            // Duration / trigger / delay / easing all stored on the same nodes
            // as preset effects — fall through to the regular extraction below.
            // Re-use the unified Read path by setting marker locals.
            var motionDur = 500;
            if (int.TryParse(animMotion?.CommonBehavior?.CommonTimeNode?.Duration, out var mdur)) motionDur = mdur;
            else if (int.TryParse(effectCTn.Duration, out var mdur2)) motionDur = mdur2;
            animNode.Format["duration"] = motionDur;
            var ntm = effectCTn.NodeType?.Value;
            animNode.Format["trigger"] = ntm == TimeNodeValues.AfterEffect ? "afterPrevious"
                : ntm == TimeNodeValues.WithEffect ? "withPrevious"
                : "onClick";
            if (effectCTn.Acceleration?.HasValue == true && effectCTn.Acceleration.Value > 0)
                animNode.Format["easein"] = (int)(effectCTn.Acceleration.Value / 1000);
            if (effectCTn.Deceleration?.HasValue == true && effectCTn.Deceleration.Value > 0)
                animNode.Format["easeout"] = (int)(effectCTn.Deceleration.Value / 1000);
            // Walk up for delay (mid cTn pattern).
            CommonTimeNode? midM = null;
            var curM = effectCTn.Parent;
            for (int wd = 0; wd < 5 && curM != null; wd++)
            {
                if (curM is CommonTimeNode cand && cand != effectCTn && cand.PresetId == null)
                { midM = cand; break; }
                curM = curM.Parent;
            }
            var dlyValM = midM?.StartConditionList?.GetFirstChild<Condition>()?.Delay?.Value;
            if (dlyValM != null && dlyValM != "0"
                && int.TryParse(dlyValM, out var dMsM) && dMsM > 0)
                animNode.Format["delay"] = dMsM;
            return;
        }

        var presetId = effectCTn.PresetId?.Value ?? 0;
        var clsVal = effectCTn.PresetClass?.Value;
        var cls = clsVal == TimeNodePresetClassValues.Exit ? "exit"
                : clsVal == TimeNodePresetClassValues.Emphasis ? "emphasis"
                : "entrance";

        var animEffect = effectCTn.Descendants<AnimateEffect>().FirstOrDefault();
        var filter = animEffect?.Filter?.Value ?? "";

        // CONSISTENCY(anim-preset-map): use shared resolver in Animations.cs so
        // sub-path Get returns the same effect name as slide-level shape Get.
        var effectName = ResolveAnimEffectName(filter, (int)presetId, cls);

        animNode.Format["effect"] = effectName;
        animNode.Format["class"] = cls;
        animNode.Format["presetId"] = presetId;

        // CONSISTENCY(anim-direction-readback): mirror the slide-level shape Get
        // direction decode in Animations.cs (`Read direction from presetSubtype`).
        // Without this key, dump emit cannot round-trip directional effects
        // (fly-down → fly-up on replay, since AddAnimation defaults direction).
        var presetSubtype = effectCTn.PresetSubtype?.Value ?? 0;
        var dirStr = presetSubtype switch
        {
            8 => "left",
            2 => "right",
            1 when effectName is "fly" or "wipe" or "crawl" => "up",
            4 when effectName is "fly" or "wipe" or "crawl" => "down",
            _ => (string?)null
        };
        if (dirStr != null) animNode.Format["direction"] = dirStr;

        // bt-2 fix: surface trigger (encoded as effectCTn.NodeType in OOXML).
        // ClickEffect → onclick, AfterEffect → afterPrevious, WithEffect → withPrevious.
        var nt = effectCTn.NodeType?.Value;
        animNode.Format["trigger"] = nt == TimeNodeValues.AfterEffect ? "afterPrevious"
            : nt == TimeNodeValues.WithEffect ? "withPrevious"
            : "onClick";

        var dur = 500;
        if (int.TryParse(animEffect?.CommonBehavior?.CommonTimeNode?.Duration, out var d)) dur = d;
        else if (int.TryParse(effectCTn.Descendants<AnimateScale>().FirstOrDefault()?.CommonBehavior?.CommonTimeNode?.Duration, out var d2)) dur = d2;
        else if (int.TryParse(effectCTn.Descendants<AnimateRotation>().FirstOrDefault()?.CommonBehavior?.CommonTimeNode?.Duration, out var d3)) dur = d3;
        else if (int.TryParse(effectCTn.Descendants<Animate>().FirstOrDefault()?.CommonBehavior?.CommonTimeNode?.Duration, out var d4)) dur = d4;
        else if (int.TryParse(effectCTn.Duration, out var d5)) dur = d5;
        animNode.Format["duration"] = dur;

        if (effectCTn.Acceleration?.HasValue == true && effectCTn.Acceleration.Value > 0)
            animNode.Format["easein"] = (int)(effectCTn.Acceleration.Value / 1000);
        if (effectCTn.Deceleration?.HasValue == true && effectCTn.Deceleration.Value > 0)
            animNode.Format["easeout"] = (int)(effectCTn.Deceleration.Value / 1000);

        // L2 props: emit only when the underlying cTn attribute is present.
        // OOXML repeatCount is 1000ths-of-a-count or the literal "indefinite";
        // canonical readback is the plain integer count or "indefinite".
        var rcRaw = effectCTn.RepeatCount?.Value;
        if (!string.IsNullOrEmpty(rcRaw))
        {
            if (rcRaw.Equals("indefinite", StringComparison.OrdinalIgnoreCase))
                animNode.Format["repeat"] = "indefinite";
            else if (int.TryParse(rcRaw, System.Globalization.NumberStyles.Integer,
                         System.Globalization.CultureInfo.InvariantCulture, out var rcMilli)
                     && rcMilli >= 1000)
                animNode.Format["repeat"] = rcMilli / 1000;
        }
        var restartVal = effectCTn.Restart?.Value;
        if (restartVal != null)
        {
            // Round-trip canonical enum spelling rather than the SDK ToString().
            string? canon = null;
            if (restartVal == TimeNodeRestartValues.Always) canon = "always";
            else if (restartVal == TimeNodeRestartValues.WhenNotActive) canon = "whenNotActive";
            else if (restartVal == TimeNodeRestartValues.Never) canon = "never";
            if (canon != null) animNode.Format["restart"] = canon;
        }
        if (effectCTn.AutoReverse?.Value == true)
            animNode.Format["autoReverse"] = true;

        // Delay (stored on midCTn start condition)
        CommonTimeNode? midCTn = null;
        var cur = effectCTn.Parent;
        for (int walkDepth = 0; walkDepth < 5 && cur != null; walkDepth++)
        {
            if (cur is CommonTimeNode candidate && candidate != effectCTn
                && candidate.PresetId == null)
            {
                midCTn = candidate;
                break;
            }
            cur = cur.Parent;
        }
        var midDelayVal = midCTn?.StartConditionList?.GetFirstChild<Condition>()?.Delay?.Value;
        if (midDelayVal != null && midDelayVal != "0"
            && int.TryParse(midDelayVal, out var dMs) && dMs > 0)
            animNode.Format["delay"] = dMs;
    }

    // R8-8: walk a Get-produced slide tree and harvest every paragraph or run
    // sub-node, applying the same TextContains / attribute filters Query
    // applies at the shape level. Recurses into shape/group/placeholder/table
    // bodies so all text-bearing surfaces participate.
    private static void CollectParagraphsOrRuns(
        DocumentNode node, bool wantRun, string? textContains,
        Dictionary<string, (string Value, bool Negate)>? attrs,
        List<DocumentNode> results)
    {
        if (node.Children == null) return;
        foreach (var child in node.Children)
        {
            var ct = child.Type;
            bool isPara = ct is "paragraph" or "p";
            bool isRun = ct is "run" or "r";
            bool match = wantRun ? isRun : isPara;
            if (match)
            {
                if (textContains != null
                    && !(child.Text ?? "").Contains(textContains, StringComparison.OrdinalIgnoreCase))
                {
                    // attribute matchers may still filter; skip on text miss
                }
                else if (MatchesGenericAttributes(child, attrs))
                {
                    results.Add(child);
                }
            }
            // Recurse: runs live one level under paragraphs, paragraphs under
            // shape/placeholder/cell bodies. Group/table descend as well.
            CollectParagraphsOrRuns(child, wantRun, textContains, attrs, results);
        }
    }
}
