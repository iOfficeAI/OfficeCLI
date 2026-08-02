// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    private string AddChart(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        // CONSISTENCY(host-part-rel): same routing as AddPicture (round23 E) and
        // AddHyperlink (round23 C). When the parent paragraph lives in a Header/Footer
        // part, the chart rel must live on that part — otherwise r:id in headerN.xml
        // points to a rel only present in document.xml.rels and Word reports broken.
        // Resolve the host part (MainDocument / Header / Footer / Footnotes /
        // Endnotes / Comments) so the ChartPart rel registers where the chart
        // actually lives. Mirrors AddPicture/AddHyperlink. Without this,
        // footnote/endnote charts landed under document.xml.rels and Word 422'd.
        OpenXmlPart chartMainPart = ResolveHostPart(parent);

        // Parse chart data. Use TryGetValue(case-insensitive) so reads
        // are recorded by TrackingPropertyDictionary.
        string chartType = "column";
        if (properties.TryGetValue("charttype", out var ctVal) || properties.TryGetValue("type", out ctVal))
            chartType = ctVal;
        var chartTitle = properties.GetValueOrDefault("title");
        var categories = Core.ChartHelper.ParseCategories(properties);
        var seriesData = Core.ChartHelper.ParseSeriesData(properties);

        if (seriesData.Count == 0)
            throw new ArgumentException("Chart requires data. Use: data=\"Series1:1,2,3;Series2:4,5,6\" " +
                "or series1=\"Revenue:100,200,300\"");

        // Dimensions (default: 15cm x 10cm)
        long chartCx = properties.TryGetValue("width", out var chartWStr) ? ParseEmu(chartWStr) : 5400000;
        long chartCy = properties.TryGetValue("height", out var chStr) ? ParseEmu(chStr) : 3600000;

        var docPropId = NextDocPropId();
        // BUG-R7-02 (T-2): explicit `name` prop was previously ignored —
        // dump emitted name=… on round-trip but Add silently dropped it,
        // so the chart's shape name reverted to its title every replay.
        // Honor caller intent first; fall back to title, then synthesize.
        // CONSISTENCY(empty-string-fallback): mirror AddPicture's
        // !IsNullOrEmpty guard — `??` only short-circuits on null, so a
        // literal name="" would otherwise pin the chart's shape name to
        // empty instead of falling through to title.
        var chartName = (properties.TryGetValue("name", out var chartNameOverride)
                         && !string.IsNullOrEmpty(chartNameOverride))
            ? chartNameOverride
            : (chartTitle ?? $"Chart {docPropId}");

        // Extended chart types (cx:chart) — funnel, treemap, sunburst, boxWhisker, histogram
        if (Core.ChartExBuilder.IsExtendedChartType(chartType))
        {
            var cxChartSpace = Core.ChartExBuilder.BuildExtendedChartSpace(
                chartType, chartTitle, categories, seriesData, properties);
            var extChartPart = chartMainPart.AddNewPart<ExtendedChartPart>();
            extChartPart.ChartSpace = cxChartSpace;
            extChartPart.ChartSpace.Save();

            // CONSISTENCY(chartex-sidecars): see PowerPointHandler.Add.Media.cs
            // for the full rationale. Word's chartEx host has the same hard
            // requirement on rId1 (embedded xlsx) + rId2 (style) + rId3 (colors).
            var embPart = extChartPart.AddNewPart<EmbeddedPackagePart>(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "rId1");
            var xlsxBytes = Core.ChartExResources.BuildMinimalEmbeddedXlsx(categories, seriesData);
            using (var emsr = new MemoryStream(xlsxBytes))
                embPart.FeedData(emsr);

            var stylePart = extChartPart.AddNewPart<ChartStylePart>("rId2");
            using (var styleStream = Core.ChartExResources.OpenChartStyleXml())
                stylePart.FeedData(styleStream);

            var colorPart = extChartPart.AddNewPart<ChartColorStylePart>("rId3");
            using (var colorStream = Core.ChartExResources.OpenChartColorStyleXml())
                colorPart.FeedData(colorStream);

            var cxRelId = chartMainPart.GetIdOfPart(extChartPart);
            var cxChartRef = new DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing.RelId { Id = cxRelId };

            var cxGraphic = new A.Graphic(
                new A.GraphicData(cxChartRef)
                { Uri = "http://schemas.microsoft.com/office/drawing/2014/chartex" }
            );
            var cxFrame = BuildChartFrame(cxGraphic, chartCx, chartCy, docPropId, chartName, properties);

            var cxRun = new Run(new Drawing(cxFrame));
            Paragraph cxPara;
            if (parent is Paragraph existingCxPara)
            {
                // CONSISTENCY(add-index): honor --index / --after / --before (#76).
                var cxChildren = existingCxPara.ChildElements.ToList();
                if (index.HasValue && index.Value < cxChildren.Count)
                    existingCxPara.InsertBefore(cxRun, cxChildren[index.Value]);
                else
                    existingCxPara.AppendChild(cxRun);
                cxPara = existingCxPara;
            }
            else
            {
                cxPara = new Paragraph(cxRun);
                AssignParaId(cxPara);
                InsertAtIndexOrAppend(parent, cxPara, index);
            }

            // Return document-order position so it matches the resolver
            // (GetAllWordCharts). CountWordCharts is insertion-order and
            // disagrees whenever --before/--after inserts mid-document.
            var cxAllCharts = GetAllWordCharts();
            var cxDocOrderIdx = cxAllCharts.FindIndex(c => ReferenceEquals(c.Container, cxFrame));
            return $"/chart[{(cxDocOrderIdx >= 0 ? cxDocOrderIdx + 1 : cxAllCharts.Count)}]";
        }

        // BUG-R6A(BUG2): Build chart content BEFORE adding part (invalid type or
        // a chart-type cardinality violation throws, which must not leave an
        // empty/malformed ChartPart on disk). Mirrors the Excel/Pptx ordering.
        var chartSpace = Core.ChartHelper.BuildChartSpace(chartType, chartTitle, categories, seriesData, properties);
        var chartPart = chartMainPart.AddNewPart<ChartPart>();
        chartPart.ChartSpace = chartSpace;

        // Apply deferred properties (axisTitle, dataLabels, etc.) via SetChartProperties
        // Must be called BEFORE Save() so the in-memory DOM is still available.
        // CONSISTENCY(tracking-deferred-filter): see PowerPointHandler.Add.Media.cs —
        // .Where() over TrackingPropertyDictionary marks every key consumed and
        // silently swallows real typos. Iterate Keys + TryGetValue per match instead.
        var deferredProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var dk in properties.Keys.ToList())
        {
            if (Core.ChartHelper.IsDeferredKey(dk) && properties.TryGetValue(dk, out var dv))
                deferredProps[dk] = dv;
        }
        if (deferredProps.Count > 0)
            Core.ChartHelper.SetChartProperties(chartPart, deferredProps);
        else
            chartPart.ChartSpace.Save();

        // BUG-DUMP-CHART-SIDECARS: re-attach the source chart's sidecar parts
        // (chartStyle / chartColorStyle / themeOverride / embedded data
        // workbook) captured by the dump, so the rebuilt native chart keeps its
        // theme, custom colours and editable data instead of a bare default.
        AttachChartSidecars(chartPart, properties);

        // Re-attach a captured c:userShapes overlay (logo/photo/annotation drawn
        // on top of the chart) so it round-trips through dump→batch.
        AttachChartUserShapes(chartPart, properties);

        var chartRelId = chartMainPart.GetIdOfPart(chartPart);

        // Build Drawing frame (inline or floating anchor) with ChartReference.
        var chartGraphic = new A.Graphic(
            new A.GraphicData(
                new DocumentFormat.OpenXml.Drawing.Charts.ChartReference { Id = chartRelId }
            )
            { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" }
        );
        var frame = BuildChartFrame(chartGraphic, chartCx, chartCy, docPropId, chartName, properties);

        var chartRun = new Run(new Drawing(frame));
        Paragraph chartPara;
        if (parent is Paragraph existingChartPara)
        {
            // CONSISTENCY(add-index): honor --index / --after / --before (#76).
            var chartChildren = existingChartPara.ChildElements.ToList();
            if (index.HasValue && index.Value < chartChildren.Count)
                existingChartPara.InsertBefore(chartRun, chartChildren[index.Value]);
            else
                existingChartPara.AppendChild(chartRun);
            chartPara = existingChartPara;
        }
        else
        {
            chartPara = new Paragraph(chartRun);
            AssignParaId(chartPara);
            InsertAtIndexOrAppend(parent, chartPara, index);
        }

        // Return document-order position (matches GetAllWordCharts resolver).
        var allCharts = GetAllWordCharts();
        var docOrderIdx = allCharts.FindIndex(c => ReferenceEquals(c.Container, frame));
        return $"/chart[{(docOrderIdx >= 0 ? docOrderIdx + 1 : allCharts.Count)}]";
    }

    /// <summary>
    /// Re-attach a native chart's sidecar parts captured by the dump under
    /// <c>sidecar.{style|colors|themeOverride|package}.data</c> (base64 data
    /// URIs). chartStyle/chartColorStyle/themeOverride are related to the chart
    /// part by relationship type (no in-XML reference); the embedded data
    /// workbook is wired via a fresh <c>&lt;c:externalData r:id&gt;</c>. No-op
    /// when the chart carried no sidecars (e.g. a freshly authored chart).
    /// </summary>
    private void AttachChartSidecars(ChartPart chartPart, Dictionary<string, string> properties)
    {
        static (string Ct, byte[] Bytes)? ParseDataUri(string v)
        {
            if (string.IsNullOrEmpty(v) || !v.StartsWith("data:", StringComparison.Ordinal)) return null;
            var comma = v.IndexOf(";base64,", StringComparison.Ordinal);
            if (comma < 0) return null;
            var ct = v.Substring(5, comma - 5);
            try { return (ct, System.Convert.FromBase64String(v[(comma + 8)..])); }
            catch { return null; }
        }

        void Feed(OpenXmlPart part, byte[] bytes)
        {
            using var ms = new MemoryStream(bytes);
            part.FeedData(ms);
        }

        if (properties.TryGetValue("sidecar.style.data", out var styleV)
            && ParseDataUri(styleV) is { } st)
            Feed(chartPart.AddNewPart<ChartStylePart>(), st.Bytes);

        if (properties.TryGetValue("sidecar.colors.data", out var colorsV)
            && ParseDataUri(colorsV) is { } co)
            Feed(chartPart.AddNewPart<ChartColorStylePart>(), co.Bytes);

        if (properties.TryGetValue("sidecar.themeOverride.data", out var toV)
            && ParseDataUri(toV) is { } to)
            Feed(chartPart.AddNewPart<ThemeOverridePart>(), to.Bytes);

        if (properties.TryGetValue("sidecar.package.data", out var pkgV)
            && ParseDataUri(pkgV) is { } pk)
        {
            var pkgPart = chartPart.AddNewPart<EmbeddedPackagePart>(pk.Ct);
            Feed(pkgPart, pk.Bytes);
            var pkgRelId = chartPart.GetIdOfPart(pkgPart);
            var chartSpace = chartPart.ChartSpace;
            if (chartSpace != null
                && chartSpace.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.ExternalData>() == null)
            {
                var extData = new DocumentFormat.OpenXml.Drawing.Charts.ExternalData
                {
                    Id = pkgRelId,
                    AutoUpdate = new DocumentFormat.OpenXml.Drawing.Charts.AutoUpdate { Val = false },
                };
                // CT_ChartSpace order: …chart, spPr, txPr, externalData,
                // printSettings, userShapes, extLst. Insert before the first of
                // those trailing elements if present, else append.
                var anchor = chartSpace.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.PrintSettings>() as OpenXmlElement
                    ?? chartSpace.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.UserShapes>() as OpenXmlElement
                    ?? chartSpace.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.ChartSpaceExtensionList>() as OpenXmlElement;
                if (anchor != null) chartSpace.InsertBefore(extData, anchor);
                else chartSpace.AppendChild(extData);
                chartSpace.Save();
            }
        }
    }

    /// <summary>
    /// Re-create a chart's <c>c:userShapes</c> overlay (chartshapes part + its
    /// embedded images) from the dump carrier props (<c>userShapesXml</c> +
    /// <c>userShapes.partN.*</c>), and wire the <c>&lt;c:userShapes r:id&gt;</c>
    /// reference into the chartSpace. No-op when the chart carries no overlay.
    /// </summary>
    private void AttachChartUserShapes(ChartPart chartPart, Dictionary<string, string> properties)
    {
        if (!properties.TryGetValue("userShapesXml", out var userShapesXml)
            || string.IsNullOrEmpty(userShapesXml))
            return;

        // Strip the `userShapes.` prefix so MaterializeInlinedParts sees the
        // canonical part{N}.relId / part{N}.data / ext{N}.* keys.
        var carrierProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        const string prefix = "userShapes.";
        foreach (var k in properties.Keys.ToList())
        {
            if (k.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)
                && properties.TryGetValue(k, out var v))
                carrierProps[k.Substring(prefix.Length)] = v;
        }

        var cdp = chartPart.AddNewPart<ChartDrawingPart>();
        var rewrite = MaterializeInlinedParts(cdp, carrierProps, "chartUserShapes");
        var finalXml = rewrite(userShapesXml);
        using (var s = cdp.GetStream(FileMode.Create, FileAccess.Write))
        using (var w = new StreamWriter(s))
            w.Write(finalXml);

        var cdpRelId = chartPart.GetIdOfPart(cdp);
        var chartSpace = chartPart.ChartSpace;
        if (chartSpace == null) return;
        // c:userShapes is CT_RelId — its sole attribute is r:id. Set it
        // generically to stay independent of the SDK property name.
        var userShapes = new DocumentFormat.OpenXml.Drawing.Charts.UserShapes();
        userShapes.SetAttribute(new OpenXmlAttribute(
            "r", "id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", cdpRelId));
        // c:userShapes is the last CT_ChartSpace element before c:extLst.
        var extLst = chartSpace.ChildElements
            .FirstOrDefault(e => e.LocalName == "extLst");
        if (extLst != null) chartSpace.InsertBefore(userShapes, extLst);
        else chartSpace.AppendChild(userShapes);
        chartSpace.Save();
    }

    private string AddPicture(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        if (!properties.TryGetValue("path", out var imgPath) && !properties.TryGetValue("src", out imgPath))
            throw new ArgumentException("'src' property is required for picture type");
        // R49: `/chart[N]` resolves to the paragraph hosting the chart; AddPicture
        // would silently re-route the image into body and return a bogus path
        // (`/chart[1]/r[K]`) that 404s on Get. Mirrors the R47 AddTextbox guard.
        if (parentPath.StartsWith("/chart[", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException(
                $"Cannot add a picture to a chart path ('{parentPath}'). " +
                "Charts don't host picture children — use /body or a table cell as parent.");

        // Buffer the image bytes so we can both feed the image part and sniff
        // the native pixel dimensions for auto aspect-ratio calculations.
        var (rawStream, imgPartType) = OfficeCli.Core.ImageSource.Resolve(imgPath);
        using var rawStreamDispose = rawStream;
        using var imgStream = new MemoryStream();
        rawStream.CopyTo(imgStream);
        imgStream.Position = 0;

        var mainPart = _doc.MainDocumentPart!;
        // BUG-R14A: route through ResolveHostPart so the ImagePart and its
        // r:embed relationship land on whatever part actually holds the
        // <w:drawing> — header/footer AND footnote/endnote/comment. OOXML
        // resolves r:embed against the rels of the part containing the
        // drawing, so a picture added into a footnote/endnote/comment body
        // must register its image rel on word/_rels/footnotes.xml.rels (etc.),
        // not document.xml.rels — otherwise the blip dangles and validation
        // fails ([Semantic] r:embed does not exist). Previously this path did
        // its own Header/Footer-only resolution and defaulted footnote/endnote/
        // comment to MainDocumentPart. Mirrors the comment-hyperlink-rel fix
        // (BUG-R13B(BUG2)) that added the comments branch to ResolveHostPart.
        OpenXmlPart imgHostPart = ResolveHostPart(parent);

        // AddImagePart is a generic extension on parts implementing
        // ISupportedRelationship<…, ImagePart> (MainDocument / Header / Footer /
        // Footnotes / Endnotes / Comments). Dispatch by runtime type so the
        // rel lands on the correct part.
        ImagePart AddImg(PartTypeInfo t) => imgHostPart switch
        {
            MainDocumentPart mdp => mdp.AddImagePart(t),
            HeaderPart hp => hp.AddImagePart(t),
            FooterPart fp => fp.AddImagePart(t),
            FootnotesPart fnp => fnp.AddImagePart(t),
            EndnotesPart enp => enp.AddImagePart(t),
            WordprocessingCommentsPart cp => cp.AddImagePart(t),
            _ => throw new InvalidOperationException(
                $"Host part type {imgHostPart.GetType().Name} does not support image parts"),
        };

        string relId;
        string? svgRelId = null;
        Stream? fallbackDimStream = null;  // source for TryGetDimensions when raster is the fallback
        if (imgPartType == ImagePartType.Svg)
        {
            // OOXML SVG embedding: main blip points to a PNG fallback, and
            // a:blip/a:extLst carries an asvg:svgBlip referencing the SVG
            // part. Modern Office picks up the SVG; older versions render
            // the PNG. See SvgImageHelper for namespace/URI details.
            var svgPart = AddImg(ImagePartType.Svg);
            svgPart.FeedData(imgStream);
            imgStream.Position = 0;
            svgRelId = imgHostPart.GetIdOfPart(svgPart);

            MemoryStream pngStream;
            if (properties.TryGetValue("fallback", out var fallbackPath) && !string.IsNullOrWhiteSpace(fallbackPath))
            {
                var (fbRaw, fbType) = OfficeCli.Core.ImageSource.Resolve(fallbackPath);
                using var fbDispose = fbRaw;
                pngStream = new MemoryStream();
                fbRaw.CopyTo(pngStream);
                pngStream.Position = 0;
                var fbPart = AddImg(fbType);
                fbPart.FeedData(pngStream);
                pngStream.Position = 0;
                relId = imgHostPart.GetIdOfPart(fbPart);
            }
            else
            {
                var pngPart = AddImg(ImagePartType.Png);
                pngPart.FeedData(new MemoryStream(OfficeCli.Core.SvgImageHelper.TransparentPng1x1, writable: false));
                relId = imgHostPart.GetIdOfPart(pngPart);
                pngStream = new MemoryStream(OfficeCli.Core.SvgImageHelper.TransparentPng1x1, writable: false);
            }
            fallbackDimStream = pngStream;
        }
        else
        {
            var imagePart = AddImg(imgPartType);
            imagePart.FeedData(imgStream);
            imgStream.Position = 0;
            relId = imgHostPart.GetIdOfPart(imagePart);
        }

        // Determine dimensions. When only one axis is supplied, compute the
        // other from the image's native pixel aspect ratio. When neither is
        // supplied, width defaults to 6 inches and height follows the aspect
        // ratio (or a 4 inch fallback when the image header cannot be read).
        // CONSISTENCY(picture-size-alias): accept "w"/"h" as short aliases for
        // "width"/"height" — mirrors pptx shape and xlsx picture behavior.
        bool hasWidth = properties.TryGetValue("width", out var widthStr)
            || properties.TryGetValue("w", out widthStr);
        bool hasHeight = properties.TryGetValue("height", out var heightStr)
            || properties.TryGetValue("h", out heightStr);
        long cxEmu = hasWidth ? ParseEmu(widthStr!) : 5486400;  // 6 inches fallback
        long cyEmu = hasHeight ? ParseEmu(heightStr!) : 3657600; // 4 inches fallback
        // OOXML CT_PositiveSize2D / ST_PositiveCoordinate require strictly
        // positive width and height. Negative/zero EMUs would write a
        // schema-invalid <wp:extent> that Word rejects (file repair / 422).
        if (hasWidth && cxEmu <= 0)
            throw new ArgumentException($"Invalid 'width' value: '{widthStr}'. Picture width must be > 0 (OOXML ST_PositiveCoordinate).");
        if (hasHeight && cyEmu <= 0)
            throw new ArgumentException($"Invalid 'height' value: '{heightStr}'. Picture height must be > 0 (OOXML ST_PositiveCoordinate).");

        if (!hasWidth || !hasHeight)
        {
            var dims = OfficeCli.Core.ImageSource.TryGetDimensions(imgStream);
            if (dims is { Width: > 0, Height: > 0 } d)
            {
                double ratio = (double)d.Height / d.Width;
                if (hasWidth && !hasHeight)
                    cyEmu = (long)(cxEmu * ratio);
                else if (!hasWidth && hasHeight)
                    cxEmu = (long)(cyEmu / ratio);
                else
                    cyEmu = (long)(cxEmu * ratio);
            }
        }

        // BUG-R5-02: data URIs (data:image/png;base64,iVBOR...) contain
        // multiple slashes inside the base64 payload, so Path.GetFileName
        // returns a meaningless tail like "png;base64,iVBOR..." which then
        // becomes both the picture name AND the alt text. Detect data: /
        // base64-blob inputs and fall back to a neutral placeholder unless
        // the caller supplied an explicit alt= or name=.
        string DefaultPictureName()
        {
            if (string.IsNullOrEmpty(imgPath)) return "image";
            if (imgPath.StartsWith("data:", StringComparison.OrdinalIgnoreCase)) return "image";
            // Heuristic for raw base64 (no scheme): no path separator and length
            // is implausibly long for a real filename.
            if (imgPath.Length > 256 && imgPath.IndexOf('/') < 0 && imgPath.IndexOf('\\') < 0) return "image";
            try { return Path.GetFileName(imgPath); }
            catch { return "image"; }
        }
        // `name` is the picture's docPr Name (Word's Object Browser label;
        // typically "Picture 1", "Picture 2", …). `alt` is the docPr
        // Description (Word's Alt Text field, used by screen readers).
        // Source documents almost always set these to DIFFERENT values, so
        // collapsing both onto `altText` poisons dump round-trip — d2 would
        // emit the long alt text as the name. Take them separately; fall
        // back through name → alt → DefaultPictureName so callers passing
        // only one still get a sensible result.
        // CONSISTENCY(picture-alt): full alias set on input (alt canonical;
        // altText/alttext/description aliases), matching Set and the shared
        // picture schema contract. Set already accepted alttext/description —
        // Add silently ignoring them was an Add/Set asymmetry.
        var altFromProps = new[] { "alt", "altText", "alttext", "description" }
            .Select(k => properties.GetValueOrDefault(k))
            .FirstOrDefault(v => !string.IsNullOrEmpty(v));
        var pictureName = properties.TryGetValue("name", out var nameOverride) && !string.IsNullOrEmpty(nameOverride)
            ? nameOverride
            : (altFromProps ?? DefaultPictureName());
        // altText: explicit `alt=` wins. When absent, leave the Description
        // attribute blank — auto-stamping `altText = pictureName` meant Get
        // surfaced a phantom `alt=<name>` key on dump round-trip for sources
        // whose docPr had no Description (07_example_online_convert.docx).
        // Empty string keeps the OOXML attr off entirely (DocProperties
        // serialises Description="" as no attr).
        var altText = altFromProps ?? "";

        var imgDocPropId = NextDocPropId();
        // BUG-DUMP-R29-1: parse the optional effectExtent prop ("l,t,r,b", 4
        // EMU ints — may be negative) captured by the dump emit. Word's inline
        // layout height depends on this margin; without it the rebuild
        // collapses to 0/0/0/0 and downstream content drifts up. Absent →
        // null → CreateImageRun/CreateAnchorImageRun default to 0/0/0/0
        // (interactive Add back-compat). Tolerant of whitespace/bad input.
        (long L, long T, long R, long B)? effectExtent = null;
        if (properties.TryGetValue("effectExtent", out var eeStr) && !string.IsNullOrWhiteSpace(eeStr))
        {
            var eeParts = eeStr.Split(',');
            if (eeParts.Length == 4
                && long.TryParse(eeParts[0].Trim(), out var eeL)
                && long.TryParse(eeParts[1].Trim(), out var eeT)
                && long.TryParse(eeParts[2].Trim(), out var eeR)
                && long.TryParse(eeParts[3].Trim(), out var eeB))
            {
                effectExtent = (eeL, eeT, eeR, eeB);
            }
        }
        Run imgRun;
        // BUG-R4-BT3: a non-"none" `wrap` value implies floating placement —
        // wrap only has meaning on a <wp:anchor>. Previously, callers passing
        // `wrap=square|tight|topBottom|behind|inFront` without an explicit
        // `anchor=true` got an inline picture and the wrap was silently
        // dropped (also affected dump round-trip of floating pictures).
        bool wrapImpliesAnchor = properties.TryGetValue("wrap", out var implicitWrap)
            && !string.IsNullOrEmpty(implicitWrap)
            && !string.Equals(implicitWrap, "none", StringComparison.OrdinalIgnoreCase)
            && !string.Equals(implicitWrap, "inline", StringComparison.OrdinalIgnoreCase);
        // BUG-DUMP11-06: `anchor` is overloaded — historically a bool flag for
        // floating placement, but Get also surfaces the hyperlink anchor name
        // when the picture's run is wrapped in <w:hyperlink w:anchor="...">.
        // Treat bool-recognized values (true/false/yes/no/0/1/on/off) as the
        // floating switch; treat any other non-empty string as a hyperlink
        // bookmark name attached to the picture's drawing.
        bool hasAnchorProp = properties.TryGetValue("anchor", out var anchorVal)
            && !string.IsNullOrEmpty(anchorVal);
        bool anchorIsBool = hasAnchorProp && ParseHelpers.IsValidBooleanString(anchorVal);
        bool anchorIsFloating = hasAnchorProp && anchorIsBool && IsTruthy(anchorVal);
        string? hyperlinkAnchorName = hasAnchorProp && !anchorIsBool ? anchorVal : null;
        if (anchorIsFloating || wrapImpliesAnchor)
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
            // Relative <wp:align> keyword per axis (left/center/right;
            // top/bottom/center/inside/outside). When present it overrides the
            // posOffset form so an aligned float honours Word's relative
            // placement instead of collapsing to the margin origin.
            var hAlign = properties.TryGetValue("halign", out var hAlignStr) && !string.IsNullOrEmpty(hAlignStr)
                ? hAlignStr : null;
            var vAlign = properties.TryGetValue("valign", out var vAlignStr) && !string.IsNullOrEmpty(vAlignStr)
                ? vAlignStr : null;
            // BUG-DUMP-R26-1: round-trip the anchor z-order. CreateImageNode now
            // surfaces relativeHeight; honour it here so overlapping floats keep
            // their distinct stacking order instead of collapsing to 1U.
            uint relHeight = properties.TryGetValue("relativeHeight", out var relHeightStr)
                && uint.TryParse(relHeightStr, out var rh) ? rh : 1U;
            // Wrap distances (gap between the float and the text wrapping around
            // it): "T,B,L,R" EMU prop captured by the dump. Absent → null →
            // CreateAnchorImageRun keeps the interactive defaults.
            (uint T, uint B, uint L, uint R)? wrapDist = null;
            if (properties.TryGetValue("wrapDist", out var wdStr) && !string.IsNullOrWhiteSpace(wdStr))
            {
                var wd = wdStr.Split(',');
                if (wd.Length == 4
                    && uint.TryParse(wd[0].Trim(), out var wdT)
                    && uint.TryParse(wd[1].Trim(), out var wdB)
                    && uint.TryParse(wd[2].Trim(), out var wdL)
                    && uint.TryParse(wd[3].Trim(), out var wdR))
                {
                    wrapDist = (wdT, wdB, wdL, wdR);
                }
            }
            // BUG-R24-WRAPPOLY: forward a captured custom wrapTight/wrapThrough
            // polygon so the exact text-flow boundary round-trips (a source
            // polygon hugging the image tighter than the default full square
            // shifted wrapped/below text by several px).
            var wrapPolygon = properties.GetValueOrDefault("wrap.polygon");
            // BUG-DUMP-NAR-WP14: wp14 relative sizing (sizeRelH/sizeRelV =
            // "relativeFrom;pct", e.g. "page;0"). The dump captures the anchor's
            // <wp14:sizeRelH>/<wp14:sizeRelV> percentage-of-page sizing; without
            // re-applying it the picture falls back to its absolute extent and
            // reflows. Absent → null → no wp14 child emitted.
            var sizeRelH = properties.GetValueOrDefault("sizeRelH");
            var sizeRelV = properties.GetValueOrDefault("sizeRelV");
            // BUG-DUMP-WRAPSIDE: forward the wrapSquare text side (left/right/
            // largest) so a one-sided wrap round-trips instead of defaulting to
            // bothSides and reflowing the text around the float.
            var wrapSide = properties.GetValueOrDefault("wrap.side") ?? properties.GetValueOrDefault("wrapSide");
            imgRun = CreateAnchorImageRun(relId, cxEmu, cyEmu, altText, wrapType, hPos, vPos, hRel, vRel, behind, imgDocPropId, pictureName, hAlign, vAlign, relHeight, effectExtent, wrapDist, wrapPolygon, sizeRelH, sizeRelV, wrapSide);
        }
        else
        {
            imgRun = CreateImageRun(relId, cxEmu, cyEmu, altText, imgDocPropId, pictureName, effectExtent);
        }

        // BUG-DUMP-R28-PICHIDDEN: restore the drawing's hidden flag
        // (<wp:docPr hidden="1">). CreateImageRun/CreateAnchorImageRun never set
        // it, so a hidden monochrome print logo replayed visible and rendered
        // (black) on top of the visible colour logo. Apply it to the wp:docPr.
        if (properties.TryGetValue("hidden", out var hiddenVal) && IsTruthy(hiddenVal))
        {
            var hiddenDocPr = imgRun.Descendants<DW.DocProperties>().FirstOrDefault();
            if (hiddenDocPr != null) hiddenDocPr.Hidden = true;
        }

        // Accessibility: mark the image decorative (screen readers skip it). Stored
        // as an adec:decorative extension under <wp:docPr>. Word treats a decorative
        // image's alt text as absent, which is expected.
        if (properties.TryGetValue("decorative", out var decVal) && IsTruthy(decVal))
        {
            var decDocPr = imgRun.Descendants<DW.DocProperties>().FirstOrDefault();
            if (decDocPr != null) SetPictureDecorative(decDocPr);
        }

        // Wire the asvg:svgBlip extension after the run is built. Walking
        // the Drawing to find the Blip keeps CreateImageRun /
        // CreateAnchorImageRun signature-stable for non-SVG callers.
        if (svgRelId != null)
        {
            var addedBlip = imgRun.Descendants<A.Blip>().FirstOrDefault();
            if (addedBlip != null)
                OfficeCli.Core.SvgImageHelper.AppendSvgExtension(addedBlip, svgRelId);
        }

        // BUG-DUMP-R45-4: re-inject image-level visual content the dump captured
        // as verbatim XML (blip recolor/alpha children + spPr effectLst/outerShdw
        // drop shadow). These have no flat-key representation; the fixed
        // CreateImageRun/CreateAnchorImageRun rebuild would otherwise drop them.
        // Each prop is present only when the source carried that content, so a
        // plain picture is untouched.
        if (properties.TryGetValue("blipEffects", out var blipEffectsXml)
            && !string.IsNullOrWhiteSpace(blipEffectsXml))
        {
            ApplyBlipEffects(imgRun, blipEffectsXml);
        }
        // Whole-spPr verbatim (xfrm flip flags, a content <a:ext> that differs
        // from the frame's wp:extent, bwMode, explicit noFill/ln, effectLst) —
        // supersedes the narrower spEffects injection when present.
        if (properties.TryGetValue("spPrXml", out var spPrXmlVal)
            && !string.IsNullOrWhiteSpace(spPrXmlVal))
        {
            ApplySpPrVerbatim(imgRun, spPrXmlVal);
        }
        else if (properties.TryGetValue("spEffects", out var spEffectsXml)
            && !string.IsNullOrWhiteSpace(spEffectsXml))
        {
            ApplySpPrEffects(imgRun, spEffectsXml);
        }

        // BUG-DUMP-R51-1: round-trip a click-hyperlink on the image. A clickable
        // image (e.g. a logo linking to a URL) carries <a:hlinkClick r:id="…"> on
        // its <pic:cNvPr> (and, for external targets, a TargetMode="External"
        // relationship). The dump now surfaces the resolved URL as `link`; re-create
        // the external rel on the drawing's host part and attach the hlinkClick so
        // the image stays clickable instead of flattening to a plain image. A bare
        // anchor (internal bookmark, no scheme/relative target) is stored as the
        // hlinkClick @w:anchor attribute with no rel — mirrors the run-level
        // <w:hyperlink w:anchor> path. Reuses AddHyperlinkRelationship, the same
        // rel-creation helper the run-level hyperlink Add/Set paths use.
        if (properties.TryGetValue("link", out var linkVal) && !string.IsNullOrEmpty(linkVal))
        {
            var picCNvPr = imgRun.Descendants<PIC.NonVisualDrawingProperties>().FirstOrDefault();
            if (picCNvPr != null)
            {
                // <a:hlinkClick> always references its target through an r:id
                // relationship (DrawingML has no bare @anchor like w:hyperlink).
                // An absolute URI or relative target round-trips as a
                // TargetMode=External rel; a fragment "#anchor" round-trips as an
                // internal (isExternal=false) rel with Target="#anchor", matching
                // the run-level hyperlink Add path's fragment handling.
                bool isFragment = linkVal.StartsWith('#');
                Uri? linkUri;
                if (isFragment)
                {
                    linkUri = new Uri(linkVal, UriKind.Relative);
                }
                else if (Uri.TryCreate(linkVal, UriKind.Absolute, out linkUri))
                {
                    // BUG-DUMP-PIC-UNSAFE-LINK: a clickable image whose click-link
                    // carries an unsafe scheme (javascript:/data:/vbscript: — common
                    // in web-exported docs that wrap thumbnails in javascript:popUp())
                    // must NOT abort the whole picture. The link is a decoration on
                    // the image; RequireSafeScheme used to throw here and the entire
                    // `add picture` op failed, DROPPING the image — silently losing
                    // content and reflowing the page (lost pages on dump→batch). Word
                    // never executes a javascript: link anyway, so drop just the
                    // unsafe link and keep the image (same security outcome as the
                    // throw, without sacrificing the picture).
                    if (!Core.HyperlinkUriValidator.IsSafeScheme(linkVal))
                        linkUri = null;
                }
                else
                {
                    Uri.TryCreate(linkVal, UriKind.Relative, out linkUri);
                }
                if (linkUri != null)
                {
                    var linkRelId = imgHostPart switch
                    {
                        MainDocumentPart mdp => mdp.AddHyperlinkRelationship(linkUri, !isFragment).Id,
                        HeaderPart hp => hp.AddHyperlinkRelationship(linkUri, !isFragment).Id,
                        FooterPart fp => fp.AddHyperlinkRelationship(linkUri, !isFragment).Id,
                        FootnotesPart fnp => fnp.AddHyperlinkRelationship(linkUri, !isFragment).Id,
                        EndnotesPart enp => enp.AddHyperlinkRelationship(linkUri, !isFragment).Id,
                        WordprocessingCommentsPart cp => cp.AddHyperlinkRelationship(linkUri, !isFragment).Id,
                        _ => mainPart.AddHyperlinkRelationship(linkUri, !isFragment).Id,
                    };
                    picCNvPr.HyperlinkOnClick = new A.HyperlinkOnClick { Id = linkRelId };
                }
            }
        }

        string resultPath;
        Paragraph imgPara;
        if (parent is Paragraph existingPara)
        {
            // Use ChildElements for index lookup to match ResolveAnchorPosition
            // (which counts pPr). If index points at pPr, clamp forward.
            var imgChildren = existingPara.ChildElements.ToList();
            if (index.HasValue && index.Value < imgChildren.Count)
            {
                var refElement = imgChildren[index.Value];
                if (refElement is ParagraphProperties)
                {
                    if (index.Value + 1 < imgChildren.Count)
                        existingPara.InsertBefore(imgRun, imgChildren[index.Value + 1]);
                    else
                        existingPara.AppendChild(imgRun);
                }
                else
                {
                    existingPara.InsertBefore(imgRun, refElement);
                }
            }
            else
            {
                existingPara.AppendChild(imgRun);
            }
            imgPara = existingPara;
            // CONSISTENCY(run-path-index): align the returned r[N] index with
            // navigation's r[N] resolution, which uses Descendants<Run>() and
            // skips comment-reference runs. GetAllRuns encapsulates both rules.
            var imgRunIdx = PathIndex.FromArrayIndex(GetAllRuns(existingPara).IndexOf(imgRun));
            // CONSISTENCY(para-path-canonical): canonicalize to paraId-form.
            resultPath = $"{ReplaceTrailingParaSegment(parentPath, existingPara)}/r[{imgRunIdx}]";
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
                AssignParaId(imgPara);
                // Prevent fixed line spacing (inherited from Normal style) from
                // clipping the image to the text line height.
                imgPara.PrependChild(new ParagraphProperties(
                    new SpacingBetweenLines { Line = "240", LineRule = LineSpacingRuleValues.Auto }));
                imgCell.AppendChild(imgPara);
            }
            var imgPIdx = PathIndex.FromArrayIndex(imgCell.Elements<Paragraph>().ToList().IndexOf(imgPara));
            resultPath = $"{parentPath}/{BuildParaPathSegment(imgPara, imgPIdx)}";
        }
        else
        {
            imgPara = new Paragraph(imgRun);
            AssignParaId(imgPara);
            // Prevent fixed line spacing (inherited from Normal style) from
            // clipping the image to the text line height.
            imgPara.PrependChild(new ParagraphProperties(
                new SpacingBetweenLines { Line = "240", LineRule = LineSpacingRuleValues.Auto }));

            // Use ChildElements for index lookup so that tables and sectPr
            // siblings do not shift the effective insertion position. This
            // matches ResolveAnchorPosition, which computes anchor indices
            // against ChildElements.
            var allChildren = parent.ChildElements.ToList();
            if (index.HasValue && index.Value < allChildren.Count)
            {
                var refElement = allChildren[index.Value];
                parent.InsertBefore(imgPara, refElement);
                var imgPIdx = PathIndex.FromArrayIndex(parent.Elements<Paragraph>().ToList().IndexOf(imgPara));
                resultPath = $"{parentPath}/{BuildParaPathSegment(imgPara, imgPIdx)}";
            }
            else
            {
                AppendToParent(parent, imgPara);
                var imgPIdx = parent.Elements<Paragraph>().Count();
                resultPath = $"{parentPath}/{BuildParaPathSegment(imgPara, imgPIdx)}";
            }
        }

        // Apply run-level properties carried on the picture's Format (lang.*,
        // noproof, bold, color, …). Sources often stamp <w:lang> / <w:noProof>
        // on the picture's run — without this pass they were silently dropped
        // on replay, surfacing as phantom keys-missing on the second dump.
        // Iterate AddPicture's full property bag and apply anything the run-
        // formatting helper recognises; AddPicture-specific keys (width,
        // height, alt, name, wrap, anchor, …) are not in the helper's
        // vocabulary so they pass through untouched.
        OpenXmlCompositeElement? imgRunRPr = null;
        // ACCOUNTING(handler-as-truth): the foreach below dereferences the
        // dict via IEnumerable.GetEnumerator on the Dictionary static type,
        // which bypasses TrackingPropertyDictionary's overridden enumerator.
        // The crop keys are consumed by the `lk is "crop"…` branch but never
        // get marked as read → false-positive UNSUPPORTED warning. Explicitly
        // probe each so the tracker records them.
        properties.TryGetValue("crop", out _);
        properties.TryGetValue("cropLeft", out _);
        properties.TryGetValue("cropRight", out _);
        properties.TryGetValue("cropTop", out _);
        properties.TryGetValue("cropBottom", out _);
        foreach (var (key, value) in properties)
        {
            var lk = key.ToLowerInvariant();
            if (lk is "width" or "height" or "alt" or "name" or "src"
                or "wrap" or "anchor" or "hposition" or "vposition"
                or "hrelative" or "vrelative" or "halign" or "valign" or "behindtext"
                or "tooltip" or "tgtframe" or "tgtframe" or "history" or "url"
                or "relid" or "id" or "contenttype" or "filesize"
                or "src.svg"
                // BUG-DUMP-R45-4: verbatim-XML props applied above, not rPr keys.
                or "blipeffects" or "speffects")
                continue;
            // BUG-DUMP-CROP: crop is a blipFill property, not a run-rPr key —
            // ApplyRunFormatting would silently drop it, so the dump→batch
            // `add picture --prop crop=…` op (emitted by CreateImageNode's crop
            // readback) never re-applied the source rectangle. Route it through
            // the shared writer that Set uses.
            if (lk is "crop" or "cropleft" or "cropright" or "croptop" or "cropbottom")
            {
                var blipFillAdd = imgRun.GetFirstChild<Drawing>()
                    ?.Descendants<DocumentFormat.OpenXml.Drawing.Pictures.BlipFill>().FirstOrDefault();
                if (blipFillAdd != null) ApplyCropToBlipFill(blipFillAdd, key, value);
                continue;
            }
            imgRunRPr ??= imgRun.GetFirstChild<RunProperties>() ?? imgRun.PrependChild(new RunProperties());
            ApplyRunFormatting(imgRunRPr, key, value);
        }

        // BUG-DUMP11-06: a hyperlink-wrapped picture's `anchor` attr (the
        // Word-level <w:hyperlink w:anchor="bookmark"> wrapping) round-trips
        // by re-wrapping the inserted Run in a fresh Hyperlink. Navigation's
        // run-parent-is-hyperlink branch already surfaces the anchor on the
        // picture node. Pass-through the optional metadata attrs (tooltip /
        // tgtFrame / history / url) for symmetry with AddHyperlink.
        if (hyperlinkAnchorName != null)
        {
            var hlWrap = new Hyperlink { Anchor = hyperlinkAnchorName };
            if (properties.TryGetValue("tooltip", out var picTip)) hlWrap.Tooltip = picTip;
            if ((properties.TryGetValue("tgtFrame", out var picTgt)
                 || properties.TryGetValue("tgtframe", out picTgt))
                && !string.IsNullOrEmpty(picTgt))
                hlWrap.TargetFrame = picTgt;
            if (properties.TryGetValue("history", out var picHist) && IsTruthy(picHist))
                hlWrap.History = OnOffValue.FromBoolean(true);

            var imgRunParent = imgRun.Parent;
            if (imgRunParent != null)
            {
                // Replace the run in-place with a Hyperlink wrapper so
                // sibling order and the resultPath (which addresses the run
                // via Descendants<Run>()) remain valid.
                imgRun.InsertAfterSelf(hlWrap);
                imgRun.Remove();
                hlWrap.AppendChild(imgRun);
            }
        }

        return resultPath;
    }

    // ==================== OLE Object Insertion ====================
    //
    // Inserts an <w:object> wrapper containing:
    //   1. VML shapetype _x0000_t75 (picture frame, well-known shape ID)
    //   2. VML v:shape bound to an icon preview ImagePart
    //   3. o:OLEObject naming the ProgID and referencing an
    //      EmbeddedObjectPart / EmbeddedPackagePart (the binary payload)
    //
    // Defaults are tuned so callers can just say `--type ole --prop src=...`:
    //   - ProgID auto-detected from src extension (via OleHelper)
    //   - Backing part kind auto-chosen (Package for .docx/.xlsx/.pptx, Object otherwise)
    //   - Icon preview = tiny PNG placeholder
    //   - Dimensions default to 2in × 0.75in (matches Office's show-as-icon frame)
    //
    // Caller can override: progId, width, height, icon (png/jpg/emf file path),
    // display (icon|content). display=content flips DrawAspect to "Content".
    // dump→batch round-trip for an ActiveX form-control run (<w:object> hosting
    // <w:control r:id> + a VML preview <v:imagedata r:id> — no o:OLEObject).
    // props carry the verbatim <w:r> XML plus one part{N}.relId/part{N}.data
    // pair per package part the object references (preview image, activeX
    // persistence XML) and part{N}.child{M}.* for parts nested under those
    // (the activeX binary blob). Parts are recreated with FRESH relationship
    // ids — the source ids would collide with the rebuilt main part's existing
    // rels — and the run XML's r:id refs are rewritten to match. Child parts
    // keep their SOURCE rel ids: they are scoped to the freshly created parent
    // part, so the verbatim part bytes' internal refs resolve untouched.
    // Unified verbatim part-owning carrier (`add inlinedparts`) — supersedes the
    // former per-element carrier verbs (chartpart / diagram / vmlshape /
    // drawingshape / activex), which differed ONLY in a runXml marker check and
    // all delegated here. The runXml is the verbatim run whose drawing/pict/control
    // references its parts via r:id / r:dm / r:embed / ...; part{N}.* payloads carry
    // every part the element owns (and their children + external rels). The element
    // kind is self-evident from runXml + the part content types — CreateInlinedPart
    // routes by content type, never by verb — so a single verb covers all of them.
    // Used only by dump→batch (machine-produced); the old verb names stay accepted
    // as input aliases. The marker check is now the union of the former per-verb ones.
    private string AddInlinedPartsRun(OpenXmlElement parent, string parentPath, Dictionary<string, string> properties, string opName)
    {
        properties ??= new Dictionary<string, string>();
        if (!properties.TryGetValue("runXml", out var marker) || string.IsNullOrEmpty(marker)
            || !(marker.Contains("c:chart", StringComparison.Ordinal)
                 || marker.Contains("relIds", StringComparison.Ordinal)
                 || marker.Contains("<w:pict", StringComparison.Ordinal)
                 || marker.Contains("<w:drawing", StringComparison.Ordinal)
                 || marker.Contains("<w:control", StringComparison.Ordinal)))
            throw new ArgumentException(
                "inlinedparts requires --prop runXml containing a part-owning element "
                + "(<c:chart>, <dgm:relIds>, <w:pict>, <w:drawing> or <w:control>)");
        var runXml = properties["runXml"];
        var mainPart = _doc.MainDocumentPart!;
        // CONSISTENCY(host-part-rel): same routing as AddOle — parts referenced
        // from a header/footer-hosted run must attach to that part.
        OpenXmlPart hostPart = mainPart;
        {
            var headerAncestor = parent as Header ?? parent.Ancestors<Header>().FirstOrDefault();
            if (headerAncestor != null)
            {
                var hp = mainPart.HeaderParts.FirstOrDefault(p => ReferenceEquals(p.Header, headerAncestor));
                if (hp != null) hostPart = hp;
            }
            else
            {
                var footerAncestor = parent as Footer ?? parent.Ancestors<Footer>().FirstOrDefault();
                if (footerAncestor != null)
                {
                    var fp = mainPart.FooterParts.FirstOrDefault(p => ReferenceEquals(p.Footer, footerAncestor));
                    if (fp != null) hostPart = fp;
                }
            }
        }

        var rewriteRelIds = MaterializeInlinedParts(hostPart, properties, opName);

        var rewritten = rewriteRelIds(runXml);

        var axRun = new Run(rewritten);

        string resultPath;
        if (parent is Paragraph axPara)
        {
            axPara.AppendChild(axRun);
            var axRunIdx = PathIndex.FromArrayIndex(GetAllRuns(axPara).IndexOf(axRun));
            // CONSISTENCY(para-path-canonical): canonicalize to paraId-form.
            resultPath = $"{ReplaceTrailingParaSegment(parentPath, axPara)}/r[{axRunIdx}]";
        }
        else if (parent is TableCell axCell)
        {
            var firstCellPara = axCell.Elements<Paragraph>().FirstOrDefault();
            Paragraph hostPara;
            if (firstCellPara != null && !firstCellPara.Elements<Run>().Any())
            {
                firstCellPara.AppendChild(axRun);
                hostPara = firstCellPara;
            }
            else
            {
                hostPara = new Paragraph(axRun);
                AssignParaId(hostPara);
                axCell.AppendChild(hostPara);
            }
            var axPIdx = PathIndex.FromArrayIndex(axCell.Elements<Paragraph>().ToList().IndexOf(hostPara));
            var axCellRunIdx = PathIndex.FromArrayIndex(GetAllRuns(hostPara).IndexOf(axRun));
            resultPath = $"{parentPath}/{BuildParaPathSegment(hostPara, axPIdx)}/r[{axCellRunIdx}]";
        }
        else
        {
            var hostPara = new Paragraph(axRun);
            AssignParaId(hostPara);
            AppendToParent(parent, hostPara);
            var axPIdx = PathIndex.FromArrayIndex(parent.Elements<Paragraph>().ToList().IndexOf(hostPara));
            resultPath = $"{parentPath}/{BuildParaPathSegment(hostPara, axPIdx)}/r[1]";
        }
        return resultPath;
    }

    // Shared by AddInlinedPartsRun and the sdtXml carrier in AddSdt: create
    // every part{N} (with part{N}.child{M}) and ext{N} relationship on
    // <paramref name="hostPart"/>, then return the two-phase rel-id rewriter
    // mapping every source id to its freshly assigned one.
    private Func<string, string> MaterializeInlinedParts(OpenXmlPart hostPart, Dictionary<string, string> properties, string opName)
    {
        // Pass 1: create every top-level part empty, so the complete old→new
        // id map exists before any content is written. A collected XML part's
        // bytes can reference a SIBLING part's host-part relationship (the
        // SmartArt data part's <dsp:dataModelExt relId> points at the
        // rendered-drawing part), so the rewrite below must cover all ids.
        var idMap = new List<(string OldId, string NewId)>();
        var pending = new List<(OpenXmlPart Part, byte[] Bytes, string Ct, int Pi)>();
        for (int pi = 1; properties.TryGetValue($"part{pi}.relId", out var oldRelId); pi++)
        {
            var dataUri = properties.GetValueOrDefault($"part{pi}.data");
            if (string.IsNullOrEmpty(oldRelId)
                || !OfficeCli.Core.OleHelper.TryDecodeDataUri(dataUri, out var bytes, out var ct)
                || bytes.Length == 0)
                throw new ArgumentException($"{opName} part{pi} requires relId and a non-empty data: URI");

            var created = CreateInlinedPart(hostPart, ct)
                ?? throw new ArgumentException($"{opName} part{pi}: unsupported content type '{ct}'");
            pending.Add((created, bytes, ct, pi));
            idMap.Add((oldRelId!, hostPart.GetIdOfPart(created)));
        }

        // External relationships (hyperlinks inside a VML textbox, linked
        // content): no part bytes — recreate the rel on the host part with a
        // fresh id and route the source id through the same rewrite map.
        const string HyperlinkRelType =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
        for (int ei = 1; properties.TryGetValue($"ext{ei}.relId", out var extOldId); ei++)
        {
            var extType = properties.GetValueOrDefault($"ext{ei}.type");
            var extTarget = properties.GetValueOrDefault($"ext{ei}.target");
            if (string.IsNullOrEmpty(extOldId) || string.IsNullOrEmpty(extType) || string.IsNullOrEmpty(extTarget))
                throw new ArgumentException($"{opName} ext{ei} requires relId, type and target");
            var extUri = new Uri(extTarget, UriKind.RelativeOrAbsolute);
            var newExtId = extType == HyperlinkRelType
                ? hostPart.AddHyperlinkRelationship(extUri, true).Id
                : hostPart.AddExternalRelationship(extType, extUri).Id;
            idMap.Add((extOldId!, newExtId));
        }

        // Two-phase rewrite (shared by the run XML and every inlined XML
        // part's bytes): a freshly assigned id can equal a *different* source
        // id still pending replacement, so route through unique placeholders.
        string RewriteRelIds(string xml)
        {
            for (int i = 0; i < idMap.Count; i++)
                xml = xml.Replace($"\"{idMap[i].OldId}\"", $"\"__OCLI_AXREL_{i}__\"", StringComparison.Ordinal);
            for (int i = 0; i < idMap.Count; i++)
                xml = xml.Replace($"\"__OCLI_AXREL_{i}__\"", $"\"{idMap[i].NewId}\"", StringComparison.Ordinal);
            return xml;
        }

        // Pass 2: feed content (XML parts get their host-part rel refs
        // rewritten; binary parts stay verbatim) and attach child parts.
        foreach (var (created, bytes, ct, pi) in pending)
        {
            var feedBytes = bytes;
            if (ct.EndsWith("+xml", StringComparison.OrdinalIgnoreCase))
            {
                var text = System.Text.Encoding.UTF8.GetString(bytes);
                var rewrittenText = RewriteRelIds(text);
                if (!ReferenceEquals(rewrittenText, text))
                    feedBytes = System.Text.Encoding.UTF8.GetBytes(rewrittenText);
            }
            using (var ms = new MemoryStream(feedBytes))
                created.FeedData(ms);

            for (int ci = 1; properties.TryGetValue($"part{pi}.child{ci}.relId", out var childRelId); ci++)
            {
                var childUri = properties.GetValueOrDefault($"part{pi}.child{ci}.data");
                if (string.IsNullOrEmpty(childRelId)
                    || !OfficeCli.Core.OleHelper.TryDecodeDataUri(childUri, out var cbytes, out var cct)
                    || cbytes.Length == 0)
                    throw new ArgumentException($"{opName} part{pi}.child{ci} requires relId and a non-empty data: URI");
                var childPart = CreateInlinedChildPart(created, cct, childRelId!)
                    ?? throw new ArgumentException($"{opName} part{pi}.child{ci}: unsupported content type '{cct}'");
                // BUG-DUMP-R71-USERSHAPES-IMG: recreate the child's OWN parts (a
                // chart userShapes drawing -> its image) BEFORE feeding the child
                // bytes, using the ORIGINAL rel id so the child's verbatim r:embed
                // resolves without rewriting. Without this the rebuilt drawing's
                // r:embed dangles ("relationship does not exist").
                for (int gi = 1; properties.TryGetValue($"part{pi}.child{ci}.gc{gi}.relId", out var gcRelId); gi++)
                {
                    var gcUri = properties.GetValueOrDefault($"part{pi}.child{ci}.gc{gi}.data");
                    if (string.IsNullOrEmpty(gcRelId)
                        || !OfficeCli.Core.OleHelper.TryDecodeDataUri(gcUri, out var gcbytes, out var gcct)
                        || gcbytes.Length == 0)
                        throw new ArgumentException($"{opName} part{pi}.child{ci}.gc{gi} requires relId and a non-empty data: URI");
                    var gcPart = CreateInlinedChildPart(childPart, gcct, gcRelId!)
                        ?? throw new ArgumentException($"{opName} part{pi}.child{ci}.gc{gi}: unsupported content type '{gcct}'");
                    using var gcms = new MemoryStream(gcbytes);
                    gcPart.FeedData(gcms);
                }
                using var cms = new MemoryStream(cbytes);
                childPart.FeedData(cms);
            }

            // Per-part external rels (e.g. a chart's <c:externalData r:id> ->
            // external oleObject workbook). The part's bytes reference these
            // ids verbatim (RewriteRelIds only touches host-part ids), so
            // recreate each ON the created part with its ORIGINAL id.
            for (int xi = 1; properties.TryGetValue($"part{pi}.ext{xi}.relId", out var pExtId); xi++)
            {
                var pExtType = properties.GetValueOrDefault($"part{pi}.ext{xi}.type");
                var pExtTarget = properties.GetValueOrDefault($"part{pi}.ext{xi}.target");
                if (string.IsNullOrEmpty(pExtId) || string.IsNullOrEmpty(pExtType) || string.IsNullOrEmpty(pExtTarget))
                    throw new ArgumentException($"{opName} part{pi}.ext{xi} requires relId, type and target");
                var pExtUri = new Uri(pExtTarget, UriKind.RelativeOrAbsolute);
                if (pExtType == HyperlinkRelType)
                    created.AddHyperlinkRelationship(pExtUri, true, pExtId!);
                else
                    created.AddExternalRelationship(pExtType, pExtUri, pExtId!);
            }
        }

        return RewriteRelIds;
    }

    // Part factory for the inlined-parts carriers. Top-level parts hang off the
    // run's host part (main document / header / footer); returns null for an
    // unrecognized content type so the caller surfaces a clear error.
    private static OpenXmlPart? CreateInlinedPart(OpenXmlPart hostPart, string ct) => ct switch
    {
        "application/vnd.ms-office.activeX+xml"
            => hostPart.AddNewPart<EmbeddedControlPersistencePart>(ct, null),
        "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml"
            => hostPart.AddNewPart<DiagramDataPart>(ct, null),
        "application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml"
            => hostPart.AddNewPart<DiagramLayoutDefinitionPart>(ct, null),
        "application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml"
            => hostPart.AddNewPart<DiagramStylePart>(ct, null),
        "application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml"
            => hostPart.AddNewPart<DiagramColorsPart>(ct, null),
        // The rendered-drawing part referenced from the data part's
        // dataModelExt extension — its relationship lives on the MAIN part
        // (ms diagramDrawing rel type), so it is a top-level part here.
        "application/vnd.ms-office.drawingml.diagramDrawing+xml"
            => hostPart.AddNewPart<DiagramPersistLayoutPart>(ct, null),
        _ when ct.StartsWith("image/", StringComparison.OrdinalIgnoreCase)
            => hostPart.AddNewPart<ImagePart>(ct, null),
        // BUG-DUMP-INKML: digital-ink (handwriting) parts (application/inkml+xml,
        // referenced from <w14:contentPart r:id>) are attached via a customXml
        // relationship — Word/the SDK reach them as a CustomXmlPart whose content
        // type is overridden to inkml+xml. The carrier previously had no arm for
        // this content type and aborted the whole inlined-parts step, dropping the
        // ink drawing. Recreate it as a CustomXmlPart with the source content type
        // so the customXml relationship + inkml payload round-trip.
        "application/inkml+xml"
            => hostPart.AddNewPart<CustomXmlPart>(ct, null),
        // BUG-DUMP-R55-VMLCHART: a VML shape / AlternateContent drawing embeds a
        // DrawingML chart (chart+xml, referenced by r:id). Without this the
        // inlined-parts materializer aborted the whole `add vmlshape` step and
        // dropped the chart. Its colour-style / style / embedded-package
        // children (when present) ride part{N}.child{M} via CreateInlinedChildPart.
        "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
            => hostPart.AddNewPart<ChartPart>(ct, null),
        "application/vnd.ms-office.chartcolorstyle+xml"
            => hostPart.AddNewPart<ChartColorStylePart>(ct, null),
        "application/vnd.ms-office.chartstyle+xml"
            => hostPart.AddNewPart<ChartStylePart>(ct, null),
        // BUG-DUMP-R40-VMLOLE: a VML shape / <o:OLEObject> embeds an OLE object
        // (legacy .xls/.doc/.ppt → application/vnd.ms-* via an `oleObject`
        // relationship) or a modern OOXML package (.xlsx/.docx/.pptx via a
        // `package` relationship). These reach the inlined-parts path as part{N}
        // and were unsupported, aborting the whole `add vmlshape` (and dropping
        // the embedded spreadsheet) on dump→batch. Route OLE-object content
        // types to an EmbeddedObjectPart (preserving the source content type so
        // .xls round-trips) and package content types to an EmbeddedPackagePart.
        _ when IsEmbeddedPackageContentType(ct)
            => hostPart.AddNewPart<EmbeddedPackagePart>(ct, null),
        _ when IsEmbeddedOleObjectContentType(ct)
            => hostPart.AddNewPart<EmbeddedObjectPart>(ct, null),
        _ => null,
    };

    // Modern OOXML formats embedded as an OLE package (rel type `package`).
    private static bool IsEmbeddedPackageContentType(string ct) =>
        ct is "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
           or "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.macroEnabled.main+xml"
           or "application/vnd.ms-excel.sheet.macroEnabled.12"
           or "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
           or "application/vnd.openxmlformats-officedocument.presentationml.presentation";

    // Legacy / generic OLE objects embedded via an `oleObject` relationship.
    private static bool IsEmbeddedOleObjectContentType(string ct) =>
        ct is "application/vnd.openxmlformats-officedocument.oleObject"
           or "application/vnd.ms-excel"
           or "application/msword"
           or "application/vnd.ms-powerpoint"
           or "application/vnd.ms-office"
        || ct.StartsWith("application/vnd.ms-", StringComparison.OrdinalIgnoreCase);

    // Child parts keep their SOURCE rel id (scoped to the freshly created
    // parent part, so the verbatim parent bytes' internal refs resolve).
    private static OpenXmlPart? CreateInlinedChildPart(OpenXmlPart parent, string ct, string relId) => ct switch
    {
        "application/vnd.ms-office.activeX" or "application/vnd.ms-office.activeX+xml"
            => parent.AddNewPart<EmbeddedControlPersistenceBinaryDataPart>(ct, relId),
        // The rendered-drawing child of a diagram data part — what Word
        // actually rasterizes for modern SmartArt.
        "application/vnd.ms-office.drawingml.diagramDrawing+xml"
            => parent.AddNewPart<DiagramPersistLayoutPart>(ct, relId),
        // BUG-DUMP-R55-VMLCHART: a chart part's own children (colour-style /
        // style / its embedded data package) keep their source rel id.
        "application/vnd.ms-office.chartcolorstyle+xml"
            => parent.AddNewPart<ChartColorStylePart>(ct, relId),
        "application/vnd.ms-office.chartstyle+xml"
            => parent.AddNewPart<ChartStylePart>(ct, relId),
        // BUG-DUMP-VMLCHART-THEMEOVERRIDE: a DrawingML chart child of a VML
        // shape can carry a <c:themeOverride> part (the chart's own theme
        // overrides — colours/fonts that differ from the document theme). It
        // reaches here as part{N}.child{M} with themeOverride+xml. Without this
        // arm CreateInlinedChildPart returned null, aborting the whole
        // `add vmlshape` step and dropping the shape's enclosed text runs.
        "application/vnd.openxmlformats-officedocument.themeOverride+xml"
            => parent.AddNewPart<ThemeOverridePart>(ct, relId),
        // BUG-DUMP-CHART-VERBATIM: a native chart's userShapes overlay drawing
        // (chartshapes+xml, the chartUserShapes child of the ChartPart — a
        // logo/annotation drawn on top of the chart). Carried as a chart child
        // so the verbatim `add chartpart` carrier round-trips it; without this
        // arm a chart with a userShapes overlay aborts the carrier and falls
        // back to the lossy typed rebuild.
        "application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml"
            => parent.AddNewPart<ChartDrawingPart>(ct, relId),
        // BUG-DUMP-VMLCHART-EMBEDPKG: a chart child of a VML shape can also be
        // the chart's embedded data workbook (spreadsheetml.sheet, the live
        // chart source) or a legacy OLE object. CreateInlinedPart already routes
        // these at the top level; mirror it for the child slot so the embedded
        // package round-trips instead of aborting the shape (and dropping its
        // text). Same content-type predicates so the two factories stay in sync.
        _ when IsEmbeddedPackageContentType(ct)
            => parent.AddNewPart<EmbeddedPackagePart>(ct, relId),
        // BUG-DUMP-CHART-OLEDATA: a native chart's <c:externalData r:id> can point
        // at a LEGACY embedded OLE workbook (relationship type oleObject →
        // oleObject1.bin) rather than a modern .xlsx package. The SDK's ChartPart
        // does not list EmbeddedObjectPart among its allowed children, so
        // AddNewPart<EmbeddedObjectPart> on a ChartPart throws "The part cannot be
        // added here" — failing the whole chart inlined-parts op and leaving the
        // chart with a dangling externalData ref (chart renders without its data).
        // Attach it via AddExtendedPart with the source oleObject relationship type
        // (which bypasses the typed-child constraint) keeping the source rel id so
        // the chart's verbatim <c:externalData r:id> resolves. Non-chart parents
        // keep the typed EmbeddedObjectPart (valid there, used by VML shapes).
        _ when IsEmbeddedOleObjectContentType(ct) && parent is ChartPart
            => parent.AddExtendedPart(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject",
                ct, ".bin", relId),
        _ when IsEmbeddedOleObjectContentType(ct)
            => parent.AddNewPart<EmbeddedObjectPart>(ct, relId),
        _ when ct.StartsWith("image/", StringComparison.OrdinalIgnoreCase)
            => parent.AddNewPart<ImagePart>(ct, relId),
        _ => null,
    };


    private string AddOle(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        properties ??= new Dictionary<string, string>();
        var srcPath = OfficeCli.Core.OleHelper.RequireSource(properties);
        OfficeCli.Core.OleHelper.WarnOnUnknownOleProps(properties);

        var mainPart = _doc.MainDocumentPart!;

        // Determine the host part that owns the parent element.
        // For /header[N] or /footer[N], the parent lives inside a
        // HeaderPart/FooterPart, so the embedded payload AND icon ImagePart
        // relationships must be attached to that part — not to
        // MainDocumentPart — otherwise OpenXmlValidator rejects the
        // cross-part r:id with a NullReferenceException.
        OpenXmlPart hostPart = mainPart;
        {
            var headerAncestor = parent as Header ?? parent.Ancestors<Header>().FirstOrDefault();
            if (headerAncestor != null)
            {
                var hp = mainPart.HeaderParts.FirstOrDefault(p => ReferenceEquals(p.Header, headerAncestor));
                if (hp != null) hostPart = hp;
            }
            else
            {
                var footerAncestor = parent as Footer ?? parent.Ancestors<Footer>().FirstOrDefault();
                if (footerAncestor != null)
                {
                    var fp = mainPart.FooterParts.FirstOrDefault(p => ReferenceEquals(p.Footer, footerAncestor));
                    if (fp != null) hostPart = fp;
                }
            }
        }

        // 1. Create the embedded binary payload part and rel id on the host part.
        // 2. Resolve ProgID.
        string embedRelId;
        string? progId;
        if (OfficeCli.Core.OleHelper.TryDecodeDataUri(srcPath, out var embedBytes, out var dataCt))
        {
            // dump→batch round-trip: src is a data: URI carrying the embedded
            // payload exactly as stored (raw package, or CFB-wrapped Ole10Native
            // for generic objects). The dump also forwards oleKind / contentType
            // / embedExt so we rebuild the same part class + content type +
            // target extension and feed the bytes verbatim (no extension sniffing,
            // no CFB re-wrap).
            var oleKind = properties.GetValueOrDefault("oleKind")
                ?? properties.GetValueOrDefault("olekind")
                ?? (dataCt.Contains("oleObject", StringComparison.OrdinalIgnoreCase) ? "object" : "package");
            var contentType = properties.GetValueOrDefault("contentType")
                ?? properties.GetValueOrDefault("contenttype")
                ?? (string.IsNullOrEmpty(dataCt) ? "application/octet-stream" : dataCt);
            var embedExt = properties.GetValueOrDefault("embedExt")
                ?? properties.GetValueOrDefault("embedext");
            (embedRelId, _) = OfficeCli.Core.OleHelper.AddEmbeddedPartFromBytes(
                hostPart, embedBytes, oleKind, contentType, embedExt);
            // BUG-DUMP-OLE-NOPROGID: ProgID is OPTIONAL in OOXML (an <o:OLEObject>
            // may omit it), so the dump omits the progId prop when the source has
            // none. Requiring it here threw on that valid input, failing the batch
            // op and silently dropping the OLE object. Accept a missing progId and
            // emit the <o:OLEObject> WITHOUT a ProgID attribute, mirroring the
            // source. (An explicit progId is still validated.)
            progId = properties.GetValueOrDefault("progId")
                ?? properties.GetValueOrDefault("progid");
            if (!string.IsNullOrEmpty(progId))
                OfficeCli.Core.OleHelper.ValidateProgId(progId);
        }
        else
        {
            (embedRelId, _) = OfficeCli.Core.OleHelper.AddEmbeddedPart(hostPart, srcPath, _filePath);
            // ProgID: explicit > auto-detected from extension.
            progId = OfficeCli.Core.OleHelper.ResolveProgId(properties, srcPath);
        }

        // 3. Create the icon preview ImagePart on the host part (same part
        //    that owns the OLE element itself). Attaching to MainDocumentPart
        //    when the OLE lives in a header/footer would produce a dangling
        //    cross-part relationship — see host part resolution above.
        var (_, iconRelId) = OfficeCli.Core.OleHelper.CreateIconPart(hostPart, properties);

        // 4. Dimensions. Word VML shapes take points in their style string.
        //    Defaults match OleHelper's 2in × 0.75in icon frame.
        long cxEmu = properties.TryGetValue("width", out var wStr)
            ? ParseEmu(wStr) : OfficeCli.Core.OleHelper.DefaultOleWidthEmu;
        long cyEmu = properties.TryGetValue("height", out var hStr)
            ? ParseEmu(hStr) : OfficeCli.Core.OleHelper.DefaultOleHeightEmu;
        // EMU → points (914400 EMU/inch, 72 points/inch).
        double cxPt = cxEmu / EmuConverter.EmuPerPointF;
        double cyPt = cyEmu / EmuConverter.EmuPerPointF;
        // Twips for w:dxaOrig/w:dyaOrig (20 twips/point). A round-tripped OLE
        // carries the source's native object box verbatim (dxaOrig/dyaOrig props)
        // so a floating OLE keeps Word's original size instead of being rescaled
        // from the display frame.
        string cxTwips = properties.TryGetValue("dxaOrig", out var dxaO) && !string.IsNullOrEmpty(dxaO)
            ? dxaO : ((long)(cxPt * 20)).ToString(System.Globalization.CultureInfo.InvariantCulture);
        string cyTwips = properties.TryGetValue("dyaOrig", out var dyaO) && !string.IsNullOrEmpty(dyaO)
            ? dyaO : ((long)(cyPt * 20)).ToString(System.Globalization.CultureInfo.InvariantCulture);

        // 5. DrawAspect: "Icon" (default) or "Content" (live preview).
        // Strict validation: unknown values throw rather than silently
        // falling back to Icon — see OleHelper.NormalizeOleDisplay.
        var display = OfficeCli.Core.OleHelper.NormalizeOleDisplay(
            properties.GetValueOrDefault("display", "icon"));
        var drawAspect = display == "content" ? "Content" : "Icon";

        // 6. ObjectID: VML requires a unique "_nnnnnnnnnn" token.
        //    Count existing OLE objects and assign a monotonic id so two
        //    OLEs added within the same wallclock second don't collide
        //    (the old scheme used ToUnixTimeSeconds()).
        var existingOleCount = mainPart.Document?.Body?.Descendants<EmbeddedObject>().Count() ?? 0;
        var oleSeq = existingOleCount + 1;
        var objectId = "_" + (1000000000 + oleSeq);

        // 7. Build the w:object XML. The shapetype + shape + OLEObject
        //    triple is the canonical form Word itself writes for OLE.
        //    ShapeID must also be unique per OLE in the document — base it
        //    on the OLE sequence (not NextDocPropId, which is shared with
        //    Drawing DocProperties and can collide). D4 gives 9999 slots.
        var shapeId = $"_x0000_i1{oleSeq:D4}";

        // Optional friendly name → v:shape alt="..." attribute.
        // CONSISTENCY(ole-name): the VML CT_OleObject complex type has no
        // Name attribute (valid attrs: Type/ProgID/ShapeID/DrawAspect/
        // ObjectID/r:id/UpdateMode/LinkType/LockedField/FieldCodes — see
        // DocumentFormat.OpenXml.Vml.Office.OleObject). Writing Name= on
        // o:OLEObject produces a schema validation error. Use the
        // surrounding v:shape element's "alt" attribute (Alternate Text,
        // closest semantic match in VML) for the friendly name. Get reads
        // it back from the same place, preserving Format["name"] round-trip.
        var shapeAltAttr = "";
        if (properties.TryGetValue("name", out var oleName) && !string.IsNullOrEmpty(oleName))
            shapeAltAttr = $" alt=\"{System.Security.SecurityElement.Escape(oleName)}\"";

        // CONSISTENCY(ole-shapetype-dedup): v:shapetype id="_x0000_t75" must be
        // unique across the whole document.xml — OOXML validation rejects
        // duplicate shapetype ids. If the document already has an
        // _x0000_t75 shapetype (left over from a prior picture/OLE insert),
        // skip re-emitting it and reference the existing one from v:shape.
        var shapetypeAlreadyExists = false;
        foreach (var existingObj in mainPart.Document?.Body?.Descendants<EmbeddedObject>() ?? Enumerable.Empty<EmbeddedObject>())
        {
            foreach (var st in existingObj.Descendants().Where(e => e.LocalName == "shapetype"))
            {
                var idAttr = st.GetAttributes().FirstOrDefault(a => a.LocalName == "id");
                if (idAttr.Value == "_x0000_t75") { shapetypeAlreadyExists = true; break; }
            }
            if (shapetypeAlreadyExists) break;
        }

        var shapetypeXml = shapetypeAlreadyExists ? "" : """
<v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f">
<v:stroke joinstyle="miter"/>
<v:formulas>
<v:f eqn="if lineDrawn pixelLineWidth 0"/>
<v:f eqn="sum @0 1 0"/>
<v:f eqn="sum 0 0 @1"/>
<v:f eqn="prod @2 1 2"/>
<v:f eqn="prod @3 21600 pixelWidth"/>
<v:f eqn="prod @3 21600 pixelHeight"/>
<v:f eqn="sum @0 0 1"/>
<v:f eqn="prod @6 1 2"/>
<v:f eqn="prod @7 21600 pixelWidth"/>
<v:f eqn="sum @8 21600 0"/>
<v:f eqn="prod @7 21600 pixelHeight"/>
<v:f eqn="sum @10 21600 0"/>
</v:formulas>
<v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/>
<o:lock v:ext="edit" aspectratio="t"/>
</v:shapetype>
""";

        // A round-tripped FLOATING OLE carries its source v:shape style verbatim
        // (position:absolute + margin + z-index + wrap), keeping the object out of
        // the text flow. An inline OLE (no shapeStyle) rebuilds the bare
        // width/height frame and marks itself o:ole="" — the floating shape omits
        // o:ole="" to match what Word writes for an absolutely-positioned OLE.
        var hasFloatStyle = properties.TryGetValue("shapeStyle", out var floatStyle)
            && !string.IsNullOrEmpty(floatStyle);
        var shapeStyleAttr = hasFloatStyle
            ? System.Security.SecurityElement.Escape(floatStyle)
            : $"width:{cxPt.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture)}pt;height:{cyPt.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture)}pt";
        var oleAttr = hasFloatStyle ? "" : " o:ole=\"\"";

        // BUG-DUMP-OLECROP: re-apply the VML <v:imagedata> crop rectangle captured
        // by the dump ("cropleft:Nf;cropright:Nf;…"), splicing each present side
        // back as a verbatim VML attribute so the preview round-trips cropped (an
        // uncropped EMF preview renders larger and pushes later pages down).
        var imageCropAttrs = "";
        if (properties.TryGetValue("crop", out var oleCrop) && !string.IsNullOrEmpty(oleCrop))
        {
            var cropSb = new System.Text.StringBuilder();
            foreach (var seg in oleCrop.Split(';', StringSplitOptions.RemoveEmptyEntries))
            {
                var kv = seg.Split(':', 2);
                if (kv.Length != 2) continue;
                var ck = kv[0].Trim().ToLowerInvariant();
                if (ck is not ("cropleft" or "croptop" or "cropright" or "cropbottom")) continue;
                cropSb.Append(' ').Append(ck).Append("=\"")
                      .Append(System.Security.SecurityElement.Escape(kv[1].Trim())).Append('"');
            }
            imageCropAttrs = cropSb.ToString();
        }

        // ProgID is optional (BUG-DUMP-OLE-NOPROGID): omit the attribute entirely
        // when the source had none, rather than emitting ProgID="".
        var progIdAttr = string.IsNullOrEmpty(progId)
            ? ""
            : $" ProgID=\"{System.Security.SecurityElement.Escape(progId)}\"";
        var oleXml = $"""
<w:object xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" w:dxaOrig="{cxTwips}" w:dyaOrig="{cyTwips}">
{shapetypeXml}<v:shape id="{shapeId}" type="#_x0000_t75" style="{shapeStyleAttr}"{oleAttr}{shapeAltAttr}>
<v:imagedata r:id="{iconRelId}" o:title=""{imageCropAttrs}/>
</v:shape>
<o:OLEObject Type="Embed"{progIdAttr} ShapeID="{shapeId}" DrawAspect="{drawAspect}" ObjectID="{objectId}" r:id="{embedRelId}"/>
</w:object>
""";
        var oleObject = new EmbeddedObject(oleXml);

        // 8. Wrap in a Run and insert it, mirroring the AddPicture positional logic.
        var oleRun = new Run(oleObject);

        // BUG-DUMP-OLERPR: re-apply the source OLE run's <w:rPr> (forwarded by
        // TryEmitOleRun). The run wrapping <w:object> can carry run typography —
        // most visibly a <w:bdr> border box, also rFonts/sz that set the host
        // line height — and a bare rebuilt run dropped it, nudging following
        // lines and reflowing the page. Mirrors the breakRunRpr re-apply.
        if (properties.TryGetValue("runRpr", out var oleRpr)
            && !string.IsNullOrWhiteSpace(oleRpr)
            && oleRpr.Contains("rPr", StringComparison.Ordinal))
        {
            try { oleRun.PrependChild(new RunProperties(oleRpr)); } catch { /* malformed: skip */ }
        }

        // If the parent is a block-level SDT, insert into its SdtContentBlock
        // (creating it if missing) instead of appending directly to the SdtBlock.
        // Direct SdtBlock child paragraphs violate the schema and get silently
        // stripped by Word on reload — which previously broke OLE persistence
        // across reopen when added inside an SDT container. See
        // OleTestTeamRound6.Word_OleInsideSdt_QueryFindsOle.
        if (parent is SdtBlock sdtBlockParent)
        {
            var contentBlock = sdtBlockParent.GetFirstChild<SdtContentBlock>();
            if (contentBlock == null)
            {
                contentBlock = new SdtContentBlock();
                sdtBlockParent.AppendChild(contentBlock);
            }
            parent = contentBlock;
        }
        // Inline SDT runs live inside a w:p parent: route the OLE to that
        // surrounding paragraph so insertion follows the normal run path.
        else if (parent is SdtRun sdtRunParent)
        {
            var contentRun = sdtRunParent.GetFirstChild<SdtContentRun>();
            if (contentRun != null)
                contentRun.AppendChild(oleRun);
            else
                sdtRunParent.AppendChild(new SdtContentRun(oleRun));
            var parentParaInline = sdtRunParent.Ancestors<Paragraph>().FirstOrDefault();
            if (parentParaInline != null)
            {
                var runs = GetAllRuns(parentParaInline);
                var runIdxInline = PathIndex.FromArrayIndex(runs.IndexOf(oleRun));
                // CONSISTENCY(para-path-canonical): canonicalize when the
                // SDT lives directly inside a paragraph (parentPath ends in
                // /p[...]); otherwise (SDT in a cell) parentPath does not
                // end in /p[...] and ReplaceTrailingParaSegment is a no-op.
                return $"{ReplaceTrailingParaSegment(parentPath, parentParaInline)}/r[{runIdxInline}]";
            }
            return parentPath + "/r[1]";
        }

        string resultPath;
        if (parent is Paragraph existingPara)
        {
            // Use ChildElements for index lookup to match ResolveAnchorPosition.
            var oleChildren = existingPara.ChildElements.ToList();
            if (index.HasValue && index.Value < oleChildren.Count)
            {
                var refElement = oleChildren[index.Value];
                if (refElement is ParagraphProperties)
                {
                    if (index.Value + 1 < oleChildren.Count)
                        existingPara.InsertBefore(oleRun, oleChildren[index.Value + 1]);
                    else
                        existingPara.AppendChild(oleRun);
                }
                else
                {
                    existingPara.InsertBefore(oleRun, refElement);
                }
            }
            else
            {
                existingPara.AppendChild(oleRun);
            }
            var olePIdx = 1;
            foreach (var para in parent.Parent?.Elements<Paragraph>() ?? Enumerable.Empty<Paragraph>())
            {
                if (ReferenceEquals(para, existingPara)) break;
                olePIdx++;
            }
            var oleRunIdx = PathIndex.FromArrayIndex(GetAllRuns(existingPara).IndexOf(oleRun));
            // CONSISTENCY(para-path-canonical): canonicalize to paraId-form.
            resultPath = $"{ReplaceTrailingParaSegment(parentPath, existingPara)}/r[{oleRunIdx}]";
        }
        else if (parent is TableCell oleCell)
        {
            var firstCellPara = oleCell.Elements<Paragraph>().FirstOrDefault();
            Paragraph olePara;
            if (firstCellPara != null && !firstCellPara.Elements<Run>().Any())
            {
                firstCellPara.AppendChild(oleRun);
                olePara = firstCellPara;
            }
            else
            {
                olePara = new Paragraph(oleRun);
                AssignParaId(olePara);
                oleCell.AppendChild(olePara);
            }
            var olePIdx = PathIndex.FromArrayIndex(oleCell.Elements<Paragraph>().ToList().IndexOf(olePara));
            // CONSISTENCY(ole-run-path): same /r[1] suffix as the else branch
            // below — the OLE run is the addressable target, not the paragraph.
            var oleCellRunIdx = PathIndex.FromArrayIndex(GetAllRuns(olePara).IndexOf(oleRun));
            resultPath = $"{parentPath}/{BuildParaPathSegment(olePara, olePIdx)}/r[{oleCellRunIdx}]";
        }
        else
        {
            var olePara = new Paragraph(oleRun);
            AssignParaId(olePara);
            var allChildren = parent.ChildElements.ToList();
            if (index.HasValue && index.Value < allChildren.Count)
            {
                var refElement = allChildren[index.Value];
                parent.InsertBefore(olePara, refElement);
            }
            else
            {
                AppendToParent(parent, olePara);
            }
            var olePIdx = PathIndex.FromArrayIndex(parent.Elements<Paragraph>().ToList().IndexOf(olePara));
            // Return the /r[1] address so callers can Set/Get/Remove the
            // OLE run directly. Picture's Add returns a paragraph-level
            // path because the paragraph Set is meaningful (font, style);
            // for OLE, the only interesting target is the run itself.
            resultPath = $"{parentPath}/{BuildParaPathSegment(olePara, olePIdx)}/r[1]";
        }
        // BUG-DUMP-DELOLE: preserve a tracked-change wrapper on the rebuilt OLE run
        // (revision.type=del/ins/moveFrom/moveTo). Mirrors AddBreak — wrap after the
        // result path is computed; no-op when no revision.type is present. A deleted
        // figure that loses its <w:del> resurrects as a live full-size object and
        // pushes the following content onto new pages.
        WrapRunsInRevision(new List<Run> { oleRun }, properties);
        return resultPath;
    }
}
