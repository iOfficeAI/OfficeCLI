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
    private string AddChart(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
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

        // Dimensions (default: 15cm x 10cm)
        long chartCx = properties.TryGetValue("width", out var chartWStr) ? ParseEmu(chartWStr) : 5400000;
        long chartCy = properties.TryGetValue("height", out var chStr) ? ParseEmu(chStr) : 3600000;

        var docPropId = NextDocPropId();
        var chartName = chartTitle ?? $"Chart {docPropId}";

        // Extended chart types (cx:chart) — funnel, treemap, sunburst, boxWhisker, histogram
        if (Core.ChartExBuilder.IsExtendedChartType(chartType))
        {
            var cxChartSpace = Core.ChartExBuilder.BuildExtendedChartSpace(
                chartType, chartTitle, categories, seriesData, properties);
            var extChartPart = chartMainPart.AddNewPart<ExtendedChartPart>();
            extChartPart.ChartSpace = cxChartSpace;
            extChartPart.ChartSpace.Save();

            var cxRelId = chartMainPart.GetIdOfPart(extChartPart);
            var cxChartRef = new DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing.RelId { Id = cxRelId };

            var cxInline = new DW.Inline(
                new DW.Extent { Cx = chartCx, Cy = chartCy },
                new DW.EffectExtent { LeftEdge = 0, TopEdge = 0, RightEdge = 0, BottomEdge = 0 },
                new DW.DocProperties { Id = docPropId, Name = chartName },
                new DW.NonVisualGraphicFrameDrawingProperties(),
                new A.Graphic(
                    new A.GraphicData(cxChartRef)
                    { Uri = "http://schemas.microsoft.com/office/drawing/2014/chartex" }
                )
            )
            {
                DistanceFromTop = 0U,
                DistanceFromBottom = 0U,
                DistanceFromLeft = 0U,
                DistanceFromRight = 0U
            };

            var cxRun = new Run(new Drawing(cxInline));
            Paragraph cxPara;
            if (parent is Paragraph existingCxPara)
            {
                existingCxPara.AppendChild(cxRun);
                cxPara = existingCxPara;
            }
            else
            {
                cxPara = new Paragraph(cxRun);
                AssignParaId(cxPara);
                AppendToParent(parent, cxPara);
            }

            var totalCharts = CountWordCharts(chartMainPart);
            return $"/chart[{totalCharts}]";
        }

        // Create ChartPart and build chart
        var chartPart = chartMainPart.AddNewPart<ChartPart>();
        chartPart.ChartSpace = Core.ChartHelper.BuildChartSpace(chartType, chartTitle, categories, seriesData, properties);

        // Apply deferred properties (axisTitle, dataLabels, etc.) via SetChartProperties
        // Must be called BEFORE Save() so the in-memory DOM is still available
        var deferredProps = properties
            .Where(kv => Core.ChartHelper.IsDeferredKey(kv.Key))
            .ToDictionary(kv => kv.Key, kv => kv.Value);
        if (deferredProps.Count > 0)
            Core.ChartHelper.SetChartProperties(chartPart, deferredProps);
        else
            chartPart.ChartSpace.Save();

        var chartRelId = chartMainPart.GetIdOfPart(chartPart);

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
            AssignParaId(chartPara);
            AppendToParent(parent, chartPara);
        }

        var totalChartIdx = CountWordCharts(chartMainPart);
        return $"/chart[{totalChartIdx}]";
    }

    private string AddPicture(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        if (!properties.TryGetValue("path", out var imgPath) && !properties.TryGetValue("src", out imgPath))
            throw new ArgumentException("'src' property is required for picture type");

        // Buffer the image bytes so we can both feed the image part and sniff
        // the native pixel dimensions for auto aspect-ratio calculations.
        var (rawStream, imgPartType) = OfficeCli.Core.ImageSource.Resolve(imgPath);
        using var rawStreamDispose = rawStream;
        using var imgStream = new MemoryStream();
        rawStream.CopyTo(imgStream);
        imgStream.Position = 0;

        var mainPart = _doc.MainDocumentPart!;
        var imagePart = mainPart.AddImagePart(imgPartType);
        imagePart.FeedData(imgStream);
        imgStream.Position = 0;
        var relId = mainPart.GetIdOfPart(imagePart);

        // Determine dimensions. When only one axis is supplied, compute the
        // other from the image's native pixel aspect ratio. When neither is
        // supplied, width defaults to 6 inches and height follows the aspect
        // ratio (or a 4 inch fallback when the image header cannot be read).
        bool hasWidth = properties.TryGetValue("width", out var widthStr);
        bool hasHeight = properties.TryGetValue("height", out var heightStr);
        long cxEmu = hasWidth ? ParseEmu(widthStr!) : 5486400;  // 6 inches fallback
        long cyEmu = hasHeight ? ParseEmu(heightStr!) : 3657600; // 4 inches fallback

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

        var altText = properties.GetValueOrDefault("alt", Path.GetFileName(imgPath));

        var imgDocPropId = NextDocPropId();
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
            imgRun = CreateAnchorImageRun(relId, cxEmu, cyEmu, altText, wrapType, hPos, vPos, hRel, vRel, behind, imgDocPropId);
        }
        else
        {
            imgRun = CreateImageRun(relId, cxEmu, cyEmu, altText, imgDocPropId);
        }

        string resultPath;
        Paragraph imgPara;
        if (parent is Paragraph existingPara)
        {
            // When --index N is supplied, insert before the Nth existing run
            // instead of always appending. Matches AddRun's index semantics.
            var runCount = existingPara.Elements<Run>().Count();
            if (index.HasValue && index.Value < runCount)
            {
                var refRun = existingPara.Elements<Run>().ElementAt(index.Value);
                existingPara.InsertBefore(imgRun, refRun);
            }
            else
            {
                existingPara.AppendChild(imgRun);
            }
            imgPara = existingPara;
            var imgRunIdx = existingPara.Elements<Run>().ToList().IndexOf(imgRun) + 1;
            resultPath = $"{parentPath}/r[{imgRunIdx}]";
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
            var imgPIdx = imgCell.Elements<Paragraph>().ToList().IndexOf(imgPara) + 1;
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
                var imgPIdx = parent.Elements<Paragraph>().ToList().IndexOf(imgPara) + 1;
                resultPath = $"{parentPath}/{BuildParaPathSegment(imgPara, imgPIdx)}";
            }
            else
            {
                AppendToParent(parent, imgPara);
                var imgPIdx = parent.Elements<Paragraph>().Count();
                resultPath = $"{parentPath}/{BuildParaPathSegment(imgPara, imgPIdx)}";
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
        var (embedRelId, _) = OfficeCli.Core.OleHelper.AddEmbeddedPart(hostPart, srcPath, _filePath);

        // 2. Resolve ProgID (explicit > auto-detected from extension).
        var progId = OfficeCli.Core.OleHelper.ResolveProgId(properties, srcPath);

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
        double cxPt = cxEmu / 12700.0;
        double cyPt = cyEmu / 12700.0;
        // Twips for w:dxaOrig/w:dyaOrig (20 twips/point).
        long cxTwips = (long)(cxPt * 20);
        long cyTwips = (long)(cyPt * 20);

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

        var oleXml = $"""
<w:object xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" w:dxaOrig="{cxTwips}" w:dyaOrig="{cyTwips}">
{shapetypeXml}<v:shape id="{shapeId}" type="#_x0000_t75" style="width:{cxPt.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture)}pt;height:{cyPt.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture)}pt" o:ole=""{shapeAltAttr}>
<v:imagedata r:id="{iconRelId}" o:title=""/>
</v:shape>
<o:OLEObject Type="Embed" ProgID="{System.Security.SecurityElement.Escape(progId)}" ShapeID="{shapeId}" DrawAspect="{drawAspect}" ObjectID="{objectId}" r:id="{embedRelId}"/>
</w:object>
""";
        var oleObject = new EmbeddedObject(oleXml);

        // 8. Wrap in a Run and insert it, mirroring the AddPicture positional logic.
        var oleRun = new Run(oleObject);

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
                var runIdxInline = runs.IndexOf(oleRun) + 1;
                return $"{parentPath}/r[{runIdxInline}]";
            }
            return parentPath + "/r[1]";
        }

        string resultPath;
        if (parent is Paragraph existingPara)
        {
            var runCount = existingPara.Elements<Run>().Count();
            if (index.HasValue && index.Value < runCount)
            {
                var refRun = existingPara.Elements<Run>().ElementAt(index.Value);
                existingPara.InsertBefore(oleRun, refRun);
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
            var oleRunIdx = existingPara.Elements<Run>().ToList().IndexOf(oleRun) + 1;
            resultPath = $"{parentPath}/r[{oleRunIdx}]";
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
            var olePIdx = oleCell.Elements<Paragraph>().ToList().IndexOf(olePara) + 1;
            // CONSISTENCY(ole-run-path): same /r[1] suffix as the else branch
            // below — the OLE run is the addressable target, not the paragraph.
            var oleCellRunIdx = olePara.Elements<Run>().ToList().IndexOf(oleRun) + 1;
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
            var olePIdx = parent.Elements<Paragraph>().ToList().IndexOf(olePara) + 1;
            // Return the /r[1] address so callers can Set/Get/Remove the
            // OLE run directly. Picture's Add returns a paragraph-level
            // path because the paragraph Set is meaningful (font, style);
            // for OLE, the only interesting target is the run itself.
            resultPath = $"{parentPath}/{BuildParaPathSegment(olePara, olePIdx)}/r[1]";
        }
        return resultPath;
    }
}
