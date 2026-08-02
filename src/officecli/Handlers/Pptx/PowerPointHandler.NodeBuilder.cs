// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler
{
    // Map a paragraph <a:pPr algn=…> to the canonical friendly token. The four
    // core values get friendly names (left/center/right/justify); the OOXML-only
    // values (justLow / dist / thaiDist) pass through as their raw token so they
    // round-trip through ParseTextAlignment instead of folding to "left".
    private static string MapTextAlignToFriendly(EnumValue<Drawing.TextAlignmentTypeValues>? algn)
    {
        var inner = algn?.InnerText;
        return inner switch
        {
            "l" => "left",
            "ctr" => "center",
            "r" => "right",
            "just" => "justify",
            null or "" => "left",
            _ => inner,   // justLow / dist / thaiDist — preserve verbatim
        };
    }

    // Emit the simple-valued CT_TextParagraphProperties attributes that were
    // previously dropped on readback: line-break / punctuation / font-alignment
    // / default-tab-size. These are heavily used by CJK source decks
    // (eaLnBrk="0" suppresses East-Asian line breaking) — dropping them rewrapped
    // text and shifted layout on round-trip. Canonical keys mirror the OOXML
    // attribute names so Add/Set re-apply them symmetrically.
    private static void EmitParagraphBreakProps(Drawing.ParagraphProperties? pProps, DocumentNode node)
    {
        if (pProps == null) return;
        if (pProps.EastAsianLineBreak?.HasValue == true)
            node.Format["eaLnBrk"] = pProps.EastAsianLineBreak.Value ? "1" : "0";
        if (pProps.LatinLineBreak?.HasValue == true)
            node.Format["latinLnBrk"] = pProps.LatinLineBreak.Value ? "1" : "0";
        if (pProps.FontAlignment?.HasValue == true && !string.IsNullOrEmpty(pProps.FontAlignment.InnerText))
            node.Format["fontAlgn"] = pProps.FontAlignment.InnerText!;
        if (pProps.DefaultTabSize?.HasValue == true)
            node.Format["defTabSz"] = FormatEmu(pProps.DefaultTabSize.Value);
    }

    // CONSISTENCY(effect-color-8digit): shadow/glow readback contract is
    // CSS-form 8-digit hex '#RRGGBBAA' (schema/help/pptx/shape.json
    // shadow.readback / glow.readback). FormatHexWithAlpha falls back to
    // 6-digit when the underlying srgbClr has no a:alpha child, which broke
    // the round-trip promise for the opaque case. Coerce hex colors emitted
    // into the composite shadow/glow strings to 8-digit; scheme color names
    // (accent1, dark1, …) pass through unchanged.
    private static string EnsureEightDigitHexForEffect(string color)
    {
        if (string.IsNullOrEmpty(color)) return color;
        // Color may carry transforms ("000000+lumMod50"). Coerce only the
        // base hex token (before the first '+'); scheme color names
        // (accent1, dark1, …) pass through unchanged.
        var plusIdx = color.IndexOf('+');
        var head = plusIdx >= 0 ? color[..plusIdx] : color;
        var tail = plusIdx >= 0 ? color[plusIdx..] : "";
        var hadHash = head.StartsWith('#');
        var hex = hadHash ? head[1..] : head;
        // All callers are shadow / innerShadow / glow composite strings of the
        // form "#RRGGBB-blur-angle-dist-opacity": the alpha is carried by the
        // dedicated trailing OPACITY field. So the color token must stay 6-digit
        // (no alpha byte). Emitting an 8-digit "#RRGGBBAA" here double-encodes
        // the same a:alpha element — opaque "#RRGGBBFF" alongside a non-100
        // opacity (R1-B4 / R4-6), and re-applies alpha twice on replay. Strip
        // any alpha byte FormatHexWithAlpha baked in, leaving opacity as the
        // single source of truth.
        if (hex.Length == 6 && hex.All(Uri.IsHexDigit))
            return $"#{hex.ToUpperInvariant()}{tail}";
        if (hex.Length == 8 && hex.All(Uri.IsHexDigit))
            return $"#{hex[..6].ToUpperInvariant()}{tail}";
        return color;
    }

    // R58 bt-2: the effectLst walker above consumes outerShdw / innerShdw /
    // glow / fillOverlay / reflection / softEdge / blur. Any other child
    // (tint, lum, hsl, alphaModFix, clrChange, duotone, biLevel, xfrm, …)
    // has no compressible string surface; falling back to verbatim-XML
    // passthrough via `effectsRaw` is the only way to round-trip it.
    // Keep this allowlist in sync with the walker; new compressible readers
    // added above must add their LocalName here so they do not trigger the
    // raw fallback on their own.
    private static readonly HashSet<string> HandledEffectListChildren =
        new(StringComparer.Ordinal)
        {
            "outerShdw", "innerShdw", "glow", "fillOverlay",
            "reflection", "softEdge", "blur",
        };

    private static bool HasUnhandledEffectListChild(Drawing.EffectList effectList)
    {
        foreach (var child in effectList.Elements())
        {
            if (!HandledEffectListChildren.Contains(child.LocalName))
                return true;
        }
        return false;
    }

    private List<DocumentNode> GetSlideChildNodes(SlidePart slidePart, int slideNum, int depth)
    {
        var children = new List<DocumentNode>();
        var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree;
        if (shapeTree == null) return children;
        BuildChildNodesIntoContainer(children, shapeTree, slidePart, slideNum, depth, $"/slide[{slideNum}]", isSlideRoot: true);
        return children;
    }

    // CONSISTENCY(pptx-group-flatten): Get/dump now descends into GroupShape
    // so group-internal picture/table/chart/connector are visible in the
    // returned tree. Each leaf carries its honest path via parentPathPrefix
    // so callers can pipe a Get-emitted path back to Set/Remove. Zoom and
    // 3DModel only enumerate at slide root — they aren't legal group content.
    private void BuildChildNodesIntoContainer(
        List<DocumentNode> children,
        OpenXmlCompositeElement container,
        SlidePart slidePart,
        int slideNum,
        int depth,
        string parentPathPrefix,
        bool isSlideRoot)
    {
        // CONSISTENCY(spTree-order): walk container.ChildElements ONCE in
        // declared order so the Children list mirrors true spTree stacking.
        // Per-type positional indices (shapeIdx, picIdx, tblIdx, chartIdx,
        // grpIdx, cxnIdx) are still per-element-type (matching ResolveShape /
        // ResolvePicture / etc. path semantics), but emission order is the
        // raw spTree order. Previously this routine bucketed by element type
        // (all Shapes first, then GraphicFrames, then Pictures, then Groups,
        // then ConnectionShapes), so dump/replay reordered a kitchen-sink slide
        // built as shape→textbox→picture→table→chart→equation into shape→
        // textbox→equation→table→chart→picture — zorder values on table/chart/
        // equation shifted on every round-trip even though geometry preserved
        // visual stacking. Bug A from round-trip strong-conclusion audit.
        var contentElements = container.ChildElements
            .Where(e => e is Shape or Picture or GraphicFrame or GroupShape or ConnectionShape).ToList();

        int shapeIdx = 0, picIdx = 0, tblIdx = 0, chartIdx = 0, grpIdx = 0, cxnIdx = 0;
        // First pass: just allocate per-type positional indices in element order
        // for nodes other than groups (groups need recursion handled inline
        // below, but their positional index is also assigned during this walk).
        foreach (var el in container.ChildElements)
        {
            switch (el)
            {
                case Shape shape:
                    shapeIdx++;
                    children.Add(ShapeToNode(shape, slideNum, shapeIdx, depth, slidePart, parentPathPrefix));
                    break;
                case GraphicFrame gf:
                    // bt-5: detect by graphicData URI FIRST. A tblPr/extLst
                    // carrying an unknown <a:ext uri> (a16:tblExt, p15:
                    // trackedChange, vendor extensions) can defeat the
                    // strongly-typed Descendants<Drawing.Table>() probe when
                    // the OpenXml SDK reads the unknown ext payload as an
                    // OpenXmlUnknownElement child of tblPr and the tblPr
                    // itself surfaces as an unknown — the typed Table walk
                    // then yields zero results and the entire graphicFrame
                    // falls through every branch, dropping silently from
                    // dump. URI-based detection is invariant against tblPr
                    // extLst content so the table survives even when the
                    // typed accessor stalls.
                    var gfGraphicData = gf.Graphic?.GraphicData;
                    var gfUri = gfGraphicData?.Uri?.Value;
                    bool gfIsTable = gf.Descendants<Drawing.Table>().Any()
                        || (gfUri?.Equals(TableGraphicDataUri, StringComparison.OrdinalIgnoreCase) == true);
                    if (gfIsTable)
                    {
                        tblIdx++;
                        children.Add(TableToNode(gf, slideNum, tblIdx, depth, parentPathPrefix));
                    }
                    else if (gf.Descendants<C.ChartReference>().Any() || IsExtendedChartFrame(gf))
                    {
                        chartIdx++;
                        children.Add(ChartToNode(gf, slidePart, slideNum, chartIdx, depth, parentPathPrefix));
                    }
                    else if (TryExtractPictureFromGraphicFrame(gf) is Picture wrappedPic)
                    {
                        // R48: graphicFrame whose <a:graphicData uri="...picture">
                        // hosts a <p:pic>. Keynote->PPT and Aspose authoring tools
                        // emit pictures in this legal but rare wrapped form
                        // (top-level <p:pic> is the common idiom). Without this
                        // branch the picture is silently dropped on dump.
                        picIdx++;
                        children.Add(PictureToNode(wrappedPic, slideNum, picIdx, slidePart, parentPathPrefix));
                    }
                    else if (IsSmartArtGraphicFrame(gf))
                    {
                        // bt-4: graphicFrame hosting <dgm:relIds> (SmartArt).
                        // Previously fell through all the typed branches and was
                        // silently absent from Get/Query — `query "shape"` /
                        // `get /slide[N]` showed 0 children, the dump emitter's
                        // smartart-aware path runs in parallel from the raw XML
                        // and can recover, but anything iterating the
                        // DocumentNode tree (HTML preview, view text, fuzzers)
                        // missed the diagram entirely. Surface as a typed
                        // "smartart" node so it shows up in the tree; the dump
                        // emitter's EmitSmartArtsForSlide still owns the
                        // diagram-part round-trip.
                        shapeIdx++;
                        children.Add(SmartArtToNode(gf, slideNum, shapeIdx, parentPathPrefix));
                    }
                    break;
                case Picture pic:
                    picIdx++;
                    children.Add(PictureToNode(pic, slideNum, picIdx, slidePart, parentPathPrefix));
                    break;
                case GroupShape grp:
                    grpIdx++;
                    children.Add(BuildGroupNode(grp, slidePart, slideNum, depth, parentPathPrefix, grpIdx, contentElements));
                    break;
                case ConnectionShape cxn:
                    cxnIdx++;
                    children.Add(ConnectorToNode(cxn, slideNum, cxnIdx, parentPathPrefix, depth, slidePart));
                    break;
            }
        }

        // Zoom and 3D model are slide-level only; they are not valid
        // children of a GroupShape per the OOXML schema, so only enumerate
        // them when we're at the slide root. These are appended AFTER the
        // shape-tree walk because they live in <p:ext> sections, not in
        // spTree native order.
        if (isSlideRoot && container is ShapeTree rootShapeTree)
        {
            var zoomElements = GetZoomElements(rootShapeTree);
            int zmIdx = 0;
            foreach (var zmEl in zoomElements)
            {
                zmIdx++;
                children.Add(ZoomToNode(zmEl, slideNum, zmIdx));
            }

            var model3dElements = GetModel3DElements(rootShapeTree);
            int m3dIdx = 0;
            foreach (var m3dEl in model3dElements)
            {
                m3dIdx++;
                children.Add(Model3DToNode(m3dEl, slideNum, m3dIdx));
            }
        }
    }

    private DocumentNode BuildGroupNode(GroupShape grp, SlidePart slidePart, int slideNum,
        int depth, string parentPathPrefix, int grpIdx, List<OpenXmlElement> contentElements)
    {
        var grpName = grp.NonVisualGroupShapeProperties?.NonVisualDrawingProperties?.Name?.Value ?? "Group";
        var grpPathSeg = BuildElementPathSegment("group", grp, grpIdx);
        var grpNode = new DocumentNode
        {
            Path = $"{parentPathPrefix}/{grpPathSeg}",
            Type = "group",
            Preview = grpName,
            ChildCount = grp.Elements<Shape>().Count() + grp.Elements<Picture>().Count()
                + grp.Elements<GraphicFrame>().Count() + grp.Elements<ConnectionShape>().Count()
                + grp.Elements<GroupShape>().Count()
        };
        grpNode.Format["name"] = grpName;
        var grpId = GetCNvPrId(grp);
        if (grpId.HasValue) grpNode.Format["id"] = grpId.Value;
        var grpCreationId = ReadCNvPrCreationId(grp);
        if (grpCreationId != null) grpNode.Format["extLst.creationId"] = grpCreationId;
        var grpXfrm = grp.GroupShapeProperties?.TransformGroup;
        if (grpXfrm?.Offset?.X != null) grpNode.Format["x"] = FormatEmu(grpXfrm.Offset.X.Value);
        if (grpXfrm?.Offset?.Y != null) grpNode.Format["y"] = FormatEmu(grpXfrm.Offset.Y.Value);
        if (grpXfrm?.Extents?.Cx != null) grpNode.Format["width"] = FormatEmu(grpXfrm.Extents.Cx.Value);
        if (grpXfrm?.Extents?.Cy != null) grpNode.Format["height"] = FormatEmu(grpXfrm.Extents.Cy.Value);
        if (grpXfrm?.Rotation != null && grpXfrm.Rotation.Value != 0)
            grpNode.Format["rotation"] = $"{grpXfrm.Rotation.Value / 60000.0:0.######}";
        // R53 bt-4: <a:chOff>/<a:chExt> define the group's child coordinate
        // system. When they diverge from <a:off>/<a:ext> (e.g. a deliberately
        // shifted internal coord system the inner shapes' x/y are computed
        // against), AddGroup defaults them to the outer rect and the inner
        // shapes silently move. Surface verbatim only when they differ from
        // the outer rect so the dump→replay round-trip stays quiet for the
        // common identity case.
        var grpChOff = grpXfrm?.ChildOffset;
        var grpChExt = grpXfrm?.ChildExtents;
        bool grpChOffDiverges = grpChOff != null
            && ((grpChOff.X?.Value ?? 0) != (grpXfrm!.Offset?.X?.Value ?? 0)
                || (grpChOff.Y?.Value ?? 0) != (grpXfrm.Offset?.Y?.Value ?? 0));
        bool grpChExtDiverges = grpChExt != null
            && ((grpChExt.Cx?.Value ?? 0) != (grpXfrm!.Extents?.Cx?.Value ?? 0)
                || (grpChExt.Cy?.Value ?? 0) != (grpXfrm.Extents?.Cy?.Value ?? 0));
        if (grpChOffDiverges)
            grpNode.Format["childOffset"] = $"{grpChOff!.X?.Value ?? 0},{grpChOff.Y?.Value ?? 0}";
        if (grpChExtDiverges)
            grpNode.Format["childExtent"] = $"{grpChExt!.Cx?.Value ?? 0},{grpChExt.Cy?.Value ?? 0}";
        // Note: p:grpSpPr has no fill child per OOXML CT_GroupShapeProperties
        // (only xfrm/scene3d/extLst), so even if a legacy file carries a
        // stray fill on grpSpPr, PowerPoint silently ignores it. Don't
        // surface a fill key on group Get — it would mislead callers into
        // thinking a fill on the group means something visible. Apply fill
        // to the child shapes instead.
        var grpZIdx = contentElements.IndexOf(grp);
        if (grpZIdx >= 0) grpNode.Format["zorder"] = grpZIdx + 1;
        // Hyperlink (nvGrpSpPr/cNvPr/a:hlinkClick) — same slot as shape/picture.
        var grpHl = grp.NonVisualGroupShapeProperties?.NonVisualDrawingProperties?
            .GetFirstChild<Drawing.HyperlinkOnClick>();
        var grpLinkUrl = ReadHyperlinkOnClickUrl(grpHl, slidePart);
        if (grpLinkUrl != null) grpNode.Format["link"] = grpLinkUrl;
        var grpTip = grpHl?.Tooltip?.Value;
        if (!string.IsNullOrEmpty(grpTip)) grpNode.Format["tooltip"] = grpTip!;

        // Recurse into the group's contents when depth allows, so callers
        // see the same iceberg-free view through Get that Query already
        // provides. Group content paths become /slide[N]/group[K]/<type>[L].
        if (depth > 0)
        {
            BuildChildNodesIntoContainer(
                grpNode.Children, grp, slidePart, slideNum, depth - 1,
                $"{parentPathPrefix}/{grpPathSeg}", isSlideRoot: false);
        }
        return grpNode;
    }

    // bt-4: SmartArt sits inside a <p:graphicFrame> whose <a:graphicData uri>
    // is "http://schemas.openxmlformats.org/drawingml/2006/diagram" and whose
    // first child is a <dgm:relIds> pointing to four diagram parts. Detect
    // by URI (the typed accessor finds dgm:relIds anywhere in the subtree,
    // matching how GetSmartArtsOnSlide probes).
    private const string SmartArtGraphicDataUri =
        "http://schemas.openxmlformats.org/drawingml/2006/diagram";

    // bt-5: table graphicFrame URI — used as a defensive fallback when the
    // strongly-typed Descendants<Drawing.Table>() probe fails because tblPr's
    // extLst contains unrecognized extension content (a16:tblExt, p15:
    // trackedChange) that confuses the SDK's tbl-subtree typing.
    private const string TableGraphicDataUri =
        "http://schemas.openxmlformats.org/drawingml/2006/table";

    internal static bool IsSmartArtGraphicFrame(GraphicFrame gf)
    {
        var data = gf.Graphic?.GraphicData;
        if (data?.Uri?.Value?.Equals(SmartArtGraphicDataUri, StringComparison.OrdinalIgnoreCase) == true)
            return true;
        // Fallback: scan descendants by local name + namespace, mirroring
        // GetSmartArtsOnSlide's recovery walk so an authored variant whose
        // graphicData URI was rewritten still matches.
        return gf.Descendants().Any(e =>
            e.LocalName == "relIds" && e.NamespaceUri == SmartArtGraphicDataUri);
    }

    private static DocumentNode SmartArtToNode(GraphicFrame gf, int slideNum, int shapeIdx, string? parentPathPrefix = null)
    {
        var name = gf.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Name?.Value ?? "SmartArt";
        var pathSeg = BuildElementPathSegment("shape", gf, shapeIdx);
        var basePath = parentPathPrefix ?? $"/slide[{slideNum}]";
        var node = new DocumentNode
        {
            Path = $"{basePath}/{pathSeg}",
            Type = "smartart",
            Preview = name,
        };
        node.Format["name"] = name;
        var id = GetCNvPrId(gf);
        if (id.HasValue) node.Format["id"] = id.Value;
        var creationId = ReadCNvPrCreationId(gf);
        if (creationId != null) node.Format["extLst.creationId"] = creationId;
        var offset = gf.Transform?.Offset;
        if (offset?.X != null) node.Format["x"] = Core.EmuConverter.FormatEmuLossy(offset.X.Value);
        if (offset?.Y != null) node.Format["y"] = Core.EmuConverter.FormatEmuLossy(offset.Y.Value);
        var extents = gf.Transform?.Extents;
        if (extents?.Cx != null) node.Format["width"] = Core.EmuConverter.FormatEmuLossy(extents.Cx.Value);
        if (extents?.Cy != null) node.Format["height"] = Core.EmuConverter.FormatEmuLossy(extents.Cy.Value);
        if (gf.Parent is ShapeTree zTree)
        {
            var contentEls = zTree.ChildElements
                .Where(e => e is Shape or Picture or GraphicFrame or GroupShape or ConnectionShape)
                .ToList();
            var zIdx = contentEls.IndexOf(gf);
            if (zIdx >= 0) node.Format["zorder"] = zIdx + 1;
        }
        return node;
    }

    private static DocumentNode TableToNode(GraphicFrame gf, int slideNum, int tblIdx, int depth, string? parentPathPrefix = null)
    {
        // bt-5: the typed Drawing.Table accessor can return zero results when
        // tblPr's extLst carries unrecognized extension content that the
        // OpenXml SDK fails to type-bind. The graphicFrame walker above falls
        // back to detecting by graphicData URI; mirror that fallback here so
        // we don't throw on .First() when typed Table is absent but the
        // graphicData URI confirms a table host.
        var table = gf.Descendants<Drawing.Table>().FirstOrDefault();
        var rows = table?.Elements<Drawing.TableRow>().ToList() ?? new List<Drawing.TableRow>();
        var cols = rows.FirstOrDefault()?.Elements<Drawing.TableCell>().Count() ?? 0;
        var name = gf.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Name?.Value ?? "Table";

        var tblPathSeg = BuildElementPathSegment("table", gf, tblIdx);
        var basePath = parentPathPrefix ?? $"/slide[{slideNum}]";
        var tblPath = $"{basePath}/{tblPathSeg}";
        var node = new DocumentNode
        {
            Path = tblPath,
            Type = "table",
            Preview = $"{name} ({rows.Count}x{cols})",
            ChildCount = rows.Count
        };

        node.Format["name"] = name;
        var tblId = GetCNvPrId(gf);
        if (tblId.HasValue) node.Format["id"] = tblId.Value;
        var tblCreationId = ReadCNvPrCreationId(gf);
        if (tblCreationId != null) node.Format["extLst.creationId"] = tblCreationId;
        node.Format["rows"] = rows.Count;
        node.Format["cols"] = cols;

        var gridCols = table?.TableGrid?.Elements<Drawing.GridColumn>().ToList();
        if (gridCols != null && gridCols.Count > 0)
            node.Format["colWidths"] = string.Join(",", gridCols.Select(gc => gc.Width?.Value is long w ? Core.EmuConverter.FormatEmuLossy(w) : "0"));

        // Table style
        var tblPr = table?.GetFirstChild<Drawing.TableProperties>();
        var tableStyleId = tblPr?.GetFirstChild<Drawing.TableStyleId>()?.InnerText;
        if (!string.IsNullOrEmpty(tableStyleId))
        {
            var styleName = OfficeCli.Core.TableStyles.TableStyleRegistry.GuidToShortName(tableStyleId);
            // CONSISTENCY(canonical-key): emit only canonical 'style'; schema lists
            // 'tableStyle' and 'tableStyleId' as input aliases (Set side) — Get
            // normalizes to canonical (style = resolved name when known, else GUID).
            node.Format["style"] = styleName ?? tableStyleId;
        }

        // TableLook flags
        if (tblPr != null)
        {
            // firstRow and bandRow default to TRUE in AddTable (an interactive
            // nicety — a bare `add table` yields a styled header + banding). But
            // the OOXML default for an ABSENT firstRow/bandRow attribute is
            // FALSE. Emit their EFFECTIVE value (absent → false) so a source
            // that omits them round-trips faithfully instead of gaining a header
            // row / banding on replay (bnc480256: <a:tblPr bandRow="1"> with no
            // firstRow replayed as firstRow="1", adding a header band + gap).
            // The other four default to false in AddTable already, so emitting
            // them only when present stays faithful.
            node.Format["firstRow"] = tblPr.FirstRow?.Value ?? false;
            if (tblPr.LastRow is not null) node.Format["lastRow"] = tblPr.LastRow.Value;
            if (tblPr.FirstColumn is not null) node.Format["firstCol"] = tblPr.FirstColumn.Value;
            if (tblPr.LastColumn is not null) node.Format["lastCol"] = tblPr.LastColumn.Value;
            node.Format["bandedRows"] = tblPr.BandRow?.Value ?? false;
            if (tblPr.BandColumn is not null) node.Format["bandedCols"] = tblPr.BandColumn.Value;
        }

        // Outer-edge border aggregation (PPT has no table-level border element).
        // Scan the outer edges across cells; emit per-side keys when uniform,
        // and 'border.all' shorthand when all four sides match.
        if (table != null) AggregateTableOuterBorders(table, rows, node);

        // Position
        var offset = gf.Transform?.Offset;
        if (offset != null)
        {
            if (offset.X is not null) node.Format["x"] = Core.EmuConverter.FormatEmuLossy(offset.X!);
            if (offset.Y is not null) node.Format["y"] = Core.EmuConverter.FormatEmuLossy(offset.Y!);
        }
        var extents = gf.Transform?.Extents;
        if (extents != null)
        {
            if (extents.Cx is not null) node.Format["width"] = Core.EmuConverter.FormatEmuLossy(extents.Cx!);
            if (extents.Cy is not null) node.Format["height"] = Core.EmuConverter.FormatEmuLossy(extents.Cy!);
        }

        // CONSISTENCY(zorder): mirror shape/picture/connector — emit when
        // parented to a ShapeTree so dump/replay preserves stacking order.
        if (gf.Parent is ShapeTree tblZTree)
        {
            var tblZContent = tblZTree.ChildElements
                .Where(e => e is Shape or Picture or GraphicFrame or GroupShape or ConnectionShape)
                .ToList();
            var tblZIdx = tblZContent.IndexOf(gf);
            if (tblZIdx >= 0) node.Format["zorder"] = tblZIdx + 1;
        }

        if (depth > 0)
        {
            int rIdx = 0;
            foreach (var row in rows)
            {
                rIdx++;
                var rowNode = new DocumentNode
                {
                    Path = $"{tblPath}/tr[{rIdx}]",
                    Type = "tr",
                    ChildCount = row.Elements<Drawing.TableCell>().Count()
                };

                // Row height
                if (row.Height?.HasValue == true)
                    rowNode.Format["height"] = Core.EmuConverter.FormatEmuLossy(row.Height.Value);

                if (depth > 1)
                {
                    int cIdx = 0;
                    foreach (var cell in row.Elements<Drawing.TableCell>())
                    {
                        cIdx++;
                        var cellText = GetCellTextWithParagraphBreaks(cell);
                        var cellNode = new DocumentNode
                        {
                            Path = $"{tblPath}/tr[{rIdx}]/tc[{cIdx}]",
                            Type = "tc",
                            Text = cellText
                        };

                        // CONSISTENCY(empty-run-preserve): mirror R43 b209ac10
                        // for table cells. PowerPoint persists an empty cell
                        // as <a:p><a:r><a:rPr lang="en-US"/><a:t/></a:r></a:p>
                        // — the run-bearing form carries IME / cursor state.
                        // AddTable's blank-cell seed uses <a:endParaRPr/>
                        // instead, so dump→replay of a run-bearing empty cell
                        // silently flips to the endParaRPr form. Surface a
                        // Format marker the emitter can read to decide
                        // whether to issue `set tc[K] text=""` (which forces
                        // the run-bearing form via AppendLineWithTabs) even
                        // though the cell text itself is empty.
                        bool hasAnyRun = cell.TextBody?.Descendants<Drawing.Run>().Any() == true;
                        if (string.IsNullOrEmpty(cellText) && hasAnyRun)
                            cellNode.Format["hasEmptyRun"] = true;

                        // Verbatim cell text-body passthrough. The batch emitter
                        // rebuilds a cell via `set tc[K] text=...`, which routes
                        // through AppendLineWithTabs and produces BARE paragraphs:
                        // every per-paragraph <a:pPr> child (lnSpc/spcBef/spcAft/
                        // buClrTx/buFontTx/buSzTx/buNone/tabLst/defRPr), the
                        // <a:lstStyle>, and every run's rich <a:rPr> (ea/latin/
                        // solidFill) are dropped — the cell text reflows off the
                        // source's intended metrics. Capture the whole <a:txBody>
                        // OuterXml so the cell Set path can re-inject it verbatim,
                        // superseding the text= rebuild. Only when the cell carries
                        // real paragraph content (skip the trivial empty-cell seed,
                        // which round-trips fine through the existing text= path).
                        // CONSISTENCY(cell-txbody-raw-passthrough): mirrors the shape
                        // lstStyleRaw / effectsRaw verbatim passthrough.
                        if (cell.TextBody != null
                            && CellTextBodyHasRichContent(cell.TextBody))
                            cellNode.Format["txBodyRaw"] = cell.TextBody.OuterXml;

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
                            // Preserve all stops (including intermediate ones) via the shared helper.
                            cellNode.Format["gradient"] = ReadGradientString(gradFill);
                            cellNode.Format["fill"] = "gradient";
                        }
                        else
                        {
                            // BUG-R6-A: Read both RgbColorModelHex and SchemeColor for cell fill
                            // (mirror shape fill behavior). Scheme colors (accent1, dark1, ...)
                            // were silently dropped before.
                            var cellFillSolid = tcPr?.GetFirstChild<Drawing.SolidFill>();
                            var cellFillColor = ReadColorFromFill(cellFillSolid);
                            if (cellFillColor != null) cellNode.Format["fill"] = cellFillColor;
                            // Explicit <a:noFill/> beats the table style's band/
                            // firstRow fill; dropping it painted the cell with the
                            // style fill on round-trip (sample18).
                            else if (tcPr?.GetFirstChild<Drawing.NoFill>() != null)
                                cellNode.Format["fill"] = "none";
                        }

                        // Cell borders (including diagonal tl2br/tr2bl)
                        if (tcPr != null) ReadTableCellBorders(tcPr, cellNode);

                        // BUG-R6-A: cell padding readback (Set wrote LeftMargin/etc; Get
                        // missed it on the NodeBuilder cell branch). Canonical key is
                        // "padding.*" per cross-handler rule (the project conventions).
                        if (tcPr?.LeftMargin?.HasValue == true)
                            cellNode.Format["padding.left"] = FormatEmu(tcPr.LeftMargin.Value);
                        if (tcPr?.RightMargin?.HasValue == true)
                            cellNode.Format["padding.right"] = FormatEmu(tcPr.RightMargin.Value);
                        if (tcPr?.TopMargin?.HasValue == true)
                            cellNode.Format["padding.top"] = FormatEmu(tcPr.TopMargin.Value);
                        if (tcPr?.BottomMargin?.HasValue == true)
                            cellNode.Format["padding.bottom"] = FormatEmu(tcPr.BottomMargin.Value);

                        // BUG-R6-A: emit colspan/rowspan on cell node (mirror Query.cs).
                        if (cell.GridSpan?.HasValue == true && cell.GridSpan.Value > 1)
                            cellNode.Format["colspan"] = cell.GridSpan.Value;
                        if (cell.RowSpan?.HasValue == true && cell.RowSpan.Value > 1)
                            cellNode.Format["rowspan"] = cell.RowSpan.Value;
                        if (cell.HorizontalMerge?.HasValue == true && cell.HorizontalMerge.Value)
                            cellNode.Format["hmerge"] = true;
                        if (cell.VerticalMerge?.HasValue == true && cell.VerticalMerge.Value)
                            cellNode.Format["vmerge"] = true;

                        // Cell text direction (a:tcPr @vert). Canonical readback
                        // mirrors the Set vocabulary (horizontal / vertical270 /
                        // vertical90 / stacked) so round-trip equality holds.
                        if (tcPr?.Vertical?.HasValue == true)
                        {
                            cellNode.Format["textdirection"] = tcPr.Vertical.InnerText switch
                            {
                                "horz" => "horizontal",
                                "vert" => "vertical90",
                                "vert270" => "vertical270",
                                "wordArtVert" => "stacked",
                                "eaVert" => "eaVert",
                                "mongolianVert" => "mongolianVert",
                                "wordArtVertRtl" => "wordArtVertRtl",
                                _ => tcPr.Vertical.InnerText
                            };
                        }

                        // Cell text wrap (a:tcPr/a:txBody/a:bodyPr @wrap).
                        // Set writes square|none on the cell's BodyProperties;
                        // mirror back as bool (false == "none", true == "square").
                        var cellBodyPr = cell.TextBody?.GetFirstChild<Drawing.BodyProperties>();
                        if (cellBodyPr?.Wrap?.HasValue == true)
                        {
                            cellNode.Format["wrap"] = cellBodyPr.Wrap.Value != Drawing.TextWrappingValues.None;
                        }

                        // Cell vertical alignment
                        if (tcPr?.Anchor?.HasValue == true)
                        {
                            var av = tcPr.Anchor.Value;
                            if (av == Drawing.TextAnchoringTypeValues.Top) cellNode.Format["valign"] = "top";
                            else if (av == Drawing.TextAnchoringTypeValues.Center) cellNode.Format["valign"] = "center";
                            else if (av == Drawing.TextAnchoringTypeValues.Bottom) cellNode.Format["valign"] = "bottom";
                            else cellNode.Format["valign"] = tcPr.Anchor.InnerText switch
                            {
                                "ctr" => "center",
                                _ => tcPr.Anchor.InnerText
                            };
                        }

                        // R56 bt-4: a:tcPr @horzOverflow (overflow|clip) — typed
                        // SDK property. Mirror Set vocabulary verbatim.
                        if (tcPr?.HorizontalOverflow?.HasValue == true)
                        {
                            var hov = tcPr.HorizontalOverflow.Value;
                            if (hov == Drawing.TextHorizontalOverflowValues.Overflow)
                                cellNode.Format["horzOverflow"] = "overflow";
                            else if (hov == Drawing.TextHorizontalOverflowValues.Clip)
                                cellNode.Format["horzOverflow"] = "clip";
                            else
                                cellNode.Format["horzOverflow"] = tcPr.HorizontalOverflow.InnerText;
                        }

                        // R56 bt-4: a:tcPr @lockText is a non-standard MS Office
                        // extension (not in the SDK enum); preserve via raw
                        // attribute iteration. Same KeyNotFoundException-safe
                        // pattern as bodyPr/rtlCol above.
                        if (tcPr != null)
                        {
                            foreach (var attr in tcPr.GetAttributes())
                            {
                                if (attr.LocalName == "lockText")
                                {
                                    var av = attr.Value ?? "";
                                    cellNode.Format["lockText"] = av == "1"
                                        || av.Equals("true", StringComparison.OrdinalIgnoreCase);
                                    break;
                                }
                            }
                        }

                        // Cell run-level formatting (font, size, bold, italic, underline, strike, color)
                        var cellFirstRun = cell.Descendants<Drawing.Run>().FirstOrDefault();
                        if (cellFirstRun?.RunProperties != null)
                        {
                            var rp = cellFirstRun.RunProperties;
                            var cellLatin = rp.GetFirstChild<Drawing.LatinFont>()?.Typeface?.Value;
                            var cellEa = rp.GetFirstChild<Drawing.EastAsianFont>()?.Typeface?.Value;
                            var cellCs = rp.GetFirstChild<Drawing.ComplexScriptFont>()?.Typeface?.Value;
                            // Bare `font` is the Latin slot alias only — see
                            // CONSISTENCY(font-bare-latin-only).
                            if (cellLatin != null) cellNode.Format["font"] = cellLatin;
                            // CONSISTENCY(canonical-keys): always emit per-script
                            // slots when present (schema declares get:true).
                            if (cellLatin != null) cellNode.Format["font.latin"] = cellLatin;
                            // CONSISTENCY(per-script-round-trip): see shape-level note.
                            if (cellEa != null) cellNode.Format["font.ea"] = cellEa;
                            if (cellCs != null) cellNode.Format["font.cs"] = cellCs;

                            if (rp.FontSize?.HasValue == true)
                                cellNode.Format["size"] = $"{rp.FontSize.Value / 100.0:0.##}pt";

                            if (rp.Bold?.HasValue == true) cellNode.Format["bold"] = rp.Bold.Value;
                            if (rp.Italic?.HasValue == true) cellNode.Format["italic"] = rp.Italic.Value;

                            if (rp.Underline?.HasValue == true && rp.Underline.Value != Drawing.TextUnderlineValues.None)
                            {
                                cellNode.Format["underline"] = rp.Underline.InnerText switch
                                {
                                    "sng" => "single",
                                    "dbl" => "double",
                                    _ => rp.Underline.InnerText
                                };
                            }
                            if (rp.Strike?.HasValue == true)
                            {
                                cellNode.Format["strike"] = rp.Strike.Value switch
                                {
                                    var v when v == Drawing.TextStrikeValues.DoubleStrike => "double",
                                    var v when v == Drawing.TextStrikeValues.NoStrike => "none",
                                    _ => "single",
                                };
                            }

                            var cellRunColor = ReadColorFromFill(rp.GetFirstChild<Drawing.SolidFill>());
                            if (cellRunColor != null) cellNode.Format["color"] = cellRunColor;

                            if (rp.Spacing?.HasValue == true)
                                cellNode.Format["spacing"] = $"{rp.Spacing.Value / 100.0:0.##}";
                            // R56 bt-3: baseline is OOXML thousandths of a percent
                            // (33000 → 33%). Emit unit-qualified `%` so the canonical
                            // readback distinguishes percent semantics from a raw
                            // numeric input on Set; the parser accepts either form.
                            if (rp.Baseline?.HasValue == true && rp.Baseline.Value != 0)
                                cellNode.Format["baseline"] = $"{rp.Baseline.Value / 1000.0:0.##}%";
                        }

                        // Cell paragraph alignment
                        var cellFirstPara = cell.TextBody?.Elements<Drawing.Paragraph>().FirstOrDefault();
                        if (cellFirstPara?.ParagraphProperties?.Alignment?.HasValue == true)
                        {
                            var alv = cellFirstPara.ParagraphProperties.Alignment.Value;
                            var align = cellFirstPara.ParagraphProperties.Alignment.InnerText;
                            if (alv == Drawing.TextAlignmentTypeValues.Left) align = "left";
                            else if (alv == Drawing.TextAlignmentTypeValues.Center) align = "center";
                            else if (alv == Drawing.TextAlignmentTypeValues.Right) align = "right";
                            else if (alv == Drawing.TextAlignmentTypeValues.Justified) align = "justify";
                            else if (align == "ctr") align = "center";
                            cellNode.Format["align"] = align;
                        }

                        // Cell paragraph direction (mirrors shape/textbox readback).
                        // Only emit when explicitly set on the first paragraph; ltr
                        // is the schema default so absence == ltr.
                        if (cellFirstPara?.ParagraphProperties?.RightToLeft?.HasValue == true)
                            cellNode.Format["direction"] = cellFirstPara.ParagraphProperties.RightToLeft.Value ? "rtl" : "ltr";

                        // BUG-R6-A: cell-level lineSpacing/spaceBefore/spaceAfter readback
                        // from first paragraph (mirrors shape paragraph aggregation —
                        // Set writes to all paragraphs; Get returns the first one's value).
                        var cellFirstPProps = cellFirstPara?.ParagraphProperties;
                        if (cellFirstPProps != null)
                        {
                            var cellLsPct = cellFirstPProps.GetFirstChild<Drawing.LineSpacing>()?.GetFirstChild<Drawing.SpacingPercent>().PercentVal();
                            if (cellLsPct.HasValue) cellNode.Format["lineSpacing"] = SpacingConverter.FormatPptLineSpacingPercent(cellLsPct.Value);
                            var cellLsPts = cellFirstPProps.GetFirstChild<Drawing.LineSpacing>()?.GetFirstChild<Drawing.SpacingPoints>()?.Val?.Value;
                            if (cellLsPts.HasValue) cellNode.Format["lineSpacing"] = SpacingConverter.FormatPptLineSpacingPoints(cellLsPts.Value);
                            var cellSb = cellFirstPProps.GetFirstChild<Drawing.SpaceBefore>()?.GetFirstChild<Drawing.SpacingPoints>()?.Val?.Value;
                            if (cellSb.HasValue) cellNode.Format["spaceBefore"] = SpacingConverter.FormatPptSpacing(cellSb.Value);
                            var cellSa = cellFirstPProps.GetFirstChild<Drawing.SpaceAfter>()?.GetFirstChild<Drawing.SpacingPoints>()?.Val?.Value;
                            if (cellSa.HasValue) cellNode.Format["spaceAfter"] = SpacingConverter.FormatPptSpacing(cellSa.Value);
                        }

                        rowNode.Children.Add(cellNode);
                    }
                }
                node.Children.Add(rowNode);
            }
        }

        return node;
    }

    // CONSISTENCY(pptx-group-flatten): single recursive walker that yields every
    // renderable element in shapeTree order, descending into GroupShape so query
    // and view sees a true union of root + group-internal content. Positional
    // counters reset per parent so BuildElementPathSegment produces stable paths
    // like `/slide[1]/group[2]/shape[3]` (or `@id=` form when cNvPr.Id present).
    // Group containers yield themselves before their children — `query "group"`
    // returns all groups at any depth; `query "shape"` returns leaf shapes only
    // because the type filter happens after yield.
    internal readonly record struct RenderableYield(
        OpenXmlElement Element, string ParentPath, string TypeName, int IndexInParent);

    private static IEnumerable<RenderableYield> EnumerateRenderableElements(
        OpenXmlCompositeElement container, string parentPath)
    {
        int shapeIdx = 0, picIdx = 0, tblIdx = 0, chartIdx = 0, cxnIdx = 0, grpIdx = 0;
        foreach (var child in container.ChildElements)
        {
            // mc:AlternateContent is parsed as OpenXmlUnknownElement by SDK so
            // strongly-typed Descendants<T> won't enter — but we walk
            // ChildElements directly here, so skip the wrapper explicitly to
            // avoid double-counting (Choice + Fallback both have <p:sp>).
            // CONSISTENCY(mc-alt-skip): the defense is at the walker level,
            // not per-call-site.
            if (child is OpenXmlUnknownElement u && u.LocalName == "AlternateContent")
                continue;

            switch (child)
            {
                case Shape s:
                    shapeIdx++;
                    yield return new RenderableYield(s, parentPath, "shape", shapeIdx);
                    break;
                case Picture p:
                    picIdx++;
                    yield return new RenderableYield(p, parentPath, "picture", picIdx);
                    break;
                case ConnectionShape cxn:
                    cxnIdx++;
                    yield return new RenderableYield(cxn, parentPath, "connector", cxnIdx);
                    break;
                case GraphicFrame gf:
                    // bt-5: mirror the BuildChildNodesIntoContainer URI-fallback
                    // so an unknown tblPr extLst doesn't dropkick the table
                    // from query/view either.
                    var uri2 = gf.Graphic?.GraphicData?.Uri?.Value;
                    bool isTable2 = gf.Descendants<Drawing.Table>().Any()
                        || (uri2?.Equals(TableGraphicDataUri, StringComparison.OrdinalIgnoreCase) == true);
                    if (isTable2)
                    {
                        tblIdx++;
                        yield return new RenderableYield(gf, parentPath, "table", tblIdx);
                    }
                    else if (gf.Descendants<C.ChartReference>().Any() || IsExtendedChartFrame(gf))
                    {
                        chartIdx++;
                        yield return new RenderableYield(gf, parentPath, "chart", chartIdx);
                    }
                    else if (TryExtractPictureFromGraphicFrame(gf) is Picture wrappedPic)
                    {
                        // R48: graphicFrame-wrapped picture (uri=".../picture") —
                        // mirror the slide-root walker so HTML preview / svg
                        // export reach the underlying <p:pic>.
                        picIdx++;
                        yield return new RenderableYield(wrappedPic, parentPath, "picture", picIdx);
                    }
                    else if (IsSmartArtGraphicFrame(gf))
                    {
                        // bt-4: SmartArt graphicFrame — surface to query/view
                        // walkers as a "smartart" yield so HTML preview, view
                        // text, and other consumers see the diagram instead of
                        // an empty slide. Indexed against shapeIdx so the path
                        // mirrors SmartArtToNode's /slide[N]/shape[K].
                        shapeIdx++;
                        yield return new RenderableYield(gf, parentPath, "smartart", shapeIdx);
                    }
                    break;
                case GroupShape g:
                    grpIdx++;
                    yield return new RenderableYield(g, parentPath, "group", grpIdx);
                    var nestedParent = $"{parentPath}/{BuildElementPathSegment("group", g, grpIdx)}";
                    foreach (var nested in EnumerateRenderableElements(g, nestedParent))
                        yield return nested;
                    break;
            }
        }
    }

    private static DocumentNode ShapeToNode(Shape shape, int slideNum, int shapeIdx, int depth, OpenXmlPart? part = null, string? parentPathPrefix = null)
    {
        var text = GetShapeText(shape);
        var name = GetShapeName(shape);
        var isTitle = IsTitle(shape);
        var isEquation = !isTitle && shape.TextBody != null
            && shape.TextBody.Descendants().Any(e => e.LocalName == "oMath" || e.LocalName == "oMathPara"
                || (e.LocalName == "m" && e.NamespaceUri == "http://schemas.microsoft.com/office/drawing/2010/main"));

        // Read <p:ph> once: schema declares phType/phIndex as get:true with
        // readback. Previously only IsTitle peeked at it for Type discrimination,
        // so phType=body/subTitle/date/footer/slidenum/header/picture/chart/
        // table/diagram/media/obj/clipArt all collapsed to Type="textbox" with
        // no Format["phType"]. Now: non-title placeholders surface as
        // Type="placeholder", and every placeholder (incl. title) emits
        // Format["phType"] in the human-readable form ParsePlaceholderType
        // accepts on input — so it round-trips through Add.
        var phElemForNode = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
            ?.GetFirstChild<PlaceholderShape>();
        var phTypeStr = phElemForNode != null ? FormatPlaceholderType(phElemForNode.Type?.Value) : null;
        var isPlaceholder = phElemForNode != null;

        var shapePathSeg = BuildElementPathSegment("shape", shape, shapeIdx);
        var basePath = parentPathPrefix ?? $"/slide[{slideNum}]";
        var shapePath = $"{basePath}/{shapePathSeg}";
        // `txBox="1"` on <p:cNvSpPr> is the on-disk marker for a dedicated text
        // container (PowerPoint's "Insert Text Box"). Without it, a shape with a
        // prstGeom — even with no authored text — is a geometry shape, not a
        // textbox. Falling back to "textbox" for every non-title/equation/
        // placeholder collapsed both flavors and broke `add --type shape`
        // round-trip (Get reported Type="textbox").
        var isTextBox = shape.NonVisualShapeProperties
            ?.NonVisualShapeDrawingProperties?.TextBox?.Value == true;
        var node = new DocumentNode
        {
            Path = shapePath,
            Type = isTitle ? "title"
                : isEquation ? "equation"
                : isPlaceholder ? "placeholder"
                : isTextBox ? "textbox"
                : "shape",
            Text = text,
            Preview = string.IsNullOrEmpty(text) ? name : (text.Length > 50 ? text[..50] + "..." : text)
        };

        node.Format["name"] = name;
        if (phTypeStr != null) node.Format["phType"] = phTypeStr;
        // Store the placeholder idx as its raw uint, NOT (int). Decks (often
        // authored by third-party editors) use high idx values up to uint.MaxValue
        // (4294967295); an (int) cast overflows that to -1, which AddPlaceholder's
        // uint.TryParse then rejects → the placeholder auto-assigns a fresh idx,
        // binds to the WRONG layout slot, and loses all inherited font/size/color
        // (whole-deck text re-styling). Mirrors the Query.cs placeholder builders,
        // which already store ph.Index.Value (uint) directly.
        if (phElemForNode?.Index?.Value is uint phIdx) node.Format["phIndex"] = phIdx;
        // A truly BARE <p:ph/> (no type attribute AND no idx) must round-trip
        // bare. FormatPlaceholderType(null) reports "body" for human-facing Get,
        // but replaying it as type="body"+idx binds the shape to the layout's
        // body slot and inherits its bullet/formatting — a bare ph inherits none
        // of that (ECMA-376 default type is "obj", unbound). Source decks use a
        // bare ph precisely to opt OUT of master styling (formatting-bullet-
        // indent: "Object without master styling" gained a bullet on replay).
        // Emit a round-trip marker; AddPlaceholder writes <p:ph/> verbatim.
        if (isPlaceholder && phElemForNode?.Type?.Value == null && phElemForNode?.Index?.Value == null)
            node.Format["phBare"] = "true";

        // CONSISTENCY(splocks-round-trip): <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
        // is the on-disk marker that the shape cannot be ungrouped (PowerPoint
        // writes it on every placeholder it generates). Without surfacing it,
        // dump→replay drops the spLocks element and downstream tools see the
        // placeholder as a free-form shape. Read the noGrp attribute; other
        // spLocks/* attrs follow the same shape if/when they're needed.
        var spLocks = shape.NonVisualShapeProperties?.NonVisualShapeDrawingProperties
            ?.GetFirstChild<Drawing.ShapeLocks>();
        if (spLocks?.NoGrouping?.Value == true)
            node.Format["noGrp"] = true;

        // CONSISTENCY(equation-formula-readback): surface the OMath as a LaTeX
        // string on Get so dump emitter can carry it as `formula=...` on
        // `add equation` — AddEquation requires formula/text or it throws.
        // Without this, equation shapes round-trip as `add equation` with no
        // formula prop and replay fails.
        if (isEquation)
        {
            var mathElements = FindShapeMathElements(shape);
            if (mathElements.Count > 0)
            {
                try
                {
                    var latex = FormulaParser.ToLatex(mathElements[0]);
                    if (!string.IsNullOrEmpty(latex))
                        node.Format["formula"] = latex;
                }
                catch
                {
                    // ToLatex may fail on exotic OMath shapes; fall through —
                    // emitter can still degrade gracefully.
                }
            }
        }

        // Cross-handler `evaluated` protocol — surface unevaluated a:fld
        // descendants on the shape node so agents can find them via Get
        // without parsing view-issues messages. Emit `false` if any dynamic
        // a:fld (slidenum / datetime*) inside this shape has an empty <a:t>;
        // omit the key entirely when there are no dynamic fields at all
        // (matches Word's pattern: only fields surface `evaluated`).
        if (shape.TextBody != null)
        {
            bool anyDynamic = false;
            bool anyUnevaluated = false;
            foreach (var fld in shape.TextBody.Descendants<Drawing.Field>())
            {
                var fldType = fld.Type?.Value ?? "";
                // Single source of truth in Helpers.IsDynamicSlideFieldTypeStatic.
                // Adding new dynamic types only needs one edit there.
                if (!IsDynamicSlideFieldTypeStatic(fldType)) continue;
                anyDynamic = true;
                var cached = string.Concat(fld.Elements<Drawing.Text>().Select(t => t.Text));
                if (cached.Length == 0) { anyUnevaluated = true; break; }
            }
            if (anyDynamic) node.Format["evaluated"] = !anyUnevaluated;
        }

        // CONSISTENCY(alt-readback): Set accepts alt/altText/description and
        // writes to NonVisualDrawingProperties.Description. Surface it on Get
        // so writes are observable.
        var shapeAlt = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Description?.Value;
        if (!string.IsNullOrEmpty(shapeAlt)) node.Format["alt"] = shapeAlt;
        var shapeId = GetCNvPrId(shape);
        if (shapeId.HasValue) node.Format["id"] = shapeId.Value;
        // bt-3: surface modern-PowerPoint creationId GUID stored on cNvPr's
        // extLst so dump→replay preserves stable shape identity (used by
        // change-tracking, comment anchors, animation timelines).
        var shapeCreationId = ReadCNvPrCreationId(shape);
        if (shapeCreationId != null) node.Format["extLst.creationId"] = shapeCreationId;
        // CONSISTENCY(istitle-bool): always emit isTitle so query selectors
        // `[isTitle=true]` and `[isTitle=false]` are both honored by the
        // AttributeFilter post-query pass (which checks node.Format directly).
        node.Format["isTitle"] = isTitle;

        // Position and size
        var xfrm = shape.ShapeProperties?.Transform2D;
        if (xfrm != null)
        {
            if (xfrm.Offset != null)
            {
                if (xfrm.Offset.X is not null) node.Format["x"] = FormatEmu(xfrm.Offset.X!);
                if (xfrm.Offset.Y is not null) node.Format["y"] = FormatEmu(xfrm.Offset.Y!);
            }
            if (xfrm.Extents != null)
            {
                if (xfrm.Extents.Cx is not null) node.Format["width"] = FormatEmu(xfrm.Extents.Cx!);
                if (xfrm.Extents.Cy is not null) node.Format["height"] = FormatEmu(xfrm.Extents.Cy!);
            }
        }

        // Shape fill
        var shapeFill = shape.ShapeProperties?.GetFirstChild<Drawing.SolidFill>();
        var shapeFillColor = ReadColorFromFill(shapeFill);
        if (shapeFillColor != null) node.Format["fill"] = shapeFillColor;
        // Gradient fill on shape
        var shapeGradFill = shape.ShapeProperties?.GetFirstChild<Drawing.GradientFill>();
        if (shapeGradFill != null)
        {
            var stops = shapeGradFill.GradientStopList?.Elements<Drawing.GradientStop>().ToList();
            if (stops != null && stops.Count >= 2)
            {
                var gc1 = ParseHelpers.FormatHexColor(stops[0].GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value ?? "");
                var gc2 = ParseHelpers.FormatHexColor(stops[^1].GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value ?? "");
                var lin = shapeGradFill.GetFirstChild<Drawing.LinearGradientFill>();
                var deg = lin?.Angle?.Value != null ? lin.Angle.Value / 60000.0 : 0.0;

                // CONSISTENCY(gradient-alpha-source): the gradient= string emitted by
                // ReadGradientString already encodes per-stop alpha (either as
                // `#RRGGBBAA` or `+alpha=permille` suffix). Emitting Format["opacity"]
                // from the first stop's alpha here triggers AddShape's "apply opacity
                // to every gradient stop" loop on replay, which overwrites the source's
                // distinct per-stop alphas with the uniform first-stop value — e.g.
                // beautiful-deck.pptx slide 1 AccentTL (radial halo: stop0 alpha=15000,
                // stop1 alpha=0) lost its outer-transparent stop and surfaced as a
                // solid 15% opaque disc on replay. Per-stop alpha rides through the
                // gradient= round-trip; do not double-emit it as a shape-level opacity.
            }
        }
        if (shape.ShapeProperties?.GetFirstChild<Drawing.NoFill>() != null) node.Format["fill"] = "none";

        // Opacity (Alpha on SolidFill color element)
        var fillColorEl = shapeFill?.GetFirstChild<Drawing.RgbColorModelHex>() as OpenXmlElement
            ?? shapeFill?.GetFirstChild<Drawing.SchemeColor>();
        var alphaVal = fillColorEl?.GetFirstChild<Drawing.Alpha>()?.Val?.Value;
        if (alphaVal.HasValue) node.Format["opacity"] = $"{alphaVal.Value / 100000.0:0.##}";

        // Shape preset/geometry
        var presetGeom = shape.ShapeProperties?.GetFirstChild<Drawing.PresetGeometry>();
        if (presetGeom?.Preset?.HasValue == true)
        {
            node.Format["geometry"] = presetGeom.Preset.InnerText;
            // CONSISTENCY(preset-adj-handles): <a:avLst><a:gd name="adj"
            // fmla="val N"/>...</a:avLst> carries per-preset adjust handle
            // values (rounded-rect corner radius, callout pointer tail,
            // arrow head proportions, etc.). Silently dropping them on
            // dump→replay reverts every preset to its default proportions
            // — a visible geometry shift. Surface as `adj=adj1:N,adj2:M`
            // (canonical key for the bag, comma-joined per handle); Set /
            // Add accept the same form on the way back.
            var avLst = presetGeom.GetFirstChild<Drawing.AdjustValueList>();
            if (avLst != null)
            {
                var guides = avLst.Elements<Drawing.ShapeGuide>()
                    .Where(g => g.Name?.HasValue == true && g.Formula?.HasValue == true)
                    .Select(g => $"{g.Name!.Value}:{g.Formula!.Value}")
                    .ToList();
                if (guides.Count > 0)
                    node.Format["adj"] = string.Join(",", guides);
            }
        }
        else
        {
            var custGeom = shape.ShapeProperties?.GetFirstChild<Drawing.CustomGeometry>();
            if (custGeom != null)
            {
                node.Format["geometry"] = "custom";
                // Raw OOXML preserves the full path-list (commands + adjust handles)
                // that ReconstructCustomGeometryPath's SVG-ish abstraction loses.
                // dump→batch needs byte-faithful replay, so emit the raw <a:custGeom>
                // XML alongside the SVG hint. Add side picks whichever it can parse.
                node.Format["customPath"] = ReconstructCustomGeometryPath(custGeom);
                node.Format["customGeometryXml"] = custGeom.OuterXml;
            }
        }

        // Gradient fill
        var gradFill = shape.ShapeProperties?.GetFirstChild<Drawing.GradientFill>();
        if (gradFill != null)
        {
            node.Format["gradient"] = ReadGradientString(gradFill);
            if (!node.Format.ContainsKey("fill"))
                node.Format["fill"] = "gradient";
            // bt-B2: when the source <a:gradFill> carries flip=... or a
            // child <a:tileRect>, the semantic gradient= form can't carry
            // them (BuildGradientFill never emits these). Surface the
            // verbatim element so Set's gradientRaw key can re-install it.
            if (HasGradientNonSemanticTuning(gradFill))
                node.Format["gradientRaw"] = gradFill.OuterXml;
        }

        // Image (blip) fill on shape
        var blipFill = shape.ShapeProperties?.GetFirstChild<Drawing.BlipFill>();
        if (blipFill != null)
        {
            node.Format["image"] = "true";
            // R9-7: surface the embedded image's file name on a companion key so
            // the readback is round-trippable, mirroring R4-7 background.src.
            // Keep Format["image"] == "true" for the long-standing bare contract.
            var blipEmbedId = blipFill.GetFirstChild<Drawing.Blip>()?.Embed?.Value;
            if (!string.IsNullOrEmpty(blipEmbedId) && part != null)
            {
                try
                {
                    var imgPart = part.GetPartById(blipEmbedId!);
                    var fileName = System.IO.Path.GetFileName(imgPart.Uri.ToString());
                    if (!string.IsNullOrEmpty(fileName))
                        node.Format["image.src"] = fileName;
                }
                catch { /* dangling rel — no src surfaced */ }
            }
            // Round-trip the blip fill's framing: the <a:srcRect> crop insets and
            // the <a:stretch><a:fillRect> stretch insets. ApplyShapeImageFill
            // formerly always wrote a child-less <a:fillRect/>, so an image
            // stretched past the shape bounds (a banner skyline with negative
            // fillRect t/b) snapped back to an exact-fit stretch and the framing
            // shifted on round-trip. Mirrors the PictureToNode fillRect/srcRect
            // readback so shape image fills frame identically.
            var shSrcRect = blipFill.GetFirstChild<Drawing.SourceRectangle>();
            if (shSrcRect != null)
            {
                var sl = shSrcRect.Left?.Value; var st = shSrcRect.Top?.Value;
                var sr = shSrcRect.Right?.Value; var sb = shSrcRect.Bottom?.Value;
                if (sl.HasValue || st.HasValue || sr.HasValue || sb.HasValue)
                    node.Format["srcRect"] = $"{sl ?? 0},{st ?? 0},{sr ?? 0},{sb ?? 0}";
            }
            var shFr = blipFill.GetFirstChild<Drawing.Stretch>()?.GetFirstChild<Drawing.FillRectangle>();
            if (shFr != null)
            {
                var fl = shFr.Left?.Value; var ft = shFr.Top?.Value;
                var frv = shFr.Right?.Value; var fb = shFr.Bottom?.Value;
                if (fl.HasValue || ft.HasValue || frv.HasValue || fb.HasValue)
                    node.Format["fillRect"] = $"{fl ?? 0},{ft ?? 0},{frv ?? 0},{fb ?? 0}";
            }
        }

        // Pattern fill on shape — round-trip the input form "preset:fg:bg".
        var patternFill = shape.ShapeProperties?.GetFirstChild<Drawing.PatternFill>();
        if (patternFill != null)
        {
            var preset = patternFill.Preset?.InnerText ?? "";
            var fgEl = patternFill.GetFirstChild<Drawing.ForegroundColor>();
            var bgEl = patternFill.GetFirstChild<Drawing.BackgroundColor>();
            var fgHex = fgEl?.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value;
            // Use the EnumValue's InnerText (the OOXML token, e.g. "tx1"); the
            // newer SDK backs SchemeColorValues with a struct whose Value.ToString()
            // is "SchemeColorValues { }", which is not a parseable color. Mirrors
            // the scheme-color readback in PowerPointHandler.Fill.cs.
            var fgScheme = fgEl?.GetFirstChild<Drawing.SchemeColor>()?.Val?.InnerText;
            var bgHex = bgEl?.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value;
            var bgScheme = bgEl?.GetFirstChild<Drawing.SchemeColor>()?.Val?.InnerText;
            var fg = fgHex != null ? ParseHelpers.FormatHexColor(fgHex) : (fgScheme ?? "");
            var bg = bgHex != null ? ParseHelpers.FormatHexColor(bgHex) : (bgScheme ?? "");
            node.Format["pattern"] = string.IsNullOrEmpty(bg) ? $"{preset}:{fg}" : $"{preset}:{fg}:{bg}";
            if (!node.Format.ContainsKey("fill"))
                node.Format["fill"] = "pattern";
        }

        // List style (from first paragraph)
        var firstParaBullet = shape.TextBody?.Elements<Drawing.Paragraph>().FirstOrDefault()?.ParagraphProperties;
        if (firstParaBullet != null)
        {
            var firstBulletRaw = ReadBulletRawFromPProps(firstParaBullet);
            if (firstBulletRaw != null)
            {
                node.Format["bulletRaw"] = firstBulletRaw;
                // R7-10: emit a re-feedable `list` companion alongside bulletRaw when
                // the bullet maps to a canonical keyword (suppressed for raw-char
                // passthroughs, preserving the bulletRaw-only contract for custom bullets).
                var firstListCanon = ReadCanonicalListKeyword(firstParaBullet);
                if (firstListCanon != null) node.Format["list"] = firstListCanon;
            }
            else
            {
                var firstList = ReadListStyleFromPProps(firstParaBullet);
                if (firstList != null) node.Format["list"] = firstList;
            }
        }

        // Collect font info
        var firstRun = shape.TextBody?.Descendants<Drawing.Run>().FirstOrDefault();
        // Heterogeneity probe: shape-level `size` and `color` summarize the
        // textbox's run formatting. When runs disagree, surfacing the first
        // run's value silently is a foot-gun — an agent reading
        // Format["color"] can't tell the textbox actually has mixed colors.
        // Drop the key in that case so `ContainsKey` is the contract.
        var allRuns = shape.TextBody?.Descendants<Drawing.Run>().ToList()
                      ?? new List<Drawing.Run>();
        bool hasMixedSize = false;
        bool hasMixedColor = false;
        // R60 fuzzer regression: bold/italic must also be heterogeneity-probed.
        // A textbox whose first run is bold (e.g. "AV AT WA — kern=0  ") and
        // whose sibling runs are roman silently surfaced bold=true at shape
        // level — dump→replay then emitted `add textbox bold=true` and the
        // shape-default styling bled bold onto every default (unstyled) run.
        // Same for italic on the subscript/superscript-alias textbox where
        // a leading italic run lifted italic onto siblings (Custom: / higher /
        // lower lost their roman style on replay).
        bool hasMixedBold = false;
        bool hasMixedItalic = false;
        if (allRuns.Count > 1)
        {
            string? firstSizeKey = null;
            string? firstColorKey = null;
            string? firstBoldKey = null;
            string? firstItalicKey = null;
            bool sizeSeen = false, colorSeen = false, boldSeen = false, italicSeen = false;
            foreach (var r in allRuns)
            {
                var rp = r.RunProperties;
                var sz = rp?.FontSize?.Value;
                var szKey = sz.HasValue ? sz.Value.ToString() : "(unset)";
                if (!sizeSeen) { firstSizeKey = szKey; sizeSeen = true; }
                else if (firstSizeKey != szKey) hasMixedSize = true;

                var col = ReadColorFromFill(rp?.GetFirstChild<Drawing.SolidFill>());
                var colKey = col ?? "(unset)";
                if (!colorSeen) { firstColorKey = colKey; colorSeen = true; }
                else if (firstColorKey != colKey) hasMixedColor = true;

                var boldKey = rp?.Bold?.HasValue == true ? (rp.Bold.Value ? "1" : "0") : "(unset)";
                if (!boldSeen) { firstBoldKey = boldKey; boldSeen = true; }
                else if (firstBoldKey != boldKey) hasMixedBold = true;

                var italicKey = rp?.Italic?.HasValue == true ? (rp.Italic.Value ? "1" : "0") : "(unset)";
                if (!italicSeen) { firstItalicKey = italicKey; italicSeen = true; }
                else if (firstItalicKey != italicKey) hasMixedItalic = true;
            }
        }
        if (firstRun?.RunProperties != null)
        {
            var fontLatinTf = firstRun.RunProperties.GetFirstChild<Drawing.LatinFont>()?.Typeface?.Value;
            var fontEaTf = firstRun.RunProperties.GetFirstChild<Drawing.EastAsianFont>()?.Typeface?.Value;
            var fontCsTf = firstRun.RunProperties.GetFirstChild<Drawing.ComplexScriptFont>()?.Typeface?.Value;
            // Bare `font` is the Latin slot alias only. Do NOT fall back to
            // <a:ea>/<a:cs> — those have their own canonical keys
            // (`font.ea`, `font.cs`) and bare `font` implying EA misrepresents
            // the OOXML. Suppressing bare `font` for ea/cs-only also keeps
            // `effective.font` (theme Latin) visible — symmetric with the
            // `font.cs`-only case.
            if (fontLatinTf != null) node.Format["font"] = fontLatinTf;
            // Per-script slots — emit canonical `font.latin` / `font.ea`
            // whenever the slot is present so schema-declared `get:true`
            // round-trips (CONSISTENCY(canonical-keys)). The redundant
            // `font` alias is kept for backward compat.
            if (fontLatinTf != null) node.Format["font.latin"] = fontLatinTf;
            // CONSISTENCY(per-script-round-trip): emit font.ea / font.cs
            // whenever the <a:ea>/<a:cs> element is present in source XML,
            // even when its typeface equals the latin slot. The previous
            // `fEa != fLatin` guard suppressed the readback for matching
            // typefaces — round-trip then dropped the explicit <a:ea>/<a:cs>
            // elements entirely. PowerPoint authors all three slots with
            // matching typefaces on every run; without this, dump→replay
            // lost every ea/cs element on Latin-only decks.
            if (fontEaTf != null) node.Format["font.ea"] = fontEaTf;
            if (fontCsTf != null) node.Format["font.cs"] = fontCsTf;

            var fontSize = firstRun.RunProperties.FontSize?.Value;
            if (fontSize.HasValue && !hasMixedSize) node.Format["size"] = $"{fontSize.Value / 100.0:0.##}pt";

            if (firstRun.RunProperties.Bold?.HasValue == true && !hasMixedBold) node.Format["bold"] = firstRun.RunProperties.Bold.Value;
            if (firstRun.RunProperties.Italic?.HasValue == true && !hasMixedItalic) node.Format["italic"] = firstRun.RunProperties.Italic.Value;
            // CONSISTENCY(rPr-cap): mirror cap attribute readback so shape-level
            // Get matches Set's allCaps/cap input (Set writes rPr cap="all"/"small").
            if (firstRun.RunProperties.Capital?.HasValue == true && firstRun.RunProperties.Capital.Value != Drawing.TextCapsValues.None)
                node.Format["cap"] = firstRun.RunProperties.Capital.InnerText;
            if (firstRun.RunProperties.Underline?.HasValue == true && firstRun.RunProperties.Underline.Value != Drawing.TextUnderlineValues.None)
            {
                var ulInner = firstRun.RunProperties.Underline.InnerText;
                node.Format["underline"] = ulInner switch
                {
                    "sng" => "single",
                    "dbl" => "double",
                    _ => ulInner
                };
            }
            // CONSISTENCY(underline-color): mirror the run-level reader so
            // shape-level Get also surfaces the underline color set on the
            // first run. Without this, `Add shape underline.color=...` round-
            // trips at the run scope only and Get on the shape drops it.
            var firstRunUFill = firstRun.RunProperties.GetFirstChild<Drawing.UnderlineFill>();
            if (firstRunUFill != null)
            {
                var firstRunUColor = ReadColorFromFill(firstRunUFill.GetFirstChild<Drawing.SolidFill>());
                if (firstRunUColor != null) node.Format["underline.color"] = firstRunUColor;
            }
            // R57 bt-1: <a:uLn> companion reader — see run-level emit below
            // for the rationale. Mirrors uFill at shape level so a single-run
            // shape with <a:uLn w=…> on the only run round-trips both color
            // and width through dump→replay.
            var firstRunULn = firstRun.RunProperties.GetFirstChild<Drawing.Underline>();
            if (firstRunULn != null)
            {
                if (!node.Format.ContainsKey("underline.color"))
                {
                    var firstRunULnColor = ReadColorFromFill(firstRunULn.GetFirstChild<Drawing.SolidFill>());
                    if (firstRunULnColor != null) node.Format["underline.color"] = firstRunULnColor;
                }
                if (firstRunULn.Width?.HasValue == true)
                    node.Format["underline.width"] = EmuConverter.FormatLineWidth(firstRunULn.Width.Value);
            }
            // BUG-PPTX-R2-03: mirror the run-level <a:ln> (text outline / glyph
            // stroke) reader at shape level — same pattern as uFill/uLn above.
            // Without this, `Add/Set shape textOutline=…` round-trips at the run
            // scope only and Get on the shape drops the key, so the Add path
            // (which forwards textOutline through effectKeys) looks broken.
            var firstRunOutline = firstRun.RunProperties.GetFirstChild<Drawing.Outline>();
            if (firstRunOutline != null)
            {
                string? toWidth = firstRunOutline.Width?.HasValue == true
                    ? EmuConverter.FormatLineWidth(firstRunOutline.Width.Value) : null;
                string? toColor = ReadColorFromFill(firstRunOutline.GetFirstChild<Drawing.SolidFill>());
                if (toWidth != null) node.Format["textOutline.width"] = toWidth;
                if (toColor != null) node.Format["textOutline.color"] = toColor;
                if (toWidth != null && toColor != null)
                    node.Format["textOutline"] = $"{toWidth}:{toColor}";
                else if (toWidth != null)
                    node.Format["textOutline"] = toWidth;
                else if (toColor != null)
                    node.Format["textOutline"] = toColor;
                else
                    node.Format["textOutline"] = "true";
                // Mirror the run-level textOutlineRaw carrier (dash/gradient/
                // cap-join beyond the width:color compound, sample10).
                if (firstRunOutline.ChildElements.Any(c => c is not Drawing.SolidFill)
                    || firstRunOutline.GetAttributes().Any(a => a.LocalName != "w"))
                    node.Format["textOutlineRaw"] = firstRunOutline.OuterXml;
            }
            if (firstRun.RunProperties.Strike?.HasValue == true)
            {
                // Emit explicit "none" too, so a round-trip Add(strike=none) → Get
                // returns the same key. PowerPoint writes <a:rPr strike="noStrike"/>
                // verbatim; dropping it silently breaks batch (dump | apply) parity.
                node.Format["strike"] = firstRun.RunProperties.Strike.Value switch
                {
                    var v when v == Drawing.TextStrikeValues.DoubleStrike => "double",
                    var v when v == Drawing.TextStrikeValues.NoStrike => "none",
                    _ => "single",
                };
            }
            // CONSISTENCY(highlight): mirror the run-level a:highlight reader at
            // shape level so `Add shape highlight=…` (first-run write) surfaces
            // on shape-level Get, same pattern as uFill/uLn/textOutline above.
            var firstRunHighlight = ReadColorFromHighlight(firstRun.RunProperties.GetFirstChild<Drawing.Highlight>());
            if (firstRunHighlight != null) node.Format["highlight"] = firstRunHighlight;

            // Character spacing on first run
            if (firstRun.RunProperties.Spacing?.HasValue == true)
                node.Format["spacing"] = $"{firstRun.RunProperties.Spacing.Value / 100.0:0.##}";
            // Baseline (superscript/subscript) — RUN-LEVEL ONLY.
            // R56 bt-3: baseline is a run-only rPr attribute (ECMA-376 §21.1.2.3.9);
            // lifting it onto the shape's Format conflates a per-run vertical
            // offset with the shape as a whole and round-trips as a bogus
            // `add shape baseline=…` op that PowerPoint's pPr / spPr schemas
            // don't accept. Suppress at shape level — the per-run reader below
            // still surfaces it correctly under the run node.

            // Text color (from first run) — solid or gradient
            var runColor = ReadColorFromFill(firstRun.RunProperties.GetFirstChild<Drawing.SolidFill>());
            if (runColor != null && !hasMixedColor) node.Format["color"] = runColor;
            var runGradFill = firstRun.RunProperties.GetFirstChild<Drawing.GradientFill>();
            if (runGradFill != null)
                node.Format["textFill"] = ReadGradientString(runGradFill);

            // Hyperlink on first run (link + tooltip — tooltip mirrors how
            // picture / group already round-trip, see below + line ~1262).
            if (part != null)
            {
                var firstHl = firstRun.RunProperties.GetFirstChild<Drawing.HyperlinkOnClick>();
                var linkUrl = ReadHyperlinkOnClickUrl(firstHl, part);
                if (linkUrl != null) node.Format["link"] = linkUrl;
                var firstTip = firstHl?.Tooltip?.Value;
                if (!string.IsNullOrEmpty(firstTip)) node.Format["tooltip"] = firstTip!;
            }

            // CONSISTENCY(rpr-attr-fallback / R21-fuzzer-1+2): surface long-tail
            // rPr attributes (kern, kumimoji, normalizeH, ...) at shape
            // level too, mirroring BuildRunNode. Without this, shape-level Add
            // can write a long-tail attr to first-run rPr but shape-level Get
            // cannot surface it unless the user descends to /shape[N]/r[1].
            // bt-6: rPr-only attrs (err, dirty, smtClean, lang) are filtered
            // out — they describe the run's editor state, not the shape.
            FillUnknownRunProps(firstRun.RunProperties, node, shapeLevel: true);
        }
        else
        {
            // Runless shape — `add shape --prop bold=… --prop size=… --prop color=…
            // --prop font=…` lands those defaults on the first paragraph's
            // <a:endParaRPr> (since there is no <a:r> to host an rPr). Surface
            // them at shape level so dump→replay round-trips byte-equal.
            var endRPrShape = shape.TextBody?.Elements<Drawing.Paragraph>()
                .FirstOrDefault()?.GetFirstChild<Drawing.EndParagraphRunProperties>();
            if (endRPrShape != null)
            {
                var epLatin = endRPrShape.GetFirstChild<Drawing.LatinFont>()?.Typeface?.Value;
                var epEa = endRPrShape.GetFirstChild<Drawing.EastAsianFont>()?.Typeface?.Value;
                var epCs = endRPrShape.GetFirstChild<Drawing.ComplexScriptFont>()?.Typeface?.Value;
                if (epLatin != null) { node.Format["font"] = epLatin; node.Format["font.latin"] = epLatin; }
                if (epEa != null) node.Format["font.ea"] = epEa;
                if (epCs != null) node.Format["font.cs"] = epCs;
                if (endRPrShape.FontSize?.Value is { } epSz)
                    node.Format["size"] = $"{epSz / 100.0:0.##}pt";
                if (endRPrShape.Bold?.HasValue == true)
                    node.Format["bold"] = endRPrShape.Bold.Value;
                if (endRPrShape.Italic?.HasValue == true)
                    node.Format["italic"] = endRPrShape.Italic.Value;
                // CONSISTENCY(rpr-attr-fallback): underline/strike on runless
                // shapes land on endParaRPr (Add's RunPropTargets fallback,
                // same as bold/italic) — surface them here, mirroring the
                // run-present readers at line ~1141/~1198.
                if (endRPrShape.Underline?.HasValue == true && endRPrShape.Underline.Value != Drawing.TextUnderlineValues.None)
                {
                    var epUlInner = endRPrShape.Underline.InnerText;
                    node.Format["underline"] = epUlInner switch
                    {
                        "sng" => "single",
                        "dbl" => "double",
                        _ => epUlInner
                    };
                }
                if (endRPrShape.Strike?.HasValue == true)
                {
                    node.Format["strike"] = endRPrShape.Strike.Value switch
                    {
                        var v when v == Drawing.TextStrikeValues.DoubleStrike => "double",
                        var v when v == Drawing.TextStrikeValues.NoStrike => "none",
                        _ => "single",
                    };
                }
                // BUG3: cap readback on runless shapes — mirror the run-present
                // path at line ~1139 which uses Capital.InnerText ("all"/"small").
                var epCapAttr = endRPrShape.GetAttributes().FirstOrDefault(a => a.LocalName == "cap");
                if (epCapAttr.Value != null && epCapAttr.Value != "none")
                    node.Format["cap"] = epCapAttr.Value;
                var epColor = ReadColorFromFill(endRPrShape.GetFirstChild<Drawing.SolidFill>());
                if (epColor != null) node.Format["color"] = epColor;
            }
        }

        // Shape-level hyperlink (on NonVisualDrawingProperties). Route through
        // the shared ReadHyperlinkOnClickUrl helper so named-action targets
        // (firstslide/lastslide/nextslide/previousslide) and internal slide
        // jumps (slide[N]) round-trip — the previous inline reader only saw
        // external HyperlinkRelationship URIs.
        if (part != null && !node.Format.ContainsKey("link"))
        {
            var nvDp = shape.NonVisualShapeProperties?.NonVisualDrawingProperties;
            var hlClick = nvDp?.GetFirstChild<Drawing.HyperlinkOnClick>();
            var shapeLinkUrl = ReadHyperlinkOnClickUrl(hlClick, part);
            if (shapeLinkUrl != null) node.Format["link"] = shapeLinkUrl;
            var shapeTip = hlClick?.Tooltip?.Value;
            if (!string.IsNullOrEmpty(shapeTip) && !node.Format.ContainsKey("tooltip"))
                node.Format["tooltip"] = shapeTip!;
        }

        // Line/border
        var outline = shape.ShapeProperties?.GetFirstChild<Drawing.Outline>();
        if (outline != null)
        {
            var lineSolidFill = outline.GetFirstChild<Drawing.SolidFill>();
            var lineColor = ReadColorFromFill(lineSolidFill);
            if (lineColor != null) node.Format["line"] = lineColor;
            var lineIsNone = outline.GetFirstChild<Drawing.NoFill>() != null;
            if (lineIsNone) node.Format["line"] = "none";
            // Gradient on the line — round-trippable spec form so dump→batch
            // replay rebuilds the gradient instead of falling back to bare
            // <a:ln/> (which inherits theme's default thin black stroke).
            var lineGradFill = outline.GetFirstChild<Drawing.GradientFill>();
            if (lineGradFill != null)
            {
                node.Format["line.gradient"] = ReadGradientString(lineGradFill);
            }
            // When line=none, suppress the residual width readback so users don't
            // see a stale lineWidth from a prior color-set assignment.
            if (!lineIsNone && outline.Width?.HasValue == true) node.Format["lineWidth"] = FormatLineWidth(outline.Width.Value);
            var dash = outline.GetFirstChild<Drawing.PresetDash>();
            if (dash?.Val?.HasValue == true)
            {
                // emit the canonical OOXML token (lgDash/lgDashDot/lgDashDotDot/
                // sysDot/sysDash/sysDashDot/sysDashDotDot) so the readback survives a
                // round-trip through the input parser. Previously emitted 'longdash[dot]'
                // aliases that aren't accepted by the (case-strict) PresetLineDashValues
                // constructor — a future round-trip would throw.
                var dashValue = dash.Val.InnerText ?? "";
                node.Format["lineDash"] = dashValue switch
                {
                    "solid" => "solid",
                    "dot" => "dot",
                    "dash" => "dash",
                    "dashDot" => "dashDot",
                    "lgDash" => "lgDash",
                    "lgDashDot" => "lgDashDot",
                    "lgDashDotDot" => "lgDashDotDot",
                    "sysDot" => "sysDot",
                    "sysDash" => "sysDash",
                    "sysDashDot" => "sysDashDot",
                    "sysDashDotDot" => "sysDashDotDot",
                    _ => dashValue
                };
            }
            // R64 bt-3: <a:custDash> sibling to (not nested in) prstDash —
            // mirror connector readback so shape outlines with author-defined
            // dash patterns round-trip via lineDashRaw passthrough.
            var custDash = outline.GetFirstChild<Drawing.CustomDash>();
            if (custDash != null)
            {
                node.Format["lineDashRaw"] = custDash.OuterXml;
            }
            // lineCap / lineJoin / cmpd / lineAlign readback. Previously
            // these attributes were accepted on input but silently dropped; the
            // bidirectional gap meant users couldn't see whether the value stuck.
            if (outline.CapType?.HasValue == true)
            {
                var capRaw = outline.CapType.InnerText ?? "";
                node.Format["lineCap"] = capRaw switch
                {
                    "rnd" => "round",
                    "sq" => "square",
                    "flat" => "flat",
                    _ => capRaw
                };
            }
            if (outline.CompoundLineType?.HasValue == true)
                node.Format["cmpd"] = outline.CompoundLineType.InnerText ?? "";
            if (outline.Alignment?.HasValue == true)
                node.Format["lineAlign"] = outline.Alignment.InnerText ?? "";
            if (outline.GetFirstChild<Drawing.Round>() != null)
                node.Format["lineJoin"] = "round";
            else if (outline.GetFirstChild<Drawing.LineJoinBevel>() != null)
                node.Format["lineJoin"] = "bevel";
            else
            {
                // R61 bt-2: surface the miter limit (`<a:miter lim="N"/>` attribute)
                // alongside lineJoin=miter — previously only the bare join token
                // was emitted and lim was silently dropped on round-trip.
                var shapeMiterEl = outline.GetFirstChild<Drawing.Miter>();
                if (shapeMiterEl != null)
                {
                    node.Format["lineJoin"] = "miter";
                    if (shapeMiterEl.Limit?.HasValue == true)
                        node.Format["miterLimit"] = shapeMiterEl.Limit.Value;
                }
            }
            // head/tail arrowheads on shape outlines.
            var shapeHeadEnd = outline.GetFirstChild<Drawing.HeadEnd>();
            if (shapeHeadEnd?.Type?.HasValue == true)
                node.Format["headEnd"] = shapeHeadEnd.Type.InnerText ?? "";
            var shapeTailEnd = outline.GetFirstChild<Drawing.TailEnd>();
            if (shapeTailEnd?.Type?.HasValue == true)
                node.Format["tailEnd"] = shapeTailEnd.Type.InnerText ?? "";
            var lineColorEl = lineSolidFill?.GetFirstChild<Drawing.RgbColorModelHex>() as OpenXmlElement
                ?? lineSolidFill?.GetFirstChild<Drawing.SchemeColor>();
            var lineAlpha = lineColorEl?.GetFirstChild<Drawing.Alpha>()?.Val?.Value;
            if (lineAlpha.HasValue) node.Format["lineOpacity"] = $"{lineAlpha.Value / 100000.0:0.##}";
        }

        // Effects (shadow, glow, reflection) — shape level ONLY. The shape's
        // <p:spPr><a:effectLst> is the only source for shape-level effect=
        // keys; run-level <a:rPr><a:effectLst> rides through RunToNode below.
        // bt-1: an earlier "fall back to first run's effectLst when shape has
        // no fill" heuristic mis-promoted run-level shadow/glow onto the
        // textbox shape AND (via single-run collapse) onto the paragraph,
        // producing duplicate writes on replay. Each level reads its own
        // effectLst now; CONSISTENCY(run-context-explicit) holds end-to-end.
        var effectList = shape.ShapeProperties?.GetFirstChild<Drawing.EffectList>();
        var activeEffectList = effectList?.HasChildren == true ? effectList : null;
        if (activeEffectList != null)
        {
            var outerShadow = activeEffectList.GetFirstChild<Drawing.OuterShadow>();
            if (outerShadow != null)
            {
                var shadowColor = EnsureEightDigitHexForEffect(ReadColorFromElement(outerShadow) ?? "000000");
                var blurPt = outerShadow.BlurRadius?.HasValue == true ? $"{outerShadow.BlurRadius.Value / EmuConverter.EmuPerPointF:0.##}" : "4";
                var angleDeg = outerShadow.Direction?.HasValue == true ? $"{outerShadow.Direction.Value / 60000.0:0.##}" : "45";
                var distPt = outerShadow.Distance?.HasValue == true ? $"{outerShadow.Distance.Value / EmuConverter.EmuPerPointF:0.##}" : "3";
                var alphaEl = outerShadow.Descendants<Drawing.Alpha>().FirstOrDefault();
                // OOXML default: <a:outerShdw> without <a:alpha> is fully opaque
                // (the shadow inherits the color element's alpha; an absent alpha
                // means 100%). Defaulting to "40" used to mask explicit
                // alpha=FF inputs as a 40% shadow on round-trip.
                var opacity = alphaEl?.Val?.HasValue == true ? $"{alphaEl.Val.Value / 1000.0:0.##}" : "100";
                node.Format["shadow"] = $"{shadowColor}-{blurPt}-{angleDeg}-{distPt}-{opacity}";
                // bt-2: a source-authored <a:outerShdw> can carry scale/skew
                // attributes (sx/sy stretch the shadow, kx/ky skew it, algn
                // anchors the projection, rotWithShape toggles whether the
                // shadow rotates with its host). BuildOuterShadow only knows
                // the COLOR-BLUR-ANGLE-DIST-OPACITY tuple, so re-build drops
                // every extra attr on the floor. Mirror reflectionRaw — when
                // any non-compressible attr is present, surface the element's
                // OuterXml as `shadowRaw` so Set can re-install it verbatim.
                if (!IsPlainOuterShadow(outerShadow))
                    node.Format["shadowRaw"] = outerShadow.OuterXml;
            }
            // bt-1: surface <a:innerShdw> as `innerShadow=COLOR-BLUR-ANGLE-DIST-OPACITY`
            // (parallel to `shadow=` for outer). Without this readback the dump
            // path silently dropped every inner-shadow effect; replay then
            // produced a flat shape with no inset highlight (Apple-style
            // pressed-button look, decorative card insets).
            var innerShadow = activeEffectList.GetFirstChild<Drawing.InnerShadow>();
            if (innerShadow != null)
            {
                var isColor = EnsureEightDigitHexForEffect(ReadColorFromElement(innerShadow) ?? "000000");
                var isBlur = innerShadow.BlurRadius?.HasValue == true ? $"{innerShadow.BlurRadius.Value / EmuConverter.EmuPerPointF:0.##}" : "4";
                var isAngle = innerShadow.Direction?.HasValue == true ? $"{innerShadow.Direction.Value / 60000.0:0.##}" : "45";
                var isDist = innerShadow.Distance?.HasValue == true ? $"{innerShadow.Distance.Value / EmuConverter.EmuPerPointF:0.##}" : "3";
                var isAlpha = innerShadow.Descendants<Drawing.Alpha>().FirstOrDefault();
                var isOpacity = isAlpha?.Val?.HasValue == true ? $"{isAlpha.Val.Value / 1000.0:0.##}" : "100";
                node.Format["innerShadow"] = $"{isColor}-{isBlur}-{isAngle}-{isDist}-{isOpacity}";
                // R56 bt-2: parallel to outer shadowRaw — when the inner
                // shadow's color child carries lumMod/lumOff/shade/tint/
                // satMod/hueMod transforms, the compressed innerShadow=
                // form embeds them via "+lumMod50+lumOff50" suffix syntax,
                // then appends "-BLUR-ANGLE-DIST-OPACITY". The dash-tail
                // after a transform chain is undocumented; surface OuterXml
                // so Set can re-install the source <a:innerShdw> verbatim.
                if (!IsPlainInnerShadow(innerShadow))
                    node.Format["innerShadowRaw"] = innerShadow.OuterXml;
            }
            var glow = activeEffectList.GetFirstChild<Drawing.Glow>();
            if (glow != null)
            {
                var glowColor = EnsureEightDigitHexForEffect(ReadColorFromElement(glow) ?? "000000");
                var radiusPt = glow.Radius?.HasValue == true ? $"{glow.Radius.Value / EmuConverter.EmuPerPointF:0.##}" : "8";
                var glowAlpha = glow.Descendants<Drawing.Alpha>().FirstOrDefault();
                // OOXML default: <a:glow> without <a:alpha> is fully opaque.
                var glowOpacity = glowAlpha?.Val?.HasValue == true ? $"{glowAlpha.Val.Value / 1000.0:0.##}" : "100";
                node.Format["glow"] = $"{glowColor}-{radiusPt}-{glowOpacity}";
            }
            // bt-1: <a:fillOverlay blend="darken|lighten|mult|over|screen"><a:solidFill>…</>
            // is a composited fill on top of the shape's main fill — modern
            // PowerPoint decks use it for theme-tinted card insets. The
            // effectLst walker only consumed the well-known outerShdw / glow /
            // reflection / softEdge children, so fillOverlay was silently
            // dropped on dump→replay. Surface OuterXml as fillOverlayRaw
            // (passthrough mirrors shadowRaw / reflectionRaw); a compressed
            // `fillOverlay=blend:color` form would lose alpha/gradient overlay
            // variants, so raw is the conservative choice.
            var fillOverlay = activeEffectList.GetFirstChild<Drawing.FillOverlay>();
            if (fillOverlay != null)
                node.Format["fillOverlayRaw"] = fillOverlay.OuterXml;
            var reflEl = activeEffectList.GetFirstChild<Drawing.Reflection>();
            if (reflEl != null)
            {
                // CONSISTENCY(reflection-exact-match): Set accepts both named
                // presets (tight/half/full → 55000/90000/100000) and bare
                // percent (0-100 → pct*1000). The previous bucketed readback
                // (`>=95000 full, >=70000 half, else tight`) made every
                // non-preset numeric round-trip as the nearest preset name,
                // silently rewriting the user's input. Now: exact-preset
                // values emit the preset name, everything else emits the
                // integer percent the Set side accepts.
                //
                // bt-B1: a source-authored <a:reflection> can carry arbitrary
                // tuning attributes (blurRad, stA, endA, dist, dir, fadeDir,
                // sx/sy, kx/ky, algn, …). ApplyReflection rebuilds the
                // element from just the preset bucket and drops anything
                // outside its known field set, so a custom-tuned reflection
                // collapses to the nearest preset. Detect "this isn't a
                // bare preset" by comparing the captured attribute set
                // against ApplyReflection's emit (stA=52000 endA=300, the
                // matching endPos, all other attrs default) — when any
                // additional attribute is set, surface the element's
                // OuterXml as `reflectionRaw` so Set can re-install it
                // verbatim. The preset key is still emitted alongside for
                // human readability / non-passthrough consumers.
                var endPos = reflEl.EndPosition?.Value ?? 0;
                node.Format["reflection"] = endPos switch
                {
                    55000 => "tight",
                    90000 => "half",
                    100000 => "full",
                    _ => (endPos / 1000).ToString(System.Globalization.CultureInfo.InvariantCulture),
                };
                bool isPlainPreset = IsPlainReflectionPreset(reflEl);
                if (!isPlainPreset)
                    node.Format["reflectionRaw"] = reflEl.OuterXml;
            }
            var softEdge = activeEffectList.GetFirstChild<Drawing.SoftEdge>();
            if (softEdge?.Radius?.HasValue == true)
                // Unit-qualified pt — matches the cross-format canonical from
                // the project conventions (line.width "0.75pt", padding "12pt", glow
                // "4pt"). The bare numeric form here was the lone outlier on
                // the effects readback surface and broke dump round-trip
                // when set softEdge=<value> re-parses the readback.
                node.Format["softEdge"] = $"{softEdge.Radius.Value / EmuConverter.EmuPerPointF:0.##}pt";

            // R58 bt-1: shape-level <a:blur rad="…" grow="…"/> on spPr/effectLst
            // was silently dropped on dump→replay — Set/Add already supported
            // the `blur` key (ApplyBlur in Effects.cs), but NodeBuilder never
            // emitted it, so source-authored shape blur (decorative depth-of-
            // field, frosted-card looks) round-tripped to nothing. Emit as
            // `blur=<rad>pt:<grow>` so the grow attribute (default true in the
            // OOXML schema, but explicitly authored in real decks) survives.
            var blur = activeEffectList.GetFirstChild<Drawing.Blur>();
            if (blur != null)
            {
                var radPt = blur.Radius?.HasValue == true
                    ? $"{blur.Radius.Value / EmuConverter.EmuPerPointF:0.##}pt"
                    : "0pt";
                // OOXML default for <a:blur grow> is true; emit the boolean
                // verbatim so an explicit `grow="0"` survives the round-trip.
                var grow = blur.Grow?.HasValue == true ? blur.Grow.Value : true;
                node.Format["blur"] = $"{radPt}:{(grow ? "true" : "false")}";
            }

            // R58 bt-2: surface the entire effectLst as `effectsRaw=<OuterXml>`
            // when it contains any child the compressed walker above does not
            // consume. Source decks carry tint / lum / hsl / alphaModFix /
            // clrChange / duotone / etc. children on the spPr effectLst — the
            // reported case was `<a:effectLst><a:tint amt="50000"/></a:effectLst>`
            // — and the walker silently drops them on dump-then-replay. Mirror
            // the effectDagRaw / fillOverlayRaw passthrough: a verbatim XML
            // surface is the only round-trip-safe form, since these children
            // have no compressible string representation. On Set, effectsRaw
            // replaces the entire effectLst (it wins over the compressed
            // shadow= / glow= / reflection= / softEdge= / blur= keys that may
            // also be emitted from the same source effectLst).
            // CONSISTENCY(effects-raw-passthrough): mirrors effectDagRaw / fillOverlayRaw.
            if (HasUnhandledEffectListChild(activeEffectList))
                node.Format["effectsRaw"] = activeEffectList.OuterXml;
        }

        // R52 bt-2: surface <a:effectDag> as `effectDagRaw=<OuterXml>`.
        // EffectDag is the advanced compositing tree (cont/sib containers with
        // blur + fillOverlay + outerShdw, used for layered Aero-style glass /
        // depth treatments). It's a sibling to effectLst on the spPr, and
        // PowerPoint accepts both — the spec composes them in document order.
        // The walker above only consumes effectLst children, so any source-
        // authored effectDag was silently dropped on dump→replay. Mirror the
        // shadowRaw / fillOverlayRaw passthrough — there is no compressible
        // string form that survives the cont/sib nesting + blur radius +
        // fillOverlay blend modes a real effectDag can carry.
        var effectDag = shape.ShapeProperties?.GetFirstChild<Drawing.EffectDag>();
        if (effectDag != null)
            node.Format["effectDagRaw"] = effectDag.OuterXml;

        // 3D rotation (scene3d)
        var scene3d = shape.ShapeProperties?.GetFirstChild<Drawing.Scene3DType>();
        if (scene3d != null)
        {
            var cam = scene3d.Camera;
            var rot3d = cam?.Rotation;
            if (rot3d != null)
            {
                var rx = rot3d.Latitude?.Value ?? 0;
                var ry = rot3d.Longitude?.Value ?? 0;
                var rz = rot3d.Revolution?.Value ?? 0;
                if (rx != 0 || ry != 0 || rz != 0)
                    node.Format["rot3d"] = $"{rx / 60000.0:0.##},{ry / 60000.0:0.##},{rz / 60000.0:0.##}";
            }
            // Camera preset (a:scene3d/a:camera/@prst). Previously dropped on
            // round-trip — EnsureScene3D hard-coded OrthographicFront, so any
            // source preset (perspectiveContrastingRightFacing, isometricTopUp,
            // ...) was silently rewritten to the default. Emit the OOXML
            // camelCase name verbatim; Set re-applies via PresetCameraValues.
            if (cam?.Preset?.HasValue == true)
                node.Format["camera"] = cam.Preset.InnerText ?? "";
            var lightRig = scene3d.LightRig;
            if (lightRig?.Rig?.HasValue == true) node.Format["lighting"] = lightRig.Rig.InnerText;
            // R55 bt-4: surface lightRig @dir + <a:rot> child. Both are
            // schema-legal extensions of <a:lightRig>, and PowerPoint writes
            // them on the Reflection (light-rotation) Effect Options. Before
            // this readback, ApplyLightRig wrote a fresh <a:lightRig> with
            // only @rig, dropping any source dir / rot child on round-trip.
            if (lightRig?.Direction?.HasValue == true)
                node.Format["lightingDir"] = lightRig.Direction.InnerText;
            var lightRot = lightRig?.Rotation;
            if (lightRot != null)
            {
                var lat = lightRot.Latitude?.Value ?? 0;
                var lon = lightRot.Longitude?.Value ?? 0;
                var rev = lightRot.Revolution?.Value ?? 0;
                node.Format["lightingRot"] = $"{lat}:{lon}:{rev}";
            }
        }

        // 3D format (sp3d)
        var sp3d = shape.ShapeProperties?.GetFirstChild<Drawing.Shape3DType>();
        if (sp3d != null)
        {
            if (sp3d.ExtrusionHeight?.HasValue == true && sp3d.ExtrusionHeight.Value != 0)
                node.Format["depth"] = $"{sp3d.ExtrusionHeight.Value / EmuConverter.EmuPerPointF:0.##}";
            if (sp3d.PresetMaterial?.HasValue == true)
                node.Format["material"] = sp3d.PresetMaterial.InnerText;
            var bevelT = sp3d.BevelTop;
            if (bevelT != null) node.Format["bevel"] = FormatBevel(bevelT);
            var bevelB = sp3d.BevelBottom;
            if (bevelB != null) node.Format["bevelBottom"] = FormatBevel(bevelB);
            // extrusionClr / contourClr — additional sp3d children. The
            // sp3d emit dropped them silently before this readback; with no
            // Format key surfacing the value, the Setter had nothing to
            // round-trip and the colored extrusion edge rendered black.
            var extColor = ReadColorFromElement(sp3d.ExtrusionColor);
            if (!string.IsNullOrEmpty(extColor))
                node.Format["extrusionColor"] = extColor;
            var contColor = ReadColorFromElement(sp3d.ContourColor);
            if (!string.IsNullOrEmpty(contColor))
                node.Format["contourColor"] = contColor;
        }

        // Flip
        if (xfrm?.HorizontalFlip?.Value == true) node.Format["flipH"] = true;
        if (xfrm?.VerticalFlip?.Value == true) node.Format["flipV"] = true;

        // Z-order (1-based position among content elements: 1 = back, N = front)
        if (shape.Parent is ShapeTree zTree)
        {
            var contentEls = zTree.ChildElements
                .Where(e => e is Shape or Picture or GraphicFrame or GroupShape or ConnectionShape)
                .ToList();
            var zIdx = contentEls.IndexOf(shape);
            if (zIdx >= 0) node.Format["zorder"] = zIdx + 1;
        }

        // Rotation (plain number in degrees, no suffix, so Set can consume the value directly)
        if (xfrm?.Rotation != null && xfrm.Rotation.Value != 0)
            node.Format["rotation"] = $"{xfrm.Rotation.Value / 60000.0:0.######}";

        // Text margin
        var bodyPr = shape.TextBody?.Elements<Drawing.BodyProperties>().FirstOrDefault();
        if (bodyPr != null)
        {
            // Textbox-level RTL (a:bodyPr rtlCol). OpenXml SDK doesn't expose
            // rtlCol as a typed property AND GetAttribute(localName, ns)
            // THROWS KeyNotFoundException when the attribute is absent, so
            // iterate the attribute list to find rtlCol safely.
            string? rtlColAttr = null;
            foreach (var attr in bodyPr.GetAttributes())
            {
                if (attr.LocalName == "rtlCol") { rtlColAttr = attr.Value; break; }
            }
            if (!string.IsNullOrEmpty(rtlColAttr) && !node.Format.ContainsKey("direction"))
            {
                bool rtlColOn = rtlColAttr == "1" || rtlColAttr.Equals("true", StringComparison.OrdinalIgnoreCase);
                node.Format["direction"] = rtlColOn ? "rtl" : "ltr";
            }

            var lIns = bodyPr.LeftInset;
            var tIns = bodyPr.TopInset;
            var rIns = bodyPr.RightInset;
            var bIns = bodyPr.BottomInset;
            if (lIns != null || tIns != null || rIns != null || bIns != null)
            {
                // If all four are the same, show as single value
                if (lIns == tIns && tIns == rIns && rIns == bIns && lIns != null)
                    node.Format["margin"] = FormatEmu(lIns.Value);
                else
                    // CONSISTENCY(margin-sparse-roundtrip): emit "-" for any
                    // side absent on the source bodyPr instead of substituting
                    // the PowerPoint default (91440/45720). ApplyTextMargin's
                    // 4-form path treats "-" as "leave unset" so dump->replay
                    // preserves a source bodyPr like lIns/tIns/rIns-only
                    // without fabricating a default bIns="45720".
                    node.Format["margin"] = string.Join(",",
                        new[] { lIns, tIns, rIns, bIns }.Select(
                            v => v != null && v.HasValue ? FormatEmu(v.Value) : "-"));
            }

            // Vertical alignment — map XML enum to user-friendly name.
            // CONSISTENCY(valign-vocab): shape vertical-anchor input accepts both
            // "middle" and "center" (see ExcelHandler.Add.Drawings.cs:693 etc.),
            // but the OOXML enum is "ctr" — historically readback emitted "center"
            // which broke `set valign=middle` → `get valign=middle` round-trips.
            // Emit "middle" to match the canonical input alias users actually type
            // for the vertical axis. Table-cell valign (NodeBuilder.cs:391-394)
            // is a separate code path and keeps "center" since cells share the
            // halign/valign vocabulary.
            if (bodyPr.Anchor?.HasValue == true)
            {
                var vaInner = bodyPr.Anchor.InnerText;
                node.Format["valign"] = vaInner switch
                {
                    "t" => "top",
                    "ctr" => "middle",
                    "b" => "bottom",
                    _ => vaInner
                };
            }

            // Text direction (a:bodyPr @vert). Mirrors the cell-level reader at
            // line 344. Set accepts vertical90 / vertical270 / stacked etc; Get
            // must surface them so round-trip works.
            if (bodyPr.Vertical?.HasValue == true)
            {
                node.Format["textdirection"] = bodyPr.Vertical.InnerText switch
                {
                    "horz" => "horizontal",
                    "vert" => "vertical90",
                    "vert270" => "vertical270",
                    "wordArtVert" => "stacked",
                    "eaVert" => "eaVert",
                    "mongolianVert" => "mongolianVert",
                    "wordArtVertRtl" => "wordArtVertRtl",
                    _ => bodyPr.Vertical.InnerText
                };
            }

            // Multi-column text (a:bodyPr @numCol / @spcCol). Set writes
            // ColumnCount/ColumnSpacing; Get must surface them so dump→replay
            // and the HTML preview round-trip the column layout.
            if (bodyPr.ColumnCount?.HasValue == true && bodyPr.ColumnCount.Value > 1)
            {
                node.Format["columns"] = bodyPr.ColumnCount.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
                if (bodyPr.ColumnSpacing?.HasValue == true)
                    node.Format["columnSpacing"] = $"{Units.EmuToPt(bodyPr.ColumnSpacing.Value):0.##}pt";
            }

            // TextWarp (WordArt)
            var prstTxWarp = bodyPr.GetFirstChild<Drawing.PresetTextWarp>();
            if (prstTxWarp?.Preset?.HasValue == true)
            {
                node.Format["textWarp"] = prstTxWarp.Preset.InnerText;
                // Verbatim form when the warp carries adjust values — the
                // preset-name-only readback loses <a:avLst><a:gd fmla=…>,
                // flattening the curve amount on dump→replay.
                if (prstTxWarp.AdjustValueList?.HasChildren == true)
                    node.Format["textWarpRaw"] = prstTxWarp.OuterXml;
            }

            // 3D text: <a:scene3d>/<a:sp3d> INSIDE bodyPr (camera/lighting +
            // extrusion/bevel — distinct from the shape-level pair on spPr).
            // Previously dropped entirely, so a WordArt-3D deck replayed flat
            // (sample12). Verbatim raw carriers, spliced back by the
            // textScene3dRaw / textSp3dRaw Set cases.
            var bodyScene3d = bodyPr.GetFirstChild<Drawing.Scene3DType>();
            if (bodyScene3d != null)
                node.Format["textScene3dRaw"] = bodyScene3d.OuterXml;
            var bodySp3d = bodyPr.GetFirstChild<Drawing.Shape3DType>();
            if (bodySp3d != null)
                node.Format["textSp3dRaw"] = bodySp3d.OuterXml;

            // Word-wrap (a:bodyPr @wrap = "square" | "none"). Set already
            // accepts wrap=true/false and writes Square/None; Get must
            // surface it so dump→replay preserves the attribute. Match the
            // cell-level reader at line 397.
            if (bodyPr.Wrap?.HasValue == true)
                node.Format["wrap"] = bodyPr.Wrap.Value != Drawing.TextWrappingValues.None;

            // anchorCtr (center the text BLOCK horizontally) and upright
            // (keep glyphs upright inside a rotated body) — simple bodyPr
            // attributes previously dropped on dump→replay (sample16 /
            // sample13: anchor-centered eaVert columns drifted, rotated
            // upright text flipped).
            if (bodyPr.AnchorCenter?.HasValue == true)
                node.Format["anchorCtr"] = bodyPr.AnchorCenter.Value;
            if (bodyPr.UpRight?.HasValue == true)
                node.Format["upright"] = bodyPr.UpRight.Value;
            // Overflow clipping (clip-vertical-overflow: vertOverflow="clip"
            // dropped → overflowing text replayed visible).
            if (bodyPr.VerticalOverflow?.HasValue == true)
                node.Format["vertOverflow"] = bodyPr.VerticalOverflow.InnerText;
            if (bodyPr.HorizontalOverflow?.HasValue == true)
                node.Format["horzOverflow"] = bodyPr.HorizontalOverflow.InnerText;

            // AutoFit — surface only when the source bodyPr carries an
            // explicit child. An empty <a:bodyPr/> inherits from the
            // layout/master cascade; emitting "none" as the default forces
            // <a:noAutoFit/> on replay (Add/Set autoFit=none injects the
            // child), breaking shrink-to-fit on every placeholder dump→
            // replay cycle. CONSISTENCY(empty-bodyPr-inherits).
            var normAutoFit = bodyPr.GetFirstChild<Drawing.NormalAutoFit>();
            if (normAutoFit != null)
            {
                node.Format["autoFit"] = "normal";
                // Preserve the shrink-on-overflow scale PowerPoint computed and
                // stored (fontScale="92500" = 92.5%, lnSpcReduction similar).
                // Without this the box rebuilt at 100% and text re-flowed across
                // the whole deck. OOXML thousandths-of-percent, emitted raw.
                if (normAutoFit.FontScale?.Value is int afFs && afFs != 100000)
                    node.Format["fontScale"] = afFs;
                if (normAutoFit.LineSpaceReduction?.Value is int afLr && afLr != 0)
                    node.Format["lnSpcReduction"] = afLr;
            }
            else if (bodyPr.GetFirstChild<Drawing.ShapeAutoFit>() != null) node.Format["autoFit"] = "shape";
            else if (bodyPr.GetFirstChild<Drawing.NoAutoFit>() != null) node.Format["autoFit"] = "none";
        }

        // Shape-level txBody <a:lstStyle> with content (lvl1pPr..lvl9pPr:
        // lnSpc/defTabSz/algn/fonts). The NodeBuilder models per-paragraph
        // props but never the txBody lstStyle, so rebuild emitted an empty
        // <a:lstStyle/> and text reflowed off the source's per-level metrics.
        // Surface the whole element's OuterXml verbatim so AddShape can
        // re-inject it in CT_TextBody order (after bodyPr, before first p).
        // Skip empty stubs (<a:lstStyle/>) — they inherit from the cascade
        // and an empty re-inject is a no-op. CONSISTENCY(lstStyle-raw-passthrough):
        // mirrors customGeometryXml / effectsRaw verbatim passthrough.
        var shapeLstStyle = shape.TextBody?.GetFirstChild<Drawing.ListStyle>();
        if (shapeLstStyle != null && shapeLstStyle.HasChildren)
            node.Format["lstStyleRaw"] = shapeLstStyle.OuterXml;

        // Text alignment (from first paragraph). Only surface when explicitly
        // present in the source XML; the previous else-branch hard-coded
        // align=left whenever pPr/algn was absent, which baked an explicit
        // value into every round-trip and broke inheritance from the layout/
        // master defRPr cascade. Callers that need the effective alignment
        // can read Format["effective.align"] (cascade-resolved separately).
        var firstPara = shape.TextBody?.Elements<Drawing.Paragraph>().FirstOrDefault();
        if (firstPara?.ParagraphProperties?.Alignment?.HasValue == true)
        {
            node.Format["align"] = MapTextAlignToFriendly(firstPara.ParagraphProperties.Alignment);
        }

        // Paragraph spacing and indent (from first paragraph)
        var pProps = firstPara?.ParagraphProperties;
        if (pProps != null)
        {
            var lsPct = pProps.GetFirstChild<Drawing.LineSpacing>()?.GetFirstChild<Drawing.SpacingPercent>().PercentVal();
            if (lsPct.HasValue) node.Format["lineSpacing"] = SpacingConverter.FormatPptLineSpacingPercent(lsPct.Value);
            var lsPts = pProps.GetFirstChild<Drawing.LineSpacing>()?.GetFirstChild<Drawing.SpacingPoints>()?.Val?.Value;
            if (lsPts.HasValue) node.Format["lineSpacing"] = SpacingConverter.FormatPptLineSpacingPoints(lsPts.Value);
            var sb = pProps.GetFirstChild<Drawing.SpaceBefore>()?.GetFirstChild<Drawing.SpacingPoints>()?.Val?.Value;
            if (sb.HasValue) node.Format["spaceBefore"] = SpacingConverter.FormatPptSpacing(sb.Value);
            var sa = pProps.GetFirstChild<Drawing.SpaceAfter>()?.GetFirstChild<Drawing.SpacingPoints>()?.Val?.Value;
            if (sa.HasValue) node.Format["spaceAfter"] = SpacingConverter.FormatPptSpacing(sa.Value);
            if (pProps.Indent?.HasValue == true) node.Format["indent"] = FormatPptIndentPoints(pProps.Indent.Value);
            if (pProps.LeftMargin?.HasValue == true) node.Format["marginLeft"] = FormatPptIndentPoints(pProps.LeftMargin.Value);
            if (pProps.RightMargin?.HasValue == true) node.Format["marginRight"] = FormatPptIndentPoints(pProps.RightMargin.Value);
            // Reading direction (Arabic / Hebrew). Only emit when explicitly
            // set so LTR docs don't get a noisy `direction=ltr` everywhere.
            if (pProps.RightToLeft?.HasValue == true)
                node.Format["direction"] = pProps.RightToLeft.Value ? "rtl" : "ltr";
            EmitParagraphBreakProps(pProps, node);
        }
        // Inherit direction from slideLayout / slideMaster placeholder defaults
        // when the shape itself doesn't declare one. Surfaced as
        // `effective.direction` (mirrors the Word effective.* idiom).
        if (!node.Format.ContainsKey("direction") && part is SlidePart slidePart)
        {
            // R8-4: route the txStyles probe by placeholder type. Title
            // placeholders inherit only from titleStyle, body / subTitle from
            // bodyStyle, everything else from otherStyle. Pre-fix, the helper
            // walked txStyles.ChildElements blindly and returned the first
            // child with rtl=1 — so a master with bodyStyle rtl=1 leaked
            // direction onto a titleStyle-rtl-absent title placeholder.
            var phForDir = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
                ?.GetFirstChild<PlaceholderShape>();
            var phTypeForDir = phForDir?.Type?.HasValue == true ? phForDir.Type.Value : (PlaceholderValues?)null;
            bool? inherited = ResolveInheritedDirection(slidePart, phTypeForDir, isTitle);
            if (inherited.HasValue)
                node.Format["effective.direction"] = inherited.Value ? "rtl" : "ltr";
        }

        // Count paragraphs regardless of depth
        if (shape.TextBody != null)
        {
            var paragraphs = shape.TextBody.Elements<Drawing.Paragraph>().ToList();
            node.ChildCount = paragraphs.Count;

            // Include paragraph and run hierarchy at depth > 0
            if (depth > 0)
            {
                int paraIdx = 0;
                foreach (var para in paragraphs)
                {
                    var paraText = string.Join("", para.Elements<Drawing.Run>()
                        .Select(r => r.Text?.Text ?? ""));
                    var paraRuns = para.Elements<Drawing.Run>().ToList();

                    var paraNode = new DocumentNode
                    {
                        Path = $"{shapePath}/paragraph[{paraIdx + 1}]",
                        Type = "paragraph",
                        Text = paraText,
                        ChildCount = paraRuns.Count
                    };

                    // Add paragraph formatting info
                    var paraPProps = para.ParagraphProperties;
                    if (paraPProps?.Alignment?.HasValue == true)
                    {
                        paraNode.Format["align"] = MapTextAlignToFriendly(paraPProps.Alignment);
                    }
                    if (paraPProps?.Level?.HasValue == true) paraNode.Format["level"] = paraPProps.Level.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
                    if (paraPProps?.Indent?.HasValue == true) paraNode.Format["indent"] = FormatPptIndentPoints(paraPProps.Indent.Value);
                    if (paraPProps?.LeftMargin?.HasValue == true) paraNode.Format["marginLeft"] = FormatPptIndentPoints(paraPProps.LeftMargin.Value);
                    if (paraPProps?.RightMargin?.HasValue == true) paraNode.Format["marginRight"] = FormatPptIndentPoints(paraPProps.RightMargin.Value);
                    // R53 fuzzer-1: per-paragraph bullet (<a:buChar>, <a:buAutoNum>,
                    // <a:buNone>) was only captured on the shape's first paragraph
                    // (Format["list"] above), so dump→replay reseeded every
                    // paragraph after the first as flush-left plain text. Surface
                    // the canonical `list` value on every paragraph node so
                    // AddParagraph/Set list= re-applies the per-paragraph marker.
                    if (paraPProps != null)
                    {
                        // bulletRaw carries the full bullet group verbatim
                        // (buFont/buClr/buSzPct/buChar/…) so colored/sized/
                        // Wingdings bullets round-trip; the lossy `list` keyword
                        // is emitted only as a fallback when there's no raw block.
                        var paraBulletRaw = ReadBulletRawFromPProps(paraPProps);
                        if (paraBulletRaw != null)
                        {
                            paraNode.Format["bulletRaw"] = paraBulletRaw;
                            // R7-10: re-feedable `list` companion when canonical (suppressed
                            // for raw-char passthroughs). Mirrors R4-7 background.src companion.
                            var paraListCanon = ReadCanonicalListKeyword(paraPProps);
                            if (paraListCanon != null) paraNode.Format["list"] = paraListCanon;
                        }
                        else
                        {
                            var paraList = ReadListStyleFromPProps(paraPProps);
                            if (paraList != null) paraNode.Format["list"] = paraList;
                        }
                        // R65 bt-2: <a:tabLst>/<a:tab pos algn/> — custom tab
                        // stops were silently dropped on Get / dump, collapsing
                        // tab-aligned paragraphs after batch replay. Surface as
                        // the canonical compound `tabs=pos:algn,pos:algn,…` so
                        // AddParagraph / SetParagraph can re-stamp the list
                        // (CONSISTENCY(pptx-bare-as-points) for pos units).
                        var paraTabs = ReadTabsFromPProps(paraPProps);
                        if (paraTabs != null) paraNode.Format["tabs"] = paraTabs;
                        // Paragraph-level <a:defRPr> verbatim. Runs without their
                        // own rPr inherit size/bold/font/color from here before
                        // the layout/master bodyStyle cascade; the granular para
                        // keys never captured it, so a bare-run paragraph rebuilt
                        // at the master body size/weight instead of the authored
                        // default (e.g. 40pt-bold-Helvetica → 52pt-regular).
                        var paraDefRPr = paraPProps.GetFirstChild<Drawing.DefaultRunProperties>();
                        if (paraDefRPr != null) paraNode.Format["defRPrRaw"] = paraDefRPr.OuterXml;
                    }
                    if (paraPProps?.RightToLeft?.HasValue == true)
                        paraNode.Format["direction"] = paraPProps.RightToLeft.Value ? "rtl" : "ltr";
                    EmitParagraphBreakProps(paraPProps, paraNode);
                    var pLsPct = paraPProps?.GetFirstChild<Drawing.LineSpacing>()?.GetFirstChild<Drawing.SpacingPercent>().PercentVal();
                    if (pLsPct.HasValue) paraNode.Format["lineSpacing"] = SpacingConverter.FormatPptLineSpacingPercent(pLsPct.Value);
                    var pLsPts = paraPProps?.GetFirstChild<Drawing.LineSpacing>()?.GetFirstChild<Drawing.SpacingPoints>()?.Val?.Value;
                    if (pLsPts.HasValue) paraNode.Format["lineSpacing"] = SpacingConverter.FormatPptLineSpacingPoints(pLsPts.Value);
                    var pSb = paraPProps?.GetFirstChild<Drawing.SpaceBefore>()?.GetFirstChild<Drawing.SpacingPoints>()?.Val?.Value;
                    if (pSb.HasValue) paraNode.Format["spaceBefore"] = SpacingConverter.FormatPptSpacing(pSb.Value);
                    var pSa = paraPProps?.GetFirstChild<Drawing.SpaceAfter>()?.GetFirstChild<Drawing.SpacingPoints>()?.Val?.Value;
                    if (pSa.HasValue) paraNode.Format["spaceAfter"] = SpacingConverter.FormatPptSpacing(pSa.Value);

                    // R48: surface <a:endParaRPr sz=…> on EMPTY paragraphs
                    // (no Run children). Designers use a runless paragraph
                    // with a small endParaRPr font size as a vertical spacer
                    // between visible runs in the same shape — e.g. four
                    // cards on aionui.pptx slide 7 carry an
                    // `<a:p><a:endParaRPr sz="400"/></a:p>` (4pt spacer)
                    // between the heading and body. The dump path treats an
                    // empty paragraph as `text=""` and replay reseeds a
                    // default-size empty <a:r> which inflates the spacer to
                    // the inherited body size (~18-24pt), shifting the body
                    // copy down by 15-20pt per spacer (~30-40px total
                    // across two such paragraphs). Capture sz as
                    // `size` on the empty-paragraph node so AddParagraph
                    // can re-stamp it via the existing size= setter.
                    if (paraRuns.Count == 0)
                    {
                        var endRPr = para.GetFirstChild<Drawing.EndParagraphRunProperties>();
                        var endSz = endRPr?.FontSize?.Value;
                        if (endSz.HasValue)
                            paraNode.Format["size"] = $"{endSz.Value / 100.0:0.##}pt";
                    }

                    // Include runs at depth > 1, interleaved with <a:br>
                    // line breaks in source document order. Previously the
                    // emit walked only <a:r> elements, so an in-paragraph
                    // <a:br/> (PPT's hard line break inside a run sequence)
                    // was silently dropped on dump→replay. AddLineBreak is
                    // the replay-side counterpart (Add.Text.cs).
                    if (depth > 1)
                    {
                        int runIdx = 0;
                        int brIdx = 0;
                        foreach (var child in para.Elements())
                        {
                            if (child is Drawing.Run runChild)
                            {
                                paraNode.Children.Add(RunToNode(runChild,
                                    $"{shapePath}/paragraph[{paraIdx + 1}]/run[{runIdx + 1}]", part));
                                runIdx++;
                            }
                            else if (child is Drawing.Break brChild)
                            {
                                brIdx++;
                                var brNode = new DocumentNode
                                {
                                    Path = $"{shapePath}/paragraph[{paraIdx + 1}]/linebreak[{brIdx}]",
                                    Type = "linebreak",
                                    Text = string.Empty,
                                };
                                // <a:br><a:rPr sz=…/></a:br>: the break's run
                                // properties set the empty line's height. A bare
                                // replayed <a:br/> inherits the paragraph size
                                // instead (sample19: 14pt breaks came back 24pt,
                                // shifting everything below by 10pt per break).
                                var brRPr = brChild.GetFirstChild<Drawing.RunProperties>();
                                if (brRPr != null)
                                    brNode.Format["rPrRaw"] = brRPr.OuterXml;
                                paraNode.Children.Add(brNode);
                            }
                        }
                    }

                    // CONSISTENCY(effective-X-mirror): see WordHandler.StyleList.cs
                    // ResolveEffectiveParagraphStyleProperties — para-level keys
                    // (align / lineSpacing / spaceBefore / spaceAfter) resolved
                    // through the same 8-layer cascade as runs.
                    PopulateEffectiveParagraphPropertiesPpt(paraNode, shape, para, part);

                    node.Children.Add(paraNode);
                    paraIdx++;
                }
            }
        }

        // Animation (requires SlidePart to access Timing tree)
        if (part is SlidePart animSlidePart)
            ReadShapeAnimation(animSlidePart, shape, node);

        // Populate effective.* properties from slide layout/master inheritance
        PopulateEffectiveShapeProperties(node, shape, part);

        return node;
    }

    private static DocumentNode RunToNode(Drawing.Run run, string path, OpenXmlPart? part = null)
    {
        var node = new DocumentNode
        {
            Path = path,
            Type = "run",
            Text = run.Text?.Text ?? ""
        };

        if (run.RunProperties != null)
        {
            var fLatin = run.RunProperties.GetFirstChild<Drawing.LatinFont>()?.Typeface?.Value;
            var fEa = run.RunProperties.GetFirstChild<Drawing.EastAsianFont>()?.Typeface?.Value;
            var fCs = run.RunProperties.GetFirstChild<Drawing.ComplexScriptFont>()?.Typeface?.Value;
            // Schema: run-level `font` is write-only (get:false). Get
            // canonicalizes the readback to per-script keys
            // (font.latin / font.ea / font.cs). Emitting both bare `font`
            // and `font.latin` violates the no-duplicate-alias rule in the
            // the project conventions "Canonical DocumentNode.Format Rules".
            if (fLatin != null) node.Format["font.latin"] = fLatin;
            // CONSISTENCY(per-script-round-trip): emit font.ea / font.cs
            // unconditionally when the source <a:ea>/<a:cs> element is
            // present, even when its typeface matches latin — see the
            // shape-level emit above for rationale.
            if (fEa != null) node.Format["font.ea"] = fEa;
            if (fCs != null) node.Format["font.cs"] = fCs;
            var fs = run.RunProperties.FontSize?.Value;
            if (fs.HasValue) node.Format["size"] = $"{fs.Value / 100.0:0.##}pt";
            // R52 bt-1: surface explicit b="0"/i="0" too — designers stamp
            // `<a:rPr b="0"/>` to override an inherited bold from defRPr/lstStyle
            // (e.g. roman headlines under an italic-by-default theme run). Same
            // HasValue probe as the shape-level emit at line 1060.
            if (run.RunProperties.Bold?.HasValue == true) node.Format["bold"] = run.RunProperties.Bold.Value;
            if (run.RunProperties.Italic?.HasValue == true) node.Format["italic"] = run.RunProperties.Italic.Value;
            // CONSISTENCY(run-rtl): rPr carries an rtl attribute too; Set on a
            // run path writes here. Drawing.RunProperties doesn't expose it as a
            // typed property — read the raw attribute so Get on /paragraph/run
            // round-trips run-context rtl=true.
            foreach (var rAttr in run.RunProperties.GetAttributes())
            {
                if (rAttr.LocalName == "rtl" && !string.IsNullOrEmpty(rAttr.Value))
                {
                    node.Format["rtl"] = rAttr.Value is "1" or "true" ? "true" : "false";
                    break;
                }
            }
            if (run.RunProperties.Underline?.HasValue == true && run.RunProperties.Underline.Value != Drawing.TextUnderlineValues.None)
            {
                node.Format["underline"] = run.RunProperties.Underline.InnerText switch
                {
                    "sng" => "single",
                    "dbl" => "double",
                    _ => run.RunProperties.Underline.InnerText
                };
            }
            // CONSISTENCY(underline-color): mirror docx Get vocabulary —
            // 'underline.color' is the canonical dotted key.
            var uFill = run.RunProperties.GetFirstChild<Drawing.UnderlineFill>();
            if (uFill != null)
            {
                var uFillColor = ReadColorFromFill(uFill.GetFirstChild<Drawing.SolidFill>());
                if (uFillColor != null) node.Format["underline.color"] = uFillColor;
            }
            // R57 bt-1: <a:uLn> (Drawing.Underline) carries underline-LINE
            // styling — width (`w` EMU) plus a SolidFill child that paints
            // the stroke. Distinct from <a:uFill> above (uFill paints the
            // same stroke via a different schema slot — both are valid
            // OOXML; source decks often pick uLn when they also want a
            // custom width since uFill has no w attribute). Without reading
            // uLn, a run with <a:uLn w="38100"><a:solidFill>FF0000</a:solidFill></a:uLn>
            // round-trips as bare `underline=single` and loses both the
            // colour and the 3pt width.
            var uLn = run.RunProperties.GetFirstChild<Drawing.Underline>();
            if (uLn != null)
            {
                if (!node.Format.ContainsKey("underline.color"))
                {
                    var uLnColor = ReadColorFromFill(uLn.GetFirstChild<Drawing.SolidFill>());
                    if (uLnColor != null) node.Format["underline.color"] = uLnColor;
                }
                if (uLn.Width?.HasValue == true)
                    node.Format["underline.width"] = EmuConverter.FormatLineWidth(uLn.Width.Value);
            }
            if (run.RunProperties.Strike?.HasValue == true)
            {
                // Emit explicit "none" too — mirrors first-run reader above.
                node.Format["strike"] = run.RunProperties.Strike.Value switch
                {
                    var v when v == Drawing.TextStrikeValues.DoubleStrike => "double",
                    var v when v == Drawing.TextStrikeValues.NoStrike => "none",
                    _ => "single",
                };
            }
            if (run.RunProperties.Spacing?.HasValue == true)
                node.Format["spacing"] = $"{run.RunProperties.Spacing.Value / 100.0:0.##}";
            // R56 bt-3: emit baseline as percent (33000 → "33%"); see cell-level
            // reader above for the canonical-form rationale.
            if (run.RunProperties.Baseline?.HasValue == true && run.RunProperties.Baseline.Value != 0)
                node.Format["baseline"] = $"{run.RunProperties.Baseline.Value / 1000.0:0.##}%";
            // Color (solid or gradient)
            var runFillColor = ReadColorFromFill(run.RunProperties.GetFirstChild<Drawing.SolidFill>());
            if (runFillColor != null) node.Format["color"] = runFillColor;
            var runGrad = run.RunProperties.GetFirstChild<Drawing.GradientFill>();
            if (runGrad != null) node.Format["textFill"] = ReadGradientString(runGrad);
            // CONSISTENCY(highlight): a:highlight readback — Add/Set write it
            // via SetRunOrShapeProperties; Get must surface it for round-trip.
            var runHighlight = ReadColorFromHighlight(run.RunProperties.GetFirstChild<Drawing.Highlight>());
            if (runHighlight != null) node.Format["highlight"] = runHighlight;
            // Hyperlink (link + tooltip — round-trips Add/Set 'tooltip=…').
            if (part != null)
            {
                var runHl = run.RunProperties.GetFirstChild<Drawing.HyperlinkOnClick>();
                var linkUrl = ReadHyperlinkOnClickUrl(runHl, part);
                if (linkUrl != null) node.Format["link"] = linkUrl;
                var runTip = runHl?.Tooltip?.Value;
                if (!string.IsNullOrEmpty(runTip)) node.Format["tooltip"] = runTip!;
            }

            // Effects on the run's own <a:rPr><a:effectLst>. Set on run path
            // writes here (ApplyTextShadow / ApplyTextGlow / ApplyTextReflection
            // / ApplyTextSoftEdge); Get must read it back at run level for
            // round-trip. Format strings mirror the shape-level effect readers.
            var runEffectList = run.RunProperties.GetFirstChild<Drawing.EffectList>();
            if (runEffectList != null)
            {
                var rOuterShadow = runEffectList.GetFirstChild<Drawing.OuterShadow>();
                if (rOuterShadow != null)
                {
                    var sColor = EnsureEightDigitHexForEffect(ReadColorFromElement(rOuterShadow) ?? "000000");
                    var blurPt = rOuterShadow.BlurRadius?.HasValue == true ? $"{rOuterShadow.BlurRadius.Value / EmuConverter.EmuPerPointF:0.##}" : "4";
                    var angleDeg = rOuterShadow.Direction?.HasValue == true ? $"{rOuterShadow.Direction.Value / 60000.0:0.##}" : "45";
                    var distPt = rOuterShadow.Distance?.HasValue == true ? $"{rOuterShadow.Distance.Value / EmuConverter.EmuPerPointF:0.##}" : "3";
                    var alphaEl = rOuterShadow.Descendants<Drawing.Alpha>().FirstOrDefault();
                    var opacity = alphaEl?.Val?.HasValue == true ? $"{alphaEl.Val.Value / 1000.0:0.##}" : "100";
                    node.Format["shadow"] = $"{sColor}-{blurPt}-{angleDeg}-{distPt}-{opacity}";
                }
                var rGlow = runEffectList.GetFirstChild<Drawing.Glow>();
                if (rGlow != null)
                {
                    var gColor = EnsureEightDigitHexForEffect(ReadColorFromElement(rGlow) ?? "000000");
                    var radiusPt = rGlow.Radius?.HasValue == true ? $"{rGlow.Radius.Value / EmuConverter.EmuPerPointF:0.##}" : "8";
                    var gAlphaEl = rGlow.Descendants<Drawing.Alpha>().FirstOrDefault();
                    var gOpacity = gAlphaEl?.Val?.HasValue == true ? $"{gAlphaEl.Val.Value / 1000.0:0.##}" : "100";
                    node.Format["glow"] = $"{gColor}-{radiusPt}-{gOpacity}";
                }
                var rRefl = runEffectList.GetFirstChild<Drawing.Reflection>();
                if (rRefl != null)
                {
                    // CONSISTENCY(reflection-exact-match): mirror the shape-level
                    // reader at line 1040 — exact-preset matches emit the name,
                    // anything else emits the integer percent so Set round-trips.
                    var endPos = rRefl.EndPosition?.Value ?? 0;
                    node.Format["reflection"] = endPos switch
                    {
                        55000 => "tight",
                        90000 => "half",
                        100000 => "full",
                        _ => (endPos / 1000).ToString(System.Globalization.CultureInfo.InvariantCulture),
                    };
                }
                var rSoftEdge = runEffectList.GetFirstChild<Drawing.SoftEdge>();
                if (rSoftEdge?.Radius?.HasValue == true)
                    node.Format["softEdge"] = $"{rSoftEdge.Radius.Value / EmuConverter.EmuPerPointF:0.##}pt";
                // R62 bt-5: <a:fillOverlay blend=…><a:srgbClr/></a:fillOverlay>
                // can sit on a run's <a:rPr><a:effectLst> too (per-character
                // tinted overlay), parallel to the shape-spPr emit above. The
                // run reader walked outerShdw/glow/reflection/softEdge but
                // skipped fillOverlay, so source-authored run-level overlays
                // round-tripped to nothing — Get returned `size`+`lang` only.
                // Mirror the shape-level passthrough: surface the verbatim
                // <a:fillOverlay…/> as `fillOverlayRaw` so Set can re-install
                // it on the same run's rPr without inventing a compressed
                // `blend:color` form that would lose alpha / gradient overlay
                // variants. CONSISTENCY(effects-raw-passthrough).
                var rFillOverlay = runEffectList.GetFirstChild<Drawing.FillOverlay>();
                if (rFillOverlay != null)
                    node.Format["fillOverlayRaw"] = rFillOverlay.OuterXml;
            }

            // R61 bt-1: <a:ln> (text outline / stroke) on rPr. Distinct from
            // shape-level <a:ln> on spPr — the run-level outline strokes the
            // glyph edges of THIS run only. Reader scans rPr's first child
            // (Drawing.Outline is schema-order bucket 1 in DrawingRunPropChildOrder),
            // surfaces width as canonical pt+EMU form (FormatLineWidth, mirroring
            // shape lineWidth and run underline.width), and reads the colour out
            // of an <a:solidFill> child (matching shape line.color readback). The
            // compound `textOutline=width:color` form mirrors the connector `line=
            // color:width:dash` compound — same `:` delimiter the rest of the
            // pptx Set/Add line family uses (SplitCompoundLineValue). Without
            // this, run-level <a:ln w="6350"><a:solidFill>FF0000</a:solidFill></a:ln>
            // silently dropped on dump→batch replay.
            var runOutline = run.RunProperties.GetFirstChild<Drawing.Outline>();
            if (runOutline != null)
            {
                string? toWidth = runOutline.Width?.HasValue == true
                    ? EmuConverter.FormatLineWidth(runOutline.Width.Value) : null;
                string? toColor = ReadColorFromFill(runOutline.GetFirstChild<Drawing.SolidFill>());
                if (toWidth != null) node.Format["textOutline.width"] = toWidth;
                if (toColor != null) node.Format["textOutline.color"] = toColor;
                // Compound canonical (width:color). Surfaces even when only one
                // part is present so the bare <a:ln/> (theme-inherited stroke,
                // no attrs) still emits a marker that round-trips.
                if (toWidth != null && toColor != null)
                    node.Format["textOutline"] = $"{toWidth}:{toColor}";
                else if (toWidth != null)
                    node.Format["textOutline"] = toWidth;
                else if (toColor != null)
                    node.Format["textOutline"] = toColor;
                else
                    node.Format["textOutline"] = "true";
                // The width:color compound can't represent dash patterns,
                // gradient strokes, caps/joins etc. (sample10: a sysDot
                // WordArt outline replayed solid). Carry the verbatim <a:ln>
                // whenever it holds anything beyond width + solidFill; the
                // emitter suppresses the semantic keys in favour of the raw.
                bool outlineBeyondCompound =
                    runOutline.ChildElements.Any(c => c is not Drawing.SolidFill)
                    || runOutline.GetAttributes().Any(a => a.LocalName != "w");
                if (outlineBeyondCompound)
                    node.Format["textOutlineRaw"] = runOutline.OuterXml;
            }

            // Bitmap text fill: <a:rPr><a:blipFill> paints the glyphs with an
            // image (WordArt picture fill, sample10). No semantic key can
            // express it — carry verbatim; the emitter also re-creates the
            // referenced ImagePart with the pinned rId.
            var runBlipFill = run.RunProperties.GetFirstChild<Drawing.BlipFill>();
            if (runBlipFill != null)
                node.Format["textFillRaw"] = runBlipFill.OuterXml;

            // Long-tail OOXML fallback. drawingML rPr carries most properties
            // as attributes on rPr itself (kern, spc, lang, dirty, smtClean,
            // normalizeH, baseline, ...), with sub-elements for fills/fonts/
            // hyperlinks. Symmetric with the run-context Set fallback in
            // SetRunOrShapeProperties.
            FillUnknownRunProps(run.RunProperties, node);
        }

        // Populate effective.* properties from slide layout/master inheritance
        PopulateEffectiveRunProperties(node, run, part);

        return node;
    }

    // OOXML attribute names already mapped to canonical Format keys by the
    // curated run reader. Skip these in the long-tail fallback so we don't
    // emit `b: "1"` alongside `bold: true`, `sz: "2400"` alongside `size: "24pt"`.
    private static readonly System.Collections.Generic.HashSet<string> CuratedRunAttrs =
        new(System.StringComparer.Ordinal)
    {
        "b", "i", "u", "strike", "sz", "spc", "baseline", "rtl",
    };

    // CONSISTENCY(rpr-bool-set): mirrors DrawingRunBoolAttrs in
    // ShapeProperties.cs — these rPr attributes are OOXML xsd:boolean and
    // must read back as canonical "true"/"false" (not "1"/"0") so user
    // Add/Set vocabulary round-trips through Get.
    private static readonly System.Collections.Generic.HashSet<string> RunBoolAttrNames =
        new(System.StringComparer.Ordinal)
    {
        "b", "i", "noProof", "normalizeH", "dirty", "err", "smtClean", "kumimoji",
    };

    private static readonly System.Collections.Generic.HashSet<string> CuratedRunChildren =
        new(System.StringComparer.Ordinal)
    {
        "latin", "ea", "cs", "solidFill", "gradFill", "hlinkClick", "effectLst",
        // R61 bt-1: <a:ln> (text outline) — surfaced via the curated
        // textOutline / textOutline.color / textOutline.width keys.
        // Without this the fallback re-emits a bare <a:ln/> as `ln: true`
        // alongside the curated keys.
        "ln",
    };

    // bt-6: rPr attributes that describe the run's editor state (spell-check
    // marker, dirty/needs-recheck flag, smart-tag clean flag, BCP-47 language)
    // are meaningless when surfaced as shape-level Format keys. The shape
    // doesn't have a single language or spelling state — it just aggregates
    // first-run properties for convenience. Skip these in shape-level scans.
    // R58 bt-4: normalizeH (CJK width normalization) is a per-run typography
    // toggle. The shape-level lift used to fan it out onto the shape Format
    // bag, so a single-run shape's normalizeH replayed as a shape-level
    // `add textbox normalizeH=true` and a multi-run shape silently broadcast
    // the first run's flag to every subsequent run on dump→replay. Filter
    // it here so it stays on the run where it lives.
    private static readonly System.Collections.Generic.HashSet<string> RunOnlyAttrs =
        new(System.StringComparer.Ordinal)
    {
        "err", "dirty", "smtClean", "lang", "normalizeH",
    };

    private static void FillUnknownRunProps(Drawing.RunProperties? rPr, DocumentNode node, bool shapeLevel = false)
    {
        if (rPr == null) return;

        // Walk attributes on rPr itself.
        foreach (var attr in rPr.GetAttributes())
        {
            var name = attr.LocalName;
            if (string.IsNullOrEmpty(name)) continue;
            if (CuratedRunAttrs.Contains(name)) continue;
            if (shapeLevel && RunOnlyAttrs.Contains(name)) continue;
            if (node.Format.ContainsKey(name)) continue;
            // CONSISTENCY(rpr-bool-readback): normalize OOXML xsd:boolean
            // attrs ("1"/"0", "true"/"false") to canonical "true"/"false"
            // so Add/Set values round-trip without forcing callers to
            // memorize the wire form. Mirrors the bool set declared in
            // DrawingRunBoolAttrs (ShapeProperties.cs).
            if (RunBoolAttrNames.Contains(name))
            {
                var v = attr.Value;
                node.Format[name] = v is "1" or "true" or "True" ? "true" : "false";
            }
            else
            {
                node.Format[name] = attr.Value;
            }
        }

        // Walk leaf children that match the OOXML "child-with-val" or "toggle"
        // pattern symmetric with TryCreateTypedChild's accepted shapes.
        foreach (var child in rPr.ChildElements)
        {
            var name = child.LocalName;
            if (string.IsNullOrEmpty(name)) continue;
            if (CuratedRunChildren.Contains(name)) continue;
            if (node.Format.ContainsKey(name)) continue;
            if (child.ChildElements.Count > 0) continue;

            string? valAttr = null;
            int attrCount = 0;
            foreach (var a in child.GetAttributes())
            {
                attrCount++;
                if (a.LocalName.Equals("val", System.StringComparison.OrdinalIgnoreCase))
                    valAttr = a.Value;
            }
            if (valAttr != null) node.Format[name] = valAttr;
            else if (attrCount == 0) node.Format[name] = true;
        }
    }

    private static DocumentNode PictureToNode(Picture pic, int slideNum, int picIdx, SlidePart? slidePart = null, string? parentPathPrefix = null)
    {
        var name = pic.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value ?? "Picture";
        var alt = pic.NonVisualPictureProperties?.NonVisualDrawingProperties?.Description?.Value;

        // Detect video/audio
        var nvPr = pic.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties;
        var isVideo = nvPr?.GetFirstChild<Drawing.VideoFromFile>() != null;
        var isAudio = nvPr?.GetFirstChild<Drawing.AudioFromFile>() != null;
        var mediaType = isVideo ? "video" : isAudio ? "audio" : "picture";

        var picPathSeg = BuildElementPathSegment("picture", pic, picIdx);
        var basePath = parentPathPrefix ?? $"/slide[{slideNum}]";
        var node = new DocumentNode
        {
            Path = $"{basePath}/{picPathSeg}",
            Type = mediaType,
            Preview = name
        };

        node.Format["name"] = name;
        var picId = GetCNvPrId(pic);
        if (picId.HasValue) node.Format["id"] = picId.Value;
        var picCreationId = ReadCNvPrCreationId(pic);
        if (picCreationId != null) node.Format["extLst.creationId"] = picCreationId;
        // CONSISTENCY(media-alt-readback): emit alt for video/audio too — Set
        // accepts it and ViewAsIssues flags missing alt on audio/video <p:pic>
        // descendants, so Get must surface it to close the round-trip.
        if (!string.IsNullOrEmpty(alt)) node.Format["alt"] = alt;
        else node.Format["alt"] = "(missing)";

        // A picture rendered from mermaid (render=image) stamps `mermaid:<source>`
        // into its alt-text as the regeneration carrier. Surface the bare source as
        // a first-class Format["mermaid"] so dump/get/agents read it directly
        // instead of string-slicing the accessibility text. Storage stays in alt
        // (round-trips for free); this is a read-side convenience only.
        if (!string.IsNullOrEmpty(alt)
            && alt.StartsWith(Core.Diagram.MermaidImageRenderer.SourceTag, StringComparison.Ordinal))
            node.Format["mermaid"] = alt.Substring(Core.Diagram.MermaidImageRenderer.SourceTag.Length);

        // CONSISTENCY(picture-relid): mirror docx (WordHandler.ImageHelpers
        // emits Format["relId"] on the run-picture) and xlsx. Without this
        // key, TryExtractBinary refuses to extract the image — `get --save`
        // would fail on every pptx picture. contentType/fileSize follow the
        // same pattern as the Query picture branch so Get and Query agree.
        var embedRel = pic.BlipFill?.Blip?.Embed?.Value;
        // For video/audio nodes, BlipFill.Blip.Embed points to the poster
        // thumbnail image (PNG), not the media itself — reading its
        // contentType/fileSize would surface "image/png" + ~70 bytes for a
        // 100MB mp4. Prefer the media rel (VideoFromFile.Link /
        // AudioFromFile.Link) on the nvPr so the readback reflects the
        // actual media part. Fall back to the embed rel for plain pictures.
        string? mediaRel = null;
        if (isVideo) mediaRel = nvPr?.GetFirstChild<Drawing.VideoFromFile>()?.Link?.Value;
        else if (isAudio) mediaRel = nvPr?.GetFirstChild<Drawing.AudioFromFile>()?.Link?.Value;
        var sourceRel = mediaRel ?? embedRel;
        if (!string.IsNullOrEmpty(sourceRel))
        {
            node.Format["relId"] = sourceRel!;
            if (slidePart != null)
            {
                try
                {
                    // Media rels (video/audio) live in the
                    // DataPartReferenceRelationships collection and resolve
                    // to a MediaDataPart, not a regular OpenXmlPart — the
                    // standard GetPartById throws on them. Try the media
                    // relationship surface first; fall back to the part-by-id
                    // path for picture embeds.
                    if (mediaRel != null)
                    {
                        var dpRel = slidePart.DataPartReferenceRelationships
                            .FirstOrDefault(r => r.Id == mediaRel);
                        var dataPart = dpRel?.DataPart;
                        if (dataPart != null)
                        {
                            node.Format["contentType"] = dataPart.ContentType;
                            using var s = dataPart.GetStream();
                            node.Format["fileSize"] = s.Length;
                        }
                    }
                    else
                    {
                        var srcPart = slidePart.GetPartById(sourceRel!);
                        if (srcPart != null)
                        {
                            node.Format["contentType"] = srcPart.ContentType;
                            // Dispose the stream so the underlying ZipArchiveEntry is
                            // released — otherwise a subsequent DeletePart (e.g. when
                            // Set replaces a picture's source) throws "Cannot delete an
                            // entry currently open for writing".
                            using var s = srcPart.GetStream();
                            node.Format["fileSize"] = s.Length;
                        }
                    }
                }
                catch { /* rel may not resolve on the slide part — leave as relId-only */ }
            }
        }

        // Read media timing (volume, autoplay) from slide Timing tree
        if ((isVideo || isAudio) && slidePart != null)
        {
            var shapeId = pic.NonVisualPictureProperties?.NonVisualDrawingProperties?.Id?.Value;
            if (shapeId != null)
                ReadMediaTimingProperties(slidePart, shapeId.Value, node);

            // p14:trim
            var p14Media = nvPr?.Descendants<DocumentFormat.OpenXml.Office2010.PowerPoint.Media>().FirstOrDefault();
            var trim = p14Media?.MediaTrim;
            if (trim != null)
            {
                if (trim.Start?.Value != null) node.Format["trimStart"] = trim.Start.Value;
                if (trim.End?.Value != null) node.Format["trimEnd"] = trim.End.Value;
            }
        }

        // Position and size
        var picXfrm = pic.ShapeProperties?.Transform2D;
        if (picXfrm?.Offset != null)
        {
            if (picXfrm.Offset.X is not null) node.Format["x"] = FormatEmu(picXfrm.Offset.X!);
            if (picXfrm.Offset.Y is not null) node.Format["y"] = FormatEmu(picXfrm.Offset.Y!);
        }
        if (picXfrm?.Extents != null)
        {
            if (picXfrm.Extents.Cx is not null) node.Format["width"] = FormatEmu(picXfrm.Extents.Cx!);
            if (picXfrm.Extents.Cy is not null) node.Format["height"] = FormatEmu(picXfrm.Extents.Cy!);
        }
        if (picXfrm?.Rotation != null && picXfrm.Rotation.Value != 0)
            node.Format["rotation"] = $"{picXfrm.Rotation.Value / 60000.0:0.######}";

        // Flip — CONSISTENCY(shape-picture-parity): mirror ShapeToNode.
        if (picXfrm?.HorizontalFlip?.Value == true) node.Format["flipH"] = true;
        if (picXfrm?.VerticalFlip?.Value == true) node.Format["flipV"] = true;

        // CONSISTENCY(picture-geometry): a picture can be "cropped to shape" via a
        // non-rectangle <a:prstGeom> on its spPr (e.g. prst="ellipse"). AddPicture
        // stamps a default rect, so without surfacing the preset the crop-to-shape
        // is lost on dump→replay — the picture renders as a plain rectangle. Emit
        // the preset name; AddPicture accepts it via the `geometry`/`shape` prop.
        // Skip "rect" — it is AddPicture's default, so emitting it on every plain
        // picture would only add noise.
        var picPresetName = pic.ShapeProperties?
            .GetFirstChild<Drawing.PresetGeometry>()?.Preset?.InnerText;
        if (!string.IsNullOrEmpty(picPresetName) && picPresetName != "rect")
        {
            node.Format["geometry"] = picPresetName;
            // Preset adjust handles on the crop shape — mirror ShapeToNode's
            // adj readback (round2DiagRect corner radius etc.); dropping them
            // reverts the crop to default proportions (custom-shape-bitmap-fill).
            var picAvLst = pic.ShapeProperties?
                .GetFirstChild<Drawing.PresetGeometry>()?.GetFirstChild<Drawing.AdjustValueList>();
            if (picAvLst != null)
            {
                var picGuides = picAvLst.Elements<Drawing.ShapeGuide>()
                    .Where(g => g.Name?.HasValue == true && g.Formula?.HasValue == true)
                    .Select(g => $"{g.Name!.Value}:{g.Formula!.Value}")
                    .ToList();
                if (picGuides.Count > 0)
                    node.Format["adj"] = string.Join(",", picGuides);
            }
        }

        // Crop-to-CUSTOM-shape: a picture cropped to a freeform carries
        // <a:custGeom> instead of prstGeom (sample11). Surface the verbatim
        // XML — AddPicture consumes customGeometryXml the same way AddShape
        // does, so the crop path replays byte-faithfully instead of falling
        // back to the default rect.
        var picCustGeom = pic.ShapeProperties?.GetFirstChild<Drawing.CustomGeometry>();
        if (picCustGeom != null)
            node.Format["customGeometryXml"] = picCustGeom.OuterXml;

        // Shape fill ON the picture's spPr — visible wherever the image
        // doesn't cover the frame (e.g. a negative srcRect outset leaves a
        // border that the spPr fill paints; sample17's black surround).
        // `fill` on picture Add is taken by the fit-mode alias, so use a
        // dedicated key.
        var picSpFill = ReadColorFromFill(pic.ShapeProperties?.GetFirstChild<Drawing.SolidFill>());
        if (picSpFill != null)
            node.Format["frameFill"] = picSpFill;

        // 3D on a picture: <a:scene3d>/<a:sp3d> on the pic's spPr (camera
        // rotation + extrusion/bevel). ShapeToNode reads these semantically
        // for shapes, but pictures had no readback at all — a 3D-rotated
        // picture replayed flat (Scene3d_pureImage). Verbatim carriers,
        // spliced back by AddPicture.
        var picScene3d = pic.ShapeProperties?.GetFirstChild<Drawing.Scene3DType>();
        if (picScene3d != null)
            node.Format["scene3dRaw"] = picScene3d.OuterXml;
        var picSp3d = pic.ShapeProperties?.GetFirstChild<Drawing.Shape3DType>();
        if (picSp3d != null)
            node.Format["sp3dRaw"] = picSp3d.OuterXml;

        // CONSISTENCY(zorder): mirror shape/connector — emit for any
        // ShapeTree-rooted picture so Add(picture, zorder=N) round-trips.
        if (pic.Parent is ShapeTree picZTree)
        {
            var picZContent = picZTree.ChildElements
                .Where(e => e is Shape or Picture or GraphicFrame or GroupShape or ConnectionShape)
                .ToList();
            var picZIdx = picZContent.IndexOf(pic);
            if (picZIdx >= 0) node.Format["zorder"] = picZIdx + 1;
        }

        // Opacity (via AlphaModulateFixedEffect on blip)
        var picBlip = pic.BlipFill?.GetFirstChild<Drawing.Blip>();
        var alphaModFix = picBlip?.GetFirstChild<Drawing.AlphaModulationFixed>();
        if (alphaModFix?.Amount?.HasValue == true)
            node.Format["opacity"] = $"{alphaModFix.Amount.Value / 100000.0:0.##}";

        // R57 bt-3: surface <a:blip cstate="email|print|hqprint|screen|none">
        // as the `compressionState` Format key. cstate is an ATTRIBUTE on the
        // blip element itself (not a child), so the R56 blip-children raw-set
        // passthrough does NOT cover it — without this readback the
        // PowerPoint Picture Format → Compress Pictures choice was silently
        // dropped on dump→replay and the image came back at the source's
        // default compression. Skip `none` since that's the schema default.
        if (picBlip?.CompressionState?.HasValue == true)
        {
            var cstate = picBlip.CompressionState.InnerText;
            if (!string.IsNullOrEmpty(cstate) && cstate != "none")
                node.Format["compressionState"] = cstate;
        }

        // bt-2: surface <a:biLevel thresh="N"/> as the `biLevel` Format key.
        // BiLevel converts the picture to 1-bit black/white using `thresh`
        // as the luminance cutoff (0..100000 permille). Without this readback
        // the dump path silently dropped the effect on round-trip — the
        // image came back in full color and the high-contrast art look
        // (mono outline / printable mode) collapsed.
        var picBiLevel = picBlip?.GetFirstChild<Drawing.BiLevel>();
        if (picBiLevel?.Threshold?.HasValue == true)
            node.Format["biLevel"] = $"{picBiLevel.Threshold.Value / 1000.0:0.##}";

        // R52 bt-1: surface <a:duotone> blip recolor as `duotone=#c1,#c2`.
        // Duotone maps the image's luminance gradient between two stops,
        // producing the printed-mono / brand-tint look. Two srgbClr children
        // are required by the schema; scheme colors are also valid stops.
        // Without this readback the duotone block was silently dropped on
        // dump→replay and the image came back full-color.
        var picDuotone = picBlip?.GetFirstChild<Drawing.Duotone>();
        if (picDuotone != null)
        {
            var duoStops = new List<string>();
            foreach (var stop in picDuotone.ChildElements)
            {
                if (stop is Drawing.RgbColorModelHex rgbStop && rgbStop.Val?.Value is { } rgbV)
                    duoStops.Add(ParseHelpers.FormatHexColor(rgbV));
                else if (stop is Drawing.SchemeColor schemeStop && schemeStop.Val?.HasValue == true
                         && !string.IsNullOrEmpty(schemeStop.Val.InnerText))
                    duoStops.Add(schemeStop.Val.InnerText);
            }
            if (duoStops.Count == 2)
                node.Format["duotone"] = string.Join(",", duoStops);
        }

        // R31 bt: surface <a:grayscl/> blip recolor as `recolor=grayscale`.
        // Set.Media writes a bare <a:grayscl/> from `recolor=grayscale`, so the
        // readback key/value must mirror that to round-trip. Without this the
        // grayscale recolor (Picture Format → Color → Recolor → Grayscale) was
        // silently dropped on Get / dump→replay and the image came back full-color.
        if (picBlip?.GetFirstChild<Drawing.Grayscale>() != null)
            node.Format["recolor"] = "grayscale";

        // Click-hyperlink on the picture (nvPicPr/cNvPr/a:hlinkClick).
        // CONSISTENCY(shape-picture-parity): pictures share the cNvPr
        // hyperlink slot with shapes; reuse the same reader.
        if (slidePart != null)
        {
            var picHl = pic.NonVisualPictureProperties?.NonVisualDrawingProperties?
                .GetFirstChild<Drawing.HyperlinkOnClick>();
            var picLinkUrl = ReadHyperlinkOnClickUrl(picHl, slidePart);
            if (picLinkUrl != null) node.Format["link"] = picLinkUrl;
            var picTip = picHl?.Tooltip?.Value;
            if (!string.IsNullOrEmpty(picTip)) node.Format["tooltip"] = picTip!;
        }

        // Brightness / contrast — stored on the blip as <a:lum bright="N"
        // contrast="M"/> (CT_LuminanceEffect; each value is percent × 1000).
        // Mirrors the write side in Set.Media.cs. Legacy files written by
        // older builds may carry an invalid <a:lumMod>/<a:lumOff> pair under
        // the blip; we still read them so existing decks display correctly
        // until they're re-Set (which migrates them to <a:lum>).
        if (picBlip != null)
        {
            int? brightVal = null, contrastVal = null;
            foreach (var kid in picBlip.ChildElements)
            {
                if (kid.NamespaceUri != "http://schemas.openxmlformats.org/drawingml/2006/main") continue;
                if (kid is Drawing.LuminanceEffect lumElem)
                {
                    if (lumElem.Brightness?.HasValue == true) brightVal = lumElem.Brightness.Value;
                    if (lumElem.Contrast?.HasValue == true) contrastVal = lumElem.Contrast.Value;
                }
                else if (kid.LocalName == "lumOff" || kid.LocalName == "lumMod")
                {
                    // Legacy invalid markup written by older builds.
                    var valAttr = kid.GetAttribute("val", "").Value;
                    if (string.IsNullOrEmpty(valAttr) || !int.TryParse(valAttr, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out var iv)) continue;
                    if (kid.LocalName == "lumOff") brightVal ??= iv;
                    else if (kid.LocalName == "lumMod") contrastVal ??= iv - 100000;
                }
            }
            if (brightVal.HasValue && brightVal.Value != 0)
                node.Format["brightness"] = $"{brightVal.Value / 1000.0:0.##}";
            if (contrastVal.HasValue && contrastVal.Value != 0)
                node.Format["contrast"] = $"{contrastVal.Value / 1000.0:0.##}";
        }

        // Shadow / glow — Set.Media writes these into spPr/effectLst via
        // shared ApplyShadow/ApplyGlow. Mirror the shape-level reader so
        // picture round-trips match shapes.
        var picEffectList = pic.ShapeProperties?.GetFirstChild<Drawing.EffectList>();
        if (picEffectList != null)
        {
            var picOuterShadow = picEffectList.GetFirstChild<Drawing.OuterShadow>();
            if (picOuterShadow != null)
            {
                var shadowColor = EnsureEightDigitHexForEffect(ReadColorFromElement(picOuterShadow) ?? "000000");
                var blurPt = picOuterShadow.BlurRadius?.HasValue == true ? $"{picOuterShadow.BlurRadius.Value / EmuConverter.EmuPerPointF:0.##}" : "4";
                var angleDeg = picOuterShadow.Direction?.HasValue == true ? $"{picOuterShadow.Direction.Value / 60000.0:0.##}" : "45";
                var distPt = picOuterShadow.Distance?.HasValue == true ? $"{picOuterShadow.Distance.Value / EmuConverter.EmuPerPointF:0.##}" : "3";
                var alphaEl = picOuterShadow.Descendants<Drawing.Alpha>().FirstOrDefault();
                var opacity = alphaEl?.Val?.HasValue == true ? $"{alphaEl.Val.Value / 1000.0:0.##}" : "40";
                node.Format["shadow"] = $"{shadowColor}-{blurPt}-{angleDeg}-{distPt}-{opacity}";
            }
            var picGlow = picEffectList.GetFirstChild<Drawing.Glow>();
            if (picGlow != null)
            {
                var glowColor = EnsureEightDigitHexForEffect(ReadColorFromElement(picGlow) ?? "000000");
                var radiusPt = picGlow.Radius?.HasValue == true ? $"{picGlow.Radius.Value / EmuConverter.EmuPerPointF:0.##}" : "8";
                var glowAlpha = picGlow.Descendants<Drawing.Alpha>().FirstOrDefault();
                var glowOpacity = glowAlpha?.Val?.HasValue == true ? $"{glowAlpha.Val.Value / 1000.0:0.##}" : "75";
                node.Format["glow"] = $"{glowColor}-{radiusPt}-{glowOpacity}";
            }
            // Semantic shadow= backfills defaults for absent dist/dir (a
            // source outerShdw with only blurRad+algn replays with dist=3pt
            // dir=45° added). Carry the verbatim effectLst so replay is
            // byte-faithful; the emitter suppresses the semantic keys.
            node.Format["effectsRaw"] = picEffectList.OuterXml;
        }

        // Picture border: <a:ln> on the pic's spPr (the white frame around a
        // crop-to-shape picture, custom-shape-bitmap-fill). No readback
        // existed — carry verbatim.
        var picLn = pic.ShapeProperties?.GetFirstChild<Drawing.Outline>();
        if (picLn != null)
            node.Format["lineRaw"] = picLn.OuterXml;

        // Crop
        var srcRect = pic.BlipFill?.GetFirstChild<Drawing.SourceRectangle>();
        if (srcRect != null)
        {
            var cl = srcRect.Left?.Value ?? 0;
            var ct = srcRect.Top?.Value ?? 0;
            var cr = srcRect.Right?.Value ?? 0;
            var cb = srcRect.Bottom?.Value ?? 0;
            if (cl != 0 || ct != 0 || cr != 0 || cb != 0)
                // OOXML stores crop edges in perMille (thousandths of a percent),
                // so `12345` = 12.345% and every fractional digit through the
                // third is significant. The Set side accepts `cropLeft=12.345`
                // and stores `<a:srcRect l="12345"/>`; the readback must round
                // to 3 decimals so dump → batch → replay is byte-stable.
                // Pre-R12 the formatter was `0.##` (2 decimals), so `12.345`
                // round-tripped as `12.35` — silent precision loss.
                node.Format["crop"] = $"{cl / 1000.0:0.###},{ct / 1000.0:0.###},{cr / 1000.0:0.###},{cb / 1000.0:0.###}";
        }

        // bt-2 / bt-5: blipFill mode — <a:tile .../> vs
        // <a:stretch><a:fillRect .../>. PictureToNode previously surfaced
        // only srcRect (crop). Without these readbacks <a:tile> collapsed
        // into a default <a:stretch>, and any <a:fillRect> insets (positive
        // for letterbox padding, negative for outset cover) were zeroed.
        // R47 944a9c7d covered the srcRect negative-crop range; this is the
        // tile / fillRect axis.
        var picTile = pic.BlipFill?.GetFirstChild<Drawing.Tile>();
        var picStretch = pic.BlipFill?.GetFirstChild<Drawing.Stretch>();
        if (picTile != null)
        {
            node.Format["fillMode"] = "tile";
            // Sx/Sy default to 100000 (100% — same scale as source). Emit only
            // when present; AddPicture's tilescale single-axis fallback uses
            // 1.0 absent input.
            if (picTile.HorizontalRatio?.HasValue == true)
                node.Format["tileSx"] = picTile.HorizontalRatio.Value;
            if (picTile.VerticalRatio?.HasValue == true)
                node.Format["tileSy"] = picTile.VerticalRatio.Value;
            if (picTile.HorizontalOffset?.HasValue == true)
                node.Format["tileTx"] = picTile.HorizontalOffset.Value;
            if (picTile.VerticalOffset?.HasValue == true)
                node.Format["tileTy"] = picTile.VerticalOffset.Value;
            if (picTile.Alignment?.HasValue == true)
                node.Format["tileAlgn"] = picTile.Alignment.InnerText;
            if (picTile.Flip?.HasValue == true)
                node.Format["tileFlip"] = picTile.Flip.InnerText;
        }
        else if (picStretch != null)
        {
            // bt-5: positive insets shrink the stretched area (letterbox);
            // negative insets oversize it (cover-style crop, already
            // honored by the srcRect path R47 fixed). Either sign rides
            // through fillRect verbatim — the AddPicture stretch path used
            // to always write a child-less <a:fillRect/>, losing every
            // authored inset on round-trip.
            var fr = picStretch.GetFirstChild<Drawing.FillRectangle>();
            if (fr != null)
            {
                var fl = fr.Left?.Value;
                var ft = fr.Top?.Value;
                var frVal = fr.Right?.Value;
                var fb = fr.Bottom?.Value;
                if (fl.HasValue || ft.HasValue || frVal.HasValue || fb.HasValue)
                    node.Format["fillRect"] = $"{fl ?? 0},{ft ?? 0},{frVal ?? 0},{fb ?? 0}";
            }
            else
            {
                // Bare <a:stretch/> — real PowerPoint renders a negative
                // srcRect differently with vs without an explicit
                // <a:fillRect/> (sample17). Flag it so AddPicture writes
                // the same bare form back.
                node.Format["stretchBare"] = "true";
            }
        }

        return node;
    }

    /// <summary>
    /// Read volume and autoplay from the slide timing tree for a media shape.
    /// </summary>
    private static void ReadMediaTimingProperties(SlidePart slidePart, uint shapeId, DocumentNode node)
    {
        var timing = slidePart.Slide?.GetFirstChild<Timing>();
        if (timing == null) return;

        var shapeIdStr = shapeId.ToString();

        // Read volume from p:video/p:audio → cMediaNode
        foreach (var mediaNode in timing.Descendants<CommonMediaNode>())
        {
            var target = mediaNode.TargetElement?.GetFirstChild<ShapeTarget>();
            if (target?.ShapeId?.Value != shapeIdStr) continue;

            if (mediaNode.Volume?.HasValue == true)
                node.Format["volume"] = (int)(mediaNode.Volume.Value / 1000.0);
            // Loop-until-Stopped: cMediaNode's cTn has
            // repeatCount="indefinite" when looped.
            var loopCTn = mediaNode.CommonTimeNode;
            if (loopCTn?.RepeatCount?.Value is string rc
                && rc.Equals("indefinite", StringComparison.OrdinalIgnoreCase))
                node.Format["loop"] = true;
            break;
        }

        // Read autoplay from main sequence: look for cmd="playFrom(0)" targeting this shape
        // with nodeType="afterEffect" (autoplay) vs "clickEffect" (click-to-play)
        foreach (var cmd in timing.Descendants<Command>())
        {
            if (cmd.CommandName?.Value != "playFrom(0)") continue;
            var cmdTarget = cmd.CommonBehavior?.TargetElement?.GetFirstChild<ShapeTarget>();
            if (cmdTarget?.ShapeId?.Value != shapeIdStr) continue;

            // Found the playback command — check its parent cTn for nodeType
            var parentCTn = cmd.Parent as CommonTimeNode
                ?? cmd.Ancestors<CommonTimeNode>().FirstOrDefault();
            if (parentCTn?.NodeType?.Value == TimeNodeValues.AfterEffect)
                node.Format["autoPlay"] = true;
            break;
        }
    }

    private static Shape CreateTextShape(uint id, string name, string text, bool isTitle, bool isTextBox = false,
                                         PlaceholderValues? placeholderType = null, uint? placeholderIndex = null)
    {
        var shape = new Shape();
        var appNvPr = new ApplicationNonVisualDrawingProperties();
        if (isTitle)
        {
            appNvPr.AppendChild(new PlaceholderShape { Type = PlaceholderValues.Title });
        }
        else if (placeholderType.HasValue)
        {
            var ph = new PlaceholderShape { Type = placeholderType.Value };
            if (placeholderIndex.HasValue) ph.Index = placeholderIndex.Value;
            appNvPr.AppendChild(ph);
        }
        // OOXML `<p:cNvSpPr txBox="1"/>` is the only on-disk marker that
        // distinguishes a dedicated text container from a geometry shape that
        // happens to carry text. Without it, every shape with a prstGeom and
        // empty/short text is indistinguishable on readback.
        var cNvSpPr = new NonVisualShapeDrawingProperties();
        if (isTextBox)
            cNvSpPr.TextBox = true;
        shape.NonVisualShapeProperties = new NonVisualShapeProperties(
            new NonVisualDrawingProperties { Id = id, Name = name },
            cNvSpPr,
            appNvPr
        );
        var spPr = new ShapeProperties();
        if (isTitle)
        {
            // Default title position: top-center area of standard 16:9 slide.
            // EMU values picked to round-trip exactly through FormatEmu in pt
            // (12700 EMU/pt) so Get readback is unit-qualified, not raw emu.
            spPr.Transform2D = new Drawing.Transform2D
            {
                Offset = new Drawing.Offset { X = 838200, Y = 365125 },    // 66pt, 28.75pt
                Extents = new Drawing.Extents { Cx = 10515600, Cy = 1320800 } // 828pt, 104pt
            };
        }
        else
        {
            // Default body/content position: below title.
            // Cx=10515600=828pt, Cy=4349750=342.5pt, Y=1825625=143.75pt — all
            // pt-exact so FormatEmu emits clean readback (was Cy=4351338 →
            // 342.625pt, which fell back to raw emu).
            spPr.Transform2D = new Drawing.Transform2D
            {
                Offset = new Drawing.Offset { X = 838200, Y = 1825625 },   // 66pt, 143.75pt
                Extents = new Drawing.Extents { Cx = 10515600, Cy = 4349750 } // 828pt, 342.5pt
            };
        }
        shape.ShapeProperties = spPr;
        var body = new TextBody(
            new Drawing.BodyProperties(),
            new Drawing.ListStyle()
        );
        // CONSISTENCY(text-escape-boundary): \n / \t resolution at CLI --prop;
        // text arrives here with real newlines and tabs already.
        if (string.IsNullOrEmpty(text))
        {
            // Decorator shapes (no text) must not seed a default <a:r> with
            // lang="en-US" — that lang attribute leaks back through
            // FillUnknownRunProps to shape-level Format on round-trip, so a
            // source <p:sp> with no rPr lang gains lang=en-US after Add→Get.
            // Mirror what PowerPoint emits for an empty text body: a single
            // empty paragraph with no run, no endParaRPr lang. (DRIFT-3)
            body.AppendChild(new Drawing.Paragraph());
        }
        else
        {
            var lines = text.Split('\n');
            foreach (var line in lines)
            {
                var para = new Drawing.Paragraph();
                AppendLineWithTabs(para, line, seg => new Drawing.Run(
                    new Drawing.RunProperties { Language = "en-US" },
                    new Drawing.Text { Text = seg }
                ));
                body.AppendChild(para);
            }
        }
        shape.TextBody = body;
        return shape;
    }

    private static DocumentNode ConnectorToNode(ConnectionShape cxn, int slideNum, int cxnIdx, string? parentPathPrefix = null, int depth = 0, OpenXmlPart? part = null)
    {
        var name = cxn.NonVisualConnectionShapeProperties?.NonVisualDrawingProperties?.Name?.Value ?? "Connector";
        var cxnPathSeg = BuildElementPathSegment("connector", cxn, cxnIdx);
        var basePath = parentPathPrefix ?? $"/slide[{slideNum}]";
        var node = new DocumentNode
        {
            Path = $"{basePath}/{cxnPathSeg}",
            Type = "connector",
            Preview = name
        };
        node.Format["name"] = name;
        var cxnId = GetCNvPrId(cxn);
        if (cxnId.HasValue) node.Format["id"] = cxnId.Value;
        var cxnCreationId = ReadCNvPrCreationId(cxn);
        if (cxnCreationId != null) node.Format["extLst.creationId"] = cxnCreationId;

        var spPr = cxn.ShapeProperties;
        var xfrm = spPr?.GetFirstChild<Drawing.Transform2D>();
        if (xfrm != null)
        {
            if (xfrm.Offset?.X != null) node.Format["x"] = FormatEmu(xfrm.Offset.X!);
            if (xfrm.Offset?.Y != null) node.Format["y"] = FormatEmu(xfrm.Offset.Y!);
            if (xfrm.Extents?.Cx != null) node.Format["width"] = FormatEmu(xfrm.Extents.Cx!);
            if (xfrm.Extents?.Cy != null) node.Format["height"] = FormatEmu(xfrm.Extents.Cy!);
        }

        // Fill (solid fill on the connector shape itself, not on the outline)
        var cxnFill = ReadColorFromFill(spPr?.GetFirstChild<Drawing.SolidFill>());
        if (cxnFill != null) node.Format["fill"] = cxnFill;
        if (spPr?.GetFirstChild<Drawing.NoFill>() != null) node.Format["fill"] = "none";

        var geom = spPr?.GetFirstChild<Drawing.PresetGeometry>();
        if (geom?.Preset?.HasValue == true)
            // CONSISTENCY(canonical-key): canonical 'shape'; 'preset' was legacy key.
            node.Format["shape"] = geom.Preset.InnerText;

        var ln = spPr?.GetFirstChild<Drawing.Outline>();
        var lnIsNone = ln?.GetFirstChild<Drawing.NoFill>() != null;
        if (!lnIsNone && ln?.Width?.HasValue == true)
            node.Format["lineWidth"] = FormatLineWidth(ln.Width.Value);
        var cxnDash = ln?.GetFirstChild<Drawing.PresetDash>();
        if (cxnDash?.Val?.HasValue == true)
        {
            // emit canonical OOXML token (see shape readback).
            var dashValue = cxnDash.Val.InnerText ?? "";
            node.Format["lineDash"] = dashValue switch
            {
                "solid" => "solid",
                "dot" => "dot",
                "dash" => "dash",
                "dashDot" => "dashDot",
                "lgDash" => "lgDash",
                "lgDashDot" => "lgDashDot",
                "lgDashDotDot" => "lgDashDotDot",
                "sysDot" => "sysDot",
                "sysDash" => "sysDash",
                "sysDashDot" => "sysDashDot",
                "sysDashDotDot" => "sysDashDotDot",
                _ => dashValue
            };
        }
        // R64 bt-3: <a:custDash><a:ds d="N" sp="N"/>...</a:custDash> — author-
        // defined dash pattern as alternating dash-length / space-length stops
        // (1000ths of em). prstDash and custDash are mutually exclusive in
        // CT_LineProperties (EG_LineDashProperties choice). No compressible
        // string form — mirror shadowRaw / fillOverlayRaw and surface the
        // OuterXml as lineDashRaw so dump→batch replay rebuilds the custom
        // pattern verbatim instead of dropping the dash and falling back to
        // a solid stroke.
        var cxnCustDash = ln?.GetFirstChild<Drawing.CustomDash>();
        if (cxnCustDash != null)
        {
            node.Format["lineDashRaw"] = cxnCustDash.OuterXml;
        }
        // R58 bt-3: <a:ln cap="..." cmpd="..."> attributes — connector readback
        // was previously dropping cap (lineCap) and cmpd (compound line) on
        // round-trip even though the shape outline branch above already
        // surfaces both. Mirror the shape mapping (rnd→round, sq→square) so
        // canonical Format keys match across shape/connector outlines, and
        // dump→batch replay rebuilds the double-line / round-cap stroke.
        if (ln?.CapType?.HasValue == true)
        {
            var cxnCapRaw = ln.CapType.InnerText ?? "";
            node.Format["lineCap"] = cxnCapRaw switch
            {
                "rnd" => "round",
                "sq" => "square",
                "flat" => "flat",
                _ => cxnCapRaw
            };
        }
        if (ln?.CompoundLineType?.HasValue == true)
            node.Format["cmpd"] = ln.CompoundLineType.InnerText ?? "";

        // R61 bt-2: <a:ln> line-join children (<a:round/>, <a:bevel/>, <a:miter lim="N"/>)
        // — connector readback was previously dropping all three on round-trip even
        // though the shape outline branch above (~line 1331) already surfaces them.
        // Mirror shape mapping (round/bevel/miter) and surface the miter limit as
        // miterLimit (OOXML lim attribute: 1000ths of a percent, e.g. 800000 = 800%).
        // Without this, dump→batch replay drops <a:miter lim="800000"/> entirely.
        if (ln?.GetFirstChild<Drawing.Round>() != null)
            node.Format["lineJoin"] = "round";
        else if (ln?.GetFirstChild<Drawing.LineJoinBevel>() != null)
            node.Format["lineJoin"] = "bevel";
        else
        {
            var miterEl = ln?.GetFirstChild<Drawing.Miter>();
            if (miterEl != null)
            {
                node.Format["lineJoin"] = "miter";
                if (miterEl.Limit?.HasValue == true)
                    node.Format["miterLimit"] = miterEl.Limit.Value;
            }
        }

        // Gradient on the connector line — emit round-trippable spec so dump→batch
        // replay rebuilds the gradient instead of falling back to a bare <a:ln/>
        // (which would inherit the theme's default thin stroke). Mirrors the shape
        // outline gradient readback above.
        var cxnLineGradFill = ln?.GetFirstChild<Drawing.GradientFill>();
        if (cxnLineGradFill != null)
        {
            node.Format["line.gradient"] = ReadGradientString(cxnLineGradFill);
        }
        var solidFill = ln?.GetFirstChild<Drawing.SolidFill>();
        var rgb = solidFill?.GetFirstChild<Drawing.RgbColorModelHex>();
        // CONSISTENCY(canonical-key): canonical 'color'; 'lineColor' was legacy.
        // Use ReadColorFromFill so scheme-color line= (accent1, dark1, …) round-trips
        // through Get; the prior rgb-only branch silently dropped a:schemeClr.
        var cxnColor = ReadColorFromFill(solidFill);
        if (cxnColor != null) node.Format["color"] = cxnColor;

        // Line opacity
        var cxnColorEl = rgb as OpenXmlElement ?? solidFill?.GetFirstChild<Drawing.SchemeColor>();
        var cxnAlpha = cxnColorEl?.GetFirstChild<Drawing.Alpha>()?.Val?.Value;
        if (cxnAlpha.HasValue) node.Format["lineOpacity"] = $"{cxnAlpha.Value / 100000.0:0.##}";

        // Head/tail end arrows
        var headEnd = ln?.GetFirstChild<Drawing.HeadEnd>();
        if (headEnd?.Type?.HasValue == true)
            node.Format["headEnd"] = headEnd.Type.InnerText;
        var tailEnd = ln?.GetFirstChild<Drawing.TailEnd>();
        if (tailEnd?.Type?.HasValue == true)
            node.Format["tailEnd"] = tailEnd.Type.InnerText;

        // Rotation
        if (xfrm?.Rotation?.HasValue == true && xfrm.Rotation.Value != 0)
            node.Format["rotation"] = $"{xfrm.Rotation.Value / 60000.0:0.######}";

        // Flip — a bent/elbow connector's actual routing depends on flipH/flipV
        // (they mirror the path within its bounding box). AddConnector already
        // honors both; ConnectorToNode previously read only rotation, so a
        // flipped connector round-tripped with mirrored routing (wrong bends).
        if (xfrm?.HorizontalFlip?.Value == true) node.Format["flipH"] = true;
        if (xfrm?.VerticalFlip?.Value == true) node.Format["flipV"] = true;

        // Z-order (1-based position among content elements: 1 = back, N = front).
        // CONSISTENCY(zorder): shape/picture/group all emit zorder when parent is a
        // ShapeTree; connector belongs to the same set and was previously omitted —
        // round-trip of Add(connector, zorder=N) silently dropped the value.
        if (cxn.Parent is ShapeTree cxnTree)
        {
            var contentEls = cxnTree.ChildElements
                .Where(e => e is Shape or Picture or GraphicFrame or GroupShape or ConnectionShape)
                .ToList();
            var cxnZIdx = contentEls.IndexOf(cxn);
            if (cxnZIdx >= 0) node.Format["zorder"] = cxnZIdx + 1;
        }

        // Connection info (startShape/endShape)
        var cxnDrawProps = cxn.NonVisualConnectionShapeProperties?.NonVisualConnectorShapeDrawingProperties;
        var startCxn = cxnDrawProps?.StartConnection;
        if (startCxn?.Id?.HasValue == true)
        {
            node.Format["startShape"] = startCxn.Id.Value;
            if (startCxn.Index?.HasValue == true)
                node.Format["startIdx"] = startCxn.Index.Value;
        }
        var endCxn = cxnDrawProps?.EndConnection;
        if (endCxn?.Id?.HasValue == true)
        {
            node.Format["endShape"] = endCxn.Id.Value;
            if (endCxn.Index?.HasValue == true)
                node.Format["endIdx"] = endCxn.Index.Value;
        }

        // R53 bt-2: surface <a:cxnSpLocks noChangeShapeType="1"> — the
        // shape-type lock PowerPoint emits to pin a connector primitive in
        // place. Reader-only support of the most common bit; AddConnector
        // honors `lockShapeType=true` to re-emit on dump→replay.
        var cxnLocks = cxnDrawProps?.GetFirstChild<Drawing.ConnectionShapeLocks>();
        if (cxnLocks?.NoChangeShapeType?.Value == true)
            node.Format["lockShapeType"] = "true";

        // R57 bt-4: surface inline text label on <p:cxnSp> — PowerPoint
        // (and most flowchart authoring tools) attach a <p:txBody> child to
        // connectors to render an in-line label between the two endpoints.
        // The OOXML p:cxnSp schema does NOT declare <p:txBody> as a typed
        // child (only nvCxnSpPr / spPr / style / extLst), so the OpenXml SDK
        // surfaces the entire <p:txBody> subtree as an OpenXmlUnknownElement
        // tree. Reparse its OuterXml into a strongly-typed Presentation.TextBody
        // so the existing paragraph/run walker can extract text + chars props.
        // Without this, Get/dump silently drop every connector label and
        // round-trip wipes the text entirely.
        var cxnTextBody = ResolveConnectorTextBody(cxn);
        if (cxnTextBody != null)
        {
            var cxnParas = cxnTextBody.Elements<Drawing.Paragraph>().ToList();
            // Surface aggregate text so `view text` / single-line consumers
            // see the label without descending into children.
            var aggText = string.Join("\n", cxnParas
                .Select(p => string.Join("", p.Elements<Drawing.Run>().Select(r => r.Text?.Text ?? ""))));
            if (!string.IsNullOrEmpty(aggText)) node.Text = aggText;
            node.ChildCount = cxnParas.Count;

            if (depth > 0)
            {
                int paraIdx = 0;
                foreach (var para in cxnParas)
                {
                    paraIdx++;
                    var paraRuns = para.Elements<Drawing.Run>().ToList();
                    var paraText = string.Join("", paraRuns.Select(r => r.Text?.Text ?? ""));
                    var paraNode = new DocumentNode
                    {
                        Path = $"{node.Path}/paragraph[{paraIdx}]",
                        Type = "paragraph",
                        Text = paraText,
                        ChildCount = paraRuns.Count
                    };
                    var paraPProps = para.ParagraphProperties;
                    if (paraPProps?.Alignment?.HasValue == true)
                    {
                        paraNode.Format["align"] = MapTextAlignToFriendly(paraPProps.Alignment);
                    }
                    EmitParagraphBreakProps(paraPProps, paraNode);
                    if (depth > 1)
                    {
                        int runIdx = 0;
                        foreach (var runChild in para.Elements<Drawing.Run>())
                        {
                            runIdx++;
                            paraNode.Children.Add(RunToNode(runChild,
                                $"{node.Path}/paragraph[{paraIdx}]/run[{runIdx}]", part));
                        }
                    }
                    node.Children.Add(paraNode);
                }
            }
        }

        return node;
    }

    // R57 bt-4: connector <p:txBody> lives outside the SDK's typed p:cxnSp
    // schema and parses as an OpenXmlUnknownElement subtree. Locate it by
    // LocalName ("txBody") and reparse its OuterXml as a strongly-typed
    // Presentation.TextBody so callers can walk a:p / a:r / a:t through the
    // regular typed accessors. Also tolerates the legal-but-rare case where
    // the SDK does recognize the txBody (e.g. if a future schema update lifts
    // the restriction): GetFirstChild<TextBody>() returns it directly.
    internal static DocumentFormat.OpenXml.Presentation.TextBody? ResolveConnectorTextBody(ConnectionShape cxn)
    {
        var typed = cxn.GetFirstChild<DocumentFormat.OpenXml.Presentation.TextBody>();
        if (typed != null) return typed;
        var unk = cxn.ChildElements.OfType<OpenXmlUnknownElement>()
            .FirstOrDefault(e => e.LocalName == "txBody");
        if (unk == null) return null;
        try
        {
            return new DocumentFormat.OpenXml.Presentation.TextBody(unk.OuterXml);
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Reconstruct an SVG-like path string from a CustomGeometry element's path list.
    /// </summary>
    private static string ReconstructCustomGeometryPath(Drawing.CustomGeometry custGeom)
    {
        var sb = new StringBuilder();
        var pathList = custGeom.GetFirstChild<Drawing.PathList>();
        if (pathList == null) return "custom";

        foreach (var path in pathList.Elements<Drawing.Path>())
        {
            foreach (var child in path.ChildElements)
            {
                switch (child)
                {
                    case Drawing.MoveTo mt:
                        var mPt = mt.GetFirstChild<Drawing.Point>();
                        if (mPt != null)
                            sb.Append($"M{mPt.X?.Value ?? "0"},{mPt.Y?.Value ?? "0"} ");
                        break;
                    case Drawing.LineTo lt:
                        var lPt = lt.GetFirstChild<Drawing.Point>();
                        if (lPt != null)
                            sb.Append($"L{lPt.X?.Value ?? "0"},{lPt.Y?.Value ?? "0"} ");
                        break;
                    case Drawing.CubicBezierCurveTo cb:
                        var pts = cb.Elements<Drawing.Point>().ToList();
                        if (pts.Count >= 3)
                            sb.Append($"C{pts[0].X?.Value ?? "0"},{pts[0].Y?.Value ?? "0"} {pts[1].X?.Value ?? "0"},{pts[1].Y?.Value ?? "0"} {pts[2].X?.Value ?? "0"},{pts[2].Y?.Value ?? "0"} ");
                        break;
                    case Drawing.QuadraticBezierCurveTo qb:
                        var qPts = qb.Elements<Drawing.Point>().ToList();
                        if (qPts.Count >= 2)
                            sb.Append($"Q{qPts[0].X?.Value ?? "0"},{qPts[0].Y?.Value ?? "0"} {qPts[1].X?.Value ?? "0"},{qPts[1].Y?.Value ?? "0"} ");
                        break;
                    case Drawing.ArcTo at:
                        sb.Append($"A{at.WidthRadius?.Value ?? "0"},{at.HeightRadius?.Value ?? "0"} ");
                        break;
                    case Drawing.CloseShapePath:
                        sb.Append("Z ");
                        break;
                }
            }
        }

        return sb.ToString().Trim();
    }

    // GUID → CLI short-name lookup moved to Core/TableStyles/TableStyleRegistry.
    // Call OfficeCli.Core.TableStyles.TableStyleRegistry.GuidToShortName(guid).

    // Table-level border aggregation. PPT OOXML has no <a:tblBorders>; the
    // visual "table border" is the union of outer cell borders. We sample the
    // outer edge cells: top of row 1, bottom of last row, left of column 1,
    // right of last column. If every cell along an edge agrees, emit a
    // canonical 'border.<side>' summary; if all four sides match, also emit
    // 'border.all'. Mixed/empty edges are simply omitted (consumers should
    // descend to per-cell readback to inspect heterogeneous borders).
    private static void AggregateTableOuterBorders(
        Drawing.Table table,
        List<Drawing.TableRow> rows,
        DocumentNode node)
    {
        if (rows.Count == 0) return;
        string? FormatBorder(OpenXmlCompositeElement? lp)
        {
            if (lp == null) return null;
            if (lp.GetFirstChild<Drawing.NoFill>() != null) return "none";
            var solidFill = lp.GetFirstChild<Drawing.SolidFill>();
            if (solidFill == null) return null;
            var color = ReadColorFromFill(solidFill);
            var wAttr = lp.GetAttributes().FirstOrDefault(a => a.LocalName == "w");
            var dash = lp.GetFirstChild<Drawing.PresetDash>();
            var parts = new List<string>();
            if (!string.IsNullOrEmpty(wAttr.Value) && long.TryParse(wAttr.Value, out var w) && w > 0)
                parts.Add(FormatEmu(w));
            parts.Add(dash?.Val?.HasValue == true ? dash.Val.InnerText! : "solid");
            if (color != null) parts.Add(color);
            return string.Join(" ", parts);
        }

        string? AggregateEdge(IEnumerable<Drawing.TableCell> cells, Func<Drawing.TableCellProperties, OpenXmlCompositeElement?> pick)
        {
            string? agreed = null;
            bool first = true;
            int count = 0;
            foreach (var cell in cells)
            {
                count++;
                var tcPr = cell.TableCellProperties ?? cell.GetFirstChild<Drawing.TableCellProperties>();
                var v = tcPr == null ? null : FormatBorder(pick(tcPr));
                if (first) { agreed = v; first = false; }
                else if (v != agreed) return null; // edge not uniform
            }
            return count == 0 ? null : agreed;
        }

        var topCells = rows[0].Elements<Drawing.TableCell>();
        var bottomCells = rows[^1].Elements<Drawing.TableCell>();
        var leftCells = rows.Select(r => r.Elements<Drawing.TableCell>().FirstOrDefault()).Where(c => c != null)!;
        var rightCells = rows.Select(r => r.Elements<Drawing.TableCell>().LastOrDefault()).Where(c => c != null)!;

        var top = AggregateEdge(topCells, t => t.TopBorderLineProperties);
        var bottom = AggregateEdge(bottomCells!, t => t.BottomBorderLineProperties);
        var left = AggregateEdge(leftCells!, t => t.LeftBorderLineProperties);
        var right = AggregateEdge(rightCells!, t => t.RightBorderLineProperties);

        if (top != null) node.Format["border.top"] = top;
        if (bottom != null) node.Format["border.bottom"] = bottom;
        if (left != null) node.Format["border.left"] = left;
        if (right != null) node.Format["border.right"] = right;

        if (top != null && top == bottom && top == left && top == right)
            node.Format["border.all"] = top;
    }

    // ==================== Effective Properties Resolution (PPT) ====================
    // CONSISTENCY(effective-X-mirror): see PowerPointHandler.StyleList.cs.
    // PopulateEffectiveShapeProperties / PopulateEffectiveRunProperties live
    // there now alongside the unified cascade resolver. What remains here are
    // direction-inheritance helpers (rtl is intentionally out-of-scope for the
    // per-key cascade — see project_pptx_dump_design_decisions.md).


    /// <summary>
    /// Gets the presentation-level DefaultTextStyle by navigating from a SlidePart.
    /// </summary>
    private static OpenXmlCompositeElement? GetPresentationDefaultTextStyle(SlidePart slidePart)
    {
        // Navigate: SlidePart → SlideLayoutPart → SlideMasterPart → PresentationPart → Presentation
        var masterPart = slidePart.SlideLayoutPart?.SlideMasterPart;
        if (masterPart == null) return null;

        // The SlideMasterPart's parent relationships include the PresentationPart
        // We can access the Presentation through the package
        foreach (var rel in masterPart.Parts)
        {
            if (rel.OpenXmlPart is PresentationPart presPart)
                return presPart.Presentation?.DefaultTextStyle;
        }

        return null;
    }

    /// <summary>
    /// Walk slideLayout → slideMaster placeholder defaults looking for an
    /// explicit pPr.RightToLeft. Returns the first hit (true/false) or null
    /// when no ancestor declares a direction. Used by ShapeToNode to populate
    /// `effective.direction` when the slide-level shape doesn't set it itself.
    /// </summary>
    private static bool? ResolveInheritedDirection(SlidePart slidePart, PlaceholderValues? phType = null, bool isTitle = false)
    {
        bool? Probe(OpenXmlElement? root)
        {
            if (root == null) return null;
            foreach (var sp in root.Descendants<Shape>())
            {
                foreach (var para in sp.TextBody?.Elements<Drawing.Paragraph>() ?? Enumerable.Empty<Drawing.Paragraph>())
                {
                    var rtl = para.ParagraphProperties?.RightToLeft;
                    if (rtl?.HasValue == true) return rtl.Value;
                }
            }
            return null;
        }

        var layoutHit = Probe(slidePart.SlideLayoutPart?.SlideLayout?.CommonSlideData?.ShapeTree);
        if (layoutHit.HasValue) return layoutHit;

        var masterHit = Probe(slidePart.SlideLayoutPart?.SlideMasterPart?.SlideMaster?.CommonSlideData?.ShapeTree);
        if (masterHit.HasValue) return masterHit;

        // Final fallback: master-wide <p:txStyles> defaults
        // (bodyStyle/titleStyle/otherStyle Level1 lvl1pPr rtl). Set on
        // /slidelayout[N] or /slidemaster[N] with --prop direction=rtl writes
        // here; this is the only inheritance surface for blank layouts that
        // ship without placeholder shapes.
        var txStyles = slidePart.SlideLayoutPart?.SlideMasterPart?.SlideMaster?.TextStyles;
        if (txStyles != null)
        {
            // R8-4: route by placeholder type. titleStyle is the inheritance
            // surface for Title / CenteredTitle; bodyStyle for Body / SubTitle
            // / Object; otherStyle for everything else and for non-placeholder
            // shapes (mirrors ResolveEffectiveBold / ResolveEffectiveColor —
            // the otherStyle surface is the canonical default for free shapes).
            OpenXmlCompositeElement? styleList;
            if (isTitle || phType == PlaceholderValues.Title || phType == PlaceholderValues.CenteredTitle)
                styleList = txStyles.TitleStyle;
            else if (phType == PlaceholderValues.Body || phType == PlaceholderValues.SubTitle || phType == PlaceholderValues.Object)
                styleList = txStyles.BodyStyle;
            else
                styleList = txStyles.OtherStyle;

            if (styleList != null)
            {
                var lvl1 = styleList.GetFirstChild<Drawing.Level1ParagraphProperties>();
                var rtl = lvl1?.RightToLeft;
                if (rtl?.HasValue == true) return rtl.Value;
            }
        }
        return null;
    }
}
