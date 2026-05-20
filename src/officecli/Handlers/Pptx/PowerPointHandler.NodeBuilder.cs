// Copyright 2025 OfficeCLI (officecli.ai)
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
        if (hex.Length == 6 && hex.All(Uri.IsHexDigit))
        {
            // Schema readback contract: '#RRGGBBAA'. Ensure both '#' and the
            // trailing 'FF' alpha byte are present.
            return $"#{hex.ToUpperInvariant()}FF{tail}";
        }
        if (hex.Length == 8 && hex.All(Uri.IsHexDigit) && !hadHash)
        {
            // Already 8-digit but lacking '#' — add the leading hash.
            return $"#{hex.ToUpperInvariant()}{tail}";
        }
        return color;
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
        int shapeIdx = 0;
        foreach (var shape in container.Elements<Shape>())
        {
            children.Add(ShapeToNode(shape, slideNum, shapeIdx + 1, depth, slidePart, parentPathPrefix));
            shapeIdx++;
        }

        int tblIdx = 0;
        int chartIdx = 0;
        foreach (var gf in container.Elements<GraphicFrame>())
        {
            if (gf.Descendants<Drawing.Table>().Any())
            {
                tblIdx++;
                children.Add(TableToNode(gf, slideNum, tblIdx, depth, parentPathPrefix));
            }
            else if (gf.Descendants<C.ChartReference>().Any() || IsExtendedChartFrame(gf))
            {
                chartIdx++;
                children.Add(ChartToNode(gf, slidePart, slideNum, chartIdx, depth, parentPathPrefix));
            }
        }

        int picIdx = 0;
        foreach (var pic in container.Elements<Picture>())
        {
            children.Add(PictureToNode(pic, slideNum, picIdx + 1, slidePart, parentPathPrefix));
            picIdx++;
        }

        var contentElements = container.ChildElements
            .Where(e => e is Shape or Picture or GraphicFrame or GroupShape or ConnectionShape).ToList();

        int grpIdx = 0;
        foreach (var grp in container.Elements<GroupShape>())
        {
            grpIdx++;
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
            var grpXfrm = grp.GroupShapeProperties?.TransformGroup;
            if (grpXfrm?.Offset?.X != null) grpNode.Format["x"] = FormatEmu(grpXfrm.Offset.X.Value);
            if (grpXfrm?.Offset?.Y != null) grpNode.Format["y"] = FormatEmu(grpXfrm.Offset.Y.Value);
            if (grpXfrm?.Extents?.Cx != null) grpNode.Format["width"] = FormatEmu(grpXfrm.Extents.Cx.Value);
            if (grpXfrm?.Extents?.Cy != null) grpNode.Format["height"] = FormatEmu(grpXfrm.Extents.Cy.Value);
            if (grpXfrm?.Rotation != null && grpXfrm.Rotation.Value != 0)
                grpNode.Format["rotation"] = $"{grpXfrm.Rotation.Value / 60000.0:0.######}";
            var grpFillColor = ReadColorFromFill(grp.GroupShapeProperties?.GetFirstChild<Drawing.SolidFill>());
            if (grpFillColor != null) grpNode.Format["fill"] = grpFillColor;
            else if (grp.GroupShapeProperties?.GetFirstChild<Drawing.NoFill>() != null) grpNode.Format["fill"] = "none";
            else if (grp.GroupShapeProperties?.GetFirstChild<Drawing.GradientFill>() != null) grpNode.Format["fill"] = "gradient";
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
            children.Add(grpNode);
        }

        int cxnIdx = 0;
        foreach (var cxn in container.Elements<ConnectionShape>())
        {
            cxnIdx++;
            children.Add(ConnectorToNode(cxn, slideNum, cxnIdx, parentPathPrefix));
        }

        // Zoom and 3D model are slide-level only; they are not valid
        // children of a GroupShape per the OOXML schema, so only enumerate
        // them when we're at the slide root.
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

    private static DocumentNode TableToNode(GraphicFrame gf, int slideNum, int tblIdx, int depth, string? parentPathPrefix = null)
    {
        var table = gf.Descendants<Drawing.Table>().First();
        var rows = table.Elements<Drawing.TableRow>().ToList();
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
        node.Format["rows"] = rows.Count;
        node.Format["cols"] = cols;

        // Table style
        var tblPr = table.GetFirstChild<Drawing.TableProperties>();
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
            if (tblPr.FirstRow is not null) node.Format["firstRow"] = tblPr.FirstRow.Value;
            if (tblPr.LastRow is not null) node.Format["lastRow"] = tblPr.LastRow.Value;
            if (tblPr.FirstColumn is not null) node.Format["firstCol"] = tblPr.FirstColumn.Value;
            if (tblPr.LastColumn is not null) node.Format["lastCol"] = tblPr.LastColumn.Value;
            if (tblPr.BandRow is not null) node.Format["bandedRows"] = tblPr.BandRow.Value;
            if (tblPr.BandColumn is not null) node.Format["bandedCols"] = tblPr.BandColumn.Value;
        }

        // Outer-edge border aggregation (PPT has no table-level border element).
        // Scan the outer edges across cells; emit per-side keys when uniform,
        // and 'border.all' shorthand when all four sides match.
        AggregateTableOuterBorders(table, rows, node);

        // Position
        var offset = gf.Transform?.Offset;
        if (offset != null)
        {
            if (offset.X is not null) node.Format["x"] = FormatEmu(offset.X!);
            if (offset.Y is not null) node.Format["y"] = FormatEmu(offset.Y!);
        }
        var extents = gf.Transform?.Extents;
        if (extents != null)
        {
            if (extents.Cx is not null) node.Format["width"] = FormatEmu(extents.Cx!);
            if (extents.Cy is not null) node.Format["height"] = FormatEmu(extents.Cy!);
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
                    rowNode.Format["height"] = FormatEmu(row.Height.Value);

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
                        }

                        // Cell borders (including diagonal tl2br/tr2bl)
                        if (tcPr != null) ReadTableCellBorders(tcPr, cellNode);

                        // BUG-R6-A: cell padding readback (Set wrote LeftMargin/etc; Get
                        // missed it on the NodeBuilder cell branch). Canonical key is
                        // "padding.*" per cross-handler rule (root CLAUDE.md).
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
                            if (cellEa != null && cellEa != cellLatin) cellNode.Format["font.ea"] = cellEa;
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
                            if (rp.Baseline?.HasValue == true && rp.Baseline.Value != 0)
                                cellNode.Format["baseline"] = $"{rp.Baseline.Value / 1000.0:0.##}";
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
                            var cellLsPct = cellFirstPProps.GetFirstChild<Drawing.LineSpacing>()?.GetFirstChild<Drawing.SpacingPercent>()?.Val?.Value;
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
                    if (gf.Descendants<Drawing.Table>().Any())
                    {
                        tblIdx++;
                        yield return new RenderableYield(gf, parentPath, "table", tblIdx);
                    }
                    else if (gf.Descendants<C.ChartReference>().Any() || IsExtendedChartFrame(gf))
                    {
                        chartIdx++;
                        yield return new RenderableYield(gf, parentPath, "chart", chartIdx);
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
        if (phElemForNode?.Index?.Value is uint phIdx) node.Format["phIndex"] = (int)phIdx;

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
        if (isTitle) node.Format["isTitle"] = true;

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

                // Gradient opacity (from first stop's alpha)
                var gradAlpha = stops[0].GetFirstChild<Drawing.RgbColorModelHex>()?.GetFirstChild<Drawing.Alpha>()?.Val?.Value
                    ?? stops[0].GetFirstChild<Drawing.SchemeColor>()?.GetFirstChild<Drawing.Alpha>()?.Val?.Value;
                if (gradAlpha.HasValue) node.Format["opacity"] = $"{gradAlpha.Value / 100000.0:0.##}";
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
        }

        // Image (blip) fill on shape
        var blipFill = shape.ShapeProperties?.GetFirstChild<Drawing.BlipFill>();
        if (blipFill != null) node.Format["image"] = "true";

        // Pattern fill on shape — round-trip the input form "preset:fg:bg".
        var patternFill = shape.ShapeProperties?.GetFirstChild<Drawing.PatternFill>();
        if (patternFill != null)
        {
            var preset = patternFill.Preset?.InnerText ?? "";
            var fgEl = patternFill.GetFirstChild<Drawing.ForegroundColor>();
            var bgEl = patternFill.GetFirstChild<Drawing.BackgroundColor>();
            var fgHex = fgEl?.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value;
            var fgScheme = fgEl?.GetFirstChild<Drawing.SchemeColor>()?.Val?.Value.ToString();
            var bgHex = bgEl?.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value;
            var bgScheme = bgEl?.GetFirstChild<Drawing.SchemeColor>()?.Val?.Value.ToString();
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
            var charBullet = firstParaBullet.GetFirstChild<Drawing.CharacterBullet>();
            var autoBullet = firstParaBullet.GetFirstChild<Drawing.AutoNumberedBullet>();
            if (charBullet != null)
            {
                var charVal = charBullet.Char?.Value ?? "•";
                node.Format["list"] = charVal switch
                {
                    "•" or "●" or "○" => "bullet",
                    "–" or "—" or "-" => "dash",
                    "►" or "▶" or "▸" or "➤" => "arrow",
                    "✓" or "✔" => "check",
                    "★" or "☆" or "⭐" => "star",
                    _ => charVal
                };
            }
            else if (autoBullet?.Type?.HasValue == true)
            {
                var autoVal = autoBullet.Type.InnerText;
                node.Format["list"] = autoVal switch
                {
                    "arabicPeriod" or "arabicParenR" or "arabicPlain" or "arabicParenBoth" => "numbered",
                    "romanLcPeriod" or "romanLcParenR" or "romanLcParenBoth" => "romanLc",
                    "romanUcPeriod" or "romanUcParenR" or "romanUcParenBoth" => "romanUc",
                    "alphaLcPeriod" or "alphaLcParenR" or "alphaLcParenBoth" => "alphaLc",
                    "alphaUcPeriod" or "alphaUcParenR" or "alphaUcParenBoth" => "alphaUc",
                    _ => autoVal
                };
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
        if (allRuns.Count > 1)
        {
            string? firstSizeKey = null;
            string? firstColorKey = null;
            bool sizeSeen = false, colorSeen = false;
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
            if (fontEaTf != null && fontEaTf != fontLatinTf) node.Format["font.ea"] = fontEaTf;
            if (fontCsTf != null) node.Format["font.cs"] = fontCsTf;

            var fontSize = firstRun.RunProperties.FontSize?.Value;
            if (fontSize.HasValue && !hasMixedSize) node.Format["size"] = $"{fontSize.Value / 100.0:0.##}pt";

            if (firstRun.RunProperties.Bold?.HasValue == true) node.Format["bold"] = firstRun.RunProperties.Bold.Value;
            if (firstRun.RunProperties.Italic?.HasValue == true) node.Format["italic"] = firstRun.RunProperties.Italic.Value;
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

            // Character spacing on first run
            if (firstRun.RunProperties.Spacing?.HasValue == true)
                node.Format["spacing"] = $"{firstRun.RunProperties.Spacing.Value / 100.0:0.##}";
            // Baseline (superscript/subscript)
            if (firstRun.RunProperties.Baseline?.HasValue == true && firstRun.RunProperties.Baseline.Value != 0)
                node.Format["baseline"] = $"{firstRun.RunProperties.Baseline.Value / 1000.0:0.##}";

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
            // rPr attributes (lang, kern, kumimoji, normalizeH, ...) at shape
            // level too, mirroring BuildRunNode. Without this, shape-level Add
            // can write `lang` to first-run rPr but shape-level Get cannot
            // surface it unless the user descends to /shape[N]/r[1] explicitly.
            FillUnknownRunProps(firstRun.RunProperties, node);
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
            else if (outline.GetFirstChild<Drawing.Miter>() != null)
                node.Format["lineJoin"] = "miter";
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

        // Effects (shadow, glow, reflection) — check shape-level first, then text run-level
        var effectList = shape.ShapeProperties?.GetFirstChild<Drawing.EffectList>();
        // Fall back to first text run's effectLst (used for fill=none shapes)
        var textEffectList = effectList == null || (!effectList.HasChildren)
            ? shape.TextBody?.Descendants<Drawing.RunProperties>()
                .Select(rp => rp.GetFirstChild<Drawing.EffectList>())
                .FirstOrDefault(el => el != null)
            : null;
        var activeEffectList = effectList?.HasChildren == true ? effectList : textEffectList;
        if (activeEffectList != null)
        {
            var outerShadow = activeEffectList.GetFirstChild<Drawing.OuterShadow>();
            if (outerShadow != null)
            {
                var shadowColor = EnsureEightDigitHexForEffect(ReadColorFromElement(outerShadow) ?? "000000");
                var blurPt = outerShadow.BlurRadius?.HasValue == true ? $"{outerShadow.BlurRadius.Value / 12700.0:0.##}" : "4";
                var angleDeg = outerShadow.Direction?.HasValue == true ? $"{outerShadow.Direction.Value / 60000.0:0.##}" : "45";
                var distPt = outerShadow.Distance?.HasValue == true ? $"{outerShadow.Distance.Value / 12700.0:0.##}" : "3";
                var alphaEl = outerShadow.Descendants<Drawing.Alpha>().FirstOrDefault();
                // OOXML default: <a:outerShdw> without <a:alpha> is fully opaque
                // (the shadow inherits the color element's alpha; an absent alpha
                // means 100%). Defaulting to "40" used to mask explicit
                // alpha=FF inputs as a 40% shadow on round-trip.
                var opacity = alphaEl?.Val?.HasValue == true ? $"{alphaEl.Val.Value / 1000.0:0.##}" : "100";
                node.Format["shadow"] = $"{shadowColor}-{blurPt}-{angleDeg}-{distPt}-{opacity}";
            }
            var glow = activeEffectList.GetFirstChild<Drawing.Glow>();
            if (glow != null)
            {
                var glowColor = EnsureEightDigitHexForEffect(ReadColorFromElement(glow) ?? "000000");
                var radiusPt = glow.Radius?.HasValue == true ? $"{glow.Radius.Value / 12700.0:0.##}" : "8";
                var glowAlpha = glow.Descendants<Drawing.Alpha>().FirstOrDefault();
                // OOXML default: <a:glow> without <a:alpha> is fully opaque.
                var glowOpacity = glowAlpha?.Val?.HasValue == true ? $"{glowAlpha.Val.Value / 1000.0:0.##}" : "100";
                node.Format["glow"] = $"{glowColor}-{radiusPt}-{glowOpacity}";
            }
            var reflEl = activeEffectList.GetFirstChild<Drawing.Reflection>();
            if (reflEl != null)
            {
                // Map endPosition back to type: tight=55000, half=90000, full=100000
                var endPos = reflEl.EndPosition?.Value ?? 0;
                if (endPos >= 95000) node.Format["reflection"] = "full";
                else if (endPos >= 70000) node.Format["reflection"] = "half";
                else node.Format["reflection"] = "tight";
            }
            var softEdge = activeEffectList.GetFirstChild<Drawing.SoftEdge>();
            if (softEdge?.Radius?.HasValue == true)
                // Unit-qualified pt — matches the cross-format canonical from
                // root CLAUDE.md (line.width "0.75pt", padding "12pt", glow
                // "4pt"). The bare numeric form here was the lone outlier on
                // the effects readback surface and broke dump round-trip
                // when set softEdge=<value> re-parses the readback.
                node.Format["softEdge"] = $"{softEdge.Radius.Value / 12700.0:0.##}pt";
        }

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
            var lightRig = scene3d.LightRig;
            if (lightRig?.Rig?.HasValue == true) node.Format["lighting"] = lightRig.Rig.InnerText;
        }

        // 3D format (sp3d)
        var sp3d = shape.ShapeProperties?.GetFirstChild<Drawing.Shape3DType>();
        if (sp3d != null)
        {
            if (sp3d.ExtrusionHeight?.HasValue == true && sp3d.ExtrusionHeight.Value != 0)
                node.Format["depth"] = $"{sp3d.ExtrusionHeight.Value / 12700.0:0.##}";
            if (sp3d.PresetMaterial?.HasValue == true)
                node.Format["material"] = sp3d.PresetMaterial.InnerText;
            var bevelT = sp3d.BevelTop;
            if (bevelT != null) node.Format["bevel"] = FormatBevel(bevelT);
            var bevelB = sp3d.BevelBottom;
            if (bevelB != null) node.Format["bevelBottom"] = FormatBevel(bevelB);
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
                    node.Format["margin"] = $"{FormatEmu(lIns ?? 91440)},{FormatEmu(tIns ?? 45720)},{FormatEmu(rIns ?? 91440)},{FormatEmu(bIns ?? 45720)}";
            }

            // Vertical alignment — map XML enum to user-friendly name (like POI TextAlign)
            if (bodyPr.Anchor?.HasValue == true)
            {
                var vaInner = bodyPr.Anchor.InnerText;
                node.Format["valign"] = vaInner switch
                {
                    "t" => "top",
                    "ctr" => "center",
                    "b" => "bottom",
                    _ => vaInner
                };
            }

            // TextWarp (WordArt)
            var prstTxWarp = bodyPr.GetFirstChild<Drawing.PresetTextWarp>();
            if (prstTxWarp?.Preset?.HasValue == true)
                node.Format["textWarp"] = prstTxWarp.Preset.InnerText;

            // AutoFit
            if (bodyPr.GetFirstChild<Drawing.NormalAutoFit>() != null) node.Format["autoFit"] = "normal";
            else if (bodyPr.GetFirstChild<Drawing.ShapeAutoFit>() != null) node.Format["autoFit"] = "shape";
            else node.Format["autoFit"] = "none";
        }

        // Text alignment (from first paragraph). Only surface when explicitly
        // present in the source XML; the previous else-branch hard-coded
        // align=left whenever pPr/algn was absent, which baked an explicit
        // value into every round-trip and broke inheritance from the layout/
        // master defRPr cascade. Callers that need the effective alignment
        // can read Format["effective.align"] (cascade-resolved separately).
        var firstPara = shape.TextBody?.Elements<Drawing.Paragraph>().FirstOrDefault();
        if (firstPara?.ParagraphProperties?.Alignment?.HasValue == true)
        {
            var alInner = firstPara.ParagraphProperties.Alignment.InnerText;
            node.Format["align"] = alInner switch
            {
                "l" => "left",
                "ctr" => "center",
                "r" => "right",
                "just" => "justify",
                _ => alInner
            };
        }

        // Paragraph spacing and indent (from first paragraph)
        var pProps = firstPara?.ParagraphProperties;
        if (pProps != null)
        {
            var lsPct = pProps.GetFirstChild<Drawing.LineSpacing>()?.GetFirstChild<Drawing.SpacingPercent>()?.Val?.Value;
            if (lsPct.HasValue) node.Format["lineSpacing"] = SpacingConverter.FormatPptLineSpacingPercent(lsPct.Value);
            var lsPts = pProps.GetFirstChild<Drawing.LineSpacing>()?.GetFirstChild<Drawing.SpacingPoints>()?.Val?.Value;
            if (lsPts.HasValue) node.Format["lineSpacing"] = SpacingConverter.FormatPptLineSpacingPoints(lsPts.Value);
            var sb = pProps.GetFirstChild<Drawing.SpaceBefore>()?.GetFirstChild<Drawing.SpacingPoints>()?.Val?.Value;
            if (sb.HasValue) node.Format["spaceBefore"] = SpacingConverter.FormatPptSpacing(sb.Value);
            var sa = pProps.GetFirstChild<Drawing.SpaceAfter>()?.GetFirstChild<Drawing.SpacingPoints>()?.Val?.Value;
            if (sa.HasValue) node.Format["spaceAfter"] = SpacingConverter.FormatPptSpacing(sa.Value);
            if (pProps.Indent?.HasValue == true) node.Format["indent"] = FormatEmu(pProps.Indent.Value);
            if (pProps.LeftMargin?.HasValue == true) node.Format["marginLeft"] = FormatEmu(pProps.LeftMargin.Value);
            if (pProps.RightMargin?.HasValue == true) node.Format["marginRight"] = FormatEmu(pProps.RightMargin.Value);
            // Reading direction (Arabic / Hebrew). Only emit when explicitly
            // set so LTR docs don't get a noisy `direction=ltr` everywhere.
            if (pProps.RightToLeft?.HasValue == true)
                node.Format["direction"] = pProps.RightToLeft.Value ? "rtl" : "ltr";
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
                        var paraAlignVal = paraPProps.Alignment.Value;
                        paraNode.Format["align"] = paraAlignVal == Drawing.TextAlignmentTypeValues.Center ? "center"
                            : paraAlignVal == Drawing.TextAlignmentTypeValues.Right ? "right"
                            : paraAlignVal == Drawing.TextAlignmentTypeValues.Justified ? "justify"
                            : "left";
                    }
                    if (paraPProps?.Level?.HasValue == true) paraNode.Format["level"] = paraPProps.Level.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
                    if (paraPProps?.Indent?.HasValue == true) paraNode.Format["indent"] = FormatEmu(paraPProps.Indent.Value);
                    if (paraPProps?.LeftMargin?.HasValue == true) paraNode.Format["marginLeft"] = FormatEmu(paraPProps.LeftMargin.Value);
                    if (paraPProps?.RightMargin?.HasValue == true) paraNode.Format["marginRight"] = FormatEmu(paraPProps.RightMargin.Value);
                    if (paraPProps?.RightToLeft?.HasValue == true)
                        paraNode.Format["direction"] = paraPProps.RightToLeft.Value ? "rtl" : "ltr";
                    var pLsPct = paraPProps?.GetFirstChild<Drawing.LineSpacing>()?.GetFirstChild<Drawing.SpacingPercent>()?.Val?.Value;
                    if (pLsPct.HasValue) paraNode.Format["lineSpacing"] = SpacingConverter.FormatPptLineSpacingPercent(pLsPct.Value);
                    var pLsPts = paraPProps?.GetFirstChild<Drawing.LineSpacing>()?.GetFirstChild<Drawing.SpacingPoints>()?.Val?.Value;
                    if (pLsPts.HasValue) paraNode.Format["lineSpacing"] = SpacingConverter.FormatPptLineSpacingPoints(pLsPts.Value);
                    var pSb = paraPProps?.GetFirstChild<Drawing.SpaceBefore>()?.GetFirstChild<Drawing.SpacingPoints>()?.Val?.Value;
                    if (pSb.HasValue) paraNode.Format["spaceBefore"] = SpacingConverter.FormatPptSpacing(pSb.Value);
                    var pSa = paraPProps?.GetFirstChild<Drawing.SpaceAfter>()?.GetFirstChild<Drawing.SpacingPoints>()?.Val?.Value;
                    if (pSa.HasValue) paraNode.Format["spaceAfter"] = SpacingConverter.FormatPptSpacing(pSa.Value);

                    // Include runs at depth > 1
                    if (depth > 1)
                    {
                        int runIdx = 0;
                        foreach (var run in paraRuns)
                        {
                            paraNode.Children.Add(RunToNode(run,
                                $"{shapePath}/paragraph[{paraIdx + 1}]/run[{runIdx + 1}]", part));
                            runIdx++;
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
            // root CLAUDE.md "Canonical DocumentNode.Format Rules".
            if (fLatin != null) node.Format["font.latin"] = fLatin;
            if (fEa != null && fEa != fLatin) node.Format["font.ea"] = fEa;
            if (fCs != null) node.Format["font.cs"] = fCs;
            var fs = run.RunProperties.FontSize?.Value;
            if (fs.HasValue) node.Format["size"] = $"{fs.Value / 100.0:0.##}pt";
            if (run.RunProperties.Bold?.Value == true) node.Format["bold"] = true;
            if (run.RunProperties.Italic?.Value == true) node.Format["italic"] = true;
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
            if (run.RunProperties.Baseline?.HasValue == true && run.RunProperties.Baseline.Value != 0)
                node.Format["baseline"] = $"{run.RunProperties.Baseline.Value / 1000.0:0.##}";
            // Color (solid or gradient)
            var runFillColor = ReadColorFromFill(run.RunProperties.GetFirstChild<Drawing.SolidFill>());
            if (runFillColor != null) node.Format["color"] = runFillColor;
            var runGrad = run.RunProperties.GetFirstChild<Drawing.GradientFill>();
            if (runGrad != null) node.Format["textFill"] = ReadGradientString(runGrad);
            // Hyperlink (link + tooltip — round-trips Add/Set 'tooltip=…').
            if (part != null)
            {
                var runHl = run.RunProperties.GetFirstChild<Drawing.HyperlinkOnClick>();
                var linkUrl = ReadHyperlinkOnClickUrl(runHl, part);
                if (linkUrl != null) node.Format["link"] = linkUrl;
                var runTip = runHl?.Tooltip?.Value;
                if (!string.IsNullOrEmpty(runTip)) node.Format["tooltip"] = runTip!;
            }

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
        "b", "i", "u", "strike", "sz", "spc", "baseline",
    };

    private static readonly System.Collections.Generic.HashSet<string> CuratedRunChildren =
        new(System.StringComparer.Ordinal)
    {
        "latin", "ea", "cs", "solidFill", "gradFill", "hlinkClick",
    };

    private static void FillUnknownRunProps(Drawing.RunProperties? rPr, DocumentNode node)
    {
        if (rPr == null) return;

        // Walk attributes on rPr itself.
        foreach (var attr in rPr.GetAttributes())
        {
            var name = attr.LocalName;
            if (string.IsNullOrEmpty(name)) continue;
            if (CuratedRunAttrs.Contains(name)) continue;
            if (node.Format.ContainsKey(name)) continue;
            node.Format[name] = attr.Value;
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
        if (!isVideo && !isAudio)
        {
            if (!string.IsNullOrEmpty(alt)) node.Format["alt"] = alt;
            else node.Format["alt"] = "(missing)";
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

        // Brightness / contrast — stored on the same blip as
        // a:lumOff (brightness) and a:lumMod (contrast). Mirrors the
        // write side in Set.Media.cs (`case "brightness" or "contrast"`).
        // Note: the SDK's Blip class doesn't strong-type lumMod/lumOff as
        // direct children (they're effect children per the a:CT_Blip
        // schema but live in a content-group the SDK marks "unknown" once
        // round-tripped). Read via LocalName/attribute so we tolerate both
        // the strongly-typed append we do in Set.Media and the unknown
        // element form we see on re-parse.
        if (picBlip != null)
        {
            int? lumModVal = null, lumOffVal = null;
            foreach (var kid in picBlip.ChildElements)
            {
                if (kid.NamespaceUri != "http://schemas.openxmlformats.org/drawingml/2006/main") continue;
                var valAttr = kid.GetAttribute("val", "").Value;
                if (string.IsNullOrEmpty(valAttr) || !int.TryParse(valAttr, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out var iv)) continue;
                if (kid.LocalName == "lumMod") lumModVal = iv;
                else if (kid.LocalName == "lumOff") lumOffVal = iv;
            }
            if (lumOffVal.HasValue && lumOffVal.Value != 0)
                node.Format["brightness"] = $"{lumOffVal.Value / 1000.0:0.##}";
            if (lumModVal.HasValue && lumModVal.Value != 100000)
                node.Format["contrast"] = $"{(lumModVal.Value - 100000) / 1000.0:0.##}";
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
                var blurPt = picOuterShadow.BlurRadius?.HasValue == true ? $"{picOuterShadow.BlurRadius.Value / 12700.0:0.##}" : "4";
                var angleDeg = picOuterShadow.Direction?.HasValue == true ? $"{picOuterShadow.Direction.Value / 60000.0:0.##}" : "45";
                var distPt = picOuterShadow.Distance?.HasValue == true ? $"{picOuterShadow.Distance.Value / 12700.0:0.##}" : "3";
                var alphaEl = picOuterShadow.Descendants<Drawing.Alpha>().FirstOrDefault();
                var opacity = alphaEl?.Val?.HasValue == true ? $"{alphaEl.Val.Value / 1000.0:0.##}" : "40";
                node.Format["shadow"] = $"{shadowColor}-{blurPt}-{angleDeg}-{distPt}-{opacity}";
            }
            var picGlow = picEffectList.GetFirstChild<Drawing.Glow>();
            if (picGlow != null)
            {
                var glowColor = EnsureEightDigitHexForEffect(ReadColorFromElement(picGlow) ?? "000000");
                var radiusPt = picGlow.Radius?.HasValue == true ? $"{picGlow.Radius.Value / 12700.0:0.##}" : "8";
                var glowAlpha = picGlow.Descendants<Drawing.Alpha>().FirstOrDefault();
                var glowOpacity = glowAlpha?.Val?.HasValue == true ? $"{glowAlpha.Val.Value / 1000.0:0.##}" : "75";
                node.Format["glow"] = $"{glowColor}-{radiusPt}-{glowOpacity}";
            }
        }

        // Crop
        var srcRect = pic.BlipFill?.GetFirstChild<Drawing.SourceRectangle>();
        if (srcRect != null)
        {
            var cl = srcRect.Left?.Value ?? 0;
            var ct = srcRect.Top?.Value ?? 0;
            var cr = srcRect.Right?.Value ?? 0;
            var cb = srcRect.Bottom?.Value ?? 0;
            if (cl != 0 || ct != 0 || cr != 0 || cb != 0)
                node.Format["crop"] = $"{cl / 1000.0:0.##},{ct / 1000.0:0.##},{cr / 1000.0:0.##},{cb / 1000.0:0.##}";
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
            // Default title position: top-center area of standard 16:9 slide
            spPr.Transform2D = new Drawing.Transform2D
            {
                Offset = new Drawing.Offset { X = 838200, Y = 365125 },    // ~2.33cm, ~1.01cm
                Extents = new Drawing.Extents { Cx = 10515600, Cy = 1325563 } // ~29.21cm, ~3.68cm
            };
        }
        else
        {
            // Default body/content position: below title
            spPr.Transform2D = new Drawing.Transform2D
            {
                Offset = new Drawing.Offset { X = 838200, Y = 1825625 },   // ~2.33cm, ~5.07cm
                Extents = new Drawing.Extents { Cx = 10515600, Cy = 4351338 } // ~29.21cm, ~12.09cm
            };
        }
        shape.ShapeProperties = spPr;
        var body = new TextBody(
            new Drawing.BodyProperties(),
            new Drawing.ListStyle()
        );
        // CONSISTENCY(text-escape-boundary): \n / \t resolution at CLI --prop;
        // text arrives here with real newlines and tabs already.
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
        shape.TextBody = body;
        return shape;
    }

    private static DocumentNode ConnectorToNode(ConnectionShape cxn, int slideNum, int cxnIdx, string? parentPathPrefix = null)
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

        return node;
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
