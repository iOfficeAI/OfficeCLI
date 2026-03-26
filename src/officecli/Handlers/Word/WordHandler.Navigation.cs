// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    // ==================== Navigation ====================

    private DocumentNode GetRootNode(int depth)
    {
        var node = new DocumentNode { Path = "/", Type = "document" };
        var children = new List<DocumentNode>();

        var mainPart = _doc.MainDocumentPart;
        if (mainPart?.Document?.Body != null)
        {
            children.Add(new DocumentNode
            {
                Path = "/body",
                Type = "body",
                ChildCount = mainPart.Document.Body.ChildElements.Count
            });
        }

        if (mainPart?.StyleDefinitionsPart != null)
        {
            children.Add(new DocumentNode
            {
                Path = "/styles",
                Type = "styles",
                ChildCount = mainPart.StyleDefinitionsPart.Styles?.ChildElements.Count ?? 0
            });
        }

        int headerIdx = 0;
        if (mainPart?.HeaderParts != null)
        {
            foreach (var _ in mainPart.HeaderParts)
            {
                children.Add(new DocumentNode
                {
                    Path = $"/header[{headerIdx + 1}]",
                    Type = "header"
                });
                headerIdx++;
            }
        }

        int footerIdx = 0;
        if (mainPart?.FooterParts != null)
        {
            foreach (var _ in mainPart.FooterParts)
            {
                children.Add(new DocumentNode
                {
                    Path = $"/footer[{footerIdx + 1}]",
                    Type = "footer"
                });
                footerIdx++;
            }
        }

        if (mainPart?.NumberingDefinitionsPart != null)
        {
            children.Add(new DocumentNode { Path = "/numbering", Type = "numbering" });
        }

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

        // Page size from last section properties (document default)
        var sectPr = mainPart?.Document?.Body?.GetFirstChild<SectionProperties>()
            ?? mainPart?.Document?.Body?.Descendants<SectionProperties>().LastOrDefault();
        if (sectPr != null)
        {
            var pageSize = sectPr.GetFirstChild<PageSize>();
            if (pageSize?.Width?.Value != null) node.Format["pageWidth"] = pageSize.Width.Value;
            if (pageSize?.Height?.Value != null) node.Format["pageHeight"] = pageSize.Height.Value;
            if (pageSize?.Orient?.Value != null) node.Format["orientation"] = pageSize.Orient.InnerText;
            var margins = sectPr.GetFirstChild<PageMargin>();
            if (margins != null)
            {
                if (margins.Top?.Value != null) node.Format["marginTop"] = margins.Top.Value;
                if (margins.Bottom?.Value != null) node.Format["marginBottom"] = margins.Bottom.Value;
                if (margins.Left?.Value != null) node.Format["marginLeft"] = margins.Left.Value;
                if (margins.Right?.Value != null) node.Format["marginRight"] = margins.Right.Value;
            }
        }

        node.Children = children;
        node.ChildCount = children.Count;
        return node;
    }

    private record PathSegment(string Name, int? Index, string? StringIndex = null);

    private static List<PathSegment> ParsePath(string path)
    {
        var segments = new List<PathSegment>();
        var parts = path.Trim('/').Split('/');

        foreach (var part in parts)
        {
            var bracketIdx = part.IndexOf('[');
            if (bracketIdx >= 0)
            {
                var name = part[..bracketIdx];
                var indexStr = part[(bracketIdx + 1)..^1];
                if (int.TryParse(indexStr, out var idx))
                    segments.Add(new PathSegment(name, idx));
                else
                    segments.Add(new PathSegment(name, null, indexStr));
            }
            else
            {
                segments.Add(new PathSegment(part, null));
            }
        }

        return segments;
    }

    private OpenXmlElement? NavigateToElement(List<PathSegment> segments)
        => NavigateToElement(segments, out _);

    private OpenXmlElement? NavigateToElement(List<PathSegment> segments, out string? availableContext)
    {
        availableContext = null;
        if (segments.Count == 0) return null;

        var first = segments[0];

        // Handle bookmark[Name] as top-level path
        if (first.Name.ToLowerInvariant() == "bookmark" && first.StringIndex != null)
        {
            var body = _doc.MainDocumentPart?.Document?.Body;
            return body?.Descendants<BookmarkStart>()
                .FirstOrDefault(b => b.Name?.Value == first.StringIndex);
        }

        OpenXmlElement? current = first.Name.ToLowerInvariant() switch
        {
            "body" => _doc.MainDocumentPart?.Document?.Body,
            "styles" => _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles,
            "header" => _doc.MainDocumentPart?.HeaderParts.ElementAtOrDefault((first.Index ?? 1) - 1)?.Header,
            "footer" => _doc.MainDocumentPart?.FooterParts.ElementAtOrDefault((first.Index ?? 1) - 1)?.Footer,
            "numbering" => _doc.MainDocumentPart?.NumberingDefinitionsPart?.Numbering,
            "settings" => _doc.MainDocumentPart?.DocumentSettingsPart?.Settings,
            "comments" => _doc.MainDocumentPart?.WordprocessingCommentsPart?.Comments,
            _ => null
        };

        string parentPath = "/" + first.Name + (first.Index.HasValue ? $"[{first.Index}]" : "");

        for (int i = 1; i < segments.Count && current != null; i++)
        {
            var seg = segments[i];
            IEnumerable<OpenXmlElement> children;
            if (current is Body body2 && (seg.Name.ToLowerInvariant() == "p" || seg.Name.ToLowerInvariant() == "tbl"))
            {
                // Only count direct body-level paragraphs/tables, skip those inside SdtBlock containers
                children = seg.Name.ToLowerInvariant() == "p"
                    ? body2.Elements<Paragraph>().Cast<OpenXmlElement>()
                    : body2.Elements<Table>().Cast<OpenXmlElement>();
            }
            else if (current is Body body3 && seg.Name == "oMathPara")
            {
                // oMathPara can be direct body children or wrapped inside w:p elements
                var mathParas = new List<OpenXmlElement>();
                foreach (var el in body3.ChildElements)
                {
                    if (el.LocalName == "oMathPara" || el is M.Paragraph)
                        mathParas.Add(el);
                    else if (el is Paragraph wp)
                    {
                        var inner = wp.ChildElements.FirstOrDefault(c => c.LocalName == "oMathPara" || c is M.Paragraph);
                        if (inner != null) mathParas.Add(inner);
                    }
                }
                children = mathParas;
            }
            else
            {
                children = seg.Name.ToLowerInvariant() switch
                {
                    "p" => current.Elements<Paragraph>().Cast<OpenXmlElement>(),
                    "r" => current.Descendants<Run>()
                        .Where(r => r.GetFirstChild<CommentReference>() == null)
                        .Cast<OpenXmlElement>(),
                    "tbl" => current.Elements<Table>().Cast<OpenXmlElement>(),
                    "tr" => current.Elements<TableRow>().Cast<OpenXmlElement>(),
                    "tc" => current.Elements<TableCell>().Cast<OpenXmlElement>(),
                    "sdt" => current.ChildElements
                        .Where(e => e is SdtBlock || e is SdtRun).Cast<OpenXmlElement>(),
                    _ => current.ChildElements.Where(e => e.LocalName == seg.Name).Cast<OpenXmlElement>()
                };
            }

            var childList = children.ToList();
            var next = seg.Index.HasValue
                ? childList.ElementAtOrDefault(seg.Index.Value - 1)
                : childList.FirstOrDefault();

            if (next == null)
            {
                availableContext = BuildAvailableContext(current, parentPath, seg.Name, childList.Count);
                return null;
            }

            parentPath += "/" + seg.Name + (seg.Index.HasValue ? $"[{seg.Index}]" : "");
            current = next;
        }

        return current;
    }

    /// <summary>
    /// Build a context string describing available children when navigation fails.
    /// </summary>
    private static string BuildAvailableContext(OpenXmlElement parent, string parentPath, string requestedType, int matchCount)
    {
        if (matchCount > 0)
            return $"Available at {parentPath}: {requestedType}[1]..{requestedType}[{matchCount}]";

        // List distinct child types at this level
        var childTypes = parent.ChildElements
            .GroupBy(c => c.LocalName)
            .Select(g => $"{g.Key}({g.Count()})")
            .Take(10)
            .ToList();

        return childTypes.Count > 0
            ? $"No {requestedType} found at {parentPath}. Available children: {string.Join(", ", childTypes)}"
            : $"No children at {parentPath}";
    }

    private DocumentNode ElementToNode(OpenXmlElement element, string path, int depth)
    {
        var node = new DocumentNode { Path = path, Type = element.LocalName };

        if (element is BookmarkStart bkStart)
        {
            node.Type = "bookmark";
            node.Format["name"] = bkStart.Name?.Value ?? "";
            node.Format["id"] = bkStart.Id?.Value ?? "";
            var bkText = GetBookmarkText(bkStart);
            if (!string.IsNullOrEmpty(bkText))
                node.Text = bkText;
            return node;
        }

        if (element is Paragraph para)
        {
            node.Type = "paragraph";
            node.Text = GetParagraphText(para);
            node.Style = GetStyleName(para);
            node.Preview = node.Text?.Length > 50 ? node.Text[..50] + "..." : node.Text;
            node.ChildCount = GetAllRuns(para).Count();

            var pProps = para.ParagraphProperties;
            if (pProps != null)
            {
                if (pProps.ParagraphStyleId?.Val?.Value != null)
                    node.Format["style"] = pProps.ParagraphStyleId.Val.Value;
                if (pProps.Justification?.Val != null)
                {
                    var alignText = pProps.Justification.Val.InnerText;
                    var alignValue = alignText == "both" ? "justify" : alignText;
                    node.Format["alignment"] = alignValue;
                }
                if (pProps.SpacingBetweenLines != null)
                {
                    if (pProps.SpacingBetweenLines.Before?.Value != null)
                    {
                        node.Format["spaceBefore"] = SpacingConverter.FormatWordSpacing(pProps.SpacingBetweenLines.Before.Value);
                    }
                    if (pProps.SpacingBetweenLines.After?.Value != null)
                    {
                        node.Format["spaceAfter"] = SpacingConverter.FormatWordSpacing(pProps.SpacingBetweenLines.After.Value);
                    }
                    if (pProps.SpacingBetweenLines.Line?.Value != null)
                    {
                        node.Format["lineSpacing"] = SpacingConverter.FormatWordLineSpacing(
                            pProps.SpacingBetweenLines.Line.Value,
                            pProps.SpacingBetweenLines.LineRule?.InnerText);
                    }
                }
                if (pProps.Indentation?.FirstLine?.Value != null)
                    node.Format["firstLineIndent"] = pProps.Indentation.FirstLine.Value;
                if (pProps.Indentation?.Left?.Value != null)
                    node.Format["leftIndent"] = pProps.Indentation.Left.Value;
                if (pProps.Indentation?.Right?.Value != null)
                    node.Format["rightIndent"] = pProps.Indentation.Right.Value;
                if (pProps.Indentation?.Hanging?.Value != null)
                    node.Format["hangingIndent"] = pProps.Indentation.Hanging.Value;
                if (pProps.KeepNext != null)
                {
                    node.Format["keepNext"] = true;
                }
                if (pProps.KeepLines != null)
                {
                    node.Format["keepLines"] = true;
                }
                if (pProps.PageBreakBefore != null)
                    node.Format["pageBreakBefore"] = true;
                if (pProps.WidowControl != null)
                {
                    // Val == null or Val == true means enabled; Val == false means explicitly disabled
                    var wcVal = pProps.WidowControl.Val;
                    node.Format["widowControl"] = wcVal == null || wcVal.Value;
                }
                if (pProps.Shading != null)
                {
                    var shdVal = pProps.Shading.Val?.InnerText ?? "";
                    var shdFill = pProps.Shading.Fill?.Value;
                    var shdColor = pProps.Shading.Color?.Value;
                    if (string.Equals(shdVal, "clear", StringComparison.OrdinalIgnoreCase)
                        && !string.IsNullOrEmpty(shdFill)
                        && string.IsNullOrEmpty(shdColor))
                    {
                        node.Format["shd"] = ParseHelpers.FormatHexColor(shdFill);
                    }
                    else
                    {
                        var shdParts = new List<string>();
                        if (!string.IsNullOrEmpty(shdVal)) shdParts.Add(shdVal);
                        if (!string.IsNullOrEmpty(shdFill)) shdParts.Add(ParseHelpers.FormatHexColor(shdFill));
                        if (!string.IsNullOrEmpty(shdColor)) shdParts.Add(ParseHelpers.FormatHexColor(shdColor));
                        node.Format["shd"] = string.Join(";", shdParts);
                    }
                }

                var pBdr = pProps.ParagraphBorders;
                if (pBdr != null)
                {
                    ReadBorder(pBdr.TopBorder, "pbdr.top", node);
                    ReadBorder(pBdr.BottomBorder, "pbdr.bottom", node);
                    ReadBorder(pBdr.LeftBorder, "pbdr.left", node);
                    ReadBorder(pBdr.RightBorder, "pbdr.right", node);
                    ReadBorder(pBdr.BetweenBorder, "pbdr.between", node);
                    ReadBorder(pBdr.BarBorder, "pbdr.bar", node);
                }

                var numProps = pProps.NumberingProperties;
                if (numProps != null)
                {
                    if (numProps.NumberingId?.Val?.Value != null)
                    {
                        var numIdVal = numProps.NumberingId.Val.Value;
                        node.Format["numid"] = numIdVal;
                        var ilvlVal = numProps.NumberingLevelReference?.Val?.Value ?? 0;
                        node.Format["numlevel"] = ilvlVal;
                        var numFmt = GetNumberingFormat(numIdVal, ilvlVal);
                        node.Format["numFmt"] = numFmt;
                        node.Format["listStyle"] = numFmt.ToLowerInvariant() == "bullet" ? "bullet" : "ordered";
                        var start = GetStartValue(numIdVal, ilvlVal);
                        if (start != null)
                            node.Format["start"] = start.Value;
                    }
                }
            }

            // First-run formatting on the paragraph node (like PPTX does for shapes)
            var firstRun = para.Elements<Run>().FirstOrDefault(r => r.GetFirstChild<Text>() != null);
            if (firstRun?.RunProperties != null)
            {
                var rp = firstRun.RunProperties;
                var pFont = rp.RunFonts?.Ascii?.Value;
                if (pFont != null && !node.Format.ContainsKey("font")) node.Format["font"] = pFont;
                if (rp.FontSize?.Val?.Value != null && !node.Format.ContainsKey("size"))
                    node.Format["size"] = $"{int.Parse(rp.FontSize.Val.Value) / 2.0:0.##}pt";
                if (rp.Bold != null && !node.Format.ContainsKey("bold")) node.Format["bold"] = true;
                if (rp.Italic != null && !node.Format.ContainsKey("italic")) node.Format["italic"] = true;
                if (rp.Color?.Val?.Value != null && !node.Format.ContainsKey("color"))
                    node.Format["color"] = ParseHelpers.FormatHexColor(rp.Color.Val.Value);
                else if (rp.Color?.ThemeColor?.HasValue == true && !node.Format.ContainsKey("color"))
                    node.Format["color"] = rp.Color.ThemeColor.InnerText;
                if (rp.Underline?.Val != null && !node.Format.ContainsKey("underline"))
                    node.Format["underline"] = rp.Underline.Val.InnerText;
                if (rp.Strike != null && !node.Format.ContainsKey("strike"))
                    node.Format["strike"] = true;
                if (rp.Highlight?.Val != null && !node.Format.ContainsKey("highlight"))
                    node.Format["highlight"] = rp.Highlight.Val.InnerText;
            }

            if (depth > 0)
            {
                int runIdx = 0;
                foreach (var run in GetAllRuns(para))
                {
                    node.Children.Add(ElementToNode(run, $"{path}/r[{runIdx + 1}]", depth - 1));
                    runIdx++;
                }
            }
        }
        else if (element is Run run)
        {
            node.Type = "run";
            node.Text = GetRunText(run);
            var font = GetRunFont(run);
            if (font != null) node.Format["font"] = font;
            var size = GetRunFontSize(run);
            if (size != null) node.Format["size"] = size;
            if (run.RunProperties?.Bold != null) node.Format["bold"] = true;
            if (run.RunProperties?.Italic != null) node.Format["italic"] = true;
            if (run.RunProperties?.Color?.Val?.Value != null) node.Format["color"] = ParseHelpers.FormatHexColor(run.RunProperties.Color.Val.Value);
            else if (run.RunProperties?.Color?.ThemeColor?.HasValue == true) node.Format["color"] = run.RunProperties.Color.ThemeColor.InnerText;
            if (run.RunProperties?.Underline?.Val != null) node.Format["underline"] = run.RunProperties.Underline.Val.InnerText;
            if (run.RunProperties?.Strike != null) node.Format["strike"] = true;
            if (run.RunProperties?.Highlight?.Val != null) node.Format["highlight"] = run.RunProperties.Highlight.Val.InnerText;
            if (run.RunProperties?.Caps != null) node.Format["caps"] = true;
            if (run.RunProperties?.SmallCaps != null) node.Format["smallcaps"] = true;
            if (run.RunProperties?.DoubleStrike != null) node.Format["dstrike"] = true;
            if (run.RunProperties?.Vanish != null) node.Format["vanish"] = true;
            if (run.RunProperties?.Outline != null) node.Format["outline"] = true;
            if (run.RunProperties?.Shadow != null) node.Format["shadow"] = true;
            if (run.RunProperties?.Emboss != null) node.Format["emboss"] = true;
            if (run.RunProperties?.Imprint != null) node.Format["imprint"] = true;
            if (run.RunProperties?.NoProof != null) node.Format["noproof"] = true;
            if (run.RunProperties?.RightToLeftText != null) node.Format["rtl"] = true;
            if (run.RunProperties?.VerticalTextAlignment?.Val?.Value == VerticalPositionValues.Superscript)
                node.Format["superscript"] = true;
            if (run.RunProperties?.VerticalTextAlignment?.Val?.Value == VerticalPositionValues.Subscript)
                node.Format["subscript"] = true;
            if (run.RunProperties?.Shading?.Fill?.Value != null)
            {
                node.Format["shading"] = ParseHelpers.FormatHexColor(run.RunProperties.Shading.Fill.Value);
            }
            // w14 text effects
            ReadW14TextEffects(run.RunProperties, node);
            // Image properties if run contains a Drawing
            var runDrawing = run.GetFirstChild<Drawing>();
            if (runDrawing != null)
            {
                node.Type = "picture";
                var docProps = runDrawing.Descendants<DW.DocProperties>().FirstOrDefault();
                if (docProps?.Description?.Value != null) node.Format["alt"] = docProps.Description.Value;
                var extent = runDrawing.Descendants<DW.Extent>().FirstOrDefault();
                if (extent?.Cx != null) node.Format["width"] = $"{extent.Cx.Value / 360000.0:F1}cm";
                if (extent?.Cy != null) node.Format["height"] = $"{extent.Cy.Value / 360000.0:F1}cm";
            }
            if (run.Parent is Hyperlink hlParent && hlParent.Id?.Value != null)
            {
                try
                {
                    var rel = _doc.MainDocumentPart?.HyperlinkRelationships.FirstOrDefault(r => r.Id == hlParent.Id.Value);
                    if (rel != null) node.Format["link"] = rel.Uri.ToString();
                }
                catch { }
            }
        }
        else if (element is Hyperlink hyperlink)
        {
            node.Type = "hyperlink";
            node.Text = string.Concat(hyperlink.Descendants<Text>().Select(t => t.Text));
            var relId = hyperlink.Id?.Value;
            if (relId != null)
            {
                try
                {
                    var rel = _doc.MainDocumentPart?.HyperlinkRelationships
                        .FirstOrDefault(r => r.Id == relId);
                    if (rel != null) node.Format["link"] = rel.Uri.ToString();
                }
                catch { }
            }
            // Read run formatting from the first run inside the hyperlink
            var hlRun = hyperlink.Elements<Run>().FirstOrDefault(r => r.GetFirstChild<Text>() != null);
            if (hlRun?.RunProperties != null)
            {
                var rp = hlRun.RunProperties;
                if (rp.RunFonts?.Ascii?.Value != null) node.Format["font"] = rp.RunFonts.Ascii.Value;
                if (rp.FontSize?.Val?.Value != null)
                    node.Format["size"] = $"{int.Parse(rp.FontSize.Val.Value) / 2.0:0.##}pt";
                if (rp.Bold != null) node.Format["bold"] = true;
                if (rp.Italic != null) node.Format["italic"] = true;
                if (rp.Color?.Val?.Value != null) node.Format["color"] = ParseHelpers.FormatHexColor(rp.Color.Val.Value);
                else if (rp.Color?.ThemeColor?.HasValue == true) node.Format["color"] = rp.Color.ThemeColor.InnerText;
                if (rp.Underline?.Val != null) node.Format["underline"] = rp.Underline.Val.InnerText;
                if (rp.Strike != null) node.Format["strike"] = true;
                if (rp.Highlight?.Val != null) node.Format["highlight"] = rp.Highlight.Val.InnerText;
            }
        }
        else if (element is Table table)
        {
            node.Type = "table";
            node.ChildCount = table.Elements<TableRow>().Count();
            var firstRow = table.Elements<TableRow>().FirstOrDefault();
            // Use grid column count (from TableGrid) instead of cell count for accurate column reporting
            var gridColCount = table.GetFirstChild<TableGrid>()?.Elements<GridColumn>().Count();
            node.Format["cols"] = gridColCount ?? firstRow?.Elements<TableCell>().Count() ?? 0;

            var tp = table.GetFirstChild<TableProperties>();
            if (tp != null)
            {
                // Table style
                if (tp.TableStyle?.Val?.Value != null)
                    node.Format["style"] = tp.TableStyle.Val.Value;
                // Table borders
                var tblBorders = tp.TableBorders;
                if (tblBorders != null)
                {
                    ReadBorder(tblBorders.TopBorder, "border.top", node);
                    ReadBorder(tblBorders.BottomBorder, "border.bottom", node);
                    ReadBorder(tblBorders.LeftBorder, "border.left", node);
                    ReadBorder(tblBorders.RightBorder, "border.right", node);
                    ReadBorder(tblBorders.InsideHorizontalBorder, "border.insideH", node);
                    ReadBorder(tblBorders.InsideVerticalBorder, "border.insideV", node);
                }
                // Table width
                if (tp.TableWidth?.Width?.Value != null)
                {
                    var wType = tp.TableWidth.Type?.Value;
                    node.Format["width"] = wType == TableWidthUnitValues.Pct
                        ? (int.Parse(tp.TableWidth.Width.Value) / 50) + "%"
                        : tp.TableWidth.Width.Value;
                }
                // Alignment
                if (tp.TableJustification?.Val?.Value != null)
                    node.Format["alignment"] = tp.TableJustification.Val.InnerText;
                // Indent
                if (tp.TableIndentation?.Width?.Value != null)
                    node.Format["indent"] = tp.TableIndentation.Width.Value;
                // Cell spacing
                if (tp.TableCellSpacing?.Width?.Value != null)
                    node.Format["cellSpacing"] = tp.TableCellSpacing.Width.Value;
                // Layout
                if (tp.TableLayout?.Type?.Value != null)
                    node.Format["layout"] = tp.TableLayout.Type.Value == TableLayoutValues.Fixed ? "fixed" : "auto";
                // Default cell margin (padding)
                var dcm = tp.TableCellMarginDefault;
                if (dcm?.TopMargin?.Width?.Value != null)
                    node.Format["padding.top"] = dcm.TopMargin.Width.Value;
                if (dcm?.BottomMargin?.Width?.Value != null)
                    node.Format["padding.bottom"] = dcm.BottomMargin.Width.Value;
                if (dcm?.TableCellLeftMargin?.Width?.Value != null)
                    node.Format["padding.left"] = dcm.TableCellLeftMargin.Width.Value;
                if (dcm?.TableCellRightMargin?.Width?.Value != null)
                    node.Format["padding.right"] = dcm.TableCellRightMargin.Width.Value;
            }

            // Column widths from grid
            var gridCols = table.GetFirstChild<TableGrid>()?.Elements<GridColumn>().ToList();
            if (gridCols != null && gridCols.Count > 0)
                node.Format["colWidths"] = string.Join(",", gridCols.Select(g => g.Width?.Value ?? "0"));

            if (depth > 0)
            {
                int rowIdx = 0;
                foreach (var row in table.Elements<TableRow>())
                {
                    var rowNode = new DocumentNode
                    {
                        Path = $"{path}/tr[{rowIdx + 1}]",
                        Type = "row",
                        ChildCount = row.Elements<TableCell>().Count()
                    };
                    ReadRowProps(row, rowNode);
                    if (depth > 1)
                    {
                        int cellIdx = 0;
                        foreach (var cell in row.Elements<TableCell>())
                        {
                            var cellNode = new DocumentNode
                            {
                                Path = $"{path}/tr[{rowIdx + 1}]/tc[{cellIdx + 1}]",
                                Type = "cell",
                                Text = string.Join("", cell.Descendants<Text>().Select(t => t.Text)),
                                ChildCount = cell.Elements<Paragraph>().Count()
                            };
                            ReadCellProps(cell, cellNode);
                            if (depth > 2)
                            {
                                int pIdx = 0;
                                foreach (var cellPara in cell.Elements<Paragraph>())
                                {
                                    cellNode.Children.Add(ElementToNode(cellPara, $"{path}/tr[{rowIdx + 1}]/tc[{cellIdx + 1}]/p[{pIdx + 1}]", depth - 3));
                                    pIdx++;
                                }
                            }
                            rowNode.Children.Add(cellNode);
                            cellIdx++;
                        }
                    }
                    node.Children.Add(rowNode);
                    rowIdx++;
                }
            }
        }
        else if (element is TableCell directCell)
        {
            node.Type = "cell";
            node.Text = string.Join("", directCell.Descendants<Text>().Select(t => t.Text));
            node.ChildCount = directCell.Elements<Paragraph>().Count();
            ReadCellProps(directCell, node);
            if (depth > 0)
            {
                int pIdx = 0;
                foreach (var cellPara in directCell.Elements<Paragraph>())
                {
                    node.Children.Add(ElementToNode(cellPara, $"{path}/p[{pIdx + 1}]", depth - 1));
                    pIdx++;
                }
            }
        }
        else if (element is TableRow directRow)
        {
            node.Type = "row";
            node.ChildCount = directRow.Elements<TableCell>().Count();
            ReadRowProps(directRow, node);
        }
        else if (element is SdtBlock sdtBlockNode)
        {
            node.Type = "sdt";
            var sdtProps = sdtBlockNode.SdtProperties;
            if (sdtProps != null)
            {
                var alias = sdtProps.GetFirstChild<SdtAlias>();
                if (alias?.Val?.Value != null) node.Format["alias"] = alias.Val.Value;
                var tagEl = sdtProps.GetFirstChild<Tag>();
                if (tagEl?.Val?.Value != null) node.Format["tag"] = tagEl.Val.Value;
                var lockEl = sdtProps.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Lock>();
                if (lockEl?.Val?.Value != null) node.Format["lock"] = lockEl.Val.InnerText;
                var sdtId = sdtProps.GetFirstChild<SdtId>();
                if (sdtId?.Val?.Value != null) node.Format["id"] = sdtId.Val.Value;

                // Determine SDT type (check specific types first, text last as fallback)
                if (sdtProps.GetFirstChild<SdtContentDropDownList>() != null) node.Format["sdtType"] = "dropdown";
                else if (sdtProps.GetFirstChild<SdtContentComboBox>() != null) node.Format["sdtType"] = "combobox";
                else if (sdtProps.GetFirstChild<SdtContentDate>() != null) node.Format["sdtType"] = "date";
                else if (sdtProps.GetFirstChild<SdtContentText>() != null) node.Format["sdtType"] = "text";
                else node.Format["sdtType"] = "richtext";

                // Read dropdown/combobox items
                var ddl = sdtProps.GetFirstChild<SdtContentDropDownList>();
                var combo = sdtProps.GetFirstChild<SdtContentComboBox>();
                var listItems = ddl?.Elements<ListItem>() ?? combo?.Elements<ListItem>();
                if (listItems != null)
                {
                    var items = listItems.Select(li => li.DisplayText?.Value ?? li.Value?.Value ?? "").ToList();
                    if (items.Count > 0) node.Format["items"] = string.Join(",", items);
                }
            }
            node.Text = string.Concat(sdtBlockNode.Descendants<Text>().Select(t => t.Text));
            var sdtContent = sdtBlockNode.SdtContentBlock;
            node.ChildCount = sdtContent?.ChildElements.Count ?? 0;
        }
        else if (element is SdtRun sdtRunNode)
        {
            node.Type = "sdt";
            var sdtProps = sdtRunNode.SdtProperties;
            if (sdtProps != null)
            {
                var alias = sdtProps.GetFirstChild<SdtAlias>();
                if (alias?.Val?.Value != null) node.Format["alias"] = alias.Val.Value;
                var tagEl = sdtProps.GetFirstChild<Tag>();
                if (tagEl?.Val?.Value != null) node.Format["tag"] = tagEl.Val.Value;
                var lockEl = sdtProps.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Lock>();
                if (lockEl?.Val?.Value != null) node.Format["lock"] = lockEl.Val.InnerText;
                var sdtId = sdtProps.GetFirstChild<SdtId>();
                if (sdtId?.Val?.Value != null) node.Format["id"] = sdtId.Val.Value;

                if (sdtProps.GetFirstChild<SdtContentDropDownList>() != null) node.Format["sdtType"] = "dropdown";
                else if (sdtProps.GetFirstChild<SdtContentComboBox>() != null) node.Format["sdtType"] = "combobox";
                else if (sdtProps.GetFirstChild<SdtContentDate>() != null) node.Format["sdtType"] = "date";
                else if (sdtProps.GetFirstChild<SdtContentText>() != null) node.Format["sdtType"] = "text";
                else node.Format["sdtType"] = "richtext";

                var ddl = sdtProps.GetFirstChild<SdtContentDropDownList>();
                var combo = sdtProps.GetFirstChild<SdtContentComboBox>();
                var listItems = ddl?.Elements<ListItem>() ?? combo?.Elements<ListItem>();
                if (listItems != null)
                {
                    var items = listItems.Select(li => li.DisplayText?.Value ?? li.Value?.Value ?? "").ToList();
                    if (items.Count > 0) node.Format["items"] = string.Join(",", items);
                }
            }
            node.Text = string.Concat(sdtRunNode.Descendants<Text>().Select(t => t.Text));
        }
        else if (element.LocalName == "oMathPara" || element is M.Paragraph)
        {
            node.Type = "equation";
            node.Format["mode"] = "display";
            // Extract LaTeX via FormulaParser
            var oMath = element.Descendants<M.OfficeMath>().FirstOrDefault();
            if (oMath != null)
            {
                try { node.Text = Core.FormulaParser.ToLatex(oMath); }
                catch { node.Text = element.InnerText; }
            }
            else
            {
                node.Text = element.InnerText;
            }
        }
        else if (element is M.OfficeMath inlineMath)
        {
            node.Type = "equation";
            node.Format["mode"] = "inline";
            try { node.Text = Core.FormulaParser.ToLatex(inlineMath); }
            catch { node.Text = element.InnerText; }
        }
        else
        {
            // Generic fallback: collect XML attributes and child val patterns
            foreach (var attr in element.GetAttributes())
                node.Format[attr.LocalName] = attr.Value;
            foreach (var child in element.ChildElements)
            {
                if (child.ChildElements.Count == 0)
                {
                    foreach (var attr in child.GetAttributes())
                    {
                        if (attr.LocalName.Equals("val", StringComparison.OrdinalIgnoreCase))
                        {
                            node.Format[child.LocalName] = attr.Value;
                            break;
                        }
                    }
                }
            }

            var innerText = element.InnerText;
            if (!string.IsNullOrEmpty(innerText))
                node.Text = innerText.Length > 200 ? innerText[..200] + "..." : innerText;
            if (string.IsNullOrEmpty(innerText))
            {
                var outerXml = element.OuterXml;
                node.Preview = outerXml.Length > 200 ? outerXml[..200] + "..." : outerXml;
            }

            node.ChildCount = element.ChildElements.Count;
            if (depth > 0)
            {
                var typeCounters = new Dictionary<string, int>();
                foreach (var child in element.ChildElements)
                {
                    var name = child.LocalName;
                    typeCounters.TryGetValue(name, out int idx);
                    node.Children.Add(ElementToNode(child, $"{path}/{name}[{idx + 1}]", depth - 1));
                    typeCounters[name] = idx + 1;
                }
            }
        }

        return node;
    }

    private static void ReadRowProps(TableRow row, DocumentNode node)
    {
        var trPr = row.TableRowProperties;
        if (trPr == null) return;
        var rh = trPr.GetFirstChild<TableRowHeight>();
        if (rh?.Val?.Value != null)
        {
            node.Format["height"] = rh.Val.Value;
            if (rh.HeightType?.Value == HeightRuleValues.Exact)
                node.Format["height.rule"] = "exact";
        }
        if (trPr.GetFirstChild<TableHeader>() != null)
            node.Format["header"] = true;
    }

    private static void ReadCellProps(TableCell cell, DocumentNode node)
    {
        var tcPr = cell.TableCellProperties;
        if (tcPr != null)
        {
            // Borders (including diagonal — like POI CTTcBorders)
            var cb = tcPr.TableCellBorders;
            if (cb != null)
            {
                ReadBorder(cb.TopBorder, "border.top", node);
                ReadBorder(cb.BottomBorder, "border.bottom", node);
                ReadBorder(cb.LeftBorder, "border.left", node);
                ReadBorder(cb.RightBorder, "border.right", node);
                ReadBorder(cb.TopLeftToBottomRightCellBorder, "border.tl2br", node);
                ReadBorder(cb.TopRightToBottomLeftCellBorder, "border.tr2bl", node);
            }
            // Shading — check for gradient (w14:gradFill in mc:AlternateContent) first
            var mcNs = "http://schemas.openxmlformats.org/markup-compatibility/2006";
            var gradAc = tcPr.ChildElements
                .FirstOrDefault(e => e.LocalName == "AlternateContent" && e.NamespaceUri == mcNs);
            if (gradAc != null && gradAc.InnerXml.Contains("gradFill"))
            {
                // Parse gradient colors and angle from w14:gradFill XML
                var colors = new List<string>();
                foreach (var match in System.Text.RegularExpressions.Regex.Matches(
                    gradAc.InnerXml, @"val=""([0-9A-Fa-f]{6})"""))
                {
                    colors.Add(((System.Text.RegularExpressions.Match)match).Groups[1].Value);
                }
                var angleMatch = System.Text.RegularExpressions.Regex.Match(
                    gradAc.InnerXml, @"ang=""(\d+)""");
                var angle = angleMatch.Success ? int.Parse(angleMatch.Groups[1].Value) / 60000.0 : 0.0;
                var angleStr = angle % 1 == 0 ? $"{(int)angle}" : $"{angle:0.##}";
                if (colors.Count >= 2)
                {
                    node.Format["fill"] = $"gradient;{ParseHelpers.FormatHexColor(colors[0])};{ParseHelpers.FormatHexColor(colors[1])};{angleStr}";
                }
                else if (colors.Count == 1)
                {
                    node.Format["fill"] = ParseHelpers.FormatHexColor(colors[0]);
                }
            }
            else
            {
                var shd = tcPr.Shading;
                if (shd?.Fill?.Value != null)
                {
                    node.Format["fill"] = ParseHelpers.FormatHexColor(shd.Fill.Value);
                }
            }
            // Width
            if (tcPr.TableCellWidth?.Width?.Value != null)
                node.Format["width"] = tcPr.TableCellWidth.Width.Value;
            // Vertical alignment
            if (tcPr.TableCellVerticalAlignment?.Val?.Value != null)
                node.Format["valign"] = tcPr.TableCellVerticalAlignment.Val.InnerText;
            // Vertical merge
            if (tcPr.VerticalMerge != null)
                node.Format["vmerge"] = tcPr.VerticalMerge.Val?.Value == MergedCellValues.Restart ? "restart" : "continue";
            // Grid span
            if (tcPr.GridSpan?.Val?.Value != null && tcPr.GridSpan.Val.Value > 1)
                node.Format["gridSpan"] = tcPr.GridSpan.Val.Value;
            // Cell padding/margins
            var mar = tcPr.TableCellMargin;
            if (mar != null)
            {
                if (mar.TopMargin?.Width?.Value != null) node.Format["padding.top"] = mar.TopMargin.Width.Value;
                if (mar.BottomMargin?.Width?.Value != null) node.Format["padding.bottom"] = mar.BottomMargin.Width.Value;
                if (mar.LeftMargin?.Width?.Value != null) node.Format["padding.left"] = mar.LeftMargin.Width.Value;
                if (mar.RightMargin?.Width?.Value != null) node.Format["padding.right"] = mar.RightMargin.Width.Value;
            }
            // Text direction
            if (tcPr.TextDirection?.Val?.Value != null)
                node.Format["textDirection"] = tcPr.TextDirection.Val.InnerText;
            // No wrap
            if (tcPr.NoWrap != null)
                node.Format["nowrap"] = true;
        }
        // Alignment from first paragraph
        var firstPara = cell.Elements<Paragraph>().FirstOrDefault();
        var just = firstPara?.ParagraphProperties?.Justification?.Val;
        if (just != null)
            node.Format["alignment"] = just.InnerText;
        // Run-level formatting from first run (mirrors PPTX table cell behavior)
        var firstRun = cell.Descendants<Run>().FirstOrDefault();
        if (firstRun?.RunProperties != null)
        {
            var rPr = firstRun.RunProperties;
            if (rPr.RunFonts?.Ascii?.Value != null) node.Format["font"] = rPr.RunFonts.Ascii.Value;
            if (rPr.FontSize?.Val?.Value != null) node.Format["size"] = $"{int.Parse(rPr.FontSize.Val.Value) / 2.0:0.##}pt";
            if (rPr.Bold != null) node.Format["bold"] = true;
            if (rPr.Italic != null) node.Format["italic"] = true;
            if (rPr.Color?.Val?.Value != null) node.Format["color"] = ParseHelpers.FormatHexColor(rPr.Color.Val.Value);
            else if (rPr.Color?.ThemeColor?.HasValue == true) node.Format["color"] = rPr.Color.ThemeColor.InnerText;
            if (rPr.Underline?.Val != null) node.Format["underline"] = rPr.Underline.Val.InnerText;
            if (rPr.Strike != null) node.Format["strike"] = true;
            if (rPr.Highlight?.Val != null) node.Format["highlight"] = rPr.Highlight.Val.InnerText;
        }
    }

    private static void ReadBorder(BorderType? border, string key, DocumentNode node)
    {
        if (border?.Val == null) return;
        var style = border.Val?.InnerText ?? "none";
        var size = border.Size?.Value ?? 0u;
        var color = border.Color?.Value;
        var space = border.Space?.Value ?? 0u;
        var parts = new List<string> { style };
        if (size > 0 || color != null || space > 0) parts.Add(size.ToString());
        if (color != null || space > 0) parts.Add(color is not null ? ParseHelpers.FormatHexColor(color) : "auto");
        if (space > 0) parts.Add(space.ToString());
        node.Format[key] = string.Join(";", parts);
    }
}
