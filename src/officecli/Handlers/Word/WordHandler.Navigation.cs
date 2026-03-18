// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;

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
    {
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

        for (int i = 1; i < segments.Count && current != null; i++)
        {
            var seg = segments[i];
            IEnumerable<OpenXmlElement> children;
            if (current is Body body2 && (seg.Name.ToLowerInvariant() == "p" || seg.Name.ToLowerInvariant() == "tbl"))
            {
                // Flatten sdt containers when navigating body-level paragraphs/tables
                children = seg.Name.ToLowerInvariant() == "p"
                    ? GetBodyElements(body2).OfType<Paragraph>().Cast<OpenXmlElement>()
                    : GetBodyElements(body2).OfType<Table>().Cast<OpenXmlElement>();
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
                    _ => current.ChildElements.Where(e => e.LocalName == seg.Name).Cast<OpenXmlElement>()
                };
            }

            current = seg.Index.HasValue
                ? children.ElementAtOrDefault(seg.Index.Value - 1)
                : children.FirstOrDefault();
        }

        return current;
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
                    node.Format["alignment"] = pProps.Justification.Val.InnerText;
                if (pProps.SpacingBetweenLines != null)
                {
                    if (pProps.SpacingBetweenLines.Before?.Value != null)
                        node.Format["spacebefore"] = pProps.SpacingBetweenLines.Before.Value;
                    if (pProps.SpacingBetweenLines.After?.Value != null)
                        node.Format["spaceafter"] = pProps.SpacingBetweenLines.After.Value;
                    if (pProps.SpacingBetweenLines.Line?.Value != null)
                        node.Format["linespacing"] = pProps.SpacingBetweenLines.Line.Value;
                }
                if (pProps.Indentation?.FirstLine?.Value != null)
                    node.Format["firstlineindent"] = pProps.Indentation.FirstLine.Value;
                if (pProps.Indentation?.Left?.Value != null)
                    node.Format["leftindent"] = pProps.Indentation.Left.Value;
                if (pProps.Indentation?.Right?.Value != null)
                    node.Format["rightindent"] = pProps.Indentation.Right.Value;
                if (pProps.Indentation?.Hanging?.Value != null)
                    node.Format["hangingindent"] = pProps.Indentation.Hanging.Value;
                if (pProps.KeepNext != null)
                    node.Format["keepnext"] = true;
                if (pProps.KeepLines != null)
                    node.Format["keeplines"] = true;
                if (pProps.PageBreakBefore != null)
                    node.Format["pagebreakbefore"] = true;
                if (pProps.WidowControl != null)
                    node.Format["widowcontrol"] = true;
                if (pProps.Shading != null)
                    node.Format["shading"] = pProps.Shading.Fill?.Value ?? pProps.Shading.Color?.Value ?? "";

                var pBdr = pProps.ParagraphBorders;
                if (pBdr != null)
                {
                    ReadBorder(pBdr.TopBorder, "pBdr.top", node);
                    ReadBorder(pBdr.BottomBorder, "pBdr.bottom", node);
                    ReadBorder(pBdr.LeftBorder, "pBdr.left", node);
                    ReadBorder(pBdr.RightBorder, "pBdr.right", node);
                    ReadBorder(pBdr.BetweenBorder, "pBdr.between", node);
                    ReadBorder(pBdr.BarBorder, "pBdr.bar", node);
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
            if (run.RunProperties?.Color?.Val?.Value != null) node.Format["color"] = run.RunProperties.Color.Val.Value;
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
                node.Format["shading"] = run.RunProperties.Shading.Fill.Value;
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
        }
        else if (element is Table table)
        {
            node.Type = "table";
            node.ChildCount = table.Elements<TableRow>().Count();
            var firstRow = table.Elements<TableRow>().FirstOrDefault();
            node.Format["cols"] = firstRow?.Elements<TableCell>().Count() ?? 0;

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
            // Borders
            var cb = tcPr.TableCellBorders;
            if (cb != null)
            {
                ReadBorder(cb.TopBorder, "border.top", node);
                ReadBorder(cb.BottomBorder, "border.bottom", node);
                ReadBorder(cb.LeftBorder, "border.left", node);
                ReadBorder(cb.RightBorder, "border.right", node);
            }
            // Shading
            var shd = tcPr.Shading;
            if (shd?.Fill?.Value != null)
                node.Format["shd"] = shd.Fill.Value;
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
                node.Format["gridspan"] = tcPr.GridSpan.Val.Value;
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
    }

    private static void ReadBorder(BorderType? border, string key, DocumentNode node)
    {
        if (border?.Val == null) return;
        var style = border.Val.InnerText;
        var size = border.Size?.Value ?? 0u;
        var color = border.Color?.Value;
        var space = border.Space?.Value ?? 0u;
        var parts = new List<string> { style };
        if (size > 0 || color != null || space > 0) parts.Add(size.ToString());
        if (color != null || space > 0) parts.Add(color ?? "auto");
        if (space > 0) parts.Add(space.ToString());
        node.Format[key] = string.Join(";", parts);
    }
}
