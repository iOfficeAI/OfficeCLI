using System.IO.Compression;
using System.Text;
using System.Text.Json.Nodes;
using System.Xml.Linq;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class HwpxHandler
{
    // ==================== Query Layer ====================

    public DocumentNode Get(string path, int depth = 1)
    {
        if (string.IsNullOrEmpty(path))
            throw new ArgumentException("Path cannot be empty");

        if (path == "/")
            return GetRootNode(depth);

        var element = ResolvePath(path);
        return BuildDocumentNode(element, path, depth);
    }

    private DocumentNode GetRootNode(int depth)
    {
        var node = new DocumentNode
        {
            Path = "/",
            Type = "hwpx-document",
            ChildCount = _doc.Sections.Count,
        };

        // Document metadata
        node.Format["sections"] = _doc.Sections.Count;
        node.Format["hasHeader"] = _doc.Header != null;

        if (depth > 0)
        {
            foreach (var sec in _doc.Sections)
            {
                var secNode = new DocumentNode
                {
                    Path = $"/section[{sec.Index + 1}]",
                    Type = "section",
                    ChildCount = sec.Root.Elements().Count(),
                    Preview = $"Section {sec.Index + 1}: {sec.Paragraphs.Count} paragraphs, {sec.Tables.Count} tables"
                };

                if (depth > 1)
                {
                    PopulateSectionChildren(secNode, sec, depth - 1);
                }

                node.Children.Add(secNode);
            }
        }

        return node;
    }

    private DocumentNode BuildDocumentNode(XElement element, string path, int depth)
    {
        var localName = element.Name.LocalName;

        return localName switch
        {
            "p" => BuildParagraphNode(element, path, depth),
            "tbl" => BuildTableNode(element, path, depth),
            "tr" => BuildTableRowNode(element, path, depth),
            "tc" => BuildTableCellNode(element, path, depth),
            "run" => BuildRunNode(element, path),
            "sec" => BuildSectionNode(element, path, depth),
            _ => BuildGenericNode(element, path, depth)
        };
    }

    private DocumentNode BuildParagraphNode(XElement para, string path, int depth)
    {
        var text = HwpxKorean.Normalize(ExtractParagraphText(para));
        var styleInfo = GetParagraphStyleInfo(para);

        var node = new DocumentNode
        {
            Path = path,
            Type = !string.IsNullOrEmpty(styleInfo.HeadingLevel) 
                ? $"heading{styleInfo.HeadingLevel}" 
                : "paragraph",
            Text = text,
            Preview = text.Length > 100 ? text[..100] + "…" : text,
            ChildCount = para.Elements(HwpxNs.Hp + "run").Count(),
        };

        // Format properties
        node.Format["alignment"] = styleInfo.Alignment;
        if (styleInfo.HeadingLevel != null)
            node.Format["headingLevel"] = int.Parse(styleInfo.HeadingLevel);
        
        var styleIdRef = para.Attribute("styleIDRef")?.Value;
        if (styleIdRef != null)
            node.Format["styleIDRef"] = styleIdRef;

        var paraPrIdRef = para.Attribute("paraPrIDRef")?.Value;
        if (paraPrIdRef != null)
            node.Format["paraPrIDRef"] = paraPrIdRef;

        // Populate children (runs) if depth allows
        if (depth > 1)
        {
            int runIdx = 0;
            foreach (var run in para.Elements(HwpxNs.Hp + "run"))
            {
                runIdx++;
                var runPath = $"{path}/run[{runIdx}]";
                node.Children.Add(BuildRunNode(run, runPath));
            }
        }

        return node;
    }

    private DocumentNode BuildRunNode(XElement run, string path)
    {
        var text = string.Join("", run.Elements(HwpxNs.Hp + "t").Select(t => t.Value));
        text = HwpxKorean.Normalize(text);

        var node = new DocumentNode
        {
            Path = path,
            Type = "run",
            Text = text,
        };

        var charPrIdRef = run.Attribute("charPrIDRef")?.Value;
        if (charPrIdRef != null)
        {
            node.Format["charPrIDRef"] = charPrIdRef;

            // Look up character properties from header.xml
            if (_doc.Header != null)
            {
                var charPr = _doc.Header.Descendants(HwpxNs.Hh + "charPr")
                    .FirstOrDefault(cp => cp.Attribute("id")?.Value == charPrIdRef);
                if (charPr != null)
                {
                    var height = (int?)charPr.Attribute("height");
                    if (height.HasValue)
                        node.Format["fontSize"] = $"{height.Value / 100.0}pt";

                    var textColor = charPr.Attribute("textColor")?.Value;
                    if (textColor != null)
                        node.Format["color"] = textColor;

                    if (charPr.Element(HwpxNs.Hh + "bold") != null)
                        node.Format["bold"] = true;
                    if (charPr.Element(HwpxNs.Hh + "italic") != null)
                        node.Format["italic"] = true;

                    var fontRef = charPr.Element(HwpxNs.Hh + "fontRef");
                    if (fontRef != null)
                    {
                        node.Format["fontHangul"] = fontRef.Attribute("hangul")?.Value;
                        node.Format["fontLatin"] = fontRef.Attribute("latin")?.Value;
                    }
                }
            }
        }

        return node;
    }

    private DocumentNode BuildTableNode(XElement tbl, string path, int depth)
    {
        var rowCnt = (int?)tbl.Attribute("rowCnt") ?? 0;
        var colCnt = (int?)tbl.Attribute("colCnt") ?? 0;

        var node = new DocumentNode
        {
            Path = path,
            Type = "table",
            Preview = $"Table {rowCnt}×{colCnt}",
            ChildCount = rowCnt,
        };

        node.Format["rowCount"] = rowCnt;
        node.Format["colCount"] = colCnt;

        var borderFill = tbl.Attribute("borderFillIDRef")?.Value;
        if (borderFill != null)
            node.Format["borderFillIDRef"] = borderFill;

        if (depth > 1)
        {
            int trIdx = 0;
            foreach (var tr in tbl.Elements(HwpxNs.Hp + "tr"))
            {
                trIdx++;
                var trPath = $"{path}/tr[{trIdx}]";
                node.Children.Add(BuildTableRowNode(tr, trPath, depth - 1));
            }
        }

        return node;
    }

    private DocumentNode BuildTableRowNode(XElement tr, string path, int depth)
    {
        var cells = tr.Elements(HwpxNs.Hp + "tc").ToList();
        var node = new DocumentNode
        {
            Path = path,
            Type = "tableRow",
            ChildCount = cells.Count,
        };

        if (depth > 1)
        {
            int tcIdx = 0;
            foreach (var tc in cells)
            {
                tcIdx++;
                var tcPath = $"{path}/tc[{tcIdx}]";
                node.Children.Add(BuildTableCellNode(tc, tcPath, depth - 1));
            }
        }

        return node;
    }

    private DocumentNode BuildTableCellNode(XElement tc, string path, int depth)
    {
        // Extract cell address (dual-format support)
        var (row, col, rowSpan, colSpan) = GetCellAddr(tc);

        var node = new DocumentNode
        {
            Path = path,
            Type = "tableCell",
        };

        node.Format["row"] = row;
        node.Format["col"] = col;
        node.Format["rowSpan"] = rowSpan;
        node.Format["colSpan"] = colSpan;

        // Extract cell text from subList paragraphs
        var subList = tc.Element(HwpxNs.Hp + "subList");
        if (subList != null)
        {
            var cellText = new StringBuilder();
            foreach (var p in subList.Elements(HwpxNs.Hp + "p"))
            {
                var pText = HwpxKorean.Normalize(ExtractParagraphText(p));
                if (cellText.Length > 0 && !string.IsNullOrEmpty(pText))
                    cellText.Append('\n');
                cellText.Append(pText);
            }
            node.Text = cellText.ToString();
            node.Preview = node.Text.Length > 50 ? node.Text[..50] + "…" : node.Text;
            node.ChildCount = subList.Elements(HwpxNs.Hp + "p").Count();
        }

        return node;
    }

    private DocumentNode BuildSectionNode(XElement sec, string path, int depth)
    {
        var section = _doc.Sections.FirstOrDefault(s => s.Root == sec);
        var node = new DocumentNode
        {
            Path = path,
            Type = "section",
            ChildCount = sec.Elements().Count(),
        };

        if (section != null)
        {
            node.Preview = $"Section {section.Index + 1}: {section.Paragraphs.Count}p, {section.Tables.Count}tbl";
        }

        // Section properties
        var secPr = sec.Descendants(HwpxNs.Hp + "secPr").FirstOrDefault();
        if (secPr != null)
        {
            node.Format["textDirection"] = secPr.Attribute("textDirection")?.Value;
            var pagePr = secPr.Element(HwpxNs.Hp + "pagePr");
            if (pagePr != null)
            {
                node.Format["pageWidth"] = (int?)pagePr.Attribute("width");
                node.Format["pageHeight"] = (int?)pagePr.Attribute("height");
                node.Format["landscape"] = pagePr.Attribute("landscape")?.Value;
            }
        }

        if (depth > 0 && section != null)
        {
            PopulateSectionChildren(node, section, depth);
        }

        return node;
    }

    private void PopulateSectionChildren(DocumentNode node, HwpxSection section, int depth)
    {
        int pIdx = 0, tblIdx = 0;
        foreach (var child in section.Root.Elements())
        {
            var localName = child.Name.LocalName;
            if (localName == "p")
            {
                pIdx++;
                var childPath = $"{node.Path}/p[{pIdx}]";
                node.Children.Add(BuildParagraphNode(child, childPath, depth - 1));
            }
            else if (localName == "tbl")
            {
                tblIdx++;
                var childPath = $"{node.Path}/tbl[{tblIdx}]";
                node.Children.Add(BuildTableNode(child, childPath, depth - 1));
            }
        }
    }

    private DocumentNode BuildGenericNode(XElement element, string path, int depth)
    {
        var node = new DocumentNode
        {
            Path = path,
            Type = element.Name.LocalName,
            ChildCount = element.Elements().Count(),
        };

        // Copy all attributes to format
        foreach (var attr in element.Attributes())
        {
            node.Format[attr.Name.LocalName] = attr.Value;
        }

        if (element.HasElements && depth > 1)
        {
            // Per-type counters so paths are resolvable: a[1], b[1], a[2] — not a[1], b[2], a[3]
            var childCounts = new Dictionary<string, int>();
            foreach (var child in element.Elements())
            {
                var localName = child.Name.LocalName;
                childCounts.TryGetValue(localName, out int count);
                childCounts[localName] = ++count;
                var childPath = $"{path}/{MapElementToPathName(localName)}[{count}]";
                node.Children.Add(BuildDocumentNode(child, childPath, depth - 1));
            }
        }
        else if (!element.HasElements)
        {
            node.Text = element.Value;
        }

        return node;
    }

    public List<DocumentNode> Query(string selector)
    {
        if (string.IsNullOrEmpty(selector))
            throw new ArgumentException("Selector cannot be empty");

        var elements = ExecuteSelector(selector);
        return elements.Select(e => BuildDocumentNode(e, BuildPath(e), 1)).ToList();
    }

    private void SaveSection(XElement element)
    {
        // Walk up to find which section this element belongs to
        var current = element;
        while (current != null)
        {
            if (current.Name.Namespace == HwpxNs.Hs && current.Name.LocalName == "sec")
                break;
            current = current.Parent;
        }

        if (current == null)
            throw new InvalidOperationException("Cannot determine section for element");

        var section = _doc.Sections.First(s => s.Root == current);
        var entryName = section.EntryPath;
        // Delete-and-recreate pattern (avoids trailing bytes from SetLength(0))
        // For new sections, entry may not exist yet — just create.
        var entry = _doc.Archive.GetEntry(entryName);
        entry?.Delete();
        var newEntry = _doc.Archive.CreateEntry(entryName, CompressionLevel.Optimal);
        using var stream = newEntry.Open();
        // CRITICAL: Hancom requires single-line (minified) XML without BOM.
        // XDocument.ToString() pretty-prints by default — use DisableFormatting.
        var xmlStr = HwpxPacker.MinifyXml(section.Document.ToString(SaveOptions.DisableFormatting));
        // Restore 2016 namespace for hp10 prefix — Hancom expects this original URI
        xmlStr = HwpxPacker.RestoreOriginalNamespaces(xmlStr);
        // Prepend XML declaration (without BOM)
        xmlStr = "<?xml version='1.0' encoding='UTF-8'?>" + xmlStr;
        var bytes = System.Text.Encoding.UTF8.GetBytes(xmlStr);
        stream.Write(bytes, 0, bytes.Length);
    }
}
