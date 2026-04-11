using System.Text;
using System.Text.Json.Nodes;
using System.Xml.Linq;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class HwpxHandler
{
    // ==================== View Layer ====================

    public string ViewAsText(int? startLine = null, int? endLine = null,
                              int? maxLines = null, HashSet<string>? cols = null)
    {
        var sb = new StringBuilder();
        int lineNum = 0;
        int emitted = 0;

        foreach (var (section, para, path) in _doc.AllContentInOrder())
        {
            lineNum++;
            if (startLine.HasValue && lineNum < startLine.Value) continue;
            if (endLine.HasValue && lineNum > endLine.Value) break;

            var rawText = ExtractParagraphText(para);
            var text = HwpxKorean.Normalize(rawText);

            if (maxLines.HasValue && emitted >= maxLines.Value)
            {
                sb.AppendLine($"... (more lines)");
                break;
            }

            sb.AppendLine($"{lineNum}. {text}");
            emitted++;
        }

        return sb.ToString().TrimEnd();
    }

    public string ViewAsAnnotated(int? startLine = null, int? endLine = null,
                                   int? maxLines = null, HashSet<string>? cols = null)
    {
        var sb = new StringBuilder();
        int lineNum = 0;
        int emitted = 0;

        foreach (var (section, para, localIdx) in _doc.AllParagraphs())
        {
            lineNum++;
            if (startLine.HasValue && lineNum < startLine.Value) continue;
            if (endLine.HasValue && lineNum > endLine.Value) break;
            if (maxLines.HasValue && emitted >= maxLines.Value)
            {
                int remaining = CountRemainingParagraphs(lineNum);
                if (remaining > 0)
                    sb.AppendLine($"... ({remaining} more lines)");
                break;
            }

            var path = $"/section[{section.Index + 1}]/p[{localIdx + 1}]";
            var styleInfo = GetParagraphStyleInfo(para);
            var runs = ExtractAnnotatedRuns(para);
            var text = string.Join("", runs.Select(r => r.Text));
            text = HwpxKorean.Normalize(text);

            // Build annotation prefix
            var annotations = new List<string>();
            if (!string.IsNullOrEmpty(styleInfo.HeadingLevel))
                annotations.Add($"h{styleInfo.HeadingLevel}");
            if (styleInfo.Alignment != "LEFT")
                annotations.Add(styleInfo.Alignment.ToLowerInvariant());

            var prefix = annotations.Count > 0 ? $"[{string.Join(",", annotations)}] " : "";
            sb.AppendLine($"{lineNum}. {path} {prefix}{text}");
            emitted++;
        }

        return sb.ToString().TrimEnd();
    }

    public string ViewAsOutline()
    {
        var sb = new StringBuilder();

        foreach (var (section, para, localIdx) in _doc.AllParagraphs())
        {
            var styleInfo = GetParagraphStyleInfo(para);
            if (string.IsNullOrEmpty(styleInfo.HeadingLevel)) continue;

            var level = int.Parse(styleInfo.HeadingLevel);
            var indent = new string(' ', (level - 1) * 2);
            var text = HwpxKorean.Normalize(ExtractParagraphText(para));
            var preview = text.Length > 80 ? text[..80] + "…" : text;
            var path = $"/section[{section.Index + 1}]/p[{localIdx + 1}]";

            sb.AppendLine($"{indent}h{level}: {preview} ({path})");
        }

        return sb.Length > 0 ? sb.ToString().TrimEnd() : "(no headings found)";
    }

    public string ViewAsStats()
    {
        int totalParas = 0, totalTables = 0, totalChars = 0, totalWords = 0;
        int totalImages = 0;

        foreach (var sec in _doc.Sections)
        {
            totalParas += sec.Paragraphs.Count;
            totalTables += sec.Tables.Count;
            totalImages += sec.Root.Descendants(HwpxNs.Hp + "img").Count();

            foreach (var p in sec.Paragraphs)
            {
                var text = HwpxKorean.Normalize(ExtractParagraphText(p));
                totalChars += text.Length;
                totalWords += CountWords(text);
            }
        }

        var sb = new StringBuilder();
        sb.AppendLine($"Sections:   {_doc.Sections.Count}");
        sb.AppendLine($"Paragraphs: {totalParas}");
        sb.AppendLine($"Tables:     {totalTables}");
        sb.AppendLine($"Images:     {totalImages}");
        sb.AppendLine($"Characters: {totalChars}");
        sb.AppendLine($"Words:      {totalWords}");

        // Page info — iterate ALL sections for aggregate stats; use first secPr for page size reference
        foreach (var sec in _doc.Sections)
        {
            var secPr = sec.Root.Descendants(HwpxNs.Hp + "secPr").FirstOrDefault();
            var pagePr = secPr?.Element(HwpxNs.Hp + "pagePr");
            if (pagePr != null)
            {
                var width = (int?)pagePr.Attribute("width") ?? 0;
                var height = (int?)pagePr.Attribute("height") ?? 0;
                sb.AppendLine($"Page size:  {FormatHwpUnit(width)} × {FormatHwpUnit(height)}");
                break; // Report first section's page size; add per-section loop if needed
            }
        }

        // Metadata
        var meta = GetMetadata();
        if (meta.TryGetValue("title", out var mTitle) && !string.IsNullOrEmpty(mTitle))
            sb.AppendLine($"Title:      {mTitle}");
        if (meta.TryGetValue("creator", out var mCreator) && !string.IsNullOrEmpty(mCreator))
            sb.AppendLine($"Creator:    {mCreator}");

        return sb.ToString().TrimEnd();
    }

    public JsonNode ViewAsStatsJson()
    {
        int totalParas = 0, totalTables = 0, totalChars = 0, totalWords = 0;
        int totalImages = 0;

        foreach (var sec in _doc.Sections)
        {
            totalParas += sec.Paragraphs.Count;
            totalTables += sec.Tables.Count;
            totalImages += sec.Root.Descendants(HwpxNs.Hp + "img").Count();

            foreach (var p in sec.Paragraphs)
            {
                var text = HwpxKorean.Normalize(ExtractParagraphText(p));
                totalChars += text.Length;
                totalWords += CountWords(text);
            }
        }

        return new JsonObject
        {
            ["sections"] = _doc.Sections.Count,
            ["paragraphs"] = totalParas,
            ["tables"] = totalTables,
            ["images"] = totalImages,
            ["characters"] = totalChars,
            ["words"] = totalWords,
        };
    }

    public JsonNode ViewAsOutlineJson()
    {
        var items = new JsonArray();

        foreach (var (section, para, localIdx) in _doc.AllParagraphs())
        {
            var styleInfo = GetParagraphStyleInfo(para);
            if (string.IsNullOrEmpty(styleInfo.HeadingLevel)) continue;

            var level = int.Parse(styleInfo.HeadingLevel);
            var text = HwpxKorean.Normalize(ExtractParagraphText(para));
            var path = $"/section[{section.Index + 1}]/p[{localIdx + 1}]";

            items.Add(new JsonObject
            {
                ["level"] = level,
                ["text"] = text,
                ["path"] = path,
            });
        }

        return items;
    }

    public JsonNode ViewAsTextJson(int? startLine = null, int? endLine = null,
                                    int? maxLines = null, HashSet<string>? cols = null)
    {
        var lines = new JsonArray();
        int lineNum = 0;
        int emitted = 0;

        foreach (var (section, para, path) in _doc.AllContentInOrder())
        {
            lineNum++;
            if (startLine.HasValue && lineNum < startLine.Value) continue;
            if (endLine.HasValue && lineNum > endLine.Value) break;
            if (maxLines.HasValue && emitted >= maxLines.Value) break;

            var text = HwpxKorean.Normalize(ExtractParagraphText(para));

            lines.Add(new JsonObject
            {
                ["line"] = lineNum,
                ["path"] = path,
                ["text"] = text,
            });
            emitted++;
        }

        return new JsonObject
        {
            ["lines"] = lines,
            ["totalLines"] = lineNum,
        };
    }

    public List<DocumentIssue> ViewAsIssues(string? issueType = null, int? limit = null)
    {
        var issues = new List<DocumentIssue>();
        int issueId = 0;

        // Check for empty paragraphs
        foreach (var (section, para, localIdx) in _doc.AllParagraphs())
        {
            var text = ExtractParagraphText(para);
            if (string.IsNullOrWhiteSpace(text))
            {
                // Skip — empty paragraphs are normal spacing
                continue;
            }

            // Check for PUA characters (corruption indicator)
            if (text.Any(c => c >= '\uE000' && c <= '\uF8FF'))
            {
                issues.Add(new DocumentIssue
                {
                    Id = $"HWPX-{++issueId:D3}",
                    Type = IssueType.Content,
                    Severity = IssueSeverity.Warning,
                    Path = $"/section[{section.Index + 1}]/p[{localIdx + 1}]",
                    Message = "Paragraph contains Private Use Area characters",
                    Context = text[..Math.Min(text.Length, 50)]
                });
            }
        }

        // Check for tables with inconsistent column counts
        foreach (var (section, tbl, tblIdx) in _doc.AllTables())
        {
            var rows = tbl.Elements(HwpxNs.Hp + "tr").ToList();
            if (rows.Count == 0) continue;

            var expectedCols = (int?)tbl.Attribute("colCnt") ?? -1;
            foreach (var (row, rowIdx) in rows.Select((r, i) => (r, i)))
            {
                // Sum colSpan values (handles merged cells); GetCellAddr is defined in this partial class
                var colSpanSum = row.Elements(HwpxNs.Hp + "tc")
                    .Sum(tc => (int?)GetCellAddr(tc).ColSpan ?? 1);
                if (expectedCols >= 0 && colSpanSum != expectedCols)
                {
                    issues.Add(new DocumentIssue
                    {
                        Id = $"HWPX-{++issueId:D3}",
                        Type = IssueType.Structure,
                        Severity = IssueSeverity.Error,
                        Path = $"/section[{section.Index + 1}]/tbl[{tblIdx + 1}]/tr[{rowIdx + 1}]",
                        Message = $"Row colSpan sum {colSpanSum} != expected {expectedCols}",
                        Context = null
                    });
                }
            }
        }

        // Check for missing header.xml
        if (_doc.Header == null)
        {
            issues.Add(new DocumentIssue
            {
                Id = $"HWPX-{++issueId:D3}",
                Type = IssueType.Structure,
                Severity = IssueSeverity.Warning,
                Path = "/",
                Message = "Document missing header.xml (style definitions unavailable)",
                Context = null
            });
        }

        // Filter by type
        if (issueType != null)
        {
            var filterType = Enum.Parse<IssueType>(issueType, ignoreCase: true);
            issues = issues.Where(i => i.Type == filterType).ToList();
        }

        // Apply limit
        if (limit.HasValue)
            issues = issues.Take(limit.Value).ToList();

        return issues;
    }

    // ==================== Forms ====================

    public string ViewAsForms()
    {
        var sb = new StringBuilder();
        int count = 0;
        foreach (var sec in _doc.Sections)
        {
            foreach (var run in sec.Root.Descendants(HwpxNs.Hp + "run"))
            {
                var ctrl = run.Element(HwpxNs.Hp + "ctrl");
                var fieldBegin = ctrl?.Element(HwpxNs.Hp + "fieldBegin");
                if (fieldBegin?.Attribute("type")?.Value != "CLICK_HERE") continue;

                count++;
                var instId = fieldBegin.Attribute("id")?.Value ?? "?";
                // Direction = help text for the field
                var direction = fieldBegin.Descendants(HwpxNs.Hp + "stringParam")
                    .FirstOrDefault(p => p.Attribute("name")?.Value == "Direction")?.Value ?? "";
                // Find display text in the next run
                var nextRun = run.ElementsAfterSelf(HwpxNs.Hp + "run").FirstOrDefault();
                var text = nextRun?.Elements(HwpxNs.Hp + "t").FirstOrDefault()?.Value ?? "";
                var isDefault = text == direction;

                sb.AppendLine($"  [{instId}] \"{text}\"{(isDefault ? " (default)" : "")}");
            }
        }
        sb.Insert(0, $"Form fields (CLICK_HERE): {count}\n");
        return sb.ToString().TrimEnd();
    }

    // ==================== Styles ====================

    public string ViewAsStyles()
    {
        if (_doc.Header?.Root == null) return "(no header.xml)";
        var sb = new StringBuilder();
        var styles = _doc.Header.Root.Descendants(HwpxNs.Hh + "style").ToList();
        sb.AppendLine($"Styles: {styles.Count}");
        foreach (var style in styles)
        {
            var id = style.Attribute("id")?.Value ?? "?";
            var name = style.Attribute("name")?.Value ?? "(unnamed)";
            var engName = style.Attribute("engName")?.Value ?? "";
            var type = style.Attribute("type")?.Value ?? "PARA";
            var charPrId = style.Attribute("charPrIDRef")?.Value ?? "0";
            var paraPrId = style.Attribute("paraPrIDRef")?.Value ?? "0";
            var eng = !string.IsNullOrEmpty(engName) ? $" ({engName})" : "";
            sb.AppendLine($"  [{id}] {name}{eng} [{type}] charPr={charPrId} paraPr={paraPrId}");
        }
        return sb.ToString().TrimEnd();
    }

    // ==================== HTML Preview ====================

    public string ViewAsHtml(int? page = null)
    {
        var sb = new StringBuilder();
        sb.AppendLine("<!DOCTYPE html><html lang=\"ko\"><head><meta charset=\"UTF-8\">");
        sb.AppendLine("<title>HWPX Preview</title>");
        sb.AppendLine("<style>");
        sb.AppendLine(HwpxHtmlCss());
        sb.AppendLine("</style></head><body><div class=\"page\">");

        foreach (var (section, element, path) in _doc.AllContentInOrder())
        {
            switch (element.Name.LocalName)
            {
                case "p":
                    var wrappedTbl = element.Descendants(HwpxNs.Hp + "tbl").FirstOrDefault();
                    if (wrappedTbl != null)
                        sb.Append(TableToHtml(wrappedTbl));
                    else
                        sb.Append(ParagraphToHtml(element));
                    break;
            }
        }

        sb.AppendLine("</div></body></html>");
        return sb.ToString();
    }

    // ==================== HTML Helpers ====================

    private string ParagraphToHtml(XElement p)
    {
        var styleInfo = GetParagraphStyleInfo(p);
        var tag = "p";

        if (!string.IsNullOrEmpty(styleInfo.HeadingLevel))
        {
            var level = Math.Clamp(int.Parse(styleInfo.HeadingLevel), 1, 6);
            tag = $"h{level}";
        }

        var paraCss = GetParaPrCss(p.Attribute("paraPrIDRef")?.Value ?? "0");

        var sb = new StringBuilder();
        sb.Append($"<{tag}");
        if (!string.IsNullOrEmpty(paraCss)) sb.Append($" style=\"{paraCss}\"");
        sb.Append('>');

        foreach (var run in p.Elements(HwpxNs.Hp + "run"))
            sb.Append(RunToHtml(run));

        sb.Append($"</{tag}>");
        return sb.ToString();
    }

    private string RunToHtml(XElement run)
    {
        var sb = new StringBuilder();
        var charPrId = run.Attribute("charPrIDRef")?.Value ?? "0";
        var css = GetCharPrCss(charPrId);
        var charPr = FindCharPr(charPrId);
        var hasBold = charPr?.Element(HwpxNs.Hh + "bold") != null;
        var hasItalic = charPr?.Element(HwpxNs.Hh + "italic") != null;
        var ulEl = charPr?.Element(HwpxNs.Hh + "underline");
        var hasUnderline = ulEl != null && ulEl.Attribute("type")?.Value != "NONE";
        var soEl = charPr?.Element(HwpxNs.Hh + "strikeout");
        var hasStrikeout = soEl != null && soEl.Attribute("shape")?.Value != "NONE";
        var hasSup = charPr?.Element(HwpxNs.Hh + "supscript") != null;
        var hasSub = charPr?.Element(HwpxNs.Hh + "subscript") != null;

        if (!string.IsNullOrEmpty(css)) sb.Append($"<span style=\"{css}\">");
        if (hasBold) sb.Append("<b>");
        if (hasItalic) sb.Append("<i>");
        if (hasUnderline) sb.Append("<u>");
        if (hasStrikeout) sb.Append("<s>");
        if (hasSup) sb.Append("<sup>");
        if (hasSub) sb.Append("<sub>");

        foreach (var child in run.Elements())
        {
            switch (child.Name.LocalName)
            {
                case "t":
                    sb.Append(TextWithMarkpenToHtml(child));
                    break;
                case "lineBreak":
                    sb.Append("<br/>");
                    break;
                case "tab":
                    sb.Append("&emsp;");
                    break;
                case "equation":
                    var script = child.Element(HwpxNs.Hp + "script")?.Value
                        ?? child.Attribute("script")?.Value ?? child.Value;
                    sb.Append($"<span class=\"hwpx-eq\" title=\"{EscapeHtml(script)}\">[{EscapeHtml(script.Trim())}]</span>");
                    break;
                case "pic":
                    sb.Append(PicToHtml(child));
                    break;
            }
        }

        if (hasSub) sb.Append("</sub>");
        if (hasSup) sb.Append("</sup>");
        if (hasStrikeout) sb.Append("</s>");
        if (hasUnderline) sb.Append("</u>");
        if (hasItalic) sb.Append("</i>");
        if (hasBold) sb.Append("</b>");
        if (!string.IsNullOrEmpty(css)) sb.Append("</span>");

        return sb.ToString();
    }

    private static string TextWithMarkpenToHtml(XElement t)
    {
        var sb = new StringBuilder();
        foreach (var node in t.Nodes())
        {
            if (node is System.Xml.Linq.XText text)
                sb.Append(EscapeHtml(text.Value));
            else if (node is XElement el)
            {
                if (el.Name.LocalName == "markpenBegin")
                {
                    var color = el.Attribute("color")?.Value ?? "#FFFF00";
                    sb.Append($"<mark style=\"background:{color}\">");
                }
                else if (el.Name.LocalName == "markpenEnd")
                    sb.Append("</mark>");
            }
        }
        return sb.ToString();
    }

    private string TableToHtml(XElement tbl)
    {
        var sb = new StringBuilder();
        sb.Append("<table>");
        foreach (var tr in tbl.Elements(HwpxNs.Hp + "tr"))
        {
            sb.Append("<tr>");
            foreach (var tc in tr.Elements(HwpxNs.Hp + "tc"))
            {
                var cellSpan = tc.Element(HwpxNs.Hp + "cellSpan");
                var colspan = (int?)cellSpan?.Attribute("colSpan") ?? 1;
                var rowspan = (int?)cellSpan?.Attribute("rowSpan") ?? 1;
                var subList = tc.Element(HwpxNs.Hp + "subList");
                var vAlign = subList?.Attribute("vertAlign")?.Value?.ToLowerInvariant() ?? "top";

                var bfId = tc.Attribute("borderFillIDRef")?.Value;
                var cellCss = $"vertical-align:{vAlign}";
                if (bfId != null)
                {
                    var bgColor = GetBorderFillBgColor(bfId);
                    if (bgColor != null) cellCss += $";background:{bgColor}";
                }

                sb.Append("<td");
                if (colspan > 1) sb.Append($" colspan=\"{colspan}\"");
                if (rowspan > 1) sb.Append($" rowspan=\"{rowspan}\"");
                sb.Append($" style=\"{cellCss}\">");

                if (subList != null)
                {
                    foreach (var cp in subList.Elements(HwpxNs.Hp + "p"))
                        sb.Append(ParagraphToHtml(cp));
                }
                sb.Append("</td>");
            }
            sb.Append("</tr>");
        }
        sb.Append("</table>");
        return sb.ToString();
    }

    private string PicToHtml(XElement pic)
    {
        var imgEl = pic.Descendants().FirstOrDefault(e => e.Name.LocalName == "img");
        var src = imgEl?.Attribute("src")?.Value ?? imgEl?.Attribute("binaryItemIDRef")?.Value;
        if (src != null)
        {
            var binData = _doc.GetBinData(src);
            if (binData != null)
            {
                var ext = Path.GetExtension(src).ToLowerInvariant();
                var mime = ext switch { ".png" => "image/png", ".gif" => "image/gif", ".bmp" => "image/bmp", _ => "image/jpeg" };
                return $"<img src=\"data:{mime};base64,{Convert.ToBase64String(binData)}\" style=\"max-width:100%\"/>";
            }
        }
        return "<span class=\"hwpx-img\">[image]</span>";
    }

    private string GetCharPrCss(string charPrId)
    {
        var charPr = FindCharPr(charPrId);
        if (charPr == null) return "";
        var parts = new List<string>();
        var height = (int?)charPr.Attribute("height") ?? 1000;
        parts.Add($"font-size:{height / 100.0:0.#}pt");
        var color = charPr.Attribute("textColor")?.Value;
        if (color != null && color != "#000000") parts.Add($"color:{color}");
        var fontRef = charPr.Element(HwpxNs.Hh + "fontRef");
        if (fontRef != null)
        {
            var hangulRef = fontRef.Attribute("hangul")?.Value ?? "0";
            var fontName = GetFontName("HANGUL", hangulRef);
            if (fontName != null) parts.Add($"font-family:'{fontName}',sans-serif");
        }
        return string.Join(";", parts);
    }

    private string GetParaPrCss(string paraPrId)
    {
        if (_doc.Header?.Root == null) return "";
        var paraPr = _doc.Header.Root.Descendants(HwpxNs.Hh + "paraPr")
            .FirstOrDefault(p => p.Attribute("id")?.Value == paraPrId);
        if (paraPr == null) return "";
        var parts = new List<string>();
        var align = paraPr.Element(HwpxNs.Hh + "align")?.Attribute("horizontal")?.Value;
        if (align != null && align != "JUSTIFY")
            parts.Add($"text-align:{align.ToLowerInvariant()}");
        else if (align == "JUSTIFY")
            parts.Add("text-align:justify");
        var margin = paraPr.Element(HwpxNs.Hh + "margin");
        if (margin != null)
        {
            var indent = (int?)margin.Attribute("indent") ?? 0;
            if (indent != 0) parts.Add($"text-indent:{indent / 283.46:0.#}mm");
            var left = (int?)margin.Attribute("left") ?? 0;
            if (left != 0) parts.Add($"margin-left:{left / 283.46:0.#}mm");
        }
        var ls = paraPr.Element(HwpxNs.Hh + "lineSpacing");
        if (ls != null)
        {
            var lsType = ls.Attribute("type")?.Value;
            var lsVal = (int?)ls.Attribute("value") ?? 160;
            if (lsType == "PERCENT") parts.Add($"line-height:{lsVal / 100.0:0.##}");
        }
        return string.Join(";", parts);
    }

    private string? GetBorderFillBgColor(string bfId)
    {
        var bf = _doc.Header?.Root?.Descendants(HwpxNs.Hh + "borderFill")
            .FirstOrDefault(b => b.Attribute("id")?.Value == bfId);
        var winBrush = bf?.Descendants(HwpxNs.Hc + "winBrush").FirstOrDefault();
        return winBrush?.Attribute("faceColor")?.Value;
    }

    private string? GetFontName(string lang, string fontRef)
    {
        var fontface = _doc.Header?.Root?.Descendants(HwpxNs.Hh + "fontface")
            .FirstOrDefault(f => f.Attribute("lang")?.Value == lang);
        var font = fontface?.Elements(HwpxNs.Hh + "font")
            .FirstOrDefault(f => f.Attribute("id")?.Value == fontRef);
        return font?.Attribute("face")?.Value;
    }

    private static string EscapeHtml(string text)
        => text.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;");

    private static string HwpxHtmlCss() => """
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { background: #e8e8e8; font-family: '함초롬돋움', 'Malgun Gothic', sans-serif; }
        .page { max-width: 210mm; margin: 20px auto; padding: 20mm 25mm; background: #fff;
                box-shadow: 0 2px 8px rgba(0,0,0,0.15); min-height: 297mm; }
        p { margin: 2px 0; font-size: 10pt; line-height: 1.6; }
        h1 { font-size: 16pt; margin: 12px 0 4px; }
        h2 { font-size: 14pt; margin: 10px 0 4px; }
        h3 { font-size: 12pt; margin: 8px 0 4px; }
        h4, h5, h6 { font-size: 11pt; margin: 6px 0 4px; }
        table { border-collapse: collapse; width: 100%; margin: 8px 0; }
        td, th { border: 1px solid #000; padding: 4px 8px; font-size: 10pt; }
        .hwpx-eq { font-family: 'HancomEQN', serif; color: #333; background: #f5f5f5;
                   padding: 2px 6px; border-radius: 3px; font-size: 0.9em; }
        .hwpx-img { color: #999; font-style: italic; }
        mark { padding: 1px 2px; }
        @media print { body { background: #fff; } .page { box-shadow: none; margin: 0; padding: 20mm; } }
        """;

    /// <summary>Extract all text from a paragraph's hp:run/hp:t elements.</summary>
    private static string ExtractParagraphText(XElement para)
    {
        var runs = para.Elements(HwpxNs.Hp + "run");
        var sb = new StringBuilder();
        foreach (var run in runs)
        {
            foreach (var t in run.Elements(HwpxNs.Hp + "t"))
            {
                sb.Append(t.Value);
            }
            // Handle equations — extract Hancom equation script text
            // Element name is hp:equation (confirmed by hwpxlib). hp:eqEdit is legacy HWP5 class name.
            var eqEl = run.Element(HwpxNs.Hp + "equation")
                ?? run.Element(HwpxNs.Hp + "eqEdit")
                ?? run.Descendants().FirstOrDefault(e =>
                    e.Name.LocalName == "equation" || e.Name.LocalName == "eqEdit");
            if (eqEl != null)
            {
                var script = eqEl.Element(HwpxNs.Hp + "script")?.Value
                    ?? eqEl.Attribute("script")?.Value
                    ?? eqEl.Value;
                if (!string.IsNullOrEmpty(script))
                    sb.Append($"[eq: {script}]");
            }
            // Handle line breaks
            if (run.Element(HwpxNs.Hp + "lineBreak") != null)
                sb.Append('\n');
            if (run.Element(HwpxNs.Hp + "tab") != null)
                sb.Append('\t');
        }
        return sb.ToString();
    }

    /// <summary>Extract runs with formatting annotations.</summary>
    private static List<(string Text, Dictionary<string, string> Format)> ExtractAnnotatedRuns(XElement para)
    {
        var result = new List<(string, Dictionary<string, string>)>();
        foreach (var run in para.Elements(HwpxNs.Hp + "run"))
        {
            var text = string.Join("", run.Elements(HwpxNs.Hp + "t").Select(t => t.Value));
            if (string.IsNullOrEmpty(text)) continue;

            var format = new Dictionary<string, string>();
            var charPrIdRef = run.Attribute("charPrIDRef")?.Value;
            if (charPrIdRef != null)
                format["charPrIDRef"] = charPrIdRef;

            result.Add((text, format));
        }
        return result;
    }

    /// <summary>Get paragraph style info from attributes and header.xml lookup.</summary>
    private (string? HeadingLevel, string Alignment) GetParagraphStyleInfo(XElement para)
    {
        var styleIdRef = para.Attribute("styleIDRef")?.Value;
        var paraPrIdRef = para.Attribute("paraPrIDRef")?.Value;

        string? headingLevel = null;
        string alignment = "LEFT";

        // Look up style in header.xml
        if (_doc.Header != null && styleIdRef != null)
        {
            var style = _doc.Header.Root!.Descendants(HwpxNs.Hh + "style")
                .FirstOrDefault(s => s.Attribute("id")?.Value == styleIdRef);
            if (style != null)
            {
                var name = style.Attribute("name")?.Value ?? "";
                // Korean heading styles: "개요 1", "개요 2", etc.
                var headingMatch = System.Text.RegularExpressions.Regex.Match(name, @"개요\s*(\d+)");
                if (headingMatch.Success)
                    headingLevel = headingMatch.Groups[1].Value;
                // English heading styles
                var engMatch = System.Text.RegularExpressions.Regex.Match(name, @"(?i)heading\s*(\d+)");
                if (engMatch.Success)
                    headingLevel = engMatch.Groups[1].Value;
            }
        }

        // Look up paragraph properties for alignment and heading
        if (_doc.Header != null && paraPrIdRef != null)
        {
            var paraPr = _doc.Header.Root!.Descendants(HwpxNs.Hh + "paraPr")
                .FirstOrDefault(p => p.Attribute("id")?.Value == paraPrIdRef);
            if (paraPr != null)
            {
                // Real HWPX: alignment is a child element <hh:align horizontal="LEFT"/>
                var alignEl = paraPr.Element(HwpxNs.Hh + "align");
                alignment = alignEl?.Attribute("horizontal")?.Value ?? "LEFT";

                // Heading detection via paraPr > heading element (type="OUTLINE")
                if (headingLevel == null)
                {
                    var heading = paraPr.Element(HwpxNs.Hh + "heading");
                    if (heading?.Attribute("type")?.Value == "OUTLINE"
                        && int.TryParse(heading.Attribute("level")?.Value, out var hl) && hl >= 1)
                        headingLevel = hl.ToString();
                }
            }
        }

        return (headingLevel, alignment);
    }

    private int CountRemainingParagraphs(int currentLine)
    {
        int total = _doc.AllParagraphs().Count();
        return Math.Max(0, total - currentLine);
    }

    private static int CountWords(string text)
    {
        if (string.IsNullOrWhiteSpace(text)) return 0;
        // Korean: each syllable cluster counts as a word boundary
        // Simple heuristic: split on whitespace, count non-empty
        return text.Split(Array.Empty<char>(), StringSplitOptions.RemoveEmptyEntries).Length;
    }

    private static string FormatHwpUnit(int hwpUnit)
    {
        var mm = hwpUnit / 283.46;
        return $"{mm:0.#}mm";
    }
}
