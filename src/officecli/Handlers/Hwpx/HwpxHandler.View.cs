using System.Text;
using System.Text.Json.Nodes;
using System.Xml.Linq;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class HwpxHandler
{
    private sealed record FormFieldInfo(string Type, string Id, string Name, string Text, string? HelpText, bool IsDefault);

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

        // Level 7: BinData integrity — orphan/missing binary references
        var referencedBinData = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var sec in _doc.Sections)
        {
            foreach (var el in sec.Root.Descendants())
            {
                var binRef = el.Attribute("binaryItemIDRef")?.Value;
                if (binRef != null) referencedBinData.Add(binRef);
            }
        }
        var actualBinData = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var entry in _doc.Archive.Entries)
        {
            if (entry.FullName.Contains("BinData/", StringComparison.OrdinalIgnoreCase))
                actualBinData.Add(System.IO.Path.GetFileNameWithoutExtension(entry.FullName));
        }
        foreach (var missing in referencedBinData.Except(actualBinData))
        {
            issues.Add(new DocumentIssue
            {
                Id = $"HWPX-{++issueId:D3}", Type = IssueType.Structure,
                Severity = IssueSeverity.Error, Path = "/BinData",
                Message = $"Referenced binary '{missing}' not found in archive",
                Context = null
            });
        }
        foreach (var orphan in actualBinData.Except(referencedBinData))
        {
            issues.Add(new DocumentIssue
            {
                Id = $"HWPX-{++issueId:D3}", Type = IssueType.Structure,
                Severity = IssueSeverity.Info, Path = "/BinData",
                Message = $"Orphan binary '{orphan}' not referenced by any element",
                Context = null
            });
        }

        // Level 8: Field pair validation — unclosed fieldBegin/fieldEnd
        foreach (var sec in _doc.Sections)
        {
            var fieldBegins = sec.Root.Descendants(HwpxNs.Hp + "fieldBegin").ToList();
            var fieldEnds = sec.Root.Descendants(HwpxNs.Hp + "fieldEnd").ToList();
            if (fieldBegins.Count != fieldEnds.Count)
            {
                issues.Add(new DocumentIssue
                {
                    Id = $"HWPX-{++issueId:D3}", Type = IssueType.Structure,
                    Severity = IssueSeverity.Warning,
                    Path = $"/section[{sec.Index + 1}]",
                    Message = $"Field count mismatch: {fieldBegins.Count} opens vs {fieldEnds.Count} closes",
                    Context = null
                });
            }
        }

        // Level 9: Section count consistency — manifest vs actual
        if (_doc.ManifestDoc != null)
        {
            var manifestSections = _doc.ManifestDoc.Descendants()
                .Count(e => e.Attribute("media-type")?.Value == "application/xml"
                    && (e.Attribute("href")?.Value?.StartsWith("section") ?? false));
            if (manifestSections != _doc.Sections.Count)
            {
                issues.Add(new DocumentIssue
                {
                    Id = $"HWPX-{++issueId:D3}", Type = IssueType.Structure,
                    Severity = IssueSeverity.Error, Path = "/content.hpf",
                    Message = $"Section count mismatch: manifest={manifestSections}, loaded={_doc.Sections.Count}",
                    Context = null
                });
            }
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

    public string ViewAsForms(bool auto = true)
    {
        var sb = new StringBuilder();
        var fields = EnumerateInteractiveFormFields().ToList();
        foreach (var field in fields)
        {
            var nameSuffix = string.IsNullOrEmpty(field.Name) ? "" : $" {field.Name}";
            sb.AppendLine($"  [{field.Id}] {field.Type}{nameSuffix}: \"{field.Text}\"{(field.IsDefault ? " (default)" : "")}");
        }
        sb.Insert(0, $"Form fields: {fields.Count}\n");

        if (auto)
        {
            var recognized = RecognizeFormFields();
            if (recognized.Count > 0)
            {
                var adjacentCount = recognized.Count(f => f.Strategy == "adjacent");
                var headerDataCount = recognized.Count(f => f.Strategy == "header-data");
                var strategySummary = new List<string>();
                if (adjacentCount > 0) strategySummary.Add($"{adjacentCount} adjacent");
                if (headerDataCount > 0) strategySummary.Add($"{headerDataCount} header-data");
                var otherCount = recognized.Count - adjacentCount - headerDataCount;
                if (otherCount > 0) strategySummary.Add($"{otherCount} other");

                sb.AppendLine();
                sb.AppendLine($"Forms: {recognized.Count} fields recognized ({string.Join(", ", strategySummary)})");
                sb.AppendLine();

                // Compute column widths
                int labelW = Math.Max(5, recognized.Max(f => f.Label.Length));
                int valueW = Math.Max(5, recognized.Max(f => f.Value.Length));
                int pathW = Math.Max(4, recognized.Max(f => f.Path.Length));
                int stratW = Math.Max(8, recognized.Max(f => f.Strategy.Length));

                // Cap widths to keep output readable
                labelW = Math.Min(labelW, 20);
                valueW = Math.Min(valueW, 24);
                pathW = Math.Min(pathW, 44);

                sb.AppendLine($"  {"Label".PadRight(labelW)}  {"Value".PadRight(valueW)}  {"Path".PadRight(pathW)}  Strategy");
                sb.AppendLine($"  {new string('\u2500', labelW + 2 + valueW + 2 + pathW + 2 + stratW)}");

                foreach (var f in recognized)
                {
                    var label = f.Label.Length > labelW ? f.Label[..(labelW - 1)] + "\u2026" : f.Label.PadRight(labelW);
                    var value = f.Value.Length > valueW ? f.Value[..(valueW - 1)] + "\u2026" : f.Value.PadRight(valueW);
                    var path = f.Path.Length > pathW ? f.Path[..(pathW - 1)] + "\u2026" : f.Path.PadRight(pathW);
                    sb.AppendLine($"  {label}  {value}  {path}  {f.Strategy}");
                }
            }

            // F8: Form confidence score
            int totalTables = _doc.Sections.Sum(s => s.Tables.Count);
            if (totalTables > 0)
            {
                var formTablePaths = recognized
                    .Select(f => System.Text.RegularExpressions.Regex.Match(f.Path, @"^/section\[\d+\]/tbl\[\d+\]").Value)
                    .Where(p => !string.IsNullOrEmpty(p))
                    .Distinct()
                    .Count();
                double confidence = (double)formTablePaths / totalTables;
                sb.AppendLine();
                sb.AppendLine($"Form confidence: {confidence:P0} ({formTablePaths}/{totalTables} tables are form-like)");
            }
        }

        return sb.ToString().TrimEnd();
    }

    /// <summary>JSON output for forms view. Supports CLICK_HERE + auto-recognized fields.</summary>
    public JsonNode ViewAsFormsJson(bool auto = true)
    {
        var result = new JsonObject();

        var clickFields = new JsonArray();
        var formFields = new JsonArray();
        foreach (var field in EnumerateInteractiveFormFields())
        {
            if (field.Type == "CLICK_HERE")
            {
                clickFields.Add(new JsonObject {
                    ["id"] = field.Id, ["text"] = field.Text,
                    ["helpText"] = field.HelpText, ["isDefault"] = field.IsDefault
                });
            }

            formFields.Add(new JsonObject {
                ["id"] = field.Id,
                ["type"] = field.Type,
                ["name"] = field.Name,
                ["text"] = field.Text,
                ["helpText"] = field.HelpText,
                ["isDefault"] = field.IsDefault
            });
        }
        result["clickHere"] = clickFields;
        result["formFields"] = formFields;

        if (auto)
        {
            var autoFields = new JsonArray();
            foreach (var f in RecognizeFormFields())
            {
                autoFields.Add(new JsonObject {
                    ["label"] = f.Label, ["value"] = f.Value,
                    ["path"] = f.Path, ["row"] = f.Row, ["col"] = f.Col,
                    ["strategy"] = f.Strategy
                });
            }
            result["forms"] = autoFields;
        }

        return result;
    }

    private IEnumerable<FormFieldInfo> EnumerateInteractiveFormFields()
    {
        foreach (var sec in _doc.Sections)
        {
            foreach (var run in sec.Root.Descendants(HwpxNs.Hp + "run"))
            {
                var ctrl = run.Element(HwpxNs.Hp + "ctrl");
                var fieldBegin = ctrl?.Element(HwpxNs.Hp + "fieldBegin");
                var fieldType = fieldBegin?.Attribute("type")?.Value;
                if (fieldType is not ("CLICK_HERE" or "CHECKBOX" or "DROPDOWN")) continue;

                var field = fieldBegin!;
                var id = field.Attribute("id")?.Value ?? "?";
                var name = field.Attribute("name")?.Value ?? "";
                var helpText = field.Descendants()
                    .FirstOrDefault(p => p.Attribute("name")?.Value is "Direction" or "Label")
                    ?.Value;
                var nextRun = run.ElementsAfterSelf(HwpxNs.Hp + "run").FirstOrDefault();
                var text = nextRun?.Elements(HwpxNs.Hp + "t").FirstOrDefault()?.Value ?? "";
                var isDefault = !string.IsNullOrEmpty(helpText) && text == helpText;

                yield return new FormFieldInfo(fieldType, id, name, text, helpText, isDefault);
            }
        }
    }

    // ==================== Object Finder (Plan 82) ====================

    private static readonly string[] DefaultObjectTypes = ["picture", "field", "bookmark", "equation"];

    /// <summary>List objects of specified type(s) with paths and previews.</summary>
    public string ViewAsObjects(string? objectType = null)
    {
        var types = objectType != null ? [objectType] : DefaultObjectTypes;
        var sb = new StringBuilder();
        int total = 0;

        foreach (var type in types)
        {
            List<System.Xml.Linq.XElement> elements;
            try { elements = ExecuteSelector(type); }
            catch { continue; }

            // formfield: list interactive form fields only
            if (type == "formfield")
                elements = elements.Where(e => e.Attribute("type")?.Value is "CLICK_HERE" or "CHECKBOX" or "DROPDOWN").ToList();

            if (elements.Count == 0) continue;
            total += elements.Count;
            sb.AppendLine($"{type}: {elements.Count}");
            foreach (var el in elements)
            {
                var path = BuildPath(el);
                var preview = GetElementText(el);
                if (preview.Length > 60) preview = preview[..60] + "…";
                if (string.IsNullOrWhiteSpace(preview)) preview = $"({el.Name.LocalName})";

                // Extra info per type
                var extra = type switch
                {
                    "picture" or "img" => el.Attribute("binaryItemIDRef")?.Value is { } r ? $" [{r}]" : "",
                    "field" or "formfield" => el.Attribute("type")?.Value is { } t ? $" [{t}]" : "",
                    "equation" => "",
                    "bookmark" => el.Attribute("name")?.Value is { } n ? $" [{n}]" : "",
                    _ => ""
                };
                sb.AppendLine($"  {path}{extra}: {preview}");
            }
            sb.AppendLine();
        }

        if (total == 0)
            return "(no objects found)";
        sb.Insert(0, $"Objects: {total}\n\n");
        return sb.ToString().TrimEnd();
    }

    /// <summary>JSON output for object finder.</summary>
    public JsonNode ViewAsObjectsJson(string? objectType = null)
    {
        var types = objectType != null ? [objectType] : DefaultObjectTypes;
        var result = new JsonObject();

        foreach (var type in types)
        {
            List<System.Xml.Linq.XElement> elements;
            try { elements = ExecuteSelector(type); }
            catch { continue; }

            if (type == "formfield")
                elements = elements.Where(e => e.Attribute("type")?.Value is "CLICK_HERE" or "CHECKBOX" or "DROPDOWN").ToList();

            if (elements.Count == 0) continue;

            var arr = new JsonArray();
            foreach (var el in elements)
            {
                var obj = new JsonObject
                {
                    ["path"] = BuildPath(el),
                    ["text"] = GetElementText(el)
                };
                // Type-specific attributes
                if (el.Attribute("binaryItemIDRef")?.Value is { } binRef) obj["binaryRef"] = binRef;
                if (el.Attribute("type")?.Value is { } ft) obj["fieldType"] = ft;
                if (el.Attribute("name")?.Value is { } bname) obj["name"] = bname;
                arr.Add(obj);
            }
            result[type] = arr;
        }

        return result;
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

    // ==================== Table Map (Plan 71) ====================

    /// <summary>
    /// Display all tables with grid structure, recognized labels, and cell paths.
    /// </summary>
    public string ViewAsTables()
    {
        var sb = new StringBuilder();
        int tblCount = 0;

        foreach (var (sec, tbl, localTblIdx) in _doc.AllTables())
        {
            tblCount++;
            var (grid, cellList) = BuildTableGrid(tbl);
            if (cellList.Count == 0) continue;

            int maxRow = grid.GetLength(0), maxCol = grid.GetLength(1);
            var basePath = $"/section[{sec.Index + 1}]/tbl[{localTblIdx + 1}]";
            sb.AppendLine($"Table {tblCount} ({basePath}, {maxRow}×{maxCol}):");

            // Grid visualization
            for (int r = 0; r < maxRow; r++)
            {
                sb.Append($"  [{r}] ");
                for (int c = 0; c < maxCol; c++)
                {
                    var cell = grid[r, c];
                    if (cell == null) { sb.Append("·  "); continue; }

                    // Skip duplicate merged cell refs (only show on first occurrence)
                    var (cr, cc, rs, cs) = GetCellAddr(cell);
                    if (cr != r || cc != c) { sb.Append("↕  "); continue; }

                    var text = ExtractCellText(cell).Trim();
                    var preview = text.Length > 12 ? text[..12] + "…" : text;
                    if (string.IsNullOrEmpty(preview)) preview = "(empty)";

                    var span = (rs > 1 || cs > 1) ? $"[{rs}×{cs}]" : "";
                    sb.Append($"{preview}{span}  ");
                }
                sb.AppendLine();
            }

            // Recognized fields for this table
            var fields = new List<RecognizedField>();
            var tableGrid = grid; // reuse
            var seen = new HashSet<XElement>();
            foreach (var (tc, row, col, rowSpan, colSpan) in cellList)
            {
                if (seen.Contains(tc)) continue;
                seen.Add(tc);
                var cellText = ExtractCellText(tc);
                if (!IsLabelCell(cellText)) continue;
                int targetCol = col + colSpan;
                if (targetCol < maxCol)
                {
                    var valueCell = grid[row, targetCol];
                    if (valueCell != null && valueCell != tc)
                    {
                        var value = ExtractCellText(valueCell).Trim();
                        if (!string.IsNullOrEmpty(value))
                            fields.Add(new RecognizedField(
                                NormalizeLabel(cellText), value, basePath, row, col, "adjacent"));
                    }
                }
            }
            if (fields.Count > 0)
            {
                sb.AppendLine($"  Labels: {fields.Count}");
                foreach (var f in fields)
                    sb.AppendLine($"    {f.Label}: {f.Value} (r{f.Row},c{f.Col})");
            }
            sb.AppendLine();
        }

        if (tblCount == 0)
            sb.AppendLine("(no tables)");
        else
            sb.Insert(0, $"Tables: {tblCount}\n\n");

        return sb.ToString().TrimEnd();
    }

    // ==================== Markdown Export (Plan 72) ====================

    /// <summary>Export document as GitHub Flavored Markdown.</summary>
    public string ViewAsMarkdown()
    {
        var sb = new StringBuilder();

        foreach (var (section, element, path) in _doc.AllContentInOrder())
        {
            var localName = element.Name.LocalName;
            if (localName == "p")
            {
                // Check if this paragraph is inside a table cell (skip — handled by table renderer)
                if (element.Ancestors().Any(a => a.Name.LocalName == "tc")) continue;

                var styleInfo = GetParagraphStyleInfo(element);
                var mdLine = ParagraphToMarkdown(element);
                if (string.IsNullOrWhiteSpace(mdLine)) { sb.AppendLine(); continue; }

                if (!string.IsNullOrEmpty(styleInfo.HeadingLevel))
                {
                    var level = Math.Clamp(int.Parse(styleInfo.HeadingLevel), 1, 6);
                    sb.AppendLine($"{new string('#', level)} {mdLine}");
                }
                else
                {
                    sb.AppendLine(mdLine);
                }
                sb.AppendLine();
            }
        }

        // Render tables
        foreach (var (sec, tbl, localTblIdx) in _doc.AllTables())
        {
            var (grid, cellList) = BuildTableGrid(tbl);
            if (cellList.Count == 0) continue;
            int maxRow = grid.GetLength(0), maxCol = grid.GetLength(1);

            // F5: Single-cell tables → emit as structured text instead of table
            if (maxRow == 1 && maxCol == 1 && cellList.Count == 1)
            {
                var cellText = ExtractCellText(cellList[0].Tc).Trim();
                if (!string.IsNullOrEmpty(cellText))
                {
                    var lines = cellText.Split('\n');
                    foreach (var line in lines)
                    {
                        var trimmed = line.Trim();
                        if (string.IsNullOrEmpty(trimmed)) { sb.AppendLine(); continue; }
                        var m = System.Text.RegularExpressions.Regex.Match(trimmed, @"^(\d+[.)]|[가-하][.]|[a-z][.)]) (.+)$");
                        if (m.Success)
                            sb.AppendLine($"**{m.Groups[1].Value}** {m.Groups[2].Value}");
                        else
                            sb.AppendLine(trimmed);
                    }
                    sb.AppendLine();
                }
                continue;
            }

            // F6: Pseudo-table demotion — skip tables with <=3 rows and >=30% empty cells
            if (maxRow <= 3)
            {
                int totalCells = maxRow * maxCol;
                int emptyCells = 0;
                for (int r = 0; r < maxRow; r++)
                    for (int c = 0; c < maxCol; c++)
                    {
                        var cell = grid[r, c];
                        if (cell == null || string.IsNullOrWhiteSpace(ExtractCellText(cell)))
                            emptyCells++;
                    }
                if (totalCells > 0 && (double)emptyCells / totalCells >= 0.3)
                {
                    for (int r = 0; r < maxRow; r++)
                        for (int c = 0; c < maxCol; c++)
                        {
                            var cell = grid[r, c];
                            if (cell == null) continue;
                            var (cr, cc, _, _) = GetCellAddr(cell);
                            if (cr != r || cc != c) continue;
                            var text = ExtractCellText(cell).Trim();
                            if (!string.IsNullOrEmpty(text))
                                sb.AppendLine(text);
                        }
                    sb.AppendLine();
                    continue;
                }
            }

            for (int r = 0; r < maxRow; r++)
            {
                sb.Append("| ");
                for (int c = 0; c < maxCol; c++)
                {
                    var cell = grid[r, c];
                    if (cell == null) { sb.Append("| "); continue; }
                    var (cr, cc, _, _) = GetCellAddr(cell);
                    if (cr != r || cc != c) { sb.Append("| "); continue; } // merged continuation
                    var text = ExtractCellText(cell).Trim().Replace("\n", " ").Replace("|", "\\|");
                    sb.Append($"{text} | ");
                }
                sb.AppendLine();

                // Separator after header row
                if (r == 0)
                {
                    sb.Append("| ");
                    for (int c = 0; c < maxCol; c++)
                        sb.Append("--- | ");
                    sb.AppendLine();
                }
            }
            sb.AppendLine();
        }

        return sb.ToString().Trim();
    }

    private string ParagraphToMarkdown(XElement p)
    {
        var sb = new StringBuilder();
        foreach (var run in p.Elements(HwpxNs.Hp + "run"))
            sb.Append(RunToMarkdown(run));
        return sb.ToString().Trim();
    }

    private string RunToMarkdown(XElement run)
    {
        var sb = new StringBuilder();
        var charPrId = run.Attribute("charPrIDRef")?.Value ?? "0";
        var charPr = FindCharPr(charPrId);
        var hasBold = charPr?.Element(HwpxNs.Hh + "bold") != null;
        var hasItalic = charPr?.Element(HwpxNs.Hh + "italic") != null;
        var soEl = charPr?.Element(HwpxNs.Hh + "strikeout");
        var hasStrikeout = soEl != null && soEl.Attribute("shape")?.Value != "NONE";

        var textParts = new StringBuilder();
        foreach (var child in run.Elements())
        {
            switch (child.Name.LocalName)
            {
                case "t":
                    textParts.Append(child.Value);
                    break;
                case "lineBreak":
                    textParts.Append("  \n"); // MD hard line break
                    break;
                case "tab":
                    textParts.Append('\t');
                    break;
                case "equation":
                    var script = child.Element(HwpxNs.Hp + "script")?.Value
                        ?? child.Attribute("script")?.Value ?? child.Value;
                    textParts.Append($"`{script.Trim()}`");
                    break;
                case "img": case "picture":
                    var src = child.Attribute("binaryItemIDRef")?.Value ?? "image";
                    textParts.Append($"![{src}]({src})");
                    break;
            }
        }

        var text = textParts.ToString();
        if (string.IsNullOrEmpty(text)) return "";

        // F4: GFM tilde escape — prevent false strikethrough from literal tildes
        // Must happen BEFORE strikethrough wrapping
        if (!hasStrikeout)
            text = text.Replace("~", @"\~");

        if (hasStrikeout) text = $"~~{text}~~";
        if (hasBold && hasItalic) text = $"***{text}***";
        else if (hasBold) text = $"**{text}**";
        else if (hasItalic) text = $"*{text}*";

        sb.Append(text);
        return sb.ToString();
    }

    /// <summary>JSON output for table map view.</summary>
    public JsonNode ViewAsTablesJson()
    {
        var result = new JsonObject();
        var tablesArr = new JsonArray();

        foreach (var (sec, tbl, localTblIdx) in _doc.AllTables())
        {
            var (grid, cellList) = BuildTableGrid(tbl);
            if (cellList.Count == 0) continue;

            int maxRow = grid.GetLength(0), maxCol = grid.GetLength(1);
            var basePath = $"/section[{sec.Index + 1}]/tbl[{localTblIdx + 1}]";

            var tblObj = new JsonObject
            {
                ["path"] = basePath,
                ["rows"] = maxRow,
                ["cols"] = maxCol
            };

            // Cells grid
            var cellsArr = new JsonArray();
            for (int r = 0; r < maxRow; r++)
            {
                var rowArr = new JsonArray();
                for (int c = 0; c < maxCol; c++)
                {
                    var cell = grid[r, c];
                    if (cell == null) { rowArr.Add((JsonNode?)null); continue; }
                    var (cr, cc, rs, cs) = GetCellAddr(cell);
                    if (cr != r || cc != c) { rowArr.Add("↕"); continue; }
                    var text = ExtractCellText(cell).Trim();
                    rowArr.Add(new JsonObject
                    {
                        ["text"] = text,
                        ["path"] = $"{basePath}/tr[{r + 1}]/tc[{c + 1}]",
                        ["rowSpan"] = rs,
                        ["colSpan"] = cs
                    });
                }
                cellsArr.Add(rowArr);
            }
            tblObj["cells"] = cellsArr;

            tablesArr.Add(tblObj);
        }

        result["tables"] = tablesArr;
        return result;
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

        // F3: Legal appendix heading detection (별표/별지/별첨, 제N조 관련)
        if (headingLevel == null)
        {
            var text = ExtractParagraphText(para);
            if (System.Text.RegularExpressions.Regex.IsMatch(text, @"^\s*\[?별[표지첨]\s*(?:\d+\s*)?(?:의\s*\d+\s*)?(?:\]|$)"))
                headingLevel = "2";
            else if (System.Text.RegularExpressions.Regex.IsMatch(text, @"^\s*\(제\s*\d+\s*조\s*관련\)"))
                headingLevel = "3";
            // G3: Space-tolerant legal heading detection
            else
            {
                var compacted = System.Text.RegularExpressions.Regex.Replace(text.TrimStart(), @"\s+", "");
                if (System.Text.RegularExpressions.Regex.IsMatch(compacted, @"^제\d+[장편](?![에의은을로서와가는도])"))
                    headingLevel = "1";
                else if (System.Text.RegularExpressions.Regex.IsMatch(compacted, @"^제\d+[절관](?![에의은을로서와가는도])"))
                    headingLevel = "2";
            }
        }

        // Plan 99.9.I3: Font-size ratio heading detection (fallback when outline level not set)
        if (headingLevel == null && _doc.Header != null)
        {
            var charPrIdRef = para.Elements(HwpxNs.Hp + "run")
                .FirstOrDefault()?.Attribute("charPrIDRef")?.Value;
            if (charPrIdRef != null)
            {
                var charPr = FindCharPr(charPrIdRef);
                if (charPr != null)
                {
                    double fontSize = GetFontSizePt(charPr);
                    double baseFontSize = _baseFontSizePt ??= ComputeBaseFontSize();
                    if (baseFontSize > 0)
                    {
                        double ratio = fontSize / baseFontSize;
                        if (ratio >= 1.5) headingLevel = "1";       // H1: 150%+
                        else if (ratio >= 1.3) headingLevel = "2";  // H2: 130%+
                        else if (ratio >= 1.15) headingLevel = "3"; // H3: 115%+
                    }
                }
            }
        }

        return (headingLevel, alignment);
    }

    /// <summary>
    /// Plan 99.9.I3: Compute base (body) font size by finding the most frequent font size across all paragraphs.
    /// Used as denominator for heading ratio detection.
    /// </summary>
    private double ComputeBaseFontSize()
    {
        var sizeCounts = new Dictionary<double, int>();
        foreach (var (_, para, _) in _doc.AllParagraphs())
        {
            var charPrIdRef = para.Elements(HwpxNs.Hp + "run")
                .FirstOrDefault()?.Attribute("charPrIDRef")?.Value;
            if (charPrIdRef == null) continue;
            var charPr = FindCharPr(charPrIdRef);
            if (charPr == null) continue;
            double size = GetFontSizePt(charPr);
            sizeCounts[size] = sizeCounts.GetValueOrDefault(size) + 1;
        }
        return sizeCounts.Count > 0
            ? sizeCounts.MaxBy(kv => kv.Value).Key
            : 10.0; // default 10pt
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
