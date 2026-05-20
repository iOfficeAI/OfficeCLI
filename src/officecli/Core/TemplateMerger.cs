// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Core;

/// <summary>
/// Merges a template Office document with JSON data by replacing {{key}} placeholders.
/// Supports DOCX, XLSX, and PPTX formats.
/// </summary>
internal static class TemplateMerger
{
    // Allow optional outer whitespace ({{ name }}), hyphenated keys
    // ({{user-id}}), inner spaces ({{ first name }}), and array indexing
    // ({{items[0]}}). Outer \s* is stripped; inner spaces are preserved as
    // part of the key. The captured group is also Trim()'d at the use site
    // so a placeholder like "{{  name  }}" resolves the same as "{{name}}".
    // Match is non-greedy on the inner segment so trailing whitespace
    // followed by }} is not absorbed into the key.
    private static readonly Regex PlaceholderPattern = new(@"\{\{\s*(\w[\w.\-\[\] ]*?)\s*\}\}", RegexOptions.Compiled);

    /// <summary>
    /// Result of a merge operation.
    /// </summary>
    public record MergeResult(int ReplacedCount, List<string> UnresolvedPlaceholders, List<string> UsedKeys);

    /// <summary>
    /// Parse merge data from a string argument. If the value ends with .json and the file exists,
    /// read from file; otherwise parse as inline JSON.
    /// </summary>
    public static Dictionary<string, string> ParseMergeData(string dataArg)
    {
        string jsonText;

        if (dataArg.EndsWith(".json", StringComparison.OrdinalIgnoreCase) && File.Exists(dataArg))
        {
            jsonText = File.ReadAllText(dataArg);
        }
        else
        {
            jsonText = dataArg;
        }

        var jsonNode = JsonNode.Parse(jsonText)
            ?? throw new CliException("Invalid JSON data: parsed to null")
            {
                Code = "invalid_json",
                Suggestion = "Provide valid JSON object, e.g. '{\"name\":\"Alice\"}'"
            };

        if (jsonNode is not JsonObject jsonObj)
            throw new CliException("JSON data must be an object (not array or primitive)")
            {
                Code = "invalid_json",
                Suggestion = "Provide a JSON object, e.g. '{\"name\":\"Alice\"}'"
            };

        var data = new Dictionary<string, string>();
        // Pass 1: literal top-level keys win. {{a.b}} with data {"a.b":"X"}
        // resolves to "X" regardless of whether {"a":{"b":...}} also exists.
        // CONSISTENCY(merge-literal-key-precedence): mirrors the hyphen-key
        // contract (R21) — literal lookup is the canonical path; nested
        // dot-path flattening is a convenience layered on top, not a
        // replacement.
        foreach (var kvp in jsonObj)
        {
            data[kvp.Key] = kvp.Value?.ToString() ?? "";
        }
        // Pass 2: flatten nested objects into dot paths ("a"→{"b":"v"} →
        // "a.b":"v") and arrays into bracket paths ("items"→[v0,v1] →
        // "items[0]":"v0", "items[1]":"v1"). Only fill keys that pass 1 did
        // NOT already write, so a literal "a.b" or "items[0]" sibling at the
        // root stays authoritative.
        foreach (var kvp in jsonObj)
        {
            if (kvp.Value is JsonObject nested)
                FlattenNested(nested, kvp.Key, data);
            else if (kvp.Value is JsonArray arr)
                FlattenArray(arr, kvp.Key, data);
        }
        return data;
    }

    /// <summary>
    /// Walk a JsonArray and emit <c>prefix[i]</c> entries for each element.
    /// Nested objects/arrays recurse so e.g. <c>users[0].name</c> and
    /// <c>matrix[0][1]</c> resolve. Existing literal sibling keys win, same
    /// precedence rule as <see cref="FlattenNested"/>.
    /// </summary>
    private static void FlattenArray(JsonArray arr, string prefix, Dictionary<string, string> data)
    {
        for (int i = 0; i < arr.Count; i++)
        {
            var path = $"{prefix}[{i}]";
            var item = arr[i];
            if (item is JsonObject childObj)
            {
                FlattenNested(childObj, path, data);
                if (!data.ContainsKey(path))
                    data[path] = item.ToString();
            }
            else if (item is JsonArray childArr)
            {
                FlattenArray(childArr, path, data);
            }
            else if (!data.ContainsKey(path))
            {
                data[path] = item?.ToString() ?? "";
            }
        }
    }

    /// <summary>
    /// Recursively walk a JsonObject and write <c>prefix.child</c> entries into
    /// <paramref name="data"/>, skipping any path a literal sibling already
    /// populated. Nested arrays are skipped (no canonical placeholder index
    /// syntax). Primitive leaves use the same <c>ToString()</c> projection as
    /// the top-level pass for output parity.
    /// </summary>
    private static void FlattenNested(JsonObject obj, string prefix, Dictionary<string, string> data)
    {
        foreach (var kvp in obj)
        {
            var path = $"{prefix}.{kvp.Key}";
            if (kvp.Value is JsonObject child)
            {
                // Recurse first, then write the object's own ToString only
                // as a last resort (matches how Pass 1 stores top-level
                // objects: stringified JSON when no deeper match exists).
                FlattenNested(child, path, data);
                if (!data.ContainsKey(path))
                    data[path] = kvp.Value?.ToString() ?? "";
            }
            else if (kvp.Value is JsonArray nestedArr)
            {
                FlattenArray(nestedArr, path, data);
            }
            else if (!data.ContainsKey(path))
            {
                data[path] = kvp.Value?.ToString() ?? "";
            }
        }
    }

    /// <summary>
    /// Merge a template document with data. Copies template to output, then replaces placeholders.
    /// </summary>
    public static MergeResult Merge(string templatePath, string outputPath, Dictionary<string, string> data)
    {
        if (!File.Exists(templatePath))
            throw new CliException($"Template file not found: {templatePath}")
            {
                Code = "file_not_found",
                Suggestion = "Check the template file path."
            };

        File.Copy(templatePath, outputPath, overwrite: true);

        var ext = Path.GetExtension(outputPath).ToLowerInvariant();
        return ext switch
        {
            ".docx" => MergeDocx(outputPath, data),
            ".xlsx" => MergeXlsx(outputPath, data),
            ".pptx" => MergePptx(outputPath, data),
            _ => throw new CliException($"Unsupported file type for merge: {ext}")
            {
                Code = "unsupported_type",
                ValidValues = [".docx", ".xlsx", ".pptx"]
            }
        };
    }

    private static MergeResult MergeDocx(string filePath, Dictionary<string, string> data)
    {
        var usedKeys = new HashSet<string>();
        int totalReplacements = 0;

        // CONSISTENCY(merge-single-pass): walk every <w:t> in body + aux parts
        // in one pass with a single-pass regex substitute. The earlier
        // per-key handler.Set(find/replace) loop fed each substituted value
        // back through the next iteration, so a value like "{{name}}" inside
        // data["greeting"] would itself be replaced — and only keys whose
        // placeholder still survived the cascade counted as "used".
        ReplacePlaceholdersInDocx(filePath, data, usedKeys, count => totalReplacements += count);

        // Scan for unresolved placeholders
        var unresolved = ScanUnresolvedDocx(filePath);

        return new MergeResult(totalReplacements, unresolved, usedKeys.ToList());
    }

    private static void ReplacePlaceholdersInDocx(string filePath, Dictionary<string, string> data, HashSet<string> usedKeys, Action<int> bumpReplacements)
    {
        using var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(filePath, true);
        var mainPart = doc.MainDocumentPart;
        if (mainPart == null) return;

        IEnumerable<(DocumentFormat.OpenXml.Packaging.OpenXmlPart Part, DocumentFormat.OpenXml.OpenXmlPartRootElement? Root)> allParts()
        {
            if (mainPart.Document != null)
                yield return (mainPart, mainPart.Document);
            foreach (var hp in mainPart.HeaderParts)
                yield return (hp, hp.Header);
            foreach (var fp in mainPart.FooterParts)
                yield return (fp, fp.Footer);
            if (mainPart.FootnotesPart?.Footnotes != null)
                yield return (mainPart.FootnotesPart, mainPart.FootnotesPart.Footnotes);
            if (mainPart.EndnotesPart?.Endnotes != null)
                yield return (mainPart.EndnotesPart, mainPart.EndnotesPart.Endnotes);
            if (mainPart.WordprocessingCommentsPart?.Comments != null)
                yield return (mainPart.WordprocessingCommentsPart, mainPart.WordprocessingCommentsPart.Comments);
        }

        foreach (var (part, root) in allParts())
        {
            if (root == null) continue;
            bool changed = false;
            foreach (var t in root.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>())
            {
                var original = t.Text ?? "";
                if (original.Length == 0 || !original.Contains("{{")) continue;
                var replaced = SinglePassReplace(original, data, out var matched, usedKeys, bumpReplacements);
                if (matched && replaced != original)
                {
                    t.Text = replaced;
                    t.Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve;
                    changed = true;
                }
            }
            if (changed) root.Save();
        }
    }

    /// <summary>
    /// Replace every <c>{{key}}</c> in <paramref name="input"/> in a single
    /// left-to-right scan. The substituted value is never re-fed through
    /// the next iteration, so a kvp value containing <c>{{other}}</c> stays
    /// literal. Tracks which keys actually appeared in the template
    /// (<paramref name="usedKeys"/>) and the total number of substitutions
    /// (<paramref name="bumpReplacements"/>). Placeholders whose name is
    /// not in <paramref name="data"/> are left intact so
    /// <c>ScanUnresolved*</c> can report them.
    /// </summary>
    private static string SinglePassReplace(string input, Dictionary<string, string> data, out bool matched, HashSet<string>? usedKeys = null, Action<int>? bumpReplacements = null)
    {
        matched = false;
        if (string.IsNullOrEmpty(input) || !input.Contains("{{"))
            return input;

        int localCount = 0;
        var result = PlaceholderPattern.Replace(input, m =>
        {
            var key = m.Groups[1].Value;
            if (data.TryGetValue(key, out var replacement))
            {
                usedKeys?.Add(key);
                localCount++;
                return replacement;
            }
            return m.Value;
        });
        if (localCount > 0) bumpReplacements?.Invoke(localCount);
        matched = localCount > 0;
        return result;
    }

    private static List<string> ScanUnresolvedDocx(string filePath)
    {
        var unresolved = new HashSet<string>();
        using var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(filePath, false);
        var body = doc.MainDocumentPart?.Document?.Body;
        if (body == null) return unresolved.ToList();

        foreach (var para in body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Paragraph>())
        {
            var text = string.Concat(para.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().Select(t => t.Text));
            foreach (Match match in PlaceholderPattern.Matches(text))
            {
                unresolved.Add(match.Groups[1].Value);
            }
        }

        // Also scan headers and footers
        var mainPart = doc.MainDocumentPart;
        if (mainPart != null)
        {
            foreach (var headerPart in mainPart.HeaderParts)
            {
                foreach (var para in headerPart.Header?.Descendants<DocumentFormat.OpenXml.Wordprocessing.Paragraph>() ?? Enumerable.Empty<DocumentFormat.OpenXml.Wordprocessing.Paragraph>())
                {
                    var text = string.Concat(para.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().Select(t => t.Text));
                    foreach (Match match in PlaceholderPattern.Matches(text))
                        unresolved.Add(match.Groups[1].Value);
                }
            }
            foreach (var footerPart in mainPart.FooterParts)
            {
                foreach (var para in footerPart.Footer?.Descendants<DocumentFormat.OpenXml.Wordprocessing.Paragraph>() ?? Enumerable.Empty<DocumentFormat.OpenXml.Wordprocessing.Paragraph>())
                {
                    var text = string.Concat(para.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().Select(t => t.Text));
                    foreach (Match match in PlaceholderPattern.Matches(text))
                        unresolved.Add(match.Groups[1].Value);
                }
            }
        }

        return unresolved.OrderBy(x => x).ToList();
    }

    private static MergeResult MergeXlsx(string filePath, Dictionary<string, string> data)
    {
        var usedKeys = new HashSet<string>();
        int totalReplacements = 0;

        using var doc = SpreadsheetDocument.Open(filePath, true);
        var workbookPart = doc.WorkbookPart;
        if (workbookPart == null)
            return new MergeResult(0, new List<string>(), new List<string>());

        // Get shared string table
        var sstPart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
        var sst = sstPart?.SharedStringTable;

        foreach (var worksheetPart in workbookPart.WorksheetParts)
        {
            var sheetData = worksheetPart.Worksheet?.GetFirstChild<SheetData>();
            if (sheetData == null) continue;

            foreach (var row in sheetData.Elements<Row>())
            {
                foreach (var cell in row.Elements<Cell>())
                {
                    var cellText = GetCellText(cell, sst);
                    if (string.IsNullOrEmpty(cellText) || !cellText.Contains("{{")) continue;

                    var newText = SinglePassReplace(cellText, data, out _, usedKeys, count => totalReplacements += count);

                    if (newText != cellText)
                    {
                        SetCellText(cell, newText);
                    }
                }
            }
            worksheetPart.Worksheet?.Save();
        }

        // Scan for unresolved
        var unresolved = ScanUnresolvedXlsx(doc);

        return new MergeResult(totalReplacements, unresolved, usedKeys.ToList());
    }

    private static string GetCellText(Cell cell, SharedStringTable? sst)
    {
        if (cell.DataType?.Value == CellValues.InlineString)
            return cell.InlineString?.InnerText ?? "";

        var value = cell.CellValue?.Text ?? "";

        if (cell.DataType?.Value == CellValues.SharedString && sst != null)
        {
            if (int.TryParse(value, out int idx))
            {
                var item = sst.Elements<SharedStringItem>().ElementAtOrDefault(idx);
                return item?.InnerText ?? value;
            }
        }

        if (cell.DataType?.Value == CellValues.String)
            return value;

        return value;
    }

    private static void SetCellText(Cell cell, string text)
    {
        // Set as inline string to avoid shared string table complexity
        cell.DataType = CellValues.InlineString;
        cell.CellValue = null;
        cell.InlineString = new InlineString(new DocumentFormat.OpenXml.Spreadsheet.Text(text));
    }

    private static List<string> ScanUnresolvedXlsx(SpreadsheetDocument doc)
    {
        var unresolved = new HashSet<string>();
        var workbookPart = doc.WorkbookPart;
        if (workbookPart == null) return unresolved.ToList();

        var sstPart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
        var sst = sstPart?.SharedStringTable;

        foreach (var worksheetPart in workbookPart.WorksheetParts)
        {
            var sheetData = worksheetPart.Worksheet?.GetFirstChild<SheetData>();
            if (sheetData == null) continue;

            foreach (var row in sheetData.Elements<Row>())
            {
                foreach (var cell in row.Elements<Cell>())
                {
                    var text = GetCellText(cell, sst);
                    foreach (Match match in PlaceholderPattern.Matches(text))
                        unresolved.Add(match.Groups[1].Value);
                }
            }
        }

        return unresolved.OrderBy(x => x).ToList();
    }

    private static MergeResult MergePptx(string filePath, Dictionary<string, string> data)
    {
        var usedKeys = new HashSet<string>();
        int totalReplacements = 0;

        using var doc = PresentationDocument.Open(filePath, true);
        var presentationPart = doc.PresentationPart;
        if (presentationPart == null)
            return new MergeResult(0, new List<string>(), new List<string>());

        foreach (var slidePart in presentationPart.SlideParts)
        {
            // Process shapes on slide
            var shapeTree = slidePart.Slide?.CommonSlideData?.ShapeTree;
            if (shapeTree != null)
            {
                foreach (var shape in shapeTree.Elements<Shape>())
                {
                    totalReplacements += ReplaceInTextBody(shape.TextBody, data, usedKeys);
                }
            }

            // Process notes
            var notesPart = slidePart.NotesSlidePart;
            if (notesPart != null)
            {
                var notesShapeTree = notesPart.NotesSlide?.CommonSlideData?.ShapeTree;
                if (notesShapeTree != null)
                {
                    foreach (var shape in notesShapeTree.Elements<Shape>())
                    {
                        totalReplacements += ReplaceInTextBody(shape.TextBody, data, usedKeys);
                    }
                }
                notesPart.NotesSlide?.Save();
            }

            slidePart.Slide?.Save();
        }

        // Scan for unresolved
        var unresolved = ScanUnresolvedPptx(doc);

        return new MergeResult(totalReplacements, unresolved, usedKeys.ToList());
    }

    private static int ReplaceInTextBody(OpenXmlElement? textBody, Dictionary<string, string> data, HashSet<string> usedKeys)
    {
        if (textBody == null) return 0;
        int replacements = 0;

        foreach (var para in textBody.Elements<Drawing.Paragraph>())
        {
            replacements += ReplaceInParagraph(para, data, usedKeys);
        }

        return replacements;
    }

    /// <summary>
    /// Replace placeholders in a Drawing.Paragraph. Handles text split across multiple runs
    /// by concatenating run text, finding placeholders, and rebuilding runs.
    /// </summary>
    private static int ReplaceInParagraph(Drawing.Paragraph para, Dictionary<string, string> data, HashSet<string> usedKeys)
    {
        var runs = para.Elements<Drawing.Run>().ToList();
        if (runs.Count == 0) return 0;

        // Concatenate all run text
        var fullText = string.Concat(runs.Select(r => r.Text?.Text ?? ""));
        if (!fullText.Contains("{{")) return 0;

        int replacements = 0;
        var newText = SinglePassReplace(fullText, data, out _, usedKeys, count => replacements += count);

        if (replacements == 0) return 0;

        // Replace: keep first run with new text and its formatting, remove the rest
        var firstRun = runs[0];
        if (firstRun.Text == null)
            firstRun.Text = new Drawing.Text(newText);
        else
            firstRun.Text.Text = newText;

        // Remove remaining runs
        for (int i = 1; i < runs.Count; i++)
        {
            runs[i].Remove();
        }

        return replacements;
    }

    private static List<string> ScanUnresolvedPptx(PresentationDocument doc)
    {
        var unresolved = new HashSet<string>();
        var presentationPart = doc.PresentationPart;
        if (presentationPart == null) return unresolved.ToList();

        foreach (var slidePart in presentationPart.SlideParts)
        {
            var shapeTree = slidePart.Slide?.CommonSlideData?.ShapeTree;
            if (shapeTree != null)
            {
                foreach (var shape in shapeTree.Elements<Shape>())
                {
                    ScanTextBody(shape.TextBody, unresolved);
                }
            }

            var notesPart = slidePart.NotesSlidePart;
            if (notesPart != null)
            {
                var notesShapeTree = notesPart.NotesSlide?.CommonSlideData?.ShapeTree;
                if (notesShapeTree != null)
                {
                    foreach (var shape in notesShapeTree.Elements<Shape>())
                    {
                        ScanTextBody(shape.TextBody, unresolved);
                    }
                }
            }
        }

        return unresolved.OrderBy(x => x).ToList();
    }

    private static void ScanTextBody(OpenXmlElement? textBody, HashSet<string> unresolved)
    {
        if (textBody == null) return;

        foreach (var para in textBody.Elements<Drawing.Paragraph>())
        {
            var text = string.Concat(para.Elements<Drawing.Run>().Select(r => r.Text?.Text ?? ""));
            foreach (Match match in PlaceholderPattern.Matches(text))
                unresolved.Add(match.Groups[1].Value);
        }
    }
}
