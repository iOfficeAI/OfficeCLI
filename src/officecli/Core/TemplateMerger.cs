// Copyright 2025 OfficeCli (officecli.ai)
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
public static class TemplateMerger
{
    private static readonly Regex PlaceholderPattern = new(@"\{\{(\w[\w.]*)\}\}", RegexOptions.Compiled);

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
        foreach (var kvp in jsonObj)
        {
            data[kvp.Key] = kvp.Value?.ToString() ?? "";
        }
        return data;
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
        using (var handler = new Handlers.WordHandler(filePath, editable: true))
        {
            foreach (var kvp in data)
            {
                var placeholder = "{{" + kvp.Key + "}}";
                handler.Set("/", new Dictionary<string, string>
                {
                    ["find"] = placeholder,
                    ["replace"] = kvp.Value
                });
            }
        }

        // Scan for unresolved placeholders
        var unresolved = ScanUnresolvedDocx(filePath);
        // Keys that were provided and are not still unresolved were successfully replaced
        var usedKeys = data.Keys.Where(k => !unresolved.Contains(k)).ToList();

        return new MergeResult(usedKeys.Count, unresolved, usedKeys);
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
                foreach (var para in headerPart.Header.Descendants<DocumentFormat.OpenXml.Wordprocessing.Paragraph>())
                {
                    var text = string.Concat(para.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().Select(t => t.Text));
                    foreach (Match match in PlaceholderPattern.Matches(text))
                        unresolved.Add(match.Groups[1].Value);
                }
            }
            foreach (var footerPart in mainPart.FooterParts)
            {
                foreach (var para in footerPart.Footer.Descendants<DocumentFormat.OpenXml.Wordprocessing.Paragraph>())
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

                    var newText = cellText;
                    foreach (var kvp in data)
                    {
                        var placeholder = "{{" + kvp.Key + "}}";
                        if (newText.Contains(placeholder))
                        {
                            newText = newText.Replace(placeholder, kvp.Value);
                            usedKeys.Add(kvp.Key);
                            totalReplacements++;
                        }
                    }

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

        var newText = fullText;
        int replacements = 0;

        foreach (var kvp in data)
        {
            var placeholder = "{{" + kvp.Key + "}}";
            if (newText.Contains(placeholder))
            {
                newText = newText.Replace(placeholder, kvp.Value);
                usedKeys.Add(kvp.Key);
                replacements++;
            }
        }

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
