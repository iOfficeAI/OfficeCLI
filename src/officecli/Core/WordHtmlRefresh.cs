// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml.Packaging;

namespace OfficeCli.Core;

/// <summary>
/// HTML-based refresh fallback. Mirrors the TOC update pipeline
/// but uses the browser's pagination instead of Word's layout engine —
/// page numbers may differ from what F9 in Word would produce, but the
/// values are internally consistent with officecli's own HTML preview.
/// </summary>
internal static class WordHtmlRefresh
{
    public static bool RefreshViaHtml(string docx)
    {
        try
        {
            string htmlSnapshot;
            using (var doc = WordprocessingDocument.Open(docx, isEditable: true))
            {
                WordTocBuilder.RegenerateAllTocs(doc);
                doc.MainDocumentPart!.Document!.Save();
            }

            using (var handler = (Handlers.WordHandler)Handlers.DocumentHandlerFactory.Open(docx, editable: false))
                htmlSnapshot = handler.ViewAsHtml(null);

            var tmpHtml = Path.Combine(Path.GetTempPath(), $"officecli_refresh_{Guid.NewGuid():N}.html");
            HtmlScreenshot.PaginationResult? pagination;
            try
            {
                File.WriteAllText(tmpHtml, htmlSnapshot);
                pagination = HtmlScreenshot.GetPaginationFromDom(tmpHtml);
            }
            finally { try { File.Delete(tmpHtml); } catch { } }

            if (pagination == null) return false;

            using (var doc = WordprocessingDocument.Open(docx, isEditable: true))
            {
                ApplyPageNumbers(doc, pagination.AnchorPageMap);
                doc.MainDocumentPart!.Document!.Save();

                var part = doc.ExtendedFilePropertiesPart ?? doc.AddExtendedFilePropertiesPart();
                if (part.Properties == null)
                    part.Properties = new DocumentFormat.OpenXml.ExtendedProperties.Properties();
                if (part.Properties.Pages == null)
                    part.Properties.Pages = new DocumentFormat.OpenXml.ExtendedProperties.Pages();
                part.Properties.Pages.Text = pagination.TotalPages.ToString();
                part.Properties.Save();
            }
            return true;
        }
        catch { return false; }
    }

    static void ApplyPageNumbers(WordprocessingDocument doc, Dictionary<string, int> map)
    {
        var body = doc.MainDocumentPart?.Document?.Body;
        if (body == null) return;
        // Walk all PAGEREF fields. The instr text " PAGEREF _TocXXX \h "
        // identifies the bookmark; the very next Run after the separate
        // fldChar holds the cached page number Text we want to rewrite.
        foreach (var p in body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Paragraph>())
        {
            DocumentFormat.OpenXml.Wordprocessing.FieldCode? instr = null;
            foreach (var r in p.Descendants<DocumentFormat.OpenXml.Wordprocessing.Run>())
            {
                var fc = r.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.FieldChar>();
                if (fc?.FieldCharType?.Value == DocumentFormat.OpenXml.Wordprocessing.FieldCharValues.Begin)
                {
                    instr = null;
                }
                else if (instr == null)
                {
                    var ic = r.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.FieldCode>();
                    if (ic != null && ic.Text != null && ic.Text.TrimStart().StartsWith("PAGEREF", StringComparison.OrdinalIgnoreCase))
                        instr = ic;
                }
                else if (fc?.FieldCharType?.Value == DocumentFormat.OpenXml.Wordprocessing.FieldCharValues.Separate)
                {
                    var resultRun = r.NextSibling<DocumentFormat.OpenXml.Wordprocessing.Run>();
                    if (resultRun != null)
                    {
                        var anchor = ExtractPagerefAnchor(instr.Text!);
                        if (anchor != null && map.TryGetValue(anchor, out var pgNum))
                        {
                            var t = resultRun.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Text>();
                            if (t != null) t.Text = pgNum.ToString();
                        }
                    }
                    instr = null;
                }
            }
        }
    }

    static string? ExtractPagerefAnchor(string instrText)
    {
        var m = System.Text.RegularExpressions.Regex.Match(instrText, @"PAGEREF\s+(\S+)");
        return m.Success ? m.Groups[1].Value : null;
    }
}
