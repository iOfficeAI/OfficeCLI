// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
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
        // The TOC regeneration below writes the package through the SDK, and it
        // runs before we know whether a pagination backend exists at all — so a
        // refresh that ends up reporting failure used to leave a half-updated
        // document behind: the TOC entries were expanded (with the placeholder
        // page number 0 the caller is supposed to overwrite) and fresh bookmarks
        // were inserted, all after a 1-exit-code "refresh failed". Keep the bytes
        // as opened and put them back on every failure path, so a failed refresh
        // is a no-op on disk — the contract ElementRollback already holds for a
        // failed element-level Set/Add.
        var original = TryReadAllBytes(docx);
        try
        {
            string htmlSnapshot;
            using (var doc = WordprocessingDocument.Open(docx, isEditable: true))
            {
                WordTocBuilder.RegenerateAllTocs(doc);
                doc.MainDocumentPart!.Document!.Save();
            }

            using (var handler = Handlers.DocumentHandlerFactory.Open(docx, editable: false))
            {
                Handlers.Rendering.RenderingBootstrap.EnsureRegistered();
                var renderer = Rendering.RendererRegistry.Default.Resolve(
                    "docx", Rendering.RenderOutputKind.Html, Rendering.RenderMode.Static);
                htmlSnapshot = renderer!.Render(
                    new Handlers.Rendering.HandlerRenderInput(handler, "docx"),
                    new Rendering.RenderOptions()).Text!;
            }

            var tmpHtml = Path.Combine(Path.GetTempPath(), $"officecli_refresh_{Guid.NewGuid():N}.html");
            HtmlScreenshot.PaginationResult? pagination;
            try
            {
                File.WriteAllText(tmpHtml, htmlSnapshot);
                pagination = HtmlScreenshot.GetPaginationFromDom(tmpHtml);
            }
            finally { try { File.Delete(tmpHtml); } catch { } }

            if (pagination == null)
            {
                RestoreOriginal(docx, original);
                return false;
            }

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
        catch
        {
            RestoreOriginal(docx, original);
            return false;
        }
    }

    /// <summary>Bytes of the package as opened, or null when they could not be
    /// read — there is then nothing to put back, which is the pre-fix behaviour.</summary>
    static byte[]? TryReadAllBytes(string docx)
    {
        try { return File.ReadAllBytes(docx); } catch { return null; }
    }

    /// <summary>Put the snapshot back after a failed refresh. Best-effort: a
    /// restore that fails must not turn "refresh failed" into a thrown error,
    /// and every failure exit reports false either way.</summary>
    static void RestoreOriginal(string docx, byte[]? original)
    {
        if (original == null) return;
        try { File.WriteAllBytes(docx, original); } catch { }
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
