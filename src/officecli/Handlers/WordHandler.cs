// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class WordHandler : IDocumentHandler
{
    private readonly WordprocessingDocument _doc;
    private readonly string _filePath;
    private HashSet<string> _usedParaIds = new(StringComparer.OrdinalIgnoreCase);
    private int _nextParaId = 0x100000;
    public int LastFindMatchCount { get; internal set; }

    /// <summary>
    /// Props that the most recent Add() call could not consume. Surfaced to
    /// the CLI layer so silent-drops on the curated surface (e.g.
    /// `add /styles --prop font.eastAsia=...`) become visible warnings
    /// instead of "Added" lies. Reset at the start of each Add.
    /// </summary>
    public List<string> LastAddUnsupportedProps { get; internal set; } = new();

    public WordHandler(string filePath, bool editable)
    {
        _filePath = filePath;
        _doc = WordprocessingDocument.Open(filePath, editable);
        WordStrictAttributeSanitizer.Sanitize(_doc);
        if (editable)
        {
            EnsureAllParaIds();
            EnsureDocPropIds();
        }
    }

    /// <summary>
    /// Resolve a picture-run path to the embedded image's bytes and content
    /// type. Returns null if the path doesn't point at a Drawing-bearing
    /// run, or the run carries no resolvable rId/embed target.
    ///
    /// <para>
    /// Used by <c>BatchEmitter</c> to round-trip pictures through batch
    /// dumps — the bytes are encoded as a data URI in the emitted
    /// `src=` prop and re-imported via <c>ImageSource.Resolve</c> on replay.
    /// </para>
    /// </summary>
    /// <summary>
    /// Returns true if the run at <paramref name="runPath"/> wraps a chart
    /// (c:chart inside a Drawing's graphicData). BatchEmitter uses this to
    /// distinguish chart-bearing runs from picture/OLE/background runs that
    /// also surface as type="picture" in Get — without this, an unsupported
    /// drawing's failed image extraction would consume the next chart spec
    /// and render at the wrong paragraph.
    /// </summary>
    public bool IsChartRun(string runPath)
    {
        var segments = ParsePath(runPath);
        var element = NavigateToElement(segments);
        if (element is not Run run) return false;
        var drawing = run.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Drawing>();
        if (drawing == null) return false;
        return drawing
            .Descendants<DocumentFormat.OpenXml.Drawing.Charts.ChartReference>()
            .Any();
    }

    /// <summary>
    /// Outer XML of the element at <paramref name="path"/>. BatchEmitter
    /// uses this as a raw-XML fallback for content that has no typed Add
    /// path — wps:wsp background shapes being the motivating case. Returns
    /// null if the path doesn't resolve.
    /// </summary>
    public string? GetElementXml(string path)
    {
        try
        {
            var segments = ParsePath(path);
            var element = NavigateToElement(segments);
            return element?.OuterXml;
        }
        catch
        {
            return null;
        }
    }

    public (byte[] Bytes, string ContentType)? GetImageBinary(string runPath)
    {
        // Parse + navigate via the same machinery Get/Set use so paraId
        // anchors and positional indices behave consistently.
        var segments = ParsePath(runPath);
        var element = NavigateToElement(segments);
        if (element is not Run run) return null;

        var drawing = run.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Drawing>();
        if (drawing == null) return null;

        var blip = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().FirstOrDefault();
        var embedId = blip?.Embed?.Value;
        if (string.IsNullOrEmpty(embedId)) return null;

        // CONSISTENCY(host-part-rel): mirror the AddPicture host-part lookup
        // — image part may be attached to a header/footer part rather than
        // the main document part, depending on where the run lives.
        var hostPart = ResolveImageHostPart(run);
        try
        {
            var part = hostPart.GetPartById(embedId);
            using var src = part.GetStream();
            using var ms = new MemoryStream();
            src.CopyTo(ms);
            return (ms.ToArray(), part.ContentType);
        }
        catch
        {
            return null;
        }
    }

    private OpenXmlPart ResolveImageHostPart(Run run)
    {
        var headerAncestor = run.Ancestors<Header>().FirstOrDefault();
        if (headerAncestor != null)
        {
            var hp = _doc.MainDocumentPart!.HeaderParts
                .FirstOrDefault(p => ReferenceEquals(p.Header, headerAncestor));
            if (hp != null) return hp;
        }
        var footerAncestor = run.Ancestors<Footer>().FirstOrDefault();
        if (footerAncestor != null)
        {
            var fp = _doc.MainDocumentPart!.FooterParts
                .FirstOrDefault(p => ReferenceEquals(p.Footer, footerAncestor));
            if (fp != null) return fp;
        }
        return _doc.MainDocumentPart!;
    }

    // ==================== Raw Layer ====================

    public string Raw(string partPath, int? startRow = null, int? endRow = null, HashSet<string>? cols = null)
    {
        var mainPart = _doc.MainDocumentPart;
        if (mainPart == null) return "(no main part)";

        return partPath.ToLowerInvariant() switch
        {
            "/document" or "/word/document.xml" => mainPart.Document?.OuterXml ?? "",
            "/styles" or "/word/styles.xml" => mainPart.StyleDefinitionsPart?.Styles?.OuterXml ?? "(no styles)",
            "/settings" or "/word/settings.xml" => mainPart.DocumentSettingsPart?.Settings?.OuterXml ?? "(no settings)",
            "/numbering" or "/word/numbering.xml" => mainPart.NumberingDefinitionsPart?.Numbering?.OuterXml ?? "(no numbering)",
            "/comments" => mainPart.WordprocessingCommentsPart?.Comments?.OuterXml ?? "(no comments)",
            "/theme" or "/word/theme/theme1.xml" => mainPart.ThemePart?.Theme?.OuterXml ?? "(no theme)",
            _ when partPath.StartsWith("/header") => GetHeaderRawXml(partPath),
            _ when partPath.StartsWith("/footer") => GetFooterRawXml(partPath),
            _ when partPath.StartsWith("/chart") => GetChartRawXml(partPath),
            _ => throw new ArgumentException($"Unknown part: {partPath}. Available: /document, /styles, /settings, /numbering, /comments, /theme, /header[n], /footer[n], /chart[n]")
        };
    }

    public void RawSet(string partPath, string xpath, string action, string? xml)
    {
        var mainPart = _doc.MainDocumentPart
            ?? throw new InvalidOperationException("No main document part");

        OpenXmlPartRootElement rootElement;
        var lowerPath = partPath.ToLowerInvariant();

        if (lowerPath is "/document" or "/")
            rootElement = mainPart.Document ?? throw new InvalidOperationException("No document");
        else if (lowerPath is "/styles")
            rootElement = mainPart.StyleDefinitionsPart?.Styles ?? throw new InvalidOperationException("No styles part");
        else if (lowerPath is "/settings")
            rootElement = mainPart.DocumentSettingsPart?.Settings ?? throw new InvalidOperationException("No settings part");
        else if (lowerPath is "/numbering")
        {
            // CONSISTENCY(raw-set-create-missing-part): see /theme branch.
            var numPart = mainPart.NumberingDefinitionsPart ?? mainPart.AddNewPart<NumberingDefinitionsPart>();
            if (numPart.Numbering == null)
            {
                numPart.Numbering = new Numbering();
                numPart.Numbering.Save();
            }
            rootElement = numPart.Numbering;
        }
        else if (lowerPath is "/comments")
            rootElement = mainPart.WordprocessingCommentsPart?.Comments ?? throw new InvalidOperationException("No comments part");
        else if (lowerPath is "/theme")
        {
            // CONSISTENCY(raw-set-create-missing-part): blank docs created via
            // BlankDocCreator have no ThemePart; dump→batch round-trip from a
            // real Word/python-docx file emits raw-set /theme replace which
            // would otherwise abort the whole batch. Lazily add the theme part
            // and an empty <a:theme> root so RawXmlHelper.Execute can match
            // /a:theme and replace it with the dumped XML.
            var themePart = mainPart.ThemePart ?? mainPart.AddNewPart<ThemePart>();
            if (themePart.Theme == null)
            {
                themePart.Theme = new DocumentFormat.OpenXml.Drawing.Theme(
                    new DocumentFormat.OpenXml.Drawing.ThemeElements());
                themePart.Theme.Save();
            }
            rootElement = themePart.Theme;
        }
        else if (lowerPath.StartsWith("/header"))
        {
            var idx = 0;
            var bracketIdx = partPath.IndexOf('[');
            if (bracketIdx >= 0)
                int.TryParse(partPath[(bracketIdx + 1)..].TrimEnd(']'), out idx);
            var headerPart = mainPart.HeaderParts.ElementAtOrDefault(idx - 1)
                ?? throw new ArgumentException($"header[{idx}] not found");
            rootElement = headerPart.Header ?? throw new InvalidOperationException($"Corrupt file: header[{idx}] data missing");
        }
        else if (lowerPath.StartsWith("/footer"))
        {
            var idx = 0;
            var bracketIdx = partPath.IndexOf('[');
            if (bracketIdx >= 0)
                int.TryParse(partPath[(bracketIdx + 1)..].TrimEnd(']'), out idx);
            var footerPart = mainPart.FooterParts.ElementAtOrDefault(idx - 1)
                ?? throw new ArgumentException($"footer[{idx}] not found");
            rootElement = footerPart.Footer ?? throw new InvalidOperationException($"Corrupt file: footer[{idx}] data missing");
        }
        else if (lowerPath.StartsWith("/chart"))
        {
            var idx = 0;
            var bracketIdx = partPath.IndexOf('[');
            if (bracketIdx >= 0)
                int.TryParse(partPath[(bracketIdx + 1)..].TrimEnd(']'), out idx);
            var chartPart = mainPart.ChartParts.ElementAtOrDefault(idx - 1)
                ?? throw new ArgumentException($"chart[{idx}] not found");
            rootElement = chartPart.ChartSpace ?? throw new InvalidOperationException($"Corrupt file: chart[{idx}] data missing");
        }
        else
            throw new ArgumentException($"Unknown part: {partPath}. Available: /document, /styles, /settings, /numbering, /header[n], /footer[n], /chart[n]");

        var affected = RawXmlHelper.Execute(rootElement, xpath, action, xml);
        rootElement.Save();
        // CONSISTENCY(paraid-global-uniqueness): RawSet may inject paragraphs
        // carrying paraIds the handler hasn't seen — without re-scanning,
        // _usedParaIds and _nextParaId stay stale and the next AddBreak /
        // AddParagraph could allocate a colliding paraId. Especially
        // dangerous in resident mode where one process serves many commands
        // across the same _usedParaIds set. Re-run EnsureAllParaIds after
        // every successful raw mutation so the global pool stays accurate.
        EnsureAllParaIds();
        // BUG-R5-01: do not emit chatter from inside the handler — the CLI
        // wrappers (CommandBuilder.Raw raw-set + batch run raw-set) print
        // their own structured message. Writing here pollutes batch --json
        // output (extra stdout lines escaped into result.message strings).
        _ = affected;
    }

    public List<ValidationError> Validate() => RawXmlHelper.ValidateDocument(_doc);

    public void Dispose()
    {
        _doc.Dispose();
        // CONSISTENCY(word-self-close): the OpenXml SDK serializes empty
        // elements with a space before the self-close (`<w:br />`). Several
        // downstream consumers (and test regexes) look for the canonical
        // `<w:br/>` / `<w:tab/>` form. Normalize the persisted document.xml
        // in place so the saved package matches the canonical short form.
        // Only applied to word/document.xml; styles/settings/numbering are
        // left untouched since the space form is schema-equivalent.
        try { NormalizeSelfClosingInDocx(_filePath); } catch { /* best-effort */ }
    }

    private static void NormalizeSelfClosingInDocx(string path)
    {
        if (!System.IO.File.Exists(path)) return;
        using var fs = new System.IO.FileStream(path, System.IO.FileMode.Open, System.IO.FileAccess.ReadWrite);
        using var za = new System.IO.Compression.ZipArchive(fs, System.IO.Compression.ZipArchiveMode.Update, leaveOpen: false);
        var entry = za.GetEntry("word/document.xml");
        if (entry == null) return;
        string xml;
        using (var rs = entry.Open())
        using (var sr = new System.IO.StreamReader(rs))
            xml = sr.ReadToEnd();
        // Collapse "<w:br />" → "<w:br/>" and "<w:tab />" → "<w:tab/>"
        // (no-attribute empty elements only).
        var normalized = System.Text.RegularExpressions.Regex.Replace(
            xml, @"<w:(br|tab) />", "<w:$1/>");
        if (normalized == xml) return;
        entry.Delete();
        var newEntry = za.CreateEntry("word/document.xml");
        using var ws = newEntry.Open();
        using var sw = new System.IO.StreamWriter(ws, new System.Text.UTF8Encoding(false));
        sw.Write(normalized);
    }

    // (private helpers, navigation, selector, style/list, image helpers moved to Word/ partial files)
}
