// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    public string Add(string parentPath, string type, InsertPosition? position, Dictionary<string, string> properties)
    {
        // CONSISTENCY(prop-key-case): property keys are case-insensitive
        // ("SRC"/"src"/"Src" all resolve the same). Normalize once at the
        // dispatch entry so every AddXxx helper can rely on TryGetValue("src").
        properties = properties == null
            ? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            : properties.Comparer == StringComparer.OrdinalIgnoreCase
                ? properties
                : new Dictionary<string, string>(properties, StringComparer.OrdinalIgnoreCase);

        // Reset per-Add diagnostic. Helpers that detect silent-drop props
        // (currently only AddStyle) populate this; the CLI layer surfaces
        // it as a WARNING line so curated-surface gaps stop being silent.
        LastAddUnsupportedProps = new List<string>();

        // Reject negative --index up front with a clean message instead of
        // letting it fall through and surface as a raw .NET
        // ArgumentOutOfRangeException from collection indexing. Applies to
        // every parent (/body, /styles, /header[N], ...).
        if (position?.Index.HasValue == true && position.Index.Value < 0)
            throw new ArgumentException("--index must be non-negative.");

        var body = _doc.MainDocumentPart?.Document?.Body
            ?? throw new InvalidOperationException("Document body not found");

        OpenXmlElement parent;
        if (parentPath is "/" or "" or "/body")
        {
            parent = body;
            parentPath = "/body";
        }
        else if (parentPath == "/styles")
        {
            var stylesPart = _doc.MainDocumentPart!.StyleDefinitionsPart
                ?? _doc.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles ??= new Styles();
            parent = stylesPart.Styles;
        }
        else if (parentPath == "/numbering")
        {
            var numberingPart = _doc.MainDocumentPart!.NumberingDefinitionsPart
                ?? _doc.MainDocumentPart.AddNewPart<NumberingDefinitionsPart>();
            numberingPart.Numbering ??= new Numbering();
            parent = numberingPart.Numbering;
        }
        else if (TryResolveFootnoteOrEndnoteBody(parentPath, out var fnBody, out var canonicalPath))
        {
            // Route /footnote[@footnoteId=N] / /footnote[N] (and endnote
            // equivalents) to the footnote/endnote element itself so block-
            // level adds (paragraph, run, ...) land inside its body.
            parent = fnBody!;
            parentPath = canonicalPath!;
        }
        else
        {
            List<PathSegment> parts;
            try
            {
                parts = ParsePath(parentPath);
            }
            catch (Exception ex) when (ex is not ArgumentException and not InvalidOperationException)
            {
                throw new ArgumentException($"Malformed parent path '{parentPath}'. Check selector brackets and escape sequences.", ex);
            }
            parent = NavigateToElement(parts, out var ctx)
                ?? throw new ArgumentException($"Path not found: {parentPath}" + (ctx != null ? $". {ctx}" : ""));
        }

        // Reject add operations whose parent/child combination would produce
        // schema-invalid OOXML (e.g. /body/sectPr accepting a paragraph child,
        // or /body/p[N] accepting a nested paragraph/table). `position` is
        // passed because some parent/child combos are legal *only* with a
        // specific anchor form — notably block-level adds under a paragraph
        // parent via `find:` (the paragraph is split and the block is
        // promoted to a body-level sibling between the halves).
        ValidateParentChild(parent, parentPath, type, position);

        int? index;
        try
        {
            // Resolve --after/--before to index (handles find: prefix for text-based anchoring)
            index = ResolveAnchorPosition(parent, parentPath, position);
        }
        catch (ArgumentOutOfRangeException ex)
        {
            throw new ArgumentException($"Invalid anchor for --after/--before. Check selector syntax (e.g. p[2], r[@paraId=...]).", ex);
        }
        catch (Exception ex) when (ex is not ArgumentException and not InvalidOperationException)
        {
            throw new ArgumentException($"Invalid anchor for --after/--before: {ex.GetType().Name}. Check selector syntax.", ex);
        }

        // Handle find: prefix — text-based anchoring
        if (index == FindAnchorIndex && position != null)
        {
            var anchorValue = (position.After ?? position.Before)!;
            var findValue = anchorValue["find:".Length..]; // strip "find:" prefix
            var isAfter = position.After != null;
            return AddAtFindPosition(parent, parentPath, type, findValue, isAfter, null, properties);
        }

        string resultPath;
        try
        {
        resultPath = type.ToLowerInvariant() switch
        {
            "paragraph" or "p" => AddParagraph(parent, parentPath, index, properties),
            "equation" or "formula" or "math" => AddEquation(parent, parentPath, index, properties),
            "run" or "r" => AddRun(parent, parentPath, index, properties),
            "table" or "tbl" => AddTable(parent, parentPath, index, properties),
            "row" or "tr" => AddRow(parent, parentPath, index, properties),
            "cell" or "tc" => AddCell(parent, parentPath, index, properties),
            "tab" or "tabstop" => AddTab(parent, parentPath, index, properties),
            "chart" => AddChart(parent, parentPath, index, properties),
            "picture" or "image" or "img" => AddPicture(parent, parentPath, index, properties),
            "ole" or "oleobject" or "object" or "embed" => AddOle(parent, parentPath, index, properties),
            "comment" => AddComment(parent, parentPath, index, properties),
            "bookmark" => AddBookmark(parent, parentPath, index, properties),
            "hyperlink" or "link" => AddHyperlink(parent, parentPath, index, properties),
            "section" or "sectionbreak" => AddSection(parent, parentPath, index, properties),
            "footnote" => AddFootnote(parent, parentPath, index, properties),
            "endnote" => AddEndnote(parent, parentPath, index, properties),
            "toc" or "tableofcontents" => AddToc(parent, parentPath, index, properties),
            "style" => AddStyle(parent, parentPath, index, properties),
            "num" => AddNum(parent, parentPath, index, properties),
            "abstractnum" => AddAbstractNum(parent, parentPath, index, properties),
            "lvl" or "level" => AddLvl(parent, parentPath, index, properties),
            "header" => AddHeader(parent, parentPath, index, properties),
            "footer" => AddFooter(parent, parentPath, index, properties),
            "field" or "pagenum" or "pagenumber" or "page" or "numpages" or "sectionpages" or "section"
                or "date" or "createdate" or "savedate" or "printdate" or "edittime" or "time"
                or "author" or "lastsavedby" or "title" or "subject" or "filename"
                or "numwords" or "numchars" or "revnum" or "template" or "comments" or "doccomments" or "keywords"
                or "mergefield" or "ref" or "pageref" or "noteref" or "seq" or "styleref" or "docproperty" or "if"
                => AddField(parent, parentPath, index, properties, type),
            "pagebreak" or "columnbreak" or "break" => AddBreak(parent, parentPath, index, properties, type),
            "sdt" or "contentcontrol" => AddSdt(parent, parentPath, index, properties),
            "watermark" => AddWatermark(parent, parentPath, index, properties),
            "formfield" => AddFormField(parent, parentPath, index, properties),
            // Reject tracked-revision element types. Falling through to
            // AddDefault produces schema-invalid XML (unnamespaced attrs —
            // OOXML needs w:author/w:id/w:date) and, without --index,
            // clobbers the target paragraph's existing runs (data loss).
            // There is also no way to express the required <w:r><w:t>
            // content via --prop. Revisions are authored by word processors
            // with track-changes enabled; route users back to the normal
            // inline add flow. Mirrors footnote/endnote/comment rejection
            // added in round 6.
            "ins" or "del" or "moveto" or "movefrom" =>
                throw new ArgumentException(
                    $"Cannot add '{type}' directly. Tracked revisions (<w:ins>/<w:del>/<w:moveTo>/<w:moveFrom>) are authored by word processors with track-changes enabled. To insert content that reviewers see as a tracked change, add the run normally (--type run --prop text=...) and enable track-changes in Word."),
            _ => AddDefault(parent, parentPath, index, properties, type),
        };
        }
        catch (ArgumentOutOfRangeException ex)
        {
            // Surface as a clean ArgumentException (CLI layer formats Message).
            // Scrub the raw .NET parameter noise.
            throw new ArgumentException($"Invalid index or anchor for add '{type}'. Check --index / --after / --before values.", ex);
        }

        _doc.MainDocumentPart?.Document?.Save();
        return resultPath;
    }

    /// <summary>
    /// Resolve a top-level /footnote[...] or /endnote[...] path to the
    /// corresponding Footnote/Endnote element (so block-level adds land in
    /// its content). Returns false for anything else. Supports the two
    /// emitted predicate shapes: [@footnoteId=N]/[@endnoteId=N] and [N].
    /// </summary>
    private bool TryResolveFootnoteOrEndnoteBody(string parentPath, out OpenXmlElement? fnBody, out string? canonicalPath)
    {
        fnBody = null;
        canonicalPath = null;

        var fnMatch = System.Text.RegularExpressions.Regex.Match(
            parentPath, @"^/footnote\[(?:@footnoteId=)?(\d+)\]$");
        if (fnMatch.Success)
        {
            var fnId = int.Parse(fnMatch.Groups[1].Value);
            var fn = _doc.MainDocumentPart?.FootnotesPart?.Footnotes?
                .Elements<Footnote>().FirstOrDefault(f => f.Id?.Value == fnId);
            if (fn == null)
                throw new ArgumentException($"Footnote {fnId} not found");
            fnBody = fn;
            canonicalPath = $"/footnote[@footnoteId={fnId}]";
            return true;
        }

        var enMatch = System.Text.RegularExpressions.Regex.Match(
            parentPath, @"^/endnote\[(?:@endnoteId=)?(\d+)\]$");
        if (enMatch.Success)
        {
            var enId = int.Parse(enMatch.Groups[1].Value);
            var en = _doc.MainDocumentPart?.EndnotesPart?.Endnotes?
                .Elements<Endnote>().FirstOrDefault(e => e.Id?.Value == enId);
            if (en == null)
                throw new ArgumentException($"Endnote {enId} not found");
            fnBody = en;
            canonicalPath = $"/endnote[@endnoteId={enId}]";
            return true;
        }

        return false;
    }

    /// <summary>
    /// Reject add operations whose parent/child combination would produce
    /// schema-invalid OOXML. Keeps validation cheap: just the handful of
    /// categories that corrupt documents silently.
    /// </summary>
    private static void ValidateParentChild(OpenXmlElement parent, string parentPath, string type, InsertPosition? position = null)
    {
        var t = type?.ToLowerInvariant() ?? "";
        // `find:` anchors on block-level types under a paragraph parent are
        // legal: AddAtFindPosition splits the paragraph at the anchor and
        // promotes the block to a body-level sibling between the halves.
        // This matches Word's native "cursor mid-sentence → Insert → Table"
        // behavior. Same latitude for section/toc.
        bool isFindAnchor =
            position != null &&
            ((position.After?.StartsWith("find:", StringComparison.Ordinal) ?? false)
             || (position.Before?.StartsWith("find:", StringComparison.Ordinal) ?? false));

        // /body/sectPr cannot contain added children via `add` — the section
        // element only holds layout primitives (pgSz, pgMar, cols, ...), all
        // of which are managed via `set` on /body/sectPr instead.
        if (parent is SectionProperties)
        {
            throw new ArgumentException(
                $"Cannot add '{type}' under {parentPath}. SectionProperties only holds layout metadata; use 'officecli set' to modify pgSz, pgMar, cols, etc.");
        }

        if (parent is Paragraph)
        {
            // Block-level constructs can't nest inside a paragraph — unless
            // the caller used a `find:` anchor, in which case AddAtFindPosition
            // splits the paragraph and promotes the block to a body sibling.
            switch (t)
            {
                case "paragraph":
                case "p":
                case "table":
                case "tbl":
                case "section":
                case "sectionbreak":
                case "toc":
                case "tableofcontents":
                    if (isFindAnchor) break;
                    throw new ArgumentException(
                        $"Cannot add '{type}' under {parentPath}: a paragraph cannot contain another paragraph, table, section break, or TOC. Add at /body instead, or use --after/--before find:<text> to split this paragraph at the anchor.");
                case "sectpr":
                    // Raw <w:sectPr> as a direct child of <w:p> is schema-invalid.
                    // sectPr may only live inside <w:pPr> (paragraph-level break)
                    // or at the end of <w:body> (document final section).
                    // Block `--from` clones that would produce <w:p><w:sectPr/></w:p>.
                    throw new ArgumentException(
                        $"Cannot add '{type}' under {parentPath}: raw <w:sectPr> cannot be a direct child of a paragraph (it must live inside <w:pPr>). Use `--type section` to create a proper paragraph-level section break.");
            }
        }

        if (parent is Body)
        {
            switch (t)
            {
                case "row":
                case "tr":
                    throw new ArgumentException(
                        $"Cannot add '{type}' under {parentPath}: rows must be added under a table (/body/tbl[N]).");
                case "cell":
                case "tc":
                    throw new ArgumentException(
                        $"Cannot add '{type}' under {parentPath}: cells must be added under a row (/body/tbl[N]/tr[M]).");
                case "run":
                case "r":
                case "hyperlink":
                case "link":
                    // Inline-level elements can't be direct body children — they
                    // must live inside a paragraph. Reject CopyFrom that would
                    // produce <w:r>/<w:hyperlink> as a body child.
                    // (bookmark/field/pagebreak are wrapped or pair-inserted by
                    // their Add* helpers when targeting /body, so allowed.)
                    throw new ArgumentException(
                        $"Cannot add '{type}' under {parentPath}: inline-level elements must live inside a paragraph (/body/p[N]).");
                case "sectpr":
                    // Raw <w:sectPr> as a direct body child is a singleton managed
                    // implicitly by the document; block direct clone-via-from that
                    // would produce two <w:sectPr> children. Note: `--type section`
                    // is a distinct legit operation (creates a paragraph whose pPr
                    // carries a sectPr — a section break) and is allowed.
                    throw new ArgumentException(
                        $"Cannot add '{type}' under {parentPath}: body-level <w:sectPr> is a singleton. Use 'officecli set /body/sectPr' to modify it, or add a section break via `--type section` (which creates a paragraph-level break).");
                case "style":
                    throw new ArgumentException(
                        $"Cannot add 'style' under {parentPath}: styles belong under /styles, not /body.");
            }
        }

        // <w:tc> (TableCell) accepts only block-level elements: paragraph,
        // table, sdt, tcPr, customXml. Reject bare runs/hyperlinks/cells
        // cloned directly into a cell via --from, mirroring Table/TableRow.
        if (parent is TableCell)
        {
            switch (t)
            {
                case "paragraph":
                case "p":
                case "table":
                case "tbl":
                case "sdt":
                case "contentcontrol":
                    break;
                // Inline content with explicit cell-wrap helpers in
                // AddPicture/AddOle (Add.Media.cs) — they wrap the run in a
                // Paragraph inside the cell, satisfying the OOXML block-level
                // requirement transparently.
                case "picture":
                case "image":
                case "img":
                case "ole":
                case "oleobject":
                case "object":
                case "embed":
                    break;
                case "cell":
                case "tc":
                    throw new ArgumentException(
                        $"Cannot add '{type}' under {parentPath}: cells cannot be nested inside cells. Add cells under a row (/body/tbl[N]/tr[M]).");
                default:
                    throw new ArgumentException(
                        $"Cannot add '{type}' under {parentPath}: table cells only accept paragraphs, tables, or SDTs (block-level content). Add the element inside a paragraph first.");
            }
        }

        // Global: 'style' belongs only under /styles, never anywhere else.
        if (t == "style" && parent is not Styles)
        {
            throw new ArgumentException(
                $"Cannot add 'style' under {parentPath}: styles belong under /styles.");
        }

        // Global: 'num' / 'abstractNum' belong only under /numbering. Mirrors
        // the 'style'/'styles' pairing — definition parts have a single allowed
        // parent path so users don't have to guess where they go.
        if ((t == "num" || t == "abstractnum") && parent is not Numbering)
        {
            throw new ArgumentException(
                $"Cannot add '{type}' under {parentPath}: numbering definitions belong under /numbering.");
        }

        // /numbering only accepts numbering definitions (num, abstractNum). Reject everything else
        // so a stray --type p doesn't corrupt numbering.xml.
        if (parent is Numbering)
        {
            if (t != "num" && t != "abstractnum")
                throw new ArgumentException(
                    $"Cannot add '{type}' under /numbering. /numbering only holds numbering definitions — use --type num (with --prop abstractNumId=N) or --type abstractNum.");
        }

        // 'tab' (tab stop) lives in a paragraph's pPr/tabs container, or in a
        // paragraph/table style's pPr/tabs container. Reject anywhere else so
        // users get a useful pointer instead of falling through to AddDefault
        // and writing a stray <w:tab> at the wrong level.
        if (t == "tab" || t == "tabstop")
        {
            if (parent is Style stl)
            {
                var stType = stl.Type?.Value;
                if (stType != StyleValues.Paragraph && stType != StyleValues.Table)
                    throw new ArgumentException(
                        $"Cannot add 'tab' under {parentPath}: style '{stl.StyleId?.Value}' is type=" +
                        $"{stl.Type?.InnerText ?? "(unset)"}. Tab stops require a paragraph or table style.");
            }
            else if (parent is not Paragraph)
            {
                throw new ArgumentException(
                    $"Cannot add 'tab' under {parentPath}: tab stops belong inside a paragraph (e.g. /body/p[N]) " +
                    $"or a paragraph-typed style (e.g. /styles/Heading1).");
            }
        }


        // <w:tbl> only accepts tblPr, tblGrid, tr, sdt, customXml as children.
        // Reject anything else (paragraph, table, section, toc, break, ...) so
        // Word doesn't open a corrupted document silently.
        if (parent is Table)
        {
            switch (t)
            {
                case "row":
                case "tr":
                    break;
                default:
                    throw new ArgumentException(
                        $"Cannot add '{type}' under {parentPath}: tables only accept rows (/body/tbl[N]/tr). Use --type row.");
            }
        }

        // <w:tr> only accepts trPr, tc, sdt, customXml as children.
        if (parent is TableRow)
        {
            switch (t)
            {
                case "cell":
                case "tc":
                    break;
                default:
                    throw new ArgumentException(
                        $"Cannot add '{type}' under {parentPath}: table rows only accept cells (/body/tbl[N]/tr[M]/tc). Use --type cell.");
            }
        }

        // <w:sdt>/<w:sdtContent> wrappers don't accept arbitrary children as
        // direct kids. SdtBlock/SdtRun only hold sdtPr + sdtContent; any
        // block-level add under /body/sdt[N] belongs under
        // /body/sdt[N]/sdtContent. Reject the degenerate path with a
        // pointer to the content wrapper instead of silently producing
        // <w:p> as a direct child of <w:sdt> (schema-invalid).
        if (parent is SdtBlock || parent is SdtRun)
        {
            throw new ArgumentException(
                $"Cannot add '{type}' directly under {parentPath}. SDT (content control) elements only contain <w:sdtPr> and <w:sdtContent>. Add under {parentPath}/sdtContent instead.");
        }

        // /styles is the StyleDefinitions root. It only holds <w:style>,
        // <w:docDefaults>, and latentStyles. Every other type (paragraph,
        // table, toc, section, sdt, pagebreak, ...) would corrupt styles.xml.
        if (parent is Styles)
        {
            if (t != "style")
                throw new ArgumentException(
                    $"Cannot add '{type}' under /styles. /styles only holds style definitions — use --type style with --prop id=... --prop name=... (and basedOn/font/size/etc.) to add one.");
        }
    }

    public (string RelId, string PartPath) AddPart(string parentPartPath, string partType, Dictionary<string, string>? properties = null)
    {
        var mainPart = _doc.MainDocumentPart!;

        switch (partType.ToLowerInvariant())
        {
            case "chart":
                var chartPart = mainPart.AddNewPart<ChartPart>();
                var relId = mainPart.GetIdOfPart(chartPart);
                // Initialize with minimal valid ChartSpace
                chartPart.ChartSpace = new C.ChartSpace(
                    new C.Chart(new C.PlotArea(new C.Layout()))
                );
                chartPart.ChartSpace.Save();
                var chartIdx = mainPart.ChartParts.ToList().IndexOf(chartPart);
                return (relId, $"/chart[{chartIdx + 1}]");

            case "header":
                var headerPart = mainPart.AddNewPart<HeaderPart>();
                var hRelId = mainPart.GetIdOfPart(headerPart);
                headerPart.Header = new Header(new Paragraph());
                headerPart.Header.Save();
                var hIdx = mainPart.HeaderParts.ToList().IndexOf(headerPart);
                return (hRelId, $"/header[{hIdx + 1}]");

            case "footer":
                var footerPart = mainPart.AddNewPart<FooterPart>();
                var fRelId = mainPart.GetIdOfPart(footerPart);
                footerPart.Footer = new Footer(new Paragraph());
                footerPart.Footer.Save();
                var fIdx = mainPart.FooterParts.ToList().IndexOf(footerPart);
                return (fRelId, $"/footer[{fIdx + 1}]");

            default:
                throw new ArgumentException(
                    $"Unknown part type: {partType}. Supported: chart, header, footer");
        }
    }


    private void SetDocumentProperties(Dictionary<string, string> properties, List<string>? unsupported = null)
    {
        var doc = _doc.MainDocumentPart?.Document
            ?? throw new InvalidOperationException("Document not found");

        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "pagebackground" or "background":
                    doc.DocumentBackground = new DocumentBackground { Color = value };
                    // Enable background display in settings
                    var settingsPart = _doc.MainDocumentPart!.DocumentSettingsPart
                        ?? _doc.MainDocumentPart.AddNewPart<DocumentSettingsPart>();
                    settingsPart.Settings ??= new Settings();
                    if (settingsPart.Settings.GetFirstChild<DisplayBackgroundShape>() == null)
                        settingsPart.Settings.AddChild(new DisplayBackgroundShape());
                    settingsPart.Settings.Save();
                    break;

                case "defaultfont":
                    // Delegate to TrySetDocDefaults which uses EnsureRunPropsDefault()
                    // to create the DocDefaults chain when absent (e.g. blank documents).
                    TrySetDocDefaults("docdefaults.font", value);
                    break;
                case "defaultfontsize":
                    TrySetDocDefaults("docdefaults.fontsize", value);
                    break;

                case "pagewidth":
                    EnsureSectionProperties().GetFirstChild<PageSize>()!.Width = ParseTwips(value);
                    break;
                case "pageheight":
                    EnsureSectionProperties().GetFirstChild<PageSize>()!.Height = ParseTwips(value);
                    break;
                case "margintop":
                    EnsurePageMargin().Top = (int)ParseTwips(value);
                    break;
                case "marginbottom":
                    EnsurePageMargin().Bottom = (int)ParseTwips(value);
                    break;
                case "marginleft":
                    EnsurePageMargin().Left = ParseTwips(value);
                    break;
                case "marginright":
                    EnsurePageMargin().Right = ParseTwips(value);
                    break;

                // Core document properties
                case "title":
                    _doc.PackageProperties.Title = value;
                    break;
                case "author" or "creator":
                    _doc.PackageProperties.Creator = value;
                    break;
                case "subject":
                    _doc.PackageProperties.Subject = value;
                    break;
                case "keywords":
                    _doc.PackageProperties.Keywords = value;
                    break;
                case "description":
                    _doc.PackageProperties.Description = value;
                    break;
                case "category":
                    _doc.PackageProperties.Category = value;
                    break;
                case "lastmodifiedby":
                    _doc.PackageProperties.LastModifiedBy = value;
                    break;
                case "revision":
                    _doc.PackageProperties.Revision = value;
                    break;

                case "protection":
                {
                    var protSettingsPart = _doc.MainDocumentPart!.DocumentSettingsPart
                        ?? _doc.MainDocumentPart.AddNewPart<DocumentSettingsPart>();
                    protSettingsPart.Settings ??= new Settings();

                    // Remove existing protection
                    var existing = protSettingsPart.Settings.GetFirstChild<DocumentProtection>();
                    existing?.Remove();

                    if (!string.Equals(value, "none", StringComparison.OrdinalIgnoreCase))
                    {
                        var editValue = value.ToLowerInvariant() switch
                        {
                            "forms" => DocumentProtectionValues.Forms,
                            "readonly" => DocumentProtectionValues.ReadOnly,
                            "comments" => DocumentProtectionValues.Comments,
                            "trackedchanges" => DocumentProtectionValues.TrackedChanges,
                            _ => DocumentProtectionValues.Forms
                        };
                        var prot = new DocumentProtection
                        {
                            Edit = new EnumValue<DocumentProtectionValues>(editValue),
                            Enforcement = new OnOffValue(true)
                        };
                        protSettingsPart.Settings.AppendChild(prot);
                    }

                    protSettingsPart.Settings.Save();
                    break;
                }

                case "acceptallchanges" or "accept-changes" or "acceptchanges":
                    if (value.Equals("all", StringComparison.OrdinalIgnoreCase) || IsTruthy(value))
                        AcceptAllChanges();
                    break;
                case "rejectallchanges" or "reject-changes" or "rejectchanges":
                    if (value.Equals("all", StringComparison.OrdinalIgnoreCase) || IsTruthy(value))
                        RejectAllChanges();
                    break;

                default:
                    // Try document settings, section layout, compatibility, and docDefaults
                    var lowerKey = key.ToLowerInvariant();
                    if (!TrySetDocSetting(lowerKey, value)
                        && !TrySetSectionLayout(lowerKey, value)
                        && !TrySetCompatibility(lowerKey, value)
                        && !TrySetDocDefaults(lowerKey, value)
                        && !Core.ThemeHandler.TrySetTheme(_doc.MainDocumentPart?.ThemePart, lowerKey, value)
                        && !Core.ExtendedPropertiesHandler.TrySetExtendedProperty(
                            Core.ExtendedPropertiesHandler.GetOrCreateExtendedPart(_doc), lowerKey, value))
                        unsupported?.Add(key);
                    break;
            }
        }
    }

    private SectionProperties EnsureSectionProperties()
    {
        var body = _doc.MainDocumentPart!.Document!.Body!;
        var sectPr = body.GetFirstChild<SectionProperties>();
        if (sectPr == null)
        {
            sectPr = new SectionProperties();
            body.AppendChild(sectPr);
        }
        if (sectPr.GetFirstChild<PageSize>() == null)
        {
            var pgSz = new PageSize { Width = WordPageDefaults.A4WidthTwips, Height = WordPageDefaults.A4HeightTwips };
            // Schema order: pgSz must come before pgMar, cols, and docGrid
            var firstNonRef = sectPr.ChildElements.FirstOrDefault(c =>
                c is not HeaderReference && c is not FooterReference && c is not SectionType);
            if (firstNonRef != null)
                firstNonRef.InsertBeforeSelf(pgSz);
            else
                sectPr.AppendChild(pgSz);
        }
        return sectPr;
    }

    private PageMargin EnsurePageMargin()
    {
        var sectPr = EnsureSectionProperties();
        var margin = sectPr.GetFirstChild<PageMargin>();
        if (margin == null)
        {
            margin = new PageMargin { Top = 1440, Bottom = 1440, Left = 1800, Right = 1800 };
            // Insert after PageSize to maintain CT_SectPr schema order: pgSz → pgMar → ...
            var pgSz = sectPr.GetFirstChild<PageSize>();
            if (pgSz != null)
                pgSz.InsertAfterSelf(margin);
            else
                sectPr.AddChild(margin, throwOnError: false);
        }
        return margin;
    }
}
