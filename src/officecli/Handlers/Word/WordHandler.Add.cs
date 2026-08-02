// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
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
        Modified = true;
        // The signature is non-nullable, but the body uses `type?.Equals(...)`
        // below to short-circuit header/footer routing — that null-conditional
        // makes the C# flow analyzer treat `type` as nullable from that point
        // on, surfacing CS8604 at the ValidateParentChild call. Validate up
        // front so the analyzer (and any caller violating the signature) gets
        // a clean failure instead of a NRE down the line.
        ArgumentNullException.ThrowIfNull(type);

        // CONSISTENCY(prop-key-case): property keys are case-insensitive
        // ("SRC"/"src"/"Src" all resolve the same). Normalize once at the
        // dispatch entry so every AddXxx helper can rely on TryGetValue("src").
        properties = properties switch
        {
            null => new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase),
            // Preserve TrackingPropertyDictionary so handler-as-truth read
            // tracking survives this entry normalization.
            OfficeCli.Core.TrackingPropertyDictionary => properties,
            var p when p.Comparer == StringComparer.OrdinalIgnoreCase => p,
            _ => new Dictionary<string, string>(properties, StringComparer.OrdinalIgnoreCase),
        };

        // Reset per-Add diagnostic. Helpers that detect silent-drop props
        // (currently only AddStyle) populate this; the CLI layer surfaces
        // it as a WARNING line so curated-surface gaps stop being silent.
        LastAddUnsupportedProps = new List<string>();
        LastAddWarnings = new List<string>();
        LastUnrecognizedLatex = new List<string>();

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
        else if (type.Equals("header", StringComparison.OrdinalIgnoreCase)
                 || type.Equals("footer", StringComparison.OrdinalIgnoreCase))
        {
            // /section[N] for header/footer add: NavigateToElement only
            // resolves break-paragraph carriers (n <= sectParas.Count); the
            // final body-level sectPr (n == sectParas.Count + 1) has no
            // carrier paragraph. AddHeader/AddFooter map parentPath →
            // sectPr via ResolveTargetSectPrForHeaderFooter (string-based,
            // independent of `parent`), so route through with parent=body.
            var sectMatch = System.Text.RegularExpressions.Regex.Match(
                parentPath, @"^/section\[(\d+)\]/?$",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            if (sectMatch.Success)
            {
                parent = body;
            }
            else
            {
                List<PathSegment> parts;
                try { parts = ParsePath(parentPath); }
                catch (Exception ex) when (ex is not ArgumentException and not InvalidOperationException)
                {
                    throw new ArgumentException($"Malformed parent path '{parentPath}'. Check selector brackets and escape sequences.", ex);
                }
                parent = NavigateToElement(parts, out var ctx)
                    ?? throw new ArgumentException($"Path not found: {parentPath}" + (ctx != null ? $". {ctx}" : ""));
            }
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

        // Only a body-level add changes the body's direct paragraph/table set
        // that GetBodyChildIndex caches. Run-adds (parent = a paragraph) and
        // table-cell adds (parent = a cell paragraph) don't, so they must NOT
        // clear it — clearing on every add turned the table-heavy gov_ndrc
        // replay O(n²) (each later /body/tbl[last()] navigation rebuilt the
        // whole index).
        if (parent is Body) ClearBodyChildIndex();

        // PERF(nav-child-cache): a run/cell add under a paragraph (parent != Body)
        // does NOT change the body-direct index above, but it DOES change the
        // paragraph's run set — so a cached /<para>/r[K] (or /<row>/tc[K]) list
        // must be dropped or a later resolve returns stale runs. This must fire
        // on EXIT, not here at entry: an --after/--before add navigates to its
        // anchor (e.g. /body/tbl[1]/tr[2]) DURING the add, which repopulates the
        // very row/cell cache from the PRE-insert tree; an entry-only clear would
        // then be overwritten and leave the positional insert resolving against a
        // stale index (symptom: `add row --before tr[2]` lands at the wrong slot).
        // Always-armed (cheap Dictionary.Clear; build-time nav caches are empty,
        // so the append fast-path is untouched). Mirrors the _anchorCacheGuard
        // exit-timing rationale below.
        using var _navChildCacheGuard = new NavCacheClearGuard(this);

        // --after/--before poisons the cache mid-Add: ResolveAnchorPosition
        // navigates to the anchor, which REBUILDS the child-index cache from
        // the pre-mutation tree; the positional insert then leaves it stale
        // for the rest of the session (symptom: "No tbl found at /body" while
        // the same error lists tbl(1) as available — navigation read the
        // poisoned cache, the error message enumerated the live DOM).
        // AddParagraph invalidates after its own positional insert, but other
        // types (table, ...) did not — so arm an exit-invalidate guard, the
        // same pattern Remove/Move/Swap/CopyFrom use (BodyCacheGuard, "must
        // invalidate on exit, after the structural change has happened").
        // Conditional on an anchor position: the append hot path (batch replay
        // of thousands of paragraphs) must keep its caches or it turns O(n²).
        using var _anchorCacheGuard = parent is Body && (position?.After != null || position?.Before != null)
            ? new BodyCacheGuard(this)
            : default;

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
            // `diagram` is overloaded: the mermaid synthesizer (mermaid/text/dsl/src)
            // and the dump→batch verbatim carrier that rebuilds a native OOXML
            // SmartArt diagram from raw parts (carries `runXml`). Route to mermaid
            // only when it is NOT the parts carrier; `flowchart` is always mermaid.
            "flowchart" => AddDiagram(parent, parentPath, index, properties),
            "diagram" when !properties.ContainsKey("runXml")
                => AddDiagram(parent, parentPath, index, properties),
            "markdown" or "md" => AddMarkdown(parent, parentPath, index, properties),
            "run" or "r" => AddRun(parent, parentPath, index, properties),
            "table" or "tbl" => AddTable(parent, parentPath, index, properties),
            "row" or "tr" => AddRow(parent, parentPath, index, properties),
            "col" or "column" => AddTableColumn(parent, parentPath, index, properties),
            "cell" or "tc" => AddCell(parent, parentPath, index, properties),
            "tab" or "tabstop" => AddTab(parent, parentPath, index, properties),
            "ptab" or "positionaltab" => AddPtab(parent, parentPath, index, properties),
            "chart" => AddChart(parent, parentPath, index, properties),
            "picture" or "image" or "img" => AddPicture(parent, parentPath, index, properties),
            "ole" or "oleobject" or "object" or "embed" => AddOle(parent, parentPath, index, properties),
            // Unified verbatim part-owning carrier (dump→batch only). The former
            // per-element verbs (chartpart/diagram/smartart/vmlshape/drawingshape/
            // activex) differed only in a marker check and all delegated to the
            // same routine; they remain accepted as input aliases for hand-written
            // batches, but the emitter now emits the single canonical `inlinedparts`.
            "inlinedparts" or "chartpart" or "activex" or "diagram" or "smartart"
                or "vmlshape" or "drawingshape"
                => AddInlinedPartsRun(parent, parentPath, properties, "inlinedparts"),
            "comment" => AddComment(parent, parentPath, index, properties),
            "bookmark" => AddBookmark(parent, parentPath, index, properties),
            "bookmarkend" => AddBookmarkEnd(parent, parentPath, index, properties),
            "permstart" or "permend" => AddPerm(parent, parentPath, index, properties, type),
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
            "textbox" or "txbx" => AddTextbox(parent, parentPath, index, properties),
            "shape" or "sp" => AddShape(parent, parentPath, index, properties),
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
            // Reject standalone comment range markers. Falling through to
            // AddDefault triggers schema-aware insertion via Particle.Set
            // which CLEARS existing run children of the paragraph (data
            // loss). The atomic, safe path is `add --type comment` which
            // creates both range markers + comment text together.
            "commentrangestart" or "commentrangeend" or "commentreference" =>
                throw new ArgumentException(
                    $"Cannot add '{type}' directly. Adding a bare comment range marker into a paragraph destroys existing runs (schema-aware sequence reset). Use `add --type comment --prop start=... --prop end=... --prop text=...` to create the comment atomically."),
            // Reject altChunk: it embeds alternate-format payloads (HTML/RTF
            // fragments) via OOXML relationship-bound parts. AddDefault would
            // fall through to TryCreateTypedElement which writes user props
            // as raw unnamespaced attrs (e.g. src=...) — schema-invalid; Word
            // rejects the file. Batch dump already warns+drops altChunk for
            // the same reason.
            "altchunk" =>
                throw new ArgumentException(
                    "Cannot add 'altChunk' directly. altChunk embeds alternate-format payloads via OOXML relationships which require a curated implementation. Use the batch import or raw-set path for round-trip fidelity."),
            _ => AddDefault(parent, parentPath, index, properties, type),
        };
        }
        catch (ArgumentOutOfRangeException ex)
        {
            // Surface as a clean ArgumentException (CLI layer formats Message).
            // Scrub the raw .NET parameter noise.
            throw new ArgumentException($"Invalid index or anchor for add '{type}'. Check --index / --after / --before values.", ex);
        }

        SaveDoc();
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
        // of which are managed via `set` on /body/sectPr instead. EXCEPTION:
        // header/footer adds are routed by section selector; the actual part
        // attachment runs via ResolveTargetSectPrForHeaderFooter.
        if (parent is SectionProperties && t != "header" && t != "footer")
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
                // The inlined-parts carrier wraps the run in a cell paragraph, same
                // as AddOle — block-level schema requirement satisfied. Old verb
                // aliases kept alongside the unified `inlinedparts`.
                case "inlinedparts":
                case "activex":
                case "diagram":
                case "smartart":
                case "vmlshape":
                case "drawingshape":
                    break;
                // BUG-FIX(B2): bookmark is an inline-level construct, but
                // AddBookmark redirects into the cell's first paragraph
                // (auto-creating one if needed) so the resulting XML stays
                // schema-valid (cell only accepts block-level children).
                case "bookmark":
                case "bookmarkend":
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

        // /numbering only accepts numbering definitions. Reject stray curated
        // types (a typo'd --type p) so they can't corrupt numbering.xml. A
        // namespace-prefixed type (e.g. w:abstractNum, w:num, w:numPicBullet) is
        // an explicit generic-add request — the dump→batch recursive emitter
        // rebuilds the whole subtree this way, bypassing the curated
        // abstractNum/num seeding (which auto-fills 9 default levels). Let those
        // through to AddDefault, mirroring how /styles children (w:pPr, w:rPr,
        // w:tblStylePr, …) reach the generic path. CONSISTENCY(numbering-typed-decomp).
        if (parent is Numbering)
        {
            bool prefixedGeneric = t.Contains(':');
            if (t != "num" && t != "abstractnum" && !prefixedGeneric)
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
                case "col":
                case "column":
                    // 'col'/'column' is a virtual element synthesized by
                    // AddTableColumn (gridCol + per-row tc). OOXML has no
                    // <w:col> child; the gate is opened here so dispatch
                    // reaches the column helper.
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

        // CONSISTENCY(set-atomicity): multi-prop set must be all-or-nothing. The
        // resident process keeps the doc in memory, so a throw partway through this
        // foreach would otherwise leave earlier props applied while the command exits
        // non-zero — visible to the next read. Snapshot Document OuterXml on entry;
        // any exception restores the whole document tree before re-throwing. The body
        // ref captured outside is invalid after restore — callers of doc.Body must
        // re-resolve via _doc.MainDocumentPart.Document.Body if they cache it.
        var atomicSnapshot = doc.OuterXml;
        try
        {
        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "pagebackground" or "background":
                    // w:background/@w:color is ST_HexColor (bare RRGGBB or
                    // "auto") — strip a leading '#' and resolve named/rgb()
                    // forms like every other color write. Get emits the
                    // canonical "#FFFFFF" form (FormatHexColor); without this
                    // the '#' leaked verbatim into the attribute and a
                    // dump→batch round-trip produced schema-invalid OOXML.
                    doc.DocumentBackground = new DocumentBackground
                    {
                        Color = OfficeCli.Core.ParseHelpers.SanitizeColorForOoxml(value).Rgb
                    };
                    // Enable background display in settings
                    var settingsPart = _doc.MainDocumentPart!.DocumentSettingsPart
                        ?? _doc.MainDocumentPart.AddNewPart<DocumentSettingsPart>();
                    settingsPart.Settings ??= new Settings();
                    if (settingsPart.Settings.GetFirstChild<DisplayBackgroundShape>() == null)
                        settingsPart.Settings.AddChild(new DisplayBackgroundShape());
                    settingsPart.Settings.Save();
                    break;

                case "recalcfields":
                    // Compute + write cached values we can do WITHOUT a layout
                    // engine: today that's SEQ numbering (document-order count).
                    // PAGE/PAGEREF/TOC page numbers need pagination — pair with
                    // `--prop updateFields=true` to defer those to Word.
                    if (value.Trim().ToLowerInvariant() is "seq" or "all" or "true" or "")
                        RecalcSeqFields();
                    else
                        (unsupported ??= new()).Add($"recalcFields={value} (supported: seq)");
                    break;

                case "defaultfont":
                    // Delegate to TrySetDocDefaults which uses EnsureRunPropsDefault()
                    // to create the DocDefaults chain when absent (e.g. blank documents).
                    TrySetDocDefaults("docdefaults.font", value);
                    break;
                case "defaultfontsize":
                    TrySetDocDefaults("docdefaults.fontsize", value);
                    break;

                // Dump→batch fidelity: a source body sectPr that OMITS <w:pgSz>
                // (deferring to Word's application default — US Letter) must NOT
                // inherit the blank template's stamped A4 pgSz on rebuild. The
                // emitter signals the absence with pageSize="none"; remove the
                // element so the rebuilt sectPr also defers to the app default.
                // ("none" is the established remove sentinel on sectPr children —
                // see pageStart/lineNumbers/valign/pgBorders in
                // WordHandler.Set.SectionLayout.cs.) Independent of pageMargin so
                // a source with one but not the other round-trips correctly.
                case "pagesize":
                    if (string.Equals(value, "none", StringComparison.OrdinalIgnoreCase))
                        BodySectionPropertiesForRemove()?.RemoveAllChildren<PageSize>();
                    break;
                case "pagemargin":
                    if (string.Equals(value, "none", StringComparison.OrdinalIgnoreCase))
                        BodySectionPropertiesForRemove()?.RemoveAllChildren<PageMargin>();
                    break;

                // BUG-DUMP-R31-1: emitter signal that the source body carried a
                // <w:sectPr> element (possibly childless). Materialize it so a
                // bare <w:sectPr/> round-trips — a missing sectPr renders at a
                // different page width than an empty one. EnsureSectionProperties
                // creates the body sectPr if absent; no child is added, so an
                // otherwise-empty source stays empty. Combined with pageSize=none /
                // pageMargin=none, this also suppresses the drop-the-empty-sectPr
                // path below (RemovedBothPageGeometry sees sectPr=present).
                case "sectpr":
                    if (string.Equals(value, "present", StringComparison.OrdinalIgnoreCase))
                        EnsureBareSectionProperties();
                    break;

                case "pagewidth" or "width":
                {
                    var twW = ParseTwips(value);
                    Core.WordPageDefaults.ValidatePageDim(twW, "pageWidth");
                    EnsureSectionProperties().GetFirstChild<PageSize>()!.Width = twW;
                    break;
                }
                case "pageheight" or "height":
                {
                    var twH = ParseTwips(value);
                    Core.WordPageDefaults.ValidatePageDim(twH, "pageHeight");
                    EnsureSectionProperties().GetFirstChild<PageSize>()!.Height = twH;
                    break;
                }
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
                case "marginheader":
                    EnsurePageMargin().Header = ParseTwips(value);
                    break;
                case "marginfooter":
                    EnsurePageMargin().Footer = ParseTwips(value);
                    break;
                case "margingutter":
                    EnsurePageMargin().Gutter = ParseTwips(value);
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
                case "revisionnumber":
                    _doc.PackageProperties.Revision = value;
                    break;

                case "protection":
                {
                    var protSettingsPart = _doc.MainDocumentPart!.DocumentSettingsPart
                        ?? _doc.MainDocumentPart.AddNewPart<DocumentSettingsPart>();
                    protSettingsPart.Settings ??= new Settings();

                    var existing = protSettingsPart.Settings.GetFirstChild<DocumentProtection>();

                    if (string.Equals(value, "none", StringComparison.OrdinalIgnoreCase))
                    {
                        // Explicit "none" still removes the element.
                        existing?.Remove();
                    }
                    else
                    {
                        var editValue = value.ToLowerInvariant() switch
                        {
                            "forms" => DocumentProtectionValues.Forms,
                            "readonly" => DocumentProtectionValues.ReadOnly,
                            "comments" => DocumentProtectionValues.Comments,
                            "trackedchanges" => DocumentProtectionValues.TrackedChanges,
                            _ => DocumentProtectionValues.Forms
                        };
                        // BUG-DUMP-PROTECTION-ENFORCE: honor an accompanying
                        // protectionEnforced flag so a protection mode that is
                        // DEFINED-but-NOT-ENFORCED in the source (w:enforcement="0")
                        // round-trips as unenforced. Forcing enforcement on flips
                        // Word into form-fill mode and shifts every line ~12px.
                        // Default to enforced when the flag is absent so the
                        // single-command `set / --prop protection=forms` still means
                        // "enforce".
                        bool enforce = !properties.TryGetValue("protectionEnforced", out var enfVal)
                            || enfVal == null || IsTruthy(enfVal);
                        if (existing != null)
                        {
                            // Update Edit + Enforcement in place; preserve any
                            // crypto attributes (cryptSpinCount/hash/salt/...)
                            // that were injected via raw-set. A replace-new
                            // path would silently destroy the password payload.
                            existing.Edit = new EnumValue<DocumentProtectionValues>(editValue);
                            existing.Enforcement = new OnOffValue(enforce);
                        }
                        else
                        {
                            var prot = new DocumentProtection
                            {
                                Edit = new EnumValue<DocumentProtectionValues>(editValue),
                                Enforcement = new OnOffValue(enforce)
                            };
                            // CONSISTENCY(settings-schema-order): w:documentProtection
                            // must precede w:compat / w:charSpacingControl in w:settings
                            // (CT_Settings sequence). Plain AppendChild lands after
                            // any pre-existing compat block and fails OOXML validation
                            // (R12 minor). Reuse the existing helper.
                            InsertBeforeCompatibility(protSettingsPart.Settings, prot);
                        }
                    }

                    protSettingsPart.Settings.Save();
                    break;
                }

                case "protectionenforced":
                {
                    // BUG-DUMP-PROTECTION-ENFORCE: enforcement state for
                    // documentProtection. Apply to the existing protection element
                    // if one is present; a no-op when there is none (the
                    // `protection` case — which may run before or after this one in
                    // the same multi-prop op, dict order being unspecified — is
                    // authoritative and reads this flag directly). Having a real
                    // case also stops the generic-setting fallthrough from emitting
                    // a spurious warning on a round-tripped protectionEnforced prop.
                    var enfPart = _doc.MainDocumentPart?.DocumentSettingsPart;
                    var existingProt = enfPart?.Settings?.GetFirstChild<DocumentProtection>();
                    if (existingProt != null)
                    {
                        existingProt.Enforcement = new OnOffValue(IsTruthy(value));
                        enfPart!.Settings!.Save();
                    }
                    break;
                }

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

        // Geometry-less round-trip: distinguish "source had NO body sectPr at
        // all" from "source had a pgSz-less sectPr that carries OTHER real
        // content" (docGrid linesAndChars, cols, headers, …). The emitter
        // signals a pgSz-less source with pageSize=none / pageMargin=none, but
        // a truly sectPr-less source emits NOTHING else in this `set /` step —
        // its props are exactly {pageSize:none, pageMargin:none}. After the
        // remove sentinels strip pgSz/pgMar, the rebuild target is left with
        // only the blank template's stamped <w:docGrid w:type="default"/> — a
        // present-but-empty sectPr that Word resolves to the schema default
        // (US Letter), flipping the page size away from the app default the
        // sectPr-less source rendered at. So: when BOTH remove signals fired
        // and the resulting body sectPr has no real content (only a
        // default/empty docGrid, nothing else), drop the sectPr entirely so the
        // rebuild matches the source's sectPr-less state. A sectPr that still
        // holds real content (a non-default docGrid, cols, headers, pgSz, …) is
        // never removed — those keys were applied earlier in this same loop.
        if (RemovedBothPageGeometry(properties))
        {
            var bodySectPr = _doc.MainDocumentPart?.Document?.Body?.GetFirstChild<SectionProperties>();
            if (bodySectPr != null && IsEmptyDefaultSectionProperties(bodySectPr))
                bodySectPr.Remove();
        }
        // BUG-DUMP-R31-1: the source carried a childless <w:sectPr/> (sectPr=
        // present + both geometry-none sentinels). The drop path above is
        // suppressed so the element survives, but the blank rebuild target still
        // holds the template's default <w:docGrid w:type="default"/>. Strip that
        // trivial default so the rebuilt sectPr is bare like the source — a bare
        // <w:sectPr/> and a sectPr carrying a stray default docGrid resolve to
        // the same page width, but a byte-faithful empty round-trip is cleaner
        // and matches the source's IsEmptyDefaultSectionProperties shape.
        else if (KeptEmptySectPr(properties))
        {
            var bodySectPr = _doc.MainDocumentPart?.Document?.Body?.GetFirstChild<SectionProperties>();
            if (bodySectPr != null && IsEmptyDefaultSectionProperties(bodySectPr))
                bodySectPr.RemoveAllChildren<DocGrid>();
        }
        }
        catch
        {
            // Restore the in-memory Document tree from the pre-mutation snapshot so the
            // failed command leaves no partial state. Re-throw so the CLI surface still
            // reports the original error and exits non-zero. Document(string) accepts
            // OuterXml form per the OpenXmlElement(outerXml) constructor contract.
            _doc.MainDocumentPart!.Document = new Document(atomicSnapshot);
            throw;
        }
    }

    /// <summary>
    /// True when this `set /` carried BOTH the pageSize=none AND pageMargin=none
    /// remove sentinels — the emitter's signal that the source body sectPr
    /// omitted both &lt;w:pgSz&gt; and &lt;w:pgMar&gt;. Only this combination can
    /// indicate a truly sectPr-less source; one without the other is a partial
    /// geometry omission that keeps its sectPr.
    /// </summary>
    private static bool RemovedBothPageGeometry(Dictionary<string, string> properties)
    {
        bool pgSzNone = false, pgMarNone = false, sectPrPresent = false;
        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "pagesize" when string.Equals(value, "none", StringComparison.OrdinalIgnoreCase):
                    pgSzNone = true;
                    break;
                case "pagemargin" when string.Equals(value, "none", StringComparison.OrdinalIgnoreCase):
                    pgMarNone = true;
                    break;
                // BUG-DUMP-R31-1: the source body genuinely carried a <w:sectPr>
                // (even childless). A bare sectPr must round-trip — its absence
                // changes the rendered page width — so the drop-the-empty-sectPr
                // path must NOT fire even though both geometry-none sentinels did.
                case "sectpr" when string.Equals(value, "present", StringComparison.OrdinalIgnoreCase):
                    sectPrPresent = true;
                    break;
            }
        }
        return pgSzNone && pgMarNone && !sectPrPresent;
    }

    /// <summary>
    /// BUG-DUMP-R31-1: True when this `set /` carried BOTH geometry-none
    /// sentinels AND the sectPr=present marker — the source had a childless
    /// &lt;w:sectPr/&gt;. The element must be KEPT (not dropped) but trimmed of
    /// the blank template's default docGrid so it round-trips bare.
    /// </summary>
    private static bool KeptEmptySectPr(Dictionary<string, string> properties)
    {
        bool pgSzNone = false, pgMarNone = false, sectPrPresent = false;
        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "pagesize" when string.Equals(value, "none", StringComparison.OrdinalIgnoreCase):
                    pgSzNone = true;
                    break;
                case "pagemargin" when string.Equals(value, "none", StringComparison.OrdinalIgnoreCase):
                    pgMarNone = true;
                    break;
                case "sectpr" when string.Equals(value, "present", StringComparison.OrdinalIgnoreCase):
                    sectPrPresent = true;
                    break;
            }
        }
        return pgSzNone && pgMarNone && sectPrPresent;
    }

    /// <summary>
    /// True when a body-level sectPr carries no real content: it is either empty
    /// or holds nothing but a single &lt;w:docGrid&gt; whose only meaningful state
    /// is the blank template's default (type absent or "default", no linePitch /
    /// charSpace). Such a sectPr is what the blank rebuild is left with after a
    /// truly sectPr-less source's pgSz/pgMar removals; Word resolves it to the
    /// schema-default page size (US Letter), so it must be dropped to match the
    /// source's app-default fallback. Any other child (pgSz, pgMar, cols,
    /// headers, a non-default docGrid, …) means the sectPr is real and is kept.
    /// </summary>
    private static bool IsEmptyDefaultSectionProperties(SectionProperties sectPr)
    {
        foreach (var child in sectPr.ChildElements)
        {
            if (child is DocGrid dg)
            {
                // A docGrid is "trivial" only if it carries the blank default and
                // no line/char grid state. linePitch/charSpace presence (or a
                // non-default type) means a real CJK grid the source intended.
                bool typeIsDefault = dg.Type?.Value == null || dg.Type.Value == DocGridValues.Default;
                if (!typeIsDefault || dg.LinePitch?.Value != null || dg.CharacterSpace?.Value != null)
                    return false;
                continue;
            }
            // Any non-docGrid child is real section content.
            return false;
        }
        return true;
    }

    /// <summary>
    /// Reports whether the body-level sectPr explicitly carries a
    /// &lt;w:pgSz&gt; / &lt;w:pgMar&gt;. Used by the batch emitter to decide
    /// whether to emit a `pageSize=none` / `pageMargin=none` remove signal:
    /// a source that OMITS these defers to Word's application default
    /// (US Letter), and the rebuilt blank's stamped A4 must be stripped so
    /// both render identically. Reports independently per element.
    /// </summary>
    internal (bool hasPageSize, bool hasPageMargin) BodySectionPageGeometryPresence()
    {
        var sectPr = _doc.MainDocumentPart?.Document?.Body?.GetFirstChild<SectionProperties>();
        if (sectPr == null) return (false, false);
        return (sectPr.GetFirstChild<PageSize>() != null,
                sectPr.GetFirstChild<PageMargin>() != null);
    }

    /// <summary>
    /// BUG-DUMP-R31-1: True when the source body carries a &lt;w:sectPr&gt;
    /// element at all (even a childless one). The pageSize=none / pageMargin=none
    /// sentinels alone can't distinguish "source had NO sectPr" from "source had
    /// an EMPTY &lt;w:sectPr/&gt;" — and the two render at DIFFERENT page widths
    /// in real Word (a bare sectPr resolves to one default geometry, a missing
    /// sectPr to another). The emitter uses this to decide whether the drop-the-
    /// empty-sectPr path may fire (only when the source truly had none).
    /// </summary>
    internal bool BodyHasSectionProperties()
        => _doc.MainDocumentPart?.Document?.Body?.GetFirstChild<SectionProperties>() != null;

    /// <summary>
    /// Raw body-sectPr page geometry in native OOXML twips, keyed by the same
    /// batch prop names EmitSection uses (pageWidth/pageHeight, marginTop/…).
    /// Used by the batch emitter to write bare-twip values instead of the
    /// cm-rounded strings Get emits for human readback: twip→cm→twip is lossy
    /// (e.g. 1418 twips → "2.5cm" → 1417), so dump→batch round-trips drifted by
    /// ±1 twip. Bare integers parse back as exact twips (ParseTwips fallthrough),
    /// so this keeps the round-trip byte-exact while leaving the canonical cm
    /// Get output untouched. Only present keys are returned; absent pgSz/pgMar
    /// (and their per-attribute absences) leave the slot out so the
    /// pageSize=none / pageMargin=none sentinel path is unaffected.
    /// </summary>
    internal Dictionary<string, string> BodySectionPageGeometryTwips()
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var sectPr = _doc.MainDocumentPart?.Document?.Body?.GetFirstChild<SectionProperties>();
        if (sectPr == null) return result;
        var pgSz = sectPr.GetFirstChild<PageSize>();
        if (pgSz?.Width?.Value != null) result["pageWidth"] = pgSz.Width.Value.ToString();
        if (pgSz?.Height?.Value != null) result["pageHeight"] = pgSz.Height.Value.ToString();
        var pgMar = sectPr.GetFirstChild<PageMargin>();
        if (pgMar != null)
        {
            if (pgMar.Top?.Value != null) result["marginTop"] = ((uint)Math.Abs(pgMar.Top.Value)).ToString();
            if (pgMar.Bottom?.Value != null) result["marginBottom"] = ((uint)Math.Abs(pgMar.Bottom.Value)).ToString();
            if (pgMar.Left?.Value != null) result["marginLeft"] = pgMar.Left.Value.ToString();
            if (pgMar.Right?.Value != null) result["marginRight"] = pgMar.Right.Value.ToString();
            if (pgMar.Header?.Value != null) result["marginHeader"] = pgMar.Header.Value.ToString();
            if (pgMar.Footer?.Value != null) result["marginFooter"] = pgMar.Footer.Value.ToString();
            if (pgMar.Gutter?.Value != null) result["marginGutter"] = pgMar.Gutter.Value.ToString();
        }
        // BUG-DUMP-R25-4: cols @w:space on the body-level sectPr also drifts
        // through cm (708→"1.25cm"→709). Override the canonical columnSpace key
        // with raw twips, mirroring the pgSz/pgMar handling above. The inline-
        // section carrier path is fixed independently in Navigation.cs.
        var cols = sectPr.GetFirstChild<Columns>();
        if (cols?.Space?.Value != null && uint.TryParse(cols.Space.Value, out var colSpaceTwips))
            result["columnSpace"] = colSpaceTwips.ToString();
        return result;
    }

    /// <summary>
    /// Returns the existing body-level sectPr WITHOUT auto-stamping a
    /// PageSize/PageMargin. Used only by the pageSize/pageMargin remove
    /// sentinels: EnsureSectionProperties() re-stamps a default PageSize as a
    /// side effect, which would resurrect the very element a sibling
    /// `pageSize=none` had just removed (order-dependent corruption when both
    /// removes are in one `set /` call). Null when no sectPr exists yet.
    /// </summary>
    private SectionProperties? BodySectionPropertiesForRemove()
        => _doc.MainDocumentPart?.Document?.Body?.GetFirstChild<SectionProperties>();

    /// <summary>
    /// BUG-DUMP-R31-1: ensure the body has a &lt;w:sectPr&gt; element WITHOUT
    /// stamping the default A4 pgSz the way EnsureSectionProperties does. Used by
    /// the sectPr=present round-trip signal: an empty source sectPr must stay
    /// empty (pgSz/pgMar were stripped by the pageSize=none/pageMargin=none
    /// sentinels), and re-stamping A4 here would re-introduce the geometry the
    /// sentinels just removed.
    /// </summary>
    private SectionProperties EnsureBareSectionProperties()
    {
        var body = _doc.MainDocumentPart!.Document!.Body!;
        var sectPr = body.GetFirstChild<SectionProperties>();
        if (sectPr == null)
        {
            sectPr = new SectionProperties();
            body.AppendChild(sectPr);
        }
        return sectPr;
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
