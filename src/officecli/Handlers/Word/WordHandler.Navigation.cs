// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    // Perf: append-monotonic body caches for batch replay. Each body append
    // otherwise costs O(n) — navigation scans all paragraphs to resolve
    // /body/p[last()], and AddParagraph re-Count()s Elements<Paragraph>() to
    // name the result path — making an N-paragraph replay O(N²). Self-validating
    // against the live tree (see AppendedBodyParaCount), so no mutation site has
    // to remember to invalidate.
    private OpenXmlElement? _lastBodyParagraph;
    private int _bodyParaCount = -1;

    // The Table element created by the most recent AddTable call. Lets the
    // markdown expander populate cells on the element directly instead of
    // re-navigating "/body/tbl[N]" (which rebuilds the body child-index, O(body),
    // per table — O(M²) over an M-table interleave). Written only by AddTable and
    // consumed immediately after; not a cache (no invalidation semantics).
    private Table? _lastAddedTable;

    private void InvalidateBodyParaCache() { _lastBodyParagraph = null; _bodyParaCount = -1; ClearBodyChildIndex(); }

    // using-scope that invalidates the append/child caches when the enclosing
    // mutation method RETURNS. Remove/Move/Swap/CopyFrom navigate first (which
    // rebuilds the child-index cache to the PRE-mutation tree) and only then
    // mutate, so invalidating at entry would leave a stale cache behind. We must
    // invalidate on exit, after the structural change has happened.
    private readonly struct BodyCacheGuard : System.IDisposable
    {
        private readonly WordHandler _h;
        public BodyCacheGuard(WordHandler h) => _h = h;
        // Null-safe so `default(BodyCacheGuard)` is a no-op: Add() arms the
        // guard conditionally (only for --after/--before body-level adds).
        public void Dispose() => _h?.InvalidateBodyParaCache();
    }

    // Clears ONLY the run/row/cell nav child-index caches on exit. Set() arms
    // this when a mutation might rewrite a container's child set — most notably
    // `set text`, which replaces ALL runs of a paragraph/cell (3 runs → 1),
    // and revision accept/reject, which can delete inserted runs. Exit-timing
    // (not entry) so a set that navigates through the mutated segment mid-op —
    // repopulating the very cache it invalidates — still ends clean. Property-
    // only sets whose keys are all in NavCacheSafeSetKeys leave the guard
    // disarmed, so a per-run attribute-set batch keeps hitting the cache
    // (arming it there would rebuild O(R) per set → O(R²)). Null-safe so
    // `default(NavCacheClearGuard)` is a no-op.
    private readonly struct NavCacheClearGuard : System.IDisposable
    {
        private readonly WordHandler _h;
        public NavCacheClearGuard(WordHandler h) => _h = h;
        public void Dispose() => _h?.ClearNavChildCaches();
    }

    // ==================== Navigation ====================

    /// <summary>
    /// OOXML toggle element (Bold, Italic, Strike, Caps, …) is "ON" when the
    /// element exists AND its <c>w:val</c> attribute is either absent or
    /// truthy. <c>&lt;w:b/&gt;</c> means ON; <c>&lt;w:b w:val="0"/&gt;</c>
    /// and <c>&lt;w:b w:val="false"/&gt;</c> mean explicitly OFF. Pure
    /// null-checks on the element flip the OFF case back to ON, corrupting
    /// canonical Get readback (BUG-R2-04). Use this helper at every
    /// toggle-readback site so the override is honored.
    /// </summary>
    // BUG-R4B(BUG1): width-bearing OOXML attributes (w:tblInd/@w,
    // w:tblCellSpacing/@w, w:tcMar margins, …) are typed Int32Value in the SDK
    // and throw FormatException on the first .Value read when a producer emits
    // a decimal string ("0.0", "9440.0"). Read the raw InnerText and parse
    // leniently (truncate) so get/dump survive such files. Returns null when
    // absent or unparsable.
    private static int? SafeWidth(Int32Value? w) =>
        w is null ? null : ParseHelpers.LenientInt(w.InnerText);

    private static int? SafeWidth(StringValue? w) =>
        w is null ? null : ParseHelpers.LenientInt(w.InnerText);

    private static int? SafeWidth(Int16Value? w) =>
        w is null ? null : ParseHelpers.LenientInt(w.InnerText);

    private static bool IsToggleOn(Bold? t)   => t != null && (t.Val == null || t.Val.Value);
    private static bool IsToggleOn(Italic? t) => t != null && (t.Val == null || t.Val.Value);
    private static bool IsToggleOn(Strike? t) => t != null && (t.Val == null || t.Val.Value);
    private static bool IsToggleOn(DoubleStrike? t) => t != null && (t.Val == null || t.Val.Value);
    private static bool IsToggleOn(Caps? t) => t != null && (t.Val == null || t.Val.Value);
    private static bool IsToggleOn(SmallCaps? t) => t != null && (t.Val == null || t.Val.Value);
    private static bool IsToggleOn(Vanish? t) => t != null && (t.Val == null || t.Val.Value);
    private static bool IsToggleOn(Outline? t) => t != null && (t.Val == null || t.Val.Value);
    private static bool IsToggleOn(Shadow? t) => t != null && (t.Val == null || t.Val.Value);
    private static bool IsToggleOn(Emboss? t) => t != null && (t.Val == null || t.Val.Value);
    private static bool IsToggleOn(Imprint? t) => t != null && (t.Val == null || t.Val.Value);
    private static bool IsToggleOn(NoProof? t) => t != null && (t.Val == null || t.Val.Value);
    // BUG-DUMP-R35-TRBOOL (project-wide): generic CT_OnOff / CT_OnOffOnly reader
    // for the toggle elements that have no typed overload above (tcPr/style/row
    // markers like w:noWrap, w:hideMark, w:semiHidden, w:tblHeader, …). The bare
    // element is ON; an explicit w:val="0"/"false"/"off" is OFF. Reading
    // "element present → true" flips an explicit-OFF marker to ON on dump→batch.
    // OnOffType (OnOffValue Val) and OnOffOnlyType (EnumValue Val) expose
    // different Val CLR types, so read the raw w:val attribute text uniformly.
    private static bool IsToggleOn(OpenXmlElement? e)
    {
        if (e == null) return false;
        foreach (var a in e.GetAttributes())
            if (a.LocalName == "val")
                return a.Value is not ("0" or "false" or "off");
        return true; // bare element (no w:val) = ON
    }

    private DocumentNode GetRootNode(int depth)
    {
        var node = new DocumentNode { Path = "/", Type = "document" };
        var children = new List<DocumentNode>();

        var mainPart = _doc.MainDocumentPart;
        if (mainPart?.Document?.Body != null)
        {
            children.Add(new DocumentNode
            {
                Path = "/body",
                Type = "body",
                ChildCount = mainPart.Document.Body.ChildElements.Count
            });
        }

        if (mainPart?.StyleDefinitionsPart != null)
        {
            children.Add(new DocumentNode
            {
                Path = "/styles",
                Type = "styles",
                ChildCount = mainPart.StyleDefinitionsPart.Styles?.ChildElements.Count ?? 0
            });
        }

        int headerIdx = 0;
        if (mainPart?.HeaderParts != null)
        {
            foreach (var _ in mainPart.HeaderParts)
            {
                children.Add(new DocumentNode
                {
                    Path = $"/header[{headerIdx + 1}]",
                    Type = "header"
                });
                headerIdx++;
            }
        }

        int footerIdx = 0;
        if (mainPart?.FooterParts != null)
        {
            foreach (var _ in mainPart.FooterParts)
            {
                children.Add(new DocumentNode
                {
                    Path = $"/footer[{footerIdx + 1}]",
                    Type = "footer"
                });
                footerIdx++;
            }
        }

        if (mainPart?.NumberingDefinitionsPart != null)
        {
            children.Add(new DocumentNode { Path = "/numbering", Type = "numbering" });
        }

        // CONSISTENCY(footnotes-container): mirror /footnotes/footnote[N] enumeration
        // (Navigation.cs:785) — user entries only (id > 0), excluding separator/
        // continuation system rows so child counts match what `query footnote` returns.
        if (mainPart?.FootnotesPart?.Footnotes != null)
        {
            int fnCount = mainPart.FootnotesPart.Footnotes.Elements<Footnote>()
                .Count(f => f.Id?.Value > 0);
            if (fnCount > 0)
            {
                children.Add(new DocumentNode
                {
                    Path = "/footnotes",
                    Type = "footnotes",
                    ChildCount = fnCount
                });
            }
        }

        if (mainPart?.EndnotesPart?.Endnotes != null)
        {
            int enCount = mainPart.EndnotesPart.Endnotes.Elements<Endnote>()
                .Count(e => e.Id?.Value > 0);
            if (enCount > 0)
            {
                children.Add(new DocumentNode
                {
                    Path = "/endnotes",
                    Type = "endnotes",
                    ChildCount = enCount
                });
            }
        }

        if (mainPart?.WordprocessingCommentsPart?.Comments != null)
        {
            int cCount = mainPart.WordprocessingCommentsPart.Comments.Elements<Comment>().Count();
            if (cCount > 0)
            {
                children.Add(new DocumentNode
                {
                    Path = "/comments",
                    Type = "comments",
                    ChildCount = cCount
                });
            }
        }

        // Core document properties
        var props = _doc.PackageProperties;
        if (props.Title != null) node.Format["title"] = props.Title;
        if (props.Creator != null) node.Format["author"] = props.Creator;
        if (props.Subject != null) node.Format["subject"] = props.Subject;
        if (props.Keywords != null) node.Format["keywords"] = props.Keywords;
        if (props.Description != null) node.Format["description"] = props.Description;
        if (props.Category != null) node.Format["category"] = props.Category;
        if (props.LastModifiedBy != null) node.Format["lastModifiedBy"] = props.LastModifiedBy;
        if (props.Revision != null) node.Format["revisionNumber"] = props.Revision;
        if (props.Created != null) node.Format["created"] = props.Created.Value.ToString("o");
        if (props.Modified != null) node.Format["modified"] = props.Modified.Value.ToString("o");

        // BUG-DUMP10-03: surface the document-level page background color
        // (<w:document><w:background w:color="…"/>…). Without this, dump
        // dropped the page background entirely. Set side already accepts
        // the canonical `background` key (see WordHandler.Add.cs:565).
        if (mainPart?.Document?.GetFirstChild<DocumentBackground>() is { } bgEl
            && bgEl.Color?.Value is { Length: > 0 } bgColor)
        {
            node.Format["background"] = ParseHelpers.FormatHexColor(bgColor);
        }

        // Page size from last section properties (document default)
        var sectPr = mainPart?.Document?.Body?.GetFirstChild<SectionProperties>()
            ?? mainPart?.Document?.Body?.Descendants<SectionProperties>().LastOrDefault();
        if (sectPr != null)
        {
            // CONSISTENCY(root-vs-section-readback): surface the section's
            // <w:sectPrChange> format-revision marker (author/date) at "/" too,
            // matching BuildSectionNode in WordHandler.Query.cs. Without it the
            // dump→batch EmitSection path (which reads from Get("/")) could not
            // fold the marker into the section op and it was dropped on
            // round-trip — unlike tblPrChange/trPrChange/tcPrChange/pPrChange.
            var sectPrChange = sectPr.GetFirstChild<SectionPropertiesChange>();
            if (sectPrChange != null)
            {
                if (!string.IsNullOrEmpty(sectPrChange.Author?.Value))
                    node.Format["sectPrChange.author"] = sectPrChange.Author!.Value!;
                if (sectPrChange.Date?.Value is DateTime sDate)
                    node.Format["sectPrChange.date"] = sDate.ToString("o");
                if (sectPrChange.Id?.Value is { } sId)
                    node.Format["sectPrChange.id"] = sId.ToString();
                // BUG-DUMP-R43-9: carry the verbatim prior-sectPr snapshot.
                var sectPrev = sectPrChange.GetFirstChild<PreviousSectionProperties>();
                if (sectPrev != null && sectPrev.HasChildren)
                    node.Format["sectPrChange.beforeXml"] = sectPrev.OuterXml;
            }
            var pageSize = sectPr.GetFirstChild<PageSize>();
            if (pageSize?.Width?.Value != null) node.Format["pageWidth"] = FormatTwipsToCm(pageSize.Width.Value);
            if (pageSize?.Height?.Value != null) node.Format["pageHeight"] = FormatTwipsToCm(pageSize.Height.Value);
            if (pageSize?.Orient?.Value != null) node.Format["orientation"] = pageSize.Orient.InnerText;
            var margins = sectPr.GetFirstChild<PageMargin>();
            if (margins != null)
            {
                if (margins.Top?.Value != null) node.Format["marginTop"] = FormatTwipsToCm((uint)Math.Abs(margins.Top.Value));
                if (margins.Bottom?.Value != null) node.Format["marginBottom"] = FormatTwipsToCm((uint)Math.Abs(margins.Bottom.Value));
                if (margins.Left?.Value != null) node.Format["marginLeft"] = FormatTwipsToCm(margins.Left.Value);
                if (margins.Right?.Value != null) node.Format["marginRight"] = FormatTwipsToCm(margins.Right.Value);
                // header/footer-from-edge distances + binding gutter (pgMar
                // @header/@footer/@gutter) — needed for dump→batch fidelity.
                if (margins.Header?.Value != null) node.Format["marginHeader"] = FormatTwipsToCm(margins.Header.Value);
                if (margins.Footer?.Value != null) node.Format["marginFooter"] = FormatTwipsToCm(margins.Footer.Value);
                if (margins.Gutter?.Value != null) node.Format["marginGutter"] = FormatTwipsToCm(margins.Gutter.Value);
            }

            // CONSISTENCY(root-vs-section-readback): the body-level sectPr surfaced at /
            // and at /section[N] (for the final section) must yield the same Format keys
            // so set/get round-trips at either path. Mirror BuildSectionNode in
            // WordHandler.Query.cs:786-863 — keep encoding identical (restart maps
            // "newPage"→"restartPage", "newSection"→"restartSection").
            var pgNumType = sectPr.GetFirstChild<PageNumberType>();
            if (pgNumType?.Start?.Value != null)
                node.Format["pageStart"] = pgNumType.Start.Value;
            if (pgNumType?.Format?.Value != null)
                node.Format["pageNumFmt"] = pgNumType.Format.InnerText;
            // BUG-DUMP11-01: w:pgNumType also carries chapStyle (heading style
            // index for chapter numbering) and chapSep (separator between
            // chapter and page numbers). Surfaced here so the body sectPr
            // round-trips chapter-numbering config.
            if (pgNumType?.ChapterStyle?.Value != null)
                node.Format["chapStyle"] = pgNumType.ChapterStyle.Value;
            if (pgNumType?.ChapterSeparator?.Value != null)
                node.Format["chapSep"] = pgNumType.ChapterSeparator.InnerText;

            if (IsToggleOn(sectPr.GetFirstChild<TitlePage>()))
                node.Format["titlePage"] = true;

            // BUG-DUMP-SECT-PAPERSRC: <w:paperSrc w:first/@w:other> selects the
            // printer paper-source bin for the first page vs the rest. Surfaced
            // as dotted paperSrc.first / paperSrc.other so dump→batch round-trips
            // the printer tray config (mirrors the /section[N] readback in
            // WordHandler.Query.cs and the AddSection/Set replay path).
            var bodyPaperSrc = sectPr.GetFirstChild<PaperSource>();
            if (bodyPaperSrc != null)
            {
                if (bodyPaperSrc.First?.Value != null)
                    node.Format["paperSrc.first"] = bodyPaperSrc.First.Value;
                if (bodyPaperSrc.Other?.Value != null)
                    node.Format["paperSrc.other"] = bodyPaperSrc.Other.Value;
            }

            // Surface pgBorders per-side detail (val/sz/color/space) plus the
            // offsetFrom position. Mirrors the paragraph/table per-side border
            // convention (ReadBorder → key + key.sz/.color/.space), keyed under
            // the pgBorders.<side> prefix. The earlier presence-only shorthand
            // (`pgBorders=box`) discarded per-side line style/weight/color and
            // the offsetFrom attribute, so a blue double 1.5pt page border
            // round-tripped to a thin black single line. The `box`/`none`
            // shorthand is still accepted on Set for backward-compat.
            ReadPageBorders(sectPr.GetFirstChild<PageBorders>(), node);

            // Section-level RTL (Arabic / Hebrew page direction).
            if (IsToggleOn(sectPr.GetFirstChild<BiDi>()))
                node.Format["direction"] = "rtl";

            // <w:rtlGutter/> places the binding gutter on the right side.
            if (IsToggleOn(sectPr.GetFirstChild<GutterOnRight>()))
                node.Format["rtlGutter"] = true;

            // BUG-DUMP11-03: <w:noEndnote/> on a section suppresses endnote
            // collection at section end. Bare on/off toggle (no val attr).
            if (IsToggleOn(sectPr.GetFirstChild<NoEndnote>()))
                node.Format["noEndnote"] = true;

            // BUG-DUMP-SECT-FORMPROT: <w:formProt/> locks the section's content
            // except form fields. BUG-DUMP-R40-4: ST_OnOff — bare = ON,
            // <w:formProt w:val="false"/> = OFF. Surface the actual value (was a
            // presence-only test that emitted true for an explicit-false source,
            // flipping protection false→true on round-trip).
            var formProtN = sectPr.GetFirstChild<FormProtection>();
            if (formProtN != null)
                node.Format["formProt"] = formProtN.Val == null || formProtN.Val.Value;

            var lnNum = sectPr.GetFirstChild<LineNumberType>();
            if (lnNum != null)
            {
                var countBy = lnNum.CountBy?.Value ?? 1;
                // BUG-DUMP-SECT-LNDIST: only surface lineNumbers when the source
                // actually carries a @w:restart attr. Defaulting to "continuous"
                // made the emitter fabricate a spurious restart="continuous" on a
                // source that had only @countBy/@distance — round-trip injected an
                // attribute the original lacked. The countBy/distance/start sub-keys
                // below carry the line-number intent independently.
                if (lnNum.Restart?.InnerText is string bodyLnRestart)
                    node.Format["lineNumbers"] = bodyLnRestart switch
                    {
                        "newPage" => "restartPage",
                        "newSection" => "restartSection",
                        _ => "continuous"
                    };
                if (countBy != 1) node.Format["lineNumberCountBy"] = countBy;
                // BUG-DUMP11-02: w:lnNumType/@w:start was silently dropped.
                // Surface as canonical lineNumberStart key.
                if (lnNum.Start?.Value is short lnStart)
                    node.Format["lineNumberStart"] = (int)lnStart;
                // BUG-DUMP-SECT-LNDIST: w:lnNumType/@w:distance (gutter twips
                // between the line-number column and body text) was dropped —
                // exact sibling-attr parallel to @start. Surface as canonical
                // lineNumberDistance.
                if (lnNum.Distance?.Value is string lnDistRaw
                    && int.TryParse(lnDistRaw, out var lnDist))
                    node.Format["lineNumberDistance"] = lnDist;
            }

            // BUG-DUMP11-04: header / footer references (default / first /
            // even) — mirror BuildSectionNode in WordHandler.Query.cs so
            // Get('/') and /section[N] surface the same headerRef.<type> /
            // footerRef.<type> keys.
            if (mainPart != null)
            {
                string? primaryHeader = null;
                foreach (var href in sectPr.Elements<HeaderReference>())
                {
                    if (href.Id?.Value == null) continue;
                    var refType = href.Type?.InnerText ?? "default";
                    try
                    {
                        var part = mainPart.GetPartById(href.Id.Value) as DocumentFormat.OpenXml.Packaging.HeaderPart;
                        if (part != null)
                        {
                            var idx = mainPart.HeaderParts.ToList().IndexOf(part);
                            if (idx >= 0)
                            {
                                var pathRef = $"/header[{idx + 1}]";
                                node.Format[$"headerRef.{refType}"] = pathRef;
                                if (primaryHeader == null || refType == "default") primaryHeader = pathRef;
                            }
                        }
                    }
                    catch { /* dangling rel — skip */ }
                }
                if (primaryHeader != null) node.Format["headerRef"] = primaryHeader;

                string? primaryFooter = null;
                foreach (var fref in sectPr.Elements<FooterReference>())
                {
                    if (fref.Id?.Value == null) continue;
                    var refType = fref.Type?.InnerText ?? "default";
                    try
                    {
                        var part = mainPart.GetPartById(fref.Id.Value) as DocumentFormat.OpenXml.Packaging.FooterPart;
                        if (part != null)
                        {
                            var idx = mainPart.FooterParts.ToList().IndexOf(part);
                            if (idx >= 0)
                            {
                                var pathRef = $"/footer[{idx + 1}]";
                                node.Format[$"footerRef.{refType}"] = pathRef;
                                if (primaryFooter == null || refType == "default") primaryFooter = pathRef;
                            }
                        }
                    }
                    catch { /* dangling rel — skip */ }
                }
                if (primaryFooter != null) node.Format["footerRef"] = primaryFooter;
            }
        }

        // Document protection
        var settings = _doc.MainDocumentPart?.DocumentSettingsPart?.Settings;
        var docProtection = settings?.GetFirstChild<DocumentProtection>();
        if (docProtection != null)
        {
            var editText = docProtection.Edit?.InnerText;
            node.Format["protection"] = editText switch
            {
                "readOnly" => "readOnly",
                "comments" => "comments",
                "trackedChanges" => "trackedChanges",
                "forms" => "forms",
                _ => "none"
            };
            var enforced = docProtection.Enforcement?.Value;
            node.Format["protectionEnforced"] = enforced == true || enforced == null && docProtection.Edit != null;
        }
        else
        {
            node.Format["protection"] = "none";
            node.Format["protectionEnforced"] = false;
        }

        // Document-level settings (DocGrid, CJK, print/display, font embedding, layout flags, columns, etc.)
        PopulateDocSettings(node);
        PopulateCompatibility(node);
        PopulateDocDefaults(node);

        // Theme and Extended Properties
        Core.ThemeHandler.PopulateTheme(_doc.MainDocumentPart?.ThemePart, node);
        Core.ExtendedPropertiesHandler.PopulateExtendedProperties(_doc.ExtendedFilePropertiesPart, node);

        node.Children = children;
        node.ChildCount = children.Count;
        return node;
    }

    private record PathSegment(string Name, int? Index, string? StringIndex = null);

    /// <summary>
    /// Resolve InsertPosition (After/Before anchor path) to a 0-based int? index.
    /// Anchor path can be full (/body/p[@paraId=xxx]) or short (p[@paraId=xxx]).
    /// </summary>
    private int? ResolveAnchorPosition(OpenXmlElement parent, string parentPath, InsertPosition? position)
    {
        if (position == null) return null;
        if (position.Index.HasValue) return position.Index;

        var anchorPath = position.After ?? position.Before!;

        // Catch bare attribute selector without element wrapper, e.g. @paraId=XXX instead of p[@paraId=XXX]
        if (System.Text.RegularExpressions.Regex.IsMatch(anchorPath, @"^@(\w+)=(.+)$"))
            throw new ArgumentException($"Invalid anchor path \"{anchorPath}\". Did you mean: p[{anchorPath}]?");

        // Handle find: prefix — text-based anchoring within a paragraph
        if (anchorPath.StartsWith("find:", StringComparison.OrdinalIgnoreCase))
        {
            // Return a sentinel value; actual handling done in Add via AddAtFindPosition
            return FindAnchorIndex;
        }

        // Normalize: if short form (no leading /), prepend parentPath
        if (!anchorPath.StartsWith("/"))
            anchorPath = parentPath.TrimEnd('/') + "/" + anchorPath;

        // Top-level /watermark[N]? special case. Watermarks are stored in
        // the header parts, not the body — there is no body-level sibling
        // that represents the watermark. `add --type watermark` returns
        // "/watermark" as the new element's identity; to keep that path
        // round-trippable as --after/--before, treat it as a no-op
        // positional hint: --after /watermark appends to parent, --before
        // /watermark prepends. Callers needing a specific body position
        // should pass an explicit /body/p[N] anchor instead.
        {
            var wmMatch = System.Text.RegularExpressions.Regex.Match(anchorPath, @"^/watermark(?:\[(\d+)\])?$", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            if (wmMatch.Success)
            {
                // Honour the positional-hint contract only when a watermark
                // actually exists in the doc. Otherwise fall through so the
                // standard "Anchor element not found" error fires — matching
                // /chart[1] and other absent-anchor behaviour. An explicit
                // index beyond the number of watermarks (there's at most one)
                // is out-of-range — error instead of silently appending.
                var wmExists = FindWatermark() != null;
                var wmCount = wmExists ? 1 : 0;
                if (wmMatch.Groups[1].Success)
                {
                    var wmIdx = int.Parse(wmMatch.Groups[1].Value);
                    if (wmIdx < 1 || wmIdx > wmCount)
                        throw new ArgumentException($"Anchor element not found: {anchorPath}");
                }
                else if (!wmExists)
                {
                    throw new ArgumentException($"Anchor element not found: {anchorPath}");
                }
                return position.After != null ? (int?)null : 0;
            }
        }

        // Virtual table column anchor: /body/tbl[N]/col[N]. ParsePath would
        // fail because <w:col> doesn't exist in OOXML. Used by `add column
        // --before/--after col[K]` and `add --from col[K] --before/--after col[J]`.
        // Validates that the anchor exists in the named table.
        {
            var colAnchorMatch = System.Text.RegularExpressions.Regex.Match(
                anchorPath, @"^/body/tbl\[(\d+)\]/col\[(\d+)\]$");
            if (colAnchorMatch.Success)
            {
                var anchorTableIdx = int.Parse(colAnchorMatch.Groups[1].Value);
                var anchorColIdx = int.Parse(colAnchorMatch.Groups[2].Value);
                var body = _doc.MainDocumentPart?.Document?.Body;
                var tables = body?.Elements<Table>().ToList() ?? new List<Table>();
                if (anchorTableIdx < 1 || anchorTableIdx > tables.Count)
                    throw new ArgumentException($"Anchor table not found: {anchorPath} (total tables at /body: {tables.Count})");
                var anchorGrid = tables[anchorTableIdx - 1].GetFirstChild<TableGrid>();
                var gridColCount = anchorGrid?.Elements<GridColumn>().Count() ?? 0;
                if (anchorColIdx < 1 || anchorColIdx > gridColCount)
                    throw new ArgumentException($"Anchor column not found: {anchorPath} (total columns: {gridColCount})");
                return position.After != null ? anchorColIdx : anchorColIdx - 1;
            }
        }

        var segments = ParsePath(anchorPath);
        var anchor = NavigateToElement(segments, out var ctx)
            ?? throw new ArgumentException($"Anchor element not found: {anchorPath}" + (ctx != null ? $". {ctx}" : ""));

        // Body-level <w:sectPr> (direct child of Body) must remain the last
        // child of body. `--after /body/sectPr` has no valid placement;
        // silently routing to "before sectPr" (the old behaviour) misleads
        // the caller. Reject with a clear error. Paragraph-level sectPr
        // (inside w:pPr) is unaffected — its carrier paragraph is the
        // anchor, not the sectPr itself.
        if (position.After != null && anchor is SectionProperties && anchor.Parent is Body)
        {
            throw new ArgumentException(
                "Cannot insert after body-level sectPr; it must remain the last child of body. " +
                "Use --before /body/sectPr (or omit the anchor to append before sectPr).");
        }

        // Find anchor's position among parent's children
        var siblings = parent.ChildElements.ToList();
        // /body/oMathPara[N] resolves to the inner M.Paragraph/oMathPara element;
        // when it lives inside a pure wrapper w:p, the wrapper is the actual
        // body child. Re-target the anchor to that wrapper so --after/--before
        // can find it among body siblings.
        if ((anchor is M.Paragraph || anchor.LocalName == "oMathPara")
            && anchor.Parent is Paragraph wrapAnchor
            && IsOMathParaWrapperParagraph(wrapAnchor)
            && parent.ChildElements.Contains(wrapAnchor))
        {
            anchor = wrapAnchor;
        }
        var anchorIdx = siblings.IndexOf(anchor);
        if (anchorIdx < 0)
            throw new ArgumentException($"Anchor element is not a child of {parentPath}: {anchorPath}");

        // CONSISTENCY(table-row-anchor): when inserting into a <w:tbl>, the
        // body's child list also contains tblPr / tblGrid / tblPrEx, but
        // AddRow indexes against parent.Elements<TableRow>() — using the
        // ChildElements offset there would push past the tail and silently
        // AppendChild. Translate the anchor's position into row-only space
        // so the AddRow contract (index = row-only index) holds.
        if (parent is Table tbl && anchor is TableRow trAnchor)
        {
            var rows = tbl.Elements<TableRow>().ToList();
            var rowIdx = rows.IndexOf(trAnchor);
            if (rowIdx < 0)
                throw new ArgumentException($"Anchor row is not a row of {parentPath}: {anchorPath}");
            if (position.After != null)
                return rowIdx + 1 >= rows.Count ? null : rowIdx + 1;
            return rowIdx;
        }

        if (position.After != null)
        {
            // Insert after anchor: if last child, return null (append)
            return anchorIdx + 1 >= siblings.Count ? null : anchorIdx + 1;
        }
        else
        {
            // Insert before anchor
            return anchorIdx;
        }
    }

    /// <summary>Sentinel value indicating find: anchor needs text-based resolution.</summary>
    private const int FindAnchorIndex = -99999;

    /// <summary>
    /// Build an SDT path segment using @sdtId= if available, otherwise positional index.
    /// </summary>
    private static string BuildSdtPathSegment(OpenXmlElement sdt, int positionalIndex)
    {
        var sdtProps = (sdt is SdtBlock sb ? sb.SdtProperties : (sdt as SdtRun)?.SdtProperties);
        var sdtIdVal = sdtProps?.GetFirstChild<SdtId>()?.Val?.Value;
        return sdtIdVal != null
            ? $"sdt[@sdtId={sdtIdVal}]"
            : $"sdt[{positionalIndex}]";
    }

    /// <summary>
    /// Build a paragraph path segment using @paraId= if available, otherwise positional index.
    /// E.g. "p[@paraId=1A2B3C4D]" or "p[3]".
    /// </summary>
    private static string BuildParaPathSegment(Paragraph para, int positionalIndex)
    {
        var paraId = para.ParagraphId?.Value;
        return !string.IsNullOrEmpty(paraId)
            ? $"p[@paraId={paraId}]"
            : $"p[{positionalIndex}]";
    }

    // 1-based position of a just-appended paragraph among its parent's Paragraph
    // children, for naming the result path. O(1) for the body via the
    // append-monotonic cache; other parents Count() directly (not the hot path).
    // Self-validating: the incremental count is trusted only when the paragraph
    // cached last time is the element immediately preceding this freshly-appended
    // one — any out-of-band insert/remove/non-paragraph sibling breaks the chain
    // and forces one O(n) reseed, so no mutation site must invalidate manually.
    // Also refreshes _lastBodyParagraph to keep /body/p[last()] navigation warm.
    private int AppendBodyParaFast(OpenXmlElement parent, Paragraph para)
    {
        // The Open XML SDK keeps children in a singly-linked list, so
        // InsertBefore(refChild) and PreviousSibling() scan for the predecessor
        // — O(N) each, making an N-paragraph append O(N²). InsertAfterSelf and
        // NextSibling are O(1). So append right after the cached last body
        // paragraph (whose NextSibling is the trailing sectPr) and bump the
        // cached count. Cold cache or an out-of-band mutation (NextSibling no
        // longer the sectPr) falls back to the O(N) append + recount, which then
        // reseeds the cache for the next run.
        if (parent is Body fastBody && _bodyParaCount >= 0
            && _lastBodyParagraph is Paragraph anchor
            && ReferenceEquals(anchor.Parent, fastBody))
        {
            // The insertion point is "immediately before the trailing sectPr".
            // Usually that is right after the cached last paragraph, but a
            // non-paragraph block appended since (most commonly a table, in a
            // heading/paragraph/list/TABLE interleave) now sits between it and
            // the sectPr, so anchor.NextSibling() is no longer the sectPr. Walk
            // forward over those few intervening blocks to the last child before
            // the sectPr and InsertAfterSelf there — still O(1) amortized (the
            // hop count is the non-paragraph blocks added since the last
            // paragraph, ~1 per table). This replaces the old cold path that
            // fell back to InsertBefore + Elements<Paragraph>().Count(), both
            // O(N): one such recount per interleaved table made an N-block mixed
            // document O(N²). _bodyParaCount stays valid because a table append
            // does not change the body-direct paragraph count.
            OpenXmlElement tail = anchor;
            var next = tail.NextSibling();
            while (next != null && next is not SectionProperties)
            {
                tail = next;
                next = tail.NextSibling();
            }
            if (next is SectionProperties)
            {
                tail.InsertAfterSelf(para);
                _lastBodyParagraph = para;
                return ++_bodyParaCount;
            }
        }
        AppendToParent(parent, para);
        if (parent is Body coldBody)
        {
            _bodyParaCount = coldBody.Elements<Paragraph>().Count();
            _lastBodyParagraph = para;
            return _bodyParaCount;
        }
        return parent.Elements<Paragraph>().Count();
    }

    private static List<PathSegment> ParsePath(string path)
    {
        var segments = new List<PathSegment>();
        // Reject leading double-slash up front — the subsequent Trim('/') would
        // otherwise eat the second slash and silently resolve "//body" → /body,
        // "//header[1]" → /header[1], producing inconsistent behavior next to
        // "//section[1]" which already errors out as Path-not-found via the regex
        // dispatch. The earlier-dispatch regexes anchor on `^/` so they don't
        // match `^//…` either; failures fall through here and we now reject.
        if (path.StartsWith("//"))
            throw new ArgumentException(
                $"Malformed path '{path}'. Path must start with exactly one '/'.");
        // Reject trailing slash up front — the subsequent Trim('/') would
        // otherwise silently absorb it and produce a path that looks valid
        // (e.g. "/body/p[1]/" → "body/p[1]") while any callers
        // concatenating onto the raw input would end up with doubled
        // separators like "/body/p[1]//r[2]" in the returned path.
        if (path.Length > 1 && path.EndsWith("/"))
            throw new ArgumentException(
                $"Malformed path '{path}'. Trailing '/' is not allowed.");
        var parts = path.Trim('/').Split('/');

        // BUG-DUMP-R33-STYLEID: the segment immediately after "/styles" is a
        // style ID, not an OOXML element type, so it must NOT go through the
        // element-type alias map. A style whose id is "paragraph" / "run" /
        // "table" (real-world docs author these) otherwise had its path segment
        // rewritten ("paragraph" -> "p"), and "/styles/paragraph" then resolved
        // as "find a <w:p> under /styles" — failing every Add/Set/Get on that
        // style (e.g. adding a tab stop to it dropped on dump->batch replay).
        string prevSegName = "";

        foreach (var part in parts)
        {
            bool afterStyles = prevSegName == "styles";
            // Reject degenerate empty segments from trailing/duplicate slashes
            // (e.g. "/body/p[1]/" or "/body//p[1]"). Without this, ParsePath
            // would silently swallow the empty part and return a garbled
            // navigable path.
            if (part.Length == 0)
                throw new ArgumentException(
                    $"Malformed path '{path}'. Empty path segment (check for trailing or duplicate '/').");

            var bracketIdx = part.IndexOf('[');
            if (bracketIdx >= 0)
            {
                // Only single-predicate form is supported. Reject malformed
                // selectors like "p[1][2]" or "p[1]trailing" where content
                // follows the first closing ']'. Without this the trailing
                // junk is silently swallowed (e.g. "p[1][2]" would resolve
                // to "p[1]") which hides typos.
                if (!part.EndsWith("]"))
                    throw new ArgumentException(
                        $"Malformed path segment '{part}'. Expected 'name[index]' or 'name[@attr=value]'.");
                var firstClose = part.IndexOf(']');
                if (firstClose != part.Length - 1)
                    throw new ArgumentException(
                        $"Malformed path segment '{part}'. Multiple predicates are not supported — use a single 'name[...]' form.");

                var rawName = part[..bracketIdx];
                var name = afterStyles ? rawName : Core.PathAliases.Resolve(rawName);
                var indexStr = part[(bracketIdx + 1)..^1];
                // Reject empty predicate "p[]" which Int32.TryParse silently
                // rejects but which then falls through as a StringIndex of "".
                if (indexStr.Length == 0)
                    throw new ArgumentException(
                        $"Malformed path segment '{part}'. Empty predicate — expected 'name[index]' or 'name[@attr=value]'.");
                if (int.TryParse(indexStr, out var idx))
                {
                    if (idx <= 0)
                        throw new ArgumentException(
                            $"Malformed path segment '{part}'. Index predicate must be a positive integer (1-based), got '{indexStr}'.");
                    segments.Add(new PathSegment(name, idx));
                }
                else
                {
                    // Only accept a tightly specified set of string predicates:
                    //   last()
                    //   @attr=value   where attr is a simple identifier
                    //                 ([A-Za-z_][A-Za-z0-9_]*) and value is
                    //                 either bare-word (no whitespace, not
                    //                 starting with '@' or quote) or
                    //                 double-quoted.
                    // Anything else (e.g. "XYZ", " 1", "@=X", "@paraId",
                    //   "@w:paraId=X", "@attr='X'") is rejected up front so
                    //   typos cannot silently hit the FirstOrDefault()
                    //   fallback in NavigateToElement.
                    var normalizedPredicate = ValidateAndNormalizePredicate(part, indexStr);
                    segments.Add(new PathSegment(name, null, normalizedPredicate));
                }
                prevSegName = name;
            }
            else
            {
                var name = afterStyles ? part : Core.PathAliases.Resolve(part);
                segments.Add(new PathSegment(name, null));
                prevSegName = name;
            }
        }

        return segments;
    }

    /// <summary>
    /// Validate a string predicate (the content inside [...] that isn't an
    /// integer) and return its normalized form. Accepted grammar:
    ///   last()
    ///   @ident=value            (bare value: no whitespace, no quotes, no '@')
    ///   @ident="quoted value"   (double-quoted value)
    /// Everything else throws ArgumentException so typos like "p[XYZ]",
    /// "p[ 1]", "p[@paraId]" (no =), "p[@=X]", "p[@w:paraId=X]" are rejected
    /// instead of silently falling through to childList.FirstOrDefault().
    /// </summary>
    private static string ValidateAndNormalizePredicate(string part, string predicate)
    {
        if (predicate == "last()")
            return predicate;

        if (predicate.Length > 0 && predicate[0] == '@')
        {
            // Must have '=' and a non-empty identifier before it.
            var eq = predicate.IndexOf('=');
            if (eq <= 1)
                throw new ArgumentException(
                    $"Malformed path segment '{part}'. Attribute predicate must be '[@name=value]' with a non-empty attribute name.");

            var attr = predicate[1..eq];
            // Simple identifier: [A-Za-z_][A-Za-z0-9_]*
            if (!System.Text.RegularExpressions.Regex.IsMatch(attr, "^[A-Za-z_][A-Za-z0-9_]*$"))
                throw new ArgumentException(
                    $"Malformed path segment '{part}'. Attribute name '{attr}' is not a simple identifier (no prefixes/colons).");

            var value = predicate[(eq + 1)..];
            if (value.Length == 0)
                throw new ArgumentException(
                    $"Malformed path segment '{part}'. Attribute predicate value is empty.");

            // Accept double-quoted value — strip quotes so downstream
            // comparisons (which use bare string equality) work uniformly.
            if (value.Length >= 2 && value[0] == '"' && value[^1] == '"')
            {
                var inner = value[1..^1];
                if (inner.Contains('"'))
                    throw new ArgumentException(
                        $"Malformed path segment '{part}'. Quoted attribute value must not contain embedded double quotes.");
                return $"@{attr}={inner}";
            }

            // Bare value: no whitespace, no quotes, no leading '@'.
            if (value[0] == '@' || value[0] == '\'' || value[0] == '"')
                throw new ArgumentException(
                    $"Malformed path segment '{part}'. Attribute value must be bare-word or double-quoted.");
            foreach (var c in value)
            {
                if (char.IsWhiteSpace(c))
                    throw new ArgumentException(
                        $"Malformed path segment '{part}'. Attribute value must not contain whitespace (use double quotes).");
            }
            return predicate;
        }

        throw new ArgumentException(
            $"Malformed path segment '{part}'. Predicate must be a positive integer, 'last()', or '[@attr=value]'.");
    }

    // PERF: cache the flattened+filtered body paragraph/table lists per Body
    // instance. /body/p[N] and /body/tbl[N] are resolved by index; without
    // the cache, dumping a 14k-paragraph doc made 14k Get calls × 14k walks
    // → O(n²). Invalidation is by explicit clear: ClearBodyChildIndex() is
    // called on every structural add (Add() entry) and on Remove/Move/Swap/
    // CopyFrom/raw-set (via InvalidateBodyParaCache / RawSet). Property-only
    // Set calls do NOT clear (correct — they don't change which paragraph sits
    // at index N), so a batch of set /body/p[N] keeps hitting the cache.
    //
    // The guard must NOT recompute body.ChildElements.Count: the SDK stores
    // children in a singly-linked list, so .Count is O(n), which made even a
    // cache *hit* O(n) and turned batch set/get of /body/p[N] back into O(n²).
    private readonly Dictionary<OpenXmlElement, (List<OpenXmlElement> paras, List<OpenXmlElement> tables)>
        _bodyChildIndexCache = new();

    // Per-body-child → owning section, built lazily by FindOwningSectionProperties.
    private Dictionary<OpenXmlElement, SectionProperties?>? _owningSectionCache;

    // Per-body paraId → paragraph, built lazily by GetBodyParaById. Resolving
    // /body/p[@paraId=X] via FirstOrDefault scan is O(n); dump emits one
    // Get(/body/p[@paraId]) per paragraph, so that scan made dump O(n²).
    private Dictionary<OpenXmlElement, Dictionary<string, Paragraph>>? _bodyParaByIdCache;

    // Per-scope-root bookmark w:id → BookmarkStart, built lazily by
    // FindBookmarkStartById. ResolveBookmarkEndName / IsContentSpanBookmark(end)
    // resolved a standalone <w:bookmarkEnd> to its paired start via
    // body.Descendants<BookmarkStart>().FirstOrDefault(id) — O(bookmarks) per
    // call. dump emits one such lookup per bookmarkEnd, so a document with N
    // bookmarks cost O(N²) (a 6940-bookmark FedRAMP SSP spent tens of seconds
    // here alone). Keyed by scope root (Body) so the map is built once.
    private Dictionary<OpenXmlElement, Dictionary<string, BookmarkStart>>? _bookmarkStartByIdCache;

    // Per-container child-index caches for the SAME O(n²) shape the body caches
    // above fix, but one level down: resolving /<para>/r[K], /<tbl>/tr[K] and
    // /<tr>/tc[K] for K=1..M re-materialized the (filtered/flattened) child list
    // on EVERY navigation. dump emits one Get/GetElementXml per run/row/cell, so
    // a run-dense paragraph (thousands of char-level runs from a PDF conversion)
    // or a large table cost O(M²). Key = the container element (paragraph/table/
    // row); value = the same List the uncached path produced, so ElementAt is
    // O(1) and the `as List` reuse at the navigation tail avoids a per-call copy.
    // Invalidation piggybacks on ClearBodyChildIndex (coarse: any structural
    // mutation drops all three — dump is read-only and hit-heavy, mutations rare;
    // "clear too much" is safe, "clear too little" is the only correctness bug).
    private readonly Dictionary<OpenXmlElement, List<OpenXmlElement>> _navRunIndexCache = new();
    private readonly Dictionary<OpenXmlElement, List<OpenXmlElement>> _navRowIndexCache = new();
    private readonly Dictionary<OpenXmlElement, List<OpenXmlElement>> _navCellIndexCache = new();

    // Drop the body child-index + owning-section + paraId caches after a
    // structural mutation. Called from Add() (body-level) and InvalidateBodyParaCache.
    private void ClearBodyChildIndex()
    {
        _bodyChildIndexCache.Clear();
        _owningSectionCache = null;
        _bodyParaByIdCache = null;
        _bookmarkStartByIdCache = null;
        ClearNavChildCaches();
    }

    // Drop the run / row / cell child-index caches. Split out from
    // ClearBodyChildIndex because these must invalidate on a WIDER set of
    // mutations than the body-direct caches: adding/removing a run under a
    // PARAGRAPH (parent != Body) leaves the body child-index valid but makes a
    // cached /<para>/r[K] list stale. Add() therefore calls this unconditionally
    // at entry (not gated on `parent is Body`), while body-direct mutations reach
    // it via ClearBodyChildIndex. A bare Dictionary.Clear() is O(1)-ish and dump
    // (the hit-heavy path) never mutates, so over-clearing costs nothing.
    private void ClearNavChildCaches()
    {
        _navRunIndexCache.Clear();
        _navRowIndexCache.Clear();
        _navCellIndexCache.Clear();
    }

    // Cached, filtered run list for /<container>/r[K] resolution. The filter is
    // byte-identical to the inline `"r" =>` switch arm it replaces (skip runs
    // inside SdtRun / SimpleField / a nested textbox / an mc:AlternateContent
    // wrapper, and comment-reference runs) so the run set NodeBuilder surfaces
    // stays aligned. Key = `container` (the element the r[K] index is relative
    // to) because the textbox/AlternateContent skips are expressed relative to it.
    private List<OpenXmlElement> GetNavRunIndex(OpenXmlElement container)
    {
        if (_navRunIndexCache.TryGetValue(container, out var cached))
            return cached;
        var runs = container.Descendants<Run>()
            .Where(r => r.GetFirstChild<CommentReference>() == null)
            .Where(r => r.Ancestors<SdtRun>().FirstOrDefault() == null)
            .Where(r => r.Ancestors<SimpleField>().FirstOrDefault() == null)
            .Where(r =>
            {
                var tbc = r.Ancestors<TextBoxContent>().FirstOrDefault();
                if (tbc == null) return true;
                foreach (var anc in tbc.Ancestors())
                {
                    if (ReferenceEquals(anc, container)) return false;
                }
                return true;
            })
            .Where(r =>
            {
                foreach (var anc in r.Ancestors())
                {
                    if (ReferenceEquals(anc, container)) break;
                    if (anc.LocalName == "AlternateContent"
                        || anc.LocalName == "Choice"
                        || anc.LocalName == "Fallback")
                        return false;
                }
                return true;
            })
            .Cast<OpenXmlElement>()
            .ToList();
        _navRunIndexCache[container] = runs;
        return runs;
    }

    private List<OpenXmlElement> GetNavRowIndex(Table table)
    {
        if (_navRowIndexCache.TryGetValue(table, out var cached))
            return cached;
        var rows = GetTableRowsFlattened(table).Cast<OpenXmlElement>().ToList();
        _navRowIndexCache[table] = rows;
        return rows;
    }

    private List<OpenXmlElement> GetNavCellIndex(TableRow row)
    {
        if (_navCellIndexCache.TryGetValue(row, out var cached))
            return cached;
        var cells = GetRowCellsFlattened(row).Cast<OpenXmlElement>().ToList();
        _navCellIndexCache[row] = cells;
        return cells;
    }

    // O(1) /body/p[@paraId=X] over body-direct (incl. customXml) paragraphs.
    // Returns null for ids that live only inside tables/sdt — the caller falls
    // back to a Descendants scan for those (rare).
    private Paragraph? GetBodyParaById(Body body, string paraId)
    {
        _bodyParaByIdCache ??= new();
        if (!_bodyParaByIdCache.TryGetValue(body, out var map))
        {
            map = new(StringComparer.OrdinalIgnoreCase);
            foreach (var p in GetBodyParagraphIndex(body).OfType<Paragraph>())
                if (p.ParagraphId?.Value is { } id && !map.ContainsKey(id))
                    map[id] = p; // first-wins, matches the old FirstOrDefault
            _bodyParaByIdCache[body] = map;
        }
        return map.TryGetValue(paraId, out var found) ? found : null;
    }

    private List<OpenXmlElement> GetBodyParagraphIndex(Body body) => GetBodyChildIndex(body).paras;
    private List<OpenXmlElement> GetBodyTableIndex(Body body) => GetBodyChildIndex(body).tables;

    private (List<OpenXmlElement> paras, List<OpenXmlElement> tables) GetBodyChildIndex(Body body)
    {
        if (_bodyChildIndexCache.TryGetValue(body, out var entry))
            return (entry.paras, entry.tables);

        var flat = new List<OpenXmlElement>();
        void Collect(OpenXmlElement el)
        {
            foreach (var c in el.ChildElements)
            {
                if (c is CustomXmlBlock cx) Collect(cx);
                else flat.Add(c);
            }
        }
        Collect(body);

        var paras = new List<OpenXmlElement>();
        var tables = new List<OpenXmlElement>();
        foreach (var e in flat)
        {
            if (e is Paragraph p && !IsOMathParaWrapperParagraph(p)) paras.Add(p);
            else if (e is Table t) tables.Add(t);
        }
        _bodyChildIndexCache[body] = (paras, tables);
        return (paras, tables);
    }

    /// <summary>
    /// Return the raw OuterXml of the element at <paramref name="path"/>, or
    /// null if the path does not resolve. Unlike <see cref="Raw"/> (which only
    /// addresses whole package parts), this navigates the document DOM. Used by
    /// the dump emitter to round-trip rich block containers — a Table-of-Contents
    /// SDT, say — verbatim instead of flattening them to text.
    /// </summary>
    internal string? RawElementXml(string path)
    {
        try
        {
            var segments = ParsePath(path);
            return NavigateToElement(segments)?.OuterXml;
        }
        catch
        {
            return null;
        }
    }

    // BUG-DUMP-R35-2: recover the whitespace-preserved text of each <w:r> that
    // sits inside a <w:smartTag>/<w:customXml> wrapper in <paramref name="para"/>,
    // in document order. smartTag/customXml parse as OpenXmlUnknownElement and
    // the SDK's unknown-element reader discards a run's insignificant whitespace
    // (a `<w:t> </w:t>` with no xml:space="preserve" collapses to `<w:t/>`), so
    // the space is unrecoverable from the parsed DOM. The part stream still has
    // it; re-parse it whitespace-preserved (System.Xml.Linq) and return the
    // wrapper-run texts so the synthesizer can fill in a collapsed run.
    //
    // The matching raw paragraph is located by its NON-EMPTY wrapper-run text
    // fingerprint (those survive in both DOM and raw, so the sequence is a stable
    // key) — robust against paragraphs living in tables/headers without needing a
    // positional path. Returns an empty list on any failure (caller falls back to
    // the parsed DOM text, i.e. the prior behavior).
    private List<string> RecoverWrapperRunTexts(Paragraph para)
    {
        var empty = new List<string>();
        const string wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        // DOM-side fingerprint: the non-empty wrapper-run texts of this paragraph.
        var domKey = new List<string>();
        bool paraHasWrapper = false;
        foreach (var unkRun in para.Descendants<DocumentFormat.OpenXml.OpenXmlUnknownElement>())
        {
            if (unkRun.LocalName != "r" || unkRun.NamespaceUri != wNs) continue;
            bool inWrapper = false;
            for (var anc = unkRun.Parent; anc != null && anc != para; anc = anc.Parent)
                if (anc is DocumentFormat.OpenXml.OpenXmlUnknownElement uw
                    && uw.NamespaceUri == wNs
                    && (uw.LocalName == "smartTag" || uw.LocalName == "customXml"))
                { inWrapper = true; break; }
            if (!inWrapper) continue;
            paraHasWrapper = true;
            var t = unkRun.InnerText;
            // Fingerprint on NON-WHITESPACE runs only: the parsed DOM has already
            // dropped insignificant-whitespace runs (the exact loss we recover),
            // so a whitespace-only run never appears here. The raw side must
            // exclude them too or the counts diverge and no paragraph matches.
            if (!string.IsNullOrWhiteSpace(t)) domKey.Add(t);
        }
        if (!paraHasWrapper) return empty;

        try
        {
            var part = para.Ancestors<DocumentFormat.OpenXml.OpenXmlPartRootElement>()
                .FirstOrDefault()?.OpenXmlPart
                ?? (OpenXmlPart?)_doc.MainDocumentPart;
            if (part == null) return empty;
            System.Xml.Linq.XDocument xdoc;
            using (var stream = part.GetStream(System.IO.FileMode.Open, System.IO.FileAccess.Read))
                xdoc = System.Xml.Linq.XDocument.Load(stream,
                    System.Xml.Linq.LoadOptions.PreserveWhitespace);
            System.Xml.Linq.XNamespace w = wNs;
            foreach (var rawP in xdoc.Descendants(w + "p"))
            {
                // Collect this raw paragraph's wrapper-run texts (all, in order)
                // plus its non-empty fingerprint for matching.
                var allTexts = new List<string>();
                var nonEmpty = new List<string>();
                foreach (var rawR in rawP.Descendants(w + "r"))
                {
                    // Only runs whose ancestor chain (within rawP) includes a
                    // smartTag/customXml wrapper.
                    bool inWrap = rawR.Ancestors()
                        .TakeWhile(a => a != rawP)
                        .Any(a => a.Name == w + "smartTag" || a.Name == w + "customXml")
                        || rawR.Ancestors(w + "smartTag").Any()
                        || rawR.Ancestors(w + "customXml").Any();
                    if (!inWrap) continue;
                    var txt = string.Concat(rawR.Descendants(w + "t").Select(t => t.Value));
                    allTexts.Add(txt);
                    if (!string.IsNullOrWhiteSpace(txt)) nonEmpty.Add(txt);
                }
                if (nonEmpty.Count == domKey.Count
                    && nonEmpty.SequenceEqual(domKey, StringComparer.Ordinal))
                    return allTexts;
            }
        }
        catch { /* fall back to parsed DOM text */ }
        return empty;
    }

    private OpenXmlElement? NavigateToElement(List<PathSegment> segments)
        => NavigateToElement(segments, out _, out _);

    private OpenXmlElement? NavigateToElement(List<PathSegment> segments, out string? availableContext)
        => NavigateToElement(segments, out availableContext, out _);

    private OpenXmlElement? NavigateToElement(List<PathSegment> segments, out string? availableContext, out string resolvedPath)
    {
        resolvedPath = "";
        availableContext = null;
        if (segments.Count == 0) return null;

        var first = segments[0];

        // Handle bookmark[@name=...] as top-level path
        if (first.Name.ToLowerInvariant() == "bookmark" && first.StringIndex != null
            && first.StringIndex.StartsWith("@name=", StringComparison.OrdinalIgnoreCase))
        {
            var targetName = first.StringIndex["@name=".Length..];
            var body = _doc.MainDocumentPart?.Document?.Body;
            return body?.Descendants<BookmarkStart>()
                .FirstOrDefault(b => b.Name?.Value == targetName);
        }

        // Handle /bookmark[N] (1-based positional, document order). Skips
        // _GoBack and other reserved bookmarks (names starting with '_') so
        // the index matches what `query bookmark` returns.
        if (first.Name.ToLowerInvariant() == "bookmark" && segments.Count == 1
            && first.Index.HasValue)
        {
            var body = _doc.MainDocumentPart?.Document?.Body;
            if (body != null)
            {
                var bks = body.Descendants<BookmarkStart>()
                    .Where(b => !(b.Name?.Value ?? "").StartsWith("_", StringComparison.Ordinal))
                    .ToList();
                var n = first.Index.Value;
                if (n >= 1 && n <= bks.Count) return bks[n - 1];
            }
        }

        // BUG-R36-B5: top-level /sdt[N] alias. The schema documents both
        // /sdt[N] and /body/p[N]/sdt[M], but only the body-anchored form
        // resolved. Resolve /sdt[N] positionally over body-level SdtBlock
        // elements (document order), mirroring the /bookmark[N] alias above.
        if (first.Name.ToLowerInvariant() == "sdt" && segments.Count == 1
            && first.Index.HasValue)
        {
            var body = _doc.MainDocumentPart?.Document?.Body;
            if (body != null)
            {
                var sdts = body.Descendants<SdtBlock>().Cast<OpenXmlElement>()
                    .Concat(body.Descendants<SdtRun>().Cast<OpenXmlElement>())
                    .ToList();
                var n = first.Index.Value;
                if (n >= 1 && n <= sdts.Count) return sdts[n - 1];
            }
        }
        if (first.Name.ToLowerInvariant() == "sdt" && segments.Count == 1
            && first.StringIndex != null
            && first.StringIndex.StartsWith("@sdtId=", StringComparison.OrdinalIgnoreCase))
        {
            var body = _doc.MainDocumentPart?.Document?.Body;
            if (body != null
                && int.TryParse(first.StringIndex["@sdtId=".Length..], out var targetId))
            {
                return body.Descendants<SdtBlock>().Cast<OpenXmlElement>()
                    .Concat(body.Descendants<SdtRun>().Cast<OpenXmlElement>())
                    .FirstOrDefault(s =>
                        (s as SdtBlock)?.SdtProperties?.GetFirstChild<SdtId>()?.Val?.Value == targetId
                        || (s as SdtRun)?.SdtProperties?.GetFirstChild<SdtId>()?.Val?.Value == targetId);
            }
        }

        // Top-level /section[N] anchor routing. `add --type section` returns
        // "/section[N]" as the new element's identity; resolving it to the
        // carrier paragraph (the one whose pPr holds the Nth sectPr) lets
        // callers use it directly as --after/--before. Body-level sectPr
        // (the final section) is intentionally NOT an anchor target here —
        // it must remain the last child of body; anchor use is rejected in
        // ResolveAnchorPosition.
        if (first.Name.ToLowerInvariant() == "section" && segments.Count == 1 && first.Index.HasValue)
        {
            var body = _doc.MainDocumentPart?.Document?.Body;
            if (body != null)
            {
                var n = first.Index.Value;
                var sectParas = body.Elements<Paragraph>()
                    .Where(p => p.ParagraphProperties?.GetFirstChild<SectionProperties>() != null)
                    .ToList();
                if (n >= 1 && n <= sectParas.Count)
                    return sectParas[n - 1];
            }
        }

        // Top-level /chart[N] anchor routing. `add --type chart` returns
        // "/chart[N]" as the new element's identity; resolve it to the
        // body-level paragraph containing the Nth chart drawing so callers
        // can use the returned path directly as --after/--before.
        if (first.Name.ToLowerInvariant() == "chart" && segments.Count == 1 && first.Index.HasValue)
        {
            var charts = GetAllWordCharts();
            var n = first.Index.Value;
            if (n >= 1 && n <= charts.Count)
            {
                OpenXmlElement? cur = charts[n - 1].Container;
                while (cur != null && cur is not Paragraph) cur = cur.Parent;
                if (cur is Paragraph chartPara) return chartPara;
            }
        }

        // Top-level /toc[N] anchor routing. `add --type toc` returns
        // "/toc[N]" as the new element's identity; resolve it to the Nth
        // body paragraph whose descendants include a FieldCode starting
        // with "TOC" (mirrors AddToc's counting logic) so callers can use
        // the returned path directly as --after/--before.
        if (first.Name.ToLowerInvariant() == "toc" && segments.Count == 1 && first.Index.HasValue)
        {
            var body = _doc.MainDocumentPart?.Document?.Body;
            if (body != null)
            {
                var tocParas = body.Elements<Paragraph>()
                    .Where(p => p.Descendants<FieldCode>().Any(fc =>
                        fc.Text != null && fc.Text.TrimStart().StartsWith("TOC", StringComparison.OrdinalIgnoreCase)))
                    .ToList();
                var n = first.Index.Value;
                if (n >= 1 && n <= tocParas.Count)
                    return tocParas[n - 1];
            }
        }

        // Top-level /formfield[N] anchor routing. `add --type formfield`
        // returns "/formfield[N]" as the new element's identity; resolve it to
        // the body-level paragraph containing the Nth form field's begin-run
        // so callers can use the returned path directly as --after/--before.
        // R14-bug4: also accept /formfield[@name=NAME] (the schema-documented
        // stable form, signalled by first.StringIndex starting with "@name=").
        if (first.Name.ToLowerInvariant() == "formfield" && segments.Count == 1)
        {
            var allFf = FindFormFields();
            (FieldInfo Field, FormFieldData FfData) hit = default;
            if (first.Index.HasValue)
            {
                var n = first.Index.Value;
                if (n >= 1 && n <= allFf.Count) hit = allFf[n - 1];
            }
            else if (first.StringIndex != null
                && first.StringIndex.StartsWith("@name=", StringComparison.OrdinalIgnoreCase))
            {
                var target = first.StringIndex["@name=".Length..];
                hit = allFf.FirstOrDefault(ff =>
                    ff.FfData.GetFirstChild<FormFieldName>()?.Val?.Value == target);
            }
            if (hit.Field != null)
            {
                var beginRun = hit.Field.BeginRun;
                OpenXmlElement? cur = beginRun;
                while (cur != null && cur is not Paragraph) cur = cur.Parent;
                return cur ?? beginRun;
            }
        }

        // Resolve the 0-based part index for a positional header/footer segment.
        // last() must map to the LAST part by enumeration (creation) order, the
        // same as /body/p[last()] — without this it fell through to
        // (Index ?? 1) - 1 = 0 and silently resolved /header[last()] to the
        // FIRST header, so a `set /header[last()]` wrote into the wrong header
        // (content loss when several headers of different types exist).
        int PartIndex(int count) => first.StringIndex == "last()" ? count - 1 : (first.Index ?? 1) - 1;

        OpenXmlElement? current = first.Name.ToLowerInvariant() switch
        {
            "body" => _doc.MainDocumentPart?.Document?.Body,
            "styles" => _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles,
            "header" => _doc.MainDocumentPart?.HeaderParts is { } hps
                ? hps.ToList() is var hl && hl.Count > 0 ? hl.ElementAtOrDefault(PartIndex(hl.Count))?.Header : null
                : null,
            "footer" => _doc.MainDocumentPart?.FooterParts is { } fps
                ? fps.ToList() is var fl && fl.Count > 0 ? fl.ElementAtOrDefault(PartIndex(fl.Count))?.Footer : null
                : null,
            "numbering" => _doc.MainDocumentPart?.NumberingDefinitionsPart?.Numbering,
            "settings" => _doc.MainDocumentPart?.DocumentSettingsPart?.Settings,
            "comments" => _doc.MainDocumentPart?.WordprocessingCommentsPart?.Comments,
            // /footnotes and /endnotes are container aliases so that
            // /footnotes/footnote[N] and /endnotes/endnote[N] work as
            // documented in the help text. The Nth user note is also
            // selectable directly via /footnote[N] (positional) or
            // /footnote[@footnoteId=N] (id-based) — those paths bypass
            // this switch via the `current == null` block below.
            "footnotes" => _doc.MainDocumentPart?.FootnotesPart?.Footnotes,
            "endnotes" => _doc.MainDocumentPart?.EndnotesPart?.Endnotes,
            _ => null
        };

        string parentPath = "/" + first.Name + (first.Index.HasValue ? $"[{first.Index}]" : "");

        // Top-level /footnote[@footnoteId=N] / /footnote[N] routing. Mirrors
        // WordHandler.Add.cs's TryResolveFootnoteOrEndnoteBody so that paths
        // returned by `add` under a footnote/endnote are round-trippable via
        // `get` and usable as --after/--before anchors.
        if (current == null)
        {
            var fname = first.Name.ToLowerInvariant();
            if (fname == "footnote")
            {
                int? fnId = first.Index;
                if (fnId == null && first.StringIndex != null
                    && first.StringIndex.StartsWith("@footnoteId=", StringComparison.OrdinalIgnoreCase)
                    && int.TryParse(first.StringIndex["@footnoteId=".Length..], out var idv))
                {
                    fnId = idv;
                }
                if (fnId != null)
                {
                    current = _doc.MainDocumentPart?.FootnotesPart?.Footnotes?
                        .Elements<Footnote>().FirstOrDefault(f => f.Id?.Value == fnId.Value);
                    parentPath = $"/footnote[@footnoteId={fnId}]";
                }
            }
            else if (fname == "endnote")
            {
                int? enId = first.Index;
                if (enId == null && first.StringIndex != null
                    && first.StringIndex.StartsWith("@endnoteId=", StringComparison.OrdinalIgnoreCase)
                    && int.TryParse(first.StringIndex["@endnoteId=".Length..], out var idv))
                {
                    enId = idv;
                }
                if (enId != null)
                {
                    current = _doc.MainDocumentPart?.EndnotesPart?.Endnotes?
                        .Elements<Endnote>().FirstOrDefault(e => e.Id?.Value == enId.Value);
                    parentPath = $"/endnote[@endnoteId={enId}]";
                }
            }
        }

        for (int i = 1; i < segments.Count && current != null; i++)
        {
            var seg = segments[i];
            IEnumerable<OpenXmlElement> children;
            // When the current element is a block-level SDT, transparently
            // descend into its SdtContentBlock so paths like
            // /body/sdt[@sdtId=X]/p[N] resolve to paragraphs physically
            // nested inside the content wrapper. Mirrors GetBodyElements()
            // which already flattens SdtBlock when iterating body children.
            if (current is SdtBlock navSdtBlock)
            {
                var contentBlock = navSdtBlock.GetFirstChild<SdtContentBlock>();
                if (contentBlock != null) current = contentBlock;
            }
            else if (current is SdtRun navSdtRun)
            {
                var contentRun = navSdtRun.GetFirstChild<SdtContentRun>();
                if (contentRun != null) current = contentRun;
            }

            // Allow an explicit "/sdtContent" segment as a no-op selector: after
            // the transparent descend above, `current` is already the
            // SdtContent{Block,Run}. This keeps the ValidateParentChild hint
            // ("Add under <sdt>/sdtContent instead") literally navigable.
            if (seg.Name.Equals("sdtContent", StringComparison.OrdinalIgnoreCase)
                && (current is SdtContentBlock || current is SdtContentRun))
            {
                parentPath += "/sdtContent";
                continue;
            }

            // O(1) fast-path for /body/p[last()] — batch run-add resolves this
            // once per run (100k+ times in a large replay). Trusted when the
            // cached last paragraph is still parented to this body and its
            // NextSibling is the trailing sectPr, i.e. it is genuinely last
            // (NextSibling is O(1) on the SDK's singly-linked list, unlike a
            // full ToList()+LastOrDefault scan). Emit the same "/p[N]" segment
            // the scan below would, so resolvedPath is identical. A stale hit
            // falls through to the scan.
            if (seg.StringIndex == "last()" && seg.Name == "p" && current is Body
                && _bodyParaCount >= 0
                && _lastBodyParagraph is Paragraph lbpFast
                && ReferenceEquals(lbpFast.Parent, current)
                && lbpFast.NextSibling() is SectionProperties)
            {
                parentPath += $"/p[{_bodyParaCount}]";
                current = lbpFast;
                continue;
            }

            if (current is Body body2 && (seg.Name.ToLowerInvariant() == "p" || seg.Name.ToLowerInvariant() == "tbl"))
            {
                // Only count direct body-level paragraphs/tables, skip those inside SdtBlock containers.
                // #6: paragraphs whose sole content is m:oMathPara are
                // counted via the /body/oMathPara[N] path instead, so the
                // /body/p[N] enumeration skips them to match HTML-preview
                // data-path attribution (which also skips them).
                // BUG-DUMP8-01/02: w:customXml body wrappers are non-structural —
                // recursively flatten so paragraphs/tables nested inside one
                // (or several) levels of CustomXmlBlock surface in the same
                // /body/p[N] / /body/tbl[N] enumeration. Mirrors the listing
                // logic in WalkBodyChild for `get /body`; without this, path
                // resolution diverged from listing and `get /body/p[1]` threw
                // "Path not found" on customXml-wrapped paragraphs.
                // PERF: cache the filtered lists per Body instance + child count.
                // Without the cache, dumping a doc with N body paragraphs costs
                // O(N²) because the dump emitter calls Get("/body/p[K]") for
                // every K in 1..N, and each call re-walked body.ChildElements.
                // Real-world 14k-paragraph doc: 5+ minutes → seconds.
                children = seg.Name.ToLowerInvariant() == "p"
                    ? GetBodyParagraphIndex(body2)
                    : GetBodyTableIndex(body2);
            }
            else if (current is Body body3 && seg.Name == "oMathPara")
            {
                // oMathPara can be direct body children or wrapped inside w:p elements
                var mathParas = new List<OpenXmlElement>();
                foreach (var el in body3.ChildElements)
                {
                    if (el.LocalName == "oMathPara" || el is M.Paragraph)
                        mathParas.Add(el);
                    else if (el is Paragraph wp && IsOMathParaWrapperParagraph(wp))
                    {
                        // Only pure-wrapper paragraphs (pPr + single oMathPara child)
                        // — otherwise /body/p[N] and /body/oMathPara[M] would both
                        // address the same paragraph (mixed prose + inline math),
                        // causing Get/Set/Remove to diverge by callsite.
                        var inner = wp.ChildElements.FirstOrDefault(c => c.LocalName == "oMathPara" || c is M.Paragraph);
                        if (inner != null) mathParas.Add(inner);
                    }
                }
                children = mathParas;
            }
            else
            {
                children = seg.Name.ToLowerInvariant() switch
                {
                    "p" => current.Elements<Paragraph>().Cast<OpenXmlElement>(),
                    // BUG-D1-TXBX-RUN-INDEX: /<para>/r[N] must enumerate the
                    // SAME runs NodeBuilder surfaces as paragraph children —
                    // i.e. skip runs that live INSIDE a child textbox
                    // (their canonical path is /<para>/textbox[N]/p[M]/r[K]).
                    // Without this filter, a paragraph hosting 3 textbox
                    // drawings (Card A / B / C side-by-side) resolves r[2] to
                    // a text run INSIDE Card A's textbox content instead of
                    // Card B's drawing run, so WordBatchEmitter probes the
                    // wrong XML, IsTextboxDrawing returns false, and Cards
                    // B/C silently degrade to plain `add r` rows on dump
                    // round-trip — losing the textboxes entirely and
                    // misaligning every subsequent /body/textbox[K] index.
                    // Mirrors GetAllRuns in WordHandler.Helpers (also skips
                    // SimpleField/SdtRun-nested runs for the same path-
                    // stability reason).
                    // PERF(nav-child-cache): the filter body lives in
                    // GetNavRunIndex, which memoizes the resulting List per
                    // container so resolving r[K] for K=1..R is O(R) total, not
                    // O(R²). Returning the cached List (not a Cast wrapper) lets
                    // the `as List` reuse at the navigation tail skip a re-copy.
                    "r" => GetNavRunIndex(current),
                    "tbl" => current.Elements<Table>().Cast<OpenXmlElement>(),
                    "tr" => current is Table trHostTable
                        ? GetNavRowIndex(trHostTable)
                        : current.Elements<TableRow>().Cast<OpenXmlElement>(),
                    "tc" => current is TableRow tcHostRow
                        ? GetNavCellIndex(tcHostRow)
                        : current.Elements<TableCell>().Cast<OpenXmlElement>(),
                    "sdt" => current.ChildElements
                        .Where(e => e is SdtBlock || e is SdtRun).Cast<OpenXmlElement>(),
                    // v5.7-cont: /body/textbox[N] → walk descendant drawings,
                    // pick the Nth wps:txbx host, return its w:txbxContent
                    // so child p[M] resolves naturally via the next loop iter.
                    "textbox" => current.Descendants<Drawing>()
                        .Where(d => d.InnerXml.Contains("<wps:txbx") || d.InnerXml.Contains("txBox=\"1\""))
                        .Select(d => (OpenXmlElement?)d.Descendants().FirstOrDefault(e =>
                            e.LocalName == "txbxContent"
                            && e.NamespaceUri == "http://schemas.openxmlformats.org/wordprocessingml/2006/main"))
                        .Where(e => e != null)
                        .Cast<OpenXmlElement>(),
                    // v5.7-cont: /body/shape[N] → walk descendant drawings,
                    // pick the Nth wps:wsp that isn't itself a textbox. Returns
                    // the wps:wsp element so children resolve in its scope.
                    "shape" => current.Descendants<Drawing>()
                        .Where(d => d.InnerXml.Contains("<wps:wsp")
                                 && !d.InnerXml.Contains("<wps:txbx")
                                 && !d.InnerXml.Contains("txBox=\"1\""))
                        .Select(d => (OpenXmlElement?)d.Descendants().FirstOrDefault(e =>
                            e.LocalName == "wsp"
                            && e.NamespaceUri == "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"))
                        .Where(e => e != null)
                        .Cast<OpenXmlElement>(),
                    // /body/group[N] → the Nth wpg:wgp group drawing (e.g. a
                    // `diagram` is emitted as one group so it stays adjustable as
                    // a unit). Returns the <wpg:wgp> element; `set width/height`
                    // scales the whole group (mirrors the pptx /slide[N]/group[K]).
                    "group" => current.Descendants<Drawing>()
                        .Select(d => (OpenXmlElement?)d.Descendants().FirstOrDefault(e =>
                            e.LocalName == "wgp"
                            && e.NamespaceUri == "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"))
                        .Where(e => e != null)
                        .Cast<OpenXmlElement>(),
                    // /<para>/tab[N] and /styles/<id>/tab[N] descend
                    // transparently through pPr/tabs (or StyleParagraph-
                    // Properties/tabs) so the user-facing path stays flat
                    // instead of leaking the OOXML containers (.../pPr/tabs/tab).
                    // Symmetric with how AddTab returns the flat form.
                    "tab" when current is Paragraph navParaT
                        => navParaT.ParagraphProperties?.GetFirstChild<Tabs>()?.Elements<TabStop>().Cast<OpenXmlElement>()
                           ?? Enumerable.Empty<OpenXmlElement>(),
                    "tab" when current is Style navStyleT
                        => navStyleT.StyleParagraphProperties?.GetFirstChild<Tabs>()?.Elements<TabStop>().Cast<OpenXmlElement>()
                           ?? Enumerable.Empty<OpenXmlElement>(),
                    // /styles/<key> resolves <key> as a styleId or styleName
                    // (matches Set.Dispatch.cs's regex+OR matching), so paths
                    // like /styles/Heading1 are navigable for Add/Get/Set.
                    // The segment name here IS the key, not an OOXML local-
                    // name; downstream FirstOrDefault picks the (single) match.
                    _ when current is Styles navStylesContainer
                        => navStylesContainer.Elements<Style>().Where(s =>
                            string.Equals(s.StyleId?.Value, seg.Name, StringComparison.Ordinal)
                            || string.Equals(s.StyleName?.Val?.Value, seg.Name, StringComparison.Ordinal))
                           .Cast<OpenXmlElement>(),
                    // CONSISTENCY(footnotes-container): /footnotes/footnote[N]
                    // enumerates user footnotes only (id > 0), matching what
                    // `query footnote` returns and the positional /footnote[N]
                    // routing used by Add. The schema's separator/continuation
                    // entries (id=-1, id=0) are excluded so positional indexes
                    // line up across paths.
                    "footnote" when current is Footnotes fns
                        => fns.Elements<Footnote>().Where(f => f.Id?.Value > 0).Cast<OpenXmlElement>(),
                    "endnote" when current is Endnotes ens
                        => ens.Elements<Endnote>().Where(e => e.Id?.Value > 0).Cast<OpenXmlElement>(),
                    // BUG-DUMP-FLDSIMPLE-IMG: address the Nth drawing-bearing run INSIDE
                    // a paragraph's <w:fldSimple> results so the path-based picture
                    // pipeline (GetElementXml / GetImageBinary / GetDrawingShapeEmitData)
                    // can ship an image cached in a simple field. (Case label is lower —
                    // this switch is on seg.Name.ToLowerInvariant().) Order matches the
                    // fldSimple decomposition in the paragraph emit.
                    "fldimgrun" when current is Paragraph fldImgPara
                        => fldImgPara.Elements<SimpleField>()
                            .SelectMany(f => f.Elements<Run>()
                                .Where(r => r.GetFirstChild<Drawing>() != null))
                            .Cast<OpenXmlElement>(),
                    _ => current.ChildElements.Where(e => e.LocalName == seg.Name).Cast<OpenXmlElement>()
                };
            }

            // Reuse the list when the switch already produced one (e.g.
            // GetBodyParagraphIndex/GetBodyTableIndex return a cached
            // List<OpenXmlElement>). ToList() would copy it O(n) on EVERY
            // navigation, so resolving /body/p[N] for N=1..n — batch set/get of
            // paragraphs — was O(n²). The list is only read here (ElementAt /
            // LastOrDefault / IndexOf), so sharing the cached reference is safe.
            var childList = children as List<OpenXmlElement> ?? children.ToList();
            OpenXmlElement? next;
            if (seg.Index.HasValue)
                next = childList.ElementAtOrDefault(seg.Index.Value - 1);
            else if (seg.StringIndex == "last()")
                next = childList.LastOrDefault();
            else if (seg.StringIndex != null && seg.StringIndex.StartsWith("@paraId=", StringComparison.OrdinalIgnoreCase))
            {
                var targetId = seg.StringIndex["@paraId=".Length..];
                // CONSISTENCY(paraid-global-uniqueness): paraId is globally
                // unique across body/headers/footers/footnotes/endnotes/
                // comments (EnsureAllParaIds scans every part). Resolve by
                // descendants too — direct-child-only scan made cell paras
                // unreachable from the canonical /body/p[@paraId=...] form
                // that AddPtab/AddBreak/AddField return for cell parents.
                // Body case uses the O(1) paraId map (dump resolves one
                // /body/p[@paraId] per paragraph — the scan was O(n²)).
                next = current is Body navParaByIdBody
                    ? GetBodyParaById(navParaByIdBody, targetId)
                    : childList.OfType<Paragraph>()
                        .FirstOrDefault(p => string.Equals(p.ParagraphId?.Value, targetId, StringComparison.OrdinalIgnoreCase));
                if (next == null)
                {
                    next = (current as OpenXmlElement)?.Descendants<Paragraph>()
                        .FirstOrDefault(p => string.Equals(p.ParagraphId?.Value, targetId, StringComparison.OrdinalIgnoreCase));
                }
            }
            else if (seg.StringIndex != null && seg.StringIndex.StartsWith("@textId=", StringComparison.OrdinalIgnoreCase))
            {
                var targetId = seg.StringIndex["@textId=".Length..];
                next = childList.OfType<Paragraph>()
                    .FirstOrDefault(p => string.Equals(p.TextId?.Value, targetId, StringComparison.OrdinalIgnoreCase));
                if (next == null)
                {
                    next = (current as OpenXmlElement)?.Descendants<Paragraph>()
                        .FirstOrDefault(p => string.Equals(p.TextId?.Value, targetId, StringComparison.OrdinalIgnoreCase));
                }
            }
            else if (seg.StringIndex != null && seg.StringIndex.StartsWith("@commentId=", StringComparison.OrdinalIgnoreCase))
            {
                var targetId = seg.StringIndex["@commentId=".Length..];
                next = childList.OfType<Comment>()
                    .FirstOrDefault(c => c.Id?.Value == targetId);
            }
            else if (seg.StringIndex != null && seg.StringIndex.StartsWith("@name=", StringComparison.OrdinalIgnoreCase))
            {
                // Generic @name=... selector, used by bookmarkStart[@name=X]
                // so that the path returned by AddBookmark is navigable.
                var targetName = seg.StringIndex["@name=".Length..];
                next = childList.FirstOrDefault(e =>
                    e is BookmarkStart bs && string.Equals(bs.Name?.Value, targetName, StringComparison.Ordinal));
            }
            else if (seg.StringIndex != null && seg.StringIndex.StartsWith("@sdtId=", StringComparison.OrdinalIgnoreCase))
            {
                var targetId = seg.StringIndex["@sdtId=".Length..];
                next = childList.Where(e => e is SdtBlock or SdtRun)
                    .FirstOrDefault(e =>
                    {
                        var sdtId = (e is SdtBlock sb ? sb.SdtProperties : (e as SdtRun)?.SdtProperties)
                            ?.GetFirstChild<SdtId>()?.Val?.Value;
                        return sdtId?.ToString() == targetId;
                    });
            }
            // CONSISTENCY(id-selectors): mirror @paraId/@commentId/@sdtId — accept @id= for
            // numbering/abstractNum (w:abstractNumId@val) and numbering/num (w:num@numId).
            else if (seg.StringIndex != null && seg.StringIndex.StartsWith("@id=", StringComparison.OrdinalIgnoreCase))
            {
                var targetId = seg.StringIndex["@id=".Length..];
                next = childList.FirstOrDefault(e => e switch
                {
                    AbstractNum an => an.AbstractNumberId?.Value.ToString() == targetId,
                    NumberingInstance ni => ni.NumberID?.Value.ToString() == targetId,
                    // BUG-DUMP-BMEND-IDPATH: a bookmarkEnd/bookmarkStart addressed by
                    // w:id must resolve by id, NOT positionally. The dump emits a
                    // bookmarkEnd's path as bookmarkEnd[@id=N]; without this, the
                    // numeric form bookmarkEnd[N] was treated as a positional ordinal,
                    // so an end whose id != its document-order position (e.g. the
                    // first bookmarkEnd in a stack whose id is 4 but the 4th end has
                    // id 7) re-read the WRONG element, producing a duplicate bookmark
                    // id on round-trip.
                    BookmarkEnd be => be.Id?.Value?.ToString() == targetId,
                    BookmarkStart bs => bs.Id?.Value?.ToString() == targetId,
                    _ => false,
                });
            }
            else if (seg.StringIndex != null && seg.StringIndex.StartsWith("@", StringComparison.Ordinal))
            {
                // Unrecognized attribute predicate — throw rather than silently returning
                // the first element. ValidateAndNormalizePredicate accepts any @ident=value
                // syntactically, but not every attribute maps to a Word OOXML concept.
                // Comment on the gap: expand the dispatch chain above when a new attribute
                // needs to be addressable (e.g. @bookmarkId=, @w14:paraId=).
                var eq = seg.StringIndex.IndexOf('=');
                var attrName = eq > 0 ? seg.StringIndex[1..eq] : seg.StringIndex[1..];
                throw new ArgumentException(
                    $"Attribute predicate '@{attrName}' is not a recognized Word path attribute. " +
                    $"Supported attributes: @paraId, @textId, @commentId, @sdtId, @id, @name.");
            }
            else
                next = childList.FirstOrDefault();

            if (next == null)
            {
                availableContext = BuildAvailableContext(current, parentPath, seg.Name, childList.Count);
                return null;
            }

            // Build path segment: prefer stable ID when available, fallback to positional.
            // Use the resolved element's LocalName (always canonical lowercase for OOXML)
            // rather than seg.Name (which echoes user capitalization like 'P'), so the
            // returned path round-trips cleanly and matches Query's canonical form.
            // Style is exempt — /styles/<id> uses the user-supplied styleId/Name as the key.
            var canonName = (next is Style) ? seg.Name : next.LocalName;
            // BUG-D1-TXBX-PATH: the textbox / shape segments resolve to
            // wps:wsp + w:txbxContent OOXML elements whose LocalName is
            // "wsp" / "txbxContent" — neither is navigable as a path
            // segment (the navigable name is the user-facing "textbox" /
            // "shape" form handled at the children-selector switch above).
            // Preserve the user-supplied segment name so resolvedPath stays
            // re-navigable; without this, descendant Get calls on children
            // like /<host>/textbox[N]/tbl[K]/tr[J] fail with
            // "No txbxContent found at /body".
            // Keep seg.Name verbatim (no lowercasing): path matching for the
            // literal "txbxContent" is case-sensitive, so lowercasing broke
            // round-trip when the user supplied the literal form (issue #258).
            if (canonName == "txbxContent" || canonName == "wsp")
                canonName = seg.Name;
            if (next is Paragraph navPara && !string.IsNullOrEmpty(navPara.ParagraphId?.Value))
            {
                parentPath += "/" + canonName + $"[@paraId={navPara.ParagraphId.Value}]";
            }
            else if (next is Comment navComment && navComment.Id?.Value != null)
            {
                parentPath += "/" + canonName + $"[@commentId={navComment.Id.Value}]";
            }
            else if (next is Style navStyle)
            {
                // Style is keyed by styleId — emit /styles/<id> without a
                // positional [N] suffix to match Query's canonical form.
                parentPath += "/" + (navStyle.StyleId?.Value ?? seg.Name);
            }
            else if (next is SdtBlock or SdtRun)
            {
                var sdtProps = (next is SdtBlock sb2 ? sb2.SdtProperties : (next as SdtRun)?.SdtProperties);
                var sdtIdVal = sdtProps?.GetFirstChild<SdtId>()?.Val?.Value;
                if (sdtIdVal != null)
                    parentPath += "/" + canonName + $"[@sdtId={sdtIdVal}]";
                else
                {
                    var posIdx = PathIndex.FromArrayIndex(childList.IndexOf(next));
                    parentPath += "/" + canonName + $"[{posIdx}]";
                }
            }
            else
            {
                var posIdx = PathIndex.FromArrayIndex(childList.IndexOf(next));
                parentPath += "/" + canonName + $"[{posIdx}]";
            }
            current = next;
        }

        resolvedPath = parentPath;
        return current;
    }

    /// <summary>
    /// Build a context string describing available children when navigation fails.
    /// </summary>
    private static string BuildAvailableContext(OpenXmlElement parent, string parentPath, string requestedType, int matchCount)
    {
        if (matchCount > 0)
            return $"Available at {parentPath}: {requestedType}[1]..{requestedType}[{matchCount}]";

        // List distinct child types at this level. R2-bt-1: a display equation
        // is a w:p wrapping a single m:oMathPara. The path resolver (rule #6 in
        // ResolvePath / WalkBodyChild) does NOT count such wrapper paragraphs
        // under /body/p[N] — they are addressed as /body/oMathPara[N] instead.
        // The raw LocalName grouping below counted the wrapper w:p as p(1),
        // producing the self-contradictory "No p found … Available children:
        // p(1)" message: the listing claimed a p the resolver refuses to index.
        // Reclassify wrapper paragraphs as oMathPara here so the listing agrees
        // with what /body/p[N] and /body/oMathPara[N] actually resolve.
        bool atBody = parent is Body;
        var childTypes = parent.ChildElements
            .GroupBy(c => atBody && c is Paragraph wp && IsOMathParaWrapperParagraph(wp)
                ? "oMathPara"
                : c.LocalName)
            .Select(g => $"{g.Key}({g.Count()})")
            .Take(10)
            .ToList();

        return childTypes.Count > 0
            ? $"No {requestedType} found at {parentPath}. Available children: {string.Join(", ", childTypes)}"
            : $"No children at {parentPath}";
    }

    /// <summary>
    /// R4-bt-1: compute the canonical, RESOLVABLE path of a math element
    /// (m:oMath inline, or m:oMathPara display) that lives in a body-level
    /// paragraph — used after an equation mode switch MOVES the element, so the
    /// Set response can report the new path instead of the stale pre-move one.
    /// Mirrors the body enumeration in ResolvePath exactly: a pure
    /// oMathPara-wrapper paragraph is addressed as /body/oMathPara[N]; any other
    /// paragraph is /body/p[N], and the math element hangs off it as
    /// /oMath[K] or /oMathPara[K]. Returns null if the element is not under a
    /// direct body paragraph (caller then keeps the original path).
    /// </summary>
    internal string? ComputeMathElementPath(OpenXmlElement mathEl)
    {
        var body = _doc.MainDocumentPart?.Document?.Body;
        if (body == null) return null;
        var hostPara = mathEl.Ancestors<Paragraph>().FirstOrDefault();
        if (hostPara == null) return null;

        bool isDisplay = mathEl is M.Paragraph || mathEl.LocalName == "oMathPara";

        // Shared tail: given the resolved paragraph-path prefix, append the
        // immediate container (/hyperlink[H] when the math sits in a hyperlink)
        // and the math index (/oMathPara[K] display, /oMath[K] inline).
        string? MathTail(string paraPrefix)
        {
            var container = mathEl.Parent!;
            var prefix = paraPrefix;
            if (container is Hyperlink hlC)
            {
                int hIdx = hostPara.Elements<Hyperlink>().ToList()
                    .FindIndex(h => ReferenceEquals(h, hlC)) + 1;
                if (hIdx <= 0) return null;
                prefix += $"/hyperlink[{hIdx}]";
            }
            else if (!ReferenceEquals(container, hostPara)) return null;
            if (isDisplay)
            {
                int k = container.Elements<M.Paragraph>().ToList()
                    .FindIndex(e => ReferenceEquals(e, mathEl)) + 1;
                return k > 0 ? $"{prefix}/oMathPara[{k}]" : null;
            }
            int kk = container.Elements<M.OfficeMath>().ToList()
                .FindIndex(e => ReferenceEquals(e, mathEl)) + 1;
            return kk > 0 ? $"{prefix}/oMath[{kk}]" : null;
        }

        // Footnote/endnote-hosted equations: build the note-scoped path
        // (/footnotes/footnote[N]/p[@paraId=X]/…) so a mode switch inside a note
        // reports a resolvable path instead of the stale pre-switch one. N
        // enumerates user notes (Id>0), matching the resolver and Add.
        if (hostPara.Parent is Footnote or Endnote)
        {
            var noteEl = hostPara.Parent!;
            bool isFn = noteEl is Footnote;
            var siblings = (isFn
                ? (noteEl.Parent as Footnotes)?.Elements<Footnote>()
                    .Where(f => f.Id?.Value > 0).Cast<OpenXmlElement>()
                : (noteEl.Parent as Endnotes)?.Elements<Endnote>()
                    .Where(e => e.Id?.Value > 0).Cast<OpenXmlElement>())?.ToList();
            if (siblings == null) return null;
            int nIdx = siblings.FindIndex(n => ReferenceEquals(n, noteEl)) + 1;
            if (nIdx <= 0) return null;
            int pIdx0 = noteEl.Elements<Paragraph>().ToList()
                .FindIndex(p => ReferenceEquals(p, hostPara)) + 1;
            var seg = BuildParaPathSegment(hostPara, pIdx0);
            var notePrefix = isFn
                ? $"/footnotes/footnote[{nIdx}]"
                : $"/endnotes/endnote[{nIdx}]";
            return MathTail($"{notePrefix}/{seg}");
        }

        if (!ReferenceEquals(hostPara.Parent, body)) return null;

        // Pure oMathPara-wrapper paragraphs are addressed at body level as
        // /body/oMathPara[N]; the wrapped m:oMathPara IS the target itself.
        if (IsOMathParaWrapperParagraph(hostPara))
        {
            int n = 0;
            foreach (var el in body.ChildElements)
            {
                if (el.LocalName == "oMathPara" || el is M.Paragraph) n++;
                else if (el is Paragraph wp && IsOMathParaWrapperParagraph(wp))
                {
                    n++;
                    if (ReferenceEquals(wp, hostPara)) return $"/body/oMathPara[{n}]";
                }
            }
            return null;
        }

        // Otherwise the host is a regular /body/p[N] (skipping pure-wrapper
        // paragraphs, matching GetBodyParagraphIndex), and the math element is a
        // positional child of it.
        int pIdx = 0;
        foreach (var el in body.Elements<Paragraph>())
        {
            if (IsOMathParaWrapperParagraph(el)) continue; // counted under oMathPara[N]
            pIdx++;
            // The math element sits directly in the paragraph OR nested inside a
            // w:hyperlink child (dump→batch round-trips equations in a hyperlink);
            // MathTail handles both and returns the resolvable /body/p[N]/… path.
            if (ReferenceEquals(el, hostPara))
                return MathTail($"/body/p[{pIdx}]");
        }
        return null;
    }

    private DocumentNode BookmarkStartToNode(BookmarkStart bkStart, DocumentNode node)
    {
        node.Type = "bookmark";
        node.Format["name"] = bkStart.Name?.Value ?? "";
        node.Format["id"] = bkStart.Id?.Value ?? "";
        // BUG-DUMP-BMSPAN: flag a content-wrapping bookmark so the emitter
        // splits it into a positioned start (`open=true`) + a separate
        // `end=true` op after the wrapped runs, preserving the range. Holds
        // for the BookmarkEnd both inside the same paragraph and as a body-
        // direct sibling (POI/Word emit the End after </w:p>).
        if (IsContentSpanBookmark(bkStart))
            node.Format["_spanOpen"] = true;
        // BUG-DUMP10-04: for cross-paragraph bookmark spans, walk
        // forward over sibling paragraphs in the same body and
        // surface the BookmarkEnd's paragraph offset (0-based).
        // 0 = same paragraph (default; AddBookmark places End next to
        // Start). >0 = the End sits N paragraphs after the Start.
        // Without this, dump emitted only the BookmarkStart and
        // AddBookmark always re-emitted the End in the same paragraph,
        // collapsing every multi-paragraph bookmark on round-trip.
        var bkStartId = bkStart.Id?.Value;
        if (!string.IsNullOrEmpty(bkStartId)
            && bkStart.Ancestors<Paragraph>().FirstOrDefault() is { } startPara
            && startPara.Parent is OpenXmlElement bodyParent)
        {
            var siblings = bodyParent.Elements<Paragraph>().ToList();
            int startIdx = siblings.IndexOf(startPara);
            if (startIdx >= 0)
            {
                for (int i = startIdx; i < siblings.Count; i++)
                {
                    var endHere = siblings[i].Descendants<BookmarkEnd>()
                        .FirstOrDefault(be => be.Id?.Value == bkStartId);
                    if (endHere != null)
                    {
                        int offset = i - startIdx;
                        if (offset > 0) node.Format["endPara"] = offset;
                        break;
                    }
                }
            }
        }
        // BUG-DUMP-R32-4: a table-column-range bookmark carries
        // w:colFirst/w:colLast (a rectangular column-span bookmark over table
        // columns). These were dropped on dump, downgrading the bookmark to a
        // plain point bookmark. Surface them so AddBookmark re-stamps the attrs.
        // Mirrors the permStart ColumnFirst/ColumnLast reads above.
        if (bkStart.ColumnFirst?.Value != null)
            node.Format["colFirst"] = bkStart.ColumnFirst.Value.ToString();
        if (bkStart.ColumnLast?.Value != null)
            node.Format["colLast"] = bkStart.ColumnLast.Value.ToString();
        // BUG-DUMP-BMDISPLACED: a bookmark adjacent to a custom-XML / SDT
        // boundary (e.g. a TOC heading bookmark sitting just before the TOC's
        // <w:sdt>) carries w:displacedByCustomXml ("next"/"prev") — it tells
        // Word which side of the structured-tag the marker resolves to. Dropped
        // on dump, the bookmark's position shifted across the SDT boundary and
        // every PAGEREF/TOC entry referencing it rendered "Error! Bookmark not
        // defined." Surface it so AddBookmark re-stamps the attribute.
        if (bkStart.DisplacedByCustomXml is { InnerText: { Length: > 0 } dbcx })
            node.Format["displacedByCustomXml"] = dbcx;
        var bkText = GetBookmarkText(bkStart);
        if (!string.IsNullOrEmpty(bkText))
            node.Text = bkText;
        return node;
    }

    // BUG-DUMP-PERM: surface a ranged editing-permission start marker
    // (<w:permStart>) with all its attributes so dump→batch round-trips the
    // editable-region delimiter. Mirrors BookmarkStartToNode.
    private static DocumentNode PermStartToNode(PermStart permStart, DocumentNode node)
    {
        node.Type = "permStart";
        if (permStart.Id?.Value != null) node.Format["id"] = permStart.Id.Value.ToString();
        if (permStart.EditorGroup?.Value != null && permStart.EditorGroup.HasValue)
            node.Format["edGrp"] = permStart.EditorGroup.InnerText;
        if (permStart.Ed?.Value is { Length: > 0 } edUser) node.Format["ed"] = edUser;
        if (permStart.ColumnFirst?.Value != null) node.Format["colFirst"] = permStart.ColumnFirst.Value.ToString();
        if (permStart.ColumnLast?.Value != null) node.Format["colLast"] = permStart.ColumnLast.Value.ToString();
        return node;
    }

    private DocumentNode FootnoteToNode(Footnote fnEl, DocumentNode node, string path, int depth)
    {
        node.Type = "footnote";
        // Strip the reference-mark leading space (CONSISTENCY with Query
        // get-by-id and `query footnote`). Without this branch the
        // generic InnerText fallback below would return " fn-text".
        node.Text = GetFootnoteText(fnEl);
        if (fnEl.Id?.Value != null) node.Format["id"] = fnEl.Id.Value;
        if (fnEl.Type?.Value != null) node.Format["type"] = fnEl.Type.InnerText;
        // R44 minor-5: surface first-run formatting on footnote node so
        // bold/italic/size/color set via Add/Set round-trip through Get.
        // Mirrors the hyperlink firstRun pattern at line ~2746 above.
        var fnFirstRun = fnEl.Descendants<Run>().FirstOrDefault(r => r.GetFirstChild<Text>() != null);
        if (fnFirstRun?.RunProperties != null)
        {
            var rp = fnFirstRun.RunProperties;
            if (rp.RunFonts?.Ascii?.Value != null) node.Format["font"] = rp.RunFonts.Ascii.Value;
            if (rp.FontSize?.Val?.Value != null)
                node.Format["size"] = $"{int.Parse(rp.FontSize.Val.Value) / 2.0:0.##}pt";
            if (rp.Bold != null) node.Format["bold"] = IsToggleOn(rp.Bold);
            if (rp.Italic != null) node.Format["italic"] = IsToggleOn(rp.Italic);
            if (rp.Color?.ThemeColor?.HasValue == true) node.Format["color"] = rp.Color.ThemeColor.InnerText;
            else if (rp.Color?.Val?.Value != null) node.Format["color"] = ParseHelpers.FormatHexColor(rp.Color.Val.Value);
            if (rp.Underline?.Val != null) node.Format["underline"] = rp.Underline.Val.InnerText;
            if (rp.Strike != null) node.Format["strike"] = IsToggleOn(rp.Strike);
            if (rp.Highlight?.Val != null) node.Format["highlight"] = rp.Highlight.Val.InnerText;
        }
        // R20-wbt-1: surface direction from the first content paragraph's
        // pPr.BiDi so the cascade (already applied by ApplyFootnoteEndnoteFormatKeys)
        // round-trips through Get. Mirrors the paragraph readback below.
        var fnBidi = fnEl.Descendants<Paragraph>().FirstOrDefault()?.ParagraphProperties?.GetFirstChild<BiDi>();
        if (fnBidi != null)
            node.Format["direction"] = TryReadOnOff(fnBidi.Val) == true ? "rtl" : "ltr";
        // BUG-DUMP8-05/06: Paragraph branch surfaces inline w:sym (as
        // sym= run children) and m:oMath (as equation children) but the
        // Footnote branch returned early after flat text/format, so
        // sym and oMath inside footnote bodies were silently dropped.
        // Walk descendant runs/equations and surface them as children
        // on the footnote node, mirroring the paragraph walker's keys.
        if (depth > 0)
        {
            int fnSymIdx = 0;
            foreach (var symRun in fnEl.Descendants<Run>())
            {
                var symEl = symRun.GetFirstChild<SymbolChar>();
                if (symEl?.Char?.Value == null) continue;
                var symFontVal = symEl.Font?.Value ?? "";
                var symNode = new DocumentNode
                {
                    Type = "run",
                    Path = $"{path}/r[{fnSymIdx + 1}]",
                };
                symNode.Format["sym"] = $"{symFontVal}:{symEl.Char.Value}";
                node.Children.Add(symNode);
                fnSymIdx++;
            }
            AddNoteEquationChildren(fnEl, node, path, depth);
        }
        return node;
    }

    // Emit RESOLVABLE child paths for equations inside a footnote/endnote. The
    // former flat `Descendants<M.OfficeMath>()` + `equation[N]` labelling produced
    // paths (/footnotes/footnote[N]/equation[K]) that no element matches — the
    // resolver keys off LocalName. Walk per host paragraph and emit the same
    // resolvable segments AddEquation returns: p[@paraId=X]/oMath[K] (inline) and
    // p[@paraId=X]/oMathPara[K] (display, wrapper w:p added by the R5 fix),
    // including hyperlink-nested math.
    private void AddNoteEquationChildren(OpenXmlElement noteEl, DocumentNode node, string path, int depth)
    {
        int pIdx = 0;
        foreach (var p in noteEl.Elements<Paragraph>())
        {
            pIdx++;
            var seg = BuildParaPathSegment(p, pIdx);
            void EmitFrom(OpenXmlElement container, string cSeg)
            {
                int di = 0;
                foreach (var disp in container.Elements<M.Paragraph>())
                    node.Children.Add(ElementToNode(disp, $"{cSeg}/oMathPara[{++di}]", depth - 1));
                int ii = 0;
                foreach (var inl in container.Elements<M.OfficeMath>())
                    node.Children.Add(ElementToNode(inl, $"{cSeg}/oMath[{++ii}]", depth - 1));
            }
            EmitFrom(p, $"{path}/{seg}");
            var hyperlinks = p.Elements<Hyperlink>().ToList();
            for (int h = 0; h < hyperlinks.Count; h++)
                EmitFrom(hyperlinks[h], $"{path}/{seg}/hyperlink[{h + 1}]");
        }
    }

    private DocumentNode EndnoteToNode(Endnote enEl, DocumentNode node, string path, int depth)
    {
        node.Type = "endnote";
        node.Text = GetFootnoteText(enEl);
        if (enEl.Id?.Value != null) node.Format["id"] = enEl.Id.Value;
        if (enEl.Type?.Value != null) node.Format["type"] = enEl.Type.InnerText;
        // R44 minor-5: mirror footnote firstRun readback for endnote.
        var enFirstRun = enEl.Descendants<Run>().FirstOrDefault(r => r.GetFirstChild<Text>() != null);
        if (enFirstRun?.RunProperties != null)
        {
            var rp = enFirstRun.RunProperties;
            if (rp.RunFonts?.Ascii?.Value != null) node.Format["font"] = rp.RunFonts.Ascii.Value;
            if (rp.FontSize?.Val?.Value != null)
                node.Format["size"] = $"{int.Parse(rp.FontSize.Val.Value) / 2.0:0.##}pt";
            if (rp.Bold != null) node.Format["bold"] = IsToggleOn(rp.Bold);
            if (rp.Italic != null) node.Format["italic"] = IsToggleOn(rp.Italic);
            if (rp.Color?.ThemeColor?.HasValue == true) node.Format["color"] = rp.Color.ThemeColor.InnerText;
            else if (rp.Color?.Val?.Value != null) node.Format["color"] = ParseHelpers.FormatHexColor(rp.Color.Val.Value);
            if (rp.Underline?.Val != null) node.Format["underline"] = rp.Underline.Val.InnerText;
            if (rp.Strike != null) node.Format["strike"] = IsToggleOn(rp.Strike);
            if (rp.Highlight?.Val != null) node.Format["highlight"] = rp.Highlight.Val.InnerText;
        }
        var enBidi = enEl.Descendants<Paragraph>().FirstOrDefault()?.ParagraphProperties?.GetFirstChild<BiDi>();
        if (enBidi != null)
            node.Format["direction"] = TryReadOnOff(enBidi.Val) == true ? "rtl" : "ltr";
        // CONSISTENCY with Footnote: surface inline w:sym / m:oMath
        // descendants so dump round-trips them through batch.
        if (depth > 0)
        {
            int enSymIdx = 0;
            foreach (var symRun in enEl.Descendants<Run>())
            {
                var symEl = symRun.GetFirstChild<SymbolChar>();
                if (symEl?.Char?.Value == null) continue;
                var symFontVal = symEl.Font?.Value ?? "";
                var symNode = new DocumentNode
                {
                    Type = "run",
                    Path = $"{path}/r[{enSymIdx + 1}]",
                };
                symNode.Format["sym"] = $"{symFontVal}:{symEl.Char.Value}";
                node.Children.Add(symNode);
                enSymIdx++;
            }
            AddNoteEquationChildren(enEl, node, path, depth);
        }
        return node;
    }

    private DocumentNode CommentToNode(Comment comment, DocumentNode node)
    {
        node.Type = "comment";
        node.Text = string.Join("", comment.Descendants<Text>().Select(t => t.Text));
        if (comment.Author?.Value != null) node.Format["author"] = comment.Author.Value;
        if (comment.Initials?.Value != null) node.Format["initials"] = comment.Initials.Value;
        if (comment.Id?.Value != null) node.Format["id"] = comment.Id.Value;
        if (comment.Date?.Value != null) node.Format["date"] = comment.Date.Value.ToString("o");
        if (comment.Id?.Value != null)
        {
            var anchorPath = FindCommentAnchorPath(comment.Id.Value);
            if (anchorPath != null) node.Format["anchoredTo"] = anchorPath;
        }
        // commentsExtended.xml (w15): resolved-state + reply-parent. `done` is
        // emitted on every comment (enables `query 'comment[done=false]'`);
        // `parentId` only on replies (the parent comment's w:id).
        var (cmtParentId, cmtDone) = ReadCommentExInfo(comment);
        node.Format["done"] = cmtDone ? "true" : "false";
        if (cmtParentId != null) node.Format["parentId"] = cmtParentId;
        // R21-WB-1: surface direction from the first content paragraph's
        // pPr.BiDi so the cascade (already applied by ApplyCommentFormatKeys)
        // round-trips through Get. Mirrors footnote/endnote readback above.
        var cmtBidi = comment.Descendants<Paragraph>().FirstOrDefault()?.ParagraphProperties?.GetFirstChild<BiDi>();
        if (cmtBidi != null)
            node.Format["direction"] = TryReadOnOff(cmtBidi.Val) == true ? "rtl" : "ltr";
        return node;
    }

    private DocumentNode SectionPropertiesToNode(SectionProperties sectPrEl, string path)
    {
        // CONSISTENCY(section-readback): /body/sectPr[N] should surface
        // the same Format keys as /section[N] so direction, page size,
        // margins, etc. are visible regardless of which path the caller
        // used. Delegate to BuildSectionNode but preserve the original
        // path the caller asked for.
        return BuildSectionNode(sectPrEl, path);
    }

    // BUG-DUMP-DELININS / BUG-DUMP-MOVE-DEL: a run wrapped by an outer revision
    // (ins / moveFrom / moveTo) AND an inner <w:del> must surface the inner
    // deletion as revision.nested.* so the emitter rebuilds the nested
    // <w:ins|moveFrom|moveTo><w:del> stack. Without it the <w:del> wrapper was
    // dropped and the deleted text resurfaced as live, accepted content (a silent
    // meaning change). ECMA-376 permits ins/moveFrom/moveTo to wrap a del.
    private static void CaptureNestedDeletion(Run run, DocumentNode node)
    {
        var nestedDel = run.Ancestors<DeletedRun>().FirstOrDefault();
        if (nestedDel == null) return;
        node.Format["revision.nested.type"] = "del";
        if (!string.IsNullOrEmpty(nestedDel.Author?.Value))
            node.Format["revision.nested.author"] = nestedDel.Author!.Value!;
        if (nestedDel.Date?.Value is DateTime nestedDelDate)
            node.Format["revision.nested.date"] = nestedDelDate.ToString("o");
        if (nestedDel.Id?.Value is { } nestedDelId)
            node.Format["revision.nested.id"] = nestedDelId.ToString();
    }

    // BUG-DUMP-H99: detect a run that PACKS an entire field's fldChar chain
    // (begin + <w:instrText> + separate + cached result + end) into a SINGLE
    // <w:r>, instead of Word's usual one-structural-element-per-run shape. Such
    // runs are emitted by some converters (RTF→docx, Markdown→docx) and a few
    // apps for footer PAGE / STYLEREF fields. RunToNode surfaces only the FIRST
    // structural element (GetFirstChild<FieldChar>()), so the instruction,
    // separator, cached result and end fldChar were silently lost — and
    // CollapseFieldChains then saw a lone begin and warn-dropped the whole field
    // (the page number vanished). A run is "packed" when it carries at least one
    // field-structural element (fldChar / instrText / delInstrText) alongside any
    // other emittable child; the caller splits it into the same one-element-per-
    // run node stream Word emits so the collapse pass rebuilds the field.
    private static bool IsPackedFieldRun(Run run)
    {
        var fldChars = run.Elements<FieldChar>().ToList();
        if (fldChars.Count == 0) return false;
        bool hasBegin = fldChars.Any(f => string.Equals(
            f.FieldCharType?.InnerText, "begin", StringComparison.OrdinalIgnoreCase));
        bool hasEnd = fldChars.Any(f => string.Equals(
            f.FieldCharType?.InnerText, "end", StringComparison.OrdinalIgnoreCase));
        // Only a SELF-CONTAINED field (begin … end packed in one <w:r>) is split.
        // A run carrying only a partial fragment (e.g. begin + instrText with the
        // separate/end in sibling runs, or a lone begin — authored-garbage shapes)
        // is left to RunToNode's first-match-wins type precedence: splitting it
        // wouldn't help the round-trip (it still collapses to an unmatched-begin
        // warn-drop) and would break the documented single-node contract
        // (WordRunSpecialContentBugFixTests.FieldChar_PlusInstrText_InSameRun_StableType).
        return hasBegin && hasEnd;
    }

    private DocumentNode RunToNode(Run run, DocumentNode node, string path)
    {
        node.Type = "run";
        node.Text = GetRunText(run);
        // BUG-DUMP-PGNUM: a run containing <w:pgNum/> (page-number placeholder,
        // valid inside header/footer runs) has no scalar add/set representation
        // — GetRunText surfaces no glyph for it and the typed `add r` path drops
        // it entirely on dump→batch round-trip, leaving an uncropped/blank run.
        // Same class as the ruby / rich-break raw-set fallbacks: stamp a sentinel
        // so TryEmitPgNumRun re-inserts the verbatim <w:r> at its source position.
        // node.Type stays "run" — the flag only routes the emitter, not the
        // readback. A co-located <w:cr/> (which the typed path otherwise demotes
        // to <w:br/>) rides along inside the same verbatim run XML.
        if (run.GetFirstChild<PageNumber>() != null)
            node.Format["_hasPgNum"] = true;
        // BUG-DUMP-R40-3: a run containing <w:noBreakHyphen/> (non-breaking
        // hyphen glyph) or <w:softHyphen/> (discretionary hyphen) — siblings of
        // <w:t>/<w:tab>/<w:br> inside the run — has no scalar add/set
        // representation. GetRunText surfaces the Unicode glyph (U+2011 / U+00AD)
        // as Text, but the typed `add p text="…"`/`add r` path persists that
        // glyph as literal <w:t> text and the structural hyphen element vanishes
        // from the round-trip. Mirror the pgNum/dateField raw-set fallback: stamp
        // a sentinel so the run stays on the explicit-run path and
        // TryEmitHyphenRun re-inserts the verbatim <w:r> (the hyphen element AND
        // any co-located <w:t> text) at its source position. node.Type/node.Text
        // are untouched — the flag only routes the emitter.
        // Record the KIND ("soft"/"noBreak") so every emit site (TryEmitHyphenRun
        // AND EmitStructuredHyperlink's trailing-run path) can rebuild the right
        // structural element without re-probing the raw XML. _hasHyphen stays
        // truthy (non-empty string) so the existing routing checks are unaffected.
        if (run.GetFirstChild<SoftHyphen>() != null)
            node.Format["_hasHyphen"] = "soft";
        else if (run.GetFirstChild<NoBreakHyphen>() != null)
            node.Format["_hasHyphen"] = "noBreak";
        // BUG-DUMP-R40-2: surface <w:annotationRef/> (the comment-reference mark
        // that opens every Word-authored comment body). GetRunText emits no
        // glyph for it, so the run looked empty and the typed `add comment`/`add
        // r` path dropped the mark — the rebuilt comment lost its clickable
        // reference mark and the run's rStyle="CommentReference". Stamp a flag so
        // EmitComments rebuilds the annotationRef run verbatim (with its rStyle).
        if (run.GetFirstChild<AnnotationReferenceMark>() != null)
            node.Format["annotationRef"] = true;
        // BUG-DUMP-DATEFIELD: a run containing a Word date-component placeholder
        // element — <w:dayLong/> / <w:dayShort/> / <w:monthLong/> / <w:monthShort/>
        // / <w:yearLong/> / <w:yearShort/> (Word substitutes the current date at
        // render) — has no scalar add/set representation. GetRunText surfaces a
        // human "[dayLong]" sentinel (fine for `view text`), but the typed
        // `add p text="…"` collapse persists that sentinel as LITERAL <w:t> text
        // and the date element vanishes from the round-trip. Mirror the pgNum
        // raw-set fallback: stamp a sentinel so ShouldCollapseSingleRun keeps the
        // run on the explicit-run path and TryEmitDateFieldRun re-inserts the
        // verbatim <w:r> (co-located <w:t> text AND the date element) at its
        // source position. node.Type / node.Text are untouched — the flag only
        // routes the emitter, not the readback or the human sentinel.
        if (run.GetFirstChild<DayLong>() != null
            || run.GetFirstChild<DayShort>() != null
            || run.GetFirstChild<MonthLong>() != null
            || run.GetFirstChild<MonthShort>() != null
            || run.GetFirstChild<YearLong>() != null
            || run.GetFirstChild<YearShort>() != null)
            node.Format["_hasDateField"] = true;
        // BUG-DUMP7-01: surface <w:sym w:font=… w:char=…/> as a `sym`
        // Format key (font:hex). GetRunText also surfaces the resolved
        // Unicode glyph as Text so the run looks non-empty, but Text
        // alone is lossy — Wingdings F0E0 ↦ U+F0E0 would replay as a
        // plain text run in a non-symbol font and the glyph would
        // disappear. AddRun consumes `sym=` to rebuild SymbolChar.
        var symEl = run.GetFirstChild<SymbolChar>();
        if (symEl?.Char?.Value != null)
        {
            var symFontVal = symEl.Font?.Value ?? "";
            node.Format["sym"] = $"{symFontVal}:{symEl.Char.Value}";
        }
        // BUG-DUMP4-02: surface track-change attribution from any
        // InsertedRun/DeletedRun ancestor wrapping this run. Descendants<Run>
        // unwraps the wrapper so the run looks plain on the curated
        // surface; without this the author/date attribution silently
        // disappears on dump round-trip even though the inner text
        // survives.
        var insAncestor = run.Ancestors<InsertedRun>().FirstOrDefault();
        var moveFromAncestor = insAncestor == null ? run.Ancestors<MoveFromRun>().FirstOrDefault() : null;
        var moveToAncestor = (insAncestor == null && moveFromAncestor == null)
            ? run.Ancestors<MoveToRun>().FirstOrDefault() : null;
        if (insAncestor != null)
        {
            node.Format["revision.type"] = "ins";
            if (!string.IsNullOrEmpty(insAncestor.Author?.Value))
                node.Format["revision.author"] = insAncestor.Author!.Value!;
            if (insAncestor.Date?.Value is DateTime insDate)
                node.Format["revision.date"] = insDate.ToString("o");
            if (insAncestor.Id?.Value is { } insId)
                node.Format["revision.id"] = insId.ToString();
            // BUG-DUMP-DELININS: a <w:ins> may itself contain a <w:del> (one
            // reviewer inserts text, a second reviewer deletes that insertion).
            // The run is then BOTH inserted and deleted and its text rides in
            // <w:delText>. A single revision.type can't carry both wrappers, so
            // the inner del was dropped — the deletion round-tripped as a live
            // insertion (<w:delText> rebuilt as <w:t>, the deleted content
            // silently un-deleted; this is the cd241 delText-loss class).
            // A run with BOTH an ins AND a del ancestor is ins-outer/del-inner
            // (ECMA-376 permits ins⊃del). Capture the inner del as revision.nested.*
            // so the emitter rebuilds the <w:ins><w:del> stack. (moveFrom/moveTo may
            // also wrap a del — same capture, see those branches.)
            CaptureNestedDeletion(run, node);
        }
        else if (moveFromAncestor != null)
        {
            // CONSISTENCY(revision-wrapper-readback): w:moveFrom/w:moveTo wrap
            // a run identically to w:ins/w:del; without a Get-side branch the
            // wrapper attribution silently disappeared on dump round-trip and
            // the BatchEmitter re-emitted the run as a plain `add r` (losing
            // the move pairing entirely). Add accepts revision.type=moveFrom/
            // moveTo + author/date/id.
            node.Format["revision.type"] = "moveFrom";
            if (!string.IsNullOrEmpty(moveFromAncestor.Author?.Value))
                node.Format["revision.author"] = moveFromAncestor.Author!.Value!;
            if (moveFromAncestor.Date?.Value is DateTime mfDate)
                node.Format["revision.date"] = mfDate.ToString("o");
            if (moveFromAncestor.Id?.Value is { } mfId)
                node.Format["revision.id"] = mfId.ToString();
            CaptureNestedDeletion(run, node);   // BUG-DUMP-MOVE-DEL: moveFrom⊃del
        }
        else if (moveToAncestor != null)
        {
            node.Format["revision.type"] = "moveTo";
            if (!string.IsNullOrEmpty(moveToAncestor.Author?.Value))
                node.Format["revision.author"] = moveToAncestor.Author!.Value!;
            if (moveToAncestor.Date?.Value is DateTime mtDate)
                node.Format["revision.date"] = mtDate.ToString("o");
            if (moveToAncestor.Id?.Value is { } mtId)
                node.Format["revision.id"] = mtId.ToString();
            CaptureNestedDeletion(run, node);   // BUG-DUMP-MOVE-DEL: moveTo⊃del
        }
        else
        {
            var delAncestor = run.Ancestors<DeletedRun>().FirstOrDefault();
            if (delAncestor != null)
            {
                node.Format["revision.type"] = "del";
                if (!string.IsNullOrEmpty(delAncestor.Author?.Value))
                    node.Format["revision.author"] = delAncestor.Author!.Value!;
                if (delAncestor.Date?.Value is DateTime delDate)
                    node.Format["revision.date"] = delDate.ToString("o");
                if (delAncestor.Id?.Value is { } delId)
                    node.Format["revision.id"] = delId.ToString();
            }
            else
            {
                // AddRun writes <w:rPrChange> for `trackChange=format`. The
                // rPrChange block carries the same author/date attribution
                // as the ins/del wrappers, but rides inside <w:rPr> rather
                // than wrapping the run.
                var rPrChange = run.RunProperties?.GetFirstChild<RunPropertiesChange>();
                if (rPrChange != null)
                {
                    node.Format["revision.type"] = "format";
                    if (!string.IsNullOrEmpty(rPrChange.Author?.Value))
                        node.Format["revision.author"] = rPrChange.Author!.Value!;
                    if (rPrChange.Date?.Value is DateTime rDate)
                        node.Format["revision.date"] = rDate.ToString("o");
                    if (rPrChange.Id?.Value is { } rId)
                        node.Format["revision.id"] = rId.ToString();
                    // BUG: rPrChange's inner snapshot (PreviousRunProperties
                    // when written by SetRevision; plain RunProperties when
                    // written by AddRun's empty-marker path; foreign producers
                    // may use either) carries the pre-change run properties
                    // (bold/italic/color/size/…). AddRun's rPrChange branch
                    // re-stamps the marker with an empty snapshot, losing
                    // what Word's Reject Change would restore. Surface
                    // beforeLost so BatchEmitter can warn. See
                    // WordHandler.Set.Revision.cs note on which strongly-typed
                    // child class round-trips through OpenXml 3.x.
                    var prevPrev = rPrChange.GetFirstChild<PreviousRunProperties>();
                    var prevPlain = rPrChange.GetFirstChild<RunProperties>();
                    OpenXmlElement? prevRun = (prevPrev != null && prevPrev.HasChildren)
                        ? prevPrev
                        : (prevPlain != null && prevPlain.HasChildren) ? prevPlain : null;
                    // BUG-DUMP-R43-8: capture the verbatim prior-rPr snapshot so
                    // the emitter can restore it (revision.beforeXml) instead of
                    // dropping it and emitting an empty marker. Falls back to the
                    // beforeLost warning only when the inner snapshot is empty
                    // (nothing to carry).
                    if (prevRun != null)
                        node.Format["revision.beforeXml"] = prevRun.OuterXml;
                }
            }
        }
        // CONSISTENCY(canonical-keys): mirror style Get (WordHandler.Query.cs:546-553) —
        // emit per-script font slots, no flat "font" alias. R6 BUG-1: previously
        // collapsed all 4 slots into a single "font" via GetRunFont (Ascii first).
        var rFonts = run.RunProperties?.RunFonts;
        if (rFonts != null)
        {
            // CONSISTENCY(canonical-keys): collapse Ascii+HighAnsi into
            // `font.latin` (canonical per schema docx/run.json) when they
            // match — the round-trip case for `font.latin=` Set. Differing
            // slots fall back to legacy `font.ascii` / `font.hAnsi` keys.
            var ascii = string.IsNullOrEmpty(rFonts.Ascii?.Value) ? null : rFonts.Ascii!.Value;
            var hAnsi = string.IsNullOrEmpty(rFonts.HighAnsi?.Value) ? null : rFonts.HighAnsi!.Value;
            if (ascii != null && hAnsi != null && ascii == hAnsi)
                node.Format["font.latin"] = ascii;
            else
            {
                if (ascii != null && hAnsi != null)
                {
                    node.Format["font.ascii"] = ascii;
                    node.Format["font.hAnsi"] = hAnsi;
                }
                // BUG-DUMP-FONT-LATIN: ascii-only (hAnsi absent) must emit
                // font.ascii, NOT font.latin. font.latin sets BOTH slots on
                // replay, so a source <w:rFonts ascii="X"/> round-tripped to
                // ascii="X" hAnsi="X" — silently rebinding extended-Latin
                // glyphs (which had fallen back to the style/docDefaults font)
                // to X. Same for hAnsi-only.
                else if (ascii != null) node.Format["font.ascii"] = ascii;
                else if (hAnsi != null) node.Format["font.hAnsi"] = hAnsi;
            }
            if (!string.IsNullOrEmpty(rFonts.EastAsia?.Value)) node.Format["font.ea"] = rFonts.EastAsia!.Value!;
            // BUG-DUMP14-03: theme-font slots (asciiTheme/hAnsiTheme/
            // eastAsiaTheme/cstheme) bind a run to a theme major/minor
            // font instead of a literal face name. Without surfacing
            // them, documents using theme fonts lose all font bindings
            // on round-trip (only literal Ascii/HighAnsi were read).
            if (rFonts.AsciiTheme?.HasValue == true)
                node.Format["font.asciiTheme"] = rFonts.AsciiTheme.InnerText;
            if (rFonts.HighAnsiTheme?.HasValue == true)
                node.Format["font.hAnsiTheme"] = rFonts.HighAnsiTheme.InnerText;
            if (rFonts.EastAsiaTheme?.HasValue == true)
                node.Format["font.eaTheme"] = rFonts.EastAsiaTheme.InnerText;
            if (rFonts.ComplexScriptTheme?.HasValue == true)
                node.Format["font.csTheme"] = rFonts.ComplexScriptTheme.InnerText;
            // <w:rFonts w:hint="eastAsia|cs|default"/> selects which font slot
            // renders ambiguous characters (CJK digits/punctuation). Dropping it
            // changes glyph widths on round-trip — a tight CJK line can rewrap.
            // Round-trips via the font.hint key (ApplyRunFormatting writes it).
            if (rFonts.Hint?.HasValue == true)
                node.Format["font.hint"] = rFonts.Hint.InnerText;
        }
        // <w:lang/> three slots: val (latin) / eastAsia / bidi (cs).
        // CONSISTENCY(canonical-keys): mirror font.latin/font.ea/font.cs vocabulary.
        var rLang = run.RunProperties?.GetFirstChild<Languages>();
        if (rLang != null)
        {
            if (rLang.Val?.Value != null) node.Format["lang.latin"] = rLang.Val.Value;
            if (rLang.EastAsia?.Value != null) node.Format["lang.ea"] = rLang.EastAsia.Value;
            if (rLang.Bidi?.Value != null) node.Format["lang.cs"] = rLang.Bidi.Value;
        }
        var size = GetRunFontSize(run);
        if (size != null) node.Format["size"] = size;
        if (run.RunProperties?.Bold != null) node.Format["bold"] = IsToggleOn(run.RunProperties.Bold);
        if (run.RunProperties?.Italic != null) node.Format["italic"] = IsToggleOn(run.RunProperties.Italic);
        // Complex-script readback (font.cs / size.cs / bold.cs / italic.cs).
        // See WordHandler.I18n.cs.
        ReadComplexScriptRunFormatting(run.RunProperties, null, node.Format);
        // BUG-DUMP-R44-1: direct run color must carry BOTH the explicit hex val
        // AND the theme linkage (themeColor/themeShade/themeTint) via the shared
        // ';themeColor=…' tail convention. The old code emitted only the theme
        // name when ThemeColor was present, dropping an explicit w:val (e.g.
        // <w:color w:val="FFFFFF" w:themeColor="background1"/> baked white text
        // that rebuilt as val="auto" → invisible black-on-black). Mirrors the
        // style-scoped fix (StyleColorWithThemeTail, BUG-DUMP-R43-3).
        if (run.RunProperties?.Color != null && StyleColorWithThemeTail(run.RunProperties.Color) is { } runColor)
            node.Format["color"] = runColor;
        if (run.RunProperties?.Underline?.Val != null) node.Format["underline"] = run.RunProperties.Underline.Val.InnerText;
        // CONSISTENCY(underline-color): backfilled from style Get edc8f884.
        if (run.RunProperties?.Underline?.Color?.Value != null)
            node.Format["underline.color"] = ParseHelpers.FormatHexColor(run.RunProperties.Underline.Color.Value);
        if (run.RunProperties?.Strike != null) node.Format["strike"] = IsToggleOn(run.RunProperties.Strike);
        if (run.RunProperties?.Highlight?.Val != null) node.Format["highlight"] = run.RunProperties.Highlight.Val.InnerText;
        if (run.RunProperties?.Caps != null) node.Format["caps"] = IsToggleOn(run.RunProperties.Caps);
        if (run.RunProperties?.SmallCaps != null) node.Format["smallcaps"] = IsToggleOn(run.RunProperties.SmallCaps);
        if (run.RunProperties?.DoubleStrike != null) node.Format["dstrike"] = IsToggleOn(run.RunProperties.DoubleStrike);
        if (run.RunProperties?.Vanish != null) node.Format["vanish"] = IsToggleOn(run.RunProperties.Vanish);
        if (run.RunProperties?.Outline != null) node.Format["outline"] = IsToggleOn(run.RunProperties.Outline);
        if (run.RunProperties?.Shadow != null) node.Format["shadow"] = IsToggleOn(run.RunProperties.Shadow);
        if (run.RunProperties?.Emboss != null) node.Format["emboss"] = IsToggleOn(run.RunProperties.Emboss);
        if (run.RunProperties?.Imprint != null) node.Format["imprint"] = IsToggleOn(run.RunProperties.Imprint);
        if (run.RunProperties?.NoProof != null) node.Format["noproof"] = IsToggleOn(run.RunProperties.NoProof);
        if (run.RunProperties?.RightToLeftText != null)
        {
            // <w:rtl/> with no Val attribute implies true; <w:rtl w:val="0"/>
            // is an explicit off-override (overrides inherited docDefaults).
            // CONSISTENCY(canonical-key): paragraphs and sections surface
            // this property as Format["direction"]="rtl"|"ltr"; runs must
            // match so users see one canonical key across scopes (R16-bt-1).
            var rtlVal = run.RunProperties.RightToLeftText.Val;
            var on = rtlVal == null ? true : rtlVal.Value;
            node.Format["direction"] = on ? "rtl" : "ltr";
        }
        if (run.RunProperties?.VerticalTextAlignment?.Val?.Value == VerticalPositionValues.Superscript)
            node.Format["superscript"] = true;
        if (run.RunProperties?.VerticalTextAlignment?.Val?.Value == VerticalPositionValues.Subscript)
            node.Format["subscript"] = true;
        // ApplyRunFormatting writes <w:position> for `position` (raised /
        // lowered baseline offset in half-points). Mirror it on the Get
        // side so the round-trip key survives.
        var posVal = run.RunProperties?.GetFirstChild<Position>()?.Val?.Value;
        if (!string.IsNullOrEmpty(posVal)
            && int.TryParse(posVal, System.Globalization.NumberStyles.Integer,
                System.Globalization.CultureInfo.InvariantCulture, out var posHalfPts))
            node.Format["position"] = $"{posHalfPts / 2.0:0.##}pt";
        if (run.RunProperties?.Spacing?.Val?.HasValue == true)
            node.Format["charSpacing"] = $"{run.RunProperties.Spacing.Val.Value / 20.0:0.##}pt";
        // BUG-DUMP22-08: <w:bdr/> (character border) is multi-attribute
        // (val + sz + color + space) so the long-tail FillUnknownChildProps
        // skipped it (attrCount > 1), leaving only the surface bare key
        // with no sub-attrs. Emit the colon-encoded compound form that
        // ApplyRunFormatting consumes on replay so dump round-trips
        // preserve size and color.
        var rBdr = run.RunProperties?.GetFirstChild<Border>();
        if (rBdr?.Val?.HasValue == true)
        {
            var bdrStyle = rBdr.Val!.InnerText;
            var bdrSize = rBdr.Size?.Value;
            var bdrColor = rBdr.Color?.Value;
            var bdrSpace = rBdr.Space?.Value;
            // BUG-DUMP-R36-1: append w:shadow / w:frame as positional segments
            // 5/6 (STYLE;SIZE;COLOR;SPACE;SHADOW;FRAME). Emit only when present
            // so a plain border keeps its 4-field shape (back-compat) and does
            // not gain a spurious shadow="false"/frame="false" on replay.
            var segs = new System.Collections.Generic.List<string>
            {
                bdrStyle ?? "single",
                bdrSize?.ToString() ?? "",
                string.IsNullOrEmpty(bdrColor) ? "" : ParseHelpers.FormatHexColor(bdrColor),
                bdrSpace?.ToString() ?? "0"
            };
            bool? bdrShadow = rBdr.Shadow?.Value;
            bool? bdrFrame = rBdr.Frame?.Value;
            if (bdrShadow.HasValue || bdrFrame.HasValue)
                segs.Add(bdrShadow == true ? "true" : "false");
            if (bdrFrame.HasValue)
                segs.Add(bdrFrame == true ? "true" : "false");
            node.Format["bdr"] = string.Join(';', segs);
        }
        if (run.RunProperties?.Shading != null)
        {
            // CONSISTENCY(shd-canonical-fill): solid run shading reads back as
            // the canonical `fill` key (matches table cells / paragraphs); true
            // pattern/theme keeps the shading.val/.fill/.color detail keys (the
            // dump→batch fold consumes them). w:highlight and the w14 text
            // shadow are separate elements handled elsewhere — only w:shd here.
            ReadShadingCanonical(run.RunProperties.Shading, node);
        }
        // w14 text effects
        ReadW14TextEffects(run.RunProperties, node);
        // BUG-DUMP10-01: w:eastAsianLayout (vert/combine/vertCompress)
        // is a multi-attribute child the long-tail FillUnknownChildProps
        // skips (it only handles single-val/no-attr leaves). Without an
        // explicit reader, vertical-text and two-lines-in-one CJK layout
        // was silently dropped on dump→batch round-trip. Set side is
        // covered by TypedAttributeFallback.TrySet which creates the
        // dotted child + attr automatically.
        if (run.RunProperties?.GetFirstChild<EastAsianLayout>() is { } eal)
        {
            // BUG-DUMP-EALID: w:id links runs in the same two-lines-in-one
            // combine group (a semantic grouping). Curated reader emitted
            // vert/combine/vertCompress/combineBrackets but dropped w:id on
            // dump→batch round-trip. Set side already accepts eastAsianLayout.id.
            // BUG-DUMP-R32-1: read eastAsianLayout attrs as RAW strings, never
            // the strongly-typed SDK accessors. w:combine / w:vert are declared
            // ST_OnOff in the schema, but Word's "Combine Characters" feature
            // writes w:combine="lines" / "letters" (and brackets variants), which
            // the typed getter (eal.Combine.Value) parses as ST_OnOff and THROWS
            // ("text value invalid… only true/false/on/off/0/1"), aborting the
            // ENTIRE dump. GetAttribute returns whatever string the source had so
            // we round-trip "lines"/"letters"/"1"/etc. verbatim.
            string? RawEalAttr(string localName) =>
                eal.GetAttributes().FirstOrDefault(a =>
                    a.LocalName.Equals(localName, StringComparison.Ordinal)).Value;
            var ealId             = RawEalAttr("id");
            var ealVert           = RawEalAttr("vert");
            var ealCombine        = RawEalAttr("combine");
            var ealVertCompress   = RawEalAttr("vertCompress");
            var ealCombineBracket = RawEalAttr("combineBrackets");
            if (!string.IsNullOrEmpty(ealId)) node.Format["eastAsianLayout.id"] = ealId;
            // vert / combine are ST_OnOff but Word may emit non-boolean values
            // (combine="lines"/"letters"). Emit "1" for the truthy boolean forms
            // (keeps the prior dotted-key contract) and pass any other value
            // through verbatim so the rebuild reproduces the source attribute.
            if (!string.IsNullOrEmpty(ealVert))
                node.Format["eastAsianLayout.vert"] = NormalizeEalOnOff(ealVert);
            if (!string.IsNullOrEmpty(ealCombine))
                node.Format["eastAsianLayout.combine"] = NormalizeEalOnOff(ealCombine);
            if (!string.IsNullOrEmpty(ealVertCompress))
                node.Format["eastAsianLayout.vertCompress"] = ealVertCompress;
            if (!string.IsNullOrEmpty(ealCombineBracket))
                node.Format["eastAsianLayout.combineBrackets"] = ealCombineBracket;
        }
        // BUG-DUMP-R32-1: collapse the boolean ST_OnOff spellings to "1" (the
        // dotted-key contract the apply side already understands) while passing
        // any non-boolean value (combine="lines"/"letters") through unchanged.
        static string NormalizeEalOnOff(string raw) =>
            raw is "true" or "on" or "1" ? "1"
            : raw is "false" or "off" or "0" ? "0"
            : raw;
        // Long-tail fallback: surface every rPr child the curated reader
        // didn't consume. Symmetric with the Set-side TryCreateTypedChild
        // fallback in SetElementRun (WordHandler.Set.Element.cs).
        FillUnknownChildProps(run.RunProperties, node);
        // Image properties if run contains a Drawing.
        // BUG-R5-T3: previously this branch wrote only id/name/alt/width/
        // height/relId — wrap/hPosition/vPosition/hRelative/vRelative/
        // behindText for floating pictures were silently dropped, which
        // also broke dump→batch round-trip (WordBatchEmitter relies on Get).
        // Reuse CreateImageNode (the canonical picture-node builder) and
        // merge its Format bag into the run node.
        var runDrawing = run.GetFirstChild<Drawing>();
        if (runDrawing != null)
        {
            var picNode = CreateImageNode(runDrawing, run, path);
            node.Type = picNode.Type;
            if (!string.IsNullOrEmpty(picNode.Text)) node.Text = picNode.Text;
            foreach (var kv in picNode.Format)
                node.Format[kv.Key] = kv.Value;
        }
        // OLE object if run contains an EmbeddedObject. The underlying
        // logic is the same as CreateOleNode — reuse it so Get/Query
        // return identical shapes.
        var runOle = run.GetFirstChild<EmbeddedObject>();
        if (runOle != null)
        {
            // CONSISTENCY(ole-host-part): mirror Query.cs's header/footer
            // OLE handling — the EmbeddedObjectPart relationship lives on
            // the owning Header/Footer part, not the MainDocumentPart.
            // Walk ancestors to find the host part so CreateOleNode can
            // populate contentType/fileSize instead of returning orphan.
            OpenXmlPart? hostPart = _doc.MainDocumentPart;
            var headerAncestor = run.Ancestors<Header>().FirstOrDefault();
            if (headerAncestor != null && _doc.MainDocumentPart != null)
            {
                var hp = _doc.MainDocumentPart.HeaderParts
                    .FirstOrDefault(p => ReferenceEquals(p.Header, headerAncestor));
                if (hp != null) hostPart = hp;
            }
            else
            {
                var footerAncestor = run.Ancestors<Footer>().FirstOrDefault();
                if (footerAncestor != null && _doc.MainDocumentPart != null)
                {
                    var fp = _doc.MainDocumentPart.FooterParts
                        .FirstOrDefault(p => ReferenceEquals(p.Footer, footerAncestor));
                    if (fp != null) hostPart = fp;
                }
            }
            var oleNode = CreateOleNode(runOle, run, path, hostPart);
            // Keep the node's path as-is, but swap in the OLE-sourced
            // type/format bag.
            node.Type = oleNode.Type;
            foreach (var kv in oleNode.Format)
                node.Format[kv.Key] = kv.Value;
            if (!string.IsNullOrEmpty(oleNode.Text))
                node.Text = oleNode.Text;
        }
        // CONSISTENCY(run-special-content): runs that primarily carry inline
        // structure (ptab, fldChar, instrText, tab, break) instead of a
        // <w:t> payload were previously surfaced as opaque
        // {type:"run", text:""} placeholders — six of these in a row in
        // header/footer paragraphs (PAGE field begin/instr/separate/end +
        // ptab anchors), all indistinguishable. Upgrade the node.Type so
        // callers walking paragraph.children can rebuild left/center/right
        // alignment regions and detect field markers without reparsing the
        // raw OOXML themselves. Mirrors the type=picture / type=ole
        // pattern above.
        //
        // Each block is gated on `node.Type == "run"` so that:
        //   (a) Drawing/EmbeddedObject (already upgraded above to
        //       picture/ole) wins over a co-residing <w:br>/<w:tab> —
        //       picture+break is a real Word emission and the picture
        //       identity must not be silently overwritten;
        //   (b) the first matching structural element wins when several
        //       coexist in one run (rare but possible), keeping node.Type
        //       single-valued and deterministic. ptab is checked first
        //       (most semantically distinctive), then fieldChar, then
        //       instrText, then tab, then break.
        if (node.Type == "run")
        {
            var ptabEl = run.GetFirstChild<PositionalTab>();
            if (ptabEl != null)
            {
                // Open XML SDK v3 enum .ToString() returns "FooValues { }"
                // — use .InnerText to get the actual XML attribute value
                // ("center", "right", "begin", etc.). Same trap as the
                // LineSpacingRuleValues note in the Word handler conventions.
                var ptabAlign = ptabEl.Alignment?.HasValue == true ? ptabEl.Alignment.InnerText : null;
                var ptabRelTo = ptabEl.RelativeTo?.HasValue == true ? ptabEl.RelativeTo.InnerText : null;
                var ptabLeader = ptabEl.Leader?.HasValue == true ? ptabEl.Leader.InnerText : null;
                // BUG-DUMP-DELTAB sibling (BUG-DUMP-PTABTEXT): when the run ALSO
                // carries text (<w:t> or <w:delText>), the standalone `add ptab`
                // type-upgrade would drop that co-resident text — and unlike a
                // plain tab the ptab can't ride back through GetRunText's \t, so
                // keeping it a plain run would instead drop the ptab. Keep the run
                // a `run` AND carry the ptab as inline props so AddRun rebuilds
                // <w:ptab/> ahead of the text, preserving BOTH.
                bool hasText = run.Elements<Text>().Any() || run.Elements<DeletedText>().Any();
                if (hasText)
                {
                    node.Format["ptabInline"] = "true";
                    if (ptabAlign != null) node.Format["ptabInline.align"] = ptabAlign;
                    if (ptabRelTo != null) node.Format["ptabInline.relativeTo"] = ptabRelTo;
                    if (ptabLeader != null) node.Format["ptabInline.leader"] = ptabLeader;
                }
                else
                {
                    node.Type = "ptab";
                    if (ptabAlign != null) node.Format["align"] = ptabAlign;
                    if (ptabRelTo != null) node.Format["relativeTo"] = ptabRelTo;
                    if (ptabLeader != null) node.Format["leader"] = ptabLeader;
                }
            }
        }
        if (node.Type == "run")
        {
            var fldCharEl = run.GetFirstChild<FieldChar>();
            if (fldCharEl != null)
            {
                node.Type = "fieldChar";
                if (fldCharEl.FieldCharType?.HasValue == true)
                    node.Format["fieldCharType"] = fldCharEl.FieldCharType.InnerText;
                // CONSISTENCY(field-cache-stale): expose dirty so audit
                // tools can verify whether Set instr / Set cached
                // properly flagged the owning field for recompute. The
                // attribute persists in OOXML; surfacing it via Get
                // closes the loop the Round 3 dirty fix opened.
                if (fldCharEl.Dirty?.Value == true)
                    node.Format["dirty"] = true;
                // BUG-DUMP-R37-4: <w:fldChar w:fldLock="true"> — the field is
                // locked against F9/recalc. Lives on the begin fldChar; surface
                // it so CollapseFieldChains can carry it onto the synthetic
                // field node and AddField re-applies it on replay.
                if (fldCharEl.FieldLock?.Value == true)
                    node.Format["fldLock"] = true;
                if (fldCharEl.FormFieldData != null)
                {
                    node.Format["hasFormFieldData"] = true;
                    // R14-bug1+2: surface the ffData payload on the fieldChar
                    // begin node so WordBatchEmitter.CollapseFieldChains can
                    // emit a `type=formfield` synth carrying the full
                    // ffData props (name, default, maxLength, checkBox.size,
                    // checkBox.default, ddList items, helpText, statusText,
                    // entryMacro, exitMacro, calcOnExit, textInput.type,
                    // textInput.format). Without these keys, the dump
                    // emitted a bare `add field instr=FORMTEXT` row that
                    // walked the BuildFieldAddProps default arm and lost
                    // every ffData wrapper on replay.
                    var ffd = fldCharEl.FormFieldData;
                    var ffName = ffd.GetFirstChild<FormFieldName>()?.Val?.Value;
                    if (!string.IsNullOrEmpty(ffName))
                        node.Format["ffName"] = ffName;
                    var ffEnabled = ffd.GetFirstChild<Enabled>();
                    if (ffEnabled != null)
                        node.Format["ffEnabled"] = ffEnabled.Val?.Value ?? true;
                    var ffHelp = ffd.GetFirstChild<HelpText>()?.Val?.Value;
                    if (!string.IsNullOrEmpty(ffHelp))
                        node.Format["ffHelpText"] = ffHelp;
                    var ffStatus = ffd.GetFirstChild<StatusText>()?.Val?.Value;
                    if (!string.IsNullOrEmpty(ffStatus))
                        node.Format["ffStatusText"] = ffStatus;
                    var ffEntry = ffd.GetFirstChild<EntryMacro>()?.Val?.Value;
                    if (!string.IsNullOrEmpty(ffEntry))
                        node.Format["ffEntryMacro"] = ffEntry;
                    var ffExit = ffd.GetFirstChild<ExitMacro>()?.Val?.Value;
                    if (!string.IsNullOrEmpty(ffExit))
                        node.Format["ffExitMacro"] = ffExit;
                    var ffCalc = ffd.GetFirstChild<CalculateOnExit>();
                    if (ffCalc != null)
                        node.Format["ffCalcOnExit"] = ffCalc.Val?.Value ?? true;

                    var ti = ffd.GetFirstChild<TextInput>();
                    var cb = ffd.GetFirstChild<CheckBox>();
                    var dd = ffd.GetFirstChild<DropDownListFormField>();
                    if (ti != null)
                    {
                        node.Format["ffType"] = "text";
                        var tDef = ti.GetFirstChild<DefaultTextBoxFormFieldString>()?.Val?.Value;
                        if (!string.IsNullOrEmpty(tDef))
                            node.Format["ffDefault"] = tDef;
                        var tMax = ti.GetFirstChild<MaxLength>()?.Val?.Value;
                        if (tMax != null && tMax.Value != 0)
                            node.Format["ffMaxLength"] = (int)tMax;
                        var tTyp = ti.GetFirstChild<TextBoxFormFieldType>()?.Val?.InnerText;
                        if (!string.IsNullOrEmpty(tTyp))
                            node.Format["ffTextType"] = tTyp;
                        var tFmt = ti.GetFirstChild<Format>()?.Val?.Value;
                        if (!string.IsNullOrEmpty(tFmt))
                            node.Format["ffTextFormat"] = tFmt;
                    }
                    else if (cb != null)
                    {
                        node.Format["ffType"] = "checkbox";
                        var cChecked = cb.GetFirstChild<Checked>();
                        var cDefault = cb.GetFirstChild<DefaultCheckBoxFormFieldState>();
                        // BUG-DUMP-FFCHECKBOX-DEFAULT: <w:checked> (current state)
                        // and <w:default> (initial/reset state) are independent —
                        // surface them DISTINCTLY, each only when present (mirror the
                        // dropdown ffResult/ffDefault split, BUG-DUMP-R27-3). The old
                        // readback collapsed `checked ?? default` into one ffChecked
                        // and dropped <w:default>, so a checkbox whose default differed
                        // from its current state round-tripped with the default flipped
                        // and the explicit current marker lost or a spurious one added.
                        // (<w:checked/> with no w:val means checked=true.)
                        if (cChecked != null) node.Format["ffChecked"] = cChecked.Val?.Value ?? true;
                        if (cDefault != null) node.Format["ffDefault"] = cDefault.Val?.Value ?? true;
                        var cSize = cb.GetFirstChild<FormFieldSize>()?.Val?.Value;
                        if (!string.IsNullOrEmpty(cSize))
                            node.Format["ffCheckBoxSize"] = cSize;
                    }
                    else if (dd != null)
                    {
                        node.Format["ffType"] = "dropdown";
                        var items = dd.Elements<ListEntryFormField>()
                            .Select(li => li.Val?.Value ?? "").ToList();
                        if (items.Count > 0)
                            node.Format["ffItems"] = string.Join(",", items);
                        // BUG-DUMP-R27-3: <w:ddList> carries TWO distinct indices —
                        // <w:result> (DropDownListSelection) is the CURRENT
                        // selection, <w:default> (DefaultDropDownListItemIndex) is the
                        // default entry. The old readback stored the SELECTION
                        // under `ffDefault` (conflating the two) and never read
                        // the real default, so on emit neither <w:result> nor
                        // <w:default> was re-applied — the dropdown reverted to
                        // the first entry. Surface them as separate keys.
                        var dSel = dd.GetFirstChild<DropDownListSelection>()?.Val?.Value;
                        if (dSel != null)
                            node.Format["ffResult"] = (int)dSel;
                        var dDef = dd.GetFirstChild<DefaultDropDownListItemIndex>()?.Val?.Value;
                        if (dDef != null)
                            node.Format["ffDefault"] = (int)dDef;
                    }
                }
            }
        }
        if (node.Type == "run")
        {
            var instrEl = run.GetFirstChild<FieldCode>();
            // BUG-DUMP-DELFIELD: a field instruction nested inside a <w:del>
            // wrapper carries its code in <w:delInstrText> (DeletedFieldCode),
            // not <w:instrText> (FieldCode). Without this fallback the deleted
            // field code stays an opaque "run" node, CollapseFieldChains never
            // sees the HYPERLINK/PAGE/… instruction, and the field collapses
            // to a bare `add r` — losing both the field structure AND the
            // enclosing deletion (tracked-deleted text resurrected as live).
            // Treat DeletedFieldCode identically to FieldCode so the collapse
            // pass parses the instruction; the run's revision.* attribution
            // (set above from the DeletedRun ancestor) rides along.
            var delInstrEl = instrEl == null ? run.GetFirstChild<DeletedFieldCode>() : null;
            if (instrEl != null)
            {
                node.Type = "instrText";
                node.Format["instruction"] = instrEl.Text ?? "";
                // CONSISTENCY(canonical-keys): also surface the
                // instruction as node.Text so selector text-contains
                // searches (`instrText[text~=PAGE]`) and Get readback
                // agree. Without this, MatchesRunSelector's
                // GetRunText fallback hits the <w:instrText> content
                // while Navigation hands callers an empty Text — the
                // two surfaces disagreed on what the run "says".
                node.Text = instrEl.Text ?? "";
            }
            else if (delInstrEl != null)
            {
                node.Type = "instrText";
                node.Format["instruction"] = delInstrEl.Text ?? "";
                node.Text = delInstrEl.Text ?? "";
            }
        }
        // CONSISTENCY(run-text-tab): the type-upgrade for tab/break runs
        // checks "no Text element" (not "node.Text empty") because
        // GetRunText now surfaces TabChar as \t in node.Text. A pure
        // <w:r><w:tab/></w:r> run has no <w:t> child but node.Text="\t".
        // BUG-DUMP-DELTAB: also exclude runs carrying <w:delText> (a tracked
        // deletion). GetRunText surfaces delText into node.Text; without this
        // exclusion a deleted run mixing <w:tab/> + <w:delText> was reclassified
        // tab-only and node.Text wiped to "", silently dropping the deleted text.
        if (node.Type == "run" && !run.Elements<Text>().Any() && !run.Elements<DeletedText>().Any())
        {
            var tabEls = run.Elements<TabChar>().ToList();
            // BUG-DUMP-R25-2: a tab-only run carrying MULTIPLE <w:tab/> chars
            // (no <w:t>) must keep its count. The single-`tab` upgrade below
            // sets node.Text="" and the emitter then hardcodes one "\t", so the
            // extra tabs were dropped on dump round-trip. Mirror the multi-break
            // case directly below: GetRunText already surfaced them as
            // node.Text="\t\t\t" (one \t per TabChar) and AddText splits run
            // text on \t back into one TabChar each — so for the multi-tab case
            // keep the run as a `run` and let the text path carry the count.
            // Only a SINGLE tab takes the type=tab upgrade.
            if (tabEls.Count == 1)
            {
                node.Type = "tab";
                node.Text = "";
            }
            // tabEls.Count > 1 falls through: node stays type="run" with
            // node.Text="\t\t…" (set by GetRunText), round-tripped via add r.
        }
        // CONSISTENCY(run-text-break): gate on "no Text element" (not
        // "node.Text empty"), same as the tab upgrade above. GetRunText
        // surfaces a soft <w:br/> (textWrapping) as \n in node.Text, so a
        // pure <w:r><w:br/></w:r> break run is no longer Text-empty and the
        // old IsNullOrEmpty gate skipped the upgrade — leaving it as a run
        // with text="\n" that the emitter mis-rendered. A mixed run
        // <w:t>foo</w:t><w:br/> still has a <w:t> child, so it stays a run
        // (text="foo\n") and the inline break is preserved as \n.
        if (node.Type == "run" && !run.Elements<Text>().Any() && !run.Elements<DeletedText>().Any())
        {
            var breakEl = run.GetFirstChild<Break>();
            if (breakEl != null)
            {
                // BUG-R10A(BUG2): a breaks-only run carrying MULTIPLE soft line
                // breaks (<w:br/><w:br/><w:br/> with no <w:t>) must keep its
                // count. The single-`break` upgrade below collapses the whole
                // run to one `add pagebreak type=line`, dropping the extra
                // breaks on dump round-trip. GetRunText already surfaced them as
                // node.Text="\n\n\n" (one \n per line break / CR), and
                // AppendTextWithBreaks rebuilds one <w:br/> per \n — so for the
                // multi-line-break case keep the run as a `run` and let the text
                // path carry the count. Only a SINGLE break (or any page/column
                // break, which has no \n source representation) takes the
                // type=break upgrade. A text-flanked run keeps its <w:t> child
                // and never reaches this branch.
                var lineBreakLikeCount =
                    run.Elements<Break>().Count(b =>
                        b.Type == null || b.Type.Value == BreakValues.TextWrapping)
                    + run.Elements<CarriageReturn>().Count();
                bool hasNonLineBreak = run.Elements<Break>().Any(b =>
                    b.Type != null && b.Type.Value != BreakValues.TextWrapping);
                if (lineBreakLikeCount > 1 && !hasNonLineBreak)
                {
                    // Leave node.Type == "run"; node.Text is already "\n…\n".
                }
                else
                {
                    node.Type = "break";
                    node.Text = "";
                    // Normalize "textWrapping" → "line" on emit. OOXML treats
                    // a typeless <w:br/> as textWrapping (the default), but
                    // AddBreak's user-facing vocab uses "line"; without
                    // normalisation, dump round-trip emits `type=line` from
                    // typeless source and `type=textWrapping` from the
                    // explicitly-stamped replay target — semantically
                    // identical, byte-different.
                    if (breakEl.Type?.HasValue == true)
                    {
                        var bt = breakEl.Type.InnerText;
                        node.Format["breakType"] = string.Equals(bt, "textWrapping", StringComparison.OrdinalIgnoreCase)
                            ? "line"
                            : bt;
                    }
                    // <w:br w:clear="all|left|right|none"/> — a text-wrapping
                    // break's float-clearing behavior. Dropping it left spacer
                    // lines beside a floating table instead of below it, and
                    // the layout below merged upward.
                    if (breakEl.Clear?.HasValue == true)
                        node.Format["breakClear"] = breakEl.Clear.InnerText;
                }
            }
        }

        if (run.Parent is Hyperlink hlParent)
        {
            // BUG-DUMP10-05: a hyperlink wrapper with neither r:id nor
            // anchor (tooltip-only / history-only) used to fall through
            // both branches below, leaving the run with no Format keys
            // that would trigger the WordBatchEmitter hyperlink-emit guard.
            // Surface a sentinel so the wrapper survives even when there
            // is no destination — required for w:hyperlink[@w:tooltip]
            // bookmarks-style hover popups.
            node.Format["isHyperlink"] = true;
            if (hlParent.Id?.Value != null)
            {
                try
                {
                    var rel = ResolveHyperlinkRelationship(hlParent, hlParent.Id.Value);
                    // CONSISTENCY(docx-hyperlink-canonical-url): schema docx/hyperlink.json
                    // declares `url` as the canonical key; `link` is accepted as an input
                    // alias by Add/Set but Get normalizes output to `url`.
                    // Use OriginalString rather than ToString() — System.Uri normalises
                    // a bare authority by appending a trailing `/` (e.g.
                    // `https://example.com` → `https://example.com/`), which would
                    // surface as a get-side drift the on-disk .rels Target lacks.
                    if (rel != null) node.Format["url"] = rel.Uri.OriginalString;
                }
                catch { }
            }
            // CONSISTENCY(internal-anchor-hyperlink): runs inside an
            // internal anchor hyperlink (w:hyperlink[@w:anchor]) had no
            // r:id, so `anchor` was never surfaced on the run. The
            // WordBatchEmitter hyperlink branch keys off Format["anchor"]/
            // ["url"] to emit `add hyperlink`; without anchor the run
            // was demoted to a plain `add r` and the link was lost on
            // dump→batch round-trip.
            if (hlParent.Anchor?.Value != null)
                node.Format["anchor"] = hlParent.Anchor.Value;
            // BUG-DUMP24-02: w:docLocation is a separate "location in
            // target document" attribute, distinct from w:anchor. Surface
            // it so dump→batch round-trips the wrapping hyperlink fully.
            if (hlParent.DocLocation?.Value != null)
                node.Format["docLocation"] = hlParent.DocLocation.Value;
            // BUG-DUMP10-02: surface the tooltip / tgtFrame / history
            // attributes from the wrapping hyperlink so dump→batch
            // round-trip preserves them. Same canonical keys as the
            // standalone Hyperlink branch below.
            if (hlParent.Tooltip?.Value != null)
                node.Format["tooltip"] = hlParent.Tooltip.Value;
            if (hlParent.TargetFrame?.Value != null)
                node.Format["tgtFrame"] = hlParent.TargetFrame.Value;
            // BUG-DUMP-HISTFALSE: OOXML default for w:history is true, so an
            // explicit w:history="false" must be surfaced — emitting only on
            // ==true dropped it and flipped the link to history-on on round-trip.
            if (hlParent.History?.Value is bool hlpHist)
                node.Format["history"] = hlpHist;
        }

        // Populate effective.* properties from style inheritance.
        // CONSISTENCY(run-special-content): runs whose primary payload
        // is a structural inline element (ptab/fieldChar/instrText/tab/
        // break) carry no glyph for font/size/color to apply to;
        // emitting effective.size / effective.font.* on them only
        // floods output with noise and primes audit tools to misread
        // cosmetic styles on a "fldChar end" marker as meaningful.
        // Picture/ole runs are gated for the same reason — their
        // typography is irrelevant to the embedded media.
        var parentPara = run.Ancestors<Paragraph>().FirstOrDefault();
        if (parentPara != null && node.Type == "run")
            PopulateEffectiveRunProperties(node, run, parentPara);

        // Same noise-suppression for direct rPr-level keys read before
        // the type upgrade above (font.*/size/bold/...): on a pure MARKER
        // run (fieldChar / instrText / break) the rPr has no glyph to paint,
        // so surfacing font/size/color is noise that primes audit tools to
        // misread cosmetic styling on a structural marker as meaningful —
        // strip it so the bag shows only the role-defining keys
        // (fieldCharType, instr, breakType, …).
        //
        // BUG-DUMP-TABRPR: tab and ptab are deliberately NOT in this list
        // (they were, historically, under the same "no glyph" rationale —
        // but that was wrong for them). A tab / positional tab is a SIZED
        // inline element: Word routinely stamps rPr on these runs (font/size/
        // szCs drive the line height and the leader-dot glyphs the tab
        // paints), and stripping it dropped that rPr on every dump→batch
        // round-trip, leaving an empty <w:rPr/>. The earlier strip never
        // needed them anyway — the tab/ptab assertions
        // (Req5Round12FuzzTabSpecialTests) only check a run shows no
        // typography AFTER Set rejected/never-applied any (so the run
        // genuinely has none to read); they don't require hiding typography
        // that a source document already authored.
        if (node.Type is "fieldChar" or "instrText" or "break")
        {
            // BUG-DUMP-R46-FFSIZE: a FORMCHECKBOX / FORMTEXT begin fieldChar
            // legitimately carries the field-run's typography — a <w:sizeAuto/>
            // checkbox sizes its glyph to the run's FONT SIZE, and AddFormField
            // stamps the dumped size/bold/color/font back onto every rebuilt
            // field run. Stripping it (as noise) shrank the dumped form field to
            // the docDefaults size, enlarging the checkbox and inflating each
            // list row. Keep the field-run formatting on a form-field begin node
            // (it rides the FieldResultFormatKeys channel into `add formfield`);
            // every other field marker still sheds its noise typography.
            bool keepFieldRunFmt = node.Format.ContainsKey("hasFormFieldData");
            foreach (var noiseKey in TypographyOnlyKeys)
            {
                if (keepFieldRunFmt && FieldRunFormatKeepKeys.Contains(noiseKey))
                    continue;
                node.Format.Remove(noiseKey);
            }
        }
        return node;
    }

    private DocumentNode HyperlinkToNode(Hyperlink hyperlink, DocumentNode node)
    {
        node.Type = "hyperlink";
        node.Text = string.Concat(hyperlink.Descendants<Text>().Select(t => t.Text));
        var relId = hyperlink.Id?.Value;
        if (relId != null)
        {
            try
            {
                var rel = ResolveHyperlinkRelationship(hyperlink, relId);
                // CONSISTENCY(docx-hyperlink-canonical-url): see note above.
                if (rel != null) node.Format["url"] = rel.Uri.OriginalString;
            }
            catch { }
        }
        // Internal-anchor hyperlink (`add --type hyperlink --prop anchor=Foo`)
        // sets w:hyperlink/@w:anchor instead of @r:id. Surface it so set/get
        // round-trips and users can debug why a link points where it does.
        if (hyperlink.Anchor?.Value != null)
            node.Format["anchor"] = hyperlink.Anchor.Value;
        // BUG-DUMP24-02: w:docLocation is a separate "location in target
        // document" attribute, distinct from w:anchor. Surface it so
        // dump→batch round-trips it.
        if (hyperlink.DocLocation?.Value != null)
            node.Format["docLocation"] = hyperlink.DocLocation.Value;
        // BUG-DUMP10-02: tooltip / tgtFrame / history attributes are
        // independent of url/anchor — surface them so dump→batch
        // preserves the hover popup, target window, and history flag.
        if (hyperlink.Tooltip?.Value != null)
            node.Format["tooltip"] = hyperlink.Tooltip.Value;
        if (hyperlink.TargetFrame?.Value != null)
            node.Format["tgtFrame"] = hyperlink.TargetFrame.Value;
        // BUG-DUMP-HISTFALSE: OOXML default for w:history is true, so an
        // explicit w:history="false" must be surfaced — emitting only on
        // ==true dropped it and flipped the link to history-on on round-trip.
        if (hyperlink.History?.Value is bool hlHist)
            node.Format["history"] = hlHist;
        // Read run formatting from the first run inside the hyperlink
        var hlRun = hyperlink.Elements<Run>().FirstOrDefault(r => r.GetFirstChild<Text>() != null);
        if (hlRun?.RunProperties != null)
        {
            var rp = hlRun.RunProperties;
            if (rp.RunFonts?.Ascii?.Value != null) node.Format["font"] = rp.RunFonts.Ascii.Value;
            // BUG-DUMP17-07: surface per-script font slot so dump→batch
            // round-trip preserves font.cs on hyperlink runs.
            if (rp.RunFonts?.ComplexScript?.Value != null) node.Format["font.cs"] = rp.RunFonts.ComplexScript.Value;
            if (rp.FontSize?.Val?.Value != null)
                node.Format["size"] = $"{int.Parse(rp.FontSize.Val.Value) / 2.0:0.##}pt";
            if (rp.Bold != null) node.Format["bold"] = IsToggleOn(rp.Bold);
            if (rp.Italic != null) node.Format["italic"] = IsToggleOn(rp.Italic);
            if (rp.Color?.ThemeColor?.HasValue == true) node.Format["color"] = rp.Color.ThemeColor.InnerText;
            else if (rp.Color?.Val?.Value != null) node.Format["color"] = ParseHelpers.FormatHexColor(rp.Color.Val.Value);
            if (rp.Underline?.Val != null) node.Format["underline"] = rp.Underline.Val.InnerText;
            // CONSISTENCY(underline-color): backfilled from style Get edc8f884.
            if (rp.Underline?.Color?.Value != null)
                node.Format["underline.color"] = ParseHelpers.FormatHexColor(rp.Underline.Color.Value);
            if (rp.Strike != null) node.Format["strike"] = IsToggleOn(rp.Strike);
            if (rp.Highlight?.Val != null) node.Format["highlight"] = rp.Highlight.Val.InnerText;
            var rStyle = rp.GetFirstChild<RunStyle>();
            if (rStyle?.Val?.Value != null)
                node.Format["rStyle"] = rStyle.Val.Value;
        }
        return node;
    }

    private DocumentNode TableToNode(Table table, DocumentNode node, string path, int depth)
    {
        node.Type = "table";
        var flatRows = GetTableRowsFlattened(table);
        node.ChildCount = flatRows.Count;
        var firstRow = flatRows.FirstOrDefault();
        // Use grid column count (from TableGrid) instead of cell count for accurate column reporting
        var gridColCount = table.GetFirstChild<TableGrid>()?.Elements<GridColumn>().Count();
        // CONSISTENCY(format-stringy): user-facing numeric counts are
        // stored as strings to match other Word format keys (size "14pt",
        // spacing "12pt"). Avoids object-vs-int comparison surprises.
        node.Format["cols"] = (gridColCount ?? (firstRow != null ? GetRowCellsFlattened(firstRow).Count : 0)).ToString();
        node.Format["rows"] = node.ChildCount.ToString();
        // _gridCols: actual <w:gridCol> count (0 when TableGrid is missing
        // or empty), unbiased by the row-cell fallback that `cols` uses for
        // backward-compat. EmitTable reads this to decide whether to emit
        // `gridCols=0` on the dumped `add table` so AddTable leaves the
        // <w:tblGrid/> empty — preserving sources whose cells encode width
        // via tcW (or auto-fit). Underscore-prefixed to mark it as
        // internal-only (not a user-facing Set/Add key).
        node.Format["_gridCols"] = (gridColCount ?? 0).ToString();

        // BUG-DUMP-R43-9: a <w:tblGridChange> records the prior column-grid as
        // a tracked change. It carries no author/date (schema has only w:id) and
        // no property-fold host, so unlike the other *PrChange markers it has no
        // attribution to re-stamp — capture the FULL marker verbatim (mirrors
        // cellMerge.xml) and let EmitTable raw-set it back into the tblGrid.
        // Without this, a standalone tblGridChange (one not produced as a side
        // effect of a width Set) was dropped entirely on round-trip.
        var gridChange = table.GetFirstChild<TableGrid>()?.GetFirstChild<TableGridChange>();
        if (gridChange != null)
            node.Format["tblGridChange.xml"] = gridChange.OuterXml;

        var tp = table.GetFirstChild<TableProperties>();
        if (tp != null)
        {
            // tblPrChange: `set table + trackChange.author` snapshots the
            // prior tblPr and stamps author/date. Surfaced under
            // tblPrChange.* so EmitTable can emit a follow-up
            // `set /body/tbl[N]` step carrying trackChange.author/date
            // (re-runs the snapshot+stamp on replay). Distinct namespace
            // from any future bare `trackChange=` on tables.
            var tblPrChange = tp.GetFirstChild<TablePropertiesChange>();
            if (tblPrChange != null)
            {
                if (!string.IsNullOrEmpty(tblPrChange.Author?.Value))
                    node.Format["tblPrChange.author"] = tblPrChange.Author!.Value!;
                if (tblPrChange.Date?.Value is DateTime tDate)
                    node.Format["tblPrChange.date"] = tDate.ToString("o");
                // BUG-DUMP-R43-9: carry the verbatim prior-tblPr snapshot.
                var tblPrev = tblPrChange.GetFirstChild<PreviousTableProperties>();
                if (tblPrev != null && tblPrev.HasChildren)
                    node.Format["tblPrChange.beforeXml"] = tblPrev.OuterXml;
            }
            // Table style
            // BUG-R3-05: empty Val (set via legacy code that wrote tblStyle
            // with empty string) must NOT surface as a "style" key.
            if (!string.IsNullOrEmpty(tp.TableStyle?.Val?.Value))
                node.Format["style"] = tp.TableStyle.Val.Value!;
            // BUG-DUMP-R36-2: tblStyleRowBandSize / tblStyleColBandSize control
            // how many rows/cols make up one stripe when a banded table style
            // applies (default 1). Previously dropped on round-trip — band size 2
            // stripes every 2nd row instead of every row, a visible change.
            // Walk children by local name rather than GetFirstChild<T>: in
            // CT_TblPr these elements precede tblW, and real-world producers that
            // emit them AFTER tblW make the strict SDK parser type them as
            // OpenXmlUnknownElement (the typed accessor then returns null).
            // Our own AddTable writes them inside an mc:AlternateContent
            // Requires="w" guard (CT_TblPr has no slot for them — see
            // AddTable's band-guard comment), so additionally unwrap one
            // level of direct AlternateContent/Choice. Scoped unwrap, not
            // Descendants(): a tblPrChange snapshot must not leak its
            // prior band sizes onto the live table.
            foreach (var tpChild in tp.ChildElements.SelectMany(c =>
                         c is AlternateContent
                             ? c.ChildElements.OfType<AlternateContentChoice>()
                                .SelectMany(ch => ch.ChildElements)
                             : new[] { c }.AsEnumerable()))
            {
                var ln = tpChild.LocalName;
                if (ln is not ("tblStyleRowBandSize" or "tblStyleColBandSize")) continue;
                string? valAttr = null;
                foreach (var a in tpChild.GetAttributes())
                    if (a.LocalName == "val") { valAttr = a.Value; break; }
                if (!int.TryParse(valAttr, out var bandVal)) continue;
                if (ln == "tblStyleRowBandSize") node.Format["rowBandSize"] = bandVal;
                else node.Format["colBandSize"] = bandVal;
            }
            // Table borders. `LeftBorder`/`RightBorder` only catch
            // <w:left>/<w:right>; bidi-aware sources use <w:start>/<w:end>
            // which the SDK does NOT alias onto Left/Right (the typed
            // properties stay null). Walk all border children by local
            // name and map both forms onto the same canonical key — the
            // alternative is dropping borders for any doc whose tblBorders
            // uses the start/end naming (three-line-table2.docx).
            var tblBorders = tp.TableBorders;
            if (tblBorders != null)
            {
                ReadBorder(tblBorders.TopBorder, "border.top", node);
                ReadBorder(tblBorders.BottomBorder, "border.bottom", node);
                ReadBorder(tblBorders.InsideHorizontalBorder, "border.insideH", node);
                ReadBorder(tblBorders.InsideVerticalBorder, "border.insideV", node);
                foreach (var bChild in tblBorders.ChildElements)
                {
                    if (bChild is BorderType bt)
                    {
                        var ln = bChild.LocalName;
                        if (ln.Equals("left", StringComparison.OrdinalIgnoreCase)
                            || ln.Equals("start", StringComparison.OrdinalIgnoreCase))
                            ReadBorder(bt, "border.left", node);
                        else if (ln.Equals("right", StringComparison.OrdinalIgnoreCase)
                                 || ln.Equals("end", StringComparison.OrdinalIgnoreCase))
                            ReadBorder(bt, "border.right", node);
                    }
                }
            }
            // Table width
            // BUG-R4B(BUG1): read raw InnerText (decimal-tolerant) instead of
            // .Value, which throws on producers that emit w:w="9440.0".
            if (SafeWidth(tp.TableWidth?.Width) is int twWidth)
            {
                var wType = tp.TableWidth!.Type?.Value;
                // BUG-DUMP19-03: type=auto must round-trip as "auto", not
                // collapse to a bare dxa integer (Width="0").
                node.Format["width"] = wType == TableWidthUnitValues.Pct
                    ? FormatPctWidth(twWidth)
                    : wType == TableWidthUnitValues.Auto
                        ? "auto"
                        : twWidth.ToString(System.Globalization.CultureInfo.InvariantCulture);
            }
            else if (tp.TableWidth?.Type?.Value == TableWidthUnitValues.Auto)
            {
                // Some producers emit <w:tblW w:type="auto"/> without w:w.
                node.Format["width"] = "auto";
            }
            else
            {
                // Internal-only marker: source had no <w:tblW> element at
                // all. EmitTable reads this to tell AddTable to skip the
                // default-tblW stamp; without it, replay grows
                // a <w:tblW w:w="<sum-of-gridCol>" w:type="dxa"/> that
                // the source never had, and the next dump surfaces a
                // phantom `width=…` key.
                node.Format["_noTblW"] = true;
            }
            // Alignment
            if (tp.TableJustification?.Val?.Value != null)
                node.Format["align"] = tp.TableJustification.Val.InnerText;
            // Indent
            // BUG-R4B(BUG1): decimal-tolerant width read (w:tblInd w:w="0.0").
            // BUG-DUMP-R34-TBLIND: preserve the indent UNIT. A pct-typed tblInd
            // (w:type="pct", value in fiftieths-of-a-percent) was read as a bare
            // int and re-added as dxa twips, so a "2%" table indent (≈180 twips on
            // a letter page) collapsed to 100 twips — shifting the whole table
            // (and its bordered answer boxes) left. Encode pct as the same "X%"
            // form table width uses, so Add/Set re-parse the unit.
            if (SafeWidth(tp.TableIndentation?.Width) is int tblIndW)
                node.Format["indent"] = tp.TableIndentation!.Type?.Value == TableWidthUnitValues.Pct
                    ? FormatPctWidth(tblIndW)
                    : (object)tblIndW;
            // Cell spacing
            if (SafeWidth(tp.TableCellSpacing?.Width) is int tblCsW)
                node.Format["cellSpacing"] = tblCsW;
            // Layout — emit "autofit" (not "auto") so the readback token
            // matches the canonical input vocabulary documented in the
            // table add/set help. Set accepts both "auto" and "autofit"
            // (anything not "fixed" maps to Autofit), so this only affects
            // get and is round-trip safe with the dump/replay pipeline.
            if (tp.TableLayout?.Type?.Value != null)
                node.Format["layout"] = tp.TableLayout.Type.Value == TableLayoutValues.Fixed ? "fixed" : "autofit";
            // BUG-DUMP-R40-5: <w:tblLook> conditional-formatting bitmask
            // (firstRow/lastRow/firstColumn/lastColumn/noHBand/noVBand) controls
            // which banded/conditional table-style facets apply. Previously
            // dropped on dump→batch, so AddTable's default seed (04A0 = firstRow +
            // firstColumn) leaked onto every table — wrongly enabling first-column
            // bold/shading under conditional styles. The combined w:val hex is the
            // authoritative form (Word reads both w:val and the decomposed boolean
            // attrs, but w:val wins). Surface the verbatim hex so EmitTable/AddTable
            // round-trip every bit. Walk children by local name (CT_TblPr orders
            // tblLook last; a producer emitting attrs the SDK can't type still
            // exposes them via GetAttributes). If the source had NO tblLook, emit
            // no key — AddTable then suppresses the default seed.
            foreach (var tpChild in tp.ChildElements)
            {
                if (tpChild.LocalName != "tblLook") continue;
                string? lookVal = null;
                foreach (var a in tpChild.GetAttributes())
                    if (a.LocalName == "val") { lookVal = a.Value; break; }
                if (!string.IsNullOrEmpty(lookVal))
                {
                    // Normalize to 4-digit uppercase hex (Word writes "0620").
                    if (int.TryParse(lookVal, System.Globalization.NumberStyles.HexNumber,
                            System.Globalization.CultureInfo.InvariantCulture, out var lookBits))
                        node.Format["tblLook"] = lookBits.ToString("X4", System.Globalization.CultureInfo.InvariantCulture);
                    else
                        node.Format["tblLook"] = lookVal;
                }
                else
                {
                    // No w:val — reconstruct the bitmask from the decomposed
                    // boolean attributes so the round-trip is still lossless.
                    int bits = 0;
                    foreach (var a in tpChild.GetAttributes())
                    {
                        bool on = a.Value == "1" || a.Value == "true";
                        if (!on) continue;
                        switch (a.LocalName)
                        {
                            case "firstRow": bits |= 0x0020; break;
                            case "lastRow": bits |= 0x0040; break;
                            case "firstColumn": bits |= 0x0080; break;
                            case "lastColumn": bits |= 0x0100; break;
                            case "noHBand": bits |= 0x0200; break;
                            case "noVBand": bits |= 0x0400; break;
                        }
                    }
                    node.Format["tblLook"] = bits.ToString("X4", System.Globalization.CultureInfo.InvariantCulture);
                }
                break;
            }
            // Direction (CT_TblPrBase / w:bidiVisual). Mirrors paragraph
            // direction vocabulary. bidiVisual is a CT_OnOff toggle: present
            // with no val (or val=true) is ON (RTL); val="0"/false is an
            // explicit OFF (LTR). Read the val — a presence-only check turned a
            // source's explicit `<w:bidiVisual w:val="0"/>` into direction=rtl,
            // and AddTable then stamped a bare (ON) `<w:bidiVisual/>`, visually
            // mirroring the columns. Emit rtl only when the toggle is actually
            // ON; an OFF/absent toggle leaves no key, and WordBatchEmitter
            // pins direction=ltr for that case (see EmitTable).
            var tblBidi = tp.GetFirstChild<BiDiVisual>();
            if (tblBidi != null)
            {
                // CT_OnOff: no val attribute (or a truthy val) is ON; an
                // explicit falsey val ("0"/"false"/"off") is OFF. Read the raw
                // attribute text so the check is robust regardless of how the
                // SDK surfaces the toggle's typed value.
                var bidiRaw = tblBidi.Val?.InnerText;
                bool bidiOn = bidiRaw is null
                    || !(bidiRaw is "0" or "false" or "off");
                if (bidiOn)
                    node.Format["direction"] = "rtl";
            }
            // Default cell margin (padding)
            var dcm = tp.TableCellMarginDefault;
            // BUG-R4B(BUG1): decimal-tolerant margin reads.
            if (SafeWidth(dcm?.TopMargin?.Width) is int dcmT)
                node.Format["padding.top"] = dcmT;
            if (SafeWidth(dcm?.BottomMargin?.Width) is int dcmB)
                node.Format["padding.bottom"] = dcmB;
            if (SafeWidth(dcm?.TableCellLeftMargin?.Width) is int dcmL)
                node.Format["padding.left"] = dcmL;
            if (SafeWidth(dcm?.TableCellRightMargin?.Width) is int dcmR)
                node.Format["padding.right"] = dcmR;
            // Table-level shading (w:tblPr/w:shd). Mirror paragraph shading
            // pattern: split into shading.val/.fill/.color sub-keys.
            // WordBatchEmitter's shading-fold collapses these into a single
            // semicolon-encoded `shading=VAL;FILL[;COLOR]` value, which
            // AddTable consumes via the existing "shading" case.
            // BUG-DUMP22-09: floating-table position (<w:tblpPr/>) and
            // overlap (<w:tblOverlap/>) — both were silently dropped on
            // dump, leaving floating tables stuck inline on round-trip.
            // Surface tblpPr's six attrs as tblp.* dotted keys (using the
            // OOXML attribute local names verbatim) plus tblOverlap as a
            // dotted sibling so AddTable's TypedAttributeFallback can
            // re-create the elements verbatim. CONSISTENCY(canonical-keys):
            // dotted-segment-as-element-prefix matches ind.firstLine and
            // pBdr.top patterns.
            var tblpPr = tp.GetFirstChild<TablePositionProperties>();
            if (tblpPr != null)
            {
                if (tblpPr.HorizontalAnchor?.HasValue == true)
                    node.Format["tblp.horzAnchor"] = tblpPr.HorizontalAnchor.InnerText;
                if (tblpPr.VerticalAnchor?.HasValue == true)
                    node.Format["tblp.vertAnchor"] = tblpPr.VerticalAnchor.InnerText;
                if (tblpPr.TablePositionX?.HasValue == true)
                    node.Format["tblp.tblpX"] = tblpPr.TablePositionX.Value!;
                if (tblpPr.TablePositionY?.HasValue == true)
                    node.Format["tblp.tblpY"] = tblpPr.TablePositionY.Value!;
                if (tblpPr.TablePositionXAlignment?.HasValue == true)
                    node.Format["tblp.tblpXSpec"] = tblpPr.TablePositionXAlignment.InnerText;
                if (tblpPr.TablePositionYAlignment?.HasValue == true)
                    node.Format["tblp.tblpYSpec"] = tblpPr.TablePositionYAlignment.InnerText;
                if (tblpPr.LeftFromText?.HasValue == true)
                    node.Format["tblp.leftFromText"] = tblpPr.LeftFromText.Value!;
                if (tblpPr.RightFromText?.HasValue == true)
                    node.Format["tblp.rightFromText"] = tblpPr.RightFromText.Value!;
                if (tblpPr.TopFromText?.HasValue == true)
                    node.Format["tblp.topFromText"] = tblpPr.TopFromText.Value!;
                if (tblpPr.BottomFromText?.HasValue == true)
                    node.Format["tblp.bottomFromText"] = tblpPr.BottomFromText.Value!;
            }
            var tblOverlap = tp.GetFirstChild<TableOverlap>();
            if (tblOverlap?.Val?.HasValue == true)
                node.Format["tblOverlap.val"] = tblOverlap.Val.InnerText;
            if (tp.Shading != null)
            {
                var tShdVal = tp.Shading.Val?.InnerText;
                var tShdFill = tp.Shading.Fill?.Value;
                var tShdColor = tp.Shading.Color?.Value;
                if (!string.IsNullOrEmpty(tShdVal)) node.Format["shading.val"] = tShdVal;
                if (!string.IsNullOrEmpty(tShdFill)) node.Format["shading.fill"] = ParseHelpers.FormatHexColor(tShdFill);
                if (!string.IsNullOrEmpty(tShdColor)) node.Format["shading.color"] = ParseHelpers.FormatHexColor(tShdColor);
                ReadShadingTheme(tp.Shading, node);
            }

            // BUG-R3-01: tblLook readback — Set wrote the XML correctly, but
            // Get never read it back (Set/Get round-trip gap). Emit both the
            // short-form lowercase keys (firstrow/lastrow/bandrow — match
            // Set's case-insensitive vocabulary and project canonical
            // pattern: vmerge/colspan) AND OOXML-attribute-name camelCase
            // keys (firstRow/bandedRows — verbatim attribute names) so
            // batch round-trip works either way. The two forms exist for
            // historical-vocabulary parity; values are kept consistent
            // across both keys (lowercase stores "true"/"false" string,
            // camelCase stores bool).
            // BUG-R4-01/06: Get emits ONLY canonical camelCase keys
            // (firstRow/lastRow/firstCol/lastCol/bandedRows/bandedCols).
            // Set still accepts lowercase aliases (firstrow/bandrow/etc)
            // as input — see Set.Element.cs. Internal hex `tblLook.val`
            // is NOT surfaced (was a dump-poisoning impl detail).
            var tblLookRead = tp.GetFirstChild<TableLook>();
            if (tblLookRead != null)
            {
                if (tblLookRead.FirstRow?.HasValue == true && tblLookRead.FirstRow.Value)
                    node.Format["firstRow"] = true;
                if (tblLookRead.LastRow?.HasValue == true && tblLookRead.LastRow.Value)
                    node.Format["lastRow"] = true;
                if (tblLookRead.FirstColumn?.HasValue == true && tblLookRead.FirstColumn.Value)
                    node.Format["firstCol"] = true;
                if (tblLookRead.LastColumn?.HasValue == true && tblLookRead.LastColumn.Value)
                    node.Format["lastCol"] = true;
                // banding semantics are inverted: noHBand=true means NO banding.
                // Emit only when banding IS active (noHBand=false explicitly set).
                if (tblLookRead.NoHorizontalBand?.HasValue == true && !tblLookRead.NoHorizontalBand.Value)
                    node.Format["bandedRows"] = true;
                if (tblLookRead.NoVerticalBand?.HasValue == true && !tblLookRead.NoVerticalBand.Value)
                    node.Format["bandedCols"] = true;
            }

            // Accessibility: table caption / description. Set writes
            // <w:tblCaption w:val="…"/> and <w:tblDescription w:val="…"/>
            // (see Set.Element.cs table branch). Without the readback,
            // get/dump silently drops these on round-trip.
            var tblCaption = tp.GetFirstChild<TableCaption>();
            if (!string.IsNullOrEmpty(tblCaption?.Val?.Value))
                node.Format["caption"] = tblCaption.Val.Value!;
            var tblDescription = tp.GetFirstChild<TableDescription>();
            if (!string.IsNullOrEmpty(tblDescription?.Val?.Value))
                node.Format["description"] = tblDescription.Val.Value!;
        }

        // Column widths from grid
        var gridCols = table.GetFirstChild<TableGrid>()?.Elements<GridColumn>().ToList();
        if (gridCols != null && gridCols.Count > 0)
            node.Format["colWidths"] = string.Join(",", gridCols.Select(g => (g.Width?.Value ?? "0") + "dxa"));

        if (depth > 0)
        {
            int rowIdx = 0;
            foreach (var row in GetTableRowsFlattened(table))
            {
                var rowNode = new DocumentNode
                {
                    Path = $"{path}/tr[{rowIdx + 1}]",
                    Type = "row",
                    ChildCount = GetRowCellsFlattened(row).Count
                };
                ReadRowProps(row, rowNode);
                if (depth > 1)
                {
                    int cellIdx = 0;
                    foreach (var cell in GetRowCellsFlattened(row))
                    {
                        var cellNode = new DocumentNode
                        {
                            Path = $"{path}/tr[{rowIdx + 1}]/tc[{cellIdx + 1}]",
                            Type = "cell",
                            Text = string.Join("", cell.Descendants<Text>().Select(t => t.Text)),
                            // CONSISTENCY(cell-children): include nested Table and
                            // block-SDT children alongside Paragraphs. BUG-R11A(BUG1).
                            ChildCount = cell.Elements<OpenXmlElement>().Count(e => e is Paragraph || e is Table || e is SdtBlock)
                        };
                        ReadCellProps(cell, cellNode);
                        if (depth > 2)
                        {
                            int cellPIdx = 0, cellTblIdx = 0, cellSdtIdx = 0;
                            foreach (var cellChild in cell.Elements<OpenXmlElement>())
                            {
                                if (cellChild is Paragraph cellPara)
                                {
                                    cellPIdx++;
                                    var cParaSegment = BuildParaPathSegment(cellPara, cellPIdx);
                                    cellNode.Children.Add(ElementToNode(cellPara, $"{path}/tr[{rowIdx + 1}]/tc[{cellIdx + 1}]/{cParaSegment}", depth - 3));
                                }
                                else if (cellChild is Table cellTbl)
                                {
                                    cellTblIdx++;
                                    cellNode.Children.Add(ElementToNode(cellTbl, $"{path}/tr[{rowIdx + 1}]/tc[{cellIdx + 1}]/tbl[{cellTblIdx}]", depth - 3));
                                }
                                else if (cellChild is SdtBlock cellSdt)
                                {
                                    cellSdtIdx++;
                                    cellNode.Children.Add(ElementToNode(cellSdt, $"{path}/tr[{rowIdx + 1}]/tc[{cellIdx + 1}]/{BuildSdtPathSegment(cellSdt, cellSdtIdx)}", depth - 3));
                                }
                            }
                        }
                        rowNode.Children.Add(cellNode);
                        cellIdx++;
                    }
                }
                node.Children.Add(rowNode);
                rowIdx++;
            }
        }
        return node;
    }

    private DocumentNode TableCellToNode(TableCell directCell, DocumentNode node, string path, int depth)
    {
        node.Type = "cell";
        node.Text = string.Join("", directCell.Descendants<Text>().Select(t => t.Text));
        // CONSISTENCY(cell-children): include nested Table and block-SDT children
        // alongside Paragraphs. BUG-R11A(BUG1): without SdtBlock here a block
        // content control that is a direct cell child was invisible to Get (and
        // therefore dropped by the dump cell walk that reads cellNode.Children).
        node.ChildCount = directCell.Elements<OpenXmlElement>().Count(e => e is Paragraph || e is Table || e is SdtBlock);
        ReadCellProps(directCell, node);
        if (depth > 0)
        {
            int dcPIdx = 0, dcTblIdx = 0, dcSdtIdx = 0;
            foreach (var dcChild in directCell.Elements<OpenXmlElement>())
            {
                if (dcChild is Paragraph cellPara)
                {
                    dcPIdx++;
                    var dcParaSegment = BuildParaPathSegment(cellPara, dcPIdx);
                    node.Children.Add(ElementToNode(cellPara, $"{path}/{dcParaSegment}", depth - 1));
                }
                else if (dcChild is Table dcTbl)
                {
                    dcTblIdx++;
                    node.Children.Add(ElementToNode(dcTbl, $"{path}/tbl[{dcTblIdx}]", depth - 1));
                }
                else if (dcChild is SdtBlock dcSdt)
                {
                    dcSdtIdx++;
                    node.Children.Add(ElementToNode(dcSdt, $"{path}/{BuildSdtPathSegment(dcSdt, dcSdtIdx)}", depth - 1));
                }
            }
        }
        return node;
    }

    private DocumentNode TableRowToNode(TableRow directRow, DocumentNode node, string path, int depth)
    {
        node.Type = "row";
        node.ChildCount = GetRowCellsFlattened(directRow).Count;
        ReadRowProps(directRow, node);
        if (depth > 0)
        {
            int cellIdx = 0;
            foreach (var cell in GetRowCellsFlattened(directRow))
            {
                var cellNode = new DocumentNode
                {
                    Path = $"{path}/tc[{cellIdx + 1}]",
                    Type = "cell",
                    Text = string.Join("", cell.Descendants<Text>().Select(t => t.Text)),
                    // CONSISTENCY(cell-children): include nested Table and block-SDT
                    // children alongside Paragraphs. BUG-R11A(BUG1).
                    ChildCount = cell.Elements<OpenXmlElement>().Count(e => e is Paragraph || e is Table || e is SdtBlock)
                };
                ReadCellProps(cell, cellNode);
                if (depth > 1)
                {
                    int drPIdx = 0, drTblIdx = 0, drSdtIdx = 0;
                    foreach (var drChild in cell.Elements<OpenXmlElement>())
                    {
                        if (drChild is Paragraph cellPara)
                        {
                            drPIdx++;
                            var drParaSegment = BuildParaPathSegment(cellPara, drPIdx);
                            cellNode.Children.Add(ElementToNode(cellPara, $"{path}/tc[{cellIdx + 1}]/{drParaSegment}", depth - 2));
                        }
                        else if (drChild is Table drTbl)
                        {
                            drTblIdx++;
                            cellNode.Children.Add(ElementToNode(drTbl, $"{path}/tc[{cellIdx + 1}]/tbl[{drTblIdx}]", depth - 2));
                        }
                        else if (drChild is SdtBlock drSdt)
                        {
                            drSdtIdx++;
                            cellNode.Children.Add(ElementToNode(drSdt, $"{path}/tc[{cellIdx + 1}]/{BuildSdtPathSegment(drSdt, drSdtIdx)}", depth - 2));
                        }
                    }
                }
                node.Children.Add(cellNode);
                cellIdx++;
            }
        }
        return node;
    }

    private DocumentNode SdtBlockToNode(SdtBlock sdtBlockNode, DocumentNode node)
    {
        node.Type = "sdt";
        var sdtProps = sdtBlockNode.SdtProperties;
        if (sdtProps != null)
        {
            var alias = sdtProps.GetFirstChild<SdtAlias>();
            if (alias?.Val?.Value != null) node.Format["alias"] = alias.Val.Value;
            var tagEl = sdtProps.GetFirstChild<Tag>();
            if (tagEl?.Val?.Value != null) node.Format["tag"] = tagEl.Val.Value;
            var lockEl = sdtProps.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Lock>();
            if (lockEl?.Val?.Value != null) node.Format["lock"] = lockEl.Val.InnerText;
            var sdtId = sdtProps.GetFirstChild<SdtId>();
            if (sdtId?.Val?.Value != null) node.Format["id"] = sdtId.Val.Value;

            // Determine SDT type (check specific types first, text last as fallback)
            // BUG-DUMP-R42-7/8: <w:group/> and <w:picture/> markers identify a
            // grouping / picture content control; without reading them the
            // control was reported (and later rebuilt) as a generic rich-text SDT.
            var checkBoxEl = sdtProps.GetFirstChild<DocumentFormat.OpenXml.Office2010.Word.SdtContentCheckBox>();
            if (sdtProps.GetFirstChild<SdtContentGroup>() != null) node.Format["type"] = "group";
            else if (sdtProps.GetFirstChild<SdtContentPicture>() != null) node.Format["type"] = "picture";
            else if (checkBoxEl != null) node.Format["type"] = "checkbox";
            else if (sdtProps.GetFirstChild<SdtContentDropDownList>() != null) node.Format["type"] = "dropdown";
            else if (sdtProps.GetFirstChild<SdtContentComboBox>() != null) node.Format["type"] = "combobox";
            else if (sdtProps.GetFirstChild<SdtContentDate>() != null) node.Format["type"] = "date";
            else if (sdtProps.GetFirstChild<SdtContentText>() != null) node.Format["type"] = "text";
            else node.Format["type"] = "richtext";

            // Checkbox checked state (w14:checked val 1/0 → bool).
            if (checkBoxEl != null)
                node.Format["checked"] = checkBoxEl.Checked?.Val?.InnerText == "1"
                    || string.Equals(checkBoxEl.Checked?.Val?.InnerText, "true", StringComparison.OrdinalIgnoreCase);

            // Read date format for date controls
            var dateContent = sdtProps.GetFirstChild<SdtContentDate>();
            if (dateContent?.DateFormat?.Val?.Value != null)
                node.Format["format"] = dateContent.DateFormat.Val.Value;

            // Editable status
            node.Format["editable"] = IsSdtEditable(sdtProps);

            // Placeholder detection. `placeholder` (showingPlcHdr flag) and
            // `placeholderText` (docPart reference) are INDEPENDENT in OOXML: a
            // control can declare a placeholder gallery while displaying real
            // content (showingPlcHdr absent). Surface each on its own so the
            // docPart reference round-trips even when not currently shown.
            if (sdtProps.GetFirstChild<ShowingPlaceholder>() != null)
                node.Format["placeholder"] = true;
            var plcHdrText = sdtProps.GetFirstChild<SdtPlaceholder>()?.DocPartReference?.Val?.Value;
            if (plcHdrText != null) node.Format["placeholderText"] = plcHdrText;

            ReadSdtExtraProps(sdtProps, node);
        }
        node.Text = string.Concat(sdtBlockNode.Descendants<Text>().Select(t => t.Text));
        var sdtContent = sdtBlockNode.SdtContentBlock;
        node.ChildCount = sdtContent?.ChildElements.Count ?? 0;
        return node;
    }

    private DocumentNode SdtRunToNode(SdtRun sdtRunNode, DocumentNode node)
    {
        node.Type = "sdt";
        var sdtProps = sdtRunNode.SdtProperties;
        if (sdtProps != null)
        {
            var alias = sdtProps.GetFirstChild<SdtAlias>();
            if (alias?.Val?.Value != null) node.Format["alias"] = alias.Val.Value;
            var tagEl = sdtProps.GetFirstChild<Tag>();
            if (tagEl?.Val?.Value != null) node.Format["tag"] = tagEl.Val.Value;
            var lockEl = sdtProps.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Lock>();
            if (lockEl?.Val?.Value != null) node.Format["lock"] = lockEl.Val.InnerText;
            var sdtId = sdtProps.GetFirstChild<SdtId>();
            if (sdtId?.Val?.Value != null) node.Format["id"] = sdtId.Val.Value;

            // BUG-DUMP-R42-7/8: surface group / picture content-control markers.
            var checkBoxElRun = sdtProps.GetFirstChild<DocumentFormat.OpenXml.Office2010.Word.SdtContentCheckBox>();
            if (sdtProps.GetFirstChild<SdtContentGroup>() != null) node.Format["type"] = "group";
            else if (sdtProps.GetFirstChild<SdtContentPicture>() != null) node.Format["type"] = "picture";
            else if (checkBoxElRun != null) node.Format["type"] = "checkbox";
            else if (sdtProps.GetFirstChild<SdtContentDropDownList>() != null) node.Format["type"] = "dropdown";
            else if (sdtProps.GetFirstChild<SdtContentComboBox>() != null) node.Format["type"] = "combobox";
            else if (sdtProps.GetFirstChild<SdtContentDate>() != null) node.Format["type"] = "date";
            else if (sdtProps.GetFirstChild<SdtContentText>() != null) node.Format["type"] = "text";
            else node.Format["type"] = "richtext";

            if (checkBoxElRun != null)
                node.Format["checked"] = checkBoxElRun.Checked?.Val?.InnerText == "1"
                    || string.Equals(checkBoxElRun.Checked?.Val?.InnerText, "true", StringComparison.OrdinalIgnoreCase);

            // Read date format for date controls
            var dateContentRun = sdtProps.GetFirstChild<SdtContentDate>();
            if (dateContentRun?.DateFormat?.Val?.Value != null)
                node.Format["format"] = dateContentRun.DateFormat.Val.Value;

            // Editable status
            node.Format["editable"] = IsSdtEditable(sdtProps);

            // Placeholder detection — `placeholder` (showingPlcHdr) and
            // `placeholderText` (docPart) are independent (see SdtBlockToNode).
            if (sdtProps.GetFirstChild<ShowingPlaceholder>() != null)
                node.Format["placeholder"] = true;
            var plcHdrTextRun = sdtProps.GetFirstChild<SdtPlaceholder>()?.DocPartReference?.Val?.Value;
            if (plcHdrTextRun != null) node.Format["placeholderText"] = plcHdrTextRun;

            ReadSdtExtraProps(sdtProps, node);
        }
        node.Text = string.Concat(sdtRunNode.Descendants<Text>().Select(t => t.Text));
        return node;
    }

    // BUG-DUMP-SDTPROPS: read the SDT sdtPr children that the typed dump→batch
    // path previously dropped — list items, date-picker selected value/calendar/
    // language/store-as, and combo/dropdown current selection (lastValue).
    // Shared by SdtBlockToNode and SdtRunToNode so block and inline controls
    // surface the identical canonical keys.
    private static void ReadSdtExtraProps(SdtProperties sdtProps, DocumentNode node)
    {
        // Date-picker: surface the actual selected value + locale/calendar so a
        // populated date control round-trips, not just its display format.
        var date = sdtProps.GetFirstChild<SdtContentDate>();
        if (date != null)
        {
            if (date.FullDate?.Value != null)
                node.Format["date.fullDate"] = date.FullDate.Value.ToString("yyyy-MM-ddTHH:mm:ssZ");
            if (date.Calendar?.Val != null && date.Calendar.Val.HasValue)
                node.Format["date.calendar"] = date.Calendar.Val.InnerText;
            if (date.LanguageId?.Val?.Value != null)
                node.Format["date.lid"] = date.LanguageId.Val.Value;
            if (date.SdtDateMappingType?.Val != null && date.SdtDateMappingType.Val.HasValue)
                node.Format["date.storeMappedDataAs"] = date.SdtDateMappingType.Val.InnerText;
        }

        // Dropdown / combo: list items + current selection (lastValue).
        var ddl = sdtProps.GetFirstChild<SdtContentDropDownList>();
        var combo = sdtProps.GetFirstChild<SdtContentComboBox>();
        var listItems = ddl?.Elements<ListItem>() ?? combo?.Elements<ListItem>();
        if (listItems != null)
        {
            // BUG-R5-07: SDT ListItems carry distinct DisplayText and
            // Value attrs. Real Word docs commonly differ (e.g.
            // "Draft|DRAFT"). Emit the pipe form when value !=
            // displayText so dump→add round-trips. ParseSdtItems on
            // the Add side accepts both bare and piped forms.
            var items = listItems.Select(li =>
            {
                var disp = li.DisplayText?.Value ?? li.Value?.Value ?? "";
                var val = li.Value?.Value ?? li.DisplayText?.Value ?? "";
                return disp == val ? disp : $"{disp}|{val}";
            }).ToList();
            if (items.Count > 0) node.Format["items"] = string.Join(",", items);
        }
        if (ddl?.LastValue?.Value is { Length: > 0 } ddlLast)
            node.Format["dropDown.lastValue"] = ddlLast;
        if (combo?.LastValue?.Value is { Length: > 0 } comboLast)
            node.Format["comboBox.lastValue"] = comboLast;

        // BUG-DUMP-R25-5: <w:dataBinding> links the control to a customXml data
        // store (xpath + storeItemID, plus the namespace prefixMappings the
        // xpath uses). Dropping it degrades a bound control into a static one.
        // Surface the three attrs so AddSdt can rebuild the element. Mirrors
        // how the date/placeholder sdtPr children are surfaced above.
        var dataBinding = sdtProps.GetFirstChild<DataBinding>();
        if (dataBinding != null)
        {
            if (dataBinding.XPath?.Value is { Length: > 0 } xp)
                node.Format["dataBinding.xpath"] = xp;
            if (dataBinding.StoreItemId?.Value is { Length: > 0 } sid)
                node.Format["dataBinding.storeItemID"] = sid;
            if (dataBinding.PrefixMappings?.Value is { Length: > 0 } pm)
                node.Format["dataBinding.prefixMappings"] = pm;
        }
    }

    private DocumentNode OfficeMathToNode(M.OfficeMath inlineMath, DocumentNode node)
    {
        node.Type = "equation";
        node.Format["mode"] = "inline";
        try { node.Text = Core.FormulaParser.ToLatex(inlineMath); }
        catch { node.Text = inlineMath.InnerText; }
        if (string.IsNullOrEmpty(node.Text))
            node.Text = inlineMath.InnerText;
        return node;
    }

    private DocumentNode HeaderFooterToNode(OpenXmlElement element, DocumentNode node, string path, int depth)
    {
        // Header/Footer: enumerate block-level children. Tables are valid
        // block-level OOXML inside hdr/ftr (same schema as body), so list
        // them alongside paragraphs. Mirrors body-listing logic above.
        node.Type = element is Header ? "header" : "footer";
        node.Text = string.Concat(element.Descendants<Text>().Select(t => t.Text));
        // BUG-R11A(BUG3): include block-SDT children. A header/footer body may be
        // wrapped in (possibly nested) <w:sdt><w:sdtContent>; without SdtBlock
        // here Get returned zero children and the dump emitted an empty part —
        // dropping the entire header/footer body (PAGE/NUMPAGES fields and all).
        node.ChildCount = element.Elements<Paragraph>().Count()
            + element.Elements<Table>().Count()
            + element.Elements<SdtBlock>().Count();
        if (depth > 0)
        {
            int pIdx = 0, tblIdx = 0, sdtIdx = 0;
            foreach (var child in element.ChildElements)
            {
                if (child is Paragraph hfPara)
                {
                    pIdx++;
                    var paraSegment = BuildParaPathSegment(hfPara, pIdx);
                    node.Children.Add(ElementToNode(hfPara, $"{path}/{paraSegment}", depth - 1));
                }
                else if (child is Table)
                {
                    tblIdx++;
                    node.Children.Add(ElementToNode(child, $"{path}/tbl[{tblIdx}]", depth - 1));
                }
                else if (child is SdtBlock hfSdt)
                {
                    sdtIdx++;
                    node.Children.Add(ElementToNode(hfSdt, $"{path}/{BuildSdtPathSegment(hfSdt, sdtIdx)}", depth - 1));
                }
            }
        }
        return node;
    }

    private DocumentNode BodyToNode(Body bodyNode, DocumentNode node, string path, int depth)
    {
        // CONSISTENCY(body-listing): enumerate body children using the
        // same p[N]/oMathPara[M] counting rules as NavigateToElement so
        // `get /body` emits paths that `get <path>` can resolve. The
        // generic fallback would count every LocalName, listing wrapper
        // <w:p> (pure oMathPara) as p[2] even though the resolver skips
        // them. Mirrors the logic in WordHandler.View.ViewAsText.
        node.ChildCount = bodyNode.ChildElements.Count;
        if (depth > 0)
        {
            int pIdx = 0, tblIdx = 0, mathParaIdx = 0, sdtIdx = 0;
            // BUG-DUMP7-04: w:customXml body wrappers are non-structural —
            // their inner paragraphs and tables should appear as direct
            // body children (with shared p/tbl/sdt counters) so the
            // wrapper itself is invisible to dump but its content
            // round-trips. Recursively flatten any depth of customXml
            // nesting. Without this, the wrapper fell to the generic
            // else and its children were never enumerated.
            void WalkBodyChild(OpenXmlElement child)
            {
                if (child.LocalName == "oMathPara" || child is M.Paragraph)
                {
                    mathParaIdx++;
                    node.Children.Add(ElementToNode(child, $"{path}/oMathPara[{mathParaIdx}]", depth - 1));
                }
                else if (child is Paragraph bPara)
                {
                    if (IsOMathParaWrapperParagraph(bPara))
                    {
                        mathParaIdx++;
                        node.Children.Add(ElementToNode(bPara, $"{path}/oMathPara[{mathParaIdx}]", depth - 1));
                    }
                    else
                    {
                        pIdx++;
                        var bSeg = BuildParaPathSegment(bPara, pIdx);
                        node.Children.Add(ElementToNode(bPara, $"{path}/{bSeg}", depth - 1));
                    }
                }
                else if (child is Table)
                {
                    tblIdx++;
                    node.Children.Add(ElementToNode(child, $"{path}/tbl[{tblIdx}]", depth - 1));
                }
                else if (child is SdtBlock)
                {
                    sdtIdx++;
                    node.Children.Add(ElementToNode(child, $"{path}/sdt[{sdtIdx}]", depth - 1));
                }
                else if (child is CustomXmlBlock cxBlock)
                {
                    foreach (var inner in cxBlock.ChildElements)
                        WalkBodyChild(inner);
                }
                else if (child is BookmarkEnd bodyBkEnd)
                {
                    // BUG-DUMP-BMSPAN: a body-direct BookmarkEnd (POI/Word emit
                    // the End after </w:p> for a paragraph-anchored bookmark).
                    // For a content-wrapping bookmark, surface it as a named
                    // span-end node so EmitBody replays a positioned
                    // `add bookmark name=X end=true` after the wrapped paragraph,
                    // preserving the range. Empty bookmarks need no end node —
                    // the combined start op already recreates the pair.
                    var beNode = ElementToNode(child, $"{path}/bookmarkEnd[1]", depth - 1);
                    var beName = ResolveBookmarkEndName(bodyBkEnd);
                    if (IsContentSpanBookmark(bodyBkEnd) && !string.IsNullOrEmpty(beName))
                        beNode.Format["name"] = beName!;
                    node.Children.Add(beNode);
                }
                else if (child is MoveFromRangeStart or MoveFromRangeEnd
                                  or MoveToRangeStart or MoveToRangeEnd)
                {
                    // BUG-DUMP-R43-10: block-level tracked-move range markers
                    // (<w:moveFromRangeStart>/<w:moveFromRangeEnd>/<w:moveToRangeStart>/
                    // <w:moveToRangeEnd>) appear as direct body children — siblings of
                    // paragraphs — when a whole block was moved with track-changes.
                    // (Run-level moveFrom/moveTo are handled elsewhere.) Surface each
                    // with its localName Type and stash the verbatim outer XML under
                    // _rawMoveRangeXml so EmitBody re-inserts it via a body-level
                    // raw-set, preserving id/name/author/date/colFirst/colLast.
                    var mvNode = ElementToNode(child, $"{path}/{child.LocalName}[1]", depth - 1);
                    mvNode.Format["_rawMoveRangeXml"] = child.OuterXml;
                    node.Children.Add(mvNode);
                }
                else
                {
                    // Non-structural (sectPr etc.) — keep localName naming
                    node.Children.Add(ElementToNode(child, $"{path}/{child.LocalName}[1]", depth - 1));
                }
            }
            foreach (var child in bodyNode.ChildElements)
                WalkBodyChild(child);
        }
        return node;
    }

    private DocumentNode ParagraphToNode(Paragraph para, DocumentNode node, string path, int depth)
    {
        node.Type = "paragraph";
        node.Text = GetParagraphText(para);
        node.Style = GetStyleName(para);
        node.Preview = node.Text?.Length > 50 ? node.Text[..50] + "..." : node.Text;
        node.ChildCount = GetAllRuns(para).Count();

        if (!string.IsNullOrEmpty(para.ParagraphId?.Value))
            node.Format["paraId"] = para.ParagraphId.Value;
        // textId intentionally NOT exposed in Format: Set() rewrites it on
        // every mutation (see WordHandler.Set.cs "para.TextId = GenerateParaId()"),
        // which would let an AI agent comparing consecutive Get snapshots see
        // spurious diffs and mistake idempotent edits for real changes. paraId
        // is stable and sufficient for identity. The underlying w14:textId
        // attribute is still present in the OOXML; only the user-facing
        // DocumentNode.Format projection hides it.

        var pProps = para.ParagraphProperties;
        // AddParagraph writes <w:pPrChange> for `trackChange=format`. The
        // pPrChange block carries author/date attribution alongside a
        // baseline snapshot of the pre-format pPr — mirror what the run
        // side does for <w:rPrChange>.
        var pPrChange = pProps?.GetFirstChild<ParagraphPropertiesChange>();
        if (pPrChange != null)
        {
            node.Format["revision.type"] = "format";
            if (!string.IsNullOrEmpty(pPrChange.Author?.Value))
                node.Format["revision.author"] = pPrChange.Author!.Value!;
            if (pPrChange.Date?.Value is DateTime pDate)
                node.Format["revision.date"] = pDate.ToString("o");
            if (pPrChange.Id?.Value is { } pcId)
                node.Format["revision.id"] = pcId.ToString();
            // BUG: pPrChange's previous-pPr snapshot (alignment, spacing,
            // indent, …) is what Word's Reject Change restores. Add v1
            // re-stamps the marker with empty previous-pPr, losing the
            // snapshot on dump→batch round-trip. SDK 3.x writes the snapshot
            // as ParagraphPropertiesExtended (SetRevision path) or
            // PreviousParagraphProperties (AddParagraph's empty-marker path /
            // foreign producers). Probe both.
            var prevExt = pPrChange.GetFirstChild<ParagraphPropertiesExtended>();
            var prevPpr = pPrChange.GetFirstChild<PreviousParagraphProperties>();
            // BUG-DUMP-R43-8: carry the verbatim prior-pPr snapshot so the
            // emitter restores it via revision.beforeXml instead of stamping an
            // empty <w:pPr/> marker. The inner element's OuterXml round-trips
            // through AddParagraph's pPrChange InnerXml assignment.
            // BUG-DUMP-PPRCHANGE-CS-EMPTYSNAP: emit beforeXml even when the prior
            // snapshot is EMPTY (the element exists but has no children — Word's
            // "format changed from the default" marker). Prefer the populated
            // element, but fall back to an empty one so the key is still present.
            // Without it, an empty-snapshot pPrChange on a paragraph that also
            // carries complex-script (.cs) run props made the replayed `set` hit
            // the "RTL cascade properties not supported with trackChange" guard
            // (which is bypassed only when revision.beforeXml is supplied) — the
            // op failed and the cell's content was dropped. An empty snapshot
            // round-trips as an empty <w:pPr/> via ApplyBeforeXmlSnapshot, so no
            // smearing occurs.
            OpenXmlElement? prevPpEl =
                  (prevExt != null && prevExt.HasChildren) ? prevExt
                : (prevPpr != null && prevPpr.HasChildren) ? prevPpr
                : (OpenXmlElement?)prevExt ?? prevPpr;
            if (prevPpEl != null)
                node.Format["revision.beforeXml"] = prevPpEl.OuterXml;
        }
        // paraMarkIns: `<w:pPr><w:rPr><w:ins .../></w:rPr></w:pPr>` records
        // that the paragraph mark itself was inserted as a tracked change —
        // distinct from pPrChange (format-change snapshot) and from any
        // content-run wrappers. Surfaced under a paraMarkIns.* namespace so
        // dump can translate it back into AddParagraph's bare-trackChange.*
        // form on replay (which re-creates both the ¶ mark and any
        // accompanying content wrapping in one step). Kept distinct from
        // `trackChange=format` to avoid clobbering pPrChange attribution
        // when both are present on the same paragraph.
        var pmrpRev = pProps?.ParagraphMarkRunProperties;
        if (pmrpRev != null)
        {
            var pMarkIns = pmrpRev.GetFirstChild<Inserted>();
            if (pMarkIns != null)
            {
                if (!string.IsNullOrEmpty(pMarkIns.Author?.Value))
                    node.Format["paraMarkIns.author"] = pMarkIns.Author!.Value!;
                if (pMarkIns.Date?.Value is DateTime piDate)
                    node.Format["paraMarkIns.date"] = piDate.ToString("o");
                if (pMarkIns.Id?.Value is { } piId)
                    node.Format["paraMarkIns.id"] = piId.ToString();
            }
            // paraMarkDel: <w:pPr><w:rPr><w:del .../></w:rPr></w:pPr> records
            // that the paragraph mark was *deleted* as a tracked change
            // (paragraph-join revision). Mirror paraMarkIns so dump round-trip
            // preserves it.
            var pMarkDel = pmrpRev.GetFirstChild<Deleted>();
            if (pMarkDel != null)
            {
                if (!string.IsNullOrEmpty(pMarkDel.Author?.Value))
                    node.Format["paraMarkDel.author"] = pMarkDel.Author!.Value!;
                if (pMarkDel.Date?.Value is DateTime pdDate)
                    node.Format["paraMarkDel.date"] = pdDate.ToString("o");
                if (pMarkDel.Id?.Value is { } pdId)
                    node.Format["paraMarkDel.id"] = pdId.ToString();
            }
        }
        if (pProps != null)
        {
            if (pProps.ParagraphStyleId?.Val?.Value != null)
            {
                // CONSISTENCY(style-dual-key): `style` carries the OOXML
                // styleId (canonical handle used by basedOn/pStyle/rStyle).
                // `styleName` carries the user-facing display name. Both
                // are emitted so query selectors can pick precision
                // (styleId=/styleName=) or convenience (style=, lenient).
                node.Format["style"] = pProps.ParagraphStyleId.Val.Value;
                node.Format["styleId"] = pProps.ParagraphStyleId.Val.Value;
                var displayName = GetStyleName(para);
                if (!string.IsNullOrEmpty(displayName))
                    node.Format["styleName"] = displayName;
            }
            if (pProps.Justification?.Val != null)
            {
                var alignText = pProps.Justification.Val.InnerText;
                var alignValue = alignText == "both" ? "justify" : alignText;
                node.Format["align"] = alignValue;
            }
            if (pProps.SpacingBetweenLines != null)
            {
                if (pProps.SpacingBetweenLines.Before?.Value != null)
                {
                    node.Format["spaceBefore"] = SpacingConverter.FormatWordSpacingNonNegative(pProps.SpacingBetweenLines.Before.Value);
                }
                if (pProps.SpacingBetweenLines.After?.Value != null)
                {
                    node.Format["spaceAfter"] = SpacingConverter.FormatWordSpacingNonNegative(pProps.SpacingBetweenLines.After.Value);
                }
                if (pProps.SpacingBetweenLines.Line?.Value != null)
                {
                    node.Format["lineSpacing"] = SpacingConverter.FormatWordLineSpacing(
                        pProps.SpacingBetweenLines.Line.Value,
                        pProps.SpacingBetweenLines.LineRule?.InnerText);
                }
                if (pProps.SpacingBetweenLines.LineRule?.HasValue == true)
                {
                    node.Format["lineRule"] = pProps.SpacingBetweenLines.LineRule.InnerText;
                }
                // CONSISTENCY(ind-chars): mirror style-level Get (Query.cs)
                // for the chars-unit space-before/after slots so P1-7
                // round-trip works on paragraphs as well as styles.
                if (pProps.SpacingBetweenLines.BeforeLines?.Value != null)
                {
                    node.Format["spaceBeforeLines"] = pProps.SpacingBetweenLines.BeforeLines.Value;
                }
                if (pProps.SpacingBetweenLines.AfterLines?.Value != null)
                {
                    node.Format["spaceAfterLines"] = pProps.SpacingBetweenLines.AfterLines.Value;
                }
                // BUG-DUMP-R44-4: the auto-spacing on/off toggles
                // (w:beforeAutospacing / w:afterAutospacing — Word's "automatic
                // spacing between paragraphs of the same style") were never read
                // back, so dump→batch silently dropped them. Emit as bool toggles
                // under spaceBeforeAuto / spaceAfterAuto (canonical, matching the
                // spaceBefore* key family).
                if (pProps.SpacingBetweenLines.BeforeAutoSpacing?.Value != null)
                {
                    node.Format["spaceBeforeAuto"] = pProps.SpacingBetweenLines.BeforeAutoSpacing.Value;
                }
                if (pProps.SpacingBetweenLines.AfterAutoSpacing?.Value != null)
                {
                    node.Format["spaceAfterAuto"] = pProps.SpacingBetweenLines.AfterAutoSpacing.Value;
                }
            }
            if (pProps.Indentation != null)
            {
                // Malformed sources (some legal-document / HTML-export
                // generators) split the indent across TWO <w:ind> elements —
                // e.g. <w:ind w:start="360"/><w:ind w:firstLine="360"/>.
                // GetFirstChild returns only the first, so a firstLine/hanging
                // carried on a later element was dropped and the paragraph's
                // first-line indent vanished on round-trip. Word merges the
                // duplicate elements; coalesce each attribute across EVERY
                // <w:ind> child (first element that sets it wins) so the
                // readback matches what Word renders.
                var allInd = pProps.Elements<Indentation>().ToList();
                string? PickStr(Func<Indentation, string?> sel)
                    => allInd.Select(sel).FirstOrDefault(v => v != null);
                int? PickInt(Func<Indentation, int?> sel)
                    => allInd.Select(sel).FirstOrDefault(v => v != null);
                // CONSISTENCY(unit-qualified-spacing): indents return "Xpt" via SpacingConverter,
                // matching spaceBefore/spaceAfter (Canonical DocumentNode.Format Rules).
                var firstLineV = PickStr(i => i.FirstLine?.Value);
                if (firstLineV != null) node.Format["firstLineIndent"] = SpacingConverter.FormatWordSpacing(firstLineV);
                var hangingV = PickStr(i => i.Hanging?.Value);
                if (hangingV != null) node.Format["hangingIndent"] = SpacingConverter.FormatWordSpacing(hangingV);
                // CONSISTENCY(ind-start-end): modern Word writes <w:ind w:start>/<w:end> instead of left/right.
                var leftTwips = PickStr(i => i.Left?.Value ?? i.Start?.Value);
                if (leftTwips != null) node.Format["indent"] = SpacingConverter.FormatWordSpacing(leftTwips);
                var rightTwips = PickStr(i => i.Right?.Value ?? i.End?.Value);
                if (rightTwips != null) node.Format["rightIndent"] = SpacingConverter.FormatWordSpacing(rightTwips);
                // CONSISTENCY(ind-chars): chars-unit indents (Chinese typography) — backfilled from style Get edc8f884.
                var firstLineChars = PickInt(i => i.FirstLineChars?.Value);
                if (firstLineChars != null) node.Format["firstLineChars"] = firstLineChars.Value;
                var hangingChars = PickInt(i => i.HangingChars?.Value);
                if (hangingChars != null) node.Format["hangingChars"] = hangingChars.Value;
                var leftChars = PickInt(i => i.LeftChars?.Value ?? i.StartCharacters?.Value);
                if (leftChars != null) node.Format["leftChars"] = leftChars.Value;
                var rightChars = PickInt(i => i.RightChars?.Value ?? i.EndCharacters?.Value);
                if (rightChars != null) node.Format["rightChars"] = rightChars.Value;
            }
            if (pProps.KeepNext != null)
            {
                var v = pProps.KeepNext.Val;
                node.Format["keepNext"] = v == null || v.Value;
            }
            if (pProps.KeepLines != null)
            {
                var v = pProps.KeepLines.Val;
                node.Format["keepLines"] = v == null || v.Value;
            }
            if (pProps.PageBreakBefore != null)
            {
                var v = pProps.PageBreakBefore.Val;
                node.Format["pageBreakBefore"] = v == null || v.Value;
            }
            if (pProps.WidowControl != null)
            {
                // Val == null or Val == true means enabled; Val == false means explicitly disabled
                var wcVal = pProps.WidowControl.Val;
                node.Format["widowControl"] = wcVal == null || wcVal.Value;
            }
            if (pProps.BiDi != null)
            {
                // <w:bidi/> default Val is true; explicit Val=false toggles
                // it off. Emit canonical 'direction' so writers can clone
                // the paragraph with the same key they used to set it.
                // R8-fuzz-5: pProps.BiDi.Val.Value invokes OnOffValue.Parse
                // and throws FormatException on garbage attribute text
                // (e.g. <w:bidi w:val="garbage"/>). Skip the key on
                // unparseable input — Get must never crash on a doc that
                // disk-loaded fine, even when validate would flag the same
                // attribute as schema-invalid.
                bool? bidiOn = TryReadOnOff(pProps.BiDi.Val);
                if (bidiOn.HasValue)
                    node.Format["direction"] = bidiOn.Value ? "rtl" : "ltr";
            }
            if (pProps.ContextualSpacing != null)
            {
                var csVal = pProps.ContextualSpacing.Val;
                node.Format["contextualSpacing"] = csVal == null || csVal.Value;
            }
            if (pProps.Shading != null)
            {
                // CONSISTENCY(shd-canonical-fill): solid paragraph shading reads
                // back as the canonical `fill` key (matches table cells / runs);
                // true pattern/theme keeps the shading.val/.fill/.color detail keys.
                ReadShadingCanonical(pProps.Shading, node);
            }

            var pBdr = pProps.ParagraphBorders;
            if (pBdr != null)
            {
                ReadBorder(pBdr.TopBorder, "pbdr.top", node);
                ReadBorder(pBdr.BottomBorder, "pbdr.bottom", node);
                ReadBorder(pBdr.LeftBorder, "pbdr.left", node);
                ReadBorder(pBdr.RightBorder, "pbdr.right", node);
                ReadBorder(pBdr.BetweenBorder, "pbdr.between", node);
                ReadBorder(pBdr.BarBorder, "pbdr.bar", node);
            }

            var numProps = pProps.NumberingProperties;
            if (numProps != null && numProps.NumberingId?.Val?.Value != null)
            {
                var numIdVal = numProps.NumberingId.Val.Value;
                node.Format["numId"] = numIdVal.ToString();
                var ilvlVal = numProps.NumberingLevelReference?.Val?.Value ?? 0;
                // R29-3: surface under the canonical key 'numLevel' (paragraph.json
                // declares numLevel canonical with ilvl as an input alias; Get
                // normalizes to the single canonical key). Style/abstractNum-level
                // contexts keep 'ilvl' as their own canonical key.
                node.Format["numLevel"] = ilvlVal.ToString();
                // numId=0 is the OOXML "remove numbering" sentinel — the paragraph
                // explicitly opts out of any inherited list style. Skip numFmt /
                // listStyle / start lookup so Get does not falsely advertise a list.
                if (numIdVal != 0)
                {
                    var numFmt = GetNumberingFormat(numIdVal, ilvlVal);
                    node.Format["numFmt"] = numFmt;
                    node.Format["listStyle"] = numFmt.ToLowerInvariant() == "bullet" ? "bullet" : "ordered";
                    var start = GetStartValue(numIdVal, ilvlVal);
                    if (start != null)
                        node.Format["start"] = start.Value;
                }
            }
            else if (numProps != null)
            {
                // BUG-DUMP-H77: a direct <w:numPr> with NO <w:numId> child — a bare
                // list-level override (<w:ilvl> only) or a tracked numbering-insertion
                // (<w:numPr><w:ins/></w:numPr>, no ilvl/numId). The numId branch above
                // is skipped, and the style-inheritance fallback would either miss it
                // or wrongly promote inherited numbering — so the whole numPr was
                // silently dropped on `add p`. Surface numLevel directly (NO numId, so
                // no numFmt/listStyle/start lookup and NO numInherited flag); the
                // numPrIns.* readback below fires for both numId and numId-less numPr.
                var bareIlvl = numProps.NumberingLevelReference?.Val?.Value;
                if (bareIlvl.HasValue)
                    node.Format["numLevel"] = bareIlvl.Value.ToString();
            }
            else
            {
                // Fall back to the style chain — paragraphs that inherit numbering
                // from styles like ListBullet / ListNumber don't have a direct numPr,
                // but Get should still surface the effective list metadata.
                var inherited = ResolveNumPrFromStyle(para);
                if (inherited.HasValue)
                {
                    var (inhId, inhLvl) = inherited.Value;
                    node.Format["numId"] = inhId.ToString();
                    // R29-3: canonical key 'numLevel' (see direct-numPr branch above).
                    node.Format["numLevel"] = inhLvl.ToString();
                    // BUG-DUMP26-01: flag style-inherited values so WordBatchEmitter
                    // can suppress them on `add p` — they're already covered by
                    // the paragraph's style and emitting them would semantically
                    // promote inherited→explicit on round-trip. Mirrors the
                    // round-1 first-run hoist precedent.
                    node.Format["numInherited"] = "true";
                    var numFmt = GetNumberingFormat(inhId, inhLvl);
                    node.Format["numFmt"] = numFmt;
                    node.Format["listStyle"] = numFmt.ToLowerInvariant() == "bullet" ? "bullet" : "ordered";
                    var start = GetStartValue(inhId, inhLvl);
                    if (start != null)
                        node.Format["start"] = start.Value;
                }
            }
            // BUG-DUMP-R49-1 / BUG-DUMP-H77: <w:numPr><w:ins .../> is a tracked
            // insertion of the list-numbering assignment (Reviewing pane: "Formatted:
            // List Paragraph"). Surface as numPrIns.* so the batch emitter can replay
            // <w:numPr><w:ins> after the paragraph is created (no first-class Add/Set
            // vocabulary for numPr tracked changes; verbatim is the safest round-trip).
            // Fires for BOTH a numId-bearing numPr AND a numId-less one — a tracked
            // numbering insertion frequently carries no numId/ilvl at all.
            if (numProps != null)
            {
                var numPrIns = numProps.GetFirstChild<Inserted>();
                if (numPrIns != null)
                {
                    if (!string.IsNullOrEmpty(numPrIns.Author?.Value))
                        node.Format["numPrIns.author"] = numPrIns.Author!.Value!;
                    if (numPrIns.Date?.Value is DateTime npiDate)
                        node.Format["numPrIns.date"] = npiDate.ToString("o");
                    if (numPrIns.Id?.Value is { } npiId)
                        node.Format["numPrIns.id"] = npiId.ToString();
                }
            }

            // CONSISTENCY(outline-lvl): backfilled from style Get edc8f884. Paragraph-level outlineLvl overrides style.
            if (pProps.OutlineLevel?.Val?.Value != null)
                node.Format["outlineLvl"] = (int)pProps.OutlineLevel.Val.Value;

            // CONSISTENCY(tabs): backfilled from style Get edc8f884.
            if (pProps.Tabs != null)
            {
                var tabList = new List<Dictionary<string, object?>>();
                foreach (var tab in pProps.Tabs.Elements<TabStop>())
                {
                    var t = new Dictionary<string, object?>();
                    if (tab.Position?.Value != null) t["pos"] = tab.Position.Value;
                    if (tab.Val?.HasValue == true) t["val"] = tab.Val.InnerText;
                    if (tab.Leader?.HasValue == true) t["leader"] = tab.Leader.InnerText;
                    if (t.Count > 0) tabList.Add(t);
                }
                if (tabList.Count > 0) node.Format["tabs"] = tabList;
            }

            // Long-tail fallback: surface every pPr child the curated reader
            // didn't consume. Symmetric with the Set-side TryCreateTypedChild
            // fallback in SetElementParagraph (WordHandler.Set.Element.cs).
            FillUnknownChildProps(pProps, node);

            // CONSISTENCY(add-set-symmetry): inline section break.
            // A paragraph carrying <w:sectPr> inside its <w:pPr> is the
            // OOXML representation of a mid-document section break (the
            // last paragraph before the break holds the section's
            // properties). AddSection on /body produces exactly this
            // shape, but Get used to expose nothing — leaving the
            // paragraph indistinguishable from a regular empty para.
            // Surface it as `sectionBreak` (Add prop name match) plus
            // companion section-property keys readers expect.
            var inlineSectPr = pProps.GetFirstChild<SectionProperties>();
            if (inlineSectPr != null)
            {
                // BUG-DUMP-R31-1: a mid-document <w:sectPr> may be CHILDLESS
                // (no <w:type>, no pgSz/pgMar) — the source author deferred the
                // break kind and page geometry to Word's defaults, exactly like
                // an empty FINAL body sectPr. Emitting a fabricated
                // sectionBreak="nextPage" here (and a default pgSz/pgMar on
                // rebuild) injects a <w:type>/geometry the source never had.
                // Surface the REAL <w:type> (null when absent) and flag a truly
                // childless sectPr so the emitter omits `type` and the apply
                // produces a bare <w:sectPr/>. A mid sectPr WITH real type or
                // geometry keeps emitting those keys below.
                var sectMark = inlineSectPr.GetFirstChild<SectionType>()?.Val?.InnerText;
                node.Format["sectionBreak"] = sectMark ?? "nextPage";
                if (sectMark != null)
                    node.Format["sectionBreak.type"] = sectMark;
                if (!inlineSectPr.HasChildren)
                    node.Format["sectionBreak.empty"] = true;
                // BUG-DUMP-SECT-TYPEINJECT: a sectPr WITH children (pgSz/docGrid/
                // …) but NO <w:type> deferred the break kind to the OOXML default
                // (nextPage). The `empty` flag above only covers a fully childless
                // sectPr, so this non-empty/no-type case fell through and AddSection
                // default-stamped <w:type w:val="nextPage"/> — an explicit section
                // page break the source never had (+1 page). Signal it so the
                // emitter forwards `notype=true` and AddSection skips the stamp.
                else if (sectMark == null)
                    node.Format["sectionBreak.notype"] = true;

                // Per-section page layout when overridden on this break.
                // Emit native OOXML twips (bare integers) rather than the
                // cm-rounded human form: these sectionBreak.* keys exist ONLY
                // for the dump→batch round-trip (canonical Get readback for a
                // section is via `query section`), and twip→cm→twip rounds to
                // 2 decimals so it drifts ±1 twip per cycle (1418→"2.5cm"→1417).
                // Bare integers parse back as exact twips (ParseTwips
                // fallthrough), so the rebuilt inline sectPr matches byte-for-
                // byte. Mirrors the body-sectPr fix in BodySectionPageGeometryTwips.
                var pgSz = inlineSectPr.GetFirstChild<PageSize>();
                if (pgSz?.Width?.Value != null)
                    node.Format["sectionBreak.pageWidth"] = pgSz.Width.Value.ToString();
                if (pgSz?.Height?.Value != null)
                    node.Format["sectionBreak.pageHeight"] = pgSz.Height.Value.ToString();
                if (pgSz?.Orient?.Value != null)
                    node.Format["sectionBreak.orientation"] = pgSz.Orient.InnerText;

                var pgMar = inlineSectPr.GetFirstChild<PageMargin>();
                if (pgMar != null)
                {
                    if (pgMar.Top?.Value != null)
                        node.Format["sectionBreak.marginTop"] = ((uint)Math.Abs(pgMar.Top.Value)).ToString();
                    if (pgMar.Bottom?.Value != null)
                        node.Format["sectionBreak.marginBottom"] = ((uint)Math.Abs(pgMar.Bottom.Value)).ToString();
                    if (pgMar.Left?.Value != null)
                        node.Format["sectionBreak.marginLeft"] = pgMar.Left.Value.ToString();
                    if (pgMar.Right?.Value != null)
                        node.Format["sectionBreak.marginRight"] = pgMar.Right.Value.ToString();
                    // header/footer-from-edge + binding gutter (mirror the root
                    // and /section[N] readbacks) so a mid-document section
                    // break round-trips its full pgMar, not just the 4 edges.
                    if (pgMar.Header?.Value != null)
                        node.Format["sectionBreak.marginHeader"] = pgMar.Header.Value.ToString();
                    if (pgMar.Footer?.Value != null)
                        node.Format["sectionBreak.marginFooter"] = pgMar.Footer.Value.ToString();
                    if (pgMar.Gutter?.Value != null)
                        node.Format["sectionBreak.marginGutter"] = pgMar.Gutter.Value.ToString();
                }

                var pgNum = inlineSectPr.GetFirstChild<PageNumberType>();
                if (pgNum?.Start?.Value != null)
                    node.Format["sectionBreak.pageStart"] = pgNum.Start.Value;
                if (pgNum?.Format?.Value != null)
                    node.Format["sectionBreak.pageNumFmt"] = pgNum.Format.InnerText;
                // BUG-DUMP-SECT-CHAPNUM: chapter-number page numbering (chapStyle =
                // heading style index, chapSep = separator between chapter and page,
                // e.g. "1-1"/"2.3") on a mid-document section carrier. The body-sectPr
                // readback emits both; without mirroring them here a non-trailing
                // section's page numbers silently lost their chapter prefix/separator.
                if (pgNum?.ChapterStyle?.Value != null)
                    node.Format["sectionBreak.chapStyle"] = pgNum.ChapterStyle.Value;
                if (pgNum?.ChapterSeparator?.Value != null)
                    node.Format["sectionBreak.chapSep"] = pgNum.ChapterSeparator.InnerText;

                if (inlineSectPr.GetFirstChild<TitlePage>() != null)
                    node.Format["sectionBreak.titlePage"] = true;

                // BUG-DUMP-SECT-NOENDNOTE: <w:noEndnote/> suppresses endnote
                // collection for the section. The body-sectPr readback emits it;
                // omitting it here let suppressed endnotes reappear in a non-trailing
                // section on round-trip.
                if (IsToggleOn(inlineSectPr.GetFirstChild<NoEndnote>()))
                    node.Format["sectionBreak.noEndnote"] = "true";

                // BUG-DUMP-SECT-PAPERSRC: printer paper-source bins on a
                // mid-document section carrier. Surface as sectionBreak.paperSrc.*
                // so the carrier sectPr round-trips the printer tray config.
                var sbPaperSrc = inlineSectPr.GetFirstChild<PaperSource>();
                if (sbPaperSrc != null)
                {
                    if (sbPaperSrc.First?.Value != null)
                        node.Format["sectionBreak.paperSrc.first"] = sbPaperSrc.First.Value;
                    if (sbPaperSrc.Other?.Value != null)
                        node.Format["sectionBreak.paperSrc.other"] = sbPaperSrc.Other.Value;
                }

                // BUG-DUMP-SECT-FORMPROT: <w:formProt/> on a mid-document section
                // carrier. BUG-DUMP-R40-4: ST_OnOff — surface the actual on/off
                // value (presence-only test flipped an explicit-false to true on
                // round-trip; AddSection omits the element on a falsey value).
                var inlineFormProt = inlineSectPr.GetFirstChild<FormProtection>();
                if (inlineFormProt != null)
                    node.Format["sectionBreak.formProt"] = inlineFormProt.Val == null || inlineFormProt.Val.Value;

                // BUG-DUMP9-06: Columns / VerticalTextAlignmentOnPage on
                // an inline sectPr carrier were silently dropped — only
                // the root sectPr reader handled them. Surface as
                // sectionBreak.columns / sectionBreak.vAlign so dump
                // round-trips the carrier sectPr.
                var sbCols = inlineSectPr.GetFirstChild<Columns>();
                if (sbCols != null)
                {
                    if (sbCols.ColumnCount?.Value != null)
                        node.Format["sectionBreak.columns"] = (int)sbCols.ColumnCount.Value;
                    // BUG-DUMP-R25-4: emit cols @w:space as RAW TWIPS (bare int),
                    // not 2-decimal cm. pgMar/pgSz already round-trip raw twips;
                    // the cols space was the lone cm-drift survivor (708→"1.25cm"
                    // →709 on replay). Set's columns.space case feeds ParseTwips,
                    // which reads a bare integer exactly.
                    if (sbCols.Space?.Value != null && uint.TryParse(sbCols.Space.Value, out var sbColSpaceTwips))
                        node.Format["sectionBreak.columnSpace"] = sbColSpaceTwips.ToString();
                    if (sbCols.EqualWidth?.Value != null)
                        node.Format["sectionBreak.columns.equalWidth"] = sbCols.EqualWidth.Value;
                    if (sbCols.Separator?.Value == true)
                        node.Format["sectionBreak.columns.separator"] = true;
                    // BUG-DUMP-R25-3: surface the explicit per-column widths/
                    // spaces for an unequal-width (equalWidth="false") inline
                    // section. Without these the carrier <w:cols equalWidth="0">
                    // round-tripped with no <w:col> children, collapsing the
                    // source's uneven columns to equal width. Mirrors the body-
                    // section colWidths/colSpaces readback.
                    var sbColDefs = sbCols.Elements<Column>().ToList();
                    if (sbColDefs.Count > 0)
                    {
                        node.Format["sectionBreak.colWidths"] = string.Join(",", sbColDefs.Select(c => c.Width?.Value ?? "0"));
                        node.Format["sectionBreak.colSpaces"] = string.Join(",", sbColDefs.Select(c => c.Space?.Value ?? "0"));
                    }
                }

                var sbVAlign = inlineSectPr.GetFirstChild<VerticalTextAlignmentOnPage>();
                if (sbVAlign?.Val != null)
                    node.Format["sectionBreak.vAlign"] = sbVAlign.Val.InnerText;

                // BUG-DUMP-SECT-TEXTDIR: section page text flow on a mid-document
                // section carrier. Surface as sectionBreak.textDirection so the
                // carrier sectPr round-trips — distinct from cell-level tcPr.
                var sbTextDir = inlineSectPr.GetFirstChild<TextDirection>();
                if (sbTextDir?.Val != null)
                    node.Format["sectionBreak.textDirection"] = sbTextDir.Val.InnerText;

                // BUG-DUMP-SECT-RTL: a mid-document section carrier can be RTL
                // (Hebrew/Arabic/Persian multi-section docs) via <w:bidi/> (page
                // direction) and <w:rtlGutter/> (binding gutter on the right).
                // The body-sectPr readback reads both; without surfacing them on
                // the carrier every non-trailing section silently flipped LTR and
                // the gutter moved to the left. AddSection consumes bidi/rtlGutter.
                if (IsToggleOn(inlineSectPr.GetFirstChild<BiDi>()))
                    node.Format["sectionBreak.bidi"] = "rtl";
                if (IsToggleOn(inlineSectPr.GetFirstChild<GutterOnRight>()))
                    node.Format["sectionBreak.rtlGutter"] = "true";

                // Page border on a mid-document section carrier (e.g. a cover
                // page that boxes only its own first page via display="firstPage").
                // Surface per-side detail + offsetFrom/zOrder/display under the
                // sectionBreak.pgBorders.* prefix so the carrier sectPr round-trips
                // — mirrors the body-sectPr ReadPageBorders path. Without this the
                // <w:pgBorders> child was dropped entirely on rebuild (the body
                // sectPr's Get path never sees a carrier paragraph's sectPr).
                var sbPgBorders = inlineSectPr.GetFirstChild<PageBorders>();
                if (sbPgBorders != null)
                {
                    ReadBorder(sbPgBorders.TopBorder, "sectionBreak.pgBorders.top", node);
                    ReadBorder(sbPgBorders.LeftBorder, "sectionBreak.pgBorders.left", node);
                    ReadBorder(sbPgBorders.BottomBorder, "sectionBreak.pgBorders.bottom", node);
                    ReadBorder(sbPgBorders.RightBorder, "sectionBreak.pgBorders.right", node);
                    if (sbPgBorders.OffsetFrom?.InnerText is { } sbOff)
                        node.Format["sectionBreak.pgBorders.offsetFrom"] = sbOff;
                    if (sbPgBorders.ZOrder?.InnerText is { } sbZ)
                        node.Format["sectionBreak.pgBorders.zOrder"] = sbZ;
                    if (sbPgBorders.Display?.InnerText is { } sbDisp)
                        node.Format["sectionBreak.pgBorders.display"] = sbDisp;
                }

                // BUG-DUMP-SECT-FOOTNOTE: footnote/endnote numbering on a
                // mid-document section carrier. Surface as sectionBreak.footnotePr.*
                // / sectionBreak.endnotePr.* so the carrier sectPr round-trips —
                // without this, footnote markers reverted from i/ii to 1/2.
                var sbFn = inlineSectPr.GetFirstChild<FootnoteProperties>();
                if (sbFn != null)
                {
                    if (sbFn.NumberingFormat?.Val != null)
                        node.Format["sectionBreak.footnotePr.numFmt"] = sbFn.NumberingFormat.Val.InnerText;
                    if (sbFn.NumberingRestart?.Val != null)
                        node.Format["sectionBreak.footnotePr.numRestart"] = sbFn.NumberingRestart.Val.InnerText;
                    if (sbFn.NumberingStart?.Val != null)
                        node.Format["sectionBreak.footnotePr.numStart"] = (int)sbFn.NumberingStart.Val.Value;
                    if (sbFn.FootnotePosition?.Val != null)
                        node.Format["sectionBreak.footnotePr.pos"] = sbFn.FootnotePosition.Val.InnerText;
                }
                var sbEn = inlineSectPr.GetFirstChild<EndnoteProperties>();
                if (sbEn != null)
                {
                    if (sbEn.NumberingFormat?.Val != null)
                        node.Format["sectionBreak.endnotePr.numFmt"] = sbEn.NumberingFormat.Val.InnerText;
                    if (sbEn.NumberingRestart?.Val != null)
                        node.Format["sectionBreak.endnotePr.numRestart"] = sbEn.NumberingRestart.Val.InnerText;
                    if (sbEn.NumberingStart?.Val != null)
                        node.Format["sectionBreak.endnotePr.numStart"] = (int)sbEn.NumberingStart.Val.Value;
                    if (sbEn.EndnotePosition?.Val != null)
                        node.Format["sectionBreak.endnotePr.pos"] = sbEn.EndnotePosition.Val.InnerText;
                }

                var lnNum = inlineSectPr.GetFirstChild<LineNumberType>();
                if (lnNum != null)
                {
                    // BUG-DUMP-SECT-LNDIST: only surface lineNumbers when @w:restart
                    // is present — defaulting to "continuous" fabricated a spurious
                    // restart="continuous" on a carrier that had only @countBy/@distance.
                    if (lnNum.Restart?.InnerText is string sbLnRestart)
                        node.Format["sectionBreak.lineNumbers"] = sbLnRestart switch
                        {
                            "newPage" => "restartPage",
                            "newSection" => "restartSection",
                            _ => "continuous"
                        };
                    if (lnNum.CountBy?.Value is short cb && cb > 1)
                        node.Format["sectionBreak.lineNumberCountBy"] = cb;
                    // BUG-DUMP-SECT-LNSTART: w:lnNumType/@w:start (first line number)
                    // on a mid-document carrier — body-sectPr readback emits it; the
                    // carrier dropped it, restarting line numbering at the default.
                    if (lnNum.Start?.Value is short sbLnStart)
                        node.Format["sectionBreak.lineNumberStart"] = (int)sbLnStart;
                    // BUG-DUMP-SECT-LNDIST: w:lnNumType/@w:distance (gutter twips).
                    if (lnNum.Distance?.Value is string sbLnDistRaw
                        && int.TryParse(sbLnDistRaw, out var sbLnDist))
                        node.Format["sectionBreak.lineNumberDistance"] = sbLnDist;
                }

                // BUG-DUMP-SECGRID: the document grid (<w:docGrid>) on a
                // mid-document section break governs CJK line pitch / lines-
                // per-page; dropping it on dump→batch reflowed every page of
                // that section (a major pagination-drift source). Mirror the
                // root / section readback (docGrid.type / .linePitch / .charSpace);
                // AddSection replays them via its docGrid.* typed fallback.
                var sbGrid = inlineSectPr.GetFirstChild<DocGrid>();
                if (sbGrid != null)
                {
                    if (sbGrid.Type?.HasValue == true)
                        node.Format["sectionBreak.docGrid.type"] = sbGrid.Type.InnerText;
                    if (sbGrid.LinePitch?.Value != null)
                        node.Format["sectionBreak.docGrid.linePitch"] = sbGrid.LinePitch.Value;
                    // BUG-R7B(BUG3): charSpace is ST_DecimalNumber (signed), but
                    // Word stores a negative value as its unsigned 32-bit wrap
                    // (e.g. -6145 -> 4294961151) which overflows Int32Value.Value
                    // and threw "Value was either too large or too small for an
                    // Int32" on dump. Read the raw text and wrap unsigned->signed,
                    // mirroring the lenient root/section readback in
                    // WordHandler.Navigation.DocSettings.cs / Query.cs. A
                    // non-numeric value is skipped rather than crashing the dump.
                    if (sbGrid.CharacterSpace != null)
                    {
                        var rawCs = sbGrid.CharacterSpace.InnerText;
                        if (long.TryParse(rawCs, System.Globalization.NumberStyles.Integer,
                                System.Globalization.CultureInfo.InvariantCulture, out var csVal))
                        {
                            if (csVal > int.MaxValue) csVal -= 4294967296L; // 2^32 unsigned->signed
                            node.Format["sectionBreak.docGrid.charSpace"] = (int)csVal;
                        }
                    }
                }
            }
        }

        // BUG-DUMP9-02: surface paragraph-mark-only run formatting under
        // the `markRPr.*` namespace whenever pPr/rPr exists. The
        // run-fallback path below promotes mark rPr to bare keys only
        // when there are no runs (round-1 hoisting fix); when runs are
        // present, mark-only formatting on the ¶ glyph used to be
        // silently dropped on dump round-trip. Emit dedicated keys so
        // replay can target ParagraphMarkRunProperties without conflating
        // with run-level formatting.
        var pmrpForDump = para.ParagraphProperties?.ParagraphMarkRunProperties;
        // Local shared readback for the ¶-mark glyph props that were missing
        // from BOTH mark paths (the dotted markRPr.* block below for
        // text-bearing paragraphs AND the bare-key fallback for run-less /
        // collapsed paragraphs): caps/smallCaps, dstrike, the text effects
        // (outline/shadow/emboss/imprint), vertAlign, character border,
        // vanish and specVanish. caps/smallCaps change the glyph width (and
        // an empty paragraph's line height), vertAlign moves the baseline,
        // specVanish is Word's style-separator marker — all silently dropped
        // on dump→batch. Key forms mirror the run reader (RunToNode) so
        // replay routes through ApplyRunFormatting unchanged.
        void ReadMarkGlyphProps(OpenXmlCompositeElement src, string prefix)
        {
            void Emit(string key, object val)
            {
                var k = prefix + key;
                if (!node.Format.ContainsKey(k)) node.Format[k] = val;
            }
            void Tog(OnOffType? el, string key)
            {
                if (el != null) Emit(key, el.Val == null || el.Val.Value);
            }
            Tog(src.GetFirstChild<Caps>(), "caps");
            Tog(src.GetFirstChild<SmallCaps>(), "smallcaps");
            Tog(src.GetFirstChild<DoubleStrike>(), "dstrike");
            Tog(src.GetFirstChild<Outline>(), "outline");
            Tog(src.GetFirstChild<Shadow>(), "shadow");
            Tog(src.GetFirstChild<Emboss>(), "emboss");
            Tog(src.GetFirstChild<Imprint>(), "imprint");
            Tog(src.GetFirstChild<Vanish>(), "vanish");
            Tog(src.GetFirstChild<SpecVanish>(), "specVanish");
            var va = src.GetFirstChild<VerticalTextAlignment>()?.Val?.Value;
            if (va == VerticalPositionValues.Superscript) Emit("superscript", true);
            else if (va == VerticalPositionValues.Subscript) Emit("subscript", true);
            // <w:position> raise/lower — pt form, same as the run reader
            // (the dotted markRPr.position raw-half-point emit above wins
            // when it already ran; ApplyRunFormatting accepts both forms).
            var mgPos = src.GetFirstChild<Position>()?.Val?.Value;
            if (!string.IsNullOrEmpty(mgPos)
                && int.TryParse(mgPos, System.Globalization.NumberStyles.Integer,
                    System.Globalization.CultureInfo.InvariantCulture, out var mgPosHp))
                Emit("position", $"{mgPosHp / 2.0:0.##}pt");
            // Character border — same colon-encoded compound form as the run
            // reader (STYLE;SIZE;COLOR;SPACE[;SHADOW[;FRAME]]).
            var mgBdr = src.GetFirstChild<Border>();
            if (mgBdr?.Val?.HasValue == true)
            {
                var segs = new List<string>
                {
                    mgBdr.Val!.InnerText ?? "single",
                    mgBdr.Size?.Value.ToString() ?? "",
                    string.IsNullOrEmpty(mgBdr.Color?.Value) ? "" : ParseHelpers.FormatHexColor(mgBdr.Color!.Value!),
                    mgBdr.Space?.Value.ToString() ?? "0"
                };
                bool? mgShadow = mgBdr.Shadow?.Value;
                bool? mgFrame = mgBdr.Frame?.Value;
                if (mgShadow.HasValue || mgFrame.HasValue)
                    segs.Add(mgShadow == true ? "true" : "false");
                if (mgFrame.HasValue)
                    segs.Add(mgFrame == true ? "true" : "false");
                Emit("bdr", string.Join(';', segs));
            }
        }
        // Suppress markRPr.* dotted keys when the paragraph has no
        // text-bearing runs — the bare keys below (size, font.latin, …)
        // already cover markRPr via the firstRun-fallback path. Emitting
        // both forms on an empty paragraph means dump→batch→dump
        // surfaces phantom markRPr.* keys even after AddParagraph
        // routed the formatting correctly (BUG-DUMP-MARKRPR-DOUBLE).
        // The dotted form's purpose is to distinguish the ¶ glyph's
        // formatting from the visible text — only meaningful when text
        // runs exist.
        var hasTextRun = para.Elements<Run>()
            .Any(r => r.GetFirstChild<Text>() != null
                      && !string.IsNullOrEmpty(r.GetFirstChild<Text>()?.Text));
        // BUG-DUMP-R27-2: a paragraph whose ONLY run is an inline drawing/
        // picture (no <w:t> text) has hasTextRun == false, so the dotted
        // block was suppressed. But the bare-key firstRun-fallback below
        // ALSO can't carry the ¶-mark formatting: firstRun (a run with a
        // <w:t>) is null, so it falls back to markRp and emits BARE keys
        // (highlight / font.ea / …) onto the paragraph node — which the
        // single-picture-run emit path then drops (the picture op carries
        // its own font.ea but silently loses highlight). The para-mark
        // formatting on a non-text run paragraph is just as distinct from
        // the (non-textual) run content as in the text-run case, so emit
        // the dotted form here too. The firstRun-fallback markRp branch is
        // narrowed (below) to fire only when NO runs exist at all, so the
        // two forms stay mutually exclusive (no DOUBLE).
        // BUG-DUMP-MARKRPR-HYPERLINK: count runs nested in hyperlinks/SDTs/
        // smartTags too, not just direct-child runs. A cell paragraph whose
        // sole content is a hyperlink (a language-link cell, "EN"/"SP", …) has
        // NO direct <w:r> child, so hasAnyRun was false and the dotted markRPr.*
        // block below was skipped — yet the firstRun-fallback also can't carry
        // it (firstRun is null but the bare-key path is gated on !hasAnyRun via
        // a still-direct-children check), so the ¶-mark <w:rPr> (the font/size
        // that sets the cell line height) was dropped entirely on round-trip,
        // collapsing the line and drifting the table. Descendants<Run> makes the
        // dotted form fire so the mark rPr round-trips; the bare-key fallback
        // (gated on its own !hasAnyRun below) stays off, so no DOUBLE emit.
        var hasAnyRun = para.Descendants<Run>().Any();
        if (pmrpForDump != null && (hasTextRun || hasAnyRun))
        {
            var b = pmrpForDump.GetFirstChild<Bold>();
            if (b != null) node.Format["markRPr.bold"] = IsToggleOn(b);
            var i = pmrpForDump.GetFirstChild<Italic>();
            if (i != null) node.Format["markRPr.italic"] = IsToggleOn(i);
            var s = pmrpForDump.GetFirstChild<Strike>();
            if (s != null) node.Format["markRPr.strike"] = IsToggleOn(s);
            var u = pmrpForDump.GetFirstChild<Underline>();
            if (u?.Val?.HasValue == true) node.Format["markRPr.underline"] = u.Val.InnerText;
            var fs = pmrpForDump.GetFirstChild<FontSize>();
            if (fs?.Val?.Value != null)
                node.Format["markRPr.size"] = $"{int.Parse(fs.Val.Value) / 2.0:0.##}pt";
            // BUG-DUMP-MARKRPR-CS: the ¶-mark's complex-script + kern slots
            // (szCs / bCs / iCs / kern) were never emitted under markRPr.*, so
            // dump→batch silently dropped them on every paragraph that has text
            // runs (the empty-paragraph fallback path already reads them via
            // ReadComplexScriptRunFormatting). Mirror that readback here so the
            // ¶ glyph's CJK metrics survive the round-trip. ApplyRunFormatting
            // consumes size.cs / bold.cs / italic.cs / kern on replay.
            var fsCs = pmrpForDump.GetFirstChild<FontSizeComplexScript>();
            if (fsCs?.Val?.Value is string fsCsVal && int.TryParse(fsCsVal, out var fsCsHp))
                node.Format["markRPr.size.cs"] = $"{fsCsHp / 2.0:0.##}pt";
            var bCsMark = pmrpForDump.GetFirstChild<BoldComplexScript>();
            if (bCsMark != null && (bCsMark.Val == null || bCsMark.Val.Value))
                node.Format["markRPr.bold.cs"] = true;
            var iCsMark = pmrpForDump.GetFirstChild<ItalicComplexScript>();
            if (iCsMark != null && (iCsMark.Val == null || iCsMark.Val.Value))
                node.Format["markRPr.italic.cs"] = true;
            var kernMark = pmrpForDump.GetFirstChild<Kern>();
            if (kernMark?.Val?.HasValue == true)
                node.Format["markRPr.kern"] = kernMark.Val.Value.ToString();
            var clr = pmrpForDump.GetFirstChild<Color>();
            // BUG-DUMP-R44-1: paragraph-mark color must carry BOTH the hex val
            // AND the theme linkage via the shared ';themeColor=…' tail, same as
            // the direct run color path — emitting only the theme name dropped an
            // explicit w:val (e.g. <w:color w:val="1F497D" w:themeColor="text2"/>
            // rebuilt as val="auto").
            if (clr != null && StyleColorWithThemeTail(clr) is { } markClr)
                node.Format["markRPr.color"] = markClr;
            var rf = pmrpForDump.GetFirstChild<RunFonts>();
            // BUG-DUMP-MARKRPR-HANSI / BUG-DUMP-FONT-LATIN: collapse to
            // font.latin only when BOTH slots are present and equal. Divergent
            // slots emit both; ascii-only emits font.ascii and hAnsi-only
            // emits font.hAnsi — NEVER font.latin for a single slot, since
            // font.latin sets both on replay and would silently rebind the ¶
            // mark's missing slot (extended-Latin falls back to the
            // style/docDefaults font when hAnsi is absent).
            var markAscii = rf?.Ascii?.Value;
            var markHAnsi = rf?.HighAnsi?.Value;
            if (markAscii != null && markHAnsi != null)
            {
                if (markAscii == markHAnsi)
                    node.Format["markRPr.font.latin"] = markAscii;
                else
                {
                    node.Format["markRPr.font.ascii"] = markAscii;
                    node.Format["markRPr.font.hAnsi"] = markHAnsi;
                }
            }
            else if (markAscii != null)
                node.Format["markRPr.font.ascii"] = markAscii;
            else if (markHAnsi != null)
                node.Format["markRPr.font.hAnsi"] = markHAnsi;
            if (rf?.EastAsia?.Value != null)
                node.Format["markRPr.font.ea"] = rf.EastAsia.Value;
            if (rf?.ComplexScript?.Value != null)
                node.Format["markRPr.font.cs"] = rf.ComplexScript.Value;
            // Theme-bound slots (<w:rFonts w:asciiTheme="minorHAnsi" …/>):
            // the ¶ mark's font sets the line height of an empty spacer
            // paragraph, so dropping a theme binding (mark renders in the
            // docDefaults face instead of the theme face) changes each
            // spacer's height slightly and the accumulated drift reflows
            // page breaks across the whole document.
            if (rf?.AsciiTheme?.HasValue == true)
                node.Format["markRPr.font.asciiTheme"] = rf.AsciiTheme.InnerText;
            if (rf?.HighAnsiTheme?.HasValue == true)
                node.Format["markRPr.font.hAnsiTheme"] = rf.HighAnsiTheme.InnerText;
            if (rf?.EastAsiaTheme?.HasValue == true)
                node.Format["markRPr.font.eaTheme"] = rf.EastAsiaTheme.InnerText;
            if (rf?.ComplexScriptTheme?.HasValue == true)
                node.Format["markRPr.font.csTheme"] = rf.ComplexScriptTheme.InnerText;
            // ¶-mark font hint + character spacing (mirror the run-level
            // font.hint / charSpacing readback so the paragraph mark's glyph
            // properties round-trip too — see RunToNode).
            if (rf?.Hint?.HasValue == true)
                node.Format["markRPr.font.hint"] = rf.Hint.InnerText;
            var pmSpacing = pmrpForDump.GetFirstChild<Spacing>();
            if (pmSpacing?.Val?.HasValue == true)
                node.Format["markRPr.charSpacing"] = $"{pmSpacing.Val.Value / 20.0:0.##}pt";
            // BUG-DUMP-R41-3: ¶-mark vertical position (<w:pPr><w:rPr><w:position
            // w:val="-10"/>). The run-level <w:position> already round-trips (see
            // RunToNode → Format["position"]); the markRPr whitelist surfaced
            // size/color/kern/charSpacing/… but never <w:position>, so a paragraph
            // mark's raise/lower was silently dropped on the dump → batch round-trip.
            // Emit the raw half-point val under markRPr.position; the general
            // markRPr.* dispatch routes it through ApplyRunFormatting's "position"
            // case on replay (Add.Text.cs), mirroring the run-level readback.
            var pmPos = pmrpForDump.GetFirstChild<Position>();
            if (pmPos?.Val?.Value is string pmPosVal && !string.IsNullOrEmpty(pmPosVal))
                node.Format["markRPr.position"] = pmPosVal;
            // ¶-mark <w:rtl/> (mark-only RTL, no pPr <w:bidi/>) — see the
            // empty-paragraph fallback for the rationale; same dotted key so
            // ApplyRunFormatting's rtl case restores it without touching the
            // paragraph direction cascade.
            var pmRtl = pmrpForDump.GetFirstChild<RightToLeftText>();
            if (pmRtl != null)
                node.Format["markRPr.rtl"] = TryReadOnOff(pmRtl.Val) != false;
            // BUG-DUMP-MARKRPR-SNAPGRID: the ¶-mark's <w:snapToGrid> toggle. On a
            // doc with a <w:docGrid>, the mark's snapToGrid="0" keeps the
            // terminating line off the grid (sets its height); the markRPr
            // allowlist surfaced bold/size/color/position/rtl/… but never
            // snapToGrid, so dump→batch dropped it and the line re-snapped to the
            // grid, shifting metrics and reflowing the page. Emit a canonical bool
            // (CT_OnOff); ApplyRunFormatting's snapToGrid case restores the
            // explicit OFF form on replay.
            var pmSnap = pmrpForDump.GetFirstChild<SnapToGrid>();
            if (pmSnap != null)
                node.Format["markRPr.snapToGrid"] = TryReadOnOff(pmSnap.Val) != false;
            // Glyph props historically absent from this whitelist (caps family,
            // text effects, vertAlign, bdr, vanish, specVanish) — see the
            // shared readback above.
            ReadMarkGlyphProps(pmrpForDump, "markRPr.");
            var hl = pmrpForDump.GetFirstChild<Highlight>();
            if (hl?.Val?.HasValue == true) node.Format["markRPr.highlight"] = hl.Val.InnerText;
            // BUG-DUMP-R27-1: ¶-mark character shading (<w:pPr><w:rPr><w:shd/>).
            // Run-level shd already round-trips (shading.val/.fill/.color folded
            // into a single `shading` key); the markRPr whitelist supported
            // bold/color/kern/size/underline/font.* but not shd, so para-mark
            // shading was silently dropped. Emit the same semicolon-encoded
            // VAL;FILL[;COLOR] form `markRPr.highlight` uses (single key) so it
            // routes through ApplyRunFormatting's `shading` case on replay.
            var pmShd = pmrpForDump.GetFirstChild<Shading>();
            if (pmShd != null)
            {
                var pmShdVal = pmShd.Val?.InnerText;
                var pmShdFill = pmShd.Fill?.Value;
                var pmShdColor = pmShd.Color?.Value;
                // Skip the OOXML "no shading" form (clear + no fill/color) —
                // mirrors the FilterEmittableProps shading-fold drop.
                bool effectivelyNone = string.Equals(pmShdVal, "clear", StringComparison.OrdinalIgnoreCase)
                    && string.IsNullOrEmpty(pmShdFill) && string.IsNullOrEmpty(pmShdColor);
                if (!effectivelyNone)
                {
                    var v = string.IsNullOrEmpty(pmShdVal) ? "clear" : pmShdVal;
                    string folded;
                    if (!string.IsNullOrEmpty(pmShdColor))
                        folded = $"{v};{(string.IsNullOrEmpty(pmShdFill) ? "" : ParseHelpers.FormatHexColor(pmShdFill))};{ParseHelpers.FormatHexColor(pmShdColor)}";
                    else if (!string.IsNullOrEmpty(pmShdFill))
                        folded = $"{v};{ParseHelpers.FormatHexColor(pmShdFill)}";
                    else
                        folded = v;
                    // BUG-DUMP-R41-4: carry the ¶-mark shd theme linkage as
                    // key=val tails (mirrors the run/paragraph/cell shading
                    // fold); ParseShadingValue strips them on replay.
                    if (pmShd.ThemeFill?.HasValue == true) folded += $";themeFill={pmShd.ThemeFill.InnerText}";
                    if (pmShd.ThemeFillShade?.Value is { } pmTfs) folded += $";themeFillShade={pmTfs}";
                    if (pmShd.ThemeFillTint?.Value is { } pmTft) folded += $";themeFillTint={pmTft}";
                    if (pmShd.ThemeColor?.HasValue == true) folded += $";themeColor={pmShd.ThemeColor.InnerText}";
                    if (pmShd.ThemeShade?.Value is { } pmTsh) folded += $";themeShade={pmTsh}";
                    if (pmShd.ThemeTint?.Value is { } pmTt) folded += $";themeTint={pmTt}";
                    node.Format["markRPr.shading"] = folded;
                }
            }
            // BUG-DUMP-MARKRPR-LANG: the ¶-mark's <w:lang> slots (val=latin /
            // eastAsia / bidi=cs) were never emitted under markRPr.*, so
            // dump→batch silently dropped them on every text-bearing paragraph.
            // Run-level lang already round-trips as lang.latin/lang.ea/lang.cs
            // (see RunToNode); mirror that here. ApplyRunFormatting consumes
            // markRPr.lang.latin / .ea / .cs on replay (Add.Text.cs markRPr.*
            // dispatch → ApplyRunFormatting handles lang.* multi-slot).
            var pmLang = pmrpForDump.GetFirstChild<Languages>();
            if (pmLang != null)
            {
                if (pmLang.Val?.Value != null) node.Format["markRPr.lang.latin"] = pmLang.Val.Value;
                if (pmLang.EastAsia?.Value != null) node.Format["markRPr.lang.ea"] = pmLang.EastAsia.Value;
                if (pmLang.Bidi?.Value != null) node.Format["markRPr.lang.cs"] = pmLang.Bidi.Value;
            }
            // schemas/help/docx/paragraph.json declares rStyle add+set+get;
            // Add.Text.cs:437 writes <w:rStyle> into ParagraphMarkRunProperties,
            // but Get used to drop it. Emit at the paragraph-level canonical
            // key (no markRPr prefix) to match the schema's declaration.
            var rs = pmrpForDump.GetFirstChild<RunStyle>();
            if (rs?.Val?.Value != null)
            {
                // Bare `rStyle` on add-paragraph styles BOTH the mark and the
                // implicit text run (BUG-R6-03). When the source carries the
                // style on the MARK ONLY (a quote paragraph whose runs stay
                // on the paragraph style), echoing it bare would restyle the
                // rebuilt text run and override the paragraph style's italic/
                // color. Emit the mark-only shape under markRPr.rStyle, which
                // targets ParagraphMarkRunProperties exclusively on replay.
                var rsFirstTextRun = para.Elements<Run>()
                    .FirstOrDefault(r => r.GetFirstChild<Text>() != null);
                var rsRunVal = rsFirstTextRun?.RunProperties?.GetFirstChild<RunStyle>()?.Val?.Value;
                if (string.Equals(rsRunVal, rs.Val.Value, StringComparison.Ordinal))
                    node.Format["rStyle"] = rs.Val.Value;
                else
                    node.Format["markRPr.rStyle"] = rs.Val.Value;
            }
            // BUG-DUMP-MARKRPR-VERBATIM (class fix): the dotted markRPr.* keys
            // above are a hardcoded ALLOWLIST — any ¶-mark rPr child not on it
            // (w:em CJK emphasis, w:effect, w:w letter-scaling, the w14:*
            // OpenType-extension elements, …) was silently dropped on round-trip.
            // Emit the WHOLE ¶-mark <w:rPr> verbatim as a single key so EVERY
            // property survives; AddParagraph applies it as the authoritative
            // mark rPr and skips the per-property dotted apply (the dotted keys
            // stay emitted for human/other-consumer readability but are inert on
            // replay when markRPr.xml is present). Mirrors AddFootnote's verbatim
            // referenceMarkRPr. The revision paraMarkIns/Del markers live in a
            // SEPARATE namespace (not <w:rPr> children) and are unaffected.
            // OuterXml keeps the xmlns:w (and any w14:/mc:) declarations the
            // standalone fragment needs to re-parse on replay (new
            // ParagraphMarkRunProperties(xml) drops children whose prefix can't
            // resolve, so the declaration must stay). The redundant xmlns:w is
            // cosmetic bloat but harmless and idempotent (the next dump re-emits
            // the same OuterXml). Do NOT strip it — stripping breaks the apply.
            if (pmrpForDump.HasChildren)
                node.Format["markRPr.xml"] = pmrpForDump.OuterXml;
        }

        // First-run formatting on the paragraph node (like PPTX does for shapes).
        // Fall back to ParagraphMarkRunProperties when no runs exist (e.g. empty paragraph
        // that had formatting applied via Set before any text was added).
        var firstRun = para.Elements<Run>().FirstOrDefault(r => r.GetFirstChild<Text>() != null);
        // BUG-DUMP-R27-2: only fall back to ParagraphMarkRunProperties for the
        // bare-key path when the paragraph has NO runs at all (truly empty ¶).
        // A picture/drawing-only paragraph has runs but no text run; its
        // ¶-mark formatting is now carried by the dotted markRPr.* block above,
        // so emitting bare keys here too would double-emit (and the bare
        // highlight would be dropped by the single-picture-run emit anyway).
        var paraRp = firstRun?.RunProperties
            ?? (firstRun == null && !hasAnyRun ? para.ParagraphProperties?.ParagraphMarkRunProperties as OpenXmlCompositeElement : null);
        if (paraRp != null)
        {
            RunProperties? rp = paraRp as RunProperties ?? null;
            ParagraphMarkRunProperties? markRp = paraRp as ParagraphMarkRunProperties ?? null;

            // BUG-R12C: an empty paragraph's ¶-mark rStyle. The dotted
            // markRPr.* block above is suppressed when the paragraph has no
            // runs, and it is the bare-key fallback (this block, markRp != null)
            // that must carry the ¶-glyph formatting — but it never surfaced
            // <w:rStyle>. The referenced character style's size sets the empty
            // paragraph's line height; a title-page spacer styled BookTitle
            // (16pt) collapsed to the default Normal height on dump→batch,
            // shifting every block below it upward. Emit bare `rStyle` (a bare
            // rStyle on a run-less paragraph routes to ParagraphMarkRunProperties
            // only — see the dotted-block note above and Add.Text.cs).
            if (markRp != null
                && markRp.GetFirstChild<RunStyle>()?.Val?.Value is { } emptyParaMarkRStyle
                && !node.Format.ContainsKey("rStyle")
                && !node.Format.ContainsKey("markRPr.rStyle"))
                node.Format["rStyle"] = emptyParaMarkRStyle;

            // CONSISTENCY(canonical-keys): mirror style Get (WordHandler.Query.cs:546-553) —
            // emit per-script font slots, no flat "font" alias. R6 BUG-1: previously only
            // emitted Ascii under "font" key, dropping eastAsia/hAnsi/cs slots.
            var pRunFonts = rp?.RunFonts ?? markRp?.GetFirstChild<RunFonts>();
            if (pRunFonts != null)
            {
                // CONSISTENCY(canonical-keys): schema (docx/run.json,
                // docx/paragraph.json) declares `font.latin` and `font.ea`
                // as canonical. Collapse Ascii+HighAnsi to `font.latin`
                // when they match (the round-trip case for `font.latin=`
                // Set). When they differ, emit both legacy slots so no
                // information is lost.
                var ascii = pRunFonts.Ascii?.Value;
                var hAnsi = pRunFonts.HighAnsi?.Value;
                if (ascii != null && hAnsi != null && ascii == hAnsi)
                {
                    if (!node.Format.ContainsKey("font.latin"))
                        node.Format["font.latin"] = ascii;
                }
                else if (ascii != null && hAnsi != null)
                {
                    // Two slots, divergent values — fall back to legacy keys.
                    if (!node.Format.ContainsKey("font.ascii"))
                        node.Format["font.ascii"] = ascii;
                    if (!node.Format.ContainsKey("font.hAnsi"))
                        node.Format["font.hAnsi"] = hAnsi;
                }
                // BUG-DUMP-FONT-LATIN: ascii-only → font.ascii, hAnsi-only →
                // font.hAnsi (NOT font.latin, which sets both on replay and
                // would rebind the absent slot).
                else if (ascii != null)
                {
                    if (!node.Format.ContainsKey("font.ascii"))
                        node.Format["font.ascii"] = ascii;
                }
                else if (hAnsi != null)
                {
                    if (!node.Format.ContainsKey("font.hAnsi"))
                        node.Format["font.hAnsi"] = hAnsi;
                }
                if (!string.IsNullOrEmpty(pRunFonts.EastAsia?.Value) && !node.Format.ContainsKey("font.ea"))
                    node.Format["font.ea"] = pRunFonts.EastAsia!.Value!;
                // BUG-DUMP15-03: surface theme-font slots on the paragraph
                // node (leaked from first run rPr) so dump→batch round-trip
                // preserves theme bindings. Mirrors the run-level readback
                // at the typed-Run branch below.
                if (pRunFonts.AsciiTheme?.HasValue == true && !node.Format.ContainsKey("font.asciiTheme"))
                    node.Format["font.asciiTheme"] = pRunFonts.AsciiTheme.InnerText;
                if (pRunFonts.HighAnsiTheme?.HasValue == true && !node.Format.ContainsKey("font.hAnsiTheme"))
                    node.Format["font.hAnsiTheme"] = pRunFonts.HighAnsiTheme.InnerText;
                if (pRunFonts.EastAsiaTheme?.HasValue == true && !node.Format.ContainsKey("font.eaTheme"))
                    node.Format["font.eaTheme"] = pRunFonts.EastAsiaTheme.InnerText;
                if (pRunFonts.ComplexScriptTheme?.HasValue == true && !node.Format.ContainsKey("font.csTheme"))
                    node.Format["font.csTheme"] = pRunFonts.ComplexScriptTheme.InnerText;
                // BUG-DUMP-R31-2: the bare-key firstRun/empty-¶ fallback emitted
                // every rFonts slot (latin/ea/themes) EXCEPT the <w:hint> font-
                // slot selector. An empty paragraph whose ¶-mark rPr carried
                // <w:rFonts w:hint="eastAsia"/> (and, by the same path, a single-
                // run paragraph hoisting its first run's rFonts) therefore lost
                // the hint on dump → it never round-tripped. w:hint selects which
                // font slot renders boundary CJK glyphs; dropping it can rebind
                // the glyph to the wrong physical font. Mirror the dotted-markRPr
                // and RunToNode emits — AddParagraph's bare-key path routes
                // font.hint back through ApplyRunFormatting on replay.
                if (pRunFonts.Hint?.HasValue == true && !node.Format.ContainsKey("font.hint"))
                    node.Format["font.hint"] = pRunFonts.Hint.InnerText;
            }

            var fsVal = rp?.FontSize?.Val?.Value ?? markRp?.GetFirstChild<FontSize>()?.Val?.Value;
            if (fsVal != null && !node.Format.ContainsKey("size"))
                node.Format["size"] = $"{int.Parse(fsVal) / 2.0:0.##}pt";

            var boldEl = rp?.Bold ?? markRp?.GetFirstChild<Bold>();
            if (boldEl != null && !node.Format.ContainsKey("bold")) node.Format["bold"] = IsToggleOn(boldEl);

            var italicEl = rp?.Italic ?? markRp?.GetFirstChild<Italic>();
            if (italicEl != null && !node.Format.ContainsKey("italic")) node.Format["italic"] = IsToggleOn(italicEl);

            // Complex-script readback (font.cs / size.cs / bold.cs / italic.cs).
            // See WordHandler.I18n.cs.
            ReadComplexScriptRunFormatting(rp, markRp, node.Format);

            var colorEl = rp?.Color ?? markRp?.GetFirstChild<Color>();
            if (colorEl != null && !node.Format.ContainsKey("color"))
            {
                // BUG-DUMP-R47-4: a single-run paragraph collapsed into `add p`
                // must carry the run color's FULL theme linkage (hex val +
                // themeColor + themeTint/themeShade) via the shared
                // ';themeColor=…' tail — same as the un-collapsed run path
                // (RunToNode StyleColorWithThemeTail) and the ¶-mark path. The
                // old code emitted only the theme name, so a run color like
                // <w:color w:val="548DD4" w:themeColor="text2" w:themeTint="99"/>
                // rebuilt as val="auto" + themeColor with no tint — a visibly
                // different (untinted) color on round-trip.
                if (StyleColorWithThemeTail(colorEl) is { } pFirstRunColor)
                    node.Format["color"] = pFirstRunColor;
            }

            var ulEl = rp?.Underline ?? markRp?.GetFirstChild<Underline>();
            if (ulEl?.Val != null && !node.Format.ContainsKey("underline"))
                node.Format["underline"] = ulEl.Val.InnerText;
            // CONSISTENCY(underline-color): backfilled from style Get edc8f884.
            if (ulEl?.Color?.Value != null && !node.Format.ContainsKey("underline.color"))
                node.Format["underline.color"] = ParseHelpers.FormatHexColor(ulEl.Color.Value);

            var strikeEl = rp?.Strike ?? (OpenXmlLeafElement?)markRp?.GetFirstChild<Strike>();
            if (strikeEl != null && !node.Format.ContainsKey("strike")) node.Format["strike"] = true;

            // BUG-DUMP-R62-MARKVANISH: an empty paragraph's ¶-mark <w:vanish/>
            // (hidden text) makes the paragraph zero-height. The canonical case is
            // the spacer paragraph Word inserts between two adjacent tables: marked
            // vanish, it collapses so the tables render flush; un-marked, it renders
            // at its ¶-glyph height and opens a visible gap that pushes every row
            // below down (a whole-document reflow on a form-heavy template). The
            // bare-key fallback read every other mark toggle (bold/italic/strike/
            // color/kern) but not vanish, so the spacer reappeared on dump→batch.
            // Mirror the run reader's vanish readback; AddParagraph's bare-key path
            // applies it to the ¶ mark rPr (run-less) or the run (single-run).
            var vanishEl = rp?.Vanish ?? (Vanish?)markRp?.GetFirstChild<Vanish>();
            if (vanishEl != null && IsToggleOn(vanishEl) && !node.Format.ContainsKey("vanish"))
                node.Format["vanish"] = true;

            var hlEl = rp?.Highlight ?? markRp?.GetFirstChild<Highlight>();
            if (hlEl?.Val != null && !node.Format.ContainsKey("highlight"))
                node.Format["highlight"] = hlEl.Val.InnerText;

            // BUG-DUMP-MARKRPR-CS: empty-paragraph ¶-mark kern. Text
            // paragraphs route the mark's kern through markRPr.kern above;
            // here we only cover the empty-paragraph case (markRp != null
            // ⇒ no text run) so we don't hoist a first run's kern onto the
            // paragraph node. AddParagraph's bare-key path applies it back
            // to the ¶ mark rPr on replay.
            if (markRp != null && !node.Format.ContainsKey("kern"))
            {
                var kEmpty = markRp.GetFirstChild<Kern>();
                if (kEmpty?.Val?.HasValue == true)
                    node.Format["kern"] = kEmpty.Val.Value.ToString();
            }
            // BUG-DUMP-R27-1 / BUG-DUMP-R44-3: empty-paragraph ¶-mark character
            // shading (<w:shd> inside ParagraphMarkRunProperties). This MUST be
            // emitted under the EXPLICIT `markRPr.shading` key (same as the
            // text/picture path in the pmrpForDump block above), NOT the bare
            // `shading` key. Bare `shading` is now unambiguously a paragraph-level
            // DIRECT <w:pPr><w:shd> (whole-line banner) on replay — emitting it
            // here would relocate a genuine ¶-mark (pilcrow-only) shading onto the
            // whole-line background, AND conversely a genuine direct pPr/shd on a
            // no-text banner paragraph (e.g. an empty full-width banner bar) would
            // be wrongly pushed down to the ¶ mark. The `markRPr.shading` key
            // routes through AddParagraph's markRPr.* branch to the mark rPr.
            if (markRp != null && !node.Format.ContainsKey("markRPr.shading"))
            {
                var shdEmpty = markRp.GetFirstChild<Shading>();
                if (shdEmpty != null)
                {
                    var sv = shdEmpty.Val?.InnerText;
                    var sf = shdEmpty.Fill?.Value;
                    var sc = shdEmpty.Color?.Value;
                    bool none = string.Equals(sv, "clear", StringComparison.OrdinalIgnoreCase)
                        && string.IsNullOrEmpty(sf) && string.IsNullOrEmpty(sc);
                    if (!none)
                    {
                        var v = string.IsNullOrEmpty(sv) ? "clear" : sv;
                        if (!string.IsNullOrEmpty(sc))
                            node.Format["markRPr.shading"] = $"{v};{(string.IsNullOrEmpty(sf) ? "" : ParseHelpers.FormatHexColor(sf))};{ParseHelpers.FormatHexColor(sc)}";
                        else if (!string.IsNullOrEmpty(sf))
                            node.Format["markRPr.shading"] = $"{v};{ParseHelpers.FormatHexColor(sf)}";
                        else
                            node.Format["markRPr.shading"] = v;
                    }
                }
            }
            // BUG-R10A(BUG3): empty-paragraph ¶-mark character spacing
            // (<w:spacing> inside ParagraphMarkRunProperties). The
            // text-bearing-paragraph path surfaces this as
            // `markRPr.charSpacing` (see pmrpForDump block above), but the
            // empty-paragraph fallback handled neither bare `charSpacing` nor
            // `markRPr.charSpacing` — so `set --prop markRPr.charSpacing=2pt`
            // on an empty paragraph wrote the <w:spacing> correctly yet `get`
            // never surfaced it and dump→batch dropped it. Use the same
            // canonical key as the text-para path so replay rebuilds the
            // <w:spacing> via AddParagraph's markRPr.charSpacing handling.
            if (markRp != null && !node.Format.ContainsKey("markRPr.charSpacing"))
            {
                var spEmpty = markRp.GetFirstChild<Spacing>();
                if (spEmpty?.Val?.HasValue == true)
                    node.Format["markRPr.charSpacing"] = $"{spEmpty.Val.Value / 20.0:0.##}pt";
            }
            // BUG-DUMP-MARKRPR-LANG: empty-paragraph ¶-mark <w:lang> slots.
            // Text-bearing paragraphs route the mark's lang through
            // markRPr.lang.* (see pmrpForDump block above); the empty-paragraph
            // fallback emits the bare lang.latin/lang.ea/lang.cs keys, which
            // AddParagraph applies back to the ¶ mark rPr on replay (no text
            // run ⇒ the lang switch targets ParagraphMarkRunProperties).
            if (markRp != null)
            {
                var langEmpty = markRp.GetFirstChild<Languages>();
                if (langEmpty != null)
                {
                    if (langEmpty.Val?.Value != null && !node.Format.ContainsKey("lang.latin"))
                        node.Format["lang.latin"] = langEmpty.Val.Value;
                    if (langEmpty.EastAsia?.Value != null && !node.Format.ContainsKey("lang.ea"))
                        node.Format["lang.ea"] = langEmpty.EastAsia.Value;
                    if (langEmpty.Bidi?.Value != null && !node.Format.ContainsKey("lang.cs"))
                        node.Format["lang.cs"] = langEmpty.Bidi.Value;
                }
                // ¶-mark <w:rtl/> on a run-less paragraph (Arabic forms carry
                // it on empty spacer paragraphs WITHOUT pPr <w:bidi/>). The
                // markRPr whitelist never surfaced the slot, so dozens of marks
                // lost their RTL on dump→batch and the form reflowed. Dotted
                // key — the bare `rtl` key would route through the paragraph
                // direction cascade and fabricate a <w:bidi/> the source
                // doesn't have.
                var rtlEmpty = markRp.GetFirstChild<RightToLeftText>();
                if (rtlEmpty != null && !node.Format.ContainsKey("markRPr.rtl"))
                    node.Format["markRPr.rtl"] = TryReadOnOff(rtlEmpty.Val) != false;
            }
            // Glyph props historically absent from the bare-key fallback
            // (caps family, text effects, vertAlign, bdr, specVanish) — the
            // RUN-LESS ¶ mark only (markRp != null). The firstRun hoist must
            // NOT surface them as bare paragraph keys: run-level content is
            // emitted separately (run/field ops carry their own readback), so
            // hoisting would double-apply — replay's no-text path writes the
            // hoisted copy onto the ¶ mark and the XML gains a phantom
            // duplicate (e.g. a 6th vertAlign on a 5-run superscript field).
            if (markRp != null)
                ReadMarkGlyphProps(markRp, "");
        }

        // Populate effective.* properties from style inheritance
        PopulateEffectiveParagraphProperties(node, para);

        if (depth > 0)
        {
            // BUG-DUMP13-02: interleave typed Runs and inline M.OfficeMath
            // equations in DOM order so paragraphs like `r1 / m:oMath / r2`
            // emit r1, equation, r2 (not r1, r2, equation). Previously
            // GetAllRuns appended every run first and the inline-equation
            // loop below appended all equations afterwards as a separate
            // group, so DOM order was lost on dump round-trip.
            //
            // We compute a DOM-position index per element via a single
            // descendant walk (Descendants() yields document order) and
            // use it to sort only the run+equation slice, leaving other
            // categories (sdt/bookmark/field/etc.) in their original
            // append order.
            int runIdx = 0;
            int inlineEqIdx = 0;
            var descendantPos = new Dictionary<OpenXmlElement, int>(ReferenceEqualityComparer.Instance);
            int dpi = 0;
            foreach (var d in para.Descendants())
                descendantPos[d] = dpi++;
            // BUG-DUMP-R29-SMARTTAG: document-position of each child appended to
            // node.Children, parallel to the list. The smartTag/customXml wrapper
            // runs are synthesized in a separate pass below and were appended at
            // the tail — so a mid-paragraph <w:smartTag> run ("PRICE <AND> TERMS")
            // replayed at the END of the paragraph. Tracking positions lets that
            // pass INSERT each wrapper run at its true document order instead.
            var childPositions = new List<int>();

            var runs = GetAllRuns(para);
            // BUG-DUMP9-04: m:oMath nested inside w:hyperlink is a
            // grandchild of the paragraph and was silently dropped.
            // BUG-DUMP8-03: include m:oMath nested inside w:ins/w:del
            // change-track wrappers — they are paragraph grandchildren,
            // not direct children, and were silently dropped on dump.
            var inlineEqsAll = para.Elements<M.OfficeMath>()
                .Concat(para.Elements<InsertedRun>().SelectMany(ins => ins.Elements<M.OfficeMath>()))
                .Concat(para.Elements<DeletedRun>().SelectMany(del => del.Elements<M.OfficeMath>()))
                .Concat(para.Elements<Hyperlink>().SelectMany(hl => hl.Elements<M.OfficeMath>()))
                .ToList();
            // BUG-DUMP15-04: paragraph hyperlink children for hyperlink-
            // scoped equation paths. m:oMath inside w:hyperlink must
            // surface as /…/p[N]/hyperlink[K]/equation[M] so dump→batch
            // replays the equation INSIDE the hyperlink rather than
            // alongside it. Index hyperlinks by their position among
            // the paragraph's direct Hyperlink children.
            var paraHyperlinks = para.Elements<Hyperlink>().ToList();

            // Merge runs and inline equations by DOM position, then emit
            // in that interleaved order.
            // BUG-DUMP15-02: bare <w:fldChar>/<w:instrText> direct children
            // of <w:p> (not wrapped in a <w:r>) are parsed as
            // OpenXmlUnknownElement and silently dropped from the children
            // list, which left CollapseFieldChains nothing to stitch and
            // dump→batch round-trips lost the entire HYPERLINK chain.
            // Surface them as synthetic fieldChar/instrText nodes so the
            // emitter can collapse them into a `field` row.
            const string wNs2 = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            var bareFieldUnknowns = para.Elements<DocumentFormat.OpenXml.OpenXmlUnknownElement>()
                .Where(u => u.NamespaceUri == wNs2
                    && (u.LocalName == "fldChar" || u.LocalName == "instrText"))
                .ToList();
            // BUG-DUMP25-01: include direct-child BookmarkStart elements in
            // the DOM-ordered merge so a bookmark sitting between two runs
            // surfaces as `r, bookmark, r` rather than the legacy
            // `r, r, bookmark` (every bookmark hoisted to the tail of
            // node.Children). The trailing standalone bookmark loop below
            // is now skipped when this branch surfaces them.
            // BUG-DUMP-BMINDEL: a bookmark can live INSIDE a revision wrapper
            // (<w:ins>/<w:del>) — e.g. a PAGEREF/REF target bookmark sitting in a
            // tracked-deleted run. Those are NOT direct <w:p> children, so the
            // bare Elements<BookmarkStart>() walk dropped them and every field
            // pointing at one rendered "Error! Bookmark not defined." Include the
            // ins/del-nested starts too (mirrors inlineEqsAll above); descendantPos
            // already positions them by DOM order.
            // BUG-DUMP-HYPERLINK-BOOKMARK: a <w:bookmarkStart> can sit INSIDE a
            // <w:hyperlink> (a cross-reference target placed on a linked phrase) —
            // a paragraph grandchild, so the bare Elements<BookmarkStart>() walk
            // dropped it and every PAGEREF/REF pointing at it rebuilt as a dangling
            // "Error! Bookmark not defined." Surface the hyperlink-nested starts too
            // (mirrors the ins/del-nested handling above); descendantPos positions
            // them by DOM order and the bookmark emit replays them at that offset.
            var paraBookmarks = para.Elements<BookmarkStart>()
                .Concat(para.Elements<InsertedRun>().SelectMany(ins => ins.Elements<BookmarkStart>()))
                .Concat(para.Elements<DeletedRun>().SelectMany(del => del.Elements<BookmarkStart>()))
                .Concat(para.Elements<Hyperlink>().SelectMany(hl => hl.Elements<BookmarkStart>()))
                .ToList();
            // BUG-DUMP-BMSPAN: a bookmark that WRAPS content (runs/equations
            // between BookmarkStart and the matching BookmarkEnd) must round-
            // trip with the End placed AFTER the wrapped content, not adjacent
            // to the Start. Surface the matching BookmarkEnd as its own DOM-
            // ordered "bookmarkEnd" node so the emitter can replay a separate,
            // positioned `add bookmark ... end=true` op (mirrors the existing
            // two-marker model; the End may even live in a downstream
            // paragraph — cross-paragraph spans surface their End there).
            // Empty/zero-length bookmarks (End immediately follows Start) keep
            // the single combined `add bookmark` op so they stay empty.
            var paraBookmarkEnds = para.Elements<BookmarkEnd>()
                .Concat(para.Elements<InsertedRun>().SelectMany(ins => ins.Elements<BookmarkEnd>()))
                .Concat(para.Elements<DeletedRun>().SelectMany(del => del.Elements<BookmarkEnd>()))
                .Concat(para.Elements<Hyperlink>().SelectMany(hl => hl.Elements<BookmarkEnd>()))
                .Where(be => be.Id?.Value != null && IsContentSpanBookmark(be))
                .ToList();
            // BUG-DUMP-PERM: ranged editing-permission markers (<w:permStart>/
            // <w:permEnd>) are positioned paragraph children just like bookmark
            // markers — they delimit a region a group/user may edit inside a
            // protected document. Surface them in the DOM-ordered merge so the
            // emitter replays a positioned `add permStart`/`add permEnd` op at
            // each marker's original offset (mirrors the bookmark two-marker
            // path). Without this they were dropped entirely on round-trip.
            var paraPermStarts = para.Elements<PermStart>().ToList();
            var paraPermEnds = para.Elements<PermEnd>().ToList();
            // BUG-DUMP-FLDSIMPLE-ORDER: direct-child <w:fldSimple> (e.g. a footer
            // STYLEREF/PAGE) must surface at its DOCUMENT position in the merged
            // child list, not hoisted to the tail. The legacy standalone loop
            // below appended every fldSimple after the positional merge, so a
            // paragraph shaped "text <tab> <fldSimple> text <pageField>" round-
            // tripped as "text <pageField> <tab> text <fldSimple>" — visible
            // content (a running-header STYLEREF beside a page number) reordered.
            // Surface direct-child fldSimple as a positioned "fldSimple" kind;
            // the hyperlink-nested fldSimple loop further down keeps its own
            // hyperlink-scoped path and is unaffected.
            var paraSimpleFields = para.Elements<SimpleField>().ToList();
            // BUG-DUMP-RUBY: ruby-bearing runs (a <w:r> wrapping <w:ruby>, the
            // CJK phonetic guide) are excluded from GetAllRuns (their inner
            // <w:rt>/<w:rubyBase> runs would otherwise flatten into sequential
            // plain runs, dropping the wrapper). Surface the outer run as a
            // "ruby" kind in the DOM-ordered merge so it emits at its original
            // intra-paragraph position; the emitter re-inserts the verbatim
            // <w:r><w:ruby>…</w:r> via raw-set. Revision-wrapped ruby
            // (<w:ins>/<w:del> ancestor) is left to the legacy DUMP7-10 path.
            var paraRubyRuns = para.Descendants<Run>()
                .Where(r => r.GetFirstChild<Ruby>() != null
                    && r.Ancestors<InsertedRun>().FirstOrDefault() == null
                    && r.Ancestors<DeletedRun>().FirstOrDefault() == null)
                .ToList();
            // BUG-DUMP-R42-9: <w:bdo> (bidirectional override) is a run-container
            // — surface each as a "bdo" kind in the DOM-ordered merge so the
            // emitter re-inserts the verbatim <w:bdo>…</w:bdo> via raw-set,
            // preserving the w:val direction and the wrapped runs. GetAllRuns
            // already drops the inner runs (so they don't double-emit as plain
            // runs). Mirrors the ruby raw-set path.
            var paraBdos = para.Elements<BidirectionalOverride>().ToList();
            // BUG-DUMP-R43-7: <w:dir> (BidirectionalEmbedding) is the third bidi
            // run-container — mirror the bdo merge so the verbatim wrapper replays.
            var paraDirs = para.Elements<BidirectionalEmbedding>().ToList();
            var ordered = runs.Where(r => r.GetFirstChild<Ruby>() == null)
                .Select(r => (pos: descendantPos.TryGetValue(r, out var p) ? p : int.MaxValue, kind: "run", el: (OpenXmlElement)r))
                .Concat(paraRubyRuns.Select(r => (pos: descendantPos.TryGetValue(r, out var p) ? p : int.MaxValue, kind: "ruby", el: (OpenXmlElement)r)))
                .Concat(paraBdos.Select(b => (pos: descendantPos.TryGetValue(b, out var p) ? p : int.MaxValue, kind: "bdo", el: (OpenXmlElement)b)))
                .Concat(paraDirs.Select(b => (pos: descendantPos.TryGetValue(b, out var p) ? p : int.MaxValue, kind: "dir", el: (OpenXmlElement)b)))
                .Concat(inlineEqsAll.Select(e => (pos: descendantPos.TryGetValue(e, out var p) ? p : int.MaxValue, kind: "eq", el: (OpenXmlElement)e)))
                .Concat(bareFieldUnknowns.Select(u => (pos: descendantPos.TryGetValue(u, out var p) ? p : int.MaxValue, kind: u.LocalName == "fldChar" ? "fieldChar" : "instrText", el: (OpenXmlElement)u)))
                .Concat(paraSimpleFields.Select(f => (pos: descendantPos.TryGetValue(f, out var p) ? p : int.MaxValue, kind: "fldSimple", el: (OpenXmlElement)f)))
                .Concat(paraBookmarks.Select(b => (pos: descendantPos.TryGetValue(b, out var p) ? p : int.MaxValue, kind: "bookmark", el: (OpenXmlElement)b)))
                .Concat(paraBookmarkEnds.Select(b => (pos: descendantPos.TryGetValue(b, out var p) ? p : int.MaxValue, kind: "bookmarkEnd", el: (OpenXmlElement)b)))
                .Concat(paraPermStarts.Select(b => (pos: descendantPos.TryGetValue(b, out var p) ? p : int.MaxValue, kind: "permStart", el: (OpenXmlElement)b)))
                .Concat(paraPermEnds.Select(b => (pos: descendantPos.TryGetValue(b, out var p) ? p : int.MaxValue, kind: "permEnd", el: (OpenXmlElement)b)))
                .OrderBy(t => t.pos)
                .ToList();
            int bareFieldIdx = 0;
            int fldSimpleMergeIdx = 0;
            int fldImgRunIdx = 0;   // BUG-DUMP-FLDSIMPLE-IMG: paragraph-scoped, matches the fldimgrun[] selector
            foreach (var entry in ordered)
            {
                int _childCountBefore = node.Children.Count;
                if (entry.kind == "fldSimple")
                {
                    // BUG-DUMP-FLDSIMPLE-ORDER: emit a direct-child <w:fldSimple>
                    // as a positioned `field` node (same shape as the legacy
                    // standalone loop, which is now removed for direct children).
                    var fld = (SimpleField)entry.el;
                    var instr = fld.Instruction?.Value ?? "";
                    // BUG-DUMP-FLDSIMPLE-IMG: a <w:fldSimple> whose result holds a
                    // <w:drawing> (e.g. REF SHAPE caching the referenced inline image)
                    // cannot round-trip through the text-only `field` node below — the
                    // image was silently dropped. Decompose into a complex field that
                    // keeps the picture in its result: synthesized begin/instr/separate
                    // markers (inline raw) + a real picture node (ships the image bytes
                    // and rebinds the blip rel via the resolvable fldimgrun[] path) + an
                    // end marker. Mirrors BUG-DUMP-R28-INCLUDEPICTURE for fldChar chains;
                    // fldSimple already round-trips as a complex field, so it is faithful.
                    if (fld.Descendants<Drawing>().Any())
                    {
                        string EscXml(string s) => s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;");
                        DocumentNode MarkerNode(string xml) => new DocumentNode
                        {
                            Type = "field",
                            Path = $"{path}/field[{fldSimpleMergeIdx + 1}]",
                            Format = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
                            {
                                ["_fieldMarkerRaw"] = true,
                                ["_markerInlineXml"] = xml,
                            }
                        };
                        node.Children.Add(MarkerNode(
                            "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>"
                            + "<w:r><w:instrText xml:space=\"preserve\">" + EscXml(instr) + "</w:instrText></w:r>"
                            + "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>"));
                        var pendingRaw = new System.Text.StringBuilder();
                        void FlushRaw()
                        {
                            if (pendingRaw.Length == 0) return;
                            node.Children.Add(MarkerNode(pendingRaw.ToString()));
                            pendingRaw.Clear();
                        }
                        foreach (var rr in fld.Elements<Run>())
                        {
                            var rd = rr.GetFirstChild<Drawing>();
                            if (rd != null)
                            {
                                FlushRaw();
                                node.Children.Add(CreateImageNode(rd, rr, $"{path}/fldimgrun[{++fldImgRunIdx}]"));
                            }
                            else pendingRaw.Append(rr.OuterXml);
                        }
                        FlushRaw();
                        node.Children.Add(MarkerNode("<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>"));
                        fldSimpleMergeIdx++;
                    }
                    else if (fld.Elements<DeletedRun>().Any() || fld.Elements<InsertedRun>().Any()
                             || fld.Elements<MoveFromRun>().Any() || fld.Elements<MoveToRun>().Any())
                    {
                        // BUG-DUMP-H79: a <w:fldSimple> whose cached result contains a
                        // tracked change (<w:del>/<w:ins>/<w:moveFrom>/<w:moveTo>) cannot
                        // round-trip through the text-only `field` node below — the
                        // displayText scan uses Descendants<Text>(), which excludes
                        // <w:delText>, so the deleted result text was silently dropped,
                        // and the typed `add field` path has no model for a deleted result
                        // run. Decompose into a complex field whose result content is
                        // emitted VERBATIM (each fldSimple child OuterXml, preserving the
                        // <w:del>/<w:ins> wrapper), bracketed by synthesized begin/instr/
                        // separate/end markers. Mirrors the drawing-result decomposition
                        // above and the complex-field-result del fix.
                        string EscInstr(string s) => s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;");
                        var chain = new System.Text.StringBuilder();
                        chain.Append("<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>")
                             .Append("<w:r><w:instrText xml:space=\"preserve\">").Append(EscInstr(instr)).Append("</w:instrText></w:r>")
                             .Append("<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>");
                        foreach (var child in fld.Elements())
                            chain.Append(child.OuterXml);
                        chain.Append("<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>");
                        node.Children.Add(new DocumentNode
                        {
                            Type = "field",
                            Path = $"{path}/field[{fldSimpleMergeIdx + 1}]",
                            Format = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
                            {
                                ["_fieldMarkerRaw"] = true,
                                ["_markerInlineXml"] = chain.ToString(),
                            }
                        });
                        fldSimpleMergeIdx++;
                    }
                    else
                    {
                        var displayText = string.Join("", fld.Descendants<Text>().Select(t => t.Text));
                        var fldNode = new DocumentNode
                        {
                            Type = "field",
                            Text = displayText,
                            Path = $"{path}/field[{fldSimpleMergeIdx + 1}]",
                        };
                        fldNode.Format["instruction"] = instr.Trim();
                        var instrUpper = instr.Trim().Split(' ', 2)[0].ToUpperInvariant();
                        if (!string.IsNullOrEmpty(instrUpper))
                            fldNode.Format["fieldType"] = instrUpper.ToLowerInvariant();
                        if (fld.Dirty?.Value == true) fldNode.Format["dirty"] = true;
                        if (fld.FieldLock?.Value == true) fldNode.Format["fldLock"] = true;
                        fldNode.Format["evaluated"] = displayText.Length > 0;
                        node.Children.Add(fldNode);
                        fldSimpleMergeIdx++;
                    }
                }
                else if (entry.kind == "run")
                {
                    // BUG-DUMP-H99: a run that packs an entire field's fldChar
                    // chain into one <w:r> must be split into the same one-
                    // structural-element-per-run node stream Word emits, so
                    // CollapseFieldChains rebuilds the field instead of seeing a
                    // lone begin and warn-dropping it. Uniform run formatting (the
                    // packed run carries a single rPr) round-trips fully via the
                    // typed field path. See IsPackedFieldRun.
                    if (entry.el is Run packedRun && IsPackedFieldRun(packedRun))
                    {
                        var packedRpr = packedRun.GetFirstChild<RunProperties>();
                        string? hlParentForPacked = null;
                        if (packedRun.Parent is Hyperlink packedHl)
                        {
                            int hlIdxP = paraHyperlinks.IndexOf(packedHl);
                            if (hlIdxP >= 0) hlParentForPacked = $"{path}/hyperlink[{hlIdxP + 1}]";
                        }
                        bool packedInCustomXml = packedRun.Ancestors<CustomXmlRun>().FirstOrDefault() != null;
                        foreach (var packedChild in packedRun.ChildElements)
                        {
                            if (packedChild is RunProperties) continue;
                            if (packedChild is LastRenderedPageBreak) continue;
                            var synthRun = new Run();
                            if (packedRpr != null)
                                synthRun.AppendChild((RunProperties)packedRpr.CloneNode(true));
                            synthRun.AppendChild(packedChild.CloneNode(true));
                            var packedNode = ElementToNode(synthRun, $"{path}/r[{runIdx + 1}]", depth - 1);
                            if (hlParentForPacked != null)
                                packedNode.Format["_hyperlinkParent"] = hlParentForPacked;
                            if (packedInCustomXml)
                                packedNode.Format["_wrapperFlattened"] = true;
                            node.Children.Add(packedNode);
                            runIdx++;
                        }
                    }
                    else
                    {
                    var runNode = ElementToNode(entry.el, $"{path}/r[{runIdx + 1}]", depth - 1);
                    // BUG-DUMP-R35-2: unlike <w:smartTag> (OpenXmlUnknownElement in
                    // this SDK build — handled by the unknown-subtree synthesizer
                    // below), a run-level <w:customXml> parses as a TYPED
                    // CustomXmlRun, so its inner runs arrive here via GetAllRuns
                    // and the wrapper is flattened by the typed path. Mark them so
                    // the emitter surfaces the same deterministic flatten warning.
                    // CONSISTENCY(wrapper-flatten-warning).
                    if (entry.el.Ancestors<CustomXmlRun>().FirstOrDefault() != null)
                        runNode.Format["_wrapperFlattened"] = true;
                    // BUG-DUMP18-02: surface a hyperlink-scoped subpath on
                    // runs that are direct children of <w:hyperlink>. The
                    // canonical Path stays flat (/…/r[N]) for back-compat
                    // with every existing caller; WordBatchEmitter's
                    // CollapseFieldChains carries this hint to the synth
                    // field-add row so a fldChar-chain field inside a
                    // hyperlink replays INSIDE the hyperlink instead of
                    // alongside it. Mirrors the SimpleField hyperlink-
                    // scope path emitted below.
                    if (entry.el.Parent is Hyperlink runHl)
                    {
                        int hlIdxRun = paraHyperlinks.IndexOf(runHl);
                        if (hlIdxRun >= 0)
                            runNode.Format["_hyperlinkParent"] = $"{path}/hyperlink[{hlIdxRun + 1}]";
                    }
                    node.Children.Add(runNode);
                    runIdx++;
                    }
                }
                else if (entry.kind == "eq")
                {
                    // BUG-DUMP15-04: equations whose immediate parent is
                    // <w:hyperlink> get a hyperlink-scoped path so the
                    // emitter can place the equation INSIDE the hyperlink
                    // on replay.
                    // R3-bt-1: emit the RESOLVABLE child segment `oMath[N]`, not
                    // `equation[N]`. An inline equation is a bare <m:oMath>
                    // (LocalName "oMath") inside the w:p/w:hyperlink, so the path
                    // resolver matches it via LocalName == "oMath" — `equation`
                    // matched no element and get/set/remove on the listed
                    // `…/equation[1]` failed ("No equation found … Available:
                    // oMath(1)"). This now agrees with what `query equation`
                    // already returns (…/oMath[N]).
                    string eqPath;
                    if (entry.el.Parent is Hyperlink eqHl)
                    {
                        int hlIdx = paraHyperlinks.IndexOf(eqHl);
                        int hlEqIdx = eqHl.Elements<M.OfficeMath>()
                            .ToList().IndexOf((M.OfficeMath)entry.el);
                        eqPath = $"{path}/hyperlink[{hlIdx + 1}]/oMath[{hlEqIdx + 1}]";
                    }
                    else
                    {
                        eqPath = $"{path}/oMath[{inlineEqIdx + 1}]";
                        inlineEqIdx++;
                    }
                    var eqNode = ElementToNode(entry.el, eqPath, depth - 1);
                    // BUG-DUMP-EQVERBATIM: stash the verbatim <m:oMath> so the
                    // emitter can round-trip it exactly (the LaTeX formula string
                    // drops per-run <w:rPr> like rFonts="Cambria Math"). Captured
                    // here from the live element — the equation[N] path doesn't
                    // resolve through GetElementXml.
                    eqNode.Format["_omathXml"] = entry.el.OuterXml;
                    node.Children.Add(eqNode);
                }
                else if (entry.kind == "ruby")
                {
                    // BUG-DUMP-RUBY: emit the ruby-bearing run at its DOM
                    // position. The node carries the verbatim outer-run XML
                    // (stashed under _rawRubyXml) so the emitter re-inserts the
                    // <w:r><w:ruby>…</w:r> via a raw-set append, preserving
                    // <w:rt> (furigana) + <w:rubyBase> (base text) + <w:rubyPr>
                    // (alignment/hps/lid). Base text surfaces in Text so view/
                    // readback isn't empty.
                    var rubyRun = (Run)entry.el;
                    var rubyEl = rubyRun.GetFirstChild<Ruby>()!;
                    var rubyBase = rubyEl.ChildElements.FirstOrDefault(c => c.LocalName == "rubyBase");
                    var baseText = rubyBase == null ? "" :
                        string.Concat(rubyBase.Descendants<Text>().Select(t => t.Text));
                    var rubyNode = new DocumentNode
                    {
                        Type = "ruby",
                        Text = baseText,
                        Path = $"{path}/r[{runIdx + 1}]",
                    };
                    rubyNode.Format["_rawRubyXml"] = rubyRun.OuterXml;
                    node.Children.Add(rubyNode);
                    runIdx++;
                }
                else if (entry.kind == "bdo")
                {
                    // BUG-DUMP-R42-9: emit the <w:bdo> bidirectional-override
                    // wrapper at its DOM position. The node carries the verbatim
                    // outer XML (stashed under _rawBdoXml) so the emitter re-inserts
                    // the <w:bdo>…</w:bdo> via a raw-set append, preserving the
                    // w:val direction and the wrapped runs' visual char ordering.
                    // Inner-run text surfaces in Text so view/readback isn't empty.
                    var bdoEl = (BidirectionalOverride)entry.el;
                    var bdoText = string.Concat(bdoEl.Descendants<Text>().Select(t => t.Text));
                    var bdoNode = new DocumentNode
                    {
                        Type = "bdo",
                        Text = bdoText,
                        Path = $"{path}/r[{runIdx + 1}]",
                    };
                    bdoNode.Format["_rawBdoXml"] = bdoEl.OuterXml;
                    if (bdoEl.Val?.HasValue == true)
                        bdoNode.Format["bdo.val"] = bdoEl.Val.InnerText;
                    node.Children.Add(bdoNode);
                    runIdx++;
                }
                else if (entry.kind == "dir")
                {
                    // BUG-DUMP-R43-7: emit the <w:dir> bidirectional-embedding
                    // wrapper at its DOM position. The node carries the verbatim
                    // outer XML (stashed under _rawDirXml) so the emitter re-inserts
                    // the <w:dir>…</w:dir> via a raw-set append, preserving the
                    // w:val direction and the wrapped runs. Mirrors the bdo path.
                    var dirEl = (BidirectionalEmbedding)entry.el;
                    var dirText = string.Concat(dirEl.Descendants<Text>().Select(t => t.Text));
                    var dirNode = new DocumentNode
                    {
                        Type = "dir",
                        Text = dirText,
                        Path = $"{path}/r[{runIdx + 1}]",
                    };
                    dirNode.Format["_rawDirXml"] = dirEl.OuterXml;
                    if (dirEl.Val?.HasValue == true)
                        dirNode.Format["dir.val"] = dirEl.Val.InnerText;
                    node.Children.Add(dirNode);
                    runIdx++;
                }
                else if (entry.kind == "bookmark")
                {
                    // BUG-DUMP25-01: emit BookmarkStart at its DOM position
                    // (sandwiched between sibling runs/equations) so dump→
                    // batch round-trips preserve mid-paragraph bookmark
                    // offsets like Word's _GoBack resume-cursor mark.
                    // Path index counts bookmarks among themselves to
                    // stay 1-based, mirroring the legacy bmIdx counter.
                    int bmPathIdx = paraBookmarks.IndexOf((BookmarkStart)entry.el);
                    // _spanOpen is set inside BookmarkStartToNode (ElementToNode).
                    node.Children.Add(ElementToNode(entry.el, $"{path}/bookmark[{bmPathIdx + 1}]", depth - 1));
                }
                else if (entry.kind == "bookmarkEnd")
                {
                    // BUG-DUMP-BMSPAN: standalone BookmarkEnd marker for a
                    // content-wrapping bookmark. Carries only the matching
                    // bookmark name so the emitter can replay a positioned
                    // `add bookmark name=X end=true` after the wrapped runs.
                    var be = (BookmarkEnd)entry.el;
                    var beNode = new DocumentNode
                    {
                        Type = "bookmarkEnd",
                        // BUG-DUMP-BMEND-IDPATH: address by w:id explicitly. A bare
                        // numeric bracket is resolved positionally, but a bookmarkEnd's
                        // id need not equal its document-order position (stacked TOC
                        // anchors where one closes inside a field result), which
                        // re-read the wrong end and duplicated a bookmark id.
                        Path = $"{path}/bookmarkEnd[@id={be.Id?.Value}]",
                    };
                    var matchName = ResolveBookmarkEndName(be);
                    if (!string.IsNullOrEmpty(matchName))
                        beNode.Format["name"] = matchName!;
                    node.Children.Add(beNode);
                }
                else if (entry.kind == "permStart")
                {
                    // BUG-DUMP-PERM: surface the permStart marker at its DOM
                    // position with all its attributes so the emitter can replay
                    // a positioned `add permStart`.
                    node.Children.Add(PermStartToNode((PermStart)entry.el,
                        new DocumentNode { Path = $"{path}/permStart[{paraPermStarts.IndexOf((PermStart)entry.el) + 1}]" }));
                }
                else if (entry.kind == "permEnd")
                {
                    var pe = (PermEnd)entry.el;
                    var peNode = new DocumentNode
                    {
                        Type = "permEnd",
                        Path = $"{path}/permEnd[{PathIndex.FromArrayIndex(paraPermEnds.IndexOf(pe))}]",
                    };
                    if (pe.Id?.Value != null) peNode.Format["id"] = pe.Id.Value.ToString();
                    node.Children.Add(peNode);
                }
                else
                {
                    // BUG-DUMP15-02: synthesize fieldChar/instrText nodes
                    // for bare unknown elements so CollapseFieldChains can
                    // stitch the field. Mirrors the Run-based shape.
                    var u = (DocumentFormat.OpenXml.OpenXmlUnknownElement)entry.el;
                    var bn = new DocumentNode
                    {
                        Type = entry.kind,
                        Path = $"{path}/r[{runIdx + 1}]",
                    };
                    runIdx++;
                    if (entry.kind == "fieldChar")
                    {
                        var fct = u.GetAttribute("fldCharType", wNs2).Value;
                        if (!string.IsNullOrEmpty(fct))
                            bn.Format["fieldCharType"] = fct;
                    }
                    else // instrText
                    {
                        bn.Format["instruction"] = u.InnerText;
                        bn.Text = u.InnerText;
                    }
                    node.Children.Add(bn);
                    bareFieldIdx++;
                }
                // BUG-DUMP-R29-SMARTTAG: record this entry's doc-position for
                // every child it appended, keeping childPositions parallel to
                // node.Children (ascending, since `ordered` is pos-sorted).
                for (int _z = node.Children.Count - _childCountBefore; _z > 0; _z--)
                    childPositions.Add(entry.pos);
            }
            // BUG-DUMP5-06/07: <w:ruby> and <w:smartTag> aren't registered
            // as typed paragraph children in the OpenXml SDK schema set we
            // load — RawSet-injected fragments and SDK-untracked content
            // from real-world docx files surface them as
            // OpenXmlUnknownElement, so Descendants<Run>() inside
            // GetAllRuns skips every nested run (the inner <w:r> is also
            // an unknown element, not a typed Run). Walk the unknown
            // subtrees and synthesize plain `run` DocumentNodes from any
            // <w:r>/<w:t> children we find so the inner text round-trips
            // through dump→batch instead of vanishing.
            const string wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            // BUG-DUMP-R35-2: smartTag/customXml wrappers parse as
            // OpenXmlUnknownElement; the SDK's unknown-element reader DISCARDS
            // insignificant whitespace, so a `<w:r><w:t> </w:t></w:r>` (no
            // xml:space="preserve") between two nested wrappers collapses to an
            // empty `<w:t/>` in the parsed DOM (verified: para.OuterXml shows the
            // self-closed tag). The space still exists in the part stream, so —
            // when the parsed run text is empty — recover the run's true,
            // whitespace-preserved text from the raw part XML by run ordinal.
            // A typed Run keeps its <w:t> text node (Text.Text returns " " even
            // without preserve), so this loss is specific to the unknown path.
            var rawWrapperTexts = RecoverWrapperRunTexts(para);
            int wrapperRunOrdinal = -1;
            foreach (var unkRun in para.Descendants<DocumentFormat.OpenXml.OpenXmlUnknownElement>())
            {
                if (unkRun.LocalName != "r" || unkRun.NamespaceUri != wNs) continue;
                // Is this unknown run nested under a smartTag/customXml wrapper
                // (vs a ruby/ins/del subtree)? Only the wrapper case advances the
                // raw-text recovery cursor and triggers the flatten warning.
                bool inWrapper = false;
                for (var anc = unkRun.Parent; anc != null; anc = anc.Parent)
                {
                    if (anc is DocumentFormat.OpenXml.OpenXmlUnknownElement uw
                        && uw.NamespaceUri == wNs
                        && (uw.LocalName == "smartTag" || uw.LocalName == "customXml"))
                    { inWrapper = true; break; }
                    if (anc == para) break;
                }
                if (inWrapper) wrapperRunOrdinal++;
                // BUG-DUMP-ALTCONTENT-DOUBLE: a run inside an <mc:AlternateContent>
                // (a WPS/DrawingML shape with a VML <mc:Fallback>) or inside a
                // <w:txbxContent> is part of a textbox/drawing that is round-tripped
                // VERBATIM via a raw-set (the textbox/drawing emit). The SDK parses
                // the AlternateContent / VML Fallback subtree as OpenXmlUnknownElement,
                // so its inner <w:r> reaches THIS synthesizer — and because both the
                // mc:Choice and mc:Fallback branches hold the SAME text, synthesizing
                // them as plain runs duplicated a shape's text ("AustraliaIndonesia")
                // up to 4x in the body. Skip any run whose ancestor chain crosses an
                // AlternateContent/Choice/Fallback wrapper or a txbxContent.
                bool inDrawingWrapper = false;
                for (var anc = unkRun.Parent; anc != null && anc != para; anc = anc.Parent)
                {
                    var ln = anc.LocalName;
                    if (ln == "AlternateContent" || ln == "Choice" || ln == "Fallback"
                        || ln == "txbxContent" || ln == "txbx" || ln == "pict"
                        || ln == "textbox")
                    { inDrawingWrapper = true; break; }
                }
                if (inDrawingWrapper) continue;
                // Only surface runs whose direct parent is an unknown
                // wrapper (ruby/rt/rubyBase/smartTag/customXml). Runs
                // whose parent is a typed Paragraph would already be
                // typed Runs and reached via GetAllRuns above; if they
                // somehow surface as unknown here it's because the
                // entire paragraph is malformed and we'd duplicate.
                // BUG-DUMP7-10: also accept InsertedRun/DeletedRun
                // ancestors — w:del>w:ruby in a malformed doc parses
                // ruby as unknown but the typed w:del wrapper still
                // sits between para and the unknown subtree, so the
                // ancestor (not just direct parent) needs the typed
                // change-track wrapper allowance.
                if (unkRun.Parent is not DocumentFormat.OpenXml.OpenXmlUnknownElement
                    && unkRun.Ancestors<InsertedRun>().FirstOrDefault() == null
                    && unkRun.Ancestors<DeletedRun>().FirstOrDefault() == null)
                    continue;
                var sbInner = new System.Text.StringBuilder();
                foreach (var tEl in unkRun.Descendants<DocumentFormat.OpenXml.OpenXmlUnknownElement>())
                {
                    if (tEl.NamespaceUri != wNs) continue;
                    // BUG-DUMP7-10: a w:del-wrapped ruby's inner runs
                    // carry their text in <w:delText>, not <w:t>.
                    // Without delText/instrText the "base"/"rt" text
                    // dropped silently and the paragraph surfaced empty.
                    if (tEl.LocalName == "t"
                        || tEl.LocalName == "delText"
                        || tEl.LocalName == "instrText")
                        sbInner.Append(tEl.InnerText);
                }
                var recoveredText = sbInner.ToString();
                // BUG-DUMP-R35-2: when the parsed run text is empty but a
                // whitespace-preserved text was recovered from the raw part XML
                // at this wrapper-run ordinal, use it so a space-only run between
                // two nested smartTags survives dump→batch ("John Smith", not
                // "JohnSmith"). Only consult the raw cursor for wrapper runs.
                if (recoveredText.Length == 0 && inWrapper
                    && wrapperRunOrdinal >= 0 && wrapperRunOrdinal < rawWrapperTexts.Count)
                    recoveredText = rawWrapperTexts[wrapperRunOrdinal];
                // BUG-DUMP-SMARTTAG-BR: a wrapper run whose only content is a
                // <w:br/> / <w:cr/> (a line break between two nested smartTags —
                // e.g. the <br/> separating "123 Main St." from "Olympia, WA" in
                // a multi-line address) carries no text, so the recovery above
                // left it empty. Dropping it (the bare continue below) joined the
                // two lines and compressed the block, drifting the page.
                // Synthesize a typed break node so the inline line break survives
                // the wrapper flatten; insert it at the wrapper run's true
                // document position like the text-run synth path does.
                if (recoveredText.Length == 0)
                {
                    var brkSep = unkRun.ChildElements.FirstOrDefault(c =>
                        c.NamespaceUri == wNs && (c.LocalName == "br" || c.LocalName == "cr"));
                    if (brkSep != null)
                    {
                        var brkNode = new DocumentNode { Type = "break", Path = $"{path}/r[{runIdx + 1}]" };
                        var brkT = brkSep.GetAttributes()
                            .FirstOrDefault(a => a.LocalName == "type" && a.NamespaceUri == wNs).Value;
                        brkNode.Format["breakType"] = string.IsNullOrEmpty(brkT) ? "line" : brkT;
                        var brkPos = descendantPos.TryGetValue(unkRun, out var bp) ? bp : int.MaxValue;
                        int brkIdx = childPositions.FindIndex(cp => cp > brkPos);
                        if (brkIdx < 0) brkIdx = node.Children.Count;
                        node.Children.Insert(brkIdx, brkNode);
                        childPositions.Insert(brkIdx, brkPos);
                        runIdx++;
                    }
                    continue;
                }
                var synthNode = new DocumentNode
                {
                    Type = "run",
                    Text = recoveredText,
                    Path = $"{path}/r[{runIdx + 1}]",
                };
                // BUG-DUMP-R29-SMARTTAG-RPR: carry the wrapped run's own rPr into
                // the synthesized node. The unknown-subtree path previously emitted
                // the run as bare text, dropping its bold/size/font/etc. A bold,
                // 10pt "AND" inside a heading ("INSPECTION <smartTag>AND</smartTag>
                // ACCEPTANCE") came back non-bold AND at the docDefaults 11pt — the
                // taller mis-sized word grew the heading line, and across many such
                // headings the extra height cascaded into whole-document pagination
                // drift. The rPr is itself an OpenXmlUnknownElement, so map its
                // children to the same canonical run keys RunToNode emits.
                var unkRPr = unkRun.ChildElements.FirstOrDefault(c =>
                    c.LocalName == "rPr" && c.NamespaceUri == wNs);
                if (unkRPr != null)
                {
                    string? Attr(OpenXmlElement el, string n) => el.GetAttributes()
                        .FirstOrDefault(a => a.LocalName == n && a.NamespaceUri == wNs).Value;
                    bool ToggleOn(OpenXmlElement el)
                    { var v = Attr(el, "val"); return v is null or "1" or "true" or "on"; }
                    foreach (var ch in unkRPr.ChildElements)
                    {
                        if (ch.NamespaceUri != wNs) continue;
                        switch (ch.LocalName)
                        {
                            case "b": synthNode.Format["bold"] = ToggleOn(ch); break;
                            case "bCs": synthNode.Format["bold.cs"] = ToggleOn(ch); break;
                            case "i": synthNode.Format["italic"] = ToggleOn(ch); break;
                            case "iCs": synthNode.Format["italic.cs"] = ToggleOn(ch); break;
                            case "caps": synthNode.Format["caps"] = ToggleOn(ch); break;
                            case "smallCaps": synthNode.Format["smallcaps"] = ToggleOn(ch); break;
                            case "strike": synthNode.Format["strike"] = ToggleOn(ch); break;
                            case "sz":
                                if (Attr(ch, "val") is { } sv && int.TryParse(sv, out var szi))
                                    synthNode.Format["size"] = $"{szi / 2.0:0.##}pt";
                                break;
                            case "szCs":
                                if (Attr(ch, "val") is { } scv && int.TryParse(scv, out var szci))
                                    synthNode.Format["size.cs"] = $"{szci / 2.0:0.##}pt";
                                break;
                            case "color": if (Attr(ch, "val") is { } cv) synthNode.Format["color"] = cv; break;
                            case "highlight": if (Attr(ch, "val") is { } hv) synthNode.Format["highlight"] = hv; break;
                            case "u": if (Attr(ch, "val") is { } uv) synthNode.Format["underline"] = uv; break;
                            case "rStyle": if (Attr(ch, "val") is { } rsv) synthNode.Format["rStyle"] = rsv; break;
                            case "vertAlign":
                                var va = Attr(ch, "val");
                                if (va == "superscript") synthNode.Format["superscript"] = true;
                                else if (va == "subscript") synthNode.Format["subscript"] = true;
                                break;
                            case "rFonts":
                                if (Attr(ch, "ascii") is { } fa) synthNode.Format["font.latin"] = fa;
                                else if (Attr(ch, "hAnsi") is { } fh) synthNode.Format["font.latin"] = fh;
                                if (Attr(ch, "eastAsia") is { } fe) synthNode.Format["font.ea"] = fe;
                                if (Attr(ch, "cs") is { } fc) synthNode.Format["font.cs"] = fc;
                                break;
                        }
                    }
                }
                // BUG-DUMP-R35-2: mark a wrapper-flattened run so the emitter can
                // surface a deterministic "wrapper flattened" warning (the inner
                // run text/formatting is preserved; only the smartTag/customXml
                // wrapper element is dropped — consistent with Word often
                // stripping these and the project's flatten precedents).
                if (inWrapper)
                    synthNode.Format["_wrapperFlattened"] = true;
                // BUG-DUMP7-10: preserve trackChange attribution from
                // the typed w:ins/w:del ancestor so the round-trip
                // re-emits the wrapper (mirrors the typed-Run branch
                // at the top of this method).
                var insAnc = unkRun.Ancestors<InsertedRun>().FirstOrDefault();
                if (insAnc != null)
                {
                    synthNode.Format["revision.type"] = "ins";
                    if (!string.IsNullOrEmpty(insAnc.Author?.Value))
                        synthNode.Format["revision.author"] = insAnc.Author!.Value!;
                    if (insAnc.Date?.Value is DateTime insAncDate)
                        synthNode.Format["revision.date"] = insAncDate.ToString("o");
                }
                else
                {
                    var delAnc = unkRun.Ancestors<DeletedRun>().FirstOrDefault();
                    if (delAnc != null)
                    {
                        synthNode.Format["revision.type"] = "del";
                        if (!string.IsNullOrEmpty(delAnc.Author?.Value))
                            synthNode.Format["revision.author"] = delAnc.Author!.Value!;
                        if (delAnc.Date?.Value is DateTime delAncDate)
                            synthNode.Format["revision.date"] = delAncDate.ToString("o");
                    }
                    else
                    {
                        // BUG-DUMP-SMARTTAG-DELWRAP: when the tracked-change wrapper
                        // sits INSIDE a <w:smartTag>/<w:customXml> (itself an
                        // OpenXmlUnknownElement), the <w:ins>/<w:del>/<w:moveFrom>/
                        // <w:moveTo> between the wrapper and this run also parses as
                        // an OpenXmlUnknownElement — the typed Ancestors<> probes
                        // above both miss it, so a deletion nested in a smartTag lost
                        // its revision entirely and round-tripped as live <w:t> text
                        // (delText silently un-deleted). Walk the unknown-element
                        // ancestors for the w:ns revision wrapper and read its
                        // w:author/w:date attributes by name.
                        var revAnc = unkRun.Ancestors<DocumentFormat.OpenXml.OpenXmlUnknownElement>()
                            .FirstOrDefault(a => a.NamespaceUri == wNs
                                && a.LocalName is "ins" or "del" or "moveFrom" or "moveTo");
                        if (revAnc != null)
                        {
                            string? RevAttr(string n) => revAnc.GetAttributes()
                                .FirstOrDefault(a => a.LocalName == n && a.NamespaceUri == wNs).Value;
                            synthNode.Format["revision.type"] = revAnc.LocalName;
                            if (RevAttr("author") is { Length: > 0 } revAuthor)
                                synthNode.Format["revision.author"] = revAuthor;
                            if (RevAttr("date") is { Length: > 0 } revDate)
                                synthNode.Format["revision.date"] = revDate;
                            if (RevAttr("id") is { Length: > 0 } revId)
                                synthNode.Format["revision.id"] = revId;
                        }
                    }
                }
                // BUG-DUMP-R29-SMARTTAG: insert at the wrapper run's true document
                // position instead of appending at the tail, so a mid-paragraph
                // smartTag/customXml run ("PRICE <AND> TERMS") replays in order.
                // The /r[runIdx] path index keeps its tail numbering (typed runs
                // already claimed /r[1..M]; these synth runs are emitted as plain
                // `add r text=…` and never path-resolved, so the index is inert) —
                // only the node.Children ORDER drives the emit sequence.
                var wrapPos = descendantPos.TryGetValue(unkRun, out var wp) ? wp : int.MaxValue;
                int insertIdx = childPositions.FindIndex(cp => cp > wrapPos);
                if (insertIdx < 0) insertIdx = node.Children.Count;
                node.Children.Insert(insertIdx, synthNode);
                childPositions.Insert(insertIdx, wrapPos);
                runIdx++;
            }
            // BUG-DUMP25-01: BookmarkStart children are now surfaced
            // inside the DOM-ordered `ordered` merge above, so a
            // bookmark between two runs round-trips at its original
            // intra-paragraph offset. The legacy standalone loop here
            // (which appended every bookmark at the tail of
            // node.Children) is intentionally left empty.
            // BUG-DUMP4-06: surface inline SdtRun (content control) children
            // so WordBatchEmitter can re-emit a typed `add sdt` row carrying
            // alias/tag/type metadata. Without this, GetAllRuns unwrapped
            // the SdtRun's inner Run as a plain `add r` and the metadata
            // was silently dropped on dump round-trip.
            int sdtRunIdx = 0;
            foreach (var sdtR in para.Elements<SdtRun>())
            {
                node.Children.Add(ElementToNode(sdtR, $"{path}/sdt[{sdtRunIdx + 1}]", depth - 1));
                sdtRunIdx++;
            }
            // BUG-DUMP-BARE-BR: a <w:br/> / <w:cr/> that is a DIRECT child of
            // <w:p> (not wrapped in a <w:r>) is schema-invalid, so the SDK loads
            // it as an OpenXmlUnknownElement rather than a typed Break — and Word
            // still renders the line break. The run walk above only enumerates
            // <w:r> children, so these bare breaks were dropped, merging the lines
            // on round-trip. Surface each as a typed break node (mirroring the
            // smartTag-wrapped bare-break path) so the emitter replays it.
            foreach (var bareBr in para.ChildElements)
            {
                if (bareBr.NamespaceUri != wNs ||
                    (bareBr.LocalName != "br" && bareBr.LocalName != "cr"))
                    continue;
                var bareBrNode = new DocumentNode { Type = "break", Path = $"{path}/r[{node.Children.Count + 1}]" };
                var bbType = bareBr.GetAttributes()
                    .FirstOrDefault(a => a.LocalName == "type" && a.NamespaceUri == wNs).Value;
                bareBrNode.Format["breakType"] = string.IsNullOrEmpty(bbType) ? "line" : bbType;
                node.Children.Add(bareBrNode);
            }
            // BUG-DUMP7-03 / BUG-DUMP8-03 / BUG-DUMP9-04: inline <m:oMath>
            // children (including those nested inside w:ins/w:del/w:hyperlink
            // wrappers) are now interleaved with runs at the top of this
            // block (BUG-DUMP13-02) so DOM order is preserved. The
            // `inlineEqIdx` counter declared there carries forward into the
            // block-level oMathPara branch below.
            // BUG-DUMP12-02: surface block-level <m:oMathPara> children of a
            // mixed-content paragraph (paragraph that ALSO has ordinary
            // runs/hyperlinks/etc) as display equation nodes. The pure-wrapper
            // case is handled at the body level via the LocalName=="oMathPara"
            // branch in WalkBodyChild + IsOMathParaWrapperParagraph; the
            // mixed-content case falls through to plain p[N] and was silently
            // dropping the equation. We only emit when the para is NOT a pure
            // oMathPara wrapper, to avoid double-counting against the body
            // /oMathPara[M] addressing.
            if (!IsOMathParaWrapperParagraph(para))
            {
                // R4-bt-2: a display equation in a mixed-content paragraph is an
                // m:oMathPara (M.Paragraph) child. Emit the RESOLVABLE
                // oMathPara[N] segment — the resolver matches it by LocalName
                // ("oMathPara"), so the old equation[N] segment listed here did
                // not resolve via get/set/remove. Index among the paragraph's
                // own oMathPara children (a separate counter from the inline
                // oMath one) so the positional path matches what the resolver
                // enumerates.
                int mathParaIdx = 0;
                foreach (var blockEq in para.Elements<M.Paragraph>())
                {
                    mathParaIdx++;
                    node.Children.Add(ElementToNode(blockEq, $"{path}/oMathPara[{mathParaIdx}]", depth - 1));
                }
            }
            // BUG-DUMP6-01: surface <w:fldSimple> children as typed `field`
            // nodes so WordBatchEmitter can re-emit `add field` with the
            // instruction preserved. Without this, GetAllRuns descended into
            // SimpleField and surfaced the inner display run as a plain run,
            // silently dropping the w:instr attribute.
            // BUG-DUMP-FLDSIMPLE-ORDER: direct-child fldSimple is now emitted
            // INSIDE the positional `ordered` merge above (kind "fldSimple"),
            // so it lands at its document position instead of the child tail.
            // Only the hyperlink-NESTED fldSimple (a paragraph grandchild) is
            // handled here — BUG-DUMP9-03 / BUG-DUMP18-02: it must surface as
            // /…/p[N]/hyperlink[K]/field[M] so dump→batch replays the field
            // INSIDE the hyperlink rather than alongside it.
            for (int hlI = 0; hlI < paraHyperlinks.Count; hlI++)
            {
                var hl = paraHyperlinks[hlI];
                int perHlFldIdx = 0;
                foreach (var fld in hl.Elements<SimpleField>())
                {
                    var instr = fld.Instruction?.Value ?? "";
                    var displayText = string.Join("",
                        fld.Descendants<Text>().Select(t => t.Text));
                    var fldNode = new DocumentNode
                    {
                        Type = "field",
                        Text = displayText,
                        Path = $"{path}/hyperlink[{hlI + 1}]/field[{perHlFldIdx + 1}]",
                    };
                    fldNode.Format["instruction"] = instr.Trim();
                    var instrUpper = instr.Trim().Split(' ', 2)[0].ToUpperInvariant();
                    if (!string.IsNullOrEmpty(instrUpper))
                        fldNode.Format["fieldType"] = instrUpper.ToLowerInvariant();
                    var fldDirtyHl = fld.Dirty?.Value == true;
                    if (fldDirtyHl) fldNode.Format["dirty"] = true;
                    // BUG-DUMP-R37-4: see sibling fldSimple branch above.
                    if (fld.FieldLock?.Value == true) fldNode.Format["fldLock"] = true;
                    fldNode.Format["evaluated"] = displayText.Length > 0;
                    node.Children.Add(fldNode);
                    perHlFldIdx++;
                }
            }
        }
        return node;
    }

    // A wpg:wgp group (e.g. a `diagram`) read back as a clean "group" node:
    // the user-facing /body/group[N] path plus x/y/width/height, instead of the
    // raw wgp/grpSpPr/wsp internals the generic builder would emit. Mirrors the
    // pptx group readback so an agent can read the size, compute a target, and
    // `set /body/group[N] --prop width/height`. Node text stays at /body/textbox[K].
    private DocumentNode WgpGroupToNode(OpenXmlElement wgp, DocumentNode node, string path)
    {
        node.Type = "group";
        node.Path = System.Text.RegularExpressions.Regex.Replace(path, @"/wgp\[(\d+)\]$", "/group[$1]");

        OpenXmlElement? Child(OpenXmlElement? p, string local) =>
            p?.ChildElements.FirstOrDefault(e => e.LocalName == local);

        var ext = Child(Child(Child(wgp, "grpSpPr"), "xfrm"), "ext");
        if (ext != null)
        {
            if (ReadUnqualifiedLong(ext, "cx") is { } cx) node.Format["width"] = Core.EmuConverter.FormatEmu(cx);
            if (ReadUnqualifiedLong(ext, "cy") is { } cy) node.Format["height"] = Core.EmuConverter.FormatEmu(cy);
        }
        var anchor = wgp.Ancestors().FirstOrDefault(e => e.LocalName is "anchor" or "inline");
        if (anchor != null)
        {
            long? PosOffset(string dir)
            {
                var off = Child(Child(anchor, dir), "posOffset");
                return long.TryParse(off?.InnerText, out var v) ? v : null;
            }
            if (PosOffset("positionH") is { } x) node.Format["x"] = Core.EmuConverter.FormatEmu(x);
            if (PosOffset("positionV") is { } y) node.Format["y"] = Core.EmuConverter.FormatEmu(y);
        }
        return node;
    }

    private DocumentNode ElementToNode(OpenXmlElement element, string path, int depth)
    {
        var node = new DocumentNode { Path = path, Type = element.LocalName };

        if (element is BookmarkStart bkStart)
            return BookmarkStartToNode(bkStart, node);

        if (element is PermStart permStartEl)
            return PermStartToNode(permStartEl, node);

        if (element is PermEnd permEndEl)
        {
            node.Type = "permEnd";
            if (permEndEl.Id?.Value != null) node.Format["id"] = permEndEl.Id.Value.ToString();
            return node;
        }

        if (element is Footnote fnEl)
            return FootnoteToNode(fnEl, node, path, depth);

        if (element is Endnote enEl)
            return EndnoteToNode(enEl, node, path, depth);

        if (element is Comment comment)
            return CommentToNode(comment, node);

        if (element is SectionProperties sectPrEl)
            return SectionPropertiesToNode(sectPrEl, path);

        if (element is Paragraph para)
            return ParagraphToNode(para, node, path, depth);
        else if (element is Run run)
            return RunToNode(run, node, path);
        else if (element is Hyperlink hyperlink)
            return HyperlinkToNode(hyperlink, node);
        else if (element is Table table)
            return TableToNode(table, node, path, depth);
        else if (element is TableCell directCell)
            return TableCellToNode(directCell, node, path, depth);
        else if (element is TableRow directRow)
            return TableRowToNode(directRow, node, path, depth);
        else if (element is SdtBlock sdtBlockNode)
            return SdtBlockToNode(sdtBlockNode, node);
        else if (element is SdtRun sdtRunNode)
            return SdtRunToNode(sdtRunNode, node);
        else if (element.LocalName == "wgp"
                 && element.NamespaceUri == "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup")
            return WgpGroupToNode(element, node, path);
        else if (element.LocalName == "oMathPara" || element is M.Paragraph)
        {
            node.Type = "equation";
            node.Format["mode"] = "display";
            // BUG-DUMP19-02: surface m:oMathParaPr/m:jc as Format["align"] so
            // block-equation alignment round-trips. Without this the value is
            // silently dropped on read-back.
            var mathPPr = element.GetFirstChild<M.ParagraphProperties>();
            var jcVal = mathPPr?.Justification?.Val?.InnerText;
            if (!string.IsNullOrEmpty(jcVal))
            {
                node.Format["align"] = jcVal switch
                {
                    "centerGroup" => "centerGroup",
                    _ => jcVal // "left" | "center" | "right"
                };
            }
            // BUG-DUMP-EQDISPLAY-PPR: a display equation wrapped in <w:p> carries
            // the paragraph's line spacing (e.g. line=360 / 1.5x) and before/after
            // that set the equation line's height. The dump previously surfaced
            // only mode/align/formula, so the wrapper paragraph's spacing was
            // dropped on round-trip — the equation collapsed to single spacing,
            // compressing the page and drifting later content across boundaries.
            // Forward the wrapper pPr spacing so TryEmitDisplayEquation + AddEquation
            // can re-apply it to the rebuilt wrapper paragraph.
            if (element.Parent is Paragraph eqWrapP && eqWrapP.ParagraphProperties is { } eqWrapPpr)
            {
                // Granular spacing keys are kept for human-readable round-trips
                // and back-compat, but the wrapper paragraph also carries a
                // paragraph-mark <w:rPr> (font on the ¶ mark) and pStyle that
                // co-determine the equation line's height. Re-applying only
                // spacing+jc while dropping the mark rPr changed the line box and
                // drifted pagination WORSE than dropping pPr entirely. Carry the
                // whole pPr verbatim so AddEquation can restore it intact
                // (CONSISTENCY(verbatim-ppr-supersede): same pattern as chart
                // spPr / paragraph pPr verbatim round-trips).
                node.Format["wrapperPpr"] = eqWrapPpr.OuterXml;
                if (eqWrapPpr.SpacingBetweenLines is { } eqSp)
                {
                    if (eqSp.Before?.Value != null)
                        node.Format["spaceBefore"] = SpacingConverter.FormatWordSpacing(eqSp.Before.Value);
                    if (eqSp.After?.Value != null)
                        node.Format["spaceAfter"] = SpacingConverter.FormatWordSpacing(eqSp.After.Value);
                    if (eqSp.Line?.Value != null)
                        node.Format["lineSpacing"] = SpacingConverter.FormatWordLineSpacing(
                            eqSp.Line.Value, eqSp.LineRule?.InnerText);
                    if (eqSp.LineRule?.HasValue == true)
                        node.Format["lineRule"] = eqSp.LineRule.InnerText;
                }
                if (eqWrapPpr.Justification?.Val?.InnerText is { Length: > 0 } eqWrapJc)
                    node.Format["wrapperAlign"] = eqWrapJc == "both" ? "justify" : eqWrapJc;
            }
            // Extract LaTeX via FormulaParser
            var oMath = element.Descendants<M.OfficeMath>().FirstOrDefault();
            if (oMath != null)
            {
                // BUG-DUMP-EQVERBATIM (display): carry the verbatim <m:oMath> so
                // AddEquation rebuilds from it instead of the lossy LaTeX string,
                // preserving every math-run <w:rPr> (rFonts="Cambria Math", sizes).
                node.Format["xml"] = oMath.OuterXml;
                try { node.Text = Core.FormulaParser.ToLatex(oMath); }
                catch { node.Text = element.InnerText; }
            }
            else
            {
                node.Text = element.InnerText;
            }
        }
        else if (element is M.OfficeMath inlineMath)
            return OfficeMathToNode(inlineMath, node);
        else if (element is Header or Footer)
            return HeaderFooterToNode(element, node, path, depth);
        else if (element is Body bodyNode)
            return BodyToNode(bodyNode, node, path, depth);
        else
        {
            // Generic fallback: collect XML attributes and child val patterns
            foreach (var attr in element.GetAttributes())
                node.Format[attr.LocalName] = attr.Value;
            foreach (var child in element.ChildElements)
            {
                if (child.ChildElements.Count == 0)
                {
                    foreach (var attr in child.GetAttributes())
                    {
                        if (attr.LocalName.Equals("val", StringComparison.OrdinalIgnoreCase))
                        {
                            node.Format[child.LocalName] = attr.Value;
                            break;
                        }
                    }
                }
            }

            var innerText = element.InnerText;
            if (!string.IsNullOrEmpty(innerText))
                node.Text = innerText.Length > 200 ? innerText[..200] + "..." : innerText;
            if (string.IsNullOrEmpty(innerText))
            {
                var outerXml = element.OuterXml;
                node.Preview = outerXml.Length > 200 ? outerXml[..200] + "..." : outerXml;
            }

            node.ChildCount = element.ChildElements.Count;
            if (depth > 0)
            {
                var typeCounters = new Dictionary<string, int>();
                foreach (var child in element.ChildElements)
                {
                    var name = child.LocalName;
                    typeCounters.TryGetValue(name, out int idx);
                    node.Children.Add(ElementToNode(child, $"{path}/{name}[{idx + 1}]", depth - 1));
                    typeCounters[name] = idx + 1;
                }
            }
        }

        return node;
    }

    private static void ReadRowProps(TableRow row, DocumentNode node)
    {
        // BUG-DUMP-R24-4: per-row <w:tblPrEx> (table property exceptions) —
        // border / jc / indent / layout / cellMar overrides that apply to THIS
        // row only (rows that differ from the table default). Capture the
        // verbatim element so every CT_TblPrEx child round-trips losslessly; the
        // row apply path (SetElementTableRow) re-inserts it as the row's first
        // child. Read before the trPr early-return so a row with tblPrEx but no
        // trPr still round-trips its exceptions.
        var tblPrEx = row.GetFirstChild<TablePropertyExceptions>();
        if (tblPrEx != null && !string.IsNullOrEmpty(tblPrEx.InnerXml))
            node.Format["tblPrEx"] = tblPrEx.OuterXml;
        var trPr = row.TableRowProperties;
        if (trPr == null) return;
        // trPrChange: `set row + trackChange.author` snapshots prior trPr.
        // Surfaced under trPrChange.* for round-trip; EmitTable emits a
        // follow-up `set tr[N]` with trackChange.author/date to reproduce.
        var trPrChange = trPr.GetFirstChild<TableRowPropertiesChange>();
        if (trPrChange != null)
        {
            if (!string.IsNullOrEmpty(trPrChange.Author?.Value))
                node.Format["trPrChange.author"] = trPrChange.Author!.Value!;
            if (trPrChange.Date?.Value is DateTime trDate)
                node.Format["trPrChange.date"] = trDate.ToString("o");
            // BUG-DUMP-R43-9: carry the verbatim prior-trPr snapshot.
            var trPrev = trPrChange.GetFirstChild<PreviousTableRowProperties>();
            if (trPrev != null && trPrev.HasChildren)
                node.Format["trPrChange.beforeXml"] = trPrev.OuterXml;
        }
        // BUG-DUMP-R40-6: row-level tracked-change marker. <w:trPr><w:ins>/<w:del>
        // marks the whole row as inserted/deleted with track-changes on (CT_TrPr,
        // distinct from the run-level InsertedRun/DeletedRun wrapper). Previously
        // unread, so the marker vanished on dump→batch and the inserted/deleted
        // row lost its revision attribution. Surface via the same canonical
        // revision.* creation keys the run reader uses (revision.type=ins|del +
        // author/date/id) so the row emitter can re-emit <w:trPr><w:ins>/<w:del>.
        var rowIns = trPr.GetFirstChild<Inserted>();
        var rowDel = rowIns == null ? trPr.GetFirstChild<Deleted>() : null;
        if (rowIns != null)
        {
            node.Format["revision.type"] = "ins";
            if (!string.IsNullOrEmpty(rowIns.Author?.Value))
                node.Format["revision.author"] = rowIns.Author!.Value!;
            if (rowIns.Date?.Value is DateTime rowInsDate)
                node.Format["revision.date"] = rowInsDate.ToString("o");
            if (rowIns.Id?.Value is { } rowInsId)
                node.Format["revision.id"] = rowInsId.ToString();
        }
        else if (rowDel != null)
        {
            node.Format["revision.type"] = "del";
            if (!string.IsNullOrEmpty(rowDel.Author?.Value))
                node.Format["revision.author"] = rowDel.Author!.Value!;
            if (rowDel.Date?.Value is DateTime rowDelDate)
                node.Format["revision.date"] = rowDelDate.ToString("o");
            if (rowDel.Id?.Value is { } rowDelId)
                node.Format["revision.id"] = rowDelId.ToString();
        }
        var rh = trPr.GetFirstChild<TableRowHeight>();
        // BUG-DUMP-R34-FLOATHEIGHT: some editors write a NON-INTEGER trHeight
        // (w:val="1821.8200000000002"). The SDK types @w:val as UInt16, so
        // touching rh.Val.Value parses the string and throws FormatException —
        // aborting the entire dump. Read the raw InnerText and parse tolerantly,
        // rounding a fractional twip to the nearest integer.
        long? rhTwips = null;
        if (rh?.Val?.InnerText is { Length: > 0 } rhRaw)
        {
            if (long.TryParse(rhRaw, System.Globalization.NumberStyles.Integer,
                    System.Globalization.CultureInfo.InvariantCulture, out var rhExact))
                rhTwips = rhExact;
            else if (double.TryParse(rhRaw, System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture, out var rhFloat))
                rhTwips = (long)Math.Round(rhFloat);
        }
        if (rhTwips != null)
        {
            // BUG-DUMP-R25-1: emit row height as RAW TWIPS ("{n}dxa") not
            // 2-decimal cm. cm round-tripping drifted the val (302→300,
            // 734→731, …) because Round(twips*2.54/1440) loses precision.
            // dxa is the same exact-twip convention already used for
            // colWidths/width readback (ParseTwips strips the "dxa" suffix
            // with no scaling). Set still accepts "2cm" on the input side.
            node.Format["height"] = rhTwips.Value + "dxa";
            // BUG-DUMP-R25-1: round-trip the height rule faithfully. docx
            // CT_Height @w:hRule defaults to "auto" when absent — Word treats
            // an absent hRule as auto row-sizing. Only emit height.rule when
            // the source actually carried an explicit exact/atLeast; emitting
            // it for auto would let Add/Set inject a spurious atLeast.
            if (rh?.HeightType?.Value == HeightRuleValues.Exact)
                node.Format["height.rule"] = "exact";
            else if (rh?.HeightType?.Value == HeightRuleValues.AtLeast)
                node.Format["height.rule"] = "atLeast";
        }
        // BUG-DUMP-R35-TRBOOL: these are CT_TrPr on/off toggles (CT_OnOff). An
        // element with `w:val="0"` (or "false"/"off") means the toggle is OFF —
        // NOT the same as the bare element, which is ON. Reading "present →
        // true" flipped an explicit `<w:tblHeader w:val="0"/>` to header=true,
        // so dump→batch re-emitted a bare `<w:tblHeader/>` (= ON). A first row
        // wrongly marked tblHeader is treated by Word as a repeating header that
        // it refuses to orphan at a page bottom, pushing the whole table to the
        // next page (a blank page + reflow on a Canva-style title-block table).
        // The bare element (no w:val) is ON; an explicit "0"/"false"/"off" is
        // OFF — leave the key unset when OFF (false = absent default) so dump→
        // batch never re-emits a bare ON element. See IsToggleOn(OpenXmlElement?).
        if (IsToggleOn(trPr.GetFirstChild<TableHeader>()))
            node.Format["header"] = true;
        if (IsToggleOn(trPr.GetFirstChild<CantSplit>()))
            node.Format["cantSplit"] = true;
        // BUG-DUMP-R37-3: <w:hidden/> marks the whole row not displayed/printed
        // (CT_TrPr). Previously unread — a hidden row reappeared on dump→batch.
        // Mirror the header/cantSplit toggle reads; Add/Set grow matching cases.
        if (IsToggleOn(trPr.GetFirstChild<Hidden>()))
            node.Format["hidden"] = true;
        // BUG-DUMP-R24-1: row-level <w:jc> in <w:trPr> horizontally positions
        // the WHOLE ROW on the page (CT_TrPr). Distinct from table-level
        // tblPr/jc and from cell/paragraph alignment — use canonical key
        // `rowAlign` so it never collides with cell/para `alignment`.
        var rowJc = trPr.GetFirstChild<TableJustification>();
        if (rowJc?.Val?.Value is TableRowAlignmentValues rowAlignVal)
        {
            if (rowAlignVal == TableRowAlignmentValues.Center) node.Format["rowAlign"] = "center";
            else if (rowAlignVal == TableRowAlignmentValues.Right) node.Format["rowAlign"] = "right";
            else if (rowAlignVal == TableRowAlignmentValues.Left) node.Format["rowAlign"] = "left";
        }
        // cnfStyle (conditional-formatting bitmask) — mirror the cell reader
        // at ~line 4131 so table-row conditional formatting round-trips.
        var rowCnf = trPr.GetFirstChild<ConditionalFormatStyle>();
        if (rowCnf?.Val?.Value is string rowCnfVal && !string.IsNullOrEmpty(rowCnfVal))
            node.Format["cnfStyle"] = rowCnfVal;
        // BUG-DUMP-R42-2: <w:gridBefore>/<w:gridAfter> (column-count skips) and
        // their paired preferred widths <w:wBefore>/<w:wAfter> (CT_TrPr). A row
        // skips N leading/trailing grid columns and reserves a preferred width
        // for the skipped span — produces a ragged/indented table edge. Were
        // previously unread, so dump→batch dropped the skip and collapsed the
        // ragged edge into a full-width row. gridBefore/gridAfter carry an
        // integer @w:val column count; wBefore/wAfter carry @w:w (twips) + @w:type
        // (dxa/pct/auto/nil) — surfaced through the same unit-qualified width form
        // the cell tcW reader emits (see FormatTableWidth) so the type round-trips
        // losslessly. Add/Set grow matching cases (Add.Table.cs, Set.Element.cs).
        var gridBefore = trPr.GetFirstChild<GridBefore>();
        if (gridBefore?.Val?.Value is { } gbVal)
            node.Format["gridBefore"] = gbVal.ToString(System.Globalization.CultureInfo.InvariantCulture);
        var wBefore = trPr.GetFirstChild<WidthBeforeTableRow>();
        if (wBefore != null && FormatTableWidth(wBefore.Width, wBefore.Type?.Value) is { } wBeforeStr)
            node.Format["wBefore"] = wBeforeStr;
        var gridAfter = trPr.GetFirstChild<GridAfter>();
        if (gridAfter?.Val?.Value is { } gaVal)
            node.Format["gridAfter"] = gaVal.ToString(System.Globalization.CultureInfo.InvariantCulture);
        var wAfter = trPr.GetFirstChild<WidthAfterTableRow>();
        if (wAfter != null && FormatTableWidth(wAfter.Width, wAfter.Type?.Value) is { } wAfterStr)
            node.Format["wAfter"] = wAfterStr;
        // BUG-DUMP-R62-ROWCELLSPACING: row-level <w:trPr><w:tblCellSpacing> sets
        // the spacing BETWEEN cells for THIS row (CT_TrPr) — distinct from the
        // table-level tblPr/tblCellSpacing read at ~line 2874. Previously unread,
        // so a form-table whose every row carries cellSpacing="20" collapsed
        // flush on dump→batch: each row shed its inter-cell gap, and the lost
        // per-row height accumulated into a multi-row vertical drift that reflowed
        // the whole document (all pages went red). Surface under the same
        // `cellSpacing` key the table reader uses — the row node is distinct, so
        // there's no collision — and let SetElementTableRow + RowOnlyKeys
        // round-trip it.
        if (trPr.GetFirstChild<TableCellSpacing>() is { } rowCellSpacing
            && SafeWidth(rowCellSpacing.Width) is int rowCsW)
            node.Format["cellSpacing"] = rowCsW;
    }

    // BUG-DUMP-R42-2 / BUG-DUMP-R42-6: shared width readback for OOXML
    // CT_TblWidth-shaped elements (tcW / wBefore / wAfter). Mirrors the cell
    // tcW reader: pct stored as fifths-of-percent ('N%'), auto/nil round-trip
    // as their bare type names, dxa as '{twips}dxa'. nil is a DISTINCT value
    // from '0dxa' ("no preferred width" vs "1-twip explicit") — returning the
    // literal "nil" keeps BUG-DUMP-R42-6's nil cell width from collapsing to
    // dxa. Returns null when @w:w can't be parsed and the type isn't auto/nil.
    private static string? FormatTableWidth(StringValue? rawWidth, TableWidthUnitValues? type)
    {
        if (type == TableWidthUnitValues.Nil) return "nil";
        if (type == TableWidthUnitValues.Auto) return "auto";
        var w = SafeWidth(rawWidth);
        if (w is not int twips) return null;
        if (type == TableWidthUnitValues.Pct)
            return FormatPctWidth(twips);
        return twips.ToString(System.Globalization.CultureInfo.InvariantCulture) + "dxa";
    }

    // OOXML stores pct widths in fifths-of-a-percent (5000 = 100%), so the
    // exact percentage always fits in two decimals. Integer division here
    // (720/50 → "14%") shaved up to 0.98% off every column on round-trip,
    // which re-wraps cell text and reflows whole pages.
    private static string FormatPctWidth(int fiftieths) =>
        (fiftieths / 50.0).ToString("0.##", System.Globalization.CultureInfo.InvariantCulture) + "%";

    private static void ReadCellProps(TableCell cell, DocumentNode node)
    {
        var tcPr = cell.TableCellProperties;
        if (tcPr != null)
        {
            // tcPrChange: `set cell + trackChange.author` snapshots prior tcPr.
            // Surfaced under tcPrChange.* for round-trip; EmitTable emits a
            // follow-up `set tr[N]/tc[M]` with trackChange.author/date.
            var tcPrChange = tcPr.GetFirstChild<TableCellPropertiesChange>();
            if (tcPrChange != null)
            {
                if (!string.IsNullOrEmpty(tcPrChange.Author?.Value))
                    node.Format["tcPrChange.author"] = tcPrChange.Author!.Value!;
                if (tcPrChange.Date?.Value is DateTime tcDate)
                    node.Format["tcPrChange.date"] = tcDate.ToString("o");
                // BUG-DUMP-R43-9: carry the verbatim prior-tcPr snapshot so the
                // Table emitter restores it (via revision.beforeXml) instead of
                // re-stamping an empty <w:tcPr/> marker.
                var tcPrev = tcPrChange.GetFirstChild<PreviousTableCellProperties>();
                if (tcPrev != null && tcPrev.HasChildren)
                    node.Format["tcPrChange.beforeXml"] = tcPrev.OuterXml;
            }
            // BUG-DUMP-R51-2: cell-level tracked insertion/deletion
            // (<w:tcPr><w:cellIns>/<w:cellDel>) — marks a whole table cell as
            // inserted/deleted under Track Changes (distinct from row-level
            // <w:trPr><w:ins>/<w:del> and from run-level revisions on the cell's
            // content). Previously unread, so a tracked cell insert/delete was
            // dropped on dump→batch while row-level trIns/trDel survived. Surface
            // it as cellRevision.type=ins|del + cellRevision.author/.date/.id so
            // the Table emitter can re-stamp the marker; a distinct key namespace
            // keeps it clear of the run-content revision.* on the cell's runs.
            var cellIns = tcPr.GetFirstChild<CellInsertion>();
            var cellDel = cellIns == null ? tcPr.GetFirstChild<CellDeletion>() : null;
            if (cellIns != null)
            {
                node.Format["cellRevision.type"] = "ins";
                if (!string.IsNullOrEmpty(cellIns.Author?.Value))
                    node.Format["cellRevision.author"] = cellIns.Author!.Value!;
                if (cellIns.Date?.Value is DateTime cellInsDate)
                    node.Format["cellRevision.date"] = cellInsDate.ToString("o");
                if (cellIns.Id?.Value is { } cellInsId)
                    node.Format["cellRevision.id"] = cellInsId.ToString();
            }
            else if (cellDel != null)
            {
                node.Format["cellRevision.type"] = "del";
                if (!string.IsNullOrEmpty(cellDel.Author?.Value))
                    node.Format["cellRevision.author"] = cellDel.Author!.Value!;
                if (cellDel.Date?.Value is DateTime cellDelDate)
                    node.Format["cellRevision.date"] = cellDelDate.ToString("o");
                if (cellDel.Id?.Value is { } cellDelId)
                    node.Format["cellRevision.id"] = cellDelId.ToString();
            }
            // Borders (including diagonal)
            var cb = tcPr.TableCellBorders;
            if (cb != null)
            {
                ReadBorder(cb.TopBorder, "border.top", node);
                ReadBorder(cb.BottomBorder, "border.bottom", node);
                ReadBorder(cb.LeftBorder, "border.left", node);
                ReadBorder(cb.RightBorder, "border.right", node);
                ReadBorder(cb.TopLeftToBottomRightCellBorder, "border.tl2br", node);
                ReadBorder(cb.TopRightToBottomLeftCellBorder, "border.tr2bl", node);
            }
            // Shading — check for gradient (w14:gradFill in mc:AlternateContent) first
            var mcNs = "http://schemas.openxmlformats.org/markup-compatibility/2006";
            var gradAc = tcPr.ChildElements
                .FirstOrDefault(e => e.LocalName == "AlternateContent" && e.NamespaceUri == mcNs);
            if (gradAc != null && gradAc.InnerXml.Contains("gradFill"))
            {
                // Parse gradient colors and angle from w14:gradFill XML
                var colors = new List<string>();
                foreach (var match in System.Text.RegularExpressions.Regex.Matches(
                    gradAc.InnerXml, @"val=""([0-9A-Fa-f]{6})"""))
                {
                    colors.Add(((System.Text.RegularExpressions.Match)match).Groups[1].Value);
                }
                var angleMatch = System.Text.RegularExpressions.Regex.Match(
                    gradAc.InnerXml, @"ang=""(\d+)""");
                var angle = angleMatch.Success ? int.Parse(angleMatch.Groups[1].Value) / 60000.0 : 0.0;
                var angleStr = angle % 1 == 0 ? $"{(int)angle}" : $"{angle:0.##}";
                if (colors.Count >= 2)
                {
                    node.Format["fill"] = $"gradient;{ParseHelpers.FormatHexColor(colors[0])};{ParseHelpers.FormatHexColor(colors[1])};{angleStr}";
                }
                else if (colors.Count == 1)
                {
                    node.Format["fill"] = ParseHelpers.FormatHexColor(colors[0]);
                }
            }
            else
            {
                var shd = tcPr.Shading;
                if (shd != null)
                {
                    // The cell help schema declares `fill` as the canonical key
                    // (set:true get:true, readback "#RRGGBB uppercase, or
                    // 'gradient'") with shd/shading only as Set-side aliases.
                    // A solid cell background is <w:shd w:val="clear"|"solid"
                    // w:fill="RRGGBB"/> — fully expressible as a single `fill`
                    // value, so emit the canonical key (matches sibling
                    // color/align/valign round-trip; mirrors the gradient branch
                    // above which already emits `fill`). The gradient/solidFill
                    // branch above handles synthetic gradients.
                    //
                    // A real pattern shading (w:val = pct*/stripe/cross), a
                    // separate pattern Color, or theme-linkage attrs cannot be
                    // collapsed into one solid color — those keep the
                    // shading.val/.fill/.color/.theme* detail keys (consumed by
                    // the dump→batch fold in WordBatchEmitter.Filters.cs). When
                    // shading.* detail is present, ExtractCellOnlyProps drops the
                    // `fill` alias so they don't double-apply (BUG-DUMP21-02).
                    //
                    // <w:shd w:val="clear" w:fill="auto"/> is OOXML's "no
                    // shading" — emit nothing (matches a cell with no shd).
                    var cShdVal = shd.Val?.InnerText;
                    var cShdFill = shd.Fill?.Value;
                    var cShdColor = shd.Color?.Value;
                    bool hasFillColor = !string.IsNullOrEmpty(cShdFill)
                        && !string.Equals(cShdFill, "auto", StringComparison.OrdinalIgnoreCase);
                    bool isSolidVal = string.IsNullOrEmpty(cShdVal)
                        || string.Equals(cShdVal, "clear", StringComparison.OrdinalIgnoreCase)
                        || string.Equals(cShdVal, "solid", StringComparison.OrdinalIgnoreCase);
                    bool hasPatternColor = !string.IsNullOrEmpty(cShdColor);
                    bool hasTheme = shd.ThemeFill?.HasValue == true
                        || shd.ThemeFillShade?.Value != null || shd.ThemeFillTint?.Value != null
                        || shd.ThemeColor?.HasValue == true
                        || shd.ThemeShade?.Value != null || shd.ThemeTint?.Value != null;

                    // <w:shd w:val="clear" w:fill="auto"/> (and bare clear/solid
                    // with no fill color, no pattern color, no theme) is OOXML's
                    // "no shading" form — emit nothing, identical to a cell with
                    // no <w:shd> at all (mirrors the batch-emitter
                    // shadingIsEffectivelyNone skip).
                    bool effectivelyNone = isSolidVal && !hasFillColor
                        && !hasPatternColor && !hasTheme;

                    if (effectivelyNone)
                    {
                        // intentionally emit no key
                    }
                    else if (isSolidVal && hasFillColor && !hasPatternColor && !hasTheme)
                    {
                        node.Format["fill"] = ParseHelpers.FormatHexColor(cShdFill!);
                    }
                    else
                    {
                        // Pattern / theme / pattern-color cell: keep the detail
                        // keys verbatim (unchanged from before — emits shading.fill
                        // even for the "auto" sentinel so the dump round-trip sees
                        // the same shape).
                        if (!string.IsNullOrEmpty(cShdVal)) node.Format["shading.val"] = cShdVal;
                        if (!string.IsNullOrEmpty(cShdFill)) node.Format["shading.fill"] = ParseHelpers.FormatHexColor(cShdFill);
                        if (hasPatternColor) node.Format["shading.color"] = ParseHelpers.FormatHexColor(cShdColor!);
                        ReadShadingTheme(shd, node);
                    }
                }
            }
            // Width
            // BUG-DUMP6-04: preserve w:tcW @type semantics. Mirror the table-level
            // width readback above (line ~1930) — pct widths are stored as
            // fifths-of-percent, so divide by 50 and append '%' so dump→batch
            // can recognize and re-emit pct cell widths.
            // BUG-R4-05: emit width with explicit unit suffix (dxa/%) — root
            // the project conventions mandates unit-qualified width readback. Bare integer
            // ("3000") is the historic bug.
            // BUG-R4B(BUG1): decimal-tolerant cell-width read.
            if (SafeWidth(tcPr.TableCellWidth?.Width) is int cwRaw)
            {
                var cwType = tcPr.TableCellWidth!.Type?.Value;
                // BUG-DUMP-R42-6: type=nil ("no preferred width") is a DISTINCT
                // value from "0dxa" (1-twip explicit). Previously both nil and
                // a zero @w:w collapsed to "0dxa", so a nil cell width round-
                // tripped to <w:tcW w:w="1" w:type="dxa"/> — changing the width
                // semantics. Route through FormatTableWidth so nil/auto surface
                // as their bare type names; only a true dxa zero stays "0dxa".
                if (cwType == TableWidthUnitValues.Nil)
                    node.Format["width"] = "nil";
                else if (cwType == TableWidthUnitValues.Pct)
                    node.Format["width"] = FormatPctWidth(cwRaw);
                else if (cwType == TableWidthUnitValues.Auto)
                    node.Format["width"] = "auto";
                else if (cwRaw == 0)
                    node.Format["width"] = "0dxa";
                else
                    node.Format["width"] = cwRaw.ToString(System.Globalization.CultureInfo.InvariantCulture) + "dxa";
            }
            // Vertical alignment
            if (tcPr.TableCellVerticalAlignment?.Val?.Value != null)
                node.Format["valign"] = tcPr.TableCellVerticalAlignment.Val.InnerText;
            // Vertical merge
            if (tcPr.VerticalMerge != null)
                node.Format["vmerge"] = tcPr.VerticalMerge.Val?.Value == MergedCellValues.Restart ? "restart" : "continue";
            // Horizontal merge — same toggle pattern as vmerge: ST_Merge val=restart
            // marks the leading cell of a horizontal span, bare <w:hMerge/> marks the
            // continuation cells. Without this read block dump→batch silently dropped
            // every horizontal span on round-trip.
            if (tcPr.HorizontalMerge != null)
                node.Format["hmerge"] = tcPr.HorizontalMerge.Val?.Value == MergedCellValues.Restart ? "restart" : "continue";
            // Grid span
            if (tcPr.GridSpan?.Val?.Value != null && tcPr.GridSpan.Val.Value > 1)
                node.Format["colspan"] = tcPr.GridSpan.Val.Value;
            // Cell padding/margins
            var mar = tcPr.TableCellMargin;
            if (mar != null)
            {
                // BUG-R4B(BUG1): decimal-tolerant cell-margin reads.
                // BUG-DUMP-R37-2: logical start/end margins (<w:start>/<w:end>,
                // the bidi-aware spelling) were silently dropped — only the
                // physical <w:left>/<w:right> variants were read. Fall back to
                // StartMargin/EndMargin when the physical slot is absent,
                // mirroring the HtmlPreview CSS reader (LeftMargin ?? StartMargin).
                if (SafeWidth(mar.TopMargin?.Width) is int mT) node.Format["padding.top"] = mT;
                if (SafeWidth(mar.BottomMargin?.Width) is int mB) node.Format["padding.bottom"] = mB;
                if (SafeWidth(mar.LeftMargin?.Width ?? mar.StartMargin?.Width) is int mL) node.Format["padding.left"] = mL;
                if (SafeWidth(mar.RightMargin?.Width ?? mar.EndMargin?.Width) is int mR) node.Format["padding.right"] = mR;
            }
            // Text direction
            if (tcPr.TextDirection?.Val?.Value != null)
                node.Format["textDirection"] = tcPr.TextDirection.Val.InnerText;
            // No wrap (CT_OnOff — honor an explicit w:val="0" = OFF)
            if (IsToggleOn(tcPr.NoWrap))
                node.Format["nowrap"] = true;
            // BUG-R3-03: cnfStyle (conditional formatting bitfield).
            var cnfRead = tcPr.GetFirstChild<ConditionalFormatStyle>();
            if (cnfRead?.Val?.Value is string cnfVal && !string.IsNullOrEmpty(cnfVal))
                node.Format["cnfStyle"] = cnfVal;
            // BUG-DUMP-CELLTAIL: the cell reader is a curated list with no
            // FillUnknownChildProps long-tail (unlike run rPr / paragraph pPr),
            // so any tcPr child outside the curated set was silently dropped on
            // dump→batch. Surface the two common toggles explicitly (mirrors the
            // row-level cantSplit/tblHeader reads); Add/Set already support both.
            if (IsToggleOn(tcPr.GetFirstChild<HideMark>()))
                node.Format["hideMark"] = true;
            if (IsToggleOn(tcPr.GetFirstChild<TableCellFitText>()))
                node.Format["tcFitText"] = true;
            // BUG-DUMP-R32-3: <w:cellMerge> is a tracked-change marker (a cell
            // split/merge made under Track Changes) carrying
            // vMerge/vMergeOrig/id/author/date. It's neither a curated tcPr key
            // nor reachable through any FillUnknownChildProps fallback, so it was
            // dropped SILENTLY on dump→batch — a tracked-change loss (unlike
            // pPrChange/rPrChange which warn). Surface its verbatim OuterXml so
            // the Table emitter can re-apply it via a raw-set into the cell's
            // tcPr. Schema: cellMerge has no typed val/attr the curated setters
            // model, so the raw XML is the faithful round-trip carrier.
            if (tcPr.GetFirstChild<CellMerge>() is { } cellMergeEl)
                node.Format["cellMerge.xml"] = cellMergeEl.OuterXml;
        }
        // BUG-R4-05: when no per-cell tcW is set, synthesize width from the
        // parent table's tblGrid/gridCol so Get always exposes a unit-qualified
        // width (matches the cross-handler width contract). CONSISTENCY(add-set-symmetry):
        // Add intentionally does not stamp per-cell tcW (BUG-R6-06) — width
        // lives in tblGrid as the schema intends — so Get must back-fill.
        if (!node.Format.ContainsKey("width"))
        {
            var parentTbl = cell.Ancestors<Table>().FirstOrDefault();
            var parentRow = cell.Parent as TableRow;
            if (parentTbl != null && parentRow != null)
            {
                // BUG-DUMP-GRIDIDX: the cell's GRID-COLUMN index is NOT its
                // ordinal position in the row — a preceding cell with
                // gridSpan>1 occupies multiple grid columns, so the starting
                // grid column is the SUM of preceding cells' spans. Using the
                // raw ordinal mis-reads the column width for every cell after
                // a horizontally-merged one (and made the derived row total
                // exceed tblGrid, overflowing the page on rebuild).
                var rowCells = GetRowCellsFlattened(parentRow);
                var cellPos = rowCells.IndexOf(cell);
                var cellIdx = 0;
                for (int ci = 0; ci < cellPos; ci++)
                    cellIdx += (int)(rowCells[ci].TableCellProperties?.GridSpan?.Val?.Value ?? 1);
                var gridCols = parentTbl.GetFirstChild<TableGrid>()?.Elements<GridColumn>().ToList();
                if (gridCols != null && cellPos >= 0 && cellIdx < gridCols.Count)
                {
                    // Account for gridSpan — sum spanned cols.
                    var span = (tcPr?.GridSpan?.Val?.Value ?? 1);
                    long total = 0;
                    for (int gi = cellIdx; gi < Math.Min(cellIdx + span, gridCols.Count); gi++)
                    {
                        if (uint.TryParse(gridCols[gi].Width?.Value, out var gv))
                            total += gv;
                    }
                    if (total > 0)
                    {
                        node.Format["width"] = total + "dxa";
                        // BUG-DUMP-AUTOFITW: flag this width as derived from
                        // tblGrid (the cell had NO source <w:tcW>). EmitTable
                        // suppresses the fabricated width for AUTOFIT tables so
                        // Word's column solver keeps inheriting from tblGrid
                        // instead of being over-constrained by a synthetic tcW
                        // (which shifts boundaries, worst with gridSpan/vMerge).
                        // Internal-only marker (underscore-prefixed); the plain
                        // `get` width readback is unchanged.
                        node.Format["_widthDerived"] = true;
                    }
                }
            }
        }
        // Alignment from first paragraph
        var firstPara = cell.Elements<Paragraph>().FirstOrDefault();
        var just = firstPara?.ParagraphProperties?.Justification?.Val;
        if (just != null)
            node.Format["align"] = just.InnerText;
        // Direction: <w:bidi/> on the first cell paragraph maps to canonical
        // direction=rtl. Mirrors paragraph readback canonical key. R20-bt-2:
        // also surface direction=rtl when the enclosing table carries
        // <w:bidiVisual/> on tblPr — cells inherit table-level visual RTL
        // even without their own pPr.bidi.
        if (firstPara?.ParagraphProperties?.BiDi != null)
            node.Format["direction"] = "rtl";
        else if (cell.Ancestors<Table>().FirstOrDefault()
                     ?.GetFirstChild<TableProperties>()?.GetFirstChild<BiDiVisual>() != null)
            node.Format["direction"] = "rtl";
        // Run-level formatting from first run (mirrors PPTX table cell behavior)
        var firstRun = cell.Descendants<Run>().FirstOrDefault();
        if (firstRun?.RunProperties != null)
        {
            var rPr = firstRun.RunProperties;
            if (rPr.RunFonts?.Ascii?.Value != null) node.Format["font"] = rPr.RunFonts.Ascii.Value;
            if (rPr.FontSize?.Val?.Value != null) node.Format["size"] = $"{int.Parse(rPr.FontSize.Val.Value) / 2.0:0.##}pt";
            if (rPr.Bold != null) node.Format["bold"] = IsToggleOn(rPr.Bold);
            if (rPr.Italic != null) node.Format["italic"] = IsToggleOn(rPr.Italic);
            if (rPr.Color?.Val?.Value != null) node.Format["color"] = ParseHelpers.FormatHexColor(rPr.Color.Val.Value);
            else if (rPr.Color?.ThemeColor?.HasValue == true) node.Format["color"] = rPr.Color.ThemeColor.InnerText;
            if (rPr.Underline?.Val != null) node.Format["underline"] = rPr.Underline.Val.InnerText;
            // CONSISTENCY(underline-color): backfilled from style Get edc8f884.
            if (rPr.Underline?.Color?.Value != null)
                node.Format["underline.color"] = ParseHelpers.FormatHexColor(rPr.Underline.Color.Value);
            if (rPr.Strike != null) node.Format["strike"] = IsToggleOn(rPr.Strike);
            if (rPr.Highlight?.Val != null) node.Format["highlight"] = rPr.Highlight.Val.InnerText;
        }
    }

    // Surface a <w:pgBorders> element as per-side detail keyed under the
    // pgBorders.<side> prefix (mirrors ReadBorder for paragraph/table borders),
    // plus the offsetFrom position attribute. Page borders have no insideH /
    // insideV sides (CT_PageBorders is top/left/bottom/right only). The Set
    // side re-materialises these into a PageBorders with real per-side
    // val/sz/color/space + offsetFrom — see TrySetSectionLayout pgborders.*.
    private static void ReadPageBorders(PageBorders? pgBorders, DocumentNode node)
    {
        if (pgBorders == null) return;
        ReadBorder(pgBorders.TopBorder, "pgBorders.top", node);
        ReadBorder(pgBorders.LeftBorder, "pgBorders.left", node);
        ReadBorder(pgBorders.BottomBorder, "pgBorders.bottom", node);
        ReadBorder(pgBorders.RightBorder, "pgBorders.right", node);
        // offsetFrom: "page" (border measured from page edge) vs "text"
        // (from text margin). Default in OOXML is "text" — surface only when
        // present so a source that omits it round-trips without re-stamping.
        if (pgBorders.OffsetFrom?.InnerText is { } off)
            node.Format["pgBorders.offsetFrom"] = off;
        // BUG-DUMP-R44-5: zOrder (front/back — border drawn IN FRONT of vs
        // BEHIND text) and display (allPages/firstPage/notFirstPage — which
        // pages the page border appears on) are meaning-changing attributes that
        // were dropped on round-trip. Surface only when present so a source that
        // omits them round-trips without re-stamping the OOXML defaults.
        if (pgBorders.ZOrder?.InnerText is { } z)
            node.Format["pgBorders.zOrder"] = z;
        if (pgBorders.Display?.InnerText is { } disp)
            node.Format["pgBorders.display"] = disp;
    }

    private static void ReadBorder(BorderType? border, string key, DocumentNode node)
    {
        if (border?.Val == null) return;
        // CONSISTENCY(canonical-keys): emit val on the parent key plus .sz/.color/.space sub-keys
        // (matches Excel border.* schema). No compound semicolon-joined string — that was a private
        // encoding that diverged from both OOXML and the rest of the project.
        node.Format[key] = border.Val?.InnerText ?? "none";
        if (border.Size?.Value is uint sz) node.Format[$"{key}.sz"] = sz;
        if (border.Color?.Value is { } c) node.Format[$"{key}.color"] = ParseHelpers.FormatHexColor(c);
        if (border.Space?.Value is uint sp) node.Format[$"{key}.space"] = sp;
        // BUG-DUMP-R36-1: w:shadow (drop-shadow) / w:frame (border on the OUTSIDE
        // edge of text) are render-relevant and were previously dropped on
        // round-trip. Emit as .shadow / .frame sub-keys only when present so a
        // plain border round-trips without a spurious shadow="false".
        if (border.Shadow?.Value is bool sh) node.Format[$"{key}.shadow"] = sh;
        if (border.Frame?.Value is bool fr) node.Format[$"{key}.frame"] = fr;
        // BUG-DUMP-R41-2: w:themeColor / w:themeShade / w:themeTint encode the
        // border's link to the theme color slot (accent1, …) plus the
        // shade/tint applied to it. The resolved w:color hex was kept, but the
        // theme linkage was dropped on round-trip — Word would no longer
        // recolor the border when the document theme changes. Emit each as a
        // sub-key only when present so a plain (non-themed) border keeps the
        // legacy shape. Mirrors the run/shading themeFill readback below.
        if (border.ThemeColor?.HasValue == true) node.Format[$"{key}.themeColor"] = border.ThemeColor.InnerText ?? "";
        if (border.ThemeShade?.Value is { } tsh) node.Format[$"{key}.themeShade"] = tsh;
        if (border.ThemeTint?.Value is { } tt) node.Format[$"{key}.themeTint"] = tt;
    }

    // BUG-DUMP-R41-4: surface the theme-linkage attributes on a <w:shd> under
    // the shading.* namespace so they round-trip alongside the resolved fill.
    // <w:shd> carries TWO theme slots: w:themeFill/themeFillShade/themeFillTint
    // (the linkage for the FILL color) and w:themeColor/themeShade/themeTint
    // (the linkage for the pattern COLOR). The resolved w:fill / w:color hex was
    // kept, but the theme linkage was dropped — Word would no longer recolor
    // the shading when the document theme changes. Emit each sub-key only when
    // present so a plain (non-themed) shading keeps the legacy 3-key shape.
    // WordBatchEmitter's shading fold appends these as `key=val` tail segments;
    // ParseShadingValue strips them and ApplyShadingTheme re-stamps them.
    // CONSISTENCY(shd-canonical-fill): emit a solid <w:shd> background as the
    // canonical `fill` key, matching the table-cell shading reader (~line 5938).
    // A solid background is <w:shd w:val="clear"|"solid" w:fill="RRGGBB"/> —
    // fully expressible as one color, so emit `fill` (#RRGGBB uppercase via
    // FormatHexColor). A real pattern (w:val = pct*/stripe/cross), a separate
    // pattern color, or theme-linkage attrs cannot collapse to one solid color
    // and keep the shading.val/.fill/.color/.theme* detail keys (consumed by the
    // dump→batch fold in WordBatchEmitter.Filters.cs). <w:shd w:val="clear"
    // w:fill="auto"/> ("no shading") emits nothing.
    private static void ReadShadingCanonical(Shading shd, DocumentNode node)
    {
        var shdVal = shd.Val?.InnerText;
        var shdFill = shd.Fill?.Value;
        var shdColor = shd.Color?.Value;
        bool hasFillColor = !string.IsNullOrEmpty(shdFill)
            && !string.Equals(shdFill, "auto", StringComparison.OrdinalIgnoreCase);
        bool isSolidVal = string.IsNullOrEmpty(shdVal)
            || string.Equals(shdVal, "clear", StringComparison.OrdinalIgnoreCase)
            || string.Equals(shdVal, "solid", StringComparison.OrdinalIgnoreCase);
        bool hasPatternColor = !string.IsNullOrEmpty(shdColor);
        bool hasTheme = shd.ThemeFill?.HasValue == true
            || shd.ThemeFillShade?.Value != null || shd.ThemeFillTint?.Value != null
            || shd.ThemeColor?.HasValue == true
            || shd.ThemeShade?.Value != null || shd.ThemeTint?.Value != null;

        bool effectivelyNone = isSolidVal && !hasFillColor && !hasPatternColor && !hasTheme;

        if (effectivelyNone)
        {
            // intentionally emit no key (matches no <w:shd> at all)
        }
        else if (isSolidVal && hasFillColor && !hasPatternColor && !hasTheme)
        {
            node.Format["fill"] = ParseHelpers.FormatHexColor(shdFill!);
        }
        else
        {
            // Pattern / theme / pattern-color: keep the detail keys verbatim.
            if (!string.IsNullOrEmpty(shdVal)) node.Format["shading.val"] = shdVal;
            if (!string.IsNullOrEmpty(shdFill)) node.Format["shading.fill"] = ParseHelpers.FormatHexColor(shdFill);
            if (hasPatternColor) node.Format["shading.color"] = ParseHelpers.FormatHexColor(shdColor!);
            ReadShadingTheme(shd, node);
        }
    }

    private static void ReadShadingTheme(Shading shd, DocumentNode node)
    {
        if (shd.ThemeFill?.HasValue == true) node.Format["shading.themeFill"] = shd.ThemeFill.InnerText ?? "";
        if (shd.ThemeFillShade?.Value is { } tfs) node.Format["shading.themeFillShade"] = tfs;
        if (shd.ThemeFillTint?.Value is { } tft) node.Format["shading.themeFillTint"] = tft;
        if (shd.ThemeColor?.HasValue == true) node.Format["shading.themeColor"] = shd.ThemeColor.InnerText ?? "";
        if (shd.ThemeShade?.Value is { } tsh) node.Format["shading.themeShade"] = tsh;
        if (shd.ThemeTint?.Value is { } tt) node.Format["shading.themeTint"] = tt;
    }

    // OOXML localNames that curated style/paragraph/run readers already map
    // to canonical keys. FillUnknownChildProps skips these so the long-tail
    // fallback doesn't re-expose them under their bare OOXML names alongside
    // the canonical key (e.g. avoid emitting both `bold: true` and `b: true`).
    private static readonly System.Collections.Generic.HashSet<string> CuratedStyleLocalNames =
        new(System.StringComparer.Ordinal)
    {
        // rPr-side (covered by curated style/paragraph/run readers)
        "b", "bCs", "i", "iCs", "sz", "szCs", "u", "color", "strike", "rFonts",
        "highlight", "caps", "smallCaps", "dstrike", "vanish",
        "outline", "shadow", "emboss", "imprint", "noProof", "rtl",
        "vertAlign", "spacing", "shd",
        // BUG-DUMP22-08: <w:bdr/> is multi-attribute (val+sz+color+space).
        // Curated reader emits the colon-encoded compound form; suppress
        // the long-tail fallback so the bare `bdr=single` name doesn't
        // co-emit alongside the canonical encoded value.
        "bdr",
        // BUG-DUMP10-01: <w:eastAsianLayout/> is a multi-attribute element
        // surfaced by the curated reader as eastAsianLayout.vert / .combine
        // dotted keys. Skip the long-tail fallback so it doesn't double-emit
        // the bare element name with a `true` value.
        "eastAsianLayout",
        // pPr-side
        "jc", "ind", "outlineLvl", "widowControl",
        "keepNext", "keepLines", "pageBreakBefore", "contextualSpacing",
        "pBdr", "numPr", "tabs", "pStyle",
        // bidi maps to canonical `direction` in style/paragraph readback;
        // skip the long-tail fallback to avoid emitting both `direction: rtl`
        // and `bidi: true` for the same <w:bidi/> child element.
        "bidi",
        // Container elements covered by the curated paragraph-mark / run-property
        // reader (see paraRp block ~line 1004). Without this, an empty <w:rPr/>
        // left behind by Set bold=false (etc.) would surface as `rPr: true` via
        // the long-tail fallback. fuzz-1.
        "rPr",
        // BUG-R7-09 / F-3: <w:lang/> is a multi-slot element (val=latin /
        // eastAsia / bidi). The curated reader emits each slot as
        // lang.latin / lang.ea / lang.cs. Word/WPS occasionally write a bare
        // <w:lang/> with no attributes as a "reset to default language"
        // sentinel — the long-tail fallback would then surface that as
        // `lang: true`, which Set parses as a BCP-47 tag and rejects with
        // "Invalid BCP-47 'true'". Skip lang here so the canonical .latin/
        // .ea/.cs reader stays the single source of truth.
        "lang",
    };

    // Long-tail OOXML fallback: walk a properties container (rPr/pPr/...) and
    // surface every leaf child whose localName isn't already covered by the
    // curated reader. Shape is symmetric with the Add/Set side:
    //
    //   - child with no attrs            → Format[name] = true
    //     (toggle, matches GenericXmlQuery.TryCreateTypedChild bare-toggle).
    //   - child with one `val` attr only → Format[name] = val
    //     (scalar, matches GenericXmlQuery.TryCreateTypedChild val-leaf).
    //   - child with any other attrs     → Format[name.attr] = value per attr
    //     (dotted, matches TypedAttributeFallback.TrySet single-level shape
    //     `elementLocal.attrLocal`). Every typed attr surfaces, including
    //     `val` when accompanied by other attrs (so themed colors / multi-
    //     slot indents / spacing round-trip in full).
    //
    // Nested-children elements are emitted as raw flag toggles only — the
    // dotted reflection covers leaf attrs, and 3+ segment nested reflection
    // is intentionally out of scope (raw-XML escape handles the deep cases).
    private static void FillUnknownChildProps(OpenXmlElement? container, DocumentNode node)
    {
        if (container == null) return;
        foreach (var child in container.ChildElements)
        {
            var name = child.LocalName;
            if (string.IsNullOrEmpty(name)) continue;
            if (CuratedStyleLocalNames.Contains(name)) continue;
            if (child.ChildElements.Count > 0) continue;

            // <w:cnfStyle> (CT_Cnf) must NOT be decomposed into dotted
            // cnfStyle.firstRow / .lastRow / … keys: the SDK canonicalizes every
            // individual bit attribute into the combined 12-bit @val string on
            // parse, so a decomposed replay can't round-trip — set cnfStyle.lastRow=0
            // reparses to <w:cnfStyle w:val="000000000000"/> and wipes the value
            // accumulated by the earlier bit keys. Emit the canonical combined @val
            // as a single key (matching the table-cell reader at ~line 4129); it
            // replays through one <w:cnfStyle w:val="…"/> set, which is lossless.
            if (child is ConditionalFormatStyle cnf)
            {
                if (!node.Format.ContainsKey(name)
                    && cnf.Val?.Value is string cnfVal && !string.IsNullOrEmpty(cnfVal))
                    node.Format[name] = cnfVal;
                continue;
            }

            var typedAttrs = new System.Collections.Generic.List<DocumentFormat.OpenXml.OpenXmlAttribute>();
            foreach (var a in child.GetAttributes()) typedAttrs.Add(a);

            if (typedAttrs.Count == 0)
            {
                if (!node.Format.ContainsKey(name))
                    node.Format[name] = true;
                continue;
            }

            if (typedAttrs.Count == 1
                && typedAttrs[0].LocalName.Equals("val", System.StringComparison.OrdinalIgnoreCase))
            {
                if (!node.Format.ContainsKey(name))
                {
                    var raw = typedAttrs[0].Value ?? "";
                    // CT_OnOff toggles (kinsoku, snapToGrid, webHidden, suppressLineNumbers,
                    // …) expose a typed Val of type OnOffValue. Emit a canonical bool so the
                    // long-tail toggles match curated booleans (bold/keepNext → true/false)
                    // instead of the raw "1"/"0" the XML carries. Enum/string children
                    // (em, effect, textAlignment, …) keep their literal val.
                    //
                    // Every CT_OnOff element derives from the shared OnOffType base, so a
                    // direct `is` check is equivalent to inspecting Val's runtime type —
                    // and is trim/AOT-safe where GetType().GetProperty("Val") (IL2075) is not.
                    node.Format[name] =
                        child is DocumentFormat.OpenXml.Wordprocessing.OnOffType
                            ? OfficeCli.Core.ParseHelpers.IsTruthySafe(raw)
                            : raw;
                }
                continue;
            }

            // Multi-attribute element → dotted `<name>.<attr>` keys. Symmetric
            // with TypedAttributeFallback.TrySet on the Add/Set side, so
            // dump→replay round-trips through the same reflection path that
            // already accepts `ind.firstLine=240`, `spacing.line=480`, etc.
            foreach (var a in typedAttrs)
            {
                if (string.IsNullOrEmpty(a.LocalName)) continue;
                var key = $"{name}.{a.LocalName}";
                if (node.Format.ContainsKey(key)) continue;
                node.Format[key] = a.Value ?? "";
            }
        }
    }
}
