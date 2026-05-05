// Copyright 2025 OfficeCli (officecli.ai)
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
        if (props.Revision != null) node.Format["revision"] = props.Revision;
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

            if (sectPr.GetFirstChild<TitlePage>() != null)
                node.Format["titlePage"] = true;

            // Section-level RTL (Arabic / Hebrew page direction).
            if (sectPr.GetFirstChild<BiDi>() != null)
                node.Format["direction"] = "rtl";

            // <w:rtlGutter/> places the binding gutter on the right side.
            if (sectPr.GetFirstChild<GutterOnRight>() != null)
                node.Format["rtlGutter"] = true;

            var lnNum = sectPr.GetFirstChild<LineNumberType>();
            if (lnNum != null)
            {
                var countBy = lnNum.CountBy?.Value ?? 1;
                var restartVal = lnNum.Restart?.InnerText ?? "continuous";
                node.Format["lineNumbers"] = restartVal switch
                {
                    "newPage" => "restartPage",
                    "newSection" => "restartSection",
                    _ => "continuous"
                };
                if (countBy != 1) node.Format["lineNumberCountBy"] = countBy;
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

        foreach (var part in parts)
        {
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

                var name = Core.PathAliases.Resolve(part[..bracketIdx]);
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
            }
            else
            {
                segments.Add(new PathSegment(Core.PathAliases.Resolve(part), null));
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
                OpenXmlElement? cur = charts[n - 1].Inline;
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
        if (first.Name.ToLowerInvariant() == "formfield" && segments.Count == 1 && first.Index.HasValue)
        {
            var allFf = FindFormFields();
            var n = first.Index.Value;
            if (n >= 1 && n <= allFf.Count)
            {
                var beginRun = allFf[n - 1].Field.BeginRun;
                // Walk up to the nearest Paragraph so the anchor is a direct
                // child of the body (matching what the user typically passes
                // as --parent /body). If no paragraph ancestor (shouldn't
                // happen for a valid form field), fall back to the begin run.
                OpenXmlElement? cur = beginRun;
                while (cur != null && cur is not Paragraph) cur = cur.Parent;
                return cur ?? beginRun;
            }
        }

        OpenXmlElement? current = first.Name.ToLowerInvariant() switch
        {
            "body" => _doc.MainDocumentPart?.Document?.Body,
            "styles" => _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles,
            "header" => _doc.MainDocumentPart?.HeaderParts.ElementAtOrDefault((first.Index ?? 1) - 1)?.Header,
            "footer" => _doc.MainDocumentPart?.FooterParts.ElementAtOrDefault((first.Index ?? 1) - 1)?.Footer,
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
                var flat = new List<OpenXmlElement>();
                void CollectBodyChildren(OpenXmlElement el)
                {
                    foreach (var c in el.ChildElements)
                    {
                        if (c is CustomXmlBlock cx) CollectBodyChildren(cx);
                        else flat.Add(c);
                    }
                }
                CollectBodyChildren(body2);
                children = seg.Name.ToLowerInvariant() == "p"
                    ? flat.OfType<Paragraph>()
                        .Where(p => !IsOMathParaWrapperParagraph(p))
                        .Cast<OpenXmlElement>()
                    : flat.OfType<Table>().Cast<OpenXmlElement>();
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
                    "r" => current.Descendants<Run>()
                        .Where(r => r.GetFirstChild<CommentReference>() == null)
                        .Cast<OpenXmlElement>(),
                    "tbl" => current.Elements<Table>().Cast<OpenXmlElement>(),
                    "tr" => current.Elements<TableRow>().Cast<OpenXmlElement>(),
                    "tc" => current.Elements<TableCell>().Cast<OpenXmlElement>(),
                    "sdt" => current.ChildElements
                        .Where(e => e is SdtBlock || e is SdtRun).Cast<OpenXmlElement>(),
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
                    _ => current.ChildElements.Where(e => e.LocalName == seg.Name).Cast<OpenXmlElement>()
                };
            }

            var childList = children.ToList();
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
                next = childList.OfType<Paragraph>()
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
                    _ => false,
                });
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
                    var posIdx = childList.IndexOf(next) + 1;
                    parentPath += "/" + canonName + $"[{posIdx}]";
                }
            }
            else
            {
                var posIdx = childList.IndexOf(next) + 1;
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

        // List distinct child types at this level
        var childTypes = parent.ChildElements
            .GroupBy(c => c.LocalName)
            .Select(g => $"{g.Key}({g.Count()})")
            .Take(10)
            .ToList();

        return childTypes.Count > 0
            ? $"No {requestedType} found at {parentPath}. Available children: {string.Join(", ", childTypes)}"
            : $"No children at {parentPath}";
    }

    private DocumentNode ElementToNode(OpenXmlElement element, string path, int depth)
    {
        var node = new DocumentNode { Path = path, Type = element.LocalName };

        if (element is BookmarkStart bkStart)
        {
            node.Type = "bookmark";
            node.Format["name"] = bkStart.Name?.Value ?? "";
            node.Format["id"] = bkStart.Id?.Value ?? "";
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
            var bkText = GetBookmarkText(bkStart);
            if (!string.IsNullOrEmpty(bkText))
                node.Text = bkText;
            return node;
        }

        if (element is Footnote fnEl)
        {
            node.Type = "footnote";
            // Strip the reference-mark leading space (CONSISTENCY with Query
            // get-by-id and `query footnote`). Without this branch the
            // generic InnerText fallback below would return " fn-text".
            node.Text = GetFootnoteText(fnEl);
            if (fnEl.Id?.Value != null) node.Format["id"] = fnEl.Id.Value;
            if (fnEl.Type?.Value != null) node.Format["type"] = fnEl.Type.InnerText;
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
                int fnEqIdx = 0;
                foreach (var fnEq in fnEl.Descendants<M.OfficeMath>())
                {
                    node.Children.Add(ElementToNode(fnEq, $"{path}/equation[{fnEqIdx + 1}]", depth - 1));
                    fnEqIdx++;
                }
            }
            return node;
        }

        if (element is Endnote enEl)
        {
            node.Type = "endnote";
            node.Text = GetFootnoteText(enEl);
            if (enEl.Id?.Value != null) node.Format["id"] = enEl.Id.Value;
            if (enEl.Type?.Value != null) node.Format["type"] = enEl.Type.InnerText;
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
                int enEqIdx = 0;
                foreach (var enEq in enEl.Descendants<M.OfficeMath>())
                {
                    node.Children.Add(ElementToNode(enEq, $"{path}/equation[{enEqIdx + 1}]", depth - 1));
                    enEqIdx++;
                }
            }
            return node;
        }

        if (element is Comment comment)
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
            // R21-WB-1: surface direction from the first content paragraph's
            // pPr.BiDi so the cascade (already applied by ApplyCommentFormatKeys)
            // round-trips through Get. Mirrors footnote/endnote readback above.
            var cmtBidi = comment.Descendants<Paragraph>().FirstOrDefault()?.ParagraphProperties?.GetFirstChild<BiDi>();
            if (cmtBidi != null)
                node.Format["direction"] = TryReadOnOff(cmtBidi.Val) == true ? "rtl" : "ltr";
            return node;
        }

        if (element is SectionProperties sectPrEl)
        {
            // CONSISTENCY(section-readback): /body/sectPr[N] should surface
            // the same Format keys as /section[N] so direction, page size,
            // margins, etc. are visible regardless of which path the caller
            // used. Delegate to BuildSectionNode but preserve the original
            // path the caller asked for.
            return BuildSectionNode(sectPrEl, path);
        }

        if (element is Paragraph para)
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
                        node.Format["spaceBefore"] = SpacingConverter.FormatWordSpacing(pProps.SpacingBetweenLines.Before.Value);
                    }
                    if (pProps.SpacingBetweenLines.After?.Value != null)
                    {
                        node.Format["spaceAfter"] = SpacingConverter.FormatWordSpacing(pProps.SpacingBetweenLines.After.Value);
                    }
                    if (pProps.SpacingBetweenLines.Line?.Value != null)
                    {
                        node.Format["lineSpacing"] = SpacingConverter.FormatWordLineSpacing(
                            pProps.SpacingBetweenLines.Line.Value,
                            pProps.SpacingBetweenLines.LineRule?.InnerText);
                    }
                }
                if (pProps.Indentation != null)
                {
                    var ind = pProps.Indentation;
                    // CONSISTENCY(unit-qualified-spacing): indents return "Xpt" via SpacingConverter,
                    // matching spaceBefore/spaceAfter (Canonical DocumentNode.Format Rules).
                    if (ind.FirstLine?.Value != null) node.Format["firstLineIndent"] = SpacingConverter.FormatWordSpacing(ind.FirstLine.Value);
                    if (ind.Hanging?.Value != null) node.Format["hangingIndent"] = SpacingConverter.FormatWordSpacing(ind.Hanging.Value);
                    // CONSISTENCY(ind-start-end): modern Word writes <w:ind w:start>/<w:end> instead of left/right.
                    var leftTwips = ind.Left?.Value ?? ind.Start?.Value;
                    if (leftTwips != null) node.Format["indent"] = SpacingConverter.FormatWordSpacing(leftTwips);
                    var rightTwips = ind.Right?.Value ?? ind.End?.Value;
                    if (rightTwips != null) node.Format["rightIndent"] = SpacingConverter.FormatWordSpacing(rightTwips);
                    // CONSISTENCY(ind-chars): chars-unit indents (Chinese typography) — backfilled from style Get edc8f884.
                    if (ind.FirstLineChars?.Value != null) node.Format["firstLineChars"] = ind.FirstLineChars.Value;
                    if (ind.HangingChars?.Value != null) node.Format["hangingChars"] = ind.HangingChars.Value;
                    var leftChars = ind.LeftChars?.Value ?? ind.StartCharacters?.Value;
                    if (leftChars != null) node.Format["leftChars"] = leftChars;
                    var rightChars = ind.RightChars?.Value ?? ind.EndCharacters?.Value;
                    if (rightChars != null) node.Format["rightChars"] = rightChars;
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
                    // CONSISTENCY(canonical-keys): split shading into shading.val/.fill/.color sub-keys
                    // matching the OOXML attribute structure. No compound semicolon string.
                    var shdVal = pProps.Shading.Val?.InnerText;
                    var shdFill = pProps.Shading.Fill?.Value;
                    var shdColor = pProps.Shading.Color?.Value;
                    if (!string.IsNullOrEmpty(shdVal)) node.Format["shading.val"] = shdVal;
                    if (!string.IsNullOrEmpty(shdFill)) node.Format["shading.fill"] = ParseHelpers.FormatHexColor(shdFill);
                    if (!string.IsNullOrEmpty(shdColor)) node.Format["shading.color"] = ParseHelpers.FormatHexColor(shdColor);
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
                        node.Format["numLevel"] = inhLvl.ToString();
                        var numFmt = GetNumberingFormat(inhId, inhLvl);
                        node.Format["numFmt"] = numFmt;
                        node.Format["listStyle"] = numFmt.ToLowerInvariant() == "bullet" ? "bullet" : "ordered";
                        var start = GetStartValue(inhId, inhLvl);
                        if (start != null)
                            node.Format["start"] = start.Value;
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
                    var sectMark = inlineSectPr.GetFirstChild<SectionType>()?.Val?.InnerText;
                    node.Format["sectionBreak"] = sectMark ?? "nextPage";

                    // Per-section page layout when overridden on this break.
                    var pgSz = inlineSectPr.GetFirstChild<PageSize>();
                    if (pgSz?.Width?.Value != null)
                        node.Format["sectionBreak.pageWidth"] = FormatTwipsToCm(pgSz.Width.Value);
                    if (pgSz?.Height?.Value != null)
                        node.Format["sectionBreak.pageHeight"] = FormatTwipsToCm(pgSz.Height.Value);
                    if (pgSz?.Orient?.Value != null)
                        node.Format["sectionBreak.orientation"] = pgSz.Orient.InnerText;

                    var pgMar = inlineSectPr.GetFirstChild<PageMargin>();
                    if (pgMar != null)
                    {
                        if (pgMar.Top?.Value != null)
                            node.Format["sectionBreak.marginTop"] = FormatTwipsToCm((uint)Math.Abs(pgMar.Top.Value));
                        if (pgMar.Bottom?.Value != null)
                            node.Format["sectionBreak.marginBottom"] = FormatTwipsToCm((uint)Math.Abs(pgMar.Bottom.Value));
                        if (pgMar.Left?.Value != null)
                            node.Format["sectionBreak.marginLeft"] = FormatTwipsToCm(pgMar.Left.Value);
                        if (pgMar.Right?.Value != null)
                            node.Format["sectionBreak.marginRight"] = FormatTwipsToCm(pgMar.Right.Value);
                    }

                    var pgNum = inlineSectPr.GetFirstChild<PageNumberType>();
                    if (pgNum?.Start?.Value != null)
                        node.Format["sectionBreak.pageStart"] = pgNum.Start.Value;
                    if (pgNum?.Format?.Value != null)
                        node.Format["sectionBreak.pageNumFmt"] = pgNum.Format.InnerText;

                    if (inlineSectPr.GetFirstChild<TitlePage>() != null)
                        node.Format["sectionBreak.titlePage"] = true;

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
                        if (sbCols.Space?.Value != null && uint.TryParse(sbCols.Space.Value, out var sbColSpaceTwips))
                            node.Format["sectionBreak.columnSpace"] = FormatTwipsToCm(sbColSpaceTwips);
                        if (sbCols.EqualWidth?.Value != null)
                            node.Format["sectionBreak.columns.equalWidth"] = sbCols.EqualWidth.Value;
                        if (sbCols.Separator?.Value == true)
                            node.Format["sectionBreak.columns.separator"] = true;
                    }

                    var sbVAlign = inlineSectPr.GetFirstChild<VerticalTextAlignmentOnPage>();
                    if (sbVAlign?.Val != null)
                        node.Format["sectionBreak.vAlign"] = sbVAlign.Val.InnerText;

                    var lnNum = inlineSectPr.GetFirstChild<LineNumberType>();
                    if (lnNum != null)
                    {
                        node.Format["sectionBreak.lineNumbers"] = lnNum.Restart?.InnerText switch
                        {
                            "newPage" => "restartPage",
                            "newSection" => "restartSection",
                            _ => "continuous"
                        };
                        if (lnNum.CountBy?.Value is short cb && cb > 1)
                            node.Format["sectionBreak.lineNumberCountBy"] = cb;
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
            if (pmrpForDump != null)
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
                var clr = pmrpForDump.GetFirstChild<Color>();
                if (clr != null)
                {
                    if (clr.ThemeColor?.HasValue == true)
                        node.Format["markRPr.color"] = clr.ThemeColor.InnerText;
                    else if (clr.Val?.Value != null)
                        node.Format["markRPr.color"] = ParseHelpers.FormatHexColor(clr.Val.Value);
                }
                var rf = pmrpForDump.GetFirstChild<RunFonts>();
                if (rf?.Ascii?.Value != null)
                    node.Format["markRPr.font.latin"] = rf.Ascii.Value;
                if (rf?.EastAsia?.Value != null)
                    node.Format["markRPr.font.ea"] = rf.EastAsia.Value;
                if (rf?.ComplexScript?.Value != null)
                    node.Format["markRPr.font.cs"] = rf.ComplexScript.Value;
                var hl = pmrpForDump.GetFirstChild<Highlight>();
                if (hl?.Val?.HasValue == true) node.Format["markRPr.highlight"] = hl.Val.InnerText;
            }

            // First-run formatting on the paragraph node (like PPTX does for shapes).
            // Fall back to ParagraphMarkRunProperties when no runs exist (e.g. empty paragraph
            // that had formatting applied via Set before any text was added).
            var firstRun = para.Elements<Run>().FirstOrDefault(r => r.GetFirstChild<Text>() != null);
            var paraRp = firstRun?.RunProperties
                ?? (firstRun == null ? para.ParagraphProperties?.ParagraphMarkRunProperties as OpenXmlCompositeElement : null);
            if (paraRp != null)
            {
                RunProperties? rp = paraRp as RunProperties ?? null;
                ParagraphMarkRunProperties? markRp = paraRp as ParagraphMarkRunProperties ?? null;

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
                    else if (ascii != null)
                    {
                        if (!node.Format.ContainsKey("font.latin"))
                            node.Format["font.latin"] = ascii;
                    }
                    else if (hAnsi != null)
                    {
                        if (!node.Format.ContainsKey("font.latin"))
                            node.Format["font.latin"] = hAnsi;
                    }
                    if (!string.IsNullOrEmpty(pRunFonts.EastAsia?.Value) && !node.Format.ContainsKey("font.ea"))
                        node.Format["font.ea"] = pRunFonts.EastAsia!.Value!;
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
                    // Prefer theme color over Val when both set (Val often
                    // "auto" when ThemeColor is the authoritative source).
                    if (colorEl.ThemeColor?.HasValue == true)
                        node.Format["color"] = colorEl.ThemeColor.InnerText;
                    else if (colorEl.Val?.Value != null)
                        node.Format["color"] = ParseHelpers.FormatHexColor(colorEl.Val.Value);
                }

                var ulEl = rp?.Underline ?? markRp?.GetFirstChild<Underline>();
                if (ulEl?.Val != null && !node.Format.ContainsKey("underline"))
                    node.Format["underline"] = ulEl.Val.InnerText;
                // CONSISTENCY(underline-color): backfilled from style Get edc8f884.
                if (ulEl?.Color?.Value != null && !node.Format.ContainsKey("underline.color"))
                    node.Format["underline.color"] = ParseHelpers.FormatHexColor(ulEl.Color.Value);

                var strikeEl = rp?.Strike ?? (OpenXmlLeafElement?)markRp?.GetFirstChild<Strike>();
                if (strikeEl != null && !node.Format.ContainsKey("strike")) node.Format["strike"] = true;

                var hlEl = rp?.Highlight ?? markRp?.GetFirstChild<Highlight>();
                if (hlEl?.Val != null && !node.Format.ContainsKey("highlight"))
                    node.Format["highlight"] = hlEl.Val.InnerText;
            }

            // Populate effective.* properties from style inheritance
            PopulateEffectiveParagraphProperties(node, para);

            if (depth > 0)
            {
                int runIdx = 0;
                foreach (var run in GetAllRuns(para))
                {
                    node.Children.Add(ElementToNode(run, $"{path}/r[{runIdx + 1}]", depth - 1));
                    runIdx++;
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
                foreach (var unkRun in para.Descendants<DocumentFormat.OpenXml.OpenXmlUnknownElement>())
                {
                    if (unkRun.LocalName != "r" || unkRun.NamespaceUri != wNs) continue;
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
                    if (sbInner.Length == 0) continue;
                    var synthNode = new DocumentNode
                    {
                        Type = "run",
                        Text = sbInner.ToString(),
                        Path = $"{path}/r[{runIdx + 1}]",
                    };
                    // BUG-DUMP7-10: preserve trackChange attribution from
                    // the typed w:ins/w:del ancestor so the round-trip
                    // re-emits the wrapper (mirrors the typed-Run branch
                    // at the top of this method).
                    var insAnc = unkRun.Ancestors<InsertedRun>().FirstOrDefault();
                    if (insAnc != null)
                    {
                        synthNode.Format["trackChange"] = "ins";
                        if (!string.IsNullOrEmpty(insAnc.Author?.Value))
                            synthNode.Format["trackChange.author"] = insAnc.Author!.Value!;
                        if (insAnc.Date?.Value is DateTime insAncDate)
                            synthNode.Format["trackChange.date"] = insAncDate.ToString("o");
                    }
                    else
                    {
                        var delAnc = unkRun.Ancestors<DeletedRun>().FirstOrDefault();
                        if (delAnc != null)
                        {
                            synthNode.Format["trackChange"] = "del";
                            if (!string.IsNullOrEmpty(delAnc.Author?.Value))
                                synthNode.Format["trackChange.author"] = delAnc.Author!.Value!;
                            if (delAnc.Date?.Value is DateTime delAncDate)
                                synthNode.Format["trackChange.date"] = delAncDate.ToString("o");
                        }
                    }
                    node.Children.Add(synthNode);
                    runIdx++;
                }
                // CONSISTENCY(bookmark-roundtrip): surface BookmarkStart
                // children so BatchEmitter can re-emit `add bookmark` rows.
                // Without this, dump silently dropped every bookmark on a
                // paragraph, breaking REF/HYPERLINK targets.
                int bmIdx = 0;
                foreach (var bm in para.Elements<BookmarkStart>())
                {
                    node.Children.Add(ElementToNode(bm, $"{path}/bookmark[{bmIdx + 1}]", depth - 1));
                    bmIdx++;
                }
                // BUG-DUMP4-06: surface inline SdtRun (content control) children
                // so BatchEmitter can re-emit a typed `add sdt` row carrying
                // alias/tag/type metadata. Without this, GetAllRuns unwrapped
                // the SdtRun's inner Run as a plain `add r` and the metadata
                // was silently dropped on dump round-trip.
                int sdtRunIdx = 0;
                foreach (var sdtR in para.Elements<SdtRun>())
                {
                    node.Children.Add(ElementToNode(sdtR, $"{path}/sdt[{sdtRunIdx + 1}]", depth - 1));
                    sdtRunIdx++;
                }
                // BUG-DUMP7-03: surface inline <m:oMath> children as typed
                // `equation` nodes (mode=inline) so BatchEmitter can re-emit
                // `add equation mode=inline`. Without this, GetAllRuns
                // ignored the OfficeMath element entirely and the paragraph
                // surfaced empty — dump emitted just `add p` and the formula
                // was silently dropped. Block-level oMathPara is already
                // handled at the body level via the LocalName=="oMathPara"
                // branch below.
                // BUG-DUMP8-03: include m:oMath nested inside w:ins/w:del
                // change-track wrappers — they are paragraph grandchildren,
                // not direct children, and were silently dropped on dump.
                int inlineEqIdx = 0;
                // BUG-DUMP9-04: m:oMath nested inside w:hyperlink is a
                // grandchild of the paragraph and was silently dropped.
                var inlineEqs = para.Elements<M.OfficeMath>()
                    .Concat(para.Elements<InsertedRun>().SelectMany(ins => ins.Elements<M.OfficeMath>()))
                    .Concat(para.Elements<DeletedRun>().SelectMany(del => del.Elements<M.OfficeMath>()))
                    .Concat(para.Elements<Hyperlink>().SelectMany(hl => hl.Elements<M.OfficeMath>()));
                foreach (var inlineEq in inlineEqs)
                {
                    node.Children.Add(ElementToNode(inlineEq, $"{path}/equation[{inlineEqIdx + 1}]", depth - 1));
                    inlineEqIdx++;
                }
                // BUG-DUMP6-01: surface <w:fldSimple> children as typed `field`
                // nodes so BatchEmitter can re-emit `add field` with the
                // instruction preserved. Without this, GetAllRuns descended
                // into SimpleField and surfaced the inner display run as a
                // plain run, silently dropping the w:instr attribute.
                // BUG-DUMP9-03: w:fldSimple nested inside w:hyperlink is a
                // grandchild of the paragraph and was silently dropped.
                int fldSimpleIdx = 0;
                var fldSimples = para.Elements<SimpleField>()
                    .Concat(para.Elements<Hyperlink>().SelectMany(hl => hl.Elements<SimpleField>()));
                foreach (var fld in fldSimples)
                {
                    var instr = fld.Instruction?.Value ?? "";
                    var displayText = string.Join("",
                        fld.Descendants<Text>().Select(t => t.Text));
                    var fldNode = new DocumentNode
                    {
                        Type = "field",
                        Text = displayText,
                        Path = $"{path}/field[{fldSimpleIdx + 1}]",
                    };
                    fldNode.Format["instruction"] = instr.Trim();
                    var instrUpper = instr.Trim().Split(' ', 2)[0].ToUpperInvariant();
                    if (!string.IsNullOrEmpty(instrUpper))
                        fldNode.Format["fieldType"] = instrUpper.ToLowerInvariant();
                    node.Children.Add(fldNode);
                    fldSimpleIdx++;
                }
            }
        }
        else if (element is Run run)
        {
            node.Type = "run";
            node.Text = GetRunText(run);
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
            if (insAncestor != null)
            {
                node.Format["trackChange"] = "ins";
                if (!string.IsNullOrEmpty(insAncestor.Author?.Value))
                    node.Format["trackChange.author"] = insAncestor.Author!.Value!;
                if (insAncestor.Date?.Value is DateTime insDate)
                    node.Format["trackChange.date"] = insDate.ToString("o");
            }
            else
            {
                var delAncestor = run.Ancestors<DeletedRun>().FirstOrDefault();
                if (delAncestor != null)
                {
                    node.Format["trackChange"] = "del";
                    if (!string.IsNullOrEmpty(delAncestor.Author?.Value))
                        node.Format["trackChange.author"] = delAncestor.Author!.Value!;
                    if (delAncestor.Date?.Value is DateTime delDate)
                        node.Format["trackChange.date"] = delDate.ToString("o");
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
                    else if (ascii != null) node.Format["font.latin"] = ascii;
                    else if (hAnsi != null) node.Format["font.latin"] = hAnsi;
                }
                if (!string.IsNullOrEmpty(rFonts.EastAsia?.Value)) node.Format["font.ea"] = rFonts.EastAsia!.Value!;
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
            if (run.RunProperties?.Color?.ThemeColor?.HasValue == true) node.Format["color"] = run.RunProperties.Color.ThemeColor.InnerText;
            else if (run.RunProperties?.Color?.Val?.Value != null) node.Format["color"] = ParseHelpers.FormatHexColor(run.RunProperties.Color.Val.Value);
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
            if (run.RunProperties?.Spacing?.Val?.HasValue == true)
                node.Format["charSpacing"] = $"{run.RunProperties.Spacing.Val.Value / 20.0:0.##}pt";
            if (run.RunProperties?.Shading?.Fill?.Value != null)
            {
                node.Format["shading"] = ParseHelpers.FormatHexColor(run.RunProperties.Shading.Fill.Value);
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
                if (eal.Vertical?.Value == true) node.Format["eastAsianLayout.vert"] = "1";
                if (eal.Combine?.Value == true) node.Format["eastAsianLayout.combine"] = "1";
                if (eal.VerticalCompress?.HasValue == true)
                    node.Format["eastAsianLayout.vertCompress"] = eal.VerticalCompress.InnerText;
                if (eal.CombineBrackets?.HasValue == true)
                    node.Format["eastAsianLayout.combineBrackets"] = eal.CombineBrackets.InnerText;
            }
            // Long-tail fallback: surface every rPr child the curated reader
            // didn't consume. Symmetric with the Set-side TryCreateTypedChild
            // fallback in SetElementRun (WordHandler.Set.Element.cs).
            FillUnknownChildProps(run.RunProperties, node);
            // Image properties if run contains a Drawing.
            // BUG-R5-T3: previously this branch wrote only id/name/alt/width/
            // height/relId — wrap/hPosition/vPosition/hRelative/vRelative/
            // behindText for floating pictures were silently dropped, which
            // also broke dump→batch round-trip (BatchEmitter relies on Get).
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
                    node.Type = "ptab";
                    // Open XML SDK v3 enum .ToString() returns "FooValues { }"
                    // — use .InnerText to get the actual XML attribute value
                    // ("center", "right", "begin", etc.). Same trap as the
                    // LineSpacingRuleValues note in WordHandler CLAUDE.md.
                    if (ptabEl.Alignment?.HasValue == true)
                        node.Format["align"] = ptabEl.Alignment.InnerText;
                    if (ptabEl.RelativeTo?.HasValue == true)
                        node.Format["relativeTo"] = ptabEl.RelativeTo.InnerText;
                    if (ptabEl.Leader?.HasValue == true)
                        node.Format["leader"] = ptabEl.Leader.InnerText;
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
                    if (fldCharEl.FormFieldData != null)
                        node.Format["hasFormFieldData"] = true;
                }
            }
            if (node.Type == "run")
            {
                var instrEl = run.GetFirstChild<FieldCode>();
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
            }
            // CONSISTENCY(run-text-tab): the type-upgrade for tab/break runs
            // checks "no Text element" (not "node.Text empty") because
            // GetRunText now surfaces TabChar as \t in node.Text. A pure
            // <w:r><w:tab/></w:r> run has no <w:t> child but node.Text="\t".
            if (node.Type == "run" && !run.Elements<Text>().Any())
            {
                var tabEl = run.GetFirstChild<TabChar>();
                if (tabEl != null)
                {
                    node.Type = "tab";
                    node.Text = "";
                }
            }
            if (node.Type == "run" && string.IsNullOrEmpty(node.Text))
            {
                var breakEl = run.GetFirstChild<Break>();
                if (breakEl != null)
                {
                    node.Type = "break";
                    if (breakEl.Type?.HasValue == true)
                        node.Format["breakType"] = breakEl.Type.InnerText;
                }
            }

            if (run.Parent is Hyperlink hlParent)
            {
                // BUG-DUMP10-05: a hyperlink wrapper with neither r:id nor
                // anchor (tooltip-only / history-only) used to fall through
                // both branches below, leaving the run with no Format keys
                // that would trigger the BatchEmitter hyperlink-emit guard.
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
                        if (rel != null) node.Format["url"] = rel.Uri.ToString();
                    }
                    catch { }
                }
                // CONSISTENCY(internal-anchor-hyperlink): runs inside an
                // internal anchor hyperlink (w:hyperlink[@w:anchor]) had no
                // r:id, so `anchor` was never surfaced on the run. The
                // BatchEmitter hyperlink branch keys off Format["anchor"]/
                // ["url"] to emit `add hyperlink`; without anchor the run
                // was demoted to a plain `add r` and the link was lost on
                // dump→batch round-trip.
                if (hlParent.Anchor?.Value != null)
                    node.Format["anchor"] = hlParent.Anchor.Value;
                // BUG-DUMP10-02: surface the tooltip / tgtFrame / history
                // attributes from the wrapping hyperlink so dump→batch
                // round-trip preserves them. Same canonical keys as the
                // standalone Hyperlink branch below.
                if (hlParent.Tooltip?.Value != null)
                    node.Format["tooltip"] = hlParent.Tooltip.Value;
                if (hlParent.TargetFrame?.Value != null)
                    node.Format["tgtFrame"] = hlParent.TargetFrame.Value;
                if (hlParent.History?.Value == true)
                    node.Format["history"] = true;
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
            // the type upgrade above (font.*/size/bold/...): they are valid
            // OOXML but irrelevant to special-content runs, where node.Type
            // already conveys the semantic role. Strip them for ptab /
            // fieldChar / instrText / tab / break so audit tools see a
            // clean property bag (alignment, fieldCharType, instr,
            // breakType, etc.).
            if (node.Type is "ptab" or "fieldChar" or "instrText" or "tab" or "break")
            {
                foreach (var noiseKey in TypographyOnlyKeys)
                    node.Format.Remove(noiseKey);
            }
        }
        else if (element is Hyperlink hyperlink)
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
                    if (rel != null) node.Format["url"] = rel.Uri.ToString();
                }
                catch { }
            }
            // Internal-anchor hyperlink (`add --type hyperlink --prop anchor=Foo`)
            // sets w:hyperlink/@w:anchor instead of @r:id. Surface it so set/get
            // round-trips and users can debug why a link points where it does.
            if (hyperlink.Anchor?.Value != null)
                node.Format["anchor"] = hyperlink.Anchor.Value;
            // BUG-DUMP10-02: tooltip / tgtFrame / history attributes are
            // independent of url/anchor — surface them so dump→batch
            // preserves the hover popup, target window, and history flag.
            if (hyperlink.Tooltip?.Value != null)
                node.Format["tooltip"] = hyperlink.Tooltip.Value;
            if (hyperlink.TargetFrame?.Value != null)
                node.Format["tgtFrame"] = hyperlink.TargetFrame.Value;
            if (hyperlink.History?.Value == true)
                node.Format["history"] = true;
            // Read run formatting from the first run inside the hyperlink
            var hlRun = hyperlink.Elements<Run>().FirstOrDefault(r => r.GetFirstChild<Text>() != null);
            if (hlRun?.RunProperties != null)
            {
                var rp = hlRun.RunProperties;
                if (rp.RunFonts?.Ascii?.Value != null) node.Format["font"] = rp.RunFonts.Ascii.Value;
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
            }
        }
        else if (element is Table table)
        {
            node.Type = "table";
            node.ChildCount = table.Elements<TableRow>().Count();
            var firstRow = table.Elements<TableRow>().FirstOrDefault();
            // Use grid column count (from TableGrid) instead of cell count for accurate column reporting
            var gridColCount = table.GetFirstChild<TableGrid>()?.Elements<GridColumn>().Count();
            // CONSISTENCY(format-stringy): user-facing numeric counts are
            // stored as strings to match other Word format keys (size "14pt",
            // spacing "12pt"). Avoids object-vs-int comparison surprises.
            node.Format["cols"] = (gridColCount ?? firstRow?.Elements<TableCell>().Count() ?? 0).ToString();
            node.Format["rows"] = node.ChildCount.ToString();

            var tp = table.GetFirstChild<TableProperties>();
            if (tp != null)
            {
                // Table style
                if (tp.TableStyle?.Val?.Value != null)
                    node.Format["style"] = tp.TableStyle.Val.Value;
                // Table borders
                var tblBorders = tp.TableBorders;
                if (tblBorders != null)
                {
                    ReadBorder(tblBorders.TopBorder, "border.top", node);
                    ReadBorder(tblBorders.BottomBorder, "border.bottom", node);
                    ReadBorder(tblBorders.LeftBorder, "border.left", node);
                    ReadBorder(tblBorders.RightBorder, "border.right", node);
                    ReadBorder(tblBorders.InsideHorizontalBorder, "border.insideH", node);
                    ReadBorder(tblBorders.InsideVerticalBorder, "border.insideV", node);
                }
                // Table width
                if (tp.TableWidth?.Width?.Value != null)
                {
                    var wType = tp.TableWidth.Type?.Value;
                    node.Format["width"] = wType == TableWidthUnitValues.Pct
                        ? (int.Parse(tp.TableWidth.Width.Value) / 50) + "%"
                        : tp.TableWidth.Width.Value;
                }
                // Alignment
                if (tp.TableJustification?.Val?.Value != null)
                    node.Format["align"] = tp.TableJustification.Val.InnerText;
                // Indent
                if (tp.TableIndentation?.Width?.Value != null)
                    node.Format["indent"] = tp.TableIndentation.Width.Value;
                // Cell spacing
                if (tp.TableCellSpacing?.Width?.Value != null)
                    node.Format["cellSpacing"] = tp.TableCellSpacing.Width.Value;
                // Layout
                if (tp.TableLayout?.Type?.Value != null)
                    node.Format["layout"] = tp.TableLayout.Type.Value == TableLayoutValues.Fixed ? "fixed" : "auto";
                // Direction (CT_TblPrBase / w:bidiVisual). Mirrors paragraph
                // direction vocabulary; presence-only readback (no bidiVisual
                // means no key — LTR is the default).
                if (tp.GetFirstChild<BiDiVisual>() != null)
                    node.Format["direction"] = "rtl";
                // Default cell margin (padding)
                var dcm = tp.TableCellMarginDefault;
                if (dcm?.TopMargin?.Width?.Value != null)
                    node.Format["padding.top"] = dcm.TopMargin.Width.Value;
                if (dcm?.BottomMargin?.Width?.Value != null)
                    node.Format["padding.bottom"] = dcm.BottomMargin.Width.Value;
                if (dcm?.TableCellLeftMargin?.Width?.Value != null)
                    node.Format["padding.left"] = dcm.TableCellLeftMargin.Width.Value;
                if (dcm?.TableCellRightMargin?.Width?.Value != null)
                    node.Format["padding.right"] = dcm.TableCellRightMargin.Width.Value;
            }

            // Column widths from grid
            var gridCols = table.GetFirstChild<TableGrid>()?.Elements<GridColumn>().ToList();
            if (gridCols != null && gridCols.Count > 0)
                node.Format["colWidths"] = string.Join(",", gridCols.Select(g => g.Width?.Value ?? "0"));

            if (depth > 0)
            {
                int rowIdx = 0;
                foreach (var row in table.Elements<TableRow>())
                {
                    var rowNode = new DocumentNode
                    {
                        Path = $"{path}/tr[{rowIdx + 1}]",
                        Type = "row",
                        ChildCount = row.Elements<TableCell>().Count()
                    };
                    ReadRowProps(row, rowNode);
                    if (depth > 1)
                    {
                        int cellIdx = 0;
                        foreach (var cell in row.Elements<TableCell>())
                        {
                            var cellNode = new DocumentNode
                            {
                                Path = $"{path}/tr[{rowIdx + 1}]/tc[{cellIdx + 1}]",
                                Type = "cell",
                                Text = string.Join("", cell.Descendants<Text>().Select(t => t.Text)),
                                // CONSISTENCY(cell-children): include nested Table children alongside Paragraphs.
                                ChildCount = cell.Elements<OpenXmlElement>().Count(e => e is Paragraph || e is Table)
                            };
                            ReadCellProps(cell, cellNode);
                            if (depth > 2)
                            {
                                int cellPIdx = 0, cellTblIdx = 0;
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
        }
        else if (element is TableCell directCell)
        {
            node.Type = "cell";
            node.Text = string.Join("", directCell.Descendants<Text>().Select(t => t.Text));
            // CONSISTENCY(cell-children): include nested Table children alongside Paragraphs.
            node.ChildCount = directCell.Elements<OpenXmlElement>().Count(e => e is Paragraph || e is Table);
            ReadCellProps(directCell, node);
            if (depth > 0)
            {
                int dcPIdx = 0, dcTblIdx = 0;
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
                }
            }
        }
        else if (element is TableRow directRow)
        {
            node.Type = "row";
            node.ChildCount = directRow.Elements<TableCell>().Count();
            ReadRowProps(directRow, node);
            if (depth > 0)
            {
                int cellIdx = 0;
                foreach (var cell in directRow.Elements<TableCell>())
                {
                    var cellNode = new DocumentNode
                    {
                        Path = $"{path}/tc[{cellIdx + 1}]",
                        Type = "cell",
                        Text = string.Join("", cell.Descendants<Text>().Select(t => t.Text)),
                        // CONSISTENCY(cell-children): include nested Table children alongside Paragraphs.
                        ChildCount = cell.Elements<OpenXmlElement>().Count(e => e is Paragraph || e is Table)
                    };
                    ReadCellProps(cell, cellNode);
                    if (depth > 1)
                    {
                        int drPIdx = 0, drTblIdx = 0;
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
                        }
                    }
                    node.Children.Add(cellNode);
                    cellIdx++;
                }
            }
        }
        else if (element is SdtBlock sdtBlockNode)
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
                if (sdtProps.GetFirstChild<SdtContentDropDownList>() != null) node.Format["type"] = "dropdown";
                else if (sdtProps.GetFirstChild<SdtContentComboBox>() != null) node.Format["type"] = "combobox";
                else if (sdtProps.GetFirstChild<SdtContentDate>() != null) node.Format["type"] = "date";
                else if (sdtProps.GetFirstChild<SdtContentText>() != null) node.Format["type"] = "text";
                else node.Format["type"] = "richtext";

                // Read date format for date controls
                var dateContent = sdtProps.GetFirstChild<SdtContentDate>();
                if (dateContent?.DateFormat?.Val?.Value != null)
                    node.Format["format"] = dateContent.DateFormat.Val.Value;

                // Editable status
                node.Format["editable"] = IsSdtEditable(sdtProps);

                // Placeholder detection
                var showingPlcHdr = sdtProps.GetFirstChild<ShowingPlaceholder>();
                if (showingPlcHdr != null)
                {
                    node.Format["placeholder"] = true;
                    var plcHdrText = sdtProps.GetFirstChild<SdtPlaceholder>()?.DocPartReference?.Val?.Value;
                    if (plcHdrText != null) node.Format["placeholderText"] = plcHdrText;
                }

                // Read dropdown/combobox items
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
            }
            node.Text = string.Concat(sdtBlockNode.Descendants<Text>().Select(t => t.Text));
            var sdtContent = sdtBlockNode.SdtContentBlock;
            node.ChildCount = sdtContent?.ChildElements.Count ?? 0;
        }
        else if (element is SdtRun sdtRunNode)
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

                if (sdtProps.GetFirstChild<SdtContentDropDownList>() != null) node.Format["type"] = "dropdown";
                else if (sdtProps.GetFirstChild<SdtContentComboBox>() != null) node.Format["type"] = "combobox";
                else if (sdtProps.GetFirstChild<SdtContentDate>() != null) node.Format["type"] = "date";
                else if (sdtProps.GetFirstChild<SdtContentText>() != null) node.Format["type"] = "text";
                else node.Format["type"] = "richtext";

                // Editable status
                node.Format["editable"] = IsSdtEditable(sdtProps);

                // Placeholder detection
                var showingPlcHdrRun = sdtProps.GetFirstChild<ShowingPlaceholder>();
                if (showingPlcHdrRun != null)
                {
                    node.Format["placeholder"] = true;
                    var plcHdrTextRun = sdtProps.GetFirstChild<SdtPlaceholder>()?.DocPartReference?.Val?.Value;
                    if (plcHdrTextRun != null) node.Format["placeholderText"] = plcHdrTextRun;
                }

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
            }
            node.Text = string.Concat(sdtRunNode.Descendants<Text>().Select(t => t.Text));
        }
        else if (element.LocalName == "oMathPara" || element is M.Paragraph)
        {
            node.Type = "equation";
            node.Format["mode"] = "display";
            // Extract LaTeX via FormulaParser
            var oMath = element.Descendants<M.OfficeMath>().FirstOrDefault();
            if (oMath != null)
            {
                try { node.Text = Core.FormulaParser.ToLatex(oMath); }
                catch { node.Text = element.InnerText; }
            }
            else
            {
                node.Text = element.InnerText;
            }
        }
        else if (element is M.OfficeMath inlineMath)
        {
            node.Type = "equation";
            node.Format["mode"] = "inline";
            try { node.Text = Core.FormulaParser.ToLatex(inlineMath); }
            catch { node.Text = element.InnerText; }
        }
        else if (element is Header or Footer)
        {
            // Header/Footer: enumerate block-level children. Tables are valid
            // block-level OOXML inside hdr/ftr (same schema as body), so list
            // them alongside paragraphs. Mirrors body-listing logic above.
            node.Type = element is Header ? "header" : "footer";
            node.Text = string.Concat(element.Descendants<Text>().Select(t => t.Text));
            node.ChildCount = element.Elements<Paragraph>().Count() + element.Elements<Table>().Count();
            if (depth > 0)
            {
                int pIdx = 0, tblIdx = 0;
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
                }
            }
        }
        else if (element is Body bodyNode)
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
                    else
                    {
                        // Non-structural (sectPr etc.) — keep localName naming
                        node.Children.Add(ElementToNode(child, $"{path}/{child.LocalName}[1]", depth - 1));
                    }
                }
                foreach (var child in bodyNode.ChildElements)
                    WalkBodyChild(child);
            }
        }
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
        var trPr = row.TableRowProperties;
        if (trPr == null) return;
        var rh = trPr.GetFirstChild<TableRowHeight>();
        if (rh?.Val?.Value != null)
        {
            node.Format["height"] = rh.Val.Value;
            if (rh.HeightType?.Value == HeightRuleValues.Exact)
                node.Format["height.rule"] = "exact";
        }
        if (trPr.GetFirstChild<TableHeader>() != null)
            node.Format["header"] = true;
        if (trPr.GetFirstChild<CantSplit>() != null)
            node.Format["cantSplit"] = true;
    }

    private static void ReadCellProps(TableCell cell, DocumentNode node)
    {
        var tcPr = cell.TableCellProperties;
        if (tcPr != null)
        {
            // Borders (including diagonal — like POI CTTcBorders)
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
                if (shd?.Fill?.Value != null)
                {
                    node.Format["fill"] = ParseHelpers.FormatHexColor(shd.Fill.Value);
                }
            }
            // Width
            // BUG-DUMP6-04: preserve w:tcW @type semantics. Mirror the table-level
            // width readback above (line ~1930) — pct widths are stored as
            // fifths-of-percent, so divide by 50 and append '%' so dump→batch
            // can recognize and re-emit pct cell widths.
            if (tcPr.TableCellWidth?.Width?.Value != null)
            {
                var cwType = tcPr.TableCellWidth.Type?.Value;
                if (cwType == TableWidthUnitValues.Pct
                    && int.TryParse(tcPr.TableCellWidth.Width.Value, out var pctRaw))
                    node.Format["width"] = (pctRaw / 50) + "%";
                else
                    node.Format["width"] = tcPr.TableCellWidth.Width.Value;
            }
            // Vertical alignment
            if (tcPr.TableCellVerticalAlignment?.Val?.Value != null)
                node.Format["valign"] = tcPr.TableCellVerticalAlignment.Val.InnerText;
            // Vertical merge
            if (tcPr.VerticalMerge != null)
                node.Format["vmerge"] = tcPr.VerticalMerge.Val?.Value == MergedCellValues.Restart ? "restart" : "continue";
            // Grid span
            if (tcPr.GridSpan?.Val?.Value != null && tcPr.GridSpan.Val.Value > 1)
                node.Format["colspan"] = tcPr.GridSpan.Val.Value;
            // Cell padding/margins
            var mar = tcPr.TableCellMargin;
            if (mar != null)
            {
                if (mar.TopMargin?.Width?.Value != null) node.Format["padding.top"] = mar.TopMargin.Width.Value;
                if (mar.BottomMargin?.Width?.Value != null) node.Format["padding.bottom"] = mar.BottomMargin.Width.Value;
                if (mar.LeftMargin?.Width?.Value != null) node.Format["padding.left"] = mar.LeftMargin.Width.Value;
                if (mar.RightMargin?.Width?.Value != null) node.Format["padding.right"] = mar.RightMargin.Width.Value;
            }
            // Text direction
            if (tcPr.TextDirection?.Val?.Value != null)
                node.Format["textDirection"] = tcPr.TextDirection.Val.InnerText;
            // No wrap
            if (tcPr.NoWrap != null)
                node.Format["nowrap"] = true;
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
            if (rPr.Strike != null) node.Format["strike"] = true;
            if (rPr.Highlight?.Val != null) node.Format["highlight"] = rPr.Highlight.Val.InnerText;
        }
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
    // curated reader. Shape is symmetric with GenericXmlQuery.TryCreateTypedChild
    // on the Set side: child-with-val → Format[name]=val; toggle (no attrs) →
    // Format[name]=true. Multi-attribute / nested children are skipped — the
    // generic Set path can't write them, so exposing them would produce keys
    // that don't round-trip.
    private static void FillUnknownChildProps(OpenXmlElement? container, DocumentNode node)
    {
        if (container == null) return;
        foreach (var child in container.ChildElements)
        {
            var name = child.LocalName;
            if (string.IsNullOrEmpty(name)) continue;
            if (CuratedStyleLocalNames.Contains(name)) continue;
            if (node.Format.ContainsKey(name)) continue;
            if (child.ChildElements.Count > 0) continue;

            string? valAttr = null;
            int attrCount = 0;
            foreach (var a in child.GetAttributes())
            {
                attrCount++;
                if (a.LocalName.Equals("val", System.StringComparison.OrdinalIgnoreCase))
                    valAttr = a.Value;
            }
            if (valAttr != null)
                node.Format[name] = valAttr;
            else if (attrCount == 0)
                node.Format[name] = true;
            // else: complex multi-attribute element — skip, curated reader
            // is expected to cover it (e.g. rFonts is in CuratedStyleLocalNames).
        }
    }
}
