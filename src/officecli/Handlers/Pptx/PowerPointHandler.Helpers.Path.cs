// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler
{

    /// <summary>
    /// Normalize cell[R,C] shorthand to tr[R]/tc[C] in paths.
    /// E.g. /slide[1]/table[1]/cell[2,3] → /slide[1]/table[1]/tr[2]/tc[3]
    /// Also handles trailing segments: /slide[1]/table[1]/cell[2,3]/txBody → /slide[1]/table[1]/tr[2]/tc[3]/txBody
    /// </summary>
    /// <summary>
    /// CONSISTENCY(path-stability): the per-handler path-pattern regexes are mostly
    /// case-sensitive. DOCX folds case via ToLowerInvariant on every segment name
    /// (Navigation.cs); we mirror that here by lowercasing the alphabetic LocalName
    /// portion of every `<name>[index]` segment so `/SLIDE[1]/SHAPE[2]` is treated
    /// identically to `/slide[1]/shape[2]` and routes through the structured matchers
    /// instead of falling through to the raw-XML default.
    /// </summary>
    private static string NormalizePptxPathSegmentCasing(string path)
    {
        if (string.IsNullOrEmpty(path) || path == "/") return path;
        // Lowercase only the LocalName before '[' or '/' or end-of-segment. Preserve
        // bracketed identifiers (placeholder[Title 1]), attribute selectors (@role=ROLE),
        // and named arguments verbatim — only the leading element-name token is folded.
        return Regex.Replace(path, @"(?<=^|/)([A-Za-z][A-Za-z0-9]*)",
            m => m.Value.ToLowerInvariant());
    }

    /// <summary>
    /// Resolve a content-add parent path that is either a slide (/slide[N]) or a
    /// (possibly nested) group inside a slide (/slide[N]/group[K]/group[L]/…).
    /// Returns the owning slide part, the slide's shape tree (for shape-id
    /// allocation and from/to lookups), the container the new element should be
    /// inserted into (the group when one is addressed, else the shape tree), the
    /// 1-based slide index, and the canonical return-path prefix
    /// (e.g. "/slide[2]/group[1]"). Returns null when the path is neither a slide
    /// nor a slide-group path, so callers keep their own slide-only error message.
    ///
    /// CONSISTENCY(group-inner-shape-add): mirrors the nested-group resolution in
    /// AddShape so picture/connector/media adds accept the same group parents that
    /// dump emits for grouped content.
    /// </summary>
    private (SlidePart slidePart, ShapeTree shapeTree, OpenXmlCompositeElement insertContainer,
             int slideIdx, string returnPathPrefix)?
        ResolveSlideOrGroupAddParent(string parentPath)
    {
        var groupMatch = Regex.Match(parentPath, @"^/slide\[(\d+)\]((?:/group\[\d+\])*)$");
        if (!groupMatch.Success) return null;

        var slideIdx = int.Parse(groupMatch.Groups[1].Value);
        var slideParts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > slideParts.Count)
            throw new ArgumentException($"Slide {slideIdx} not found (total: {slideParts.Count})");
        var slidePart = slideParts[PathIndex.ToArrayIndex(slideIdx)];
        var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree
            ?? throw new InvalidOperationException("Slide has no shape tree");

        OpenXmlCompositeElement insertContainer = shapeTree;
        var groupSegs = groupMatch.Groups[2].Value;
        if (groupSegs.Length > 0)
        {
            foreach (Match seg in Regex.Matches(groupSegs, @"/group\[(\d+)\]"))
            {
                var grpIdx = int.Parse(seg.Groups[1].Value);
                var groups = insertContainer.Elements<GroupShape>().ToList();
                if (grpIdx < 1 || grpIdx > groups.Count)
                    throw new ArgumentException(
                        $"Group {grpIdx} not found under {insertContainer.LocalName} on slide {slideIdx} (total: {groups.Count})");
                insertContainer = groups[grpIdx - 1];
            }
        }
        var returnPathPrefix = $"/slide[{slideIdx}]{groupSegs}";
        return (slidePart, shapeTree, insertContainer, slideIdx, returnPathPrefix);
    }

    private static string NormalizeCellPath(string path)
    {
        // Reject malformed segment separators that previously slipped past
        // the regex matchers and ended up exposing raw OOXML local names
        // (e.g. `Get("/slide[1]/")` returned type=sld, `Get("//slide[1]")`
        // returned sld). DOCX already rejects these forms; bring PPTX/XLSX
        // up to parity with an explicit error rather than silent leakage.
        if (path.Length > 1 && path != "/" && path.EndsWith("/"))
            throw new ArgumentException($"Invalid path '{path}': trailing '/' is not allowed.");
        if (path.StartsWith("//"))
            throw new ArgumentException($"Invalid path '{path}': leading '//' is not allowed.");
        if (path.Contains("//"))
            throw new ArgumentException($"Invalid path '{path}': empty path segment ('//') is not allowed.");
        // CONSISTENCY(table-path-long-form): pptx the project conventions documents long form
        // /slide[N]/table[K]/row[R]/cell[C] as canonical. Query/Add already alias
        // row→tr and cell→tc at their dispatch layer; mirror that here so Get/Set
        // /Remove parse paths also accept long form. Short OOXML form (tr/tc)
        // continues to work unchanged.
        path = Regex.Replace(path, @"cell\[(\d+),\s*(\d+)\]", m => $"tr[{m.Groups[1].Value}]/tc[{m.Groups[2].Value}]");
        // Alias only inside /table[K]/... — never globally, to avoid colliding
        // with hypothetical future top-level "row"/"cell" segments.
        // BUG-004: the table segment must also match selector forms
        // (table[@name=Foo], table[@id=N]) — @name paths previously skipped
        // row/cell aliasing entirely, so Set on /table[@name=X]/row/cell fell
        // through to an element that rejects text (UNSUPPORTED props: text).
        path = Regex.Replace(path, @"(/table\[[^\]]+\](?:/[^/]+)*?)/row\[(\d+)\]", m => $"{m.Groups[1].Value}/tr[{m.Groups[2].Value}]");
        path = Regex.Replace(path, @"(/tr\[\d+\])/cell\[(\d+)\]", m => $"{m.Groups[1].Value}/tc[{m.Groups[2].Value}]");
        // CONSISTENCY(table-path-long-form): same parity for the column axis.
        // schemas/help/pptx/table-column.json declares element=column with
        // alias col, and Add accepts --type column. Get/Set/Remove must also
        // accept the long form so all five ops share one path vocabulary.
        path = Regex.Replace(path, @"(/table\[[^\]]+\])/column\[(\d+)\]", m => $"{m.Groups[1].Value}/col[{m.Groups[2].Value}]");
        return path;
    }

    /// <summary>
    /// Resolve InsertPosition (After/Before anchor path) to a 0-based int? index for PPT.
    /// Anchor path can be full (/slide[1]/shape[@id=X]) or short (shape[@id=X]).
    /// </summary>
    /// <summary>Sentinel value for find: anchor resolution.</summary>
    private const int FindAnchorIndex = -99999;

    private int? ResolveAnchorPosition(string parentPath, InsertPosition? position)
    {
        if (position == null) return null;
        if (position.Index.HasValue) return position.Index;

        var anchorPath = position.After ?? position.Before!;

        // Catch bare attribute selector without element wrapper, e.g. @id=XXX instead of shape[@id=XXX]
        if (Regex.IsMatch(anchorPath, @"^@(\w+)=(.+)$"))
            throw new ArgumentException($"Invalid anchor path \"{anchorPath}\". Did you mean: shape[{anchorPath}]?");

        // Handle find: prefix — text-based anchoring
        if (anchorPath.StartsWith("find:", StringComparison.OrdinalIgnoreCase))
            return FindAnchorIndex;

        // Normalize: if short form, prepend parentPath
        if (!anchorPath.StartsWith("/"))
            anchorPath = parentPath.TrimEnd('/') + "/" + anchorPath;

        // Resolve @id=/@name= in the anchor path
        anchorPath = ResolveIdPath(anchorPath);

        // For slide-level anchors (/slide[N])
        var slideMatch = Regex.Match(anchorPath, @"^/slide\[(\d+)\]$");
        if (slideMatch.Success)
        {
            var slideIdx = int.Parse(slideMatch.Groups[1].Value) - 1; // 0-based
            var slideCount = GetSlideParts().Count();
            if (slideIdx < 0 || slideIdx >= slideCount)
                throw new ArgumentException($"Anchor slide not found: {anchorPath} (total slides: {slideCount})");
            if (position.After != null)
                return slideIdx + 1 >= slideCount ? null : slideIdx + 1;
            else
                return slideIdx;
        }

        // For element-level anchors. CONSISTENCY(pptx-group-flatten): allow
        // optional /group[K] ancestors so anchors like /slide[1]/group[2]/shape[3]
        // resolve to the position inside the group's children.
        var elemMatch = Regex.Match(anchorPath, @"^/slide\[(\d+)\]((?:/group\[\d+\])*)/(\w+)\[(\d+)\]$");
        if (elemMatch.Success)
        {
            var slideIdx = int.Parse(elemMatch.Groups[1].Value);
            var elemGroupChain = elemMatch.Groups[2].Value;
            var elemIdx = int.Parse(elemMatch.Groups[4].Value) - 1; // 0-based
            // Validate that the anchor element exists
            var slideParts = GetSlideParts().ToList();
            if (slideIdx < 1 || slideIdx > slideParts.Count)
                throw new ArgumentException($"Anchor slide not found: {anchorPath} (total slides: {slideParts.Count})");
            OpenXmlCompositeElement? anchorContainer = GetSlide(slideParts[PathIndex.ToArrayIndex(slideIdx)]).CommonSlideData?.ShapeTree;
            if (anchorContainer != null && !string.IsNullOrEmpty(elemGroupChain))
            {
                foreach (Match gm in Regex.Matches(elemGroupChain, @"/group\[(\d+)\]"))
                {
                    var gIdx = int.Parse(gm.Groups[1].Value);
                    var groupsAtScope = anchorContainer.Elements<GroupShape>().ToList();
                    if (gIdx < 1 || gIdx > groupsAtScope.Count)
                        throw new ArgumentException($"Anchor group {gIdx} not found in scope (have {groupsAtScope.Count})");
                    anchorContainer = groupsAtScope[gIdx - 1];
                }
            }
            if (anchorContainer != null)
            {
                var contentChildren = anchorContainer.ChildElements
                    .Where(e => e is not NonVisualGroupShapeProperties && e is not GroupShapeProperties)
                    .ToList();
                if (elemIdx < 0 || elemIdx >= contentChildren.Count)
                    throw new ArgumentException($"Anchor element not found: {anchorPath} (total elements in scope: {contentChildren.Count})");
            }
            if (position.After != null)
                return elemIdx + 1; // InsertAtPosition handles bounds
            else
                return elemIdx;
        }

        // Table sub-element anchors: /slide[N]/table[K]/(tr|row|col|column)[N]
        // Used by `add --type row/col --before/--after` on PPT tables. The
        // anchor's positional index is all we need — the dispatcher (AddRow /
        // AddColumn) consumes the returned index against the table's own
        // tr/gridCol list.
        var tableSubMatch = Regex.Match(anchorPath, @"^/slide\[(\d+)\]/table\[(\d+)\]/(tr|row|col|column)\[(\d+)\]$");
        if (tableSubMatch.Success)
        {
            var subIdx = int.Parse(tableSubMatch.Groups[4].Value) - 1; // 0-based
            if (position.After != null)
                return subIdx + 1;
            else
                return subIdx;
        }

        throw new ArgumentException($"Cannot resolve anchor path: {anchorPath}");
    }

    /// <summary>
    /// Resolve @id= and @name= attribute selectors in a PPT path to positional indices.
    /// E.g. /slide[1]/shape[@id=5] → /slide[1]/shape[N] where N is the positional index of shape with cNvPr.Id=5.
    /// </summary>
    private string ResolveIdPath(string path)
    {
        // Null/empty paths are a valid "duplicate in place" / "no target"
        // signal from CopyFrom and friends; pass them through untouched so
        // downstream dispatch can interpret the null itself.
        if (path == null) return path!;
        // Quick check: if no [@, nothing to resolve
        if (!path.Contains("[@"))
            return path;

        // Iterate matches left-to-right so we can rewrite the prefix as we go;
        // each successive @id=/@name= resolves relative to whatever group context
        // the earlier (already-rewritten) prefix established.
        var sb = new System.Text.StringBuilder();
        var cursor = 0;
        var rewritten = path;
        // Support quoted attr values so a name containing ']' (e.g. PowerPoint's
        // auto-generated "Shape [1] copy") survives the predicate parse: the
        // unquoted fallback stops at the first ']' as before.
        // The predicate group is repeatable so `placeholder[@type=body][@idx=1]`
        // (issue #260 — same chained-bracket form StyleList uses for
        // /master/ph[@type=..][@idx=..] style paths) parses as ONE segment.
        var matches = Regex.Matches(path, @"(\w+)((?:\[@\w+=(?:'[^']*'|""[^""]*""|[^\]]+)\])+)");
        foreach (Match m in matches)
        {
            var elementType = m.Groups[1].Value.ToLowerInvariant();

            // Parse the chained predicates: [(key, value), ...].
            var predicates = new List<(string Key, string Value)>();
            foreach (Match pm in Regex.Matches(m.Groups[2].Value,
                @"\[@(\w+)=(?:'([^']*)'|""([^""]*)""|([^\]]+))\]"))
            {
                // Three alternation captures: single-quoted (2), double-quoted
                // (3), unquoted (4). Trim is still useful for the unquoted form
                // because the schema documents @name=Foo Bar (no quotes).
                string v;
                if (pm.Groups[2].Success) v = pm.Groups[2].Value;
                else if (pm.Groups[3].Success) v = pm.Groups[3].Value;
                else v = pm.Groups[4].Value.Trim('"', '\'', ' ');
                predicates.Add((pm.Groups[1].Value.ToLowerInvariant(), v));
            }

            bool isPlaceholder = elementType == "placeholder";
            // Everything except placeholder keeps the legacy contract: exactly
            // one @id=/@name= predicate. Any other shape (StyleList's
            // `ph[@type=..][@idx=..]` style paths, raw-XML attribute paths)
            // passes through UNTOUCHED so its own parser still sees it.
            if (!isPlaceholder
                && !(predicates.Count == 1 && predicates[0].Key is "id" or "name"))
            {
                sb.Append(path, cursor, m.Index + m.Length - cursor);
                cursor = m.Index + m.Length;
                continue;
            }

            sb.Append(path, cursor, m.Index - cursor);
            var prefix = sb.ToString();

            var attrName = predicates[0].Key;
            var attrValue = predicates[0].Value;

            // CONSISTENCY(master-layout-shape-edit): @id=/@name= resolution must
            // also work when the prefix is a slidemaster or slidelayout shape
            // container — Add returns `/slidemaster[N]/shape[@id=K]` so the
            // same path must round-trip through Get/Set/Remove.
            ShapeTree? shapeTree;
            var nestedMlMatch = Regex.Match(prefix, @"^/slidemaster\[(\d+)\]/slidelayout\[(\d+)\]", RegexOptions.IgnoreCase);
            var masterMlMatch = Regex.Match(prefix, @"^/slidemaster\[(\d+)\]", RegexOptions.IgnoreCase);
            var layoutMlMatch = Regex.Match(prefix, @"^/slidelayout\[(\d+)\]", RegexOptions.IgnoreCase);
            if (nestedMlMatch.Success)
            {
                var mIdx = int.Parse(nestedMlMatch.Groups[1].Value);
                var lIdx = int.Parse(nestedMlMatch.Groups[2].Value);
                var masters = _doc.PresentationPart?.SlideMasterParts?.ToList() ?? [];
                if (mIdx < 1 || mIdx > masters.Count)
                    throw new ArgumentException($"Slide master {mIdx} not found (total: {masters.Count})");
                var layouts = masters[mIdx - 1].SlideLayoutParts?.ToList() ?? [];
                if (lIdx < 1 || lIdx > layouts.Count)
                    throw new ArgumentException($"Slide layout {lIdx} not found under master {mIdx} (total: {layouts.Count})");
                shapeTree = layouts[lIdx - 1].SlideLayout?.CommonSlideData?.ShapeTree;
                if (shapeTree == null)
                    throw new ArgumentException($"Slide layout {lIdx} has no shape tree");
            }
            else if (masterMlMatch.Success && !prefix.Contains("/slidelayout[", StringComparison.OrdinalIgnoreCase))
            {
                var mIdx = int.Parse(masterMlMatch.Groups[1].Value);
                var masters = _doc.PresentationPart?.SlideMasterParts?.ToList() ?? [];
                if (mIdx < 1 || mIdx > masters.Count)
                    throw new ArgumentException($"Slide master {mIdx} not found (total: {masters.Count})");
                shapeTree = masters[mIdx - 1].SlideMaster?.CommonSlideData?.ShapeTree;
                if (shapeTree == null)
                    throw new ArgumentException($"Slide master {mIdx} has no shape tree");
            }
            else if (layoutMlMatch.Success)
            {
                var lIdx = int.Parse(layoutMlMatch.Groups[1].Value);
                var allLayouts = (_doc.PresentationPart?.SlideMasterParts ?? Enumerable.Empty<SlideMasterPart>())
                    .SelectMany(m => m.SlideLayoutParts ?? Enumerable.Empty<SlideLayoutPart>()).ToList();
                if (lIdx < 1 || lIdx > allLayouts.Count)
                    throw new ArgumentException($"Slide layout {lIdx} not found (total: {allLayouts.Count})");
                shapeTree = allLayouts[lIdx - 1].SlideLayout?.CommonSlideData?.ShapeTree;
                if (shapeTree == null)
                    throw new ArgumentException($"Slide layout {lIdx} has no shape tree");
            }
            else
            {
                var slideMatch = Regex.Match(prefix, @"/slide\[(\d+)\]");
                if (!slideMatch.Success)
                    throw new ArgumentException($"Cannot resolve @{attrName}= outside of a slide context: {path}");
                var slideIdx = int.Parse(slideMatch.Groups[1].Value);

                var slideParts = GetSlideParts().ToList();
                if (slideIdx < 1 || slideIdx > slideParts.Count)
                    throw new ArgumentException($"Slide {slideIdx} not found (total: {slideParts.Count})");
                var slidePart = slideParts[PathIndex.ToArrayIndex(slideIdx)];
                shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree;
                if (shapeTree == null)
                    throw new ArgumentException($"Slide {slideIdx} has no shape tree");
            }

            // CONSISTENCY(group-id-scope): if the prefix has /group[N] segments
            // after /slide[N], scope the @id=/@name= search inside that nested
            // group's shape tree, not the slide-level shape tree.
            OpenXmlElement scope = shapeTree;
            var groupMatches = Regex.Matches(prefix, @"/group\[(\d+)\]");
            foreach (Match gm in groupMatches)
            {
                var gIdx = int.Parse(gm.Groups[1].Value);
                var groups = scope.Elements<GroupShape>().ToList();
                if (gIdx < 1 || gIdx > groups.Count)
                    throw new ArgumentException($"Group {gIdx} not found in scope (total: {groups.Count})");
                scope = groups[gIdx - 1];
            }

            var positionalIdx = isPlaceholder
                ? FindPlaceholderByBinding(scope, predicates)
                : FindElementByAttrInScope(scope, elementType, attrName, attrValue);
            var replacement = $"{m.Groups[1].Value}[{positionalIdx}]";
            sb.Append(replacement);
            cursor = m.Index + m.Length;
        }
        sb.Append(path, cursor, path.Length - cursor);
        return sb.ToString();
    }

    /// <summary>
    /// Resolve [last()] predicates to numeric indices by walking the path
    /// left-to-right and counting siblings of that element type at the
    /// resolved prefix. Mirrors XPath last() semantics so all downstream
    /// regex-based dispatch only ever sees numeric indices.
    /// CONSISTENCY(path-stability): handles slide root + shape-tree types
    /// (shape/picture/table/chart/connector/group/placeholder) + table tr/tc.
    /// Unrecognized parent contexts pass through unchanged so the existing
    /// "Invalid path index 'last()'" error still fires for unsupported cases.
    /// </summary>
    private string ResolveLastPredicates(string path)
    {
        if (string.IsNullOrEmpty(path) || !path.Contains("[last()]", StringComparison.OrdinalIgnoreCase))
            return path;

        var segments = path.TrimStart('/').Split('/');
        var rebuilt = new System.Text.StringBuilder();
        for (int i = 0; i < segments.Length; i++)
        {
            var seg = segments[i];
            var bracket = seg.IndexOf('[');
            if (bracket > 0 && seg.EndsWith("]", StringComparison.Ordinal))
            {
                var name = seg[..bracket];
                var idx = seg[(bracket + 1)..^1];
                if (idx.Equals("last()", StringComparison.OrdinalIgnoreCase))
                {
                    var prefix = rebuilt.ToString(); // already-resolved prefix, "" or "/slide[3]/..."
                    var count = CountLastSiblings(prefix, name.ToLowerInvariant());
                    if (count <= 0)
                        throw new ArgumentException($"Cannot resolve [last()] in segment '{seg}': no '{name}' siblings found at '{(prefix.Length == 0 ? "/" : prefix)}'.");
                    seg = $"{name}[{count}]";
                }
            }
            rebuilt.Append('/').Append(seg);
        }
        return rebuilt.ToString();
    }

    /// <summary>
    /// Count siblings of <paramref name="elementType"/> at the resolved
    /// <paramref name="prefix"/>. Prefix is empty (root) or a fully numeric
    /// path. Returns 0 when no count rule applies.
    /// </summary>
    private int CountLastSiblings(string prefix, string elementType)
    {
        // Root scope: /slide, /slidemaster, /slidelayout
        if (prefix.Length == 0)
        {
            return elementType switch
            {
                "slide" => GetSlideParts().Count(),
                "slidemaster" => _doc.PresentationPart?.SlideMasterParts?.Count() ?? 0,
                _ => 0,
            };
        }

        // Slide-scoped: /slide[N]
        var slideMatch = System.Text.RegularExpressions.Regex.Match(prefix, @"^/slide\[(\d+)\](.*)$");
        if (slideMatch.Success)
        {
            var slideIdx = int.Parse(slideMatch.Groups[1].Value);
            var slideParts = GetSlideParts().ToList();
            if (slideIdx < 1 || slideIdx > slideParts.Count) return 0;
            var slidePart = slideParts[PathIndex.ToArrayIndex(slideIdx)];
            var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree;
            if (shapeTree == null) return 0;

            var rest = slideMatch.Groups[2].Value;
            // Direct slide children (no further nesting in prefix)
            if (string.IsNullOrEmpty(rest))
                return CountInShapeContainer(shapeTree, elementType);

            // /slide[N]/group[M]/...[last()]
            OpenXmlElement scope = shapeTree;
            var groupMatches = System.Text.RegularExpressions.Regex.Matches(rest, @"/group\[(\d+)\]");
            int consumed = 0;
            foreach (System.Text.RegularExpressions.Match gm in groupMatches)
            {
                if (gm.Index != consumed) break; // non-contiguous; bail
                var gIdx = int.Parse(gm.Groups[1].Value);
                var groups = scope.Elements<GroupShape>().ToList();
                if (gIdx < 1 || gIdx > groups.Count) return 0;
                scope = groups[gIdx - 1];
                consumed = gm.Index + gm.Length;
            }
            var tail = rest[consumed..];
            if (string.IsNullOrEmpty(tail))
                return CountInShapeContainer(scope, elementType);

            // /slide[N]/.../table[M]/{tr|tc}[last()]
            var tblMatch = System.Text.RegularExpressions.Regex.Match(tail, @"^/table\[(\d+)\](.*)$");
            if (tblMatch.Success)
            {
                var tblIdx = int.Parse(tblMatch.Groups[1].Value);
                var tables = scope.Elements<DocumentFormat.OpenXml.Presentation.GraphicFrame>()
                    .Where(gf => gf.Descendants<Drawing.Table>().Any())
                    .ToList();
                if (tblIdx < 1 || tblIdx > tables.Count) return 0;
                var table = tables[tblIdx - 1].Descendants<Drawing.Table>().FirstOrDefault();
                if (table == null) return 0;
                var tableTail = tblMatch.Groups[2].Value;
                if (string.IsNullOrEmpty(tableTail))
                {
                    return elementType switch
                    {
                        "tr" or "row" => table.Elements<Drawing.TableRow>().Count(),
                        _ => 0,
                    };
                }
                // /tr[K]
                var trMatch = System.Text.RegularExpressions.Regex.Match(tableTail, @"^/tr\[(\d+)\]$");
                if (trMatch.Success && (elementType == "tc" || elementType == "cell"))
                {
                    var trIdx = int.Parse(trMatch.Groups[1].Value);
                    var rows = table.Elements<Drawing.TableRow>().ToList();
                    if (trIdx < 1 || trIdx > rows.Count) return 0;
                    return rows[trIdx - 1].Elements<Drawing.TableCell>().Count();
                }
            }
        }
        return 0;
    }

    /// <summary>
    /// Count direct children of <paramref name="container"/> matching the
    /// PPTX element-type vocabulary used by paths (shape, picture, table,
    /// chart, connector, group, placeholder, textbox, title).
    /// </summary>
    private static int CountInShapeContainer(OpenXmlElement container, string elementType)
    {
        return elementType switch
        {
            "shape" or "textbox" or "title" or "equation" => container.Elements<Shape>().Count(),
            "picture" or "pic" or "image" => container.Elements<Picture>().Count(),
            "table" => container.Elements<GraphicFrame>().Count(gf => gf.Descendants<Drawing.Table>().Any()),
            "chart" => container.Elements<GraphicFrame>().Count(gf =>
                gf.Descendants<DocumentFormat.OpenXml.Drawing.Charts.ChartReference>().Any() || IsExtendedChartFrame(gf)),
            "connector" or "connection" => container.Elements<ConnectionShape>().Count(),
            "group" => container.Elements<GroupShape>().Count(),
            "placeholder" or "ph" => container.Elements<Shape>()
                .Count(s => s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape != null),
            _ => 0,
        };
    }

    /// <summary>
    /// Resolve a placeholder binding selector (issue #260) to the 1-based
    /// positional index among the scope's placeholders. Supported predicate
    /// keys, chainable as `placeholder[@type=body][@idx=1]`:
    ///   @idx  — the p:ph idx binding id (ECMA-376 §19.3.1.36; absent = "0")
    ///   @type — the p:ph type (absent = "body" per schema default)
    ///   @id / @name — the shape's cNvPr, same semantics as other types.
    /// A selector matching more than one placeholder (e.g. body idx=1 and
    /// pic idx=1 sharing @idx) is an error prompting a narrower predicate,
    /// never a silent first-match.
    /// </summary>
    private static int FindPlaceholderByBinding(OpenXmlElement scope, List<(string Key, string Value)> predicates)
    {
        foreach (var (key, _) in predicates)
            if (key is not ("idx" or "type" or "id" or "name"))
                throw new ArgumentException(
                    $"Unsupported placeholder attribute selector '@{key}='. Supported: @idx=, @type=, @id=, @name=.");

        var placeholders = scope.Elements<Shape>()
            .Select((s, i) => (Shape: s,
                Ph: s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape))
            .Where(t => t.Ph != null)
            .ToList();

        string PhIdx(PlaceholderShape ph) => ph.Index?.Value.ToString() ?? "0";
        string PhType(PlaceholderShape ph) => ph.Type?.InnerText ?? "body";
        string Describe((Shape Shape, PlaceholderShape? Ph) t) =>
            $"type={PhType(t.Ph!)} idx={PhIdx(t.Ph!)}" +
            (t.Shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value is { Length: > 0 } n
                ? $" name=\"{n}\"" : "");

        var hits = new List<int>();
        for (int i = 0; i < placeholders.Count; i++)
        {
            var (shape, ph) = placeholders[i];
            bool all = true;
            foreach (var (key, value) in predicates)
            {
                var actual = key switch
                {
                    "idx" => PhIdx(ph!),
                    "type" => PhType(ph!),
                    "id" => shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value.ToString(),
                    _ => shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value,
                };
                if (!string.Equals(actual, value, StringComparison.OrdinalIgnoreCase)) { all = false; break; }
            }
            if (all) hits.Add(i + 1);
        }

        var selector = string.Concat(predicates.Select(p => $"[@{p.Key}={p.Value}]"));
        if (hits.Count == 0)
            throw new ArgumentException(
                $"No placeholder matches placeholder{selector}. Available: " +
                (placeholders.Count == 0 ? "(none)"
                    : string.Join(", ", placeholders.Select((t, i) => $"placeholder[{i + 1}] ({Describe(t)})"))));
        if (hits.Count > 1)
            throw new ArgumentException(
                $"placeholder{selector} is ambiguous — matches " +
                string.Join(", ", hits.Select(h => $"placeholder[{h}] ({Describe(placeholders[h - 1])})")) +
                ". Narrow it with a combined selector, e.g. placeholder[@type=body][@idx=1].");
        return hits[0];
    }

    /// <summary>
    /// Find the 1-based positional index of an element within its type group by @id= or @name=.
    /// </summary>
    private static int FindElementByAttr(ShapeTree shapeTree, string elementType, string attrName, string attrValue)
        => FindElementByAttrInScope(shapeTree, elementType, attrName, attrValue);

    /// <summary>
    /// Like <see cref="FindElementByAttr"/> but searches direct children of any
    /// container element (ShapeTree or GroupShape). Used to scope @id=/@name=
    /// lookups inside nested groups.
    /// </summary>
    private static int FindElementByAttrInScope(OpenXmlElement scope, string elementType, string attrName, string attrValue)
    {
        var elements = elementType switch
        {
            "shape" or "textbox" or "title" or "equation" => scope.Elements<Shape>()
                .Select(s => (element: (OpenXmlElement)s, nvPr: s.NonVisualShapeProperties?.NonVisualDrawingProperties)).ToList(),
            "picture" or "pic" or "image" => scope.Elements<Picture>()
                .Select(p => (element: (OpenXmlElement)p, nvPr: p.NonVisualPictureProperties?.NonVisualDrawingProperties)).ToList(),
            "table" => scope.Elements<GraphicFrame>()
                .Where(gf => gf.Descendants<Drawing.Table>().Any())
                .Select(gf => (element: (OpenXmlElement)gf, nvPr: gf.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties)).ToList(),
            "chart" => scope.Elements<GraphicFrame>()
                .Where(gf => gf.Descendants<DocumentFormat.OpenXml.Drawing.Charts.ChartReference>().Any() || IsExtendedChartFrame(gf))
                .Select(gf => (element: (OpenXmlElement)gf, nvPr: gf.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties)).ToList(),
            "connector" or "connection" => scope.Elements<ConnectionShape>()
                .Select(c => (element: (OpenXmlElement)c, nvPr: c.NonVisualConnectionShapeProperties?.NonVisualDrawingProperties)).ToList(),
            "group" => scope.Elements<GroupShape>()
                .Select(g => (element: (OpenXmlElement)g, nvPr: g.NonVisualGroupShapeProperties?.NonVisualDrawingProperties)).ToList(),
            "video" or "audio" => scope.Elements<Picture>()
                .Select(p => (element: (OpenXmlElement)p, nvPr: p.NonVisualPictureProperties?.NonVisualDrawingProperties)).ToList(),
            _ => throw new ArgumentException($"Unknown element type '{elementType}' for @{attrName}= addressing")
        };

        for (int i = 0; i < elements.Count; i++)
        {
            var nvPr = elements[i].nvPr;
            if (nvPr == null) continue;

            if (attrName == "id" && nvPr.Id?.Value.ToString() == attrValue)
                return i + 1;
            if (attrName == "name" && MatchesShapeName(nvPr.Name?.Value, attrValue))
                return i + 1;
        }

        throw new ArgumentException($"No {elementType} found with @{attrName}={attrValue}");
    }

    /// <summary>
    /// Build a path segment using @id= if the element has a cNvPr.Id, otherwise use positional index.
    /// E.g. "shape[@id=5]" or "shape[2]".
    /// </summary>
    internal static string BuildElementPathSegment(string elementType, OpenXmlElement element, int positionalIndex)
    {
        var id = GetCNvPrId(element);
        return id.HasValue
            ? $"{elementType}[@id={id.Value}]"
            : $"{elementType}[{positionalIndex}]";
    }
}
