// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using OfficeCli.Core;
using OfficeCli.Handlers;

namespace OfficeCli;

static partial class CommandBuilder
{
    private static Command BuildGetCommand(Option<bool> jsonOption)
    {
        var getFileArg = new Argument<FileInfo>("file") { Description = "Office document path (required even with open/close mode)" };
        var pathArg = new Argument<string>("path") { Description = "DOM path (e.g. /body/p[1]) or 'selected' to read the current watch selection" };
        pathArg.DefaultValueFactory = _ => "/";
        var depthOpt = new Option<int>("--depth") { Description = "Depth of child nodes to include" };
        depthOpt.DefaultValueFactory = _ => 1;
        var saveOpt = new Option<string?>("--save") { Description = "Extract the backing binary payload (picture/ole/media) to this file path" };

        var getCommand = new Command("get", "Get a document node by path");
        getCommand.Add(getFileArg);
        getCommand.Add(pathArg);
        getCommand.Add(depthOpt);
        getCommand.Add(saveOpt);
        getCommand.Add(jsonOption);

        getCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(getFileArg)!;
            var path = MsysPathHint.Restore(result.GetValue(pathArg)!)!;
            var depth = result.GetValue(depthOpt);
            // CONSISTENCY(dos-hardening): cap user-supplied depth so a huge
            // --depth on a deeply-nested doc can't drive the node-building
            // recursion (and its O(n^2) InnerText/OuterXml-per-node cost) into
            // a multi-minute hang or stack overflow. See DocumentLimits.
            if (depth > DocumentLimits.MaxRecursionDepth)
                depth = DocumentLimits.MaxRecursionDepth;
            var savePath = result.GetValue(saveOpt);

            // Special pseudo-path "selected" — query the running watch process
            // for the currently-selected element paths and resolve them to nodes.
            if (string.Equals(path, "selected", StringComparison.OrdinalIgnoreCase))
            {
                return GetSelectedAction(file.FullName, depth, json);
            }

            if (TryResident(file.FullName, req =>
            {
                req.Command = "get";
                req.Json = json;
                req.Args["path"] = path;
                req.Args["depth"] = depth.ToString();
                if (!string.IsNullOrEmpty(savePath)) req.Args["save"] = savePath;
            }, json) is {} rc) return rc;

            using var handler = DocumentHandlerFactory.Open(file.FullName);
            var node = handler.Get(path, depth);

            // CONSISTENCY(get-not-found-exit): some handler Get paths surface
            // "not found" via DocumentNode { Type = "error" } instead of
            // throwing (e.g. /numbering/abstractNum[@id=999]). Other paths
            // throw and exit 1 via SafeRun. Treat error-type nodes the same
            // way so callers get a consistent non-zero exit on missing paths.
            if (string.Equals(node.Type, "error", StringComparison.Ordinal))
            {
                var err = node.Text ?? $"Path not found: {path}";
                if (json)
                    Console.WriteLine(OutputFormatter.WrapEnvelopeError(err));
                else
                    Console.Error.WriteLine($"Error: {err}");
                return 1;
            }

            // --save <path>: extract the binary payload backing an OLE /
            // picture / media node to disk. The handler exposes this via
            // TryExtractBinary which looks up the node's relId and copies
            // the part's stream. When the node has no backing binary, we
            // surface a clear error instead of silently succeeding.
            if (!string.IsNullOrEmpty(savePath))
            {
                if (!handler.TryExtractBinary(path, savePath, out var contentType, out var byteCount))
                {
                    var err = $"Node at '{path}' has no binary payload to extract (only ole/picture/media/embedded nodes can be saved).";
                    if (json)
                        Console.WriteLine(OutputFormatter.WrapEnvelopeError(err));
                    else
                        Console.Error.WriteLine($"Error: {err}");
                    return 1;
                }
                node.Format["savedTo"] = savePath;
                node.Format["savedBytes"] = byteCount;
                if (!string.IsNullOrEmpty(contentType))
                    node.Format["savedContentType"] = contentType!;
            }

            if (json)
                // Unified envelope contract: single-path get returns the same
                // {matches, results: [...]} shape as `get selected` and `query`,
                // so agents and scripts can use one jq path everywhere. Text
                // mode keeps the rich single-node rendering.
                Console.WriteLine(OutputFormatter.WrapEnvelope(
                    OutputFormatter.FormatNodes(new List<DocumentNode> { node }, OutputFormat.Json)));
            else
                Console.WriteLine(OutputFormatter.FormatNode(node, OutputFormat.Text));
            return 0;
        }, json); });

        return getCommand;
    }

    private static int GetSelectedAction(string filePath, int depth, bool json)
    {
        var paths = WatchNotifier.QuerySelection(filePath);
        if (paths == null)
        {
            var msg = $"no watch running for {Path.GetFileName(filePath)}. Start one with: officecli watch \"{filePath}\"";
            if (json)
                Console.WriteLine(OutputFormatter.WrapEnvelopeError(msg));
            else
                Console.Error.WriteLine($"Error: {msg}");
            return 1;
        }

        // Resolve each path to a DocumentNode. Skip paths that no longer exist
        // (e.g. element removed since selection was made) — silently drop them.
        var nodes = new List<OfficeCli.Core.DocumentNode>();
        if (paths.Length > 0)
        {
            using var handler = DocumentHandlerFactory.Open(filePath);
            foreach (var p in paths)
            {
                try
                {
                    var n = handler.Get(p, depth);
                    if (n != null) nodes.Add(n);
                }
                catch
                {
                    // path no longer resolves — drop
                }
            }
        }

        // Flatten row/column nodes into their children so text output is
        // grep-friendly (one cell per line instead of a single "/Sheet1/col[C]" line).
        var flat = new List<OfficeCli.Core.DocumentNode>();
        foreach (var n in nodes)
        {
            if (n.Children.Count > 0 && n.Type is "column" or "row")
                flat.AddRange(n.Children);
            else
                flat.Add(n);
        }

        if (json)
        {
            Console.WriteLine(OutputFormatter.WrapEnvelope(
                OutputFormatter.FormatNodes(flat, OutputFormat.Json)));
        }
        else
        {
            Console.WriteLine(OutputFormatter.FormatNodes(flat, OutputFormat.Text));
        }
        return 0;
    }

    private static Command BuildQueryCommand(Option<bool> jsonOption)
    {
        var queryFileArg = new Argument<FileInfo>("file") { Description = "Office document path (required even with open/close mode)" };
        var selectorArg = new Argument<string>("selector") { Description = "CSS-like selector (e.g. paragraph[style=Normal] > run[font!=Arial])" };

        var queryFindOpt = new Option<string?>("--find") { Description = "Filter results to elements containing this text (case-insensitive substring)" };
        var queryCompactOpt = new Option<bool>("--compact") { Description = "One line per element in document order: path<TAB>[label]<TAB>\"text(truncated at 60, … mark)\"; empty text shows (empty); tables fold to [table RxC]. Final line is always 'total: N of M elements / K slides' (pptx) or 'total: N of M elements' (docx — never gains a container segment): N = element lines above (lineCount-1 == N proves you read everything), M = all top-level frames. Full-document listing: selector '*' lists all top-level frames/blocks in document order and makes N == M (docx '*' folds tables without descending into cell paragraphs; 'paragraph, table' descends). Labels are a closed set (pptx: title/placeholder/textbox/shape/picture/chart/connector/group/equation + 'table RxC'; docx: style name). This format is a stability contract: columns/labels may be added, never changed or reordered. pptx/docx only (xlsx: use 'view text --range'). Add columns with --fields." };
        var queryFieldsOpt = new Option<string?>("--fields") { Description = "Comma-separated Format keys appended as extra k=v columns in --compact output (e.g. x,y,width)" };

        var queryCommand = new Command("query", "Query document elements with CSS-like selectors");
        queryCommand.Add(queryFileArg);
        queryCommand.Add(selectorArg);
        queryCommand.Add(jsonOption);
        queryCommand.Add(queryFindOpt);
        queryCommand.Add(queryCompactOpt);
        queryCommand.Add(queryFieldsOpt);

        queryCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(queryFileArg)!;
            var selector = MsysPathHint.Restore(result.GetValue(selectorArg)!)!;
            var textFilter = result.GetValue(queryFindOpt);
            var compact = result.GetValue(queryCompactOpt);
            var fields = result.GetValue(queryFieldsOpt);
            if (compact && json)
                throw new OfficeCli.Core.CliException("--compact is a plain-text line format; drop --json (or drop --compact for the JSON tree).") { Code = "invalid_value" };

            if (TryResident(file.FullName, req =>
            {
                req.Command = "query";
                req.Json = json;
                req.Args["selector"] = selector;
                if (textFilter != null) req.Args["find"] = textFilter;
                if (compact) req.Args["compact"] = "true";
                if (fields != null) req.Args["fields"] = fields;
            }, json) is {} rc) return rc;

            var format = json ? OutputFormat.Json : OutputFormat.Text;

            using var handler = DocumentHandlerFactory.Open(file.FullName);
            // CONSISTENCY(cell-selector-alias): the Excel cell selector accepts short
            // aliases (bold -> font.bold, size -> font.size, ...). FilterSelector
            // applies the same normalization, runs the boolean and/or engine, and
            // routes a pure-AND (flat) selector through the exact legacy path.
            // SelectorTargetsCells strips an optional sheet prefix (Sheet1!cell...,
            // /Sheet1/cell...) before the element check — without it, sheet-scoped
            // cell selectors skip alias normalization and drop all matches.
            Func<string, string>? keyResolver =
                handler is OfficeCli.Handlers.ExcelHandler
                && OfficeCli.Handlers.ExcelHandler.SelectorTargetsCells(selector)
                    ? OfficeCli.Handlers.ExcelHandler.ResolveCellAttributeAlias : null;
            var (results, warnings) = OfficeCli.Core.AttributeFilter.FilterSelector(selector, handler.Query, keyResolver);
            if (!string.IsNullOrEmpty(textFilter))
                results = results.Where(n => n.Text != null && OfficeCli.Core.AttributeFilter.MatchesTextFilter(n.Text, textFilter)).ToList();
            if (compact)
            {
                foreach (var w in warnings) Console.Error.WriteLine(w.Message);
                Console.WriteLine(FormatNodesCompact(handler, results, fields));
                return 0;
            }
            if (json)
            {
                // CONSISTENCY(query-json-children): Query returns nodes with empty
                // Children but populated ChildCount (handlers build query nodes at
                // depth=0 to avoid expensive subtree walks). For --json output we
                // hydrate children via Get(path, depth=1) so consumers see the same
                // shape that `get --json` produces.
                for (int i = 0; i < results.Count; i++)
                {
                    var n = results[i];
                    if (n.ChildCount > 0 && n.Children.Count == 0 && !string.IsNullOrEmpty(n.Path))
                    {
                        try
                        {
                            var hydrated = handler.Get(n.Path, depth: 1);
                            if (hydrated?.Children != null && hydrated.Children.Count > 0)
                                n.Children.AddRange(hydrated.Children);
                        }
                        catch { /* path may not be Get-resolvable; leave as-is */ }
                    }
                }
                var cliWarnings = warnings.Select(w => new OfficeCli.Core.CliWarning
                {
                    Message = w.Message,
                    Code = w.Code,
                    Kind = w.Kind,
                    Key = w.Key,
                    Value = w.Value,
                    Available = w.Available,
                    Suggestion = w.Suggestion,
                }).ToList();
                Console.WriteLine(OutputFormatter.WrapEnvelope(
                    OutputFormatter.FormatNodes(results, OutputFormat.Json),
                    cliWarnings.Count > 0 ? cliWarnings : null));
            }
            else
            {
                foreach (var w in warnings) Console.Error.WriteLine(w.Message);
                var output = OutputFormatter.FormatNodes(results, OutputFormat.Text);
                if (!string.IsNullOrEmpty(output))
                    Console.WriteLine(output);
                if (results.Count == 0)
                {
                    var ext = file.Extension.ToLowerInvariant().TrimStart('.');
                    Console.Error.WriteLine($"No matches. Run 'officecli {ext} query' for selector syntax.");
                }
            }
            return 0;
        }, json); });

        return queryCommand;
    }

    /// <summary>
    /// `query --compact` line format. STABILITY CONTRACT — this output is
    /// consumed programmatically (parsed line-by-line, counted for
    /// read-completeness accounting); treat every token below as API:
    /// column order, the TAB separator, the `…` truncation mark, `(empty)`,
    /// the `[label]` bracket form, and the total-line shape may only gain
    /// NEW trailing columns, never change or reorder existing ones. Any
    /// change lands in the CHANGELOG.
    ///
    ///   {path}\t[{label}]\t"{text ≤60 chars, \t/\n/"/\\ escaped}"
    ///   {path}\t[table {R}x{C}]                     (tables fold; no text col)
    ///   {path}\t[{label}]\t(empty)                  (no text)
    ///   ...--fields k1,k2 appends \tk1=v1\tk2=v2 columns (missing key → k=)
    ///   total: {N} of {M} elements / {K} slides     (pptx)
    ///   total: {N} of {M} elements                  (docx)
    ///
    /// N = exactly the number of lines above the total line (post-filter;
    /// a folded table counts as 1) so `lineCount - 1 == N` proves the reader
    /// saw the whole result. M = all top-level frames in the document
    /// (pptx: shapes/pictures/tables/charts/connectors/groups across slides;
    /// docx: body-level blocks). The total line is always emitted (N=0
    /// included) and is always the last line, exactly once. The docx total
    /// has NO container segment — that absence is itself frozen (appending
    /// one later would be a total-line change, which the contract forbids).
    /// Full listing: '*' on both formats (docx aliases it to top-level
    /// body blocks in document order — see WordHandler.QueryDispatch).
    ///
    /// Element lines are in document order: pptx sorts by slide index then
    /// z-order (multi-type selectors like '*' would otherwise group by type),
    /// docx follows document flow. Labels are a CLOSED SET per format —
    /// pptx: title/placeholder/textbox/shape/picture/chart/connector/group/
    /// equation + the folded 'table RxC'; docx: the paragraph's style name
    /// (open set of values, fixed [style] position). New label values may be
    /// added; existing ones never change meaning.
    /// </summary>
    internal static string FormatNodesCompact(IDocumentHandler handler, List<DocumentNode> results, string? fields)
    {
        if (handler is ExcelHandler)
            throw new OfficeCli.Core.CliException(
                "--compact is not supported for xlsx: 'view text' is already the compact per-row form ([/Sheet1/row[N]] A1=v ...). Use 'view text' or 'view text --range Sheet1!A1:C10'.")
            { Code = "invalid_value" };

        var fieldList = string.IsNullOrWhiteSpace(fields)
            ? null
            : fields.Split(',').Select(f => f.Trim()).Where(f => f.Length > 0).ToList();

        // Document order (see contract above): multi-type selectors return
        // results grouped per element type; re-sort pptx frames by slide index
        // then z-order so the line sequence mirrors the deck. Stable sort keeps
        // relative order where either key is missing. docx query results
        // already follow document flow.
        if (handler is PowerPointHandler)
        {
            results = results
                .Select((n, i) => (n, i))
                .OrderBy(t => System.Text.RegularExpressions.Regex.Match(t.n.Path, @"^/slide\[(\d+)\]") is { Success: true } m
                    ? int.Parse(m.Groups[1].Value) : int.MaxValue)
                .ThenBy(t => t.n.Format.TryGetValue("zorder", out var z) && int.TryParse(z?.ToString(), out var zi)
                    ? zi : int.MaxValue)
                .ThenBy(t => t.i)
                .Select(t => t.n)
                .ToList();
        }

        var sb = new System.Text.StringBuilder();
        foreach (var n in results)
        {
            sb.Append(n.Path);
            if (n.Type == "table" && n.Format.TryGetValue("rows", out var r) && n.Format.TryGetValue("cols", out var c))
            {
                sb.Append('\t').Append($"[table {r}x{c}]");
            }
            else
            {
                // pptx Type already carries the semantic label (title/placeholder/
                // textbox/shape/...); docx paragraphs label by style name.
                var label = !string.IsNullOrEmpty(n.Style) ? n.Style : n.Type;
                sb.Append('\t').Append('[').Append(label).Append(']');
                sb.Append('\t').Append(CompactText(n.Text));
            }
            if (fieldList != null)
                foreach (var f in fieldList)
                    sb.Append('\t').Append(f).Append('=')
                      .Append(n.Format.TryGetValue(f, out var v) && v != null ? v.ToString() : "");
            sb.Append('\n');
        }

        var (total, containerSuffix) = CountCompactDenominator(handler);
        sb.Append($"total: {results.Count} of {total} elements{containerSuffix}");
        return sb.ToString();
    }

    private static string CompactText(string? text)
    {
        if (string.IsNullOrEmpty(text)) return "(empty)";
        var t = text.Replace("\\", "\\\\").Replace("\t", "\\t")
                    .Replace("\r", "").Replace("\n", "\\n").Replace("\"", "\\\"");
        t = OfficeCli.Core.DisplayText.Truncate(t, 60);
        return "\"" + t + "\"";
    }

    /// <summary>
    /// Denominator for the compact total line: every top-level frame in the
    /// document, independent of the user's selector, so the agent can compare
    /// "what I matched" against "what exists".
    /// </summary>
    private static (int Total, string ContainerSuffix) CountCompactDenominator(IDocumentHandler handler)
    {
        if (handler is PowerPointHandler)
        {
            int slides = handler.Query("slide").Count;
            // Frame selectors overlap (title/textbox are shapes); dedupe by path.
            var paths = new HashSet<string>();
            foreach (var sel in new[] { "shape", "picture", "table", "chart", "connector", "group" })
            {
                try { foreach (var n in handler.Query(sel)) paths.Add(n.Path); }
                catch { /* selector unsupported on this doc — skip */ }
            }
            return (paths.Count, $" / {slides} slides");
        }
        // docx: body-level content blocks (paragraphs + tables + sdt + ...).
        // sectPr is layout metadata, not an addressable element — exclude it
        // so the denominator matches what a selector can actually reach.
        try
        {
            var body = handler.Get("/body", depth: 1);
            return (body.Children.Count(c => c.Type != "section"), "");
        }
        catch
        {
            return (0, "");
        }
    }
}
