// Copyright 2025 OfficeCli (officecli.ai)
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
            var path = result.GetValue(pathArg)!;
            var depth = result.GetValue(depthOpt);
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
                Console.WriteLine(OutputFormatter.WrapEnvelope(
                    OutputFormatter.FormatNode(node, OutputFormat.Json)));
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

        var queryTextOpt = new Option<string?>("--text") { Description = "Filter results to elements containing this text (case-insensitive)" };

        var queryCommand = new Command("query", "Query document elements with CSS-like selectors");
        queryCommand.Add(queryFileArg);
        queryCommand.Add(selectorArg);
        queryCommand.Add(jsonOption);
        queryCommand.Add(queryTextOpt);

        queryCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(queryFileArg)!;
            var selector = result.GetValue(selectorArg)!;
            var textFilter = result.GetValue(queryTextOpt);

            if (TryResident(file.FullName, req =>
            {
                req.Command = "query";
                req.Json = json;
                req.Args["selector"] = selector;
                if (textFilter != null) req.Args["text"] = textFilter;
            }, json) is {} rc) return rc;

            var format = json ? OutputFormat.Json : OutputFormat.Text;

            using var handler = DocumentHandlerFactory.Open(file.FullName);
            var filters = OfficeCli.Core.AttributeFilter.Parse(selector);
            var (results, warnings) = OfficeCli.Core.AttributeFilter.ApplyWithWarnings(handler.Query(selector), filters);
            if (!string.IsNullOrEmpty(textFilter))
                results = results.Where(n => n.Text != null && n.Text.Contains(textFilter, StringComparison.OrdinalIgnoreCase)).ToList();
            if (json)
            {
                var cliWarnings = warnings.Select(w => new OfficeCli.Core.CliWarning { Message = w, Code = "filter_warning" }).ToList();
                Console.WriteLine(OutputFormatter.WrapEnvelope(
                    OutputFormatter.FormatNodes(results, OutputFormat.Json),
                    cliWarnings.Count > 0 ? cliWarnings : null));
            }
            else
            {
                foreach (var w in warnings) Console.Error.WriteLine(w);
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
}
