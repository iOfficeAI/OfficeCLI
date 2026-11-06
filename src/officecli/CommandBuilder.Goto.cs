// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using OfficeCli.Core;

namespace OfficeCli;

static partial class CommandBuilder
{
    // ==================== goto ====================
    //
    // Push a one-shot scroll target to all SSE clients of a running watch.
    // Does not open the file, does not mutate cached HTML, does not bump
    // the version — pure runtime navigation. Mirrors mark/unmark in being
    // a separate top-level command that talks to watch over the named
    // pipe (CONSISTENCY(watch-runtime-cmd)).
    //
    // Paths are resolved by the watch server against its cached rendered
    // HTML. Server-issued mark ids target viewer-only highlighted fragments.

    private static Command BuildGotoCommand(Option<bool> jsonOption, string name = "goto")
    {
        var fileArg = new Argument<FileInfo>("file") { Description = "Office document path" };
        var pathArg = new Argument<string?>("path")
        {
            Description = "Element path to center (Word, Excel, or PowerPoint)",
            Arity = ArgumentArity.ZeroOrOne,
        };
        var markIdOpt = new Option<string?>("--mark-id")
        {
            Description = "Server-issued viewer mark id to center",
        };

        var cmd = new Command(name,
            "Center a document path or viewer mark in the running watch viewer(s). Broadcast to all SSE clients of the file.");
        cmd.Add(fileArg);
        cmd.Add(pathArg);
        cmd.Add(markIdOpt);
        cmd.Add(jsonOption);

        cmd.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(fileArg)!;
            var path = OfficeCli.Core.MsysPathHint.Restore(result.GetValue(pathArg));
            var markId = result.GetValue(markIdOpt)?.Trim();
            var hasPath = !string.IsNullOrWhiteSpace(path);
            var hasMark = !string.IsNullOrWhiteSpace(markId);
            if (hasPath == hasMark)
            {
                var err = "Specify exactly one document path or --mark-id.";
                if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeError(err));
                else Console.Error.WriteLine(err);
                return 2;
            }

            if (!WatchServer.IsWatching(file.FullName))
            {
                var err = $"No watch process is running for {file.Name}.";
                if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeError(err));
                else Console.Error.WriteLine(err);
                return 1;
            }

            ScrollResult scroll;
            string target;
            if (hasMark)
            {
                scroll = WatchNotifier.TryScrollMark(file.FullName, markId!);
                target = $"mark {markId}";
            }
            else
            {
                var normalizedPath = path!.Trim();
                // Preserve legacy Word paragraph/table aliases and anchors.
                var legacySelector = file.Extension.Equals(
                    ".docx",
                    StringComparison.OrdinalIgnoreCase)
                    ? WatchMessage.ExtractWordScrollTarget(normalizedPath)
                    : null;
                scroll = legacySelector != null
                    ? WatchNotifier.TryScroll(file.FullName, legacySelector)
                    : WatchNotifier.TryScrollPath(file.FullName, normalizedPath);
                target = normalizedPath;
            }
            if (scroll.Kind == ScrollResult.K.NotFound)
            {
                var err = $"Cannot scroll to '{target}': {scroll.Error}.";
                if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeError(err));
                else Console.Error.WriteLine(err);
                return 1;
            }
            if (scroll.Kind == ScrollResult.K.NoWatch)
            {
                var err = $"No watch process is running for {file.Name}.";
                if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeError(err));
                else Console.Error.WriteLine(err);
                return 1;
            }

            var msg = $"Scrolled watcher(s) to {target}";
            if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeText(msg));
            else Console.WriteLine(msg);
            return 0;
        }, json); });

        return cmd;
    }
}
