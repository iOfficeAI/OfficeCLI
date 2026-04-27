// Copyright 2025 OfficeCli (officecli.ai)
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
    // Word: path like /body/p[5] or /body/table[2] — resolves via
    // WatchMessage.ExtractWordScrollTarget. PPT/Excel: not yet wired in
    // (anchor coverage is the gap, not the command itself).

    private static Command BuildGotoCommand(Option<bool> jsonOption)
    {
        var fileArg = new Argument<FileInfo>("file") { Description = "Office document path (.docx)" };
        var pathArg = new Argument<string>("path") { Description = "Element path to scroll to (e.g. /body/p[5], /body/table[1], /body/table[1]/tr[2]/tc[3])" };

        var cmd = new Command("goto",
            "Scroll the running watch viewer(s) to the given element. Path resolves to an HTML anchor; broadcast to all SSE clients of the file. Word: paragraph, table, table row, table cell.");
        cmd.Add(fileArg);
        cmd.Add(pathArg);
        cmd.Add(jsonOption);

        cmd.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(fileArg)!;
            var path = result.GetValue(pathArg)!;

            var selector = WatchMessage.ExtractWordScrollTarget(path);
            if (selector == null)
            {
                var err = $"Cannot resolve scroll target for path '{path}'. Supported: /body/p[N], /body/paragraph[N], /body/table[N], /body/table[N]/tr[R], /body/table[N]/tr[R]/tc[C].";
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

            WatchNotifier.NotifyIfWatching(file.FullName, new WatchMessage
            {
                Action = "scroll",
                ScrollTo = selector,
            });

            var msg = $"Scrolled watcher(s) to {path} ({selector})";
            if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeText(msg));
            else Console.WriteLine(msg);
            return 0;
        }, json); });

        return cmd;
    }
}
