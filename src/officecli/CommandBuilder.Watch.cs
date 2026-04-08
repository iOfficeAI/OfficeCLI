// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using OfficeCli.Core;

namespace OfficeCli;

static partial class CommandBuilder
{
    private static Command BuildWatchCommand()
    {
        var watchFileArg = new Argument<FileInfo>("file") { Description = "Office document path (.pptx, .xlsx, .docx)" };
        var watchPortOpt = new Option<int>("--port") { Description = "HTTP port for preview server" };
        watchPortOpt.DefaultValueFactory = _ => 18080;

        var watchCommand = new Command("watch", "Start a live preview server that auto-refreshes when the document changes");
        watchCommand.Add(watchFileArg);
        watchCommand.Add(watchPortOpt);

        watchCommand.SetAction(result => SafeRun(() =>
        {
            var file = result.GetValue(watchFileArg)!;
            var port = result.GetValue(watchPortOpt);

            // Render initial HTML from existing file content
            string? initialHtml = null;
            if (file.Exists)
            {
                try
                {
                    using var handler = DocumentHandlerFactory.Open(file.FullName, editable: false);
                    if (handler is OfficeCli.Handlers.PowerPointHandler ppt)
                        initialHtml = ppt.ViewAsHtml();
                    else if (handler is OfficeCli.Handlers.ExcelHandler excel)
                        initialHtml = excel.ViewAsHtml();
                    else if (handler is OfficeCli.Handlers.WordHandler word)
                        initialHtml = word.ViewAsHtml();
                }
                catch { /* ignore — will show waiting page */ }
            }

            using var cts = new CancellationTokenSource();
            Console.CancelKeyPress += (_, e) => { e.Cancel = true; cts.Cancel(); };

            using var watch = new WatchServer(file.FullName, port, initialHtml: initialHtml);
            // BUG-BT-R302: SIGTERM (pkill, kill) does NOT run `using` finally
            // blocks, so the WatchServer.Dispose() pipe-socket cleanup never
            // runs and stale CoreFxPipe_* files accumulate in $TMPDIR. Hook
            // ProcessExit so a graceful SIGTERM still triggers Dispose. SIGKILL
            // is unrecoverable by definition (kernel-level), so this only
            // covers cooperative shutdown.
            AppDomain.CurrentDomain.ProcessExit += (_, _) =>
            {
                try { watch.Dispose(); } catch { /* best effort */ }
            };
            watch.RunAsync(cts.Token).GetAwaiter().GetResult();
            return 0;
        }));

        return watchCommand;
    }

    private static Command BuildUnwatchCommand()
    {
        var unwatchFileArg = new Argument<FileInfo>("file") { Description = "Office document path (.pptx, .xlsx, .docx)" };
        var unwatchCommand = new Command("unwatch", "Stop the watch preview server for the document");
        unwatchCommand.Add(unwatchFileArg);

        unwatchCommand.SetAction(result => SafeRun(() =>
        {
            var file = result.GetValue(unwatchFileArg)!;
            if (WatchNotifier.SendClose(file.FullName))
                Console.WriteLine($"Watch stopped for {file.Name}");
            else
                Console.Error.WriteLine($"No watch running for {file.Name}");
            return 0;
        }));

        return unwatchCommand;
    }
}
