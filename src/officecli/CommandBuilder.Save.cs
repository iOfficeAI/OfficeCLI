// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using OfficeCli.Core;

namespace OfficeCli;

static partial class CommandBuilder
{
    // ==================== save command ====================
    //
    // Flush the resident's in-memory document to disk WITHOUT ending the
    // session. Only meaningful inside a running resident — agents that build
    // a workbook incrementally and need mid-build snapshots (e.g. for a
    // third-party Excel viewer that ingests the .xlsx package directly) use
    // this to recover the parse-amortization benefit of resident mode while
    // still publishing snapshots on demand.
    //
    // Non-resident mode is rejected on purpose: each non-resident command
    // already opens-mutates-closes (close = save), so there's no pending
    // in-memory state to flush. A no-op success would invite confused
    // "save just in case" code; an error tells the user they're in the
    // wrong mode.
    private static Command BuildSaveCommand(Option<bool> jsonOption)
    {
        var saveFileArg = new Argument<FileInfo>("file") { Description = "Office document path" };
        var saveCommand = new Command("save", "Flush in-memory changes to disk, keeping the resident running. Run before a non-officecli program reads the file (officecli's own reads always see edits; a direct disk reader sees the pre-edit file until a flush). A live resident also auto-flushes shortly after going idle (adaptive 2-10s; see OFFICECLI_RESIDENT_FLUSH: each|auto|<seconds>|off). No-op if no resident is active.");
        saveCommand.Add(saveFileArg);
        saveCommand.Add(jsonOption);

        saveCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(saveFileArg)!;
            var filePath = file.FullName;

            // Probe without auto-starting. `save` is meaningless without a
            // pre-existing in-memory session, so we deliberately skip the
            // TryResident auto-start path that other verbs use.
            if (!ResidentClient.TryConnect(filePath, out _))
            {
                if (ResidentRecoveryMarker.TryConsume(filePath))
                    throw ResidentRecoveryMarker.CreateLossException(filePath);
                // No resident session to flush. In the non-resident model the
                // document on disk is already current (each mutation eager-saved),
                // so save is a no-op SUCCESS rather than an error — keeping
                // "edit, then save/close" a safe habit regardless of backend.
                var msg = $"{file.Name} is already saved to disk.";
                if (json)
                    Console.WriteLine(OutputFormatter.WrapEnvelopeText(msg));
                else
                    Console.WriteLine(msg);
                return 0;
            }

            var request = new ResidentRequest { Command = "save", Json = json };
            var response = ResidentClient.TrySend(
                filePath, request,
                maxRetries: ResidentBusyMaxRetries,
                connectTimeoutMs: ResidentBusyConnectTimeoutMs);
            if (response == null)
            {
                var msg = $"Resident for {file.Name} is running but the save command could not be delivered (main pipe busy or unresponsive).";
                if (json)
                    Console.WriteLine(OutputFormatter.WrapEnvelopeError(msg));
                else
                    Console.Error.WriteLine($"Error: {msg}");
                return 3;
            }

            if (!string.IsNullOrEmpty(response.Stdout))
                Console.WriteLine(response.Stdout);
            if (!string.IsNullOrEmpty(response.Stderr))
                Console.Error.WriteLine(response.Stderr);
            return response.ExitCode;
        }, json); });

        return saveCommand;
    }
}
