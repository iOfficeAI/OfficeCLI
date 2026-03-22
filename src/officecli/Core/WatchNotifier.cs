// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.IO.Pipes;
using System.Text;

namespace OfficeCli.Core;

/// <summary>
/// Sends a refresh notification to a running watch process (if any).
/// Non-blocking, fire-and-forget. Silently does nothing if no watch is running.
/// </summary>
public static class WatchNotifier
{
    public static void NotifyIfWatching(string filePath, string? changedPath = null)
    {
        try
        {
            var pipeName = WatchServer.GetWatchPipeName(filePath);
            using var client = new NamedPipeClientStream(".", pipeName, PipeDirection.InOut);
            client.Connect(100); // 100ms timeout — fast fail if no watch

            using var writer = new StreamWriter(client, Encoding.UTF8, leaveOpen: true) { AutoFlush = true };
            using var reader = new StreamReader(client, Encoding.UTF8, leaveOpen: true);

            // Send: "refresh" or "refresh:/slide[1]/shape[2]"
            var message = string.IsNullOrEmpty(changedPath) ? "refresh" : $"refresh:{changedPath}";
            writer.WriteLine(message);
            reader.ReadLine(); // wait for ack
        }
        catch
        {
            // No watch process running — silently ignore
        }
    }

    /// <summary>
    /// Send a close command to a running watch process.
    /// Returns true if the watch was successfully closed.
    /// </summary>
    public static bool SendClose(string filePath)
    {
        try
        {
            var pipeName = WatchServer.GetWatchPipeName(filePath);
            using var client = new NamedPipeClientStream(".", pipeName, PipeDirection.InOut);
            client.Connect(200);

            using var writer = new StreamWriter(client, Encoding.UTF8, leaveOpen: true) { AutoFlush = true };
            using var reader = new StreamReader(client, Encoding.UTF8, leaveOpen: true);

            writer.WriteLine("close");
            reader.ReadLine();
            return true;
        }
        catch
        {
            return false;
        }
    }
}
