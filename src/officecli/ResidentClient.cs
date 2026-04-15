// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.IO.Pipes;
using System.Text;

namespace OfficeCli;

public static class ResidentClient
{
    /// <summary>
    /// Check if a resident is running for this file (without consuming a connection).
    /// Just tries to connect briefly.
    /// </summary>
    public static bool TryConnect(string filePath, out string pipeName)
    {
        pipeName = ResidentServer.GetPipeName(filePath);
        try
        {
            using var client = new NamedPipeClientStream(".", pipeName + "-ping", PipeDirection.InOut);
            client.Connect(100); // 100ms timeout

            // Ping to verify it's the right file
            var pingRequest = new ResidentRequest { Command = "__ping__" };
            var json = System.Text.Json.JsonSerializer.Serialize(pingRequest, ResidentJsonContext.Default.ResidentRequest);
            PipeWriteLine(client, json);

            var responseLine = PipeReadLine(client);
            if (responseLine == null) return false;

            var response = System.Text.Json.JsonSerializer.Deserialize<ResidentResponse>(responseLine, ResidentJsonContext.Default.ResidentResponse);
            if (response == null) return false;

            // Stdout contains the file path when responding to ping
            if (string.IsNullOrEmpty(response.Stdout)) return false;
            var residentFilePath = Path.GetFullPath(response.Stdout);
            var requestedFilePath = Path.GetFullPath(filePath);
            return string.Equals(residentFilePath, requestedFilePath, StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Send a command to the resident server in a single connection.
    /// Returns null if no resident is running or the file doesn't match.
    /// </summary>
    /// <param name="connectTimeoutMs">
    /// How long to wait for the server to accept the pipe connection. Default
    /// 100ms suits the "is a resident listening at all?" fast-fail path; when
    /// the caller has already confirmed the resident is alive (e.g. via
    /// <see cref="TryConnect"/>), pass a longer value (seconds) so the command
    /// waits for its turn in the serialized command queue instead of silently
    /// dropping under load.
    /// </param>
    public static ResidentResponse? TrySend(string filePath, ResidentRequest request, int maxRetries = 0, int connectTimeoutMs = 100)
    {
        var pipeName = ResidentServer.GetPipeName(filePath);
        for (int attempt = 0; attempt <= maxRetries; attempt++)
        {
            try
            {
                using var client = new NamedPipeClientStream(".", pipeName, PipeDirection.InOut);
                client.Connect(connectTimeoutMs);

                var json = System.Text.Json.JsonSerializer.Serialize(request, ResidentJsonContext.Default.ResidentRequest);
                PipeWriteLine(client, json);

                var responseLine = PipeReadLine(client);
                if (responseLine == null) continue;

                var response = System.Text.Json.JsonSerializer.Deserialize<ResidentResponse>(responseLine, ResidentJsonContext.Default.ResidentResponse);
                if (response != null) return response;
            }
            catch
            {
                if (attempt == maxRetries) return null;
                Thread.Sleep(50 * (attempt + 1)); // brief backoff before retry
            }
        }
        return null;
    }

    /// <summary>
    /// Ask a running resident to change its idle timeout. Used by `open`
    /// to upgrade a short-lived resident that `create` auto-started
    /// (60s) up to the normal 12min interactive window. Served by the
    /// ping pipe, so it succeeds even while the main pipe is busy.
    /// Returns false if the resident isn't running, the value is out
    /// of range, or the RPC failed.
    /// </summary>
    public static bool SendSetIdleTimeout(string filePath, int seconds)
    {
        var pipeName = ResidentServer.GetPipeName(filePath) + "-ping";
        try
        {
            using var client = new NamedPipeClientStream(".", pipeName, PipeDirection.InOut);
            client.Connect(200);

            var request = new ResidentRequest { Command = "__set-idle-timeout__" };
            request.Args["seconds"] = seconds.ToString();
            var json = System.Text.Json.JsonSerializer.Serialize(request, ResidentJsonContext.Default.ResidentRequest);
            PipeWriteLine(client, json);

            var responseLine = PipeReadLine(client);
            if (responseLine == null) return false;

            var response = System.Text.Json.JsonSerializer.Deserialize<ResidentResponse>(responseLine, ResidentJsonContext.Default.ResidentResponse);
            return response != null && response.ExitCode == 0;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Send a close command to the resident server.
    /// </summary>
    public static bool SendClose(string filePath)
    {
        // Send close via the dedicated ping pipe (always responsive)
        var pipeName = ResidentServer.GetPipeName(filePath) + "-ping";
        try
        {
            using var client = new NamedPipeClientStream(".", pipeName, PipeDirection.InOut);
            client.Connect(200);

            var request = new ResidentRequest { Command = "__close__" };
            var json = System.Text.Json.JsonSerializer.Serialize(request, ResidentJsonContext.Default.ResidentRequest);
            PipeWriteLine(client, json);

            var responseLine = PipeReadLine(client);
            if (responseLine == null) return false;

            var response = System.Text.Json.JsonSerializer.Deserialize<ResidentResponse>(responseLine, ResidentJsonContext.Default.ResidentResponse);
            return response != null && response.ExitCode == 0;
        }
        catch
        {
            return false;
        }
    }

    // ==================== Pipe I/O helpers ====================
    //
    // On Windows, StreamReader/StreamWriter deadlock on named pipes under .NET 11
    // preview — the managed stream wrapper's internal buffering stalls reads even
    // when bytes are available on the wire.  Raw byte I/O avoids the issue.
    //
    // On Linux/macOS, StreamReader/StreamWriter work fine and are faster (buffered
    // reads), so we keep using them.

    private const int MaxLineLength = 1_048_576; // 1 MB safety limit

    private static void PipeWriteLine(Stream pipe, string line)
    {
        if (!OperatingSystem.IsWindows())
        {
            using var writer = new StreamWriter(pipe, Encoding.UTF8, leaveOpen: true) { AutoFlush = true };
            writer.WriteLine(line);
            return;
        }
        var bytes = Encoding.UTF8.GetBytes(line + "\n");
        pipe.Write(bytes, 0, bytes.Length);
        pipe.Flush();
    }

    private static string? PipeReadLine(Stream pipe)
    {
        if (!OperatingSystem.IsWindows())
        {
            using var reader = new StreamReader(pipe, Encoding.UTF8, leaveOpen: true);
            return reader.ReadLine();
        }
        var buffer = new byte[1];
        var lineBytes = new List<byte>(256);
        while (true)
        {
            var bytesRead = pipe.Read(buffer, 0, 1);
            if (bytesRead == 0) return lineBytes.Count > 0 ? Encoding.UTF8.GetString(lineBytes.ToArray()) : null;
            if (buffer[0] == (byte)'\n')
            {
                if (lineBytes.Count > 0 && lineBytes[^1] == (byte)'\r')
                    lineBytes.RemoveAt(lineBytes.Count - 1);
                return Encoding.UTF8.GetString(lineBytes.ToArray());
            }
            if (lineBytes.Count >= MaxLineLength)
                return null;
            lineBytes.Add(buffer[0]);
        }
    }
}
