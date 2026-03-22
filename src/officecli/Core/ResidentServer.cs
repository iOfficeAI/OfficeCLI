// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.IO.Pipes;
using System.Security.Cryptography;
using System.Text;

namespace OfficeCli.Core;

public class ResidentServer : IDisposable
{
    private readonly IDocumentHandler _handler;
    private readonly string _filePath;
    private readonly string _pipeName;
    private CancellationTokenSource _cts = new();
    private readonly SemaphoreSlim _commandLock = new(1, 1);
    private readonly TimeSpan _idleTimeout = TimeSpan.FromMinutes(12);
    private CancellationTokenSource _idleCts = new();
    private bool _disposed;

    public string PipeName => _pipeName;

    public ResidentServer(string filePath, bool editable = true)
    {
        _filePath = Path.GetFullPath(filePath);
        _pipeName = GetPipeName(_filePath);
        _handler = DocumentHandlerFactory.Open(_filePath, editable);
    }

    public static string GetPipeName(string filePath)
    {
        var fullPath = Path.GetFullPath(filePath);
        if (OperatingSystem.IsWindows())
            fullPath = fullPath.ToUpperInvariant();
        var hash = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(fullPath)))[..16];
        return $"officecli-{hash}";
    }

    public async Task RunAsync(CancellationToken externalToken = default)
    {
        using var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(_cts.Token, externalToken);
        var token = linkedCts.Token;

        // Start ping responder on a dedicated pipe (never blocked by business commands)
        var pingTask = RunPingResponderAsync(token);

        // Start idle watchdog
        var idleTask = RunIdleWatchdogAsync(token);

        // Main command loop - accept connections concurrently, serialize command execution
        while (!token.IsCancellationRequested)
        {
            var server = new NamedPipeServerStream(_pipeName, PipeDirection.InOut,
                NamedPipeServerStream.MaxAllowedServerInstances,
                PipeTransmissionMode.Byte, PipeOptions.Asynchronous);
            try
            {
                await server.WaitForConnectionAsync(token);
                // Handle client asynchronously so we can accept the next connection
                _ = HandleClientWithLockAsync(server, token);
            }
            catch (OperationCanceledException)
            {
                await server.DisposeAsync();
                break;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Resident error: {ex.Message}");
                await server.DisposeAsync();
            }
        }

        // Both tasks observe the same token; swallow cancellation on shutdown
        try { await pingTask; } catch (OperationCanceledException) { }
        try { await idleTask; } catch (OperationCanceledException) { }
    }

    private void ResetIdleTimer()
    {
        // Cancel the old idle CTS to restart the delay; don't Dispose here to avoid
        // a race with RunIdleWatchdogAsync reading the token concurrently.
        var oldCts = Interlocked.Exchange(ref _idleCts, new CancellationTokenSource());
        oldCts.Cancel();
    }

    private async Task RunIdleWatchdogAsync(CancellationToken token)
    {
        while (!token.IsCancellationRequested)
        {
            try
            {
                // Snapshot the current idle CTS; ResetIdleTimer() swaps it to restart the wait
                var idleCts = Volatile.Read(ref _idleCts);
                using var linked = CancellationTokenSource.CreateLinkedTokenSource(idleCts.Token, token);
                await Task.Delay(_idleTimeout, linked.Token);

                // Reached here = idle timeout elapsed without reset
                Console.Error.WriteLine($"Resident idle for {_idleTimeout.TotalMinutes} minutes, closing.");
                _cts.Cancel();
                break;
            }
            catch (OperationCanceledException) when (!token.IsCancellationRequested)
            {
                // _idleCts was cancelled (timer reset), loop and wait again
            }
        }
    }

    private async Task RunPingResponderAsync(CancellationToken token)
    {
        var pingPipeName = _pipeName + "-ping";
        while (!token.IsCancellationRequested)
        {
            var server = new NamedPipeServerStream(pingPipeName, PipeDirection.InOut,
                NamedPipeServerStream.MaxAllowedServerInstances,
                PipeTransmissionMode.Byte, PipeOptions.Asynchronous);
            try
            {
                await server.WaitForConnectionAsync(token);
                using var reader = new StreamReader(server, Encoding.UTF8, leaveOpen: true);
                using var writer = new StreamWriter(server, Encoding.UTF8, leaveOpen: true) { AutoFlush = true };

                var requestLine = await reader.ReadLineAsync(token);
                if (requestLine != null)
                {
                    var request = System.Text.Json.JsonSerializer.Deserialize<ResidentRequest>(requestLine, ResidentJsonContext.Default.ResidentRequest);
                    if (request?.Command == "__ping__")
                    {
                        var response = MakeResponse(0, _filePath, "");
                        await writer.WriteLineAsync(response.AsMemory(), token);
                    }
                    else if (request?.Command == "__close__")
                    {
                        var response = MakeResponse(0, "Closing resident.", "");
                        await writer.WriteLineAsync(response.AsMemory(), token);
                        _cts.Cancel();
                        break;
                    }
                }
            }
            catch (OperationCanceledException)
            {
                break;
            }
            catch
            {
                // Ignore ping errors
            }
            finally
            {
                await server.DisposeAsync();
            }
        }
    }

    private async Task HandleClientWithLockAsync(NamedPipeServerStream server, CancellationToken token)
    {
        try
        {
            await _commandLock.WaitAsync(token);
            try
            {
                ResetIdleTimer();
                await HandleClientAsync(server, token);
            }
            finally
            {
                _commandLock.Release();
                ResetIdleTimer();
            }
        }
        catch (OperationCanceledException) { }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Resident error: {ex.Message}");
        }
        finally
        {
            await server.DisposeAsync();
        }
    }

    private async Task HandleClientAsync(NamedPipeServerStream server, CancellationToken token)
    {
        using var reader = new StreamReader(server, Encoding.UTF8, leaveOpen: true);
        using var writer = new StreamWriter(server, Encoding.UTF8, leaveOpen: true) { AutoFlush = true };

        var requestLine = await reader.ReadLineAsync(token);
        if (requestLine == null) return;

        var response = ProcessRequest(requestLine);
        await writer.WriteLineAsync(response.AsMemory(), token);
    }

    private string ProcessRequest(string requestLine)
    {
        ResidentRequest? request = null;
        try
        {
            request = System.Text.Json.JsonSerializer.Deserialize<ResidentRequest>(requestLine, ResidentJsonContext.Default.ResidentRequest);
            if (request == null)
                return MakeResponse(1, "", "Invalid request");

            // Capture stdout/stderr (safe: _commandLock serializes all commands)
            var stdoutWriter = new StringWriter();
            var stderrWriter = new StringWriter();
            var origOut = Console.Out;
            var origErr = Console.Error;
            Console.SetOut(stdoutWriter);
            Console.SetError(stderrWriter);

            try
            {
                ExecuteCommand(request);
            }
            finally
            {
                Console.SetOut(origOut);
                Console.SetError(origErr);
            }

            var stdout = stdoutWriter.ToString().TrimEnd('\r', '\n');
            var stderr = stderrWriter.ToString().TrimEnd('\r', '\n');

            if (request.Json)
            {
                // JSON mode: server builds the envelope so client just passes through
                var warnings = BuildWarnings(stderr);
                var envelope = IsJson(stdout)
                    ? OutputFormatter.WrapEnvelope(stdout, warnings)
                    : OutputFormatter.WrapEnvelopeText(stdout, warnings);
                return MakeResponse(0, envelope, "");
            }

            return MakeResponse(0, stdout, stderr);
        }
        catch (Exception ex)
        {
            if (request?.Json == true)
            {
                // JSON mode: wrap error in envelope
                return MakeResponse(1, OutputFormatter.WrapErrorEnvelope(ex), "");
            }
            return MakeResponse(1, "", ex.Message);
        }
    }

    private static bool IsJson(string s)
    {
        var trimmed = s.AsSpan().TrimStart();
        return trimmed.Length > 0 && (trimmed[0] == '{' || trimmed[0] == '[');
    }

    private static List<CliWarning>? BuildWarnings(string stderr)
    {
        if (string.IsNullOrEmpty(stderr)) return null;
        var lines = stderr.Split('\n', StringSplitOptions.RemoveEmptyEntries);
        if (lines.Length == 0) return null;
        return lines.Select(line =>
        {
            var warning = new CliWarning { Message = line.Trim() };
            if (line.Contains("UNSUPPORTED")) warning.Code = "unsupported_property";
            else if (line.Contains("VALIDATION")) warning.Code = "validation_error";
            else warning.Code = "warning";
            return warning;
        }).ToList();
    }

    private void ExecuteCommand(ResidentRequest request)
    {
        var format = request.Json ? OutputFormat.Json : OutputFormat.Text;

        switch (request.Command)
        {
            case "view":
                ExecuteView(request, format);
                break;
            case "get":
                ExecuteGet(request, format);
                break;
            case "query":
                ExecuteQuery(request, format);
                break;
            case "set":
                ExecuteSet(request);
                WatchNotifier.NotifyIfWatching(_filePath, request.GetArg("path"));
                break;
            case "add":
                ExecuteAdd(request);
                WatchNotifier.NotifyIfWatching(_filePath, request.GetArg("parent"));
                break;
            case "remove":
                ExecuteRemove(request);
                WatchNotifier.NotifyIfWatching(_filePath, request.GetArg("path"));
                break;
            case "move":
                ExecuteMove(request);
                WatchNotifier.NotifyIfWatching(_filePath, request.GetArg("path"));
                break;
            case "raw":
                ExecuteRaw(request);
                break;
            case "raw-set":
                ExecuteRawSet(request);
                WatchNotifier.NotifyIfWatching(_filePath);
                break;
            case "add-part":
                ExecuteAddPart(request);
                WatchNotifier.NotifyIfWatching(_filePath);
                break;
            case "validate":
                ExecuteValidate();
                break;
            default:
                Console.Error.WriteLine($"Unknown command: {request.Command}");
                break;
        }
    }

    private void ExecuteView(ResidentRequest req, OutputFormat format)
    {
        var mode = req.GetArg("mode", "text")!;
        var start = req.GetIntArg("start");
        var end = req.GetIntArg("end");
        var maxLines = req.GetIntArg("max-lines");
        var issueType = req.GetArgOrNull("type");
        var limit = req.GetIntArg("limit");
        var cols = req.GetCols("cols");

        if (mode!.ToLowerInvariant() is "html" or "h")
        {
            if (_handler is OfficeCli.Handlers.PowerPointHandler pptHandler)
            {
                var html = pptHandler.ViewAsHtml(start, end);

                if (req.Json)
                {
                    Console.Write(html);
                }
                else
                {
                    var htmlPath = Path.Combine(Path.GetTempPath(), $"officecli_preview_{Path.GetFileNameWithoutExtension(_filePath)}_{DateTime.Now:HHmmss}.html");
                    File.WriteAllText(htmlPath, html);
                    Console.WriteLine(htmlPath);
                    try
                    {
                        var psi = new System.Diagnostics.ProcessStartInfo(htmlPath) { UseShellExecute = true };
                        System.Diagnostics.Process.Start(psi);
                    }
                    catch { /* silently ignore if browser can't be opened */ }
                }
            }
            else
            {
                Console.Error.WriteLine("HTML preview is only supported for .pptx files.");
            }
            return;
        }

        var output = mode!.ToLowerInvariant() switch
        {
            "text" or "t" => _handler.ViewAsText(start, end, maxLines, cols),
            "annotated" or "a" => _handler.ViewAsAnnotated(start, end, maxLines, cols),
            "outline" or "o" => _handler.ViewAsOutline(),
            "stats" or "s" => _handler.ViewAsStats(),
            "issues" or "i" => OutputFormatter.FormatIssues(_handler.ViewAsIssues(issueType, limit), format),
            _ => $"Unknown mode: {mode}. Available: text, annotated, outline, stats, issues, html"
        };

        if (req.Json && mode is not ("issues" or "i"))
            Console.WriteLine(OutputFormatter.FormatView(mode, output, format));
        else
            Console.WriteLine(output);
    }

    private void ExecuteGet(ResidentRequest req, OutputFormat format)
    {
        var path = req.GetArg("path", "/");
        var depth = req.GetIntArg("depth") ?? 1;
        var node = _handler.Get(path, depth);
        Console.WriteLine(OutputFormatter.FormatNode(node, format));
    }

    private void ExecuteQuery(ResidentRequest req, OutputFormat format)
    {
        var selector = req.GetArg("selector", "");
        var filters = AttributeFilter.Parse(selector);
        var (results, warnings) = AttributeFilter.ApplyWithWarnings(_handler.Query(selector), filters);
        foreach (var w in warnings) Console.Error.WriteLine(w);
        Console.WriteLine(OutputFormatter.FormatNodes(results, format));
    }

    private void ExecuteSet(ResidentRequest req)
    {
        var path = req.GetArg("path", "/");
        var properties = req.GetProps();
        var unsupported = _handler.Set(path, properties);
        var applied = properties.Where(kv => !unsupported.Contains(kv.Key)).ToList();
        if (applied.Count > 0)
            Console.WriteLine($"Updated {path}: {string.Join(", ", applied.Select(kv => $"{kv.Key}={kv.Value}"))}");
        if (unsupported.Count > 0)
            Console.Error.WriteLine($"UNSUPPORTED props (use raw-set instead): {string.Join(", ", unsupported)}");
    }

    private void ExecuteAdd(ResidentRequest req)
    {
        var parentPath = req.GetArg("parent", "/body");
        var from = req.GetArgOrNull("from");
        var index = req.GetIntArg("index");

        if (!string.IsNullOrEmpty(from))
        {
            var resultPath = _handler.CopyFrom(from, parentPath, index);
            Console.WriteLine($"Copied to {resultPath}");
        }
        else
        {
            var type = req.GetArg("type", "");
            var properties = req.GetProps();
            var resultPath = _handler.Add(parentPath, type, index, properties);
            Console.WriteLine($"Added {type} at {resultPath}");
        }
    }

    private void ExecuteRemove(ResidentRequest req)
    {
        var path = req.GetArg("path", "/");
        _handler.Remove(path);
        Console.WriteLine($"Removed {path}");
    }

    private void ExecuteMove(ResidentRequest req)
    {
        var path = req.GetArg("path", "/");
        var to = req.GetArgOrNull("to");
        var index = req.GetIntArg("index");
        var resultPath = _handler.Move(path, to, index);
        Console.WriteLine($"Moved to {resultPath}");
    }

    private void ExecuteRaw(ResidentRequest req)
    {
        var partPath = req.GetArg("part", "/document");
        var startRow = req.GetIntArg("start");
        var endRow = req.GetIntArg("end");
        var cols = req.GetCols("cols");
        Console.WriteLine(_handler.Raw(partPath, startRow, endRow, cols));
    }

    private void ExecuteRawSet(ResidentRequest req)
    {
        var partPath = req.GetArg("part", "/document");
        var xpath = req.GetArg("xpath", "");
        var action = req.GetArg("action", "");
        var xml = req.GetArgOrNull("xml");

        var errorsBefore = _handler.Validate().Select(e => e.Description).ToHashSet();
        _handler.RawSet(partPath, xpath, action, xml);

        var errorsAfter = _handler.Validate();
        var newErrors = errorsAfter.Where(e => !errorsBefore.Contains(e.Description)).ToList();
        if (newErrors.Count > 0)
        {
            Console.WriteLine($"VALIDATION: {newErrors.Count} new error(s) introduced:");
            foreach (var err in newErrors)
            {
                Console.WriteLine($"  [{err.ErrorType}] {err.Description}");
                if (err.Path != null) Console.WriteLine($"    Path: {err.Path}");
                if (err.Part != null) Console.WriteLine($"    Part: {err.Part}");
            }
        }
    }

    private void ExecuteAddPart(ResidentRequest req)
    {
        var parent = req.GetArg("parent", "/");
        var type = req.GetArg("type", "");
        var errorsBefore = _handler.Validate().Select(e => e.Description).ToHashSet();
        var (relId, partPath) = _handler.AddPart(parent, type);
        Console.WriteLine($"Created {type} part: relId={relId} path={partPath}");

        var errorsAfter = _handler.Validate();
        var newErrors = errorsAfter.Where(e => !errorsBefore.Contains(e.Description)).ToList();
        if (newErrors.Count > 0)
        {
            Console.WriteLine($"VALIDATION: {newErrors.Count} new error(s) introduced:");
            foreach (var err in newErrors)
            {
                Console.WriteLine($"  [{err.ErrorType}] {err.Description}");
                if (err.Path != null) Console.WriteLine($"    Path: {err.Path}");
                if (err.Part != null) Console.WriteLine($"    Part: {err.Part}");
            }
        }
    }

    private void ExecuteValidate()
    {
        var errors = _handler.Validate();
        if (errors.Count == 0)
        {
            Console.WriteLine("Validation passed: no errors found.");
        }
        else
        {
            Console.WriteLine($"Found {errors.Count} validation error(s):");
            foreach (var err in errors)
            {
                Console.WriteLine($"  [{err.ErrorType}] {err.Description}");
                if (err.Path != null) Console.WriteLine($"    Path: {err.Path}");
                if (err.Part != null) Console.WriteLine($"    Part: {err.Part}");
            }
        }
    }

    private static string MakeResponse(int exitCode, string stdout, string stderr)
    {
        var response = new ResidentResponse { ExitCode = exitCode, Stdout = stdout, Stderr = stderr };
        return System.Text.Json.JsonSerializer.Serialize(response, ResidentJsonContext.Default.ResidentResponse);
    }

    public void Dispose()
    {
        if (!_disposed)
        {
            _disposed = true;
            _cts.Cancel();

            // Run the entire shutdown sequence on a background thread.
            // A watchdog on the calling thread ensures the process always exits.
            var shutdownTask = Task.Run(() =>
            {
                // Wait for any in-flight command to finish (preserves data integrity)
                _commandLock.Wait();
                _commandLock.Release();

                try { _handler.Dispose(); }
                catch (Exception ex) { Console.Error.WriteLine($"Warning: handler dispose error: {ex.Message}"); }

                _commandLock.Dispose();
            });

            // Watchdog: if shutdown takes longer than 10min, force exit
            if (!shutdownTask.Wait(TimeSpan.FromMinutes(10)))
            {
                Console.Error.WriteLine("Warning: shutdown timed out after 10 minutes, forcing exit.");
                Environment.Exit(1);
            }

            _cts.Dispose();
            _idleCts.Dispose();
        }
    }
}

public class ResidentRequest
{
    public string Command { get; set; } = "";
    public Dictionary<string, string> Args { get; set; } = new();
    public string[]? Props { get; set; }
    public bool Json { get; set; }

    public string GetArg(string key, string defaultValue = "")
    {
        return Args.TryGetValue(key, out var val) ? val : defaultValue;
    }

    public string? GetArgOrNull(string key)
    {
        return Args.TryGetValue(key, out var val) ? val : null;
    }

    public int? GetIntArg(string key)
    {
        if (Args.TryGetValue(key, out var val) && int.TryParse(val, out var n))
            return n;
        return null;
    }

    public HashSet<string>? GetCols(string key)
    {
        var val = GetArgOrNull(key);
        if (val == null) return null;
        return new HashSet<string>(val.Split(',').Select(c => c.Trim().ToUpperInvariant()));
    }

    public Dictionary<string, string> GetProps()
    {
        var properties = new Dictionary<string, string>();
        foreach (var prop in Props ?? Array.Empty<string>())
        {
            var eqIdx = prop.IndexOf('=');
            if (eqIdx > 0)
                properties[prop[..eqIdx]] = prop[(eqIdx + 1)..];
        }
        return properties;
    }
}

public class ResidentResponse
{
    public int ExitCode { get; set; }
    public string Stdout { get; set; } = "";
    public string Stderr { get; set; } = "";
}

[System.Text.Json.Serialization.JsonSourceGenerationOptions]
[System.Text.Json.Serialization.JsonSerializable(typeof(ResidentRequest))]
[System.Text.Json.Serialization.JsonSerializable(typeof(ResidentResponse))]
internal partial class ResidentJsonContext : System.Text.Json.Serialization.JsonSerializerContext;
