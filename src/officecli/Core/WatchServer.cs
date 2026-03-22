// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeCli.Core;

public class WatchServer : IDisposable
{
    private readonly string _filePath;
    private readonly string _pipeName;
    private readonly int _port;
    private readonly TcpListener _tcpListener;
    private readonly List<NetworkStream> _sseClients = new();
    private readonly object _sseLock = new();
    private CancellationTokenSource _cts = new();
    private string _currentHtml = "";
    private bool _disposed;

    private const string SseScript = """
        <script>
        (function() {
            var es = new EventSource('/events');
            es.addEventListener('update', function(e) {
                var msg = JSON.parse(e.data);
                if (msg.action === 'full') {
                    location.reload();
                    return;
                }
                var slideNum = msg.slide;
                if (msg.action === 'replace') {
                    var el = document.querySelector('.slide-container[data-slide="' + slideNum + '"]');
                    if (el) {
                        var tmp = document.createElement('div');
                        tmp.innerHTML = msg.html;
                        var newEl = tmp.firstElementChild;
                        el.parentNode.replaceChild(newEl, el);
                        // re-scale the new slide
                        if (typeof scaleSlides === 'function') scaleSlides();
                        if (typeof buildThumbs === 'function') buildThumbs();
                    } else {
                        location.reload();
                    }
                } else if (msg.action === 'remove') {
                    var el = document.querySelector('.slide-container[data-slide="' + slideNum + '"]');
                    if (el) {
                        el.remove();
                        // renumber remaining slides
                        document.querySelectorAll('.slide-container').forEach(function(c, i) {
                            c.setAttribute('data-slide', i + 1);
                        });
                        if (typeof buildThumbs === 'function') buildThumbs();
                    }
                    // Update page counter
                    var counter = document.querySelector('.page-counter');
                    if (counter) {
                        var total = document.querySelectorAll('.slide-container').length;
                        counter.textContent = '1 / ' + total;
                    }
                } else if (msg.action === 'add') {
                    var main = document.querySelector('.main');
                    if (main) {
                        var tmp = document.createElement('div');
                        tmp.innerHTML = msg.html;
                        var newEl = tmp.firstElementChild;
                        main.appendChild(newEl);
                        if (typeof scaleSlides === 'function') scaleSlides();
                        if (typeof buildThumbs === 'function') buildThumbs();
                    }
                    var counter = document.querySelector('.page-counter');
                    if (counter) {
                        var total = document.querySelectorAll('.slide-container').length;
                        counter.textContent = '1 / ' + total;
                    }
                }
            });
        })();
        </script>
        """;

    public WatchServer(string filePath, int port)
    {
        _filePath = Path.GetFullPath(filePath);
        _pipeName = GetWatchPipeName(_filePath);
        _port = port;
        _tcpListener = new TcpListener(IPAddress.Loopback, _port);
    }

    public static string GetWatchPipeName(string filePath)
    {
        var fullPath = Path.GetFullPath(filePath);
        if (OperatingSystem.IsWindows())
            fullPath = fullPath.ToUpperInvariant();
        var hash = Convert.ToHexString(
            System.Security.Cryptography.SHA256.HashData(Encoding.UTF8.GetBytes(fullPath)))[..16];
        return $"officecli-watch-{hash}";
    }

    public async Task RunAsync(CancellationToken externalToken = default)
    {
        using var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(_cts.Token, externalToken);
        var token = linkedCts.Token;

        RefreshFullHtml();

        _tcpListener.Start();
        Console.WriteLine($"Watch: http://localhost:{_port}");
        Console.WriteLine($"Watching: {_filePath}");
        Console.WriteLine("Press Ctrl+C to stop.");

        var pipeTask = RunPipeListenerAsync(token);

        while (!token.IsCancellationRequested)
        {
            try
            {
                var client = await _tcpListener.AcceptTcpClientAsync(token);
                _ = HandleClientAsync(client, token);
            }
            catch (OperationCanceledException) { break; }
            catch (SocketException) { break; }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Watch HTTP error: {ex.Message}");
            }
        }

        try { await pipeTask; } catch (OperationCanceledException) { }
    }

    private void RefreshFullHtml()
    {
        try
        {
            var response = ResidentClient.TrySend(_filePath, new ResidentRequest
            {
                Command = "view",
                Args = new Dictionary<string, string> { ["mode"] = "html" },
                Json = true
            });

            if (response != null && response.ExitCode == 0 && !string.IsNullOrEmpty(response.Stdout))
            {
                _currentHtml = response.Stdout;
                return;
            }

            using var handler = DocumentHandlerFactory.Open(_filePath);
            if (handler is OfficeCli.Handlers.PowerPointHandler pptHandler)
            {
                _currentHtml = pptHandler.ViewAsHtml();
            }
            else
            {
                _currentHtml = "<html><body><p>HTML preview is only supported for .pptx files.</p></body></html>";
            }
        }
        catch (Exception ex)
        {
            _currentHtml = $"<html><body><p>Error: {ex.Message}</p></body></html>";
        }
    }

    /// <summary>
    /// Extract slide number from a path like /slide[2]/shape[1] → 2.
    /// Returns 0 if path is "/" (root-level add like adding a slide).
    /// </summary>
    private static int ExtractSlideNum(string? path)
    {
        if (string.IsNullOrEmpty(path)) return 0;
        var match = Regex.Match(path, @"/slide\[(\d+)\]");
        if (match.Success && int.TryParse(match.Groups[1].Value, out var num))
            return num;
        return 0; // root-level operation
    }

    /// <summary>
    /// Render a single slide HTML using resident or direct file access.
    /// </summary>
    private string? RenderSlideHtml(int slideNum)
    {
        try
        {
            // Try resident: the handler stays in memory
            var response = ResidentClient.TrySend(_filePath, new ResidentRequest
            {
                Command = "view",
                Args = new Dictionary<string, string> { ["mode"] = "html" },
                Json = true
            });

            if (response != null && response.ExitCode == 0 && !string.IsNullOrEmpty(response.Stdout))
            {
                // Resident returned full HTML; we need to update _currentHtml and extract the slide
                _currentHtml = response.Stdout;
            }
            else
            {
                using var handler = DocumentHandlerFactory.Open(_filePath);
                if (handler is OfficeCli.Handlers.PowerPointHandler pptHandler)
                {
                    var slideHtml = pptHandler.RenderSlideHtml(slideNum);
                    if (slideHtml != null) return slideHtml;
                    _currentHtml = pptHandler.ViewAsHtml();
                }
            }

            // Extract slide fragment from full HTML
            return ExtractSlideFragment(_currentHtml, slideNum);
        }
        catch
        {
            return null;
        }
    }

    private static string? ExtractSlideFragment(string fullHtml, int slideNum)
    {
        var marker = $"data-slide=\"{slideNum}\"";
        var idx = fullHtml.IndexOf(marker, StringComparison.Ordinal);
        if (idx < 0) return null;

        // Find the opening <div that contains this marker
        var start = fullHtml.LastIndexOf("<div ", idx, StringComparison.Ordinal);
        if (start < 0) return null;

        // Find matching closing </div> by counting nesting
        var depth = 0;
        var pos = start;
        while (pos < fullHtml.Length)
        {
            var nextOpen = fullHtml.IndexOf("<div", pos, StringComparison.OrdinalIgnoreCase);
            var nextClose = fullHtml.IndexOf("</div>", pos, StringComparison.OrdinalIgnoreCase);

            if (nextClose < 0) break;

            if (nextOpen >= 0 && nextOpen < nextClose)
            {
                depth++;
                pos = nextOpen + 4;
            }
            else
            {
                depth--;
                if (depth == 0)
                {
                    return fullHtml[start..(nextClose + 6)];
                }
                pos = nextClose + 6;
            }
        }

        return null;
    }

    private int GetSlideCount()
    {
        try
        {
            var response = ResidentClient.TrySend(_filePath, new ResidentRequest
            {
                Command = "view",
                Args = new Dictionary<string, string> { ["mode"] = "html" },
                Json = true
            });

            if (response != null && response.ExitCode == 0 && !string.IsNullOrEmpty(response.Stdout))
            {
                _currentHtml = response.Stdout;
            }
            else
            {
                using var handler = DocumentHandlerFactory.Open(_filePath);
                if (handler is OfficeCli.Handlers.PowerPointHandler pptHandler)
                    return pptHandler.GetSlideCount();
            }
        }
        catch { }

        // Count from cached HTML
        return Regex.Matches(_currentHtml, @"data-slide=""\d+""").Count;
    }

    private async Task RunPipeListenerAsync(CancellationToken token)
    {
        while (!token.IsCancellationRequested)
        {
            var server = new System.IO.Pipes.NamedPipeServerStream(
                _pipeName, System.IO.Pipes.PipeDirection.InOut,
                System.IO.Pipes.NamedPipeServerStream.MaxAllowedServerInstances,
                System.IO.Pipes.PipeTransmissionMode.Byte,
                System.IO.Pipes.PipeOptions.Asynchronous);
            try
            {
                await server.WaitForConnectionAsync(token);
                using var reader = new StreamReader(server, Encoding.UTF8, leaveOpen: true);
                using var writer = new StreamWriter(server, Encoding.UTF8, leaveOpen: true) { AutoFlush = true };

                var message = await reader.ReadLineAsync(token);
                await writer.WriteLineAsync("ok".AsMemory(), token);

                if (message == "close")
                {
                    Console.WriteLine("Watch closed by remote command.");
                    _cts.Cancel();
                    break;
                }
                else if (message != null && message.StartsWith("refresh"))
                {
                    string? changedPath = null;
                    if (message.Contains(':'))
                        changedPath = message[(message.IndexOf(':') + 1)..];

                    HandleChange(changedPath);
                }
            }
            catch (OperationCanceledException) { break; }
            catch { /* ignore pipe errors */ }
            finally
            {
                await server.DisposeAsync();
            }
        }
    }

    private void HandleChange(string? changedPath)
    {
        var slideNum = ExtractSlideNum(changedPath);

        if (slideNum == 0)
        {
            // Root-level change (add/remove slide) — need to figure out what happened
            var oldCount = Regex.Matches(_currentHtml, @"data-slide=""\d+""").Count;
            RefreshFullHtml();
            var newCount = Regex.Matches(_currentHtml, @"data-slide=""\d+""").Count;

            if (newCount > oldCount)
            {
                // Slide added — render the new slide and push "add"
                var slideHtml = ExtractSlideFragment(_currentHtml, newCount);
                if (slideHtml != null)
                {
                    SendSseEvent("add", newCount, slideHtml);
                    return;
                }
            }
            else if (newCount < oldCount)
            {
                // Slide removed — figure out which one
                // For simplicity, find which slide number is missing
                for (int i = 1; i <= oldCount; i++)
                {
                    if (ExtractSlideFragment(_currentHtml, i) == null || i > newCount)
                    {
                        SendSseEvent("remove", i, null);
                        return;
                    }
                }
            }

            // Fallback: full reload
            SendSseEvent("full", 0, null);
        }
        else
        {
            // Slide-level change — render just that slide
            var slideHtml = RenderSlideHtml(slideNum);
            if (slideHtml != null)
            {
                SendSseEvent("replace", slideNum, slideHtml);
                // Also update _currentHtml so new page loads show latest state
                RefreshFullHtml();
            }
            else
            {
                RefreshFullHtml();
                SendSseEvent("full", 0, null);
            }
        }
    }

    private void SendSseEvent(string action, int slideNum, string? html)
    {
        // Build JSON manually to avoid dependency
        var sb = new StringBuilder();
        sb.Append("{\"action\":\"").Append(action).Append('"');
        sb.Append(",\"slide\":").Append(slideNum);
        if (html != null)
        {
            sb.Append(",\"html\":\"");
            // Escape JSON string
            foreach (var ch in html)
            {
                switch (ch)
                {
                    case '"': sb.Append("\\\""); break;
                    case '\\': sb.Append("\\\\"); break;
                    case '\n': sb.Append("\\n"); break;
                    case '\r': sb.Append("\\r"); break;
                    case '\t': sb.Append("\\t"); break;
                    default:
                        if (ch < 0x20)
                            sb.Append($"\\u{(int)ch:X4}");
                        else
                            sb.Append(ch);
                        break;
                }
            }
            sb.Append('"');
        }
        sb.Append('}');

        var json = sb.ToString();

        lock (_sseLock)
        {
            var dead = new List<NetworkStream>();
            foreach (var client in _sseClients)
            {
                try
                {
                    var data = Encoding.UTF8.GetBytes($"event: update\ndata: {json}\n\n");
                    client.Write(data);
                    client.Flush();
                }
                catch
                {
                    dead.Add(client);
                }
            }
            foreach (var d in dead) _sseClients.Remove(d);
        }
    }

    private async Task HandleClientAsync(TcpClient client, CancellationToken token)
    {
        try
        {
            var stream = client.GetStream();
            var requestLine = await ReadHttpRequestAsync(stream, token);

            if (requestLine.Contains("GET /events"))
            {
                await HandleSseAsync(stream, token);
            }
            else
            {
                var html = InjectSseScript(_currentHtml);
                var body = Encoding.UTF8.GetBytes(html);
                var header = Encoding.UTF8.GetBytes(
                    $"HTTP/1.1 200 OK\r\nContent-Type: text/html; charset=utf-8\r\nContent-Length: {body.Length}\r\nConnection: close\r\n\r\n");
                await stream.WriteAsync(header, token);
                await stream.WriteAsync(body, token);
                client.Close();
            }
        }
        catch
        {
            try { client.Close(); } catch { }
        }
    }

    private static async Task<string> ReadHttpRequestAsync(NetworkStream stream, CancellationToken token)
    {
        var buffer = new byte[4096];
        var read = await stream.ReadAsync(buffer, token);
        var request = Encoding.UTF8.GetString(buffer, 0, read);
        var idx = request.IndexOf('\r');
        return idx >= 0 ? request[..idx] : request;
    }

    private async Task HandleSseAsync(NetworkStream stream, CancellationToken token)
    {
        var header = Encoding.UTF8.GetBytes(
            "HTTP/1.1 200 OK\r\nContent-Type: text/event-stream; charset=utf-8\r\nCache-Control: no-cache\r\nConnection: keep-alive\r\nAccess-Control-Allow-Origin: *\r\n\r\n");
        await stream.WriteAsync(header, token);

        lock (_sseLock) { _sseClients.Add(stream); }

        try
        {
            while (!token.IsCancellationRequested)
            {
                await Task.Delay(30000, token);
                var heartbeat = Encoding.UTF8.GetBytes(": heartbeat\n\n");
                await stream.WriteAsync(heartbeat, token);
            }
        }
        catch { }
        finally
        {
            lock (_sseLock) { _sseClients.Remove(stream); }
        }
    }

    private static string InjectSseScript(string html)
    {
        var idx = html.LastIndexOf("</body>", StringComparison.OrdinalIgnoreCase);
        if (idx >= 0)
            return html[..idx] + SseScript + html[idx..];
        return html + SseScript;
    }

    public void Dispose()
    {
        if (!_disposed)
        {
            _disposed = true;
            _cts.Cancel();
            try { _tcpListener.Stop(); } catch { }
            _cts.Dispose();
        }
    }
}
