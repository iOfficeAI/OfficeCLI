// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Net.Sockets;
using System.Net.WebSockets;
using System.Text;
using System.Text.Json;

namespace OfficeCli.Core;

/// <summary>
/// Maintains a warm, in-memory Chrome DevTools Protocol (CDP) WebSocket connection
/// for ultra-fast (~35ms) HTML-to-PNG screenshot rendering within a ResidentServer lifecycle.
/// Supports zero-disk in-memory HTML injection over data URIs.
/// </summary>
public sealed class CdpScreenshotClient : IDisposable
{
    private Process? _chromeProcess;
    private ClientWebSocket? _ws;
    private int _messageId;
    private int _port;
    private readonly object _lock = new();
    private bool _disposed;

    private static int GetFreeTcpPort()
    {
        using var l = new TcpListener(System.Net.IPAddress.Loopback, 0);
        l.Start();
        int port = ((System.Net.IPEndPoint)l.LocalEndpoint).Port;
        l.Stop();
        return port;
    }

    private static string EscapeJsonString(string s)
    {
        return "\"" + s.Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("\n", "\\n").Replace("\r", "\\r") + "\"";
    }

    public async Task EnsureStartedAsync(CancellationToken ct = default)
    {
        if (_ws != null && _ws.State == WebSocketState.Open) return;

        var bin = HtmlScreenshot.FindChrome();
        if (bin == null) throw new InvalidOperationException("No Chromium-family browser found on system.");

        _port = GetFreeTcpPort();
        var psi = new ProcessStartInfo
        {
            FileName = bin,
            Arguments = $"--headless=new --disable-gpu --no-sandbox --hide-scrollbars --remote-debugging-port={_port} --remote-allow-origins=*",
            UseShellExecute = false,
            CreateNoWindow = true,
        };

        _chromeProcess = Process.Start(psi);
        if (_chromeProcess == null) throw new InvalidOperationException("Failed to launch browser process.");

        string? wsUrl = await PollPageWebSocketUrlAsync(_port, ct);
        if (string.IsNullOrEmpty(wsUrl))
        {
            Dispose();
            throw new InvalidOperationException("Failed to connect to CDP HTTP debug endpoint.");
        }

        _ws = new ClientWebSocket();
        await _ws.ConnectAsync(new Uri(wsUrl), ct);

        await SendRawCdpAsync("{\"id\":1,\"method\":\"Page.enable\"}", 1, ct);
    }

    private static async Task<string?> PollPageWebSocketUrlAsync(int port, CancellationToken ct)
    {
        using var http = new HttpClient { Timeout = TimeSpan.FromMilliseconds(500) };
        string targetUrl = $"http://127.0.0.1:{port}/json/list";

        for (int i = 0; i < 30; i++)
        {
            try
            {
                var jsonStr = await http.GetStringAsync(targetUrl, ct);
                using var doc = JsonDocument.Parse(jsonStr);
                foreach (var el in doc.RootElement.EnumerateArray())
                {
                    if (el.TryGetProperty("type", out var typeProp) && typeProp.GetString() == "page")
                    {
                        if (el.TryGetProperty("webSocketDebuggerUrl", out var wsProp))
                            return wsProp.GetString();
                    }
                }
            }
            catch
            {
                await Task.Delay(50, ct);
            }
        }
        return null;
    }

    private async Task<JsonElement> SendRawCdpAsync(string reqJson, int msgId, CancellationToken ct)
    {
        if (_ws == null || _ws.State != WebSocketState.Open)
            await EnsureStartedAsync(ct);

        byte[] reqBytes = Encoding.UTF8.GetBytes(reqJson);
        await _ws!.SendAsync(new ArraySegment<byte>(reqBytes), WebSocketMessageType.Text, true, ct);

        var buffer = new byte[8192];
        using var ms = new MemoryStream();

        while (true)
        {
            WebSocketReceiveResult result;
            do
            {
                result = await _ws.ReceiveAsync(new ArraySegment<byte>(buffer), ct);
                if (result.MessageType == WebSocketMessageType.Close)
                {
                    await _ws.CloseAsync(WebSocketCloseStatus.NormalClosure, "Closing", ct);
                    throw new InvalidOperationException("WebSocket closed unexpectedly.");
                }
                ms.Write(buffer, 0, result.Count);
            }
            while (!result.EndOfMessage);

            ms.Position = 0;
            using var doc = JsonDocument.Parse(ms);
            var root = doc.RootElement.Clone();
            ms.SetLength(0);

            if (root.TryGetProperty("id", out var idProp) && idProp.GetInt32() == msgId)
            {
                if (root.TryGetProperty("result", out var resProp))
                    return resProp;
                if (root.TryGetProperty("error", out var errProp))
                    throw new InvalidOperationException($"CDP Error: {errProp.GetRawText()}");
            }
        }
    }

    private async Task WaitForLoadEventAsync(CancellationToken ct)
    {
        if (_ws == null || _ws.State != WebSocketState.Open) return;

        var buffer = new byte[4096];
        using var ms = new MemoryStream();

        using var cts = CancellationTokenSource.CreateLinkedTokenSource(ct);
        cts.CancelAfter(3000); // 3-second safety timeout for load event

        try
        {
            while (!cts.IsCancellationRequested)
            {
                WebSocketReceiveResult result;
                do
                {
                    result = await _ws.ReceiveAsync(new ArraySegment<byte>(buffer), cts.Token);
                    ms.Write(buffer, 0, result.Count);
                }
                while (!result.EndOfMessage);

                ms.Position = 0;
                using var doc = JsonDocument.Parse(ms);
                var root = doc.RootElement.Clone();
                ms.SetLength(0);

                if (root.TryGetProperty("method", out var methodProp) && methodProp.GetString() == "Page.loadEventFired")
                {
                    return;
                }
            }
        }
        catch (OperationCanceledException) { }
    }

    public async Task<byte[]?> CaptureScreenshotAsync(string inputHtmlOrPath, int w = 1600, int h = 1200, bool isRawHtml = false, CancellationToken ct = default)
    {
        try
        {
            await EnsureStartedAsync(ct);

            int id1 = Interlocked.Increment(ref _messageId);
            string metricsJson = $"{{\"id\":{id1},\"method\":\"Emulation.setDeviceMetricsOverride\",\"params\":{{\"width\":{w},\"height\":{h},\"deviceScaleFactor\":1,\"mobile\":false}}}}";
            await SendRawCdpAsync(metricsJson, id1, ct);

            string dataUrl;
            if (isRawHtml)
            {
                string b64Html = Convert.ToBase64String(Encoding.UTF8.GetBytes(inputHtmlOrPath));
                dataUrl = "data:text/html;charset=utf-8;base64," + b64Html + "#screenshot";
            }
            else
            {
                dataUrl = new Uri(Path.GetFullPath(inputHtmlOrPath)).AbsoluteUri + "#screenshot";
            }

            int id2 = Interlocked.Increment(ref _messageId);
            string encodedUrl = EscapeJsonString(dataUrl);
            string navJson = $"{{\"id\":{id2},\"method\":\"Page.navigate\",\"params\":{{\"url\":{encodedUrl}}}}}";
            await SendRawCdpAsync(navJson, id2, ct);

            if (!isRawHtml)
                await WaitForLoadEventAsync(ct);

            int id3 = Interlocked.Increment(ref _messageId);
            string shotJson = $"{{\"id\":{id3},\"method\":\"Page.captureScreenshot\",\"params\":{{\"format\":\"png\"}}}}";
            var result = await SendRawCdpAsync(shotJson, id3, ct);

            if (result.TryGetProperty("data", out var dataProp))
            {
                string? b64 = dataProp.GetString();
                if (!string.IsNullOrEmpty(b64))
                    return Convert.FromBase64String(b64);
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"[CDP Warm Screenshot Error]: {ex.Message}");
            Dispose();
        }
        return null;
    }

    public void Dispose()
    {
        lock (_lock)
        {
            if (_disposed) return;
            _disposed = true;

            try { _ws?.Dispose(); } catch { }
            _ws = null;

            try
            {
                if (_chromeProcess != null && !_chromeProcess.HasExited)
                {
                    _chromeProcess.Kill(true);
                    _chromeProcess.Dispose();
                }
            }
            catch { }
            _chromeProcess = null;
        }
    }
}
