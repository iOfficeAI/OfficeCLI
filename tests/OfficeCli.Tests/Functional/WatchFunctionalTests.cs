// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.IO.Pipes;
using System.Net;
using System.Net.Sockets;
using System.Text;
using FluentAssertions;
using OfficeCli;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Functional tests for the Watch feature: pipe name generation, slide number extraction,
/// pipe notification, HTTP serving, SSE incremental updates, and single-slide rendering.
/// </summary>
public class WatchFunctionalTests : IDisposable
{
    private readonly string _path;
    private PowerPointHandler _handler;

    public WatchFunctionalTests()
    {
        _path = Path.Combine(Path.GetTempPath(), $"test_watch_{Guid.NewGuid():N}.pptx");
        BlankDocCreator.Create(_path);
        _handler = new PowerPointHandler(_path, editable: true);
    }

    public void Dispose()
    {
        _handler.Dispose();
        if (File.Exists(_path)) File.Delete(_path);
    }

    private PowerPointHandler Reopen()
    {
        _handler.Dispose();
        _handler = new PowerPointHandler(_path, editable: true);
        return _handler;
    }

    /// <summary>Find a free TCP port by binding to port 0 and reading the assigned port.</summary>
    private static int GetFreePort()
    {
        var listener = new TcpListener(IPAddress.Loopback, 0);
        listener.Start();
        var port = ((IPEndPoint)listener.LocalEndpoint).Port;
        listener.Stop();
        return port;
    }

    // ==================== Pipe name generation ====================

    [Fact]
    public void GetWatchPipeName_IsDeterministic()
    {
        var name1 = WatchServer.GetWatchPipeName(_path);
        var name2 = WatchServer.GetWatchPipeName(_path);
        name1.Should().Be(name2);
    }

    [Fact]
    public void GetWatchPipeName_StartsWithPrefix()
    {
        var name = WatchServer.GetWatchPipeName(_path);
        name.Should().StartWith("officecli-watch-");
    }

    [Fact]
    public void GetWatchPipeName_DifferentFilesProduceDifferentNames()
    {
        var name1 = WatchServer.GetWatchPipeName("/tmp/a.pptx");
        var name2 = WatchServer.GetWatchPipeName("/tmp/b.pptx");
        name1.Should().NotBe(name2);
    }

    [Fact]
    public void GetWatchPipeName_DiffersFromResidentPipeName()
    {
        var watchName = WatchServer.GetWatchPipeName(_path);
        var residentName = ResidentServer.GetPipeName(_path);
        watchName.Should().NotBe(residentName);
    }

    // ==================== Single-slide rendering ====================

    [Fact]
    public void RenderSlideHtml_ReturnsFragmentForExistingSlide()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Hello", ["x"] = "2cm", ["y"] = "3cm", ["width"] = "10cm", ["height"] = "4cm" });

        var html = _handler.RenderSlideHtml(1);

        html.Should().NotBeNull();
        html.Should().Contain("data-slide=\"1\"");
        html.Should().Contain("slide-container");
        html.Should().Contain("Hello");
    }

    [Fact]
    public void RenderSlideHtml_ReturnsNullForInvalidSlide()
    {
        _handler.Add("/", "slide", null, new());

        _handler.RenderSlideHtml(0).Should().BeNull();
        _handler.RenderSlideHtml(2).Should().BeNull();
    }

    [Fact]
    public void RenderSlideHtml_ReflectsModifications()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new() { ["text"] = "Before", ["x"] = "1cm", ["y"] = "1cm", ["width"] = "10cm", ["height"] = "3cm" });

        var before = _handler.RenderSlideHtml(1);
        before.Should().Contain("Before");

        _handler.Set("/slide[1]/shape[1]", new() { ["text"] = "After" });

        var after = _handler.RenderSlideHtml(1);
        after.Should().Contain("After");
        after.Should().NotContain("Before");
    }

    [Fact]
    public void GetSlideCount_ReturnsCorrectCount()
    {
        _handler.GetSlideCount().Should().Be(0);

        _handler.Add("/", "slide", null, new());
        _handler.GetSlideCount().Should().Be(1);

        _handler.Add("/", "slide", null, new());
        _handler.GetSlideCount().Should().Be(2);

        _handler.Remove("/slide[2]");
        _handler.GetSlideCount().Should().Be(1);
    }

    // ==================== ViewAsHtml data-slide attribute ====================

    [Fact]
    public void ViewAsHtml_ContainsDataSlideAttributes()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/", "slide", null, new());

        var html = _handler.ViewAsHtml();

        html.Should().Contain("data-slide=\"1\"");
        html.Should().Contain("data-slide=\"2\"");
    }

    // ==================== Pipe notification ====================

    [Fact]
    public void WatchNotifier_SilentlyIgnoresWhenNoWatch()
    {
        // Should not throw when no watch process is running
        var act = () => WatchNotifier.NotifyIfWatching(_path, "/slide[1]");
        act.Should().NotThrow();
    }

    [Fact]
    public async Task WatchServer_ReceivesPipeNotification()
    {
        using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(10));
        var pipeName = WatchServer.GetWatchPipeName(_path);
        string? receivedMessage = null;

        // Start a pipe listener (simulating what WatchServer does)
        var listenerTask = Task.Run(async () =>
        {
            var server = new NamedPipeServerStream(
                pipeName, PipeDirection.InOut,
                NamedPipeServerStream.MaxAllowedServerInstances,
                PipeTransmissionMode.Byte, PipeOptions.Asynchronous);
            try
            {
                await server.WaitForConnectionAsync(cts.Token);
                using var reader = new StreamReader(server, Encoding.UTF8, leaveOpen: true);
                using var writer = new StreamWriter(server, Encoding.UTF8, leaveOpen: true) { AutoFlush = true };

                receivedMessage = await reader.ReadLineAsync(cts.Token);
                await writer.WriteLineAsync("ok".AsMemory(), cts.Token);
            }
            finally
            {
                await server.DisposeAsync();
            }
        }, cts.Token);

        // Give the listener time to start
        await Task.Delay(200, cts.Token);

        // Send notification
        WatchNotifier.NotifyIfWatching(_path, "/slide[1]/shape[2]");

        await listenerTask;

        receivedMessage.Should().Be("refresh:/slide[1]/shape[2]");
    }

    [Fact]
    public async Task WatchServer_ReceivesRefreshWithoutPath()
    {
        using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(10));
        var pipeName = WatchServer.GetWatchPipeName(_path);
        string? receivedMessage = null;

        var listenerTask = Task.Run(async () =>
        {
            var server = new NamedPipeServerStream(
                pipeName, PipeDirection.InOut,
                NamedPipeServerStream.MaxAllowedServerInstances,
                PipeTransmissionMode.Byte, PipeOptions.Asynchronous);
            try
            {
                await server.WaitForConnectionAsync(cts.Token);
                using var reader = new StreamReader(server, Encoding.UTF8, leaveOpen: true);
                using var writer = new StreamWriter(server, Encoding.UTF8, leaveOpen: true) { AutoFlush = true };

                receivedMessage = await reader.ReadLineAsync(cts.Token);
                await writer.WriteLineAsync("ok".AsMemory(), cts.Token);
            }
            finally
            {
                await server.DisposeAsync();
            }
        }, cts.Token);

        await Task.Delay(200, cts.Token);
        WatchNotifier.NotifyIfWatching(_path);
        await listenerTask;

        receivedMessage.Should().Be("refresh");
    }

    // ==================== HTTP server ====================

    [Fact]
    public async Task WatchServer_ServesHtmlOnHttp()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Add("/slide[1]", "shape", null, new() { ["text"] = "HttpTest", ["x"] = "1cm", ["y"] = "1cm", ["width"] = "10cm", ["height"] = "3cm" });
        _handler.Dispose();

        var port = GetFreePort();
        using var watch = new WatchServer(_path, port);
        using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(10));

        var serverTask = watch.RunAsync(cts.Token);

        // Wait for server to start
        await Task.Delay(500, cts.Token);

        // Fetch HTML
        using var client = new TcpClient("localhost", port);
        var stream = client.GetStream();
        var request = Encoding.UTF8.GetBytes("GET / HTTP/1.1\r\nHost: localhost\r\n\r\n");
        await stream.WriteAsync(request, cts.Token);

        var buffer = new byte[65536];
        var read = await stream.ReadAsync(buffer, cts.Token);
        var response = Encoding.UTF8.GetString(buffer, 0, read);

        response.Should().Contain("HTTP/1.1 200 OK");
        response.Should().Contain("text/html");
        response.Should().Contain("HttpTest");
        response.Should().Contain("EventSource"); // SSE script injected

        cts.Cancel();
        try { await serverTask; } catch (OperationCanceledException) { }

        // Re-open handler for Dispose
        _handler = new PowerPointHandler(_path, editable: true);
    }

    // ==================== Batch + Watch integration ====================

    [Fact]
    public async Task Batch_SendsWatchNotification_WhenWatchIsRunning()
    {
        // Arrange: create a slide so the file is valid
        _handler.Add("/", "slide", null, new());
        _handler.Dispose();

        using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(10));
        var pipeName = WatchServer.GetWatchPipeName(_path);
        string? receivedMessage = null;

        // Simulate WatchServer listening on the named pipe
        var listenerTask = Task.Run(async () =>
        {
            var server = new NamedPipeServerStream(
                pipeName, PipeDirection.InOut,
                NamedPipeServerStream.MaxAllowedServerInstances,
                PipeTransmissionMode.Byte, PipeOptions.Asynchronous);
            try
            {
                await server.WaitForConnectionAsync(cts.Token);
                using var reader = new StreamReader(server, Encoding.UTF8, leaveOpen: true);
                using var writer = new StreamWriter(server, Encoding.UTF8, leaveOpen: true) { AutoFlush = true };
                receivedMessage = await reader.ReadLineAsync(cts.Token);
                await writer.WriteLineAsync("ok".AsMemory(), cts.Token);
            }
            finally { await server.DisposeAsync(); }
        }, cts.Token);

        await Task.Delay(200, cts.Token); // let listener start

        // Act: run batch command via CLI
        var batchJson = System.Text.Json.JsonSerializer.Serialize(new[]
        {
            new { command = "add", parent = "/", type = "slide", props = new Dictionary<string, string>() },
        });
        var batchFile = Path.Combine(Path.GetTempPath(), $"batch_{Guid.NewGuid():N}.json");
        await File.WriteAllTextAsync(batchFile, batchJson, cts.Token);
        try
        {
            var root = CommandBuilder.BuildRootCommand();
            root.Parse(["batch", _path, "--input", batchFile]).Invoke();
        }
        finally { File.Delete(batchFile); }

        await listenerTask;

        // Assert: watch received a "refresh" notification
        receivedMessage.Should().StartWith("refresh");
    }

    // ==================== SSE endpoint ====================

    [Fact]
    public async Task WatchServer_SseEndpointReturnsEventStream()
    {
        _handler.Add("/", "slide", null, new());
        _handler.Dispose();

        var port = GetFreePort();
        using var watch = new WatchServer(_path, port);
        using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(10));

        var serverTask = watch.RunAsync(cts.Token);
        await Task.Delay(500, cts.Token);

        using var client = new TcpClient("localhost", port);
        var stream = client.GetStream();
        var request = Encoding.UTF8.GetBytes("GET /events HTTP/1.1\r\nHost: localhost\r\n\r\n");
        await stream.WriteAsync(request, cts.Token);

        var buffer = new byte[4096];
        var read = await stream.ReadAsync(buffer, cts.Token);
        var response = Encoding.UTF8.GetString(buffer, 0, read);

        response.Should().Contain("HTTP/1.1 200 OK");
        response.Should().Contain("text/event-stream");

        cts.Cancel();
        try { await serverTask; } catch (OperationCanceledException) { }

        _handler = new PowerPointHandler(_path, editable: true);
    }
}
