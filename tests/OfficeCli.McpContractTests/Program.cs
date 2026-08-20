using System.Collections.Concurrent;
using System.Text;
using System.Text.Json;
using OfficeCli;

static JsonElement Id(string json)
{
    using var document = JsonDocument.Parse(json);
    return document.RootElement.Clone();
}

static JsonDocument Parse(string json) => JsonDocument.Parse(json);

static void Require(bool condition, string message)
{
    if (!condition)
        throw new InvalidOperationException(message);
}

using (var initialize = Parse(McpServer.HandleInitialize(Id("17"))))
{
    var root = initialize.RootElement;
    Require(root.GetProperty("jsonrpc").GetString() == "2.0", "initialize must use JSON-RPC 2.0");
    Require(root.GetProperty("id").GetInt32() == 17, "initialize must preserve numeric request IDs");
    Require(root.GetProperty("result").GetProperty("protocolVersion").GetString() == "2024-11-05", "protocol version drifted");
    Require(root.GetProperty("result").GetProperty("capabilities").TryGetProperty("tools", out _), "tools capability missing");
}

using (var tools = Parse(McpServer.HandleToolsList(Id("\"tools-1\""))))
{
    var root = tools.RootElement;
    Require(root.GetProperty("id").GetString() == "tools-1", "tools/list must preserve string request IDs");
    var definitions = root.GetProperty("result").GetProperty("tools");
    Require(definitions.GetArrayLength() == 1, "server must advertise exactly one tool");
    var tool = definitions[0];
    Require(tool.GetProperty("name").GetString() == "officecli", "unexpected MCP tool name");
    Require(tool.GetProperty("inputSchema").GetProperty("type").GetString() == "object", "tool schema must be an object");
}

var callRoot = Id("""
    {"jsonrpc":"2.0","id":"call-a","method":"tools/call","params":{"name":"officecli","arguments":{"command":["skills","list"]}}}
    """);
var concurrentResponses = await Task.WhenAll(
    McpServer.HandleToolsCallAsync(Id("\"call-a\""), callRoot, CancellationToken.None),
    McpServer.HandleToolsCallAsync(Id("\"call-b\""), callRoot, CancellationToken.None));
var expectedIds = new[] { "call-a", "call-b" };
for (var index = 0; index < concurrentResponses.Length; index++)
{
    using var response = Parse(concurrentResponses[index]);
    var root = response.RootElement;
    Require(root.GetProperty("id").GetString() == expectedIds[index], "concurrent response IDs crossed");
    var result = root.GetProperty("result");
    Require(!result.GetProperty("isError").GetBoolean(), "skills/list MCP call failed");
    var content = result.GetProperty("content");
    Require(content.GetArrayLength() > 0, "structured content must not be flattened away");
    Require(content[0].GetProperty("type").GetString() == "text", "structured content type missing");
}

using var firstCancellation = new CancellationTokenSource();
using var secondCancellation = new CancellationTokenSource();
var inFlight = new ConcurrentDictionary<string, McpServer.InFlightRequest>();
inFlight["101"] = new(firstCancellation);
inFlight["102"] = new(secondCancellation);
using (var cancellation = Parse("""
    {"jsonrpc":"2.0","method":"notifications/cancelled","params":{"requestId":101,"reason":"test"}}
    """))
{
    McpServer.CancelRequest(cancellation.RootElement, inFlight);
}
Require(firstCancellation.IsCancellationRequested, "target request was not cancelled");
Require(!secondCancellation.IsCancellationRequested, "cancelling one request affected another request");

var slowStarted = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
McpServer.CommandExecutorOverride = async (arguments, token) =>
{
    var command = arguments.GetProperty("command").GetString();
    if (command == "slow")
    {
        slowStarted.SetResult();
        await Task.Delay(Timeout.InfiniteTimeSpan, token);
    }
    if (command == "image")
    {
        return (new McpServer.McpContent[]
        {
            new("text", Text: "preview"),
            new("image", Data: "iVBORw0KGgo=", MimeType: "image/png"),
            new("resource", Text: "resource body", MimeType: "text/plain", Uri: "file:///report.txt", Name: "report"),
        }, false);
    }
    return (new McpServer.McpContent[] { new("text", Text: "{\"ok\":true,\"request\":\"fast\"}") }, false);
};

try
{
    var stdioInput = string.Join('\n', new[]
    {
        "{\"jsonrpc\":\"2.0\",\"id\":1,\"method\":\"initialize\",\"params\":{\"protocolVersion\":\"2024-11-05\",\"capabilities\":{},\"clientInfo\":{\"name\":\"contract\",\"version\":\"1\"}}}",
        "{\"jsonrpc\":\"2.0\",\"id\":2,\"method\":\"tools/call\",\"params\":{\"name\":\"officecli\",\"arguments\":{\"command\":\"slow\"}}}",
        "{\"jsonrpc\":\"2.0\",\"id\":3,\"method\":\"tools/call\",\"params\":{\"name\":\"officecli\",\"arguments\":{\"command\":\"fast\"}}}",
        "{\"jsonrpc\":\"2.0\",\"method\":\"notifications/cancelled\",\"params\":{\"requestId\":2,\"reason\":\"contract test\"}}",
        "{\"jsonrpc\":\"2.0\",\"id\":4,\"method\":\"tools/call\",\"params\":{\"name\":\"officecli\",\"arguments\":{\"command\":\"image\"}}}",
        "",
    });
    await using var input = new MemoryStream(Encoding.UTF8.GetBytes(stdioInput));
    await using var output = new MemoryStream();
    await McpServer.RunAsync(input, output, enableUpgradeCheck: false);
    await slowStarted.Task.WaitAsync(TimeSpan.FromSeconds(1));

    output.Position = 0;
    using var outputReader = new StreamReader(output, Encoding.UTF8, false, 1024, leaveOpen: true);
    var lines = (await outputReader.ReadToEndAsync()).Split('\n', StringSplitOptions.RemoveEmptyEntries);
    var responses = lines.Select(line => JsonDocument.Parse(line)).ToArray();
    try
    {
        Require(responses.All(response => response.RootElement.GetProperty("jsonrpc").GetString() == "2.0"), "stdout contained non-JSON-RPC output");
        Require(responses.Any(response => response.RootElement.GetProperty("id").GetInt32() == 1), "initialize response missing");
        Require(responses.Any(response => response.RootElement.GetProperty("id").GetInt32() == 3), "concurrent fast response missing");
        Require(!responses.Any(response => response.RootElement.GetProperty("id").GetInt32() == 2), "cancelled request emitted a late response");
        var structured = responses.Single(response => response.RootElement.GetProperty("id").GetInt32() == 3)
            .RootElement.GetProperty("result").GetProperty("structuredContent");
        Require(structured.GetProperty("ok").GetBoolean(), "structuredContent was flattened");
        var richContent = responses.Single(response => response.RootElement.GetProperty("id").GetInt32() == 4)
            .RootElement.GetProperty("result").GetProperty("content");
        Require(richContent[1].GetProperty("type").GetString() == "image", "image content shape drifted");
        Require(richContent[2].GetProperty("resource").GetProperty("uri").GetString() == "file:///report.txt", "resource content shape drifted");
    }
    finally
    {
        foreach (var response in responses) response.Dispose();
    }
}
finally
{
    McpServer.CommandExecutorOverride = null;
}

Console.WriteLine("MCP contract tests passed.");
