using System.Collections.Concurrent;
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

Console.WriteLine("MCP contract tests passed.");
