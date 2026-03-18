// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.Json.Serialization;

namespace OfficeCli.Core;

public class BatchItem
{
    [JsonPropertyName("command")]
    public string Command { get; set; } = "";

    [JsonPropertyName("path")]
    public string? Path { get; set; }

    [JsonPropertyName("parent")]
    public string? Parent { get; set; }

    [JsonPropertyName("type")]
    public string? Type { get; set; }

    [JsonPropertyName("from")]
    public string? From { get; set; }

    [JsonPropertyName("index")]
    public int? Index { get; set; }

    [JsonPropertyName("to")]
    public string? To { get; set; }

    [JsonPropertyName("props")]
    public Dictionary<string, string>? Props { get; set; }

    [JsonPropertyName("selector")]
    public string? Selector { get; set; }

    [JsonPropertyName("mode")]
    public string? Mode { get; set; }

    [JsonPropertyName("depth")]
    public int? Depth { get; set; }

    [JsonPropertyName("part")]
    public string? Part { get; set; }

    [JsonPropertyName("xpath")]
    public string? Xpath { get; set; }

    [JsonPropertyName("action")]
    public string? Action { get; set; }

    [JsonPropertyName("xml")]
    public string? Xml { get; set; }

    public ResidentRequest ToResidentRequest()
    {
        var req = new ResidentRequest { Command = Command };

        if (Path != null) req.Args["path"] = Path;
        if (Parent != null) req.Args["parent"] = Parent;
        if (Type != null) req.Args["type"] = Type;
        if (From != null) req.Args["from"] = From;
        if (Index.HasValue) req.Args["index"] = Index.Value.ToString();
        if (To != null) req.Args["to"] = To;
        if (Selector != null) req.Args["selector"] = Selector;
        if (Mode != null) req.Args["mode"] = Mode;
        if (Depth.HasValue) req.Args["depth"] = Depth.Value.ToString();
        if (Part != null) req.Args["part"] = Part;
        if (Xpath != null) req.Args["xpath"] = Xpath;
        if (Action != null) req.Args["action"] = Action;
        if (Xml != null) req.Args["xml"] = Xml;

        if (Props != null)
            req.Props = Props.Select(kv => $"{kv.Key}={kv.Value}").ToArray();

        return req;
    }
}

public class BatchResult
{
    [JsonPropertyName("success")]
    public bool Success { get; set; }

    [JsonPropertyName("output")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Output { get; set; }

    [JsonPropertyName("error")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Error { get; set; }
}

[JsonSourceGenerationOptions]
[JsonSerializable(typeof(List<BatchItem>))]
[JsonSerializable(typeof(List<BatchResult>))]
internal partial class BatchJsonContext : JsonSerializerContext;
