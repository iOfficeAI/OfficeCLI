// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.Json.Serialization;

namespace OfficeCli.Core;

/// <summary>
/// Represents a node in the document DOM tree.
/// This is the universal abstraction across Word/Excel/PowerPoint.
/// </summary>
public class DocumentNode
{
    [JsonPropertyName("path")]
    public string Path { get; set; } = "";
    [JsonPropertyName("type")]
    public string Type { get; set; } = "";
    [JsonPropertyName("text")]
    public string? Text { get; set; }
    [JsonPropertyName("preview")]
    public string? Preview { get; set; }
    [JsonPropertyName("style")]
    public string? Style { get; set; }
    [JsonPropertyName("childCount")]
    public int ChildCount { get; set; }
    [JsonPropertyName("format")]
    public Dictionary<string, object?> Format { get; set; } = new();
    [JsonPropertyName("children")]
    public List<DocumentNode> Children { get; set; } = new();
}
