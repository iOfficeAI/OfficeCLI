// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Text.Json;
using System.Text.Json.Serialization;

namespace OfficeCli.Deck;

public sealed record DeckSpec
{
    public int SchemaVersion { get; init; }
    public long Revision { get; init; }
    public string Stage { get; init; } = "ready";
    public DeckMetadata Metadata { get; init; } = new();
    public DeckThemeSelection Theme { get; init; } = new();
    public List<DeckSlide> Slides { get; init; } = [];
    public List<DeckAsset> Assets { get; init; } = [];
    public Dictionary<string, JsonElement> Extensions { get; init; } = [];
}

public sealed record DeckMetadata
{
    public string Title { get; init; } = "";
    public string? Goal { get; init; }
    public string? Audience { get; init; }
    public string Language { get; init; } = "en-US";
    public string AspectRatio { get; init; } = "16:9";
    public string? Author { get; init; }
}

public sealed record DeckThemeSelection
{
    public string Id { get; init; } = "business-light";
    public string? Mode { get; init; }
    public Dictionary<string, string> BrandTokens { get; init; } = [];
}

public sealed record DeckSlide
{
    public string Id { get; init; } = "";
    public string Role { get; init; } = "statement";
    public string LayoutId { get; init; } = "statement";
    public string? Title { get; init; }
    public string? Notes { get; init; }
    public bool Hidden { get; init; }
    public List<DeckBlock> Blocks { get; init; } = [];
    public Dictionary<string, JsonElement> Controls { get; init; } = [];
    /// <summary>Optional preferred same-slide layout IDs (P1.6). Export uses LayoutId only.</summary>
    public List<string>? Candidates { get; init; }
}

public sealed record DeckBlock
{
    public string Id { get; init; } = "";
    public string Type { get; init; } = "text";
    public string? Slot { get; init; }
    public string? Text { get; init; }
    public string? Value { get; init; }
    public string? Label { get; init; }
    public string? AssetId { get; init; }
    public List<string> Items { get; init; } = [];
    public JsonElement? Data { get; init; }
}

public sealed record DeckAsset
{
    public string Id { get; init; } = "";
    public string Path { get; init; } = "";
    public string Type { get; init; } = "image";
    public string Status { get; init; } = "ready";
    public string? Alt { get; init; }
    public string? Source { get; init; }
    public string? Model { get; init; }
    public string? PromptSummary { get; init; }
}

public sealed record DeckCatalog(
    string Version,
    string Hash,
    IReadOnlyList<DeckTheme> Themes,
    IReadOnlyList<DeckLayout> Layouts);

public sealed record DeckTheme(
    string Id,
    string Label,
    IReadOnlyDictionary<string, string> Tokens);

public sealed record DeckLayout(
    string Id,
    string Role,
    string Label,
    IReadOnlyList<DeckSlot> Slots,
    IReadOnlyList<DeckControl> Controls,
    string OverflowStrategy = "diagnose",
    IReadOnlyList<string>? AlternativeLayoutIds = null);

public sealed record DeckSlot(
    string Id,
    double X,
    double Y,
    double Width,
    double Height,
    IReadOnlyList<string> Accepts,
    bool Required = false,
    bool Toggleable = false,
    int? MaxLength = null,
    int? MaxItems = null);

public sealed record DeckControl(
    string Id,
    string Type,
    string Label,
    JsonElement DefaultValue,
    IReadOnlyList<string>? Options = null,
    double? Min = null,
    double? Max = null,
    double? Step = null);

public sealed record DeckDiagnostic(
    string Severity,
    string Code,
    string Message,
    string? Path = null,
    string? SlideId = null,
    string? BlockId = null,
    string? Suggestion = null);

public sealed record DeckValidationResult(bool Valid, IReadOnlyList<DeckDiagnostic> Diagnostics);

public sealed record DeckBuildResult(bool Success, string Output, long Revision);

public sealed record DeckPreviewScene(
    long Revision,
    string ThemeId,
    IReadOnlyDictionary<string, string> ThemeTokens,
    IReadOnlyList<DeckPreviewSlide> Slides);

public sealed record DeckPreviewSlide(
    string Id,
    string LayoutId,
    string? Title,
    bool Hidden,
    IReadOnlyList<DeckPreviewElement> Elements);

public sealed record DeckPreviewElement(
    string Id,
    string Type,
    string Slot,
    double X,
    double Y,
    double Width,
    double Height,
    string Text,
    string? AssetPath = null,
    JsonElement? Data = null);

public sealed record DeckLayoutQueryRequest(
    string? Role = null,
    int? ItemCount = null,
    bool? HasChart = null,
    bool? NeedsMedia = null,
    bool? HasTable = null,
    string? Query = null,
    int Limit = 8);

public sealed record DeckLayoutQueryHit(
    string LayoutId,
    string Role,
    string Label,
    double Score,
    int Capacity,
    IReadOnlyList<string> Accepts,
    IReadOnlyList<string> Reasons,
    IReadOnlyList<string> AlternativeLayoutIds);

public sealed record DeckLayoutQueryResult(
    DeckLayoutQueryRequest Query,
    string CatalogVersion,
    string CatalogHash,
    int ResultCount,
    IReadOnlyList<DeckLayoutQueryHit> Results);

[JsonSourceGenerationOptions(
    PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase,
    PropertyNameCaseInsensitive = true,
    DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
    WriteIndented = true)]
[JsonSerializable(typeof(DeckSpec))]
[JsonSerializable(typeof(DeckCatalogSource))]
[JsonSerializable(typeof(DeckCatalog))]
[JsonSerializable(typeof(DeckValidationResult))]
[JsonSerializable(typeof(DeckBuildResult))]
[JsonSerializable(typeof(DeckPreviewScene))]
[JsonSerializable(typeof(DeckLayoutQueryRequest))]
[JsonSerializable(typeof(DeckLayoutQueryHit))]
[JsonSerializable(typeof(DeckLayoutQueryResult))]
[JsonSerializable(typeof(DeckThemeRemapReport))]
[JsonSerializable(typeof(DeckThemeRemapResult))]
[JsonSerializable(typeof(DeckThemeRemapSlide))]
[JsonSerializable(typeof(DeckScaffoldReport))]
[JsonSerializable(typeof(DeckScaffoldResult))]
[JsonSerializable(typeof(DeckScaffoldCliResult))]
[JsonSerializable(typeof(Dictionary<string, int>))]
internal partial class DeckJsonContext : JsonSerializerContext;
