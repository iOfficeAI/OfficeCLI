// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace OfficeCli.Deck;

public static class DeckCatalogLoader
{
    private const string ResourceName = "OfficeCli.Deck.catalog.json";
    private static readonly Lazy<DeckCatalog> Catalog = new(LoadCore);

    public static DeckCatalog Load() => Catalog.Value;

    private static DeckCatalog LoadCore()
    {
        using var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(ResourceName)
            ?? throw new InvalidOperationException("Embedded presentation catalog is unavailable.");
        using var reader = new StreamReader(stream, Encoding.UTF8);
        var json = reader.ReadToEnd();
        var source = JsonSerializer.Deserialize(json, DeckJsonContext.Default.DeckCatalogSource)
            ?? throw new InvalidOperationException("Embedded presentation catalog is invalid.");
        var hash = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(json))).ToLowerInvariant();
        var layouts = source.Layouts.Select(layout => layout with
        {
            Slots = layout.Slots.Select(slot => slot with
            {
                MaxLength = slot.MaxLength ?? Math.Max(40, (int)(slot.Width * slot.Height * 900)),
                MaxItems = slot.MaxItems ?? (slot.Accepts.Any(type => type is "list" or "timeline") ? 8 : null),
            }).ToList(),
            AlternativeLayoutIds = layout.AlternativeLayoutIds ?? [],
        }).ToList();
        return new DeckCatalog(source.Version, hash, source.Themes, layouts);
    }
}

internal sealed record DeckCatalogSource
{
    public string Version { get; init; } = "1";
    public List<DeckTheme> Themes { get; init; } = [];
    public List<DeckLayout> Layouts { get; init; } = [];
}
