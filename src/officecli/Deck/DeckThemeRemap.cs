// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Globalization;
using System.Text.Json;

namespace OfficeCli.Deck;

/// <summary>
/// Remap a DeckSpec theme.id, report slides whose layouts may need a same-role swap,
/// and suggest alternatives via <see cref="DeckLayoutQuery"/>. Original CSBU WorkMate flow
/// (no AGPL / third-party theme grids).
/// </summary>
public static class DeckThemeRemap
{
    public static DeckThemeRemapResult Remap(
        DeckSpec spec,
        string toThemeId,
        DeckThemeRemapOptions? options = null,
        DeckCatalog? catalog = null)
    {
        ArgumentNullException.ThrowIfNull(spec);
        if (string.IsNullOrWhiteSpace(toThemeId))
            throw new ArgumentException("Target theme id is required.", nameof(toThemeId));

        options ??= new DeckThemeRemapOptions();
        catalog ??= DeckCatalogLoader.Load();
        var toId = toThemeId.Trim();
        var themeById = catalog.Themes.ToDictionary(theme => theme.Id, StringComparer.Ordinal);
        if (!themeById.TryGetValue(toId, out var toTheme))
            throw new ArgumentException($"Unknown theme '{toId}'.", nameof(toThemeId));

        var fromId = spec.Theme.Id;
        themeById.TryGetValue(fromId, out var fromTheme);
        var fromMode = InferMode(fromTheme);
        var toMode = InferMode(toTheme);
        var modeChange = fromMode != toMode ? $"{fromMode}->{toMode}" : null;
        var layoutById = catalog.Layouts.ToDictionary(layout => layout.Id, StringComparer.Ordinal);
        var limit = Math.Clamp(options.Limit <= 0 ? 5 : options.Limit, 1, 20);

        var slides = new List<DeckThemeRemapSlide>();
        foreach (var slide in spec.Slides)
        {
            var knownLayout = layoutById.ContainsKey(slide.LayoutId);
            var hints = InferContentHints(slide, layoutById.GetValueOrDefault(slide.LayoutId));
            var query = DeckLayoutQuery.Query(
                new DeckLayoutQueryRequest(
                    Role: slide.Role,
                    ItemCount: hints.ItemCount,
                    HasChart: hints.HasChart,
                    NeedsMedia: hints.NeedsMedia,
                    HasTable: hints.HasTable,
                    Limit: limit),
                catalog);

            var alternatives = query.Results
                .Where(hit => !string.Equals(hit.LayoutId, slide.LayoutId, StringComparison.Ordinal))
                .Take(limit)
                .ToList();
            var currentHit = query.Results.FirstOrDefault(hit =>
                string.Equals(hit.LayoutId, slide.LayoutId, StringComparison.Ordinal));
            var topHit = query.Results.FirstOrDefault();

            var reasons = new List<string>();
            if (!knownLayout) reasons.Add("unknown_layout");
            if (modeChange != null) reasons.Add("theme_mode_shift");
            if (topHit != null
                && !string.Equals(topHit.LayoutId, slide.LayoutId, StringComparison.Ordinal)
                && (currentHit is null || topHit.Score - currentHit.Score >= 2.0))
            {
                reasons.Add("better_alternative");
            }
            if (knownLayout
                && currentHit is null
                && query.Results.Count > 0)
            {
                reasons.Add("not_in_top");
            }

            // needsRemap: unknown layout always; otherwise mode shift or a clearly better alt.
            var needsRemap = reasons.Contains("unknown_layout")
                || reasons.Contains("better_alternative")
                || (reasons.Contains("theme_mode_shift") && alternatives.Count > 0 && currentHit is null);

            slides.Add(new DeckThemeRemapSlide(
                SlideId: slide.Id,
                Role: slide.Role,
                LayoutId: slide.LayoutId,
                NeedsRemap: needsRemap,
                Reasons: reasons,
                Alternatives: alternatives));
        }

        var report = new DeckThemeRemapReport(
            FromThemeId: fromId,
            ToThemeId: toId,
            ModeChange: modeChange,
            CatalogVersion: catalog.Version,
            CatalogHash: catalog.Hash,
            GeneratedAt: DateTimeOffset.UtcNow.ToString("O", CultureInfo.InvariantCulture),
            NeedsRemapCount: slides.Count(item => item.NeedsRemap),
            Slides: slides);

        DeckSpec? updated = null;
        if (options.Apply)
        {
            var extensions = new Dictionary<string, JsonElement>(spec.Extensions, StringComparer.Ordinal);
            if (options.WriteReport)
            {
                var reportJson = JsonSerializer.SerializeToElement(report, DeckJsonContext.Default.DeckThemeRemapReport);
                extensions["themeRemap"] = reportJson;
            }

            var nextMode = options.SetMode ? toMode : spec.Theme.Mode;
            updated = spec with
            {
                Revision = spec.Revision + 1,
                Theme = spec.Theme with
                {
                    Id = toId,
                    Mode = nextMode,
                },
                Extensions = extensions,
            };
        }

        return new DeckThemeRemapResult(
            Applied: options.Apply,
            SpecPath: options.SpecPath,
            Report: report,
            Spec: updated);
    }

    public static void SaveSpec(DeckSpec spec, string path)
    {
        ArgumentNullException.ThrowIfNull(spec);
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("Spec path is required.", nameof(path));
        var json = JsonSerializer.Serialize(spec, DeckJsonContext.Default.DeckSpec);
        File.WriteAllText(path, json + Environment.NewLine);
    }

    private static string InferMode(DeckTheme? theme)
    {
        if (theme is null) return "light";
        if (!theme.Tokens.TryGetValue("background", out var background) || string.IsNullOrWhiteSpace(background))
            return "light";
        var hex = background.Trim().TrimStart('#');
        if (hex.Length is not (6 or 8)) return "light";
        if (!int.TryParse(hex[..2], NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var r)
            || !int.TryParse(hex[2..4], NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var g)
            || !int.TryParse(hex[4..6], NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var b))
            return "light";
        // Relative luminance threshold — dark backgrounds → dark mode.
        var luminance = (0.2126 * r + 0.7152 * g + 0.0722 * b) / 255.0;
        return luminance < 0.45 ? "dark" : "light";
    }

    private static (int? ItemCount, bool? HasChart, bool? NeedsMedia, bool? HasTable) InferContentHints(
        DeckSlide slide,
        DeckLayout? layout)
    {
        var blocks = slide.Blocks ?? [];
        var hasChart = blocks.Any(block => string.Equals(block.Type, "chart", StringComparison.Ordinal));
        var needsMedia = blocks.Any(block =>
            string.Equals(block.Type, "image", StringComparison.Ordinal)
            || string.Equals(block.Type, "shape", StringComparison.Ordinal));
        var hasTable = blocks.Any(block => string.Equals(block.Type, "table", StringComparison.Ordinal));

        var itemish = blocks.Count(block => block.Type is "metric" or "list" or "timeline" or "text");
        if (itemish == 0 && layout != null)
            itemish = DeckLayoutQuery.EstimateCapacity(layout);

        bool? chartHint = hasChart ? true : null;
        bool? mediaHint = needsMedia ? true : null;
        bool? tableHint = hasTable ? true : null;
        int? itemCount = itemish > 0 ? itemish : null;
        return (itemCount, chartHint, mediaHint, tableHint);
    }
}

public sealed record DeckThemeRemapOptions(
    bool Apply = false,
    bool WriteReport = true,
    bool SetMode = true,
    int Limit = 5,
    string? SpecPath = null);

public sealed record DeckThemeRemapSlide(
    string SlideId,
    string Role,
    string LayoutId,
    bool NeedsRemap,
    IReadOnlyList<string> Reasons,
    IReadOnlyList<DeckLayoutQueryHit> Alternatives);

public sealed record DeckThemeRemapReport(
    string FromThemeId,
    string ToThemeId,
    string? ModeChange,
    string CatalogVersion,
    string CatalogHash,
    string GeneratedAt,
    int NeedsRemapCount,
    IReadOnlyList<DeckThemeRemapSlide> Slides);

public sealed record DeckThemeRemapResult(
    bool Applied,
    string? SpecPath,
    DeckThemeRemapReport Report,
    DeckSpec? Spec);
