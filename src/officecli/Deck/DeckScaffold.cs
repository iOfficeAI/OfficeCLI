// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace OfficeCli.Deck;

/// <summary>
/// Outline a long WorkMate deck from goal / audience / page count with section
/// transitions and role-mix heuristics. Original CSBU WorkMate flow (not a
/// third-party goal-spec clone). Emits stage=outline slides with empty blocks.
/// </summary>
public static class DeckScaffold
{
    private static readonly string[] DefaultContentRoles =
    [
        "metrics", "comparison", "process", "trend", "breakdown", "statement",
        "risks", "actions", "context", "observation", "team", "case", "result",
        "distribution", "relationship", "image"
    ];

    public static DeckScaffoldResult Scaffold(DeckScaffoldRequest request, DeckCatalog? catalog = null)
    {
        ArgumentNullException.ThrowIfNull(request);
        catalog ??= DeckCatalogLoader.Load();

        var pages = Math.Clamp(request.Pages <= 0 ? 12 : request.Pages, 4, 60);
        var seed = string.IsNullOrWhiteSpace(request.Seed)
            ? DeriveSeed(request.Goal, request.Audience, pages)
            : request.Seed.Trim();
        var themeId = string.IsNullOrWhiteSpace(request.ThemeId) ? "csbu-workmate" : request.ThemeId.Trim();
        var themeById = catalog.Themes.ToDictionary(theme => theme.Id, StringComparer.Ordinal);
        if (!themeById.TryGetValue(themeId, out var theme))
            throw new ArgumentException($"Unknown theme '{themeId}'.", nameof(request));

        var language = string.IsNullOrWhiteSpace(request.Language) ? "en-US" : request.Language.Trim();
        var title = string.IsNullOrWhiteSpace(request.Title)
            ? Truncate(request.Goal?.Trim() ?? "WorkMate deck", 72)
            : request.Title.Trim();

        var rng = new Random(StableHash(seed));
        var plan = PlanRoles(pages, request.Goal, request.Audience, rng);
        var layoutByRole = catalog.Layouts
            .GroupBy(layout => layout.Role, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.ToList(), StringComparer.Ordinal);

        var slides = new List<DeckSlide>(plan.Count);
        var sectionIndex = 0;
        for (var i = 0; i < plan.Count; i++)
        {
            var entry = plan[i];
            var layout = PickLayout(layoutByRole, entry.Role, rng)
                ?? catalog.Layouts.FirstOrDefault(layout => layout.Role == "statement")
                ?? catalog.Layouts[0];
            var slideId = $"s{i + 1:00}-{SanitizeId(entry.Role)}";
            var slideTitle = entry.Kind switch
            {
                "cover" => title,
                "agenda" => language.StartsWith("zh", StringComparison.OrdinalIgnoreCase) ? "议程" : "Agenda",
                "transition" => language.StartsWith("zh", StringComparison.OrdinalIgnoreCase)
                    ? $"第 {++sectionIndex} 部分"
                    : $"Section {++sectionIndex}",
                "closing" => language.StartsWith("zh", StringComparison.OrdinalIgnoreCase) ? "总结与下一步" : "Summary & next steps",
                _ => PlaceholderTitle(entry.Role, i + 1, language),
            };

            var candidates = layoutByRole.TryGetValue(entry.Role, out var sameRole)
                ? sameRole.Where(item => item.Id != layout.Id).Take(3).Select(item => item.Id).ToList()
                : null;

            slides.Add(new DeckSlide
            {
                Id = slideId,
                Role = entry.Role,
                LayoutId = layout.Id,
                Title = slideTitle,
                Notes = entry.Kind == "transition"
                    ? "Section transition — confirm narrative beat before fill."
                    : null,
                Blocks = [],
                Controls = new(),
                Candidates = candidates is { Count: > 0 } ? candidates : null,
            });
        }

        var report = new DeckScaffoldReport(
            Seed: seed,
            Pages: pages,
            ThemeId: themeId,
            CatalogVersion: catalog.Version,
            CatalogHash: catalog.Hash,
            GeneratedAt: DateTimeOffset.UtcNow.ToString("O", CultureInfo.InvariantCulture),
            RoleMix: plan.GroupBy(item => item.Role).ToDictionary(group => group.Key, group => group.Count(), StringComparer.Ordinal),
            SectionBreaks: plan.Count(item => item.Kind == "transition"),
            Heuristics: BuildHeuristicNotes(pages, request.Goal, request.Audience));

        var extensions = new Dictionary<string, JsonElement>(StringComparer.Ordinal)
        {
            ["deckScaffold"] = JsonSerializer.SerializeToElement(report, DeckJsonContext.Default.DeckScaffoldReport),
        };

        var mode = InferMode(theme);
        var spec = new DeckSpec
        {
            SchemaVersion = 1,
            Revision = 1,
            Stage = "outline",
            Metadata = new DeckMetadata
            {
                Title = title,
                Goal = string.IsNullOrWhiteSpace(request.Goal) ? null : request.Goal.Trim(),
                Audience = string.IsNullOrWhiteSpace(request.Audience) ? null : request.Audience.Trim(),
                Language = language,
                AspectRatio = "16:9",
                Author = "CSBU WorkMate",
            },
            Theme = new DeckThemeSelection
            {
                Id = themeId,
                Mode = mode,
            },
            Slides = slides,
            Assets = [],
            Extensions = extensions,
        };

        return new DeckScaffoldResult(Spec: spec, Report: report);
    }

    public static void SaveSpec(DeckSpec spec, string path)
    {
        ArgumentNullException.ThrowIfNull(spec);
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("Spec path is required.", nameof(path));
        var json = JsonSerializer.Serialize(spec, DeckJsonContext.Default.DeckSpec);
        File.WriteAllText(path, json + Environment.NewLine);
    }

    private static List<ScaffoldRoleEntry> PlanRoles(int pages, string? goal, string? audience, Random rng)
    {
        var preferred = PreferRolesFromText(goal, audience);
        var contentRoles = preferred.Concat(DefaultContentRoles).Distinct(StringComparer.Ordinal).ToArray();
        var plan = new List<ScaffoldRoleEntry>(pages) { new("cover", "cover") };

        var wantsAgenda = pages >= 8;
        var closingSlots = pages >= 10 ? 2 : 1; // actions + closing when long
        var reservedTail = closingSlots;
        var reservedHead = 1 + (wantsAgenda ? 1 : 0);
        var middleBudget = Math.Max(1, pages - reservedHead - reservedTail);

        if (wantsAgenda)
            plan.Add(new("agenda", "breakdown"));

        // Decide how many section transitions fit in the middle without starving content.
        var sectionEvery = pages >= 24 ? 5 : pages >= 16 ? 6 : pages >= 10 ? 7 : 0;
        var transitionCount = 0;
        if (sectionEvery > 0 && middleBudget >= sectionEvery + 2)
            transitionCount = Math.Max(1, (middleBudget - 1) / (sectionEvery + 1));
        var contentCount = Math.Max(1, middleBudget - transitionCount);

        var contentQueue = new List<string>(contentCount);
        for (var i = 0; i < contentCount; i++)
        {
            string role;
            if (i < 2 && preferred.Count == 0)
                role = i == 0 ? "statement" : "metrics";
            else if (preferred.Count > 0 && rng.NextDouble() < 0.55)
                role = preferred[rng.Next(preferred.Count)];
            else
                role = contentRoles[rng.Next(contentRoles.Length)];
            contentQueue.Add(role);
        }

        // Interleave transitions every sectionEvery content slides.
        var emittedContent = 0;
        var transitionsLeft = transitionCount;
        foreach (var role in contentQueue)
        {
            if (transitionsLeft > 0 && emittedContent > 0 && emittedContent % sectionEvery == 0)
            {
                plan.Add(new("transition", "transition"));
                transitionsLeft--;
            }
            plan.Add(new("content", role));
            emittedContent++;
        }
        while (transitionsLeft-- > 0)
            plan.Insert(Math.Max(1, plan.Count - 1), new("transition", "transition"));

        if (closingSlots == 2)
            plan.Add(new("content", "actions"));
        plan.Add(new("closing", "closing"));

        while (plan.Count > pages)
        {
            // Prefer dropping a middle content slide over cover/closing.
            var idx = plan.FindLastIndex(item => item.Kind == "content" && item.Role is not ("actions"));
            if (idx <= 0) idx = plan.Count - 2;
            plan.RemoveAt(Math.Clamp(idx, 1, plan.Count - 2));
        }
        while (plan.Count < pages)
            plan.Insert(plan.Count - 1, new("content", contentRoles[rng.Next(contentRoles.Length)]));

        return plan;
    }

    private static List<string> PreferRolesFromText(string? goal, string? audience)
    {
        var text = $"{goal} {audience}".ToLowerInvariant();
        var hits = new List<string>();
        void Add(string role, params string[] needles)
        {
            if (needles.Any(needle => text.Contains(needle, StringComparison.Ordinal)))
                hits.Add(role);
        }

        Add("metrics", "kpi", "metric", "增长", "revenue", "指标", "数字");
        Add("trend", "trend", "forecast", "趋势", "waterfall", "漏斗", "funnel");
        Add("comparison", "compare", "vs", "对比", "选项", "option", "swot");
        Add("process", "process", "roadmap", "流程", "里程碑", "timeline");
        Add("risks", "risk", "风险", "compliance", "合规");
        Add("team", "team", "组织", "org", "people");
        Add("case", "case", "客户", "customer", "story");
        Add("actions", "action", "next step", "下一步", "todo");
        return hits;
    }

    private static DeckLayout? PickLayout(Dictionary<string, List<DeckLayout>> byRole, string role, Random rng)
    {
        if (!byRole.TryGetValue(role, out var list) || list.Count == 0) return null;
        // Prefer shorter ids (often the canonical layout) with light randomness
        var ordered = list.OrderBy(layout => layout.Id.Length).ThenBy(layout => layout.Id, StringComparer.Ordinal).ToList();
        var pick = Math.Min(ordered.Count - 1, rng.Next(0, Math.Min(3, ordered.Count)));
        return ordered[pick];
    }

    private static IReadOnlyList<string> BuildHeuristicNotes(int pages, string? goal, string? audience)
    {
        var notes = new List<string>
        {
            pages >= 16 ? "long_deck_section_transitions" : "short_deck_linear",
            "outline_stage_empty_blocks",
            "candidates_pinned_same_role",
        };
        if (!string.IsNullOrWhiteSpace(goal)) notes.Add("goal_biased_roles");
        if (!string.IsNullOrWhiteSpace(audience)) notes.Add("audience_biased_roles");
        return notes;
    }

    private static string PlaceholderTitle(string role, int index, string language)
    {
        if (language.StartsWith("zh", StringComparison.OrdinalIgnoreCase))
            return $"{role} · 第 {index} 页";
        return $"{char.ToUpperInvariant(role[0])}{role[1..]} · slide {index}";
    }

    private static string DeriveSeed(string? goal, string? audience, int pages)
    {
        var raw = $"{goal}|{audience}|{pages}";
        var hash = SHA256.HashData(Encoding.UTF8.GetBytes(raw));
        return Convert.ToHexString(hash)[..12].ToLowerInvariant();
    }

    private static int StableHash(string seed)
    {
        var bytes = SHA256.HashData(Encoding.UTF8.GetBytes(seed));
        return BitConverter.ToInt32(bytes, 0) & int.MaxValue;
    }

    private static string SanitizeId(string role)
    {
        var chars = role.Select(ch => char.IsLetterOrDigit(ch) ? ch : '-').ToArray();
        return new string(chars).Trim('-').ToLowerInvariant();
    }

    private static string Truncate(string value, int max)
        => value.Length <= max ? value : value[..(max - 1)].TrimEnd() + "…";

    private static string InferMode(DeckTheme theme)
    {
        if (!theme.Tokens.TryGetValue("background", out var background) || string.IsNullOrWhiteSpace(background))
            return "light";
        var hex = background.Trim().TrimStart('#');
        if (hex.Length is not (6 or 8)) return "light";
        if (!int.TryParse(hex[..2], NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var r)
            || !int.TryParse(hex[2..4], NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var g)
            || !int.TryParse(hex[4..6], NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var b))
            return "light";
        var luminance = (0.2126 * r + 0.7152 * g + 0.0722 * b) / 255.0;
        return luminance < 0.45 ? "dark" : "light";
    }

    private sealed record ScaffoldRoleEntry(string Kind, string Role);
}

public sealed record DeckScaffoldRequest(
    string? Goal = null,
    string? Audience = null,
    int Pages = 12,
    string? Title = null,
    string? Language = "en-US",
    string? ThemeId = "csbu-workmate",
    string? Seed = null);

public sealed record DeckScaffoldReport(
    string Seed,
    int Pages,
    string ThemeId,
    string CatalogVersion,
    string CatalogHash,
    string GeneratedAt,
    IReadOnlyDictionary<string, int> RoleMix,
    int SectionBreaks,
    IReadOnlyList<string> Heuristics);

public sealed record DeckScaffoldResult(
    DeckSpec Spec,
    DeckScaffoldReport Report);

public sealed record DeckScaffoldCliResult(
    bool Written,
    string SpecPath,
    int SlideCount,
    string Stage,
    DeckScaffoldReport Report);

