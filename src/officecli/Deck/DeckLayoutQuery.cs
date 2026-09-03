// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Deck;

/// <summary>
/// Capacity-aware layout ranking over the embedded WorkMate catalog.
/// Original CSBU WorkMate implementation — inspired by the layout:query workflow idea only.
/// </summary>
public static class DeckLayoutQuery
{
    public static DeckLayoutQueryResult Query(DeckLayoutQueryRequest request, DeckCatalog? catalog = null)
    {
        ArgumentNullException.ThrowIfNull(request);
        catalog ??= DeckCatalogLoader.Load();
        var limit = Math.Clamp(request.Limit <= 0 ? 8 : request.Limit, 1, 50);
        var roleFilter = string.IsNullOrWhiteSpace(request.Role) ? null : request.Role.Trim();
        var textQuery = string.IsNullOrWhiteSpace(request.Query) ? null : request.Query.Trim();

        var scored = new List<DeckLayoutQueryHit>();
        foreach (var layout in catalog.Layouts)
        {
            if (roleFilter != null
                && !string.Equals(layout.Role, roleFilter, StringComparison.OrdinalIgnoreCase))
                continue;

            var capacity = EstimateCapacity(layout);
            var accepts = layout.Slots
                .SelectMany(slot => slot.Accepts)
                .Distinct(StringComparer.Ordinal)
                .OrderBy(value => value, StringComparer.Ordinal)
                .ToList();
            var reasons = new List<string>();
            double score = 0;

            if (roleFilter != null)
            {
                score += 20;
                reasons.Add("role_match");
            }

            if (request.ItemCount is int itemCount && itemCount > 0)
            {
                var target = Math.Max(1, itemCount);
                var delta = Math.Abs(capacity - target);
                var capacityScore = Math.Max(0, 10 - delta * 2);
                score += capacityScore;
                if (delta == 0) reasons.Add("capacity_exact");
                else if (delta == 1) reasons.Add("capacity_near");
                else if (capacityScore > 0) reasons.Add("capacity_partial");
                else reasons.Add("capacity_mismatch");
            }

            if (request.HasChart is bool wantsChart)
            {
                var acceptsChart = accepts.Contains("chart", StringComparer.Ordinal);
                if (wantsChart)
                {
                    score += acceptsChart ? 6 : -4;
                    reasons.Add(acceptsChart ? "chart_ok" : "chart_missing");
                }
                else if (acceptsChart)
                {
                    // Mild preference for non-chart layouts when chart is explicitly not needed.
                    score -= 1;
                    reasons.Add("chart_unneeded");
                }
            }

            if (request.NeedsMedia is bool wantsMedia)
            {
                var acceptsImage = accepts.Contains("image", StringComparer.Ordinal);
                if (wantsMedia)
                {
                    score += acceptsImage ? 5 : -3;
                    reasons.Add(acceptsImage ? "media_ok" : "media_missing");
                }
                else if (acceptsImage)
                {
                    score -= 0.5;
                    reasons.Add("media_unneeded");
                }
            }

            if (request.HasTable is bool wantsTable)
            {
                var acceptsTable = accepts.Contains("table", StringComparer.Ordinal);
                if (wantsTable)
                {
                    score += acceptsTable ? 4 : -2;
                    reasons.Add(acceptsTable ? "table_ok" : "table_missing");
                }
            }

            if (textQuery != null)
            {
                var haystack = $"{layout.Id} {layout.Label} {layout.Role}";
                if (haystack.Contains(textQuery, StringComparison.OrdinalIgnoreCase))
                {
                    score += 3;
                    reasons.Add("text_match");
                }
                else
                {
                    // Soft-filter: keep but do not boost; still return within role.
                    score -= 0.25;
                }
            }

            if (layout.AlternativeLayoutIds is { Count: > 0 })
            {
                score += 1;
                reasons.Add("has_alternatives");
            }

            // Prefer denser, better-labeled catalog entries slightly for stable ordering.
            score += Math.Min(2, layout.Controls.Count * 0.1);

            scored.Add(new DeckLayoutQueryHit(
                LayoutId: layout.Id,
                Role: layout.Role,
                Label: layout.Label,
                Score: Math.Round(score, 2),
                Capacity: capacity,
                Accepts: accepts,
                Reasons: reasons,
                AlternativeLayoutIds: layout.AlternativeLayoutIds ?? []));
        }

        var results = scored
            .OrderByDescending(hit => hit.Score)
            .ThenBy(hit => hit.LayoutId, StringComparer.Ordinal)
            .Take(limit)
            .ToList();

        return new DeckLayoutQueryResult(
            Query: new DeckLayoutQueryRequest(
                Role: roleFilter,
                ItemCount: request.ItemCount,
                HasChart: request.HasChart,
                NeedsMedia: request.NeedsMedia,
                HasTable: request.HasTable,
                Query: textQuery,
                Limit: limit),
            CatalogVersion: catalog.Version,
            CatalogHash: catalog.Hash,
            ResultCount: results.Count,
            Results: results);
    }

    /// <summary>
    /// Best-effort module capacity from catalog metadata (moduleCount control max, else content slots).
    /// </summary>
    public static int EstimateCapacity(DeckLayout layout)
    {
        var moduleCount = layout.Controls.FirstOrDefault(control =>
            string.Equals(control.Id, "moduleCount", StringComparison.Ordinal));
        if (moduleCount?.Max is double max && max >= 1)
            return (int)Math.Round(max);

        if (moduleCount != null && moduleCount.DefaultValue.ValueKind == System.Text.Json.JsonValueKind.Number
            && moduleCount.DefaultValue.TryGetDouble(out var defaultValue) && defaultValue >= 1)
            return (int)Math.Round(defaultValue);

        var contentSlots = layout.Slots.Count(slot =>
            slot.Id is not ("title" or "subtitle" or "notes" or "kicker"));
        return Math.Max(1, contentSlots);
    }
}
