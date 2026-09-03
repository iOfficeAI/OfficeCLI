// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Globalization;
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeCli.Handlers;

namespace OfficeCli.Deck;

public static class DeckService
{
    public const long MaxSpecBytes = 2 * 1024 * 1024;
    public const long MaxAssetBytes = 25 * 1024 * 1024;
    public const long MaxTotalAssetBytes = 200 * 1024 * 1024;

    public static DeckSpec LoadSpec(string path)
    {
        var info = new FileInfo(path);
        if (!info.Exists) throw new FileNotFoundException($"Deck spec not found: {path}", path);
        if (info.Length > MaxSpecBytes) throw new InvalidDataException("Deck spec exceeds the 2 MB limit.");
        using var stream = File.OpenRead(path);
        var spec = JsonSerializer.Deserialize(stream, DeckJsonContext.Default.DeckSpec)
            ?? throw new InvalidDataException("Deck spec is empty or invalid.");
        return NormalizeSpec(spec);
    }

    private static DeckSpec NormalizeSpec(DeckSpec spec)
    {
        var theme = spec.Theme ?? new DeckThemeSelection();
        var slides = (spec.Slides ?? []).Select(slide => slide with
        {
            Blocks = (slide.Blocks ?? []).Select(block => block with
            {
                Items = block.Items ?? [],
            }).ToList(),
            Controls = slide.Controls ?? [],
        }).ToList();

        return spec with
        {
            Metadata = spec.Metadata ?? new DeckMetadata(),
            Theme = theme with { BrandTokens = theme.BrandTokens ?? [] },
            Slides = slides,
            Assets = spec.Assets ?? [],
            Extensions = spec.Extensions ?? [],
        };
    }

    public static DeckPreviewScene RenderPreview(DeckSpec spec, string specPath)
    {
        var catalog = DeckCatalogLoader.Load();
        EnsureValid(spec, catalog, specPath);
        return CreatePreviewScene(spec, specPath, catalog);
    }

    private static DeckPreviewScene CreatePreviewScene(DeckSpec spec, string specPath, DeckCatalog catalog)
    {
        var theme = EffectiveTheme(catalog.Themes.Single(item => item.Id == spec.Theme.Id), spec.Theme);
        var layoutById = catalog.Layouts.ToDictionary(item => item.Id, StringComparer.Ordinal);
        var assetById = spec.Assets.ToDictionary(item => item.Id, StringComparer.Ordinal);
        var slides = spec.Slides.Select(slide =>
        {
            var layout = layoutById[slide.LayoutId];
            var elements = MapElements(slide, layout, assetById, specPath);
            return new DeckPreviewSlide(slide.Id, slide.LayoutId, slide.Title, slide.Hidden, elements);
        }).ToList();
        return new DeckPreviewScene(spec.Revision, theme.Id, theme.Tokens, slides);
    }

    public static string Build(DeckSpec spec, string specPath, string outputPath, long? expectedRevision = null)
    {
        if (spec.Stage != "ready")
            throw new InvalidDataException("Only a ready DeckSpec can be built as PPTX. Confirm and complete the outline first.");
        EnsureExpectedRevision(spec.Revision, expectedRevision);
        var catalog = DeckCatalogLoader.Load();
        EnsureValid(spec, catalog, specPath);
        var scene = CreatePreviewScene(spec, specPath, catalog);
        DebugDeck($"scene elements: {string.Join(", ", scene.Slides.SelectMany(slide => slide.Elements).Select(element => element.Type))}");
        var target = Path.GetFullPath(outputPath);
        var directory = Path.GetDirectoryName(target) ?? Directory.GetCurrentDirectory();
        Directory.CreateDirectory(directory);
        var temp = Path.Combine(directory, $".{Path.GetFileNameWithoutExtension(target)}.{Guid.NewGuid():N}.tmp.pptx");

        try
        {
            BlankDocCreator.Create(temp, spec.Metadata.Language);
            using (var handler = new PowerPointHandler(temp, editable: true))
            {
                var theme = EffectiveTheme(catalog.Themes.Single(item => item.Id == spec.Theme.Id), spec.Theme);
                foreach (var _ in scene.Slides)
                    handler.Add("/", "slide", null, new Dictionary<string, string> { ["layout"] = "blank" });

                for (var index = 0; index < scene.Slides.Count; index++)
                {
                    var slideNumber = index + 1;
                    var slideProps = new Dictionary<string, string>
                    {
                        ["background"] = theme.Tokens["background"],
                        ["hidden"] = scene.Slides[index].Hidden ? "true" : "false",
                    };
                    if (!string.IsNullOrWhiteSpace(scene.Slides[index].Title))
                        slideProps["name"] = scene.Slides[index].Title!;
                    handler.Set($"/slide[{slideNumber}]", slideProps);
                    foreach (var element in scene.Slides[index].Elements)
                        AddElement(handler, slideNumber, element, theme);
                    var notes = spec.Slides[index].Notes;
                    if (!string.IsNullOrWhiteSpace(notes))
                        handler.Add($"/slide[{slideNumber}]", "notes", null,
                             new Dictionary<string, string> { ["text"] = notes });
                }
                handler.Save();
            }
            using (var saved = PresentationDocument.Open(temp, false))
            {
                var slideParts = saved.PresentationPart?.SlideParts.ToList() ?? [];
                DebugDeck($"saved parts: slides={slideParts.Count}, charts={slideParts.Sum(slide => slide.Parts.Count(part => part.OpenXmlPart is ChartPart))}");
            }
            ValidateGeneratedPackage(temp);
            if (expectedRevision.HasValue)
                EnsureExpectedRevision(LoadSpec(specPath).Revision, expectedRevision);
            AtomicReplace(temp, target);
            DebugPackageParts(target, "target");
            return target;
        }
        catch
        {
            try { if (File.Exists(temp)) File.Delete(temp); } catch { }
            throw;
        }
    }

    private static void EnsureExpectedRevision(long actualRevision, long? expectedRevision)
    {
        if (expectedRevision.HasValue && actualRevision != expectedRevision.Value)
            throw new InvalidDataException(
                $"DeckSpec revision is stale: expected {expectedRevision.Value}, found {actualRevision}.");
    }

    private static void AtomicReplace(string temp, string target)
    {
        if (new FileInfo(target).LinkTarget != null)
            throw new InvalidDataException("Refusing to replace a symbolic-link PPTX target.");
        File.Move(temp, target, overwrite: true);
    }

    public static DeckValidationResult Validate(DeckSpec spec, string specPath)
    {
        var structural = DeckValidator.Validate(spec);
        var diagnostics = structural.Diagnostics.ToList();
        var referencedAssets = spec.Slides
            .SelectMany(slide => slide.Blocks)
            .Where(block => block.AssetId != null)
            .Select(block => block.AssetId!)
            .ToHashSet(StringComparer.Ordinal);
        long totalBytes = 0;
        foreach (var asset in spec.Assets)
        {
            string path;
            try { path = ResolveAssetPath(specPath, asset.Path); }
            catch (InvalidDataException)
            {
                continue;
            }
            if (asset.Status != "ready")
            {
                diagnostics.Add(new DeckDiagnostic(spec.Stage == "ready" && referencedAssets.Contains(asset.Id) ? "error" : "warning", "asset_not_ready",
                    $"Asset '{asset.Id}' is not ready.",
                    Suggestion: "Upload, generate, or replace the media before export."));
                continue;
            }
            if (!File.Exists(path))
            {
                diagnostics.Add(new DeckDiagnostic(referencedAssets.Contains(asset.Id) ? "error" : "warning", "asset_file_missing",
                    $"Asset file for '{asset.Id}' was not found."));
                continue;
            }
            var length = new FileInfo(path).Length;
            totalBytes += length;
            if (length > MaxAssetBytes)
                diagnostics.Add(new DeckDiagnostic("error", "asset_too_large", $"Asset '{asset.Id}' exceeds the 25 MB limit."));
        }
        if (totalBytes > MaxTotalAssetBytes)
            diagnostics.Add(new DeckDiagnostic("error", "asset_total_too_large", "Deck assets exceed the 200 MB total limit."));
        return new DeckValidationResult(!diagnostics.Any(item => item.Severity == "error"), diagnostics);
    }

    private static void EnsureValid(DeckSpec spec, DeckCatalog catalog, string specPath)
    {
        var structural = DeckValidator.Validate(spec, catalog);
        var validation = structural.Valid ? Validate(spec, specPath) : structural;
        if (!validation.Valid)
            throw new InvalidDataException(string.Join(Environment.NewLine,
                validation.Diagnostics.Where(item => item.Severity == "error").Select(item => $"[{item.Code}] {item.Message}")));
    }

    private static DeckTheme EffectiveTheme(DeckTheme theme, DeckThemeSelection selection)
    {
        if (selection.BrandTokens.Count == 0) return theme;
        var tokens = theme.Tokens.ToDictionary(item => item.Key, item => item.Value, StringComparer.Ordinal);
        foreach (var (key, value) in selection.BrandTokens) tokens[key] = value;
        return theme with { Tokens = tokens };
    }

    private static void ValidateGeneratedPackage(string path)
    {
        using var document = PresentationDocument.Open(path, false);
        var errors = new OpenXmlValidator().Validate(document).Take(10).ToList();
        if (errors.Count > 0)
            throw new InvalidDataException("Generated PPTX failed OOXML validation: "
                + string.Join(" | ", errors.Select(error => error.Description)));
        if (document.PresentationPart?.Presentation?.SlideIdList == null)
            throw new InvalidDataException("Generated PPTX has no slide list.");
    }

    private static IReadOnlyList<DeckPreviewElement> MapElements(
        DeckSlide slide,
        DeckLayout layout,
        IReadOnlyDictionary<string, DeckAsset> assets,
        string specPath)
    {
        var slots = layout.Slots.ToDictionary(item => item.Id, StringComparer.Ordinal);
        var occupied = new HashSet<string>(StringComparer.Ordinal);
        var elements = new List<DeckPreviewElement>();
        foreach (var block in slide.Blocks)
        {
            if (block.Slot != null && !IsSlotVisible(slide, block.Slot))
                continue;
            DeckSlot slot;
            if (block.Slot != null && slots.TryGetValue(block.Slot, out var explicitSlot))
            {
                slot = explicitSlot;
                occupied.Add(slot.Id);
            }
            else
            {
                slot = layout.Slots.FirstOrDefault(candidate =>
                           !occupied.Contains(candidate.Id)
                           && candidate.Accepts.Contains(block.Type, StringComparer.Ordinal))
                       ?? throw new InvalidDataException(
                           $"Block '{block.Id}' has no unoccupied compatible slot in layout '{layout.Id}'.");
                occupied.Add(slot.Id);
            }
            slot = AdjustSlot(layout.Id, slot, slide);
            if (slot.Width <= 0 || slot.Height <= 0)
                continue;
            var text = BlockText(block);
            string? assetPath = null;
            if (block.AssetId != null && assets.TryGetValue(block.AssetId, out var asset) && asset.Status == "ready")
                assetPath = ResolveAssetPath(specPath, asset.Path);
            var data = OverlayChartControls(slide, block);
            elements.Add(new DeckPreviewElement(block.Id, block.Type, slot.Id, slot.X, slot.Y, slot.Width, slot.Height, text, assetPath, data));
        }
        return elements;
    }

    private static DeckSlot AdjustSlot(string layoutId, DeckSlot slot, DeckSlide slide)
    {
        if ((layoutId is "image-text" or "two-column" or "cover-split" or "quote-split"
                or "cover-banner" or "image-left-bullets" or "cover-dark-band"
                or "image-quote" or "image-stats" or "cover-photo-stack"
                or "image-callout-overlay" or "image-split-caption")
            && ControlString(slide, "mediaSide", "left") == "right")
            return slot with { X = 1 - slot.X - slot.Width };

        if (!IsSlotVisible(slide, "insight"))
        {
            if ((layoutId is "chart" or "chart-radar" or "chart-insight-right"
                    or "chart-waterfall" or "chart-funnel" or "distribution-pie-focus"
                    or "result-chart-proof") && slot.Id == "chart")
                return slot with { X = 0.06, Width = 0.88 };
            if ((layoutId is "data-table" or "table-callouts") && slot.Id == "table")
                return slot with { X = 0.06, Width = 0.88 };
            if ((layoutId is "risk-matrix-simple" or "risks-matrix" or "decision-matrix") && slot.Id == "matrix")
                return slot with { X = 0.06, Width = 0.88 };
        }

        if (!IsSlotVisible(slide, "callout"))
        {
            if ((layoutId is "bullets-callout" or "comparison-score-table" or "observation-chart-note"
                    or "metrics-callout-side" or "image-callout-overlay")
                && (slot.Id is "content" or "table" or "chart" or "metric1" or "metric2" or "visual"))
            {
                if (slot.Id is "content" or "table" or "chart" or "visual")
                    return slot with { X = 0.06, Width = 0.88 };
            }
            if ((layoutId is "chart-callout-bottom" or "process-checkpoint" or "breakdown-key-points"
                    or "tradeoff-matrix-lite" or "risks-owner-table" or "actions-raci-lite"
                    or "distribution-stacked-bars" or "statement-callout" or "closing-cta-metrics"
                    or "case-three-phase" or "context-split-callout")
                && (slot.Id is "chart" or "table" or "matrix" or "body" or "content"
                    or "step1" or "step2" or "step3" or "step4" or "left" or "right"
                    or "col1" or "col2" or "col3" or "metric1" or "metric2" or "metric3"))
            {
                // keep authored geometry; callout just collapses via visibility below
            }
        }

        if (slot.Id == "insight" && !IsSlotVisible(slide, "insight"))
            return slot with { Width = 0, Height = 0 };
        if (slot.Id == "callout" && !IsSlotVisible(slide, "callout"))
            return slot with { Width = 0, Height = 0 };
        if (slot.Id == "footer" && !IsSlotVisible(slide, "footer"))
            return slot with { Width = 0, Height = 0 };
        if (slot.Id == "kicker" && !IsSlotVisible(slide, "kicker"))
            return slot with { Width = 0, Height = 0 };

        if ((layoutId is "comparison" or "two-column" or "toc" or "before-after" or "pros-cons"
                or "mitigation-plan" or "bullets-two" or "metrics-highlight" or "chart-compare"
                or "toc-two-column" or "vs-scorecard" or "statement-split" or "image-left-bullets"
                or "feature-vs" or "cost-benefit" or "kpi-vs-target" or "side-by-side-kpis"
                or "closing-split-cta" or "risks-mitigation-grid" or "result-before-after"
                or "case-study" or "case-challenge-solution" or "relationship-pairs"
                or "chart-dual-panel" or "comparison-criteria" or "ask-split-footer"
                or "risks-two-track" or "context-split-callout" or "actions-two-column"
                or "trend-dual-charts" or "process-lanes-2" or "bullets-callout")
            && (slot.Id is "left" or "right" or "kpi" or "support" or "body" or "content" or "callout"))
        {
            var balance = Math.Clamp(ControlDouble(slide, "balance", 50), 35, 65) / 100;
            const double start = 0.06;
            const double total = 0.88;
            const double gap = 0.06;
            var usable = total - gap;
            var leftWidth = usable * balance;
            var rightWidth = usable - leftWidth;
            return slot.Id is "left" or "kpi" or "content"
                ? slot with { X = start, Width = leftWidth }
                : slot with { X = start + leftWidth + gap, Width = rightWidth };
        }

        if ((layoutId is "comparison-table" or "risk" or "risk-heatmap" or "data-table"
                or "risk-matrix-simple" or "table-callouts" or "chart-insight-right"
                or "decision-matrix" or "risks-matrix" or "risks-heatmap-lite"
                or "context-brief" or "observation-quote-data" or "distribution-pie-focus"
                or "chart-waterfall" or "chart-funnel")
            && (slot.Id is "left" or "summary" or "insight" or "table" or "matrix" or "chart" or "body"))
        {
            // Insight toggle already collapsed/expanded above for table/matrix/chart companions.
            if ((layoutId is "data-table" or "risk-matrix-simple" or "table-callouts" or "chart-insight-right"
                    or "decision-matrix" or "risks-matrix" or "distribution-pie-focus"
                    or "chart-waterfall" or "chart-funnel")
                && !IsSlotVisible(slide, "insight")
                && (slot.Id is "table" or "matrix" or "chart"))
                return slot with { X = 0.06, Width = 0.88 };

            var balance = Math.Clamp(ControlDouble(slide, "balance", 35), 25, 50) / 100;
            const double start = 0.06;
            const double total = 0.88;
            const double gap = 0.04;
            var usable = total - gap;
            var leftWidth = usable * balance;
            var rightWidth = usable - leftWidth;
            if (slot.Id is "left" or "summary" or "insight")
                return slot with { X = start, Width = leftWidth };
            return slot with { X = start + leftWidth + gap, Width = rightWidth };
        }

        if ((layoutId is "swot" or "swot-compact")
            && (slot.Id is "strengths" or "weaknesses" or "opportunities" or "threats"))
        {
            var balance = Math.Clamp(ControlDouble(slide, "balance", 50), 40, 60) / 100;
            const double start = 0.06;
            const double total = 0.88;
            const double gap = 0.04;
            var usable = total - gap;
            var leftWidth = usable * balance;
            var rightWidth = usable - leftWidth;
            var isLeft = slot.Id is "strengths" or "opportunities";
            return isLeft
                ? slot with { X = start, Width = leftWidth }
                : slot with { X = start + leftWidth + gap, Width = rightWidth };
        }

        if (layoutId == "pest"
            && (slot.Id is "political" or "economic" or "social" or "technological"))
        {
            var balance = Math.Clamp(ControlDouble(slide, "balance", 50), 40, 60) / 100;
            const double start = 0.06;
            const double total = 0.88;
            const double gap = 0.04;
            var usable = total - gap;
            var leftWidth = usable * balance;
            var rightWidth = usable - leftWidth;
            var isLeft = slot.Id is "political" or "social";
            return isLeft
                ? slot with { X = start, Width = leftWidth }
                : slot with { X = start + leftWidth + gap, Width = rightWidth };
        }

        return PackModuleSlots(layoutId, slot, slide);
    }

    private static DeckSlot PackModuleSlots(string layoutId, DeckSlot slot, DeckSlide slide)
    {
        var moduleIds = layoutId switch
        {
            "metrics" or "kpi-trio" or "cover-kpi-strip" or "metrics-callout" or "kpi-radar-sidecar"
                or "chart-with-kpis" or "image-stats" or "context-facts" or "case-metrics"
                or "result-metrics" or "metrics-with-footer" or "closing-cta-metrics"
                => new[] { "metric1", "metric2", "metric3" },
            "metrics-duo" or "quote-metrics" or "result-summary" or "cover-dual-metric"
                or "metrics-callout-side" or "case-quote-result"
                => new[] { "metric1", "metric2" },
            "metrics-row-4" or "metrics-strip" or "kpi-sparkline-row" or "context-metrics-strip"
                => new[] { "metric1", "metric2", "metric3", "metric4" },
            "metrics-grid-compact" => new[] { "metric1", "metric2", "metric3", "metric4", "metric5", "metric6" },
            "cards" or "agenda-cards" or "toc-cards" or "risks-top3" or "observation-callouts"
                or "actions-priority" or "relationship-map-lite" or "breakdown-icon-row"
                or "risks-priority-cards" or "result-three-up"
                => new[] { "card1", "card2", "card3" },
            "cards-four" or "team-org-lite" or "breakdown-numbered-cards" or "breakdown-quad"
                or "observation-grid" or "relationship-hub"
                => new[] { "card1", "card2", "card3", "card4" },
            "three-column" or "comparison-three" or "breakdown-pillars" or "option-score"
                or "process-swimlane-lite" or "team-roles" or "distribution-segments"
                or "stakeholder-grid" or "team-roles-footer" or "case-three-phase"
                or "process-lanes-2"
                => new[] { "col1", "col2", "col3" },
            "option-cards-4" or "comparison-four" or "comparison-columns-4"
                or "distribution-four-seg" or "stakeholder-map-4"
                => new[] { "col1", "col2", "col3", "col4" },
            "process-steps" or "process-horizontal" or "cycle-4" or "agenda-timeline" or "roadmap-milestones"
                or "toc-timeline" or "journey-steps" or "closing-roadmap" or "context-timeline"
                or "double-diamond" or "process-checkpoint" or "case-timeline" or "chapter-progress"
                => new[] { "step1", "step2", "step3", "step4" },
            "process-vertical" => new[] { "step1", "step2", "step3", "step4" },
            "process-5" or "process-vertical-5" => new[] { "step1", "step2", "step3", "step4", "step5" },
            "process-6" => new[] { "step1", "step2", "step3", "step4", "step5", "step6" },
            "team" or "team-row" or "team-grid" => new[] { "member1", "member2", "member3", "member4" },
            "team-cards-5" or "team-lead-grid" => new[] { "member1", "member2", "member3", "member4", "member5" },
            "closing-contacts" or "closing-contacts-footer" => new[] { "member1", "member2", "member3" },
            "funnel" or "funnel-wide" or "pipeline-stages" => new[] { "stage1", "stage2", "stage3", "stage4" },
            "gallery-two" => new[] { "visual1", "visual2" },
            "gallery-three" or "image-three-up" or "gallery-caption-row"
                => new[] { "visual1", "visual2", "visual3" },
            "image-mosaic-4" => new[] { "visual1", "visual2", "visual3", "visual4" },
            "five-forces" => new[] { "rivalry", "entrants", "substitutes", "suppliers", "buyers" },
            _ => Array.Empty<string>(),
        };
        if (moduleIds.Length == 0) return slot;
        var index = Array.IndexOf(moduleIds, slot.Id);
        if (index < 0) return slot;

        // Prefer columns control for col* packs when present; else moduleCount.
        var defaultCount = moduleIds.Length;
        var countControl = moduleIds[0].StartsWith("col", StringComparison.Ordinal)
            && slide.Controls.ContainsKey("columns")
            ? "columns"
            : "moduleCount";
        var count = (int)Math.Clamp(ControlDouble(slide, countControl, defaultCount), 1, moduleIds.Length);
        var visibleIds = moduleIds.Take(count).Where(id => IsSlotVisible(slide, id)).ToArray();
        var visibleIndex = Array.IndexOf(visibleIds, slot.Id);
        if (visibleIndex < 0)
            return slot with { Width = 0, Height = 0 };
        // Keep authored geometry for non-row packs (funnel stages, 2x2 cycle, gallery).
        if (layoutId is "funnel" or "funnel-wide" or "cycle-4" or "gallery-two" or "gallery-three"
                or "process-vertical" or "cards-four" or "option-cards-4" or "comparison-four"
                or "pipeline-stages" or "image-mosaic-4" or "team-org-lite" or "five-forces"
                or "kpi-radar-sidecar" or "image-stats" or "case-metrics" or "process-5"
                or "chart-with-kpis" or "metrics-callout" or "context-facts" or "relationship-map-lite"
                or "quote-metrics" or "result-summary" or "breakdown-quad" or "observation-grid"
                or "relationship-hub" or "team-lead-grid" or "metrics-grid-compact"
                or "process-vertical-5" or "cover-dual-metric" or "case-quote-result"
                or "metrics-callout-side")
            return slot;

        var (start, total, gap) = DensityPackMetrics(slide);
        var packCount = Math.Max(1, visibleIds.Length);
        var usable = total - gap * Math.Max(0, packCount - 1);
        var widths = new double[packCount];
        var focus = (int)Math.Round(ControlDouble(slide, "focusIndex", -1));
        if (focus >= 0 && focus < packCount && packCount > 1)
        {
            // Emphasize one module (~25% weight); siblings share the remainder evenly.
            const double boost = 1.25;
            const double sibling = 1.0;
            var weightSum = boost + sibling * (packCount - 1);
            for (var i = 0; i < packCount; i++)
                widths[i] = usable * ((i == focus ? boost : sibling) / weightSum);
        }
        else
        {
            var width = usable / packCount;
            for (var i = 0; i < packCount; i++)
                widths[i] = width;
        }
        double x = start;
        for (var i = 0; i < visibleIndex; i++)
            x += widths[i] + gap;
        return slot with { X = x, Width = widths[visibleIndex] };
    }

    private static (double Start, double Total, double Gap) DensityPackMetrics(DeckSlide slide)
    {
        return ControlString(slide, "density", "comfortable") switch
        {
            "compact" => (0.04, 0.92, 0.02),
            "spacious" => (0.08, 0.84, 0.05),
            _ => (0.06, 0.88, 0.03),
        };
    }


    private static JsonElement? OverlayChartControls(DeckSlide slide, DeckBlock block)
    {
        if (block.Type != "chart" || !block.Data.HasValue || block.Data.Value.ValueKind != JsonValueKind.Object)
            return block.Data;

        var hasChartType = slide.Controls.TryGetValue("chartType", out var chartType)
            && chartType.ValueKind == JsonValueKind.String;
        var hasLegend = slide.Controls.ContainsKey("showLegend");
        var hasAxis = slide.Controls.ContainsKey("showAxisLabels");
        if (!hasChartType && !hasLegend && !hasAxis)
            return block.Data;

        using var document = JsonDocument.Parse(block.Data.Value.GetRawText());
        using var stream = new MemoryStream();
        using (var writer = new Utf8JsonWriter(stream))
        {
            writer.WriteStartObject();
            var replacedType = false;
            var replacedLegend = false;
            var replacedAxis = false;
            foreach (var property in document.RootElement.EnumerateObject())
            {
                if (hasChartType && property.NameEquals("chartType"))
                {
                    writer.WritePropertyName("chartType");
                    chartType.WriteTo(writer);
                    replacedType = true;
                }
                else if (hasLegend && property.NameEquals("legend"))
                {
                    writer.WritePropertyName("legend");
                    writer.WriteStringValue(ControlBool(slide, "showLegend", true) ? "right" : "none");
                    replacedLegend = true;
                }
                else if (hasAxis && property.NameEquals("axisVisible"))
                {
                    writer.WritePropertyName("axisVisible");
                    writer.WriteBooleanValue(ControlBool(slide, "showAxisLabels", true));
                    replacedAxis = true;
                }
                else
                {
                    property.WriteTo(writer);
                }
            }
            if (hasChartType && !replacedType)
            {
                writer.WritePropertyName("chartType");
                chartType.WriteTo(writer);
            }
            if (hasLegend && !replacedLegend)
            {
                writer.WritePropertyName("legend");
                writer.WriteStringValue(ControlBool(slide, "showLegend", true) ? "right" : "none");
            }
            if (hasAxis && !replacedAxis)
            {
                writer.WritePropertyName("axisVisible");
                writer.WriteBooleanValue(ControlBool(slide, "showAxisLabels", true));
            }
            writer.WriteEndObject();
        }
        return JsonDocument.Parse(stream.ToArray()).RootElement.Clone();
    }

    internal static string SlotVisibilityControlId(string slotId) => $"slot.{slotId}.visible";

    /// <summary>
    /// Generic per-slot visibility from slide.controls["slot.&lt;id&gt;.visible"].
    /// Falls back to legacy showInsight for the insight slot.
    /// </summary>
    internal static bool IsSlotVisible(DeckSlide slide, string slotId)
    {
        var key = SlotVisibilityControlId(slotId);
        if (slide.Controls.TryGetValue(key, out var value)
            && value.ValueKind is JsonValueKind.True or JsonValueKind.False)
            return value.GetBoolean();
        if (slotId == "insight")
            return ControlBool(slide, "showInsight", true);
        if (slotId == "callout")
            return ControlBool(slide, "showCallout", true);
        if (slotId == "footer")
            return ControlBool(slide, "showFooter", true);
        if (slotId is "kicker" or "eyebrow")
            return ControlBool(slide, "showKicker", true);
        return true;
    }

    private static string ControlString(DeckSlide slide, string id, string fallback) =>
        slide.Controls.TryGetValue(id, out var value) && value.ValueKind == JsonValueKind.String
            ? value.GetString() ?? fallback
            : fallback;

    private static bool ControlBool(DeckSlide slide, string id, bool fallback) =>
        slide.Controls.TryGetValue(id, out var value) && value.ValueKind is JsonValueKind.True or JsonValueKind.False
            ? value.GetBoolean()
            : fallback;

    private static double ControlDouble(DeckSlide slide, string id, double fallback) =>
        slide.Controls.TryGetValue(id, out var value) && value.TryGetDouble(out var number) ? number : fallback;

    private static string ResolveAssetPath(string specPath, string relativePath)
    {
        var root = Path.GetFullPath(Path.GetDirectoryName(Path.GetFullPath(specPath)) ?? Directory.GetCurrentDirectory());
        var candidate = Path.GetFullPath(Path.Combine(root, relativePath));
        var rootPrefix = root.EndsWith(Path.DirectorySeparatorChar) ? root : root + Path.DirectorySeparatorChar;
        if (!candidate.StartsWith(rootPrefix, StringComparison.OrdinalIgnoreCase))
            throw new InvalidDataException("Asset path escapes the deck directory.");
        var current = root;
        foreach (var segment in Path.GetRelativePath(root, candidate).Split(
            [Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar], StringSplitOptions.RemoveEmptyEntries))
        {
            current = Path.Combine(current, segment);
            if ((File.Exists(current) || Directory.Exists(current))
                && File.GetAttributes(current).HasFlag(FileAttributes.ReparsePoint))
                throw new InvalidDataException("Asset paths cannot traverse symbolic links or reparse points.");
        }
        return candidate;
    }

    private static string BlockText(DeckBlock block) => block.Type switch
    {
        "list" or "timeline" => string.Join("\n", block.Items.Select(item => $"• {item}")),
        "metric" => string.Join("\n", new[] { block.Value, block.Label }.Where(value => !string.IsNullOrWhiteSpace(value))),
        "quote" => string.IsNullOrWhiteSpace(block.Text) ? "" : $"“{block.Text}”",
        "chart" => block.Text ?? block.Label ?? "Chart",
        "table" => block.Text ?? block.Label ?? "Table",
        _ => block.Text ?? block.Value ?? block.Label ?? "",
    };

    private static void AddElement(PowerPointHandler handler, int slideNumber, DeckPreviewElement element, DeckTheme theme)
    {
        var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["name"] = $"wm-{element.Id}",
            ["text"] = element.Text,
            ["x"] = Cm(element.X * 33.867),
            ["y"] = Cm(element.Y * 19.05),
            ["width"] = Cm(element.Width * 33.867),
            ["height"] = Cm(element.Height * 19.05),
            ["font"] = theme.Tokens["fontFamily"],
            ["font.color"] = theme.Tokens["text"],
            ["font.size"] = element.Type == "metric" ? "30pt" : "18pt",
            ["margin"] = "0.12cm",
            ["autofit"] = "shrink",
        };
        var slidePath = $"/slide[{slideNumber}]";
        if (element.Type == "chart")
        {
            AddChart(handler, slidePath, element, props);
            return;
        }
        if (element.Type == "table")
        {
            AddTable(handler, slidePath, element, props);
            return;
        }
        if (element.Type == "image" && element.AssetPath != null && File.Exists(element.AssetPath))
        {
            props.Remove("text");
            props["src"] = element.AssetPath;
            props["alt"] = element.Text;
            props["fit"] = "cover";
            handler.Add(slidePath, "picture", null, props);
            return;
        }
        if (element.Type is "metric" or "quote" or "shape")
        {
            props["preset"] = "roundRect";
            props["fill"] = theme.Tokens["surface"];
            props["line"] = "none";
        }
        handler.Add(slidePath, "shape", null, props);
    }

    private static void AddChart(PowerPointHandler handler, string slidePath, DeckPreviewElement element,
        Dictionary<string, string> props)
    {
        var data = RequireObjectData(element);
        var categories = RequireArray(data, "categories").Select(JsonText).ToArray();
        var series = RequireArray(data, "series").Select(item =>
        {
            if (item.ValueKind != JsonValueKind.Object)
                throw new InvalidDataException($"Chart block '{element.Id}' series entries must be objects.");
            var name = item.TryGetProperty("name", out var nameValue) ? JsonText(nameValue) : "Series";
            var values = RequireArray(item, "values").Select(value =>
            {
                if (!value.TryGetDouble(out var number) || !double.IsFinite(number))
                    throw new InvalidDataException($"Chart block '{element.Id}' contains a non-finite value.");
                return number.ToString(CultureInfo.InvariantCulture);
            });
            return $"{SanitizeChartToken(name)}:{string.Join(',', values)}";
        }).ToArray();
        if (series.Length == 0)
            throw new InvalidDataException($"Chart block '{element.Id}' requires at least one series.");
        props.Remove("text");
        props.Remove("font");
        props.Remove("font.color");
        props.Remove("font.size");
        props.Remove("margin");
        props.Remove("autofit");
        props["chartType"] = data.TryGetProperty("chartType", out var chartType) ? JsonText(chartType) : "column";
        props["categories"] = string.Join(',', categories.Select(SanitizeChartToken));
        props["data"] = string.Join(';', series);
        if (!string.IsNullOrWhiteSpace(element.Text)) props["title"] = element.Text;
        if (data.TryGetProperty("legend", out var legendProp))
            props["legend"] = JsonText(legendProp);
        if (data.TryGetProperty("axisVisible", out var axisProp))
            props["axisVisible"] = axisProp.ValueKind is JsonValueKind.True or JsonValueKind.False
                ? (axisProp.GetBoolean() ? "true" : "false")
                : JsonText(axisProp);
        var path = handler.Add(slidePath, "chart", null, props);
        DebugDeck($"added chart '{element.Id}' at {path}");
    }

    private static void AddTable(PowerPointHandler handler, string slidePath, DeckPreviewElement element,
        Dictionary<string, string> props)
    {
        var data = RequireObjectData(element);
        var rows = RequireArray(data, "rows");
        if (rows.Count == 0)
            throw new InvalidDataException($"Table block '{element.Id}' requires at least one row.");
        var encodedRows = rows.Select(row =>
        {
            if (row.ValueKind != JsonValueKind.Array)
                throw new InvalidDataException($"Table block '{element.Id}' rows must be arrays.");
            return string.Join(',', row.EnumerateArray().Select(value => CsvCell(JsonText(value))));
        });
        props.Remove("text");
        props.Remove("font");
        props.Remove("font.color");
        props.Remove("font.size");
        props.Remove("margin");
        props.Remove("autofit");
        props["data"] = string.Join(';', encodedRows);
        var path = handler.Add(slidePath, "table", null, props);
        DebugDeck($"added table '{element.Id}' at {path}");
    }

    private static JsonElement RequireObjectData(DeckPreviewElement element)
    {
        if (!element.Data.HasValue || element.Data.Value.ValueKind != JsonValueKind.Object)
            throw new InvalidDataException($"{element.Type} block '{element.Id}' requires an object data payload.");
        return element.Data.Value;
    }

    private static List<JsonElement> RequireArray(JsonElement parent, string property)
    {
        if (!parent.TryGetProperty(property, out var value) || value.ValueKind != JsonValueKind.Array)
            throw new InvalidDataException($"Property '{property}' must be an array.");
        return value.EnumerateArray().ToList();
    }

    private static string JsonText(JsonElement value) => value.ValueKind == JsonValueKind.String
        ? value.GetString() ?? ""
        : value.ToString();

    private static string SanitizeChartToken(string value) => value.Replace(',', ' ').Replace(';', ' ').Replace(':', ' ');

    private static string CsvCell(string value) => value.IndexOfAny([',', ';', '"']) >= 0
        ? $"\"{value.Replace("\"", "\"\"")}\""
        : value;

    private static string Cm(double value) => value.ToString("0.###", CultureInfo.InvariantCulture) + "cm";

    private static void DebugDeck(string message)
    {
        if (Environment.GetEnvironmentVariable("OFFICECLI_DECK_DEBUG") == "1")
            Console.Error.WriteLine($"[deck-debug] {message}");
    }

    private static void DebugPackageParts(string path, string label)
    {
        if (Environment.GetEnvironmentVariable("OFFICECLI_DECK_DEBUG") != "1") return;
        using var document = PresentationDocument.Open(path, false);
        var slideParts = document.PresentationPart?.SlideParts.ToList() ?? [];
        var charts = slideParts.SelectMany(slide => slide.Parts)
            .Where(part => part.OpenXmlPart is ChartPart)
            .Select(part => part.OpenXmlPart.Uri.ToString())
            .ToList();
        DebugDeck($"{label} parts: slides={slideParts.Count}, charts={charts.Count}, chartUris={string.Join(",", charts)}");
    }
}
