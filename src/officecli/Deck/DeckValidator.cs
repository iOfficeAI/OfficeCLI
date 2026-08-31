// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Deck;

public static class DeckValidator
{
    public const int MaxSlides = 100;
    public const int MaxBlocksPerSlide = 64;
    public const long MaxSafeRevision = 9_007_199_254_740_991;

    private static readonly HashSet<string> BlockTypes =
        ["text", "list", "metric", "image", "chart", "table", "timeline", "quote", "shape"];
    private static readonly HashSet<string> ChartTypes =
        ["column", "bar", "line", "area", "pie", "doughnut", "scatter", "radar"];

    public static DeckValidationResult Validate(DeckSpec spec, DeckCatalog? catalog = null)
    {
        catalog ??= DeckCatalogLoader.Load();
        var diagnostics = new List<DeckDiagnostic>();

        if (spec.SchemaVersion != 1)
            diagnostics.Add(Error("unsupported_schema", "schemaVersion must be 1.", "/schemaVersion"));
        if (spec.Revision < 0 || spec.Revision > MaxSafeRevision)
            diagnostics.Add(Error("invalid_revision", $"revision must be between 0 and {MaxSafeRevision}.", "/revision"));
        if (spec.Stage is not ("outline" or "ready"))
            diagnostics.Add(Error("invalid_stage", "stage must be outline or ready.", "/stage"));
        if (string.IsNullOrWhiteSpace(spec.Metadata.Title))
            diagnostics.Add(Error("title_required", "metadata.title is required.", "/metadata/title"));
        else if (ContainsPlaceholder(spec.Metadata.Title))
            diagnostics.Add(Error("placeholder_text", "metadata.title contains an unresolved placeholder.", "/metadata/title"));
        if (spec.Metadata.AspectRatio != "16:9")
            diagnostics.Add(Error("unsupported_aspect_ratio", "Only the 16:9 aspect ratio is supported in schema v1.", "/metadata/aspectRatio"));

        var themeIds = catalog.Themes.Select(theme => theme.Id).ToHashSet(StringComparer.Ordinal);
        if (!themeIds.Contains(spec.Theme.Id))
            diagnostics.Add(Error("unknown_theme", $"Unknown theme '{spec.Theme.Id}'.", "/theme/id"));
        if (spec.Theme.Mode is not (null or "light" or "dark"))
            diagnostics.Add(Error("invalid_theme_mode", "theme.mode must be light or dark.", "/theme/mode"));
        var allowedBrandTokens = new HashSet<string>(["background", "surface", "text", "mutedText", "accent", "fontFamily"],
            StringComparer.Ordinal);
        foreach (var token in spec.Theme.BrandTokens.Keys.Where(token => !allowedBrandTokens.Contains(token)))
            diagnostics.Add(Error("unsupported_brand_token", $"Brand token '{token}' is not supported in schema v1.",
                $"/theme/brandTokens/{token}"));
        foreach (var (token, value) in spec.Theme.BrandTokens.Where(item => item.Key != "fontFamily"))
        {
            if (value.Length != 6 || value.Any(character => !Uri.IsHexDigit(character)))
                diagnostics.Add(Error("invalid_brand_color", $"Brand color '{token}' must be a six-digit RGB hex value.",
                    $"/theme/brandTokens/{token}"));
        }
        if (spec.Theme.BrandTokens.TryGetValue("fontFamily", out var fontFamily)
            && (string.IsNullOrWhiteSpace(fontFamily) || fontFamily.Length > 80))
            diagnostics.Add(Error("invalid_brand_font", "fontFamily must be a non-empty font name of at most 80 characters.",
                "/theme/brandTokens/fontFamily"));
        if (spec.Slides.Count == 0)
            diagnostics.Add(Error("slides_required", "At least one slide is required.", "/slides"));
        if (spec.Slides.Count > MaxSlides)
            diagnostics.Add(Error("slide_limit", $"A deck can contain at most {MaxSlides} slides.", "/slides"));

        var layoutById = catalog.Layouts.ToDictionary(layout => layout.Id, StringComparer.Ordinal);
        var slideIds = new HashSet<string>(StringComparer.Ordinal);
        for (var slideIndex = 0; slideIndex < spec.Slides.Count; slideIndex++)
        {
            var slide = spec.Slides[slideIndex];
            var slidePath = $"/slides/{slideIndex}";
            if (string.IsNullOrWhiteSpace(slide.Id))
                diagnostics.Add(Error("slide_id_required", "Every slide needs a stable id.", $"{slidePath}/id"));
            else if (!slideIds.Add(slide.Id))
                diagnostics.Add(Error("duplicate_slide_id", $"Duplicate slide id '{slide.Id}'.", $"{slidePath}/id", slide.Id));
            if (ContainsPlaceholder(slide.Title ?? "") || ContainsPlaceholder(slide.Notes ?? ""))
                diagnostics.Add(Error("placeholder_text", "Slide title or notes contain an unresolved placeholder.", slidePath,
                    slide.Id));
            if (!layoutById.TryGetValue(slide.LayoutId, out var layout))
                diagnostics.Add(Error("unknown_layout", $"Unknown layout '{slide.LayoutId}'.", $"{slidePath}/layoutId", slide.Id));
            if (slide.Blocks.Count > MaxBlocksPerSlide)
                diagnostics.Add(Error("block_limit", $"A slide can contain at most {MaxBlocksPerSlide} blocks.", $"{slidePath}/blocks", slide.Id));
            if (layout != null)
                ValidateControls(slide, layout, slidePath, diagnostics);

            var blockIds = new HashSet<string>(StringComparer.Ordinal);
            var occupiedSlots = new Dictionary<string, string>(StringComparer.Ordinal);
            for (var blockIndex = 0; blockIndex < slide.Blocks.Count; blockIndex++)
            {
                var block = slide.Blocks[blockIndex];
                var blockPath = $"{slidePath}/blocks/{blockIndex}";
                if (string.IsNullOrWhiteSpace(block.Id))
                    diagnostics.Add(Error("block_id_required", "Every block needs a stable id.", $"{blockPath}/id", slide.Id));
                else if (!blockIds.Add(block.Id))
                    diagnostics.Add(Error("duplicate_block_id", $"Duplicate block id '{block.Id}'.", $"{blockPath}/id", slide.Id, block.Id));
                if (!BlockTypes.Contains(block.Type))
                    diagnostics.Add(Error("unknown_block_type", $"Unknown block type '{block.Type}'.", $"{blockPath}/type", slide.Id, block.Id));
                if (string.IsNullOrWhiteSpace(block.Slot))
                {
                    if (spec.Stage == "ready")
                        diagnostics.Add(Error("block_slot_required", "Every block in a ready deck must be assigned to a layout slot.",
                            $"{blockPath}/slot", slide.Id, block.Id,
                            "Assign the block to a compatible slot or choose a layout that supports it."));
                }
                else if (layout != null)
                {
                    var slot = layout.Slots.FirstOrDefault(candidate => candidate.Id == block.Slot);
                    if (slot == null)
                        diagnostics.Add(Error("unknown_slot", $"Layout '{layout.Id}' has no slot '{block.Slot}'.", $"{blockPath}/slot", slide.Id, block.Id));
                    else
                    {
                        if (!occupiedSlots.TryAdd(slot.Id, block.Id))
                            diagnostics.Add(Error("duplicate_slot_assignment",
                                $"Slot '{slot.Id}' is already assigned to block '{occupiedSlots[slot.Id]}'.", $"{blockPath}/slot", slide.Id, block.Id,
                                "Assign each block to a different compatible slot or choose another layout."));
                        if (!slot.Accepts.Contains(block.Type, StringComparer.Ordinal))
                            diagnostics.Add(Error("slot_type_mismatch", $"Slot '{slot.Id}' does not accept block type '{block.Type}'.", $"{blockPath}/type", slide.Id, block.Id));
                        else
                            ValidateSlotCapacity(block, slot, layout, blockPath, slide.Id, diagnostics);
                    }
                }
                if (block.Type == "image" && string.IsNullOrWhiteSpace(block.AssetId))
                    diagnostics.Add(Error("image_asset_required", "Image blocks require assetId.", $"{blockPath}/assetId", slide.Id, block.Id));
                var visibleText = string.Join(' ', new[] { block.Text, block.Value, block.Label }
                    .Where(value => !string.IsNullOrWhiteSpace(value)).Concat(block.Items));
                if (ContainsPlaceholder(visibleText))
                    diagnostics.Add(Error("placeholder_text", "Content contains an unresolved placeholder.", blockPath,
                        slide.Id, block.Id, "Replace the placeholder with final content before export."));
                ValidateStructuredData(block, blockPath, slide.Id, diagnostics);
            }

            if (spec.Stage == "ready" && layout != null)
            {
                foreach (var required in layout.Slots.Where(slot => slot.Required))
                {
                    if (!slide.Blocks.Any(block => block.Slot == required.Id))
                        diagnostics.Add(Error("required_slot_empty", $"Required slot '{required.Id}' is empty.", $"{slidePath}/blocks", slide.Id,
                            suggestion: $"Add a compatible block with slot '{required.Id}'."));
                }
            }
        }

        ValidateAssets(spec, diagnostics);
        return new DeckValidationResult(!diagnostics.Any(item => item.Severity == "error"), diagnostics);
    }

    private static void ValidateAssets(DeckSpec spec, List<DeckDiagnostic> diagnostics)
    {
        var ids = new HashSet<string>(StringComparer.Ordinal);
        for (var index = 0; index < spec.Assets.Count; index++)
        {
            var asset = spec.Assets[index];
            var path = $"/assets/{index}";
            if (string.IsNullOrWhiteSpace(asset.Id) || !ids.Add(asset.Id))
                diagnostics.Add(Error("invalid_asset_id", "Asset ids must be non-empty and unique.", $"{path}/id"));
            if (asset.Type != "image")
                diagnostics.Add(Error("unsupported_asset_type", "DeckSpec v1 assets must use type 'image'.", $"{path}/type"));
            if (asset.Status is not ("pending" or "ready" or "error"))
                diagnostics.Add(Error("invalid_asset_status", "Asset status must be pending, ready, or error.", $"{path}/status"));
            if (string.IsNullOrWhiteSpace(asset.Path) || Path.IsPathRooted(asset.Path)
                || Uri.TryCreate(asset.Path, UriKind.Absolute, out _)
                || asset.Path.Split(['/', '\\'], StringSplitOptions.RemoveEmptyEntries).Contains(".."))
                diagnostics.Add(Error("unsafe_asset_path", "Asset paths must be relative and cannot contain '..'.", $"{path}/path"));
        }
        var known = spec.Assets.Select(asset => asset.Id).ToHashSet(StringComparer.Ordinal);
        foreach (var slide in spec.Slides)
        foreach (var block in slide.Blocks.Where(block => block.AssetId != null && !known.Contains(block.AssetId)))
            diagnostics.Add(Error("asset_not_found", $"Asset '{block.AssetId}' is not declared.", slideId: slide.Id, blockId: block.Id));
    }

    private static void ValidateStructuredData(DeckBlock block, string blockPath, string slideId,
        List<DeckDiagnostic> diagnostics)
    {
        if (block.Type is not ("chart" or "table")) return;
        if (!block.Data.HasValue || block.Data.Value.ValueKind != System.Text.Json.JsonValueKind.Object)
        {
            diagnostics.Add(Error("block_data_required", $"{block.Type} blocks require an object data payload.",
                $"{blockPath}/data", slideId, block.Id));
            return;
        }
        var data = block.Data.Value;
        if (block.Type == "table")
        {
            if (!data.TryGetProperty("rows", out var rows) || rows.ValueKind != System.Text.Json.JsonValueKind.Array
                || rows.GetArrayLength() == 0 || rows.EnumerateArray().Any(row => row.ValueKind != System.Text.Json.JsonValueKind.Array))
                diagnostics.Add(Error("invalid_table_data", "Table data.rows must be a non-empty array of row arrays.",
                    $"{blockPath}/data/rows", slideId, block.Id));
            else if (rows.GetArrayLength() > 100 || rows.EnumerateArray().Any(row => row.GetArrayLength() > 20))
                diagnostics.Add(Error("unsafe_table_range", "Presentation tables are limited to 100 rows and 20 columns.",
                    $"{blockPath}/data/rows", slideId, block.Id));
            return;
        }
        if (data.TryGetProperty("chartType", out var chartType)
            && (chartType.ValueKind != System.Text.Json.JsonValueKind.String
                || !ChartTypes.Contains(chartType.GetString() ?? "")))
            diagnostics.Add(Error("unsupported_chart_type",
                $"chartType must be one of: {string.Join(", ", ChartTypes.Order())}.",
                $"{blockPath}/data/chartType", slideId, block.Id));
        if (!data.TryGetProperty("categories", out var categories)
            || categories.ValueKind != System.Text.Json.JsonValueKind.Array)
            diagnostics.Add(Error("invalid_chart_categories", "Chart data.categories must be an array.",
                $"{blockPath}/data/categories", slideId, block.Id));
        if (!data.TryGetProperty("series", out var series) || series.ValueKind != System.Text.Json.JsonValueKind.Array
            || series.GetArrayLength() == 0 || series.EnumerateArray().Any(item =>
                item.ValueKind != System.Text.Json.JsonValueKind.Object
                || !item.TryGetProperty("values", out var values)
                || values.ValueKind != System.Text.Json.JsonValueKind.Array
                || values.EnumerateArray().Any(value => !value.TryGetDouble(out var number) || !double.IsFinite(number))))
            diagnostics.Add(Error("invalid_chart_series", "Chart data.series must contain named series with finite numeric values.",
                $"{blockPath}/data/series", slideId, block.Id));
        else if (data.TryGetProperty("categories", out categories)
            && categories.ValueKind == System.Text.Json.JsonValueKind.Array
            && (categories.GetArrayLength() > 500 || series.EnumerateArray().Any(item =>
                item.GetProperty("values").GetArrayLength() != categories.GetArrayLength())))
            diagnostics.Add(Error("unsafe_chart_range",
                "Chart categories and series values must have matching lengths and cannot exceed 500 points.",
                $"{blockPath}/data", slideId, block.Id));
    }

    private static void ValidateSlotCapacity(DeckBlock block, DeckSlot slot, DeckLayout layout, string blockPath,
        string slideId, List<DeckDiagnostic> diagnostics)
    {
        var textLength = new[] { block.Text, block.Value, block.Label }
            .Where(value => !string.IsNullOrEmpty(value)).Sum(value => value!.Length)
            + block.Items.Sum(item => item.Length);
        if (slot.MaxLength.HasValue && textLength > slot.MaxLength.Value)
            diagnostics.Add(Error("slot_text_overflow", $"Content exceeds slot '{slot.Id}' length guidance ({slot.MaxLength}).",
                blockPath, slideId, block.Id, LayoutSuggestion(layout)));
        if (slot.MaxItems.HasValue && block.Items.Count > slot.MaxItems.Value)
            diagnostics.Add(Error("slot_item_overflow", $"Content exceeds slot '{slot.Id}' item limit ({slot.MaxItems}).",
                $"{blockPath}/items", slideId, block.Id, LayoutSuggestion(layout)));
    }

    private static void ValidateControls(DeckSlide slide, DeckLayout layout, string slidePath,
        List<DeckDiagnostic> diagnostics)
    {
        var definitions = layout.Controls.ToDictionary(control => control.Id, StringComparer.Ordinal);
        foreach (var (id, value) in slide.Controls)
        {
            if (!definitions.TryGetValue(id, out var control))
            {
                diagnostics.Add(Error("unknown_layout_control", $"Layout '{layout.Id}' has no control '{id}'.",
                    $"{slidePath}/controls/{id}", slide.Id));
                continue;
            }
            var valid = control.Type switch
            {
                "toggle" => value.ValueKind is System.Text.Json.JsonValueKind.True or System.Text.Json.JsonValueKind.False,
                "select" => value.ValueKind == System.Text.Json.JsonValueKind.String
                    && control.Options?.Contains(value.GetString() ?? "", StringComparer.Ordinal) == true,
                "range" => value.TryGetDouble(out var number)
                    && (!control.Min.HasValue || number >= control.Min.Value)
                    && (!control.Max.HasValue || number <= control.Max.Value),
                _ => false,
            };
            if (!valid)
                diagnostics.Add(Error("invalid_layout_control", $"Control '{id}' has an invalid value.",
                    $"{slidePath}/controls/{id}", slide.Id));
        }
    }

    private static string LayoutSuggestion(DeckLayout layout) => layout.AlternativeLayoutIds is { Count: > 0 }
        ? $"Shorten the content or switch to: {string.Join(", ", layout.AlternativeLayoutIds)}."
        : "Shorten the content or split it across slides.";

    private static bool ContainsPlaceholder(string value) =>
        value.Contains("{{", StringComparison.Ordinal)
        || value.Contains("<TODO>", StringComparison.OrdinalIgnoreCase)
        || value.Contains("lorem", StringComparison.OrdinalIgnoreCase)
        || value.Contains("xxxx", StringComparison.OrdinalIgnoreCase);

    private static DeckDiagnostic Error(string code, string message, string? path = null, string? slideId = null,
        string? blockId = null, string? suggestion = null) =>
        new("error", code, message, path, slideId, blockId, suggestion);
}
