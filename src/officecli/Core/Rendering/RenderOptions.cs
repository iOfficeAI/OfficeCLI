// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Core.Rendering;

/// <summary>
/// A render request. Unifies the per-handler entry points that exist today
/// (WordHandler.ViewAsHtml(pageFilter, gridCols, gridCellWpx),
/// PowerPointHandler.ViewAsHtml(start, end, gridCols, viewportPx, includeSlideNumbers),
/// ExcelHandler.ViewAsHtml(), FormatHandlerProxy.ViewAsHtml(page)) into a single
/// options object so dispatch no longer branches on the concrete handler type.
/// <para>
/// Unset fields mean "renderer default". A renderer ignores options it does not
/// support (e.g. Excel ignores paging/grid) rather than erroring.
/// </para>
/// </summary>
public sealed class RenderOptions
{
    /// <summary>Which artifact to produce. Exactly one kind.</summary>
    public RenderOutputKind Output { get; init; } = RenderOutputKind.Html;

    /// <summary>Static (view/export) vs Watch (live preview with interaction contract).</summary>
    public RenderMode Mode { get; init; } = RenderMode.Static;

    /// <summary>First page to render (1-based, inclusive). Null = from the start.</summary>
    public int? StartPage { get; init; }

    /// <summary>Last page to render (1-based, inclusive). Null = to the end.</summary>
    public int? EndPage { get; init; }

    /// <summary>Handler-specific page filter string (Word's <c>pageFilter</c>). Null = all.
    /// When both this and Start/End are set, the renderer prefers the explicit range.</summary>
    public string? PageFilter { get; init; }

    /// <summary>Number of columns when laying pages out in a grid. 0 = none/auto.</summary>
    public int GridColumns { get; init; }

    /// <summary>Grid cell width in CSS px. 0 = renderer default.</summary>
    public int GridCellWidthPx { get; init; }

    /// <summary>Viewport width in px for slide-style scaling (pptx). 0 = renderer default.</summary>
    public int ViewportPx { get; init; }

    /// <summary>Whether PPTX HTML renders slide-number placeholders and slidenum fields.</summary>
    public bool IncludeSlideNumbers { get; init; } = true;

    /// <summary>Target raster width in px for Png/Pdf output. 0 = renderer default.</summary>
    public int RasterWidthPx { get; init; }

    /// <summary>Target raster height in px for Png/Pdf output. 0 = renderer default.</summary>
    public int RasterHeightPx { get; init; }
}
