// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System;
using OfficeCli.Core.Rendering;

namespace OfficeCli.Handlers.Rendering;

/// <summary>
/// Adapts the existing in-tree HTML/SVG rendering to the <see cref="IRenderer"/> seam,
/// WITHOUT changing how rendering works. Each adapter is a thin, stateless wrapper that
/// forwards a <see cref="RenderOptions"/> to the handler's existing ViewAs* method, so
/// output is byte-identical to a direct call (covered by RendererSeamParityTests).
/// <para>
/// These register at priority 0; an alternative renderer registered above them is
/// selected instead.
/// </para>
/// </summary>
public sealed class WordBasicRenderer : IRenderer
{
    public RenderCapabilities Capabilities { get; } = new()
    {
        Name = "basic-html-docx",
        Priority = 0,
        SupportedFormats = new[] { "docx" },
        SupportedOutputs = RenderOutputKind.Html,
        SupportsWatch = true,
    };

    public bool IsAvailable => true;

    public RenderResult Render(IRenderInput input, RenderOptions options)
    {
        var h = Handler<WordHandler>(input);
        var html = h.ViewAsHtml(options.PageFilter, options.GridColumns, options.GridCellWidthPx);
        return RenderResult.Html(html);
    }

    private static T Handler<T>(IRenderInput input) where T : class
        => (input as HandlerRenderInput)?.Handler as T
           ?? throw new InvalidOperationException($"{typeof(T).Name} render input expected");
}

public sealed class ExcelBasicRenderer : IRenderer
{
    public RenderCapabilities Capabilities { get; } = new()
    {
        Name = "basic-html-xlsx",
        Priority = 0,
        SupportedFormats = new[] { "xlsx" },
        SupportedOutputs = RenderOutputKind.Html,
        SupportsWatch = true,
    };

    public bool IsAvailable => true;

    public RenderResult Render(IRenderInput input, RenderOptions options)
    {
        var h = (input as HandlerRenderInput)?.Handler as ExcelHandler
                ?? throw new InvalidOperationException("ExcelHandler render input expected");
        return RenderResult.Html(h.ViewAsHtml());
    }
}

public sealed class PptBasicRenderer : IRenderer
{
    // PowerPointHandler.ViewAsHtml's viewport default; 0 in options means "use it".
    private const int DefaultViewportPx = 1600;

    public RenderCapabilities Capabilities { get; } = new()
    {
        Name = "basic-pptx",
        Priority = 0,
        SupportedFormats = new[] { "pptx" },
        SupportedOutputs = RenderOutputKind.Html | RenderOutputKind.Svg,
        SupportsWatch = true,
    };

    public bool IsAvailable => true;

    public RenderResult Render(IRenderInput input, RenderOptions options)
    {
        var h = (input as HandlerRenderInput)?.Handler as PowerPointHandler
                ?? throw new InvalidOperationException("PowerPointHandler render input expected");

        if (options.Output == RenderOutputKind.Svg)
            return RenderResult.Svg(h.ViewAsSvg(options.StartPage ?? 1));

        var viewport = options.ViewportPx > 0 ? options.ViewportPx : DefaultViewportPx;
        var html = h.ViewAsHtml(options.StartPage, options.EndPage, options.GridColumns, viewport,
            options.IncludeSlideNumbers);
        return RenderResult.Html(html);
    }
}

/// <summary>
/// Registers the built-in basic renderers. Idempotent: registers shared singleton
/// instances, so repeated calls on the same registry are no-ops.
/// </summary>
public static class RenderingBootstrap
{
    private static readonly IRenderer Word = new WordBasicRenderer();
    private static readonly IRenderer Excel = new ExcelBasicRenderer();
    private static readonly IRenderer Ppt = new PptBasicRenderer();

    public static void EnsureRegistered(RendererRegistry? registry = null)
    {
        registry ??= RendererRegistry.Default;
        registry.Register(Word);
        registry.Register(Excel);
        registry.Register(Ppt);
        // Let any out-of-tree renderers add themselves once the built-ins are present.
        registry.RaiseComposingOnce();
    }
}
