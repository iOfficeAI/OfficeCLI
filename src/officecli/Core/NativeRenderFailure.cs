// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System;

namespace OfficeCli.Core;

/// <summary>Which step of a native render failed.</summary>
internal enum NativeRenderStage
{
    /// <summary>The application could not be started (absent, unregistered, or not the expected app).</summary>
    Launch,
    /// <summary>The application started but could not open the document.</summary>
    Open,
    /// <summary>The document opened but exporting it failed or produced nothing.</summary>
    Render,
    /// <summary>The render did not finish within its time budget.</summary>
    Timeout,
    /// <summary>The application rendered the document, but it has fewer pages than the one requested.</summary>
    PageOutOfRange,
}

/// <summary>Tags an exception with the native render step it came from.</summary>
internal sealed class NativeRenderStageException(NativeRenderStage stage, Exception inner)
    : Exception(inner.Message, inner)
{
    public NativeRenderStage Stage { get; } = stage;
}

/// <summary>
/// The requested page does not exist in the application's own pagination.
/// Not a render failure to fall back from: the application's page count is the
/// authoritative one, so the request itself is out of range.
/// </summary>
internal sealed class NativePageOutOfRangeException(int totalPages)
    : Exception($"the document has {totalPages} page(s)")
{
    public int TotalPages { get; } = totalPages;
}

/// <summary>
/// Why a native render returned no image. Lets <c>--render native</c> report the
/// actual cause instead of assuming the application is not installed.
/// </summary>
internal sealed record NativeRenderFailure(NativeRenderStage Stage, string? Detail, int? TotalPages = null)
{
    public static NativeRenderFailure FromException(Exception e) => e switch
    {
        NativePageOutOfRangeException p => new(NativeRenderStage.PageOutOfRange, p.Message, p.TotalPages),
        NativeRenderStageException s => new(s.Stage, Describe(s.InnerException ?? s)),
        _ => new(NativeRenderStage.Render, Describe(e)),
    };

    /// <summary>
    /// When the native application rendered the document but the requested page
    /// is past its last page, throw the same <c>out_of_range</c> error the page
    /// bound check raises — in every render mode. Falling back to the HTML
    /// preview here would return a page the application itself does not have.
    /// </summary>
    public static void ThrowIfPageOutOfRange(NativeRenderFailure? failure, string? pageSpec)
    {
        if (failure is { Stage: NativeRenderStage.PageOutOfRange, TotalPages: { } total })
            throw new CliException($"--page {pageSpec?.Trim()} out of range (total pages: {total}).") { Code = "out_of_range" };
    }

    public static NativeRenderFailure TimedOut(int ms) => new(NativeRenderStage.Timeout, $"no result after {ms / 1000}s");

    public static NativeRenderFailure NothingRendered() => new(NativeRenderStage.Render, "no pages were produced");

    static string Describe(Exception e) => $"{e.GetType().Name}: {e.Message}";

    /// <summary>
    /// Build the error for an explicit <c>--render native</c> that produced no image.
    /// <paramref name="failure"/> is null when the native path was never attempted
    /// (non-Windows), which reports the same as a launch failure.
    /// </summary>
    public static CliException ToCliException(string app, NativeRenderFailure? failure)
    {
        var detail = failure?.Detail is { Length: > 0 } d ? $" Last error: {d}" : "";
        return (failure?.Stage ?? NativeRenderStage.Launch) switch
        {
            NativeRenderStage.Open => new CliException($"--render native: Microsoft {app} could not open the file.{detail}")
            {
                Code = "native_open_failed",
                Suggestion = $"{app} is available, so the file itself is the likely cause (locked, corrupt, or unsupported). Use --render html or --render auto.",
            },
            NativeRenderStage.Render => new CliException($"--render native: Microsoft {app} opened the file but rendering failed.{detail}")
            {
                Code = "native_render_failed",
                Suggestion = "Use --render html or --render auto.",
            },
            NativeRenderStage.Timeout => new CliException($"--render native: Microsoft {app} did not finish rendering in time.{detail}")
            {
                Code = "native_render_timeout",
                Suggestion = "Retry, or use --render html or --render auto.",
            },
            _ => new CliException($"--render native requires Windows with Microsoft {app} installed.{detail}")
            {
                Code = "native_unavailable",
                Suggestion = "Use --render html or --render auto.",
            },
        };
    }
}
