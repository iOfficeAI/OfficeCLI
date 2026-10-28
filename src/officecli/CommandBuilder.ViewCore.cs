// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using OfficeCli.Core;

namespace OfficeCli;

/// <summary>Parsed arguments of `view`, shared by the CLI front-end and the resident.</summary>
internal sealed record ViewArgs(
    string Mode, int? Start, int? End, int? MaxLines, string? IssueType, int? Limit, string? ColsStr,
    string? PageFilter, bool Browser, string? OutArg, string? ClipArg, int ScreenshotWidth, int ScreenshotHeight,
    int GridCols, string RenderMode, bool WithPages, bool Json);

/// <summary>
/// How a native (Word / PowerPoint) render gets at the file while a resident
/// holds it: the resident's editable handler owns a write handle that blocks
/// the Office app, so it is released before the render and reopened after.
/// The CLI front-end passes null (its handler is read-only and coexists).
/// </summary>
internal sealed class NativeRenderLock
{
    public required Action Release { get; init; }
    public required Func<IDocumentHandler> Reopen { get; init; }
}

static partial class CommandBuilder
{
    /// <summary>
    /// The `view` command proper — every mode (text, annotated, outline,
    /// stats, issues, forms, html, svg, screenshot) for an already-open
    /// handler. This is the single implementation: the CLI front-end calls
    /// it after option parsing, and the resident calls it with the request's
    /// arguments, so the two paths cannot drift (issue #485). `pdf` is
    /// handled by the front-end alone (exporter plugin, no handler).
    /// </summary>
    /// <summary>A produced file's path: bare on stdout, or in the read envelope under --json.</summary>
    private static void EmitPathOrText(string text, bool json)
    {
        if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeText(text));
        else Console.WriteLine(text);
    }

    internal static int RunView(IDocumentHandler handler, string filePath, ViewArgs a, NativeRenderLock? nativeLock)
    {
        var json = a.Json;
        var mode = a.Mode;
        var start = a.Start; var end = a.End; var maxLines = a.MaxLines;
        var issueType = a.IssueType; var limit = a.Limit; var colsStr = a.ColsStr;
        var pageFilter = a.PageFilter; var browser = a.Browser; var outArg = a.OutArg; var clipArg = a.ClipArg;
        var screenshotWidth = a.ScreenshotWidth; var screenshotHeight = a.ScreenshotHeight;
        var gridCols = a.GridCols; var renderMode = a.RenderMode; var withPages = a.WithPages;

        // Line windows and limits are 1-based counts; zero, negatives and an
        // inverted window used to be accepted and quietly produce nothing.
        if (start is <= 0 || end is <= 0)
            throw new OfficeCli.Core.CliException($"Invalid --start/--end: line numbers are 1-based and must be >= 1.") { Code = "invalid_value" };
        if (start.HasValue && end.HasValue && end.Value < start.Value)
            throw new OfficeCli.Core.CliException($"Invalid --start {start}/--end {end}: end is before start.") { Code = "invalid_value", Suggestion = $"Use --start {end} --end {start}." };
        if (maxLines is <= 0)
            throw new OfficeCli.Core.CliException("Invalid --max-lines: must be >= 1.") { Code = "invalid_value" };
        if (limit is <= 0)
            throw new OfficeCli.Core.CliException("Invalid --limit: must be >= 1.") { Code = "invalid_value" };
        if (screenshotWidth <= 0 || screenshotHeight <= 0)
            throw new OfficeCli.Core.CliException($"Invalid screenshot size {screenshotWidth}x{screenshotHeight}: width and height must be positive.")
            { Code = "invalid_value", Suggestion = "Omit the options for the default size, or pass positive pixel values." };
        // xlsx renders the active sheet; --page / --grid have no meaning there
        // and were silently ignored.
        if (handler is OfficeCli.Handlers.ExcelHandler && mode.ToLowerInvariant() is "html" or "h" or "screenshot" or "p"
            && (!string.IsNullOrEmpty(pageFilter) || gridCols != 0))
            throw new OfficeCli.Core.CliException("--page and --grid are not supported for xlsx (the active sheet is rendered).")
            { Code = "invalid_value", Suggestion = "Use --range Sheet!A1:C3 to capture part of a sheet." };

        var format = json ? OutputFormat.Json : OutputFormat.Text;
        var cols = colsStr != null ? new HashSet<string>(colsStr.Split(',').Select(c => c.Trim().ToUpperInvariant())) : null;


        if (mode.ToLowerInvariant() is "html" or "h")
        {
            string? html = null;
            if (handler is OfficeCli.Handlers.PowerPointHandler pptHandler)
            {
                // BUG-R36-B7: --page on pptx html previously fell through to
                // start/end via the parser default (no value), so --page 99
                // silently rendered all slides. Honor --page with strict
                // range checking, matching SVG mode's CONSISTENCY(strict-page).
                var (pStart, pEnd) = ParsePptHtmlPage(pageFilter, start, end, pptHandler);
                html = RenderViaRegistry(handler, "pptx",
                    new OfficeCli.Core.Rendering.RenderOptions { StartPage = pStart, EndPage = pEnd });
            }
            else if (handler is OfficeCli.Handlers.ExcelHandler)
                html = RenderViaRegistry(handler, "xlsx", new OfficeCli.Core.Rendering.RenderOptions());
            else if (handler is OfficeCli.Handlers.WordHandler)
            {
                ValidateDocxPageSpec(pageFilter);
                html = RenderViaRegistry(handler, "docx",
                    new OfficeCli.Core.Rendering.RenderOptions { PageFilter = pageFilter });
            }
            else if (handler is OfficeCli.Core.Plugins.FormatHandlerProxy proxy)
                html = proxy.ViewAsHtml(int.TryParse(pageFilter, out var p) ? p : (int?)null);

            if (html != null)
            {
                if (outArg != null || browser)
                {
                    // --out: write to the requested path. --browser without --out:
                    // write to a temp file and open it. With both, write to --out
                    // and open that.
                    // SECURITY: when falling back to a temp file, include a random
                    // token so the preview path is not predictable. A predictable
                    // path (HHmmss only) lets a local attacker pre-place a symlink at
                    // the expected location, causing File.WriteAllText to follow it
                    // and overwrite an arbitrary victim file with preview HTML. It
                    // also caused collisions between concurrent `view html`
                    // invocations of the same file.
                    var htmlPath = outArg ?? Path.Combine(Path.GetTempPath(), $"officecli_preview_{Path.GetFileNameWithoutExtension(Path.GetFileName(filePath))}_{DateTime.Now:HHmmss}_{Guid.NewGuid():N}.html");
                    File.WriteAllText(htmlPath, html);
                    EmitPathOrText(Path.GetFullPath(htmlPath), json);
                    if (browser)
                    {
                        try
                        {
                            var psi = new System.Diagnostics.ProcessStartInfo(htmlPath) { UseShellExecute = true };
                            System.Diagnostics.Process.Start(psi);
                        }
                        catch { /* silently ignore if browser can't be opened */ }
                    }
                }
                else
                {
                    // Default: output HTML to stdout (--json: in the read envelope).
                    if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeText(html));
                    else Console.WriteLine(html.TrimEnd('\r', '\n')); // exactly one trailing newline, as the resident returns it
                }
            }
            else
            {
                throw new OfficeCli.Core.CliException("HTML preview is only supported for .pptx, .xlsx, and .docx files.")
                {
                    Code = "unsupported_type",
                    Suggestion = "Use a .pptx, .xlsx, or .docx file, or use mode 'text' or 'annotated' for other formats.",
                    ValidValues = ["text", "annotated", "outline", "stats", "issues"]
                };
            }
            return 0;
        }

        if (mode.ToLowerInvariant() is "screenshot" or "p")
        {
            // Screenshot mode: render the same HTML preview as `view html`, then
            // headless-screenshot the temp HTML to a PNG. Mirrors svg's pattern of
            // a dedicated mode that produces a file + prints the path.
            // --grid N tiles slides into an N-column thumbnail grid (pptx only).
            //
            // CONSISTENCY(screenshot-default-first-page): screenshot mode defaults
            // to a single bounded visual unit (pptx → slide 1, docx → page 1, xlsx
            // → active sheet). Without this, multi-slide/multi-page docs render
            // the full HTML stacked vertically and get silently cropped by the
            // viewport height (default 1200) — a footgun. To capture all
            // slides/pages, use --page explicitly (e.g. --page 1-N) or --grid N
            // for pptx thumbnails. xlsx is naturally first-sheet via CSS
            // `.sheet-content { display:none }` + `.active` on sheet 0.
            string? html = null;
            byte[]? directPng = null;
            // --clip captures a data-path-addressed region out of the HTML
            // preview, so the native/direct-PNG backends (whole pages only)
            // are bypassed. pptx: the clip's slide ordinal selects the page
            // when --page wasn't given; docx: all pages render so the target
            // can sit anywhere (--page still narrows explicitly).
            if (clipArg != null)
                renderMode = "html";
            if (handler is OfficeCli.Handlers.PowerPointHandler pptHandler)
            {
                var effectiveFilter = pageFilter;
                if (clipArg != null && string.IsNullOrEmpty(effectiveFilter)
                    && System.Text.RegularExpressions.Regex.Match(clipArg, @"^/slide\[(\d+)\]") is { Success: true } slideM)
                    effectiveFilter = slideM.Groups[1].Value;
                if (string.IsNullOrEmpty(effectiveFilter) && start is null && end is null && gridCols == 0)
                    effectiveFilter = "1";
                var (pStart, pEnd) = ParsePptHtmlPage(effectiveFilter, start, end, pptHandler);

                // Native path (mirrors docx --render auto/native/html): on Windows
                // with the presentation app installed, export the slide(s) to PNG
                // with the OS-native engine. Grid mode is HTML-only, so native is
                // skipped when --grid is set. The slide's 96-DPI native pixels are
                // the default export size; a custom --screenshot-width overrides it
                // (aspect-matched height). A range stacks vertically.
                var (nativeW, nativeH) = pptHandler.GetSlideNativePixels();
                // -1 = auto: pick columns from the slide count + slide aspect so the
                // composed contact sheet is ≈ square (landscape slides → fewer cols).
                int gridColsResolved = gridCols < 0
                    ? OfficeCli.Core.HtmlScreenshot.AutoGridColumns((pEnd ?? pptHandler.GetSlideCount()) - (pStart ?? 1) + 1, nativeW, nativeH)
                    : gridCols;
                int exportW = nativeW, exportH = nativeH;
                if (!(screenshotWidth == 1600 && screenshotHeight == 1200))
                {
                    exportW = screenshotWidth;
                    exportH = screenshotHeight == 1200
                        ? Math.Max(1, (int)Math.Round(screenshotWidth * (double)nativeH / nativeW))
                        : screenshotHeight;
                }
                OfficeCli.Core.NativeRenderFailure? nativeFailure = null;
                if (renderMode != "html" && OperatingSystem.IsWindows())
                {
                    // A resident holds a write handle that blocks the Office app:
                    // release it for the native render and reopen afterwards.
                    nativeLock?.Release();
                    try
                    {
                        try
                        {
                            if (gridColsResolved > 0)
                            {
                                const int gap = 12, pad = 12;
                                int cellW = Math.Max(1, (int)Math.Round((screenshotWidth - 2 * pad - (gridColsResolved - 1) * gap) / (double)gridColsResolved));
                                int cellH = Math.Max(1, (int)Math.Round(cellW * (double)nativeH / nativeW));
                                directPng = OfficeCli.Core.PowerPointPngBackend.RenderGrid(filePath, pStart ?? 1, pEnd ?? pptHandler.GetSlideCount(), cellW, cellH, gridColsResolved, gap, pad, out nativeFailure);
                            }
                            else
                            {
                                directPng = OfficeCli.Core.PowerPointPngBackend.Render(filePath, pStart ?? 1, pEnd ?? pStart ?? 1, exportW, exportH, out nativeFailure);
                            }
                        }
                        catch (Exception ex) { directPng = null; nativeFailure = OfficeCli.Core.NativeRenderFailure.FromException(ex); }
                }
                    finally
                    {
                        if (nativeLock != null) { handler = nativeLock.Reopen(); pptHandler = (OfficeCli.Handlers.PowerPointHandler)handler; }
                    }
                }
                if (renderMode == "native" && directPng == null)
                    throw OfficeCli.Core.NativeRenderFailure.ToCliException("PowerPoint", nativeFailure);

                if (directPng == null)
                {
                    // Sheet (range / grid) and single-slide viewport sizing are
                    // shared with the resident path — see PptHtmlScreenshotViewport.
                    int htmlGridCols = PptHtmlScreenshotViewport(gridCols, gridColsResolved, pStart, pEnd,
                        pptHandler.GetSlideCount(), nativeW, nativeH, ref screenshotWidth, ref screenshotHeight);
                    html = RenderViaRegistry(pptHandler, "pptx", new OfficeCli.Core.Rendering.RenderOptions
                    { StartPage = pStart, EndPage = pEnd, GridColumns = htmlGridCols, ViewportPx = screenshotWidth })!;
                }
            }
            else if (handler is OfficeCli.Handlers.ExcelHandler excelHandler)
                html = RenderViaRegistry(excelHandler, "xlsx", new OfficeCli.Core.Rendering.RenderOptions { CellRange = clipArg })!;
            else if (handler is OfficeCli.Handlers.WordHandler wordHandlerGrid && DocxWantsSheet(gridCols, pageFilter))
            {
                // Contact sheet: --grid N/auto over the whole document, or a
                // multi-page --page range (one column of just those pages).
                // Layout maths are shared with the resident path — see
                // DocxSheetLayout. Native-first on Windows with Word; the HTML
                // preview grid is the fallback.
                var sheet = DocxSheetLayout(wordHandlerGrid, pageFilter, gridCols, screenshotWidth, screenshotHeight);
                OfficeCli.Core.NativeRenderFailure? nativeFailure = null;
                if (renderMode != "html" && OperatingSystem.IsWindows())
                {
                    // A resident holds a write handle that blocks the Office app:
                    // release it for the native render and reopen afterwards.
                    nativeLock?.Release();
                    try
                    {
                        try { directPng = OfficeCli.Core.WordPdfBackend.RenderGrid(filePath, sheet.NativePageSpec, sheet.CellWpx, sheet.CellHpx, sheet.Cols, DocxSheetGap, DocxSheetPad, out nativeFailure); }
                        catch (Exception ex) { directPng = null; nativeFailure = OfficeCli.Core.NativeRenderFailure.FromException(ex); }
                }
                    finally
                    {
                        if (nativeLock != null) { handler = nativeLock.Reopen(); wordHandlerGrid = (OfficeCli.Handlers.WordHandler)handler; }
                    }
                }
                if (renderMode == "native" && directPng == null)
                    throw OfficeCli.Core.NativeRenderFailure.ToCliException("Word", nativeFailure);
                if (directPng == null)
                {
                    // HTML fallback: layoutGrid tiles in-browser; size the viewport
                    // to fit the rows so window-size backends don't crop.
                    html = RenderViaRegistry(wordHandlerGrid, "docx", sheet.HtmlOptions)!;
                    (screenshotWidth, screenshotHeight) = (sheet.ViewportW, sheet.ViewportH);
                }
            }
            else if (handler is OfficeCli.Handlers.WordHandler wordHandler)
            {
                // --clip: render every page (filter stays null) unless the
                // caller narrowed with --page — the target may be anywhere.
                var effectiveFilter = clipArg != null
                    ? pageFilter
                    : (string.IsNullOrEmpty(pageFilter) ? "1" : pageFilter);
                ValidateDocxPageSpec(effectiveFilter);
                // A page past the end used to come back as a blank image with
                // exit 0. Page 1 always exists; anything else is checked against
                // the preview's page count.
                // (`N` and the one-page range `N-N` both land here.)
                var singleSel = ParseDocxPageSelection(effectiveFilter, int.MaxValue);
                // Normalise `N-N` to `N` so the page is framed exactly like a
                // bare page number (native-size viewport, no wrapper chrome).
                if (singleSel.Count == 1) effectiveFilter = singleSel[0].ToString();
                if (singleSel.Count == 1 && singleSel[0] > 1)
                {
                    var total = DocxPreviewPageCount(wordHandler);
                    if (total.HasValue && singleSel[0] > total.Value)
                        throw new OfficeCli.Core.CliException($"--page {singleSel[0]} out of range (total pages: {total.Value}).") { Code = "out_of_range" };
                }
                OfficeCli.Core.NativeRenderFailure? nativeFailure = null;
                if (renderMode != "html" && OperatingSystem.IsWindows())
                {
                    // A resident holds a write handle that blocks the Office app:
                    // release it for the native render and reopen afterwards.
                    nativeLock?.Release();
                    try
                    {
                        // effectiveFilter is only null under --range, which forces
                        // renderMode=html — this native branch is then unreachable.
                        try { directPng = OfficeCli.Core.WordPdfBackend.Render(filePath, effectiveFilter!, out nativeFailure); }
                        catch (Exception ex) { directPng = null; nativeFailure = OfficeCli.Core.NativeRenderFailure.FromException(ex); }
                }
                    finally
                    {
                        if (nativeLock != null) { handler = nativeLock.Reopen(); wordHandler = (OfficeCli.Handlers.WordHandler)handler; }
                    }
                }
                if (renderMode == "native" && directPng == null)
                    throw OfficeCli.Core.NativeRenderFailure.ToCliException("Word", nativeFailure);
                if (directPng == null)
                {
                    html = RenderViaRegistry(wordHandler, "docx",
                        new OfficeCli.Core.Rendering.RenderOptions { PageFilter = effectiveFilter })!;

                    // HTML-path screenshot of a single page: size the viewport to the
                    // page's 96-DPI native pixels so the PNG is the page, padding-free
                    // (the preview scales the page to fill + drops its chrome when
                    // headless). Default dims -> native px; a custom --screenshot-width
                    // -> that width with an aspect-matched height. The Windows COM path
                    // (directPng != null) renders real pages and is untouched.
                    DocxSinglePageViewport(wordHandler, effectiveFilter, gridCols, ref screenshotWidth, ref screenshotHeight);
                }
            }

            // A renderer that paints its own pixels supplies them directly, sitting
            // between the native backend (which wins when it ran) and the HTML→headless
            // fallback. An explicit --render html opts out. The default install has no
            // PNG-capable renderer, so this is a no-op there.
            if (directPng == null && renderMode != "html")
            {
                var pngFmt = handler switch
                {
                    OfficeCli.Handlers.PowerPointHandler => "pptx",
                    OfficeCli.Handlers.ExcelHandler => "xlsx",
                    OfficeCli.Handlers.WordHandler => "docx",
                    _ => null
                };
                if (pngFmt != null)
                    directPng = RenderPngBytesViaRegistry(handler, pngFmt,
                        new OfficeCli.Core.Rendering.RenderOptions
                        {
                            Output = OfficeCli.Core.Rendering.RenderOutputKind.Png,
                            PageFilter = pageFilter,
                            RasterWidthPx = screenshotWidth,
                            RasterHeightPx = screenshotHeight,
                        });
            }

            if (html == null && directPng == null)
            {
                throw new OfficeCli.Core.CliException("Screenshot mode is only supported for .pptx, .xlsx, and .docx files.")
                {
                    Code = "unsupported_type",
                    Suggestion = "Use a .pptx, .xlsx, or .docx file.",
                    ValidValues = ["text", "annotated", "outline", "stats", "issues", "html", "svg", "screenshot"]
                };
            }

            var pngPath = outArg ?? Path.Combine(Path.GetTempPath(), $"officecli_screenshot_{Path.GetFileNameWithoutExtension(Path.GetFileName(filePath))}_{DateTime.Now:HHmmss}_{Guid.NewGuid():N}.png");
            if (directPng != null)
            {
                File.WriteAllBytes(pngPath, directPng);
            }
            else
            {
                // SECURITY: random token in temp filename — same rationale as the html/--browser path.
                var tmpHtml = Path.Combine(Path.GetTempPath(), $"officecli_preview_{Path.GetFileNameWithoutExtension(Path.GetFileName(filePath))}_{DateTime.Now:HHmmss}_{Guid.NewGuid():N}.html");
                File.WriteAllText(tmpHtml, html!);
                OfficeCli.Core.HtmlScreenshot.Result r;
                try
                {
                    r = clipArg != null
                        ? OfficeCli.Core.HtmlScreenshot.CaptureClipped(tmpHtml, pngPath,
                            OfficeCli.Core.HtmlScreenshot.ResolveClipDataPaths(clipArg))
                        : OfficeCli.Core.HtmlScreenshot.Capture(tmpHtml, pngPath, screenshotWidth, screenshotHeight);
                }
                finally { try { File.Delete(tmpHtml); } catch { /* ignore */ } }
                if (!r.Ok && r.Error == "clip_target_not_found")
                {
                    throw new OfficeCli.Core.CliException(
                        $"--range target '{clipArg}' matched no rendered element.")
                    {
                        Code = "range_target_not_found",
                        Suggestion = "Use a data-path the HTML preview emits (xlsx: 'Sheet1!A1:C3' within the sheet's used area; pptx: '/slide[N]/shape[K]'; docx: '/body/table[N]' or '/body/p[N]'). `query` lists element paths.",
                    };
                }
                if (!r.Ok)
                {
                    throw new OfficeCli.Core.CliException(
                        "No headless browser available. Install Chrome/Edge/Chromium or Firefox, or `pip install playwright && playwright install chromium`."
                        + (r.Error != null ? $" Last error: {r.Error}" : ""))
                    { Code = "no_screenshot_backend" };
                }
            }
            EmitPathOrText(Path.GetFullPath(pngPath), json);
            // Informational slide count on stderr in text mode only: under
            // --json the resident would turn the line into a warning the direct
            // path does not have.
            if (!json && handler is OfficeCli.Handlers.PowerPointHandler pptCount)
                Console.Error.WriteLine($"[pages] total={pptCount.GetSlideCount()}");
            if (browser)
            {
                try
                {
                    var psi = new System.Diagnostics.ProcessStartInfo(pngPath) { UseShellExecute = true };
                    System.Diagnostics.Process.Start(psi);
                }
                catch { /* silently ignore if image viewer can't be opened */ }
            }
            return 0;
        }

        if (mode.ToLowerInvariant() is "svg" or "g")
        {
            if (handler is OfficeCli.Handlers.PowerPointHandler pptSvgHandler)
            {
                // CONSISTENCY(view-page): SVG mode honors --page like html mode; --page wins over --start
                int slideNum = 1;
                if (!string.IsNullOrEmpty(pageFilter))
                {
                    // Same spec grammar and codes as html/screenshot; svg
                    // renders exactly one slide, so a range is refused rather
                    // than silently reduced to its first slide.
                    var (svgStart, svgEnd) = ParsePptHtmlPage(pageFilter, null, null, pptSvgHandler);
                    if (svgStart != svgEnd)
                        throw new OfficeCli.Core.CliException($"Invalid --page value '{pageFilter}': svg renders a single slide.")
                        { Code = "invalid_value", Suggestion = "Pass one slide number, or use screenshot/html for a range." };
                    slideNum = svgStart ?? 1;
                }
                else if (start.HasValue && start.Value > 0)
                {
                    slideNum = start.Value;
                }
                var svg = RenderViaRegistry(handler, "pptx",
                    new OfficeCli.Core.Rendering.RenderOptions
                    { Output = OfficeCli.Core.Rendering.RenderOutputKind.Svg, StartPage = slideNum })!;

                if (browser)
                {
                    string outPath;
                    if (svg.Contains("data-formula"))
                    {
                        // Wrap SVG in HTML shell for KaTeX formula rendering
                        // GUID keeps the path unpredictable so a local attacker
                        // can't pre-plant a symlink at it and have WriteAllText
                        // clobber a victim file (CWE-59) — matches the sibling
                        // preview/screenshot temp writers in this file.
                        outPath = Path.Combine(Path.GetTempPath(), $"officecli_slide{slideNum}_{Path.GetFileNameWithoutExtension(Path.GetFileName(filePath))}_{DateTime.Now:HHmmss}_{Guid.NewGuid():N}.html");
                        // CONSISTENCY(katex-mirror): mirror-first with CDN fallback — see Core/KatexAssets.
                        var html = $"<!DOCTYPE html><html><head><meta charset='UTF-8'><link rel='stylesheet' href='{OfficeCli.Core.KatexAssets.CssUrl}' onerror=\"{OfficeCli.Core.KatexAssets.CssOnErrorJs}\"><script defer src='{OfficeCli.Core.KatexAssets.JsUrl}' onerror=\"{OfficeCli.Core.KatexAssets.JsOnErrorJs("")}\"></script><style>body{{margin:0;display:flex;justify-content:center;background:#f0f0f0}}</style></head><body>{svg}<script>window.addEventListener('load',function(){{document.querySelectorAll('[data-formula]').forEach(function(el){{try{{katex.render(el.getAttribute('data-formula'),el,{{throwOnError:false,displayMode:true}})}}catch(e){{}}}})}})</script></body></html>";
                        File.WriteAllText(outPath, html);
                    }
                    else
                    {
                        outPath = Path.Combine(Path.GetTempPath(), $"officecli_slide{slideNum}_{Path.GetFileNameWithoutExtension(Path.GetFileName(filePath))}_{DateTime.Now:HHmmss}_{Guid.NewGuid():N}.svg");
                        File.WriteAllText(outPath, svg);
                    }
                    EmitPathOrText(outPath, json);
                    try
                    {
                        var psi = new System.Diagnostics.ProcessStartInfo(outPath) { UseShellExecute = true };
                        System.Diagnostics.Process.Start(psi);
                    }
                    catch { /* silently ignore if browser can't be opened */ }
                }
                else
                {
                    if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeText(svg));
                    else Console.WriteLine(svg.TrimEnd('\r', '\n'));
                }
            }
            else if (handler is OfficeCli.Core.Plugins.FormatHandlerProxy svgProxy)
            {
                int? svgPage = null;
                if (!string.IsNullOrEmpty(pageFilter)
                    && int.TryParse(pageFilter.Split(',')[0].Split('-')[0].Trim(), out var sp))
                    svgPage = sp;
                var svg = svgProxy.ViewAsSvg(svgPage);
                if (svg is null)
                    throw new OfficeCli.Core.CliException(
                        $"SVG preview is not supported by the format-handler plugin for {Path.GetExtension(filePath)}.")
                    { Code = "unsupported_type" };
                if (browser)
                {
                    var outPath = Path.Combine(Path.GetTempPath(),
                        $"officecli_preview_{Path.GetFileNameWithoutExtension(Path.GetFileName(filePath))}_{DateTime.Now:HHmmss}_{Guid.NewGuid():N}.svg");
                    File.WriteAllText(outPath, svg);
                    EmitPathOrText(outPath, json);
                    try
                    {
                        var psi = new System.Diagnostics.ProcessStartInfo(outPath) { UseShellExecute = true };
                        System.Diagnostics.Process.Start(psi);
                    }
                    catch { /* silently ignore if viewer can't be opened */ }
                }
                else
                {
                    if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeText(svg));
                    else Console.WriteLine(svg.TrimEnd('\r', '\n'));
                }
            }
            else
            {
                throw new OfficeCli.Core.CliException("SVG preview is only supported for .pptx files.")
                {
                    Code = "unsupported_type",
                    Suggestion = "Use a .pptx file, or use mode 'text' or 'annotated' for other formats.",
                    ValidValues = ["text", "annotated", "outline", "stats", "issues", "html", "svg", "screenshot"]
                };
            }
            return 0;
        }

        int? withPagesValue = null;
        if (withPages && (mode.ToLowerInvariant() is "stats" or "s") && handler is OfficeCli.Handlers.WordHandler wordHandlerForCount)
        {
            if (OperatingSystem.IsWindows())
            {
                nativeLock?.Release();
                try { withPagesValue = OfficeCli.Core.WordPdfBackend.GetPageCount(filePath); } catch { withPagesValue = null; }
                finally { if (nativeLock != null) { handler = nativeLock.Reopen(); wordHandlerForCount = (OfficeCli.Handlers.WordHandler)handler; } }
            }
            if (withPagesValue == null)
            {
                var tmpHtml = Path.Combine(Path.GetTempPath(), $"officecli_pc_{Path.GetFileNameWithoutExtension(Path.GetFileName(filePath))}_{Guid.NewGuid():N}.html");
                try
                {
                    File.WriteAllText(tmpHtml, RenderViaRegistry(wordHandlerForCount, "docx", new OfficeCli.Core.Rendering.RenderOptions())!);
                    withPagesValue = OfficeCli.Core.HtmlScreenshot.GetPageCountFromDom(tmpHtml);
                }
                finally { try { File.Delete(tmpHtml); } catch { } }
            }
            if (withPagesValue == null)
                throw new OfficeCli.Core.CliException("--page-count: failed to get page count (Word backend and HTML fallback both unavailable).")
                { Code = "page_count_unavailable" };
        }

        if (json)
        {
            // Structured JSON output — no Content string wrapping
            var modeKey = mode.ToLowerInvariant();
            if (modeKey is "stats" or "s")
            {
                var statsJson = handler.ViewAsStatsJson();
                if (withPagesValue.HasValue) statsJson["pages"] = withPagesValue.Value;
                Console.WriteLine(OutputFormatter.WrapEnvelope(statsJson.ToJsonString(OutputFormatter.PublicJsonOptions)));
            }
            else if (modeKey is "outline" or "o")
                Console.WriteLine(OutputFormatter.WrapEnvelope(handler.ViewAsOutlineJson().ToJsonString(OutputFormatter.PublicJsonOptions)));
            else if (modeKey is "text" or "t")
                Console.WriteLine(OutputFormatter.WrapEnvelope(handler.ViewAsTextJson(start, end, maxLines, cols, clipArg).ToJsonString(OutputFormatter.PublicJsonOptions)));
            else if (modeKey is "annotated" or "a")
                Console.WriteLine(OutputFormatter.WrapEnvelope(
                    OutputFormatter.FormatView(mode, handler.ViewAsAnnotated(start, end, maxLines, cols), OutputFormat.Json)));
            else if (modeKey is "issues" or "i")
                Console.WriteLine(OutputFormatter.WrapEnvelope(
                    OutputFormatter.FormatIssues(handler.ViewAsIssues(issueType, limit), OutputFormat.Json)));
            else if (modeKey is "forms" or "f")
            {
                if (handler is OfficeCli.Handlers.WordHandler wordFormsHandler)
                    Console.WriteLine(OutputFormatter.WrapEnvelope(wordFormsHandler.ViewAsFormsJson().ToJsonString(OutputFormatter.PublicJsonOptions)));
                else if (handler is OfficeCli.Core.Plugins.FormatHandlerProxy formsProxy)
                {
                    var formsJson = formsProxy.ViewAsFormsJson();
                    if (formsJson is null)
                        throw new OfficeCli.Core.CliException($"Forms view is not supported by the format-handler plugin for {Path.GetExtension(filePath)}.")
                        { Code = "unsupported_type" };
                    Console.WriteLine(OutputFormatter.WrapEnvelope(formsJson.ToJsonString(OutputFormatter.PublicJsonOptions)));
                }
                else
                    throw new OfficeCli.Core.CliException("Forms view is only supported for .docx files.")
                    {
                        Code = "unsupported_type",
                        ValidValues = ["text", "annotated", "outline", "stats", "issues", "html", "svg", "screenshot", "pdf", "forms"]
                    };
            }
            else
                throw new OfficeCli.Core.CliException($"Unknown mode: {mode}. Available: text, annotated, outline, stats, issues, html, svg, screenshot, pdf, forms")
                {
                    Code = "invalid_value",
                    ValidValues = ["text", "annotated", "outline", "stats", "issues", "html", "svg", "screenshot", "pdf", "forms"]
                };
        }
        else
        {
            var output = mode.ToLowerInvariant() switch
            {
                "text" or "t" => handler.ViewAsText(start, end, maxLines, cols, clipArg),
                "annotated" or "a" => handler.ViewAsAnnotated(start, end, maxLines, cols),
                "outline" or "o" => handler.ViewAsOutline(),
                "stats" or "s" => withPagesValue.HasValue
                    ? $"Pages: {withPagesValue}\n" + handler.ViewAsStats()
                    : handler.ViewAsStats(),
                "issues" or "i" => OutputFormatter.FormatIssues(handler.ViewAsIssues(issueType, limit), OutputFormat.Text),
                "forms" or "f" => handler switch
                {
                    OfficeCli.Handlers.WordHandler wfh => wfh.ViewAsForms(),
                    OfficeCli.Core.Plugins.FormatHandlerProxy fp
                        => fp.ViewAsFormsJson()?.ToJsonString(OutputFormatter.PublicJsonOptions)
                           ?? throw new OfficeCli.Core.CliException($"Forms view is not supported by the format-handler plugin for {Path.GetExtension(filePath)}.")
                               { Code = "unsupported_type" },
                    _ => throw new OfficeCli.Core.CliException("Forms view is only supported for .docx files.")
                    {
                        Code = "unsupported_type",
                        ValidValues = ["text", "annotated", "outline", "stats", "issues", "html", "svg", "screenshot", "pdf", "forms"]
                    }
                },
                _ => throw new OfficeCli.Core.CliException($"Unknown mode: {mode}. Available: text, annotated, outline, stats, issues, html, svg, screenshot, pdf, forms")
                {
                    Code = "invalid_value",
                    ValidValues = ["text", "annotated", "outline", "stats", "issues", "html", "svg", "screenshot", "pdf", "forms"]
                }
            };
            // An empty view (e.g. --start past the end) prints nothing rather
            // than a lone blank line, on both paths.
            if (output.Length > 0) Console.WriteLine(output);
        }
        return 0;
    }
}
