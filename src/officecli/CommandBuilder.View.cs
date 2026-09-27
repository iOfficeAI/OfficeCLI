// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using OfficeCli.Core;
using OfficeCli.Handlers;

namespace OfficeCli;

static partial class CommandBuilder
{
    private static Command BuildViewCommand(Option<bool> jsonOption)
    {
        var viewFileArg = new Argument<FileInfo>("file") { Description = "Office document path (.docx, .xlsx, .pptx)" };
        var viewModeArg = new Argument<string>("mode") { Description = "View mode: text, annotated, outline, stats, issues, html, svg, screenshot, pdf, forms. text mode (xlsx): each cell is rendered as <A1>=<value>, tab-separated; empty cells are omitted, so a sparse row lists only its populated cells (e.g. 'B2=120000\\tD2=Beijing')." };
        var startLineOpt = new Option<int?>("--start") { Description = "Start line/paragraph number" };
        var endLineOpt = new Option<int?>("--end") { Description = "End line/paragraph number" };
        var maxLinesOpt = new Option<int?>("--max-lines") { Description = "Maximum number of lines/rows/slides to output (truncates with total count)" };
        var issueTypeOpt = new Option<string?>("--type") { Description = IssueSubtypes.TypeHelpDescription() };
        var limitOpt = new Option<int?>("--limit") { Description = "Limit number of results" };

        var colsOpt = new Option<string?>("--cols") { Description = "Column filter, comma-separated (Excel only, e.g. A,B,C)" };
        var pageOpt = new Option<string?>("--page") { Description = "Page filter (e.g. 1, 2-5, 1,3,5). html mode: default=all. screenshot mode: default=1 (use --page 1-N to capture more, or --grid N for a whole-doc thumbnail contact sheet)." };
        var browserOpt = new Option<bool>("--browser") { Description = "Open output in browser (html / svg modes)" };
        var outOpt = new Option<string?>("--out", "-o") { Description = "Output file path (html, screenshot, pdf modes; defaults to stdout for html, a temp file for screenshot)" };
        var clipOpt = new Option<string?>("--range") { Description = "Restrict output to a region. Screenshot mode: an xlsx cell range ('Sheet1!A1:C3' or '/Sheet1/A1:C3') or any element data-path ('/slide[1]/shape[@id=N]', '/body/table[1]'); the PNG is cropped to the target's bounding box. Text mode (xlsx only): a cell range or single cell — emits just those rows/cells, saving context on large sheets. Not the character-offset `range=` prop of `set` (that one formats a text span like 3:7)." };
        var screenshotWidthOpt = new Option<int>("--screenshot-width") { Description = "Screenshot viewport width (default 1600)", DefaultValueFactory = _ => 1600 };
        var screenshotHeightOpt = new Option<int>("--screenshot-height") { Description = "Screenshot viewport height (default 1200)", DefaultValueFactory = _ => 1200 };
        var gridOpt = new Option<string?>("--grid")
        {
            Description = "Tile pages/slides into a thumbnail contact sheet (screenshot mode, pptx + docx). Bare --grid (or --grid auto) picks a column count that keeps the sheet roughly square; pass a number (e.g. --grid 3) to force columns. Omit = off.",
            Arity = ArgumentArity.ZeroOrOne, // allow bare --grid (no value) → auto
        };
        var renderOpt = new Option<string>("--render") { Description = "Screenshot rendering path (docx/pptx): auto (default; native on Windows w/ Word/PowerPoint, html elsewhere), native (force OS-native, error if unavailable), html", DefaultValueFactory = _ => "auto" };
        var withPagesOpt = new Option<bool>("--page-count") { Description = "stats mode (docx only): also report total page count via Word repagination (Win + Word required; slow on long docs)" };

        var viewCommand = new Command("view", "View document in different modes");
        viewCommand.Add(viewFileArg);
        viewCommand.Add(viewModeArg);
        viewCommand.Add(startLineOpt);
        viewCommand.Add(endLineOpt);
        viewCommand.Add(maxLinesOpt);
        viewCommand.Add(issueTypeOpt);
        viewCommand.Add(limitOpt);
        viewCommand.Add(colsOpt);
        viewCommand.Add(pageOpt);
        viewCommand.Add(browserOpt);
        viewCommand.Add(outOpt);
        viewCommand.Add(clipOpt);
        viewCommand.Add(screenshotWidthOpt);
        viewCommand.Add(screenshotHeightOpt);
        viewCommand.Add(gridOpt);
        viewCommand.Add(renderOpt);
        viewCommand.Add(withPagesOpt);
        viewCommand.Add(jsonOption);

        viewCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(viewFileArg)!;
            var mode = result.GetValue(viewModeArg)!;
            var start = result.GetValue(startLineOpt);
            var end = result.GetValue(endLineOpt);
            var maxLines = result.GetValue(maxLinesOpt);
            var issueType = IssueSubtypes.Validate(result.GetValue(issueTypeOpt));
            var limit = result.GetValue(limitOpt);
            var colsStr = result.GetValue(colsOpt);
            var pageFilter = result.GetValue(pageOpt);
            var browser = result.GetValue(browserOpt);
            var outArg = result.GetValue(outOpt);
            var clipArg = result.GetValue(clipOpt);
            var screenshotWidth = result.GetValue(screenshotWidthOpt);
            var screenshotHeight = result.GetValue(screenshotHeightOpt);
            // --grid has three states: absent → off (0), present with no value
            // (bare --grid) → auto (-1), present with a value → parse it.
            var gridResult = result.GetResult(gridOpt);
            var gridCols = gridResult is null ? 0
                : gridResult.Tokens.Count == 0 ? -1
                : ParseGridSpec(gridResult.Tokens[0].Value);
            var renderMode = (result.GetValue(renderOpt) ?? "auto").ToLowerInvariant();
            if (renderMode is not ("auto" or "native" or "html"))
                throw new OfficeCli.Core.CliException($"Invalid --render value: {renderMode}. Valid: auto, native, html") { Code = "invalid_render", ValidValues = ["auto", "native", "html"] };
            var withPages = result.GetValue(withPagesOpt);

            // pdf mode runs entirely through an exporter plugin (no handler
            // open, no resident hop — the plugin gets a snapshot of the
            // source and writes the PDF). Handled before TryResident
            // because exporter invocation needs the file lock released, and
            // ExporterInvoker closes the resident itself when present.
            if (mode.ToLowerInvariant() is "pdf")
            {
                var pdfPath = outArg ?? Path.ChangeExtension(file.FullName, "pdf");
                var exp = OfficeCli.Core.Plugins.ExporterInvoker.Run(file.FullName, ".pdf", pdfPath);
                if (json)
                {
                    Console.WriteLine(OutputFormatter.WrapEnvelopeText(exp.OutputPath));
                }
                else
                {
                    Console.WriteLine(Path.GetFullPath(exp.OutputPath));
                    if (exp.ResidentClosed)
                        Console.Error.WriteLine($"[note] resident closed to release lock; reopen with `officecli open` if needed");
                }
                if (browser)
                {
                    try
                    {
                        var psi = new System.Diagnostics.ProcessStartInfo(exp.OutputPath) { UseShellExecute = true };
                        System.Diagnostics.Process.Start(psi);
                    }
                    catch { /* silently ignore if no default PDF viewer */ }
                }
                return 0;
            }

            // Try resident first
            if (TryResident(file.FullName, req =>
            {
                req.Command = "view";
                req.Json = json;
                req.Args["mode"] = mode;
                if (start.HasValue) req.Args["start"] = start.Value.ToString();
                if (end.HasValue) req.Args["end"] = end.Value.ToString();
                if (maxLines.HasValue) req.Args["max-lines"] = maxLines.Value.ToString();
                if (issueType != null) req.Args["type"] = issueType;
                if (limit.HasValue) req.Args["limit"] = limit.Value.ToString();
                if (colsStr != null) req.Args["cols"] = colsStr;
                if (pageFilter != null) req.Args["page"] = pageFilter;
                if (browser) req.Args["browser"] = "true";
                if (outArg != null) req.Args["out"] = outArg;
                if (clipArg != null) req.Args["range"] = clipArg;
                req.Args["screenshot-width"] = screenshotWidth.ToString();
                req.Args["screenshot-height"] = screenshotHeight.ToString();
                if (gridCols != 0) req.Args["grid"] = gridCols.ToString(); // -1 = auto
                if (renderMode != "auto") req.Args["render"] = renderMode;
                if (withPages) req.Args["page-count"] = "true";
            }, json) is {} rc) return rc;

            var format = json ? OutputFormat.Json : OutputFormat.Text;
            var cols = colsStr != null ? new HashSet<string>(colsStr.Split(',').Select(c => c.Trim().ToUpperInvariant())) : null;

            using var handler = DocumentHandlerFactory.Open(file.FullName);

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
                    html = RenderViaRegistry(handler, "docx",
                        new OfficeCli.Core.Rendering.RenderOptions { PageFilter = pageFilter });
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
                        var htmlPath = outArg ?? Path.Combine(Path.GetTempPath(), $"officecli_preview_{Path.GetFileNameWithoutExtension(file.Name)}_{DateTime.Now:HHmmss}_{Guid.NewGuid():N}.html");
                        File.WriteAllText(htmlPath, html);
                        Console.WriteLine(Path.GetFullPath(htmlPath));
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
                        // Default: output HTML to stdout
                        Console.Write(html);
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
                    Exception? nativeFailure = null;
                    bool nativeAttempted = renderMode != "html" && OperatingSystem.IsWindows();
                    if (nativeAttempted)
                    {
                        try
                        {
                            if (gridColsResolved > 0)
                            {
                                const int gap = 12, pad = 12;
                                int cellW = Math.Max(1, (int)Math.Round((screenshotWidth - 2 * pad - (gridColsResolved - 1) * gap) / (double)gridColsResolved));
                                int cellH = Math.Max(1, (int)Math.Round(cellW * (double)nativeH / nativeW));
                                directPng = OfficeCli.Core.PowerPointPngBackend.RenderGrid(file.FullName, pStart ?? 1, pEnd ?? pptHandler.GetSlideCount(), cellW, cellH, gridColsResolved, gap, pad);
                            }
                            else
                            {
                                directPng = OfficeCli.Core.PowerPointPngBackend.Render(file.FullName, pStart ?? 1, pEnd ?? pStart ?? 1, exportW, exportH);
                            }
                        }
                        catch (Exception e) { nativeFailure = e; directPng = null; }
                    }
                    if (renderMode == "native" && directPng == null)
                        throw OfficeCli.Core.NativeRenderDiagnostics.Create(
                            "Microsoft PowerPoint", nativeAttempted, nativeFailure);

                    if (directPng == null)
                    {
                        html = RenderViaRegistry(pptHandler, "pptx", new OfficeCli.Core.Rendering.RenderOptions
                        { StartPage = pStart, EndPage = pEnd, GridColumns = gridColsResolved, ViewportPx = screenshotWidth })!;

                        // The generic 4:3 viewport (1600×1200) letterboxes a single slide with
                        // canvas padding. When capturing one slide (not a multi-slide range or
                        // grid), size the viewport to the slide so the PNG is the slide,
                        // padding-free (ViewAsHtml scales the slide to fill + zeroes the headless
                        // page padding). Default dims -> the slide's 96-DPI native pixels;
                        // a custom --screenshot-width -> that width with an aspect-matched height.
                        // Multi-slide ranges stack vertically and keep the tall viewport.
                        if (pStart == pEnd && gridCols == 0)
                        {
                            if (screenshotWidth == 1600 && screenshotHeight == 1200)
                                (screenshotWidth, screenshotHeight) = (nativeW, nativeH);
                            else if (screenshotHeight == 1200)
                                screenshotHeight = Math.Max(1, (int)Math.Round(screenshotWidth * (double)nativeH / nativeW));
                        }
                    }
                }
                else if (handler is OfficeCli.Handlers.ExcelHandler excelHandler)
                    html = RenderViaRegistry(excelHandler, "xlsx", new OfficeCli.Core.Rendering.RenderOptions())!;
                else if (handler is OfficeCli.Handlers.WordHandler wordHandlerGrid && gridCols != 0)
                {
                    // Contact-sheet grid: tile every page into an N-column (or auto)
                    // thumbnail grid for a one-shot whole-document overview.
                    // Native-first, mirroring pptx and single-page docx: on Windows
                    // with Word, RenderGrid rasterizes each real-Word page and tiles
                    // it; elsewhere (or on failure) we fall back to the HTML preview
                    // grid. Column count + cell size are computed once below and used
                    // by both paths.
                    const int gap = 12, pad = 12;
                    const int maxDim = 1920; // mirror HtmlScreenshot.CapDim's LLM-image ceiling
                    const int scrollbar = 17;
                    var (npW, npH) = wordHandlerGrid.GetPageNativePixels();

                    // Page count needs a real layout pass; the preview's paginator
                    // publishes it via <title>PAGES:N> on dump-dom (independent of the
                    // grid, which only reflows AFTER pagination). Count first so
                    // `--grid auto` can size the column count to the document.
                    int pageCount = 1;
                    var tmpForCount = Path.Combine(Path.GetTempPath(), $"officecli_gridcount_{Path.GetFileNameWithoutExtension(file.Name)}_{Guid.NewGuid():N}.html");
                    try
                    {
                        File.WriteAllText(tmpForCount, RenderViaRegistry(wordHandlerGrid, "docx", new OfficeCli.Core.Rendering.RenderOptions())!);
                        pageCount = OfficeCli.Core.HtmlScreenshot.GetPageCountFromDom(tmpForCount) ?? 1;
                    }
                    catch { /* fall back to 1 row */ }
                    finally { try { File.Delete(tmpForCount); } catch { /* ignore */ } }

                    // -1 = auto: pick columns that keep the composed sheet ≈ square.
                    int docGridCols = gridCols < 0 ? OfficeCli.Core.HtmlScreenshot.AutoGridColumns(pageCount, npW, npH) : gridCols;
                    int rows = Math.Max(1, (pageCount + docGridCols - 1) / docGridCols);

                    // Pre-cap to the 1920 ceiling OURSELVES and recompute cellW from the
                    // final width, so cellW and the capture viewport stay consistent
                    // (CapDim shrinking the viewport while cellW stayed fixed is what
                    // collapsed a 3-col request to 2 cols). After this, CapDim is a no-op.
                    // Subtract a scrollbar allowance so this matches layoutGrid's
                    // clientWidth-derived cell size → the row-height estimate (and thus
                    // the captured viewport) tracks the real grid with little slack.
                    double vpW = screenshotWidth;
                    double cellW = Math.Max(1.0, (vpW - scrollbar - 2.0 * pad - (docGridCols - 1) * gap) / docGridCols);
                    double cellH = cellW * npH / npW;
                    double vpH = pad * 2 + rows * cellH + (rows - 1) * gap;
                    double over = Math.Max(vpW, vpH) / maxDim;
                    if (over > 1.0) { vpW /= over; cellW /= over; cellH /= over; vpH /= over; }

                    // Native-first: render each real-Word page and tile (Windows + Word).
                    Exception? nativeFailure = null;
                    bool nativeAttempted = renderMode != "html" && OperatingSystem.IsWindows();
                    if (nativeAttempted)
                    {
                        try { directPng = OfficeCli.Core.WordPdfBackend.RenderGrid(file.FullName, $"1-{pageCount}", (int)Math.Round(cellW), (int)Math.Round(cellH), docGridCols, gap, pad); }
                        catch (Exception e) { nativeFailure = e; directPng = null; }
                    }
                    if (renderMode == "native" && directPng == null)
                        throw OfficeCli.Core.NativeRenderDiagnostics.Create(
                            "Microsoft Word", nativeAttempted, nativeFailure);
                    if (directPng == null)
                    {
                        // HTML fallback: layoutGrid tiles in-browser; size the viewport
                        // to fit the rows so window-size backends don't crop.
                        html = RenderViaRegistry(wordHandlerGrid, "docx", new OfficeCli.Core.Rendering.RenderOptions
                        { GridColumns = docGridCols, GridCellWidthPx = (int)Math.Round(cellW) })!;
                        screenshotWidth = Math.Max(1, (int)Math.Round(vpW));
                        screenshotHeight = Math.Max(1, (int)Math.Ceiling(vpH));
                    }
                }
                else if (handler is OfficeCli.Handlers.WordHandler wordHandler)
                {
                    // --clip: render every page (filter stays null) unless the
                    // caller narrowed with --page — the target may be anywhere.
                    var effectiveFilter = clipArg != null
                        ? pageFilter
                        : (string.IsNullOrEmpty(pageFilter) ? "1" : pageFilter);
                    Exception? nativeFailure = null;
                    bool nativeAttempted = renderMode != "html" && OperatingSystem.IsWindows();
                    if (nativeAttempted)
                    {
                        // effectiveFilter is only null under --range, which forces
                        // renderMode=html — this native branch is then unreachable.
                        try { directPng = OfficeCli.Core.WordPdfBackend.Render(file.FullName, effectiveFilter!); }
                        catch (Exception e) { nativeFailure = e; directPng = null; }
                    }
                    if (renderMode == "native" && directPng == null)
                        throw OfficeCli.Core.NativeRenderDiagnostics.Create(
                            "Microsoft Word", nativeAttempted, nativeFailure);
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
                        if (int.TryParse(effectiveFilter, out _) && gridCols == 0)
                        {
                            var (nativeW, nativeH) = wordHandler.GetPageNativePixels();
                            if (screenshotWidth == 1600 && screenshotHeight == 1200)
                                (screenshotWidth, screenshotHeight) = (nativeW, nativeH);
                            else if (screenshotHeight == 1200)
                                screenshotHeight = Math.Max(1, (int)Math.Round(screenshotWidth * (double)nativeH / nativeW));
                        }
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

                var pngPath = outArg ?? Path.Combine(Path.GetTempPath(), $"officecli_screenshot_{Path.GetFileNameWithoutExtension(file.Name)}_{DateTime.Now:HHmmss}_{Guid.NewGuid():N}.png");
                if (directPng != null)
                {
                    File.WriteAllBytes(pngPath, directPng);
                }
                else
                {
                    // SECURITY: random token in temp filename — same rationale as the html/--browser path.
                    var tmpHtml = Path.Combine(Path.GetTempPath(), $"officecli_preview_{Path.GetFileNameWithoutExtension(file.Name)}_{DateTime.Now:HHmmss}_{Guid.NewGuid():N}.html");
                    File.WriteAllText(tmpHtml, html!);
                    var r = clipArg != null
                        ? OfficeCli.Core.HtmlScreenshot.CaptureClipped(tmpHtml, pngPath,
                            OfficeCli.Core.HtmlScreenshot.ResolveClipDataPaths(clipArg))
                        : OfficeCli.Core.HtmlScreenshot.Capture(tmpHtml, pngPath, screenshotWidth, screenshotHeight);
                    try { File.Delete(tmpHtml); } catch { /* ignore */ }
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
                Console.WriteLine(Path.GetFullPath(pngPath));
                if (handler is OfficeCli.Handlers.PowerPointHandler pptCount)
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
                        var firstTok = pageFilter.Split(',')[0].Split('-')[0].Trim();
                        // CONSISTENCY(strict-page): reject non-positive --page
                        // values explicitly instead of silently rendering
                        // slide 1, mirroring how 0 / negatives are surfaced
                        // elsewhere in the CLI.
                        if (!int.TryParse(firstTok, out var p))
                            throw new ArgumentException(
                                $"Invalid --page value '{pageFilter}': expected a positive slide number.");
                        if (p <= 0)
                            throw new ArgumentException(
                                $"Invalid --page value '{pageFilter}': slide number must be >= 1.");
                        slideNum = p;
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
                            outPath = Path.Combine(Path.GetTempPath(), $"officecli_slide{slideNum}_{Path.GetFileNameWithoutExtension(file.Name)}_{DateTime.Now:HHmmss}_{Guid.NewGuid():N}.html");
                            // CONSISTENCY(katex-mirror): mirror-first with CDN fallback — see Core/KatexAssets.
                            var html = $"<!DOCTYPE html><html><head><meta charset='UTF-8'><link rel='stylesheet' href='{OfficeCli.Core.KatexAssets.CssUrl}' onerror=\"{OfficeCli.Core.KatexAssets.CssOnErrorJs}\"><script defer src='{OfficeCli.Core.KatexAssets.JsUrl}' onerror=\"{OfficeCli.Core.KatexAssets.JsOnErrorJs("")}\"></script><style>body{{margin:0;display:flex;justify-content:center;background:#f0f0f0}}</style></head><body>{svg}<script>window.addEventListener('load',function(){{document.querySelectorAll('[data-formula]').forEach(function(el){{try{{katex.render(el.getAttribute('data-formula'),el,{{throwOnError:false,displayMode:true}})}}catch(e){{}}}})}})</script></body></html>";
                            File.WriteAllText(outPath, html);
                        }
                        else
                        {
                            outPath = Path.Combine(Path.GetTempPath(), $"officecli_slide{slideNum}_{Path.GetFileNameWithoutExtension(file.Name)}_{DateTime.Now:HHmmss}_{Guid.NewGuid():N}.svg");
                            File.WriteAllText(outPath, svg);
                        }
                        Console.WriteLine(outPath);
                        try
                        {
                            var psi = new System.Diagnostics.ProcessStartInfo(outPath) { UseShellExecute = true };
                            System.Diagnostics.Process.Start(psi);
                        }
                        catch { /* silently ignore if browser can't be opened */ }
                    }
                    else
                    {
                        Console.Write(svg);
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
                            $"SVG preview is not supported by the format-handler plugin for {file.Extension}.")
                        { Code = "unsupported_type" };
                    if (browser)
                    {
                        var outPath = Path.Combine(Path.GetTempPath(),
                            $"officecli_preview_{Path.GetFileNameWithoutExtension(file.Name)}_{DateTime.Now:HHmmss}_{Guid.NewGuid():N}.svg");
                        File.WriteAllText(outPath, svg);
                        Console.WriteLine(outPath);
                        try
                        {
                            var psi = new System.Diagnostics.ProcessStartInfo(outPath) { UseShellExecute = true };
                            System.Diagnostics.Process.Start(psi);
                        }
                        catch { /* silently ignore if viewer can't be opened */ }
                    }
                    else
                    {
                        Console.Write(svg);
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
                    try { withPagesValue = OfficeCli.Core.WordPdfBackend.GetPageCount(file.FullName); } catch { withPagesValue = null; }
                }
                if (withPagesValue == null)
                {
                    var tmpHtml = Path.Combine(Path.GetTempPath(), $"officecli_pc_{Path.GetFileNameWithoutExtension(file.Name)}_{Guid.NewGuid():N}.html");
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
                            throw new OfficeCli.Core.CliException($"Forms view is not supported by the format-handler plugin for {file.Extension}.")
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
                    throw new OfficeCli.Core.CliException($"Unknown mode: {mode}. Available: text, annotated, outline, stats, issues, html, svg, screenshot, forms")
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
                               ?? throw new OfficeCli.Core.CliException($"Forms view is not supported by the format-handler plugin for {file.Extension}.")
                                   { Code = "unsupported_type" },
                        _ => throw new OfficeCli.Core.CliException("Forms view is only supported for .docx files.")
                        {
                            Code = "unsupported_type",
                            ValidValues = ["text", "annotated", "outline", "stats", "issues", "html", "svg", "screenshot", "pdf", "forms"]
                        }
                    },
                    _ => throw new OfficeCli.Core.CliException($"Unknown mode: {mode}. Available: text, annotated, outline, stats, issues, html, svg, screenshot, forms")
                    {
                        Code = "invalid_value",
                        ValidValues = ["text", "annotated", "outline", "stats", "issues", "html", "svg", "screenshot", "pdf", "forms"]
                    }
                };
                Console.WriteLine(output);
            }
            return 0;
        }, json); });

        return viewCommand;
    }

    /// <summary>
    /// BUG-R36-B7 helper. Resolve --page (and fallback --start/--end) into a
    /// validated (startSlide, endSlide) pair for pptx html previews. Rejects
    /// non-positive numbers and indices past the slide count instead of
    /// silently rendering the whole deck.
    /// </summary>
    // Interpret the --grid value: absent/empty → 0 (off), "auto" → -1 (pick a
    // column count that keeps the sheet roughly square), a non-negative integer
    // → that explicit column count. Anything else is a hard error.
    private static int ParseGridSpec(string? spec)
    {
        if (string.IsNullOrWhiteSpace(spec)) return 0;
        spec = spec.Trim();
        if (spec.Equals("auto", StringComparison.OrdinalIgnoreCase)) return -1;
        if (int.TryParse(spec, out var n) && n >= 0) return n;
        throw new OfficeCli.Core.CliException($"Invalid --grid value: {spec}. Use a column count (e.g. 3) or 'auto'.")
        { Code = "invalid_value", ValidValues = ["auto", "1", "2", "3", "4"] };
    }

    // Render through the renderer registry rather than calling the handler directly.
    // The built-in renderers forward to the handler's ViewAs* methods (output is
    // byte-identical), so this is behavior-preserving; it also lets an alternative
    // renderer registered for the format be selected instead. Returns null when no
    // renderer covers the request, preserving the unsupported-type path.
    internal static string? RenderViaRegistry(
        OfficeCli.Core.IDocumentHandler handler, string formatId,
        OfficeCli.Core.Rendering.RenderOptions options)
    {
        OfficeCli.Handlers.Rendering.RenderingBootstrap.EnsureRegistered();
        var renderer = OfficeCli.Core.Rendering.RendererRegistry.Default
            .Resolve(formatId, options.Output, options.Mode);
        return renderer?.Render(
            new OfficeCli.Handlers.Rendering.HandlerRenderInput(handler, formatId), options).Text;
    }

    // PNG bytes from a renderer that paints its own pixels, or null when none is
    // registered for this format (the built-in renderers emit HTML only, so the
    // default install returns null here and the HTML→headless path is used). The
    // native (real Office) backend, when it ran, has already produced the pixels and
    // this is not reached.
    internal static byte[]? RenderPngBytesViaRegistry(
        OfficeCli.Core.IDocumentHandler handler, string formatId,
        OfficeCli.Core.Rendering.RenderOptions options)
    {
        OfficeCli.Handlers.Rendering.RenderingBootstrap.EnsureRegistered();
        var renderer = OfficeCli.Core.Rendering.RendererRegistry.Default
            .Resolve(formatId, OfficeCli.Core.Rendering.RenderOutputKind.Png, options.Mode);
        if (renderer == null) return null;
        return renderer.Render(
            new OfficeCli.Handlers.Rendering.HandlerRenderInput(handler, formatId), options).Bytes;
    }

    private static (int? start, int? end) ParsePptHtmlPage(
        string? pageFilter, int? start, int? end,
        OfficeCli.Handlers.PowerPointHandler pptHandler)
    {
        if (string.IsNullOrEmpty(pageFilter)) return (start, end);
        var slideCount = pptHandler.Query("slide").Count;
        var firstTok = pageFilter.Split(',')[0].Trim();
        // Range form "M-N"
        if (firstTok.Contains('-'))
        {
            var parts = firstTok.Split('-', 2);
            if (!int.TryParse(parts[0], out var ps) || !int.TryParse(parts[1], out var pe))
                throw new ArgumentException($"Invalid --page value '{pageFilter}': expected N or M-N or comma list.");
            if (ps <= 0 || pe <= 0)
                throw new ArgumentException($"Invalid --page value '{pageFilter}': slide number must be >= 1.");
            if (ps > slideCount)
                throw new ArgumentException($"--page {ps} out of range (total slides: {slideCount}).");
            return (ps, Math.Min(pe, slideCount));
        }
        if (!int.TryParse(firstTok, out var p))
            throw new ArgumentException($"Invalid --page value '{pageFilter}': expected a positive slide number.");
        if (p <= 0)
            throw new ArgumentException($"Invalid --page value '{pageFilter}': slide number must be >= 1.");
        if (p > slideCount)
            throw new ArgumentException($"--page {p} out of range (total slides: {slideCount}).");
        return (p, p);
    }
}
