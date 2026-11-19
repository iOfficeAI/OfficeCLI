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
                // Absolute: the resident resolves relative paths against ITS
                // working directory, not the caller's.
                if (outArg != null) req.Args["out"] = Path.GetFullPath(outArg);
                if (clipArg != null) req.Args["range"] = clipArg;
                req.Args["screenshot-width"] = screenshotWidth.ToString();
                req.Args["screenshot-height"] = screenshotHeight.ToString();
                if (gridCols != 0) req.Args["grid"] = gridCols.ToString(); // -1 = auto
                if (renderMode != "auto") req.Args["render"] = renderMode;
                if (withPages) req.Args["page-count"] = "true";
            }, json) is {} rc) return rc;

            using var handler = DocumentHandlerFactory.Open(file.FullName);
            return RunView(handler, file.FullName, new ViewArgs(mode, start, end, maxLines, issueType, limit, colsStr,
                pageFilter, browser, outArg, clipArg, screenshotWidth, screenshotHeight, gridCols, renderMode, withPages, json), null);
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
        // 0 is not a column count (it used to be accepted and silently turn the grid off).
        if (int.TryParse(spec, out var n) && n >= 1) return n;
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

    // ---- Screenshot viewport sizing shared by the standalone and resident paths ----
    // ResidentServer.cs mirrors the screenshot branches of this file; the sizing
    // maths live here once so the two paths cannot drift apart again (issue #485).

    /// <summary>
    /// Size the HTML-screenshot viewport for a pptx capture and return the grid
    /// column count the preview should tile with (0 = plain stack).
    /// A multi-slide range without --grid is a one-column contact sheet; a grid
    /// mirrors the preview CSS (24px total padding, 12px gap) and the 1920px
    /// capture ceiling so the PNG is exactly the sheet. A single slide at the
    /// default dims gets the slide's 96-DPI native pixels, padding-free.
    /// </summary>
    internal static int PptHtmlScreenshotViewport(int gridCols, int gridColsResolved, int? pStart, int? pEnd,
        int slideCount, int nativeW, int nativeH, ref int screenshotWidth, ref int screenshotHeight)
    {
        int htmlGridCols = gridColsResolved > 0 ? gridColsResolved
            : (pStart.HasValue && pEnd.HasValue && pStart.Value != pEnd.Value ? 1 : 0);
        if (htmlGridCols > 0)
        {
            const int pad = 24, gap = 12, maxDim = 1920;
            int count = (pEnd ?? slideCount) - (pStart ?? 1) + 1;
            int rows = Math.Max(1, (count + htmlGridCols - 1) / htmlGridCols);
            double vpW = screenshotWidth == 1600 && screenshotHeight == 1200 && htmlGridCols == 1
                ? nativeW + pad : screenshotWidth;
            double cellW = Math.Max(1.0, (vpW - pad - (htmlGridCols - 1) * gap) / htmlGridCols);
            double cellH = cellW * nativeH / nativeW;
            double vpH = pad + rows * cellH + (rows - 1) * gap;
            double over = Math.Max(vpW, vpH) / maxDim;
            if (over > 1.0) { vpW /= over; vpH /= over; }
            screenshotWidth = Math.Max(1, (int)Math.Round(vpW));
            screenshotHeight = Math.Max(1, (int)Math.Ceiling(vpH));
        }
        else if (pStart == pEnd && gridCols == 0)
        {
            if (screenshotWidth == 1600 && screenshotHeight == 1200)
                (screenshotWidth, screenshotHeight) = (nativeW, nativeH);
            else if (screenshotHeight == 1200)
                screenshotHeight = Math.Max(1, (int)Math.Round(screenshotWidth * (double)nativeH / nativeW));
        }
        return htmlGridCols;
    }

    internal const int DocxSheetGap = 12, DocxSheetPad = 12;

    /// <summary>Does a docx screenshot request take the contact-sheet path?</summary>
    internal static bool DocxWantsSheet(int gridCols, string? pageFilter)
        => gridCols != 0 || IsMultiPageFilter(pageFilter);

    internal sealed record DocxSheet(int Cols, int CellWpx, int CellHpx, int ViewportW, int ViewportH,
        string NativePageSpec, OfficeCli.Core.Rendering.RenderOptions HtmlOptions);

    /// <summary>
    /// Column count, cell size, viewport and page selection for a docx contact
    /// sheet. The page count needs a real layout pass (the preview publishes it
    /// via &lt;title&gt;PAGES:N&gt;). The sheet is pre-capped to the 1920px
    /// ceiling so cell size and viewport stay consistent; a one-column sheet at
    /// the default dims keeps the page's native width before the cap.
    /// </summary>
    internal static DocxSheet DocxSheetLayout(OfficeCli.Handlers.WordHandler wordHandler, string? pageFilter, int gridCols,
        int screenshotWidth, int screenshotHeight)
    {
        const int gap = DocxSheetGap, pad = DocxSheetPad, maxDim = 1920, scrollbar = 17;
        var (npW, npH) = wordHandler.GetPageNativePixels();

        int pageCount = DocxPreviewPageCount(wordHandler) ?? 1;

        // Pages actually tiled: the --page selection clamped to the document,
        // or every page when no filter was given. A selection that starts past
        // the last page is an error, not an empty sheet.
        ValidateDocxPageSpec(pageFilter);
        var selectedPages = ParseDocxPageSelection(pageFilter, pageCount);
        if (!string.IsNullOrWhiteSpace(pageFilter)
            && (selectedPages.Count == 0 || ParseDocxPageSelection(pageFilter, int.MaxValue).Count != selectedPages.Count))
            throw new OfficeCli.Core.CliException($"--page {pageFilter} out of range (total pages: {pageCount}).") { Code = "out_of_range" };
        int tiled = selectedPages.Count > 0 ? selectedPages.Count : pageCount;
        string nativePageSpec = selectedPages.Count > 0 ? string.Join(",", selectedPages) : $"1-{pageCount}";

        // -1 = auto: columns that keep the sheet ≈ square; 0 (range without --grid) = one column.
        int cols = gridCols < 0 ? OfficeCli.Core.HtmlScreenshot.AutoGridColumns(tiled, npW, npH)
            : gridCols == 0 ? 1 : gridCols;
        int rows = Math.Max(1, (tiled + cols - 1) / cols);

        double vpW = screenshotWidth == 1600 && screenshotHeight == 1200 && cols == 1
            ? npW + scrollbar + 2.0 * pad : screenshotWidth;
        double cellW = Math.Max(1.0, (vpW - scrollbar - 2.0 * pad - (cols - 1) * gap) / cols);
        double cellH = cellW * npH / npW;
        double vpH = pad * 2 + rows * cellH + (rows - 1) * gap;
        double over = Math.Max(vpW, vpH) / maxDim;
        if (over > 1.0) { vpW /= over; cellW /= over; cellH /= over; vpH /= over; }

        var html = new OfficeCli.Core.Rendering.RenderOptions
        { PageFilter = selectedPages.Count > 0 ? pageFilter : null, GridColumns = cols, GridCellWidthPx = (int)Math.Round(cellW) };
        return new DocxSheet(cols, (int)Math.Round(cellW), (int)Math.Round(cellH),
            Math.Max(1, (int)Math.Round(vpW)), Math.Max(1, (int)Math.Ceiling(vpH)), nativePageSpec, html);
    }

    /// <summary>
    /// Single docx page through the HTML path: size the viewport to the page's
    /// 96-DPI native pixels (default dims) or aspect-match a custom width.
    /// </summary>
    internal static void DocxSinglePageViewport(OfficeCli.Handlers.WordHandler wordHandler, string? effectiveFilter, int gridCols,
        ref int screenshotWidth, ref int screenshotHeight)
    {
        if (!int.TryParse(effectiveFilter, out _) || gridCols != 0) return;
        var (nativeW, nativeH) = wordHandler.GetPageNativePixels();
        if (screenshotWidth == 1600 && screenshotHeight == 1200)
            (screenshotWidth, screenshotHeight) = (nativeW, nativeH);
        else if (screenshotHeight == 1200)
            screenshotHeight = Math.Max(1, (int)Math.Round(screenshotWidth * (double)nativeH / nativeW));
    }

    /// <summary>True when a --page filter names more than one page (a range or a list).</summary>
    internal static bool IsMultiPageFilter(string? pageFilter)
    {
        if (string.IsNullOrWhiteSpace(pageFilter)) return false;
        var pages = ParseDocxPageSelection(pageFilter, int.MaxValue);
        return pages.Count > 1;
    }

    /// <summary>
    /// Reject a malformed docx --page spec up front (non-numeric, zero or
    /// negative pages, a range whose end precedes its start) instead of
    /// rendering every page or a blank page with exit 0.
    /// </summary>
    internal static void ValidateDocxPageSpec(string? pageFilter)
    {
        if (string.IsNullOrWhiteSpace(pageFilter)) return;
        foreach (var part in pageFilter.Split(','))
        {
            var t = part.Trim();
            if (t.Length == 0)
                throw new OfficeCli.Core.CliException($"Invalid --page value '{pageFilter}': empty list item.") { Code = "invalid_value" };
            if (t.Contains('-'))
            {
                var r = t.Split('-', 2);
                if (!int.TryParse(r[0].Trim(), out var from) || !int.TryParse(r[1].Trim(), out var to))
                    throw new OfficeCli.Core.CliException($"Invalid --page value '{pageFilter}': expected N, M-N or a comma list.") { Code = "invalid_value" };
                if (from <= 0 || to <= 0)
                    throw new OfficeCli.Core.CliException($"Invalid --page value '{pageFilter}': page number must be >= 1.") { Code = "invalid_value" };
                if (to < from)
                    throw new OfficeCli.Core.CliException($"Invalid --page value '{pageFilter}': range end is before its start.") { Code = "invalid_value", Suggestion = $"Use --page {to}-{from}." };
            }
            else if (!int.TryParse(t, out var n))
                throw new OfficeCli.Core.CliException($"Invalid --page value '{pageFilter}': expected a positive page number.") { Code = "invalid_value" };
            else if (n <= 0)
                throw new OfficeCli.Core.CliException($"Invalid --page value '{pageFilter}': page number must be >= 1.") { Code = "invalid_value" };
        }
    }

    /// <summary>
    /// Page count of a docx as the HTML preview paginates it (a real layout
    /// pass; the preview publishes the count via &lt;title&gt;PAGES:N&gt;).
    /// Null when the count could not be read.
    /// </summary>
    internal static int? DocxPreviewPageCount(OfficeCli.Handlers.WordHandler wordHandler)
    {
        var tmp = Path.Combine(Path.GetTempPath(), $"officecli_gridcount_{Guid.NewGuid():N}.html");
        try
        {
            File.WriteAllText(tmp, RenderViaRegistry(wordHandler, "docx", new OfficeCli.Core.Rendering.RenderOptions())!);
            return OfficeCli.Core.HtmlScreenshot.GetPageCountFromDom(tmp);
        }
        catch { return null; }
        finally { try { File.Delete(tmp); } catch { /* ignore */ } }
    }

    /// <summary>
    /// Pages exist only after the preview's script has laid the document out,
    /// so <see cref="DocxPreviewPageCount"/> needs a chrome-family headless
    /// browser. Without one a --page past the end cannot be rejected here, and
    /// refusing would also refuse pages that do exist. The preview is produced
    /// anyway — it shows an out-of-range notice itself when none of the
    /// requested pages exists — and the caller is told the check was skipped.
    /// </summary>
    internal static void WarnDocxPageNotChecked(string? pageFilter) =>
        Console.Error.WriteLine(
            $"Warning: --page {pageFilter?.Trim()} could not be checked against the page count " +
            "(no Chrome/Chromium/Edge available to count pages). " +
            "If the page does not exist, the preview shows an out-of-range notice instead of content.");

    /// <summary>
    /// Expand a docx --page filter ("1", "2-5", "1,3,5", "2-4,7") into the
    /// sorted distinct pages within 1..total. Unparsable parts are ignored,
    /// matching the preview's own lenient filter.
    /// </summary>
    internal static List<int> ParseDocxPageSelection(string? pageFilter, int total)
    {
        var pages = new SortedSet<int>();
        if (string.IsNullOrWhiteSpace(pageFilter)) return new List<int>();
        foreach (var part in pageFilter.Split(','))
        {
            var t = part.Trim();
            if (t.Contains('-'))
            {
                var r = t.Split('-', 2);
                if (int.TryParse(r[0].Trim(), out var from) && int.TryParse(r[1].Trim(), out var to))
                {
                    to = Math.Min(to, total);
                    for (int p = Math.Max(1, from); p <= to; p++) pages.Add(p);
                }
            }
            else if (int.TryParse(t, out var n) && n >= 1 && n <= total)
                pages.Add(n);
        }
        return pages.ToList();
    }

    private static (int? start, int? end) ParsePptHtmlPage(
        string? pageFilter, int? start, int? end,
        OfficeCli.Handlers.PowerPointHandler pptHandler)
    {
        if (string.IsNullOrEmpty(pageFilter)) return (start, end);
        var slideCount = pptHandler.Query("slide").Count;
        var spec = pageFilter.Trim();
        // A comma list cannot be rendered as one contiguous slide run; refuse
        // it instead of silently showing only the first item.
        if (spec.Contains(','))
            throw new OfficeCli.Core.CliException($"Invalid --page value '{pageFilter}': pptx takes a single slide (N) or a contiguous range (M-N), not a list.")
            { Code = "invalid_value", Suggestion = "Use --page M-N, or run one command per slide." };
        // Range form "M-N"
        if (spec.Contains('-'))
        {
            var parts = spec.Split('-', 2);
            if (!int.TryParse(parts[0].Trim(), out var ps) || !int.TryParse(parts[1].Trim(), out var pe))
                throw new OfficeCli.Core.CliException($"Invalid --page value '{pageFilter}': expected N or M-N.") { Code = "invalid_value" };
            if (ps <= 0 || pe <= 0)
                throw new OfficeCli.Core.CliException($"Invalid --page value '{pageFilter}': slide number must be >= 1.") { Code = "invalid_value" };
            if (pe < ps)
                throw new OfficeCli.Core.CliException($"Invalid --page value '{pageFilter}': range end is before its start.") { Code = "invalid_value", Suggestion = $"Use --page {pe}-{ps}." };
            if (ps > slideCount || pe > slideCount)
                throw new OfficeCli.Core.CliException($"--page {pageFilter} out of range (total slides: {slideCount}).") { Code = "out_of_range" };
            return (ps, pe);
        }
        if (!int.TryParse(spec, out var p))
            throw new OfficeCli.Core.CliException($"Invalid --page value '{pageFilter}': expected a positive slide number.") { Code = "invalid_value" };
        if (p <= 0)
            throw new OfficeCli.Core.CliException($"Invalid --page value '{pageFilter}': slide number must be >= 1.") { Code = "invalid_value" };
        if (p > slideCount)
            throw new OfficeCli.Core.CliException($"--page {p} out of range (total slides: {slideCount}).") { Code = "out_of_range" };
        return (p, p);
    }
}
