// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using OfficeCli.Core;
using OfficeCli.Handlers;
using OfficeCli.Handlers.Hwp;

namespace OfficeCli;

static partial class CommandBuilder
{
    private static Command BuildViewCommand(Option<bool> jsonOption)
    {
        var viewFileArg = new Argument<FileInfo>("file") { Description = "Office document path (.docx, .xlsx, .pptx, .hwpx, experimental .hwp)" };
        var viewModeArg = new Argument<string>("mode") { Description = "View mode: text, annotated, outline, stats, issues, html, svg, screenshot, pdf, forms, styles, tables, markdown, objects, fields, field, native" };
        var startLineOpt = new Option<int?>("--start") { Description = "Start line/paragraph number" };
        var endLineOpt = new Option<int?>("--end") { Description = "End line/paragraph number" };
        var maxLinesOpt = new Option<int?>("--max-lines") { Description = "Maximum number of lines/rows/slides to output (truncates with total count)" };
        var issueTypeOpt = new Option<string?>("--type") { Description = IssueSubtypes.TypeHelpDescription() };
        var limitOpt = new Option<int?>("--limit") { Description = "Limit number of results" };

        var colsOpt = new Option<string?>("--cols") { Description = "Column filter, comma-separated (Excel only, e.g. A,B,C)" };
        var pageOpt = new Option<string?>("--page") { Description = "Page filter (e.g. 1, 2-5, 1,3,5). html mode: default=all. screenshot mode: default=1 (use --page 1-N to capture more, or --grid N for pptx thumbnails)." };
        var browserOpt = new Option<bool>("--browser") { Description = "Open output in browser or image viewer (html / svg / screenshot modes)" };
        var outOpt = new Option<string?>("--out", "-o") { Description = "Output file path (screenshot mode; defaults to a temp file)" };
        var screenshotWidthOpt = new Option<int>("--screenshot-width") { Description = "Screenshot viewport width (default 1600)", DefaultValueFactory = _ => 1600 };
        var screenshotHeightOpt = new Option<int>("--screenshot-height") { Description = "Screenshot viewport height (default 1200)", DefaultValueFactory = _ => 1200 };
        var gridOpt = new Option<int>("--grid") { Description = "Tile slides into an N-column thumbnail grid (screenshot mode, pptx only; 0 = off)", DefaultValueFactory = _ => 0 };
        var renderOpt = new Option<string>("--render") { Description = "Screenshot rendering path (docx only): auto (default; native on Windows w/ Word, html elsewhere), native (force OS-native, error if unavailable), html", DefaultValueFactory = _ => "auto" };
        var withPagesOpt = new Option<bool>("--page-count") { Description = "stats mode (docx only): also report total page count via Word repagination (Win + Word required; slow on long docs)" };
        var autoOpt = new Option<bool>("--auto") { Description = "Auto-recognize label-value fields in tables (hwpx forms only)" };
        var objectTypeOpt = new Option<string?>("--object-type") { Description = "Object type filter: picture, field, bookmark, equation, formfield (hwpx objects mode)" };
        var nativeOpOpt = new Option<string?>("--op") { Description = "HWP rhwp native read operation for native view mode" };
        var nativeArgOpt = new Option<string[]>("--native-arg") { Description = "HWP native view argument (key=value), repeatable", AllowMultipleArgumentsPerToken = true };
        var fieldNameOpt = new Option<string?>("--field-name") { Description = "Field name for HWP/HWPX field read mode" };
        var fieldIdOpt = new Option<int?>("--field-id") { Description = "Field id for HWP/HWPX field read mode" };
        var sectionOpt = new Option<int?>("--section") { Description = "HWP rhwp section index for table/page operations" };
        var parentParaOpt = new Option<int?>("--parent-para") { Description = "HWP rhwp parent paragraph index for table operations" };
        var controlOpt = new Option<int?>("--control") { Description = "HWP rhwp control index for table operations" };
        var cellOpt = new Option<int?>("--cell") { Description = "HWP rhwp cell index for table operations" };
        var cellParaOpt = new Option<int?>("--cell-para") { Description = "HWP rhwp cell paragraph index for table operations" };
        var offsetOpt = new Option<int?>("--offset") { Description = "HWP rhwp text offset for table cell read" };
        var countOpt = new Option<int?>("--count") { Description = "HWP rhwp count/limit for table cell read" };
        var maxParentParaOpt = new Option<int?>("--max-parent-para") { Description = "HWP rhwp scan upper bound for parent paragraphs" };
        var maxControlOpt = new Option<int?>("--max-control") { Description = "HWP rhwp scan upper bound for controls" };
        var maxCellOpt = new Option<int?>("--max-cell") { Description = "HWP rhwp scan upper bound for cells" };
        var maxCellParaOpt = new Option<int?>("--max-cell-para") { Description = "HWP rhwp scan upper bound for cell paragraphs" };
        var includeEmptyOpt = new Option<bool>("--include-empty") { Description = "Include empty HWP table cells in scan output" };

        var viewCommand = new Command("view", BuildViewDescription());
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
        viewCommand.Add(screenshotWidthOpt);
        viewCommand.Add(screenshotHeightOpt);
        viewCommand.Add(gridOpt);
        viewCommand.Add(renderOpt);
        viewCommand.Add(withPagesOpt);
        viewCommand.Add(autoOpt);
        viewCommand.Add(objectTypeOpt);
        viewCommand.Add(nativeOpOpt);
        viewCommand.Add(nativeArgOpt);
        viewCommand.Add(fieldNameOpt);
        viewCommand.Add(fieldIdOpt);
        viewCommand.Add(sectionOpt);
        viewCommand.Add(parentParaOpt);
        viewCommand.Add(controlOpt);
        viewCommand.Add(cellOpt);
        viewCommand.Add(cellParaOpt);
        viewCommand.Add(offsetOpt);
        viewCommand.Add(countOpt);
        viewCommand.Add(maxParentParaOpt);
        viewCommand.Add(maxControlOpt);
        viewCommand.Add(maxCellOpt);
        viewCommand.Add(maxCellParaOpt);
        viewCommand.Add(includeEmptyOpt);
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
            var screenshotWidth = result.GetValue(screenshotWidthOpt);
            var screenshotHeight = result.GetValue(screenshotHeightOpt);
            var gridCols = result.GetValue(gridOpt);
            var renderMode = (result.GetValue(renderOpt) ?? "auto").ToLowerInvariant();
            if (renderMode is not ("auto" or "native" or "html"))
                throw new OfficeCli.Core.CliException($"Invalid --render value: {renderMode}. Valid: auto, native, html") { Code = "invalid_render", ValidValues = ["auto", "native", "html"] };
            var withPages = result.GetValue(withPagesOpt);
            var autoRecognize = result.GetValue(autoOpt);
            var objectTypeFilter = result.GetValue(objectTypeOpt);
            var nativeOp = result.GetValue(nativeOpOpt);
            var nativeArgs = result.GetValue(nativeArgOpt);
            var fieldName = result.GetValue(fieldNameOpt);
            var fieldId = result.GetValue(fieldIdOpt);
            var hwpViewArgs = new Dictionary<string, string>(StringComparer.Ordinal);
            AddHwpViewOption(hwpViewArgs, "--section", result.GetValue(sectionOpt));
            AddHwpViewOption(hwpViewArgs, "--parent-para", result.GetValue(parentParaOpt));
            AddHwpViewOption(hwpViewArgs, "--control", result.GetValue(controlOpt));
            AddHwpViewOption(hwpViewArgs, "--cell", result.GetValue(cellOpt));
            AddHwpViewOption(hwpViewArgs, "--cell-para", result.GetValue(cellParaOpt));
            AddHwpViewOption(hwpViewArgs, "--offset", result.GetValue(offsetOpt));
            AddHwpViewOption(hwpViewArgs, "--count", result.GetValue(countOpt));
            AddHwpViewOption(hwpViewArgs, "--max-parent-para", result.GetValue(maxParentParaOpt));
            AddHwpViewOption(hwpViewArgs, "--max-control", result.GetValue(maxControlOpt));
            AddHwpViewOption(hwpViewArgs, "--max-cell", result.GetValue(maxCellOpt));
            AddHwpViewOption(hwpViewArgs, "--max-cell-para", result.GetValue(maxCellParaOpt));
            if (result.GetValue(includeEmptyOpt))
                hwpViewArgs["--include-empty"] = "true";

            // pdf mode runs entirely through an exporter plugin (no handler
            // open, no resident hop — the plugin gets a snapshot of the
            // source and writes the PDF). Handled before TryResident
            // because exporter invocation needs the file lock released, and
            // ExporterInvoker closes the resident itself when present.
            var lowerMode = mode.ToLowerInvariant();
            var earlyExtension = Path.GetExtension(file.FullName);
            var bridgeOwnsPdf = string.Equals(earlyExtension, ".hwp", StringComparison.OrdinalIgnoreCase)
                || (string.Equals(earlyExtension, ".hwpx", StringComparison.OrdinalIgnoreCase)
                    && (HwpEngineSelector.IsExperimentalBridgeEnabled()
                        || HwpEngineSelector.CanUseInstalledRuntime(
                            HwpCapabilityConstants.FormatHwpx,
                            HwpCapabilityConstants.OperationExportPdf)));
            if (lowerMode is "pdf" && !bridgeOwnsPdf)
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
                req.Args["screenshot-width"] = screenshotWidth.ToString();
                req.Args["screenshot-height"] = screenshotHeight.ToString();
                if (gridCols > 0) req.Args["grid"] = gridCols.ToString();
                if (renderMode != "auto") req.Args["render"] = renderMode;
                if (withPages) req.Args["page-count"] = "true";
                if (autoRecognize) req.Args["auto"] = "true";
                if (objectTypeFilter != null) req.Args["object-type"] = objectTypeFilter;
            }, json) is {} rc) return rc;

            var format = json ? OutputFormat.Json : OutputFormat.Text;
            var cols = colsStr != null ? new HashSet<string>(colsStr.Split(',').Select(c => c.Trim().ToUpperInvariant())) : null;

            var extension = Path.GetExtension(file.FullName);

            // Binary .hwp: route through HWP engine (bridge when experimental, else unsupported)
            if (string.Equals(extension, ".hwp", StringComparison.OrdinalIgnoreCase))
                return HandleHwpView(file.FullName, HwpFormat.Hwp, mode, pageFilter, json, fieldName, fieldId, outArg, hwpViewArgs, nativeOp, nativeArgs);

            // HWPX stays on the custom XML handler by default. The rhwp bridge can be
            // opted into for read/render smoke coverage without changing stable HWPX behavior.
            var hwpxModeKey = mode.Trim().ToLowerInvariant();
            var hwpxOperation = HwpViewOperationForMode(hwpxModeKey);
            if (string.Equals(extension, ".hwpx", StringComparison.OrdinalIgnoreCase)
                && hwpxOperation != null
                && (HwpEngineSelector.IsExperimentalBridgeEnabled()
                    || HwpEngineSelector.CanUseInstalledRuntime(
                        HwpCapabilityConstants.FormatHwpx,
                        hwpxOperation)))
                return HandleHwpView(file.FullName, HwpFormat.Hwpx, mode, pageFilter, json, fieldName, fieldId, outArg, hwpViewArgs, nativeOp, nativeArgs);

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
                    html = pptHandler.ViewAsHtml(pStart, pEnd);
                }
                else if (handler is OfficeCli.Handlers.ExcelHandler excelHandler)
                    html = excelHandler.ViewAsHtml();
                else if (handler is OfficeCli.Handlers.WordHandler wordHandler)
                    html = wordHandler.ViewAsHtml(pageFilter);
                else if (handler is OfficeCli.Handlers.HwpxHandler hwpxHandler)
                    html = hwpxHandler.ViewAsHtml();
                else if (handler is OfficeCli.Core.Plugins.FormatHandlerProxy proxy)
                    html = proxy.ViewAsHtml(int.TryParse(pageFilter, out var p) ? p : (int?)null);

                if (html != null)
                {
                    if (browser)
                    {
                        // --browser: write to temp file and open in browser
                        // SECURITY: include a random token so the preview path is not predictable.
                        // A predictable path (HHmmss only) lets a local attacker pre-place a symlink
                        // at the expected location, causing File.WriteAllText to follow it and
                        // overwrite an arbitrary victim file with preview HTML. It also caused
                        // collisions between concurrent `view html` invocations of the same file.
                        var htmlPath = Path.Combine(Path.GetTempPath(), $"officecli_preview_{Path.GetFileNameWithoutExtension(file.Name)}_{DateTime.Now:HHmmss}_{Guid.NewGuid():N}.html");
                        File.WriteAllText(htmlPath, html);
                        Console.WriteLine(htmlPath);
                        try
                        {
                            var psi = new System.Diagnostics.ProcessStartInfo(htmlPath) { UseShellExecute = true };
                            System.Diagnostics.Process.Start(psi);
                        }
                        catch { /* silently ignore if browser can't be opened */ }
                    }
                    else
                    {
                        // Default: output HTML to stdout
                        Console.Write(html);
                    }
                }
                else
                {
                    throw new OfficeCli.Core.CliException("HTML preview is only supported for .pptx, .xlsx, .docx, and .hwpx files.")
                    {
                        Code = "unsupported_type",
                        Suggestion = "Use a .pptx, .xlsx, .docx, or .hwpx file, or use mode 'text' or 'annotated' for other formats.",
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
                if (handler is OfficeCli.Handlers.PowerPointHandler pptHandler)
                {
                    var effectiveFilter = pageFilter;
                    if (string.IsNullOrEmpty(effectiveFilter) && start is null && end is null && gridCols == 0)
                        effectiveFilter = "1";
                    var (pStart, pEnd) = ParsePptHtmlPage(effectiveFilter, start, end, pptHandler);
                    html = pptHandler.ViewAsHtml(pStart, pEnd, gridCols, screenshotWidth);
                }
                else if (handler is OfficeCli.Handlers.ExcelHandler excelHandler)
                    html = excelHandler.ViewAsHtml();
                else if (handler is OfficeCli.Handlers.WordHandler wordHandler)
                {
                    var effectiveFilter = string.IsNullOrEmpty(pageFilter) ? "1" : pageFilter;
                    if (renderMode != "html" && OperatingSystem.IsWindows())
                    {
                        try { directPng = OfficeCli.Core.WordPdfBackend.Render(file.FullName, effectiveFilter); }
                        catch { directPng = null; }
                    }
                    if (renderMode == "native" && directPng == null)
                        throw new OfficeCli.Core.CliException("--render native requires Windows with Microsoft Word installed.")
                        { Code = "native_unavailable", Suggestion = "Use --render html or --render auto." };
                    if (directPng == null) html = wordHandler.ViewAsHtml(effectiveFilter);
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
                    var r = OfficeCli.Core.HtmlScreenshot.Capture(tmpHtml, pngPath, screenshotWidth, screenshotHeight);
                    try { File.Delete(tmpHtml); } catch { /* ignore */ }
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
                    var svg = pptSvgHandler.ViewAsSvg(slideNum);

                    if (browser)
                    {
                        string outPath;
                        if (svg.Contains("data-formula"))
                        {
                            // Wrap SVG in HTML shell for KaTeX formula rendering
                            outPath = Path.Combine(Path.GetTempPath(), $"officecli_slide{slideNum}_{Path.GetFileNameWithoutExtension(file.Name)}_{DateTime.Now:HHmmss}.html");
                            var html = $"<!DOCTYPE html><html><head><meta charset='UTF-8'><link rel='stylesheet' href='https://cdn.jsdelivr.net/npm/katex@0.16.11/dist/katex.min.css'><script defer src='https://cdn.jsdelivr.net/npm/katex@0.16.11/dist/katex.min.js'></script><style>body{{margin:0;display:flex;justify-content:center;background:#f0f0f0}}</style></head><body>{svg}<script>window.addEventListener('load',function(){{document.querySelectorAll('[data-formula]').forEach(function(el){{try{{katex.render(el.getAttribute('data-formula'),el,{{throwOnError:false,displayMode:true}})}}catch(e){{}}}})}})</script></body></html>";
                            File.WriteAllText(outPath, html);
                        }
                        else
                        {
                            outPath = Path.Combine(Path.GetTempPath(), $"officecli_slide{slideNum}_{Path.GetFileNameWithoutExtension(file.Name)}_{DateTime.Now:HHmmss}.svg");
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
                        File.WriteAllText(tmpHtml, wordHandlerForCount.ViewAsHtml(null));
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
                    Console.WriteLine(OutputFormatter.WrapEnvelope(handler.ViewAsTextJson(start, end, maxLines, cols).ToJsonString(OutputFormatter.PublicJsonOptions)));
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
                    else if (handler is OfficeCli.Handlers.HwpxHandler hwpxFormsHandler)
                        Console.WriteLine(OutputFormatter.WrapEnvelope(hwpxFormsHandler.ViewAsFormsJson(autoRecognize).ToJsonString(OutputFormatter.PublicJsonOptions)));
                    else if (handler is OfficeCli.Core.Plugins.FormatHandlerProxy formsProxy)
                    {
                        var formsJson = formsProxy.ViewAsFormsJson();
                        if (formsJson is null)
                            throw new OfficeCli.Core.CliException($"Forms view is not supported by the format-handler plugin for {file.Extension}.")
                            { Code = "unsupported_type" };
                        Console.WriteLine(OutputFormatter.WrapEnvelope(formsJson.ToJsonString(OutputFormatter.PublicJsonOptions)));
                    }
                    else
                        throw new OfficeCli.Core.CliException("Forms view is only supported for .docx and .hwpx files.")
                        {
                            Code = "unsupported_type",
                            ValidValues = ["text", "annotated", "outline", "stats", "issues", "html", "svg", "screenshot", "pdf", "forms", "tables", "objects"]
                        };
                }
                else if (modeKey is "tables" or "tbl")
                {
                    if (handler is OfficeCli.Handlers.HwpxHandler hwpxTblHandler)
                        Console.WriteLine(OutputFormatter.WrapEnvelope(hwpxTblHandler.ViewAsTablesJson().ToJsonString(OutputFormatter.PublicJsonOptions)));
                    else
                        throw new OfficeCli.Core.CliException("Tables view is only supported for .hwpx files.")
                        { Code = "unsupported_type" };
                }
                else if (modeKey is "objects" or "obj")
                {
                    if (handler is OfficeCli.Handlers.HwpxHandler hwpxObjHandler)
                        Console.WriteLine(OutputFormatter.WrapEnvelope(hwpxObjHandler.ViewAsObjectsJson(objectTypeFilter).ToJsonString(OutputFormatter.PublicJsonOptions)));
                    else
                        throw new OfficeCli.Core.CliException("Objects view is only supported for .hwpx files.")
                        { Code = "unsupported_type" };
                }
                else
                    throw new OfficeCli.Core.CliException($"Unknown mode: {mode}. Available: text, annotated, outline, stats, issues, html, svg, screenshot, pdf, forms, tables, objects")
                    {
                        Code = "invalid_value",
                        ValidValues = ["text", "annotated", "outline", "stats", "issues", "html", "svg", "screenshot", "pdf", "forms", "tables", "objects"]
                    };
            }
            else
            {
                var output = mode.ToLowerInvariant() switch
                {
                    "text" or "t" => handler.ViewAsText(start, end, maxLines, cols),
                    "annotated" or "a" => handler.ViewAsAnnotated(start, end, maxLines, cols),
                    "outline" or "o" => handler.ViewAsOutline(),
                    "stats" or "s" => withPagesValue.HasValue
                        ? $"Pages: {withPagesValue}\n" + handler.ViewAsStats()
                        : handler.ViewAsStats(),
                    "issues" or "i" => OutputFormatter.FormatIssues(handler.ViewAsIssues(issueType, limit), OutputFormat.Text),
                    "forms" or "f" => handler switch
                    {
                        OfficeCli.Handlers.WordHandler wfh => wfh.ViewAsForms(),
                        OfficeCli.Handlers.HwpxHandler hfh => hfh.ViewAsForms(autoRecognize),
                        OfficeCli.Core.Plugins.FormatHandlerProxy fp
                            => fp.ViewAsFormsJson()?.ToJsonString(OutputFormatter.PublicJsonOptions)
                               ?? throw new OfficeCli.Core.CliException($"Forms view is not supported by the format-handler plugin for {file.Extension}.")
                                   { Code = "unsupported_type" },
                        _ => throw new OfficeCli.Core.CliException("Forms view is only supported for .docx, .hwpx, or a plugin that supports forms view.")
                        {
                            Code = "unsupported_type",
                            ValidValues = ["text", "annotated", "outline", "stats", "issues", "html", "svg", "screenshot", "pdf", "forms", "tables", "markdown", "objects", "styles"]
                        }
                    },
                    "styles" => handler is OfficeCli.Handlers.HwpxHandler hsh
                        ? hsh.ViewAsStyles()
                        : throw new OfficeCli.Core.CliException("Styles view is only supported for .hwpx files.")
                        {
                            Code = "unsupported_type",
                            ValidValues = ["text", "annotated", "outline", "stats", "issues", "html", "styles", "tables"]
                        },
                    "tables" or "tbl" => handler is OfficeCli.Handlers.HwpxHandler htbl
                        ? htbl.ViewAsTables()
                        : throw new OfficeCli.Core.CliException("Tables view is only supported for .hwpx files.")
                        {
                            Code = "unsupported_type",
                            ValidValues = ["text", "annotated", "outline", "stats", "issues", "html", "styles", "tables", "markdown"]
                        },
                    "markdown" or "md" => handler is OfficeCli.Handlers.HwpxHandler hmd
                        ? hmd.ViewAsMarkdown()
                        : throw new OfficeCli.Core.CliException("Markdown view is only supported for .hwpx files.")
                        {
                            Code = "unsupported_type",
                            ValidValues = ["text", "annotated", "outline", "stats", "issues", "html", "styles", "tables", "markdown", "objects"]
                        },
                    "objects" or "obj" => handler is OfficeCli.Handlers.HwpxHandler hobj
                        ? hobj.ViewAsObjects(objectTypeFilter)
                        : throw new OfficeCli.Core.CliException("Objects view is only supported for .hwpx files.")
                        {
                            Code = "unsupported_type",
                            ValidValues = ["text", "annotated", "outline", "stats", "issues", "html", "styles", "tables", "markdown", "objects"]
                        },
                    _ => throw new OfficeCli.Core.CliException($"Unknown mode: {mode}. Available: text, annotated, outline, stats, issues, html, svg, screenshot, pdf, forms, tables, markdown, objects, styles")
                    {
                        Code = "invalid_value",
                        ValidValues = ["text", "annotated", "outline", "stats", "issues", "html", "svg", "screenshot", "pdf", "forms", "tables", "markdown", "objects", "styles"]
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

    private static void AddHwpViewOption(Dictionary<string, string> args, string key, int? value)
    {
        if (value.HasValue)
            args[key] = value.Value.ToString();
    }

    private static int HandleHwpView(
        string filePath,
        HwpFormat format,
        string mode,
        string? pageFilter,
        bool json,
        string? fieldName = null,
        int? fieldId = null,
        string? outArg = null,
        IReadOnlyDictionary<string, string>? viewArgs = null,
        string? nativeOp = null,
        string[]? nativeArgs = null)
    {
        var modeKey = mode.Trim().ToLowerInvariant();
        var formatKey = format == HwpFormat.Hwp
            ? HwpCapabilityConstants.FormatHwp
            : HwpCapabilityConstants.FormatHwpx;
        var operation = HwpViewOperationForMode(modeKey);

        if (!HwpEngineSelector.IsExperimentalBridgeEnabled()
            && !HwpEngineSelector.CanUseInstalledRuntime(formatKey, operation))
        {
            var label = format == HwpFormat.Hwp ? "Binary .hwp" : "HWPX";
            throw new HwpEngineException(
                $"{label} bridge view requires packaged rhwp sidecars or OFFICECLI_HWP_ENGINE=rhwp-experimental.",
                HwpCapabilityConstants.ReasonBridgeNotEnabled,
                "Run ./dev-install.sh, or set OFFICECLI_HWP_ENGINE=rhwp-experimental and install rhwp-officecli-bridge.",
                [
                    HwpCapabilityConstants.OperationReadText,
                    HwpCapabilityConstants.OperationRenderSvg,
                    HwpCapabilityConstants.OperationRenderPng,
                    HwpCapabilityConstants.OperationExportPdf,
                    HwpCapabilityConstants.OperationExportMarkdown,
                    HwpCapabilityConstants.OperationThumbnail,
                    HwpCapabilityConstants.OperationDocumentInfo,
                    HwpCapabilityConstants.OperationDiagnostics,
                    HwpCapabilityConstants.OperationDumpControls,
                    HwpCapabilityConstants.OperationDumpPages,
                    HwpCapabilityConstants.OperationListFields,
                    HwpCapabilityConstants.OperationReadField,
                    HwpCapabilityConstants.OperationReadTableCell,
                    HwpCapabilityConstants.OperationScanCells,
                    HwpCapabilityConstants.OperationNativeRead
                ],
                formatKey,
                operation,
                HwpCapabilityConstants.EngineNone,
                HwpCapabilityConstants.ModeNone);
        }

        var engine = HwpEngineSelector.GetEngine(formatKey, operation);
        var fileInfo = new FileInfo(filePath);
        var ct = CancellationToken.None;

        if (modeKey is "text" or "t")
        {
            var request = new HwpReadRequest(format, filePath, fileInfo.Length, json);
            var result = engine.ReadTextAsync(request, ct).GetAwaiter().GetResult();
            if (json)
            {
                var envelope = new System.Text.Json.Nodes.JsonObject
                {
                    ["success"] = true,
                    ["data"] = new System.Text.Json.Nodes.JsonObject
                    {
                        ["text"] = result.Text,
                        ["engine"] = result.Engine,
                        ["engineVersion"] = result.EngineVersion
                    },
                    ["warnings"] = HwpCapabilityJsonMapper.ToJsonArray(result.Warnings)
                };
                Console.WriteLine(envelope.ToJsonString(OfficeCli.Core.OutputFormatter.PublicJsonOptions));
            }
            else
            {
                Console.WriteLine(result.Text);
            }
            return 0;
        }

        if (modeKey is "svg" or "g")
        {
            var outDir = Path.Combine(Path.GetTempPath(), $"officecli_hwp_svg_{Guid.NewGuid():N}");
            Directory.CreateDirectory(outDir);
            var request = new HwpRenderRequest(
                format, filePath, outDir,
                pageFilter ?? "all", fileInfo.Length, json);
            var result = engine.RenderSvgAsync(request, ct).GetAwaiter().GetResult();
            if (json)
            {
                var pagesArr = new System.Text.Json.Nodes.JsonArray();
                foreach (var p in result.Pages)
                    pagesArr.Add((System.Text.Json.Nodes.JsonNode?)new System.Text.Json.Nodes.JsonObject
                    {
                        ["page"] = p.Page, ["path"] = p.SvgPath, ["sha256"] = p.Sha256
                    });
                var envelope = new System.Text.Json.Nodes.JsonObject
                {
                    ["success"] = true,
                    ["data"] = new System.Text.Json.Nodes.JsonObject
                    {
                        ["pages"] = pagesArr,
                        ["manifest"] = result.ManifestPath,
                        ["engine"] = result.Engine,
                        ["engineVersion"] = result.EngineVersion
                    },
                    ["warnings"] = HwpCapabilityJsonMapper.ToJsonArray(result.Warnings)
                };
                Console.WriteLine(envelope.ToJsonString(OfficeCli.Core.OutputFormatter.PublicJsonOptions));
            }
            else
            {
                foreach (var p in result.Pages)
                    Console.WriteLine($"Page {p.Page}: {p.SvgPath}");
            }
            return 0;
        }

        if (modeKey is "png" or "pdf" or "markdown" or "md" or "thumbnail" or "info" or "diagnostics" or "diag" or "dump" or "controls" or "pages" or "dump-pages" or "table-cell" or "cell" or "tables" or "cells" or "native" or "native-op")
        {
            var args = new Dictionary<string, string>(StringComparer.Ordinal);
            if (viewArgs != null)
                foreach (var entry in viewArgs)
                    args[entry.Key] = entry.Value;
            string bridgeCommand;
            var effectiveOperation = operation ?? HwpCapabilityConstants.OperationReadText;
            if (modeKey is "png")
            {
                bridgeCommand = "render-png";
                args["--out-dir"] = outArg != null
                    ? Path.GetFullPath(outArg)
                    : Path.Combine(Path.GetTempPath(), $"officecli_hwp_png_{Guid.NewGuid():N}");
                args["--page"] = pageFilter ?? "all";
                Directory.CreateDirectory(args["--out-dir"]);
            }
            else if (modeKey is "pdf")
            {
                bridgeCommand = "export-pdf";
                args["--output"] = outArg != null
                    ? Path.GetFullPath(outArg)
                    : Path.GetFullPath(Path.ChangeExtension(filePath, ".pdf"));
                args["--page"] = pageFilter ?? "all";
            }
            else if (modeKey is "markdown" or "md")
            {
                bridgeCommand = "export-markdown";
                args["--page"] = pageFilter ?? "all";
            }
            else if (modeKey is "thumbnail")
            {
                bridgeCommand = "thumbnail";
                args["--output"] = outArg != null
                    ? Path.GetFullPath(outArg)
                    : Path.Combine(Path.GetTempPath(), $"officecli_hwp_thumbnail_{Guid.NewGuid():N}.png");
            }
            else if (modeKey is "info")
            {
                bridgeCommand = "document-info";
            }
            else if (modeKey is "diagnostics" or "diag")
            {
                bridgeCommand = "diagnostics";
            }
            else if (modeKey is "dump" or "controls")
            {
                bridgeCommand = "dump-controls";
            }
            else if (modeKey is "pages" or "dump-pages")
            {
                bridgeCommand = "dump-pages";
                if (!string.IsNullOrWhiteSpace(pageFilter))
                    args["--page"] = pageFilter;
            }
            else if (modeKey is "table-cell" or "cell")
            {
                bridgeCommand = "get-cell-text";
            }
            else if (modeKey is "native" or "native-op")
            {
                if (string.IsNullOrWhiteSpace(nativeOp))
                    throw new HwpEngineException(
                        "HWP native view requires --op <rhwp-native-op>.",
                        HwpCapabilityConstants.ReasonUnsupportedOperation,
                        "Example: officecli view input.hwp native --op get-style-list --json",
                        [HwpCapabilityConstants.OperationNativeRead],
                        formatKey,
                        HwpCapabilityConstants.OperationNativeRead,
                        HwpCapabilityConstants.EngineRhwpBridge,
                        HwpCapabilityConstants.ModeExperimental);
                ValidateHwpNativeViewRequest(formatKey, nativeOp, nativeArgs ?? Array.Empty<string>());
                bridgeCommand = "native-op";
                args["--op"] = nativeOp;
                foreach (var (key, value) in ParsePropsArray(nativeArgs ?? Array.Empty<string>()))
                {
                    var normalized = key.StartsWith("--", StringComparison.Ordinal) ? key : $"--{key}";
                    args[normalized] = value;
                }
            }
            else
            {
                bridgeCommand = "scan-cells";
            }

            var request = new HwpJsonViewRequest(format, filePath, fileInfo.Length, effectiveOperation, bridgeCommand, args, json);
            var result = engine.ViewJsonAsync(request, ct).GetAwaiter().GetResult();
            if (json)
            {
                var data = (System.Text.Json.Nodes.JsonObject)result.Data.DeepClone();
                data["engine"] = result.Engine;
                data["engineVersion"] = result.EngineVersion;
                var envelope = new System.Text.Json.Nodes.JsonObject
                {
                    ["success"] = true,
                    ["data"] = data,
                    ["warnings"] = HwpCapabilityJsonMapper.ToJsonArray(result.Warnings)
                };
                Console.WriteLine(envelope.ToJsonString(OfficeCli.Core.OutputFormatter.PublicJsonOptions));
            }
            else if (result.Data["markdown"]?.GetValue<string>() is { } markdown)
            {
                Console.WriteLine(markdown);
            }
            else if (result.Data["dump"]?.GetValue<string>() is { } dump)
            {
                Console.WriteLine(dump);
            }
            else if (result.Data["pdf"]?["path"]?.GetValue<string>() is { } pdfPath)
            {
                Console.WriteLine(pdfPath);
            }
            else
            {
                Console.WriteLine(result.Data.ToJsonString(OfficeCli.Core.OutputFormatter.PublicJsonOptions));
            }
            return 0;
        }

        if (modeKey is "fields")
        {
            var request = new HwpFieldListRequest(format, filePath, fileInfo.Length, json);
            var result = engine.ListFieldsAsync(request, ct).GetAwaiter().GetResult();
            if (json)
            {
                var envelope = new System.Text.Json.Nodes.JsonObject
                {
                    ["success"] = true,
                    ["data"] = result.Fields.DeepClone(),
                    ["engine"] = result.Engine,
                    ["engineVersion"] = result.EngineVersion,
                    ["warnings"] = HwpCapabilityJsonMapper.ToJsonArray(result.Warnings)
                };
                Console.WriteLine(envelope.ToJsonString(OfficeCli.Core.OutputFormatter.PublicJsonOptions));
            }
            else
            {
                Console.WriteLine(result.Fields.ToJsonString(OfficeCli.Core.OutputFormatter.PublicJsonOptions));
            }
            return 0;
        }

        if (modeKey is "field")
        {
            var request = new HwpFieldReadRequest(format, filePath, fieldName, fieldId, fileInfo.Length, json);
            var result = engine.ReadFieldAsync(request, ct).GetAwaiter().GetResult();
            if (json)
            {
                var envelope = new System.Text.Json.Nodes.JsonObject
                {
                    ["success"] = true,
                    ["data"] = result.Field.DeepClone(),
                    ["engine"] = result.Engine,
                    ["engineVersion"] = result.EngineVersion,
                    ["warnings"] = HwpCapabilityJsonMapper.ToJsonArray(result.Warnings)
                };
                Console.WriteLine(envelope.ToJsonString(OfficeCli.Core.OutputFormatter.PublicJsonOptions));
            }
            else
            {
                Console.WriteLine(result.Field.ToJsonString(OfficeCli.Core.OutputFormatter.PublicJsonOptions));
            }
            return 0;
        }

        throw new HwpEngineException(
            $"{formatKey} bridge view mode '{mode}' is not supported. Use text, svg, png, pdf, markdown, thumbnail, info, diagnostics, dump, pages, fields, field, table-cell, tables, or native.",
            HwpCapabilityConstants.ReasonUnsupportedOperation,
            null,
            [
                HwpCapabilityConstants.OperationReadText,
                HwpCapabilityConstants.OperationRenderSvg,
                HwpCapabilityConstants.OperationRenderPng,
                HwpCapabilityConstants.OperationExportPdf,
                HwpCapabilityConstants.OperationExportMarkdown,
                HwpCapabilityConstants.OperationThumbnail,
                HwpCapabilityConstants.OperationDocumentInfo,
                HwpCapabilityConstants.OperationDiagnostics,
                HwpCapabilityConstants.OperationDumpControls,
                HwpCapabilityConstants.OperationDumpPages,
                HwpCapabilityConstants.OperationListFields,
                HwpCapabilityConstants.OperationReadField,
                HwpCapabilityConstants.OperationReadTableCell,
                HwpCapabilityConstants.OperationScanCells,
                HwpCapabilityConstants.OperationNativeRead
            ],
            formatKey,
            null,
            HwpCapabilityConstants.EngineRhwpBridge,
            HwpCapabilityConstants.ModeExperimental);
    }
}
