// Copyright 2025 OfficeCli (officecli.ai)
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
        var viewModeArg = new Argument<string>("mode") { Description = "View mode: text, annotated, outline, stats, issues, html, svg, screenshot, forms" };
        var startLineOpt = new Option<int?>("--start") { Description = "Start line/paragraph number" };
        var endLineOpt = new Option<int?>("--end") { Description = "End line/paragraph number" };
        var maxLinesOpt = new Option<int?>("--max-lines") { Description = "Maximum number of lines/rows/slides to output (truncates with total count)" };
        var issueTypeOpt = new Option<string?>("--type") { Description = "Issue type filter: format, content, structure" };
        var limitOpt = new Option<int?>("--limit") { Description = "Limit number of results" };

        var colsOpt = new Option<string?>("--cols") { Description = "Column filter, comma-separated (Excel only, e.g. A,B,C)" };
        var pageOpt = new Option<string?>("--page") { Description = "Page filter for html mode (e.g. 1, 2-5, 1,3,5)" };
        var browserOpt = new Option<bool>("--browser") { Description = "Open output in browser (html / svg modes)" };
        var outOpt = new Option<string?>("--out", "-o") { Description = "Output file path (screenshot mode; defaults to a temp file)" };
        var screenshotWidthOpt = new Option<int>("--screenshot-width") { Description = "Screenshot viewport width (default 1600)", DefaultValueFactory = _ => 1600 };
        var screenshotHeightOpt = new Option<int>("--screenshot-height") { Description = "Screenshot viewport height (default 1200)", DefaultValueFactory = _ => 1200 };
        var gridOpt = new Option<int>("--grid") { Description = "Tile slides into an N-column thumbnail grid (screenshot mode, pptx only; 0 = off)", DefaultValueFactory = _ => 0 };

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
        viewCommand.Add(screenshotWidthOpt);
        viewCommand.Add(screenshotHeightOpt);
        viewCommand.Add(gridOpt);
        viewCommand.Add(jsonOption);

        viewCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(viewFileArg)!;
            var mode = result.GetValue(viewModeArg)!;
            var start = result.GetValue(startLineOpt);
            var end = result.GetValue(endLineOpt);
            var maxLines = result.GetValue(maxLinesOpt);
            var issueType = result.GetValue(issueTypeOpt);
            var limit = result.GetValue(limitOpt);
            var colsStr = result.GetValue(colsOpt);
            var pageFilter = result.GetValue(pageOpt);
            var browser = result.GetValue(browserOpt);
            var outArg = result.GetValue(outOpt);
            var screenshotWidth = result.GetValue(screenshotWidthOpt);
            var screenshotHeight = result.GetValue(screenshotHeightOpt);
            var gridCols = result.GetValue(gridOpt);

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
                    html = pptHandler.ViewAsHtml(pStart, pEnd);
                }
                else if (handler is OfficeCli.Handlers.ExcelHandler excelHandler)
                    html = excelHandler.ViewAsHtml();
                else if (handler is OfficeCli.Handlers.WordHandler wordHandler)
                    html = wordHandler.ViewAsHtml(pageFilter);

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
                string? html = null;
                if (handler is OfficeCli.Handlers.PowerPointHandler pptHandler)
                {
                    var (pStart, pEnd) = ParsePptHtmlPage(pageFilter, start, end, pptHandler);
                    html = pptHandler.ViewAsHtml(pStart, pEnd, gridCols, screenshotWidth);
                }
                else if (handler is OfficeCli.Handlers.ExcelHandler excelHandler)
                    html = excelHandler.ViewAsHtml();
                else if (handler is OfficeCli.Handlers.WordHandler wordHandler)
                    html = wordHandler.ViewAsHtml(pageFilter);

                if (html == null)
                {
                    throw new OfficeCli.Core.CliException("Screenshot mode is only supported for .pptx, .xlsx, and .docx files.")
                    {
                        Code = "unsupported_type",
                        Suggestion = "Use a .pptx, .xlsx, or .docx file.",
                        ValidValues = ["text", "annotated", "outline", "stats", "issues", "html", "svg", "screenshot"]
                    };
                }

                // SECURITY: random token in temp filename — same rationale as the html/--browser path.
                var tmpHtml = Path.Combine(Path.GetTempPath(), $"officecli_preview_{Path.GetFileNameWithoutExtension(file.Name)}_{DateTime.Now:HHmmss}_{Guid.NewGuid():N}.html");
                File.WriteAllText(tmpHtml, html);
                var pngPath = outArg ?? Path.Combine(Path.GetTempPath(), $"officecli_screenshot_{Path.GetFileNameWithoutExtension(file.Name)}_{DateTime.Now:HHmmss}_{Guid.NewGuid():N}.png");
                var r = OfficeCli.Core.HtmlScreenshot.Capture(tmpHtml, pngPath, screenshotWidth, screenshotHeight);
                try { File.Delete(tmpHtml); } catch { /* ignore */ }
                if (!r.Ok)
                {
                    throw new OfficeCli.Core.CliException(
                        "No headless browser available. Install Chrome/Edge/Chromium or Firefox, or `pip install playwright && playwright install chromium`."
                        + (r.Error != null ? $" Last error: {r.Error}" : ""))
                    { Code = "no_screenshot_backend" };
                }
                Console.WriteLine(Path.GetFullPath(pngPath));
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

            if (json)
            {
                // Structured JSON output — no Content string wrapping
                var modeKey = mode.ToLowerInvariant();
                if (modeKey is "stats" or "s")
                    Console.WriteLine(OutputFormatter.WrapEnvelope(handler.ViewAsStatsJson().ToJsonString(OutputFormatter.PublicJsonOptions)));
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
                    else
                        throw new OfficeCli.Core.CliException("Forms view is only supported for .docx files.")
                        {
                            Code = "unsupported_type",
                            ValidValues = ["text", "annotated", "outline", "stats", "issues", "html", "svg", "screenshot", "forms"]
                        };
                }
                else
                    throw new OfficeCli.Core.CliException($"Unknown mode: {mode}. Available: text, annotated, outline, stats, issues, html, svg, screenshot, forms")
                    {
                        Code = "invalid_value",
                        ValidValues = ["text", "annotated", "outline", "stats", "issues", "html", "svg", "screenshot", "forms"]
                    };
            }
            else
            {
                var output = mode.ToLowerInvariant() switch
                {
                    "text" or "t" => handler.ViewAsText(start, end, maxLines, cols),
                    "annotated" or "a" => handler.ViewAsAnnotated(start, end, maxLines, cols),
                    "outline" or "o" => handler.ViewAsOutline(),
                    "stats" or "s" => handler.ViewAsStats(),
                    "issues" or "i" => OutputFormatter.FormatIssues(handler.ViewAsIssues(issueType, limit), OutputFormat.Text),
                    "forms" or "f" => handler is OfficeCli.Handlers.WordHandler wfh
                        ? wfh.ViewAsForms()
                        : throw new OfficeCli.Core.CliException("Forms view is only supported for .docx files.")
                        {
                            Code = "unsupported_type",
                            ValidValues = ["text", "annotated", "outline", "stats", "issues", "html", "svg", "screenshot", "forms"]
                        },
                    _ => throw new OfficeCli.Core.CliException($"Unknown mode: {mode}. Available: text, annotated, outline, stats, issues, html, svg, screenshot, forms")
                    {
                        Code = "invalid_value",
                        ValidValues = ["text", "annotated", "outline", "stats", "issues", "html", "svg", "screenshot", "forms"]
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
}
