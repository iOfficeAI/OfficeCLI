// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Diagnostics;
using System.Runtime.InteropServices;

namespace OfficeCli.Core;

/// <summary>
/// Headless HTML→PNG screenshot via shell-out to whichever browser is available.
/// Tries playwright CLI → Chromium-family (Chrome/Edge/Chromium) → Firefox.
/// No embedded browser engine; binary stays small.
/// </summary>
internal static class HtmlScreenshot
{
    public sealed record Result(bool Ok, string Backend, string? Error);

    public sealed record PaginationResult(int TotalPages, Dictionary<string, int> AnchorPageMap);

    /// Run a chromium-family browser in dump-dom mode against the given HTML
    /// and parse the document title for "PAGES:N|MAP:anchor=p,anchor=p,...".
    /// The HTML must set the title from JS after layout settles.
    /// <summary>
    /// Render <paramref name="htmlPath"/> in a chrome-family browser with a
    /// virtual-time budget (so async JS such as mermaid.js finishes) and return
    /// the final serialized DOM, or null if no chrome-family browser is found or
    /// the run fails. Callers extract whatever they need from the DOM (title,
    /// a rendered &lt;svg&gt;, …).
    /// </summary>
    public static string? DumpDom(string htmlPath, int timeoutMs = 60000, string[]? extraArgs = null)
    {
        var url = new Uri(Path.GetFullPath(htmlPath)).AbsoluteUri + "#screenshot";
        var bin = FindChrome();
        if (bin == null) return null;
        var args = new List<string>
        {
            "--headless=new",
            "--disable-gpu",
            "--no-sandbox",
            "--virtual-time-budget=15000",
            "--timeout=20000",  // wall-clock backstop: a stalled resource is not rescued by virtual time (issue #181)
        };
        if (extraArgs != null) args.AddRange(extraArgs);
        args.Add("--dump-dom");
        args.Add(url);
        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = bin,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                // CONSISTENCY(child-stream-encoding): see BlankDocCreator.
                StandardOutputEncoding = System.Text.Encoding.UTF8,
                StandardErrorEncoding = System.Text.Encoding.UTF8,
                UseShellExecute = false,
                CreateNoWindow = true,
            };
            foreach (var a in args) psi.ArgumentList.Add(a);
            using var p = Process.Start(psi);
            if (p == null) return null;
            var stdout = p.StandardOutput.ReadToEnd();
            if (!p.WaitForExit(timeoutMs)) { try { p.Kill(true); } catch { } return null; }
            return stdout;
        }
        catch { return null; }
    }

    /// <summary>True when a chrome-family browser (Chrome/Chromium/Edge) is available.</summary>
    public static bool HasChromeFamily() => FindChrome() != null;

    /// <summary>
    /// Screenshot <paramref name="htmlPath"/> in a chrome-family browser at an exact
    /// pixel window size, with a virtual-time budget (so async JS such as mermaid.js
    /// finishes before capture) and a HiDPI scale factor for crisp raster output.
    /// Returns true on a non-empty PNG at <paramref name="outPath"/>.
    /// </summary>
    public static bool CaptureChromeSized(string htmlPath, string outPath, int w, int h,
                                          int scale = 2, int timeoutMs = 60000)
    {
        var bin = FindChrome();
        if (bin == null) return false;
        outPath = Path.GetFullPath(outPath);
        var outDir = Path.GetDirectoryName(outPath);
        if (!string.IsNullOrEmpty(outDir)) Directory.CreateDirectory(outDir);
        var url = new Uri(Path.GetFullPath(htmlPath)).AbsoluteUri + "#screenshot";
        var args = new[]
        {
            "--headless=new",
            "--disable-gpu",
            "--no-sandbox",
            "--hide-scrollbars",
            $"--force-device-scale-factor={scale}",
            $"--window-size={w},{h}",
            "--virtual-time-budget=15000",
            "--timeout=20000",  // wall-clock backstop: a stalled resource is not rescued by virtual time (issue #181)
            "--default-background-color=00000000",
            $"--screenshot={outPath}",
            url,
        };
        var (ok, _) = RunBinary(bin, args);
        return ok && File.Exists(outPath) && new FileInfo(outPath).Length > 0;
    }

    public static PaginationResult? GetPaginationFromDom(string htmlPath, int timeoutMs = 60000)
    {
        var stdout = DumpDom(htmlPath, timeoutMs);
        if (stdout == null) return null;
        try
        {
            var m = System.Text.RegularExpressions.Regex.Match(stdout, @"<title>PAGES:(\d+)(?:\|MAP:([^<]*))?</title>");
            if (!m.Success || !int.TryParse(m.Groups[1].Value, out var n)) return null;
            var map = new Dictionary<string, int>();
            if (m.Groups[2].Success && m.Groups[2].Value.Length > 0)
            {
                foreach (var pair in m.Groups[2].Value.Split(','))
                {
                    var eq = pair.IndexOf('=');
                    if (eq > 0 && int.TryParse(pair[(eq + 1)..], out var pgNum))
                        map[pair[..eq]] = pgNum;
                }
            }
            return new PaginationResult(n, map);
        }
        catch { return null; }
    }

    public static int? GetPageCountFromDom(string htmlPath, int timeoutMs = 60000)
        => GetPaginationFromDom(htmlPath, timeoutMs)?.TotalPages;

    /// <summary>
    /// Pick a column count for a <paramref name="count"/>-item thumbnail contact
    /// sheet so the composed image is roughly square — used by `--grid auto`.
    /// Each cell is <paramref name="cellW"/>×<paramref name="cellH"/>; the grid
    /// image has aspect (rows·cellH)/(cols·cellW) = (count/cols²)·(cellH/cellW),
    /// so cols ≈ √(count·aspect) makes it ≈ 1. Portrait pages (aspect &gt; 1) get
    /// more columns, landscape slides (aspect &lt; 1) fewer. Clamped to [1, count].
    /// </summary>
    public static int AutoGridColumns(int count, double cellW, double cellH)
    {
        if (count <= 1) return 1;
        double aspect = cellH / Math.Max(1.0, cellW);
        int cols = (int)Math.Round(Math.Sqrt(count * aspect));
        return Math.Clamp(cols, 1, count);
    }

    public static Result Capture(string htmlPath, string outPath, int width = 1600, int height = 1200)
    {
        var url = new Uri(Path.GetFullPath(htmlPath)).AbsoluteUri + "#screenshot";
        outPath = Path.GetFullPath(outPath);
        var outDir = Path.GetDirectoryName(outPath);
        if (!string.IsNullOrEmpty(outDir)) Directory.CreateDirectory(outDir);

        // Cap to <= 1920px to stay within multi-image LLM limits.
        var (w, h) = CapDim(width, height, 1920);

        string? lastError = null;
        foreach (var (name, runner) in Backends())
        {
            var (ok, err) = runner(url, outPath, w, h);
            if (ok && File.Exists(outPath) && new FileInfo(outPath).Length > 0)
                return new Result(true, name, null);
            if (err != null) lastError = $"{name}: {err}";
        }
        return new Result(false, "", lastError ?? "no headless backend available");
    }

    /// <summary>
    /// Screenshot only the region covered by the given <c>data-path</c> targets
    /// (a single element, or several whose union bounding box is captured —
    /// e.g. the two corner cells of an xlsx range). Three passes, all plain
    /// chrome invocations: pass 1 injects a measure script that resolves the
    /// targets (un-hiding an inactive sheet tab's content) and reports the
    /// union rect + document extent through the document title; pass 2
    /// captures the WHOLE page at a viewport covering the target (layout is
    /// identical to the measure pass, so the coordinates hold — no transform
    /// tricks, which Chrome's paint culling and ~500px window clamp defeat);
    /// pass 3 loads that PNG in a bare wrapper page offset by the rect origin
    /// and captures at exactly the rect's size — a pure pixel crop.
    /// Chrome-family only (the injection needs a scriptable engine); returns
    /// a Result whose Error is "clip_target_not_found" when no target matched
    /// a rendered element.
    /// </summary>
    public static Result CaptureClipped(string htmlPath, string outPath, IReadOnlyList<string> dataPaths,
                                        int padPx = 0, int scale = 2)
    {
        if (FindChrome() == null)
            return new Result(false, "", "clip mode requires a Chrome-family browser (Chrome/Edge/Chromium)");
        var html = File.ReadAllText(htmlPath);
        var pathsJs = "[" + string.Join(",", dataPaths.Select(p =>
            "'" + p.Replace("\\", "\\\\").Replace("'", "\\'") + "'")) + "]";
        // Shared prelude: resolve targets and force-display hidden ancestors so
        // an inactive sheet's cells become measurable/visible.
        // A data-path can match several nodes (the pptx sidebar thumbnails
        // clone slide content), so pick the VISIBLE match with the largest
        // box; only when every match is hidden (an inactive xlsx sheet tab)
        // un-hide the first one's display:none ancestors and use it.
        var prelude =
            "var _clipPaths=" + pathsJs + ";" +
            "function _clipPick(p){var cands=document.querySelectorAll('[data-path=\"'+p+'\"]');" +
            "var best=null,bestA=0;cands.forEach(function(el){var r=el.getBoundingClientRect();" +
            "var a=r.width*r.height;if(a>bestA){bestA=a;best=el;}});" +
            "if(best)return best;if(cands.length===0)return null;" +
            "var el=cands[0];for(var an=el;an&&an!==document.documentElement;an=an.parentElement){" +
            "if(getComputedStyle(an).display==='none')an.style.display='block';}return el;}" +
            "function _clipEls(){var els=[];_clipPaths.forEach(function(p){" +
            "var el=_clipPick(p);if(el)els.push(el);});return els;}";

        // The rect is reported via console.log -> --enable-logging=stderr in
        // the SAME chrome invocation that captures the full page: headless
        // dump-dom uses a different, fixed viewport (~1024x513, --window-size
        // ignored) than screenshot mode, and any preview whose layout depends
        // on viewport height (the xlsx sheet stack) measures differently
        // there — a separate measure pass can never be trusted. One process,
        // one layout, one frame. The settle delay runs late in the virtual
        // time budget so font loads and the page's own layout JS (pptx shape
        // positioning) have finished; virtual time fast-forwards it, so the
        // wall-clock cost is nil. No requestAnimationFrame anywhere: rAF
        // fires unreliably under --virtual-time-budget.
        var measureScript =
            "<script>function _clipRun(){" +
            prelude +
            "var els=_clipEls();if(els.length===0){console.log('CLIPRECT:NOTFOUND');return;}" +
            "var x1=1e9,y1=1e9,x2=-1e9,y2=-1e9;els.forEach(function(el){var r=el.getBoundingClientRect();" +
            "var x=r.left+window.scrollX,y=r.top+window.scrollY;" +
            "if(x<x1)x1=x;if(y<y1)y1=y;if(x+r.width>x2)x2=x+r.width;if(y+r.height>y2)y2=y+r.height;});" +
            "console.log('CLIPRECT:'+Math.floor(x1)+','+Math.floor(y1)+','+Math.ceil(x2-x1)+','+Math.ceil(y2-y1)" +
            "+','+window.innerWidth+','+window.innerHeight);}" +
            // Un-hide EARLY, measure LATE: revealing an inactive sheet's
            // content triggers the page's own observers (resize/height
            // adjustments), so a rect taken in the same tick as the un-hide
            // reads the pre-adjustment layout while the paint shows the
            // post-adjustment one. The 11-second virtual gap lets every
            // observer settle before the coordinates are read.
            "function _clipArm(){setTimeout(function(){_clipEls();},800);setTimeout(_clipRun,12000);}" +
            "if(document.readyState==='complete')_clipArm();else window.addEventListener('load',_clipArm);</script>";

        string Inject(string doc, string script) =>
            doc.Contains("</body>", StringComparison.OrdinalIgnoreCase)
                ? doc.Replace("</body>", script + "</body>")
                : doc + script;

        var measurePath = htmlPath + ".clipmeasure.html";
        var fullPath = htmlPath + ".clipfull.png";
        var wrapPath = htmlPath + ".clipwrap.html";
        try
        {
            File.WriteAllText(measurePath, Inject(html, measureScript));
            // Combined measure + whole-page capture. When the target extends
            // past the current viewport, retry AT the enlarged viewport (the
            // enlargement can reflow viewport-dependent layout, so the rect
            // must come from the same run that painted the final PNG).
            // Headless chrome's JS/layout viewport is SHORTER than the
            // requested window (87px of window chrome on macOS; 0 on some
            // platforms), and the final --screenshot paint re-lays-out at the
            // FULL window height — so a rect measured by page JS lives in a
            // different vertical layout than the painted pixels whenever the
            // page's layout depends on viewport height (the xlsx sheet
            // stack). Self-calibrate: the measure run reports its own
            // innerHeight, giving delta = window - viewport; the measure runs
            // with the window ENLARGED by delta (JS layout = target height)
            // and the capture runs with the plain window (paint layout = the
            // same target height). No hardcoded platform constant.
            const int maxViewport = 4000;
            int vw = 800, vh = 600, x = 0, y = 0, w = 0, h = 0, deltaH = 0;
            for (var attempt = 0; ; attempt++)
            {
                var (ok, stderr) = RunChromeCapture(measurePath, fullPath, vw, vh + deltaH, scale);
                if (!ok)
                    return new Result(false, "chrome", "measure run failed");
                var rectM = System.Text.RegularExpressions.Regex.Match(
                    stderr ?? "", @"CLIPRECT:(-?\d+),(-?\d+),(\d+),(\d+),(\d+),(\d+)");
                if (!rectM.Success)
                {
                    if ((stderr ?? "").Contains("CLIPRECT:NOTFOUND"))
                        return new Result(false, "chrome", "clip_target_not_found");
                    return new Result(false, "chrome", "clip rect not reported (console logging unavailable?)");
                }
                x = Math.Max(0, int.Parse(rectM.Groups[1].Value) - padPx);
                y = Math.Max(0, int.Parse(rectM.Groups[2].Value) - padPx);
                w = int.Parse(rectM.Groups[3].Value) + 2 * padPx;
                h = int.Parse(rectM.Groups[4].Value) + 2 * padPx;
                var innerH = int.Parse(rectM.Groups[6].Value);
                if (w <= 0 || h <= 0) return new Result(false, "chrome", "clip target has an empty bounding box");
                var newDelta = (vh + deltaH) - innerH;
                var needW = Math.Max(vw, x + w);
                var needH = Math.Max(vh, y + h);
                if (needW > maxViewport || needH > maxViewport)
                    return new Result(false, "chrome",
                        $"clip target extends to ({needW},{needH}) CSS px — beyond the {maxViewport}px capture ceiling");
                var settled = needW == vw && needH == vh && innerH == vh;
                if (settled || attempt >= 4) { vw = needW; vh = needH; break; }
                vw = needW; vh = needH; deltaH = Math.Max(0, newDelta);
            }
            // Oversized targets drop the HiDPI factor to stay within the
            // multi-image LLM ceiling; device pixels only, CSS layout intact.
            if (Math.Max(w, h) * scale > 1920) scale = 1;
            // The whole-page capture: plain window (vw, vh) paints at exactly
            // the (vw, vh) layout the calibrated measure ran in.
            {
                var (okF, _) = RunChromeCapture(measurePath, fullPath, vw, vh, scale);
                if (!okF || !File.Exists(fullPath) || new FileInfo(fullPath).Length == 0)
                    return new Result(false, "chrome", "whole-page capture failed");
            }
            var fullW = vw;
            var fullH = vh;
            try
            {
                // Pass 3: pure pixel crop — show the full PNG offset by the
                // rect origin in a bare page and capture at the rect's size.
                // The PNG already carries the HiDPI scale, so the wrapper
                // renders it at its CSS size and captures at the same factor.
                var fullUri = new Uri(Path.GetFullPath(fullPath)).AbsoluteUri;
                File.WriteAllText(wrapPath,
                    "<!DOCTYPE html><html><head><title>clip</title><style>html,body{margin:0;padding:0;overflow:hidden}</style></head>" +
                    $"<body><img src=\"{fullUri}\" style=\"position:absolute;left:{-x}px;top:{-y}px;width:{fullW}px;height:{fullH}px\"></body></html>");
                var ok = CaptureChromeSized(wrapPath, outPath, w, h, scale);
                return ok
                    ? new Result(true, "chrome", null)
                    : new Result(false, "chrome", "clipped capture produced no image");
            }
            finally
            {
                try { File.Delete(fullPath); } catch { /* ignore */ }
                try { File.Delete(wrapPath); } catch { /* ignore */ }
            }
        }
        finally { try { File.Delete(measurePath); } catch { /* ignore */ } }
    }

    /// <summary>
    /// Parse a `--clip` target into the data-path list CaptureClipped consumes.
    /// Two forms: an xlsx range `/Sheet1/A1:C3` (also `Sheet1!A1:C3`) resolves
    /// to its two corner cell paths; anything else is a single element
    /// data-path used verbatim (`/slide[1]/shape[2]`, `/body/table[1]`, …).
    /// </summary>
    public static List<string> ResolveClipDataPaths(string clip)
    {
        var c = clip.Trim();
        // Sheet1!A1:C3 → /Sheet1/A1:C3
        var bang = c.IndexOf('!');
        if (bang > 0 && !c.StartsWith('/'))
            c = "/" + c[..bang] + "/" + c[(bang + 1)..];
        var m = System.Text.RegularExpressions.Regex.Match(
            c, @"^(/[^/]+)/([A-Za-z]{1,3}\d+):([A-Za-z]{1,3}\d+)$");
        if (m.Success)
        {
            return
            [
                $"{m.Groups[1].Value}/{m.Groups[2].Value.ToUpperInvariant()}",
                $"{m.Groups[1].Value}/{m.Groups[3].Value.ToUpperInvariant()}",
            ];
        }
        return [c];
    }

    private static IEnumerable<(string, Func<string, string, int, int, (bool, string?)>)> Backends()
    {
        yield return ("playwright", TryPlaywright);
        yield return ("chrome", TryChrome);
        yield return ("firefox", TryFirefox);
    }

    private static (int, int) CapDim(int w, int h, int limit)
    {
        var m = Math.Max(w, h);
        if (m <= limit) return (w, h);
        var s = (double)limit / m;
        return (Math.Max(1, (int)(w * s)), Math.Max(1, (int)(h * s)));
    }

    // ----- Playwright CLI -----------------------------------------------------------------

    private static (bool, string?) TryPlaywright(string url, string outPath, int w, int h)
    {
        var pw = WhichFirst("playwright");
        if (pw == null) return (false, null);
        var args = new[] { "screenshot", $"--viewport-size={w},{h}", "--full-page", url, outPath };
        return RunBinary(pw, args);
    }

    // ----- Chromium family ---------------------------------------------------------------

    private static (bool, string?) TryChrome(string url, string outPath, int w, int h)
    {
        var bin = FindChrome();
        if (bin == null) return (false, null);
        var args = new[]
        {
            "--headless=new",
            "--disable-gpu",
            "--no-sandbox",
            "--hide-scrollbars",
            $"--window-size={w},{h}",
            // Without these caps, new-headless --screenshot waits for the
            // page's external resources (CDN fonts / KaTeX css+js) to settle;
            // a slow or stalled request makes it sit until RunBinary's
            // 2-minute kill — per page, which reads as a hang on headless
            // Linux (issue #181). The virtual-time budget lets async JS
            // (KaTeX, mermaid) finish and bounds slow-but-completing loads;
            // --timeout is the wall-clock backstop for a STALLED request
            // (connection accepted, never answered), which virtual time
            // does NOT rescue — verified against a never-responding server.
            "--virtual-time-budget=15000",
            "--timeout=20000",
            $"--screenshot={outPath}",
            url,
        };
        return RunBinary(bin, args);
    }

    private static string? FindChrome()
    {
        string[] names = ["google-chrome", "google-chrome-stable", "chromium", "chromium-browser",
                          "chrome", "microsoft-edge", "microsoft-edge-stable", "msedge"];
        var pathHit = WhichFirst(names);
        if (pathHit != null) return pathHit;

        var abs = new List<string>();
        if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
        {
            abs.AddRange(new[]
            {
                "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
                "/Applications/Chromium.app/Contents/MacOS/Chromium",
                "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
            });
        }
        else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
        {
            abs.AddRange(new[]
            {
                "/usr/bin/google-chrome", "/usr/bin/chromium", "/usr/bin/chromium-browser",
                "/snap/bin/chromium", "/snap/bin/google-chrome",
                "/usr/bin/microsoft-edge", "/usr/bin/microsoft-edge-stable",
            });
        }
        else if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            string[] roots = [
                Environment.GetEnvironmentVariable("PROGRAMFILES") ?? @"C:\Program Files",
                Environment.GetEnvironmentVariable("PROGRAMFILES(X86)") ?? @"C:\Program Files (x86)",
                Environment.GetEnvironmentVariable("LOCALAPPDATA") ?? "",
            ];
            string[] suffixes = [
                @"Google\Chrome\Application\chrome.exe",
                @"Chromium\Application\chrome.exe",
                @"Microsoft\Edge\Application\msedge.exe",
            ];
            foreach (var r in roots)
                if (!string.IsNullOrEmpty(r))
                    foreach (var s in suffixes) abs.Add(Path.Combine(r, s));
        }
        return abs.FirstOrDefault(File.Exists);
    }

    // ----- Firefox -----------------------------------------------------------------------

    private static (bool, string?) TryFirefox(string url, string outPath, int w, int h)
    {
        var bin = FindFirefox();
        if (bin == null) return (false, null);
        // Firefox: `--headless --screenshot=<out> --window-size=W,H <URL>`.
        // Note: no `=new` headless variant; --force-device-scale-factor not supported.
        var args = new[] { "--headless", $"--screenshot={outPath}", $"--window-size={w},{h}", url };
        return RunBinary(bin, args);
    }

    private static string? FindFirefox()
    {
        var pathHit = WhichFirst("firefox", "firefox-esr");
        if (pathHit != null) return pathHit;

        var abs = new List<string>();
        if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
        {
            abs.AddRange(new[]
            {
                "/Applications/Firefox.app/Contents/MacOS/firefox",
                "/Applications/Firefox Developer Edition.app/Contents/MacOS/firefox",
                "/Applications/Firefox Nightly.app/Contents/MacOS/firefox",
            });
        }
        else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
        {
            abs.AddRange(new[] { "/usr/bin/firefox", "/usr/bin/firefox-esr", "/snap/bin/firefox" });
        }
        else if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            foreach (var r in new[]
            {
                Environment.GetEnvironmentVariable("PROGRAMFILES") ?? @"C:\Program Files",
                Environment.GetEnvironmentVariable("PROGRAMFILES(X86)") ?? @"C:\Program Files (x86)",
            })
                if (!string.IsNullOrEmpty(r)) abs.Add(Path.Combine(r, @"Mozilla Firefox\firefox.exe"));
        }
        return abs.FirstOrDefault(File.Exists);
    }

    // ----- Helpers -----------------------------------------------------------------------

    /// <summary>Find the first of <paramref name="names"/> on PATH (honouring
    /// Windows PATHEXT). Shared executable lookup — same mechanism used to detect
    /// playwright/chrome/firefox, so callers like the mermaid renderer detect mmdc
    /// identically.</summary>
    public static string? Which(params string[] names) => WhichFirst(names);

    private static string? WhichFirst(params string[] names)
    {
        var pathSep = RuntimeInformation.IsOSPlatform(OSPlatform.Windows) ? ';' : ':';
        var pathExt = RuntimeInformation.IsOSPlatform(OSPlatform.Windows)
            ? (Environment.GetEnvironmentVariable("PATHEXT") ?? ".COM;.EXE;.BAT;.CMD").Split(';')
            : new[] { "" };
        var paths = (Environment.GetEnvironmentVariable("PATH") ?? "").Split(pathSep);

        foreach (var name in names)
        {
            foreach (var dir in paths)
            {
                if (string.IsNullOrEmpty(dir)) continue;
                foreach (var ext in pathExt)
                {
                    var candidate = Path.Combine(dir, name + ext);
                    if (File.Exists(candidate)) return candidate;
                }
            }
        }
        return null;
    }

    /// <summary>
    /// One chrome run that paints a screenshot AND returns stderr, where
    /// --enable-logging=stderr surfaces the page's console.log lines — the
    /// channel CaptureClipped uses to read the measured rect out of the very
    /// same invocation that produced the pixels. Both streams are drained
    /// asynchronously before the bounded wait (undrained pipes deadlock at
    /// ~64KB).
    /// </summary>
    private static (bool Ok, string? Stderr) RunChromeCapture(string htmlPath, string outPath, int w, int h, int scale)
    {
        var bin = FindChrome();
        if (bin == null) return (false, null);
        outPath = Path.GetFullPath(outPath);
        var url = new Uri(Path.GetFullPath(htmlPath)).AbsoluteUri + "#screenshot";
        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = bin,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                // CONSISTENCY(child-stream-encoding): see BlankDocCreator.
                StandardOutputEncoding = System.Text.Encoding.UTF8,
                StandardErrorEncoding = System.Text.Encoding.UTF8,
                UseShellExecute = false,
                CreateNoWindow = true,
            };
            foreach (var a in new[]
            {
                "--headless=new",
                "--disable-gpu",
                "--no-sandbox",
                "--hide-scrollbars",
                "--enable-logging=stderr",
                "--v=0",
                $"--force-device-scale-factor={scale}",
                $"--window-size={w},{h}",
                "--virtual-time-budget=15000",
                "--timeout=20000",
                $"--screenshot={outPath}",
                url,
            }) psi.ArgumentList.Add(a);
            using var p = Process.Start(psi);
            if (p == null) return (false, null);
            var outTask = p.StandardOutput.ReadToEndAsync();
            var errTask = p.StandardError.ReadToEndAsync();
            if (!p.WaitForExit(120_000))
            {
                try { p.Kill(true); } catch { /* ignore */ }
                return (false, null);
            }
            return (p.ExitCode == 0, errTask.GetAwaiter().GetResult());
        }
        catch { return (false, null); }
    }

    private static (bool, string?) RunBinary(string bin, string[] args)
    {
        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = bin,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                // CONSISTENCY(child-stream-encoding): see BlankDocCreator.
                StandardOutputEncoding = System.Text.Encoding.UTF8,
                StandardErrorEncoding = System.Text.Encoding.UTF8,
                UseShellExecute = false,
                CreateNoWindow = true,
            };
            foreach (var a in args) psi.ArgumentList.Add(a);
            using var p = Process.Start(psi);
            if (p == null) return (false, "process did not start");
            if (!p.WaitForExit(120_000))
            {
                try { p.Kill(true); } catch { /* ignore */ }
                return (false, "timeout after 120s");
            }
            if (p.ExitCode != 0)
            {
                var stderr = p.StandardError.ReadToEnd();
                var lastLine = stderr.Trim().Split('\n').LastOrDefault() ?? $"exit {p.ExitCode}";
                return (false, lastLine);
            }
            return (true, null);
        }
        catch (Exception e)
        {
            return (false, e.Message);
        }
    }
}
