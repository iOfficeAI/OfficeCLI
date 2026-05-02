// Copyright 2025 OfficeCli (officecli.ai)
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

    public static Result Capture(string htmlPath, string outPath, int width = 1600, int height = 1200)
    {
        var url = new Uri(Path.GetFullPath(htmlPath)).AbsoluteUri;
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

    private static (bool, string?) RunBinary(string bin, string[] args)
    {
        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = bin,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
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
