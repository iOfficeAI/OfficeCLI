// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Threading;

namespace OfficeCli.Core;

/// <summary>
/// OS-native PNG rendering for .pptx on Windows: drives the installed
/// presentation application through its automation interface to export each
/// requested slide straight to a PNG, then stitches a multi-slide range
/// vertically. Failures are surfaced to the caller; auto mode catches them and
/// falls back to the HTML screenshot path. The COM/IDispatch plumbing and the
/// PNG stitch are shared with <see cref="WordPdfBackend"/>.
/// </summary>
[SupportedOSPlatform("windows")]
internal static class PowerPointPngBackend
{
    static readonly Guid G_App = new("91493441-5A91-11CF-8700-00AA0060263B");

    /// Render slides [startSlide..endSlide] (1-based, inclusive) to a single PNG
    /// at width×height pixels. A range is stitched top-to-bottom. Runs on a
    /// dedicated STA thread and preserves failures from that thread.
    public static byte[]? Render(string pptx, int startSlide, int endSlide, int width, int height, int timeoutMs = 60000)
    {
        // Keep within the multi-image LLM ceiling, same 1920 long-edge cap as the HTML path.
        var m = Math.Max(width, height);
        if (m > 1920) { var s = 1920.0 / m; width = Math.Max(1, (int)(width * s)); height = Math.Max(1, (int)(height * s)); }

        byte[]? result = null;
        Exception? error = null;
        var th = new Thread(() =>
        {
            var tmp = new List<string>();
            try { result = RenderCore(pptx, startSlide, endSlide, width, height, timeoutMs, tmp); }
            catch (Exception e) { error = e; }
            finally { foreach (var f in tmp) { try { File.Delete(f); } catch { /* ignore */ } } }
        });
        th.SetApartmentState(ApartmentState.STA);
        th.IsBackground = true;
        th.Start();
        if (!th.Join(timeoutMs + 30000))
            throw new TimeoutException("PowerPoint native rendering timed out.");
        if (error != null) ExceptionDispatchInfo.Capture(error).Throw();
        return result;
    }

    /// Render slides [startSlide..endSlide] (1-based; endSlide <= 0 means "to the
    /// last slide") into an N-column thumbnail grid. Each slide is exported at
    /// cellW×cellH and tiled with the given gap/padding (pixels) on a white
    /// background. Cells are scaled down if the composed image would exceed the
    /// 1920 long-edge ceiling. Failures are surfaced to the caller.
    public static byte[]? RenderGrid(string pptx, int startSlide, int endSlide, int cellW, int cellH, int cols, int gap, int pad, int timeoutMs = 120000)
    {
        byte[]? result = null;
        Exception? error = null;
        var th = new Thread(() =>
        {
            var tmp = new List<string>();
            try { result = RenderGridCore(pptx, startSlide, endSlide, cellW, cellH, cols, gap, pad, timeoutMs, tmp); }
            catch (Exception e) { error = e; }
            finally { foreach (var f in tmp) { try { File.Delete(f); } catch { /* ignore */ } } }
        });
        th.SetApartmentState(ApartmentState.STA);
        th.IsBackground = true;
        th.Start();
        if (!th.Join(timeoutMs + 30000))
            throw new TimeoutException("PowerPoint native grid rendering timed out.");
        if (error != null) ExceptionDispatchInfo.Capture(error).Throw();
        return result;
    }

    static byte[]? RenderGridCore(string pptx, int startSlide, int endSlide, int cellW, int cellH, int cols, int gap, int pad, int timeoutMs, List<string> tmp)
    {
        if (cols < 1) cols = 1;
        var clsid = G_App; var iid = WordPdfBackend.G_IDispatch;
        WordPdfBackend.CoCreateInstance(ref clsid, IntPtr.Zero, 4, ref iid, out var app);
        try
        {
            var name = (string?)WordPdfBackend.DispGet(app, "Name") ?? "";
            if (!name.Contains("PowerPoint", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("app_not_authentic: " + name);
            try { WordPdfBackend.DispSet(app, "DisplayAlerts", 1); } catch { /* alerts-none; ignore if unsettable */ }

            var presentations = (IntPtr)WordPdfBackend.DispGet(app, "Presentations")!;
            try
            {
                var pres = (IntPtr)WordPdfBackend.DispMethod(presentations, "Open", Path.GetFullPath(pptx), -1, 0, 0)!;
                try
                {
                    var slides = (IntPtr)WordPdfBackend.DispGet(pres, "Slides")!;
                    try
                    {
                        var count = WaitForCount(slides, timeoutMs);
                        int s = Math.Max(1, startSlide);
                        int e = Math.Min(count, endSlide <= 0 ? count : endSlide);
                        if (e < s) return null;
                        int n = e - s + 1;
                        int rows = (n + cols - 1) / cols;

                        // Scale cells down so the composed grid stays within the 1920 long-edge ceiling.
                        int totalW = pad * 2 + cols * cellW + (cols - 1) * gap;
                        int totalH = pad * 2 + rows * cellH + (rows - 1) * gap;
                        int m = Math.Max(totalW, totalH);
                        if (m > 1920) { var sc = 1920.0 / m; cellW = Math.Max(1, (int)(cellW * sc)); cellH = Math.Max(1, (int)(cellH * sc)); }

                        var pngs = new List<byte[]>();
                        for (int i = s; i <= e; i++)
                        {
                            var slide = (IntPtr)WordPdfBackend.DispMethod(slides, "Item", i)!;
                            try
                            {
                                var outFile = Path.Combine(Path.GetTempPath(), $"_pg_{Guid.NewGuid():N}.png");
                                tmp.Add(outFile);
                                WordPdfBackend.DispMethod(slide, "Export", outFile, "PNG", cellW, cellH);
                                pngs.Add(File.ReadAllBytes(outFile));
                            }
                            finally { Marshal.Release(slide); }
                        }
                        return pngs.Count == 0 ? null : WordPdfBackend.StitchGrid(pngs, cols, gap, pad);
                    }
                    finally { Marshal.Release(slides); }
                }
                finally { try { WordPdfBackend.DispMethod(pres, "Close"); } catch { } Marshal.Release(pres); }
            }
            finally { Marshal.Release(presentations); }
        }
        finally { try { WordPdfBackend.DispMethod(app, "Quit"); } catch { } Marshal.Release(app); }
    }

    static byte[]? RenderCore(string pptx, int startSlide, int endSlide, int width, int height, int timeoutMs, List<string> tmp)
    {
        var clsid = G_App; var iid = WordPdfBackend.G_IDispatch;
        WordPdfBackend.CoCreateInstance(ref clsid, IntPtr.Zero, 4, ref iid, out var app);
        try
        {
            var name = (string?)WordPdfBackend.DispGet(app, "Name") ?? "";
            if (!name.Contains("PowerPoint", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("app_not_authentic: " + name);
            try { WordPdfBackend.DispSet(app, "DisplayAlerts", 1); } catch { /* alerts-none; ignore if unsettable */ }

            var presentations = (IntPtr)WordPdfBackend.DispGet(app, "Presentations")!;
            try
            {
                // Open(FileName, ReadOnly=-1, Untitled=0, WithWindow=0): read-only, no window.
                var pres = (IntPtr)WordPdfBackend.DispMethod(presentations, "Open", Path.GetFullPath(pptx), -1, 0, 0)!;
                try
                {
                    var slides = (IntPtr)WordPdfBackend.DispGet(pres, "Slides")!;
                    try
                    {
                        var count = WaitForCount(slides, timeoutMs);
                        var pngs = new List<byte[]>();
                        for (int n = startSlide; n <= endSlide && n <= count; n++)
                        {
                            if (n < 1) continue;
                            var slide = (IntPtr)WordPdfBackend.DispMethod(slides, "Item", n)!;
                            try
                            {
                                var outFile = Path.Combine(Path.GetTempPath(), $"_p_{Guid.NewGuid():N}.png");
                                tmp.Add(outFile);
                                // Export(FileName, FilterName, ScaleWidth, ScaleHeight).
                                WordPdfBackend.DispMethod(slide, "Export", outFile, "PNG", width, height);
                                pngs.Add(File.ReadAllBytes(outFile));
                            }
                            finally { Marshal.Release(slide); }
                        }
                        return pngs.Count == 0 ? null : WordPdfBackend.Stitch(pngs);
                    }
                    finally { Marshal.Release(slides); }
                }
                finally { try { WordPdfBackend.DispMethod(pres, "Close"); } catch { } Marshal.Release(pres); }
            }
            finally { Marshal.Release(presentations); }
        }
        finally { try { WordPdfBackend.DispMethod(app, "Quit"); } catch { } Marshal.Release(app); }
    }

    /// Poll the slide count until it is readable (large decks finish loading
    /// asynchronously), then return it.
    static int WaitForCount(IntPtr slides, int timeoutMs)
    {
        var deadline = Environment.TickCount64 + Math.Min(timeoutMs, 30000);
        while (Environment.TickCount64 < deadline)
        {
            try { return (int)WordPdfBackend.DispGet(slides, "Count")!; }
            catch { Thread.Sleep(200); }
        }
        return (int)WordPdfBackend.DispGet(slides, "Count")!;
    }
}
