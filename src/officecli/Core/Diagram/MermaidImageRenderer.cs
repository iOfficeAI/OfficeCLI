// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeCli.Core.Diagram;

/// <summary>
/// Optional high-fidelity path for the <c>diagram</c> element: render the mermaid
/// source with the <b>real mermaid.js</b> to a PNG — covering every mermaid diagram
/// type (gantt / pie / class / state / er / git / mindmap / …) that the native
/// shape synthesizer does not, at full fidelity.
///
/// <para>Backend cascade, best tool first:
/// <list type="number">
///   <item><b>mmdc</b> (the official mermaid-cli), if installed — purpose-built,
///   one call to a tight PNG.</item>
///   <item><b>Chrome-family</b> browser the user already has (via
///   <see cref="HtmlScreenshot"/>): render mermaid.js in a page and screenshot it.
///   Only mermaid.min.js (~3.5 MB) is fetched to a local cache on first use
///   (mirror → CDN); if that fails the page loads mermaid from the CDN live.</item>
///   <item>otherwise the caller falls back to the native synthesizer
///   (<see cref="DiagramCompiler"/>) — zero dependencies, fully editable shapes.</item>
/// </list>
/// PNG (not SVG) throughout: Office cannot draw mermaid's <c>&lt;foreignObject&gt;</c>
/// HTML labels, so a raster that bakes in the browser's own rendering is required.</para>
/// </summary>
/// <summary>
/// The mermaid source is syntactically invalid (as opposed to an infrastructure
/// failure — missing browser, mermaid.js download, screenshot crash). Derives from
/// <see cref="ArgumentException"/> so the CLI surfaces it as a bad-input error
/// (<c>success:false</c> with the message) rather than a process failure, and so the
/// diagram Add path does NOT fall back to the native synthesizer — every backend
/// rejects the same broken source, and the point is to feed the parse error (with its
/// line number) back so the caller can fix it.
/// </summary>
public sealed class MermaidSyntaxException : ArgumentException
{
    public MermaidSyntaxException(string message) : base(message) { }
}

public static class MermaidImageRenderer
{
    // Pin a major version so cache + mirror + CDN agree and rendering is stable.
    private const string MermaidVersion = "11";
    // Own mirror first (offline-first, no third-party dependency at steady state),
    // then the public CDN as a fallback.
    private const string MirrorUrl =
        "https://d.officecli.ai/assets/mermaid-" + MermaidVersion + ".min.js";
    private const string CdnUrl =
        "https://cdn.jsdelivr.net/npm/mermaid@" + MermaidVersion + "/dist/mermaid.min.js";

    // ESM builds, used only when a style option (theme / layout / look) is set.
    // The UMD global (mermaid.min.js above) does not render a source carrying a
    // YAML `---` frontmatter block (run() marks it processed but emits no svg —
    // verified headless), and ELK ships as an ESM-only layout package that must
    // be registered via mermaid.registerLayoutLoaders. So a styled diagram is
    // rendered by importing the ESM mermaid + (for elk) the ESM elk loader.
    // These are NOT locally cached: the ESM build pulls dozens of relative
    // ./chunks/*.mjs and cross-package deps (dayjs/khroma/dompurify), so a
    // single-file cache cannot satisfy them — the page imports them live from
    // the CDN. Offline styling would need a pre-bundled single-file ESM hosted
    // on the mirror (future infra task); the plain UMD path stays offline.
    private const string ElkVersion = "0.1";
    private const string MermaidEsmUrl =
        "https://cdn.jsdelivr.net/npm/mermaid@" + MermaidVersion + "/dist/mermaid.esm.min.mjs";
    private const string ElkEsmUrl =
        "https://cdn.jsdelivr.net/npm/@mermaid-js/layout-elk@" + ElkVersion
        + "/dist/mermaid-layout-elk.esm.min.mjs";

    // Accepted style-option values. Kept in sync with schemas/help/*/diagram.json.
    private static readonly HashSet<string> Themes =
        new(StringComparer.OrdinalIgnoreCase) { "default", "dark", "neutral", "forest", "base" };
    private static readonly HashSet<string> Layouts =
        new(StringComparer.OrdinalIgnoreCase) { "dagre", "elk" };
    private static readonly HashSet<string> Looks =
        new(StringComparer.OrdinalIgnoreCase) { "classic", "handdrawn" };

    /// <summary>Sentinel prefix stamped into the rendered picture's alt-text so the
    /// mermaid source travels inside the document and the diagram stays regenerable.</summary>
    public const string SourceTag = "mermaid:";

    /// <summary>
    /// Readability floor for the adaptive one-page default: the smallest fit scale
    /// at which a diagram shrunk to fit a single page/slide stays legible. Mermaid's
    /// default node text is ~16px (=12pt at 96 DPI); shrinking it by this factor
    /// lands near 6.5pt — about the floor for readable print. When one-page fit
    /// would drop below this, the caller grows the canvas (poster) instead so a long
    /// flowchart does not become an unreadable sliver.
    /// </summary>
    public const double ReadableFitFloor = 0.55;

    /// <summary>
    /// Adaptive-default decision: would fitting a diagram of natural pixel size
    /// (<paramref name="wPx"/>×<paramref name="hPx"/>) into a one-page box
    /// (<paramref name="boxWcm"/>×<paramref name="boxHcm"/>) shrink it below the
    /// readability floor? The raster px are read as 96-DPI CSS pixels. True → the
    /// caller should grow the canvas rather than fit. False (incl. degenerate
    /// inputs) → fit to one page is fine.
    /// </summary>
    public static bool ExceedsOnePageReadably(double wPx, double hPx, double boxWcm, double boxHcm)
    {
        if (wPx <= 0 || hPx <= 0 || boxWcm <= 0 || boxHcm <= 0) return false;
        double natWcm = wPx / 96.0 * 2.54, natHcm = hPx / 96.0 * 2.54;
        double s = Math.Min(boxWcm / natWcm, boxHcm / natHcm);
        return s < ReadableFitFloor;
    }

    /// <summary>
    /// Bake the requested style options into the mermaid source as a leading
    /// <c>--- config: … ---</c> frontmatter block, so they render AND round-trip
    /// (the composed source is what gets stamped into alt-text). Returns the
    /// source unchanged when no option is set. Rejects unknown values with a
    /// message listing the valid ones. When the source already carries its own
    /// frontmatter or an <c>%%{init}%%</c> directive, the source wins and the
    /// options are ignored (caller may warn) — merging into an existing block is
    /// out of scope and would risk producing a malformed document.
    /// </summary>
    public static string ComposeSource(string mermaid, string? theme, string? layout, string? look)
    {
        theme = string.IsNullOrWhiteSpace(theme) ? null : theme.Trim();
        layout = string.IsNullOrWhiteSpace(layout) ? null : layout.Trim();
        look = string.IsNullOrWhiteSpace(look) ? null : look.Trim();
        if (theme == null && layout == null && look == null) return mermaid;

        if (theme != null && !Themes.Contains(theme))
            throw new ArgumentException($"unknown diagram theme '{theme}'. Valid: {string.Join(", ", Themes)}.");
        if (layout != null && !Layouts.Contains(layout))
            throw new ArgumentException($"unknown diagram layout '{layout}'. Valid: {string.Join(", ", Layouts)}.");
        if (look != null && !Looks.Contains(look))
            throw new ArgumentException($"unknown diagram look '{look}'. Valid: classic, handDrawn.");

        var lead = mermaid.TrimStart();
        if (lead.StartsWith("---", StringComparison.Ordinal) || lead.StartsWith("%%{", StringComparison.Ordinal))
            return mermaid; // source already declares config — do not double-inject

        var sb = new StringBuilder("---\nconfig:\n");
        if (theme != null) sb.Append("  theme: ").Append(theme.ToLowerInvariant()).Append('\n');
        if (layout != null) sb.Append("  layout: ").Append(layout.ToLowerInvariant()).Append('\n');
        // look's canonical mermaid spelling is camelCase handDrawn; normalize.
        if (look != null)
            sb.Append("  look: ")
              .Append(look.Equals("handdrawn", StringComparison.OrdinalIgnoreCase) ? "handDrawn" : "classic")
              .Append('\n');
        sb.Append("---\n").Append(mermaid);
        return sb.ToString();
    }

    /// <summary>True when the (already composed) source carries a style frontmatter,
    /// so it must be rendered via the ESM path rather than the UMD global.</summary>
    private static bool SourceNeedsEsm(string source)
    {
        var lead = source.TrimStart();
        if (!lead.StartsWith("---", StringComparison.Ordinal)) return false;
        var end = lead.IndexOf("\n---", 3, StringComparison.Ordinal);
        if (end < 0) return false;
        var fm = lead.Substring(0, end);
        return fm.Contains("theme:", StringComparison.OrdinalIgnoreCase)
            || fm.Contains("layout:", StringComparison.OrdinalIgnoreCase)
            || fm.Contains("look:", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>True when the composed source selects the ELK layout, which needs
    /// the ESM elk loader registered in the page.</summary>
    private static bool SourceNeedsElk(string source)
    {
        var lead = source.TrimStart();
        if (!lead.StartsWith("---", StringComparison.Ordinal)) return false;
        var end = lead.IndexOf("\n---", 3, StringComparison.Ordinal);
        if (end < 0) return false;
        return Regex.IsMatch(lead.Substring(0, end), @"layout:\s*elk\b", RegexOptions.IgnoreCase);
    }

    private static string CacheDir => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".officecli", "cache");
    private static string CachedJsPath => Path.Combine(CacheDir, $"mermaid-{MermaidVersion}.min.js");

    /// <summary>True when any image backend is available: mmdc, or a chrome-family browser.</summary>
    public static bool IsAvailable() => TryLocateMmdc(out _) || HtmlScreenshot.HasChromeFamily();

    /// <summary>
    /// Daily-refresh hook, called from <see cref="UpdateChecker"/>'s once-per-24h
    /// background process (already talking to the mirror). Revalidates an <b>already
    /// cached</b> mermaid.js against the mirror with a conditional request and updates
    /// it if the server's copy changed. Never pre-downloads (first-use owns that),
    /// never blocks, never throws. Only the chrome backend uses this cache; mmdc ships
    /// its own mermaid.
    /// </summary>
    public static void RefreshCacheIfPresent()
    {
        try
        {
            if (!File.Exists(CachedJsPath)) return; // refresh only what the user actually uses
            using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(20) };
            using var req = new HttpRequestMessage(HttpMethod.Get, MirrorUrl);
            req.Headers.IfModifiedSince = new DateTimeOffset(File.GetLastWriteTimeUtc(CachedJsPath));
            using var resp = http.SendAsync(req).GetAwaiter().GetResult();
            if (resp.StatusCode == System.Net.HttpStatusCode.NotModified) return; // unchanged → keep cache
            if (!resp.IsSuccessStatusCode) return;                                 // mirror hiccup → keep cache
            var bytes = resp.Content.ReadAsByteArrayAsync().GetAwaiter().GetResult();
            if (bytes.Length > 500_000)
                File.WriteAllBytes(CachedJsPath, bytes);
        }
        catch { /* best effort — the existing cache stays usable */ }
    }

    /// <summary>
    /// Render <paramref name="mermaid"/> to a temporary PNG file and return its path
    /// (caller owns + deletes it). Tries mmdc first (purpose-built, one call), then a
    /// chrome-family browser; degrades between them so a broken mmdc still yields a
    /// render. Throws <see cref="InvalidOperationException"/> only when no backend
    /// works (message carries the underlying tool's error).
    /// </summary>
    public static string RenderToPngFile(string mermaid, string? background = null)
    {
        Exception? failure = null;
        if (TryLocateMmdc(out var mmdc))
        {
            try { return RenderViaMmdc(mermaid, mmdc, background); }
            catch (MermaidSyntaxException) { throw; } // bad input — the browser would reject it too
            catch (Exception e) { failure = e; }
        }
        if (HtmlScreenshot.HasChromeFamily())
        {
            try { return RenderViaChrome(mermaid, background); }
            catch (MermaidSyntaxException) { throw; } // surface the parse error, don't mask it
            catch (Exception e) { failure ??= e; }
        }
        throw failure ?? new InvalidOperationException(
            "render=image needs mermaid-cli (mmdc) or a headless browser (Chrome/Chromium/Edge). "
            + "Install one, or use render=native for the built-in synthesizer.");
    }

    // ----- mmdc (official mermaid-cli) --------------------------------------------------

    private static string? _mmdcExe;
    private static bool _mmdcProbed;

    /// <summary>Locate mmdc: OFFICECLI_MMDC (explicit path) wins, else <c>mmdc</c> on PATH.</summary>
    private static bool TryLocateMmdc(out string exe)
    {
        if (!_mmdcProbed)
        {
            _mmdcProbed = true;
            _mmdcExe = ProbeMmdc();
        }
        exe = _mmdcExe ?? "";
        return _mmdcExe != null;
    }

    /// <summary>Heuristic: does mmdc's stderr/stdout describe a source syntax problem
    /// (vs a crash / environment fault)? mmdc surfaces mermaid's own parser text.</summary>
    private static bool LooksLikeSyntaxError(string msg)
    {
        if (string.IsNullOrEmpty(msg)) return false;
        return msg.Contains("Parse error", StringComparison.OrdinalIgnoreCase)
            || msg.Contains("Lexical error", StringComparison.OrdinalIgnoreCase)
            || msg.Contains("No diagram type detected", StringComparison.OrdinalIgnoreCase)
            || msg.Contains("UnknownDiagramError", StringComparison.OrdinalIgnoreCase)
            || msg.Contains("Expecting ", StringComparison.Ordinal);
    }

    private static string? ProbeMmdc()
    {
        // OFFICECLI_MMDC (explicit path) wins; otherwise find `mmdc` on PATH via the
        // same shared lookup used for chrome/playwright (WhichFirst handles PATHEXT,
        // so "mmdc" resolves mmdc.cmd on Windows).
        var env = Environment.GetEnvironmentVariable("OFFICECLI_MMDC");
        if (!string.IsNullOrWhiteSpace(env) && File.Exists(env)) return env;
        return HtmlScreenshot.Which("mmdc");
    }

    /// <summary>Validate a user background value for the mmdc <c>-b</c> flag / page
    /// style: <c>transparent</c>, a #hex, or a plain CSS color word. Rejects
    /// anything with whitespace/quotes so it can't break out of the flag or the
    /// inline style. Null/empty → transparent.</summary>
    private static string SafeBackground(string? bg)
    {
        if (string.IsNullOrWhiteSpace(bg)) return "transparent";
        bg = bg.Trim();
        return Regex.IsMatch(bg, @"^(transparent|#[0-9a-fA-F]{3,8}|[a-zA-Z]{1,20})$")
            ? bg
            : "transparent";
    }

    private static string RenderViaMmdc(string mermaid, string exe, string? background)
    {
        var inPath = Path.Combine(Path.GetTempPath(), $"ocli_mmd_{Guid.NewGuid():N}.mmd");
        var outPath = Path.ChangeExtension(inPath, ".png");
        File.WriteAllText(inPath, mermaid);
        try
        {
            var psi = new ProcessStartInfo(exe)
            {
                RedirectStandardError = true,
                RedirectStandardOutput = true,
                // CONSISTENCY(child-stream-encoding): see BlankDocCreator.
                StandardOutputEncoding = System.Text.Encoding.UTF8,
                StandardErrorEncoding = System.Text.Encoding.UTF8,
                UseShellExecute = false,
                CreateNoWindow = true,
            };
            psi.ArgumentList.Add("-i"); psi.ArgumentList.Add(inPath);
            psi.ArgumentList.Add("-o"); psi.ArgumentList.Add(outPath);
            psi.ArgumentList.Add("-b"); psi.ArgumentList.Add(SafeBackground(background));
            psi.ArgumentList.Add("-s"); psi.ArgumentList.Add("2"); // HiDPI for crisp raster
            var pcfg = Environment.GetEnvironmentVariable("OFFICECLI_MMDC_PUPPETEER");
            if (!string.IsNullOrWhiteSpace(pcfg) && File.Exists(pcfg))
            {
                psi.ArgumentList.Add("-p"); psi.ArgumentList.Add(pcfg);
            }

            using var p = Process.Start(psi)
                ?? throw new InvalidOperationException("failed to start mmdc.");
            // Async-drain both streams: the serial stderr-then-stdout reads
            // interlocked when mmdc filled the stdout pipe first (bounded by
            // the 120s kill below, but a wasted two minutes per diagram).
            var errTask = p.StandardError.ReadToEndAsync();
            var outTask = p.StandardOutput.ReadToEndAsync();
            if (!p.WaitForExit(120_000))
            {
                try { p.Kill(true); } catch { /* best effort */ }
                throw new InvalidOperationException("mmdc timed out after 120s.");
            }
            if (p.ExitCode != 0 || !File.Exists(outPath))
            {
                var msg = $"{errTask.Result}{outTask.Result}".Trim();
                // A parse/unknown-type failure is bad input, not a broken mmdc; class
                // it as syntax so the Add path surfaces it (and does not fall back).
                if (LooksLikeSyntaxError(msg))
                    throw new MermaidSyntaxException(
                        $"mermaid syntax error: {msg} "
                        + "(fix the mermaid source, or use render=native for the built-in subset).");
                throw new InvalidOperationException($"mmdc failed (exit {p.ExitCode}). {msg}".Trim());
            }
            return outPath;
        }
        finally { try { File.Delete(inPath); } catch { /* best effort */ } }
    }

    // ----- chrome-family browser (mermaid.js in a page → sized screenshot) --------------

    /// <summary>Two chrome passes: dump the DOM to read the diagram's viewBox, then
    /// screenshot at exactly that size (HiDPI). PNG bakes in the browser's rendering
    /// so mermaid's foreignObject labels — invisible to Office as SVG — appear.</summary>
    private static string RenderViaChrome(string mermaid, string? background)
    {
        // A styled source (theme/layout/look frontmatter) is rendered by the ESM
        // build (the UMD global does not render frontmatter, and elk is ESM-only);
        // a plain source keeps the offline-cached UMD path unchanged.
        var html = SourceNeedsEsm(mermaid)
            ? BuildHtmlEsm(mermaid, SourceNeedsElk(mermaid), SafeBackground(background))
            : BuildHtml(mermaid, ResolveMermaidJsRef(), SafeBackground(background));
        var htmlPath = Path.Combine(Path.GetTempPath(), $"ocli_mmd_{Guid.NewGuid():N}.html");
        File.WriteAllText(htmlPath, html);
        try
        {
            var dom = HtmlScreenshot.DumpDom(htmlPath)
                ?? throw new InvalidOperationException("headless browser produced no output.");

            // The <title> is the AUTHORITATIVE outcome — not the presence of an <svg>.
            // On a syntax error mermaid.parse() rejects (we capture the message) but
            // STILL injects its red "Syntax error" bomb graphic into the DOM anyway
            // (suppressErrorRendering doesn't stop it). So a viewBox is present even on
            // failure; keying off the svg would screenshot the bomb and "succeed".
            // Trust the title: MMDREADY = real render, MMDSYNTAX = bad input, else infra.
            if (dom.Contains("<title>MMDSYNTAX</title>", StringComparison.Ordinal))
                throw new MermaidSyntaxException(
                    "mermaid syntax error: " + ExtractMermaidMessage(dom)
                    + "\n(fix the mermaid source, or use render=native for the built-in subset).");
            if (!dom.Contains("<title>MMDREADY</title>", StringComparison.Ordinal))
                throw new InvalidOperationException(
                    dom.Contains("<title>MMDERR</title>", StringComparison.Ordinal)
                    ? "mermaid failed to render: " + ExtractMermaidMessage(dom)
                    : "mermaid produced no diagram (mermaid.js failed to load or the render timed out).");

            var (w, h) = ParseSvgSize(dom);
            if (w <= 0 || h <= 0)
                throw new InvalidOperationException("mermaid rendered but produced no measurable svg viewBox.");

            var pngPath = Path.ChangeExtension(htmlPath, ".png");
            if (!HtmlScreenshot.CaptureChromeSized(htmlPath, pngPath,
                    (int)Math.Ceiling(w) + 2, (int)Math.Ceiling(h) + 2))
                throw new InvalidOperationException("headless screenshot failed.");
            return pngPath;
        }
        finally { try { File.Delete(htmlPath); } catch { /* best effort */ } }
    }

    /// <summary>Cache → one-time download → live CDN. Returns a URL usable as a
    /// &lt;script src&gt; (a <c>file://</c> for a cached/downloaded copy, else the CDN).</summary>
    private static string ResolveMermaidJsRef()
    {
        try
        {
            if (File.Exists(CachedJsPath) && new FileInfo(CachedJsPath).Length > 500_000)
                return new Uri(CachedJsPath).AbsoluteUri;

            Directory.CreateDirectory(CacheDir);
            foreach (var url in new[] { MirrorUrl, CdnUrl }) // mirror first, CDN fallback
            {
                try
                {
                    using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(30) };
                    var bytes = http.GetByteArrayAsync(url).GetAwaiter().GetResult();
                    if (bytes.Length > 500_000)
                    {
                        File.WriteAllBytes(CachedJsPath, bytes);
                        return new Uri(CachedJsPath).AbsoluteUri;
                    }
                }
                catch { /* try next source */ }
            }
        }
        catch { /* fall through to live CDN */ }
        return CdnUrl; // every download failed → reference the CDN directly in the page
    }

    private static string BuildHtml(string mermaid, string jsRef, string background)
    {
        // Pass the source as base64 so no mermaid character can break out of the
        // HTML/JS context. Render explicitly (startOnLoad:false + mermaid.run) and
        // stamp the title so failures are recoverable from the dumped DOM.
        var b64 = Convert.ToBase64String(Encoding.UTF8.GetBytes(mermaid));
        return
            "<!DOCTYPE html><html><head><meta charset=\"utf-8\">"
            // svg{display:block}: an inline <svg> sits on the text baseline, leaving a
            // descender gap BELOW it inside the inline-block wrapper. The fixed-height
            // screenshot window then clips those few pixels → the diagram's bottom edge
            // (e.g. a sequence diagram's bottom actor boxes) gets cut. block removes it.
            + "<style>html,body{margin:0;padding:0;background:" + background + ";font-size:0}"
            + "#d{display:inline-block}#d svg{display:block}</style>"
            + $"<script src=\"{jsRef}\"></script></head>"
            // Hidden sink for a failure message. mermaid's parse error is multi-line
            // with an aligned caret ('^') under the offending column; document.title
            // would flatten the newlines and misalign the caret, so carry the verbatim
            // message here (a <pre> preserves whitespace) and use the title only as the
            // one-word outcome SIGNAL (MMDREADY / MMDSYNTAX / MMDERR).
            + "<body><pre id=\"mmderr\" style=\"display:none\"></pre>"
            + "<div id=\"d\" class=\"mermaid\"></div><script>"
            // atob yields a BYTE string (one Latin-1 char per byte); decode those
            // bytes back as UTF-8 so CJK/emoji in the mermaid source survive. A bare
            // atob() would render "提交" as mojibake ("æ¤").
            + $"const src=new TextDecoder().decode(Uint8Array.from(atob(\"{b64}\"),c=>c.charCodeAt(0)));"
            + "document.getElementById('d').textContent=src;"
            + "window.addEventListener('load',async()=>{try{"
            // htmlLabels:false → mermaid emits real SVG <text> instead of <foreignObject>
            // (HTML), which Office's SVG renderer cannot display — otherwise every
            // node/label comes out blank. securityLevel:loose allows the run.
            // suppressErrorRendering: on invalid syntax mermaid otherwise silently
            // renders a red "Syntax error in text" bomb graphic and returns success —
            // we would screenshot the bomb and embed it. With this off, run() throws.
            + "mermaid.initialize({startOnLoad:false,securityLevel:'loose',htmlLabels:false,"
            + "flowchart:{htmlLabels:false},class:{htmlLabels:false},suppressErrorRendering:true});"
            // Validate first: mermaid.parse() throws a precise, line-numbered error
            // ('Parse error on line N: … Expecting X, got Y') without touching the DOM.
            // Stamp it under a DISTINCT title so the host tells a user syntax error
            // (surface it, let the agent fix the source) apart from an infra failure
            // (browser/mermaid.js problem — fall back to the native synthesizer).
            + "try{await mermaid.parse(src);}catch(pe){"
            + "document.getElementById('mmderr').textContent=(pe&&pe.message?pe.message:String(pe));"
            + "document.title='MMDSYNTAX';return;}"
            + "await mermaid.run({nodes:[document.getElementById('d')]});"
            // Tighten the SVG to its REAL content bounds. mermaid's own viewBox
            // overshoots for some types (sequence diagrams reserve far more width/
            // height than they draw), which otherwise bakes a big transparent band
            // into the screenshot. getBBox() is the true rendered geometry; rewrite
            // viewBox + width/height to it (+small pad) so the capture is a tight crop.
            // pad clears ink that getBBox ignores: getBBox returns pure geometry, but
            // mermaid drop-shadows (filter:drop-shadow(3px 5px 2px …)) paint several px
            // past it — a 4px pad clipped the bottom actor boxes of a sequence diagram.
            // 14 covers the largest shadow (offset 5 + blur 2) with margin to spare.
            + "try{const s=document.querySelector('#d svg');if(s){const b=s.getBBox();"
            + "const p=14,x=b.x-p,y=b.y-p,w=Math.ceil(b.width+2*p),h=Math.ceil(b.height+2*p);"
            + "s.setAttribute('viewBox',x+' '+y+' '+w+' '+h);"
            + "s.setAttribute('width',w);s.setAttribute('height',h);"
            + "s.style.maxWidth=w+'px';s.style.width=w+'px';s.style.height=h+'px';}}catch(e){}"
            + "document.title='MMDREADY';"
            + "}catch(e){document.getElementById('mmderr').textContent="
            + "(e&&e.message?e.message:String(e));document.title='MMDERR';}});"
            + "</script></body></html>";
    }

    /// <summary>
    /// ESM variant of <see cref="BuildHtml"/> for a styled source: imports the ESM
    /// mermaid build (which, unlike the UMD global, renders a frontmatter-carrying
    /// source) and, when the source selects ELK, the ESM elk loader — registered
    /// via <c>registerLayoutLoaders</c> before init. Same title/mmderr/getBBox
    /// contract as BuildHtml so the DOM-dump reader is unchanged. Assets load live
    /// from the CDN (see the ESM-url comment above); on failure the title stays
    /// unset → the host treats it as an infra failure and falls back to native.
    /// </summary>
    private static string BuildHtmlEsm(string mermaid, bool needsElk, string background)
    {
        var b64 = Convert.ToBase64String(Encoding.UTF8.GetBytes(mermaid));
        var elkImport = needsElk
            ? $"const elk=(await import(\"{ElkEsmUrl}\")).default;mermaid.registerLayoutLoaders(elk);"
            : "";
        return
            "<!DOCTYPE html><html><head><meta charset=\"utf-8\">"
            + "<style>html,body{margin:0;padding:0;background:" + background + ";font-size:0}"
            + "#d{display:inline-block}#d svg{display:block}</style></head>"
            + "<body><pre id=\"mmderr\" style=\"display:none\"></pre>"
            + "<div id=\"d\" class=\"mermaid\"></div><script type=\"module\">"
            + "try{"
            + $"const mermaid=(await import(\"{MermaidEsmUrl}\")).default;"
            + elkImport
            + "const src=new TextDecoder().decode(Uint8Array.from(atob(\"" + b64 + "\"),c=>c.charCodeAt(0)));"
            + "document.getElementById('d').textContent=src;"
            + "mermaid.initialize({startOnLoad:false,securityLevel:'loose',htmlLabels:false,"
            + "flowchart:{htmlLabels:false},class:{htmlLabels:false},suppressErrorRendering:true});"
            + "try{await mermaid.parse(src);}catch(pe){"
            + "document.getElementById('mmderr').textContent=(pe&&pe.message?pe.message:String(pe));"
            + "document.title='MMDSYNTAX';throw pe;}"
            + "await mermaid.run({nodes:[document.getElementById('d')]});"
            + "try{const s=document.querySelector('#d svg');if(s){const b=s.getBBox();"
            + "const p=14,x=b.x-p,y=b.y-p,w=Math.ceil(b.width+2*p),h=Math.ceil(b.height+2*p);"
            + "s.setAttribute('viewBox',x+' '+y+' '+w+' '+h);"
            + "s.setAttribute('width',w);s.setAttribute('height',h);"
            + "s.style.maxWidth=w+'px';s.style.width=w+'px';s.style.height=h+'px';}}catch(e){}"
            + "document.title='MMDREADY';"
            + "}catch(e){if(document.title!=='MMDSYNTAX'){"
            + "document.getElementById('mmderr').textContent=(e&&e.message?e.message:String(e));"
            + "document.title='MMDERR';}}"
            + "</script></body></html>";
    }

    /// <summary>Pull the verbatim failure message out of the hidden &lt;pre id="mmderr"&gt;
    /// in the dumped DOM. The &lt;pre&gt; preserves mermaid's newlines (so its caret '^'
    /// stays aligned under the offending column); the browser HTML-escapes the text, so
    /// undo the common entities WITHOUT collapsing whitespace. Falls back to a generic
    /// note if the sink is missing.</summary>
    private static string ExtractMermaidMessage(string dom)
    {
        var m = Regex.Match(dom, "<pre id=\"mmderr\"[^>]*>(.*?)</pre>", RegexOptions.Singleline);
        if (!m.Success || string.IsNullOrWhiteSpace(m.Groups[1].Value))
            return "(no detail reported)";
        var s = m.Groups[1].Value
            .Replace("&amp;", "&").Replace("&lt;", "<").Replace("&gt;", ">")
            .Replace("&quot;", "\"").Replace("&#39;", "'");
        return s.Trim('\n', '\r', ' ', '\t');
    }

    /// <summary>Read the rendered diagram's CSS-pixel size from the svg viewBox in
    /// the dumped DOM (mermaid writes <c>viewBox="0 0 W H"</c>). (0,0) if not found.</summary>
    private static (double w, double h) ParseSvgSize(string dom)
    {
        var m = Regex.Match(dom, @"<svg[^>]*\bviewBox=""[\d.\-]+\s+[\d.\-]+\s+([\d.]+)\s+([\d.]+)""",
            RegexOptions.IgnoreCase);
        if (m.Success
            && double.TryParse(m.Groups[1].Value, System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out var w)
            && double.TryParse(m.Groups[2].Value, System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out var h))
            return (w, h);
        return (0, 0);
    }
}
