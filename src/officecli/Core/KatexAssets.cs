// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Core;

/// <summary>
/// Single source of truth for KaTeX asset URLs in generated HTML.
/// Mirrors the mermaid sourcing policy (see <c>MermaidImageRenderer</c>):
/// own mirror first (no third-party dependency at steady state, reachable in
/// networks where jsdelivr is not), public CDN as the fallback. The mirror
/// hosts the full dist subtree (js + css + fonts/) under one immutable
/// versioned prefix, so the css's relative <c>url(fonts/…)</c> references
/// resolve against whichever origin actually served it.
///
/// Unlike mermaid there is no local file cache: mermaid.js is executed by
/// the CLI's own headless browser against a throwaway file, while these URLs
/// are baked into preview HTML that users keep and open elsewhere — a
/// <c>file://~/.officecli/…</c> reference would break the moment the HTML
/// leaves the machine. Load robustness comes from the mirror→CDN onerror
/// chain plus the screenshot path's --virtual-time-budget/--timeout caps.
/// </summary>
internal static class KatexAssets
{
    public const string Version = "0.16.11";

    private const string MirrorBase = "https://d.officecli.ai/assets/katex-" + Version;
    private const string CdnBase = "https://cdn.jsdelivr.net/npm/katex@" + Version + "/dist";

    public static string CssUrl => MirrorBase + "/katex.min.css";
    public static string JsUrl => MirrorBase + "/katex.min.js";
    public static string CdnCssUrl => CdnBase + "/katex.min.css";
    public static string CdnJsUrl => CdnBase + "/katex.min.js";

    // The mirror-first / CDN-retry / final-fallback chain used to live in
    // onload="" / onerror="" attributes on the two tags. The preview pages now
    // ship a Content-Security-Policy that refuses inline handler attributes
    // (see PreviewCsp), so the same chain is attached by a nonced script
    // instead. It listens in the capture phase on the document and is emitted
    // BEFORE the tags, so an error that fires immediately (cached failure)
    // cannot slip past it. The tags themselves stay parser-inserted — the
    // deferred script must still run before DOMContentLoaded, which the
    // renderers' init code relies on.
    private const string CssId = "ocli-katex-css";
    private const string JsId = "ocli-katex-js";

    /// <summary>
    /// The &lt;link&gt; + &lt;script&gt; pair that loads KaTeX, with its fallback
    /// chain. <paramref name="finalFallbackJs"/> runs when both the mirror and
    /// the CDN script fail.
    /// </summary>
    public static string HeadTags(string finalFallbackJs) =>
        $"<script{PreviewCsp.NonceAttr}>(function(){{"
        + "document.addEventListener('load',function(e){var t=e.target;"
        + $"if(t&&t.id==='{CssId}')t.media='all';}},true);"
        + "document.addEventListener('error',function(e){var t=e.target;if(!t)return;"
        + $"if(t.id==='{CssId}'){{if(!t.dataset.f){{t.dataset.f=1;t.href='{CdnCssUrl}'}}else{{t.remove()}}}}"
        + $"else if(t.id==='{JsId}'){{var s=document.createElement('script');s.src='{CdnJsUrl}';"
        + "s.onerror=function(){" + finalFallbackJs + "};document.head.appendChild(s)}"
        + "},true);})();</script>\n"
        + $"<link id=\"{CssId}\" rel=\"stylesheet\" href=\"{CssUrl}\" media=\"print\">\n"
        + $"<script id=\"{JsId}\" defer src=\"{JsUrl}\"{PreviewCsp.NonceAttr}></script>";
}
