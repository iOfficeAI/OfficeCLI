// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Security.Cryptography;
using System.Text.RegularExpressions;

namespace OfficeCli.Core;

/// <summary>
/// Content-Security-Policy for the generated preview pages (view html / svg,
/// screenshot, watch). Defence in depth behind the per-value escaping in the
/// renderers: even if a document-controlled string reaches the markup, it
/// cannot run script, because the policy only executes scripts that carry
/// this process's nonce — inline event-handler attributes, javascript: URLs
/// and un-nonced &lt;script&gt; elements are refused by the browser.
///
/// Every script the renderers emit must therefore be written with
/// <see cref="NonceAttr"/>, and must attach its handlers with
/// addEventListener rather than on*="" attributes.
///
/// The nonce is per process, not per render: one command may render a full
/// page and separate fragments (pptx slides pushed to watch), and they have to
/// agree. The watch server re-stamps incoming HTML with its own nonce
/// (<see cref="Rebind"/>), so a browser tab keeps one nonce for its lifetime.
/// </summary>
internal static class PreviewCsp
{
    private static readonly Lazy<string> _nonce = new(() =>
        Convert.ToHexString(RandomNumberGenerator.GetBytes(16)).ToLowerInvariant());

    public static string Nonce => _nonce.Value;

    /// <summary>` nonce="…"` — append inside every emitted &lt;script&gt; start tag.</summary>
    public static string NonceAttr => $" nonce=\"{Nonce}\"";

    // style-src keeps 'unsafe-inline': the previews position everything with
    // inline styles. Styling is not the boundary here; script execution is.
    // 'strict-dynamic' lets a nonced script load its own dependencies (KaTeX
    // CDN fallback, three.js module graph) without a host allowlist — a CDN
    // host in script-src would let injected markup load any package from it.
    private static string Policy(string nonce) =>
        "default-src 'self' data: blob: https:; " +
        $"script-src 'nonce-{nonce}' 'strict-dynamic'; " +
        "style-src 'self' 'unsafe-inline' https:; " +
        "img-src 'self' data: blob: https: http:; " +
        "object-src 'none'; frame-src 'none'; base-uri 'none'; form-action 'none'";

    /// <summary>
    /// The policy as a &lt;meta&gt; tag. Must be the first thing in &lt;head&gt;
    /// after the charset: a meta policy only governs what follows it.
    /// </summary>
    public static string MetaTag =>
        $"<meta http-equiv=\"Content-Security-Policy\" content=\"{Policy(Nonce)}\">";

    private static readonly Regex NonceInPolicy =
        new(@"<meta http-equiv=""Content-Security-Policy"" content=""[^""]*'nonce-([0-9a-f]{32})'", RegexOptions.Compiled);

    /// <summary>The nonce declared by the first policy meta in <paramref name="html"/>, or null.</summary>
    public static string? ExtractNonce(string? html)
    {
        if (string.IsNullOrEmpty(html)) return null;
        var m = NonceInPolicy.Match(html);
        return m.Success ? m.Groups[1].Value : null;
    }

    /// <summary>
    /// Re-stamp HTML rendered by another process (nonce <paramref name="from"/>)
    /// with nonce <paramref name="to"/>. The token is 128 random bits, so a
    /// plain replace cannot touch document content.
    /// </summary>
    public static string Rebind(string html, string? from, string to)
    {
        if (string.IsNullOrEmpty(html) || string.IsNullOrEmpty(from) || from == to) return html;
        return html.Replace(from, to, StringComparison.Ordinal);
    }
}
