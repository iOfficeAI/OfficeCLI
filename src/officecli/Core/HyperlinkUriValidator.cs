// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Core;

/// <summary>
/// Cross-handler allowlist for user-supplied hyperlink URI schemes.
///
/// PowerPoint/Excel/Word would otherwise happily write any URI a caller
/// hands in (javascript:, file://, data:, vbscript:) into the document's
/// .rels Target. That round-trips cleanly but lets a malicious caller plant
/// click-bait that triggers script execution or local-file exfiltration on
/// recipients who follow the link. The OOXML format itself does not gate
/// scheme — every Office product applies its own runtime warning UI on top
/// — so we reject unsafe schemes at write time and keep the document clean.
///
/// Handler-internal targets (PowerPoint's ppaction://, slide://, named
/// actions like firstslide/nextslide, fragment anchors like #_ftn1,
/// in-workbook references like Sheet!A1, or any non-absolute URI) are
/// resolved by the handler before this validator is consulted; callers only
/// pass strings here once they have been classified as an external URI.
/// </summary>
public static class HyperlinkUriValidator
{
    // Schemes that survive an Office "is this link safe?" prompt without
    // user warnings. http/https/mailto are the everyday cases; ftp/sms/tel
    // /news are the standard PowerPoint "Action button" set; ppaction is
    // PowerPoint's internal navigation pseudo-scheme and is allowed so a
    // caller can paste a ppaction:// URI it read from another file.
    private static readonly HashSet<string> AllowedSchemes = new(StringComparer.OrdinalIgnoreCase)
    {
        "http",
        "https",
        "mailto",
        "ftp",
        "ftps",
        "sftp",
        "news",
        "tel",
        "sms",
        "ppaction",
    };

    /// <summary>
    /// Validate an external hyperlink URI's scheme. Throws ArgumentException
    /// with a deterministic, agent-readable message when the scheme is not
    /// in the allowlist. Empty / null input is a no-op so the caller's own
    /// "missing URL" diagnostic remains the surfaced error.
    /// </summary>
    public static void RequireSafeScheme(string url, string contextKey = "link")
    {
        if (string.IsNullOrEmpty(url)) return;
        if (!Uri.TryCreate(url, UriKind.Absolute, out var uri)) return; // not absolute → handler-internal path, not our concern
        var scheme = uri.Scheme;
        if (string.IsNullOrEmpty(scheme)) return;
        if (AllowedSchemes.Contains(scheme)) return;
        throw new ArgumentException(
            $"Invalid {contextKey} URL scheme '{scheme}:': only http, https, mailto, ftp, ftps, sftp, news, tel, sms, and ppaction targets are accepted. " +
            "javascript:, file:, data:, vbscript:, and similar schemes are rejected to prevent click-bait redirection in shared documents.");
    }
}
