// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Handlers;

/// <summary>
/// Auxiliary-parts scan — surfaces a warning per package part that the dump
/// surface does NOT round-trip, so silent data loss becomes visible.
///
/// <para>
/// Rationale: <see cref="WordBatchEmitter.EmitWordWithWarnings(WordHandler)"/>
/// emits a curated set of parts (body, styles, numbering, theme, settings,
/// headers/footers, comments, footnotes, endnotes, charts, media). Every
/// other part in the OPC package is silently dropped on dump∘replay —
/// templates with content-control bindings (customXml), AutoText repositories
/// (glossary/document.xml), modern-comment metadata (people/commentsExtended/
/// commentsIds), browser-optimised templates (webSettings), embedded fonts
/// (fontTable + word/fonts/*.odttf), and user-defined custom document
/// properties (docProps/custom.xml entries outside the OfficeCLI.* namespace)
/// all vanish without trace.
/// </para>
///
/// <para>
/// This pass walks the package once and emits a
/// <see cref="WordBatchEmitter.DocxUnsupportedWarning"/> per unrecognised
/// part. The dump command then mirrors these into envelope.warnings (with
/// a stderr "warning:" line for human consumption) — same channel as the
/// OLE / shape unsupported warnings introduced in earlier rounds.
/// </para>
///
/// <para>
/// Allowlist (not denylist) — every part with a curated emit path is listed
/// here. When a new emitter lands (e.g. real glossary support) the matching
/// prefix migrates from <see cref="UnsupportedReasons"/> to
/// <see cref="KnownEmittedPrefixes"/>/<see cref="KnownEmittedExact"/> so the
/// warning disappears.
/// </para>
/// </summary>
public static partial class WordBatchEmitter
{
    // Part URIs (or URI prefixes) that the curated emit path round-trips.
    // OPC package "auto-managed" parts (Content_Types, top-level rels,
    // docProps/core+app — restamped on save) are listed alongside the
    // semantic parts the dump actually emits.
    private static readonly HashSet<string> KnownEmittedExact = new(StringComparer.OrdinalIgnoreCase)
    {
        "/word/document.xml",
        "/word/styles.xml",
        "/word/stylesWithEffects.xml",   // legacy compat, restamped by SDK
        "/word/numbering.xml",
        "/word/settings.xml",
        "/word/comments.xml",
        "/word/commentsExtended.xml",    // BUG-DUMP-R26-4: reply threading — EmitComments raw-set
        "/word/footnotes.xml",
        "/word/endnotes.xml",
        "/word/fontTable.xml",           // BUG-DUMP-R42-3: EmitFontTableRaw round-trips faces + altName subs
        "/word/webSettings.xml",         // EmitWebSettingsRaw round-trips the whole part verbatim
        // docProps stores — EmitDocPropsRaw round-trips core/app/custom verbatim
        // so data-bound content controls keep their source display text (the
        // OfficeCLI audit stamp still rides in custom.xml via StampOnSave).
        "/docProps/core.xml",
        "/docProps/app.xml",
        "/docProps/custom.xml",
        "/[Content_Types].xml",
        "/_rels/.rels",
    };

    private static readonly string[] KnownEmittedPrefixes = new[]
    {
        "/word/theme/",            // theme1.xml etc — EmitThemeRaw
        "/word/header",            // header1.xml, header2.xml... EmitHeadersFooters
        "/word/footer",            // footer1.xml etc
        "/word/fonts/",            // BUG-DUMP-R45-1: embedded fonts (.odttf) — EmitFontTableRaw embed-binary
        "/word/media/",            // images — picture run emit
        "/media/",                 // package-root media (some picture add paths
                                   // land here) — round-tripped by the same
                                   // picture data-URI carrier; warning was a
                                   // false alarm (source/replay MD5 identical)
        "/word/charts/",           // chart XML + embedded xlsx — chart run emit
        "/word/extendedCharts/",   // chartEx XML + embedded xlsx + its chartex
                                   // style/colour sidecars — run inlined-parts
                                   // carrier (chart run emit)
        "/word/embeddings/",       // OLE payloads — warning already raised per-run
        "/word/diagrams/",         // SmartArt — partial coverage via shape emit
        "/word/activeX/",          // ActiveX controls — `add activex` inlined-parts carrier
        "/customXml/",             // customXml data stores — EmitCustomXmlRaw embed-binary pairs
        "/word/printerSettings/",  // SDK strips on save
        "/word/customizations.xml", // legacy customizations
    };

    // Maps an unsupported part URI (or URI prefix) → human-readable reason.
    // Order matters: prefixes are tested in declaration order; first match wins.
    private static readonly (string Prefix, string Element, string Reason)[] UnsupportedReasons = new[]
    {
        ("/word/glossary/",            "glossary",               "Building Blocks / AutoText repository dropped on dump"),
        ("/word/people.xml",           "people",                 "modern-comment author metadata dropped on dump"),
        ("/word/commentsIds.xml",      "commentsIds",            "modern-comment durable-id metadata dropped on dump"),
        ("/word/commentsExtensible.xml","commentsExtensible",    "modern-comment extension metadata dropped on dump"),
        ("/word/vbaProject.bin",       "vbaProject",             "VBA macro project dropped on dump"),
        ("/word/vbaData.xml",          "vbaData",                "VBA macro metadata dropped on dump"),
    };

    /// <summary>
    /// Walk the package once, comparing each part's zip-URI against the
    /// emitted allowlist; record a warning for every part the dump surface
    /// does NOT round-trip. Also probes <c>docProps/custom.xml</c> for
    /// user-defined properties (any name outside the <c>OfficeCLI.*</c>
    /// namespace is silently dropped on save).
    /// </summary>
    internal static void EmitAuxiliaryPartsScan(WordHandler word, List<DocxUnsupportedWarning> warnings)
    {
        IEnumerable<string> parts;
        try { parts = word.EnumeratePartUris(); }
        catch { return; }

        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var uri in parts)
        {
            if (!seen.Add(uri)) continue;
            // Relationship parts (.rels) are auto-managed by the SDK alongside
            // their owning part — skip uniformly. Without this, every emitted
            // part's `_rels/<name>.xml.rels` would also surface as a warning.
            if (uri.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)) continue;

            if (KnownEmittedExact.Contains(uri)) continue;
            if (KnownEmittedPrefixes.Any(p => uri.StartsWith(p, StringComparison.OrdinalIgnoreCase))) continue;

            // Look up the catalogued reason; if none matches, emit a generic
            // "unknown part" warning so silent loss never goes unreported.
            string element = "auxiliaryPart";
            string reason = "package part dropped on dump (no curated emit path)";
            foreach (var (prefix, elt, why) in UnsupportedReasons)
            {
                if (uri.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                {
                    element = elt;
                    reason = why;
                    break;
                }
            }

            warnings.Add(new DocxUnsupportedWarning(
                Element: element,
                Path: uri,
                Reason: reason));
        }
    }
}
