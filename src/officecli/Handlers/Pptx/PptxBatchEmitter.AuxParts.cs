// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Handlers;

/// <summary>
/// Auxiliary-parts scan — surfaces a warning per package part that the dump
/// surface does NOT round-trip, so silent data loss becomes visible.
///
/// <para>
/// Rationale (mirrors <see cref="WordBatchEmitter.EmitAuxiliaryPartsScan"/>):
/// the typed pptx emit path covers presentation / masters / layouts / slides /
/// notes / theme / media / charts / embeddings / SmartArt diagrams / comments
/// (legacy + modern). Every other part in the OPC package is silently dropped
/// on dump∘replay — view-pane settings, custom table styles, handout masters,
/// printer settings, content-control bindings (customXml), embedded fonts,
/// programmability tags, and user-defined custom document properties all
/// vanish without trace.
/// </para>
///
/// <para>
/// This pass walks the package once (SDK part graph ∪ raw zip entries — the
/// union catches both linked parts and orphan zip entries the SDK refuses to
/// surface) and records a <see cref="PptxBatchEmitter.UnsupportedWarning"/>
/// per unrecognised part. The dump command then mirrors these into envelope
/// warnings (with a stderr "warning:" line for human consumption) — same
/// channel as the per-slide unsupported warnings already emitted by
/// ProbeUnsupportedOnSlide.
/// </para>
///
/// <para>
/// Allowlist (not denylist) — every part with a curated emit path is listed
/// here. When a new emitter lands, the matching prefix migrates from
/// <see cref="UnsupportedReasons"/> to <see cref="KnownEmittedPrefixes"/> /
/// <see cref="KnownEmittedExact"/> so the warning disappears.
/// </para>
/// </summary>
public static partial class PptxBatchEmitter
{
    // Part URIs (or URI prefixes) that the curated emit path round-trips.
    // OPC package "auto-managed" parts (Content_Types, top-level rels,
    // docProps/core+app — restamped on save) are listed alongside the
    // semantic parts the dump actually emits.
    private static readonly HashSet<string> AuxKnownEmittedExact = new(StringComparer.OrdinalIgnoreCase)
    {
        "/ppt/presentation.xml",
        "/ppt/tableStyles.xml",          // custom table-style catalogue — EmitTableStyles
        // OPC auto-managed
        "/docProps/core.xml",            // restamped by OfficeCliMetadata
        "/docProps/app.xml",             // restamped by OfficeCliMetadata
        "/[Content_Types].xml",
        "/_rels/.rels",
    };

    private static readonly string[] AuxKnownEmittedPrefixes = new[]
    {
        "/ppt/theme/",            // theme1.xml etc — EmitThemeRaw
        "/ppt/slideMasters/",     // EmitMasterRaw
        "/ppt/slideLayouts/",     // EmitLayoutRaw
        "/ppt/slides/",           // EmitSlide (typed per-slide emit)
        "/ppt/notesMasters/",     // EmitNotesMasterRaw
        "/ppt/notesSlides/",      // EmitNotes per-slide
        "/ppt/media/",            // picture/media embed — EmitPicture / EmitMediaForSlide
        "/media/",                // package-root media — same picture/media carriers
        "/ppt/embeddings/",       // chart xlsx / OLE payloads — EmitChart / EmitOleForSlide
        "/ppt/charts/",           // chart XML — EmitChart
        "/ppt/diagrams/",         // SmartArt — EmitSmartArtsForSlide
        "/ppt/printerSettings/",  // presentation-level ExtendedPart — carried via add-part extpart (GetPresentationExtendedParts)
    };

    // Maps an unsupported part URI (or URI prefix) → human-readable reason.
    // Order matters: prefixes are tested in declaration order; first match wins.
    //
    // Note on commentAuthors / commentAuthorsExtended: the comment emit chain
    // (EmitComments / EmitModernComments) re-creates these parts on replay
    // WHEN comments exist. They're listed here anyway because a deck can
    // carry an authors part with no comments (legacy author cleanup leftover)
    // — in that case our emit chain produces nothing, and the part silently
    // disappears. Surfacing the warning every time is the safer default;
    // when comments do exist, the matching add rows confirm the authors
    // metadata will be regenerated.
    private static readonly (string Prefix, string Element, string Reason)[] AuxUnsupportedReasons = new[]
    {
        ("/ppt/commentAuthors.xml",         "commentAuthors",         "legacy comment-authors metadata dropped on dump (regenerated on replay only if slides carry legacy comments)"),
        ("/ppt/commentAuthorsExtended.xml", "commentAuthorsExtended", "modern (Office 365) comment-authors extension dropped on dump (regenerated on replay only if slides carry modern threaded comments)"),
        ("/ppt/comments/",                  "legacyComment",          "legacy slide comment part dropped on dump (round-trips via per-slide `add comment` rows when EnumerateComments surfaces the content)"),
        ("/ppt/modernComments/",            "modernComment",          "modern threaded comment part dropped on dump (round-trips via per-slide `add modernComment` rows when EnumerateModernComments surfaces the content)"),
        ("/ppt/tableStyles.xml",            "tableStyles",            "custom table-style catalogue dropped on dump"),
        ("/ppt/viewProps.xml",              "viewProps",              "view-pane / zoom / sorter settings dropped on dump"),
        ("/ppt/handoutMasters/",            "handoutMaster",          "handout master dropped on dump"),
        ("/ppt/customXml/",                 "customXml",              "customXml part dropped on dump (custom data store / content-control bindings)"),
        ("/customXml/",                     "customXml",              "customXml part dropped on dump (custom data store / content-control bindings)"),
        ("/ppt/fonts/",                     "embeddedFont",           "embedded font binary (.fntdata) dropped on dump"),
        ("/ppt/tags/",                      "tags",                   "programmability tags (custom data binding) dropped on dump"),
        ("/ppt/vbaProject.bin",             "vbaProject",             "VBA macro project dropped on dump"),
    };

    /// <summary>
    /// Walk the package once, comparing each part's zip-URI against the
    /// emitted allowlist; record a warning for every part the dump surface
    /// does NOT round-trip. Also probes <c>docProps/custom.xml</c> for
    /// user-defined properties (any name outside the <c>OfficeCLI.*</c>
    /// namespace is silently dropped on save).
    /// </summary>
    internal static void EmitAuxiliaryPartsScan(PowerPointHandler ppt, List<UnsupportedWarning> warnings)
    {
        IEnumerable<string> parts;
        try { parts = ppt.EnumeratePartUris(); }
        catch { return; }

        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var uri in parts)
        {
            if (!seen.Add(uri)) continue;
            // Relationship parts (.rels) are auto-managed by the SDK alongside
            // their owning part — skip uniformly so every emitted part's
            // `_rels/<name>.xml.rels` does not surface as a warning.
            if (uri.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)) continue;
            // Zip directory entries (trailing slash) and the package root
            // surface as `/ppt/`, `/ppt/customXml/`, etc. via the raw-zip
            // pass; they carry no payload, so warning once per directory
            // would just duplicate the per-file warning already emitted.
            if (uri.EndsWith("/", StringComparison.Ordinal)) continue;

            if (AuxKnownEmittedExact.Contains(uri)) continue;
            if (AuxKnownEmittedPrefixes.Any(p => uri.StartsWith(p, StringComparison.OrdinalIgnoreCase))) continue;

            // Special-case: docProps/custom.xml — OfficeCliMetadata always
            // restamps OfficeCLI.* entries; user-authored entries are silently
            // dropped on save. Warn only if the part carries non-OfficeCLI
            // properties; the auto-stamped pair (OfficeCLI.Version + .LastModified)
            // is expected and not a loss.
            if (string.Equals(uri, "/docProps/custom.xml", StringComparison.OrdinalIgnoreCase))
            {
                var userProps = ppt.EnumerateCustomDocPropertyNames()
                    .Where(n => !n.StartsWith("OfficeCLI.", StringComparison.Ordinal))
                    .ToList();
                if (userProps.Count > 0)
                {
                    warnings.Add(new UnsupportedWarning(
                        Element: "customDocProperty",
                        SlidePath: uri,
                        Reason: $"user-defined custom document properties dropped on dump ({string.Join(", ", userProps)})"));
                }
                continue;
            }

            // Look up the catalogued reason; if none matches, emit a generic
            // "unknown part" warning so silent loss never goes unreported.
            string element = "auxiliaryPart";
            string reason = "package part dropped on dump (no curated emit path)";
            foreach (var (prefix, elt, why) in AuxUnsupportedReasons)
            {
                if (uri.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                {
                    element = elt;
                    reason = why;
                    break;
                }
            }

            warnings.Add(new UnsupportedWarning(
                Element: element,
                SlidePath: uri,
                Reason: reason));
        }
    }
}
