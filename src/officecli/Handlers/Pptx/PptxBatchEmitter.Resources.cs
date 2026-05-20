// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using OfficeCli.Core;

namespace OfficeCli.Handlers;

public static partial class PptxBatchEmitter
{
    // CONSISTENCY(emit-resources-mirror): mirrors WordBatchEmitter.Resources.cs
    // — each whole-part-XML block emits as a single raw-set replace. Theme /
    // master / layout / notesMaster carry rich structured XML (clrScheme,
    // fontScheme, txStyles, fmtScheme, …) that has no typed Set vocabulary; the
    // natural operation is "swap the whole block". Replay's raw-set overwrites
    // whatever the blank deck stamped during BlankDocCreator.

    // CONSISTENCY(raw-xmlns-canonicalize): mirrors
    // WordBatchEmitter.Resources.CanonicalizeRawXml. RawXmlHelper.Execute
    // propagates the root's xmlns declarations onto every direct child so the
    // SDK's InnerXml setter can resolve prefixes (SDK does not inherit root
    // xmlns scope when parsing inner content). After replay, the part's XML
    // carries redundant xmlns:p / xmlns:a attrs on each child of /theme,
    // /slideMaster[N], /slideLayout[N] — observed first-replay growth on a
    // blank-deck round-trip: 16657 → 17923 bytes (≈1.2 KB across 7 raw-set
    // parts), then stable on subsequent rounds. Canonicalise on emit so the
    // first-pass (clean source) and second-pass (post-replay bloated) shapes
    // collapse identically.
    private static string CanonicalizeRawXml(string xml)
    {
        if (string.IsNullOrEmpty(xml) || !xml.StartsWith("<")) return xml;
        try
        {
            var doc = System.Xml.Linq.XDocument.Parse(xml);
            if (doc.Root == null) return xml;
            var rootNsAttrs = doc.Root.Attributes()
                .Where(a => a.IsNamespaceDeclaration)
                .ToDictionary(a => a.Name, a => a.Value);
            foreach (var desc in doc.Root.Descendants())
            {
                var toRemove = desc.Attributes()
                    .Where(a => a.IsNamespaceDeclaration
                                && rootNsAttrs.TryGetValue(a.Name, out var v)
                                && v == a.Value)
                    .ToList();
                foreach (var a in toRemove) a.Remove();
            }
            return doc.Root.ToString(System.Xml.Linq.SaveOptions.DisableFormatting);
        }
        catch
        {
            // Malformed XML — leave as-is rather than corrupting.
            return xml;
        }
    }

    private static void EmitThemeRaw(PowerPointHandler ppt, List<BatchItem> items)
    {
        // Pptx Raw("/theme") returns the presentation-level theme part (first
        // master's theme). Multi-master decks have additional theme parts
        // attached to each master, but the existing Raw/RawSet surface only
        // addresses the primary one — keep parity until per-master theme
        // raw-set lands. Skip silently when the source has none.
        string xml;
        try { xml = ppt.Raw("/theme"); }
        catch { return; }
        if (string.IsNullOrEmpty(xml) || !xml.StartsWith("<") || xml == "(no theme)")
            return;
        xml = CanonicalizeRawXml(xml);

        items.Add(new BatchItem
        {
            Command = "raw-set",
            Part = "/theme",
            Xpath = "/a:theme",
            Action = "replace",
            Xml = xml
        });
    }

    private static void EmitNotesMasterRaw(PowerPointHandler ppt, List<BatchItem> items)
    {
        if (!ppt.HasNotesMaster) return;
        string xml;
        try { xml = ppt.Raw("/notesMaster"); }
        catch { return; }
        if (string.IsNullOrEmpty(xml) || !xml.StartsWith("<")) return;
        xml = CanonicalizeRawXml(xml);

        items.Add(new BatchItem
        {
            Command = "raw-set",
            Part = "/notesMaster",
            Xpath = "/p:notesMaster",
            Action = "replace",
            Xml = xml
        });
    }

    private static void EmitMasterRaw(PowerPointHandler ppt, List<BatchItem> items)
    {
        var n = ppt.SlideMasterCount;
        for (int i = 1; i <= n; i++) EmitMasterRawOne(ppt, i, items);
    }

    private static bool EmitMasterRawOne(PowerPointHandler ppt, int idx, List<BatchItem> items)
    {
        string xml;
        try { xml = ppt.Raw($"/slideMaster[{idx}]"); }
        catch { return false; }
        if (string.IsNullOrEmpty(xml) || !xml.StartsWith("<")) return false;
        xml = CanonicalizeRawXml(xml);

        items.Add(new BatchItem
        {
            Command = "raw-set",
            Part = $"/slideMaster[{idx}]",
            Xpath = "/p:sldMaster",
            Action = "replace",
            Xml = xml
        });
        return true;
    }

    private static void EmitLayoutRaw(PowerPointHandler ppt, List<BatchItem> items)
    {
        var n = ppt.SlideLayoutCount;
        for (int i = 1; i <= n; i++) EmitLayoutRawOne(ppt, i, items);
    }

    private static bool EmitLayoutRawOne(PowerPointHandler ppt, int idx, List<BatchItem> items)
    {
        string xml;
        try { xml = ppt.Raw($"/slideLayout[{idx}]"); }
        catch { return false; }
        if (string.IsNullOrEmpty(xml) || !xml.StartsWith("<")) return false;
        xml = CanonicalizeRawXml(xml);

        items.Add(new BatchItem
        {
            Command = "raw-set",
            Part = $"/slideLayout[{idx}]",
            Xpath = "/p:sldLayout",
            Action = "replace",
            Xml = xml
        });
        return true;
    }

    private static bool EmitNoteSlideRawOne(PowerPointHandler ppt, int idx, List<BatchItem> items)
    {
        string xml;
        try { xml = ppt.Raw($"/noteSlide[{idx}]"); }
        catch { return false; }
        if (string.IsNullOrEmpty(xml) || !xml.StartsWith("<")) return false;
        xml = CanonicalizeRawXml(xml);

        items.Add(new BatchItem
        {
            Command = "raw-set",
            Part = $"/noteSlide[{idx}]",
            Xpath = "/p:notes",
            Action = "replace",
            Xml = xml
        });
        return true;
    }
}
