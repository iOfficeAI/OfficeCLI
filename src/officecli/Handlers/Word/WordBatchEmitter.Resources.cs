// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using OfficeCli.Core;

namespace OfficeCli.Handlers;

public static partial class WordBatchEmitter
{

    // RawXmlHelper.Execute propagates the root's xmlns declarations onto every
    // direct child so the SDK's InnerXml setter can resolve prefixes (SDK does
    // not inherit root xmlns scope when parsing inner content). After replay,
    // the part's XML carries redundant xmlns attrs on each child, which the
    // next dump reads back verbatim — phantom bloat that breaks idempotency.
    //
    // Canonicalize on emit: parse the part's XML, drop child-element xmlns
    // declarations that match the root's declarations, re-serialize. The
    // first-pass emit (source's clean XML) and second-pass emit (target's
    // bloated XML) both collapse to the same canonical shape.
    // Schema order for <w:ind>'s attributes. SDK serialises in this order
    // on write, so source-side OuterXml (which mirrors the on-disk order
    // from the original producer) and replay-target OuterXml (SDK's
    // canonical) can disagree on attribute order alone. Re-sort to a fixed
    // canonical order so both passes emit identical bytes.
    private static readonly string[] s_indAttrOrder =
    [
        "start", "end", "left", "right",
        "hanging", "firstLine",
        "startChars", "endChars", "leftChars", "rightChars",
        "hangingChars", "firstLineChars",
    ];

    private static void SortIndAttrs(System.Xml.Linq.XElement ind)
    {
        var attrs = ind.Attributes().ToList();
        // Keep xmlns declarations first (in original order), then sort
        // typed attrs by the schema-order table, then unknown attrs by name.
        var nsDecls = attrs.Where(a => a.IsNamespaceDeclaration).ToList();
        var typed = attrs.Where(a => !a.IsNamespaceDeclaration).ToList();
        int OrderKey(System.Xml.Linq.XAttribute a)
        {
            var idx = Array.IndexOf(s_indAttrOrder, a.Name.LocalName);
            return idx < 0 ? 99 : idx;
        }
        var sorted = typed.OrderBy(OrderKey).ThenBy(a => a.Name.LocalName, StringComparer.Ordinal).ToList();
        ind.RemoveAttributes();
        foreach (var a in nsDecls) ind.Add(a);
        foreach (var a in sorted) ind.Add(a);
    }

    private static void RenameAttr(System.Xml.Linq.XElement el, string fromLocal, string toLocal, string ns)
    {
        var fromName = System.Xml.Linq.XName.Get(fromLocal, ns);
        var toName = System.Xml.Linq.XName.Get(toLocal, ns);
        var src = el.Attribute(fromName);
        if (src == null) return;
        if (el.Attribute(toName) != null) { src.Remove(); return; }
        // Preserve attribute order: re-build the attribute list with the
        // rename applied in-place. SetAttributeValue(newName) by itself would
        // append the new attr at the tail and shift byte order.
        var rebuilt = el.Attributes()
            .Select(a => a.Name == fromName
                ? new System.Xml.Linq.XAttribute(toName, a.Value)
                : new System.Xml.Linq.XAttribute(a.Name, a.Value))
            .ToList();
        el.RemoveAttributes();
        foreach (var a in rebuilt) el.Add(a);
    }

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
            // SDK normalises bidi-aware <w:ind w:start="…"> ↔ <w:ind w:left="…">
            // (and end ↔ right) on serialisation depending on the document's
            // bidi state. The two forms are byte-different but semantically
            // equivalent in non-bidi documents. Canonicalise to the bidi-
            // aware names AND fix the attribute order so the dump pair emits
            // identical bytes regardless of SDK's choice.
            var wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            foreach (var ind in doc.Descendants(System.Xml.Linq.XName.Get("ind", wNs)))
            {
                RenameAttr(ind, "left", "start", wNs);
                RenameAttr(ind, "right", "end", wNs);
                // BIDI-aware character-count variants also drift through SDK
                // normalisation. proof_fixed family: <w:ind … w:leftChars="0" …>
                // → SDK rewrites as <w:ind … w:startChars="0" …>.
                RenameAttr(ind, "leftChars", "startChars", wNs);
                RenameAttr(ind, "rightChars", "endChars", wNs);
                SortIndAttrs(ind);
            }
            // Stabilise root attribute order: SDK serialises xmlns attrs in
            // an internal order that can shift when mc:Ignorable / other
            // typed attrs change, so byte-equal round-trip needs a canonical
            // ordering. Emit xmlns attrs first (sorted by prefix; default
            // xmlns first if any), then non-xmlns attrs (sorted by name).
            var root = doc.Root;
            var allAttrs = root.Attributes().ToList();
            foreach (var a in allAttrs) a.Remove();
            var nsAttrs = allAttrs.Where(a => a.IsNamespaceDeclaration)
                .OrderBy(a => a.Name == System.Xml.Linq.XNamespace.Xmlns + "xmlns" ? "" : a.Name.LocalName,
                         StringComparer.Ordinal)
                .ToList();
            var otherAttrs = allAttrs.Where(a => !a.IsNamespaceDeclaration)
                .OrderBy(a => a.Name.NamespaceName, StringComparer.Ordinal)
                .ThenBy(a => a.Name.LocalName, StringComparer.Ordinal)
                .ToList();
            foreach (var a in nsAttrs) root.Add(a);
            foreach (var a in otherAttrs) root.Add(a);
            return root.ToString(System.Xml.Linq.SaveOptions.DisableFormatting);
        }
        catch
        {
            // Malformed XML — leave as-is rather than corrupting.
            return xml;
        }
    }

    // <w:footnotePr>/<w:endnotePr> in settings.xml carry <w:footnote>/<w:endnote>
    // child refs pointing at the separator + continuationSeparator notes
    // (typically id="-1" and id="0") that live in footnotes.xml / endnotes.xml.
    // The dump round-trips note CONTENT via typed `add footnote`/`add endnote`
    // (body-referenced notes only) — it never recreates a separator-only notes
    // part. So when a source carries those separator refs but no body-referenced
    // notes (a footnotes.xml holding ONLY the id -1/0 separators, which every
    // Word doc has), replaying the settings raw-set leaves the refs pointing at
    // notes parts that don't exist in the blank target, and the referential
    // validator rejects the result ("w:footnote … does not exist in part
    // /MainDocumentPart/FootnotesPart"). Word auto-manages separators, so
    // dropping these refs is lossless; strip them while keeping footnotePr/
    // endnotePr and any real config children (pos, numFmt, numStart, …).
    //
    // ALSO strips any settings element carrying a relationship reference
    // (an attribute in the `r:` namespace — r:id / r:embed / r:link). The dump
    // never recreates settings.xml.rels, so e.g. <w:attachedTemplate r:id="…"/>
    // (the pointer to Normal.dotm that nearly every real Word document carries)
    // dangles on replay — the OOXML validator NRE'd "before producing results"
    // and real Word refused to open the file. Dropping the pointer is lossless
    // (Word falls back to the Normal template). Same family as <w:mailMerge>'s
    // data-source r:id etc.; remove the whole referencing element.
    private static string StripDanglingNoteSeparatorRefs(
        string settingsXml, bool keepFootnoteSeps, bool keepEndnoteSeps)
    {
        if (string.IsNullOrEmpty(settingsXml) || !settingsXml.StartsWith("<")) return settingsXml;
        // Fast path: nothing to strip unless a note-properties block or a
        // relationship-bearing element (r: prefixed attribute) is present.
        if (!settingsXml.Contains("notePr") && !settingsXml.Contains(":id=")
            && !settingsXml.Contains(":embed=") && !settingsXml.Contains(":link="))
            return settingsXml;
        try
        {
            var doc = System.Xml.Linq.XDocument.Parse(settingsXml);
            if (doc.Root == null) return settingsXml;
            var wNs = (System.Xml.Linq.XNamespace)"http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            var rNs = (System.Xml.Linq.XNamespace)"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            var removed = false;
            // BUG-DUMP-R57-NOTESEP: keep the separator (-1) / continuationSeparator
            // (0) refs when the document has body-referenced notes — then `add
            // footnote`/`add endnote` recreates the notes part WITH those two
            // separators, so the refs resolve. Dropping them made Word fall back
            // to the DEFAULT separator, whose height differs from the source's
            // custom one; on a footnote-dense page that shifted the body text area
            // enough to flip a page break and cascade a multi-page reflow.
            //
            // BUG-DUMP-R58-NOTENOTICE: keep ONLY ids -1 and 0. footnotePr may also
            // reference other reserved special notes (continuationNotice, often
            // id=1) which the dump does NOT round-trip — and `add footnote`
            // renumbers the surviving body notes from 1 up, so the dropped
            // continuationNotice's id gets reused by a real BODY note. A kept ref
            // to that id then declares a body footnote as a document-wide special
            // note, which Word rejects outright ("file may be corrupted") even
            // though the SDK validator passes. So a kept ref is safe only for the
            // -1/0 separators the rebuild reliably recreates; strip every other
            // referenced id (and strip all refs when the part won't be recreated).
            foreach (var pr in doc.Descendants(wNs + "footnotePr").ToList())
            {
                foreach (var sep in pr.Elements(wNs + "footnote").ToList())
                {
                    var id = sep.Attribute(wNs + "id")?.Value;
                    if (keepFootnoteSeps && (id == "-1" || id == "0")) continue;
                    sep.Remove();
                    removed = true;
                }
            }
            foreach (var pr in doc.Descendants(wNs + "endnotePr").ToList())
            {
                foreach (var sep in pr.Elements(wNs + "endnote").ToList())
                {
                    var id = sep.Attribute(wNs + "id")?.Value;
                    if (keepEndnoteSeps && (id == "-1" || id == "0")) continue;
                    sep.Remove();
                    removed = true;
                }
            }
            // Drop any element with a dangling relationship reference (settings
            // .xml.rels is not round-tripped). attachedTemplate is the common one.
            foreach (var el in doc.Descendants().ToList())
            {
                if (el.Attributes().Any(a => a.Name.Namespace == rNs))
                {
                    el.Remove();
                    removed = true;
                }
            }
            if (!removed) return settingsXml;
            return doc.Root.ToString(System.Xml.Linq.SaveOptions.DisableFormatting);
        }
        catch
        {
            return settingsXml;
        }
    }

    // A footnotes/endnotes part that holds ONLY the reserved separator (-1) and
    // continuationSeparator (0) special notes is "default" — and droppable —
    // when each separator's paragraph carries nothing but the bare separator
    // glyph mark (<w:separator/> / <w:continuationSeparator/>). Word auto-manages
    // that default, so a blank target recreates an equivalent on open.
    //
    // But a separator can be CUSTOMIZED: real templates push a PAGE/NUMPAGES
    // field, "- N -" page-number text, or rule formatting into the separator
    // note's runs. Dropping such a part loses authored content silently. This
    // probe returns true when a notes part's separator/continuationSeparator
    // notes carry ANY run content beyond the bare glyph mark — the signal that
    // the whole part must be raw-emitted (and its settings footnotePr refs kept).
    //
    // Conservative by construction: only -1/0 special notes are inspected (body
    // notes are round-tripped via `add footnote`/`add endnote` regardless), and
    // any parse failure returns false (fall back to the existing drop path).
    private static bool HasCustomNoteSeparator(string notesXml)
    {
        if (string.IsNullOrEmpty(notesXml) || !notesXml.StartsWith("<")) return false;
        // Fast path: a default separator paragraph is just <w:separator/> /
        // <w:continuationSeparator/>; any field / instrText / drawn text — or a
        // <w:pPr> (paragraph spacing, see below) — means there is custom content
        // worth inspecting.
        if (!notesXml.Contains("fldChar") && !notesXml.Contains("instrText")
            && !notesXml.Contains("<w:t") && !notesXml.Contains("<w:drawing")
            && !notesXml.Contains("<w:pict") && !notesXml.Contains("<w:pPr"))
            return false;
        try
        {
            var doc = System.Xml.Linq.XDocument.Parse(notesXml);
            if (doc.Root == null) return false;
            var wNs = (System.Xml.Linq.XNamespace)"http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            foreach (var note in doc.Root.Elements())
            {
                if (note.Name.LocalName is not ("footnote" or "endnote")) continue;
                var type = note.Attribute(wNs + "type")?.Value;
                if (type is not ("separator" or "continuationSeparator")) continue;
                // Any run carrying a field char, field instruction, drawn text,
                // or drawing/picture is custom content the bare glyph mark lacks.
                // BUG-DUMP-NOTESEP-SPACING: a non-empty <w:pPr> (most often
                // <w:spacing w:after="0"/>, which tightens the footnote area) is
                // ALSO custom — AddFootnote seeds a bare pPr-less separator, so
                // dropping it lets Word's default after-spacing grow the footnote
                // area and reflow the body. Round-trip the whole part to keep it.
                bool custom = note.Descendants(wNs + "fldChar").Any()
                    || note.Descendants(wNs + "instrText").Any()
                    || note.Descendants(wNs + "t").Any()
                    || note.Descendants(wNs + "drawing").Any()
                    || note.Descendants(wNs + "pict").Any()
                    || note.Descendants(wNs + "pPr").Any(p => p.HasElements);
                if (custom) return true;
            }
            return false;
        }
        catch { return false; }
    }

    // Raw-emit word/footnotes.xml / word/endnotes.xml as a whole-part replace
    // when the source carries a CUSTOMIZED separator (HasCustomNoteSeparator)
    // but NO body-referenced notes — the only case the `add footnote`/`add
    // endnote` path does not already recreate the part. The settings raw-set
    // keeps the -1/0 footnotePr refs in this case (see EmitSettingsRaw), so the
    // refs resolve against the part we recreate here. A doc with body notes
    // recreates the part through Add (separators included) and skips this; a doc
    // with a plain default separator drops the part as before.
    //
    // Apply side: `raw-set /footnotes` / `/endnotes` create-or-replace the part
    // (WordHandler.RawSet, mirroring the /numbering and /theme branches).
    private static string SafeRaw(WordHandler word, string zipUri)
    {
        try { return word.Raw(zipUri); }
        catch { return ""; }
    }

    private static void EmitNoteSeparatorsRaw(WordHandler word, List<BatchItem> items)
    {
        EmitOneNoteSeparatorRaw(word, items, "footnote", "/word/footnotes.xml",
            "/footnotes", "/w:footnotes");
        EmitOneNoteSeparatorRaw(word, items, "endnote", "/word/endnotes.xml",
            "/endnotes", "/w:endnotes");
    }

    private static void EmitOneNoteSeparatorRaw(
        WordHandler word, List<BatchItem> items, string queryKind,
        string zipUri, string semanticPart, string rootXpath)
    {
        // Body notes recreate the part via Add — skip (Query filters -1/0).
        bool hasBody = false;
        try { hasBody = word.Query(queryKind).Count > 0; } catch { }
        if (hasBody) return;

        string xml;
        try { xml = word.Raw(zipUri); }
        catch { return; } // source has no such part
        xml = CanonicalizeRawXml(xml);
        if (!HasCustomNoteSeparator(xml)) return; // plain default — drop as before

        items.Add(new BatchItem
        {
            Command = "raw-set",
            Part = semanticPart,
            Xpath = rootXpath,
            Action = "replace",
            Xml = xml
        });
    }

    // BUG-DUMP-NOTENOTICE-FIDELITY: the HAS-BODY-NOTES complement of
    // EmitNoteSeparatorsRaw. When a source has body footnotes/endnotes, the
    // `add footnote`/`add endnote` body walk recreates the notes part — but
    // only with DEFAULT separator (-1) and continuationSeparator (0) notes
    // carrying the bare glyph mark. Two losses follow:
    //   1. A CUSTOMIZED separator/continuationSeparator (real "[Footnote
    //      continued …]" text, a PAGE field, a rule) is replaced by the bare
    //      default — its authored content vanishes.
    //   2. A continuationNotice special note (Word's "[Footnote continued on
    //      next page]", typically id=1) is dropped entirely, and because
    //      AddFootnote renumbers body notes from 1 up, that id gets REUSED by a
    //      real body note. EmitSettingsRaw's R58 strip drops the now-dangling
    //      ref to avoid the "file may be corrupted" Word reports for it.
    //
    // This fixup restores both with full fidelity, running AFTER the body walk
    // (so the part already exists and the body-note id range is known):
    //   - For -1 / 0: a targeted raw-set REPLACE swaps the seeded default note
    //     for the source's verbatim one (only when the source note is custom).
    //   - For continuationNotice (and any other reserved special id beyond
    //     -1/0): re-id it to a FRESH id above the rebuilt body range
    //     (max body id + 1) so it cannot collide with a body note, append it,
    //     and re-add the settings footnotePr ref at the new id via a targeted
    //     raw-set on <w:footnotePr> (overriding the R58 strip).
    //
    // Not a whole-part replace (that would clobber the body notes Add just
    // created). Mirrors EmitNoteSeparatorsRaw's conservatism: parse failure or
    // a non-custom source falls back to the existing (lossy) default path.
    private static void EmitNoteSpecialNotesFixup(WordHandler word, List<BatchItem> items)
    {
        EmitOneNoteSpecialNotesFixup(word, items, "footnote", "/word/footnotes.xml",
            "/footnotes", "footnotePr");
        EmitOneNoteSpecialNotesFixup(word, items, "endnote", "/word/endnotes.xml",
            "/endnotes", "endnotePr");
    }

    private static void EmitOneNoteSpecialNotesFixup(
        WordHandler word, List<BatchItem> items, string queryKind,
        string zipUri, string semanticPart, string settingsNotePr)
    {
        // Only the has-body-notes case — the no-body case is EmitNoteSeparatorsRaw.
        int bodyNoteCount;
        try { bodyNoteCount = word.Query(queryKind).Count; }
        catch { return; }
        if (bodyNoteCount == 0) return;

        string xml;
        try { xml = word.Raw(zipUri); }
        catch { return; }
        if (string.IsNullOrEmpty(xml) || !xml.StartsWith("<")) return;
        xml = CanonicalizeRawXml(xml);

        // BUG-DUMP-NOTE-RAWREF-WONTOPEN: the per-reference `add footnote`/`add
        // endnote` body walk only fires for references the walk actually visits.
        // When EVERY reference to a body note lives inside a raw-emitted region
        // — an SDT content-control carrier, a verbatim field/textbox block —
        // the walk never sees it, so NO `add <kind>` is emitted and the rebuild
        // never creates the FootnotesPart/EndnotesPart. The raw region still
        // carries the verbatim `<w:footnoteReference w:id="N"/>`, so the rebuilt
        // body references a note id that does not exist → Word reports the file
        // is corrupt and refuses to open. (The targeted per-note raw-sets below
        // would also fail: their child XPath matches nothing in the missing
        // part.) Recover by emitting the WHOLE notes part verbatim: RawSet on
        // "/w:footnotes" lazily creates the part with the source's exact note
        // bodies. Because nothing was added, no body renumbering happened, so
        // every raw reference keeps its original id and resolves cleanly.
        int emittedAdds = items.Count(it =>
            it.Command == "add"
            && string.Equals(it.Type, queryKind, StringComparison.OrdinalIgnoreCase));
        if (emittedAdds == 0)
        {
            bool rawHasRef = items.Any(it =>
                it.Command == "raw-set"
                && it.Xml != null
                && it.Xml.Contains($"{queryKind}Reference", StringComparison.Ordinal));
            if (rawHasRef)
            {
                items.Add(new BatchItem
                {
                    Command = "raw-set",
                    Part = semanticPart,
                    Xpath = $"/w:{queryKind}s",
                    Action = "replace",
                    Xml = xml
                });
            }
            // No `add <kind>` AND no raw reference → genuine orphan note bodies;
            // WarnOrphanNotes already surfaced the drop. Either way the targeted
            // per-note fixup below cannot run (no part to patch), so stop here.
            return;
        }

        System.Xml.Linq.XDocument doc;
        try { doc = System.Xml.Linq.XDocument.Parse(xml); }
        catch { return; }
        if (doc.Root == null) return;
        var wNs = (System.Xml.Linq.XNamespace)"http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        // A note carries CUSTOM content when it holds drawn text / a field /
        // a drawing beyond the bare separator glyph mark — same probe shape as
        // HasCustomNoteSeparator, applied per note.
        //
        // BUG-DUMP-NOTESEP-SPACING: ALSO custom when the separator paragraph
        // carries non-empty paragraph properties (a <w:pPr> with any child —
        // most often <w:spacing w:after="0"/>, which tightens the footnote
        // area). AddFootnote seeds a BARE separator paragraph (no pPr), so Word
        // applies its DEFAULT after-spacing on replay → the footnote area grows
        // taller → body text shifts → a ±1 page reflow across the document.
        // Treating a spacing-only separator as "custom" makes the verbatim
        // raw-set restore below fire and round-trip the pPr. (The dominant
        // common cause of the P1 footnote-reflow cluster.)
        static bool IsCustom(System.Xml.Linq.XElement note, System.Xml.Linq.XNamespace w)
            => note.Descendants(w + "t").Any()
            || note.Descendants(w + "fldChar").Any()
            || note.Descendants(w + "instrText").Any()
            || note.Descendants(w + "drawing").Any()
            || note.Descendants(w + "pict").Any()
            || note.Descendants(w + "pPr").Any(p => p.HasElements);

        // The rebuild renumbers body notes 1..bodyNoteCount, so the next free
        // id sits at bodyNoteCount+1. Re-id every restored continuationNotice
        // upward from there (multiple are possible in principle).
        int nextFreeId = bodyNoteCount + 1;
        // continuationNotice ids to re-add to settings footnotePr (new ids).
        var noticeRefIds = new List<int>();
        bool anyFixup = false;

        foreach (var note in doc.Root.Elements())
        {
            if (note.Name.LocalName is not ("footnote" or "endnote")) continue;
            var idStr = note.Attribute(wNs + "id")?.Value;
            if (!int.TryParse(idStr, out var id)) continue;
            var type = note.Attribute(wNs + "type")?.Value;

            if (id == -1 || id == 0)
            {
                // separator / continuationSeparator — only restore when custom;
                // a bare default note already matches what AddFootnote seeded.
                if (type is not ("separator" or "continuationSeparator")) continue;
                if (!IsCustom(note, wNs)) continue;
                items.Add(new BatchItem
                {
                    Command = "raw-set",
                    Part = semanticPart,
                    Xpath = $"/w:{queryKind}s/w:{queryKind}[@w:id='{id}']",
                    Action = "replace",
                    Xml = note.ToString(System.Xml.Linq.SaveOptions.DisableFormatting)
                });
                anyFixup = true;
            }
            else if (id > 0 && type is "continuationNotice")
            {
                // continuationNotice (or any reserved special note with a
                // positive id) — only the source body refs use the 2..N range;
                // a special note never has a body reference, so re-id it to a
                // fresh id above the rebuilt body range and append it.
                int freshId = nextFreeId++;
                note.SetAttributeValue(wNs + "id", freshId.ToString());
                items.Add(new BatchItem
                {
                    Command = "raw-set",
                    Part = semanticPart,
                    Xpath = $"/w:{queryKind}s",
                    Action = "append",
                    Xml = note.ToString(System.Xml.Linq.SaveOptions.DisableFormatting)
                });
                noticeRefIds.Add(freshId);
                anyFixup = true;
            }
        }

        if (!anyFixup) return;

        // Re-add the continuationNotice ref(s) to settings footnotePr at their
        // fresh ids. EmitSettingsRaw kept only -1/0 (R58 strip); append the
        // remapped notice refs after the kept separators via a targeted replace
        // of the whole <w:footnotePr> block. Build it from the kept -1/0 refs
        // plus the new notice refs so the order stays separator → contSep →
        // notice (the order Word writes). When the source settings had no
        // footnotePr the rebuild's blank one is empty — still emit, so the
        // notice ref resolves against the appended note.
        if (noticeRefIds.Count == 0) return;
        var sb = new System.Text.StringBuilder();
        sb.Append("<w:").Append(settingsNotePr)
          .Append(" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">");
        sb.Append($"<w:{queryKind} w:id=\"-1\"/>");
        sb.Append($"<w:{queryKind} w:id=\"0\"/>");
        foreach (var rid in noticeRefIds)
            sb.Append($"<w:{queryKind} w:id=\"{rid}\"/>");
        sb.Append("</w:").Append(settingsNotePr).Append('>');
        items.Add(new BatchItem
        {
            Command = "raw-set",
            Part = "/settings",
            Xpath = $"/w:settings/w:{settingsNotePr}",
            Action = "replace",
            Xml = sb.ToString()
        });
    }

    // <w:numPicBullet> defines a picture (image) list bullet; a level opts into
    // it with <w:lvlPicBulletId>. The picture lives in word/media/* referenced
    // by numbering.xml.rels (r:id inside the numPicBullet's VML/drawing). The
    // dump round-trips numbering.xml verbatim via raw-set but never recreates
    // the numbering part's rels or the media binary, so the r:id dangles on
    // replay — real Word then refuses to open the file ("may be corrupt") and
    // the SDK validator NREs walking the broken numPicBullet. Strip the
    // numPicBullet definitions AND the lvlPicBulletId opt-ins so the level
    // falls back to its own <w:lvlText> glyph (already round-tripped). Lossy
    // for picture bullets only; mirrors the dangling footnote/endnote separator
    // ref strip and the external-rel SDT fallback.
    private static string StripDanglingPicBullets(string numberingXml, out bool stripped)
    {
        stripped = false;
        if (string.IsNullOrEmpty(numberingXml) || !numberingXml.StartsWith("<")) return numberingXml;
        if (!numberingXml.Contains("PicBullet")) return numberingXml; // fast path
        try
        {
            var doc = System.Xml.Linq.XDocument.Parse(numberingXml);
            if (doc.Root == null) return numberingXml;
            var wNs = (System.Xml.Linq.XNamespace)"http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            var removed = false;
            foreach (var el in doc.Descendants(wNs + "numPicBullet")
                         .Concat(doc.Descendants(wNs + "lvlPicBulletId")).ToList())
            {
                el.Remove();
                removed = true;
            }
            if (!removed) return numberingXml;
            stripped = true;
            return doc.Root.ToString(System.Xml.Linq.SaveOptions.DisableFormatting);
        }
        catch
        {
            return numberingXml;
        }
    }

    // Round-trip the source's <w:docDefaults> (the document-wide rPr/pPr
    // baseline inside styles.xml) VERBATIM via raw-set replace. This is the
    // root fix for "blank-default pollution": BlankDocCreator stamps an
    // opinionated docDefaults (Calibri, sz=22/11pt, szCs=22) that a source
    // omitting a slot — calibre/pandoc exports routinely carry only szCs, or
    // only a complex-script font, leaving the Latin size/lang/textAlignment to
    // Word's application default — would otherwise inherit on replay, rendering
    // at the wrong size/font. Per-property emits (docDefaults.font.latin,
    // docDefaults.fontSize, …) only covered the slots the source set
    // EXPLICITLY and could not express "this slot is absent", so the blank's
    // value leaked through. Replacing the whole block makes the rebuilt
    // docDefaults byte-identical to the source — including its absences — so
    // Word applies the same defaults to both. Mirrors the theme/settings/
    // numbering raw-emit rationale (structured XML edited as a block).
    // Raw-set replace can leave extra namespace declarations (xmlns:w14,
    // xmlns:mc) attached to the replaced element itself even when nothing in
    // the subtree uses them, so a dump taken after one replay emits a
    // byte-different fragment than the original dump. Strip declarations whose
    // namespace is unused by any element or attribute in the subtree before
    // serializing, so repeated dump→replay→dump cycles stay byte-identical.
    private static System.Xml.Linq.XElement StripUnusedNsDeclarations(System.Xml.Linq.XElement el)
    {
        var used = new HashSet<System.Xml.Linq.XNamespace>();
        foreach (var d in el.DescendantsAndSelf())
        {
            used.Add(d.Name.Namespace);
            foreach (var a in d.Attributes())
                if (!a.IsNamespaceDeclaration && a.Name.Namespace != System.Xml.Linq.XNamespace.None)
                    used.Add(a.Name.Namespace);
        }
        foreach (var d in el.DescendantsAndSelf())
            d.Attributes()
                .Where(a => a.IsNamespaceDeclaration && !used.Contains((System.Xml.Linq.XNamespace)a.Value))
                .ToList()
                .ForEach(a => a.Remove());
        return el;
    }

    private static void EmitDocDefaultsRaw(WordHandler word, List<BatchItem> items)
    {
        string stylesXml;
        try { stylesXml = word.Raw("/styles"); }
        catch { return; }
        if (string.IsNullOrEmpty(stylesXml) || !stylesXml.StartsWith("<")) return;
        string? dd;
        try
        {
            var doc = System.Xml.Linq.XDocument.Parse(stylesXml);
            var wNs = (System.Xml.Linq.XNamespace)"http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            var el = doc.Root?.Element(wNs + "docDefaults");
            dd = el == null ? null : StripUnusedNsDeclarations(el).ToString(System.Xml.Linq.SaveOptions.DisableFormatting);
        }
        catch { return; }

        // BUG-DUMP-R32-2: the dump→batch rebuild starts from the blank-`create`
        // template, whose styles.xml DOES carry an opinionated <w:docDefaults>
        // (Calibri, sz=22/11pt). When the SOURCE styles.xml has NO docDefaults,
        // the template's leaked through unchanged — its sz=22 default compressed
        // line height enough to reflow a table across the page break (SSIM
        // 0.816 → 0.899 once stripped). Faithfully round-trip the source's
        // docDefaults-presence: emit a `remove` so the rebuilt styles.xml has no
        // docDefaults either, matching the source.
        if (string.IsNullOrEmpty(dd))
        {
            items.Add(new BatchItem
            {
                Command = "raw-set",
                Part = "/styles",
                Xpath = "/w:styles/w:docDefaults",
                Action = "remove",
            });
            return;
        }

        items.Add(new BatchItem
        {
            Command = "raw-set",
            Part = "/styles",
            Xpath = "/w:styles/w:docDefaults",
            Action = "replace",
            Xml = dd
        });
    }

    // <w:latentStyles> (the built-in style visibility/priority table Word
    // writes on every authored document) was dropped on dump: the blank
    // rebuild has none and nothing re-emitted it. Mostly UI metadata, but its
    // defaults (defSemiHidden/defUIPriority and per-style lsdExceptions)
    // change how Word surfaces styles. Round-trip verbatim: insert after the
    // docDefaults block (CT_Styles order: docDefaults, latentStyles, style*),
    // or prepend when the source has no docDefaults.
    private static void EmitLatentStylesRaw(WordHandler word, List<BatchItem> items)
    {
        string stylesXml;
        try { stylesXml = word.Raw("/styles"); }
        catch { return; }
        if (string.IsNullOrEmpty(stylesXml) || !stylesXml.StartsWith("<")) return;
        string? ls;
        bool hasDocDefaults;
        try
        {
            var doc = System.Xml.Linq.XDocument.Parse(stylesXml);
            var wNs = (System.Xml.Linq.XNamespace)"http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            var lsEl = doc.Root?.Element(wNs + "latentStyles");
            ls = lsEl == null ? null : StripUnusedNsDeclarations(lsEl).ToString(System.Xml.Linq.SaveOptions.DisableFormatting);
            hasDocDefaults = doc.Root?.Element(wNs + "docDefaults") != null;
        }
        catch { return; }
        if (string.IsNullOrEmpty(ls)) return; // source had none; blank has none — consistent

        items.Add(new BatchItem
        {
            Command = "raw-set",
            Part = "/styles",
            Xpath = hasDocDefaults ? "/w:styles/w:docDefaults" : "/w:styles",
            Action = hasDocDefaults ? "insertafter" : "prepend",
            Xml = ls
        });
    }

    // BUG-R4B(BUG6): the theme part (word/theme/theme1.xml) can carry a
    // <a:blipFill><a:blip r:embed="rIdN"/> referencing an image relationship in
    // theme1.xml.rels (a custom fmtScheme bg fill). The dump round-trips the
    // theme XML verbatim via raw-set but never recreates the theme part's rels
    // or the media binary, so the r:embed dangles on replay and the rebuilt
    // theme1.xml fails validation ("relationship 'rId1' ... does not exist").
    // Theme-image round-trip is a separate feature; here we just ensure the
    // rebuilt theme has NO dangling reference: strip the unreconstructable
    // blip/blipFill cleanly (falling back to the previous fill in the same
    // *StyleLst, or dropping the entry) and signal the loss via a warning.
    // Mirrors StripDanglingNoteSeparatorRefs / StripDanglingPicBullets.
    private static string StripDanglingThemeBlipRefs(string themeXml, out bool stripped)
    {
        stripped = false;
        if (string.IsNullOrEmpty(themeXml) || !themeXml.StartsWith("<")) return themeXml;
        // Fast path: only act when a relationship-bearing attribute is present.
        if (!themeXml.Contains(":embed=") && !themeXml.Contains(":link="))
            return themeXml;
        try
        {
            var doc = System.Xml.Linq.XDocument.Parse(themeXml);
            if (doc.Root == null) return themeXml;
            var aNs = (System.Xml.Linq.XNamespace)"http://schemas.openxmlformats.org/drawingml/2006/main";
            var rNs = (System.Xml.Linq.XNamespace)"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            var removed = false;
            // Any blipFill whose blip carries a relationship reference cannot be
            // reconstructed (image binary + rels not round-tripped). Remove the
            // whole <a:blipFill> so the parent fill list entry is replaced with a
            // schema-valid neutral fill rather than a dangling one.
            foreach (var blipFill in doc.Descendants(aNs + "blipFill").ToList())
            {
                var blip = blipFill.Element(aNs + "blip");
                var hasRel = blip != null
                    && blip.Attributes().Any(a => a.Name.Namespace == rNs);
                if (!hasRel) continue;
                // Replace the unreconstructable blipFill with a neutral
                // <a:solidFill><a:schemeClr val="phClr"/></a:solidFill> — the
                // canonical placeholder fill used throughout fmtScheme style
                // lists — so the surrounding *StyleLst keeps a valid child count
                // and Word still renders a (plain) fill.
                var placeholder = new System.Xml.Linq.XElement(aNs + "solidFill",
                    new System.Xml.Linq.XElement(aNs + "schemeClr",
                        new System.Xml.Linq.XAttribute("val", "phClr")));
                blipFill.ReplaceWith(placeholder);
                removed = true;
            }
            // Defensive: drop any other element still carrying an r:embed/r:link
            // (e.g. an a:blip that is not inside an a:blipFill).
            foreach (var el in doc.Descendants().ToList())
            {
                if (el.Attributes().Any(a => a.Name.Namespace == rNs
                    && (a.Name.LocalName == "embed" || a.Name.LocalName == "link")))
                {
                    el.Remove();
                    removed = true;
                }
            }
            if (!removed) return themeXml;
            stripped = true;
            return doc.Root.ToString(System.Xml.Linq.SaveOptions.DisableFormatting);
        }
        catch
        {
            return themeXml;
        }
    }

    private static void EmitThemeRaw(WordHandler word, List<BatchItem> items,
                                     List<DocxUnsupportedWarning>? warnings = null)
    {
        // Theme carries clrScheme + fontScheme + fmtScheme — pure structured
        // XML that users rarely modify property-by-property; the natural
        // operation is "swap the entire theme block". Raw-set replace fits
        // that model exactly. Word.Raw returns the literal string
        // "(no theme)" when the part is missing.
        //
        // ALWAYS emit, even for source docs that have no theme part. The
        // blank target auto-stamps theme1.xml (for Word render
        // parity), so silently skipping the emit caused dump∘replay∘dump
        // to drift by +1 item every pass: dump-1 saw no theme and
        // emitted nothing; replay left blank's theme in place; dump-2
        // saw blank's theme and emitted it. Dump-1 now emits an empty
        // <a:theme/> placeholder for theme-less sources, which the apply
        // path overwrites blank's seeded theme with — making dump-2 see
        // the same empty theme and emit the same placeholder. Fixed point.
        string xml;
        try { xml = word.Raw("/theme"); }
        catch { xml = ""; }
        // BUG-DUMP-R37-5: a theme-LESS source (word.Raw returns the literal
        // "(no theme)" sentinel) must round-trip with NO theme part. The blank
        // rebuild template auto-stamps a theme1.xml whose minorFont resolves CJK
        // glyphs differently from Word's app-default theme, tightening CJK line
        // height and drifting body text upward. Faithfully round-trip the
        // source's theme-absence: emit a `remove` so the rebuilt doc has no
        // theme part either (RawSet /theme + action=remove deletes the part and
        // its document.xml.rels relationship — no dangling ref). Mirrors
        // EmitDocDefaultsRaw's remove branch for a source lacking docDefaults.
        if (string.Equals(xml.Trim(), "(no theme)", StringComparison.Ordinal))
        {
            // The "(no theme)" sentinel fires for TWO distinct source shapes:
            // (a) the theme PART is genuinely absent — round-trip the absence
            //     with a remove (BUG-DUMP-R37-5, below); and
            // (b) the part EXISTS but is degenerate (0-byte / unreadable root —
            //     ThemePart.Theme is null), which Word tolerates. Emitting the
            //     remove here deleted the rebuilt doc's theme part outright;
            //     fall through instead so the schema-complete default-theme
            //     branch below emits a replace, mirroring what the source doc
            //     effectively renders with.
            // CONSISTENCY(empty-theme-default): same default-theme reuse as the
            // no-themeElements branch below (BlankDocCreator.BuildDefaultTheme).
            bool themePartExists = false;
            try
            {
                themePartExists = word.EnumeratePartUris().Any(u =>
                    u.StartsWith("/word/theme/", StringComparison.OrdinalIgnoreCase)
                    && !u.EndsWith(".rels", StringComparison.OrdinalIgnoreCase));
            }
            catch { /* enumeration failed — treat as absent (prior behavior) */ }
            if (!themePartExists)
            {
                items.Add(new BatchItem
                {
                    Command = "raw-set",
                    Part = "/theme",
                    Xpath = "/a:theme",
                    Action = "remove",
                });
                return;
            }
            xml = ""; // degenerate part → schema-complete default theme below
        }
        xml = CanonicalizeRawXml(xml);
        // A bare <a:theme/> (or <a:theme name="Office Theme"/>) is schema-INVALID:
        // <a:theme> requires a child <a:themeElements> (clrScheme + fontScheme +
        // fmtScheme). Replaying that placeholder over the blank target's valid
        // theme1.xml produced a file real Word refuses to open ("file may be
        // corrupt"); the source docx that triggered this carried a 0-byte
        // theme1.xml, which Word tolerates but the SDK read back as an empty
        // theme. Emit the SAME complete theme a blank doc stamps instead: the
        // result is Word-openable AND keeps the dump→replay→dump item count
        // stable (replay writes a real theme, dump-2 reads it back and emits
        // one theme item, same as dump-1) — the original reason this site
        // always emits something rather than skipping theme-less sources.
        if (string.IsNullOrEmpty(xml) || !xml.StartsWith("<") || !xml.Contains("themeElements"))
            xml = BlankDocCreator.BuildDefaultTheme(null, null).OuterXml;

        // BUG-R4B(BUG6): scrub dangling theme image references (the image binary
        // and theme1.xml.rels are not round-tripped) so the rebuilt theme
        // validates clean. Warn so the loss is visible.
        xml = StripDanglingThemeBlipRefs(xml, out var blipStripped);
        if (blipStripped)
            warnings?.Add(new DocxUnsupportedWarning(
                Element: "theme.blipFill",
                Path: "/theme",
                Reason: "theme image fill (a:blipFill r:embed) dropped — theme-part image round-trip is not supported; the fill was replaced with a neutral placeholder to keep the rebuilt theme valid"));

        items.Add(new BatchItem
        {
            Command = "raw-set",
            Part = "/theme",
            Xpath = "/a:theme",
            Action = "replace",
            Xml = xml
        });
    }

    private static void EmitSettingsRaw(WordHandler word, List<BatchItem> items)
    {
        // Settings carries dozens of feature flags + compat shims that
        // surface on root.Format only piecemeal — and not all of them are
        // wired through Set's case table. Wholesale raw-set is the simplest
        // way to keep Word feature toggles (evenAndOddHeaders, mirrorMargins,
        // schema-pegged compat options, …) round-tripped without
        // per-property allowlisting.
        //
        // ALWAYS emit, even for source docs without a settings part. The
        // blank target auto-stamps a settings.xml (characterSpacingControl
        // + compat block), so silently skipping the emit caused the same
        // idempotency drift as EmitThemeRaw: dump-1 saw no settings and
        // emitted nothing, dump-2 saw blank's leftover and emitted it.
        // Empty placeholder clears blank's seeded settings so dump-2
        // reads the same empty state and emits the same placeholder.
        string xml;
        try { xml = word.Raw("/settings"); }
        catch { xml = ""; }
        xml = CanonicalizeRawXml(xml);
        if (string.IsNullOrEmpty(xml) || !xml.StartsWith("<"))
            xml = "<w:settings xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" />";
        // BUG-DUMP-R57-NOTESEP: body-referenced notes mean the corresponding
        // notes part (with its -1/0 separators) is recreated on replay, so the
        // settings separator refs resolve and must be preserved (dropping them
        // changed the rendered separator height and reflowed footnote-dense
        // pages). Query filters out the reserved -1/0 separators, so a non-empty
        // result means real body notes exist.
        bool hasBodyFootnotes = false, hasBodyEndnotes = false;
        try { hasBodyFootnotes = word.Query("footnote").Count > 0; } catch { }
        try { hasBodyEndnotes = word.Query("endnote").Count > 0; } catch { }
        // BUG-DUMP-NOTESEP-CUSTOM: a separator-only notes part is recreated by
        // EmitNoteSeparatorsRaw when its separator is CUSTOMIZED (PAGE field,
        // "- N -" text, …). In that case the -1/0 footnotePr refs must survive
        // the strip too, so they resolve against the raw-emitted part — same
        // reasoning as the body-notes case, different recreation path.
        bool keepFootnoteSeps = hasBodyFootnotes
            || HasCustomNoteSeparator(SafeRaw(word, "/word/footnotes.xml"));
        bool keepEndnoteSeps = hasBodyEndnotes
            || HasCustomNoteSeparator(SafeRaw(word, "/word/endnotes.xml"));
        xml = StripDanglingNoteSeparatorRefs(xml, keepFootnoteSeps, keepEndnoteSeps);

        items.Add(new BatchItem
        {
            Command = "raw-set",
            Part = "/settings",
            Xpath = "/w:settings",
            Action = "replace",
            Xml = xml
        });
    }

    private static void EmitNumberingRaw(WordHandler word, List<BatchItem> items)
        => EmitNumberingRaw(word, items, null, RecursiveNumberingDecomp);

    private static void EmitNumberingRaw(WordHandler word, List<BatchItem> items, List<DocxUnsupportedWarning>? warnings)
        => EmitNumberingRaw(word, items, warnings, RecursiveNumberingDecomp);

    private static void EmitNumberingRaw(WordHandler word, List<BatchItem> items, List<DocxUnsupportedWarning>? warnings, bool recursiveDecomp)
    {
        // Numbering models list templates (abstractNum + num pairs, each
        // abstractNum holds 9 levels with their own pPr / numFmt / lvlText).
        // RECURSIVE-NUMBERING-DECOMP (default ON): decompose the whole
        // <w:numbering> subtree into typed `add` ops — one `add w:abstractNum` /
        // `add w:num` per definition, recursing into every child (multiLevelType,
        // lvl, start, numFmt, lvlText, pPr, rPr, abstractNumId, lvlOverride, …)
        // via the generic prefixed-add path (the same machinery that decomposes
        // styles). Falls back to the verbatim raw-set replace below on ANY
        // residue (picture bullets carry a VML r:id; external rels / unknown
        // namespaces) — that path already round-trips the pic-bullet binaries.
        // The blank document creates an empty numbering part, so a single
        // replace on the part root is sufficient for the fallback.
        string xml;
        try { xml = word.Raw("/numbering"); }
        catch { return; }
        xml = CanonicalizeRawXml(xml);
        if (string.IsNullOrEmpty(xml) || !xml.StartsWith("<")) return;
        // Skip when numbering is empty (just `<w:numbering/>` with no children).
        if (!xml.Contains("<w:abstractNum") && !xml.Contains("<w:num "))
            return;

        if (recursiveDecomp)
        {
            var typedOps = TryDecomposeNumbering(xml);
            if (typedOps != null)
            {
                items.AddRange(typedOps);
                return;
            }
        }
        // BUG-DUMP-R45-2: round-trip the picture-bullet image binaries (word/media/*)
        // so the <w:numPicBullet> definition + its <v:imagedata r:id> + each level's
        // <w:lvlPicBulletId> opt-in can STAY. Read the NumberingDefinitionsPart's
        // ImageParts keyed by the rel id the VML references; for each ref the
        // numbering XML still carries, emit a companion `embed-binary` op so the
        // apply side rebuilds the ImagePart with the SAME r:id (RawEmbedBinary).
        // Only fall back to StripDanglingPicBullets when NO image binary can be
        // carried (orphaned rel) — then the level falls back to its <w:lvlText>
        // glyph. A doc with numbering but no pic bullets is unchanged.
        var picBins = word.GetNumberingImageParts();
        var carriedPicRelIds = picBins
            .Where(b => xml.Contains($"r:id=\"{b.RelId}\"", StringComparison.Ordinal))
            .ToList();
        if (carriedPicRelIds.Count == 0)
        {
            // No image binary to back the numPicBullet refs — strip to stay valid.
            xml = StripDanglingPicBullets(xml, out var picBulletsStripped);
            if (picBulletsStripped)
                warnings?.Add(new DocxUnsupportedWarning(
                    Element: "numbering.numPicBullet",
                    Path: "/numbering",
                    Reason: "picture list bullet (w:numPicBullet / w:lvlPicBulletId) dropped because its image binary could not be read; affected levels fall back to their w:lvlText glyph"));
        }
        xml = ReorderLvlChildren(xml);

        items.Add(new BatchItem
        {
            Command = "raw-set",
            Part = "/numbering",
            Xpath = "/w:numbering",
            Action = "replace",
            Xml = xml
        });

        // Companion binary ops AFTER the XML replace, so the NumberingDefinitionsPart
        // (created lazily by the replace) exists when RawEmbedBinary attaches each
        // ImagePart. Skip refs the (possibly stripped) XML no longer carries.
        foreach (var (relId, bytes, contentType) in carriedPicRelIds)
        {
            if (!xml.Contains($"r:id=\"{relId}\"", StringComparison.Ordinal)) continue;
            items.Add(new BatchItem
            {
                Command = "raw-set",
                Part = "/numbering",
                Xpath = relId,
                Action = "embed-binary",
                Xml = $"data:{contentType};base64,{Convert.ToBase64String(bytes)}",
            });
        }
    }

    // BUG-DUMP-R42-3: round-trip word/fontTable.xml. The part declares font
    // faces (panose/charset/family/pitch/sig) plus <w:altName> substitutions
    // (e.g. 方正小标宋简体 altName=方正舒体). Dropping it makes Word substitute a
    // different face — a real visual change (a title rendered cursive/brush in
    // the source renders plain in the rebuild). Mirror the theme/numbering raw
    // round-trip: emit a `raw-set /fontTable` replace carrying the verbatim
    // <w:fonts> XML.
    //
    // IMPORTANT: fontTable.xml may reference embedded font binaries
    // (w:embedRegular/w:embedBold/w:embedItalic/w:embedBoldItalic, each
    // carrying an r:id → font/*.odttf). The .odttf binaries and the part's
    // rels are NOT round-tripped, so preserving those r:ids would dangle. Strip
    // the embed elements (keeping the face declarations + altName subs — the
    // rendering-relevant part) so the rebuilt part validates with no dangling
    // rel. A doc with NO fontTable emits nothing.
    // customXml data stores (item.xml + itemProps.xml): SDT content controls
    // bind to a store through the itemProps datastore-item GUID, so shipping
    // the part bytes verbatim (with the item recreated under its SOURCE rel id
    // so the props op can address it) restores the bindings. Previously the
    // whole /customXml tree was warn-dropped.
    private static void EmitCustomXmlRaw(WordHandler word, List<BatchItem> items)
    {
        foreach (var (relId, bytes, ct, props) in word.GetCustomXmlEmitData())
        {
            items.Add(new BatchItem
            {
                Command = "raw-set",
                Part = "/customXml",
                Xpath = relId,
                Action = "embed-binary",
                Xml = $"data:{ct};base64,{Convert.ToBase64String(bytes)}",
            });
            if (props is { } p)
            {
                items.Add(new BatchItem
                {
                    Command = "raw-set",
                    Part = $"/customXml/{relId}",
                    Xpath = p.RelId,
                    Action = "embed-binary",
                    Xml = $"data:{p.ContentType};base64,{Convert.ToBase64String(p.Bytes)}",
                });
            }
        }
    }

    // docProps/core.xml + app.xml + custom.xml — document properties.
    //
    // Cover pages and headers routinely host data-bound content controls
    // (`<w:sdt>` with `<w:dataBinding w:xpath="…">`) whose DISPLAYED text is
    // pulled from these property stores, not from the cached run text:
    //   • core.xml   dc:title / dc:subject / dc:creator …  (coreProperties)
    //   • app.xml    Company / Manager / TitlesOfParts …    (extended-properties)
    //   • custom.xml user-defined name/value pairs           (custom-properties)
    // Without round-tripping the stores, the blank rebuild stamps OfficeCLI
    // defaults (Application=OfficeCLI, creator=OfficeCLI, no Company/title),
    // so every bound control renders EMPTY — the cover title/company/contact
    // vanish even though the SDT structure round-trips perfectly.
    //
    // Emit each part verbatim as a normal `raw-set replace` whose xpath is the
    // part's root element (replacing the root IS replacing the whole part — no
    // bespoke action verb). The apply side recognises the docProps part path and
    // rewrites the whole zip entry after the package closes (the SDK won't
    // persist a mid-session docProps write); see WordHandler.StashWholePartReplace.
    // A dump→batch rebuild reproduces the source, so all three are carried
    // verbatim — the source authoring identity (app.xml Application, core.xml
    // creator, custom.xml user props) is the faithful result; the OfficeCLI
    // audit stamp is a create/edit concern, not a reconstruction one. Previously
    // these were treated as auto-managed (restamped to OfficeCLI defaults) and
    // silently dropped on dump, blanking every data-bound control.
    private static void EmitDocPropsRaw(WordHandler word, List<BatchItem> items)
    {
        foreach (var partUri in new[] { "/docProps/core.xml", "/docProps/app.xml", "/docProps/custom.xml" })
        {
            string xml;
            try { xml = word.Raw(partUri); }
            catch { continue; } // source lacks this part — nothing to carry
            if (string.IsNullOrWhiteSpace(xml) || !xml.TrimStart().StartsWith("<")) continue;
            xml = CanonicalizeRawXml(xml);
            if (string.IsNullOrEmpty(xml) || !xml.StartsWith("<")) continue;
            // The xpath is the part's root element — replacing the root element
            // IS replacing the whole part, so the standard `replace` action with
            // the root xpath is the honest description (no bespoke action verb).
            // The apply side recognises the docProps part path and rewrites the
            // whole entry; see WordHandler.RawSet / StashWholePartReplace.
            var rootXpath = "/" + RootElementName(xml);
            items.Add(new BatchItem
            {
                Command = "raw-set",
                Part = partUri,
                Xpath = rootXpath,
                Action = "replace",
                Xml = xml,
            });
        }
    }

    // Extract the (possibly prefixed) qualified name of the first element in an
    // XML string — used as the root xpath for whole-part docProps replaces.
    private static string RootElementName(string xml)
    {
        var i = xml.IndexOf('<');
        while (i >= 0 && i + 1 < xml.Length && (xml[i + 1] == '?' || xml[i + 1] == '!'))
            i = xml.IndexOf('<', i + 1);
        if (i < 0) return "*";
        int start = i + 1;
        int end = start;
        while (end < xml.Length && xml[end] != ' ' && xml[end] != '>' && xml[end] != '\t'
               && xml[end] != '\r' && xml[end] != '\n' && xml[end] != '/')
            end++;
        return end > start ? xml[start..end] : "*";
    }

    // word/webSettings.xml (web-publishing div/frame settings). Verbatim
    // whole-part raw-set; the apply side creates the part lazily. Previously
    // warn-dropped.
    private static void EmitWebSettingsRaw(WordHandler word, List<BatchItem> items)
    {
        string xml;
        try { xml = word.Raw("/webSettings"); }
        catch { return; }
        if (string.Equals(xml.Trim(), "(no webSettings)", StringComparison.Ordinal))
            return;
        xml = CanonicalizeRawXml(xml);
        if (string.IsNullOrEmpty(xml) || !xml.StartsWith("<")) return;
        items.Add(new BatchItem
        {
            Command = "raw-set",
            Part = "/webSettings",
            Xpath = "/w:webSettings",
            Action = "replace",
            Xml = xml
        });
    }

    private static void EmitFontTableRaw(WordHandler word, List<BatchItem> items,
                                         List<DocxUnsupportedWarning>? warnings = null)
    {
        string xml;
        try { xml = word.Raw("/fonttable"); }
        catch { return; }
        if (string.Equals(xml.Trim(), "(no fontTable)", StringComparison.Ordinal))
            return; // source had no fontTable — emit nothing
        xml = CanonicalizeRawXml(xml);
        if (string.IsNullOrEmpty(xml) || !xml.StartsWith("<") || !xml.Contains("<w:font"))
            return; // empty <w:fonts/> — nothing rendering-relevant to carry

        // BUG-DUMP-R45-1: round-trip the embedded font binaries (word/fonts/*.odttf)
        // so the <w:embed*> refs can stay. Read the FontParts keyed by the rel id
        // the embed elements reference; for each ref we can carry, emit a companion
        // `embed-binary` op so the apply side rebuilds the FontPart with the SAME
        // r:id (RawEmbedBinary). The .odttf bytes round-trip raw — the obfuscation
        // key (w:fontKey) already rides in the verbatim fontTable XML. Only fall
        // back to stripping refs whose binary could not be read (orphaned rel), so
        // the rebuilt part never carries a dangling embed ref. A fontTable that
        // declares NO embeds (faces only) emits nothing extra.
        var fontBins = word.GetEmbeddedFontParts();
        var carriedFontRelIds = new HashSet<string>(
            fontBins.Select(b => b.RelId), StringComparer.OrdinalIgnoreCase);

        // Strip ONLY the embed refs we cannot back with a binary; keep the rest.
        xml = StripUncarriedEmbeddedFontRefs(xml, carriedFontRelIds, out var embedStripped);
        if (embedStripped)
            warnings?.Add(new DocxUnsupportedWarning(
                Element: "fontTable.embeddedFont",
                Path: "/fontTable",
                Reason: "an embedded font binary reference (w:embedRegular/Bold/Italic/BoldItalic) was dropped because its font binary could not be read; the font-face declarations and altName substitutions are preserved"));

        items.Add(new BatchItem
        {
            Command = "raw-set",
            Part = "/fontTable",
            Xpath = "/w:fonts",
            Action = "replace",
            Xml = xml
        });

        // Companion binary ops AFTER the XML replace, so the FontTablePart (created
        // lazily by the replace) exists when RawEmbedBinary attaches each FontPart.
        // Skip any ref the (possibly stripped) XML no longer carries.
        foreach (var (relId, bytes, contentType) in fontBins)
        {
            if (!xml.Contains($"r:id=\"{relId}\"", StringComparison.Ordinal)) continue;
            items.Add(new BatchItem
            {
                Command = "raw-set",
                Part = "/fontTable",
                Xpath = relId,
                Action = "embed-binary",
                Xml = $"data:{contentType};base64,{Convert.ToBase64String(bytes)}",
            });
        }
    }

    // BUG-DUMP-R45-1: strip ONLY the <w:embed*> elements whose r:id is NOT in the
    // carried set (binary unreadable), keeping every ref we round-trip. Regex-level
    // (tolerant of attribute order / self-closing forms), matching the StripDangling*
    // family. When carriedRelIds covers all embeds this is a no-op.
    private static string StripUncarriedEmbeddedFontRefs(
        string xml, HashSet<string> carriedRelIds, out bool stripped)
    {
        var before = xml;
        xml = System.Text.RegularExpressions.Regex.Replace(
            xml,
            @"<w:embed(?:Regular|Bold|Italic|BoldItalic)\b[^>]*?/>",
            m =>
            {
                var idm = System.Text.RegularExpressions.Regex.Match(m.Value, @"r:id=""([^""]+)""");
                return idm.Success && carriedRelIds.Contains(idm.Groups[1].Value)
                    ? m.Value
                    : string.Empty;
            });
        stripped = !string.Equals(before, xml, StringComparison.Ordinal);
        return xml;
    }

    // BUG-DUMP-R28-3: a source <w:lvl> may store its children in an order that
    // is tolerated by Word but violates the CT_Lvl schema sequence — most
    // commonly <w:legacy> emitted BEFORE <w:suff>/<w:lvlText> (legacy list
    // templates from older Word exports). The dump round-trips numbering.xml
    // verbatim via raw-set, so the out-of-order children reach the rebuilt
    // part unchanged; the SDK validator then rejects the FIRST element that
    // appears after the schema state machine has advanced past its slot
    // (e.g. "<w:suff> unexpected" once <w:legacy> has been seen). Real Word is
    // lenient, but `validate` and strict consumers fail. Reorder each <w:lvl>'s
    // children into the canonical CT_Lvl sequence so the rebuilt numbering.xml
    // validates. Unknown/unlisted children keep their relative order and sort
    // after the known ones (defensive — CT_Lvl has no extension point, but a
    // future/vendor element shouldn't be dropped). Mirrors StripDanglingPicBullets'
    // parse-edit-reserialize shape.
    private static readonly string[] _ctLvlChildOrder =
    {
        "start", "numFmt", "lvlRestart", "pStyle", "isLgl", "suff",
        "lvlText", "lvlPicBulletId", "legacy", "lvlJc", "pPr", "rPr"
    };

    private static string ReorderLvlChildren(string numberingXml)
    {
        if (string.IsNullOrEmpty(numberingXml) || !numberingXml.StartsWith("<")) return numberingXml;
        if (!numberingXml.Contains("<w:lvl")) return numberingXml; // fast path
        try
        {
            var doc = System.Xml.Linq.XDocument.Parse(numberingXml);
            if (doc.Root == null) return numberingXml;
            var wNs = (System.Xml.Linq.XNamespace)"http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            int RankOf(System.Xml.Linq.XElement e)
            {
                if (e.Name.Namespace != wNs) return _ctLvlChildOrder.Length;
                int idx = Array.IndexOf(_ctLvlChildOrder, e.Name.LocalName);
                return idx < 0 ? _ctLvlChildOrder.Length : idx;
            }
            var changed = false;
            foreach (var lvl in doc.Descendants(wNs + "lvl").ToList())
            {
                var kids = lvl.Elements().ToList();
                if (kids.Count < 2) continue;
                // Stable sort by CT_Lvl rank; only rewrite when order differs.
                var sorted = kids
                    .Select((el, i) => (el, i))
                    .OrderBy(t => RankOf(t.el))
                    .ThenBy(t => t.i)
                    .Select(t => t.el)
                    .ToList();
                bool reordered = false;
                for (int i = 0; i < kids.Count; i++)
                {
                    if (!ReferenceEquals(kids[i], sorted[i])) { reordered = true; break; }
                }
                if (!reordered) continue;
                foreach (var k in kids) k.Remove();
                foreach (var s in sorted) lvl.Add(s);
                changed = true;
            }
            if (!changed) return numberingXml;
            return doc.Root.ToString(System.Xml.Linq.SaveOptions.DisableFormatting);
        }
        catch
        {
            return numberingXml;
        }
    }

    private static void EmitHeadersFooters(WordHandler word, List<BatchItem> items,
                                           List<DocxUnsupportedWarning>? warnings = null)
    {
        var root = word.Get("/");
        if (root.Children == null) return;
        // BUG-X4-T2: header/footer parts carry no `type` key on Get; the
        // section's `headerRef.default|first|even` (and `footerRef.*`)
        // entries are the only place the part's role is recorded. Build a
        // reverse lookup so EmitHeaderFooterPart can emit the right
        // `type` prop (default/first/even) instead of always emitting
        // "default" — which on a doc with both default + first headers
        // throws "Header of type 'default' already exists" on replay.
        // In addition to (path → type), track which section's headerRef /
        // footerRef points at the part. Multi-section docs with per-section
        // default headers used to all emit `add header parent="/"` —
        // AddHeader resolves "/" to a single sectPr, so the 2nd-and-later
        // default headers tripped "Header of type 'default' already exists"
        // on replay. Emit `parent=/section[N]` so each header targets its
        // true owning section (mirrors ResolveTargetSectPrForHeaderFooter's
        // /section[N] resolver).
        // A single header/footer PART may be referenced by MORE THAN ONE type
        // in the same section — Word commonly points both the `even` and the
        // `default` headerReference at one part (so odd AND even pages show the
        // same running header without authoring two copies). Keep the full LIST
        // of (type, section) refs per part, not just the first: collapsing to a
        // single ref dropped the `default` reference, so odd pages (which use
        // the default header when evenAndOddHeaders is off, or when there is no
        // titlePg) rendered with NO header at all. Each ref is emitted as its
        // own `add header` (a content copy referenced by that type) — the
        // rebuild carries N small part copies instead of one shared part, but
        // renders identically.
        var headerPathInfo = new Dictionary<string, List<(string Type, string? SectionPath)>>(StringComparer.OrdinalIgnoreCase);
        var footerPathInfo = new Dictionary<string, List<(string Type, string? SectionPath)>>(StringComparer.OrdinalIgnoreCase);
        var headerRefSeen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var footerRefSeen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        // headerRef.<type> / footerRef.<type> live on **section** nodes
        // (see WordHandler.Query.cs:902), not on root. An earlier fix
        // scanned root.Format and silently found nothing, so every emitted
        // header/footer was typed "default" — round-trip failed when a doc
        // had both default + first headers. Walk all section children to
        // build the path→type map.
        // Attribute each referenced header/footer part to the section whose
        // sectPr references it, using the REPLAY-side document-order ordinal.
        // EnumerateSectionHeaderFooterRefs walks every sectPr including inline
        // ones nested in an SDT (which `query section` misses) — required
        // because a dump unwraps an SDT-wrapped section into a normal body
        // paragraph, so its sectPr joins the /section[N] sequence and a header
        // attributed by the SDT-blind ordinal would land on the wrong section
        // (the symptom: a landscape figure section's header rendered on the
        // first portrait page and the real first-page header vanished).
        foreach (var sref in word.EnumerateSectionHeaderFooterRefs())
        {
            var parent = sref.IsFinal ? "/" : $"/section[{sref.ReplayOrdinal}]";
            foreach (var (type, partPath) in sref.Headers)
                if (headerRefSeen.Add($"{partPath}|{type}|{parent}"))
                    (headerPathInfo.TryGetValue(partPath, out var hl) ? hl
                        : (headerPathInfo[partPath] = new List<(string, string?)>()))
                        .Add((type, parent));
            foreach (var (type, partPath) in sref.Footers)
                if (footerRefSeen.Add($"{partPath}|{type}|{parent}"))
                    (footerPathInfo.TryGetValue(partPath, out var fl) ? fl
                        : (footerPathInfo[partPath] = new List<(string, string?)>()))
                        .Add((type, parent));
        }

        int hIdx = 0, fIdx = 0;
        foreach (var child in root.Children)
        {
            if (child.Type == "header")
            {
                // Skip orphaned header parts (present in the package but
                // not referenced by any section's w:headerReference). Re-
                // emitting them as `add header type=default` collides with
                // the real default header on batch replay ("Header of type
                // 'default' already exists"). Only re-emit parts that a
                // section actually links to.
                if (!headerPathInfo.TryGetValue(child.Path, out var hRefs)) continue;
                foreach (var (type, section) in hRefs)
                {
                    hIdx++;
                    EmitHeaderFooterPart(word, child.Path, "header", hIdx, items, type, section, warnings);
                }
            }
            else if (child.Type == "footer")
            {
                // Same orphan guard as header above.
                if (!footerPathInfo.TryGetValue(child.Path, out var fRefs)) continue;
                foreach (var (type, section) in fRefs)
                {
                    fIdx++;
                    EmitHeaderFooterPart(word, child.Path, "footer", fIdx, items, type, section, warnings);
                }
            }
        }
    }

    private static void EmitHeaderFooterPart(WordHandler word, string sourcePath, string kind,
                                             int targetIndex, List<BatchItem> items,
                                             string subTypeOverride = "default",
                                             string? sectionParent = null,
                                             List<DocxUnsupportedWarning>? warnings = null)
    {
        var partNode = word.Get(sourcePath);
        // BUG-DUMP9-08: tables are valid block-level OOXML inside hdr/ftr
        // (same schema as body) and Navigation surfaces them as `table`-typed
        // children, but the previous filter only kept paragraphs and silently
        // dropped tables. Iterate in source order, tracking per-type indices
        // so paragraph and table paths line up with replay output.
        // BUG-R11A(BUG3): include block-SDT children. A header/footer body can be
        // wrapped in (possibly nested) <w:sdt><w:sdtContent>; without `sdt` here
        // the walk produced zero content ops and the entire part body (PAGE/
        // NUMPAGES fields and all) was dropped on dump → batch.
        var blockChildren = (partNode.Children ?? new List<DocumentNode>())
            .Where(c => c.Type == "paragraph" || c.Type == "p"
                     || c.Type == "table" || c.Type == "tbl"
                     || c.Type == "sdt")
            .ToList();
        // partNode.Format does not expose `type`; the caller resolves the
        // role (default/first/even) from the section's headerRef.* / footerRef.*
        // map and passes it via subTypeOverride.
        var subType = subTypeOverride;

        // BUG-R6B(BUG2): a non-standard w:type (e.g. "odd", not in ST_HdrFtr
        // {even,default,first}) is pre-existing source rot. validate/get/dump
        // now all degrade gracefully on it; AddHeader/AddFooter on replay would
        // still reject "odd", so normalize the emitted op to "default" and warn
        // rather than emit a self-unreplayable script. Strict round-trip
        // fidelity isn't possible for a value the schema doesn't recognise.
        if (!string.Equals(subType, "default", StringComparison.OrdinalIgnoreCase)
            && !string.Equals(subType, "first", StringComparison.OrdinalIgnoreCase)
            && !string.Equals(subType, "even", StringComparison.OrdinalIgnoreCase))
        {
            warnings?.Add(new DocxUnsupportedWarning(
                Element: kind,
                Path: sourcePath,
                Reason: $"non-standard {kind}Reference w:type '{subType}' (not in {{default, first, even}}); emitted as 'default'"));
            subType = "default";
        }

        // Create the part with just its role (default/first/even). AddHeader/
        // AddFooter seed an empty auto paragraph; EmitParagraph(autoPresent:
        // true) on paras[0] then routes through CollapseFieldChains so a
        // PAGE-field header (the canonical case) round-trips as a typed
        // `add field` row instead of being baked into static "1" text on the
        // seed paragraph (BUG-X4-T3). Run-level formatting on multi-run
        // first paragraphs is preserved by the per-run emit path below.
        var addHeaderProps = new Dictionary<string, string> { ["type"] = subType };
        // First-page header auto-stamps <w:titlePg/> on its section (UX:
        // without titlePg, Word silently ignores type="first" headerRef).
        // Source may have headerRef-first WITHOUT titlePg — preserve that
        // shape by passing noTitlePg=true so AddHeader skips the auto-stamp.
        // Otherwise the next dump would emit a phantom `titlePage=true` key.
        if ((kind == "header" || kind == "footer")
            && string.Equals(subType, "first", StringComparison.OrdinalIgnoreCase)
            && sectionParent != null)
        {
            try
            {
                var sectionNode = word.Get(sectionParent);
                bool sourceHadTitlePg = sectionNode.Format.TryGetValue("titlePage", out var tpv)
                                     && tpv is bool b && b;
                if (!sourceHadTitlePg)
                    addHeaderProps["noTitlePg"] = "true";
            }
            catch { /* section path unresolved — fall through with auto-stamp */ }
        }
        // CONSISTENCY(headerfooter-noEvenAndOdd-opt-out): even-{header,footer}
        // auto-stamps <w:evenAndOddHeaders/> in /settings. Source whose settings
        // lacks the toggle (rare but real — Word renders inconsistently across
        // versions) gets a phantom toggle injected on replay. Suppress by
        // surfacing `noEvenAndOddHeaders=true` so AddHeader/AddFooter skip the
        // stamp. The settings raw-set already replaced /settings with the
        // source xml before this add executes.
        if ((kind == "header" || kind == "footer")
            && string.Equals(subType, "even", StringComparison.OrdinalIgnoreCase))
        {
            try
            {
                // BUG-DUMP-R4-02: `Get("/settings")` returns a node whose
                // Format dict is empty — PopulateDocSettings is only called
                // by GetRootNode, not when /settings is resolved directly.
                // Reading `Format["evenAndOddHeaders"]` off the settings node
                // therefore always returned false, so dump emitted a phantom
                // `noEvenAndOddHeaders=true` even when the source's
                // settings.xml carried the toggle. Read from root, which IS
                // populated, mirroring the `titlePage` check above (that one
                // reads off /section[N] which also runs its own populator).
                var rootNode = word.Get("/");
                bool sourceHadToggle = rootNode.Format.TryGetValue("evenAndOddHeaders", out var ev)
                                     && ev is bool eb && eb;
                if (!sourceHadToggle)
                    addHeaderProps["noEvenAndOddHeaders"] = "true";
            }
            catch { /* settings unreadable — fall through */ }
        }
        items.Add(new BatchItem
        {
            Command = "add",
            // Route per-section headers/footers to their owning section
            // (e.g. /section[2]) instead of root "/", so multi-section docs
            // that carry one default header per section don't collide on
            // replay. Falls back to "/" when the part is not owned by any
            // section in the harvested map (defensive — EmitHeadersFooters
            // already filters orphans before reaching here).
            Parent = sectionParent ?? "/",
            Type = kind,
            Props = addHeaderProps
        });

        var partTargetPath = $"/{kind}[{targetIndex}]";
        // BUG-R5B(BUG1): a header/footer body can host a textbox-bearing run
        // (e.g. a centered page-number textbox in the footer). TryEmitTextbox —
        // which AddTextbox supports for /header[N] and /footer[N] hosts — bails
        // out when ctx is null, and EmitParagraph was previously called with no
        // ctx here, so the textbox (and the PAGE field inside it) was silently
        // dropped. Build a part-scoped ctx so the textbox emit path fires and
        // unsupported-content warnings surface. Footnote/endnote/chart cursors
        // are part-local (header/footer rarely carry them, and the body ctx's
        // cursors must not be consumed from here).
        var hfCtx = new BodyEmitContext(
            FootnoteTexts: new List<string>(),
            EndnoteTexts: new List<string>(),
            FootnoteCursor: new NoteCursor(),
            EndnoteCursor: new NoteCursor(),
            ChartSpecs: new List<ChartSpec>(),
            ChartCursor: new NoteCursor(),
            ParaIdToTargetIdx: null,
            DeferredBookmarks: new List<BatchItem>(),
            TextboxCounters: new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase),
            SourceTextboxCounters: new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase),
            TableOrdinalBox: new int[1],
            CurrentCellXPathBox: new string?[1],
            CurrentCellPartBox: new string?[1],
            MovePairIds: word.BuildMovePairIdMap(),
            RawPassedParaIds: new HashSet<string>(StringComparer.OrdinalIgnoreCase),
            Warnings: warnings ?? new List<DocxUnsupportedWarning>());
        int pIdx = 0, tblIdx = 0;
        bool sawFirstPara = false;
        // BUG-DUMP-R2-NESTED-LEAD (header/footer site): a header/footer body
        // may begin with a table (CT_HdrFtr allows it). `add header`/`add footer`
        // auto-seeds an empty leading paragraph; when the first source child is a
        // table that seed has no source counterpart. Suppress the seed-reuse so
        // any later paragraph adds AFTER the table instead of overwriting the
        // leading seed, then drop the phantom seed below.
        //
        // BUG-R11A(BUG3): an SDT-wrapped header/footer body (the whole body is a
        // block <w:sdt>) is the same shape — its first child is neither a
        // paragraph the seed can host. Generalize the leading-seed opt-out to
        // "first child is a table OR a block-SDT".
        bool firstChildIsNonPara = blockChildren.Count > 0
            && (blockChildren[0].Type == "table" || blockChildren[0].Type == "tbl"
                || blockChildren[0].Type == "sdt");
        // The host part for raw-set: /document for body, otherwise the part path.
        var hfRawPart = partTargetPath;
        var hfRootXPath = kind == "header" ? "/w:hdr" : "/w:ftr";
        foreach (var child in blockChildren)
        {
            if (child.Type == "table" || child.Type == "tbl")
            {
                tblIdx++;
                // BUG-R11A(BUG1): pass the part-scoped hfCtx (was ctx: null) so a
                // block <w:sdt> nested in a header/footer table cell is emitted
                // via EmitTable's cell-SDT branch. hfCtx's note/chart cursors are
                // fresh and part-local, so threading it here consumes nothing the
                // body walk relies on.
                EmitTable(word, child.Path, tblIdx, items, ctx: hfCtx,
                          parentTablePath: null, containerPath: partTargetPath);
            }
            else if (child.Type == "sdt")
            {
                // BUG-R11A(BUG3): a block <w:sdt> that is a direct child of the
                // header/footer body. Reuse the cell-SDT machinery: rich block
                // content (the canonical PAGE/NUMPAGES footer, and the nested
                // <w:sdt><w:sdtContent><w:sdt> shape) round-trips verbatim via a
                // raw-set into the part root — injecting the OUTER sdt preserves
                // any nesting. Text-shaped controls go through the typed
                // `add sdt` path targeting the part. `cellHasContent: sawFirstPara`
                // chooses prepend (lands ahead of the auto-seed when this SDT is
                // the leading body content) vs append (after preceding paragraphs).
                EmitCellSdt(word, child.Path, partTargetPath, hfRootXPath, hfRawPart,
                            cellHasContent: sawFirstPara, items, hfCtx);
            }
            else
            {
                pIdx++;
                EmitParagraph(word, child.Path, partTargetPath, pIdx, items,
                              autoPresent: !sawFirstPara && !firstChildIsNonPara, hfCtx);
                sawFirstPara = true;
            }
        }
        // Remove the unconsumed auto-seeded leading paragraph (see above).
        if (firstChildIsNonPara)
        {
            items.Add(new BatchItem
            {
                Command = "remove",
                Path = $"{partTargetPath}/p[1]",
            });
        }

        // BUG-DUMP-HDRFTR-STRUCT-BOOKMARK: re-insert any <w:bookmarkStart>/
        // <w:bookmarkEnd> that sat at the <w:hdr>/<w:ftr> ROOT level (between block
        // paragraphs, not inside one). The block walk above only emits paragraph/
        // table/sdt content, so a header/footer-scoped cross-reference target was
        // dropped, leaving a dangling REF/PAGEREF. Replay each verbatim at its
        // source position via raw-set into this part. (Paragraph-level header/footer
        // bookmarks already survive through EmitParagraph; only root-direct-child
        // markers need this.)
        foreach (var (bmXml, relXpath, action) in word.GetPartRootStructuralBookmarks(sourcePath))
        {
            items.Add(new BatchItem
            {
                Command = "raw-set",
                Part = hfRawPart,
                Xpath = relXpath == "." ? hfRootXPath : $"{hfRootXPath}/{relXpath}",
                Action = action,
                Xml = bmXml,
            });
        }
    }

    private static void EmitComments(WordHandler word, List<BatchItem> items,
                                     Dictionary<string, int> paraIdToTargetIdx,
                                     HashSet<string>? rawPassedParaIds = null)
    {
        rawPassedParaIds ??= new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var comments = word.Query("comment");
        int targetCommentIdx = 0;  // 1-based index of the comment as it will be rebuilt
        int sourceCommentIdx = 0;  // 1-based positional index in the source comments part
        foreach (var c in comments)
        {
            targetCommentIdx++;
            sourceCommentIdx++;
            var props = FilterEmittableProps(c.Format);

            // BUG-R9A(BUG1): emit the comment body STRUCTURALLY instead of
            // flattening it to a single `text` prop. A comment body may carry
            // multiple runs (each with its own rPr) and multiple paragraphs;
            // the old flatten-to-`text` path discarded all per-run formatting
            // and any paragraph beyond the first (silent data loss). Strategy:
            //   - `add comment` carries the FIRST paragraph's FIRST run text +
            //     that run's rPr (ApplyCommentFormatKeys applies them to the
            //     lone run present at creation time).
            //   - remaining runs in the first paragraph -> `add run` into
            //     /comments/comment[N]/p[1].
            //   - additional paragraphs -> `add paragraph` into
            //     /comments/comment[N], then `add run` per run.
            // Mirrors how /body content runs/paragraphs are emitted. Plain and
            // empty comments still round-trip (single run / no run).
            //
            // Enumerate the source comment's paragraphs WITH their run children.
            // Use the positional comment index (word.Query("comment") returns
            // comments in source order, so loop position == positional index)
            // and Get(path, depth:2) so each paragraph node carries populated
            // run Children — word.Query enumerates collection children at
            // depth 0 (empty Children), which would silently re-flatten.
            var bodyParas = new List<DocumentNode>();
            for (int pIdx = 1; ; pIdx++)
            {
                DocumentNode? para;
                try { para = word.Get($"/comments/comment[{sourceCommentIdx}]/p[{pIdx}]", depth: 2); }
                catch { break; }
                if (para == null) break;
                bodyParas.Add(para);
            }

            var firstParaRuns = bodyParas.Count > 0
                ? bodyParas[0].Children.Where(IsRoundTrippableCommentRun).ToList()
                : new List<DocumentNode>();

            // BUG-DUMP-R26-4: preserve the comment's first-paragraph w14:paraId.
            // commentsExtended.xml threads replies via w15:commentEx paraIdParent,
            // keyed by these paraIds — regenerating them (EnsureAllParaIds stamps
            // a fresh id on the AddComment-built body) would silently break the
            // reply link even when the threading part itself is preserved. Forward
            // the source paraId so AddComment stamps it onto the comment body.
            if (bodyParas.Count > 0
                && bodyParas[0].Format.TryGetValue("paraId", out var cpid)
                && cpid != null && !string.IsNullOrEmpty(cpid.ToString()))
            {
                props["commentParaId"] = cpid.ToString()!;
            }

            // BUG-DUMP-R40-2: carry the comment's first-paragraph pStyle (Word
            // authors comments under the "CommentText" paragraph style). The old
            // emit took props only from the comment-node Format (no pStyle), so
            // AddComment built a default-styled paragraph and the comment lost
            // its comment-pane styling. ApplyCommentFormatKeys -> ApplyParagraph
            // LevelProperty consumes `style` and stamps the pStyle.
            if (bodyParas.Count > 0
                && bodyParas[0].Format.TryGetValue("style", out var cStyle)
                && cStyle != null && !string.IsNullOrEmpty(cStyle.ToString()))
            {
                props["style"] = cStyle.ToString()!;
            }

            // BUG-DUMP-NOTE-PBDR (comment parity): the comment's first paragraph can
            // carry direct paragraph formatting (alignment / indent / spacing / a
            // paragraph border <w:pBdr> / shading / markRPr) — none of which
            // CommentToNode surfaces onto the comment node, so the `add comment` op
            // (built from c.Format) dropped them all (only 2nd+ paragraphs, emitted
            // via FilterEmittableProps(para.Format), kept their pPr). Forward the
            // first paragraph's pPr keys (the same set EmitNoteReference forwards);
            // ApplyCommentFormatKeys -> ApplyParagraphLevelProperty applies them.
            if (bodyParas.Count > 0)
            {
                foreach (var (k, v) in FilterEmittableProps(bodyParas[0].Format))
                {
                    if (v == null) continue;
                    // allowNumPr:false — AddComment has no <w:numPr> rebuild path
                    // (unlike AddFootnote/AddEndnote), so a comment's first-para
                    // list membership is not forwarded. See IsForwardableNoteFirstParaKey.
                    if (IsForwardableNoteFirstParaKey(k, allowNumPr: false) && !props.ContainsKey(k))
                        props[k] = v.ToString()!;
                }
            }

            // BUG-R6B(BUG1): always emit `text`, even when empty. An empty
            // comment (no inline text, or only an empty table) is valid OOXML;
            // omitting `text` produced a dump op that AddComment refused to
            // replay ("'text' property is required"), silently dropping the
            // comment on round-trip. AddComment now accepts text="".
            // The first run's text + rPr ride on `add comment`; if there is no
            // first run (empty comment) fall back to empty text.
            // A leading tab/ptab run must NOT be swallowed as the seed text
            // (its Text is empty, so the <w:tab/> would be lost) — leave it for
            // the structural body-run pass below. Mirror EmitNoteReference.
            int commentSeedSkip = 0;
            if (firstParaRuns.Count > 0
                && (firstParaRuns[0].Type == "run" || firstParaRuns[0].Type == "r")
                // BUG-DUMP-NOTE-HYPHEN: don't flatten a structural-hyphen first run
                // into the `add comment text=` seed — let EmitContainerBodyRuns emit
                // it as `add r hyphen=` so <w:softHyphen/>/<w:noBreakHyphen/> survive.
                && !firstParaRuns[0].Format.ContainsKey("_hasHyphen"))
            {
                var firstRun = firstParaRuns[0];
                props["text"] = firstRun.Text ?? string.Empty;
                MergeRunFormatProps(props, firstRun);
                commentSeedSkip = 1;
            }
            else if (firstParaRuns.Count > 0)
            {
                props["text"] = "";
            }
            else
            {
                // BUG-DUMP-NOTE-EMPTYLEAD (comment parity): when the comment's
                // first paragraph is empty but later paragraphs hold the content,
                // seed p[1] EMPTY — the structural pass below (the pi>=1 loop)
                // re-emits those later paragraphs. The whole-comment c.Text
                // fallback is only for a SINGLE degenerate paragraph; with later
                // paragraphs it duplicated the body. Mirrors the note seed guard.
                props["text"] = bodyParas.Count > 1 ? string.Empty : (c.Text ?? string.Empty);
            }
            // Map anchoredTo (source paraId path) -> target paragraph index.
            // anchoredTo looks like "/body/p[@paraId=00100000]"; parse and
            // resolve via the paraId map we built during EmitBody.
            string parentTarget = "/body/p[1]";  // safe fallback to first body para
            // BUG-DUMP-H103: true when this comment's range markers live in a
            // paragraph emitted VERBATIM by EmitCrossParagraphFieldMember (a TOC /
            // cross-paragraph field span). Such a paragraph already carries the
            // <w:commentRangeStart/End/Reference> markers with their SOURCE id, so
            // the typed `add comment` must place NO range markers (it would
            // duplicate the start under a fresh id) and must REUSE the source id
            // so the verbatim markers resolve to this definition.
            bool anchorIsRawPassed = false;
            if (props.TryGetValue("anchoredTo", out var anchor))
            {
                // BUG-R4 (DBF-R4-01): a comment anchored inside a table cell
                // resolves to "/body/tbl[N]/tr[M]/tc[K]/p[J]" — a positional
                // path that is structurally stable across dump→batch (the table
                // is re-created with fresh paraIds, so the body paraId map can't
                // help). Pass it through verbatim so the comment re-anchors in
                // the cell instead of falling back to /body/p[1].
                if (anchor.Contains("/tbl[", StringComparison.OrdinalIgnoreCase))
                {
                    parentTarget = anchor;
                }
                else
                {
                    var pid = ExtractParaId(anchor);
                    if (pid != null && paraIdToTargetIdx.TryGetValue(pid, out var idx))
                        parentTarget = $"/body/p[{idx}]";
                    if (pid != null && rawPassedParaIds.Contains(pid))
                        anchorIsRawPassed = true;
                }
                props.Remove("anchoredTo");
            }
            // BUG-DUMP4-03: emit the 1-based run index where the source
            // CommentRangeStart sits inside its paragraph so replay can
            // narrow the anchor instead of widening to the entire para.
            // 0 means "before all runs" (paragraph start); >=1 means
            // "after run N". AddComment already accepts a run-targeted
            // parent path (/body/p[N]/r[M]), but we keep the prop on the
            // paragraph-level emit so the wire format stays uniform with
            // the existing parent-resolution logic — replay can switch on
            // runStart later without changing the schema.
            // BUG-DUMP-R26-3: when the comment range END lives in a DIFFERENT
            // paragraph than its start, the range spans paragraphs. The old
            // single-op `add comment` crammed rangeStart+rangeEnd+ref into the
            // start paragraph, collapsing the span to one paragraph. Detect the
            // multi-paragraph case and emit a two-marker round-trip: the `add
            // comment` carries rangeOpen=true (places only the start), then a
            // follow-up `add comment rangeEnd=true` closes the range at the end
            // paragraph. Resolved end target path is stashed for after the
            // `add comment` op is appended (replay order: start then end).
            string? rangeEndParent = null;
            int rangeEndRunIdx = 0;
            if (anchorIsRawPassed && c.Format.TryGetValue("id", out var rawCid) && rawCid != null)
            {
                // BUG-DUMP-H103: definition-only emit. The range start/end and the
                // reference run are ALREADY present in the verbatim raw-passed
                // paragraph(s) with the source id; reuse that source id on the
                // definition so they resolve, and place NO body markers here. `id`
                // is preserved (not removed below); `range=none` tells AddComment to
                // create the comment + body but skip all anchor placement.
                props["id"] = rawCid.ToString()!;
                props["range"] = "none";
            }
            else if (c.Format.TryGetValue("id", out var cid) && cid != null)
            {
                var runStart = word.FindCommentAnchorRunIndex(cid.ToString()!);
                // 0 = before all runs (paragraph start); always emit so
                // replay knows the anchor is positional, not whole-paragraph.
                props["runStart"] = runStart.ToString();
                // BUG-DUMP-COMMENT-POINTREF: a zero-width / point-anchored
                // comment in the source carries only <w:commentReference> (no
                // commentRangeStart/End). Carry range=false so AddComment
                // replays a reference-only run instead of synthesizing a
                // spurious range — preserving the point comment's identity.
                if (!word.CommentHasRange(cid.ToString()!))
                {
                    props["range"] = "false";
                }
                else if (word.FindCommentRangeEnd(cid.ToString()!) is { } endInfo)
                {
                    // Resolve the END paragraph to the same target-index space
                    // the start anchor uses. Table-cell paths pass through
                    // verbatim (positionally stable); body paras map via paraId.
                    string? endTarget = null;
                    if (endInfo.path.Contains("/tbl[", StringComparison.OrdinalIgnoreCase))
                        endTarget = endInfo.path;
                    else
                    {
                        var endPid = ExtractParaId(endInfo.path);
                        if (endPid != null && paraIdToTargetIdx.TryGetValue(endPid, out var eIdx))
                            endTarget = $"/body/p[{eIdx}]";
                        // Positional fallback: a source paragraph with no
                        // w14:paraId surfaces as /body/p[N]. Top-level body
                        // paragraphs replay 1:1 positionally in EmitBody, so the
                        // positional path is a valid target anchor as-is.
                        else if (System.Text.RegularExpressions.Regex.IsMatch(
                                     endInfo.path, @"^/body/p\[\d+\]$"))
                            endTarget = endInfo.path;
                    }
                    // Only split into a two-marker op when the end paragraph is
                    // genuinely different from the start paragraph. A single-
                    // paragraph range still round-trips through the one-op path.
                    if (endTarget != null && endTarget != parentTarget)
                    {
                        props["rangeOpen"] = "true";
                        rangeEndParent = endTarget;
                        rangeEndRunIdx = endInfo.runIndex;
                    }
                }
            }
            // The comment id is allocated by AddComment on the target side;
            // do not propagate the source id (would conflict on replay).
            // EXCEPTION (BUG-DUMP-H103): a definition-only comment whose verbatim
            // raw-passed paragraph already holds source-id range markers MUST keep
            // that source id so the markers resolve — `range=none` set it above.
            if (!anchorIsRawPassed)
                props.Remove("id");
            // BUG-X7-04 (T-4): previously dropped `date` so dump→replay always
            // re-stamped the comment with the SDK's "now". That breaks
            // archival / audit-trail use cases where the source timestamp is
            // load-bearing. Preserve it; AddComment accepts an explicit
            // ISO-8601 date and the SDK will use it instead of stamping.

            items.Add(new BatchItem
            {
                Command = "add",
                Parent = parentTarget,
                Type = "comment",
                Props = props
            });

            // BUG-DUMP-R26-3: close a multi-paragraph comment range at its end
            // paragraph. The `add comment rangeOpen=true` above placed only the
            // CommentRangeStart; this op places the CommentRangeEnd + reference
            // run at the (different) end paragraph so the comment scopes the
            // full span. Emitted immediately after the open so the LIFO match in
            // AddComment pairs them correctly.
            if (rangeEndParent != null)
            {
                items.Add(new BatchItem
                {
                    Command = "add",
                    Parent = rangeEndParent,
                    Type = "comment",
                    Props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["rangeEnd"] = "true",
                        ["runEnd"] = rangeEndRunIdx.ToString(),
                    }
                });
            }

            // BUG-R9A(BUG1): structural emit of the remainder of the comment
            // body. The target comment is rebuilt at /comments/comment[N]
            // where N == targetCommentIdx (comments replay in source order).
            string targetCommentPath = $"/comments/comment[{targetCommentIdx}]";

            // Remaining runs of the first paragraph (run [1] already rode on
            // `add comment`). p[1] always exists after `add comment`.
            // BUG-R13A: coalesce hyperlink runs so a hyperlink in the comment
            // body round-trips as a typed `add hyperlink` (was dropped as a
            // flat `add r` with unsupported url/isHyperlink props).
            EmitContainerBodyRuns(word, firstParaRuns.Skip(commentSeedSkip).ToList(),
                $"{targetCommentPath}/p[1]", items);

            // Additional paragraphs (paragraph [1] is the `add comment` body).
            for (int pi = 1; pi < bodyParas.Count; pi++)
            {
                var para = bodyParas[pi];
                var paraProps = FilterEmittableProps(para.Format);
                paraProps.Remove("text");  // text is carried by per-run emits below
                items.Add(new BatchItem
                {
                    Command = "add",
                    Parent = targetCommentPath,
                    Type = "paragraph",
                    Props = paraProps.Count > 0 ? paraProps : null
                });
                var runs = para.Children.Where(IsRoundTrippableCommentRun).ToList();
                // AddParagraph with no `text` produces an empty paragraph; emit
                // each run so per-run formatting survives. The new paragraph is
                // the (pi+1)-th paragraph of the comment.
                EmitContainerBodyRuns(word, runs, $"{targetCommentPath}/p[{pi + 1}]", items);
            }
        }

        // BUG-DUMP-R26-4: round-trip word/commentsExtended.xml (modern comment-
        // reply threading). Emitted once, AFTER every `add comment` so all the
        // comment paragraphs (with their preserved w14:paraId) exist on the
        // target before the threading part references them. Whole-part replace.
        var commentsExXml = word.GetCommentsExtendedXml();
        if (!string.IsNullOrEmpty(commentsExXml))
        {
            items.Add(new BatchItem
            {
                Command = "raw-set",
                Part = "/commentsExtended",
                Action = "replace",
                Xml = commentsExXml,
            });
        }
    }

    // BUG-R9A(BUG1): a comment-body run is round-trippable as a plain `add run`
    // only when it is text-carrying (no drawing / field / break / footnote-ref
    // structure). Comment bodies in practice hold plain text runs; richer
    // structure inside a comment is rare and out of scope here — skip such runs
    // rather than mis-emit them as plain text.
    private static bool IsRoundTrippableCommentRun(DocumentNode run)
    {
        return run.Type == "run" || run.Type == "r";
    }

    // BUG-R9A(BUG1): emit one comment-body run as `add run`, carrying its text
    // and rPr (italic/bold/color/size/font/…). Mirrors EmitPlainOrHyperlinkRun
    // for /body runs, minus the hyperlink/revision special-casing (comment
    // bodies don't carry those in the supported round-trip).
    private static void EmitCommentRun(WordHandler word, DocumentNode run, string paraTargetPath, List<BatchItem> items, int hlBaseline = 0)
    {
        // BUG-R13A: a run flattened out of a <w:hyperlink> wrapper carries
        // url/anchor/isHyperlink (and _hyperlinkParent) Format keys that
        // `add r` does not understand — emitting it as a flat `add r` silently
        // dropped the hyperlink wrapper + URL on replay (only the link text
        // survived as a plain run). Route such runs through the body walker's
        // EmitPlainOrHyperlinkRun, which emits a proper typed `add hyperlink`
        // op (rebuilding the <w:hyperlink> + rel relationship). Plain runs fall
        // through to the flat `add r` path unchanged. Multi-run hyperlinks are
        // coalesced upstream in EmitContainerBodyRuns; this single-run guard
        // covers the in-loop callers that emit one run at a time.
        if (run.Format.ContainsKey("url") || run.Format.ContainsKey("anchor")
            || run.Format.ContainsKey("isHyperlink"))
        {
            EmitPlainOrHyperlinkRun(word, run, paraTargetPath, items, null, hlBaseline);
            return;
        }
        // Tab-only run (<w:r><w:tab/></w:r>, Type=="tab", empty Text): the
        // generic path below emitted an EMPTY run and the tab vanished,
        // shifting every footnote/endnote/comment line that aligns its text
        // after the reference mark. Mirror the body walker's TryEmitTabRun:
        // AddText splits "\t" back into a TabChar.
        if (run.Type == "tab")
        {
            var tabProps = FilterEmittableProps(run.Format);
            tabProps["text"] = "\t";
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = paraTargetPath,
                Type = "r",
                Props = tabProps
            });
            return;
        }
        // BUG-DUMP-NOTE-HYPHEN: a run carrying a structural <w:softHyphen/> /
        // <w:noBreakHyphen/> (RunToNode stamps _hasHyphen) must emit a typed
        // `add r --prop hyphen=soft|noBreak` — the generic path below persists
        // GetRunText's cached U+00AD/U+2011 glyph as literal <w:t> text and drops
        // the structural hyphen element. The body walk handles this via
        // TryEmitHyphenRun; mirror it here so footnote/endnote/comment runs do too.
        if (run.Format.ContainsKey("_hasHyphen"))
        {
            var hyProps = FilterEmittableProps(run.Format);
            hyProps.Remove("_hasHyphen");
            hyProps["hyphen"] = run.Format.TryGetValue("_hasHyphen", out var hk)
                && string.Equals(hk?.ToString(), "soft", StringComparison.OrdinalIgnoreCase)
                ? "soft" : "noBreak";
            if (!string.IsNullOrEmpty(run.Text)) hyProps["text"] = run.Text!;
            else hyProps.Remove("text");
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = paraTargetPath,
                Type = "r",
                Props = hyProps
            });
            return;
        }
        var rProps = FilterEmittableProps(run.Format);
        if (!string.IsNullOrEmpty(run.Text))
            rProps["text"] = run.Text!;
        items.Add(new BatchItem
        {
            Command = "add",
            Parent = paraTargetPath,
            Type = "r",
            Props = rProps.Count > 0 ? rProps : null
        });
    }

    // BUG-R13A: emit a sequence of container-body (comment / footnote / endnote)
    // runs, coalescing consecutive runs that share the same <w:hyperlink>
    // wrapper into one structured `add hyperlink` (+ per-run `add r`) so a
    // multi-run formatted hyperlink survives the round-trip with every run's
    // rPr intact. Reuses the body-paragraph walker's CoalesceHyperlinkRuns /
    // EmitPlainOrHyperlinkRun machinery (single source of truth for hyperlink
    // emit). Non-hyperlink runs pass through EmitCommentRun unchanged.
    private static void EmitContainerBodyRuns(WordHandler word, List<DocumentNode> runs, string paraTargetPath, List<BatchItem> items)
    {
        // BUG-R14B: capture the hyperlink baseline ONCE for this container body
        // so multi-run hyperlinks re-index from 1 within it (mirrors the body
        // walker; per-run capture would mis-reset a 2nd hyperlink to index 1).
        int hlBaseline = items.Count(it => it.Type == "hyperlink"
            && string.Equals(it.Parent, paraTargetPath, StringComparison.Ordinal));
        foreach (var run in CoalesceHyperlinkRuns(runs))
            EmitCommentRun(word, run, paraTargetPath, items, hlBaseline);
    }

    // BUG-R9A(BUG1): fold a run's rPr format keys into the `add comment` prop
    // bag so ApplyCommentFormatKeys applies them to the comment's first run.
    // `text` is set separately by the caller; never copy paragraph/comment-level
    // keys here (the run node carries only run-level format).
    private static void MergeRunFormatProps(Dictionary<string, string> props, DocumentNode run)
    {
        var rProps = FilterEmittableProps(run.Format);
        foreach (var (k, v) in rProps)
        {
            if (string.Equals(k, "text", StringComparison.OrdinalIgnoreCase)) continue;
            props[k] = v;
        }
    }

    // BUG-R12A(BUG3): emit a footnote/endnote STRUCTURALLY (mirrors the R9
    // comment-body fix EmitComments above) instead of flattening the note body
    // to a single `text` prop. A note body may carry multiple runs (each with
    // its own rPr) and multiple paragraphs; the old flatten-to-`text` path
    // (BodyEmitContext.FootnoteTexts/EndnoteTexts) discarded all per-run
    // formatting and any paragraph beyond the first (silent data loss).
    //
    // Strategy (identical to EmitComments):
    //   - `add footnote`/`add endnote` (anchored on the body carrier paragraph)
    //     carries the FIRST content paragraph's FIRST content run text + that
    //     run's rPr (ApplyFootnoteEndnoteFormatKeys applies them to the lone
    //     authored run AddFootnote/AddEndnote seeds at creation time).
    //   - remaining runs in the first paragraph -> `add r` into
    //     /footnote[N]/p[1] (or /endnote[N]/p[1]).
    //   - additional paragraphs -> `add paragraph` into /footnote[N], then
    //     `add r` per run.
    //
    // <paramref name="kind"/> is "footnote" or "endnote"; <paramref
    // name="sourceNoteIdx"/> is the 1-based positional index in the source note
    // part (== document-order reference cursor + 1, since references walk in
    // order). <paramref name="targetNoteIdx"/> is the 1-based index of the note
    // as it will be rebuilt — equal to the number of `add footnote`/`add
    // endnote` ops already emitted (including this one). The note reference mark
    // run (footnoteRef/endnoteRef, empty text) is skipped: AddFootnote/AddEndnote
    // recreates it on replay.
    private static void EmitNoteReference(WordHandler word, string kind, int sourceNoteIdx,
                                          int targetNoteIdx, string carrierPath, List<BatchItem> items,
                                          DocumentNode? bodyRefRun = null)
    {
        // BUG-DUMP-ENDNOTE-ID: the source-side `/{kind}[N]` path resolves by
        // note Id (== N), NOT by ordinal position among the user notes —
        // /endnote[2] means "endnote whose w:id=2", not "the 2nd endnote". The
        // 1-based document-order reference cursor (sourceNoteIdx) only equals the
        // Id when the part's user notes start at id 1 (the convention Word and
        // our own AddFootnote/AddEndnote use: separators at id -1/0, first user
        // note at id 1). Some editors number endnote separators at id 0/1, so the
        // first user endnote is id 2 and /endnote[1] resolves to the
        // continuationSeparator (empty body) — every endnote body was silently
        // dropped while the footnote path round-tripped by coincidence of id
        // convention. Translate the ordinal cursor to the real source note Id by
        // enumerating user notes (id > 0) in document order, then address the
        // source by id-qualified path. The rebuilt-side targetNotePath stays
        // positional: AddFootnote/AddEndnote always allocate ids 1..N, so on the
        // rebuilt part ordinal == id.
        int sourceNoteId = ResolveUserNoteId(word, kind, sourceNoteIdx);

        // Count the note's paragraphs from its raw XML (deterministic). A
        // depth-N note Get returns EMPTY children — it does not enumerate its
        // <w:p> grandchildren — and, inside the dump session, out-of-range
        // /<kind>[N]/p[K] does NOT reliably throw (it clamped, producing a flood
        // of empty paragraphs), so neither the children list nor a Get-until-
        // throw loop is a safe bound. The raw XML <w:p…> open-tag count is.
        string sourceNotePath = $"/{kind}[@{kind}Id={sourceNoteId}]";
        var noteXml = word.GetElementXml(sourceNotePath);
        // BUG-DUMP-R27-5: enumerate the note's DIRECT block children (w:p AND
        // w:tbl) in document order. The old code regex-counted EVERY <w:p> open
        // (which includes paragraphs nested inside table cells) and walked
        // /<kind>[N]/p[K] positionally — so a note containing a <w:tbl>
        // double-counted the cell paragraphs as if they were top-level note
        // paragraphs, addressed out-of-range /<kind>[N]/p[K] slots (clamping to
        // empty), and never emitted the table at all. Walk the direct children
        // with a depth-tracked scan (mirrors ComputeParagraphChildDocOrder) so
        // tables route through EmitTable against the note host below.
        var directChildren = EnumerateNoteDirectChildren(noteXml);

        // Resolve each direct-paragraph child to its positional /<kind>[N]/p[K]
        // path (K is the 1-based index AMONG DIRECT paragraphs, which is exactly
        // how the handler indexes /<kind>[N]/p[K]). Tables keep their direct
        // 1-based tbl ordinal for the EmitTable source path.
        var bodyParas = new List<DocumentNode>();
        // Block-order list parallel to directChildren: each entry is either a
        // resolved paragraph node (kind "p") or a 1-based table ordinal.
        var blockOrder = new List<(string Kind, DocumentNode? Para, int TblOrdinal)>();
        int directParaIdx = 0;
        int directTblIdx = 0;
        foreach (var ck in directChildren)
        {
            if (ck == "p")
            {
                directParaIdx++;
                DocumentNode? para = null;
                try { para = word.Get($"{sourceNotePath}/p[{directParaIdx}]", depth: 2); }
                catch { /* leave null */ }
                if (para != null) bodyParas.Add(para);
                blockOrder.Add(("p", para, 0));
            }
            else if (ck == "tbl")
            {
                directTblIdx++;
                blockOrder.Add(("tbl", null, directTblIdx));
            }
        }

        var firstParaRuns = bodyParas.Count > 0
            ? bodyParas[0].Children.Where(c => IsRoundTrippableNoteRun(word, c)).ToList()
            : new List<DocumentNode>();

        // `add footnote`/`add endnote` requires a non-empty `text` (AddFootnote/
        // AddEndnote throw without it). Carry the first content run's text + rPr;
        // fall back to the concatenated note text only when no content run
        // resolves (degenerate/empty note).
        var noteProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        // BUG-DUMP-R40-1: carry the note's first-paragraph pStyle. AddFootnote/
        // AddEndnote hardcode pStyle="FootnoteText"/"EndnoteText" on the
        // synthesized note paragraph, but the source note may reference a
        // DIFFERENT style id (e.g. "style24" / "Endnote" from another editor). The old
        // emit dropped the source style, so the rebuilt note carried a DANGLING
        // pStyle="EndnoteText" (not present in the source styles.xml) and lost
        // the note's hanging indent / size / line-number suppression. Forward the
        // source style id so ApplyFootnoteEndnoteFormatKeys -> ApplyParagraph
        // LevelProperty overrides the hardcoded default with the real style.
        if (bodyParas.Count > 0)
        {
            var srcNoteStyle = bodyParas[0].Format.TryGetValue("style", out var noteStyle)
                && noteStyle != null ? noteStyle.ToString() : null;
            // BUG-DUMP: AddFootnote/AddEndnote hard-code pStyle=FootnoteText/
            // EndnoteText on the synthesized note paragraph. When the SOURCE note
            // paragraph carries NO pStyle (it inherits the default paragraph
            // style — e.g. Normal, which has a non-zero spaceAfter), stamping
            // FootnoteText (spaceAfter=0) collapses the inter-paragraph gap and
            // the note renders shorter, shifting the page. Emit an explicit empty
            // style to signal "no pStyle" so the apply side strips the hard-coded
            // default; a real source style id is forwarded verbatim (R40-1).
            noteProps["style"] = string.IsNullOrEmpty(srcNoteStyle) ? "" : srcNoteStyle!;
        }
        // Forward the note's first-paragraph PARAGRAPH-level formatting. The
        // `add footnote`/`add endnote` step only seeds p[1] with text + style +
        // the ref-mark/run rPr; without this, a note paragraph's explicit line
        // spacing (Arabic notes carry <w:spacing w:line="200" w:lineRule="exact">),
        // direction, indent, and ¶-mark rPr were dropped — the note rendered
        // taller, pulling body content down and shifting page breaks. Carry the
        // paragraph-scoped keys (spacing/line/ind/jc/direction + the markRPr.*
        // family); ApplyFootnoteEndnoteFormatKeys routes them through
        // ApplyParagraphLevelProperty / the markRPr.* branch. Skip effective.*
        // (style-resolved, not authored), the run-text keys MergeRunFormatProps
        // already carries, and text/style handled above.
        if (bodyParas.Count > 0)
        {
            // BUG-DUMP-NOTE-PBDR: source from FilterEmittableProps so multi-segment
            // paragraph props (pbdr.<side> + .sz/.color/.space, shading) arrive
            // FOLDED into the single compound value ApplyParagraphLevelProperty
            // parses — forwarding the raw sub-keys dropped the border weight/color.
            foreach (var (k, v) in FilterEmittableProps(bodyParas[0].Format))
            {
                if (v == null) continue;
                // allowNumPr:true — AddFootnote/AddEndnote rebuild a direct
                // <w:numPr> (BUG-DUMP-NOTE-NUMPR), so numId/numLevel are forwarded
                // here. NOT numFmt/listStyle/start (would trigger ad-hoc numbering-
                // definition creation, BUG-DUMP26-01; the /numbering raw-set holds it).
                bool isParaKey = IsForwardableNoteFirstParaKey(k, allowNumPr: true);
                // Don't forward style-INHERITED numbering (the pStyle, forwarded
                // separately, supplies it) — promoting inherited->explicit would
                // duplicate it. numInherited is skipped by FilterEmittableProps, so
                // read it from the raw first-para Format.
                if ((k == "numId" || k == "numLevel")
                    && bodyParas[0].Format.TryGetValue("numInherited", out var noteNi)
                    && string.Equals(noteNi?.ToString(), "true", StringComparison.OrdinalIgnoreCase))
                    isParaKey = false;
                if (isParaKey && !noteProps.ContainsKey(k))
                    noteProps[k] = v.ToString()!;
            }
        }
        // BUG-DUMP-R42-1: capture the ref-mark run's char-style link. Word's
        // note ref mark carries <w:rStyle w:val="FootnoteReference"/> (or
        // "EndnoteReference"), linking it to the style in styles.xml that
        // defines the superscript appearance. AddFootnote/AddEndnote used to
        // hardcode an inline <w:vertAlign w:val="superscript"/> instead,
        // severing the style link (style still in styles.xml, run no longer
        // references it). Forward the source rStyle so AddFootnote/AddEndnote
        // restores it; when the source mark had no rStyle, leave the prop unset
        // and AddFootnote/AddEndnote falls back to the inline superscript.
        if (bodyParas.Count > 0)
        {
            var refMarkRun = bodyParas[0].Children.FirstOrDefault(c =>
                (c.Type == "run" || c.Type == "r")
                && string.IsNullOrEmpty(c.Text)
                && (word.GetElementXml(c.Path)?.Contains("footnoteRef", StringComparison.Ordinal) == true
                    || word.GetElementXml(c.Path)?.Contains("endnoteRef", StringComparison.Ordinal) == true));
            if (refMarkRun != null)
            {
                var refRaw = word.GetElementXml(refMarkRun.Path);
                if (!string.IsNullOrEmpty(refRaw))
                {
                    var rsMatch = System.Text.RegularExpressions.Regex.Match(
                        refRaw, "<w:rStyle\\s+w:val=\"([^\"]*)\"");
                    if (rsMatch.Success && !string.IsNullOrEmpty(rsMatch.Groups[1].Value))
                        noteProps["referenceStyle"] = rsMatch.Groups[1].Value;
                    // The in-note mark run can ALSO carry direct formatting
                    // (rFonts/sz alongside the rStyle) — same hazard as the
                    // body-side reference run. Carry the verbatim <w:rPr>.
                    try
                    {
                        var refRunEl = System.Xml.Linq.XElement.Parse(refRaw);
                        var wNs3 = (System.Xml.Linq.XNamespace)"http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                        var refRPrEl = refRunEl.Element(wNs3 + "rPr");
                        if (refRPrEl != null)
                            noteProps["referenceMarkRPr"] = refRPrEl.ToString(System.Xml.Linq.SaveOptions.DisableFormatting);
                    }
                    catch { /* keep the rStyle-only fallback */ }
                }
            }
        }
        // The BODY-side reference run (the <w:footnoteReference>/<w:endnoteReference>
        // host in the document text) can carry direct formatting beyond the
        // FootnoteReference char style — real documents shrink the superscript
        // mark with run-level rFonts/sz (e.g. Gill Sans sz=18). AddFootnote/
        // AddEndnote rebuilt that run with ONLY the rStyle, so the mark rendered
        // at the inherited size, inflating every host line and shifting the page.
        // Carry the run's verbatim <w:rPr> so the apply side restores it.
        if (bodyRefRun != null)
        {
            // BUG-DUMP-NOTEREF-CUSTOMMARK-DEL: the body reference run may itself be a
            // TRACKED REVISION (most commonly a <w:del> wrapping a deleted note
            // reference). EmitNoteReference routes the run to `add footnote`/`add
            // endnote`, bypassing the normal run emit's revision wrapping — so without
            // this the deletion attribution was lost and the deleted reference
            // resurfaced as a live, accepted reference (and its custom mark, which
            // rides in <w:delText>, became live <w:t>). Carry revision.* so the apply
            // re-wraps the rebuilt reference run.
            foreach (var rk in new[] { "revision.type", "revision.author", "revision.date", "revision.id" })
                if (bodyRefRun.Format.TryGetValue(rk, out var rv) && rv != null
                    && !string.IsNullOrEmpty(rv.ToString()))
                    noteProps["reference." + rk] = rv.ToString()!;
            var bodyRunXml = word.GetElementXml(bodyRefRun.Path);
            if (!string.IsNullOrEmpty(bodyRunXml))
            {
                try
                {
                    var runEl = System.Xml.Linq.XElement.Parse(bodyRunXml);
                    var wNs2 = (System.Xml.Linq.XNamespace)"http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                    var rPrEl = runEl.Element(wNs2 + "rPr");
                    if (rPrEl != null)
                        noteProps["referenceRPr"] = rPrEl.ToString(System.Xml.Linq.SaveOptions.DisableFormatting);
                    // BUG-DUMP-NOTEREF-CUSTOMMARK: a note reference may use a
                    // CUSTOM mark instead of an auto-number — Word sets
                    // <w:footnoteReference w:customMarkFollows="1" w:id="N"/> and
                    // the literal mark glyph lives in a SIBLING <w:t> in the SAME
                    // body run (e.g. "*", "†"). The typed rebuild emitted a bare
                    // <w:footnoteReference w:id="N"/>, dropping BOTH the attribute
                    // and the glyph (the asterisk vanished from body text). Carry
                    // the flag + the mark text so AddFootnote/AddEndnote restore them.
                    var refChild = runEl.Element(wNs2 + "footnoteReference")
                                   ?? runEl.Element(wNs2 + "endnoteReference");
                    var cmf = refChild?.Attribute(wNs2 + "customMarkFollows")?.Value;
                    if (cmf is "1" or "true" or "on")
                    {
                        noteProps["referenceCustomMarkFollows"] = "1";
                        // BUG-DUMP-NOTEREF-CUSTOMMARK-DEL: in a TRACKED-DELETION
                        // reference run (<w:del>) the custom mark glyph rides in
                        // <w:delText>, not <w:t> — reading only <w:t> dropped the
                        // mark (e.g. a deleted footnote "1" silently lost the "1").
                        // Concat both so the mark survives whether the run is live
                        // or a deletion.
                        var markText = string.Concat(
                            runEl.Elements()
                                .Where(e => e.Name == wNs2 + "t" || e.Name == wNs2 + "delText")
                                .Select(e => e.Value));
                        noteProps["referenceCustomMark"] = markText;
                        // BUG-DUMP-H97: an academic-style custom mark is often a
                        // SYMBOL glyph (<w:sym w:font="Symbol" w:char="F020"/>), not
                        // <w:t> text. Reading only <w:t>/<w:delText> captured an empty
                        // mark → customMarkFollows="true" with no mark = dangling flag
                        // + the visible marker glyph vanished. Capture the <w:sym>
                        // font+char so AddFootnote/AddEndnote can rebuild it.
                        var symEl = runEl.Elements(wNs2 + "sym").FirstOrDefault();
                        if (symEl != null)
                        {
                            var symFont = symEl.Attribute(wNs2 + "font")?.Value ?? "";
                            var symChar = symEl.Attribute(wNs2 + "char")?.Value ?? "";
                            if (symChar.Length > 0)
                                noteProps["referenceCustomMarkSym"] = symFont + ":" + symChar;
                        }
                    }
                }
                catch { /* malformed run XML — keep the rStyle-only fallback */ }
            }
        }
        // How many leading runs the `add <kind>` seed consumes. The seed run
        // can only carry TEXT (it becomes the lone authored <w:t> run after the
        // refmark); a leading tab/ptab run must NOT be swallowed here — its
        // Text is empty, so consuming it flattens the <w:tab/> to nothing and
        // de-indents every note that tabs after its reference mark (e.g. a
        // footnote "<refmark><tab><hyperlink>"). Leave those for the structural
        // body-run pass below.
        int noteSeedSkip = 0;
        if (firstParaRuns.Count > 0
            && (firstParaRuns[0].Type == "run" || firstParaRuns[0].Type == "r")
            // BUG-DUMP-NOTE-HYPHEN: a first content run carrying a structural
            // hyphen must NOT be flattened into the `add <kind> text=` seed (which
            // persists the cached U+00AD/U+2011 glyph and drops the element). Seed
            // empty and let EmitContainerBodyRuns emit it as `add r hyphen=`.
            && !firstParaRuns[0].Format.ContainsKey("_hasHyphen"))
        {
            var firstRun = firstParaRuns[0];
            // Emit the FIRST content run's text VERBATIM. AddFootnote/AddEndnote
            // no longer prepend a synthetic leading space and GetFootnoteText no
            // longer trims one, so the apply side stores exactly what we emit —
            // the source first <w:t> round-trips byte-faithfully (an Arabic note
            // starting "خاص…" stays "خاص…", not " خاص…"). A genuinely authored
            // leading space is preserved for the same reason.
            noteProps["text"] = firstRun.Text ?? string.Empty;
            MergeRunFormatProps(noteProps, firstRun);
            noteSeedSkip = 1;
        }
        else if (firstParaRuns.Count > 0)
        {
            // First round-trippable run is a tab/ptab: seed empty text (just the
            // refmark) and let EmitContainerBodyRuns emit it (and the rest) in
            // order — it round-trips the tab as `add r text="\t"`.
            noteProps["text"] = "";
        }
        else
        {
            // BUG-DUMP-R27-5: a note whose FIRST direct child is a <w:tbl> (no
            // leading paragraph) has NO authored leading text. OOXML still
            // requires the <w:*Ref/> mark to live in the note's first
            // paragraph, so `add <kind>` always fabricates one — but its text
            // must be EMPTY (just the refmark), never the note's concatenated
            // descendant text. The old fallback pulled `Get(note).Text`, which
            // walks Descendants<Text> and so vacuumed every TABLE CELL's text
            // into a phantom leading run (e.g. " t1at1bt2at2b"), duplicating
            // the cell content that EmitTable re-emits below. Only fall back to
            // the note's own text when the note actually leads with a paragraph
            // (degenerate/empty-run paragraph) — for a table-leading note the
            // refmark paragraph stays text-less and the table round-trips
            // through the blockOrder EmitTable pass.
            bool leadsWithTable = directChildren.Count > 0 && directChildren[0] == "tbl";
            // BUG-DUMP-NOTE-EMPTYLEAD: a note whose FIRST paragraph is empty but
            // which has SUBSEQUENT block children (a leading blank line before
            // the ref-mark/content paragraph — e.g. a footnote authored as
            // <w:p/><w:p><w:footnoteRef/> text</w:p>) must seed p[1] EMPTY. The
            // content lives in the later paragraph(s) that the structural body-run
            // pass re-emits below. The whole-note Get(sourceNotePath).Text fallback
            // is meant ONLY for a SINGLE degenerate paragraph whose visible text
            // sits in non-round-trippable runs; when later blocks exist it
            // vacuumed the content paragraph's text into the seed AND the
            // structural pass emitted that paragraph again, DUPLICATING the note
            // body — the doubled note rendered taller, pulled body content down,
            // and shifted every subsequent page break. Mirrors the leadsWithTable
            // guard (R27-5), which already excludes table content from the seed.
            bool hasLaterBlocks = directChildren.Count > 1;
            string fallback = "";
            if (!leadsWithTable && !hasLaterBlocks)
            {
                try { fallback = word.Get(sourceNotePath).Text ?? ""; }
                catch { /* leave empty */ }
            }
            noteProps["text"] = fallback;
        }

        items.Add(new BatchItem
        {
            Command = "add",
            Parent = carrierPath,
            Type = kind,
            Props = noteProps
        });

        // Structural emit of the remainder. The target note is rebuilt at
        // /<kind>[targetNoteIdx] (notes replay in reference order; the reserved
        // separator / continuationSeparator notes, ids -1/0, are excluded by
        // Query and by the positional /<kind>[N] path index). The first authored
        // run + p[1] already exist after the `add <kind>` above.
        string targetNotePath = $"/{kind}[{targetNoteIdx}]";

        // BUG-DUMP-NOTE-TABSTOPS: the note's first paragraph can carry explicit
        // tab stops — most importantly <w:tab w:val="clear"/> entries that CLEAR
        // inherited tab stops (Arabic UN notes clear 7 default stops). `add
        // <kind>` seeds p[1] with style/spacing/indent/rPr but never the tabs, so
        // the cleared stops reappeared and shifted tabbed footnote content
        // horizontally. Emit them the same way a body paragraph does (EmitTabStops
        // handles every val incl. "clear"); the per-paragraph loop below does the
        // same for the note's subsequent paragraphs.
        if (bodyParas.Count > 0 && bodyParas[0].Format.TryGetValue("tabs", out var noteTabs))
            EmitTabStops($"{targetNotePath}/p[1]", noteTabs, items);

        // BUG-R13A: coalesce hyperlink runs so a hyperlink inside a footnote/
        // endnote body round-trips as a typed `add hyperlink` (was dropped as a
        // flat `add r` carrying unsupported url/isHyperlink props).
        EmitContainerBodyRuns(word, firstParaRuns.Skip(noteSeedSkip).ToList(),
            $"{targetNotePath}/p[1]", items);

        // BUG-DUMP-R27-5: walk the remaining DIRECT block children in document
        // order. The first paragraph (note p[1], the ref-mark carrier) was
        // emitted by the `add <kind>` above, so skip the first "p" entry.
        // Target-side paragraph indices count only emitted paragraphs (`add
        // paragraph` builds /<kind>[N]/p[last()]); tables interleave via
        // EmitTable against the note host.
        bool firstParaSkipped = false;
        int targetParaOrdinal = 1; // p[1] already exists
        foreach (var (blockKind, paraNode, tblOrdinal) in blockOrder)
        {
            if (blockKind == "p")
            {
                if (!firstParaSkipped) { firstParaSkipped = true; continue; }
                if (paraNode == null) continue;
                var paraProps = FilterEmittableProps(paraNode.Format);
                paraProps.Remove("text"); // text carried by per-run emits below
                items.Add(new BatchItem
                {
                    Command = "add",
                    Parent = targetNotePath,
                    Type = "paragraph",
                    Props = paraProps.Count > 0 ? paraProps : null
                });
                targetParaOrdinal++;
                // BUG-DUMP-NOTE-TABSTOPS: carry this note paragraph's tab stops
                // (incl. clear-type) — same as p[1] above and the body path.
                if (paraNode.Format.TryGetValue("tabs", out var subParaTabs))
                    EmitTabStops($"{targetNotePath}/p[{targetParaOrdinal}]", subParaTabs, items);
                var runs = paraNode.Children.Where(c => IsRoundTrippableNoteRun(word, c)).ToList();
                EmitContainerBodyRuns(word, runs, $"{targetNotePath}/p[{targetParaOrdinal}]", items);
            }
            else // "tbl" — reuse the body table emitter against the note host.
            {
                // ctx is null here: EmitNoteReference has no BodyEmitContext,
                // and the ctx-driven paths in EmitTable (global //w:tbl ordinal
                // for cell-SDT raw-sets) only fire for containerPath=="/body".
                // A note-hosted table routes every cell through the typed emit.
                EmitTable(word, $"{sourceNotePath}/tbl[{tblOrdinal}]", tblOrdinal,
                    items, ctx: null, parentTablePath: null, containerPath: targetNotePath);
            }
        }
    }

    // BUG-DUMP-R27-5: enumerate a footnote/endnote's DIRECT block-level children
    // (top-level <w:p> and <w:tbl>) in document order from its raw XML. A
    // depth-tracked scan keeps paragraphs nested inside table cells (or nested
    // tables) from being counted as note-level blocks — the bug that flattened
    // a footnote table into out-of-range positional paragraph emits. The first
    // element open encountered is the <w:footnote>/<w:endnote> wrapper itself
    // (depth 0); its direct children are at depth 1.
    private static List<string> EnumerateNoteDirectChildren(string? noteXml)
    {
        var result = new List<string>();
        if (string.IsNullOrEmpty(noteXml)) return result;
        int depth = -1; // becomes 0 when the note wrapper opens
        bool seenWrapper = false;
        foreach (System.Text.RegularExpressions.Match m in
                 System.Text.RegularExpressions.Regex.Matches(
                     noteXml, @"<(/?)w:([A-Za-z]+)\b[^>]*?(/?)>"))
        {
            var closing = m.Groups[1].Value == "/";
            var name = m.Groups[2].Value;
            var selfClose = m.Groups[3].Value == "/";
            if (!seenWrapper)
            {
                if (!closing) { seenWrapper = true; depth = 0; }
                continue;
            }
            if (closing) { depth--; continue; }
            if (depth == 0 && (name == "p" || name == "tbl"))
                result.Add(name);
            if (!selfClose) depth++;
        }
        return result;
    }

    // Enumerate a block SDT's DIRECT sdtContent children (top-level <w:p> /
    // <w:tbl>) in document order from its raw XML, so the unwrap fallback can
    // address each by ordinal (`/sdt[N]/p[K]`, `/sdt[N]/tbl[K]`). The scan
    // anchors on the FIRST <w:sdtContent> open (the block content, after
    // sdtPr/sdtEndPr) and depth-tracks so a NESTED sdt's own paragraphs are
    // not counted as this SDT's block children. Mirrors EnumerateNoteDirectChildren.
    private static List<string> EnumerateSdtContentDirectChildren(string? sdtXml)
    {
        var result = new List<string>();
        if (string.IsNullOrEmpty(sdtXml)) return result;
        int depth = -1; // becomes 0 when <w:sdtContent> opens
        bool inContent = false;
        foreach (System.Text.RegularExpressions.Match m in
                 System.Text.RegularExpressions.Regex.Matches(
                     sdtXml, @"<(/?)w:([A-Za-z]+)\b[^>]*?(/?)>"))
        {
            var closing = m.Groups[1].Value == "/";
            var name = m.Groups[2].Value;
            var selfClose = m.Groups[3].Value == "/";
            if (!inContent)
            {
                if (!closing && name == "sdtContent") { inContent = true; depth = 0; }
                continue;
            }
            if (closing)
            {
                depth--;
                if (depth < 0) break; // </w:sdtContent>
                continue;
            }
            // BUG-R16C: a nested block <w:sdt> directly inside the unwrapped
            // sdtContent (e.g. a cover wrapper grouping data-bound title +
            // subtitle controls) must be surfaced too — otherwise the unwrap
            // emits only the sibling paragraphs/tables and the nested controls
            // (and their text) vanish on dump.
            if (depth == 0 && (name == "p" || name == "tbl" || name == "sdt"))
                result.Add(name);
            if (!selfClose) depth++;
        }
        return result;
    }

    // BUG-DUMP-ENDNOTE-ID: map a 1-based document-order user-note ordinal to the
    // real OOXML note Id. `query footnote`/`query endnote` returns user notes
    // (id > 0, separators excluded) in document order with id-qualified paths
    // (/endnote[@endnoteId=2]); this is the same set the reference cursor counts.
    // The Nth reference therefore corresponds to the Nth entry's Id. Falls back
    // to the ordinal itself (legacy id==ordinal assumption) when the path can't
    // be parsed or the ordinal is out of range — preserves the prior behaviour
    // for the well-formed Word-convention case rather than throwing.
    private static int ResolveUserNoteId(WordHandler word, string kind, int ordinal)
    {
        var notes = word.Query(kind);
        if (ordinal >= 1 && ordinal <= notes.Count)
        {
            var m = System.Text.RegularExpressions.Regex.Match(
                notes[ordinal - 1].Path, $@"@{kind}Id=(-?\d+)");
            if (m.Success && int.TryParse(m.Groups[1].Value, out var id))
                return id;
        }
        return ordinal;
    }

    // BUG-R12A(BUG3): a note-body run is round-trippable as a plain `add r` only
    // when it is a text-carrying run that is NOT the note reference mark
    // (<w:footnoteRef/> / <w:endnoteRef/>, which renders the superscript marker
    // and is recreated by AddFootnote/AddEndnote). Richer structure (drawings,
    // fields, nested notes) inside a note body is rare and out of scope — skip
    // such runs rather than mis-emit them as plain text. Mirrors
    // IsRoundTrippableCommentRun, plus the ref-mark exclusion.
    //
    // BUG-DUMP-ENDNOTE-ID: the ref-mark exclusion must reject only a *pure*
    // ref-mark run (the <w:*Ref/> with no body text). Word emits the ref mark
    // and the note text in SEPARATE runs, but some editors fuse them into a
    // single <w:r><w:*Ref/><w:t>body</w:t></w:r>. Rejecting any run that merely
    // *contains* the ref child dropped that fused run's entire body text — the
    // root of "endnote bodies silently dropped". Get's .Text already excludes
    // the ref mark (it contributes no <w:t>), and AddFootnote/AddEndnote rebuilds
    // the ref mark from scratch, so a fused run round-trips correctly as a plain
    // text run; only a text-less ref mark is dropped.
    private static bool IsRoundTrippableNoteRun(WordHandler word, DocumentNode run)
    {
        // Tab-only runs align note text after the reference mark; EmitCommentRun
        // round-trips them as `add r text="\t"`. Excluding them silently
        // de-indented every footnote that tabs before its content.
        if (run.Type == "tab") return true;
        if (run.Type != "run" && run.Type != "r") return false;
        var raw = word.GetElementXml(run.Path);
        if (!string.IsNullOrEmpty(raw)
            && (raw.Contains("footnoteRef", StringComparison.Ordinal)
                || raw.Contains("endnoteRef", StringComparison.Ordinal))
            && string.IsNullOrEmpty(run.Text))
            return false; // a pure reference mark — recreated by AddFootnote/AddEndnote
        return true;
    }

    // Emit a body-level SDT (Content Control). Simple SDTs (a single text run,
    // dropdown/combobox/date pickers) round-trip as a typed `add /body --type
    // sdt` carrying type/alias/tag/items/format + the visible text — all of
    // which AddSdt rebuilds. Without this, SDTs were silently dropped from dump
    // output (BUG-X2-06 / X2-3).
    //
    // Rich BLOCK SDTs are different: a Table of Contents, or any content control
    // wrapping multiple paragraphs / hyperlinks / fields / a table, carries
    // block structure the text-only path cannot express — it concatenates every
    // inner paragraph into one `text` run, collapsing a multi-line TOC into a
    // single line. Round-trip the whole <w:sdt> verbatim via raw-set instead,
    // inserted before the body's trailing sectPr so it lands at the same spot
    // the sequential `add /body` items build up to (AppendToParent inserts body
    // children before that sectPr). Same rationale as the theme/settings/
    // numbering raw emits: structured XML edited as a block, not per-property.
    private static void EmitSdt(WordHandler word, string sourcePath, List<BatchItem> items, BodyEmitContext ctx)
    {
        var rawXml = word.RawElementXml(sourcePath);
        // BUG-DUMP-COMMENT-IN-SDT: strip in-sdtContent comment-range markers from the
        // verbatim slice — they keep their SOURCE id while the comment is renumbered
        // dense + re-anchored via EmitComments/AddComment, leaving a dangling stale-id
        // marker pair that makes the rebuilt doc fail to open in Word (validate
        // dangling-reference). Mirrors the COMMENT-IN-MATH strip; the comment survives
        // through its typed re-anchor.
        if (rawXml != null) rawXml = WordHandler.StripVerbatimCommentMarkers(rawXml);
        if (!string.IsNullOrEmpty(rawXml) && IsRichBlockSdt(rawXml!))
        {
            // External relationship references (hyperlink r:id, image r:embed/
            // r:link) would dangle in the blank target — raw injection does not
            // recreate the matching rels. Ship the SDT through the inlined-parts
            // carrier instead (verbatim sdtXml + part{N}/ext{N} data, rel ids
            // rewritten on replay), same as the activex/diagram/vmlshape runs.
            // Only when a referenced part can't be resolved fall back to the
            // text emit and surface the loss.
            if (HasExternalRelRef(rawXml!))
            {
                var sdtData = word.GetSdtEmitData(sourcePath);
                if (sdtData != null)
                {
                    var carrierProps = PackInlinedPartsProps(sdtData);
                    // BUG-DUMP-COMMENT-IN-SDT: same strip as the verbatim raw-set path
                    // — the inlined-parts carrier ships sdtContent verbatim too.
                    carrierProps["sdtXml"] = WordHandler.StripVerbatimCommentMarkers(carrierProps["runXml"]);
                    carrierProps.Remove("runXml");
                    items.Add(new BatchItem
                    {
                        Command = "add",
                        Parent = "/body",
                        Type = "sdt",
                        Props = carrierProps,
                    });
                    // The carrier ships the whole SDT (including any <w:tbl> in its
                    // content) verbatim, without routing through EmitTable, so the
                    // shipped tables never bump ctx.TableOrdinalBox. At replay those
                    // tables still exist in document order and count toward the
                    // `(//w:tbl)[N]` XPath that later cell-SDT / tblGrid raw-sets
                    // resolve against — leaving the ordinal short makes every
                    // following table's selector land one (or more) tables early, so
                    // a sibling table's cell-SDT raw-set wraps the wrong cell's
                    // drawing in a spurious nested SDT that the next SDK re-save drops.
                    // Bump the box by the table count of the shipped sdtXml so the
                    // emitter's `(//w:tbl)` numbering stays in lockstep with replay.
                    // CONSISTENCY(tbl-ordinal): mirrors EmitTable's `++TableOrdinalBox[0]`.
                    if (ctx != null
                        && carrierProps.TryGetValue("sdtXml", out var shippedXml)
                        && !string.IsNullOrEmpty(shippedXml))
                    {
                        ctx.TableOrdinalBox[0] += System.Text.RegularExpressions.Regex
                            .Matches(shippedXml, "<w:tbl[ >]").Count;
                    }
                    return;
                }
                // Unreconstructable references (a header/footer rel inside an
                // SDT-wrapped sectPr, a chart): UNWRAP — emit the SDT's inner
                // block children through the normal body walk so the content,
                // its drawings and any inline section break survive. Only the
                // content-control wrapper itself is lost; warn deterministically
                // (mirrors the customXml wrapper-flattening contract). The old
                // flatten-to-text fallback dropped whole sections (a mid-
                // document portrait/landscape boundary vanished and every page
                // after it flipped orientation).
                ctx.Warnings.Add(new DocxUnsupportedWarning(
                    Element: "sdt.richContent",
                    Path: sourcePath,
                    Reason: "content control wrapper dropped on dump (rich block content references parts the sdt carrier cannot ship); its inner content is emitted unwrapped"));
                // Get on a block SDT surfaces NO children (the navigator does
                // not descend into sdtContent), so walk the raw sdtContent
                // block children by ordinal and emit each through its own
                // navigable path (`/sdt[N]/p[K]`, `/sdt[N]/tbl[K]`). This
                // preserves the inner paragraphs, their drawings AND any
                // inline section break (the landscape boundary that was
                // vanishing). Bail to the text emit only when the SDT exposes
                // no addressable block children at all.
                int sdtParaOrdinal = 0, sdtTblOrdinal = 0, sdtNestedOrdinal = 0;
                bool sdtEmittedAny = false;
                // A cross-paragraph field (a cached TOC inside this content
                // control) must NOT be re-emitted paragraph-by-paragraph — the
                // opener's fldChar(begin) has no matching end in its own
                // paragraph, so the typed emit collapses it and drops the first
                // cached entry. Raw-pass every paragraph of such a span verbatim,
                // mirroring the body walk's EmitCrossParagraphFieldMember.
                var sdtSpanEnd = new Dictionary<int, int>();
                foreach (var (s, e) in word.GetSdtContentCrossParagraphFieldSpanRanges(sourcePath))
                    sdtSpanEnd[s] = e;
                int? activeSdtSpanEnd = null;
                foreach (var kind in EnumerateSdtContentDirectChildren(rawXml!))
                {
                    if (kind == "p")
                    {
                        sdtParaOrdinal++;
                        if (activeSdtSpanEnd == null && sdtSpanEnd.TryGetValue(sdtParaOrdinal, out var spEnd))
                            activeSdtSpanEnd = spEnd;
                        if (activeSdtSpanEnd != null)
                        {
                            var rawP = word.GetElementXml($"{sourcePath}/p[{sdtParaOrdinal}]");
                            if (!string.IsNullOrEmpty(rawP))
                                items.Add(new BatchItem
                                {
                                    Command = "raw-set",
                                    Part = "/document",
                                    Xpath = "//w:body/w:sectPr",
                                    Action = "insertbefore",
                                    Xml = rawP
                                });
                            else
                                EmitParagraph(word, $"{sourcePath}/p[{sdtParaOrdinal}]", "/body", 1,
                                              items, autoPresent: false, ctx);
                            if (sdtParaOrdinal >= activeSdtSpanEnd.Value) activeSdtSpanEnd = null;
                        }
                        else
                        {
                            EmitParagraph(word, $"{sourcePath}/p[{sdtParaOrdinal}]", "/body", 1,
                                          items, autoPresent: false, ctx);
                        }
                        sdtEmittedAny = true;
                    }
                    else if (kind == "tbl")
                    {
                        sdtTblOrdinal++;
                        EmitTable(word, $"{sourcePath}/tbl[{sdtTblOrdinal}]", sdtTblOrdinal, items, ctx);
                        sdtEmittedAny = true;
                    }
                    else if (kind == "sdt")
                    {
                        // BUG-R16C: recurse into a nested block SDT so its content
                        // (e.g. a data-bound cover title/subtitle) survives the
                        // outer wrapper's unwrap. EmitSdt raw-sets the nested
                        // control verbatim at body level (preserving its
                        // dataBinding), so the bound text still renders.
                        sdtNestedOrdinal++;
                        EmitSdt(word, $"{sourcePath}/sdt[{sdtNestedOrdinal}]", items, ctx);
                        sdtEmittedAny = true;
                    }
                }
                if (!sdtEmittedAny)
                    EmitSdtTyped(word, sourcePath, "/body", items);
                return;
            }
            else
            {
                items.Add(new BatchItem
                {
                    Command = "raw-set",
                    Part = "/document",
                    Xpath = "//w:body/w:sectPr",
                    Action = "insertbefore",
                    Xml = rawXml
                });
                // CONSISTENCY(tbl-ordinal): this rich-block-SDT (no external rel)
                // is shipped verbatim WITHOUT routing its inner <w:tbl> through
                // EmitTable, so EmitTable's `++TableOrdinalBox` never counts them —
                // yet the later `(//w:tbl)[N]` cell-SDT/SdtRow/tblGrid raw-set
                // selectors count ALL tables in document order, including these.
                // Leaving the ordinal short made every following table's selector
                // land N tables early, dropping a locked SdtRow + dropdown-bound
                // cells. Bump by the shipped XML's table count — mirrors the
                // HasExternalRelRef carrier branch above (and the textbox carrier).
                if (ctx != null && !string.IsNullOrEmpty(rawXml))
                    ctx.TableOrdinalBox[0] += System.Text.RegularExpressions.Regex
                        .Matches(rawXml, "<w:tbl[ >]").Count;
                return;
            }
        }

        EmitSdtTyped(word, sourcePath, "/body", items);
    }

    // BUG-R11A(BUG1): block <w:sdt> that is a DIRECT CHILD of a table cell.
    // Mirrors body-level EmitSdt: rich block content (multi-paragraph / field /
    // table / drawing / line-break) round-trips verbatim via raw-set appended
    // into the just-emitted cell; everything text-shaped goes through the typed
    // `add sdt` path targeting the cell. Without this, the cell child walk in
    // EmitTable enumerated only paragraphs and nested tables, so a cell-nested
    // SDT (and its content) was silently dropped on dump → round-trip data loss.
    //
    // <paramref name="cellXPath"/> is the `(//w:tbl)[N]/w:tr[r]/w:tc[c]` selector
    // that resolves to the target cell at replay time (built from the document-
    // order table ordinal so it is stable regardless of later tables / nesting).
    // <paramref name="rawPart"/> is the host part ("/document" for body tables,
    // "/header[N]" / "/footer[N]" otherwise). <paramref name="cellHasContent"/>
    // decides prepend vs append so the SDT keeps document order relative to the
    // cell's auto-seeded leading paragraph.
    // Returns true when the SDT was raw-set into the cell AHEAD of the auto-seed
    // paragraph (the insert-after-tcPr branch), so the caller knows a spurious
    // empty seed paragraph is left behind and must be removed when the cell has
    // no real paragraph of its own (BUG-DUMP-R36-CELLSDT). All other paths
    // (append after existing content, typed `add sdt`, header/footer prepend,
    // external-rel flatten) consume or never create that seed, so return false.
    private static bool EmitCellSdt(WordHandler word, string sourcePath, string cellTargetPath,
                                    string cellXPath, string rawPart, bool cellHasContent,
                                    List<BatchItem> items, BodyEmitContext ctx)
    {
        var rawXml = word.RawElementXml(sourcePath);
        // BUG-DUMP-COMMENT-IN-SDT: strip in-sdtContent comment-range markers from the
        // verbatim slice — they keep their SOURCE id while the comment is renumbered
        // dense + re-anchored via EmitComments/AddComment, leaving a dangling stale-id
        // marker pair that makes the rebuilt doc fail to open in Word (validate
        // dangling-reference). Mirrors the COMMENT-IN-MATH strip; the comment survives
        // through its typed re-anchor.
        if (rawXml != null) rawXml = WordHandler.StripVerbatimCommentMarkers(rawXml);
        if (!string.IsNullOrEmpty(rawXml) && IsRichBlockSdt(rawXml!))
        {
            if (HasExternalRelRef(rawXml!))
            {
                // BUG-DUMP-CELLSDT-CARRIER: mirror the body-level EmitSdt path —
                // a cell content control wrapping a hyperlink or image used to
                // flatten to plain text (the rel-bearing raw XML can't be raw-set
                // without dangling). Ship it through the inlined-parts carrier
                // instead (verbatim sdtXml + part/ext data; rel ids rewritten on
                // replay), so the link/image and the control's rich structure
                // survive. Only fall back to the text flatten when a referenced
                // part genuinely can't be resolved.
                var sdtData = word.GetSdtEmitData(sourcePath);
                bool carrierHostIsCell = System.Text.RegularExpressions.Regex.IsMatch(
                    cellXPath, @"/w:tc(\[\d+\])?$");
                if (sdtData != null)
                {
                    var carrierProps = PackInlinedPartsProps(sdtData);
                    // BUG-DUMP-COMMENT-IN-SDT: same strip as the verbatim raw-set path
                    // — the inlined-parts carrier ships sdtContent verbatim too.
                    carrierProps["sdtXml"] = WordHandler.StripVerbatimCommentMarkers(carrierProps["runXml"]);
                    carrierProps.Remove("runXml");
                    items.Add(new BatchItem
                    {
                        Command = "add",
                        Parent = cellTargetPath,
                        Type = "sdt",
                        Props = carrierProps,
                    });
                    // The carrier's `add sdt` APPENDS the control (AppendToParent),
                    // landing it after the cell's auto-seed <w:p> when it is the
                    // leading content ([seed, sdt]). Drop that now-leading seed so
                    // the cell matches the source shape (SDT first); when the SDT
                    // is NOT leading (cellHasContent) it appends after real content
                    // and no seed remains to remove. Only genuine cell hosts have
                    // the auto-seed paragraph (header/footer roots do not).
                    if (carrierHostIsCell && !cellHasContent)
                        items.Add(new BatchItem
                        {
                            Command = "raw-set",
                            Part = rawPart,
                            Xpath = $"{cellXPath}/w:p[1]",
                            Action = "remove",
                        });
                    return false;
                }
                ctx.Warnings.Add(new DocxUnsupportedWarning(
                    Element: "sdt.richContent",
                    Path: sourcePath,
                    Reason: "content control in a table cell with rich block content AND external relationship references (hyperlinks/images) flattened to text on dump"));
            }
            else
            {
                // BUG-DUMP-R27-4: CT_Tc requires <w:tcPr> (when present) to be
                // the cell's FIRST child, before any block content. Prepending
                // the rich SDT to the cell landed it BEFORE <w:tcPr>
                // (<w:tc><w:sdt/><w:tcPr/>…) → "unexpected child element tcPr"
                // and an invalid file. For the empty-cell case, anchor the SDT
                // on the cell's auto-seeded leading paragraph with `insertbefore`
                // — AddTable always seeds exactly one <w:p> per cell, and CT_Tc
                // orders that <w:p> after any <w:tcPr>, so inserting before it
                // lands the SDT after tcPr (if present) and ahead of the seed,
                // preserving CT_Tc order and the source's "SDT is the cell's
                // leading content" shape regardless of whether the cell has a
                // tcPr. The append case (cell already has emitted content)
                // already lands after tcPr + that content.
                //
                // BUG-DUMP-CELLSDT-NOTCPR: an earlier revision targeted the cell's
                // <w:tcPr> with `insertafter`, assuming "the rebuilt cell always
                // carries a tcPr (AddTable seeds the cell width)". That stopped
                // being true once the grid width became canonical on <w:tblGrid>
                // and AddTable began emitting bare cells (<w:tc><w:p/></w:tc> with
                // no tcPr). A cell whose sole content is a rich block SDT then has
                // no tcPr, so `{cellXPath}/w:tcPr` matched nothing and replay threw
                // ("XPath matched no elements: …/w:tc[1]/w:tcPr"), dropping the SDT
                // entirely (round-trip data loss). Anchoring on the always-present
                // seed <w:p> instead is robust to tcPr presence.
                //
                // BUG-DUMP-R28-4: the insert-before-seed placement is a TABLE-CELL
                // rule and must fire ONLY when the host xpath actually resolves to
                // a <w:tc>. This helper is reused for header/footer-body block
                // SDTs (EmitHeaderFooter passes the /w:hdr or /w:ftr root as the
                // host xpath); a header/footer root has no auto-seeded cell
                // paragraph, so it keeps the plain prepend into the host root.
                // Gate on a genuine cell host (xpath ending in `…/w:tc` or
                // `…/w:tc[N]`).
                bool hostIsCell = System.Text.RegularExpressions.Regex.IsMatch(
                    cellXPath, @"/w:tc(\[\d+\])?$");
                if (cellHasContent)
                {
                    items.Add(new BatchItem
                    {
                        Command = "raw-set",
                        Part = rawPart,
                        Xpath = cellXPath,
                        Action = "append",
                        Xml = rawXml
                    });
                }
                else
                {
                    // BUG-DUMP-H86: anchor each leading SDT with insertbefore the
                    // host's auto-seed <w:p> (/w:p[1]) — for BOTH a table cell
                    // (AddTable seeds one <w:p> per cell) AND a header/footer root
                    // (the part is created with a seed <w:p>, removed later by
                    // EmitHeaderFooter's firstChildIsNonPara pass). Successive
                    // insertbefore raw-sets stack in document order
                    // ([SDT1, SDT2, SDT3, seed]); the original header/footer branch
                    // used bare `prepend` to the root, which REVERSES multiple SDTs
                    // ([SDT3, SDT2, SDT1]) — the header/footer-path gap left by the
                    // H85 cell fix. insertbefore lands ahead of the seed exactly like
                    // the old prepend for a single SDT, so single-SDT behavior is
                    // unchanged.
                    items.Add(new BatchItem
                    {
                        Command = "raw-set",
                        Part = rawPart,
                        Xpath = $"{cellXPath}/w:p[1]",
                        Action = "insertbefore",
                        Xml = rawXml
                    });
                    // The cell caller drops the now-unconsumed seed when the cell has
                    // no real paragraph (returns true → cellSdtLeftSeed); the
                    // header/footer caller runs its own seed removal and ignores the
                    // return.
                    if (hostIsCell) return true;
                }
                return false;
            }
        }

        EmitSdtTyped(word, sourcePath, cellTargetPath, items);
        return false;
    }

    // Shared typed `add sdt` emit. Whitelists the Get-canonical keys AddSdt
    // consumes plus the visible text; targets <paramref name="parentPath"/>
    // (/body for body-level SDTs, a cell path for cell-nested ones). AddSdt
    // accepts both Body and TableCell parents, so the same emit serves both.
    // BUG-DUMP-SDTPROPS: canonical Get keys the typed `add sdt` path forwards.
    // Shared by EmitSdtTyped (block SDT) and EmitInlineSdt (inline SDT) so both
    // round-trip the identical set of form-control properties. `editable` is a
    // Get readback (negation of `lock`); `id` is allocated at creation — neither
    // forwarded.
    internal static readonly string[] SdtTypedEmitKeys =
    {
        "type", "alias", "tag", "items", "format", "lock",
        // checkbox checked state — a bare-glyph checkbox now round-trips through
        // the typed path (see HasSpecialSdtTypeMarker); without this the state
        // would silently reset to unchecked on a dump->batch cycle.
        "checked",
        "placeholder", "placeholderText",
        "date.fullDate", "date.calendar", "date.lid", "date.storeMappedDataAs",
        "comboBox.lastValue", "dropDown.lastValue",
        // BUG-DUMP-R25-5: customXml data-store binding (xpath / storeItemID /
        // prefixMappings). Without these the control degrades to static on
        // round-trip. AddSdt rebuilds <w:dataBinding> from the three keys.
        "dataBinding.xpath", "dataBinding.storeItemID", "dataBinding.prefixMappings",
    };

    /// <summary>Stringify a Get-canonical sdt Format value for typed emit,
    /// lowercasing C# bool ToString() ("True"/"False" → "true"/"false"). Boolean
    /// props (checked / placeholder) surface as CLR bools from Get; emitting the
    /// capitalized form would split one concept into two vocabulary tokens for a
    /// consumer that reads the dump as text. AddSdt's IsTruthy accepts either.</summary>
    internal static string NormalizeSdtEmitValue(object v)
    {
        var s = v.ToString() ?? "";
        return s is "True" or "False" ? s.ToLowerInvariant() : s;
    }

    private static void EmitSdtTyped(WordHandler word, string sourcePath, string parentPath,
                                     List<BatchItem> items)
    {
        DocumentNode sdt;
        try { sdt = word.Get(sourcePath); }
        catch { return; }

        var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        // Whitelist Get-canonical keys that AddSdt consumes. `editable` is a
        // Get readback (negation of `lock`), the source-side `id` is allocated
        // at creation, so neither is forwarded.
        //
        // BUG-DUMP-SDTPROPS: forward the form-control sdtPr children the typed
        // emit previously dropped — `lock` (content-control locking), the
        // placeholder docPart/showing-placeholder flag, the date-picker selected
        // value + locale/calendar/store-as, and the combo/dropdown current
        // selection. Each has a matching AddSdt case; the Get reader surfaces the
        // canonical key (ReadSdtExtraProps + placeholder detection).
        foreach (var key in SdtTypedEmitKeys)
        {
            if (sdt.Format.TryGetValue(key, out var v) && v != null)
            {
                var s = NormalizeSdtEmitValue(v);
                if (s.Length > 0) props[key] = s;
            }
        }
        if (!string.IsNullOrEmpty(sdt.Text))
            props["text"] = sdt.Text!;

        items.Add(new BatchItem
        {
            Command = "add",
            Parent = parentPath,
            Type = "sdt",
            Props = props
        });
    }

    // A block SDT is "rich" when its content carries structure the text-only
    // typed emit cannot reproduce: more than one paragraph, more than one run,
    // any run-level rPr, or any hyperlink / complex field / table / drawing /
    // break / tab. Such SDTs round-trip verbatim via raw-set; everything else
    // (single plain text run, form-control pickers) stays on the introspectable
    // typed `add sdt` path.
    // BUG-DUMP-R27-4: an SDT whose sdtPr carries a special-type marker —
    // <w15:repeatingSection> / <w15:repeatingSectionItem> (the "repeat +" UI)
    // or <w:docPartObj> (a building-block / Quick-Part gallery registration) —
    // cannot round-trip through the typed `add sdt` path. The typed emit reads
    // only text/lock/placeholder/combo-dropdown sdtPr children and would
    // reclassify the control as a generic richtext SDT, silently dropping the
    // repeating-section structure and the gallery descriptors. Treat the marker
    // as "rich" so the whole <w:sdt> raw-sets verbatim (same passthrough the
    // nested-inline-SDT case uses), preserving the SDT BEHAVIOR. Namespace
    // prefixes are matched loosely (w15:/w14:/etc.) by local element name.
    internal static bool HasSpecialSdtTypeMarker(string sdtXml)
        => System.Text.RegularExpressions.Regex.IsMatch(
               sdtXml, @"<[A-Za-z0-9]+:repeatingSection(Item)?[ />]")
        || System.Text.RegularExpressions.Regex.IsMatch(
               sdtXml, @"<w:docPartObj[ />]");
    // A <w14:checkbox> content control is no longer forced to raw-set here: the
    // typed `add sdt --prop type=checkbox --prop checked=X` path now rebuilds the
    // sdtPr marker (BuildSdtCheckBox) AND the box glyph, and `checked` rides the
    // typed-emit whitelist (SdtTypedEmitKeys), so a checkbox whose content is the
    // bare glyph round-trips through the typed path. A checkbox carrying richer
    // content (a formatted glyph run with <w:rPr>, extra runs, nested markers,
    // …) still returns rich from the general IsRich*Sdt checks below and stays
    // raw-set for fidelity. (Historically checkbox short-circuited to raw here
    // because the typed add could not reproduce the type or checked state.)

    private static bool IsRichBlockSdt(string sdtXml)
    {
        if (HasSpecialSdtTypeMarker(sdtXml))
            return true;
        // <w:p> / <w:p attr...> — but not <w:pPr>, <w:pict>, <w:proofErr> (the
        // char after "w:p" must be a space or '>').
        if (System.Text.RegularExpressions.Regex.Matches(sdtXml, "<w:p[ >]").Count > 1)
            return true;
        // BUG-R12A(BUG1b): a single-paragraph block SDT whose content carries
        // multiple runs or any run-level formatting (bold/color/size/font/…)
        // cannot round-trip through the flat `add sdt text=` path — AddSdt seeds
        // one unformatted run from the concatenated text, so "FIRST"+"SECOND"
        // (2nd bold/red) comes back as a single plain "FIRSTSECOND" run. The
        // run-level richness check here was previously only applied to inline
        // (run-level) SDTs (IsRichInlineSdt); body/cell/header/footer BLOCK
        // SDTs flattened. Raw-set the SDT verbatim (no rels) so per-run rPr
        // survives. Restrict the run-count probe to CONTENT runs by counting
        // <w:r> opens — sdtPr/sdtEndPr carry no <w:r>, so no false positives.
        if (System.Text.RegularExpressions.Regex.Matches(sdtXml, "<w:r[ >]").Count > 1)
            return true;
        // Any run-level rPr inside a content run (the rPr sits under <w:r>; a
        // pPr's <w:rPr> paragraph-mark formatting is matched too, which is also
        // worth preserving verbatim and the typed path can't express it).
        if (sdtXml.Contains("<w:rPr", StringComparison.Ordinal))
            return true;
        // A content paragraph carrying a pStyle (e.g. a placeholder cover-title
        // SDT whose inner <w:p> is styled "Title") cannot round-trip through the
        // flat `add sdt text=` path — AddSdt seeds a default-styled paragraph, so
        // the pStyle is lost and the placeholder renders at body-text size and
        // top-of-page position instead of the styled title. Raw-set verbatim so
        // the inner paragraph style (and the showingPlcHdr placeholder) survive.
        if (sdtXml.Contains("<w:pStyle", StringComparison.Ordinal))
            return true;
        // BUG-DUMP-SDT-PPR: a content paragraph carrying direct paragraph-level
        // formatting (jc / ind / framePr / keepNext / spacing / cnfStyle / numPr /
        // suppressLineNumbers / the CJK kinsoku family / …) in its <w:pPr> cannot
        // round-trip through the flat `add sdt text=` path — AddSdt seeds a
        // default paragraph and drops ALL pPr. The pStyle/rPr triggers above only
        // cover styled or run/mark-formatted paragraphs; a plain-run paragraph with
        // rich pPr fell to the lossy path. Any <w:pPr> with a child element means
        // direct paragraph formatting is present → raw-set verbatim. (An empty
        // <w:pPr/> or <w:pPr></w:pPr> has no children and won't match.)
        if (System.Text.RegularExpressions.Regex.IsMatch(sdtXml, "<w:pPr>\\s*<w:"))
            return true;
        // BUG-DUMP-SDT-NESTED: a NESTED <w:sdt> (content control inside this one)
        // can't round-trip through the flat `add sdt text=` path — AddSdt seeds a
        // single plain run from the concatenated text, dropping the inner SDT
        // wrapper (its tag/id/type). The outer's own <w:sdt> opening tag is one
        // match; a second means a nested control → raw-set verbatim. (<w:sdtPr> /
        // <w:sdtContent> / <w:sdtEndPr> don't match "<w:sdt" + space/'>'.)
        if (System.Text.RegularExpressions.Regex.Matches(sdtXml, "<w:sdt[ >]").Count > 1)
            return true;
        // BUG-DUMP-H79: an SDT whose content carries a tracked change
        // (<w:del>/<w:ins>/<w:moveFrom>/<w:moveTo>) cannot round-trip through the
        // flat `add sdt text=` path — that path serializes only live text, so a
        // del-only content paragraph (no live runs) flattens to empty and the
        // deletion is silently dropped. None of the run/rPr/pPr triggers above
        // fire for a pure <w:del> paragraph (the <w:r> sits inside <w:del>, the
        // pPr is empty). Treat any tracked-change wrapper as rich → raw-set
        // verbatim. (Same meta-pattern as the complex-field-result del fix: a
        // tracked-change wrapper must force the verbatim path.)
        if (sdtXml.Contains("<w:del", StringComparison.Ordinal)
            || sdtXml.Contains("<w:ins", StringComparison.Ordinal)
            || sdtXml.Contains("<w:moveFrom", StringComparison.Ordinal)
            || sdtXml.Contains("<w:moveTo", StringComparison.Ordinal))
            return true;
        // BUG-DUMP-H94: an SDT whose content carries a range/anchor marker
        // (<w:bookmarkStart/End>, <w:commentRangeStart/End> / <w:commentReference>,
        // <w:permStart/End>) cannot round-trip through the flat `add sdt text=`
        // path — that path seeds only the text and omits all inner markers, so the
        // bookmark / comment-range / permission silently vanishes (balanced
        // start+end drop together, so no marker-imbalance tripwire). None of the
        // run/rPr/pPr triggers fire for a plain paragraph + a bookmark. Force the
        // verbatim raw-set path. Same classifier-gap pattern as the tracked-change
        // trigger above (H79) and the repeatingSection trigger (H84).
        if (sdtXml.Contains("<w:bookmark", StringComparison.Ordinal)
            || sdtXml.Contains("<w:comment", StringComparison.Ordinal)
            || sdtXml.Contains("<w:perm", StringComparison.Ordinal))
            return true;
        return sdtXml.Contains("<w:hyperlink", StringComparison.Ordinal)
            || sdtXml.Contains("<w:fldChar", StringComparison.Ordinal)
            || sdtXml.Contains("w:instrText", StringComparison.Ordinal)
            || sdtXml.Contains("<w:fldSimple", StringComparison.Ordinal)
            || sdtXml.Contains("<w:tbl", StringComparison.Ordinal)
            || sdtXml.Contains("<w:drawing", StringComparison.Ordinal)
            // BUG-DUMPR2: a single-paragraph SDT can still carry intra-run
            // structure the text-only typed emit can't reproduce — a line break
            // or tab. sdt.Text concatenates run text and drops <w:br/>/<w:tab/>,
            // so "a<w:br/>b" flattened to "ab". Treat their presence as rich so
            // the SDT round-trips verbatim via raw-set (no rels involved).
            || sdtXml.Contains("<w:br", StringComparison.Ordinal)
            || sdtXml.Contains("<w:tab", StringComparison.Ordinal)
            || sdtXml.Contains("<w:cr", StringComparison.Ordinal)
            // BUG-DUMP-H95: other text-less run-content elements the typed
            // `add sdt text=` path drops because they produce no <w:t> — a symbol
            // (<w:sym>), a positional tab (<w:ptab> — distinct from <w:tab>, so the
            // <w:tab> trigger above does NOT cover it), and the hyphen markers
            // (<w:noBreakHyphen> turns "co-op" into "coop" when lost; <w:softHyphen>
            // for completeness). Same text-less-content reason as <w:br>/<w:tab>/<w:cr>.
            || sdtXml.Contains("<w:sym", StringComparison.Ordinal)
            || sdtXml.Contains("<w:ptab", StringComparison.Ordinal)
            || sdtXml.Contains("<w:noBreakHyphen", StringComparison.Ordinal)
            || sdtXml.Contains("<w:softHyphen", StringComparison.Ordinal)
            // BUG-DUMP-EQUATION-SDT: an equation content control's math content
            // (<m:oMath>/<m:oMathPara>) lives in m: runs, not <w:r>, so the run
            // checks above miss it and the typed path dropped the equation. Treat
            // math content or the <w:equation/> sdtPr marker as rich → raw-set
            // verbatim. (Block-level equation SDTs mirror the inline fix.)
            || sdtXml.Contains("<m:oMath", StringComparison.Ordinal)
            || sdtXml.Contains("<w:equation", StringComparison.Ordinal);
    }

    // Raw injection of an <w:sdt> into the blank target preserves the element
    // verbatim but cannot recreate the package relationships its r:id/r:embed/
    // r:link attributes point at — those would dangle. Detect them so the
    // caller can fall back to the (lossy but valid) text emit.
    private static bool HasExternalRelRef(string xml)
        => xml.Contains("r:id=", StringComparison.Ordinal)
        || xml.Contains("r:embed=", StringComparison.Ordinal)
        || xml.Contains("r:link=", StringComparison.Ordinal);

    private static void EmitSection(WordHandler word, List<BatchItem> items)
    {
        var root = word.Get("/");
        // BUG-DUMP-PROTECTION-ENFORCE: a document can DEFINE a protection mode
        // (w:edit="forms") without ENFORCING it (w:enforcement="0") — Word then
        // renders the doc normally. The `protection` Set handler used to always
        // stamp w:enforcement=1, and this emitter dropped protectionEnforced as
        // "dump-only metadata", so an unenforced-forms source round-tripped to
        // ENFORCED forms → Word switched to form-fill mode and pushed every line
        // down a constant ~12px (a pervasive visual drift with no content change).
        // Keep protectionEnforced so the `protection`/`protectionEnforced` Set
        // cases can restore the source enforcement state. For protection="none"
        // there is nothing to enforce: drop both keys so round-trips stay clean.
        if (root.Format.TryGetValue("protection", out var protVal)
            && string.Equals(protVal?.ToString(), "none", StringComparison.OrdinalIgnoreCase))
        {
            root.Format.Remove("protection");
            root.Format.Remove("protectionEnforced");
        }
        var blankBaseline = _blankRootBaseline.Value;
        var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var (k, v) in root.Format)
        {
            bool include = RootScalarKeys.Contains(k);
            if (!include)
            {
                foreach (var pref in RootPrefixGroups)
                {
                    if (k.StartsWith(pref, StringComparison.OrdinalIgnoreCase))
                    {
                        include = true;
                        break;
                    }
                }
            }
            if (!include) continue;
            // docDefaults round-trips verbatim via EmitDocDefaultsRaw now —
            // skip the per-property emit here so the two paths don't fight
            // (and so source-absent slots aren't re-stamped from the blank).
            if (k.StartsWith("docDefaults.", StringComparison.OrdinalIgnoreCase)) continue;
            if (v == null) continue;
            var s = v switch { bool b => b ? "true" : "false", _ => v.ToString() ?? "" };
            if (s.Length == 0) continue;
            // Skip when the source's value already matches what BlankDocCreator
            // would stamp. Otherwise dump-then-replay leaves blank's value on
            // the target unchanged, but the SECOND dump picks it up (because
            // the value is now explicit in the part) and emits a `set /` row
            // dump-1 had skipped — losing idempotency. Symmetry: dump-2
            // applies the same rule and also skips. The existing
            // docDefaults.font.latin="" clear below is the inverse case
            // (blank's value is undesirable — actively clear it).
            if (blankBaseline.TryGetValue(k, out var blankVal)
                && string.Equals(blankVal, s, StringComparison.Ordinal))
            {
                continue;
            }
            props[k] = s;
        }
        // BUG-DUMP-PROTECTION-ENFORCE: protectionEnforced must be emitted even when
        // it matches the blank baseline (false). The `protection` Set writer defaults
        // enforcement to TRUE, and the typed `set / protection=...` op replays AFTER
        // the verbatim settings raw-set — so an unenforced-forms source (enforcement
        // ="0") needs an explicit protectionEnforced=false to stop the writer from
        // re-enforcing it (which flips Word into form-fill mode, shifting every line).
        // The baseline-skip above drops a false value (blank default is also false),
        // so force it into props here whenever a protection mode is present.
        if (root.Format.TryGetValue("protection", out var protForce)
            && !string.Equals(protForce?.ToString(), "none", StringComparison.OrdinalIgnoreCase)
            && root.Format.TryGetValue("protectionEnforced", out var enfForce) && enfForce != null)
        {
            props["protectionEnforced"] = enfForce is bool eb ? (eb ? "true" : "false")
                : enfForce.ToString() ?? "true";
        }
        // NOTE: docDefaults (fonts, size, lang, spacing, …) is no longer
        // emitted property-by-property here — EmitDocDefaultsRaw round-trips
        // the whole <w:docDefaults> block verbatim, which also handles the
        // "source omits a slot the blank stamped" pollution the old
        // per-property clears (bare-font rewrite, the BUG-X6-05 font.latin=""
        // clear) were patching one slot at a time.
        //
        // Page-geometry absence: when the source body sectPr OMITS <w:pgSz>
        // (and/or <w:pgMar>), Get returns no pageWidth/pageHeight/marginTop/…
        // keys, so nothing above stamps them — but the rebuild target is a
        // blank doc whose sectPr already carries the template's A4 pgSz +
        // default pgMar. Left untouched, the rebuild renders A4 while real
        // Word renders the pgSz-less source as its application default (US
        // Letter) → whole-document re-wrap. Emit an explicit remove signal
        // (`pageSize=none` / `pageMargin=none`; "none" is the established
        // sectPr-child remove sentinel) so the rebuilt sectPr also defers to
        // the app default. Independent per element — a source with pgSz but
        // no pgMar (or vice versa) is handled correctly. When the source HAS
        // the element the normal pageWidth/marginTop emit above carries it,
        // and no remove signal is emitted.
        // Page geometry round-trip fix: the loop above sourced pageWidth/
        // pageHeight/marginTop/… from Get's canonical cm strings, which round
        // twips to 2 decimals (1418 twips → "2.5cm"). Replaying that through
        // ParseTwips yields 1417 — a ±1-twip drift on every dump→batch cycle.
        // Overwrite each PRESENT geometry key with its native-twip integer
        // (bare numbers parse back as exact twips), so the rebuild's pgSz/pgMar
        // match the source byte-for-byte. Only keys already in `props` are
        // overwritten — the blank-baseline skip above and the pageSize=none /
        // pageMargin=none sentinels below stay in force.
        // BUG-DUMP-R29-PGSZ: emit each PRESENT geometry key's exact twips even when
        // the cm-based blank-baseline skip above dropped it. Get's canonical cm
        // string collapses near-equal widths to the same value (A4 source 11907
        // twips and the blank template's 11906 twips both render "21cm"), so the
        // skip wrongly treated the source as "same as blank" and emitted NO
        // pageWidth — the rebuild kept the blank's 11906. That 1-twip narrower
        // page is enough to flip a borderline line wrap, adding a line that
        // cascades into whole-document pagination drift. Sourcing straight from
        // the sectPr twips (bare numbers parse back as exact twips) guarantees the
        // rebuilt pgSz/pgMar match the source byte-for-byte. rawTwips only holds
        // keys whose sectPr child is present, so a source that OMITS pgSz/pgMar
        // still falls through to the pageSize=none / pageMargin=none sentinels.
        var rawTwips = word.BodySectionPageGeometryTwips();
        foreach (var (gk, gv) in rawTwips)
        {
            props[gk] = gv;
        }
        // pgBorders fold: Get emits pgBorders.<side> + pgBorders.<side>.sz/
        // .color/.space as separate keys (mirrors pbdr.* / border.*). Set's
        // pgborders.<side> case parses a single semicolon-encoded
        // STYLE;SIZE;COLOR;SPACE value, so fold the sub-keys into the bare
        // side key and drop them. pgBorders.offsetFrom passes through verbatim
        // (it's a standalone Set key). Without folding, the 3-segment sub-keys
        // hit UNSUPPORTED on replay and the per-side weight/color/space were
        // lost — the page border collapsed to the box default.
        FoldPgBordersProps(props);
        var (hasPgSz, hasPgMar) = word.BodySectionPageGeometryPresence();
        if (!hasPgSz) props["pageSize"] = "none";
        if (!hasPgMar) props["pageMargin"] = "none";
        // BUG-DUMP-R31-1: a childless <w:sectPr/> (no pgSz, no pgMar) is NOT the
        // same as a missing sectPr — real Word renders the two at different page
        // widths. The pageSize=none / pageMargin=none sentinels above would let
        // the apply's drop-the-empty-sectPr path remove the element entirely,
        // collapsing "empty sectPr present" into "no sectPr". Emit an explicit
        // sectPr=present marker whenever the source body actually carries a
        // sectPr element, so the apply keeps a bare <w:sectPr/> instead of
        // dropping it. Absent on a truly sectPr-less source — there the drop
        // path stays in force.
        if (word.BodyHasSectionProperties()) props["sectPr"] = "present";
        // sectPrChange round-trip — fold the source's <w:sectPrChange>
        // format-revision marker (author/date) into the section `set /` op as
        // a revision.type=format + revision.author (+ .date) triplet, mirroring
        // FoldRevisionIntoProps for tblPrChange/trPrChange/tcPrChange (see
        // WordBatchEmitter.Table.cs). The before-snapshot is intentionally NOT
        // reconstructed — Set's section path writes an EMPTY-snapshot
        // <w:sectPrChange> (same shape as the table/paragraph markers whose
        // baseline can't be recovered from a dump). Without this fold the
        // marker was the only non-Run *PrChange dropped on round-trip.
        if (FoldRevisionIntoProps(root.Format, "sectPrChange", props))
        {
            // Carry the stable w:id too (FoldRevisionIntoProps handles only
            // author/date — shared with the table path, which doesn't surface
            // an id). Section readback surfaces sectPrChange.id; preserving it
            // keeps the marker's identity stable across the round-trip.
            var sectChangeId = TryStringFormat(root.Format, "sectPrChange.id");
            if (sectChangeId != null) props["revision.id"] = sectChangeId;
        }
        if (props.Count == 0) return;
        items.Add(new BatchItem
        {
            Command = "set",
            Path = "/",
            Props = props
        });
    }

    // Fold pgBorders.<side>.sz/.color/.space sub-keys into the bare
    // pgBorders.<side> key as a STYLE;SIZE;COLOR;SPACE value (the form Set's
    // pgborders.<side> case parses via ParseBorderValue). Mirrors the pbdr.* /
    // border.* fold in WordBatchEmitter.Filters.cs. pgBorders.offsetFrom is a
    // standalone Set key and is left untouched.
    private static void FoldPgBordersProps(Dictionary<string, string> props)
    {
        var fold = new Dictionary<string, (string? style, string? sz, string? color, string? space)>(
            StringComparer.OrdinalIgnoreCase);
        var subKeys = new List<string>();
        foreach (var (key, val) in props)
        {
            if (!key.StartsWith("pgBorders.", StringComparison.OrdinalIgnoreCase)) continue;
            var parts = key.Split('.');
            // parts[0]=pgBorders, parts[1]=side|offsetFrom
            if (parts.Length < 2) continue;
            // offsetFrom is a flat key — not a per-side border. Leave it alone.
            // BUG-DUMP-R44-5: zOrder / display are likewise flat pgBorders attrs,
            // not per-side borders — exclude them from the per-side fold so they
            // pass through verbatim to the section set step.
            if (parts.Length == 2 &&
                (string.Equals(parts[1], "offsetFrom", StringComparison.OrdinalIgnoreCase)
                 || string.Equals(parts[1], "zOrder", StringComparison.OrdinalIgnoreCase)
                 || string.Equals(parts[1], "display", StringComparison.OrdinalIgnoreCase)))
                continue;
            var side = $"{parts[0]}.{parts[1]}"; // pgBorders.top
            fold.TryGetValue(side, out var cur);
            if (parts.Length == 2)
            {
                cur.style = val;
            }
            else if (parts.Length == 3)
            {
                switch (parts[2].ToLowerInvariant())
                {
                    case "sz": cur.sz = val; break;
                    case "color": cur.color = val; break;
                    case "space": cur.space = val; break;
                }
                subKeys.Add(key); // 3-segment sub-keys get dropped after folding
            }
            fold[side] = cur;
        }
        foreach (var sk in subKeys) props.Remove(sk);
        foreach (var (side, folded) in fold)
        {
            if (folded.style == null) continue;
            var sz = folded.sz ?? "";
            var col = folded.color ?? "";
            var sp = folded.space ?? "";
            var v = folded.style;
            if (folded.sz != null || folded.color != null || folded.space != null)
                v += ";" + sz;
            if (folded.color != null || folded.space != null)
                v += ";" + col;
            if (folded.space != null)
                v += ";" + sp;
            props[side] = v;
        }
    }

    // BUG-DUMP-STYLE-TABS: render a style's pPr tab-stop list (the `tabs` Format
    // value — IEnumerable<Dictionary>) into the POS[:ALIGN[:LEADER]] comma-joined
    // shorthand that AddStyle's ApplyTabsShorthand consumes, so style tab stops
    // round-trip on the `add style` op instead of via unresolvable per-stop
    // `add tab parent=/styles/<id>` rows. Mirrors EmitTabStops' field reads.
    internal static string BuildTabsShorthand(object? tabsVal)
    {
        if (tabsVal is not System.Collections.Generic.IEnumerable<Dictionary<string, object?>> list)
            return "";
        var segs = new List<string>();
        foreach (var t in list)
        {
            if (!t.TryGetValue("pos", out var p) || p == null) continue;
            var pos = p.ToString();
            if (string.IsNullOrEmpty(pos)) continue;
            var val = t.TryGetValue("val", out var v) && v != null ? v.ToString() ?? "" : "";
            var leader = t.TryGetValue("leader", out var l) && l != null ? l.ToString() ?? "" : "";
            // ApplyTabsShorthand defaults an empty ALIGN to left, so "pos::leader"
            // is valid when a leader is present without an explicit alignment.
            string seg = !string.IsNullOrEmpty(leader) ? $"{pos}:{val}:{leader}"
                       : !string.IsNullOrEmpty(val) ? $"{pos}:{val}"
                       : pos!;
            segs.Add(seg);
        }
        return string.Join(",", segs);
    }

    private static void EmitStyles(WordHandler word, List<BatchItem> items, bool recursiveStyleDecomp)
    {
        // Use query() rather than walking Get("/styles").Children — the
        // positional /styles/style[N] children Get returns are not
        // addressable on the Get side (style paths resolve by id, not by
        // index). Query produces id-based paths and excludes docDefaults.
        var styles = word.Query("style");
        // STYLE-RAW-FALLBACK: scalar Format keys (basedOn / spaceAfter /
        // font / size / …) cannot express a TABLE style's visual formatting:
        // its style-level <w:tblPr> (borders, band sizes, cell margins), its
        // <w:tblStylePr> conditional-formatting blocks (firstRow / lastRow /
        // band1Vert / …), and table-level <w:shd>/<w:tcPr>/<w:trPr>. A table
        // that draws its borders/shading/banding from a table style (no inline
        // <w:tblBorders> on the table itself) therefore lost ALL visual
        // formatting on round-trip — the rebuilt style emitted only scalars,
        // dropping tblBorders/tblStylePr/shd, and Word rendered it as plain
        // borderless text. Give table styles a raw-set replace fallback that
        // round-trips the whole <w:style> element verbatim — exactly the
        // pattern docDefaults / theme / settings already use. The scalar `add
        // style` still runs first (creating the style + handling id collisions
        // via AddStyle's upsert/suffix path); the raw-set then swaps the
        // freshly-added <w:style> for the source's verbatim copy, so no
        // double-apply and no scalar/raw drift. Mirrors EmitDocDefaultsRaw.
        var rawStyleByMatchAttr = BuildRawTableStyleMap(word);
        // BUG-R18C: styleId → verbatim XML of the LAST occurrence, for ids that
        // appear more than once. Word renders a duplicate styleId via its last
        // definition; raw-set-replace the first-occurrence scalar emit with it.
        var lastDuplicateStyleXml = BuildLastDuplicateStyleMap(word);
        // Blank-baseline cleanup: BlankDocCreator always stamps a Normal
        // style (for Word render parity — Calibri 11pt, 1.08x
        // line). When the source has no entry for styleId="Normal",
        // skipping the emit leaks the blank's stamped Normal into the
        // replay target — dump-2's dump then emits it as a phantom
        // `add /styles Normal`, breaking idempotency. Always prepend a
        // remove-Normal so target's styles end up matching source's
        // (idempotent: Remove of a missing style is a soft success).
        // When source HAS Normal, EmitStyles below recreates it via the
        // builtin-name upsert path; the redundant remove is harmless and
        // keeps the wire format independent of source/blank divergence.
        bool sourceHasNormal = styles.Any(s =>
            string.Equals(s.Format.TryGetValue("id", out var v) ? v?.ToString() : null,
                          "Normal", StringComparison.Ordinal));
        if (!sourceHasNormal)
        {
            items.Add(new BatchItem
            {
                Command = "remove",
                Path = "/styles/Normal",
            });
        }
        // Dedupe by styleId. A styleId is effectively a key — OOXML requires
        // it unique — but real-world sources (third-party editors / merged docs) carry
        // duplicates (e.g. 88 <w:style> elements, 58 unique ids). Word itself
        // tolerates this by keeping the FIRST occurrence and ignoring the rest
        // (it opens the file fine). Mirror that: emit each styleId once. Without
        // this, every duplicate replayed as an `add style` that failed the id
        // uniqueness check. (Get-by-id also resolves all duplicates to the first
        // style anyway, so the extra emits were redundant copies.)
        var seenStyleIds = new HashSet<string>(StringComparer.Ordinal);
        foreach (var stub in styles)
        {
            // CONSISTENCY(slash-in-style-id): style ids/names containing '/'
            // produce paths like /styles/Style/With/Slash that the path
            // parser splits on. Get fails. Fall back to the Query stub —
            // we lose pPr/rPr details but at least the style stub
            // (id/name/type/basedOn) round-trips, instead of dropping the
            // style entirely (BUG BT-3).
            DocumentNode full;
            try { full = word.Get(stub.Path); }
            catch { full = stub; }
            var props = FilterEmittableProps(full.Format);
            // Ensure id is present (Add requires it for /styles target).
            if (!props.ContainsKey("id") && !props.ContainsKey("styleId"))
            {
                if (props.TryGetValue("name", out var n)) props["id"] = n;
                else continue;
            }
            var emitId = props.GetValueOrDefault("id") ?? props.GetValueOrDefault("styleId");
            if (!string.IsNullOrEmpty(emitId) && !seenStyleIds.Add(emitId))
                continue; // duplicate styleId — keep first, skip the rest (Word's behavior)
            // BUG-DUMP-STYLE-TABS: a style's pPr tab stops must round-trip via the
            // `tabs=` shorthand prop on the `add style` op — NOT as separate
            // `add tab parent=/styles/<id>` rows. Unlike a paragraph (/body/p[N]
            // resolves as a tab-add parent), `/styles/<id>` is not navigable for
            // tab insertion, so the per-stop ops failed ("Path not found:
            // /styles/TextBox") and the style's tab strip was dropped. AddStyle
            // already consumes `tabs=` via ApplyTabsShorthand; build the shorthand
            // here so FilterEmittableProps' drop of the (non-stringable) tabs list
            // is compensated inline on the style op itself.
            if (!props.ContainsKey("tabs") && !props.ContainsKey("tabstops")
                && full.Format.TryGetValue("tabs", out var styleTabsForProp))
            {
                var tabsShorthand = BuildTabsShorthand(styleTabsForProp);
                if (!string.IsNullOrEmpty(tabsShorthand)) props["tabs"] = tabsShorthand;
            }
            // BUG-X6-03: built-in style ids (Normal / Heading1-9 / Title /
            // …) collide with the blank template's reservations on a
            // fresh batch target. AddStyle is now idempotent for those
            // specific ids (upsert: drop existing + re-add). For non-
            // built-in ids the strict "already exists" check still
            // applies. Emit `add` uniformly so the wire format stays a
            // simple `add`-only stream regardless of style provenance.
            // STYLE-RAW-FALLBACK: the verbatim <w:style> XML that the raw-set
            // path replaces with — table styles' tblPr/tblStylePr/shd/trPr/tcPr
            // and any rPr/pPr child the scalar emit drops. Keyed by emitId so an
            // id collision/suffix still lands on the right element.
            // BUG-R18C: a duplicate styleId resolves (in Word) to its LAST
            // occurrence; prefer that verbatim definition over the table-style
            // map (first) and over the scalar emit (which read the first).
            string? rawStyleReplace = null;
            if (!string.IsNullOrEmpty(emitId)
                && lastDuplicateStyleXml.TryGetValue(emitId, out var dupXml))
                rawStyleReplace = dupXml;
            else if (!string.IsNullOrEmpty(emitId)
                && rawStyleByMatchAttr.TryGetValue(emitId, out var rawStyleXml))
                rawStyleReplace = rawStyleXml;

            // RECURSIVE-STYLE-DECOMP (opt-in via OFFICECLI_RECURSIVE_STYLE_DECOMP):
            // instead of a verbatim raw-set replace, decompose the <w:style>
            // subtree into typed `add` ops — a shell `add style` (identity attrs
            // + name child) plus one `add <child>` per descendant element. Falls
            // back to the raw-set replace when the subtree carries content the
            // typed path can't round-trip losslessly (external rel, unknown
            // namespace, mixed text, pathological depth) — see
            // TryDecomposeStyleChildren. The whole-element raw-set stays the
            // safety net; this only shrinks it where decomposition is provably
            // lossless. Relies on the generic add appending repeatable same-name
            // children (e.g. multiple tblStylePr) rather than collapsing them.
            List<BatchItem>? recursiveOps = null;
            if (recursiveStyleDecomp && rawStyleReplace != null && !string.IsNullOrEmpty(emitId))
                recursiveOps = TryDecomposeStyleChildren(rawStyleReplace, $"/styles/{emitId}");

            if (recursiveOps != null)
            {
                // Shell add: only the style's identity + own attributes + name.
                // Every other child (basedOn, uiPriority, pPr, rPr, tblPr,
                // tblStylePr, …) is rebuilt by the recursive ops below, so the
                // scalar formatting props are intentionally dropped here.
                var shellProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                foreach (var k in s_styleShellProps)
                    if (props.TryGetValue(k, out var pv))
                    {
                        // BUG-DUMP-STYLE-EMPTY-NAME: a <w:style> with no <w:name>
                        // child surfaces as name="" in props. Emitting it makes the
                        // shell `add style` materialize a spurious <w:name w:val=""/>
                        // the source never had — breaking the recursive path's
                        // exact-element-multiset guarantee. (The legacy raw-set path
                        // hides this because its verbatim whole-style replace
                        // overwrites the shell's empty name; the recursive path has no
                        // such replace.) Skip the empty name so nameless styles —
                        // typically auto table sub-styles — round-trip unchanged.
                        if (k == "name" && string.IsNullOrEmpty(pv)) continue;
                        shellProps[k] = pv;
                    }
                items.Add(new BatchItem
                {
                    Command = "add",
                    Parent = "/styles",
                    Type = "style",
                    Props = shellProps
                });
                items.AddRange(recursiveOps);
            }
            else
            {
                items.Add(new BatchItem
                {
                    Command = "add",
                    Parent = "/styles",
                    Type = "style",
                    Props = props
                });
                // BUG-X4-T1 / BUG-DUMP-STYLE-TABS: style tab stops are folded into
                // the `tabs=` prop above — the old per-stop `add tab` emit failed
                // to resolve and is retired for styles.
                if (rawStyleReplace != null)
                {
                    items.Add(new BatchItem
                    {
                        Command = "raw-set",
                        Part = "/styles",
                        Xpath = $"/w:styles/w:style[@w:styleId='{emitId}']",
                        Action = "replace",
                        Xml = rawStyleReplace
                    });
                }
            }
        }
    }

    // Props copied onto the shell `add style` in RECURSIVE-STYLE-DECOMP mode:
    // the style's identity + its own <w:style> attributes + the name child.
    // All formatting children are rebuilt by the recursive child ops.
    private static readonly string[] s_styleShellProps =
        ["id", "styleId", "type", "name", "default", "customStyle"];

    // Recursive style decomposition: emit each <w:style> as typed `add` ops
    // rather than a verbatim raw-set replace (residue falls back to raw-set).
    // ON by default; set OFFICECLI_RECURSIVE_STYLE_DECOMP=0 to fall back to the
    // legacy whole-style raw-set path. Settable (internal) so tests can select
    // the path deterministically without depending on process-env / static-init
    // timing under parallel test execution.
    internal static bool RecursiveStyleDecomp =
        Environment.GetEnvironmentVariable("OFFICECLI_RECURSIVE_STYLE_DECOMP") != "0";

    // RECURSIVE-NUMBERING-DECOMP: emit the <w:numbering> part as typed `add`
    // ops (one `add w:abstractNum`/`add w:num` per definition + recursive
    // children) instead of a verbatim raw-set replace. ON by default; set
    // OFFICECLI_RECURSIVE_NUMBERING_DECOMP=0 to fall back to the legacy
    // whole-part raw-set path. Settable (internal) so tests can pin the path
    // deterministically without depending on process-env / static-init timing.
    internal static bool RecursiveNumberingDecomp =
        Environment.GetEnvironmentVariable("OFFICECLI_RECURSIVE_NUMBERING_DECOMP") != "0";

    // Well-known OOXML namespace → canonical prefix, the inverse of the add
    // path's CommonNamespaces. A namespace absent here is treated as residue
    // (the generic add can't resolve an unknown prefix), so the caller raw-sets.
    private static readonly Dictionary<string, string> s_nsToPrefix = new(StringComparer.Ordinal)
    {
        ["http://schemas.openxmlformats.org/wordprocessingml/2006/main"] = "w",
        ["http://schemas.openxmlformats.org/officeDocument/2006/relationships"] = "r",
        ["http://schemas.openxmlformats.org/drawingml/2006/main"] = "a",
        ["http://schemas.openxmlformats.org/markup-compatibility/2006"] = "mc",
        ["http://schemas.openxmlformats.org/officeDocument/2006/math"] = "m",
        // NOTE: Office Word extension wordml namespaces (w14/w15/w16cid/…) are
        // deliberately NOT listed here. They appear in real numbering/styles as
        // attributes (w15:restartNumberingAfterBreak on w:abstractNum, w14:paraId,
        // …) — but the strict OOXML validator REQUIRES those extension attributes
        // to be covered by the part-root's mc:Ignorable, which the typed
        // child-add path does not reproduce (it only attaches children to the
        // blank's root). Emitting them as typed props therefore yields a
        // schema-INVALID part. Until part-root mc:Ignorable round-trips, an
        // element carrying an extension attribute is treated as residue, so the
        // whole part falls back to the verbatim raw-set (which keeps mc:Ignorable
        // and validates). See TryDecomposeNumbering.
    };

    private const int StyleDecompMaxDepth = 40;

    // Decompose a verbatim <w:style> element into typed `add` child ops under
    // stylePath (the shell `add style` creates the element + its <w:name>).
    // Returns null on any residue the typed path can't round-trip — the caller
    // then emits the verbatim raw-set replace instead. The source tree is
    // finite and acyclic; the depth cap guards pathological nesting / overflow.
    private static List<BatchItem>? TryDecomposeStyleChildren(string styleXml, string stylePath)
    {
        if (string.IsNullOrEmpty(styleXml) || !styleXml.StartsWith("<")) return null;
        System.Xml.Linq.XElement styleEl;
        DocumentFormat.OpenXml.Wordprocessing.Style sdkStyle;
        try { styleEl = System.Xml.Linq.XElement.Parse(styleXml); }
        catch { return null; }
        // Parse the same <w:style> through the SDK so each child carries its
        // schema-typed identity: a child the style context can't hold parses as
        // OpenXmlUnknownElement, which TryEmitElementAdd treats as residue. Any
        // parse failure → whole-part raw-set (the existing safety net).
        try { sdkStyle = new DocumentFormat.OpenXml.Wordprocessing.Style(styleXml); }
        catch { return null; }
        var wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var ops = new List<BatchItem>();
        var counts = new Dictionary<string, int>(StringComparer.Ordinal);
        var xChildren = styleEl.Elements().ToList();
        var sdkChildren = sdkStyle.Elements().ToList();
        if (xChildren.Count != sdkChildren.Count) return null;
        for (int i = 0; i < xChildren.Count; i++)
        {
            var child = xChildren[i];
            var sdkChild = sdkChildren[i];
            if (child.Name.LocalName != sdkChild.LocalName) return null;
            // <w:name> is created by the shell `add style` (from the name prop);
            // skip it so it isn't added twice. Its localName is unique, so the
            // ordinals of the other children are unaffected.
            if (child.Name.LocalName == "name" && child.Name.NamespaceName == wNs)
                continue;
            counts.TryGetValue(child.Name.LocalName, out var c);
            counts[child.Name.LocalName] = c + 1;
            var childPath = $"{stylePath}/{child.Name.LocalName}[{c + 1}]";
            if (!TryEmitElementAdd(child, sdkChild, stylePath, childPath, ops, 0))
                return null;
        }
        return ops;
    }

    // Decompose a verbatim <w:numbering> element into typed `add` ops: one
    // `add w:abstractNum` / `add w:num` (/ `add w:numIdMacAtCleanup`) per direct
    // child, each recursing into its full subtree via TryEmitElementAdd. No
    // shell step is needed — unlike <w:style> (whose <w:name> identity child is
    // created by the curated `add style`), an abstractNum/num is fully described
    // by its attributes + children, so the generic prefixed-add path rebuilds it
    // verbatim. Children replay in source order; the cardinality-aware generic
    // add keeps abstractNum* before num* (first num has no num sibling → AddChild
    // places it in CT_Numbering schema order). Returns null on any residue (a
    // picture-bullet numPicBullet carries a VML r:id; unknown namespaces) — the
    // caller then ships the whole part verbatim via raw-set.
    //
    // Root attributes (mc:Ignorable, extra xmlns) on <w:numbering> are NOT
    // reproduced — the children attach to the blank's existing root. This is
    // vestigial-safe by construction: a prefix listed in mc:Ignorable that is
    // actually USED (a w14:/wp14: element or attribute) would make
    // TryEmitElementAdd return false → whole-part raw-set fallback (which keeps
    // mc:Ignorable). So whenever this typed path SUCCEEDS, mc:Ignorable lists
    // only unused prefixes and its loss changes nothing — same class as the
    // SDK's w:left→w:start / xmlns canonicalization the dump already accepts.
    private static List<BatchItem>? TryDecomposeNumbering(string numberingXml)
    {
        if (string.IsNullOrEmpty(numberingXml) || !numberingXml.StartsWith("<")) return null;
        System.Xml.Linq.XElement numEl;
        DocumentFormat.OpenXml.Wordprocessing.Numbering sdkNum;
        try { numEl = System.Xml.Linq.XElement.Parse(numberingXml); }
        catch { return null; }
        try { sdkNum = new DocumentFormat.OpenXml.Wordprocessing.Numbering(numberingXml); }
        catch { return null; }
        var ops = new List<BatchItem>();
        var counts = new Dictionary<string, int>(StringComparer.Ordinal);
        var xChildren = numEl.Elements().ToList();
        var sdkChildren = sdkNum.Elements().ToList();
        if (xChildren.Count != sdkChildren.Count) return null;
        for (int i = 0; i < xChildren.Count; i++)
        {
            var child = xChildren[i];
            var sdkChild = sdkChildren[i];
            if (child.Name.LocalName != sdkChild.LocalName) return null;
            counts.TryGetValue(child.Name.LocalName, out var c);
            counts[child.Name.LocalName] = c + 1;
            var childPath = $"/numbering/{child.Name.LocalName}[{c + 1}]";
            if (!TryEmitElementAdd(child, sdkChild, "/numbering", childPath, ops, 0))
                return null;
        }
        return ops.Count > 0 ? ops : null;
    }

    // Emit `add <localName> parent=<parentPath>` for el (attributes as
    // namespace-prefixed props), then recurse into its children. childPath is
    // el's own resolved path (for addressing its children). Returns false on
    // residue.
    private static bool TryEmitElementAdd(
        System.Xml.Linq.XElement el, DocumentFormat.OpenXml.OpenXmlElement sdkEl,
        string parentPath, string elPath, List<BatchItem> ops, int depth)
    {
        if (depth > StyleDecompMaxDepth) return false;
        if (!s_nsToPrefix.TryGetValue(el.Name.NamespaceName, out var prefix))
            return false; // unknown element namespace → residue
        // BUG-DUMP-STYLE-TCPR (generalized): an element is only decomposable into
        // a typed `add` if the generic add can REBUILD it in this parent context.
        // The authority on that is the SDK schema itself: TryCreateTypedElement
        // (the generic-add builder) returns null exactly when the element resolves
        // to an OpenXmlUnknownElement under the reconstructed parent, replaying as
        // "Unknown element type 'w:X'". We get the same verdict for free: `sdkEl`
        // is the SAME element as `el`, parsed by the SDK under its real (typed)
        // parent — a style-context <w:tcPr> is StyleTableCellProperties, a
        // style-context <w:rPr> is StyleRunProperties, and each rejects the
        // children the flat-XML tree carries (w:tcBorders under the former,
        // w:rtl under the latter) by parsing them as OpenXmlUnknownElement.
        // Treat any such element as residue → the recursion unwinds to
        // TryDecomposeStyleChildren returning null → the caller ships the whole
        // <w:style> verbatim via raw-set. This replaces the former per-element
        // (w:tcBorders-only) block: instead of listing each unreconstructable
        // element as it is discovered (w:tcBorders, then w:rtl, then the next
        // RTL/CJK style child), the SDK's own type table draws the line once, so
        // no future style child of this class silently breaks replay. Everything
        // the generic add CAN rebuild in a style context (tblPr/tblCellMar,
        // tblStylePr/rPr, pPr, …) parses as a typed element and still decomposes
        // — verified by RecursiveStyleDecompTests + StyleDecompUnknownProbeTests.
        if (sdkEl is DocumentFormat.OpenXml.OpenXmlUnknownElement)
            return false;
        // Direct (non-whitespace) text content has no typed `add` representation.
        foreach (var node in el.Nodes())
            if (node is System.Xml.Linq.XText t && !string.IsNullOrWhiteSpace(t.Value))
                return false;
        var props = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var a in el.Attributes())
        {
            if (a.IsNamespaceDeclaration) continue;
            var ans = a.Name.NamespaceName;
            string key;
            if (string.IsNullOrEmpty(ans))
            {
                key = a.Name.LocalName; // unqualified attribute → bare key
            }
            else if (s_nsToPrefix.TryGetValue(ans, out var ap))
            {
                if (ap == "r") return false; // external relationship → would dangle
                key = $"{ap}:{a.Name.LocalName}";
            }
            else
            {
                return false; // unknown attribute namespace → residue
            }
            props[key] = a.Value;
        }
        ops.Add(new BatchItem
        {
            Command = "add",
            Parent = parentPath,
            Type = $"{prefix}:{el.Name.LocalName}",
            Props = props
        });
        // Walk the flat-XML children in lockstep with the SDK-parsed children so
        // each recursion carries the SDK's schema verdict for that node. Both
        // enumerations exclude text/whitespace and preserve source order (unknown
        // elements included), so they align 1:1. Any desync — different length or
        // a mismatched localName at the same slot — means the SDK parsed the
        // subtree differently than the flat XML implies; treat that as residue
        // (whole-part raw-set) rather than emit a possibly-wrong add.
        var xChildren = el.Elements().ToList();
        var sdkChildren = sdkEl.Elements().ToList();
        if (xChildren.Count != sdkChildren.Count) return false;
        var counts = new Dictionary<string, int>(StringComparer.Ordinal);
        for (int i = 0; i < xChildren.Count; i++)
        {
            var child = xChildren[i];
            var sdkChild = sdkChildren[i];
            if (child.Name.LocalName != sdkChild.LocalName) return false;
            counts.TryGetValue(child.Name.LocalName, out var c);
            counts[child.Name.LocalName] = c + 1;
            var childPath = $"{elPath}/{child.Name.LocalName}[{c + 1}]";
            if (!TryEmitElementAdd(child, sdkChild, elPath, childPath, ops, depth + 1))
                return false;
        }
        return true;
    }

    // STYLE-RAW-FALLBACK helper: parse the source styles.xml once and return a
    // map from styleId → verbatim <w:style> XML, for ALL styles.
    //
    // Originally restricted to TABLE styles (whose tblPr/tblStylePr/shd/trPr/
    // tcPr have no scalar Format representation). But the scalar emit also
    // silently drops rPr/pPr children it has no key for — e.g. a paragraph
    // style whose rPr carries <w:bdr> (a run border box): the dump emitted
    // `shading=` from the sibling <w:shd> but no border key, so a Heading with
    // a colored border box round-tripped as a plain filled heading, reflowing
    // the whole document (SSIM 0.69). Rather than chase every missing rPr/pPr
    // child key (the same hardcoded-allowlist class fixed verbatim for the ¶
    // mark and docDefaults), round-trip EVERY style's <w:style> element
    // verbatim. The scalar `add style` still runs first (creating the style +
    // built-in id upsert / collision suffix); the raw-set then swaps it for the
    // source's exact copy, so no scalar/raw drift and no dropped children. The
    // keying id is each style's own w:styleId — callers match it against the id
    // the `add` step used.
    private static Dictionary<string, string> BuildRawTableStyleMap(WordHandler word)
    {
        var map = new Dictionary<string, string>(StringComparer.Ordinal);
        string stylesXml;
        try { stylesXml = word.Raw("/styles"); }
        catch { return map; }
        if (string.IsNullOrEmpty(stylesXml) || !stylesXml.StartsWith("<")) return map;
        try
        {
            var doc = System.Xml.Linq.XDocument.Parse(stylesXml);
            var wNs = (System.Xml.Linq.XNamespace)"http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            foreach (var styleEl in doc.Root?.Elements(wNs + "style") ?? Enumerable.Empty<System.Xml.Linq.XElement>())
            {
                var idAttr = styleEl.Attribute(wNs + "styleId");
                var styleId = idAttr?.Value;
                if (string.IsNullOrEmpty(styleId)) continue;
                // Dedupe: keep the first occurrence, matching EmitStyles' own
                // first-wins styleId dedup (Word tolerates duplicate ids).
                if (map.ContainsKey(styleId)) continue;
                map[styleId] = StripUnusedNsDeclarations(styleEl).ToString(System.Xml.Linq.SaveOptions.DisableFormatting);
            }
        }
        catch { return new Dictionary<string, string>(StringComparer.Ordinal); }
        return map;
    }

    // BUG-R18C: a styles.xml with a DUPLICATE styleId (two <w:style> elements
    // sharing one id — common when a template merge leaves a built-in stub
    // ahead of the customized definition) does not render via the FIRST
    // occurrence. Word resolves a duplicate styleId to the LAST definition (the
    // customization wins), but EmitStyles' scalar emit reads the style via
    // Get-by-id, which Navigation resolves to the FIRST element — so a
    // customized Heading1 (border, before/after spacing, bold, theme colour)
    // sitting in the second occurrence was emitted as the plain first stub.
    // Headings then lost their spacing on rebuild and every page's body
    // reflowed upward.
    //
    // Return styleId → verbatim XML of the LAST occurrence, ONLY for styleIds
    // that appear more than once. EmitStyles raw-set-replaces the just-added
    // (first-occurrence) style with this last-occurrence definition, so the
    // rebuilt style matches what Word actually renders. Single-occurrence
    // styles are left to the scalar emit (unchanged).
    private static Dictionary<string, string> BuildLastDuplicateStyleMap(WordHandler word)
    {
        var empty = new Dictionary<string, string>(StringComparer.Ordinal);
        string stylesXml;
        try { stylesXml = word.Raw("/styles"); }
        catch { return empty; }
        if (string.IsNullOrEmpty(stylesXml) || !stylesXml.StartsWith("<")) return empty;
        try
        {
            var doc = System.Xml.Linq.XDocument.Parse(stylesXml);
            var wNs = (System.Xml.Linq.XNamespace)"http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            var counts = new Dictionary<string, int>(StringComparer.Ordinal);
            var lastXml = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (var styleEl in doc.Root?.Elements(wNs + "style") ?? Enumerable.Empty<System.Xml.Linq.XElement>())
            {
                var styleId = styleEl.Attribute(wNs + "styleId")?.Value;
                if (string.IsNullOrEmpty(styleId)) continue;
                counts[styleId] = counts.GetValueOrDefault(styleId) + 1;
                lastXml[styleId] = StripUnusedNsDeclarations(styleEl).ToString(System.Xml.Linq.SaveOptions.DisableFormatting);
            }
            var dups = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (var kv in counts)
                if (kv.Value > 1) dups[kv.Key] = lastXml[kv.Key];
            return dups;
        }
        catch { return empty; }
    }

    private sealed class NoteCursor { public int Index; }
}
