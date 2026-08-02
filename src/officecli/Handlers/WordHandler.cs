// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class WordHandler : IDocumentHandler, Rendering.IRenderModelHost
{
    private readonly WordprocessingDocument _doc;

    object? Rendering.IRenderModelHost.RenderModel => _doc;
    private readonly string _filePath;
    private HashSet<string> _usedParaIds = new(StringComparer.OrdinalIgnoreCase);
    private int _nextParaId = 0x100000;
    public int LastFindMatchCount { get; internal set; }

    // StyleId → Style index, lazily built to make style-chain resolution O(1)
    // per hop. Before this, ResolveNumPrFromStyle / ResolveEffectiveRunProperties
    // / GetParagraphListStyle and the HtmlPreview/Query style walks each did a
    // LINEAR Elements<Style>().FirstOrDefault(StyleId==id) PER basedOn hop —
    // O(paragraphs × chainDepth × totalStyles), which HANGS on heavily templated
    // docs (deep basedOn chains × thousands of styles).
    //
    // MUTATION-SAFE: the cache validates against the live <w:styles> element's
    // reference identity ONLY (O(1) — a child-element COUNT check would be O(n)
    // in the SDK's linked-list ChildElements and re-introduce the per-hop O(n)
    // we are eliminating). A freshly created/rebuilt StyleDefinitionsPart swaps
    // the reference and self-invalidates. Same-instance mutations (Add appends a
    // Style, Remove drops one) keep the reference, so those two sites explicitly
    // call InvalidateStyleIndex(). `set /styles/X` only mutates a style's
    // PROPERTIES (never its id), so it needs no invalidation — the index maps
    // StyleId→Style by reference and the same object is still the right target.
    private Dictionary<string, Style>? _styleByIdCache;
    private Styles? _styleByIdCacheOwner;

    /// <summary>Drop the StyleId→Style index. Called by the (few) paths that add
    /// or remove a &lt;w:style&gt; on the EXISTING styles element.</summary>
    private void InvalidateStyleIndex()
    {
        _styleByIdCache = null;
        _styleByIdCacheOwner = null;
    }

    /// <summary>
    /// O(1) StyleId lookup through a lazily-built, reference-validated index.
    /// Replaces the per-hop linear <c>Elements&lt;Style&gt;().FirstOrDefault(
    /// s =&gt; s.StyleId?.Value == id)</c> scan used throughout style-chain
    /// resolution. Returns null when no styles part exists or no style matches.
    /// On a duplicate StyleId (malformed doc) the FIRST in document order wins —
    /// matching the original FirstOrDefault semantics.
    /// </summary>
    private Style? FindStyleById(string? styleId)
    {
        if (styleId == null) return null;
        var styles = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles;
        if (styles == null) return null;

        if (_styleByIdCache == null || !ReferenceEquals(styles, _styleByIdCacheOwner))
        {
            var dict = new Dictionary<string, Style>(StringComparer.Ordinal);
            foreach (var s in styles.Elements<Style>())
            {
                var id = s.StyleId?.Value;
                // First-wins on duplicate id (FirstOrDefault parity).
                if (id != null && !dict.ContainsKey(id)) dict[id] = s;
            }
            _styleByIdCache = dict;
            _styleByIdCacheOwner = styles;
        }

        return _styleByIdCache.TryGetValue(styleId, out var found) ? found : null;
    }

    // Number of elements a no-slash selector Set matched and mutated (Sheet1!row[...]).
    // Read by the CLI/resident to echo the multi-element change count.
    public int LastSelectorSetCount { get; internal set; }

    // Backing FileStream — mirrors the PPT pattern. Opening via a shared
    // FileStream (FileShare.Read in editable mode) lets external readers
    // observe the file while the handler is alive, which is required for
    // mid-session `save` snapshots to be useful to third-party consumers
    // (issue #114). The package writes through the stream; the on-disk
    // bytes lag _doc until _doc.Save() runs.
    private FileStream? _backingStream;
    // Editable sessions open _doc over this in-memory copy of the package so
    // saves swap the file atomically (temp + File.Replace) instead of rewriting
    // it in place. Null for read-only sessions (they never write). See
    // AtomicWriteBack / OfficeCli.Core.AtomicPackageWriter.
    private MemoryStream? _packageStream;

    // Whole-part payloads (docProps/core|app|custom.xml) staged by
    // RawReplaceWholePart and written straight into the saved zip after the
    // package closes (FlushPendingWholeParts). The Open-XML SDK does not
    // reliably flush docProps mutations made mid-session on a stream-opened
    // editable package — a typed/stream write is silently reverted to the
    // open-time content by _doc.Save(). Rewriting the zip entries post-close
    // (the same mechanism NormalizeSelfClosingInDocx uses for document.xml)
    // sidesteps that entirely. Keyed by zip entry name (no leading slash).
    private Dictionary<string, string>? _pendingWholeParts;

    // Part root elements mutated by a raw-set during a DeferSave batch. The
    // per-op rootElement.Save() (a whole-part re-serialize) and the four global
    // id sweeps are skipped during defer — a document with thousands of raw-set
    // replaces against one large part (e.g. a 4070-style styles.xml fixed up
    // verbatim, 1224 raw-sets) otherwise pays O(raw-sets × part-size) in
    // redundant serialization plus O(raw-sets × doc) in id scans. Each touched
    // root is Saved once, and the id sweeps run once, at FinalizeDeferredIds.
    private HashSet<OpenXmlPartRootElement>? _deferredRawSetRoots;

    /// <summary>
    /// Props that the most recent Add() call could not consume. Surfaced to
    /// the CLI layer so silent-drops on the curated surface (e.g.
    /// `add /styles --prop font.eastAsia=...`) become visible warnings
    /// instead of "Added" lies. Reset at the start of each Add.
    /// </summary>
    public List<string> LastAddUnsupportedProps { get; internal set; } = new();

    /// <summary>
    /// Advisory warnings from the most recent Add() call (e.g. unknown
    /// style id referenced but stored as-is). Surfaced to the CLI layer
    /// as stderr WARNING lines, non-fatal. Reset at the start of each Add.
    /// </summary>
    public List<string> LastAddWarnings { get; internal set; } = new();

    /// <summary>
    /// Advisory warnings from the most recent Set() call (e.g. unknown
    /// style id referenced as-is). Surfaced to the CLI layer as stderr
    /// WARNING lines, non-fatal — kept out of the unsupported-prop list so
    /// the write still counts as applied. Reset at the start of each Set.
    /// </summary>
    public List<string> LastSetWarnings { get; internal set; } = new();

    /// <summary>
    /// R4-bt-1: when the most recent Set() MOVED the target element to a new
    /// position (an equation mode switch unwraps/wraps the m:oMath, changing its
    /// canonical path), this carries the NEW resolvable path so the CLI reports
    /// "Updated &lt;new path&gt;" instead of the old path that no longer
    /// resolves. Null when no move occurred. Reset at the start of each Set.
    /// </summary>
    public string? LastSetNewPath { get; internal set; }

    /// <summary>
    /// LaTeX commands/environments the most recent Add()/Set() equation parse
    /// did not recognize and silently rendered as literal text. Surfaced to the
    /// CLI layer as <c>unrecognized_latex_command</c> warnings (exit code 2),
    /// mirroring the <c>unsupported_property</c> UX — lenient accept is kept
    /// (the equation is still written). Reset at the start of each Add/Set.
    /// </summary>
    public List<string> LastUnrecognizedLatex { get; internal set; } = new();

    /// <summary>
    /// Set true by Add/Set/Remove/RawSet, consumed by Save/Dispose to decide
    /// whether to stamp <c>docProps/custom.xml</c> with an OfficeCLI audit
    /// trail. Pure Get/Query sessions leave this false and never touch the
    /// file's metadata.
    /// </summary>
    internal bool Modified { get; set; }

    /// <summary>
    /// When true, per-mutation <c>Document.Save()</c> calls are skipped — the
    /// in-memory DOM stays authoritative and is serialized once at Dispose
    /// (AutoSave) / explicit flush. Set by the batch driver around a replay so
    /// N mutations cost O(N) instead of O(N²) (each Save re-serializes the whole
    /// growing part). Single-command paths leave this false and save eagerly.
    /// </summary>
    public bool DeferSave { get; set; }

    /// <summary>
    /// When true, Dispose releases the document WITHOUT serializing the
    /// in-memory DOM — the backing stream is closed first so the package's
    /// autosave has nowhere to write. Used by atomic-batch rollback: the
    /// on-disk file (flushed to the pre-batch state) stays authoritative and
    /// the poisoned DOM is simply dropped. Only meaningful when no
    /// intra-session flush has happened since the state to preserve.
    /// </summary>
    public bool DiscardOnDispose { get; set; }

    /// <summary>
    /// Serialize the main document part to the backing store, unless
    /// <see cref="DeferSave"/> is set (batch replay defers to one save at the
    /// end). Every mutation path (Add/Set/Remove/Move/Swap) routes its
    /// per-operation <c>Document.Save()</c> through here so the eager-vs-deferred
    /// decision lives in ONE place — N batch mutations cost one serialize, not N.
    /// </summary>
    private void SaveDoc()
    {
        if (!DeferSave) _doc.MainDocumentPart?.Document?.Save();
    }

    /// <summary>
    /// Enumerate every <see cref="OpenXmlPart"/> in the package (transitive
    /// walk via the SDK's own <c>GetAllParts</c> extension) yielding each
    /// part's zip-URI (<c>OpenXmlPart.Uri.OriginalString</c>). Used by the
    /// batch emitter's auxiliary-parts scan to surface warnings for parts the
    /// dump surface does not round-trip (customXml, glossary, webSettings,
    /// fontTable, embedded fonts, modern-comment metadata, user docProps).
    /// </summary>
    internal IEnumerable<string> EnumeratePartUris()
    {
        // Two complementary sources:
        //   1. SDK part graph (GetAllParts) — only reachable via the
        //      relationship graph; misses orphan parts but matches every
        //      real Word file produced by Office / python-docx / WPS.
        //   2. Raw zip entries — catches orphan parts (fuzzer-injected
        //      content, partial-rel files dropped by failed save), parts
        //      whose `_rels/.rels` reference disappeared, and anything the
        //      SDK refuses to surface as a typed part.
        // Union them so the aux-parts scan can warn on both cases. Skips
        // duplicates inside the consumer.
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var part in _doc.GetAllParts())
        {
            var u = part.Uri.OriginalString;
            if (seen.Add(u)) yield return u;
        }
        // Raw zip walk — opens the file independently. Use FileShare.ReadWrite
        // because the SDK's _backingStream already holds the file handle in
        // editable mode (FileShare.Read on our side). Falling back to the SDK
        // package via reflection is also possible (RawXmlHelper does this for
        // TryReadByZipUri), but the raw zip is simpler and sees the full
        // on-disk truth.
        if (File.Exists(_filePath))
        {
            List<string> zipEntries;
            try
            {
                using var fs = File.Open(_filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using var zip = new System.IO.Compression.ZipArchive(fs, System.IO.Compression.ZipArchiveMode.Read);
                zipEntries = zip.Entries.Select(e => e.FullName).ToList();
            }
            catch { yield break; }
            foreach (var name in zipEntries)
            {
                if (string.IsNullOrEmpty(name)) continue;
                var u = name.StartsWith("/") ? name : "/" + name;
                if (seen.Add(u)) yield return u;
            }
        }
    }

    /// <summary>
    /// Surface the <c>docProps/custom.xml</c> property names (if any).
    /// Used by the batch emitter to detect user-defined custom document
    /// properties whose values are silently dropped by dump (only
    /// <c>OfficeCLI.*</c> values are auto-restamped on save; user-authored
    /// props vanish). Returns an empty list if no custom part exists.
    /// </summary>
    internal IReadOnlyList<string> EnumerateCustomDocPropertyNames()
    {
        var part = _doc.CustomFilePropertiesPart;
        if (part?.Properties == null) return Array.Empty<string>();
        var names = new List<string>();
        foreach (var p in part.Properties.Elements<DocumentFormat.OpenXml.CustomProperties.CustomDocumentProperty>())
        {
            var n = p.Name?.Value;
            if (!string.IsNullOrEmpty(n)) names.Add(n!);
        }
        return names;
    }

    // One section's header/footer references for the batch emitter.
    internal sealed record SectionHfRef(
        bool IsFinal,
        // replay-side `/section[N]` ordinal (1-based document order among
        // INLINE sectPr, i.e. those carried in a paragraph mark); 0 for the
        // body-final sectPr (addressed as "/").
        int ReplayOrdinal,
        // type ("default"/"first"/"even") → header part Get path "/header[N]"
        List<(string Type, string PartPath)> Headers,
        List<(string Type, string PartPath)> Footers);

    // Enumerate EVERY section's header/footer references in document order,
    // INCLUDING inline sectPr nested inside an SDT (which `query section` /
    // the body-paragraph /section[N] walk does not surface). The batch
    // emitter needs this because a dump unwraps an SDT-wrapped section into a
    // normal body paragraph on replay, so its sectPr joins the document-order
    // /section[N] sequence — and a header attributed by the SDT-blind
    // ordinal would land on the wrong section. Resolves each reference's
    // rel id to the part's /header[N] / /footer[N] Get path (mainPart part
    // index + 1), matching the readback in BuildSectionNode.
    internal List<SectionHfRef> EnumerateSectionHeaderFooterRefs()
    {
        var result = new List<SectionHfRef>();
        var mainPart = _doc.MainDocumentPart;
        var body = mainPart?.Document?.Body;
        if (mainPart == null || body == null) return result;
        var headerParts = mainPart.HeaderParts.ToList();
        var footerParts = mainPart.FooterParts.ToList();

        List<(string, string)> ResolveRefs<TRef>(SectionProperties sp,
            Func<TRef, string?> idOf, Func<TRef, string?> typeOf,
            Func<string, OpenXmlPart?> partFor, List<OpenXmlPart> parts, string kind)
            where TRef : OpenXmlElement
        {
            var refs = new List<(string, string)>();
            foreach (var r in sp.Elements<TRef>())
            {
                var id = idOf(r);
                if (string.IsNullOrEmpty(id)) continue;
                var type = typeOf(r) ?? "default";
                try
                {
                    var part = partFor(id!);
                    var idx = part != null ? parts.IndexOf(part) : -1;
                    if (idx >= 0) refs.Add((type, $"/{kind}[{idx + 1}]"));
                }
                catch { /* dangling rel — skip */ }
            }
            return refs;
        }

        int inlineOrdinal = 0;
        // Descendants<SectionProperties>() yields all sectPr in document
        // order, including those inside SDTs / cells. The body-final sectPr
        // is a direct Body child; every other is carried in a ParagraphProperties.
        foreach (var sp in body.Descendants<SectionProperties>())
        {
            bool isFinal = sp.Parent is Body;
            int ord = 0;
            if (!isFinal) ord = ++inlineOrdinal;
            var headers = ResolveRefs<HeaderReference>(sp,
                r => r.Id?.Value, r => r.Type?.InnerText,
                id => mainPart.GetPartById(id) as DocumentFormat.OpenXml.Packaging.HeaderPart,
                headerParts.Cast<OpenXmlPart>().ToList(), "header");
            var footers = ResolveRefs<FooterReference>(sp,
                r => r.Id?.Value, r => r.Type?.InnerText,
                id => mainPart.GetPartById(id) as DocumentFormat.OpenXml.Packaging.FooterPart,
                footerParts.Cast<OpenXmlPart>().ToList(), "footer");
            if (headers.Count > 0 || footers.Count > 0)
                result.Add(new SectionHfRef(isFinal, ord, headers, footers));
        }
        return result;
    }

    public WordHandler(string filePath, bool editable)
    {
        _filePath = filePath;
        var share = editable ? FileShare.Read : FileShare.ReadWrite;
        var access = editable ? FileAccess.ReadWrite : FileAccess.Read;
        _backingStream = new FileStream(filePath, FileMode.Open, access, share);
        try
        {
            if (editable)
            {
                // Crash-atomic saves: work against an in-memory copy of the
                // package so the on-disk file is only ever swapped atomically
                // (temp + File.Replace via AtomicWriteBack), never rewritten in
                // place — a process death mid-write otherwise truncates the
                // original with no fallback. Read-only sessions keep streaming
                // from the FileStream: they never write.
                _backingStream.Position = 0;
                _packageStream = new MemoryStream();
                _backingStream.CopyTo(_packageStream);
                _packageStream.Position = 0;
            }
            _doc = WordprocessingDocument.Open((Stream?)_packageStream ?? _backingStream, editable);
            WordStrictAttributeSanitizer.Sanitize(_doc);
            if (editable)
            {
                EnsureAllParaIds();
                EnsureDocPropIds();
                EnsureSdtIds();
                EnsureBookmarkIds();
            }
        }
        catch
        {
            // A failed open must not leak the backing FileStream — the
            // factory's repair-and-retry paths (FixXmlEncoding /
            // StripDanglingPackageRels) reopen the file for in-place fixes
            // and would hit "file is being used by another process".
            _doc?.Dispose();
            _packageStream?.Dispose();
            _backingStream.Dispose();
            throw;
        }
    }

    /// <summary>
    /// Resolve a picture-run path to the embedded image's bytes and content
    /// type. Returns null if the path doesn't point at a Drawing-bearing
    /// run, or the run carries no resolvable rId/embed target.
    ///
    /// <para>
    /// Used by <c>WordBatchEmitter</c> to round-trip pictures through batch
    /// dumps — the bytes are encoded as a data URI in the emitted
    /// `src=` prop and re-imported via <c>ImageSource.Resolve</c> on replay.
    /// </para>
    /// </summary>
    /// <summary>
    /// Returns true if the run at <paramref name="runPath"/> wraps a chart
    /// (c:chart inside a Drawing's graphicData). WordBatchEmitter uses this to
    /// distinguish chart-bearing runs from picture/OLE/background runs that
    /// also surface as type="picture" in Get — without this, an unsupported
    /// drawing's failed image extraction would consume the next chart spec
    /// and render at the wrong paragraph.
    /// </summary>
    public bool IsChartRun(string runPath)
    {
        var segments = ParsePath(runPath);
        var element = NavigateToElement(segments);
        if (element is not Run run) return false;
        var drawing = run.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Drawing>();
        if (drawing == null) return false;
        return drawing
            .Descendants<DocumentFormat.OpenXml.Drawing.Charts.ChartReference>()
            .Any();
    }

    /// <summary>
    /// Outer XML of the element at <paramref name="path"/>. WordBatchEmitter
    /// uses this as a raw-XML fallback for content that has no typed Add
    /// path — wps:wsp background shapes being the motivating case. Returns
    /// null if the path doesn't resolve.
    /// </summary>
    public string? GetElementXml(string path)
    {
        try
        {
            var segments = ParsePath(path);
            var element = NavigateToElement(segments);
            return element?.OuterXml;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Concatenated OuterXml of the contiguous run of sibling elements from
    /// <paramref name="firstPath"/> to <paramref name="lastPath"/> (inclusive).
    /// The dump→batch nested-field / rich-field-result raw-set walks a field's
    /// begin..end slice; that slice may interleave bookmarkStart / bookmarkEnd
    /// children whose query paths (indexed by w:id) don't round-trip through
    /// NavigateToElement (positional indexing). Resolving only the begin and end
    /// runs — which always navigate — and then iterating the parent's actual
    /// child elements between them captures every interleaved element verbatim,
    /// regardless of whether each has an individually-navigable path. Returns
    /// null when either endpoint fails to resolve or they aren't siblings.
    /// </summary>
    public string? GetSiblingRangeXml(string firstPath, string lastPath)
    {
        try
        {
            var first = NavigateToElement(ParsePath(firstPath));
            var last = NavigateToElement(ParsePath(lastPath));
            if (first == null || last == null) return null;
            if (!ReferenceEquals(first.Parent, last.Parent) || first.Parent == null) return null;
            var sb = new System.Text.StringBuilder();
            bool collecting = false;
            foreach (var child in first.Parent.ChildElements)
            {
                if (ReferenceEquals(child, first)) collecting = true;
                if (collecting) sb.Append(child.OuterXml);
                if (ReferenceEquals(child, last)) return sb.Length > 0 ? sb.ToString() : null;
            }
            // last never encountered after first — order inverted or detached.
            return null;
        }
        catch
        {
            return null;
        }
    }

    public (byte[] Bytes, string ContentType)? GetImageBinary(string runPath)
    {
        // Parse + navigate via the same machinery Get/Set use so paraId
        // anchors and positional indices behave consistently.
        var segments = ParsePath(runPath);
        var element = NavigateToElement(segments);
        if (element is not Run run) return null;

        var drawing = run.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Drawing>();
        if (drawing == null) return null;

        var blip = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().FirstOrDefault();
        var embedId = blip?.Embed?.Value;
        if (string.IsNullOrEmpty(embedId)) return null;

        // CONSISTENCY(host-part-rel): mirror the AddPicture host-part lookup
        // — image part may be attached to a header/footer part rather than
        // the main document part, depending on where the run lives.
        var hostPart = ResolveImageHostPart(run);
        try
        {
            var part = hostPart.GetPartById(embedId);
            using var src = part.GetStream();
            using var ms = new MemoryStream();
            src.CopyTo(ms);
            return (ms.ToArray(), part.ContentType);
        }
        catch
        {
            return null;
        }
    }

    // dump→batch round-trip carrier for a Word OLE object run. The dump emits
    // these so AddOle can rebuild the embedded part + icon + VML frame without
    // an external src file. EmbeddedBytes is the payload exactly as stored
    // (raw package, or CFB-wrapped Ole10Native) — fed back verbatim.
    internal sealed record OleEmitData(
        byte[] EmbeddedBytes, string OleKind, string EmbeddedContentType, string EmbeddedExt,
        byte[]? IconBytes, string? IconContentType,
        string? ProgId, string? Display, string? Width, string? Height, string? Name,
        // The verbatim VML v:shape style (only when it floats — carries
        // position:absolute / margin-left / margin-top / z-index / wrap hints)
        // plus the w:object native size, so a floating OLE round-trips out of the
        // text flow instead of collapsing to inline (which pushes content down).
        string? ShapeStyle, string? DxaOrig, string? DyaOrig,
        // The VML <v:imagedata> crop rectangle (cropleft/croptop/cropright/
        // cropbottom, each a VML "Nf" 1/65536 fraction), serialized as a
        // semicolon-joined "name:value" list. Dropped on round-trip → Word
        // renders the full uncropped EMF preview, inflating the object and
        // pushing every later page down. Captured verbatim for AddOle to splice back.
        string? Crop);

    private const string RelNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    /// <summary>
    /// dump→batch: extract everything needed to faithfully re-add an OLE object
    /// run — the embedded payload bytes (raw) plus part kind / content type /
    /// target extension, the icon image bytes, and the VML display metadata
    /// (progId, drawAspect→display, width/height, friendly name). Returns null
    /// when the run is not an OLE object or its parts can't be resolved, so the
    /// emitter falls back to the warn-and-drop path.
    /// </summary>
    internal OleEmitData? GetOleEmitData(string runPath)
    {
        OpenXmlElement? element;
        try { element = NavigateToElement(ParsePath(runPath)); }
        catch { return null; }
        if (element is not Run run) return null;
        var oleObj = run.GetFirstChild<EmbeddedObject>();
        if (oleObj == null) return null;

        var oleElement = oleObj.Descendants().FirstOrDefault(e => e.LocalName == "OLEObject");
        if (oleElement == null) return null;
        string? embedRelId = null, progId = null, drawAspect = null;
        foreach (var a in oleElement.GetAttributes())
        {
            if (a.LocalName == "ProgID") progId = a.Value;
            else if (a.LocalName == "DrawAspect") drawAspect = a.Value;
            else if (a.LocalName == "id" && a.NamespaceUri == RelNs) embedRelId = a.Value;
        }
        if (string.IsNullOrEmpty(embedRelId)) return null;

        var hostPart = ResolveImageHostPart(run);
        OpenXmlPart embedPart;
        try { embedPart = hostPart.GetPartById(embedRelId); }
        catch { return null; }
        byte[] embeddedBytes;
        try
        {
            using var s = embedPart.GetStream();
            using var ms = new MemoryStream();
            s.CopyTo(ms);
            embeddedBytes = ms.ToArray();
        }
        catch { return null; }
        var oleKind = embedPart is EmbeddedPackagePart ? "package" : "object";
        var embedExt = System.IO.Path.GetExtension(embedPart.Uri.ToString()).TrimStart('.');

        byte[]? iconBytes = null;
        string? iconCt = null;
        var imageData = oleObj.Descendants().FirstOrDefault(e => e.LocalName == "imagedata");
        if (imageData != null)
        {
            var iconRelId = imageData.GetAttributes()
                .FirstOrDefault(a => a.LocalName == "id" && a.NamespaceUri == RelNs).Value;
            if (!string.IsNullOrEmpty(iconRelId))
            {
                try
                {
                    var ip = hostPart.GetPartById(iconRelId);
                    using var s = ip.GetStream();
                    using var ms = new MemoryStream();
                    s.CopyTo(ms);
                    iconBytes = ms.ToArray();
                    iconCt = ip.ContentType;
                }
                catch { /* icon is best-effort; AddOle falls back to placeholder */ }
            }
        }

        // BUG-DUMP-OLECROP: capture the VML <v:imagedata> crop attributes verbatim
        // (cropleft/croptop/cropright/cropbottom). Same imageData element as the
        // icon lookup above. Without these AddOle rebuilds an uncropped imagedata.
        string? crop = null;
        if (imageData != null)
        {
            var cropParts = new List<string>();
            foreach (var cropAttr in new[] { "cropleft", "croptop", "cropright", "cropbottom" })
            {
                var cv = imageData.GetAttributes().FirstOrDefault(a => a.LocalName == cropAttr).Value;
                if (!string.IsNullOrEmpty(cv)) cropParts.Add($"{cropAttr}:{cv}");
            }
            if (cropParts.Count > 0) crop = string.Join(";", cropParts);
        }

        string? display = string.IsNullOrEmpty(drawAspect)
            ? null
            : (drawAspect.Equals("Content", StringComparison.OrdinalIgnoreCase) ? "content" : "icon");

        string? width = null, height = null, name = null, shapeStyle = null;
        var shape = oleObj.Descendants().FirstOrDefault(e => e.LocalName == "shape");
        if (shape != null)
        {
            var alt = shape.GetAttributes().FirstOrDefault(a => a.LocalName == "alt").Value;
            if (!string.IsNullOrEmpty(alt)) name = alt;
            var style = shape.GetAttributes().FirstOrDefault(a => a.LocalName == "style").Value;
            if (!string.IsNullOrEmpty(style))
            {
                bool isFloating = false;
                foreach (var seg in style.Split(';', StringSplitOptions.RemoveEmptyEntries))
                {
                    var kv = seg.Split(':', 2);
                    if (kv.Length != 2) continue;
                    var k = kv[0].Trim().ToLowerInvariant();
                    // Keep the raw VML literal (e.g. "77pt") — AddOle's ParseEmu
                    // accepts pt/cm/in, so the frame dimensions round-trip exactly.
                    if (k == "width") width = kv[1].Trim();
                    else if (k == "height") height = kv[1].Trim();
                    else if (k == "position" && kv[1].Trim().Equals("absolute", StringComparison.OrdinalIgnoreCase))
                        isFloating = true;
                }
                // Carry the verbatim style only when the OLE floats — an inline
                // shape's style is just width/height, which AddOle rebuilds anyway,
                // and forcing a verbatim inline style would defeat the width/height
                // props. A floating shape's style holds the absolute position +
                // z-index + wrap hints that keep it out of the text flow.
                if (isFloating) shapeStyle = style;
            }
        }

        // w:object native size (w:dxaOrig/w:dyaOrig) — preserve it verbatim so a
        // floating OLE keeps Word's original object box; otherwise AddOle derives
        // it from the display size, which can rescale the rendered object.
        string? dxaOrig = oleObj.GetAttributes()
            .FirstOrDefault(a => a.LocalName == "dxaOrig").Value;
        string? dyaOrig = oleObj.GetAttributes()
            .FirstOrDefault(a => a.LocalName == "dyaOrig").Value;

        return new OleEmitData(embeddedBytes, oleKind, embedPart.ContentType, embedExt,
            iconBytes, iconCt, progId, display, width, height, name,
            shapeStyle, dxaOrig, dyaOrig, crop);
    }

    // dump→batch round-trip carrier for an ActiveX form-control run — a
    // <w:object> hosting <w:control r:id> plus a VML preview <v:imagedata r:id>
    // instead of an o:OLEObject. The dump ships the verbatim run XML and every
    // referenced package part's bytes (preview image, activeX persistence XML,
    // and that part's nested binary blob); AddActiveX rebuilds the parts and
    // rewrites the run's r:id refs, so the control needs no external files.
    internal sealed record ActiveXPartData(
        string RelId, byte[] Bytes, string ContentType, List<ActiveXPartData> Children,
        List<ActiveXExternalData> Externals);
    // Externals: relationships with TargetMode=External (hyperlinks inside a
    // VML textbox, linked content) — no part bytes, just type + target URI.
    // Appear both at run level (ActiveXEmitData.Externals) and per collected
    // part (ActiveXPartData.Externals) — e.g. a carried chart part whose
    // <c:externalData r:id> points at an external oleObject workbook.
    internal sealed record ActiveXExternalData(string RelId, string Type, string Target);
    internal sealed record ActiveXEmitData(
        string RunXml, List<ActiveXPartData> Parts, List<ActiveXExternalData> Externals);

    /// <summary>
    /// dump→batch: extract the verbatim run XML and all referenced parts for an
    /// ActiveX control run. Returns null when the run hosts no &lt;w:control&gt;
    /// or any referenced part can't be resolved, so the emitter falls back to
    /// the warn-and-drop path.
    /// </summary>
    // BUG-DUMP-FF-ROWLEVEL-BOOKMARK: every <w:bookmarkStart> name anywhere in the
    // main document body. Legacy form fields are wrapped in a same-name bookmark
    // so REF fields can target them, but Word frequently places that bookmark at
    // ROW level (a direct <w:tr> child sitting between two <w:tc>) rather than as
    // a sibling of the field's run. The table emitter cannot round-trip a
    // between-cell bookmark, so EmitParagraph's same-paragraph name check never
    // sees it and wrongly pins noBookmark — AddFormField then skips the wrapping
    // bookmark and ALL form-field bookmarks vanish, which makes Word refuse to
    // open a form-heavy document. Consulting the document-wide name set lets the
    // form-field emit recognise the row-level bookmark and have AddFormField
    // recreate it (inside the cell — functionally equivalent for Word's form
    // model). A field whose source genuinely had NO bookmark (a bare checkbox
    // grid, BUG-DUMP-FFCHECKBOX-BOOKMARK) still stays bookmark-less.
    internal HashSet<string> GetAllBookmarkNames()
    {
        var set = new HashSet<string>(StringComparer.Ordinal);
        var body = _doc.MainDocumentPart?.Document?.Body;
        if (body == null) return set;
        foreach (var bs in body.Descendants<BookmarkStart>())
            if (bs.Name?.Value is { Length: > 0 } nm) set.Add(nm);
        return set;
    }

    // BUG-DUMP-TABLE-STRUCT-BOOKMARK: a <w:bookmarkStart>/<w:bookmarkEnd> placed
    // at TABLE-STRUCTURE level — a direct child of <w:tbl> (between two <w:tr>) or
    // of <w:tr> (between two <w:tc>) — is valid OOXML and a common cross-reference
    // target (PAGEREF/REF \h), but the typed `add table` emit only walks rows and
    // cells, so these markers were dropped on round-trip → dangling references
    // ("Error! Bookmark not defined."). Return each structural marker's verbatim
    // OuterXml plus a table-relative xpath + insert action so EmitTable can
    // re-insert it at its source position via raw-set. Markers inside a cell's
    // paragraphs (the normal case) are unaffected — they ride the cell emit.
    internal List<(string Xml, string RelXpath, string Action)> GetTableStructuralBookmarks(string tablePath)
    {
        var result = new List<(string, string, string)>();
        OpenXmlElement? el;
        try { el = NavigateToElement(ParsePath(tablePath)); }
        catch { return result; }
        if (el is not Table tbl) return result;

        // tbl-level markers: direct children of <w:tbl> that are bookmark markers,
        // positioned relative to the running <w:tr> index.
        int totalRows = tbl.Elements<TableRow>().Count();
        int trIdx = 0;
        foreach (var child in tbl.ChildElements)
        {
            if (child is TableRow) { trIdx++; continue; }
            // BUG-DUMP-BLOCK-PERM: <w:permStart>/<w:permEnd> (editable-region markers
            // in a protected doc) are valid block-level children of <w:tbl>/<w:tr>
            // just like bookmarks, and were dropped the same way (enumerated only at
            // paragraph scope) — leaving an unbalanced/unbounded protected range.
            // Capture them alongside structural bookmarks.
            if (child is BookmarkStart || child is BookmarkEnd || child is PermStart || child is PermEnd)
            {
                string rel, action;
                if (trIdx == 0) { rel = "w:tr[1]"; action = "before"; }       // before first row
                else if (trIdx < totalRows) { rel = $"w:tr[{trIdx + 1}]"; action = "before"; }
                else { rel = $"w:tr[{trIdx}]"; action = "after"; }            // after last row
                result.Add((child.OuterXml, rel, action));
            }
        }

        // tr-level markers: direct children of a <w:tr> that are bookmark markers,
        // positioned relative to the running <w:tc> index within that row.
        trIdx = 0;
        foreach (var row in tbl.Elements<TableRow>())
        {
            trIdx++;
            int totalCells = row.Elements<TableCell>().Count();
            if (totalCells == 0) continue;
            int tcIdx = 0;
            foreach (var rc in row.ChildElements)
            {
                if (rc is TableCell)
                {
                    tcIdx++;
                    // BUG-DUMP-PERM-DUP: a cell-DIRECT <w:permStart>/<w:permEnd> (the
                    // displacedByCustomXml="next" markers Word places at <w:tc> level
                    // right before an inline <w:sdt>) is ALSO enumerated by
                    // GetCellStructuralBookmarks, which EmitTable invokes per cell and
                    // which emits each marker as its own faithful 1:1 raw-set op.
                    // Scanning cell-direct perms here too double-emitted every such
                    // permStart (22 vs 18 → duplicate ids → validate regression /
                    // corrupted protection ranges) while the matching permEnd was
                    // emitted once. Leave cell-direct markers to the per-cell helper;
                    // only count the cell so tr-level positioning below stays correct.
                    // The tbl-direct / tr-direct perms below are NOT inside a <w:tc>,
                    // so the per-cell helper does not see them and they stay here.
                    continue;
                }
                if (rc is BookmarkStart || rc is BookmarkEnd || rc is PermStart || rc is PermEnd)
                {
                    string rel, action;
                    if (tcIdx == 0) { rel = $"w:tr[{trIdx}]/w:tc[1]"; action = "before"; }
                    else if (tcIdx < totalCells) { rel = $"w:tr[{trIdx}]/w:tc[{tcIdx + 1}]"; action = "before"; }
                    else { rel = $"w:tr[{trIdx}]/w:tc[{tcIdx}]"; action = "after"; }
                    result.Add((rc.OuterXml, rel, action));
                }
            }
        }
        return CoalesceStructuralBookmarks(result);
    }

    // BUG-DUMP-STRUCT-BOOKMARK-ORDER: two structural markers replayed as separate
    // raw-set ops at the SAME anchor+action reverse on insert — two "after X" ops
    // put the second BEFORE the first, so a zero-length bookmark's End lands ahead
    // of its Start. The bookmark id-balancer then sees an unpaired End + an orphan
    // Start, closes the orphan with an extra zero-length End, and produces a
    // DUPLICATE bookmark id (schema-invalid, Word repair-on-open). Merge runs of
    // consecutive entries that share (RelXpath, Action) into ONE raw-set whose XML
    // concatenates them in source order, so the pair inserts atomically and stays
    // ordered Start→End.
    private static List<(string Xml, string RelXpath, string Action)> CoalesceStructuralBookmarks(
        List<(string Xml, string RelXpath, string Action)> markers)
    {
        if (markers.Count < 2) return markers;
        var merged = new List<(string Xml, string RelXpath, string Action)>();
        foreach (var m in markers)
        {
            if (merged.Count > 0)
            {
                var last = merged[^1];
                if (string.Equals(last.RelXpath, m.RelXpath, StringComparison.Ordinal)
                    && string.Equals(last.Action, m.Action, StringComparison.Ordinal))
                {
                    merged[^1] = (last.Xml + m.Xml, last.RelXpath, last.Action);
                    continue;
                }
            }
            merged.Add(m);
        }
        return merged;
    }

    // BUG-DUMP-HDRFTR-STRUCT-BOOKMARK: a <w:bookmarkStart>/<w:bookmarkEnd> that is
    // a DIRECT child of a <w:hdr>/<w:ftr> root (between block paragraphs/tables, not
    // inside one) is dropped on round-trip — EmitHeaderFooterPart walks only the
    // block children (paragraphs/tables/sdts), so a header/footer-scoped
    // cross-reference target vanishes (header/footer paragraph-level bookmarks
    // already survive via EmitParagraph). Return each root-level marker's verbatim
    // OuterXml + a part-root-relative xpath + insert action so the caller replays it
    // via raw-set. Positioned by paragraph index (the reliable anchor in a hdr/ftr;
    // structural bookmarks interleaved with tables/sdts are vanishingly rare).
    internal List<(string Xml, string RelXpath, string Action)> GetPartRootStructuralBookmarks(string partSourcePath)
    {
        var result = new List<(string, string, string)>();
        OpenXmlElement? el;
        try { el = NavigateToElement(ParsePath(partSourcePath)); }
        catch { return result; }
        if (el is not Header && el is not Footer) return result;

        int totalParas = el.Elements<Paragraph>().Count();
        int pIdx = 0;
        foreach (var child in el.ChildElements)
        {
            if (child is Paragraph) { pIdx++; continue; }
            if (child is BookmarkStart || child is BookmarkEnd)
            {
                string rel, action;
                if (totalParas == 0) { rel = "."; action = "append"; }        // no paragraph anchor
                else if (pIdx == 0) { rel = "w:p[1]"; action = "before"; }     // before first paragraph
                else if (pIdx < totalParas) { rel = $"w:p[{pIdx + 1}]"; action = "before"; }
                else { rel = $"w:p[{pIdx}]"; action = "after"; }               // after last paragraph
                result.Add((child.OuterXml, rel, action));
            }
        }
        return CoalesceStructuralBookmarks(result);
    }

    // BUG-DUMP-H97: <w:bookmarkStart>/<w:bookmarkEnd> (and <w:permStart>/<w:permEnd>)
    // that are DIRECT children of a <w:tc> — between <w:tcPr> and the cell's first
    // paragraph, or between/after cell paragraphs — were dropped on round-trip: the
    // cell content walk emits only p/tbl/sdt children, so a cell-level (often
    // column-span, colFirst/colLast) bookmark slipped through. The canonical source
    // is Google Docs, which emits cell nav anchors as leading tc-direct-child
    // bookmarks. Mirrors GetPartRootStructuralBookmarks (header/footer root): return
    // each marker's verbatim OuterXml (namespace-correct, preserving
    // colFirst/colLast) + its paragraph-relative position for the caller to raw-set
    // into the rebuilt cell.
    internal List<(string Xml, string RelXpath, string Action)> GetCellStructuralBookmarks(string cellSourcePath)
    {
        var result = new List<(string, string, string)>();
        OpenXmlElement? el;
        try { el = NavigateToElement(ParsePath(cellSourcePath)); }
        catch { return result; }
        if (el is not TableCell) return result;

        int totalParas = el.Elements<Paragraph>().Count();
        int pIdx = 0;
        foreach (var child in el.ChildElements)
        {
            if (child is Paragraph) { pIdx++; continue; }
            if (child is BookmarkStart || child is BookmarkEnd || child is PermStart || child is PermEnd)
            {
                string rel, action;
                if (totalParas == 0) { rel = "."; action = "append"; }
                else if (pIdx == 0) { rel = "w:p[1]"; action = "before"; }
                else if (pIdx < totalParas) { rel = $"w:p[{pIdx + 1}]"; action = "before"; }
                else { rel = $"w:p[{pIdx}]"; action = "after"; }
                result.Add((child.OuterXml, rel, action));
            }
        }
        // NB: do NOT coalesce. A coalesced single fragment carrying both
        // <w:bookmarkStart> and <w:bookmarkEnd> raw-set into a cell triggers a
        // duplicate bookmarkEnd (bookmark-id processing of the inserted fragment);
        // emitting each marker as its own raw-set op round-trips 1:1 (verified).
        return result;
    }

    // BUG-DUMP-BLOCK-PERM: <w:permStart>/<w:permEnd> (editable-region markers in a
    // protected doc) that are DIRECT children of <w:body> — between top-level
    // paragraphs/tables, not inside a <w:p> — were dropped on round-trip (perm
    // markers were enumerated only at paragraph scope), silently leaving an
    // unbalanced/unbounded protected range. Body-direct BOOKMARKS already survive
    // (EmitBody's bookmark case), so this captures only the perm markers, positioned
    // by top-level paragraph index, for the caller to replay via raw-set against
    // //w:body. (Table-direct perm markers ride GetTableStructuralBookmarks.)
    internal List<(string Xml, string RelXpath, string Action)> GetBodyStructuralPermMarkers()
    {
        var result = new List<(string, string, string)>();
        var body = _doc.MainDocumentPart?.Document?.Body;
        if (body == null) return result;
        int totalParas = body.Elements<Paragraph>().Count();
        int pIdx = 0;
        foreach (var child in body.ChildElements)
        {
            if (child is Paragraph) { pIdx++; continue; }
            if (child is PermStart || child is PermEnd)
            {
                string rel, action;
                if (totalParas == 0) { rel = "."; action = "append"; }
                else if (pIdx == 0) { rel = "w:p[1]"; action = "before"; }
                else if (pIdx < totalParas) { rel = $"w:p[{pIdx + 1}]"; action = "before"; }
                else { rel = $"w:p[{pIdx}]"; action = "after"; }
                result.Add((child.OuterXml, rel, action));
            }
        }
        return CoalesceStructuralBookmarks(result);
    }

    // BUG-DUMP-R72-FF-BOOKMARK-COUNT: per-name occurrence count of source body
    // bookmarks. The form-field noBookmark decision is count-aware, not boolean:
    // a doc with one <w:bookmarkStart name="Check1"> but 26 checkbox fields all
    // named "Check1" must recreate exactly ONE Check1 bookmark, not 26. The set
    // form (GetAllBookmarkNames) can't tell "1 field had it" from "all 26 did".
    internal Dictionary<string, int> GetAllBookmarkNameCounts()
    {
        var counts = new Dictionary<string, int>(StringComparer.Ordinal);
        var body = _doc.MainDocumentPart?.Document?.Body;
        if (body == null) return counts;
        foreach (var bs in body.Descendants<BookmarkStart>())
            if (bs.Name?.Value is { Length: > 0 } nm)
                counts[nm] = counts.TryGetValue(nm, out var c) ? c + 1 : 1;
        return counts;
    }

    internal ActiveXEmitData? GetActiveXEmitData(string runPath)
    {
        OpenXmlElement? element;
        try { element = NavigateToElement(ParsePath(runPath)); }
        catch { return null; }
        if (element is not Run run) return null;
        var obj = run.GetFirstChild<EmbeddedObject>();
        if (obj == null) return null;
        if (!obj.Descendants().Any(e => e.LocalName == "control")) return null;
        return CollectInlinedPartsEmitData(run, obj);
    }

    /// <summary>
    /// dump→batch: collect the OLE / preview-image parts referenced by r:id
    /// INSIDE an &lt;m:oMath&gt; (a MathType / Equation.DSMT4 object embeds an
    /// &lt;o:OLEObject r:id&gt; binary payload plus a &lt;v:imagedata r:id&gt;
    /// preview). These ride along in the verbatim math XML the equation emit
    /// ships, but their embedding relationship + binary part are not otherwise
    /// recreated — leaving a dangling r:id (a silent embedding loss and a
    /// validator NullReferenceException) on replay. Returns the carrier parts so
    /// the equation op can base64-inline them (part{N}.*) and AddEquation can
    /// rematerialize + rewrite the ids; null when the math references no parts.
    /// </summary>
    internal ActiveXEmitData? GetMathInlinedPartsEmitData(string mathPath)
    {
        OpenXmlElement? element = NavigateMathElement(mathPath);
        if (element == null) return null;
        // The math element (m:oMath / m:oMathPara) is both the carrier body and
        // the subtree to scan for r:id references. scanRawXml: the OLE object's
        // relationship ids live in schema-invalid <w:object>-in-<m:r> nesting the
        // typed walk can't see, so recover them from the raw OuterXml.
        return CollectInlinedPartsEmitData(
            ResolveImageHostPart(element), element, element, scanRawXml: true);
    }

    // The emit-side equation node Path uses the `/equation[N]` segment (mirroring
    // the `add equation` command vocabulary), but the path navigator addresses
    // the underlying element as `/oMath[N]` (or `/oMathPara[N]` for display). Try
    // the path verbatim, then with that final-segment normalization, so the
    // equation emit can resolve its own math element regardless of which form it
    // carries.
    private OpenXmlElement? NavigateMathElement(string mathPath)
    {
        foreach (var candidate in new[]
                 {
                     mathPath,
                     mathPath.Replace("/equation[", "/oMath[", StringComparison.Ordinal),
                     mathPath.Replace("/equation[", "/oMathPara[", StringComparison.Ordinal),
                 })
        {
            try
            {
                var el = NavigateToElement(ParsePath(candidate));
                if (el != null) return el;
            }
            catch { /* try the next normalization */ }
        }
        return null;
    }

    /// <summary>
    /// dump→batch: extract the verbatim run XML and all referenced parts for a
    /// SmartArt diagram run. The &lt;dgm:relIds&gt; element references the
    /// data / layout / quickStyle / colors parts via r:dm / r:lo / r:qs / r:cs,
    /// and the data part nests the rendered-drawing part. Same carrier shape as
    /// the ActiveX path; returns null when the run hosts no diagram or any
    /// referenced part can't be resolved.
    /// </summary>
    internal ActiveXEmitData? GetDiagramEmitData(string runPath)
    {
        OpenXmlElement? element;
        try { element = NavigateToElement(ParsePath(runPath)); }
        catch { return null; }
        if (element is not Run run) return null;
        var drawing = run.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Drawing>();
        if (drawing == null) return null;
        if (!drawing.Descendants().Any(e => e.LocalName == "relIds")) return null;
        return CollectInlinedPartsEmitData(run, drawing);
    }

    /// <summary>
    /// dump→batch: extract the verbatim run XML and all referenced parts /
    /// external relationships for a legacy VML shape run (&lt;w:pict&gt;).
    /// Covers textbox content carrying hyperlinks (external rels) and
    /// v:imagedata image parts. Returns null when the run hosts no pict or a
    /// reference can't be resolved.
    /// </summary>
    internal ActiveXEmitData? GetVmlShapeEmitData(string runPath)
    {
        OpenXmlElement? element;
        try { element = NavigateToElement(ParsePath(runPath)); }
        catch { return null; }
        if (element is not Run run) return null;
        // Word and other editors export VML textboxes wrapped in mc:AlternateContent
        // (Choice = modern wps drawing, Fallback = w:pict) — probe the whole
        // run subtree, and collect rel refs from the whole run so the Choice
        // branch's references resolve on replay too.
        if (!run.Descendants<DocumentFormat.OpenXml.Wordprocessing.Picture>().Any())
            return null;
        return CollectInlinedPartsEmitData(run, run);
    }

    /// <summary>
    /// dump→batch: extract the verbatim &lt;w:sdt&gt; XML and every referenced
    /// part / external relationship for a rich BLOCK content control (a cover
    /// page with anchored textboxes and a logo image). Same carrier shape as
    /// the ActiveX / diagram / vmlshape paths; returns null when the SDT
    /// references nothing (the plain raw-set passthrough already handles that)
    /// or a reference can't be resolved.
    /// </summary>
    /// <summary>
    /// dump→batch: extract the verbatim run XML and referenced parts for a
    /// modern DrawingML shape run (&lt;w:drawing&gt; hosting wps:wsp with an
    /// image blipFill). Same carrier shape as the vmlshape path.
    /// </summary>
    internal ActiveXEmitData? GetDrawingShapeEmitData(string runPath)
    {
        OpenXmlElement? element;
        try { element = NavigateToElement(ParsePath(runPath)); }
        catch { return null; }
        if (element is not Run run) return null;
        if (!run.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>().Any())
            return null;
        return GuardCarrierContentTypes(CollectInlinedPartsEmitData(run, run));
    }

    // The inlined-parts factories (CreateInlinedPart) cover activeX, diagram
    // and image parts. A subtree referencing anything else — a chart, a
    // header/footer (an SDT wrapping a sectPr carries headerReference rel
    // ids) — must NOT take the carrier: the apply side would reject the
    // content type and fail the step. Returning null lets the caller fall
    // back to its legacy path (typed chart emit, warn-and-flatten, …).
    private static ActiveXEmitData? GuardCarrierContentTypes(ActiveXEmitData? data)
    {
        if (data == null) return null;
        static bool Supported(string ct) =>
            ct is "application/vnd.ms-office.activeX+xml"
               or "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml"
               or "application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml"
               or "application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml"
               or "application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml"
               or "application/vnd.ms-office.drawingml.diagramDrawing+xml"
            || ct.StartsWith("image/", StringComparison.OrdinalIgnoreCase);
        foreach (var part in data.Parts)
            if (!Supported(part.ContentType)) return null;
        return data;
    }

    internal ActiveXEmitData? GetSdtEmitData(string sdtPath)
    {
        OpenXmlElement? element;
        try { element = NavigateToElement(ParsePath(sdtPath)); }
        catch { return null; }
        // Block AND inline (run-level) content controls — an inline SDT
        // wrapping a picture (w:sdt inside w:p) needs the same carrier.
        if (element is SdtBlock sdtBlock)
            return GuardCarrierContentTypes(CollectInlinedPartsEmitData(sdtBlock, sdtBlock));
        if (element is SdtRun sdtRun)
            return GuardCarrierContentTypes(CollectInlinedPartsEmitData(sdtRun, sdtRun));
        return null;
    }

    /// <summary>
    /// dump→batch: capture a chart's <c>c:userShapes</c> overlay drawing (the
    /// chartshapes part — a logo/annotation/picture drawn on top of the chart in
    /// Word's chart editor) plus the images it embeds, so AddChart can re-attach
    /// it. Returns null when the chart at <paramref name="runPath"/> has no
    /// userShapes part. The shipped XML is the part's raw stream (self-contained
    /// namespaces), not OuterXml.
    /// </summary>
    internal ActiveXEmitData? GetChartUserShapesEmitData(string runPath)
    {
        var chartPart = ResolveChartPartFromRunPath(runPath);
        if (chartPart == null) return null;
        var cdp = chartPart.GetPartsOfType<ChartDrawingPart>().FirstOrDefault();
        if (cdp == null) return null;

        string xml;
        try
        {
            using var s = cdp.GetStream(FileMode.Open, FileAccess.Read);
            using var r = new StreamReader(s);
            xml = r.ReadToEnd();
        }
        catch { return null; }
        if (string.IsNullOrWhiteSpace(xml)) return null;

        var parts = new List<ActiveXPartData>();
        foreach (var rel in cdp.Parts)
        {
            byte[] bytes;
            try { bytes = ReadPartBytes(rel.OpenXmlPart); }
            catch { return null; }
            parts.Add(new ActiveXPartData(
                rel.RelationshipId, bytes, rel.OpenXmlPart.ContentType,
                new List<ActiveXPartData>(), new List<ActiveXExternalData>()));
        }
        var externals = new List<ActiveXExternalData>();
        foreach (var ext in cdp.ExternalRelationships)
            externals.Add(new ActiveXExternalData(ext.Id, ext.RelationshipType, ext.Uri.OriginalString));

        if (parts.Count == 0 && externals.Count == 0) return null;
        // Only image overlays round-trip (CreateInlinedPart covers image/*).
        return GuardCarrierContentTypes(new ActiveXEmitData(xml, parts, externals));
    }

    // BUG-DUMP-CHART-SIDECARS: the typed `add chart` path rebuilds a native
    // DrawingML chart from semantic props and does NOT carry the source chart's
    // sidecar parts — its chartStyle, chartColorStyle, themeOverride, and the
    // embedded data workbook (referenced by <c:externalData>). Gather them so
    // AddChart can re-attach them. Returns (role, contentType, bytes) where role
    // is one of style/colors/themeOverride/package. The ChartDrawingPart
    // (userShapes) and image parts are excluded — those have their own carriers.
    internal List<(string Role, string ContentType, byte[] Bytes)>? GetChartSidecarEmitData(string runPath)
    {
        var chartPart = ResolveChartPartFromRunPath(runPath);
        if (chartPart == null) return null;
        var result = new List<(string, string, byte[])>();
        foreach (var rel in chartPart.Parts)
        {
            var part = rel.OpenXmlPart;
            string? role = part switch
            {
                ChartStylePart => "style",
                ChartColorStylePart => "colors",
                ThemeOverridePart => "themeOverride",
                EmbeddedPackagePart => "package",
                EmbeddedObjectPart => "package",
                _ => null,
            };
            if (role == null) continue;
            byte[] bytes;
            try { bytes = ReadPartBytes(part); }
            catch { continue; }
            result.Add((role, part.ContentType, bytes));
        }
        return result.Count > 0 ? result : null;
    }

    private ChartPart? ResolveChartPartFromRunPath(string runPath)
    {
        OpenXmlElement? element;
        try { element = NavigateToElement(ParsePath(runPath)); }
        catch { return null; }
        if (element is not Run run) return null;
        var drawing = run.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Drawing>();
        if (drawing == null) return null;
        var chartRef = drawing
            .Descendants<DocumentFormat.OpenXml.Drawing.Charts.ChartReference>().FirstOrDefault();
        if (chartRef?.Id?.Value == null) return null;
        try { return ResolveImageHostPart(run).GetPartById(chartRef.Id.Value) as ChartPart; }
        catch { return null; }
    }

    /// <summary>
    /// dump→batch: capture a native DrawingML chart run VERBATIM — the run's
    /// <c>&lt;w:drawing&gt;</c> (referencing the chart part via
    /// <c>&lt;c:chart r:id&gt;</c>) plus the chart part's bytes and every sidecar
    /// it owns (chartStyle, chartColorStyle, themeOverride, the userShapes
    /// overlay drawing, an embedded data workbook, and any external-workbook
    /// relationship). This SUPERSEDES the typed <c>add chart</c> rebuild, whose
    /// BuildChartProps round-trip de-references the data (numRef→numLit, drops
    /// strRef category labels, ptCount data points, dLbls, externalData) and
    /// visibly compresses the rendered chart. Mirrors the diagram/activeX
    /// inlined-parts carrier. Returns null when the run hosts no chart reference
    /// or a referenced part can't be resolved, so the caller falls back to the
    /// typed path.
    /// </summary>
    internal ActiveXEmitData? GetChartVerbatimEmitData(string runPath)
    {
        OpenXmlElement? element;
        try { element = NavigateToElement(ParsePath(runPath)); }
        catch { return null; }
        if (element is not Run run) return null;
        var drawing = run.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Drawing>();
        if (drawing == null) return null;
        if (!drawing.Descendants<DocumentFormat.OpenXml.Drawing.Charts.ChartReference>().Any())
            return null;
        // GATE (editability vs fidelity): only supersede the typed `add chart`
        // path for charts the typed BuildChartProps round-trip actually mangles —
        // those whose data is a LIVE cached reference (<c:numRef>/<c:strRef>) or
        // links an external workbook (<c:externalData>), or that carry a
        // userShapes overlay / mc:AlternateContent the typed path drops. Charts
        // AUTHORED via `add chart` embed LITERAL data (<c:numLit>/<c:strLit>) and
        // carry none of these markers — keep them on the typed path so the dump
        // stays human/agent-editable (`add chart --prop data=…`) instead of an
        // opaque verbatim blob. This is the deliberate dump-as-code boundary.
        var chartPart = ResolveChartPartFromRunPath(runPath);
        if (chartPart == null || !ChartNeedsVerbatim(chartPart)) return null;
        // No GuardCarrierContentTypes here: chart + its sidecar content types are
        // explicitly supported by CreateInlinedPart / CreateInlinedChildPart, so
        // the guard's conservative allowlist (which excludes charts on purpose)
        // must not veto this carrier.
        return CollectInlinedPartsEmitData(run, drawing);
    }

    // A chart whose typed-prop round-trip is known-lossy: live data references
    // (numRef/strRef de-reference to numLit/strLit, dropping ptCount points and
    // category labels), an external-workbook link, a userShapes overlay, or an
    // mc:AlternateContent block the typed rebuild can't represent.
    private static bool ChartNeedsVerbatim(ChartPart chartPart)
    {
        string xml;
        try
        {
            using var s = chartPart.GetStream(FileMode.Open, FileAccess.Read);
            using var r = new StreamReader(s);
            xml = r.ReadToEnd();
        }
        catch { return false; }
        return xml.Contains("<c:numRef", StringComparison.Ordinal)
            || xml.Contains("<c:strRef", StringComparison.Ordinal)
            || xml.Contains("<c:externalData", StringComparison.Ordinal)
            || xml.Contains("<c:userShapes", StringComparison.Ordinal)
            || xml.Contains("AlternateContent", StringComparison.Ordinal);
    }

    // Shared collector for the `add activex` / `add diagram` carriers: every
    // relationship-namespace attribute in the subtree (r:id, r:dm, r:lo, …)
    // names a part on the run's host part; ship each part's bytes plus its
    // direct children (activeX binary blob, diagram rendered drawing).
    private ActiveXEmitData? CollectInlinedPartsEmitData(OpenXmlElement run, OpenXmlElement subtree)
        => CollectInlinedPartsEmitData(ResolveImageHostPart(run), run, subtree);

    // Host-explicit overload: a chart's userShapes drawing lives on the
    // ChartDrawingPart (not a header/footer/main part reachable by walking the
    // run's ancestors), so the caller supplies the host whose relationships the
    // r:embed/r:id references resolve against. <paramref name="runXmlElement"/>
    // is the element whose OuterXml is shipped verbatim (the carrier body).
    private ActiveXEmitData? CollectInlinedPartsEmitData(
        OpenXmlPart hostPart, OpenXmlElement runXmlElement, OpenXmlElement subtree,
        bool scanRawXml = false)
    {
        var relIds = new List<string>();
        // BUG-DUMP-OLE-IN-OMATH: schema-invalid nesting (a <w:object> hosting an
        // <o:OLEObject r:id> inside an <m:r> math run) is loaded by the SDK as raw
        // content the typed object model does not expose as navigable children, so
        // subtree.Descendants() finds none of its relationship ids. Scan the raw
        // OuterXml via XLinq (namespace-aware, prefix-independent) to recover them.
        // Opt-in (math only) so the activex/diagram callers keep their exact
        // behaviour. Deduped against the typed walk below.
        if (scanRawXml)
        {
            try
            {
                var raw = System.Xml.Linq.XElement.Parse(subtree.OuterXml);
                foreach (var el in raw.DescendantsAndSelf())
                    foreach (var a in el.Attributes())
                        if (a.Name.NamespaceName == RelNs
                            && !string.IsNullOrEmpty(a.Value) && !relIds.Contains(a.Value))
                            relIds.Add(a.Value);
            }
            catch { /* malformed fragment — fall back to the typed walk */ }
        }
        foreach (var el in subtree.Descendants())
        {
            foreach (var a in el.GetAttributes())
            {
                if (a.NamespaceUri == RelNs
                    && !string.IsNullOrEmpty(a.Value) && !relIds.Contains(a.Value!))
                    relIds.Add(a.Value!);
            }
        }
        if (relIds.Count == 0) return null;

        var parts = new List<ActiveXPartData>();
        var externals = new List<ActiveXExternalData>();
        for (int idx = 0; idx < relIds.Count; idx++)
        {
            var relId = relIds[idx];
            OpenXmlPart part;
            byte[] bytes;
            try
            {
                part = hostPart.GetPartById(relId);
                bytes = ReadPartBytes(part);
            }
            catch
            {
                // Not a part — maybe an EXTERNAL relationship (a hyperlink
                // inside a VML textbox, linked content). Those carry no bytes;
                // ship type + target so the apply side recreates the rel.
                var hyper = hostPart.HyperlinkRelationships.FirstOrDefault(h => h.Id == relId);
                if (hyper != null)
                {
                    externals.Add(new ActiveXExternalData(
                        relId, hyper.RelationshipType, hyper.Uri.OriginalString));
                    continue;
                }
                var ext = hostPart.ExternalRelationships.FirstOrDefault(x => x.Id == relId);
                if (ext != null)
                {
                    externals.Add(new ActiveXExternalData(
                        relId, ext.RelationshipType, ext.Uri.OriginalString));
                    continue;
                }
                return null;
            }
            var children = new List<ActiveXPartData>();
            foreach (var child in part.Parts)
            {
                byte[] cb;
                try { cb = ReadPartBytes(child.OpenXmlPart); }
                catch { return null; }
                // BUG-DUMP-R71-USERSHAPES-IMG: a child part can own its OWN parts
                // — a chart's <c:userShapes> ChartDrawingPart references an image
                // (drawingN.xml r:embed -> media/imageN). Without capturing that
                // grandchild the rebuilt drawing's r:embed dangles ("relationship
                // does not exist"). Collect one more nesting level so the image
                // (and any other child-of-child) round-trips. Grandchildren are
                // recreated under the child with their ORIGINAL rel id, so the
                // child's verbatim XML refs resolve without rewriting.
                var grandchildren = new List<ActiveXPartData>();
                foreach (var gc in child.OpenXmlPart.Parts)
                {
                    byte[] gcb;
                    try { gcb = ReadPartBytes(gc.OpenXmlPart); }
                    catch { return null; }
                    grandchildren.Add(new ActiveXPartData(
                        gc.RelationshipId, gcb, gc.OpenXmlPart.ContentType,
                        new List<ActiveXPartData>(), new List<ActiveXExternalData>()));
                }
                children.Add(new ActiveXPartData(
                    child.RelationshipId, cb, child.OpenXmlPart.ContentType,
                    grandchildren, new List<ActiveXExternalData>()));
            }
            // A collected part can carry its OWN external relationships — e.g. a
            // chart part whose <c:externalData r:id> links an external oleObject
            // workbook (TargetMode=External). Capture them so the apply side
            // recreates the rel on the part; otherwise the verbatim part bytes
            // reference a now-dangling id and the host app refuses to render.
            var partExternals = new List<ActiveXExternalData>();
            foreach (var hyper in part.HyperlinkRelationships)
                partExternals.Add(new ActiveXExternalData(
                    hyper.Id, hyper.RelationshipType, hyper.Uri.OriginalString));
            foreach (var ext in part.ExternalRelationships)
                partExternals.Add(new ActiveXExternalData(
                    ext.Id, ext.RelationshipType, ext.Uri.OriginalString));
            parts.Add(new ActiveXPartData(relId, bytes, part.ContentType, children, partExternals));

            // A collected XML part's CONTENT can reference further host-part
            // relationships via bare relId="…" attributes — SmartArt's data
            // part points at its rendered-drawing part through
            // <dsp:dataModelExt relId="rId12"> resolved against the MAIN part,
            // not the data part. Chase those so the rendered drawing ships too.
            if (part.ContentType.EndsWith("+xml", StringComparison.OrdinalIgnoreCase))
            {
                var xmlText = System.Text.Encoding.UTF8.GetString(bytes);
                foreach (System.Text.RegularExpressions.Match m in
                         System.Text.RegularExpressions.Regex.Matches(xmlText, "relId=\"([^\"]+)\""))
                {
                    var extra = m.Groups[1].Value;
                    if (!string.IsNullOrEmpty(extra) && !relIds.Contains(extra))
                    {
                        try { hostPart.GetPartById(extra); }
                        catch { continue; }
                        relIds.Add(extra);
                    }
                }
            }
        }
        // A run that references ONLY external rels (a VML textbox whose sole
        // refs are hyperlinks) is still a valid carrier.
        if (parts.Count == 0 && externals.Count == 0) return null;
        return new ActiveXEmitData(runXmlElement.OuterXml, parts, externals);
    }

    private static byte[] ReadPartBytes(OpenXmlPart part)
    {
        using var s = part.GetStream();
        using var ms = new MemoryStream();
        s.CopyTo(ms);
        return ms.ToArray();
    }

    private OpenXmlPart ResolveImageHostPart(OpenXmlElement run)
    {
        // BUG-R14A: self-or-ancestor — `run` may itself BE the part-root (the SDT
        // carrier in AddSdt passes the Header/Footer element directly as parent
        // when an `add sdt parent=/footer[N]` lands a rich content control at the
        // footer root). Ancestors<Footer>() excludes self, so a bare
        // Ancestors lookup fell through to MainDocumentPart and registered the
        // SDT's hyperlink/image relationship on document.xml.rels instead of the
        // footer's — leaving the r:id in word/footerN.xml dangling and the file
        // unopenable in Word. Mirrors ResolveHostPart's self-or-ancestor walk.
        var headerAncestor = run as Header ?? run.Ancestors<Header>().FirstOrDefault();
        if (headerAncestor != null)
        {
            var hp = _doc.MainDocumentPart!.HeaderParts
                .FirstOrDefault(p => ReferenceEquals(p.Header, headerAncestor));
            if (hp != null) return hp;
        }
        var footerAncestor = run as Footer ?? run.Ancestors<Footer>().FirstOrDefault();
        if (footerAncestor != null)
        {
            var fp = _doc.MainDocumentPart!.FooterParts
                .FirstOrDefault(p => ReferenceEquals(p.Footer, footerAncestor));
            if (fp != null) return fp;
        }
        return _doc.MainDocumentPart!;
    }

    // BUG-DUMP-R45-1: enumerate the embedded font binaries (word/fonts/*.odttf)
    // attached to the FontTablePart, keyed by the relationship id the
    // fontTable.xml <w:embed*> elements reference. The dump emits each as a
    // companion `raw-set /fontTable --action embed-binary` carrying the bytes as
    // a data: URI; the apply side rebuilds the FontPart with the SAME r:id so the
    // round-tripped embed refs (kept verbatim in the fontTable XML) resolve.
    // The .odttf bytes are obfuscated with the w:fontKey already in the XML —
    // they round-trip raw, no de/re-obfuscation needed.
    // customXml data stores (item.xml + itemProps.xml pairs). SDT content
    // controls bind to a store through the itemProps datastore-item GUID, not
    // a relationship id, so recreating the parts with their bytes verbatim is
    // sufficient for the bindings to resolve. The dump ships each pair as
    // `raw-set /customXml --action embed-binary` (item) plus
    // `raw-set /customXml/<itemRelId> --action embed-binary` (its props child).
    internal List<(string RelId, byte[] Bytes, string ContentType,
                   (string RelId, byte[] Bytes, string ContentType)? Props)> GetCustomXmlEmitData()
    {
        var result = new List<(string, byte[], string, (string, byte[], string)?)>();
        var mainPart = _doc.MainDocumentPart;
        if (mainPart == null) return result;
        foreach (var part in mainPart.CustomXmlParts)
        {
            byte[] bytes;
            try { bytes = ReadPartBytes(part); }
            catch { continue; }
            (string, byte[], string)? props = null;
            var pp = part.CustomXmlPropertiesPart;
            if (pp != null)
            {
                try { props = (part.GetIdOfPart(pp), ReadPartBytes(pp), pp.ContentType); }
                catch { /* props are optional; item still round-trips */ }
            }
            result.Add((mainPart.GetIdOfPart(part), bytes, part.ContentType, props));
        }
        return result;
    }

    internal List<(string RelId, byte[] Bytes, string ContentType)> GetEmbeddedFontParts()
    {
        var result = new List<(string, byte[], string)>();
        var ftp = _doc.MainDocumentPart?.FontTablePart;
        if (ftp == null) return result;
        foreach (var fp in ftp.FontParts)
        {
            string relId;
            try { relId = ftp.GetIdOfPart(fp); }
            catch { continue; }
            if (string.IsNullOrEmpty(relId)) continue;
            try
            {
                using var s = fp.GetStream();
                using var ms = new MemoryStream();
                s.CopyTo(ms);
                result.Add((relId, ms.ToArray(), fp.ContentType));
            }
            catch { /* unreadable font part — skip; emitter falls back to strip */ }
        }
        return result;
    }

    // BUG-DUMP-R45-2: enumerate the image binaries (word/media/*) attached to
    // the NumberingDefinitionsPart, keyed by the relationship id the
    // numbering.xml <w:numPicBullet> VML <v:imagedata r:id> references. Mirrors
    // GetEmbeddedFontParts — the dump emits each as a companion
    // `raw-set /numbering --action embed-binary` so the apply side rebuilds the
    // ImagePart with the SAME r:id and the kept numPicBullet VML resolves.
    internal List<(string RelId, byte[] Bytes, string ContentType)> GetNumberingImageParts()
    {
        var result = new List<(string, byte[], string)>();
        var ndp = _doc.MainDocumentPart?.NumberingDefinitionsPart;
        if (ndp == null) return result;
        foreach (var ip in ndp.ImageParts)
        {
            string relId;
            try { relId = ndp.GetIdOfPart(ip); }
            catch { continue; }
            if (string.IsNullOrEmpty(relId)) continue;
            try
            {
                using var s = ip.GetStream();
                using var ms = new MemoryStream();
                s.CopyTo(ms);
                result.Add((relId, ms.ToArray(), ip.ContentType));
            }
            catch { /* unreadable image part — skip */ }
        }
        return result;
    }

    // ==================== Raw Layer ====================

    public string Raw(string partPath, int? startRow = null, int? endRow = null, HashSet<string>? cols = null)
    {
        if (partPath == null) throw new ArgumentNullException(nameof(partPath));
        var mainPart = _doc.MainDocumentPart;
        if (mainPart == null) return "(no main part)";

        // CONSISTENCY(zip-uri-lookup): see RawXmlHelper. Any path ending in
        // .xml or .rels is resolved against the package directly.
        if (RawXmlHelper.IsZipUriPath(partPath))
        {
            var xml = RawXmlHelper.TryReadByZipUri(_doc, _filePath, partPath)
                ?? throw new ArgumentException(
                    $"Unknown part: {partPath}. The path was treated as a zip-internal URI " +
                    $"but no matching part exists in the package. " +
                    $"Use semantic paths (/document, /styles, /header[N]) for stable identification.");
            return xml;
        }

        return partPath.ToLowerInvariant() switch
        {
            "/document" => mainPart.Document?.OuterXml ?? "",
            "/styles" => mainPart.StyleDefinitionsPart?.Styles?.OuterXml ?? "(no styles)",
            "/settings" => mainPart.DocumentSettingsPart?.Settings?.OuterXml ?? "(no settings)",
            "/numbering" => mainPart.NumberingDefinitionsPart?.Numbering?.OuterXml ?? "(no numbering)",
            "/comments" => mainPart.WordprocessingCommentsPart?.Comments?.OuterXml ?? "(no comments)",
            "/footnotes" => mainPart.FootnotesPart?.Footnotes?.OuterXml ?? "(no footnotes)",
            "/endnotes" => mainPart.EndnotesPart?.Endnotes?.OuterXml ?? "(no endnotes)",
            "/theme" => mainPart.ThemePart?.Theme?.OuterXml ?? "(no theme)",
            // BUG-DUMP-R42-3: word/fontTable.xml declares font faces +
            // altName substitutions (e.g. 方正小标宋简体 altName=方正舒体).
            // Without it Word substitutes different faces — a real visual
            // change. FontTablePart has a typed Fonts root we read verbatim.
            "/fonttable" => mainPart.FontTablePart?.Fonts?.OuterXml ?? "(no fontTable)",
            "/websettings" => mainPart.WebSettingsPart?.WebSettings?.OuterXml ?? "(no webSettings)",
            _ when partPath.StartsWith("/header") => GetHeaderRawXml(partPath),
            _ when partPath.StartsWith("/footer") => GetFooterRawXml(partPath),
            _ when partPath.StartsWith("/chart") => GetChartRawXml(partPath),
            _ => throw new ArgumentException($"Unknown part: {partPath}. Available: /document, /styles, /settings, /numbering, /comments, /footnotes, /endnotes, /theme, /header[n], /footer[n], /chart[n]")
        };
    }

    public void RawSet(string partPath, string xpath, string action, string? xml)
    {
        Modified = true;
        ClearBodyChildIndex(); // raw-set may rewrite the body / its paragraph set
        if (partPath == null) throw new ArgumentNullException(nameof(partPath));
        var mainPart = _doc.MainDocumentPart
            ?? throw new InvalidOperationException("No main document part");

        // BUG-DUMP-R45-1 / BUG-DUMP-R45-2: companion binary-attach action for the
        // /fontTable and /numbering raw-set replace. The dump emits the part XML
        // verbatim (keeping <w:embed*> / <w:numPicBullet> refs) PLUS one
        // `embed-binary` op per referenced binary, carrying the bytes as a
        // `data:<ct>;base64,…` URI in `xml` and the SOURCE relationship id in
        // `xpath`. Rebuild the FontPart / ImagePart with that exact r:id so the
        // already-replaced XML's refs resolve — no XML rewrite, no dangling rel.
        // Mirrors the OLE base64-inline round-trip (OleHelper.AddEmbeddedPartFromBytes).
        if (string.Equals(action, "embed-binary", StringComparison.OrdinalIgnoreCase))
        {
            RawEmbedBinary(mainPart, partPath, xpath, xml);
            return;
        }

        // docProps whole-part replace (core/app/custom.xml). EmitDocPropsRaw
        // emits a normal `raw-set replace` whose xpath is the part root, but the
        // Open-XML SDK does not persist a docProps mutation made mid-session on a
        // stream-opened package (both typed-root and stream writes are reverted
        // by _doc.Save()). So stage the whole-part XML and rewrite the zip entry
        // after the package closes (FlushPendingWholeParts) instead of taking
        // the normal xpath path. Recognised by part path, not action, so it
        // stays a plain `replace` in the dump.
        if (IsDocPropsWholePart(partPath)
            && string.Equals(action, "replace", StringComparison.OrdinalIgnoreCase))
        {
            StashWholePartReplace(partPath, xml);
            return;
        }

        if (RawXmlHelper.IsZipUriPath(partPath))
        {
            var part = RawXmlHelper.FindPartByZipUri(_doc, partPath)
                ?? throw new ArgumentException(
                    $"Unknown part: {partPath}. The path was treated as a zip-internal URI " +
                    $"but no matching part exists in the package. " +
                    $"Use semantic paths (/document, /styles, /header[N]) for stable identification.");
            RawXmlHelper.Execute(part, xpath, action, xml);
            return;
        }

        OpenXmlPartRootElement rootElement;
        var lowerPath = partPath.ToLowerInvariant();

        // Fast-path: a raw-set that replaces ONE <w:style> by styleId on /styles.
        // The dump emits one such op per verbatim-fidelity style (1223 of them in
        // a 4070-style WPS doc). The generic path below re-serializes the whole
        // styles part (rootElement.OuterXml + XDocument.Parse + InnerXml=) per op
        // — O(part) each, which dominated a multi-minute batch. Swap the single
        // element on the live SDK DOM instead (O(fragment)). Any deviation —
        // different xpath shape, styleId absent, fragment the SDK won't parse —
        // returns false and falls through to the proven generic path below, so
        // this changes performance, not behavior.
        if (lowerPath == "/styles"
            && string.Equals(action, "replace", StringComparison.OrdinalIgnoreCase)
            && xml != null
            && TryReplaceSingleStyleFast(mainPart, xpath, xml))
        {
            return;
        }

        // Fast-path: a raw-set that inserts ONE <w:p> immediately before the
        // body's trailing sectPr. The dump emits one such op per cross-paragraph
        // field-span member (EmitCrossParagraphFieldMember) — e.g. 785 of them for
        // a large back-of-book INDEX field, perfectly consecutive at end of batch.
        // The generic path re-serializes the whole document part per op
        // (rootElement.OuterXml + XDocument.Parse + InnerXml=) — O(part) each,
        // turning a 7MB / 11k-paragraph batch into O(n²) (minutes). Insert on the
        // live SDK DOM instead (O(fragment), O(1) via the append-monotonic cache).
        // Same opt-not-behavior contract as the /styles fast-path: any deviation
        // (xpath shape, missing body/sectPr, unparseable fragment) falls through.
        if (lowerPath is "/document" or "/"
            && (string.Equals(action, "before", StringComparison.OrdinalIgnoreCase)
                || string.Equals(action, "insertbefore", StringComparison.OrdinalIgnoreCase))
            && xml != null
            && TryInsertBodyParaBeforeSectPrFast(mainPart, xpath, xml))
        {
            return;
        }

        // Fast-path: a raw-set rooted at a single body table by its global
        // document-order ordinal — (//w:tbl)[N] or (//w:tbl)[N]/<tail>. The dump
        // emits one such op per verbatim cell-content fallback (rich field result,
        // nested SDT, VML textbox) plus tblGrid / cellMerge round-trips — e.g. 2580
        // of them for a 757-table report. The generic path re-serializes the whole
        // multi-MB document part per op (rootElement.OuterXml + XDocument.Parse +
        // InnerXml=) — O(part) each, turning a table-heavy batch into O(ops × part)
        // (minutes). Resolve the Nth table on the live SDK DOM and run the XPath
        // engine against that table's subtree only (O(table), a few KB). Same
        // opt-not-behavior contract as the /styles and sectPr fast-paths above: any
        // deviation (xpath shape, ordinal out of range, 0 matches, unparseable
        // fragment) returns false and falls through to the proven generic path.
        if (lowerPath is "/document" or "/"
            && TryRawSetWithinTableFast(mainPart, xpath, action, xml))
        {
            return;
        }

        if (lowerPath is "/document" or "/")
            rootElement = mainPart.Document ?? throw new InvalidOperationException("No document");
        else if (lowerPath is "/styles")
            rootElement = mainPart.StyleDefinitionsPart?.Styles ?? throw new InvalidOperationException("No styles part");
        else if (lowerPath is "/settings")
            rootElement = mainPart.DocumentSettingsPart?.Settings ?? throw new InvalidOperationException("No settings part");
        else if (lowerPath is "/numbering")
        {
            // CONSISTENCY(raw-set-create-missing-part): see /theme branch.
            var numPart = mainPart.NumberingDefinitionsPart ?? mainPart.AddNewPart<NumberingDefinitionsPart>();
            if (numPart.Numbering == null)
            {
                numPart.Numbering = new Numbering();
                numPart.Numbering.Save();
            }
            rootElement = numPart.Numbering;
        }
        else if (lowerPath is "/websettings")
        {
            // CONSISTENCY(raw-set-create-missing-part): blank docs have no
            // webSettings part; dump→batch from a Word-authored source emits
            // raw-set /webSettings replace, so create the part lazily.
            var wsPart = mainPart.WebSettingsPart ?? mainPart.AddNewPart<WebSettingsPart>();
            if (wsPart.WebSettings == null)
            {
                wsPart.WebSettings = new WebSettings();
                wsPart.WebSettings.Save();
            }
            rootElement = wsPart.WebSettings;
        }
        else if (lowerPath is "/comments")
            rootElement = mainPart.WordprocessingCommentsPart?.Comments ?? throw new InvalidOperationException("No comments part");
        else if (lowerPath is "/footnotes")
        {
            // CONSISTENCY(raw-set-create-missing-part): blank docs have no
            // FootnotesPart; a dump→batch round-trip emits raw-set /footnotes
            // replace ONLY when the source carries a CUSTOMIZED separator note
            // (PAGE field / "- N -" text) and has no body notes — the gap the
            // typed `add footnote` path does not cover (it recreates the part
            // with DEFAULT separators only). Lazily create the part + an empty
            // <w:footnotes> root so RawXmlHelper.Execute can match /w:footnotes
            // and replace it with the dumped XML, which re-introduces the custom
            // separator the kept settings footnotePr -1/0 refs resolve against.
            var fnPart = mainPart.FootnotesPart ?? mainPart.AddNewPart<FootnotesPart>();
            if (fnPart.Footnotes == null)
            {
                fnPart.Footnotes = new Footnotes();
                fnPart.Footnotes.Save();
            }
            rootElement = fnPart.Footnotes;
        }
        else if (lowerPath is "/endnotes")
        {
            // CONSISTENCY(raw-set-create-missing-part): mirrors /footnotes.
            var enPart = mainPart.EndnotesPart ?? mainPart.AddNewPart<EndnotesPart>();
            if (enPart.Endnotes == null)
            {
                enPart.Endnotes = new Endnotes();
                enPart.Endnotes.Save();
            }
            rootElement = enPart.Endnotes;
        }
        else if (lowerPath is "/theme")
        {
            // BUG-DUMP-R37-5: a theme-LESS source must round-trip with NO theme
            // part — the blank rebuild template auto-stamps theme1.xml, whose
            // minorFont resolves CJK glyphs to a different fallback than Word's
            // app-default theme, tightening CJK line height and drifting body
            // text. EmitThemeRaw emits `action=remove` on /theme when the source
            // had none; delete the whole ThemePart (DeletePart also drops the
            // document.xml.rels theme relationship, so no dangling ref remains).
            // Mirrors EmitDocDefaultsRaw's remove path for a source lacking
            // docDefaults. Idempotent: a re-dump of the rebuilt (theme-less) doc
            // sees no theme and emits the same remove.
            if (string.Equals(action, "remove", StringComparison.OrdinalIgnoreCase)
                && string.Equals(partPath, "/theme", StringComparison.OrdinalIgnoreCase)
                && string.Equals(xpath, "/a:theme", StringComparison.OrdinalIgnoreCase))
            {
                if (mainPart.ThemePart != null)
                    mainPart.DeletePart(mainPart.ThemePart);
                return;
            }
            // CONSISTENCY(raw-set-create-missing-part): blank docs created via
            // BlankDocCreator have no ThemePart; dump→batch round-trip from a
            // real Word/python-docx file emits raw-set /theme replace which
            // would otherwise abort the whole batch. Lazily add the theme part
            // and an empty <a:theme> root so RawXmlHelper.Execute can match
            // /a:theme and replace it with the dumped XML.
            var themePart = mainPart.ThemePart ?? mainPart.AddNewPart<ThemePart>();
            if (themePart.Theme == null)
            {
                themePart.Theme = new DocumentFormat.OpenXml.Drawing.Theme(
                    new DocumentFormat.OpenXml.Drawing.ThemeElements());
                themePart.Theme.Save();
            }
            rootElement = themePart.Theme;
        }
        else if (lowerPath is "/fonttable")
        {
            // BUG-DUMP-R42-3: round-trip word/fontTable.xml (font-face
            // declarations + altName substitutions). The blank rebuild template
            // has no FontTablePart; lazily create one with an empty <w:fonts>
            // root so RawXmlHelper.Execute can match /w:fonts and replace it
            // with the dumped XML. The emitter strips any embedded-font
            // references (w:embedRegular/Bold/Italic/BoldItalic r:id) first, so
            // no dangling rel survives (the .odttf binaries are not round-tripped).
            // CONSISTENCY(raw-set-create-missing-part): mirrors the /theme +
            // /numbering branches.
            var fontTablePart = mainPart.FontTablePart ?? mainPart.AddNewPart<FontTablePart>();
            if (fontTablePart.Fonts == null)
            {
                fontTablePart.Fonts = new DocumentFormat.OpenXml.Wordprocessing.Fonts();
                fontTablePart.Fonts.Save();
            }
            rootElement = fontTablePart.Fonts;
        }
        else if (lowerPath.StartsWith("/header"))
        {
            var idx = 0;
            var bracketIdx = partPath.IndexOf('[');
            if (bracketIdx >= 0)
                int.TryParse(partPath[(bracketIdx + 1)..].TrimEnd(']'), out idx);
            var headerPart = mainPart.HeaderParts.ElementAtOrDefault(idx - 1)
                ?? throw new ArgumentException($"header[{idx}] not found");
            rootElement = headerPart.Header ?? throw new InvalidOperationException($"Corrupt file: header[{idx}] data missing");
        }
        else if (lowerPath.StartsWith("/footer"))
        {
            var idx = 0;
            var bracketIdx = partPath.IndexOf('[');
            if (bracketIdx >= 0)
                int.TryParse(partPath[(bracketIdx + 1)..].TrimEnd(']'), out idx);
            var footerPart = mainPart.FooterParts.ElementAtOrDefault(idx - 1)
                ?? throw new ArgumentException($"footer[{idx}] not found");
            rootElement = footerPart.Footer ?? throw new InvalidOperationException($"Corrupt file: footer[{idx}] data missing");
        }
        else if (lowerPath.StartsWith("/chart"))
        {
            var idx = 0;
            var bracketIdx = partPath.IndexOf('[');
            if (bracketIdx >= 0)
                int.TryParse(partPath[(bracketIdx + 1)..].TrimEnd(']'), out idx);
            var chartPart = mainPart.ChartParts.ElementAtOrDefault(idx - 1)
                ?? throw new ArgumentException($"chart[{idx}] not found");
            rootElement = chartPart.ChartSpace ?? throw new InvalidOperationException($"Corrupt file: chart[{idx}] data missing");
        }
        else if (lowerPath is "/commentsextended")
        {
            // BUG-DUMP-R26-4: round-trip word/commentsExtended.xml (modern
            // comment-reply threading; w15:commentEx paraIdParent links a reply
            // to its parent comment, keyed by the comment paragraphs' w14:paraId).
            // The part has no clean typed-root API across SDK versions, so feed
            // the whole part stream directly. `replace` (the only action the dump
            // emits) overwrites the part body verbatim; the part is lazily created
            // on a blank rebuilt doc. The linkage stays valid because EmitComments
            // now preserves the comment paragraphs' source w14:paraId.
            var exPart = mainPart.WordprocessingCommentsExPart
                ?? mainPart.AddNewPart<WordprocessingCommentsExPart>();
            if (!string.Equals(action, "replace", StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException(
                    $"/commentsExtended supports only the 'replace' action (got '{action}').");
            if (string.IsNullOrEmpty(xml))
                throw new ArgumentException("/commentsExtended replace requires XML.");
            using (var ms = new System.IO.MemoryStream(System.Text.Encoding.UTF8.GetBytes(xml)))
                exPart.FeedData(ms);
            EnsureAllParaIds();
            return;
        }
        else
            throw new ArgumentException($"Unknown part: {partPath}. Available: /document, /styles, /settings, /numbering, /header[n], /footer[n], /chart[n]");

        var affected = RawXmlHelper.Execute(rootElement, xpath, action, xml);
        // During a DeferSave batch, defer the whole-part re-serialize and the
        // four global id sweeps to FinalizeDeferredIds (one Save per touched
        // root, one sweep total) instead of paying them per raw-set. The
        // Execute above already applied the edit to the in-memory SDK DOM, so
        // the deferred passes and the final _doc.Save() see the changes. Outside
        // a batch (one-shot raw-set / resident mid-session) keep the eager path.
        if (DeferSave)
        {
            (_deferredRawSetRoots ??= new HashSet<OpenXmlPartRootElement>()).Add(rootElement);
            _ = affected;
            return;
        }
        rootElement.Save();
        // CONSISTENCY(paraid-global-uniqueness): RawSet may inject paragraphs
        // carrying paraIds the handler hasn't seen — without re-scanning,
        // _usedParaIds and _nextParaId stay stale and the next AddBreak /
        // AddParagraph could allocate a colliding paraId. Especially
        // dangerous in resident mode where one process serves many commands
        // across the same _usedParaIds set. Re-run EnsureAllParaIds after
        // every successful raw mutation so the global pool stays accurate.
        EnsureAllParaIds();
        // CONSISTENCY(docpr-global-uniqueness): same hazard for <wp:docPr> ids.
        // RawSet may inject a drawing whose docPr id collides with one a typed
        // add already allocated via NextDocPropId (e.g. dump preserving a
        // non-textbox shape replays its source id verbatim while the textbox
        // alongside it was renumbered onto the same value). EnsureDocPropIds
        // otherwise only runs at open, so an in-session raw mutation would
        // leave the duplicate on disk; re-run it here to dedupe to the lowest
        // free id — the same correction paraId gets above.
        EnsureDocPropIds();
        // CONSISTENCY(sdt-global-uniqueness): same hazard for sdt <w:id>.
        // NextSdtId allocates max+1 for typed adds, but a raw-set can inject an
        // sdt carrying a colliding id. EnsureSdtIds otherwise runs only at open.
        EnsureSdtIds();
        // CONSISTENCY(bookmark-global-uniqueness): same hazard for bookmark
        // <w:id>. AddBookmark scans body bookmarkStarts for max+1, but a raw-set
        // (verbatim <w:sdt> cover block) can inject a bookmark whose id collides
        // with one a structured add already used. EnsureBookmarkIds otherwise
        // runs only at open.
        EnsureBookmarkIds();
        // BUG-R5-01: do not emit chatter from inside the handler — the CLI
        // wrappers (CommandBuilder.Raw raw-set + batch run raw-set) print
        // their own structured message. Writing here pollutes batch --json
        // output (extra stdout lines escaped into result.message strings).
        _ = affected;
    }

    // Compiled matcher for the dump's single-style raw-set xpath
    // (/w:styles/w:style[@w:styleId='ID']). See the /styles fast-path in RawSet.
    private static readonly System.Text.RegularExpressions.Regex _singleStyleByIdXpath =
        new(@"^/w:styles/w:style\[@w:styleId='([^']*)'\]$",
            System.Text.RegularExpressions.RegexOptions.Compiled);

    // Replace exactly one <w:style> (matched by styleId) on the live styles DOM,
    // bypassing the generic whole-part serialize/parse round-trip. Returns false
    // (caller falls through to the generic path) for any xpath that is not the
    // exact single-style-by-id shape, a styleId not present, or a fragment the
    // SDK cannot parse — so it is a pure optimization with no behavior change.
    // A <w:style> carries no paraId / docPr / sdt / bookmark, so the post-replace
    // global id sweeps the generic path runs are no-ops here and are skipped; the
    // whole-part Save is deferred to FinalizeDeferredIds during a batch, mirroring
    // the generic /styles raw-set under DeferSave.
    private bool TryReplaceSingleStyleFast(MainDocumentPart mainPart, string xpath, string xml)
    {
        var m = _singleStyleByIdXpath.Match(xpath);
        if (!m.Success) return false;
        var styleId = m.Groups[1].Value;
        var styles = mainPart.StyleDefinitionsPart?.Styles;
        if (styles == null) return false;
        var target = styles.Elements<Style>()
            .FirstOrDefault(s => string.Equals(s.StyleId?.Value, styleId, StringComparison.Ordinal));
        if (target == null) return false;
        Style newStyle;
        try { newStyle = new Style(xml); }
        catch { return false; }
        target.InsertBeforeSelf(newStyle);
        target.Remove();
        if (DeferSave)
            (_deferredRawSetRoots ??= new HashSet<OpenXmlPartRootElement>()).Add(styles);
        else
            styles.Save();
        return true;
    }

    // Compiled matcher for a raw-set xpath rooted at a single table by global
    // document-order ordinal: (//w:tbl)[N] optionally followed by a child path.
    // See the cell-append fast-path in RawSet.
    private static readonly System.Text.RegularExpressions.Regex _tableCellAppendXpath =
        new(@"^\(//w:tbl\)\[(\d+)\]/w:tr\[(\d+)\]/w:tc\[(\d+)\]$",
            System.Text.RegularExpressions.RegexOptions.Compiled);

    // Namespace prefixes declared on the synthetic <w:body> wrapper used to parse
    // a verbatim cell-content fragment, so a fragment that relies on an inherited
    // prefix (one not among its own inline xmlns) still parses. Mirrors the prefix
    // set the generic raw-set path has in scope; Word-specific and intentionally
    // narrow. A fragment whose prefix is missing here simply fails to parse and
    // falls through to the generic path — never silently wrong.
    private static readonly (string prefix, string uri)[] _wordFragmentNs =
    {
        ("w",   "http://schemas.openxmlformats.org/wordprocessingml/2006/main"),
        ("r",   "http://schemas.openxmlformats.org/officeDocument/2006/relationships"),
        ("a",   "http://schemas.openxmlformats.org/drawingml/2006/main"),
        ("wp",  "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"),
        ("mc",  "http://schemas.openxmlformats.org/markup-compatibility/2006"),
        ("pic", "http://schemas.openxmlformats.org/drawingml/2006/picture"),
        ("w14", "http://schemas.microsoft.com/office/word/2010/wordml"),
        ("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"),
        ("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"),
        ("a14", "http://schemas.microsoft.com/office/drawing/2010/main"),
        ("v",   "urn:schemas-microsoft-com:vml"),
        ("o",   "urn:schemas-microsoft-com:office:office"),
    };

    // Append a verbatim cell-content fragment — the dump's
    // raw-set (//w:tbl)[N]/w:tr[M]/w:tc[K] append for a rich field result / nested
    // SDT / VML textbox — directly onto the live cell, bypassing the generic
    // whole-document serialize/parse round-trip. The generic path re-serializes the
    // whole multi-MB document part per op (rootElement.OuterXml + XDocument.Parse +
    // InnerXml=), making a table-heavy batch O(ops × part) — minutes for a
    // 757-table report. Here the existing table content is never re-serialized
    // (which is what would otherwise stamp redundant xmlns onto every rebuilt
    // tbl/tr/tc/p); only the fragment is parsed and grafted, so the saved bytes
    // match the generic path (both keep the fragment's own inline xmlns).
    // Returns false — caller falls through to the proven generic path — for any
    // xpath that is not the exact cell-append shape, an out-of-range index, or a
    // fragment the SDK cannot parse. Gated to DeferSave (batch): the generic path
    // runs the per-op global id sweeps (paraId/docPr/sdt/bookmark) eagerly outside
    // a batch, and verbatim cell content carries those ids; under DeferSave they
    // run once in FinalizeDeferredIds over the deferred roots, which this mirrors.
    private bool TryRawSetWithinTableFast(MainDocumentPart mainPart, string xpath, string action, string? xml)
    {
        if (!DeferSave) return false;
        if (xml == null) return false;
        if (!string.Equals(action, "append", StringComparison.OrdinalIgnoreCase)) return false;
        var m = _tableCellAppendXpath.Match(xpath);
        if (!m.Success) return false;
        var document = mainPart.Document;
        if (document == null) return false;
        if (!int.TryParse(m.Groups[1].Value, out var n) || n < 1) return false;
        if (!int.TryParse(m.Groups[2].Value, out var r) || r < 1) return false;
        if (!int.TryParse(m.Groups[3].Value, out var c) || c < 1) return false;
        // (//w:tbl)[N] — Nth <w:tbl> in document order (the SDK's Descendants<T>()
        // walk is pre-order DFS, the same order XPath's // axis yields); then the
        // Mth direct <w:tr> child and that row's Kth direct <w:tc> child, matching
        // the literal /w:tr[M]/w:tc[K] child-axis steps. Any out-of-range index
        // falls through to the generic path.
        var table = document.Descendants<Table>().ElementAtOrDefault(n - 1);
        if (table == null) return false;
        var row = table.Elements<TableRow>().ElementAtOrDefault(r - 1);
        if (row == null) return false;
        var cell = row.Elements<TableCell>().ElementAtOrDefault(c - 1);
        if (cell == null) return false;
        // Parse the fragment inside a synthetic <w:body> that declares the common
        // Word prefixes (so a fragment relying on an inherited prefix still parses)
        // then graft its children onto the live cell. The fragment keeps its own
        // inline xmlns exactly as the generic path's append does, and — crucially —
        // the surrounding table is left untouched, so no redundant xmlns is stamped
        // onto existing content.
        var sb = new System.Text.StringBuilder("<w:body");
        foreach (var (prefix, uri) in _wordFragmentNs)
            sb.Append(" xmlns:").Append(prefix).Append("=\"").Append(uri).Append('"');
        sb.Append('>').Append(xml).Append("</w:body>");
        Body holder;
        try { holder = new Body(sb.ToString()); }
        catch { return false; }
        var appended = holder.ChildElements.ToList();
        if (appended.Count == 0) return false;
        foreach (var child in appended)
        {
            child.Remove();
            cell.AppendChild(child);
        }
        (_deferredRawSetRoots ??= new HashSet<OpenXmlPartRootElement>()).Add(document);
        return true;
    }

    // Compiled matcher for the dump's cross-paragraph field-member xpath
    // (/w:document/w:body/w:sectPr). See the body-para fast-path in RawSet.
    private static readonly System.Text.RegularExpressions.Regex _bodySectPrXpath =
        new(@"^/w:document/w:body/w:sectPr$",
            System.Text.RegularExpressions.RegexOptions.Compiled);

    // Insert exactly one <w:p> immediately before the body's trailing sectPr on
    // the live document DOM, bypassing the generic whole-part serialize/parse
    // round-trip. Returns false (caller falls through to the generic path) for any
    // xpath that is not the exact /w:document/w:body/w:sectPr shape, a missing
    // body/sectPr, or a fragment the SDK cannot parse as a <w:p> — so it is a pure
    // optimization with no behavior change. Gated to DeferSave (batch): outside a
    // batch the generic path runs the per-op global id sweeps (paraId/docPr/sdt/
    // bookmark) eagerly, and a verbatim <w:p> can carry those ids; under DeferSave
    // those sweeps are deferred to FinalizeDeferredIds exactly as the generic
    // DeferSave raw-set path defers them, so the fast path matches its behavior.
    private bool TryInsertBodyParaBeforeSectPrFast(MainDocumentPart mainPart, string xpath, string xml)
    {
        if (!DeferSave) return false;
        if (!_bodySectPrXpath.IsMatch(xpath)) return false;
        var body = mainPart.Document?.Body;
        if (body == null) return false;
        // The body-level sectPr is always the last child of <w:body> (ECMA-376).
        if (body.LastChild is not SectionProperties sectPr)
            return false; // no trailing sectPr to anchor before — fall through
        Paragraph newPara;
        try { newPara = new Paragraph(xml); }
        catch { return false; }
        if (newPara.LocalName != "p") return false; // fragment must be a paragraph
        // O(1) insert via the append-monotonic cache (mirrors AppendBodyParaFast):
        // the cached last body paragraph sits right before sectPr, so
        // InsertAfterSelf keeps sectPr last and avoids the SDK singly-linked
        // list's O(N) InsertBefore predecessor scan (which made N consecutive
        // member inserts O(N²)).
        if (_bodyParaCount >= 0
            && _lastBodyParagraph is Paragraph anchor
            && ReferenceEquals(anchor.Parent, body)
            && anchor.NextSibling() is SectionProperties)
        {
            anchor.InsertAfterSelf(newPara);
            _bodyParaCount++;
        }
        else
        {
            // Cold cache / out-of-band mutation: O(N) insert + reseed.
            sectPr.InsertBeforeSelf(newPara);
            _bodyParaCount = body.Elements<Paragraph>().Count();
        }
        _lastBodyParagraph = newPara;
        (_deferredRawSetRoots ??= new HashSet<OpenXmlPartRootElement>()).Add(mainPart.Document!);
        return true;
    }

    // BUG-DUMP-R45-1 / BUG-DUMP-R45-2: apply an `embed-binary` companion op.
    // partPath selects the host part (/fontTable → FontPart, /numbering →
    // ImagePart); xpath carries the source relationship id to reuse; xml is a
    // `data:<contentType>;base64,…` URI. The FontTablePart / NumberingDefinitions
    // Part is created lazily by the preceding `raw-set replace`, so it already
    // exists here. AddNewPart<T>(contentType, id) attaches the part with the
    // exact r:id, so the verbatim <w:embed*> / <v:imagedata> ref resolves.
    private void RawEmbedBinary(MainDocumentPart mainPart, string partPath, string? xpath, string? xml)
    {
        var relId = xpath;
        if (string.IsNullOrEmpty(relId))
            throw new ArgumentException("embed-binary requires the relationship id in 'xpath'.");
        if (!OleHelper.TryDecodeDataUri(xml, out var bytes, out var contentType) || bytes.Length == 0)
            throw new ArgumentException("embed-binary requires a non-empty data: URI in 'xml'.");

        var lowerPath = partPath.ToLowerInvariant();
        if (lowerPath is "/fonttable")
        {
            var ftp = mainPart.FontTablePart
                ?? throw new InvalidOperationException("No fontTable part for embed-binary");
            var ct = string.IsNullOrEmpty(contentType)
                ? "application/vnd.openxmlformats-officedocument.obfuscatedFont"
                : contentType;
            var fp = ftp.AddNewPart<FontPart>(ct, relId);
            using var ms = new MemoryStream(bytes);
            fp.FeedData(ms);
        }
        else if (lowerPath is "/numbering")
        {
            var ndp = mainPart.NumberingDefinitionsPart
                ?? throw new InvalidOperationException("No numbering part for embed-binary");
            var ct = string.IsNullOrEmpty(contentType) ? "image/png" : contentType;
            var ip = ndp.AddNewPart<ImagePart>(ct, relId);
            using var ms = new MemoryStream(bytes);
            ip.FeedData(ms);
        }
        else if (lowerPath is "/customxml")
        {
            // customXml data store item — recreate with the SOURCE rel id so
            // the companion props op can address it as /customXml/<relId>.
            // Nothing in any XML references this id, so reuse is collision-
            // safe (the rebuilt main part's auto ids use a different scheme).
            var ct = string.IsNullOrEmpty(contentType) ? "application/xml" : contentType;
            var cxp = mainPart.AddNewPart<CustomXmlPart>(ct, relId);
            using var ms = new MemoryStream(bytes);
            cxp.FeedData(ms);
        }
        else if (lowerPath.StartsWith("/customxml/", StringComparison.Ordinal))
        {
            // itemProps child of a previously attached customXml item part.
            var itemRelId = partPath.Substring("/customXml/".Length);
            OpenXmlPart item;
            try { item = mainPart.GetPartById(itemRelId); }
            catch
            {
                throw new InvalidOperationException(
                    $"embed-binary: no customXml item part with rel id '{itemRelId}'");
            }
            if (item is not CustomXmlPart customItem)
                throw new InvalidOperationException(
                    $"embed-binary: rel id '{itemRelId}' is not a customXml part");
            var ct = string.IsNullOrEmpty(contentType)
                ? "application/vnd.openxmlformats-officedocument.customXmlProperties+xml"
                : contentType;
            var propsPart = customItem.AddNewPart<CustomXmlPropertiesPart>(ct, relId);
            using var ms = new MemoryStream(bytes);
            propsPart.FeedData(ms);
        }
        else
        {
            throw new ArgumentException(
                $"embed-binary not supported for part '{partPath}' (only /fontTable, /numbering, /customXml).");
        }
    }

    // The three docProps stores whose whole-part replace must take the
    // staged-zip-rewrite path (the SDK won't persist them mid-session).
    private static bool IsDocPropsWholePart(string partPath)
    {
        var p = partPath.ToLowerInvariant();
        return p is "/docprops/core.xml" or "/docprops/app.xml" or "/docprops/custom.xml";
    }

    // Stage a docProps whole-part replace. EmitDocPropsRaw round-trips
    // core/app/custom.xml verbatim so data-bound content controls (cover title /
    // company / contact) keep their source-authored display text. The payload is
    // STAGED here, not written now: the Open-XML SDK does not reliably persist a
    // docProps mutation made mid-session on a stream-opened editable package —
    // both a typed-root assignment and a direct part-stream write are silently
    // reverted to the open-time content when _doc.Save() flushes (verified
    // empirically; even the documented `set / --prop extended.company=…` path is
    // lost the same way). The staged XML is written straight into the saved zip
    // after the package closes (FlushPendingWholeParts), the same post-close
    // zip-rewrite mechanism NormalizeSelfClosingInDocx uses for document.xml.
    //
    // Verbatim for all three: a dump→batch rebuild reproduces the source, so the
    // source authoring identity (app.xml Application, core.xml creator,
    // custom.xml user props) is the faithful result — the OfficeCLI audit stamp
    // is a create/edit-time concern, not a reconstruction one.
    private void StashWholePartReplace(string partPath, string? xml)
    {
        if (string.IsNullOrWhiteSpace(xml))
            throw new ArgumentException("raw-set replace on docProps requires non-empty 'xml'.");
        var content = xml.Trim();

        var entryName = (partPath.StartsWith('/') ? partPath[1..] : partPath);
        // Ensure a standard prolog so the rewritten part matches Office output
        // (the staged content is element-only — Raw() strips the prolog).
        if (!content.StartsWith("<?xml", StringComparison.Ordinal))
            content = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" + content;

        (_pendingWholeParts ??= new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase))
            [entryName] = content;
        Modified = true;
    }

    // Write every staged whole-part payload into the saved zip on disk,
    // replacing the entry the SDK wrote. The caller must have released the
    // file first (_doc + _backingStream closed) — this runs from Dispose,
    // after _doc.Dispose(). Clears the staging map after writing.
    //
    // Resident note: a mid-session flush (the `save` command or the idle-autosave)
    // calls handler.Save() while keeping the package open, so staged docProps land
    // on disk only when the resident closes and the handler Disposes — a narrower
    // rule than the general "resident flushes on save/close/idle-autosave" model
    // (ordinary edits do flush mid-session; only these staged whole-parts wait for
    // close). dump→batch rebuilds run non-resident (OFFICECLI_NO_AUTO_RESIDENT=1),
    // so this gap does not affect the round-trip path.
    /// <summary>
    /// Atomic-batch rollback support. Staged whole-part payloads
    /// (docProps raw-set replacements) are session state that only lands on
    /// disk at a NON-discard Dispose — handler.Save() cannot flush them, so
    /// the resident's pre-batch flush barrier does not cover them. A rollback
    /// therefore snapshots the staging map BEFORE the batch (batch-staged
    /// entries must roll back too) and adopts the snapshot into the
    /// replacement handler, instead of dropping confirmed pre-batch edits
    /// with the poisoned DOM.
    /// </summary>
    internal Dictionary<string, string>? SnapshotPendingWholeParts()
        => _pendingWholeParts == null ? null
            : new Dictionary<string, string>(_pendingWholeParts, StringComparer.OrdinalIgnoreCase);

    internal void AdoptPendingWholeParts(Dictionary<string, string>? staged)
    {
        if (staged == null || staged.Count == 0) return;
        _pendingWholeParts = new Dictionary<string, string>(staged, StringComparer.OrdinalIgnoreCase);
        Modified = true;
    }

    private void FlushPendingWholeParts(string path)
    {
        if (_pendingWholeParts == null || _pendingWholeParts.Count == 0) return;
        var pending = _pendingWholeParts;
        _pendingWholeParts = null;
        if (!System.IO.File.Exists(path)) return;
        try
        {
            using var fs = new System.IO.FileStream(path, System.IO.FileMode.Open, System.IO.FileAccess.ReadWrite);
            using var za = new System.IO.Compression.ZipArchive(fs, System.IO.Compression.ZipArchiveMode.Update, leaveOpen: false);
            foreach (var (entryName, content) in pending)
            {
                za.GetEntry(entryName)?.Delete();
                var entry = za.CreateEntry(entryName);
                using var w = new System.IO.StreamWriter(entry.Open(), new System.Text.UTF8Encoding(false));
                w.Write(content);
            }
        }
        catch { /* best-effort: a malformed package or locked file leaves the
                   SDK-saved docProps in place rather than aborting the save */ }
    }


    /// <summary>
    /// BUG-DUMP-R26-4: read word/commentsExtended.xml verbatim for the dump
    /// emitter, or null when the part is absent. The part carries modern
    /// comment-reply threading (w15:commentEx paraIdParent) keyed by comment
    /// paragraph w14:paraId; EmitComments raw-sets this XML back on replay and
    /// preserves the comment paraIds so the linkage survives round-trip.
    /// </summary>
    internal string? GetCommentsExtendedXml()
    {
        var exPart = _doc.MainDocumentPart?.WordprocessingCommentsExPart;
        if (exPart == null) return null;
        using var reader = new System.IO.StreamReader(exPart.GetStream(System.IO.FileMode.Open, System.IO.FileAccess.Read));
        var xml = reader.ReadToEnd();
        return string.IsNullOrWhiteSpace(xml) ? null : xml;
    }

    public List<ValidationError> Validate() => RawXmlHelper.ValidateDocument(_doc, _filePath);

    /// <summary>
    /// Run the global id-uniqueness passes once, just before a deferred-batch
    /// flush. During <see cref="DeferSave"/> replay the per-mutation id
    /// management that normally runs after each raw-set is skipped, so an
    /// `add p` carrying a paragraph-mark revision id preserved from the source
    /// can collide with a content-run ins/del id generated by a later item.
    /// EnsureAllParaIds (which also dedupes revision w:ids), EnsureDocPropIds and
    /// EnsureSdtIds reconcile the whole document in one O(N) sweep at the end,
    /// matching the after-raw-set normalization but without the per-op O(N²) cost
    /// the deferred path exists to avoid. Best-effort; never block the save.
    /// </summary>
    private void FinalizeDeferredIds()
    {
        // Flush the whole-part re-serialize that each batch raw-set deferred:
        // one Save per touched root, not one per op. Must run before the id
        // sweeps so the sweeps (and the subsequent _doc.Save) see the persisted
        // DOM. The edits already live in the in-memory SDK DOM via Execute's
        // InnerXml set, so this is purely the serialization the per-op path
        // would have done.
        if (_deferredRawSetRoots != null)
        {
            foreach (var root in _deferredRawSetRoots)
            {
                try { root.Save(); }
                catch { /* best-effort; a corrupt root must not block the flush */ }
            }
            _deferredRawSetRoots = null;
        }
        try
        {
            EnsureAllParaIds();
            EnsureDocPropIds();
            EnsureSdtIds();
            EnsureBookmarkIds();
        }
        catch { /* id normalization is best-effort; never block the flush */ }
        try { NormalizeAllRunPropsSchemaOrder(); }
        catch { /* schema-order normalization is best-effort; never block the flush */ }
        try { EnsureUniqueMoveRunIds(); }
        catch { /* move-id dedup is best-effort; never block the flush */ }
    }

    // BUG-DUMP-R71-MOVE-DUP-ID: a move that spans multiple runs/paragraphs
    // round-trips through per-run `add r revision.type=moveFrom revision.id=X`,
    // each wrapping its run in a <w:moveFrom w:id="X">; N runs then produce N
    // MoveFromRun elements all sharing id X. The strict validator rejects
    // same-type duplicate w:id. The CLI deliberately shares the id within a
    // move so the Move_{id} range-marker NAME pairs the moveFrom with its moveTo
    // (see BuildMovePairIdMap / WrapRunAsMoveFrom) — and that pairing rides on
    // the range-marker name, NOT on the content-run id. So at batch finalization
    // we can re-assign each MoveFromRun / MoveToRun a unique id (leaving the
    // range markers and their names untouched) to satisfy validation without
    // disturbing the move pairing. Word's own files use unique content-run ids
    // with name-based pairing, so this matches the native shape. Interactive
    // (non-DeferSave) WrapRunAsMove* is unchanged — this pass runs only here.
    private void EnsureUniqueMoveRunIds()
    {
        var main = _doc.MainDocumentPart;
        if (main == null) return;
        IEnumerable<OpenXmlElement?> roots = new OpenXmlElement?[] { main.Document?.Body }
            .Concat(main.HeaderParts.Select(h => (OpenXmlElement?)h.Header))
            .Concat(main.FooterParts.Select(f => (OpenXmlElement?)f.Footer))
            .Append(main.FootnotesPart?.Footnotes)
            .Append(main.EndnotesPart?.Endnotes);
        var rootList = roots.Where(r => r != null).Cast<OpenXmlElement>().ToList();

        // Allocate fresh ids above the global max of every move-element w:id in
        // play so a reassigned move-run id can't collide with another.
        long next = 0;
        void TrackMax(string? v) { if (v != null && long.TryParse(v, out var n) && n > next) next = n; }
        foreach (var root in rootList)
        {
            foreach (var e in root.Descendants<MoveFromRun>()) TrackMax(e.Id?.Value);
            foreach (var e in root.Descendants<MoveToRun>()) TrackMax(e.Id?.Value);
            foreach (var e in root.Descendants<MoveFromRangeStart>()) TrackMax(e.Id?.Value);
            foreach (var e in root.Descendants<MoveFromRangeEnd>()) TrackMax(e.Id?.Value);
            foreach (var e in root.Descendants<MoveToRangeStart>()) TrackMax(e.Id?.Value);
            foreach (var e in root.Descendants<MoveToRangeEnd>()) TrackMax(e.Id?.Value);
        }

        // Per-type uniqueness: the validator only flags same-type duplicate w:id,
        // so an element may share an id with a different-typed move element (one
        // move's worth is fine) — the fault is N elements of the SAME type all
        // carrying the collapsed pairing id (N MoveFromRun content wrappers, or
        // the per-fragment range markers a cross-paragraph span emits). Keep the
        // first occurrence's id (the lone single-run case is untouched) and
        // re-stamp the rest. Range-marker w:name is left intact, so the
        // Move_{id} name pairing — the only thing accept/reject and the dump
        // round-trip rely on — survives.
        void DedupById(Func<OpenXmlElement, string?> get, Action<OpenXmlElement, string> set,
                       Func<OpenXmlElement, bool> match)
        {
            var seen = new HashSet<string>(StringComparer.Ordinal);
            foreach (var root in rootList)
                foreach (var el in root.Descendants().Where(match))
                {
                    var cur = get(el);
                    if (string.IsNullOrEmpty(cur)) continue;
                    if (!seen.Add(cur!))
                        set(el, (++next).ToString(System.Globalization.CultureInfo.InvariantCulture));
                }
        }
        DedupById(e => (e as MoveFromRun)?.Id?.Value, (e, v) => ((MoveFromRun)e).Id = v, e => e is MoveFromRun);
        DedupById(e => (e as MoveToRun)?.Id?.Value, (e, v) => ((MoveToRun)e).Id = v, e => e is MoveToRun);
        DedupById(e => (e as MoveFromRangeStart)?.Id?.Value, (e, v) => ((MoveFromRangeStart)e).Id = v, e => e is MoveFromRangeStart);
        DedupById(e => (e as MoveFromRangeEnd)?.Id?.Value, (e, v) => ((MoveFromRangeEnd)e).Id = v, e => e is MoveFromRangeEnd);
        DedupById(e => (e as MoveToRangeStart)?.Id?.Value, (e, v) => ((MoveToRangeStart)e).Id = v, e => e is MoveToRangeStart);
        DedupById(e => (e as MoveToRangeEnd)?.Id?.Value, (e, v) => ((MoveToRangeEnd)e).Id = v, e => e is MoveToRangeEnd);
    }

    // BUG-DUMP-R71-RPR-ORDER: document-wide CT_RPr / CT_ParaRPr child-order
    // normalization, run once at batch finalization (after the raw-set flush so
    // it also catches verbatim-injected runs). A run/paragraph-mark rPr can be
    // assembled across the curated Add path, ApplyRunFormatting, raw AppendChild,
    // TypedAttributeFallback tail-appends, and verbatim raw-set — any of which
    // can leave a child out of CT_RPr order (sz after u, ins after rStyle),
    // which strict OOXML validation rejects. Re-seat every standard (non-w14)
    // rPr child into its schema slot. Content-preserving (pure reorder); only
    // touches rPr that were already out of order. Covers the main document plus
    // header/footer/footnote/endnote parts.
    private void NormalizeAllRunPropsSchemaOrder()
    {
        var main = _doc.MainDocumentPart;
        if (main == null) return;
        IEnumerable<OpenXmlElement?> roots = new OpenXmlElement?[] { main.Document?.Body }
            .Concat(main.HeaderParts.Select(h => (OpenXmlElement?)h.Header))
            .Concat(main.FooterParts.Select(f => (OpenXmlElement?)f.Footer))
            .Append(main.FootnotesPart?.Footnotes)
            .Append(main.EndnotesPart?.Endnotes);
        foreach (var root in roots)
        {
            if (root == null) continue;
            foreach (var rPr in root.Descendants<RunProperties>().ToList())
                NormalizeRunPropsSchemaOrder(rPr);
            foreach (var pmRpr in root.Descendants<ParagraphMarkRunProperties>().ToList())
            {
                NormalizeRunPropsSchemaOrder(pmRpr);
                // The paragraph-mark rPr itself must sit at its CT_PPr slot
                // (after numPr/spacing, before sectPr/pPrChange); a verbatim or
                // mixed-path pPr build can land it out of order ("unexpected
                // child element rPr"). Re-seat the element within its pPr.
                if (pmRpr.Parent is OpenXmlCompositeElement pPrParent)
                    Core.SchemaOrder.Place(pPrParent, pmRpr);
            }
        }
    }

    /// <summary>
    /// BUG-R7B(BUG2): document-wide id reconciliation, callable mid-session.
    /// Under <see cref="DeferSave"/> the per-raw-set id passes
    /// (EnsureDocPropIds / EnsureAllParaIds / EnsureSdtIds) are skipped, so a
    /// batch that raw-sets header/footer parts can leave the in-memory document
    /// with colliding wp:docPr ids until the next save/close runs
    /// <see cref="FinalizeDeferredIds"/>. A `validate` (or any flush) issued
    /// between the batch and the save would then see the duplicates. The
    /// resident batch path calls this right after the batch loop so the
    /// in-memory tree matches the on-disk state save/close would produce —
    /// the same document-scoped passes, run once.
    /// </summary>
    public void ReconcileGlobalIds() => FinalizeDeferredIds();

    public void Save()
    {
        if (DeferSave) FinalizeDeferredIds();
        // Mid-session flush. The Dispose-time NormalizeSelfClosingInDocx step
        // is intentionally skipped here — it requires opening the file as a
        // Zip with read-write access, which can't be done while the backing
        // stream still holds the file. The on-disk snapshot will have
        // `<w:br />` form instead of the canonical `<w:br/>` form; both are
        // schema-valid OOXML.
        if (Modified)
        {
            try { OfficeCli.Core.OfficeCliMetadata.StampOnSave(_doc); }
            catch { /* best-effort audit trail */ }
        }
        _doc.Save();
        if (_packageStream != null) AtomicWriteBack();
        else _backingStream?.Flush();
    }

    // Crash-atomic flush of the in-memory package to disk (temp + File.Replace).
    // No-op for read-only sessions (_packageStream == null), which never write.
    // See OfficeCli.Core.AtomicPackageWriter.
    private void AtomicWriteBack(Action<string>? postProcessTemp = null)
    {
        if (_packageStream == null || _backingStream == null) return;
        OfficeCli.Core.AtomicPackageWriter.Flush(
            _packageStream, _filePath,
            releaseLock: () => { _backingStream!.Dispose(); _backingStream = null; },
            reopenLock: () => { _backingStream = new FileStream(_filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.Read); },
            postProcessTemp: postProcessTemp);
    }

    public void Dispose()
    {
        if (DiscardOnDispose)
        {
            // Atomic-batch rollback: never serialize the poisoned DOM. Close
            // the backing stream FIRST so the package's dispose-time autosave
            // has nowhere to write, then swallow its complaint. The on-disk
            // file keeps the last flushed (pre-batch) state.
            _backingStream?.Dispose();
            _backingStream = null;
            try { _doc.Dispose(); } catch { /* autosave hit the closed stream — intended */ }
            _packageStream?.Dispose();
            _packageStream = null;
            return;
        }
        if (DeferSave) FinalizeDeferredIds();
        if (Modified)
        {
            try { OfficeCli.Core.OfficeCliMetadata.StampOnSave(_doc); }
            catch { /* best-effort audit trail */ }
        }
        try { _doc.Save(); } catch { /* read-only or already disposed */ }
        if (_packageStream != null)
        {
            // Editable: read the in-memory package into the atomic temp BEFORE
            // _doc.Dispose() (the SDK may close the stream it was opened over on
            // dispose), applying the post-close in-place rewrites to the TEMP so
            // the whole result lands in one atomic File.Replace. The rewrites:
            //  - FlushPendingWholeParts: docProps whole-part edits the SDK won't
            //    persist on a stream-opened package (see RawReplaceWholePart).
            //  - NormalizeSelfClosingInDocx: canonicalize `<w:br />` -> `<w:br/>`
            //    (schema-equivalent; several consumers/tests want the short form).
            try
            {
                AtomicWriteBack(tmp =>
                {
                    FlushPendingWholeParts(tmp);
                    NormalizeSelfClosingInDocx(tmp);
                });
            }
            catch { /* best-effort: the original file is left intact */ }
            _doc.Dispose();
            _packageStream.Dispose();
            _packageStream = null;
            _backingStream?.Dispose();
            _backingStream = null;
        }
        else
        {
            // Read-only session (no _packageStream): we never wrote anything,
            // so leave the on-disk bytes strictly untouched. The legacy
            // post-close NormalizeSelfClosingInDocx ran here too, but a zip
            // opened in ZipArchiveMode.Update rewrites the whole archive on
            // dispose UNCONDITIONALLY — even when no entry changed. Files
            // officecli wrote are already in .NET's container form, so the
            // rewrite was byte-idempotent and invisible; any foreign producer
            // (Word, POI) saw a pure view/get/query/validate change the
            // file's bytes and mtime on first read (issue #231). Normalization still runs on every editable save
            // via AtomicWriteBack above, which is the only path that may
            // legitimately alter the file. _pendingWholeParts is only staged
            // by mutations, which require an editable session, so nothing can
            // be pending here.
            _doc.Dispose();
            _backingStream?.Dispose();
            _backingStream = null;
        }
    }

    private static void NormalizeSelfClosingInDocx(string path)
    {
        if (!System.IO.File.Exists(path)) return;
        using var fs = new System.IO.FileStream(path, System.IO.FileMode.Open, System.IO.FileAccess.ReadWrite);
        using var za = new System.IO.Compression.ZipArchive(fs, System.IO.Compression.ZipArchiveMode.Update, leaveOpen: false);
        var entry = za.GetEntry("word/document.xml");
        if (entry == null) return;
        string xml;
        using (var rs = entry.Open())
        using (var sr = new System.IO.StreamReader(rs))
            xml = sr.ReadToEnd();
        // Collapse "<w:br />" → "<w:br/>" and "<w:tab />" → "<w:tab/>"
        // (no-attribute empty elements only).
        var normalized = System.Text.RegularExpressions.Regex.Replace(
            xml, @"<w:(br|tab) />", "<w:$1/>");
        if (normalized == xml) return;
        entry.Delete();
        var newEntry = za.CreateEntry("word/document.xml");
        using var ws = newEntry.Open();
        using var sw = new System.IO.StreamWriter(ws, new System.Text.UTF8Encoding(false));
        sw.Write(normalized);
    }

    // (private helpers, navigation, selector, style/list, image helpers moved to Word/ partial files)
}
