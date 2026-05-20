// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler : IDocumentHandler
{
    private readonly PresentationDocument _doc;
    private readonly string _filePath;
    private HashSet<uint> _usedShapeIds = new();
    private uint _nextShapeId = 10000;
    public int LastFindMatchCount { get; internal set; }

    // Backing FileStream when we open via stream (shared-read mode). null
    // when the package owns its own file handle via PresentationDocument.Open(path).
    private FileStream? _backingStream;

    public PowerPointHandler(string filePath, bool editable)
    {
        _filePath = filePath;
        // Open via a shared FileStream so external readers (e.g. test harness
        // ZipFile.OpenRead while the handler is alive) don't hit the macOS
        // flock exclusive lock that PresentationDocument.Open(path, editable)
        // would acquire. The package writes through to the stream; we call
        // _doc.Save() in Dispose() to flush before closing the stream.
        var share = editable ? FileShare.Read : FileShare.ReadWrite;
        var access = editable ? FileAccess.ReadWrite : FileAccess.Read;
        _backingStream = new FileStream(filePath, FileMode.Open, access, share);
        _doc = PresentationDocument.Open(_backingStream, editable);
        if (editable)
            InitShapeIdCounter();
    }

    /// <summary>
    /// Get the slide dimensions from the presentation. Falls back to 16:9 (33.867cm × 19.05cm).
    /// </summary>
    private (long width, long height) GetSlideSize()
    {
        var sldSz = _doc.PresentationPart?.Presentation?.GetFirstChild<SlideSize>();
        return (sldSz?.Cx?.Value ?? SlideSizeDefaults.Widescreen16x9Cx, sldSz?.Cy?.Value ?? SlideSizeDefaults.Widescreen16x9Cy);
    }

    // ==================== Raw Layer ====================

    // CONSISTENCY(zip-uri-lookup): see ExcelHandler.cs / RawXmlHelper —
    // any partPath ending in `.xml` is resolved as a literal zip URI via
    // the package's part tree, no per-handler alias table needed.

    public string Raw(string partPath, int? startRow = null, int? endRow = null, HashSet<string>? cols = null)
    {
        if (partPath == null) throw new ArgumentNullException(nameof(partPath));
        var presentationPart = _doc.PresentationPart;
        if (presentationPart == null) return "(empty)";

        if (RawXmlHelper.IsZipUriPath(partPath))
        {
            var xml = RawXmlHelper.TryReadByZipUri(_doc, _filePath, partPath)
                ?? throw new ArgumentException(
                    $"Unknown part: {partPath}. The path was treated as a zip-internal URI " +
                    $"but no matching part exists in the package. " +
                    $"Use semantic paths (/presentation, /slide[N], /slideMaster[N]) for stable identification.");
            return xml;
        }

        if (partPath == "/" || partPath == "/presentation")
            return presentationPart.Presentation?.OuterXml ?? "(empty)";

        if (partPath == "/theme")
            return presentationPart.ThemePart?.Theme?.OuterXml ?? "(no theme)";

        var slideMatch = Regex.Match(partPath, @"^/slide\[(\d+)\]$");
        if (slideMatch.Success)
        {
            var idx = int.Parse(slideMatch.Groups[1].Value);
            var slideParts = GetSlideParts().ToList();
            if (idx >= 1 && idx <= slideParts.Count)
                return GetSlide(slideParts[idx - 1]).OuterXml;
            throw new ArgumentException($"slide[{idx}] not found (total: {slideParts.Count})");
        }

        // CONSISTENCY(raw-rawset-symmetry): RawSet supports master/layout/noteSlide;
        // Raw must too, otherwise users can't read back what they just wrote.
        var masterMatch = Regex.Match(partPath, @"^/slideMaster\[(\d+)\]$");
        if (masterMatch.Success)
        {
            var idx = int.Parse(masterMatch.Groups[1].Value);
            var masters = presentationPart.SlideMasterParts.ToList();
            if (idx < 1 || idx > masters.Count)
                throw new ArgumentException($"slideMaster[{idx}] not found (total: {masters.Count})");
            return masters[idx - 1].SlideMaster?.OuterXml
                ?? throw new InvalidOperationException("Corrupt file: slide master data missing");
        }

        var layoutMatch = Regex.Match(partPath, @"^/slideLayout\[(\d+)\]$");
        if (layoutMatch.Success)
        {
            var idx = int.Parse(layoutMatch.Groups[1].Value);
            var layouts = presentationPart.SlideMasterParts
                .SelectMany(m => m.SlideLayoutParts).ToList();
            if (idx < 1 || idx > layouts.Count)
                throw new ArgumentException($"slideLayout[{idx}] not found (total: {layouts.Count})");
            return layouts[idx - 1].SlideLayout?.OuterXml
                ?? throw new InvalidOperationException("Corrupt file: slide layout data missing");
        }

        var noteMatch = Regex.Match(partPath, @"^/noteSlide\[(\d+)\]$");
        if (noteMatch.Success)
        {
            var idx = int.Parse(noteMatch.Groups[1].Value);
            var slideParts = GetSlideParts().ToList();
            if (idx < 1 || idx > slideParts.Count)
                throw new ArgumentException($"slide[{idx}] not found (total: {slideParts.Count})");
            var notesPart = slideParts[idx - 1].NotesSlidePart
                ?? throw new ArgumentException($"Slide {idx} has no notes");
            return notesPart.NotesSlide?.OuterXml
                ?? throw new InvalidOperationException("Corrupt file: notes slide data missing");
        }

        // CONSISTENCY(raw-rawset-symmetry): /notesMaster surfaces the
        // presentation-level NotesMasterPart's XML so PptxBatchEmitter can
        // raw-set it on replay (mirrors theme/master/layout treatment).
        if (partPath == "/notesMaster")
        {
            return presentationPart.NotesMasterPart?.NotesMaster?.OuterXml
                ?? throw new ArgumentException("No notes master part");
        }

        throw new ArgumentException($"Unknown part: {partPath}. Available: /presentation, /theme, /slide[N], /slideMaster[N], /slideLayout[N], /noteSlide[N], /notesMaster");
    }

    public void RawSet(string partPath, string xpath, string action, string? xml)
    {
        if (partPath == null) throw new ArgumentNullException(nameof(partPath));
        var presentationPart = _doc.PresentationPart
            ?? throw new InvalidOperationException("No presentation part");

        if (RawXmlHelper.IsZipUriPath(partPath))
        {
            var part = RawXmlHelper.FindPartByZipUri(_doc, partPath)
                ?? throw new ArgumentException(
                    $"Unknown part: {partPath}. The path was treated as a zip-internal URI " +
                    $"but no matching part exists in the package. " +
                    $"Use semantic paths (/presentation, /slide[N], /slideMaster[N]) for stable identification.");
            RawXmlHelper.Execute(part, xpath, action, xml);
            return;
        }

        OpenXmlPartRootElement rootElement;

        if (partPath is "/" or "/presentation")
        {
            rootElement = presentationPart.Presentation
                ?? throw new InvalidOperationException("No presentation");
        }
        else if (partPath == "/theme")
        {
            rootElement = presentationPart.ThemePart?.Theme
                ?? throw new ArgumentException("No theme part");
        }
        else if (Regex.Match(partPath, @"^/slide\[(\d+)\]$") is { Success: true } slideMatch)
        {
            var idx = int.Parse(slideMatch.Groups[1].Value);
            var slideParts = GetSlideParts().ToList();
            if (idx < 1 || idx > slideParts.Count)
                throw new ArgumentException($"Slide {idx} not found (total: {slideParts.Count})");
            rootElement = GetSlide(slideParts[idx - 1]);
        }
        else if (Regex.Match(partPath, @"^/slideMaster\[(\d+)\]$") is { Success: true } masterMatch)
        {
            var idx = int.Parse(masterMatch.Groups[1].Value);
            var masters = presentationPart.SlideMasterParts.ToList();
            if (idx < 1)
                throw new ArgumentException($"SlideMaster {idx} not found (total: {masters.Count})");
            if (idx > masters.Count)
            {
                // CONSISTENCY(grow-on-rawset): mirrors the slideLayout branch.
                // Source decks with multiple slideMasters (template kits, decks
                // assembled from several themes) emit raw-set on /slideMaster[2..N];
                // blank target only stamped /slideMaster[1], so the replay used to
                // fail every additional master AND every layout owned by those
                // missing masters. Auto-grow to idx so the raw-set replace has a
                // root element to swap out; the replace then carries in the real
                // sldLayoutIdLst, which GrowSlideLayoutParts consults when the
                // subsequent /slideLayout[K] raw-sets land.
                GrowSlideMasterParts(idx);
                masters = presentationPart.SlideMasterParts.ToList();
                if (idx > masters.Count)
                    throw new ArgumentException($"SlideMaster {idx} not found (total: {masters.Count})");
            }
            rootElement = masters[idx - 1].SlideMaster
                ?? throw new InvalidOperationException("Corrupt file: slide master data missing");
        }
        else if (Regex.Match(partPath, @"^/slideLayout\[(\d+)\]$") is { Success: true } layoutMatch)
        {
            var idx = int.Parse(layoutMatch.Groups[1].Value);
            var layouts = presentationPart.SlideMasterParts
                .SelectMany(m => m.SlideLayoutParts).ToList();
            if (idx < 1)
                throw new ArgumentException($"SlideLayout {idx} not found (total: {layouts.Count})");
            if (idx > layouts.Count)
            {
                // BUG-J: Replay scenario — source deck has N layouts (e.g. 11) but
                // the blank target only stamped K (5). EmitMasterRaw already
                // replaced the master's sldLayoutIdLst to reference all N rIds,
                // but the SlideLayoutPart objects for K+1..N don't exist yet, so
                // raw-set fails with "SlideLayout {idx} not found". Auto-grow the
                // missing layout parts under the appropriate master based on the
                // post-master-replace sldLayoutIdLst, then re-resolve.
                GrowSlideLayoutParts(idx);
                layouts = presentationPart.SlideMasterParts
                    .SelectMany(m => m.SlideLayoutParts).ToList();
                if (idx > layouts.Count)
                    throw new ArgumentException($"SlideLayout {idx} not found (total: {layouts.Count})");
            }
            rootElement = layouts[idx - 1].SlideLayout
                ?? throw new InvalidOperationException("Corrupt file: slide layout data missing");
        }
        else if (Regex.Match(partPath, @"^/noteSlide\[(\d+)\]$") is { Success: true } noteMatch)
        {
            var idx = int.Parse(noteMatch.Groups[1].Value);
            var slideParts = GetSlideParts().ToList();
            if (idx < 1 || idx > slideParts.Count)
                throw new ArgumentException($"Slide {idx} not found (total: {slideParts.Count})");
            var notesPart = slideParts[idx - 1].NotesSlidePart
                ?? throw new ArgumentException($"Slide {idx} has no notes");
            rootElement = notesPart.NotesSlide
                ?? throw new InvalidOperationException("Corrupt file: notes slide data missing");
        }
        else if (partPath == "/notesMaster")
        {
            // CONSISTENCY(grow-on-rawset): blank pptx files have no
            // NotesMasterPart, but PptxBatchEmitter emits a raw-set /notesMaster
            // on any deck that has one. Create the part on demand so dump-replay
            // can stamp the source notes master back in (mirrors GrowSlideLayoutParts
            // in the slideLayout branch above).
            var nmPart = presentationPart.NotesMasterPart;
            if (nmPart == null)
            {
                nmPart = presentationPart.AddNewPart<NotesMasterPart>();
                // Seed a minimal placeholder so the raw-set "replace" action has
                // a NotesMaster root element to swap out; raw-replace builds the
                // real content from the supplied XML on the next line.
                nmPart.NotesMaster = new NotesMaster(
                    new CommonSlideData(new ShapeTree(
                        new NonVisualGroupShapeProperties(
                            new NonVisualDrawingProperties { Id = 1, Name = "" },
                            new NonVisualGroupShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties()),
                        new GroupShapeProperties(new DocumentFormat.OpenXml.Drawing.TransformGroup()))),
                    new ColorMap
                    {
                        Background1 = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Light1,
                        Text1 = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Dark1,
                        Background2 = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Light2,
                        Text2 = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Dark2,
                        Accent1 = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Accent1,
                        Accent2 = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Accent2,
                        Accent3 = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Accent3,
                        Accent4 = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Accent4,
                        Accent5 = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Accent5,
                        Accent6 = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Accent6,
                        Hyperlink = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Hyperlink,
                        FollowedHyperlink = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.FollowedHyperlink,
                    });
            }
            rootElement = nmPart.NotesMaster!;
        }
        else
        {
            throw new ArgumentException($"Unknown part: {partPath}. Available: /presentation, /theme, /slide[N], /slideMaster[N], /slideLayout[N], /noteSlide[N], /notesMaster");
        }

        var affected = RawXmlHelper.Execute(rootElement, xpath, action, xml);
        rootElement.Save();
        // After a /slideMaster[N] raw-set the master's <p:sldLayoutIdLst> is
        // the source's authoritative layout count. Blank decks ship with a
        // pre-stamped 5-layout master, so a 1-layout source replays to a
        // 5-layout deck — dump→batch→dump grows from 8 ops to 12 because
        // the 4 extra blank layouts survive. Prune SlideLayoutParts whose
        // rId is no longer in the post-replace sldLayoutIdLst so the
        // replayed deck mirrors the source's layout set exactly. The grow
        // path (line ~203) handles the opposite case (source has MORE).
        if (Regex.IsMatch(partPath, @"^/slideMaster\[\d+\]$") && rootElement is SlideMaster sm)
        {
            var mp = sm.SlideMasterPart;
            if (mp != null)
            {
                var declaredRids = new HashSet<string>(
                    sm.SlideLayoutIdList?.Elements<SlideLayoutId>()
                        .Select(e => e.RelationshipId?.Value ?? "")
                        .Where(s => !string.IsNullOrEmpty(s))
                    ?? Enumerable.Empty<string>(),
                    StringComparer.Ordinal);
                foreach (var pair in mp.Parts.ToList())
                {
                    if (pair.OpenXmlPart is SlideLayoutPart lp
                        && !declaredRids.Contains(pair.RelationshipId))
                    {
                        // The orphan layout's rels (theme/image links etc.) drop
                        // with the part; DeletePart cascades.
                        mp.DeletePart(lp);
                    }
                }
            }
        }
        // BUG-R43: raw-set may have inserted/removed shape XML directly (incl.
        // cNvPr ids). The cached _usedShapeIds set is now stale, so the next
        // Add() can hand out an id that already exists in the tree, producing
        // duplicate cNvPr ids that PowerPoint silently rejects. Rebuild the
        // shape-id index from the live tree after every raw-set.
        InitShapeIdCounter();
        // BUG-R5-01: silent — CLI wrappers print their own structured message.
        _ = affected;
    }

    // BUG-J: Auto-grow SlideLayoutPart objects so raw-set replay can target
    // /slideLayout[targetGlobalIdx] even when the blank deck only stamped a
    // subset. The master's sldLayoutIdLst (already replaced by EmitMasterRaw)
    // is the source of truth: each entry holds the rId that the new
    // SlideLayoutPart must register with so the master-layout relationship
    // matches. We walk masters in declaration order, compute each master's
    // declared layout count from sldLayoutIdLst, and create missing parts
    // under whichever master's range contains targetGlobalIdx. Newly created
    // parts get a stub root (the imminent replace overwrites it).
    private void GrowSlideLayoutParts(int targetGlobalIdx)
    {
        var presentationPart = _doc.PresentationPart
            ?? throw new InvalidOperationException("No presentation part");
        var masters = presentationPart.SlideMasterParts.ToList();
        int seen = 0;
        foreach (var mp in masters)
        {
            var declared = mp.SlideMaster?.SlideLayoutIdList?.Elements<SlideLayoutId>().ToList()
                ?? new List<SlideLayoutId>();
            int declaredCount = declared.Count;
            int existingCount = mp.SlideLayoutParts.Count();
            // This master "owns" global indices (seen+1)..(seen+declaredCount).
            int rangeStart = seen + 1;
            int rangeEnd = seen + declaredCount;
            if (targetGlobalIdx >= rangeStart && targetGlobalIdx <= rangeEnd)
            {
                // Create missing parts for slots existingCount+1 .. declaredCount.
                for (int slot = existingCount; slot < declaredCount; slot++)
                {
                    var declaredId = declared[slot];
                    var rId = declaredId.RelationshipId?.Value;
                    SlideLayoutPart newPart;
                    if (!string.IsNullOrEmpty(rId) && !mp.Parts.Any(p => p.RelationshipId == rId))
                    {
                        newPart = mp.AddNewPart<SlideLayoutPart>(rId);
                    }
                    else
                    {
                        // Either rId missing in sldLayoutIdLst or already taken
                        // (corruption guard) — let OpenXml allocate a new one
                        // and patch the sldLayoutIdLst entry to match.
                        newPart = mp.AddNewPart<SlideLayoutPart>();
                        var newRid = mp.GetIdOfPart(newPart);
                        declaredId.RelationshipId = newRid;
                    }
                    // Stub root — the raw-set replace immediately rewrites it.
                    newPart.SlideLayout = new SlideLayout(
                        new CommonSlideData(
                            new ShapeTree(
                                new NonVisualGroupShapeProperties(
                                    new NonVisualDrawingProperties { Id = 1, Name = "" },
                                    new NonVisualGroupShapeDrawingProperties(),
                                    new ApplicationNonVisualDrawingProperties()),
                                new GroupShapeProperties()))
                    ) { Type = SlideLayoutValues.Blank };
                    newPart.SlideLayout.Save();
                    // Layouts must point back to their master.
                    newPart.AddPart(mp);
                }
                if (mp.SlideMaster != null) mp.SlideMaster.Save();
                return;
            }
            seen += declaredCount;
        }
        // targetGlobalIdx is beyond every master's declared range; caller will
        // raise the canonical "not found" error after we return.
    }

    // CONSISTENCY(grow-on-rawset): mirror of GrowSlideLayoutParts for the
    // SlideMasterPart side. Multi-master source decks (template kits, decks
    // assembled from multiple themes) emit raw-set on /slideMaster[2..N], but
    // BlankDocCreator only stamps one master. Create enough placeholder
    // SlideMasterParts (each with a minimal SlideMaster root plus its own
    // SlideLayoutIdList stub) and register them in the presentation's
    // sldMasterIdLst so the raw-set replace has a root element to swap, and
    // so subsequent /slideLayout[K] raw-sets can find their owning master via
    // GrowSlideLayoutParts.
    private void GrowSlideMasterParts(int targetIdx)
    {
        var presentationPart = _doc.PresentationPart
            ?? throw new InvalidOperationException("No presentation part");
        var presentation = presentationPart.Presentation
            ?? throw new InvalidOperationException("No presentation");
        var sldMasterIdLst = presentation.SlideMasterIdList
            ?? throw new InvalidOperationException("Presentation has no SlideMasterIdList");

        var existing = presentationPart.SlideMasterParts.Count();
        if (targetIdx <= existing) return;

        // Pick a SlideMasterId base that won't collide with the existing IDs.
        var existingIds = sldMasterIdLst.Elements<SlideMasterId>()
            .Select(e => e.Id?.Value ?? 0u)
            .ToHashSet();
        uint nextId = 2147483648u;
        while (existingIds.Contains(nextId)) nextId++;

        for (int i = existing; i < targetIdx; i++)
        {
            var newPart = presentationPart.AddNewPart<SlideMasterPart>();
            var rId = presentationPart.GetIdOfPart(newPart);
            // Minimal SlideMaster root with an empty SlideLayoutIdList so
            // GrowSlideLayoutParts sees the right declared count once the
            // imminent raw-set replace overwrites this stub with the real
            // master XML (which carries the source's actual sldLayoutIdLst).
            newPart.SlideMaster = new SlideMaster(
                new CommonSlideData(new ShapeTree(
                    new NonVisualGroupShapeProperties(
                        new NonVisualDrawingProperties { Id = 1, Name = "" },
                        new NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new DocumentFormat.OpenXml.Drawing.TransformGroup()))),
                new ColorMap
                {
                    Background1 = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Light1,
                    Text1 = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Dark1,
                    Background2 = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Light2,
                    Text2 = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Dark2,
                    Accent1 = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Accent1,
                    Accent2 = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Accent2,
                    Accent3 = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Accent3,
                    Accent4 = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Accent4,
                    Accent5 = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Accent5,
                    Accent6 = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Accent6,
                    Hyperlink = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Hyperlink,
                    FollowedHyperlink = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.FollowedHyperlink,
                },
                new SlideLayoutIdList()
            );
            newPart.SlideMaster.Save();

            // Every SlideMasterPart must reference a ThemePart. The replay's
            // raw-set replaces only the master XML, not the package
            // relationships — without a theme link here the package fails
            // validation. Share the presentation's primary theme; a richer
            // multi-theme deck can later raw-set its own theme parts.
            if (presentationPart.ThemePart != null)
                newPart.AddPart(presentationPart.ThemePart);

            sldMasterIdLst.AppendChild(new SlideMasterId { Id = nextId++, RelationshipId = rId });
        }
        presentation.Save();
    }

    public (string RelId, string PartPath) AddPart(string parentPartPath, string partType, Dictionary<string, string>? properties = null)
    {
        var presentationPart = _doc.PresentationPart
            ?? throw new InvalidOperationException("No presentation part");

        switch (partType.ToLowerInvariant())
        {
            case "chart":
                // Charts go under a SlidePart
                var slideMatch = System.Text.RegularExpressions.Regex.Match(
                    parentPartPath, @"^/slide\[(\d+)\]$");
                if (!slideMatch.Success)
                    throw new ArgumentException(
                        "Chart must be added under a slide: add-part <file> '/slide[N]' --type chart");

                var slideIdx = int.Parse(slideMatch.Groups[1].Value);
                var slideParts = GetSlideParts().ToList();
                if (slideIdx < 1 || slideIdx > slideParts.Count)
                    throw new ArgumentException($"Slide index {slideIdx} out of range");

                var slidePart = slideParts[slideIdx - 1];
                var chartPart = slidePart.AddNewPart<DocumentFormat.OpenXml.Packaging.ChartPart>();
                var relId = slidePart.GetIdOfPart(chartPart);

                chartPart.ChartSpace = new DocumentFormat.OpenXml.Drawing.Charts.ChartSpace(
                    new DocumentFormat.OpenXml.Drawing.Charts.Chart(
                        new DocumentFormat.OpenXml.Drawing.Charts.PlotArea(
                            new DocumentFormat.OpenXml.Drawing.Charts.Layout()
                        )
                    )
                );
                chartPart.ChartSpace.Save();

                var chartIdx = slidePart.ChartParts.ToList().IndexOf(chartPart);
                return (relId, $"/slide[{slideIdx}]/chart[{chartIdx + 1}]");

            default:
                throw new ArgumentException(
                    $"Unknown part type: {partType}. Supported: chart");
        }
    }

    public List<ValidationError> Validate() => RawXmlHelper.ValidateDocument(_doc);

    public void Dispose()
    {
        // Save through the package (flush in-memory edits to the underlying
        // stream) before disposing. When we own the backing FileStream, the
        // package would otherwise leave the on-disk file in whatever state
        // the last auto-flush left it — for the stream-Open path this can
        // truncate to zero bytes and look like a corrupted zip on reopen.
        try { _doc.Save(); } catch { /* read-only or already disposed */ }
        _doc.Dispose();
        _backingStream?.Dispose();
        _backingStream = null;
    }

    // Internal accessors used by PptxBatchEmitter (resource enumeration).
    // Keep the PresentationPart itself private; expose only the counts and
    // a binary getter that the emitter needs.
    internal int SlideMasterCount =>
        _doc.PresentationPart?.SlideMasterParts.Count() ?? 0;
    internal int SlideLayoutCount =>
        _doc.PresentationPart?.SlideMasterParts.SelectMany(m => m.SlideLayoutParts).Count() ?? 0;
    internal bool HasNotesMaster =>
        _doc.PresentationPart?.NotesMasterPart != null;
    // Exposed for PptxBatchEmitter so it can iterate slides without going
    // through Get("/") — Get("/") fans out into per-slide deep walks that
    // can throw at SDK validation time on vendor templates with foreign
    // attributes (gov_bja, 1.pptx, ...). The emitter now uses this count
    // plus per-slide try/catch to keep the dump going on partial corruption.
    internal int SlideCount => GetSlideParts().Count();

    internal bool SlideHasNotes(int slideIdx)
    {
        var parts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > parts.Count) return false;
        return parts[slideIdx - 1].NotesSlidePart != null;
    }

    // Resolve a /slide[N]/picture[M] path's image bytes for base64-inline emit.
    // Mirrors WordHandler.GetImageBinary's contract: returns null if the path
    // does not resolve to a Picture with an embedded ImagePart.
    public (byte[] Bytes, string ContentType)? GetImageBinary(string picturePath)
    {
        // Accept both `picture[N]` positional and `picture[@id=N]` cNvPr-id
        // segment forms (BuildElementPathSegment emits @id= when the shape
        // carries a cNvPr id, which Pictures always do).
        var m = Regex.Match(picturePath,
            @"^/slide\[(\d+)\]/(?:.+/)?picture\[(?:@id=)?(\d+)\]$");
        if (!m.Success) return null;
        var slideIdx = int.Parse(m.Groups[1].Value);
        var idOrIdx = int.Parse(m.Groups[2].Value);
        var byId = picturePath.Contains("@id=", StringComparison.Ordinal);
        var parts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > parts.Count) return null;
        var slidePart = parts[slideIdx - 1];
        var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree;
        if (shapeTree == null) return null;
        var pictures = shapeTree.Descendants<Picture>().ToList();
        Picture? pic = null;
        if (byId)
        {
            pic = pictures.FirstOrDefault(p =>
            {
                var pid = p.NonVisualPictureProperties?.NonVisualDrawingProperties?.Id?.Value;
                return pid.HasValue && pid.Value == (uint)idOrIdx;
            });
        }
        else
        {
            if (idOrIdx >= 1 && idOrIdx <= pictures.Count) pic = pictures[idOrIdx - 1];
        }
        if (pic == null) return null;
        var blip = pic.BlipFill?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Blip>();
        var embedId = blip?.Embed?.Value;
        if (string.IsNullOrEmpty(embedId)) return null;
        try
        {
            var part = slidePart.GetPartById(embedId);
            using var src = part.GetStream();
            using var ms = new MemoryStream();
            src.CopyTo(ms);
            return (ms.ToArray(), part.ContentType);
        }
        catch { return null; }
    }

    // ==================== Private Helpers ====================

    private static Slide GetSlide(SlidePart part) =>
        part.Slide ?? throw new InvalidOperationException("Corrupt file: slide data missing");

    private IEnumerable<SlidePart> GetSlideParts()
    {
        var presentation = _doc.PresentationPart?.Presentation;
        var slideIdList = presentation?.GetFirstChild<SlideIdList>();
        if (slideIdList == null) yield break;

        foreach (var slideId in slideIdList.Elements<SlideId>())
        {
            var relId = slideId.RelationshipId?.Value;
            if (relId == null) continue;
            yield return (SlidePart)_doc.PresentationPart!.GetPartById(relId);
        }
    }

}
